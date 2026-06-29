# Durable Submission Record Schema and Implementation Plan

This document reconciles the mockup in `form_submit_tool.html` with the current Form Manager flow observed in the read-only global files:

- `a:\GLOBAL_6-next\www\includes\forms.inc` - front-end form processor.
- `a:\GLOBAL_6-next\admin\apprenewadmin.asp` - admin submission list/detail tool.

Those global files are reference material only. The plan below describes the schema and integration points needed to produce the mockup behavior without editing those files directly in this project.

> **Reference implementation:** Local, patched copies of both global files (plus the new helper class and Oracle DDL) live under [`../patched/`](../patched/). They are not deployed and the originals were never edited; see [`../patched/README.md`](../patched/README.md) for the change list and deploy order. Search those files for `FORM SUBMIT LEDGER PATCH` to review every edit.

## What The Global Code Already Does

`forms.inc` already creates and updates an `app_renew` row during form progress and submit. The key durable anchor is `saveFormData(formStatus)`, which inserts or updates `app_renew(app_renew_id, app_renew_type, app_renew_status, membership_id)` and can save the current session data.

The processor already invokes the core operations the mockup displays:

- Payment: payment pages call `paymentStatus "Pending"`, `paymentStatus "Complete"`, or `paymentStatus "Failed"`. The result is stored in `app_renew_detail` as `PAYMENT_STATUS` and `PAYMENT_DATE_TIME`.
- Uploaded files: page and config branches call `cRenew.SaveFiles`.
- Record updates: save-record pages and page options call `cRenew.SaveData`.
- PDF: configured pages call `cRenew.PDF`, `cRenew.PDFQuickPost`, or rebuild helpers.
- Emails: page options call `cRenew.EmailPage` and `cRenew.EmailPagesOnclick`.
- Awaiting Updates: the processor sets `session("DO_AWAIT")`, and admin checks `mem_await` / `mem_det_await` for unreleased rows.
- Billing / batch transactions: payment and page options call `batchTransaction`, `batchTransactionPending`, and related helpers.
- Custom code: page option `execute_code` is executed during submit flow.

`apprenewadmin.asp` already renders a list with Form Type, Status, PDF, Awaiting Updates, Payment, dates, and links to a modify/detail view. It derives:

- PDF from `form_pdf.status`.
- Awaiting Updates from `mem_await` / `mem_det_await`.
- Payment from `app_renew_detail` fields `card_amount`, `payment_status`, and `payment_date_time`.
- Legacy status from `app_renew.app_renew_status`.

The missing piece is not the existence of the operations. The missing piece is a durable operation ledger that says what was expected to happen, what did happen, what is still running, and what can be safely retried.

## Reconciliation With The Mockup

The mockup should be understood as a new observability and recovery layer around the existing processor.

The existing `app_renew_status` values (`PROGRESS`, `PENDING`, `APPROVED`, `COMPLETED`, `DENIED`, `SAVED`, `PAY_LATER`, `DELETED`) are workflow/admin statuses. They should not be stretched to explain every execution step.

The mockup's durable `state` is a separate lifecycle:

```text
draft -> awaiting_payment -> payment_complete -> processing -> completed
                         \-> abandoned
processing -> failed
```

Definitions:

- `draft`: an `app_renew` row exists, but final submission has not happened.
- `awaiting_payment`: the user submitted and was sent to payment; the gateway result is not known.
- `payment_complete`: payment is confirmed, but post-submit operations have not all completed.
- `processing`: the submission is eligible to run post-submit operations and at least one expected operation is queued, running, retrying, pending, or awaiting review.
- `completed`: every expected operation is done or not needed.
- `failed`: an expected operation exhausted retry / failed in a way that requires attention.
- `abandoned`: the submission will not complete automatically, usually due to timeout, missing payment return, or manual abandonment.

The mockup's operation cards map to current hooks as follows:

- Payment maps to `paymentStatus`, payment config (`payment_page`, `payment_processor`), `CARD_AMOUNT`, and bypasses where a conditional page skips the configured payment page.
- Uploaded files maps to `cRenew.SaveFiles` and the session file arrays / file persistence work.
- Record updates maps to `cRenew.SaveData` plus Awaiting Updates holds in `mem_await` / `mem_det_await`.
- PDF maps to `cRenew.PDF`, `cRenew.PDFQuickPost`, `pdfRebuildByID`, and `form_pdf`.
- Emails maps to `cRenew.EmailPage` and `cRenew.EmailPagesOnclick`.
- Billing history maps to batch transaction helpers.
- Custom code maps to page option `execute_code`.

Form-level configuration and per-submission triggers must be separated:

- Form-level configuration is expected to be the same for every submission of a form code, such as RENEW requiring payment and generating PDFs.
- Triggered behavior is conditional on a field value for a specific submission, such as `payment_ty = K` bypassing `pay.asp`, or `name_change = Yes` requiring an Awaiting Updates review.

## Proposed Schema

The schema below adds durable execution tracking without replacing the legacy `app_renew` row. `app_renew_id` remains the anchor so existing admin links, logs, and restore flows continue to work.

## Oracle Data Type Recommendations

Use Oracle-native types that fit the existing Classic ASP / Oracle style while still making the new ledger precise enough for recovery.

- IDs: use `NUMBER(12)` for `app_renew_id`, `operation_id`, `item_id`, and `event_id`. This is large enough for long-running sequence growth and matches the existing sequence-oriented code style.
- Existing member IDs: use `NUMBER(12)` for `membership_id`. Do not assume every in-progress form has a real membership id; `0` and null-like values exist in the legacy flow.
- Codes / states / operation keys: use `VARCHAR2(... CHAR)` rather than byte semantics so labels remain predictable if non-ASCII values ever appear. Keep operational enums short:
  - `durable_state VARCHAR2(30 CHAR)`
  - `operation_key VARCHAR2(50 CHAR)`
  - `state VARCHAR2(30 CHAR)`
  - `source_scope VARCHAR2(30 CHAR)`
- Booleans: use `CHAR(1)` with `check (... in ('Y','N'))`, not `VARCHAR2(1)`, for values such as `configured`, `ready_to_process`, and `retryable`.
- Timestamps: prefer `TIMESTAMP(6) WITH LOCAL TIME ZONE` for all new `*_at` columns. It preserves ordering and avoids server/time-zone ambiguity. If the current Oracle version or ASP data-access layer makes that painful, use `DATE` as a compatibility fallback and keep all writes through `sysdate`.
- Money: use `NUMBER(12,2)` for payment amounts, with `currency_code CHAR(3 CHAR)` when a column is needed. Keep the existing `app_renew_detail.card_amount` as legacy source data.
- Large text: use `VARCHAR2(4000 CHAR)` for short user-visible messages and `CLOB` for payloads, stack traces, config snapshots, and raw provider responses.
- JSON: use `CLOB` plus an `IS JSON` check constraint if the Oracle version supports it. If not, still keep JSON in `CLOB` and validate in helper code.
- File hashes: use `VARCHAR2(64 CHAR)` for SHA-256 hex hashes, or `VARCHAR2(128 CHAR)` if the hash algorithm may vary.
- Paths / URLs / email targets: use `VARCHAR2(1000 CHAR)` for normal targets and `VARCHAR2(2000 CHAR)` where gateway URLs or storage paths may be long.
- Counters / attempts: use `NUMBER(5)` for retry attempts and `NUMBER(9)` for counts. These should never be unbounded.

Recommended defaults:

```sql
created_at timestamp(6) with local time zone default systimestamp not null
updated_at timestamp(6) with local time zone default systimestamp not null
```

Compatibility fallback:

```sql
created_at date default sysdate not null
updated_at date default sysdate not null
```

Use one style consistently. The SQL below shows the preferred timestamp type.

### `form_submit_record`

One row per `app_renew_id`.

```sql
create table form_submit_record (
    app_renew_id number(12) primary key,
    app_renew_type varchar2(50 char) not null,
    membership_id number(12),

    durable_state varchar2(30 char) not null,
    legacy_app_renew_status varchar2(30 char),

    submitted_at timestamp(6) with local time zone,
    ready_to_process char(1 char) default 'N' not null,
    processing_started_at timestamp(6) with local time zone,
    completed_at timestamp(6) with local time zone,
    abandoned_at timestamp(6) with local time zone,

    current_operation_key varchar2(50 char),
    last_error varchar2(4000 char),
    retry_after timestamp(6) with local time zone,

    created_at timestamp(6) with local time zone default systimestamp not null,
    updated_at timestamp(6) with local time zone default systimestamp not null,

    constraint ck_fsr_ready_to_process check (ready_to_process in ('Y','N')),
    constraint ck_fsr_durable_state check (durable_state in (
        'draft','awaiting_payment','payment_complete','processing',
        'completed','failed','abandoned'
    ))
);
```

Allowed `durable_state` values:

```text
draft
awaiting_payment
payment_complete
processing
completed
failed
abandoned
```

This row answers the mockup's top-level State column. It should not replace `app_renew.app_renew_status`; instead, it gives the recovery worker a durable lifecycle that is specific to execution.

### `form_submit_operation`

One row per operation expected or considered for a submission.

```sql
create table form_submit_operation (
    operation_id number(12) primary key,
    app_renew_id number(12) not null,
    operation_key varchar2(50 char) not null,

    configured char(1 char) not null,
    source_scope varchar2(30 char) not null,
    source_label varchar2(200 char),
    trigger_field varchar2(100 char),
    trigger_operator varchar2(20 char),
    trigger_value varchar2(400 char),

    state varchar2(30 char) not null,
    expected_count number(9) default 0 not null,
    done_count number(9) default 0 not null,
    failed_count number(9) default 0 not null,

    attempts number(5) default 0 not null,
    max_attempts number(5) default 0 not null,
    next_attempt_at timestamp(6) with local time zone,
    job_key varchar2(100 char),

    started_at timestamp(6) with local time zone,
    completed_at timestamp(6) with local time zone,
    last_error varchar2(4000 char),
    notes varchar2(4000 char),

    created_at timestamp(6) with local time zone default systimestamp not null,
    updated_at timestamp(6) with local time zone default systimestamp not null,

    constraint uq_form_submit_operation unique (app_renew_id, operation_key),
    constraint ck_fso_configured check (configured in ('Y','N')),
    constraint ck_fso_operation_key check (operation_key in (
        'payment','files','records','pdf','emails','billing','custom_code'
    )),
    constraint ck_fso_source_scope check (source_scope in (
        'form_config','triggered','not_configured','not_needed','legacy_inferred'
    )),
    constraint ck_fso_state check (state in (
        'not_configured','not_needed','ready','queued','running',
        'pending','retrying','awaiting','done','failed'
    ))
);
```

Allowed `operation_key` values initially:

```text
payment
files
records
pdf
emails
billing
custom_code
```

Allowed `source_scope` values:

```text
form_config
triggered
not_configured
not_needed
legacy_inferred
```

Allowed `state` values:

```text
not_configured
not_needed
ready
queued
running
pending
retrying
awaiting
done
failed
```

This row answers each operation card header and the compact grid cells. Counts are operation-level totals, such as `1/2` emails sent or `3/4` files persisted.

### `form_submit_operation_item`

Optional child rows for operations with multiple pieces, such as files and emails.

```sql
create table form_submit_operation_item (
    item_id number(12) primary key,
    operation_id number(12) not null,
    app_renew_id number(12) not null,
    operation_key varchar2(50 char) not null,

    item_key varchar2(200 char) not null,
    item_label varchar2(500 char),
    item_target varchar2(1000 char),

    state varchar2(30 char) not null,
    source_scope varchar2(30 char),
    trigger_field varchar2(100 char),
    trigger_value varchar2(400 char),

    temp_location varchar2(1000 char),
    permanent_location varchar2(1000 char),
    content_hash varchar2(128 char),
    size_bytes number(12),

    attempts number(5) default 0 not null,
    max_attempts number(5) default 0 not null,
    next_attempt_at timestamp(6) with local time zone,
    last_error varchar2(4000 char),

    created_at timestamp(6) with local time zone default systimestamp not null,
    updated_at timestamp(6) with local time zone default systimestamp not null,

    constraint uq_form_submit_operation_item unique (operation_id, item_key),
    constraint ck_fsoi_state check (state in (
        'not_configured','not_needed','ready','queued','running',
        'pending','retrying','awaiting','done','failed'
    ))
);
```

Examples:

- One row per uploaded file selected by the user.
- One row per configured / triggered email that should send.
- One row per generated PDF artifact if PDF variants are needed later.

This is what makes it possible to show "3 files uploaded, 2 saved, 1 stuck in temp" without hiding it behind a boolean.

### `form_submit_event`

Append-only timeline for detail view and troubleshooting.

```sql
create table form_submit_event (
    event_id number(12) primary key,
    app_renew_id number(12) not null,
    operation_key varchar2(50 char),
    item_id number(12),

    event_type varchar2(100 char) not null,
    severity varchar2(20 char) default 'info' not null,
    message varchar2(4000 char) not null,
    payload_json clob,

    created_at timestamp(6) with local time zone default systimestamp not null,

    constraint ck_fse_severity check (severity in ('debug','info','warn','error'))
);
```

This powers the mockup's pipeline checkpoints and makes recovery actions explainable to staff.

### `form_submit_config_snapshot`

One snapshot of the relevant form config at submission time.

```sql
create table form_submit_config_snapshot (
    app_renew_id number(12) primary key,
    app_renew_type varchar2(50 char) not null,
    config_version varchar2(100 char),
    config_json clob not null,
    created_at timestamp(6) with local time zone default systimestamp not null
);
```

This prevents confusing history when a form's configuration changes after a submission was created. Admin can show the expectations that were true when that submission was finalized.

## Oracle Schema Change Suggestions

The safest approach is additive: create new ledger tables beside the existing Form Manager tables and avoid altering `app_renew` until the new flow has proven itself.

### New sequences

Use explicit sequences because the existing Classic ASP codebase already uses sequence helpers such as `getSequence(...)`.

```sql
create sequence seq_form_submit_operation_id start with 1 increment by 1 nocache;
create sequence seq_form_submit_operation_item_id start with 1 increment by 1 nocache;
create sequence seq_form_submit_event_id start with 1 increment by 1 nocache;
```

If the environment standard is cached sequences, use a small cache such as `cache 100`; the IDs do not need to be gap-free.

### Foreign keys

Add foreign keys where the existing schema allows it. If legacy deployments avoid formal FKs, still create the same indexes and enforce relationships in helper code.

```sql
alter table form_submit_record
    add constraint fk_fsr_app_renew
    foreign key (app_renew_id) references app_renew(app_renew_id);

alter table form_submit_operation
    add constraint fk_fso_record
    foreign key (app_renew_id) references form_submit_record(app_renew_id);

alter table form_submit_operation_item
    add constraint fk_fsoi_operation
    foreign key (operation_id) references form_submit_operation(operation_id);

alter table form_submit_event
    add constraint fk_fse_record
    foreign key (app_renew_id) references form_submit_record(app_renew_id);
```

Do not use cascading deletes as the default. `app_renew` is soft-deleted in admin, and the ledger is audit/recovery data. If purge tooling is later added, make deletion explicit.

### Indexes for admin and recovery

```sql
create index ix_fsr_type_state_updated
    on form_submit_record(app_renew_type, durable_state, updated_at);

create index ix_fsr_retry
    on form_submit_record(durable_state, retry_after);

create index ix_fso_app_state
    on form_submit_operation(app_renew_id, state);

create index ix_fso_recovery
    on form_submit_operation(state, next_attempt_at, operation_key);

create index ix_fsoi_app_operation_state
    on form_submit_operation_item(app_renew_id, operation_key, state);

create index ix_fse_app_created
    on form_submit_event(app_renew_id, created_at);
```

These support the list view, detail timeline, and recovery worker. Add a function-based or filtered-style index only after observing actual query plans; keep the first migration simple.

### JSON constraints

If supported by the Oracle version:

```sql
alter table form_submit_event
    add constraint ck_fse_payload_json check (payload_json is json);

alter table form_submit_config_snapshot
    add constraint ck_fscs_config_json check (config_json is json);
```

If not supported, omit these constraints and validate JSON in the helper layer before inserting.

### Compatibility view for admin

Create a view that summarizes the operation rows into the mockup's grid shape. This lets `apprenewadmin.asp` adopt the new data with a smaller SQL change.

```sql
create or replace view v_form_submit_admin as
select
    r.app_renew_id,
    r.app_renew_type,
    r.membership_id,
    r.durable_state,
    r.legacy_app_renew_status,
    max(case when o.operation_key = 'payment' then o.state end) payment_state,
    max(case when o.operation_key = 'files' then o.state end) files_state,
    max(case when o.operation_key = 'files' then o.done_count end) files_done_count,
    max(case when o.operation_key = 'files' then o.expected_count end) files_expected_count,
    max(case when o.operation_key = 'pdf' then o.state end) pdf_state,
    max(case when o.operation_key = 'emails' then o.state end) emails_state,
    max(case when o.operation_key = 'emails' then o.done_count end) emails_done_count,
    max(case when o.operation_key = 'emails' then o.expected_count end) emails_expected_count,
    max(case when o.operation_key = 'records' then o.state end) records_state,
    r.updated_at
from form_submit_record r
left join form_submit_operation o on o.app_renew_id = r.app_renew_id
group by
    r.app_renew_id,
    r.app_renew_type,
    r.membership_id,
    r.durable_state,
    r.legacy_app_renew_status,
    r.updated_at;
```

This is intentionally generic. If admin needs payment amount/date, keep reading the existing legacy fields or add operation-specific columns after the ledger is populated reliably.

### Existing table changes

Recommended first migration: no required changes to existing tables.

Optional later changes:

- Add `app_renew.form_submit_state VARCHAR2(30 CHAR)` only as a denormalized cache if list performance requires it. The source of truth should remain `form_submit_record.durable_state`.
- Add `app_renew.form_submit_updated_at` only if the admin list must sort without joining the ledger.
- Add a back-reference from `form_pdf` to `form_submit_operation.operation_id` only if PDF retries need exact artifact lineage.
- Do not add more `app_renew_detail` pseudo-fields for new operation states. That table already acts as a field/value store; using it for a recovery ledger will make counts, retries, and item-level status harder to query.

## Operation Semantics

Payment:

- `configured = Y`, `source_scope = form_config` when the form has a configured payment page / processor.
- `state = ready` before final submit reaches payment.
- `state = awaiting` after redirect to gateway before a known callback/result.
- `state = done` when payment is confirmed.
- `state = failed` when a known gateway failure occurs.
- `state = not_needed` when a conditional flow bypasses the configured payment page, with `source_scope = triggered`, `trigger_field = payment_ty`, `trigger_value = K`, and `notes` explaining the alternate page.
- **List view display** (legacy-admin scan pattern; maps from canonical `state` above):
  - `done` → `Complete ($amount)`
  - `awaiting` / `pending` → `Pending ($amount)` — payment **not** completed
  - `ready` → `Not paid ($amount)`
  - `failed` → `Failed ($amount)`
  - `not_configured` / `not_needed` / bypass → `NA`
- Canonical `state` is stored in `form_submit_operation`; display labels are a view/UI concern.

Files:

- Create one operation item for every uploaded file the user actually selected.
- `expected_count` is the number of selected files, not the number of upload fields.
- `done_count` is the number persisted to permanent storage.
- Optional upload fields with no uploaded files should be `not_needed`, not failed.

Records / Awaiting Updates:

- `configured = Y` if this form can update CRM records.
- `state = awaiting` when an active `mem_await` or `mem_det_await` row exists with no approval date.
- `source_scope = triggered` when the hold is caused by a field value, such as `name_change = Yes` or `felony_record = Yes`.
- "Awaiting Updates = Yes" in admin should mean actual persisted hold now, not merely "configured to possibly hold."
- Review status should be derived:
  - `Auto`: no review needed.
  - `Awaiting`: active hold exists now.
  - `Approved`: a hold existed and has been released.

PDF:

- `configured = Y` when the form or page config says to generate a PDF.
- `ready` before the prerequisite step.
- `queued` / `running` when job execution is pending or active.
- `done` when the generated artifact is persisted and visible in `form_pdf` or equivalent storage.
- `failed` when generation or persistence fails.

Emails:

- If a form type has no emails, the operation is `not_configured` and the grid should show `NA` / dash.
- Form-config emails should appear for every submission of that form type, with per-submission state differences only.
- Triggered email pages should record `source_scope = triggered` and the triggering field condition.
- Counts exclude `not_needed` items. Example: if the normal payment receipt is skipped because payment was bypassed, it should be shown as not sent / not needed, while the triggered alternate receipt email can be done.

Billing:

- Map to batch transaction helpers.
- `not_configured` for form types that do not create billing history.
- `queued` behind payment / record updates when appropriate.

Custom code:

- Map to page option `execute_code`.
- Treat configured custom code as an operation even if it currently runs inline.
- Capture success/failure so staff can see whether the script ran.

## Implementation Plan

### Phase 1: Read-only derived dashboard

Goal: make admin clearer without changing the processor.

- Add read-only derivation queries around existing tables:
  - `app_renew` for identity and legacy status.
  - `app_renew_detail` for payment amount/status/date and dynamic field values.
  - `form_pdf` for PDF status.
  - `mem_await` / `mem_det_await` for active and historical Awaiting Updates.
  - Existing logs for best-effort failure clues.
- Use the mockup vocabulary in the admin UI but mark uncertain derived values as `legacy_inferred`.
- Keep Form Type filtering central. Different form codes should have different expected configuration; submissions of the same form code should not appear to have different form-level config.

### Phase 2: Add schema and helper API

Goal: give the processor a small, idempotent API to record intent and execution.

Create helper routines such as:

```text
EnsureSubmitRecord(app_renew_id, app_renew_type, membership_id)
SyncLegacyStatus(app_renew_id, form_status)
SetDurableState(app_renew_id, durable_state)
RecomputeSubmitState(app_renew_id)
SnapshotConfig(app_renew_id)
EnsureOperation(app_renew_id, operation_key, configured, source_scope, ...)
SetOperationState(app_renew_id, operation_key, state, ...)
SetOperationTrigger(app_renew_id, operation_key, field_name, operator, value)
EnsureOperationItem(app_renew_id, operation_key, item_key, ...)
SetOperationItemState(item_id, state, ...)
EnsureFileItem(app_renew_id, field_name, original_file_name, ...)
EnsureEmailItem(app_renew_id, email_key, recipient, subject, ...)
AppendSubmitEvent(app_renew_id, operation_key, event_type, severity, message, payload_json)
```

These helpers must be idempotent because `forms.inc` can be re-entered by refreshes, redirects, PDF/email modes, and recovery.

### Phase 3: Instrument existing operation boundaries

Goal: record facts at the point they already happen.

Recommended boundaries:

- Before redirecting to payment: set `payment = awaiting` and `durable_state = awaiting_payment`.
- Inside `paymentStatus`: update payment operation and event rows.
- Around `cRenew.SaveFiles`: create file items, then mark each persisted / failed.
- Around `cRenew.SaveData`: mark records running/done/failed or awaiting.
- When `mem_await` / `mem_det_await` rows are created: set records operation to `awaiting` with trigger details where available.
- When awaiting rows are approved: set records operation to `done` or release it back to `queued`.
- Around `cRenew.PDF`: mark PDF queued/running/done/failed.
- Around `cRenew.EmailPage` / `EmailPagesOnclick`: create email items and mark per-email result.
- Around batch transaction helpers: mark billing state.
- Around page `execute_code`: mark custom code state.

### Suggested `forms.inc` changes

Do not scatter raw SQL throughout `forms.inc`. Add a small helper include/class and call it from the existing operation boundaries.

Suggested include:

```asp
<!--#include virtual="/classes/cFormSubmitLedger.inc"-->
```

Suggested object lifecycle near the existing `cRenew` setup:

```asp
dim cLedger
set cLedger = new cFormSubmitLedger
cLedger.AppRenewID = session("APP_RENEW_ID")
cLedger.AppRenewType = appRenewCd
cLedger.MembershipID = session("MEMBERSHIP_ID")
```

The helper should be safe to call even before `session("APP_RENEW_ID")` exists. Once `saveFormData` creates the id, call `cLedger.RefreshFromSession` or set `AppRenewID` again.

#### 1. `saveFormData(formStatus)`

This is the most important anchor because it creates or updates `app_renew`.

After `session("APP_RENEW_ID")` is known, call:

```asp
cLedger.EnsureSubmitRecord formStatus
cLedger.SyncLegacyStatus formStatus
```

Recommended durable-state mapping inside the helper:

```text
PROGRESS -> draft
SAVED    -> draft
PENDING  -> processing, unless an active Awaiting Updates hold exists
APPROVED -> completed, if all configured operations are done
COMPLETED -> completed, if all configured operations are done
DENIED   -> failed or abandoned, depending on local policy
DELETED  -> abandoned
PAY_LATER -> awaiting_payment or abandoned, depending on whether payment is still expected
```

Do not rely only on `formStatus` for durable state. Treat it as one input; operation rows should determine whether the submission is actually complete.

Also call `cLedger.SnapshotConfig` when the user reaches final submit or first transitions out of `draft`. The snapshot should include at least:

- `payment_page`, `payment_processor`, and payment amount rules.
- `save_files_page`, `save_record_page`, `save_pending_page`, and `pdf_page`.
- Form/page PDF settings.
- Form/page email settings.
- `approve_app`, `email_await`, and Awaiting Updates trigger rules when discoverable.
- Page options for `save_record`, `save_pending`, `create_pdf`, `email_pages_onclick`, `record_zero_trans`, `record_batch_trans`, and `execute_code`.

#### 2. Payment routing and `paymentStatus(stat)`

Before redirecting to the configured payment page or gateway:

```asp
cLedger.EnsureOperation "payment", "Y", "form_config", "Configured payment page"
cLedger.SetOperationState "payment", "awaiting", "Redirected to payment gateway"
cLedger.SetDurableState "awaiting_payment"
```

Inside `paymentStatus(stat)`, after the existing `app_renew_detail` writes succeed:

```asp
if stat = "Complete" then
    cLedger.SetOperationState "payment", "done", "Payment confirmed"
    cLedger.SetDurableState "payment_complete"
elseif stat = "Pending" then
    cLedger.SetOperationState "payment", "awaiting", "Payment pending"
    cLedger.SetDurableState "awaiting_payment"
elseif stat = "Failed" then
    cLedger.SetOperationState "payment", "failed", "Payment failed"
    cLedger.SetDurableState "failed"
end if
```

For bypass cases like `payment_ty = K`, record the configured payment operation as not needed for this submission:

```asp
cLedger.EnsureOperation "payment", "Y", "triggered", "Configured payment page bypassed"
cLedger.SetOperationTrigger "payment", "payment_ty", "=", "K"
cLedger.SetOperationState "payment", "not_needed", "Configured payment page was bypassed by alternate Form Pages flow"
```

This should happen where the redirect/page logic determines that the configured payment page will not be used.

#### 3. `cRenew.SaveFiles`

Before calling `cRenew.SaveFiles`:

```asp
cLedger.EnsureOperation "files", "Y", "form_config", "Uploaded files"
cLedger.SetOperationState "files", "running", "Saving uploaded files"
```

After the call, create/update one item per selected uploaded file:

```asp
cLedger.EnsureFileItem fieldName, originalFileName, tempPath, expectedPermanentPath, fileSize
cLedger.SetFileItemState fieldName, originalFileName, "done", permanentPath
```

On failure:

```asp
cLedger.SetFileItemState fieldName, originalFileName, "failed", err.Description
cLedger.SetOperationState "files", "failed", "One or more uploaded files did not persist"
```

Important: optional upload fields with zero selected files should produce `files = not_needed`, not `failed`. If three files are selected in one field, create three operation item rows.

#### 4. `cRenew.SaveData`

Before calling `cRenew.SaveData`:

```asp
cLedger.EnsureOperation "records", "Y", "form_config", "Record updates"
cLedger.SetOperationState "records", "running", "Updating CRM/member records"
```

After success:

```asp
cLedger.SetOperationState "records", "done", "Record updates completed"
```

If the submit flow creates an Awaiting Updates hold, set records to awaiting instead of done:

```asp
cLedger.SetOperationState "records", "awaiting", "Awaiting staff review before record update"
cLedger.SetOperationTrigger "records", triggerField, "=", triggerValue
cLedger.AppendEvent "records", "awaiting_updates_created", "info", "Awaiting Updates hold persisted"
cLedger.SetDurableState "processing"
```

The helper can detect active holds by querying `mem_await` and `mem_det_await` for `approval_dt is null` using the current `app_renew_id`.

#### 5. Awaiting Updates release / approval

This may live outside `forms.inc` in the admin approval path. Wherever the hold is released:

```asp
cLedger.AppendEvent "records", "awaiting_updates_released", "info", "Reviewer released Awaiting Updates hold"
cLedger.SetOperationState "records", "queued", "Record update ready after staff release"
```

If the release action immediately applies record updates, mark `records = done` after that succeeds.

#### 6. `cRenew.PDF`, `cRenew.PDFQuickPost`, and rebuild

Before PDF generation:

```asp
cLedger.EnsureOperation "pdf", "Y", "form_config", "PDF generation"
cLedger.SetOperationState "pdf", "running", "Generating PDF"
```

After success and persistence:

```asp
cLedger.SetOperationState "pdf", "done", "PDF generated and stored"
```

If PDF work is queued to another process, use `queued` first and let the worker set `running` / `done`.

On error:

```asp
cLedger.SetOperationState "pdf", "failed", err.Description
```

#### 7. `cRenew.EmailPage` and `cRenew.EmailPagesOnclick`

Before sending configured form/page emails:

```asp
cLedger.EnsureOperation "emails", "Y", "form_config", "Configured form/page emails"
cLedger.SetOperationState "emails", "running", "Sending configured emails"
```

Create one item per email target/template:

```asp
cLedger.EnsureEmailItem emailKey, recipient, subject, sourceScope
cLedger.SetEmailItemState emailKey, "done", "Email sent"
```

For form types with no email configuration:

```asp
cLedger.EnsureOperation "emails", "N", "not_configured", "No emails configured"
cLedger.SetOperationState "emails", "not_configured", "This form type does not send emails"
```

For triggered email pages, record the trigger:

```asp
cLedger.EnsureEmailItem emailKey, recipient, subject, "triggered"
cLedger.SetEmailItemTrigger emailKey, "payment_ty", "=", "K"
```

#### 8. Batch / billing transaction helpers

Wrap calls such as `batchTransaction`, `batchTransactionPending`, and `record_batch_trans` handling:

```asp
cLedger.EnsureOperation "billing", "Y", "form_config", "Billing history"
cLedger.SetOperationState "billing", "running", "Recording billing transaction"
```

After success, mark `done`. If the form type has no billing, ensure `billing = not_configured` so the admin grid shows NA instead of a blank ambiguous value.

#### 9. Page `execute_code`

Before `execute cRenew.PageOption("execute_code")`:

```asp
cLedger.EnsureOperation "custom_code", "Y", "form_config", "Page custom code"
cLedger.SetOperationState "custom_code", "running", "Executing configured page code"
```

After success:

```asp
cLedger.SetOperationState "custom_code", "done", "Custom code completed"
```

On error, capture the error and keep the operation retry policy explicit. Some custom code may not be safe to auto-retry unless it is known idempotent.

#### 10. End-of-submit reconciliation in `forms.inc`

Near the end of a submit request, after page operations have run, call:

```asp
cLedger.RecomputeSubmitState
```

The helper should set:

- `completed` when every configured/triggered operation is `done`, `not_needed`, or `not_configured`.
- `processing` when any operation is `ready`, `queued`, `running`, `pending`, `retrying`, or `awaiting`.
- `failed` when any required operation is `failed` and not retryable.
- `awaiting_payment` when payment is required and still awaiting gateway result.

Do not let `doPDF=1` or `doEmail=1` debug/render modes create duplicate submit records. Those paths should append events only when they are intentionally rebuilding/re-sending for an existing `app_renew_id`.

### Phase 4: Add recovery worker

Goal: make "Replay Pipeline" safe.

- Select records in `processing`, `payment_complete`, `failed` with retryable operations, or `awaiting_payment` past a reconciliation threshold.
- For each operation, decide whether it is complete, retryable, waiting, or terminal.
- Resume only idempotent steps.
- Use operation items for granular retry, especially failed file persistence and failed emails.
- Append every recovery decision to `form_submit_event`.

### Phase 5: Replace mock data with real admin data

Goal: make `apprenewadmin.asp` show the mockup's state using real schema.

- List view reads `form_submit_record` plus summarized `form_submit_operation` rows.
- Detail view reads all operations, items, and events for the selected `app_renew_id`.
- Existing columns can remain, but their meaning should be tightened:
  - State: durable execution state.
  - Status: legacy/admin `app_renew_status`.
  - Awaiting Updates: active persisted hold now.
  - Review: derived review status (`Auto`, `Awaiting`, `Approved`).
  - Payment: list shows `Complete ($)` / `Pending ($)` / `Not paid` / `Failed` / `NA` from the `payment` operation row; Payment Date-Time from legacy `app_renew_detail` or operation `completed_at`.
  - PDF/Files/Emails: operation summaries.
- Keep the old admin actions (approve, pend, deny, delete, rebuild PDF) but add safe recovery actions that operate through the new helper API.

## Open Decisions

- Table naming: use a neutral prefix like `form_submit_*`, or align to legacy with `app_renew_submit_*`.
- Payload storage: Oracle `clob` JSON is flexible; separate normalized columns should still exist for commonly queried status fields.
- Job queue ownership: decide whether PDF/email/file retry uses an existing queue or a new recovery worker.
- Awaiting Updates history: confirm whether released rows remain in `mem_await` / `mem_det_await` with `approval_dt`; if not, add an event row when released.
- Approval wording: resolved — the mockup and admin patch use **Review** for the derived column (`Auto` / `Awaiting` / `Approved`), keeping legacy **Status** for `app_renew_status` values like `APPROVED`.

## Acceptance Criteria

- A PM can switch Form Type and see consistent form-level expectations for all submissions of that type.
- A submission with no payment shows Payment as `NA` on the list (with reason in detail).
- A submission with payment completed shows Payment as `Complete ($amount)` on the list.
- A submission awaiting gateway or payment confirmation shows Payment as `Pending ($amount)` on the list.
- A submission with payment bypassed shows the triggering field/value and does not wait on gateway reconciliation.
- Optional upload fields with no files do not look failed.
- Multiple files in one upload field are tracked individually.
- If one file remains in temp storage, the grid and detail view show exactly which file needs retry.
- Emails are `NA` for form types with no email config.
- Form-config emails are consistent across submissions of the same form type.
- Triggered emails show the field condition that caused them.
- Awaiting Updates `Yes` means the hold exists now and was persisted.
- Review status distinguishes `Auto`, `Awaiting`, and `Approved`.
- Recovery can resume incomplete work using durable operation state, without rerunning completed items unnecessarily.
