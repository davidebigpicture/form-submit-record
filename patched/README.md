# Patched reference files (local only)

These are **local, patched copies** of two global files plus the new
dependencies they require. They implement the durable submission ledger /
observability layer described in
[`../docs/durable_submission_schema_plan.md`](../docs/durable_submission_schema_plan.md)
and visualized in [`../form_submit_tool.html`](../form_submit_tool.html).

> **Do NOT treat these as the source of truth for the globals.** The originals
> in `a:\GLOBAL_6-next\...` were used as a read-only reference and were **not**
> edited. Diff these copies against the globals before deploying anything.

## Files

| File | Origin | What changed |
|------|--------|--------------|
| `www/includes/forms.inc` | copy of `a:\GLOBAL_6-next\www\includes\forms.inc` | Instrumented at existing operation boundaries to record durable submission state. |
| `admin/apprenewadmin.asp` | copy of `a:\GLOBAL_6-next\admin\apprenewadmin.asp` | Adds **Submission State / Emails / Review** columns sourced from the ledger. |
| `classes/cFormSubmitLedger.inc` | **new** | Idempotent, fail-safe helper class that writes the ledger. |
| `db/form_submit_ledger.sql` | **new** | Oracle DDL aligned with the plan doc: tables, sequences, indexes, FKs, `v_form_submit_admin` view. |

## Canonical vocabulary

Operation `state` values are shared across the mockup (`OP_STATE`), `db/form_submit_ledger.sql` (`ck_fso_state`), and `cFormSubmitLedger.inc`. UI display labels may differ from stored values (e.g. canonical `awaiting` with `method: "card"` renders as "Awaiting gateway"). Do not persist display-only states like `awaiting_gateway` or `not_required`.

Every change in the two ASP files is wrapped in clearly marked blocks:

```
' === FORM SUBMIT LEDGER PATCH ===
...
' === END FORM SUBMIT LEDGER PATCH ===
```

Search for `FORM SUBMIT LEDGER PATCH` to review every edit.

## Design guarantees

- **Never breaks the form.** Every DB call in `cFormSubmitLedger` runs under
  `on error resume next`. If the schema is missing (migration not yet applied),
  the helper silently no-ops. The form processor behaves exactly as before.
- **Idempotent.** All writes are `MERGE`/`UPDATE` keyed by `app_renew_id`
  (+ `operation_key` / `item_key`), so refreshes, the `doPDF`/`doEmail` passes,
  redirects, and recovery re-runs are safe.
- **Additive only.** No existing table is altered. The admin tool keeps working
  unchanged if `db/form_submit_ledger.sql` has not been run (the new columns are
  hidden until `form_submit_record` exists).
- **CRLF + tab formatting preserved** to match the originals (verified).

## Deploy order

1. Apply the schema: run `db/form_submit_ledger.sql` in the client Oracle schema.
2. Deploy `classes/cFormSubmitLedger.inc` to `/classes/`.
3. Deploy the patched `www/includes/forms.inc` (writes the ledger).
4. Deploy the patched `admin/apprenewadmin.asp` (reads the ledger).

Steps 2-4 are safe to deploy before step 1; the columns simply stay empty/hidden
until the tables exist.

## What `forms.inc` records (instrumentation points)

| Boundary in `forms.inc` | Ledger call |
|-------------------------|-------------|
| Object setup (top) | `set cLedger = new cFormSubmitLedger` |
| End of `saveFormData(formStatus)` | `EnsureSubmitRecord formStatus` |
| End of `paymentStatus(stat)` | `RecordPaymentStatus stat` (maps Complete/Pending/Failed) |
| Page op `save_record` / `save_record_page` | `records` running -> `MarkRecordsAfterSave` (done, or awaiting if an Awaiting Updates hold exists) |
| Page op `save_files_page` | `files` running -> `SyncFilesFromSession` (boundary record) |
| Page op `email_field` | `emails` running -> done |
| Page op `create_pdf` / `pdf_page` | `pdf` running -> done |
| Page op `record_zero_trans` / `record_batch_trans` | `billing` running -> done |
| Page `execute_code` | `custom_code` running -> done |
| End of page operations | `RecomputeSubmitState` |

## Known limitations / next phase

- **Per-file items.** `cAppRenew.SaveFiles` is a global class we do not edit, so
  the patch records `files` at the operation level via `SyncFilesFromSession`.
  True per-file rows (the "3 of 3 persisted, 1 stuck in temp" view) require
  hooks **inside** `SaveFiles` or enumeration of the session upload arrays.
  `cFormSubmitLedger` already exposes `EnsureFileItem` / `SetFileItemState` for
  when that hook is added.
- **Per-email items.** Same pattern: `EnsureEmailItem` / `SetEmailItemState`
  exist; the current patch records `emails` at the operation level.
- **Payment bypass (`payment_ty = K`).** `MarkPaymentBypassed` is implemented;
  wire it where the alternate Form Pages flow decides to skip `pay.asp`.
- **Awaiting Updates release.** `apprenewadmin.asp` approve/deny handlers can
  call `AppendEvent "records","awaiting_updates_released",...` +
  `SetOperationState "records","done"/"queued"` to drive the **Review** column to
  `Approved`. Not wired here to keep the admin diff read-only.
- **Config snapshot.** `SnapshotConfig(jsonText)` caps payloads at ~3900 chars
  (plain SQL literal limit); large snapshots should bind a CLOB later.
- **Recovery worker** (Phase 4) is out of scope for these files.
