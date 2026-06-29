# Submit Record Enhancement
## Current Legacy Admin Fields

- Values

- Payment

- Reconcile
    
- Status

- Uploaded Files

- Update Records

- Awaiting Updates

- Review / Auto (derived; not legacy `app_renew_status = APPROVED`)

- Parents, Children, Grandchildren, Sibling, Other, Grandparent completed

- PDF

- Billing History

- Emails

- Custom Code

## Clarifications From Mockup Review

- Form Type controls the expected form-level configuration. Submissions of the same form code should not appear to have different form-level setup.
- Per-submission field values can still trigger behavior, such as payment bypass, a special receipt page, or Awaiting Updates.
- CONTACT examples should show no payment and no emails. Name-change submissions require uploaded proof and can trigger Awaiting Updates; email/phone-only changes should not.
- Awaiting Updates `Yes` should mean an actual persisted unreleased hold exists now, not merely that the form is configured to potentially hold.
- Payment list column must answer “was payment completed?” at a glance, using legacy-admin labels:
  - `Complete ($amount)` when `ops.payment.state = done`
  - `Pending ($amount)` when `awaiting` or `pending` (not completed)
  - `Not paid ($amount)` when `ready` (required but not started)
  - `Failed ($amount)` when `failed`
  - `NA` when not required, not configured, or bypassed (e.g. `payment_ty = K`)
- The Review column (renamed from "Approval" in the mockup) is derived review status:
  - `Auto`: no staff review needed.
  - `Awaiting`: active review hold now.
  - `Approved`: staff reviewed and released the hold.
- The legacy `app_renew_status` is not enough to explain durable execution. A separate durable state is needed: `draft`, `awaiting_payment`, `payment_complete`, `processing`, `completed`, `failed`, `abandoned`.

## Reconciled Implementation Notes

- `forms.inc` already creates/updates the `app_renew` row and calls operation hooks for payment, files, records, PDF, emails, billing/batch transactions, and custom code.
- `apprenewadmin.asp` already shows coarse values from `app_renew`, `app_renew_detail`, `mem_await` / `mem_det_await`, and `form_pdf`.
- The required enhancement is a durable operation ledger around those hooks, with operation-level and item-level states.
- Operation `state` values in the mockup, DDL (`ck_fso_state`), and helper must stay aligned. Display labels may differ by surface:
  - **List Payment column:** `Complete ($)`, `Pending ($)`, `Not paid`, `Failed`, `NA` (legacy-admin scan pattern).
  - **Detail payment card:** may use richer labels (e.g. canonical `awaiting` + card → “Awaiting gateway”).
- See `durable_submission_schema_plan.md` for the proposed schema and rollout plan.

## Canonical operation vocabulary

Shared across `form_submit_tool.html` (`OP_STATE`), `patched/db/form_submit_ledger.sql`, and `cFormSubmitLedger.inc`:

`not_configured`, `not_needed`, `ready`, `queued`, `running`, `pending`, `retrying`, `awaiting`, `done`, `failed`

Do not store display-only states such as `awaiting_gateway` or `not_required` in the database.

## List view column map (mockup)

| Column | Source | PM-facing values |
|--------|--------|------------------|
| State | `form_submit_record.durable_state` | `draft`, `processing`, `completed`, … |
| Status | `app_renew.app_renew_status` | Pending, Completed, In Progress, … |
| Payment | `form_submit_operation` where `operation_key = payment` | `Complete ($)`, `Pending ($)`, `Not paid`, `Failed`, `NA` |
| Review | derived from `records` op + `mem_await` | `Auto`, `Awaiting`, `Approved` |
| Awaiting Updates | active `mem_await` / `mem_det_await` hold | `Yes` / `No` |