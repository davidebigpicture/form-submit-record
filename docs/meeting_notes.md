# Submit Record Enhancement
## Current Legacy Admin Fields

- Values

- Payment

- Reconcile
    
- Status

- Uploaded Files

- Update Records

- Awaiting Updates

- Approval / Auto

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
- The Approval column is better understood as Review status:
  - `Auto`: no staff review needed.
  - `Awaiting`: active review hold now.
  - `Approved`: staff reviewed and released the hold.
- The legacy `app_renew_status` is not enough to explain durable execution. A separate durable state is needed: `draft`, `awaiting_payment`, `payment_complete`, `processing`, `completed`, `failed`, `abandoned`.

## Reconciled Implementation Notes

- `forms.inc` already creates/updates the `app_renew` row and calls operation hooks for payment, files, records, PDF, emails, billing/batch transactions, and custom code.
- `apprenewadmin.asp` already shows coarse values from `app_renew`, `app_renew_detail`, `mem_await` / `mem_det_await`, and `form_pdf`.
- The required enhancement is a durable operation ledger around those hooks, with operation-level and item-level states.
- See `durable_submission_schema_plan.md` for the proposed schema and rollout plan.