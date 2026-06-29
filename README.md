# Form Submit Record

Mockup and design artifacts for a durable, recoverable form submission pipeline in Form Manager.

## Contents

| File | Description |
|------|-------------|
| `form_submit_tool.html` | Interactive admin mockup for submission record management |
| `form_submit_tool-legacy.html` | Earlier version of the mockup |
| `_check.js` | Quick Node script to validate grid rendering in the mockup |
| `docs/readme.md` | Mermaid flowchart of the durable submit pipeline |
| `docs/chat.md` | Design discussion notes |
| `docs/meeting_notes.md` | Feature checklist from planning sessions |
| `docs/durable_submission_schema_plan.md` | Proposed schema and implementation plan reconciled against `forms.inc` and `apprenewadmin.asp` |
| `patched/` | Local patched copies of `forms.inc`, `apprenewadmin.asp`, `cFormSubmitLedger.inc`, and Oracle DDL (not deployed) |

## Viewing the mockup

Open `form_submit_tool.html` in a browser. No build step required.

The mockup is intentionally richer than the current global ASP admin tool. It shows the desired durable submission record: form-type-specific expected configuration, per-operation execution state, per-item file/email tracking, Awaiting Updates review state, and recovery actions.

**List view highlights** (what PMs scan first):

- **Payment** — legacy-admin wording: `Complete ($)`, `Pending ($)`, `Not paid`, `Failed`, or `NA` (maps from canonical `ops.payment.state`: `done`, `awaiting`/`pending`, `ready`, `failed`, `not_configured`/`not_needed`).
- **Review** — derived from the active hold: `Auto`, `Awaiting`, or `Approved` (not the legacy `app_renew_status = APPROVED` Status column).
- **State** — durable execution lifecycle (`draft`, `processing`, `completed`, etc.), separate from payment completion.

## Current Design Status

The design has been reconciled against the current read-only global files:

- `a:\GLOBAL_6-next\www\includes\forms.inc` - front-end form processor.
- `a:\GLOBAL_6-next\admin\apprenewadmin.asp` - current admin tool for viewing submissions.

Those global files already contain the main operation boundaries (`SaveFiles`, `SaveData`, `PDF`, email page hooks, payment status writes, Awaiting Updates checks). The missing layer is durable operation-level intent/execution tracking. See `docs/durable_submission_schema_plan.md` for the proposed schema and rollout plan.

## Validate grid rendering

```bash
node _check.js
```
