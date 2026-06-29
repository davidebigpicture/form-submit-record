## Final Form Submit Flow (Durable + Recoverable)

This flow is the target architecture shown in `form_submit_tool.html`. It has been reconciled with the current global ASP files:

- `forms.inc` already creates/updates the `app_renew` row through `saveFormData(formStatus)`.
- `forms.inc` already calls the main operation hooks: payment status, `SaveFiles`, `SaveData`, `PDF`, email page sends, batch transactions, and custom page code.
- `apprenewadmin.asp` already reads the legacy submission list from `app_renew`, `app_renew_detail`, `mem_await` / `mem_det_await`, and `form_pdf`.

The new work is to add a durable operation ledger around those existing hooks, not to replace the entire form processor.

```mermaid
flowchart TD

    A[Start Form<br>Create or reuse app_renew_id] --> B[User fills pages<br>Redis / Temp / app_renew_page_submit]

    B --> C{Payment Required?}

    %% ========================
    %% PAYMENT PATH
    %% ========================
    C -->|Yes| D[Mark payment operation awaiting<br>Redirect to Payment]

    D --> E[Payment Gateway]

    %% success or uncertain (crash)
    E --> F[Return from Payment]

    F --> G[Mark payment done<br>payment_complete = true]

    %% reconcile path
    F --> H[Recovery Job Checks Submission]
    H --> I{Payment Known?}

    I -->|Yes| G
    I -->|No| J[Reconcile with Gateway]
    J --> K{Was Payment Captured?}

    K -->|Yes| G
    K -->|No| Z[Mark Abandoned<br>(after timeout)]

    %% ========================
    %% NO PAYMENT PATH
    %% ========================
    C -->|No| L[Final Certify / Submit Page]

    L --> M[Mark Submission<br>ready_to_process = true]

    %% ========================
    %% COMMON PIPELINE ENTRY
    %% ========================
    G --> N[Begin Processing Pipeline]
    M --> N

    %% ========================
    %% PIPELINE STEPS
    %% ========================
    N --> O[Save Uploaded Files<br>(idempotent + per-file tracked)]

    O --> P{Awaiting Updates<br>Triggered?}

    P -->|Yes| P1[Persist review hold<br>mem_await / mem_det_await]
    P1 --> P2[Wait for staff release]
    P2 --> Q[Update CRM Record<br>(transactional)]

    P -->|No| Q[Update CRM Record<br>(transactional)]

    Q --> R[Generate PDF]

    R --> S[Create Billing History]

    S --> T[Send Emails]

    T --> U[Run Custom Code]

    U --> V[Mark Submission Complete]

    %% ========================
    %% RECOVERY LOOP
    %% ========================
    H --> W{Submission Incomplete?}

    W -->|Yes| N
    W -->|No| Y[No Action]

    %% ========================
    %% NOTES
    %% ========================
    Z --> X[End]
    V --> X
    Y --> X
```

## Legacy Mapping

- `app_renew.app_renew_status` remains the legacy/admin status (`PROGRESS`, `PENDING`, `APPROVED`, `COMPLETED`, etc.).
- The proposed durable state is separate: `draft`, `awaiting_payment`, `payment_complete`, `processing`, `completed`, `failed`, `abandoned`.
- Payment currently persists to `app_renew_detail` fields `PAYMENT_STATUS` and `PAYMENT_DATE_TIME`; the new schema should mirror that into a `payment` operation row.
- Awaiting Updates currently derives from active rows in `mem_await` / `mem_det_await`; the new schema should treat active unreleased rows as an actual `records = awaiting` state.
- PDF currently derives from `form_pdf.status`; the new schema should track expected / queued / running / done / failed around PDF generation.
- Files, emails, billing history, and custom code need explicit operation rows because the current admin list does not persist enough execution detail to support retry and PM-friendly status.

See `durable_submission_schema_plan.md` for the proposed schema and phased implementation plan.