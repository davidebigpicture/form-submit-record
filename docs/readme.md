## Final Form Submit Flow (Durable + Recoverable)

```mermaid
flowchart TD

    A[Start Form<br>Create SubmissionId] --> B[User fills pages<br>Redis / Temp]

    B --> C{Payment Required?}

    %% ========================
    %% PAYMENT PATH
    %% ========================
    C -->|Yes| D[Redirect to Payment]

    D --> E[Payment Gateway]

    %% success or uncertain (crash)
    E --> F[Return from Payment]

    F --> G[Mark Submission<br>payment_complete = true]

    %% reconcile path
    F --> H[Recovery Job Checks Submission]
    H --> I{Payment Known?}

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
    N --> O[Save Uploaded Files<br>(idempotent)]

    O --> P[Update CRM Record<br>(transactional)]

    P --> Q[Generate PDF]

    Q --> R[Create Billing History]

    R --> S[Send Emails]

    S --> T[Run Custom Code]

    T --> U[Mark Submission Complete]

    %% ========================
    %% RECOVERY LOOP
    %% ========================
    H --> V{Submission Incomplete?}

    V -->|Yes| N
    V -->|No| W[No Action]

    %% ========================
    %% NOTES
    %% ========================
    Z --> X[End]
    U --> X
```