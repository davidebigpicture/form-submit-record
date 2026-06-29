
        // Sample records modeled after the legacy list, extended with the durable submission
        // record fields called out in docs/meeting_notes.md.
        const records = [
            {
                id: 13148,
                name: "Gregory King",
                membershipId: "4018144",
                state: "processing",
                status: "Pending",
                statusTone: "warn",
                stage: "Awaiting processing",
                stageTone: "warn",
                paymentDate: "6/6/2026 7:46:22 PM",
                expiration: "6/30/2027",
                created: "6/6/2026 7:45:37 PM",
                lastModified: "6/6/2026 7:55:58 PM",
                lastSubmitted: "6/6/2026 7:47:22 PM",
                ops: {
                    payment: { required: true, method: "card", amount: "$100.00", state: "done", note: "Card charged and confirmed by the gateway." },
                    files: { fields: [{ label: "Continuing education certificate", required: false, max: 5, uploaded: 0, persisted: 0, items: [] }] },
                    pdf: { configured: true, state: "running", note: "Building now — the job is running in the queue." },
                    emails: [
                        { name: "Payment receipt", to: "gregory.king@example.com", trigger: "After payment", configured: true, state: "done" },
                        { name: "Signed PDF copy", to: "records@wvpsboard.gov", trigger: "After PDF", configured: true, state: "queued" },
                    ],
                    records: { configured: true, state: "pending", note: "Updating the licensure database for this applicant." },
                    billing: { configured: true, state: "queued", note: "Billing history entry is queued behind the records update." },
                    customCode: { configured: true, state: "queued", note: "Post-renewal script runs at the end of the pipeline." },
                },
                booleans: {
                    filesSaved: false,
                    updateRecords: false,
                    awaitingUpdates: false,
                    approval: false,
                    auto: true,
                    pdfJob: false,
                    billingHistory: false,
                    emails: false,
                    customCode: false,
                    parents: true,
                    children: false,
                    grandchildren: null,
                    sibling: null,
                    other: null,
                    grandparent: true,
                    readyToProcess: true,
                },
                steps: ["done", "done", "done", "pending", "pending", "pending"],
                stepsLabel: "3 / 6",
                updateStats: { pending: true },
                pipeline: [
                    ["7:45 PM", "Submission record created", "Durable row reserved before final certify."],
                    ["7:46 PM", "Payment captured", "Card charged and confirmed by the gateway."],
                    ["7:47 PM", "User certified / submitted", "ready_to_process set true — pipeline started."],
                    ["7:47 PM", "PDF build started", "Running in the job queue."],
                    ["—", "CRM updated", "In progress."],
                    ["—", "Emails + billing + custom code", "Queued behind the records update."],
                ],
                recovery: { tone: "warn", text: "Processing did not finish after certify. Recovery worker will resume the pipeline.", action: "Resume pipeline" },
            },
            {
                id: 13142,
                name: "(no name on record)",
                membershipId: "0",
                state: "draft",
                status: "In Progress",
                statusTone: "neutral",
                stage: "Draft / no certify yet",
                stageTone: "neutral",
                paymentDate: "",
                expiration: "",
                created: "6/6/2026 3:07:33 PM",
                lastModified: "6/6/2026 3:31:14 PM",
                lastSubmitted: "6/6/2026 3:31:15 PM",
                ops: {
                    payment: { required: true, method: "card", amount: "$100.00", state: "ready", note: "Payment will be requested after the applicant submits the form." },
                    files: { fields: [{ label: "Continuing education certificate", required: false, max: 5, uploaded: 0, persisted: 0, items: [] }] },
                    pdf: { configured: true, state: "ready", note: "Will build once the applicant submits." },
                    emails: [
                        { name: "Payment receipt", to: "(applicant email)", trigger: "After payment", configured: true, state: "ready" },
                        { name: "Signed PDF copy", to: "records@wvpsboard.gov", trigger: "After PDF", configured: true, state: "ready" },
                    ],
                    records: { configured: true, state: "ready", note: "Runs after the applicant submits." },
                    billing: { configured: true, state: "ready", note: "Runs after the applicant submits." },
                    customCode: { configured: true, state: "ready", note: "Post-renewal script runs at the end of the pipeline." },
                },
                booleans: {
                    filesSaved: false,
                    updateRecords: false,
                    awaitingUpdates: false,
                    approval: false,
                    auto: false,
                    pdfJob: false,
                    billingHistory: false,
                    emails: false,
                    customCode: false,
                    parents: null,
                    children: null,
                    grandchildren: null,
                    sibling: null,
                    other: null,
                    grandparent: null,
                    readyToProcess: false,
                },
                steps: ["done", "pending", "pending", "pending", "pending", "pending"],
                stepsLabel: "1 / 6",
                updateStats: { pending: true },
                pipeline: [
                    ["3:07 PM", "Submission record created", "Durable row exists even though user has not certified."],
                    ["—", "Pages persisted + certified", "User is still editing."],
                    ["—", "Files / CRM / PDF / notifications", "Will start after certify or payment."],
                ],
                recovery: { tone: "warn", text: "Draft is older than 24 hours — safe to abandon after timeout.", action: "Mark abandoned" },
            },
            {
                id: 13139,
                name: "Aaron Teets",
                membershipId: "4017040",
                state: "awaiting_payment",
                status: "In Progress",
                statusTone: "neutral",
                stage: "Payment pending",
                stageTone: "warn",
                paymentDate: "",
                expiration: "6/30/2026",
                created: "6/5/2026 5:13:12 PM",
                lastModified: "6/5/2026 5:14:07 PM",
                lastSubmitted: "6/5/2026 5:14:08 PM",
                ops: {
                    payment: { required: true, method: "card", amount: "$100.00", state: "awaiting", note: "Redirected to the card gateway 18 hours ago with no result returned. Reconcile to confirm whether the card was actually charged." },
                    files: { fields: [{ label: "Continuing education certificate", required: false, max: 5, uploaded: 2, persisted: 2, items: [{ name: "ce_hours_2026.pdf", size: "402 KB", state: "done" }, { name: "ce_hours_ethics.pdf", size: "118 KB", state: "done" }] }] },
                    pdf: { configured: true, state: "ready", note: "Held until the payment result is known." },
                    emails: [
                        { name: "Payment receipt", to: "aaron.teets@example.com", trigger: "After payment", configured: true, state: "ready" },
                        { name: "Signed PDF copy", to: "records@wvpsboard.gov", trigger: "After PDF", configured: true, state: "ready" },
                    ],
                    records: { configured: true, state: "ready", note: "Held until payment is confirmed." },
                    billing: { configured: true, state: "ready", note: "Held until payment is confirmed." },
                    customCode: { configured: true, state: "ready", note: "Post-renewal script runs at the end of the pipeline." },
                },
                booleans: {
                    filesSaved: true,
                    updateRecords: false,
                    awaitingUpdates: true,
                    approval: false,
                    auto: false,
                    pdfJob: false,
                    billingHistory: false,
                    emails: false,
                    customCode: false,
                    parents: true,
                    children: false,
                    grandchildren: null,
                    sibling: null,
                    other: null,
                    grandparent: true,
                    readyToProcess: false,
                },
                steps: ["done", "done", "done", "blocked", "pending", "pending"],
                stepsLabel: "3 / 6",
                updateStats: { pending: true },
                pipeline: [
                    ["5:13 PM", "Submission record created", "Durable row before payment redirect."],
                    ["5:13 PM", "Files saved", "Uploads persisted."],
                    ["5:14 PM", "Redirected to payment gateway", "Submission id and token sent with redirect."],
                    ["—", "Payment callback", "No return seen — reconcile job armed."],
                    ["—", "CRM + PDF + emails", "Held until payment certainty."],
                ],
                recovery: { tone: "bad", text: "Payment return missing for 18 hours. Reconcile required before pipeline can resume.", action: "Reconcile with gateway" },
            },
            {
                id: 13155,
                name: "Diane Whitfield",
                membershipId: "4019870",
                state: "processing",
                status: "In Progress",
                statusTone: "neutral",
                stage: "Processing",
                stageTone: "warn",
                paymentDate: "",
                expiration: "6/30/2027",
                created: "6/14/2026 11:02:09 AM",
                lastModified: "6/14/2026 11:08:44 AM",
                lastSubmitted: "6/14/2026 11:05:20 AM",
                booleans: {
                    readyToProcess: true,
                    awaitingUpdates: false,
                    approval: false,
                    auto: true,
                    parents: true,
                    children: true,
                    grandchildren: false,
                    sibling: null,
                    other: null,
                    grandparent: true,
                },
                steps: ["done", "done", "done", "pending", "pending", "pending"],
                stepsLabel: "3 / 6",
                updateStats: { pending: true },
                ops: {
                    payment: { required: true, bypassed: true, bypassField: "payment_ty", bypassValue: "K", bypassPage: "pay.asp", state: "not_needed" },
                    files: { fields: [
                        { label: "Continuing education certificates", required: false, max: 5, uploaded: 3, persisted: 3, items: [{ name: "ce_block_a.pdf", size: "512 KB", state: "done" }, { name: "ce_block_b.pdf", size: "288 KB", state: "done" }, { name: "ce_ethics.pdf", size: "96 KB", state: "done" }] },
                        { label: "License card photo", required: false, max: 1, uploaded: 1, persisted: 1, items: [{ name: "license_front.jpg", size: "174 KB", state: "done" }] },
                    ] },
                    pdf: { configured: true, state: "running", note: "Building now — the configured payment page was bypassed, so the pipeline runs without waiting on a gateway." },
                    emails: [
                        { name: "Payment receipt", to: "diane.whitfield@example.com", trigger: "After payment", configured: true, state: "not_needed", note: "no payment was taken" },
                        { name: "Check Receipt page", to: "diane.whitfield@example.com", trigger: "On submit (alternate page)", configured: true, state: "done", why: { source: "trigger", text: "payment_ty = K — the Check Receipt email page is sent instead of the card receipt." } },
                        { name: "Signed PDF copy", to: "records@wvpsboard.gov", trigger: "After PDF", configured: true, state: "queued" },
                    ],
                    records: { configured: true, state: "pending", note: "Updating the licensure database." },
                    billing: { configured: true, state: "queued", note: "Queued behind the records update." },
                    customCode: { configured: true, state: "queued", note: "Post-renewal script runs at the end of the pipeline." },
                },
                pipeline: [
                    ["11:02 AM", "Submission record created", "Durable row reserved."],
                    ["11:04 AM", "Files saved", "All 4 uploads persisted to permanent storage."],
                    ["11:05 AM", "Submitted via alternate page", "Configured payment page (pay.asp) bypassed because payment_ty = K — no gateway involved."],
                    ["11:05 AM", "PDF build started", "Running in the job queue."],
                    ["—", "CRM + emails + billing", "Processing normally — nothing is gated on payment."],
                ],
                recovery: { tone: "warn", text: "Alternate payment flow (payment_ty = K): the configured payment page (pay.asp) was bypassed, so no gateway is involved. The pipeline is processing normally.", action: "Resume pipeline" },
            },
            {
                id: 13136,
                name: "Aaron Smith",
                membershipId: "4018551",
                state: "processing",
                status: "In Progress",
                statusTone: "neutral",
                stage: "Awaiting updates",
                stageTone: "warn",
                paymentDate: "6/3/2026 10:48:12 PM",
                expiration: "6/30/2026",
                created: "6/3/2026 10:48:12 PM",
                lastModified: "6/3/2026 10:59:18 PM",
                lastSubmitted: "6/3/2026 10:59:19 PM",
                ops: {
                    payment: { required: true, method: "card", amount: "$100.00", state: "done", note: "Card charged and confirmed by the gateway." },
                    files: { fields: [{ label: "Continuing education certificate", required: false, max: 5, uploaded: 2, persisted: 2, items: [{ name: "ce_hours_2026.pdf", size: "360 KB", state: "done" }, { name: "ethics_cert.pdf", size: "122 KB", state: "done" }] }] },
                    pdf: { configured: true, state: "ready", note: "Held until the staff review is released." },
                    emails: [
                        { name: "Payment receipt", to: "aaron.smith@example.com", trigger: "After payment", configured: true, state: "done" },
                        { name: "Signed PDF copy", to: "records@wvpsboard.gov", trigger: "After PDF", configured: true, state: "ready" },
                    ],
                    records: {
                        configured: true, state: "pending",
                        note: "Paused — the record update has not been applied yet.",
                        awaiting: { source: "trigger", text: "felony_record = Yes — a reviewer must approve before the renewal updates the record." },
                    },
                    billing: { configured: true, state: "ready", note: "Runs after the record update is released." },
                    customCode: { configured: true, state: "ready", note: "Runs after the record update is released." },
                },
                booleans: {
                    filesSaved: true,
                    updateRecords: true,
                    awaitingUpdates: true,
                    approval: false,
                    auto: false,
                    pdfJob: false,
                    billingHistory: true,
                    emails: false,
                    customCode: false,
                    parents: true,
                    children: true,
                    grandchildren: false,
                    sibling: false,
                    other: null,
                    grandparent: true,
                    readyToProcess: true,
                },
                steps: ["done", "done", "done", "pending", "pending", "pending"],
                stepsLabel: "3 / 6",
                updateStats: { pending: true },
                pipeline: [
                    ["10:48 PM", "Submission record created", "Durable row before payment."],
                    ["10:50 PM", "Payment captured", "Callback marked payment_complete = true."],
                    ["10:55 PM", "Files saved", "Both uploads persisted."],
                    ["10:55 PM", "Awaiting Updates triggered", "felony_record = Yes — record update paused for staff review."],
                    ["—", "Record update + billing", "Held until a reviewer releases the submission."],
                    ["—", "PDF + emails + custom code", "Held behind the record update."],
                ],
                recovery: { tone: "warn", text: "Awaiting Updates was triggered (felony_record = Yes). A reviewer must release the submission before the record is updated and the pipeline continues.", action: "Open review queue" },
            },
            {
                id: 13117,
                name: "Ambrose Gmeiner",
                membershipId: "4016804",
                state: "completed",
                status: "Completed",
                statusTone: "ok",
                stage: "Closed",
                stageTone: "ok",
                issues: [
                    { severity: "bad", text: "File save failure: 1 of 3 continuing-education files (ce_block_c.pdf) is still stuck in temporary storage and was never persisted to permanent storage. Payment was captured and downstream steps ran on an incomplete file set — this upload MUST be retried before the submission can be considered durable." },
                ],
                paymentDate: "5/20/2026 4:37:58 PM",
                expiration: "6/30/2026",
                created: "5/20/2026 4:34:49 PM",
                lastModified: "5/20/2026 4:37:57 PM",
                lastSubmitted: "5/20/2026 4:37:12 PM",
                ops: {
                    payment: { required: true, method: "card", amount: "$100.00", state: "done", note: "Card charged and confirmed." },
                    files: { fields: [
                        { label: "Continuing education certificates", required: false, max: 5, uploaded: 3, persisted: 2, items: [
                            { name: "ce_block_a.pdf", size: "385 KB", state: "done" },
                            { name: "ce_block_b.pdf", size: "241 KB", state: "done" },
                            { name: "ce_block_c.pdf", size: "198 KB", state: "failed", error: "Still in temporary storage — never persisted to permanent storage. Needs a retry." },
                        ] },
                        { label: "License card photo", required: false, max: 1, uploaded: 1, persisted: 1, items: [{ name: "license_card_front.jpg", size: "180 KB", state: "done" }] },
                        { label: "Supporting documents", required: false, max: 5, uploaded: 0, persisted: 0, items: [] },
                    ] },
                    pdf: { configured: true, state: "done", note: "Built and stored." },
                    emails: [
                        { name: "Payment receipt", to: "ambrose.g@example.com", trigger: "After payment", configured: true, state: "done" },
                        { name: "Signed PDF copy", to: "records@wvpsboard.gov", trigger: "After PDF", configured: true, state: "retrying", attempts: "attempt 2 of 5", note: "mailbox temporarily rejected the message" },
                    ],
                    records: { configured: true, state: "done", note: "Licensure database updated." },
                    billing: { configured: true, state: "done", note: "Billing history entry created." },
                    customCode: { configured: true, state: "done", note: "Post-submit script ran." },
                },
                booleans: {
                    filesSaved: false,
                    updateRecords: true,
                    awaitingUpdates: false,
                    approval: true,
                    auto: true,
                    pdfJob: true,
                    billingHistory: true,
                    emails: true,
                    customCode: true,
                    parents: true,
                    children: true,
                    grandchildren: true,
                    sibling: true,
                    other: true,
                    grandparent: true,
                    readyToProcess: true,
                },
                steps: ["done", "done", "blocked", "done", "done", "done"],
                stepsLabel: "5 / 6 • 1 failed",
                updateStats: { updated: 3, inserted: 0, fields: 9 },
                pipeline: [
                    ["4:34 PM", "Submission record created", "Durable row reserved."],
                    ["4:37 PM", "Payment captured", "Gateway returned cleanly."],
                    ["4:37 PM", "Files saved", "FAILED — 1 of 2 uploads missing from S3. Needs retry."],
                    ["4:37 PM", "CRM + PDF", "Ran against partial file set."],
                    ["4:37 PM", "Billing history + emails + custom code", "All done."],
                ],
                recovery: { tone: "bad", text: "File save failed on a completed submission. Retry the failed uploads to make this record durable.", action: "Retry Failed Uploads" },
            },
            {
                id: 13160,
                name: "Maria Alvarez",
                membershipId: "4019233",
                state: "completed",
                status: "Completed",
                statusTone: "ok",
                stage: "Closed",
                stageTone: "ok",
                paymentDate: "6/12/2026 9:02:11 AM",
                expiration: "6/30/2027",
                created: "6/12/2026 8:58:40 AM",
                lastModified: "6/12/2026 9:03:30 AM",
                lastSubmitted: "6/12/2026 9:00:05 AM",
                ops: {
                    payment: { required: true, method: "card", amount: "$100.00", state: "done", note: "Card charged and confirmed." },
                    files: { fields: [{ label: "Continuing education certificate", required: false, max: 5, uploaded: 2, persisted: 2, items: [{ name: "continuing_ed_certificate.pdf", size: "412 KB", state: "done" }, { name: "license_card_front.jpg", size: "180 KB", state: "done" }] }] },
                    pdf: { configured: true, state: "done", note: "Built and stored." },
                    emails: [
                        { name: "Payment receipt", to: "maria.alvarez@example.com", trigger: "After payment", configured: true, state: "done" },
                        { name: "Signed PDF copy", to: "records@wvpsboard.gov", trigger: "After PDF", configured: true, state: "done" },
                    ],
                    records: { configured: true, state: "done", note: "Licensure database updated." },
                    billing: { configured: true, state: "done", note: "Billing history entry created." },
                    customCode: { configured: true, state: "done", note: "Post-submit script ran." },
                },
                booleans: {
                    filesSaved: true,
                    updateRecords: true,
                    awaitingUpdates: false,
                    approval: true,
                    auto: true,
                    pdfJob: true,
                    billingHistory: true,
                    emails: true,
                    customCode: true,
                    parents: true,
                    children: true,
                    grandchildren: true,
                    sibling: true,
                    other: true,
                    grandparent: true,
                    readyToProcess: true,
                },
                steps: ["done", "done", "done", "done", "done", "done"],
                stepsLabel: "6 / 6",
                updateStats: { updated: 4, inserted: 2, fields: 18 },
                pipeline: [
                    ["8:58 AM", "Submission record created", "Durable row reserved before payment."],
                    ["9:02 AM", "Payment captured", "Gateway callback confirmed payment_complete = true."],
                    ["9:02 AM", "Files saved", "Both uploads verified in S3."],
                    ["9:02 AM", "CRM updated", "Idempotent upsert applied."],
                    ["9:03 AM", "PDF + billing history", "Generated and stored."],
                    ["9:03 AM", "Emails + custom code", "All notifications sent — pipeline complete."],
                ],
                recovery: { tone: "ok", text: "All pipeline steps completed successfully and durably. No action required.", action: "View receipt" },
            },
            {
                id: 13098,
                name: "Nathaniel Boone",
                membershipId: "4015220",
                state: "abandoned",
                status: "Failed",
                statusTone: "bad",
                stage: "Abandoned after retries",
                stageTone: "bad",
                paymentDate: "5/2/2026 1:14:02 PM",
                expiration: "6/30/2026",
                created: "5/2/2026 1:10:45 PM",
                lastModified: "5/4/2026 2:00:00 AM",
                lastSubmitted: "5/2/2026 1:12:30 PM",
                ops: {
                    payment: { required: true, method: "card", amount: "$100.00", state: "done", note: "Card charged and confirmed — money was collected." },
                    files: { fields: [{ label: "Continuing education certificate", required: false, max: 5, uploaded: 1, persisted: 1, items: [{ name: "ce_summary.pdf", size: "210 KB", state: "done" }] }] },
                    pdf: { configured: true, state: "queued", note: "Never started — the pipeline halted at the CRM step." },
                    emails: [
                        { name: "Payment receipt", to: "n.boone@example.com", trigger: "After payment", configured: true, state: "done" },
                        { name: "Signed PDF copy", to: "records@wvpsboard.gov", trigger: "After PDF", configured: true, state: "queued" },
                    ],
                    records: { configured: true, state: "failed", note: "3 attempts failed (1:13, 1:43, 2:13 PM) with HTTP 500. Retry budget exhausted and the job was moved to the dead-letter queue." },
                    billing: { configured: true, state: "queued", note: "Never reached." },
                    customCode: { configured: true, state: "queued", note: "Never reached — the pipeline halted at the CRM step." },
                },
                booleans: {
                    filesSaved: true,
                    updateRecords: false,
                    awaitingUpdates: true,
                    approval: false,
                    auto: false,
                    pdfJob: false,
                    billingHistory: false,
                    emails: false,
                    customCode: false,
                    parents: true,
                    children: false,
                    grandchildren: null,
                    sibling: null,
                    other: null,
                    grandparent: true,
                    readyToProcess: true,
                },
                steps: ["done", "done", "done", "blocked", "pending", "pending"],
                stepsLabel: "3 / 6 • abandoned",
                updateStats: { pending: true },
                issues: [
                    { severity: "bad", text: "CRM update step failed on 3 consecutive retries (1:13 PM, 1:43 PM, 2:13 PM) — the downstream provider kept returning HTTP 500. After the retry budget was exhausted the submission was moved to the dead-letter queue and marked abandoned. Payment was already captured, so this record needs a manual completion or a refund decision." },
                ],
                pipeline: [
                    ["1:10 PM", "Submission record created", "Durable row reserved before payment."],
                    ["1:13 PM", "Payment captured", "payment_complete = true."],
                    ["1:13 PM", "CRM update — attempt 1", "FAILED — provider returned HTTP 500."],
                    ["1:43 PM", "CRM update — attempt 2 (retry)", "FAILED — provider returned HTTP 500."],
                    ["2:13 PM", "CRM update — attempt 3 (retry)", "FAILED — provider returned HTTP 500. Retry budget exhausted."],
                    ["2:13 PM", "Moved to dead-letter queue", "Marked abandoned — needs manual intervention."],
                ],
                recovery: { tone: "bad", text: "Abandoned after 3 failed CRM retries. Payment was captured — replay the pipeline once the provider is healthy, or escalate for a manual fix / refund.", action: "Replay pipeline" },
            },
        ];

        // CONTACT — "Contact Change" form type. Different form-level config than RENEW:
        // no payment, no billing, no custom code. "Awaiting Updates" is NOT a form-level
        // setting here — it is triggered only when the applicant submits a name change
        // (which requires a proof upload). Email/phone-only changes skip uploads + review.
        const contactRecords = [
            {
                id: 13201,
                formType: "CONTACT",
                name: "Janet Pearce",
                membershipId: "4012880",
                state: "completed",
                status: "Completed",
                statusTone: "ok",
                stage: "Closed",
                stageTone: "ok",
                paymentDate: "",
                expiration: "",
                created: "6/20/2026 9:14:02 AM",
                lastModified: "6/20/2026 9:15:10 AM",
                lastSubmitted: "6/20/2026 9:14:40 AM",
                ops: {
                    payment: { required: false, configured: false, state: "not_configured" },
                    files: { fields: [{ label: "Proof of name change", required: false, max: 2, uploaded: 0, persisted: 0, items: [] }] },
                    pdf: { configured: true, state: "done", note: "Confirmation PDF built and stored." },
                    emails: [],
                    records: { configured: true, state: "done", note: "Email + phone updated on the contact record immediately — no review needed for this change type." },
                    billing: { configured: false, state: "not_configured" },
                    customCode: { configured: false, state: "not_configured" },
                },
                booleans: {
                    readyToProcess: true,
                    awaitingUpdates: false,
                    approval: true,
                    auto: true,
                    parents: null,
                    children: null,
                    grandchildren: null,
                    sibling: null,
                    other: null,
                    grandparent: null,
                },
                steps: ["done", "done", "done", "done"],
                stepsLabel: "Done",
                updateStats: { updated: 1, inserted: 0, fields: 2 },
                pipeline: [
                    ["9:14 AM", "Submission record created", "Durable row reserved before certify."],
                    ["9:14 AM", "User certified / submitted", "Email + phone change only — no proof upload, no review needed."],
                    ["9:14 AM", "Record updated", "Contact email + phone applied immediately."],
                    ["9:15 AM", "Confirmation PDF built", "PDF generated and stored — pipeline complete. This form type does not send emails."],
                ],
                recovery: { tone: "ok", text: "All steps completed. A simple email/phone change needs no payment, no uploads, and no staff review.", action: "View receipt" },
            },
            {
                id: 13202,
                formType: "CONTACT",
                name: "Robert Castillo",
                membershipId: "4013447",
                state: "processing",
                status: "In Progress",
                statusTone: "neutral",
                stage: "Awaiting updates",
                stageTone: "warn",
                paymentDate: "",
                expiration: "",
                created: "6/21/2026 2:31:50 PM",
                lastModified: "6/21/2026 2:34:05 PM",
                lastSubmitted: "6/21/2026 2:33:12 PM",
                ops: {
                    payment: { required: false, configured: false, state: "not_configured" },
                    files: { fields: [{ label: "Proof of name change", required: true, max: 2, uploaded: 1, persisted: 1, items: [{ name: "marriage_certificate.pdf", size: "640 KB", state: "done" }] }] },
                    pdf: { configured: true, state: "ready", note: "Held until the name-change review is released." },
                    emails: [],
                    records: {
                        configured: true, state: "pending",
                        note: "Paused — the name change has not been applied to the record yet.",
                        awaiting: { source: "trigger", text: "name_change = Yes — staff must review the uploaded proof before the contact record is updated." },
                    },
                    billing: { configured: false, state: "not_configured" },
                    customCode: { configured: false, state: "not_configured" },
                },
                booleans: {
                    readyToProcess: true,
                    awaitingUpdates: true,
                    approval: false,
                    auto: false,
                    parents: null,
                    children: null,
                    grandchildren: null,
                    sibling: null,
                    other: null,
                    grandparent: null,
                },
                steps: ["done", "done", "pending", "pending"],
                stepsLabel: "2 / 4",
                updateStats: { pending: true },
                pipeline: [
                    ["2:31 PM", "Submission record created", "Durable row reserved before certify."],
                    ["2:33 PM", "Files saved", "Proof of name change (marriage_certificate.pdf) persisted to permanent storage."],
                    ["2:33 PM", "Awaiting Updates triggered", "name_change = Yes — record update paused for staff review."],
                    ["—", "Record update", "Held until a reviewer approves the proof."],
                    ["—", "Confirmation PDF", "Held behind the record update. This form type does not send emails."],
                ],
                recovery: { tone: "warn", text: "Awaiting Updates was triggered (name_change = Yes). A reviewer must approve the uploaded proof before the name change is applied.", action: "Open review queue" },
            },
            {
                id: 13203,
                formType: "CONTACT",
                name: "Linda Okafor",
                membershipId: "4014902",
                state: "completed",
                status: "Completed",
                statusTone: "ok",
                stage: "Closed",
                stageTone: "ok",
                paymentDate: "",
                expiration: "",
                created: "6/18/2026 11:40:22 AM",
                lastModified: "6/19/2026 9:05:48 AM",
                lastSubmitted: "6/18/2026 11:41:30 AM",
                ops: {
                    payment: { required: false, configured: false, state: "not_configured" },
                    files: { fields: [{ label: "Proof of name change", required: true, max: 2, uploaded: 1, persisted: 1, items: [{ name: "court_order_name_change.pdf", size: "512 KB", state: "done" }] }] },
                    pdf: { configured: true, state: "done", note: "Confirmation PDF built and stored after the review was released." },
                    emails: [],
                    records: {
                        configured: true, state: "done", reviewed: true,
                        note: "Name change applied after staff reviewed and released the Awaiting Updates hold.",
                    },
                    billing: { configured: false, state: "not_configured" },
                    customCode: { configured: false, state: "not_configured" },
                },
                booleans: {
                    readyToProcess: true,
                    awaitingUpdates: false,
                    approval: true,
                    auto: false,
                    parents: null,
                    children: null,
                    grandchildren: null,
                    sibling: null,
                    other: null,
                    grandparent: null,
                },
                steps: ["done", "done", "done", "done"],
                stepsLabel: "Done",
                updateStats: { updated: 1, inserted: 0, fields: 3 },
                pipeline: [
                    ["6/18 11:40 AM", "Submission record created", "Durable row reserved before certify."],
                    ["6/18 11:41 AM", "Files saved", "Proof of name change (court_order_name_change.pdf) persisted."],
                    ["6/18 11:41 AM", "Awaiting Updates triggered", "name_change = Yes — held for staff review."],
                    ["6/19 9:05 AM", "Reviewer released hold", "Staff approved the proof — record update resumed."],
                    ["6/19 9:05 AM", "Record updated + confirmation PDF", "Name change applied; confirmation PDF built. This form type does not send emails."],
                ],
                recovery: { tone: "ok", text: "Name change reviewed, approved, and applied. All steps completed. No action required.", action: "View receipt" },
            },
        ];

        // Each form type carries its own expected configuration; switching the Form Type
        // dropdown swaps the visible submission set.
        const recordSets = { RENEW: records, CONTACT: contactRecords };
        let activeType = "RENEW";
        function currentRecords() {
            return recordSets[activeType] || records;
        }

        const grid = document.getElementById("gridBody");
        const detailPanel = document.getElementById("detailPanel");
        let activeId = null;

        function icon(name, cls = "") {
            return `<svg class="ic ${cls}" viewBox="0 0 24 24" aria-hidden="true"><use href="#ic-${name}"></use></svg>`;
        }

        function boolCell(value) {
            if (value === true) return '<span class="bool t" title="True">✓</span>';
            if (value === false) return '<span class="bool f" title="False">·</span>';
            return '<span class="bool na" title="N/A">—</span>';
        }

        function chip(tone, label) {
            return `<span class="chip ${tone}">${label}</span>`;
        }

        function stepsStrip(steps, label) {
            const segs = steps.map((s) => `<span class="seg ${s}"></span>`).join("");
            return `<span class="steps">${segs}<span class="label">${label}</span></span>`;
        }

        // Unified operation-status vocabulary (shared by list + detail).
        // Canonical values — must match patched/db/form_submit_ledger.sql ck_fso_state.
        // UI may map a canonical state to a friendlier label (e.g. payment awaiting + card -> "Awaiting gateway").
        const OP_STATE = {
            not_configured: { label: "Not set up", tone: "neutral", icon: "minus" },
            not_needed: { label: "Not needed", tone: "neutral", icon: "minus" },
            ready: { label: "Set up", tone: "info", icon: "check-circle" },
            queued: { label: "Queued", tone: "warn", icon: "clock", anim: "pulse" },
            running: { label: "Running", tone: "warn", icon: "refresh", anim: "spin" },
            pending: { label: "In progress", tone: "warn", icon: "clock", anim: "pulse" },
            retrying: { label: "Retrying", tone: "warn", icon: "refresh", anim: "spin" },
            awaiting: { label: "Awaiting response", tone: "warn", icon: "clock", anim: "pulse" },
            done: { label: "Done", tone: "ok", icon: "check-circle" },
            failed: { label: "Failed", tone: "bad", icon: "alert" },
        };

        function stateTone(state) {
            return (OP_STATE[state] || OP_STATE.not_configured).tone;
        }

        function opChip(state, opts = {}) {
            const m = OP_STATE[state] || OP_STATE.not_configured;
            const anim = m.anim === "spin" ? " live" : (m.anim === "pulse" ? " pulse-dot" : "");
            const label = opts.label != null ? opts.label : m.label;
            const title = opts.title || label;
            return `<span class="chip ${m.tone}${anim}" title="${title}">${icon(m.icon)}${label}</span>`;
        }

        // Files: a field may hold several uploaded files; each must be persisted (temp -> S3).
        function filesSummary(files) {
            const fields = (files && files.fields) || [];
            let uploaded = 0, persisted = 0, failed = 0, retrying = 0, pending = 0;
            fields.forEach((f) => {
                uploaded += f.uploaded || 0;
                persisted += f.persisted || 0;
                (f.items || []).forEach((it) => {
                    if (it.state === "failed") failed++;
                    else if (it.state === "retrying") retrying++;
                    else if (it.state === "pending" || it.state === "running" || it.state === "queued") pending++;
                });
            });
            let state = "not_needed";
            if (uploaded > 0) {
                if (failed > 0) state = "failed";
                else if (retrying > 0) state = "retrying";
                else if (pending > 0 || persisted < uploaded) state = "pending";
                else state = "done";
            }
            return { uploaded, persisted, failed, retrying, pending, state };
        }

        function gridFilesCell(files) {
            const s = filesSummary(files);
            if (s.uploaded === 0) return '<span class="bool na" title="No files uploaded">—</span>';
            return opChip(s.state, { label: `${s.persisted}/${s.uploaded}`, title: `${s.persisted} of ${s.uploaded} uploaded files saved to permanent storage` });
        }

        function gridEmailsCell(emails) {
            // "not_needed" emails are configured for the form but skipped on this submission
            // (e.g. a payment receipt when payment was bypassed) — leave them out of the count.
            const cfg = (emails || []).filter((e) => e.configured && e.state !== "not_needed");
            if (!cfg.length) return '<span class="bool na" title="No emails to send for this submission">—</span>';
            const by = (s) => cfg.filter((e) => e.state === s).length;
            const done = by("done"), failed = by("failed"), retrying = by("retrying"), running = by("running");
            const waiting = cfg.length - done - failed - retrying - running; // queued / ready / pending / awaiting
            // Icon reflects the in-flight email(s): spinner = actively working (retrying/running),
            // clock = waiting its turn. The count only says how many are already sent.
            let state = "pending";
            if (failed) state = "failed";
            else if (retrying) state = "retrying";
            else if (running) state = "running";
            else if (done === cfg.length) state = "done";
            const parts = [`${done} of ${cfg.length} sent`];
            if (failed) parts.push(`${failed} failed`);
            if (retrying) parts.push(`${retrying} retrying after a failure (spinner)`);
            if (running) parts.push(`${running} sending now (spinner)`);
            if (waiting) parts.push(`${waiting} queued / waiting (clock)`);
            return opChip(state, { label: `${done}/${cfg.length}`, title: parts.join(" · ") });
        }

        function gridPdfCell(pdf) {
            if (!pdf || !pdf.configured) return '<span class="bool na" title="PDF not configured">—</span>';
            return opChip(pdf.state);
        }

        function paymentAwaitingLabel(p) {
            if (p && p.method === "card") return "Awaiting gateway";
            return (OP_STATE.awaiting || {}).label || "Awaiting";
        }

        function gridPaymentCell(p) {
            if (!p) return '<span class="bool na" title="No payment data">—</span>';
            if (!p.required || p.state === "not_configured") {
                return '<span class="chip neutral" title="This form type does not require payment">NA</span>';
            }
            if (p.bypassed || p.state === "not_needed") {
                return `<span class="chip neutral" title="Configured payment page bypassed — no gateway charge">${icon('minus')}NA</span>` +
                    (p.bypassField ? ` <span style="font-size:11px;color:var(--bp-muted)">(${p.bypassField} = ${p.bypassValue})</span>` : "");
            }
            const amt = p.amount || "";
            const amtParen = amt ? ` (${amt})` : "";
            // List view uses legacy-admin wording so PMs can scan completion at a glance.
            switch (p.state) {
                case "done":
                    return `<span class="chip ok" title="Payment completed${p.note ? " — " + p.note : ""}">${icon('check-circle')}Complete${amtParen}</span>`;
                case "awaiting":
                    return `<span class="chip warn pulse-dot" title="Payment not completed — ${p.note || paymentAwaitingLabel(p)}">${icon('clock')}Pending${amtParen}</span>`;
                case "pending":
                    return `<span class="chip warn pulse-dot" title="Payment not completed — ${p.note || "awaiting confirmation"}">${icon('clock')}Pending${amtParen}</span>`;
                case "failed":
                    return `<span class="chip bad" title="Payment failed${p.note ? " — " + p.note : ""}">${icon('alert')}Failed${amtParen}</span>`;
                case "ready":
                    return `<span class="chip neutral" title="Payment required but not collected yet">${icon('clock')}Not paid${amtParen}</span>`;
                default:
                    return `<span class="bool na" title="Payment status unknown">—</span>`;
            }
        }

        const STATE_ORDER = ["draft", "awaiting_payment", "payment_complete", "processing", "completed"];
        const STATE_TONE = {
            draft: "neutral",
            awaiting_payment: "warn",
            payment_complete: "ok",
            processing: "warn",
            completed: "ok",
            failed: "bad",
            abandoned: "bad",
        };
        // Plain-language meaning of each durable state (shown on hover for PMs).
        const STATE_DESC = {
            draft: "Started but not yet submitted — the applicant is still filling out the form. Nothing has run.",
            awaiting_payment: "Submitted and sent to the payment page — waiting for the payment result to come back.",
            payment_complete: "Payment confirmed. The post-submission steps are about to run.",
            processing: "Submitted (and paid, if required). The system is now running the back-end steps — saving files, updating the record, building the PDF, sending any emails. This is the working state between submit and done.",
            completed: "Every step finished and was saved durably. Nothing left to do.",
            failed: "A step failed and the submission could not finish on its own — needs attention.",
            abandoned: "Never completed — e.g. payment was never finished, or retries were exhausted.",
        };
        function stateChip(state, hasIssues) {
            const warn = hasIssues ? ' <span class="state-warn" title="has unresolved issues">⚠</span>' : '';
            const desc = STATE_DESC[state] || "durable submission state";
            return `<span class="chip ${STATE_TONE[state] || 'neutral'}" title="${desc}">${state}${warn}</span>`;
        }
        function opCardTone(state, configured) {
            if (configured === false) return "neutral";
            return stateTone(state);
        }

        // Review status — derived from the ACTUAL hold state (one source of truth),
        // not a separate flag. Distinct from legacy app_renew_status = APPROVED.
        function awaitingHold(record) {
            const r = record.ops && record.ops.records;
            return !!(r && r.awaiting);
        }
        function reviewStatus(record) {
            const r = record.ops && record.ops.records;
            if (r && r.awaiting) return { label: "Awaiting", tone: "warn", title: "Held for staff review right now — " + r.awaiting.text };
            if (r && r.reviewed) return { label: "Approved", tone: "ok", title: "A staff member reviewed and released this submission." };
            if (record.state === "draft") return { label: "—", tone: "neutral", title: "Not submitted yet — no review status." };
            return { label: "Auto", tone: "neutral", title: "Approved automatically — no staff review was required for this submission." };
        }
        function reviewCell(record) {
            const a = reviewStatus(record);
            return `<span class="chip ${a.tone}" title="${a.title}">${a.label}</span>`;
        }

        function hasRelationshipData(b) {
            if (!b) return false;
            return ["parents", "children", "grandchildren", "sibling", "other", "grandparent"].some(
                (k) => b[k] !== null && b[k] !== undefined
            );
        }

        // Why is an operation active? "form config" = set for every submission of this
        // form; "triggered" = turned on by a field value on this particular submission.
        function whyInline(why) {
            if (!why) return "";
            const tag = why.source === "trigger"
                ? '<span class="why-tag trigger">triggered</span>'
                : '<span class="why-tag config">form config</span>';
            return `${tag}<span>${why.text}</span>`;
        }
        function whyLine(why) {
            return why ? `<div class="op-why">${whyInline(why)}</div>` : "";
        }

        function opCard(o) {
            const cfgLine = (o.configured === undefined)
                ? ""
                : `<div class="op-config"><span class="dot ${o.configured ? "on" : "off"}"></span>${o.configured ? "Configured for this form" : "Not configured for this form"}</div>`;
            return `<div class="op-card ${o.tone || "neutral"}${o.span ? " span2" : ""}">
                <div class="op-head"><span class="op-name">${icon(o.icon)} ${o.name}</span>${o.statusHtml}</div>
                ${cfgLine}
                ${o.desc ? `<div class="op-desc">${o.desc}</div>` : ""}
                ${whyLine(o.why)}
                ${o.body || ""}
                ${o.action ? `<div class="op-action">${o.action}</div>` : ""}
            </div>`;
        }

        function paymentOpCard(p) {
            if (!p) return "";
            if (!p.required || p.state === "not_configured") {
                return opCard({
                    icon: "dollar-sign", name: "Payment", tone: "neutral", configured: false,
                    statusHtml: `<span class="chip neutral">${icon('minus')}Not required</span>`,
                    desc: "This form flow does not require a payment.",
                });
            }
            if (p.bypassed || p.state === "not_needed") {
                return opCard({
                    icon: "dollar-sign", name: "Payment", tone: "neutral",
                    statusHtml: `<span class="chip neutral">${icon('minus')}NA (${p.bypassField} = ${p.bypassValue})</span>`,
                    desc: `The configured payment page (<code>${p.bypassPage || "pay.asp"}</code>) was bypassed by an alternate page (set in the <strong>Form Pages</strong> tool) that handles Create PDF / Update Record / Set Submission Status directly. No payment gateway was involved.`,
                    why: { source: "trigger", text: `${p.bypassField} = ${p.bypassValue} routed this submission to the alternate page.` },
                });
            }
            const methodLabel = p.method === "card" ? "Pay by card" : "Payment";
            let statusHtml, action = "";
            switch (p.state) {
                case "done": statusHtml = opChip("done", { label: "Paid" }); break;
                case "awaiting":
                    statusHtml = opChip("awaiting", { label: paymentAwaitingLabel(p) });
                    action = `<button class="btn">${icon("credit-card")} Reconcile with gateway</button>`;
                    break;
                case "pending": statusHtml = opChip("pending"); break;
                case "failed":
                    statusHtml = opChip("failed");
                    action = `<button class="btn">${icon("credit-card")} Reconcile with gateway</button>`;
                    break;
                case "ready": statusHtml = opChip("ready"); break;
                default: statusHtml = `<span class="chip neutral">—</span>`;
            }
            const tone = p.state === "done" ? "ok" : p.state === "failed" ? "bad" : p.state === "ready" ? "info" : "warn";
            const desc = `${methodLabel}${p.amount ? ` — <strong>${p.amount}</strong>` : ""}.${p.note ? " " + p.note : ""}`;
            return opCard({ icon: "dollar-sign", name: "Payment", tone, statusHtml, desc, action });
        }

        function filesOpCard(files) {
            const s = filesSummary(files);
            const tone = s.uploaded === 0 ? "neutral" : opCardTone(s.state);
            const statusHtml = s.uploaded === 0
                ? `<span class="chip neutral">${icon('minus')}No files uploaded</span>`
                : opChip(s.state, { label: `${s.persisted} of ${s.uploaded} saved` });
            const fieldsHtml = ((files && files.fields) || [])
                .filter((f) => (f.uploaded || 0) > 0 || f.required)
                .map((f) => {
                    const items = (f.items || []).map((it) => {
                        const lab = it.state === "done" ? "Saved"
                            : it.state === "failed" ? "Stuck in temp"
                            : it.state === "retrying" ? "Retrying"
                            : it.state === "pending" ? "Saving…"
                            : (OP_STATE[it.state] || OP_STATE.pending).label;
                        return `<li class="${it.state === "failed" ? "bad" : ""}">
                            <span class="uf-fname">${it.name}${it.error ? `<div class="err">${it.error}</div>` : ""}</span>
                            <span class="uf-size">${it.size || ""}</span>
                            ${opChip(it.state, { label: lab })}
                        </li>`;
                    }).join("");
                    const need = (f.uploaded || 0) - (f.persisted || 0);
                    return `<div class="upload-field">
                        <div class="uf-head">
                            <span class="uf-name">${f.label}</span>
                            <span class="uf-meta">${f.uploaded || 0} uploaded · ${f.persisted || 0} saved${need > 0 ? ` · ${need} pending` : ""} · ${f.required ? "required" : "optional"}${f.max ? ` · up to ${f.max}` : ""}</span>
                        </div>
                        ${items ? `<ul class="upload-files">${items}</ul>` : ""}
                    </div>`;
                }).join("");
            const need = s.uploaded - s.persisted;
            const desc = s.uploaded === 0
                ? "No files were uploaded with this submission."
                : `Every uploaded file must be moved from temporary storage to permanent storage. ${need > 0 ? `<strong>${need} file(s) still need to be saved.</strong>` : "All uploaded files are saved."}`;
            const action = s.state === "failed" ? `<button class="btn danger">${icon("upload")} Retry failed saves</button>` : "";
            return opCard({ icon: "paperclip", name: "Uploaded files", tone, statusHtml, desc, body: fieldsHtml, action, span: true });
        }

        function pdfOpCard(pdf) {
            if (!pdf) return "";
            const tone = opCardTone(pdf.state, pdf.configured);
            const statusHtml = pdf.configured ? opChip(pdf.state) : `<span class="chip neutral">${icon('minus')}Not set up</span>`;
            const desc = pdf.configured
                ? (pdf.note || "")
                : "This form is not configured to generate a PDF.";
            const why = pdf.configured
                ? (pdf.why || { source: "config", text: "Set at the form level — every submission of this form builds a PDF." })
                : null;
            return opCard({ icon: "printer", name: "PDF document", tone, statusHtml, desc, why, configured: pdf.configured });
        }

        // Derive a plain-language "why" for an email from its trigger, unless one is
        // given explicitly (used for field-condition email pages).
        function emailWhy(e) {
            const t = (e.trigger || "").toLowerCase();
            let text;
            if (t.includes("after pdf")) text = "Emails the generated PDF after it is built.";
            else if (t.includes("after payment")) text = "Sent automatically after a successful payment.";
            else if (t.includes("on submit")) text = "Sent on every submission of this form.";
            else text = "Configured at the form level.";
            return { source: "config", text };
        }

        function emailsOpCard(emails) {
            const cfg = (emails || []).filter((e) => e.configured);
            // "active" = configured emails that actually need to send on this submission.
            const active = cfg.filter((e) => e.state !== "not_needed");
            const done = active.filter((e) => e.state === "done").length;
            let overall = "pending";
            if (!active.length) overall = cfg.length ? "not_needed" : "not_configured";
            else if (active.some((e) => e.state === "failed")) overall = "failed";
            else if (active.some((e) => e.state === "retrying")) overall = "retrying";
            else if (done === active.length) overall = "done";
            const tone = active.length ? opCardTone(overall) : "neutral";
            const statusHtml = active.length
                ? opChip(overall, { label: `${done} of ${active.length} sent` })
                : `<span class="chip neutral">${icon('minus')}${cfg.length ? "None to send" : "NA"}</span>`;
            const list = cfg.map((e) => {
                const lab = e.state === "done" ? "Sent" : e.state === "not_needed" ? "Not sent" : undefined;
                const extra = e.attempts ? ` (${e.attempts})` : "";
                const why = e.why || emailWhy(e);
                return `<li>
                    <span class="em-main">${e.name}</span>
                    ${opChip(e.state, { label: lab })}
                    <span class="em-sub">${e.trigger} → ${e.to}${extra}${e.note ? ` — ${e.note}` : ""}</span>
                    <span class="em-why">${whyInline(why)}</span>
                </li>`;
            }).join("");
            const desc = cfg.length
                ? "Emails are sent through the job queue. Each one is tracked from queued → sent (or failed / retrying)."
                : "Not applicable — this form type does not send any emails.";
            return opCard({ icon: "send", name: "Emails", tone, statusHtml, desc, body: cfg.length ? `<ul class="email-list">${list}</ul>` : "", span: true });
        }

        function recordsOpCard(op, updateStats) {
            if (!op) return "";
            const tone = op.awaiting ? "warn" : opCardTone(op.state, op.configured);
            const statusHtml = op.configured === false
                ? `<span class="chip neutral">${icon('minus')}Not set up</span>`
                : op.awaiting ? opChip("pending", { label: "Awaiting review" }) : opChip(op.state);
            const holdHtml = op.awaiting
                ? `<div class="op-hold">${icon("pause-circle")} <span><strong>Awaiting Updates</strong> — held for staff review before the record is updated.</span></div>${whyLine(op.awaiting)}`
                : "";
            return opCard({ icon: "database", name: "Record updates (CRM)", tone, statusHtml, desc: op.note || "", body: holdHtml + updateStatsLine({ updateStats }) });
        }

        function genericOpCard(name, ic, op) {
            if (!op) return "";
            const configured = op.configured !== false;
            const tone = opCardTone(op.state, configured);
            const statusHtml = configured ? opChip(op.state) : `<span class="chip neutral">${icon('minus')}Not set up</span>`;
            return opCard({ icon: ic, name, tone, statusHtml, desc: op.note || "", configured });
        }
        function recordLifecycle(state) {
            const isTerminal = state === "failed" || state === "abandoned";
            const currentIdx = STATE_ORDER.indexOf(state);
            const steps = STATE_ORDER.map((s, i) => {
                let cls = "";
                if (!isTerminal && currentIdx >= 0) {
                    if (i < currentIdx) cls = "done";
                    else if (i === currentIdx) cls = "current";
                }
                return `<span class="lc-step ${cls}">${s}</span>`;
            }).join('<span class="lc-arrow">→</span>');
            const tail = isTerminal
                ? `<span class="lc-arrow">/</span><span class="lc-step failed">${state}</span>`
                : "";
            return `<div class="lc-row">${steps}${tail}</div>`;
        }

        function renderGrid() {
            const list = currentRecords();
            const pageInfo = document.getElementById("pageInfo");
            if (pageInfo) pageInfo.textContent = `1-${list.length} of ${list.length} values`;
            grid.innerHTML = list.map((record, i) => `
                <tr data-id="${record.id}" class="${record.id === activeId ? "selected" : ""}">
                    <td class="check" data-col="check"><input type="checkbox" title="select"></td>
                    <td class="num" data-col="num">${i + 1}.</td>
                    <td data-col="edit"><a href="#" onclick="return false;">Edit ${record.id}</a></td>
                    <td data-col="state">${stateChip(record.state, record.issues && record.issues.length > 0)}</td>
                    <td data-col="status">${chip(record.statusTone, record.status)}</td>
                    <td data-col="stage">${chip(record.stageTone, record.stage)}</td>
                    <td class="nowrap" data-col="name">${record.name}</td>
                    <td class="nowrap" data-col="membershipId">${record.membershipId}</td>
                    <td data-col="pdf">${gridPdfCell(record.ops.pdf)}</td>
                    <td data-col="files">${gridFilesCell(record.ops.files)}</td>
                    <td data-col="emails">${gridEmailsCell(record.ops.emails)}</td>
                    <td data-col="awaitingUpdates" title="Is this submission parked in the staff-review hold right now? Yes = it actually entered and persisted the awaiting-updates state (not just 'configured to maybe need review'). Clears to No once a reviewer releases it.">${awaitingHold(record) ? '<span class="chip warn">Yes</span>' : '<span class="chip neutral">No</span>'}</td>
                    <td data-col="review" title="Review status: Auto = no staff review needed; Awaiting = held for staff review now; Approved = a staff member reviewed and released it.">${reviewCell(record)}</td>
                    <td data-col="payment">${gridPaymentCell(record.ops.payment)}</td>
                    <td data-col="steps">${stepsStrip(record.steps, record.stepsLabel)}</td>
                    <td class="nowrap" data-col="paymentDate">${record.paymentDate}</td>
                    <td class="nowrap" data-col="expiration">${record.expiration}</td>
                    <td class="nowrap" data-col="created">${record.created}</td>
                    <td class="nowrap" data-col="lastModified">${record.lastModified}</td>
                    <td class="nowrap" data-col="lastSubmitted">${record.lastSubmitted}</td>
                </tr>
            `).join("");

            grid.querySelectorAll("tr").forEach((row) => {
                row.addEventListener("click", (e) => {
                    if (e.target instanceof HTMLInputElement || e.target instanceof HTMLAnchorElement) return;
                    activeId = Number(row.dataset.id);
                    renderGrid();
                    renderDetail();
                    detailPanel.scrollIntoView({ behavior: "smooth", block: "start" });
                });
            });
        }

        function boolItem(label, value, opts = {}) {
            let cls = value === true ? "t" : value === false ? "f" : "na";
            let txt = value === true ? "TRUE" : value === false ? "false" : "n/a";
            if (opts.errorWhenFalse && value === false) {
                cls = "bad t";
                txt = "FAILED — RETRY";
            } else if (opts.warnWhenFalse && value === false) {
                cls = "warn t";
                txt = "PENDING";
            }
            return `<div class="item ${cls}"><span class="lbl">${label}</span><span class="val">${txt}</span></div>`;
        }

        function updateStatsLine(record) {
            const s = record.updateStats;
            if (!s) return "";
            if (s.pending) {
                return `<div class="update-stats"><span class="us pending">${icon('clock')} Record updates pending — run when the pipeline reaches this step</span></div>`;
            }
            return `<div class="update-stats">
                <span class="us">${icon('database')} <strong>${s.updated}</strong> records updated</span>
                <span class="us"><strong>${s.inserted}</strong> inserted</span>
                <span class="us"><strong>${s.fields}</strong> fields changed</span>
            </div>`;
        }

        function renderDetail() {
            if (activeId == null) {
                detailPanel.hidden = true;
                detailPanel.innerHTML = "";
                return;
            }
            detailPanel.hidden = false;
            const list = currentRecords();
            const record = list.find((r) => r.id === activeId) || list[0];
            const b = record.booleans;
            const formType = record.formType || "RENEW";

            const debugLinks = ["index.asp", "alert.asp", "contact.asp", "legal.asp", "explain.asp", "pdh.asp", "certify.asp", "receipt.asp", "receipt_nopay.asp"]
                .map((l) => `<li><a href="#" onclick="return false;">${l}</a></li>`).join("");

            const recoveryToneIcon = { warn: "alert-triangle", bad: "alert", ok: "check-circle" };
            const pickActionIcon = (txt) => {
                const t = (txt || "").toLowerCase();
                if (/reconcile|gateway/.test(t)) return "credit-card";
                if (/abandon/.test(t)) return "trash";
                if (/retry|upload/.test(t)) return "upload";
                if (/approval/.test(t)) return "check-circle";
                if (/resume|pipeline|replay/.test(t)) return "refresh";
                return "chevron-right";
            };
            const recoveryBanner = record.recovery
                ? `<div class="recovery-banner ${record.recovery.tone}">
                       <span>${icon(recoveryToneIcon[record.recovery.tone] || "alert-triangle")} <strong>Recovery:</strong> ${record.recovery.text}</span>
                       <span class="actions"><button>${icon(pickActionIcon(record.recovery.action))} ${record.recovery.action}</button></span>
                   </div>`
                : "";

            const issuesBanner = (record.issues && record.issues.length)
                ? `<div class="issues-banner ${record.issues.some(i => i.severity === 'bad') ? '' : 'warn'}">
                       <div class="ib-title">${icon('alert-triangle')} Marked ${record.state} — durable record reports unresolved issues</div>
                       <ul>${record.issues.map(i => `<li>${i.text}</li>`).join('')}</ul>
                   </div>`
                : "";

            const ops = record.ops;
            const operationsSection = `
                <div class="state-section">
                    <div class="state-title">${icon("layers")} Operations <span class="new-marker">new</span></div>
                    <div class="op-grid">
                        ${paymentOpCard(ops.payment)}
                        ${pdfOpCard(ops.pdf)}
                        ${recordsOpCard(ops.records, record.updateStats)}
                        ${genericOpCard("Billing history", "file-text", ops.billing)}
                        ${genericOpCard("Custom code", "layers", ops.customCode)}
                        ${filesOpCard(ops.files)}
                        ${emailsOpCard(ops.emails)}
                    </div>
                </div>`;

            const lifecycleStrip = `
                <div class="lifecycle lc-embedded">
                    <div class="lc-title">This submission's state</div>
                    ${recordLifecycle(record.state)}
                </div>`;

            detailPanel.innerHTML = `
                <div class="detail-header">
                    <div class="title">Application — ${record.name}</div>
                    <div class="id">app_${formType.toLowerCase()}_id=${record.id}</div>
                    <button class="btn-list" type="button">${icon('list')} List</button>
                </div>

                ${lifecycleStrip}

                ${issuesBanner}

                ${recoveryBanner}

                <div class="detail-body">
                    <aside class="detail-meta">
                        <div class="meta-row"><span class="k">Created:</span><span class="v">${record.created}</span></div>
                        <div class="meta-row"><span class="k">Type:</span><span class="v">${formType}</span></div>
                        <div class="meta-row"><span class="k">Status:</span><span class="v">${chip(record.statusTone, record.status)}</span></div>
                        <div class="meta-row"><span class="k">State:</span><span class="v">${stateChip(record.state, record.issues && record.issues.length > 0)} <span class="new-marker">durable</span></span></div>
                        <div class="meta-row"><span class="k">Stage:</span><span class="v">${chip(record.stageTone, record.stage)}</span></div>
                        <div class="meta-row"><span class="k">Internal ID:</span><span class="v">${record.membershipId}</span></div>
                        <div class="meta-row"><span class="k">Submission ID:</span><span class="v"><code>SUB-${record.id}</code> <span class="new-marker">durable</span></span></div>
                        <div class="meta-row" title="Review status: Auto = no staff review needed; Awaiting = held for staff review now; Approved = a staff member reviewed and released it."><span class="k">Review:</span><span class="v">${reviewCell(record)}</span></div>
                        <div class="meta-row"><span class="k">Payment:</span><span class="v">${gridPaymentCell(record.ops.payment)}</span></div>
                        <div class="meta-row"><span class="k">Files:</span><span class="v">${gridFilesCell(record.ops.files)}</span></div>
                        <div class="meta-row"><span class="k">PDF:</span><span class="v">${gridPdfCell(record.ops.pdf)}</span></div>
                        <div class="meta-row"><span class="k">Emails:</span><span class="v">${gridEmailsCell(record.ops.emails)}</span></div>
                        <div class="meta-row"><span class="k">Database Record:</span><span class="v"><a href="#" onclick="return false;">Link</a></span></div>
                        <div class="meta-row">
                            <span class="k">Debug Links:</span>
                            <span class="v"><ul class="debug-list">${debugLinks}</ul></span>
                        </div>
                        <div class="meta-row"><span class="k">Restore Links:</span><span class="v"><a href="#" onclick="return false;">Show</a></span></div>
                    </aside>

                    <div class="detail-state">
                        ${operationsSection}

                        <div class="state-section">
                            <div class="state-title">Status flags</div>
                            <div class="bool-grid">
                                ${boolItem("Submission record created", true)}
                                ${boolItem("ready_to_process", b.readyToProcess, { warnWhenFalse: true })}
                                ${boolItem("Awaiting updates", awaitingHold(record))}
                                ${boolItem("Approved", b.approval, { warnWhenFalse: true })}
                                ${boolItem("Auto", b.auto)}
                            </div>
                        </div>

                        ${hasRelationshipData(b) ? `
                        <div class="state-section">
                            <div class="state-title">Relationship completion <span class="new-marker">new</span></div>
                            <div class="bool-grid">
                                ${boolItem("Parents", b.parents)}
                                ${boolItem("Children", b.children)}
                                ${boolItem("Grandchildren", b.grandchildren)}
                                ${boolItem("Sibling", b.sibling)}
                                ${boolItem("Other", b.other)}
                                ${boolItem("Grandparent", b.grandparent)}
                            </div>
                        </div>
                        ` : ""}

                        <div class="state-section">
                            <div class="state-title">Pipeline checkpoints</div>
                            <div class="pipeline">
                                ${record.pipeline.map(([when, what, note]) => `
                                    <div class="step">
                                        <span class="when">${when}</span>
                                        <span class="what">${what}</span>
                                        <span class="note">${note}</span>
                                    </div>
                                `).join("")}
                            </div>
                        </div>
                    </div>
                </div>

                <div class="actions-row">
                    <button class="btn">${icon('pause-circle')} Pend</button>
                    <button class="btn">${icon('x-circle')} Deny</button>
                    <button class="btn">${icon('file-text')} Rebuild PDF</button>
                    <button class="btn">${icon('maximize')} Show Data - Full</button>
                    <button class="btn primary">${icon('refresh')} Replay Pipeline <span class="new-marker">new</span></button>
                    <button class="btn">${icon('credit-card')} Force Reconcile <span class="new-marker">new</span></button>
                    ${filesSummary(record.ops.files).state === "failed"
                ? `<button class="btn danger">${icon('upload')} Retry Failed Uploads <span class="new-marker">new</span></button>`
                : ''}
                </div>
            `;

            const listBtn = detailPanel.querySelector(".btn-list");
            if (listBtn) {
                listBtn.addEventListener("click", () => {
                    activeId = null;
                    renderGrid();
                    renderDetail();
                    document.querySelector(".grid-wrap").scrollIntoView({ behavior: "smooth", block: "start" });
                });
            }
        }

        // Column selector — toggle optional columns; injects a <style> rule per hidden column.
        const COLUMNS = [
            { key: "state", label: "State", on: true },
            { key: "status", label: "Status", on: true },
            { key: "stage", label: "Stage", on: false },
            { key: "membershipId", label: "MEMBERSHIP_ID", on: false },
            { key: "pdf", label: "PDF", on: true },
            { key: "files", label: "Files", on: true },
            { key: "emails", label: "Emails", on: true },
            { key: "awaitingUpdates", label: "Awaiting Updates", on: false },
            { key: "review", label: "Review", on: true },
            { key: "payment", label: "Payment", on: true },
            { key: "steps", label: "Steps", on: true },
            { key: "paymentDate", label: "Payment Date-Time", on: false },
            { key: "expiration", label: "Expiration", on: false },
            { key: "created", label: "Created", on: false },
            { key: "lastModified", label: "Last Modified", on: false },
            { key: "lastSubmitted", label: "Last Submitted", on: false },
        ];
        const colVisStyle = document.createElement("style");
        document.head.appendChild(colVisStyle);
        const colBtn = document.getElementById("colBtn");
        const colMenu = document.getElementById("colMenu");

        function applyColumnVisibility() {
            const hidden = COLUMNS.filter((c) => !c.on).map((c) => `[data-col="${c.key}"]`);
            colVisStyle.textContent = hidden.length ? hidden.join(",") + "{display:none;}" : "";
        }

        colMenu.innerHTML = COLUMNS.map((c) =>
            `<label><input type="checkbox" data-colkey="${c.key}" ${c.on ? "checked" : ""}> ${c.label}</label>`
        ).join("");
        colMenu.querySelectorAll("input").forEach((input) => {
            input.addEventListener("change", () => {
                const col = COLUMNS.find((c) => c.key === input.dataset.colkey);
                col.on = input.checked;
                applyColumnVisibility();
            });
        });
        colBtn.addEventListener("click", (e) => {
            e.stopPropagation();
            const open = colMenu.hidden;
            colMenu.hidden = !open;
            colBtn.setAttribute("aria-expanded", String(open));
        });
        document.addEventListener("click", (e) => {
            if (!colMenu.hidden && !colMenu.contains(e.target) && e.target !== colBtn) {
                colMenu.hidden = true;
                colBtn.setAttribute("aria-expanded", "false");
            }
        });
        applyColumnVisibility();

        // Form Type drives the expected configuration — switching it swaps submission sets.
        const formTypeSel = document.getElementById("formTypeFilter");
        if (formTypeSel) {
            formTypeSel.addEventListener("change", () => {
                activeType = recordSets[formTypeSel.value] ? formTypeSel.value : "RENEW";
                activeId = null;
                renderGrid();
                renderDetail();
            });
        }

        renderGrid();
        renderDetail();
    