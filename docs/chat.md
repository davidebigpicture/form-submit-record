Yes — and that’s an important correction, and honestly it makes your design **stronger** than what I described earlier.

You’re doing this:

```
Create Submission ✅
→ Then go to Payment
```

instead of:

```
Payment → then create Submission  ❌
```

👉 That’s the **right call**, especially in your environment.

***

# ✅ What this guarantees (very concretely)

With your ordering:

```
CreateSubmission()
→ Payment()
→ rest of pipeline...
```

### Now every failure scenario becomes recoverable

***

## ✅ Case: crash DURING payment

```
Submission ✅ (exists)
Payment ❌ (unknown/failed)
```

What you have:

* A durable intent record
* Enough data to:
  * retry payment
  * abandon safely

👉 No data loss

***

## ✅ Case: crash AFTER payment (your big concern)

```
Submission ✅
Payment ✅
💥 crash before SaveRecord / PDF / etc.
```

What you have:

* Submission record ✅
* Payment likely recorded (or reconcilable)

👉 Recovery job can:

```
→ detect payment_complete = true
→ resume remaining steps
```

✅ This is exactly what you want

***

## ✅ Case: crash mid SaveRecord (your earlier example)

```
Submission ✅
Payment ✅
Files partial ✅❌
CRM ❌
```

👉 Recovery has:

* submission\_id
* expected files
* payment info

So it can safely:

```
→ finish files
→ run UpdateRecord (transactional)
→ continue pipeline
```

***

# 🧠 Subtle but important upgrade your approach gives you

By creating the submission **before payment**, you’ve separated:

### ✅ “Intent”

> the user is trying to submit something

from:

### ✅ “Execution”

> the system is finishing all the steps

***

That’s actually a **clean architectural boundary**, even if you didn’t frame it that way:

```
User Intent (Submission exists)
            ↓
Execution Workflow (idempotent steps)
```

***

# ⚠️ One thing this introduces (you should be aware of)

Now you *will* get:

### “Abandoned submissions”

Cases like:

* user never completes payment
* browser closed
* payment failed

So you need a policy:

***

## ✅ Add submission states

At minimum:

```
draft
awaiting_payment
payment_complete
processing
completed
failed
```

***

## ✅ Add cleanup or visibility

Options:

* show incomplete submissions in admin
* auto-expire after N hours
* allow user to resume payment

***

# ✅ One more piece you still need (critical)

Even with “submission before payment,” you still need:

## 🔁 Payment reconciliation

Because this can still happen:

```
Payment succeeds
💥 crash before you mark payment_complete
```

So you must be able to:

```
→ query payment provider by submission_id / reference
→ update submission later
```

Otherwise:

* you have money collected
* but `payment_complete = false`

***

# ✅ Updated clean mental model (your system now)

```
Submission (durable intent)
    ↓
Payment (external, must be reconciled)
    ↓
Workflow Steps (idempotent + retryable)
    ↓
Completion
```

***

# 💬 Straight answer to your point

> “If it dies during or directly after payment, we still captured Submission Intent”

✅ Yes — and that’s the *most important guarantee* you can add.

It changes your system from:

> ❌ “we might lose what the user tried to do”

to:

> ✅ “we always know what they tried to do, and we can finish it later”

***

# 🧩 If I refine your diagram slightly based on this

The key structural shift is:

```
[Create Submission]  ← Anchor starts here (before payment)
        ↓
[Payment]
        ↓
[All remaining steps]
```

Recovery always attaches to the submission.

***

# ✅ Bottom line

You made the correct call:

* ✅ Submission before payment = durable intent
* ✅ Payment after = recoverable external step
* ✅ Everything after = idempotent workflow

That combination is what actually makes the system resilient to:

* app pool crashes
* mid-request death
* network failures

***

If you want, I can adjust the diagram to **explicitly show the “awaiting payment” vs “processing” state transitions** — that’s usually the next useful step for aligning dev + PM.

***

# Addendum: Reconciliation With Current Global ASP Flow

The current global files already establish the right anchor points:

- `forms.inc` creates or updates `app_renew` before the full submit workflow is complete.
- Payment writes `PAYMENT_STATUS` and `PAYMENT_DATE_TIME` into `app_renew_detail`.
- Page/config branches call the real work: `SaveFiles`, `SaveData`, `PDF`, email page sends, batch transaction helpers, and custom `execute_code`.
- `apprenewadmin.asp` lists existing submissions and derives coarse PDF, Awaiting Updates, and Payment values from `form_pdf`, `mem_await` / `mem_det_await`, and `app_renew_detail`.

The mockup should therefore not be viewed as a replacement form processor. It is a durable execution ledger and clearer admin surface layered around the existing processor.

The missing schema needs to answer four questions the current tables only answer indirectly:

1. What was expected to happen for this form type?
2. Which expectations were triggered by this particular submission's field values?
3. Which operation or item is queued, running, retrying, awaiting review, done, failed, or not needed?
4. What can a recovery worker safely retry without duplicating already completed work?

See `durable_submission_schema_plan.md` for the concrete table proposal and phased implementation plan.
