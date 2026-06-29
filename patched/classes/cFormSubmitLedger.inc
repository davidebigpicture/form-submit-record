<%
'==========================================================================
' cFormSubmitLedger.inc
' Durable form-submission ledger helper (additive observability layer).
'
' Records intent + execution facts for the durable submission pipeline
' described in docs/durable_submission_schema_plan.md and visualized in
' form_submit_tool.html.
'
' DESIGN RULES
'   * Never break the form. Every DB call is wrapped in on-error and any
'     failure (including missing tables before the migration is applied) is
'     swallowed; the worst case is that the dashboard simply has less data.
'   * Idempotent. forms.inc is re-entered by refreshes, redirects, the
'     doPDF/doEmail passes and recovery, so every write is a MERGE/UPDATE
'     keyed by app_renew_id (+ operation_key / item_key).
'   * Self-contained. Opens its own ADODB connection from
'     Application("ConnectionString") unless a connection is supplied via
'     the Conn property.
'
' Requires the Oracle objects created by db/form_submit_ledger.sql.
'==========================================================================

class cFormSubmitLedger
    private mConn
    private mOwnConn
    private mAppRenewID
    private mAppRenewType
    private mMembershipID
    private mLastError
    private mDisabled

    private sub Class_Initialize
        set mConn = nothing
        mOwnConn = false
        mAppRenewID = ""
        mAppRenewType = ""
        mMembershipID = ""
        mLastError = ""
        mDisabled = false
    end sub

    private sub Class_Terminate
        on error resume next
        if mOwnConn then
            if not (mConn is nothing) then
                if mConn.State <> 0 then mConn.Close
            end if
        end if
        set mConn = nothing
        on error goto 0
    end sub

    '---- public properties -------------------------------------------------
    public property let Conn(c)
        set mConn = c
    end property
    public property let AppRenewID(v)
        mAppRenewID = v & ""
    end property
    public property get AppRenewID
        AppRenewID = mAppRenewID
    end property
    public property let AppRenewType(v)
        mAppRenewType = v & ""
    end property
    public property let MembershipID(v)
        mMembershipID = v & ""
    end property
    public property get LastError
        LastError = mLastError
    end property

    ' Pull ids from the ASP session when the caller did not set them.
    public sub RefreshFromSession()
        on error resume next
        if mAppRenewID & "" = "" then mAppRenewID = Session("APP_RENEW_ID") & ""
        if mAppRenewType & "" = "" then mAppRenewType = Session("APP_RENEW_CD") & ""
        if mMembershipID & "" = "" then mMembershipID = Session("MEMBERSHIP_ID") & ""
        on error goto 0
    end sub

    ' Explicit context set; preferred from forms.inc which uses its own
    ' "session" variable (intrinsic in classic mode, CacheManager in redis).
    public sub SetContext(idVal, typeVal, memVal)
        if idVal & "" <> "" then mAppRenewID = idVal & ""
        if typeVal & "" <> "" then mAppRenewType = typeVal & ""
        if memVal & "" <> "" then mMembershipID = memVal & ""
    end sub

    '---- small helpers -----------------------------------------------------
    private function IIf(cond, a, b)
        if cond then IIf = a else IIf = b
    end function

    private function Q(s)
        Q = "'" & replace(s & "", "'", "''") & "'"
    end function

    ' Returns an integer literal string or "null". Digit-sanitized so it is
    ' both injection-safe and free of the clng() overflow that number(12) ids
    ' and large byte sizes would trigger (VBScript Long maxes near 2.1e9).
    private function NumOrNull(v)
        dim raw, k, ch, outv, neg, dotpos
        raw = trim(v & "")
        if raw = "" then
            NumOrNull = "null"
            exit function
        end if
        if not isNumeric(raw) then
            NumOrNull = "null"
            exit function
        end if
        neg = ""
        if left(raw, 1) = "-" then neg = "-"
        dotpos = instr(raw, ".")
        if dotpos > 0 then raw = left(raw, dotpos - 1)
        outv = ""
        for k = 1 to len(raw)
            ch = mid(raw, k, 1)
            if ch >= "0" and ch <= "9" then outv = outv & ch
        next
        if outv = "" then
            NumOrNull = "null"
        else
            NumOrNull = neg & outv
        end if
    end function

    ' Digit-only app_renew_id for use as a numeric SQL literal (overflow-safe).
    private function IDsql()
        dim raw, k, ch, outv
        raw = trim(mAppRenewID & "")
        outv = ""
        for k = 1 to len(raw)
            ch = mid(raw, k, 1)
            if ch >= "0" and ch <= "9" then outv = outv & ch
        next
        if outv = "" then outv = "-1"
        IDsql = outv
    end function

    private function EnsureContext()
        EnsureContext = false
        if mDisabled then exit function
        on error resume next
        RefreshFromSession
        if not (isNumeric(mAppRenewID & "") and (mAppRenewID & "") <> "") then exit function
        if mConn is nothing then
            set mConn = Server.CreateObject("ADODB.Connection")
            mConn.Open Application("ConnectionString")
            mOwnConn = true
        end if
        if err.number <> 0 then
            mLastError = "EnsureContext: " & err.Description
            err.clear
            mDisabled = true   ' avoid hammering a broken connection for the rest of the request
            exit function
        end if
        on error goto 0
        EnsureContext = true
    end function

    private sub ExecSql(sql, label)
        on error resume next
        mConn.Execute sql
        if err.number <> 0 then
            mLastError = label & ": " & err.Description
            err.clear
        end if
        on error goto 0
    end sub

    private function MapDurableState(formStatus)
        dim s : s = ucase(trim(formStatus & ""))
        select case s
            case "PROGRESS", "SAVED", "INVITE"
                MapDurableState = "draft"
            case "PENDING"
                MapDurableState = "processing"
            case "APPROVED", "COMPLETED"
                MapDurableState = "completed"
            case "DENIED"
                MapDurableState = "failed"
            case "DELETED"
                MapDurableState = "abandoned"
            case "PAY_LATER"
                MapDurableState = "awaiting_payment"
            case else
                MapDurableState = "processing"
        end select
    end function

    '---- submission record -------------------------------------------------
    public sub EnsureSubmitRecord(formStatus)
        if not EnsureContext() then exit sub
        dim st : st = MapDurableState(formStatus)
        dim submittedAt
        submittedAt = IIf(ucase(trim(formStatus & "")) = "PENDING", "systimestamp", "null")
        dim sql
        sql = "merge into form_submit_record t " &_
            "using (select " & IDsql() & " app_renew_id from dual) s " &_
            "on (t.app_renew_id = s.app_renew_id) " &_
            "when matched then update set " &_
                "app_renew_type = " & Q(mAppRenewType) & ", " &_
                "membership_id = " & NumOrNull(mMembershipID) & ", " &_
                "legacy_app_renew_status = " & Q(formStatus) & ", " &_
                "submitted_at = nvl(t.submitted_at, " & submittedAt & "), " &_
                "updated_at = systimestamp " &_
            "when not matched then insert " &_
                "(app_renew_id, app_renew_type, membership_id, durable_state, legacy_app_renew_status, ready_to_process, submitted_at, created_at, updated_at) " &_
                "values (" & IDsql() & ", " & Q(mAppRenewType) & ", " & NumOrNull(mMembershipID) & ", " & Q(st) & ", " & Q(formStatus) & ", 'N', " & submittedAt & ", systimestamp, systimestamp)"
        ExecSql sql, "EnsureSubmitRecord"
    end sub

    public sub SyncLegacyStatus(formStatus)
        if not EnsureContext() then exit sub
        ExecSql "update form_submit_record set legacy_app_renew_status = " & Q(formStatus) & ", updated_at = systimestamp where app_renew_id = " & IDsql(), "SyncLegacyStatus"
    end sub

    public sub SetDurableState(state)
        if not EnsureContext() then exit sub
        dim extra : extra = ""
        if lcase(trim(state & "")) = "completed" then extra = ", completed_at = systimestamp"
        ExecSql "update form_submit_record set durable_state = " & Q(state) & extra & ", updated_at = systimestamp where app_renew_id = " & IDsql(), "SetDurableState"
    end sub

    '---- operations --------------------------------------------------------
    public sub EnsureOperation(opKey, configured, sourceScope, sourceLabel)
        if not EnsureContext() then exit sub
        dim cfg : cfg = ucase(left(configured & "N", 1))
        if cfg <> "Y" then cfg = "N"
        dim sql
        sql = "merge into form_submit_operation t " &_
            "using (select " & IDsql() & " app_renew_id, " & Q(opKey) & " operation_key from dual) s " &_
            "on (t.app_renew_id = s.app_renew_id and t.operation_key = s.operation_key) " &_
            "when matched then update set configured = " & Q(cfg) & ", source_scope = " & Q(sourceScope) & ", source_label = " & Q(sourceLabel) & ", updated_at = systimestamp " &_
            "when not matched then insert (operation_id, app_renew_id, operation_key, configured, source_scope, source_label, state, created_at, updated_at) " &_
            "values (seq_form_submit_operation_id.nextval, " & IDsql() & ", " & Q(opKey) & ", " & Q(cfg) & ", " & Q(sourceScope) & ", " & Q(sourceLabel) & ", 'ready', systimestamp, systimestamp)"
        ExecSql sql, "EnsureOperation"
    end sub

    public sub SetOperationState(opKey, state, note)
        if not EnsureContext() then exit sub
        dim sql
        sql = "merge into form_submit_operation t " &_
            "using (select " & IDsql() & " app_renew_id, " & Q(opKey) & " operation_key from dual) s " &_
            "on (t.app_renew_id = s.app_renew_id and t.operation_key = s.operation_key) " &_
            "when matched then update set " &_
                "state = " & Q(state) & ", " &_
                "notes = " & Q(note) & ", " &_
                "started_at = case when t.started_at is null and " & Q(state) & " in ('queued','running','retrying','pending') then systimestamp else t.started_at end, " &_
                "completed_at = case when " & Q(state) & " = 'done' then systimestamp else t.completed_at end, " &_
                "attempts = case when " & Q(state) & " in ('running','retrying') then nvl(t.attempts, 0) + 1 else t.attempts end, " &_
                "updated_at = systimestamp " &_
            "when not matched then insert (operation_id, app_renew_id, operation_key, configured, source_scope, state, notes, created_at, updated_at) " &_
            "values (seq_form_submit_operation_id.nextval, " & IDsql() & ", " & Q(opKey) & ", 'Y', 'form_config', " & Q(state) & ", " & Q(note) & ", systimestamp, systimestamp)"
        ExecSql sql, "SetOperationState"
        AppendEvent opKey, "state_changed", "info", state & IIf(note & "" <> "", " - " & note, "")
    end sub

    public sub SetOperationTrigger(opKey, field, op, value)
        if not EnsureContext() then exit sub
        ExecSql "update form_submit_operation set source_scope = 'triggered', trigger_field = " & Q(field) & ", trigger_operator = " & Q(op) & ", trigger_value = " & Q(value) & ", updated_at = systimestamp where app_renew_id = " & IDsql() & " and operation_key = " & Q(opKey), "SetOperationTrigger"
    end sub

    '---- operation items (files / emails) ----------------------------------
    public sub EnsureOperationItem(opKey, itemKey, itemLabel, sourceScope)
        if not EnsureContext() then exit sub
        EnsureOperation opKey, "Y", sourceScope, opKey
        dim sql
        sql = "merge into form_submit_operation_item t " &_
            "using (select " & IDsql() & " app_renew_id, " & Q(opKey) & " operation_key, " & Q(itemKey) & " item_key from dual) s " &_
            "on (t.app_renew_id = s.app_renew_id and t.operation_key = s.operation_key and t.item_key = s.item_key) " &_
            "when matched then update set item_label = " & Q(itemLabel) & ", source_scope = " & Q(sourceScope) & ", updated_at = systimestamp " &_
            "when not matched then insert (item_id, operation_id, app_renew_id, operation_key, item_key, item_label, source_scope, state, created_at, updated_at) " &_
            "values (seq_form_submit_operation_item_id.nextval, " &_
                "(select operation_id from form_submit_operation where app_renew_id = " & IDsql() & " and operation_key = " & Q(opKey) & "), " &_
                IDsql() & ", " & Q(opKey) & ", " & Q(itemKey) & ", " & Q(itemLabel) & ", " & Q(sourceScope) & ", 'ready', systimestamp, systimestamp)"
        ExecSql sql, "EnsureOperationItem"
    end sub

    public sub SetOperationItemState(opKey, itemKey, state, note)
        if not EnsureContext() then exit sub
        ExecSql "update form_submit_operation_item set state = " & Q(state) & ", notes = " & Q(note) & ", " &_
            "completed_at = case when " & Q(state) & " = 'done' then systimestamp else completed_at end, " &_
            "attempts = case when " & Q(state) & " in ('running','retrying') then nvl(attempts, 0) + 1 else attempts end, " &_
            "updated_at = systimestamp " &_
            "where app_renew_id = " & IDsql() & " and operation_key = " & Q(opKey) & " and item_key = " & Q(itemKey), "SetOperationItemState"
        RecountOperationItems opKey
    end sub

    public sub RecountOperationItems(opKey)
        if not EnsureContext() then exit sub
        ExecSql "update form_submit_operation o set " &_
            "expected_count = (select count(*) from form_submit_operation_item i where i.app_renew_id = o.app_renew_id and i.operation_key = o.operation_key), " &_
            "done_count = (select count(*) from form_submit_operation_item i where i.app_renew_id = o.app_renew_id and i.operation_key = o.operation_key and i.state = 'done'), " &_
            "failed_count = (select count(*) from form_submit_operation_item i where i.app_renew_id = o.app_renew_id and i.operation_key = o.operation_key and i.state = 'failed'), " &_
            "updated_at = systimestamp " &_
            "where o.app_renew_id = " & IDsql() & " and o.operation_key = " & Q(opKey), "RecountOperationItems"
    end sub

    public sub EnsureFileItem(fieldName, originalFileName, tempPath, expectedPath, fileSize)
        if not EnsureContext() then exit sub
        EnsureOperation "files", "Y", "form_config", "Uploaded files"
        dim k : k = fieldName & "|" & originalFileName
        EnsureOperationItem "files", k, originalFileName, "form_config"
        ExecSql "update form_submit_operation_item set temp_location = " & Q(tempPath) & ", expected_location = " & Q(expectedPath) & ", size_bytes = " & NumOrNull(fileSize) & ", updated_at = systimestamp " &_
            "where app_renew_id = " & IDsql() & " and operation_key = 'files' and item_key = " & Q(k), "EnsureFileItem"
    end sub

    public sub SetFileItemState(fieldName, originalFileName, state, location)
        if not EnsureContext() then exit sub
        dim k : k = fieldName & "|" & originalFileName
        dim loc : loc = ""
        if lcase(state) = "done" then loc = ", permanent_location = " & Q(location)
        ExecSql "update form_submit_operation_item set state = " & Q(state) & loc & ", notes = " & Q(location) & ", updated_at = systimestamp " &_
            "where app_renew_id = " & IDsql() & " and operation_key = 'files' and item_key = " & Q(k), "SetFileItemState"
        RecountOperationItems "files"
    end sub

    public sub EnsureEmailItem(emailKey, recipient, subject, sourceScope)
        if not EnsureContext() then exit sub
        EnsureOperation "emails", "Y", sourceScope, "Configured emails"
        EnsureOperationItem "emails", emailKey, recipient & " / " & subject, sourceScope
    end sub

    public sub SetEmailItemState(emailKey, state, note)
        SetOperationItemState "emails", emailKey, state, note
    end sub

    public sub SetEmailItemTrigger(emailKey, field, op, value)
        if not EnsureContext() then exit sub
        ExecSql "update form_submit_operation_item set source_scope = 'triggered', trigger_field = " & Q(field) & ", trigger_operator = " & Q(op) & ", trigger_value = " & Q(value) & ", updated_at = systimestamp " &_
            "where app_renew_id = " & IDsql() & " and operation_key = 'emails' and item_key = " & Q(emailKey), "SetEmailItemTrigger"
    end sub

    '---- events ------------------------------------------------------------
    public sub AppendEvent(opKey, eventType, severity, message)
        if not EnsureContext() then exit sub
        dim opSql
        if opKey & "" = "" then opSql = "null" else opSql = Q(opKey)
        ExecSql "insert into form_submit_event (event_id, app_renew_id, operation_key, event_type, severity, message, created_at) " &_
            "values (seq_form_submit_event_id.nextval, " & IDsql() & ", " & opSql & ", " & Q(eventType) & ", " & Q(severity) & ", " & Q(left(message & "", 3900)) & ", systimestamp)", "AppendEvent"
    end sub

    '---- config snapshot ---------------------------------------------------
    public sub SnapshotConfig(jsonText)
        if not EnsureContext() then exit sub
        dim body : body = jsonText & ""
        if body = "" then body = "{}"
        ' Plain SQL literals cap at 4000 bytes; large snapshots should use a
        ' bound CLOB parameter in a later phase. Cap defensively for now.
        if len(body) > 3900 then body = left(body, 3900)
        dim sql
        sql = "merge into form_submit_config_snapshot t " &_
            "using (select " & IDsql() & " app_renew_id from dual) s on (t.app_renew_id = s.app_renew_id) " &_
            "when matched then update set app_renew_type = " & Q(mAppRenewType) & ", config_json = " & Q(body) & " " &_
            "when not matched then insert (app_renew_id, app_renew_type, config_json, created_at) " &_
            "values (" & IDsql() & ", " & Q(mAppRenewType) & ", " & Q(body) & ", systimestamp)"
        ExecSql sql, "SnapshotConfig"
    end sub

    '---- convenience derivations ------------------------------------------
    ' True when an unreleased Awaiting Updates hold exists for this submission.
    public function HasActiveHold()
        HasActiveHold = false
        if not EnsureContext() then exit function
        on error resume next
        dim rs
        set rs = mConn.Execute("select count(*) c from (" &_
            "select app_renew_id from mem_await where approval_dt is null and app_renew_id = " & IDsql() & " " &_
            "union all " &_
            "select app_renew_id from mem_det_await where approval_dt is null and app_renew_id = " & IDsql() & ")")
        if err.number = 0 then
            if not rs.eof then
                if isNumeric(rs("c") & "") then
                    if clng(rs("c")) > 0 then HasActiveHold = true
                end if
            end if
        else
            err.clear
        end if
        set rs = nothing
        on error goto 0
    end function

    ' Mark records done, or awaiting when an Awaiting Updates hold is present.
    public sub MarkRecordsAfterSave()
        if not EnsureContext() then exit sub
        EnsureOperation "records", "Y", "form_config", "Record updates"
        if HasActiveHold() then
            SetOperationState "records", "awaiting", "Awaiting staff review before record update"
            AppendEvent "records", "awaiting_updates_created", "info", "Active Awaiting Updates hold detected"
            SetDurableState "processing"
        else
            SetOperationState "records", "done", "Record updates completed"
        end if
    end sub

    ' Map a legacy payment status string to durable payment + submission state.
    public sub RecordPaymentStatus(stat)
        if not EnsureContext() then exit sub
        EnsureOperation "payment", "Y", "form_config", "Configured payment page"
        dim s : s = lcase(trim(stat & ""))
        if s = "complete" or s = "completed" or s = "paid" then
            SetOperationState "payment", "done", "Payment confirmed (" & stat & ")"
            SetDurableState "payment_complete"
        elseif s = "pending" then
            SetOperationState "payment", "awaiting", "Payment pending (" & stat & ")"
            SetDurableState "awaiting_payment"
        elseif s = "failed" or s = "declined" or s = "error" then
            SetOperationState "payment", "failed", "Payment failed (" & stat & ")"
            SetDurableState "failed"
        else
            SetOperationState "payment", "running", "Payment status: " & stat
        end if
    end sub

    ' Configured payment page bypassed for this submission (e.g. payment_ty=K).
    public sub MarkPaymentBypassed(triggerField, triggerValue)
        if not EnsureContext() then exit sub
        EnsureOperation "payment", "Y", "triggered", "Configured payment page bypassed"
        if triggerField & "" <> "" then SetOperationTrigger "payment", triggerField, "=", triggerValue
        SetOperationState "payment", "not_needed", "Configured payment page bypassed by alternate Form Pages flow"
    end sub

    ' Boundary record for files. cAppRenew.SaveFiles owns real persistence and
    ' is a global class we cannot edit, so per-file item granularity should be
    ' added inside SaveFiles (or by enumerating session upload arrays) later.
    public sub SyncFilesFromSession()
        if not EnsureContext() then exit sub
        EnsureOperation "files", "Y", "form_config", "Uploaded files"
        SetOperationState "files", "done", "Uploaded files persisted (boundary record)"
    end sub

    '---- end-of-submit reconciliation -------------------------------------
    public sub RecomputeSubmitState()
        if not EnsureContext() then exit sub
        dim a
        a = "(select count(*) from form_submit_operation o where o.app_renew_id = r.app_renew_id"
        dim sql
        sql = "update form_submit_record r set " &_
            "durable_state = (case " &_
                "when " & a & ") = 0 then r.durable_state " &_
                "when " & a & " and o.state = 'failed') > 0 and " & a & " and o.state in ('ready','queued','running','pending','retrying','awaiting')) = 0 then 'failed' " &_
                "when " & a & " and o.operation_key = 'payment' and o.state = 'awaiting') > 0 then 'awaiting_payment' " &_
                "when " & a & " and o.state in ('ready','queued','running','pending','retrying','awaiting')) > 0 then 'processing' " &_
                "else 'completed' end), " &_
            "completed_at = (case when " & a & ") > 0 and " & a & " and o.state in ('ready','queued','running','pending','retrying','awaiting','failed')) = 0 then systimestamp else r.completed_at end), " &_
            "ready_to_process = (case when " & a & " and o.state in ('ready','queued','retrying')) > 0 then 'Y' else r.ready_to_process end), " &_
            "updated_at = systimestamp " &_
            "where r.app_renew_id = " & IDsql()
        ExecSql sql, "RecomputeSubmitState"
    end sub
end class
%>
