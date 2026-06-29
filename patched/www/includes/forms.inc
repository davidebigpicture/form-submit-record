<!-- #include virtual="/www/includes/sanitize.inc" -->
<!-- #include virtual="/classes/cFormSubmitLedger.inc" -->
<%
if request.servervariables("SCRIPT_NAME") <> "/www/epay_result.asp" then
	Response.CacheControl = "no-cache,no-Store"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
end if

if appRenewCd&"" <> "" then
	redis = true
else
	redis = false
	if request.servervariables("SCRIPT_NAME") = "/www/multipay.asp" then
		Set TypeLib = CreateObject("Scriptlet.TypeLib")
		sessionID = replace(Mid(TypeLib.Guid, 2, 36),"-","")
		set TypeLib = nothing		
		set session = Server.CreateObject("CacheManager")
		session.Init sessionID
		session("APP_RENEW_CD") = "TEST"
		session("APP_RENEW_ID") = request.querystring("formID")
		redis = true
	else
		appRenewCd = session("APP_RENEW_CD")
	end if
end if

if redis then
	startSession(0)
end if

if request.querystring("action") = "restore" then		
	session("APP_RENEW_ID") = request.querystring("formID")
end if
	
server.scriptTimeout = 60 * 60

dim cForm, cRenew, connOracle, cSqlForm
dim compPageFound, arrComp

dim pageCd, arrScript, nextPage, requiredMissing, prevPage, arrPages, arrFields, idxField, idxField2, arrFieldNames, arrFieldCols, pageID, formDesc, formPath
dim formStart, formEnd, formOpen, sysNow, sysDate, updateVal, arrSkip, displayType, prodTbl, doSkip
dim reqMisMsg, rsConfig, redirectPage, arrSysVars(20), maxDataLength, multiArray, objFields, cLogIt, sessionTimeout, arrKeys, idxKeys, appRenewID, objFieldsPage, objFieldsType
dim isFormScanner, isDev, doAjax, doRadio, doDrop
dim strReqDebug, officeIP
dim redisDebug, redisDebugKey
dim loadStart, loadEnd, loadTime, timeoutExit, timeoutMessage, timeoutCounter, timeoutLabel
dim pageNoLoadDefaults, pageNoCurrent
dim logFormID
dim inactivityMsg
dim loadDefaultsDelay
dim formRecType
dim redisArr, redisKey
dim cAwait
' === FORM SUBMIT LEDGER PATCH (durable submission observability) ===
' Additive helper; never throws, connects lazily, no-ops if schema absent.
dim cLedger
set cLedger = new cFormSubmitLedger
' === END FORM SUBMIT LEDGER PATCH ===
dim msg
dim cognitoID, redisSSO, rsCognito, cognitoFound, cognito, displayForm, cognitoProviderType
dim msgClear, arrReqOC, idxReqOC

if request("action") = "loadDebug" or request.form("do") = "exit" or request.querystring("preview") = "1" or request.querystring("doPDF") = "1" or request.querystring("doEmail") = "1" then
else
	if session("MEMBERSHIP_ID")&"" <> "" and session("MEMBERSHIP_ID")&"" <> "0" then
		if session("HAS_AWAITING_UPDATES")&"" = "" then
			set cAwait = new cAwaitingUpdatesO
			session("HAS_AWAITING_UPDATES") = cAwait.HasUpdates(session("MEMBERSHIP_ID"))&""
			set cAwait = nothing
		end if
	end if
end if

set cCV = new cCodeValueO
cCV.LoadFromCodeClassAndValue "APP_RENEW_TYPE_CD",appRenewCd
formRecType = cCV.MinorCode
set cCV = nothing

loadDefaultsDelay = False

logFormID = session("APP_RENEW_ID")

if request.form("auto_save") = "1" then
	set cLogIt = new cLog
	cLogIt.WriteLog "forms\form_" & logFormID & ".log","start autoSave"
	set cLogIt = nothing
else
	set cLogIt = new cLog
end if

loadStart = now

doAjax = false
if request.form("method") = "ajax" then
	doAjax = true
end if

isDev = false
if Application("DEV_DOMAIN") = request.servervariables("SERVER_NAME") then
	isDev = true
end if

if request.querystring("isDev") = "false" or session("isDev")&"" = "false" then
	session("isDev") = "false"
	isDev = false
end if

if Application("DEV_DOMAIN") = request.servervariables("SERVER_NAME") then
	if request.querystring("isDev") = "true" or session("isDev")&"" = "true" then
		session("isDev") = "true"
		isDev = true
	end if
end if

dim epayURL
if isDev then
	epayURL = "https://epay2dev.wvsto.com"
else
	epayURL = "https://epay2.wvsto.com"
end if

arrSysVars(0) = "DATETIME"
arrSysVars(1) = "YEAR_TYPE"
arrSysVars(2) = "HOUR_NUM"
arrSysVars(3) = "MONTH_NUM"
arrSysVars(4) = "MONTH_NAME"
arrSysVars(5) = "DAY_NUM"
arrSysVars(6) = "YEAR_NUM"
arrSysVars(7) = "SYSTEM_DATE"
arrSysVars(8) = "LAST_FIRST"
arrSysVars(9) = "LAST_FIRST_TYPE"
arrSysVars(10) = "FULL_NAME"
arrSysVars(11) = "MONTH_NUM_TWO"
arrSysVars(12) = "DAY_NAME"
arrSysVars(13) = "MEMBERSHIP_ID"
arrSysVars(14) = "REDISSESSIONID"
arrSysVars(15) = "MONTH_FIRST_DAY"
arrSysVars(16) = "MONTH_LAST_DAY"
arrSysVars(17) = "YEAR_FIRST_DAY"
arrSysVars(18) = "YEAR_LAST_DAY"
arrSysVars(19) = "DAY_NUM_TWO"
arrSysVars(20) = "LEAP_YEAR"

reqMisMsg = "Required fields missing on a previous page."

prodTbl = "batch_product"
if Application("ORG_ID") = "2" or Application("ORG_ID") = "200" or Application("ORG_ID") = "175" then
	prodTbl = "product"
end if

dbConnect connOracle, Application("ConnectionString")

if session("APP_RENEW_ID")&"" = "" then
	saveFormData("PROGRESS")
end if

if request.querystring("doPDF") = "1" and request.querystring("sessionID") <> "" then
	if redis then
		sessionID = request.querystring("sessionID")

		set session = nothing

		set session = Server.CreateObject("CacheManager")
		session.Init sessionID
	else
		session.abandon
		set redisSession = Server.CreateObject("CacheManager")
		redisSession.Init request.querystring("sessionID")
		arrSesKeys = split(redisSession.GetCacheList,",")
		if isArray(arrSesKeys) then
			for x = lbound(arrSesKeys) to ubound(arrSesKeys)
				item = arrSesKeys(x)
				session(item) = redisSession(item)
			next
		end if
		set redisSession = nothing
	end if

	redisKey = "FORM_" & Application("ORG_ID") & "_" & session("APP_RENEW_ID")
	set redisArr = Server.CreateObject("CacheManager")
	redisArr.SetDefaultExpiration 365,0,0,0
	redisArr.Init redisKey
	arrKeys = split(redisArr.GetCacheList,",")
	if isArray(arrKeys) then
		for x = lbound(arrKeys) to ubound(arrKeys)
			item = arrKeys(x)
			if isArray(redisArr(item)) Then
				if ucase(item) = "ARR_PAGES" or ucase(item) = "ARR_FIELDS" or ucase(item) = "ARR_FIELD_COLS" or ucase(item) = "ARR_COMBINE_FIELDS" or ucase(item) = "ARR_COMBINE_FIELDS_JS" or ucase(item) = "ARR_PAGE" or ucase(item) = "ARR_PAGE_COLS" or ucase(item) = "ARR_CONFIG" or ucase(item) = "ARR_CONFIG_COLS" or ucase(item) = "ARR_BATCH_PRODS" then
				else
					session(item) = redisArr(item)
				end if
			end if
		next
	end if
	set redisArr = nothing
end if

if request.querystring("doPDF") = "1" and request.querystring("table") <> "" then
	session.abandon
	set cSqlForm = new cRunSql
	cSqlForm.Conn = connOracle
	cSqlForm.sqlStr = "select key,value from " & request.querystring("table")
	set rs = cSqlForm.Execute
	do while not rs.eof
	    session(rs("key")&"") = cStr(rs("value")&"")
	    rs.MoveNext
	loop
	rs.close
	set rs = nothing
	set cSqlForm = nothing
end if

if request.querystring("doPDF") = "1" and request.querystring("app_renew_id") <> "" then
	session.abandon
	set cSqlForm = new cRunSql
	cSqlForm.Conn = connOracle
	cSqlForm.sqlStr = "select field_name,data from app_renew_detail where app_renew_id = ?"
	cSqlForm.AddParam(request.querystring("app_renew_id"))
	set rs = cSqlForm.Execute
	do while not rs.eof
	    session(rs("field_name")&"") = cStr(rs("data")&"")
	    rs.MoveNext
	loop
	rs.close
	set rs = nothing
	set cSqlForm = nothing
end if

arrScript = split(request.servervariables("SCRIPT_NAME"),"/")

clientIP = request.servervariables("HTTP_X_FORWARDED_FOR")
if clientIP & " " = " " then
	clientIP = Request.ServerVariables("REMOTE_ADDR")
end if

isFormScanner = false

officeIP = false

if clientIP & " " <> " " then
	if clientIP = "52.27.220.2" or clientIP = "54.184.141.173" or clientIP = "24.111.243.250" or left(clientIP,6) = "172.31" then
		officeIP = true
		if request.querystring("officeIP") = "false" or session("officeIP")&"" = "false" then
			session("officeIP") = "false"
			officeIP = false
		end if
		if request.querystring("officeIP") = "true" or session("officeIP")&"" = "true" then
			officeIP = true
			session("officeIP") = "true"
		end if
	end if
	if clientIP = "64.235.150.121" then
		isFormScanner = true
	else
		'is IP in scanner range of 64.235.153.0/24  
		isFormScanner = compareIpsForm(split(clientIP,"."), Array(64,235,153,0), Array(64,235,153,255))
	end if
end if

pageCd = arrScript(ubound(arrScript))
pageName = pageCd

formLog("querystring = " & request.querystring)

if request.form("do")&"" = "" and request.querystring = "" then
	set cRenew = new cAppRenew
	cRenew.redis = redis
	if cRenew.ConfigOption("login_page")&"" = "" then
		if cRenew.PageOption("order_by") = "1" then
			if isDev then
				%><script>alert("Clearing session because this is the first page of a form with no 'Login Page' set in config.");</script><%
			end if
			cRenew.clearFormSession()
			set cRenew = nothing
			closeAll()
			%><script>window.location = "<%=pageCd%>?reload=1";</script><%
			response.end
		end if
	else
		if cstr(cRenew.PageID(pageCd)) = cRenew.ConfigOption("login_page")&"" then
			if cRenew.ConfigOption("login_clear_session") = "Y" then
				if isDev then
					%><script>alert("Clearing session because 'Clear Session on Login Page' is set to 'Y' in config.");</script><%
				end if
				cRenew.clearFormSession()
				set cRenew = nothing
				closeAll()
				%><script>window.location = "<%=pageCd%>?reload=1";</script><%
				response.end
			end if
		end if
	end if
	set cRenew = nothing
end if

if session("APP_RENEW_CD")&"" <> "" and appRenewCd&"" <> "" then

	'moving from one form type to another
	if request.querystring("do") = "redirect" then
		session("ARR_PAGES") = ""
		session("ARR_FIELDS") = ""
		session("ARR_COMBINE_FIELDS") = ""
		session("ARR_COMBINE_FIELDS_JS") = ""
		session("ARR_FIELD_COLS") = ""
		session("ARR_PAGE") = ""
		session("ARR_PAGE_COLS") = ""
		session("ARR_CONFIG") = ""
		session("ARR_CONFIG_COLS") = ""
		session("ARR_BATCH_PRODS") = ""

		session("DEFAULTS_LOADED") = false
		set cSqlForm = new cRunSql
		cSqlForm.Conn = connOracle
		cSqlForm.SqlStr = "update app_renew set app_renew_type = ? where app_renew_id = ?"
		cSqlForm.AddParam(appRenewCd)
		cSqlForm.AddParam(session("APP_RENEW_ID"))
		cSqlForm.Execute
		set cSqlForm = nothing

		if request.querystring("useother")&"" = "1" and request.querystring("OTHER_ID")&"" <> "" then
			session("MEMBERSHIP_ID") = request.querystring("OTHER_ID")
			parentID = getFieldValue(request.querystring("OTHER_ID"),"ORG_RELATE_ID")
			if parentID&"" <> "" Then
				if instr(parentID,",") Then
					parentID = ""
					session("PARENT_ID") = ""
				else
					session("PARENT_ID") = parentID
					session("PARENT_TYPE") = getFieldValue(parentID,"ORG_SUB_TY_CD")
				end if
			end if
		elseif request.querystring("usechild")&"" = "1" and request.QueryString("child_id")&"" <> "" then
			session("PARENT_ID") = session("MEMBERSHIP_ID")
			set cMember = new cMemberValuesO
			cMember.AddLoadCol("ORG_SUB_TY_CD")
			cMember.ValuesOnly = true
			cMember.LoadFromMembershipID session("PARENT_ID")
			session("PARENT_TYPE") = cMember.Value("ORG_SUB_TY_CD")
			set cMember = nothing
			session("MEMBERSHIP_ID") = request.QueryString("child_id")
			if redis then
				session.remove("CHILD_ID")
				session.remove("CHILD_TYPE")
			else
				session.contents.remove("CHILD_ID")
				session.contents.remove("CHILD_TYPE")
			end if

			set cSqlForm = new cRunSql
			cSqlForm.Conn = connOracle
			cSqlForm.SqlStr = "update app_renew set membership_id = ? where app_renew_id = ?"
			cSqlForm.AddParam(session("MEMBERSHIP_ID"))
			cSqlForm.AddParam(session("APP_RENEW_ID"))
			cSqlForm.Execute
			set cSqlForm = nothing
		elseif request.querystring("usechild")&"" = "1" and session("CHILD_ID")&"" <> "" then
			session("PARENT_ID") = session("MEMBERSHIP_ID")
			set cMember = new cMemberValuesO
			cMember.AddLoadCol("ORG_SUB_TY_CD")
			cMember.ValuesOnly = true
			cMember.LoadFromMembershipID session("PARENT_ID")
			session("PARENT_TYPE") = cMember.Value("ORG_SUB_TY_CD")
			set cMember = nothing
			session("MEMBERSHIP_ID") = session("CHILD_ID")
			if redis then
				session.remove("CHILD_ID")
				session.remove("CHILD_TYPE")
			else
				session.contents.remove("CHILD_ID")
				session.contents.remove("CHILD_TYPE")
			end if

			set cSqlForm = new cRunSql
			cSqlForm.Conn = connOracle
			cSqlForm.SqlStr = "update app_renew set membership_id = ? where app_renew_id = ?"
			cSqlForm.AddParam(session("MEMBERSHIP_ID"))
			cSqlForm.AddParam(session("APP_RENEW_ID"))
			cSqlForm.Execute
			set cSqlForm = nothing
		else
			if request.querystring("CHILD_ID")&"" <> "" then
				session("CHILD_ID") = request.querystring("CHILD_ID")
				set cMember = new cMemberValuesO
				cMember.AddLoadCol("ORG_SUB_TY_CD")
				cMember.ValuesOnly = true
				cMember.LoadFromMembershipID request.querystring("CHILD_ID")
				session("CHILD_TYPE") = cMember.Value("ORG_SUB_TY_CD")
				set cMember = nothing
			end if

			if session("CHILD_ID")&"" <> "" and session("CHILD_TYPE")&"" = "" then
				set cMember = new cMemberValuesO
				cMember.AddLoadCol("ORG_SUB_TY_CD")
				cMember.ValuesOnly = true
				cMember.LoadFromMembershipID session("CHILD_ID")
				session("CHILD_TYPE") = cMember.Value("ORG_SUB_TY_CD")
				set cMember = nothing
			end if
		end if

		saveFormData("PROGRESS")
			
		if request.querystring("CHILD_ID")&"" <> "" then
			%>
			<script>
			var loc = window.location.href.split("&child_id")[0];
			window.location.href = loc;
			</script>
			<%
			response.end
		end if
	end if

	if request.querystring("do") = "redirect" then
		session("ARR_PAGES") = ""
		session("ARR_FIELDS") = ""
		session("ARR_COMBINE_FIELDS") = ""
		session("ARR_COMBINE_FIELDS_JS") = ""
		session("ARR_FIELD_COLS") = ""
		session("ARR_PAGE") = ""
		session("ARR_PAGE_COLS") = ""
		session("ARR_CONFIG") = ""
		session("ARR_CONFIG_COLS") = ""
		session("ARR_BATCH_PRODS") = ""
		session("APP_RENEW_CD") = appRenewCd
		redirectForm(request.ServerVariables("SCRIPT_NAME") & "?skip=1")
		response.end
	end if

	if not doAjax then
		if appRenewCd <> session("APP_RENEW_CD") then
			if Application("ORG_ID") = "200" then
				membershipID = session("MEMBERSHIP_ID")
				session.abandon
				redirectForm(request.servervariables("SCRIPT_NAME") & "?membershipID=" & membershipID & "&formCode=" & appRenewCd)
			else
				set cRenew = new cAppRenew
				cRenew.redis = redis
				cRenew.clearFormSession()
				set cRenew = nothing
				'session.abandon
				set cSqlForm = new cRunSql
				cSqlForm.Conn = connOracle
				cSqlForm.SqlStr = "select login_page from app_renew_config where app_renew_type = ?"
				cSqlForm.AddParam(appRenewCd)
				set rsConfig = cSqlForm.Execute
				set cSqlForm = nothing
				if rsConfig.eof then
					rsConfig.close
					set rsConfig = nothing
				else
					if rsConfig("login_page")&"" <> "" then
						set cSqlForm = new cRunSql
						cSqlForm.Conn = connOracle
						cSqlForm.SqlStr = "select page from app_renew_pages where app_renew_pages_id = ?"
						cSqlForm.AddParam(rsConfig("login_page"))
						set rsConfig = cSqlForm.Execute
						if rsConfig.eof then
						else
							formPath = rsConfig("page")
						end if
						rsConfig.close
						set rsConfig = nothing
						set cSqlForm = nothing
					else
						rsConfig.close
						set rsConfig = nothing
					end if
				end if
				if formPath&"" = "" then
					set cCV = new cCodeValueO
					cCV.LoadFromCodeClassAndValue "APP_RENEW_TYPE_CD", appRenewCd
					formPath = cCV.CodeValueLongDesc
					set cCV = nothing
				end if
				formLog("mismatched form code, redirecting to " & formPath)
				formLog("page code = " & appRenewCd)
				formLog("session code = " & session("APP_RENEW_CD"))
				formLog("session ID = " & session("redissessionid"))
				formLog("cookie ID = " & Request.Cookies("sessionIDwww"))
				if isDev then
					session.abandon
					closeAll()
					%><script>alert("redirecting to <%=formPath%> because a form of a different type is already in progress");document.location = "<%=formPath%>?do=exit";</script><%
					response.end
				else
					session.abandon
					redirectForm(formPath&"?do=exit")
				end if
			end if
		end if
	end if
end if

if request.servervariables("SCRIPT_NAME") = "/www/multipay.asp" then
	set cSqlForm = new cRunSql
	cSqlForm.Conn = connOracle
	cSqlForm.SqlStr = "select app_renew_type from app_renew where app_renew_id = ?"
	cSqlForm.AddParam(request.querystring("formID"))
	set rs = cSqlForm.Execute
	set cSqlForm = nothing
	if rs.eof then
	else
		appRenewCd = rs("app_renew_type")
	end if
	rs.Close
	set rs = nothing
end if

if redis then
	session("APP_RENEW_CD") = appRenewCd
end if

if isArray(session("ARR_PAGES")) Then
	arrPages = session("ARR_PAGES")
else
	set cSqlForm = new cRunSql
	cSqlForm.Conn = connOracle
	cSqlForm.SqlStr = "select page, order_by, name, hide_progress, app_renew_pages_id from app_renew_pages where app_renew_type = ? and order_by >= 0 " &_
		"order by order_by"
	cSqlForm.AddParam(appRenewCd)
	set rs  = cSqlForm.Execute
	set cSqlForm = nothing

	if rs.eof then
		arrPages = ""
	else
		arrPages = rs.GetRows()
	end if
	rs.Close
	set rs = nothing
	session("ARR_PAGES") = arrPages
end if

if not isArray(arrPages) then
	%><div align="center">Pages setup incomplete</div><%
	connOracle.Close
	set connOracle = nothing
	set cRenew = nothing
	response.end
end if

if isArray(session("ARR_FIELD_COLS")) then
	arrFieldCols = session("ARR_FIELD_COLS")
else
	set cSqlForm = new cRunSql
	cSqlForm.Conn = connOracle
	cSqlForm.SqlStr = "select COLUMN_NAME from all_col_comments where table_name = 'APP_RENEW_FIELDS'"
	set rs = cSqlForm.Execute
	set cSqlForm = nothing
	if rs.eof Then
		arrFieldCols = ""
	else
		arrFieldCols = rs.GetRows()
	end if
	rs.Close
	set rs = nothing
	session("ARR_FIELD_COLS") = arrFieldCols
end if

if isArray(arrFieldCols) then
	redim arrFieldNames(ubound(arrFieldCols,2))
	for idxField = lbound(arrFieldCols,2) to ubound(arrFieldCols,2)
		arrFieldNames(idxField) = arrFieldCols(0,idxField)
	next
end if

if isArray(session("ARR_FIELDS")) Then
	arrFields = session("ARR_FIELDS")
else
	set cSqlForm = new cRunSql
	cSqlForm.Conn = connOracle
	cSqlForm.SqlStr = "select "
	for idxField = lbound(arrFieldNames) to ubound(arrFieldNames)
		cSqlForm.AddSqlStr(arrFieldNames(idxField))
		if idxField < ubound(arrFieldNames) then	
			cSqlForm.AddSqlStr(",")
		end if
	next
	cSqlForm.AddSqlStr(" from app_renew_fields where app_renew_type = ?")
	cSqlForm.AddParam(appRenewCd)
	set rs = cSqlForm.Execute
	set cSqlForm = nothing
	if rs.eof Then
		arrFields = ""
	else
		arrFields = rs.GetRows()
	end if
	rs.Close
	set rs = nothing
	session("ARR_FIELDS") = arrFields
end if

set objFields = CreateObject("Scripting.Dictionary")
set objFieldsPage = CreateObject("Scripting.Dictionary")
set objFieldsType = CreateObject("Scripting.Dictionary")
if isArray(arrFields) then
	for idxField = lbound(arrFields,2) to ubound(arrFields,2)
		objFields(clng(arrFields(fieldNameIndex("APP_RENEW_FIELDS_ID"),idxField))) = arrFields(fieldNameIndex("FIELD_NAME"),idxField)
		objFieldsPage(clng(arrFields(fieldNameIndex("APP_RENEW_FIELDS_ID"),idxField))) = arrFields(fieldNameIndex("PAGE"),idxField)
		objFieldsType(clng(arrFields(fieldNameIndex("APP_RENEW_FIELDS_ID"),idxField))) = arrFields(fieldNameIndex("FIELD_TYPE"),idxField)
	next
end if

set cForm = new cFormTools
set cRenew = new cAppRenew

cognito = false
if cstr(cRenew.PageID(pageCd)) = cRenew.ConfigOption("login_page")&"" then
	if cRenew.ConfigOption("login_type") = "EXIST" then
		if getCodeValueDesc("APPLICATION_DB","SSO_COGNITO_WWW")&"" = "1" then
			cognito = true
		end if
	end if
end if

if cRenew.ConfigOption("load_defaults_delay") = "Y" then
	if not session("DEFAULTS_LOADED") then
		loadDefaultsDelay = true
	end if
end if

timeoutCounter = cRenew.ConfigOption("timeout_counter")

if timeoutCounter = "Y" then
	timeoutMessage = "Your session has expired."
	if cRenew.ConfigOption("timeout_message")&"" <> "" then
		timeoutMessage = cRenew.ConfigOption("timeout_message")
	end if

	timeoutLabel = "Session Time Remaining:"
	if cRenew.ConfigOption("timeout_label")&"" <> "" then
		timeoutLabel = cRenew.ConfigOption("timeout_label")
	end if

	timeoutExit = cRenew.ConfigOption("exit_path")
	if timeoutExit&"" = "" then
		if cRenew.ConfigOption("login_page")&"" <> "" then
			timeoutExit = cRenew.PageCD(cRenew.ConfigOption("login_page"))
		end if
		if timeoutExit&"" = "" then
			if cRenew.ConfigOption("load_defaults_page")&"" <> "" then
				timeoutExit = cRenew.PageCD(cRenew.ConfigOption("load_defaults_page"))
			end if
		end if
		if timeoutExit&"" = "" then
			timeoutExit = "/"
		end if
	end if
end if

cRenew.redis = redis

maxDataLength = cRenew.ColumnSize("APP_RENEW_DETAIL","DATA")

displayType  = cRenew.ConfigOption("DISPLAY_TYPE") 
if displayType&"" = "" or displayType&"" = "N" then
  displayType = "TABLE"
end if

if cRenew.ConfigOption("timeout") & "" <> "" then
	sessionTimeout = cint(cRenew.ConfigOption("timeout"))
else
	sessionTimeout = 60
end if

if redis then
	startSession(sessionTimeout)
end if

if cRenew.ConfigOption("sys_date") & "" <> "" then
	if instr(cRenew.ConfigOption("sys_date"),"T") then
		sysNow = cdate(split(cRenew.ConfigOption("sys_date"),"T")(0) & " " & split(cRenew.ConfigOption("sys_date"),"T")(1))
	else
		sysNow = cdate(cRenew.ConfigOption("sys_date"))
	end if
else
	'sysNow = now
	sysNow = dateAdd("h",cRenew.ConfigOption("time_zone_offset"),now)
end if

if cRenew.ConfigOption("sys_date") & "" <> "" then
	if instr(cRenew.ConfigOption("sys_date"),"T") then
		sysDate = cdate(split(cRenew.ConfigOption("sys_date"),"T")(0))
	elseif instr(cRenew.ConfigOption("sys_date")," ") then
		sysDate = cdate(split(cRenew.ConfigOption("sys_date")," ")(0))
	else
		sysDate = cDate(cRenew.ConfigOption("sys_date"))
	end if
else
	'sysDate = date
	sysDate = cdate(split(sysNow," ")(0))
end if

if request("doPDF") = "1" then
	sysDate = FormatDateTime(session("DATETIME"),2)
	sysNow = cdate(session("DATETIME"))
else
	session("DATETIME") = sysNow
end if

if year(sysDate) mod 2 = 0 then
	session("YEAR_TYPE") = "EVEN"
else
	session("YEAR_TYPE") = "ODD"
end if

if year(sysDate) mod 4 = 0 then
	session("LEAP_YEAR") = "Y"
else
	session("LEAP_YEAR") = "N"
end if

session("HOUR_NUM") = hour(sysNow)
session("MONTH_NUM") = month(sysDate)
if month(sysDate) < 10 then
	session("MONTH_NUM_TWO") = "0" & month(sysDate)
else
	session("MONTH_NUM_TWO") = month(sysDate)
end if
session("MONTH_NAME") = monthName(month(sysDate))
session("DAY_NUM") = day(sysDate)
if day(sysDate) < 10 then
	session("DAY_NUM_TWO") = "0" & day(sysDate)
else
	session("DAY_NUM_TWO") = day(sysDate)
end if
session("DAY_NAME") = WeekdayName(Weekday(sysDate))
session("YEAR_NUM") = year(sysDate)
session("MONTH_FIRST_DAY") = cStr(session("MONTH_NUM_TWO") & "/01/" & year(sysDate))
session("MONTH_LAST_DAY") = cStr(dateAdd("d",-1,dateAdd("m",1,session("MONTH_FIRST_DAY"))))
if len(session("MONTH_LAST_DAY")) = 9 then
	session("MONTH_LAST_DAY") = "0" & session("MONTH_LAST_DAY")
end if
session("YEAR_FIRST_DAY") = "01/01/" & year(sysDate)
session("YEAR_LAST_DAY") = "12/31/" & year(sysDate)
session("SYSTEM_DATE") = month(sysDate)&"/"&day(sysDate)&"/"&year(sysDate)
if session("LAST_NAME")&"" <> "" then
	session("LAST_FIRST") = ucase(left(session("LAST_NAME"),1))
	if Asc(session("LAST_FIRST")&"") >= 65 and Asc(session("LAST_FIRST")&"") < 77 then	
		session("LAST_FIRST_TYPE")	= "1"
	else
		session("LAST_FIRST_TYPE")	= "2"
	end if
	if session("FIRST_NAME")&"" <> "" then
		session("FULL_NAME") = session("FIRST_NAME") & " " & session("LAST_NAME")
	end if
end if

formOpen = true

if request("doPDF") = "1" then
else
	if cRenew.ConfigOption("form_avail_start")&"" <> "" then
		formStart = replace(cRenew.ConfigOption("form_avail_start"),"[current year]",year(sysDate))
		if instr(formStart,"[odd year]") then
			if year(sysDate) mod 2 = 0 then
				formStart = replace(formStart,"[odd year]",year(sysDate)-1)
			else
				formStart = replace(formStart,"[odd year]",year(sysDate))
			end if
		end if
		if instr(formStart,"[even year]") then
			if year(sysDate) mod 2 = 0 then
				formStart = replace(formStart,"[even year]",year(sysDate))
			else
				formStart = replace(formStart,"[even year]",year(sysDate)-1)
			end if
		end if
		if isDate(formStart) then
			formStart = cdate(formStart)
		end if
	end if

	if cRenew.ConfigOption("form_avail_end")&"" <> "" then
		formEnd = replace(cRenew.ConfigOption("form_avail_end"),"[current year]",year(sysDate))
		if instr(formEnd,"[odd year]") then
			if year(sysDate) mod 2 = 0 then
				formEnd = replace(formEnd,"[odd year]",year(sysDate)-1)
			else
				formEnd = replace(formEnd,"[odd year]",year(sysDate))
			end if
		end if
		if instr(formEnd,"[even year]") then
			if year(sysDate) mod 2 = 0 then
				formEnd = replace(formEnd,"[even year]",year(sysDate))
			else
				formEnd = replace(formEnd,"[even year]",year(sysDate)-1)
			end if
		end if
		if isDate(formEnd) then
			formEnd = cdate(formEnd)
			if isDate(formStart) then
				if formEnd < formStart then
					if sysDate < formEnd then
						formStart = dateAdd("yyyy",-1,formStart)
					else
						if instr(cRenew.ConfigOption("form_avail_end"),"[even year]") or instr(cRenew.ConfigOption("form_avail_end"),"[odd year]") then
							formEnd = dateAdd("yyyy",2,formEnd)
						else
							formEnd = dateAdd("yyyy",1,formEnd)
						end if
					end if
				end if
			end if
			if instr(cstr(formEnd)," ") then
				if sysNow > formEnd then
					formOpen = false
				end if
			else
				if sysDate > formEnd then
					formOpen = false
				end if
			end if
		end if
	end if

	if cRenew.ConfigOption("form_avail_start")&"" <> "" then
		if instr(cStr(formStart)," ") then
			if sysNow < formStart then
				formOpen = false
			end if
		else
			if sysDate < formStart then
				formOpen = false
			end if
		end if
	end if
end if

if not formOpen then
	if cRenew.ConfigOption("unavail_exceptions")&"" <> "" then
		arrUnavailExceptions = split(cRenew.ConfigOption("unavail_exceptions"),", ")
		for i = lbound(arrUnavailExceptions) to ubound(arrUnavailExceptions)
			if clientIP = arrUnavailExceptions(i) then
				formOpen = true
			end if
		next
	end if
end if

dim cCVTitle
set cCVTitle = new cCodeValueO
cCVTitle.LoadFromCodeClassAndValue "APP_RENEW_TYPE_CD", appRenewCd
formDesc = cCVTitle.CodeValueDesc
set cCVTitle = nothing

globe_title = formDesc
if cRenew.PageOption("name")&"" <> "" then
	globe_title = globe_title & " - " & cRenew.PageOption("name")
end if

pageID = cstr(cRenew.PageID(pageCd))

if request("action") = "loadDebug" or request.form("do") = "exit" or request.querystring("preview") = "1" or request.querystring("doPDF") = "1" or request.querystring("doEmail") = "1" then
else
	if cRenew.ConfigOption("load_defaults_page")&"" <> "" then
		for i = lbound(arrPages,2) to ubound(arrPages,2)
			if arrPages(4,i)&"" = cRenew.ConfigOption("load_defaults_page")&"" then
				pageNoLoadDefaults = cint(arrPages(1,i))
			end if
			if arrPages(4,i)&"" = pageID&"" then
				pageNoCurrent = cint(arrPages(1,i))
			end if
		next
		if pageNoCurrent > pageNoLoadDefaults then
			if session("DEFAULTS_LOADED")&"" = "" then
				inactivityMsg = "The form has timed out due to inactivity. The form will return you to the beginning."
				if isDev then
					inactivityMsg = inactivityMsg & "\n\ncurrent page = " & pageCd & "\n" &_
					"current page order = " & pageNoCurrent & "\n" &_
					"load defaults page = " & cRenew.pageCD(cRenew.ConfigOption("load_defaults_page")) & "\n" &_
					"load defaults page order = " & pageNoLoadDefaults & "\n" &_
					"DEFAULTS_LOADED = " & session("DEFAULTS_LOADED")
				end if
				%><script>alert("<%=inactivityMsg%>");document.location = "<%=aspxPage(cRenew.PageCD(cRenew.ConfigOption("load_defaults_page")))%>";</script><%
				closeAll()
				response.end
			end if
		end if
	end if
end if

if cRenew.ConfigOption("req_check_page")&"" <> "" then
	if request.form("do") = "submit" and pageID&"" = cRenew.ConfigOption("payment_page")&"" then
		requiredMissing = finalFieldCheck
		if requiredMissing&"" <> "" then
%><script>alert("<%=reqMisMsg%>");document.location="<%=aspxPage(requiredMissing)%>";</script><%
			closeAll()
			response.end
		end if
	end if
end if

if request.servervariables("REQUEST_METHOD") = "POST" or request.querystring("invite") = "1" then
	if request.querystring("do") = "batchSave" then		
	else
		saveFormData("PROGRESS")
		if request.form("auto_save") = "1" then
		else
			logPageSubmit
		end if
	end if
end if

dim restoreStatus
if request.querystring("action") = "restore" then		
	set cSqlForm = new cRunSql
	cSqlForm.Conn = connOracle
	cSqlForm.SqlStr = "select membership_id,app_renew_status,app_renew_type from app_renew where app_renew_id = ?"
	cSqlForm.AddParam(request.QueryString("formID"))
	set rs = cSqlForm.Execute
	set cSqlForm = nothing
	if rs.eof Then
	else
		session("APP_RENEW_CD") = rs("app_renew_type")
		if redis then
			appRenewCd = rs("app_renew_type")
		end if
		session("MEMBERSHIP_ID") = rs("membership_id")
		restoreStatus = rs("app_renew_status")
		cRenew.LoadDefaults
	end if
	rs.close
	set rs = nothing
end if

if request.querystring("do") = "submit" and request.querystring("membershipID") <> "" then
	session("MEMBERSHIP_ID") = cStr(request.querystring("membershipID"))
	session("PARENT_ID") = cStr(request.querystring("parentID"))
	session("APP_RENEW_CD") = cStr(request.querystring("formCode"))
	session("APP_RENEW_ID") = ""
	goNext
elseif request.querystring("membershipID") <> "" then
	session("MEMBERSHIP_ID") = cStr(request.querystring("membershipID"))
	session("PARENT_ID") = cStr(request.querystring("parentID"))
	session("APP_RENEW_CD") = cStr(request.querystring("formCode"))
	session("APP_RENEW_ID") = ""
	redirectForm(request.servervariables("SCRIPT_NAME"))
elseif request.form("do") = "submitReload" then
	loginCheck
	cForm.SaveForm
elseif request.form("do") = "submit" then
	loginCheck

	if request.form("form_id")&"" <> "" and session("MEMBERSHIP_ID")&"" <> "" then
		if request.form("form_id")&"" <> session("MEMBERSHIP_ID")&"" then
			formLog("id mismatch form = " & request.form("form_id") & " session = " & session("MEMBERSHIP_ID"))
			session.abandon
		%><script>alert("Form ID mismatch detected.  Only one form per browser is supported.");document.location = "<%if cRenew.ConfigOption("login_page")&"" <> "" then%><%=cRenew.PageCD(cRenew.ConfigOption("login_page"))%><%else%>/<%end if%>";</script><%
		end if
	end if

	'check if form is already in progress when logging in
	if pageID&"" = cRenew.ConfigOption("login_page")&"" then
		if session("DEFAULTS_LOADED") then
			if isDev then
				%><script>alert("redirecting to <%=aspxPage(cRenew.PageCD(cRenew.ConfigOption("login_page")))%> because a form is already in progress");document.location = "<%=aspxPage(cRenew.PageCD(cRenew.ConfigOption("login_page")))%>?do=exit";</script><%
				closeAll()
				response.end
			else
				redirectForm(cRenew.PageCD(cRenew.ConfigOption("login_page")) & "?do=exit")
			end if
		end if
	end if

	'issue found with task 19284 so disabling for now
	'check if form has timed out on load default values page
	if pageID&"" = cRenew.ConfigOption("load_defaults_page")&"" then
		if session("DEFAULTS_LOADED")&"" = "" then
			inactivityMsg = "The form has timed out due to inactivity. The form will return you to the beginning."
			if isDev then
				inactivityMsg = inactivityMsg & "\n\ndo = " & request.form("do") & "\n" &_
				"current page = " & pageCd & "\n" &_
				"load defaults page = " & cRenew.pageCD(cRenew.ConfigOption("load_defaults_page")) & "\n" &_
				"DEFAULTS_LOADED = " & session("DEFAULTS_LOADED")
			end if
			%><script>alert("<%=inactivityMsg%>");document.location = "<%=aspxPage(cRenew.PageCD(cRenew.ConfigOption("load_defaults_page")))%>";</script><%
			closeAll()
			response.end
		end if
	end if

	if session("CARD_NUM") <> "" then
		session("CARD_NUM_RECEIPT") = "****" & right(session("CARD_NUM"),4)
	end if

	'execute code
	formLog("start execute code")
	for idxField = lbound(arrFields,2) to ubound(arrFields,2)
		if arrFields(fieldNameIndex("PAGE"),idxField) = pageCd and arrFields(fieldNameIndex("EXECUTE_CODE"),idxField)&"" <> "" then
			formLog("EXECUTE_CODE = " & arrFields(fieldNameIndex("EXECUTE_CODE"),idxField))
			formLog("FIELD_NAME = " & arrFields(fieldNameIndex("FIELD_NAME"),idxField))
			formLog("DEFAULT_VALUE = " & arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField))
			formLog("form value = " & request.form(arrFields(fieldNameIndex("FIELD_NAME"),idxField)))
			if request.form(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField) then
				dim arrExecute, idxExecute
				arrExecute = split(arrFields(fieldNameIndex("EXECUTE_CODE"),idxField),";")
				formLog("ubound(arrExecute) = " & ubound(arrExecute))
				for idxExecute = lbound(arrExecute) to ubound(arrExecute)
					formLog("idxExecute = " & idxExecute)
					formLog("executing " & arrExecute(idxExecute))
					execute arrExecute(idxExecute)
				next
			end if
		end if
	next
	formLog("end execute code")

	'recaptcha
	for idxField = lbound(arrFields,2) to ubound(arrFields,2)
		if arrFields(fieldNameIndex("PAGE"),idxField) = pageCd and arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "RECAPTCHA" then
			dim cRecaptcha
			set cRecaptcha = new cGoogleRecaptchaNew
			if not cRecaptcha.Verify(Application("GOOGLE_RECAPTCHA_PRIVATE_KEY"),request.Form("g-recaptcha-response"),request.ServerVariables("REMOTE_ADDR")) then
				cForm.AddErrorMsg("reCAPTCHA failed")
			end if
			set cRecaptcha = nothing
			idxField = ubound(arrFields,2) + 1
		end if
	next

	'address verification
	for idxField = lbound(arrFields,2) to ubound(arrFields,2)
		if arrFields(fieldNameIndex("PAGE"),idxField) = pageCd and (arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "ADD_VER" or arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "ADD_VER_REP") then
			verifyAddress arrFields(fieldNameIndex("FIELD_DESCRIPTION"),idxField),arrFields(fieldNameIndex("FIELD_TYPE"),idxField),arrFields(fieldNameIndex("COMBINE_FIELDS"),idxField)
		end if
	next

	'Currency Format fields
	for idxField = lbound(arrFields,2) to ubound(arrFields,2)
		if arrFields(fieldNameIndex("PAGE"),idxField) = pageCd and arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "CURRENCY_FORMAT" then
			if session(cRenew.getFieldName(arrFields(fieldNameIndex("RELATED_FIELD"),idxField)))&"" <> "" then
				session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = FormatCurrency(session(cRenew.getFieldName(arrFields(fieldNameIndex("RELATED_FIELD"),idxField))))
			end if
		end if
	next

	'DEA Number fields
	dim dea1, dea2, dea3, dea4, dea5, dea6, dea7, deaNum, deaErr, deaSum1, deaSum2
	for idxField = lbound(arrFields,2) to ubound(arrFields,2)
		if arrFields(fieldNameIndex("PAGE"),idxField) = pageCd and arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "DEA_NUMBER" then
			if arrFields(fieldNameIndex("TOTAL_MIN_MSG"),idxField)&"" <> "" then
				deaErr = arrFields(fieldNameIndex("TOTAL_MIN_MSG"),idxField)
			else
				deaErr = "Invalid DEA Number value"
			end if
			deaNum = request.form(arrFields(fieldNameIndex("FIELD_NAME"),idxField))
			if deaNum&"" <> "" then
				Set re = New RegExp
				re.Pattern = "[A-Z]{2}\d{7}" 
				if re.Test(deaNum) then
					dea1 = cint(mid(deaNum,3,1))
					dea2 = cint(mid(deaNum,4,1))
					dea3 = cint(mid(deaNum,5,1))
					dea4 = cint(mid(deaNum,6,1))
					dea5 = cint(mid(deaNum,7,1))
					dea6 = cint(mid(deaNum,8,1))
					dea7 = cint(mid(deaNum,9,1))
					deaSum1 = dea1 + dea3 + dea5
					deaSum2 = (dea2 + dea4 + dea6) * 2
					if right(deaNum,1) = right(cstr(deaSum1 + deaSum2),1) then
					else
						cForm.AddErrorMsg(deaErr)
					end if
				else
					cForm.AddErrorMsg(deaErr)
				end if
				set re = nothing
			end if
		end if
	next

	'min/max values for Date fields
	dim minDateVal, maxDateVal, minMonth, minDay, minYear, maxMonth, maxDay, maxYear, arrDtField, idxDtField, dtField
	for idxField = lbound(arrFields,2) to ubound(arrFields,2)
		if arrFields(fieldNameIndex("PAGE"),idxField) = pageCd and arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "DATE" and (arrFields(fieldNameIndex("OPT_VAL_1"),idxField)&"" <> "" or arrFields(fieldNameIndex("OPT_VAL_2"),idxField)&"" <> "" or arrFields(fieldNameIndex("RELATED_FIELD"),idxField)&"" <> "" or arrFields(fieldNameIndex("RELATED_FIELD_2"),idxField)&"" <> "") then
			if session(arrFields(fieldNameIndex("FIELD_NAME"),idxField))&"" <> "" then
				arrDtField = split(session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)),"\\")
				for idxDtField = lbound(arrDtField) to ubound(arrDtField)
					dtField = arrDtField(idxDtField)
					if arrFields(fieldNameIndex("OPT_VAL_1"),idxField)&"" <> "" or arrFields(fieldNameIndex("RELATED_FIELD"),idxField)&"" <> "" then
					  if arrFields(fieldNameIndex("OPT_VAL_1"),idxField)&"" <> "" then
						minDateVal = yearFunction(arrFields(fieldNameIndex("OPT_VAL_1"),idxField))
					  else
						minDateVal = session(cRenew.getFieldName(arrFields(fieldNameIndex("RELATED_FIELD"),idxField)))
					  end if
					  if isDate(minDateVal) then
						minMonth = month(str)
						minYear = year(str)
						minDay = day(str)
						minDateVal = cDate(minDateVal)
						if isDate(dtField) then
							if cDate(dtField) >= minDateVal then
							else
								cForm.AddErrorMsg(arrFields(fieldNameIndex("FIELD_DESCRIPTION"),idxField) & " must be " & minDateVal & " or later.")
							end if
						else
							cForm.AddErrorMsg(arrFields(fieldNameIndex("FIELD_DESCRIPTION"),idxField) & " must be " & minDateVal & " or later.")
						end if
					  end if
					end if
					if arrFields(fieldNameIndex("OPT_VAL_2"),idxField)&"" <> "" or arrFields(fieldNameIndex("RELATED_FIELD_2"),idxField)&"" <> "" then
					  if arrFields(fieldNameIndex("OPT_VAL_2"),idxField)&"" <> "" then
						maxDateVal = yearFunction(arrFields(fieldNameIndex("OPT_VAL_2"),idxField))
					  else
						maxDateVal = session(cRenew.getFieldName(arrFields(fieldNameIndex("RELATED_FIELD_2"),idxField)))
					  end if
					  if isDate(maxDateVal) then
						maxMonth = month(str)
						maxYear = year(str)
						maxDay = day(str)
						maxDateVal = cDate(maxDateVal)
						if isDate(dtField) then
							if cDate(dtField) <= maxDateVal then
							else
								cForm.AddErrorMsg(arrFields(fieldNameIndex("FIELD_DESCRIPTION"),idxField) & " must be " & maxDateVal & " or earlier.")
							end if
						else
							cForm.AddErrorMsg(arrFields(fieldNameIndex("FIELD_DESCRIPTION"),idxField) & " must be " & maxDateVal & " or earlier.")
						end if
					  end if
					end if
				next
			end if
		end if
	next

	'min/max values for total fields
	dim minField, maxField, minVal, maxVal
	for idxField = lbound(arrFields,2) to ubound(arrFields,2)
		if arrFields(fieldNameIndex("PAGE"),idxField) = pageCd and arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "TOTAL" and (arrFields(fieldNameIndex("RELATED_FIELD_2"),idxField)&"" <> "" or arrFields(fieldNameIndex("RELATED_FIELD_3"),idxField)&"" <> "" or arrFields(fieldNameIndex("TOTAL_MIN"),idxField)&"" <> "" or arrFields(fieldNameIndex("TOTAL_MAX"),idxField)&"" <> "") then
			if arrFields(fieldNameIndex("TOTAL_MIN"),idxField)&"" <> "" then
				if isNumeric(request.form(arrFields(fieldNameIndex("FIELD_NAME"),idxField))) then
					if cdbl(request.form(arrFields(fieldNameIndex("FIELD_NAME"),idxField))) < cdbl(arrFields(fieldNameIndex("TOTAL_MIN"),idxField)) then
						if arrFields(fieldNameIndex("TOTAL_MIN_MSG"),idxField)&"" <> "" then
							cForm.AddErrorMsg(arrFields(fieldNameIndex("TOTAL_MIN_MSG"),idxField))
							if isDev then
								cForm.AddErrorMsg(arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " = " & request.form(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) & " minimum = " & arrFields(fieldNameIndex("TOTAL_MIN"),idxField))
							end if
						else
							cForm.AddErrorMsg(URLDecode(arrFields(fieldNameIndex("FIELD_DESCRIPTION"),idxField)) & " does not meet the minimum requirement of " & arrFields(fieldNameIndex("TOTAL_MIN"),idxField))
						end if
					end if
				end if
			elseif arrFields(fieldNameIndex("RELATED_FIELD_2"),idxField)&"" <> "" then
				minField = cRenew.getFieldName(arrFields(fieldNameIndex("RELATED_FIELD_2"),idxField))
				minVal = session(minField)
				if isNumeric(minVal) then
					if cdbl(request.form(arrFields(fieldNameIndex("FIELD_NAME"),idxField))) < cdbl(minVal) then
						if arrFields(fieldNameIndex("TOTAL_MIN_MSG"),idxField)&"" <> "" then
							cForm.AddErrorMsg(arrFields(fieldNameIndex("TOTAL_MIN_MSG"),idxField))
							if isDev then
								cForm.AddErrorMsg(arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " = " & request.form(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) & " minimum = " & minVal)
							end if
						else
							cForm.AddErrorMsg(URLDecode(arrFields(fieldNameIndex("FIELD_DESCRIPTION"),idxField)) & " does not meet the minimum requirement of " & minVal)
						end if
					end if
				end if
			end if
			if arrFields(fieldNameIndex("TOTAL_MAX"),idxField)&"" <> "" then
				if isNumeric(request.form(arrFields(fieldNameIndex("FIELD_NAME"),idxField))) then
					if cdbl(request.form(arrFields(fieldNameIndex("FIELD_NAME"),idxField))) > cdbl(arrFields(fieldNameIndex("TOTAL_MAX"),idxField)) then
						if arrFields(fieldNameIndex("TOTAL_MAX_MSG"),idxField)&"" <> "" then
							cForm.AddErrorMsg(arrFields(fieldNameIndex("TOTAL_MAX_MSG"),idxField))
							if isDev then
								cForm.AddErrorMsg(arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " = " & request.form(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) & " maximum = " & arrFields(fieldNameIndex("TOTAL_MAX"),idxField))
							end if
						else
							cForm.AddErrorMsg(URLDecode(arrFields(fieldNameIndex("FIELD_DESCRIPTION"),idxField)) & " exceeds the maximum requirement of " & arrFields(fieldNameIndex("TOTAL_MAX"),idxField))
						end if
					end if
				end if
			elseif arrFields(fieldNameIndex("RELATED_FIELD_3"),idxField)&"" <> "" then
				maxField = cRenew.getFieldName(arrFields(fieldNameIndex("RELATED_FIELD_3"),idxField))
				maxVal = session(maxField)
				if isNumeric(maxVal) then
					if cdbl(request.form(arrFields(fieldNameIndex("FIELD_NAME"),idxField))) > cdbl(maxVal) then
						if arrFields(fieldNameIndex("TOTAL_MIN_MSG"),idxField)&"" <> "" then
							cForm.AddErrorMsg(arrFields(fieldNameIndex("TOTAL_MIN_MSG"),idxField))
							if isDev then
								cForm.AddErrorMsg(arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " = " & request.form(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) & " maximum = " & maxVal)
							end if
						else
							cForm.AddErrorMsg(URLDecode(arrFields(fieldNameIndex("FIELD_DESCRIPTION"),idxField)) & " exceeds the maximum requirement of " & maxVal)
						end if
					end if
				end if
			end if
		end if
	next

	'look for email confirm fields
	for idxField = lbound(arrFields,2) to ubound(arrFields,2)
		if arrFields(fieldNameIndex("PAGE"),idxField) = pageCd and (arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "PASSWORD_CONFIRM" or arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "EMAIL_CONFIRM") and arrFields(fieldNameIndex("RELATED_FIELD"),idxField)&"" <> "" then
			for idxField2 = lbound(arrFields,2) to ubound(arrFields,2)
				if arrFields(fieldNameIndex("APP_RENEW_FIELDS_ID"),idxField2)&"" = arrFields(fieldNameIndex("RELATED_FIELD"),idxField)&"" then
					if ucase(request.form(arrFields(fieldNameIndex("FIELD_NAME"),idxField2))) <> ucase(request.form(arrFields(fieldNameIndex("FIELD_NAME"),idxField))) then
						cForm.AddErrorMsg(URLDecode(arrFields(fieldNameIndex("FIELD_DESCRIPTION"),idxField)) & " does not match " & URLDecode(arrFields(fieldNameIndex("FIELD_DESCRIPTION"),idxField2)))
					end if
				end if
			next
		end if
	next

	'restricted characters
	if cRenew.ConfigOption("restrict_chars")&"" <> "" Then
		dim restField, idxRest, restChar
		for idxField = lbound(arrFields,2) to ubound(arrFields,2)
			if arrFields(fieldNameIndex("PAGE"),idxField) = pageCd and arrFields(fieldNameIndex("RESTRICT_CHARS"),idxField) = "Y" then
				restField = session(arrFields(fieldNameIndex("FIELD_NAME"),idxField))&""
				for idxRest = 0 to len(cRenew.ConfigOption("restrict_chars"))
					restChar = right(left(cRenew.ConfigOption("restrict_chars"),idxRest+1),1)
					if instr(restField,restChar) Then
						if cRenew.ConfigOption("restrict_chars_msg")&"" <> "" then
							cForm.AddErrorMsg(cRenew.ConfigOption("restrict_chars_msg") & " - " & URLDecode(arrFields(fieldNameIndex("FIELD_DESCRIPTION"),idxField)))
						else
							cForm.AddErrorMsg("Invalid characters entered - " & URLDecode(arrFields(fieldNameIndex("FIELD_DESCRIPTION"),idxField)))
						end if
					end if
				next
			end if
		next
	end if

	'onClick fields to require check for drop and drop uploader fields
	for idxField = lbound(arrFields,2) to ubound(arrFields,2)
		if arrFields(fieldNameIndex("FIELDS_REQUIRE"),idxField)&"" <> "" then
			if session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField) then
				arrReqOC = split(arrFields(fieldNameIndex("FIELDS_REQUIRE"),idxField),", ")
				for idxReqOC = lbound(arrReqOC) to ubound(arrReqOC)
					for idxField2 = lbound(arrFields,2) to ubound(arrFields,2)
						if arrFields(fieldNameIndex("APP_RENEW_FIELDS_ID"),idxField2)&"" = arrReqOC(idxReqOC)&"" and pageCd = arrFields(fieldNameIndex("PAGE"),idxField2) then
							if arrFields(fieldNameIndex("FIELD_TYPE"),idxField2) = "FILE_DRAG" then
								if isArray(session("arr_files_" & arrFields(fieldNameIndex("FIELD_NAME"),idxField2))) then
									if ubound(session("arr_files_" & arrFields(fieldNameIndex("FIELD_NAME"),idxField2)),2) < 0 then
										cForm.AddErrorMsg("File upload for " & arrFields(fieldNameIndex("FIELD_DESCRIPTION"),idxField2) & " required")
									end if
								else
									cForm.AddErrorMsg("File upload for " & arrFields(fieldNameIndex("FIELD_DESCRIPTION"),idxField2) & " required")
								end if
							end if
						end if
					next
				next
			end if
		end if
	next

	numericRangeFields
	numericComparisons
	dateAddFields
	dateRangeFields
	dateComparisons
	dateDifferences
	combineFields(false)
	cRenew.LoadRequiredFields(cForm)
	cRenew.loadFieldTypes(cForm)
	requiredExplainFields
	cForm.Validateform

	if cForm.ErrorMsg = "" then

		if pageID&"" = cRenew.ConfigOption("save_record_page")&"" then
			cRenew.LoadFinalSequences
		end if

		numericComparisons
		dateComparisons
		dateDifferences
		combineFields(false)

		if cRenew.ConfigOption("PDF_GENERATION_TYPE") = "PAGE" then
			'generate PDF for this page
			if cRenew.PageOption("NO_PDF")&"" <> "Y" then
				'cRenew.PDFPage
			end if
		end if

		if pageID&"" = cRenew.ConfigOption("existing_record_page")&"" then
			existingRecordCheck
		end if

		if loadDefaultsDelay = true then
			delayLog "onclick page to redirect to",true
		else
			delayLog "onclick page to redirect to",false
			for idxField = lbound(arrFields,2) to ubound(arrFields,2)
				if arrFields(fieldNameIndex("PAGE"),idxField) = pageCd and arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "RADIO" and arrFields(fieldNameIndex("REDIRECT_PAGE"),idxField)&"" <> "" then
					if request.form(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = yearFunction(arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField)&"") or session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = yearFunction(arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField)&"") then
						if pageID&"" = cRenew.ConfigOption("save_record_page")&"" then
							cRenew.SaveData
						elseif pageID&"" = cRenew.ConfigOption("save_files_page")&"" then
							cRenew.SaveFiles
						end if
						if pageID&"" = cRenew.ConfigOption("pdf_page")&"" then
							cRenew.PDF
						end if
						redirectPage = arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField)
						if isDev and request("doEmail") <> "1" and request("doPDF") <> "1" then
						%><script>alert("redirecting to <%=aspxPage(arrFields(fieldNameIndex("REDIRECT_PAGE"),idxField))%> because <%=arrFields(fieldNameIndex("FIELD_NAME"),idxField)%> = <%=yearFunction(arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField))%>");document.location = "<%=aspxPage(arrFields(fieldNameIndex("REDIRECT_PAGE"),idxField))%>?redirect=1";</script><%
							closeAll()
							response.end
						else
							redirectForm(redirectPage)
						end if
					end if
				end if
			next
		end if
	end if

	if pageID&"" = cRenew.ConfigOption("login_page")&"" then
		if cRenew.ConfigOption("login_type") = "USER" or (cRenew.ConfigOption("login_type") = "USER_GSSO" and session("GOOGLE_EMAIL")&"" <> "") then
			if doAjax then
				session("MEMBERSHIP_ID") = ""
			end if
			loginUser request.form("username"), request.form("password")
		end if
		if cRenew.ConfigOption("login_type") = "EXIST" or cRenew.ConfigOption("login_type") = "EXIST_IF" then
			if doAjax then
				session("MEMBERSHIP_ID") = ""
			end if
			loginExist
			if cognito and session("COGNITO_ID")&"" <> "" Then
				formLog("COGNITO_ID = " & session("COGNITO_ID"))
				set cMember = new cMemberValuesO
				cMember.AddLoadCol("COGNITO_ID")
				cMember.ValuesOnly = true	
				cMember.LoadFromMembershipID session("MEMBERSHIP_ID")
				if cMember.Value("COGNITO_ID")&"" <> session("COGNITO_ID") Then	
					formLog("Updating COGNITO_ID on MEMBERSHIP record")
					cMember.UpdateValue session("MEMBERSHIP_ID"),"COGNITO_ID",session("COGNITO_ID")
				else
					formLog("COGNITO_ID already set on MEMBERSHIP record")
				end if
				set cMember = nothing
				setCognitoCookie
			end if
			if session("MEMBERSHIP_ID")&"" = "" or session("MEMBERSHIP_ID") = "0" then
			else
				if cRenew.ConfigOption("TWO_FACTOR_EMAIL_FIELD")&"" <> "" or cRenew.ConfigOption("TWO_FACTOR_PHONE_FIELD")&"" <> "" then
					redirectForm("/www/twofactor/")
				end if
			end if
			if doAjax then
				response.clear
				if session("MEMBERSHIP_ID")&"" <> "" then
			        Response.Status = "200 - Success"
					response.write session("MEMBERSHIP_ID")
				else
			        response.Status = "401 - Invalid credentials"
				end if
				closeAll()
				response.end
			end if
		end if
		if cRenew.ConfigOption("login_type") = "NEW" then
			loginSaved
		end if
	end if
elseif request.querystring("action") = "loadDebug" then		
	loadDebug
	loginCheck
elseif request.querystring("do") = "batchSave" then		
	loadDebug
	if request.querystring("addSession")&"" <> "" then
		dim arrAddSes
		arrAddSess = split(request.querystring("addSession"),":")
		session(arrAddSess(0)) = arrAddSess(1)
	end if
	if request.querystring("await") = "1" then
		session("DO_AWAIT") = true
	end if
	set cRenew = nothing

	set cRenew = new cAppRenew
	cRenew.redis = redis
	cRenew.SaveData
	if request.querystring("pdf") = "1" then
		cRenew.PDF
	end if
	batchTransaction
	set cSqlForm = new cRunSql
	cSqlForm.Conn = connOracle
	cSqlForm.SqlStr = "update app_renew set app_renew_status = 'APPROVED' where app_renew_id = ?"
	cSqlForm.AddParam(request.querystring("formID"))
	cSqlForm.Execute
	set cSqlForm = nothing
elseif request.querystring("ExternalAuthID")&"" <> "" and cRenew.ConfigOption("login_type") = "USER_GSSO" then
	if doAjax then
		session("MEMBERSHIP_ID") = ""
	end if
	loginGoogleSSO
else
	formLog("start loading page")
	if cRenew.ConfigOption("log_session_page") = "Y" then
		logSession("top")
	end if
	if cRenew.ConfigOption("payment_processor") = "WVEPAY_QUICKPOST" and request.querystring("ID") <> "" then
		appRenewID = session("APP_RENEW_ID")
		if left(request.servervariables("SCRIPT_NAME"),16) = "/www/epay_result" then
			session("WV_EPAY_ID") = cStr(request.querystring("ID"))
			if finishWVEpay then
				cRenew.SaveData
				batchTransaction

				set cSqlForm = new cRunSql
				cSqlForm.Conn = connOracle
				cSqlForm.SqlStr = "update app_renew set app_renew_status = 'PENDING', membership_id = ? where app_renew_id = ?"
				cSqlForm.AddParam(session("MEMBERSHIP_ID"))
				cSqlForm.AddParam(appRenewID)
				cSqlForm.Execute
				set cSqlForm = nothing

				if session("PARENT_ID")&"" <> "" then
					set cSqlForm = new cRunSql
					cSqlForm.Conn = connOracle
					cSqlForm.SqlStr = "delete from app_renew_detail where app_renew_id = ? and field_name = 'PARENT_ID'"
					cSqlForm.AddParam(appRenewID)
					cSqlForm.Execute
					set cSqlForm = nothing

					set cSqlForm = new cRunSql
					cSqlForm.Conn = connOracle
					cSqlForm.SqlStr = "insert into app_renew_detail (app_renew_id,field_name,data) values (?,'PARENT_ID',?)"
					cSqlForm.AddParam(appRenewID)
					cSqlForm.AddParam(session("PARENT_ID"))
					cSqlForm.Execute
					set cSqlForm = nothing
				end if

				if cRenew.ConfigOption("pdf_page")&"" <> "" then
					cRenew.PDFQuickPost
				end if
				closeAll()
				response.end
			end if
		else
			if finishWVEpay then
				set cSqlForm = new cRunSql
				cSqlForm.Conn = connOracle
				cSqlForm.SqlStr = "select membership_id from app_renew where app_renew_id = ?"
				cSqlForm.AddParam(appRenewID)
				set rs = cSqlForm.Execute
				if rs.eof then
				else
					session("MEMBERSHIP_ID") = cStr(rs("MEMBERSHIP_ID"))
				end if
				rs.close
				set rs = nothing
				set cSqlForm = nothing
			else
				prevPage = getPrevPage
				%><script language="Javascript">document.location = "<%=aspxPage(prevPage)%>";</script><%
				closeAll()
				response.end
			end if
		end if
	end if
	if cRenew.ConfigOption("payment_processor") = "WVEPAY" and request.querystring("ID") <> "" then
		session("WV_EPAY_ID") = cStr(request.querystring("ID"))
		dim prevPageID

		prevPage = getPrevPage
		nextPage = getNextPage

		prevPageId = cRenew.PageID(prevPage)
		if prevPageID&"" = cRenew.ConfigOption("payment_page")&"" then
			if session("MEMBERSHIP_ID")&"" = "" or session("MEMBERSHIP_ID") = "0" then
				set cNewMem = new cNewMemberO
				session("MEMBERSHIP_ID") = cNewMem.MembershipID
				membershipID = session("MEMBERSHIP_ID")
				cNewMem.SetValue "ORG_SUB_TY_CD",formRecType
				if cRenew.configOption("approve_app") = "N" then
					cNewMem.setValue "M_STAT_CD", "ACTIVE"
				else
					cNewMem.setValue "M_STAT_CD", "PENDING"
				end if
				cNewMem.InsertRecord
				set cNewMem = nothing
			end if
			if finishWVEpay then
				if pageID&"" = cRenew.ConfigOption("save_pending_page")&"" then
					if cRenew.ConfigOption("approve_app") = "N" then
						saveFormData("COMPLETED")
					else
						saveFormData("PENDING")
					end if
				else
					if prevPageID&"" = cRenew.ConfigOption("save_pending_page")&"" then
						saveFormData("PENDING")
					else
					end if
				end if
				'cRenew.SaveData
				batchTransaction
				redirectForm(nextPage & "?do=submit")
			else
			%><script language="Javascript">document.location = "<%=aspxPage(prevPage)%>";</script><%
				closeAll()
				response.end
			end if
		end if
	end if

	if instr(lcase(request.servervariables("SCRIPT_NAME")),"preview") then
	else
		setRequiredExplainFields

		loginCheck

		if pageID&"" = cRenew.ConfigOption("load_defaults_page")&"" then
			if request("doPDF") = "1" then
			else
				if not session("DEFAULTS_LOADED") then
					cRenew.LoadDefaults
					numericComparisons
					dateComparisons
					dateDifferences
					session("DEFAULTS_LOADED") = true
					if cRenew.ConfigOption("load_defaults_delay") = "Y" then
						loadDefaultsDelay = false
					end if
				end if
			end if
		end if
	end if

	updateFields false

	combineFields(false)

	if instr(lcase(request.servervariables("SCRIPT_NAME")),"preview") then
	else
		if request("doPDF")&"" = "1" or request("doEmail")&"" = "1" then
		else
			session("CARD_AMOUNT") = amountDue
		end if
	end if
end if

if request.querystring("action") = "restore" then		
	set cSqlForm = new cRunSql
	cSqlForm.Conn = connOracle
	cSqlForm.SqlStr = "select d.field_name,d.data from app_renew_page_sub_detail d, app_renew_page_submit s where d.app_renew_page_submit_id = s.app_renew_page_submit_id and s.app_renew_id = ? and d.data_type = 'RF'"
	cSqlForm.AddParam(request.QueryString("formID"))
	set rs = cSqlForm.Execute
	set cSqlForm = nothing
	do while not rs.eof
		session(rs("field_name")) = rs("data")
		rs.MoveNext
	loop
	rs.close
	set rs = nothing
	saveFormData(restoreStatus)
end if

if request("doPDF") = "1" then
else

	numericRangeFields
	dateAddFields
	dateRangeFields
	numericComparisons
	dateComparisons
	dateDifferences
	combineFields(false)

	'child form link fields
	dim childLink, arrChild, idxChild, childIDField, childLinkField
	for idxField = lbound(arrFields,2) to ubound(arrFields,2)
		if arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "CHILD_FORM_LINK" then
			childIDField = cRenew.getFieldName(arrFields(fieldNameIndex("RELATED_FIELD"),idxField))
			arrChild = split(session(childIDField),"\\")

			childLink = ""

			for idxChild = lbound(arrChild) to ubound(arrChild)
				childLink = childLink & "<a href=""" & arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField) & "?do=redirect&usechild=1&child_id=" & arrChild(idxChild) & """>" & arrFields(fieldNameIndex("FIELD_DESCRIPTION"),idxField) & "</a>"
				if idxChild < ubound(arrChild) then
					childLink = childLink & "\\"
				end if
			next
			childLinkField = cRenew.getFieldName(arrFields(fieldNameIndex("RELATED_FIELD_2"),idxField))
			session(childLinkField) = childLink		
		end if
	next

	'other form link fields
	dim otherLink, arrOther, idxOther, otherIDField, otherLinkField
	for idxField = lbound(arrFields,2) to ubound(arrFields,2)
		if arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "OTHER_FORM_LINK" then
			otherIDField = cRenew.getFieldName(arrFields(fieldNameIndex("RELATED_FIELD"),idxField))
			arrOther = split(session(otherIDField),"\\")

			otherLink = ""

			for idxOther = lbound(arrOther) to ubound(arrOther)
				otherLink = otherLink & "<a href=""" & arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField) & "?do=redirect&useother=1&other_id=" & arrOther(idxOther) & """>" & arrFields(fieldNameIndex("FIELD_DESCRIPTION"),idxField) & "</a>"
				if idxOther < ubound(arrOther) then
					otherLink = otherLink & "\\"
				end if
			next
			otherLinkField = cRenew.getFieldName(arrFields(fieldNameIndex("RELATED_FIELD_2"),idxField))
			session(otherLinkField) = otherLink		
		end if
	next

	if loadDefaultsDelay = true then
		delayLog "client-side skip",true
	else
		delayLog "client-side skip",false
		for idxField = lbound(arrFields,2) to ubound(arrFields,2)
			if arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "RADIO" and arrFields(fieldNameIndex("PAGES_SKIP"),idxField)&"" <> "" then
				doSkip = false
				if arrFields(fieldNameIndex("RADIO_NOT_EQUAL"),idxField) = "Y" then
					if session(arrFields(fieldNameIndex("FIELD_NAME"),idxField))&"" <> yearFunction(arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField)&"") then
						doSkip = true
					end if
				else
					if session(arrFields(fieldNameIndex("FIELD_NAME"),idxField))&"" = yearFunction(arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField)&"") then
						doSkip = true
					end if
				end if
				if doSkip then
					arrSkip = split(arrFields(fieldNameIndex("PAGES_SKIP"),idxField),", ")
					for i = lbound(arrSkip) to ubound(arrSkip)
						if ucase(pageName) = ucase(arrSkip(i)) then
							session("SKIPPED_" & pageName) = "1"
							if isDev and request("doEmail") <> "1" and request("doPDF") <> "1" then
							%><script>alert("skipping <%=aspxPage(pageName)%> because <%=arrFields(fieldNameIndex("FIELD_NAME"),idxField)%> <%if arrFields(fieldNameIndex("RADIO_NOT_EQUAL"),idxField) = "Y" then%><><%else%>=<%end if%> <%=yearFunction(arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField))%>");document.location = "<%=aspxPage(getNextPage)%>?skip=1&tim=<%=dateDiff("s","01/01/1970",now())%>";</script><%
								closeAll()
								response.end
							else
								redirectForm(aspxPage(getNextPage) & "?skip=1&tim=" & dateDiff("s","01/01/1970",now()))
							end if
						end if
					next
				end if
			end if
		next
	end if
end if

formLog("start loading custom code")
dim clientFSO, clientTS, clientFileName, clientFileContents
clientFileContents = ""
clientFileName = request.ServerVariables("APPL_PHYSICAL_PATH") & "\includes\client_forms.inc"
Set clientFSO = Server.CreateObject("Scripting.FileSystemObject")
if clientFSO.FileExists(clientFileName) then
	Set clientTS = clientFSO.OpenTextFile(clientFileName, 1, false)
	If Not clientTS.AtEndOfStream  Then
		clientFileContents = clientTS.ReadAll
	End If
	clientTS.Close
	Set clientTS = Nothing
end if
Set clientFSO = Nothing
if instr(clientFileContents,"""" & appRenewCd & """") > 0 then
	formLog("custom code exists for this form")
end if
%>
<!--#include virtual="/includes/client_forms.inc"-->
<%
formLog("end loading custom code")

if request.form("do") = "submit" or request.querystring("do") = "submit" then
	if cRenew.ConfigOption("payment_processor") = "KANPAY" then
		prevPage = getPrevPage

		prevPageId = cRenew.PageID(prevPage)
		if prevPageID&"" = cRenew.ConfigOption("payment_page")&"" then
			batchTransaction
		end if
	end if
	if pageID&"" = cRenew.ConfigOption("payment_page")&"" then

		if cRenew.ConfigOption("payment_processor") = "KANPAY" or cRenew.ConfigOption("payment_processor") = "STRIPE" or cRenew.ConfigOption("payment_processor") = "WVEPAY" or cRenew.ConfigOption("payment_processor") = "WVEPAY_QUICKPOST" then
		else
			validateCCFields
		end if

		if cForm.ErrorMsg = "" then
			paymentStatus "Pending"
			session("CARD_CONFIRMATION") = getSequence("SEQ_VERISIGN_LOG_ID")
			if cRenew.ConfigOption("payment_processor") = "PAYPAL" then
				if PayPal then
					paymentStatus "Complete"
					if pageID&"" = cRenew.ConfigOption("save_record_page")&"" then
						cRenew.SaveData
					elseif pageID&"" = cRenew.ConfigOption("save_files_page")&"" then
						cRenew.SaveFiles
					end if
					batchTransaction
				else
					paymentStatus "Failed"
				end if
			elseif cRenew.ConfigOption("payment_processor") = "PAYPAL_WPP" then
				if PayPalWPP then
					paymentStatus "Complete"
					if pageID&"" = cRenew.ConfigOption("save_record_page")&"" then
						cRenew.SaveData
					elseif pageID&"" = cRenew.ConfigOption("save_files_page")&"" then
						cRenew.SaveFiles
					end if
					batchTransaction
				else
					paymentStatus "Failed"
				end if
			elseif cRenew.ConfigOption("payment_processor") = "PROFITSTARS" then
				if profitStars then
					paymentStatus "Complete"
					if pageID&"" = cRenew.ConfigOption("save_record_page")&"" then
						cRenew.SaveData
					elseif pageID&"" = cRenew.ConfigOption("save_files_page")&"" then
						cRenew.SaveFiles
					end if
					batchTransactionPending
				else
					paymentStatus "Failed"
				end if
			elseif cRenew.ConfigOption("payment_processor") = "AUTHORIZE" then
				if authorize then
					paymentStatus "Complete"
					if pageID&"" = cRenew.ConfigOption("save_record_page")&"" then
						cRenew.SaveData
					elseif pageID&"" = cRenew.ConfigOption("save_files_page")&"" then
						cRenew.SaveFiles
					end if
					batchTransaction
				else
					paymentStatus "Failed"
				end if
			elseif cRenew.ConfigOption("payment_processor") = "ELAVON" then
				if elavon then
					paymentStatus "Complete"
					if pageID&"" = cRenew.ConfigOption("save_record_page")&"" then
						cRenew.SaveData
					elseif pageID&"" = cRenew.ConfigOption("save_files_page")&"" then
						cRenew.SaveFiles
					end if
					batchTransaction
				else
					paymentStatus "Failed"
				end if
			elseif cRenew.ConfigOption("payment_processor") = "MONERIS" then
				if moneris then
					paymentStatus "Complete"
					if pageID&"" = cRenew.ConfigOption("save_record_page")&"" then
						cRenew.SaveData
					elseif pageID&"" = cRenew.ConfigOption("save_files_page")&"" then
						cRenew.SaveFiles
					end if
					batchTransaction
				else
					paymentStatus "Failed"
				end if
			elseif cRenew.ConfigOption("payment_processor") = "CLOVER" then
				if clover then
					paymentStatus "Complete"
					if pageID&"" = cRenew.ConfigOption("save_record_page")&"" then
						cRenew.SaveData
					elseif pageID&"" = cRenew.ConfigOption("save_files_page")&"" then
						cRenew.SaveFiles
					end if
					batchTransaction
				else
					paymentStatus "Failed"
				end if
			elseif cRenew.ConfigOption("payment_processor") = "WVEPAY" or cRenew.ConfigOption("payment_processor") = "WVEPAY_QUICKPOST" then
				if pageID&"" = cRenew.ConfigOption("save_record_page")&"" then
					cRenew.SaveData
				end if
				if pageID&"" = cRenew.ConfigOption("pdf_page")&"" then
					if session("MEMBERSHIP_ID")&"" = "" or session("MEMBERSHIP_ID") = "0" then
						set cNewMem = new cNewMemberO
						session("MEMBERSHIP_ID") = cNewMem.MembershipID
						membershipID = session("MEMBERSHIP_ID")
						cNewMem.SetValue "ORG_SUB_TY_CD",formRecType
						if cRenew.configOption("approve_app") = "N" then
							cNewMem.setValue "M_STAT_CD", "ACTIVE"
						else
							cNewMem.setValue "M_STAT_CD", "PENDING"
						end if
						cNewMem.InsertRecord
						set cNewMem = nothing
					end if
					if pageID&"" = cRenew.ConfigOption("save_pending_page")&"" then
						saveFormData("PENDING")
					else
					end if
					cRenew.PDF
				else
					if pageID&"" = cRenew.ConfigOption("save_pending_page")&"" then
						saveFormData("PENDING")
					else
					end if
				end if
				if cint("0" & session("CARD_AMOUNT")) > 0 then
					if session("pType")&"" <> "" then
						if session("pType") = "1" then
							nextPage = getNextPage
							redirectForm(nextPage & "?do=submit")
						else
							wvEPay
						end if
					else
						wvEPay
					end if
				else
					nextPage = getNextPage
					if pageID&"" = cRenew.ConfigOption("save_pending_page")&"" then
						if cRenew.ConfigOption("approve_app") = "N" then
							saveFormData("COMPLETED")
						else
							saveFormData("PENDING")
						end if
					end if
					cRenew.SaveData
					batchTransaction
					redirectForm(nextPage & "?do=submit")
				end if
			elseif cRenew.ConfigOption("payment_processor") = "KANPAY" then
				kanpay
			elseif cRenew.ConfigOption("payment_processor") = "STRIPE" then
				if stripe then
					paymentStatus "Complete"
					if pageID&"" = cRenew.ConfigOption("save_record_page")&"" then
						cRenew.SaveData
					elseif pageID&"" = cRenew.ConfigOption("save_files_page")&"" then
						cRenew.SaveFiles
					end if
					batchTransaction
				else
					paymentStatus "Failed"
				end if
			else	
				batchTransaction
			end if
		end if
	end if

	if cForm.ErrorMsg = "" then
		if pageID&"" = cRenew.ConfigOption("req_check_page")&"" or cRenew.PageOption("req_check_page") = "Y" then
			requiredMissing = finalFieldCheck
		end if
		
		if requiredMissing&"" <> "" then
		else
			if cRenew.PageOption("execute_code")&"" <> "" then
				' === FORM SUBMIT LEDGER PATCH ===
				cLedger.SetContext session("APP_RENEW_ID"), appRenewCd, session("MEMBERSHIP_ID")
				cLedger.SetOperationState "custom_code", "running", "Executing configured page code"
				' === END FORM SUBMIT LEDGER PATCH ===
				execute cRenew.PageOption("execute_code")
				' === FORM SUBMIT LEDGER PATCH ===
				cLedger.SetOperationState "custom_code", "done", "Custom code completed"
				' === END FORM SUBMIT LEDGER PATCH ===
			end if

			if pageID&"" <> cRenew.ConfigOption("payment_page")&"" then
				if pageID&"" = cRenew.ConfigOption("save_record_page")&"" then
					' === FORM SUBMIT LEDGER PATCH ===
					cLedger.SetContext session("APP_RENEW_ID"), appRenewCd, session("MEMBERSHIP_ID")
					cLedger.SetOperationState "records", "running", "Updating records (save_record_page)"
					' === END FORM SUBMIT LEDGER PATCH ===
					cRenew.SaveData
					' === FORM SUBMIT LEDGER PATCH ===
					cLedger.MarkRecordsAfterSave
					' === END FORM SUBMIT LEDGER PATCH ===
				elseif pageID&"" = cRenew.ConfigOption("save_files_page")&"" then
					' === FORM SUBMIT LEDGER PATCH ===
					cLedger.SetContext session("APP_RENEW_ID"), appRenewCd, session("MEMBERSHIP_ID")
					cLedger.SetOperationState "files", "running", "Saving uploaded files (save_files_page)"
					' === END FORM SUBMIT LEDGER PATCH ===
					cRenew.SaveFiles
					' === FORM SUBMIT LEDGER PATCH ===
					' Boundary record; per-file items require deeper SaveFiles hooks.
					cLedger.SyncFilesFromSession
					' === END FORM SUBMIT LEDGER PATCH ===
				end if
			end if
			if pageID&"" = cRenew.ConfigOption("pdf_page")&"" then
				' === FORM SUBMIT LEDGER PATCH ===
				cLedger.SetContext session("APP_RENEW_ID"), appRenewCd, session("MEMBERSHIP_ID")
				cLedger.SetOperationState "pdf", "running", "Generating PDF (pdf_page)"
				' === END FORM SUBMIT LEDGER PATCH ===
				cRenew.PDF
				' === FORM SUBMIT LEDGER PATCH ===
				cLedger.SetOperationState "pdf", "done", "PDF generated (pdf_page)"
				' === END FORM SUBMIT LEDGER PATCH ===
			end if

			if pageID&"" = cRenew.ConfigOption("save_pending_page")&"" then
				if cRenew.ConfigOption("approve_app") = "N" then
					saveFormData("COMPLETED")
				else
					saveFormData("PENDING")
				end if
			end if

			arrComp = session("ARR_PAGES_COMPLETE")
			if isArray(arrComp) then
				compPageFound = false

				for i = lbound(arrComp) to ubound(arrComp)
					if arrComp(i) = pageName then
						compPageFound = true
					end if
				next
				if compPageFound then
				else
					i = ubound(arrComp) + 1
					redim preserve arrComp(i)
					arrComp(i) = pageName
				end if
			else
				redim arrComp(0)
				arrComp(0) = pageName
			end if
			session("ARR_PAGES_COMPLETE") = arrComp
			goNext
		end if
	end if
end if

%><!--#include virtual="/header.asp"--><%

dim version
version = getCodeValueDesc("APPLICATION_DB","VERSION")
Application("VERSION") = version
session("VERSION") = version
if (isDev or officeIP) and request("doEmail") <> "1" and request("doPDF") <> "1" then
	%><div class="debug">Version <%=version%></div><%
end if

%><script>var formID = "<%=session("APP_RENEW_ID")%>";var redis = <%=lcase(redis&"")%>;var isDev = <%=lcase(isDev)%>;<%if isDev then%>var resourcePath = "<%=replace(Application("ResourcePath"),"\","\\")%>";<%end if%></script><%

if formOpen then

%><div class="form-display"><%
if isDev and request("doPDF") = "1" and request("doEmail") <> "1" then
	%><div align="center"><%=pageName%></div><%
end if
if isDev and request("doEmail") <> "1" and request("doPDF") <> "1" then
	%><div class="debug"><%
	if isArray(arrPages) then
		%>Go to: <select name="pages" id="pages" onchange="goPage(this);"><%
		for i = lbound(arrPages,2) to ubound(arrPages,2)
			%><option value="<%=aspxPage(arrPages(0,i))%>"<%if arrPages(0,i) = pageName then%> selected<%end if%>><%=replace(arrPages(0,i),".asp","")%> - <%=arrPages(2,i)%></option><%
		next
		%></select><%
	end if
	%><br><br><input type="checkbox" name="doPDF" id="doPDF" onclick="document.location='<%=aspxPage(request.servervariables("SCRIPT_NAME"))%><%if request.querystring("doPDF") = "1" then%><%else%>?doPDF=1<%end if%>';"<%if request.querystring("doPDF") = "1" then%> checked<%end if%>> <label for="doPDF">Show as PDF</label><br></div><%
end if

if officeIP and request("doEmail") <> "1" and request("doPDF") <> "1" then
	dim officeURL, officeURLQS
	officeURL = aspxPage(request.servervariables("SCRIPT_NAME"))
	officeURLQS = ""
	if request.querystring&"" <> "" then
		for each item in request.querystring
			if item = "officeIP" then
			else
				if officeURLQS <> "" then
					officeURLQS = officeURLQS & "&"
				end if
				officeURLQS = officeURLQS & item & "=" & request.querystring(item)
			end if
		next
	end if
	if officeURLQS <> "" then
		officeURL = officeURL & "?" & officeURLQS & "&"
	else
		officeURL = officeURL & "?"
	end if
	officeURL = officeURL & "officeIP=false"
	%><div class="debug"><input type="checkbox" name="doOfficeIPFalse" id="doOfficeIPFalse" onclick="document.location='<%=officeURL%>';"> <label for="doOfficeIPFalse">Show as Non-Office IP</label><br></div><%
end if

if isDev and request("doEmail") <> "1" and request("doPDF") <> "1" then
	dim devURL, devURLQS
	devURL = aspxPage(request.servervariables("SCRIPT_NAME"))
	devURLQS = ""
	if request.querystring&"" <> "" then
		for each item in request.querystring
			if item = "isDev" then
			else
				if devURLQS <> "" then
					devURLQS = devURLQS & "&"
				end if
				devURLQS = devURLQS & item & "=" & request.querystring(item)
			end if
		next
	end if
	if devURLQS <> "" then
		devURL = devURL & "?" & devURLQS & "&"
	else
		devURL = devURL & "?"
	end if
	devURL = devURL & "isDev=false"
	%><div class="debug"><input type="checkbox" name="isDevFalse" id="isDevFalse" onclick="document.location='<%=devURL%>';"> <label for="isDevFalse">Show as Live Site</label><br></div><%
end if

if request("doPDF") = "1" then
else
	if officeIP then
		showDebug
	end if
end if

if requiredMissing&"" <> "" then
	if isDev or officeIP then
		if strReqDebug&"" <> "" then
			%><div class="debug error alert alert-warning" id="divReqMissDebug"><%=strReqDebug%></div><br><%
		end if
	end if
	%><div class="error" id="divReqMiss"><%=reqMisMsg%></div><br><button type="button" id="btnSubmit" name="btnSubmit" class="btn btn-primary" onClick="document.location='<%=aspxPage(requiredMissing)%>';">Continue</button><%
end if

if request.form("do") = "goBack" then
	formLog("go back")
	prevPage = getPrevPage
	if prevPage&"" <> "" then
	    cForm.saveForm
		redirectForm(prevPage)
	end if
end if

if request.form("do") = "saveContinue" then
	cForm.saveForm
    if cRenew.PageOption("save_validate") = "Y" then
		numericComparisons
		dateComparisons
		dateDifferences
		combineFields(false)
		cRenew.LoadRequiredFields(cForm)
		cRenew.loadFieldTypes(cForm)
		requiredExplainFields
		cForm.Validateform
	end if

	if cForm.ErrorMsg = "" then
		formLog("saved to continue")
		on error goto 0
		cForm.saveForm
		on error resume next
		goNext
	else
		response.write "<div class=""error alert alert-warning"">" & cForm.ErrorMsg & "</div>"
	end if
end if

if request.form("do") = "saveForLater" then
	cForm.saveForm
	numericComparisons
	dateComparisons
	dateDifferences
	combineFields(false)
    if cRenew.PageOption("save_validate") = "Y" then
		cRenew.LoadRequiredFields(cForm)
		cRenew.loadFieldTypes(cForm)
		requiredExplainFields
		cForm.Validateform
	end if

	if cForm.ErrorMsg = "" then
		formLog("saved for later")
		on error goto 0
		cForm.saveForm
		saveFormData("SAVED")
    
		dim saveLoginPage
		saveLoginPage = "index.asp"
		if cRenew.PageCD(cRenew.ConfigOption("login_page")) <> "" then
			saveLoginPage = cRenew.PageCD(cRenew.ConfigOption("login_page"))
		end if
		%><div class="alert alert-success success">You have successfully saved your application. Please click <a href="<%=saveLoginPage%>">here</a> to login and access your saved application.</div><%
		For Each cookie in Response.Cookies
			Response.Cookies(cookie).Expires = DateAdd("d",-1,now())
		Next
	else
		response.write "<div class=""error alert alert-warning"">" & cForm.ErrorMsg & "</div>"
	end if
end if
%><script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script><script src="https://code.jquery.com/jquery-migrate-3.0.1.min.js"></script><%
if instr(ucase(request.servervariables("HTTP_USER_AGENT")),"FIREFOX") then
	%><script type="text/javascript" src="/www/jquery/jquery-ui-1.13.0/jquery-ui.min.js"></script><%
end if

%><script>var startTime = "<%=loadStart%>";</script><%

if isDev then
	%><script type='text/javascript' src='/www/js/forms.js?version=<%=version%>&tim=<%=dateDiff("s",cdate("01/01/1970"),now)%>'></script><%
else
	%><script type='text/javascript' src='/www/js/forms.js?version=<%=version%>'></script><%
end if

if timeoutCounter = "Y" then
	if cRenew.PageOption("end_session") = "Y" then
	else
		if request("doPDF")&"" = "1" or request("doEmail")&"" = "1" then
		else
			%><script>var timeoutInterval = <%=cint(sessionTimeout)*60%>;<%if timeoutMessage&"" <> "" then%>timeoutMessage = "<%=timeoutMessage%>";<%end if%><%if timeoutExit&"" <> "" then%>timeOutLocation = "<%=timeoutExit%>";<%end if%></script><%
			%><br><div id="divTimeout"><%=timeoutLabel%>&nbsp;<span id="timeremaining"></span></div><br><%
		end if
	end if
end if

passwordExpired

on error resume next
top
on error goto 0

'response.write "field = " & cRenew.getFieldName(396365)&"<BR>"
'response.write "desc = " & cRenew.getFieldDesc("eor_remarks")&"<BR>"

if loadDefaultsDelay = true then
	delayLog "client-side hide",true
else
	delayLog "client-side hide",false
	set cSqlForm = new cRunSql
	cSqlForm.Conn = connOracle
	cSqlForm.SqlStr = "select field_name, field_description, default_value, fields_hide, radio_not_equal from app_renew_fields where app_renew_type = ? and field_type = 'RADIO' and fields_hide is not null "
	cSqlForm.AddParam(appRenewCd)
	if cRenew.PageOption("onclick_page_specific") = "Y" then
		cSqlForm.AddSqlStr("and page = ?")
		cSqlForm.AddParam(pageCd)
	else
		cSqlForm.AddSqlStr("and nvl(page,'NONE') not in (select page from app_renew_pages where app_renew_type = ? and onclick_page_specific = 'Y')")
		cSqlForm.AddParam(appRenewCd)
	end if
	set rsFields = cSqlForm.Execute
	set cSqlForm = nothing
	do while not rsFields.eof
		doRadio = false
		if rsFields("RADIO_NOT_EQUAL") = "Y" then
			if session(rsFields("FIELD_NAME"))&"" <> yearFunction(rsFields("DEFAULT_VALUE")&"") then
				doRadio = true
			end if
		else
			if session(rsFields("FIELD_NAME"))&"" = yearFunction(rsFields("DEFAULT_VALUE")&"") then
				doRadio = true
			end if
		end if
		if doRadio then
			arrTmp = split(rsFields("FIELDS_HIDE"),", ")
			for i = lbound(arrTmp) to ubound(arrTmp)
				if replace(arrTmp(i)," ","") <> "" then
					if objFields.Exists(clng(arrTmp(i))) then
						cForm.SetAttribute objFields.Item(clng(arrTmp(i))),"style=""display:none"""
						msg = "hiding " & objFields.Item(clng(arrTmp(i))) & " because " & rsFields("FIELD_NAME") & " = " & yearFunction(rsFields("DEFAULT_VALUE"))
						cRenew.AddMsg(msg)
						formLog(msg)
					end if
				end if
			next
		end if
		rsFields.MoveNext
	loop
	rsFields.close
	set rsFields = nothing
end if

%><div id="divMessage" class="debug error alert alert-warning" style="display:none"></div><%

dim activeWork
activeWork = false
if activeWork then
	%><br><div align="center"><h2><b><font color="red">Active work being done on forms. You may see debug output, errors, etc. Please wait at least 30 minutes to see if an issue is resolved before submitting a task.</font></b></h2></div><%
end if

if request.form("do") = "saveForLater" and cForm.ErrorMsg = "" then
else
	if requiredMissing&"" <> "" then
	else
		if request("doPDF") <> "1" then
			if cRenew.PageOption("save_record") = "Y" then
				if session("MEMBERSHIP_ID")&"" = "" or session("MEMBERSHIP_ID") = "0" then
					set cNewMem = new cNewMemberO
					session("MEMBERSHIP_ID") = cNewMem.MembershipID
					membershipID = session("MEMBERSHIP_ID")
					cNewMem.SetValue "ORG_SUB_TY_CD",formRecType
					if cRenew.configOption("approve_app") = "N" then
						cNewMem.setValue "M_STAT_CD", "ACTIVE"
					else
						cNewMem.setValue "M_STAT_CD", "PENDING"
					end if
					cNewMem.InsertRecord
					set cNewMem = nothing
				end if
			end if
		end if
		if cognito then
	        If request.querystring("ExternalAuthId") <> "" Then

                session("COGNITO_ID") = ""

				redisKey = "ExternalAuth:" & request.querystring("ExternalAuthId")
				set redisSSO = Server.CreateObject("CacheManager")
				redisSSO.Init(redisKey)
				cognitoID = redisSSO("aws_access_username")
				set redisSSO = Nothing

				if cognitoID&"" <> "" then
					formLog("COGNITO_ID = " & cognitoID)
	                session("COGNITO_ID") = cognitoID
					cognitoFound = false
					set cSql = new cRunSql
					cSql.Conn = connOracle
					cSql.SqlStr = "select membership_id from membership where cognito_id = ?"
					cSql.AddParam(cognitoID)
					set rsCognito = cSql.Execute
					set cSql = nothing
					if rsCognito.eof Then
						formLog("no MEMBERSHIP record found")
					else
						formLog("found MEMBERSHIP record " & rsCognito("membership_id"))
						cognitoFound = true
						session("MEMBERSHIP_ID") = rsCognito("membership_id")
					end if
					rsCognito.close
					set rsCognito = nothing
					if cognitoFound then
						setCognitoCookie
						goNext
					else
						%>
						<br>
						<div align="center">
							Log in below to complete Single Sign-On
						</div>
						<%
					end if
				end if
			else
				cognitoProvidertype = getCodeValueDesc("APPLICATION_DB","SSO_COGNITO_TYPE_WWW")
				If Request.Cookies("cognitoSSOCookie") <> "" Then
					Dim strCognitoSSORedirect
					strCognitoSSORedirect = "/gdn/auth/index?AuthenticationOptionsToPresent=C&returnUrl=https://" & _
						Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("SCRIPT_NAME")
					if cognitoProvidertype&"" <> "" and lcase(cognitoProvidertype&"") <> "none" then
						strCognitoSSORedirect = strCognitoSSORedirect & "&SpecificProviderName=" & cognitoProvidertype
					end if
					redirectForm(strCognitoSSORedirect)
				End If
				%>
				<div align="center">
					<a class="btn btn-primary" href="javascript:sso('C','<%=cognitoProvidertype%>');"><i class="fa fa-sign-in"></i> Continue with BP Single Sign-On</a>
				</div>
				<%
			end if
		end if
		displayForm = true
		if cognito then
	        If request.querystring("ExternalAuthId") <> "" Then
			else
				displayForm = false
			end if
		end if
		if displayForm then
			%><!--#include virtual="/www/includes/display_form.inc"--><%
		end if
	end if
end if


if isArray(arrPages) and cRenew.ConfigOption("progress")&"" = "Y" then
	if request("doEmail") <> "1" then
		showProgress
	end if
end if

on error resume next
bottom
on error goto 0

if requiredMissing&"" <> "" then
else
	combineFields(true)
	%><script language="Javascript">function goPage(ddl) {	document.location = ddl.value + "?action=loadDebug&formID=<%=session("APP_RENEW_ID")%>"; }<%
	%>var btnSubmit = document.getElementById("btnSubmit");<%
	%>var btnSave = document.getElementById("btnSave");<%
	%>var btnExit = document.getElementById("btnExit");<%
	if cRenew.ConfigOption("OFFLINE_DISABLE_BUTTONS") = "Y" then
		%>checkOffline();var myVar = setInterval(checkOffline, 10000);<%
	end if

	if pageName = "unavailable.asp" then
	%>if (btnSave) hideElement(btnSave);if (btnSubmit) hideElement(btnSubmit);if (btnExit) hideElement(btnExit);<%
	end if

	if pageName = "index.asp" then
	%>if (btnSave)hideElement(btnSave);if (btnExit) hideElement(btnExit);<%
	end if

	if pageName = "pay.asp" then
	%>if (btnSave) hideElement(btnSave);<%
	end if

	if pageName = "success.asp" or pageName = "receipt.asp" then
	%>if (btnSave) hideElement(btnSave);if (btnSubmit) hideElement(btnSubmit);<%
	end if

	updateFields true

	msgClear = ""
	if loadDefaultsDelay = true then
		delayLog "client-side clear/hide/show/require/not require",true
	else
		delayLog "client-side clear/hide/show/require/not require",false
		set cSqlForm = new cRunSql
		cSqlForm.Conn = connOracle
		cSqlForm.SqlStr = "select parent.field_name, parent.field_description, child.default_value, child.fields_clear, child.fields_show, child.fields_require, child.fields_hide, child.radio_not_equal from app_renew_fields parent, app_renew_fields child " &_
			"where child.parent_id = parent.app_renew_fields_id " &_
			"and parent.app_renew_type = ? and parent.field_type = 'DROP' " &_
			"and (child.fields_clear is not null or child.fields_hide is not null or child.fields_show is not null or child.fields_require is not null)"
		cSqlForm.AddParam(appRenewCd)
		if cRenew.PageOption("onclick_page_specific") = "Y" then
			cSqlForm.AddSqlStr("and parent.page = ?")
			cSqlForm.AddParam(pageCd)
		else
			cSqlForm.AddSqlStr("and nvl(parent.page,'NONE') not in (select page from app_renew_pages where app_renew_type = ? and onclick_page_specific = 'Y')")
			cSqlForm.AddParam(appRenewCd)
		end if
		set rsFields = cSqlForm.Execute
		set cSqlForm = nothing
		do while not rsFields.eof
			doDrop = false
			if rsFields("RADIO_NOT_EQUAL") = "Y" then
				if session(rsFields("FIELD_NAME"))&"" <> yearFunction(rsFields("DEFAULT_VALUE")&"") then
					doDrop = true
				end if
			else
				if session(rsFields("FIELD_NAME"))&"" = yearFunction(rsFields("DEFAULT_VALUE")&"") then
					doDrop = true
				end if
			end if
			if doDrop then
				if rsFields("FIELDS_CLEAR")&"" <> "" then
					arrTmp = split(rsFields("FIELDS_CLEAR"),", ")
					for i = lbound(arrTmp) to ubound(arrTmp)
						if replace(arrTmp(i)," ","") <> "" then
							if objFields.Exists(clng(arrTmp(i))) then
								fieldName = objFields.Item(clng(arrTmp(i)))
								msg = "clearing " & fieldName & " because " & rsFields("FIELD_NAME") & " = " & session(rsFields("FIELD_NAME"))
								if isDev then
									msgClear = msgClear & msg & "\n"
								end if
								cRenew.AddMsg(msg)
								formLog(msg)
								%>clearElementID("<%=fieldName%>");<%
								session(fieldName) = ""
								session(fieldName) = ""
								session("arr_files_" & fieldName) = ""
							end if
						end if
					next
				end if
				if rsFields("FIELDS_HIDE")&"" <> "" then
					arrTmp = split(rsFields("FIELDS_HIDE"),", ")
					for i = lbound(arrTmp) to ubound(arrTmp)
						if replace(arrTmp(i)," ","") <> "" then
							if objFields.Exists(clng(arrTmp(i))) then
								fieldName = objFields.Item(clng(arrTmp(i)))
								msg = "hiding " & fieldName & " because " & rsFields("FIELD_NAME") & " = " & session(rsFields("FIELD_NAME"))
								cRenew.AddMsg(msg)
								formLog(msg)
							%>hideElementID("btnAdd<%=fieldName%>");hideElementID("lbl<%=fieldName%>");hideElementID("<%=fieldName%>");<%
							end if
						end if
					next
				end if
				if rsFields("FIELDS_SHOW")&"" <> "" then
					arrTmp = split(rsFields("FIELDS_SHOW"),", ")
					for i = lbound(arrTmp) to ubound(arrTmp)
						if replace(arrTmp(i)," ","") <> "" then
							if objFields.Exists(clng(arrTmp(i))) then
								fieldName = objFields.Item(clng(arrTmp(i)))
								msg = "showing " & fieldName & " because " & rsFields("FIELD_NAME") & " = " & session(rsFields("FIELD_NAME"))
								cRenew.AddMsg(msg)
								formLog(msg)
							%>showElementID("spn<%=fieldName%>");showElementID("lbl<%=fieldName%>");showElementID("btnAdd<%=fieldName%>");showElementID("<%=fieldName%>");<%
							end if
						end if
					next
				end if
				if rsFields("FIELDS_REQUIRE")&"" <> "" then
					arrTmp = split(rsFields("FIELDS_REQUIRE"),", ")
					for i = lbound(arrTmp) to ubound(arrTmp)
						if replace(arrTmp(i)," ","") <> "" then
							if objFields.Exists(clng(arrTmp(i))) then
								fieldName = objFields.Item(clng(arrTmp(i)))
								msg = "requiring " & fieldName & " because " & rsFields("FIELD_NAME") & " = " & session(rsFields("FIELD_NAME"))
								cRenew.AddMsg(msg)
								formLog(msg)
							%>requireElementID("<%=fieldName%>");<%
							end if
						end if
					next
				end if
			end if
			rsFields.MoveNext
		loop
		rsFields.close
		set rsFields = nothing
		set cSqlForm = nothing

		set cSqlForm = new cRunSql
		cSqlForm.Conn = connOracle
		cSqlForm.SqlStr = "select field_name, field_description, default_value, fields_clear, fields_hide, fields_show, fields_require, radio_not_equal, fields_unrequire from app_renew_fields where app_renew_type = ? and field_type = 'RADIO' " &_
			"and (fields_clear is not null or fields_hide is not null or fields_show is not null or fields_require is not null or fields_unrequire is not null) "
		cSqlForm.AddParam(appRenewCd)
		if cRenew.PageOption("onclick_page_specific") = "Y" then
			cSqlForm.AddSqlStr("and page = ?")
			cSqlForm.AddParam(pageCd)
		else
			cSqlForm.AddSqlStr("and nvl(page,'NONE') not in (select page from app_renew_pages where app_renew_type = ? and onclick_page_specific = 'Y')")
			cSqlForm.AddParam(appRenewCd)
		end if
		set rsFields = cSqlForm.Execute
		set cSqlForm = nothing
		do while not rsFields.eof
			doRadio = false
			if rsFields("RADIO_NOT_EQUAL") = "Y" then
				if session(rsFields("FIELD_NAME"))&"" <> yearFunction(rsFields("DEFAULT_VALUE")&"") then
					doRadio = true
				end if
			else
				if session(rsFields("FIELD_NAME"))&"" = yearFunction(rsFields("DEFAULT_VALUE")&"") then
					doRadio = true
				end if
			end if
			if doRadio then
				if rsFields("FIELDS_CLEAR")&"" <> "" then
					arrTmp = split(rsFields("FIELDS_CLEAR"),", ")
					for i = lbound(arrTmp) to ubound(arrTmp)
						if replace(arrTmp(i)," ","") <> "" then
							if objFields.Exists(clng(arrTmp(i))) then
								fieldName = objFields.Item(clng(arrTmp(i)))
								msg = "clearing " & fieldName & " because " & rsFields("FIELD_NAME") & " = " & session(rsFields("FIELD_NAME"))
								if isDev then
									msgClear = msgClear & msg & "\n"
								end if
								cRenew.AddMsg(msg)
								formLog(msg)
								%>clearElementID("<%=fieldName%>");<%
								session(fieldName) = ""
								session("arr_files_" & fieldName) = ""
							end if
						end if
					next
				end if
				if rsFields("FIELDS_HIDE")&"" <> "" then
					arrTmp = split(rsFields("FIELDS_HIDE"),", ")
					for i = lbound(arrTmp) to ubound(arrTmp)
						if replace(arrTmp(i)," ","") <> "" then
							if objFields.Exists(clng(arrTmp(i))) then
								fieldName = objFields.Item(clng(arrTmp(i)))
								msg = "hiding " & fieldName & " because " & rsFields("FIELD_NAME") & " = " & session(rsFields("FIELD_NAME"))
								cRenew.AddMsg(msg)
								formLog(msg)
							%>hideElementID("btnAdd<%=fieldName%>");hideElementID("lbl<%=fieldName%>");hideElementID("<%=fieldName%>");<%
							end if
						end if
					next
				end if
				if rsFields("FIELDS_SHOW")&"" <> "" then
					arrTmp = split(rsFields("FIELDS_SHOW"),", ")
					for i = lbound(arrTmp) to ubound(arrTmp)
						if replace(arrTmp(i)," ","") <> "" then
							if objFields.Exists(clng(arrTmp(i))) then
								fieldName = objFields.Item(clng(arrTmp(i)))
								msg = "showing " & fieldName & " because " & rsFields("FIELD_NAME") & " = " & session(rsFields("FIELD_NAME"))
								cRenew.AddMsg(msg)
								formLog(msg)
							%>showElementID("spn<%=fieldName%>");showElementID("lbl<%=fieldName%>");showElementID("btnAdd<%=fieldName%>");showElementID("<%=fieldName%>");<%
							end if
						end if
					next
				end if
				if rsFields("FIELDS_REQUIRE")&"" <> "" then
					arrTmp = split(rsFields("FIELDS_REQUIRE"),", ")
					for i = lbound(arrTmp) to ubound(arrTmp)
						if replace(arrTmp(i)," ","") <> "" then
							if objFields.Exists(clng(arrTmp(i))) then
								fieldName = objFields.Item(clng(arrTmp(i)))
								msg = "requiring " & fieldName & " because " & rsFields("FIELD_NAME") & " = " & session(rsFields("FIELD_NAME"))
								cRenew.AddMsg(msg)
								formLog(msg)
							%>requireElementID("<%=fieldName%>");<%
							end if
						end if
					next
				end if
				if rsFields("FIELDS_UNREQUIRE")&"" <> "" then
					arrTmp = split(rsFields("FIELDS_UNREQUIRE"),", ")
					for i = lbound(arrTmp) to ubound(arrTmp)
						if replace(arrTmp(i)," ","") <> "" then
							if objFields.Exists(clng(arrTmp(i))) then
								fieldName = objFields.Item(clng(arrTmp(i)))
								msg = "unrequiring " & fieldName & " because " & rsFields("FIELD_NAME") & " = " & session(rsFields("FIELD_NAME"))
								cRenew.AddMsg(msg)
								formLog(msg)
							%>unrequireElementID("<%=fieldName%>");<%
							end if
						end if
					next
				end if
			end if
			rsFields.MoveNext
		loop
		rsFields.close
		set rsFields = nothing
		set cSqlForm = nothing

		set cSqlForm = new cRunSql
		cSqlForm.Conn = connOracle
		cSqlForm.SqlStr = "select field_name, field_description, fields_show_check, fields_require_check, fields_hide_check, fields_show_uncheck, fields_require_uncheck, fields_hide_uncheck " &_
			"from app_renew_fields where app_renew_type = ? and field_type = 'CHECK' " &_
			"and (fields_show_check is not null or fields_require_check is not null or fields_hide_check is not null or fields_show_uncheck is not null or fields_require_uncheck is not null or fields_hide_uncheck is not null)"
		cSqlForm.AddParam(appRenewCd)
		if cRenew.PageOption("onclick_page_specific") = "Y" then
			cSqlForm.AddSqlStr("and page = ?")
			cSqlForm.AddParam(pageCd)
		else
			cSqlForm.AddSqlStr("and nvl(page,'NONE') not in (select page from app_renew_pages where app_renew_type = ? and onclick_page_specific = 'Y')")
			cSqlForm.AddParam(appRenewCd)
		end if
		set rsFields = cSqlForm.Execute
		set cSqlForm = nothing
		do while not rsFields.eof
			if session(rsFields("FIELD_NAME"))&"" = "1" then
				if rsFields("FIELDS_SHOW_CHECK")&"" <> "" then
					arrTmp = split(rsFields("FIELDS_SHOW_CHECK"),", ")
					for i = lbound(arrTmp) to ubound(arrTmp)
						if replace(arrTmp(i)," ","") <> "" then
							if objFields.Exists(clng(arrTmp(i))) then
								fieldName = objFields.Item(clng(arrTmp(i)))
								msg = "showing " & fieldName & " because " & rsFields("FIELD_NAME") & " = 1 (checked)"
								cRenew.AddMsg(msg)
								formLog(msg)
							%>showElementID("spn<%=fieldName%>");showElementID("lbl<%=fieldName%>");showElementID("btnAdd<%=fieldName%>");showElementID("<%=fieldName%>");<%
							end if
						end if
					next
				end if
				if rsFields("FIELDS_REQUIRE_CHECK")&"" <> "" then
					arrTmp = split(rsFields("FIELDS_REQUIRE_CHECK"),", ")
					for i = lbound(arrTmp) to ubound(arrTmp)
						if replace(arrTmp(i)," ","") <> "" then
							if objFields.Exists(clng(arrTmp(i))) then
								fieldName = objFields.Item(clng(arrTmp(i)))
								msg = "requiring " & fieldName & " because " & rsFields("FIELD_NAME") & " = 1 (checked)"
								cRenew.AddMsg(msg)
								formLog(msg)
							%>requireElementID("<%=fieldName%>");<%
							end if
						end if
					next
				end if
				if rsFields("FIELDS_HIDE_CHECK")&"" <> "" then
					arrTmp = split(rsFields("FIELDS_HIDE_CHECK"),", ")
					for i = lbound(arrTmp) to ubound(arrTmp)
						if replace(arrTmp(i)," ","") <> "" then
							if objFields.Exists(clng(arrTmp(i))) then
								fieldName = objFields.Item(clng(arrTmp(i)))
								msg = "hiding " & fieldName & " because " & rsFields("FIELD_NAME") & " = 1 (checked)"
								cRenew.AddMsg(msg)
								formLog(msg)
							%>hideElementID("btnAdd<%=fieldName%>");hideElementID("lbl<%=fieldName%>");hideElementID("<%=fieldName%>");<%
							end if
						end if
					next
				end if
			end if
			if session(rsFields("FIELD_NAME"))&"" = "" then
				if rsFields("FIELDS_SHOW_UNCHECK")&"" <> "" then
					arrTmp = split(rsFields("FIELDS_SHOW_UNCHECK"),", ")
					for i = lbound(arrTmp) to ubound(arrTmp)
						if replace(arrTmp(i)," ","") <> "" then
							if objFields.Exists(clng(arrTmp(i))) then
								fieldName = objFields.Item(clng(arrTmp(i)))
								msg = "showing " & fieldName & " because " & rsFields("FIELD_NAME") & " is empty (unchecked)"
								cRenew.AddMsg(msg)
								formLog(msg)
							%>showElementID("spn<%=fieldName%>");showElementID("lbl<%=fieldName%>");showElementID("btnAdd<%=fieldName%>");showElementID("<%=fieldName%>");<%
							end if
						end if
					next
				end if
				if rsFields("FIELDS_REQUIRE_UNCHECK")&"" <> "" then
					arrTmp = split(rsFields("FIELDS_REQUIRE_UNCHECK"),", ")
					for i = lbound(arrTmp) to ubound(arrTmp)
						if replace(arrTmp(i)," ","") <> "" then
							if objFields.Exists(clng(arrTmp(i))) then
								fieldName = objFields.Item(clng(arrTmp(i)))
								msg = "requiring " & fieldName & " because " & rsFields("FIELD_NAME") & " is empty (unchecked)"
								cRenew.AddMsg(msg)
								formLog(msg)
							%>requireElementID("<%=fieldName%>");<%
							end if
						end if
					next
				end if
				if rsFields("FIELDS_HIDE_UNCHECK")&"" <> "" then
					arrTmp = split(rsFields("FIELDS_HIDE_UNCHECK"),", ")
					for i = lbound(arrTmp) to ubound(arrTmp)
						if replace(arrTmp(i)," ","") <> "" then
							if objFields.Exists(clng(arrTmp(i))) then
								fieldName = objFields.Item(clng(arrTmp(i)))
								msg = "hiding " & fieldName & " because " & rsFields("FIELD_NAME") & " is empty (unchecked)"
								cRenew.AddMsg(msg)
								formLog(msg)
							%>hideElementID("btnAdd<%=fieldName%>");hideElementID("lbl<%=fieldName%>");hideElementID("<%=fieldName%>");<%
							end if
						end if
					next
				end if
			end if
			rsFields.MoveNext
		loop
		rsFields.close
		set rsFields = nothing
		set cSqlForm = nothing
	end if

	if isDev then
		if msgClear <> "" Then
			%>alert("<%=msgClear%>");<%
		end if
	end if

	numericComparisons
	dateComparisons
	dateDifferences
	
	%>var bottomFunctionsComplete = false;bottomFunctions();bottomFunctionsComplete = true;<%

	'correct date formats and dates and numbers in field groups
	arrFieldTypes = cForm.fieldTypesArr
	if isArray(arrFieldTypes) then
		for idxFT = lbound(arrFieldTypes,2) to ubound(arrFieldTypes,2)
			fieldName = arrFieldTypes(0,idxFT)
			fieldType = arrFieldTypes(2,idxFT)
			dataType = arrFieldTypes(3,idxFT)
			codeClass = arrFieldTypes(4,idxFT)
			if dataType = "DATE" then
				fieldData = session(fieldName)
				if fieldData&"" <> "" then
					if instr(fieldData,"\\") then
						arrFieldData = split(fieldData&"","\\")
						for idxFieldData = lbound(arrFieldData) to ubound(arrFieldData)
							fieldData = arrFieldData(idxFieldData)
							if isDate(fieldData) then
								tmpMonth = month(cdate(fieldData))
								tmpDay = day(cdate(fieldData))
								tmpYear = year(cdate(fieldData))
								if tmpMonth < 10 then
									tmpMonth = "0" & tmpMonth
								end if
								if tmpDay < 10 then
									tmpDay = "0" & tmpDay
								end if
								tmpDate = tmpYear & "-" & tmpMonth & "-" & tmpDay
								%>if (document.getElementsByName("<%=fieldName%>")[<%=idxFieldData%>]) document.getElementsByName("<%=fieldName%>")[<%=idxFieldData%>].value = "<%=tmpDate%>";<%
							end if
						next
					end if
				end if
			end if
			if dataType = "NUMBER" or dataType = "MONTH" or (fieldType = "DROP" and request("doPDF")&"" <> "1") or dataType = "YEARS" then
				fieldData = session(fieldName)
				if fieldData&"" <> "" then
					if instr(fieldData,"\\") then
						arrFieldData = split(fieldData&"","\\")
						for idxFieldData = lbound(arrFieldData) to ubound(arrFieldData)
							fieldData = arrFieldData(idxFieldData)
							%>if (document.getElementsByName("<%=fieldName%>")[<%=idxFieldData%>]) document.getElementsByName("<%=fieldName%>")[<%=idxFieldData%>].value = "<%=fieldData%>";<%
						next
					end if
				end if
			end if
		next
	end if
	
	%>copyTextAreas();</script><%

end if

%></div><%
if request("doPDF") <> "1" and request("doEmail") <> "1" then
	' === FORM SUBMIT LEDGER PATCH ===
	cLedger.SetContext session("APP_RENEW_ID"), appRenewCd, session("MEMBERSHIP_ID")
	' === END FORM SUBMIT LEDGER PATCH ===
	if cRenew.PageOption("save_record") = "Y" then
		' === FORM SUBMIT LEDGER PATCH ===
		cLedger.SetOperationState "records", "running", "Updating records (page save_record)"
		' === END FORM SUBMIT LEDGER PATCH ===
		cRenew.SaveData
		' === FORM SUBMIT LEDGER PATCH ===
		' Sets records = awaiting when an Awaiting Updates hold exists, else done.
		cLedger.MarkRecordsAfterSave
		' === END FORM SUBMIT LEDGER PATCH ===
	end if
	if cRenew.PageOption("save_pending") = "Y" then
		saveFormData("PENDING")
	end if
	if cRenew.PageOption("email_field") <> "" then
		' === FORM SUBMIT LEDGER PATCH ===
		cLedger.SetOperationState "emails", "running", "Sending configured page email"
		' === END FORM SUBMIT LEDGER PATCH ===
		if session("DO_AWAIT") = true then
			if cRenew.ConfigOption("email_await") = "Y" then
				cRenew.EmailPage(isFormScanner)
			end if
		else
			cRenew.EmailPage(isFormScanner)
		end if
		' === FORM SUBMIT LEDGER PATCH ===
		cLedger.SetOperationState "emails", "done", "Configured page email sent"
		' === END FORM SUBMIT LEDGER PATCH ===
	end if
	if cRenew.PageOption("create_pdf") = "Y" then
		' === FORM SUBMIT LEDGER PATCH ===
		cLedger.SetOperationState "pdf", "running", "Generating PDF (page create_pdf)"
		' === END FORM SUBMIT LEDGER PATCH ===
		cRenew.PDF
		' === FORM SUBMIT LEDGER PATCH ===
		cLedger.SetOperationState "pdf", "done", "PDF generated (page create_pdf)"
		' === END FORM SUBMIT LEDGER PATCH ===
	end if
	if cRenew.PageOption("record_zero_trans") = "Y" then
		if session("CARD_AMOUNT")&"" <> "" then
			if clng(session("CARD_AMOUNT")) = 0 then
				' === FORM SUBMIT LEDGER PATCH ===
				cLedger.SetOperationState "billing", "running", "Recording zero-dollar batch transaction"
				' === END FORM SUBMIT LEDGER PATCH ===
				batchTransaction
				' === FORM SUBMIT LEDGER PATCH ===
				cLedger.SetOperationState "billing", "done", "Batch transaction recorded"
				' === END FORM SUBMIT LEDGER PATCH ===
			end if
		end if
	end if
	if cRenew.PageOption("record_batch_trans") = "Y" then
		if session("CARD_AMOUNT")&"" <> "" then
			' === FORM SUBMIT LEDGER PATCH ===
			cLedger.SetOperationState "billing", "running", "Recording batch transaction"
			' === END FORM SUBMIT LEDGER PATCH ===
			batchTransaction
			' === FORM SUBMIT LEDGER PATCH ===
			cLedger.SetOperationState "billing", "done", "Batch transaction recorded"
			' === END FORM SUBMIT LEDGER PATCH ===
		end if
	end if
	' === FORM SUBMIT LEDGER PATCH ===
	' Reconcile durable submission state after this page's operations ran.
	cLedger.RecomputeSubmitState
	' === END FORM SUBMIT LEDGER PATCH ===
	nextPage = getNextPage
	if nextPage&"" = "" or cRenew.PageOption("email_pages_onclick") = "Y" then
		if loadDefaultsDelay = true then
			delayLog "onclick page to email",true
		else
			delayLog "onclick page to email",false
			cRenew.EmailPagesOnclick(isFormScanner)
		end if
	end if
	if cRenew.PageOption("end_session") = "Y" then
		if request.querystring("action") = "loadDebug" then		
		else
			formLog("ending session because page option End Session set to Y")
			cRenew.clearFormSession()
			'session.abandon
		end if
	end if
end if

else
	showDebug
	%><div><%=cRenew.ConfigOption("unavail_msg")%></div><%
end if

if request.querystring("doPDF") = "1" then
else
	if officeIP then
		cRenew.ShowMsg()
	end if
end if
if request.form("do") = "saveForLater" then
	if cForm.ErrorMsg = "" then
		formLog("ending session because user clicked Save for Later")
		cRenew.clearFormSession()
	end if
end if

if cRenew.ConfigOption("log_session_page") = "Y" then
	logSession("bottom")
end if

loadEnd = now
loadTime = dateDiff("s",loadStart,loadEnd)
session("LOAD_TIME_" & pageCd) = loadtime

if isDev or officeIP then
	if request.querystring("doPDF")&"" = "1" or request.querystring("doEmail")&"" = "1" then
	else
		response.write "<br>started loading at " & loadStart & "<br>"
		response.write "finished loading at " & loadEnd & "<br>"
		response.write "load time = " & loadTime & " seconds<br>"
	end if
end if

formLog("end loading page")

if timeoutCounter = "Y" then
	%><script>addLoadEvent(startTimer);</script><%
end if

if pageID&"" = cRenew.ConfigOption("load_defaults_page")&"" then
	if cRenew.PageOption("end_session") = "Y" then
	else
		saveFormData("PROGRESS")
	end if
end if

if cint(cRenew.ConfigOption("auto_save_interval")) > 0 then
	%><script>autoSave(<%=cRenew.ConfigOption("auto_save_interval")%>);</script><%
end if

if isDev then
	%><script type='text/javascript' src='/www/js/deps.js?version=<%=version%>&tim=<%=dateDiff("s",cdate("01/01/1970"),now)%>'></script><%
else
	%><script type='text/javascript' src='/www/js/deps.js?version=<%=version%>'></script><%
end if

%><!--#include virtual="/footer.asp"--><%

if request.querystring("doPDF") = "1" and request.querystring("sessionID") <> "" then
	if redis then
	else
		session.abandon
	end if
end if

closeAll()

sub paymentStatus(stat)
	session("PAYMENT_STATUS") = stat
	session("PAYMENT_DATE_TIME") = sysNow

	set cSql = new cRunSql
	cSql.Conn = connOracle
	cSql.SqlStr = "delete from app_renew_detail where app_renew_id = ? and lower(field_name) in ('payment_status','payment_date_time')"
	cSql.AddParam(session("APP_RENEW_ID"))
	cSql.Execute
	set cSql = nothing

	set cSql = new cRunSql
	cSql.Conn = connOracle
	cSql.SqlStr = "insert into app_renew_detail (app_renew_id,field_name,data) values (?,?,?)"
	cSql.AddParam(session("APP_RENEW_ID"))
	cSql.AddParam("PAYMENT_STATUS")
	cSql.AddParam(stat)
	cSql.Execute
	set cSql = nothing

	set cSql = new cRunSql
	cSql.Conn = connOracle
	cSql.SqlStr = "insert into app_renew_detail (app_renew_id,field_name,data) values (?,?,?)"
	cSql.AddParam(session("APP_RENEW_ID"))
	cSql.AddParam("PAYMENT_DATE_TIME")
	cSql.AddParam(session("PAYMENT_DATE_TIME"))
	cSql.Execute
	set cSql = nothing

	' === FORM SUBMIT LEDGER PATCH ===
	' Map the legacy payment status to durable payment + submission state.
	cLedger.SetContext session("APP_RENEW_ID"), appRenewCd, session("MEMBERSHIP_ID")
	cLedger.RecordPaymentStatus stat
	' === END FORM SUBMIT LEDGER PATCH ===

end sub

sub setCognitoCookie
	Response.Cookies("cognitoSSOCookie") = True
	Response.Cookies("cognitoSSOCookie").Expires = DateAdd("d", 400, Now())
	Response.Cookies("cognitoSSOCookie").Path = "/"
end sub

sub existingForm(formCode,formStatus,existField,newField,onlyNumbers)
	dim rsEA
	session("EXISTING_FORM") = "0"
	set cSql = new cRunSql
	cSql.Conn = connOracle
	cSql.SqlStr = "select a.app_renew_id from app_renew a, app_renew_detail d " &_
		"where d.app_renew_id = a.app_renew_id " &_
		"and a.app_renew_status = ? and a.APP_RENEW_TYPE = ? " &_
		"and lower(d.field_name) = ? "
	if onlyNumbers then
		cSql.AddSqlStr("and REGEXP_REPLACE(d.data, '[^0-9]+', '') = ?")
	else
		cSql.AddSqlStr("and lower(d.data) = ?")
	end if
	cSql.AddParam(formStatus)
	cSql.AddParam(formCode)
	cSql.AddParam(lcase(existField))
	if onlyNumbers then
		cSql.AddParam(numericOnly(session(newField)))
	else
		cSql.AddParam(lcase(session(newField)))
	end if
	set rsEA = cSql.execute
	set cSql = nothing
	if rsEA.eof then
	else
		session("EXISTING_FORM") = "1"
	end if
	rsEA.close
	set rsEA = nothing
end sub

sub existingFormByIDType(formCode,formStatus,recordIDField,setFieldName,setFieldValueFound,setFieldValueNotFound)
	dim rsEA, arrStat, arrCode, idxStat, idxCode, formID

	formLog("start existingFormByIDType")
	formLog(formCode & "," & formStatus & "," & recordIDField & "," & setFieldName & "," & setFieldValueFound & "," & setFieldValueNotFound)
	formID = session(recordIDField)
	arrStat = split(formStatus,",")
	arrCode = split(formCode,",")
	
	session(setFieldName) = setFieldValueNotFound
	set cSql = new cRunSql
	cSql.Conn = connOracle
	cSql.SqlStr = "select app_renew_id from app_renew " &_
		"where membership_id = ? and app_renew_status in ("
	for idxStat = lbound(arrStat) to ubound(arrStat)
		cSql.AddSqlStr("?")
		if idxStat < ubound(arrStat) then
			cSql.AddSqlStr(",")
		end if
	next
	cSql.AddSqlStr(") and app_renew_type in (")
	for idxCode = lbound(arrCode) to ubound(arrCode)
		cSql.AddSqlStr("?")
		if idxCode < ubound(arrCode) then
			cSql.AddSqlStr(",")
		end if
	next
	cSql.AddSqlStr(")")
	cSql.AddParam(formID)
	for idxStat = lbound(arrStat) to ubound(arrStat)
		cSql.AddParam(arrStat(idxStat))
	next
	for idxCode = lbound(arrCode) to ubound(arrCode)
		cSql.AddParam(arrCode(idxCode))
	next
	set rsEA = cSql.execute
	set cSql = nothing
	if rsEA.eof then
		formLog("no existing forms found")
	else
		formLog("existing forms found")
		session(setFieldName) = setFieldValueFound
	end if
	rsEA.close
	set rsEA = nothing
	formLog("end existingFormByIDType")
end sub

sub logSession(pagePos)
	if request.querystring("doPDF") = "1" or request.querystring("doEmail") = "1" Then
		exit sub
	end if
	formLog("start all session values " & pagePos)
	if redis then
		arrKeys = split(session.GetCacheList,",")
		if isArray(arrKeys) then
			url = "https://forms.ebigpicture.com/forms.ashx?action=logSession&key=" & session("redisSessionID") & "&script_name=" & request.servervariables("SCRIPT_NAME") & "&pos=" & pagePos
			if pagePos = "top" then
				ajaxURL url,"GET",false
			else
				ajaxURL url,"GET",true
			end if
		end if
	else
		for each item in session.contents
			if isArray(session(item)) then
				formLog(item & " is array")
			else
				formLog(item & " = " & session(item))
			end if
		next
	end if
	formLog("end all session values " & pagePos)
end sub

sub setPayLaterCardAmount
	dim payTotal, arrPay
	payTotal = 0
	payChecked = false
	for i = 1 to request.form("chkPay").count
		arrPay = split(request.form("chkPay")(i),"|")
		payTotal = payTotal + clng(arrPay(2))
	next
	session("CARD_AMOUNT") = payTotal
end sub

sub updateAllChildRecords(recordType,fieldToUpdate,fieldWithValue)
	if request("doPDF")&"" = "1" or request("doEmail")&"" = "1" or request("action")&"" = "loadDebug" then
		exit Sub
	end if

	dim rsUC, cMemUC

	set cSql = new cRunSql
	cSql.Conn = connOracle
	cSql.SqlStr = "select m.membership_id from membership m, mor_finder f " &_
		"where f.membership_id = m.membership_id and f.org_relate_id = ? " &_
		"and m.org_sub_ty_cd = ?"
	cSql.AddParam(session("MEMBERSHIP_ID"))
	cSql.AddParam(recordType)
	set rsUC = cSql.execute
	set cSql = nothing
	if rsUC.eof then
	else
		do while not rsUC.eof
			set cMemUC = new cMemberValuesO
			cMemUC.UpdateValue rsUC("MEMBERSHIP_ID"),fieldToUpdate,session(fieldWithValue)
			set cMemUC = nothing
			rsUC.MoveNext
		loop
	end if
	rsUC.close
	set rsUC = nothing
end sub

sub addYears(fieldName,numYrs)
	dim idxYr
%>
<script>
var optionText;
var optionValue;
<%for idxYr = year(date)+1 to year(date)+numYrs%>
optionText = "<%=idxYr%>";
optionValue = "<%=idxYr%>";
$('#<%=fieldName%>').append(`<option value="${optionValue}">${optionText}</option>`);
<%next%>
</script>
<%
end sub

sub validateCCFields

	dim cardNum, cardMo, cardYr, cardCVV

	cardNum = request.form("CARD_NUM")
	cardMo = request.form("CARD_EXP_MO")
	cardYr = request.form("CARD_EXP_YR")
	cardCVV = request.form("CARD_CVV")

	if cardNum&"" <> "" then
		for idxField = lbound(arrFields,2) to ubound(arrFields,2)
			if arrFields(fieldNameIndex("PAGE"),idxField) = pageCd and arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "CCNOAMEX" then
				if left(cardNum,2) = "34" or left(cardNum,2) = "37" then
					cForm.AddErrorMsg("American Express is NOT accepted")
				end if
				idxField = ubound(arrFields,2) + 1
			end if
		next

		if left(cardNum,1) = "2" or left(cardNum,1) = "3" or left(cardNum,1) = "4" or left(cardNum,1) = "5" or left(cardNum,1) = "6" then
			if left(cardNum,1) = "3" then
				if len(cardNum) <> 15 then
					cForm.AddErrorMsg "Invalid card number"
				end if
			elseif len(cardNum) <> 16 then
				cForm.AddErrorMsg "Invalid card number"
			elseif not IsValidLuhn(cardNum) then
				cForm.AddErrorMsg "Invalid card number"
			end if
			if left(cardNum,1) = "3" Then
				if len(cardCVV) <> 4 then
					cForm.AddErrorMsg "Invalid CVV"
				end if
			else
				if len(cardCVV) <> 3 then
					cForm.AddErrorMsg "Invalid CVV"
				end if
			end if
		else
			cForm.AddErrorMsg "Invalid card number"
		end if
		if cint(cardYr) = year(date) then
			if cint(cardMo) < month(date) then
				cForm.AddErrorMsg "Invalid expiration date"
			end if
		end if
	end if
end sub

Function IsValidLuhn(cardNumber)
    Dim sum, alt, i, digit

    sum = 0
    alt = False
    
    ' Loop from right to left
    For i = Len(cardNumber) To 1 Step -1
        digit = Mid(cardNumber, i, 1)
        
        ' Skip non-digits (just in case)
        If IsNumeric(digit) Then
            digit = CInt(digit)
            
            If alt Then
                digit = digit * 2
                If digit > 9 Then
                    digit = digit - 9
                End If
            End If
            
            sum = sum + digit
            alt = Not alt
        End If
    Next
    
    ' Valid if sum is multiple of 10
    IsValidLuhn = (sum Mod 10 = 0)
End Function

sub passwordExpired
	dim rsPassExp, passLastChanged, passChange

	if session("MEMBERSHIP_ID")&"" <> "0" then
		if cRenew.ConfigOption("forgot_password_password_field")&"" <> "" then
			if cRenew.ConfigOption("change_password_date")&"" <> "" or cRenew.ConfigOption("change_password_days")&"" <> "" then
				passChange = false

				set cSqlForm = new cRunSql
				cSqlForm.Conn = connOracle
				cSqlForm.SqlStr = "select to_char(modify_dt,'mm/dd/yyyy') modify_dt from member_detail where membership_id = ? and column_id = (select column_id from org_column_detail where column_desc = ?)"			
				cSqlForm.AddParam(session("MEMBERSHIP_ID"))
				cSqlForm.AddParam(cRenew.ConfigOption("forgot_password_password_field"))
				set rsPassExp = cSqlForm.Execute
				set cSqlForm = nothing
				if rsPassExp.eof then
					passChange = true
				else
					passLastChanged = cdate(rsPassExp("modify_dt"))
					if cRenew.ConfigOption("change_password_days")&"" <> "" then
						if dateDiff("d",passLastChanged,date) > cint(cRenew.ConfigOption("change_password_days")) then
							passChange = true
						end if
					elseif cRenew.ConfigOption("change_password_date")&"" <> "" then
						if passLastChanged < cdate(cRenew.ConfigOption("change_password_date")) then
							passChange = true
						end if
					end if
				end if
				rsPassExp.close
				set rsPassExp = nothing
				if passChange then
					session("CHANGE_PASSWORD") = true
					session("CHANGE_PASSWORD_ID") = session("MEMBERSHIP_ID")
					session("FORGOT_PASSWORD_CODE") = "1"

					set cCV = new cCodeValueO
					cCV.LoadFromCodeClassAndValue "APP_RENEW_TYPE_CD", session("APP_RENEW_CD")
					formPath = cCV.CodeValueLongDesc
					set cCV = nothing

					session("FORGOT_PASSWORD_LOGIN_PAGE") = formPath & "/" & cRenew.PageCD(cRenew.ConfigOption("login_page"))
					redirect("/www/forms/changepassword/?expired=true")
					response.end
				end if
			end if
		end if
	end if
end sub

sub verifyAddress(verDesc,verTyp,verFields)
	dim arrAdd, addLine1Field, addCityField, addStateField, addZipField, addZip4Field
	dim addLine1Value, addCityValue, addStateValue, addZipValue
	dim addJSON, addXML, addResp, addVerURL, addMsg

	addVerURL = "https://us-street.api.smartystreets.com/street-address?" &_
		"auth-id=55bea508-3623-50b8-fb02-2f033ccd8087&" &_
		"auth-token=fGqt971z8BaIXVlplXUW&" &_
		"candidates=1&" &_
		"match=strict&license=us-core-cloud"

	arrAdd = split(verFields,", ")
	if arrAdd(0)&"" <> "" then
		addLine1Field = cRenew.getFieldName(arrAdd(0))
		addLine1Value = session(addLine1Field)&""
	end if
	if arrAdd(1)&"" <> "" then
		addCityField = cRenew.getFieldName(arrAdd(1))
		addCityValue = session(addCityField)&""
	end if
	if arrAdd(2)&"" <> "" then
		addStateField = cRenew.getFieldName(arrAdd(2))
		addStateValue = session(addStateField)&""
	end if
	if arrAdd(3)&"" <> "" then
		addZipField = cRenew.getFieldName(arrAdd(3))
		addZipValue = session(addZipField)&""
	end if
	if arrAdd(4)&"" <> "" then
		addZip4Field = cRenew.getFieldName(arrAdd(4))
	end if

	if addLine1Value&"" <> "" and addCityValue&"" <> "" and addStateValue&"" <> "" and addZipField <> "" then

		addVerURL = addVerURL & "&street=" & addLine1Value
		addVerURL = addVerURL & "&city=" & addCityValue
		addVerURL = addVerURL & "&state=" & addStateValue
		addVerURL = addVerURL & "&zipcode=" & addZipValue

		Set addXML = CreateObject("MSXML2.XMLHTTP.6.0")
		addXML.open "GET",addVerURL,False
		addXML.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		addXML.send()
		addResp = addXML.responseText
		set addXML = nothing  

		if instr(addResp,"Active subscription required") then
			addMsg = "Address verification subscription expired"
		elseif addResp&"" <> "[]" then
			if verTyp = "ADD_VER_REP" then
				addResp = replace(addResp,"[{","{")
				addResp = replace(addResp,"}]","}")
				Set addJSON = New aspJSON
				addJSON.loadJSON(addResp)
				session(addLine1Field) = addJSON.data("delivery_line_1")
				session(addCityField) = addJSON.data("components")("city_name")
				session(addStateField) = addJSON.data("components")("state_abbreviation")
				session(addZipField) = addJSON.data("components")("zipcode")
				if addZip4Field&"" <> "" then
					session(addZip4Field) = addJSON.data("components")("plus4_code")
				end if
				set addJSON = nothing
			end if
		else
			addMsg = "Invalid Address"
			if isDev then
				addMsg = addMsg & "<BR><BR>Request - " & addVerURL & "<BR><BR>Response - " & addResp
			end if
		end if
	else
		addMsg = "Incomplete address"
	end if

	if addMsg&"" <> "" then
		cForm.AddErrorMsg(verDesc & " error - " & addMsg)
	end if

end sub

sub formLog(logTxt)

if request.form("auto_save") = "1" then
else
	if session("APP_RENEW_ID")&"" <> "" then
		cLogIt.BatchLog(pageCd & "|" & logTxt)
	end if
end if

end sub

sub updateFields(doJS)

if loadDefaultsDelay = true then
	delayLog "onclick page fields to update",true
	exit sub
else
	delayLog "onclick page fields to update",false
end if


for idxField = lbound(arrFields,2) to ubound(arrFields,2)
	if (arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "YN" or arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "RADIO") and arrFields(fieldNameIndex("FIELD_UPDATE"),idxField)&"" <> "" then
		if arrFields(fieldNameIndex("RADIO_NOT_EQUAL"),idxField) = "Y" then
			if (arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "RADIO" and cstr(session(arrFields(fieldNameIndex("FIELD_NAME"),idxField))&"") <> yearFunction(arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField)&"")) or _
				(arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "YN" and session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) <> "Yes") then
				if objFields.Exists(clng(arrFields(fieldNameIndex("FIELD_UPDATE"),idxField))) then
					fieldName = objFields.Item(clng(arrFields(fieldNameIndex("FIELD_UPDATE"),idxField)))
					if arrFields(fieldNameIndex("MISC"),idxField)&"" <> "" then
						updateVal = arrFields(fieldNameIndex("MISC"),idxField)&""
						if instr(updateVal,"[current") then
							updateVal = replace(updateVal,"[current date]",cstr(sysDate))
							for i = 100 to 1 step -1
								updateVal = replace(updateVal,"[current year - " & i & "]",year(sysDate)-i)
							next

							for i = 1 to 100
								updateVal = replace(updateVal,"[current year + " & i & "]",year(sysDate)+i)
							next

							updateVal = replace(updateVal,"[current year]",year(sysDate))
						end if
					elseif arrFields(fieldNameIndex("related_field"),idxField)&"" <> "" then
						if objFields.Exists(clng(arrFields(fieldNameIndex("related_field"),idxField))) then
							relatedFieldName = objFields.Item(clng(arrFields(fieldNameIndex("related_field"),idxField)))
							updateVal = session(relatedFieldName)
						end if
					end if
					msg = "updating " & fieldName & " to " & updateVal & " because " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " <> " & session(arrFields(fieldNameIndex("FIELD_NAME"),idxField))
					cRenew.AddMsg(msg)
					formLog(msg)
					if doJS then
						%>updateElementID("<%=fieldName%>","<%=updateVal%>");<%
					end if
					session(fieldName) = updateVal
				end if
			end if
		else
			if (arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "RADIO" and cstr(session(arrFields(fieldNameIndex("FIELD_NAME"),idxField))&"") = yearFunction(arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField)&"")) or _
				(arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "YN" and session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "Yes") then
				if objFields.Exists(clng(arrFields(fieldNameIndex("FIELD_UPDATE"),idxField))) then
					fieldName = objFields.Item(clng(arrFields(fieldNameIndex("FIELD_UPDATE"),idxField)))
					if arrFields(fieldNameIndex("MISC"),idxField)&"" <> "" then
						updateVal = arrFields(fieldNameIndex("MISC"),idxField)&""
						if instr(updateVal,"[current") then
							updateVal = replace(updateVal,"[current date]",cstr(sysDate))
							for i = 100 to 1 step -1
								updateVal = replace(updateVal,"[current year - " & i & "]",year(sysDate)-i)
							next

							for i = 1 to 100
								updateVal = replace(updateVal,"[current year + " & i & "]",year(sysDate)+i)
							next

							updateVal = replace(updateVal,"[current year]",year(sysDate))
						end if
					elseif arrFields(fieldNameIndex("related_field"),idxField)&"" <> "" then
						if objFields.Exists(clng(arrFields(fieldNameIndex("related_field"),idxField))) then
							relatedFieldName = objFields.Item(clng(arrFields(fieldNameIndex("related_field"),idxField)))
							updateVal = session(relatedFieldName)
						end if
					end if
					msg = "updating " & fieldName & " to " & updateVal & " because " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " = " & session(arrFields(fieldNameIndex("FIELD_NAME"),idxField))
					cRenew.AddMsg(msg)
					formLog(msg)
					if doJS then
						%>updateElementID("<%=fieldName%>","<%=updateVal%>");<%
					end if
					session(fieldName) = updateVal
				end if
			end if
		end if
	end if
next

end sub

sub numericRangeFields

dim numCompField, numCompFieldVal, minCompVal, maxCompVal, numMinField, numMaxField

if loadDefaultsDelay = true then
	delayLog "numeric range",true
	exit sub
else
	delayLog "numeric range",false
end if

formLog("start numericRangeFields")

for idxField = lbound(arrFields,2) to ubound(arrFields,2)
	if arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "NUMERIC_RANGE" then
		numMinField = ""
		numMaxField = ""
		numCompField = cRenew.getFieldName(arrFields(fieldNameIndex("RELATED_FIELD"),idxField))
		if arrFields(fieldNameIndex("RELATED_FIELD_2"),idxField)&"" <> "" then
			numMinField = cRenew.getFieldName(arrFields(fieldNameIndex("RELATED_FIELD_2"),idxField))
		end if
		if arrFields(fieldNameIndex("RELATED_FIELD_3"),idxField)&"" <> "" then
			numMaxField = cRenew.getFieldName(arrFields(fieldNameIndex("RELATED_FIELD_3"),idxField))
		end if
	
		if session(numCompField)&"" <> "" then
			if isNumeric(session(numCompField)) then
				numCompFieldVal = cdbl(session(numCompField))
			end if
		end if
		if arrFields(fieldNameIndex("TOTAL_MIN"),idxField)&"" <> "" then
			minCompVal = cdbl(arrFields(fieldNameIndex("TOTAL_MIN"),idxField))
		elseif numMinField&"" <> "" then
			minCompVal = cdbl(session(numMinField))
		else
			minCompVal = 0
		end if
		if arrFields(fieldNameIndex("TOTAL_MAX"),idxField)&"" <> "" then
			maxCompVal = cdbl(arrFields(fieldNameIndex("TOTAL_MAX"),idxField))
		elseif numMaxField&"" <> "" then
			maxCompVal = cdbl(session(numMaxField))
		else
			maxCompVal = 1000000
		end if

		if numCompFieldVal&"" <> "" then
			if isNumeric(numCompFieldVal) then
				if numCompFieldVal >= minCompVal and numCompFieldVal <= maxCompVal then
					msg = "setting " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " to " & arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField) & " because " & numCompFieldVal & " >= " & minCompVal & " and <= " & maxCompVal
					cRenew.AddMsg(msg)
					formLog(msg)
					session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField)
				end if
			end if
		end if
	end if
next

formLog("end numericRangeFields")

end sub

sub dateAddFields

formLog("start dateAddFields")

'Date Add fields
for idxField = lbound(arrFields,2) to ubound(arrFields,2)
	if arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "DATE_ADD" then
		if session(cRenew.getFieldName(arrFields(fieldNameIndex("RELATED_FIELD"),idxField)))&"" <> "" then
			session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = DateAddField(session(cRenew.getFieldName(arrFields(fieldNameIndex("RELATED_FIELD"),idxField))),arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField))
			formLog("Date Add Field " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " set to " & session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) & " from " & session(cRenew.getFieldName(arrFields(fieldNameIndex("RELATED_FIELD"),idxField)))  & " and " & arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField))
			msg = "Date Add Field " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " set to " & session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) & " from " & session(cRenew.getFieldName(arrFields(fieldNameIndex("RELATED_FIELD"),idxField)))  & " and " & arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField)
			cRenew.AddMsg(msg)
			formLog(msg)
		end if
	end if
next

formLog("end dateAddFields")

end Sub

sub dateRangeFields

dim dtCompField, dtMinField, dtMaxField, dtCompVal, dtMinVal, dtMaxVal, dateRangeMessage

formLog("start dateRangeFields")

for idxField = lbound(arrFields,2) to ubound(arrFields,2)
	if arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "DATE_RANGE" then
		dtCompField = cRenew.getFieldName(arrFields(fieldNameIndex("RELATED_FIELD"),idxField))
		dtMinField = cRenew.getFieldName(arrFields(fieldNameIndex("RELATED_FIELD_2"),idxField))
		dtMaxField = cRenew.getFieldName(arrFields(fieldNameIndex("RELATED_FIELD_3"),idxField))
		formLog("comparison field name = " & dtCompField)
		formLog("minimum field name = " & dtMinField)
		formLog("maximum field name = " & dtMaxField)
		dtMinVal = cdate("01/01/1800")
		dtMaxVal = cdate("01/01/2100")
		formLog("comparison field value = " & session(dtCompField))
		formLog("minimum field value = " & session(dtMinField))
		formLog("maximum field value = " & session(dtMaxField))
		if trim(session(dtCompField)&"") <> "" then
			if isDate(session(dtCompField)) then
				dtCompVal = cDate(session(dtCompField))
			else
				formLog("comparison field value is not a date")
			end if
		end if
		if trim(session(dtMinField)&"") <> "" then
			if isDate(session(dtMinField)) then
				dtMinVal = cDate(session(dtMinField))
			else
				formLog("minimum field value is not a date")
			end if
		else
			formLog("minimum field value is empty so using default of 01/01/1800")
		end if
		if trim(session(dtMaxField)&"") <> "" then
			if isDate(session(dtMaxField)) then
				dtMaxVal = cDate(session(dtMaxField))
			else
				formLog("maximum field value is not a date")
			end if
		else
			formLog("maximum field value is empty so using default of 01/01/2100")
		end if
		if session(dtCompField)&"" <> "" then
			if isDate(session(dtCompField))&"" <> "" then
				if dtCompVal >= dtMinVal and dtCompVal <= dtMaxVal then
					session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField)
					dateRangeMessage = "setting " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " to " & arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField) & " because " & dtCompVal & " >= " & dtMinVal & " and <= " & dtMaxVal
				else
					dateRangeMessage = "NOT setting " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " to " & arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField) & " because " & dtCompVal & " IS NOT  >= " & dtMinVal & " and <= " & dtMaxVal
				end if
			else
				dateRangeMessage = "NOT setting " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " because " & dtCompField & " is not a date"
			end if
		else
			dateRangeMessage = "NOT setting " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " because " & dtCompField & " is empty"
		end if
		msg = dateRangeMessage
		cRenew.AddMsg(msg)
		formLog(msg)
	end if
next

formLog("end dateRangeFields")

end sub

sub combineFields(doJS)

if loadDefaultsDelay = true then
	delayLog "combine fields",true
	exit sub
else
	delayLog "combine fields",false
end if

formLog("start combineFields")

'combination fields
dim arrCombine, rsCombine, combineField, combineFieldName, combineFieldValue
dim arrTmp
dim arrAllCombine, idxCombine
dim arrCombineFields, arrCombineFieldsJS, idxCombineField

if isArray(session("ARR_COMBINE_FIELDS")) Then
	arrCombineFields = session("ARR_COMBINE_FIELDS")
else
	set cSqlForm = new cRunSql
	cSqlForm.Conn = connOracle
	cSqlForm.SqlStr = "select distinct FIELD_NAME,COMBINE_FIELDS,NO_COMBINE_UNDERSCORE from app_renew_fields where app_renew_type = ? and replace(replace(combine_fields,', ',''),',','') is not null"
	cSqlForm.AddParam(appRenewCd)
	set rsCombine = cSqlForm.Execute
	set cSqlForm = nothing
	if rsCombine.eof then
		arrCombineFields = ""
	else
		arrCombineFields = rsCombine.GetRows()
	end if
	rsCombine.Close
	set rsCombine = nothing
	session("ARR_COMBINE_FIELDS") = arrCombineFields
end if

if isArray(arrCombineFields) then
	for idxCombineField = lbound(arrCombineFields,2) to ubound(arrCombineFields,2)
		combineField = ""
		arrCombine = split(replace(arrCombineFields(1,idxCombineField)&""," ",""),",")
		msg = "combining fields to create " & arrCombineFields(0,idxCombineField)
		cRenew.AddMsg(msg)
		formLog(msg)
		for i = lbound(arrCombine) to ubound(arrCombine)
			if arrCombine(i) <> "" then
				if isNumeric(arrCombine(i)) then
					if objFields.Exists(clng(arrCombine(i))) then 
						combineFieldName = objFields.Item(clng(arrCombine(i)))
						combineFieldValue = session(combineFieldName)
						if arrCombineFields(2,idxCombineField) = "Y" then
						else
							if i > 0 then
								combineField = combineField & "_"
							end if
						end if
						msg = "adding field " & combineFieldName & " with value of " & combineFieldValue
						cRenew.AddMsg(msg)
						formLog(msg)
						combineField = combineField & combineFieldValue
					end if
				end if				
			end if
		next
		msg = "combine field " & arrCombineFields(0,idxCombineField) & " created with value of " & combineField
		cRenew.AddMsg(msg)
		formLog(msg)
		session(arrCombineFields(0,idxCombineField)) = combineField
	next
end if

if doJS then
	if isArray(session("ARR_COMBINE_FIELDS_JS")) Then
		arrCombineFieldsJS = session("ARR_COMBINE_FIELDS_JS")
	else
		set cSqlForm = new cRunSql
		cSqlForm.Conn = connOracle
		cSqlForm.SqlStr = "select distinct FIELD_NAME,COMBINE_FIELDS,NO_COMBINE_UNDERSCORE from app_renew_fields where app_renew_type = ? and replace(replace(combine_fields,', ',''),',','') is not null " &_
			"and (replace(replace(fields_hide,', ',''),',','') is not null or replace(replace(fields_show,', ',''),',','') is not null or replace(replace(fields_require,', ',''),',','') is not null or replace(replace(fields_unrequire,', ',''),',','') is not null)"
		cSqlForm.AddParam(appRenewCd)
		set rsCombine = cSqlForm.Execute
		set cSqlForm = nothing
		if rsCombine.eof then
			arrCombineFieldsJS = ""
		else
			arrCombineFieldsJS = rsCombine.GetRows()
		end if
		rsCombine.Close
		set rsCombine = nothing
		session("ARR_COMBINE_FIELDS_JS") = arrCombineFieldsJS
	end if
	if isArray(arrCombineFieldsJS) then
		%>
		<script>
			const session = {};<%
			if redis then
				arrSesKeys = split(session.GetCacheList,",")
				if isArray(arrSesKeys) then
					for x = lbound(arrSesKeys) to ubound(arrSesKeys)
						item = arrSesKeys(x)
						if isArray(session(item)) then
						else
							if instr(session(item),"""") or instr(session(item),vbcrlf) or right(lcase(item),10) = "_has_value" or right(lcase(item),6) = "_count" then
							else
								%>session["<%=lcase(item)%>"] = "<%=session(item)%>";<%
							end if
						end if
					next
				end if
			else
				for each item in session.contents
					if isArray(session(item)) then
					else
						if instr(session(item),"""") or instr(session(item),vbcrlf) or right(lcase(item),10) = "_has_value" or right(lcase(item),6) = "_count" then
						else
							%>session["<%=lcase(item)%>"] = "<%=session(item)%>";<%
						end if
					end if
				next
			end if
			%>function combineFields() {<%
		for idxCombineField = lbound(arrCombineFieldsJS,2) to ubound(arrCombineFieldsJS,2)
			%>
				var <%=arrCombineFieldsJS(0,idxCombineField)%> = <%
			combineField = ""
			arrCombine = split(replace(arrCombineFieldsJS(1,idxCombineField)&""," ",""),",")
			for i = lbound(arrCombine) to ubound(arrCombine)
				if arrCombine(i) <> "" then
					if isNumeric(arrCombine(i)) then
						if objFields.Exists(clng(arrCombine(i))) then 
							combineFieldName = objFields.Item(clng(arrCombine(i)))
							combineFieldValue = session(combineFieldName)
							if arrCombineFieldsJS(2,idxCombineField) = "Y" then
							else
								if i > 0 then
									combineField = combineField & "_"
								end if
							end if
							if isArray(arrAllCombine) then
								idxCombine = ubound(arrAllCombine) + 1
								redim preserve arrAllCombine(idxCombine)
							else
								idxCombine = 0
								redim arrAllCombine(idxCombine)
							end if
							arrAllCombine(idxCombine) = combineFieldName
							if i > 0 then
								if arrCombineFieldsJS(2,idxCombineField) = "Y" then
									%> + <%
								else
									%> + "_" + <%
								end if
							end if
							%>getFieldValue("<%=combineFieldName%>","<%=objFieldsType(clng(arrCombine(i)))%>")<%
						end if
					end if				
				end if
			next
			%>;
			<%
			for idxField = lbound(arrFields,2) to ubound(arrFields,2)
				if arrFields(fieldNameIndex("FIELD_NAME"),idxField) = arrCombineFieldsJS(0,idxCombineField) then
					%>
					if (<%=arrCombineFieldsJS(0,idxCombineField)%> == "<%=arrFields(fieldNameIndex("default_value"),idxField)%>") {
					<%
					if arrFields(fieldNameIndex("FIELDS_HIDE"),idxField)&"" <> "" then
						arrTmp = split(arrFields(fieldNameIndex("FIELDS_HIDE"),idxField),", ")
						for i = lbound(arrTmp) to ubound(arrTmp)
							if replace(arrTmp(i)," ","") <> "" then
								if objFields.Exists(clng(arrTmp(i))) then
									%>
									hideElementID("<%=objFields.Item(clng(arrTmp(i)))%>");
									<%
								end if
							end if
						next
					end if
					if arrFields(fieldNameIndex("FIELDS_SHOW"),idxField)&"" <> "" then
						arrTmp = split(arrFields(fieldNameIndex("FIELDS_SHOW"),idxField),", ")
						for i = lbound(arrTmp) to ubound(arrTmp)
							if replace(arrTmp(i)," ","") <> "" then
								if objFields.Exists(clng(arrTmp(i))) then
									%>
									showElementID("<%=objFields.Item(clng(arrTmp(i)))%>");
									<%
								end if
							end if
						next
					end if
					if arrFields(fieldNameIndex("FIELDS_REQUIRE"),idxField)&"" <> "" then
						arrTmp = split(arrFields(fieldNameIndex("FIELDS_REQUIRE"),idxField),", ")
						for i = lbound(arrTmp) to ubound(arrTmp)
							if replace(arrTmp(i)," ","") <> "" then
								if objFields.Exists(clng(arrTmp(i))) then
									%>
									requireElementID("<%=objFields.Item(clng(arrTmp(i)))%>");
									<%
								end if
							end if
						next
					end if
					if arrFields(fieldNameIndex("FIELDS_UNREQUIRE"),idxField)&"" <> "" then
						arrTmp = split(arrFields(fieldNameIndex("FIELDS_UNREQUIRE"),idxField),", ")
						for i = lbound(arrTmp) to ubound(arrTmp)
							if replace(arrTmp(i)," ","") <> "" then
								if objFields.Exists(clng(arrTmp(i))) then
									%>
									unrequireElementID("<%=objFields.Item(clng(arrTmp(i)))%>");
									<%
								end if
							end if
						next
					end if
					%>
					}
					<%
				end if
			next
		next
		%>
			}
		<%
		if isArray(arrAllCombine) then
			for idxCombine = lbound(arrAllCombine) to ubound(arrAllCombine)
			%>
			var <%=arrAllCombine(idxCombine)%>s = document.getElementsByName("<%=arrAllCombine(idxCombine)%>");
			for (var x=0;x<<%=arrAllCombine(idxCombine)%>s.length;x++)
				<%=arrAllCombine(idxCombine)%>s[x].setAttribute( "onchange", "combineFields()" );
			<%
			next
		end if
		%>
			combineFields();
		</script>
		<%
	end if
end if

formLog("end combineFields")

end sub

sub numericComparisons


dim numComp1, numComp2 

if loadDefaultsDelay = true then
	delayLog "numeric comparisons",true
	exit sub
else
	delayLog "numeric comparisons",false
end if

formLog("start numericComparisons")

for idxField = lbound(arrFields,2) to ubound(arrFields,2)
	if arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "NUMERIC_COMPARE" then
		if arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField)&"" <> "" then
			numComp1 = arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField)
		else
			numComp1 = session(arrFields(fieldNameIndex("FIELD_NAME"),idxField))
		end if
		numComp1 = numComp1&""
		if objFields.Exists(clng("0" & arrFields(fieldNameIndex("RELATED_FIELD"),idxField))) then
			numComp2 = session(objFields.Item(clng("0" & arrFields(fieldNameIndex("RELATED_FIELD"),idxField))))
		else
			numComp2 = ""
		end if
		if isNumeric(numComp2) and isNumeric(numComp1) then
			numComp1 = cdbl(numComp1)
			numComp2 = cdbl(numComp2)
			if arrFields(fieldNameIndex("LOGIN_COMPARE"),idxField) = "=" then
				if numComp1 = numComp2 then
					msg = "setting " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " to 1 because " & numComp1 & " = " & numComp2
					session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "1"
				else
					msg = "setting " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " to 0 because " & numComp1 & " <> " & numComp2
					session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "0"
				end if
				cRenew.AddMsg(msg)
				formLog(msg)
			end if
			if arrFields(fieldNameIndex("LOGIN_COMPARE"),idxField) = ">" then
				if numComp1 > numComp2 then
					msg = "setting " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " to 1 because " & numComp1 & " > " & numComp2
					session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "1"
				else
					msg = "setting " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " to 0 because " & numComp1 & " <= " & numComp2
					session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "0"
				end if
				cRenew.AddMsg(msg)
				formLog(msg)
			end if
			if arrFields(fieldNameIndex("LOGIN_COMPARE"),idxField) = ">=" then
				if numComp1 >= numComp2 then
					msg = "setting " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " to 1 because " & numComp1 & " >= " & numComp2
					session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "1"
				else
					msg = "setting " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " to 0 because " & numComp1 & " < " & numComp2
					session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "0"
				end if
				cRenew.AddMsg(msg)
				formLog(msg)
			end if
			if arrFields(fieldNameIndex("LOGIN_COMPARE"),idxField) = "<" then
				if numComp1 < numComp2 then
					msg = "setting " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " to 1 because " & numComp1 & " < " & numComp2
					session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "1"
				else
					msg = "setting " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " to 0 because " & numComp1 & " >= " & numComp2
					session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "0"
				end if
				cRenew.AddMsg(msg)
				formLog(msg)
			end if
			if arrFields(fieldNameIndex("LOGIN_COMPARE"),idxField) = "<=" then
				if numComp1 <= numComp2 then
					msg = "setting " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " to 1 because " & numComp1 & " <= " & numComp2
					session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "1"
				else
					msg = "setting " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " to 0 because " & numComp1 & " > " & numComp2
					session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "0"
				end if
				cRenew.AddMsg(msg)
				formLog(msg)
			end if
		end if
	end if
next

formLog("end numericComparisons")

end sub

sub dateComparisons

dim dateComp1, dateComp2

if loadDefaultsDelay = true then
	delayLog "date comparisons",true
	exit sub
else
	delayLog "date comparisons",false
end if

formLog("start dateComparisons")

for idxField = lbound(arrFields,2) to ubound(arrFields,2)
	if arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "DATE_COMPARE" then
		if instr(arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField)&"","[current") then
			dateComp1 = replace(arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField)&"","[current date]",cstr(sysDate))
		elseif arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField)&"" <> "" then
			dateComp1 = session(arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField))
		else
			dateComp1 = session(arrFields(fieldNameIndex("FIELD_NAME"),idxField))
		end if
		dateComp1 = dateComp1&""
		for i = 1 to 100
			dateComp1 = replace(dateComp1,"[current date + " & i  & " days]",dateAdd("d",i,sysDate))
		next
		for i = 1 to 100
			dateComp1 = replace(dateComp1,"[current date - " & i  & " days]",dateAdd("d",-i,sysDate))
		next
		for i = 1 to 100
			dateComp1 = replace(dateComp1,"[current date + " & i  & " months]",dateAdd("m",i,sysDate))
		next
		for i = 1 to 100
			dateComp1 = replace(dateComp1,"[current date - " & i  & " months]",dateAdd("m",-i,sysDate))
		next
		for i = 1 to 100
			dateComp1 = replace(dateComp1,"[current date + " & i  & " years]",dateAdd("yyyy",i,sysDate))
		next
		for i = 1 to 100
			dateComp1 = replace(dateComp1,"[current date - " & i  & " years]",dateAdd("yyyy",-i,sysDate))
		next
		for i = 100 to 1 step -1
			dateComp1 = replace(dateComp1,"[current year - " & i & "]",year(sysDate)-i)
		next
		for i = 1 to 100
			dateComp1 = replace(dateComp1,"[current year + " & i & "]",year(sysDate)+i)
		next
		dateComp1 = replace(dateComp1,"[current year]",year(sysDate))

		if objFields.Exists(clng("0" & arrFields(fieldNameIndex("RELATED_FIELD"),idxField))) then
			dateComp2 = session(objFields.Item(clng("0" & arrFields(fieldNameIndex("RELATED_FIELD"),idxField))))
		else
			dateComp2 = ""
		end if

		if isDate(dateComp2) and isDate(dateComp1) then
			dateComp1 = cdate(dateComp1)
			dateComp2 = cdate(dateComp2)
			if arrFields(fieldNameIndex("LOGIN_COMPARE"),idxField) = "=" then
				if dateComp1 = dateComp2 then
					msg = "setting " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " to 1 because " & dateComp1 & " = " & dateComp2
					session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "1"
				else
					msg = "setting " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " to 0 because " & dateComp1 & " <> " & dateComp2
					session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "0"
				end if
				cRenew.AddMsg(msg)
				formLog(msg)
			end if
			if arrFields(fieldNameIndex("LOGIN_COMPARE"),idxField) = ">" then
				if dateComp1 > dateComp2 then
					msg = "setting " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " to 1 because " & dateComp1 & " > " & dateComp2
					session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "1"
				else
					msg = "setting " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " to 0 because " & dateComp1 & " <= " & dateComp2
					session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "0"
				end if
				cRenew.AddMsg(msg)
				formLog(msg)
			end if
			if arrFields(fieldNameIndex("LOGIN_COMPARE"),idxField) = ">=" then
				if dateComp1 >= dateComp2 then
					msg = "setting " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " to 1 because " & dateComp1 & " >= " & dateComp2
					session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "1"
				else
					msg = "setting " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " to 0 because " & dateComp1 & " < " & dateComp2
					session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "0"
				end if
				cRenew.AddMsg(msg)
				formLog(msg)
			end if
			if arrFields(fieldNameIndex("LOGIN_COMPARE"),idxField) = "<" then
				if dateComp1 < dateComp2 then
					msg = "setting " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " to 1 because " & dateComp1 & " < " & dateComp2
					session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "1"
				else
					msg = "setting " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " to 0 because " & dateComp1 & " >= " & dateComp2
					session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "0"
				end if
				cRenew.AddMsg(msg)
				formLog(msg)
			end if
			if arrFields(fieldNameIndex("LOGIN_COMPARE"),idxField) = "<=" then
				if dateComp1 <= dateComp2 then
					msg = "setting " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " to 1 because " & dateComp1 & " <= " & dateComp2
					session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "1"
				else
					msg = "setting " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " to 0 because " & dateComp1 & " > " & dateComp2
					session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "0"
				end if
				cRenew.AddMsg(msg)
				formLog(msg)
			end if
		end if
	end if
next

formLog("end dateComparisons")

end sub

sub dateDifferences

dim dateComp1, dateComp2, dateDifference

if loadDefaultsDelay = true then
	delayLog "date differences",true
	exit sub
else
	delayLog "date differences",false
end if

formLog("start dateDifferences")

for idxField = lbound(arrFields,2) to ubound(arrFields,2)
	if arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "DATE_DIFF" then
		if objFields.Exists(clng("0" & arrFields(fieldNameIndex("RELATED_FIELD"),idxField))) then
			dateComp1 = session(objFields.Item(clng("0" & arrFields(fieldNameIndex("RELATED_FIELD"),idxField))))
		else
			dateComp1 = ""
		end if
		if objFields.Exists(clng("0" & arrFields(fieldNameIndex("RELATED_FIELD_2"),idxField))) then
			dateComp2 = session(objFields.Item(clng("0" & arrFields(fieldNameIndex("RELATED_FIELD_2"),idxField))))
		else
			dateComp2 = ""
		end if

		if isDate(dateComp2) and isDate(dateComp1) then
			dateComp1 = cdate(dateComp1)
			dateComp2 = cdate(dateComp2)
			dateDifference = abs(dateDiff(arrFields(fieldNameIndex("DATE_DIFF"),idxField),dateComp1,dateComp2))
			msg = "setting " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " to " & dateDifference & " - Difference = " & arrFields(fieldNameIndex("DATE_DIFF"),idxField) & " Date 1 = " & dateComp1 & " Date 2 = " & dateComp2
			cRenew.AddMsg(msg)
			formLog(msg)
			session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = dateDiff(arrFields(fieldNameIndex("DATE_DIFF"),idxField),dateComp1,dateComp2)
		end if
	end if
next

formLog("end dateDifferences")

end sub

sub loadSavedAppRenew

dim rsAppRenew, rsSavedAppRenew

set cSqlForm = new cRunSql
cSqlForm.Conn = connOracle
cSqlForm.SqlStr = "select APP_RENEW_ID from app_renew where membership_id = ? and app_renew_status = 'SAVED' order by app_renew_id desc"
cSqlForm.AddParam(session("MEMBERSHIP_ID"))
set rsAppRenew = cSqlForm.Execute
set cSqlForm = nothing
if rsAppRenew.eof and rsAppRenew.bof then
else
	set cSqlForm = new cRunSql
	cSqlForm.Conn = connOracle
	cSqlForm.SqlStr = "select FIELD_NAME,DATA from app_renew_detail where app_renew_id = ? and upper(field_name) not in ("
	for i = lbound(arrSysVars) to ubound(arrSysVars)
		cSqlForm.AddSqlStr("upper('" & arrSysVars(i) & "')")
		if i < ubound(arrSysVars) then
			cSqlForm.AddSqlStr(",")
		end if
	next
	cSqlForm.AddSqlStr(")")
	cSqlForm.AddParam(rsAppRenew("APP_RENEW_ID"))
	set rsSavedAppRenew = cSqlForm.Execute
	do while not rsSavedAppRenew.eof
		session(rsSavedAppRenew("FIELD_NAME")) = cstr(rsSavedAppRenew("DATA"))
		rsSavedAppRenew.MoveNext
	loop
	rsSavedAppRenew.close
	set rsSavedAppRenew = nothing
	set cSqlForm = nothing
end if

rsAppRenew.close
set rsAppRenew = nothing

end sub

sub saveAppRenewForLater
	saveFormData("SAVED")
end sub

sub saveAppRenew
	saveFormData("PENDING")
end sub

sub saveAppRenewPending
	saveFormData("PENDING")
end sub

sub saveFormPayLater
	saveFormData("PAY_LATER")
end sub

sub saveFormDataSession()

	if request("action") = "loadDebug" or request.form("do") = "exit" or request.querystring("preview") = "1" or request.querystring("doPDF") = "1" or request.querystring("doEmail") = "1" then
		exit sub
	end if

    dim cSqlSave, item, idxRow, idxCol, pageCd, sesVal, rsARD, arrD, idxD
    dim fieldName, fieldData, detailFound, arID

	if redis then
		url = "https://forms.ebigpicture.com/forms.ashx?action=saveFormDataSession&key=" & session("redisSessionID") & "&script_name=" & request.servervariables("SCRIPT_NAME")
		ajaxURL url,"GET",false
		exit sub
	end if

	formLog("start saveFormDataSession")

    maxDataLength = 32767

    renewID = session("APP_RENEW_ID")

	formLog("start clearing items no longer in session from app_renew_detail")
	set cSqlSave = new cRunSql
	cSqlSave.Conn = connOracle
    cSqlSave.SqlStr = "select field_name,data from app_renew_detail where app_renew_id = ?"
    cSqlSave.AddParam(renewID)
    set rsARD = cSqlSave.Execute
    set cSqlSave = nothing
    if rsARD.eof then
    else
        arrD = rsARD.GetRows
    end if
    rsARD.close
    set rsARD = nothing

    if isArray(arrD) then
        for idxD = lbound(arrD,2) to ubound(arrD,2)
            fieldName = arrD(0,idxD)&""
            fieldData = arrD(1,idxD)&""
            sesVal = session(fieldName)&""
            if fieldData&"" <> "" then
                if sesVal = "" then
            	    formLog("deleting " & fieldName & " from app_renew_detail because table value = " & fieldData & " and session value = " & sesVal)
	                set cSqlSave = new cRunSql
	                cSqlSave.Conn = connOracle
                    cSqlSave.SqlStr = "delete from app_renew_detail where app_renew_id = ? and upper(field_name) = upper(?)"
                    cSqlSave.AddParam(renewID)
                    cSqlSave.AddParam(fieldName)
                    cSqlSave.Execute
                    set cSqlSave = nothing
                end if
            end if
        next
    end if
        
    formLog("end clearing items no longer in session from app_renew_detail")

	redisKey = "FORM_" & Application("ORG_ID") & "_" & session("APP_RENEW_ID")
    formLog("redis key " & redisKey & " created for saving arrays")
    set redisArr = Server.CreateObject("CacheManager")
	redisArr.SetDefaultExpiration 365,0,0,0
	redisArr.Init redisKey
	redisArr.Abandon
	
	for each item in session.contents
		if isArray(session(item)) then
			formLog("saving array " & item & " to redis cache")
			redisArr(item) = session(Item)
		else
			if instr(lcase(item),"arrrelate") or lcase(item) = "password" or right(item,1) = ")" then
			else
				if item&"" <> "" then
					sesVal = session(item)&""
					if sesVal&"" <> "" then
						if len(sesVal) > maxDataLength then
							sesVal = left(sesVal,maxDataLength)
						end if
						fieldData = ""
						if isArray(arrD) then
							for idxD = lbound(arrD,2) to ubound(arrD,2)
								fieldName = arrD(0,idxD)&""
								if lcase(fieldName) = lcase(item) then
									fieldData = arrD(1,idxD)&""
									idxD = ubound(arrD,2) + 1
								end if
							next
						end if
						if fieldData = "" and sesVal <> "" then
							formLog("start inserting " & item & " = " & sesVal)
							set cSqlSave = new cRunSql
							cSqlSave.Conn = connOracle
							cSqlSave.SqlStr = "insert into app_renew_detail (app_renew_id,field_name,data) values (?,?,?)"
							cSqlSave.AddParam(renewID)
							cSqlSave.AddParam(item)
							cSqlSave.AddParam(sesVal)
							formLog(cSqlSave.SqlStr & " ," & renewID & "," & item & "," & sesVal)
							cSqlSave.Execute
							set cSqlSave = nothing
							formLog("end inserting " & item & " = " & sesVal)
						elseif fieldData <> "" and sesVal <> "" then
							if fieldData <> sesVal then
								formLog("start updating " & item & " = " & sesVal)
								set cSqlSave = new cRunSql
								cSqlSave.Conn = connOracle
								cSqlSave.SqlStr = "update app_renew_detail set data = ? where lower(field_name) = lower(?) and app_renew_id = ?"
								formLog(cSqlSave.SqlStr & " ," & sesVal & "," & item & "," & renewID)
								cSqlSave.AddParam(sesVal)
								cSqlSave.AddParam(item)
								cSqlSave.AddParam(renewID)
								cSqlSave.Execute
								set cSqlSave = nothing
								formLog("end updating " & item & " = " & sesVal)
							end if
						end if
					end if
				end if
			end if
		end if
	next

	set redisArr = nothing

	formLog("end saveFormDataSession")

end sub

sub saveFormData(formStatus)

if request("doPDF")&"" = "1" or request("doEmail")&"" = "1" then
	formLog("doPDF = 1 or doEmail = 1 so exiting saveFormData")
	exit Sub
end if

formLog("start saveFormData - " & formStatus)

if session("MEMBERSHIP_ID") = "" then
    session("MEMBERSHIP_ID") = 0
end if

dim renewID, cSqlSave, rsSave, saveSession, formName

if session("SAVED_APPRENEW_ID")&"" <> "" then
	session("APP_RENEW_ID") = session("SAVED_APPRENEW_ID")
end if

if session("APP_RENEW_ID")&"" <> "" then

	renewID = session("APP_RENEW_ID")
	if formStatus = "PROGRESS" then
		set cSqlSave = new cRunSql
		cSqlSave.Conn = connOracle
		if session("MEMBERSHIP_ID") <> "0" then
			cSqlSave.SqlStr = "update app_renew set membership_id = ?, app_renew_type = ? where app_renew_id = ?"
			cSqlSave.AddParam(session("MEMBERSHIP_ID"))
		else
			cSqlSave.SqlStr = "update app_renew set app_renew_type = ?  where app_renew_id = ?"
		end if
		cSqlSave.AddParam(appRenewCd)
		cSqlSave.AddParam(renewID)
		cSqlSave.Execute
		set cSqlSave = nothing
	else
		set cSqlSave = new cRunSql
		cSqlSave.Conn = connOracle
		if session("MEMBERSHIP_ID") <> "0" then
			cSqlSave.SqlStr = "update app_renew set membership_id = ?, app_renew_status = ?, app_renew_type = ?  where app_renew_id = ?"
			cSqlSave.AddParam(session("MEMBERSHIP_ID"))
		else
			cSqlSave.SqlStr = "update app_renew set app_renew_status = ?, app_renew_type = ?  where app_renew_id = ?"
		end if
		cSqlSave.AddParam(formStatus)
		cSqlSave.AddParam(appRenewCd)
		cSqlSave.AddParam(renewID)
		cSqlSave.Execute
		set cSqlSave = nothing
	end if
else
	renewID = getSequence("SEQ_APP_RENEW_ID")
	session("APP_RENEW_ID") = renewID

	on error resume next
	response.cookies("APP_RENEW_ID") = renewID
	response.cookies("APP_RENEW_ID").Expires = DateAdd("N",sessionTimeout,Now)
	response.cookies("APP_RENEW_CD") = appRenewCd
	response.cookies("APP_RENEW_CD").Expires = DateAdd("N",sessionTimeout,Now)
	on error goto 0

	set cSqlSave = new cRunSql
	cSqlSave.Conn = connOracle
	cSqlSave.SqlStr = "insert into app_renew(app_renew_id,app_renew_type,app_renew_status,membership_id) values (?,?,?,?)"
	cSqlSave.AddParam(renewID)
	cSqlSave.AddParam(appRenewCd)
	cSqlSave.AddParam(formStatus)
	cSqlSave.AddParam(session("MEMBERSHIP_ID"))
	cSqlSave.Execute
	set cSqlSave = nothing
end if

formName = ""
if session("COMPANY") & "" <> "" then
	formName = session("COMPANY")
elseif session("LAST_NAME")&"" = "" then
	if session("FIRST_NAME")&"" <> "" and session("MAIDEN_NAME")&"" <> "" Then
		formName = session("FIRST_NAME") & " " & session("MAIDEN_NAME")
	end if
elseif session("FIRST_NAME")&"" <> "" and session("LAST_NAME")&"" <> "" Then
	formName = session("FIRST_NAME") & " " & session("LAST_NAME")
end if
if formName <> "" then
	set cSqlSave = new cRunSql
	cSqlSave.Conn = connOracle
	cSqlSave.SqlStr = "update app_renew set name = ? where app_renew_id = ? and name is null"
	cSqlSave.AddParam(formName)
	cSqlSave.AddParam(renewID)
	cSqlSave.execute
	set cSqlSave = nothing
end if

saveSession = false

dim cRenewSave 
set cRenewSave = new cAppRenew
if cRenewSave.ConfigOption("save_session_type") = "A" then
	formLog("saving session data because config is set to save all page submits")
	saveSession = true
end if

if cRenewSave.ConfigOption("save_session_type") = "C" then
	if formStatus = "PROGRESS" then
		formLog("NOT saving session data because formStatus = " & formStatus & " and config is set to only save non-progress forms")
	else
		formLog("saving session data because formStatus = " & formStatus & " and config is set to only save non-progress forms")
		saveSession = true
	end if
end if
dim hasValueFields
hasValueFields = cRenewSave.ConfigOption("has_value_fields")&""
set cRenewSave = nothing

if saveSession then
	if hasValueFields = "Y" then
		if redis then
			arrKeys = split(session.GetCacheList,",")

			dim idxSes,itemSes

			for idxSes = lbound(arrKeys) to ubound(arrKeys)
				itemSes = arrKeys(idxSes)
				if instr(ucase(itemSes),"_HAS_VALUE") then
					session(itemSes) = ""
					session.remove(itemSes)
				end if
			next
			for idxSes = lbound(arrKeys) to ubound(arrKeys)
				itemSes = arrKeys(idxSes)
				if instr(ucase(itemSes),"_HAS_VALUE") then
				else
					if isArray(session(itemSes)) then
					elseif right(ucase(itemSes),6) = "_COUNT" then
					else
						if session(itemSes)&"" <> "" then
							session(itemSes & "_HAS_VALUE") = "1"
						else
							session(itemSes & "_HAS_VALUE") = "0"
						end if
					end if
				end if
			next
		else
			for each item in session.contents
				if instr(ucase(item),"_HAS_VALUE") then
					session(item) = ""
					session.contents.remove(itemSes)
				end if
			next
			for each item in session.contents
				if instr(ucase(item),"_HAS_VALUE") then
				else
					if isArray(session(item)) then
					elseif right(ucase(item),6) = "_COUNT" then
					else
						if session(item)&"" <> "" then
							session(item & "_HAS_VALUE") = "1"
						else
							session(item & "_HAS_VALUE") = "0"
						end if
					end if
				end if
			next
		end if
	end if

	saveFormDataSession()
end if

' === FORM SUBMIT LEDGER PATCH ===
' Mirror the app_renew row just written into the durable submission ledger.
' (doPDF / doEmail passes already exited at the top of this sub.)
cLedger.SetContext session("APP_RENEW_ID"), appRenewCd, session("MEMBERSHIP_ID")
cLedger.EnsureSubmitRecord formStatus
' === END FORM SUBMIT LEDGER PATCH ===

formLog("end saveFormData - " & formStatus)

end sub

sub logPageSubmit()

dim cSqlPS, sqlPS, pageSubmitID, pageSubData, redisPS, redisPSID, typeLib, cRenewLog

appRenewID = session("APP_RENEW_ID")
pageSubmitID = getSequence("SEQ_APP_RENEW_PAGE_SUBMIT_ID")

set cSqlPS = new cRunSql
cSqlPS.Conn = connOracle
cSqlPS.SqlStr = "insert into APP_RENEW_PAGE_SUBMIT (APP_RENEW_PAGE_SUBMIT_ID,APP_RENEW_ID) values (?,?)"
cSqlPS.AddParam(pageSubmitID)
cSqlPS.AddParam(appRenewID)
cSqlPS.Execute
set cSqlPS = nothing

sqlPS = "insert into APP_RENEW_PAGE_SUB_DETAIL (APP_RENEW_PAGE_SUBMIT_ID,DATA_TYPE,FIELD_NAME,DATA) values (?,?,?,?)"

dim arrRS(2), idxRS
arrRS(0) = "HTTP_USER_AGENT"
arrRS(1) = "HTTP_X_FORWARDED_FOR"
arrRS(2) = "REMOTE_ADDR"

'request.servervariables
for idxRS = lbound(arrRS) to ubound(arrRS)
	set cSqlPS = new cRunSql
	cSqlPS.Conn = connOracle
	cSqlPS.SqlStr = sqlPS
	cSqlPS.AddParam(pageSubmitID)
	cSqlPS.AddParam("RS")
	cSqlPS.AddParam(arrRS(idxRS))
	pageSubData = request.servervariables(arrRS(idxRS))
	if len(pageSubData) > 6000 then
		pageSubData = left(pageSubData,6000)
	end if
    if item = "SCRIPT_NAME" then
        pageSubData = request("script_name")
    end if
	cSqlPS.AddParam(pageSubData)
	cSqlPS.Execute
	set cSqlPS = nothing
next

'request.form
for each item in request.form
	if lcase(item) = "password" then
	else
		set cSqlPS = new cRunSql
		cSqlPS.Conn = connOracle
		cSqlPS.SqlStr = sqlPS
		cSqlPS.AddParam(pageSubmitID)
		cSqlPS.AddParam("RF")
		cSqlPS.AddParam(item)
		pageSubData = request.form(item)
		if len(pageSubData) > 6000 then
			pageSubData = left(pageSubData,6000)
		end if
		cSqlPS.AddParam(pageSubData)
		cSqlPS.Execute
		set cSqlPS = nothing
	end if
next

set cRenewLog = new cAppRenew
if cRenewLog.ConfigOption("page_submit_log_type")&"" = "" or cRenewLog.ConfigOption("page_submit_log_type")&"" = "A" then
	'request.querystring
	for each item in request.querystring
		if lcase(item) = "password" then
		else
			set cSqlPS = new cRunSql
			cSqlPS.Conn = connOracle
			cSqlPS.SqlStr = sqlPS
			cSqlPS.AddParam(pageSubmitID)
			cSqlPS.AddParam("RQ")
			cSqlPS.AddParam(item)
			pageSubData = request.querystring(item)
			if len(pageSubData) > 6000 then
				pageSubData = left(pageSubData,6000)
			end if
			cSqlPS.AddParam(pageSubData)
			cSqlPS.Execute
			set cSqlPS = nothing
		end if
	next

	'cookies
	for each item in request.Cookies
		if request.Cookies(item).HasKeys then
		else
			set cSqlPS = new cRunSql
			cSqlPS.Conn = connOracle
			cSqlPS.SqlStr = sqlPS
			cSqlPS.AddParam(pageSubmitID)
			cSqlPS.AddParam("C")
			cSqlPS.AddParam(item)
			pageSubData = request.Cookies(item)
			if len(pageSubData) > 6000 then
				pageSubData = left(pageSubData,6000)
			end if
			cSqlPS.AddParam(pageSubData)
			cSqlPS.Execute
			set cSqlPS = nothing
		end if
	next

	if redis then
		redisPSID = session("redissessionid")
	else
		Set TypeLib = CreateObject("Scriptlet.TypeLib")
		redisPSID = replace(Mid(TypeLib.Guid, 2, 36),"-","")
		set TypeLib = nothing
		set redisPS = Server.CreateObject("CacheManager")
		redisPS.Init redisPSID
		for each item in session.contents
			redisPS(item) = session(item)
		next
		set redisPS = nothing
	end if

	url = "https://forms.ebigpicture.com/forms.ashx?action=logPageSubmit&key=" & redisPSID & "&page_sub_id=" & pageSubmitID & "&script_name=" & request.servervariables("SCRIPT_NAME")
	ajaxURL url,"GET",true
end if

set cRenewLog = nothing

end sub

sub delayLog(theProcess,skipProcess)
	dim delayMsg

	delayMsg = "running " & theProcess & " code"
	if skipProcess Then
		delayMsg = "NOT " & delayMsg
	end if
	formLog(delayMsg)
	formLog("DEFAULTS_LOADED = " & session("DEFAULTS_LOADED"))
	formLog("LOAD_DEFAULTS_DELAY = " & cRenew.ConfigOption("load_defaults_delay"))
end sub

function DateAddField(fieldValue,defaultValue)

dim arrDtType(1,2), idxDtTyp1, idxDtTyp2, retVal, dtVal, monVal, dtAddFound, dayVal

retVal = fieldValue

if not isDate(fieldValue) Then
	exit function
end if

dtVal = cDate(fieldValue)
monVal = month(dtVal)
dayVal = day(dtVal)

arrDtType(0,0) = "years"
arrDtType(1,0) = "yyyy"
arrDtType(0,1) = "months"
arrDtType(1,1) = "m"
arrDtType(0,2) = "days"
arrDtType(1,2) = "d"

dtAddFound = false

do while not dtAddFound
	for idxDtTyp = lbound(arrDtType,2) to ubound(arrDtType,2)
		if instr(defaultValue,arrDtType(0,idxDtTyp)) then
			for idxDtTyp2 = 1 to 2000
				if defaultValue = "[current value + " & idxDtTyp2 & " " & arrDtType(0,idxDtTyp) & "]" then
					retVal = cstr(dateAdd(arrDtType(1,idxDtTyp),idxDtTyp2,dtVal))
					exit do
				end if
				if defaultValue = "[current value - " & idxDtTyp2 & " " & arrDtType(0,idxDtTyp) & "]" then
					retVal = cstr(dateAdd(arrDtType(1,idxDtTyp),-idxDtTyp2,dtVal))
					exit do
				end if
			next
		end if
	next
loop

if retVal&"" <> "" then
	if monVal = 2 and arrDtType(1,idxDtTyp) = "yyyy" then
		if year(cDate(retVal)) mod 4 = 0 then
			if dayVal = 28 or dayVal = 29 then
				retVal = "02/29/" & year(cDate(retVal))
			end if
		else
			if dayVal = 28 or dayVal = 29 then
				retVal = "02/28/" & year(cDate(retVal))
			end if
		end if
	end if
end if

DateAddField = retVal

end function

function saveToFile(fileContents, fileType)

dim oFs, oTextFile, fileName
fileName = fileType&udate(sysNow)&".txt"
set oFs = server.createobject("Scripting.FileSystemObject")
set oTextFile = oFs.OpenTextFile(Application("ResourcePath")&"\"&fileName, 2, True)
oTextFile.Write fileContents
oTextFile.Close
set oTextFile = nothing
set oFs = nothing
saveToFile = fileName

end function

function saveToHTMLFile(fileContents, fileType)

dim oFs, oTextFile, fileName
fileName = fileType&udate(sysNow)&".html"
set oFs = server.createobject("Scripting.FileSystemObject")
response.write Application("ResourcePath")&"\"&fileName & "<BR>"
set oTextFile = oFs.CreateTextFile(Application("ResourcePath")&"\"&fileName, 2, True)
oTextFile.Write fileContents
oTextFile.Close
set oTextFile = nothing
set oFs = nothing
saveToHTMLFile = fileName

end function


Sub Error(Method)
    Response.Write "'" & Method & "' failed with a '" & varReturn & "'<br>"
    Set objTK = Nothing
    Response.End
End Sub



function UDate(oldDate)
UDate = DateDiff("s", "01/01/1970 00:00:00", oldDate)
end function

Function getBatchID(batch_num)
dim rsBatch, cSqlBatch

set cSqlBatch = new cRunSql
cSqlBatch.conn = connOracle

cSqlBatch.sqlStr = "select batch_id from batch where batch_num = ? and rownum < 2"
cSqlBatch.AddParam(batch_num)
set rsBatch = cSqlBatch.Execute
getBatchID = rsBatch("batch_id") 
rsBatch.close
set rsBatch = nothing

set cSqlBatch = nothing

End Function

Function getBatch_NUM

Dim mnt,yr,today, batchDate

if Application("ORG_ID") = "775" and hour(sysNow) > 22 then
	batchDate = dateAdd("d",1,sysDate)
else
	batchDate = sysDate
end if

mnt = Month(batchDate)

if len(mnt) < 2 then mnt = "0"&mnt end if

yr = Year(batchDate)
yr = Cstr(yr)
yr = Right(yr,2)

today = day(batchDate)

if len(today) < 2 then today = "0"&today end if
	
getBatch_NUM = yr&""&mnt&""&today&""&"WEB"

End Function

Sub updateBatchInfo(batch_num,amount)
dim rsBatch, batchStatus, cSqlBatch, BATCH_ID

set cSqlBatch = new cRunSql
cSqlBatch.conn = connOracle

batchStatus = "OPEN"

if Application("ORG_ID") = "825" or Application("ORG_ID") = "2" then
	batchStatus = "CLOSED"
end if

cSqlBatch.sqlStr = "select batch_num from batch where batch_num = '"&batch_num&"' and rownum < 2"
set rsBatch = cSqlBatch.Execute
set cSqlBatch = nothing

set cSqlBatch = new cRunSql
cSqlBatch.conn = connOracle
if rsBatch.EOF and rsBatch.BOF then		 '' --------- no batchs yet today.
	cSqlBatch.SqlStr = "insert into batch (batch_num,status_cd,amount,org_id) values ('"&batch_num&"','" & batchStatus & "','"&amount&"','" & Application("ORG_ID") & "') "
else
	BATCH_ID = getBatchID(batch_num)
	cSqlBatch.SqlStr = "update batch set amount = (select sum(amount) from transaction_log where trans_status_cd in ('PENDING','APPROVED') and batch_id = ?) where batch_id = ?"
	cSqlBatch.AddParam(BATCH_ID)
	cSqlBatch.AddParam(BATCH_ID)
End If

cSqlBatch.Execute
rsBatch.close
set rsBatch = nothing
set cSqlBatch = nothing

End Sub

Sub updateBatch_Member_Finder(batch_id,membership_id)

dim rsBMF, cSqlBatch

session("BATCH_MEMBER_ID") = getSequence("SEQ_BATCH_MEMBER_ID")

set cSqlBatch = new cRunSql
cSqlBatch.conn = connOracle

cSqlBatch.SqlStr = "select batch_id from batch_member_finder where batch_id = ? and membership_id = ?"
cSqlBatch.AddParam(batch_id)
cSqlBatch.AddParam(membership_id)
set rsBMF = cSqlBatch.Execute
set cSqlBatch = nothing

if rsBMF.eof and rsBMF.bof then
	set cSqlBatch = new cRunSql
	cSqlBatch.conn = connOracle
    cSqlBatch.SqlStr = " insert into batch_member_finder (batch_member_id,batch_id,member_id, membership_id) values (?,?,(select org_spec_id from membership where membership_id = ?),?)"
	cSqlBatch.AddParam(session("BATCH_MEMBER_ID"))
    cSqlBatch.AddParam(batch_id)
    cSqlBatch.AddParam(membership_id)
    cSqlBatch.AddParam(membership_id)
    cSqlBatch.Execute
	set cSqlBatch = nothing
end if

rsBMF.close
set rsBMF = nothing

End Sub

function finalFieldCheck

dim returnPage, arrRetPage, fieldCheck
dim objReq, arrReq

set objReq = CreateObject("Scripting.Dictionary")

finalFieldCheck = ""

arrRetPage = split(request.querystring("FROM_PAGE"),"/")

set cSqlForm = new cRunSql
cSqlForm.conn = connOracle
cSqlForm.SqlStr = "select app_renew_fields.FIELD_NAME " &_
	",app_renew_fields.REQUIRED " &_
	",app_renew_fields.FIELD_TYPE " &_
	",app_renew_fields.PAGE " &_
	",app_renew_fields.FIELDS_REQUIRE " &_
	",app_renew_fields.DEFAULT_VALUE " &_
	"from app_renew_fields, app_renew_pages " &_
	"where upper(app_renew_pages.page) = upper(app_renew_fields.page) and app_renew_pages.app_renew_type = app_renew_fields.app_renew_type " &_
	"and app_renew_fields.app_renew_type = ? and app_renew_fields.page is not null and lower(app_renew_fields.page) <> 'index.asp' " &_
	"and lower(app_renew_fields.page) <> 'pay.asp' and app_Renew_pages.order_by > 0 " &_
	"and ((app_renew_fields.required = 'Y') or (app_renew_fields.field_type = 'RADIO' and app_renew_fields.fields_require is not null)) " &_
	"and app_renew_pages.required_check_exempt = 'N' " &_
	"order by app_renew_pages.order_by "
cSqlForm.AddParam(appRenewCd)
set rsFields = cSqlForm.Execute

strReqDebug = ""

do while not rsFields.eof and finalFieldCheck = ""
	if rsFields("REQUIRED") = "Y" then
		fieldCheck = ucase(rsFields("FIELD_NAME"))
		if rsFields("FIELD_TYPE") = "FILE" or rsFields("FIELD_TYPE") = "FILEBTN" then
			fieldCheck = "file_" & ucase(rsFields("FIELD_NAME"))
		end if
		if session(fieldCheck) = "" then
			if strReqDebug <> "" then
				strReqDebug = strReqDebug  & "<BR>"
			end if
			if request.QueryString("FROM_PAGE") <> "" then
				strReqDebug = strReqDebug & fieldCheck & " on " & replace(request.QueryString("FROM_PAGE"),arrRetPage(ubound(arrRetPage)),rsFields("PAGE")) & " is required and empty"
				objReq(fieldCheck) = replace(request.QueryString("FROM_PAGE"),arrRetPage(ubound(arrRetPage)),rsFields("PAGE"))
			else
				strReqDebug = strReqDebug & fieldCheck & " on " & rsFields("PAGE") & " is required and empty"
				objReq(fieldCheck) = rsFields("PAGE")
			end if
		end if
	end if
	if rsFields("FIELDS_REQUIRE")&"" <> "" then
		if session(rsFields("FIELD_NAME"))&"" = yearFunction(rsFields("DEFAULT_VALUE")&"") then
			arrTmp = split(rsFields("FIELDS_REQUIRE"),", ")
			for i = lbound(arrTmp) to ubound(arrTmp)
				if replace(arrTmp(i)," ","") <> "" then
					if objFields.Exists(clng(arrTmp(i))) then
						fieldName = objFields.Item(clng(arrTmp(i)))
						if objFieldsType(clng(arrTmp(i))) = "FILE" or objFieldsType(clng(arrTmp(i))) = "FILEBTN" then
							fieldName = "file_" & fieldName
						end if
						if session(fieldName) = "" then
							if objReq.Exists(fieldName) then
							else
								if strReqDebug <> "" then
									strReqDebug = strReqDebug  & "<BR>"
								end if
								strReqDebug = strReqDebug & fieldName & " on " & objFieldsPage(clng(arrTmp(i))) & " is empty and required because " & rsFields("FIELD_NAME") & " = " & session(rsFields("FIELD_NAME"))
								objReq(fieldName) = objFieldsPage(clng(arrTmp(i)))
							end if
						end if
					end if
				end if
			next
		end if
	end if
    rsFields.MoveNext
loop

rsFields.close
set rsFields = nothing
set cSqlForm = nothing


set cSqlForm = new cRunSql
cSqlForm.conn = connOracle
cSqlForm.SqlStr = "select app_renew_fields.FIELD_NAME " &_
	",app_renew_fields.DEFAULT_VALUE " &_
	",app_renew_fields.FIELDS_UNREQUIRE " &_
	"from app_renew_fields, app_renew_pages " &_
	"where upper(app_renew_pages.page) = upper(app_renew_fields.page) and app_renew_pages.app_renew_type = app_renew_fields.app_renew_type " &_
	"and app_renew_fields.app_renew_type = ? and app_renew_fields.page is not null and lower(app_renew_fields.page) <> 'index.asp' " &_
	"and lower(app_renew_fields.page) <> 'pay.asp' and app_Renew_pages.order_by > 0 " &_
	"and app_renew_fields.field_type = 'RADIO' and app_renew_fields.fields_unrequire is not null " &_
	"order by app_renew_pages.order_by "
cSqlForm.AddParam(appRenewCd)
set rsFields = cSqlForm.Execute
do while not rsFields.eof
	if rsFields("FIELDS_UNREQUIRE")&"" <> "" then
		if session(rsFields("FIELD_NAME"))&"" = yearFunction(rsFields("DEFAULT_VALUE")&"") then
			arrTmp = split(rsFields("FIELDS_UNREQUIRE"),", ")
			for i = lbound(arrTmp) to ubound(arrTmp)
				if replace(arrTmp(i)," ","") <> "" then
					if objFields.Exists(clng(arrTmp(i))) then
						fieldName = objFields.Item(clng(arrTmp(i)))
						if objFieldsType(clng(arrTmp(i))) = "FILE" or objFieldsType(clng(arrTmp(i))) = "FILEBTN" then
							fieldName = "file_" & fieldName
						end if
						if objReq.Exists(fieldName) then
							if strReqDebug <> "" then
								strReqDebug = strReqDebug  & "<BR>"
							end if
							strReqDebug = strReqDebug & fieldName & " on " & objFieldsPage(clng(arrTmp(i))) & " is empty but NOT required because " & rsFields("FIELD_NAME") & " = " & session(rsFields("FIELD_NAME"))
							objReq.Remove(fieldName)
						end if
					end if
				end if
			next
		end if
	end if
	rsFields.MoveNext
loop
rsFields.close
set rsFields = nothing
set cSqlForm = nothing

if objReq.Count > 0 then
	arrReq = objReq.Items
	finalFieldCheck = arrReq(0)
end if

objReq.RemoveAll()
set objReq = nothing

end function

function finalCheck

dim returnPage, arrRetPage, fieldCheck, cSqlFinal

set cSqlFinal = new cRunSql
cSqlFinal.Conn = connOracle

arrRetPage = split(request.querystring("FROM_PAGE"),"/")

finalCheck = true

cSqlFinal.SqlStr = "select app_renew_fields.FIELD_NAME,app_renew_fields.FIELD_TYPE,app_renew_fields.PAGE from app_renew_fields, app_renew_pages " &_
	"where upper(app_renew_pages.page) = upper(app_renew_fields.page) and app_renew_pages.app_renew_type = app_renew_fields.app_renew_type " &_
	"and app_renew_fields.app_renew_type = ? and app_renew_fields.required = 'Y' and app_renew_fields.page is not null and lower(app_renew_fields.page) <> 'index.asp' " &_
	"and lower(app_renew_fields.page) <> 'pay.asp' and app_Renew_pages.order_by > 0 " &_
	"order by app_renew_pages.order_by "
cSqlFinal.AddParam(appRenewCd)
set rsFields = cSqlFinal.Execute

do while not rsFields.eof and finalCheck
	fieldCheck = ucase(rsFields("FIELD_NAME"))
	if rsFields("FIELD_TYPE") = "FILE" or rsFields("FIELD_TYPE") = "FILEBTN" then
		fieldCheck = "file_" & ucase(rsFields("FIELD_NAME"))
	end if
    if session(fieldCheck) = "" then
        finalCheck = false
		if request.querystring("FROM_PAGE") <> "" then
			returnPage = replace(request.querystring("FROM_PAGE"),arrRetPage(ubound(arrRetPage)),rsFields("PAGE"))
		else
			returnPage = rsFields("PAGE")
		end if
    end if
    rsFields.MoveNext
loop

rsFields.close
set rsFields = nothing

set cSqlFinal = nothing

if finalCheck then
else
	%><div class="error" id="divReqMiss">Required fields missing on a previous page.</div><button type="button" id="btnSubmit" name="btnSubmit" class="btn btn-primary" onClick="document.location='<%=aspxPage(returnPage)%>';">Continue</button><%
end if

end function

function addArray(arr,str)

dim idxNew

if isArray(arr) then
	idxNew = ubound(arr) + 1
	redim preserve arr(idxNew)
else
	idxNew = 0
	redim arr(idxNew)
end if
arr(idxNew) = str

addArray = arr

end function

function addGroup(group,str,delim)

if group&"" <> "" then
	group = group & delim
end if

group = group & str

addGroup = group

end function

sub groupRows(arr)

dim arrVal
if request.querystring("doPDF") = "1" then
	for i = lbound(arr) to ubound(arr)
%>
document.getElementById("<%=arr(i)%>").innerHTML = "<%=replace(session(arr(i))&"","\\","<br>")%>";
<%
	next
else
	arrVal = split(session(arr(0))&"","\\")
	if ubound(arrVal) > 0 then
		for i = lbound(arrVal)+1 to ubound(arrVal)
		%>
addGroupRowNew('<%=arr(0)%>',99);
		<%
		next
	end if

	for i = lbound(arr) to ubound(arr)
	%>
var arr<%=arr(i)%> = "<%=session(arr(i))&""%>".split("\\");
var fld<%=arr(i)%> = document.getElementsByName("<%=arr(i)%>");
	<%
	next
	for i = lbound(arr) to ubound(arr)
	%>
for (var idx<%=arr(i)%>=0;idx<%=arr(i)%><arr<%=arr(i)%>.length;idx<%=arr(i)%>++) {
	if ('value' in fld<%=arr(i)%>[idx<%=arr(i)%>]) {
		fld<%=arr(i)%>[idx<%=arr(i)%>].value = arr<%=arr(i)%>[idx<%=arr(i)%>];
	} else {
		setInnerText(fld<%=arr(i)%>[idx<%=arr(i)%>],arr<%=arr(i)%>[idx<%=arr(i)%>]);
	}
}
<%
	next
end if

end sub

sub groupTables(arr,dest)

dim arrVal
if request.querystring("doPDF") = "1" then
	for i = lbound(arr) to ubound(arr)
%>
//document.getElementById("<%=arr(i)%>").innerHTML = "<%=replace(session(arr(i))&"","\\","<br>")%>";
<%
	next
else
	arrVal = split(session(arr(0))&"","\\")
	if ubound(arrVal) > 0 then
		for i = lbound(arrVal)+1 to ubound(arrVal)
		%>
addGroupTableFull('<%=arr(0)%>',99,'<%=dest%>','');
		<%
		next
	end if

	for i = lbound(arr) to ubound(arr)
	%>
var arr<%=arr(i)%> = "<%=session(arr(i))&""%>".split("\\");
var fld<%=arr(i)%> = document.getElementsByName("<%=arr(i)%>");
	<%
	next
	for i = lbound(arr) to ubound(arr)
	%>
for (var idx<%=arr(i)%>=0;idx<%=arr(i)%><arr<%=arr(i)%>.length;idx<%=arr(i)%>++) {
	if ('value' in fld<%=arr(i)%>[idx<%=arr(i)%>]) {
		fld<%=arr(i)%>[idx<%=arr(i)%>].value = arr<%=arr(i)%>[idx<%=arr(i)%>];
	} else {
		setInnerText(fld<%=arr(i)%>[idx<%=arr(i)%>],arr<%=arr(i)%>[idx<%=arr(i)%>]);
	}
}
<%
	next
end if

end sub

sub popGroupFields(arr)

if request.querystring("doPDF") = "1" then
	for i = lbound(arr) to ubound(arr)
%>
document.getElementById("<%=arr(i)%>").innerHTML = "<%=replace(session(arr(i))&"","\\","<br>")%>";
<%
	next
else
	for i = lbound(arr) to ubound(arr)
	%>
var arr<%=arr(i)%> = "<%=session(arr(i))&""%>".split("\\");
var fld<%=arr(i)%> = document.getElementsByName("<%=arr(i)%>");
	<%
	next
	for i = lbound(arr) to ubound(arr)
	%>
for (var idx<%=arr(i)%>=0;idx<%=arr(i)%><arr<%=arr(i)%>.length;idx<%=arr(i)%>++) {
	if ('value' in fld<%=arr(i)%>[idx<%=arr(i)%>]) {
		fld<%=arr(i)%>[idx<%=arr(i)%>].value = arr<%=arr(i)%>[idx<%=arr(i)%>];
	} else {
		setInnerText(fld<%=arr(i)%>[idx<%=arr(i)%>],arr<%=arr(i)%>[idx<%=arr(i)%>]);
	}
}
<%
	next
end if

end sub

sub requiredExplainFields

dim questionPage

questionPage = ""

for idxField = lbound(arrFields,2) to ubound(arrFields,2)
	if arrFields(fieldNameIndex("PAGE"),idxField) = pageCd and arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "YN_EXPLAIN" then
		questionPage = 	arrFields(fieldNameIndex("QUESTION_PAGE"),idxField)
	end if
next

if instr(pageCd,"explain") or questionPage <> "" then
	if questionPage&"" = "" then
		questionPage = replace(pageCd,"explain.asp","") & ".asp"
	end if

	if loadDefaultsDelay = true then
		delayLog "onclick fields to require explanation",true
	else
		delayLog "onclick fields to require explanation",false
		for idxField = lbound(arrFields,2) to ubound(arrFields,2)
			if arrFields(fieldNameIndex("PAGE"),idxField) = questionPage and arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "YN" then
				if session(arrFields(fieldNameIndex("FIELD_NAME"),idxField))&"" = "0" or lcase(session(arrFields(fieldNameIndex("FIELD_NAME"),idxField))) = "yes" then
					if arrFields(fieldNameIndex("YN_REQUIRE_FILE"),idxField) = "Y" and arrFields(fieldNameIndex("YN_REQUIRE_EXP"),idxField) = "Y" then
						if session("file_" & arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "" or session("exp_" & arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "" then
							cForm.AddErrorMsg "Explanation and File Upload for ""Yes"" response to question '" & URLDecode(arrFields(fieldNameIndex("FIELD_DESCRIPTION"),idxField)) & "'  is required"
						end if
					elseif arrFields(fieldNameIndex("YN_REQUIRE_FILE"),idxField) = "Y" then
						if session("file_" & arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "" then
							cForm.AddErrorMsg "File Upload for ""Yes"" response to question '" & URLDecode(arrFields(fieldNameIndex("FIELD_DESCRIPTION"),idxField)) & "'  is required"
						end if
					elseif arrFields(fieldNameIndex("YN_REQUIRE_EXP"),idxField) = "Y" then
						if session("exp_" & arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "" then
							cForm.AddErrorMsg "Explanation for ""Yes"" response to question '" & URLDecode(arrFields(fieldNameIndex("FIELD_DESCRIPTION"),idxField)) & "'  is required"
						end if
					else
						if session("exp_" & arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "" and session("file_" & arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "" then
							cForm.AddErrorMsg "Explanation or File Upload for ""Yes"" response to question '" & URLDecode(arrFields(fieldNameIndex("FIELD_DESCRIPTION"),idxField)) & "'  is required"
						end if
					end if
				end if
			end if
		next
	end if
end if

for idxField = lbound(arrFields,2) to ubound(arrFields,2)
	if arrFields(fieldNameIndex("PAGE"),idxField) = pageCd and arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "EXPLANATION" and arrFields(fieldNameIndex("REQUIRED"),idxField) = "Y" then
		if session("exp_" & arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "" and not isArray(session("ARR_FILES_" & arrFields(fieldNameIndex("FIELD_NAME"),idxField))) then
			cForm.AddErrorMsg "Explanation or File Upload for question '" & URLDecode(arrFields(fieldNameIndex("FIELD_DESCRIPTION"),idxField)) & "' is required"
		end if
	end if
next

for idxField = lbound(arrFields,2) to ubound(arrFields,2)
	if arrFields(fieldNameIndex("PAGE"),idxField) = pageCd and arrFields(fieldNameIndex("REQUIRE_EXPLANATION"),idxField) = "Y" then
		if session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = cstr(arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField)&"") then
			if arrFields(fieldNameIndex("YN_REQUIRE_FILE"),idxField)&"" = "Y" and arrFields(fieldNameIndex("YN_REQUIRE_EXP"),idxField)&"" = "Y" then
				if session(arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "" or session("exp_" & arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "" then
					cForm.AddErrorMsg "Explanation and File Upload for question '" & URLDecode(arrFields(fieldNameIndex("FIELD_DESCRIPTION"),idxField)) & "' is required"
				end if
			elseif arrFields(fieldNameIndex("YN_REQUIRE_FILE"),idxField)&"" = "Y" then
				if not isArray(session("ARR_FILES_" & arrFields(fieldNameIndex("FIELD_NAME"),idxField))) then
					cForm.AddErrorMsg "File Upload for question '" & URLDecode(arrFields(fieldNameIndex("FIELD_DESCRIPTION"),idxField)) & "' is required"
				end if
			elseif arrFields(fieldNameIndex("YN_REQUIRE_EXP"),idxField)&"" = "Y" then
				if session("exp_" & arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "" then
					cForm.AddErrorMsg "Explanation for question '" & URLDecode(arrFields(fieldNameIndex("FIELD_DESCRIPTION"),idxField)) & "' is required"
				end if
			else
				if session("exp_" & arrFields(fieldNameIndex("FIELD_NAME"),idxField)) = "" and not isArray(session("ARR_FILES_" & arrFields(fieldNameIndex("FIELD_NAME"),idxField))) then
					cForm.AddErrorMsg "Explanation or File Upload for question '" & URLDecode(arrFields(fieldNameIndex("FIELD_DESCRIPTION"),idxField)) & "' is required"
				end if
			end if
		end if
	end if
next

end sub

sub setRequiredExplainFields

dim questionPage

questionPage = ""

for idxField = lbound(arrFields,2) to ubound(arrFields,2)
	if arrFields(fieldNameIndex("PAGE"),idxField) = pageCd and arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "YN_EXPLAIN" then
		questionPage = 	arrFields(fieldNameIndex("QUESTION_PAGE"),idxField)
	end if
next

if instr(pageCd,"explain") or questionPage <> "" then
	if questionPage&"" = "" then
		questionPage = replace(pageCd,"explain.asp","") & ".asp"
	end if

	for idxField = lbound(arrFields,2) to ubound(arrFields,2)
		if arrFields(fieldNameIndex("PAGE"),idxField) = questionPage and arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "YN_EXPLAIN" then
			if session(arrFields(fieldNameIndex("FIELD_NAME"),idxField))&"" = "0" or lcase(session(arrFields(fieldNameIndex("FIELD_NAME"),idxField))) = "yes" then
				if arrFields(fieldNameIndex("YN_REQUIRE_EXP"),idxField) = "Y" then
					cForm.AddRequired "exp_" & arrFields(fieldNameIndex("FIELD_NAME"),idxField),"Explanation for question '" & URLDecode(arrFields(fieldNameIndex("FIELD_DESCRIPTION"),idxField)) & "'"
				end if
			end if
		end if
	next
end if

end sub

function getNextPage

	dim nextPg, curPgIdx, idxPg, idxSkip
	dim rsExplainFields, rsQuestionFields
	dim skipServerSide

	skipServerSide = false

	if cRenew.ConfigOption("SKIP_SERVER_SIDE") = "1" then
		skipServerSide = true
		if loadDefaultsDelay = true then
			skipServerSide = false
			delayLog "server-side skip",true
		end if
	end if

	if skipServerSide then
		delayLog "server-side skip",false
		combineFields(false)
		curPgIdx = 0

		nextPg = ""

		if isArray(arrPages) then	

			for i = lbound(arrPages,2) to ubound(arrPages,2)
				if arrPages(0,i) = pageName then
					curPgIdx = i
				end if
			next

			for idxPg = curPgIdx to ubound(arrPages,2)
				if idxPg < ubound(arrPages,2) and nextPg = "" then
					nextPg = arrPages(0,idxPg+1)
					for idxField = lbound(arrFields,2) to ubound(arrFields,2)
						if arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "RADIO" and arrFields(fieldNameIndex("PAGES_SKIP"),idxField)&"" <> "" then
							doSkip = false
							if arrFields(fieldNameIndex("RADIO_NOT_EQUAL"),idxField) = "Y" then
								if session(arrFields(fieldNameIndex("FIELD_NAME"),idxField))&"" <> yearFunction(arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField)&"") then
									doSkip = true
								end if
							else
								if session(arrFields(fieldNameIndex("FIELD_NAME"),idxField))&"" = yearFunction(arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField)&"") then
									doSkip = true
								end if
							end if
							if doSkip then
								arrSkip = split(arrFields(fieldNameIndex("PAGES_SKIP"),idxField),", ")
								for idxSkip = lbound(arrSkip) to ubound(arrSkip)
									if ucase(nextPg) = ucase(arrSkip(idxSkip)) then
										if arrFields(fieldNameIndex("RADIO_NOT_EQUAL"),idxField) = "Y" then
											formLog("skip " & ucase(arrSkip(idxSkip)) & " because " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " <> " & yearFunction(arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField)))
										else								
											formLog("skip " & ucase(arrSkip(idxSkip)) & " because " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " = " & yearFunction(arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField)))
										end if
										nextPg = ""
									end if
								next
							end if
						end if
					next

					for idxField = lbound(arrFields,2) to ubound(arrFields,2)
						if arrFields(fieldNameIndex("PAGE"),idxField) = pageCd and arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "YN_EXPLAIN" and arrFields(fieldNameIndex("ROW_NUM"),idxField)&"" <> "" and arrFields(fieldNameIndex("COL_NUM"),idxField)&"" <> "" then
							questionPage = arrFields(fieldNameIndex("QUESTION_PAGE"),idxField)
							questionPageAttributes = arrFields(fieldNameIndex("ATTRIBUTES"),idxField)
							if questionPage&"" = "" then
							  if right(pageCd,11) = "explain.asp" then
								questionPage = replace(pageCd,"explain.asp","") & ".asp"
							  end if
							end if

							  if questionPage&"" <> "" then
								  yesCount = 0
								  set cSqlForm = new cRunSql
								  cSqlForm.Conn = connOracle
								  cSqlForm.SqlStr = "select FIELD_NAME from app_renew_fields where app_renew_type = ? and lower(page) = ? and field_type = 'YN' order by row_num, col_num"
								  cSqlForm.AddParam(appRenewCd)
								  cSqlForm.AddParam(questionPage)
								  set rsQuestionFields = cSqlForm.Execute
								  set cSqlForm = nothing
								  do while not rsQuestionFields.eof
									if session(rsQuestionFields("FIELD_NAME")) = "Yes" then
									  yesCount = yesCount + 1
									end if
									rsQuestionFields.MoveNext
								  loop
								  rsQuestionFields.close
								  set rsQuestionFields = nothing
								  set cSqlForm = nothing
      
								  noCount = 0
								  set cSqlForm = new cRunSql
								  cSqlForm.Conn = connOracle
								  cSqlForm.SqlStr = "select FIELD_NAME from app_renew_fields where app_renew_type = ? and lower(page) = ? and field_type = 'NY' order by row_num, col_num"
								  cSqlForm.AddParam(appRenewCd)
								  cSqlForm.AddParam(questionPage)
								  set rsQuestionFields = cSqlForm.Execute
								  set cSqlForm = nothing
								  do while not rsQuestionFields.eof
									if session(rsQuestionFields("FIELD_NAME")) = "No" then
									  noCount = noCount + 1
									end if
									rsQuestionFields.MoveNext
								  loop
								  rsQuestionFields.close
								  set rsQuestionFields = nothing
								  set cSqlForm = nothing

								if yesCount = 0 and noCount = 0 then
									if instr(arrFields(fieldNameIndex("ATTRIBUTES"),idxField)&"","readonly") = 0 then
									else
										nextPg = ""
									end if
								end if
							  end if
							end if
						next
				end if
			next
		end if

		getNextPage = nextPg
	else
		if isArray(arrPages) then	
			for i = lbound(arrPages,2) to ubound(arrPages,2)
				if arrPages(0,i) = pageName and i < ubound(arrPages,2) then
					getNextPage = arrPages(0,i+1)
				end if
			next
		end if	
	end if

end function 

function getNextPageByID(pageID)
	dim nextPg

	if isArray(arrPages) then	
		for i = lbound(arrPages,2) to ubound(arrPages,2)
			if clng(arrPages(4,i)) = clng(pageID) and i < ubound(arrPages,2) then
				getNextPageByID = arrPages(0,i+1)
			end if
		next
	end if
end function 

function getPrevPage
	dim retVal
	retVal = ""
	if isArray(arrPages) then
		for i = ubound(arrPages,2) to lbound(arrPages,2) step -1
			if arrPages(0,i) = pageName and i > 0 then
				retVal = arrPages(0,i-1)
			end if
		next
	end if
	getPrevPage = retVal
end function 

sub showDebug

dim debugForm, debugItem
dim idxRelate, arrRelate, idxRelate2, isFormScanner, clientIP

clientIP = request.servervariables("HTTP_X_FORWARDED_FOR")
if clientIP & " " = " " then
	clientIP = Request.ServerVariables("REMOTE_ADDR")
end if

isFormScanner = false

if clientIP & " " <> " " then
	if clientIP = "64.235.150.121" then
		isFormScanner = true
	else
		'is IP in scanner range of 64.235.153.0/24  
		isFormScanner = compareIpsForm(split(clientIP,"."), Array(64,235,153,0), Array(64,235,153,255))
	end if
end if

if isDev or request.querystring("action") = "loadDebug" or officeIP then
	if request("doPDF") = "1" then
	else
		debugForm = true
	end if
end if

if isFormScanner then
	debugForm = false
end if

if debugForm then
	%><div class="debug"><br><%
		%><div class='error alert alert-info'><%
			%>form code is <%=session("APP_RENEW_CD")%><br><%
			%>form record type is <%=formRecType%><br><%
			%>form id is <%=session("APP_RENEW_ID")%><br><%
			%>membership id is <%=session("MEMBERSHIP_ID")%><br><%
			%>session timeout set to <%=sessionTimeout%><br><%
			%>system date/time type set to <%=sysNow%><br><%
			%>display type set to <%=displayType%><%
			if formStart&"" <> "" then
				%><br>form start set to <%=formStart%><%
			end if
			if formEnd&"" <> "" then
				%><br>form end set to <%=formEnd%><%
			end if
			%><br>IP address is <%=clientIP%><%
			if redis then
				%><br>redisSessionID = <%=session("redisSessionID")%><%
				%><br>redisTimeout = <%=session("redisTimeout")%><%
				%><br>redisExpiration = <%=session("redisExpiration")%><%
				If Request.Cookies("sessionIDwww") Is Nothing Then
					Response.Write "<br>redis Cookie does not exist"
				Else
					Response.Write "<br>Cookie sessionIDwww exists with value: " & Request.Cookies("sessionIDwww")
				End If
			End If
		%></div><%
	if request("doEmail") = "1" then
	else
		if redis then
			redisDebugKey = session("redissessionid")
		else
			Set TypeLib = CreateObject("Scriptlet.TypeLib")
			redisDebugKey = replace(Mid(TypeLib.Guid, 2, 36),"-","")
			set TypeLib = nothing
			set redisDebug = Server.CreateObject("CacheManager")
			redisDebug.Init redisDebugKey
			for each item in session.contents
				redisDebug(item) = session(item)
			next
			set redisDebug = nothing
		end if
		%><br><a href="<%=aspxPage(request.ServerVariables("SCRIPT_NAME"))%>?action=loadDebug&formID=<%=session("APP_RENEW_ID")%>">Debug Link</a><br><br><%
		%><input type="checkbox" onClick="toggleEditLinks(this);" id="chkEditLinks" name="chkEditLinks"> <label for="chkEditLinks">Show Edit Links</label><br><%
		%><input type="checkbox" onClick="toggleDebug(this,'<%=pageCd%>','<%=cRenew.PageOption("save_record")%>','<%=cRenew.PageOption("save_pending")%>','<%=cRenew.PageOption("create_pdf")%>','<%=cRenew.PageOption("end_session")%>','<%=redisDebugKey%>');" id="chkDebug" name="chkDebug"> <label for="chkDebug">Debug</label><br><%
		%><div id="divDebug" style="display:none">Fetching debug...</div><%
		%><input type="checkbox" onClick="toggleFormLog(this,'<%=session("APP_RENEW_ID")%>');" id="chkFormLog" name="chkFormLog"> <label for="chkFormLog">Form Log</label><br><%
		%><div id="divFormLog" style="display:none">Fetching form log...</div><%
	end if
	%></div><%
else
	if officeIP then
		%><!--Debug Link - https://<%=request.ServerVariables("SERVER_NAME")%><%=request.ServerVariables("SCRIPT_NAME")%>?action=loadDebug&formID=<%=session("APP_RENEW_ID")%>--><%
	end if
end if

end sub

function aspxPage(page)
	if request("ajax") = "1" then
		aspxPage = replace(page,".asp",".aspx")
	else
		aspxPage = page
	end if
end function

sub goNext

if request("doPDF") <> "1" then

	nextPage = getNextPage

	if nextPage&"" <> "" then
		%><script language="Javascript">document.location = "<%=aspxPage(nextPage)%>";</script><%
		closeAll()
		response.end
	else
		cRenew.EmailPages(isFormScanner)
		if session("TASK_FORMS") then
			%><script language="Javascript">document.location = "/www/task_forms.asp";</script><%
			closeAll()
			response.end
		end if
	end if
end if

end sub


sub batchTransaction

if request("action") = "loadDebug" or request.form("do") = "exit" or request.querystring("preview") = "1" or request.querystring("doPDF") = "1" or request.querystring("doEmail") = "1" then
	exit sub
end if

dim arrProductID
dim BATCH_ID

BATCH_NUM = getBatch_NUM
session("BATCH_NUM") = BATCH_NUM
updateBatchInfo BATCH_NUM,session("CARD_AMOUNT")
BATCH_ID = getBatchID(BATCH_NUM)
updateBatch_Member_Finder BATCH_ID,session("MEMBERSHIP_ID")

arrProductID = split(session("PRODUCT_ID")&"",",")
for i = lbound(arrProductID) to ubound(arrProductID)
	set cSqlForm = new cRunSql
	cSqlForm.conn = connOracle
	cSqlForm.SqlStr = "insert into TRANSACTION_LOG (TRAN_SEQ_ID,BATCH_ID,TRANS_TYPE,TRANS_CD,TRANS_SOURCE,TRANS_STATUS_CD,RESULT_CD, " &_
	"AMOUNT,MEMBERSHIP_ID,PAYMENT_TYPE_CD,CLIENT_IP, CREATE_DT,CK_NAME,CARDNUM,CK_DATE,TXN_ID,PPREF,FULL_PAYPAL_RESPONSE, " &_
	"FIRST_NAME,LAST_NAME,ADDRESS_LINE_1,CITY,STATE,ZIP,BATCH_MEMBER_ID,APP_RENEW_ID) VALUES " &_
	"(?,?,'PAY',(select product_desc_brief from " & prodTbl & " where product_id = ?),'WEB','APPROVED','0',(select cost_amt from " & prodTbl & " where product_id = ?),?," &_
	"?,?, sysdate,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
	cSqlForm.AddParam(session("CARD_CONFIRMATION"))
	cSqlForm.AddParam(BATCH_ID)
	cSqlForm.AddParam(arrProductID(i))
	cSqlForm.AddParam(arrProductID(i))
	cSqlForm.AddParam(Session("MEMBERSHIP_ID"))
	if session("CHECK_NUM")&"" <> "" then
		cSqlForm.AddParam("K")
	elseif session("PAYMENT_TYPE")&"" <> "" then
		cSqlForm.AddParam(session("PAYMENT_TYPE"))
	else
		cSqlForm.AddParam("C")
	end if
	cSqlForm.AddParam(Request.ServerVariables("REMOTE_ADDR"))
	if session("CHECK_NAME")&"" <> "" then
		cSqlForm.AddParam(session("CHECK_NAME"))
	else
		cSqlForm.AddParam(Session("CARD_FIRST_NAME") & " " & session("CARD_LAST_NAME"))
	end if
	if session("CHECK_NUM")&"" <> "" then
		cSqlForm.AddParam(session("CHECK_NUM"))
	else
		cSqlForm.AddParam(session("CARD_NUM"))
	end if
	if session("CHECK_DATE")&"" <> "" then
		cSqlForm.AddParam(session("CHECK_DATE"))
	else
		cSqlForm.AddParam(sysDate)
	end if
	cSqlForm.AddParam(session("PNREF"))
	cSqlForm.AddParam(session("PPREF"))
	cSqlForm.AddParam(session("TRANS_RESPONSE"))
	cSqlForm.AddParam(session("CARD_FIRST_NAME"))
	cSqlForm.AddParam(session("CARD_LAST_NAME"))
	cSqlForm.AddParam(session("CARD_ADDRESS_LINE_1"))
	cSqlForm.AddParam(session("CARD_CITY"))
	cSqlForm.AddParam(session("CARD_STATE"))
	cSqlForm.AddParam(session("CARD_ZIP"))
	cSqlForm.AddParam(session("BATCH_MEMBER_ID"))
	cSqlForm.AddParam(session("APP_RENEW_ID"))
	cSqlForm.Execute
	set cSqlForm = nothing
next
updateBatchInfo BATCH_NUM,session("CARD_AMOUNT")
%>
<%end sub%>

<%sub batchTransactionPending%>
<%
dim arrProductID

dim BATCH_ID

BATCH_NUM = getBatch_NUM

session("BATCH_NUM") = BATCH_NUM
updateBatchInfo BATCH_NUM,session("CARD_AMOUNT")
BATCH_ID = getBatchID(BATCH_NUM)
updateBatch_Member_Finder BATCH_ID,session("MEMBERSHIP_ID")

arrProductID = split(session("PRODUCT_ID")&"",",")
for i = lbound(arrProductID) to ubound(arrProductID)
	set cSqlForm = new cRunSql
	cSqlForm.conn = connOracle
	cSqlForm.SqlStr = "insert into TRANSACTION_LOG (TRAN_SEQ_ID,BATCH_ID,TRANS_TYPE,TRANS_CD,TRANS_SOURCE,TRANS_STATUS_CD,RESULT_CD, " &_
	"AMOUNT,MEMBERSHIP_ID,PAYMENT_TYPE_CD,CLIENT_IP, CREATE_DT,CK_NAME,CARDNUM,CK_DATE,TXN_ID,PPREF,FULL_PAYPAL_RESPONSE, " &_
	"FIRST_NAME,LAST_NAME,ADDRESS_LINE_1,CITY,STATE,ZIP) VALUES " &_
	"(?,?,'PAY',(select product_desc_brief from " & prodTbl & " where product_id = ?),'WEB','PENDING','0',(select cost_amt from " & prodTbl & " where product_id = ?),?," &_
	"?,?, sysdate,?,?,?,?,?,?,?,?,?,?,?,?)"
	cSqlForm.AddParam(session("CARD_CONFIRMATION"))
	cSqlForm.AddParam(BATCH_ID)
	cSqlForm.AddParam(arrProductID(i))
	cSqlForm.AddParam(arrProductID(i))
	cSqlForm.AddParam(Session("MEMBERSHIP_ID"))
	if session("CHECK_NUM")&"" <> "" then
		cSqlForm.AddParam("K")
	else
		cSqlForm.AddParam("C")
	end if
	cSqlForm.AddParam(Request.ServerVariables("REMOTE_ADDR"))
	if session("CHECK_NAME")&"" <> "" then
		cSqlForm.AddParam(session("CHECK_NAME"))
	else
		cSqlForm.AddParam(Session("CARD_FIRST_NAME") & " " & session("CARD_LAST_NAME"))
	end if
	if session("CHECK_NUM")&"" <> "" then
		cSqlForm.AddParam(session("CHECK_NUM"))
	else
		cSqlForm.AddParam(session("CARD_NUM"))
	end if
	if session("CHECK_DATE")&"" <> "" then
		cSqlForm.AddParam(session("CHECK_DATE"))
	else
		cSqlForm.AddParam(sysDate)
	end if
	cSqlForm.AddParam(session("PNREF"))
	cSqlForm.AddParam(session("PPREF"))
	cSqlForm.AddParam(session("TRANS_RESPONSE"))
	cSqlForm.AddParam(session("CARD_FIRST_NAME"))
	cSqlForm.AddParam(session("CARD_LAST_NAME"))
	cSqlForm.AddParam(session("CARD_ADDRESS_LINE_1"))
	cSqlForm.AddParam(session("CARD_CITY"))
	cSqlForm.AddParam(session("CARD_STATE"))
	cSqlForm.AddParam(session("CARD_ZIP"))
	cSqlForm.Execute
	set cSqlForm = nothing
next
updateBatchInfo BATCH_NUM,session("CARD_AMOUNT")
%>
<%end sub%>

<%function finishWVEPay%>
<%
dim PaymentDoc, pDoc, card, amt, convFee

formLog("begin retrieving WV ePay results")
set PaymentDoc = RetrieveWVEpay(request.querystring("ID"))
formLog("result = " & PaymentDoc.xml)
formLog("end retrieving WV ePay results")

if instr(PaymentDoc.xml,"ProcessState") then
	set PDoc = server.CreateObject("Microsoft.XMLDOM")
	PDoc.loadxml(PaymentDoc.xml)

	oStatus = PDoc.selectSingleNode("//PayRespBlk/Status").Attributes.GetNamedItem("ProcessState").Text
	if oStatus = "Failed" then
		oErr = PDoc.selectSingleNode("//PayRespBlk/Status").Attributes.GetNamedItem("StatusDesc").Text
	else
		if instr(PaymentDoc.xml,"ACH Header") then
			card = PDoc.selectSingleNode("//Payment/ACH").Attributes.GetNamedItem("Account").Text
			session("CARD_NAME") = PDoc.selectSingleNode("//Payment/ACH").Attributes.GetNamedItem("Name").Text
			session("CARD_FIRST_NAME") = PDoc.selectSingleNode("//Payment/ACH").Attributes.GetNamedItem("Name").Text
			session("PAYMENT_TYPE") = "A"
		else
			card = PDoc.selectSingleNode("//Payment/Card").Attributes.GetNamedItem("Number").Text
			session("CARD_NAME") = PDoc.selectSingleNode("//Payment/Card").Attributes.GetNamedItem("Name").Text
			session("CARD_FIRST_NAME") = PDoc.selectSingleNode("//Payment/Card").Attributes.GetNamedItem("Name").Text
		end if
		amt = PDoc.selectSingleNode("//CurrencyAmount").Attributes.GetNamedItem("Amount").Text
		on error resume next
		convFee = PDoc.selectSingleNode("//FromAgency/Agency").Attributes.GetNamedItem("ConvienceFee").Text
		on error goto 0
		session("CARD_NUM_RECEIPT") = "****" & right(card,4)
		session("CARD_AMOUNT") = amt
		session("CARD_AMOUNT_EPAY") = amt
		session("CONV_FEE") = convFee
		session("PNREF") = PDoc.selectSingleNode("//TransRefBlk/TransId").Attributes.GetNamedItem("IOTPTransId").Text
		if instr(session("PNREF"),"@") then
			session("PNREF_CONF") = split(session("PNREF"),"@")(0)
		end if
		finishWVEPay = true
		paymentStatus "Complete"
	end if
	set pDoc = nothing
end if

set PaymentDoc = nothing
    
if oStatus = "Failed" then
	paymentStatus "Failed"
	%>
	<script language="Javascript">
		alert("Payment failed - <%=oErr%>");
	</script>
	<%
	finishWVEPay = false
elseif oStatus = "CompletedOk" then
else
	%>
	<script language="Javascript">
		alert("Payment failed");
	</script>
	<%
	finishWVEPay = false
end if
%>
<%end function%>

<%function RetrieveWVEpay(UID)
    Dim oXMLHTTP,XmlDoc, wvStr, epayKey
    
	if epayURL&"" = "" then
		if isDev then
			epayURL = "https://epay2dev.wvsto.com"
		else
			epayURL = "https://epay2.wvsto.com"
		end if
	end if

	ePayKey = getCodeValueDesc("APPLICATION_DB","WV_EPAY_KEY")

	if cRenew.ConfigOption("epay_key")&"" <> "" Then
		ePayKey = cRenew.ConfigOption("epay_key")
	end if
    
    set XmlDoc = server.CreateObject("Microsoft.XMLDOM")

	if ePayKey&"" <> "" and epayKey <> "WV_EPAY_KEY" then
		epayURL = epayURL & "/ePay/ePayService.asmx/RetrieveCached"
	else
		epayURL = epayURL & "/ePay/ePayService.asmx/RetrieveCachedData"
	end if
	formLog("posting to " & epayURL)    
    Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
    oXMLHTTP.open "POST",epayURL,False
    oXMLHTTP.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
	wvStr = "GroupID=" & UID
	if ePayKey&"" <> "" and epayKey <> "WV_EPAY_KEY" then
	    wvStr = wvStr & "&Key=" & ePayKey
	end if
	formLog(wvStr)
    oXMLHTTP.send wvStr
    strResponse = oXMLHTTP.responseText
    set  oXMLHTTP = nothing  

    xmlDoc.LoadXML(strResponse)

    set RetrieveWVEpay = XmlDoc
	set xmlDOC = nothing
    
end function
%>

<%sub kanpay%>
<%
	session("KANPAY_NEXT_PAGE") = getNextPage
	session("KANPAY_PAYMENT_PAGE") = request.servervariables("SCRIPT_NAME")
	redirectForm("/www/kanpay/")
%>
<%end sub%>

<%sub wvEPay%>
<%
Dim XMLString, XmlDoc, objNodeList, strValidationStatus, strValidationUID, strResponse, oXMLHTTP
dim strXML

set XMLString = GenerateXML(Amount)
strXML = XMLString.xml
if session("payment") = "ACH" then
	strXML = replace(strXML,"</XMLPayment>","<ACH><BankName></BankName><RoutingNumber></RoutingNumber><AccountNumber></AccountNumber><AccountType></AccountType></ACH><PaymentOptions><Payment Type=""Check"" Option=""4"" ImageURL=""/ePayCustomPage/Images/ACHImage.gif"" Enabled=""True"" /></PaymentOptions></XMLPayment>")
else
	strXML = replace(strXML,"</XMLPayment>","<CreditCards><Cards Type=""American Express"" CardID=""AMER"" /><Cards Type=""Discover"" CardID=""DISC"" /><Cards Type=""Master Card"" CardID=""MAST"" /><Cards Type=""Visa"" CardID=""VISA"" /></CreditCards><PaymentOptions><Payment Type=""Credit Card"" Option=""1"" ImageURL=""/ePayCustomPage/Images/CreditCardImage.gif"" Enabled=""True""/></PaymentOptions> </XMLPayment>")
end if

formLog("begin retrieving WV ePay ID")
Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
oXMLHTTP.open "POST",epayURL & "/ePay/ePayService.asmx/XMLPayment",False
oXMLHTTP.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
formLog("request = " & strXML)
oXMLHTTP.send "XMLNd=" & strXML
strResponse = oXMLHTTP.responseText
formLog("response = " & strResponse)
set  oXMLHTTP = nothing  
formLog("end retrieving WV ePay ID")

strStart = instr(strResponse,"Status=""")
strValidationStatus = right(strResponse,(len(strResponse)-strStart)-7)
strEnd = instr(strValidationStatus,"""")
strValidationStatus = left(strValidationStatus,strEnd-1)

'Make sure eveything is valid    
if strValidationStatus = "Valid" then
	formLog("redirecting to WV ePay")
	strStart = instr(strResponse," UID=""")
	strValidationUID = right(strResponse,(len(strResponse)-strStart)-5)
	strEnd = instr(strValidationUID,"""")
	strValidationUID = left(strValidationUID,strEnd-1)
	session("WV_EPAY_INIT_ID") = strValidationUID

	set cSql = new cRunSql
	cSql.Conn = connOracle
	cSql.SqlStr = "delete from app_renew_detail where upper(field_name) = 'WV_EPAY_INIT_ID' and app_renew_id = ?"
	cSql.AddParam(session("APP_RENEW_ID"))
	cSql.Execute
	set cSql = nothing

	set cSql = new cRunSql
	cSql.Conn = connOracle
	cSql.SqlStr = "insert into app_renew_detail (field_name,data,app_renew_id) values ('WV_EPAY_INIT_ID',?,?)"
	cSql.AddParam(strValidationUID)
	cSql.AddParam(session("APP_RENEW_ID"))
	cSql.Execute
	set cSql = nothing

	paymentStatus "Pending"
	'Redirect the pages to the ePay system
	formLog("redirecting to " & epayURL & "/ePayCustomPage/ePay.aspx?Page=" & strValidationUID)
	if isDev and officeIP then
	%>
	<script>
		alert("You are being redirected to the third party payment site for WV ePay.\n\nErrors on that site are not software issues.  Development staff here cannot assist you.\n\nYou must contact the client and/or processor.  Only send a task on this end when specific instructions for software (not config) changes are available.");
		document.location="<%=epayURL & "/ePayCustomPage/ePay.aspx?Page=" & strValidationUID%>";
		</script>
	<%
		closeAll()
	else
		redirectForm(epayURL & "/ePayCustomPage/ePay.aspx?Page=" & strValidationUID)
	end if
	response.end
else
	formLog("WV ePay ID error")
	response.write strValidationStatus
	if isDev Then
		response.write "<BR><BR>url = " & epayURL & "/ePay/ePayService.asmx/XMLPayment<BR>"
		response.write "request = <textarea>" & strXML & "</textarea><BR>"
		response.write "response = <textarea>" & strResponse & "</textarea><BR>"
		response.write "<BR><BR><strong>This means we received a bad response back from the processor.  This is NOT a software issue on this end.  Review the setup and contact the processor with the request and response from above.</strong>"
		if isDev and officeIP then
		%>
		<script>alert("This means we received a bad response back from the processor.  This is NOT a software issue on this end.  Review the setup and contact the processor with the request and response from the screen.")
		<%
		end if
	end if
end if
set xmlString = nothing
response.end
%>
<%end sub%>

<%
private function GenerateXML(Amount)
Dim XmlDoc, ePayID, cCV, arrEP(1,10)
Dim XMLPayment, XMLAgency, XMLOptions, XMLBilling, XMLBillingOption
Dim XMLAttr, epayKey

set XmlDoc = server.CreateObject("Microsoft.XMLDOM")

set XMLPayment = XmlDoc.createElement("XMLPayment")
           
set XMLOptions = XmlDoc.createElement("ePayOptions")
            
set XMLAttr = XmlDoc.CreateAttribute("PreQuestionare")
XMLAttr.Value = "False"
XMLOptions.setAttributeNode XMLAttr

set XMLAttr = XmlDoc.CreateAttribute("PostQuestionare")
XMLAttr.Value = "False"
XMLOptions.setAttributeNode XMLAttr

set XMLAttr = XmlDoc.CreateAttribute("QuickInsert")
XMLAttr.Value = "False"
XMLOptions.setAttributeNode XMLAttr	 

if Application("EPAY_CSS_URL")&"" <> "" then
	set XMLAttr = XmlDoc.CreateAttribute("CSS")
	XMLAttr.Value = Application("EPAY_CSS_URL")
	XMLOptions.setAttributeNode XMLAttr	 
end if
	    
XMLPayment.AppendChild(XMLOptions)

set XMLBilling = XmlDoc.createElement("Billing")
XMLPayment.AppendChild(XMLBilling)

arrEP(0,0) = "FirstName"
arrEP(1,0) = "FIRST_NAME"
arrEP(0,1) = "LastName"
arrEP(1,1) = "LAST_NAME"
arrEP(0,2) = "Address1"
arrEP(1,2) = "ADDRESS_LINE_1"
arrEP(0,3) = "Address2"
arrEP(1,3) = "ADDRESS_LINE_2"
arrEP(0,4) = "City"
arrEP(1,4) = "CITY"
arrEP(0,5) = "Country"
arrEP(1,5) = "COUNTRY"
arrEP(0,6) = "State"
arrEP(1,6) = "STATE"
arrEP(0,7) = "Zip"
arrEP(1,7) = "ZIP"
arrEP(0,8) = "Phone"
arrEP(1,8) = "PHONE"
arrEP(0,9) = "Email"
arrEP(1,9) = "EMAIL"
arrEP(0,10) = "CompanyName"
arrEP(1,10) = "COMPANY"
	    
for i = lbound(arrEP,2) to ubound(arrEP,2)
	set XMLBillingOption = XmlDoc.createElement(arrEP(0,i))
	xmlBillingOption.Text = replace(session(arrEP(1,i))&"","&","and")
	XMLBilling.AppendChild(XMLBillingOption)
	set XMLBillingOption = nothing
next

set XMLAgency = XmlDoc.createElement("Agency")

set cCV = new cCodeValueO
cCV.LoadFromCodeClassAndValue "APP_RENEW_TYPE_CD",appRenewCd
formPath = cCV.CodeValueLongDesc

if cRenew.ConfigOption("epay_id")&"" <> "" then
	epayID = cRenew.ConfigOption("epay_id")
else
	ePayID = cCV.OptionValue1
end if

set cCV = nothing

ePayKey = getCodeValueDesc("APPLICATION_DB","WV_EPAY_KEY")

if cRenew.ConfigOption("epay_key")&"" <> "" Then
	ePayKey = cRenew.ConfigOption("epay_key")
end if
    
if ePayID&"" = "" then
%>
<div class='error alert alert-warning'>ePay Configuration Incomplete</div>
<%
	response.end
end if

set XMLAttr = XmlDoc.CreateAttribute("ID")
XMLAttr.Value = ePayID
XMLAgency.setAttributeNode XMLAttr

if ePayKey&"" <> "" and epayKey <> "WV_EPAY_KEY" then
	set XMLAttr = XmlDoc.CreateAttribute("Key")
	XMLAttr.Value = epayKey
	XMLAgency.setAttributeNode XMLAttr
end if

'Add the amount to charge
set XMLAttr = XmlDoc.CreateAttribute("Amount")
XMLAttr.Value = session("CARD_AMOUNT") 
XMLAgency.setAttributeNode XMLAttr

'Change the type when going from test to production
set XMLAttr = XmlDoc.CreateAttribute("Type")
if isDev then
	XMLAttr.Value = "TestOK"
	'XMLAttr.Value = "TestFailer"
else
	XMLAttr.Value = "Production"
end if
'Production
'TestOK
'TestFailer
'XMLAttr.Value = "Production"
'XMLAttr.Value = "TestFailer"

XMLAgency.setAttributeNode XMLAttr
'This is the page it is coming from
set XMLAttr = XmlDoc.CreateAttribute("Page")
XMLAttr.Value = Application("SECURE_URL") & Request.ServerVariables("SCRIPT_NAME")
XMLAgency.setAttributeNode XMLAttr

'This is the it should return to
set XMLAttr = XmlDoc.CreateAttribute("ReturnPage")
if right(formPath,1) = "/" then
	XMLAttr.Value = Application("SECURE_URL") & formPath & getNextPage
else
	XMLAttr.Value = Application("SECURE_URL") & formPath & "/" & getNextPage
end if
XMLAgency.setAttributeNode XMLAttr

set XMLAttr = XmlDoc.CreateAttribute("FormID")
XMLAttr.Value = session("APP_RENEW_ID")
XMLAgency.setAttributeNode XMLAttr

XMLPayment.AppendChild(XMLAgency)

Dim objPI
'Create the xml processing instruction.
Set objPI = XmlDoc.createProcessingInstruction("xml", "version='1.0' encoding='ISO-8859-1'")
'Append the processing instruction to the XML document.
XmlDoc.AppendChild(objPI)

XmlDoc.AppendChild(XMLPayment)
			     
set GenerateXML = XmlDoc   

set xXMLPayment = nothing

set xmlDoc = nothing

set xmlAgency = nothing

set xmlOptions = nothing

set XMLBilling = nothing

set XMLAttr = nothing

set objPI = nothing
   
end function
%>

<%function runProfitStarsTrans(parmList)%>
<%
Dim objRequest, SERVER_URL
Dim requestID, idxerr
Dim post_response

SERVER_URL = Application("MRI_SERVER_URL") & "APIPayment/ProcessPayment"

Set objRequest = Server.CreateObject("Microsoft.XMLHTTP")
objRequest.open "POST", SERVER_URL, false
objRequest.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
objRequest.setRequestHeader "Accept", "application/xml"
objRequest.send parmList
post_response = objRequest.responseText
Set objRequest = nothing

runProfitStarsTrans = post_response

'response.write "parmList = " & parmList & "<BR>"
'response.write "<textarea>" & post_response & "</textarea><BR>"
'response.end
%>
<%end function%>

<%function profitStars%>
<%
if session("CARD_NUM")&"" = "" then
	profitStars = true
	exit function
end if

if session("NameID")&"" = "" then
	session("NameID") = getNameID
end if

dim parmList, transResponse, result, RESPMSG, xmlDoc, rsProd, prodDesc

parmList = ""
prodDesc = ""

set cSqlForm = new cRunSql
cSqlForm.conn = connOracle
cSqlForm.SqlStr = "select product_desc from " & prodTbl & " where product_id in (" & session("PRODUCT_ID") & ")"
cSqlForm.AddParam (session("PRODUCT_ID"))
set rsProd = cSqlForm.Execute
do while not rsProd.eof
	if prodDesc <> "" then
		prodDesc = prodDesc & vbcrlf
	end if
	prodDesc = prodDesc & rsProd("PRODUCT_DESC")
	rsProd.MoveNext
loop
rsProd.close
set rsProd = nothing
set cSqlForm = nothing

parmList = parmList & "HoldOnly=True"
parmList = parmList & "&nameOnAccount=" & request.form("CARD_FIRST_NAME") & " " & request.form("CARD_LAST_NAME")
parmList = parmList & "&address1=" & request.form("CARD_ADDRESS_LINE_1")
parmList = parmList & "&address2=" & request.form("CARD_ADDRESS_LINE_2")
parmList = parmList & "&city=" + request.form("CARD_CITY") 
parmList = parmList & "&state=" + request.form("CARD_STATE") 
parmList = parmList & "&postalCode=" + request.form("CARD_ZIP") 
parmList = parmList & "&country=" + request.form("CARD_COUNTRY") 
parmList = parmList & "&accountNumber=" & request.form("CARD_NUM")
parmList = parmList + "&cvv=" & request.form("CARD_CVV")
parmList = parmList & "&cardMonth=" & request.form("CARD_EXP_MO")
parmList = parmList & "&cardYear=" & request.form("CARD_EXP_YR")
parmList = parmList & "&amount=" & session("CARD_AMOUNT")
parmList = parmList & "&transactionDateTime=" & now
parmList = parmList & "&nameId=" & session("NameID")
parmList = parmList & "&description=" & prodDesc
parmList = parmList & "&convenienceFee=0"
parmList = parmList & "&transactionNumber=" & session("CARD_CONFIRMATION")

transResponse = runProfitStarsTrans(parmList)

set xmlDoc = server.CreateObject("Microsoft.XMLDOM")
xmlDoc.loadxml(transResponse)

on error resume next
RESULT = xmlDoc.selectSingleNode("//processingResult").Attributes.GetNamedItem("success").Text
on error goto 0

if RESULT&"" = "true" then
	session("PNREF") = xmlDoc.selectSingleNode("//processingResult/messages/referencenumber").text
	session("PPREF") = session("PNREF") 
	profitStars = true
else
	if instr(transResponse,"processingResult") then
		cForm.AddErrorMsg "Transaction failed - " & xmlDoc.selectSingleNode("//processingResult/errors/error").text
	end if
	if isDev or officeIP then
		cForm.AddErrorMsg("parmlist - " & parmlist)
		cForm.AddErrorMsg("transResponse - " & transResponse)
		paymentProcessorError
	end if
	profitStars = false
end if

set xmlDoc = nothing
%>
<%end function%>

<%dim paypalWPPURL%>

<%function runPayPalWPPTrans(parmList, transID)%>
<%Dim objWinHttp, SERVER_URL
Dim requestID

SERVER_URL = "https://api-3t.paypal.com/nvp"

if isDev then
	SERVER_URL = "https://api-3t.sandbox.paypal.com/nvp"
end if

paypalWPPURL = SERVER_URL

Set objWinHttp = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
objWinHttp.Open "POST", SERVER_URL, False
objWinHttp.setRequestHeader "Content-Type", "text/namevalue" ' for XML, use text/xml
objWinHttp.SetRequestHeader "X-VPS-Timeout", "30"
objWinHttp.SetRequestHeader "X-VPS-Request-ID", transID

'Send Parameter List
on error resume next
objWinHttp.Send parmList
if err.number <> 0 then
    dim idxerr
    for idxerr = 1 to 5
        objWinHttp.Send parmlist
        if err.number = 0 then
            i = 5
        end if
    next
end if
on error goto 0

runPayPalWPPTrans = objWinHttp.ResponseText
Set objWinHttp = Nothing
%>
<%end function%>

<%function PayPalWPP%>
<% 
if session("CARD_NUM")&"" = "" then
	PayPalWPP = true
	exit function
end if

if session("PAYPAL_SUBMITTED") = "1" then
	'cForm.ErrorMsg = "Payment already submitted."
	'PayPalWPP = false
	'exit function
end if

session("PAYPAL_SUBMITTED") = "1"

dim parmList, transResponse, result, RESPMSG

parmList = ""

parmList = parmList & "METHOD=DoDirectPayment"
parmList = parmList & "&PAYMENTACTION=Sale"
parmList = parmList & "&VERSION=76.0"
parmList = parmList & "&CURRENCYCODE=USD"
parmList = parmList & "&IPADDRESS=" & request.servervariables("HTTP_X_FORWARDED_FOR")
parmList = parmList & "&FIRSTNAME=" & request.form("CARD_FIRST_NAME")
parmList = parmList & "&LASTNAME=" & request.form("CARD_LAST_NAME")
parmList = parmList & "&STREET=" & request.form("CARD_ADDRESS_LINE_1")
parmList = parmList & "&CITY=" & request.form("CARD_CITY")
parmList = parmList & "&STATE=" & request.form("CARD_STATE")

'set the account from the html form after cleaning it
'cleanup spaces, /, \, and -
Set regEx = New RegExp
regEx.Global = true
regEx.IgnoreCase = True
regEx.pattern = "\s|\-|\\|\/"
cleanCCNum = trim(regEx.replace(request.form("CARD_NUM"), ""))
parmList = parmList & "&ACCT=" & cleanCCNum
set regEx = nothing

parmList = parmList & "&CVV2=" & request.form("CARD_CVV")

parmList = parmList & "&EXPDATE=" & request.form("CARD_EXP_MO")&request.form("CARD_EXP_YR")
parmList = parmList & "&USER=" & cRenew.ConfigOption("paypal_user")
parmList = parmList & "&PWD=" & cRenew.ConfigOption("paypal_password")
parmList = parmList & "&SIGNATURE=" & cRenew.ConfigOption("paypal_signature")
parmList = parmList & "&AMT=" & session("CARD_AMOUNT")

formLog("initiating PayPal transaction")
transResponse = runPayPalWPPTrans(parmList,session("CARD_CONFIRMATION"))
formLog("PayPal response received")

session("TRANS_RESPONSE") = transResponse

RESULT = ShowPayPalWPPResponse(transResponse,"ACK")
RESPMSG = ShowPayPalWPPResponse(transResponse,"L_LONGMESSAGE0")
		   
if RESULT&"" = "Success" or RESULT&"" = "SuccessWithWarning" then
	session("PNREF") = ShowPayPalResponse(transResponse,"PNREF")
	session("PPREF") = ShowPayPalResponse(transResponse,"PPREF")
	PayPalWPP = true
else
	cForm.AddErrorMsg "Transaction failed - " & RESPMSG
	if isDev or officeIP then
		cForm.AddErrorMsg("endpoint - " & paypalWPPURL)
		cForm.AddErrorMsg("parmlist - " & parmlist)
		cForm.AddErrorMsg("transResponse - " & transResponse)
		paymentProcessorError
	end if
	PayPalWPP = false
	session("PAYPAL_SUBMITTED") = "0"
end if
%>
<%End function%>

<%
Function ShowPayPalWPPResponse(transResponse, Param)

dim arrWPP, arrWPPDet, MyValue

arrWPP = split(transResponse&"","&")

for i = lbound(arrWPP) to ubound(arrWPP)
	arrWPPDet = split(arrWPP(i),"=")
	if ucase(arrWPPDet(0)) = ucase(Param) then
		MyValue = URLDecode(arrWPPDet(1))
	end if
next

ShowPayPalWPPResponse = MyValue

End Function
%>

<%function PayPal%>
<% 
if session("CARD_NUM")&"" = "" then
	PayPal = true
	exit function
end if

if session("PAYPAL_SUBMITTED") = "1" then
	cForm.ErrorMsg = "Payment already submitted."
	PayPal = false
	exit function
end if

session("PAYPAL_SUBMITTED") = "1"

dim parmList, transResponse, result, RESPMSG

parmList = ""

'set the TRXTYPE(Transaction Type) from default
parmList = parmList & "TRXTYPE=S"

'set the TENDER(Credit Card etc...) from default
parmList = parmList & "&TENDER=C"

'set the INVNUM from default
''parmList = parmList & "&INVNUM=" & getBatch_NUM
parmList = parmList & "&INVNUM=" & session("CARD_CONFIRMATION")

'set the COMMENT1 from default
parmList = parmList & "&NAME=" & request.form("CARD_FIRST_NAME") & " " & request.form("CARD_LAST_NAME")

'set the zip from the html form
parmList = parmList & "&ZIP=" + request.form("CARD_ZIP") 

'set the address from the html form
parmList = parmList & "&STREET["&Len(request.form("CARD_ADDRESS_LINE_1"))&"]=" & request.form("CARD_ADDRESS_LINE_1")

'set the account from the html form after cleaning it
'cleanup spaces, /, \, and -
Set regEx = New RegExp
regEx.Global = true
regEx.IgnoreCase = True
regEx.pattern = "\s|\-|\\|\/"
cleanCCNum = trim(regEx.replace(request.form("CARD_NUM"), ""))
parmList = parmList & "&ACCT=" & cleanCCNum
set regEx = nothing

if cRenew.ConfigOption("paypal_partner")&"" <> "" then
	parmList = parmList & "&PWD=" & cRenew.ConfigOption("paypal_password")
	parmList = parmList & "&USER=" & cRenew.ConfigOption("paypal_user")
	parmList = parmList & "&VENDOR=" & cRenew.ConfigOption("paypal_vendor")
	parmList = parmList + "&PARTNER=" & cRenew.ConfigOption("paypal_partner")
	if cRenew.ConfigOption("paypal_comment1")&"" <> "" then
		if objFields.Exists(clng(cRenew.ConfigOption("paypal_comment1"))) then
			parmList = parmList + "&COMMENT1=" & session(objFields.Item(clng(cRenew.ConfigOption("paypal_comment1"))))
		end if
	end if
	if cRenew.ConfigOption("paypal_comment2")&"" <> "" then
		if objFields.Exists(clng(cRenew.ConfigOption("paypal_comment2"))) then
			parmList = parmList + "&COMMENT2=" & session(objFields.Item(clng(cRenew.ConfigOption("paypal_comment2"))))
		end if
	end if
else
	parmList = parmList & "&PWD=" & Application("PAYPAL_PASSWORD")
	parmList = parmList & "&USER=" & Application("PAYPAL_USER")
	parmList = parmList & "&VENDOR=" & Application("PAYPAL_VENDOR")
	parmList = parmList + "&PARTNER=" & Application("PAYPAL_PARTNER")
end if

'set cvv2
parmList = parmList + "&CVV2=" & request.form("CARD_CVV")

'set the expiration date form the HTML form
parmList = parmList & "&EXPDATE=" & request.form("CARD_EXP_MO")&""&right(request.form("CARD_EXP_YR"),2)

parmList = parmList & "&AMT=" & session("CARD_AMOUNT")

formLog("initiating PayPal transaction")
transResponse = runPayPalTrans(parmList,session("CARD_CONFIRMATION"))
formLog("PayPal response received - " & transResponse)
session("TRANS_RESPONSE") = transResponse

RESULT = ShowPayPalResponse(transResponse,"RESULT")
RESPMSG = ShowPayPalResponse(transResponse,"RESPMSG")
		   
if RESULT&"" = "0" then
	session("PNREF") = ShowPayPalResponse(transResponse,"PNREF")
	session("PPREF") = ShowPayPalResponse(transResponse,"PPREF")
	formLog("PNREF = " & session("PNREF"))
	formLog("PPREF = " & session("PPREF"))
	PayPal = true
else
	cForm.AddErrorMsg "Transaction failed - " & RESPMSG
	if isDev or officeIP then
		cForm.AddErrorMsg("parmlist - " & parmlist)
		cForm.AddErrorMsg("transResponse - " & transResponse)
		paymentProcessorError
	end if
	PayPal = false
	session("PAYPAL_SUBMITTED") = "0"
end if
%>
<%End function%>

<%
Function ShowPayPalResponse(transResponse, Param)

dim MyValue, curString, varString, name, value

curString = transResponse

Do while Len(curString) <> 0

	if InStr(curString,"&") Then
	   varString = Left(curString, InStr(curString , "&" ) -1)
	else
	   varString = curString
	end if

	name = Left(varString, InStr(varString, "=" ) -1)
	value = Right(varString, Len(varString) - (Len(name)+1))

	if name = Param Then
	   MyValue = value
	   Exit Do
	end if

	if Len(curString) <> Len(varString) Then
	   curString = Right(curString, Len(curString) - (Len(varString)+1))
	else
	   curString = ""
	end if
Loop

ShowPayPalResponse = MyValue

End Function
%>

<%function runPayPalTrans(parmList, transID)%>
<%Dim objWinHttp, SERVER_URL
Dim requestID

SERVER_URL = "https://payflowpro.paypal.com"

if isDev then
	SERVER_URL = "https://pilot-payflowpro.paypal.com"
end if

Set objWinHttp = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
objWinHttp.Open "POST", SERVER_URL, False
objWinHttp.setRequestHeader "Content-Type", "text/namevalue" ' for XML, use text/xml
objWinHttp.SetRequestHeader "X-VPS-Timeout", "30"
objWinHttp.SetRequestHeader "X-VPS-Request-ID", transID

'Send Parameter List
on error resume next
objWinHttp.Send parmList
if err.number <> 0 then
    dim idxerr
    for idxerr = 1 to 5
        objWinHttp.Send parmlist
        if err.number = 0 then
            i = 5
        end if
    next
end if
on error goto 0

runPayPalTrans = objWinHttp.ResponseText
Set objWinHttp = Nothing
%>
<%end function%>

<%function getSequence(sequenceName)%>
<%
dim rsSeq, cSqlSeq

set cSqlSeq = new cRunSql
cSqlSeq.Conn = connOracle
cSqlSeq.SqlStr = "Select "&sequenceName&".NextVal from DUAL"
set rsSeq = cSqlSeq.Execute
getSequence = rsSeq("NextVal")
rsSeq.close
set rsSeq = nothing
set cSqlSeq = nothing
%>
<%end function%>

<%function getSequenceMin(sequenceName,seqMin)%>
<%
dim rsSeq, cSqlSeq, seq

set cSqlSeq = new cRunSql
cSqlSeq.Conn = connOracle
cSqlSeq.SqlStr = "Select "&sequenceName&".NextVal from DUAL"
set rsSeq = cSqlSeq.Execute
seq = rsSeq("NextVal")
rsSeq.close
set rsSeq = nothing
set cSqlSeq = nothing

if len(seq) < seqMin then
	do while len(seq) < seqMin
		seq = "0" & seq
	loop
end if
getSequenceMin = seq
%>
<%end function%>

<%function authorize%>
<%
if session("CARD_NUM")&"" = "" then
	authorize = true
	exit function
end if

if session("AUTHORIZE_SUBMITTED") = "1" then
	cForm.ErrorMsg = "Payment already submitted."
	authorize = false
	exit function
end if

session("AUTHORIZE_SUBMITTED") = "1"

Dim post_url
post_url = "https://secure.authorize.net/gateway/transact.dll"
if isDev then
	post_url = "https://test.authorize.net/gateway/transact.dll"
end if

Dim post_values
Set post_values = CreateObject("Scripting.Dictionary")
post_values.CompareMode = vbTextCompare

'the API Login ID and Transaction Key must be replaced with valid values
if cRenew.ConfigOption("authorize_api_id")&"" <> "" then
	post_values.Add "x_login", cRenew.ConfigOption("authorize_api_id")
	post_values.Add "x_tran_key", cRenew.ConfigOption("authorize_transaction_key")
else
	post_values.Add "x_login", Application("AUTHORIZE_API_ID")
	post_values.Add "x_tran_key", Application("AUTHORIZE_TRANSACTION_KEY")
end if

if cRenew.ConfigOption("authorize_comment1")&"" <> "" then
	if objFields.Exists(clng(cRenew.ConfigOption("authorize_comment1"))) then
		post_values.Add "description", session(objFields.Item(clng(cRenew.ConfigOption("authorize_comment1"))))
	end if
end if

post_values.Add "x_delim_data", "TRUE"
post_values.Add "x_delim_char", "|"
post_values.Add "x_relay_response", "FALSE"

post_values.Add "x_type", "AUTH_CAPTURE"
post_values.Add "x_method", "CC"
post_values.Add "x_card_num", request.form("CARD_NUM")
post_values.Add "x_card_code", request.form("CARD_CVV")
post_values.Add "x_exp_date", request.form("CARD_EXP_MO")&""&request.form("CARD_EXP_YR")

post_values.Add "x_amount", session("CARD_AMOUNT")
post_values.Add "x_invoice_num", session("CARD_CONFIRMATION")
post_values.Add "invoiceNumber", session("CARD_CONFIRMATION")

post_values.Add "x_first_name", request.form("CARD_FIRST_NAME")
post_values.Add "x_last_name", request.form("CARD_LAST_NAME")
post_values.Add "x_address", request.form("CARD_ADDRESS_LINE_1")
post_values.Add "x_state", request.form("CARD_STATE")
post_values.Add "x_zip", request.form("CARD_ZIP")

post_values.Add "x_cust_id", session("MEMBERSHIP_ID")

Dim post_string 
post_string = ""
For Each Key In post_values
  post_string=post_string & Key & "=" & Server.URLEncode(post_values(Key)) & "&"
Next
post_string = Left(post_string,Len(post_string)-1)

formLog("initiating Authorize transaction")

' We use xmlHTTP to submit the input values and record the response
Dim objRequest, post_response
Set objRequest = Server.CreateObject("Microsoft.XMLHTTP")
objRequest.open "POST", post_url, false
objRequest.send post_string
post_response = objRequest.responseText
Set objRequest = nothing

formLog("Authorize response received")

'response.write "post_response = " & post_response & "<BR>"
' the response string is broken into an array using the specified delimiting character
Dim authorizeResponse
authorizeResponse = split(post_response, post_values("x_delim_char"))
session("TRANS_RESPONSE") = post_response
if authorizeResponse(0) = "1" then
	authorize = true
	session("PNREF") = authorizeResponse(6)
else
	authorize = false
	cForm.AddErrorMsg "Transaction failed - " & authorizeResponse(3)	
	if isDev or officeIP then
		cForm.AddErrorMsg "post_string - " & post_string & "<BR>"
		cForm.AddErrorMsg "post_response - " & post_response & "<BR>"
		paymentProcessorError
	end if
	session("AUTHORIZE_SUBMITTED") = "0"
end if
set post_values = nothing
%>
<%end function%>

<%function elavonResponseField(post_response,field)%>
<%
dim elavonXML, elavonNodes, elavonNode

set elavonXML = server.CreateObject("Microsoft.XMLDOM")
elavonXML.loadxml(post_response)
Set elavonNodes = elavonXML.getElementsByTagName(field)
for each elavonNode in elavonNodes
	elavonResponseField = elavonNode.Text
next
set elavonNodes = nothing
set elavonXML = nothing
%>
<%end function%>

<%function monerisResponseField(post_response,field)%>
<%
dim monerisXML, monerisNodes, monerisnNode

set monerisXML = server.CreateObject("Microsoft.XMLDOM")
monerisXML.loadxml(post_response)
Set monerisNodes = monerisXML.getElementsByTagName(field)
for each monerisnNode in monerisNodes
	monerisResponseField = monerisnNode.Text
next
set monerisNodes = nothing
set monerisXML = nothing
%>
<%end function%>

<%function moneris%>
<%
Dim post_url, post_string, retVal
Dim monerisResponseCode, monerisISO, monerisMessage

post_url = "https://www3.moneris.com/gateway2/servlet/MpgRequest"
if isDev then
	post_url = "https://esqa.moneris.com/gateway2/servlet/MpgRequest"
end if

session("CARD_CONFIRMATION") = getSequence("SEQ_VERISIGN_LOG_ID")

post_string = "<?xml version=""1.0"" encoding=""UTF-8""?><request>" &_
	"<store_id>" & cRenew.ConfigOption("moneris_store_id") & "</store_id>" &_
	"<api_token>" & cRenew.ConfigOption("moneris_api_key") & "</api_token>" &_
	"<purchase>" &_
	"<pan>" & request.form("CARD_NUM") & "</pan>" &_
	"<expdate>" & request.form("CARD_EXP_MO")&""&right(request.form("CARD_EXP_YR"),2) & "</expdate>" &_
	"<amount>" & FormatNumber(session("CARD_AMOUNT"),,,,0) & "</amount>" &_
	"<order_id>" & session("CARD_CONFIRMATION") & "</order_id>" &_
	"<crypt_type>7</crypt_type>" &_
	"</purchase>" &_
	"</request>"


formLog("initiating Moneris transaction")

Dim objRequest, post_response
Set objRequest = CreateObject("MSXML2.XMLHTTP.6.0")
objRequest.open "POST", post_url, false
objRequest.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
objRequest.send(post_string)
post_response = objRequest.responseText
Set objRequest = nothing

if instr(post_response,"</receipt></response>") then
else
	post_response = post_response & "</receipt></response>"
end if

formLog("Moneris response received - " & post_response)

session("TRANS_RESPONSE") = post_response

monerisResponseCode = monerisResponseField(post_response,"ResponseCode")
monerisISO = monerisResponseField(post_response,"ISO")
monerisMessage = monerisResponseField(post_response,"Message")
if monerisResponseCode = "027" and left(monerisMessage,8) = "APPROVED" then
	if monerisISO = "00" or monerisISO = "01" or monerisISO = "10" or monerisISO = "85" then
		retval = true
		session("PNREF") = monerisResponseField(post_response,"ReferenceNum")
	else
		retVal = false
	end if
end if
if not retVal then
	cForm.AddErrorMsg "Payment failed - " & monerisMessage
	if isDev or officeIP then
		cForm.AddErrorMsg "post_url - " & post_url & "<BR>"
		cForm.AddErrorMsg "post_string - <textarea>" & post_string & "</textarea><BR>"
		cForm.AddErrorMsg "post_response - <textarea>" & post_response & "</textarea><BR>"
		paymentProcessorError
	end if
end if
moneris = retVal
%>
<%end function%>

<%function clover%>
<%
Dim post_url, post_string, retVal
Dim brand, cloverToken, oJSON
Dim objRequest, post_response

if cRenew.ConfigOption("clover_public_key")&"" = "" or cRenew.ConfigOption("clover_secret_key")&"" = "" then
	cForm.AddErrorMsg "Invalid Clover setup"
	clover = False
	exit function
end if

post_url = "https://token.clover.com/v1/tokens"
if isDev then
	post_url = "https://token-sandbox.dev.clover.com/v1/tokens"
end if

if left(request.form("CARD_NUM"),1) = "4" then	
	brand = "VISA"
elseif left(request.form("CARD_NUM"),1) = "6" then	
	brand = "DISCOVER"
elseif left(request.form("CARD_NUM"),1) = "3" then	
	brand = "AMEX"
elseif left(request.form("CARD_NUM"),1) = "5" or left(request.form("CARD_NUM"),1) = "2" then	
	brand = "MC"
else 
	brand = ""
end if

' Construct the JSON payload

formLog("initiating Clover transaction")

post_string = "{""card"":{""number"":""" & request.form("CARD_NUM") & """," &_
	"""exp_month"":""" & request.form("CARD_EXP_MO") & """," & _
	"""exp_year"":""" & right(request.form("CARD_EXP_YR"),2) & """," & _
	"""cvv"":""" & request.form("CARD_CVV") & """," & _
	"""brand"":""" & brand & """}}"

formLog("requesting Clover card token")

Set objRequest = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
objRequest.Open "POST", post_url, False
objRequest.setRequestHeader "Accept", "application/json"
objRequest.setRequestHeader "Content-Type", "application/json"
objRequest.setRequestHeader "apikey", cRenew.ConfigOption("clover_public_key")
objRequest.Send post_string
post_response = objRequest.responseText

formLog("Clover card token response received - " & post_response)

Set oJSON = New aspJSON
oJSON.loadJSON(post_response)
if objRequest.status = 200 then
	cloverToken = oJSON.data("id")
else
	if oJSON.data("message") = "Invalid JSON format" then
		cForm.AddErrorMsg "Payment failed"
	elseif oJSON.data("message") = "401 Unauthorized" then
		cForm.AddErrorMsg "Payment failed"
		if isDev or officeIP then
			cForm.AddErrorMsg "Unauthorized - review credentials<BR>"
		end if
	else
		cForm.AddErrorMsg "Payment failed - " & oJSON.data("error")("message")
	end if
	if isDev or officeIP then
		cForm.AddErrorMsg "post_url - " & post_url & "<BR>"
		cForm.AddErrorMsg "post_string - <textarea>" & post_string & "</textarea><BR>"
		cForm.AddErrorMsg "post_response - <textarea>" & post_response & "</textarea><BR>"
		paymentProcessorError
	end if
end if
set oJSON = nothing
Set objRequest = Nothing

if cloverToken&"" <> "" then
	post_url = "https://scl.clover.com/v1/charges"
	if isDev then
		post_url = "https://scl-sandbox.dev.clover.com/v1/charges"
	end if

	session("CARD_CONFIRMATION") = getSequence("SEQ_VERISIGN_LOG_ID")

	formLog("sending Clover payment request")
	
	post_string = "token=" & Server.URLEncode(cloverToken) & "&url=" & Server.URLEncode(post_url) & "&key=" & Server.URLEncode(cRenew.ConfigOption("clover_secret_key")) & "&amount=" & Server.URLEncode((cint(session("CARD_AMOUNT"))*100)) & "&ip=" & clientIP

	Set objRequest = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
	objRequest.open "POST", "https://" & request.ServerVariables("SERVER_NAME") & "/www/clover/", false
	objRequest.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
	objRequest.setRequestHeader "Accept","application/json"
	on error resume next
	objRequest.send(post_string)
	If Err.Number <> 0 Then
		Response.Write "Error: " & Err.Description
		Response.End()
	End If
	On Error GoTo 0

	post_response = objRequest.responseText
	formLog("Clover payment response received - " & post_response)
	session("TRANS_RESPONSE") = post_response

	Set oJSON = New aspJSON
	oJSON.loadJSON(post_response)
	if objRequest.status = 200 then
		retVal = true
		session("PNREF") = oJSON.data("id")
	else
		retVal = false
		if oJSON.data("message") = "Invalid JSON format" then
			cForm.AddErrorMsg "Payment failed"
		elseif oJSON.data("message") = "401 Unauthorized" then
			cForm.AddErrorMsg "Payment failed"
			if isDev or officeIP then
				cForm.AddErrorMsg "Unauthorized - review credentials<BR>"
			end if
		else
			cForm.AddErrorMsg "Payment failed - " & oJSON.data("error")("message")
		end if
		if isDev or officeIP then
			cForm.AddErrorMsg "post_url - " & post_url & "<BR>"
			cForm.AddErrorMsg "post_string - {""amount"":" & (cint(session("CARD_AMOUNT"))*100) & ",""currency"":""usd"",""source"":""" & cloverToken & """}" & "<BR>"
			cForm.AddErrorMsg "post_response - <textarea>" & post_response & "</textarea><BR>"
			paymentProcessorError
		end if
	end if
	Set objRequest = nothing
else
	retVal = false
end if

clover = retVal
%>
<%end function%>

<%function elavon%>
<%
Dim post_url, elavonErr, elavonResponse
post_url = "https://api.convergepay.com/VirtualMerchant/processxml.do"
if isDev then
	post_url = "https://api.demo.convergepay.com/VirtualMerchantDemo/processxml.do"
end if

Dim post_values
Set post_values = CreateObject("Scripting.Dictionary")
post_values.CompareMode = vbTextCompare

'the API Login ID and Transaction Key must be replaced with valid values
post_values.Add "ssl_merchant_id", cRenew.ConfigOption("elavon_merchant_id")
post_values.Add "ssl_user_id", cRenew.ConfigOption("elavon_user_id")
post_values.Add "ssl_pin", cRenew.ConfigOption("elavon_pin")
post_values.Add "ssl_transaction_type", "ccsale"
if isDev then
	post_values.Add "ssl_test_mode", "true"
else
	post_values.Add "ssl_test_mode", "false"
end if

post_values.Add "ssl_card_number", request.form("CARD_NUM")
post_values.Add "ssl_cvv2cvc2_indicator", "1"
post_values.Add "ssl_cvv2cvc2", request.form("CARD_CVV")
post_values.Add "ssl_exp_date", request.form("CARD_EXP_MO")&""&right(request.form("CARD_EXP_YR"),2)

post_values.Add "ssl_amount", session("CARD_AMOUNT")
post_values.Add "ssl_invoice_num", VERISIGN_LOG_ID

post_values.Add "ssl_first_name", request.form("CARD_FIRST_NAME")
post_values.Add "ssl_last_name", request.form("CARD_LAST_NAME")
post_values.Add "ssl_avs_address", request.form("CARD_ADDRESS_LINE_1")
post_values.Add "ssl_avs_zip", request.form("CARD_ZIP")
post_values.Add "ssl_city", request.form("CARD_CITY")
post_values.Add "ssl_state", request.form("CARD_STATE")

post_values.Add "ssl_cust_id", session("MEMBERSHIP_ID")

if cRenew.ConfigOption("elavon_dob")&"" <> "" then
	if objFields.Exists(clng(cRenew.ConfigOption("elavon_dob"))) then
		post_values.Add "Client_DOB", session(objFields.Item(clng(cRenew.ConfigOption("elavon_dob"))))
	end if
end if
if cRenew.ConfigOption("elavon_ssn")&"" <> "" then
	if objFields.Exists(clng(cRenew.ConfigOption("elavon_ssn"))) then
		post_values.Add "Last_4_Social", right(session(objFields.Item(clng(cRenew.ConfigOption("elavon_ssn"))))&"",4)
	end if
end if
if isDev then
else
	if cRenew.ConfigOption("elavon_email")&"" <> "" then
		if objFields.Exists(clng(cRenew.ConfigOption("elavon_email"))) then
			post_values.Add "ssl_email", session(objFields.Item(clng(cRenew.ConfigOption("elavon_email"))))
		end if
	end if
end if
if cRenew.ConfigOption("elavon_phone")&"" <> "" then
	if objFields.Exists(clng(cRenew.ConfigOption("elavon_phone"))) then
		post_values.Add "ssl_phone", session(objFields.Item(clng(cRenew.ConfigOption("elavon_phone"))))
	end if
end if
if cRenew.ConfigOption("elavon_lic_no")&"" <> "" then
	if objFields.Exists(clng(cRenew.ConfigOption("elavon_lic_no"))) then
		post_values.Add "License_Number", session(objFields.Item(clng(cRenew.ConfigOption("elavon_lic_no"))))
	end if
end if

Dim post_string

post_string = "xmldata=<txn>"
For Each Key In post_values
	post_string = post_string & "<" & key & ">" & Server.URLEncode(post_values(Key)) & "</" & key & ">"
next
post_string = post_string & "</txn>"

formLog("initiating Elavon transaction")

' We use xmlHTTP to submit the input values and record the response
Dim objRequest, post_response
Set objRequest = CreateObject("MSXML2.XMLHTTP.6.0")
objRequest.open "POST", post_url, false
objRequest.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
objRequest.send(post_string)
post_response = objRequest.responseText
Set objRequest = nothing

formLog("Elavon response received")
formLog("Elavon response = " & post_response)

session("TRANS_RESPONSE") = post_response

' the response string is broken into an array using the specified delimiting character
elavonResponse = elavonResponseField(post_response,"ssl_result")
if elavonResponse = "0" then
	session("PNREF") = elavonResponseField(post_response,"ssl_txn_id")
	session("PPREF") = elavonResponseField(post_response,"ssl_approval_code")
	elavon = true
else
	elavon = false
	elavonErr = elavonResponseField(post_response,"errorName")
	if elavonErr&"" = "" then
		elavonErr = elavonResponseField(post_response,"ssl_result_message")
	end if

	cForm.AddErrorMsg "Transaction failed - " & elavonErr
	if isDev or officeIP then
		cForm.AddErrorMsg "post_url - " & post_url & "<BR>"
		cForm.AddErrorMsg "post_string - <textarea>" & post_string & "</textarea><BR>"
		cForm.AddErrorMsg "post_response - <textarea>" & post_response & "</textarea><BR>"
		paymentProcessorError
	end if
end if
set post_values = nothing
%>
<%end function%>

<%function stripe
	dim objWinHttp, oJson, stripeStatus, paymentID

	Set objWinHttp = Server.CreateObject("MSXML2.serverXMLHTTP")
	objWinHttp.Open "GET", "https://api.stripe.com/v1/payment_intents/" & session("stripe_payment_intent_id"), False
	objWinHttp.setRequestHeader "Authorization", "Bearer " & cRenew.ConfigOption("stripe_secret_key") 
	objWinHttp.Send()

	Set oJSON = New aspJSON
	oJSON.loadJSON(objWinHTTP.ResponseText)
	session("TRANS_RESPONSE") = objWinHTTP.ResponseText
	stripeStatus = oJSON.data("status")&""
	paymentID = oJSON.data("latest_charge")
	set oJSON = nothing

	if stripeStatus <> "succeeded" then
		cForm.ErrorMsg = "Payment failed"
		stripe = false
	else
		session("PNREF") = session("stripe_payment_intent_id")
		session("PPREF") = paymentID
		if instr(objWinHTTP.ResponseText,"last4") then
			session("CARD_NUM") = replace(replace(split(split(objWinHTTP.ResponseText,"last4")(1),",")(0),""": """,""),"""","")
			session("CARD_NUM_RECEIPT") = session("CARD_NUM")
		end if
		stripe = true
	end if

	set objWinHTTP = nothing

end function
%>

<%
sub loginGoogleSSO
	dim redisSSO, googleEmail, rsLogin

    redisKey = "externalauth:" & request.querystring("ExternalAuthId")
	set redisSSO = Server.CreateObject("CacheManager")
    redisSSO.Init redisKey 
    googleEmail = redisSSO("email")
    set redisSSO = Nothing

	if doAjax then
		response.clear
		Response.ContentType = "application/json; charset=utf-8"
	end if

	if googleEmail&"" <> "" then
		set cSqlForm = new cRunSql
		cSqlForm.conn = connOracle
        cSqlForm.sqlStr = "select LOGIN_ID from login where inactive = 'N' and lower(google_email) = lower(?)"
		cSqlForm.AddParam(googleEmail)
		set rsLogin = cSqlForm.Execute
		set cSqlForm = nothing
		if rsLogin.eof then
			session("GOOGLE_EMAIL") = googleEmail
		else
			session("LOGIN_ID") = rsLogin("LOGIN_ID")
			set cSqlForm = new cRunSql
			cSqlForm.conn = connOracle
			cSqlForm.sqlStr = "select membership_id from member_detail m, org_column_detail o " &_
				"where o.column_id = m.column_id " &_
				"and o.form_field_type = 'SYSUSER' " &_
				"and m.data = (select login from login where login_id = ?) "
			cSqlForm.AddParam(session("LOGIN_ID"))
			set rs = cSqlForm.Execute
			if rs.eof then
				cForm.AddErrorMsg("No CRM record found")
			else
				session("MEMBERSHIP_ID") = rs("MEMBERSHIP_ID")
			end if
			rs.close
			set rs = nothing
			set cSqlForm = nothing
			rsLogin.Close
			set rsLogin = nothing
			goNext
		end if
		rsLogin.Close
		set rsLogin = nothing
	else
		if doAjax then
			response.Status = "401 - Invalid credentials"
		else
			cForm.AddErrorMsg("No user found")
		end if
	end if
end sub

sub loginUser(userName,password)
	if userName&"" <> "" and password&"" <> "" then
		dim urlSalt, hashPassword, oXMLHTTP, params, rsLogin, url
		dim loginName, loginFirstName, loginLastName, loginGroup
		dim accountLocked, maxAttempts, passwordExpiration, passwordExpired

		passwordExpired = false

		set cCV = new cCodeValueO
		cCV.LoadFromCodeClassAndValue "APPLICATION_DB","PASSWORD_EXPIRATION"
		if cCV.eof then
			passwordExpiration = 1000
		else
			if isNumeric(cCV.CodeValueDesc) then
				passwordExpiration = cint(cCV.CodeValueDesc)
			end if
		end if
		set cCV = nothing

		set cCV = new cCodeValueO
		cCV.LoadFromCodeClassAndValue "APPLICATION_DB","MAX_LOGIN_ATTEMPTS"
		if cCV.eof then
			maxAttempts = 100
		else
			if isNumeric(cCV.CodeValueDesc) then
				maxAttempts = cint(cCV.CodeValueDesc)
			end if
		end if
		set cCV = nothing

		session("LOGIN_ID") = ""

		set cSqlForm = new cRunSql
		cSqlForm.conn = connOracle
		cSqlForm.SqlStr = "select LOGIN,SALT,LOGIN_ID,FIRST_NAME,LAST_NAME,GROUP_TY_CD,PASSWORD_HASH,PASSWORD_MODIFY_DT,ACCOUNT_LOCKED from login where upper(login) = upper(?) and inactive = 'N'"
		cSqlForm.AddParam(userName)
		set rsLogin = cSqlForm.Execute
		if rsLogin.eof then
		else
			hashPassword = getPasswordHash(password,rsLogin("SALT").value)
            If request.form("username") = rsLogin("login") and rsLogin("PASSWORD_HASH") = hashPassword Then
				lastModifyDt = cdate(rsLogin("PASSWORD_MODIFY_DT"))
                
				If dateDiff("d", lastModifyDt, sysDate) >= passwordExpiration Then
					passwordExpired = true
					session("CHANGE_PASSWORD_ID") = rsLogin("LOGIN_ID")
					if doAjax then
						session("LOGIN_ID") = rsLogin("LOGIN_ID")
					end if
				else
					session("LOGIN_ID") = rsLogin("LOGIN_ID")
					loginName = rsLogin("LOGIN")
					loginFirstName = rsLogin("FIRST_NAME")
					loginLastName = rsLogin("LAST_NAME")
					loginGroup = rsLogin("GROUP_TY_CD")

					session("LoginGroup") = loginGroup

					session("LoginName") = loginName

					if rsLogin("ACCOUNT_LOCKED")&"" = "Y" then
						accountLocked = true
					else
						accountLocked = false
					end if

					if session("GOOGLE_EMAIL")&"" <> "" then
						cSqlForm.SqlStr = "update login set google_email = ? where login_id = ?"
						cSqlForm.AddParam(session("GOOGLE_EMAIL"))
						cSqlForm.AddParam(rsLogin("LOGIN_ID"))
						cSqlForm.Execute
					end if
				end if
			end if
		end if
		rsLogin.close
		set rsLogin = nothing
		set cSqlForm = nothing

		if doAjax then
			response.clear
		    Response.ContentType = "application/json; charset=utf-8"
		end if
					
		if passwordExpired then
			if doAjax then
				response.Status = "401 - Unauthorized"
				response.write "Password expired"
			else
				redirectForm("/www/change_password.asp?return=" & request.servervariables("SCRIPT_NAME"))
			end if
		elseif session("LOGIN_ID")&"" = "" then
            If CStr(session("LOGIN_ATTEMPTS")) = "" Then
                session("LOGIN_ATTEMPTS") = 0
            End If
			session("LOGIN_ATTEMPTS") = session("LOGIN_ATTEMPTS") + 1
			if cint(session("LOGIN_ATTEMPTS")) >= maxAttempts then
				if doAjax then
					response.Status = "401 - Maximum login attempts exceeded.  Account locked."
				else
					cForm.AddErrorMsg("Maximum login attempts exceeded.  Account is locked.")
				end if
				set cSqlForm = new cRunSql
				cSqlForm.conn = connOracle
				cSqlForm.SqlStr = "update login set account_locked = 'Y' where login = ?"
				cSqlForm.AddParam(request.form("username"))
				cSqlForm.Execute
				set cSqlForm = nothing
			else
				if doAjax then
					response.Status = "401 - Invalid credentials"
				else
					if cRenew.ConfigOption("login_fail_msg")&"" <> "" then
						cForm.ErrorMsg = cRenew.ConfigOption("login_fail_msg")
					else
						cForm.addErrorMsg("No record was found with those credentials.")
					end if
				end if
			end if
		elseif accountLocked then
			if doAjax then
				response.Status = "403 - Forbidden"
				response.write "Account locked"
			else
				cForm.AddErrorMsg("Account is locked.")
			end if
		else
            session("LOGIN_ATTEMPTS") = 0
			set cSqlForm = new cRunSql
			cSqlForm.conn = connOracle
			cSqlForm.sqlStr = "select membership_id from member_detail m, org_column_detail o " &_
				"where o.column_id = m.column_id " &_
				"and o.form_field_type = 'SYSUSER' " &_
				"and m.data = (select login from login where login_id = ?) "
			cSqlForm.AddParam(session("LOGIN_ID"))
			set rs = cSqlForm.Execute
			if rs.eof then
			else
				session("MEMBERSHIP_ID") = rs("MEMBERSHIP_ID")
			end if
			rs.close
			set rs = nothing
			set cSqlForm = nothing
			if doAjax then
			    Response.Status = "200 - Success"
				response.write "{""userdetails"": { ""userName"": """ & loginName & """,""firstName"": """ & loginFirstName & """,""lastName"": """ & loginLastName & """,""login_id"": """ & session("LOGIN_ID") & """,""group"": """ & loginGroup & """ }}"
			end if
		end if
		if doAjax then
			closeAll()
			response.end()
		end if
	end if
end sub

sub existingRecordCheck

formLog("start existingRecordCheck")

dim rsLogin, saveSql, rsField, rsMembership, parentType, loginPrefix, loginDate, formVal
dim checkType, checkCount

checkType = cRenew.ConfigOption("existing_record_check_type")

parentType = getParentType(appRenewCd)

if parentType & "" <> "" then
	set cSqlForm = new cRunSql
	cSqlForm.conn = connOracle
	cSqlForm.SqlStr = "select app_renew_type from app_renew_fields where app_renew_type = ? and page = ? and column_name is not null and org_sub_ty_cd = ? and existing_record_field = 'Y'"
	cSqlForm.AddParam(appRenewCd)
	cSqlForm.AddParam(pageCd)
	cSqlForm.AddParam(parentType)
	set rsLogin = cSqlForm.Execute
	if rsLogin.eof then
		parentType = ""
	end if
	rsLogin.close
	set rsLogin = nothing
	set cSqlForm = nothing
end if

set cSqlForm = new cRunSql
cSqlForm.conn = connOracle
cSqlForm.SqlStr = "select app_renew_fields.COLUMN_NAME " &_
		",app_renew_fields.COLUMN_TYPE " &_
		",app_renew_fields.LOGIN_COMPARE " &_
		",app_renew_fields.numeric_only " &_
		",app_renew_fields.ORG_SUB_TY_CD " &_
		",app_renew_fields.FIELD_NAME " &_
		",(select data_type from org_column_detail where org_sub_ty_cd = (select minor_code from code_value where code_class = 'APP_RENEW_TYPE_CD' and code_value = ?) and column_desc = app_renew_fields.column_name) db_data_type " &_
		"from app_renew_fields where app_renew_type = ? and existing_record_field = 'Y' and column_name is not null and (org_sub_ty_cd is null "
cSqlForm.AddParam(appRenewCd)
cSqlForm.AddParam(appRenewCd)
if parentType &"" <> "" then
	cSqlForm.SqlStr = cSqlForm.SqlStr & " or org_sub_ty_cd = ?"
	cSqlForm.AddParam(parentType)
end if
cSqlForm.SqlStr = cSqlForm.SqlStr & ")"
set rsLogin = cSqlForm.Execute
set cSqlForm = nothing

set cSqlForm = new cRunSql
cSqlForm.conn = connOracle
if parentType &"" = "" then
	cSqlForm.SqlStr = "select membership_id,org_relate_id from membership where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = (select minor_code from code_value where code_class = 'APP_RENEW_TYPE_CD' and code_value = ?) "
else
	cSqlForm.SqlStr = "select child.membership_id,child.org_relate_id from membership child, membership parent " &_
		"where to_char(child.org_relate_id) = to_char(parent.membership_id) " &_
		"and child.m_stat_cd = 'ACTIVE' and child.org_sub_ty_cd = (select minor_code from code_value where code_class = 'APP_RENEW_TYPE_CD' and code_value = ?) " &_
		"and parent.m_stat_cd = 'ACTIVE' and parent.org_sub_ty_cd = ? "
end if
cSqlForm.AddParam(appRenewCd)
if parentType &"" <> "" then
	cSqlForm.AddParam(parentType)
end if
cSqlForm.DisplayErrors = true

checkCount = 0
do while not rsLogin.eof
	if session(rsLogin("FIELD_NAME"))&"" <> "" then
		checkCount = checkCount + 1
		if checkCount > 1 then
			cSqlForm.AddSqlStr(checkType & " ")
		else
			cSqlForm.AddSqlStr("AND (")
		end if
		if parentType & "" = "" then
			if rsLogin("COLUMN_TYPE") = "MEMBER" then
				if rsLogin("COLUMN_NAME") = "BIRTH_DT" or rsLogin("COLUMN_NAME") = "LIC_EXPIRE_DT" then
					loginCompare = "="
					if rsLogin("LOGIN_COMPARE")&"" <> "" then
						loginCompare = rsLogin("LOGIN_COMPARE")
					end if
					cSqlForm.SqlStr = cSqlForm.SqlStr & rsLogin("COLUMN_NAME") & " " & loginCompare & " to_date(?,'mm/dd/yyyy') "
				elseif rsLogin("COLUMN_NAME") = "SOC_SECURITY" and len(session(rsLogin("FIELD_NAME"))) = 4 then
					cSqlForm.SqlStr = cSqlForm.SqlStr & "substr(" & rsLogin("COLUMN_NAME") & ",-4) = ? "
				elseif rsLogin("numeric_only") = "Y" or rsLogin("COLUMN_NAME") = "SOC_SECURITY" or rsLogin("COLUMN_NAME") = "PHONE_DAY" or rsLogin("COLUMN_NAME") = "PHONE_EV"  or rsLogin("COLUMN_NAME") = "ZIP" then
					cSqlForm.SqlStr = cSqlForm.SqlStr & "REGEXP_REPLACE(" & rsLogin("COLUMN_NAME") & ", '[^0-9]+', '') = ? "
				else
					cSqlForm.SqlStr = cSqlForm.SqlStr & "upper(" & rsLogin("COLUMN_NAME") & ") = upper(?) "
				end if
			else
				if rsLogin("db_data_type") = "DATE" then
					loginCompare = "="
					if rsLogin("LOGIN_COMPARE")&"" <> "" then
						loginCompare = rsLogin("LOGIN_COMPARE")
					end if
					cSqlForm.SqlStr = cSqlForm.SqlStr & "to_date(nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = membership.membership_id and column_desc = (membership_id,'" & rsLogin("COLUMN_NAME") & "'),''),'mm/dd/yyyy') " & loginCompare & " to_date(?,'mm/dd/yyyy') "
				elseif rsLogin("numeric_only") = "Y" then
					cSqlForm.SqlStr = cSqlForm.SqlStr & "REGEXP_REPLACE(nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = membership.membership_id and column_desc = '" & rsLogin("COLUMN_NAME") & "'),''), '[^0-9]+', '') = ? "
				else
					cSqlForm.SqlStr = cSqlForm.SqlStr & "upper(nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = membership.membership_id and column_desc = '" & rsLogin("COLUMN_NAME") & "'),'')) = upper(?) "
				end if
			end if
		else
			if rsLogin("ORG_SUB_TY_CD")&"" = "" then
				loginPrefix = "child"
			else
				loginPrefix = "parent"
			end if
			if rsLogin("COLUMN_TYPE") = "MEMBER" then
				if rsLogin("COLUMN_NAME") = "BIRTH_DT" or rsLogin("COLUMN_NAME") = "LIC_EXPIRE_DT" then
					loginCompare = "="
					if rsLogin("LOGIN_COMPARE")&"" <> "" then
						loginCompare = rsLogin("LOGIN_COMPARE")
					end if
					cSqlForm.SqlStr = cSqlForm.SqlStr & loginPrefix & "." & rsLogin("COLUMN_NAME") & " " & loginCompare & " to_date(?,'mm/dd/yyyy') "
				elseif rsLogin("COLUMN_NAME") = "SOC_SECURITY" and len(session(rsLogin("FIELD_NAME"))) = 4 then
					cSqlForm.SqlStr = cSqlForm.SqlStr & "substr(" & loginPrefix & "." & rsLogin("COLUMN_NAME") & ",-4) = ? "
				elseif rsLogin("numeric_only") = "Y" or rsLogin("COLUMN_NAME") = "SOC_SECURITY" or rsLogin("COLUMN_NAME") = "PHONE_DAY" or rsLogin("COLUMN_NAME") = "PHONE_EV"  or rsLogin("COLUMN_NAME") = "ZIP" then
					cSqlForm.SqlStr = cSqlForm.SqlStr & "REGEXP_REPLACE(" & loginPrefix & "." & rsLogin("COLUMN_NAME") & ", '[^0-9]+', '') = ? "
				else
					cSqlForm.SqlStr = cSqlForm.SqlStr & "upper(" & loginPrefix & "." & rsLogin("COLUMN_NAME") & ") = upper(?) "
				end if
			else
				cSqlForm.SqlStr = cSqlForm.SqlStr & "AND upper(nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = " & loginPrefix & "." & "membership_id and column_desc = '" & rsLogin("COLUMN_NAME") & "'),'')) = upper(?) "
			end if
		end if
		if rsLogin("COLUMN_TYPE") = "MEMBER" and (rsLogin("COLUMN_NAME") = "BIRTH_DT" or rsLogin("COLUMN_NAME") = "LIC_EXPIRE_DT") then
			formVal = session(rsLogin("FIELD_NAME"))
			if isDate(formVal) then
				loginDate = mmddyyyy(formVal)
			else
				loginDate = "01/01/1800"
			end if
			cSqlForm.AddParam(loginDate)
		else
			if rsLogin("db_data_type") = "DATE" then
				formVal = session(rsLogin("FIELD_NAME"))
				if isDate(formVal) then
					loginDate = mmddyyyy(formVal)
				else
					loginDate = "01/01/1800"
				end if
				cSqlForm.AddParam(loginDate)
			elseif rsLogin("COLUMN_NAME") = "SOC_SECURITY" or rsLogin("COLUMN_NAME") = "PHONE_DAY" or rsLogin("COLUMN_NAME") = "PHONE_EV"  or rsLogin("COLUMN_NAME") = "ZIP" then
				cSqlForm.AddParam(numericOnly(session(rsLogin("FIELD_NAME"))))
			else
				cSqlForm.AddParam(session(rsLogin("FIELD_NAME")))
			end if
		end if
	end if
	rsLogin.MoveNext
loop
rsLogin.close
set rsLogin = nothing

if checkCount > 0 then
	cSqlForm.AddSqlStr(")")
end if

formLog("existing record check sql = " & cSqlForm.SqlStr)
formLog("existing record check params = " & cSqlForm.Params)

set rsMembership = cSqlForm.Execute

if rsMembership.eof then
	formLog("No existing record found, continuing form")
	session("EXISTING_RECORD_ID") = "0"
else
	formLog("Found existing record " & rsMembership("MEMBERSHIP_ID"))
	session("EXISTING_RECORD_ID") = "1"
	session("MEMBERSHIP_ID") = rsMembership("MEMBERSHIP_ID")

	set cMember = new cMemberValuesO
	cMember.AddLoadCol("ORG_RELATE_ID")
	cMember.AddLoadCol("ORG_RELATE_ID_2")
	cMember.AddLoadCol("ORG_SUB_TY_CD")
	cMember.LoadFromMembershipID rsMembership("MEMBERSHIP_ID")
	if Application("ORG_ID") = "175" and cMember.Value("ORG_SUB_TY_CD") = "MEMBER" then
		session("PARENT_ID") = cMember.Value("ORG_RELATE_ID_2")
	else
		if cMember.Value("ORG_RELATE_ID") <> "" then
			if cMember.Value("ORG_RELATE_ID") = numericOnly(cMember.Value("ORG_RELATE_ID")) then
				session("PARENT_ID") = cMember.Value("ORG_RELATE_ID")
			end if
		end if
	end if
	set cMember = nothing

	cRenew.LoadDefaults
end if

rsMembership.close
set rsMembership = nothing

set cSqlForm = nothing

formLog("end existingRecordCheck")

end sub

function getPasswordHash(pass,salt)

	dim httpHash, urlHash, paramsHash

    urlHash = "ws.ebigpicture.com/hash.aspx"
    urlHash = "https://" & urlHash
    paramsHash = "action=hash&password=" & pass & "&salt=" & salt

    Set httpHash = CreateObject("MSXML2.XMLHTTP.6.0")
    httpHash.open "POST",urlHash,False
    httpHash.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
    httpHash.send paramsHash
    getPasswordHash = httpHash.responseText
    set httpHash = nothing

end function

sub loginExist

formLog("start loginExist")

dim rsLogin, saveSql, rsField, rsMembership, parentType, grandparentType, loginPrefix, loginDate, formVal, loginCompare
dim hashFieldDB, hashFieldForm, cCol, arrCols, hashFuncCol, hashFieldFormValue, hashFieldRecTyp, arrChildTypes, rsChildTypes, childType

childType = ""
parentType = getParentType(appRenewCd)
grandparentType = getGrandparentType(appRenewCd)

set cSqlForm = new cRunSql
cSqlForm.DisplayErrors = true
cSqlForm.Conn = connOracle
cSqlForm.SqlStr = "select code_value from code_value where code_class = 'ORG_SUB_TY_CD' and inactive = 'N' and minor_code = ?"
cSqlForm.AddParam(formRecType)
set rsChildTypes = cSqlForm.Execute
set cSqlForm = nothing
do while not rsChildTypes.eof and childType = ""
	formLog("looking for fields of childType = " & rsChildTypes("code_value"))
	set cSqlForm = new cRunSql
	cSqlForm.conn = connOracle
	cSqlForm.SqlStr = "select app_renew_type from app_renew_fields where app_renew_type = ? and page = ? and column_name is not null and org_sub_ty_cd = ?"
	cSqlForm.AddParam(appRenewCd)
	cSqlForm.AddParam(pageCd)
	cSqlForm.AddParam(rsChildTypes("code_value"))
	set rsLogin = cSqlForm.Execute
	if rsLogin.eof then
		formLog("eof")
		childType = ""
	else
		formLog("found")
		childType = rsChildTypes("code_value")
	end if
	rsLogin.close
	set rsLogin = nothing
	set cSqlForm = nothing
	rsChildTypes.MoveNext
loop
rsChildTypes.close
set rsChildTypes = nothing

if childType & "" <> "" then
	set cSqlForm = new cRunSql
	cSqlForm.conn = connOracle
	cSqlForm.SqlStr = "select app_renew_type from app_renew_fields where app_renew_type = ? and page = ? and column_name is not null and org_sub_ty_cd = ?"
	cSqlForm.AddParam(appRenewCd)
	cSqlForm.AddParam(pageCd)
	cSqlForm.AddParam(childType)
	set rsLogin = cSqlForm.Execute
	if rsLogin.eof then
		childType = ""
	end if
	rsLogin.close
	set rsLogin = nothing
	set cSqlForm = nothing
end if

if parentType & "" <> "" then
	set cSqlForm = new cRunSql
	cSqlForm.conn = connOracle
	cSqlForm.SqlStr = "select app_renew_type from app_renew_fields where app_renew_type = ? and page = ? and column_name is not null and org_sub_ty_cd = ?"
	cSqlForm.AddParam(appRenewCd)
	cSqlForm.AddParam(pageCd)
	cSqlForm.AddParam(parentType)
	set rsLogin = cSqlForm.Execute
	if rsLogin.eof then
		parentType = ""
	end if
	rsLogin.close
	set rsLogin = nothing
	set cSqlForm = nothing
end if

if grandparentType & "" <> "" then
	set cSqlForm = new cRunSql
	cSqlForm.conn = connOracle
	cSqlForm.SqlStr = "select app_renew_type from app_renew_fields where app_renew_type = ? and page = ? and column_name is not null and org_sub_ty_cd = ?"
	cSqlForm.AddParam(appRenewCd)
	cSqlForm.AddParam(pageCd)
	cSqlForm.AddParam(grandparentType)
	set rsLogin = cSqlForm.Execute
	if rsLogin.eof then
		grandparentType = ""
	else
		parentType = getParentType(appRenewCd)
	end if
	rsLogin.close
	set rsLogin = nothing
	set cSqlForm = nothing
end if

if childType <> "" then
	parentType = ""
	grandParentType = ""
end if

formLog("childType = " & childType)
formLog("parentType = " & parentType)
formLog("grandParentType = " & grandParentType)

set cSqlForm = new cRunSql
cSqlForm.conn = connOracle
cSqlForm.SqlStr = "select app_renew_fields.COLUMN_NAME " &_
		",app_renew_fields.COLUMN_TYPE " &_
		",app_renew_fields.LOGIN_COMPARE " &_
		",app_renew_fields.numeric_only " &_
		",app_renew_fields.ORG_SUB_TY_CD " &_
		",app_renew_fields.FIELD_NAME " &_
		",(select data_type from org_column_detail where org_sub_ty_cd = (select minor_code from code_value where code_class = 'APP_RENEW_TYPE_CD' and code_value = ?) and column_desc = app_renew_fields.column_name) db_data_type " &_
		",(select form_field_type from org_column_detail where org_sub_ty_cd = (select minor_code from code_value where code_class = 'APP_RENEW_TYPE_CD' and code_value = ?) and column_desc = app_renew_fields.column_name) db_field_type " &_
		",app_renew_fields.FIELD_TYPE " &_
		"from app_renew_fields where app_renew_type = ? and page = ? and column_name is not null and (org_sub_ty_cd is null "
cSqlForm.AddParam(appRenewCd)
cSqlForm.AddParam(appRenewCd)
cSqlForm.AddParam(appRenewCd)
cSqlForm.AddParam(pageCd)
if childType &"" <> "" then
	cSqlForm.SqlStr = cSqlForm.SqlStr & " or org_sub_ty_cd = ?"
	cSqlForm.AddParam(childType)
end if
if parentType &"" <> "" then
	cSqlForm.SqlStr = cSqlForm.SqlStr & " or org_sub_ty_cd = ?"
	cSqlForm.AddParam(parentType)
end if
if grandparentType &"" <> "" then
	cSqlForm.SqlStr = cSqlForm.SqlStr & " or org_sub_ty_cd = ?"
	cSqlForm.AddParam(grandparentType)
end if
cSqlForm.SqlStr = cSqlForm.SqlStr & ")"
set rsLogin = cSqlForm.Execute
set cSqlForm = nothing

if rsLogin.eof then
	formLog("No login fields configured")
	if isDev then
		cForm.ErrorMsg = "No login fields configured."
	else
		if cRenew.ConfigOption("login_fail_msg")&"" <> "" then
			cForm.ErrorMsg = cRenew.ConfigOption("login_fail_msg")
		else
			cForm.ErrorMsg = "No record was found with those credentials."
		end if
	end if
	rsLogin.close
	set rsLogin = nothing
	exit sub
end if

set cSqlForm = new cRunSql
cSqlForm.conn = connOracle
if childType <> "" then
	cSqlForm.SqlStr = "select membership.membership_id,membership.org_relate_id,membership.org_relate_id_2,membership.org_sub_ty_cd,child.membership_id child_id,child.org_sub_ty_cd child_type from membership, membership child, mor_finder " &_
		"where membership.membership_id = mor_finder.org_relate_id " &_
		"and child.membership_id = mor_finder.membership_id " &_
		"and membership.m_stat_cd = 'ACTIVE' and membership.org_sub_ty_cd = (select minor_code from code_value where code_class = 'APP_RENEW_TYPE_CD' and code_value = ?) " &_
		"and child.m_stat_cd = 'ACTIVE' and child.org_sub_ty_cd = ? "
elseif parentType &"" = "" then
	cSqlForm.SqlStr = "select membership_id,org_relate_id from membership where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = (select minor_code from code_value where code_class = 'APP_RENEW_TYPE_CD' and code_value = ?) "
elseif grandparentType&"" <> "" then
	cSqlForm.SqlStr = "select membership.membership_id,membership.org_relate_id,parent.org_relate_id grandparent_id,membership.org_sub_ty_cd from membership, membership parent, membership grandparent, mor_finder, mor_finder mor2 " &_
		"where membership.membership_id = mor_finder.membership_id " &_
		"and parent.membership_id = mor_finder.org_relate_id " &_
		"and parent.membership_id = mor2.membership_id " &_
		"and grandparent.membership_id = mor2.org_relate_id " &_
		"and membership.m_stat_cd = 'ACTIVE' and membership.org_sub_ty_cd = (select minor_code from code_value where code_class = 'APP_RENEW_TYPE_CD' and code_value = ?) " &_
		"and parent.m_stat_cd = 'ACTIVE' and parent.org_sub_ty_cd = ? " &_
		"and grandparent.m_stat_cd = 'ACTIVE' and grandparent.org_sub_ty_cd = ? "
else
	cSqlForm.SqlStr = "select membership.membership_id,membership.org_relate_id,membership.org_relate_id_2,membership.org_sub_ty_cd from membership, membership parent, mor_finder " &_
		"where membership.membership_id = mor_finder.membership_id " &_
		"and parent.membership_id = mor_finder.org_relate_id " &_
		"and membership.m_stat_cd = 'ACTIVE' and membership.org_sub_ty_cd = (select minor_code from code_value where code_class = 'APP_RENEW_TYPE_CD' and code_value = ?) " &_
		"and parent.m_stat_cd = 'ACTIVE' and parent.org_sub_ty_cd = ? "
end if
cSqlForm.AddParam(appRenewCd)
if childType &"" <> "" then
	cSqlForm.AddParam(childType)
end if
if parentType &"" <> "" then
	cSqlForm.AddParam(parentType)
end if
if grandparentType &"" <> "" then
	cSqlForm.AddParam(grandparentType)
end if
cSqlForm.DisplayErrors = true

do while not rsLogin.eof
	if request.form(rsLogin("FIELD_NAME"))&"" <> "" then
		hashFuncCol = false
		if parentType & "" = "" and childType & "" = "" then
			if rsLogin("COLUMN_TYPE") = "MEMBER" then
				if rsLogin("COLUMN_NAME") = "BIRTH_DT" or rsLogin("COLUMN_NAME") = "LIC_EXPIRE_DT" then
					loginCompare = "="
					if rsLogin("LOGIN_COMPARE")&"" <> "" then
						loginCompare = rsLogin("LOGIN_COMPARE")
					end if
					cSqlForm.SqlStr = cSqlForm.SqlStr & "AND " & rsLogin("COLUMN_NAME") & " " & loginCompare & " to_date(?,'mm/dd/yyyy') "
				elseif rsLogin("COLUMN_NAME") = "SOC_SECURITY" and len(request.form(rsLogin("FIELD_NAME"))) = 4 then
					cSqlForm.SqlStr = cSqlForm.SqlStr & "AND substr(" & rsLogin("COLUMN_NAME") & ",-4) = ? "
				elseif rsLogin("numeric_only") = "Y" then
					cSqlForm.SqlStr = cSqlForm.SqlStr & "AND REGEXP_REPLACE(" & rsLogin("COLUMN_NAME") & ", '[^0-9]+', '') = ? "
				elseif rsLogin("field_type") = "PASSWORD" then
					cSqlForm.SqlStr = cSqlForm.SqlStr & "AND " & rsLogin("COLUMN_NAME") & " = ? "
				else
					cSqlForm.SqlStr = cSqlForm.SqlStr & "AND upper(" & rsLogin("COLUMN_NAME") & ") = upper(?) "
				end if
			else
				if rsLogin("db_data_type") = "DATE" then
					loginCompare = "="
					if rsLogin("LOGIN_COMPARE")&"" <> "" then
						loginCompare = rsLogin("LOGIN_COMPARE")
					end if
					cSqlForm.SqlStr = cSqlForm.SqlStr & "AND to_date(nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = membership.membership_id and column_desc = (membership_id,'" & rsLogin("COLUMN_NAME") & "'),''),'mm/dd/yyyy') " & loginCompare & " to_date(?,'mm/dd/yyyy') "
				elseif rsLogin("numeric_only") = "Y" then
					cSqlForm.SqlStr = cSqlForm.SqlStr & "AND REGEXP_REPLACE(nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = membership.membership_id and column_desc = '" & rsLogin("COLUMN_NAME") & "'),''), '[^0-9]+', '') = ? "
				elseif rsLogin("db_field_type") = "PASSWORD_HASHED" then
					hashFieldDB = rsLogin("COLUMN_NAME")
					hashFieldForm = rsLogin("FIELD_NAME")
				else
					if request.form(rsLogin("FIELD_NAME")).count > 1 then
						cSqlForm.SqlStr = cSqlForm.SqlStr & " and ("
						for i = 1 to request.form(rsLogin("FIELD_NAME")).count
							if rsLogin("db_data_type") = "FUNCTION" then
								if rsLogin("field_type") = "SSN4" then
									cSqlForm.SqlStr = cSqlForm.SqlStr & " substr(upper(to_char(" & rsLogin("COLUMN_NAME") & "(membership.membership_id))),-4) = upper(?) "
								else
									cSqlForm.SqlStr = cSqlForm.SqlStr & " upper(to_char(" & rsLogin("COLUMN_NAME") & "(membership.membership_id))) = upper(?) "
								end if
							else
								cSqlForm.SqlStr = cSqlForm.SqlStr & " upper(nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = membership.membership_id and column_desc = '" & rsLogin("COLUMN_NAME") & "'),'')) = upper(?) "
							end if
							if i < request.form(rsLogin("FIELD_NAME")).count then
								cSqlForm.SqlStr = cSqlForm.SqlStr & " or "
							end if
						next
						cSqlForm.SqlStr = cSqlForm.SqlStr & ") "
					else
						if rsLogin("db_data_type") = "FUNCTION" then
							if rsLogin("field_type") = "SSN4" then
								cSqlForm.SqlStr = cSqlForm.SqlStr & "AND substr(upper(to_char(" & rsLogin("COLUMN_NAME") & "(membership.membership_id))),-4) = upper(?) "
							else
								set cCol = new cOrgColumnDetailO
								cCol.LoadFromColumnDesc rsLogin("COLUMN_NAME")
								if replace(cCol.OrgDisplayColumns&"",", ","") <> "" then
									arrCols = split(replace(cCol.OrgDisplayColumns,", ",""),",")
									if ubound(arrCols) = 0 then
										set cCol = nothing
										set cCol = new cOrgColumnDetailO
										cCol.LoadFromColumnDesc replace(replace(arrCols(0),"DETAIL::",""),"MEMBER::","")
										if cCol.FieldType = "PASSWORD_HASHED" then
											hashFuncCol = true
											hashFieldDB = rsLogin("COLUMN_NAME")
											hashFieldForm = rsLogin("FIELD_NAME")
											hashFieldRecTyp = cCol.OrgSubTyCd
										end if
									end if
								end if
								set cCol = nothing
								if hashFuncCol then
								else
									cSqlForm.SqlStr = cSqlForm.SqlStr & "AND upper(to_char(" & rsLogin("COLUMN_NAME") & "(membership.membership_id))) = upper(?) "
								end if
							end if
						elseif rsLogin("field_type") = "PASSWORD" then
							cSqlForm.SqlStr = cSqlForm.SqlStr & "AND nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = membership.membership_id and column_desc = '" & rsLogin("COLUMN_NAME") & "'),'') = ? "
						else
							cSqlForm.SqlStr = cSqlForm.SqlStr & "AND upper(nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = membership.membership_id and column_desc = '" & rsLogin("COLUMN_NAME") & "'),'')) = upper(?) "
						end if
					end if
				end if
			end if
		else
			if rsLogin("ORG_SUB_TY_CD")&"" = "" then
				loginPrefix = "membership"
			elseif rsLogin("ORG_SUB_TY_CD") = childType then
				loginPrefix = "child"
			elseif rsLogin("ORG_SUB_TY_CD") = grandparentType then
				loginPrefix = "grandparent"
			else
				loginPrefix = "parent"
			end if
			if rsLogin("COLUMN_TYPE") = "MEMBER" then
				if rsLogin("COLUMN_NAME") = "BIRTH_DT" or rsLogin("COLUMN_NAME") = "LIC_EXPIRE_DT" then
					loginCompare = "="
					if rsLogin("LOGIN_COMPARE")&"" <> "" then
						loginCompare = rsLogin("LOGIN_COMPARE")
					end if
					cSqlForm.SqlStr = cSqlForm.SqlStr & "AND " & loginPrefix & "." & rsLogin("COLUMN_NAME") & " " & loginCompare & " to_date(?,'mm/dd/yyyy') "
				elseif rsLogin("COLUMN_NAME") = "SOC_SECURITY" and len(request.form(rsLogin("FIELD_NAME"))) = 4 then
					cSqlForm.SqlStr = cSqlForm.SqlStr & "AND substr(" & loginPrefix & "." & rsLogin("COLUMN_NAME") & ",-4) = ? "
				else
					cSqlForm.SqlStr = cSqlForm.SqlStr & "AND upper(" & loginPrefix & "." & rsLogin("COLUMN_NAME") & ") = upper(?) "
				end if
			else
				if request.form(rsLogin("FIELD_NAME")).count > 1 then
					cSqlForm.SqlStr = cSqlForm.SqlStr & " and ("
					for i = 1 to request.form(rsLogin("FIELD_NAME")).count
						if rsLogin("db_data_type") = "FUNCTION" then
							if rsLogin("field_type") = "SSN4" then
								cSqlForm.SqlStr = cSqlForm.SqlStr & " substr(upper(to_char(" & rsLogin("COLUMN_NAME") & "(" & loginPrefix & ".membership_id))),-4) = upper(?) "
							else
								cSqlForm.SqlStr = cSqlForm.SqlStr & " upper(to_char(" & rsLogin("COLUMN_NAME") & "(" & loginPrefix & ".membership_id))) = upper(?) "
							end if
						else
							cSqlForm.SqlStr = cSqlForm.SqlStr & " upper(nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = " & loginPrefix & "." & "membership_id and column_desc = '" & rsLogin("COLUMN_NAME") & "'),'')) = upper(?) "
						end if
						if i < request.form(rsLogin("FIELD_NAME")).count then
							cSqlForm.SqlStr = cSqlForm.SqlStr & " or "
						end if
					next
					cSqlForm.SqlStr = cSqlForm.SqlStr & ") "
				elseif rsLogin("db_field_type") = "PASSWORD_HASHED" then
					hashFieldDB = rsLogin("COLUMN_NAME")
					hashFieldForm = rsLogin("FIELD_NAME")
				else
					if rsLogin("field_type") = "SSN4" then
						cSqlForm.SqlStr = cSqlForm.SqlStr & "AND substr(upper(to_char(" & rsLogin("COLUMN_NAME") & "(" & loginPrefix & ".membership_id))),-4) = upper(?) "
					elseif rsLogin("db_data_type") = "FUNCTION" then
						cSqlForm.SqlStr = cSqlForm.SqlStr & "AND upper(to_char(" & rsLogin("COLUMN_NAME") & "(" & loginPrefix & ".membership_id))) = upper(?) "
					else
						cSqlForm.SqlStr = cSqlForm.SqlStr & "AND upper(nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = " & loginPrefix & "." & "membership_id and column_desc = '" & rsLogin("COLUMN_NAME") & "'),'')) = upper(?) "
					end if
				end if
			end if
		end if
		if rsLogin("COLUMN_TYPE") = "MEMBER" and (rsLogin("COLUMN_NAME") = "BIRTH_DT" or rsLogin("COLUMN_NAME") = "LIC_EXPIRE_DT") then
			formVal = request.form(rsLogin("FIELD_NAME"))
			if isDate(formVal) then
				loginDate = mmddyyyy(formVal)
			else
				loginDate = "01/01/1800"
			end if
			cSqlForm.AddParam(loginDate)
		else
			if request.form(rsLogin("FIELD_NAME")).count > 1 then
				for i = 1 to request.form(rsLogin("FIELD_NAME")).count
					cSqlForm.AddParam(request.form(rsLogin("FIELD_NAME"))(i))
				next
			else
				if rsLogin("db_data_type") = "DATE" then
					formVal = request.form(rsLogin("FIELD_NAME"))
					if isDate(formVal) then
						loginDate = mmddyyyy(formVal)
					else
						loginDate = "01/01/1800"
					end if
					cSqlForm.AddParam(loginDate)
				elseif rsLogin("db_field_type") = "PASSWORD_HASHED" then
				elseif hashFuncCol then
				else
					cSqlForm.AddParam(request.form(rsLogin("FIELD_NAME")))
				end if
			end if
		end if
	else
		if doAjax then
			formLog("Login field " & rsLogin("FIELD_NAME") & " empty")
			rsLogin.close
			set rsLogin = nothing
			set cSqlForm = nothing
			exit sub
		end if
	end if
	rsLogin.MoveNext
loop
rsLogin.close
set rsLogin = nothing

if cRenew.ConfigOption("login_sql_save_id")&"" <> "" then
	dim cSqlSave, rsSqlSave, sqlSave
	set cSqlSave = new cRunSql
	cSqlSave.Conn = connOracle
	cSqlSave.SqlStr = "select sql_string from sql_save where sql_save_id = ?"
	cSqlSave.AddParam(cRenew.ConfigOption("login_sql_save_id"))
	set rsSqlSave = cSqlSave.Execute
	if rsSqlSave.eof then
	else
		sqlSave = dynDate(rsSqlSave("sql_string"))
	end if 
	rsSqlSave.close
	set rsSqlSave = nothing
	set cSqlSave = nothing
	if parentType &"" = "" then
		cSqlForm.AddSqlStr(" and membership_id in (select membership_id from (" & sqlSave & "))")
	else
		cSqlForm.AddSqlStr(" and membership.membership_id in (select membership_id from (" & sqlSave & "))")
	end if
end if

cSqlForm.SqlStr = FunctionReplace(cSqlForm.SqlStr)

formLog("login sql = " & cSqlForm.SqlStr)
formLog("login params = " & cSqlForm.Params)

set rsMembership = cSqlForm.Execute
set cSqlForm = nothing

if rsMembership.eof then
	if cRenew.ConfigOption("login_type") = "EXIST_IF" then
		formLog("No existing record found, continuing form")
	else
		formLog("No record found")
		if cRenew.ConfigOption("login_fail_msg")&"" <> "" then
			cForm.ErrorMsg = cRenew.ConfigOption("login_fail_msg")
		else
			cForm.addErrorMsg("No record was found with those credentials.")
		end if
	end if
else
	if hashFieldDB&"" <> "" and hashFieldForm&"" <> "" then
		formLog("HASH Found existing record " & rsMembership("MEMBERSHIP_ID"))
		set cMember = new cMemberValuesO
		cMember.AddLoadCol(hashFieldDB)
	    cMember.LoadFromMembershipID rsMembership("MEMBERSHIP_ID")
		if left(hashFieldDB,13) = "Child_Record_" then
			set cSqlForm = new cRunSql
			cSqlForm.conn = connOracle
			cSqlForm.SqlStr ="select f.membership_id from membership m, mor_finder f, membership c where m.membership_id = f.org_relate_id and f.membership_id = c.membership_id and c.m_stat_cd = 'ACTIVE' and c.org_sub_ty_cd = ? and m.membership_id = ?"
			cSqlForm.AddParam(hashFieldRecTyp)
			cSqlForm.AddParam(rsMembership("MEMBERSHIP_ID"))
			set rsMembership = cSqlForm.Execute
			set cSqlForm = nothing
			if rsMembership.eof then
			else
				hashFieldFormValue = getPassWordHash(request.form(hashFieldForm),rsMembership("MEMBERSHIP_ID"))
			end if
		else
			hashFieldFormValue = getPassWordHash(request.form(hashFieldForm),rsMembership("MEMBERSHIP_ID"))
		end if
		if hashFieldFormValue = cMember.Value(hashFieldDB) then
			formLog("HASH password match")
		else
			formLog("HASH password does not match")
			if cRenew.ConfigOption("login_fail_msg")&"" <> "" then
				cForm.ErrorMsg = cRenew.ConfigOption("login_fail_msg")
			else
				cForm.addErrorMsg("No record was found with those credentials.")
			end if
			set cMember = nothing
			rsMembership.close
			set rsMembership = nothing
			exit sub
		end if
		set cMember = nothing
	end if

	formLog("Found record " & rsMembership("MEMBERSHIP_ID"))

	set cSqlForm = new cRunSql
	cSqlForm.conn = connOracle
	cSqlForm.SqlStr = "select app_renew_type from app_renew where app_renew_type = ? and app_renew_status = 'PENDING' and membership_id = ?"
	cSqlForm.AddParam(appRenewCd)
	cSqlForm.AddParam(rsMembership("MEMBERSHIP_ID"))
	set rsPending = cSqlForm.Execute
	set cSqlForm = nothing
	if rsPending.eof then
		session("PENDING_FORM") = "0"
	else
		session("PENDING_FORM") = "1"
	end if
	rsPending.close
	set rsPending = nothing

	set cSqlForm = new cRunSql
	cSqlForm.conn = connOracle
	cSqlForm.SqlStr = "select app_renew_type from app_renew where app_renew_type = ? and app_renew_status = 'COMPLETED' and membership_id = ?"
	cSqlForm.AddParam(appRenewCd)
	cSqlForm.AddParam(rsMembership("MEMBERSHIP_ID"))
	set rsPending = cSqlForm.Execute
	set cSqlForm = nothing
	if rsPending.eof then
		session("COMPLETED_FORM") = "0"
	else
		session("COMPLETED_FORM") = "1"
	end if
	rsPending.close
	set rsPending = nothing

	if cRenew.ConfigOption("login_await_check") = "Y" then
		set cAwait = new cAwaitingUpdatesO
		if cAwait.HasUpdates(rsMembership("MEMBERSHIP_ID")) then
			if cRenew.ConfigOption("login_await_msg")&"" <> "" then
				cForm.ErrorMsg = cRenew.ConfigOption("login_await_msg")
			else
				cForm.ErrorMsg = "Your record is under review. Please contact the Board if you have questions."
			end if
		else
			session("MEMBERSHIP_ID") = cStr(rsMembership("MEMBERSHIP_ID"))
		end if
		set cAwait = nothing
	else
		session("MEMBERSHIP_ID") = cStr(rsMembership("MEMBERSHIP_ID"))
	end if
	if session("MEMBERSHIP_ID") <> "" and childType & "" <> "" then
		session("CHILD_ID") = rsMembership("child_id")
		session("CHILD_TYPE") = rsMembership("child_type")
	end if
	if session("MEMBERSHIP_ID") <> "" and parentType & "" <> "" then
		if Application("ORG_ID") = "175" and rsMembership("ORG_SUB_TY_CD") = "MEMBER" then
			session("PARENT_ID") = cstr(rsMembership("ORG_RELATE_ID_2"))
		else
			session("PARENT_ID") = cstr(rsMembership("ORG_RELATE_ID"))
		end if
		if cRenew.ConfigOption("login_await_check") = "Y" then
			set cAwait = new cAwaitingUpdatesO
			if cAwait.HasUpdates(session("PARENT_ID")) then
				if cRenew.ConfigOption("login_await_msg")&"" <> "" then
					cForm.ErrorMsg = cRenew.ConfigOption("login_await_msg")
				else
					cForm.ErrorMsg = "Your record is under review. Please contact the Board if you have questions."
				end if
				session("MEMBERSHIP_ID") = ""
				session("PARENT_ID") = ""
			end if
			set cAwait = nothing
		end if
	end if
	if session("MEMBERSHIP_ID") <> "" and grandparentType & "" <> "" then
		session("GRANDPARENT_ID") = cstr(rsMembership("GRANDPARENT_ID"))
		if cRenew.ConfigOption("login_await_check") = "Y" then
			set cAwait = new cAwaitingUpdatesO
			if cAwait.HasUpdates(session("GRANDPARENT_ID")) then
				if cRenew.ConfigOption("login_await_msg")&"" <> "" then
					cForm.ErrorMsg = cRenew.ConfigOption("login_await_msg")
				else
					cForm.ErrorMsg = "Your record is under review. Please contact the Board if you have questions."
				end if
				session("MEMBERSHIP_ID") = ""
				session("PARENT_ID") = ""
			end if
			set cAwait = nothing
		end if
	end if
end if

rsMembership.close
set rsMembership = nothing

set cSqlForm = nothing

session("AWAITING_UPDATES") = "0"
set cAwait = new cAwaitingUpdatesO
if cAwait.HasUpdates(session("MEMBERSHIP_ID")) then
	session("AWAITING_UPDATES") = "1"
end if
set cAwait = nothing

formLog("end loginExist")

end sub

sub loginSaved

formLog("start loginSaved")

dim rsSaved, saveSql, cSqlSaved, rsField, cSqlTmp, tmpDate, parentType, grandparentType, arrSaved, idxSave, arrTmp, tmpTbl, arrTmpTbl, rsSaved2, savedFileName, arrFilesSaved
dim saveID

parentType = getParentType(appRenewCd)
grandparentType = getGrandparentType(appRenewCd)

set cSqlTmp = new cRunSql
cSqlTmp.Conn = connOracle

set cSqlForm = new cRunSql
cSqlForm.Conn = connOracle
cSqlForm.SqlStr = "select FIELD_NAME,RELATED_FIELD from app_renew_fields where app_renew_type = ? and page = ? and related_field is not null"
cSqlForm.AddParam(appRenewCd)
cSqlForm.AddParam(pageCd)
set rsSaved = cSqlForm.Execute
set cSqlForm = nothing

if rsSaved.eof then
	cForm.ErrorMsg = "No saved record was found with those credentials."
	rsSaved.close
	set rsSaved = nothing
	exit sub
else
	arrTmp = rsSaved.GetRows
	redim arrSaved(4,ubound(arrTmp,2))
	for idxSave = lbound(arrTmp,2) to ubound(arrTmp,2)
		arrSaved(0,idxSave) = arrTmp(0,idxSave)
		arrSaved(1,idxSave) = arrTmp(1,idxSave)
		set cSqlSaved = new cRunSql
		cSqlSaved.Conn = connOracle
		cSqlSaved.SqlStr = "select COLUMN_NAME,FIELD_NAME,FIELD_TYPE from app_renew_fields where app_renew_fields_id = ?"
		cSqlSaved.AddParam(arrSaved(1,idxSave))
		set rsField = cSqlSaved.execute
		if rsField.eof then
		else
			arrSaved(2,idxSave) = rsField("COLUMN_NAME")
			arrSaved(3,idxSave) = rsField("FIELD_NAME")
			arrSaved(4,idxSave) = rsField("FIELD_TYPE")
		end if
		rsField.close
		set rsField = nothing
		set cSqlSaved = nothing
	next
	set cSqlForm = new cRunSql
	cSqlForm.Conn = connOracle
	cSqlForm.SqlStr = "select a.app_renew_id,s.create_dt submit_dt from app_renew a, app_renew_page_submit s  "
	for idxSave = lbound(arrSaved,2) to ubound(arrSaved,2)
		if arrSaved(2,idxSave) = "BIRTH_DT" or arrSaved(4,idxSave) = "DATE" then
			tmpTbl = "TEMP_"& getSequence("seq_temp_table")

			if isArray(arrTmpTbl) then
				redim preserve arrTmpTbl(idxSave)
			else
				redim arrTmpTbl(idxSave)
			end if
			arrTmpTbl(idxSave) = tmpTbl

			set cSqlTmp = new cRunSql
			cSqlTmp.Conn = connOracle
			cSqlTmp.SqlStr = "create table " & tmpTbl & " (app_renew_id number, data date)"
			cSqlTmp.Execute
			set cSqlTmp = nothing
			cSqlForm.AddSqlStr("," & tmpTbl & " d" & idxSave)
		else
			cSqlForm.AddSqlStr(",app_renew_detail d" & idxSave)
		end if
	next
	cSqlForm.AddSqlStr(" where s.app_renew_id = a.app_renew_id and a.app_renew_type = ? and a.app_renew_status = 'SAVED' ")
	cSqlForm.AddParam(appRenewCd)
	for idxSave = lbound(arrSaved,2) to ubound(arrSaved,2)
		cSqlForm.AddSqlStr("and d" & idxSave & ".app_renew_id = a.app_renew_id ")
	next
	for idxSave = lbound(arrSaved,2) to ubound(arrSaved,2)
		if arrSaved(2,idxSave) = "BIRTH_DT" or arrSaved(4,idxSave) = "DATE" then
			set cSqlTmp = new cRunSql
			cSqlTmp.Conn = connOracle
			cSqlTmp.SqlStr = "select APP_RENEW_ID,DATA from app_renew_detail where upper(field_name) = upper(?) and data is not null and app_renew_id in (select app_renew_id from app_renew where app_renew_status in ('SAVED','PROGRESS') and app_renew_type = ?)"
			cSqlTmp.AddParam(arrSaved(3,idxSave))
			cSqlTmp.AddParam(appRenewCd)
			set rsTmp = cSqlTmp.Execute
			do while not rsTmp.Eof
				if isDate(rsTmp("DATA")) then
					tmpDate = mmddyyyy(rsTmp("DATA"))
					if cdate(tmpDate) = cdate("12/30/1899") then
					else
						cSqlTmp.SqlStr = "insert into " & arrTmpTbl(idxSave) & " values (?,to_date(?,'mm/dd/yyyy'))"
						cSqlTmp.AddParam(rsTmp("APP_RENEW_ID"))
						cSqlTmp.AddParam(tmpDate)
						cSqlTmp.Execute
					end if
				end if
				rsTmp.MoveNext
			loop
			rsTmp.close
			set rsTmp = nothing
			set cSqlTmp = nothing
			cSqlForm.AddSqlStr("and to_date(to_char(d" & idxSave & ".data,'mm/dd/yyyy'),'mm/dd/yyyy') = to_date(?,'mm/dd/yyyy') ")
			cSqlForm.AddParam(mmddyyyy(request.form(arrSaved(0,idxSave))))
		elseif (instr(ucase(arrSaved(3,idxSave)),"SSN") or arrSaved(2,idxSave) = "SOC_SECURITY") and len(request.form(arrSaved(0,idxSave))) = 4 then
			cSqlForm.SqlStr = cSqlForm.SqlStr & "and upper(d" & idxSave & ".field_name) = upper(?) and substr(d" & idxSave & ".data,-4) = ? "
			cSqlForm.AddParam(arrSaved(3,idxSave))
			cSqlForm.AddParam(request.form(arrSaved(0,idxSave)))
		else
			cSqlForm.SqlStr = cSqlForm.SqlStr & "and upper(d" & idxSave & ".field_name) = upper(?) and upper(d" & idxSave & ".data) = upper(?) "
			cSqlForm.AddParam(arrSaved(3,idxSave))
			cSqlForm.AddParam(request.form(arrSaved(0,idxSave)))
		end if
	next
	cSqlForm.AddSqlStr(" order by submit_dt desc")

	formLog("login sql = " & cSqlForm.SqlStr)
	formLog("login params = " & cSqlForm.Params)

	set rsSaved = cSqlForm.Execute
	set cSqlForm = nothing

	if rsSaved.eof then
		formLog("no saved record found")
		cForm.ErrorMsg = "No saved record was found with those credentials."
		rsSaved.close
		set rsSaved = nothing
	else
		saveID = rsSaved("APP_RENEW_ID")
		formLog("saved record " & saveID & " found")
		set cSqlForm = new cRunSql
		cSqlForm.Conn = connOracle
		cSqlForm.SqlStr = "select FIELD_NAME,DATA from app_renew_detail where app_renew_id  = ? and upper(field_name) not in ("
		cSqlForm.AddParam(saveID)
		for i = lbound(arrSysVars) to ubound(arrSysVars)
			cSqlForm.AddSqlStr("upper('" & arrSysVars(i) & "')")
			if i < ubound(arrSysVars) then
				cSqlForm.AddSqlStr(",")
			end if
		next
		cSqlForm.AddSqlStr(") ")
		cSqlForm.AddSqlStr("and upper(field_name) not like '%HAS_VALUE' ")
		cSqlForm.AddSqlStr("and upper(field_name) not like 'ARR_FILES_%' ")
		cSqlForm.AddSqlStr("and data is not null ")
		cSqlForm.DisplayErrors = true
		set rsSaved = cSqlForm.Execute

		do while not rsSaved.eof
			if instr(lcase(rsSaved("FIELD_NAME")&""),"arrrelate") then
			else
				session(rsSaved("FIELD_NAME")) = cstr(rsSaved("DATA")&"")
			end if
			rsSaved.MoveNext
		loop
		rsSaved.Close
		set rsSaved = nothing

		redisKey = "FORM_" & Application("ORG_ID") & "_" & session("APP_RENEW_ID")
		set redisArr = Server.CreateObject("CacheManager")
		redisArr.SetDefaultExpiration 365,0,0,0
		redisArr.Init redisKey
		arrKeys = split(redisArr.GetCacheList,",")
		if isArray(arrKeys) then
			for x = lbound(arrKeys) to ubound(arrKeys)
				item = arrKeys(x)
				if isArray(redisArr(item)) Then
					if ucase(item) = "ARR_PAGES" or ucase(item) = "ARR_FIELDS" or ucase(item) = "ARR_FIELD_COLS" or ucase(item) = "ARR_COMBINE_FIELDS" or ucase(item) = "ARR_COMBINE_FIELDS_JS" or ucase(item) = "ARR_PAGE" or ucase(item) = "ARR_PAGE_COLS" or ucase(item) = "ARR_CONFIG" or ucase(item) = "ARR_CONFIG_COLS" or ucase(item) = "ARR_BATCH_PRODS" then
					else
						formLog("loading array " & item & " from saved redis cache")
						session(item) = redisArr(item)
					end if
				end if
			next
		end if
		set redisArr = nothing
	end if
end if 

if isArray(arrTmpTbl) then
	for idxSave = lbound(arrTmpTbl) to ubound(arrTmpTbl)
		if arrTmpTbl(idxSave)&"" <> "" then
			set cSqlTmp = new cRunSql
			cSqlTmp.Conn = connOracle
			cSqlTmp.SqlStr = "drop table " & arrTmpTbl(idxSave)
			cSqlTmp.Execute
			set cSqlTmp = nothing
		end if
	next
end if

formLog("end loginSaved")

if doAjax then
	response.clear
	Response.Status = "200 - Success"
	closeAll()
	response.end
end if

end sub

sub loadDebug

if not isNumeric(request.querystring("formID")) or not officeIP then
	Response.Status = "403 Forbidden"
	Response.End
end if

dim rsDebug, cSqlDebug, arrFilesDebug, debugFileName, idxDebug, rsDebug2, idxDebugFile, idxDebug2, arrRelateDebug, idxDebugRelate
dim relateRows, relateCols

set cSqlDebug = new cRunSql
cSqlDebug.Conn = connOracle
cSqlDebug.SqlStr = "select FIELD_NAME,DATA from app_renew_detail where app_renew_id = ? and upper(field_name) not like 'ARRRELATE%' and upper(field_name) not like 'ARR_FILES_%' and upper(field_name) not in ("
for i = lbound(arrSysVars) to ubound(arrSysVars)
	cSqlDebug.AddSqlStr("upper('" & arrSysVars(i) & "')")
	if i < ubound(arrSysVars) then
		cSqlDebug.AddSqlStr(",")
	end if
next
cSqlDebug.AddSqlStr(")")
cSqlDebug.AddParam(cStr(request.querystring("formID")))
set rsDebug = cSqlDebug.Execute
if rsDebug.eof then
else
	'session.abandon
	do while not rsDebug.eof
		'response.write "setting " & rsDebug("FIELD_NAME") & " to " & rsDebug("DATA") & "<BR>"
		session(rsDebug("FIELD_NAME")) = cStr(rsDebug("DATA")&"")
		rsDebug.MoveNext
	loop
end if
rsDebug.close
set rsDebug = nothing
set cSqlDebug = nothing

if session("MEMBERSHIP_ID")&"" = "" or session("MEMBERSHIP_ID")&"" = "0" then
	set cSqlDebug = new cRunSql
	cSqlDebug.Conn = connOracle
	cSqlDebug.SqlStr = "select membership_id from app_renew where app_renew_id = ?"
	cSqlDebug.AddParam(cStr(request.querystring("formID")))
	set rsDebug = cSqlDebug.Execute
	if rsDebug.eof then
	else
		if rsDebug("MEMBERSHIP_ID")&"" <> "0" then
			session("MEMBERSHIP_ID") = rsDebug("MEMBERSHIP_ID")
		end if
	end if
	rsDebug.close
	set rsDebug = nothing
	set cSqlDebug = nothing
end if

redisKey = "FORM_" & Application("ORG_ID") & "_" & session("APP_RENEW_ID")
set redisArr = Server.CreateObject("CacheManager")
redisArr.SetDefaultExpiration 365,0,0,0
redisArr.Init redisKey
arrKeys = split(redisArr.GetCacheList,",")
if isArray(arrKeys) then
	for x = lbound(arrKeys) to ubound(arrKeys)
		item = arrKeys(x)
		if isArray(redisArr(item)) Then
			if ucase(item) = "ARR_PAGES" or ucase(item) = "ARR_FIELDS" or ucase(item) = "ARR_FIELD_COLS" or ucase(item) = "ARR_COMBINE_FIELDS" or ucase(item) = "ARR_COMBINE_FIELDS_JS" or ucase(item) = "ARR_PAGE" or ucase(item) = "ARR_PAGE_COLS" or ucase(item) = "ARR_CONFIG" or ucase(item) = "ARR_CONFIG_COLS" or ucase(item) = "ARR_BATCH_PRODS" then
			else
				session(item) = redisArr(item)
			end if
		end if
	next
end if
set redisArr = nothing

end sub

sub writeNodes(xml)
%>
<textarea rows="10" columns="50">
<%
dim objRootElement
Set objRootElement = xml.documentElement
Response.Write(objRootElement.xml)
%>
</textarea>
<%
set objRootElement = nothing
end sub

sub showProgress
%><div class="formprogress"><%
%><div class="formprogress-heading"><%=formDesc%> Progress</div><%
%><table id="tblProgress"><%
for i = lbound(arrPages,2) to ubound(arrPages,2)
	if cint(arrPages(1,i)) > 0 and arrPages(3,i) <> "Y" then
	%><tr id="tbrProg<%=arrPages(0,i)%>"><%
		%><td><%
			%><input type="checkbox" disabled<%if isArray(session("ARR_PAGES_COMPLETE")) then%><%for j=lbound(session("ARR_PAGES_COMPLETE")) to ubound(session("ARR_PAGES_COMPLETE"))%><%if session("ARR_PAGES_COMPLETE")(j) = arrPages(0,i) then%> checked<%end if%><%next%><%END IF%>><%
		%></td><%
		%><td><%
			%><a href="<%=arrPages(0,i)%>"><%if arrPages(2,i) <> "" then%><%=arrPages(2,i)%><%else%><%=arrPages(0,i)%><%end if%></a><%
		%></td><%
	%></tr><%
	end if
next
	%><tr><%
		%><td colspan="2"><%
			%>Finish<%
		%></td><%
	%></tr><%
%></table><%
%></div><%
end sub

function amountDue
dim prodID, arrProdNo, idxProdNo, groupProdID, rs2
dim indProdID, prodDesc, prodAmt, rsP, rsP2
dim arrProdID, prodIDX, arrProdConf, prodField, numProd
dim arrBatchProds, idxBatchProds, idxBatchProd

if isArray(session("ARR_BATCH_PRODS")) then
	arrBatchProds = session("ARR_BATCH_PRODS")
else
	set cSqlForm = new cRunSql
	cSqlForm.conn = connOracle
	cSqlForm.SqlStr = "select PRODUCT_ID,PRODUCT_DESC,COST_AMT from " & prodTbl
	set rsP = cSqlForm.Execute
	set cSqlForm = nothing
	if rsP.eof Then
		arrBatchProds = ""
	else
		arrBatchProds = rsP.GetRows()
	end if
	rsP.close
	set rsP = nothing
	session("ARR_BATCH_PRODS") = arrBatchProds
end if

amountDue = 0

if loadDefaultsDelay = true then
	delayLog "onclick products to use",true
else
	delayLog "onclick products to use",false
	prodID = ""
	for idxField = lbound(arrFields,2) to ubound(arrFields,2)
		if arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "RADIO" and arrFields(fieldNameIndex("PRODUCT_NO"),idxField)&"" <> "" then
			if session(arrFields(fieldNameIndex("FIELD_NAME"),idxField))&"" = yearFunction(arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField)&"") then
				arrProdNo = split(arrFields(fieldNameIndex("PRODUCT_NO"),idxField),", ")
				for idxProdNo = lbound(arrProdNo) to ubound(arrProdNo)
					if prodID = "" then
					else
						prodID = prodID & ","
					end if
					indProdID = split(cRenew.ConfigOption("product")&"",", ")(arrProdNo(idxProdNo))

					if arrFields(fieldNameIndex("field_prod_qty"),idxField)&"" <> "" then
						numProd = session(cRenew.getFieldName(arrFields(fieldNameIndex("field_prod_qty"),idxField)))	
						if numProd&"" <> "" then
							if isNumeric(numProd) then
								numProd = cDbl(numProd)
								set cSqlForm = new cRunSql
								cSqlForm.conn = connOracle
								cSqlForm.SqlStr = "select PRODUCT_DESC_BRIEF, PRODUCT_DESC, COST_AMT, CATEGORY from " & prodTbl & " where product_id = ?"
								cSqlForm.AddParam(indProdID)
								set rsP = cSqlForm.Execute
								set cSqlForm = nothing
								if rsP.eof then
								else
									session(arrFields(fieldNameIndex("FIELD_NAME"),idxField) & "_AMOUNT") = numProd * cdbl(rsP("COST_AMT"))
									set cSqlForm = new cRunSql
									cSqlForm.conn = connOracle
									cSqlForm.SqlStr = "select PRODUCT_ID from " & prodTbl & " where PRODUCT_DESC_BRIEF = ?"
									cSqlForm.AddParam(rsP("PRODUCT_DESC_BRIEF") & "_" & numProd)
									set rsP2 = cSqlForm.Execute
									set cSqlForm = nothing
									if rsP2.eof then
										indProdID = getSequence("SEQ_PRODUCT_ID")
										set cSqlForm = new cRunSql
										cSqlForm.Conn = connOracle
										cSqlForm.SqlStr = "insert into batch_product(product_id,org_id,product_desc,product_desc_brief,cost_amt,inactive,category) values (?,?,?,?,?,?,?)"
										cSqlForm.AddParam(indProdID)
										cSqlForm.AddParam(Application("ORG_ID"))
										cSqlForm.AddParam(rsP("PRODUCT_DESC") & " - " & numProd)
										cSqlForm.AddParam(rsP("PRODUCT_DESC_BRIEF") & "_" & numProd)
										cSqlForm.AddParam(numProd * cdbl(rsP("COST_AMT")))
										cSqlForm.AddParam("N")
										cSqlForm.AddParam(rsP("CATEGORY"))
										cSqlForm.Execute
										set cSqlForm = nothing
									else
										indProdID = rsP2("PRODUCT_ID")
									end if
									rsP2.close
									set rsP2 = nothing
								end if
								rsP.close
								set rsP = nothing
							end if
						end if
					end if

					if isArray(arrBatchProds) then
						for idxBatchProd = lbound(arrBatchProds,2) to ubound(arrBatchProds,2)
							if indProdID&"" = arrBatchProds(0,idxBatchProd)&"" then
								prodDesc = arrBatchProds(1,idxBatchProd)
								prodAmt = arrBatchProds(2,idxBatchProd)
								idxBatchProd = ubound(arrBatchProds,2) + 1
							end if
						next
					end if
					msg = "adding product " & arrProdNo(idxProdNo)+1 & ". " & prodDesc & " " & FormatCurrency(prodAmt) & " because " & arrFields(fieldNameIndex("FIELD_NAME"),idxField) & " = " & yearFunction(arrFields(fieldNameIndex("DEFAULT_VALUE"),idxField))
					cRenew.AddMsg(msg)
					formLog(msg)
					prodID = prodID & indProdID
				next
			end if
		end if
	next
end if
'response.end
if prodID <> "" then
	session("PRODUCT_ID") = prodID
end if

if session("PRODUCT_ID")&"" = "" and cRenew.ConfigOption("product")&"" <> "" then
	session("PRODUCT_ID") = split(cRenew.ConfigOption("product")&"",", ")(0)
end if
if session("PRODUCT_ID")&"" <> "" and session("PRODUCT_ID")&"" <> "0" then

	arrProdID = split(session("PRODUCT_ID")&"",",")
	arrProdConf = split(cRenew.ConfigOption("product")&"",", ")

	for i = lbound(arrProdID) to ubound(arrProdID)
		prodIDX = -1
		for j = lbound(arrProdConf) to ubound(arrProdConf)
			if isNumeric(arrProdConf(j)) then
				if cint(arrProdID(i)) = cint(arrProdConf(j)) then
					prodIDX = j
				end if
			end if
		next
		if j >=0 then
			for idxField = lbound(arrFields,2) to ubound(arrFields,2)
				if arrFields(fieldNameIndex("FIELD_TYPE"),idxField) = "PRODUCT_FEE" and arrFields(fieldNameIndex("PRODUCT_NO"),idxField)&"" = prodIDX&"" then
					prodField = arrFields(fieldNameIndex("FIELD_NAME"),idxField)
					
					if isArray(arrBatchProds) then
						for idxBatchProd = lbound(arrBatchProds,2) to ubound(arrBatchProds,2)
							if arrProdConf(prodIDX)&"" = arrBatchProds(0,idxBatchProd)&"" then
								session(prodField) = cStr(arrBatchProds(2,idxBatchProd))
								idxBatchProd = ubound(arrBatchProds,2) + 1
							end if
						next
					end if
				end if
			next
		end if
	next

	amountDue = 0

	arrProductID = split(session("PRODUCT_ID")&"",",")
	for i = lbound(arrProductID) to ubound(arrProductID)
		if isArray(arrBatchProds) then
			for idxBatchProd = lbound(arrBatchProds,2) to ubound(arrBatchProds,2)
				if arrProductID(i)&"" = arrBatchProds(0,idxBatchProd)&"" then
					amountDue = amountDue + cdbl(arrBatchProds(2,idxBatchProd))
					idxBatchProd = ubound(arrBatchProds,2) + 1
				end if
			Next
		end if
	next

	if session("GROUP_PAY_ID")&"" <> "" then
		set cSqlForm = new cRunSql
		cSqlForm.conn = connOracle
		cSqlForm.SqlStr = "select PRODUCT_ID from pay_later where group_pay_id = ? and status = 'PENDING' " 
		cSqlForm.AddParam(session("GROUP_PAY_ID"))
		set rs = cSqlForm.Execute
		set cSqlForm = nothing
		do while not rs.eof
			set cSqlForm = new cRunSql
			cSqlForm.conn = connOracle
			cSqlForm.SqlStr = "select sum(cost_amt) cost_amt from " & prodTbl & " where product_id = ?"
			cSqlForm.AddParam(rs("PRODUCT_ID"))
			set rs2 = cSqlForm.Execute
			if rs2.eof then
			else
				amountDue = amountDue + cdbl(rs2("COST_AMT"))
			end if
			rs2.close
			set rs2 = nothing
			set cSqlForm = nothing
			rs.MoveNext
		loop
		rs.close
		set rs = nothing
	end if

end if

end function

sub loginCheck
	dim loginID

	if cRenew.PageOption("require_login") = "Y" then
		if cRenew.ConfigOption("login_type") = "USER" then
			loginID = session("LOGIN_ID")&""
		else
			loginID = session("MEMBERSHIP_ID")&""
		end if
		if (loginID = "" or loginID = "0") and request("doPDF") <> "1" then
			if cRenew.PageCD(cRenew.ConfigOption("login_page"))&"" <> "" then
				if isDev and request("doEmail") <> "1" and request("doPDF") <> "1" then
				%>
				<script>
				alert("redirecting to <%=aspxPage(cRenew.PageCD(cRenew.ConfigOption("login_page")))%> because Require Login is set to Y for this page");
				document.location = "<%=aspxPage(cRenew.PageCD(cRenew.ConfigOption("login_page")))%>";
				</script>
				<%
					closeAll()
					response.end
				else
					redirectForm(cRenew.PageCD(cRenew.ConfigOption("login_page")))
				end if
			elseif cRenew.ConfigOption("exit_path")&"" <> "" then
				if isDev and request("doEmail") <> "1" and request("doPDF") <> "1" then
				%>
				<script>
				alert("redirecting to <%=cRenew.ConfigOption("exit_path")%> because Require Login is set to Y for this page and Exit Path is set");
				document.location = "<%=cRenew.ConfigOption("exit_path")%>";
				</script>
				<%
					closeAll()
					response.end
				else
					redirectForm(cRenew.ConfigOption("exit_path"))
				end if
			else
				if isDev and request("doEmail") <> "1" and request("doPDF") <> "1" then
				%>
				<script>
				alert("redirecting to / because Require Login is set to Y and no Login Page has been specified");
				document.location = "/";
				</script>
				<%
					closeAll()
					response.end
				else
					redirectForm("/")
				end if
			end if
		end if
	end if
end sub
%>

<%
function yearFunction(str)

	if instr(str&"","[current year") then
		for i = 100 to 1 step -1
			str = replace(str&"","[current year - " & i & "]",year(sysDate)-i)
		next

		for i = 1 to 100
			str = replace(str&"","[current year + " & i & "]",year(sysDate)+i)
		next

		str = replace(str&"","[current year]",year(sysDate))
	end if

	if instr(str&"","[current date") then
		for i = 1 to 100
			str = replace(str&"","[current date + " & i  & " years]",dateAdd("yyyy",i,sysDate))
		next
		for i = 1 to 100
			str = replace(str&"","[current date - " & i  & " years]",dateAdd("yyyy",-i,sysDate))
		next
		for i = 1 to 365
			str = replace(str&"","[current date - " & i  & " days]",dateAdd("d",-i,sysDate))
		next
		for i = 1 to 365
			str = replace(str&"","[current date + " & i  & " days]",dateAdd("d",i,sysDate))
		next
		for i = 1 to 12
			str = replace(str&"","[current date - " & i  & " months]",dateAdd("m",-i,sysDate))
		next
		for i = 1 to 12
			str = replace(str&"","[current date + " & i  & " months]",dateAdd("m",i,sysDate))
		next
		for i = 1 to 52
			str = replace(str&"","[current date - " & i  & " weeks]",dateAdd("ww",-i,sysDate))
		next
		for i = 1 to 52
			str = replace(str&"","[current date + " & i  & " weeks]",dateAdd("ww",i,sysDate))
		next
	end if

	str = replace(str&"","[current date]",sysDate)

	if instr(str&"","[current month") then
		for i = 12 to 1 step -1
			str = replace(str&"","[current month - " & i & "]",monthName(month(dateAdd("m",-i,sysDate))))
		next

		for i = 1 to 12
			str = replace(str&"","[current month + " & i & "]",monthName(month(dateAdd("m",i,sysDate))))
		next

		str = replace(str&"","[current month]",monthName(month(sysDate)))
	end if

	if instr(str&"","[current day") then
		for i = 7 to 1 step -1
			str = replace(str&"","[current day - " & i & "]",WeekdayName(Weekday(dateAdd("d",-i,sysDate))))
		next

		for i = 1 to 7
			str = replace(str&"","[current day + " & i & "]",WeekdayName(Weekday(dateAdd("d",i,sysDate))))
		next

		str = replace(str&"","[current day]",WeekdayName(Weekday(sysDate)))
	end if

	yearFunction = str

end function

function mmddyyyy(dt)
	dim tmpMonth, tmpDay, tmpYear

	if len(dt) = 6 then
		dt = left(dt,2) & "/" & right(left(dt,4),2) & "/" & right(dt,2)
	end if

	tmpMonth = month(cdate(dt))
	tmpDay = day(cdate(dt))
	tmpYear = year(cdate(dt))
	if tmpMonth < 10 then
		tmpMonth = "0" & tmpMonth
	end if
	if tmpDay < 10 then
		tmpDay = "0" & tmpDay
	end if
	mmddyyyy = tmpMonth & "/" & tmpDay & "/" & tmpYear
end function

function CompareIPsForm(arSrc, arIP1, arIP2)
	dim t
    For t = 0 To 3
        If cint(arSrc(t)) >= cint(arIP1(t)) And cint(arSrc(t)) <= cint(arIP2(t)) Then
            If cint(arSrc(t)) <> cint(arIP1(t)) And cint(arSrc(t)) <> cint(arIP2(t)) Then
                CompareIPsForm = True
                Exit Function
            End If
        Else
            CompareIPsForm = False
            Exit Function
        End If
    Next
    CompareIPsForm = True

end function%>

<%function dynDate(sql)%>
<%
dim RegExpObj, Pattern, Matches, arrOp, unitType, numUnits, dateOp

arrOp = split("+,-",",")

if instr(sql,"*date+") > 0 or instr(sql,"*date-") then
    for i = lbound(arrOp) to ubound(arrOp)

	dateOp = arrOp(i)
	Set RegExpObj = New RegExp

	With RegExpObj

	Pattern = CHR(39)&"\*date\" & dateOp & "(.|\n)*?"&CHR(39)

	.Pattern = Pattern
	.IgnoreCase = True
	.Global = True
	End With
	Set Matches = RegExpObj.Execute(sql)

	For Each Item in Matches
	    arrDateStr = split(item, dateOp)
	    unitType = left(arrDateStr(1),1)
	    numUnits = replace(right(arrDateStr(1),len(arrDateStr(1))-1),"'","")
	    if unitType = "m" then
		    sql = replace(sql,item,"to_char(add_months(sysdate," & dateOp & " "&numUnits&"),'mm/dd/yyyy')")
	    end if
	    if unitType = "d" then
		    sql = replace(sql,item,"to_char(sysdate " & dateOp & " "&numUnits&",'mm/dd/yyyy')")
	    end if
	    if unitType = "y" then
		    sql = replace(sql,item,"to_char(add_months(sysdate," & dateOp & " " &(numUnits*12)&"),'mm/dd/yyyy')")
	    end if
	Next

	Set RegExpObj = nothing
    next
end if

if instr(sql,"*ms") > 0 or instr(sql,"*me") > 0 or instr(sql,"*ys") > 0 or instr(sql,"*ye") > 0 or instr(sql,"*ws") > 0 or instr(sql,"*we") > 0 then
	dim arrPN(5),startEnd,nextMonth, monthEnd,nextYear,yearEnd,weekStart,weekEnd

	arrPN(0) = "*ms"
	arrPN(1) = "*me"
	arrPN(2) = "*ys"
	arrPN(3) = "*ye"
	arrPN(4) = "*ws"
	arrPN(5) = "*we"

    for i = lbound(arrOp) to ubound(arrOp)

	    dateOp = arrOp(i)
	    for j = lbound(arrPN) to ubound(arrPN)
	        Set RegExpObj = New RegExp

	        With RegExpObj

	        Pattern = CHR(39)&"\" & arrPN(j) & "\" & dateOp & "(.|\n)*?"&CHR(39)

	        'response.write Pattern & "<BR>"

	        .Pattern = Pattern
	        .IgnoreCase = True
	        .Global = True
	        End With
	        Set Matches = RegExpObj.Execute(sql)

	        For Each Item in Matches
		        'response.write "item = " & item & "<BR>"
	            arrDateStr = split(item, dateOp)
	            unitType = right(left(arrDateStr(0),3),1)
		        startEnd = right(left(arrDateStr(0),4),1)
	            numUnits = replace(arrDateStr(1),"'","")
		        'response.write "unitType = " & unitType & "<BR>"
		        'response.write "numUnits = " & numUnits & "<BR>"
		        'response.write "startEnd = " & startEnd & "<BR>"
		        'response.write "dateOp = " & dateOp & "<BR>"
	            if unitType = "m" then
			        if startEnd = "s" then
				        searchDate = dateAdd("m",dateOp&numUnits,cdate(month(date)&"/01/"&year(date) ))
		    	        sql = replace(sql,item,"'" & searchDate & "'")
			        else
				        nextMonth = dateAdd("m",1,date)
				        monthEnd = dateAdd("d",-1,cdate(month(nextMonth) & "/01/" & year(nextMonth)))
				        searchDate = dateAdd("m",dateOp&numUnits,monthEnd)
		    	        sql = replace(sql,item,"'" & searchDate & "'")
			        end if
	            end if
	            if unitType = "y" then
			        if startEnd = "s" then
				        searchDate = dateAdd("yyyy",dateOp&numUnits,cdate("01/01/"&year(date) ))
		    	        sql = replace(sql,item,"'" & searchDate & "'")
			        else
				        nextYear = dateAdd("yyyy",1,date)
				        'response.write "nextYear " & nextYear & "<BR>"
				        yearEnd = dateAdd("d",-1,cdate("01/01/" & year(nextYear)))
				        'response.write "yearEnd " & yearEnd & "<BR>"
				        searchDate = dateAdd("yyyy",dateOp&numUnits,yearEnd)
		    	        sql = replace(sql,item,"'" & searchDate & "'")
			        end if
	            end if
	            if unitType = "w" then
                    weekDayNum = WeekDay(date)
                    weekStart = dateAdd("d",-(weekDayNum-1),date)
                    weekEnd = dateAdd("d",6,weekStart)
                    numUnits = numUnits * 7
			        if startEnd = "s" then
				        searchDate = dateAdd("d",dateOp&numUnits,weekStart)
		    	        sql = replace(sql,item,"'" & searchDate & "'")
			        else
				        searchDate = dateAdd("d",dateOp&numUnits,weekEnd)
		    	        sql = replace(sql,item,"'" & searchDate & "'")
			        end if
	            end if
	        Next

	        Set RegExpObj = nothing
	    next
    next
end if

'response.write "dyndate sql = " & sqlStr
'response.end
sql = replace(sql,"to_date('*date','mm/dd/yyyy')","to_date(to_char(sysdate,'mm/dd/yyyy'),'mm/dd/yyyy')")
sql = replace(sql,"to_date('*date', 'mm/dd/yyyy')","to_date(to_char(sysdate,'mm/dd/yyyy'),'mm/dd/yyyy')")

'sql = replace(sql,"&#44;",",")
dynDate = sql
%>
<%end function%>

<%
sub redirect(page)
	redirectForm(page)
end sub

sub redirectForm(page)

	if request.querystring("ajax") = "1" then
		page = replace(page,".asp",".aspx")
	end if

	on error resume next
	closeAll()

	if request.querystring("ajax") = "1" then
		response.clear
		%>
		<script>document.location='<%=page%>';</script>
		<%
	else
		response.redirect(page)
	end if

	response.end
end sub

sub closeAll

	on error resume next

	if isObject(requestForm) then
		set requestForm = nothing
	end if

	if isObject(requestQueryString) then
		set requestQueryString = nothing
	end if

	if isObject(objFieldsType) then
		set objFieldsType = nothing
	end if

	if isObject(objFieldsPage) then
		set objFieldsPage = nothing
	end if

	if isObject(objFields) then
		set objFields = nothing
	end if

	if isObject(cForm) then
		set cForm = nothing
	end if

	if isObject(cRenew) then
		set cRenew = nothing
	end if

	if isObject(conn) then
		if not (conn is nothing) then
			if conn.State then
				if conn.State = 1 then
					dbDisconnect conn
				end if
			end if
		end if
	end if

	if isObject(connOracle) then
		if not (connOracle is nothing) then
			if connOracle.State then
				if connOracle.State = 1 then
					dbDisconnect connOracle
				end if
			end if
		end if
	end if

	if isObject(connMysql) then
		if not (connMySql is nothing) then
			if connMysql.State then
				if connMysql.State = 1 then
					dbDisconnect connMysql
				end if
			end if
		end if
	end if

	if isObject(connb) then
		if not (connb is nothing) then
			if connb.State then
				if connb.State = 1 then
					dbDisconnect connb
				end if
			end if
		end if
	end if

	if redis then
		if isObject(session) then
			set session = nothing
		end if
	end if
	if request.form("auto_save") = "1" then
		set cLogIt = new cLog
		cLogIt.WriteLog "forms\form_" & logFormID & ".log","finish autoSave"
		set cLogIt = nothing
	else
		cLogIt.FlushLog "forms\form_" & logFormID & ".log"
		set cLogIt = nothing
	end if
end sub

function getSessionKeys

dim arrSesKeys

if redis then
	arrSesKeys = split(session.GetCacheList,",")
else
	for each item in session.contents
		if isArray(arrSesKeys) then
			idxKeys = ubound(arrSesKeys)+1
			redim preserve arrSesKeys(idxKeys)
		else
			idxKeys = 0
			redim arrSesKeys(idxKeys)
		end if
		arrSesKeys(idxKeys) = item
	next
end if

getSessionKeys = arrSesKeys

end function

function functionReplace(sql)

dim arrRT, idxRT, tmpSql, idxFunc

tmpSql = "select code_value from code_value where code_class = 'ORG_SUB_TY_CD' and inactive = 'N'"
set rs = connOracle.execute(tmpSql)
if rs.eof then
else
	sql = replace(sql,"SELECT","select")
    arrRT = rs.GetRows
    for idxRT = lbound(arrRT,2) to ubound(arrRT,2)
		if instr(sql,"Parent_Record_") then
			sql = replace(sql,"Parent_Record_" & arrRT(0,idxRT) & "(membership.membership_id)","(" & parentRecordFunction("Parent_Record_" & arrRT(0,idxRT),arrRT(0,idxRT)) & ")")
			for idxFunc = 2 to 10
				sql = replace(sql,"Parent_Record_" & idxFunc & "_" & arrRT(0,idxRT) & "(membership.membership_id)","(" & parentRecordFunction("Parent_Record_" & idxFunc & "_" & arrRT(0,idxRT),arrRT(0,idxRT)) & ")")
			next
		end if
		if instr(sql,"Grandparent_Record_") then
			sql = replace(sql,"Grandparent_Record_" & arrRT(0,idxRT) & "(membership.membership_id)","(" & grandParentRecordFunction("Grandparent_Record_" & arrRT(0,idxRT),arrRT(0,idxRT)) & ")")
			for idxFunc = 2 to 10
				sql = replace(sql,"Grandparent_Record_" & idxFunc & "_" & arrRT(0,idxRT) & "(membership.membership_id)","(" & grandParentRecordFunction("Grandparent_Record_" & idxFunc & "_" & arrRT(0,idxRT),arrRT(0,idxRT)) & ")")
			next
		end if
		if instr(sql,"Grandchild_Record_") then
			sql = replace(sql,"Grandchild_Record_" & arrRT(0,idxRT) & "(membership.membership_id)","(" & grandChildRecordFunction("Grandchild_Record_" & arrRT(0,idxRT),arrRT(0,idxRT)) & ")")
			for idxFunc = 2 to 10
				sql = replace(sql,"Grandchild_Record_" & idxFunc & "_" & arrRT(0,idxRT) & "(membership.membership_id)","(" & grandChildRecordFunction("Grandchild_Record_" & idxFunc & "_" & arrRT(0,idxRT),arrRT(0,idxRT)) & ")")
			next
		end if
		if instr(sql,"Child_Record_") then
			sql = replace(sql,"Child_Record_" & arrRT(0,idxRT) & "(membership.membership_id)","(" & childRecordFunction("Child_Record_" & arrRT(0,idxRT),arrRT(0,idxRT)) & ")")
			for idxFunc = 2 to 10
				sql = replace(sql,"Child_Record_" & idxFunc & "_" & arrRT(0,idxRT) & "(membership.membership_id)","(" & childRecordFunction("Child_Record_" & idxFunc & "_" & arrRT(0,idxRT),arrRT(0,idxRT)) & ")")
			next
		end if
    next
end if
rs.close
set rs = nothing

functionReplace = sql

end function

function childRecordFunction(functionName,recordType)

dim displayCols, arCols, idx, displayCol, newSql, codeClass, sqlStr, childType

sqlStr = "select other_record_type from org_column_detail where upper(column_desc) = upper('" & functionName & "')"
set rs = connOracle.execute(sqlStr)
if rs.eof then
else
    childType = rs("other_record_type")
	if childType&"" <> "" then
		sqlStr = "select org_display_columns from org_column_detail where upper(column_desc) =  upper('" & functionName & "')"
		set rs = connOracle.execute(sqlStr)
		if rs.eof then
		else
			displayCols = rs("org_display_columns")&""
			arrCols = split(displayCols,", ")
		end if
	end if
end if

newSql = ""

if isArray(arrCols) and childType&"" <> "" then
    newSql = "select "
    for idx = lbound(arrCols) to ubound(arrCols)
        displayCol = arrCols(idx)
        if displayCol <> "" then
            if idx > 0 then
                newSql = newSql & " || ' ' || "
            end if
            if left(displayCol,8) = "MEMBER::" then     
                sqlStr = "select code_class from membership_column_detail where org_sub_ty_cd = '" & childType & "' and column_desc = replace('" & displayCol & "','MEMBER::')"
                set rs = connOracle.execute(sqlStr)
                if rs.eof then
                else
                    codeClass = rs("code_class")&""
                    if codeClass <> "" then    
                        newSql = newSql & "(select code_value_desc from code_value where code_class = '" & codeClass & "' and code_value = '" & replace(displayCol,"MEMBER::","") & ")"  
                    else    
                        newSql = newSql & "c." & replace(displayCol,"MEMBER::","")      
                    end if
                end if
            else      
                sqlStr = "select code_class from org_column_detail where org_sub_ty_cd = '" & childType & "' and column_desc = replace('" & displayCol & "','DETAIL::')"
                set rs = connOracle.execute(sqlStr)
                if rs.eof then
                else
                    codeClass = rs("code_class")&""
                    if codeClass <> "" then    
                        newSql = newSql & "(select code_value_desc from code_value where code_class = '" &  codeClass & "' and code_value = nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = c.membership_id and column_desc = '" & replace(displayCol,"DETAIL::","") & "'),''))"    
                    else   
                        newSql = newSql & "nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = c.membership_id and column_desc = '" &  replace(displayCol,"DETAIL::","") & "'),'')"
                    end if
                end if    
            end if   
        end if
    next
    newSql = newSql & " child_rec from membership c, mor_finder f " &_
		"where c.membership_id = f.membership_id " &_
		"and f.org_relate_id = membership.membership_id  " &_
		"and c.org_sub_ty_cd = '" & childType & "' " &_
		"and c.m_stat_cd = 'ACTIVE' "      

	newSql = "select listagg(child_rec,', ') within group (order by child_rec) ""child_rec"" from (" & newSql & ")"
    childRecordFunction = newSql

end if

end function

function grandChildRecordFunction(functionName,recordType)
dim displayCols, grandChildType, arCols, idx, displayCol, newSql, codeClass, sqlStr, childType

sqlStr = "select other_record_type from org_column_detail where upper(column_desc) = upper('" & functionName & "')"
set rs = connOracle.execute(sqlStr)
if rs.eof then
else
    grandChildType = rs("other_record_type")
	sqlStr = "select minor_code from code_value where code_class = 'ORG_SUB_TY_CD' and code_value = '" & grandChildType & "'"
	set rs = connOracle.execute(sqlStr)
	if rs.eof then
	else
		childType = rs("minor_code")
	end if
	if grandChildType&"" <> "" and childType&"" <> "" then
		sqlStr = "select org_display_columns from org_column_detail where upper(column_desc) =  upper('" & functionName & "')"
		set rs = connOracle.execute(sqlStr)
		if rs.eof then
		else
			displayCols = rs("org_display_columns")&""
			arrCols = split(displayCols,", ")
		end if
	end if
end if

newSql = ""

if isArray(arrCols) and grandChildType&"" <> "" and childType&"" <> "" then
    newSql = "select "
    for idx = lbound(arrCols) to ubound(arrCols)
        displayCol = arrCols(idx)
        if displayCol <> "" then
            if idx > 0 then
                newSql = newSql & " || ' ' || "
            end if
            if left(displayCol,8) = "MEMBER::" then     
                sqlStr = "select code_class from membership_column_detail where org_sub_ty_cd = '" & grandChildType & "' and column_desc = replace('" & displayCol & "','MEMBER::')"
                set rs = connOracle.execute(sqlStr)
                if rs.eof then
                else
                    codeClass = rs("code_class")&""
                    if codeClass <> "" then    
                        newSql = newSql & "(select code_value_desc from code_value where code_class = '" & codeClass & "' and code_value = '" & replace(displayCol,"MEMBER::","") & ")"  
                    else    
                        newSql = newSql & "g." & replace(displayCol,"MEMBER::","")      
                    end if
                end if
            else      
                sqlStr = "select code_class from org_column_detail where org_sub_ty_cd = '" & grandChildType & "' and column_desc = replace('" & displayCol & "','DETAIL::')"
                set rs = connOracle.execute(sqlStr)
                if rs.eof then
                else
                    codeClass = rs("code_class")&""
                    if codeClass <> "" then    
                        newSql = newSql & "(select code_value_desc from code_value where code_class = '" &  codeClass & "' and code_value = nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = g.membership_id and column_desc = '" & replace(displayCol,"DETAIL::","") & "'),''))"    
                    else   
                        newSql = newSql & "nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = g.membership_id and column_desc = '" &  replace(displayCol,"DETAIL::","") & "'),'')"
                    end if
                end if    
            end if   
        end if
    next
    newSql = newSql & " grandchild_rec from membership g, mor_finder f, mor_finder f2 " &_
		"where f2.org_relate_id = f.membership_id " &_
		"and g.membership_id = f2.membership_id " &_
		"and f.org_relate_id = membership.membership_id  " &_
		"and g.org_sub_ty_cd = '" & grandChildType & "' " &_
		"and g.m_stat_cd = 'ACTIVE' "      

	newSql = "select listagg(grandchild_rec,', ') within group (order by grandchild_rec) ""grandchild_rec"" from (" & newSql & ")"
    grandChildRecordFunction = newSql
end if

end function

function grandParentRecordFunction(functionName,recordType)

dim displayCols, grandParentType, arCols, idx, displayCol, newSql, codeClass, sqlStr, parentType

sqlStr = "select minor_code from code_value where code_class = 'ORG_SUB_TY_CD' and code_value = (select minor_code from code_value where code_class = 'ORG_SUB_TY_CD' and code_value = '" & recordType & "')"
set rs = connOracle.execute(sqlStr)
if rs.eof then
else
    grandParentType = rs("minor_code")
	sqlStr = "select minor_code from code_value where code_class = 'ORG_SUB_TY_CD' and code_value = '" & recordType & "'"
	set rs = connOracle.execute(sqlStr)
	if rs.eof then
	else
		parentType = rs("minor_code")
	end if

	if grandParentType&"" <> "" and parentType&"" <> "" then
		sqlStr = "select org_display_columns from membership_column_detail where upper(column_desc) = 'ORG_RELATE_ID' and org_sub_ty_cd = '" & parentType & "'"
		set rs = connOracle.execute(sqlStr)
		if rs.eof then
		else
			displayCols = rs("org_display_columns")&""
			arrCols = split(displayCols,", ")
		end if
	end if
end if

newSql = ""

if isArray(arrCols) and grandParentType&"" <> "" and parentType&"" <> "" then
    newSql = "select "
    for idx = lbound(arrCols) to ubound(arrCols)
        displayCol = arrCols(idx)
        if displayCol <> "" then
            if idx > 0 then
                newSql = newSql & " || ' ' || "
            end if
            if left(displayCol,8) = "MEMBER::" then     
                sqlStr = "select code_class from membership_column_detail where org_sub_ty_cd = '" & grandParentType & "' and column_desc = replace('" & displayCol & "','MEMBER::')"
                set rs = connOracle.execute(sqlStr)
                if rs.eof then
                else
                    codeClass = rs("code_class")&""
                    if codeClass <> "" then    
                        newSql = newSql & "(select code_value_desc from code_value where code_class = '" & codeClass & "' and code_value = '" & replace(displayCol,"MEMBER::","") & ")"  
                    else    
                        newSql = newSql & replace(displayCol,"MEMBER::","")      
                    end if
                end if
            else      
                sqlStr = "select code_class from org_column_detail where org_sub_ty_cd = '" & grandParentType & "' and column_desc = replace('" & displayCol & "','DETAIL::')"
                set rs = connOracle.execute(sqlStr)
                if rs.eof then
                else
                    codeClass = rs("code_class")&""
                    if codeClass <> "" then    
                        newSql = newSql & "(select code_value_desc from code_value where code_class = '" &  codeClass & "' and code_value = nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = m.membership_id and column_desc = '" & replace(displayCol,"DETAIL::","") & "'),''))"    
                    else   
                        newSql = newSql & "nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = m.membership_id and column_desc = '" &  replace(displayCol,"DETAIL::","") & "'),'')"
                    end if
                end if    
            end if   
        end if
    next
    newSql = newSql & " grandparent_rec FROM MEMBERSHIP M,MOR_FINDER F, MOR_FINDER G WHERE M.M_STAT_CD = 'ACTIVE' AND M.MEMBERSHIP_ID = G.ORG_RELATE_ID AND G.MEMBERSHIP_ID = F.ORG_RELATE_ID AND M.ORG_SUB_TY_CD = '" & grandParentType & "' AND F.MEMBERSHIP_ID = MEMBERSHIP.MEMBERSHIP_ID"      
	newSql = "select listagg(grandparent_rec,', ') within group (order by grandparent_rec) ""grandparent_rec"" from (" & newSql & ")"
    grandParentRecordFunction = newSql
end if

end function

function parentRecordFunction(functionName,recordType)

dim displayCols, parentType, arCols, idx, displayCol, newSql, codeClass, sqlStr

sqlStr = "select org_display_columns from org_column_detail where upper(column_desc) = upper('" & functionName & "')"
set rs = connOracle.execute(sqlStr)
if rs.eof then
else
    displayCols = rs("org_display_columns")&""
    arrCols = split(displayCols,", ")
end if
 
sqlStr = "select minor_code from code_value where code_class = 'ORG_SUB_TY_CD' and code_value = '" & recordType & "'"
set rs = connOracle.execute(sqlStr)
if rs.eof then
else
    parentType = rs("minor_code")
end if

newSql = ""

if isArray(arrCols) and parentType&"" <> "" then
    newSql = "select "
    for idx = lbound(arrCols) to ubound(arrCols)
        displayCol = arrCols(idx)
        if displayCol <> "" then
            if idx > 0 then
                newSql = newSql & " || ' ' || "
            end if
            if left(displayCol,8) = "MEMBER::" then     
                sqlStr = "select code_class from membership_column_detail where org_sub_ty_cd = '" & parentType & "' and column_desc = replace('" & displayCol & "','MEMBER::')"
                set rs = connOracle.execute(sqlStr)
                if rs.eof then
                else
                    codeClass = rs("code_class")&""
                    if codeClass <> "" then    
                        newSql = newSql & "(select code_value_desc from code_value where code_class = '" & codeClass & "' and code_value = '" & replace(displayCol,"MEMBER::","") & ")"  
                    else    
                        newSql = newSql & replace(displayCol,"MEMBER::","")      
                    end if
                end if
            else      
                sqlStr = "select code_class from org_column_detail where org_sub_ty_cd = '" & parentType & "' and column_desc = replace('" & displayCol & "','DETAIL::')"
                set rs = connOracle.execute(sqlStr)
                if rs.eof then
                else
                    codeClass = rs("code_class")&""
                    if codeClass <> "" then    
                        newSql = newSql & "(select code_value_desc from code_value where code_class = '" &  codeClass & "' and code_value = nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = m.membership_id and column_desc = '" & replace(displayCol,"DETAIL::","") & "'),''))"    
                    else    
                        newSql = newSql & "nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = m.membership_id and column_desc = '" &  replace(displayCol,"DETAIL::","") & "'),'')"
                    end if
                end if    
            end if   
        end if
    next
    newSql = newSql & " parent_rec FROM MEMBERSHIP M,MOR_FINDER F WHERE M.M_STAT_CD = 'ACTIVE' AND M.MEMBERSHIP_ID = F.ORG_RELATE_ID AND M.ORG_SUB_TY_CD = '" &  parentType & "' AND F.MEMBERSHIP_ID = MEMBERSHIP.MEMBERSHIP_ID"      
	newSql = "select listagg(parent_rec,', ') within group (order by parent_rec) ""parent_rec"" from (" & newSql & ")"
    parentRecordFunction = newSql
end if

end function

function sortArray(arr)

dim sortMe, num, idx

set sortMe = Server.CreateObject("System.Collections.ArrayList")
for idx = lbound(arr) to ubound(arr)
    sortMe.Add arr(idx)
next
sortMe.Sort
idx = 0
for each num in sortMe
    arr(idx) = num
    idx = idx + 1
next

set sortMe = nothing

sortArray = arr
  
end function

function sortArrayOLD(arr)
    dim connSort, sortTbl, cSqlSort, idxArr, rsSort, strSort, arrSort, arrSql

	Set sortLib = CreateObject("Scriptlet.TypeLib")
	sortTbl = "TMP_SORT_" & replace(Mid(sortLib.Guid, 2, 36),"-","")
	set sortLib = nothing		

    Set connSort = Server.CreateObject("ADODB.Connection") 
    connSort.Open Application("ConnectionString") 

    set cSqlSort = new cRunSql
    cSqlSort.Conn = connSort
    cSqlSort.SqlStr = "create table " & sortTbl & " (key varchar2(1000))"
    cSqlSort.Execute
    set cSqlSort = nothing

    for idxArr = lbound(arr) to ubound(arr)
        set cSqlSort = new cRunSql
        cSqlSort.Conn = connSort
        cSqlSort.SqlStr = "insert into " & sortTbl & " (key) values (?)"
        cSqlSort.AddParam(arr(idxARr))
        cSqlSort.Execute
        set cSqlSort = nothing
    next

    redim arrSort(ubound(arr))

    set cSqlSort = new cRunSql
    cSqlSort.Conn = connSort
    cSqlSort.SqlStr = "select key from " & sortTbl & " order by upper(key)"
    set rsSort = cSqlSort.Execute
    set cSqlSort = nothing
    if rsSort.eof then
    else
        arrSql = rsSort.GetRows
        for idxArr = lbound(arrSql,2) to ubound(arrSql,2)
            arrSort(idxArr) = arrSql(0,idxArr)
        next

        sortArray = arrSort
    end if

    set cSqlSort = new cRunSql
    cSqlSort.Conn = connSort
    cSqlSort.SqlStr = "drop table " & sortTbl
    cSqlSort.Execute
    set cSqlSort = nothing
end function

function longDate(dirty)

dim arrTmp

if dirty <> "" then
    arrTmp = split(FormatDateTime(dirty,1),", ")
    longDate = arrTmp(1) & ", " & arrTmp(2)
end if

end function

function fieldNameIndex(fieldName)

dim idxFN, retVal

for idxFN = lbound(arrFieldNames) to ubound(arrFieldNames)
	if arrFieldNames(idxFN) = ucase(fieldName) then
		retVal = idxFN
		idxVN = ubound(arrFieldNames) + 1
	end if
next

fieldNameIndex = retVal

end function

sub paymentProcessorError
	cForm.AddErrorMsg("<BR><strong>This error is generated by the payment processor and is NOT a software issue on this end.  You must contact the client and/or payment processor.</strong>")
	if isDev and officeIP then
		%><script>alert("The payment processor has generated an error which is displayed on the screen.\n\nThis is NOT a software issue on this end.  Development staff cannot assist you.\n\nYou must contact the client and/or processor.  Only send a task on this end when specific instructions for software (not config) changes are available.")</script><%
	end if
end sub
%>