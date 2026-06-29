<!--#include virtual="/admin/admin/topshell.asp"-->
<!--#include virtual="/classes/cAwaitingUpdatesO.inc"-->
<!--#include virtual="/classes/cDocumentRepositoryO.inc"-->
<!--#include virtual="/classes/cUtilities.inc"-->
<!--#include virtual="/classes/cSendEmail.inc"-->
<!--#include virtual="/classes/cAppRenew.inc"-->
<!--#include virtual="/classes/cCodeValueO.inc"-->
<!--#include virtual="/classes/cFormSubmitLedger.inc"-->
<!--#include file="includes/tools.inc"-->
<!--#include virtual="/classes/cFileCloud.inc" -->
<!--#include virtual="/classes/aspJSON1.17.asp"-->
<!--#include virtual="/www/includes/dbtools.inc"-->
<!--#include virtual="/www/includes/hash.inc"-->
<%if Application("DEV_DOMAIN") = request.ServerVariables("SERVER_NAME") then%>
<script src="https://code.jquery.com/jquery-3.0.0.min.js"></script>
<script src="https://code.jquery.com/jquery-migrate-3.0.1.js"></script>
<%else%>
<script src="https://code.jquery.com/jquery-3.0.0.js"></script>
<script src="https://code.jquery.com/jquery-migrate-3.0.1.min.js"></script>
<%end if

Server.ScriptTimeout = 60*60

dim scriptName, rsAppStatus, connOracle, connb

devDomain = Application("DEV_DOMAIN")

scriptName = request.servervariables("SCRIPT_NAME")

dbConnect connOracle,Application("ConnectionString")

checkCodeClassO connOracle,"APP_RENEW_TYPE_CD","Application/Renewal Types"

dim cLogIt
set cLogIt = new cLog

if Application("CLIENT_ID")&"" = "" then
    Application("CLIENT_ID") = getCodeValueDesc("APPLICATION_DB","CLIENT_ID")
end if

dim connectionStringG,wsURL,xml,strJSON,oJSON,provider,dataSource,userID,password,connG,rsPDF,arrPDF,idxPDF
connectionStringG = Application("ConnectionStringG")&""
if connectionStringG = "" then
	wsURL = "https://ws.ebigpicture.com/GlobalDBConnectionInformation/"

	Set xml = Server.CreateObject("MSXML2.ServerXMLHTTP")
	xml.Open "GET", wsURL, False
	xml.setRequestHeader "Content-Type", "application/json"
	xml.Send()
	strJSON = xml.responseText
	set xml = nothing

	Set oJSON = New aspJSON
	oJSON.loadJSON(strJSON)
	provider = oJSON.data("Provider")
	dataSource = oJSON.data("Data Source")
	userID = oJSON.data("User ID")
	password = oJSON.data("Password")
	set oJSON = nothing

	connectionStringG = "Provider=" & provider & ";Data Source=" & dataSource & ";User ID=" & userID & ";Password=" & password & ";"
	Application("ConnectionStringG") = connectionStringG
end if

'ShowQueryStringVariables

if request("action") = "approve" then
	approve
elseif request("action") = "deny" then
	deny
elseif request("action") = "rebuildPDF" then
	rebuildPDF
elseif request("action") = "pend" then
	pend
elseif request("action") = "modify" then
	modify
elseif request("action") = "deleteView" then
	doDelete
    cLogIt.FlushLog "apprenewadmin\apprenewadmin_" & month(date)&day(date)&year(date) & ".log"
    set cLogIt = nothing
	redirectForm(scriptName & "?app_renew_cd=" & request("app_renew_cd") & "&app_renew_status=" & request("app_renew_status"))
elseif request("action") = "approveView" then
	doMassApprove
    cLogIt.FlushLog "apprenewadmin\apprenewadmin_" & month(date)&day(date)&year(date) & ".log"
    set cLogIt = nothing
	redirectForm(scriptName & "?app_renew_cd=" & request("app_renew_cd") & "&app_renew_status=" & request("app_renew_status"))
elseif request("action") = "rebuildMDRenewReport" then
    mdRenewalForm request("app_renew_id")
	modify
elseif request("action") = "lastSub" then
    lastSub
    cLogIt.FlushLog "apprenewadmin\apprenewadmin_" & month(date)&day(date)&year(date) & ".log"
    set cLogIt = nothing
    redirectForm(scriptName & "?action=modify&app_renew_id=" & request("app_renew_id"))
elseif request("action") = "subFormCopy" then
    subFormCopy
    cLogIt.FlushLog "apprenewadmin\apprenewadmin_" & month(date)&day(date)&year(date) & ".log"
    set cLogIt = nothing
    redirectForm(scriptName & "?action=modify&app_renew_id=" & request("app_renew_id"))
elseif request("action") = "doSearchLogs" then
	searchLogs
else
	mainMenu
end if

cLogIt.FlushLog "apprenewadmin\apprenewadmin_" & month(date)&day(date)&year(date) & ".log"
set cLogIt = nothing

dbDisconnect connOracle
%>

<!--#include virtual="/admin/admin/bottomshell.asp"-->

<%
sub searchPath(logPath,searchVal)
    dim folder, files, file, TS, fileContents, fileCounter, formID, rsAR
    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	set folder = objFSO.GetFolder(logPath)
	for each file in folder.Files
        if file.name <> "form_.log" and left(file.name,5) = "form_" and right(file.name,4) = ".log" then
    	    Set TS = objFSO.OpenTextFile(file.path, 1, false)
	        If Not TS.AtEndOfStream  Then
		        fileContents = lcase(TS.ReadAll)
                if instr(fileContents,searchVal) then
                    fileCounter = fileCounter + 1
                    formID = split(split(file.name,".")(0),"_")(1)
                    %>
                    <tr>
                        <td>
                            <%=fileCounter%>.
                        </td>
                        <td>
                            <a href="javascript:fetchLog('<%="forms\\" & file.name%>','<%=formID%>');"><%=file.name%></a>
                            <table align="center" id="tblLog<%=formID%>" style="display:none">
                                <tr><td>Fetching data...</td></tr>
                            </table>
                        </td>
                        <td>
                            <%
                                set cSql = new cRunSql
                                cSql.Conn = connOracle
                                cSql.SqlStr = "select code_value_desc from app_renew, code_value where code_value = app_renew_type and code_class = 'APP_RENEW_TYPE_CD' and app_renew_id = ?"
                                cSql.AddParam(formID)
                                set rsAR = cSql.Execute
                                set cSql = nothing
                                if rsAR.eof then
                                else
                                %>
                                <a href="<%=request.servervariables("SCRIPT_NAME")%>?action=modify&app_renew_id=<%=formID%>"><%=rsAR("code_value_desc")%></a>
                                <%
                                end if
                                rsAR.close
                                set rsAR = nothing
                            %>
                        </td>
                    </tr>
                    <%
                    response.flush
                end if
	        End If        
            set TS = nothing
        end if
    next
    set folder = nothing
    set objFSO = nothing
end sub

sub searchLogs

    dim searchVal, arrPath, objFSO, logPath

    searchVal = lcase(request.form("search"))

    if searchVal = "" then
        %><div align="center">No value entered to search</div><br /><%
        mainMenu
    else
        %><div align="center">Searching for <%=searchVal%></div><br /><%
        response.flush
        %>
        <script>
        function fetchLog(logPath,formID) {
            var el = document.getElementById("tblLog" + formID);
            if (el.style.display == "none") {
                if (el.innerHTML.indexOf("Fetching data...") >= 0) {
                    $("#tblLog" + formID).show();
                    $.post("/admin/admin/ajax.asp",
                    {
                        action: "showLogFile",
                        log_file: logPath,
                    },
                    function (data, status) {
                        if (status == "success") {
                            $("#tblLog" + formID).html(data);
                        } else {
                            alert("showLog failed");
                        }
                    });
                }
            }
        }
        </script>
        <table width="100%" border="0" cellpadding="2" cellspacing="0">
        <tr><td colspan="3"><input type="button" value="List" name="list" class="adminButton" onMouseOver="this.className='adminButtonHover'" onMouseOut="this.className='adminButton'" onClick="parent.location='<%=request.servervariables("SCRIPT_NAME")%>?app_renew_cd=<%=appRenewType%>&app_renew_status=<%=appRenewStatus%>'"></td></tr>
        <tr class="adminTitlebar">
	        <td>
	        &nbsp;
	        </td>
	        <td>
	            Log File
	        </td>
	        <td>
	            Submission
	        </td>
        </tr>
<%
    end if

    fileCounter = 0

    arrPath = split(request.servervariables("APPL_PHYSICAL_PATH")&"","\")

    dim xml
    
	Set xml = Server.CreateObject("Microsoft.XMLHTTP")
	xml.Open "GET", "https://ws.ebigpicture.com/logpath/", False
	xml.Send(data)
	logPath = xml.ResponseText & "\" & arrPath(3) & "\forms"
	set xml = nothing

    searchPath logPath,searchVal
    %></table><%
end sub

sub subFormCopy

set cSql = new cRunSql
cSql.conn = connOracle
cSql.SqlStr = "select field_name,data from app_renew_page_sub_detail where app_renew_page_submit_id = ? and data_type = 'RF'"
cSql.AddParam(request.QueryString("page_sub_id"))
set rs = cSql.Execute
set cSql = nothing
do while not rs.eof
    set cSql = new cRunSql
    cSql.conn = connOracle
    cSql.SqlStr = "delete from app_renew_detail where app_renew_id = ? and lower(field_name) = lower(?)"
    cSql.AddParam(request.QueryString("app_renew_id"))
    cSql.AddParam(rs("field_name"))
    cSql.Execute
    set cSql = nothing

    set cSql = new cRunSql
    cSql.conn = connOracle
    cSql.SqlStr = "insert into app_renew_detail (app_renew_id,field_name,data) values (?,?,?) " 
    cSql.AddParam(request.QueryString("app_renew_id"))
    cSql.AddParam(rs("field_name"))
    cSql.AddParam(rs("data"))
    cSql.Execute
    set cSql = nothing
    rs.MoveNext
loop
rs.close
set rs = nothing

end sub

sub lastSub

set cSql = new cRunSql
cSql.conn = connOracle
cSql.SqlStr = "delete from app_renew_detail where app_renew_id = ?"
cSql.AddParam(request.QueryString("app_renew_id"))
cSql.Execute
set cSql = nothing

set cSql = new cRunSql
cSql.conn = connOracle
cSql.SqlStr = "insert into app_renew_detail (app_renew_id,field_name,data) select s.app_renew_id,d.field_name,d.data from app_renew_page_sub_detail d, app_renew_page_submit s " &_
    "where s.app_renew_page_submit_id = d.app_renew_page_submit_id " &_
    "and d.data_type = 'S' " &_
    "and s.app_renew_page_submit_id = ? " &_
    "and d.field_name not like '%(%' and d.field_name not like '%)%' " 
cSql.AddParam(request.QueryString("page_sub_id"))
cSql.Execute
set cSql = nothing

end sub 

sub rebuildPDF()
dim formID, cRenew

response.write "<div align=""center"">Building PDF, please wait......</div>"
response.flush
formID = request("id")

set cRenew = new cAppRenew
cRenew.pdfRebuildByID(formID)
set cRenew = nothing
%>
<script>setTimeout(function(){document.location="<%=request.ServerVariables("SCRIPT_NAME")%>?action=modify&app_renew_id=<%=request.querystring("id")%>"; },30000);</script>
<%
response.end
end sub

sub mdRenewalForm(id)

dim theDoc, theURL, i, cDocRep, pdfFile, rs

set cSql = new cRunSql
cSql.conn = connOracle
cSql.SqlStr = "select membership_id from app_renew where app_renew_id = ?"
cSql.AddParam(id)
set rs = cSql.Execute
if rs.eof then
    membershipID = 0
else
    membershipID = rs("MEMBERSHIP_ID")
end if
rs.close
set rs = nothing
set cSql = nothing

Set theDoc = Server.CreateObject("ABCpdf11.Doc")
theDoc.HtmlOptions.UseScript = true
theDoc.MediaBox = "0 0 612 792"
theDoc.Color = "255 255 255"

theDoc.Rect = "0 0 612 792"
theDoc.HtmlOptions.BrowserWidth = 1200
theDoc.HtmlOptions.PageCacheEnabled = false	
theDoc.SetInfo 0, "HostWebBrowser", "0"
theDoc.HtmlOptions.FontProtection = false
theDoc.HtmlOptions.FontSubstitute = false
theDoc.HtmlOptions.FontEmbed = true
theDoc.HtmlOptions.ContentCount = 20
theDoc.HtmlOptions.RetryCount = 10
theDoc.HtmlOptions.Timeout = 15000

theDoc.HtmlOptions.Engine = 1

theDoc.Page = theDoc.AddPage()
if i = 0 then
	theDoc.SetInfo theDoc.Root, "/OpenAction", "[ " & theDoc.Page & " 0 R /XYZ null null 1 ]"
end if

theURL = "https://" & request.ServerVariables("SERVER_NAME") & "/db/mdRenewalPdf.asp?renewID=" & id

attCount = 0
        
on error resume next
        
theID = theDoc.AddImageUrl(theURL)
do while err.number <> 0 and attCount < 5
	response.Write "attCount = " & attCount & "<BR>"
	response.Write "err = " & err.Description & "<BR>"
	response.write theURL & "<BR>"
	response.flush
	err.Clear
	attCount = attCount + 1
    theID = theDoc.AddImageUrl(theURL)
loop
	    
if err.number <> 0 then
	set cMail = new cSendEmail
	cMail.Subject = "PDF Generation Error - WVBOM MD Renewal Form"
	cMail.TextBody = "URL = " & theURL & vbcrlf & vbcrlf & "ERROR = " & err.Description & vbcrlf & vbcrlf & "MEMBERSHIP_ID = " & membershipID
	cMail.Recipient = "errors@albertsonconsulting.com"
	cMail.From = "errors@albertsonconsulting.com"
	cMail.Send
	set cMail = nothing
end if
	    
on error goto 0
	    
theLoop = true
While theLoop
	theDoc.FrameRect()
	If Not theDoc.Chainable(theID) Then
		theLoop = false
	else
		theDoc.Page = theDoc.AddPage()
		theID = theDoc.AddImageToChain(theID)
	End If
Wend

For i = 1 To theDoc.PageCount
    theDoc.PageNumber = i
    theDoc.Flatten()
Next

pdfFile = "md_renewal_" & membershipID & datediff("s",cdate("01/01/1970"),now) & ".pdf"
theDoc.Save Application("ResourcePath")&"\"&pdfFile
theDoc.Clear()
set theDoc = nothing

set cDocRep = new cDocumentRepositoryO
cDocRep.InsertAppRenew membershipID, pdfFile, year(date) & " Application for Renewal of Medical Doctor License Report", id
set cDocRep = nothing

end sub
%>

<%
sub doDelete
dim i, userName, rsAR, cAwait, formStatus

userName = getSSI (Application("ADMIN_URL")&"/admin/cgi-bin/include.pl", "decrypt="&request.cookies("username"))

for i = 1 to request("deleteView").Count
    set cSql = new cRunSql
    cSql.conn = connOracle
    cSql.SqlStr = "select membership_id, app_renew_status from app_renew where app_renew_id = ?"
	cSql.AddParam(request("deleteView")(i))
    set rsAR = cSql.Execute
    set cSql = nothing

    formStatus = rsAR("app_renew_status")

    set cAwait = new cAwaitingUpdatesO
    if cAwait.hasUpdates(rsAR("MEMBERSHIP_ID")) then
        response.write "<div align=""center"">Form " & request("deleteView")(i) & " has Awaiting Updates and cannot be deleted</div>"
        rsAR.close
        set rsAR = nothing
        set cAwait = nothing
        response.end
    elseif (formStatus = "PENDING" and not officeIP) or formStatus = "COMPLETED" or formStatus = "APPROVED" then
        response.write "<div align=""center"">Form " & request("deleteView")(i) & " has a status of " & formStatus & " and cannot be deleted</div>"
        response.end
    else
        set cSql = new cRunSql
        cSql.conn = connOracle
	    cSql.sqlStr = "update app_renew set app_renew_status = 'DELETED' where app_renew_id = ?"
	    cSql.AddParam(request("deleteView")(i))
	    cSql.execute
        set cSql = nothing

        logit("form " & request("deleteView")(i) & " deleted by " & userName)
    end if
    rsAR.close
    set rsAR = nothing
    set cAwait = nothing
next
end sub
%>

<%
sub doMassApprove
dim i

for i = 1 to request("deleteView").Count
    set cSql = new cRunSql
    cSql.conn = connOracle
	cSql.sqlStr = "update app_renew set app_renew_status = 'APPROVED' where app_renew_id = ?"
	cSql.AddParam(request("deleteView")(i))
	cSql.execute
    set cSql = nothing
    approveUpdate(request("deleteView")(i))
next
end sub
%>

<%sub pend%>
<%
set cSql = new cRunSql
cSql.Conn = connOracle
cSql.SqlStr = "update app_renew set app_renew_status = 'PENDING' where app_renew_id = ?"
cSql.AddParam(request.querystring("id"))
cSql.Execute
set cSql = nothing

If Application("ORG_ID") = "1450" then
    set cSql = new cRunSql
    cSql.Conn = connOracle
    cSql.SqlStr = "select APP_RENEW_TYPE,MEMBERSHIP_ID from app_renew where app_renew_id = ?"
    cSql.AddParam(request.querystring("id"))
    set rs = cSql.Execute
    set cSql = nothing

    appRenewType = rs("APP_RENEW_TYPE")
    membershipID = rs("MEMBERSHIP_ID")

    set cSql = new cRunSql
    cSql.Conn = connOracle
    cSql.SqlStr = "update membership set m_stat_cd = 'ACTIVE' where to_char(membership_id) = (select org_relate_id from membership where membership_id = ?)"
    cSql.AddParam(membershipID)
    cSql.Execute
    set cSql = nothing
end if

redirectAppRenew(request.querystring("id"))
%>
<%end sub%>

<%sub deny%>
<%
dim oXMLHTTP

Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
oXMLHTTP.open "POST","https://" & request.ServerVariables("SERVER_NAME") & "/db/live/form_approve_deny_newsletter.asp",True
oXMLHTTP.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
oXMLHTTP.send("form_id=" & request.querystring("id") & "&action=D")
set oXMLHTTP = nothing

set cSql = new cRunSql
cSql.Conn = connOracle
cSql.SqlStr = "update app_renew set app_renew_status = 'DENIED' where app_renew_id = ?"
cSql.AddParam(request.querystring("id"))
cSql.Execute
set cSql = nothing
removeAwait

If Application("ORG_ID") = "3200" then
    set cMember = new cMemberValuesO
    cMember.UpdateValue getFormValByID(request.querystring("id"),"MEMBERSHIP_ID"),"M_STAT_CD","INACTIVE"
    set cMember = nothing
end if

If Application("ORG_ID") = "3100" then
    set cMember = new cMemberValuesO
    cMember.UpdateValue getFormValByID(request.querystring("id"),"MEMBERSHIP_ID"),"M_STAT_CD","INACTIVE"
    cMember.UpdateValue getFormValByID(request.querystring("id"),"CHILD_ID"),"M_STAT_CD","INACTIVE"
    set cMember = nothing
end if

If Application("ORG_ID") = "1800" then
    set cMember = new cMemberValuesO
    cMember.UpdateValue getFormValByID(request.querystring("id"),"MEMBERSHIP_ID"),"M_STAT_CD","INACTIVE"
    cMember.UpdateValue getFormValByID(request.querystring("id"),"CHILD_ID"),"M_STAT_CD","INACTIVE"
    set cMember = nothing
end if

If Application("ORG_ID") = "1900" then
    set cSql = new cRunSql
    cSql.Conn = connOracle
    cSql.SqlStr = "select * from app_renew where app_renew_id = ?"
    cSql.AddParam(request.querystring("id"))
    set rs = cSql.Execute
    set cSql = nothing

    appRenewType = rs("APP_RENEW_TYPE")
    membershipID = rs("MEMBERSHIP_ID")

    if appRenewType = "EMG_TRAIN" or appRenewType = "NA_RECRUIT" then
        set cMember = new cMemberValuesO
        cMember.UpdateValue getFormValByID(request.querystring("id"),"MEMBERSHIP_ID"),"M_STAT_CD","INACTIVE"
        cMember.UpdateValue getFormValByID(request.querystring("id"),"CHILD_ID"),"M_STAT_CD","INACTIVE"
        set cMember = nothing
    else
        set cMember = new cMemberValuesO
        cMember.LoadFromMembershipID membershipID
        cMember.UpdateValue membershipID,"M_STAT_CD","INACTIVE"
        cMember.UpdateValue cMember.Value("ORG_RELATE_ID"),"M_STAT_CD","INACTIVE"
        set cMember = nothing
    end if
end if

redirectAppRenew(request.querystring("id"))
%>
<%end sub%>

<%sub redirectAppRenew(app_renew_id)%>
<%
dim rsTmp,appRenewCode
set cSql = new cRunSql
cSql.Conn = connOracle
cSql.SqlStr = "select app_renew_type from app_renew where app_renew_id = ? "
cSql.AddParam(app_renew_id)
set rsTmp = cSql.Execute
appRenewCode = rsTmp("APP_RENEW_TYPE")
rsTmp.close
set rsTmp = nothing
set cSql = nothing

cLogIt.FlushLog "apprenewadmin\apprenewadmin_" & month(date)&day(date)&year(date) & ".log"
set cLogIt = nothing

redirectForm(scriptName & "?app_renew_cd=" & appRenewCode)
%>
<%end sub%>

<%sub approveExist(appID, existID, cMember, rsAppRenew, newID)%>
<%
dim cMainRecord

logit("approveExist " & existID & " memid = " & rsAppRenew("MEMBERSHIP_ID") & " newID = " & newID)
cSql.SqlStr = "select app_renew_fields.column_name, app_renew_detail.data, app_renew_fields.append_data,app_renew_fields.update_record " &_
    "from app_renew_detail, app_renew_fields, app_renew " &_
    "where upper(app_renew_fields.field_name) = upper(app_renew_detail.field_name) " &_
    "and app_renew.app_renew_id = app_renew_detail.app_renew_id " &_
    "and app_renew_fields.app_renew_type = app_renew.app_renew_type " &_
    "and app_renew_fields.column_name is not null " &_
    "and app_renew_fields.update_record in ('V','Y') " &_
    "and app_renew_detail.app_renew_id = ?"
cSql.AddParam(appID)

set cMainRecord = new cMemberValuesO
cMainRecord.AddLoadColMember("ORG_SUB_TY_CD")
cMainRecord.NoDetailData = true
cMainRecord.LoadFromMembershipID rsAppRenew("MEMBERSHIP_ID")
if cMainRecord.Value("ORG_SUB_TY_CD") = cMember.Value("ORG_SUB_TY_CD") then
    cSql.AddSqlStr(" and app_renew_fields.org_sub_ty_cd is null ")
else
    cSql.AddSqlStr(" and app_renew_fields.org_sub_ty_cd = ? ")
    cSql.AddParam(cMember.Value("ORG_SUB_TY_CD"))
end if
logIt(cSql.SqlStr)
logIt("ORG_SUB_TY_CD = " & cMember.Value("ORG_SUB_TY_CD"))
set cMainRecord = nothing
set rsAppRenewDetail = cSql.Execute

if rsAppRenewDetail.eof and rsAppRenewDetail.bof then
    logIt("eof = no columns to update")
else
    arrAppRenewDetail = rsAppRenewDetail.GetRows
end if

if isArray(arrAppRenewDetail) then
    for i = lbound(arrAppRenewDetail,2) to ubound(arrAppRenewDetail,2)
'		response.write arrAppRenewDetail(0,i) & " " & arrAppRenewDetail(1,i) & " " & cMember.Value(arrAppRenewDetail(0,i))&"<Br>"
        if arrAppRenewDetail(2,i) = "Y" then
            if arrAppRenewDetail(3,i) = "V" then
                if arrAppRenewDetail(1,i)&"" <> "" then
                    logIt("appending " & arrAppRenewDetail(0,i) & " with " & arrAppRenewDetail(1,i) & " for " & existID)
                    cMember.appendValue existID,arrAppRenewDetail(0,i),arrAppRenewDetail(1,i)
                else
                    logIt("NOT appending " & arrAppRenewDetail(0,i) & " because field is empty")
                end if
            else
                logIt("appending " & arrAppRenewDetail(0,i) & " with " & arrAppRenewDetail(1,i) & " for " & existID)
                cMember.appendValue existID,arrAppRenewDetail(0,i),arrAppRenewDetail(1,i)
            end if
        else
            if arrAppRenewDetail(3,i) = "V" then
                if arrAppRenewDetail(1,i)&"" <> "" then
                    logIt("updating " & arrAppRenewDetail(0,i) & " to " & arrAppRenewDetail(1,i) & " for " & existID)
                    cMember.updateValue existID,arrAppRenewDetail(0,i),arrAppRenewDetail(1,i)
                else
                    logIt("NOT updating " & arrAppRenewDetail(0,i) & " because field is empty")
                end if
            else
                logIt("updating " & arrAppRenewDetail(0,i) & " to " & arrAppRenewDetail(1,i) & " for " & existID)
                cMember.updateValue existID,arrAppRenewDetail(0,i),arrAppRenewDetail(1,i)
            end if
        end if
    Next
end if

rsAppRenewDetail.close
set rsAppRenewDetail = nothing

dim newRecordID
if clng("0" & newID) > 0 then
    newRecordID = newID
else
    newRecordID = rsAppRenew("MEMBERSHIP_ID")
end if

logIt("newRecordID = " & newRecordID)

set cDoc = new cDocumentRepositoryO
logIt("updating resource id from " & newRecordID & " to " & existID)
cDoc.UpdateMembershipID newRecordID,existID
set cDoc = nothing

if Application("ORG_ID") = "1900" then 
    set cNew = new cMemberValuesO
    cNew.LoadFromMembershipID newRecordID
    set cDoc = new cDocumentRepositoryO
    logIt("updating resource id from " & cNew.Value("ORG_RELATE_ID") & " to " & request.QueryString("existInd"))
    cDoc.UpdateMembershipID cNew.Value("ORG_RELATE_ID"),request.QueryString("existInd")
    set cDoc = nothing
    set cNew = nothing
end if

logIt("updating org_relate_id from " & newRecordID & " to " & existID)
cSql.SqlStr = "update membership set org_relate_id = ? where org_relate_id = ?"
cSql.AddParam(existID)
cSql.AddParam(newRecordID)
cSql.Execute

cSql.SqlStr = "update transaction_log set membership_id = ? where membership_id = ?"
cSql.AddParam(existID)
cSql.AddParam(newRecordID)
cSql.Execute

cSql.SqlStr = "update batch_member_finder set membership_id = ? where membership_id = ?"
cSql.AddParam(existID)
cSql.AddParam(newRecordID)
cSql.Execute
%>
<%end sub%>

<%sub approve%>
<%
dim rsRenewalFields, rsAppRenewDetail, rsAppRenew, cMember, appRenewType, arrAppRenewDetail, membershipID, contactID, cContact, newMembershipID, cPar, licenseID, cLicense, rs
dim oXMLHTTP, url

set cSql = new cRunSql
cSql.DisplayErrors = true
cSql.Conn = connOracle
cSql.SqlStr = "update app_renew set app_renew_status = 'APPROVED' where app_renew_id = ?"
cSql.AddParam(request.querystring("id"))
cSql.Execute

cSql.SqlStr = "select APP_RENEW_TYPE,MEMBERSHIP_ID from app_renew where app_renew_id = ?"
cSql.AddParam(request.querystring("id"))
set rsAppRenew = cSql.Execute
if rsAppRenew.eof then
else
    appRenewType = rsAppRenew("APP_RENEW_TYPE")
    membershipID = rsAppRenew("MEMBERSHIP_ID")
    newMembershipID = rsAppRenew("MEMBERSHIP_ID")
end if

set cMember = new cMemberValuesO

if Application("ORG_ID") = "1450" then
    cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
    if appRenewType = "COMPL_BIZ" then
        wvbopCompApprove cMember
    end if
end if

'ndhhs
if Application("ORG_ID") = "1900" then
    if appRenewType = "QRU_App" then
    	doSSI "https://" & request.ServerVariables("SERVER_NAME") & "/async_db/scripts/roster.asp","emsID=" & membershipID
    end if
end if

'nvot
if Application("ORG_ID") = "1500" then
    if request("exist") <> "" then
        cMember.LoadFromMembershipID membershipID
        approveExist request.querystring("id"), request("exist"), cMember, rsAppRenew, 0
        cMember.UpdateValue membershipID,"M_STAT_CD","INACTIVE"
        cSql.SqlStr = "select MEMBERSHIP_ID from membership where org_sub_ty_cd = 'Credential' and org_relate_id = ?"
        cSql.AddParam(membershipID)
        set rs = cSql.Execute
        if rs.eof then
        else
            cMember.UpdateValue rs("MEMBERSHIP_ID"),"ORG_RELATE_ID",request("exist")
        end if
    else
        cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
    end if
end if

'ndmirt
if Application("ORG_ID") = "1555" then
    if request("exist") <> "" then
        cMember.LoadFromMembershipID membershipID
        approveExist request.querystring("id"), request("exist"), cMember, rsAppRenew, 0
        cMember.UpdateValue membershipID,"M_STAT_CD","INACTIVE"
        cSql.SqlStr = "select MEMBERSHIP_ID from membership where org_sub_ty_cd = 'CREDENTIAL' and org_relate_id = ?"
        cSql.AddParam(membershipID)
        set rs = cSql.Execute
        if rs.eof then
        else
            cMember.UpdateValue rs("MEMBERSHIP_ID"),"ORG_RELATE_ID",request("exist")
        end if
        set cSql = nothing
        set cMember = nothing
        redirectAppRenew(request.querystring("id"))
    else
        cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
    end if
end if

if Application("ORG_ID") = "1400" then
    cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
end if

if Application("ORG_ID") = "1300" then
    if appRenewType = "admissionevaloffical" or appRenewType = "TRNFREXAMAPP" or appRenewType = "EXAMAPP" or appRenewType = "RECIPCERTWV" then
        if request("exist") <> "" then
            cMember.LoadFromMembershipID membershipID
            approveExist request.querystring("id"), request("exist"), cMember, rsAppRenew, 0
            redirectAppRenew(request.querystring("id"))
        else
            cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
        end if
    else
        cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
        if appRenewType = "INTENTEXAM" then
            if month(date) <= 3 then
                cMember.UpdateValue membershipID,"Testing_Window_MEMBER","Jan_Feb"
            elseif month(date) <=6 then
                cMember.UpdateValue membershipID,"Testing_Window_MEMBER","Apr_May"
            elseif month(date) <=9 then
                cMember.UpdateValue membershipID,"Testing_Window_MEMBER","Jul_Aug"
            else
                cMember.UpdateValue membershipID,"Testing_Window_MEMBER","Oct_Nov"
            end if
        end if
    end if
end if

if Application("ORG_ID") = "750" then
    cSql.SqlStr = "select MEMBERSHIP_ID from membership where org_sub_ty_cd = 'LIC' and m_stat_cd = 'ACTIVE' and org_relate_id = ?"
    cSql.AddParam(membershipID)
    set rs = cSql.Execute
    if rs.eof then
    else
        licenseID = rs("MEMBERSHIP_ID")
    end if
end if

if request("exist") <> "" then
    membershipID = request("exist") 
end if

cMember.LoadFromMembershipID membershipID

if Application("ORG_ID") = "15" then
    if appRenewType = "Program Application" then
        cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
    end if

    if appRenewType = "Supervision Plan" then
        if request("exist") <> "" then
            approveExist request.querystring("id"), request("exist"), cMember, rsAppRenew, request("new_id")
            redirectAppRenew(request.querystring("id"))
       else
            cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
       end if
    end if

    if appRenewType = "Provider Application" then
        if request("exist") <> "" then
            approveExist request.querystring("id"), request("exist"), cMember, rsAppRenew, 0
            if appRenewType = "Provider Application" then
                curExpDt = cMember.Value("LIC_EXPIRE_DT")
		        if curExpDt & "" <> "" then
			        if isDate(curExpDt) then
				        newExpDt = dateAdd("yyyy",1,curExpDt)
				        cMember.updateValue membershipID, "LIC_EXPIRE_DT", newExpDt
			        end if
		        end if
		        'cMember.updateValue membershipID,"Date_Received_PROVIDER",date
            end if
        else
            cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
            cMember.updateValue membershipID,"Date_Received_PROVIDER",date
            cMember.updateValue membershipID,"Date_Issued_PROVIDER",date
            cMember.updateValue membershipID, "LIC_EXPIRE_DT", month(now()) & "/" & day(now()) & "/" & (year(now())+1)
        end if
    end if
    if appRenewType = "License Application" then
        if request("exist") <> "" then
            approveExist request.querystring("id"), request("exist"), cMember, rsAppRenew, 0
            cMember.UpdateValue newMembershipID,"M_STAT_CD","INACTIVE"
            redirectAppRenew(request.querystring("id"))
        else
            cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
        end if
    end if
end if

'ndsbpe
if Application("ORG_ID") = "4800" then
    if appRenewType = "APPLY" and getFormVal("supform") = "N" and getFormVal("supCom") = "Y" then
        dbConnect connb,Application("ConnectionStringB")
        set cSqlB = new cRunSql
        cSqlB.Conn = connB
        cSqlB.SqlStr = "select SUBJECT,SENDER_EMAIL,MESSAGE from newsletter where upper(subject) = upper('Notification of Supervision Relationship')"
        set rs = cSqlB.Execute
        set cSqlB = nothing
        if rs.eof then
        else
            set cEmail = new cSendEmail
            cEmail.Subject = rs("SUBJECT")
            cEmail.From = rs("SENDER_EMAIL")
            if request.servervariables("SERVER_NAME") = devDomain then
                cEmail.Recipient = "zoe@ebigpicture.com"
            else
                cEmail.Recipient = getFormVal("supEmail")
            end if
            msg = rs("MESSAGE")
            msg = replace(msg,"[FIRST_NAME]",getFormVal("FIRST_NAME"))
            msg = replace(msg,"[LAST_NAME]",getFormVal("LAST_NAME"))
            msg = replace(msg,"[APP_NO]",getFormVal("APP_NO"))
            cEmail.HTMLBody = msg
            cEmail.Send
            set cEmail = nothing
        end if
        rs.close
        set rs = nothing
        dbDisconnect connB
    end if
end if

if Application("ORG_ID") = "750" then
    if (appRenewType = "CAAN" or appRenewType = "CAPN" or appRenewType = "UDAAPP" or appRenewType = "UMAAPP" or appRenewType = "UDARAPP" or appRenewType = "UMARAPP") then
        if request("exist") <> "" then
            approveExist request.querystring("id"), request("exist"), cMember, rsAppRenew, 0
            cMember.LoadFromMembershipID request.querystring("id")
            if request("existLic")&"" = "" then
                cSql.SqlStr = "update membership set m_stat_cd = 'ACTIVE' where m_stat_cd = 'PENDING' and org_relate_id = ?"
                cSql.AddParam(request("exist"))
                cSql.Execute
            end if
        else
            cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
            cMember.LoadFromMembershipID membershipID
        end if

        set cPar = new cMemberValuesO
        cPar.LoadFromMembershipID cMember.Value("ORG_RELATE_ID")

        if appRenewType <> "CAAN" and appRenewType <> "CAPN" and appRenewType <> "UDAAPP" and appRenewType <> "UDARAPP" and appRenewType <> "UMAAPP" and appRenewType <> "UMARAPP" then
            dbConnect connb,Application("ConnectionStringB")
            set cSql = new cRunSql
            cSql.Conn = connB
            if left(appRenewType,3) = "UDA" then
                cSql.SqlStr = "select SUBJECT,SENDER_EMAIL,MESSAGE from newsletter where upper(subject) = upper('UDA Application Approved')"
            else
                cSql.SqlStr = "select SUBJECT,SENDER_EMAIL,MESSAGE from newsletter where upper(subject) = upper('UMA Application Approved')"
            end if
            set rs = cSql.Execute
            set cSql = nothing
            if rs.eof then
            else
                set cEmail = new cSendEmail
                cEmail.Subject = replace(replace(rs("SUBJECT"),"UDA ",""),"UMA ","")
                cEmail.From = rs("SENDER_EMAIL")
                if request.servervariables("SERVER_NAME") = devDomain then
                    cEmail.Recipient = "bdeaver@albertsonconsulting.com"
                else
                    cEmail.Recipient = cMember.Value("EMAIL")
                end if
                msg = rs("MESSAGE")
                msg = replace(msg,"<FIRST_NAME>",rsEmpty("FIRST_NAME"))
                msg = replace(msg,"<LAST_NAME>",rsEmpty("LAST_NAME"))
                cEmail.HTMLBody = msg
                cEmail.Send
                set cEmail = nothing
            end if
            rs.close
            set rs = nothing
            dbDisconnect connB
        end if

        set cPar = nothing

        if request("exist") <> "" then
            redirectAppRenew(request.querystring("id"))
        end if
    end if
end if

if Application("ORG_ID") = "3900" then
    if right(appRenewType,4) = "_APP" then
        if request("exist") <> "" then
            set cSql2 = new cRunSql
            cSql2.Conn = connOracle
            cSql2.SqlStr = "select membership_id from membership where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'EQU' and org_relate_id = ?"
            cSql2.AddParam(membershipID)
            set rs = cSql2.Execute
            do while not rs.eof
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"M_STAT_CD","INACTIVE"
                rs.MoveNext
            loop
            rs.close
            set rs = nothing
            set cSql2 = nothing

            approveExist request.querystring("id"), request("exist"), cMember, rsAppRenew, 0
            logIt("deleting " & newMembershipID)
            cMember.UpdateValue newMembershipID,"M_STAT_CD","INACTIVE"
        else
            approveUpdate(request.querystring("id"))
            cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
        end if

        set cSql = new cRunSql
        cSql.Conn = connOracle
        cSql.SqlStr = "select membership_id from membership where m_stat_cd = 'PENDING' and org_sub_ty_cd = 'EQU' and org_relate_id = ?"
        cSql.AddParam(newMembershipID)
        set rs = cSql.Execute
        do while not rs.eof
            cMember.UpdateValue rs("MEMBERSHIP_ID"),"M_STAT_CD","ACTIVE"
            if request("exist") <> "" then
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"ORG_RELATE_ID",membershipID
            end if
            rs.MoveNext
        loop
        rs.close
        set rs = nothing
        set cSql = nothing
        redirectAppRenew(request.querystring("id"))
    end if
end if

if Application("ORG_ID") = "900" then
    if appRenewType = "CCCA License Application" or appRenewType = "License Application" or appRenewType = "Provisional Permit Application" then
        if request("exist") <> "" then
            approveExist request.querystring("id"), request("exist"), cMember, rsAppRenew, 0
            redirectAppRenew(request.querystring("id"))
        else
            cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
        end if
    end if
end if

if Application("ORG_ID") = "800" then
    if appRenewType = "DVM_Apply" or appRenewType = "CAET_Apply" or appRenewType = "Euth Fac Reg" or appRenewType = "Vet Facility Reg" or appRenewType = "RVT_Apply" then
        if request("exist") <> "" then
            approveExist request.querystring("id"), request("exist"), cMember, rsAppRenew, 0
        else
            cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
        end if
    end if
end if

if Application("ORG_ID") = "2900" then
    if request("exist") <> "" then
        approveExist request.querystring("id"), request("exist"), cMember, rsAppRenew, 0
        logIt("deleting " & newMembershipID)
        cMember.UpdateValue newMembershipID,"M_STAT_CD","INACTIVE"
        redirectAppRenew(request.querystring("id"))
    else
        cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
    end if
end if

if Application("ORG_ID") = "2400" then
    if request("exist") <> "" then
        approveExist request.querystring("id"), request("exist"), cMember, rsAppRenew, 0
        logIt("deleting " & newMembershipID)
        cMember.UpdateValue newMembershipID,"M_STAT_CD","INACTIVE"
        redirectAppRenew(request.querystring("id"))
    else
        if appRenewType = "BRANCH_APP" then
			dim newBranchID, curBranchID
			set cParent = new cNewMemberO
			cParent.SetValue "ORG_SUB_TY_CD","BRANCH"
			cParent.SetValue "M_STAT_CD","ACTIVE"
			newBranchID = cParent.MembershipID
			cParent.InsertRecord
			set cParent = nothing

            cSql.SqlStr = "delete from app_renew_detail where app_renew_id = ? and upper(field_name) = upper(?)"
            cSql.AddParam(request.querystring("id"))
            cSql.AddParam("PARENT_ID")
            cSql.Execute

            cSql.SqlStr = "insert into app_renew_detail (app_renew_id,field_name,data) values (?,?,?)"
            cSql.AddParam(request.querystring("id"))
            cSql.AddParam("PARENT_ID")
            cSql.AddParam(newBranchID)
            cSql.Execute

			set cMember = new cMemberValuesO
			cMember.LoadFromMembershipID membershipID
			curBranchID = cMember.Value("ORG_RELATE_ID")
			if curBranchID&"" <> "" then
				newBranchID = curBranchID & ", " & newBranchID
			end if
			cMember.UpdateValue membershipID, "ORG_RELATE_ID", newBranchID
			set cMember = nothing
        else
            cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
        end if
    end if
end if

if Application("ORG_ID") = "4100" then
    if appRenewType = "APP_VET" or appRenewType = "APP_TECH" then
        if request("existInd")&"" <> "" then
            cMember.LoadFromMembershipID request("existInd")
            approveExist request.querystring("id"), request("existInd"), cMember, rsAppRenew, newMemberID
        else
            cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
        end if
        set cMember = nothing
        approveUpdate(request.querystring("id"))
        if request("existInd")&"" <> "" then
            logIt("deleting " & membershipID)
            set cMember = new cMemberValuesO
            cMember.UpdateValue membershipID,"M_STAT_CD","INACTIVE"
            set cMember = nothing
        end if
        redirectAppRenew(request.querystring("id"))
    end if
end if

if Application("ORG_ID") = "3200" then
    if appRenewType = "APP" or appRenewType = "APP_RECIP" or appRenewType = "INTERN_APP" then
        if request("existInd")&"" <> "" then
            cMember.LoadFromMembershipID request("existInd")
            approveExist request.querystring("id"), request("existInd"), cMember, rsAppRenew, newMemberID
            logIt("deleting " & membershipID)
            cMember.UpdateValue membershipID,"M_STAT_CD","INACTIVE"
        else
            cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
        end if
        set cMember = nothing
        redirectAppRenew(request.querystring("id"))
    end if
end if

if Application("ORG_ID") = "3100" then
    if appRenewType = "APP" then
        if request("existInd")&"" <> "" then
            cMember.LoadFromMembershipID request("existInd")
            approveExist request.querystring("id"), request("existInd"), cMember, rsAppRenew, newMemberID
            logIt("deleting " & membershipID)
            cMember.UpdateValue membershipID,"M_STAT_CD","INACTIVE"
            if request("existLic")&"" <> "" then
                newLicID = membershipID
                set cLic = new cMemberValuesO
                cLic.LoadFromMembershipID request("existLic")
                approveExist request.querystring("id"), request("existLic"), cLic, rsAppRenew, getFormVal("CHILD_ID")
                set cLic = nothing
                cMember.UpdateValue getFormVal("CHILD_ID"),"M_STAT_CD","INACTIVE"
            else
                cMember.UpdateValue getFormVal("CHILD_ID"),"M_STAT_CD","ACTIVE"
            end if
        else
            cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
            cMember.UpdateValue getFormVal("CHILD_ID"),"M_STAT_CD","ACTIVE"
        end if
        set cMember = nothing
        redirectAppRenew(request.querystring("id"))
    end if
end if

if Application("ORG_ID") = "1800" then
    if appRenewType = "APP_LMSW" then
        if request("existInd")&"" <> "" then
            cMember.LoadFromMembershipID request("existInd")
            approveExist request.querystring("id"), request("existInd"), cMember, rsAppRenew, newMemberID
            logIt("deleting " & membershipID)
            cMember.UpdateValue membershipID,"M_STAT_CD","INACTIVE"
            cMember.UpdateValue getFormVal("CHILD_ID"),"ORG_RELATE_ID",request("existInd")
        else
            cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
        end if
        cMember.UpdateValue getFormVal("CHILD_ID"),"M_STAT_CD","ACTIVE"
        set cMember = nothing
        redirectAppRenew(request.querystring("id"))
    end if

    if appRenewType = "APP_LSW" then
        if request("existInd")&"" <> "" then
            cMember.LoadFromMembershipID request("existInd")
            approveExist request.querystring("id"), request("existInd"), cMember, rsAppRenew, newMemberID
            logIt("deleting " & membershipID)
            cMember.UpdateValue membershipID,"M_STAT_CD","INACTIVE"
            cMember.UpdateValue getFormVal("CHILD_ID"),"ORG_RELATE_ID",request("existInd")
        else
            cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
        end if
        cMember.UpdateValue getFormVal("CHILD_ID"),"M_STAT_CD","ACTIVE"
        set cMember = nothing
        redirectAppRenew(request.querystring("id"))
    end if

    if appRenewType = "APP" then
        if request("existInd")&"" <> "" then
            cMember.LoadFromMembershipID request("existInd")
            approveExist request.querystring("id"), request("existInd"), cMember, rsAppRenew, newMemberID
            logIt("deleting " & membershipID)
            cMember.UpdateValue membershipID,"M_STAT_CD","INACTIVE"
            if request("existLic")&"" <> "" then
                cMember.UpdateValue getFormVal("CHILD_ID"),"M_STAT_CD","INACTIVE"
            else
                cMember.UpdateValue getFormVal("CHILD_ID"),"M_STAT_CD","ACTIVE"
            end if
        else
            cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
            cMember.UpdateValue getFormVal("CHILD_ID"),"M_STAT_CD","ACTIVE"
        end if
        set cMember = nothing
        redirectAppRenew(request.querystring("id"))
    end if
end if

if Application("ORG_ID") = "1900" then
    if appRenewType = "MA_APP" or appRenewType = "MA2_APP" then
        logIt(appRenewType & " " & request.querystring("id"))
        cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
        if appRenewType = "MA_APP" then
            registry = getFormVal("registry")
        end if
        if appRenewType = "MA2_APP" then
            registry = getFormVal("registryna")
        end if
        parentID = getFormVal("PARENT_ID")
        logIt("parentID = " & parentID)

        cSql.SqlStr = "select MEMBERSHIP_ID from membership where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'LIC' and org_relate_id = ? " &_
            "and f_getmbrdetails(membership_id,'License_Type_LIC') = ?"
        cSql.AddParam(parentID)
        cSql.AddParam(registry)
        set rs = cSql.Execute
        if rs.eof then
            logIt("no existing record found for License_Type_LIC matching registry value " & registry)
        else
            logIt("existing record " & rs("MEMBERSHIP_ID") & " found for License_Type_LIC matching registry " & registry)
            cMember.UpdateValue rs("MEMBERSHIP_ID"),"LIC_EXPIRE_DT",getFormVal("MA_EXPIRE")
            cMember.UpdateValue rs("MEMBERSHIP_ID"),"Re_certification_Date_LIC",date
            cMember.UpdateValue rs("MEMBERSHIP_ID"),"Status_Change_Date_LIC",date
            cMember.UpdateValue rs("MEMBERSHIP_ID"),"Status_LIC",getFormVal("LIC_STATUS")
        end if    

        if getFormVal("cnarenew") = "1" then
            logit("cnarenew - 1")

            cSql.SqlStr = "select MEMBERSHIP_ID from membership where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'LIC' and org_relate_id = ? " &_
                "and f_getmbrdetails(membership_id,'License_Type_LIC') = '1' and f_getmbrdetails(membership_id,'Status_LIC') = '1' "
            cSql.AddParam(parentID)
            set rs = cSql.Execute
            if rs.eof then
                logIt("no existing record found for License_Type_LIC = 1 and Status_LIC = 1")
            else
                logIt("existing record " & rs("MEMBERSHIP_ID") & " found for License_Type_LIC = 1 and Status_LIC = 1")
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"LIC_EXPIRE_DT",dateAdd("yyyy",2,cdate(getFormVal("employmentenddate")))
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"Re_certification_Date_LIC",date
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"Status_Change_Date_LIC",date
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"LDW_LIC",getFormVal("employmentenddate")
            end if
            rs.close
            set rs = nothing
        end if

        if getFormVal("narenew") = "1" then
            logit("narenew - 1")
            cSql.SqlStr = "select MEMBERSHIP_ID from membership where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'LIC' and org_relate_id = ? " &_
                "and f_getmbrdetails(membership_id,'License_Type_LIC') = '3' and f_getmbrdetails(membership_id,'Status_LIC') = '1' "
            cSql.AddParam(parentID)
            set rs = cSql.Execute
            if rs.eof then
                logIt("no existing record found for License_Type_LIC = 3 and Status_LIC = 1")
            else
                logIt("existing record " & rs("MEMBERSHIP_ID") & " found for License_Type_LIC = 3 and Status_LIC = 1")
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"LIC_EXPIRE_DT","09/30/" & year(date)+2
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"Re_certification_Date_LIC",date
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"Status_Change_Date_LIC",date
            end if
            rs.close
            set rs = nothing
        end if
        if request("existLic")&"" <> "" then
            newLicID = membershipID
            set cLic = new cMemberValuesO
            cLic.LoadFromMembershipID request("existLic")
            approveExist request.querystring("id"), request("existLic"), cLic, rsAppRenew, newLicID
            set cLic = nothing
            logIt("deleting " & newLicID)
            cMember.UpdateValue newLicID,"M_STAT_CD","INACTIVE"
        else
            logIt("activating " & newLicID)
            cMember.UpdateValue newLicID,"M_STAT_CD","ACTIVE"
            logIt("setting org_relate_id to " & request("existInd"))
            cMember.UpdateValue newLicID,"ORG_RELATE_ID",request("existInd")
        end if
        set cMember = nothing
        redirectAppRenew(request.querystring("id"))
    end if

    if appRenewType = "HHA2_APP" then
        cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
        dim pay,rhistory1,rhistory2,rhistory3,rhistory4,rhistory5,rhistory6,parentID
        pay = getFormVal("pay")
        rhistory1 = getFormVal("rhistory1")
        rhistory2 = getFormVal("rhistory2")
        rhistory3 = getFormVal("rhistory3")
        rhistory4 = getFormVal("parhistory4y")
        rhistory5 = getFormVal("rhistory5")
        rhistory6 = getFormVal("rhistory6")
        parentID = getFormVal("PARENT_ID")

        logIt("start HHA2")
        logIt("parentID = " & parentID)
        logIt("rhistory1 = " & rhistory1)
        logIt("rhistory2 = " & rhistory2)
        logIt("rhistory3 = " & rhistory3)
        logIt("rhistory4 = " & rhistory4)
        logIt("rhistory5 = " & rhistory5)
        logIt("rhistory6 = " & rhistory6)

        cSql.SqlStr = "select MEMBERSHIP_ID from membership where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'LIC' and org_relate_id = ? " &_
            "and f_getmbrdetails(membership_id,'License_Type_LIC') = '2' and f_getmbrdetails(membership_id,'Status_LIC') = '2' "
        cSql.AddParam(parentID)
        set rs = cSql.Execute
        if rs.eof then
            logIt("no existing record found")
        else
            logIt("existing record " & rs("MEMBERSHIP_ID"))
            if pay = "CC" and rhistory1 = "No" and rhistory2 = "No" and rhistory3 = "No" and rhistory4 = "No" and rhistory5 = "No" and rhistory6 = "No" then
                logIt("scenario 1")
                cSql.SqlStr = "select membership_id,f_getmbrdetails(membership_id,'License_Type_LIC') License_type,f_getmbrdetails(membership_id,'Status_LIC') Status " &_
                    "from membership where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'LIC' and org_relate_id = ? " &_
                    "and f_getmbrdetails(membership_id,'License_Type_LIC') in ('1','3')"
                cSql.AddParam(parentID)
                set rs2 = cSql.Execute
                if rs2.eof then
                    logIt("no License_Type_LIC 1 or 3 record found")
                    if month(date) < 8 then
                        logIt("Jan - Jul")
                        cMember.UpdateValue rs("MEMBERSHIP_ID"),"LIC_EXPIRE_DT",getFormVal("LIC_EXP")
                    else
                        logIt("Aug - Dec")
                        cMember.UpdateValue rs("MEMBERSHIP_ID"),"LIC_EXPIRE_DT",getFormVal("LIC_EXP2")
                    end if
                else
                    logIt("License_Type_LIC 1 or 3 record " & rs2("MEMBERSHIP_ID"))
                    if rs2("Status") = "1" then
                        logIt("Status 1")
                        if rs2("License_type") = "1" then
                            logIt("License_Type_LIC 1")
                            cMember.UpdateValue rs("MEMBERSHIP_ID"),"LIC_EXPIRE_DT",getFormVal("CNA_EXPIRE")
                        end if
                        if rs2("License_type") = "3" then
                            logIt("License_Type_LIC 3")
                            cMember.UpdateValue rs("MEMBERSHIP_ID"),"LIC_EXPIRE_DT",getFormVal("NA_EXPIRE")
                        end if
                    end if
                    if rs2("Status") = "2" then
                        logIt("Status 2")
                        if month(date) < 8 then
                            logIt("Jan - Jul")
                            cMember.UpdateValue rs("MEMBERSHIP_ID"),"LIC_EXPIRE_DT",getFormVal("LIC_EXP")
                        else
                            logIt("Aug - Dec")
                            cMember.UpdateValue rs("MEMBERSHIP_ID"),"LIC_EXPIRE_DT",getFormVal("LIC_EXP2")
                        end if
                    end if
                end if
                rs2.close
                set rs2 = nothing

                cMember.UpdateValue rs("MEMBERSHIP_ID"),"LIC_APPROV_DT",getFormVal("INIT_DATE")
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"Status_LIC",getFormVal("REG_STATUS")
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"Status_Change_Date_LIC",getFormVal("STATUS_CHANGE")
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"Paid_Amount_MEMBER",getFormVal("CARD_AMOUNT")
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"Payment_Received_Date_MEMBER",getFormVal("PAY_RECVD")
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"Payment_Status_MEMBER",getFormVal("PAY_STATUS")
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"Transaction_Type_MEMBER",getFormVal("TRANS_TYPE")
            end if

            if pay = "OT" and rhistory1 = "No" and rhistory2 = "No" and rhistory3 = "No" and rhistory4 = "No" and rhistory5 = "No" and rhistory6 = "No" then
                logIt("scenario 2")
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"Status_LIC",getFormVal("REG_STATUS")
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"Status_Change_Date_LIC",getFormVal("STATUS_CHANGE")
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"Payment_Status_MEMBER",getFormVal("PAY_STATUSCK")
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"Transaction_Type_MEMBER",getFormVal("TRANS_TYPE")
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"LIC_APPROV_DT",""
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"LIC_EXPIRE_DT",""
            end if

            if pay = "CC" and (rhistory1 = "Yes" or rhistory2 = "Yes" or rhistory3 = "Yes" or rhistory4 = "Yes" or rhistory5 = "Yes" or rhistory6 = "Yes") then
                logIt("scenario 3")
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"Status_LIC",getFormVal("REG_STATUS")
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"Status_Change_Date_LIC",getFormVal("STATUS_CHANGE")
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"Paid_Amount_MEMBER",getFormVal("CARD_AMOUNT")
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"Payment_Status_MEMBER",getFormVal("PAY_STATUS")
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"Payment_Received_Date_MEMBER",getFormVal("PAY_RECVD")
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"Transaction_Type_MEMBER",getFormVal("TRANS_TYPE")
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"LIC_APPROV_DT",""
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"LIC_EXPIRE_DT",""
            end if

            if pay = "OT" and (rhistory1 = "Yes" or rhistory2 = "Yes" or rhistory3 = "Yes" or rhistory4 = "Yes" or rhistory5 = "Yes" or rhistory6 = "Yes") then
                logIt("scenario 4")
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"Status_LIC",getFormVal("REG_STATUS")
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"Status_Change_Date_LIC",getFormVal("STATUS_CHANGE")
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"Payment_Status_MEMBER",getFormVal("PAY_STATUSCK")
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"Transaction_Type_MEMBER",getFormVal("TRANS_TYPE")
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"LIC_APPROV_DT",""
                cMember.UpdateValue rs("MEMBERSHIP_ID"),"LIC_EXPIRE_DT",""
            end if
        end if
        rs.close
        set rs = nothing
        redirectAppRenew(request.querystring("id"))
    end if

    if appRenewType = "EMG_TRAIN" or appRenewType = "NA_RECRUIT" then
        newMemberID = membershipID
        if request("existInd")&"" <> "" then
            set cMember = nothing
            set cMember = new cMemberValuesO
            cMember.LoadFromMembershipID request("existInd")
            newLicID = getFormVal("CHILD_ID")
            approveExist request.querystring("id"), request("existInd"), cMember, rsAppRenew, newMemberID
            logIt("deleting " & newMemberID)
            cMember.UpdateValue newMemberID,"M_STAT_CD","INACTIVE"

            if request("existLic")&"" <> "" then
                set cLic = new cMemberValuesO
                cLic.LoadFromMembershipID request("existLic")
                approveExist request.querystring("id"), request("existLic"), cLic, rsAppRenew, newLicID
                set cLic = nothing
                logIt("deleting " & newLicID)
                cMember.UpdateValue newLicID,"M_STAT_CD","INACTIVE"
            else
                logIt("activating " & newLicID)
                cMember.UpdateValue newLicID,"M_STAT_CD","ACTIVE"
                logIt("setting org_relate_id to " & request("existInd"))
                cMember.UpdateValue newLicID,"ORG_RELATE_ID",request("existInd")
            end if
        else
            cMember.UpdateValue getFormVal("CHILD_ID"),"M_STAT_CD","ACTIVE"
            cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
        end if
        set cMember = nothing
        redirectAppRenew(request.querystring("id"))
    end if

    if appRenewType = "NA_APP" or appRenewType = "CNA_APP" or appRenewType = "HHA_APP" then
        newLicID = membershipID
        cSql.SqlStr = "select data from app_renew_detail where app_renew_id = ? and upper(field_name) = 'PARENT_ID'"
        cSql.AddParam(request.querystring("id"))
        set rs = cSql.Execute
        if rs.eof then
        else
            newMemberID = rs("DATA")
            cMember.LoadFromMembershipID newMemberID
            if cMember.Value("ORG_SPEC_ID")&"" = "" then
                cMember.UpdateValue newMemberID,"ORG_SPEC_ID",getSequenceMin("SEQ_LICENSE_NO",1)
            end if
        end if

        if appRenewType = "NA_APP" or appRenewType = "HHA_APP" then
            cSql.SqlStr = "select app_renew_id from app_renew_detail where app_renew_id = ? and lower(field_name) in ('ques1','ques2','ques3','ques4','ques5','ques6') and data = 'Yes'"
            cSql.AddParam(request.querystring("id"))
            set rs = cSql.Execute
            if rs.eof then
            else
                cMember.Updatevalue membershipID,"Status_LIC","22"
            end if
        end if

        if request("existInd")&"" <> "" then
            cMember.LoadFromMembershipID request("existInd")
            approveExist request.querystring("id"), request("existInd"), cMember, rsAppRenew, newMemberID
            logIt("deleting " & newMemberID)
            cMember.UpdateValue newMemberID,"M_STAT_CD","INACTIVE"

            if request("existLic")&"" <> "" then
                set cLic = new cMemberValuesO
                cLic.LoadFromMembershipID request("existLic")
                approveExist request.querystring("id"), request("existLic"), cLic, rsAppRenew, newLicID
                set cLic = nothing
                logIt("deleting " & newLicID)
                cMember.UpdateValue newLicID,"M_STAT_CD","INACTIVE"
            else
                logIt("activating " & newLicID)
                cMember.UpdateValue newLicID,"M_STAT_CD","ACTIVE"
                logIt("setting org_relate_id to " & request("existInd"))
                cMember.UpdateValue newLicID,"ORG_RELATE_ID",request("existInd")
            end if
        else
            cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
        end if
        set cMember = nothing
        redirectAppRenew(request.querystring("id"))
    end if
end if

if Application("ORG_ID") = "1450" and appRenewType = "COVID-19_APP" then
    orgSpecID = ""
    permitType = getFormVal("permitTyp")
    set cMember = new cMemberValuesO
    cMember.LoadFromMembershipId membershipID
    if permitType = "IN" then
        orgSpecID = "IN" & getSequenceMin("SEQ_IN_LIC_TYPE_ID",7)
    elseif permitType = "PT" then
        orgSpecID = "PT" & getSequenceMin("SEQ_PT_LIC_TYPE_ID",7)
    elseif permitType = "RP" then
        orgSpecID = "RP" & getSequenceMin("SEQ_RP_LIC_TYPE_ID",7)
    end if
    cMember.UpdateValue membershipID,"ORG_SPEC_ID",orgSpecID
    cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
    if request("existInd")&"" <> "" then
        cSql.SqlStr = "select org_relate_id from membership where membership_id = ?" 
        cSql.AddParam(membershipID)
        set rs = cSql.Execute
        newParentID = rs("ORG_RELATE_ID")
        cMember.LoadFromMembershipID request("existInd")
        approveExist request.querystring("id"), request("existInd"), cMember, rsAppRenew, newParentID
        cMember.UpdateValue newParentID,"M_STAT_CD","INACTIVE"
        logIt("deleting " & newParentID)
    end if
end if

if Application("ORG_ID") = "1450" and appRenewType <> "COVID-19_APP" then
    dim arrWVBOPSeqs(10,2)
    arrWVBOPSeqs(0,0) = "RP_APP"
    arrWVBOPSeqs(0,1) = "LIC_NO"
    arrWVBOPSeqs(0,2) = "SEQ_RP_LIC_TYPE_ID"

    arrWVBOPSeqs(1,0) = "CCP_APP"
    arrWVBOPSeqs(1,1) = "LIC_NO_CC"
    arrWVBOPSeqs(1,2) = "SEQ_CC_CRED_TYPE_ID"

    arrWVBOPSeqs(2,0) = "LP_APP"
    arrWVBOPSeqs(2,1) = "LIC_NO"
    arrWVBOPSeqs(2,2) = "SEQ_LP_CRED_TYPE_ID"

    arrWVBOPSeqs(3,0) = "MOP_APP"
    arrWVBOPSeqs(3,1) = "LIC_NO"
    arrWVBOPSeqs(3,2) = "SEQ_MO_CRED_TYPE_ID"

    arrWVBOPSeqs(4,0) = "MFR_APP"
    arrWVBOPSeqs(4,1) = "LIC_NO"
    arrWVBOPSeqs(4,2) = "SEQ_MR_CRED_TYPE_ID"

    arrWVBOPSeqs(5,0) = "TPLP_APP"
    arrWVBOPSeqs(5,1) = "LIC_NO"
    arrWVBOPSeqs(5,2) = "SEQ_3PL_CRED_TYPE_ID"

    arrWVBOPSeqs(6,0) = "WDD_APP"
    arrWVBOPSeqs(6,1) = "LIC_NO"
    arrWVBOPSeqs(6,2) = "SEQ_WD_CRED_TYPE_ID"

    arrWVBOPSeqs(7,0) = "IN_APP"
    arrWVBOPSeqs(7,1) = "LIC_NO"
    arrWVBOPSeqs(7,2) = "SEQ_IN_LIC_TYPE_ID"

    arrWVBOPSeqs(8,0) = "NT_APP"
    arrWVBOPSeqs(8,1) = "LIC_NO"
    arrWVBOPSeqs(8,2) = "SEQ_NT_LIC_TYPE_ID"

    arrWVBOPSeqs(9,0) = "PT_APP"
    arrWVBOPSeqs(9,1) = "LIC_NO"
    arrWVBOPSeqs(9,2) = "SEQ_PT_LIC_TYPE_ID"

    arrWVBOPSeqs(10,0) = "TT_APP"
    arrWVBOPSeqs(10,1) = "LIC_NO"
    arrWVBOPSeqs(10,2) = "SEQ_TT_LIC_TYPE_ID"

    for i = lbound(arrWVBOPSeqs,1) to ubound(arrWVBOPSeqs,1)
        if appRenewType = arrWVBOPSeqs(i,0) then
            cSql.SqlStr = "select DATA from app_renew_detail where app_renew_id = ? and lower(field_name) = lower(?)"
            cSql.AddParam(request.querystring("id"))
            cSql.AddParam(arrWVBOPSeqs(i,1))
            set rs = cSql.Execute
            if rs.eof then
            else
                if appRenewType = "PT_APP" or appRenewType = "NT_APP" then
                    cMember.Updatevalue getFormVal("CHILD_ID"),"ORG_SPEC_ID",rs("DATA") & getSequenceMin(arrWVBOPSeqs(i,2),7)
                else
                    if request("exist") <> "" then
                        cMember.UpdateValue request("exist"),"ORG_SPEC_ID",rs("DATA") & getSequenceMin(arrWVBOPSeqs(i,2),7)
                    else
                        cMember.UpdateValue membershipID,"ORG_SPEC_ID",rs("DATA") & getSequenceMin(arrWVBOPSeqs(i,2),7)
                    end if
                end if
            end if
            rs.close
            set rs = nothing
        end if
    next

    if appRenewType = "CSP_APP" or appRenewType = "PHAR_APP" then
        cSql.SqlStr = "select DATA from app_renew_detail where app_renew_id = ? and upper(field_name) = 'LICENSE_NUMBER'"
        cSql.AddParam(request.querystring("id"))
        set rs = cSql.Execute
        if rs.eof then
        else
            if request("exist") <> "" then
                cMember.UpdateValue request("exist"),"ORG_SPEC_ID",rs("DATA") & getSequenceMin("SEQ_" & rs("DATA") & "_CRED_TYPE_ID",7)
            else
                cMember.UpdateValue membershipID,"ORG_SPEC_ID",rs("DATA") & getSequenceMin("SEQ_" & rs("DATA") & "_CRED_TYPE_ID",7)
            end if
        end if
        rs.close
        set rs = nothing
    end if

    if appRenewType = "NT_APP" or appRenewType = "TT_APP" or appRenewType = "IN_APP" or appRenewType = "RP_APP" or appRenewType = "PT_APP" then
        approveUpdate(request.querystring("id"))
        if request("existInd")&"" <> "" then
            cSql.SqlStr = "select org_relate_id from membership where membership_id = ?" 
            cSql.AddParam(rsAppRenew("MEMBERSHIP_ID"))
            set rs = cSql.Execute
            newParentID = rs("ORG_RELATE_ID")
            rs.close
            set rs = nothing
            cMember.LoadFromMembershipID request("existInd")
            approveExist request.querystring("id"), request("existInd"), cMember, rsAppRenew, newParentID
            cMember.UpdateValue newParentID,"M_STAT_CD","INACTIVE"
            logIt("deleting " & newParentID)
            if appRenewType = "TT_APP" then
                logIt("activating " & membershipID)
                cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
                logIt("settiing org_relate_id to " & request("existInd"))
                cMember.UpdateValue membershipID,"ORG_RELATE_ID",request("existInd")
            else
                if request("existLic")&"" <> "" then
                    set cLic = new cMemberValuesO
                    cLic.LoadFromMembershipID request("existLic")
                    approveExist request.querystring("id"), request("existLic"), cLic, rsAppRenew, 0
                    set cLic = nothing
                    logIt("deleting " & rsAppRenew("MEMBERSHIP_ID"))
                    cMember.UpdateValue rsAppRenew("MEMBERSHIP_ID"),"M_STAT_CD","INACTIVE"
                else
                    logIt("activating " & membershipID)
                    cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
                    logIt("settiing org_relate_id to " & request("existInd"))
                    cMember.UpdateValue membershipID,"ORG_RELATE_ID",request("existInd")
                end if
            end if
            redirectAppRenew(request.querystring("id"))
        else
            cMember.UpdateValue getFormVal("CHILD_ID"),"M_STAT_CD","ACTIVE"
            cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
        end if
    end if

    if appRenewType = "CCP_APP" or appRenewType = "MOP_APP" or appRenewType = "WDD_APP" or appRenewType = "TPLP_APP" or appRenewType = "MFR_APP" or appRenewType = "LP_APP" or appRenewType = "PHAR_APP" or appRenewType = "CSP_APP" then
        if request("exist")&"" <> "" then
            if request("existFac")&"" <> "" then
                cSql.SqlStr = "select org_relate_id from membership where membership_id = ?"
                cSql.AddParam(newMembershipID)
                set rs = cSql.Execute
                newFacId = rs("ORG_RELATE_ID")
                rs.close
                set rs = nothing

                logIt("newMembershipID = " & newMembershipID)
                logIt("newFacId = " & newFacId)
                logIt("existFac = " & request("existFac"))
            end if

            if appRenewType = "CSP_APP" or appRenewType = "TPLP_APP" or appRenewType = "WDD_APP" or appRenewType = "MOP_APP" or appRenewType = "MFR_APP" or appRenewType = "LP_APP" or appRenewType = "PHAR_APP" then
                cSql.SqlStr = "select membership_id from membership where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'Sub' and (org_relate_id = ? or org_relate_id_2 = ?)"
                cSql.AddParam(request("exist"))
                cSql.AddParam(request("exist"))
                set rs = cSql.Execute
                if rs.eof then
                else
                    existSubID = rs("MEMBERSHIP_ID")
                end if
                rs.close
                set rs = nothing

                cSql.SqlStr = "select membership_id from membership where org_sub_ty_cd = 'Sub' and (org_relate_id = ? or org_relate_id_2 = ?)"
                cSql.AddParam(rsAppRenew("MEMBERSHIP_ID"))
                cSql.AddParam(rsAppRenew("MEMBERSHIP_ID"))
                set rs = cSql.Execute
                if rs.eof then
                else
                    newSubID = rs("MEMBERSHIP_ID")
                end if
                rs.close
                set rs = nothing

                logIt("newSubID = " & newSubID)
                logIt("existSubID = " & existSubID)

                if newSubID&"" <> "" and existSubID&"" <> "" then
                    set cSub = new cMemberValuesO
                    cSub.LoadFromMembershipID existSubID
                    approveExist request.querystring("id"), existSubID, cSub, rsAppRenew, newSubID
                    cSub.UpdateValue newSubID,"M_STAT_CD","INACTIVE"
                    set cSub = nothing
                end if
            end if

            approveExist request.querystring("id"), request("exist"), cMember, rsAppRenew, 0
            cMember.UpdateValue rsAppRenew("MEMBERSHIP_ID"),"M_STAT_CD","INACTIVE"

            if request("existFac")&"" <> "" then
                set cFac = new cMemberValuesO
                cFac.LoadFromMembershipID request("existFac")
                approveExist request.querystring("id"), request("existFac"), cFac, rsAppRenew, newFacId
                set cFac = nothing
                cMember.UpdateValue newFacId,"M_STAT_CD","INACTIVE"
            end if
        else
            cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
        end if
    end if
end if

'if Application("ORG_ID") = "375" and appRenewType <> "PA_RENEW"  then
if Application("ORG_ID") = "375" then
    cSql.SqlStr = "select app_renew_fields.column_name, app_renew_detail.data " &_
	    "from app_renew_detail, app_renew_fields, app_renew " &_
	    "where upper(app_renew_fields.field_name) = upper(app_renew_detail.field_name) " &_
	    "and app_renew.app_renew_id = app_renew_detail.app_renew_id " &_
	    "and app_renew_fields.app_renew_type = app_renew.app_renew_type " &_
	    "and app_renew_fields.column_name is not null " &_
	    "and app_renew_fields.update_record = 'Y' " &_
	    "and app_renew_detail.app_renew_id = ?"
    cSql.AddParam(request.querystring("id"))
    set rsAppRenewDetail = cSql.Execute

    if rsAppRenewDetail.eof and rsAppRenewDetail.bof then
    else
	    arrAppRenewDetail = rsAppRenewDetail.GetRows
    end if

    if isArray(arrAppRenewDetail) then
	    for i = lbound(arrAppRenewDetail,2) to ubound(arrAppRenewDetail,2)
    '		response.write arrAppRenewDetail(0,i) & " " & arrAppRenewDetail(1,i) & " " & cMember.Value(arrAppRenewDetail(0,i))&"<Br>"
		    if arrAppRenewDetail(1,i) & "" <> cMember.Value(arrAppRenewDetail(0,i))&"" then
			    cMember.updateValue membershipID,arrAppRenewDetail(0,i),arrAppRenewDetail(1,i)
		    end if
	    Next
    end if


    rsAppRenewDetail.close
    set rsAppRenewDetail = nothing
end if
dim curExpDt, newExpDt

'wvpebd
if Application("ORG_ID") = "775" then
	if appRenewType = "REAPP_PE_APP" then
        approveExist request.querystring("id"), membershipID, cMember, rsAppRenew, 0
    end if
	if appRenewType = "REINSTATE_PE_APP" then
        approveExist request.querystring("id"), membershipID, cMember, rsAppRenew, 0
    end if

	if appRenewType = "PE App" or appRenewType = "COA_App" or appRenewType = "EI_App" then
        if request("exist") <> "" then
            approveExist request.querystring("id"), request("exist"), cMember, rsAppRenew, 0
        else
		    cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"M_STAT_CD","ACTIVE"
		end if
	end if
end if

'WVBM
if Application("ORG_ID") = "725" then
    if appRenewType = "CORP_APP" or appRenewType = "PLLC_APP" then
		cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"M_STAT_CD","ACTIVE"
    end if
    if appRenewType = "ILDR" or appRenewType = "DPM_APP" or appRenewType = "PA_APP" then
        if request("exist") <> "" then
            approveExist request.querystring("id"), request("exist"), cMember, rsAppRenew, 0
            cMember.UpdateValue getFormValByID(request.querystring("id"),"MEMBERSHIP_ID"),"M_STAT_CD","INACTIVE"
            redirectAppRenew(request.querystring("id"))
        else
    		cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"M_STAT_CD","ACTIVE"
        end if
    end if
end if

'WV
if Application("ORG_ID") = "375" then
	if appRenewType = "UAAN" then
		cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"M_STAT_CD","ACTIVE"
    end if

	if appRenewType = "PLLC_APP" then
		cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"M_STAT_CD","ACTIVE"
		cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"Application_Date_PLLC",date
	end if

	if appRenewType = "CORP_APP" then
		cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"M_STAT_CD","ACTIVE"
		cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"Application_Date_Corporations",date
	end if

	if appRenewType = "DO_APP" then
		cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"M_STAT_CD","ACTIVE"
		cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"Application_Date_BoardXP",date
	end if
	
	if appRenewType = "PA_APP" then
		cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"M_STAT_CD","ACTIVE"
		cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"App_Date_Physician_Assistants",date
	end if
	
	if appRenewType = "RES_APP" then
		cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"M_STAT_CD","ACTIVE"
		cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"ED_Permit_Year_Residents","A"
		cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"Beginning_Date_Residents", date
	end if
	
	if appRenewType = "DO_RENEW" then
		curExpDt = cMember.Value("Lic_Exp_Date_BoardXP")
		if curExpDt & "" <> "" then
			if isDate(curExpDt) then
				newExpDt = dateAdd("yyyy",2,curExpDt)
				cMember.updateValue rsAppRenew("MEMBERSHIP_ID"), "Lic_Exp_Date_BoardXP", newExpDt
			end if
		end if
		cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"Lic_Renewal_Date__BoardXP",date
		'cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"Date_Paid_BoardXP",date

        dim doCSL
        
        cSql.SqlStr = "select DATA from app_renew_detail where app_renew_id = ? and upper(field_name) = 'CSLIC' and upper(data) = 'Y'"
        cSql.AddParam(request.querystring("id"))
        set rsAppRenewDetail = cSql.Execute
        if rsAppRenewDetail.eof then
            doCSL = false
        else
            doCSL = true
        end if
        rsAppRenewDetail.close
        set rsAppRenewDetail = nothing
        
        if doCSL then
		    if month(date) < 4 then
			    cMember.updateValue rsAppRenew("MEMBERSHIP_ID"), "CSL_Exp_Date__BoardXP", "06/30/" & year(date)
		    else
			    cMember.updateValue rsAppRenew("MEMBERSHIP_ID"), "CSL_Exp_Date__BoardXP", "06/30/" & year(date)+1
		    end if
		    cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"CSL_Renewal_Date_BoardXP",date
        end if
	end if
	
	if appRenewType = "PA_RENEW" then
		curExpDt = cMember.Value("Expiration__Physician_Assistants")
		if curExpDt & "" <> "" then
			if isDate(curExpDt) then
				newExpDt = dateAdd("yyyy",2,curExpDt)
				cMember.updateValue rsAppRenew("MEMBERSHIP_ID"), "Expiration__Physician_Assistants", newExpDt
			end if
		end if
		cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"Date_Paid_Physician_Assistants",date
	end if
	
	if appRenewType = "CSL_RENEW" then
		if month(date) < 4 then
			cMember.updateValue rsAppRenew("MEMBERSHIP_ID"), "CSL_Exp_Date__BoardXP", "06/30/" & year(date)
		else
			cMember.updateValue rsAppRenew("MEMBERSHIP_ID"), "CSL_Exp_Date__BoardXP", "06/30/" & year(date)+1
		end if
		cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"CSL_Renewal_Date_BoardXP",date
	end if
	
	if appRenewType = "CSL_APP" then
		if month(date) < 4 then
			cMember.updateValue rsAppRenew("MEMBERSHIP_ID"), "CSL_Exp_Date__BoardXP", "06/30/" & year(date)
		else
			cMember.updateValue rsAppRenew("MEMBERSHIP_ID"), "CSL_Exp_Date__BoardXP", "06/30/" & year(date)+1
		end if
		cMember.updateValue rsAppRenew("MEMBERSHIP_ID"), "WVCSLOrgIssueDate_BoardXP",date
	end if
	
	if appRenewType = "CORP_RENEW" then
		curExpDt = cMember.Value("Expiration_Date_Corporations")
		if curExpDt & "" <> "" then
			if isDate(curExpDt) then
				newExpDt = dateAdd("yyyy",2,curExpDt)
				cMember.updateValue rsAppRenew("MEMBERSHIP_ID"), "Expiration_Date_Corporations", newExpDt
			end if
		end if
		cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"Date_Paid_Corporations",date
	end if
	
	if appRenewType = "PLLC_RENEW" then
		curExpDt = cMember.Value("Expiration_Date_PLLC")
		if curExpDt & "" <> "" then
			if isDate(curExpDt) then
				newExpDt = dateAdd("yyyy",1,curExpDt)
				cMember.updateValue rsAppRenew("MEMBERSHIP_ID"), "Expiration_Date_PLLC", newExpDt
			end if
		end if
		cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"Date_Paid_PLLC",date
	end if
	
	if appRenewType = "RES_RENEW" then
		if cMember.Value("Beginning_Date_Residents") = "" then
			cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"Beginning_Date_Residents","07/01/" & year(date) + 1
			cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"End_Date__Residents","06/30/" & year(date) + 1
		elseif cMember.Value("First_Renewal_Begin_Date_Residents") = "" then
			cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"First_Renewal_Begin_Date_Residents","07/01/" & year(date)
			cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"First_Renewal_End_Date_Residents","06/30/" & year(date) + 1
		elseif cMember.Value("Second_Rnewal_Begin_Date_Residents") = "" then
			cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"Second_Rnewal_Begin_Date_Residents","07/01/" & year(date)
			cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"Second_Renewal_End_Date_Residents","06/30/" & year(date) + 1
		elseif cMember.Value("Third_Renewal_Begin_Dat_Residents") = "" then
			cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"Third_Renewal_Begin_Dat_Residents","07/01/" & year(date)
			cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"Third_Renewal_End_Date_Residents","06/30/" & year(date) + 1
		elseif cMember.Value("Fourth_Renewal_Begin_Date_Residents") = "" then
			cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"Fourth_Renewal_Begin_Date_Residents","07/01/" & year(date)
			cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"Fourth_Renewal_End_Date_Residents","06/30/" & year(date) + 1
		end if
		
		cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"Date_Paid_Residents",date
		
		if cMember.Value("ED_Permit_Year_Residents") = "A" then
			cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"ED_Permit_Year_Residents","B"
		end if
		if cMember.Value("ED_Permit_Year_Residents") = "B" then
			cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"ED_Permit_Year_Residents","C"
		end if
		if cMember.Value("ED_Permit_Year_Residents") = "C" then
			cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"ED_Permit_Year_Residents","E"
		end if
		if cMember.Value("ED_Permit_Year_Residents") = "D" then
			cMember.updateValue rsAppRenew("MEMBERSHIP_ID"),"ED_Permit_Year_Residents","D"
		end if
	end if
end if

acceptAwait

MEMBERSHIP_ID = rsAppRenew("MEMBERSHIP_ID")

rsAppRenew.close
set rsAppRenew = nothing
set cSql = nothing

approveUpdate(request.querystring("id"))

if Application("ORG_ID") = "1450" then
    cMember.LoadFromMembershipID MEMBERSHIP_ID
    if cMember.Value("ORG_SUB_TY_CD") = "LIC" and cMember.Value("cpeActivity_Date_LIC")&"" = "" then
        if cMember.Value("ORG_RELATE_ID")&"" <> "" and cMember.Value("NABP_e_Profile_ID_LIC")&"" <> "" then
            wvbopEProfileID = cMember.Value("NABP_e_Profile_ID_LIC")
            sqlStr = "select birth_dt from membership where membership_id = " & cMember.Value("ORG_RELATE_ID")
            set rsEmpty = connOracle.Execute(sqlStr)
            if rsEmpty.eof then
            else
                wvbopBirthDt = rsEmpty("birth_dt")
                wvbopBirthDt = month(wvbopBirthDt) & "-" & day(wvbopBirthDt) & "-" & year(wvbopBirthDt)
            end if
            rsEmpty.close
            set rsEmpty = nothing
            if wvbopBirthDt&"" <> "" then
                nabpURL = Application("NABP_URL") & "/api/nabp/getprofile/" & wvbopEProfileID & "/" & wvbopBirthDt & "/07-01-" & year(date)-2 & "/06-30-" & year(date)

		        Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
		        oXMLHTTP.open "POST",nabpURL,False
		        oXMLHTTP.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
                oXMLHTTP.setRequestHeader "Accept","application/xml"
		        oXMLHTTP.send()
		        strResponse = oXMLHTTP.responseText
		        set oXMLHTTP = nothing
                set XmlDoc = server.CreateObject("Microsoft.XMLDOM")
                xmlDoc.LoadXML(strResponse)

                dim arrNABP(1,9)
                arrNABP(0,0) = "scores/applicationId"
                arrNABP(1,0) = "scoresAppID_LIC"
                arrNABP(0,1) = "scores/examDate"
                arrNABP(1,1) = "scoresExamDate_LIC"
                arrNABP(0,2) = "scores/attemptNumber"
                arrNABP(1,2) = "scoresAttemptNo_LIC"
                arrNABP(0,3) = "disciplined"
                arrNABP(1,3) = "NABP_Discipline_LIC"
                arrNABP(0,4) = "scores/result"
                arrNABP(1,4) = "scoresResult_LIC"
                arrNABP(0,5) = "scores/exam"
                arrNABP(1,5) = "scoresExam_LIC"
                arrNABP(0,6) = "scores/postDate"
                arrNABP(1,6) = "scoresPostDate_LIC"
                arrNABP(0,7) = "invalidatedScores/examName"
                arrNABP(1,7) = "Exam_Name_LIC"
                arrNABP(0,8) = "invalidatedScores/state"
                arrNABP(1,8) = "Exam_State_LIC"
                arrNABP(0,9) = "invalidatedScores/examDate"
                arrNABP(1,9) = "Invalid_Exam_Date_LIC"

                for j = lbound(arrNABP,2) to ubound(arrNABP,2)
                    set nodes = xmlDoc.selectNodes("//root/scoreResults/" & arrNABP(0,j) ) 
                    for i = 0 to nodes.length-1
                        if j < 10 then
                            cMember.AppendValue MEMBERSHIP_ID,arrNABP(1,j),nodes(i).text
                        else
                            cMember.UpdateValue MEMBERSHIP_ID,arrNABP(1,j),nodes(i).text
                        end if
                    next
                    set nodes = nothing
                next

                set xmlDoc = nothing
            end if
        end if
    end if
end if

if Application("ORG_ID") = "60" then
    cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
    if appRenewType = "Renew" then
        cMember.LoadFromMembershipID membershipID
        email = cMember.Value("EMAIL")
        if Application("DEV_DOMAIN") = request.ServerVariables("SERVER_NAME") then
            email = "melissa@ebigpicture.com"
        end if
        URL = "https://" & request.ServerVariables("SERVER_NAME") & "/db/licensewriter.asp?ty=doc2&LABELS="&membershipID&"&membership_id="&membershipID&"&email=" & email
        Set oXMLHTTP = CreateObject("Msxml2.ServerXMLHTTP.6.0")
        oXMLHTTP.open "POST",URL,False
        oXMLHTTP.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
        oXMLHTTP.setRequestHeader "Accept", "application/xml"
        oXMLHTTP.send()
        set  oXMLHTTP = nothing  
    end if
end if

if Application("ORG_ID") = "30" then
    cMember.UpdateValue membershipID,"M_STAT_CD","ACTIVE"
    if appRenewType = "PLFM" then
        cMember.UpdateValue membershipID,"ORG_SPEC_ID",getSequenceMin("SEQ_PVR_CASE_NO",1)
    end if
    if appRenewType = "RENEW" or appRenewType = "UAP/MA_RENEW" then
        set cSql = new cRunSql
        cSql.conn = connOracle
        cSql.SqlStr = "select DATA from app_renew_detail where app_renew_id = ? and field_name = 'ISSUE_DT'"
        cSql.AddParam(request.querystring("id"))
        'cSql.Debug
        set rs = cSql.Execute
        if rs.eof then
        else
            if rs("DATA")&"" <> "" then
                cMember.UpdateValue membershipID,"ISSUE_DATE",date
            end if
        end if
        rs.close
        set rs = nothing
        set cSql = nothing
    end if

    if appRenewType = "RENEW" then
        set cSql = new cRunSql
        cSql.conn = connOracle
        cSql.SqlStr = "select DATA from app_renew_detail where app_renew_id = ? and field_name = 'APRN_ISSUE_DT'"
        cSql.AddParam(request.querystring("id"))
        'cSql.Debug
        set rs = cSql.Execute
        if rs.eof then
        else
            if rs("DATA")&"" <> "" then
                cMember.UpdateValue membershipID,"APRN_ISSUE_DATE",date
            end if
        end if
        rs.close
        set rs = nothing
        set cSql = nothing
    end if
end if

set cMember = nothing

Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
oXMLHTTP.open "POST","https://" & request.ServerVariables("SERVER_NAME") & "/db/live/form_approve_deny_newsletter.asp",True
oXMLHTTP.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
oXMLHTTP.send("form_id=" & request.querystring("id") & "&action=A")
set oXMLHTTP = nothing

redirectAppRenew(request.querystring("id"))
%>
<%end sub%>

<%sub mainMenu%>

<%
    dim recsPerPage, recStart, recEnd, numPages, pageNo
    recsPerPage = 100
%>

<script language="javascript">
    // doDelete
    function doDelete() {
        var isChecked = false;
        var multiChecked = false;

        if (document.form.deleteView) {
            if (document.form.deleteView.length) {
                for (x = 0; x < document.form.deleteView.length; x++) {
                    if (document.form.deleteView[x].checked) {
                        if (isChecked) {
                            multiChecked = true;
                        } else {
                            isChecked = true;
                        }
                    }
                }
            } else {
                if (document.form.deleteView.checked)
                    isChecked = true;
            }
        }
        
        if (multiChecked) {
            if (confirm("Are you sure you want to permanently delete these forms?")) {
                document.form.action.value = "deleteView";
                document.form.submit();
            }
        } else if (isChecked) {
            confirmMsg = "Are you sure you want to permanently delete this form?"

            if (confirm(confirmMsg)) {
                document.form.action.value = "deleteView";
                document.form.submit();
            }
        } else {
        alert("You didn't select any forms. Click the check box next to the forms you want to delete.");
        }
    }

    // doApprove
    function doApprove() {
        var isChecked = false;
        var multiChecked = false;

        if (document.form.deleteView) {
            if (document.form.deleteView.length) {
                for (x = 0; x < document.form.deleteView.length; x++) {
                    if (document.form.deleteView[x].checked) {
                        if (isChecked) {
                            multiChecked = true;
                        } else {
                            isChecked = true;
                        }
                    }
                }
            } else {
                if (document.form.deleteView.checked)
                    isChecked = true;
            }
        }
        
        if (multiChecked) {
            if (confirm("Are you sure you want to approve these forms?")) {
                document.form.action.value = "approveView";
                document.form.submit();
            }
        } else if (isChecked) {
            confirmMsg = "Are you sure you want to approve this form?"

            if (confirm(confirmMsg)) {
                document.form.action.value = "approveView";
                document.form.submit();
            }
        } else {
            alert("You didn't select any forms. Click the check box next to the forms you want to approve.");
        }
    }

    // toggleAll
    function toggleAll() {

        if (document.form.deleteViews.checked) {
            checkStat = true;
        } else {
            checkStat = false;
        }

        if (document.form.deleteView.length) {
            for (x = 0; x < document.form.deleteView.length; x++) {
                document.form.deleteView[x].checked = checkStat;
            }
        } else {
            document.form.deleteView.checked = checkStat;
        }
    }

    function doSearch() {
        document.getElementById("action").value = "doSearch";
        document.getElementById("form").submit();
    }

    function doSearchLogs() {
        document.getElementById("action").value = "doSearchLogs";
        document.getElementById("form").submit();
    }
</script>
<%
dim switch, lineClass, rsAppRenewCodes, appRenewCode, appRenewStatus, arrRS, searchVal, hideBot
dim cMem, cPar, cChi, idxMem
dim arrMem, arrPar, arrChi
dim rsAU, arrAU, idxAU, hasAU, arrAwait

sqlStr = getCodeValuesO("APP_RENEW_TYPE_CD")
set rsAppRenewCodes = connOracle.execute(sqlStr)

if rsAppRenewCodes.eof and rsAppRenewCodes.bof then
else
	arrCodes = rsAppRenewCodes.GetRows
	rsAppRenewCodes.MoveFirst
	appRenewCode = rsAppRenewCodes("code_value")
	rsAppRenewCodes.MoveFirst
end if

if len(request("app_renew_cd")) > 0 then
	appRenewCode = request("app_renew_cd")
    session("app_renew_cd") = appRenewCode
elseif session("app_renew_cd") <> "" then
    appRenewCode = session("app_renew_cd")
end if

appRenewStatus = "PENDING"

if request("action") = "doSearch" then
    searchVal = request("search")
end if

if len(request("app_renew_status")) > 0 then
	appRenewStatus = request("app_renew_status")
end if

hideBot = "N"
if session("HIDE_BOT")&"" <> "" then
    hideBot = session("HIDE_BOT")
end if
if request.form("hide_bot")&"" <> "" then
    hideBot = request.form("hide_bot")
    session("HIDE_BOT") = hideBot
end if

dim showDeleteViews
showDeleteViews = true
if Application("MASS_APPROVE_FORMS") = "1" and appRenewStatus = "PENDING" then
elseif (appRenewStatus = "PENDING" and not officeIP) or appRenewStatus = "COMPLETED" or appRenewStatus = "APPROVED" or appRenewStatus = "DELETED" then
    showDeleteViews = false
end if

set cSql = new cRunSql
cSql.Conn = connoracle
cSql.DisplayErrors = true
cSql.SqlStr = "select a.membership_id,a.app_renew_id,a.name, (select code_value_desc from code_value where code_class = 'APP_RENEW_TYPE_CD' and code_value = a.app_renew_type) app_renew_type, a.create_dt, nvl(a.modify_dt,a.create_dt) mod_dt, a.app_renew_status " &_
    ",(select max(s.create_dt) from app_renew_page_submit s where s.app_renew_id = a.app_renew_id) submit_dt, a.app_renew_type type_code " &_
    ",nvl((select data from app_renew_detail where lower(field_name) = 'card_amount' and app_renew_detail.app_renew_id = a.app_renew_id),'') card_amount " &_
    ",'' await " &_
    ",nvl((select min(data) from app_renew_detail where lower(field_name) = 'child_id' and app_renew_detail.app_renew_id = a.app_renew_id),'0') child_id " &_
    ",nvl((select min(data) from app_renew_detail where lower(field_name) = 'parent_id' and app_renew_detail.app_renew_id = a.app_renew_id),'0') parent_id " &_
    ",nvl((select min(data) from app_renew_detail where lower(field_name) = 'payment_status' and app_renew_detail.app_renew_id = a.app_renew_id),'') payment_status " &_
    ",nvl((select min(data) from app_renew_detail where lower(field_name) = 'payment_date_time' and app_renew_detail.app_renew_id = a.app_renew_id),'') payment_date_time " &_
	"from app_renew a where 1 = 1 "
if appRenewStatus&"" <> "" and appRenewStatus <> "ALL" then
    cSql.AddSqlStr(" and app_renew_status = ? ")
    cSql.AddParam(appRenewStatus)
end if
if appRenewCode&"" <> "" and appRenewCode&"" <> "ALL" then
    cSql.AddSqlStr(" and app_renew_type = ? ")
    cSql.AddParam(appRenewCode)
end if
if request("action") = "doSearch" then
    cSql.AddSqlStr("and a.app_renew_id in (")
    cSql.AddSqlStr("select app_renew_id from app_renew where upper(name) like upper(?)")
    cSql.AddParam("%" & searchVal & "%")
    cSql.AddSqlStr("union ")
    cSql.AddSqlStr("select app_renew_id from app_renew_detail where upper(data) like upper(?)")
    cSql.AddParam("%" & searchVal & "%")
    cSql.AddSqlStr("union ")
    cSql.AddSqlStr("select app_renew_id from app_renew_page_submit, app_renew_page_sub_detail where app_renew_page_sub_detail.app_renew_page_submit_id = app_renew_page_submit.app_renew_page_submit_id and upper(app_renew_page_sub_detail.data) like upper(?)")
    cSql.AddParam("%" & searchVal & "%")
    cSql.AddSqlStr(")")
end if
if appRenewStatus = "PROGRESS" or appRenewStatus = "ALL" then
    if hideBot = "Y" then
        cSql.AddSqlStr("and app_renew_id not in (select app_renew_id from app_renew where app_renew_status = 'PROGRESS' and membership_id = 0 minus select app_renew_id from app_renew_page_submit) ")
    end if
end if
if request("sort")&"" <> "" then
    cSql.AddSqlStr("order by " & request("sort") & " " & request("order"))
else
    cSql.AddSqlStr("order by mod_dt desc, create_dt desc")
end if
set rs = cSql.Execute
if rs.eof then
else
    arrRS = rs.GetRows
    maxPosition = ubound(arrRS,2) + 1

    recStart = 0

    if request("start") <> "" then
        recStart = request("start")
    end if

    recEnd = (recStart + recsPerPage) - 1
    
    if recEnd > ubound(arrRS,2) then
        recEnd = ubound(arrRS,2)
    end if
    
    totRecs = ubound(arrRS,2) + 1
    
    if (totRecs mod recsPerPage) = 0 then
        numPages = totRecs / recsPerPage
    else
        arrNumPages = split((totRecs / recsPerPage),".")
        numPages = cint(arrNumPages(0)) + 1
    end if
    
    pageNo = (recStart + recsPerPage) / recsPerPage
end if
rs.close
set rs = nothing
set cSql = nothing

' === FORM SUBMIT LEDGER PATCH ===
' Detect whether the durable ledger schema is present so the admin grid can
' surface durable state / emails / review columns. Degrades silently when the
' migration (db/form_submit_ledger.sql) has not been applied yet.
dim ledgerEnabled, cLedChk, rsLedChk
dim arrLedger, idxLed, rsLed
dim ledState, ledEmails, ledEmailsCnt, ledFiles, ledFilesCnt, ledRecords, ledPay, ledReleased
ledgerEnabled = false
on error resume next
set cLedChk = new cRunSql
cLedChk.Conn = connOracle
cLedChk.SqlStr = "select count(*) c from all_tables where table_name = 'FORM_SUBMIT_RECORD'"
set rsLedChk = cLedChk.Execute
if err.number = 0 then
    if not rsLedChk.eof then
        if isNumeric(rsLedChk("c") & "") then
            if clng(rsLedChk("c")) > 0 then ledgerEnabled = true
        end if
    end if
end if
err.clear
on error goto 0
set rsLedChk = nothing
set cLedChk = nothing
' === END FORM SUBMIT LEDGER PATCH ===

dim colspan
colspan = 9
if appRenewCode = "ALL" then
    colspan = colspan + 1
end if
if appRenewStatus = "ALL" then
    colspan = colspan + 1
end if
if Application("DEV_DOMAIN") = request.ServerVariables("SERVER_NAME") or officeIP then
    colspan = colspan + 2
end if
' === FORM SUBMIT LEDGER PATCH ===
if ledgerEnabled then
    colspan = colspan + 3
end if
' === END FORM SUBMIT LEDGER PATCH ===
%>
<form name="form" id="form" action="<%=scriptName%>" method="post">
<input type="hidden" name="action" id="action" value="" />
<table width="100%" border="0" cellpadding="2" cellspacing="0">
<tr>
    <td></td>
    <td align="center">
        <%if Application("MASS_APPROVE_FORMS") = "1" and appRenewStatus = "PENDING" then%>
        <input type="button" value="Approve" class="adminButton" onClick="doApprove() "/>
        <%end if%>
        <%
        if (appRenewStatus = "PENDING" and not officeIP) or appRenewStatus = "COMPLETED" or appRenewStatus = "APPROVED" or appRenewStatus = "DELETED" then
        else
        %>
        <input type="button" value="Delete" class="adminButton" onClick="doDelete() "/>
        <%end if%>
    </td>
	<td colspan="<%=colspan%>" align="right" class="adminText">
		<table border="0" cellpadding="2" cellspacing="0">
			<tr class="adminText">
				<td align="right">
					Search:
				</td>
				<td>
                    <input class="adminText" name="search" id="search" value="<%=searchVal%>" /><input class="adminButton" type="button" value="Search" onClick="doSearch();"/><%if officeIP then%> <input class="adminButton" type="button" value="Search Logs" onClick="doSearchLogs();"/><%end if%>
                </td>
            </tr>
			<tr class="adminText">
				<td align="right">
					Form Type:
				</td>
				<td>
					<select name="app_renew_cd" class="adminText" onChange="document.form.submit()">
                        <option value="ALL">All</option>
<%do while not rsAppRenewCodes.eof%>
						<option value="<%=rsAppRenewCodes("code_value")%>"<%if appRenewCode = rsAppRenewCodes("code_value") then%> selected<%end if%>><%=rsAppRenewCodes("code_value_desc")%><%if request.ServerVariables("SERVER_NAME") = Application("DEV_DOMAIN") then%> - <%=rsAppRenewCodes("code_value")%><%end if%></option>
	<%rsAppRenewCodes.MoveNext%>
<%loop%>
					</select>
				</td>
			</tr>
			<tr class="adminText">
				<td align="right">
					Status:
				</td>
				<td>
					<select name="app_renew_status" class="adminText" onChange="document.form.submit()">
                        <option value="ALL">All</option>
					    <option value="PENDING"<%if appRenewStatus = "PENDING" then%> selected<%end if%>>Pending</option>
					    <option value="APPROVED"<%if appRenewStatus = "APPROVED" then%> selected<%end if%>>Approved</option>
					    <option value="DENIED"<%if appRenewStatus = "DENIED" then%> selected<%end if%>>Denied</option>
					    <option value="PROGRESS"<%if appRenewStatus = "PROGRESS" then%> selected<%end if%>>In Progress</option>
					    <option value="SAVED"<%if appRenewStatus = "SAVED" then%> selected<%end if%>>Saved</option>
					    <option value="PAY_LATER"<%if appRenewStatus = "PAY_LATER" then%> selected<%end if%>>Pay Later</option>
					    <option value="COMPLETED"<%if appRenewStatus = "COMPLETED" then%> selected<%end if%>>Completed</option>
					    <option value="INVITE"<%if appRenewStatus = "INVITE" then%> selected<%end if%>>Invited</option>
					    <option value="DELETED"<%if appRenewStatus = "DELETED" then%> selected<%end if%>>Deleted</option>
					</select>
				</td>
			</tr>
            <%if appRenewStatus = "PROGRESS" or appRenewStatus = "ALL" then%>
			<tr class="adminText">
				<td align="right">
					Hide Bot Forms:
				</td>
				<td>
                    <input type="radio" name="hide_bot" id="hide_bot_y" value="Y" <%if  hideBot = "Y" then%>checked <%end if%> onclick="document.form.submit()" /> <label for="hide_bot_y">Yes</label>
                    <input type="radio" name="hide_bot" id="hide_bot_n" value="N" <%if  hideBot = "N" then%>checked <%end if%> onclick="document.form.submit()" /> <label for="hide_bot_n">No</label>
                </td>
			</tr>
            <%end if%>
		</table>
	</td>
</tr>
<tr class="adminTitlebar">
	<td>
	&nbsp;
	</td>
	<td align="center">
    <%if showDeleteViews then%>
	<input type="checkbox" name="deleteViews" value="all" onClick="toggleAll()">
    <%end if%>
	</td>
	<td>
	&nbsp;
	</td>
    <%if appRenewCode = "ALL" then%>
	<td>
	Form Type
	</td>
    <%end if%>
    <%if appRenewStatus = "ALL" then%>
	<td>
	Status
	</td>
    <%end if%>
	<td>
        <a href="<%=request.servervariables("SCRIPT_NAME")%>?sort=name&app_renew_cd=<%=appRenewCode%>&app_renew_status=<%=appRenewStatus%><%if request("action") = "doSearch" then%>&action=doSearch&search=<%=searchVal%><%end if%>">
	        <img src="/admin/admin/grafix/sort_up.gif" title="Sort by Name" /> 
        </a>
        Name
        <a href="<%=request.servervariables("SCRIPT_NAME")%>?sort=name&order=desc&app_renew_cd=<%=appRenewCode%>&app_renew_status=<%=appRenewStatus%><%if request("action") = "doSearch" then%>&action=doSearch&search=<%=searchVal%><%end if%>">
	        <img src="/admin/admin/grafix/sort_down.gif" title="Sort by Name in descending order" /> 
        </a>
	</td>
    <%if Application("DEV_DOMAIN") = request.ServerVariables("SERVER_NAME") or officeIP then%>
	<td>
        MEMBERSHIP_ID
	</td>
    <%end if%>
	<td>
	PDF
	</td>
	<td>
	Awaiting Updates
	</td>
	<td>
	Payment
	</td>
	<td>
	Payment Date-Time
	</td>
	<td>
	Expiration Date
	</td>
	<td>
        <a href="<%=request.servervariables("SCRIPT_NAME")%>?sort=create_dt&app_renew_cd=<%=appRenewCode%>&app_renew_status=<%=appRenewStatus%><%if request("action") = "doSearch" then%>&action=doSearch&search=<%=searchVal%><%end if%>">
	        <img src="/admin/admin/grafix/sort_up.gif" title="Sort by Created" /> 
        </a>
	Created
        <a href="<%=request.servervariables("SCRIPT_NAME")%>?sort=create_dt&order=desc&app_renew_cd=<%=appRenewCode%>&app_renew_status=<%=appRenewStatus%><%if request("action") = "doSearch" then%>&action=doSearch&search=<%=searchVal%><%end if%>">
	        <img src="/admin/admin/grafix/sort_down.gif" title="Sort by Created in descending order" /> 
        </a>
	</td>
	<td>
        <a href="<%=request.servervariables("SCRIPT_NAME")%>?sort=modify_dt&app_renew_cd=<%=appRenewCode%>&app_renew_status=<%=appRenewStatus%><%if request("action") = "doSearch" then%>&action=doSearch&search=<%=searchVal%><%end if%>">
	        <img src="/admin/admin/grafix/sort_up.gif" title="Sort by Last Modified" /> 
        </a>
	Last Modified
        <a href="<%=request.servervariables("SCRIPT_NAME")%>?sort=modify_dt&order=desc&app_renew_cd=<%=appRenewCode%>&app_renew_status=<%=appRenewStatus%><%if request("action") = "doSearch" then%>&action=doSearch&search=<%=searchVal%><%end if%>">
	        <img src="/admin/admin/grafix/sort_down.gif" title="Sort by Last Modified in descending order" /> 
        </a>
	</td>
    <%if Application("DEV_DOMAIN") = request.ServerVariables("SERVER_NAME") or officeIP then%>
	<td>
        <a href="<%=request.servervariables("SCRIPT_NAME")%>?sort=submit_dt&app_renew_cd=<%=appRenewCode%>&app_renew_status=<%=appRenewStatus%><%if request("action") = "doSearch" then%>&action=doSearch&search=<%=searchVal%><%end if%>">
	        <img src="/admin/admin/grafix/sort_up.gif" title="Sort by Last Submitted" /> 
        </a>
	Last Submitted
        <a href="<%=request.servervariables("SCRIPT_NAME")%>?sort=submit_dt&order=desc&app_renew_cd=<%=appRenewCode%>&app_renew_status=<%=appRenewStatus%><%if request("action") = "doSearch" then%>&action=doSearch&search=<%=searchVal%><%end if%>">
	        <img src="/admin/admin/grafix/sort_down.gif" title="Sort by Last Submitted in descending order" /> 
        </a>
	</td>
    <%end if%>
	<%if ledgerEnabled then%>
	<td>
	Submission State
	</td>
	<td>
	Emails
	</td>
	<td>
	Review
	</td>
	<%end if%>
</tr>
<%if not isArray(arrRS) then%>
<tr class="adminText">
	<td colspan="<%=colspan+2%>" align="center">
		There are no submitted forms in this view.
	</td>
</tr>
<%
else
    for i = recStart to recEnd
        if arrRS(0,i)&"" <> "0" then
            if isArray(arrMem) then
                idxMem = ubound(arrMem) + 1
                redim preserve arrMem(idxMem)
            else
                idxMem = 0
                redim arrMem(idxMem)
            end if
            arrMem(idxMem) = arrRS(0,i)
        end if
        if arrRS(11,i)&"" <> "0" then
            if isArray(arrChi) then
                idxMem = ubound(arrChi) + 1
                redim preserve arrChi(idxMem)
            else
                idxMem = 0
                redim arrChi(idxMem)
            end if
            arrChi(idxMem) = arrRS(11,i)
        end if
        if arrRS(12,i)&"" <> "0" then
            if isArray(arrPar) then
                idxMem = ubound(arrPar) + 1
                redim preserve arrPar(idxMem)
            else
                idxMem = 0
                redim arrPar(idxMem)
            end if
            arrPar(idxMem) = arrRS(12,i)
        end if
    next

    set cSql = new cRunSql
    cSql.Conn = connOracle
    cSql.SqlStr = "select app_renew_id from mem_await where approval_dt is null and app_renew_id in ("
    for i = recStart to recEnd
        cSql.AddSqlStr("?")
        if i < recEnd then
            cSql.AddSqlStr(",")
        end if
    next
    cSql.AddSqlStr(") union select app_renew_id from mem_det_await where approval_dt is null and app_renew_id in (")
    for i = recStart to recEnd
        cSql.AddSqlStr("?")
        if i < recEnd then
            cSql.AddSqlStr(",")
        end if
    next
    cSql.AddSqlStr(")")
    for i = recStart to recEnd
        cSql.AddParam(arrRS(1,i))
    next
    for i = recStart to recEnd
        cSql.AddParam(arrRS(1,i))
    next
    set rsAU = cSql.Execute
    set cSql = nothing
    if rsAU.eof then
    else
        arrAU = rsAU.GetRows
    end if
    rsAU.close
    set rsAU = nothing

    set cSql = new cRunSql
    cSql.Conn = connOracle
    cSql.SqlStr = "select app_renew_id from mem_await where app_renew_id in ("
    for i = recStart to recEnd
        cSql.AddSqlStr("?")
        if i < recEnd then
            cSql.AddSqlStr(",")
        end if
    next
    cSql.AddSqlStr(") union select app_renew_id from mem_det_await where app_renew_id in (")
    for i = recStart to recEnd
        cSql.AddSqlStr("?")
        if i < recEnd then
            cSql.AddSqlStr(",")
        end if
    next
    cSql.AddSqlStr(")")
    for i = recStart to recEnd
        cSql.AddParam(arrRS(1,i))
    next
    for i = recStart to recEnd
        cSql.AddParam(arrRS(1,i))
    next
    set rsAU = cSql.Execute
    set cSql = nothing
    if rsAU.eof then
    else
        arrAwait = rsAU.GetRows
    end if
    rsAU.close
    set rsAU = nothing

    set cMem = new cMemberValuesO
    cMem.AddLoadCol("LIC_EXPIRE_DT")
    cMem.ValuesOnly = true
    if isArray(arrMem) then
        cMem.LoadFromMembershipID arrMem
    end if

    set cChi = new cMemberValuesO
    cChi.AddLoadCol("LIC_EXPIRE_DT")
    cChi.ValuesOnly = true
    if isArray(arrChi) then
        cChi.LoadFromMembershipID arrChi
    end if

    set cPar = new cMemberValuesO
    cPar.AddLoadCol("LIC_EXPIRE_DT")
    cPar.ValuesOnly = true
    if isArray(arrPar) then
        cPar.LoadFromMembershipID arrPar
    end if

    Set connG = Server.CreateObject("ADODB.Connection")
    connG.Open connectionStringG
    set cSql = new cRunSql
    cSql.Conn = connG
    cSql.SqlStr = "select form_id,status from form_pdf where client_id = ? and form_id in ("
    cSql.AddParam(Application("CLIENT_ID"))
    for i = recStart to recEnd
        cSql.AddSqlStr("?")
        if i < recend then
            cSql.AddSqlStr(",")
        end if
    next
    cSql.AddSqlStr(")")
    for i = recStart to recEnd
        cSql.AddParam(arrRS(1,i))
    next
    set rsPDF = cSql.Execute
    set cSql = nothing
    if rsPDF.eof then
    else
        arrPDF = rsPDF.GetRows()
    end if
    rsPDF.Close
    set rsPDF = nothing
    connG.close
    set connG = nothing

    ' === FORM SUBMIT LEDGER PATCH ===
    ' Batch-load durable ledger state for the app_renew_ids shown on this page.
    if ledgerEnabled then
        set cSql = new cRunSql
        cSql.Conn = connOracle
        cSql.SqlStr = "select r.app_renew_id, r.durable_state" &_
            ", (select o.state from form_submit_operation o where o.app_renew_id = r.app_renew_id and o.operation_key = 'emails') emails_state" &_
            ", (select o.done_count || '/' || o.expected_count from form_submit_operation o where o.app_renew_id = r.app_renew_id and o.operation_key = 'emails') emails_count" &_
            ", (select o.state from form_submit_operation o where o.app_renew_id = r.app_renew_id and o.operation_key = 'files') files_state" &_
            ", (select o.done_count || '/' || o.expected_count from form_submit_operation o where o.app_renew_id = r.app_renew_id and o.operation_key = 'files') files_count" &_
            ", (select o.state from form_submit_operation o where o.app_renew_id = r.app_renew_id and o.operation_key = 'records') records_state" &_
            ", (select o.state from form_submit_operation o where o.app_renew_id = r.app_renew_id and o.operation_key = 'payment') payment_state" &_
            ", (select count(*) from form_submit_event e where e.app_renew_id = r.app_renew_id and e.event_type = 'awaiting_updates_released') released_cnt" &_
            " from form_submit_record r where r.app_renew_id in ("
        for i = recStart to recEnd
            cSql.AddSqlStr("?")
            if i < recEnd then
                cSql.AddSqlStr(",")
            end if
        next
        cSql.AddSqlStr(")")
        for i = recStart to recEnd
            cSql.AddParam(arrRS(1,i))
        next
        set rsLed = nothing
        on error resume next
        set rsLed = cSql.Execute
        if err.number = 0 then
            if not rsLed.eof then
                arrLedger = rsLed.GetRows
            end if
        end if
        err.clear
        on error goto 0
        if not (rsLed is nothing) then
            on error resume next
            rsLed.close
            on error goto 0
        end if
        set rsLed = nothing
        set cSql = nothing
    end if
    ' === END FORM SUBMIT LEDGER PATCH ===

    switch = 0
    for i = recStart to recEnd
	    switch = 1 - switch
	    if switch > 0 then
		    lineClass = "activeLineItem1"
	    else
		    lineClass = "activeLineItem2"
	    end if
%>
<tr class="<%=lineClass%>">
	<td align="center" class="adminText">
		<%=i+1%>.
	</td>
	<td align="center" class="adminText">
        <%
        hasAU = false
        if isArray(arrAU) then
            for idxAU = lbound(arrAU,2) to ubound(arrAU,2)
                if arrAU(0,idxAU)&"" = arrRS(1,i)&"" then
                    hasAU = true
                    exit for
                end if
            next
        end if
        if hasAU then
        elseif (arrRS(6,i) = "PENDING" and not officeIP) or arrRS(6,i) = "COMPLETED" or arrRS(6,i) = "APPROVED" or arrRS(6,i) = "DELETED" then
        else
         %>
		<input type="checkbox" name="deleteView" value="<%=arrRS(1,i)%>" />
        <%
        end if 
        %>
	</td>
	<td align="left" class="adminText">
		<a href="<%=scriptName%>?action=modify&app_renew_id=<%=arrRS(1,i)%>&start=<%=recStart%><%if request("action") = "doSearch" then%>&search=<%=searchVal%><%end if%>" class="adminLink">Edit <%=arrRS(1,i)%></a>
	</td>
    <%if appRenewCode = "ALL" then%>
	<td>
	    <%=arrRS(3,i)%>
	</td>
    <%end if%>
    <%if appRenewStatus = "ALL" then%>
	<td>
	    <%if arrRS(6,i) = "PENDING" then%>Pending<%elseif arrRS(6,i) = "INVITE" then%>Invited<%elseif arrRS(6,i) = "COMPLETED" then%>Completed<%elseif arrRS(6,i) = "DENIED" then%>Denied<%elseif arrRS(6,i) = "SAVED" then%>Saved<%elseif arrRS(6,i) = "PROGRESS" then%>In Progress<%elseif arrRS(6,i) = "PAY_LATER" then%>Pay Later<%elseif arrRS(6,i) = "DELETED" then%>Deleted<%else%>Approved<%end if%>
	</td>
    <%end if%>
	<td align="left" class="adminText">
		<%=arrRS(2,i)%>
	</td>
    <%if Application("DEV_DOMAIN") = request.ServerVariables("SERVER_NAME") or officeIP then%>
	<td align="left" class="adminText">
		<%=arrRS(0,i)%>
        <%
        if arrRS(8,i) = "WHOLE_RENEW" then
            groupPayID = ""
            set cSql = new cRunSql
            cSql.Conn = connoracle
            cSql.SqlStr = "select data from app_renew_detail where app_renew_id = ? and lower(field_name) = 'group_pay_id'"
            cSql.AddParam(arrRS(1,i))
            set rs = cSql.Execute
            set cSql = nothing
            if rs.eof then
            else
                groupPayID = rs("data")
            end if
            rs.close
            set rs = nothing
            if groupPayID&"" <> "" then
                set cSql = new cRunSql
                cSql.Conn = connoracle
                cSql.SqlStr = "select a.membership_id from pay_later p, app_renew a where a.app_renew_id = p.form_id and p.group_pay_id = ? and membership_id <> ?"
                cSql.AddParam(groupPayID)
                cSql.AddParam(arrRS(1,i))
                set rs = cSql.Execute
                set cSql = nothing
                do while not rs.eof
                    response.write ", " & rs("membership_id")
                    rs.MoveNext
                loop
                rs.close
                set rs = nothing
            end if
        end if
        %>
	</td>
    <%end if%>
	<td>
        <%
        if isArray(arrPDF) then
            for idxPDF = lbound(arrPDF,2) to ubound(arrPDF,2)
                if arrPDF(0,idxPDF)&"" = arrRS(1,i)&"" then
                    response.write arrPDF(1,idxPDF)
                    exit for
                end if
            next
        end if
        %>
	</td>
	<td align="left" class="adminText">
		<%
        hasAU = false
        if isArray(arrAwait) then
            for idxAU = lbound(arrAwait,2) to ubound(arrAwait,2)
                if arrAwait(0,idxAU)&"" = arrRS(1,i)&"" then
                    hasAU = true
                    exit for
                end if
            next
        end if
        if hasAU then
            response.write "Yes"
        else
            response.write "No"
        end if
        %>
	</td>
	<td align="left" class="adminText">
		<%
        if arrRS(9,i)&"" <> "" then
            if cint(arrRS(9,i)) = "0" then
                response.write "NA"
            elseif arrRS(13,i)&"" = "" then
            else
                response.write arrRS(13,i) & " (" & formatCurrency(arrRS(9,i)) & ")"
            end if
        end if
        %>
	</td>
	<td align="left" class="adminText">
        <%=arrRS(14,i)%>
	</td>
	<td align="left" class="adminText">
        <%
        if arrRS(0,i)&"" <> "0" then
            cMem.MembershipID = arrRS(0,i)
            if cMem.Value("LIC_EXPIRE_DT")&"" <> "" then
                response.write cMem.Value("LIC_EXPIRE_DT")
            elseif arrRS(11,i)&"" <> "0" then
                cChi.MembershipID = arrRS(11,i)
                if cChi.Value("LIC_EXPIRE_DT")&"" <> "" then
                    response.write cChi.Value("LIC_EXPIRE_DT")
                end if
            elseif arrRS(12,i)&"" <> "0" then
                cPar.MembershipID = arrRS(12,i)
                if cPar.Value("LIC_EXPIRE_DT")&"" <> "" then
                    response.write cPar.Value("LIC_EXPIRE_DT")
                end if
            end if
        end if
        %>
	</td>
    <td align="left" class="adminText">
		<%=timeZoneAdjust(arrRS(4,i))%>
	</td>
	<td align="left" class="adminText">
		<%=timeZoneAdjust(arrRS(5,i))%>
	</td>
    <%if Application("DEV_DOMAIN") = request.ServerVariables("SERVER_NAME") or officeIP then%>
	<td align="left" class="adminText">
		<%=timeZoneAdjust(arrRS(7,i))%>
	</td>
    <%end if%>
	<%if ledgerEnabled then%>
	<td align="left" class="adminText">
		<%
		ledState = "" : ledEmails = "" : ledEmailsCnt = "" : ledFiles = "" : ledFilesCnt = "" : ledRecords = "" : ledPay = "" : ledReleased = 0
		if isArray(arrLedger) then
			for idxLed = lbound(arrLedger,2) to ubound(arrLedger,2)
				if arrLedger(0,idxLed)&"" = arrRS(1,i)&"" then
					ledState = arrLedger(1,idxLed)&""
					ledEmails = arrLedger(2,idxLed)&""
					ledEmailsCnt = arrLedger(3,idxLed)&""
					ledFiles = arrLedger(4,idxLed)&""
					ledFilesCnt = arrLedger(5,idxLed)&""
					ledRecords = arrLedger(6,idxLed)&""
					ledPay = arrLedger(7,idxLed)&""
					ledReleased = arrLedger(8,idxLed)
					exit for
				end if
			next
		end if
		if ledState = "" then
			response.write "&mdash;"
		else
			response.write ledState
		end if
		%>
	</td>
	<td align="left" class="adminText">
		<%
		if ledEmails = "" or ledEmails = "not_configured" then
			response.write "NA"
		else
			response.write ledEmails
			if ledEmailsCnt&"" <> "" and ledEmailsCnt <> "/" then
				response.write " (" & ledEmailsCnt & ")"
			end if
		end if
		%>
	</td>
	<td align="left" class="adminText">
		<%
		if ledRecords = "awaiting" then
			response.write "Awaiting"
		elseif isNumeric(ledReleased&"") and ledReleased&"" <> "" then
			if clng(ledReleased) > 0 then
				response.write "Approved"
			else
				response.write "Auto"
			end if
		else
			response.write "Auto"
		end if
		%>
	</td>
	<%end if%>
</tr>	
<%
    next
    set cPar = nothing
    set cChi = nothing
    set cMem = nothing
end if
%>
</table>

<%if numPages > 1 then%>
<div align="center" class="adminTitlebar">
<%
    for i = 1 to numPages
        if i > 1 and i <= numPages and (((i-1) mod 10) <> 0) then
            response.Write "&nbsp;&#183;&nbsp;"
        end if
        if pageNo <> i then
            nextRecStart = (i-1) * recsPerPage
%>
    <a href="<%=request.servervariables("SCRIPT_NAME")%>?app_renew_status=<%=appRenewStatus%>&app_renew_cd=<%=appRenewCode%>&start=<%=nextRecStart%><%if request("action") = "doSearch" then%>&action=doSearch&search=<%=searchVal%><%end if%><%if request("sort")&"" <> "" then%>&sort=<%=request("sort")%><%if request("order")&"" <> "" then%>&order=<%=request("order")%><%end if%><%end if%>" class="adminNavLink">
    <%
        end if
        response.Write i
        if pageNo <> i then
    %>
    </a>
    <%
        end if
        if (i mod 10) = 0 then
            response.Write "<br>"
        end if
    next
%>
</div>
<%
end if
if totRecs > 0 then
%>
<div align="center" class="lineItemsTotal">
    <%=recStart+1%>-<%=recEnd+1%> of <%=totRecs%> values
</div>
<%end if%>

</form>
<%end sub%>

<%
sub modify()
dim rsAppRenew, appRenewType, ssn, rsDup, dupCheck, appRenewStatus, fein, existID, birthDt, facName, facPhone, zip, phone, existIndID, existLicID
dim dupSSN, dupBirth, dupSSNBirth, coName, dupCO, dupNameSSN, dupSSNBirthName, dupSSNType, dupNameSSNType, dupSSNName
dim dupSSNLic, dupNameSSNLic, dupSSNBirthNameLic
dim dupCredential, lastName, firstName, credDef, cContact, membershipID
dim licType, recType, rsMem, objFSO
dim TS, logTxt, logFile, arrPath, parentID, pdfFile, ajaxFile
%>

<script language="Javascript">
<!--
function dup() {
	if (confirm("Duplicate records exist. Do you want to proceed?"))
		document.location = "<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>";
}

function hideLink(lnk) {
    lnk.style.display = "none";
}

-->
</script>

<%
set cSql = new cRunSql
cSql.Conn = connOracle
cSql.SqlStr = "select APP_RENEW_ID,MEMBERSHIP_ID,APP_RENEW_TYPE,APP_RENEW_STATUS,CREATE_DT,NAME from app_Renew where app_renew_id = ?"
cSql.AddParam(request.querystring("app_renew_id"))
set rsAppRenew = cSql.Execute
if rsAppRenew.eof then
    %><div align="center">There is no form submission with ID <%=request.querystring("app_renew_id")%></div><%
    exit sub
end if
appRenewType = rsAppRenew("APP_RENEW_TYPE")
appRenewStatus = rsAppRenew("app_renew_status")

cSql.SqlStr = "select org_sub_ty_Cd from membership where membership_id = ?"
cSql.AddParam(rsAppREnew("MEMBERSHIP_ID"))
set rsMem = cSql.Execute
if not rsMem.eof then
    recType = rsMem("ORG_SUB_TY_CD")
end if
rsMem.Close
set rsMem = nothing

set cSql = nothing
%>
<table border="0" cellpadding="2" cellspacing="0" align="center" class="adminText">
	<tr><td class="adminSubTitle" align="center" colspan="3">Application</td></tr>
	
	
	<tr>
		<td align="right" colspan="3" class="adminText">
			<input type="button" value="List" name="list" class="adminButton" onMouseOver="this.className='adminButtonHover'" onMouseOut="this.className='adminButton'" onClick="parent.location='<%=request.servervariables("SCRIPT_NAME")%>?app_renew_cd=<%=appRenewType%>&app_renew_status=<%=appRenewStatus%>&start=<%=request("start")%><%if request("search") <> "" then%>&action=doSearch&search=<%=request("search")%><%end if%>'">
			<hr size="1" color="#000000">
		</td>
	</tr>	
	
<tr><td align="right">Created:</td><td>&nbsp;</td><td><%=timeZoneAdjust(rsAppRenew("CREATE_DT"))%></td></tr>
<tr><td align="right">Type:</td><td>&nbsp;</td><td><a href="/admin/admin/app_renew_config.asp?app_renew_cd=<%=rsAppRenew("APP_RENEW_TYPE")%>&do=edit" target="_blank"><%=rsAppRenew("APP_RENEW_TYPE")%></a></td></tr>
<tr><td align="right">Status:</td><td>&nbsp;</td><td><%if appRenewStatus = "PENDING" then%>Pending<%elseif appRenewStatus = "INVITE" then%>Invited<%elseif appRenewStatus = "COMPLETED" then%>Completed<%elseif appRenewStatus = "DENIED" then%>Denied<%elseif appRenewStatus = "SAVED" then%>Saved<%elseif appRenewStatus = "PROGRESS" then%>In Progress<%elseif appRenewStatus = "PAY_LATER" then%>Pay Later<%else%>Approved<%end if%></td></tr>
<tr><td align="right">Internal ID:</td><td>&nbsp;</td><td><%=rsAppRenew("MEMBERSHIP_ID")%></td></tr>
<tr><td align="right">Database Record:</td><td>&nbsp;</td><td><a href="https://<%=request.ServerVariables("SERVER_NAME")%>/db/live/universal.asp?PK=<%=rsAppRenew("MEMBERSHIP_ID")%>&VIEW=MASTERVIEW&ORG_SUB_TY_CD=<%=recType%>" target="_blank">Link</a>
<%if getFormVal("EXISTING_RECORD_ID")&"" = "1" then%><strong><font color="red">This application will update an existing record.</font></strong><%end if%>
</td></tr>
<tr><td align="right">Name:</td><td>&nbsp;</td><td><%=rsAppRenew("NAME")%></td></tr>
<%if rsAppRenew("APP_RENEW_TYPE") = "WHOLE_RENEW" then%>
<tr><td align="right">Group Payment ID:</td><td>&nbsp;</td><td><%=getFormVal("group_pay_id")%></td></tr>
<tr><td align="right">License Numbers:</td><td>&nbsp;</td><td><%
    dim rsPay, rsForm

	set cSql = new cRunSql
	cSql.conn = connOracle
	cSql.SqlStr = "select FORM_ID from pay_later where group_pay_id = ?"
	cSql.AddParam(getFormVal("group_pay_id"))
	set rsPay = cSql.Execute
	set cSql = nothing

    set cMember = new cMemberValuesO
    cMember.AddLoadCol("Wholesale_License")
    cMember.ValuesOnly = true

    cMember.LoadFromMembershipId rsAppRenew("MEMBERSHIP_ID")
    response.write cMember.Value("Wholesale_License") & ", "
    do while not rsPay.eof
		set cSql = new cRunSql
		cSql.conn = connOracle
		cSql.SqlStr = "select membership_id from app_renew where app_renew_id = ?"
		cSql.AddParam(rsPay("FORM_ID"))
		set rsForm = cSql.Execute
		set cSql = nothing
		if rsForm.eof then
		else
			cMember.LoadFromMembershipID rsForm("MEMBERSHIP_ID")
            response.write cMember.Value("Wholesale_License")
        end if
        rsPay.MoveNext
        if rsPay.eof then
        else
            response.write ", "
        end if
    loop
    set cMember = nothing

    set rsPay = nothing
%>
</td></tr>
<%end if%>
<tr><td align="right" valign="top">Files:</td><td>&nbsp;</td><td>
<table>
<%
set cMember = new cMemberValuesO
cMember.AddLoadCol("ORG_SUB_TY_CD")
cMember.ValuesOnly = true
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
dbConnect connb,Application("ConnectionStringB")
set cSql = new cRunSql
cSql.Conn = connb
cSql.DisplayErrors = true
cSql.SqlStr = "select membership_id,title,resource_id,storage_key from document where app_renew_id = ? order by create_dt desc"
cSql.AddParam(rsAppRenew("APP_RENEW_ID"))
set cDoc = cSql.Execute
do while not cDoc.eof
    response.write "<tr>"
    cMember.LoadFromMembershipId cDoc("membership_id")
    dbLink = "/db/live/universal.asp?PK=" & cDoc("membership_id") & "&ORG_SUB_TY_CD=" & cMember.Value("ORG_SUB_TY_CD") & "&VIEW=MASTERVIEW"
    if cDoc("storage_key")&"" <> "" then
%>
<td><a href="downloadresource.asp?id=<%=cDoc("resource_id")%>" class="adminLink"><%=cDoc("Title")%></a></td><td>&nbsp;-&nbsp;</td><td><a href="<%=dblink%>" target="_blank" class="adminLink">Database Record for File</a></td>
<%
    else
%>
        <td><%=cDoc("Title")%> not on disk</td><td>&nbsp;-&nbsp;</td><td><a href="<%=dblink%>" target="_blank" class="adminLink">Database Record for File</a></td>
<%
    end if
    response.write "</tr>"
	cDoc.MoveNext
loop
cDoc.close
set cDoc  = nothing
set cSql = nothing
dbDisconnect connB
set objFSO = nothing
set cMember = nothing
%>
</table>
</td></tr>
<%if Application("DEV_DOMAIN") = request.ServerVariables("SERVER_NAME") or officeIP then%>
<tr><td align="right" valign="top">Debug Links:</td><td>&nbsp;</td><td>
<ul>
<%
    set cSql = new cRunSql
    cSql.Conn = connOracle
    cSql.SqlStr = "select code_value_long_desc from code_value where code_class = 'APP_RENEW_TYPE_CD' and code_value = ?"
    cSql.AddParam(rsAppRenew("APP_RENEW_TYPE"))
    set rsConfig = cSql.Execute
    set cSql = nothing

    set cSql = new cRunSql
    cSql.Conn = connOracle
    cSql.SqlStr = "select page from app_renew_pages where app_renew_type = ? order by order_by"
    cSql.AddParam(rsAppRenew("APP_RENEW_TYPE"))
    set rsPages = cSql.Execute
    do while not rsPages.eof
%>
<li><a href="<%=rsConfig("code_value_long_desc")%>/<%=rsPages("page")%>?action=loadDebug&formID=<%=rsAppRenew("APP_RENEW_ID")%>" class="adminLink" target="_blank"><%=rsPages("page")%></a></li>
<%
        rsPages.MoveNext
    loop
    rsPages.close
    set rsPages = nothing
    set cSql = nothing
    rsConfig.close
    set rsConfig = nothing
%>
</ul>
</td></tr>
<tr><td align="right" valign="top">Restore Links:</td><td>&nbsp;</td><td>
<ul>
<%
    set cSql = new cRunSql
    cSql.Conn = connOracle
    cSql.SqlStr = "select code_value_long_desc from code_value where code_class = 'APP_RENEW_TYPE_CD' and code_value = ?"
    cSql.AddParam(rsAppRenew("APP_RENEW_TYPE"))
    set rsConfig = cSql.Execute
    set cSql = nothing

    set cSql = new cRunSql
    cSql.Conn = connOracle
    cSql.SqlStr = "select page from app_renew_pages where app_renew_type = ? and order_by is null"
    cSql.AddParam(rsAppRenew("APP_RENEW_TYPE"))
    set rsPages = cSql.Execute
    do while not rsPages.eof
%>
<li><a href="<%=rsConfig("code_value_long_desc")%>/<%=rsPages("page")%>?action=restore&formID=<%=rsAppRenew("APP_RENEW_ID")%>" class="adminLink" target="_blank"><%=rsPages("page")%></a></li>
<%
        rsPages.MoveNext
    loop
    rsPages.close
    set rsPages = nothing
    set cSql = nothing
    rsConfig.close
    set rsConfig = nothing
end if
%>
</ul>
</td></tr>
<%
    dim rsEP,rsP
%>
<tr><td align="right" valign="top">PDFs:</td><td>&nbsp;</td><td>
<ul>
<%
    Set connG = Server.CreateObject("ADODB.Connection")
    connG.Open connectionStringG
    set cSql = new cRunSql
    cSql.Conn = connG
    cSql.SqlStr = "select status,error,form_id,start_dt,complete_dt,create_dt from form_pdf where client_id = ? and form_id = ? order by form_pdf_id"
    cSql.AddParam(Application("CLIENT_ID"))
    cSql.AddParam(rsAppRenew("APP_RENEW_ID"))
    set rsPDF = cSql.Execute
    set cSql = nothing
    do while not rsPDF.eof
%>
<li>
    <%=rsPDF("status")%>
    <%if rsPDF("status") = "ERROR" then%> - <%=rsPDF("error")%><%end if%>
    <%if rsPDF("status") = "PENDING" then%> Created <%=rsPDF("create_dt")%><%end if%>
    <%if rsPDF("status") = "PROCESSING" then%> Started <%=rsPDF("start_dt")%><%end if%>
    <%if rsPDF("status") = "COMPLETE" then%> Started <%=rsPDF("start_dt")%> Finished <%=rsPDF("complete_dt")%> Duration <%=dateDiff("s",cDate(rsPDF("start_dt")),cDate(rsPDF("complete_dt")))%> seconds<%end if%>
</li>
<%
        rsPDF.MoveNext
    loop
    rsPDF.close
    set rsPDF = nothing
    connG.Close
    set connG = nothing
%>
</ul>
</td></tr>
<tr><td align="right" valign="top">Email Pages:</td><td>&nbsp;</td><td>
<ul>
<%
    Set connG = Server.CreateObject("ADODB.Connection")
    connG.Open connectionStringG
    set cSql = new cRunSql
    cSql.Conn = connG
    cSql.SqlStr = "select page_id,status,error,form_id,start_dt,complete_dt,create_dt from form_email_page where client_id = ? and form_id = ? order by form_email_page_id"
    cSql.AddParam(Application("CLIENT_ID"))
    cSql.AddParam(rsAppRenew("APP_RENEW_ID"))
    set rsEP = cSql.Execute
    set cSql = nothing
    do while not rsEP.eof
        set cSql = new cRunSql
        cSql.Conn = connOracle
        cSql.DisplayErrors = true
        cSql.SqlStr = "select page from app_renew_pages p, app_renew a where p.app_renew_type = a.app_renew_type and a.app_renew_id = ? and p.app_renew_pages_id = ?"
        cSql.AddParam(rsEP("form_id"))
        cSql.AddParam(rsEP("page_id"))
        set rsP = cSql.Execute
        set cSql = nothing
%>
<li>
    <%=rsP("page")%> -
    <%=rsEP("status")%>
    <%if rsEP("status") = "ERROR" then%> - <%=rsEP("error")%><%end if%>
    <%if rsEP("status") = "PENDING" then%> Created <%=rsEP("create_dt")%><%end if%>
    <%if rsEP("status") = "PROCESSING" then%> Started <%=rsEP("start_dt")%><%end if%>
    <%if rsEP("status") = "SENT" then%> Started <%=rsEP("start_dt")%> Finished <%=rsEP("complete_dt")%> Duration <%=dateDiff("s",cDate(rsEP("start_dt")),cDate(rsEP("complete_dt")))%> seconds<%end if%>
</li>
<%
        rsP.close
        set rsP = nothing
        rsEP.MoveNext
    loop
    rsEP.close
    set rsEP = nothing
    connG.Close
    set connG = nothing
%>
</ul>
</td></tr>
</table>
<br>
<%
membershipId = rsAppRenew("MEMBERSHIP_ID") 
set cMember = new cMemberValuesO
cMember.LoadFromMembershipID  membershipId
parentID = cMember.Value("ORG_RELATE_ID")
set rsAppRenew = nothing


set cSql = new cRunSql
cSql.Conn = connOracle
cSql.DisplayErrors = true
cSql.SqlStr = "select app_renew_fields.column_name, app_renew_detail.data " &_
	"from app_renew_detail, app_renew_fields, app_renew " &_
	"where upper(app_renew_fields.field_name) = upper(app_renew_detail.field_name) " &_
	"and app_renew.app_renew_id = app_renew_detail.app_renew_id " &_
	"and app_renew_fields.app_renew_type = app_renew.app_renew_type " &_
	"and app_renew_fields.column_name is not null " &_
	"and app_renew_fields.update_record = 'Y' " &_
	"and app_renew_detail.app_renew_id = ?"
cSql.AddParam(request.querystring("app_renew_id"))
set rsAppRenewDetail = cSql.Execute

if rsAppRenewDetail.eof and rsAppRenewDetail.bof then
else
	arrAppRenewDetail = rsAppRenewDetail.GetRows
end if

dupCheck = false

'sdhls
if Application("ORG_ID") = "3900" then
    phone = getFormVal("phone")
    company = getFormVal("company")
    if right(appRenewType,4) = "_APP" then
        if phone&"" <> "" and company&"" <> "" then
            dim objStr
            set cCV = new cCodeValueO
            cCV.LoadFromCodeClassAndValue "APP_RENEW_TYPE_CD",appRenewType
            objStr = cCV.OptionValue1
            if objStr&"" <> "" then
                objStr = "'" & replace(objStr,", ","','") & "'"
            end if
            set cSql = nothing
            
            set cSql = new cRunSql
            cSql.Conn = connOracle
            cSql.sqlStr = "select membership_id  " &_
                "from membership " &_
                "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'EMP' " &_
                "and phone_day = ? " &_
                "and upper(company) = upper(?) " &_
                "and membership_id <> ? "
            if objStr&"" <> "" then
                cSql.AddSqlStr("and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = membership.membership_id and column_desc = 'License_Obj_Type_EMP'),'') in (" & objStr & ")")
            end if

            cSql.AddParam(phone)
            cSql.AddParam(company)
            cSql.AddParam(membershipId)
            'cSql.Debug
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                existID = rsDup("membership_id")
            end if
            rsDup.close
            set rsDup = nothing
        end if
    end if
end if
'end sdhls

'ndbvme
if Application("ORG_ID") = "4100" then
    ssn = getFormVal("ssn")
    if appRenewType = "APP_VET" then
        if ssn&"" <> "" then
            cSql.sqlStr = "select membership_id  " &_
                "from membership " &_
                "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MEMBER' " &_
                "and replace(soc_security,'-','') = replace(?,'-','') " &_
                "and membership_id <> ?"
            cSql.AddParam(ssn)
            cSql.AddParam(membershipId)
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                existIndID = rsDup("membership_id")
            end if
            rsDup.close
            set rsDup = nothing
        end if
    end if
    if appRenewType = "APP_TECH" then
        if ssn&"" <> "" then
            cSql.sqlStr = "select membership_id  " &_
                "from membership " &_
                "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'VETTECH' " &_
                "and replace(soc_security,'-','') = replace(?,'-','') " &_
                "and membership_id <> ?"
            cSql.AddParam(ssn)
            cSql.AddParam(membershipId)
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                existIndID = rsDup("membership_id")
            end if
            rsDup.close
            set rsDup = nothing
        end if
    end if
end if
'end ndbvme

'nsbla
if Application("ORG_ID") = "3200" then
    licNo = getFormVal("nvcertificate")
    firstName = getFormVal("FIRST_NAME")
    lastName = getFormVal("LAST_NAME")
    if appRenewType = "APP" or appRenewType = "APP_RECIP" or appRenewType = "INTERN_APP" then
        if licNo&"" <> "" and firstName&"" <> "" and lastName&"" <> "" then
            cSql.sqlStr = "select membership_id  " &_
                "from membership " &_
                "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MEMBER' " &_
                "and upper(first_name) = upper(?) and (upper(last_name) = upper(?) or upper(org_spec_id) = upper(?)) " &_
                "and membership_id <> ?"
            cSql.AddParam(firstName)
            cSql.AddParam(lastName)
            cSql.AddParam(licNo)
            cSql.AddParam(membershipId)
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                existIndID = rsDup("membership_id")
            end if
            rsDup.close
            set rsDup = nothing
        end if
    end if
end if
'end nsbla

'nvslp
if Application("ORG_ID") = "3100" then
    ssn = replace(getFormVal("ssn")&"","-","")
    licType = getFormVal("lictype")
    dob = getFormVal("dob")
    if getFormVal("subsp")&"" <> "" then
        subType = getFormVal("subsp")
    end if
    if getFormVal("subaud")&"" <> "" then
        subType = getFormVal("subaud")
    end if
    if getFormVal("subhas")&"" <> "" then
        subType = getFormVal("subhas")
    end if
    childID = getFormVal("CHILD_ID")
    if appRenewType = "APP" then
        if ssn&"" <> "" and licType&"" <> "" and subType&"" <> "" then
            cSql.sqlStr = "select ind.membership_id indID,lic.membership_id licID " &_
                "from membership ind, membership lic " &_
                "where to_char(ind.membership_id) = lic.org_relate_id " &_
                "and lic.m_stat_cd = 'ACTIVE' and lic.org_sub_ty_cd = 'Credential' " &_
                "and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = lic.membership_id and column_desc = 'LEVEL_CD_Credential'),'') = ? " &_
                "and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = lic.membership_id and column_desc = 'Sub_License_Credential'),'') = ? " &_
                "and ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'MEMBER' " &_
                "and substr(replace(ind.SOC_SECURITY,'-',''),-4,4) = ? and ind.birth_dt = to_date(?,'mm/dd/yyyy') " &_
                "and ind.membership_id <> ? and lic.membership_id <> ?"
            cSql.AddParam(licType)
            cSql.AddParam(subType)
            cSql.AddParam(right(ssn,4))
            cSql.AddParam(dob)
            cSql.AddParam(membershipId)
            cSql.AddParam(childID)
            'cSql.Debug
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                existIndID = rsDup("indID")
                existLicID = rsDup("licID")
            end if
            rsDup.close
            set rsDup = nothing

            if not dupCheck then
                cSql.sqlStr = "select ind.membership_id indID " &_
                    "from membership ind " &_
                    "where ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'MEMBER' " &_
                    "and substr(replace(ind.SOC_SECURITY,'-',''),-4,4) = ? and ind.birth_dt = to_date(?,'mm/dd/yyyy') " &_
                    "and ind.membership_id <> ?"
                cSql.AddParam(right(ssn,4))
                cSql.AddParam(dob)
                cSql.AddParam(membershipId)
                'cSql.Debug
                set rsDup = cSql.Execute
                if rsDup.eof then
                else
                    dupCheck = true
                    existIndID = rsDup("indID")
                end if
                rsDup.close
                set rsDup = nothing
            end if
        end if
    end if
end if

'nvswe
if Application("ORG_ID") = "1800" then
    ssn = replace(getFormVal("ssn")&"","-","")
    childID = getFormVal("CHILD_ID")
    if appRenewType = "APP_LMSW" then
        if ssn&"" <> "" then
            cSql.sqlStr = "select ind.membership_id indID " &_
                "from membership ind " &_
                "where ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'MEMBER' " &_
                "and replace(ind.SOC_SECURITY,'-','') = ? " &_
                "and ind.membership_id <> ? "
            cSql.AddParam(ssn)
            cSql.AddParam(membershipId)
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                existIndID = rsDup("indID")
            end if
            rsDup.close
            set rsDup = nothing
        end if
    end if

    if appRenewType = "APP_LSW" then
        if ssn&"" <> "" then
            cSql.sqlStr = "select ind.membership_id indID " &_
                "from membership ind " &_
                "where ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'MEMBER' " &_
                "and replace(ind.SOC_SECURITY,'-','') = ? " &_
                "and ind.membership_id <> ? "
            cSql.AddParam(ssn)
            cSql.AddParam(membershipId)
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                existIndID = rsDup("indID")
            end if
            rsDup.close
            set rsDup = nothing
        end if
    end if

    if appRenewType = "APP" then
        if ssn&"" <> "" then
            cSql.sqlStr = "select ind.membership_id indID,lic.membership_id licID " &_
                "from membership ind, membership lic " &_
                "where to_char(ind.membership_id) = lic.org_relate_id " &_
                "and lic.m_stat_cd = 'ACTIVE' and lic.org_sub_ty_cd = 'CREDENTIAL' " &_
                "and ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'MEMBER' " &_
                "and replace(ind.SOC_SECURITY,'-','') = ? " &_
                "and ind.membership_id <> ? and lic.membership_id <> ?"
            cSql.AddParam(ssn)
            cSql.AddParam(membershipId)
            cSql.AddParam(childID)
            'cSql.Debug
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                existIndID = rsDup("indID")
                existLicID = rsDup("licID")
            end if
            rsDup.close
            set rsDup = nothing

            if not dupCheck then
                cSql.sqlStr = "select ind.membership_id indID,lic.membership_id licID " &_
                    "from membership ind, membership lic " &_
                    "where to_char(ind.membership_id) = lic.org_relate_id " &_
                    "and lic.m_stat_cd = 'ACTIVE' and lic.org_sub_ty_cd = 'CREDENTIAL' " &_
                    "and ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'MEMBER' " &_
                    "and replace(ind.SOC_SECURITY,'-','') = ? " &_
                    "and ind.membership_id <> ? and lic.membership_id <> ?"
                cSql.AddParam(ssn)
                cSql.AddParam(membershipId)
                cSql.AddParam(childID)
                'cSql.Debug
                set rsDup = cSql.Execute
                if rsDup.eof then
                else
                    dupCheck = true
                    existIndID = rsDup("indID")
                end if
                rsDup.close
                set rsDup = nothing
            end if
        end if
    end if
end if

'ndsbrc
if Application("ORG_ID") = "2900" then
    ssn = replace(getFormVal("ssn")&"","-","")
    if ssn&"" <> "" then
        cSql.sqlStr = "select membership_id " &_
            "from membership " &_
            "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MEMBER' " &_
            "and replace(SOC_SECURITY,'-','') = ?  " &_
            "and membership_id <> ? "
        cSql.AddParam(ssn)
        cSql.AddParam(membershipId)
        set rsDup = cSql.Execute
        if rsDup.eof then
        else
            dupCheck = true
            existIndID = rsDup("MEMBERSHIP_ID")
        end if
        rsDup.close
        set rsDup = nothing
    end if
end if

'ndrec
if Application("ORG_ID") = "2400" then
    ssn = replace(getFormVal("ssn")&"","-","")
    lastName = getFormVal("LAST_NAME")
    firstName = getFormVal("FIRST_NAME")
    if ssn&"" <> "" then
        cSql.sqlStr = "select membership_id " &_
            "from membership " &_
            "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MEMBER' " &_
            "and replace(SOC_SECURITY,'-','') = ?  " &_
            "and membership_id <> ? "
        cSql.AddParam(ssn)
        cSql.AddParam(membershipId)
        set rsDup = cSql.Execute
        if rsDup.eof then
        else
            dupCheck = true
            existIndID = rsDup("MEMBERSHIP_ID")
        end if
        rsDup.close
        set rsDup = nothing
        if not dupCheck then
            if firstName&"" <> "" and lastName&"" <> "" then
                cSql.sqlStr = "select membership_id " &_
                    "from membership " &_
                    "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MEMBER' " &_
                    "and upper(LAST_NAME)  = upper(?) " &_
                    "and upper(FIRST_NAME)  = upper(?) " &_
                    "and membership_id <> ? "
                cSql.AddParam(lastName)
                cSql.AddParam(firstName)
                cSql.AddParam(membershipId)
                set rsDup = cSql.Execute
                if rsDup.eof then
                else
                    dupCheck = true
                    existIndID = rsDup("MEMBERSHIP_ID")
                end if
                rsDup.close
                set rsDup = nothing
            end if
        end if
    end if
end if

'ndnar
if Application("ORG_ID") = "1900" then
    if appRenewType = "NA_RECRUIT" or appRenewType = "EMG_TRAIN" then
        childID = getFormVal("CHILD_ID")
        licType = "'6'"
        ssn = replace(getFormVal("ssn")&"","-","")
        if ssn = "" then
            ssn = replace(getFormVal("loginssn")&"","-","")
        end if
        lastName = getFormVal("LAST_NAME")

        if ssn&"" <> "" and lastName&"" <> "" then
            cSql.sqlStr = "select ind.membership_id indID,lic.membership_id licID " &_
                "from membership ind, membership lic " &_
                "where to_char(lic.membership_id) = ind.org_relate_id " &_
                "and lic.m_stat_cd = 'ACTIVE' and lic.org_sub_ty_cd = 'LIC' " &_
                "and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = lic.membership_id and column_desc = 'License_Type_LIC'),'') in (" & licType & ") " &_
                "and ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'MEMBER' " &_
                "and upper(ind.LAST_NAME)  = upper(?) " &_
                "and replace(ind.SOC_SECURITY,'-','') = ? " &_
                "and ind.membership_id <> ? and lic.membership_id <> ?"
            cSql.AddParam(lastName)
            cSql.AddParam(ssn)
            cSql.AddParam(membershipId)
            cSql.AddParam(childID)
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                dupNameSSNType = true
                existIndID = rsDup("indID")
                existLicID = rsDup("licID")
            end if
            rsDup.close
            set rsDup = nothing

            if not dupCheck then
                cSql.sqlStr = "select ind.membership_id indID,lic.membership_id licID " &_
                    "from membership ind, membership lic " &_
                    "where to_char(lic.membership_id) = ind.org_relate_id " &_
                    "and lic.m_stat_cd = 'ACTIVE' and lic.org_sub_ty_cd = 'LIC' and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = lic.membership_id and column_desc = 'License_Type_LIC'),'') in (" & licType & ") " &_
                    "and ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'MEMBER' " &_
                    "and replace(ind.SOC_SECURITY,'-','') = ? " &_
                    "and ind.membership_id <> ? and lic.membership_id <> ?"
                cSql.AddParam(ssn)
                cSql.AddParam(membershipId)
                cSql.AddParam(childID)
                set rsDup = cSql.Execute
                if rsDup.eof then
                else
                    dupCheck = true
                    dupSSNType = true
                    existIndID = rsDup("indID")
                    existLicID = rsDup("licID")
                end if
                rsDup.close
                set rsDup = nothing
            end if

            if not dupCheck then
                cSql.sqlStr = "select ind.membership_id indID " &_
                    "from membership ind " &_
                    "where ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'MEMBER' " &_
                    "and upper(ind.LAST_NAME)  = upper(?) " &_
                    "and replace(ind.SOC_SECURITY,'-','') = ? " &_
                    "and ind.membership_id <> ? "
                cSql.AddParam(lastName)
                cSql.AddParam(ssn)
                cSql.AddParam(membershipID)
                set rsDup = cSql.Execute
                if rsDup.eof then
                else
                    dupCheck = true
                    dupNameSSN = true
                    existIndID = rsDup("indID")
                end if
                rsDup.close
                set rsDup = nothing
            end if

            if not dupCheck then
                cSql.sqlStr = "select ind.membership_id indID " &_
                    "from membership ind " &_
                    "where ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'MEMBER' " &_
                    "and replace(ind.SOC_SECURITY,'-','') = ? " &_
                    "and ind.membership_id <> ? "
                cSql.AddParam(ssn)
                cSql.AddParam(membershipID)
                set rsDup = cSql.Execute
                if rsDup.eof then
                else
                    dupCheck = true
                    dupSSN = true
                    existIndID = rsDup("indID")
                end if
                rsDup.close
                set rsDup = nothing
            end if
        end if
    end if

    if appRenewType = "MA_APP" or appRenewType = "MA2_APP" or appRenewType = "NA_APP" or appRenewType = "CNA_APP" or appRenewType = "HHA_APP" then
        licType = "'3'"
        if appRenewType = "CNA_APP" then
            licType = "'1'"
        end if
        if appRenewType = "MA2_APP" then
            licType = "'4'"
        end if
        if appRenewType = "MA_APP" then
            licType = "'4','5'"
        end if
        ssn = replace(getFormVal("ssn")&"","-","")
        if ssn = "" then
            ssn = replace(getFormVal("loginssn")&"","-","")
        end if
        lastName = getFormVal("LAST_NAME")

        if ssn&"" <> "" and lastName&"" <> "" then
            cSql.sqlStr = "select ind.membership_id indID,lic.membership_id licID " &_
                "from membership ind, membership lic " &_
                "where to_char(ind.membership_id) = lic.org_relate_id " &_
                "and lic.m_stat_cd = 'ACTIVE' and lic.org_sub_ty_cd = 'LIC' " &_
                "and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = lic.membership_id and column_desc = 'License_Type_LIC'),'') in (" & licType & ") " &_
                "and ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'MEMBER' " &_
                "and upper(ind.LAST_NAME)  = upper(?) " &_
                "and replace(ind.SOC_SECURITY,'-','') = ? " &_
                "and ind.membership_id <> ? and lic.membership_id <> ?"
            cSql.AddParam(lastName)
            cSql.AddParam(ssn)
            cSql.AddParam(membershipId)
            cSql.AddParam(parentID)
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                dupNameSSNType = true
                existIndID = rsDup("indID")
                existLicID = rsDup("licID")
            end if
            rsDup.close
            set rsDup = nothing

            if not dupCheck then
                cSql.sqlStr = "select ind.membership_id indID,lic.membership_id licID " &_
                    "from membership ind, membership lic " &_
                    "where to_char(ind.membership_id) = lic.org_relate_id " &_
                    "and lic.m_stat_cd = 'ACTIVE' and lic.org_sub_ty_cd = 'LIC' and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = lic.membership_id and column_desc = 'License_Type_LIC'),'') in (" & licType & ") " &_
                    "and ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'MEMBER' " &_
                    "and replace(ind.SOC_SECURITY,'-','') = ? " &_
                    "and ind.membership_id <> ? and lic.membership_id <> ?"
                cSql.AddParam(ssn)
                cSql.AddParam(membershipId)
                cSql.AddParam(parentID)
                set rsDup = cSql.Execute
                if rsDup.eof then
                else
                    dupCheck = true
                    dupSSNType = true
                    existIndID = rsDup("indID")
                    existLicID = rsDup("licID")
                end if
                rsDup.close
                set rsDup = nothing
            end if

            if not dupCheck then
                cSql.sqlStr = "select ind.membership_id indID " &_
                    "from membership ind " &_
                    "where ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'MEMBER' " &_
                    "and upper(ind.LAST_NAME)  = upper(?) " &_
                    "and replace(ind.SOC_SECURITY,'-','') = ? " &_
                    "and ind.membership_id <> ? "
                cSql.AddParam(lastName)
                cSql.AddParam(ssn)
                cSql.AddParam(parentID)
                set rsDup = cSql.Execute
                if rsDup.eof then
                else
                    dupCheck = true
                    dupNameSSN = true
                    existIndID = rsDup("indID")
                end if
                rsDup.close
                set rsDup = nothing
            end if

            if not dupCheck then
                cSql.sqlStr = "select ind.membership_id indID " &_
                    "from membership ind " &_
                    "where ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'MEMBER' " &_
                    "and replace(ind.SOC_SECURITY,'-','') = ? " &_
                    "and ind.membership_id <> ? "
                cSql.AddParam(ssn)
                cSql.AddParam(parentID)
                set rsDup = cSql.Execute
                if rsDup.eof then
                else
                    dupCheck = true
                    dupSSN = true
                    existIndID = rsDup("indID")
                end if
                rsDup.close
                set rsDup = nothing
            end if
        end if
    end if
end if

'wvbop
if Application("ORG_ID") = "1450" then
    if appRenewType = "COVID-19_APP" then
        ssn = replace(getFormVal("ssn")&"","-","")
        lastName = getFormVal("LAST_NAME")
        if ssn&"" <> "" and lastName&"" <> "" then
            cSql.sqlStr = "select ind.membership_id indID " &_
                "from membership ind " &_
                "where ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'MEMBER' " &_
                "and ind.LAST_NAME  = ? " &_
                "and replace(ind.SOC_SECURITY,'-','') = ? "
            cSql.AddParam(lastName)
            cSql.AddParam(ssn)
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                dupNameSSNType = true
                existIndID = rsDup("indID")
            end if
            rsDup.close
            set rsDup = nothing
            if not dupCheck then
                cSql.sqlStr = "select ind.membership_id indID " &_
                    "from membership ind " &_
                    "where ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'INDV' " &_
                    "and replace(ind.SOC_SECURITY,'-','') = ? "
                cSql.AddParam(ssn)
                set rsDup = cSql.Execute
                if rsDup.eof then
                else
                    dupCheck = true
                    dupSSNType = true
                    existIndID = rsDup("indID")
                end if
                rsDup.close
                set rsDup = nothing
            end if
        end if
    end if

    if appRenewType = "NT_APP" then
        ssn = replace(getFormVal("ssn")&"","-","")
        lastName = getFormVal("LAST_NAME")

        if ssn&"" <> "" and lastName&"" <> "" then
            cSql.sqlStr = "select ind.membership_id indID,lic.membership_id licID " &_
                "from membership ind, membership lic " &_
                "where to_char(ind.membership_id) = lic.org_relate_id " &_
                "and lic.m_stat_cd = 'ACTIVE' and lic.org_sub_ty_cd = 'LIC' and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = lic.membership_id and column_desc = 'License_Type_LIC'),'') = 'NT' " &_
                "and ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'MEMBER' " &_
                "and ind.LAST_NAME  = ? " &_
                "and replace(ind.SOC_SECURITY,'-','') = ? "
            cSql.AddParam(lastName)
            cSql.AddParam(ssn)
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                dupNameSSNType = true
                existIndID = rsDup("indID")
                existLicID = rsDup("licID")
            end if
            rsDup.close
            set rsDup = nothing

            if not dupCheck then
                cSql.sqlStr = "select ind.membership_id indID,lic.membership_id licID " &_
                    "from membership ind, membership lic " &_
                    "where to_char(ind.membership_id) = lic.org_relate_id " &_
                    "and lic.m_stat_cd = 'ACTIVE' and lic.org_sub_ty_cd = 'LIC' and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = lic.membership_id and column_desc = 'License_Type_LIC'),'') = 'NT' " &_
                    "and ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'INDV' " &_
                    "and replace(ind.SOC_SECURITY,'-','') = ? "
                cSql.AddParam(ssn)
                set rsDup = cSql.Execute
                if rsDup.eof then
                else
                    dupCheck = true
                    dupSSNType = true
                    existIndID = rsDup("indID")
                    existLicID = rsDup("licID")
                end if
                rsDup.close
                set rsDup = nothing
            end if

            if not dupCheck then
                cSql.sqlStr = "select ind.membership_id indID " &_
                    "from membership ind " &_
                    "where ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'INDV' " &_
                    "and ind.LAST_NAME  = ? " &_
                    "and replace(ind.SOC_SECURITY,'-','') = ? "
                cSql.AddParam(lastName)
                cSql.AddParam(ssn)
                set rsDup = cSql.Execute
                if rsDup.eof then
                else
                    dupCheck = true
                    dupNameSSN = true
                    existIndID = rsDup("indID")
                end if
                rsDup.close
                set rsDup = nothing
            end if

            if not dupCheck then
                cSql.sqlStr = "select ind.membership_id indID " &_
                    "from membership ind " &_
                    "where ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'INDV' " &_
                    "and replace(ind.SOC_SECURITY,'-','') = ? "
                cSql.AddParam(ssn)
                set rsDup = cSql.Execute
                if rsDup.eof then
                else
                    dupCheck = true
                    dupSSN = true
                    existIndID = rsDup("indID")
                end if
                rsDup.close
                set rsDup = nothing
            end if
        end if
    end if

    if appRenewType = "TT_APP" then
        ssn = replace(getFormVal("ssn")&"","-","")
        lastName = getFormVal("LAST_NAME")

        if ssn&"" <> "" and lastName&"" <> "" then
            cSql.sqlStr = "select ind.membership_id indID,lic.membership_id licID " &_
                "from membership ind, membership lic " &_
                "where to_char(ind.membership_id) = lic.org_relate_id " &_
                "and lic.m_stat_cd = 'ACTIVE' and lic.org_sub_ty_cd = 'LIC' and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = lic.membership_id and column_desc = 'License_Type_LIC'),'') = 'TT' " &_
                "and ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'INDV' " &_
                "and ind.LAST_NAME  = ? " &_
                "and replace(ind.SOC_SECURITY,'-','') = ? "
            cSql.AddParam(lastName)
            cSql.AddParam(ssn)
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                dupNameSSNType = true
                existIndID = rsDup("indID")
                existLicID = rsDup("licID")
            end if
            rsDup.close
            set rsDup = nothing

            if not dupCheck then
                cSql.sqlStr = "select ind.membership_id indID,lic.membership_id licID " &_
                    "from membership ind, membership lic " &_
                    "where to_char(ind.membership_id) = lic.org_relate_id " &_
                    "and lic.m_stat_cd = 'ACTIVE' and lic.org_sub_ty_cd = 'LIC' and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = lic.membership_id and column_desc = 'License_Type_LIC'),'') = 'TT' " &_
                    "and ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'INDV' " &_
                    "and replace(ind.SOC_SECURITY,'-','') = ? "
                cSql.AddParam(ssn)
                set rsDup = cSql.Execute
                if rsDup.eof then
                else
                    dupCheck = true
                    dupSSNType = true
                    existIndID = rsDup("indID")
                    existLicID = rsDup("licID")
                end if
                rsDup.close
                set rsDup = nothing
            end if

            if not dupCheck then
                cSql.sqlStr = "select ind.membership_id indID " &_
                    "from membership ind " &_
                    "where ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'INDV' " &_
                    "and ind.LAST_NAME  = ? " &_
                    "and replace(ind.SOC_SECURITY,'-','') = ? "
                cSql.AddParam(lastName)
                cSql.AddParam(ssn)
                set rsDup = cSql.Execute
                if rsDup.eof then
                else
                    dupCheck = true
                    dupNameSSN = true
                    existIndID = rsDup("indID")
                end if
                rsDup.close
                set rsDup = nothing
            end if

            if not dupCheck then
                cSql.sqlStr = "select ind.membership_id indID " &_
                    "from membership ind " &_
                    "where ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'INDV' " &_
                    "and replace(ind.SOC_SECURITY,'-','') = ? "
                cSql.AddParam(ssn)
                set rsDup = cSql.Execute
                if rsDup.eof then
                else
                    dupCheck = true
                    dupSSN = true
                    existIndID = rsDup("indID")
                end if
                rsDup.close
                set rsDup = nothing
            end if
        end if
    end if

    if appRenewType = "RP_APP" then
        ssn = replace(getFormVal("ssn")&"","-","")
        lastName = getFormVal("LAST_NAME")

        if ssn&"" <> "" and lastName&"" <> "" then
            cSql.sqlStr = "select ind.membership_id indID,lic.membership_id licID " &_
                "from membership ind, membership lic " &_
                "where to_char(ind.membership_id) = lic.org_relate_id " &_
                "and lic.m_stat_cd = 'ACTIVE' and lic.org_sub_ty_cd = 'LIC' and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = lic.membership_id and column_desc = 'License_Type_LIC'),'') = 'RP' " &_
                "and ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'INDV' " &_
                "and ind.LAST_NAME  = ? " &_
                "and replace(ind.SOC_SECURITY,'-','') = ? "
            cSql.AddParam(lastName)
            cSql.AddParam(ssn)
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                dupNameSSNType = true
                existIndID = rsDup("indID")
                existLicID = rsDup("licID")
            end if
            rsDup.close
            set rsDup = nothing

            if not dupCheck then
                cSql.sqlStr = "select ind.membership_id indID,lic.membership_id licID " &_
                    "from membership ind, membership lic " &_
                    "where to_char(ind.membership_id) = lic.org_relate_id " &_
                    "and lic.m_stat_cd = 'ACTIVE' and lic.org_sub_ty_cd = 'LIC' and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = lic.membership_id and column_desc = 'License_Type_LIC'),'') = 'RP' " &_
                    "and ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'INDV' " &_
                    "and replace(ind.SOC_SECURITY,'-','') = ? "
                cSql.AddParam(ssn)
                set rsDup = cSql.Execute
                if rsDup.eof then
                else
                    dupCheck = true
                    dupSSNType = true
                    existIndID = rsDup("indID")
                    existLicID = rsDup("licID")
                end if
                rsDup.close
                set rsDup = nothing
            end if

            if not dupCheck then
                cSql.sqlStr = "select ind.membership_id indID " &_
                    "from membership ind " &_
                    "where ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'INDV' " &_
                    "and ind.LAST_NAME  = ? " &_
                    "and replace(ind.SOC_SECURITY,'-','') = ? "
                cSql.AddParam(lastName)
                cSql.AddParam(ssn)
                set rsDup = cSql.Execute
                if rsDup.eof then
                else
                    dupCheck = true
                    dupNameSSN = true
                    existIndID = rsDup("indID")
                end if
                rsDup.close
                set rsDup = nothing
            end if

            if not dupCheck then
                cSql.sqlStr = "select ind.membership_id indID " &_
                    "from membership ind " &_
                    "where ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'INDV' " &_
                    "and replace(ind.SOC_SECURITY,'-','') = ? "
                cSql.AddParam(ssn)
                set rsDup = cSql.Execute
                if rsDup.eof then
                else
                    dupCheck = true
                    dupSSN = true
                    existIndID = rsDup("indID")
                end if
                rsDup.close
                set rsDup = nothing
            end if
        end if
    end if

    if appRenewType = "IN_APP" then
        ssn = replace(getFormVal("ssn")&"","-","")
        lastName = getFormVal("LAST_NAME")

        if ssn&"" <> "" and lastName&"" <> "" then
            cSql.sqlStr = "select ind.membership_id indID,lic.membership_id licID " &_
                "from membership ind, membership lic " &_
                "where to_char(ind.membership_id) = lic.org_relate_id " &_
                "and lic.m_stat_cd = 'ACTIVE' and lic.org_sub_ty_cd = 'LIC' and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = lic.membership_id and column_desc = 'License_Type_LIC'),'') = 'IN' " &_
                "and ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'INDV' " &_
                "and ind.LAST_NAME  = ? " &_
                "and replace(ind.SOC_SECURITY,'-','') = ? "
            cSql.AddParam(lastName)
            cSql.AddParam(ssn)
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                dupNameSSNType = true
                existIndID = rsDup("indID")
                existLicID = rsDup("licID")
            end if
            rsDup.close
            set rsDup = nothing

            if not dupCheck then
                cSql.sqlStr = "select ind.membership_id indID,lic.membership_id licID " &_
                    "from membership ind, membership lic " &_
                    "where to_char(ind.membership_id) = lic.org_relate_id " &_
                    "and lic.m_stat_cd = 'ACTIVE' and lic.org_sub_ty_cd = 'LIC' and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = lic.membership_id and column_desc = 'License_Type_LIC'),'') = 'IN' " &_
                    "and ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'INDV' " &_
                    "and replace(ind.SOC_SECURITY,'-','') = ? "
                cSql.AddParam(ssn)
                set rsDup = cSql.Execute
                if rsDup.eof then
                else
                    dupCheck = true
                    dupSSNType = true
                    existIndID = rsDup("indID")
                    existLicID = rsDup("licID")
                end if
                rsDup.close
                set rsDup = nothing
            end if

            if not dupCheck then
                cSql.sqlStr = "select ind.membership_id indID " &_
                    "from membership ind " &_
                    "where ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'INDV' " &_
                    "and ind.LAST_NAME  = ? " &_
                    "and replace(ind.SOC_SECURITY,'-','') = ? "
                cSql.AddParam(lastName)
                cSql.AddParam(ssn)
                set rsDup = cSql.Execute
                if rsDup.eof then
                else
                    dupCheck = true
                    dupNameSSN = true
                    existIndID = rsDup("indID")
                end if
                rsDup.close
                set rsDup = nothing
            end if

            if not dupCheck then
                cSql.sqlStr = "select ind.membership_id indID " &_
                    "from membership ind " &_
                    "where ind.m_stat_cd = 'ACTIVE' and ind.org_sub_ty_cd = 'INDV' " &_
                    "and replace(ind.SOC_SECURITY,'-','') = ? "
                cSql.AddParam(ssn)
                set rsDup = cSql.Execute
                if rsDup.eof then
                else
                    dupCheck = true
                    dupSSN = true
                    existIndID = rsDup("indID")
                end if
                rsDup.close
                set rsDup = nothing
            end if
        end if
    end if

    if appRenewType = "PHAR_APP" or appRenewType = "CSP_APP" then
        appTyp = getFormVal("APP_TYP")
        zip = getFormVal("phyZip")
        phone = getFormVal("telNo")
        if zip&"" <> "" and phone&"" <> "" then
            cSql.sqlStr = "select cred.membership_id credID,fac.membership_id facID " &_
                "from membership cred, membership fac " &_
                "where to_char(fac.membership_id) = cred.org_relate_id " &_
                "and cred.m_stat_cd = 'ACTIVE' and cred.org_sub_ty_cd = 'CRED' and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = cred.membership_id and column_desc = 'Type_CRED'),'') = ? " &_
                "and fac.m_stat_cd = 'ACTIVE' and fac.org_sub_ty_cd = 'BIZ' " &_
                "and fac.zip = ? " &_
                "and fac.phone_day = ? "
            cSql.AddParam(appTyp)
            cSql.AddParam(zip)
            cSql.AddParam(phone)
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                existID = rsDup("credID")
                existFacID = rsDup("facID")
            end if
            rsDup.close
            set rsDup = nothing
        end if
        if not dupCheck and zip&"" <> "" then
            cSql.sqlStr = "select cred.membership_id credID,fac.membership_id facID " &_
                "from membership cred, membership fac " &_
                "where to_char(fac.membership_id) = cred.org_relate_id " &_
                "and cred.m_stat_cd = 'ACTIVE' and cred.org_sub_ty_cd = 'CRED' and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = cred.membership_id and column_desc = 'Type_CRED'),'') = ? " &_
                "and fac.m_stat_cd = 'ACTIVE' and fac.org_sub_ty_cd = 'BIZ' " &_
                "and fac.zip = ? "
            cSql.AddParam(appTyp)
            cSql.AddParam(zip)
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                existID = rsDup("credID")
                existFacID = rsDup("facID")
            end if
            rsDup.close
            set rsDup = nothing
        end if
    end if

    if appRenewType = "MFR_APP" then
        zip = getFormVal("phyZip")
        phone = getFormVal("telNo")
        if zip&"" <> "" and phone&"" <> "" then
            cSql.sqlStr = "select cred.membership_id credID,fac.membership_id facID " &_
                "from membership cred, membership fac " &_
                "where to_char(fac.membership_id) = cred.org_relate_id " &_
                "and cred.m_stat_cd = 'ACTIVE' and cred.org_sub_ty_cd = 'CRED' and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = cred.membership_id and column_desc = 'Type_CRED'),'') = 'MR' " &_
                "and fac.m_stat_cd = 'ACTIVE' and fac.org_sub_ty_cd = 'BIZ' " &_
                "and fac.zip = ? " &_
                "and fac.phone_day = ? "
            cSql.AddParam(zip)
            cSql.AddParam(phone)
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                existID = rsDup("credID")
                existFacID = rsDup("facID")
            end if
            rsDup.close
            set rsDup = nothing
        end if
        if not dupCheck and zip&"" <> "" then
            cSql.sqlStr = "select cred.membership_id credID,fac.membership_id facID " &_
                "from membership cred, membership fac " &_
                "where to_char(fac.membership_id) = cred.org_relate_id " &_
                "and cred.m_stat_cd = 'ACTIVE' and cred.org_sub_ty_cd = 'CRED' and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = cred.membership_id and column_desc = 'Type_CRED'),'') = 'MR' " &_
                "and fac.m_stat_cd = 'ACTIVE' and fac.org_sub_ty_cd = 'BIZ' " &_
                "and fac.zip = ? "
            cSql.AddParam(zip)
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                existID = rsDup("credID")
                existFacID = rsDup("facID")
            end if
            rsDup.close
            set rsDup = nothing
        end if
    end if

    if appRenewType = "MOP_APP" then
        zip = getFormVal("phyZip")
        phone = getFormVal("telNo")
        if zip&"" <> "" and phone&"" <> "" then
            cSql.sqlStr = "select cred.membership_id credID,fac.membership_id facID " &_
                "from membership cred, membership fac " &_
                "where to_char(fac.membership_id) = cred.org_relate_id " &_
                "and cred.m_stat_cd = 'ACTIVE' and cred.org_sub_ty_cd = 'CRED' and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = cred.membership_id and column_desc = 'Type_CRED'),'') = 'MO' " &_
                "and fac.m_stat_cd = 'ACTIVE' and fac.org_sub_ty_cd = 'BIZ' " &_
                "and fac.zip = ? " &_
                "and fac.phone_day = ? "
            cSql.AddParam(zip)
            cSql.AddParam(phone)
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                existID = rsDup("credID")
                existFacID = rsDup("facID")
            end if
            rsDup.close
            set rsDup = nothing
        end if
        if not dupCheck and zip&"" <> "" then
            cSql.sqlStr = "select cred.membership_id credID,fac.membership_id facID " &_
                "from membership cred, membership fac " &_
                "where to_char(fac.membership_id) = cred.org_relate_id " &_
                "and cred.m_stat_cd = 'ACTIVE' and cred.org_sub_ty_cd = 'CRED' and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = cred.membership_id and column_desc = 'Type_CRED'),'') = 'MO' " &_
                "and fac.m_stat_cd = 'ACTIVE' and fac.org_sub_ty_cd = 'BIZ' " &_
                "and fac.zip = ? "
            cSql.AddParam(zip)
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                existID = rsDup("credID")
                existFacID = rsDup("facID")
            end if
            rsDup.close
            set rsDup = nothing
        end if
    end if

    if appRenewType = "WDD_APP" then
        zip = getFormVal("phyZip")
        phone = getFormVal("telNo")
        if zip&"" <> "" and phone&"" <> "" then
            cSql.sqlStr = "select cred.membership_id credID,fac.membership_id facID " &_
                "from membership cred, membership fac " &_
                "where to_char(fac.membership_id) = cred.org_relate_id " &_
                "and cred.m_stat_cd = 'ACTIVE' and cred.org_sub_ty_cd = 'CRED' and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = cred.membership_id and column_desc = 'Type_CRED'),'') = 'WD' " &_
                "and fac.m_stat_cd = 'ACTIVE' and fac.org_sub_ty_cd = 'BIZ' " &_
                "and fac.zip = ? " &_
                "and fac.phone_day = ? "
            cSql.AddParam(zip)
            cSql.AddParam(phone)
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                existID = rsDup("credID")
                existFacID = rsDup("facID")
            end if
            rsDup.close
            set rsDup = nothing
        end if
        if not dupCheck and zip&"" <> "" then
            cSql.sqlStr = "select cred.membership_id credID,fac.membership_id facID " &_
                "from membership cred, membership fac " &_
                "where to_char(fac.membership_id) = cred.org_relate_id " &_
                "and cred.m_stat_cd = 'ACTIVE' and cred.org_sub_ty_cd = 'CRED' and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = cred.membership_id and column_desc = 'Type_CRED'),'') = 'WD' " &_
                "and fac.m_stat_cd = 'ACTIVE' and fac.org_sub_ty_cd = 'BIZ' " &_
                "and fac.zip = ? "
            cSql.AddParam(zip)
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                existID = rsDup("credID")
                existFacID = rsDup("facID")
            end if
            rsDup.close
            set rsDup = nothing
        end if
    end if

    if appRenewType = "TPLP_APP" then
        zip = getFormVal("phyZip")
        phone = getFormVal("telNo")
        if zip&"" <> "" and phone&"" <> "" then
            cSql.sqlStr = "select cred.membership_id credID,fac.membership_id facID " &_
                "from membership cred, membership fac " &_
                "where to_char(fac.membership_id) = cred.org_relate_id " &_
                "and cred.m_stat_cd = 'ACTIVE' and cred.org_sub_ty_cd = 'CRED' and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = cred.membership_id and column_desc = 'Type_CRED'),'') = '3PL' " &_
                "and fac.m_stat_cd = 'ACTIVE' and fac.org_sub_ty_cd = 'BIZ' " &_
                "and fac.zip = ? " &_
                "and fac.phone_day = ? "
            cSql.AddParam(zip)
            cSql.AddParam(phone)
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                existID = rsDup("credID")
                existFacID = rsDup("facID")
            end if
            rsDup.close
            set rsDup = nothing
        end if
        if not dupCheck and zip&"" <> "" then
            cSql.sqlStr = "select cred.membership_id credID,fac.membership_id facID " &_
                "from membership cred, membership fac " &_
                "where to_char(fac.membership_id) = cred.org_relate_id " &_
                "and cred.m_stat_cd = 'ACTIVE' and cred.org_sub_ty_cd = 'CRED' and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = cred.membership_id and column_desc = 'Type_CRED'),'') = '3PL' " &_
                "and fac.m_stat_cd = 'ACTIVE' and fac.org_sub_ty_cd = 'BIZ' " &_
                "and fac.zip = ? "
            cSql.AddParam(zip)
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                existID = rsDup("credID")
                existFacID = rsDup("facID")
            end if
            rsDup.close
            set rsDup = nothing
        end if
    end if

    if appRenewType = "LP_APP" then
        zip = getFormVal("phyZip")
        phone = getFormVal("telNo")
        if zip&"" <> "" and phone&"" <> "" then
            cSql.sqlStr = "select cred.membership_id credID,fac.membership_id facID " &_
                "from membership cred, membership fac " &_
                "where to_char(fac.membership_id) = cred.org_relate_id " &_
                "and cred.m_stat_cd = 'ACTIVE' and cred.org_sub_ty_cd = 'CRED' and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = cred.membership_id and column_desc = 'Type_CRED'),'') = 'LP' " &_
                "and fac.m_stat_cd = 'ACTIVE' and fac.org_sub_ty_cd = 'BIZ' " &_
                "and fac.zip = ? " &_
                "and fac.phone_day = ? "
            cSql.AddParam(zip)
            cSql.AddParam(phone)
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                existID = rsDup("credID")
                existFacID = rsDup("facID")
            end if
            rsDup.close
            set rsDup = nothing
        end if
        if not dupCheck and zip&"" <> "" then
            cSql.sqlStr = "select cred.membership_id credID,fac.membership_id facID " &_
                "from membership cred, membership fac " &_
                "where to_char(fac.membership_id) = cred.org_relate_id " &_
                "and cred.m_stat_cd = 'ACTIVE' and cred.org_sub_ty_cd = 'CRED' and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = cred.membership_id and column_desc = 'Type_CRED'),'') = 'LP' " &_
                "and fac.m_stat_cd = 'ACTIVE' and fac.org_sub_ty_cd = 'BIZ' " &_
                "and fac.zip = ? "
            cSql.AddParam(zip)
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                existID = rsDup("credID")
                existFacID = rsDup("facID")
            end if
            rsDup.close
            set rsDup = nothing
        end if
    end if

    if appRenewType = "CCP_APP" then
        zip = getFormVal("phyZip")
        phone = getFormVal("telNo")
        if zip&"" <> "" and phone&"" <> "" then
            cSql.sqlStr = "select cred.membership_id credID,fac.membership_id facID " &_
                "from membership cred, membership fac " &_
                "where to_char(fac.membership_id) = cred.org_relate_id " &_
                "and cred.m_stat_cd = 'ACTIVE' and cred.org_sub_ty_cd = 'CRED' and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = cred.membership_id and column_desc = 'Type_CRED'),'') = 'CC' " &_
                "and fac.m_stat_cd = 'ACTIVE' and fac.org_sub_ty_cd = 'BIZ' " &_
                "and fac.zip = ? " &_
                "and fac.phone_day = ? "
            cSql.AddParam(zip)
            cSql.AddParam(phone)
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                existID = rsDup("credID")
                existFacID = rsDup("facID")
            end if
            rsDup.close
            set rsDup = nothing
        end if
        if not dupCheck and zip&"" <> "" then
            cSql.sqlStr = "select cred.membership_id credID,fac.membership_id facID " &_
                "from membership cred, membership fac " &_
                "where to_char(fac.membership_id) = cred.org_relate_id " &_
                "and cred.m_stat_cd = 'ACTIVE' and cred.org_sub_ty_cd = 'CRED' and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = cred.membership_id and column_desc = 'Type_CRED'),'') = 'CC' " &_
                "and fac.m_stat_cd = 'ACTIVE' and fac.org_sub_ty_cd = 'BIZ' " &_
                "and fac.zip = ? "
            cSql.AddParam(zip)
            set rsDup = cSql.Execute
            if rsDup.eof then
            else
                dupCheck = true
                existID = rsDup("credID")
                existFacID = rsDup("facID")
            end if
            rsDup.close
            set rsDup = nothing
        end if
    end if
end if

if Application("ORG_ID") = "1300" then
    if appRenewType = "admissionevaloffical" then
        ssn = getFormVal("ssn")
        birthDt = getFormVal("dob")
        lastName = getFormVal("LAST_NAME")
        maidenName = getFormVal("MAIDEN_NAME")
        if birthDt <> "" and ssn <> "" and lastName <> "" and maidenName <> "" then
            cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MEMBER' " &_
                "and soc_security = ? and birth_dt = to_date(?,'mm/dd/yyyy') " &_
                "and lower(last_name) = lower(?) " &_
                "and lower(maiden_name) = lower(?) " &_
                "and membership_id <> ?"
            cSql.AddParam(ssn)
            cSql.AddParam(birthDt)
            cSql.AddParam(lastName)
            cSql.AddParam(maidenName)
            cSql.AddParam(membershipId)
            'cSql.Debug
			set rsDup = cSql.Execute
			if rsDup.eof and rsDup.bof then
			else
			    existID = rsDup("MEMBERSHIP_ID")
				dupCheck = true
                dupSSN = true
			end if

            cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MEMBER' " &_
                "and soc_security = ? " &_
                "and membership_id <> ?"
            cSql.AddParam(ssn)
            cSql.AddParam(membershipId)
            'cSql.Debug
			set rsDup = cSql.Execute
			if rsDup.eof and rsDup.bof then
			else
			    existID = rsDup("MEMBERSHIP_ID")
				dupCheck = true
                dupSSN = true
			end if

            cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MEMBER' " &_
                "and soc_security = ? " &_
                "and lower(last_name) = lower(?) " &_
                "and lower(maiden_name) = lower(?) " &_
                "and membership_id <> ?"
            cSql.AddParam(ssn)
            cSql.AddParam(lastName)
            cSql.AddParam(maidenName)
            cSql.AddParam(membershipId)
            'cSql.Debug
			set rsDup = cSql.Execute
			if rsDup.eof and rsDup.bof then
			else
			    existID = rsDup("MEMBERSHIP_ID")
				dupCheck = true
                dupSSN = true
			end if

            cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MEMBER' " &_
                "and soc_security = ? and birth_dt = to_date(?,'mm/dd/yyyy') " &_
                "and lower(last_name) = lower(?) " &_
                "and membership_id <> ?"
            cSql.AddParam(ssn)
            cSql.AddParam(birthDt)
            cSql.AddParam(lastName)
            cSql.AddParam(membershipId)
            'cSql.Debug
			set rsDup = cSql.Execute
			if rsDup.eof and rsDup.bof then
			else
			    existID = rsDup("MEMBERSHIP_ID")
				dupCheck = true
                dupSSN = true
			end if

            cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MEMBER' " &_
                "and soc_security = ? and birth_dt = to_date(?,'mm/dd/yyyy') " &_
                "and lower(maiden_name) = lower(?) " &_
                "and membership_id <> ?"
            cSql.AddParam(ssn)
            cSql.AddParam(birthDt)
            cSql.AddParam(maidenName)
            cSql.AddParam(membershipId)
            'cSql.Debug
			set rsDup = cSql.Execute
			if rsDup.eof and rsDup.bof then
			else
			    existID = rsDup("MEMBERSHIP_ID")
				dupCheck = true
                dupSSN = true
			end if

            cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MEMBER' " &_
                "and soc_security = ? " &_
                "and lower(last_name) = lower(?) " &_
                "and membership_id <> ?"
            cSql.AddParam(ssn)
            cSql.AddParam(lastName)
            cSql.AddParam(membershipId)
            'cSql.Debug
			set rsDup = cSql.Execute
			if rsDup.eof and rsDup.bof then
			else
			    existID = rsDup("MEMBERSHIP_ID")
				dupCheck = true
                dupSSN = true
			end if

            cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MEMBER' " &_
                "and soc_security = ? " &_
                "and lower(maiden_name) = lower(?) " &_
                "and membership_id <> ?"
            cSql.AddParam(ssn)
            cSql.AddParam(maidenName)
            cSql.AddParam(membershipId)
            'cSql.Debug
			set rsDup = cSql.Execute
			if rsDup.eof and rsDup.bof then
			else
			    existID = rsDup("MEMBERSHIP_ID")
				dupCheck = true
                dupSSN = true
			end if

        end if      
    end if

    if appRenewType = "TRNFREXAMAPP" or appRenewType = "EXAMAPP" or appRenewType = "RECIPCERTWV" then
        ssn = getFormVal("ssn")
        birthDt = getFormVal("dob")
        lastName = getFormVal("LAST_NAME")
        if birthDt <> "" and ssn <> "" and lastName <> "" then
            cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MEMBER' " &_
                "and soc_security = ? and birth_dt = to_date(?,'mm/dd/yyyy') " &_
                "and lower(last_name) = lower(?) " &_
                "and membership_id <> ?"
            cSql.AddParam(ssn)
            cSql.AddParam(birthDt)
            cSql.AddParam(lastName)
            cSql.AddParam(membershipId)
            'cSql.Debug
			set rsDup = cSql.Execute
			if rsDup.eof and rsDup.bof then
			else
			    existID = rsDup("MEMBERSHIP_ID")
				dupCheck = true
				dupSSNBirthName = true
			end if
			if not dupCheck then
                cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                    "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MEMBER' " &_
                    "and soc_security = ? " &_
                    "and birth_dt = to_date(?,'mm/dd/yyyy')  "  &_
                    "and membership_id <> ?"
                cSql.AddParam(ssn)
                cSql.AddParam(birthDt)
                cSql.AddParam(membershipId)
                'cSql.Debug
			    set rsDup = cSql.Execute
			    if rsDup.eof and rsDup.bof then
			    else
			        existID = rsDup("MEMBERSHIP_ID")
				    dupCheck = true
				    dupSSNBirth = true
			    end if
			end if
        end if      
    end if
end if

if Application("ORG_ID") = "900" then
    if appRenewType = "License Application" or appRenewType = "Provisional Permit Application" then
        ssn = getFormVal("SSN")
        birthDt = getFormVal("DOB")
        lastName = getFormVal("LAST_NAME")
        
        if birthDt <> "" and ssn <> "" and lastName <> "" then
            cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MEMBER' " &_
                "and soc_security = ? and birth_dt = to_date(?,'mm/dd/yyyy') " &_
                "and lower(last_name) = lower(?) " &_
                "and membership_id <> ?"
            cSql.AddParam(ssn)
            cSql.AddParam(birthDt)
            cSql.AddParam(lastName)
            cSql.AddParam(membershipId)
            'cSql.Debug
			set rsDup = cSql.Execute
			if rsDup.eof and rsDup.bof then
			else
			    existID = rsDup("MEMBERSHIP_ID")
				dupCheck = true
				dupSSNBirthName = true
			end if
			if not dupCheck then
                cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                    "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MEMBER' " &_
                    "and soc_security = ? " &_
                    "and birth_dt = to_date(?,'mm/dd/yyyy')  "  &_
                    "and membership_id <> ?"
                cSql.AddParam(ssn)
                cSql.AddParam(birthDt)
                cSql.AddParam(membershipId)
                'cSql.Debug
			    set rsDup = cSql.Execute
			    if rsDup.eof and rsDup.bof then
			    else
			        existID = rsDup("MEMBERSHIP_ID")
				    dupCheck = true
				    dupSSNBirth = true
			    end if
			end if
	    end if
    end if
end if

'sduap
if Application("ORG_ID") = "750" then
    if appRenewType = "CAPN" or appRenewType = "CAAN" then
        ssn = getFormVal("ssn")
        if ssn <> "" then
            cSql.SqlStr = "select membership.membership_id from membership, membership lic " &_
                "where lic.org_relate_id = to_char(membership.membership_id) " &_
                "and membership.m_stat_cd = 'ACTIVE' and membership.org_sub_ty_cd = 'MEMBER' " &_
                "and membership.soc_security = ? " &_
                "and membership.membership_id <> ?"
            cSql.AddParam(ssn)
            cSql.AddParam(membershipId)
			set rsDup = cSql.Execute
			if rsDup.eof and rsDup.bof then
			else
			    existID = rsDup("MEMBERSHIP_ID")
				dupCheck = true
			end if
        end if
    end if

    if (appRenewType = "UDAAPP" or appRenewType = "UMAAPP") then
        ssn = getFormVal("ssn")
        birthDt = getFormVal("dob")
        lastName = getFormVal("LAST_NAME")

        if birthDt <> "" and ssn <> "" and lastName <> "" then
            cSql.SqlStr = "select membership.membership_id from membership, membership lic " &_
                "where lic.org_relate_id = to_char(membership.membership_id) " &_
                "and membership.m_stat_cd = 'ACTIVE' and membership.org_sub_ty_cd = 'MEMBER' " &_
                "and membership.soc_security = ? and membership.birth_dt = to_date(?,'mm/dd/yyyy') " &_
                "and lower(membership.last_name) = lower(?) " &_
                "and membership.membership_id <> ?"
            cSql.AddParam(ssn)
            cSql.AddParam(birthDt)
            cSql.AddParam(lastName)
            cSql.AddParam(membershipId)
            'cSql.Debug
			set rsDup = cSql.Execute
			if rsDup.eof and rsDup.bof then
			else
			    existID = rsDup("MEMBERSHIP_ID")
				dupCheck = true
				dupSSNBirthNameLic = true
			end if
			if not dupCheck then
                cSql.SqlStr = "select membership.membership_id, lic.membership_id as lic_id from membership, membership lic " &_
                    "where lic.org_relate_id = to_char(membership.membership_id) " &_
                    "and membership.m_stat_cd = 'ACTIVE' and membership.org_sub_ty_cd = 'MEMBER' " &_
                    "and membership.soc_security = ? " &_
                    "and membership.birth_dt = to_date(?,'mm/dd/yyyy')  "  &_
                    "and membership.membership_id <> ?"
                cSql.AddParam(ssn)
                cSql.AddParam(birthDt)
                cSql.AddParam(membershipId)
		'cSql.Debug
			
			    set rsDup = cSql.Execute
			    if rsDup.eof and rsDup.bof then
			    else
			        existID = rsDup("MEMBERSHIP_ID")
                    existIDLic = rsDUP("LIC_ID")
				    dupCheck = true
				    dupSSNBirthLic = true
			    end if
			end if

			if not dupCheck then
                cSql.SqlStr = "select membership.membership_id, lic.membership_id as lic_id from membership, membership lic " &_
                    "where lic.org_relate_id = to_char(membership.membership_id) " &_
                    "and membership.m_stat_cd = 'ACTIVE' and membership.org_sub_ty_cd = 'MEMBER' " &_
                    "and membership.soc_security = ? and membership.birth_dt = to_date(?,'mm/dd/yyyy') " &_
                    "and lower(membership.last_name) = lower(?) " &_
                    "and membership.membership_id <> ?"
                cSql.AddParam(ssn)
                cSql.AddParam(birthDt)
                cSql.AddParam(lastName)
                cSql.AddParam(membershipId)
		
                'cSql.Debug
			    set rsDup = cSql.Execute
			    if rsDup.eof and rsDup.bof then
			    else
			        existID = rsDup("MEMBERSHIP_ID")
                    existIDLic = rsDUP("LIC_ID")
				    dupCheck = true
				    dupSSNBirthName = true
			    end if
            end if
			if not dupCheck then
                cSql.SqlStr = "select membership.membership_id from membership " &_
                    "where membership.m_stat_cd = 'ACTIVE' and membership.org_sub_ty_cd = 'MEMBER' " &_
                    "and membership.soc_security = ? " &_
                    "and membership.birth_dt = to_date(?,'mm/dd/yyyy')  "  &_
                    "and membership.membership_id <> ?"
                cSql.AddParam(ssn)
                cSql.AddParam(birthDt)
                cSql.AddParam(membershipId)
                'cSql.Debug
			    set rsDup = cSql.Execute
			    if rsDup.eof and rsDup.bof then
			    else
			        existID = rsDup("MEMBERSHIP_ID")
				    dupCheck = true
				    dupSSNBirth = true
			    end if
			end if
	    end if
    end if
end if

'ndbswe
if Application("ORG_ID") = "15" then
    if appRenewType = "Supervision Plan" then
        planNo = getFormVal("planNo")
        if planNo&"" <> "" then
            cSql.SqlStr = "select member.membership_id from membership member " &_
                "where upper(member.ORG_SPEC_ID) = upper(?) " &_
                "and member.m_stat_cd = 'ACTIVE' and member.org_sub_ty_cd = 'PLAN' "
            cSql.AddParam(planNo)
            set rsDup = cSql.Execute
			if rsDup.eof and rsDup.bof then
			else
                childID = getFormVal("CHILD_ID")
			    existID = rsDup("MEMBERSHIP_ID")
				dupCheck = true
			end if
        end if
    end if

    if appRenewType = "Provider Application" then
        providerID = getFormVal("PROVIDERNO")

        if providerID <> "" then
            cSql.SqlStr = "select membership.membership_id from membership, member_detail " &_
                "where member_detail.membership_id = membership.membership_id " &_
                "and member_detail.column_id = (select column_id from org_column_detail where column_desc = 'Provider_Number_PROVIDER') " &_
                "and upper(member_detail.data) = upper(?) " &_
                "and membership.m_stat_cd = 'ACTIVE' and membership.org_sub_ty_cd = 'PROVIDER' "
            cSql.AddParam(ProviderID)
            'cSql.Debug
            set rsDup = cSql.Execute
			if rsDup.eof and rsDup.bof then
			else
			    existID = rsDup("MEMBERSHIP_ID")
				dupCheck = true
			end if
        
        end if
    end if
    if appRenewType = "License Application" then
        ssn = getFormVal("SSN")
        birthDt = getFormVal("DOB")
        lastName = getFormVal("LAST_NAME")
        licType = getFormVal("LIC_TYPE")

        if licType = "Upgrading" then
            cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MEMBER' " &_
                "and soc_security = ? " &_
                "and membership_id <> ?"
            cSql.AddParam(ssn)
            cSql.AddParam(membershipId)
	        set rsDup = cSql.Execute
	        if rsDup.eof and rsDup.bof then
	        else
	            existID = rsDup("MEMBERSHIP_ID")
		        dupCheck = true
		        dupSSNBirthName = true
	        end if
        elseif birthDt <> "" and ssn <> "" and lastName <> "" then
            cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MEMBER' " &_
                "and soc_security = ? and birth_dt = to_date(?,'mm/dd/yyyy') " &_
                "and lower(last_name) = lower(?) " &_
                "and membership_id <> ?"
            cSql.AddParam(ssn)
            cSql.AddParam(birthDt)
            cSql.AddParam(lastName)
            cSql.AddParam(membershipId)
            'cSql.Debug
			set rsDup = cSql.Execute
			if rsDup.eof and rsDup.bof then
			else
			    existID = rsDup("MEMBERSHIP_ID")
				dupCheck = true
				dupSSNBirthName = true
			end if
			if not dupCheck then
                cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                    "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MEMBER' " &_
                    "and soc_security = ? " &_
                    "and lower(last_name) = lower(?) "  &_
                    "and membership_id <> ?"
                cSql.AddParam(ssn)
                cSql.AddParam(lastName)
                cSql.AddParam(membershipId)
                'cSql.Debug
			    set rsDup = cSql.Execute
			    if rsDup.eof and rsDup.bof then
			    else
			        existID = rsDup("MEMBERSHIP_ID")
				    dupCheck = true
				    dupSSNName = true
			    end if
			    if not dupCheck then
                    cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                        "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MEMBER' " &_
                        "and soc_security = ? " &_
                        "and birth_dt = to_date(?,'mm/dd/yyyy')  "  &_
                        "and membership_id <> ?"
                    cSql.AddParam(ssn)
                    cSql.AddParam(birthDt)
                    cSql.AddParam(membershipId)
                    'cSql.Debug
			        set rsDup = cSql.Execute
			        if rsDup.eof and rsDup.bof then
			        else
			            existID = rsDup("MEMBERSHIP_ID")
				        dupCheck = true
				        dupSSNBirth = true
			        end if
			    end if
			end if
	    end if
    end if
end if

'nvot
if Application("ORG_ID") = "1500" then
    if appRenewType = "APP" then
        ssn = getFormVal("ssn")
        lastName = getFormVal("LAST_NAME")

        if ssn <> "" and lastName <> "" then
            cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MEMBER' " &_
                "and soc_security = ? " &_
                "and lower(last_name) = lower(?) " &_
                "and membership_id <> ?"
            cSql.AddParam(ssn)
            cSql.AddParam(lastName)
            cSql.AddParam(membershipId)
            'cSql.Debug
			set rsDup = cSql.Execute
			if rsDup.eof and rsDup.bof then
			else
			    existID = rsDup("MEMBERSHIP_ID")
				dupCheck = true
				dupSSNName = true
			end if
			if not dupCheck then
			    if not dupCheck then
                    cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                        "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MEMBER' " &_
                        "and soc_security = ? " &_
                        "and membership_id <> ?"
                    cSql.AddParam(ssn)
                    cSql.AddParam(membershipId)
                    'cSql.Debug
			        set rsDup = cSql.Execute
			        if rsDup.eof and rsDup.bof then
			        else
			            existID = rsDup("MEMBERSHIP_ID")
				        dupCheck = true
				        dupSSN = true
			        end if
			    end if
			end if
	    end if
    end if
end if

'ndmirt
if Application("ORG_ID") = "1555" then
    if appRenewType = "COND_APP" or appRenewType = "App" or appRenewType = "TEMP_APP" then
        ssn = getFormVal("ssn")
        lastName = getFormVal("LAST_NAME")

        if ssn <> "" and lastName <> "" then
            cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MEMBER' " &_
                "and soc_security = ? " &_
                "and lower(last_name) = lower(?) " &_
                "and membership_id <> ?"
            cSql.AddParam(ssn)
            cSql.AddParam(lastName)
            cSql.AddParam(membershipId)
            'cSql.Debug
			set rsDup = cSql.Execute
			if rsDup.eof and rsDup.bof then
			else
			    existID = rsDup("MEMBERSHIP_ID")
				dupCheck = true
				dupSSNName = true
			end if
			if not dupCheck then
			    if not dupCheck then
                    cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                        "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MEMBER' " &_
                        "and soc_security = ? " &_
                        "and membership_id <> ?"
                    cSql.AddParam(ssn)
                    cSql.AddParam(membershipId)
                    'cSql.Debug
			        set rsDup = cSql.Execute
			        if rsDup.eof and rsDup.bof then
			        else
			            existID = rsDup("MEMBERSHIP_ID")
				        dupCheck = true
				        dupSSN = true
			        end if
			    end if
			end if
	    end if
    end if
end if


'ndotboard
if Application("ORG_ID") = "1400" then
    if appRenewType = "App" then
        ssn = getFormVal("ssn")
        birthDt = getFormVal("dob")
        lastName = getFormVal("LAST_NAME")

        if birthDt <> "" and ssn <> "" and lastName <> "" then
            cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MEMBER' " &_
                "and soc_security = ? and birth_dt = to_date(?,'mm/dd/yyyy') " &_
                "and lower(last_name) = lower(?) " &_
                "and membership_id <> ?"
            cSql.AddParam(ssn)
            cSql.AddParam(birthDt)
            cSql.AddParam(lastName)
            cSql.AddParam(membershipId)
            'cSql.Debug
			set rsDup = cSql.Execute
			if rsDup.eof and rsDup.bof then
			else
			    existID = rsDup("MEMBERSHIP_ID")
				dupCheck = true
				dupSSNBirthName = true
			end if
			if not dupCheck then
			    if not dupCheck then
                    cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                        "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MEMBER' " &_
                        "and soc_security = ? " &_
                        "and birth_dt = to_date(?,'mm/dd/yyyy')  "  &_
                        "and membership_id <> ?"
                    cSql.AddParam(ssn)
                    cSql.AddParam(birthDt)
                    cSql.AddParam(membershipId)
                    'cSql.Debug
			        set rsDup = cSql.Execute
			        if rsDup.eof and rsDup.bof then
			        else
			            existID = rsDup("MEMBERSHIP_ID")
				        dupCheck = true
				        dupSSNBirth = true
			        end if
			    end if
			end if
	    end if
    end if
end if

if Application("ORG_ID") = "725" then
    if appRenewType = "ILDR" then
        ssn = getFormVal("Social_Security_Number")
        lastName = getFormVal("LAST_NAME")

        if ssn <> "" and lastName <> "" then
            cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MD' " &_
                "and soc_security = ? " &_
                "and lower(last_name) = lower(?)"
            cSql.AddParam(ssn)
            cSql.AddParam(lastName)
            'cSql.Debug
			set rsDup = cSql.Execute
			if rsDup.eof and rsDup.bof then
			else
			    existID = rsDup("MEMBERSHIP_ID")
				dupCheck = true
				dupNameSSN = true
			end if
			if not dupCheck then
                cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                    "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'DPM' " &_
                    "and soc_security = ? " 
                cSql.AddParam(ssn)
                'cSql.Debug
			    set rsDup = cSql.Execute
			    if rsDup.eof and rsDup.bof then
			    else
			        existID = rsDup("MEMBERSHIP_ID")
				    dupCheck = true
				    dupSSN = true
			    end if
			end if
	    end if
    end if

    if appRenewType = "DPM_APP" then
        ssn = getFormVal("SSN")
        lastName = getFormVal("LAST_NAME")
        
        if ssn <> "" and lastName <> "" then
            cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'DPM' " &_
                "and soc_security = ? " &_
                "and lower(last_name) = lower(?)"
            cSql.AddParam(ssn)
            cSql.AddParam(lastName)
            'cSql.Debug
			set rsDup = cSql.Execute
			if rsDup.eof and rsDup.bof then
			else
			    existID = rsDup("MEMBERSHIP_ID")
				dupCheck = true
				dupNameSSN = true
			end if
			if not dupCheck then
                cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                    "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'DPM' " &_
                    "and soc_security = ? " 
                cSql.AddParam(ssn)
                'cSql.Debug
			    set rsDup = cSql.Execute
			    if rsDup.eof and rsDup.bof then
			    else
			        existID = rsDup("MEMBERSHIP_ID")
				    dupCheck = true
				    dupSSN = true
			    end if
			end if
	    end if
    end if

    if appRenewType = "PA_APP" then
        ssn = ""
        lastName = ""
        
        ssn = getFormVal("SSN")
        lastName = getFormVal("LAST_NAME")

        if ssn <> "" and lastName <> "" then
            cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'PA' " &_
                "and soc_security = ? " &_
                "and lower(last_name) = lower(?)"
            cSql.AddParam(ssn)
            cSql.AddParam(lastName)
            'cSql.Debug
			set rsDup = cSql.Execute
			if rsDup.eof and rsDup.bof then
			else
			    existID = rsDup("MEMBERSHIP_ID")
				dupCheck = true
				dupNameSSN = true
			end if
			if not dupCheck then
                cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                    "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'PA' " &_
                    "and soc_security = ? " 
                cSql.AddParam(ssn)
                'cSql.Debug
			    set rsDup = cSql.Execute
			    if rsDup.eof and rsDup.bof then
			    else
			        existID = rsDup("MEMBERSHIP_ID")
				    dupCheck = true
				    dupSSN = true
			    end if
			end if
	    end if
    end if

    if appRenewType = "CORP_APP" or appRenewType = "PLLC_APP" then
        coName = getFormVal("COMPANY")

        cSql.SqlStr = "select MEMBERSHIP_ID from membership where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MLMS' and lower(nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = membership.membership_id and column_desc = 'CompanyName_MLMS'),'')) = lower(?)"
        cSql.AddParam(coName)
        set rsDup = cSql.Execute
        if rsDup.eof then
        else
            dupCO = true
            dupCheck = true
        end if
    end if
end if

if Application("ORG_ID") = "800" then
    if appRenewType = "Vet Facility Reg" or appRenewType = "Euth Fac Reg" then
        facName = getFormVal("COMPANY")
        facName = getFormVal("HOMEPHONE")

        dupCheck = false

        if facName <> "" and facPhone <> "" then
            cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'Corporations' " &_
                "and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = membership.membership_id and column_desc = 'Corporate_Phone_Corporations'),'') = ? " &_
                "and lower(nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = membership.membership_id and column_desc = 'Name_of_Corporation_Corporations'),'')) = lower(?)"
            cSql.AddParam(facPhone)
            cSql.AddParam(facName)
            'cSql.Debug
			set rsDup = cSql.Execute
			if rsDup.eof and rsDup.bof then
			else
			    existID = rsDup("MEMBERSHIP_ID")
				dupCheck = true
				dupPhoneName = true
			end if
			if not dupCheck then
                cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                    "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'Corporations' " &_
                    "and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = membership.membership_id and column_desc = 'Corporate_Phone_Corporations'),'') = ? " 
                cSql.AddParam(facPhone)
                'cSql.Debug
			    set rsDup = cSql.Execute
			    if rsDup.eof and rsDup.bof then
			    else
			        existID = rsDup("MEMBERSHIP_ID")
				    dupCheck = true
				    dupPhone = true
			    end if
			end if
	    end if
    end if
    
    if appRenewType = "DVM_Apply" or appRenewType = "RVT_Apply" or appRenewType = "CAET_Apply" then
        ssn = getFormVal("SSN")
        birthDt = getFormVal("DOB")
        lastName = getFormVal("LAST_NAME")

        if birthDt <> "" and ssn <> "" and lastName <> "" then
            cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'RTV' " &_
                "and soc_security = ? and to_date(nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = membership.membership_id and column_desc = 'Lic_Exp_Date_RTV'),''),'mm/dd/yyyy') = to_date(?,'mm/dd/yyyy') " &_
                "and lower(last_name) = lower(?) " &_
                "union " &_
                "select MEMBERSHIP_ID from membership " &_
                "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'BoardXP' " &_
                "and soc_security = ? and to_date(nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = membership.membership_id and column_desc = 'Lic_Exp_Date_BoardXP'),''),'mm/dd/yyyy') = to_date(?,'mm/dd/yyyy') " &_
                "and lower(last_name) = lower(?) " &_
                "union " &_
                "select MEMBERSHIP_ID from membership " &_
                "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MEMBER' " &_
                "and soc_security = ? and LIC_EXPIRE_DT = to_date(?,'mm/dd/yyyy') " &_
                "and lower(last_name) = lower(?) " 
            cSql.AddParam(ssn)
            cSql.AddParam(birthDt)
            cSql.AddParam(lastName)
            cSql.AddParam(ssn)
            cSql.AddParam(birthDt)
            cSql.AddParam(lastName)
            cSql.AddParam(ssn)
            cSql.AddParam(birthDt)
            cSql.AddParam(lastName)
            'cSql.Debug
			set rsDup = cSql.Execute
			if rsDup.eof and rsDup.bof then
			else
			    existID = rsDup("MEMBERSHIP_ID")
				dupCheck = true
				dupSSNBirthName = true
			end if
			if not dupCheck then
                cSql.SqlStr = "select MEMBERSHIP_ID from membership " &_
                    "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'RTV' " &_
                    "and soc_security = ? " &_
                    "and lower(last_name) = lower(?) " &_
                    "union " &_
                    "select MEMBERSHIP_ID from membership " &_
                    "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'BoardXP' " &_
                    "and soc_security = ?  " &_
                    "and lower(last_name) = lower(?) " &_
                    "union " &_
                    "select MEMBERSHIP_ID from membership " &_
                    "where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'MEMBER' " &_
                    "and soc_security = ? " &_
                    "and lower(last_name) = lower(?) " 
                cSql.AddParam(ssn)
                cSql.AddParam(lastName)
                cSql.AddParam(ssn)
                cSql.AddParam(lastName)
                cSql.AddParam(ssn)
                cSql.AddParam(lastName)
                'cSql.Debug
			    set rsDup = cSql.Execute
			    if rsDup.eof and rsDup.bof then
			    else
			        existID = rsDup("MEMBERSHIP_ID")
				    dupCheck = true
				    dupSSNName = true
			    end if
			end if
	    end if
    end if
end if

if Application("ORG_ID") = "775" then
    if appRenewType = "EI_App" or appRenewType = "PE App" then
        ssn = getFormVal("SSN")
        birthDt = getFormVal("DOB")
        if birthDt <> "" and ssn <> "" then
            cSql.SqlStr = "select MEMBERSHIP_ID from membership where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'BoardXP' and soc_security = ? and birth_dt = to_date(?,'mm/dd/yyyy')"
            cSql.AddParam(ssn)
            cSql.AddParam(birthDt)
			set rsDup = cSql.Execute
			if rsDup.eof and rsDup.bof then
			else
			    existID = rsDup("MEMBERSHIP_ID")
				dupCheck = true
				dupSSNBirth = true
			end if
			if not dupCheck then
                cSql.SqlStr = "select MEMBERSHIP_ID from membership where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'BoardXP' and soc_security = ?"
                cSql.AddParam(ssn)
			    set rsDup = cSql.Execute
			    if rsDup.eof and rsDup.bof then
			    else
			        existID = rsDup("MEMBERSHIP_ID")
				    dupCheck = true
				    dupSSN = true
			    end if
			    if not dupCheck then
                    cSql.SqlStr = "select MEMBERSHIP_ID from membership where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'BoardXP' and birth_dt = to_date(?,'mm/dd/yyyy')"
                    cSql.AddParam(birthDt)
			        set rsDup = cSql.Execute
			        if rsDup.eof and rsDup.bof then
			        else
			            existID = rsDup("MEMBERSHIP_ID")
				        dupCheck = true
				        dupBirth = true
			        end if
			    end if
			end if
	    end if
        
    end if
    
	if appRenewType = "COA_App" then
        fein = getFormVal("FEIN")

		if fein <> "" then
			cSql.SqlStr = "select MEMBERSHIP_ID from membership where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'Corporations' and nvl((select data from member_detail, org_column_detail where org_column_detail.column_id = member_detail.column_id and member_detail.membership_id = membership.membership_id and column_desc = 'Fed_Corporation_Number_Corporations'),'') = ? and membership_id <> ?"
			cSql.AddParam(fein)
            cSql.AddParam(membershipId)
			'cSql.Debug
			set rsDup = cSql.Execute
			if rsDup.eof and rsDup.bof then
			else
			    existID = rsDup("MEMBERSHIP_ID")
				dupCheck = true
			end if
		end if
	end if
end if

if Application("ORG_ID") = "375" then
	if appRenewType = "RES_APP" then
		ssn = ""
		if isArray(arrAppRenewDetail) then
			for i = lbound(arrAppRenewDetail,2) to ubound(arrAppRenewDetail,2)
				if lcase(arrAppRenewDetail(0,i)) = "soc_security" or lcase(arrAppRenewDetail(0,i)) = "ssn" then
					ssn = arrAppRenewDetail(1,i)
				end if
			next
		end if
		
		if ssn <> "" then
			cSql.SqlStr = "select MEMBERSHIP_ID from membership where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'Residents' and soc_security = ?"
			cSql.AddParam(ssn)
			set rsDup = cSql.Execute
			if rsDup.eof and rsDup.bof then
			else
				dupCheck = true
			end if
		end if
	end if
	if appRenewType = "DO_APP" then
		ssn = ""
		if isArray(arrAppRenewDetail) then
			for i = lbound(arrAppRenewDetail,2) to ubound(arrAppRenewDetail,2)
				if lcase(arrAppRenewDetail(0,i)) = "soc_security" or lcase(arrAppRenewDetail(0,i)) = "ssn" then
					ssn = arrAppRenewDetail(1,i)
				end if
			next
		end if
		
		if ssn <> "" then
			cSql.SqlStr = "select MEMBERSHIP_ID from membership where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'BoardXP' and soc_security = ?"
			cSql.AddParam(ssn)
			set rsDup = cSql.Execute
			if rsDup.eof and rsDup.bof then
			else
				dupCheck = true
			end if
		end if
	end if

	if appRenewType = "PA_APP" then
		ssn = ""
		if isArray(arrAppRenewDetail) then
			for i = lbound(arrAppRenewDetail,2) to ubound(arrAppRenewDetail,2)
				if lcase(arrAppRenewDetail(0,i)) = "soc_security" or lcase(arrAppRenewDetail(0,i)) = "ssn" then
					ssn = arrAppRenewDetail(1,i)
				end if
			next
		end if
		
		if ssn <> "" then
			cSql.SqlStr = "select MEMBERSHIP_ID from membership where m_stat_cd = 'ACTIVE' and org_sub_ty_cd = 'Physician Assistants' and soc_security = ?"
			cSql.AddParam(ssn)
			set rsDup = cSql.Execute
			if rsDup.eof and rsDup.bof then
			else
				dupCheck = true
			end if
		end if
	end if
end if
%>

<%if Application("ORG_ID") = "725" and appRenewType = "MD RENEW" then%>
<div align="center">
    <a href="<%=scriptName%>?action=rebuildMDRenewReport&app_renew_id=<%=request.querystring("app_renew_id")%>" class="adminLink">Rebuild Renewal Report</a>
</a>
</div>
<br />
<%end if%>

<%if Application("ORG_ID") = "2400" and existIndID&"" <> "" then %>
<table align="center">
<tr><td align="right"><b><font color="green">Duplicate Existing Record</font></b>:</td><td>&nbsp;</td><td><a href="https://<%=request.ServerVariables("SERVER_NAME")%>/db/live/universal.asp?PK=<%=existIndID%>&VIEW=MASTERVIEW&ORG_SUB_TY_CD=MEMBER" target="_blank">Link</a></td></tr>
</table><br />
<%end if %>

<%if Application("ORG_ID") = "1800" and appRenewType = "APP_LMSW" and existIndID&"" <> "" then %>
<table align="center">
<tr><td align="right"><b><font color="red">Internal ID for Record Created by Form</font></b>:</td><td>&nbsp;</td><td><%=membershipID%></td></tr>
<tr><td align="right"><b><font color="blue">Database Record Created by Form</font></b>:</td><td>&nbsp;</td><td><a href="https://<%=request.ServerVariables("SERVER_NAME")%>/db/live/universal.asp?PK=<%=membershipID%>&VIEW=MASTERVIEW&ORG_SUB_TY_CD=<%=recType%>" target="_blank">Link</a></td></tr>
<tr><td align="right"><b><font color="purple">Internal ID for Duplicate Existing Record</font></b>:</td><td>&nbsp;</td><td><%=existIndID%></td></tr>
<tr><td align="right"><b><font color="green">Duplicate Existing Record</font></b>:</td><td>&nbsp;</td><td><a href="https://<%=request.ServerVariables("SERVER_NAME")%>/db/live/universal.asp?PK=<%=existIndID%>&VIEW=MASTERVIEW&ORG_SUB_TY_CD=MEMBER" target="_blank">Link</a></td></tr>
</table><br />
<%end if %>

<%if Application("ORG_ID") = "3900" and right(appRenewType,4) = "_APP" and existID&"" <> "" then %>
<table align="center">
<tr><td align="right"><b><font color="red">Internal ID for Record Created by Form</font></b>:</td><td>&nbsp;</td><td><%=membershipID%></td></tr>
<tr><td align="right"><b><font color="blue">Database Record Created by Form</font></b>:</td><td>&nbsp;</td><td><a href="https://<%=request.ServerVariables("SERVER_NAME")%>/db/live/universal.asp?PK=<%=membershipID%>&VIEW=MASTERVIEW&ORG_SUB_TY_CD=<%=recType%>" target="_blank">Link</a></td></tr>
<tr><td align="right"><b><font color="purple">Internal ID for Duplicate Existing Record</font></b>:</td><td>&nbsp;</td><td><%=existID%></td></tr>
<tr><td align="right"><b><font color="green">Duplicate Existing Record</font></b>:</td><td>&nbsp;</td><td><a href="https://<%=request.ServerVariables("SERVER_NAME")%>/db/live/universal.asp?PK=<%=existID%>&VIEW=MASTERVIEW&ORG_SUB_TY_CD=EMP" target="_blank">Link</a></td></tr>
</table><br />
<%end if %>

<%if appRenewStatus = "PENDING" then%>
<div align="center">
    <%if dupCheck then%>
        <%if Application("ORG_ID") = "3900" and right(appRenewType,4) = "_APP" then%>
            Duplicate records exist.<br />
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Insert New Record</a> or 
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>&exist=<%=existID%>" class="adminLink">Update Existing Record</a>
        <%elseif Application("ORG_ID") = "775" and appRenewType = "COA_App" then%>
            Duplicate records exist.<br />
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Insert New Record</a> or 
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>&exist=<%=existID%>" class="adminLink">Update Existing Record</a>
        <%elseif Application("ORG_ID") = "2900" then%>
            <%if dupCheck then%>
            A duplicate records exist. Do you want to proceed?<br />
	        <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Insert New Record</a> or 
    	    <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>&exist=<%=existIndID%>" class="adminLink">Update Existing Record</a>
            <%else%>
	        <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Approve</a>
            <%end if%>
        <%elseif Application("ORG_ID") = "2400" then%>
            <%if dupCheck and appRenewType <> "BRANCH_APP" then%>
            Duplicate records exist. Do you want to proceed?<br />
	        <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Insert New Record</a> or 
    	    <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>&exist=<%=existIndID%>" class="adminLink">Update Existing Record</a>
            <%else%>
	        <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Approve</a>
            <%end if%>
        <%elseif Application("ORG_ID") = "4100" and (appRenewType = "APP_VET" or appRenewType = "APP_TECH") then %>
            <%if dupCheck then%>
            Duplicate records exist. Do you want to proceed?<br />
	        <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Insert New Record</a> or 
    	    <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>&existInd=<%=existIndID%>" class="adminLink">Update Existing Record</a>
            <%else%>
	        <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Approve</a>
            <%end if%>
        <%elseif Application("ORG_ID") = "3200" and (appRenewType = "APP" or appRenewType = "APP_RECIP" or appRenewType = "INTERN_APP") then %>
            <%if dupCheck then%>
            Duplicate records exist. Do you want to proceed?<br />
	        <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Insert New Record</a> or 
    	    <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>&existInd=<%=existIndID%>" class="adminLink">Update Existing Record</a>
            <%else%>
	        <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Approve</a>
            <%end if%>
        <%elseif Application("ORG_ID") = "3100" and appRenewType = "APP" then %>
            <%if dupCheck then%>
            Duplicate records exist. Do you want to proceed?<br />
	        <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Insert New Record</a> or 
    	    <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>&existInd=<%=existIndID%>&existLic=<%=existLicID%>" class="adminLink">Update Existing Record</a>
            <%else%>
	        <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Approve</a>
            <%end if%>
        <%elseif Application("ORG_ID") = "1800" and appRenewType = "APP" then %>
            <%if dupCheck then%>
            Duplicate records exist. Do you want to proceed?<br />
	        <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Insert New Record</a> or 
    	    <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>&existInd=<%=existIndID%>&existLic=<%=existLicID%>" class="adminLink">Update Existing Record</a>
            <%else%>
	        <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Approve</a>
            <%end if%>
        <%elseif Application("ORG_ID") = "1800" and appRenewType = "APP_LMSW" then %>
            <%if dupCheck then%>
            Duplicate records exist. Do you want to proceed?<br />
	        <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Insert New Record</a> or 
    	    <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>&existInd=<%=existIndID%>" class="adminLink">Update Existing Record</a>
            <%else%>
	        <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Approve</a>
            <%end if%>
        <%elseif Application("ORG_ID") = "1800" and appRenewType = "APP_LSW" then %>
            <%if dupCheck then%>
            Duplicate records exist. Do you want to proceed?<br />
	        <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Insert New Record</a> or 
    	    <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>&existInd=<%=existIndID%>" class="adminLink">Update Existing Record</a>
            <%else%>
	        <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Approve</a>
            <%end if%>
        <%elseif Application("ORG_ID") = "1900" and (appRenewType = "NA_RECRUIT" or appRenewType = "EMG_TRAIN" or appRenewType = "MA_APP" or appRenewType = "MA2_APP" or appRenewType = "NA_APP" or appRenewType = "CNA_APP" or appRenewType = "HHA_APP")  then%>
            <%if dupCheck then%>
            Duplicate records exist. Do you want to proceed?<br />
	        <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Insert New Record</a> or 
    	    <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>&existInd=<%=existIndID%>&existLic=<%=existLicID%>" class="adminLink">Update Existing Record</a>
            <%else%>
	        <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Approve</a>
            <%end if%>
        <%elseif Application("ORG_ID") = "1450" and (appRenewType = "COVID-19_APP" or appRenewType = "NT_APP" or appRenewType = "TT_APP" or appRenewType = "IN_APP" or appRenewType = "RP_APP") then%>
            <%if dupCheck then%>
            Duplicate records exist. Do you want to proceed?<br />
	        <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Insert New Record</a> or 
    	    <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>&existInd=<%=existIndID%>&existLic=<%=existLicID%>" class="adminLink">Update Existing Record</a>
            <%else%>
	        <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Approve</a>
            <%end if%>
        <%elseif Application("ORG_ID") = "1450" and (appRenewType = "CCP_APP" or appRenewType = "MOP_APP" or appRenewType = "WDD_APP" or appRenewType = "TPLP_APP" or appRenewType = "MFR_APP" or appRenewType = "LP_APP" or appRenewType = "PHAR_APP" or appRenewType = "CSP_APP") then%>
            <%'if dupCheck then%>
            <!--
            Duplicate records exist. Do you want to proceed?<br />
	        <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Insert New Record</a> or 
    	    <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>&exist=<%=existID%>&existFac=<%=existFacID%>" class="adminLink">Update Existing Record</a>
            -->
            <%'else%>
	        <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Approve</a>
            <%'end if%>
        <%elseif Application("ORG_ID") = "725" and appRenewType = "PA_APP" and dupCheck then%>
            <%if dupNameSSN then%>
            A duplicate record exists.<br />
            <%end if%>
            <%if dupSSN then%>
            A record with a duplicate SSN exists.<br />
            <%end if%>
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Insert New Record</a> or 
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>&exist=<%=existID%>" class="adminLink">Update Existing Record</a>
        <%elseif Application("ORG_ID") = "725" and (appRenewType = "ILDR" or appRenewType = "DPM_APP") and dupCheck then%>
            <%if dupNameSSN then%>
            A duplicate record exists.<br />
            <%end if%>
            <%if dupSSN then%>
            A record with a duplicate SSN exists.<br />
            <%end if%>
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Insert New Record</a> or 
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>&exist=<%=existID%>" class="adminLink">Update Existing Record</a>
        <%elseif Application("ORG_ID") = "800" and (appRenewType = "Vet Facility Reg" or appRenewType = "Euth Fac Reg") and dupCheck then%>
            <%if dupPhoneName then%>
            A duplicate record exists. Do you want to proceed?<br />
            <%end if%>
            <%if dupPhone then%>
            A record with duplicate phone number exists. Do you want to proceed?<br />
            <%end if%>
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Insert New Record</a> or 
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>&exist=<%=existID%>" class="adminLink">Update Existing Record</a>
        <%elseif Application("ORG_ID") = "15" and (appRenewType = "Provider Application" or appRenewType = "Supervision Plan") and dupCheck then%>
            A record with a matching provider number exists. Do you want to proceed?<br />
            <a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Insert New Record</a> or 
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>&exist=<%=existID%>&new_id=<%=childID%>" class="adminLink">Update Existing Record</a>
        <%elseif Application("ORG_ID") = "15" and appRenewType = "License Application" and dupCheck then%>
            <%if dupSSNBirthName then%>
            A duplicate record exists. Do you want to proceed?<br />
            <%end if%>
            <%if dupSSNName then%>
            A record with a duplicate SSN & Last Name exists. Do you want to proceed?<br />
            <%end if%>
            <%if dupSSNBirth then%>
            A record with a duplicate SSN & DOB exists. Do you want to proceed?<br />
            <%end if%>
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Insert New Record</a> or 
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>&exist=<%=existID%>" class="adminLink">Update Existing Record</a>
        <%elseif Application("ORG_ID") = "1500" and appRenewType = "APP" and dupCheck then%>
            <%if dupSSNName then%>
            Duplicate records exist. Do you want to create a new License Record?<br />
            <%end if%>
            <%if dupSSN then%>
            The Last Name does NOT match. Do you want to update the [MEMBER] Record and create a new License Record?<br />
            <%end if%>
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Insert New Record</a> or 
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>&exist=<%=existID%>" class="adminLink">Update Existing Record</a>
        <%elseif Application("ORG_ID") = "1555" and (appRenewType = "COND_APP" or appRenewType = "App" or appRenewType = "TEMP_APP") and dupCheck then%>
            <%if dupSSNName then%>
            Duplicate records exist. Do you want to create a new License Record?<br />
            <%end if%>
            <%if dupSSN then%>
            The Last Name does NOT match. Do you want to update the [MEMBER] Record and create a new License Record?<br />
            <%end if%>
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Insert New Record</a> or 
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>&exist=<%=existID%>" class="adminLink">Update Existing Record</a>
        <%elseif Application("ORG_ID") = "1400" and appRenewType = "App" and dupCheck then%>
            <%if dupSSNBirthName then%>
            A duplicate record exists. Do you want to proceed?<br />
            <%end if%>
            <%if dupSSNBirth then%>
            A record with a duplicate SSN & DOB exists. Do you want to proceed?<br />
            <%end if%>
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Insert New Record</a> or 
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>&exist=<%=existID%>" class="adminLink">Update Existing Record</a>
        <%elseif Application("ORG_ID") = "750" and (appRenewType = "CAAN" or appRenewType = "CAPN" or appRenewType = "UDAAPP" or appRenewType = "UMAAPP") and dupCheck then%>
            <%if appRenewType = "CAPN"or appRenewType = "CAAN" then%>
            A duplicate record exists. Do you want to proceed?<br />
            <%end if%>
            <%if dupSSNBirthNameLic then%>
            A duplicate record exists. Do you want to proceed?<br />
            <%end if%>
            <%if dupSSNBirthLic then%>
            A duplicate record exists but the Last Name does NOT match. Do you want to proceed?<br />
            <%end if%>
            <%if dupSSNBirthName then%>
            A duplicate Contact record exists. Do you want to proceed?<br />
            <%end if%>
            <%if dupSSNBirth then%>
            A duplicate Contact record exists but the Last Name does NOT match. Do you want to proceed?<br />
            <%end if%>
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Insert New Record</a> or 
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>&exist=<%=existID%>&existLic=<%=existIDLic%>" class="adminLink">Update Existing Record</a>
        <%elseif Application("ORG_ID") = "1300" and (appRenewType = "admissionevaloffical" or appRenewType = "TRNFREXAMAPP" or appRenewType = "EXAMAPP" or appRenewType = "RECIPCERTWV") and dupCheck then%>
            <%if dupSSNBirthName or dupSSN then%>
            A duplicate record exists. Do you want to proceed?<br />
            <%end if%>
            <%if dupSSNBirth then%>
            A record with a duplicate SSN & DOB exists. Do you want to proceed?<br />
            <%end if%>
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Insert New Record</a> or 
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>&exist=<%=existID%>" class="adminLink">Update Existing Record</a>
        <%elseif Application("ORG_ID") = "900" and (appRenewType = "License Application" or appRenewType = "Provisional Permit Application") and dupCheck then%>
            <%if dupSSNBirthName then%>
            A duplicate record exists. Do you want to proceed?<br />
            <%end if%>
            <%if dupSSNBirth then%>
            A record with a duplicate SSN & DOB exists. Do you want to proceed?<br />
            <%end if%>
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Insert New Record</a> or 
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>&exist=<%=existID%>" class="adminLink">Update Existing Record</a>
        <%elseif Application("ORG_ID") = "800" and (appRenewType = "DVM_Apply" or appRenewType = "RVT_Apply" or appRenewType = "CAET_Apply") and dupCheck then%>
            <%if dupSSNBirthName then%>
            A duplicate record exists. Do you want to proceed?<br />
            <%end if%>
            <%if dupSSNName then%>
            A record with a duplicate SSN & Last Name exists. Do you want to proceed?<br />
            <%end if%>
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Insert New Record</a> or 
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>&exist=<%=existID%>" class="adminLink">Update Existing Record</a>
        <%elseif Application("ORG_ID") = "775" and (appRenewType = "EI_App" or appRenewType = "PE App") then%>
            <%if dupSSNBirth then%>
            A duplicate record exists. Do you want to proceed?<br />
            <%end if%>
            <%if dupSSN then%>
            A record with a duplicate SSN exists. Do you want to proceed?<br />
            <%end if%>
            <%if dupBirth then%>
            A record with a duplicate DOB exists. Do you want to proceed?<br />
            <%end if%>
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Insert New Record</a> or 
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>&exist=<%=existID%>" class="adminLink">Update Existing Record</a>
        <%elseif Application("ORG_ID") = "725" and (appRenewType = "CORP_APP" or appRenewType = "PLLC_APP") then%>
            There is a record with a matching Company Name. Please check the database and notify the applicant if necessary.
        <%else%>
            <%if Application("ORG_ID") = "725" and (appRenewType = "CORP_APP" or appRenewType = "PLLC_APP") then%>
            There is a record with a matching Company Name. Please check the database and notify the applicant if necessary.<br /><br />
            <%end if%>
	<a href="javascript:dup();" class="adminLink">Approve</a>
	    <%end if%>
    <%elseif Application("ORG_ID") = "1450" and appRenewType = "COMPL_BIZ" then%>
	<a href="<%=scriptName%>?action=approve&type=COMPL_INDV&id=<%=request.querystring("app_renew_id")%>" class="adminLink" onclick="hideLink(this);">Approve Complaint Individual</a><br /><br />
	<a href="<%=scriptName%>?action=approve&type=COMPL&id=<%=request.querystring("app_renew_id")%>" class="adminLink" onclick="hideLink(this);">Approve Complaint Facility</a><br /><br />
	<a href="<%=scriptName%>?action=approve&type=BOTH&id=<%=request.querystring("app_renew_id")%>" class="adminLink" onclick="hideLink(this);">Approve Both Individual and Facility</a>
    <%elseif Application("ORG_ID") = "15" and appRenewType = "License Application" and licType = "Upgrading" then %>
    This is a license upgrade and a duplicate record does not exist. Do you want to proceed?
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">OK</a> or 
	<a href="<%=scriptName%>" class="adminLink">Cancel</a>
    <%else%>
        <%if getFormVal("EXISTING_RECORD_ID")&"" = "1" then%>
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Approve and Update Existing Record</a> or 
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>&new=1" class="adminLink">Approve and Create New Record</a>
        <%else%>
	<a href="<%=scriptName%>?action=approve&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Approve</a>
        <%end if%>
    <%end if%>
</div>
<%end if%>

<%if appRenewStatus <> "PENDING" then%>
<br>
<div align="center"><a href="<%=scriptName%>?action=pend&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Pend</a></div>
<%end if%>

<%if appRenewStatus <> "DENIED" then%>
<br>
<div align="center"><a href="<%=scriptName%>?action=deny&id=<%=request.querystring("app_renew_id")%><%
if appRenewType = "EMG_TRAIN" or appRenewType = "NA_RECRUIT" then
    response.write "&existInd=" & existIndID
    response.write "&existLic=" & existLicID
end if
%>" class="adminLink">Deny</a></div>
<%end if%>
<script language="Javascript">
    var logFile = "";
    var pdfFile = "";
    var ajaxFile = "";
    function showDetails(elName) {
        var el = document.getElementById(elName);
        if (el.style.display == "none") {
            if (el.innerHTML.indexOf("Fetching data...") >= 0) {
                if (elName == "tblDetailsFull") {
                    $.get("/admin/admin/ajax.asp?action=showFormData&app_renew_id=<%=request.querystring("app_renew_id")%>&app_renew_type=<%=appRenewType%>&type=full",
                {
                },
                function (data, status) {
                    if (status == "success") {
                        $("#tblDetailsFull").html(data);
                        dataFetched = true;
                        $('.edit').editable('<%=Application("ADMIN_URL")%>/admin/ajax.asp?action=updateAppRenewDetail', {indicator: 'Saving...',onblur: 'submit',tooltip: 'Click to edit...',cssclass: 'editing',width: 'auto',height: '20px',type: 'text'});
                    } else {
                        alert("showData failed");
                    }
                });
            } else if (elName == "tblDetails") {
                $.get("/admin/admin/ajax.asp?action=showFormData&app_renew_id=<%=request.querystring("app_renew_id")%>&app_renew_type=<%=appRenewType%>",
            {
            },
            function (data, status) {
                if (status == "success") {
                    $("#tblDetails").html(data);
                    dataFetched = true;
                } else {
                    alert("showData failed");
                }
            });
        } else if (elName == "tblSubLog") {
            var pageSubID = elName.replace("tblSub","");
            $.post("/admin/admin/ajax.asp",
            {
                action: "showLogFile",
                log_file: logFile,
            },
            function (data, status) {
                if (status == "success") {
                    $("#" + elName).html(data);
                    dataFetched = true;
                } else {
                    alert("showLog failed");
                }
            });
        } else if (elName == "tblSubAjaxLog") {
            var pageSubID = elName.replace("tblSubAjax", "");
            $.post("/admin/admin/ajax.asp",
            {
                action: "showLogFile",
                log_file: ajaxFile,
            },
            function (data, status) {
                if (status == "success") {
                    $("#" + elName).html(data);
                    dataFetched = true;
                } else {
                    alert("showLog failed");
                }
            });
        } else if (elName == "tblSubPDFLog") {
            var pageSubID = elName.replace("tblSub","");
            $.get("/admin/admin/ajax.asp",
            {
                action: "showLogFile",
                log_file: pdfFile,
            },
            function (data, status) {
                if (status == "success") {
                    $("#" + elName).html(data);
                    dataFetched = true;
                } else {
                    alert("showPDFLog failed");
                }
            });
        } else if (elName == "tblSubRedis") {
            $.get("/admin/admin/ajax.asp",
            {
                action: "showSession",
                id:  $("#redis_id").val(),
            },
            function (data, status) {
                if (status == "success") {
                    $("#" + elName).html(data);
                    dataFetched = true;
                } else {
                    alert("showSession failed");
                }
            });
        } else {
            var pageSubID = elName.replace("tblSub","");
            $.get("/admin/admin/ajax.asp?action=showPageSubmit&page_submit_id=" + pageSubID,
            {
            },
            function (data, status) {
                if (status == "success") {
                    $("#" + elName).html(data);
                    dataFetched = true;
                } else {
                    alert("showPageSubmit failed");
                }
            });
        }
    }
    el.style.display = "";
    } else {
        el.style.display = "none";
    }
    }    
    <%if officeIP then%>

    function ardDelete(ardID) {
        confirmMsg = "Are you sure you want to permanently delete this detail entry?  This cannot be undone."

        if (confirm(confirmMsg)) {
            $.get("/admin/admin/ajax.asp?action=deleteAppRenewDetail&app_renew_detail_id=" + ardID,
            {
            },
            function (data, status) {
                if (status == "success") {
                    $("#tr" + ardID).hide();
                } else {
                    alert("ardDelete failed");
                }
            });
        }
    }

    function saveNewValue() {
        var newKey = document.getElementById("new_key").value;
        var newValue = document.getElementById("new_value").value;
        if (newKey.length == 0) {
            alert("Key required");
            return;
        }

        $.post("/admin/admin/ajax.asp",
        {
            action: "addNewDetail",
            id: "<%=request.querystring("app_renew_id")%>",
            key: newKey,
            value: newValue,
        },
        function (data, status) {
            if (status == "success") {
                $("#tblDetails").hide();
                $("#tblDetails").html("<tr><td>Fetching data...</td></tr>");
                showDetails("tblDetails")
            } else {
                alert("addNewDetail failed");
            }
        });
    }
    <%end if%>
</script>
<%if officeIP then%>
<script src="js/jquery.jeditable.js" type="text/javascript" charset="utf-8"></script>
<script>
    $.fn.editable.defaults = {
        name: 'value',
        id: 'id',
        type: 'text',
        width: 'auto',
        height: 'auto',
        event: 'click.editable',
        onSAFEblur: 'cancel',
        loadtype: 'POST',
        loadtext: 'Loading...',
        placeholder: '<span class=editPlaceholder>edit</span>',
        loaddata: {},
        submitdata: {},
        ajaxoptions: {}
    };
</script>
<%end if%>
<br>
<div align="center">
    <%if Application("ORG_ID") = "375" and (appRenewType = "CSL_RENEW" or appRenewType = "CORP_APP" or appRenewType = "CORP_RENEW" or appRenewType = "DO_APP" or appRenewType = "PA_APP" or appRenewType = "PLLC_APP" or appRenewType = "PLLC_RENEW") then%>
    <a href="/licensing/pdf_rebuild.asp?id=<%=request.querystring("app_renew_id")%>" class="adminLink">Rebuild PDF</a>
    <%else%>
    <a href="<%=scriptName%>?action=rebuildPDF&id=<%=request.querystring("app_renew_id")%>" class="adminLink">Rebuild PDF</a>
    <%end if%>
</div>
<Br />
<div align="center"><a href="javascript:showDetails('tblDetails');" class="adminLink">Show Data - Minimized</a></div>
<table border="1" style="display:none" align="center" id="tblDetails">
		<tr><td>Fetching data...</td></tr>
</table>
<Br />
<div align="center"><a href="javascript:showDetails('tblDetailsFull');" class="adminLink">Show Data - Full</a></div>
<table border="1" style="display:none" align="center" id="tblDetailsFull">
		<tr><td>Fetching data...</td></tr>
</table>
<%
showLogs = false

'office
if officeIP then
    showLogs = true
end if

if showLogs then
    dim rsPSD, rsPS, pageSubID

    set cSql = new cRunSql
    cSql.conn = connOracle
    cSql.SqlStr = "select APP_RENEW_PAGE_SUBMIT_ID,CREATE_DT from app_renew_page_submit where app_renew_id = ? order by app_renew_page_submit_id"
    cSql.AddParam(request.querystring("app_renew_id"))
    set rsPS = cSql.Execute
    set cSql = nothing
    if rsPS.eof then
    else
    %>
    <br><div align="center">Page Submits</div>
    <%  
        set cCV = new cCodeValueO
        cCV.LoadFromCodeClassAndValue "APP_RENEW_TYPE_CD",appRenewType
        do while not rsPS.eof
            pageSubID = rsPS("APP_RENEW_PAGE_SUBMIT_ID")
            set cSql = new cRunSql
            cSql.Conn = connOracle
            cSql.SqlStr = "select DATA from app_renew_page_sub_detail where app_renew_page_submit_id = ? and field_name = 'SCRIPT_NAME' "
            cSql.AddParam(rsPS("APP_RENEW_PAGE_SUBMIT_ID"))
            set rsPSD = cSql.Execute
            set cSql = nothing
            if rsPSD.eof then
                set cSql = new cRunSql
                cSql.Conn = connOracle
                cSql.SqlStr = "select DATA from app_renew_page_sub_detail where app_renew_page_submit_id = ? and field_name = 'page_cd' "
                cSql.AddParam(rsPS("APP_RENEW_PAGE_SUBMIT_ID"))
                set rsPSD = cSql.Execute
                set cSql = nothing
                if rsPSD.eof then
                    pageName = ""
                else
                    if right(cCV.CodeValueLongDesc,1) = "/" then
                        pageName = cCV.CodeValueLongDesc & rsPSD("DATA")
                    else
                        pageName = cCV.CodeValueLongDesc & "/" & rsPSD("DATA")
                    end if
                end if
            else
                pageName = rsPSD("DATA")
            end if
            rsPSD.close
            set rsPSD = nothing
    %>
    <div align="center"><a href="javascript:showDetails('tblSub<%=rsPS("APP_RENEW_PAGE_SUBMIT_ID")%>');" class="adminLink"><%=rsPS("CREATE_DT")%> - <%=pageName%></a></div>
    <div align="center">Copy Form Values to Show Data - <a href="<%=request.servervariables("SCRIPT_NAME")%>?action=subFormCopy&page_sub_id=<%=rsPS("APP_RENEW_PAGE_SUBMIT_ID")%>&app_renew_id=<%=request.querystring("app_renew_id")%>" classe="adminLink">Go</a></div>

    <table align="center" style="display:none" id="tblSub<%=rsPS("APP_RENEW_PAGE_SUBMIT_ID")%>">
        <tr><td>Fetching data...</td></tr>
    </table><br />
    <%
            response.flush
            rsPS.MoveNext
        loop
        set cCV = nothing
                %>
        <br /><div align="center" class="adminText">
                Copy Last Submit to Show Data - <a href="<%=request.servervariables("SCRIPT_NAME")%>?action=lastSub&page_sub_id=<%=pageSubID%>&app_renew_id=<%=request.querystring("app_renew_id")%>" classe="adminLink">Go</a>
        </div><br />
                <%

    end if
rsPS.close
set rsPS = nothing
set cSql = nothing

set cSql = new cRunSql
cSql.conn = connOracle
cSql.SqlStr = "select DATA from app_renew_detail where app_renew_id = ? and lower(field_name) = 'redissessionid'"
cSql.AddParam(request.querystring("app_renew_id"))
set rsPS = cSql.Execute
if rsPS.eof then
else
    redisID = rsPS("DATA")
%>
<input type="hidden" name="redis_id" id="redis_id" value="<%=redisID%>" />
<br /><div align="center"><a href="javascript:showDetails('tblSubRedis');" class="adminLink">Session</a></div>
<table align="center" style="display:none" id="tblSubRedis">
    <tr><td>Fetching data...</td></tr>
</table>
<%
end if
rsPS.close
set rsPS = nothing

arrPath = split(request.servervariables("APPL_PHYSICAL_PATH")&"","\")

dim cLogIt
set cLogIt = new cLog

logFile = "forms\form_" & request.querystring("app_renew_id") & ".log"
if cLogIt.LogExists(logFile) then
    %>
    <script>logFile = "<%=replace(logFile,"\","\\")%>";</script>
    <br /><div align="center"><a href="javascript:showDetails('tblSubLog');" class="adminLink">Form Log</a></div>
    <table align="center" style="display:none" id="tblSubLog">
        <tr><td>Fetching data...</td></tr>
    </table>
    <%
end if

ajaxFile = "forms\ajax_" & request.querystring("app_renew_id") & ".log"
if cLogIt.LogExists(ajaxFile) then
    %>
    <script>ajaxFile = "<%=replace(ajaxFile,"\","\\")%>";</script>
    <br /><div align="center"><a href="javascript:showDetails('tblSubAjaxLog');" class="adminLink">Ajax Log</a></div>
    <table align="center" style="display:none" id="tblSubAjaxLog">
        <tr><td>Fetching data...</td></tr>
    </table>
    <%
end if

pdfFile = "forms\pdf_" & request.querystring("app_renew_id") & ".log"
if cLogIt.LogExists(pdfFile) then
    %>
    <script>pdfFile = "<%=replace(pdfFile,"\","\\")%>";</script>
    <br /><div align="center"><a href="javascript:showDetails('tblSubPDFLog');" class="adminLink">PDF Log</a></div>
    <table align="center" style="display:none" id="tblSubPDFLog">
        <tr><td>Fetching data...</td></tr>
    </table>
    <%
end if
end if

if request.QueryString("search")&"" <> "" then
%>
<script>
    jQuery.expr[':'].contains = function(a, i, m) {
        return jQuery(a).text().toUpperCase().indexOf(m[3].toUpperCase()) >= 0;
    };
    $("td:contains(<%=request.QueryString("search")%>)").closest("table").prev("div").css("background-color", "rgba(255,255,200,.5)");
</script>
<%end if%>
<%end sub%>

<%function getFormVal(field)%>
<%
dim formID, cSqlFV

if request.querystring("app_renew_id")&"" <> "" then
    formID = request.querystring("app_renew_id")
elseif request.QueryString("id")&"" <> "" then
    formID = request.QueryString("id")
end if
set cSqlFV = new cRunSql
cSqlFV.Conn = connOracle
cSqlFV.SqlStr = "select DATA from app_renew_detail where app_renew_id = ? and upper(field_name) = upper(?)"
cSqlFV.AddParam(formID)
cSqlFV.AddParam(field)
logIt "getFormVal sql = " & cSqlFV.SqlStr
logIt "getFormVal formID = " & formID
logIt "getFormVal field = " & field
set rsEmpty = cSqlFV.Execute
set cSqlFV = nothing
if rsEmpty.eof and rsEmpty.bof then
    logIt "getFormVal eof"
    getFormVal = ""
else
    logIt "getFormVal = " & rsEmpty("DATA")
    getFormVal = rsEmpty("DATA")
end if 
%>
<%end function%>

<%function getFormValByID(id,field)%>
<%
dim cSqlVal
set cSqlVal = new cRunSql
cSqlVal.conn = connOracle
cSqlVal.SqlStr = "select DATA from app_renew_detail where app_renew_id = ? and upper(field_name) = upper(?)"
cSqlVal.AddParam(id)
cSqlVal.AddParam(field)
set rsEmpty = cSqlVal.Execute
if rsEmpty.eof and rsEmpty.bof then
   getFormValByID = ""
else
    getFormValByID = rsEmpty("DATA")
end if 
set cSqlVal = nothing
%>
<%end function%>

<%Sub approveUpdate(app_renew_id)%>
<%
dim PARENT_ID, cParent, parentType, recordType, CHILD_ID, childType, cChild
dim GRANDCHILD_ID, grandchildType, cGrandchild
dim cNewMem, newID

set cSql = new cRunSql
cSql.conn = connOracle

set cMember = new cMemberValuesO

cSql.SqlStr = "select MEMBERSHIP_ID,APP_RENEW_TYPE from app_renew where app_renew_id = ?"
cSql.AddParam(app_renew_id)
set rs = cSql.Execute

MEMBERSHIP_ID = rs("MEMBERSHIP_ID")

cMember.LoadFromMembershipId MEMBERSHIP_ID

if request("new") = "1" then
    set cNewMem = new cNewMemberO
    newID = cNewMem.MembershipID
    cNewMem.setValue "M_STAT_CD", "ACTIVE"
    cNewMem.setValue "ORG_SUB_TY_CD",cMember.Value("ORG_SUB_TY_CD")
    cNewMem.InsertRecord
    set cNewMem = nothing
    set cMember = nothing

    set cMember = new cMemberValuesO
    cMember.LoadFromMembershipID newID

    cSql.SqlStr = "update app_renew set MEMBERSHIP_ID = ? where app_renew_id = ?"
    cSql.AddParam(newID)
    cSql.AddParam(app_renew_id)
    cSql.Execute

    cSql.SqlStr = "update app_renew_detail set data = ? where app_renew_id = ? and lower(field_name) = 'membership_id'"
    cSql.AddParam(newID)
    cSql.AddParam(app_renew_id)
    cSql.Execute

    cSql.SqlStr = "delete from app_renew_detail where app_renew_id = ? and lower(field_name) = 'existing_record_id'"
    cSql.AddParam(app_renew_id)
    cSql.Execute

    MEMBERSHIP_ID = newID
end if

cMember.LoadFromMembershipId MEMBERSHIP_ID

appRenewType = rs("APP_RENEW_TYPE")

cSql.SqlStr = "select data from app_renew_detail where app_renew_id = ? and upper(field_name) = 'PARENT_ID'"
cSql.AddParam(app_renew_id)
set rs = cSql.Execute
if rs.eof then
else
    if instr(rs("DATA"),",") then
    else
        PARENT_ID = rs("DATA")
        set cParent = new cMemberValuesO
        set cCV = new cCodeValueO
        cCV.LoadFromCodeClassAndValue "APP_RENEW_TYPE_CD",appRenewType
        recordType = cCV.MinorCode
        cCV.LoadFromCodeClassAndValue "ORG_SUB_TY_CD",recordType
        parentType = cCV.MinorCode
        set cCV = nothing
    end if
end if

if PARENT_ID&"" = "" then
    if cMember.Value("ORG_RELATE_ID")&"" <> "" then
        if instr(cMember.Value("ORG_RELATE_ID"),",") then
        else
            PARENT_ID = cMember.Value("ORG_RELATE_ID")
            set cParent = new cMemberValuesO
            set cCV = new cCodeValueO
            cCV.LoadFromCodeClassAndValue "APP_RENEW_TYPE_CD",appRenewType
            recordType = cCV.MinorCode
            cCV.LoadFromCodeClassAndValue "ORG_SUB_TY_CD",recordType
            parentType = cCV.MinorCode
            set cCV = nothing
        end if
    end if
end if

cSql.SqlStr = "select data from app_renew_detail where app_renew_id = ? and upper(field_name) = 'CHILD_ID'"
cSql.AddParam(app_renew_id)
set rs = cSql.Execute
if rs.eof then
else
    CHILD_ID = rs("DATA")
    set cChild = new cMemberValuesO
    cSql.SqlStr = "select data from app_renew_detail where app_renew_id = ? and upper(field_name) = 'CHILD_TYPE'"
    cSql.AddParam(app_renew_id)
    set rs = cSql.Execute
    if rs.eof then
    else
        childType = rs("DATA")
        cSql.SqlStr = "select data from app_renew_detail where app_renew_id = ? and upper(field_name) = 'GRANDCHILD_ID'"
        cSql.AddParam(app_renew_id)
        set rs = cSql.Execute
        if rs.eof then
        else
            GRANDCHILD_ID = rs("DATA")
            set cGrandchild = new cMemberValuesO
            cSql.SqlStr = "select data from app_renew_detail where app_renew_id = ? and upper(field_name) = 'GRANDCHILD_TYPE'"
            cSql.AddParam(app_renew_id)
            set rs = cSql.Execute
            if rs.eof then
            else
                grandchildType = rs("DATA")
            end if
        end if
    end if
end if

cSql.SqlStr = "select approve_app,approve_update from app_renew_config where app_renew_type = ?"
cSql.AddParam(appRenewType)
set rs = cSql.Execute

if not rs.eof then
    if rs("approve_app")&"" = "Y" then
        cMember.updateValue MEMBERSHIP_ID,"M_STAT_CD","ACTIVE"

        set cSql = nothing

        'child records
        set cSql = new cRunSql
        cSql.Conn = connOracle
        cSql.SqlStr = "select membership_id from membership where org_relate_id = ? and m_stat_cd = 'PENDING'"
        cSql.AddParam(MEMBERSHIP_ID)
        set rsPend = cSql.Execute
        set cSql = nothing
        do while not rsPend.eof
            set cSql = new cRunSql
            cSql.Conn = connOracle
            cSql.SqlStr = "update membership set m_stat_cd = 'ACTIVE' where membership_id = ?"
            cSql.AddParam(rsPend("membership_id"))
            cSql.Execute
            set cSql = nothing
            rsPend.MoveNext
        loop
        rsPend.close
        set rsPend = nothing
        'end child records

        'parent record
        set cSql = new cRunSql
        cSql.Conn = connOracle
        cSql.SqlStr = "select m_stat_cd from membership where to_char(membership_id) = (select org_relate_id from membership where membership_id = ?)"
        cSql.AddParam(MEMBERSHIP_ID)
        set rsPend = cSql.Execute
        set cSql = nothing
        if rsPend.eof then
        else
            if rsPend("m_stat_cd") = "PENDING" then
                set cSql = new cRunSql
                cSql.Conn = connOracle
                cSql.SqlStr = "update membership set m_stat_cd = 'ACTIVE' where to_char(membership_id) = (select org_relate_id from membership where membership_id = ?)"
                cSql.AddParam(MEMBERSHIP_ID)
                cSql.Execute
                set cSql = nothing
            end if
        end if
        'end parent record

        set cSql = new cRunSql
        cSql.Conn = connOracle
    end if

    if rs("approve_update")&"" = "Y" then

        logIt("start approveUpdate")
        'sequences
        cSql.SqlStr = "select column_name,field_name from app_renew_fields where app_renew_type = ? and field_type = 'SEQUENCE' and update_record = 'Y' and seq_min is null"
        cSql.AddParam(appRenewType)
        set rsSeq = cSql.Execute
        do while not rsSeq.eof
            cMember.updateValue MEMBERSHIP_ID,rsSeq("column_name"),getSequenceMin(rsSeq("field_name"),1)
            rsSeq.MoveNext
        loop
        rsSeq.close
        set rsSeq = nothing

        cSql.SqlStr = "select app_renew_fields.column_name, app_renew_detail.data, app_renew_fields.update_record, append_data, org_sub_ty_cd, app_renew_detail.field_name " &_
            "from app_renew_detail, app_renew_fields, app_renew " &_
            "where upper(app_renew_fields.field_name) = upper(app_renew_detail.field_name) " &_
            "and app_renew.app_renew_id = app_renew_detail.app_renew_id " &_
            "and app_renew_fields.app_renew_type = app_renew.app_renew_type " &_
            "and app_renew_fields.column_name is not null " &_
            "and app_renew_fields.update_record in ('V','Y') " &_
            "and app_renew_detail.app_renew_id = ?"
        cSql.AddParam(app_renew_id)
        set rsAppRenewDetail = cSql.Execute

        if rsAppRenewDetail.eof and rsAppRenewDetail.bof then
        else
            arrAppRenewDetail = rsAppRenewDetail.GetRows
        end if

        rsAppRenewDetail.close
        set rsAppRenewDetail = nothing

        if isArray(arrAppRenewDetail) then
            for i = lbound(arrAppRenewDetail,2) to ubound(arrAppRenewDetail,2)
                logIt("form field " &  arrAppRenewDetail(5,i) & " = " & arrAppRenewDetail(1,i))
                if arrAppRenewDetail(4,i)&"" = "" then
                    if arrAppRenewDetail(3,i) = "Y" and (arrAppRenewDetail(2,i) = "Y" or (arrAppRenewDetail(2,i) = "V" and arrAppRenewDetail(1,i)&"" <> "")) then
                        logIt("MAIN appending " &  arrAppRenewDetail(1,i) & " to " & arrAppRenewDetail(0,i))
                        cMember.AppendValue MEMBERSHIP_ID,arrAppRenewDetail(0,i),arrAppRenewDetail(1,i)
                    else
                        if arrAppRenewDetail(1,i) & "" <> cMember.Value(arrAppRenewDetail(0,i))&"" then
                            if (arrAppRenewDetail(2,i) = "Y") or (arrAppRenewDetail(2,i) = "V" and arrAppRenewDetail(1,i)&"" <> "") then
                                logIt("MAIN updating " & arrAppRenewDetail(0,i) & " to " & arrAppRenewDetail(1,i))
                                cMember.updateValue MEMBERSHIP_ID,arrAppRenewDetail(0,i),arrAppRenewDetail(1,i)
                            end if
                        end if
                    end if
                end if
                if parentType&"" <> "" then
                    if arrAppRenewDetail(4,i)&"" = parentType then
                        if arrAppRenewDetail(3,i) = "Y" and (arrAppRenewDetail(2,i) = "Y" or (arrAppRenewDetail(2,i) = "V" and arrAppRenewDetail(1,i)&"" <> "")) then
                            logIt("PARENT appending " &  arrAppRenewDetail(1,i) & " to " & arrAppRenewDetail(0,i))
                            cParent.AppendValue PARENT_ID,arrAppRenewDetail(0,i),arrAppRenewDetail(1,i)
                        else
                            if arrAppRenewDetail(1,i) & "" <> cParent.Value(arrAppRenewDetail(0,i))&"" then
                                if (arrAppRenewDetail(2,i) = "Y") or (arrAppRenewDetail(2,i) = "V" and arrAppRenewDetail(1,i)&"" <> "") then
                                    logIt("PARENT updating " & arrAppRenewDetail(0,i) & " to " & arrAppRenewDetail(1,i))
                                    cParent.updateValue PARENT_ID,arrAppRenewDetail(0,i),arrAppRenewDetail(1,i)
                                end if
                            end if
                        end if
                    end if
                end if
                if childType&"" <> "" then
                    if arrAppRenewDetail(4,i)&"" = childType then
                        if arrAppRenewDetail(3,i) = "Y" and (arrAppRenewDetail(2,i) = "Y" or (arrAppRenewDetail(2,i) = "V" and arrAppRenewDetail(1,i)&"" <> "")) then
                            logIt("CHILD appending " &  arrAppRenewDetail(1,i) & " to " & arrAppRenewDetail(0,i))
                            cChild.AppendValue CHILD_ID,arrAppRenewDetail(0,i),arrAppRenewDetail(1,i)
                        else
                            if arrAppRenewDetail(1,i) & "" <> cChild.Value(arrAppRenewDetail(0,i))&"" then
                                if (arrAppRenewDetail(2,i) = "Y") or (arrAppRenewDetail(2,i) = "V" and arrAppRenewDetail(1,i)&"" <> "") then
                                    logIt("CHILD updating " & arrAppRenewDetail(0,i) & " to " & arrAppRenewDetail(1,i))
                                    cChild.updateValue CHILD_ID,arrAppRenewDetail(0,i),arrAppRenewDetail(1,i)
                                end if
                            end if
                        end if
                    end if
                end if
                if grandchildType&"" <> "" then
                    if arrAppRenewDetail(4,i)&"" = grandchildType then
                        if arrAppRenewDetail(3,i) = "Y" and (arrAppRenewDetail(2,i) = "Y" or (arrAppRenewDetail(2,i) = "V" and arrAppRenewDetail(1,i)&"" <> "")) then
                            logIt("GRANDCHILD appending " &  arrAppRenewDetail(1,i) & " to " & arrAppRenewDetail(0,i))
                            cGrandchild.AppendValue GRANDCHILD_ID,arrAppRenewDetail(0,i),arrAppRenewDetail(1,i)
                        else
                            if arrAppRenewDetail(1,i) & "" <> cChild.Value(arrAppRenewDetail(0,i))&"" then
                                if (arrAppRenewDetail(2,i) = "Y") or (arrAppRenewDetail(2,i) = "V" and arrAppRenewDetail(1,i)&"" <> "") then
                                    logIt("GRANDCHILD updating " & arrAppRenewDetail(0,i) & " to " & arrAppRenewDetail(1,i))
                                    cGrandchild.updateValue GRANDCHILD_ID,arrAppRenewDetail(0,i),arrAppRenewDetail(1,i)
                                end if
                            end if
                        end if
                    end if
                end if
            Next
        end if

        cSql.SqlStr = "select app_renew_fields.column_name, app_renew_detail.data, app_renew_fields.fields_update_record, app_renew_fields.default_value, app_renew_fields.org_sub_ty_cd " &_
            "from app_renew_detail, app_renew_fields, app_renew " &_
            "where upper(app_renew_fields.field_name) = upper(app_renew_detail.field_name) " &_
            "and app_renew.app_renew_id = app_renew_detail.app_renew_id " &_
            "and app_renew_fields.app_renew_type = app_renew.app_renew_type " &_
            "and app_renew_fields.fields_update_record is not null " &_
            "and app_renew_detail.app_renew_id = ?"
        cSql.AddParam(app_renew_id)
        'cSql.Debug
	    set rsRenewalFields = cSql.Execute
	    do while not rsRenewalFields.eof
		    if rsRenewalFields("DATA")&"" = rsRenewalFields("DEFAULT_VALUE")&"" then
			    arrTmp = split(rsRenewalFields("FIELDS_UPDATE_RECORD"),", ")
			    for i = lbound(arrTmp) to ubound(arrTmp)
				    cSql.sqlStr = "select FIELD_NAME,COLUMN_NAME from app_renew_fields where app_renew_fields_id = ?"
				    cSql.AddParam(arrTmp(i))
				    set rs2 = cSql.Execute
				    if rs2.eof then
				    else
                        cSql.SqlStr = "select DATA from app_renew_detail where app_renew_id = ? and upper(field_name) = upper(?)"
                        cSql.AddParam(app_renew_id)
                        cSql.AddParam(rs2("FIELD_NAME"))
                        set rs3 = cSql.Execute
                        if rs3.eof then
                        else
					        fieldName = trim(rs2("FIELD_NAME"))
						    columnName = trim(rs2("COLUMN_NAME"))
						    fieldValue = rs3("DATA")
                            if rsRenewalFields("ORG_SUB_TY_CD")&"" = "" then
						        if cMember.Value(columnName)&"" <> fieldValue&"" then
                                    'response.write "updating " & columnName & " to " & fieldValue & "<BR>"
							        cMember.updateValue MEMBERSHIP_ID,columnName,fieldValue
						        end if
                            end if
                            if rsRenewalFields("ORG_SUB_TY_CD")&"" = parentType then
                                if PARENT_ID&"" <> "" then
						            if cParent.Value(columnName)&"" <> fieldValue&"" then
                                        'response.write "updating " & columnName & " to " & fieldValue & "<BR>"
							            cParent.updateValue PARENT_ID,columnName,fieldValue
						            end if
                                end if
                            end if
                            if rsRenewalFields("ORG_SUB_TY_CD")&"" = childType then
                                if CHILD_ID&"" <> "" then
						            if cChild.Value(columnName)&"" <> fieldValue&"" then
                                        'response.write "updating " & columnName & " to " & fieldValue & "<BR>"
							            cChild.updateValue CHILD_ID,columnName,fieldValue
						            end if
                                end if
                            end if
                            if rsRenewalFields("ORG_SUB_TY_CD")&"" = grandchildType then
                                if GRANDCHILD_ID&"" <> "" then
						            if cGrandchild.Value(columnName)&"" <> fieldValue&"" then
                                        'response.write "updating " & columnName & " to " & fieldValue & "<BR>"
							            cGrandchild.updateValue GRANDCHILD_ID,columnName,fieldValue
						            end if
                                end if
                            end if
                        end if
				    end if
				    rs2.close
				    set rs2 = nothing
			    next
		    end if
		    rsRenewalFields.MoveNext
	    loop
        rsRenewalFields.close
        set rsRenewalFields = nothing
        logIt("end approveUpdate")
    end if
end if

rs.close
set rs = nothing

if PARENT_ID&"" <> "" then
    if cParent.Value("M_STAT_CD") <> "ACTIVE" then
        cParent.UpdateValue PARENT_ID,"M_STAT_CD","ACTIVE"
    end if
    set cParent = nothing
end if
if CHILD_ID&"" <> "" then
    if cChild.Value("M_STAT_CD") <> "ACTIVE" then
        cChild.UpdateValue CHILD_ID,"M_STAT_CD","ACTIVE"
    end if
    set cChild = nothing
end if
if GRANDCHILD_ID&"" <> "" then
    if cGrandchild.Value("M_STAT_CD") <> "ACTIVE" then
        cGrandchild.UpdateValue GRANDCHILD_ID,"M_STAT_CD","ACTIVE"
    end if
    set cGrandchild = nothing
end if

dim rsAppVals, appVal
cSql.SqlStr = "select COLUMN_NAME,DEFAULT_VALUE,ORG_SUB_TY_CD from app_renew_fields where app_renew_type = ? and update_record = 'A' and default_value is not null"
cSql.AddParam(appRenewType)
set rsAppVals = cSql.Execute
do while not rsAppVals.eof
    appVal = replace(rsAppVals("DEFAULT_VALUE"),"[current date]",date)
    appVal = replace(appVal,"[current year]",year(date))
	for i = 1 to 10
		if rsAppVals("DEFAULT_VALUE") = "[current date + " & i & " years]" then
            appVal = dateAdd("yyyy",i,date)
        end if
		if rsAppVals("DEFAULT_VALUE") = "[current value + " & i & " years]" then
			if cMember.Value(rsAppVals("COLUMN_NAME")) <> "" then
				if isDate(cMember.Value(rsAppVals("COLUMN_NAME"))) then
					appVal = dateAdd("yyyy",i,cdate(cMember.Value(rsAppVals("COLUMN_NAME"))))
				end if
			else
				appVal = dateAdd("yyyy",i,date)
			end if
		end if
	next
    if rsAppVals("ORG_SUB_TY_CD")&"" = "" then
        cMember.UpdateValue MEMBERSHIP_ID,rsAppVals("COLUMN_NAME"),appVal
    elseif rsAppVals("ORG_SUB_TY_CD")&"" = childType then
        cMember.UpdateValue CHILD_ID,rsAppVals("COLUMN_NAME"),appVal
    elseif rsAppVals("ORG_SUB_TY_CD")&"" = parentType then
        cMember.UpdateValue PARENT_ID,rsAppVals("COLUMN_NAME"),appVal
    end if
    rsAppVals.MoveNext
loop
rsAppVals.close
set rsAppVals = nothing

cSql.SqlStr = "select FIELD_NAME,COLUMN_NAME,ORG_SUB_TY_CD from app_renew_fields where app_renew_type = ? and update_record = 'F'"
cSql.AddParam(appRenewType)
set rsAppVals = cSql.Execute
do while not rsAppVals.eof
    appVal = getFormVal(rsAppVals("FIELD_NAME"))
    if rsAppVals("ORG_SUB_TY_CD")&"" = "" then
        cMember.UpdateValue MEMBERSHIP_ID,rsAppVals("COLUMN_NAME"),appVal
    elseif rsAppVals("ORG_SUB_TY_CD")&"" = parentType then
        if PARENT_ID&"" <> "" then
            cMember.UpdateValue PARENT_ID,rsAppVals("COLUMN_NAME"),appVal
        end if
    end if
    rsAppVals.MoveNext
loop
rsAppVals.close
set rsAppVals = nothing

set cMember = nothing
%>
<%End Sub%>

<%Sub acceptAwait()%>

<%
    dim MEMBERSHIP_ID

    set cSql = new cRunSql
    cSql.DisplayErrors = true
    cSql.Conn = connOracle

    cSql.SqlStr = "select MEMBERSHIP_ID from app_renew where app_renew_id = ?"
    cSql.AddParam(request.querystring("id"))
    set rs = cSql.Execute

    MEMBERSHIP_ID = rs("MEMBERSHIP_ID")

	cSql.sqlStr = "SELECT DATE_DATA,DATA,MEMBER_FIELD "&_
	"FROM MEM_AWAIT "&_
	"WHERE MEMBERSHIP_ID = ? "&_
	"AND APPROVAL_DT IS NULL"
    cSql.AddParam(MEMBERSHIP_ID)
    'cSql.Debug
	set rs = cSql.Execute

    set cSql = nothing
	do while not rs.Eof

        set cSql = new cRunSql
        cSql.DisplayErrors = true
        cSql.Conn = connOracle
		If rs.fields("DATE_DATA").value <> "" then
			cSql.sqlStr = "UPDATE MEMBERSHIP "&_
			"SET "&rs.fields("MEMBER_FIELD").value&" = to_date(?,'mm/dd/yyyy') "&_
			"WHERE MEMBERSHIP_ID = ?"
            cSql.AddParam(rs.fields("DATE_DATA").value)
		Else
			cSql.sqlStr = "UPDATE MEMBERSHIP "&_
			"SET "&rs.fields("MEMBER_FIELD").value&" = ? "&_
			"WHERE MEMBERSHIP_ID = ?"
            cSql.AddParam(rs.fields("DATA").value)
		End If
        cSql.AddParam(MEMBERSHIP_ID)
        'cSql.Debug
		cSql.Execute
        set cSql = nothing
		rs.MoveNext
	Loop
	rs.Close

    set cSql = new cRunSql
    cSql.DisplayErrors = true
    cSql.Conn = connOracle
	cSql.SqlStr = "UPDATE MEM_AWAIT "&_
	"SET APPROVAL_DT = sysdate "&_
	"WHERE MEMBERSHIP_ID = ? "&_
	"AND APPROVAL_DT IS NULL"
    cSql.AddParam(MEMBERSHIP_ID)
    'cSql.Debug
	cSql.Execute

	'DO THE SAME FOR DETAILS

	cSql.sqlStr = "SELECT CHANGE_TYPE,COLUMN_ID,DATA "&_
	"FROM MEM_DET_AWAIT "&_
	"WHERE MEMBERSHIP_ID = ? "&_
	"AND APPROVAL_DT IS NULL"

    cSql.AddParam(MEMBERSHIP_ID)
    'cSql.Debug
	set rs = cSql.Execute
    set cSql = nothing

	do while not rs.Eof
        set cSql = new cRunSql
        cSql.Conn = connOracle
    cSql.DisplayErrors = true
        'response.write "DATA = " & rs.fields("DATA").value & "<BR>"
        'response.write "COLUMN_ID = " & rs.fields("COLUMN_ID").value & "<BR>"
        'response.write "MEMBERSHIP_ID = " & MEMBERSHIP_ID & "<BR>"
		If rs.fields("CHANGE_TYPE").value = "DELETE" then
			cSql.sqlStr = "DELETE FROM "&_
			"MEMBER_DETAIL "&_
			"WHERE MEMBERSHIP_ID = ? "&_
			"AND COLUMN_ID = ?"
            cSql.AddParam(MEMBERSHIP_ID)
            cSql.AddParam(rs.fields("COLUMN_ID").value)
		else
		    cSql.sqlStr = "select COLUMN_ID,DATA from member_detail where MEMBERSHIP_ID = ? and column_id = ?"
		'response.write sqlStr & "<BR>"
            cSql.AddParam(MEMBERSHIP_ID)
            cSql.AddParam(rs.fields("COLUMN_ID").value)
    'cSql.Debug
		    set rsEmpty = cSql.Execute
		    if rsEmpty.eof and rsEmpty.bof then
			    cSql.sqlStr = "INSERT INTO "&_
			    "MEMBER_DETAIL "&_
			    "(MEMBERSHIP_ID, COLUMN_ID, DATA) "&_
			    "VALUES (?, ?, ?)"
                cSql.AddParam(MEMBERSHIP_ID)
                cSql.AddParam(rs.fields("COLUMN_ID").value)
                cSql.AddParam(replace(cleanData(rs.fields("DATA").value&""),"[current date await]",date))
		    else
			    cSql.sqlStr = "UPDATE MEMBER_DETAIL "&_
			    "SET DATA = ? "&_
			    "WHERE MEMBERSHIP_ID = ? "&_
			    "AND COLUMN_ID = ? "
                cSql.AddParam(replace(cleanData(rs.fields("DATA").value&""),"[current date await]",date))
                cSql.AddParam(MEMBERSHIP_ID)
                cSql.AddParam(rs.fields("COLUMN_ID").value)
		    end if
		End If

    'cSql.Debug
        cSql.Execute

        set cSql = nothing

	    rs.MoveNext
	Loop
	rs.Close

    set cSql = new cRunSql
    cSql.Conn = connOracle
    'response.end
	cSql.SqlStr = "UPDATE MEM_DET_AWAIT "&_
	"SET APPROVAL_DT = sysdate "&_
	"WHERE MEMBERSHIP_ID = ? "&_
	"AND APPROVAL_DT IS NULL"

    cSql.AddParam(MEMBERSHIP_ID)
    cSql.Execute
    set cSql = nothing
	%>
<%End Sub%>

<%Sub removeAwait%>
<%
    dim MEMBERSHIP_ID

    set cSql = new cRunSql
    cSql.DisplayErrors = true
    cSql.Conn = connOracle

    cSql.SqlStr = "select MEMBERSHIP_ID from app_renew where app_renew_id = ?"
    cSql.AddParam(request.querystring("id"))
    set rs = cSql.Execute

    MEMBERSHIP_ID = rs("MEMBERSHIP_ID")

    cSql.SqlStr = "UPDATE MEM_AWAIT "&_
	"SET APPROVAL_DT = to_date('01/01/1800', 'MM/DD/YYYY') "&_
	"WHERE MEMBERSHIP_ID = ? "&_
	"AND APPROVAL_DT IS NULL"

    cSql.AddParam(MEMBERSHIP_ID)
    cSql.Execute

	'DO THE SAME FOR DETAILS
	cSql.SqlStr = "UPDATE MEM_DET_AWAIT "&_
	"SET APPROVAL_DT = to_date('01/01/1800', 'MM/DD/YYYY') "&_
	"WHERE MEMBERSHIP_ID = ? "&_
	"AND APPROVAL_DT IS NULL"

    cSql.AddParam(MEMBERSHIP_ID)
    cSql.Execute
    set cSql = nothing
	%>
<%End Sub%>

<%function getEmailText(file)%>
<%
dim objFSO, TS, emailText, emailFile

emailFile = "e:\p13231\hosting\global\admin\appemail\" & file

' Instantiate the FileSystemObject
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

'response.write dataFile

' use Opentextfile Method to Open the text File
Const ForReading = 1 
Const Create = False

Set TS = objFSO.OpenTextFile(emailFile, ForReading, Create)

If Not TS.AtEndOfStream  Then   
    emailText = ""
	Do While Not TS.AtendOfStream
		emailText = emailText & ts.ReadLine
    loop
end if
TS.Close 
set TS = Nothing
Set objFSO = Nothing

getEmailText = emailText
%>
<%end function%>

<%sub wvbopCompApprove(cMember)%>
<%
dim cNewMem, membershipID, newID, recType, rsAR, cDoc, caseNo, rsCaseNo, caseNoMo

cSql.SqlStr = "select MEMBERSHIP_ID from app_renew where app_renew_id = ?"
cSql.AddParam(request.querystring("id"))
set rsAR = cSql.Execute
membershipID = rsAR("MEMBERSHIP_ID")

cSql.SqlStr = "Select SEQ_" & year(date) & "_DISC_CASE_ID.NextVal from DUAL"
set rsCaseNo = cSql.Execute
if month(date) < 10 then
    caseNoMo = "0" & month(date)
else
    caseNoMo = month(date)
end if
caseNo = year(date) & "-" & caseNoMo & "-" & rsCaseNo("NextVal")
rsCaseNo.close
set rsCaseNo = nothing

if request("type") = "BOTH" then
    SqlStr = "update membership set comments = '" & now & "',org_sub_ty_cd = 'Disc_Faciltiy' where membership_id = " & membershipID
    connOracle.execute(sqlStr)
    cMember.UpdateValue membershipID,"COMMENTS",""

    cSql.SqlStr = "select app_renew_fields.column_name, app_renew_detail.data " &_
        "from app_renew_detail, app_renew_fields " &_
        "where upper(app_renew_fields.field_name) = upper(app_renew_detail.field_name) " &_
        "and app_renew_fields.app_renew_type = ? " &_
        "and app_renew_fields.column_name is not null " &_
        "and app_renew_fields.update_record = 'Y' " &_
        "and app_renew_detail.app_renew_id = ?"
    cSql.AddParam("COMPL")
    cSql.AddParam(request.querystring("id"))
    set rsAppRenewDetail = cSql.Execute

    if rsAppRenewDetail.eof and rsAppRenewDetail.bof then
    else
        arrAppRenewDetail = rsAppRenewDetail.GetRows
    end if

    if isArray(arrAppRenewDetail) then
        for i = lbound(arrAppRenewDetail,2) to ubound(arrAppRenewDetail,2)
    '		response.write arrAppRenewDetail(0,i) & " " & arrAppRenewDetail(1,i) & " " & cMember.Value(arrAppRenewDetail(0,i))&"<Br>"
            if arrAppRenewDetail(1,i) & "" <> cMember.Value(arrAppRenewDetail(0,i))&"" then
                cMember.updateValue membershipID,arrAppRenewDetail(0,i),arrAppRenewDetail(1,i)
            end if
        Next
        cMember.updateValue membershipID,"COMPLAINT_CASE_NUMBER_Disc_Faciltiy",caseNo & "B"
    end if

    cMember.LoadFromMembershipID membershipID

    set cNewMem = new cNewMemberO
    newID = cNewMem.MembershipID
    cNewMem.setValue "M_STAT_CD", "ACTIVE"
    cNewMem.setValue "ORG_RELATE_ID", cMember.Value("ORG_RELATE_ID")
    cNewMem.setValue "ORG_SUB_TY_CD","Disc"
    cNewMem.InsertRecord
    set cNewMem = nothing

    cMember.LoadFromMembershipID newID

    set cDoc = new cDocumentRepositoryO
    cDoc.LoadFromMembershipIDAndAppRenewID membershipID,request.querystring("id")
    do while not cDoc.eof
        cDoc.Copy cDoc.ResourceId,newID
        cDoc.MoveNext
    loop
    set cDoc = nothing

    cSql.SqlStr = "select app_renew_fields.column_name, app_renew_detail.data " &_
        "from app_renew_detail, app_renew_fields " &_
        "where upper(app_renew_fields.field_name) = upper(app_renew_detail.field_name) " &_
        "and app_renew_fields.app_renew_type = ? " &_
        "and app_renew_fields.column_name is not null " &_
        "and app_renew_fields.update_record = 'Y' " &_
        "and app_renew_detail.app_renew_id = ?"
    cSql.AddParam("COMPL_INDV")
    cSql.AddParam(request.querystring("id"))
    set rsAppRenewDetail = cSql.Execute

    if rsAppRenewDetail.eof and rsAppRenewDetail.bof then
    else
        arrAppRenewDetail = rsAppRenewDetail.GetRows
    end if

    if isArray(arrAppRenewDetail) then
        for i = lbound(arrAppRenewDetail,2) to ubound(arrAppRenewDetail,2)
    '		response.write arrAppRenewDetail(0,i) & " " & arrAppRenewDetail(1,i) & " " & cMember.Value(arrAppRenewDetail(0,i))&"<Br>"
            if arrAppRenewDetail(1,i) & "" <> cMember.Value(arrAppRenewDetail(0,i))&"" then
                cMember.updateValue newID,arrAppRenewDetail(0,i),arrAppRenewDetail(1,i)
            end if
        Next
        cMember.updateValue newID,"COMPLAINT_CASE_NUMBER_Disc",caseNo & "A"
    end if

else
    cSql.SqlStr = "select minor_code from code_value where code_class = 'APP_RENEW_TYPE_CD' and code_value = ?"
    cSql.AddParam(request("type"))
    'cSql.Debug
    set rsAR = cSql.Execute
    recType = rsAR("MINOR_CODE")
    SqlStr = "update membership set comments = '" & now & "',org_sub_ty_cd = '" & rectype & "' where membership_id = " & membershipID
    connOracle.execute(sqlStr)
    cMember.UpdateValue membershipID,"COMMENTS",""

    cSql.SqlStr = "select app_renew_fields.column_name, app_renew_detail.data " &_
        "from app_renew_detail, app_renew_fields " &_
        "where upper(app_renew_fields.field_name) = upper(app_renew_detail.field_name) " &_
        "and app_renew_fields.app_renew_type = ? " &_
        "and app_renew_fields.column_name is not null " &_
        "and app_renew_fields.update_record = 'Y' " &_
        "and app_renew_detail.app_renew_id = ?"
    cSql.AddParam(request("type"))
    cSql.AddParam(request.querystring("id"))
    'cSql.Debug
    set rsAppRenewDetail = cSql.Execute

    if rsAppRenewDetail.eof and rsAppRenewDetail.bof then
    else
        arrAppRenewDetail = rsAppRenewDetail.GetRows
    end if

    if isArray(arrAppRenewDetail) then
        for i = lbound(arrAppRenewDetail,2) to ubound(arrAppRenewDetail,2)
    		response.write arrAppRenewDetail(0,i) & " " & arrAppRenewDetail(1,i) & " " & cMember.Value(arrAppRenewDetail(0,i))&"<Br>"
            if arrAppRenewDetail(1,i) & "" <> cMember.Value(arrAppRenewDetail(0,i))&"" then
                cMember.updateValue membershipID,arrAppRenewDetail(0,i),arrAppRenewDetail(1,i)
            end if
        Next
    end if
    cMember.updateValue membershipID,"COMPLAINT_CASE_NUMBER_Disc",caseNo
    cMember.updateValue membershipID,"COMPLAINT_CASE_NUMBER_Disc_Faciltiy",caseNo
end if

rsAR.close
set rsAR = nothing

rsAppRenewDetail.close
set rsAppRenewDetail = nothing
%>
<%end sub%>

<%sub logIt(strLog)%>
<%
cLogit.BatchLog(strLog)
%>
<%end sub%>

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
