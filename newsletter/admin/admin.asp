<!-- #include file="../includes/general_includes.asp"-->
<%
	strCookies = Request.Cookies("Admin")("name")

	If strCookies = "" Then

		Response.Redirect "login.asp"

	End If

    msg = ""
    msg = Trim(Request.Cookies("msg"))

	If msg <> "" Then
		Call displayFancyMsg(getMessage(msg))
        Response.Cookies("msg") = ""
	End If


    'Auto deletes the install folder.
    'If for some reason it didn't work manaully delete the folder and it's contents
    'You can delete this after the folder is gone.
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FolderExists(Server.MapPath(strDir&"install")) Then
		fso.DeleteFolder(Server.MapPath(strDir&"install"))
	End If
	Set fso = Nothing
    'You can delete this after the folder is gone.''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



	Set Conn = Server.CreateObject("ADODB.Connection")
	Call ConnOpen(Conn)

	Set rsCommon = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT * FROM "&msdbprefix&"newsletter WHERE news_save = 'draft'"

	Call getTextRecordset(strSQL,rsCommon)

	If Not rsCommon.EOF Then
        intDCount = rsCommon.RecordCount
    Else
        intDCount = 0
    End If
	Call closeRecordset(rsCommon)

    Set rsCommon = Server.CreateObject("ADODB.Recordset")
    strSQL = "SELECT * FROM "&msdbprefix&"newsletter WHERE news_save = 'template'"

	Call getTextRecordset(strSQL,rsCommon)
	If Not rsCommon.EOF Then
        intTCount = rsCommon.RecordCount
    Else
        intTCount = 0
    End If
	Call closeRecordset(rsCommon)
	Call ConnClose(Conn)
%>
<!-- #include file="../includes/header.asp"-->
<div id="main" class="container">
    <header>
        <h4 style="text-align: center;">Choose an Option below</h4>
    </header>
    <div class="row">
        <div class="-3u 3u 12u(medium)">
            <ul class="alt">
                <li><a class="button fit" href="admin_send.asp">Send a Newsletter</a></li>
                <li><a class="button fit" href="admin_template.asp">Manage Templates (<%= intTCount %>)</a></li>
                <li><a class="button fit" href="admin_images.asp">Manage Images</a></li>
                <li><a class="button fit" href="admin_manage.asp">Manage Admins</a></li>
                
            </ul>
        </div>
        <div class="3u$ 12u$(medium)">
            <ul class="alt">
                <li><a class="button fit" href="admin_create.asp">Create Template/Draft</a></li>
                <li><a class="button fit" href="admin_drafts.asp">Manage Drafts (<%= intDCount %>)</a></li>
                <li><a class="button fit" href="admin_addresses.asp">Manage Addresses</a></li>
                <li><a class="button fit" href="admin_options.asp">Manage Options</a></li>
            </ul>
        </div>
        <div class="-3u 6u 12u$(medium)">
            <%= getResponse("http://www.aspjunction.com/gnews.asp?ref=y&amp;nv="& strVersion&"") %>
        </div>

    </div>
</div>
<!-- #include file="../includes/footer.asp"-->