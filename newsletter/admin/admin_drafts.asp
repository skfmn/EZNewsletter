<!-- #include file="../includes/general_includes.asp"-->
<%
	on error resume next
	strCookies = Request.Cookies("Admin")("name")

	If strCookies = "" Then

		Response.Redirect "login.asp"

	End If

	If Not blnTemplates Then

	    Response.Cookies("msg") = "nar"
	    Response.Redirect "admin.asp"

	End If

    msg = ""
    msg = Trim(Request.Cookies("msg"))

	If msg <> "" Then
		Call displayFancyMsg(getMessage(msg))
        Response.Cookies("msg") = ""
	End If

	If Trim(Request.Form("save")) <> "" Then

	   If Request.Form("save") = "Save as Template" Then
	       strTempSave = "template"
	   Else
	       strTempSave = "draft"
	   End If
	    
	    If Trim(Request.Form("tempid")) <> "" Then
	        intTemplateID = Trim(Request.Form("tempid"))
		Else
			intTemplateID = 0
		End If
		
		strTempTitle = DBEncode(Trim(Request.Form("temptitle")))
	    strTempDescription = DBEncode(Trim(Request.Form("tempdescr")))
		strTempBody = DBEncode(Trim(Request.Form("tempbody")))

		Set Conn = Server.CreateObject("ADODB.Connection")
		Call ConnOpen(Conn)

	    Set rsCommon = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM "&msdbprefix&"newsletter WHERE newsletterID = "&intTemplateID

		Call getTextRecordset(strSQL,rsCommon)
		If NOT rsCommon.EOF Then
			rsCommon("news_title") = strTempTitle
	        rsCommon("news_save") = strTempSave
			rsCommon("news_description") = strTempDescription
	        rsCommon("news_body") = strTempBody
			rsCommon.Update
			Response.Cookies("msg") = "updated"
		Else
		    strSQL = "INSERT INTO "&msdbprefix&"newsletter ([news_title],[news_save],[news_description],[news_body]) Values('"&strTempTitle&"','"&strTempSave&"','"&strTempDescription&"','"&strTempBody&"')"
			Call getExecuteQuery(strSQL)
			Response.Cookies("msg") = "success"
		End If
		Call closeRecordset(rsCommon)
		Call ConnClose(Conn)

        Response.Redirect "admin_drafts.asp"

	End If

	If Trim(Request.Form("delete")) <> "" Then

		  lngTempplateID = checkint(Trim(Request.Form("tempid")))

		  Set Conn = Server.CreateObject("ADODB.Connection")
		  strSQL = "DELETE FROM "&msdbprefix&"newsletter WHERE newsletterID = "&lngTempplateID

		Call ConnOpen(Conn)
		Call getExecuteQuery(strSQL)
		Call ConnClose(Conn)

	    Response.Cookies("msg") = "del"
		Response.Redirect "admin_drafts.asp"

	End If

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
        <h2 style="text-align:center;">Manage Drafts</h2>
    </header>
    <div class="row">
        <div class="-4u 4u$ 12u$(medium)">
            <h5>You have <%= intDCount %> Drafts and <%= intTCount %> Templates.</h5>
			<div class="row">
				<div class="12u$">
					<a class="picimg button fit fancybox.ajax" href="imageurls.asp">Get Image URLs</a>
                </div>
                <div class="12u$">
					<div class="select-wrapper">
						<% Call selectLoadDraft %>
					</div>
                </div>
			</div>
        </div>
    </div>
    <form action="admin_drafts.asp" id="template" method="post">
        <input type="hidden" name="tempid" id="tempid" value="" />
        <div class="row uniform">
            <div class="-1u 10u 12u(medium)">
                <div class="row">
                    <div class="6u 12u$(small)" style="padding-bottom: 10px;">
                        <label for="temptitle">Title</label>
                        <input type="text" name="temptitle" id="temptitle" value="" />
                    </div>
                    <div class="6u$ 12u$(medium)" style="padding-bottom: 10px;">
                        <label for="tempdescr">Template Description</label>
                        <input type="text" id="tempdescr" name="tempdescr" value="" />
                    </div>
                    <div class="12u  12u$(small)">
                        <textarea name="tempbody" id="tempbody" cols="65" rows="25" wrap="soft"></textarea>
                        <script>
							CKEDITOR.replace( 'tempbody');
                        </script>
                    </div>
                </div>

                <div class="row" style="padding-top:10px;">
                    <div class="4u 12u$(small)">
                        <input type="submit" name="save" class="button fit" value="Save as Draft" />
                    </div>
                    <div class="4u 12u$(small)">
                        <input type="submit" name="save" class="button fit" value="Save as Template" />
                    </div>
                    <div class="4u$ 12u$(small)">
                        <input type="submit" name="delete" class="button fit" value="Delete" onclick="return confirm('WARNING!\n Are you SURE you want to delete this template?')" />
                    </div>
                </div>

            </div>
        </div>
    </form>
</div>
<!-- #include file="../includes/footer.asp"-->