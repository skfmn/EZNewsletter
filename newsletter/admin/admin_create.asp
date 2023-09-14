<!-- #include file="../includes/general_includes.asp"-->
<%
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

    If Trim(Request.QueryString("save")) <> "" Then

        strTemplateTitle = DBEncode(Request.Form("temptitle"))
        strTemplateDescr = DBEncode(Request.Form("tempdescr"))
        strTemplateBody = DBEncode(Request.Form("tempbody"))
        strTemplateSave = DBEncode(Request.QueryString("save"))

        Set Conn = Server.CreateObject("ADODB.Connection")
        Call ConnOpen(Conn)

        Set rsCommon = Server.CreateObject("ADODB.Recordset")
        strSQL = "SELECT * FROM "&msdbprefix&"newsletter WHERE news_title = '"&strTemplateTitle&"'"

        Call getTextRecordset(strSQL,rsCommon)
        If Not rsCommon.EOF Then
            Response.Redirect "admin_create.asp?msg=nt"
        Else
            strSQL = "INSERT INTO "&msdbprefix&"newsletter(news_title,news_save,news_description,news_body) Values('"&strTemplateTitle&"','"&strTemplateSave&"','"&strTemplateDescr&"','"&strTemplateBody&"')"
            Call getExecuteQuery(strSQL)

            Response.Cookies("msg") = "tc"
            Response.Redirect "admin_create.asp"
        End If
        Call closeRecordset(rsCommon)
        Call ConnClose(Conn)

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
        <h2 style="text-align: center;">Create a Template</h2>
    </header>
    <div class="row">
        <div class="-4u 4u$ 12u$(medium)">
            <h5>You have <%= intDCount %> Drafts and <%= intTCount %> Templates.</h5>
			<div class="row">
				<div class="12u$">
					<a class="picimg button fit fancybox.ajax" href="imageurls.asp">Get Image URLs</a>
                </div>
			</div>
        </div>
    </div>

    <form method="post">
        <input type="hidden" name="recipients" value="<%'= recipients %>">
        <div class="row uniform">
            <div class="-1u 10u 12u(medium)">
                <div class="row uniform">
                    <div class="6u 12u$(small)" style="padding-bottom: 10px;">
                        <label for="temptitle">Template Title <span style="font-size: 10px;">(Titles in <span style="color: #ff0000;">RED</span> are taken)</span></label>
                        <input type="text" id="temptitle" name="temptitle" size="30" placeholder="Title of the Template" required />
                    </div>
                    <div class="6u$ 12u$(small)" style="padding-bottom: 10px;">
                        <label for="tempdescr">Template Description</label>
                        <input type="text" id="tempdescr" name="tempdescr" size="30" placeholder="Description of the Template" required />
                    </div>
                </div>
                <div class="12u  12u$(small)">
                    <textarea name="tempbody" id="tempbody" wrap="soft" style="height: 300px;" required><%= strTempBody %></textarea>
                    <script>
						CKEDITOR.replace( 'tempbody');
                    </script>
                </div>

                <div class="row" style="padding-top: 10px; text-align: center;">
                    <div class="6u 12u$(small)">
                        <input class="button fit" type="submit" name="save" value="Save as Template" formaction="admin_create.asp?save=template" />
                    </div>
                    <div class="6u$ 12u$(small)">
                        <input class="button fit" type="submit" name="save" value="Save as Draft" formaction="admin_create.asp?save=draft" />
                    </div>
                </div>
            </div>
        </div>
    </form>
</div>
<!-- #include file="../includes/footer.asp"-->