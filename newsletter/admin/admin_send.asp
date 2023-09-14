<!-- #include file="../includes/general_includes.asp"-->
<!-- #include file="../includes/_upload.asp"-->
<%
	on error resume next
	Dim DestinationPath, Form

	strCookies = Request.Cookies("Admin")("name")

	If strCookies = "" Then

		Response.Redirect "login.asp"

	End If

	If Not blnSend Then

	    Response.Cookies("msg") = "nar"
	    Response.Redirect "admin.asp"

	End If

    msg = ""
    msg = Trim(Request.Cookies("msg"))

	If msg <> "" Then
		Call displayFancyMsg(getMessage(msg))
        Response.Cookies("msg") = ""
	End If

%>
<!-- #include file="../includes/header.asp"-->
<div id="main" class="container">
    <header>
        <h2 style="text-align:center;">Send a Newsletter</h2>
    </header>
    <%
	Set Conn = Server.CreateObject("ADODB.Connection")
	Call ConnOpen(Conn)

	If Trim(Request.Form("send")) <> "" Then

		If Request.Form("selectattach") <> "" Then
		     strAttach = "admin/attachs/"&Request.Form("selectattach")
		Else
		    strAttach = ""
		End If

		If Request.Form("send") = "Send as HTML" Then
		  intSend = 1
		Else
		  intSend = 2
		End If

        Response.Write "<div id=""outputwindow"" class=""row uniform"">"&vbcrlf
        Response.Write "  <div class=""-2u 8u$"">"&vbcrlf
        Response.Write "    <h3>Output Window</h3><br />"&vbcrlf
        Response.Write "    <textarea rows=""5"">"&vbcrlf

		subject = Trim(Request.Form("subject"))
		strEmailMsg = Trim(Request.Form("tempbody"))

	    nums = Request.Form("semail").Count
		For i = 1 To nums
		    strEmail = Request.Form("semail")(i)
		    Call send_email("",strEmail,subject,strEmailMsg,strAttach,intSend)
		    Response.Flush
		Next

        Response.Write "    </textarea>"&vbcrlf
        Response.Write "  </div>"&vbcrlf
        Response.Write "</div>"&vbcrlf

	End If

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

	Set objFSO = CreateObject("Scripting.FileSystemObject")

	strSaveFile = ""
	strSaveFile = Server.MapPath(strDir&"admin\attachs")

	If Not objFSO.FolderExists(strSaveFile) Then
		objFSO.CreateFolder(strSaveFile)
	End If
	
	strPathInfo = strSaveFile
	
	Set objFolder = objFSO.GetFolder(strPathInfo)
	Set objFolderContents = objFolder.Files
%>

    <div class="row uniform">
        <div class="-1u 5u 12u$(medium)" style="vertical-align:middle;">
			<h5>You have <%= intDCount %> Drafts and <%= intTCount %> Templates.</h5>
			<div class="row">
				<div class="12u$">
					<a class="picimg button fit fancybox.ajax" href="imageurls.asp">Get Image URLs</a>
                </div>
			</div>
        </div>
		<div class="5u$ 12u$(medium)">
			<form action="uploadattach.asp?dla=no&p=s" method="post" enctype="multipart/form-data">
			<input type="hidden" name="pfrom" value="send" />
				<h5>Upload Attachment</h5>
				<div class="row">
					<div class="8u 12u$(medium)">
						<input type="file" id="userfile" name="userfile" class="button fit" />
					</div>
					<div class="4u$ 12u$(medium)">
						<input type="submit" name="submit" value="Upload" class="button fit" />
                    </div>
				</div>	
			</form>		
		</div>
        <div class="-1u 5u 12u$(medium)">
			<label for="loadtemp">Template Title</label>
            <div class="select-wrapper">
                <% Call selectLoadTemplate %>
            </div>
        </div>
		<div class="5u$ 12u$(medium)">
			<label for="tempdescr">Template Description</label>
             <input type="text" id="tempdescr" name="tempdescr" value="" size="30" required />
		</div>
    </div>
    <form action="admin_send.asp#outputwindow" id="template" method="post">
        <input type="hidden" name="tempid" id="tempid" value="" />
        <input type="hidden" name="temptitle" id="temptitle" value="" />
		<input type="hidden" name="attach" id="attach" value="" />
        <div class="row uniform">
            <div class="-1u 10u 12u(medium)">
				<div class="row">
					<div class="12u 12u$(small)" style="padding-bottom: 10px;">
						<label for="subject">Subject</label>
						<input type="text" name="subject" id="subject" value="" size="30" required />
					</div>
					<div class="12u  12u$(small)" style="margin-bottom:20px;">
						<textarea name="tempbody" id="tempbody" cols="65" rows="25" wrap="soft"></textarea>
						<script>
							CKEDITOR.replace( 'tempbody');
						</script>
					</div>
					<div class="4u 12u$(medium)">
                        <select id="semail" name="semail" size="5" style="height:75px;color:#000000;" multiple>
<%
	intCount = 0
	Set Conn = Server.CreateObject("ADODB.Connection")
	Call ConnOpen(Conn)

	Set rsCommon = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT * FROM "&msdbprefix&"addresses WHERE confirm = 'yes'"

	Call getTextRecordset(strSQL,rsCommon)
	If Not rsCommon.EOF Then
		Do While Not rsCommon.EOF

	        intCount = intCount + 1
	        Response.Write "<option value="""&DBDecode(rsCommon("email"))&""">"&intCount&". "&DBDecode(rsCommon("email"))&"</option>"&vbcrlf
			
			rsCommon.MoveNext
	       If rsCommon.EOF Then Exit Do
		Loop
	End If
	Call closeRecordset(rsCommon)
	Call ConnClose(Conn)

%>
                        </select>
					</div>
					<div class="3u 12u$(medium)">
						<input type="button" class="button fit" value="Select All Emails" onclick="selectAll('semail');" />
					</div>
					<div class="5u$ 12u$(medium)">   
						<div class="select-wrapper">
							<select name="selectattach" id="selectattach">
<% 
	If objFolder.Files.Count <> 0 Then
	    Response.Write "<option value="""">Select Attachment</option>"&vbcrlf
		For each objFileItem In objFolderContents
			Response.Write "<option value="""&objFileItem.Name&""">"&objFileItem.Name&"</option>"&vbcrlf
		Next
	Else
	    Response.Write "<option value="""">No Attachments</option>"&vbcrlf
	End If

	Set objFSO = Nothing
%>
							</select>
							<label for="selectattach">Add Attachment</label>
						</div>
                    </div>
                </div>
                <div class="row" style="padding-top: 10px;">
                    <div class="6u 12u$(small)" style="text-align: center;">
                        <input class="button fit" type="submit" name="send" value="Send as HTML" />
                    </div>
					<div class="6u$ 12u$(small)" style="text-align: center;">
						<input class="button fit" type="submit" name="send" value="Send as Plain Text" />
					</div>
                </div>
            </div>
        </div>
    </form>
</div>
<!-- #include file="../includes/footer.asp"-->