<!-- #include file="../includes/general_includes.asp"-->
<%
on error resume next
 	strCookies = Request.Cookies("Admin")("name")

	If strCookies = "" Then

		Response.Redirect "login.asp"

	End If

	If Not blnOptions Then

	    Response.Cookies("msg") = "nar"
	    Response.Redirect "admin.asp"

	End If

    msg = ""
    msg = Trim(Request.Cookies("msg"))

	If msg <> "" Then
		Call displayFancyMsg(getMessage(msg))
        Response.Cookies("msg") = ""
	End If

    Set Conn = Server.CreateObject("ADODB.Connection")
    Call ConnOpen(Conn)

    If Request.Form("chmsg") <> "" Then

        For Each i in Request.Form
	        If left(i,8) = "messages" Then

		        strFormValue = Replace(i,left(i,9),"")
		        strFormValue = Replace(strFormValue,right(i,1),"")
                strFormMsg = Request.Form(i)

                strSQL = "UPDATE " & msdbprefix & "messages SET message = '"&DBEncode(strFormMsg)&"' WHERE msg = '"&strFormValue&"'"
                Call getExecuteQuery(strSQL)

	        End If
        Next

        Response.Cookies("msg") = "mus"
        Response.Redirect "admin_options.asp"

    End If

    If Request.Form("chumsg") <> "" Then

        For Each i in Request.Form

	        If left(i,12) = "usermessages" Then

		        strFormValue = Replace(i,left(i,13),"")
		        strFormValue = Replace(strFormValue,right(i,1),"")

                strSQL = "UPDATE " & msdbprefix & "endMsg SET endMsg = '"&DBEncode(Request.Form(i))&"' WHERE endMsgName = '"&strFormValue&"'"
                Call getExecutequery(strSQL)

	        End If
        Next

        Response.Cookies("msg") = "mus"
        Response.Redirect "admin_options.asp"

    End If

   If Request.Form("chmstgs") <> "" Then

        strSiteTitle = DBEncode(Request.Form("sitetitle"))
        smtpServer = DBEncode(Request.Form("smtpserver"))
        smtpEmail = DBEncode(Request.Form("smtpemail"))
        smtpPassword = DBEncode(Request.Form("smtppassword"))
        smtpPort = Request.Form("smtpport")

        strUrlrewrite = "no"
        strUrlrewrite = Request.Form("urlrewrite")
        If strUrlrewrite = "on" Then
            blnUrlrewrite = 1
        Else
            blnUrlrewrite = 0
        End If

       strSQL = "UPDATE " & msdbprefix & "settings SET site_title = '"&strSiteTitle&"', smtp_server = '"&smtpServer&"', email_address = '"&smtpEmail&"', smtp_password = '"&smtpPassword&"', rewrite = "&blnUrlrewrite&", smtp_port = '"&smtpPort&"'"
        Call getExecutequery(strSQL)

        Response.Cookies("msg") = "siu"
        Response.Redirect "admin_options.asp"

    End If

   If Request.Form("confirmemail") <> "" Then

        strConfirmEmail = Trim(Replace(Request.Form("tempbody"),vbcrlf,""))

        strSQL = "UPDATE " & msdbprefix & "settings SET  confirm_email = '"&strConfirmEmail&"'"
        Call getExecutequery(strSQL)

        Response.Cookies("msg") = "siu"
        Response.Redirect "admin_options.asp"

    End If

%>
<!-- #include file="../includes/header.asp"-->
<div id="main" class="container">
    <header>
        <h2>Manage options</h2>
    </header>
    <div class="row uniform">
        <div class="6u 12u$(medium)">
            <h3>Admin Messages</h3>
            <div class="12u$" style="padding-bottom: 10px;">
                <div class="table-wrapper">
                    <form action="admin_options.asp" method="post">
                        <input type="hidden" name="chmsg" value="y">
                        <table>
                            <tbody>
                                <%
    Set rsCommon = Server.CreateObject("ADODB.Recordset")
    strSQL = "SELECT * FROM " & msdbprefix & "messages"

    Call getTextRecordset(strSQL, rsCommon)
    If Not rsCommon.EOF Then
        Do While Not rsCommon.EOF
            strTempMessage = DBDecode(rsCommon("message"))
            strTempMsg = rsCommon("msg")
                                %>
                                <tr>
                                    <td style="width: 30%;">
                                        <%= msgTrans(strTempMsg) %>
                                    </td>
                                    <td style="width: 70%;">
                                        <input type="text" name="messages[<%= strTempMsg %>]" value="<%= strTempMessage %>">
                                    </td>
                                </tr>
                                <%
            rsCommon.MoveNext
            If rsCommon.EOF Then Exit DO
        Loop
    End If
    Call CloseRecordset(rsCommon)
                                %>
                                <tfoot>
                                    <tr>
                                        <td colspan="2">
                                            <input type="submit" value="Save Admin Messages" class="button fit"></td>
                                    </tr>
                                </tfoot>
                            </tbody>
                        </table>
                    </form>
                </div>
            </div>
            <div class="12u$">
                <hr class="major" style="margin: 1em 0;" />
            </div>
            <h3>User Messages</h3>
            <div class="12u$" style="padding-bottom: 10px;">
                <div class="table-wrapper">
                    <form action="admin_options.asp" method="post">
                        <input type="hidden" name="chumsg" value="y">
                        <table>
                            <tbody>
                                <%
    Set rsCommon = Server.CreateObject("ADODB.Recordset")
    strSQL = "SELECT * FROM " & msdbprefix & "endMsg"

    Call getTextRecordset(strSQL, rsCommon)
    If NOT rsCommon.EOF Then
        Do While Not rsCommon.EOF
            strEndMsg = DBDecode(rsCommon("endMsg"))
            strEndMsgName = rsCommon("endMsgName")
                                %>
                                <tr>
                                    <td style="width: 30%;">
                                        <%= msgTrans(strEndMsgName) %>
                                    </td>
                                    <td style="width: 70%;">
                                        <textarea name="usermessages[<%= strEndMsgName %>]"><%= strEndMsg %></textarea>
                                    </td>
                                </tr>
                                <%
            rsCommon.MoveNext
            If rsCommon.EOF Then Exit DO
        Loop
    End If
    Call CloseRecordset(rsCommon)
                                %>
                                <tfoot>
                                    <tr>
                                        <td colspan="2">
                                            <input type="submit" value="Save User Messages" class="button fit">
                                            Use: #EMAIL# for users email address. Use: #SITETITLE# to insert your sites title.
                                        </td>
                                    </tr>
                                </tfoot>
                            </tbody>
                        </table>
                    </form>
                </div>
            </div>
        </div>
        <div class="6u$ 12u$(medium)">
            <h3>SMTP Settings <span style="font-size:16px;">+ 2</span></h3>
            <%

    Set rsCommon = Server.CreateObject("ADODB.Recordset")
    strSQL = "SELECT * FROM " & msdbprefix & "settings"

    Call getTextRecordset(strSQL, rsCommon)
    If Not rsCommon.EOF Then

        strSiteTitle = DBDecode(rsCommon("site_title"))
        blnRewrite = rsCommon("rewrite")
        smtpServer = DBDecode(rsCommon("smtp_server"))
        smtpPort = rsCommon("smtp_port")
        smtpEmail = DBDecode(rsCommon("email_address"))
        smtpPassword = DBDecode(rsCommon("smtp_password"))

        If blnRewrite Then
            strUrlchecked = "checked"
            strUrlonoff = "On"
        Else
            strUrlchecked = ""
            strUrlonoff = "Off"
        End If

        If smtpDebug = "yes" Then
            strDebchecked = "checked"
            strDebonoff = "On"
        Else
            strDebchecked= ""
            strDebonoff = "Off"
        End If

        If smtpUse = "yes" Then
            strUsechecked = "checked"
            strUseonoff = "On"
        Else
            strUsechecked = ""
            strUseonoff = "Off"
        End If

    End If
    Call closeRecordset(rsCommon)

    Call ConnClose(Conn)

    strConfirmationEmail = Replace(rsCommon("confirm_email"),vbcrlf,"")

            %>
            <div class="row">
                <div class="12u$">
                    <div class="table-wrapper">
                        <form action="admin_options.asp" method="post">
                            <input type="hidden" name="chmstgs" value="y">
                            <table>
                                <tbody>
                                    <tr>
                                        <td style="width: 30%;">Site Title
                                        </td>
                                        <td style="width: 70%;">
                                            <input type="text" name="sitetitle" value="<%= strSiteTitle %>">
                                        </td>
                                    </tr>
                                    <tr>
                                      <td style="width:30%;">
                                        URL Rewrite
                                      </td>
                                      <td style="width:70%;">
                                        <input type="checkbox" id="urlrewrite" name="urlrewrite" <%= strUrlchecked %> >
                                        <label for="urlrewrite"><span style="font-size:1.2em;font-weight:bold;"><%= strUrlonoff  %></span></label>
                                      </td>
                                    </tr>
                                    <tr>
                                        <td style="width: 30%;">SMTP Server
                                        </td>
                                        <td style="width: 70%;">
                                            <input type="text" name="smtpserver" value="<%= smtpServer %>">
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="width: 30%;">SMTP Port
                                        </td>
                                        <td style="width: 70%;">
                                            <div class="select-wrapper" style="width: 85px;">
                                                <select name="smtpport" style="width: 80px;">
                                                    <option value="25" <% if smtpPort = "25" Then Response.Write "selected" %>>25</option>
                                                    <option value="80" <% if smtpPort = "80" Then Response.Write "selected" %>>80</option>
                                                    <option value="465" <% if smtpPort = "465" Then Response.Write "selected" %>>465</option>
                                                    <option value="587" <% if smtpPort = "587" Then Response.Write "selected" %>>587</option>
                                                    <option value="2525" <% if smtpPort = "2525" Then Response.Write "selected" %>>2525</option>
                                                </select>
                                            </div>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="width: 30%;">SMTP Email address
                                        </td>
                                        <td style="width: 70%;">
                                            <input type="text" name="smtpemail" value="<%= smtpEmail %>">
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="width: 30%; vertical-align: middle;">SMTP Password
                                        </td>
                                        <td style="width: 70%; vertical-align: middle;">
                                            <div class="input-wrapper-alt">
                                                <input type="password" id="smtppwd" name="smtppassword" value="<%= smtpPassword %>"><br />
                                                <i id="shpwd" onclick="togglePass('smtppwd','shpwd')" style="cursor: pointer;" class="fa fa-eye-slash shpwd"></i>
                                            </div>
                                        </td>
                                    </tr>
                                    <tfoot>
                                        <tr>
                                            <td colspan="2">
                                                <input type="submit" value="Save Settings" class="button fit"></td>
                                        </tr>
                                    </tfoot>
                                </tbody>
                            </table>
                        </form>
                    </div>
                </div>
                <div class="12u$">
                    <hr class="major" style="margin: 1em 0;" />
                </div>
                <div class="12u$">
                    <div class="row">
                        <div class="6u 12u$(medium)"><h3>Confirmation Email</h3></div>
                        <div class="6u$ 12u$(medium)"><a class="button picimg" href="#notice" style="font-size:12px;float:right;">HELP</a></div>
                    </div>
                    
                    <form action="admin_options.asp" method="post">
                        <input type="hidden" name="confirmemail" value="yes" />
                        <textarea name="tempbody" id="tempbody" rows="25" wrap="soft"></textarea>
                        <script>
                            CKEDITOR.replace( 'tempbody', {
                            height: 250,
                            customConfig: '<%= strDIR %>assets/js/email-config.js'
                            });
                            CKEDITOR.instances.tempbody.setData('<%= Trim(strConfirmationEmail) %>');
                        </script>
                        <input class="button fit" type="submit" name="submit" value="Save Email" style="margin-top:10px;" />
                    </form>
                </div>
                <div class="12u$">
                    <hr class="major" style="margin: 1em 0;" />
                </div>
                <div class="12u$">
                    <h3>Attachments</h3>
                    <form action="uploadattach.asp?dla=no&p=o" method="post" enctype="multipart/form-data">
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
                    <label for="selectattach">Delete Attachment(s)</label>
                    <form action="uploadattach.asp?dla=yes&p=o" method="post">
                        <div class="row">
                            <div class="7u 12u$(medium)">
                                <select name="selectattach" id="selectattach" size="3" style="height: auto;" multiple>
                                    <%
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	strSaveFile = ""
	strSaveFile = Server.MapPath(strDir&"admin\attachs")

	If Not objFSO.FolderExists(strSaveFile) Then
		objFSO.CreateFolder(strSaveFile)
	End If

	strPathInfo = strSaveFile

	Set objFolder = objFSO.GetFolder(strPathInfo)
	Set objFolderContents = objFolder.Files

	If objFolder.Files.Count <> 0 Then
		For each objFileItem In objFolderContents
			Response.Write "<option value="""&objFileItem.Name&""">"&objFileItem.Name&"</option>"&vbcrlf
		Next
	Else
	    Response.Write "<option value="""">No Attachments</option>"&vbcrlf
	End If

	Set objFSO = Nothing
                                    %>
                                </select>
                            </div>
                            <div class="5u$ 12u$(medium)">
                                <input class="button fit" type="submit" name="submit" value="Delete Attachment(s)" />
                            </div>
                        </div>
                    </form>
                </div>
                <div class="12u$">
                    <hr class="major" style="margin: 1em 0;" />
                </div>
                <div class="12u$">
                    <h3>URL Rewrite</h3>
                    <span>If you have modRewrite for either Windows or Unix you can add the code below to the appropriate file then upload it to the root folder of your website. if you are using a folder other than "newsletter" you will need to change it in the files.</span>
                    <br />
                    <br />
                    <span><strong>Windows:</strong><br />
                        Create a file called web.config add the code below.</span>
                    <pre>
            <code>
&lt;?xml version="1.0" encoding="UTF-8"?&gt;
&lt;configuration&gt;
  &lt;system.webServer&gt;
    &lt;rewrite&gt;
      &lt;rules&gt;
        &lt;rule name="Rewrite confirm to friendly URL"&gt;
          &lt;match url="^confirm/([\~\-\.@_0-9a-z-]+)" /&gt;
          &lt;action type="Rewrite" url="/newsletter/includes/process.asp?token={R:1}&amp;amp;confirm=yes" /&gt;
        &lt;/rule&gt;
        &lt;rule name="Rewrite remove to friendly URL"&gt;
          &lt;match url="^remove/([\~\-\.@_0-9a-z-]+)" /&gt;
          &lt;action type="Rewrite" url="/newsletter/includes/process.asp?token={R:1}&amp;amp;cancel=yes" /&gt;
        &lt;/rule&gt;
        &lt;rule name="Rewrite subscribe to friendly URL"&gt;
          &lt;match url="^subscribe" /&gt;
          &lt;action type="Rewrite" url="/newsletter/newsletter.asp" /&gt;
        &lt;/rule&gt;
        &lt;rule name="Rewrite signup to friendly URL"&gt;
          &lt;match url="^signup" /&gt;
          &lt;action type="Rewrite" url="/newsletter/includes/process.asp?confirm=no" /&gt;
        &lt;/rule&gt;
        &lt;rule name="Rewrite unsubscribe to friendly URL"&gt;
          &lt;match url="^unsubscribe" /&gt;
          &lt;action type="Rewrite" url="/newsletter/remove.asp" /&gt;
        &lt;/rule&gt;
        &lt;rule name="Rewrite thankyou to friendly URL"&gt;
          &lt;match url="^thankyou" /&gt;
          &lt;action type="Rewrite" url="/newsletter/includes/process.asp?thank=you" /&gt;
        &lt;/rule&gt;
      &lt;rules&gt;
    &lt;/rewrite&gt;
  &lt;/system.webServer&gt;
&lt;/configuration&gt;
            </code>
          </pre>
                    <span><strong>Unix, et al:</strong><br />
                        Create a file called .htaccess add the code below.</span>
                    <pre>
            <code>
RewriteEngine on
RewriteRule ^confirm/([\~\-a-z-]+) /newsletter/includes/process.aspp?token=$1&amp;confirm=yes
RewriteRule ^remove/([\~\-a-z-]+) /newsletter/includes/process.asp?token=$1&amp;cancel=yes
RewriteRule ^subscribe /newsletter/newsletter.asp
RewriteRule ^unsubscribe /newsletter/remove.asp
RewriteRule ^thankyou /newsletter/includes/process.asp?thank=you
            </code>
          </pre>
            </div>
        </div>
    </div>
    </div>
</div>
<div style="display:none;max-width:600px;" id="notice">
    <h2>Help!</h2>
    <p>
        Use the following snippets:
        <ul>
            <li>#SITETITLE# - Sites Title</li>
            <li>#EMAIL# - Subscribers email</li>
            <li>#CR# - For the &copy; symbol</li>
            <li>#YEAR# - For the current year</li>
            <li>#CONFIRMREWRITE# - Confirmation link if Rewrite is enabled</li>
            <li>#CONFIRMNOREWRITE# - Confirmation link if Rewrite is disabled</li>
            <li>#CANCELREWRITE# - Cancellation link if Rewrite is enabled</li>
            <li>#CANCELNOREWRITE# - Cancellation link if Rewrite is disabled</li>
        </ul>
        NOTE: links will be auto generated depending on the snippet.
    </p>
</div>
<!-- #include file="../includes/footer.asp"-->