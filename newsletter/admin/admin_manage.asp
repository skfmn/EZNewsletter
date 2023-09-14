<!-- #include file="../includes/general_includes.asp"-->
<%
	strCookies = Request.Cookies("Admin")("name")

	If strCookies = "" Then

		Response.Redirect "login.asp"

	End If

	If Not blnAdminRights Then

	    Response.Cookies("msg") = "nar"
	    Response.Redirect "admin.asp"

	End If

    msg = ""
    msg = Trim(Request.Cookies("msg"))

	If msg <> "" Then
		Call displayFancyMsg(getMessage(msg))
        Response.Cookies("msg") = ""
	End If

	If Trim(Request.Form("newadmin")) = "yes" Then

	    msg = ""
		strUsername = DBEncode(Request.Form("nwnm"))
		strPassword = Trim(Request.Form("nwpwd"))

        strSalt = CStr(getSalt(len(strPassword)))

        strEncrPassword = HashEncode(strPassword&strSalt)

		Set Conn = Server.CreateObject("ADODB.Connection")
        Call ConnOpen(Conn)

        Set rsCommon = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM "&msdbprefix&"admin WHERE name= '"&strUsername&"'"

        Call getTextRecordset(strSQL,rsCommon)
        If Not rsCommon.EOF Then
            Response.Cookies("msg") = "ant"
		Else
	        strSQL = "INSERT INTO "&msdbprefix&"admin (name,pwd,salt) VALUES ('"&strUsername&"','"&strEncrPassword&"','"&strSalt&"')"
	        Conn.Execute strSQL
			Response.Cookies("msg") = "adad"
	    End If
		Call closeRecordset(rsCommon)
		Call ConnClose(Conn)

		Response.Redirect "admin_manage.asp"

	End If

    If Trim(Request.Form("chypwd")) = "yes" Then

        Set Conn = Server.CreateObject("ADODB.Connection")
        Call ConnOpen(Conn)

        Set rsCommon = Server.CreateObject("ADODB.Recordset")
        strSQL = "SELECT * FROM "&msdbprefix&"admin WHERE adminID = "&Request.Cookies("Admin")("adminID")

        Call getTextRecordset(strSQL,rsCommon)
        If NOT rsCommon.EOF Then

            strAdminName = DBEncode(Request.Form("cname"))
            strPassword = Trim(Request.Form("cpwd"))
	        strSalt = getSalt(len(strPassword))
	        strEncrPassword = HashEncode(strPassword&strSalt)

	        strSQL = "UPDATE "&msdbprefix&"admin Set salt = '"&strSalt&"', pwd = '"&strEncrPassword&"', name='"&strAdminName&"' WHERE adminID = "&Request.Cookies("Admin")("adminID")
	        Call getExecuteQuery(strSQL)

            Response.Cookies("Admin")("name") = strAdminName

            Response.Cookies("msg") = "cpwds"
	        strRedirect = "admin_manage.asp"

        End If
        Call closeRecordset(rsCommon)
        Call ConnClose(Conn)

		Response.Redirect strRedirect

    End If

	If Trim(Request.Form("chapwd")) = "yes" Then

		Set Conn = Server.CreateObject("ADODB.Connection")
        Call ConnOpen(Conn)

        Set rsCommon = Server.CreateObject("ADODB.Recordset")
        strSQL = "SELECT * FROM "&msdbprefix&"admin WHERE name = '"&DBEncode(Request.Form("cname"))&"'"

        Call getTextRecordset(strSQL,rsCommon)

        If NOT rsCommon.EOF Then
            If rsCommon("name") <> "admin" AND LCase(Trim(Request.Form("cname"))) <> "admin" Then

                strPassword = Trim(Request.Form("cpwd"))
                strSalt = CStr(getSalt(len(strPassword)))
                strEncrPassword = CStr(HashEncode(strPassword&strSalt))

                strSQL = "UPDATE "&msdbprefix&"admin Set pwd = '"&strEncrPassword&"', salt = '"&strSalt&"' WHERE name = '"&DBEncode(Request.Form("cname"))&"'"
                Call getExecuteQuery(strSQL)

                Response.Cookies("msg") = "cpwds"
                strRedirect = "admin_manage.asp"

		    Else
                Response.Cookies("msg") = "nadmin"
		        strRedirect = "admin_manage.asp"
		    End If
	    End If
        Call closeRecordset(rsCommon)
        Call ConnClose(Conn)

        Response.Redirect strRedirect

	End If

	If Trim(Request.QueryString("deleteadmin")) = "yes" Then

         intAdminID = checkint(Request.QueryString("id"))

		Set Conn = Server.CreateObject("ADODB.Connection")
		Call ConnOpen(Conn)
	    strSQL = "DELETE FROM "&msdbprefix&"admin WHERE adminID = "&intAdminID
	    Call getExecuteQuery(strSQL)

		Call ConnClose(Conn)

        Response.Cookies("msg") = "das"
		Response.Redirect "admin_manage.asp"

	End If

%>
<!-- #include file="../includes/header.asp"-->
<div id="main" class="container">
    <header>
        <h2>Manage Admins</h2>
    </header>
    <div class="row">
        <div class="6u 12u(medium)">
            <form action="admin_manage.asp" method="post">
                <input type="hidden" name="chypwd" value="yes" />
                <div class="row uniform">
                    <div class="12u$">
                        <h3>Change YOUR Info</h3>
                        <div class="12u$" style="padding-bottom: 10px;">
                            <label for="cname">Login Name</label>
                            <input type="text" id="cname" name="cname" value="<%= Request.Cookies("Admin")("name") %>" required />
                        </div>
                        <div class="12u$" style="padding-bottom: 10px;">
                            <label for="cpwd">Password</label>
                            <div class="input-wrapper-alt">
                                <input type="password" id="cpwd" name="cpwd" required />
                                <br />
                                <i id="shpwd1" onclick="togglePass('cpwd','shpwd1')" style="cursor:pointer;" class="fa fa-eye-slash shpwd"></i>
                            </div>
                            <input class="button fit" type="submit" value="Change Password" />
                        </div>
                    </div>
                </div>
            </form>

            <form action="admin_manage.asp" method="post">
                <input type="hidden" name="chapwd" value="yes" />
                <div class="row uniform">
                    <div class="12u$">
                        <h3>Change an Admins Password</h3>
                        <div class="12u$" style="padding-bottom: 10px;">
                            <label for="cname">Login Name</label>
                            <div class="select-wrapper">
                                <select id="cname" name="cname">
                                    <%
	        Set Conn = Server.CreateObject("ADODB.Connection")
            Call ConnOpen(Conn)

            Set rsCommon = Server.CreateObject("ADODB.Recordset")
            strSQL = "SELECT adminID, name FROM "&msdbprefix&"admin"

            Call getTableRecordset(msdbprefix&"admin",rsCommon)
            If Not rsCommon.EOF Then
                Do While Not rsCommon.EOF
                    strName = DBDecode(rsCommon("name"))
                    intAdminID = rsCommon("adminID")
                    If Cint(intAdminID) <> 1 Then Response.Write "            <option value="""&strName&""">"&strName&"</option>"&vbcrlf
                    rsCommon.MoveNext
                    If rsCommon.EOF Then Exit Do
                Loop
            End If
            Call closeRecordset(rsCommon)
            Call ConnClose(Conn)
                                    %>
                                </select>
                            </div>
                        </div>
                        <div class="12u$" style="padding-bottom: 10px;">
                            <label for="cpwdb">Password</label>
                            <div class="input-wrapper-alt">
                                 <input type="password" id="cpwdb" name="cpwd" required />
                                 <br />
                                 <i id="shpwd2" onclick="togglePass('cpwdb','shpwd2')" style="cursor:pointer;" class="fa fa-eye-slash shpwd"></i>
                            </div>
                        </div>
                        <div class="12u$" style="padding-bottom: 10px; text-align: center;">
                            <input class="button fit"  type="submit" value="Change Password" />
                        </div>
                    </div>
                </div>
            </form>
        </div>
        <div class="6u$ 12u(medium)">
            <form action="admin_manage.asp" method="post">
                <input type="hidden" name="newadmin" value="yes" />
                <div class="row uniform">
                    <div class="12u$">
                    <h3>Add an Admin</h3>
                    <div class="12u$" style="padding-bottom: 10px;">
                        <label for="nwnm">Admin Name</label>
                        <input type="text" id="nwnm" name="nwnm" required />
                    </div>
                    <div class="12u$" style="padding-bottom: 10px;">
                        <label for="nwpwd">Admin Password</label>
                        <div class="input-wrapper-alt">
                            <input type="text" id="nwpwd" name="nwpwd" required />
                            <br />
                            <i id="shpwd3" onclick="togglePass('nwpwd','shpwd3')" style="cursor:pointer;" class="fa fa-eye-slash shpwd"></i>
                        </div>
                    </div>
                    <div class="12u$" style="padding-bottom: 10px; text-align: center;">
                        <input class="button fit" type="submit" value="Add An Admin" />
                    </div>
                </div>
            </form>

            

                <h3>List of Admins</h3>
                <div class="12u$" style="padding-bottom: 10px;">
                    <hr />
                    NOTE: Main Admin is not listed so you won't delete it!
                   <hr />
                </div>
                <%
		    Set Conn = Server.CreateObject("ADODB.Connection")
            Call ConnOpen(Conn)

		    Set rsCommon = Server.CreateObject("ADODB.Recordset")

            Call getTableRecordset(msdbprefix&"admin",rsCommon)
            If Not rsCommon.EOF Then
                %>
                <div class="12u$" style="padding-bottom: 10px;">
                    <div class="table-wrapper">
                        <table>
                            <tbody>
                                <%
                Do While Not rsCommon.EOF
			        If Cint(rsCommon("adminID")) <> 1 Then
                                %>
                                <tr>
                                    <td>
                                        <%= rsCommon("name") %>&nbsp;&nbsp;<a onclick="return confirmSubmit('Are you SURE you want to delete this admin?','admin_manage.asp?deleteadmin=yes&id=<%= rsCommon("adminID") %>')" style="cursor: pointer; text-decoration: underline;">Delete</a>
                                        <% If blnARights Then %>
                      &nbsp;&nbsp;<a href="arights.asp?id=<%= rsCommon("adminID") %>" title="<%= DBDecode(rsCommon("name")) %>">Assign Rights</a>
                                        <% End If %>
                                    </td>
                                </tr>
                                <%
                    End If
	                rsCommon.MoveNext
                    If rsCommon.EOF Then Exit Do
	            Loop
                                %>
                            </tbody>
                        </table>
                    </div>
                    <%
        End If
        Call closeRecordset(rsCommon)
        Call ConnClose(Conn)
                    %>
                </div>
            </div>
        </div>
    </div>
</div>
<!-- #include file="../includes/footer.asp"-->