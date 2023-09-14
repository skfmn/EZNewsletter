<!-- #include file="../includes/general_includes.asp"-->
<%

	strCookies = Request.Cookies("Admin")("name")

	If strCookies = "" Then
		Response.Redirect "login.asp"
	End If

	If Not blnAddresses Then
	    Response.Cookies("msg") = "nar"
	    Response.Redirect "admin.asp"
	End If

    msg = ""
    msg = Trim(Request.Cookies("msg"))

	If msg <> "" Then
		Call displayFancyMsg(getMessage(msg))
        Response.Cookies("msg") = ""
	End If

	If Trim(Request.Form("ae")) = "y" Then
        If Trim(Request.Form("email")) <> "" Then

            strEmail = Trim(Request.Form("email"))

		    Set Conn = Server.CreateObject("ADODB.Connection")
            Call ConnOpen(Conn)

			Set rsCommon = server.CreateObject("ADODB.Recordset")
			strSQL = "SELECT * FROM "&msdbprefix&"addresses WHERE email = '"&DBEncode(strEmail)&"'"

			Call getTextRecordset(strSQL,rsCommon)

			If Not rsCommon.EOF Then
	            Response.Cookies("msg") = "aid"
			Else

				strSQL = "INSERT INTO "&msdbprefix&"addresses([email],[datDate],[confirm],[token]) Values('"&DBEncode(strEmail)&"','"&Date&"','yes','"&getGuid&"')"
				Call getExecuteQuery(strSQL)

	            Response.Cookies("msg") = "eas"

			End If
			Call closeRecordset(rsCommon)
			Call ConnClose(Conn)

		Else

	      Response.Cookies("msg") = "nea"

		End If

		Response.Redirect "admin_addresses.asp"

	End If

	If Trim(Request.Form("da")) = "y" Then

		Set Conn = Server.CreateObject("ADODB.Connection")
		Call ConnOpen(Conn)

		strEmail = Trim(Request.Form("email"))

	    If InStr(strEmail,",") > 0 Then
			strSplitEmail = Split(strEmail,",")
			For Each strItem in strSplitEmail

				Response.Write Trim(strItem)
				strSQL = "DELETE FROM "&msdbprefix&"addresses WHERE email = '"&Trim(DBEncode(strItem))&"'"
				Call getExecuteQuery(strSQL)

			Next

		Else

			strSQL = "DELETE FROM "&msdbprefix&"addresses WHERE email = '"&Trim(DBEncode(strEmail))&"'"
			Call getExecuteQuery(strSQL)

		End If
		Call ConnClose(Conn)

	    Response.Cookies("msg") = "ds"
		Response.Redirect "admin_addresses.asp"

	End If

    If Request.QueryString("p") <> "" Then

		todaysDate =DateAdd("d",-7,Date)
		strPurge = Request.QueryString("p")

		Set Conn = Server.CreateObject("ADODB.Connection")
        Call ConnOpen(Conn)

		If strPurge = "s" Then

	        strSQL = "DELETE FROM "&msdbprefix&"addresses WHERE datDate < '"&todaysDate &"' and confirm = 'no'"
	        Call getExecuteQuery(strSQL)

		    Response.Cookies("msg") =  "ds"

		Else 

	        strSQL = "DELETE FROM "&msdbprefix&"addresses WHERE confirm = 'no'"
	        Call getExecuteQuery(strSQL)

		    Response.Cookies("msg") =  "ds"

		End If

		Response.Redirect "admin_addresses.asp"

    End If

%>
<!-- #include file="../includes/header.asp"-->
<div id="main" class="container">
	<header>
		<h2>Manage Addresses</h2>
	</header>
    <div class="row">
        <div class="6u 12u(medium)">
            <div class="12u$" style="padding-bottom: 10px;">
                <label for="viewemail">View Addresses in Your List</label>
                <textarea id="viewemail" rows="2">
<%
	intCount = 0
	Set Conn = Server.CreateObject("ADODB.Connection")
	Call ConnOpen(Conn)

	Set rsCommon = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT * FROM "&msdbprefix&"addresses WHERE confirm = 'yes'"

	Call getTextRecordset(strSQL,rsCommon)
	If Not rsCommon.EOF Then
	    intCount = rsCommon.RecordCount
		Do While Not rsCommon.EOF

			Response.Write DBDecode(rsCommon("email")) & vbcrlf

			rsCommon.MoveNext
	       If rsCommon.EOF Then Exit Do
		Loop
	Else
	    Response.Write "No Addresses"
	End If
	Call closeRecordset(rsCommon)
	Call ConnClose(Conn)

	%>
                </textarea>
            </div>
			<div class="12u$ 12u$(small)" style="padding-bottom: 10px;">
                <span>There are <%= intCount %> members in your mailing list.</span>
            </div>
        </div>
		<div class="6u$ 12u(medium)">
			<form action="admin_addresses.asp" method="post">
				<input type="hidden" name="ae" value="y" />
				<div class="row uniform">
					<div class="12u$" style="padding-bottom: 10px;">
						<label for="email">Add an Email Address</label>
						<input id="email" name="email" type="text" required>
					</div>
					<div class="12u$" style="padding-bottom: 10px; text-align: center;">
						<input class="button fit" name="submit" type="submit" value="Add Address">
					</div>
				</div>
			</form>
        </div>
		<div class="6u 12u(medium)">
			<form action="admin_addresses.asp" method="post">
				<input type="hidden" name="da" value="y" />
			<div class="row uniform">
				<div class="12u$" style="padding-bottom: 10px;">
					<label for="email">Select addresses to delete</label>
					<div class="select-wrapper">
					<% Call selectDeleteEmail %>
					</div>
				</div>
				<div class="12u$" style="padding-bottom: 10px; text-align: center;">
					<input class="button fit" type="submit" value="Delete" onclick="return confirm('WARNING!!\n Are you sure you want to delete these addresses?\n This cannot be undone!')">
				</div>
			</div>
		    </form>
		</div>
        <div class="6u$ 12u$(medium)">
            <div class="row uniform">
                <div class="12u$" style="padding-bottom:10px;">
                    <label>Purge Unconfirmed email addresses</label>
                    <input type="button" onclick="return confirmSubmit('WARNING!!\n Are you sure you want to delete these unconfirmed addresses older than a week?\n This cannot be undone!','admin_addresses.asp?p=s')" class="button fit" value="Purge Unconfirmed > 7 days" />
                </div>
                <div class="12u$" style="padding-bottom:10px;text-align:center;">
                    <input type="button" onclick="return confirmSubmit('WARNING!!\n Are you sure you want to delete all unconfirmed addresses?\n This cannot be undone!','admin_addresses.asp?p=a')" class="button fit" value="Purge all Unconfirmed" />
                </div>
            </div>
        </div>
    </div>
</div>
<!-- #include file="../includes/footer.asp"-->