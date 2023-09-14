<%
    Response.Cookies("Admin").Expires = Date - 1
	Session("loggedin") = ""
	Session.Abandon()
	Response.Redirect "login.asp"
%>
