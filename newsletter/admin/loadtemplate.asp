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

	strTempTitle = ""
	strTempBody = ""
	strTempDescr = ""
	lngTempID = 0
	intTid = 0
	intTid = checkint(Trim(Request.QueryString("tid")))

  Set Conn = Server.CreateObject("ADODB.Connection")
  Call ConnOpen(Conn)

	Set rsCommon = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT * FROM "&msdbprefix&"newsletter WHERE newsletterID = "&intTid

	Call getTextRecordset(strSQL,rsCommon)
	If Not rsCommon.EOF Then
		strTempTitle = DBDecode(rsCommon("news_title"))
	    strTempDescr = DBDecode(rsCommon("news_description"))
		strTempBody = DBDecode(rsCommon("news_body"))
		lngTempID = rsCommon("newsletterID")
	End If
	Call closeRecordset(rsCommon)
	Call ConnClose(Conn)
	
	strTempBody = Replace(strTempBody,",","~")
	strTempDescr = Replace(strTempDescr,",","~")
	Response.Write strTempTitle&","&strTempDescr&","&strTempBody&","&lngTempID
 
%>