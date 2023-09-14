<!-- #include file="general_includes.asp"-->
<!DOCTYPE HTML>
<html>
<head>
<title>EZNewsletter</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<table width="100%" height="600px" cellpadding="0" cellspacing="0" border="0">
  <tr>
    <td width="25%" valign="top">&nbsp;</td>
	<td width="50%" align="center" valign="middle" style="padding:10px">
<%
on error resume next	

	strEmail = Trim(Request("email"))
	strToken = Trim(Request.QueryString("token"))
	strConfirm = Trim(Request("confirm"))

	Set Conn = Server.CreateObject("ADODB.Connection")
	Call ConnOpen(Conn)

    If strConfirm = "yes" Then 'confirm the subscription

		strSQL = "UPDATE "&msdbprefix&"addresses SET confirm = '"&strConfirm&"' WHERE token = '"&UCase(strToken)&"'"  
		Call getExecuteQuery(strSQL)

		Set rsCommon = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM "&msdbprefix&"addresses WHERE token = '"&UCase(strToken)&"'" 

	    Call getTextRecordset(strSQL,rsCommon)
		If Not rsCommon.EOF Then
	        strEmail = rsCommon("email")
		End If
		Call closeRecordset(rsCommon)

	    strUserMessage = ""
	    strUserMessage = getUserMessage("confirmed")
	    strUserMessage = Replace(strUserMessage,"#EMAIL#",strEmail)
	    strUserMessage = Replace(strUserMessage,"#SITETITLE#",strSiteTitle)

	    Response.Write strUserMessage


    ElseIf strConfirm = "no" Then 'New subscriber

	    strEmail = Trim(Request.Form("email"))
		
		Set rsCommon = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM "&msdbprefix&"addresses WHERE email = '"&strEmail&"'"

	    Call getTextRecordset(strSQL,rsCommon)
	    If Not rsCommon.EOF Then

	        strUserMessage = ""
			strUserMessage = getUserMessage("alreadysubbed")
			strUserMessage = Replace(strUserMessage,"#EMAIL#",strEmail)
			strUserMessage = Replace(strUserMessage,"#SITETITLE#",strSiteTitle)

			Response.Write strUserMessage


	    Else

	        strToken = getGuid
	        strEmailMsg = ""
		    strSQL = "INSERT INTO "&msdbprefix&"addresses(email,datDate,confirm,token) VALUES('"&strEmail&"','"&Now&"','"&strConfirm&"','"&strToken&"')"
		    Call getExecuteQuery(strSQL)
 
			subject = strSiteTitle & " Newsletter confirmation"
			
			strEmailMsg = Replace(strConfirmationEmail,"#SITETITLE#",strSiteTitle)
	        strEmailMsg = Replace(strEmailMsg,"#EMAIL#",strEmail)
	        strEmailMsg = Replace(strEmailMsg,"#CR#","&copy;")
	        strEmailMsg = Replace(strEmailMsg,"#YEAR#",Year(Date))
	        strEmailMsg = Replace(strEmailMsg,"#CONFIRMREWRITE#","<a href=""" & strHTTP & strDomain & "/confirm/" & UCase(strToken) & "/yes"">Confirm</a>")
	        strEmailMsg = Replace(strEmailMsg,"#CONFIRMNOREWRITE#","<a href=""" & strHTTP  & strDomain & strDir & "includes/process.asp?token=" & UCase(strToken) & "&confirm=yes"">Confirm</a>")
	        strEmailMsg = Replace(strEmailMsg,"#CANCELREWRITE#","<a href=""" & strHTTP  & strDomain & "/remove/" & UCase(strToken) & """>Cancel</a>")
	        strEmailMsg = Replace(strEmailMsg,"#CANCELNOREWRITE#","<a href=""" & strHTTP  & strDomain & strDir & "includes/process.asp?token=" & UCase(strToken) & "&cancel=yes"">Cancel</a>")
	
			strResponse = send_email_proccess("",strEmail,subject,strEmailMsg,"",1)

			If strResponse = "success" Then

		        strUserMessage = ""
	            strUserMessage = getUserMessage("thanks")
	            strUserMessage = Replace(strUserMessage,"#EMAIL#",strEmail)
	            strUserMessage = Replace(strUserMessage,"#SITETITLE#",strSiteTitle)

	            Response.Write strUserMessage

			Else

	            strUserMessage = ""
	            strUserMessage = getUserMessage("adderr")
	            strUserMessage = Replace(strUserMessage,"#EMAIL#",strEmail)
	            strUserMessage = Replace(strUserMessage,"#SITETITLE#",strSiteTitle)

	            Response.Write strUserMessage

			End If
		End If
		Call closeRecordset(rsCommon)

    Else

		If Request("cancel") = "yes" Then

			Set rsCommon = Server.CreateObject("ADODB.Recordset")
			strSQL = "SELECT * FROM "&msdbprefix&"addresses WHERE token = '"&UCase(strToken)&"'" 

			Call getTextRecordset(strSQL,rsCommon)
			If not rsCommon.EOF Then
	            strEmail = rsCommon("email")
				rsCommon.Delete

	            strUserMessage = ""
	            strUserMessage = getUserMessage("canceled")
	            strUserMessage = Replace(strUserMessage,"#EMAIL#",strEmail)
	            strUserMessage = Replace(strUserMessage,"#SITETITLE#",strSiteTitle)

	            Response.Write strUserMessage

			Else

	            strUserMessage = ""
	            strUserMessage = getUserMessage("notfound")
	            strUserMessage = Replace(strUserMessage,"#EMAIL#",strEmail)
	            strUserMessage = Replace(strUserMessage,"#SITETITLE#",strSiteTitle)

	            Response.Write strUserMessage

			End If
			Call closeRecordset(rsCommon)

        End If
	End If
%>
	  </td>
	  <td width="25%" valign="top">&nbsp;</td>
	</tr>
  </table>
</body>
</html>