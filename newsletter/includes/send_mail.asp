<%
  Function send_email(sFrom,sTo,sSubject,sBody,sAttach,iHTML)
	
	  Dim cdoMessage, cdoConfig, strSendMsg
		strError = "No Error"

		Set cdoMessage = Server.CreateObject("CDO.Message")
		Set cdoConfig = Server.CreateObject("CDO.Configuration")
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtpServer
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = smtpPort
		
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = smtpEmail
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = smtpPassword
		
		cdoConfig.Fields.Update
		Set cdoMessage.Configuration = cdoConfig
		cdoMessage.From =  smtpEmail
		If Trim(sFrom) <> "" Then cdoMessage.ReplyTo =  sFrom
		cdoMessage.To = sTo
		cdoMessage.Subject = sSubject
		If sAttach <> "" Then
			cdoMessage.AddAttachment Server.MapPath(strDir&sAttach)
		End If
		If Cint(iHTML) = 1 Then
			cdoMessage.HtmlBody = sBody
		Else
			cdoMessage.TextBody = sBody
		End If
		
		on error resume next
	 
		cdoMessage.Send
	 
		If Err.Number <> 0 Then
			strSendMsg = "Email send failed: " & Err.Description &vbcrlf
		Else
		    strSendMsg = "Newsletter Sent to "&sTo&"!"&vbcrlf
		End If
		
		Set cdoMessage = Nothing
		Set cdoConfig = Nothing

        Response.Write strSendMsg

	End Function

    Function send_email_proccess(sFrom,sTo,sSubject,sBody,sAttach,iHTML)
	
	  Dim cdoMessage, cdoConfig, strSendMsg
		strError = "No Error"

		Set cdoMessage = Server.CreateObject("CDO.Message")
		Set cdoConfig = Server.CreateObject("CDO.Configuration")
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtpServer
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = smtpPort
		
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = smtpEmail
		cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = smtpPassword
		
		cdoConfig.Fields.Update
		Set cdoMessage.Configuration = cdoConfig
		cdoMessage.From =  smtpEmail
		If Trim(sFrom) <> "" Then cdoMessage.ReplyTo =  sFrom
		cdoMessage.To = sTo
		cdoMessage.Subject = sSubject
		If sAttach <> "" Then
			cdoMessage.AddAttachment Server.MapPath(strDir&sAttach)
		End If
		If Cint(iHTML) = 1 Then
			cdoMessage.HtmlBody = sBody
		Else
			cdoMessage.TextBody = sBody
		End If
		
		on error resume next
	 
		cdoMessage.Send
	 
		If Err.Number <> 0 Then
			strSendMsg = "failed"
		Else
		    strSendMsg = "success"
		End If
		
		Set cdoMessage = Nothing
		Set cdoConfig = Nothing

	    send_email_proccess = strSendMsg

	End Function
%>
