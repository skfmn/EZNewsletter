<%

    strVersion = "4.0.3"

    strSaveFile = Request.ServerVariables("APPL_PHYSICAL_PATH")&strFolder&"images"
	
    ConnStr = "Provider=sqloledb;Data Source="&msdbserver&";Initial Catalog="&msdb&";User Id="&msdbid&";Password="&msdbpwd

    Set Conn = Server.CreateObject("ADODB.Connection")
    Call ConnOpen(Conn)

    Set rsCommon = Server.CreateObject("ADODB.Recordset")
  
    Call getTableRecordset(msdbprefix&"settings",rsCommon)
    If Not rsCommon.EOF Then
		strSiteTitle = DBDecode(rsCommon("site_title"))
		strDomain = DBDecode(rsCommon("domain_name"))
		smtpServer = DBDecode(rsCommon("smtp_server"))
	    smtpPort = rsCommon("smtp_port")
		smtpEmail = DBDecode(rsCommon("email_address"))
		smtpPassword = DBDecode(rsCommon("smtp_password"))
	    blnRewrite = rsCommon("rewrite")
	    blnAspupload = rsCommon("aspupload")
	    strConfirmationEmail = Trim(rsCommon("confirm_email"))
    End If
    Call closeRecordset(rsCommon)
    Call ConnClose(Conn)

	If Request.Cookies("Admin")("adminID") <> "" Then
	    Call getMyInfo(Request.Cookies("Admin")("adminID"))
	End If

	Sub getMyInfo(lMemberID)
on error resume next
		blnSend = false
		blnAddresses = false
		blnImages = false
		blnTemplates = false
	    blnOptions = false
		blnDBRights = false
		blnAdminRights = false
		blnARights = false

		If Session("loggedin") = "" Then
			Session("loggedin") = "yes"

			Set Conn = Server.CreateObject("ADODB.Connection")
			Call ConnOpen(Conn)

			Set rsCommon = Server.CreateObject("ADODB.Recordset")
			strSQL = "SELECT * FROM "&msdbprefix&"admin WHERE adminID = "&lMemberID

			Call getTextRecordset(strSQL,rsCommon)
			If Not rsCommon.EOF Then
				blnSend = rsCommon("send")
				blnAddresses = rsCommon("addresses")
				blnImages = rsCommon("images")
				blnTemplates = rsCommon("templates")
	            blnOptions = rsCommon("options")
				blnAdminRights = rsCommon("admins_rights")
				blnARights = rsCommon("arights")
			End If
			Call closeRecordset(rsCommon)
			Call ConnClose(Conn)

			Session("blnSend") = blnSend
			Session("blnAddresses") = blnAddresses
			Session("blnImages") = blnImages
			Session("blnTemplates") = blnTemplates
	        Session("blnOptions") = blnOptions
			Session("blnAdminRights") = blnAdminRights
			Session("blnARights") = blnARights

		Else

			blnSend = Session("blnSend")
			blnAddresses = Session("blnAddresses")
			blnImages = Session("blnImages")
			blnTemplates = Session("blnTemplates")
	        blnOptions = Session("blnOptions")
			blnAdminRights = Session("blnAdminRights")
			blnARights = Session("blnARights")

		End If

	End Sub
  

	Function getResponse(sURL)
		Dim strTemp
		strTemp = ""

		Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
		xmlhttp.SetOption(2) = (xmlhttp.GetOption(2) - SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS)
		xmlhttp.Open "GET", sURL, false
		xmlhttp.Send(NULL)
		If xmlhttp.readyState = 4 Then
			If xmlhttp.status <> 200 Then

			    strTemp = "<span style=""color:#FF0000"">Error "&xmlhttp.status&" - "&xmlhttp.statusText&"</span><br />"

			Else

			    strTemp = xmlhttp.ResponseText

			End If
		End If
		Set xmlhttp = Nothing

		getResponse = strTemp

	End Function

    Sub selectEmailAddy

	    Set Conn = Server.CreateObject("ADODB.Connection")
        Call ConnOpen(Conn)

        Set rsCommon = Server.CreateObject("ADODB.Recordset")
        strSQL = "SELECT * FROM "&msdbprefix&"addresses WHERE confirm = 'yes';"

        Call getTextRecordset(strSQL,rsCommon)
		If Not rsCommon.EOF Then
			Response.Write "<select id=""viewemail"" name=""viewemail"">"
			Do While Not rsCommon.EOF
				Response.Write "<option value="""&DBDecode(rsCommon("email"))&""">"&DBDecode(rsCommon("email"))&"</option>"
				rsCommon.MoveNext
				If rsCommon.EOF Then Exit Do
			Loop
			Response.Write "</select>"
		Else
			Response.Write "<select id=""viewemail"" name=""viewemail""><option>No Addresses</option></select>"
		End If
		Call closeRecordset(rsCommon)
		Call ConnClose(Conn)

    End Sub

    Sub selectDeleteEmail

        Set Conn = Server.CreateObject("ADODB.Connection")
        Call ConnOpen(Conn)

        Set rsCommon = Server.CreateObject("ADODB.Recordset")
        strSQL = "SELECT * FROM "&msdbprefix&"addresses WHERE confirm = 'yes';"

        Call getTextRecordset(strSQL,rsCommon)
        If Not rsCommon.EOF Then
            Response.Write "<select id=""email"" name=""email"">"
            Do While Not rsCommon.EOF
	            Response.Write "<option value="""&DBDecode(rsCommon("email"))&""">"&DBDecode(rsCommon("email"))&"</option>"
	            rsCommon.MoveNext
		       If rsCommon.EOF Then Exit Do
            Loop
	        Response.Write "</select>"
        Else
            Response.Write "<select id=""email"" name=""email""><option>No Addresses</option></select>"
        End If
		Call closeRecordset(rsCommon)
		Call ConnClose(Conn)

    End Sub

	Sub selectLoadTemplate

		strLoad = "no"
		strCount = 0
		Set Conn = Server.CreateObject("ADODB.Connection")
		Call ConnOpen(Conn)

		Set rsCommon = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM "&msdbprefix&"newsletter WHERE news_save = 'template'"

		Call getTextRecordset(strSQL,rsCommon)
		If Not rsCommon.EOF Then
			Response.Write "<select name=""loadtemp"" id=""loadtemp"" onchange=""return loadTemplate(this.options[this.selectedIndex].value);"">"&vbcrlf
			Response.Write "  <option value=""0"">Load Template</option>"&vbcrlf
			Do While Not rsCommon.EOF
				Response.Write "  <option value="""&rsCommon("newsletterID")&""">" & DBDecode(rsCommon("news_title")) & "</option>"&vbcrlf
				rsCommon.MoveNext
                If rsCommon.EOF Then Exit Do
			Loop
            Response.Write "</select>"&vbcrlf
			strLoad = "yes"
		Else
			Response.Write "<select>"&vbcrlf
			Response.Write "<option>No Templates</option>"&vbcrlf
			Response.Write "</select>"&vbcrlf
		End If
		Call closeRecordset(rsCommon)
		Call ConnClose(Conn)

	End Sub

	Sub selectLoadDraft

		strLoad = "no"
		strCount = 0
		Set Conn = Server.CreateObject("ADODB.Connection")
		Call ConnOpen(Conn)

		Set rsCommon = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM "&msdbprefix&"newsletter WHERE news_save = 'draft'"

		Call getTextRecordset(strSQL,rsCommon)
		If Not rsCommon.EOF Then
			Response.Write "<select name=""loadtemp"" id=""loadtemp"" onchange=""return loadTemplate(this.options[this.selectedIndex].value);"">"&vbcrlf
			Response.Write "  <option value=""0"">Load Draft</option>"&vbcrlf
			Do While Not rsCommon.EOF
				Response.Write "  <option value="""&rsCommon("newsletterID")&""">" & DBDecode(rsCommon("news_title")) & "</option>"&vbcrlf
				rsCommon.MoveNext
                If rsCommon.EOF Then Exit Do
			Loop
            Response.Write "</select>"&vbcrlf
			strLoad = "yes"
		Else
			Response.Write "<select>"&vbcrlf
			Response.Write "<option>No Drafts</option>"&vbcrlf
			Response.Write "</select>"&vbcrlf
		End If
		Call closeRecordset(rsCommon)
		Call ConnClose(Conn)

	End Sub

    Function getMessage(sMsg)
	    Dim strTemp: strTemp = ""

        Set Conn = Server.CreateObject("ADODB.Connection")
        Call ConnOpen(Conn)

	    Set rsCommon = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM "&msdbprefix&"messages WHERE msg = '"&sMsg&"'"

		Call getTextRecordset(strSQL,rsCommon)
		If Not rsCommon.EOF Then
		  strTemp = DBDecode(rsCommon("message"))
		Else
		  strTemp = sMsg
		End If
		Call closeRecordset(rsCommon)
		Call ConnClose(Conn)

		getMessage = strTemp

	End Function

    Function getUserMessage(sMsg)
	    Dim strTemp: strTemp = ""

        Set Conn = Server.CreateObject("ADODB.Connection")
        Call ConnOpen(Conn)

	    Set rsCommon = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT * FROM "&msdbprefix&"endMsg WHERE endMsgName = '"&sMsg&"'"

		Call getTextRecordset(strSQL,rsCommon)
		If Not rsCommon.EOF Then
		  strTemp = DBDecode(rsCommon("endMsg"))
		Else
		  strTemp = sMsg
		End If
		Call closeRecordset(rsCommon)
		Call ConnClose(Conn)

		getUserMessage = strTemp

	End Function

	Function msgTrans(sMsg)

		Dim strTemp: strTemp = ""

		Select  Case sMsg
			case "eas"
				strTemp = "Email address added:"
			case "aid"
				strTemp = "Address already in DB:"
			case "ds"
				strTemp = "Delete action successful:"
			case "nea"
				strTemp = "Forgot email address:"
			case "uls"
				strTemp = "File uploaded:"
			case "ids"
				strTemp = "File deleted:"
			case "nadmin"
				strTemp = "Can't change Admin info:"
			case "del"
				strTemp = "Template deleted:"
			case "das"
				strTemp = "Deleted an Admin:"
			case "adad"
				strTemp = "Added an Admin:"
			case "nt"
				strTemp = "Name taken:"
			case "tc"
				strTemp = "Template created:"
			case "car"
				strTemp = "Changed Admin Rights:"
			case "ulf"
				strTemp = "Upload failed:"
			case "nwst"
				strTemp = "Newsletter sent:"
			case "ant"
				strTemp = "Admin name taken:"
			case "confirmed"
				strTemp = "Email confirmed:"
			case "confirmerr"
				strTemp = "Confirm. email error:"
			case "ftna"
				strTemp = "Allowable extensions:"
			case "fex"
				strTemp = "File Exists:"
			case "nimg"
				strTemp = "Not an image:"
			case "tus"
				strTemp = "Template Updated:"
			case "adderr"
				strTemp = "Couldn't add to list:"
			case "thanks"
				strTemp = "Successfully added:"
			case "canceled"
			    strTemp = "Subscriber removed:"
			case "thankserr"
				strTemp = "Added - problem sending email:"
			case "alreadysubbed"
				strTemp = "Already Subscribed:"
			case "removed"
				strTemp = "Removed from list error:"
			case "notfound"
				strTemp = "Email not found:"
			case "removederr"
				strTemp = "Couldn't remove from list:"
			case "mus"
				strTemp = "Messages updated:"
			case "error"
				strTemp = "Generic error:"
			case "siu"
				strTemp = "Site info updated:"
			case "cpwds"
				strTemp = "Changed your password:"
			case "nar"
				strTemp = "No Admin Rights:"
		End Select

		msgTrans = strTemp

	End Function

	Sub sendMail

		msg = ""
		IntCount = 0

		subject = Trim(Request.Form("nwsubject"))
		strEmailMsg = Trim(Request.Form("strmsg"))
		If Request.Form("version") = "html" Then
			strEmailMsg = strEmailMsg & strHFooter
		Else
			strEmailMsg = strEmailMsg & strTFooter
		End If

		IntCount = 0
		Set Conn = Server.CreateObject("ADODB.Connection")
		Set rsNews = Server.CreateObject("ADODB.Recordset")
		Call ConnOpen(Conn)
		Call getTableRecordset("addresses",rsNews)
		rsNews.Filter = "confirm = 'yes'"
		If Not rsNews.EOF Then
			Do While Not rsNews.EOF
				intCount = intCount+1
				strEmail = rsNews("email")

				Set cdoMessage = Server.CreateObject("CDO.Message")
				Set cdoConfig = Server.CreateObject("CDO.Configuration")
				cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
				cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = smtpServer
				cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = smtpPort
				cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
				cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = fromAddr
				cdoConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = smtpPassword
				cdoConfig.Fields.Update
				Set cdoMessage.Configuration = cdoConfig
				cdoMessage.From =  "" & strSiteTitle & " <" & fromAddr & ">"
				cdoMessage.To = strEmail
				cdoMessage.Subject = subject
				If Request.Form("version") = "html" Then
					cdoMessage.HtmlBody = strEmailMsg
				Else
					cdoMessage.TextBody = strEmailMsg
				End If
				On Error Resume Next
				cdoMessage.Send
				If Err.Number <> 0 Then
					msg = "Email send failed: " & Err.Description & "."
				Else
					msg = "Email sent"
				End If
				Set cdoMessage = Nothing
				Set cdoConfig = Nothing

				rsNews.MoveNext
				If rsNews.EOF Then Exit Do
			Loop
		End If
		Call closeRecordset(rsNews)
		Call ConnClose(Conn)

		Response.Redirect "admin.asp?msg="&msg

	End Sub

	Function DBEncode(DBvalue)
		Dim fieldvalue
		fieldvalue = DBvalue

		If fieldvalue <> "" AND Not IsNull(fieldvalue) Then
		
			Set encodeRegExp = New RegExp
			encodeRegExp.Pattern = "((delete)*(select)*(update)*(into)*(drop)*(insert)*(declare)*(xp_)*(union)*)"
			encodeRegExp.IgnoreCase = True
			encodeRegExp.Global = True
			Set Matches = encodeRegExp.Execute(fieldvalue)
			For Each Match In Matches
				fieldvalue = Replace(fieldvalue,Match.Value,StrReverse(Match.Value))
			Next
			fieldvalue=replace(fieldvalue,"'","''")

		End If

		DBEncode = fieldvalue

	End Function

	Function DBDecode(DBvalue)
		Dim fieldvalue
	    fieldvalue = DBvalue

		If fieldvalue <> "" AND ( NOT IsNull(fieldvalue) ) Then

			Set encodeRegExp = New RegExp
			encodeRegExp.Pattern = "((eteled)*(tceles)*(etadpu)*(otni)*(pord)*(tresni)*(eralced)*(_px)*(noinu)*)"
			encodeRegExp.IgnoreCase = True
			encodeRegExp.Global = True
			Set Matches = encodeRegExp.Execute(fieldvalue)
			For Each Match In Matches
				fieldvalue = Replace(fieldvalue,Match.Value,StrReverse(Match.Value))
			Next
			fieldvalue = replace(fieldvalue,"''","'")
		End If

	DBDecode = fieldvalue

	End Function

	Function displayFancyMsg(sText)
%>
<div style="display: none">
    <a id="textmsg" href="#displaymsg">Message</a>
    <div id="displaymsg" style="width: 300px;">
        <h2 style="text-align: left;">Message</h2>
        <div style="text-align: center;">
            <span style="color: #FF0000;"><%= sText %></span>
        </div>
        <div class="left_menu_bottom"></div>
    </div>
</div>
</div>
<%
	End Function

	Function fileExt(s)
	    fileExt = Right(s,Len(s)-(inStrRev(s,"\",-1,1)))
	End Function

	Function ConvBytes(TBytes)
		Dim inSize, isType
		Const lnBYTE = 1
		Const lnKILO = 1024
		Const lnMEGA = 1048576
		Const lnGIGA = 1073741824               
		Const lnTERA = 1099511627776

		If TBytes < 0 Then Exit Function

		If TBytes < lnKILO Then
			inSize = TBytes
			isType = "bytes"
		Else

			If TBytes < lnMEGA Then
			    inSize = (TBytes / lnKILO)
			    isType = "kb"
			ElseIf TBytes < lnGIGA Then
			    inSize = (TBytes / lnMEGA)
			isType = "mb"
			    ElseIf TBytes < lnTERA Then
			inSize = (TBytes / lnGIGA)
			    isType = "gb"
			Else
			    inSize = (TBytes / lnTERA)
			    isType = "tb"
			End If

		End If
	 
		inSize = FormatNumber(inSize,2)

		ConvBytes = inSize & " " & isType

	End Function

	Function chkEmail(sEmail)
		Set objRegExp = New RegExp
		searchStr = "^[\w-\.]{1,}\@([\da-zA-Z-]{1,}\.){1,}[\da-zA-Z-]{2,3}$"
		objRegExp.Pattern = searchStr
		objRegExp.IgnoreCase = true
		chkEmail = objRegExp.Test(sEmail)
	End Function

	Function getGuid
	    Dim TypeLib : Set TypeLib = CreateObject("Scriptlet.TypeLib")
	    getGuid = Mid(TypeLib.Guid, 2, 15)
		getGuid = Replace(getGuid,"-","")
		Set TypeLib = Nothing
	End Function
	

	Function checkInt(iVal)
        Dim intTemp
		intTemp = 0

	    If iVal <> "" Then
			If Not IsNumeric(iVal) Then
				Call displayFancyMsg("Input was not a number!")
			Else
			    intTemp = iVal
		    End If
		Else
		    Call displayFancyMsg(txtInputWasEmpty&"Input was empty!")
		End If

		checkInt = intTemp

	End Function

	Sub trace(strText)
		Response.Write "Debug: "&strText&"<br />"&vbcrlf
	End Sub

	Sub catch(sText,sText2)

		If Err.Number <> 0 then
		    Call trace(sText&" - "&err.description)
		Else
		    Call trace(sText&" - no error")
		End If

		If sText2 <> "" Then
		    Call trace(sText&" - "&sText2)
		End If

		on error goto 0

	End Sub
%>