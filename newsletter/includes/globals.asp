<%
    Dim Conn, ConnStr, rsCommon, strSQL, strDomain, smtpServer, cdoMessage, cdoConfig, addrList, intSend
    Dim fromAddr, strSiteTitle, strFooter, recipients, intCount, strEmailMsg, subject, strEmail, strAttach
	Dim strCKSDPWD, Addrss, Addr, msg, strConfirm, strDir, strASPUpload, intDCount, intTCount, strConfirmEmail
	Dim strSplitEmail, strItem, strFolder, strSaveFile, Upload, FName, objFSO, File, strConfirmationEmail, strUserMessage 
	Dim strPathInfo, objFolder, objFolderContents, Bgcolor, counter, intCounter, objFileItem, strVersion, strChecked
	Dim strTitle, strLoad, strTempTitle, strTempBody, lngTempID, intFileCount, strRedirect, mtabs, strHTTP, strTempSave
	Dim intTid, lngMemberID, lngTempplateID, smtpPassword, strTempName, intAdminID,  smtpEmail, smtpPort, strAdminName
	Dim blnSend, blnAddresses, blnImages,blnTemplates, blnOptions, blnAdminRights, blnARights, blnAspupload, blnRewrite

	msg = Trim(Request.QueryString("msg"))
	mtabs = Trim(Request.QueryString("mtabs"))
	
	Response.ExpiresAbsolute = Now() - 2
	Response.AddHeader "pragma","no-cache"
	Response.AddHeader "cache-control","private"
	Response.CacheControl = "No-Store"
	Response.AddHeader "If-Modified-Since",now
	Response.AddHeader "Last-Modified",now
	Response.Expires = 0

	If Request.ServerVariables("HTTPS") = "off" Then
	    strHTTP = "http://"
	Else
	    strHTTP = "https://"
	End If
%>
