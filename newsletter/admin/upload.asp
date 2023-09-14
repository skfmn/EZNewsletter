<!-- #include file="../includes/general_includes.asp"-->
<!-- #include file="../includes/_upload.asp"-->
<%
on error resume next
	Dim DestinationPath, Form

	strCookies = Request.Cookies("Admin")("name")
	
	If strCookies = "" OR NOT blnImages Then

	    Response.Redirect "login.asp"
  
	End If

	If Trim(Request.QueryString("imgupld")) = "yes" Then
	
	    If blnAspupload Then

			Set Upload = Server.CreateObject("Persits.Upload.1")
	  
			Upload.SaveVirtual strDir&"admin/images"
		  
			For each File in Upload.Files
				FName = fileExt(File.Path) 
				File.SaveAsVirtual strDir&"admin/images"&"/"&FName
			Next
	 
			Set Upload = Nothing

		Else

			Set Form = New ASPForm
			DestinationPath = Server.mapPath(strDir&"admin/images")

			Form.SizeLimit = &H100000

			If Form.State = 0 Then
			    Form.Files.Save DestinationPath
			End If'

	    End If

	    Response.Cookies("msg") = "uls"
		strRedirect = "admin_images.asp"

    Elseif Trim(Request.QueryString("deleteimg")) = "yes" Then

		strSaveFile = ""
		strSaveFile = Server.MapPath(strDir&"admin\images")

		Set objFSO = CreateObject("Scripting.FileSystemObject")
		For each strItem in Request.Form
		    objFSO.DeleteFile strSaveFile&"\"&Request.Form(strItem).Item
		Next
		Set objFSO = Nothing
		
	    Response.Cookies("msg") = "ids"
		strRedirect = "admin_images.asp"

    End If
	
	Response.Redirect strRedirect
%>