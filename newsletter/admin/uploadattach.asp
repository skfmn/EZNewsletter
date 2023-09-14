<!-- #include file="../includes/general_includes.asp"-->
<!-- #include file="../includes/_upload.asp"-->
<%
on error resume next
	Dim DestinationPath, Form

	strCookies = Request.Cookies("Admin")("name")
	
	If strCookies = "" OR NOT blnImages Then

	    Response.Redirect "login.asp"
  
	End If

    If Trim(Request.QueryString("dla")) = "yes" Then

		strSaveFile = ""
		strSaveFile = Server.MapPath(strDir&"admin\attachs")

		Set objFSO = CreateObject("Scripting.FileSystemObject")
	    nums = Request.Form("selectattach").Count
		For i = 1 To nums
		    objFSO.DeleteFile strSaveFile&"\"&Request.Form("selectattach")(i)
		Next

		Set objFSO = Nothing
		
	    Response.Cookies("msg") = "ids"

	ElseIf Trim(Request.QueryString("dla")) = "no" Then

		If blnAspupload Then

			Set Upload = Server.CreateObject("Persits.Upload.1")
	  
			Upload.SaveVirtual strDir&"admin/attachs"
		  
			For each File in Upload.Files
				FName = fileExt(File.Path) 
				File.SaveAsVirtual strDir&"admin/attachs"&"/"&FName
			Next
	 
			Set Upload = Nothing

		Else

			Set Form = New ASPForm
			DestinationPath = Server.mapPath(strDir&"admin/attachs")

			Form.SizeLimit = &H100000

			If Form.State = 0 Then
				Form.Files.Save DestinationPath
			End If'

		End If

	    Response.Cookies("msg") = "uls"

    End If

	If Request.QueryString("p") = "o" Then
	    Response.Redirect "admin_options.asp"
	Else
	    Response.Redirect "admin_send.asp"
	End If
%>