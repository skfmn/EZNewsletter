<!-- #include file="../includes/general_includes.asp"-->
<%
	strCookies = Request.Cookies("Admin")("name")
	
	If strCookies = "" Then

		Response.Redirect "login.asp"
  
	End If

	If Not blnImages Then

	    Response.Cookies("msg") = "nar"
	    Response.Redirect "admin.asp"

	End If

    msg = ""
    msg = Trim(Request.Cookies("msg"))

	If msg <> "" Then
		Call displayFancyMsg(getMessage(msg))
        Response.Cookies("msg") = ""
	End If

	Set objFSO = CreateObject("Scripting.FileSystemObject")

	strSaveFile = ""
	strSaveFile = Server.MapPath(strDir&"admin\images")

	If Not objFSO.FolderExists(strSaveFile) Then
		objFSO.CreateFolder(strSaveFile)
	End If
	
	strPathInfo = strSaveFile
	
	Set objFolder = objFSO.GetFolder(strPathInfo)
	Set objFolderContents = objFolder.Files

	counter = 0
	intCounter = 0

%>
<!-- #include file="../includes/header.asp"-->
<div id="main" class="container">
  <div class="row">
		<div class="-2u 8u 12u(medium)">
      <form action="upload.asp?deleteimg=yes" method="post">
        <h4>Images</h4>
        <div class="table-wrapper">
	        <table>
		        <thead>
			        <tr>
				        <th style="width:20px;">&nbsp;</th>
				        <th>Image</th>
				        <th>Size</th>
                        <th style="text-align:right;">Delete</th>
			        </tr>
		        </thead>
		        <tbody>
  <%   
	  For each objFileItem In objFolderContents 
	    If objFileItem.Name <> "Thumbs.db" Then
		
			  counter = counter + 1
			
			  If Bgcolor = "silver" Then
				  Bgcolor = "gray"
			  Else
				  Bgcolor = "silver"
			  End if
  %>
              <tr>
                <td><%= counter %>.</td>
                <td style="text-align:left;"><a class="picimg" href="http://<%= strDomain&strDir %>admin/images/<%= objFileItem.Name %>"><%= objFileItem.Name %></a></td>
                <td><span style="font-size:16px"><%= ConvBytes(objFileItem.Size) %></span></td>
                <td style="text-align:right;">
                  <input type="checkbox" id="file<%= counter %>" name="file<%= counter %>" value="<%= objFileItem.Name %>" style="z-index:1000"/>
                  <label for="file<%= counter %>">Yes</label>
                </td>
              </tr>
  <%
      End If
    Next

    Set objFSO = Nothing
  %>
            </tbody>
            <tfoot>
              <tr>
                <td colspan="4" style="text-align:center;">
                  <input type="submit" value="Delete Selected Images">
                </td>
              </tr>
			      </tfoot>
          </table>
        </div>
      </form>
    </div>
  </div>

  <div class="row">
		<div class="-2u 8u 12u(medium)">
      <h4>Upload Images</h4>
      <form name="upldfile" action="upload.asp?imgupld=yes" method="post" enctype="multipart/form-data">
      <input type="file" name="FILE" size="20" multiple>
      <input type="submit" value="Upload Images">
      </form>
    </div>
  </div>
</div>
<!-- #include file="../includes/footer.asp"-->