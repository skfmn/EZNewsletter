<script>
function selectText(containerid) {
    if (document.selection) { // IE
        var range = document.body.createTextRange();
        range.moveToElementText(document.getElementById(containerid));
        range.select();
    } else if (window.getSelection) {
        var range = document.createRange();
        range.selectNode(document.getElementById(containerid));
        window.getSelection().removeAllRanges();
        window.getSelection().addRange(range);
    }
}
</script>
<!-- #include file="../includes/general_includes.asp"-->
<%
    'Dim strBkcolor
	strCookies = Request.Cookies("Admin")("name")
	
	If strCookies = "" OR NOT blnImages Then

	    Response.Redirect "login.asp"
  
	End If

	Set objFSO = CreateObject("Scripting.FileSystemObject")

	strPathInfo = Server.MapPath(strDir&"admin/images")

	Set objFolder = objFSO.GetFolder(strPathInfo)
	intFileCount = objFolder.Files.count
	If intFileCount = 0 Then
	    Response.Write "No Images in Folder!"
	Else
        Response.Write "<div style=""min-width:600px"">"
        Response.Write "   <div style=""position:relative;display:block;float:left;"">"
	    intCounter = 0
	    strBkcolor = "#e2effc"
	    For Each x In objFolder.Files

	          intCounter = intCounter+1

              If strBkcolor = "#e2effc" Then
                strBkcolor = "#ffffff"
              Else 
                strBkcolor = "#e2effc"
              End If

              Response.Write "<span id=""img-"&intCounter&""" onclick=""selectText('img-"&intCounter&"')"" onmouseover=""document.getElementById('place-holder-1').src='"&strHTTP&strDomain&strDir&"admin/images/"& x.Name&"'"""
              Response.Write " onmouseout=""document.getElementById('place-holder-1').src=''""; style=""background-color:"&strBkcolor&""">"&strHTTP&strDomain&strDir&"admin/images/"& x.Name&"</span><br>"
	    Next

	    Response.Write "   </div>"
        Response.Write "   <div style=""position:relative;display:block;float:left;margin-left:10px;"">"
        Response.Write "      <img src="""" id=""place-holder-1"" />"
        Response.Write "   </div>"
        Response.Write "<div>"
	End If
	
	Set objFolder = Nothing
	Set objFSO = Nothing
%>
