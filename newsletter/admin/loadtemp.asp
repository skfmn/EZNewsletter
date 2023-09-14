<!-- #include file="../includes/general_includes.asp"-->
<%
	strCookies = Request.Cookies("Admin")("name")
	
	If strCookies = "" OR NOT blnTemplates Then

		Response.Redirect "login.asp"
  
	End If

  Set Conn = Server.CreateObject("ADODB.Connection")
  Call ConnOpen(Conn)

	Set rsCommon = Server.CreateObject("ADODB.Recordset")
  strSQL = "SELECT * FROM "&msdbprefix&"newsletter WHERE news_save = 'template' OR news_save = 'both'"

  intCount = 0

	Call getTextRecordset(strSQL,rsCommon)
	If Not rsCommon.EOF Then
    intRecordCount = rsCommon.RecordCount
%>
CKEDITOR.addTemplates( 'default',
{
	templates :
		[
<%
    Do While Not rsCommon.EOF
      strNewsBody = Replace(Server.URLEncode(DBDecode(rsCommon("news_body"))),"%","\x")
      strNewsBody = Replace(strNewsBody,"+"," ")
      intCount = intCount+1 
%>
			{
        title: '<%= DBDecode(rsCommon("news_title")) %>',
        description: '<%= DBDecode(rsCommon("news_description")) %>',
        html:
					'<%= strNewsBody %>'
			}
<%
      rsCommon.MoveNext
      If rsCommon.EOF Or Cint(intCount) = Cint(intRecordCount) Then
        Exit Do
      Else
        Response.Write ","&vbcrlf
      End If
    Loop  
%>
		]
});
<%
	End If
	Call closeRecordset(rsCommon)
	Call ConnClose(Conn)
%>