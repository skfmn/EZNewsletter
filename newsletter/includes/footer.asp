<!-- Footer -->
<footer id="footer">
    <div class="copyright">
        <a href="http://www.aspjunction.com">EZNewsletter</a> Copyright &copy; 2003 - <%= Year(Date) %> <a href="http://www.aspjunction.com">ASP junction</a>  All rights reserved.
    </div>
</footer>

<!-- Scripts -->
<script src="https://code.jquery.com/jquery-1.12.4.js"></script>
<script src="https://code.jquery.com/ui/1.12.1/jquery-ui.js"></script>
<script src="../assets/js/jquery.fancybox.js"></script>
<script src="../assets/js/skel.min.js"></script>
<script src="../assets/js/main.js"></script>
<script src="../assets/js/javascript.js"></script>
<script>
  $( function() {
    var availableTags = [
<%
on error resume next
    Set Conn = Server.CreateObject("ADODB.Connection")
    Call ConnOpen(Conn)

    Set rsCommon = Server.CreateObject("ADODB.Recordset")

    Call getTableRecordset(msdbprefix&"newsletter",rsCommon)

    intCount = 0
    intRecordCount = 0
    If Not rsCommon.EOF Then
        intRecordCount = rsCommon.RecordCount
        Do While Not rsCommon.EOF

            strNewsTitle = DBDecode(rsCommon("news_title"))

            intCount = intCount+1
            Response.Write "        """&strNewsTitle&""""

            rsCommon.MoveNext
            If rsCommon.EOF OR Cint(intCount) = Cint(intRecordCount) Then
                Exit Do
            Else
                Response.Write ","&vbcrlf
            End If
        Loop
    End If
    Call closeRecordset(rsCommon)
    Call ConnClose(Conn)
%>
    ];
    $( "#temptitle" ).autocomplete({
      source: availableTags
    });
  } );
</script>
</body>
</html>