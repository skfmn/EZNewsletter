<%
  Sub ConnOpen(Conn)
    Conn.Open ConnStr
	End Sub
	
	Sub ConnClose(Conn)
	  Conn.Close: Set Conn = Nothing
	End Sub

	Sub getPageRecordset(strSQL, rsObject) 
		rsObject.Open strSQL, Conn, 3, 1, &H0002
	End Sub
	
	Sub getPageTextRecordset(strSQL, rsObject) 
		rsObject.Open strSQL, Conn, 3, 2, &H0001
	End Sub
	
	Sub getPTextRecordset(strSQL, rsObject) 
		rsObject.Open strSQL, Conn, 3, 2, &H0001
	End Sub

	Sub getTextRecordset(strSQL, rsObject)
    rsObject.CursorLocation =  3
		rsObject.Open strSQL, Conn, 3, 3, &H0001
	End Sub
	
	Sub getTableRecordset(strSQL, rsObject) 
		rsObject.Open strSQL, Conn, 3, 3, &H0002
	End Sub
	
	Sub getExecuteQuery(strSQL) 
		Conn.Execute(strSQL)
	End Sub
	
	Sub closeRecordset(rsObject)
		rsObject.Close: Set rsObject = Nothing
	End Sub
	
	Sub closeObject(oObject)
		oObject.Close: Set oObject = Nothing
	End Sub
%>
