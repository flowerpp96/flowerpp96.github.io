<%
Sub ShowOnePage( rs, Page )
	if rs.BOF=true and rs.EOF=true then
		   Tabb="<tr class=bgc2 > "
           Tabb=Tabb & " <td width=19 valign='top'>尚无纪录</td>"
           Tabb=Tabb & " <td width=77 valign='top'>尚无纪录</td>"
           Tabb=Tabb & " <td width=77 valign='top'>尚无纪录</td>"
           Tabb=Tabb & " <td width=134 align=""center""><textarea rows=""10"" name=""memo"" cols=""15"" class=""input1"">尚无纪录</textarea></td>"
           Tabb=Tabb & " <td width=78 valign='top'>尚无纪录</td></tr>"
		   Response.Write Tabb
	else
		rs.AbsolutePage = Page
		For iPage = 1 To rs.PageSize
			RsToGbook rs
			rs.MoveNext
			If rs.EOF Then Exit For
		Next
	end if
End Sub

Function OpenOrGet_Database( DbPath, SessionName )
   Dim conn

   If Not IsObject(Session(SessionName)) Then
      Set conn = Server.CreateObject("ADODB.Connection")
      conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
      Set Session(SessionName) = conn
   End If
   Set OpenOrGet_Database = Session(SessionName)
End Function

Function Open_Database( DbPath, SessionName )
   Dim conn

   Set conn = Server.CreateObject("ADODB.Connection")
   conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
   Set Session(SessionName) = conn
   Set Open_Database = Session(SessionName)
End Function


Function OpenOrGet_RsAndPageSize( conn, sql, SessionName, PageSize )
   Dim rs 

   If Not IsObject(Session(SessionName)) Then
      Set rs = Server.CreateObject("ADODB.Recordset")
      rs.Open sql, conn, adOpenStatic
      Set Session(SessionName) = rs
      rs.PageSize = PageSize
   End If
   Set OpenOrGet_RsAndPageSize = Session(SessionName)
End Function

Function Open_RsAndPageSize( conn, sql, SessionName, PageSize )
   Dim rs 

   Set rs = Server.CreateObject("ADODB.Recordset")
   rs.Open sql, conn, adOpenStatic
   Set Session(SessionName) = rs
   rs.PageSize = PageSize
   Set Open_RsAndPageSize = Session(SessionName)
End Function

%>