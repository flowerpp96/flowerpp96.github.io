
<HTML>
<HEAD>
   <TITLE>��ѯ�������</TITLE>
</HEAD>
<BODY>
 
<p align="center">
<%

'��Recordset �����HTML �ġ����
Sub RsToTable( rs )

  ' Part I��������ݿ�ġ���ͷ��
  If rs.EOF Then
       Response.Write "<Center><P><H1>�Բ���!û�в�ѯ����Ҫ����ù�!</H1></p></Center>"
       Exit Sub
  End If
  Response.Write "<CENTER><TABLE background=#eeeeee cellSpacing=0 width=""100%"">"
  Response.Write "<TR BGCOLOR=#999999>"
  For i=1 to rs.Fields.Count-1
     Response.WRITE "<TD>" & rs.Fields(i).Name & "</TD>"
  Next
  Response.Write "</TR>"

  ' Part II��������ݿ�ġ����ݡ�
  j=0
  While Not rs.EOF 
     j=j+1
     If j MOD 2=0 Then
         Response.Write "<TR BGCOLOR=#00ffff>"
     Else 
         Response.Write "<TR BGCOLOR=#ffffff>"
     End If
     For i=1 to rs.Fields.Count-1
        If i=1 Then
             Response.WRITE "<TD><A Href=dingdan.asp?ID=" & rs.Fields(0).Value & ">" & rs.Fields(i).Value & "</A></TD>"
        Else 
             Response.WRITE "<TD>" & rs.Fields(i).Value & "</TD>"
        End If
     Next
     Response.Write "</TR>"
     rs.MoveNext
   Wend
   Response.Write "</TABLE></CENTER>"
End Sub
%>
<b><font size="7" face="����" color="#800080">��ѯ���</font></b>
<p></p>
 
<%
select case Request.form("R1")
    case "V2"
         classstr="Class>=2"
    case "V3"
         classstr="Class>=3"
    case "V4"
         classstr="Class>=4"
    case "V5"
         classstr="Class=5"
end select

select case Request.form("R2")
    case "H0"
         coststr=""
    case "H1"
         coststr=" AND Cost>=100 AND Cost<400"
    case "H2"
         coststr=" AND Cost>=400 AND Cost<700"
    case "H3"
         coststr=" AND Cost>=700 AND Cost<1100"
    case "H4"
         coststr=" AND Cost>=1100"
end select
           
SQL="SELECT * FROM Table1 WHERE City='" & Request.Form("City") & "'"  _
    & " AND " & classstr   & coststr   
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("hotel.mdb")

Set Rs=conn.Execute(SQL)
RsToTable Rs
%>
</body>
</html>


