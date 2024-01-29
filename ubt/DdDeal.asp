<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 1</title>
</head>

<body>
<%
ID=Request("ID")
OtherHotel=Request("OtherHotel")
FromYear=Request("FromYear")
FromMonth=Request("FromMonth")
FromDay=Request("FromDay")
FormTime=DateSerial( FromYear, FromMonth , FromDay )
EndYear=Request("EndYear")
EndMonth=Request("EndMonth")
EndDay=Request("EndDay")
FormTime=DateSerial( EndYear, EndMonth , EndDay )


Response.Write(FormTime) 'FromYear & "年" & FromMonth & "月”& FromDay)
 %>

</body>

</html>
