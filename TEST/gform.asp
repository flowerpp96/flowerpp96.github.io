<%@ Language=VBScript %>
<%
	function sqlstr(data)
		sqlstr="'" & replace(data,"'","''") & "'"
	end function
	
	name=left(request("name"),20)
	add=left(request("add"),20)
	memo=request("memo")
	
	set conn=server.CreateObject("ADODB.connection")
	dbpath=server.MapPath("gbook.mdb")
	conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & dbpath
	sql="insert into guestbook (name,add,memo) values("
	sql=sql & sqlstr(name) & ","
	sql=sql & sqlstr(add) & ","
	sql=sql & sqlstr(memo) & ")"
	conn.Execute sql
	Response.redirect "mousheng1.asp"
%>
