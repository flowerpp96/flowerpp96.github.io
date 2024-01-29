
<html>
<head>
<%@ Language=VBScript %> <%    

	sub rstogbook(rs)
		   Tabb="<tr class=bgc2 > "
           Tabb=Tabb & " <td width=19 valign='top'>" & rs("id") & "</td>"
           Tabb=Tabb & " <td width=77 valign='top'>" & rs("name") & "</td>"
           Tabb=Tabb & " <td width=77 valign='top'>" & rs("add") & "</td>"
           Tabb=Tabb & " <td width=134 align=""center""><textarea rows=""10"" name=""memo"" cols=""15"" class=""input1"">" & rs("memo") & "</textarea></td>"
            'Tabb=Tabb & " <td width=134 align=""center""><textarea rows=""12"" name=""memo"" cols=10 class=""input1"">" & Replace(rs("memo"),VbCrlf,"<br>") & "</textarea></td>"
           Tabb=Tabb & " <td width=78 valign='top'>" & rs("datee") & "</td></tr>"
		Response.Write Tabb
	end sub
%> <!--#include file="adovbs.inc" --> <!--#include file="my.asp" --> 
<title>我们工分 - 谋生纪事</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
<!---

td,p,input,select,dl,li {font-family:宋体;font-size:12px;}
.input {   font-size:11px; color: #660066; background-color: #FFFDD7; border-color:FF0000; border: FF0000; border-style: dotted; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px}
.input1 {   font-size:10pt; color: #4685EC; background-color: #FFFDD7; border-color:0000FF; border: 0000FF; border-style: dotted; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px}
.f12 {font-family:宋体;font-size:12px;}
.f14 {font-family:宋体;font-size:14.9px;}
.f16 {font-family:宋体;font-size:16px;}
.f16a {font-size:16px;}
.f16h11 {font-family:宋体;font-size:16px;line-height:110%;}
.f16h12 {font-family:宋体;font-size:16px;line-height:120%;}

.f12cw {color:#ffffff;font-family:宋体;font-size:12px;}
.f12cb {color:#0000ff;font-family:宋体;font-size:12px;}
.f12c1 {color:#ffffff;font-family:宋体;font-size:12px;}

.f14c1 {color:#615f5f;font-weight:bold;font-family:宋体;font-size:14.9px;}
.f14c2 {color:#95381a;font-family:宋体;font-size:14.9px;}
.f16c1 {color:#615f5f;font-weight:bold;font-family:宋体;font-size:16px;}


.bgc0 {background-color:#ffcc66}
.bgc1 {background-color:#4685EC}
.bgc2 {background-color:#FDF5E1}
.bgc3 {background-color:#ffffff}
.bgc4 {background-color:#cccccc}
.bgc5 {background-color:#FBE0AA}
.bgcw {background-color:#ffffff}

.bg1 {background-repeat:no-repeat;}

a:link {text-decoration: underline; color: #6A4800;font-weight: bold}
a:visited {text-decoration: underline; color: #6A4800;font-weight: bold}
a:active {text-decoration: none}
a:hover {text-decoration: underline; color: #ff0000;font-weight: bolder}

a.aw:link {text-decoration: underline; color: #ffffff}
a.aw:visited {text-decoration: underline; color: #ffffff}
a.aw:active {text-decoration:underline;}
a.aw:hover {text-decoration: underline; color: #e7e0bc}

a.ab:link {text-decoration: none; color: #0000ff}
a.ab:visited {text-decoration: none; color: #0000ff}
a.ab:active {text-decoration: none}
a.ab:hover {text-decoration: underline; color: #ff0000}
-->
</style>
</head>

<body bgcolor="#FF0000" topmargin=0 leftmargin=0 marginwidth="0" marginheight="0">
<span class="input1"></span>
<table border=0 cellspacing=10 cellpadding=0 align="center" bgcolor="#EFFDE6" width=460>
  <!-- main title bar --> 
  <tr align="right" valign="top"> 
    <td> 
        
         <table border=0 cellspacing=0 cellpadding=0 width=387 align="center">
<tr>
<td  colspan="5">
<table width="400" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
            
          <td width="193" align="left" bgcolor="#66CCCC"><font face="黑体"><b>&nbsp;</b></font></td>
          <td align="right" width="194" bgcolor="#66CCCC"><a href="send.asp">发帖子</a></td>
          </tr>
        </table>
        <table border=0 cellspacing=0 cellpadding=0 width="400" align="center">
          <tr> 
            <td class=bgc1></td>
          </tr>
        </table>
        
</td>       
</tr>
<tr> 
            
          <td class=bgc1 width="387"></td>
          </tr>
        </table>
        
      <table width=387 border=0 cellspacing=3 cellpadding=0 align="center">
        <tr class=bgc1 > 
            
          <td height=20 width=19 nowrap><br>
            </td>
            
          <td class=f12c1 width=77 nowrap >&nbsp;姓名</td>
            
          <td class=f12c1 width=77 nowrap >&nbsp;来自</td>
            
          <td class=f12c1 width=134 nowrap >内容</td>
            
          <td class=f12c1 width=78 nowrap >留言时间</td>
          </tr>
          
          
          
<%	
sql = "Select * From GuestBook Order By datee DESC"
If Request("Page") = "" Then
   Set conn = Open_Database( Server.MapPath("gbook.mdb"), "gbook_conn")
   Set rs = Open_RsAndPageSize( conn, sql, "gbook_rs", 5 )
Else
   Set conn = OpenOrGet_Database( Server.MapPath("gbook.mdb"), "gbook_conn")
   Set rs = OpenOrGet_RsAndPageSize( conn, sql, "gbook_rs", 5 )   
End If

Page = CLng(Request("Page"))   ' CLng 不可省略
If Page < 1 Then 
		Page = 1
	elseIf Page > rs.PageCount Then 
		Page = rs.PageCount
end if
ShowOnePage rs, Page
%>



          
          
          <tr class=bgc1 > 
            
          <td height=20 width="19" nowrap><br>
            </td>
            
          <td class=f12c1 width="77" nowrap >&nbsp;姓名</td>
            
          <td class=f12c1 width="77" nowrap >&nbsp;来自</td>
            
          <td class=f12c1 width="134" nowrap >内容</td>
            
          <td class=f12c1 width="78" nowrap >留言时间</td>
          </tr>
        </table>
        <table border=0 cellspacing=0 cellpadding=0 width=400 align="center">
          <tr> 
            <td class=bgc1></td>
          </tr>
        </table>
        
      <table border=0 cellspacing=0 cellpadding=0 width=400 align="center">
        <tr> 
            
          <td class=bgc1 width="400"></td>
          </tr>
        </table>
        
        <form method="post" action="mousheng1.asp" id=form2 name=form2>
  <table width="95%" border="1" align="center">
    <tr> 
            <td nowrap valign="middle"><a href="send.asp">发帖子</A></td>

      <%
      IF PAGE>1 THEN
		Response.Write "<td  nowrap valign=""middle""><A HREF=""mousheng1.asp?page=1"">第一页</a></td>"
		Response.write "<td   nowrap valign=""middle""><a href=mousheng1.asp?page=" & (page-1) & "> 上一页</a></td>"
      end if
      if page<rs.pagecount then
		Response.Write "<td nowrap valign=""middle""><a href=mousheng1.asp?page=" & (page+1) & ">下一页</a></td>"
      		Response.Write "<td  nowrap valign=""middle""><a href=mousheng1.asp?page=" & rs.pagecount & ">最后一页</a></td>"
      end if
	%>
            <td  align="right" nowrap valign="middle">输入页码:</td>
<td  align="left" nowrap valign="middle">
              <input type="text" name="Page" size="3" class="input">
      </td>
      <td  nowrap valign="middle">页数:<font color=#ff0000><%=Page%>/<%=rs.PageCount%></font></td>
    </tr>
  </table>
</form>
<P>&nbsp;</P>    
    </td>
  </tr>
  <tr valign=top align="center"> 
    <td colspan="2" nowrap> 
      <hr>
      <span class="f12cb">Power by 俺;Build time：2000.8; Copyleft, All rights abandoned</span> 
    </td>
  </tr>
</table>
</body>
</html>