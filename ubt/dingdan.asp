<% ID=Request("ID")
            
 SQL="Table1"  
 Set conn = Server.CreateObject("ADODB.Connection")   
 conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("hotel.mdb")   
   
 Set Rs=conn.Execute(SQL) 
 For i=1  to ID-1 
       Rs.MoveNext 
 Next 
             
 %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE></TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<META content="Microsoft FrontPage 4.0" name=GENERATOR>

</HEAD>
<BODY leftMargin=5 topMargin=0>
<TABLE align=center border=0>
   <TBODY>
   <TR><TD align=middle noWrap vAlign=top width=740>
        <p align="center"><font size="7" face="隶书" color="#808000">订单</font></p>
   </TD></TR></TBODY></TABLE>
<TABLE border=0 cellPadding=2 cellSpacing=1 height=353  width="100%">    
  <FORM action=DdDeal.asp method=post name=InputData>
     <%
      Response.Write "<INPUT name=ID type=hidden value=" & ID & ">"  
     %>    
  <TBODY>    
  <TR>    
    <TD align=middle height=288 vAlign=top width="100%">    
      <DIV align=center>    
      <CENTER>    
      <TABLE align=center background=Hotels_Form.files/cui-net.htm border=0     
      cellSpacing=1 height=370 width="599">    
        <TBODY>    
        <TR>    
          <TD align=middle bgColor=#ffb38e colSpan=2 height=22 vAlign=center     
          width="591">
            <p align="center"><font size="4" color="#008080">您好，欢迎您!</font></p>
          </TD>    
        </CENTER>    
        <TR>    
          <TD align=center bgColor=#99ccff height=19     
            width="134">
            <p align="left">您选择酒店情况:</p>
          </TD>    
      <CENTER>    
          <TD align=left bgColor=#99ccff height=19     
            width="451">   
           <%   
              'RS.AbsolutePosition ID 
              Response.Write  Rs.Fields(1).Value & ",星级:"& Rs.Fields(3).Value &  ",城市:"& RS.Fields(2).Value &  _
                        ",门市价:" & Rs.Fields(4).Value &",提供价:"& Rs.Fields(5).Value    
           %>   
        <TR>    
           <TD align=center bgColor=#99ccff height=25  width="134">备选酒店: </TD>    
           <TD bgColor=#99ccff height=25 vAlign=center width="451">&nbsp;       
           <INPUT name=OtherHotel size=18></TD></TR>      
        <TR>      
          <TD align=center bgColor=#99ccff height=25 width="134">入住时间:</TD>      
          <TD bgColor=#99ccff height=25 width="451">&nbsp;      
          <SELECT name=FromYear size=1> 
               <%  yx=Year(Date) 
                   Response.Write "<OPTION SELECTED>" & yx     
                   Response.Write "<OPTION>" & yx+1     
                   Response.Write "<OPTION>" & yx+2 & "</OPTION>" 
               %>    
          </SELECT> 年      
          <SELECT name=FromMonth size=1>     
               <% For mx=1 to 12     
                    If mx=Month(Date) Then    
                        Response.Write "<OPTION SELECTED>" & mx      
                    Else    
                        Response.Write "<OPTION>" & mx     
                    End If    
                  Next    
                  Response.Write "</OPTION>"     
               %>    
          </SELECT> 月     
          <SELECT name=FromDay size=1>    
               <% For dx=1 to 31     
                    If dx=Day(Date) Then    
                        Response.Write "<OPTION SELECTED>" & dx      
                    Else    
                        Response.Write "<OPTION>" & dx     
                    End If    
                  Next    
                  Response.Write "</OPTION>"     
               %>       
        </SELECT>日&nbsp;<FONT color=#ff0000>***</FONT></TD></TR>              
        <TR>              
          <TD align=center bgColor=#99ccff height=25  width="134">退房时间:</TD>             
          <TD bgColor=#99ccff height=25 width="451">&nbsp;      
          <SELECT  name=EndYear size=1>     
              <%  yx=Year(Date) 
                   Response.Write "<OPTION SELECTED>" & yx     
                   Response.Write "<OPTION>" & yx+1     
                   Response.Write "<OPTION>" & yx+2 & "</OPTION>" 
               %>    
          </SELECT> 年      
          <SELECT name=EndMonth size=1>     
                <% For mx=1 to 12     
                    If mx=Month(Date+1) Then    
                        Response.Write "<OPTION SELECTED>" & mx      
                    Else    
                        Response.Write "<OPTION>" & mx     
                    End If    
                  Next    
                  Response.Write "</OPTION>"     
                %>    
          </SELECT> 月      
          <SELECT name=EndDay size=1>     
               <% For dx=1 to 31     
                    If dx=Day(Date+1) Then    
                        Response.Write "<OPTION SELECTED>" & dx      
                    Else    
                        Response.Write "<OPTION>" & dx     
                    End If    
                  Next    
                  Response.Write "</OPTION>"     
               %>       
          </SELECT>日&nbsp;<FONT color=#ff0000>***</FONT></TD></TR>              
        <TR>              
          <TD align=center bgColor=#99ccff height=25  width="134">订房间数: </TD>             
          <TD bgColor=#99ccff height=25 vAlign=center width="451">&nbsp;               
          <INPUT name=RoomNum size=3>  间<FONT color=#ff0000>***</FONT></TD></TR>             
        <TR>             
          <TD align=center bgColor=#99ccff height=25  width="134">联系人姓名:</TD>             
          <TD bgColor=#99ccff height=25 width="451">&nbsp;      
          <INPUT name=LinkmanName size=15>　<FONT color=#ff0000>***</FONT></TD></TR>
        <TR>
          <TD align=center bgColor=#99ccff height=25 width="134">联系人电话:</TD>
          <TD bgColor=#99ccff height=25 width="451">&nbsp;      
          <INPUT name=LinkmanTele size=30><FONT color=#ff0000>（***区号必填）</FONT></TD></TR>              
        <TR>              
          <TD align=center bgColor=#99ccff height=25  width="134">其它联系方法:</TD>             
          <TD bgColor=#99ccff height=25 width="451">&nbsp;     
          <INPUT  name=OtherLink size=40 bgcolor="#99ccff"></TD></TR>             
        <TR>             
          <TD align=center bgColor=#99ccff height=25  width="134">入住人姓名:</TD>             
          <TD bgColor=#99ccff height=25 width="451">&nbsp;      
          <INPUT name=FormalName size=15>&nbsp; <FONT color=#ff0000>***</FONT></TD></TR>             
        <TR>             
          <TD align=center bgColor=#99ccff height=25  width="134">入住人电话:</TD>             
          <TD bgColor=#99ccff height=25 width="451">&nbsp;    
          <INPUT name=FormalTele size=30></TD></TR>             
        <TR>             
          <TD align=center bgColor=#99ccff height=25  width="134">其它要求:</TD>             
          <TD bgColor=#99ccff height=25 width="451">&nbsp;     
          <INPUT name=OtherNeed size=50></TD></TR>             
        <TR>             
          <TD align=middle bgColor=#99ccff colSpan=2 height="27" width="591">    
          <INPUT name=yes type=submit value="<确定>">  &nbsp;&nbsp;&nbsp;     
          <INPUT name=cancel onclick=history.back() type=button value="<重选酒店>">              
          </TD>  
          </TR></TBODY></TABLE></CENTER></DIV>             
      </TD></TR></TBODY>  
      </TABLE>  
      </FORM>  
      </BODY>  
      </HTML>             
