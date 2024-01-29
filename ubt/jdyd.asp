<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!-- saved from url=(0036)http://www.cui-net.com/Hotels_rs.asp -->
<HTML><HEAD><TITLE>酒店预订</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<META content="Microsoft FrontPage 4.0" name=GENERATOR>

</HEAD>
<BODY leftMargin=5 topMargin=0>
<DIV align=center>  
  <INPUT name=hname type=hidden value=hvalue>  
<TABLE border=0 width="80%" height="351"> 
  <TBODY> 
  <TR>
    <TD align=middle bgColor=#FFFFFF height=49 noWrap width="100%">
      　</TD></TR>
  <TR>
    <TD align=middle bgColor=#00FF00 height=49 noWrap width="100%">
      <p align="center"><font size="6" color="#800080" face="隶书">酒店查询</font></TD></TR>
  <TR>
    <TD align=middle bgColor=#ffb38e height=23 noWrap width="100%">
      <P align=center><FONT size=3>&nbsp;请选择您的要求酒店</FONT></P></TD></TR>
  <TR>
    <TD bgColor=#99ccff height=59 width="100%">
    <FONT size=2><FONT color=#976253>&nbsp;酒店所在城市</FONT> <FONT color=#000080>&nbsp;&nbsp;                    
                       
    <form action="search1.asp" method="POST">                 
       <p>                
       <select name="City" size="1"> 
       <% SQL="Citys"
          Set conn = Server.CreateObject("ADODB.Connection")
          conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("hotel.mdb")
          Set Rs=conn.Execute(SQL)
          j=0
          While Not rs.EOF 
              j=j+1
              If j=1 Then
                  Response.Write "<OPTION selected>" & Rs.Fields(1).Value 
              Else 
                  Response.Write "<OPTION>" & Rs.Fields(1).Value  
              End If
              Rs.MoveNext   
         Wend
       %>
       <%                    
         '<OPTION   selected>北京
         '<OPTION >澳门
         '<OPTION >包头
         '<OPTION >北戴河
         '<OPTION >北海
         '<OPTION >长春
         '<OPTION >长沙
         '<OPTION >常州
         '<OPTION >成都
         '<OPTION >承德
         '<OPTION >大连
         '<OPTION >东莞
         '<OPTION >佛山
         '<OPTION >福州
         '<OPTION >广州
         '<OPTION >贵阳
         '<OPTION >桂林
         '<OPTION >哈尔滨
         '<OPTION >海口
         '<OPTION >杭州
         '<OPTION >合肥
         '<OPTION >呼合浩特
         '<OPTION  >黄山
         '<OPTION  >惠州
         '<OPTION  >济南
         '<OPTION  >江阴
         '<OPTION  >景洪
         '<OPTION  >昆明
         '<OPTION  >拉萨
         '<OPTION  >丽江
         '<OPTION  >洛阳
         '<OPTION  >梅州
         '<OPTION  >南昌
         '<OPTION  >南京
         '<OPTION  >南宁
         '<OPTION  >南通
         '<OPTION  >宁波
         '<OPTION  >青岛
         '<OPTION  >泉州
         '<OPTION  >三亚
         '<OPTION  >汕头
         '<OPTION  >上海
         '<OPTION  >深圳
         '<OPTION  >沈阳
         '<OPTION  >石家庄
         '<OPTION  >顺德
         '<OPTION  >苏州
         '<OPTION  >天津
         '<OPTION  >桐庐镇
         '<OPTION  >威海
         '<OPTION  >温州
         '<OPTION  >乌鲁木齐
         '<OPTION  >无锡
         '<OPTION  >武昌
         '<OPTION  >武汉
         '<OPTION  >武夷山
         '<OPTION  >西安
         '<OPTION  >厦门
         '<OPTION  >香港
         '<OPTION  >宜昌
         '<OPTION  >银川
         '<OPTION  >肇庆
         '<OPTION  >镇江
         '<OPTION  >郑州
         '<OPTION  >中山
         '<OPTION  >重庆
         '<OPTION  >珠海   
         %>     
      </select></p>       
         </FONT>
         </FONT>
         </TD>
         </TR>
  <TR>
     <TD bgColor=#99ccff height=17 width="100%">
     <FONT class=medium  color=#003399 size=2>&nbsp;酒店星级</FONT> 
     </TD>
  </TR>
  <TR>
    <TD align=left bgColor=#99ccff width="15%" height="18">
    <FONT  size=2>&nbsp;&nbsp;               
    <INPUT name=R1 checked type=radio value=V2 >二星级以上                     
    <INPUT name=R1 type=radio value=V3>三星级以上               
    <INPUT name=R1 type=radio value=V4>四星级以上               
    <INPUT name=R1 type=radio value=V5>五星级               
    </FONT>              
    </TD>              
  </TR>                    
  <TR>                    
    <TD bgColor=#99ccff height=18 width="100%"></TD></TR>                    
  <TR>                    
    <TD bgColor=#99ccff height=17 width="100%">              
    <FONT class=medium  color=#003399 size=2>&nbsp;价格范围</FONT>               
    </TD>              
  </TR>                    
  <TR>                    
    <TD align=left bgColor=#99ccff height="18">              
    <FONT size=2>&nbsp;&nbsp;               
    <INPUT  CHECKED name=R2 type=radio value=H0>不限               
    <INPUT name=R2 type=radio  value=H1>100-400               
    <INPUT name=R2 type=radio  value=H2>400-700               
    <INPUT name=R2 type=radio  value=H3>700-1100               
    <INPUT name=R2 type=radio  value=H4>1100以上              
    </FONT>               
    </TD>                    
  <TR>                    
    <TD bgColor=#99ccff height=18 width="100%">　</TD>                    
  <TR>                    
    <TD bgColor=#99ccff height=18 width="100%">　</TD>                    
  <TR>                    
    <TD align=middle bgColor=#99ccff height="27">                    
      <P align=center><FONT size=2>              
         <INPUT name=ChaXun style="COLOR: rgb(0,0,128); FONT-FAMILY: 黑体" type=submit value="<查 询>">　              
      </FONT></P>            
     </TD>            
  </TR>               
  </FORM>                 
  <TR>                    
    <TD height=21 width="100%">                    
      <HR align=center color=#008000 SIZE=1>                    
    </TD>            
  </TR>                    
  </TBODY>            
  </TABLE></DIV>                    
  </FORM>                   
</BODY></HTML>                   
