<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!-- saved from url=(0036)http://www.cui-net.com/Hotels_rs.asp -->
<HTML><HEAD><TITLE>�Ƶ�Ԥ��</TITLE>
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
      ��</TD></TR>
  <TR>
    <TD align=middle bgColor=#00FF00 height=49 noWrap width="100%">
      <p align="center"><font size="6" color="#800080" face="����">�Ƶ��ѯ</font></TD></TR>
  <TR>
    <TD align=middle bgColor=#ffb38e height=23 noWrap width="100%">
      <P align=center><FONT size=3>&nbsp;��ѡ������Ҫ��Ƶ�</FONT></P></TD></TR>
  <TR>
    <TD bgColor=#99ccff height=59 width="100%">
    <FONT size=2><FONT color=#976253>&nbsp;�Ƶ����ڳ���</FONT> <FONT color=#000080>&nbsp;&nbsp;                    
                       
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
         '<OPTION   selected>����
         '<OPTION >����
         '<OPTION >��ͷ
         '<OPTION >������
         '<OPTION >����
         '<OPTION >����
         '<OPTION >��ɳ
         '<OPTION >����
         '<OPTION >�ɶ�
         '<OPTION >�е�
         '<OPTION >����
         '<OPTION >��ݸ
         '<OPTION >��ɽ
         '<OPTION >����
         '<OPTION >����
         '<OPTION >����
         '<OPTION >����
         '<OPTION >������
         '<OPTION >����
         '<OPTION >����
         '<OPTION >�Ϸ�
         '<OPTION >���Ϻ���
         '<OPTION  >��ɽ
         '<OPTION  >����
         '<OPTION  >����
         '<OPTION  >����
         '<OPTION  >����
         '<OPTION  >����
         '<OPTION  >����
         '<OPTION  >����
         '<OPTION  >����
         '<OPTION  >÷��
         '<OPTION  >�ϲ�
         '<OPTION  >�Ͼ�
         '<OPTION  >����
         '<OPTION  >��ͨ
         '<OPTION  >����
         '<OPTION  >�ൺ
         '<OPTION  >Ȫ��
         '<OPTION  >����
         '<OPTION  >��ͷ
         '<OPTION  >�Ϻ�
         '<OPTION  >����
         '<OPTION  >����
         '<OPTION  >ʯ��ׯ
         '<OPTION  >˳��
         '<OPTION  >����
         '<OPTION  >���
         '<OPTION  >ͩ®��
         '<OPTION  >����
         '<OPTION  >����
         '<OPTION  >��³ľ��
         '<OPTION  >����
         '<OPTION  >���
         '<OPTION  >�人
         '<OPTION  >����ɽ
         '<OPTION  >����
         '<OPTION  >����
         '<OPTION  >���
         '<OPTION  >�˲�
         '<OPTION  >����
         '<OPTION  >����
         '<OPTION  >��
         '<OPTION  >֣��
         '<OPTION  >��ɽ
         '<OPTION  >����
         '<OPTION  >�麣   
         %>     
      </select></p>       
         </FONT>
         </FONT>
         </TD>
         </TR>
  <TR>
     <TD bgColor=#99ccff height=17 width="100%">
     <FONT class=medium  color=#003399 size=2>&nbsp;�Ƶ��Ǽ�</FONT> 
     </TD>
  </TR>
  <TR>
    <TD align=left bgColor=#99ccff width="15%" height="18">
    <FONT  size=2>&nbsp;&nbsp;               
    <INPUT name=R1 checked type=radio value=V2 >���Ǽ�����                     
    <INPUT name=R1 type=radio value=V3>���Ǽ�����               
    <INPUT name=R1 type=radio value=V4>���Ǽ�����               
    <INPUT name=R1 type=radio value=V5>���Ǽ�               
    </FONT>              
    </TD>              
  </TR>                    
  <TR>                    
    <TD bgColor=#99ccff height=18 width="100%"></TD></TR>                    
  <TR>                    
    <TD bgColor=#99ccff height=17 width="100%">              
    <FONT class=medium  color=#003399 size=2>&nbsp;�۸�Χ</FONT>               
    </TD>              
  </TR>                    
  <TR>                    
    <TD align=left bgColor=#99ccff height="18">              
    <FONT size=2>&nbsp;&nbsp;               
    <INPUT  CHECKED name=R2 type=radio value=H0>����               
    <INPUT name=R2 type=radio  value=H1>100-400               
    <INPUT name=R2 type=radio  value=H2>400-700               
    <INPUT name=R2 type=radio  value=H3>700-1100               
    <INPUT name=R2 type=radio  value=H4>1100����              
    </FONT>               
    </TD>                    
  <TR>                    
    <TD bgColor=#99ccff height=18 width="100%">��</TD>                    
  <TR>                    
    <TD bgColor=#99ccff height=18 width="100%">��</TD>                    
  <TR>                    
    <TD align=middle bgColor=#99ccff height="27">                    
      <P align=center><FONT size=2>              
         <INPUT name=ChaXun style="COLOR: rgb(0,0,128); FONT-FAMILY: ����" type=submit value="<�� ѯ>">��              
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
