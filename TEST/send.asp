<html>
<head>
<title>留言</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script language="JavaScript">
<!--
function MM_findObj(n, d) { //v3.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document); return x;
}

function MM_validateForm() { //v3.0
  var i,p,q,nm,test,num,min,max,errors='',args=MM_validateForm.arguments;
  for (i=0; i<(args.length-2); i+=3) { test=args[i+2]; val=MM_findObj(args[i]);
    if (val) { nm=val.name; if ((val=val.value)!="") {
      if (test.indexOf('isEmail')!=-1) { p=val.indexOf('@');
        if (p<1 || p==(val.length-1)) errors+='- '+nm+' must contain an e-mail address.\n';
      } else if (test!='R') { num = parseFloat(val);
        if (val!=''+num) errors+='- '+nm+' must contain a number.\n';
        if (test.indexOf('inRange') != -1) { p=test.indexOf(':');
          min=test.substring(8,p); max=test.substring(p+1);
          if (num<min || max<num) errors+='- '+nm+' must contain a number between '+min+' and '+max+'.\n';
    } } } else if (test.charAt(0) == 'R') {
    						if (nm=="name") nm="留下姓名";
    						if (nm=="add") nm="来自何方";
    						if (nm=="memo") nm="一吐为快";
    						errors += '- '+nm+' 。\n'; 
    						}}
  } if (errors) alert('您必须键入:?\n'+errors);
  document.MM_returnValue = (errors == '');
}
//-->
</script>
<style>
<!---

td,p,input,select,dl,li {font-family:宋体;font-size:12px;}
.input {   font-size: 8pt; color: #660066; background-color: #FFFDD7; border-color:0000FF; border: 0000FF; border-style: dotted; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px;}
.input1 {   font-size: 8pt; color: #660066;  border-color:FFCC00; border:FFCC00; border-style: dotted; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px; cursor: hand}
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

a:link {text-decoration: underline; color: #00ff00}
a:visited {text-decoration: underline; color: #008000}
a:active {text-decoration: none}
a:hover {text-decoration: underline; color: #ff0000}

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

<body bgcolor="#FF0000" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">
<!-- HEADER GRAPHIC--> <!--HEADER GRAPHIC_END--> 
<form method="post" action="gform.asp" name="form1">
         
<table border=0 cellspacing=10 cellpadding=0 align="center" bgcolor="#EFFDE6" width=460>
<tr>
<td colspan="3">
<table width=450 border="0" cellspacing="0" cellpadding="0" align="center">
          <tr bgcolor="#66CCCC"> 
            <td><font face="黑体"><b>　</b></font></td>
                  </tr>
        </table>
        <table border=0 cellspacing=0 cellpadding=0 width=450 align="center">
          <tr> 
            <td class=bgc1></td>
          </tr>
        </table>
        <table border=0 cellspacing=0 cellpadding=0 width=450 align="center">
          <tr> 
            <td class=bgc1></td>
          </tr>
        </table> 
</td>
</tr>      
 <tr bordercolor="#CCCC99" bgcolor="#CCFFCC"> 
      <td nowrap> 
        <div align="center">留下姓名：</div>
      </td>
      <td> 
        <input type="text" name="name" size="30" class="input">
      </td>
    </tr>
    <tr bordercolor="#CCCC99" bgcolor="#CCFFCC"> 
      <td  nowrap> 
        <div align="center">来自何方：</div>
      </td>
      <td> 
        <input type="text" name="add" size="30" class="input">
      </td>
    </tr>
    
    <tr bordercolor="#CCCC99" bgcolor="#CCFFCC"> 
      <td valign="top"  nowrap> 
        <div align="center">一吐为快：</div>
      </td>
      <td valign="top"> 
        <textarea rows="12" name="memo" cols="50" class="input"></textarea>
      </td>
    </tr>
    <tr bordercolor="#CCCC99" bgcolor="#CCFFCC"> 
      <td colspan="2"> 
        <div align="center"> 
          <input type="submit" name="submit" value="发布"  style="BACKGROUND-COLOR: rgb(0,128,192); COLOR: rgb(255,255,255); FONT-SIZE: 9pt"  onClick="MM_validateForm('name','','R','add','','R','memo','','R');return document.MM_returnValue" class="input1">
          <input type="reset" name="reset"  style="BACKGROUND-COLOR: rgb(0,128,192); COLOR: rgb(255,255,255); FONT-SIZE: 9pt" value="重写" class="input1">
           
          <input type="button" name="reset3"   style="BACKGROUND-COLOR: rgb(0,128,192); COLOR: rgb(255,255,255); FONT-SIZE: 9pt"  value="返回" onclick=location.href="mousheng1.asp" class="input1">
          <input type="button" name="Submit2"  style="BACKGROUND-COLOR: rgb(0,128,192); COLOR: rgb(255,255,255); FONT-SIZE: 9pt" value="管理区" class="input1">
        </div>
      </td>
    </tr>
      <tr valign=top align="center"> 
    <td colspan="3"> 
      <hr>
        <span class="f12cb">Power by 俺; Build time：2000.8; Copyleft, All rights abandoned.</span> 
      </td>
  </tr>
  </table>
</form>

</body>
</html>
