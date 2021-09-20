<% 
const senlontitle="姓名配对"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>姓名配对_实用查询工具大全 - Powered by Senlon!</TITLE>

</head>
<body><SCRIPT language=javascript>
<!--
function Check(theForm)
{
var name1 = theForm.name1.value;
if (name1 == "") 
{
alert("请输入您的姓名！");
theForm.name1.value="";
theForm.name1.focus();
return false;
}
if (name1.search(/[`1234567890-=\~!@#$%^&*()_+|<>;':",.?/ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz]/) != -1) 
{
alert("请务必输入简体汉字！");
theForm.name1.value = "";
theForm.name1.focus();
return false;
}
var name2 = theForm.name2.value;
if (name2 == "") 
{
alert("请输入您爱人的名字！");
theForm.name2.value="";
theForm.name2.focus();
return false;
}
if (name2.search(/[`1234567890-=\~!@#$%^&*()_+|<>;':",.?/ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz]/) != -1) 
{
alert("请务必输入简体汉字！");
theForm.name2.value = "";
theForm.name2.focus();
return false;
}

}

//-->
</SCRIPT><!--#include file="../senlon/function.asp"--><!--#include file="../senlon/conn.asp"-->  
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
  <table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
<tbody><tr>
  <td width="100%" class="ttd"><span class="red">姓名配对:</span>
  <br><br>姓名当中究竟隐藏了多少奥秘，可能至今也没有人能完全说破，这里有个趣味游戏，通过姓名笔划数看看你和爱人的关系究竟怎样。<br><br></td>
</tr>
<form name="form1" onSubmit="return Check(this)" method="post" action="">
<input type="hidden" name="act" value="ok" />
  <tr>
    <td colspan="2" class="new">
请输入您的姓名：
  <input name="name1" type="text" id="name1" size="20" maxLength="4"> 
  请输入您爱人的名字:
  <input name="name2" type="text" id="name2" size="20" maxLength="4"> 
  <input type="submit" name="Submit1" value="开始测试" style="cursor:hand;"></form></td>
    </tr><%
	if request("act")="ok" then
strname1=request.form("name1")
str1=mid(strname1,1,1)
str2=mid(strname1,2,1)
str3=mid(strname1,3,1)
str4=mid(strname1,4,1)
bihua1=getnum(str1)
bihua2=getnum(str2)
bihua3=getnum(str3)
bihua4=getnum(str4)
bihuaname1=bihua1+bihua2+bihua3+bihua4


strname2=request.form("name2")
sstr1=mid(strname2,1,1)
sstr2=mid(strname2,2,1)
sstr3=mid(strname2,3,1)
sstr4=mid(strname2,4,1)
sbihua1=getnum(sstr1)
sbihua2=getnum(sstr2)
sbihua3=getnum(sstr3)
sbihua4=getnum(sstr4)
bihuaname2=sbihua1+sbihua2+sbihua3+sbihua4
bihuac=bihuaname1+bihuaname2
if bihuac>100 then
bihuac=(bihuac mod 100)
end if
%>
<%
'计算
sql="select * from qlpdbh where bihua like '%"&bihuac&"%'"
set rs=conn.execute(sql)
if not (rs.bof and rs.eof) then
intro=rs("intro")
end if
%>
<tr bgcolor="#EFF8FE">
<td class="new" colspan="2" valign="middle">
<%="<br>双方姓名：<font color=blue>"&strname1&"&nbsp;"&strname2&" </font><br><br>"%>
<%="<font color=cc6600>关系揭密：</font>"&intro&"<font color=#CCCCCC>(仅供娱乐哦)</font><br><br>"%>
</td>
</tr><%end if%>
</tbody>
</table>
      <!--#INCLUDE FILE="../dq.asp"-->
	</div>
    </td>
    <td valign="top"><!--#INCLUDE FILE="../right.asp"--></td>
  </tr>
</table>
<!--#include file="../foot.asp"-->
</body></html>
