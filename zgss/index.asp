<% 
const senlontitle="������"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>�������,�������,������_ʵ�ò�ѯ���ߴ�ȫ - Powered by Senlon!</TITLE>

</head>
<body><SCRIPT language=javascript>
<!--
function Check(theForm)
{
var name1 = theForm.name1.value;
if (name1 == "") 
{
alert("�������������庺�ֽ��в��㣡");
theForm.name1.value="";
theForm.name1.focus();
return false;
}
if (name1.search(/[`1234567890-=\~!@#$%^&*()_+|<>;':",.?/ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz]/) != -1) 
{
alert("�����������庺�֣�");
theForm.name1.value = "";
theForm.name1.focus();
return false;
}
}
function Check2(theForm)
{
var name2 = theForm.name2.value;
if (name2 == "") 
{
alert("�������������庺�ֽ��в��㣡");
theForm.name2.value="";
theForm.name2.focus();
return false;
}
if (name2.search(/[`1234567890-=\~!@#$%^&*()_+|<>;':",.?/ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz]/) != -1) 
{
alert("��������뷱�庺�֣�");
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
  <td width="100%" class="ttd"><span class="red">�������:</span>
  <br><br>��������ഫ������ʱ�������ľ�ʦ�����������������ʷ���أ�������϶����ģ������������������ñ����ˣ���ǡ���ô��������ÿ�����⣬�ذ�����һ�ֶ�����ռ��������Ҫ�ϣ���Ҫ�����������쵻�棬Ȼ����ֽ��д�����֡��������֣����������������齻����Ҳ����˵����������ѵ������˽⣬��������������ָʾ���������򲻿ɴ桰��һ�桱����̬����������������ٰ�ʮ��س������䷨�����̲�һ��Ԣ����Զ����ռ���ߵ�˼·���кܴ���������ã��ر�����Щ��������������е��ˣ�����һ�ֲ��������ؼ����յĻ�Ȼ���ʸо���������ǿ�����Ϊ�жϼ��ף��������ˣ���ѡ���������׵�ָ���롣<br><br></td>
</tr>
<form name="form1" onSubmit="return Check(this)" method="post" action="">
<input type="hidden" name="act" value="ok" />
  <tr>
    <td colspan="2" class="new">
�������������庺�ֳ�ǩռ����<input name="name1" type="text" id="name1" size="20" onKeyUp="value=value.replace(/[^\u4E00-\u9FA5]/g,'')" maxLength="3">��������ֻȡǰ������ <input type="submit" name="Submit1" value="�ύ" style="cursor:hand;"></form></td>
    </tr><%strname=request.form("name1")
str1=mid(strname,1,1)
str2=mid(strname,2,1)
str3=mid(strname,3,1)
if str1<>"" and str2<>"" and str3<>"" then
bihua1=(getnum(str1) mod 10)
if bihua1=0 then
bihua1=1
end if
bihua2=(getnum(str2) mod 10)
if bihua2=0 then
bihua2=1
end if
bihua3=(getnum(str3) mod 10)
if bihua3=0 then
bihua3=1
end if
bihua1=cstr(bihua1)
bihua2=cstr(bihua2)
bihua3=cstr(bihua3)
bihua=bihua1&bihua2&bihua3
bihua=int(bihua)
do while bihua>=384
bihua=bihua-384
loop
if bihua<=9 then
bihua=cstr(bihua)
bihua="0"&"0"&bihua
else
if bihua<=99 then
bihua=cstr(bihua)
bihua="0"&bihua
end if
end if
bihua=cstr(bihua)
%>
<%
'�����������
on error resume next
dim sqlzhuge,zhugetitle,zhugecontent
sqlzhuge="select * from zhuge where id='"&bihua&"'"
set rszhuge=conn.execute(sqlzhuge)
if not (rszhuge.bof and rszhuge.eof) then
zhugetitle=rszhuge("title")
zhugecontent=rszhuge("content")
end if
%>
<tr bgcolor="#EFF8FE">
<td class="new" colspan="2" valign="middle">
<%="<br>���������������Ϊ��<font color=blue>"&str1&"&nbsp;"&str2&"&nbsp;"&str3&"</font><br><br>"%>
<%="<font color=cc6600>��ʫԻ����</font>"&zhugetitle&"<br><br>"%>
<%="<font color=cc6600>��ǩ�͡���</font><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&zhugecontent&"<br><br>"%>
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
