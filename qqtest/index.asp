<% 
const senlontitle="QQ�������"
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>QQ�������_ʵ�ò�ѯ���ߴ�ȫ - Powered by Senlon!</TITLE>

</head>
<body><SCRIPT language=javascript>
<!--
function Check(theForm)
{
var qq1 = theForm.qq1.value;
if (qq1 == "") 
{
alert("������QQ��");
theForm.qq1.value="";
theForm.qq1.focus();
return false;
}
var qq2 = theForm.qq2.value;
if (qq2 == "") 
{
alert("������Է�QQ��");
theForm.qq2.value="";
theForm.qq2.focus();
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
  <td width="100%" class="ttd"><span class="red">QQ�������:</span>
  <br><br>������������ҿ϶����и�QQ�Űɣ���ô����֪��QQ���Ѹ��Լ��Ĺ�ϵ������������������Ե�˫��QQ�ţ����Ͻ������ȤζС���԰ɣ� <br> <br></td>
</tr>
<form name="form1" onSubmit="return Check(this)" method="post" action="">
<input type="hidden" name="act" value="ok" />
  <tr>
    <td colspan="2" class="new">
�ҵ�QQ��
  <input name="qq1" type="text" id="qq1" size="20"  onKeyUp="value=value.replace(/[^\d]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))" />
  �Է�QQ:
  <input name="qq2" type="text" id="qq2" size="20"  onKeyUp="value=value.replace(/[^\d]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))" />
  <input type="submit" name="Submit1" value="��ʼ����" style="cursor:hand;"></form></td>
    </tr><%
	if request("act")="ok" then
strname1=request.form("qq1")
qqsz1=0
for i=1 to Len(strname1)
qqsz1=qqsz1+Mid(strname1,i,1)
next
strname2=request.form("qq2")
qqsz2=0
for i=1 to Len(strname2)
qqsz2=qqsz2+Mid(strname2,i,1)
next

qqsz=qqsz1+qqsz2
if qqsz>100 then
qqsz=(qqsz mod 100)
end if
%>
<%
'����
sql="select * from qlpdbh where bihua like '%"&qqsz&"%'"
set rs=conn.execute(sql)
if not (rs.bof and rs.eof) then
intro=rs("intro")
end if
%>
<tr bgcolor="#EFF8FE">
<td class="new" colspan="2" valign="middle">
<%="<br>˫��QQ��<font color=blue>"&strname1&"&nbsp;"&strname2&" </font><br><br>"%>
<%="<font color=cc6600>��ϵ���ܣ�</font>"&intro&"<font color=#CCCCCC>(��������)</font><br><br>"%>
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
