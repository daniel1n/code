<% 
const senlontitle="�ֻ����뼪�ײ���"
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>�ֻ����뼪�ײ���_ʵ�ò�ѯ���ߴ�ȫ - Powered by Senlon!</TITLE>
<!--#include file="../senlon/getuc.asp"-->

</head>
<body><script language="JavaScript">
<!--
function checkjx()
{
var word=document.theform.word.value;
if (word=='' || word.length>11)
{
alert('��������ȷ��11λ�ֻ���Ŷ��');
document.theform.word.focus()
return false;
}
}
//-->
</script><!--#include file="../senlon/function.asp"--><!--#include file="../senlon/conn.asp"-->  
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
  <table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
<tbody><tr>
  <td width="100%" class="ttd"><span class="red">�ֻ����뼪��:</span>
    <br> <br>    
�ֻ��Ѿ���Ϊ����������ϢϢ��ص���Ʒ������ÿ���˶���������һ���ֻ������м�����ֻ��������������һ���ĺ��ˡ��㲻������һ����ĺ��뼪��̶Ȱɣ�����ֻ�ܵ�����Ŷ��</td>
</tr>
<form name="theform" onSubmit="return checkjx();" method="post" action="">
<input type="hidden" name="act" value="ok" />
  <tr>
    <td colspan="2" class="ttd">
������Ҫ�����ĺ��룺<input type="text" name="word" size="20" maxlength="20" onKeyUp="value=value.replace(/[^\d]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))" />
<input type="submit" name="Submit1" value="���̷���" style="cursor:hand;"></form></td>
    </tr>
<%
'�ֻ�����
'on error resume next
dim word,temp
word=cstr(trim(request("word")))
word=mid(word,1)
word=int(word)
word=word/80
temp=int(word)
word=word-temp
word=int(word*80)
word=cstr(word)
if word="0" then
word="81"
end if
sql="select * from shouji where num='"&word&"'"
set rs=conn.execute(sql)
if not (rs.bof and rs.eof) then
num=rs("num")
title=rs("title")
jx=rs("jx")
jx="("&jx&")"
content=rs("content")
end if
%>
<tr>
<td class="ttd" colspan="2" valign="middle">���������ֻ��ţ�<%=request("k")%> <span class="red"><%=request("word")%></span></td>
</tr><tr>
<td class="ttd" colspan="2" valign="middle">�ֻ��ż��׷�����<font color=blue><strong><%=title%></strong></font> <span class="red"><%=jx%></span>
</td>
</tr><tr>
<td class="new" colspan="2" valign="middle">���˵ĸ��Է�����<%=content%>
</td>
</tr>
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