<% 
const senlontitle="号码吉凶测试"
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>号码吉凶测试,阿拉伯数字吉凶分析_实用查询工具大全 - Powered by Senlon!</TITLE>
<!--#include file="../senlon/getuc.asp"-->

</head>
<body><script language="JavaScript">
<!--
function checkjx()
{
var word=document.theform.word.value;
if (word=='' || word.length>20)
{
alert('请输入20位以内的数字哦！');
document.theform.word.focus()
return false;
}
}
//-->
</script>
<!--#include file="../senlon/function.asp"-->
<!--#include file="../senlon/conn.asp"-->  
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
  <table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
<tbody><tr>
  <td width="100%" class="ttd"><span class="red">号码吉凶测试:</span>
    <br><br>本数理分析系统可以分析出20位以内的阿拉伯数字的数理吉凶。有些人会问，数字为什么可能会影响一个人的运势？其实这就像姓名、风水会影响运势命运的意义是一样的。虽然这只是一个号码，但是它与您的生活息息相关，也是您与很多人沟通的桥梁！所以『吉』与『凶』关系非常大，的确不可轻忽！ <br><br> </td>
</tr>
<form name="theform" onSubmit="return checkjx();" method="post" action="">
<input type="hidden" name="act" value="ok" />
  <tr>
    <td colspan="2" class="ttd">
请输入要分析的数字：<input type="text" name="word" size="20" maxlength="20" onKeyUp="value=value.replace(/[^\d]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))" />
<input type="submit" name="Submit1" value="立刻分析" style="cursor:hand;"></form></td>
    </tr>
<%
'手机吉凶
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
<tr bgcolor="#EFF8FE">
<td class="ttd" colspan="2" valign="middle">您分析的号码：<%=request("k")%> <span class="red"><%=request("word")%></span></td>
</tr><tr bgcolor="#EFF8FE">
<td class="ttd" colspan="2" valign="middle">号码吉凶分析：<font color=blue><strong><%=title%></strong></font> <span class="red"><%=jx%></span>
</td>
</tr><tr bgcolor="#EFF8FE">
<td class="new" colspan="2" valign="middle">主人个性分析：<%=content%>
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