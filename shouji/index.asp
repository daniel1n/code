<% 
const senlontitle="手机号码吉凶测试"
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>手机号码吉凶测试_实用查询工具大全 - Powered by Senlon!</TITLE>
<!--#include file="../senlon/getuc.asp"-->

</head>
<body><script language="JavaScript">
<!--
function checkjx()
{
var word=document.theform.word.value;
if (word=='' || word.length>11)
{
alert('请输入正确的11位手机号哦！');
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
  <td width="100%" class="ttd"><span class="red">手机号码吉凶:</span>
    <br> <br>    
手机已经成为人们生活中息息相关的物品，几乎每个人都会至少有一部手机，其中吉祥的手机号码或许给你带来一定的好运。你不妨测试一下你的号码吉祥程度吧，不过只能当娱乐哦。</td>
</tr>
<form name="theform" onSubmit="return checkjx();" method="post" action="">
<input type="hidden" name="act" value="ok" />
  <tr>
    <td colspan="2" class="ttd">
请输入要分析的号码：<input type="text" name="word" size="20" maxlength="20" onKeyUp="value=value.replace(/[^\d]/g,'') " onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))" />
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
<tr>
<td class="ttd" colspan="2" valign="middle">您分析的手机号：<%=request("k")%> <span class="red"><%=request("word")%></span></td>
</tr><tr>
<td class="ttd" colspan="2" valign="middle">手机号吉凶分析：<font color=blue><strong><%=title%></strong></font> <span class="red"><%=jx%></span>
</td>
</tr><tr>
<td class="new" colspan="2" valign="middle">主人的个性分析：<%=content%>
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