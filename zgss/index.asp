<% 
const senlontitle="诸葛测字"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>诸葛神算,诸葛神数,诸葛测字_实用查询工具大全 - Powered by Senlon!</TITLE>

</head>
<body><SCRIPT language=javascript>
<!--
function Check(theForm)
{
var name1 = theForm.name1.value;
if (name1 == "") 
{
alert("请输入三个简体汉字进行测算！");
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
}
function Check2(theForm)
{
var name2 = theForm.name2.value;
if (name2 == "") 
{
alert("请输入三个繁体汉字进行测算！");
theForm.name2.value="";
theForm.name2.focus();
return false;
}
if (name2.search(/[`1234567890-=\~!@#$%^&*()_+|<>;':",.?/ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz]/) != -1) 
{
alert("请务必输入繁体汉字！");
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
  <td width="100%" class="ttd"><span class="red">诸葛神数:</span>
  <br><br>诸葛神数相传是三国时代刘备的军师诸葛亮所作。根据历史记载，诸葛亮上懂天文，下晓地理，料事如神，用兵用人，皆恰到好处。诸葛亮每遇难题，必暗自用一种独到的占卜法。心要诚，手要净，焚香向天祷告，然后，在纸上写三个字。这三个字，即是天灵与人心灵交流，也就是说，你的心事已得上天了解，而上天会对你作出指示。所以万万不可存“玩一玩”的心态。诸葛神数，共三百八十四爻，谶语句法，长短不一，寓意深远，对占卜者的思路，有很大的启发作用，特别是那些正陷于彷徨迷惘中的人，更有一种拨开云雾，重见天日的豁然开朗感觉。因此这是可以作为判断吉凶，决定进退，是选择趋吉避凶的指南针。<br><br></td>
</tr>
<form name="form1" onSubmit="return Check(this)" method="post" action="">
<input type="hidden" name="act" value="ok" />
  <tr>
    <td colspan="2" class="new">
请输入三个简体汉字抽签占卜：<input name="name1" type="text" id="name1" size="20" onKeyUp="value=value.replace(/[^\u4E00-\u9FA5]/g,'')" maxLength="3">（多输入只取前三个） <input type="submit" name="Submit1" value="提交" style="cursor:hand;"></form></td>
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
'计算诸葛神数
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
<%="<br>您输入的三个汉字为：<font color=blue>"&str1&"&nbsp;"&str2&"&nbsp;"&str3&"</font><br><br>"%>
<%="<font color=cc6600>〖诗曰〗：</font>"&zhugetitle&"<br><br>"%>
<%="<font color=cc6600>〖签释〗：</font><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"&zhugecontent&"<br><br>"%>
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
