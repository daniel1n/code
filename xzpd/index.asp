<% 
const senlontitle="星座配对"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>星座配对_实用查询工具大全 - Powered by Senlon!</TITLE>

</head>
<body><!--#include file="../senlon/function.asp"--><!--#include file="../senlon/conn.asp"-->  
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
  <table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
<tbody><tr>
  <td width="100%" class="ttd"><span class="red">星座配对:</span>
  <br><br>你相信星座配对吗？最佳的星座配能保证你的爱情浪漫永久吗？您认为星座是科学吗？星座真的有神奇力量吗？不管信或不信，测试一下看看先，仅供娱乐！<br><br></td>
</tr>
<form name="form1"  method="post" action="">
<input type="hidden" name="act" value="ok" />
  <tr>
    <td colspan="2" class="new">
我的星座：
  <SELECT name="p1" class="style3">
  <OPTION value=白羊座<%=str1%>>白羊座</OPTION>
  <OPTION value=金牛座<%=str2%>>金牛座</OPTION>
  <OPTION value=双子座<%=str3%>>双子座</OPTION>
  <OPTION value=巨蟹座<%=str4%>>巨蟹座</OPTION>
  <OPTION value=狮子座<%=str5%>>狮子座</OPTION>
  <OPTION value=处女座<%=str6%>>处女座</OPTION>
  <OPTION value=天秤座<%=str7%>>天秤座</OPTION>
  <OPTION value=天蝎座<%=str8%>>天蝎座</OPTION>
  <OPTION value=射手座<%=str9%>>射手座</OPTION>
  <OPTION value=摩羯座<%=str10%>>摩羯座</OPTION>
  <OPTION value=水瓶座<%=str11%>>水瓶座</OPTION>
  <OPTION value=双鱼座<%=str12%>>双鱼座</OPTION>
  </SELECT>
  他/她的星座
  :
  <SELECT name="p2" class="style3">
  <OPTION value=白羊座 selected>白羊座</OPTION>
  <OPTION value=金牛座>金牛座</OPTION>
  <OPTION value=双子座>双子座</OPTION>
  <OPTION value=巨蟹座>巨蟹座</OPTION>
  <OPTION value=狮子座>狮子座</OPTION>
  <OPTION value=处女座>处女座</OPTION>
  <OPTION value=天秤座>天秤座</OPTION>
  <OPTION value=天蝎座>天蝎座</OPTION>
  <OPTION value=射手座>射手座</OPTION>
  <OPTION value=摩羯座>摩羯座</OPTION>
  <OPTION value=水瓶座>水瓶座</OPTION>
  <OPTION value=双鱼座>双鱼座</OPTION>
  </SELECT>
  <input type="submit" name="Submit1" value="开始配对" style="cursor:hand;">
    </form></td>
    </tr><%
	if request("act")="ok" then
dim sqlstr
sqlstr="select * from xingzuolove where xingzuo1='"&trim(cstr(request("p1")))&"' and xingzuo2='"&trim(cstr(request("p2")))&"'"
set rs=conn.execute(sqlstr)
%>
<tr bgcolor="#EFF8FE">
<td class="new" colspan="2" valign="middle">
<%="<br>双方星座：<font color=blue>"&rs("title")&"</font><br><br>"%>
<%="<font color=red>"&rep(rs("content1"))&"</font><br><br>"%>
<%=""&rep(rs("content2"))&"<br>"%>

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
