<% 
const senlontitle="生肖配对"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>生肖配对_实用查询工具大全 - Powered by Senlon!</TITLE>

</head>
<body><!--#include file="../senlon/function.asp"--><!--#include file="../senlon/conn.asp"-->  
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
  <table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
<tbody><tr>
  <td width="100%" class="ttd"><span class="red">生肖配对:</span>
  <br><br>属相有鼠、牛、虎、兔、龙、蛇、马、羊、猴、鸡、狗、猪十二种。属相又都是由五行来区分的，而五行则存在着相生相克的道理。相生的属相较能和睦相处，而相克的属相就不那么……了。您是否还对您的爱情犹豫不决啊，那就进来看看你们的生肖配不配吧！<br><br></td>
</tr>
<form name="form1"  method="post" action="">
<input type="hidden" name="act" value="sxok" />
  <tr>
    <td colspan="2" class="new">
我的生肖：
  <SELECT name="p5" class="style3">
<OPTION value=鼠<%=str1%>>鼠</OPTION>
<OPTION value=牛<%=str2%>>牛</OPTION>
<OPTION value=虎<%=str3%>>虎</OPTION>
<OPTION value=兔<%=str4%>>兔</OPTION>
<OPTION value=龙<%=str5%>>龙</OPTION>
<OPTION value=蛇<%=str6%>>蛇</OPTION>
<OPTION value=马<%=str7%>>马</OPTION>
<OPTION value=羊<%=str8%>>羊</OPTION>
<OPTION value=猴<%=str9%>>猴</OPTION>
<OPTION value=鸡<%=str10%>>鸡</OPTION>
<OPTION value=狗<%=str11%>>狗</OPTION>
<OPTION value=猪<%=str12%>>猪</OPTION>
</SELECT>
  他/她的生肖
  :
  <SELECT name="p6" class="style3">
<OPTION value=鼠 selected>鼠</OPTION>
<OPTION value=牛>牛</OPTION>
<OPTION value=虎>虎</OPTION>
<OPTION value=兔>兔</OPTION>
<OPTION value=龙>龙</OPTION>
<OPTION value=蛇>蛇</OPTION>
<OPTION value=马>马</OPTION>
<OPTION value=羊>羊</OPTION>
<OPTION value=猴>猴</OPTION>
<OPTION value=鸡>鸡</OPTION>
<OPTION value=狗>狗</OPTION>
<OPTION value=猪>猪</OPTION>
</SELECT>
  <input type="submit" name="Submit1" value="开始配对" style="cursor:hand;">
    </form></td>
    </tr><%
	if request("act")="sxok" then
dim sqlstr
sqlstr="select * from shuxianglove where shengxiao1='"&trim(cstr(request("p5")))&"' and shengxiao2='"&trim(cstr(request("p6")))&"'"
set rs=conn.execute(sqlstr)

%>
<tr bgcolor="#EFF8FE">
<td class="new" colspan="2" valign="middle">
<%="<br>生肖配对：<font color=blue>"&rs("title")&"</font><br><br>"%>
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
</center>
</body></html>
