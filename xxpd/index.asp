<% 
const senlontitle="血型配对"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>血型配对_实用查询工具大全 - Powered by Senlon!</TITLE>

</head>
<body><!--#include file="../senlon/function.asp"--><!--#include file="../senlon/conn.asp"-->  
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
<tbody><tr>
  <td width="100%" class="ttd"><span class="red">血型配对:</span>
  <br><br>根据双方的血型，可以测试出您和恋人的缘分和婚姻哦，快来试试吧！<br><br></td>
</tr>
<form name="form1"  method="post" action="">
<input type="hidden" name="act" value="xxok" />
  <tr>
    <td colspan="2" class="new">
我的血型：
  <SELECT name="p3" class="style3">
<OPTION value=AB<%=str1%>>AB型</OPTION>
<OPTION value=A<%=str2%>>A型</OPTION>
<OPTION value=B<%=str3%>>B型</OPTION>
<OPTION value=O<%=str4%>>O型</OPTION>
</SELECT>
  他/她的血型
  :
  <SELECT name="p4" class="style3">
<OPTION value=AB>AB型</OPTION>
<OPTION value=A>A型</OPTION>
<OPTION value=B>B型</OPTION>
<OPTION value=O>O型</OPTION>
</SELECT>
  <input type="submit" name="Submit1" value="开始配对" style="cursor:hand;">
    </form></td>
    </tr><%
	if request("act")="xxok" then
sqlstr="select * from xuexinglove where xuexing1='"&trim(cstr(request("p3")))&"' and xuexing2='"&trim(cstr(request("p4")))&"'"
set rs=conn.execute(sqlstr)

%>
<tr bgcolor="#EFF8FE">
<td class="new" colspan="2" valign="middle">
<%="<br>血型配对：<font color=blue>"&rs("title1")&"</font><br><br>"%>
<%="<font color=red>"&rs("title2")&"</font><br><br>"%>
<%=""&rep(rs("content"))&"<br>"%>

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
