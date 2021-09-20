<% 
const senlontitle="心惊预测"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>心惊预测,心跳预测_实用查询工具大全 - Powered by Senlon!</TITLE>
<!--#include file="../senlon/getuc.asp"-->
<link href="../images/css.css" type=text/css rel=stylesheet>
</head>
<body><!--#include file="../senlon/function.asp"--><!--#include file="../senlon/conn.asp"-->  
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
  <table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
<tbody><tr>
  <td width="100%" class="ttd"><span class="red">心惊预测:</span>
    <br><br>感到总是会莫名的心惊？为了它，你总烦躁不安，告诉自己要镇定下来。到底心惊预测着什么，心惊的原因，心惊的征兆是什么，快来看看心惊释义怎么说吧！<br><br></td>
</tr>
<form name="form1" onSubmit="return submitchecken()" method="post" action="">
<input type="hidden" name="act" value="ok" />
  <tr>
    <td colspan="2" class="new">&nbsp;心惊时间：<select id="stime" style="WIDTH: 100px" name="stime" size="1">
<option value="1" selected="">23-01[子时]</option>
<option value="2">01-03[丑时]</option>
<option value="3">03-05[寅时]</option>
<option value="4">05-07[卯时]</option>
<option value="5">07-09[辰时]</option>
<option value="6">09-11[巳时]</option>
<option value="7">11-13[午时]</option>
<option value="8">13-15[未时]</option>
<option value="9">15-17[申时]</option>
<option value="10">17-19[酉时]</option>
<option value="11">19-21[戌时]</option>
<option value="12">21-23[亥时]</option>
</select> <input type="submit" value="开始分析" style="cursor:hand;" /></form></td>
    </tr>
<%
on error resume next
fx=request("fx")
stime=request("stime")
sql="select * from msyuce where stime='"&stime&"' and lb='心惊'"
set rs=conn.execute(sql)
if not (rs.bof and rs.eof) then
%>
<tr bgcolor="#EFF8FE">
<td class="new" colspan="2" valign="middle">
<%="<br><br>&nbsp;预测结果如下：<br><br>&nbsp;<font color=blue>"&rs("content")&"</font><br><br>"%>
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