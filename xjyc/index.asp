<% 
const senlontitle="�ľ�Ԥ��"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>�ľ�Ԥ��,����Ԥ��_ʵ�ò�ѯ���ߴ�ȫ - Powered by Senlon!</TITLE>
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
  <td width="100%" class="ttd"><span class="red">�ľ�Ԥ��:</span>
    <br><br>�е����ǻ�Ī�����ľ���Ϊ���������ܷ��겻���������Լ�Ҫ�������������ľ�Ԥ����ʲô���ľ���ԭ���ľ���������ʲô�����������ľ�������ô˵�ɣ�<br><br></td>
</tr>
<form name="form1" onSubmit="return submitchecken()" method="post" action="">
<input type="hidden" name="act" value="ok" />
  <tr>
    <td colspan="2" class="new">&nbsp;�ľ�ʱ�䣺<select id="stime" style="WIDTH: 100px" name="stime" size="1">
<option value="1" selected="">23-01[��ʱ]</option>
<option value="2">01-03[��ʱ]</option>
<option value="3">03-05[��ʱ]</option>
<option value="4">05-07[îʱ]</option>
<option value="5">07-09[��ʱ]</option>
<option value="6">09-11[��ʱ]</option>
<option value="7">11-13[��ʱ]</option>
<option value="8">13-15[δʱ]</option>
<option value="9">15-17[��ʱ]</option>
<option value="10">17-19[��ʱ]</option>
<option value="11">19-21[��ʱ]</option>
<option value="12">21-23[��ʱ]</option>
</select> <input type="submit" value="��ʼ����" style="cursor:hand;" /></form></td>
    </tr>
<%
on error resume next
fx=request("fx")
stime=request("stime")
sql="select * from msyuce where stime='"&stime&"' and lb='�ľ�'"
set rs=conn.execute(sql)
if not (rs.bof and rs.eof) then
%>
<tr bgcolor="#EFF8FE">
<td class="new" colspan="2" valign="middle">
<%="<br><br>&nbsp;Ԥ�������£�<br><br>&nbsp;<font color=blue>"&rs("content")&"</font><br><br>"%>
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