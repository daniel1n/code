<% 
const senlontitle="Ѫ�����"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>Ѫ�����_ʵ�ò�ѯ���ߴ�ȫ - Powered by Senlon!</TITLE>

</head>
<body><!--#include file="../senlon/function.asp"--><!--#include file="../senlon/conn.asp"-->  
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
<tbody><tr>
  <td width="100%" class="ttd"><span class="red">Ѫ�����:</span>
  <br><br>����˫����Ѫ�ͣ����Բ��Գ��������˵�Ե�ֺͻ���Ŷ���������԰ɣ�<br><br></td>
</tr>
<form name="form1"  method="post" action="">
<input type="hidden" name="act" value="xxok" />
  <tr>
    <td colspan="2" class="new">
�ҵ�Ѫ�ͣ�
  <SELECT name="p3" class="style3">
<OPTION value=AB<%=str1%>>AB��</OPTION>
<OPTION value=A<%=str2%>>A��</OPTION>
<OPTION value=B<%=str3%>>B��</OPTION>
<OPTION value=O<%=str4%>>O��</OPTION>
</SELECT>
  ��/����Ѫ��
  :
  <SELECT name="p4" class="style3">
<OPTION value=AB>AB��</OPTION>
<OPTION value=A>A��</OPTION>
<OPTION value=B>B��</OPTION>
<OPTION value=O>O��</OPTION>
</SELECT>
  <input type="submit" name="Submit1" value="��ʼ���" style="cursor:hand;">
    </form></td>
    </tr><%
	if request("act")="xxok" then
sqlstr="select * from xuexinglove where xuexing1='"&trim(cstr(request("p3")))&"' and xuexing2='"&trim(cstr(request("p4")))&"'"
set rs=conn.execute(sqlstr)

%>
<tr bgcolor="#EFF8FE">
<td class="new" colspan="2" valign="middle">
<%="<br>Ѫ����ԣ�<font color=blue>"&rs("title1")&"</font><br><br>"%>
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
