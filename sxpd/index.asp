<% 
const senlontitle="��Ф���"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>��Ф���_ʵ�ò�ѯ���ߴ�ȫ - Powered by Senlon!</TITLE>

</head>
<body><!--#include file="../senlon/function.asp"--><!--#include file="../senlon/conn.asp"-->  
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
  <table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
<tbody><tr>
  <td width="100%" class="ttd"><span class="red">��Ф���:</span>
  <br><br>��������ţ�������á������ߡ����򡢺����������ʮ���֡������ֶ��������������ֵģ��������������������˵ĵ���������������ܺ����ദ������˵�����Ͳ���ô�����ˡ����Ƿ񻹶����İ�����ԥ���������Ǿͽ����������ǵ���Ф�䲻��ɣ�<br><br></td>
</tr>
<form name="form1"  method="post" action="">
<input type="hidden" name="act" value="sxok" />
  <tr>
    <td colspan="2" class="new">
�ҵ���Ф��
  <SELECT name="p5" class="style3">
<OPTION value=��<%=str1%>>��</OPTION>
<OPTION value=ţ<%=str2%>>ţ</OPTION>
<OPTION value=��<%=str3%>>��</OPTION>
<OPTION value=��<%=str4%>>��</OPTION>
<OPTION value=��<%=str5%>>��</OPTION>
<OPTION value=��<%=str6%>>��</OPTION>
<OPTION value=��<%=str7%>>��</OPTION>
<OPTION value=��<%=str8%>>��</OPTION>
<OPTION value=��<%=str9%>>��</OPTION>
<OPTION value=��<%=str10%>>��</OPTION>
<OPTION value=��<%=str11%>>��</OPTION>
<OPTION value=��<%=str12%>>��</OPTION>
</SELECT>
  ��/������Ф
  :
  <SELECT name="p6" class="style3">
<OPTION value=�� selected>��</OPTION>
<OPTION value=ţ>ţ</OPTION>
<OPTION value=��>��</OPTION>
<OPTION value=��>��</OPTION>
<OPTION value=��>��</OPTION>
<OPTION value=��>��</OPTION>
<OPTION value=��>��</OPTION>
<OPTION value=��>��</OPTION>
<OPTION value=��>��</OPTION>
<OPTION value=��>��</OPTION>
<OPTION value=��>��</OPTION>
<OPTION value=��>��</OPTION>
</SELECT>
  <input type="submit" name="Submit1" value="��ʼ���" style="cursor:hand;">
    </form></td>
    </tr><%
	if request("act")="sxok" then
dim sqlstr
sqlstr="select * from shuxianglove where shengxiao1='"&trim(cstr(request("p5")))&"' and shengxiao2='"&trim(cstr(request("p6")))&"'"
set rs=conn.execute(sqlstr)

%>
<tr bgcolor="#EFF8FE">
<td class="new" colspan="2" valign="middle">
<%="<br>��Ф��ԣ�<font color=blue>"&rs("title")&"</font><br><br>"%>
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
