<% 
const senlontitle="��ָָ������"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>��ָָ������_ʵ�ò�ѯ���ߴ�ȫ - Powered by Senlon!</TITLE>
<!--#include file="../senlon/getuc.asp"-->

</head>
<body><SCRIPT language="JavaScript">
<!--
function submitchecken() {
if (document.form1.zwdm.value == "") {
alert("����������ָ�ƴ��룡");
document.form1.zwdm.focus();
return false;
}
if (document.form1.zwdm.value.length != 5) {
alert("ָ�ƴ����������Ӧ��Ϊ5��X��O����ĸ��");
document.form1.zwdm.focus();
return false;
}
return true;
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
  <td width="57%" class="ttd"><span class="red">��ָָ������:</span>
    <br><br>���о����淢�֣��˵��Ը������������ġ�������ֵ��ǣ��˵�ָ��Ҳ����������ġ������Ͼ����Ҳ�������ָ����ȫ��ͬ���ˣ����ԡ�ָ�ơ��ͱ�������������ϵ���Ҫ����֮һ����Ȼ���ǵ����������Ŵ�����ɣ������Ż����ᷢ���仯��ֻ��ָ��ʼ�ղ��ᷢ���仯��ָ�ƣ����¿ɷ�Ϊ<strong>�����ơ����ֽж�����ᣩ�͡����ơ����ֽ�������</strong>���֡�������״�Ĳ�ͬ�����Ը������Ҳ����ͬ���벻���֪����֮���������Ĳ����У�������ͨ��ָ���о��˵��Ը񣬸���ʹ�����Լ������˵��Ը���һ�����˽⣬��������̩Ȼ��֮���������£�<font color=red>�����֣�Ů���֡�</font><font color=blue>��Ĵָ��ʼ����������ᣩ��O(OPQ��O��������0����������X��XYZ��X����������5��ָ�ƴ��룬Ȼ�󰴡����̲��㡱��ť��������������</font><br><br></td>
    <td width="43%" class="ttd"><img src="../images/zw.gif" width="258" height="200"></td>
</tr>
<form name="form1" onSubmit="return submitchecken()" method="post" action="">
<input type="hidden" name="act" value="ok" />
  <tr>
    <td colspan="2" class="new">
����������ָ�ƴ��룺<input name="zwdm" type="text" id="zwdm" size="20" maxLength="5"> 
<input type="submit" name="Submit1" value="���̲���" style="cursor:hand;"></form></td>
    </tr>
<%
'�����������
on error resume next
dim zwdm,sql
zwdm=request("zwdm")
sql="select * from zhiwen where zhiwen='"&zwdm&"'"
set rs=conn.execute(sql)
if not (rs.bof and rs.eof) then
%>
<tr bgcolor="#EFF8FE">
<td class="new" colspan="2" valign="middle">
<%=""&xing&""&ming&"����ָ�ƴ���Ϊ��<font color=blue>"&rs("zhiwen")&"</font><br><br>"%>
<%="�Ը������<font color=ff1100>"&rs("jiexi")&"</font>"%>
</td>
</tr><%else%><tr bgcolor="#EFF8FE">
<td class="new" colspan="2" valign="middle">
�����Ϸ���������ָ�ƴ��룡</td>
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