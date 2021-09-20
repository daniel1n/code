<% 
const senlontitle="手指指纹算命"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>手指指纹算命_实用查询工具大全 - Powered by Senlon!</TITLE>
<!--#include file="../senlon/getuc.asp"-->

</head>
<body><SCRIPT language="JavaScript">
<!--
function submitchecken() {
if (document.form1.zwdm.value == "") {
alert("请输入您的指纹代码！");
document.form1.zwdm.focus();
return false;
}
if (document.form1.zwdm.value.length != 5) {
alert("指纹代码输入出错，应该为5个X或O的字母！");
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
  <td width="57%" class="ttd"><span class="red">手指指纹算命:</span>
    <br><br>据研究报告发现，人的性格是与生具来的。令人奇怪的是，人的指纹也是终生不变的。世界上绝对找不到两个指纹完全相同的人，所以“指纹”就被当作犯罪侦查上的重要线索之一。虽然我们的身体是由遗传所造成，且随着环境会发生变化，只有指纹始终不会发生变化。指纹，大致可分为<strong>“涡纹”（又叫斗或叫箩）和“流纹”（又叫簸箕）</strong>两种。随着形状的不同，其性格和命运也不相同。想不想得知其中之奥妙？在下面的测算中，您可以通过指纹研究人的性格，辅助使您对自己或他人的性格有一定的了解，在生活中泰然处之。方法如下：<font color=red>男左手，女右手。</font><font color=blue>从拇指开始，斗（或叫箩）用O(OPQ的O，不是零0），簸箕用X（XYZ的X）代表。输入5个指纹代码，然后按“立刻测算”按钮，即可算出结果。</font><br><br></td>
    <td width="43%" class="ttd"><img src="../images/zw.gif" width="258" height="200"></td>
</tr>
<form name="form1" onSubmit="return submitchecken()" method="post" action="">
<input type="hidden" name="act" value="ok" />
  <tr>
    <td colspan="2" class="new">
请输入您的指纹代码：<input name="zwdm" type="text" id="zwdm" size="20" maxLength="5"> 
<input type="submit" name="Submit1" value="立刻测算" style="cursor:hand;"></form></td>
    </tr>
<%
'计算诸葛神数
on error resume next
dim zwdm,sql
zwdm=request("zwdm")
sql="select * from zhiwen where zhiwen='"&zwdm&"'"
set rs=conn.execute(sql)
if not (rs.bof and rs.eof) then
%>
<tr bgcolor="#EFF8FE">
<td class="new" colspan="2" valign="middle">
<%=""&xing&""&ming&"您的指纹代码为：<font color=blue>"&rs("zhiwen")&"</font><br><br>"%>
<%="性格解析：<font color=ff1100>"&rs("jiexi")&"</font>"%>
</td>
</tr><%else%><tr bgcolor="#EFF8FE">
<td class="new" colspan="2" valign="middle">
请在上方输入您的指纹代码！</td>
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