<% 
const senlontitle="������ǩ"
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>������ǩ_ʵ�ò�ѯ���ߴ�ȫ - Powered by Senlon!</TITLE>

</head>
<body><!--#include file="../senlon/function.asp"--><!--#include file="../senlon/conn.asp"-->  
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
  <table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
<tbody><%if request("act")="jq" then
lvzulqid=request("lvzulqid")
%><tr><td align="center" class="new"><img src="../images/lvzu/<%=lvzulqid%>.gif"></td></tr><tr><td class="new"><A href="./"><font color=red>������ﷵ�س�ǩ��ҳ��</font></A> </td>
</tr><%else%>
<tr><td align="center" class="new"><img src="../images/lvzul.jpg" width="159" height="240"></td>
<td width="45%" align="center" class="new">
<%if request("act")="ok" then
on error resume next
if request.Cookies("laisuanming")("guanyin") = "" then
'���������
randomize
num=int((rnd*100)+1)
response.Cookies("laisuanming")("guanyin")=num
else
num=request.Cookies("laisuanming")("guanyin")
end if
randomize
picnum=int(rnd*3)+1
gyclicknum=request("gyclicknum")
%>���ճ��˵�&nbsp;<font style="color: #FF0000;FONT-SIZE: 26px;font-weight: bold;"><%=num%></font>&nbsp;ǩ<br><br>
<%
if gyclicknum="" then%>
<a href="?act=ok&gyclicknum=<%=gyclicknum+1%>&lvzulqid=<%=num%>" title="���������������Ȼ������￪ʼ�S���}��"><img src=../images/sign<%=picnum%>.gif width=100 height=77 border=0></a><br><br>��������������ʥ����������ǩ���������ͼ�꿪ʼ����ʥ�� 
<%else
randomize
gysmile=int((rnd*4)+1)
response.Cookies("laisuanming")("gysmile"&gyclicknum&"")=cstr(gysmile)

if gyclicknum<3 and gysmile<>"4" then%><a href="?act=ok&gyclicknum=<%=gyclicknum+1%>&lvzulqid=<%=num%>" title="���������������Ȼ������￪ʼ�S���}��"><%end if%><img src=../images/sign<%=picnum%>.gif width=100 height=77 border=0></a><br><br><%if gyclicknum<>"" then%><%if gysmile <> "4" then%>ʥ��<br><%else%>Ц��<br><%end if%>
<%end if%><br><%if gyclicknum=3 and gysmile<>"4" then%><a href="./?act=jq&lvzulqid=<%=num%>"><font color=blue>��ϲ������������������ʥ�����������쿴ǩ�ʣ�</font></a><%else%>��������������ʥ����������ǩ��Ŀǰ�����Ѿ�����<%=gyclicknum%>��<%if gysmile = "4"  then%><br>
<a href="./"><font color=red>���ǣ�������Ц���ˣ���ǩ��׼������������³�ǩ��</font></a><br><%end if%>
<%end if%><%end if%>
<%else
response.Cookies("laisuanming")("guanyin")=""
%><br><br>
<a href="?act=ok" title="���������������Ȼ������￪ʼ��ǩ"><img src="../images/yzqian.gif" width="103" height="134" border="0" /></a>
<br />
<DIV align="left" class="new2">1.��ǩǰ�Ⱥ���Ĭ�������ʦ ָ���Ծ������顣<BR>
 2.Ĭ���Լ�����������ʱ�������䡢���ھ�ס��ַ�� <BR>3.����ָ�����飬���������ҵ���˳̡����ꡢ����������......�ȡ�   <BR>4.�������ǩͲ��ʼ��ǩ��   </DIV>
<%end if%></td>
<td class="new" align="center">  <img src="../images/lvzu.jpg" width="159" height="240" border="0" /></td>
</tr>

<%end if%>
</tbody></table>
      <!--#INCLUDE FILE="../dq.asp"-->
	</div>
    </td>
    <td valign="top"><!--#INCLUDE FILE="../right.asp"--></td>
  </tr>
</table>
<!--#include file="../foot.asp"-->
</body></html>