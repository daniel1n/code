<% 
const senlontitle="�ƴ�����ǩ"
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>�ƴ�����ǩ_ʵ�ò�ѯ���ߴ�ȫ - Powered by Senlon!</TITLE>

</head>
<body>
<!--#include file="../senlon/function.asp"-->
<!--#include file="../senlon/conn.asp"-->  
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
  <table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
<tbody><%if request("act")="jq" then
gyqid=request("gyqid")
set rs=server.createobject("adodb.recordset")
sql="select * from huangdaxian where id="&gyqid
rs.open sql,conn,1,3
qianshu=rs("qianshu")
hdxname=rs("name")
title=rs("title")
shi=rep(rs("shi"))
content=rs("content")
rs.close
%><tr><td class="new"><font color=red>�ƴ�����ǩ��</font><%=qianshu%>&nbsp;<%=hdxname%>&nbsp;<%=title%> 
    <br><br><font color=red>����:</font><br><br><%=shi%><br><br><font color=red>��ǩ:</font><br><%=content%><A href="./"><br><br><font color=red>������ﷵ�س�ǩ��ҳ��</font></A> </td>
</tr><%else%>
<tr><td align="center" class="new"><img src="../images/huangdaxian2.jpg" width="150" height="209"></td>
<td width="50%" align="center" class="new">
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
<a href="?act=ok&gyclicknum=<%=gyclicknum+1%>&gyqid=<%=num%>" title="���������������Ȼ������￪ʼ�S���}��"><img src=../images/sign<%=picnum%>.gif width=100 height=77 border=0></a><br><br>��������������ʥ����������ǩ���������ͼ�꿪ʼ����ʥ�� 
<%else
randomize
gysmile=int((rnd*4)+1)
response.Cookies("laisuanming")("gysmile"&gyclicknum&"")=cstr(gysmile)

if gyclicknum<3 and gysmile<>"4" then%><a href="?act=ok&gyclicknum=<%=gyclicknum+1%>&gyqid=<%=num%>" title="���������������Ȼ������￪ʼ�S���}��"><%end if%><img src=../images/sign<%=picnum%>.gif width=100 height=77 border=0></a><br><br><%if gyclicknum<>"" then%><%if gysmile <> "4" then%>ʥ��<br><%else%>Ц��<br><%end if%>
<%end if%><br><%if gyclicknum=3 and gysmile<>"4" then%><a href="./?act=jq&gyqid=<%=num%>"><font color=blue>��ϲ������������������ʥ�����������쿴ǩ�ʣ�</font></a><%else%>��������������ʥ����������ǩ��Ŀǰ�����Ѿ�����<%=gyclicknum%>��<%if gysmile = "4"  then%><br>
<a href="./"><font color=red>���ǣ�������Ц���ˣ���ǩ��׼������������³�ǩ��</font></a><br><%end if%>
<%end if%><%end if%>
<%else
response.Cookies("laisuanming")("guanyin")=""
%><br>
<a href="?act=ok" title="���������������Ȼ������￪ʼ��ǩ"><img src="../images/hfxcq.gif" width="163" height="165" border="0" /></a>
<DIV align="left" class="new2">1.��ǩǰ�Ⱦ��ֺ�˫�ֺ�ʮ��Ĭ�� �����ɴ��ɡ�ָ���Խ򡱡�<br />
2.Ĭ���Լ��������������ڣ����䡢���ھ�ס��ַ��<br />
3.����ָ������顣���������ҵ���˳̡����ꡢ����������......
�ȵȡ�
<br />
4.�������ǩͲ��ʼ��ǩ��</DIV>
<%end if%>
</td>
<td class="new" align="center">
<img src="../images/huangdaxian1.jpg" width="150" height="209" border="0" /></td>
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