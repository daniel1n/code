<% 
const senlontitle="������ǩ"
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>������ǩ_ʵ�ò�ѯ���ߴ�ȫ - Powered by Senlon!</TITLE>

<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
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
sql="select * from guanyin where id="&gyqid
rs.open sql,conn,1,3
jieqian=rs("jieqian")
rs.close
%><tr><td align="center" class="new"><img src="../images/gy/<%=gyqid%>.gif"><br><hr>
</td></tr><tr><td class="new">������ǩ��<%=jieqian%>
    <br><br><A href="./"><font color=red>������ﷵ�س�ǩ��ҳ��</font></A> </td>
</tr><%else%>
<tr><td align="center" class="new">
  <script type="text/javascript">
AC_FL_RunContent( 'codebase','http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0','width','152','height','240','src','../images/gy','quality','high','pluginspage','http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash','movie','../images/gy' ); //end AC code
</script><noscript><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="152" height="240">
<param name="movie" value="../images/gy.swf" />
<param name="quality" value="high" />
<embed src="../images/gy.swf" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="152" height="240">
</embed> 
</object></noscript>
</td>
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
<a href="?act=ok&gyclicknum=<%=gyclicknum+1%>&gyqid=<%=num%>" title="���������������Ȼ������￪ʼ�S���}��"><img src=../images/sign<%=picnum%>.gif width=100 height=77 border=0></a><br><br>��������������ʥ����������ǩ���������ͼ�꿪ʼ����ʥ��<br>
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
<a href="?act=ok" title="���������������Ȼ������￪ʼ��ǩ"><img src="../images/qian.gif" width="97" height="189" border="0" /></a><DIV align="left" class="new2">
1.��ǩǰ�Ⱥ���Ĭ��ȿ���� �������������顣
<br />
2.Ĭ���Լ�����������ʱ�������䣬���ھ�ס��ַ��
<br />
3.����ָ�����飬���������ҵ���˳̡����ꡢ����������......�ȡ�
<br />
4.�������ǩͲ��ʼ��ǩ��</DIV>
<%end if%>
</td>
<td class="new" align="center">  <img src="../images/guanyin.jpg" width="145" height="240" border="0" /></td>
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
