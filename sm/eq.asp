<% 
const senlontitle="EQ����"
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>EQ����</TITLE>
<!--#include file="../senlon/conn.asp"-->
<!--#include file="../senlon/function.asp"--><!--#include file="../senlon/cookies.asp"-->
<%
writecookies()
%><!--#include file="../senlon/getuc.asp"-->

</head>
<body>
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <TBODY>
      <tr>
        <TD align="left" class=new style="PADDING-BOTTOM: 1px"><%
if xing<>"" then
myxz=Constellation(nian1&"-"&yue1&"-"&ri1)
%><%=xing&ming%>(<%=nian1%>-<%=yue1%>-<%=ri1%>) �ҵ�������<%=myxz%><a href="astro.asp?flag=4&astro=<%=myxz%>">[��������]</a> �ҵ�����:<%=sx%><a href="astro.asp?flag=5&xiao=<%=sx%>">[�����Ը�]</a> �ҵ�Ѫ��:<%=xuexing%><a href="astro.asp?flag=6&xuexing=<%=xuexing%>">[Ѫ�����]</a>&nbsp;<input type="button" value="�˳�����" onClick="(location='cookieclear.asp')" style="cursor:hand;COLOR: #ff0000;FONT-WEIGHT: bold;" class="button" /><%else%>��ѯ�ҵ�����:
    <form action="" method="post"><select name="y" size="1" id="y" style="font-size: 9pt"> 
            <%for i=1920 to year(date())%>
            <option value="<%=i%>" <%if i=1980 then%> selected<%end if%>><%=i%></option>
            <%next%>
          </select>
��
<select name="m" size="1" id="m" style="font-size: 9pt">
  <%for i=1 to 12%>
  <option value="<%=i%>"<%if i=month(date()) then%> selected<%end if%>><%=i%></option>
  <%next%>
</select>
��
<select name="d" size="1" id="d" style="font-size: 9pt">
  <%for i=1 to 31%>
  <option value="<%=i%>" <%if i=day(date()) then%> selected<%end if%>><%=i%></option>
  <%next%>
</select>
��(<a href="/tools/wannianli.htm" target="_blank">��������</a>)
<input name="Input2" type="submit" value="��ѯ" class="bot01"   /><input name="act" type="hidden" value="xzcx"><%if request("act")="xzcx" then
%>&nbsp;��ѯ���:<%=Constellation(""&request("y")&""&"-"&""&request("m")&""&"-"&""&request("d")&"")%></form>
<%end if%><%end if%></TD>
      </tr>
    <tr>
      <TD align="center" vAlign=top class=new style="PADDING-BOTTOM: 8px"><input name="button" type="button" class="button" style='cursor:hand;' onClick="(location='sm.asp?sm=1')" value="��������">
          <input name="button" type="button" class="button" style="cursor:hand;" onClick="(location='sm.asp?sm=2')" value="���ֲ���" />
          <input name="button" type="button" class="button" style="cursor:hand;" onClick="(location='sm.asp?sm=3')" value="�ո�����" />
          <input name="button" type="button" class="button" style="cursor:hand;" onClick="(location='sm.asp?sm=4')" value="�ƹ�����" />
          <input name="button" type="button" class="button" style="cursor:hand;" onClick="(location='sm.asp?sm=5')" value="��������" />
          <input name="button" type="button" class="button" style="cursor:hand;" onClick="(location='sm.asp?sm=6')" value="�������" />
          <input name="button" type="button" class="button" style="cursor:hand;" onClick="(location='sm.asp?sm=7')" value="�ϱ�Ϊ��" />
          <input name="button" type="button" class="button" style="cursor:hand;" onClick="(location='sm.asp?sm=8')" value="������Դ" />
          <input type="button" value="���屣��" style='cursor:hand;' onClick="(location='baojian.asp')" class="button">
          <input type="button" value="EQ����" onClick="(location='eq.asp')" style="cursor:hand;" class="button" /> 
          <input type="button" value="�����" onClick="(location='wu.asp')" class="button" />
          <input type="button" value="IQ ����" onClick="(location='iq.asp')" style="cursor:hand;" class="button" />
          <input type="button" value="ʧ������" onClick="(location='shibai.asp')" style="cursor:hand;" class="button" />
          <input type="button" value="��������" style='cursor:hand;' onClick="(location='mingren.asp')" class="button">
          <input type="button" value="�����Ը�" onClick="(location='astro.asp?flag=5&xiao=<%=sx%>')" style="cursor:hand;" class="button" />
          <input type="button" value="��������" onClick="(location='astro.asp?flag=8&flag1=������&m=<%=yue1%>&d=<%=ri1%>')" style="cursor:hand;" class="button" />
          <input type="button" value="���ջ���" onClick="(location='astro.asp?flag=8&flag1=���ջ�&m=<%=yue1%>&d=<%=ri1%>')" style="cursor:hand;" class="button" />
          <input type="button" value="Ѫ���Ը�" onClick="(location='astro.asp?flag=6&xuexing=<%=xuexing%>')" style="cursor:hand;" class="button" />
      </tr>
    </TBODY>
  </TABLE>
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <tbody>
      <tr>
        <td width="100%"  class="ttop"><%=myxz%>֮EQ����<%
		set rs=server.createobject("adodb.recordset")
		Sql="select * from ndeq where title='"&myxz&"'"
rs.open sql,conn,1,1
		%></td>
      </tr>
      <tr>
        <td class=new style="PADDING-BOTTOM: 1px" vAlign=top><%=rs("content")%></td>
      </tr>
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