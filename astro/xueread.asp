<% 
const senlontitle="Ѫ�ͽ��"
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<!--#include file="../senlon/conn.asp"-->
<!--#include file="../senlon/function.asp"-->
<!--#include file="../senlon/getuc.asp"-->
          <%flag=request("flag")
		  if flag=1 then
		astro_m=request("astro_m")
		astro_f=request("astro_f")
		title="������ "&astro_m&" .vs."&astro_f&""
		elseif flag=2 then
		 astro_m=request("astro_m")
		astro_f=request("astro_f")
		title=""&astro_m&"--"&astro_f&""
				elseif flag=3 then
		 xiao_m=request("xiao_m")
		xiao_f=request("xiao_f")
		title=""&xiao_m&"+"&xiao_f&""
		elseif flag=4 then
		title=request("astro")
		elseif flag=5 then
		title="��"&request("xiao")&"�˵��Ը�"
		elseif flag=6 then
		title=""&request("xuexing")&"��Ѫ�˵��Ը�����"
		elseif flag=7 then
		title=""&request("blood")&"��"&request("astro")&""
		elseif flag=8 then
		title=""&request("m")&"��"&request("d")&"��"&request("flag1")&""
		elseif flag=9 then
		title=""&request("11-1")&"����"&request("11-2")&""
		elseif flag=10 then
		title=""&request("12-1")&"����"&request("12-2")&""
		elseif flag=11 then
		title=""&request("star1")&"��"&request("star2")&"��"&request("13-3")&""
		elseif flag=12 then
		title="2013��ʮ����Ф����--"&request("sx")&""
		end if	
		set rs=server.createobject("adodb.recordset")
		Sql="select * from astro where title='"&title&"'"
rs.open sql,conn,1,1%>
<TITLE><%=title%>Ѫ�ͽ��_ʵ�ò�ѯ���ߴ�ȫ - Powered by Senlon!</TITLE>

</head>
<body>
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
  <table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <TBODY>
      <tr>
        <TD align="center" vAlign=top class=new style="PADDING-BOTTOM: 1px">
        <span class="red"><%=title%></span>
        <hr noshade color="#CCCCCC"></TD>
      </tr>
      <tr>
        <TD class=new style="PADDING-BOTTOM: 1px" vAlign=top><%=rs("content")%></TD>
      </tr>
	  <tr>
        <TD height='45' align="center"><a href="./"><img src="../images/fanhui.gif" alt="�����ť����" border="0" /></a></TD>
      </tr>
    </TBODY>
  </TABLE>
      <!--#INCLUDE FILE="../dq.asp"-->
	</div>
    </td>
    <td valign="top"><!--#INCLUDE FILE="../right.asp"--></td>
  </tr>
</table>
<!--#include file="../foot.asp"-->
</body></html>