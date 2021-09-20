<% 
const senlontitle="IQ测试"
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>IQ测试</TITLE>
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
%><%=xing&ming%>(<%=nian1%>-<%=yue1%>-<%=ri1%>) 我的星座：<%=myxz%><a href="astro.asp?flag=4&astro=<%=myxz%>">[星座详情]</a> 我的属相:<%=sx%><a href="astro.asp?flag=5&xiao=<%=sx%>">[属相性格]</a> 我的血型:<%=xuexing%><a href="astro.asp?flag=6&xuexing=<%=xuexing%>">[血型详解]</a>&nbsp;<input type="button" value="退出算命" onClick="(location='cookieclear.asp')" style="cursor:hand;COLOR: #ff0000;FONT-WEIGHT: bold;" class="button" /><%else%>查询我的星座:
    <form action="" method="post"><select name="y" size="1" id="y" style="font-size: 9pt"> 
            <%for i=1920 to year(date())%>
            <option value="<%=i%>" <%if i=1980 then%> selected<%end if%>><%=i%></option>
            <%next%>
          </select>
年
<select name="m" size="1" id="m" style="font-size: 9pt">
  <%for i=1 to 12%>
  <option value="<%=i%>"<%if i=month(date()) then%> selected<%end if%>><%=i%></option>
  <%next%>
</select>
月
<select name="d" size="1" id="d" style="font-size: 9pt">
  <%for i=1 to 31%>
  <option value="<%=i%>" <%if i=day(date()) then%> selected<%end if%>><%=i%></option>
  <%next%>
</select>
日(<a href="/tools/wannianli.htm" target="_blank">公历生日</a>)
<input name="Input2" type="submit" value="查询" class="bot01"   /><input name="act" type="hidden" value="xzcx"><%if request("act")="xzcx" then
%>&nbsp;查询结果:<%=Constellation(""&request("y")&""&"-"&""&request("m")&""&"-"&""&request("d")&"")%></form>
<%end if%><%end if%></TD>
      </tr>
    <tr>
      <TD align="center" vAlign=top class=new style="PADDING-BOTTOM: 8px"><input name="button" type="button" class="button" style='cursor:hand;' onClick="(location='sm.asp?sm=1')" value="生辰八字">
          <input name="button" type="button" class="button" style="cursor:hand;" onClick="(location='sm.asp?sm=2')" value="八字测算" />
          <input name="button" type="button" class="button" style="cursor:hand;" onClick="(location='sm.asp?sm=3')" value="日干论命" />
          <input name="button" type="button" class="button" style="cursor:hand;" onClick="(location='sm.asp?sm=4')" value="称骨论命" />
          <input name="button" type="button" class="button" style="cursor:hand;" onClick="(location='sm.asp?sm=5')" value="姓名测试" />
          <input name="button" type="button" class="button" style="cursor:hand;" onClick="(location='sm.asp?sm=6')" value="姓名配对" />
          <input name="button" type="button" class="button" style="cursor:hand;" onClick="(location='sm.asp?sm=7')" value="上辈为人" />
          <input name="button" type="button" class="button" style="cursor:hand;" onClick="(location='sm.asp?sm=8')" value="姓氏起源" />
          <input type="button" value="身体保健" style='cursor:hand;' onClick="(location='baojian.asp')" class="button">
          <input type="button" value="EQ曲线" onClick="(location='eq.asp')" style="cursor:hand;" class="button" /> 
          <input type="button" value="五大建议" onClick="(location='wu.asp')" class="button" />
          <input type="button" value="IQ 揭密" onClick="(location='iq.asp')" style="cursor:hand;" class="button" />
          <input type="button" value="失败剖析" onClick="(location='shibai.asp')" style="cursor:hand;" class="button" />
          <input type="button" value="星座名人" style='cursor:hand;' onClick="(location='mingren.asp')" class="button">
          <input type="button" value="属相性格" onClick="(location='astro.asp?flag=5&xiao=<%=sx%>')" style="cursor:hand;" class="button" />
          <input type="button" value="生日密码" onClick="(location='astro.asp?flag=8&flag1=生日书&m=<%=yue1%>&d=<%=ri1%>')" style="cursor:hand;" class="button" />
          <input type="button" value="生日花语" onClick="(location='astro.asp?flag=8&flag1=生日花&m=<%=yue1%>&d=<%=ri1%>')" style="cursor:hand;" class="button" />
          <input type="button" value="血型性格" onClick="(location='astro.asp?flag=6&xuexing=<%=xuexing%>')" style="cursor:hand;" class="button" />
      </tr>
    </TBODY>
  </TABLE>
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <tbody>
      <tr>
        <td width="100%"  class="ttop"><%=myxz%>之IQ测试
        <%
		set rs=server.createobject("adodb.recordset")
		Sql="select * from xlcz where title='"&myxz&"'"
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