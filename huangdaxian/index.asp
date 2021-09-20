<% 
const senlontitle="黄大仙灵签"
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>黄大仙灵签_实用查询工具大全 - Powered by Senlon!</TITLE>

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
%><tr><td class="new"><font color=red>黄大仙灵签：</font><%=qianshu%>&nbsp;<%=hdxname%>&nbsp;<%=title%> 
    <br><br><font color=red>卦语:</font><br><br><%=shi%><br><br><font color=red>解签:</font><br><%=content%><A href="./"><br><br><font color=red>点击这里返回抽签首页！</font></A> </td>
</tr><%else%>
<tr><td align="center" class="new"><img src="../images/huangdaxian2.jpg" width="150" height="209"></td>
<td width="50%" align="center" class="new">
<%if request("act")="ok" then
on error resume next
if request.Cookies("laisuanming")("guanyin") = "" then
'产生随机数
randomize
num=int((rnd*100)+1)
response.Cookies("laisuanming")("guanyin")=num
else
num=request.Cookies("laisuanming")("guanyin")
end if
randomize
picnum=int(rnd*3)+1
gyclicknum=request("gyclicknum")
%>您刚抽了第&nbsp;<font style="color: #FF0000;FONT-SIZE: 26px;font-weight: bold;"><%=num%></font>&nbsp;签<br><br>
<%
if gyclicknum="" then%>
<a href="?act=ok&gyclicknum=<%=gyclicknum+1%>&gyqid=<%=num%>" title="首先请您心无杂念，然后点这里开始擲出聖杯"><img src=../images/sign<%=picnum%>.gif width=100 height=77 border=0></a><br><br>需连续掷出三次圣杯，才是灵签！请点上面图标开始掷出圣杯 
<%else
randomize
gysmile=int((rnd*4)+1)
response.Cookies("laisuanming")("gysmile"&gyclicknum&"")=cstr(gysmile)

if gyclicknum<3 and gysmile<>"4" then%><a href="?act=ok&gyclicknum=<%=gyclicknum+1%>&gyqid=<%=num%>" title="首先请您心无杂念，然后点这里开始擲出聖杯"><%end if%><img src=../images/sign<%=picnum%>.gif width=100 height=77 border=0></a><br><br><%if gyclicknum<>"" then%><%if gysmile <> "4" then%>圣杯<br><%else%>笑杯<br><%end if%>
<%end if%><br><%if gyclicknum=3 and gysmile<>"4" then%><a href="./?act=jq&gyqid=<%=num%>"><font color=blue>恭喜，您连续掷出了三次圣杯，请点这里察看签词！</font></a><%else%>需连续掷出三次圣杯，才是灵签！目前，您已经掷出<%=gyclicknum%>次<%if gysmile = "4"  then%><br>
<a href="./"><font color=red>但是，您掷出笑杯了，此签不准，请点这里重新抽签！</font></a><br><%end if%>
<%end if%><%end if%>
<%else
response.Cookies("laisuanming")("guanyin")=""
%><br>
<a href="?act=ok" title="首先请您心无杂念，然后点这里开始抽签"><img src="../images/hfxcq.gif" width="163" height="165" border="0" /></a>
<DIV align="left" class="new2">1.抽签前先净手后双手合十虔诚默念 “大仙大仙、指点迷津”。<br />
2.默念自己姓名、出生日期，年龄、现在居住地址。<br />
3.请求指点的事情。如婚姻、事业、运程、流年、工作、财运......
等等。
<br />
4.点上面的签筒开始抽签！</DIV>
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