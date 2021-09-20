<% 
const senlontitle="吕祖灵签"
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>吕祖灵签_实用查询工具大全 - Powered by Senlon!</TITLE>

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
%><tr><td align="center" class="new"><img src="../images/lvzu/<%=lvzulqid%>.gif"></td></tr><tr><td class="new"><A href="./"><font color=red>点击这里返回抽签首页！</font></A> </td>
</tr><%else%>
<tr><td align="center" class="new"><img src="../images/lvzul.jpg" width="159" height="240"></td>
<td width="45%" align="center" class="new">
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
<a href="?act=ok&gyclicknum=<%=gyclicknum+1%>&lvzulqid=<%=num%>" title="首先请您心无杂念，然后点这里开始擲出聖杯"><img src=../images/sign<%=picnum%>.gif width=100 height=77 border=0></a><br><br>需连续掷出三次圣杯，才是灵签！请点上面图标开始掷出圣杯 
<%else
randomize
gysmile=int((rnd*4)+1)
response.Cookies("laisuanming")("gysmile"&gyclicknum&"")=cstr(gysmile)

if gyclicknum<3 and gysmile<>"4" then%><a href="?act=ok&gyclicknum=<%=gyclicknum+1%>&lvzulqid=<%=num%>" title="首先请您心无杂念，然后点这里开始擲出聖杯"><%end if%><img src=../images/sign<%=picnum%>.gif width=100 height=77 border=0></a><br><br><%if gyclicknum<>"" then%><%if gysmile <> "4" then%>圣杯<br><%else%>笑杯<br><%end if%>
<%end if%><br><%if gyclicknum=3 and gysmile<>"4" then%><a href="./?act=jq&lvzulqid=<%=num%>"><font color=blue>恭喜，您连续掷出了三次圣杯，请点这里察看签词！</font></a><%else%>需连续掷出三次圣杯，才是灵签！目前，您已经掷出<%=gyclicknum%>次<%if gysmile = "4"  then%><br>
<a href="./"><font color=red>但是，您掷出笑杯了，此签不准，请点这里重新抽签！</font></a><br><%end if%>
<%end if%><%end if%>
<%else
response.Cookies("laisuanming")("guanyin")=""
%><br><br>
<a href="?act=ok" title="首先请您心无杂念，然后点这里开始抽签"><img src="../images/yzqian.gif" width="103" height="134" border="0" /></a>
<br />
<DIV align="left" class="new2">1.抽签前先合手默念“吕祖仙师 指点迷经”三遍。<BR>
 2.默念自己姓名、出生时辰、年龄、现在居住地址。 <BR>3.请求指点事情，如婚姻、事业、运程、流年、工作、财运......等。   <BR>4.点上面的签筒开始抽签。   </DIV>
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