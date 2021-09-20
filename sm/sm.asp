<% 
const senlontitle="八字算命"
%><!--#include file="../senlon/conn.asp"-->
<!--#include file="../senlon/function.asp"-->
<!--#include file="../senlon/cookies.asp"-->
<%
writecookies()
%>
<!--#include file="../senlon/getuc.asp"-->
<!--#include file="../senlon/bzi.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>生辰八字,八字测算,日干论命,称骨论命,姓名测试,姓命配对,上辈为人,姓氏起源</TITLE>
<style type="text/css">
<!--
.fontcolor1 {color:#0000ff;}
-->
</style>

</head>
<body>
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
<%
if xing<>"" then
userbirthday=nian1&"-"&yue1&"-"&ri1
userll=DateDiff("yyyy", userbirthday, date)
%>
  <table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
  <TBODY>    
      <tr>
        <TD class=ttop style="PADDING-BOTTOM: 1px" vAlign=top>当前算命者信息</TD>
      </tr>
    <tr>
      <TD class=new style="PADDING-BOTTOM: 8px" vAlign=top>姓名：<font color="#FF0000"><%=xing&ming%></font> 性别：<%=xingbie%> <%if xuexing<>"" then%> 血型：<%=xuexing%> 型
      <%end if%>         出生:<font color="#0000ff"><%=nian1%>年<%=yue1%>月<%=ri1%>日<%=hh1%>时<%=mm1%>分</font> 今年<%=userll%>岁 属相:<%=sx%>&nbsp;星座:<%=Constellation(nian1&"-"&yue1&"-"&ri1)%>&nbsp;<input type="button" value="退出算命" onClick="(location='cookieclear.asp')" style="cursor:hand;COLOR: #ff0000;FONT-WEIGHT: bold;" class="button" /></TD>
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
          <input type="button" value="EQ 曲线" onClick="(location='eq.asp')" style="cursor:hand;" class="button" /> 
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
</TABLE>     <%
		if request("sm")=1 then%><!--#include file="bz.asp"--> <%elseif request("sm")=2 then%><!--#include file="bzcs.asp"--><%elseif request("sm")=3 then%><!--#include file="nm.asp"--><%elseif request("sm")=4 then%><!--#include file="cg.asp"-->
 <%elseif request("sm")=5 then%><!--#include file="xm.asp"--><%elseif request("sm")=6 then%><!--#include file="xmpd.asp"--><%elseif request("sm")=7 then%><!--#include file="sbwr.asp"--><%elseif request("sm")=8 then%><!--#include file="x.asp"--><%end if%><%else%>
<script language="JavaScript">
<!--
function checkbz()
{
var year=document.sm.nian.value;
var month=document.sm.yue.value;
var day=document.sm.ri.value;
var hour=document.sm.hh.value;
var now=new Date();
var nowyear=now.getYear();
var nowmonth=now.getMonth();
if (year=='')
{
alert('请选择出生年份！');
document.sm.nian.focus()
return false;
}
//if (year>nowyear || year <=nowyear-100 || isNaN(year))
if (year=='')
{
alert('请选择正确的出生年份！');
document.sm.nian.focus()
return false;
}
if ( month=='')
{
alert('请选择出生月份！');
document.sm.yue.focus()
return false;
}
if (day=='')
{
alert('请选择出生日期！');
document.sm.ri.focus()
return false;
}
if ((month==2 && day>29) || ((month==1 || month==3 || month==5 || month==7 || month==8 || month==10|| month==12) && day>31) || ((month==4 || month==6 || month==9 || month==11 ) && day>30))
{
alert('请选择正确的出生日期！');
document.sm.ri.focus()
return false;
}
if ( hour=='')
{
alert('请选择出生时间！');
document.sm.hh.focus()
return false;
}
while(document.sm.xing.value.indexOf(" ")!=-1){
document.sm.xing.value=document.sm.xing.value.replace(" ","");
}
while(document.sm.xing.value.indexOf("	")!=-1){
document.sm.xing.value=document.sm.xing.value.replace("	","");
}
if (document.sm.xing.value=='')
{
alert('请输入您的姓氏！');
document.sm.xing.focus()
return false;
}
if (document.sm.xing.value.length < 1 || document.sm.xing.value.length>2)
{
alert("错误：姓氏应在1-2个字之间！");
document.sm.xing.focus();
return (false);
}

while(document.sm.ming.value.indexOf(" ")!=-1){
document.sm.ming.value=document.sm.ming.value.replace(" ","");
}
while(document.sm.ming.value.indexOf("	")!=-1){
document.sm.ming.value=document.sm.ming.value.replace("	","");
}
if (document.sm.ming.value=='')
{
alert('请输入您的名字！');
document.sm.ming.focus()
return false;
}
if (document.sm.ming.value.length < 1 || document.sm.ming.value.length>2)
{
alert("错误：名字应在1-2个字之间！");
document.sm.ming.focus();
return (false);
}
var b=document.sm.xingbie.value;
if (b=='')
{
alert('请选择您的性别！');
document.sm.xingbie.focus()
return false;
}
}
//-->
</script>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
 <form method="post" action="sm.asp?sm=<%=request("sm")%>" name="sm"  onSubmit="return checkbz();"> <TBODY>
       <tr>
      <TD class=ttop style="PADDING-BOTTOM: 1px" vAlign=top> 输入资料立刻开始免费算命</TD>
    </tr> <tr>
      <TD class=new style="PADDING-BOTTOM: 8px" vAlign=top><span class="12black"><b><img src="../images/help.gif" width="16" height="16">&nbsp;算命说明：</b>姓名必须输入中文，生日必须输入公历（公历即阳历/新历，农历即阴历/旧历）如果不分析血型，血型可任选；如果不分析八字等，出生时分可任选；不影响其它测试结果。进入系统后可体验数十种强大的算命功能！</span></TD>
    </tr>
    <tr>
      <TD class=new style="PADDING-BOTTOM: 8px" vAlign=top><div align="center"><a title="如果您只知道生日的农历日期，不要紧，请点这里去查询公历日期" style="CURSOR: hand" href="/wannianli/" target="_blank"><font color="red">[只知道农历请点此查公历]</font></a>&nbsp;<a title="不知道出生时间怎么办" style="CURSOR: hand" href="./htm_nobirth.asp" target="_blank"><font color="red">[不知道出生时间怎么办]</font></a>&nbsp;<a title="点击此处了解算命知识" style="CURSOR: hand" href="./htm_smzs.asp" target="_blank"><font color="red">[点击此处了解算命知识]</font></a></div></TD>
    </tr>
    <tr>
      <TD align="center" vAlign=top class=new style="PADDING-BOTTOM: 8px">姓：<input type="txt" name="xing" size="4" value="" onKeypress="if ((event.keyCode != 13 && event.keyCode < 160)) event.returnValue = false;">
  	名：<input type="txt" name="ming" size="4" value="" onKeypress="if ((event.keyCode != 13 && event.keyCode < 160)) event.returnValue = false;">
  	<select name="xingbie" size="1" style="font-size: 9pt">
	<option value="" selected>性别</option>
	<option value="男">男</option>
	<option value="女">女</option>
  	</select>
  	<select name="xuexing" size="1" style="font-size: 9pt">
  	<option value="">血型</option>
  	<option value="A">A型</option>
  	<option value="B">B型</option>
  	<option value="O">O型</option>
  	<option value="AB">AB型</option>
  	</select>
  	公历生日:
       <select name="nian" size="1" style="font-size: 9pt">
      ><%for i=1900 to year(date())%><option value="<%=i%>" <%if i=year(date()) then%> selected<%end if%>><%=i%></option><%next%>
     </select>
     年
     <select size="1" name="yue" style="font-size: 9pt">
      <%for i=1 to 12%><option value="<%=i%>"<%if i=month(date()) then%> selected<%end if%>><%=i%></option><%next%>
     </select>
     月
     <select size="1" name="ri" style="font-size: 9pt">
      <%for i=1 to 31%><option value="<%=i%>" <%if i=day(date()) then%> selected<%end if%>><%=i%></option><%next%>
     </select>
     日
     <select size="1" name="hh" style="font-size: 9pt">
       <%for i=0 to 23%><option value="<%=i%>" ><%=DiZhi(i)%>&nbsp;<%=i%></option><%next%> 
     </select>
     点
     <select size="1" name="mm" style="font-size: 9pt">
       <option value="0">未知</option>
		<%for i=0 to 59%><option value="<%=i%>"><%=i%></option><%next%>
     </select>
     分 </TD>
    </tr>
    <tr>
      <TD align="center" vAlign=top class=new style="PADDING-BOTTOM: 8px">&nbsp;<input type="submit" value="立刻算命" style='cursor:hand;COLOR: #ff0000;' class="button"> </TD>
    </tr>
  </TBODY></form>
</TABLE>
<%end if%>
      <!--#INCLUDE FILE="../dq.asp"-->
	</div>
    </td>
    <td valign="top"><!--#INCLUDE FILE="../right.asp"--></td>
  </tr>
</table>
<!--#include file="../foot.asp"-->
</body></html>
