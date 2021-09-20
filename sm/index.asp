<% 
const senlontitle="免费在线算命"
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>免费在线算命,姓名算命分析,算命最准的网站_实用查询工具大全 - Powered by Senlon!</TITLE>

</head>
<body><!--#include file="../senlon/conn.asp"-->
<!--#include file="../senlon/function.asp"-->
<!--#include file="../senlon/getuc.asp"-->
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
<%
if xing<>"" then
%>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
  <TBODY>
    <tr>
      <TD class=ttop style="PADDING-BOTTOM: 1px" vAlign=top>当前姓名测算者信息</TD>
    </tr>
    <tr>
      <TD class=new style="PADDING-BOTTOM: 8px" vAlign=top>姓名：<font color="#FF0000"><%=xing&ming%></font> 性别：<%=xingbie%>
          <%if xuexing<>"" then%>
        血型：<%=xuexing%> 型
        <%end if%>
        出生:<font color="#0000ff"><%=nian1%>年<%=yue1%>月<%=ri1%>日<%=hh1%>时<%=mm1%>分</font> 今年<%=userll%>岁 属相:<%=sx%>&nbsp;星座:<%=Constellation(nian1&"-"&yue1&"-"&ri1)%>&nbsp;<input type="button" value="退出算命" onClick="(location='cookieclear.asp')" style="cursor:hand;COLOR: #ff0000;FONT-WEIGHT: bold;" class="button" /></TD>
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
<%else%><script language="JavaScript">
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
//if (year>nowyear || year <=nowyear-10 || isNaN(year))
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
 <form method="post" action="sm.asp?sm=1" name="sm"  onSubmit="return checkbz();"> <TBODY>
       <tr>
      <TD class=ttop style="PADDING-BOTTOM: 1px" vAlign=top> 免费算命</TD>
    </tr> <tr>
      <TD class=new style="PADDING-BOTTOM: 8px" vAlign=top><span class="12black"><b><img src="../images/help.gif" width="16" height="16">&nbsp;算命说明：</b>姓名必须输入中文，生日必须输入公历（公历即阳历/新历，农历即阴历/旧历）；如果不分析血型、血型可任选，如果不分析八字等、出生时分可任选，不影响其它测试结果。</span></TD>
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
      ><%for i=1930 to year(date())%><option value="<%=i%>" <%if i=year(date()) then%> selected<%end if%>><%=i%></option><%next%>
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
      <TD align="center"  vAlign=middle class=new style="PADDING-BOTTOM: 8px">
	  &nbsp;<input type="submit" value="开始算命" style='cursor:hand;COLOR: #ff0099;' class="button">
</TD></tr>	
  </TBODY></form>
</TABLE>
  <%end if%>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <TBODY>
      <tr>
        <TD colspan="3"  vAlign=top class=ttop style="PADDING-BOTTOM: 1px">抽签占卜</TD>
      </tr>
      <tr>
        <TD align="center" vAlign=top class=new style="PADDING-BOTTOM: 1px"><a href="/guanyin/" target="_blank"><img src="../images/gylq.gif" alt="观音灵签" width="176" height="89" border="0"></a></TD>
        <TD align="center" vAlign=top class=new style="PADDING-BOTTOM: 1px"><a href="/lvzu/" target="_blank"><img src="../images/yzlq.gif" alt="吕祖灵签" width="172" height="88" border="0"></a></TD>
        <TD align="center" vAlign=top class=new style="PADDING-BOTTOM: 1px"><a href="/huangdaxian/" target="_blank"><img src="../images/dxlq.gif" alt="黄大仙灵签" width="174" height="88" border="0"></a></TD>
        </tr>
      <tr>
        <TD align="center" vAlign=top class=new style="PADDING-BOTTOM: 1px"><a href="/guandi/" target="_blank"><img src="../images/dmgdm.gif" alt="关帝神签" width="174" height="88" border="0"></a></TD>
        <TD align="center" vAlign=top class=new style="PADDING-BOTTOM: 1px"><a href="/mazu/" target="_blank"><img src="../images/thlq.gif" alt="天后灵签" width="174" height="86" border="0"></a></TD>
        <TD align="center" vAlign=top class=new style="PADDING-BOTTOM: 1px"><a href="/zgss/" target="_blank"><img src="../images/zgcz.gif" alt="诸葛神算" width="172" height="88" border="0"></a></TD>
      </tr>
    </TBODY>
  </TABLE>
  <table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1">
  <tbody>
    <tr>
      <td colspan="3"  valign="top" class="ttop" style="PADDING-BOTTOM: 1px">在线排盘</td>
    </tr>
    <tr>
      <td width="13%" valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/ppbazi/" target="_blank"><img src="../images/pg1.gif" alt="四柱八字在线排盘" width="80" height="80" border="0" /></a></td>
      <td width="5%" valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/ppbazi/" target="_blank"  class="red">四柱八字</a></td>
      <td width="82%" valign="top" class="ttd" style="PADDING-BOTTOM: 1px">四柱八字（又称“生辰八字”）以人出生的年、月、日、时为四柱，合四柱之干支为八字。具体是由“年干、年支”为年柱，“月干、月支”为月柱，“日干、日支”为日柱，“时干、时支”为时柱。共四种组合（四柱），八个干支所组成（八字），故称“四柱”或“八字”全称“四柱八字”。<a href="/ppbazi/" target="_blank" class="red">点击进入>></a></td>
    </tr>
    <tr>
      <td width="13%" valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/pp6r/" target="_blank"><img src="../images/pg2.gif" alt="六壬在线排盘系统" width="80" height="80" border="0"></a></td>
      <td width="5%" valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a  class="red" href="/pp6r/" target="_blank">大六壬</a></td>
      <td width="82%" valign="top" class="ttd" style="PADDING-BOTTOM: 1px">六壬是用阴阳五行占卜吉凶的一种术数。六壬与遁甲、太乙合称三式。五行以水为首，十天干中，壬、癸分别为阳水、阴水。舍阴取阳，六十甲子中壬有六个（壬申、壬午、壬辰、壬寅、壬子、壬戌），称为六壬。六壬有六十四课，以刻有干支的天盘、地盘相叠，转动天盘后得出所值的干支及时辰，判明吉凶。<a href="/pp6r/" target="_blank" class="red">点击进入>></a></td>
    </tr>
	<tr>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/ppjkj/" target="_blank"><img src="../images/pg7.gif" alt="金口诀在线排盘系统" width="80" height="80" border="0"></a></td>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/ppjkj/" class="red" target="_blank">金口诀</a></td>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px">提到金口诀就不可以不说大六壬，因为金口诀的全称就是"大六壬神课金口诀"。因此，可知金口诀是源于大六壬脱胎形成的新的一种预测科学。金口诀自孙膑创立后，在其传人后代中口口相传，不留文字，一直不为外人所知。我们现在可以看到的有关金口诀的著作最早的就是明朝时期的《大六壬神课金口诀》。<a href="/ppjkj/" target="_blank" class="red">点击进入>></a></td>
    </tr>
    <tr>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/pp6y/" target="_blank"><img src="../images/pg3.gif" alt="纳甲六爻在线排卦系统" width="80" height="80" border="0"></a></td>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/pp6y/" class="red" target="_blank">六爻起卦</a></td>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px">六爻是我国传统预测方法的一种，往往和八字等同论。叫六爻基本是民间普及的叫法，它的别名还有“火珠林预测”“纳甲预测”“周易预测”等，书本上多称之为周易预测。<a href="/pp6y/" target="_blank" class="red">点击进入>></a></td>
    </tr>
    <tr>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/ppqmdj/" target="_blank"><img src="../images/pg4.gif" alt="奇门遁甲在线排盘" width="80" height="80" border="0"></a></td>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/ppqmdj/" class="red" target="_blank">奇门遁甲</a></td>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px">奇门遁甲，原来是中国古老的一本书，但它往往被认为是一本占卜用的书，但有的说法是说《奇门遁甲》是我国古代人民在同大自然作斗争中，经过长期观察、反复验证，总结出来的一门传统珍贵文化遗产。还有的说“奇门遁甲”是修真的功法。<a href="/ppqmdj/" target="_blank" class="red">点击进入>></a></td>
    </tr>
    <tr>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/ppxk/" target="_blank"><img src="../images/pg5.gif" alt="玄空飞星在线排盘" width="80" height="80" border="0"></a></td>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/ppxk/" class="red" target="_blank">玄空飞星</a></td>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px">玄就是时间，空就是空间，玄空学就是一门研究时间和空间为人造福的学问。玄空学由三大体系构成：河图，洛书，八卦。河图必须要记忆的是：一六为水，二七为火，三八为木，四九为金，五十为土。河图为先天。洛书必须要记忆的是：戴九履一，左三右七，二四为肩，六八为足。洛书为后天。八卦有先天八卦，有后天八卦，先天配河图，后天配洛书。先后天八卦的方位一定要记牢。<a href="/ppxk/" target="_blank" class="red">点击进入>></a></td>
    </tr>
    <tr>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/ppziwei/" target="_blank"><img src="../images/pg6.gif" alt="紫微斗数在线排盘系统" width="80" height="80" border="0"></a></a></td>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/ppziwei/" class="red" target="_blank">紫微斗数</a></td>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px">紫微斗数，是中国传统命理学的最重要的支派之一。她是以人出生的年、月、日、时确定十二宫的位置，构成命盘，结合各宫的星群组合，牵系周易卦爻，来预测一个人的名员流程、吉凶祸福的。<a href="/ppziwei/" target="_blank" class="red">点击进入>></a></td>
    </tr>
    <tr>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/taiyang/" target="_blank"><img src="../images/zty.jpg" alt="真太阳时间查询修正" width="80" height="80" border="0" /></a></td>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/taiyang/" class="red" target="_blank">真太阳时间</a></td>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px">时辰究竟应该按照什么标准确定呢？是按北京时间确定？还是按当地时间确定？其实唯一的标准就是太阳相对于地球的视角，即天文学上所谓的真太阳时。即：真太阳时＝平太阳时+真平太阳时差。凡定生时，必须按照其出生地点，推算出当地的平太阳时，再根据平太阳时推算出真太阳时为准，不能简单地沿用北京时间。<a href="/taiyang/" target="_blank" class="red">点击进入>></a></td>
    </tr>
    <tr>
      <td valign="top" class="new" style="PADDING-BOTTOM: 1px"><a href="/jingdu/" target="_blank"><img src="../images/dqjw.jpg" width="80" height="80" border="0" /></a></td>
      <td valign="top" class="new" style="PADDING-BOTTOM: 1px"><a href="/jingdu/" class="red" target="_blank">地区经度</a></td>
      <td valign="top" class="new" style="PADDING-BOTTOM: 1px">经度泛指球面坐标系的纵坐标。定义为地球面上一点与两极的连线与0度经线所在平面的夹角。以球面上的点所在辅圈相对于坐标原点所在辅圈的角距离来表示。通常特指地理坐标的经度。为了区分地球上的每一个地区，人们给经线标注了度数，这就是经度。实际上经度是两条经线所在平面之间的夹角。<a href="/jingdu/" target="_blank" class="red">点击进入>></a></td>
    </tr>
  </tbody>
</table>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1">
    <TBODY>
      <tr>
        <TD colspan="3"  vAlign=top class=ttop style="PADDING-BOTTOM: 1px">民俗预测</TD>
      </tr>
      <tr>
         <td width="13%" valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/hmjx/"><img src="../images/c4.gif" alt="手机吉凶预测" width="80" height="80" border="0"></a></TD>
         <td width="5%" valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/hmjx/" class="red" target="_blank">号码吉凶</a></TD>
         <td width="82%" valign="top" class="ttd" style="PADDING-BOTTOM: 1px">有些人会问，号码数字为什么会影响一个人的运势？其实这就像风水、阳宅会影响运势命运的意义是一样的。虽然这只是一个号码，但是它与您的生活息息相关，也代表您与所有人的沟通桥梁！<A href="/hmjx/" target="_blank" class="red">点击进入>></A></TD>
      </tr>
      <tr>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px"><a href="/ytyc/"><img src="../images/c5.gif" alt="眼跳预测 " width="80" height="80" border="0"></a></TD>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px"><a href="/ytyc/" class="red" target="_blank">眼跳预测 </a></TD>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px">眼皮跳跳，是祸是福？是灾是喜？想知道是情人想你，呃，还是老妈在念叨？又或者是老板要找你“谈话”？！别急别急，有关于眼跳吉凶、眼跳释意、眼跳征兆等的一切，眼跳释义都会告诉你哦。<A href="/ytyc/" target="_blank" class="red">点击进入>></A></TD>
      </tr>

      <tr>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px"><a href="/mryc/"><img src="../images/c6.gif" alt="面热预测" width="80" height="80" border="0"></a></TD>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px"><a href="/mryc/" class="red" target="_blank">面热预测</a></TD>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px">脸如火在烧？热的不得了？不是感冒，又没发烧，为何脸会红？不能不知道！面热预测，告诉你面热征兆,面热吉凶,面热的原因，快来算哦！<A href="/mryc/" target="_blank" class="red">点击进入>></A></TD>
      </tr>
      <tr>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px"><a href="/ptyc/"><img src="../images/c7.gif" alt="喷嚏预测" width="80" height="80" border="0"></a></TD>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px"><a href="/ptyc/" class="red" target="_blank">喷嚏预测</a></TD>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px">真的是一想,二骂,三念叨吗？这可是受到当时磁场的影响。喷嚏预测，告诉你打喷嚏的预兆，来看个究竟吧！<A href="/ptyc/" target="_blank" class="red">点击进入>></A></TD>
      </tr>
      <tr>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px"><a href="/xjyc/"><img src="../images/c8.gif" alt="心惊预测" width="80" height="80" border="0"></a></TD>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px"><a href="/xjyc/" class="red" target="_blank">心惊预测</a></TD>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px">总是会莫名的心惊，为了它，你总烦躁不安，告诉自己要镇定下来。到底心惊预测着什么，心惊的原因，心惊的征兆是什么，快来看看心惊释义怎么说吧！<A href="/xjyc/" target="_blank" class="red">点击进入>></A></TD>
      </tr>
      <tr>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px"><span class="ttd" style="PADDING-BOTTOM: 1px"><a href="/emyc/"><img src="../images/c9.jpg" alt="耳鸣预测" width="80" height="80" border="0"></a></span></TD>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px"><a href="/emyc/" class="red" target="_blank">耳鸣预测</a></TD>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px">耳朵里好象有“小蜜蜂”在“嗡嗡嗡”，或者时刻像在听“摇滚乐”？你这是耳鸣现象啦！为什么会耳鸣呢？当然有科学的解释，也有一些是科学所不能告诉你的部分哦。通过耳鸣占卜，看看这种耳鸣预测出你将来会发生什么事情哦。<a href="/emyc/" target="_blank" class="red">点击进入>></A></TD>
      </tr>	  
	  <tr>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px"><a href="/sanshishu/"><img src="../images/sss.gif" alt="财运预测" width="80" height="80" border="0"></a></TD>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px"><a href="/sanshishu/" target="_blank"  class="red">财运预测</a></TD>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px">三世书是传说中一本可以看到你前世、今生和后世的书，里面所有人的一切都包括，这个财运篇是专门从里面精选而来，可以算到你的财运属于哪种类型的，据说很准的哦，试试吧！<a href="/sanshishu/" target="_blank" class="red">点击进入>></a></TD>
      </tr>
      <tr>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px"><a href="/snsn/"><img src="../images/c1.gif" alt="生男生女" width="80" height="80" border="0"></a></TD>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px"><a  class="red" href="/snsn/" target="_blank">生男生女</a></TD>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px">想预测一下你们会生个男宝宝或是女宝宝吗？传统命理精髓，神奇预测生男生女结果。以下是一份清宫珍藏的生男生女预测表，女性年龄是以农历虚龄(即真实年龄加1岁)算,，而怀孕月份即是受孕的那个月份，当然亦是以农历算。<a href="/snsn/" target="_blank" class="red">点击进入>></a></TD>
      </tr>
      <tr>
        <TD vAlign=top class=new style="PADDING-BOTTOM: 1px"><a href="/zwsm/"><img src="../images/c3.gif" alt="指纹算命" width="80" height="80" border="0"></a></TD>
        <TD vAlign=top class=new style="PADDING-BOTTOM: 1px"><a href="/zwsm/" class="red" target="_blank">指纹算命</a></TD>
        <TD vAlign=top class=new style="PADDING-BOTTOM: 1px">据多年研究报告，人的性格是与生具来的。令人奇怪的是，人的指纹也是终生不变的。经过多年的研究，终于通过指纹可以辅助研究人的性格。<A href="/zwsm/" target="_blank" class="red">点击进入>></A></TD>
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
