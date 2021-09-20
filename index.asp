<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="senlon/getuc.asp"-->
<!--#include file="senlon/function.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>2016算命大全程序源码</title>
</head>

<body>
<!--#INCLUDE FILE="top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top"><div id="mid_left">
      <p><span class="ren">算命说明：</span><br />
      姓名必须输入中文,生日必须输入公历；如果不分析血型、血型可任选,如果不分析八字等、出生时分可任选,不影响其它测试结果.<span class="ren">－&gt;</span><A href style='cursor:hand;' onclick="window.open('http://sm.aa963.com/hdjr/index.htm','cidunongli','left=120,top=100,width=800,height=550,scrollbars=no,resizable=no,status=no')" title="如果您只知道生日的农历日期，不要紧，请点这里去查询公历日期">[公历/农历日期转换</a>] [<a href="/sm/htm_nobirth.asp" target="_blank">不知道出生时间怎么办</a>]</p>
      <hr class="shplink" />
      <form method="post" action="/sm/sm.asp?sm=1" name="sm"  onSubmit="return checkbz();"> 
        <p>姓：
        <input type="txt" name="xing" size="4" value="" onKeypress="if ((event.keyCode != 13 && event.keyCode < 160)) event.returnValue = false;">
        名：<input type="txt" name="ming" size="4" value="" onKeypress="if ((event.keyCode != 13 && event.keyCode < 160)) event.returnValue = false;">
        </p>
      <p>
        下拉选择：
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
        </p>
      <p>公历生日:
  <select name="nian" size="1" style="font-size: 9pt">
    ><%for i=1930 to year(date())%>
    <option value="<%=i%>" <%if i=year(date()) then%> selected<%end if%>><%=i%></option>
    <%next%>
  </select>
        年
        <select size="1" name="yue" style="font-size: 9pt">
          <%for i=1 to 12%>
          <option value="<%=i%>"<%if i=month(date()) then%> selected<%end if%>><%=i%></option>
          <%next%>
        </select>
        月
        <select size="1" name="ri" style="font-size: 9pt">
          <%for i=1 to 31%>
          <option value="<%=i%>" <%if i=day(date()) then%> selected<%end if%>><%=i%></option>
          <%next%>
        </select>
        日
        <select size="1" name="hh" style="font-size: 9pt">
          <%for i=0 to 23%>
          <option value="<%=i%>" ><%=DiZhi(i)%>&nbsp;<%=i%></option>
          <%next%> 
        </select>
        点
        <select size="1" name="mm" style="font-size: 9pt">
          <option value="0">未知</option>
          <%for i=0 to 59%>
          <option value="<%=i%>"><%=i%></option>
          <%next%>
        </select>
        分   　
        <input type="submit" value="开始算命" style='cursor:hand;COLOR: #ff0099;' class="anliu">
      </p>
  </form>
      <hr class="shplink" />
      <table width="100%" border="0" cellspacing="0" cellpadding="6">
        <tr>
          <td align="center"><a href="/ppbazi/index.asp"><img src="/images/pg1.gif" width="80" height="80" border="0" /></a></td>
          <td align="center"><a href="/ppziwei/index.asp"><img src="/images/pg6.gif" width="80" height="80" border="0" /></a></td>
          <td align="center"><a href="/ppqmdj/index.asp"><img src="/images/pg4.gif" width="80" height="80" border="0" /></a></td>
          <td align="center"><a href="/ppjkj/index.asp"><img src="/images/pg7.gif" width="80" height="80" border="0" /></a></td>
          <td align="center"><a href="/pp6y/index.asp"><img src="/images/pg3.gif" width="80" height="80" border="0" /></a></td>
          <td align="center"><a href="/pp6r/index.asp"><img src="/images/pg2.gif" width="80" height="80" border="0" /></a></td>
          <td align="center"><a href="/ppxk/index.asp"><img src="/images/pg5.gif" width="80" height="80" border="0" /></a></td>
          </tr>
        <tr>
          <td align="center"><a href="/ppbazi/index.asp">四柱八字</a></td>
          <td align="center"><a href="/ppziwei/index.asp">紫微斗数</a></td>
          <td align="center"><a href="/ppqmdj/index.asp">奇门遁甲</a></td>
          <td align="center"><a href="/ppjkj/index.asp">金口诀</a></td>
          <td align="center"><a href="/pp6y/index.asp">六爻起卦</a></td>
          <td align="center"><a href="/pp6r/index.asp">大六壬</a></td>
          <td align="center"><a href="/ppxk/index.asp">玄空飞星</a></td>
          </tr>
      </table>
      <p><%=(adgl.Fields.Item("addaima05").Value)%></p>
      <table width="100%" border="0" cellspacing="0" cellpadding="6">
        <tr>
          <td align="center"><a href="/sanshishu/index.asp"><img src="/images/sss.gif" width="80" height="80" border="0" /></a></td>
          <td align="center"><a href="/snsn/index.asp"><img src="/images/c1.gif" width="80" height="80" border="0" /></a></td>
          <td align="center"><a href="/shouji/index.asp"><img src="/images/c4.gif" width="80" height="80" border="0" /></a></td>
          <td align="center"><a href="/zwsm/index.asp"><img src="/images/c3.gif" width="80" height="80" border="0" /></a></td>
          <td align="center"><a href="/ytyc/index.asp"><img src="/images/c5.gif" width="80" height="80" border="0" /></a></td>
          <td align="center"><a href="/ptyc/index.asp"><img src="/images/c7.gif" width="80" height="80" border="0" /></a></td>
          <td align="center"><a href="/emyc/index.asp"><img src="/images/c9.jpg" width="80" height="80" border="0" /></a></td>
        </tr>
        <tr>
          <td align="center"><a href="/sanshishu/index.asp">预测财运</a></td>
          <td align="center"><a href="/snsn/index.asp">生男生女</a></td>
          <td align="center"><a href="/shouji/index.asp">号码吉凶</a></td>
          <td align="center"><a href="/zwsm/index.asp">指纹算命</a></td>
          <td align="center"><a href="/ytyc/index.asp">眼跳预测</a></td>
          <td align="center"><a href="/ptyc/index.asp">喷嚏预测</a></td>
          <td align="center"><a href="/emyc/index.asp">耳鸣预测</a></td>
        </tr>
    </table>
      <p><%=(adgl.Fields.Item("addaima06").Value)%></p>
    </div><!--#INCLUDE FILE="dq.asp"--></td>
    <td valign="top"><!--#INCLUDE FILE="right.asp"--></td>
  </tr>
</table>
<!--#INCLUDE FILE="foot.asp"-->
</body>
</html>
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