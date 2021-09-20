<% 
const senlontitle="星座运程"
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>星座运程查询,十二生肖解读,十二星座解读_实用查询工具大全 - Powered by Senlon!</TITLE>

</head>
<body>
<!--#include file="../senlon/conn.asp"-->
<!--#include file="../senlon/function.asp"-->
<!--#include file="../senlon/cookies.asp"-->
<%
writecookies()
%>
<!--#include file="../senlon/getuc.asp"-->
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <TBODY>
      <tr>
        <TD height='35' align="left" class=new style="PADDING-BOTTOM: 1px"><%
if xing<>"" then
myxz=Constellation(nian1&"-"&yue1&"-"&ri1)
%><%=xing&ming%>(<%=nian1%>-<%=yue1%>-<%=ri1%>) 我的星座：<font color="#FF0000"><%=myxz%></font><a href="../sm/astro.asp?flag=4&astro=<%=myxz%>">[详情]</a> 我的属相:<font color="#FF0000"><%=sx%></font><a href="../sm/astro.asp?flag=5&xiao=<%=sx%>">[性格]</a> 我的血型:<%=xuexing%><a href="../sm/astro.asp?flag=6&xuexing=<%=xuexing%>">[详解]</a> &nbsp;<input type="button" style="color:#FF0000" value="姓名测算" onClick="(location='../sm/')" >&nbsp;<input type="button" style="color:#FF0000" value="退出重查" onClick="(location='../sm/cookieclear.asp')" ><%else%>
    <form action="./" method="post"><select name="y" size="1" id="y" style="font-size: 9pt"> 
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
日(<a href="/wannianli/" target="_blank"><font color="#FF0000">阴历/阳历转换</font></a>)
<input name="Input2" type="submit" value="查询我的星座" class="bot01"   /><input name="act" type="hidden" value="xzcx"><%if request("act")="xzcx" then
%>&nbsp;&nbsp;&nbsp;您的星座是：<font color="#FF0000"><%=Constellation(""&request("y")&""&"-"&""&request("m")&""&"-"&""&request("d")&"")%></font></form>
<%end if%><%end if%></form></TD>
      </tr>
    </TBODY>
  </TABLE>
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <tbody>
      <tr>
        <td colspan="2" class="ttop">血型/星座解读:</td>
      </tr>
      <tr>
        <td width="49%" class="new"><form action=xueread.asp method=post target=_blank>
<input type=hidden name=flag value=6>
<select name="xuexing" style="width:120px;">
	<option value=O>O血型详解</option>
		<option value=A>A血型详解</option>
		<option value=B>B血型详解</option>
		<option value=AB>AB血型详解</option>
	</select>

 
	<input name="" type="submit" value="查看" class="bot05" /></p>
</form></td>
        <td width="51%" class="new"><form action=senread.asp method=post target=_blank>
<input type=hidden name=flag value=7>
	<select name="astro">
<option>白羊座</option>		<option>金牛座</option>		<option>双子座</option>

<option>巨蟹座</option>		<option>狮子座</option>		<option>处女座</option>
<option>天秤座</option>		<option>天蝎座</option>		<option>射手座</option>
<option>魔羯座</option>		<option>水瓶座</option>		<option>双鱼座</option>
	</select>
 
	<select name="blood" class="sele01">
<option value="O">O型</option><option value="A">A型</option><option value="B">B型</option><option value="AB">AB型</option>
	</select>
	<input name="" type="submit" value="星座+血型"></p>
	</form></td>
      </tr>
	  <tr>
        <td colspan="2" class="new"><a href="xueread.asp?flag=6&xuexing=o" target="_blank"><b>O型血</b></a>：<a href="senread.asp?flag=7&blood=o&astro=白羊座" target="_blank">白羊座</a>&nbsp;<a href="senread.asp?flag=7&blood=o&astro=金牛座" target="_blank">金牛座</a>&nbsp;<a href="senread.asp?flag=7&blood=o&astro=双子座" target="_blank">双子座</a>&nbsp;<a href="senread.asp?flag=7&blood=o&astro=巨蟹座" target="_blank">巨蟹座</a>&nbsp;<a href="senread.asp?flag=7&blood=o&astro=狮子座" target="_blank">狮子座</a>&nbsp;<a href="senread.asp?flag=7&blood=o&astro=处女座" target="_blank">处女座</a>&nbsp;<a href="senread.asp?flag=7&blood=o&astro=天秤座" target="_blank">天秤座</a>&nbsp;<a href="senread.asp?flag=7&blood=o&astro=天蝎座" target="_blank">天蝎座</a>&nbsp;<a href="senread.asp?flag=7&blood=o&astro=射手座" target="_blank">射手座</a>&nbsp;<a href="senread.asp?flag=7&blood=o&astro=魔羯座" target="_blank">魔羯座</a>&nbsp;<a href="senread.asp?flag=7&blood=o&astro=水瓶座" target="_blank">水瓶座</a>&nbsp;<a href="senread.asp?flag=7&blood=o&astro=双鱼座" target="_blank">双鱼座</a></td>
      </tr>
	  <tr>
        <td colspan="2" class="new"><a href="xueread.asp?flag=6&xuexing=a" target="_blank"><b>A型血</b></a>：<a href="senread.asp?flag=7&blood=a&astro=白羊座" target="_blank">白羊座</a>&nbsp;<a href="senread.asp?flag=7&blood=a&astro=金牛座" target="_blank">金牛座</a>&nbsp;<a href="senread.asp?flag=7&blood=a&astro=双子座" target="_blank">双子座</a>&nbsp;<a href="senread.asp?flag=7&blood=a&astro=巨蟹座" target="_blank">巨蟹座</a>&nbsp;<a href="senread.asp?flag=7&blood=a&astro=狮子座" target="_blank">狮子座</a>&nbsp;<a href="senread.asp?flag=7&blood=a&astro=处女座" target="_blank">处女座</a>&nbsp;<a href="senread.asp?flag=7&blood=a&astro=天秤座" target="_blank">天秤座</a>&nbsp;<a href="senread.asp?flag=7&blood=a&astro=天蝎座" target="_blank">天蝎座</a>&nbsp;<a href="senread.asp?flag=7&blood=a&astro=射手座" target="_blank">射手座</a>&nbsp;<a href="senread.asp?flag=7&blood=a&astro=魔羯座" target="_blank">魔羯座</a>&nbsp;<a href="senread.asp?flag=7&blood=a&astro=水瓶座" target="_blank">水瓶座</a>&nbsp;<a href="senread.asp?flag=7&blood=a&astro=双鱼座" target="_blank">双鱼座</a></td>
      </tr>
	  <tr>
        <td colspan="2" class="new"><a href="xueread.asp?flag=6&xuexing=b" target="_blank"><b>B型血</b></a>：<a href="senread.asp?flag=7&blood=b&astro=白羊座" target="_blank">白羊座</a>&nbsp;<a href="senread.asp?flag=7&blood=b&astro=金牛座" target="_blank">金牛座</a>&nbsp;<a href="senread.asp?flag=7&blood=b&astro=双子座" target="_blank">双子座</a>&nbsp;<a href="senread.asp?flag=7&blood=b&astro=巨蟹座" target="_blank">巨蟹座</a>&nbsp;<a href="senread.asp?flag=7&blood=b&astro=狮子座" target="_blank">狮子座</a>&nbsp;<a href="senread.asp?flag=7&blood=b&astro=处女座" target="_blank">处女座</a>&nbsp;<a href="senread.asp?flag=7&blood=b&astro=天秤座" target="_blank">天秤座</a>&nbsp;<a href="senread.asp?flag=7&blood=b&astro=天蝎座" target="_blank">天蝎座</a>&nbsp;<a href="senread.asp?flag=7&blood=b&astro=射手座" target="_blank">射手座</a>&nbsp;<a href="senread.asp?flag=7&blood=b&astro=魔羯座" target="_blank">魔羯座</a>&nbsp;<a href="senread.asp?flag=7&blood=b&astro=水瓶座" target="_blank">水瓶座</a>&nbsp;<a href="senread.asp?flag=7&blood=b&astro=双鱼座" target="_blank">双鱼座</a></td>
      </tr>
	  <tr>
        <td colspan="2" class="new"><a href="xueread.asp?flag=6&xuexing=ab" target="_blank"><b>AB型血</b></a>：<a href="senread.asp?flag=7&blood=ab&astro=白羊座" target="_blank">白羊座</a>&nbsp;<a href="senread.asp?flag=7&blood=ab&astro=金牛座" target="_blank">金牛座</a>&nbsp;<a href="senread.asp?flag=7&blood=ab&astro=双子座" target="_blank">双子座</a>&nbsp;<a href="senread.asp?flag=7&blood=ab&astro=巨蟹座" target="_blank">巨蟹座</a>&nbsp;<a href="senread.asp?flag=7&blood=ab&astro=狮子座" target="_blank">狮子座</a>&nbsp;<a href="senread.asp?flag=7&blood=ab&astro=处女座" target="_blank">处女座</a>&nbsp;<a href="senread.asp?flag=7&blood=ab&astro=天秤座" target="_blank">天秤座</a>&nbsp;<a href="senread.asp?flag=7&blood=ab&astro=天蝎座" target="_blank">天蝎座</a>&nbsp;<a href="senread.asp?flag=7&blood=ab&astro=射手座" target="_blank">射手座</a>&nbsp;<a href="senread.asp?flag=7&blood=ab&astro=魔羯座" target="_blank">魔羯座</a>&nbsp;<a href="senread.asp?flag=7&blood=ab&astro=水瓶座" target="_blank">水瓶座</a>&nbsp;<a href="senread.asp?flag=7&blood=ab&astro=双鱼座" target="_blank">双鱼座</a></td>
      </tr>    
    </tbody>
  </table>
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <tbody>
      <tr>
	  <tr>
        <td colspan="2" class="ttop">生日书/生日花:</td>
      </tr>
      <tr>
        <td class="new"><form action=../shengrishu/ method=post target=_blank>
	<input type=hidden name=flag value=8>
	<select name="m" class="sele02" id="m">
	  <option value="1" selected>月</option>
	  <option value="1">1月</option>
	  <option value="2">2月</option>
	  <option value="3">3月</option>
	  <option value="4">04</option>
	  <option value="5">05</option>
	  <option value="6">06</option>
	  <option value="7">07</option>
	  <option value="8">08</option>
	  <option value="9">09</option>

    <option value="10">10</option>
    <option value="11">11</option>
    <option value="12">12</option>
	</select>
	<select name="d" class="sele02" id="d">
	  <option value="1" selected>日</option>
	  <option value="1">01</option>
	  <option value="2">02</option>
	  <option value="3">03</option>
	  <option value="4">04</option>
	  <option value="5">05</option>
	  <option value="6">06</option>
	  <option value="7">07</option>
	  <option value="8">08</option>
	  <option value="9">09</option>
	  <option value="10">10</option>
	  <option value="11">11</option>
	  <option value="12">12</option>
	  <option value="13">13</option>
	  <option value="14">14</option>
	  <option value="15">15</option>
	  <option value="16">16</option>
	  <option value="17">17</option>
	  <option value="18">18</option>
	  <option value="19">19</option>
	  <option value="20">20</option>
	  <option value="21">21</option>
	  <option value="22">22</option>
	  <option value="23">23</option>
	  <option value="24">24</option>
	  <option value="25">25</option>
	  <option value="26">26</option>
	  <option value="27">27</option>
	  <option value="28">28</option>
	  <option value="29">29</option>
	  <option value="30">30</option>
	  <option value="31">31</option>
	</select>
	 
 
	<select name=flag1 class="sele01" id="flag1">
	  <option value="生日书">生日书</option>
	</select>
 
	<input name="" type="submit" value="查询" class="bot01"   /></p>
 </form></td>
        <td class="new"><form action=../shengrihua/ method=post target=_blank>
	<input type=hidden name=flag value=8>
	<select name="m" class="sele02" id="m">
	  <option value="1" selected>月</option>
	  <option value="1">1月</option>
	  <option value="2">2月</option>
	  <option value="3">3月</option>
	  <option value="4">04</option>
	  <option value="5">05</option>
	  <option value="6">06</option>
	  <option value="7">07</option>
	  <option value="8">08</option>
	  <option value="9">09</option>

    <option value="10">10</option>
    <option value="11">11</option>
    <option value="12">12</option>
	</select>
	<select name="d" class="sele02" id="d">
	  <option value="1" selected>日</option>
	  <option value="1">01</option>
	  <option value="2">02</option>
	  <option value="3">03</option>
	  <option value="4">04</option>
	  <option value="5">05</option>
	  <option value="6">06</option>
	  <option value="7">07</option>
	  <option value="8">08</option>
	  <option value="9">09</option>
	  <option value="10">10</option>
	  <option value="11">11</option>
	  <option value="12">12</option>
	  <option value="13">13</option>
	  <option value="14">14</option>
	  <option value="15">15</option>
	  <option value="16">16</option>
	  <option value="17">17</option>
	  <option value="18">18</option>
	  <option value="19">19</option>
	  <option value="20">20</option>
	  <option value="21">21</option>
	  <option value="22">22</option>
	  <option value="23">23</option>
	  <option value="24">24</option>
	  <option value="25">25</option>
	  <option value="26">26</option>
	  <option value="27">27</option>
	  <option value="28">28</option>
	  <option value="29">29</option>
	  <option value="30">30</option>
	  <option value="31">31</option>
	</select>
	 
 
	<select name=flag1 class="sele01" id="flag1">
	  <option value="生日花">生日花</option>
	</select>
 
	<input name="" type="submit" value="查询" class="bot01"   /></p>
 </form></td>
      </tr> 
      </tr>    
    </tbody>
  </table>
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <tbody>
      <tr>
        <td width="100%" colspan="6" class="ttop">十二星座解读</td>
      </tr>
      <tr>
        <td colspan="6" class="new2"><span class="red">·你知道自己的星座吗？你知道不同星座的性格，爱情，事业如何吗？快来点击分析一下自己的星座吧！</span></td>
      </tr>
      <tr>
        <td class="new"><a href="12xz.asp?flag=4&astro=白羊座" target="_blank"><img src="../images/xz_1.gif" alt="白羊座" width="90" height="80" border="0"></a></td>
        <td class="new"><a href="12xz.asp?flag=4&astro=金牛座" target="_blank"><img src="../images/xz_2.gif" alt="金牛座" width="90" height="80" border="0"></a></td>
        <td class="new"><a href="12xz.asp?flag=4&astro=双子座" target="_blank"><img src="../images/xz_3.gif" alt="双子座" width="90" height="80" border="0"></a></td>
        <td class="new"><a href="12xz.asp?flag=4&astro=巨蟹座" target="_blank"><img src="../images/xz_4.gif" alt="巨蟹座" width="90" height="80" border="0"></a></td>
        <td class="new"><a href="12xz.asp?flag=4&astro=狮子座" target="_blank"><img src="../images/xz_5.gif" alt="狮子座" width="90" height="80" border="0"></a></td>
        <td class="new"><a href="12xz.asp?flag=4&astro=处女座" target="_blank"><img src="../images/xz_6.gif" alt="处女座" width="90" height="80" border="0"></a></td>
      </tr>
      <tr>
        <td class="new"><DIV align="center"><STRONG>白羊座</STRONG></DIV></td>
        <td class="new"><DIV align="center"><STRONG>金牛座</STRONG></DIV></td>
        <td class="new"><DIV align="center"><STRONG>双子座</STRONG></DIV></td>
        <td class="new"><DIV align="center"><STRONG>巨蟹座</STRONG></DIV></td>
        <td class="new"><DIV align="center"><STRONG>狮子座</STRONG></DIV></td>
        <td class="new"><DIV align="center"><STRONG>处女座</STRONG></DIV></td>
      </tr>
      <tr>
        <td class="new"><a href="12xz.asp?flag=4&astro=天秤座" target="_blank"><img src="../images/xz_7.gif" alt="天秤座" width="90" height="80" border="0"></a></td>
        <td class="new"><a href="12xz.asp?flag=4&astro=天蝎座" target="_blank"><img src="../images/xz_8.gif" alt="天蝎座" width="90" height="80" border="0"></a></td>
        <td class="new"><a href="12xz.asp?flag=4&astro=射手座" target="_blank"><img src="../images/xz_9.gif" alt="射手座" width="90" height="80" border="0"></a></td>
        <td class="new"><a href="12xz.asp?flag=4&astro=魔羯座" target="_blank"><img src="../images/xz_10.gif" alt="魔羯座" width="90" height="80" border="0"></a></td>
        <td class="new"><a href="12xz.asp?flag=4&astro=水瓶座" target="_blank"><img src="../images/xz_11.gif" alt="水瓶座" width="90" height="80" border="0"></a></td>
        <td class="new"><a href="12xz.asp?flag=4&astro=双鱼座" target="_blank"><img src="../images/xz_12.gif" alt="双鱼座" width="90" height="80" border="0"></a></td>
      </tr>
      <tr>
        <td class="new"><DIV align="center"><STRONG>天秤座</STRONG></DIV></td>
        <td class="new"><DIV align="center"><STRONG>天蝎座</STRONG></DIV></td>
        <td class="new"><DIV align="center"><STRONG>射手座</STRONG></DIV></td>
        <td class="new"><DIV align="center"><STRONG>魔羯座</STRONG></DIV></td>
        <td class="new"><DIV align="center"><STRONG>水瓶座</STRONG></DIV></td>
        <td class="new"><DIV align="center"><STRONG>双鱼座</STRONG></DIV></td>
      </tr>
    </tbody>
  </table>

<!-- 今日星座开始 -->
<%
ostr=year(now)&month(now)&day(now)
%>
<table width="100%" border="0" align="center" align="center" cellpadding="0" cellspacing="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <tr>
      <td colspan="6"  valign="top" class="ttop" style="PADDING-BOTTOM: 1px">今日星座运势</td>
    </tr>
	<tr>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=1&day=<%=ostr%>"><img border="0" src="../xingzuo/images/1.gif" width="78" height="81"></a></td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=2&day=<%=ostr%>"><img border="0" src="../xingzuo/images/2.gif" width="81" height="81"></a></td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=3&day=<%=ostr%>"><img border="0" src="../xingzuo/images/3.gif" width="78" height="79"></a></td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=4&day=<%=ostr%>"><img border="0" src="../xingzuo/images/4.gif" width="80" height="79"></a></td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=5&day=<%=ostr%>"><img border="0" src="../xingzuo/images/5.gif" width="79" height="77"></a></td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=6&day=<%=ostr%>"><img border="0" src="../xingzuo/images/6.gif" width="79" height="79"></a></td>
	</tr>
	<tr>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=1&day=<%=ostr%>">白羊座</a><br>
		(3.21-4.20)</td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=2&day=<%=ostr%>">金牛座</a><br>
		(4.21-5.20)</td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=3&day=<%=ostr%>">双子座</a><br>
		(5.21-6.21)</td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=4&day=<%=ostr%>">巨蟹座</a><br>
		(6.22-7.22)</td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=5&day=<%=ostr%>">狮子座</a><br>
		(7.23-8.22)</td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=6&day=<%=ostr%>">处女座</a><br>
		(8.23-9.22)</td>
	</tr>
	<tr>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=7&day=<%=ostr%>"><img border="0" src="../xingzuo/images/7.gif" width="80" height="82"></a></td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=8&day=<%=ostr%>"><img border="0" src="../xingzuo/images/8.gif" width="80" height="78"></a></td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=9&day=<%=ostr%>"><img border="0" src="../xingzuo/images/9.gif" width="74" height="74"></a></td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=10&day=<%=ostr%>"><img border="0" src="../xingzuo/images/10.gif" width="79" height="82"></a></td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=11&day=<%=ostr%>"><img border="0" src="../xingzuo/images/11.gif" width="80" height="81"></a></td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=12&day=<%=ostr%>"><img border="0" src="../xingzuo/images/12.gif" width="78" height="77"></a></td>
	</tr>
	<tr>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=7&day=<%=ostr%>">天秤座</a><br>
		(9.23-10.22)</td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=8&day=<%=ostr%>">天蝎座</a><br>
		(10.23-11.21)</td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=9&day=<%=ostr%>">射手座</a><br>
		(11.22-12.21)</td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=10&day=<%=ostr%>">摩羯座</a><br>
		(12.22-1.19)</td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=11&day=<%=ostr%>">水瓶座</a><br>
		(1.20-2.19)</td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=12&day=<%=ostr%>">双鱼座</a><br>
		(2.20-3.20)</td>
	</tr>
</table>
<!-- 今日星座结束 -->
      <!--#INCLUDE FILE="../dq.asp"-->
	</div>
    </td>
    <td valign="top"><!--#INCLUDE FILE="../right.asp"--></td>
  </tr>
</table>
<!--#include file="../foot.asp"-->
</body></html>