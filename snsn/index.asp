<% 
const senlontitle="生男生女预测"
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>生男生女预测,清宫生男生女预测表_实用查询工具大全 - Powered by Senlon!</TITLE>
<!--#include file="../senlon/getuc.asp"-->

</head>
<body><!--#include file="../senlon/function.asp"--><!--#include file="../senlon/conn.asp"-->  
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
  <table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
<tbody><tr>
  <td width="100%" class="ttd"><span class="red">生男生女:</span>
    <br><br>本生男生女测算数据来自《清宫珍藏的生男生女预测表》，想预测一下你们会生个男宝宝或是女宝宝吗？那么就开始算算吧！结果仅供参考哦，请勿太信！<br><br></td>
</tr><%if request("act")="ok" then
mqname=request("mqname")
if mqname="" then
mqname="亲爱的511cha网友"
end if
cs=request("cs")
yue=request("yue")
nn=request("nn")
snsn=request("snsn")%>
  <tr>
    <td colspan="2" class="new"><font color="#0000FF"><%=mqname%></font>，您好<br><%if request("cs")=1 then
sql="select * from snsn where nn='"&nn&"' "
set rs=conn.execute(sql)
if not (rs.bof and rs.eof) then
baby=rs(""&yue&"")
end if
if baby="男" then
baby="小王子"
elseif baby="女" then
baby="小公主"
end if
%>恭喜您，根据推算，你很可能会有一个&nbsp;<font color="#FF0000"><%=baby%></font><%else%><%
sql="select * from snsn where nn='"&nn&"' "
set rs=conn.execute(sql)
if rs("1")=snsn then
yuef=yuef&"一月&nbsp;"
end if
if rs("2")=snsn then
yuef=yuef&"二月&nbsp;"
end if
if rs("3")=snsn then
yuef=yuef&"三月&nbsp;"
end if
if rs("4")=snsn then
yuef=yuef&"四月&nbsp;"
end if
if rs("5")=snsn then
yuef=yuef&"五月&nbsp;"
end if
if rs("6")=snsn then
yuef=yuef&"六月&nbsp;"
end if
if rs("7")=snsn then
yuef=yuef&"七月&nbsp;"
end if
if rs("8")=snsn then
yuef=yuef&"八月&nbsp;"
end if
if rs("9")=snsn then
yuef=yuef&"九月&nbsp;"
end if
if rs("10")=snsn then
yuef=yuef&"十月&nbsp;"
end if
if rs("11")=snsn then
yuef=yuef&"十一月&nbsp;"
end if
if rs("12")=snsn then
yuef=yuef&"十二月&nbsp;"
end if
sql="select * from snsn where nn='"&nn+1&"' "
set rs=conn.execute(sql)
if rs("1")=snsn then
nyuef=nyuef&"一月&nbsp;"
end if
if rs("2")=snsn then
nyuef=nyuef&"二月&nbsp;"
end if
if rs("3")=snsn then
nyuef=nyuef&"三月&nbsp;"
end if
if rs("4")=snsn then
nyuef=nyuef&"四月&nbsp;"
end if
if rs("5")=snsn then
nyuef=nyuef&"五月&nbsp;"
end if
if rs("6")=snsn then
nyuef=nyuef&"六月&nbsp;"
end if
if rs("7")=snsn then
nyuef=nyuef&"七月&nbsp;"
end if
if rs("8")=snsn then
nyuef=nyuef&"八月&nbsp;"
end if
if rs("9")=snsn then
nyuef=nyuef&"九月&nbsp;"
end if
if rs("10")=snsn then
nyuef=nyuef&"十月&nbsp;"
end if
if rs("11")=snsn then
nyuef=nyuef&"十一月&nbsp;"
end if
if rs("12")=snsn then
nyuef=nyuef&"十二月&nbsp;"
end if
if snsn="男" then
baby="小王子"
elseif snsn="女" then
baby="小公主"
end if
%>您如果想生个<font color="#FF0000"><%=baby%></font>，那么建议您在农历 <font color="#0000FF">今年 ：<%=yuef%>&nbsp;&nbsp;明年：<%=nyuef%></font> 怀孕的话机会比较大！<%end if%><br><br>
      <a class="red" href="./">重新测试</a>  </td>
<%else%>
  <tr>
    <td colspan="2" class="ttd"><font color="#FF0000">*年龄为母亲虚岁年龄，月份指怀孕月份，以农历为准。遇闰月，上半月以上个月份计算，下半月以下个月份计算。</font>
   </td>
    </tr>  <tr>
    <td colspan="2" class="ttd"><font color="#0000ff">*我要预测宝宝性别：</font>
   </td>
    </tr><form name="theform" method="post" action="">
<input type="hidden" name="act" value="ok" />
<input type="hidden" name="cs" value="1" /><tr>
    <td colspan="2" class="ttd">我的芳名叫：<input name="mqname" type="text">&nbsp;今年芳龄(虚岁)是:<select name="nn" size="1" style="font-size: 9pt">
      ><%for i=18 to 45%><option value="<%=i%>" <%if i=22 then%> selected<%end if%>><%=i%></option><%next%>
     </select>&nbsp;怀孕的月份(<a href="/wannianli/" target="_blank"><font color=red>农历</font></a>)是:<select name="yue" size="1" style="font-size: 9pt">
      ><%for i=1 to 12%><option value="<%=i%>" ><%=i%></option><%next%>
     </select>月 &nbsp;<input name="sub" type="submit" value="开始推算">
   </td></form>
    </tr> <tr>
    <td colspan="2" class="ttd"><font color="#0000ff">*我要查询适合怀孕的月份：</font>
   </td>
    </tr><form name="theform" method="post" action="">
<input type="hidden" name="act" value="ok" />
<input type="hidden" name="cs" value="2" /><tr>
    <td colspan="2" class="new">我的芳名叫：<input name="mqname" type="text">&nbsp;今年芳龄(虚岁)是:<select name="nn" size="1" style="font-size: 9pt">
      ><%for i=18 to 45%><option value="<%=i%>" <%if i=22 then%> selected<%end if%>><%=i%></option><%next%>
     </select>&nbsp;我计划生个:<select name="snsn" size="1" style="font-size: 9pt">
       <option value="男">小王子</option>
       <option value="女">小公主</option>

     </select>
     &nbsp;
     <input name="sub" type="submit" value="开始推算">
   </td></form>
    </tr><%end if%>
</tbody>
</table>
      <!--#INCLUDE FILE="../dq.asp"-->
	</div>
    </td>
    <td valign="top"><!--#INCLUDE FILE="../right.asp"--></td>
  </tr>
</table>
<!--#include file="../foot.asp"-->
</center>
</body></html>