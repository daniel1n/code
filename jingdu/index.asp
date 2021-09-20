<% 
const senlontitle="地区经度查询"
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>地区经度查询_实用查询工具大全 - Powered by Senlon!</TITLE>

</head>
<body><!--#include file="../senlon/getuc.asp"-->
<SCRIPT src="comefrom.js"></SCRIPT><body topmargin=50 leftmargin=0 onload=init()>
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
<%
'其他数据

province1=trim(request("province1"))
city1=trim(request("city1"))

if not isnumeric(jingdu) then jingdu=0
on error resume next
	db1="senlon.asp"
	Set conn1 = Server.CreateObject("ADODB.Connection")
	connstr1="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db1)
	conn1.Open connstr1

if province1<>"" and city1<>"" then
	set rs=server.createobject("ADODB.recordset")
	strsql="select * from jd where sheng='"&province1&"' and xian='"&city1&"'"
	rs.open strsql,conn1,1,1 
	if not rs.eof then
		id=rs("id")
		jingdu=rs("jingdu")
		jingdu=Split(jingdu,".")

	else
	jingdu="没有找到相关地区"
	end if
else
end if
conn1.close
set conn1=nothing

%>
  <table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" >
<tbody><form method="POST"  action="" name=form1 onSubmit="return submitchecken();">
<input type="hidden" name="act" value="ok" /><tr>
  <td width="100%" class="ttop"> 地区经度查询</td>
    </tr>  <tr>
<td class="new">省份名称：
  <SELECT name=province onchange=select()></SELECT> <INPUT name=province1 style="color:ff0000;" size=6 value="<%=province1%>" >
  县市名称：
  <SELECT name=city onchange=select()></SELECT> <INPUT name=city1 style="color:ff0000;" size=8 value="<%=city1%>" >&nbsp;</td>
</tr>     <%if request("act")="ok" then%>
 <tr>
        <td class="new"><p>&nbsp;<br>
所查地区：<font color="#FF0000"><%=province1%>-<%=city1%></font><p>
地区经度：<font color="#0000FF"><%=jingdu(0)%>度<%=jingdu(1)%>分</font><p>

　　  </td>
      </tr><%end if%>
<tr>
  <td class="new"><input type="submit" name="Submit3" value="查询地区经度" style='cursor:hand;'></td>
</tr>

</tr>
</form>
</table>

  <table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" >
    <tbody>
   
      <tr>
        <td width="100%" class="ttop">地区经度查询</td>
      </tr>
      <tr>
        <td class="new"><p>说明：本系统目前不支持中国以外的地区经度查询，支持中国县级以上城市的经度查询，城市名称结尾请不要带“市”、“县”等字，尽量写短一些，本系统数据支持模糊方式查询。（县市名称可以为空） 
 </p>
          <p>经度泛指球面坐标系的纵坐标。定义为地球面上一点与两极的连线与0度经线所在平面的夹角。以球面上的点所在辅圈相对于坐标原点所在辅圈的角距离来表示。通常特指地理坐标的经度。为了区分地球上的每一个地区，人们给经线标注了度数，这就是经度（longitude ）。实际上经度是两条经线所在平面之间的夹角。<br><br>从北极点到南极点，可以画出许多南北方向的与地球赤道垂直的大圆圈，这叫作“经圈”；构成这些圆圈的线段，就叫经线。公元1884年，国际上规定以通过英国伦敦近郊的格林尼治天文台旧址的经线作为计算经度的起点，即经度零度零分零秒，也称“本初子午线”。在它东面的为东经，共180度；在它西面的为西经，共180度。因为地球是圆的，所以东经180度和西经180度的经线是同一条经线。各国公定180度经线为“国际日期变更线”。为了避免同一地区使用两个不同的日期，国际日期变线在遇陆地时略有偏离。<br><br>国际上规定，把通过英国首都伦敦格林尼治天文台原址的那一条经线定为0°经线，也叫本初子午线。从0°经线算起，向东、向西各分作180°，以东的180°属于东经，习惯上用“E”作代号，以西的180°属于西经，习惯上用“W”作代号。东经180°和西经的180°重合在一条经线上，那就是180°经线。在地图上判读经度时应注意：从西向东，经度的度数由小到大为东经度；从西向东，经度的度数由大到小，为西经度；除0°和180°经线外，其余经线都能准确区分是东经度还是西经度。不同的经线具有不同的地方时。偏东的地方时要早，偏西的地方时要迟。每15个经度便相差一个小时。</p></td>
      </tr>
  </table>
      <!--#INCLUDE FILE="../dq.asp"-->
	</div>
    </td>
    <td valign="top"><!--#INCLUDE FILE="../right.asp"--></td>
  </tr>
</table>
<!--#include file="../foot.asp"-->
</body></html>
