 
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#5391EE"  style="MARGIN-BOTTOM: 10px; table-layout:fixed;word-wrap:break-word;">    <tbody>
<tr>
<td  width="6%" rowspan="6" bgcolor="#FFFFFF" class="new"><%xing1=mid(xing,1,1)
ming1=mid(ming,1,1)
%>
<b><%=xing1%></b><br /><br />
<%if mid(xing,2,1) <> "" then
response.Write "<b>"&mid(xing,2,1)&"</b> <br /><br />"
end if%>
<b><%=ming1%></b> <br /><br />
<%if mid(ming,2,1) <> "" then
response.Write "<b>"&mid(ming,2,1)&"</b> <br /><br />"
end if%>
<%bzwh=tgdzwh(yg1)&tgdzwh(yz1)&tgdzwh(mg1)&tgdzwh(mz1)&tgdzwh(dg1)&tgdzwh(dz1)&tgdzwh(tg1)&tgdzwh(tz1)
wnum1 = Len(Replace(bzwh, "金", Space(Len("金") + 1))) - Len(bzwh)
if wnum1>=3 then
mainw=mainw&"<strong>金</strong>旺"
end if
if wnum1=0 then
mainq=mainq&"缺<strong>金</strong>"
end if
wnum2 = Len(Replace(bzwh, "木", Space(Len("木") + 1))) - Len(bzwh)
if wnum2>=3 then
mainw=mainw&"<strong>木</strong>旺"
end if
if wnum2=0 then
mainq=mainq&"缺<strong>木</strong>"
end if
wnum3 = Len(Replace(bzwh, "水", Space(Len("水") + 1))) - Len(bzwh)
if wnum3>=3 then
mainw=mainw&"<strong>水</strong>旺"
end if
if wnum3=0 then
mainq=mainq&"缺<strong>水</strong>"
end if
wnum4 = Len(Replace(bzwh, "火", Space(Len("火") + 1))) - Len(bzwh)
if wnum4>=3 then
mainw=mainw&"<strong>火</strong>旺"
end if
if wnum4=0 then
mainq=mainq&"缺<strong>火</strong>"
end if
wnum5 = Len(Replace(bzwh, "土", Space(Len("土") + 1))) - Len(bzwh)
if wnum5>=3 then
mainw=mainw&"<strong>土</strong>旺"
end if
if wnum5=0 then
mainq=mainq&"缺<strong>土</strong>"
end if

sql="select * from wh where wh='"&tgdzwh(dg1)&"'"
set rs=conn.execute(sql)
if not (rs.bof and rs.eof) then
tnwh=rs("tnwh")
ynwh=rs("ynwh")
skzhyj=rs("skzhyj")
whzx=rs("whzx")
szwh=rs("szwh")
hyhw=rs("hyhw")
end if
rs.close

sql="select * from sjrs where wh='"&tgdzwh(dg1)&"' and sj='"&siji(yue1)&"'"
set rs=conn.execute(sql)
if not (rs.bof and rs.eof) then
sjrs=rs("sjrs")
end if
rs.close

sql="select * from sxgx where sx='"&sx&"'"
set rs=conn.execute(sql)
if not (rs.bof and rs.eof) then
sxgx=rs("sxgx")
end if
rs.close

sql="select * from rgnm where rgz='"&dgz&"'"
set rs=conn.execute(sql)
if not (rs.bof and rs.eof) then
xgfx=rs("xgfx")
aqfx=rs("aqfx")
syfx=rs("syfx")
cyfx=rs("cyfx")
jkfx=rs("jkfx")
end if
rs.close
%>
</td>
<td  width="12%" rowspan="2" bgcolor="#FFFFFF" class="new">出生时间：</td>
  <td  width="9%" bgcolor="#FFFFFF" class="new">(公历)</td>
  <td  width="12%" bgcolor="#FFFFFF" class="new"><%=nian1%>年</td>
  <td  width="10%" bgcolor="#FFFFFF" class="new"><%=yue1%>月</td>
  <td  width="10%" bgcolor="#FFFFFF" class="new"><%=ri1%>日</td>
  <td  width="11%" bgcolor="#FFFFFF" class="new"><%=hh1%>点</td>
  <td width="30%"  rowspan="6" bgcolor="#FFFFFF" class="new" style="padding-left:4px;padding-right:4px;">本命属<%=sx%>，<%=layin(ygz)%>命。五行<%=mainw%><%=mainq%>；日主天干为<b><%=tgdzwh(dg1)%></b>，生于<%=siji(yue1)%>季。<br />
    （同类<%=tnwh%>；异类<%=ynwh%>。）<br />八字五行个数 : <%=wnum1%>个金，<%=wnum2%>个木，<%=wnum3%>个水，<%=wnum4%>个火，<%=wnum5%>个土
　</td>
 </tr>
 <tr>
  <td bgcolor="#FFFFFF" class="new" >(农历)</td>
  <td bgcolor="#FFFFFF" class="new" ><%=nian2%>年</td>
  <td bgcolor="#FFFFFF" class="new" ><%=yue2%>月</td>
  <td bgcolor="#FFFFFF" class="new" ><%=ri2%>日</td>
  <td bgcolor="#FFFFFF" class="new" ><%=DiZhi(hh2)%>时</td>
 </tr>

 <tr>
  <td  colspan="2" bgcolor="#FFFFFF" class="new">八字：</td>
  <td bgcolor="#FFFFFF" class="new" ><%=ygz%>　</td>
  <td bgcolor="#FFFFFF" class="new" ><%=mgz%>　</td>
  <td bgcolor="#FFFFFF" class="new" ><%=dgz%>　</td>
  <td bgcolor="#FFFFFF" class="new" ><%=tgz%>　</td>
 </tr>
 <tr>
  <td  colspan="2" bgcolor="#FFFFFF" class="new">五行：</td>
  <td bgcolor="#FFFFFF" class="new" ><%=tgdzwh(yg1)%><%=tgdzwh(yz1)%></td>
  <td bgcolor="#FFFFFF" class="new" ><%=tgdzwh(mg1)%><%=tgdzwh(mz1)%></td>
  <td bgcolor="#FFFFFF" class="new" ><%=tgdzwh(dg1)%><%=tgdzwh(dz1)%></td>
  <td bgcolor="#FFFFFF" class="new" ><%=tgdzwh(tg1)%><%=tgdzwh(tz1)%></td>
 </tr>
 <tr>
  <td  colspan="2" bgcolor="#FFFFFF" class="new">纳音：</td>
  <td bgcolor="#FFFFFF" class="new" ><%=layin(ygz)%></td>
  <td bgcolor="#FFFFFF" class="new" ><%=layin(mgz)%></td>
  <td bgcolor="#FFFFFF" class="new" ><%=layin(dgz)%></td>
  <td bgcolor="#FFFFFF" class="new" ><%=layin(tgz)%></td>
 </tr>
</tbody>
</table>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#5391EE"  style="MARGIN-BOTTOM: 10px; table-layout:fixed;word-wrap:break-word;">
  <tbody>
    <tr>
      <td  width="16%"  bgcolor="#FFFFFF" class="new">生肖个性</td>
      <td width="84%" colspan="6" bgcolor="#FFFFFF" class="new">根据分析，您的生肖为“<%=sx%>”<br /> <%=sxgx%></td>
    </tr>
  </tbody>
</table><table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#5391EE"  style="MARGIN-BOTTOM: 10px; table-layout:fixed;word-wrap:break-word;">
  <tbody>
    <tr>
      <td  width="16%"  bgcolor="#FFFFFF" class="new">性格分析</td>
      <td width="84%" colspan="6" bgcolor="#FFFFFF" class="new"> <%=xgfx%></td>
    </tr>
  </tbody>
</table><table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#5391EE"  style="MARGIN-BOTTOM: 10px; table-layout:fixed;word-wrap:break-word;">
  <tbody>
    <tr>
      <td  width="16%"  bgcolor="#FFFFFF" class="new">爱情分析</td>
      <td width="84%" colspan="6" bgcolor="#FFFFFF" class="new"> <%=aqfx%></td>
    </tr>
  </tbody>
</table><table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#5391EE"  style="MARGIN-BOTTOM: 10px; table-layout:fixed;word-wrap:break-word;">
  <tbody>
    <tr>
      <td  width="16%"  bgcolor="#FFFFFF" class="new">事业分析</td>
      <td width="84%" colspan="6" bgcolor="#FFFFFF" class="new"> <%=syfx%></td>
    </tr>
  </tbody>
</table>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#5391EE"  style="MARGIN-BOTTOM: 10px; table-layout:fixed;word-wrap:break-word;">
  <tbody>
    <tr>
      <td  width="16%"  bgcolor="#FFFFFF" class="new">财运分析</td>
      <td width="84%" colspan="6" bgcolor="#FFFFFF" class="new"><%=cyfx%></td>
    </tr>
  </tbody>
</table>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#5391EE"  style="MARGIN-BOTTOM: 10px; table-layout:fixed;word-wrap:break-word;">
  <tbody>
    <tr>
      <td  width="16%"  bgcolor="#FFFFFF" class="new">财运分析</td>
      <td width="84%" colspan="6" bgcolor="#FFFFFF" class="new"><%=cyfx%></td>
    </tr>
  </tbody>
</table>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
  <tbody>
    <tr>
      <td class="new"><div id="ft">说明:日干论命仅对八字中日柱的信息进行分析，为片面的信息，仅供参考！</div></td>
    </tr>
  </tbody>
</table>
