
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#5391EE"  style="MARGIN-BOTTOM: 10px; table-layout:fixed;word-wrap:break-word;">    
  <tbody>
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
rgxx=rs("rgxx")
rgcz=rs("rgcz")
rgzfx=rs("rgzfx")
end if
rs.close

sql="select * from rysmn where siceng='"&yue2&"月'"
set rs=conn.execute(sql)
if not (rs.bof and rs.eof) then
mingni1=rs("mingmi")
end if
rs.close
sql="select * from rysmn where siceng='"&ri2&"日'"
set rs=conn.execute(sql)
if not (rs.bof and rs.eof) then
mingni2=rs("mingmi")
end if
rs.close
sql="select * from rysmn where siceng='"&DiZhi(hh2)&"时'"
set rs=conn.execute(sql)
if not (rs.bof and rs.eof) then
mingni3=rs("mingmi")
end if
rs.close
sql="select * from smtf where ri='"&dgz&"' and shi='"&tgz&"'"
set rs=conn.execute(sql)
if not (rs.bof and rs.eof) then
tf1=rs("tf1")
tf2=rs("tf2")
end if
rs.close
sql="select * from qtbj where tg='"&dg1&"' and dz='"&mz1&"'"
set rs=conn.execute(sql)
if not (rs.bof and rs.eof) then
qtbj=rs("content")
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
    （同类<%=tnwh%>；异类<%=ynwh%>。）<br /><hr size="1"><font color="#aaaaaa" style="font-size:12px">重要说明：本结果为系统自动分析，仅供参考，八字缺什么与补什么无关，具体应由专业老师分析！</font> (<a href="./htm_smzs.asp" target="_blank">什么是八字</a>)</td>
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
</table><table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <TBODY>
<tr>
  <td class=new>八字五行个数 : <%=wnum1%>个金，<%=wnum2%>个木，<%=wnum3%>个水，<%=wnum4%>个火，<%=wnum5%>个土
</td>
</tr>
    </TBODY>
</TABLE><table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <TBODY>
<tr>
  <td class=new>四季用神参考 : 日主天干<%=tgdzwh(dg1)%>生于<%=siji(yue1)%>季,<%=sjrs%>。
</td>
</tr><tr>
  <td class=new>穷通宝鉴调候用神参考 : <%=dg1%><%=tgdzwh(dg1)%>生于<%=mz1%>月，<%=qtbj%>
</td>
</tr>
    </TBODY>
</TABLE>
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
      <td  width="16%"  bgcolor="#FFFFFF" class="new">日干心性</td>
      <td width="84%" colspan="6" bgcolor="#FFFFFF" class="new"> <%=rgxx%></td>
    </tr>
  </tbody>
</table><table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#5391EE"  style="MARGIN-BOTTOM: 10px; table-layout:fixed;word-wrap:break-word;">
  <tbody>
    <tr>
      <td  width="16%"  bgcolor="#FFFFFF" class="new">日干支层次</td>
      <td width="84%" colspan="6" bgcolor="#FFFFFF" class="new"> <%=rgcz%></td>
    </tr>
  </tbody>
</table><table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#5391EE"  style="MARGIN-BOTTOM: 10px; table-layout:fixed;word-wrap:break-word;">
  <tbody>
    <tr>
      <td  width="16%"  bgcolor="#FFFFFF" class="new">日干支分析</td>
      <td width="84%" colspan="6" bgcolor="#FFFFFF" class="new"> <%=rgzfx%><br /><font color="#cccccc">* 根据四柱预测学部分专家学者提供的资料，归纳整理，个别字句有待考证，总体准确度较高！</font></td>
    </tr>
  </tbody>
</table>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#5391EE"  style="MARGIN-BOTTOM: 10px; table-layout:fixed;word-wrap:break-word;">    <tbody>
<tr>
<td  width="16%"  bgcolor="#FFFFFF" class="new">五行生克制化宜忌</td>
<td width="84%" colspan="6" bgcolor="#FFFFFF" class="new"><%=skzhyj%></td>
  </tr>
<tr>
  <td  bgcolor="#FFFFFF" class="new">五行之性</td>
  <td colspan="6" bgcolor="#FFFFFF" class="new"><%=whzx%></td>
</tr>
<tr>
  <td  bgcolor="#FFFFFF" class="new">四柱五行生克中对应需补的脏腑和部位</td>
  <td colspan="6" bgcolor="#FFFFFF" class="new"><%=szwh%></td>
</tr>
<tr>
  <td  bgcolor="#FFFFFF" class="new">宜从事行业与方位</td>
  <td colspan="6" bgcolor="#FFFFFF" class="new"><%=hyhw%></td>
</tr>


</tbody>
</table><table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#5391EE"  style="MARGIN-BOTTOM: 10px; table-layout:fixed;word-wrap:break-word;">
  <tbody>
    <tr>
      <td  width="4%"  bgcolor="#FFFFFF" class="new">三<br />命<br />通<br />会</td>
      <td width="96%" colspan="6" bgcolor="#FFFFFF" class="new"><%=tf1%><br /><font color="#006699"><%=tf2%></font></td>
    </tr>
  </tbody>
</table>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#5391EE"  style="MARGIN-BOTTOM: 10px; table-layout:fixed;word-wrap:break-word;">
  <tbody>
    <tr>
      <td  width="4%" rowspan="3"  bgcolor="#FFFFFF" class="new">月<br /><br />日<br /><br />时<br /><br />命<br /><br />理</td>
      <td  width="12%"  bgcolor="#FFFFFF" class="new"><%=yue2%>月生</td>
      <td width="84%" colspan="6" bgcolor="#FFFFFF" class="new"><%=mingni1%></td>
    </tr>
    <tr>
      <td  bgcolor="#FFFFFF" class="new"><%=ri2%>日生</td>
      <td colspan="6" bgcolor="#FFFFFF" class="new"><%=mingni2%></td>
    </tr>
    <tr>
      <td  bgcolor="#FFFFFF" class="new"><%=DiZhi(hh2)%>时时生</td>
      <td colspan="6" bgcolor="#FFFFFF" class="new"><%=mingni3%></td>
    </tr>
  </tbody>
</table>
