 
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
wnum1 = Len(Replace(bzwh, "��", Space(Len("��") + 1))) - Len(bzwh)
if wnum1>=3 then
mainw=mainw&"<strong>��</strong>��"
end if
if wnum1=0 then
mainq=mainq&"ȱ<strong>��</strong>"
end if
wnum2 = Len(Replace(bzwh, "ľ", Space(Len("ľ") + 1))) - Len(bzwh)
if wnum2>=3 then
mainw=mainw&"<strong>ľ</strong>��"
end if
if wnum2=0 then
mainq=mainq&"ȱ<strong>ľ</strong>"
end if
wnum3 = Len(Replace(bzwh, "ˮ", Space(Len("ˮ") + 1))) - Len(bzwh)
if wnum3>=3 then
mainw=mainw&"<strong>ˮ</strong>��"
end if
if wnum3=0 then
mainq=mainq&"ȱ<strong>ˮ</strong>"
end if
wnum4 = Len(Replace(bzwh, "��", Space(Len("��") + 1))) - Len(bzwh)
if wnum4>=3 then
mainw=mainw&"<strong>��</strong>��"
end if
if wnum4=0 then
mainq=mainq&"ȱ<strong>��</strong>"
end if
wnum5 = Len(Replace(bzwh, "��", Space(Len("��") + 1))) - Len(bzwh)
if wnum5>=3 then
mainw=mainw&"<strong>��</strong>��"
end if
if wnum5=0 then
mainq=mainq&"ȱ<strong>��</strong>"
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
<td  width="12%" rowspan="2" bgcolor="#FFFFFF" class="new">����ʱ�䣺</td>
  <td  width="9%" bgcolor="#FFFFFF" class="new">(����)</td>
  <td  width="12%" bgcolor="#FFFFFF" class="new"><%=nian1%>��</td>
  <td  width="10%" bgcolor="#FFFFFF" class="new"><%=yue1%>��</td>
  <td  width="10%" bgcolor="#FFFFFF" class="new"><%=ri1%>��</td>
  <td  width="11%" bgcolor="#FFFFFF" class="new"><%=hh1%>��</td>
  <td width="30%"  rowspan="6" bgcolor="#FFFFFF" class="new" style="padding-left:4px;padding-right:4px;">������<%=sx%>��<%=layin(ygz)%>��������<%=mainw%><%=mainq%>���������Ϊ<b><%=tgdzwh(dg1)%></b>������<%=siji(yue1)%>����<br />
    ��ͬ��<%=tnwh%>������<%=ynwh%>����<br />�������и��� : <%=wnum1%>����<%=wnum2%>��ľ��<%=wnum3%>��ˮ��<%=wnum4%>����<%=wnum5%>����
��</td>
 </tr>
 <tr>
  <td bgcolor="#FFFFFF" class="new" >(ũ��)</td>
  <td bgcolor="#FFFFFF" class="new" ><%=nian2%>��</td>
  <td bgcolor="#FFFFFF" class="new" ><%=yue2%>��</td>
  <td bgcolor="#FFFFFF" class="new" ><%=ri2%>��</td>
  <td bgcolor="#FFFFFF" class="new" ><%=DiZhi(hh2)%>ʱ</td>
 </tr>

 <tr>
  <td  colspan="2" bgcolor="#FFFFFF" class="new">���֣�</td>
  <td bgcolor="#FFFFFF" class="new" ><%=ygz%>��</td>
  <td bgcolor="#FFFFFF" class="new" ><%=mgz%>��</td>
  <td bgcolor="#FFFFFF" class="new" ><%=dgz%>��</td>
  <td bgcolor="#FFFFFF" class="new" ><%=tgz%>��</td>
 </tr>
 <tr>
  <td  colspan="2" bgcolor="#FFFFFF" class="new">���У�</td>
  <td bgcolor="#FFFFFF" class="new" ><%=tgdzwh(yg1)%><%=tgdzwh(yz1)%></td>
  <td bgcolor="#FFFFFF" class="new" ><%=tgdzwh(mg1)%><%=tgdzwh(mz1)%></td>
  <td bgcolor="#FFFFFF" class="new" ><%=tgdzwh(dg1)%><%=tgdzwh(dz1)%></td>
  <td bgcolor="#FFFFFF" class="new" ><%=tgdzwh(tg1)%><%=tgdzwh(tz1)%></td>
 </tr>
 <tr>
  <td  colspan="2" bgcolor="#FFFFFF" class="new">������</td>
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
      <td  width="16%"  bgcolor="#FFFFFF" class="new">��Ф����</td>
      <td width="84%" colspan="6" bgcolor="#FFFFFF" class="new">���ݷ�����������ФΪ��<%=sx%>��<br /> <%=sxgx%></td>
    </tr>
  </tbody>
</table><table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#5391EE"  style="MARGIN-BOTTOM: 10px; table-layout:fixed;word-wrap:break-word;">
  <tbody>
    <tr>
      <td  width="16%"  bgcolor="#FFFFFF" class="new">�Ը����</td>
      <td width="84%" colspan="6" bgcolor="#FFFFFF" class="new"> <%=xgfx%></td>
    </tr>
  </tbody>
</table><table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#5391EE"  style="MARGIN-BOTTOM: 10px; table-layout:fixed;word-wrap:break-word;">
  <tbody>
    <tr>
      <td  width="16%"  bgcolor="#FFFFFF" class="new">�������</td>
      <td width="84%" colspan="6" bgcolor="#FFFFFF" class="new"> <%=aqfx%></td>
    </tr>
  </tbody>
</table><table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#5391EE"  style="MARGIN-BOTTOM: 10px; table-layout:fixed;word-wrap:break-word;">
  <tbody>
    <tr>
      <td  width="16%"  bgcolor="#FFFFFF" class="new">��ҵ����</td>
      <td width="84%" colspan="6" bgcolor="#FFFFFF" class="new"> <%=syfx%></td>
    </tr>
  </tbody>
</table>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#5391EE"  style="MARGIN-BOTTOM: 10px; table-layout:fixed;word-wrap:break-word;">
  <tbody>
    <tr>
      <td  width="16%"  bgcolor="#FFFFFF" class="new">���˷���</td>
      <td width="84%" colspan="6" bgcolor="#FFFFFF" class="new"><%=cyfx%></td>
    </tr>
  </tbody>
</table>
<table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#5391EE"  style="MARGIN-BOTTOM: 10px; table-layout:fixed;word-wrap:break-word;">
  <tbody>
    <tr>
      <td  width="16%"  bgcolor="#FFFFFF" class="new">���˷���</td>
      <td width="84%" colspan="6" bgcolor="#FFFFFF" class="new"><%=cyfx%></td>
    </tr>
  </tbody>
</table>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
  <tbody>
    <tr>
      <td class="new"><div id="ft">˵��:�ո��������԰�������������Ϣ���з�����ΪƬ�����Ϣ�������ο���</div></td>
    </tr>
  </tbody>
</table>
