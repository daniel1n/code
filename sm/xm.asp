<%
'�����û���Ϣ
dim tiange,dige,renge1,renge2,renge
tiange=0
dige=0
renge1=0
renge2=0
renge=0
'��
dim xing1,xing2,bihua1,bihua2
bihua1=0
bihua2=0
xing1=mid(xing,1,1)
xing11=mid(xing,1,1)
bihua1=getnum(xing1)
tiange=bihua1+1
tiangee=bihua1+1
renge1=bihua1
if mid(xing,2,1) <> "" then
xing2=mid(xing,2,1)
xing22=mid(xing,2,1)
bihua2=getnum(xing2)
tiange=bihua1+bihua2
tiangee=bihua1+bihua2
renge1=bihua2 
end if
'��
dim ming1,ming2,bihua3,bihua4
bihua3=0
bihua4=0
ming1=mid(ming,1,1)
ming11=mid(ming,1,1)
bihua3=getnum(ming1)
dige=bihua3+1
digee=bihua3+1
renge2=bihua3
if mid(ming,2,1) <> "" then
ming2=mid(ming,2,1)
ming22=mid(ming,2,1)
bihua4=getnum(ming2)
dige=bihua3+bihua4
digee=bihua3+bihua4
end if
zhongge=bihua1+bihua2+bihua3+bihua4
zhonggee=bihua1+bihua2+bihua3+bihua4
'��������
renge=renge1+renge2
rengee=renge1+renge2
waige=zhongge-renge
waigee=zhonggee-rengee
if mid(xing,2,1) = "" then
waige=waige+1
waigee=waigee+1
end if
if mid(ming,2,1) = "" then
waige=waige+1
waigee=waigee+1
end if%> 
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
        <tbody>
          <tr>
            <td bgcolor="#FFFFFF" class="ttd">&nbsp;</td>
            <td align="center" bgcolor="#FFFFFF" class="ttd"><font color="ababab">����</font></td>
            <td align="center" bgcolor="#FFFFFF" class="ttd"><font color="ababab">ƴ��</font></td>
            <td align="center" bgcolor="#FFFFFF" class="ttd"><font color="ababab">�����ʻ�</font></td>
            <td align="center" bgcolor="#FFFFFF" class="ttd"><font color="ababab">��������</font></td>
          </tr>
          <tr>
            <td align="center" bgcolor="#FFFFFF" class="ttd"><%=xing1%></td>
            <td align="center" bgcolor="#FFFFFF" class="ttd"><%=GbToBig(xing1)%></td>
            <td align="center" bgcolor="#FFFFFF" class="ttd"><%=c(mid(xing,1,1))%></td>
            <td align="center" bgcolor="#FFFFFF" class="ttd"><%=bihua1%></td>
            <td align="center" bgcolor="#FFFFFF" class="ttd"><%=getzywh(xing11)%></td>
          </tr>
          <tr>
            <%if mid(xing,2,1)<>"" then%>
            <td align="center" bgcolor="#FFFFFF" class="ttd"><%=xing2%></td>
            <td align="center" bgcolor="#FFFFFF" class="ttd"><%=GbToBig(xing2)%></td>
            <td align="center" bgcolor="#FFFFFF" class="ttd"><%=c(mid(xing,2,1))%></td>
            <td align="center" bgcolor="#FFFFFF" class="ttd"><%=bihua2%></td>
            <td align="center" bgcolor="#FFFFFF" class="ttd"><%=getzywh(xing22)%></td>
          </tr>
          <%end if%>
          <tr>
            <td align="center" bgcolor="#FFFFFF" class="ttd"><%=ming1%></td>
            <td align="center" bgcolor="#FFFFFF" class="ttd"><%=GbToBig(ming1)%></td>
            <td align="center" bgcolor="#FFFFFF" class="ttd"><%=c(mid(ming,1,1))%></td>
            <td align="center" bgcolor="#FFFFFF" class="ttd"><%=bihua3%></td>
            <td align="center" bgcolor="#FFFFFF" class="ttd"><%=getzywh(ming11)%></td>
          </tr>
          <%if mid(ming,2,1)<>"" then%>
          <tr>
            <td align="center" bgcolor="#FFFFFF" class="new"><%=ming2%></td>
            <td align="center" bgcolor="#FFFFFF" class="new"><%=GbToBig(ming2)%></td>
            <td align="center" bgcolor="#FFFFFF" class="new"><%=c(mid(ming,2,1))%></td>
            <td align="center" bgcolor="#FFFFFF" class="new"><%=bihua4%></td>
            <td align="center" bgcolor="#FFFFFF" class="new"><%=getzywh(ming22)%></td>
          </tr>
          <%end if%>
        </tbody>
      </table>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
  <tbody>
    <tr>
      <td class="new" align="center" bgcolor="#FFFFFF">���-&gt; <%=tiange%> (<%=getsancai(tiange)%>)&nbsp;&nbsp;�˸�-&gt; <%=renge%> (<%=getsancai(renge)%>)&nbsp;&nbsp;�ظ�-&gt; <%=dige%> (<%=getsancai(dige)%>)&nbsp;&nbsp;���-&gt; <%=waige%> (<%=getsancai(waige)%>)&nbsp;&nbsp;�ܸ�-&gt; <%=zhongge%> (<%=getsancai(zhongge)%>)&nbsp;&nbsp;(<a href="./htm_smzs.htm" target="_blank"><font color="#ff0000">ʲô�����</font></a>)</td>
    </tr>
  </tbody>
</table>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <TBODY>
      <tr>
        <TD class=ttd><p><font color="#0000FF">���<%=tiangee%>�Ľ�����</font><font color="#ff0000">��������������������ģ����������Ӱ�첻��</font>
</TD>
      </tr> <%set rs=server.createobject("adodb.recordset")
sql="select * from 81 where num='"&tiangee&"'"
rs.open sql,conn,1,1
tgyy=rs("yy")
tgjx=rs("jx")
tgas=rs("as")
tgxx=rs("content")
rs.close
%> <tr>
        <TD class=ttd>
<%=tgyy%>&nbsp;<font color=red>(<%=tgjx%>)</font></TD>
      </tr> <tr>
        <TD class=ttd>�� ��ϸ���ͣ�
<br />
<%=tgxx%></TD>
      </tr>      <tr>
        <TD class=ttd><p><font color="#0000FF">�˸�<%=rengee%>�Ľ�����</font><font color="#ff0000">�˸����ֳ����ˣ����������������ĵ㣬Ӱ���˵�һ�����ˡ�</font>
</TD>
      </tr> <%
sql="select * from 81 where num='"&rengee&"'"
rs.open sql,conn,1,1
rgyy=rs("yy")
rgjx=rs("jx")
rgas=rs("as")
rgxx=rs("content")
rs.close
%> <tr>
        <TD class=ttd>
<%=rgyy%>&nbsp;<font color=red>(<%=rgjx%>)</font></TD>
      </tr> <tr>
        <TD class=ttd>�� ��ϸ���ͣ�
<br />
<%=rgxx%></TD>
      </tr>  <tr>
        <TD class=ttd><p><font color="#0000FF">�ظ�<%=digee%>�Ľ�����</font><font color="#ff0000">�ظ����ֳ�ǰ�ˣ�Ӱ�������꣨36�꣩��ǰ�Ļ����</font>
</TD>
      </tr> <%
sql="select * from 81 where num='"&digee&"'"
rs.open sql,conn,1,1
dgyy=rs("yy")
dgjx=rs("jx")
dgas=rs("as")
dgxx=rs("content")
rs.close
%> <tr>
        <TD class=ttd>
<%=dgyy%>&nbsp;<font color=red>(<%=dgjx%>)</font></TD>
      </tr> <tr>
        <TD class=ttd>�� ��ϸ���ͣ�
<br />
<%=dgxx%></TD>
      </tr> <tr>
        <TD class=ttd><p><font color="#0000FF">�ܸ�<%=zhonggee%>�Ľ�����</font><font color="#ff0000">�ܸ��ֳƺ��ˣ�Ӱ�������꣨36�꣩�Ժ�����ˡ�</font>
</TD>
      </tr> <%
sql="select * from 81 where num='"&zhonggee&"'"
rs.open sql,conn,1,1
zgyy=rs("yy")
zgjx=rs("jx")
zgas=rs("as")
zgxx=rs("content")
rs.close
%> <tr>
        <TD class=ttd>
<%=zgyy%>&nbsp;<font color=red>(<%=zgjx%>)</font></TD>
      </tr> <tr>
        <TD class=ttd>�� ��ϸ���ͣ�
<br />
<%=zgxx%></TD>
      </tr> <tr>
        <TD class=ttd><p><font color="#0000FF">���<%=waigee%>�Ľ�����</font><font color="#ff0000">����ֳƱ��Ӱ���˵��罻�������ǻ۵ȣ����������ص�ȥ����</font>
</TD>
      </tr> <%
sql="select * from 81 where num='"&waigee&"'"
rs.open sql,conn,1,1
wgyy=rs("yy")
wgjx=rs("jx")
wgas=rs("as")
wgxx=rs("content")
rs.close
%> <tr>
        <TD class=ttd>
<%=wgyy%>&nbsp;<font color=red>(<%=wgjx%>)</font></TD>
      </tr> <tr>
        <TD class=new>�� ��ϸ���ͣ�
<br />
<%=wgxx%></TD>
      </tr>
    </TBODY>
</TABLE>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <TBODY>
      <tr>
        <TD class=ttd><%
dim sancai
sancai=getsancai(tiange)+getsancai(renge)+getsancai(dige)
on error resume next
'���ż���
dim sqlsancai,sancaicontent
sqlsancai="select * from sancai where title='"&sancai&"'"
set rssancai=conn.execute(sqlsancai)
if not (rssancai.bof and rssancai.eof) then
sancaicontent=rssancai("content")
scyy=rssancai("yy")
scjx=rssancai("jx")
jcy=rssancai("jcy")
jcyjx=rssancai("jx1")
cgy=rssancai("cgy")
cgyjx=rssancai("jx1")
rjgx=rssancai("rjgx")
rjgxjx=rssancai("jx3")
xg=rssancai("xg")

end if
%>
<font color="#0000FF">�����������Ӱ��:</font> ����������������Ϊ��<font color="#ff0000"><%=sancai%></font>�����������������յ������ݴ˻����������һ����Ӱ�졣</TD>
      </tr><tr><td class=ttd><%=scyy%> (<%=scjx%>)
</td></tr><tr><td class=new><%=sancaicontent%>
</td></tr>
    </TBODY>
</TABLE>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <TBODY>
      <tr>
        <TD width="24%" class=ttd>
<font color="#0000FF">�Ի����˵�Ӱ��:</font></TD>
        <TD width="76%" class=ttd><%=jcy%> <%=jcyjx%></TD>
      </tr><tr><td class=ttd>
<font color="#0000FF">�Գɹ��˵�Ӱ��:</font></td>
        <td class=ttd><%=cgy%> <%=cgyjx%></td>
      </tr>
      <tr><td class=ttd>
<font color="#0000FF">���˼ʹ�ϵ��Ӱ��:</font></td>
        <td class=ttd><%=rjgx%> <%=rjgxjx%></td>
      </tr><tr><td class=new>
<font color="#0000FF">���Ը��Ӱ��:</font></td>
        <td class=new><%=xg%></td>
      </tr>
    </TBODY>
</TABLE><table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <TBODY>
      <tr>
        <TD width="24%" class=ttd>
<font color="#0000FF">�˸�<%=rengee%>����������ʾ:</font></TD>
        <TD width="76%" class=ttd><%=rgas%></TD>
      </tr><tr><td class=ttd>
<font color="#0000FF">�ظ�<%=digee%>����������ʾ:</font></td>
        <td class=ttd><%=dgas%></td>
      </tr>
      <tr><td class=ttd>
<font color="#0000FF">�ܸ�<%=zhonggee%>����������ʾ:</font></td>
        <td class=ttd><%=zgas%></td>
      </tr><tr><td class=new>
<font color="#0000FF">�ظ�<%=waigee%>����������ʾ:</font></td>
        <td class=new><%=wgas%></td>
      </tr>
    </TBODY>
</TABLE><table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <TBODY>
      <tr><%
	  xmdf=getpf(tgjx)/10+getpf(rgjx)+getpf(dgjx)+getpf(zgjx)+getpf(wgjx)/10+getpf(scjx)/4+getpf(jcyjx)/4+getpf(cgyjx)/4+getpf(rjgxjx)/4
	  if zhonggee>60 then
	  xmdf=xmdf-5
	  end if
	  xmdf=50+xmdf%>
        <TD width="67%" class=new>���������飺<br /><br />
          <%if xmdf<60 then%>������ֿ��ܲ��Ǻܺá�ǿ�ҽ����㻻���������ԣ�Ҳ����������˶��ı�ġ�
		<%elseif xmdf<70 and xmdf>=60 then%>
		������ֿ��ܲ�̫�ã�������ܵĻ����������Ըı�һ�£�Ҳ������°빦��֮Ч��
		<%elseif xmdf<80 and xmdf>=70 then%>������ֿ��ܲ�̫���룬Ҫ��Ӯ�óɹ�������ȱ��˸�������ļ����ͺ�ˮ��������������ĸ�����Ҳδ�����ɡ�
		<%elseif xmdf<80 and xmdf>=70 then%>�������һ�㡣��Ȼ����·;�л�����һЩ���ѣ���ֻҪŬ�������ǻ��кܶ��ջ�ġ�������������ĸ�����Ҳδ�����ɡ�
		<%elseif xmdf<90 and xmdf>=80 then%>
		�������ȡ�ò���������������õ���������������һ��˳���ģ�ף����ˣ�
		<%elseif xmdf>=90 then%>
		�������ȡ�÷ǳ�����������������õ����ɹ��뾪ϲ����������һ������ǧ��ע�ⲻҪʧȥ�Ͻ��ġ�<%end if%><div id="ft"><%=now()%> by Senlon</div></TD>
        <TD width="33%"  class=new><font color="0000ff">����������֣�</font><span class=pf><%=xmdf%>��</span> </TD>
      </tr>
    </TBODY>
</TABLE>
