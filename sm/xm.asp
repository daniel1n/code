<%
'处理用户信息
dim tiange,dige,renge1,renge2,renge
tiange=0
dige=0
renge1=0
renge2=0
renge=0
'姓
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
'名
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
'计算三才
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
            <td align="center" bgcolor="#FFFFFF" class="ttd"><font color="ababab">繁体</font></td>
            <td align="center" bgcolor="#FFFFFF" class="ttd"><font color="ababab">拼音</font></td>
            <td align="center" bgcolor="#FFFFFF" class="ttd"><font color="ababab">康熙笔划</font></td>
            <td align="center" bgcolor="#FFFFFF" class="ttd"><font color="ababab">字意五行</font></td>
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
      <td class="new" align="center" bgcolor="#FFFFFF">天格-&gt; <%=tiange%> (<%=getsancai(tiange)%>)&nbsp;&nbsp;人格-&gt; <%=renge%> (<%=getsancai(renge)%>)&nbsp;&nbsp;地格-&gt; <%=dige%> (<%=getsancai(dige)%>)&nbsp;&nbsp;外格-&gt; <%=waige%> (<%=getsancai(waige)%>)&nbsp;&nbsp;总格-&gt; <%=zhongge%> (<%=getsancai(zhongge)%>)&nbsp;&nbsp;(<a href="./htm_smzs.htm" target="_blank"><font color="#ff0000">什么是五格？</font></a>)</td>
    </tr>
  </tbody>
</table>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <TBODY>
      <tr>
        <TD class=ttd><p><font color="#0000FF">天格<%=tiangee%>的解析：</font><font color="#ff0000">天格数是先祖留传下来的，其数理对人影响不大。</font>
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
        <TD class=ttd>→ 详细解释：
<br />
<%=tgxx%></TD>
      </tr>      <tr>
        <TD class=ttd><p><font color="#0000FF">人格<%=rengee%>的解析：</font><font color="#ff0000">人格数又称主运，是整个姓名的中心点，影响人的一生命运。</font>
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
        <TD class=ttd>→ 详细解释：
<br />
<%=rgxx%></TD>
      </tr>  <tr>
        <TD class=ttd><p><font color="#0000FF">地格<%=digee%>的解析：</font><font color="#ff0000">地格数又称前运，影响人中年（36岁）以前的活动力。</font>
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
        <TD class=ttd>→ 详细解释：
<br />
<%=dgxx%></TD>
      </tr> <tr>
        <TD class=ttd><p><font color="#0000FF">总格<%=zhonggee%>的解析：</font><font color="#ff0000">总格又称后运，影响人中年（36岁）以后的命运。</font>
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
        <TD class=ttd>→ 详细解释：
<br />
<%=zgxx%></TD>
      </tr> <tr>
        <TD class=ttd><p><font color="#0000FF">外格<%=waigee%>的解析：</font><font color="#ff0000">外格又称变格，影响人的社交能力、智慧等，其数理不用重点去看。</font>
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
        <TD class=new>→ 详细解释：
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
'三才吉凶
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
<font color="#0000FF">对三才数理的影响:</font> 您的姓名三才配置为：<font color="#ff0000"><%=sancai%></font>。它具有如下数理诱导力，据此会对人生产生一定的影响。</TD>
      </tr><tr><td class=ttd><%=scyy%> (<%=scjx%>)
</td></tr><tr><td class=new><%=sancaicontent%>
</td></tr>
    </TBODY>
</TABLE>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <TBODY>
      <tr>
        <TD width="24%" class=ttd>
<font color="#0000FF">对基础运的影响:</font></TD>
        <TD width="76%" class=ttd><%=jcy%> <%=jcyjx%></TD>
      </tr><tr><td class=ttd>
<font color="#0000FF">对成功运的影响:</font></td>
        <td class=ttd><%=cgy%> <%=cgyjx%></td>
      </tr>
      <tr><td class=ttd>
<font color="#0000FF">对人际关系的影响:</font></td>
        <td class=ttd><%=rjgx%> <%=rjgxjx%></td>
      </tr><tr><td class=new>
<font color="#0000FF">对性格的影响:</font></td>
        <td class=new><%=xg%></td>
      </tr>
    </TBODY>
</TABLE><table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <TBODY>
      <tr>
        <TD width="24%" class=ttd>
<font color="#0000FF">人格<%=rengee%>有以下数理暗示:</font></TD>
        <TD width="76%" class=ttd><%=rgas%></TD>
      </tr><tr><td class=ttd>
<font color="#0000FF">地格<%=digee%>有以下数理暗示:</font></td>
        <td class=ttd><%=dgas%></td>
      </tr>
      <tr><td class=ttd>
<font color="#0000FF">总格<%=zhonggee%>有以下数理暗示:</font></td>
        <td class=ttd><%=zgas%></td>
      </tr><tr><td class=new>
<font color="#0000FF">地格<%=waigee%>有以下数理暗示:</font></td>
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
        <TD width="67%" class=new>总评及建议：<br /><br />
          <%if xmdf<60 then%>你的名字可能不是很好。强烈建议你换个名字试试，也许人生会因此而改变的。
		<%elseif xmdf<70 and xmdf>=60 then%>
		你的名字可能不太好，如果可能的话，不妨尝试改变一下，也许会有事半功倍之效。
		<%elseif xmdf<80 and xmdf>=70 then%>你的名字可能不太理想，要想赢得成功，必须比别人付出更多的艰辛和汗水。如果有条件，改个名字也未尝不可。
		<%elseif xmdf<80 and xmdf>=70 then%>你的名字一般。虽然人生路途中会遇到一些困难，但只要努力，还是会有很多收获的。如果有条件，改个名字也未尝不可。
		<%elseif xmdf<90 and xmdf>=80 then%>
		你的名字取得不错，如果与命理搭配得当，相信它会助你一生顺利的，祝你好运！
		<%elseif xmdf>=90 then%>
		你的名字取得非常棒，如果与命理搭配得当，成功与惊喜将会伴随你的一生。但千万注意不要失去上进心。<%end if%><div id="ft"><%=now()%> by Senlon</div></TD>
        <TD width="33%"  class=new><font color="0000ff">姓名五格评分：</font><span class=pf><%=xmdf%>分</span> </TD>
      </tr>
    </TBODY>
</TABLE>
