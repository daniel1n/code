<%if request.form("act")<>"ok" then%><%
'处理用户信息
tiange=0
dige=0
renge1=0
renge2=0
renge=0
'姓
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
end if%><SCRIPT language=javascript>
<!--
function Check(theForm)
{
var name1 = theForm.name1.value;
if (name1 == "") 
{
alert("请输入您的姓名！");
theForm.name1.value="";
theForm.name1.focus();
return false;
}
if (theForm.name1.value.length < 2 || theForm.name1.value.length>4)
{
alert("错误：名字应在2-4个字之间！");
theForm.name1.focus();
return (false);
}
if (name1.search(/[`1234567890-=\~!@#$%^&*()_+|<>;':",.?/ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz]/) != -1) 
{
alert("请务必输入简体汉字！");
theForm.name1.value = "";
theForm.name1.focus();
return false;
}
var name2 = theForm.name2.value;
if (name2 == "") 
{
alert("请输入您爱人的名字！");
theForm.name2.value="";
theForm.name2.focus();
return false;
}
if (theForm.name2.value.length < 2 || theForm.name2.value.length>4)
{
alert("错误：名字应在2-4个字之间！");
theForm.name2.focus();
return (false);
}
if (name2.search(/[`1234567890-=\~!@#$%^&*()_+|<>;':",.?/ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz]/) != -1) 
{
alert("请务必输入简体汉字！");
theForm.name2.value = "";
theForm.name2.focus();
return false;
}

}

//-->
</SCRIPT><table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#5391EE"  style="MARGIN-BOTTOM: 10px; table-layout:fixed;word-wrap:break-word;">  
  <tbody>
  <form name="form1" method="post" onSubmit="return Check(this)"  action=""> <tr>
      <td class=new bgcolor="#FFFFFF">配对测试 <input name="fs" type="radio" onClick="javacript:fs11.style.display='none';fs21.style.display='none';" value="0" checked="" />不测试八字相合程度　<input type="radio" name="fs" value="1" onClick="javacript:fs11.style.display='block';fs21.style.display='block';" />同时测试八字相合程度

    </tr>  <tr>
      <td class=new bgcolor="#FFFFFF"><strong>请输入第一个姓名：</strong>
          <input name="name1" type="text" value="<%=xing1%><%=xing2%><%=ming1%><%=ming2%>" />
          &nbsp;
          <select size="1" name="xing1">
<option value="1" <%if xing2="" then%>selected="selected"<%end if%>>单姓</option>
<option value="2"<%if xing2<>"" then%>selected="selected"<%end if%>>复姓</option>
          </select>
          &nbsp;
          <select size="1" name="sex1">
<option value="男" <%if xingbie="男" then%>selected="selected"<%end if%>>男性</option>
<option value="女" <%if xingbie="女" then%>selected="selected"<%end if%>>女性</option>
          </select><br /><br /><div id="fs11" style="display:none">
<div class="divh2"></div>公历/阳历生日：<select size="1" name="y1"><%for i=1900 to year(date())%><option value="<%=i%>" <%if nian1<>"" then%><%if i=cint(nian1) then%> selected<%end if%><%else%><%if i=year(date()) then%> selected<%end if%><%end if%>><%=i%></option><%next%></select>年<select name="m1" size="1"><%for i=1 to 12%><option value="<%=i%>" <%if yue1<>"" then%><%if i=cint(yue1) then%> selected<%end if%><%else%><%if i=month(date()) then%> selected<%end if%><%end if%>><%=i%></option><%next%></select>月<select name="d1" size="1"><%for i=1 to 31%><option value="<%=i%>" <%if ri1<>"" then%><%if i=cint(ri1) then%> selected<%end if%><%else%><%if i=day(date()) then%> selected<%end if%><%end if%>><%=i%></option><%next%></select>日<select size="1" name="h1"> <%for i=0 to 23%><option value="<%=i%>" <%if hh1<>"" then%><%if i=cint(hh1) then%> selected<%end if%><%end if%>><%=DiZhi(i)%><%=i%></option><%next%> </select>点<select size="1" name="f1"><option value="0">未知</option>
		<%for i=0 to 59%><option value="<%=i%>" <%if mm1<>"" then%><%if i=cint(mm1) then%> selected<%end if%><%end if%>><%=i%></option><%next%>
		</select>分&nbsp;<a title="如果您只知道生日的农历日期，不要紧，请点这里去查询公历日期" style="CURSOR: hand" onClick="window.open('/tools/wannianli.htm','httpcnnongli','left=0,top=0,width=680,height=480,scrollbars=no,resizable=no,status=no')" href="#wnl1"><font color="#008000">只知道农历生日请点此查询公历</font></a>
<hr size="1">
</div><br /><strong>请输入第二个姓名：</strong>
          <input type="text" name="name2" />
          &nbsp;
          <select size="1" name="xing2">
<option value="1">单姓</option>
<option value="2">复姓</option>
</select>
          &nbsp;
          <select size="1" name="sex2">
<option value="男" <%if xingbie="女" then%>selected="selected"<%end if%>>男性</option>
<option value="女" <%if xingbie="男" then%>selected="selected"<%end if%>>女性</option>
</select><input type="hidden" name="act" value="ok" /><br /><br /><div id="fs21" style="display:none">
公历/阳历生日：<select size="1" name="y2"><%for i=1900 to year(date())%><option value="<%=i%>" <%if i=year(date()) then%> selected<%end if%>><%=i%></option><%next%></select>年<select name="m2" size="1"><%for i=1 to 12%><option value="<%=i%>"<%if i=month(date()) then%> selected<%end if%>><%=i%></option><%next%></select>月<select name="d2" size="1"><%for i=1 to 31%><option value="<%=i%>" <%if i=day(date()) then%> selected<%end if%>><%=i%></option><%next%></select>日<select size="1" name="h2"> <%for i=0 to 23%><option value="<%=i%>" ><%=DiZhi(i)%><%=i%></option><%next%> </select>点<select size="1" name="f2"><option value="0">未知</option>
		<%for i=0 to 59%><option value="<%=i%>"><%=i%></option><%next%>
		</select>分
       &nbsp;<a title="如果您只知道生日的农历日期，不要紧，请点这里去查询公历日期" style="CURSOR: hand" onClick="window.open('/tools/wannianli.htm','httpcnnongli','left=0,top=0,width=680,height=480,scrollbars=no,resizable=no,status=no')" href="#wnl1"><font color="#008000">只知道农历生日请点此查询公历</font></a>
<hr size="1">
</div>
<input type="submit" value="开始测试" style="cursor:hand;" />
      </td>
    </tr> </form>
  </tbody>
</table>
<%else%>
<%fs=request.form("fs")
name1=request.form("name1")
name2=request.form("name2")
xing1=request.form("xing1")
xing2=request.form("xing2")
sex1=request.form("sex1")
sex2=request.form("sex2")
y1=request.form("y1")
m1=request.form("m1")
d1=request.form("d1")
h1=request.form("h1")
f1=request.form("f1")
y2=request.form("y2")
m2=request.form("m2")
d2=request.form("d2")
h2=request.form("h2")
f2=request.form("f2")
tiange=0
dige=0
renge1=0
renge2=0
renge=0
bihua1=0
bihua2=0
nxing1=mid(name1,1,1)
nxing11=mid(name1,1,1)
bihua1=getnum(nxing1)
tiange1=bihua1+1
tiangee1=bihua1+1
renge1=bihua1
rengee1=bihua1
'判断第一个名字 姓
if xing1 = "2" then
nxing2=mid(name1,2,1)
nxing22=mid(name1,2,1)
bihua2=getnum(nxing2)
tiange1=bihua1+bihua2
tiangee1=bihua1+bihua2
renge1=bihua2
rengee1=bihua2
bihua3=0
bihua4=0
nming1=mid(name1,3,1)
nming12=mid(name1,3,1)
bihua3=getnum(nming1)
dige1=bihua3+1
digee1=bihua3+1
renge2=bihua3
rengee2=bihua3
if mid(name1,4,1) <> "" then
nming2=mid(name1,4,1)
nming22=mid(name1,4,1)
bihua4=getnum(nming2)
dige1=bihua3+bihua4
digee1=bihua3+bihua4
end if
else
bihua3=0
bihua4=0
nming1=mid(name1,2,1)
nming12=mid(name1,2,1)
bihua3=getnum(nming1)
dige1=bihua3+1
digee1=bihua3+1
renge2=bihua3
rengee2=bihua3
waige1=waige1+1
waigee1=waigee1+1
if mid(name1,3,1) <> "" then
nming2=mid(name1,3,1)
nming22=mid(name1,3,1)
bihua4=getnum(nming2)
dige1=bihua3+bihua4
digee1=bihua3+bihua4
end if
end if
zhongge1=bihua1+bihua2+bihua3+bihua4
zhonggee1=bihua1+bihua2+bihua3+bihua4
renge1=renge1+renge2
rengee1=rengee1+rengee2
waige1=zhongge1-renge1
waigee1=zhonggee1-rengee1
if nxing2 = "" then
waige1=waige1+1
waigee1=waigee1+1
end if
if nming2 = "" then
waige1=waige1+1
waigee1=waigee1+1
end if
%><table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#5391EE"  style="MARGIN-BOTTOM: 10px; table-layout:fixed;word-wrap:break-word;">  
  <tbody>    <tr>
      <td colspan="3" class=new bgcolor="#FFFFFF">姓名：<%=name1%>  性别：<%=sex1%>&nbsp;<%if fs=1 then%>生日:<%=y1%>年<%=m1%>月<%=d1%>日<%=h1%>时<%=f1%>分<%end if%></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="2" cellspacing="1"  style="MARGIN-BOTTOM: 10px; table-layout:fixed;word-wrap:break-word;">
        <tbody>
          <tr>
            <td bgcolor="#FFFFFF"></td>
            <td align="center" bgcolor="#FFFFFF"><font color="ababab">繁体</font></td>
            <td align="center" bgcolor="#FFFFFF"><font color="ababab">拼音</font></td>
            <td align="center" bgcolor="#FFFFFF"><font color="ababab">笔划</font></td>
            <td align="center" bgcolor="#FFFFFF"><font color="ababab">五行</font></td>
            </tr>
          <tr>
            <td  align="center" bgcolor="#FFFFFF" class="new2"><%=nxing1%></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><font color="aaaaaa"><%=GbToBig(nxing1)%></font></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><font color="aaaaaa"><%=c(mid(nxing11,1,1))%></font></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><%=bihua1%></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><%=getzywh(nxing11)%></td>
          </tr>
          <tr>
         <%if nxing2<>"" then%>   <td align="center" bgcolor="#FFFFFF" class="new2"><%=nxing2%></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><font color="aaaaaa"><%=GbToBig(nxing2)%></font></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><font color="aaaaaa"><%=c(nxing22)%></font></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><%=bihua2%></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><%=getzywh(nxing22)%></td>
          </tr><%end if%>
          <tr>
            <td align="center" bgcolor="#FFFFFF" class="new2"><%=nming1%></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><font color="aaaaaa"><%=GbToBig(nming1)%></font></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><font color="aaaaaa"><%=c(nming12)%></font></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><%=bihua3%></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><%=getzywh(nming12)%></td>
          </tr>
          <%if nming2<>"" then%><tr>
            <td align="center" bgcolor="#FFFFFF" class="new2"><%=nming2%></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><font color="aaaaaa"><%=GbToBig(nming2)%></font></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><font color="aaaaaa"><%=c(nming22)%></font></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><%=bihua4%></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><%=getzywh(nming22)%></td>
          </tr><%end if%>
        </tbody>
      </table></td>
      <td width="25%" colspan="-3" align="center" bgcolor="#FFFFFF"  class="new2">天格-&gt; <%=tiange1%> (<%=getsancai(tiange1)%>) <%set rs=server.createobject("adodb.recordset")
sql="select * from 81 where num='"&tiangee1&"'"
rs.open sql,conn,1,1
tgjx=rs("jx")
rs.close
tgf1=getpf(tgjx)
%> ->(<font color=red><%=tgjx%></font>)<br />
      <p>人格-&gt; <%=renge1%> (<%=getsancai(renge1)%>)<%set rs=server.createobject("adodb.recordset")
sql="select * from 81 where num='"&rengee1&"'"
rs.open sql,conn,1,1
rgjx=rs("jx")
rs.close
rgf1=getpf(rgjx)
%> ->(<font color=red><%=rgjx%></font>)</p>        <p>地格-&gt; <%=dige1%> (<%=getsancai(dige1)%>)<%set rs=server.createobject("adodb.recordset")
sql="select * from 81 where num='"&digee1&"'"
rs.open sql,conn,1,1
dgjx=rs("jx")
rs.close
dgf1=getpf(dgjx)
%> ->(<font color=red><%=dgjx%></font>)</p></td>
      <td width="25%"  class="new2" align="center" bgcolor="#FFFFFF">外格-&gt; <%=waige1%> (<%=getsancai(waige1)%>)<%set rs=server.createobject("adodb.recordset")
sql="select * from 81 where num='"&waigee1&"'"
rs.open sql,conn,1,1
wgjx=rs("jx")
rs.close
wgf1=getpf(wgjx)
%> ->(<font color=red><%=wgjx%></font>)<br />
      <p>　</p>        <p>总格-&gt; <%=zhongge1%> (<%=getsancai(zhongge1)%>)<%set rs=server.createobject("adodb.recordset")
sql="select * from 81 where num='"&zhonggee1&"'"
rs.open sql,conn,1,1
zgjx=rs("jx")
rs.close
zgf1=getpf(zgjx)
%> ->(<font color=red><%=zgjx%></font>)</p></td>
    </tr>
    <tr>
      <td colspan="3" class=new bgcolor="#FFFFFF">姓名蕴含的个性分析：<%sancai1=getsancai(tiangee1)+getsancai(rengee1)+getsancai(digee1)
	  sqlsancai="select * from sancai where title='"&sancai1&"'"
	  set rssancai=conn.execute(sqlsancai)
if not (rssancai.bof and rssancai.eof) then
xg1=rssancai("xg")
end if
%><%=xg1%></td>
    </tr>
  </tbody>
</table>
<%
name2=request.form("name2")
name2=request.form("name2")
xing1=request.form("xing1")
xing2=request.form("xing2")
sex1=request.form("sex1")
sex2=request.form("sex2")
ntiange=0
ndige=0
nrenge1=0
nrenge2=0
nrenge=0
nbihua1=0
nbihua2=0
n2xing1=mid(name2,1,1)
n2xing11=mid(name2,1,1)
nbihua1=getnum(n2xing1)
ntiange1=nbihua1+1
ntiangee1=nbihua1+1
nrenge1=nbihua1
nrengee1=nbihua1
'判断第一个名字 姓
if xing2 = "2" then
n2xing2=mid(name2,2,1)
n2xing22=mid(name2,2,1)
nbihua2=getnum(n2xing2)
ntiange1=nbihua1+nbihua2
ntiangee1=nbihua1+nbihua2
nrenge1=nbihua2
nrengee1=nbihua2
nbihua3=0
nbihua4=0
n2ming1=mid(name2,3,1)
n2ming12=mid(name2,3,1)
nbihua3=getnum(n2ming1)
ndige1=nbihua3+1
ndigee1=nbihua3+1
nrenge2=nbihua3
nrengee2=nbihua3
if mid(name2,4,1) <> "" then
n2ming2=mid(name2,4,1)
n2ming22=mid(name2,4,1)
nbihua4=getnum(n2ming2)
ndige1=nbihua3+nbihua4
ndigee1=nbihua3+nbihua4
end if
else
nbihua3=0
nbihua4=0
n2ming1=mid(name2,2,1)
n2ming12=mid(name2,2,1)
nbihua3=getnum(n2ming1)
ndige1=nbihua3+1
ndigee1=nbihua3+1
nrenge2=nbihua3
nrengee2=nbihua3
nwaige1=nwaige1+1
nwaigee1=nwaigee1+1
if mid(name2,3,1) <> "" then
n2ming2=mid(name2,3,1)
n2ming22=mid(name2,3,1)
nbihua4=getnum(n2ming2)
ndige1=nbihua3+nbihua4
ndigee1=nbihua3+nbihua4
end if
end if
nzhongge1=nbihua1+nbihua2+nbihua3+nbihua4
nzhonggee1=nbihua1+nbihua2+nbihua3+nbihua4
nrenge1=nrenge1+nrenge2
nrengee1=nrengee1+nrengee2
nwaige1=nzhongge1-nrenge1
nwaigee1=nzhonggee1-nrengee1
if n2xing2 = "" then
nwaige1=nwaige1+1
nwaigee1=nwaigee1+1
end if
if n2ming2 = "" then
nwaige1=nwaige1+1
nwaigee1=nwaigee1+1
end if
%><table width="100%" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#5391EE"  style="MARGIN-BOTTOM: 10px; table-layout:fixed;word-wrap:break-word;">  
  <tbody> <tr>
      <td colspan="3" class=new bgcolor="#FFFFFF">姓名：<%=name2%>  性别：<%=sex2%> &nbsp;<%if fs=1 then%>生日:<%=y2%>年<%=m2%>月<%=d2%>日<%=h2%>时<%=f2%>分<%end if%></td>
    </tr>
    <tr>
      <td bgcolor="#FFFFFF"><table width="100%" border="0" align="center" cellpadding="2" cellspacing="1"  style="MARGIN-BOTTOM: 10px; table-layout:fixed;word-wrap:break-word;">
        <tbody>
          <tr>
            <td bgcolor="#FFFFFF"></td>
            <td align="center" bgcolor="#FFFFFF"><font color="ababab">繁体</font></td>
            <td align="center" bgcolor="#FFFFFF"><font color="ababab">拼音</font></td>
            <td align="center" bgcolor="#FFFFFF"><font color="ababab">笔划</font></td>
            <td align="center" bgcolor="#FFFFFF"><font color="ababab">五行</font></td>
            </tr>
          <tr>
            <td  align="center" bgcolor="#FFFFFF" class="new2"><%=n2xing1%></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><font color="aaaaaa"><%=GbToBig(n2xing1)%></font></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><font color="aaaaaa"><%=c(mid(n2xing11,1,1))%></font></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><%=nbihua1%></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><%=getzywh(n2xing11)%></td>
          </tr>
          <tr>
         <%if n2xing2<>"" then%>   <td align="center" bgcolor="#FFFFFF" class="new2"><%=n2xing2%></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><font color="aaaaaa"><%=GbToBig(n2xing2)%></font></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><font color="aaaaaa"><%=c(n2xing22)%></font></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><%=nbihua2%></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><%=getzywh(n2xing22)%></td>
          </tr><%end if%>
          <tr>
            <td align="center" bgcolor="#FFFFFF" class="new2"><%=n2ming1%></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><font color="aaaaaa"><%=GbToBig(n2ming1)%></font></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><font color="aaaaaa"><%=c(n2ming12)%></font></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><%=nbihua3%></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><%=getzywh(n2ming12)%></td>
          </tr>
          <%if n2ming2<>"" then%><tr>
            <td align="center" bgcolor="#FFFFFF" class="new2"><%=n2ming2%></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><font color="aaaaaa"><%=GbToBig(n2ming2)%></font></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><font color="aaaaaa"><%=c(n2ming22)%></font></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><%=nbihua4%></td>
            <td align="center" bgcolor="#FFFFFF" class="new2"><%=getzywh(n2ming22)%></td>
          </tr><%end if%>
        </tbody>
      </table></td>
      <td width="25%" colspan="-3" align="center" bgcolor="#FFFFFF"  class="new2">天格-&gt; <%=ntiange1%> (<%=getsancai(ntiange1)%>) <%set rs=server.createobject("adodb.recordset")
sql="select * from 81 where num='"&ntiangee1&"'"
rs.open sql,conn,1,1
tgjx=rs("jx")
rs.close
tgf2=getpf(tgjx)
%> ->(<font color=red><%=tgjx%></font>)<br />
      <p>人格-&gt; <%=nrenge1%> (<%=getsancai(nrenge1)%>)<%set rs=server.createobject("adodb.recordset")
sql="select * from 81 where num='"&nrengee1&"'"
rs.open sql,conn,1,1
rgjx=rs("jx")
rs.close
rgf2=getpf(rgjx)
%> ->(<font color=red><%=rgjx%></font>)</p>        <p>地格-&gt; <%=ndige1%> (<%=getsancai(ndige1)%>)<%set rs=server.createobject("adodb.recordset")
sql="select * from 81 where num='"&ndigee1&"'"
rs.open sql,conn,1,1
dgjx=rs("jx")
rs.close
dgf2=getpf(dgjx)
%> ->(<font color=red><%=dgjx%></font>)</p></td>
      <td width="25%"  class="new2" align="center" bgcolor="#FFFFFF">外格-&gt; <%=nwaige1%> (<%=getsancai(nwaige1)%>)<%set rs=server.createobject("adodb.recordset")
sql="select * from 81 where num='"&nwaigee1&"'"
rs.open sql,conn,1,1
wgjx=rs("jx")
rs.close
wgf2=getpf(wgjx)
%> ->(<font color=red><%=wgjx%></font>)<br />
      <p>　</p>        <p>总格-&gt; <%=nzhongge1%> (<%=getsancai(nzhongge1)%>)<%set rs=server.createobject("adodb.recordset")
sql="select * from 81 where num='"&nzhonggee1&"'"
rs.open sql,conn,1,1
zgjx=rs("jx")
rs.close
zgf2=getpf(zgjx)
%> ->(<font color=red><%=zgjx%></font>)</p></td>
    </tr>
    <tr>
      <td colspan="3" class=new bgcolor="#FFFFFF">姓名蕴含的个性分析：<%sancai1=getsancai(ntiangee1)+getsancai(nrengee1)+getsancai(ndigee1)
	  sqlsancai="select * from sancai where title='"&sancai1&"'"
	  set rssancai=conn.execute(sqlsancai)
if not (rssancai.bof and rssancai.eof) then
xg1=rssancai("xg")

end if
%><%=xg1%></td>
    </tr>
  </tbody>
</table>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <TBODY>
      <tr><%
	  n1=abs(rgf1-rgf2)
	  n2=abs(dgf1-dgf2)
	  n3=abs(zgf1-zgf2)
	  n4=abs(tgf1-tgf2)
	  n5=abs(wgf1-wgf2)
	  zf=(n1+n2+n3)+(n4+n5)/5
	  zf=round(100-zf)
	  %>
        <TD class=ttd>经过分析，[<%=name1%>]和[<%=name2%>]的姓名配对评分如下：</TD>
      </tr>      <tr>
        <TD class=new><span class="red">姓名缘份指数：<%=zf%></span></TD>
        </tr>  <tr>
        <TD class=new><%if sex1=sex2 then%><font color=red>本站是按中国民俗学的一些测算方法来计算的，暂时不支持同性缘份的测试</font><%else%><font color=green>
		 <%if zf<60 then%>你们的姓名五格可能不是太合，不过八字配合所起的作用更大，另外彼此的努力也会让你们改善关系，记住：事在人为！
		<%elseif zf<70 and zf>=60 then%>
		你们的姓名五格相合程度马马虎虎，不过八字配合所起的作用更大，继续努力改善关系，会对你们的关系有帮助的！ 
		<%elseif zf<80 and zf>=70 then%>你们的姓名五格相合一般！ 
		<%elseif zf<80 and zf>=70 then%>你们的姓名五格相合程度还不错哟！ 
		<%elseif zf<90 and zf>=80 then%>
		你们的姓名五格相合程度相当棒！ 
		<%elseif zf>=90then%>
		你们的姓名五格相合程度太棒了！！恭喜！</font><%if name1=name2 then%><br /><font color=#0000ff>^_^ 你们俩同名同姓嘛，真巧哟！</font> <%end if%><%end if%>
		<%end if%></TD>
        </tr>
    </TBODY>
</TABLE>
<%end if%>