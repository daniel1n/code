<script language="JavaScript">
<!--
function checkbz()
{
var year=document.theform.y.value;
var month=document.theform.m.value;
var day=document.theform.d.value;
var hour=document.theform.hh.value;
var now=new Date();
var nowyear=now.getYear();
var nowmonth=now.getMonth();
if (year=='')
{
alert('请选择出生年份！');
document.theform.y.focus()
return false;
}
if (year>nowyear || year <=nowyear-100 || isNaN(year))
{
alert('请选择正确的出生年份！');
document.theform.y.focus()
return false;
}
if ( month=='')
{
alert('请选择出生月份！');
document.theform.m.focus()
return false;
}
if (day=='')
{
alert('请选择出生日期！');
document.theform.d.focus()
return false;
}
if ((month==2 && day>29) || ((month==1 || month==3 || month==5 || month==7 || month==8 || month==10|| month==12) && day>31) || ((month==4 || month==6 || month==9 || month==11 ) && day>30))
{
alert('请选择正确的出生日期！');
document.theform.d.focus()
return false;
}
if ( hour=='')
{
alert('请选择出生时间！');
document.theform.hh.focus()
return false;
}

var b=document.theform.b.value;
if (b=='')
{
alert('请选择您的性别！');
document.theform.b.focus()
return false;
}
var name=document.theform.name.value;
if (name=='')
{
alert('请输入您的姓名！');
document.theform.name.focus()
return false;
}
if (document.theform.name.value.length < 2 || document.theform.name.value.length>5)
{
alert("错误：姓名应在2-5个字之间！");
document.theform.name.focus();
return (false);
}
}
//-->
</script><table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <TBODY>
      <tr>
        <TD class=ttd style="PADDING-BOTTOM: 8px" vAlign=top>* 本测算来源于民间，仅供娱乐！ 
       </TD>
      </tr> <%if request.form("act")="ok" then%><%
	  sbnum=0
	  csrq=request.form("y")&request.form("m")&request.form("d")
	  for i = 1 to len(cstr(csrq)) 
sbnum=sbnum+mid(cstr(csrq),i,1) 
next
sbnum2=0
for i = 1 to len(cstr(sbnum)) 
sbnum2=sbnum2+mid(cstr(sbnum),i,1) 
next
sbnum3=0
for i = 1 to len(cstr(sbnum2)) 
sbnum3=sbnum3+mid(cstr(sbnum2),i,1) 
next
%>
	        <tr>
        <TD class=new style="PADDING-BOTTOM: 8px" vAlign=top><p>亲爱的<%=request.form("name")%>：</p>
          <p> 您出生于<%=request.form("y")%>年<%=request.form("m")%>月<%=request.form("d")%>日，今年<%=DateDiff("yyyy", ""&request.form("y")&""&"-"&""&request.form("m")&""&"-"&""&request.form("d")&"", date)%>岁（以下资料仅供参考）：</p>
          <p> 测测你上辈子是什么人：根据您的公元出生年月日，经过电脑计算，结果为<%=sbnum3%>，表示你上辈子是 <%set rs=server.createobject("adodb.recordset")
sql="select * from sbwr where num='"&sbnum3&"'"
rs.open sql,conn,1,1
intro=rs("intro")
rs.close
%> <font color=#ff0000><%=intro%></font><br /><br /><a href="sm.asp?sm=7"><font color=#0000ff>[重新测试]</font></a> </p></TD>
      </tr><%else%><form method="POST" action="" onSubmit="return checkbz();" name="theform">   <tr>
        <TD class=new style="PADDING-BOTTOM: 8px" vAlign=top><strong>公历生日：</strong><select size="1" name="y"><%for i=1900 to year(date())%><option value="<%=i%>" <%if nian1<>"" then%><%if i=cint(nian1) then%> selected<%end if%><%else%><%if i=year(date()) then%> selected<%end if%><%end if%>><%=i%></option><%next%></select>年<select name="m" size="1"><%for i=1 to 12%><option value="<%=i%>" <%if yue1<>"" then%><%if i=cint(yue1) then%> selected<%end if%><%else%><%if i=month(date()) then%> selected<%end if%><%end if%>><%=i%></option><%next%></select>月<select name="d" size="1"><%for i=1 to 31%><option value="<%=i%>" <%if ri1<>"" then%><%if i=cint(ri1) then%> selected<%end if%><%else%><%if i=day(date()) then%> selected<%end if%><%end if%>><%=i%></option><%next%></select>日<select size="1" name="hh"> <%for i=0 to 23%><option value="<%=i%>" <%if hh1<>"" then%><%if i=cint(hh1) then%> selected<%end if%><%end if%>><%=DiZhi(i)%><%=i%></option><%next%> </select>点<select size="1" name="fen"><option value="0">未知</option>
		<%for i=0 to 59%><option value="<%=i%>" <%if mm1<>"" then%><%if i=cint(mm1) then%> selected<%end if%><%end if%>><%=i%></option><%next%>
		</select>分
       </TD>
      </tr>  <tr>
        <TD class=new style="PADDING-BOTTOM: 8px" vAlign=top><strong>姓名：</strong>
          <input onkeyup="value=value.replace(/[^\u4E00-\u9FA5]/g,'')" name="name" type="text" value="<%=xing%><%=ming%>" />
          &nbsp;
          <select size="1" name="b">
<option value="男" <%if xingbie="男" then%>selected="selected"<%end if%>>男性</option>
<option value="女" <%if xingbie="女" then%>selected="selected"<%end if%>>女性</option>
          </select>&nbsp;<input type="hidden" name="act" value="ok" /><input name="cs" type="submit" value="开始测试" />
       </TD>
      </tr></form><%end if%>
    </TBODY>
  </TABLE>
