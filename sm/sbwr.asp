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
alert('��ѡ�������ݣ�');
document.theform.y.focus()
return false;
}
if (year>nowyear || year <=nowyear-100 || isNaN(year))
{
alert('��ѡ����ȷ�ĳ�����ݣ�');
document.theform.y.focus()
return false;
}
if ( month=='')
{
alert('��ѡ������·ݣ�');
document.theform.m.focus()
return false;
}
if (day=='')
{
alert('��ѡ��������ڣ�');
document.theform.d.focus()
return false;
}
if ((month==2 && day>29) || ((month==1 || month==3 || month==5 || month==7 || month==8 || month==10|| month==12) && day>31) || ((month==4 || month==6 || month==9 || month==11 ) && day>30))
{
alert('��ѡ����ȷ�ĳ������ڣ�');
document.theform.d.focus()
return false;
}
if ( hour=='')
{
alert('��ѡ�����ʱ�䣡');
document.theform.hh.focus()
return false;
}

var b=document.theform.b.value;
if (b=='')
{
alert('��ѡ�������Ա�');
document.theform.b.focus()
return false;
}
var name=document.theform.name.value;
if (name=='')
{
alert('����������������');
document.theform.name.focus()
return false;
}
if (document.theform.name.value.length < 2 || document.theform.name.value.length>5)
{
alert("��������Ӧ��2-5����֮�䣡");
document.theform.name.focus();
return (false);
}
}
//-->
</script><table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <TBODY>
      <tr>
        <TD class=ttd style="PADDING-BOTTOM: 8px" vAlign=top>* ��������Դ����䣬�������֣� 
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
        <TD class=new style="PADDING-BOTTOM: 8px" vAlign=top><p>�װ���<%=request.form("name")%>��</p>
          <p> ��������<%=request.form("y")%>��<%=request.form("m")%>��<%=request.form("d")%>�գ�����<%=DateDiff("yyyy", ""&request.form("y")&""&"-"&""&request.form("m")&""&"-"&""&request.form("d")&"", date)%>�꣨�������Ͻ����ο�����</p>
          <p> ������ϱ�����ʲô�ˣ��������Ĺ�Ԫ���������գ��������Լ��㣬���Ϊ<%=sbnum3%>����ʾ���ϱ����� <%set rs=server.createobject("adodb.recordset")
sql="select * from sbwr where num='"&sbnum3&"'"
rs.open sql,conn,1,1
intro=rs("intro")
rs.close
%> <font color=#ff0000><%=intro%></font><br /><br /><a href="sm.asp?sm=7"><font color=#0000ff>[���²���]</font></a> </p></TD>
      </tr><%else%><form method="POST" action="" onSubmit="return checkbz();" name="theform">   <tr>
        <TD class=new style="PADDING-BOTTOM: 8px" vAlign=top><strong>�������գ�</strong><select size="1" name="y"><%for i=1900 to year(date())%><option value="<%=i%>" <%if nian1<>"" then%><%if i=cint(nian1) then%> selected<%end if%><%else%><%if i=year(date()) then%> selected<%end if%><%end if%>><%=i%></option><%next%></select>��<select name="m" size="1"><%for i=1 to 12%><option value="<%=i%>" <%if yue1<>"" then%><%if i=cint(yue1) then%> selected<%end if%><%else%><%if i=month(date()) then%> selected<%end if%><%end if%>><%=i%></option><%next%></select>��<select name="d" size="1"><%for i=1 to 31%><option value="<%=i%>" <%if ri1<>"" then%><%if i=cint(ri1) then%> selected<%end if%><%else%><%if i=day(date()) then%> selected<%end if%><%end if%>><%=i%></option><%next%></select>��<select size="1" name="hh"> <%for i=0 to 23%><option value="<%=i%>" <%if hh1<>"" then%><%if i=cint(hh1) then%> selected<%end if%><%end if%>><%=DiZhi(i)%><%=i%></option><%next%> </select>��<select size="1" name="fen"><option value="0">δ֪</option>
		<%for i=0 to 59%><option value="<%=i%>" <%if mm1<>"" then%><%if i=cint(mm1) then%> selected<%end if%><%end if%>><%=i%></option><%next%>
		</select>��
       </TD>
      </tr>  <tr>
        <TD class=new style="PADDING-BOTTOM: 8px" vAlign=top><strong>������</strong>
          <input onkeyup="value=value.replace(/[^\u4E00-\u9FA5]/g,'')" name="name" type="text" value="<%=xing%><%=ming%>" />
          &nbsp;
          <select size="1" name="b">
<option value="��" <%if xingbie="��" then%>selected="selected"<%end if%>>����</option>
<option value="Ů" <%if xingbie="Ů" then%>selected="selected"<%end if%>>Ů��</option>
          </select>&nbsp;<input type="hidden" name="act" value="ok" /><input name="cs" type="submit" value="��ʼ����" />
       </TD>
      </tr></form><%end if%>
    </TBODY>
  </TABLE>
