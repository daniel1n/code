<% 
const senlontitle="��������"
%><!--#include file="../senlon/conn.asp"-->
<!--#include file="../senlon/function.asp"-->
<!--#include file="../senlon/cookies.asp"-->
<%
writecookies()
%>
<!--#include file="../senlon/getuc.asp"-->
<!--#include file="../senlon/bzi.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>��������,���ֲ���,�ո�����,�ƹ�����,��������,�������,�ϱ�Ϊ��,������Դ</TITLE>
<style type="text/css">
<!--
.fontcolor1 {color:#0000ff;}
-->
</style>

</head>
<body>
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
<%
if xing<>"" then
userbirthday=nian1&"-"&yue1&"-"&ri1
userll=DateDiff("yyyy", userbirthday, date)
%>
  <table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
  <TBODY>    
      <tr>
        <TD class=ttop style="PADDING-BOTTOM: 1px" vAlign=top>��ǰ��������Ϣ</TD>
      </tr>
    <tr>
      <TD class=new style="PADDING-BOTTOM: 8px" vAlign=top>������<font color="#FF0000"><%=xing&ming%></font> �Ա�<%=xingbie%> <%if xuexing<>"" then%> Ѫ�ͣ�<%=xuexing%> ��
      <%end if%>         ����:<font color="#0000ff"><%=nian1%>��<%=yue1%>��<%=ri1%>��<%=hh1%>ʱ<%=mm1%>��</font> ����<%=userll%>�� ����:<%=sx%>&nbsp;����:<%=Constellation(nian1&"-"&yue1&"-"&ri1)%>&nbsp;<input type="button" value="�˳�����" onClick="(location='cookieclear.asp')" style="cursor:hand;COLOR: #ff0000;FONT-WEIGHT: bold;" class="button" /></TD>
    </tr>
    <tr>
      <TD align="center" vAlign=top class=new style="PADDING-BOTTOM: 8px"><input name="button" type="button" class="button" style='cursor:hand;' onClick="(location='sm.asp?sm=1')" value="��������">
          <input name="button" type="button" class="button" style="cursor:hand;" onClick="(location='sm.asp?sm=2')" value="���ֲ���" />
          <input name="button" type="button" class="button" style="cursor:hand;" onClick="(location='sm.asp?sm=3')" value="�ո�����" />
          <input name="button" type="button" class="button" style="cursor:hand;" onClick="(location='sm.asp?sm=4')" value="�ƹ�����" />
          <input name="button" type="button" class="button" style="cursor:hand;" onClick="(location='sm.asp?sm=5')" value="��������" />
          <input name="button" type="button" class="button" style="cursor:hand;" onClick="(location='sm.asp?sm=6')" value="�������" />
          <input name="button" type="button" class="button" style="cursor:hand;" onClick="(location='sm.asp?sm=7')" value="�ϱ�Ϊ��" />
          <input name="button" type="button" class="button" style="cursor:hand;" onClick="(location='sm.asp?sm=8')" value="������Դ" />
          <input type="button" value="���屣��" style='cursor:hand;' onClick="(location='baojian.asp')" class="button">
          <input type="button" value="EQ ����" onClick="(location='eq.asp')" style="cursor:hand;" class="button" /> 
          <input type="button" value="�����" onClick="(location='wu.asp')" class="button" />
          <input type="button" value="IQ ����" onClick="(location='iq.asp')" style="cursor:hand;" class="button" />
          <input type="button" value="ʧ������" onClick="(location='shibai.asp')" style="cursor:hand;" class="button" />
          <input type="button" value="��������" style='cursor:hand;' onClick="(location='mingren.asp')" class="button">
          <input type="button" value="�����Ը�" onClick="(location='astro.asp?flag=5&xiao=<%=sx%>')" style="cursor:hand;" class="button" />
          <input type="button" value="��������" onClick="(location='astro.asp?flag=8&flag1=������&m=<%=yue1%>&d=<%=ri1%>')" style="cursor:hand;" class="button" />
          <input type="button" value="���ջ���" onClick="(location='astro.asp?flag=8&flag1=���ջ�&m=<%=yue1%>&d=<%=ri1%>')" style="cursor:hand;" class="button" />
          <input type="button" value="Ѫ���Ը�" onClick="(location='astro.asp?flag=6&xuexing=<%=xuexing%>')" style="cursor:hand;" class="button" />
    </tr>
  </TBODY>
</TABLE>     <%
		if request("sm")=1 then%><!--#include file="bz.asp"--> <%elseif request("sm")=2 then%><!--#include file="bzcs.asp"--><%elseif request("sm")=3 then%><!--#include file="nm.asp"--><%elseif request("sm")=4 then%><!--#include file="cg.asp"-->
 <%elseif request("sm")=5 then%><!--#include file="xm.asp"--><%elseif request("sm")=6 then%><!--#include file="xmpd.asp"--><%elseif request("sm")=7 then%><!--#include file="sbwr.asp"--><%elseif request("sm")=8 then%><!--#include file="x.asp"--><%end if%><%else%>
<script language="JavaScript">
<!--
function checkbz()
{
var year=document.sm.nian.value;
var month=document.sm.yue.value;
var day=document.sm.ri.value;
var hour=document.sm.hh.value;
var now=new Date();
var nowyear=now.getYear();
var nowmonth=now.getMonth();
if (year=='')
{
alert('��ѡ�������ݣ�');
document.sm.nian.focus()
return false;
}
//if (year>nowyear || year <=nowyear-100 || isNaN(year))
if (year=='')
{
alert('��ѡ����ȷ�ĳ�����ݣ�');
document.sm.nian.focus()
return false;
}
if ( month=='')
{
alert('��ѡ������·ݣ�');
document.sm.yue.focus()
return false;
}
if (day=='')
{
alert('��ѡ��������ڣ�');
document.sm.ri.focus()
return false;
}
if ((month==2 && day>29) || ((month==1 || month==3 || month==5 || month==7 || month==8 || month==10|| month==12) && day>31) || ((month==4 || month==6 || month==9 || month==11 ) && day>30))
{
alert('��ѡ����ȷ�ĳ������ڣ�');
document.sm.ri.focus()
return false;
}
if ( hour=='')
{
alert('��ѡ�����ʱ�䣡');
document.sm.hh.focus()
return false;
}
while(document.sm.xing.value.indexOf(" ")!=-1){
document.sm.xing.value=document.sm.xing.value.replace(" ","");
}
while(document.sm.xing.value.indexOf("	")!=-1){
document.sm.xing.value=document.sm.xing.value.replace("	","");
}
if (document.sm.xing.value=='')
{
alert('�������������ϣ�');
document.sm.xing.focus()
return false;
}
if (document.sm.xing.value.length < 1 || document.sm.xing.value.length>2)
{
alert("��������Ӧ��1-2����֮�䣡");
document.sm.xing.focus();
return (false);
}

while(document.sm.ming.value.indexOf(" ")!=-1){
document.sm.ming.value=document.sm.ming.value.replace(" ","");
}
while(document.sm.ming.value.indexOf("	")!=-1){
document.sm.ming.value=document.sm.ming.value.replace("	","");
}
if (document.sm.ming.value=='')
{
alert('�������������֣�');
document.sm.ming.focus()
return false;
}
if (document.sm.ming.value.length < 1 || document.sm.ming.value.length>2)
{
alert("��������Ӧ��1-2����֮�䣡");
document.sm.ming.focus();
return (false);
}
var b=document.sm.xingbie.value;
if (b=='')
{
alert('��ѡ�������Ա�');
document.sm.xingbie.focus()
return false;
}
}
//-->
</script>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
 <form method="post" action="sm.asp?sm=<%=request("sm")%>" name="sm"  onSubmit="return checkbz();"> <TBODY>
       <tr>
      <TD class=ttop style="PADDING-BOTTOM: 1px" vAlign=top> �����������̿�ʼ�������</TD>
    </tr> <tr>
      <TD class=new style="PADDING-BOTTOM: 8px" vAlign=top><span class="12black"><b><img src="../images/help.gif" width="16" height="16">&nbsp;����˵����</b>���������������ģ����ձ������빫��������������/������ũ��������/���������������Ѫ�ͣ�Ѫ�Ϳ���ѡ��������������ֵȣ�����ʱ�ֿ���ѡ����Ӱ���������Խ��������ϵͳ���������ʮ��ǿ����������ܣ�</span></TD>
    </tr>
    <tr>
      <TD class=new style="PADDING-BOTTOM: 8px" vAlign=top><div align="center"><a title="�����ֻ֪�����յ�ũ�����ڣ���Ҫ�����������ȥ��ѯ��������" style="CURSOR: hand" href="/wannianli/" target="_blank"><font color="red">[ֻ֪��ũ�����˲鹫��]</font></a>&nbsp;<a title="��֪������ʱ����ô��" style="CURSOR: hand" href="./htm_nobirth.asp" target="_blank"><font color="red">[��֪������ʱ����ô��]</font></a>&nbsp;<a title="����˴��˽�����֪ʶ" style="CURSOR: hand" href="./htm_smzs.asp" target="_blank"><font color="red">[����˴��˽�����֪ʶ]</font></a></div></TD>
    </tr>
    <tr>
      <TD align="center" vAlign=top class=new style="PADDING-BOTTOM: 8px">�գ�<input type="txt" name="xing" size="4" value="" onKeypress="if ((event.keyCode != 13 && event.keyCode < 160)) event.returnValue = false;">
  	����<input type="txt" name="ming" size="4" value="" onKeypress="if ((event.keyCode != 13 && event.keyCode < 160)) event.returnValue = false;">
  	<select name="xingbie" size="1" style="font-size: 9pt">
	<option value="" selected>�Ա�</option>
	<option value="��">��</option>
	<option value="Ů">Ů</option>
  	</select>
  	<select name="xuexing" size="1" style="font-size: 9pt">
  	<option value="">Ѫ��</option>
  	<option value="A">A��</option>
  	<option value="B">B��</option>
  	<option value="O">O��</option>
  	<option value="AB">AB��</option>
  	</select>
  	��������:
       <select name="nian" size="1" style="font-size: 9pt">
      ><%for i=1900 to year(date())%><option value="<%=i%>" <%if i=year(date()) then%> selected<%end if%>><%=i%></option><%next%>
     </select>
     ��
     <select size="1" name="yue" style="font-size: 9pt">
      <%for i=1 to 12%><option value="<%=i%>"<%if i=month(date()) then%> selected<%end if%>><%=i%></option><%next%>
     </select>
     ��
     <select size="1" name="ri" style="font-size: 9pt">
      <%for i=1 to 31%><option value="<%=i%>" <%if i=day(date()) then%> selected<%end if%>><%=i%></option><%next%>
     </select>
     ��
     <select size="1" name="hh" style="font-size: 9pt">
       <%for i=0 to 23%><option value="<%=i%>" ><%=DiZhi(i)%>&nbsp;<%=i%></option><%next%> 
     </select>
     ��
     <select size="1" name="mm" style="font-size: 9pt">
       <option value="0">δ֪</option>
		<%for i=0 to 59%><option value="<%=i%>"><%=i%></option><%next%>
     </select>
     �� </TD>
    </tr>
    <tr>
      <TD align="center" vAlign=top class=new style="PADDING-BOTTOM: 8px">&nbsp;<input type="submit" value="��������" style='cursor:hand;COLOR: #ff0000;' class="button"> </TD>
    </tr>
  </TBODY></form>
</TABLE>
<%end if%>
      <!--#INCLUDE FILE="../dq.asp"-->
	</div>
    </td>
    <td valign="top"><!--#INCLUDE FILE="../right.asp"--></td>
  </tr>
</table>
<!--#include file="../foot.asp"-->
</body></html>
