<% 
const senlontitle="�����������"
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>�����������,������������,������׼����վ_ʵ�ò�ѯ���ߴ�ȫ - Powered by Senlon!</TITLE>

</head>
<body><!--#include file="../senlon/conn.asp"-->
<!--#include file="../senlon/function.asp"-->
<!--#include file="../senlon/getuc.asp"-->
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
<%
if xing<>"" then
%>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
  <TBODY>
    <tr>
      <TD class=ttop style="PADDING-BOTTOM: 1px" vAlign=top>��ǰ������������Ϣ</TD>
    </tr>
    <tr>
      <TD class=new style="PADDING-BOTTOM: 8px" vAlign=top>������<font color="#FF0000"><%=xing&ming%></font> �Ա�<%=xingbie%>
          <%if xuexing<>"" then%>
        Ѫ�ͣ�<%=xuexing%> ��
        <%end if%>
        ����:<font color="#0000ff"><%=nian1%>��<%=yue1%>��<%=ri1%>��<%=hh1%>ʱ<%=mm1%>��</font> ����<%=userll%>�� ����:<%=sx%>&nbsp;����:<%=Constellation(nian1&"-"&yue1&"-"&ri1)%>&nbsp;<input type="button" value="�˳�����" onClick="(location='cookieclear.asp')" style="cursor:hand;COLOR: #ff0000;FONT-WEIGHT: bold;" class="button" /></TD>
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
          <input type="button" value="EQ����" onClick="(location='eq.asp')" style="cursor:hand;" class="button" /> 
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
</TABLE>
<%else%><script language="JavaScript">
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
//if (year>nowyear || year <=nowyear-10 || isNaN(year))
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
 <form method="post" action="sm.asp?sm=1" name="sm"  onSubmit="return checkbz();"> <TBODY>
       <tr>
      <TD class=ttop style="PADDING-BOTTOM: 1px" vAlign=top> �������</TD>
    </tr> <tr>
      <TD class=new style="PADDING-BOTTOM: 8px" vAlign=top><span class="12black"><b><img src="../images/help.gif" width="16" height="16">&nbsp;����˵����</b>���������������ģ����ձ������빫��������������/������ũ��������/�����������������Ѫ�͡�Ѫ�Ϳ���ѡ��������������ֵȡ�����ʱ�ֿ���ѡ����Ӱ���������Խ����</span></TD>
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
      ><%for i=1930 to year(date())%><option value="<%=i%>" <%if i=year(date()) then%> selected<%end if%>><%=i%></option><%next%>
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
      <TD align="center"  vAlign=middle class=new style="PADDING-BOTTOM: 8px">
	  &nbsp;<input type="submit" value="��ʼ����" style='cursor:hand;COLOR: #ff0099;' class="button">
</TD></tr>	
  </TBODY></form>
</TABLE>
  <%end if%>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <TBODY>
      <tr>
        <TD colspan="3"  vAlign=top class=ttop style="PADDING-BOTTOM: 1px">��ǩռ��</TD>
      </tr>
      <tr>
        <TD align="center" vAlign=top class=new style="PADDING-BOTTOM: 1px"><a href="/guanyin/" target="_blank"><img src="../images/gylq.gif" alt="������ǩ" width="176" height="89" border="0"></a></TD>
        <TD align="center" vAlign=top class=new style="PADDING-BOTTOM: 1px"><a href="/lvzu/" target="_blank"><img src="../images/yzlq.gif" alt="������ǩ" width="172" height="88" border="0"></a></TD>
        <TD align="center" vAlign=top class=new style="PADDING-BOTTOM: 1px"><a href="/huangdaxian/" target="_blank"><img src="../images/dxlq.gif" alt="�ƴ�����ǩ" width="174" height="88" border="0"></a></TD>
        </tr>
      <tr>
        <TD align="center" vAlign=top class=new style="PADDING-BOTTOM: 1px"><a href="/guandi/" target="_blank"><img src="../images/dmgdm.gif" alt="�ص���ǩ" width="174" height="88" border="0"></a></TD>
        <TD align="center" vAlign=top class=new style="PADDING-BOTTOM: 1px"><a href="/mazu/" target="_blank"><img src="../images/thlq.gif" alt="�����ǩ" width="174" height="86" border="0"></a></TD>
        <TD align="center" vAlign=top class=new style="PADDING-BOTTOM: 1px"><a href="/zgss/" target="_blank"><img src="../images/zgcz.gif" alt="�������" width="172" height="88" border="0"></a></TD>
      </tr>
    </TBODY>
  </TABLE>
  <table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1">
  <tbody>
    <tr>
      <td colspan="3"  valign="top" class="ttop" style="PADDING-BOTTOM: 1px">��������</td>
    </tr>
    <tr>
      <td width="13%" valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/ppbazi/" target="_blank"><img src="../images/pg1.gif" alt="����������������" width="80" height="80" border="0" /></a></td>
      <td width="5%" valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/ppbazi/" target="_blank"  class="red">��������</a></td>
      <td width="82%" valign="top" class="ttd" style="PADDING-BOTTOM: 1px">�������֣��ֳơ��������֡������˳������ꡢ�¡��ա�ʱΪ������������֮��֧Ϊ���֡��������ɡ���ɡ���֧��Ϊ���������¸ɡ���֧��Ϊ���������ոɡ���֧��Ϊ��������ʱ�ɡ�ʱ֧��Ϊʱ������������ϣ����������˸���֧����ɣ����֣����ʳơ��������򡰰��֡�ȫ�ơ��������֡���<a href="/ppbazi/" target="_blank" class="red">�������>></a></td>
    </tr>
    <tr>
      <td width="13%" valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/pp6r/" target="_blank"><img src="../images/pg2.gif" alt="������������ϵͳ" width="80" height="80" border="0"></a></td>
      <td width="5%" valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a  class="red" href="/pp6r/" target="_blank">������</a></td>
      <td width="82%" valign="top" class="ttd" style="PADDING-BOTTOM: 1px">����������������ռ�����׵�һ��������������ݼס�̫�Һϳ���ʽ��������ˮΪ�ף�ʮ����У��ɡ���ֱ�Ϊ��ˮ����ˮ������ȡ������ʮ�������������������ꡢ���硢�ɳ������������ӡ����磩����Ϊ���ɡ���������ʮ�ĿΣ��Կ��и�֧�����̡����������ת�����̺�ó���ֵ�ĸ�֧��ʱ�����������ס�<a href="/pp6r/" target="_blank" class="red">�������>></a></td>
    </tr>
	<tr>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/ppjkj/" target="_blank"><img src="../images/pg7.gif" alt="��ھ���������ϵͳ" width="80" height="80" border="0"></a></td>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/ppjkj/" class="red" target="_blank">��ھ�</a></td>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px">�ᵽ��ھ��Ͳ����Բ�˵�����ɣ���Ϊ��ھ���ȫ�ƾ���"��������ν�ھ�"����ˣ���֪��ھ���Դ�ڴ�������̥�γɵ��µ�һ��Ԥ���ѧ����ھ����������������䴫�˺���пڿ��ഫ���������֣�һֱ��Ϊ������֪���������ڿ��Կ������йؽ�ھ�����������ľ�������ʱ�ڵġ���������ν�ھ�����<a href="/ppjkj/" target="_blank" class="red">�������>></a></td>
    </tr>
    <tr>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/pp6y/" target="_blank"><img src="../images/pg3.gif" alt="�ɼ���س��������ϵͳ" width="80" height="80" border="0"></a></td>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/pp6y/" class="red" target="_blank">��س����</a></td>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px">��س���ҹ���ͳԤ�ⷽ����һ�֣������Ͱ��ֵ�ͬ�ۡ�����س����������ռ��Ľз������ı������С�������Ԥ�⡱���ɼ�Ԥ�⡱������Ԥ�⡱�ȣ��鱾�϶��֮Ϊ����Ԥ�⡣<a href="/pp6y/" target="_blank" class="red">�������>></a></td>
    </tr>
    <tr>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/ppqmdj/" target="_blank"><img src="../images/pg4.gif" alt="���Ŷݼ���������" width="80" height="80" border="0"></a></td>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/ppqmdj/" class="red" target="_blank">���Ŷݼ�</a></td>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px">���Ŷݼף�ԭ�����й����ϵ�һ���飬������������Ϊ��һ��ռ���õ��飬���е�˵����˵�����Ŷݼס����ҹ��Ŵ�������ͬ����Ȼ�������У��������ڹ۲졢������֤���ܽ������һ�Ŵ�ͳ����Ļ��Ų������е�˵�����Ŷݼס�������Ĺ�����<a href="/ppqmdj/" target="_blank" class="red">�������>></a></td>
    </tr>
    <tr>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/ppxk/" target="_blank"><img src="../images/pg5.gif" alt="���շ�����������" width="80" height="80" border="0"></a></td>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/ppxk/" class="red" target="_blank">���շ���</a></td>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px">������ʱ�䣬�վ��ǿռ䣬����ѧ����һ���о�ʱ��Ϳռ�Ϊ���츣��ѧ�ʡ�����ѧ��������ϵ���ɣ���ͼ�����飬���ԡ���ͼ����Ҫ������ǣ�һ��Ϊˮ������Ϊ������Ϊľ���ľ�Ϊ����ʮΪ������ͼΪ���졣�������Ҫ������ǣ�������һ���������ߣ�����Ϊ�磬����Ϊ�㡣����Ϊ���졣������������ԣ��к�����ԣ��������ͼ�����������顣�Ⱥ�����Եķ�λһ��Ҫ���Ρ�<a href="/ppxk/" target="_blank" class="red">�������>></a></td>
    </tr>
    <tr>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/ppziwei/" target="_blank"><img src="../images/pg6.gif" alt="��΢������������ϵͳ" width="80" height="80" border="0"></a></a></td>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/ppziwei/" class="red" target="_blank">��΢����</a></td>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px">��΢���������й���ͳ����ѧ������Ҫ��֧��֮һ���������˳������ꡢ�¡��ա�ʱȷ��ʮ������λ�ã��������̣���ϸ�������Ⱥ��ϣ�ǣϵ������س����Ԥ��һ���˵���Ա���̡����׻����ġ�<a href="/ppziwei/" target="_blank" class="red">�������>></a></td>
    </tr>
    <tr>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/taiyang/" target="_blank"><img src="../images/zty.jpg" alt="��̫��ʱ���ѯ����" width="80" height="80" border="0" /></a></td>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/taiyang/" class="red" target="_blank">��̫��ʱ��</a></td>
      <td valign="top" class="ttd" style="PADDING-BOTTOM: 1px">ʱ������Ӧ�ð���ʲô��׼ȷ���أ��ǰ�����ʱ��ȷ�������ǰ�����ʱ��ȷ������ʵΨһ�ı�׼����̫������ڵ�����ӽǣ�������ѧ����ν����̫��ʱ��������̫��ʱ��ƽ̫��ʱ+��ƽ̫��ʱ�������ʱ�����밴��������ص㣬��������ص�ƽ̫��ʱ���ٸ���ƽ̫��ʱ�������̫��ʱΪ׼�����ܼ򵥵����ñ���ʱ�䡣<a href="/taiyang/" target="_blank" class="red">�������>></a></td>
    </tr>
    <tr>
      <td valign="top" class="new" style="PADDING-BOTTOM: 1px"><a href="/jingdu/" target="_blank"><img src="../images/dqjw.jpg" width="80" height="80" border="0" /></a></td>
      <td valign="top" class="new" style="PADDING-BOTTOM: 1px"><a href="/jingdu/" class="red" target="_blank">��������</a></td>
      <td valign="top" class="new" style="PADDING-BOTTOM: 1px">���ȷ�ָ��������ϵ�������ꡣ����Ϊ��������һ����������������0�Ⱦ�������ƽ��ļнǡ��������ϵĵ����ڸ�Ȧ���������ԭ�����ڸ�Ȧ�ĽǾ�������ʾ��ͨ����ָ��������ľ��ȡ�Ϊ�����ֵ����ϵ�ÿһ�����������Ǹ����߱�ע�˶���������Ǿ��ȡ�ʵ���Ͼ�����������������ƽ��֮��ļнǡ�<a href="/jingdu/" target="_blank" class="red">�������>></a></td>
    </tr>
  </tbody>
</table>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1">
    <TBODY>
      <tr>
        <TD colspan="3"  vAlign=top class=ttop style="PADDING-BOTTOM: 1px">����Ԥ��</TD>
      </tr>
      <tr>
         <td width="13%" valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/hmjx/"><img src="../images/c4.gif" alt="�ֻ�����Ԥ��" width="80" height="80" border="0"></a></TD>
         <td width="5%" valign="top" class="ttd" style="PADDING-BOTTOM: 1px"><a href="/hmjx/" class="red" target="_blank">���뼪��</a></TD>
         <td width="82%" valign="top" class="ttd" style="PADDING-BOTTOM: 1px">��Щ�˻��ʣ���������Ϊʲô��Ӱ��һ���˵����ƣ���ʵ������ˮ����լ��Ӱ���������˵�������һ���ġ���Ȼ��ֻ��һ�����룬����������������ϢϢ��أ�Ҳ�������������˵Ĺ�ͨ������<A href="/hmjx/" target="_blank" class="red">�������>></A></TD>
      </tr>
      <tr>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px"><a href="/ytyc/"><img src="../images/c5.gif" alt="����Ԥ�� " width="80" height="80" border="0"></a></TD>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px"><a href="/ytyc/" class="red" target="_blank">����Ԥ�� </a></TD>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px">��Ƥ�������ǻ��Ǹ���������ϲ����֪�����������㣬����������������߶���ֻ������ϰ�Ҫ���㡰̸���������𼱱𼱣��й����������ס��������⡢�������׵ȵ�һ�У��������嶼�������Ŷ��<A href="/ytyc/" target="_blank" class="red">�������>></A></TD>
      </tr>

      <tr>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px"><a href="/mryc/"><img src="../images/c6.gif" alt="����Ԥ��" width="80" height="80" border="0"></a></TD>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px"><a href="/mryc/" class="red" target="_blank">����Ԥ��</a></TD>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px">��������գ��ȵĲ����ˣ����Ǹ�ð����û���գ�Ϊ������죿���ܲ�֪��������Ԥ�⣬��������������,���ȼ���,���ȵ�ԭ�򣬿�����Ŷ��<A href="/mryc/" target="_blank" class="red">�������>></A></TD>
      </tr>
      <tr>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px"><a href="/ptyc/"><img src="../images/c7.gif" alt="����Ԥ��" width="80" height="80" border="0"></a></TD>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px"><a href="/ptyc/" class="red" target="_blank">����Ԥ��</a></TD>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px">�����һ��,����,����߶��������ܵ���ʱ�ų���Ӱ�졣����Ԥ�⣬������������Ԥ�ף������������ɣ�<A href="/ptyc/" target="_blank" class="red">�������>></A></TD>
      </tr>
      <tr>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px"><a href="/xjyc/"><img src="../images/c8.gif" alt="�ľ�Ԥ��" width="80" height="80" border="0"></a></TD>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px"><a href="/xjyc/" class="red" target="_blank">�ľ�Ԥ��</a></TD>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px">���ǻ�Ī�����ľ���Ϊ���������ܷ��겻���������Լ�Ҫ�������������ľ�Ԥ����ʲô���ľ���ԭ���ľ���������ʲô�����������ľ�������ô˵�ɣ�<A href="/xjyc/" target="_blank" class="red">�������>></A></TD>
      </tr>
      <tr>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px"><span class="ttd" style="PADDING-BOTTOM: 1px"><a href="/emyc/"><img src="../images/c9.jpg" alt="����Ԥ��" width="80" height="80" border="0"></a></span></TD>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px"><a href="/emyc/" class="red" target="_blank">����Ԥ��</a></TD>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px">����������С�С�۷䡱�ڡ������ˡ�������ʱ����������ҡ���֡��������Ƕ�����������Ϊʲô������أ���Ȼ�п�ѧ�Ľ��ͣ�Ҳ��һЩ�ǿ�ѧ�����ܸ�����Ĳ���Ŷ��ͨ������ռ�����������ֶ���Ԥ����㽫���ᷢ��ʲô����Ŷ��<a href="/emyc/" target="_blank" class="red">�������>></A></TD>
      </tr>	  
	  <tr>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px"><a href="/sanshishu/"><img src="../images/sss.gif" alt="����Ԥ��" width="80" height="80" border="0"></a></TD>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px"><a href="/sanshishu/" target="_blank"  class="red">����Ԥ��</a></TD>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px">�������Ǵ�˵��һ�����Կ�����ǰ���������ͺ������飬���������˵�һ�ж��������������ƪ��ר�Ŵ����澫ѡ�����������㵽��Ĳ��������������͵ģ���˵��׼��Ŷ�����԰ɣ�<a href="/sanshishu/" target="_blank" class="red">�������>></a></TD>
      </tr>
      <tr>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px"><a href="/snsn/"><img src="../images/c1.gif" alt="������Ů" width="80" height="80" border="0"></a></TD>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px"><a  class="red" href="/snsn/" target="_blank">������Ů</a></TD>
        <TD vAlign=top class=ttd style="PADDING-BOTTOM: 1px">��Ԥ��һ�����ǻ������б�������Ů�����𣿴�ͳ�����裬����Ԥ��������Ů�����������һ���幬��ص�������ŮԤ���Ů����������ũ������(����ʵ�����1��)��,���������·ݼ������е��Ǹ��·ݣ���Ȼ������ũ���㡣<a href="/snsn/" target="_blank" class="red">�������>></a></TD>
      </tr>
      <tr>
        <TD vAlign=top class=new style="PADDING-BOTTOM: 1px"><a href="/zwsm/"><img src="../images/c3.gif" alt="ָ������" width="80" height="80" border="0"></a></TD>
        <TD vAlign=top class=new style="PADDING-BOTTOM: 1px"><a href="/zwsm/" class="red" target="_blank">ָ������</a></TD>
        <TD vAlign=top class=new style="PADDING-BOTTOM: 1px">�ݶ����о����棬�˵��Ը������������ġ�������ֵ��ǣ��˵�ָ��Ҳ����������ġ�����������о�������ͨ��ָ�ƿ��Ը����о��˵��Ը�<A href="/zwsm/" target="_blank" class="red">�������>></A></TD>
      </tr>
    </TBODY>
  </TABLE>
      <!--#INCLUDE FILE="../dq.asp"-->
	</div>
    </td>
    <td valign="top"><!--#INCLUDE FILE="../right.asp"--></td>
  </tr>
</table>
<!--#include file="../foot.asp"-->
</body></html>
