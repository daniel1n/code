<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="senlon/getuc.asp"-->
<!--#include file="senlon/function.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>2016������ȫ����Դ��</title>
</head>

<body>
<!--#INCLUDE FILE="top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top"><div id="mid_left">
      <p><span class="ren">����˵����</span><br />
      ����������������,���ձ������빫�������������Ѫ�͡�Ѫ�Ϳ���ѡ,������������ֵȡ�����ʱ�ֿ���ѡ,��Ӱ���������Խ��.<span class="ren">��&gt;</span><A href style='cursor:hand;' onclick="window.open('http://sm.aa963.com/hdjr/index.htm','cidunongli','left=120,top=100,width=800,height=550,scrollbars=no,resizable=no,status=no')" title="�����ֻ֪�����յ�ũ�����ڣ���Ҫ�����������ȥ��ѯ��������">[����/ũ������ת��</a>] [<a href="/sm/htm_nobirth.asp" target="_blank">��֪������ʱ����ô��</a>]</p>
      <hr class="shplink" />
      <form method="post" action="/sm/sm.asp?sm=1" name="sm"  onSubmit="return checkbz();"> 
        <p>�գ�
        <input type="txt" name="xing" size="4" value="" onKeypress="if ((event.keyCode != 13 && event.keyCode < 160)) event.returnValue = false;">
        ����<input type="txt" name="ming" size="4" value="" onKeypress="if ((event.keyCode != 13 && event.keyCode < 160)) event.returnValue = false;">
        </p>
      <p>
        ����ѡ��
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
        </p>
      <p>��������:
  <select name="nian" size="1" style="font-size: 9pt">
    ><%for i=1930 to year(date())%>
    <option value="<%=i%>" <%if i=year(date()) then%> selected<%end if%>><%=i%></option>
    <%next%>
  </select>
        ��
        <select size="1" name="yue" style="font-size: 9pt">
          <%for i=1 to 12%>
          <option value="<%=i%>"<%if i=month(date()) then%> selected<%end if%>><%=i%></option>
          <%next%>
        </select>
        ��
        <select size="1" name="ri" style="font-size: 9pt">
          <%for i=1 to 31%>
          <option value="<%=i%>" <%if i=day(date()) then%> selected<%end if%>><%=i%></option>
          <%next%>
        </select>
        ��
        <select size="1" name="hh" style="font-size: 9pt">
          <%for i=0 to 23%>
          <option value="<%=i%>" ><%=DiZhi(i)%>&nbsp;<%=i%></option>
          <%next%> 
        </select>
        ��
        <select size="1" name="mm" style="font-size: 9pt">
          <option value="0">δ֪</option>
          <%for i=0 to 59%>
          <option value="<%=i%>"><%=i%></option>
          <%next%>
        </select>
        ��   ��
        <input type="submit" value="��ʼ����" style='cursor:hand;COLOR: #ff0099;' class="anliu">
      </p>
  </form>
      <hr class="shplink" />
      <table width="100%" border="0" cellspacing="0" cellpadding="6">
        <tr>
          <td align="center"><a href="/ppbazi/index.asp"><img src="/images/pg1.gif" width="80" height="80" border="0" /></a></td>
          <td align="center"><a href="/ppziwei/index.asp"><img src="/images/pg6.gif" width="80" height="80" border="0" /></a></td>
          <td align="center"><a href="/ppqmdj/index.asp"><img src="/images/pg4.gif" width="80" height="80" border="0" /></a></td>
          <td align="center"><a href="/ppjkj/index.asp"><img src="/images/pg7.gif" width="80" height="80" border="0" /></a></td>
          <td align="center"><a href="/pp6y/index.asp"><img src="/images/pg3.gif" width="80" height="80" border="0" /></a></td>
          <td align="center"><a href="/pp6r/index.asp"><img src="/images/pg2.gif" width="80" height="80" border="0" /></a></td>
          <td align="center"><a href="/ppxk/index.asp"><img src="/images/pg5.gif" width="80" height="80" border="0" /></a></td>
          </tr>
        <tr>
          <td align="center"><a href="/ppbazi/index.asp">��������</a></td>
          <td align="center"><a href="/ppziwei/index.asp">��΢����</a></td>
          <td align="center"><a href="/ppqmdj/index.asp">���Ŷݼ�</a></td>
          <td align="center"><a href="/ppjkj/index.asp">��ھ�</a></td>
          <td align="center"><a href="/pp6y/index.asp">��س����</a></td>
          <td align="center"><a href="/pp6r/index.asp">������</a></td>
          <td align="center"><a href="/ppxk/index.asp">���շ���</a></td>
          </tr>
      </table>
      <p><%=(adgl.Fields.Item("addaima05").Value)%></p>
      <table width="100%" border="0" cellspacing="0" cellpadding="6">
        <tr>
          <td align="center"><a href="/sanshishu/index.asp"><img src="/images/sss.gif" width="80" height="80" border="0" /></a></td>
          <td align="center"><a href="/snsn/index.asp"><img src="/images/c1.gif" width="80" height="80" border="0" /></a></td>
          <td align="center"><a href="/shouji/index.asp"><img src="/images/c4.gif" width="80" height="80" border="0" /></a></td>
          <td align="center"><a href="/zwsm/index.asp"><img src="/images/c3.gif" width="80" height="80" border="0" /></a></td>
          <td align="center"><a href="/ytyc/index.asp"><img src="/images/c5.gif" width="80" height="80" border="0" /></a></td>
          <td align="center"><a href="/ptyc/index.asp"><img src="/images/c7.gif" width="80" height="80" border="0" /></a></td>
          <td align="center"><a href="/emyc/index.asp"><img src="/images/c9.jpg" width="80" height="80" border="0" /></a></td>
        </tr>
        <tr>
          <td align="center"><a href="/sanshishu/index.asp">Ԥ�����</a></td>
          <td align="center"><a href="/snsn/index.asp">������Ů</a></td>
          <td align="center"><a href="/shouji/index.asp">���뼪��</a></td>
          <td align="center"><a href="/zwsm/index.asp">ָ������</a></td>
          <td align="center"><a href="/ytyc/index.asp">����Ԥ��</a></td>
          <td align="center"><a href="/ptyc/index.asp">����Ԥ��</a></td>
          <td align="center"><a href="/emyc/index.asp">����Ԥ��</a></td>
        </tr>
    </table>
      <p><%=(adgl.Fields.Item("addaima06").Value)%></p>
    </div><!--#INCLUDE FILE="dq.asp"--></td>
    <td valign="top"><!--#INCLUDE FILE="right.asp"--></td>
  </tr>
</table>
<!--#INCLUDE FILE="foot.asp"-->
</body>
</html>
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