<% 
const senlontitle="�����˳�"
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>�����˳̲�ѯ,ʮ����Ф���,ʮ���������_ʵ�ò�ѯ���ߴ�ȫ - Powered by Senlon!</TITLE>

</head>
<body>
<!--#include file="../senlon/conn.asp"-->
<!--#include file="../senlon/function.asp"-->
<!--#include file="../senlon/cookies.asp"-->
<%
writecookies()
%>
<!--#include file="../senlon/getuc.asp"-->
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <TBODY>
      <tr>
        <TD height='35' align="left" class=new style="PADDING-BOTTOM: 1px"><%
if xing<>"" then
myxz=Constellation(nian1&"-"&yue1&"-"&ri1)
%><%=xing&ming%>(<%=nian1%>-<%=yue1%>-<%=ri1%>) �ҵ�������<font color="#FF0000"><%=myxz%></font><a href="../sm/astro.asp?flag=4&astro=<%=myxz%>">[����]</a> �ҵ�����:<font color="#FF0000"><%=sx%></font><a href="../sm/astro.asp?flag=5&xiao=<%=sx%>">[�Ը�]</a> �ҵ�Ѫ��:<%=xuexing%><a href="../sm/astro.asp?flag=6&xuexing=<%=xuexing%>">[���]</a> &nbsp;<input type="button" style="color:#FF0000" value="��������" onClick="(location='../sm/')" >&nbsp;<input type="button" style="color:#FF0000" value="�˳��ز�" onClick="(location='../sm/cookieclear.asp')" ><%else%>
    <form action="./" method="post"><select name="y" size="1" id="y" style="font-size: 9pt"> 
            <%for i=1920 to year(date())%>
            <option value="<%=i%>" <%if i=1980 then%> selected<%end if%>><%=i%></option>
            <%next%>
          </select>
��
<select name="m" size="1" id="m" style="font-size: 9pt">
  <%for i=1 to 12%>
  <option value="<%=i%>"<%if i=month(date()) then%> selected<%end if%>><%=i%></option>
  <%next%>
</select>
��
<select name="d" size="1" id="d" style="font-size: 9pt">
  <%for i=1 to 31%>
  <option value="<%=i%>" <%if i=day(date()) then%> selected<%end if%>><%=i%></option>
  <%next%>
</select>
��(<a href="/wannianli/" target="_blank"><font color="#FF0000">����/����ת��</font></a>)
<input name="Input2" type="submit" value="��ѯ�ҵ�����" class="bot01"   /><input name="act" type="hidden" value="xzcx"><%if request("act")="xzcx" then
%>&nbsp;&nbsp;&nbsp;���������ǣ�<font color="#FF0000"><%=Constellation(""&request("y")&""&"-"&""&request("m")&""&"-"&""&request("d")&"")%></font></form>
<%end if%><%end if%></form></TD>
      </tr>
    </TBODY>
  </TABLE>
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <tbody>
      <tr>
        <td colspan="2" class="ttop">Ѫ��/�������:</td>
      </tr>
      <tr>
        <td width="49%" class="new"><form action=xueread.asp method=post target=_blank>
<input type=hidden name=flag value=6>
<select name="xuexing" style="width:120px;">
	<option value=O>OѪ�����</option>
		<option value=A>AѪ�����</option>
		<option value=B>BѪ�����</option>
		<option value=AB>ABѪ�����</option>
	</select>

 
	<input name="" type="submit" value="�鿴" class="bot05" /></p>
</form></td>
        <td width="51%" class="new"><form action=senread.asp method=post target=_blank>
<input type=hidden name=flag value=7>
	<select name="astro">
<option>������</option>		<option>��ţ��</option>		<option>˫����</option>

<option>��з��</option>		<option>ʨ����</option>		<option>��Ů��</option>
<option>�����</option>		<option>��Ы��</option>		<option>������</option>
<option>ħ����</option>		<option>ˮƿ��</option>		<option>˫����</option>
	</select>
 
	<select name="blood" class="sele01">
<option value="O">O��</option><option value="A">A��</option><option value="B">B��</option><option value="AB">AB��</option>
	</select>
	<input name="" type="submit" value="����+Ѫ��"></p>
	</form></td>
      </tr>
	  <tr>
        <td colspan="2" class="new"><a href="xueread.asp?flag=6&xuexing=o" target="_blank"><b>O��Ѫ</b></a>��<a href="senread.asp?flag=7&blood=o&astro=������" target="_blank">������</a>&nbsp;<a href="senread.asp?flag=7&blood=o&astro=��ţ��" target="_blank">��ţ��</a>&nbsp;<a href="senread.asp?flag=7&blood=o&astro=˫����" target="_blank">˫����</a>&nbsp;<a href="senread.asp?flag=7&blood=o&astro=��з��" target="_blank">��з��</a>&nbsp;<a href="senread.asp?flag=7&blood=o&astro=ʨ����" target="_blank">ʨ����</a>&nbsp;<a href="senread.asp?flag=7&blood=o&astro=��Ů��" target="_blank">��Ů��</a>&nbsp;<a href="senread.asp?flag=7&blood=o&astro=�����" target="_blank">�����</a>&nbsp;<a href="senread.asp?flag=7&blood=o&astro=��Ы��" target="_blank">��Ы��</a>&nbsp;<a href="senread.asp?flag=7&blood=o&astro=������" target="_blank">������</a>&nbsp;<a href="senread.asp?flag=7&blood=o&astro=ħ����" target="_blank">ħ����</a>&nbsp;<a href="senread.asp?flag=7&blood=o&astro=ˮƿ��" target="_blank">ˮƿ��</a>&nbsp;<a href="senread.asp?flag=7&blood=o&astro=˫����" target="_blank">˫����</a></td>
      </tr>
	  <tr>
        <td colspan="2" class="new"><a href="xueread.asp?flag=6&xuexing=a" target="_blank"><b>A��Ѫ</b></a>��<a href="senread.asp?flag=7&blood=a&astro=������" target="_blank">������</a>&nbsp;<a href="senread.asp?flag=7&blood=a&astro=��ţ��" target="_blank">��ţ��</a>&nbsp;<a href="senread.asp?flag=7&blood=a&astro=˫����" target="_blank">˫����</a>&nbsp;<a href="senread.asp?flag=7&blood=a&astro=��з��" target="_blank">��з��</a>&nbsp;<a href="senread.asp?flag=7&blood=a&astro=ʨ����" target="_blank">ʨ����</a>&nbsp;<a href="senread.asp?flag=7&blood=a&astro=��Ů��" target="_blank">��Ů��</a>&nbsp;<a href="senread.asp?flag=7&blood=a&astro=�����" target="_blank">�����</a>&nbsp;<a href="senread.asp?flag=7&blood=a&astro=��Ы��" target="_blank">��Ы��</a>&nbsp;<a href="senread.asp?flag=7&blood=a&astro=������" target="_blank">������</a>&nbsp;<a href="senread.asp?flag=7&blood=a&astro=ħ����" target="_blank">ħ����</a>&nbsp;<a href="senread.asp?flag=7&blood=a&astro=ˮƿ��" target="_blank">ˮƿ��</a>&nbsp;<a href="senread.asp?flag=7&blood=a&astro=˫����" target="_blank">˫����</a></td>
      </tr>
	  <tr>
        <td colspan="2" class="new"><a href="xueread.asp?flag=6&xuexing=b" target="_blank"><b>B��Ѫ</b></a>��<a href="senread.asp?flag=7&blood=b&astro=������" target="_blank">������</a>&nbsp;<a href="senread.asp?flag=7&blood=b&astro=��ţ��" target="_blank">��ţ��</a>&nbsp;<a href="senread.asp?flag=7&blood=b&astro=˫����" target="_blank">˫����</a>&nbsp;<a href="senread.asp?flag=7&blood=b&astro=��з��" target="_blank">��з��</a>&nbsp;<a href="senread.asp?flag=7&blood=b&astro=ʨ����" target="_blank">ʨ����</a>&nbsp;<a href="senread.asp?flag=7&blood=b&astro=��Ů��" target="_blank">��Ů��</a>&nbsp;<a href="senread.asp?flag=7&blood=b&astro=�����" target="_blank">�����</a>&nbsp;<a href="senread.asp?flag=7&blood=b&astro=��Ы��" target="_blank">��Ы��</a>&nbsp;<a href="senread.asp?flag=7&blood=b&astro=������" target="_blank">������</a>&nbsp;<a href="senread.asp?flag=7&blood=b&astro=ħ����" target="_blank">ħ����</a>&nbsp;<a href="senread.asp?flag=7&blood=b&astro=ˮƿ��" target="_blank">ˮƿ��</a>&nbsp;<a href="senread.asp?flag=7&blood=b&astro=˫����" target="_blank">˫����</a></td>
      </tr>
	  <tr>
        <td colspan="2" class="new"><a href="xueread.asp?flag=6&xuexing=ab" target="_blank"><b>AB��Ѫ</b></a>��<a href="senread.asp?flag=7&blood=ab&astro=������" target="_blank">������</a>&nbsp;<a href="senread.asp?flag=7&blood=ab&astro=��ţ��" target="_blank">��ţ��</a>&nbsp;<a href="senread.asp?flag=7&blood=ab&astro=˫����" target="_blank">˫����</a>&nbsp;<a href="senread.asp?flag=7&blood=ab&astro=��з��" target="_blank">��з��</a>&nbsp;<a href="senread.asp?flag=7&blood=ab&astro=ʨ����" target="_blank">ʨ����</a>&nbsp;<a href="senread.asp?flag=7&blood=ab&astro=��Ů��" target="_blank">��Ů��</a>&nbsp;<a href="senread.asp?flag=7&blood=ab&astro=�����" target="_blank">�����</a>&nbsp;<a href="senread.asp?flag=7&blood=ab&astro=��Ы��" target="_blank">��Ы��</a>&nbsp;<a href="senread.asp?flag=7&blood=ab&astro=������" target="_blank">������</a>&nbsp;<a href="senread.asp?flag=7&blood=ab&astro=ħ����" target="_blank">ħ����</a>&nbsp;<a href="senread.asp?flag=7&blood=ab&astro=ˮƿ��" target="_blank">ˮƿ��</a>&nbsp;<a href="senread.asp?flag=7&blood=ab&astro=˫����" target="_blank">˫����</a></td>
      </tr>    
    </tbody>
  </table>
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <tbody>
      <tr>
	  <tr>
        <td colspan="2" class="ttop">������/���ջ�:</td>
      </tr>
      <tr>
        <td class="new"><form action=../shengrishu/ method=post target=_blank>
	<input type=hidden name=flag value=8>
	<select name="m" class="sele02" id="m">
	  <option value="1" selected>��</option>
	  <option value="1">1��</option>
	  <option value="2">2��</option>
	  <option value="3">3��</option>
	  <option value="4">04</option>
	  <option value="5">05</option>
	  <option value="6">06</option>
	  <option value="7">07</option>
	  <option value="8">08</option>
	  <option value="9">09</option>

    <option value="10">10</option>
    <option value="11">11</option>
    <option value="12">12</option>
	</select>
	<select name="d" class="sele02" id="d">
	  <option value="1" selected>��</option>
	  <option value="1">01</option>
	  <option value="2">02</option>
	  <option value="3">03</option>
	  <option value="4">04</option>
	  <option value="5">05</option>
	  <option value="6">06</option>
	  <option value="7">07</option>
	  <option value="8">08</option>
	  <option value="9">09</option>
	  <option value="10">10</option>
	  <option value="11">11</option>
	  <option value="12">12</option>
	  <option value="13">13</option>
	  <option value="14">14</option>
	  <option value="15">15</option>
	  <option value="16">16</option>
	  <option value="17">17</option>
	  <option value="18">18</option>
	  <option value="19">19</option>
	  <option value="20">20</option>
	  <option value="21">21</option>
	  <option value="22">22</option>
	  <option value="23">23</option>
	  <option value="24">24</option>
	  <option value="25">25</option>
	  <option value="26">26</option>
	  <option value="27">27</option>
	  <option value="28">28</option>
	  <option value="29">29</option>
	  <option value="30">30</option>
	  <option value="31">31</option>
	</select>
	 
 
	<select name=flag1 class="sele01" id="flag1">
	  <option value="������">������</option>
	</select>
 
	<input name="" type="submit" value="��ѯ" class="bot01"   /></p>
 </form></td>
        <td class="new"><form action=../shengrihua/ method=post target=_blank>
	<input type=hidden name=flag value=8>
	<select name="m" class="sele02" id="m">
	  <option value="1" selected>��</option>
	  <option value="1">1��</option>
	  <option value="2">2��</option>
	  <option value="3">3��</option>
	  <option value="4">04</option>
	  <option value="5">05</option>
	  <option value="6">06</option>
	  <option value="7">07</option>
	  <option value="8">08</option>
	  <option value="9">09</option>

    <option value="10">10</option>
    <option value="11">11</option>
    <option value="12">12</option>
	</select>
	<select name="d" class="sele02" id="d">
	  <option value="1" selected>��</option>
	  <option value="1">01</option>
	  <option value="2">02</option>
	  <option value="3">03</option>
	  <option value="4">04</option>
	  <option value="5">05</option>
	  <option value="6">06</option>
	  <option value="7">07</option>
	  <option value="8">08</option>
	  <option value="9">09</option>
	  <option value="10">10</option>
	  <option value="11">11</option>
	  <option value="12">12</option>
	  <option value="13">13</option>
	  <option value="14">14</option>
	  <option value="15">15</option>
	  <option value="16">16</option>
	  <option value="17">17</option>
	  <option value="18">18</option>
	  <option value="19">19</option>
	  <option value="20">20</option>
	  <option value="21">21</option>
	  <option value="22">22</option>
	  <option value="23">23</option>
	  <option value="24">24</option>
	  <option value="25">25</option>
	  <option value="26">26</option>
	  <option value="27">27</option>
	  <option value="28">28</option>
	  <option value="29">29</option>
	  <option value="30">30</option>
	  <option value="31">31</option>
	</select>
	 
 
	<select name=flag1 class="sele01" id="flag1">
	  <option value="���ջ�">���ջ�</option>
	</select>
 
	<input name="" type="submit" value="��ѯ" class="bot01"   /></p>
 </form></td>
      </tr> 
      </tr>    
    </tbody>
  </table>
  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <tbody>
      <tr>
        <td width="100%" colspan="6" class="ttop">ʮ���������</td>
      </tr>
      <tr>
        <td colspan="6" class="new2"><span class="red">����֪���Լ�����������֪����ͬ�������Ը񣬰��飬��ҵ����𣿿����������һ���Լ��������ɣ�</span></td>
      </tr>
      <tr>
        <td class="new"><a href="12xz.asp?flag=4&astro=������" target="_blank"><img src="../images/xz_1.gif" alt="������" width="90" height="80" border="0"></a></td>
        <td class="new"><a href="12xz.asp?flag=4&astro=��ţ��" target="_blank"><img src="../images/xz_2.gif" alt="��ţ��" width="90" height="80" border="0"></a></td>
        <td class="new"><a href="12xz.asp?flag=4&astro=˫����" target="_blank"><img src="../images/xz_3.gif" alt="˫����" width="90" height="80" border="0"></a></td>
        <td class="new"><a href="12xz.asp?flag=4&astro=��з��" target="_blank"><img src="../images/xz_4.gif" alt="��з��" width="90" height="80" border="0"></a></td>
        <td class="new"><a href="12xz.asp?flag=4&astro=ʨ����" target="_blank"><img src="../images/xz_5.gif" alt="ʨ����" width="90" height="80" border="0"></a></td>
        <td class="new"><a href="12xz.asp?flag=4&astro=��Ů��" target="_blank"><img src="../images/xz_6.gif" alt="��Ů��" width="90" height="80" border="0"></a></td>
      </tr>
      <tr>
        <td class="new"><DIV align="center"><STRONG>������</STRONG></DIV></td>
        <td class="new"><DIV align="center"><STRONG>��ţ��</STRONG></DIV></td>
        <td class="new"><DIV align="center"><STRONG>˫����</STRONG></DIV></td>
        <td class="new"><DIV align="center"><STRONG>��з��</STRONG></DIV></td>
        <td class="new"><DIV align="center"><STRONG>ʨ����</STRONG></DIV></td>
        <td class="new"><DIV align="center"><STRONG>��Ů��</STRONG></DIV></td>
      </tr>
      <tr>
        <td class="new"><a href="12xz.asp?flag=4&astro=�����" target="_blank"><img src="../images/xz_7.gif" alt="�����" width="90" height="80" border="0"></a></td>
        <td class="new"><a href="12xz.asp?flag=4&astro=��Ы��" target="_blank"><img src="../images/xz_8.gif" alt="��Ы��" width="90" height="80" border="0"></a></td>
        <td class="new"><a href="12xz.asp?flag=4&astro=������" target="_blank"><img src="../images/xz_9.gif" alt="������" width="90" height="80" border="0"></a></td>
        <td class="new"><a href="12xz.asp?flag=4&astro=ħ����" target="_blank"><img src="../images/xz_10.gif" alt="ħ����" width="90" height="80" border="0"></a></td>
        <td class="new"><a href="12xz.asp?flag=4&astro=ˮƿ��" target="_blank"><img src="../images/xz_11.gif" alt="ˮƿ��" width="90" height="80" border="0"></a></td>
        <td class="new"><a href="12xz.asp?flag=4&astro=˫����" target="_blank"><img src="../images/xz_12.gif" alt="˫����" width="90" height="80" border="0"></a></td>
      </tr>
      <tr>
        <td class="new"><DIV align="center"><STRONG>�����</STRONG></DIV></td>
        <td class="new"><DIV align="center"><STRONG>��Ы��</STRONG></DIV></td>
        <td class="new"><DIV align="center"><STRONG>������</STRONG></DIV></td>
        <td class="new"><DIV align="center"><STRONG>ħ����</STRONG></DIV></td>
        <td class="new"><DIV align="center"><STRONG>ˮƿ��</STRONG></DIV></td>
        <td class="new"><DIV align="center"><STRONG>˫����</STRONG></DIV></td>
      </tr>
    </tbody>
  </table>

<!-- ����������ʼ -->
<%
ostr=year(now)&month(now)&day(now)
%>
<table width="100%" border="0" align="center" align="center" cellpadding="0" cellspacing="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <tr>
      <td colspan="6"  valign="top" class="ttop" style="PADDING-BOTTOM: 1px">������������</td>
    </tr>
	<tr>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=1&day=<%=ostr%>"><img border="0" src="../xingzuo/images/1.gif" width="78" height="81"></a></td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=2&day=<%=ostr%>"><img border="0" src="../xingzuo/images/2.gif" width="81" height="81"></a></td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=3&day=<%=ostr%>"><img border="0" src="../xingzuo/images/3.gif" width="78" height="79"></a></td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=4&day=<%=ostr%>"><img border="0" src="../xingzuo/images/4.gif" width="80" height="79"></a></td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=5&day=<%=ostr%>"><img border="0" src="../xingzuo/images/5.gif" width="79" height="77"></a></td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=6&day=<%=ostr%>"><img border="0" src="../xingzuo/images/6.gif" width="79" height="79"></a></td>
	</tr>
	<tr>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=1&day=<%=ostr%>">������</a><br>
		(3.21-4.20)</td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=2&day=<%=ostr%>">��ţ��</a><br>
		(4.21-5.20)</td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=3&day=<%=ostr%>">˫����</a><br>
		(5.21-6.21)</td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=4&day=<%=ostr%>">��з��</a><br>
		(6.22-7.22)</td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=5&day=<%=ostr%>">ʨ����</a><br>
		(7.23-8.22)</td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=6&day=<%=ostr%>">��Ů��</a><br>
		(8.23-9.22)</td>
	</tr>
	<tr>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=7&day=<%=ostr%>"><img border="0" src="../xingzuo/images/7.gif" width="80" height="82"></a></td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=8&day=<%=ostr%>"><img border="0" src="../xingzuo/images/8.gif" width="80" height="78"></a></td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=9&day=<%=ostr%>"><img border="0" src="../xingzuo/images/9.gif" width="74" height="74"></a></td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=10&day=<%=ostr%>"><img border="0" src="../xingzuo/images/10.gif" width="79" height="82"></a></td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=11&day=<%=ostr%>"><img border="0" src="../xingzuo/images/11.gif" width="80" height="81"></a></td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=12&day=<%=ostr%>"><img border="0" src="../xingzuo/images/12.gif" width="78" height="77"></a></td>
	</tr>
	<tr>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=7&day=<%=ostr%>">�����</a><br>
		(9.23-10.22)</td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=8&day=<%=ostr%>">��Ы��</a><br>
		(10.23-11.21)</td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=9&day=<%=ostr%>">������</a><br>
		(11.22-12.21)</td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=10&day=<%=ostr%>">Ħ����</a><br>
		(12.22-1.19)</td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=11&day=<%=ostr%>">ˮƿ��</a><br>
		(1.20-2.19)</td>
		<td><a target="_blank" href="/xingzuo/xzlist.asp?id=12&day=<%=ostr%>">˫����</a><br>
		(2.20-3.20)</td>
	</tr>
</table>
<!-- ������������ -->
      <!--#INCLUDE FILE="../dq.asp"-->
	</div>
    </td>
    <td valign="top"><!--#INCLUDE FILE="../right.asp"--></td>
  </tr>
</table>
<!--#include file="../foot.asp"-->
</body></html>