<% 
const senlontitle="������ŮԤ��"
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>������ŮԤ��,�幬������ŮԤ���_ʵ�ò�ѯ���ߴ�ȫ - Powered by Senlon!</TITLE>
<!--#include file="../senlon/getuc.asp"-->

</head>
<body><!--#include file="../senlon/function.asp"--><!--#include file="../senlon/conn.asp"-->  
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
  <table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
<tbody><tr>
  <td width="100%" class="ttd"><span class="red">������Ů:</span>
    <br><br>��������Ů�����������ԡ��幬��ص�������ŮԤ�������Ԥ��һ�����ǻ������б�������Ů��������ô�Ϳ�ʼ����ɣ���������ο�Ŷ������̫�ţ�<br><br></td>
</tr><%if request("act")="ok" then
mqname=request("mqname")
if mqname="" then
mqname="�װ���511cha����"
end if
cs=request("cs")
yue=request("yue")
nn=request("nn")
snsn=request("snsn")%>
  <tr>
    <td colspan="2" class="new"><font color="#0000FF"><%=mqname%></font>������<br><%if request("cs")=1 then
sql="select * from snsn where nn='"&nn&"' "
set rs=conn.execute(sql)
if not (rs.bof and rs.eof) then
baby=rs(""&yue&"")
end if
if baby="��" then
baby="С����"
elseif baby="Ů" then
baby="С����"
end if
%>��ϲ�����������㣬��ܿ��ܻ���һ��&nbsp;<font color="#FF0000"><%=baby%></font><%else%><%
sql="select * from snsn where nn='"&nn&"' "
set rs=conn.execute(sql)
if rs("1")=snsn then
yuef=yuef&"һ��&nbsp;"
end if
if rs("2")=snsn then
yuef=yuef&"����&nbsp;"
end if
if rs("3")=snsn then
yuef=yuef&"����&nbsp;"
end if
if rs("4")=snsn then
yuef=yuef&"����&nbsp;"
end if
if rs("5")=snsn then
yuef=yuef&"����&nbsp;"
end if
if rs("6")=snsn then
yuef=yuef&"����&nbsp;"
end if
if rs("7")=snsn then
yuef=yuef&"����&nbsp;"
end if
if rs("8")=snsn then
yuef=yuef&"����&nbsp;"
end if
if rs("9")=snsn then
yuef=yuef&"����&nbsp;"
end if
if rs("10")=snsn then
yuef=yuef&"ʮ��&nbsp;"
end if
if rs("11")=snsn then
yuef=yuef&"ʮһ��&nbsp;"
end if
if rs("12")=snsn then
yuef=yuef&"ʮ����&nbsp;"
end if
sql="select * from snsn where nn='"&nn+1&"' "
set rs=conn.execute(sql)
if rs("1")=snsn then
nyuef=nyuef&"һ��&nbsp;"
end if
if rs("2")=snsn then
nyuef=nyuef&"����&nbsp;"
end if
if rs("3")=snsn then
nyuef=nyuef&"����&nbsp;"
end if
if rs("4")=snsn then
nyuef=nyuef&"����&nbsp;"
end if
if rs("5")=snsn then
nyuef=nyuef&"����&nbsp;"
end if
if rs("6")=snsn then
nyuef=nyuef&"����&nbsp;"
end if
if rs("7")=snsn then
nyuef=nyuef&"����&nbsp;"
end if
if rs("8")=snsn then
nyuef=nyuef&"����&nbsp;"
end if
if rs("9")=snsn then
nyuef=nyuef&"����&nbsp;"
end if
if rs("10")=snsn then
nyuef=nyuef&"ʮ��&nbsp;"
end if
if rs("11")=snsn then
nyuef=nyuef&"ʮһ��&nbsp;"
end if
if rs("12")=snsn then
nyuef=nyuef&"ʮ����&nbsp;"
end if
if snsn="��" then
baby="С����"
elseif snsn="Ů" then
baby="С����"
end if
%>�����������<font color="#FF0000"><%=baby%></font>����ô��������ũ�� <font color="#0000FF">���� ��<%=yuef%>&nbsp;&nbsp;���꣺<%=nyuef%></font> ���еĻ�����Ƚϴ�<%end if%><br><br>
      <a class="red" href="./">���²���</a>  </td>
<%else%>
  <tr>
    <td colspan="2" class="ttd"><font color="#FF0000">*����Ϊĸ���������䣬�·�ָ�����·ݣ���ũ��Ϊ׼�������£��ϰ������ϸ��·ݼ��㣬�°������¸��·ݼ��㡣</font>
   </td>
    </tr>  <tr>
    <td colspan="2" class="ttd"><font color="#0000ff">*��ҪԤ�ⱦ���Ա�</font>
   </td>
    </tr><form name="theform" method="post" action="">
<input type="hidden" name="act" value="ok" />
<input type="hidden" name="cs" value="1" /><tr>
    <td colspan="2" class="ttd">�ҵķ����У�<input name="mqname" type="text">&nbsp;���근��(����)��:<select name="nn" size="1" style="font-size: 9pt">
      ><%for i=18 to 45%><option value="<%=i%>" <%if i=22 then%> selected<%end if%>><%=i%></option><%next%>
     </select>&nbsp;���е��·�(<a href="/wannianli/" target="_blank"><font color=red>ũ��</font></a>)��:<select name="yue" size="1" style="font-size: 9pt">
      ><%for i=1 to 12%><option value="<%=i%>" ><%=i%></option><%next%>
     </select>�� &nbsp;<input name="sub" type="submit" value="��ʼ����">
   </td></form>
    </tr> <tr>
    <td colspan="2" class="ttd"><font color="#0000ff">*��Ҫ��ѯ�ʺϻ��е��·ݣ�</font>
   </td>
    </tr><form name="theform" method="post" action="">
<input type="hidden" name="act" value="ok" />
<input type="hidden" name="cs" value="2" /><tr>
    <td colspan="2" class="new">�ҵķ����У�<input name="mqname" type="text">&nbsp;���근��(����)��:<select name="nn" size="1" style="font-size: 9pt">
      ><%for i=18 to 45%><option value="<%=i%>" <%if i=22 then%> selected<%end if%>><%=i%></option><%next%>
     </select>&nbsp;�Ҽƻ�����:<select name="snsn" size="1" style="font-size: 9pt">
       <option value="��">С����</option>
       <option value="Ů">С����</option>

     </select>
     &nbsp;
     <input name="sub" type="submit" value="��ʼ����">
   </td></form>
    </tr><%end if%>
</tbody>
</table>
      <!--#INCLUDE FILE="../dq.asp"-->
	</div>
    </td>
    <td valign="top"><!--#INCLUDE FILE="../right.asp"--></td>
  </tr>
</table>
<!--#include file="../foot.asp"-->
</center>
</body></html>