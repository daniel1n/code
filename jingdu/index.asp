<% 
const senlontitle="�������Ȳ�ѯ"
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>�������Ȳ�ѯ_ʵ�ò�ѯ���ߴ�ȫ - Powered by Senlon!</TITLE>

</head>
<body><!--#include file="../senlon/getuc.asp"-->
<SCRIPT src="comefrom.js"></SCRIPT><body topmargin=50 leftmargin=0 onload=init()>
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
<%
'��������

province1=trim(request("province1"))
city1=trim(request("city1"))

if not isnumeric(jingdu) then jingdu=0
on error resume next
	db1="senlon.asp"
	Set conn1 = Server.CreateObject("ADODB.Connection")
	connstr1="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db1)
	conn1.Open connstr1

if province1<>"" and city1<>"" then
	set rs=server.createobject("ADODB.recordset")
	strsql="select * from jd where sheng='"&province1&"' and xian='"&city1&"'"
	rs.open strsql,conn1,1,1 
	if not rs.eof then
		id=rs("id")
		jingdu=rs("jingdu")
		jingdu=Split(jingdu,".")

	else
	jingdu="û���ҵ���ص���"
	end if
else
end if
conn1.close
set conn1=nothing

%>
  <table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" >
<tbody><form method="POST"  action="" name=form1 onSubmit="return submitchecken();">
<input type="hidden" name="act" value="ok" /><tr>
  <td width="100%" class="ttop"> �������Ȳ�ѯ</td>
    </tr>  <tr>
<td class="new">ʡ�����ƣ�
  <SELECT name=province onchange=select()></SELECT> <INPUT name=province1 style="color:ff0000;" size=6 value="<%=province1%>" >
  �������ƣ�
  <SELECT name=city onchange=select()></SELECT> <INPUT name=city1 style="color:ff0000;" size=8 value="<%=city1%>" >&nbsp;</td>
</tr>     <%if request("act")="ok" then%>
 <tr>
        <td class="new"><p>&nbsp;<br>
���������<font color="#FF0000"><%=province1%>-<%=city1%></font><p>
�������ȣ�<font color="#0000FF"><%=jingdu(0)%>��<%=jingdu(1)%>��</font><p>

����  </td>
      </tr><%end if%>
<tr>
  <td class="new"><input type="submit" name="Submit3" value="��ѯ��������" style='cursor:hand;'></td>
</tr>

</tr>
</form>
</table>

  <table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" >
    <tbody>
   
      <tr>
        <td width="100%" class="ttop">�������Ȳ�ѯ</td>
      </tr>
      <tr>
        <td class="new"><p>˵������ϵͳĿǰ��֧���й�����ĵ������Ȳ�ѯ��֧���й��ؼ����ϳ��еľ��Ȳ�ѯ���������ƽ�β�벻Ҫ�����С������ء����֣�����д��һЩ����ϵͳ����֧��ģ����ʽ��ѯ�����������ƿ���Ϊ�գ� 
 </p>
          <p>���ȷ�ָ��������ϵ�������ꡣ����Ϊ��������һ����������������0�Ⱦ�������ƽ��ļнǡ��������ϵĵ����ڸ�Ȧ���������ԭ�����ڸ�Ȧ�ĽǾ�������ʾ��ͨ����ָ��������ľ��ȡ�Ϊ�����ֵ����ϵ�ÿһ�����������Ǹ����߱�ע�˶���������Ǿ��ȣ�longitude ����ʵ���Ͼ�����������������ƽ��֮��ļнǡ�<br><br>�ӱ����㵽�ϼ��㣬���Ի�������ϱ���������������ֱ�Ĵ�ԲȦ�����������Ȧ����������ЩԲȦ���߶Σ��ͽо��ߡ���Ԫ1884�꣬�����Ϲ涨��ͨ��Ӣ���׶ؽ����ĸ�����������̨��ַ�ľ�����Ϊ���㾭�ȵ���㣬���������������룬Ҳ�ơ����������ߡ������������Ϊ��������180�ȣ����������Ϊ��������180�ȡ���Ϊ������Բ�ģ����Զ���180�Ⱥ�����180�ȵľ�����ͬһ�����ߡ���������180�Ⱦ���Ϊ���������ڱ���ߡ���Ϊ�˱���ͬһ����ʹ��������ͬ�����ڣ��������ڱ�������½��ʱ����ƫ�롣<br><br>�����Ϲ涨����ͨ��Ӣ���׶��׶ظ�����������̨ԭַ����һ�����߶�Ϊ0�㾭�ߣ�Ҳ�б��������ߡ���0�㾭�������򶫡�����������180�㣬�Զ���180�����ڶ�����ϰ�����á�E�������ţ�������180������������ϰ�����á�W�������š�����180���������180���غ���һ�������ϣ��Ǿ���180�㾭�ߡ��ڵ�ͼ���ж�����ʱӦע�⣺�����򶫣����ȵĶ�����С����Ϊ�����ȣ������򶫣����ȵĶ����ɴ�С��Ϊ�����ȣ���0���180�㾭���⣬���ྭ�߶���׼ȷ�����Ƕ����Ȼ��������ȡ���ͬ�ľ��߾��в�ͬ�ĵط�ʱ��ƫ���ĵط�ʱҪ�磬ƫ���ĵط�ʱҪ�١�ÿ15�����ȱ����һ��Сʱ��</p></td>
      </tr>
  </table>
      <!--#INCLUDE FILE="../dq.asp"-->
	</div>
    </td>
    <td valign="top"><!--#INCLUDE FILE="../right.asp"--></td>
  </tr>
</table>
<!--#include file="../foot.asp"-->
</body></html>
