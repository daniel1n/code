<% 
const senlontitle="�����б�"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<%
ostr=year(now)&month(now)&day(now)
%>

<%
dim iid

dbpath = server.mappath("senlon.mdb")
connstr = "driver={microsoft access driver (*.mdb)};dbq="& dbpath
set conn = server.createobject("adodb.connection")
conn.open connstr



if request("id") > 0 and request("id") < 13 then
	iid = request("id")
else
	iid = 1
end if

sql = "select * from luck where id = " & iid
set rs = server.createobject("adodb.recordset")
rs.open sql,conn,3,2

if rs("date") = date then	
	istar = rs("star")
	iall = rs("all")
	ilove = rs("love")
	iwork = rs("work")
	ibusiness = rs("business")
	ihealth = rs("health")
	italk = rs("talk")
	icolor = rs("color")
	inumber = rs("number")
	isoon = rs("soon")
	imain = rs("main")
else
	iurl = "http://astro.sina.com.cn/aries.html"
	istar = "������"
	
	select case iid
	case 1
	iurl = "http://astro.sina.com.cn/pc/west/frame0_0.html"
	istar = "������"
	case 2
	iurl = "http://astro.sina.com.cn/pc/west/frame0_1.html"
	istar = "��ţ��"
	case 3
	iurl = "http://astro.sina.com.cn/pc/west/frame0_2.html"
	istar = "˫����"
	case 4
	iurl = "http://astro.sina.com.cn/pc/west/frame0_3.html"
	istar = "��з��"
	case 5
	iurl = "http://astro.sina.com.cn/pc/west/frame0_4.html"
	istar = "ʨ����"
	case 6
	iurl = "http://astro.sina.com.cn/pc/west/frame0_5.html"
	istar = "��Ů��"
	case 7
	iurl = "http://astro.sina.com.cn/pc/west/frame0_6.html"
	istar = "�����"
	case 8
	iurl = "http://astro.sina.com.cn/pc/west/frame0_7.html"
	istar = "��Ы��"
	case 9
	iurl = "http://astro.sina.com.cn/pc/west/frame0_8.html"
	istar = "������"
	case 10
	iurl = "http://astro.sina.com.cn/pc/west/frame0_9.html"
	istar = "Ħ����"
	case 11
	iurl = "http://astro.sina.com.cn/pc/west/frame0_10.html"
	istar = "ˮƿ��"
	case 12
	iurl = "http://astro.sina.com.cn/pc/west/frame0_11.html"
	istar = "˫����"
	end select

	function ichange(str)
		finalstr = ""
		for i = 1 to lenb(str)
			icharcode = ascb(midb(str,i,1))
			if icharcode < &H80 then
			finalstr = finalstr & chr(icharcode)
			else
			inextcode = ascb(midb(str,i+1,1))
			finalstr = finalstr & chr(clng(icharcode) * &H100 + cint(inextcode))
			i = i + 1
			end if
		next
		ichange = finalstr
	end function

	set iconnect = createobject("Microsoft.XMLHTTP")
	icode = iconnect.open ("GET",iurl,false) 
	iconnect.send()

	icode = ichange(iconnect.responsebody)

	if icode <> "" then
		inum = len(icode) - instr(icode,"<div class=""lotconts"">") - 22
		if inum < len(icode) - 22 then
		imain = right(icode,inum)
		inum = instr(imain,"</div>") - 1
		imain = left(imain,inum)
		end if

		inum = len(icode) - instr(icode,"�ۺ�����</h4><p>") - 11
		if inum < len(icode) - 11 then
		iall = right(icode,inum)
		inum = instr(iall,"</div>") - 1
		iall = left(iall,inum)
		iall = Ubound(split(iall,"star.gif"))
		end if
		
	
		inum = len(icode) - instr(icode,"��������</h4><p>") - 11
		if inum < len(icode) - 11 then
		ilove = right(icode,inum)
		inum = instr(ilove,"</div>") - 1
		ilove = Ubound(split(left(ilove,inum),"star.gif"))
		end if
	
		inum = len(icode) - instr(icode,"����״��</h4><p>") - 11
		if inum < len(icode) - 11 then
		iwork = right(icode,inum)
		inum = instr(iwork,"</div>") - 1
		iwork = Ubound(split(left(iwork,inum),"star.gif"))
		end if
		
		inum = len(icode) - instr(icode,"���Ͷ��</h4><p>") - 11
		if inum < len(icode) - 11 then
		ibusiness = right(icode,inum)
		inum = instr(ibusiness,"</div>") - 1
		ibusiness = Ubound(split(left(ibusiness,inum),"star.gif"))
		end if
		
		inum = len(icode) - instr(icode,"����ָ��</h4><p>") - 11
		if inum < len(icode) - 11 then
		ihealth = right(icode,inum)
		inum = instr(ihealth,"</p></div>") - 1
		ihealth = left(ihealth,inum)
		end if
		
		inum = len(icode) - instr(icode,"��̸ָ��</h4><p>") - 11
		if inum < len(icode) - 11 then
		italk = right(icode,inum)
		inum = instr(italk,"</p></div>") - 1
		italk = left(italk,inum)
		end if
		
		inum = len(icode) - instr(icode,"������ɫ</h4><p>") - 11
		if inum < len(icode) - 11 then
		icolor = right(icode,inum)
		inum = instr(icolor,"</p></div>") - 1
		icolor = left(icolor,inum)
		end if
		
		inum = len(icode) - instr(icode,"��������</h4><p>") - 11
		if inum < len(icode) - 11 then
		inumber = right(icode,inum)
		inum = instr(inumber,"</p></div>") - 1
		inumber = left(inumber,inum)
		end if
		
		inum = len(icode) - instr(icode,"��������</h4><p>") - 11
		if inum < len(icode) - 11 then
		isoon = right(icode,inum)
		inum = instr(isoon,"</p></div>") - 1
		isoon = left(isoon,inum)
		end if

		rs("all") = iall
		rs("love") = ilove
		rs("work") = iwork
		rs("business") = ibusiness
		rs("health") = ihealth
		rs("talk") = italk
		rs("color") = icolor
		rs("number") = inumber
		rs("soon") = isoon
		rs("main") = imain
		rs("date") = date
		rs.update
	
	end if

end if

rs.close
set rs = nothing
conn.close
set conn = nothing
%>

<TITLE><%=year(now)%>��<%=month(now)%>��<%=day(now)%>��<%=istar%>��������_��������_ʵ�ò�ѯ���ߴ�ȫ - Powered by Senlon!</TITLE>

</head>
<body>
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
<table width="100%" border="0" align="center" align="center" cellpadding="0" cellspacing="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
	<tr>
		<td><a target="_self" href="xzlist.asp?id=1&day=<%=ostr%>"><img border="0" src="images/1.gif" width="78" height="81"></a></td>
		<td><a target="_self" href="xzlist.asp?id=2&day=<%=ostr%>"><img border="0" src="images/2.gif" width="81" height="81"></a></td>
		<td><a target="_self" href="xzlist.asp?id=3&day=<%=ostr%>"><img border="0" src="images/3.gif" width="78" height="79"></a></td>
		<td><a target="_self" href="xzlist.asp?id=4&day=<%=ostr%>"><img border="0" src="images/4.gif" width="80" height="79"></a></td>
		<td><a target="_self" href="xzlist.asp?id=5&day=<%=ostr%>"><img border="0" src="images/5.gif" width="79" height="77"></a></td>
		<td><a target="_self" href="xzlist.asp?id=6&day=<%=ostr%>"><img border="0" src="images/6.gif" width="79" height="79"></a></td>
	</tr>
	<tr>
		<td><a target="_self" href="xzlist.asp?id=1&day=<%=ostr%>">������</a><br>
		(3.21-4.20)</td>
		<td><a target="_self" href="xzlist.asp?id=2&day=<%=ostr%>">��ţ��</a><br>
		(4.21-5.20)</td>
		<td><a target="_self" href="xzlist.asp?id=3&day=<%=ostr%>">˫����</a><br>
		(5.21-6.21)</td>
		<td><a target="_self" href="xzlist.asp?id=4&day=<%=ostr%>">��з��</a><br>
		(6.22-7.22)</td>
		<td><a target="_self" href="xzlist.asp?id=5&day=<%=ostr%>">ʨ����</a><br>
		(7.23-8.22)</td>
		<td><a target="_self" href="xzlist.asp?id=6&day=<%=ostr%>">��Ů��</a><br>
		(8.23-9.22)</td>
	</tr>
	<tr>
		<td><a target="_self" href="xzlist.asp?id=7&day=<%=ostr%>"><img border="0" src="images/7.gif" width="80" height="82"></a></td>
		<td><a target="_self" href="xzlist.asp?id=8&day=<%=ostr%>"><img border="0" src="images/8.gif" width="80" height="78"></a></td>
		<td><a target="_self" href="xzlist.asp?id=9&day=<%=ostr%>"><img border="0" src="images/9.gif" width="74" height="74"></a></td>
		<td><a target="_self" href="xzlist.asp?id=10&day=<%=ostr%>"><img border="0" src="images/10.gif" width="79" height="82"></a></td>
		<td><a target="_self" href="xzlist.asp?id=11&day=<%=ostr%>"><img border="0" src="images/11.gif" width="80" height="81"></a></td>
		<td><a target="_self" href="xzlist.asp?id=12&day=<%=ostr%>"><img border="0" src="images/12.gif" width="78" height="77"></a></td>
	</tr>
	<tr>
		<td><a target="_self" href="xzlist.asp?id=7&day=<%=ostr%>">�����</a><br>
		(9.23-10.22)</td>
		<td><a target="_self" href="xzlist.asp?id=8&day=<%=ostr%>">��Ы��</a><br>
		(10.23-11.21)</td>
		<td><a target="_self" href="xzlist.asp?id=9&day=<%=ostr%>">������</a><br>
		(11.22-12.21)</td>
		<td><a target="_self" href="xzlist.asp?id=10&day=<%=ostr%>">Ħ����</a><br>
		(12.22-1.19)</td>
		<td><a target="_self" href="xzlist.asp?id=11&day=<%=ostr%>">ˮƿ��</a><br>
		(1.20-2.19)</td>
		<td><a target="_self" href="xzlist.asp?id=12&day=<%=ostr%>">˫����</a><br>
		(2.20-3.20)</td>
	</tr>
</table>
<div style="PADDING-TOP: 3px"></div>
<table width="100%" border="0" align="center" align="center" cellpadding="0" cellspacing="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
	<tr>
		<td>
<p><a href="./xzyc.asp"><img border="0" src="images/<%=iid%>.gif" alt="<%=istar%>��������"></a></p>
<p><b><%=istar%> <%=year(now)%>��<%=month(now)%>��<%=day(now)%>������</b>��</p>
<p>�ۺ����ƣ�<%for i = 1 to iall%><img border="0" src="images/star.gif"><%next%></p>
<p>�������ƣ�<%for i = 1 to ilove%><img border="0" src="images/star.gif"><%next%></span></p>
<p>����״����<%for i = 1 to iwork%><img border="0" src="images/star.gif"><%next%></p>
<p>���Ͷ�ʣ�<%for i = 1 to ibusiness%><img border="0" src="images/star.gif"><%next%></p>
<p>����ָ����<%=ihealth%></p>
<p>��̸ָ����<%=italk%></p>
<p>������ɫ��<%=icolor%></p>
<p>�������֣�<img src="images/h_<%=inumber%>.gif" ></p>
<p>����������<%=isoon%></p>
<p>�����Ҹ棺<%=imain%></p>
</td>
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


