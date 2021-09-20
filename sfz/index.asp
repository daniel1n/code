<% 
const senlontitle="身份号码吉凶查询"
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>身份号码吉凶查询_实用查询工具大全 - Powered by Senlon!</TITLE>
<!-- #include file=conn359.asp-->

<style type="text/css">
<!--
.style1 {color: #FF0000}
.style3 {
	color: #FFFFFF;
	font-weight: bold;
}
.style5 {color: #FFFFFF}
.style6 {color: #000000}
-->
</style>
</head>
<body>
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
<table width="760" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
      <tr>
  <td class="new">
<form action="./index.asp" name=form1 method=post>
  <table border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td height="30" ><p>　　自古有云：谋事在人，成事在天。究竟我们可否争取主动，反过来操控自己的命运呢？答案是肯定的。现在我们就可以用身份证号来测试你的性格、爱情、事业与未来，让你更能掌握未来，迈出成功的第一步。<br />
        　　只要凭身份证上的号码，便可推算出你一生的命运，十分奇妙！<br />
        <br /></td>
    </tr>
    <tr>
        <td width="37%" height="30" valign="middle" align=center><p>
          身份证号码:
          <input name=xian size=25 type=text>
          <input name=submit type=submit value="查询">
          (请输入15位或者18位号码)</p>查询不了？<a href="http://sm.aa963.com/sfz/">点击这里试试&gt;&gt;</a>
          <p>重要声明：本查询不会记录查询者的任何信息，只提供所输入数字的吉凶测试，查询结果仅供参考！</p></td>
    </tr>
  </table>
</form>

</td></tr>
</table>

<table align=center border="0" cellpadding="3" cellspacing="1" width="760"  bgcolor="#5391EE">

<%

xian=request("xian")
lenx=len(xian)
if xian<>"" then

if not IsNumeric(left(xian,15)) and xian<>"" then
  response.write "<script>alert('请输入正确身份证号码!');window.history.back();</script>"
end if 

if lenx=15 or lenx=18 then 
    if lenx=15 then 
        yy="19"&mid(xian,7,2)
        mm=mid(xian,9,2)
        dd=mid(xian,11,2)
        aa=mid(xian,15,1)
    end if 
    if  lenx=18 then
        yy=mid(xian,7,4)
        mm=mid(xian,11,2)
        dd=mid(xian,13,2)
        aa=mid(xian,17,1)
    end if 
    if cint(mm)>12 or cint(dd)>31 then
      response.write "<script>alert('请输入正确的身份证号码!');window.history.back();</script>"
    end if  
   else
  response.write "<script>alert('请输入正确的身份证号码!');window.history.back();</script>"
end if 


set rs=server.createobject("adodb.recordset")
      sql="select * from sfz where bm="&left(xian,6)
      rs.open sql,conn,3,3
     
      if not rs.eof then
     
   
%>
<%
if aa mod 2=0 then 
  xb="女"
  else
  xb="男"
end if 

%>
<tr><td bgcolor=ffffff>
查询号码：<%=xian%></td></tr><td bgcolor=ffffff>
原户籍地：<%=rs("dq")%><br></td></tr><td bgcolor=ffffff>
出生年月：<%=yy&"年"&mm&"月"&dd&"日"%><br></td></tr><td bgcolor=ffffff>
性&nbsp;&nbsp;&nbsp;&nbsp;别：<%=xb%>
<%
if lenx=18 then
if mid(xian,18,1)<>cstr(sfzjy(xian)) then
  response.write "</td></tr><td bgcolor=ffffff>"
  response.write "<font color='#FF0000'>提示：身份证校验错误!可能为假！</font>"
  xian=""
  else
   response.write "</td></tr><td bgcolor=ffffff>"
   response.write "结果：身份证号码校验为合法号码!"
   'response.write xian 
end if
  else
  response.write "</td></tr><td bgcolor=ffffff>"
  response.write "新身份证："&left(xian,6)&"19"&right(xian,9)&cstr(sfzjy(xian))
  xian=left(xian,6)&"19"&right(xian,9)&cstr(sfzjy(xian))
  'response.write xian
end if
%>
</td>
</tr>
<%
     n1=mid(xian,7,1)
	 n2=mid(xian,8,1)
	 y1=mid(xian,9,1)
	 y2=mid(xian,10,1)
	 r1=mid(xian,11,1)
	 r2=mid(xian,12,1)
	'年的单双取值 
	  n=cInt(n1)+cInt(n2) 
	  if len(n)=2 then
	     n=cInt(mid(n,1,1))+cInt(mid(n,2,1))
	  end if
	 
	  n=n mod 2
	  if n=0 then 
	    n="双"
	  else 
	    n="单"
	  end if
	  '月的单双取值
	  	  y=cInt(y1)+cInt(y2) 
	  if len(y)=2 then
	     y=cInt(mid(y,1,1))+cInt(mid(y,2,1))
	  end if
	 ' response.Write "这是日的结果"&r
	  y=y mod 2
	  if y=0 then 
	    y="双"
	  else 
	    y="单" 
	  end if
	  '日的单双取值
	  r=cInt(r1)+cInt(r2) 
	  if len(r)=2 then
	     r=cInt(mid(r,1,1))+cInt(mid(r,2,1))
	  end if
	  
	  r=r mod 2
	  if r=0 then 
	    r="双"
	  else 
	    r="单" 
	  end if
	  testnum=n+y+r
	 'response.Write "这是数据的结果"&testnum
	set rs1=server.createobject("adodb.recordset")
      sql="select * from test where testnum='"&testnum&"'"
      rs1.open sql,conn,1,1
     response.write "<tr><td bgcolor=ffffff>测试结果：</td></tr>"
      if not rs1.eof then
	  response.write "<tr><td bgcolor=ffffff>"&rs1("content")&"</td></tr>"
	  else
	  response.write "<tr><td bgcolor=ffffff>511Cha提示：芸芸众生，竟无此人，无从进行命运推理！请施主输入正确的身份证再行测运，实诚功德无量！</td></tr>"
	  end if
  %>
</table>
<%
          end if 
rs.close
set rs=nothing
  conn.close
set conn=nothing
end if 

FUNCTION sfzjy(num)

 if len(num)=15 then
cID = left(num,6)&"19"&right(num,9)
  elseif len(num)=17 or len(num)=18 then
cID = left(num,17)
  end if 

nSum=mid(cID,1,1) * 7
nSum=nsum+mid(cID,2,1) * 9 
nSum=nsum+mid(cID,3,1) * 10 
nSum=nsum+mid(cID,4,1) * 5 
nSum=nsum+mid(cID,5,1) * 8 
nSum=nsum+mid(cID,6,1) * 4
nSum=nsum+mid(cID,7,1) * 2
nSum=nsum+mid(cID,8,1) * 1
nSum=nsum+mid(cID,9,1) * 6
nSum=nsum+mid(cID,10,1) * 3
nSum=nsum+mid(cID,11,1) * 7
nSum=nsum+mid(cID,12,1) * 9
nSum=nsum+mid(cID,13,1) * 10
nSum=nsum+mid(cID,14,1) * 5
nSum=nsum+mid(cID,15,1) * 8
nSum=nsum+mid(cID,16,1) * 4
nSum=nsum+mid(cID,17,1) * 2
'*计算校验位
 check_number=12-nsum mod 11
 If check_number=10 then
     check_number="X"
  elseIf check_number=12 then
     check_number="1"
  elseif check_number=11 then
     check_number="0"
 End if
 sfzjy=check_number
End function
%>
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