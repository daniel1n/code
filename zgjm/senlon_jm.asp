<% 
const senlontitle="�ܹ������б�"
%><!--#include file="../senlon/conn.asp"-->
<%
     MaxPerPage=3
   if not isempty(request("page")) then 
      currentPage=cint(request("page")) 
   else 
      currentPage=1 
   end if 
   set rs=server.createobject("adodb.recordset")
if request("act")=1 then
cxgjz="title"
elseif request("act")=2 then
cxgjz="content"
end if
sql="select * from zgjm where "&cxgjz&" like '%"& request("word") &"%'"
rs.open sql,conn,1,1
if rs.eof then
response.write("ŷ����ʾ���Բ���δ��������ؽ��Σ��뷵�����������ɣ������õ������ֻ򵥸��ʽ�������Ŷ~~<p>�㻹���Դ� http://zhgjm.aa963.com/ ��ѯ</p>")
response.end
else
allshu=rs.recordcount
rs.pagesize=MaxPerPage 
mpage=rs.pagecount     
rs.move  (currentPage-1)*MaxPerPage
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>�μ�<%=request("word")%>����ʲô��˼_�ܹ����ε�<%=currentPage%>ҳ_ʵ�ò�ѯ���ߴ�ȫ - Powered by Senlon!</TITLE>

</head>
<body><SCRIPT language=javascript>
<!--
function valid_checkdream(){
if(  document.form1.word.value.length==""  )
{window.alert("�Բ����������εĹؼ��֣�");
document.form1.word.focus();
return false;
} ;
if(  document.form1.word.value.length>20  )
{window.alert("�Բ����뽫����������20�������ڣ�");
document.form1.word.focus();
return false;
} ;

form1.submit();
return ;
}
//-->
</SCRIPT><!--#include file="../senlon/function.asp"--><!--#include file="../senlon/conn.asp"-->  
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
  <table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
<tbody>
<form method="post" action="senlon_jm.asp" name="form1" onSubmit="return valid_checkdream()" target="dream">
<input type="hidden" name="act" value="ok" />
  <tr>
    <td colspan="2" class="new">
�������������ݵĹؼ��� ��<input name="word" type="text" id="word" size="20" maxLength="20">
<input type="submit" name="Submit1" value="��ʼ����" style="cursor:hand;">
    </form></td>
    </tr>
</tbody>
</table>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
  <tr>
    <td>
<%i=0
do while not rs.eof
  i=i+1 
jmlb=rs("jmlb")
		 set rs2=server.createobject("adodb.recordset")
sql2="select * from jmlb where id="&jmlb
rs2.open sql2,conn,1,3%>
<tr><td style="line-height:200%">
<font color="green">��<%=i%>.<%=rs2("jmlb")%>��</font> <font color="red">��<%=rs("title")%>��</font><br />
<%word=request("word")
content=rs("content")
%>
<%=replace(content,word,"<font color=red>"&word&"</font>")%>
</td></tr>  <%
 
if i>=MaxPerPage then exit do
 rs.movenext
 loop 
   end if
%>

</div>
<!--��ҳ��ʼ-->
<div>
<script>
PageCount=<%=mpage%> //��ҳ��
topage=<%=currentPage%>   //��ǰͣ��ҳ
if (topage>1){document.write("<a href='?act=2&page=0&word=<%=request("word")%>' title='��һҳ'>��һҳ</a>");}
for (var i=1; i <= PageCount; i++) {
if (i <= topage+6 && i >= topage-6 || i==1 || i==PageCount){
if (i > topage+7 || i < topage-5 && i!=1 && i!=2 ){document.write(" ... ");}
if (topage==i){document.write("<font color=#d2d0d0> "+ i +" </font>");}
else{
document.write(" <a href='?act=2&page="+i+"&word=<%=request("word")%>'>["+ i +"]</a> ");
}
}
}
if (PageCount-topage>=1){document.write("<a href='?act=2&page=2&word=<%=request("word")%>' title='��һҳ'>��һҳ</a>");}
  </script>
</div>
<!--��ҳ����-->
</TD></TR>
</TBODY>
</table>
      <!--#INCLUDE FILE="../dq.asp"-->
	</div>
    </td>
    <td valign="top"><!--#INCLUDE FILE="../right.asp"--></td>
  </tr>
</table>
<!--#include file="../foot.asp"-->
</body></html>
