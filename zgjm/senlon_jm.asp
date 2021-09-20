<% 
const senlontitle="周公解梦列表"
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
response.write("欧傅提示：对不起，未搜索到相关解梦，请返回重新搜索吧，建议用单个汉字或单个词进行搜索哦~~<p>你还可以打开 http://zhgjm.aa963.com/ 查询</p>")
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
<TITLE>梦见<%=request("word")%>代表什么意思_周公解梦第<%=currentPage%>页_实用查询工具大全 - Powered by Senlon!</TITLE>

</head>
<body><SCRIPT language=javascript>
<!--
function valid_checkdream(){
if(  document.form1.word.value.length==""  )
{window.alert("对不起，请输入梦的关键字！");
document.form1.word.focus();
return false;
} ;
if(  document.form1.word.value.length>20  )
{window.alert("对不起，请将描述控制在20个字以内！");
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
请输入做梦内容的关键字 ：<input name="word" type="text" id="word" size="20" maxLength="20">
<input type="submit" name="Submit1" value="开始解梦" style="cursor:hand;">
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
<font color="green">［<%=i%>.<%=rs2("jmlb")%>］</font> <font color="red">（<%=rs("title")%>）</font><br />
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
<!--分页开始-->
<div>
<script>
PageCount=<%=mpage%> //总页数
topage=<%=currentPage%>   //当前停留页
if (topage>1){document.write("<a href='?act=2&page=0&word=<%=request("word")%>' title='上一页'>上一页</a>");}
for (var i=1; i <= PageCount; i++) {
if (i <= topage+6 && i >= topage-6 || i==1 || i==PageCount){
if (i > topage+7 || i < topage-5 && i!=1 && i!=2 ){document.write(" ... ");}
if (topage==i){document.write("<font color=#d2d0d0> "+ i +" </font>");}
else{
document.write(" <a href='?act=2&page="+i+"&word=<%=request("word")%>'>["+ i +"]</a> ");
}
}
}
if (PageCount-topage>=1){document.write("<a href='?act=2&page=2&word=<%=request("word")%>' title='下一页'>下一页</a>");}
  </script>
</div>
<!--分页结束-->
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
