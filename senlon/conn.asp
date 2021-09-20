<%
on error resume next
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath("/senlon/senlon@2006.asp")
set conn=server.createobject("ADODB.CONNECTION")
conn.open connstr
If Err Then
response.Write "连接数据库出错!"
err.Clear
Set conn = Nothing
Response.End
End If

Function HtmlEncode(Content) 
  Content = Replace(Content, ">", "&gt;") 
  Content = Replace(Content, "<", "&lt;")
  Content = Replace(Content, " ", "&nbsp;")
  Content = Replace(Content, "'", "")
  Content = Replace(Content, CHR(10),"<BR>") 
  HtmlEncode = content 
End Function
%>
