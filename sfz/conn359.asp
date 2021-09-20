

<%

  on error resume next
   connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & server.mappath("senlon.asp")
     set conn=server.createobject("ADODB.CONNECTION")
     conn.open connstr 

Function HtmlEncode(Content) 
  Content = Replace(Content, ">", "&gt;") 
  Content = Replace(Content, "<", "&lt;")
  Content = Replace(Content, " ", "&nbsp;")
  Content = Replace(Content, "'", "")
  Content = Replace(Content, CHR(10),"<BR>") 
  HtmlEncode = content 
End Function
%>
