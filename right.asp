<%
Dim newsright
Dim newsright_cmd
Dim newsright_numRows

Set newsright_cmd = Server.CreateObject ("ADODB.Command")
newsright_cmd.ActiveConnection = MM_oufuconnt_STRING
newsright_cmd.CommandText = "SELECT * FROM newsbiao ORDER BY id DESC" 
newsright_cmd.Prepared = true

Set newsright = newsright_cmd.Execute
newsright_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 10
Repeat1__index = 0
newsright_numRows = newsright_numRows + Repeat1__numRows
%>

<div id="right_00">
  <p><a href="/newslist.asp" class="heishe">最新发布</a></p>
  <hr  class="shplink"/><table width="100%" border="0" cellpadding="4">
    <% 
While ((Repeat1__numRows <> 0) AND (NOT newsright.EOF)) 
%>
  <tr>
    <td width="4%"><img src="/img/icon.gif" width="8" height="7" /></td>
    <td width="96%"><a href="/news.asp?id=<%=newsright.Fields.Item("id")%>"><%
Dim zx00
zx00=newsright.Fields.Item("title")
zx01=left(zx00,15)
response.Write(zx01)
%></a></td>
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  newsright.MoveNext()
Wend
%>
  </table>
</div>
<div id="right_02"><p><%=(adgl.Fields.Item("addaima04").Value)%></p></div>
<%
newsright.Close()
Set newsright = Nothing
%>
