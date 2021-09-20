<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#INCLUDE FILE="top.asp"-->
<%
Dim shyRecordset2__MMColParam
shyRecordset2__MMColParam = "1"
shyid=Request.QueryString("id")-1
If (shyid <> "") Then 
  shyRecordset2__MMColParam = shyid
End If
%>
<%
Dim shyRecordset2
Dim shyRecordset2_cmd
Dim shyRecordset2_numRows

Set shyRecordset2_cmd = Server.CreateObject ("ADODB.Command")
shyRecordset2_cmd.ActiveConnection = MM_oufuconnt_STRING
shyRecordset2_cmd.CommandText = "SELECT * FROM newsbiao WHERE id = ?" 
shyRecordset2_cmd.Prepared = true
shyRecordset2_cmd.Parameters.Append shyRecordset2_cmd.CreateParameter("param1", 5, 1, -1, shyRecordset2__MMColParam) ' adDouble

Set shyRecordset2 = shyRecordset2_cmd.Execute
shyRecordset2_numRows = 0
%>
<%
Dim xyRecordset2__MMColParam
xyRecordset2__MMColParam = "1"
xyid=Request.QueryString("id")+1
If (xyid <> "") Then 
  xyRecordset2__MMColParam = xyid
End If
%>
<%
Dim xyRecordset2
Dim xyRecordset2_cmd
Dim xyRecordset2_numRows

Set xyRecordset2_cmd = Server.CreateObject ("ADODB.Command")
xyRecordset2_cmd.ActiveConnection = MM_oufuconnt_STRING
xyRecordset2_cmd.CommandText = "SELECT * FROM newsbiao WHERE id = ?" 
xyRecordset2_cmd.Prepared = true
xyRecordset2_cmd.Parameters.Append xyRecordset2_cmd.CreateParameter("param1", 5, 1, -1, xyRecordset2__MMColParam) ' adDouble

Set xyRecordset2 = xyRecordset2_cmd.Execute
xyRecordset2_numRows = 0
%>
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.QueryString("id") <> "") Then 
  Recordset1__MMColParam = Request.QueryString("id")
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_oufuconnt_STRING
Recordset1_cmd.CommandText = "SELECT * FROM newsbiao WHERE id = ?" 
Recordset1_cmd.Prepared = true
Recordset1_cmd.Parameters.Append Recordset1_cmd.CreateParameter("param1", 5, 1, -1, Recordset1__MMColParam) ' adDouble

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=(Recordset1.Fields.Item("title").Value)%>_<%=(wzgl.Fields.Item("name").Value)%></title>
</head>

<body>
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top"><div id="mid_left">
    <h3><%=(Recordset1.Fields.Item("title").Value)%></h3>
    <span class="huishe">来源：<%=(Recordset1.Fields.Item("laiyun").Value)%>　日期：<%=(Recordset1.Fields.Item("timedata").Value)%></span>
<hr  class="shplink"/><p><%=(Recordset1.Fields.Item("conntmore").Value)%></p>
<p>上一篇：
  <% If Not shyRecordset2.EOF Or Not shyRecordset2.BOF Then %>
    <a href="news.asp?id=<%=(shyRecordset2.Fields.Item("id").Value)%>"><%=(shyRecordset2.Fields.Item("title").Value)%></a>
    <% End If ' end Not shyRecordset2.EOF Or NOT shyRecordset2.BOF %>
  <% If shyRecordset2.EOF And shyRecordset2.BOF Then %> 
  <span class="huishe">没有了</span>
  <% End If ' end shyRecordset2.EOF And shyRecordset2.BOF %>
</p>
<p>下一篇：
  <% If Not xyRecordset2.EOF Or Not xyRecordset2.BOF Then %>
  <a href="news.asp?id=<%=(xyRecordset2.Fields.Item("id").Value)%>"><%=(xyRecordset2.Fields.Item("title").Value)%></a>
  <% End If ' end Not xyRecordset2.EOF Or NOT xyRecordset2.BOF %>
  <% If xyRecordset2.EOF And xyRecordset2.BOF Then %>
  <span class="huishe">没有了</span>
  <% End If ' end xyRecordset2.EOF And xyRecordset2.BOF %>
</p>
    </div><!--#INCLUDE FILE="dq.asp"--></td>
    <td valign="top"><!--#INCLUDE FILE="right.asp"--></td>
  </tr>
</table>
<!--#INCLUDE FILE="foot.asp"-->
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
<%
xyRecordset2.Close()
Set xyRecordset2 = Nothing
%>
<%
shyRecordset2.Close()
Set shyRecordset2 = Nothing
%>
