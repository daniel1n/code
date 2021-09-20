<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="../login.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (true Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
<!--#include file="../../Connections/oufuconnt.asp" -->
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_oufuconnt_STRING
Recordset1_cmd.CommandText = "SELECT * FROM adbiao" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>无标题文档</title>
<link href="../../css/layout02.css" rel="stylesheet" type="text/css" />
</head>

<body>
<table width="80%" border="1" cellpadding="4">
  <tr>
    <td><strong>编号</strong></td>
    <td><strong>广告名称</strong></td>
    <td><strong>广告代码</strong></td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td>01</td>
    <td><%=(Recordset1.Fields.Item("adname01").Value)%></td>
    <td><%=(Recordset1.Fields.Item("addaima01").Value)%></td>
    <td><a href="adxg01.asp">修改</a></td>
  </tr>
  <tr>
    <td>02</td>
    <td><%=(Recordset1.Fields.Item("adname02").Value)%></td>
    <td><%=(Recordset1.Fields.Item("addaima02").Value)%></td>
    <td><a href="adxg02.asp">修改</a></td>
  </tr>
  <tr>
    <td>03</td>
    <td><%=(Recordset1.Fields.Item("adname03").Value)%></td>
    <td><%=(Recordset1.Fields.Item("addaima03").Value)%></td>
    <td><a href="adxg03.asp">修改</a></td>
  </tr>
  <tr>
    <td>04</td>
    <td><%=(Recordset1.Fields.Item("adname04").Value)%></td>
    <td><%=(Recordset1.Fields.Item("addaima04").Value)%></td>
    <td><a href="adxg04.asp">修改</a></td>
  </tr>
  <tr>
    <td>05</td>
    <td><%=(Recordset1.Fields.Item("adname05").Value)%></td>
    <td><%=(Recordset1.Fields.Item("addaima05").Value)%></td>
    <td><a href="adxg05.asp">修改</a></td>
  </tr>
  <tr>
    <td>06</td>
    <td><%=(Recordset1.Fields.Item("adname06").Value)%></td>
    <td><%=(Recordset1.Fields.Item("addaima06").Value)%></td>
    <td><a href="adxg06.asp">修改</a></td>
  </tr>
</table>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
