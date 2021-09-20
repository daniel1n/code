<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/oufuconnt.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="login.asp"
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
<%
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
<%
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function
%>
<%
If (CStr(Request("MM_update")) = "form1") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_oufuconnt_STRING
    MM_editCmd.CommandText = "UPDATE wanzhansite SET moban = ?, mulu = ?, name = ?, ymurl = ?, footdm = ? WHERE moban = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 5, 1, -1, MM_IIF(Request.Form("mb"), Request.Form("mb"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 202, 1, 50, Request.Form("ml")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 202, 1, 50, Request.Form("nm")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 202, 1, 50, Request.Form("yu")) ' adVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param5", 203, 1, 536870910, Request.Form("fd")) ' adLongVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param6", 200, 1, 536870910, Request.Form("MM_recordId")) ' adVarChar
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "right.asp?type=true"
    If (Request.QueryString <> "") Then
      If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0) Then
        MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
      Else
        MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
      End If
    End If
    Response.Redirect(MM_editRedirectUrl)
  End If
End If
%>
<%
Dim Recordset1
Dim Recordset1_cmd
Dim Recordset1_numRows

Set Recordset1_cmd = Server.CreateObject ("ADODB.Command")
Recordset1_cmd.ActiveConnection = MM_oufuconnt_STRING
Recordset1_cmd.CommandText = "SELECT * FROM wanzhansite" 
Recordset1_cmd.Prepared = true

Set Recordset1 = Recordset1_cmd.Execute
Recordset1_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>网站设置</title>
<link href="../css/layout02.css" rel="stylesheet" type="text/css" />
</head>

<body>
<form action="<%=MM_editAction%>" method="POST" name="form1" id="form1">
  <p>
    <label for="mb"></label>
    <select name="mb" id="mb">
      <option value="<%=(Recordset1.Fields.Item("moban").Value)%>"><% Dim msh 
	  msh=Recordset1.Fields.Item("moban") 
	  if msh=1 then 
	  response.write("棕黄色") 
	  else 
	  response.write("蓝白色") 
	  end if %></option>
      <option value="1">棕黄色</option>
      <option value="2">蓝白色</option>
      <%
While (NOT Recordset1.EOF)
%>
      <%
  Recordset1.MoveNext()
Wend
If (Recordset1.CursorType > 0) Then
  Recordset1.MoveFirst
Else
  Recordset1.Requery
End If
%>
    </select>
    切换模板</p>
  <p>
    <label for="ml"></label>
    <input name="ml" type="text" id="ml" value="<%=(Recordset1.Fields.Item("mulu").Value)%>" size="10" />
    安装目录
  </p>
  <p>
    <label for="nm"></label>
    <input name="nm" type="text" id="nm" value="<%=(Recordset1.Fields.Item("name").Value)%>" size="15" />
    网站名称</p>
  <p>
    <label for="yu"></label>
    <input name="yu" type="text" id="yu" value="<%=(Recordset1.Fields.Item("ymurl").Value)%>" />
  网站域名</p>
  <p>
    <label for="fd"></label>
    <textarea name="fd" id="fd" cols="45" rows="5"><%=(Recordset1.Fields.Item("footdm").Value)%></textarea>
  底部信息</p>
  <p>
    <input type="submit" name="button" id="button" value="修改" />
  　<% If Trim(Request.QueryString("type"))<>"" Then %>
  <span class="ren">修改成功</span>
  <% End If %></p>
  <input type="hidden" name="MM_update" value="form1" />
  <input type="hidden" name="MM_recordId" value="<%= Recordset1.Fields.Item("moban").Value %>" />
</form>
</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
