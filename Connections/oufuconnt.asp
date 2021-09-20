<%
' FileName="Connection_ado_conn_string.htm"
' Type="ADO" 
' DesigntimeType="ADO"
' HTTP="false"
' Catalog=""
' Schema=""
Dim MM_oufuconnt_STRING
MM_oufuconnt_STRING = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath("\senlon\#data#.asp")
%>



<%
Dim adgl
Dim adgl_cmd
Dim adgl_numRows

Set adgl_cmd = Server.CreateObject ("ADODB.Command")
adgl_cmd.ActiveConnection = MM_oufuconnt_STRING
adgl_cmd.CommandText = "SELECT * FROM adbiao" 
adgl_cmd.Prepared = true

Set adgl = adgl_cmd.Execute
adgl_numRows = 0
%>
<%
Dim wzgl
Dim wzgl_cmd
Dim wzgl_numRows

Set wzgl_cmd = Server.CreateObject ("ADODB.Command")
wzgl_cmd.ActiveConnection = MM_oufuconnt_STRING
wzgl_cmd.CommandText = "SELECT * FROM wanzhansite" 
wzgl_cmd.Prepared = true

Set wzgl = wzgl_cmd.Execute
wzgl_numRows = 0
%>