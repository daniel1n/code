<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><div id="foot_ad"><%=(adgl.Fields.Item("addaima03").Value)%></div>
    <div id="foot"><%=(wzgl.Fields.Item("footdm").Value)%></div></td>
  </tr>
</table>
<%
adgl.Close()
Set adgl = Nothing
%>
<%
wzgl.Close()
Set wzgl = Nothing
%>