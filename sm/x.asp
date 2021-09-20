<script language="javascript">
<!--
function checkname()
{
while(document.theform.xing.value.indexOf(" ")!=-1){
document.theform.xing.value=document.theform.xing.value.replace(" ","");
}
while(document.theform.xing.value.indexOf("	")!=-1){
document.theform.xing.value=document.theform.xing.value.replace("	","");
}
if (document.theform.xing.value.length < 1 || document.theform.xing.value.length>10)
{
alert("错误：姓氏应在1-10个汉字之间！");
document.theform.xing.focus();
return (false);
}
}
//-->
</script><table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <TBODY>
  
<form method="POST" action="" onSubmit="return checkname();" name="theform"><tr>
        <TD class=new style="PADDING-BOTTOM: 8px" vAlign=top>请输入姓氏：
          <input maxLength="10" type="text" name="xing" size="16" value="<%=xing%>" /> <input type="submit" value="开始搜索" style="cursor:hand;" /><input type="hidden" name="act" value="ok" />
      </TD>
      </tr></form>
    </TBODY>
  </TABLE> <%if request.form("act")="ok" then
	  sxing=request.form("xing")
	  sxing1=request.form("xing")
	  sxing2=request.form("xing")
	  else
	  sxing=xing
	  sxing1=xing
	  sxing2=xing
	  end if
%><table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <TBODY>
      <tr>
        <TD width="33%" vAlign=top class=ttd style="PADDING-BOTTOM: 8px">姓氏：&nbsp;<strong><%=sxing%></strong></TD>
        <TD width="33%" vAlign=top class=ttd style="PADDING-BOTTOM: 8px">繁体：&nbsp;<strong><%=GbToBig(sxing)%></strong></TD>
        <TD class=ttd style="PADDING-BOTTOM: 8px" vAlign=top>拼音：&nbsp;<strong><%=c(sxing1)%></strong></TD>
      </tr>
<tr>
        <TD colspan="3" vAlign=top class=new style="PADDING-BOTTOM: 8px"><strong>详细资料:</strong></TD>
      </tr> </tr>
<tr>
        <TD colspan="3" vAlign=top class=new style="PADDING-BOTTOM: 8px"><%set rs=server.createobject("adodb.recordset")
		sql="select * from xing where xing='"&sxing2&"'"
		rs.open sql,conn,1,1
if rs.eof then
%>系统没有找到此姓的资料！<%else%><%=rs("intro")%><%end if
rs.close%></TD>
      </tr>
    </TBODY>
  </TABLE>
