<% 
const senlontitle="开运金钱卦"
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>开运金钱卦_实用查询工具大全 - Powered by Senlon!</TITLE>

</head>
<body>
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
  <table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
<tbody><tr>
  <td width="100%" class="new">
<style type="text/css">
<!--
.style1 {color: #FF0000}
.style3 {
	color: #FFFFFF;
	font-weight: bold;
}
.style5 {color: #FFFFFF}
.style6 {color: #000000}
-->
</style>
<table width=100% align=center >
<tr>
  <td align=left>
<SCRIPT LANGUAGE="VBScript">
<!--
dim num1,num2
sub text0_onblur
	if len(form1.text0.value)<5 then 
		msgbox "请输入至少5个字！"
	else
		form1.botton1.disabled=false
	end if
end sub
sub botton1_onclick
	randomize()
	num1=int(60*rnd())+1
	if num1 mod 2 then
		form1.text1.value="你掷的是反"
	else
	  form1.text1.value="你掷的是正"
	end if  
	form1.botton1.disabled=true
	form1.botton2.disabled=false
end sub
sub botton2_onclick
	randomize()
	if num2 mod 2 then
	 	form1.text2.value="你掷的是反"
	 else
	   form1.text2.value="你掷的是正"
	end if  
	  form1.botton2.disabled=true
	if form1.botton1.disabled=true then 
		form1.botton0.disabled=false
		form1.submit.disabled=false
	end if
end sub
sub botton0_onclick
	form1.botton1.disabled=false
	form1.botton2.disabled=true
	form1.botton0.disabled=true
	form1.submit.disabled=true
	form1.text0.disabled=false
	form1.text1.value=""
	form1.text2.value=""
end sub
sub submit_onclick

if (num1 mod 2)  and (num2 mod 2) then
    msgbox "依旧平常、身心无恙：掷到这种卦象的人最近不会有被误解的情况发生。"
elseif   (num1 mod 2)  and not (num2 mod 2) then
    msgbox "招来横祸、舆论不当：掷到这种卦象的人如果是平常人的话，在公众场所或左邻右舍不要做一些不当的行为以及发表一些不当的言语，而公众人物的话则要更为当心，免得遭受舆论攻击。"
  
elseif  not (num1 mod 2)  and (num2 mod 2) then
    msgbox "同侪之间、会有状况：掷到这种卦象的人要小心同辈、朋友或同事..等等之间产生误解误会。 "
else
    msgbox "亲情龃龉、沟通不良：掷到这种卦象的人不管是夫妻、父子、兄弟..等等亲人之间要小心会有口角争吵产生。"
end if
end sub
-->
    </SCRIPT>
                                    </head>
                                    <h2 align="center"><br>开运金钱卦</h2>
                                  <ol>
                                      <li>掷铜板之前要想题目，想好之后开始掷。</li>
                                    <li>准备一个铜板，5元、10元、20元都可以，当你掷出人头就是”正”，反面是字的话就是”反”，掷两次，然后照顺序开始排，例如正反、反正…等等</li>
                                  </ol>
                                  <hr size="1">
                                    <form action="syw-zxcqcd2.asp" method="post" name="form1" target="_self">
                                      <h3 align="center"><strong>请先输入所要求测的事情，再进行冥想，想好再按顺序開始擲</strong></h3>
                                      <p align="center">请输入所求测的事情：<strong>
                                        <input name="text0" type="text" id="text0" size="50" maxlength="50">
                                      </strong></p>
                                      <p align="center">
                                        <input name="botton1" type="button" disabled="disabled" id="botton1" value="按下开始擲第一個銅板">
                                        &nbsp;
                                        <input name="text1" type="text" id="text1">
                                      </p>
                                      <p align="center">
                                        <input name="botton2" type="button" disabled="disabled" id="botton2" value="按下开始擲第二個銅板">
                                        &nbsp;
                                        <input name="text2" type="text" id="text2">
                                      </p>
                                      <p align="center">
                                        <input name="Submit"  type="button" disabled="disabled" value="确认">
                                        <input name="botton0" type="button" disabled="disabled" id="botton0" value="重来">
                                      </p>
                                    </form>
</td></tr>
</table>
</td>
    </tr>
</table>

      <!--#INCLUDE FILE="../dq.asp"-->
	</div>
    </td>
    <td valign="top"><!--#INCLUDE FILE="../right.asp"--></td>
  </tr>
</table>
<!--#include file="../foot.asp"-->
</body></html>