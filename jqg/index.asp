<% 
const senlontitle="���˽�Ǯ��"
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>���˽�Ǯ��_ʵ�ò�ѯ���ߴ�ȫ - Powered by Senlon!</TITLE>

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
		msgbox "����������5���֣�"
	else
		form1.botton1.disabled=false
	end if
end sub
sub botton1_onclick
	randomize()
	num1=int(60*rnd())+1
	if num1 mod 2 then
		form1.text1.value="�������Ƿ�"
	else
	  form1.text1.value="����������"
	end if  
	form1.botton1.disabled=true
	form1.botton2.disabled=false
end sub
sub botton2_onclick
	randomize()
	if num2 mod 2 then
	 	form1.text2.value="�������Ƿ�"
	 else
	   form1.text2.value="����������"
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
    msgbox "����ƽ�����������������������������������б��������������"
elseif   (num1 mod 2)  and not (num2 mod 2) then
    msgbox "������������۲�������������������������ƽ���˵Ļ����ڹ��ڳ������������᲻Ҫ��һЩ��������Ϊ�Լ�����һЩ�������������������Ļ���Ҫ��Ϊ���ģ�����������۹�����"
  
elseif  not (num1 mod 2)  and (num2 mod 2) then
    msgbox "ͬ٭֮�䡢����״�������������������ҪС��ͬ�������ѻ�ͬ��..�ȵ�֮����������ᡣ "
else
    msgbox "������������ͨ��������������������˲����Ƿ��ޡ����ӡ��ֵ�..�ȵ�����֮��ҪС�Ļ��пڽ�����������"
end if
end sub
-->
    </SCRIPT>
                                    </head>
                                    <h2 align="center"><br>���˽�Ǯ��</h2>
                                  <ol>
                                      <li>��ͭ��֮ǰҪ����Ŀ�����֮��ʼ����</li>
                                    <li>׼��һ��ͭ�壬5Ԫ��10Ԫ��20Ԫ�����ԣ�����������ͷ���ǡ��������������ֵĻ����ǡ������������Σ�Ȼ����˳��ʼ�ţ������������������ȵ�</li>
                                  </ol>
                                  <hr size="1">
                                    <form action="syw-zxcqcd2.asp" method="post" name="form1" target="_self">
                                      <h3 align="center"><strong>����������Ҫ�������飬�ٽ���ڤ�룬����ٰ�˳���_ʼ�S</strong></h3>
                                      <p align="center">���������������飺<strong>
                                        <input name="text0" type="text" id="text0" size="50" maxlength="50">
                                      </strong></p>
                                      <p align="center">
                                        <input name="botton1" type="button" disabled="disabled" id="botton1" value="���¿�ʼ�S��һ���~��">
                                        &nbsp;
                                        <input name="text1" type="text" id="text1">
                                      </p>
                                      <p align="center">
                                        <input name="botton2" type="button" disabled="disabled" id="botton2" value="���¿�ʼ�S�ڶ����~��">
                                        &nbsp;
                                        <input name="text2" type="text" id="text2">
                                      </p>
                                      <p align="center">
                                        <input name="Submit"  type="button" disabled="disabled" value="ȷ��">
                                        <input name="botton0" type="button" disabled="disabled" id="botton0" value="����">
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