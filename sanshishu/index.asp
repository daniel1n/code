<% 
const senlontitle="������"
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>������,����Ԥ��_ʵ�ò�ѯ���ߴ�ȫ - Powered by Senlon!</TITLE>
<SCRIPT language=JavaScript>
<!--
function IsCn(Str)
{
    return /^[һ-��]+$/igm.test(Str);
}
function submitchecken() {
	if (document.cidu.youname1.value == "") {
		alert("���������ϡ�");
		document.cidu.youname1.focus();
		return false;
		}
	if (IsCn(document.cidu.youname1.value)==false) {
		alert("�������,Ӧ�ö�Ϊ���֡�");
		document.cidu.youname1.focus();
		return false;
		}
	if (document.cidu.youname1.value.length > 2) {
		alert("�����������,���ܶ���2�֡�");
		document.cidu.youname1.focus();
		return false;
		}

	if (document.cidu.youname2.value == "") {
		alert("����������");
		document.cidu.youname2.focus();
		return false;
		}
	if (IsCn(document.cidu.youname2.value)==false) {
		alert("�������,Ӧ�ö�Ϊ���֡�");
		document.cidu.youname2.focus();
		return false;
		}
	if (document.cidu.youname2.value.length > 2) {
		alert("�����������,���ܶ���2�֡�");
		document.cidu.youname2.focus();
		return false;
		}
	if (document.cidu.nn.value.length != 4) {
		alert("���λ��������,����Ϊ4λ��");
		document.cidu.nian.focus();
		return false;
		}
	
	re="���������룡";
 	y=document.cidu.nian.value;
 	m=document.cidu.yue.value;

 	h=document.cidu.hh.value;

	if (y==""||y<1901||y>2050) {
		alert("��Ӧ��1901��2050֮�䡣"+re);
		document.cidu.nian.focus();
		return false;
		}
	if (m>12||m<1) {
		alert("��Ӧ��1��12֮�䡣"+re);
		document.cidu.yue.focus();
		return false;
		}


	if (h>23||h<0) {
		alert("ʱӦ��0��23֮�䡣"+re);
		document.cidu.hh.focus();
		return false;
		}
	return true;
	}
//-->
</SCRIPT>
<!--#include file="../senlon/getuc.asp"-->

</head>
<body><!--#include file="../senlon/function.asp"--><!--#include file="../senlon/conn.asp"-->  
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
  <table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
<tbody><tr>
  <td width="100%" class="ttd"><span class="red">������-����Ԥ��</span>
    <br><br>�������Ǵ�˵��һ�����Կ�����ǰ���������ͺ������飬���������˵�һ�ж��������������ƪ��ר�Ŵ����澫ѡ�����������㵽��Ĳ��������������͵ģ���˵��׼��Ŷ������Ϳ�ʼ���������Լ��Ĳ��˰ɣ�<br><br></td>
</tr>
<form method="post" action="./index.asp" name=cidu  onsubmit="return submitchecken();">
  <tr>
    <td colspan="2" class="new"><div align=center>
  	�գ�<input type=txt name=youname1 size=4 value=""  onKeypress="if ((event.keyCode != 13 && event.keyCode < 160)) event.returnValue = false;">
  	����<input type=txt name=youname2 size=4 value=""  onKeypress="if ((event.keyCode != 13 && event.keyCode < 160)) event.returnValue = false;">
  	�Ա�<select name="sex" size="1" style="font-size: 9pt">
  	<option value="��" selected>��</option>
  	<option value="Ů" >Ů</option>
  	    </select>
  	������<a href="/wannianli/" target="_blank"><font color=red>ũ������</font></a>:
  	<input name="nx" size="1" type=hidden value="����"><select name="nn" size="1" id="nn" style="font-size: 9pt">            
        
        <option value="1900" >1900</option>           
        
        <option value="1901" >1901</option>           
        
        <option value="1902" >1902</option>           
        
        <option value="1903" >1903</option>           
        
        <option value="1904" >1904</option>           
        
        <option value="1905" >1905</option>           
        
        <option value="1906" >1906</option>           
        
        <option value="1907" >1907</option>           
        
        <option value="1908" >1908</option>           
        
        <option value="1909" >1909</option>           
        
        <option value="1910" >1910</option>           
        
        <option value="1911" >1911</option>           
        
        <option value="1912" >1912</option>           
        
        <option value="1913" >1913</option>           
        
        <option value="1914" >1914</option>           
        
        <option value="1915" >1915</option>           
        
        <option value="1916" >1916</option>           
        
        <option value="1917" >1917</option>           
        
        <option value="1918" >1918</option>           
        
        <option value="1919" >1919</option>           
        
        <option value="1920" >1920</option>           
        
        <option value="1921" >1921</option>           
        
        <option value="1922" >1922</option>           
        
        <option value="1923" >1923</option>           
        
        <option value="1924" >1924</option>           
        
        <option value="1925" >1925</option>           
        
        <option value="1926" >1926</option>           
        
        <option value="1927" >1927</option>           
        
        <option value="1928" >1928</option>           
        
        <option value="1929" >1929</option>           
        
        <option value="1930" >1930</option>           
        
        <option value="1931" >1931</option>           
        
        <option value="1932" >1932</option>           
        
        <option value="1933" >1933</option>           
        
        <option value="1934" >1934</option>           
        
        <option value="1935" >1935</option>           
        
        <option value="1936" >1936</option>           
        
        <option value="1937" >1937</option>           
        
        <option value="1938" >1938</option>           
        
        <option value="1939" >1939</option>           
        
        <option value="1940" >1940</option>           
        
        <option value="1941" >1941</option>           
        
        <option value="1942" >1942</option>           
        
        <option value="1943" >1943</option>           
        
        <option value="1944" >1944</option>           
        
        <option value="1945" >1945</option>           
        
        <option value="1946" >1946</option>           
        
        <option value="1947" >1947</option>           
        
        <option value="1948" >1948</option>           
        
        <option value="1949" >1949</option>           
        
        <option value="1950" >1950</option>           
        
        <option value="1951" >1951</option>           
        
        <option value="1952" >1952</option>           
        
        <option value="1953" >1953</option>           
        
        <option value="1954" >1954</option>           
        
        <option value="1955" >1955</option>           
        
        <option value="1956" >1956</option>           
        
        <option value="1957" >1957</option>           
        
        <option value="1958" >1958</option>           
        
        <option value="1959" >1959</option>           
        
        <option value="1960" >1960</option>           
        
        <option value="1961" >1961</option>           
        
        <option value="1962" >1962</option>           
        
        <option value="1963" >1963</option>           
        
        <option value="1964" >1964</option>           
        
        <option value="1965" >1965</option>           
        
        <option value="1966" >1966</option>           
        
        <option value="1967" >1967</option>           
        
        <option value="1968" >1968</option>           
        
        <option value="1969" >1969</option>           
        
        <option value="1970" >1970</option>           
        
        <option value="1971" >1971</option>           
        
        <option value="1972" >1972</option>           
        
        <option value="1973" >1973</option>           
        
        <option value="1974" >1974</option>           
        
        <option value="1975" >1975</option>           
        
        <option value="1976" >1976</option>           
        
        <option value="1977" >1977</option>           
        
        <option value="1978" >1978</option>           
        
        <option value="1979" >1979</option>           
        
        <option value="1980" >1980</option>           
        
        <option value="1981" >1981</option>           
        
        <option value="1982" >1982</option>           
        
        <option value="1983" >1983</option>           
        
        <option value="1984" >1984</option>           
        
        <option value="1985" >1985</option>           
        
        <option value="1986" >1986</option>           
        
        <option value="1987" >1987</option>           
        
        <option value="1988" >1988</option>           
        
        <option value="1989" >1989</option>           
        
        <option value="1990" >1990</option>           
        
        <option value="1991" >1991</option>           
        
        <option value="1992" >1992</option>           
        
        <option value="1993" >1993</option>           
        
        <option value="1994" >1994</option>           
        
        <option value="1995" >1995</option>           
        
        <option value="1996" >1996</option>           
        
        <option value="1997" >1997</option>           
        
        <option value="1998" >1998</option>           
        
        <option value="1999" >1999</option>           
        
        <option value="2000" >2000</option>           
        
        <option value="2001" >2001</option>           
        
        <option value="2002" >2002</option>           
        
        <option value="2003" >2003</option>           
        
        <option value="2004" >2004</option>           
        
        <option value="2005" >2005</option>           
        
        <option value="2006" >2006</option>           
        
        <option value="2007" >2007</option>           
        
        <option value="2008" >2008</option>           
        
        <option value="2009" >2009</option>           
        
        <option value="2010" >2010</option>           
        
        <option value="2011" >2011</option>           
        
        <option value="2012" >2012</option>           
        
        <option value="2013" >2013</option>           
        
        <option value="2014" >2014</option>           
        
        <option value="2015" >2015</option>           
        
        <option value="2016" >2016</option>           
        
        <option value="2017" >2017</option>           
        
        <option value="2018" >2018</option>           
        
        <option value="2019" >2019</option>           
        
        <option value="2020" >2020</option>           
        
        <option value="2021" >2021</option>           
        
        <option value="2022" >2022</option>           
        
        <option value="2023" >2023</option>           
        
        <option value="2024" >2024</option>           
        
        <option value="2025" >2025</option>           
        
        <option value="2026" >2026</option>           
        
        <option value="2027" >2027</option>           
        
        <option value="2028" >2028</option>           
        
        <option value="2029" >2029</option>           
        
        <option value="2030" >2030</option>           
        
        <option value="2031" >2031</option>           
        
        <option value="2032" >2032</option>           
        
        <option value="2033" >2033</option>           
        
        <option value="2034" >2034</option>           
        
        <option value="2035" >2035</option>           
        
        <option value="2036" >2036</option>           
        
        <option value="2037" >2037</option>           
        
        <option value="2038" >2038</option>           
        
        <option value="2039" >2039</option>           
        
        <option value="2040" >2040</option>           
        
        <option value="2041" >2041</option>           
        
        <option value="2042" >2042</option>           
        
        <option value="2043" >2043</option>           
        
        <option value="2044" >2044</option>           
        
        <option value="2045" >2045</option>           
        
        <option value="2046" >2046</option>           
        
        <option value="2047" >2047</option>           
        
        <option value="2048" >2048</option>           
        
        <option value="2049" >2049</option>           
        
        <option value="2050" >2050</option>           
        
        <option value="2051" >2051</option>           
        
        <option value="2052" >2052</option>           
        
        <option value="2053" >2053</option>           
        
        <option value="2054" >2054</option>           
        
        <option value="2055" >2055</option>           
        
        <option value="2056" >2056</option>           
        
        <option value="2057" >2057</option>           
        
        <option value="2058" >2058</option>           
        
        <option value="2059" >2059</option>           
        
        <option value="2060" >2060</option>           
        
        <option value="2061" >2061</option>           
        
        <option value="2062" >2062</option>           
        
        <option value="2063" >2063</option>           
        
        <option value="2064" >2064</option>           
        
        <option value="2065" >2065</option>           
        
        <option value="2066" >2066</option>           
        
        <option value="2067" >2067</option>           
        
        <option value="2068" >2068</option>           
        
        <option value="2069" >2069</option>           
        
        <option value="2070" >2070</option>           
        
        <option value="2071" >2071</option>           
        
        <option value="2072" >2072</option>           
        
        <option value="2073" >2073</option>           
        
        <option value="2074" >2074</option>           
        
        <option value="2075" >2075</option>           
        
        <option value="2076" >2076</option>           
        
        <option value="2077" >2077</option>           
        
        <option value="2078" >2078</option>           
        
        <option value="2079" >2079</option>           
        
        <option value="2080" >2080</option>           
        
        <option value="2081" >2081</option>           
        
        <option value="2082" >2082</option>           
        
        <option value="2083" >2083</option>           
        
        <option value="2084" >2084</option>           
        
        <option value="2085" >2085</option>           
        
        <option value="2086" >2086</option>           
        
        <option value="2087" >2087</option>           
        
        <option value="2088" >2088</option>           
        
        <option value="2089" >2089</option>           
        
        <option value="2090" >2090</option>           
        
        <option value="2091" >2091</option>           
        
        <option value="2092" >2092</option>           
        
        <option value="2093" >2093</option>           
        
        <option value="2094" >2094</option>           
        
        <option value="2095" >2095</option>           
        
        <option value="2096" >2096</option>           
        
        <option value="2097" >2097</option>           
        
        <option value="2098" >2098</option>           
        
        <option value="2099" >2099</option>           
        
        <option value="2100" >2100</option>           
        
        <option value="2101" >2101</option>           
        
        <option value="2102" >2102</option>           
        
        <option value="2103" >2103</option>           
        
        <option value="2104" >2104</option>           
        
        <option value="2105" >2105</option>           
        
        <option value="2106" >2106</option>           
        
        <option value="2107" >2107</option>           
        
        <option value="2108" >2108</option>           
        
        <option value="2109" >2109</option>           
        
        <option value="2110" >2110</option>           
        
        <option value="2111" >2111</option>           
        
        <option value="2112" >2112</option>           
        
        <option value="2113" >2113</option>           
        
        <option value="2114" >2114</option>           
        
        <option value="2115" >2115</option>           
        
        <option value="2116" >2116</option>           
        
        <option value="2117" >2117</option>           
        
        <option value="2118" >2118</option>           
        
        <option value="2119" >2119</option>           
        
        <option value="2120" >2120</option>           
        
        <option value="2121" >2121</option>           
        
        <option value="2122" >2122</option>           
        
        <option value="2123" >2123</option>           
        
        <option value="2124" >2124</option>           
        
        <option value="2125" >2125</option>           
        
        <option value="2126" >2126</option>           
        
        <option value="2127" >2127</option>           
        
        <option value="2128" >2128</option>           
        
        <option value="2129" >2129</option>           
        
        <option value="2130" >2130</option>           
        
        <option value="2131" >2131</option>           
        
        <option value="2132" >2132</option>           
        
        <option value="2133" >2133</option>           
        
        <option value="2134" >2134</option>           
        
        <option value="2135" >2135</option>           
        
        <option value="2136" >2136</option>           
        
        <option value="2137" >2137</option>           
        
        <option value="2138" >2138</option>           
        
        <option value="2139" >2139</option>           
        
        <option value="2140" >2140</option>           
        
        <option value="2141" >2141</option>           
        
        <option value="2142" >2142</option>           
        
        <option value="2143" >2143</option>           
        
        <option value="2144" >2144</option>           
        
        <option value="2145" >2145</option>           
        
        <option value="2146" >2146</option>           
        
        <option value="2147" >2147</option>           
        
        <option value="2148" >2148</option>           
        
        <option value="2149" >2149</option>           
        
        <option value="2150" >2150</option>           
        
        <option value="2151" >2151</option>           
        
        <option value="2152" >2152</option>           
        
        <option value="2153" >2153</option>           
        
        <option value="2154" >2154</option>           
        
        <option value="2155" >2155</option>           
        
        <option value="2156" >2156</option>           
        
        <option value="2157" >2157</option>           
        
        <option value="2158" >2158</option>           
        
      </select>��
      <select size="1" name="yue" style="font-size: 9pt">           
        
        <option value="1" >1</option>           
        
        <option value="2" >2</option>           
        
        <option value="3" >3</option>           
        
        <option value="4" >4</option>           
        
        <option value="5" >5</option>           
        
        <option value="6" >6</option>           
        
        <option value="7" >7</option>           
        
        <option value="8" >8</option>           
        
        <option value="9" >9</option>           
        
        <option value="10" >10</option>           
        
        <option value="11" >11</option>           
        
        <option value="12" >12</option>           
      </select>��
   <select size="1" name="ri" style="font-size: 9pt">         
        
        <option value="1" >1</option>           
        
        <option value="2" >2</option>           
        
        <option value="3" >3</option>           
        
        <option value="4" >4</option>           
        
        <option value="5" >5</option>           
        
        <option value="6" >6</option>           
        
        <option value="7" >7</option>           
        
        <option value="8" >8</option>           
        
        <option value="9" >9</option>           
        
        <option value="10" >10</option>           
        
        <option value="11" >11</option>           
        
        <option value="12" >12</option>           
        
        <option value="13" >13</option>           
        
        <option value="14" >14</option>           
        
        <option value="15" >15</option>           
        
        <option value="16" >16</option>           
        
        <option value="17" >17</option>           
        
        <option value="18" >18</option>           
        
        <option value="19" >19</option>           
        
        <option value="20" >20</option>           
        
        <option value="21" >21</option>           
        
        <option value="22" >22</option>           
        
        <option value="23" >23</option>           
        
        <option value="24" >24</option>           
        
        <option value="25" >25</option>           
        
        <option value="26" >26</option>           
        
        <option value="27" >27</option>           
        
        <option value="28" >28</option>           
        
        <option value="29" >29</option>           
        
        <option value="30" >30</option>           
        
        <option value="31" >31</option>           
    </select>��
      <select size="1" name="hh" style="font-size: 9pt">
      <option value=9 selected>�� 0</option> 
      <option value=8 >�� 1</option> 
      <option value=8 >�� 2</option>
      <option value=7 >�� 3</option> 
      <option value=7 >�� 4</option> 
      <option value=6 >î 5</option> 
      <option value=6 >î 6</option> 
      <option value=5 >�� 7</option> 
      <option value=5 >�� 8</option> 
      <option value=4 >�� 9</option> 
      <option value=4 >�� 10</option> 
      <option value=9 >�� 11</option> 
      <option value=9 >�� 12</option> 
      <option value=8 >δ 13</option> 
      <option value=8 >δ 14</option> 
      <option value=7 >�� 15</option> 
      <option value=7 >�� 16</option> 
      <option value=6 >�� 17</option> 
      <option value=6 >�� 18</option> 
      <option value=5 >�� 19</option> 
      <option value=5 >�� 20</option> 
      <option value=4 >�� 21</option> 
      <option value=4 >�� 22</option> 
      <option value=9 >�� 23</option> 
      </select>��
      <input type="submit" name="btnAdd" value="����"></td>
    </tr>
<%
nn=Right(request("nn"), 1)
yue=request("yue")
sql="select * from sanshishu where nn='"&nn&"' "
set rs=conn.execute(sql)
if not (rs.bof and rs.eof) then
cai=rs(""&yue&"")

Select Case cai
           Case "��"    
		    content="��»���ڼ�ɸ�"
		    content1="��ҵ��»ʮ����������һ��һ���迿����׬ȡ�����ǲ��˼��ѣ�������Կ��ڿ˼󣬲��Ź��������������ٳɶ࣬���ﶼ����֪��Ψ����ԵǷ˳�����ܲ�ֹһ�μ�Ȣ��"
           Case "��"  
		    content="��»���º�����"
		    content1="��Ǯ��ȱ���䲻�ǳ�����������Ҳ��һ���Ʋ���ԣ��ȱ����ϲ��׷����������Զ̰��֪�㣬���ַǳ����ģ����ض��Ǯȴ����û��ѣ����º��������������ǵ�д��"
           Case "��"  
		     content="��»�������"
		     content1="ע��������µ�����ı������ǰ�й��������������飬�Ȳ����ѲŻ�����𾮡���ʱ��ת�䣬��Χ�߲ŵ��ķ��ƣ��б߸��̼Ҳ��Ǵ�ɻ������·����֮׬��Ǯ�������"  
		   Case "��"  
		     content="��»��������ȥ"
		     content1="һ�����ж�δ�ʽ�Ǯ���ֵĻ��ᣬ��ʹ��ˣ�Ҳ�������Ǹ���������ΪǮ�����ÿ�ȥ�ÿ죬�粻��������ƣ����ǧ��ɢ�������¿�������"  
           Case "��"  
		     content="��»��֪�㳣��"
		     content1="��ν��֪����ƶ���֡������۳����Ͳ�»Ҳƽƽ���棬��֪��ĸ��������ǲ���Ϊ�ࡣ�����˸�������������ӵ�����ǵ����꣬�������������"  
		   Case "��"  
		     content="��»��������"
		     content1="ǰ�����ú��¶࣬�������ڸ�ԣ��������������������ʵ��ס�����׳���������ֻҪ��ס��ҵ��һ����ʳ���ǣ�����̫������Ӫ��ƣ���������ʧ�ܡ�"  
           Case "��"  
		     content="��»�������޸�"
		     content1="ʮ����»�У�����»��á����ơ�������ɼ�ã�һ�����Ǯ���²���Ե����ҵ���ϰ��Ȼ�л���ɾ޸�����ʹ���˴��������˳��۵ġ��򹤻ʵۡ����뿪����������ط�չ���ɾ͸���"  
		   Case "ɷ"  
		     content="ɷ»��������µ"
		     content1="����������ֻ������Ѷϵ��Ը����ۣ������ֲ��������������¾���ʱ��Ϊʱ�������������ʱ������ٷ������������ȷ�����ѧһ�ż��ܰ���Ŭ��������Ȼ��ʳ���ǡ�"  
		   Case "��"  
		     content="��»��������Ϊ"
		     content1="���˲�����ʧ�����õĽ�Ǯ�븶����Ŭ�������ȣ���׬Ǯ����������Ϊ������»���ˣ�һ����˴�������ѧʶ�����˴���רҵ��ͷ������ҽ������ʦ�������ҵȣ��°���������ó�ԣ��"  
		   Case "��"  
		     content="��»���󸻴��"
		     content1="�Ƹ��ĵ�ס����»����ͬ���󸻴�������¸�֮��������ͻ��������Ͷ�ʸ߷�����Ŀ���������󷴶�������Դ�����������Ҫ�ɹ��˴��������������˼�������罻Ȧ�ӡ�"  
           Case "��"  
		     content="��»���ȿ����"
		     content1="��ǰ��Ӱ�죬����������Ѻ����ദ�����������Կ����󷢴�����ȿ���������������¸��������Ҹ������õ��ĵ������⡣�����ѣ����䲻�ٽ�Ǯ�����ٶ�������"  
		   Case "��"  
		     content="��»����ʳ����"
		     content1="�㲢������ѷ���ˣ�ֻϵ�Լ���ʳ���������������Ŷ��װ���ʧ���뷢��Ļ��м�̰ͼ���ݣ���ν����������Ϊ�ơ������⣬�����������ʽϲû���������ǣ����������ⱻ������"  
 End Select
%>

<tr bgcolor="#EFF8FE">
<td class="new" colspan="2" valign="middle">
<%="<br><br>&nbsp;Ԥ�������£�<br><br>&nbsp;��Ĳ�����<font size='6' color=red>��"&content&"��</font><br><br>&nbsp;&nbsp;&nbsp;&nbsp;"&content1%>

</td>
</tr><%end if%>
</tbody>
</table>
      <!--#INCLUDE FILE="../dq.asp"-->
	</div>
    </td>
    <td valign="top"><!--#INCLUDE FILE="../right.asp"--></td>
  </tr>
</table>
<!--#include file="../foot.asp"-->
</body></html>