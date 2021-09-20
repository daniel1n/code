<% 
const senlontitle="网站域名吉凶测算"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>网站域名吉凶测算_实用查询工具大全 - Powered by Senlon!</TITLE>

</head>
<body>
<!--#INCLUDE FILE="../top.asp"-->
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top">
	<div id="index">
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <tr>
      <td width="70%" colspan="2" rowspan="2">　
  <div align="center">
    <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="356" id="AutoNumber1">
    <tr>
      <td width="356" align="center"><font size="3" color="#FF0000"><b>域名神算</b></font><form name="form1" method="POST" action="./index.asp?action=domain">
  <strong><font color="#FF0000">WWW.</font></strong><input type="text" name="domain" size="20">
  <input type="submit" value="分析域名" name="B1" tabindex="1">　</form></td>
    </tr>

<%
function result(s)
s=s mod 81
if s=0 then
s=80
end if
select case s
case 1
result="大殿鸿图，信用得固，无远无近，可获成功。（吉）"
case 2
result="根基不固，摇摇欲坠，一盛一衰，劳而无功。（凶）"
case 3
result="根深蒂固，蒸蒸日上，如意吉祥，百事顺道。（吉〕"

case 4
result="坎坷前途，若人折磨，非有毅力，难望成功。（凶）"

case 5
result="阴阳合和，生意兴隆，名利双收，后福重重。（吉）"

case 6
result="万宝集门，天降幸运，立志奋发，得成大功。（吉）"

case 7
result="独劳生意，和气致祥，排除万难，必获成功。（吉）"

case 8
result="努力发达，贯彻志望，不忘进退，可望成功。（吉）"

case 9
result="虽抱奇才，有才无命，独劳无力，财力难望。（凶）"

case 10
result="乌云遮月，暗淡无光，空费心力，徒劳无功。（凶）"

case 11
result="画草木逢春，枝中沾露，稳健着实，必得人望。（吉）"

case 12
result="薄弱无力，孤立无援。外祥内苦，谋事难成。（凶）"

case 13
result="天赋吉运，能得人望，善用智慧，必获成功。（吉〕"

case 14
result="忍得苦难，必有后福，是成是败，惟靠坚毅。（凶）"

case 15
result="廉恭做事，外得人和，大事成就，一门兴隆。（吉）"

case 16
result="能获众望，成就大业，名利双收，盟主四方。（吉）"

case 17
result="排除万难，有贵人助，把握时机，可得成功。（吉）"

case 18
result="经商做事，顺利昌隆，如难慎始，百事亨通。（吉）"

case 19
result="成功虽早，慎防亏空，内外不合，障碍重重。（凶）"

case 20
result="智高志大，历尽艰难，焦心忧劳，进退两难。（凶）"

case 21
result="先历困苦，后得幸福，霜雪梅花，春来怒放。（吉）"

case 22
result="秋草逢霜，怀才不遇，忧愁怨苦，事不如意。（凶）"

case 23
result="旭日升天，名显四方，渐次进展，终成大业。（吉）"

case 24
result="锦绣前程，须靠自力，多用智谋，能奏大功。（吉）"

case 25
result="天时地利，只欠人和，讲信修睦，即可成功。（吉）"

case 26
result="波澜起伏，千变万化，凌驾万难，必可成功。（凶带吉）"

case 27
result="一成一败。一盛一衰，惟靠谨慎，可守成功。（吉带凶）"
case 28
result="鱼临旱地，难逃恶运，此数大凶，不如更名。（凶）"

case 29
result="如龙得云，青云直上，智谋奋进，才略奏功。（吉）"

case 30
result="吉凶参半，得失相伴，投机取巧，如赌一样。（吉带凶）"

case 31
result="此数大吉，名利双收，渐进向上，大业成就。功。（吉）"

case 32
result="池中之龙，风云际会，一跃上天，成功可望。（吉）"

case 33
result="意气用事，人和必失，如难慎始，必可昌隆。（吉）"

case 34
result="灾难不绝，难望成功，此数大凶，不如更名。（凶）"
case 35
result="中吉之数，进退保守，生意安稳，成就普通。（吉）"
case 36
result="波澜重叠，常陷穷困，动不如静，有才无命。（凶）"

case 37
result="逢凶化吉，吉人天相，风调雨顺，生意兴隆。（吉）"

case 38
result="名虽可得，利则难获，艺界发展，可望成功。（凶带吉）"

case 39
result="云天见月，虽有劳碌，光明坦途，指日可望。（吉）"
case 40
result="一盛一衰，沉浮不定，知难而退，自获天佑。（吉带凶）"

case 41
result="天赋吉运，德望兼备，继续努力，前途无限。（吉）"

case 42
result="事业不专，十九不成，专心进取，可望成功。（吉带凶）"

case 43
result="雨夜之花，外祥内苦，忍耐自重，转凶为吉。（吉带凶）"

case 44
result="虽用心计，事难遂愿，贪功好进，必招失败。（凶）"

case 45
result="杨柳遇春，绿叶发枝，冲破难关，一举成名。（吉）"

case 46
result="坎坷不平，艰难重重，若无耐心，难望有成。（凶）"

case 47
result="有贵人助，可成大业，虽遇不幸，浮沉不定。（吉）"

case 48
result="美华丰实，鹤立鸡群，名利俱全，繁荣富贵。（吉）"
case 49
result="遇吉则吉，遇凶则凶，惟靠谨慎，逢凶化吉。（凶）"

case 50
result="吉凶互见，一成一败，凶中有吉，吉中有凶。（吉带凶）"

case 51
result="一盛一衰，浮沉不常，自重自处，可保平安。（吉带凶）"

case 52
result="草木逢春，雨过天晴，渡过难关，即获成功。（吉）"

case 53
result="盛衰参半，外祥内苦，先吉后凶，先凶后吉。（吉带凶）"

case 54
result="虽倾全力，难望成功，此数大凶，最好改名。（凶）"

case 55
result="外观昌隆，内隐祸患，克服难关，开出泰运。（吉带凶）"
case 56
result="事与愿违，终难成功，欲速不达，有始无终。（凶）"

case 57
result="虽有困难，时来运转，旷野枯草，春来花开。（凶带吉）"

case 58
result="半凶半吉，浮沉多端，始凶终吉，难保成功。（凶带吉）"

case 59
result="遇事猜疑，难望成功，大刀阔斧，始可有成。（凶）"

case 60
result="黑暗无光，心迷意乱，出尔反尔，难定方针。（凶）"

case 61
result="云遮半月，内隐内波，应自谨慎，始保平安。（吉带凶）"


case 62
result="烦闷懊恼，事业难展，自防灾祸，始免困境。（凶）"

case 63
result="万物化育，繁荣之象，专心一意，必能成功。（吉）"

case 64
result="见异思迁，十九不成，徒劳无功，不如更名。（凶）"

case 65
result="吉运自来，能享盛名，把握机会，必获成功。（吉）"

case 66
result=" 黑夜漫长，进退维谷，内外不和，信用缺乏。（凶）"

case 67
result="独营事业，事事如意，功成名就，富贵自来。（吉）"

case 68
result=" 思虑周祥，计划力行，不失先机，可望成功。（吉）"

case 69
result=" 动摇不安，常陷逆境，不得时运，难得利润。（凶）"

case 70
result=" 惨淡经营，难免贫困，此数不吉。最好改名。（凶）"

case 71
result=" 吉凶参半，惟赖勇气，贯彻力行，始可成功。（吉带凶）"

case 72
result="利害混合，凶手吉少，得而复失，难以安顿。（凶）"

case 73
result="安乐自来，自然吉祥，力行不懈，终必成功。（吉）"


case 74
result="利不及费，坐食山空，如无智谋，难望成功。（凶）"


case 75
result="吉中带凶，欲速下达，进不如守，可保安祥。（吉带凶）"

case 76
result="此数大凶，破产之象，宜速改名，以避厄运。（凶）"

case 77
result="先苦后甘，先甘后苦，如能守成，不致失败。（吉带凶）"

case 78
result="有得有失，华而下实，须防劫财，始保安顺。（吉带凶）"
case 79
result="如走夜路，前途无兴，希望不大，劳而无功。（凶）"

case 80
result="得而复失，枉费心机，守成无贪，可保安稳。（吉带凶）"

case 81
result="最极之数还本归元，能得繁荣，发达成功。（吉）"
case else
result="不要乱丢垃圾"
end select
end function

function CtoN(ch) 
select case Lcase(ch)  '转成小写，选出对应数值
case "1","2","3","4","5","6","7","8","9","0"
CtoN=Cint(ch)
case "a","e","f","g","h","n"
CtoN= 3
case "b","d","j","k","p","q","t","x","y"
CtoN=2
case "c","i","l","o","r","s","u","v","w","z"
CtoN=1
case "m"
CtoN=4
case ".","-"
CtoN=0
case else 
CtoN=100
end select
end function

function StoC()
s=0
Domain=trim(request.form("domain"))
%><tr><td align="center" width="356">
<br>

<font color=blue>
域名： <%=domain%></font>&nbsp;&nbsp;[<a href=http://<%=domain%> target="_blank" style="font-family: arial,sans-serif; font-size: 12px">访问该站</a>]</td></tr>
<%
intlen=Len(Domain)
for i=1 to intlen step 1
s=s+CtoN(mid(Domain,i,1))
next %>

<%if s>100 then%><tr><td align="center" width="356">
<font color=blue><b>请不要乱丢垃圾</b><br>
　</font></td></tr>

<%else%>
<tr><td align="center" width="356"><font color=red > <b><%=result(s)%></b></font></td></tr>

<%
end if
end function

if request("action")="domain" then
call stoc()
end if
%>

  </table>
  </div><br>
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