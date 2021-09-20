<%
'»•≥˝ ‰»Î◊÷∑˚÷–µƒø’∏Ò---newtrim∫Ø ˝
function newtrim(str)
dim result
dim j
j=len(str)
result=""
dim i
for i=1 to j
select case mid(str,i,1)
case chr(8)
     result=result+"" '»•µÙÕÀ∏Ò
case chr(9)
     result=result+"" '»•µÙtab
case chr(10)
     result=result+"" '»•µÙªª––
case chr(13)
     result=result+"" '»•µÙªÿ≥µ
case chr(255)
     result=result+"" '»•µÙÃÿ ‚ø’∏Ò
case else
     result=result+mid(str,i,1)
end select
next
newtrim=result
end function
%>
<%
'µ√µΩ“ª∏ˆ∫∫◊÷µƒ± ª≠ ˝---getnum()∫Ø ˝
function getnum(txt)
dim sql1,bihua
sql1="select * from bihua"
set rs1=conn.execute(sql1)
  if not (rs1.bof and rs1.eof) then
    do while not rs1.eof
      if instr(1,rs1("hanzi"),txt,1) >0 then
        bihua=int(rs1("num"))
	  exit do
      else 
      rs1.movenext
      end if
    loop
end if
getnum=bihua
end function
%>
<%
'”…ÃÏ∏Ò°¢»À∏Ò°¢µÿ∏Ò ˝µ√µΩ»˝≤≈---getsancai()∫Ø ˝
function getsancai(tiange,renge,dige)
dim tian,ren,di,tiantxt,rentxt,ditxt,total
tian=(tiange mod 10)
ren=(renge mod 10)
di=(dige mod 10)
select case tian
  case 1
  tiantxt="ƒæ"
  case 2
  tiantxt="ƒæ"
  case 3
  tiantxt="ª"
  case 4
  tiantxt="ª"
  case 5
  tiantxt="Õ¡"
  case 6
  tiantxt="Õ¡"
  case 7
  tiantxt="Ω"
  case 8
  tiantxt="Ω"
  case 9
  tiantxt="ÀÆ"
  case 10
  tiantxt="ÀÆ"
  case 0
  tiantxt="ÀÆ"
end select
select case ren
  case 1
  rentxt="ƒæ"
  case 2
  rentxt="ƒæ"
  case 3
  rentxt="ª"
  case 4
  rentxt="ª"
  case 5
  rentxt="Õ¡"
  case 6
  rentxt="Õ¡"
  case 7
  rentxt="Ω"
  case 8
  rentxt="Ω"
  case 9
  rentxt="ÀÆ"
  case 10
  rentxt="ÀÆ"
  case 0
  rentxt="ÀÆ"
end select
select case di
  case 1
  ditxt="ƒæ"
  case 2
  ditxt="ƒæ"
  case 3
  ditxt="ª"
  case 4
  ditxt="ª"
  case 5
  ditxt="Õ¡"
  case 6
  ditxt="Õ¡"
  case 7
  ditxt="Ω"
  case 8
  ditxt="Ω"
  case 9
  ditxt="ÀÆ"
  case 10
  ditxt="ÀÆ"
  case 0
  ditxt="ÀÆ"
end select
total=tiantxt+rentxt+ditxt
getsancai=total
end function
%>
<%
Function hhcal(caltime) 
'π´¿˙◊™ªª≈©¿˙µƒÀ„∑® edit by huanghui
Dim WeekName(7), MonthAdd(11), NongliData(99), TianGan(9), DiZhi(11), ShuXiang(11), DayName(30), MonName(12) 
Dim curTime, curYear, curMonth, curDay, curWeekday 
Dim GongliStr, WeekdayStr, NongliStr, NongliDayStr 
Dim i, m, n, k, isEnd, bit, TheDate 
'ªÒ»°µ±«∞œµÕ≥ ±º‰ 
curTime = caltime 
'–«∆⁄√˚ 
WeekName(0) = " * " 
WeekName(1) = "–«∆⁄»’" 
WeekName(2) = "–«∆⁄“ª" 
WeekName(3) = "–«∆⁄∂˛" 
WeekName(4) = "–«∆⁄»˝" 
WeekName(5) = "–«∆⁄Àƒ" 
WeekName(6) = "–«∆⁄ŒÂ" 
WeekName(7) = "–«∆⁄¡˘" 
'ÃÏ∏…√˚≥∆ 
TianGan(0) = "º◊" 
TianGan(1) = "““" 
TianGan(2) = "±˚" 
TianGan(3) = "∂°" 
TianGan(4) = "ŒÏ" 
TianGan(5) = "º∫" 
TianGan(6) = "∏˝" 
TianGan(7) = "–¡" 
TianGan(8) = "»…" 
TianGan(9) = "πÔ" 
'µÿ÷ß√˚≥∆ 
DiZhi(0) = "◊”" 
DiZhi(1) = "≥Û" 
DiZhi(2) = "“˙" 
DiZhi(3) = "√Æ" 
DiZhi(4) = "≥Ω" 
DiZhi(5) = "À»" 
DiZhi(6) = "ŒÁ" 
DiZhi(7) = "Œ¥" 
DiZhi(8) = "…Í" 
DiZhi(9) = "”œ" 
DiZhi(10) = "–Á" 
DiZhi(11) = "∫•" 
' Ùœ‡√˚≥∆ 
ShuXiang(0) = " Û" 
ShuXiang(1) = "≈£" 
ShuXiang(2) = "ª¢" 
ShuXiang(3) = "Õ√" 
ShuXiang(4) = "¡˙" 
ShuXiang(5) = "…ﬂ" 
ShuXiang(6) = "¬Ì" 
ShuXiang(7) = "—Ú" 
ShuXiang(8) = "∫Ô" 
ShuXiang(9) = "º¶" 
ShuXiang(10) = "π∑" 
ShuXiang(11) = "÷Ì" 
'≈©¿˙»’∆⁄√˚ 
DayName(0) = "*" 
DayName(1) = "≥ı“ª" 
DayName(2) = "≥ı∂˛" 
DayName(3) = "≥ı»˝" 
DayName(4) = "≥ıÀƒ" 
DayName(5) = "≥ıŒÂ" 
DayName(6) = "≥ı¡˘" 
DayName(7) = "≥ı∆ﬂ" 
DayName(8) = "≥ı∞À" 
DayName(9) = "≥ıæ≈" 
DayName(10) = "≥ı Æ" 
DayName(11) = " Æ“ª" 
DayName(12) = " Æ∂˛" 
DayName(13) = " Æ»˝" 
DayName(14) = " ÆÀƒ" 
DayName(15) = " ÆŒÂ" 
DayName(16) = " Æ¡˘" 
DayName(17) = " Æ∆ﬂ" 
DayName(18) = " Æ∞À" 
DayName(19) = " Ææ≈" 
DayName(20) = "∂˛ Æ" 
DayName(21) = "ÿ•“ª" 
DayName(22) = "ÿ•∂˛" 
DayName(23) = "ÿ•»˝" 
DayName(24) = "ÿ•Àƒ" 
DayName(25) = "ÿ•ŒÂ" 
DayName(26) = "ÿ•¡˘" 
DayName(27) = "ÿ•∆ﬂ" 
DayName(28) = "ÿ•∞À" 
DayName(29) = "ÿ•æ≈" 
DayName(30) = "»˝ Æ" 
'≈©¿˙‘¬∑›√˚ 
MonName(0) = "*" 
MonName(1) = "’˝" 
MonName(2) = "∂˛" 
MonName(3) = "»˝" 
MonName(4) = "Àƒ" 
MonName(5) = "ŒÂ" 
MonName(6) = "¡˘" 
MonName(7) = "∆ﬂ" 
MonName(8) = "∞À" 
MonName(9) = "æ≈" 
MonName(10) = " Æ" 
MonName(11) = " Æ“ª" 
MonName(12) = "¿∞" 
'π´¿˙√ø‘¬«∞√ÊµƒÃÏ ˝ 
MonthAdd(0) = 0 
MonthAdd(1) = 31 
MonthAdd(2) = 59 
MonthAdd(3) = 90 
MonthAdd(4) = 120 
MonthAdd(5) = 151 
MonthAdd(6) = 181 
MonthAdd(7) = 212 
MonthAdd(8) = 243 
MonthAdd(9) = 273 
MonthAdd(10) = 304 
MonthAdd(11) = 334 
'≈©¿˙ ˝æ› 
NongliData(0) = 2635 
NongliData(1) = 333387 
NongliData(2) = 1701 
NongliData(3) = 1748 
NongliData(4) = 267701 
NongliData(5) = 694 
NongliData(6) = 2391 
NongliData(7) = 133423 
NongliData(8) = 1175 
NongliData(9) = 396438 
NongliData(10) = 3402 
NongliData(11) = 3749 
NongliData(12) = 331177 
NongliData(13) = 1453 
NongliData(14) = 694 
NongliData(15) = 201326 
NongliData(16) = 2350 
NongliData(17) = 465197 
NongliData(18) = 3221 
NongliData(19) = 3402 
NongliData(20) = 400202 
NongliData(21) = 2901 
NongliData(22) = 1386 
NongliData(23) = 267611 
NongliData(24) = 605 
NongliData(25) = 2349 
NongliData(26) = 137515 
NongliData(27) = 2709 
NongliData(28) = 464533 
NongliData(29) = 1738 
NongliData(30) = 2901 
NongliData(31) = 330421 
NongliData(32) = 1242 
NongliData(33) = 2651 
NongliData(34) = 199255 
NongliData(35) = 1323 
NongliData(36) = 529706 
NongliData(37) = 3733 
NongliData(38) = 1706 
NongliData(39) = 398762 
NongliData(40) = 2741 
NongliData(41) = 1206 
NongliData(42) = 267438 
NongliData(43) = 2647 
NongliData(44) = 1318 
NongliData(45) = 204070 
NongliData(46) = 3477 
NongliData(47) = 461653 
NongliData(48) = 1386 
NongliData(49) = 2413 
NongliData(50) = 330077 
NongliData(51) = 1197 
NongliData(52) = 2637 
NongliData(53) = 268877 
NongliData(54) = 3365 
NongliData(55) = 531109 
NongliData(56) = 2900 
NongliData(57) = 2922 
NongliData(58) = 398042 
NongliData(59) = 2395 
NongliData(60) = 1179 
NongliData(61) = 267415 
NongliData(62) = 2635 
NongliData(63) = 661067 
NongliData(64) = 1701 
NongliData(65) = 1748 
NongliData(66) = 398772 
NongliData(67) = 2742 
NongliData(68) = 2391 
NongliData(69) = 330031 
NongliData(70) = 1175 
NongliData(71) = 1611 
NongliData(72) = 200010 
NongliData(73) = 3749 
NongliData(74) = 527717 
NongliData(75) = 1452 
NongliData(76) = 2742 
NongliData(77) = 332397 
NongliData(78) = 2350 
NongliData(79) = 3222 
NongliData(80) = 268949 
NongliData(81) = 3402 
NongliData(82) = 3493 
NongliData(83) = 133973 
NongliData(84) = 1386 
NongliData(85) = 464219 
NongliData(86) = 605 
NongliData(87) = 2349 
NongliData(88) = 334123 
NongliData(89) = 2709 
NongliData(90) = 2890 
NongliData(91) = 267946 
NongliData(92) = 2773 
NongliData(93) = 592565 
NongliData(94) = 1210 
NongliData(95) = 2651 
NongliData(96) = 395863 
NongliData(97) = 1323 
NongliData(98) = 2707 
NongliData(99) = 265877 
'…˙≥…µ±«∞π´¿˙ƒÍ°¢‘¬°¢»’ ==> GongliStr 
curYear = Year(curTime) 
curMonth = Month(curTime) 
curDay = Day(curTime) 
GongliStr = curYear & "ƒÍ" 
If (curMonth < 10) Then 
    GongliStr = GongliStr & "0" & curMonth & "‘¬" 
Else 
    GongliStr = GongliStr & curMonth & "‘¬" 
End If 
If (curDay < 10) Then 
    GongliStr = GongliStr & "0" & curDay & "»’" 
Else 
    GongliStr = GongliStr & curDay & "»’" 
End If 
'…˙≥…µ±«∞π´¿˙–«∆⁄ ==> WeekdayStr 
curWeekday = Weekday(curTime) 
WeekdayStr = WeekName(curWeekday) 
'º∆À„µΩ≥ı º ±º‰1921ƒÍ2‘¬8»’µƒÃÏ ˝£∫1921-2-8(’˝‘¬≥ı“ª) 
TheDate = (curYear - 1921) * 365 + Int((curYear - 1921) / 4) + curDay + MonthAdd(curMonth - 1) - 38 
If ((curYear Mod 4) = 0 And curMonth > 2) Then 
    TheDate = TheDate + 1 
End If 
'º∆À„≈©¿˙ÃÏ∏…°¢µÿ÷ß°¢‘¬°¢»’ 
isEnd = 0 
m = 0 
Do 
    If (NongliData(m) < 4095) Then 
        k = 11 
    Else 
        k = 12 
    End If 
    n = k 
    Do 
        If (n < 0) Then 
            Exit Do 
        End If 
    'ªÒ»°NongliData(m)µƒµ⁄n∏ˆ∂˛Ω¯÷∆Œªµƒ÷µ 
    bit = NongliData(m) 
    For i = 1 To n Step 1 
        bit = Int(bit / 2) 
    Next 
    bit = bit Mod 2 
    If (TheDate <= 29 + bit) Then 
        isEnd = 1 
        Exit Do 
    End If 
    TheDate = TheDate - 29 - bit 
    n = n - 1 
  Loop 
  If (isEnd = 1) Then 
      Exit Do 
  End If 
  m = m + 1 
Loop 
curYear = 1921 + m 
curMonth = k - n + 1 
curDay = TheDate 
If (k = 12) Then 
    If (curMonth = (Int(NongliData(m) / 65536) + 1)) Then 
        curMonth = 1 - curMonth 
    ElseIf (curMonth > (Int(NongliData(m) / 65536) + 1)) Then 
        curMonth = curMonth - 1 
    End If 
End If 
'…˙≥…≈©¿˙ÃÏ∏…°¢µÿ÷ß°¢ Ùœ‡ ==> NongliStr
dim nongli,nlnian,nlyue,nlri,sx
nlnian = TianGan(((curYear - 4) Mod 60) Mod 10) & DiZhi(((curYear - 4) Mod 60) Mod 12) 
sx =ShuXiang(((curYear - 4) Mod 60) Mod 12)
'…˙≥…≈©¿˙‘¬°¢»’ ==> NongliDayStr 
If (curMonth < 1) Then 
    nlyue = MonName(-1 * curMonth) 
Else 
    nlyue = MonName(curMonth) 
End If 
nlri = DayName(curDay)
nonglistr=nlnian&"|"&nlyue&"|"&nlri&"|"&sx
hhcal=nonglistr
End Function
%><%Function Constellation(mDate)
if not isdate(mDate) then response.write("∑«»’∆⁄")
dim a
a=(Day(mDate) - (19 + Int(Mid("102123444423", Month(mDate), 1))))
if a>=0 then
a=0
else
a=-1
End if
'–«◊˘
    Constellation = Mid("ƒßÙ…ÀÆ∆øÀ´”„∞◊—ÚΩ≈£À´◊”æﬁ–∑ ®◊”¥¶≈ÆÃÏ≥”ÃÏ–´…‰ ÷ƒßÙ…", (Month(mDate) + a) * 2 + 1, 2) & "◊˘"
End Function%>
<%
'ÃÊªªªª––∑˚µƒπ˝¬À∫Ø ˝
Function rep(content)
rep=replace(content,vbCrLf,"<br>")
end Function
%><%function GbToBig(content)
dim s,t,c,d,i
s="∞®,∞™,∞≠,∞Æ,∞ø,∞¿,∞¬,∞”,∞’,∞⁄,∞‹,∞‰,∞Ï,∞Ì,∞Ô,∞Û,∞˜,∞˘,∞˛,±•,±¶,±®,±´,±≤,±¥,±µ,±∑,±∏,±π,±¡,± ,±œ,±–,±’,±ﬂ,±‡,±·,±‰,±Á,±Ë,±Ó,±Ò,±Ù,±ı,±ˆ,±˜,±˝,≤¶,≤ß,≤¨,≤µ,≤∑,≤π,≤Œ,≤œ,≤–,≤—,≤“,≤”,≤‘,≤’,≤÷,≤◊,≤ﬁ,≤‡,≤·,≤‚,≤„,≤Ô,≤Û,≤Ù,≤ı,≤ˆ,≤˜,≤¯,≤˘,≤˙,≤˚,≤¸,≥°,≥¢,≥§,≥•,≥¶,≥ß,≥©,≥Æ,≥µ,≥π,≥æ,≥¬,≥ƒ,≥≈,≥∆,≥Õ,≥œ,≥“,≥’,≥Ÿ,≥€,≥‹,≥›,≥„,≥Â,≥Ê,≥Ë,≥Î,≥Ï,≥Ô,≥Ò,≥Û,≥˜,≥¯,≥˙,≥˚,¥°,¥¢,¥•,¥¶,¥´,¥Ø,¥≥,¥¥,¥∏,¥ø,¥¬,¥«,¥ ,¥Õ,¥œ,¥–,¥—,¥”,¥‘,¥’,¥‹,¥Ì,¥Ô,¥¯,¥˚,µ£,µ•,µ¶,µß,µ®,µ¨,µÆ,µØ,µ±,µ≤,µ≥,µ¥,µµ,µ∑,µ∫,µª,µº,µ¡,µ∆,µÀ,µ–,µ”,µ›,µﬁ,µ„,µÊ,µÁ,µÌ,µˆ,µ˜,µ¸,µ˝,µ˛,∂§,∂•,∂ß,∂©,∂´,∂Ø,∂∞,∂≥,∂∑,∂ø,∂¿,∂¡,∂ƒ,∂∆,∂Õ,∂œ,∂–,∂“,∂”,∂‘,∂÷,∂Ÿ,∂€,∂·,∂Ï,∂Ó,∂Ô,∂Ò,∂ˆ,∂˘,∂˚,∂¸,∑°,∑¢,∑£,∑ß,∑©,∑Ø,∑∞,∑≥,∑∂,∑∑,∑π,∑√,∑ƒ,∑…,∑œ,∑—,∑◊,∑ÿ,∑‹,∑ﬂ,∑‡,∑·,∑„,∑Ê,∑Á,∑Ë,∑Î,∑Ï,∑Ì,∑Ô,∑Ù,∑¯,∏ß,∏®,∏≥,∏¥,∏∫,∏º,∏æ,∏ø,∏√,∏∆,∏«,∏…,∏œ,∏—,∏”,∏‘,∏’,∏÷,∏Ÿ,∏⁄,∏ﬁ,∏‰,∏È,∏Î,∏Û,∏ı,∏ˆ,∏¯,π®,π¨,πÆ,π±,π≥,πµ,ππ,π∫,πª,π∆,πÀ,π–,πÿ,π€,π›,πﬂ,π·,π„,πÊ,πË,πÈ,πÍ,πÎ,πÏ,πÓ,πÒ,πÛ,πÙ,πı,πˆ,π¯,π˙,π˝,∫ß,∫´,∫∫,∫“,∫◊,∫ÿ,∫·,∫‰,∫Ë,∫Ï,∫Û,∫¯,ª§,ª¶,ªß,ª©,ª™,ª≠,ªÆ,ª∞,ª≥,ªµ,ª∂,ª∑,ªπ,ª∫,ªª,ªΩ,ªæ,ª¿,ª¡,ª∆,ª—,ª”,ª‘,ªŸ,ªﬂ,ª‡,ª·,ª‚,ª„,ª‰,ªÂ,ªÊ,ªÁ,ªÎ,ªÔ,ªÒ,ªı,ªˆ,ª˜,ª˙,ª˝,º¢,º•,º¶,º®,º©,º´,º≠,º∂,º∑,º∏,ºª,º¡,º√,º∆,º«,º ,ºÃ,ºÕ,º–,º‘,º’,º÷,ºÿ,º€,º›,ºﬂ,º‡,º·,º„,º‰,ºË,ºÍ,ºÎ,ºÏ,ºÓ,ºÔ,º,ºÒ,ºÚ,ºÛ,ºı,ºˆ,º˜,º¯,º˘,º˙,º˚,º¸,Ω¢,Ω£,Ω§,Ω•,Ω¶,Ωß,Ω¨,ΩØ,Ω∞,Ω±,Ω≤,Ω¥,Ω∫,ΩΩ,Ωæ,Ωø,Ω¡,Ω¬,Ω√,Ωƒ,Ω≈,Ω»,Ω…,Ω ,ΩŒ,Ωœ,Ω’,Ω◊,Ω⁄,æ•,æ™,æ≠,æ±,æ≤,æµ,æ∂,æ∑,æ∫,æª,æ¿,æ«,æ…,æ‘,æŸ,æ›,æ‚,æÂ,æÁ,æÈ,æÓ,Ω‹,Ω‡,Ω·,ΩÎ,ΩÏ,ΩÙ,Ωı,Ωˆ,Ω˜,Ω¯,Ω˙,Ω˝,æ°,æ¢,æ£,æı,æˆ,æ˜,æ¯,æ˚,æ¸,ø•,ø™,ø≠,ø≈,ø«,øŒ,ø—,ø“,øŸ,ø‚,ø„,ø‰,øÈ,øÎ,øÌ,øÛ,øı,øˆ,ø˜,ø˘,ø˙,¿°,¿£,¿©,¿´,¿Ø,¿∞,¿≥,¿¥,¿µ,¿∂,¿∏,¿π,¿∫,¿ª,¿º,¿Ω,¿æ,¿ø,¿¿,¿¡,¿¬,¿√,¿ƒ,¿Ã,¿Õ,¿‘,¿÷,¿ÿ,¿›,¿‡,¿·,¿È,¿Î,¿Ô,¿,¿Ò,¿ˆ,¿˜,¿¯,¿˘,¿˙,¡§,¡•,¡©,¡™,¡´,¡¨,¡≠,¡Ø,¡∞,¡±,¡≤,¡≥,¡¥,¡µ,¡∂,¡∑,¡∏,¡π,¡Ω,¡æ,¡¬,¡∆,¡…,¡Õ,¡‘,¡Ÿ,¡⁄,¡€,¡›,¡ﬁ,¡‰,¡Â,¡Ë,¡È,¡Î,¡Ï,¡Û,¡ı,¡˙,¡˚,¡¸,¡˝,¬¢,¬£,¬§,¬•,¬¶,¬ß,¬®,¬´,¬¨,¬≠,¬Æ,¬Ø,¬∞,¬±,¬≤,¬≥,¬∏,¬ª,¬º,¬Ω,¬ø,¬¿,¬¡,¬¬,¬≈,¬∆,¬«,¬À,¬Ã,¬Õ,¬Œ,¬œ,¬–,¬“,¬’,¬÷,¬◊,¬ÿ,¬Ÿ,¬⁄,¬€,¬‹,¬ﬁ,¬ﬂ,¬‡,¬·,¬‚,¬Ê,¬Á,¬Ë,¬Í,¬Î,¬Ï,¬Ì,¬Ó,¬,¬Ú,¬Û,¬Ù,¬ı,¬ˆ,¬˜,¬¯,¬˘,¬˙,√°,√®,√™,√≠,√≥,√¥,√π,√ª,√æ,√≈,√∆,√«,√Ã,√Œ,√’,√÷,√Ÿ,√‡,√Â,√Ì,√,√ı,√ˆ,√˘,√˙,√˝,ƒ±,ƒ∂,ƒ∆,ƒ…,ƒ—,ƒ”,ƒ‘,ƒ’,ƒ÷,ƒŸ,ƒÂ,ƒÏ,ƒÌ,ƒ,ƒÒ,ƒÙ,ƒˆ,ƒ˜,ƒ¯,ƒ˚,ƒ¸,ƒ˛,≈°,≈¢,≈•,≈¶,≈ß,≈®,≈©,≈±,≈µ,≈∑,≈∏,≈π,≈ª,≈Ω,≈Ã,≈”,π˙,∞Æ,≈‚,≈Á,≈Ù,∆≠,∆Æ,∆µ,∆∂,∆ª,∆æ,∆¿,∆√,∆ƒ,∆À,∆Ã,∆”,∆◊,∆Í,∆Î,∆Ô,∆Ò,∆Ù,∆¯,∆˙,∆˝,«£,«§,«•,«¶,«®,«©,«´,«Æ,«Ø,«±,«≥,«¥,«µ,«π,«∫,«Ω,«æ,«ø,«¿,«¬,«≈,««,«»,«Ã,«œ,«‘,«’,«◊,«·,«‚,«„,«Í,«Î,«Ï,«Ì,«Ó,«˜,«¯,«˚,«˝,»£,»ß,»®,»∞,»¥,»µ,»√,»ƒ,»≈,»∆,»»,»Õ,»œ,»“,»Ÿ,»ﬁ,»Ì,»Ò,»Ú,»Û,»˜,»¯,»˙,»¸,…°,…•,…ß,…®,…¨,…±,…¥,…∏,…π,…¡,…¬,…ƒ,……,…À,…Õ,…’,…‹,…ﬁ,…„,…Â,…Ë,…,…Û,…Ù,…ˆ,…¯,…˘,…˛, §, •, ¶, ®, ™, ´, ¨, ±, ¥, µ, ∂, ª, ∆, Õ, Œ, ”, ‘, Ÿ, ﬁ, ‡, ‰, È, Í, Ù, ı, ˜, ˙, ˝,Àß,À´,À≠,À∞,À≥,Àµ,À∂,À∏,Àø,À«,À ,ÀÀ,ÀÃ,Àœ,À–,À”,À’,Àﬂ,À‡,À‰,ÀÁ,ÀÍ,ÀÔ,À,ÀÒ,Àı,Àˆ,À¯,Ã°,Ã¢,Ãß,ÃØ,Ã∞,Ã±,Ã≤,Ã≥,Ã∑,Ã∏,Ãæ,Ã¿,ÃÃ,ÃŒ,Ã–,Ã⁄,Ã‹,Ã‡,Ã‚,ÃÂ,ÃÎ,Ãı,Ã˘,Ã˙,Ã¸,Ã˝,Ã˛,Õ≠,Õ≥,Õ∑,Õº,Õø,Õ≈,Õ«,Õ…,Õ—,Õ“,Õ‘,Õ’,Õ÷,Õ›,Õ‡,Õ‰,ÕÂ,ÕÁ,ÕÚ,Õ¯,Œ§,Œ•,Œß,Œ™,Œ´,Œ¨,Œ≠,Œ∞,Œ±,Œ≥,ŒΩ,Œ¿,Œ¬,Œ≈,Œ∆,Œ»,Œ ,ŒÕ,ŒŒ,Œœ,Œ–,Œ—,Œÿ,ŒŸ,Œ⁄,Œ‹,Œﬁ,Œﬂ,Œ‚,ŒÎ,ŒÌ,ŒÒ,ŒÛ,Œ˝,Œ˛,œÆ,œ∞,œ≥,œ∑,œ∏,œ∫,œΩ,œø,œ¿,œ¡,œ√,œ«,œ ,œÀ,œÃ,œÕ,œŒ,œ–,œ‘,œ’,œ÷,œ◊,œÿ,œ⁄,œ€,œ‹,œﬂ,œ·,œ‚,œÁ,œÍ,œÏ,œÓ,œÙ,œ˙,œ˛,–•,–´,–≠,–Æ,–Ø,–≤,–≥,–¥,–∫,–ª,–ø,–∆,–À,–⁄,–‚,–Â,–È,–Í,–Î,–Ì,–˜,–¯,–˘,–¸,—°,—¢,—§,—ß,—´,—Ø,—∞,—±,—µ,—∂,—∑,—π,—ª,—º,—∆,—«,—»,—À,—Ã,—Œ,—œ,—’,—÷,—ﬁ,—·,—‚,—Â,—Ë,—È,—Ï,—Ó,—Ô,—Ò,—Ù,—˜,—¯,—˘,—˛,“°,“¢,“£,“§,“•,“©,“Ø,“≥,“µ,“∂,“Ω,“ø,“√,“≈,“«,“Õ,“œ,“’,“⁄,“‰,“Â,“Ë,“È,“Í,“Î,“Ï,“Ô,“Ò,“ı,“¯,“˚,”£,”§,”•,”¶,”ß,”®,”©,”™,”´,”¨,”±,”¥,”µ,”∂,”∏,”ª,”Ω,”ø,”≈,”«,” ,”À,”Ã,”Œ,”’,”ﬂ,”„,”Ê,”È,”Î,”Ï,”Ô,”ı,”˘,”¸,”˛,‘§,‘¶,‘ß,‘®,‘Ø,‘∞,‘±,‘≤,‘µ,‘∂,‘∏,‘º,‘æ,‘ø,‘¿,‘¡,‘√,‘ƒ,‘∆,‘«,‘»,‘…,‘À,‘Ã,‘Õ,‘Œ,‘œ,‘”,‘÷,‘ÿ,‘‹,‘›,‘ﬁ,‘ﬂ,‘‡,‘‰,‘Ê,‘Ó,‘,‘Ò,‘Ú,‘Û,‘Ù,‘˘,‘˙,‘˝,‘˛,’°,’¢,’©,’´,’Æ,’±,’µ,’∂,’∑,’∏,’ª,’Ω,’¿,’≈,’«,’ ,’À,’Õ,’‘,’›,’ﬁ,’‡,’‚,’Í,’Î,’Ï,’Ô,’Ú,’Û,’ı,’ˆ,’¯,÷°,÷£,÷§,÷Ø,÷∞,÷¥,÷Ω,÷ø,÷¿,÷ƒ,÷ ,÷”,÷’,÷÷,÷◊,÷⁄,÷ﬂ,÷·,÷Â,÷Á,÷Ë,÷Ì,÷Ó,÷Ô,÷Ú,÷ı,÷ˆ,÷¸,÷˝,÷˛,◊§,◊®,◊©,◊™,◊¨,◊Æ,◊Ø,◊∞,◊±,◊≥,◊¥,◊∂,◊∏,◊π,◊∫,◊ª,◊«,◊»,◊ ,◊’,◊Ÿ,◊€,◊‹,◊›,◊ﬁ,◊Á,◊È,◊Í,÷¬,÷”,√¥,Œ™,÷ª,–◊,◊º,∆Ù,∞Â,¿Ô,ˆ®,”‡,¡¥,–π"
t="∞},Ã@,µK,ê€,¬O,“\,äW,âŒ,¡T,î[,î°,ÓC,ﬁk,ΩO,éÕ,Ωâ,Ê^,÷r,ÑÉ,Ôñ,åö,àÛ,ıU,›Ö,ÿê,‰^,™N,Ç‰,ëv,øá,πP,ÆÖ,î¿,È],ﬂÖ,æé,ŸH,◊É,ﬁq,ﬁp,¸Ç,∞T,ûl,ûI,Ÿe,îP,Ôû,ì‹,¿è,„K,Òg, N,—a,Ö¢,–Q,öà,ëM,ëK,†N,…n,≈ì,Ç},úÊ,é˙,Ç»,É‘,úy,å”,‘å,îv,ìΩ,œs,í,◊ã,¿p,ÁP,Æb,ÍU,Óù,àˆ,áL,ÈL,Éî,ƒc,èS,ï≥,‚n,‹á,èÿ,âm,Íê,“r,ìŒ,∑Q,ëÕ,’\,ÚG,∞V,ﬂt,ÒY,êu,˝X,üÎ,õ_,œx,åô,Æ†,‹P,ªI,æI,·h,ôª,èN,‰z,Îr,µA,É¶,”|,Ãé,Ç˜,Øè,ÍJ,Ñì,ÂN,ºÉ,æb,ﬁo,‘~,Ÿn,¬î, [,áË,èƒ,Ö≤,úê,∏Z,Âe,ﬂ_,éß,ŸJ,ì˙,ÜŒ,‡ê,ì€,ƒë,ëÑ,’Q,èó,Æî,ìı,¸h, é,ôn,ìv,çu,∂\,åß,±I,üÙ,‡á,î≥,úÏ,ﬂf,æÜ,¸c,â|,Îä,ù’,·û,’{,µ˛,’ô,ØB,·î,Ìî,ÂV,”Ü,ñ|,Ñ”,óù,Éˆ,ÙY,†Ÿ,™ö,◊x,ŸÄ,ÂÉ,Âë,î‡,æÑ,É∂,Í†,å¶,áç,ÓD,‚g,äZ,˘Z,Ó~,”û,ê∫,I,É∫,†ñ,D,ŸE,∞l,¡P,Èy,¨m,µ\,‚C,ü©,π†,ÿú,Ôà,‘L,ºè,Ôw,èU,ŸM,ºä,âû,ä^,ëç,ºS,ÿS,ó˜,‰h,ÔL,ØÇ,ÒT,øp,÷S,¯P,ƒw,›ó,ì·,›o,Ÿx,—},ÿì,”á,ãD,ø`,‘ì,‚},…w,é÷,⁄s,∂í,⁄M,å˘,ÑÇ,‰ì,æV,çè,≈V,ÊÄ,îR,¯ù,Èw,„t,ÇÄ,Ωo,˝è,åm,Ïñ,ÿï,‚h,úœ,òã,Ÿè,âÚ,–M,Óô,Ñé,ÍP,”^,^,ëT,ÿû,èV,“é,Œ˘,öw,˝î,È|,‹â,‘é,ôô,ŸF,Ñ£,›Å,ùL,ÂÅ,á¯,ﬂ^,Òî,Ìn,ùh,Èu,˙Q,ŸR,ôM,ﬁZ,¯ô,ºt,··,âÿ,◊o,ú˚,ëÙ,áW,»A,Æã,Ñù,‘í,ë—,âƒ,ög,≠h,ﬂÄ,æè,ìQ,Üæ,Øà,ü®,úo,¸S,÷e,ì],›x,öß,ŸV,∑x,ï˛,†Z,è°,÷M,’d,¿L,»ù,úÜ,‚∑,´@,ÿõ,µú,ìÙ,ôC,∑e,á,◊I,Îu,øÉ,æÉ,òO,›ã,ºâ,îD,é◊,ÀE,Ñ©,ù˙,”ã,”õ,ÎH,¿^,ºo,äA,«v,Óa,ŸZ,‚õ,Ér,Ò{,öû,±O,à‘,π{,Èg,∆D,æ},¿O,ôz,âA,˚|,í˛,ìÏ,∫Ü,ÉÄ,úp,À],ôë,Ëb,€`,Ÿv,“ä,ÊI,≈û,Ñ¶,T,ùu,ûR,ùæ,ù{, Y,ò™,™Ñ,÷v,·u,ƒz,ù≤,Úú,ã…,îá,„q,≥C,Ée,ƒ_,Ôú,¿U,Ωg,ﬁI,›^,∑M,ÎA,πù,«o,Û@,Ωõ,Ói,Ïo,ÁR,èΩ,Ød,∏Ç,úQ,ºm,é˝,≈f,Òx,≈e,ì˛,‰è,ë÷,Ñ°,˘N,ΩÅ,Ç‹,ùç,ΩY,’],å√,æo,Â\,ÉH,÷î,ﬂM,ïx,†a,±M,Ñ≈,«G,”X,õQ,‘E,Ω^,‚x,‹ä,ÚE,È_,ÑP,Ów,ö§,’n,â®,ë©,ì∏,éÏ,—ù,’F,âK,É~,åí,µV,ïÁ,õr,Ãù,éh,∏Q,Å,ù¢,îU,Èü,œû,≈D,»R,ÅÌ,Ÿá,À{,ô⁄,îr,ª@,Í@,Ãm,ûë,◊é,îà,”[,ë–,¿|,†Ä,ûE,ì∆,Ñ⁄,ù≥,ò∑,ËD,âæ,Óê,úI,ªh,Îx,—Y,ıé,∂Y,˚ê,Öñ,ÑÓ,µ[,ï—,ûr,Î`,Çz,¬ì,…è,ﬂB,Á†,ëz,ùi,∫ü,îø,ƒò,Êú,ëŸ,üí,æö,ºZ,õˆ,É…,›v,’è,Øü,ﬂ|,ÁÇ,´C,≈R,‡è,˜[,ÑC,ŸU,˝g,‚è,úR,Ï`,éX,ÓI,s,Ñ¢,˝à,√@,áµ,ª\,â≈,în,Î],ò«,ä‰,ìß,∫t,ÃJ,±R,ÔB,è],†t,ìÔ,˚u,Ãî,Ùî,ŸT,µì,‰õ,Íë,ÛH,ÖŒ,‰X,ÇH,å“,ø|,ë],ûV,æG,én,îÅ,å\,û¥,Åy,í‡,›Ü,Çê,Åˆ,úS,æ],’ì,Ã},¡_,ﬂâ,Ëå,ªj,ÚÖ,Òò,Ωj,ãå,¨î,¥a,Œõ,ÒR,¡R,Ü·,ŸI,˚ú,Ÿu,ﬂ~,√},≤m,z,–U,ùM,÷ô,ÿà,Â^,„T,ŸQ,˜·,¸q,õ],ÊV,ÈT,êû,ÇÉ,Âi,âÙ,÷i,èõ,“í,æd,æí,èR,úÁ,ëë,È},¯Q,„ë,÷á,÷\,ÆÄ,‚c,º{,Îy,ìœ,ƒX,ê¿,Ù[,H,ƒÅ,îf,ì”,·Ñ,¯B,¬ô,˝m,Ëá,Êá,ôé,™ü,Â∏,îQ,ùÙ,‚o,º~,ƒì,ù‚,ﬁr,Øë,÷Z,öW,˙t,ö™,áI,ùa,±P,˝ã,á¯,ê€,Ÿr,áä,˘i,Ú_,Ôh,Ól,ÿö,ÃO,ë{,‘u,ùä,ÓH,ì‰,‰Å,ò„,◊V,ƒö,˝R,ÚT,ÿM,Üô,ö‚,óâ,”ô,†ø,íL,‚T,„U,ﬂw,∫û,÷t,ÂX,„Q,ùì,ú\,◊l,âq,òå,Ü‹,†ù,ÀN,èä,ìå,Ê@,òÚ,ÜÃ,ÉS,¬N,∏[,∏`,öJ,”H,›p,ö‰,ÉA,Ìï,’à,ëc,≠Ç,∏F,⁄Ö,Ö^,‹|,Úå,˝x,ÔE,ô‡,ÑÒ,Ös,˘o,◊å,à,î_,¿@,ü·,Ìg,’J,ºx,òs,Ωq,‹õ,‰J,Èc,ùô,û¢,À_,ˆw,Ÿê,Ç„,Ü ,Ú},íﬂ,ù≠,ö¢,ºÜ,∫Y,ïÒ,ÈW,ÍÑ,Ÿ†,øò,Ç˚,Ÿp,ü˝,ΩB,Ÿd,îz,ëÿ,‘O,ºù,åè,ã,ƒI,ùB,¬ï,¿K,ÑŸ,¬},éü,™{,ùÒ,‘ä,å∆,ïr,Œg,åç,◊R,ÒÇ,Ñ›,·å,Ôó,“ï,‘á,â€,´F,ò–,›î,ï¯,⁄H,åŸ,–g,ò‰,ÿQ,îµ,éõ,Îp,’l,∂ê,Ìò,’f,¥T,†q,Ωz,Ôï,¬ñ,ëZ,Ìû,‘A,’b,î\,ÃK,‘V,√C,Îm,Ωó,öq,åO,ìp,πS,øs,¨ç,Êi,´H,ìÈ,îE,îÇ,ÿù,∞c,û©,âØ,◊T,’Ñ,öU,ú´,†C,ù˝,øl,Úv,÷`,‰R,Ó},Ûw,åœ,ól,ŸN,ËF,èd,¬†,üN,„~,Ωy,Ó^,àD,âT,àF,Ój,Õë,√ì,¯r,ÒW,ÒÑ,ôE,∏D,“m,èù,û≥,ÓB,»f,æW,Ìf,ﬂ`,á˙,†ë,ûH,æS,»î,Ç•,É^,æï,÷^,–l,úÿ,¬Ñ,ºy,∑Ä,Üñ,ÆY,ìÎ,ŒÅ,úu,∏C,ÜË,Êu,ûı,’_,üo, è,Ö«,â],ÏF,Ñ’,’`,Âa,†ﬁ,“u,¡ï,„ä,ëÚ,ºö,Œr,›†,ç{,Çb,™M,èB,Âv,ır,¿w,˚y,Ÿt,„ï,Èe,Ô@,ÎU,¨F,´I,øh,W,¡w,ëó,æÄ,é˚,ËÇ,‡l,‘î,Ìë,Ìó, í,‰N,ï‘,á[,œê,Öf,í∂,îy,√{,÷C,åë,ûa,÷x,‰\,·Ö,≈d,õ∞,Án,¿C,Ãì,áu,Ìö,‘S,æw,¿m,‹é,ë“,ﬂx,∞_,Ωk,åW,ÑÏ,‘É,å§,ÒZ,”ñ,”ç,ﬂd,â∫,¯f,¯Ü,Ü°,ÅÜ,”†,Èé,üü,˚},á¿,ÓÜ,Èê,ÿW,Öí,≥é,è©,÷V,Úû,¯Ñ,óÓ,ìP,ØÉ,Íñ,∞W,B,ò”,¨é,ìu,àÚ,ﬂb,∏G,÷{,Àé,†î,Ìì,òI,»~,·t,„û,ÓU,ﬂz,Éx,è§,œÅ,Àá,É|,ëõ,¡x,‘Ñ,◊h,’x,◊g,Æê,¿[, a,Íé,„y,Ôã,ô—,ãÎ,˙ó,ë™,¿t,¨ì,Œû,†I,ü…,œâ,∑f,Ü—,ìÌ,ÇÚ,∞b,€x,‘Å,ú•,Éû,ën,‡],‚ô,™q,ﬂ[,’T,›õ,Ù~,ùO,ä ,≈c,éZ,’Z,ªn,∂R,™z,◊u,ÓA,ÒS,¯x,úY,ﬁ@,à@,ÜT,àA,æâ,ﬂh,Óä,ºs,‹S,ËÄ,é[,ªõ,êÇ,ÈÜ,ÎÖ,‡y,ÑÚ,ÎE,ﬂ\,ÃN,·j,ïû,Ìç,Îs,ûƒ,›d,îÄ,ï∫,Ÿù,⁄E,Ûv,Ëè,óó,∏^,ÿü,ìÒ,Ñt,ù…,Ÿ\,Ÿõ,ºô,Ñû,‹à,Âé,Èl,‘p,˝S,Ç˘,ö÷,±K,îÿ,›ö,ç‰,ó£,ë,æ`,èà,ùq,é§,Ÿ~,√õ,⁄w,œU,ﬁH,ÊN,ﬂ@,ÿë,·ò,Ç…,‘\,ÊÇ,Íá,íÍ,±†,™b,é¨,‡ç,◊C,øó,¬ö,àÃ,ºà,ì¥,îS,é√,Ÿ|,ÊR,ΩK,∑N,ƒ[,–\,÷a,›S,∞ô,ïÉ,ÛE,ÿi,÷T,’D,†T,≤ö,á⁄,ŸA,ËT,∫B,Òv,å£,¥u,ﬁD,Ÿç,ò∂,«f,—b,äy,â—,†Ó,ÂF,Ÿò,âã,æY,’Å,ù·,∆ù,ŸY,ùn,€ô,æC,øÇ,øv,‡u,‘{,ΩM,Ëç,ø@,Áä,¸N,ûÈ,Îb,É¥,ú ,Ü¢,Èõ,—e,ÏZ,N,ÂÄ,õ™"
c=split(s,",")
d=split(t,",")
for i=0 to 1274
content=replace(content,c(i),d(i))
next
GbToBig=content
end function%>
<%
'µ√µΩ“ª∏ˆ∫∫◊÷◊÷“‚ŒÂ––
function getzywh(txt)
dim sql1,wh
sql1="select * from hzwh"
set rs1=conn.execute(sql1)
  if not (rs1.bof and rs1.eof) then
    do while not rs1.eof
      if instr(1,rs1("hz"),txt,1) >0 then
        wh=rs1("wh")
	  exit do
      else 
      rs1.movenext
      end if
    loop
end if
getzywh=wh
rs1.close
end function
%>
<%function getsancai(sc)
sc=(sc mod 10)
select case sc
  case 1
  sctxt="ƒæ"
  case 2
  sctxt="ƒæ"
  case 3
  sctxt="ª"
  case 4
  sctxt="ª"
  case 5
  sctxt="Õ¡"
  case 6
  sctxt="Õ¡"
  case 7
  sctxt="Ω"
  case 8
  sctxt="Ω"
  case 9
  sctxt="ÀÆ"
  case 10
  sctxt="ÀÆ"
  case 0
  sctxt="ÀÆ"
end select
getsancai=sctxt
end function%><%function DiZhi(i)
select case i
case 0
dz="◊”"
case 1
dz="≥Û"
case 2
dz="≥Û"
case 3
dz="“˙"
case 4
dz="“˙"
case 5
dz="√Æ"
case 6
dz="√Æ"
case 7
dz="≥Ω"
case 8
dz="≥Ω"
case 9
dz="À»"
case 10
dz="À»"
case 11
dz="ŒÁ"
case 12
dz="ŒÁ"
case 13
dz="Œ¥"
case 14
dz="Œ¥"
case 15
dz="…Í"
case 16
dz="…Í"
case 17
dz="”œ"
case 18
dz="”œ"
case 19
dz="–Á"
case 20
dz="–Á"
case 21
dz="∫•"
case 22
dz="∫•"
case 23
dz="◊”"
end select
DiZhi=dz
end function%>
<%function getpf(sc)
select case sc
  case "¥Ûº™"
  szpf=12
  case "º™"
  szpf=8
  case "∞Îº™"
  szpf=5
  case "¥Û–◊"
  szpf=0
  case "–◊"
  szpf=1
  case "∞Î–◊"
  szpf=2
  case "∆Ω"
  szpf=4
end select
getpf=szpf
end function%><%
'ƒ…“Ù
function layin(tgdz)
sql1="select * from jiazi"
set rs1=conn.execute(sql1)
  if not (rs1.bof and rs1.eof) then
    do while not rs1.eof
      if instr(1,rs1("jiazi"),tgdz,1) >0 then
        ly=rs1("layin")
	  exit do
      else 
      rs1.movenext
      end if
    loop
end if
layin=ly
rs1.close
end function
%>
<%function tgdzwh(tgdz)
select case tgdz
  case "◊”"
  wh="ÀÆ"
  case "∫•"
  wh="ÀÆ"
  case "“˙"
  wh="ƒæ"
  case "√Æ"
  wh="ƒæ"
  case "À»"
  wh="ª"
  case "ŒÁ"
  wh="ª"
  case "…Í"
  wh="Ω"
  case "”œ"
  wh="Ω"
  case "≥Ω"
  wh="Õ¡"
  case "–Á"
  wh="Õ¡"
  case "≥Û"
  wh="Õ¡"
  case "Œ¥"
  wh="Õ¡"
    case "º◊"
  wh="ƒæ"
    case "““"
  wh="ƒæ"
    case "±˚"
  wh="ª"
    case "∂°"
  wh="ª"
    case "ŒÏ"
  wh="Õ¡"
    case "º∫"
  wh="Õ¡"
    case "∏˝"
  wh="Ω"
      case "–¡"
  wh="Ω"
      case "»…"
  wh="ÀÆ"
      case "πÔ"
  wh="ÀÆ"
end select
tgdzwh=wh
end function%>
<%function siji(yue)
select case yue
  case "1"
  sj="∂¨"
  case "2"
  sj="∂¨"
  case "3"
  sj="¥∫"
  case "4"
  sj="¥∫"
  case "5"
  sj="¥∫"
  case "6"
  sj="œƒ"
  case "7"
  sj="œƒ"
  case "8"
  sj="œƒ"
  case "9"
  sj="«Ô"
  case "10"
  sj="«Ô"
  case "11"
  sj="«Ô"
  case "12"
  sj="∂¨"
end select
siji=sj
end function%><% 
    Set d = CreateObject("Scripting.Dictionary") 
    d.add "a",-20319 
    d.add "ai",-20317 
    d.add "an",-20304 
    d.add "ang",-20295 
    d.add "ao",-20292 
    d.add "ba",-20283 
    d.add "bai",-20265 
    d.add "ban",-20257 
    d.add "bang",-20242 
    d.add "bao",-20230 
    d.add "bei",-20051 
    d.add "ben",-20036 
    d.add "beng",-20032 
    d.add "bi",-20026 
    d.add "bian",-20002 
    d.add "biao",-19990 
    d.add "bie",-19986 
    d.add "bin",-19982 
    d.add "bing",-19976 
    d.add "bo",-19805 
    d.add "bu",-19784 
    d.add "ca",-19775 
    d.add "cai",-19774 
    d.add "can",-19763 
    d.add "cang",-19756 
    d.add "cao",-19751 
    d.add "ce",-19746 
    d.add "ceng",-19741 
    d.add "cha",-19739 
    d.add "chai",-19728 
    d.add "chan",-19725 
    d.add "chang",-19715 
    d.add "chao",-19540 
    d.add "che",-19531 
    d.add "chen",-19525 
    d.add "cheng",-19515 
    d.add "chi",-19500 
    d.add "chong",-19484 
    d.add "chou",-19479 
    d.add "chu",-19467 
    d.add "chuai",-19289 
    d.add "chuan",-19288 
    d.add "chuang",-19281 
    d.add "chui",-19275 
    d.add "chun",-19270 
    d.add "chuo",-19263 
    d.add "ci",-19261 
    d.add "cong",-19249 
    d.add "cou",-19243 
    d.add "cu",-19242 
    d.add "cuan",-19238 
    d.add "cui",-19235 
    d.add "cun",-19227 
    d.add "cuo",-19224 
    d.add "da",-19218 
    d.add "dai",-19212 
    d.add "dan",-19038 
    d.add "dang",-19023 
    d.add "dao",-19018 
    d.add "de",-19006 
    d.add "deng",-19003 
    d.add "di",-18996 
    d.add "dian",-18977 
    d.add "diao",-18961 
    d.add "die",-18952 
    d.add "ding",-18783 
    d.add "diu",-18774 
    d.add "dong",-18773 
    d.add "dou",-18763 
    d.add "du",-18756 
    d.add "duan",-18741 
    d.add "dui",-18735 
    d.add "dun",-18731 
    d.add "duo",-18722 
    d.add "e",-18710 
    d.add "en",-18697 
    d.add "er",-18696 
    d.add "fa",-18526 
    d.add "fan",-18518 
    d.add "fang",-18501 
    d.add "fei",-18490 
    d.add "fen",-18478 
    d.add "feng",-18463 
    d.add "fo",-18448 
    d.add "fou",-18447 
    d.add "fu",-18446 
    d.add "ga",-18239 
    d.add "gai",-18237 
    d.add "gan",-18231 
    d.add "gang",-18220 
    d.add "gao",-18211 
    d.add "ge",-18201 
    d.add "gei",-18184 
    d.add "gen",-18183 
    d.add "geng",-18181 
    d.add "gong",-18012 
    d.add "gou",-17997 
    d.add "gu",-17988 
    d.add "gua",-17970 
    d.add "guai",-17964 
    d.add "guan",-17961 
    d.add "guang",-17950 
    d.add "gui",-17947 
    d.add "gun",-17931 
    d.add "guo",-17928 
    d.add "ha",-17922 
    d.add "hai",-17759 
    d.add "han",-17752 
    d.add "hang",-17733 
    d.add "hao",-17730 
    d.add "he",-17721 
    d.add "hei",-17703 
    d.add "hen",-17701 
    d.add "heng",-17697 
    d.add "hong",-17692 
    d.add "hou",-17683 
    d.add "hu",-17676 
    d.add "hua",-17496 
    d.add "huai",-17487 
    d.add "huan",-17482 
    d.add "huang",-17468 
    d.add "hui",-17454 
    d.add "hun",-17433 
    d.add "huo",-17427 
    d.add "ji",-17417 
    d.add "jia",-17202 
    d.add "jian",-17185 
    d.add "jiang",-16983 
    d.add "jiao",-16970 
    d.add "jie",-16942 
    d.add "jin",-16915 
    d.add "jing",-16733 
    d.add "jiong",-16708 
    d.add "jiu",-16706 
    d.add "ju",-16689 
    d.add "juan",-16664 
    d.add "jue",-16657 
    d.add "jun",-16647 
    d.add "ka",-16474 
    d.add "kai",-16470 
    d.add "kan",-16465 
    d.add "kang",-16459 
    d.add "kao",-16452 
    d.add "ke",-16448 
    d.add "ken",-16433 
    d.add "keng",-16429 
    d.add "kong",-16427 
    d.add "kou",-16423 
    d.add "ku",-16419 
    d.add "kua",-16412 
    d.add "kuai",-16407 
    d.add "kuan",-16403 
    d.add "kuang",-16401 
    d.add "kui",-16393 
    d.add "kun",-16220 
    d.add "kuo",-16216 
    d.add "la",-16212 
    d.add "lai",-16205 
    d.add "lan",-16202 
    d.add "lang",-16187 
    d.add "lao",-16180 
    d.add "le",-16171 
    d.add "lei",-16169 
    d.add "leng",-16158 
    d.add "li",-16155 
    d.add "lia",-15959 
    d.add "lian",-15958 
    d.add "liang",-15944 
    d.add "liao",-15933 
    d.add "lie",-15920 
    d.add "lin",-15915 
    d.add "ling",-15903 
    d.add "liu",-15889 
    d.add "long",-15878 
    d.add "lou",-15707 
    d.add "lu",-15701 
    d.add "lv",-15681 
    d.add "luan",-15667 
    d.add "lue",-15661 
    d.add "lun",-15659 
    d.add "luo",-15652 
    d.add "ma",-15640 
    d.add "mai",-15631 
    d.add "man",-15625 
    d.add "mang",-15454 
    d.add "mao",-15448 
    d.add "me",-15436 
    d.add "mei",-15435 
    d.add "men",-15419 
    d.add "meng",-15416 
    d.add "mi",-15408 
    d.add "mian",-15394 
    d.add "miao",-15385 
    d.add "mie",-15377 
    d.add "min",-15375 
    d.add "ming",-15369 
    d.add "miu",-15363 
    d.add "mo",-15362 
    d.add "mou",-15183 
    d.add "mu",-15180 
    d.add "na",-15165 
    d.add "nai",-15158 
    d.add "nan",-15153 
    d.add "nang",-15150 
    d.add "nao",-15149 
    d.add "ne",-15144 
    d.add "nei",-15143 
    d.add "nen",-15141 
    d.add "neng",-15140 
    d.add "ni",-15139 
    d.add "nian",-15128 
    d.add "niang",-15121 
    d.add "niao",-15119 
    d.add "nie",-15117 
    d.add "nin",-15110 
    d.add "ning",-15109 
    d.add "niu",-14941 
    d.add "nong",-14937 
    d.add "nu",-14933 
    d.add "nv",-14930 
    d.add "nuan",-14929 
    d.add "nue",-14928 
    d.add "nuo",-14926 
    d.add "o",-14922 
    d.add "ou",-14921 
    d.add "pa",-14914 
    d.add "pai",-14908 
    d.add "pan",-14902 
    d.add "pang",-14894 
    d.add "pao",-14889 
    d.add "pei",-14882 
    d.add "pen",-14873 
    d.add "peng",-14871 
    d.add "pi",-14857 
    d.add "pian",-14678 
    d.add "piao",-14674 
    d.add "pie",-14670 
    d.add "pin",-14668 
    d.add "ping",-14663 
    d.add "po",-14654 
    d.add "pu",-14645 
    d.add "qi",-14630
    d.add "qia",-14594 
    d.add "qian",-14429 
    d.add "qiang",-14407 
    d.add "qiao",-14399 
    d.add "qie",-14384 
    d.add "qin",-14379 
    d.add "qing",-14368 
    d.add "qiong",-14355 
    d.add "qiu",-14353 
    d.add "qu",-14345 
    d.add "quan",-14170 
    d.add "que",-14159 
    d.add "qun",-14151 
    d.add "ran",-14149 
    d.add "rang",-14145 
    d.add "rao",-14140 
    d.add "re",-14137 
    d.add "ren",-14135 
    d.add "reng",-14125 
    d.add "ri",-14123 
    d.add "rong",-14122 
    d.add "rou",-14112 
    d.add "ru",-14109 
    d.add "ruan",-14099 
    d.add "rui",-14097 
    d.add "run",-14094 
    d.add "ruo",-14092 
    d.add "sa",-14090 
    d.add "sai",-14087 
    d.add "san",-14083 
    d.add "sang",-13917 
    d.add "sao",-13914 
    d.add "se",-13910 
    d.add "sen",-13907 
    d.add "seng",-13906 
    d.add "sha",-13905 
    d.add "shai",-13896 
    d.add "shan",-13894 
    d.add "shang",-13878 
    d.add "shao",-13870 
    d.add "she",-13859 
    d.add "shen",-13847 
    d.add "sheng",-13831 
    d.add "shi",-13658 
    d.add "shou",-13611 
    d.add "shu",-13601 
    d.add "shua",-13406 
    d.add "shuai",-13404 
    d.add "shuan",-13400 
    d.add "shuang",-13398 
    d.add "shui",-13395 
    d.add "shun",-13391 
    d.add "shuo",-13387 
    d.add "si",-13383 
    d.add "song",-13367 
    d.add "sou",-13359 
    d.add "su",-13356 
    d.add "suan",-13343 
    d.add "sui",-13340 
    d.add "sun",-13329 
    d.add "suo",-13326 
    d.add "ta",-13318 
    d.add "tai",-13147 
    d.add "tan",-13138 
    d.add "tang",-13120 
    d.add "tao",-13107 
    d.add "te",-13096 
    d.add "teng",-13095 
    d.add "ti",-13091 
    d.add "tian",-13076 
    d.add "tiao",-13068 
    d.add "tie",-13063 
    d.add "ting",-13060 
    d.add "tong",-12888 
    d.add "tou",-12875 
    d.add "tu",-12871 
    d.add "tuan",-12860 
    d.add "tui",-12858 
    d.add "tun",-12852 
    d.add "tuo",-12849 
    d.add "wa",-12838 
    d.add "wai",-12831 
    d.add "wan",-12829 
    d.add "wang",-12812 
    d.add "wei",-12802 
    d.add "wen",-12607 
    d.add "weng",-12597 
    d.add "wo",-12594 
    d.add "wu",-12585 
    d.add "xi",-12556 
    d.add "xia",-12359 
    d.add "xian",-12346 
    d.add "xiang",-12320 
    d.add "xiao",-12300 
    d.add "xie",-12120 
    d.add "xin",-12099 
    d.add "xing",-12089 
    d.add "xiong",-12074 
    d.add "xiu",-12067 
    d.add "xu",-12058 
    d.add "xuan",-12039 
    d.add "xue",-11867 
    d.add "xun",-11861 
    d.add "ya",-11847 
    d.add "yan",-11831 
    d.add "yang",-11798 
    d.add "yao",-11781 
    d.add "ye",-11604 
    d.add "yi",-11589 
    d.add "yin",-11536 
    d.add "ying",-11358 
    d.add "yo",-11340 
    d.add "yong",-11339 
    d.add "you",-11324 
    d.add "yu",-11303 
    d.add "yuan",-11097 
    d.add "yue",-11077 
    d.add "yun",-11067 
    d.add "za",-11055 
    d.add "zai",-11052 
    d.add "zan",-11045 
    d.add "zang",-11041 
    d.add "zao",-11038 
    d.add "ze",-11024 
    d.add "zei",-11020 
    d.add "zen",-11019 
    d.add "zeng",-11018 
    d.add "zha",-11014 
    d.add "zhai",-10838 
    d.add "zhan",-10832 
    d.add "zhang",-10815 
    d.add "zhao",-10800 
    d.add "zhe",-10790 
    d.add "zhen",-10780 
    d.add "zheng",-10764 
    d.add "zhi",-10587 
    d.add "zhong",-10544 
    d.add "zhou",-10533 
    d.add "zhu",-10519 
    d.add "zhua",-10331 
    d.add "zhuai",-10329 
    d.add "zhuan",-10328 
    d.add "zhuang",-10322 
    d.add "zhui",-10315 
    d.add "zhun",-10309 
    d.add "zhuo",-10307 
    d.add "zi",-10296 
    d.add "zong",-10281 
    d.add "zou",-10274 
    d.add "zu",-10270 
    d.add "zuan",-10262 
    d.add "zui",-10260 
    d.add "zun",-10256 
    d.add "zuo",-10254 
     
    function g(num) 
    if num>0 and num<160 then 
    g=chr(num) 
    else  
    if num<-20319 or num>-10247 then 
    g="" 
    else 
    aa=d.Items 
    b=d.keys 
    for k1=d.count-1 to 0 step -1 
    if aa(k1)<=num then exit for 
    next 
    g=b(k1) 
    end if 
    end if 
    end function 
    function c(str) 
    c="" 
    for k1=1 to len(str) 
    c=c&g(asc(mid(str,k1,1))) 
    next 
    end function  %>
