<%
'ȥ�������ַ��еĿո�---newtrim����
function newtrim(str)
dim result
dim j
j=len(str)
result=""
dim i
for i=1 to j
select case mid(str,i,1)
case chr(8)
     result=result+"" 'ȥ���˸�
case chr(9)
     result=result+"" 'ȥ��tab
case chr(10)
     result=result+"" 'ȥ������
case chr(13)
     result=result+"" 'ȥ���س�
case chr(255)
     result=result+"" 'ȥ������ո�
case else
     result=result+mid(str,i,1)
end select
next
newtrim=result
end function
%>
<%
'�õ�һ�����ֵıʻ���---getnum()����
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
'������˸񡢵ظ����õ�����---getsancai()����
function getsancai(tiange,renge,dige)
dim tian,ren,di,tiantxt,rentxt,ditxt,total
tian=(tiange mod 10)
ren=(renge mod 10)
di=(dige mod 10)
select case tian
  case 1
  tiantxt="ľ"
  case 2
  tiantxt="ľ"
  case 3
  tiantxt="��"
  case 4
  tiantxt="��"
  case 5
  tiantxt="��"
  case 6
  tiantxt="��"
  case 7
  tiantxt="��"
  case 8
  tiantxt="��"
  case 9
  tiantxt="ˮ"
  case 10
  tiantxt="ˮ"
  case 0
  tiantxt="ˮ"
end select
select case ren
  case 1
  rentxt="ľ"
  case 2
  rentxt="ľ"
  case 3
  rentxt="��"
  case 4
  rentxt="��"
  case 5
  rentxt="��"
  case 6
  rentxt="��"
  case 7
  rentxt="��"
  case 8
  rentxt="��"
  case 9
  rentxt="ˮ"
  case 10
  rentxt="ˮ"
  case 0
  rentxt="ˮ"
end select
select case di
  case 1
  ditxt="ľ"
  case 2
  ditxt="ľ"
  case 3
  ditxt="��"
  case 4
  ditxt="��"
  case 5
  ditxt="��"
  case 6
  ditxt="��"
  case 7
  ditxt="��"
  case 8
  ditxt="��"
  case 9
  ditxt="ˮ"
  case 10
  ditxt="ˮ"
  case 0
  ditxt="ˮ"
end select
total=tiantxt+rentxt+ditxt
getsancai=total
end function
%>
<%
Function hhcal(caltime) 
'����ת��ũ�����㷨 edit by huanghui
Dim WeekName(7), MonthAdd(11), NongliData(99), TianGan(9), DiZhi(11), ShuXiang(11), DayName(30), MonName(12) 
Dim curTime, curYear, curMonth, curDay, curWeekday 
Dim GongliStr, WeekdayStr, NongliStr, NongliDayStr 
Dim i, m, n, k, isEnd, bit, TheDate 
'��ȡ��ǰϵͳʱ�� 
curTime = caltime 
'������ 
WeekName(0) = " * " 
WeekName(1) = "������" 
WeekName(2) = "����һ" 
WeekName(3) = "���ڶ�" 
WeekName(4) = "������" 
WeekName(5) = "������" 
WeekName(6) = "������" 
WeekName(7) = "������" 
'������� 
TianGan(0) = "��" 
TianGan(1) = "��" 
TianGan(2) = "��" 
TianGan(3) = "��" 
TianGan(4) = "��" 
TianGan(5) = "��" 
TianGan(6) = "��" 
TianGan(7) = "��" 
TianGan(8) = "��" 
TianGan(9) = "��" 
'��֧���� 
DiZhi(0) = "��" 
DiZhi(1) = "��" 
DiZhi(2) = "��" 
DiZhi(3) = "î" 
DiZhi(4) = "��" 
DiZhi(5) = "��" 
DiZhi(6) = "��" 
DiZhi(7) = "δ" 
DiZhi(8) = "��" 
DiZhi(9) = "��" 
DiZhi(10) = "��" 
DiZhi(11) = "��" 
'�������� 
ShuXiang(0) = "��" 
ShuXiang(1) = "ţ" 
ShuXiang(2) = "��" 
ShuXiang(3) = "��" 
ShuXiang(4) = "��" 
ShuXiang(5) = "��" 
ShuXiang(6) = "��" 
ShuXiang(7) = "��" 
ShuXiang(8) = "��" 
ShuXiang(9) = "��" 
ShuXiang(10) = "��" 
ShuXiang(11) = "��" 
'ũ�������� 
DayName(0) = "*" 
DayName(1) = "��һ" 
DayName(2) = "����" 
DayName(3) = "����" 
DayName(4) = "����" 
DayName(5) = "����" 
DayName(6) = "����" 
DayName(7) = "����" 
DayName(8) = "����" 
DayName(9) = "����" 
DayName(10) = "��ʮ" 
DayName(11) = "ʮһ" 
DayName(12) = "ʮ��" 
DayName(13) = "ʮ��" 
DayName(14) = "ʮ��" 
DayName(15) = "ʮ��" 
DayName(16) = "ʮ��" 
DayName(17) = "ʮ��" 
DayName(18) = "ʮ��" 
DayName(19) = "ʮ��" 
DayName(20) = "��ʮ" 
DayName(21) = "إһ" 
DayName(22) = "إ��" 
DayName(23) = "إ��" 
DayName(24) = "إ��" 
DayName(25) = "إ��" 
DayName(26) = "إ��" 
DayName(27) = "إ��" 
DayName(28) = "إ��" 
DayName(29) = "إ��" 
DayName(30) = "��ʮ" 
'ũ���·��� 
MonName(0) = "*" 
MonName(1) = "��" 
MonName(2) = "��" 
MonName(3) = "��" 
MonName(4) = "��" 
MonName(5) = "��" 
MonName(6) = "��" 
MonName(7) = "��" 
MonName(8) = "��" 
MonName(9) = "��" 
MonName(10) = "ʮ" 
MonName(11) = "ʮһ" 
MonName(12) = "��" 
'����ÿ��ǰ������� 
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
'ũ������ 
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
'���ɵ�ǰ�����ꡢ�¡��� ==> GongliStr 
curYear = Year(curTime) 
curMonth = Month(curTime) 
curDay = Day(curTime) 
GongliStr = curYear & "��" 
If (curMonth < 10) Then 
    GongliStr = GongliStr & "0" & curMonth & "��" 
Else 
    GongliStr = GongliStr & curMonth & "��" 
End If 
If (curDay < 10) Then 
    GongliStr = GongliStr & "0" & curDay & "��" 
Else 
    GongliStr = GongliStr & curDay & "��" 
End If 
'���ɵ�ǰ�������� ==> WeekdayStr 
curWeekday = Weekday(curTime) 
WeekdayStr = WeekName(curWeekday) 
'���㵽��ʼʱ��1921��2��8�յ�������1921-2-8(���³�һ) 
TheDate = (curYear - 1921) * 365 + Int((curYear - 1921) / 4) + curDay + MonthAdd(curMonth - 1) - 38 
If ((curYear Mod 4) = 0 And curMonth > 2) Then 
    TheDate = TheDate + 1 
End If 
'����ũ����ɡ���֧���¡��� 
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
    '��ȡNongliData(m)�ĵ�n��������λ��ֵ 
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
'����ũ����ɡ���֧������ ==> NongliStr
dim nongli,nlnian,nlyue,nlri,sx
nlnian = TianGan(((curYear - 4) Mod 60) Mod 10) & DiZhi(((curYear - 4) Mod 60) Mod 12) 
sx =ShuXiang(((curYear - 4) Mod 60) Mod 12)
'����ũ���¡��� ==> NongliDayStr 
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
if not isdate(mDate) then response.write("������")
dim a
a=(Day(mDate) - (19 + Int(Mid("102123444423", Month(mDate), 1))))
if a>=0 then
a=0
else
a=-1
End if
'����
    Constellation = Mid("ħ��ˮƿ˫������ţ˫�Ӿ�зʨ�Ӵ�Ů�����Ы����ħ��", (Month(mDate) + a) * 2 + 1, 2) & "��"
End Function%>
<%
'�滻���з��Ĺ��˺���
Function rep(content)
rep=replace(content,vbCrLf,"<br>")
end Function
%><%function GbToBig(content)
dim s,t,c,d,i
s="��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,¢,£,¤,¥,¦,§,¨,«,¬,­,®,¯,°,±,²,³,¸,»,¼,½,¿,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,á,è,ê,í,ó,ô,ù,û,þ,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,ı,Ķ,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,š,Ţ,ť,Ŧ,ŧ,Ũ,ũ,ű,ŵ,ŷ,Ÿ,Ź,Ż,Ž,��,��,��,��,��,��,��,ƭ,Ʈ,Ƶ,ƶ,ƻ,ƾ,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,ǣ,Ǥ,ǥ,Ǧ,Ǩ,ǩ,ǫ,Ǯ,ǯ,Ǳ,ǳ,Ǵ,ǵ,ǹ,Ǻ,ǽ,Ǿ,ǿ,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,ȣ,ȧ,Ȩ,Ȱ,ȴ,ȵ,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,ɡ,ɥ,ɧ,ɨ,ɬ,ɱ,ɴ,ɸ,ɹ,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,ʤ,ʥ,ʦ,ʨ,ʪ,ʫ,ʬ,ʱ,ʴ,ʵ,ʶ,ʻ,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,˧,˫,˭,˰,˳,˵,˶,˸,˿,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,̡,̢,̧,̯,̰,̱,̲,̳,̷,̸,̾,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,ͭ,ͳ,ͷ,ͼ,Ϳ,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,Τ,Υ,Χ,Ϊ,Ϋ,ά,έ,ΰ,α,γ,ν,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,Ϯ,ϰ,ϳ,Ϸ,ϸ,Ϻ,Ͻ,Ͽ,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,Х,Ы,Э,Ю,Я,в,г,д,к,л,п,��,��,��,��,��,��,��,��,��,��,��,��,��,ѡ,Ѣ,Ѥ,ѧ,ѫ,ѯ,Ѱ,ѱ,ѵ,Ѷ,ѷ,ѹ,ѻ,Ѽ,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,ҡ,Ң,ң,Ҥ,ҥ,ҩ,ү,ҳ,ҵ,Ҷ,ҽ,ҿ,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,ӣ,Ӥ,ӥ,Ӧ,ӧ,Ө,ө,Ӫ,ӫ,Ӭ,ӱ,Ӵ,ӵ,Ӷ,Ӹ,ӻ,ӽ,ӿ,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,Ԥ,Ԧ,ԧ,Ԩ,ԯ,԰,Ա,Բ,Ե,Զ,Ը,Լ,Ծ,Կ,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,ա,բ,թ,ի,ծ,ձ,յ,ն,շ,ո,ջ,ս,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,֡,֣,֤,֯,ְ,ִ,ֽ,ֿ,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,��,פ,ר,ש,ת,׬,׮,ׯ,װ,ױ,׳,״,׶,׸,׹,׺,׻,��,��,��,��,��,��,��,��,��,��,��,��,��,��,ô,Ϊ,ֻ,��,׼,��,��,��,��,��,��,й"
t="�},�@,�K,��,�O,�\,�W,��,�T,�[,��,�C,�k,�O,��,��,�^,�r,��,�,��,��,�U,݅,ؐ,�^,�N,��,�v,��,�P,��,��,�],߅,��,�H,׃,�q,�p,��,�T,�l,�I,�e,�P,�,��,��,�K,�g,�N,�a,��,�Q,��,�M,�K,�N,�n,œ,�},��,��,��,��,�y,��,Ԍ,�v,��,�s,�,׋,�p,�P,�b,�U,�,��,�L,�L,��,�c,�S,��,�n,܇,��,�m,�,�r,��,�Q,��,�\,�G,�V,�t,�Y,�u,�X,��,�_,�x,��,��,�P,�I,�I,�h,��,�N,�z,�r,�A,��,�|,̎,��,��,�J,��,�N,��,�b,�o,�~,�n,,�[,��,��,��,��,�Z,�e,�_,��,�J,��,��,��,��,đ,��,�Q,��,��,��,�h,ʎ,�n,�v,�u,�\,��,�I,��,��,��,��,�f,��,�c,�|,�,��,�,�{,��,ՙ,�B,�,�,�V,ӆ,�|,��,��,��,�Y,��,��,�x,ـ,�,�,��,��,��,�,��,��,�D,�g,�Z,�Z,�~,Ӟ,��,�I,��,��,�D,�E,�l,�P,�y,�m,�\,�C,��,��,؜,�,�L,��,�w,�U,�M,��,��,�^,��,�S,�S,��,�h,�L,��,�T,�p,�S,�P,�w,ݗ,��,�o,�x,�},ؓ,Ӈ,�D,�`,ԓ,�},�w,��,�s,��,�M,��,��,�,�V,��,�V,�,�R,��,�w,�t,��,�o,��,�m,�,ؕ,�h,��,��,ُ,��,�M,�,��,�P,�^,�^,�T,؞,�V,Ҏ,��,�w,��,�|,܉,Ԏ,��,�F,��,݁,�L,�,��,�^,�,�n,�h,�u,�Q,�R,�M,�Z,��,�t,��,��,�o,��,��,�W,�A,��,��,Ԓ,��,��,�g,�h,߀,��,�Q,��,��,��,�o,�S,�e,�],�x,��,�V,�x,��,�Z,��,�M,�d,�L,ȝ,��,�,�@,؛,��,��,�C,�e,��,�I,�u,��,��,�O,݋,��,�D,��,�E,��,��,Ӌ,ӛ,�H,�^,�o,�A,�v,�a,�Z,�,�r,�{,��,�O,��,�{,�g,�D,�},�O,�z,�A,�|,��,��,��,��,�p,�],��,�b,�`,�v,Ҋ,�I,Ş,��,�T,�u,�R,��,�{,�Y,��,��,�v,�u,�z,��,�,��,��,�q,�C,�e,�_,�,�U,�g,�I,�^,�M,�A,��,�o,�@,��,�i,�o,�R,��,�d,��,�Q,�m,��,�f,�x,�e,��,�,��,��,�N,��,��,��,�Y,�],��,�o,�\,�H,֔,�M,�x,�a,�M,��,�G,�X,�Q,�E,�^,�x,܊,�E,�_,�P,�w,��,�n,��,��,��,��,ѝ,�F,�K,�~,��,�V,��,�r,̝,�h,�Q,��,��,�U,�,Ϟ,�D,�R,��,ه,�{,��,�r,�@,�@,�m,��,׎,��,�[,��,�|,��,�E,��,��,��,��,�D,��,�,�I,�h,�x,�Y,��,�Y,��,��,��,�[,��,�r,�`,�z,,ɏ,�B,�,�z,�i,��,��,Ę,�,��,��,��,�Z,��,��,�v,Տ,��,�|,�,�C,�R,��,�[,�C,�U,�g,�,�R,�`,�X,�I,�s,��,��,�@,��,�\,��,�n,�],��,��,��,�t,�J,�R,�B,�],�t,��,�u,̔,��,�T,��,�,�,�H,��,�X,�H,��,�|,�],�V,�G,�n,��,�\,��,�y,��,݆,��,��,�S,�],Փ,�},�_,߉,�,�j,�,�,�j,��,��,�a,Λ,�R,�R,��,�I,��,�u,�~,�},�m,�z,�U,�M,֙,؈,�^,�T,�Q,��,�q,�],�V,�T,��,��,�i,��,�i,��,Ғ,�d,��,�R,��,��,�},�Q,�,և,�\,��,�c,�{,�y,��,�X,��,�[,�H,ā,�f,��,�,�B,,�m,�,�,��,��,�,�Q,��,�o,�~,ē,��,�r,��,�Z,�W,�t,��,�I,�a,�P,��,��,��,�r,��,�i,�_,�h,�l,ؚ,�O,�{,�u,��,�H,��,�,��,�V,Ě,�R,�T,�M,��,��,��,ә,��,�L,�T,�U,�w,��,�t,�X,�Q,��,�\,�l,�q,��,��,��,�N,��,��,�@,��,��,�S,�N,�[,�`,�J,�H,�p,��,�A,�,Ո,�c,��,�F,څ,�^,�|,�,�x,�E,��,��,�s,�o,׌,��,�_,�@,��,�g,�J,�x,�s,�q,ܛ,�J,�c,��,��,�_,�w,ِ,��,��,�},��,��,��,��,�Y,��,�W,�,٠,��,��,�p,��,�B,�d,�z,��,�O,��,��,��,�I,�B,,�K,��,�},��,�{,��,Ԋ,��,�r,�g,��,�R,�,��,�,�,ҕ,ԇ,��,�F,��,ݔ,��,�H,��,�g,��,�Q,��,��,�p,�l,��,�,�f,�T,�q,�z,�,,�Z,�,�A,�b,�\,�K,�V,�C,�m,��,�q,�O,�p,�S,�s,��,�i,�H,��,�E,��,؝,�c,��,��,�T,Մ,�U,��,�C,��,�l,�v,�`,�R,�},�w,��,�l,�N,�F,�d, ,�N,�~,�y,�^,�D,�T,�F,�j,͑,Ó,�r,�W,�,�E,�D,�m,��,��,�B,�f,�W,�f,�`,��,��,�H,�S,Ȕ,��,�^,��,�^,�l,��,,�y,��,��,�Y,��,΁,�u,�C,��,�u,��,�_,�o,ʏ,��,�],�F,��,�`,�a,��,�u,��,�,��,��,�r,ݠ,�{,�b,�M,�B,�v,�r,�w,�y,�t,�,�e,�@,�U,�F,�I,�h,�W,�w,��,��,��,�,�l,Ԕ,�,�,ʒ,�N,��,�[,ϐ,�f,��,�y,�{,�C,��,�a,�x,�\,�,�d,��,�n,�C,̓,�u,�,�S,�w,�m,܎,��,�x,�_,�k,�W,��,ԃ,��,�Z,Ӗ,Ӎ,�d,��,�f,��,��,��,Ӡ,�,��,�},��,�,�,�W,��,��,��,�V,�,��,��,�P,��,�,�W,�B,��,��,�u,��,�b,�G,�{,ˎ,��,�,�I,�~,�t,�,�U,�z,�x,��,ρ,ˇ,�|,��,�x,Ԅ,�h,�x,�g,��,�[,�a,�,�y,�,��,��,��,��,�t,��,Ξ,�I,��,ω,�f,��,��,��,�b,�x,ԁ,��,��,�n,�],�,�q,�[,�T,ݛ,�~,�O,��,�c,�Z,�Z,�n,�R,�z,�u,�A,�S,�x,�Y,�@,�@,�T,�A,��,�h,�,�s,�S,�,�[,��,��,�,�,�y,��,�E,�\,�N,�j,��,�,�s,��,�d,��,��,ٝ,�E,�v,�,��,�^,؟,��,�t,��,�\,ٛ,��,��,܈,�,�l,�p,�S,��,��,�K,��,ݚ,��,��,��,�`,��,�q,��,�~,Û,�w,�U,�H,�N,�@,ؑ,�,��,�\,�,�,��,��,�b,��,��,�C,��,,��,��,��,�S,��,�|,�R,�K,�N,�[,�\,�a,�S,��,��,�E,�i,�T,�D,�T,��,��,�A,�T,�B,�v,��,�u,�D,ٍ,��,�f,�b,�y,��,��,�F,٘,��,�Y,Ձ,��,Ɲ,�Y,�n,ۙ,�C,��,�v,�u,�{,�M,�,�@,�,�N,��,�b,��,��,��,�,�e,�Z,�N,�,��"
c=split(s,",")
d=split(t,",")
for i=0 to 1274
content=replace(content,c(i),d(i))
next
GbToBig=content
end function%>
<%
'�õ�һ��������������
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
  sctxt="ľ"
  case 2
  sctxt="ľ"
  case 3
  sctxt="��"
  case 4
  sctxt="��"
  case 5
  sctxt="��"
  case 6
  sctxt="��"
  case 7
  sctxt="��"
  case 8
  sctxt="��"
  case 9
  sctxt="ˮ"
  case 10
  sctxt="ˮ"
  case 0
  sctxt="ˮ"
end select
getsancai=sctxt
end function%><%function DiZhi(i)
select case i
case 0
dz="��"
case 1
dz="��"
case 2
dz="��"
case 3
dz="��"
case 4
dz="��"
case 5
dz="î"
case 6
dz="î"
case 7
dz="��"
case 8
dz="��"
case 9
dz="��"
case 10
dz="��"
case 11
dz="��"
case 12
dz="��"
case 13
dz="δ"
case 14
dz="δ"
case 15
dz="��"
case 16
dz="��"
case 17
dz="��"
case 18
dz="��"
case 19
dz="��"
case 20
dz="��"
case 21
dz="��"
case 22
dz="��"
case 23
dz="��"
end select
DiZhi=dz
end function%>
<%function getpf(sc)
select case sc
  case "��"
  szpf=12
  case "��"
  szpf=8
  case "�뼪"
  szpf=5
  case "����"
  szpf=0
  case "��"
  szpf=1
  case "����"
  szpf=2
  case "ƽ"
  szpf=4
end select
getpf=szpf
end function%><%
'����
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
  case "��"
  wh="ˮ"
  case "��"
  wh="ˮ"
  case "��"
  wh="ľ"
  case "î"
  wh="ľ"
  case "��"
  wh="��"
  case "��"
  wh="��"
  case "��"
  wh="��"
  case "��"
  wh="��"
  case "��"
  wh="��"
  case "��"
  wh="��"
  case "��"
  wh="��"
  case "δ"
  wh="��"
    case "��"
  wh="ľ"
    case "��"
  wh="ľ"
    case "��"
  wh="��"
    case "��"
  wh="��"
    case "��"
  wh="��"
    case "��"
  wh="��"
    case "��"
  wh="��"
      case "��"
  wh="��"
      case "��"
  wh="ˮ"
      case "��"
  wh="ˮ"
end select
tgdzwh=wh
end function%>
<%function siji(yue)
select case yue
  case "1"
  sj="��"
  case "2"
  sj="��"
  case "3"
  sj="��"
  case "4"
  sj="��"
  case "5"
  sj="��"
  case "6"
  sj="��"
  case "7"
  sj="��"
  case "8"
  sj="��"
  case "9"
  sj="��"
  case "10"
  sj="��"
  case "11"
  sj="��"
  case "12"
  sj="��"
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
