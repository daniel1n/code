<%
'set fo=server.CreateObject("calendarx.ocx")
Dim WeekName(7), MonthAdd(11), NongliData(99), TianGan(9), DiZhi(11), ShuXiang(11), DayName(30), MonName(12)
Dim curTime, curYear, curMonth, curDay, curWeekday
Dim GongliStr, WeekdayStr, NongliStr, NongliDayStr
Dim i, m, n, k, isEnd, bit, TheDate


function Form_Load(Ytime,Ntime,yue,ri,nnian)
'��ȡ��ǰϵͳʱ��
'curTime = Now()
curTime = ytime

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
NongliStr = "ũ��" & TianGan(((curYear - 4) Mod 60) Mod 10) & DiZhi(((curYear - 4) Mod 60) Mod 12) & "��"
NongliStr = NongliStr & "(" & ShuXiang(((curYear - 4) Mod 60) Mod 12) & ")"

'����ũ���¡��� ==> NongliDayStr
If (curMonth < 1) Then
    NongliDayStr = "��" & MonName(-1 * curMonth)
Else
    NongliDayStr = MonName(curMonth)
End If
NongliDayStr = NongliDayStr & "��"

NongliDayStr = NongliDayStr & DayName(curDay)
if left(curtime,9)="2004-4-29" then
NTIME="ũ��������(��)����ʮһ"
else
NTIME=NongliStr & NongliDayStr
end if
yue=curmonth
ri=curday
nnian=curyear
End function




%>
<%
function nayin(gz)
set fos=server.CreateObject("scripting.filesystemobject")
path=server.MapPath(".")&"/nayin.txt"
set nys=fos.opentextfile(path)
while not nys.atendofstream
str=nys.readline
		  if instr(str,gz)>0 then
		  response.Write(right(str,6))
		  end if
		  
		  wend
		  '��������^_^��
		  set fos=nothing
end function
'����������
function taiy(x,y)
x1=(x+1)mod 10
y1=(y+3)mod 12
response.Write(tg(x1)&dz(y1))
end function
function mingg(dm,tdz,minggongx)
YueDz=((dm+11)mod 12)
TimeDz=((tdz+11) mod 12)
Mingdz=(26-YueDz-TimeDz)mod 12
mingdzx=mingdz
mingdz=(mingdz+1)mod 12
mingtg=(minggongx+mingdz+7)mod 10
response.Write(tg(mingtg)&dz(mingdz))
end function

function cidunayin(gz)
set fos=server.CreateObject("scripting.filesystemobject")
path=server.MapPath(".")&"/nayin.txt"
set nys=fos.opentextfile(path)
while not nys.atendofstream
str=nys.readline
		  if instr(str,gz)>0 then
		  cidunayin=(right(str,6))
		  end if
		  
		  wend
		  '��������^_^��
		  set fos=nothing
end function

%>
<%'ʮ��
function sshen(tgans,ritgs)
set ssob=server.CreateObject("scripting.filesystemobject")
path=server.MapPath("shishen.txt")
set ssf=ssob.opentextfile(path)
tgx=tgans
tgy=ritgs
Youb=tgy*10*3+tgx*3
ssf.skip(Youb)
str=ssf.read(3)
response.Write "<font color=blue>"&(trim(str))&"</font>"
ssf.close
set ssob=nothing
end function
'call sshen(3,4)
%>
<%'��ʮ��
function dsshen(tgstr,ritg)
TianGan=tgstr
rizg=ritg
for ik=1 to len(TianGan)
str=left(TianGan,1)
for ij=0 to 9 
if str=tg(ij) then
icount=ij
exit for
call sshen(icount,riztg)
end if
next
next
end function
%>

<%
function DZorder(GzX)  '���ַ���
 DZs=right(gzx,1)
for i= 0 to 11
if dz(i)=dzs then
DzOrder=i+1
exit for
end if
next
end function

function Tgorder(GzX)  '��������
tgs=left(gzx,1)
for i= 0 to 9
if tg(i)=tgs then
tgOrder=i+1
exit for
end if
next
end function
%>

<%
function nayins(gz)
set fos=server.CreateObject("scripting.filesystemobject")
path=server.MapPath(".")&"/nayin.txt"
set nys=fos.opentextfile(path)
while not nys.atendofstream
str=nys.readline
		  if instr(str,gz)>0 then
		  nayins=right(str,6)
		  end if
		  
		  wend
		  '��������^_^��
		  set fos=nothing
end function

%>
