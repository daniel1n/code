<%
Dim yearlast
Dim day60
Dim a(1000)
Dim nlsc(100)
Dim sczy(100)
Dim scsy(100)
Dim jz(100)
dim  dayname(31)
dim MonName(13)
dim rn


'ʮ���

a(21) = "��"
a(22) = "��"
a(23) = "��"
a(24) = "��"
a(25) = "��"
a(26) = "��"
a(27) = "��"
a(28) = "��"
a(29) = "��"
a(20) = "��"

'ʮ����֧

a(31) = "��"
a(32) = "��"
a(33) = "��"
a(34) = "î"
a(35) = "��"
a(36) = "��"
a(37) = "��"
a(38) = "δ"
a(39) = "��"
a(40) = "��"
a(41) = "��"
a(30) = "��"

'���ˣ������ũ����������ת��

bzyear=nian1
bzmonth=yue1
bzday=ri1
bztime=hh1

dim md
'ȷ������ yz ��֧  ���˻��� qyjs
md=bzmonth * 100 + bzday
if(md>=204 and md<= 305) then
mz = 3
qyjs = ((bzmonth - 2) * 30 + bzday - 4) \ 3
end if

if(md>=306 and md<=404) then
mz = 4
qyjs = ((bzmonth - 3) * 30 + bzday - 6) \ 3
end if

if(md>=405 and md<= 504) then
mz = 5
qyjs = ((bzmonth - 4) * 30 + bzday - 5) \ 3
end if

if(md>=505 and md<=  605) then
mz = 6
qyjs = ((bzmonth - 5) * 30 + bzday - 5) \ 3
end if

if(md>=606 and md<= 706) then
mz = 7
qyjs = ((bzmonth - 6) * 30 + bzday - 6) \ 3
end if

if(md>=707 and md<= 807) then
mz = 8
qyjs = ((bzmonth - 7) * 30 + bzday - 7) \ 3
end if

if(md>=808 and md<=907) then
mz = 9
qyjs = ((bzmonth - 8) * 30 + bzday - 8) \ 3
end if

if(md>=908 and md<=1007) then
mz = 10
qyjs = ((bzmonth - 9) * 30 + bzday - 8) \ 3
end if

if(md>=1008 and md<= 1106) then
mz = 11
qyjs = ((bzmonth - 10) * 30 + bzday - 8) \ 3
end if

if(md>=1107 and md<=  1207) then
mz = 0
qyjs = ((bzmonth - 11) * 30 + bzday - 7) \ 3
end if

if(md>=1208 and md<=  1231) then
mz = 1
qyjs = (bzday - 8) \ 3
end if

if(md>=101 and md<= 105) then
mz = 1
qyjs = (30 + bzday - 4) \ 3
end if

if(md>=106 and md<=  203) then
mz = 2
qyjs = ((bzmonth - 1) * 30 + bzday - 6) \ 3
end if

'ȷ����ɺ���֧ yg ��� yz ��֧
if(md>=204 and md<=1231) then
yg = (bzyear - 3) Mod 10
yz = (bzyear - 3) Mod 12
end if
if(md>=101 and md<=203 ) then
yg = (bzyear - 4) Mod 10
yz = (bzyear - 4) Mod 12
end if

'ȷ���¸� mg �¸�

If (mz > 2 And mz <= 11) Then

mg = (yg * 2 + mz + 8) Mod 10
Else
mg = (yg * 2 + mz) Mod 10
End If

'�ӹ�Ԫ0�굽Ŀǰ��ݵ����� yearlast

yearlast = (bzyear - 1) * 5 + (bzyear - 1) \ 4 - (bzyear - 1) \ 100 + (bzyear - 1) \ 400
'����ĳ��ĳ���뵱��1��0�յ�ʱ������Ϊ��λ��yearday

For i = 1 To bzmonth - 1
Select Case i
Case 1, 3, 5, 7, 8, 10, 12
yearday = yearday + 31
Case 4, 6, 9, 11
yearday = yearday + 30
Case 2
If (bzyear Mod 4 = 0) And ((bzyear Mod 100) <> 0) Or (bzyear Mod 400 = 0) Then
yearday = yearday + 29
Else
yearday = yearday + 28
End If
End Select

next

yearday = yearday + bzday

'�����յ���ʮ������ day60
day60 = (yearlast + yearday + 6015) Mod 60

'ȷ�� �ո� dg  ��֧  dz
dg = day60 Mod 10
dz = day60 Mod 12
'ȷ�� ʱ�� tg ʱ֧ tz
tz = (bztime + 3) \ 2 Mod 12
'tg = (dg * 2 + tz + 8) Mod 10
If (tz = 0) Then
tg = (dg * 2 + tz) Mod 10

Else
tg = (dg * 2 + tz + 8) Mod 10
End If

'�� ������ʱ ת����Ϊ ��������
Select Case yz
Case 1
yzg = 0
Case 2, 8
yzg = 6
Case 3
yzg = 1
Case 4
yzg = 2
Case 5, 11
yzg = 5
Case 6
yzg = 3
Case 7
yzg = 4
Case 9
yzg = 7
Case 10
yzg = 8
Case 0
yzg = 9
End Select


'��֧�ɸ� mzg

Select Case mz
Case 1
mzg = 0
Case 2, 8
mzg = 6
Case 3
mzg = 1
Case 4
mzg = 2
Case 5, 11
mzg = 5
Case 6
mzg = 3
Case 7
mzg = 4
Case 9
mzg = 7
Case 10
mzg = 8
Case 0
mzg = 9
End Select

'��֧�ɸ� dzg

Select Case dz
Case 1
dzg = 0
Case 2, 8
dzg = 6
Case 3
dzg = 1
Case 4
dzg = 2
Case 5, 11
dzg = 5
Case 6
dzg = 3
Case 7
dzg = 4
Case 9
dzg = 7
Case 10
dzg = 8
Case 0
dzg = 9
End Select

'ʱ֧�ɸ� tzg

Select Case tz
Case 1
tzg = 0
Case 2, 8
tzg = 6
Case 3
tzg = 1
Case 4
tzg = 2
Case 5, 11
tzg = 5
Case 6
tzg = 3
Case 7
tzg = 4
Case 9
tzg = 7
Case 10
tzg = 8
Case 0
tzg = 9
End Select

'���ˣ���ɸ���֧������ɵ�ȷ������
yg1=a(20 + yg)
yz1=a(30 + yz)
mg1=a(20 + mg)
mz1=a(30 + mz)
dg1=a(20 + dg)
dz1=a(30 + dz)
tg1=a(20 + tg)
tz1=a(30 + tz)
ygz=a(20 + yg)&a(30 + yz)
mgz=a(20 + mg)&a(30 + mz)
dgz=a(20 + dg)&a(30 + dz)
tgz=a(20 + tg)&a(30 + tz)


'�������һ
bzy1=nly1
bzm1=nlm1
bzd1=nld1
bzh1=h1



dim mdn1
'ȷ������ yzn1 ��֧  ���˻��� qyjsn1
mdn1=bzm1 * 100 + bzd1
if(mdn1>=204 and mdn1<= 305) then
mzn1 = 3
qyjsn1 = ((bzm1 - 2) * 30 + bzd1 - 4) \ 3
end if

if(mdn1>=306 and mdn1<=404) then
mzn1 = 4
qyjsn1 = ((bzm1 - 3) * 30 + bzd1 - 6) \ 3
end if

if(mdn1>=405 and mdn1<= 504) then
mzn1 = 5
qyjsn1 = ((bzm1 - 4) * 30 + bzd1 - 5) \ 3
end if

if(mdn1>=505 and mdn1<=  605) then
mzn1 = 6
qyjsn1 = ((bzm1 - 5) * 30 + bzd1 - 5) \ 3
end if

if(mdn1>=606 and mdn1<= 706) then
mzn1 = 7
qyjsn1 = ((bzm1 - 6) * 30 + bzd1 - 6) \ 3
end if

if(mdn1>=707 and mdn1<= 807) then
mzn1 = 8
qyjsn1 = ((bzm1 - 7) * 30 + bzd1 - 7) \ 3
end if

if(mdn1>=808 and mdn1<=907) then
mzn1 = 9
qyjsn1 = ((bzm1 - 8) * 30 + bzd1 - 8) \ 3
end if

if(mdn1>=908 and mdn1<=1007) then
mzn1 = 10
qyjsn1 = ((bzm1 - 9) * 30 + bzd1 - 8) \ 3
end if

if(mdn1>=1008 and mdn1<= 1106) then
mzn1 = 11
qyjsn1 = ((bzm1 - 10) * 30 + bzd1 - 8) \ 3
end if

if(mdn1>=1107 and mdn1<=  1207) then
mzn1 = 0
qyjsn1 = ((bzm1 - 11) * 30 + bzd1 - 7) \ 3
end if

if(mdn1>=1208 and mdn1<=  1231) then
mzn1 = 1
qyjsn1 = (bzd1 - 8) \ 3
end if

if(mdn1>=101 and mdn1<= 105) then
mzn1 = 1
qyjsn1 = (30 + bzd1 - 4) \ 3
end if

if(mdn1>=106 and mdn1<=  203) then
mzn1 = 2
qyjsn1 = ((bzm1 - 1) * 30 + bzd1 - 6) \ 3
end if

'ȷ����ɺ���֧ ygn1 ��� yzn1 ��֧
if(mdn1>=204 and mdn1<=1231) then
ygn1 = (bzy1 - 3) Mod 10
yzn1 = (bzy1 - 3) Mod 12
end if
if(mdn1>=101 and mdn1<=203 ) then
ygn1 = (bzy1 - 4) Mod 10
yzn1 = (bzy1 - 4) Mod 12
end if

'ȷ���¸� mgn1 �¸�

If (mzn1 > 2 And mzn1 <= 11) Then

mgn1 = (ygn1 * 2 + mzn1 + 8) Mod 10
Else
mgn1 = (ygn1 * 2 + mzn1) Mod 10
End If

'�ӹ�Ԫ0�굽Ŀǰ��ݵ����� yearlast

yearlast = (bzy1 - 1) * 5 + (bzy1 - 1) \ 4 - (bzy1 - 1) \ 100 + (bzy1 - 1) \ 400
'����ĳ��ĳ���뵱��1��0�յ�ʱ������Ϊ��λ��yearday

For i = 1 To bzm1 - 1
Select Case i
Case 1, 3, 5, 7, 8, 10, 12
yearday = yearday + 31
Case 4, 6, 9, 11
yearday = yearday + 30
Case 2
If (bzy1 Mod 4 = 0) And ((bzy1 Mod 100) <> 0) Or (bzy1 Mod 400 = 0) Then
yearday = yearday + 29
Else
yearday = yearday + 28
End If
End Select

next

yearday = yearday + bzd1

'�����յ���ʮ������ day60
day60 = (yearlast + yearday + 6015) Mod 60

'ȷ�� �ո� dgn1  ��֧  dzn1
dgn1 = day60 Mod 10
dzn1 = day60 Mod 12
'ȷ�� ʱ�� tgn1 ʱ֧ tzn1
tzn1 = (bzh1 + 3) \ 2 Mod 12
'tgn1 = (dgn1 * 2 + tzn1 + 8) Mod 10
If (tzn1 = 0) Then
tgn1 = (dgn1 * 2 + tzn1) Mod 10

Else
tgn1 = (dgn1 * 2 + tzn1 + 8) Mod 10
End If

'�� ������ʱ ת����Ϊ ��������
Select Case yzn1
Case 1
yzgn1 = 0
Case 2, 8
yzgn1 = 6
Case 3
yzgn1 = 1
Case 4
yzgn1 = 2
Case 5, 11
yzgn1 = 5
Case 6
yzgn1 = 3
Case 7
yzgn1 = 4
Case 9
yzgn1 = 7
Case 10
yzgn1 = 8
Case 0
yzgn1 = 9
End Select


'��֧�ɸ� mzgn1

Select Case mzn1
Case 1
mzgn1 = 0
Case 2, 8
mzgn1 = 6
Case 3
mzgn1 = 1
Case 4
mzgn1 = 2
Case 5, 11
mzgn1 = 5
Case 6
mzgn1 = 3
Case 7
mzgn1 = 4
Case 9
mzgn1 = 7
Case 10
mzgn1 = 8
Case 0
mzgn1 = 9
End Select

'��֧�ɸ� dzgn1

Select Case dzn1
Case 1
dzgn1 = 0
Case 2, 8
dzgn1 = 6
Case 3
dzgn1 = 1
Case 4
dzgn1 = 2
Case 5, 11
dzgn1 = 5
Case 6
dzgn1 = 3
Case 7
dzgn1 = 4
Case 9
dzgn1 = 7
Case 10
dzgn1 = 8
Case 0
dzgn1 = 9
End Select

'ʱ֧�ɸ� tzgn1

Select Case tzn1
Case 1
tzgn1 = 0
Case 2, 8
tzgn1 = 6
Case 3
tzgn1 = 1
Case 4
tzgn1 = 2
Case 5, 11
tzgn1 = 5
Case 6
tzgn1 = 3
Case 7
tzgn1 = 4
Case 9
tzgn1 = 7
Case 10
tzgn1 = 8
Case 0
tzgn1 = 9
End Select

'���ˣ���ɸ���֧������ɵ�ȷ������
yg1n1=a(20 + ygn1)
yz1n1=a(30 + yzn1)
mg1n1=a(20 + mgn1)
mz1n1=a(30 + mzn1)
dg1n1=a(20 + dgn1)
dz1n1=a(30 + dzn1)
tg1n1=a(20 + tgn1)
tz1n1=a(30 + tzn1)
ygzn1=a(20 + ygn1)&a(30 + yzn1)
mgzn1=a(20 + mgn1)&a(30 + mzn1)
dgzn1=a(20 + dgn1)&a(30 + dzn1)
tgzn1=a(20 + tgn1)&a(30 + tzn1)%>