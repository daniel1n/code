<%
'去除输入字符中的空格---newtrim函数
function newtrim(str)
dim result
dim j
j=len(str)
result=""
dim i
for i=1 to j
select case mid(str,i,1)
case chr(8)
     result=result+"" '去掉退格
case chr(9)
     result=result+"" '去掉tab
case chr(10)
     result=result+"" '去掉换行
case chr(13)
     result=result+"" '去掉回车
case chr(255)
     result=result+"" '去掉特殊空格
case else
     result=result+mid(str,i,1)
end select
next
newtrim=result
end function
%>
<%
'得到一个汉字的笔画数---getnum()函数
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
'由天格、人格、地格数得到三才---getsancai()函数
function getsancai(tiange,renge,dige)
dim tian,ren,di,tiantxt,rentxt,ditxt,total
tian=(tiange mod 10)
ren=(renge mod 10)
di=(dige mod 10)
select case tian
  case 1
  tiantxt="木"
  case 2
  tiantxt="木"
  case 3
  tiantxt="火"
  case 4
  tiantxt="火"
  case 5
  tiantxt="土"
  case 6
  tiantxt="土"
  case 7
  tiantxt="金"
  case 8
  tiantxt="金"
  case 9
  tiantxt="水"
  case 10
  tiantxt="水"
  case 0
  tiantxt="水"
end select
select case ren
  case 1
  rentxt="木"
  case 2
  rentxt="木"
  case 3
  rentxt="火"
  case 4
  rentxt="火"
  case 5
  rentxt="土"
  case 6
  rentxt="土"
  case 7
  rentxt="金"
  case 8
  rentxt="金"
  case 9
  rentxt="水"
  case 10
  rentxt="水"
  case 0
  rentxt="水"
end select
select case di
  case 1
  ditxt="木"
  case 2
  ditxt="木"
  case 3
  ditxt="火"
  case 4
  ditxt="火"
  case 5
  ditxt="土"
  case 6
  ditxt="土"
  case 7
  ditxt="金"
  case 8
  ditxt="金"
  case 9
  ditxt="水"
  case 10
  ditxt="水"
  case 0
  ditxt="水"
end select
total=tiantxt+rentxt+ditxt
getsancai=total
end function
%>
<%
Function hhcal(caltime) 
'公历转换农历的算法 edit by huanghui
Dim WeekName(7), MonthAdd(11), NongliData(99), TianGan(9), DiZhi(11), ShuXiang(11), DayName(30), MonName(12) 
Dim curTime, curYear, curMonth, curDay, curWeekday 
Dim GongliStr, WeekdayStr, NongliStr, NongliDayStr 
Dim i, m, n, k, isEnd, bit, TheDate 
'获取当前系统时间 
curTime = caltime 
'星期名 
WeekName(0) = " * " 
WeekName(1) = "星期日" 
WeekName(2) = "星期一" 
WeekName(3) = "星期二" 
WeekName(4) = "星期三" 
WeekName(5) = "星期四" 
WeekName(6) = "星期五" 
WeekName(7) = "星期六" 
'天干名称 
TianGan(0) = "甲" 
TianGan(1) = "乙" 
TianGan(2) = "丙" 
TianGan(3) = "丁" 
TianGan(4) = "戊" 
TianGan(5) = "己" 
TianGan(6) = "庚" 
TianGan(7) = "辛" 
TianGan(8) = "壬" 
TianGan(9) = "癸" 
'地支名称 
DiZhi(0) = "子" 
DiZhi(1) = "丑" 
DiZhi(2) = "寅" 
DiZhi(3) = "卯" 
DiZhi(4) = "辰" 
DiZhi(5) = "巳" 
DiZhi(6) = "午" 
DiZhi(7) = "未" 
DiZhi(8) = "申" 
DiZhi(9) = "酉" 
DiZhi(10) = "戌" 
DiZhi(11) = "亥" 
'属相名称 
ShuXiang(0) = "鼠" 
ShuXiang(1) = "牛" 
ShuXiang(2) = "虎" 
ShuXiang(3) = "兔" 
ShuXiang(4) = "龙" 
ShuXiang(5) = "蛇" 
ShuXiang(6) = "马" 
ShuXiang(7) = "羊" 
ShuXiang(8) = "猴" 
ShuXiang(9) = "鸡" 
ShuXiang(10) = "狗" 
ShuXiang(11) = "猪" 
'农历日期名 
DayName(0) = "*" 
DayName(1) = "初一" 
DayName(2) = "初二" 
DayName(3) = "初三" 
DayName(4) = "初四" 
DayName(5) = "初五" 
DayName(6) = "初六" 
DayName(7) = "初七" 
DayName(8) = "初八" 
DayName(9) = "初九" 
DayName(10) = "初十" 
DayName(11) = "十一" 
DayName(12) = "十二" 
DayName(13) = "十三" 
DayName(14) = "十四" 
DayName(15) = "十五" 
DayName(16) = "十六" 
DayName(17) = "十七" 
DayName(18) = "十八" 
DayName(19) = "十九" 
DayName(20) = "二十" 
DayName(21) = "廿一" 
DayName(22) = "廿二" 
DayName(23) = "廿三" 
DayName(24) = "廿四" 
DayName(25) = "廿五" 
DayName(26) = "廿六" 
DayName(27) = "廿七" 
DayName(28) = "廿八" 
DayName(29) = "廿九" 
DayName(30) = "三十" 
'农历月份名 
MonName(0) = "*" 
MonName(1) = "正" 
MonName(2) = "二" 
MonName(3) = "三" 
MonName(4) = "四" 
MonName(5) = "五" 
MonName(6) = "六" 
MonName(7) = "七" 
MonName(8) = "八" 
MonName(9) = "九" 
MonName(10) = "十" 
MonName(11) = "十一" 
MonName(12) = "腊" 
'公历每月前面的天数 
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
'农历数据 
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
'生成当前公历年、月、日 ==> GongliStr 
curYear = Year(curTime) 
curMonth = Month(curTime) 
curDay = Day(curTime) 
GongliStr = curYear & "年" 
If (curMonth < 10) Then 
    GongliStr = GongliStr & "0" & curMonth & "月" 
Else 
    GongliStr = GongliStr & curMonth & "月" 
End If 
If (curDay < 10) Then 
    GongliStr = GongliStr & "0" & curDay & "日" 
Else 
    GongliStr = GongliStr & curDay & "日" 
End If 
'生成当前公历星期 ==> WeekdayStr 
curWeekday = Weekday(curTime) 
WeekdayStr = WeekName(curWeekday) 
'计算到初始时间1921年2月8日的天数：1921-2-8(正月初一) 
TheDate = (curYear - 1921) * 365 + Int((curYear - 1921) / 4) + curDay + MonthAdd(curMonth - 1) - 38 
If ((curYear Mod 4) = 0 And curMonth > 2) Then 
    TheDate = TheDate + 1 
End If 
'计算农历天干、地支、月、日 
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
    '获取NongliData(m)的第n个二进制位的值 
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
'生成农历天干、地支、属相 ==> NongliStr
dim nongli,nlnian,nlyue,nlri,sx
nlnian = TianGan(((curYear - 4) Mod 60) Mod 10) & DiZhi(((curYear - 4) Mod 60) Mod 12) 
sx =ShuXiang(((curYear - 4) Mod 60) Mod 12)
'生成农历月、日 ==> NongliDayStr 
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
if not isdate(mDate) then response.write("非日期")
dim a
a=(Day(mDate) - (19 + Int(Mid("102123444423", Month(mDate), 1))))
if a>=0 then
a=0
else
a=-1
End if
'星座
    Constellation = Mid("魔羯水瓶双鱼白羊金牛双子巨蟹狮子处女天秤天蝎射手魔羯", (Month(mDate) + a) * 2 + 1, 2) & "座"
End Function%>
<%
'替换换行符的过滤函数
Function rep(content)
rep=replace(content,vbCrLf,"<br>")
end Function
%><%function GbToBig(content)
dim s,t,c,d,i
s="皑,蔼,碍,爱,翱,袄,奥,坝,罢,摆,败,颁,办,绊,帮,绑,镑,谤,剥,饱,宝,报,鲍,辈,贝,钡,狈,备,惫,绷,笔,毕,毙,闭,边,编,贬,变,辩,辫,鳖,瘪,濒,滨,宾,摈,饼,拨,钵,铂,驳,卜,补,参,蚕,残,惭,惨,灿,苍,舱,仓,沧,厕,侧,册,测,层,诧,搀,掺,蝉,馋,谗,缠,铲,产,阐,颤,场,尝,长,偿,肠,厂,畅,钞,车,彻,尘,陈,衬,撑,称,惩,诚,骋,痴,迟,驰,耻,齿,炽,冲,虫,宠,畴,踌,筹,绸,丑,橱,厨,锄,雏,础,储,触,处,传,疮,闯,创,锤,纯,绰,辞,词,赐,聪,葱,囱,从,丛,凑,窜,错,达,带,贷,担,单,郸,掸,胆,惮,诞,弹,当,挡,党,荡,档,捣,岛,祷,导,盗,灯,邓,敌,涤,递,缔,点,垫,电,淀,钓,调,迭,谍,叠,钉,顶,锭,订,东,动,栋,冻,斗,犊,独,读,赌,镀,锻,断,缎,兑,队,对,吨,顿,钝,夺,鹅,额,讹,恶,饿,儿,尔,饵,贰,发,罚,阀,珐,矾,钒,烦,范,贩,饭,访,纺,飞,废,费,纷,坟,奋,愤,粪,丰,枫,锋,风,疯,冯,缝,讽,凤,肤,辐,抚,辅,赋,复,负,讣,妇,缚,该,钙,盖,干,赶,秆,赣,冈,刚,钢,纲,岗,皋,镐,搁,鸽,阁,铬,个,给,龚,宫,巩,贡,钩,沟,构,购,够,蛊,顾,剐,关,观,馆,惯,贯,广,规,硅,归,龟,闺,轨,诡,柜,贵,刽,辊,滚,锅,国,过,骇,韩,汉,阂,鹤,贺,横,轰,鸿,红,后,壶,护,沪,户,哗,华,画,划,话,怀,坏,欢,环,还,缓,换,唤,痪,焕,涣,黄,谎,挥,辉,毁,贿,秽,会,烩,汇,讳,诲,绘,荤,浑,伙,获,货,祸,击,机,积,饥,讥,鸡,绩,缉,极,辑,级,挤,几,蓟,剂,济,计,记,际,继,纪,夹,荚,颊,贾,钾,价,驾,歼,监,坚,笺,间,艰,缄,茧,检,碱,硷,拣,捡,简,俭,减,荐,槛,鉴,践,贱,见,键,舰,剑,饯,渐,溅,涧,浆,蒋,桨,奖,讲,酱,胶,浇,骄,娇,搅,铰,矫,侥,脚,饺,缴,绞,轿,较,秸,阶,节,茎,惊,经,颈,静,镜,径,痉,竞,净,纠,厩,旧,驹,举,据,锯,惧,剧,鹃,绢,杰,洁,结,诫,届,紧,锦,仅,谨,进,晋,烬,尽,劲,荆,觉,决,诀,绝,钧,军,骏,开,凯,颗,壳,课,垦,恳,抠,库,裤,夸,块,侩,宽,矿,旷,况,亏,岿,窥,馈,溃,扩,阔,蜡,腊,莱,来,赖,蓝,栏,拦,篮,阑,兰,澜,谰,揽,览,懒,缆,烂,滥,捞,劳,涝,乐,镭,垒,类,泪,篱,离,里,鲤,礼,丽,厉,励,砾,历,沥,隶,俩,联,莲,连,镰,怜,涟,帘,敛,脸,链,恋,炼,练,粮,凉,两,辆,谅,疗,辽,镣,猎,临,邻,鳞,凛,赁,龄,铃,凌,灵,岭,领,馏,刘,龙,聋,咙,笼,垄,拢,陇,楼,娄,搂,篓,芦,卢,颅,庐,炉,掳,卤,虏,鲁,赂,禄,录,陆,驴,吕,铝,侣,屡,缕,虑,滤,绿,峦,挛,孪,滦,乱,抡,轮,伦,仑,沦,纶,论,萝,罗,逻,锣,箩,骡,骆,络,妈,玛,码,蚂,马,骂,吗,买,麦,卖,迈,脉,瞒,馒,蛮,满,谩,猫,锚,铆,贸,么,霉,没,镁,门,闷,们,锰,梦,谜,弥,觅,绵,缅,庙,灭,悯,闽,鸣,铭,谬,谋,亩,钠,纳,难,挠,脑,恼,闹,馁,腻,撵,捻,酿,鸟,聂,啮,镊,镍,柠,狞,宁,拧,泞,钮,纽,脓,浓,农,疟,诺,欧,鸥,殴,呕,沤,盘,庞,国,爱,赔,喷,鹏,骗,飘,频,贫,苹,凭,评,泼,颇,扑,铺,朴,谱,脐,齐,骑,岂,启,气,弃,讫,牵,扦,钎,铅,迁,签,谦,钱,钳,潜,浅,谴,堑,枪,呛,墙,蔷,强,抢,锹,桥,乔,侨,翘,窍,窃,钦,亲,轻,氢,倾,顷,请,庆,琼,穷,趋,区,躯,驱,龋,颧,权,劝,却,鹊,让,饶,扰,绕,热,韧,认,纫,荣,绒,软,锐,闰,润,洒,萨,鳃,赛,伞,丧,骚,扫,涩,杀,纱,筛,晒,闪,陕,赡,缮,伤,赏,烧,绍,赊,摄,慑,设,绅,审,婶,肾,渗,声,绳,胜,圣,师,狮,湿,诗,尸,时,蚀,实,识,驶,势,释,饰,视,试,寿,兽,枢,输,书,赎,属,术,树,竖,数,帅,双,谁,税,顺,说,硕,烁,丝,饲,耸,怂,颂,讼,诵,擞,苏,诉,肃,虽,绥,岁,孙,损,笋,缩,琐,锁,獭,挞,抬,摊,贪,瘫,滩,坛,谭,谈,叹,汤,烫,涛,绦,腾,誊,锑,题,体,屉,条,贴,铁,厅,听,烃,铜,统,头,图,涂,团,颓,蜕,脱,鸵,驮,驼,椭,洼,袜,弯,湾,顽,万,网,韦,违,围,为,潍,维,苇,伟,伪,纬,谓,卫,温,闻,纹,稳,问,瓮,挝,蜗,涡,窝,呜,钨,乌,诬,无,芜,吴,坞,雾,务,误,锡,牺,袭,习,铣,戏,细,虾,辖,峡,侠,狭,厦,锨,鲜,纤,咸,贤,衔,闲,显,险,现,献,县,馅,羡,宪,线,厢,镶,乡,详,响,项,萧,销,晓,啸,蝎,协,挟,携,胁,谐,写,泻,谢,锌,衅,兴,汹,锈,绣,虚,嘘,须,许,绪,续,轩,悬,选,癣,绚,学,勋,询,寻,驯,训,讯,逊,压,鸦,鸭,哑,亚,讶,阉,烟,盐,严,颜,阎,艳,厌,砚,彦,谚,验,鸯,杨,扬,疡,阳,痒,养,样,瑶,摇,尧,遥,窑,谣,药,爷,页,业,叶,医,铱,颐,遗,仪,彝,蚁,艺,亿,忆,义,诣,议,谊,译,异,绎,荫,阴,银,饮,樱,婴,鹰,应,缨,莹,萤,营,荧,蝇,颖,哟,拥,佣,痈,踊,咏,涌,优,忧,邮,铀,犹,游,诱,舆,鱼,渔,娱,与,屿,语,吁,御,狱,誉,预,驭,鸳,渊,辕,园,员,圆,缘,远,愿,约,跃,钥,岳,粤,悦,阅,云,郧,匀,陨,运,蕴,酝,晕,韵,杂,灾,载,攒,暂,赞,赃,脏,凿,枣,灶,责,择,则,泽,贼,赠,扎,札,轧,铡,闸,诈,斋,债,毡,盏,斩,辗,崭,栈,战,绽,张,涨,帐,账,胀,赵,蛰,辙,锗,这,贞,针,侦,诊,镇,阵,挣,睁,狰,帧,郑,证,织,职,执,纸,挚,掷,帜,质,钟,终,种,肿,众,诌,轴,皱,昼,骤,猪,诸,诛,烛,瞩,嘱,贮,铸,筑,驻,专,砖,转,赚,桩,庄,装,妆,壮,状,锥,赘,坠,缀,谆,浊,兹,资,渍,踪,综,总,纵,邹,诅,组,钻,致,钟,么,为,只,凶,准,启,板,里,雳,余,链,泄"
t="皚,藹,礙,愛,翺,襖,奧,壩,罷,擺,敗,頒,辦,絆,幫,綁,鎊,謗,剝,飽,寶,報,鮑,輩,貝,鋇,狽,備,憊,繃,筆,畢,斃,閉,邊,編,貶,變,辯,辮,鼈,癟,瀕,濱,賓,擯,餅,撥,缽,鉑,駁,蔔,補,參,蠶,殘,慚,慘,燦,蒼,艙,倉,滄,廁,側,冊,測,層,詫,攙,摻,蟬,饞,讒,纏,鏟,産,闡,顫,場,嘗,長,償,腸,廠,暢,鈔,車,徹,塵,陳,襯,撐,稱,懲,誠,騁,癡,遲,馳,恥,齒,熾,沖,蟲,寵,疇,躊,籌,綢,醜,櫥,廚,鋤,雛,礎,儲,觸,處,傳,瘡,闖,創,錘,純,綽,辭,詞,賜,聰,蔥,囪,從,叢,湊,竄,錯,達,帶,貸,擔,單,鄲,撣,膽,憚,誕,彈,當,擋,黨,蕩,檔,搗,島,禱,導,盜,燈,鄧,敵,滌,遞,締,點,墊,電,澱,釣,調,叠,諜,疊,釘,頂,錠,訂,東,動,棟,凍,鬥,犢,獨,讀,賭,鍍,鍛,斷,緞,兌,隊,對,噸,頓,鈍,奪,鵝,額,訛,惡,餓,兒,爾,餌,貳,發,罰,閥,琺,礬,釩,煩,範,販,飯,訪,紡,飛,廢,費,紛,墳,奮,憤,糞,豐,楓,鋒,風,瘋,馮,縫,諷,鳳,膚,輻,撫,輔,賦,複,負,訃,婦,縛,該,鈣,蓋,幹,趕,稈,贛,岡,剛,鋼,綱,崗,臯,鎬,擱,鴿,閣,鉻,個,給,龔,宮,鞏,貢,鈎,溝,構,購,夠,蠱,顧,剮,關,觀,館,慣,貫,廣,規,矽,歸,龜,閨,軌,詭,櫃,貴,劊,輥,滾,鍋,國,過,駭,韓,漢,閡,鶴,賀,橫,轟,鴻,紅,後,壺,護,滬,戶,嘩,華,畫,劃,話,懷,壞,歡,環,還,緩,換,喚,瘓,煥,渙,黃,謊,揮,輝,毀,賄,穢,會,燴,彙,諱,誨,繪,葷,渾,夥,獲,貨,禍,擊,機,積,饑,譏,雞,績,緝,極,輯,級,擠,幾,薊,劑,濟,計,記,際,繼,紀,夾,莢,頰,賈,鉀,價,駕,殲,監,堅,箋,間,艱,緘,繭,檢,堿,鹼,揀,撿,簡,儉,減,薦,檻,鑒,踐,賤,見,鍵,艦,劍,餞,漸,濺,澗,漿,蔣,槳,獎,講,醬,膠,澆,驕,嬌,攪,鉸,矯,僥,腳,餃,繳,絞,轎,較,稭,階,節,莖,驚,經,頸,靜,鏡,徑,痙,競,淨,糾,廄,舊,駒,舉,據,鋸,懼,劇,鵑,絹,傑,潔,結,誡,屆,緊,錦,僅,謹,進,晉,燼,盡,勁,荊,覺,決,訣,絕,鈞,軍,駿,開,凱,顆,殼,課,墾,懇,摳,庫,褲,誇,塊,儈,寬,礦,曠,況,虧,巋,窺,饋,潰,擴,闊,蠟,臘,萊,來,賴,藍,欄,攔,籃,闌,蘭,瀾,讕,攬,覽,懶,纜,爛,濫,撈,勞,澇,樂,鐳,壘,類,淚,籬,離,裏,鯉,禮,麗,厲,勵,礫,曆,瀝,隸,倆,聯,蓮,連,鐮,憐,漣,簾,斂,臉,鏈,戀,煉,練,糧,涼,兩,輛,諒,療,遼,鐐,獵,臨,鄰,鱗,凜,賃,齡,鈴,淩,靈,嶺,領,餾,劉,龍,聾,嚨,籠,壟,攏,隴,樓,婁,摟,簍,蘆,盧,顱,廬,爐,擄,鹵,虜,魯,賂,祿,錄,陸,驢,呂,鋁,侶,屢,縷,慮,濾,綠,巒,攣,孿,灤,亂,掄,輪,倫,侖,淪,綸,論,蘿,羅,邏,鑼,籮,騾,駱,絡,媽,瑪,碼,螞,馬,罵,嗎,買,麥,賣,邁,脈,瞞,饅,蠻,滿,謾,貓,錨,鉚,貿,麽,黴,沒,鎂,門,悶,們,錳,夢,謎,彌,覓,綿,緬,廟,滅,憫,閩,鳴,銘,謬,謀,畝,鈉,納,難,撓,腦,惱,鬧,餒,膩,攆,撚,釀,鳥,聶,齧,鑷,鎳,檸,獰,甯,擰,濘,鈕,紐,膿,濃,農,瘧,諾,歐,鷗,毆,嘔,漚,盤,龐,國,愛,賠,噴,鵬,騙,飄,頻,貧,蘋,憑,評,潑,頗,撲,鋪,樸,譜,臍,齊,騎,豈,啓,氣,棄,訖,牽,扡,釺,鉛,遷,簽,謙,錢,鉗,潛,淺,譴,塹,槍,嗆,牆,薔,強,搶,鍬,橋,喬,僑,翹,竅,竊,欽,親,輕,氫,傾,頃,請,慶,瓊,窮,趨,區,軀,驅,齲,顴,權,勸,卻,鵲,讓,饒,擾,繞,熱,韌,認,紉,榮,絨,軟,銳,閏,潤,灑,薩,鰓,賽,傘,喪,騷,掃,澀,殺,紗,篩,曬,閃,陝,贍,繕,傷,賞,燒,紹,賒,攝,懾,設,紳,審,嬸,腎,滲,聲,繩,勝,聖,師,獅,濕,詩,屍,時,蝕,實,識,駛,勢,釋,飾,視,試,壽,獸,樞,輸,書,贖,屬,術,樹,豎,數,帥,雙,誰,稅,順,說,碩,爍,絲,飼,聳,慫,頌,訟,誦,擻,蘇,訴,肅,雖,綏,歲,孫,損,筍,縮,瑣,鎖,獺,撻,擡,攤,貪,癱,灘,壇,譚,談,歎,湯,燙,濤,縧,騰,謄,銻,題,體,屜,條,貼,鐵,廳,聽,烴,銅,統,頭,圖,塗,團,頹,蛻,脫,鴕,馱,駝,橢,窪,襪,彎,灣,頑,萬,網,韋,違,圍,爲,濰,維,葦,偉,僞,緯,謂,衛,溫,聞,紋,穩,問,甕,撾,蝸,渦,窩,嗚,鎢,烏,誣,無,蕪,吳,塢,霧,務,誤,錫,犧,襲,習,銑,戲,細,蝦,轄,峽,俠,狹,廈,鍁,鮮,纖,鹹,賢,銜,閑,顯,險,現,獻,縣,餡,羨,憲,線,廂,鑲,鄉,詳,響,項,蕭,銷,曉,嘯,蠍,協,挾,攜,脅,諧,寫,瀉,謝,鋅,釁,興,洶,鏽,繡,虛,噓,須,許,緒,續,軒,懸,選,癬,絢,學,勳,詢,尋,馴,訓,訊,遜,壓,鴉,鴨,啞,亞,訝,閹,煙,鹽,嚴,顔,閻,豔,厭,硯,彥,諺,驗,鴦,楊,揚,瘍,陽,癢,養,樣,瑤,搖,堯,遙,窯,謠,藥,爺,頁,業,葉,醫,銥,頤,遺,儀,彜,蟻,藝,億,憶,義,詣,議,誼,譯,異,繹,蔭,陰,銀,飲,櫻,嬰,鷹,應,纓,瑩,螢,營,熒,蠅,穎,喲,擁,傭,癰,踴,詠,湧,優,憂,郵,鈾,猶,遊,誘,輿,魚,漁,娛,與,嶼,語,籲,禦,獄,譽,預,馭,鴛,淵,轅,園,員,圓,緣,遠,願,約,躍,鑰,嶽,粵,悅,閱,雲,鄖,勻,隕,運,蘊,醞,暈,韻,雜,災,載,攢,暫,贊,贓,髒,鑿,棗,竈,責,擇,則,澤,賊,贈,紮,劄,軋,鍘,閘,詐,齋,債,氈,盞,斬,輾,嶄,棧,戰,綻,張,漲,帳,賬,脹,趙,蟄,轍,鍺,這,貞,針,偵,診,鎮,陣,掙,睜,猙,幀,鄭,證,織,職,執,紙,摯,擲,幟,質,鍾,終,種,腫,衆,謅,軸,皺,晝,驟,豬,諸,誅,燭,矚,囑,貯,鑄,築,駐,專,磚,轉,賺,樁,莊,裝,妝,壯,狀,錐,贅,墜,綴,諄,濁,茲,資,漬,蹤,綜,總,縱,鄒,詛,組,鑽,緻,鐘,麼,為,隻,兇,準,啟,闆,裡,靂,餘,鍊,洩"
c=split(s,",")
d=split(t,",")
for i=0 to 1274
content=replace(content,c(i),d(i))
next
GbToBig=content
end function%>
<%
'得到一个汉字字意五行
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
  sctxt="木"
  case 2
  sctxt="木"
  case 3
  sctxt="火"
  case 4
  sctxt="火"
  case 5
  sctxt="土"
  case 6
  sctxt="土"
  case 7
  sctxt="金"
  case 8
  sctxt="金"
  case 9
  sctxt="水"
  case 10
  sctxt="水"
  case 0
  sctxt="水"
end select
getsancai=sctxt
end function%><%function DiZhi(i)
select case i
case 0
dz="子"
case 1
dz="丑"
case 2
dz="丑"
case 3
dz="寅"
case 4
dz="寅"
case 5
dz="卯"
case 6
dz="卯"
case 7
dz="辰"
case 8
dz="辰"
case 9
dz="巳"
case 10
dz="巳"
case 11
dz="午"
case 12
dz="午"
case 13
dz="未"
case 14
dz="未"
case 15
dz="申"
case 16
dz="申"
case 17
dz="酉"
case 18
dz="酉"
case 19
dz="戌"
case 20
dz="戌"
case 21
dz="亥"
case 22
dz="亥"
case 23
dz="子"
end select
DiZhi=dz
end function%>
<%function getpf(sc)
select case sc
  case "大吉"
  szpf=12
  case "吉"
  szpf=8
  case "半吉"
  szpf=5
  case "大凶"
  szpf=0
  case "凶"
  szpf=1
  case "半凶"
  szpf=2
  case "平"
  szpf=4
end select
getpf=szpf
end function%><%
'纳音
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
  case "子"
  wh="水"
  case "亥"
  wh="水"
  case "寅"
  wh="木"
  case "卯"
  wh="木"
  case "巳"
  wh="火"
  case "午"
  wh="火"
  case "申"
  wh="金"
  case "酉"
  wh="金"
  case "辰"
  wh="土"
  case "戌"
  wh="土"
  case "丑"
  wh="土"
  case "未"
  wh="土"
    case "甲"
  wh="木"
    case "乙"
  wh="木"
    case "丙"
  wh="火"
    case "丁"
  wh="火"
    case "戊"
  wh="土"
    case "己"
  wh="土"
    case "庚"
  wh="金"
      case "辛"
  wh="金"
      case "壬"
  wh="水"
      case "癸"
  wh="水"
end select
tgdzwh=wh
end function%>
<%function siji(yue)
select case yue
  case "1"
  sj="冬"
  case "2"
  sj="冬"
  case "3"
  sj="春"
  case "4"
  sj="春"
  case "5"
  sj="春"
  case "6"
  sj="夏"
  case "7"
  sj="夏"
  case "8"
  sj="夏"
  case "9"
  sj="秋"
  case "10"
  sj="秋"
  case "11"
  sj="秋"
  case "12"
  sj="冬"
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
