<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0" class="b1" style="table-layout:fixed;word-wrap:break-word;">
    <TBODY>
<tr>
  <td class=new><span class="style2">
    <%
sub bz()

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
DayName(11) = "十一;"
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
'十神名称

a(1) = "比肩"
a(2) = "劫财"
a(3) = "食神"
a(4) = "伤官"
a(5) = "偏财"
a(6) = "正财"
a(7) = "七杀"
a(8) = "正官"
a(9) = "偏印"
a(0) = "正印"

'十天干

a(21) = "甲"
a(22) = "乙"
a(23) = "丙"
a(24) = "丁"
a(25) = "戊"
a(26) = "己"
a(27) = "庚"
a(28) = "辛"
a(29) = "壬"
a(20) = "癸"

'十二地支

a(31) = "子 "
a(32) = "丑 "
a(33) = "寅 "
a(34) = "卯 "
a(35) = "辰 "
a(36) = "巳 "
a(37) = "午 "
a(38) = "未 "
a(39) = "申 "
a(40) = "酉 "
a(41) = "戌 "
a(30) = "亥 "

'各种神煞的定义

a(11) = "夫妻恩爱不同凡响,互相爱幕如宾,如胶似膝,形影不离."
'/*内桃花*/
a(12) = "其人风流多情,难守贞操,多有外遇."
'/*外桃花*/
a(13) = "能得贵人,朋友的帮忙,使自己受益不浅."
'/*天乙*/
a(14) = "一生吉利,荣华富贵."
'/*天月二德*/
a(15) = "一生劳碌奔波"
'/*岁禄*/
a(16) = "衣食丰足,享受国家奉禄,得妻子相助"
'/*坐禄*/
a(17) = "晚年生活富态,坐享儿孙之福"
'/*归禄*/
a(18) = "诗书满腹,少年勤奋有功名之格"
'/*学堂*/
a(19) = "太平吉祥,处事少忧,即使犯了大错也会得到宽恕"
'/*天赫吉星*/

a(43) = "生平之中，收入虽多，开销亦大，较易财来财去而不易聚财。"
'/*从财格*/
a(44) = "你的头脑灵活，谋财手腕高明，最易致富。"
'/*伤官逢财*/
a(45) = "你可以从事投机性行业赚得偏财。"
'/*偏财得用*/
a(46) = "你的钱财都是一分耕耘，一分收获的得来的。"
'/*正财得用*/
a(47) = "你与异性之关系密切，异性缘佳，受异性之影响力大。"
a(50) = "你与异性之关系较不密切，受异性影响较小。"
'/*正偏财*/
a(51) = "得祖先之福荫，继承父亲之产业。"
'/*年正官*/
a(52) = "必受父母之所爱，一生少劳苦。"
'/*月正官*/
a(53) = "上有兄姊，有独立分家的倾向。"
'/*年比肩*/
a(54) = "有兄弟姊妹是养子之倾向，有独立性的事业，争财，理财之特性。"
'/*月比肩*/
a(55) = "子孙得力，老运得安乐荣华。子女相貌敦厚，性情和顺，孝顺父母。"
'/*时正官*/
a(56) = "自己为继承人，或为养子与过房。"
'/*时比肩*/
a(57) = "为人重义气，善理财。"
'/*年劫财*/
a(58) = "资财难聚，喜欢投机性的财源，自尊心强，较重视外表穿著。"
'/*月劫财*/
a(59) = "与儿子的缘份薄。"
'/*时劫财*/
a(60) = "受祖上福荫，事业可发展，平安福禄。"
'/*年食神*/
a(61) = "受双亲之恩惠得幸福。好饮食，中年发胖。"
'/*月食神*/
a(62) = "得贤慧之内助协力成家。"
'/*时食神*/
a(63) = "得不到祖业，福份减少。"
'/*年伤官*/
a(64) = "兄弟缘薄，感情不和，较不尊重父母。"
'/*月伤官*/
a(65) = "与儿子的缘份较薄，女多子少，晚运较不幸。"
'/*时伤官*/
a(66) = "继承祖业，但比较迟。"
'/*年偏财*/
a(67) = "家中主权在父，或幼年为养子。"
'/*月偏财*/
a(68) = "祖上富贵之人。"
'/*年正财*/
a(69) = "勤俭之家，双亲富贵，得父母之福荫。"
'/*月正财*/
a(70) = "晚年幸福，子孙富贵，且孝顺。"
'/*时正财*/
a(71) = "性情刚直而不屈服之特性。"
'/*时七杀*/
a(72) = "较守不住祖业，失去家庭教育。"
'/*年偏印*/
a(73) = "适合艺术家、医生、明星、星相命理学家、理发业。"
'/*月偏印*/
a(74) = "对子女较不利。"
'/*时偏印*/
a(75) = "生在富贵之家，祖先富裕。"
'/*年正印*/
a(76) = "有智慧、有慈悲心。"
'/*月正印*/
a(77) = "得子之幸福，孝养。"
'/*时正印*/
a(181) = "亲生子女不愿继承自己的产业，多半各自朝其他方向发展。"
'/*时补运=胎*/
a(182) = "有组织领导能力,精明能干,能办大事,享有声誉."
'/*将星*/
a(183) = "擅长艺术---绘画,音乐,而且勤学苦练,做事勤恳."
'/*华盖*/
a(184) = "一生奔波,但有旅行,转移或出国的喜事."
'/*驿马*/
'经验集合
a(245) = "自小聪明过人,很有异性缘,颇受单身异性的喜欢,家门多有喜庆。"
a(247) = "聪明伶俐有心计,颇俱军事才能。"
a(249) = "博学多才,早年多灾,晚年运到自然富态。"
a(240) = "有胆识,有才略,成功多出人意料之外。"
a(259) = "此为绝处逢生之象,主晚年富足。"
a(257) = "此人聪明,但多性暴好斗,有英雄气概且常能见义勇为,为民除害。"
a(258) = "聪明伶俐,少年勤奋有功名之格。"
a(237) = "杀星有制化为权,从文主富,从武主贵。"
a(230) = "食神见印,钱财年年进,主富。"
a(235) = "食神见财,不做自来,主富足。"
a(238) = "聪明秀气,人缘好,做事易成功。"
a(270) = "聪明绝顶,少年勤奋自有功名。"
a(280) = "领悟能力强,受传统约束少,多能有所作为。"
'经验集合结束

'论生时各人命运 nlsc
'生辰适合职业 scsy，生辰注意年龄 sczy


'nlsc(1) = "性急刚强，富于勤俭，须事未定，有谋久勇，是非多端，父母得力，早年发，白手成家。"
scsy(1) = "适合的职业：艺术、政治、建筑、电气、属金水事业可忌土类。 "
sczy(1) = "应该注意年限：十一岁，十八岁，三十六岁，四十九岁，五十八岁，八十八岁。 "



'nlsc(2) = "家族缘薄，离乡成功，上官近贵，父母难常，性情暴燥，二十运开，四五兴旺，晚年齐福。"
scsy(2) = "适合的职业：商业、技师、官吏、学者、饮食、加工。忌木类。"
sczy(2) = "应该注意年限：十八岁，廿三岁，卅一岁，四六岁，七二岁。 "

'nlsc(3) = "六亲相克，少年多劳，受苦，不守祖业，流浪异乡，十七败害，四十发福，晚景逢贵。 "
scsy(3) = "适合的职业：医师、音乐家、美术家、艺人、流浪职业。忌金类。"
sczy(3) = "应该注意年限：廿六岁，廿九岁，卅三岁，卅九岁，四九岁，六六岁。 "

'nlsc(4) = "福多劳少，父母难常，兄弟难靠，出外经营，利路亨通，夫妻相克，先难后易。立身成功。"
scsy(4) = "适合的职业：机械、演戏、文学、美术、宗教家、奉职。忌火类。"
sczy(4) = "应该注意年限：十六岁，廿岁，五十五岁，七二岁。 "

'nlsc(5) = "聪明伶俐，意志强固，目中无人，衣禄光辉。"
scsy(5) = "适合的职业：实业家、政治家、中介人、教员、矿业。忌木类。 "
sczy(5) = "应该注意年限：十九岁，廿七岁，卅六岁，卅九岁，六六岁。 "

'nlsc(6) = "智能非凡，自成家业，六亲无缘，离祖成家，快乐待人。"
scsy(6) = "适合的职业：评论家、刺绣、请负业、矿业、加工业。忌水类。 "
sczy(6) = "应该注意年限：卅一岁，卅六岁，卅九岁，四七岁，四九岁，八九岁。 "

'nlsc(7) = "伶俐敏捷，不守祖业，喜好出外，自作聪明。"
scsy(7) = "适合的职业：医师、护士、政治、明星、技艺、料理店、油业。忌金类。 "
sczy(7) = "应该注意年限：六岁，十二岁，廿四岁，卅三岁，四五岁，五四岁，八五岁。 "


'nlsc(8) = "父母难为，刑克妻子，兄弟无靠，中限惊恐。"
scsy(8) = "适合的职业：土木业、电气商、建筑业、木器商、酒类商。忌水类。 "
sczy(8) = "应该注意年限：十九岁，廿六岁，五六岁，七十岁。 "


'nlsc(9) = "财来财，去宜离祖业，父母无靠，夫妻和偕。 "
scsy(9) = "适合的职业：仲买人、料理、金融界、五金商、钟表、银楼。忌木类。"
sczy(9) = "应该注意年限：十九岁，廿二岁，廿八岁，卅岁，四二岁，五四岁，七二岁。 "

'nlsc(10) = "幼时辛苦，兄弟分离，父母无缘，乏子宜养。"
scsy(10) = "适合的职业：化学、采访、教员、文艺、新兴事业、加工业。忌土类。"
sczy(10) = "应该注意年限：十九岁，廿五岁，卅二岁，四九岁，七十八岁。 "


'nlsc(11) = "富有胆力，独行奋克，福禒有进，歌乐一生。"
scsy(11) = "适合的职业：诗人、作家、投机、五金、农林、米谷商、机械。忌火类。  "
sczy(11) = "应该注意年限：十六岁，廿六岁，卅五岁，四四岁，四九岁，五七岁，七八岁。 "


'nlsc(12) = "意志坚强，沉着热心，不交际人，手艺特微。 "
scsy(12) = "适合的职业：外科医、僧侣、旅馆、支配人、艺术、古董、五金。忌火类。"
sczy(12) = "应该注意年限：十一岁，廿六岁，卅六岁，卅九岁，四九岁，五六岁，七八岁。 "

'六十甲子论命（年）这里改成（日）jz
jz(1) = "为人多学小成，有始无终，心性暴燥，幼年见灾，重拜父母保养，兄弟骨肉小靠，子多刑，男人妻大而女人夫长，可谓伶俐聪明贤能之命。"
jz(2) = "为人慷慨，喜爱春风，见事多学少成，幼灾父母重拜，九流中人，夫妻无刑，儿女不孤，六亲少靠，女人贤良，纯和之命。"
jz(3) = "为人多学少成，心性不定，口快舌硬，身闲心直，手足不停不止，利官近贵，女人贤良，晓事聪明伶俐之命。"
jz(4) = "为人手足不停，身心不闲，衣禄不少，性巧聪明，作事有头无尾，女人生性好静，一生安允有幸，男人福分之命。"
jz(5) = "为人喜气春风，出入压群众，利官近贵，骨肉刑伤儿女不孤之命，女人温良贤达，有口无心，主招好夫之命。"
jz(6) = "为人聪明伶俐，有功名之分，夫妻能和顺，作事如意，牛田有分，女人衣食不少，贤良待人，男人多出风头，善有计谋，英敏之才，福厚之命。"
jz(7) = "为人口快心直，利官近贵，衣禄丰盈，男人权柄持家，女人荣夫益子，秀气之命格，男人有带固执之性格，注重欠点，受人敬佩之命。"
jz(8) = "为人有志气，一生性宽，少年多灾，头见女吉，生男有刑，夫妻和顺，女人持家兴旺，男人有建家立业，名显荣幸之命。  "
jz(9) = "为人性巧聪明，机谋多变，和气春风，功名有分，更招贤德之妻，男人多受人敬爱，姿性英敏，女人美容艳，富贵之命。  "
jz(10) = "为人心直公平，一生口便舌能，有藏衣禄，平稳足用，六亲冷淡，须事不中，为人平等，不贪不取，晚景旺相，女人助夫兴家立业之命。"
jz(11) = "为人口快舌便，身闲心不闲，有权柄智谋，为名声远播，福禄有余，女人旺夫，生财之命。 "
jz(12) = "为人和顺，幼年多灾，父母有刑，但重拜无害，夫妇和合，谐老齐眉，存心中正，中年未岁，财谷兴旺，子女有克，见迟方好。"
jz(13) = "为人胆大，有权柄及谋略，早年平平，中年成就，晚景大好，作事按机，女人饶舌絮刮之命言多必失，守己安分，幸福遁来。"
jz(14) = "为人和睦，衣禄不少，初年有财禄常在，晚景有剩骨肉，头女无情迟好，夫妻和顺，女人旺夫，持家贤良之命。"
jz(15) = "为人猛烈，易快易冷，反目无情，早年勤俭，离祖发达，主有聪明伶俐，守其和平，以礼待人，晚年大有良机，幸福之命。  "
jz(16) = "为人风流，一生衣禄丰足，自然善闲游，嬉戏不受人所欺，六亲冷淡，骨肉难为，妻招年长，配偶和合，女人与邻睦和，亲族贤达，长寿之命。  "
jz(17) = "为人春风和气，劳碌雪霜一生，利官近贵，名利双全，衣食足用，中年平顺，晚年大兴，女人勤俭"
jz(18) = "为人有机谋，多随机应变，志气过人，衣食足用，贵人扶助，中年和顺，老运发财发福，长寿之命。  "
jz(19) = "为人勤俭，父母刑伤，灾厄可折，早年有财物但不聚，晚景旺相，事应积蓄，防备乏旱，女人兴家，贤能之命。"
jz(20) = "为人心急口快，须事伶俐，救人无恩情，反招是非多，只好休管他人事，有财无库，财来财去，重物质，造基础，女人贤德能持家，初年起晚景平安之命。"
jz(21) = "为人衣禄不少，心性温柔，出入压众，初年颠倒，晚岁利达，兴家丰隆，夫妻和合，儿女见迟，女人操持兴旺，荣隆之命。  "
jz(22) = "为人口快心直，志气轩昂，衣禄足用，福寿双全，兄弟虽有，难为得力，六亲和睦，女人兴财绵远，平稳之命。"
jz(23) = "为人豪杰和顺，招财得宝，自立家业，前运勤劳，晚年荣华，清荣高气，女人老年尚好，血财旺相之命。  "
jz(24) = "为人性巧聪明，自立自营，儿女有刑，见迟方好，生平好做善事，财源旺相，女人衣禄平稳，主有天财之命。"
jz(25) = "为人计算聪明，精通文武，早生儿女克，迟生保平安，夫妻和顺，有财益之命，晚年大兴旺相，女人贤良，发达之命。 "
jz(26) = "为人口快心直，通文艺有才能，衣禄不少，男女有再娶，花烛重明之嫌，后来夫妻和睦，百年谐老，晚年发福之命。 "
jz(27) = "为人心性急，有口无心，有事大藏，易好易怒，反复无常，衣食丰足，但财物早年耗散不聚，晚景丰隆，女人有内助旺相之命。 "
jz(28) = "为人口快心直，有志气权柄，利官近贵，身闲心不闲，六亲少靠，自立成家，少年劳碌，晚年大利，女人操持家庭兴隆之命。 "
jz(29) = "为人劳碌，手足无时停，早年难守，财来财去，不聚财宝，有虚无实，晚景发财发福，女有操持旺相之命。 "
jz(30) = "为人伶俐聪明，财谷聚散，主近贵人，中年风霜，春风之徒，守己暂暂发福，三胜时败，被人反睦，晚景荣华贤良之命。 "
jz(31) = "为人和气，喜好春风，交朋群友，利官近贵，遇凶不为凶，逢凶化吉，骨肉少靠，女人口快能言，多出风刺之命。 "
jz(32) = "为人容貌端正，少年勤俭，初年平顺，兄弟少靠，子息不孤，立家兴隆，晚年大有聚财，女人持家相夫益子之命。 "
jz(33) = "为人衣食足用，交易买卖，利路亨通，自有生财，牛田有分，早年劳碌，晚景兴旺，女人牺牲血财，持家旺相发达之命。 "
jz(34) = "为人好春风，多情重恩，利官近贵，初年劳碌，身闲心苦，晚景家道兴隆，女人清秀命，半夫半财，老岁吉昌之格。 "
jz(35) = "为人和气，自营自立，早年颠倒，财谷耗散，晚景得财，宜拜师学艺，工夫之命，立业能而成功，财源广进，利路亨通，女人养育中平之命。 "
jz(36) = "为人巧计伶俐，衣食安稳，骨肉少力，六亲冷淡，儿女早见刑克，夫妻和顺，女人清闲，晚年发达之命。 "
jz(37) = "为人尊重安稳，一生衣禄无亏，主妻贤明，持家有权柄，须事能通达，遇凶化吉，贵人提拔，女人兴旺，阴家之命。 "
jz(38) = "为人心性温和，初限有惊恐之厄，虽有衣禄财帛，进退骨肉少力，晚景福寿延长，女命喜血财旺相之命。 "
jz(39) = "为人口快心直，有事不藏，男女早婚不宜，夫妻便克，儿女见迟，初年作事颠倒，中年与晚年财帛足用，女人养畜如意发福之命。 "
jz(40) = "为人酒食不少，福禄有余，凶中化吉，早年财帛不聚，多收入多支出，晚景兴隆。女人中年与晚景多牺牲利益，保持之命。 "
jz(41) = "为人衣食丰足，一生清闲，早年平平，中末丰厚有余，骨肉相亲难靠，对自己所想事业经营，廿年间可以发达之命。 "
jz(42) = "为人气象端正，且喜好春风，只带指背煞，救人无功，做好不说，早年子女刑克 ，晚得安宁，女人有发达，相夫益子之命。 "
jz(43) = "为人清闲，初年财帛耗散，心主不忧不住，宜做手艺工夫生意，不利求名，兄弟各自成家，后运发财兴旺，女人清奇巧妙之命。 "
jz(44) = "为人喜怒不常，一生口舌能便，名利有分，衣禄皆足骨肉疏，远子息迟，见女人晚景兴旺，助夫益子之命。 "
jz(45) = "为人性急，作事反复，一生劳碌辛苦，利官见贵，儿女刑伤，财帛足用，女人贤良晓事，针黹精能，子息不孤之圈。 "
jz(46) = "为人心性聪明，衣禄有足，六亲难靠，儿女早见，作事如意，百事皆通，凡事宽量，心重为吉，女人计较多变，无灾厄之命。 "
jz(47) = "为人快活，丑年有灾，利官近贵，作事敏捷，百事如意，勤俭励业，福在晚年，犯指背煞，救人无义女人心贤能事，兴旺之命。 "
jz(48) = "为人不惹闲事，百事谋求，早年不聚财物，晚景庆良机，可谓荣华富贵之命，女人丰福，立业之命。 "
jz(49) = "为人幼年有灾，中年衣食足用，男招好妻，身闲心苦，多喜多忧，兄弟少力，六亲冷淡，凡事自作自为，女人贤能之命。 "
jz(50) = "为人衣禄不少，财帛早年不聚，一生尊重，不惹是非，父母难为，骨肉少靠，夫妻和顺，儿女见迟，早宜过房之命。 "
jz(51) = "为人诚实，一生利官近贵，家道兴宁，衣食足用，财帛多招，父母有刑，重拜双亲，女人管夫，男人怕妻，命硬三分，子息长。 "
jz(52) = "为人志气轩昂，计谋巧妙，求一生近贵，则百事如意，文武皆通，女人福寿无亏之命。 "
jz(53) = "为人聪明兼伶俐，四海春风，一生衣禄无亏，身闲心荣，好交朋友，中年事业兴隆，晚景财旺，女人贤能之命。 "
jz(54) = "为人利官近贵，个性刚强，不顺人情，兄弟居长，事业显荣，女人容貌美丽，衣食丰足，贤达起家之命。 "
jz(55) = "为人志气宽宏，一生衣禄自然，容貌端正，温良性格，少年多灾，骨肉有刑，女人姊妹少靠，大有兴旺之命。 "
jz(56) = "为人口快舌便，衣禄自来，前程显达，得贵人所钦敬，则财帛旺相，百事荣昌，强公胜祖，朋友尊重而旺相之命。 "
jz(57) = "为人一生手足不住停，名行清高，利官近贵，命犯指背煞，做好不得好，救人无功劳，女人立志能兴家，六亲冷淡，晚景兴隆之命。 "
jz(58) = "为人一生伶俐，精神清爽，口能舌便，高人敬重，财帛足用，六亲冷淡，骨肉情疏，女人贤德，操持兴家之命。 "
jz(59) = "为人一生好行善事，东来西去不住不停，且多劳多管，衣食不缺，贵人提拔，百事称意，早年平常，晚年兴旺，女人善能操持贤良之命。 "
jz(0) = "为人刚直，不顺人情，财谷如意，六亲疏远，自立权衡，则晚景胜前兴，创家之命，女人持家牲畜旺相，享福延寿之命。 "



'定义部分到此结束
dim year,month,day,time
dim ra

'if(request.form("radio1")="1") then
ra=1
'else
'ra=0
'end if





if(ra=0)  then
dim nonglidata(99)
dim monthadd(12)
dim monthadd1(12)
dim k,n,m
dim thedate
dim nongliyear
dim nonglimonth
dim ylyear,ylmonth,ylday
'测试数据
thedate=0
nongliyear=nian1
nonglimonth=yue1
nongliday=ri1
'保证系统能识别为数值
nonglimonth=nonglimonth+1
nonglimonth=nonglimonth-1
nongliyear=nongliyear+1
nongliyear=nongliyear-1
nongliday=nongliday+1
nongliday=nongliday-1
runmonth=request.form("check1")
if (runmonth="on") then
runmonth=1
end if



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


MonthAdd1(0) = 0
MonthAdd1(1) = 31
MonthAdd1(2) = 60
MonthAdd1(3) = 91
MonthAdd1(4) = 121
MonthAdd1(5) = 152
MonthAdd1(6) = 182
MonthAdd1(7) = 213
MonthAdd1(8) = 244
MonthAdd1(9) = 274
MonthAdd1(10) = 305
MonthAdd1(11) = 335

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

for m=0 to nongliyear-1922 step 1

if (nonglidata(m)<4095) then
'不是闰年
k=11
else
k=12
end if

n=k
Do
If (n < 0) Then
Exit Do
End If

'从 2进制 n-1位，开始计算,到 0位结束
'获取NongliData(m)的第n个二进制位的值
bit = NongliData(m)
For i = 1 To n Step 1
'把 数据连续除 n次
bit = bit\2
Next
bit = bit Mod 2

TheDate = TheDate +29 + bit
n = n - 1
Loop
next





n=nonglimonth
Do
If (n < 0) Then
Exit Do
End If

'从 2进制 n-1位，开始计算,到 0位结束
'获取NongliData(m)的第n个二进制位的值
bit = NongliData(nongliyear-1921)
For i = 1 To n Step 1
'把 数据连续除 n次
bit = bit\2
Next
bit = bit Mod 2

TheDate = TheDate +29 + bit
n = n - 1
Loop
rn=int(nonglidata(nongliyear-1921)/65536)

rn=rn-1
rn=rn+1

if ((nonglimonth<>rn) and runmonth=1) then
response.write "此年不是这个闰月"
exit sub
end if
if(nonglidata(nongliyear-1921)>=4095) then
if (nonglimonth>(nongliyear-1921) \65536) then
bit=nonglidata(nongliyear-1921) mod 2
TheDate = TheDate +29 + bit
end if

if ((nonglimonth=nonglidata(nongliyear-1921) \65536) and runmonth=1) then
bit=nonglidata(nongliyear-1921) mod 2
TheDate = TheDate +29 + bit
end if
end if

TheDate = TheDate+nongliday


nlzh=thedate-387
'1922元旦，把农历转换成阳历
'判断是在那一年，年份=1922+k
k=0
do
for i=1922 to (1922+(nlzh\365+1)) step 1
if((nlzh<=365 and (nlzh mod 4)<>0) or (nlzh<=366 and (nlzh mod 4)=0)) then
exit do
end if
if((i mod 4)=0) then
nlzh=nlzh-366
k=k+1
elseif((i mod 4)<>0) then
nlzh=nlzh-365
k=k+1
end if
next
Loop

'对应的阳历年份
ylyear=1922+k
k=0
m=0
if(ylyear mod 4 <>0) then
for i=0 to 11 step 1
if(nlzh>(monthadd(i mod 12))) then
k=(k+1)mod 12
end if
next
ylmonth=k
if(k=0) then
ylmonth=12
end if
nlzh=nlzh-monthadd(ylmonth-1)



end if

if((ylyear mod 4)=0) then
for i=0 to 11 step 1
if(nlzh>(monthadd1(i mod 12))) then
k=(k+1)mod 12
end if
next

ylmonth=k
if(k=0) then
ylmonth=12
end if
nlzh=nlzh-monthadd1(ylmonth-1)


end if
ylday=nlzh


'到此，完成了农历向阳历的转换


year=ylyear
month=ylmonth
day=ylday

time=ri1
sex=xingbie
end if

if (ra=1) then
year=nian1
month=yue1
day=ri1
time=hh1
if xingbie="男" then 
sex=1
else
sex=0
end if
end if

'检测输入数据是否合法

Select Case month
Case 2
If ((year Mod 4 = 0) And ((year Mod 100) <> 0) Or (year Mod 400 = 0)) And (day > 29) Then
response.write("请检查输入数据是否出错！")
Exit Sub
End If

If ((year Mod 4 <> 0) And ((year Mod 100) = 0) Or (year Mod 400 <> 0)) And (day > 28) Then
response.write("请检查输入数据是否出错！")
Exit Sub
End If
Case 4, 6, 9, 11
If (day > 30) Then
response.write("请检查输入数据是否出错！")
Exit Sub
End If
End Select


dim md
'确定节气 yz 月支  起运基数 qyjs

md=month * 100 + day
if(md>=204 and md<= 305) then
mz = 3
qyjs = ((month - 2) * 30 + day - 4) \ 3
end if

if(md>=306 and md<=404) then
mz = 4
qyjs = ((month - 3) * 30 + day - 6) \ 3
end if

if(md>=405 and md<= 504) then
mz = 5
qyjs = ((month - 4) * 30 + day - 5) \ 3
end if

if(md>=505 and md<=  605) then
mz = 6
qyjs = ((month - 5) * 30 + day - 5) \ 3
end if

if(md>=606 and md<= 706) then
mz = 7
qyjs = ((month - 6) * 30 + day - 6) \ 3
end if

if(md>=707 and md<= 807) then
mz = 8
qyjs = ((month - 7) * 30 + day - 7) \ 3
end if

if(md>=808 and md<=907) then
mz = 9
qyjs = ((month - 8) * 30 + day - 8) \ 3
end if

if(md>=908 and md<=1007) then
mz = 10
qyjs = ((month - 9) * 30 + day - 8) \ 3
end if

if(md>=1008 and md<= 1106) then
mz = 11
qyjs = ((month - 10) * 30 + day - 8) \ 3
end if

if(md>=1107 and md<=  1207) then
mz = 0
qyjs = ((month - 11) * 30 + day - 7) \ 3
end if

if(md>=1208 and md<=  1231) then
mz = 1
qyjs = (day - 8) \ 3
end if

if(md>=101 and md<= 105) then
mz = 1
qyjs = (30 + day - 4) \ 3
end if

if(md>=106 and md<=  203) then
mz = 2
qyjs = ((month - 1) * 30 + day - 6) \ 3
end if

'确定年干和年支 yg 年干 yz 年支
if(md>=204 and md<=1231) then
yg = (year - 3) Mod 10
yz = (year - 3) Mod 12
end if
if(md>=101 and md<=203 ) then
yg = (year - 4) Mod 10
yz = (year - 4) Mod 12
end if

'确定月干 mg 月干

If (mz > 2 And mz <= 11) Then

mg = (yg * 2 + mz + 8) Mod 10
Else
mg = (yg * 2 + mz) Mod 10
End If

'从公元0年到目前年份的天数 yearlast

yearlast = (year - 1) * 5 + (year - 1) \ 4 - (year - 1) \ 100 + (year - 1) \ 400
'计算某月某日与当年1月0日的时间差（以日为单位）yearday

For i = 1 To month - 1
Select Case i
Case 1, 3, 5, 7, 8, 10, 12
yearday = yearday + 31
Case 4, 6, 9, 11
yearday = yearday + 30
Case 2
If (year Mod 4 = 0) And ((year Mod 100) <> 0) Or (year Mod 400 = 0) Then
yearday = yearday + 29
Else
yearday = yearday + 28
End If
End Select

next

yearday = yearday + day

'计算日的六十甲子数 day60
day60 = (yearlast + yearday + 6015) Mod 60

'确定 日干 dg  日支  dz
dg = day60 Mod 10
dz = day60 Mod 12
'确定 时干 tg 时支 tz
tz = (time + 3) \ 2 Mod 12
'tg = (dg * 2 + tz + 8) Mod 10
If (tz = 0) Then
tg = (dg * 2 + tz) Mod 10

Else
tg = (dg * 2 + tz + 8) Mod 10
End If

'到此，已经完成把 年月日时 转换成为 生辰八字的任务


'确定各地支所纳天干
'年支纳干 yzg 月支纳干 mzg  日支纳干 dzg 时支纳干 tzg
'年支纳干 yzg

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


'月支纳干 mzg

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

'日支纳干 dzg

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

'时支纳干 tzg

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

'到此，完成各地支所纳天干的确定任务

'确定各支对应的十神

'年干十神 ygs
ygs = ((yg + 11 - dg) + ((dg + 1) Mod 2) * ((yg + 10 - dg) Mod 2) * 2) Mod 10

'月干十神 mgs
mgs = ((mg + 11 - dg) + ((dg + 1) Mod 2) * ((mg + 10 - dg) Mod 2) * 2) Mod 10

'时干十神 dgs
tgs = ((tg + 11 - dg) + ((dg + 1) Mod 2) * ((tg + 10 - dg) Mod 2) * 2) Mod 10

'年支十神 yzs
yzs = ((yzg + 11 - dg) + ((dg + 1) Mod 2) * ((yzg + 10 - dg) Mod 2) * 2) Mod 10

'月支十神 mzs
mzs = ((mzg + 11 - dg) + ((dg + 1) Mod 2) * ((mzg + 10 - dg) Mod 2) * 2) Mod 10

'日支十神 dzs
dzs = ((dzg + 11 - dg) + ((dg + 1) Mod 2) * ((dzg + 10 - dg) Mod 2) * 2) Mod 10

'时支十神 tzs
tzs = ((tzg + 11 - dg) + ((dg + 1) Mod 2) * ((tzg + 10 - dg) Mod 2) * 2) Mod 10

'到此，完成年月日时，各干支十神的确定
'确定起运数


if sex=1 then

response.write("乾造：" & a(20 + yg) & a(30 + yz) & a(20 + mg) & a(30 + mz) & a(20 + dg) & a(30 + dz)_
& a(20 + tg) & a(30 + tz) & qyjs & "岁运")
else
response.write("坤造：" & a(20 + yg) & a(30 + yz) & a(20 + mg) & a(30 + mz) & a(20 + dg) & a(30 + dz)_
& a(20 + tg) & a(30 + tz) & qyjs & "岁运")
end if


response.write("<br>")
If sexok Xor (yz Mod 2) Then
response.write("大运:  ")

for i=1 to 6
response.write("     "+a(20+(((mg+10)-i) mod 10))+a(30+(((mz+12)-i) mod 12))+"     ")
next
Else
response.write("大运:  ")

for i=1 to 6
response.write("     "+a(20+(((mg+10)+i) mod 10))+a(30+(((mz+12)+i) mod 12))+"     ")
next
end if
'神杀部分
'a(ygs);
'a(mgs);
'a(tgs)
'a(yzs);
'a(mzs);
'a(dzs);
'a(tzs)

'开始解说部分
'****************************************************


response.write("<br><br><b>性格分析:</b><br>")
call west

'*******************************************************
'中国算命方法

response.write("<br>")
response.write("<b>命造简批:"+"</b>")
'1.日坐食神己瘦妻肥 日坐食神己肥夫瘦
If sex And dzs = 3 Then
response.write("<br>"+ "自己瘦，妻子肥胖" )
end if
If sex = 0 And dzs = 3 Then
response.write("<br>"+ "自己肥胖，丈夫瘦" )
end if

'与下海有关的因素: 正偏财克正偏印；官见伤官
If (ygs = 5 Or ygs = 6 Or mgs = 5 Or mgs = 6 Or tgs = 5 Or tgs = 6) And (ygs = 9 Or ygs = 0 Or _
mgs = 9 Or mgs = 0 Or tgs = 9 Or tgs = 0) Then
response.write("<br>"+"经商下海做生意" )
end if
If (ygs = 8 Or mgs = 8 Or tgs = 8) And (ygs = 4 Or mgs = 4 Or ygs = 4) Then
response.write("<br>"+"经商下海做生意" )
end if

'近视：天干金被火克
If (yg = 7 Or yg = 8 Or mg = 7 Or mg = 8 Or tg = 7 Or tg = 8) And (yg = 3 Or yg = 4 Or _
mg = 3 Or mg = 4 Or tg = 3 Or tg = 4) Then
response.write("<br>"+"要注意预防眼睛方面的疾病，容易近视")
end if
'和算命、神秘有关的因素:日支或年支为戌，他支有亥；日支或年支为亥，他支有戌
If (yz = 11 Or dz = 11) And (yz = 0 Or mz = 0 Or dz = 0 Or tz = 0) Then
response.write("<br>"+"近神学之人")
ElseIf (yz = 0 Or dz = 0) And (yz = 11 Or mz = 11 Or dz = 11 Or tz = 11) Then
response.write("<br>"+"近神学之人")
end if

'/*事业解说*/
Select Case tgs

Case 1
response.write("<br>"+"在事业方面，可以受到兄弟的帮助。" )
Case 2
response.write("<br>"+"在事业方面，可以受到姊妹的帮助" )
Case 3, 4
response.write("<br>"+"在事业方面，可以从事思考方面的职业" )
Case 5
response.write("<br>"+"在事业方面，可以受到父亲的帮助。" )
Case 6
response.write("<br>"+"在事业方面，可以受到配偶或与配偶同辈的帮助。" )
Case 7, 8
response.write("<br>"+"在事业方面，可以受到师长、上司的帮助" )
Case 9
response.write("<br>"+"在事业方面，可以受到母亲辈的帮助。" )
Case 0
response.write("<br>"+"在事业方面，可以受到长辈贵人的帮助。" )

End Select

'/*婚姻解说--男*/
If sex Then
Select Case dzs


Case 1
response.write("<br>"+"婚姻易变，多波折。")
Case 2
response.write("<br>"+"會较为晚婚，婚姻多波折。")
Case 3
response.write("<br>"+"妻贤慧能持家，得衣食财帛，妻肥胖。")
Case 4
response.write("<br>"+"妻子美貌，但志高无才能。")
Case 5
response.write("<br>"+"妾夺妻权，但有善待丈夫。不爱妻反爱妾。")
Case 6
response.write("<br>"+"得妻扶助成功，白手成家自然富贵。")
Case 7
If ((ygs <> 3) Or (mgs <> 3) Or (tgs <> 3)) Then
response.write("<br>"+"须多注意身体。")
End If
Case 8
response.write("<br>"+"独立持家，地位安定，又得妻力，或与名门之女结婚")
response.write("<br>"+"妻妾相貌敦厚，贤淑温顺。")
Case 9
response.write("<br>"+"妻做事敏捷迅速.")
Case 0
response.write("<br>"+"妻子贤淑，能得内助之力而成功。")


End Select
End If

Select Case dz Mod 4

Case 1
If ((yz = 10) Or (mz = 10)) Then
response.write("<br>"+a(11) )
End If
If (tz = 10) Then
response.write("<br>"+ a(12))
End If
Case 2
If ((yz = 7) Or (mz = 7)) Then
response.write("<br>"+ a(11))
End If
If (tz = 7) Then
response.write("<br>"+ a(12))
End If
Case 3
If ((yz = 4) Or (mz = 4)) Then
response.write("<br>"+ a(11))
End If
If (tz = 4) Then
response.write("<br>"+ a(12) )
End If
Case 0
If ((yz = 1) Or (mz = 1)) Then
response.write("<br>"+ a(11))
End If
If (tz = 1) Then
response.write("<br>"+ a(12))
End If
End Select
'/*桃花*/

Select Case dz Mod 3

Case 1
response.write("<br>"+ a(182))
'/*将星*/
Case 2
response.write("<br>"+a(183))
'/*华盖*/
Case 0
response.write("<br>"+ a(184))
'/*驿马*/
End Select

Select Case dg

Case 1, 5, 7

If ((yz = 2) Or (yz = 8) Or (mz = 2) Or (mz = 8) Or (tz = 2) Or (tz = 8)) Then
response.write("<br>"+ a(13))
End If

Case 2, 6
If ((yz = 1) Or (yz = 9) Or (mz = 1) Or (mz = 9) Or (tz = 1) Or (tz = 9)) Then
response.write("<br>"+ a(13))
End If

Case 3, 4
If ((yz = 10) Or (yz = 12) Or (mz = 12) Or (mz = 10) Or (tz = 12) Or (tz = 10) Or (dz = 12) Or (dz = 10)) Then
response.write("<br>"+a(13))
End If
Case 9, 0
If ((yz = 4) Or (yz = 6) Or (mz = 4) Or (mz = 6) Or (tz = 4) Or (tz = 6) Or (dz = 4) Or (dz = 6)) Then
response.write("<br>"+ a(13))
End If
Case 8
If ((yz = 3) Or (yz = 7) Or (mz = 3) Or (mz = 7) Or (tz = 3) Or (tz = 7) Or (dz = 3) Or (dz = 7)) Then
response.write("<br>"+a(13))
End If
End Select


'/*天乙*/

Select Case mz

Case 4
If (day60 = 21) Then
response.write("<br>"+ a(14))
End If
Case 9
If (day60 = 59) Then
response.write("<br>"+ a(14))
End If
Case 10
If (day60 = 27) Then
response.write("<br>"+ a(14))
End If
Case 1
If (day60 = 9) Then
response.write("<br>"+ a(14))
End If
Case 7
If (day60 = 3) Then
response.write("<br>"+ a(14))
End If
End Select
'/*天月二德*/


Select Case day60

Case 1, 2, 12, 3, 13, 24, 34, 45, 56, 7, 18, 29, 50
response.write("<br>"+"不愁吃穿,衣食丰足." )
End Select
'/*天福贵人*/


Select Case dg

Case 1
If (((mz = 12) Or (mz = 0)) And ((tz = 12) Or (tz = 0))) Then
response.write("<br>"+a(18))
End If
Case 2
If ((mz = 7) And (tz = 7)) Then
response.write("<br>"+a(18))
End If

Case 3, 5
If ((mz = 3) And (tz = 3)) Then
response.write("<br>"+a(18))
End If

Case 4, 6
If ((mz = 10) And (tz = 10)) Then
response.write("<br>"+a(18))
End If

Case 7
If ((mz = 6) And (tz = 6)) Then
response.write("<br>"+a(18))
End If

Case 8
If ((mz = 1) And (tz = 1)) Then
response.write("<br>"+a(18))
End If

Case 9
If ((mz = 9) And (tz = 9)) Then
response.write("<br>"+a(18))
End If
Case 0
If ((mz = 4) And (tz = 4)) Then
response.write("<br>"+a(18))
End If

End Select
'/*学堂*/


Select Case mz

Case 3, 4, 5

If (day60 = 15) Then
response.write("<br>"+ a(19))
End If

Case 6, 7, 8

If (day60 = 31) Then
response.write("<br>"+a(19))
End If

Case 9, 10, 11
If (day60 = 45) Then
response.write("<br>"+ + a(19))
End If

Case 0, 1, 2

If (day60 = 1) Then
response.write("<br>"+ a(19))
End If
End Select


'/*天赫吉星*/

If ((ygs = 8) Or (yzs = 8)) Then
response.write("<br>"+a(51))
End If

'/*年正官*/


If ((mgs = 8) Or (mzs = 8)) Then
response.write("<br>"+a(52))
End If

'/*月正官*/

If ((tgs = 8) Or (tzs = 8)) Then
response.write("<br>"+a(55))
End If

'/*时正官*/

If ((ygs = 1) Or (yzs = 1)) Then
response.write("<br>"+a(53))
End If

' /*年比肩*/

If ((mgs = 1) Or (mzs = 1)) Then
response.write("<br>"+a(54))
End If

'/*月比肩*/

If ((tgs = 1) Or (tzs = 1)) Then
response.write("<br>"+a(56))
End If

'/*时比肩*/

If ((ygs = 2) Or (yzs = 2)) Then
response.write("<br>"+a(57))
End If

'/*年劫财*/

If ((mgs = 2) Or (mzs = 2)) Then
response.write("<br>"+a(58))
End If
'/*月劫财*/

If ((tgs = 2) Or (tzs = 2)) Then
response.write("<br>"+a(59))
End If
'/*时劫财*/

If ((ygs = 3) Or (yzs = 3)) Then
response.write("<br>"+a(60))
End If
'/*年食神*/

If ((mgs = 3) Or (mzs = 3)) Then
response.write("<br>"+a(61))
End If
' /*月食神*/

If sex = 1 And ((tgs = 3) Or (tzs = 3)) Then
response.write("<br>"+a(62))
End If
'/*时食神*/

If ((ygs = 4) Or (yzs = 4)) Then
response.write("<br>"+a(63))
End If

'/*年伤官*/

If ((mgs = 4) Or (mzs = 4)) Then
response.write("<br>"+a(64))
End If
'/*月伤官*/

If ((tgs = 4) Or (tzs = 4)) Then
response.write("<br>"+a(65))
End If
' /*时伤官*/

If ((ygs = 5) Or (yzs = 5)) Then
response.write("<br>"+a(66))
End If
'/*年偏财*/

If ((mgs = 5) Or (mzs = 5)) Then
response.write("<br>"+a(67))
End If
'/*月偏财*/

If ((ygs = 6) Or (yzs = 6)) Then
response.write("<br>"+a(68))
End If

'/*年正财*/

If ((mgs = 6) Or (mzs = 6)) Then
response.write("<br>"+a(69))
End If
'/*月正财*/

If ((tgs = 6) Or (tzs = 6)) Then
response.write("<br>"+a(70))
End If
'/*时正财*/

If ((tgs = 7) Or (tzs = 7)) Then
response.write("<br>"+a(71))
End If

'/*时七杀*/

If ((ygs = 9) Or (yzs = 9)) Then
response.write("<br>"+a(72))
End If

'/*年偏印*/

If ((mgs = 9) Or (mzs = 9)) Then
response.write("<br>"+a(73))
End If

'/*月偏印*/

If ((tgs = 9) Or (tzs = 9)) Then
response.write("<br>"+a(74))
End If

'/*时偏印*/

If ((ygs = 0) Or (yzs = 0)) Then
response.write("<br>"+a(75))
End If
' /*年正印*/

If ((mgs = 0) Or (mzs = 0)) Then
response.write("<br>"+a(76))
End If
' /*月正印*/

If ((tgs = 0) Or (tzs = 0)) Then
response.write("<br>"+a(77))
End If
'/*时正印*/


If sex = 1 Then

If ((tgs = 4) And ((mgs = 5) Or (mgs = 6) Or (tgs = 5) Or (tgs = 6))) Then


response.write("<br>"+a(245))
End If

'/*4+5(6)*/
If ((tgs = 4) And ((mgs = 7) Or (ygs = 7))) Then

response.write("<br>"+a(247))
End If

'/*4+7*/
If ((tgs = 4) And ((mgs = 9) Or (ygs = 9))) Then


response.write("<br>"+a(249))
End If

'/*4+9*/
If ((tgs = 4) And ((mgs = 0) Or (ygs = 0))) Then

response.write("<br>"+a(240))
End If
'/*4+0*/
If (((tgs = 5) Or (tgs = 6)) And ((mgs = 9) Or (ygs = 9))) Then

response.write("<br>"+a(259))
End If
'/*5(6)+9*/
If (((tgs = 5) Or (tgs = 6)) And ((mgs = 7) Or (ygs = 7))) Then

response.write("<br>"+a(257))
End If

'/*5(6)+7*/
If (((tgs = 5) Or (tgs = 6)) And ((mgs = 8) Or (ygs = 8))) Then

response.write("<br>"+a(258))
End If
'/*5(6)+8*/
If (((tgs = 5) Or (tgs = 6)) And ((mgs = 4) Or (ygs = 4))) Then

response.write("<br>"+a(245))
End If

'/*5(6)+4*/
If ((tgs = 3) And ((mgs = 7) Or (ygs = 7))) Then

response.write("<br>"+a(237))
End If
'/*3+7*/
If ((tgs = 3) And ((mgs = 0) Or (ygs = 0))) Then

response.write("<br>"+a(230))
End If
'/*3+0*/
If ((tgs = 3) And ((mgs = 5) Or (ygs = 5) Or (mgs = 6) Or (ygs = 6))) Then

response.write("<br>"+a(235))
End If
'/*3+5(6)*/
If ((tgs = 3) And ((mgs = 8) Or (ygs = 8))) Then

response.write("<br>"+a(238))
End If
'/*3+8*/
If ((tgs = 7) And ((mgs = 3) Or (ygs = 3))) Then

response.write("<br>"+a(237))
End If
'/*7+3*/
If ((tgs = 7) And ((mgs = 4) Or (ygs = 4))) Then

response.write("<br>"+a(247))
End If
'/*7+4*/
If ((tgs = 7) And ((mgs = 5) Or (ygs = 5) Or (mgs = 6) Or (ygs = 6))) Then

response.write("<br>"+a(257))
End If
'/*7+5(6)*/
If ((tgs = 7) And ((mgs = 0) Or (ygs = 0))) Then

response.write("<br>"+a(270))
End If
'/*7+0*/
If ((tgs = 8) And ((mgs = 0) Or (ygs = 0))) Then

response.write("<br>"+a(280))
End If
'/*8+0*/
If ((tgs = 8) And ((mgs = 3) Or (ygs = 3))) Then

response.write("<br>"+a(238))
End If
'/*8+3*/
If ((tgs = 8) And ((mgs = 5) Or (ygs = 5) Or (mgs = 6) Or (ygs = 6))) Then

response.write("<br>"+a(258))
End If
'/*8+5(6)*/
If ((tgs = 9) And ((mgs = 5) Or (ygs = 5) Or (mgs = 6) Or (ygs = 6))) Then

response.write("<br>"+a(259))
End If
'/*9+5(6)*/
If ((tgs = 9) And ((mgs = 4) Or (ygs = 4))) Then

response.write("<br>"+a(249))
End If
'/*9+4*/
If ((tgs = 0) And ((mgs = 4) Or (ygs = 4))) Then

response.write("<br>"+a(240))
End If
'/*0+4*/
If ((tgs = 0) And ((mgs = 3) Or (ygs = 3))) Then

response.write("<br>"+a(230))
End If

' /*0+3*/
If ((tgs = 0) And ((mgs = 7) Or (ygs = 7))) Then

response.write("<br>"+a(270))
End If
'/*0+7*/
If ((tgs = 0) And ((mgs = 8) Or (ygs = 8))) Then

response.write("<br>"+a(280))
End If
'/*0+8*/
'/*0--9代十神*/
End If

'单独从年月日时的一些论断方法
'单独从生日论命
'原来是用来论年，现在用来论日（60甲子）
response.write("<br>"+jz((60 + 6 * dg - 5 * dz) Mod 60))

'单独从时辰论命

If tz = 0 Or tz = 12 Then
scdz = 12
Else
scdz = tz
End If

response.write("<br>"+ scsy(scdz) + sczy(scdz))


'解说到此结束
end sub

bz()




'西方算命部分
Sub west()
Dim west(2000)
dim year,month,day,time
year=nian1
month=yue1
day=ri1
time=hh1
'定义部分
west(1) = "是一个有野心并具有大丈夫性格的人,不喜欢受人指使,非常顽固,坚定依照自己的意愿行事. 具有领导能力,容易获得成功,很容易得到别人的信任."
west(2) = "具有双重性格.为人温柔,敏感,常因一点小事而屈服,意志消沉.虽然有点儿任性,但人际关系 良好,善于社交.想象力丰富,有写小说的能力."
west(3) = "总是给人快乐的感觉,个性乐观开朗,善辞令.为人有点任性,外柔内刚,进取心强.不喜欢受人 束缚,自尊心强,不服输."
west(4) = "年轻时的性格较凡孤僻不合群,粗暴,但随年龄的增长,人会变得正直及认真.对看不过眼的事 与人,皆会太直斥其目非,有时颇为刚腹自用."
west(5) = "喜欢与人聊天,享受与三五知己把酒言欢的乐趣.爱好旅行及追求剌激,不能忍受单调的生活和 工作.为人机智,果断,极受朋友欢迎."
west(6) = "为人较敏感,喜欢美的东西.有一颗温柔而开朗的心,人们都爱与他接近.不过,感情颇为冲动, 容易有盲目的行为."
west(7) = "为人自我中心较强,敏感.工作能力强,性格复杂,喜欢变化及追求完美.性格有时不统一,今天 是这样的,明天又会转为另一个面孔,令人无所适从."
west(8) = "是个惯彻始终的人,一旦做了某项决定,坚持到底绝不退缩.虽然树敌者很多,但支持者也不少. 交际手腕高明,适宜从事大的交易."
west(9) = "天分奇高,才华在很年轻时已得到肯定,不过受到打击后,容易一蹶不振.具有多重性格,有时是 一个博爱的人,但有时却是一个感情吝啬鬼."
west(10) = "为人有进取心,但稍嫌性急.有竞争之心,不肯服输,能积蓄财富及有一定的社会地位. 是一个野心家,有独裁之心,意志顽强,不会永远安于现状."
west(11) = "为人温柔,善良和浪漫,是一个非常幸运的人,常有不可思议的幸运事降临在身上.口才非凡, 令人信服.不过,有点儿任性,有时难免遭受挫折."
west(12) = "是一个白手兴家的典型,不愿意依靠别人,要靠自己的力量爬上高层.脾气颇大,加上有少许骄傲, 人人敬而远之,少朋友."
west(13) = "人为谨慎,好静,做事认真,脚踏实地,是一个沉默而努力向上的人, 性情内向,内热外冷,虽有 丰富感情,但不会表达出来.重视纪律,追求完美,平稳的人生."
west(14) = "性格颇为极端,若朝好的方面则名留青史;做朝坏的方面则有犯罪的倾向.注重社交,友情浓厚,反应快,观察敏锐,可以看透对方的心思."
west(15) = "有领导能力,善于策划.对于爱的人,会毫无保留的爱她.有孤僻的一面.对财产,地位及思想颇为 执着,因此造成困扰及纠纷."
west(16) = "注重精神生活而多于物质的享受.资质聪明,做事有条理,脚踏实地,能得到崇高的社会地位. 做事我行我素,喜欢变化,但因善变而令人无所适从."
west(17) = "善于思考,理智,以自我为中心,说做事大胆,所以成功和失败,都很极端.为人善良,爱思考, 有同情心,说话有保留,会顾及别人的颜面."
west(18) = "性格爱恨分明,向往权利,重友情,但感情用事,往往拖泥带水.同时又是一个意志坚强的理想 主义者.敌人虽不少,但朋友也颇多,是一个极端主义者."
west(19) = "具有才华和独特的触觉,有以自我为中心的倾向,所以当别人对你的计划有微辞时,第一个反应 是认为对方没有知识."
west(20) = "个性倔强,想象力丰富,意志力强,能掌握自己的生活.有点韧性,有时会遭到意想不到的挫折."
west(21) = "性格开朗,活泼,有积极的人生观,喜欢变化,希望从中得到快乐.人缘好,广受欢迎.为人颇偏激, 不顾社会的惯例.有时候会情绪低落,但很快会振作起来."
west(22) = "有责任感及耐心,做事多能获得成功.有组织才能,做计划工作能发挥所长,有时会以自我为中心, 不顾别人的感受."
west(23) = "喜欢变化及旅行,头脑灵活,聪明,无论做任何事都能获得成功.缺点是没有耐性,容易感到厌倦. 不过,具有学者的智慧及幽默感,而且善于交际."
west(24) = "对工作有耐性,脚踏实地,不过有时候会受到情绪的影响而意志消沉.待人宽厚,温和,是一个有 原则的人,为坚持己见,不容易让步."
west(25) = "独立性很强,不会受人左右,也不喜欢服从他人,有时候会很顽固.有随机应变的机智,因有强烈 好奇心,所以能够体验各种生活."
west(26) = "是一个有野心和自信的人,虽然平时为人温和,但遇到不公平的事,会非常泼辣.有耐力和强烈 好奇心,愿意接受任何挑战."
west(27) = "是一个理想主义者,个性温和,平稳,但也有倔强和顽固的一面.具有创造力和感化别人的力量, 并有丰富的同情心."
west(28) = "为人独立,不服输,自尊心强,不在乎任何困难,将困难视为人生历程中的挑战.具有野心,容易 树敌,时常与别人发生摩擦."
west(29) = "有宽大的同情心,喜欢帮助别人,所以颇受欢迎.其人生有多次的起伏,成功与失败都是突然而 来的,令人措手不及,不过到最后都能够建立伟大的业绩."
west(30) = "为人亲切,善良,开朗,快乐,所以人缘颇佳.具有语言,文化艺术等多方面的才能."
west(31) = "是一个严肃的人,给人一个孤高的印象,但内心是热情的,可惜不愿表达出来.做事认真,负责, 只要能好好的与人合作,必能获得成功."
west(321) = "有野心、决断力、善理财，从事这方面工作最好。注意力非常集中，使你的好憎不致于表现得太强烈。有从事军职和艺术工作的可能。 "

west(322) = "性格独立，有很好的直觉。忌冲动，务必克服卤莽的脾气。很实际，是很好的组织者，脑力丰富，有助于成为科学家或数学家。有企业精神，实际又慈悲为怀。 "

west(323) = "有幽默感、多才多艺、善社交，因服务大众而成功的机会很大。有从事文学工作而获成功的可能。热情、高贵、有野心。避免过量的饮食。 "

west(324) = "很有魅力、身体强健、痊愈力强、性情忠诚，会给你带来许多真心朋友。热心而微带虚荣，有好运、冲动、易受感动、能爱人。 "

west(325) = "创造力强，爱好美好的艺术，对文学有兴趣，善于领会，有同情心，富正义感，娱乐大众，会带来最大的成功及幸运。 "

west(326) = " 机警，性情急躁，可能因此招来一些不幸的事。应努力达成目标，以满足你的渴望。喜欢动物，是很好的赛马训练师或是赛马场场主。 "

west(327) = " 怀有野心勃勃的理想，及对目标的真诚和热烈，一旦计划或野心受挫时，会有苛刻的倾向。只要小心分析可能的困难及结果，便可避免此倾向发展。喜户外活动，母亲般的天性会帮你放松自己。灵性的结合，具崇高理想。"

west(328) = "健谈，用辞可能过份苛薄尖酸，喜神秘、辩论、争执，能助长你储存知识、发展社会背景。对家人很亲爱，但对外人可能不体贴、好骚扰，且充满怨毒。 "

west(329) = "勇敢，侵略性强，独立，不需假借外力而实现自己的目标。切忌强迫自己从事某一种行业。 "

west(330) = "会得贵人相助而成功，有种自然而像贵族般的举止。眼光锐利，有过份注重仪表的倾向，缺乏耐心，要抑制这种倾向才好。"

west(331) = "有天赋且能干，有科学头脑，能力精力都强，能应付各种环境及场面，而赢得亲友敬重，许多野心抱负，都是凭自己的天赋而实现。遵循你的预感，因为它们常常是对的。 "

west(401) = "直觉力很强，成功全靠自己的努力，善处理细节，做需要此才能的工作较易成功。一生要遭受许多磨难，有做白日梦的倾向。"

west(402) = "想像力丰富，直觉亦佳，能在医学、哲学、心理学有关人类身人领域内得到成功，在被牺牲自己野心时，切忌贪欲和破坏。 "

west(403) = "脾气暴烈，忧郁寡欢。对环境要能适意，否则此易感的性情，会成为不好的情结。喜内省，稍受激恼，就闷闷不乐和沮丧。欲望强烈，要加以抑制才好。名誉可能来自运动方面非凡的本领。"

west(404) = "有伶俐的文学头脑，慷慨大方心地善良，有做白日梦及胆小的倾向。身体好，很性感，会带来许多人的爱慕，对投机有很锐利的察觉力。"

west(405) = "有吸引力，善游说，具雄辩能力，公正，知觉敏锐，易兴奋，极喜欢自己的家人和朋友，事业上的成功，会大于私事上的成功，有艺术倾向。"

west(406) = "喜欢小孩，有艺术气质，具可爱的天性，喜美好事物，爱好玄学及神秘的文学着作，能察知别人的思想和感觉。 "

west(407) = "易受感动，喜幻想，可发展一段非常美好的的文学生涯，会挑剔别人缺点，可能天性过于敏感所致，有不少这天出生的人，献身于科学方面的研究。 "

west(408) = "很有才智，性情高贵，善于户外体能活动，可在运动方面扬名，意志坚强，行为明确，能克服许多阻难，热情而爱人，能马上交到朋友。"

west(409) = "机智，精明，有实行的特质。天生有迷人的丰采，带来许多有利的朋友。具企业精神，成功必靠自己，受逼迫时，会有不审慎的倾向。"

west(410) = "头脑灵活，易受感动，计划达不到时，有缺乏耐心和易发火的倾向，非常爱好海洋，愿意为自己所爱的人牺牲。 "

west(411) = "诚挚而忧郁，人格高尚，慈悲为怀，造福全人类减轻世人悲苦，可能是你终生的愿望，公正的报偿，荣誉成功也是必然会获得的结果，有时个性不够稳固。 "

west(412) = "个性强，讨人喜欢，诚实，爱好音乐艺术，有自我纵容的倾向，寻求奇怪不平常的地方，导致了国外旅行的愿望。 "

west(413) = "具沈默、哲学、柔顺的天性，做事讲求方法，谨慎而温柔，有与大众交易的能力。良好的教育会带给你更大的幸福和成功，绝不接受违背正义感的任何建议。 "

west(414) = "坐立不安，有渴望的天性，感情有些急躁须练习自制才行。要坚持目标，你的天赋、直觉和锐利的先见之明，方能保证成功。旅行时格外小心。对外国很有兴趣。"

west(415) = "具灵感及创造天才，喜社交活动，组织能力强，随时准备接受任何职责，多才多艺，有艺术倾向，能发挥你的潜能，就是对自己最好的一种待遇。"

west(416) = "有大胆的想像力和敏捷的脑力，勇敢，有领导别人的决断力，此特质确保你成为成功的执行者，有时显得刚愎而任性。 "

west(417) = "好沈思冥想，微带忧郁，温和，谨慎，缺乏信心，体贴并听从所爱的人的愿望，经常承担起家庭负担，体恤别人，对目标不够坚持，喜独自居住。 "

west(418) = "理想高，直觉力强，具矛盾和喜恶作剧的倾向。能力强，事业上具先见之明。一生中，亲戚朋友常给你招来许多麻烦。小孩可带给你满足和可能的荣誉。 "

west(419) = "脑筋清楚，直觉好，仁慈亲爱，有魅力，具先见之明，身体较柔弱，不可运动过度或过份劳累，努力和经验会给你带来启迪。 "

west(420) = "具有为原则和野心所支配而难捉摸的头脑，适应力强，模仿力强异于常人。计划或野心受阻时，有走极端和恼怒的倾向。天生的领导人才，任何事都想依照自己的方式。应学得圆滑些，并使用外交手腕去做事才好。 "



west(421) = "脾气傲慢，体恤，坚毅，表现出艺术方面非凡的才华。会与显要之士接触，个性讨人喜欢，很自然的立刻交到朋友。属独裁型，要避免此倾向才好。"

west(422) = "富吸引力，果断，大方，成功要靠自己天赋和努力。忠心和天赋的才能，会帮你在任何需要零碎知识的科学行业上成功。有形成贪婪、自私的实例主义的外观，所以需要有爱、感情和了解的包围。"

west(423) = "具平衡的科学头脑，仁慈亲爱的性情，慷慨大方，有时会有点顽固任性，体质柔弱，所以身心不可过份劳累。 "

west(424) = "高贵，有艺术气质，仁慈，谨慎，温和，但缺乏信心。很体贴，能顺从别人的愿望。谨防受人利用而成为代罪羔羊。 "

west(425) = "实际而保守，要是能掌权的话，最能成功。应多用谋略，学习体谅他人。追求美好事物的过程中，表现出很大的冒险精神。成功全靠自己，天生的领导人才，有成为法官的可能。 "


west(426) = "具大胆的想像力，头脑敏锐，会有学识上的成功。有夸大的天性，对衣着很讲究，喜欢用昂贵的东西。避免惹来财物上的困窘。 "

west(427) = "善于外交，仁慈，具直觉力，表现出烹调艺术方面非凡的才能，喜美食，爱自己的伴侣尤甚于别人的，对家人亲爱，宽怀大度，有纵欲饮食的倾向。 "

west(428) = "脑筋灵敏，有知觉力，学得快。富磁性，天性可爱，导致对较好事物的追求。有自行其道，想花最少代贷得遂所愿的倾向，这很容易增加敌人。不要使自己变得自我纵容，克服此妄想的倾向吧！ "


west(429) = "顽固，机智，果断。果断并能清楚表达自己的意思是最突出的特性，但有变成固执和不顾前思後的倾向。不顾传统，固执地要表达自己极端的观点，可能使你成为疯狂而有影响力的人物。 "


west(430) = "具有带头和创始的能力，有知觉力和恶意反对的倾向。健谈和深沈的先见会帮你达到事业上的成功。一生中，时时有来自亲戚方面的财富。小心训练良好的教育，会大大地帮助你发展好的一面。"

west(501) = "个性明朗，亲爱而大方，心地善良，喜帮助别人，非常喜欢家庭生活。说服力强，从事需要推销术的生涯会成功。一生中都以牺牲敌人而赢取利益。 "

west(502) = "有艺术倾向，热情，大方，多才多艺，具社交本领，能在服务大众的事业上成功。过份关心所爱的人的利益，有太紧张的倾向。记住，为不能左右的事而耽心，是很愚蠢的。 "


west(503) = "机智，有科学头脑，很实际，有发明才智，对各方面的科学都有兴趣，能吃苦，诚实，博得朋友和伴侣的尊敬。有错交朋友的危险。赋有政治方面的才能。 "

west(504) = "具高贵的气质、艺术的倾向。计划或愿望达不到时，有恼怒缺乏耐心的倾向。有很进步的脑筋和活泼的想像力，能成为很好的执政者、作家、政治家。有很好的文学天赋。 "


west(505) = "有灵妙的头脑，非凡的智力，锐利的判断力，目标坚定，善于推销术，适合做生意。可能由于心事上的不幸，导致私生活的不快乐。 "

west(506) = "脑筋细致，有怜悯心。对文学艺术和写作有浓厚兴趣，具成功作家的气质。注重外貌，希望给人深刻的印象，这可能使你疏忽自己的家庭和关心你的人。 "

west(507) = "有创作的智力，对朋友和社团有巨大的影响力，具贵族的姿势和社交上敏锐的认识力。对外的影响远超过对内，因此可能忽视你的家庭幸福。 "

west(508) = "有怜悯心，聪明。要有适当的环境、朋友，否则天性敏感，易发展成不 好的情结。避免内省的倾向。极乐于帮助比自己不幸的人。"

west(509) = "机智，决断，坚强。自我训练可避免突然发脾气。为人公正，有知觉力。经常神经紧张，天生的领导人物。要靠自己奋斗向上。 "

west(510) = "心思难测，经常在原则和野心之间游移，适应力强，表现出非凡的复制能力和模仿力，能成为娱乐界的大艺术家。任何事都想照自己的方式去做。切记，此种不计任何牺牲的个人满足，可能引起别人的憎恶。"


west(511) = "有热望的天性，明敏的心智，很有才华，对政治很有兴趣。天赋的才能，直觉和锐利的先见，亦适于法律生涯。处理药物时要小心，以免心情不佳时，而错服药物。"

west(512) = "有冒险精神，非常喜欢旅行和探险，喜团体活动和社交活动，有组织和管理众人的能力，随时准备接受工作和责任，有决断力。 "

west(513) = "有渴望的倾向，热情而敏感，独立，能独自实现自己的目标，易受感动，反应敏捷迅速。"

west(514) = "谦虚而幽默，非常喜欢海上旅行，宽宏大量，愿意帮助比你不幸的人， 终生抱负是：减轻人们的悲哀和痛苦。这会有公正的报偿，那就是名誉和成功。 "

west(515) = "富想像力和直觉力，脑筋具有好试验的倾向，爱好神秘科学。必须要找寻了解你的同伴，以免寂寞，适于人道主义方面的追求。 "

west(516) = "有野心，果断，善理财，应朝这方面发展，集中心力于目标上，可减轻你对爱憎的表露。目标的混淆，可能造成不好的情结。 "

west(517) = "有辩才，善于游说，爱好音乐和艺术，喜欢美丽高贵而讨人喜欢的东西。婚姻美满，适于做生意，或从事专门职业，以静待时机来临。 "

west(518) = "坚家不移，坚忍，身体好，痊愈力佳，能克服一生中许多肉体上的危险。迷人而可爱，有许多真诚的朋友，好运繁多，有组织能力。"

west(519) = "意志坚强，谨慎，善与众人周旋，能赢得有权势和发达的朋友的信心，投资土地可获得优利，你蔑视违背正义的任何建议。  "


west(520) = "明敏，有知觉力。对目标的热烈，可影响别人同你合作。有文学天赋，能在这方面成功。  "




west(521) = "易受感动，理想主义，友善，仁慈。为了自己的利益，有过份大方的倾向。要学习辨别真假。愿意为共同的利益而工作，因此颇受同僚和伙伴的欢迎。喜欢作水路旅行，可能因长途海上旅行，而建立有利的商业关系和成功的婚姻。 "


west(522) = "实际，现实，具坚强信念和野心勃勃的目标。有很优良的潜能，可成为作家和雄辩家。讲话非常动听。变卖财产时，可能会有法律纠纷。 "

west(523) = "精明，机警，深受更高的同情所感动，会成为人权斗士，喜欢干涉别人的事，但动机是慈悲的，有紧急应对的能力，是个亲善大使。 "

west(524) = "有吸引力，爱幻想，要学习专心致力在一工作上，以免一次想做好几件事。口才流利，是多产作家。经营有关水及化学方面的事业，可以赚钱。 "

west(525) = "有很优越的文学才能，是政治家典型，会到远地旅行，很精良的推销商，一生中成就卓着，能赢得世人认同。 "

west(526) = "敏感，害羞，感情深邃，喜欢追究，有发展、保守或孤独的倾向。试着克服缺乏安全感的感觉，不要怕对你的朋友和伴侣，表现出精明的一面，他们会赞赏你的。 "

west(527) = "有礼貌，和蔼可亲，心地仁慈敏感，具极吸引人的个性和上进心，会有许多受人尊敬和对你有利的朋友。命中注定，将有一次伟大的恋情，富领对能力，因工作上的创新而成功。别让好运冲昏你的头，阻碍你前进。 "


west(528) = "机警，易受感动，易发火，常常变得难相处，而失去许多朋友。尽量检讨此倾向。有很好的领导力，是作家、诗人及艺术家典型。 "

west(529) = "很罗曼蒂克，理想主义者，喜作梦，说话最老实，对别人的观点也很坦白地表达，世故，喜欢与重要人物来往，社交场合中，会随着经验和个性的增长，而扶摇直上。 "

west(530) = "很迷人，富吸引力，有艺术气质，思想新，工作上能创新，理想高，会有来自婚姻和合夥事业方面的利益。有长途旅行和许多有权势的朋友。 "

west(531) = "明朗，快活，仁慈，大方，公正，适于职业生涯，很依恋家庭和父母。 "

west(601) = "意志坚强，有威仪，精力旺盛，从事大众艺术能成功。避免奢侈倾向，要懂得珍惜自己所拥有的东西。天性热情，感情用事，过多冒险和不着实际的倾向，稍稍缓和些才好。 "


west(602) = "积极，坚强，有点专横。会利用自己的观点去激励别人，但遇到困难时，会坐立不安。学习容忍和耐心，才能有更好的机会去完成自己的愿望。 "

west(603) = "脑筋敏锐，安静和平的生活态度，迷人的个性，心胸坦白，善良，使所有人对你锺爱。来自有势力的朋友和感情之联系方面之利益。有把好运和成功看得太轻，而视为理所当然的倾向。 "


west(604) = "脑筋灵巧，积极，留心。艺术气质给你带来持久而布势力的朋友。大方，仁慈，慈悲为怀，逃避责任的倾向。记住，成功幸福需要努力工作，并应用才能方能达到。 "

west(605) = "会应用能力于知识和科学方面，思想独立，个性坚强固定，有否定传统，成为大胆发明家的可能，有遗产的继承。 "

west(606) = "个性猛烈凛然，很健谈，社交场合中受欢迎，能干，井然有序的头脑配合外交谋略，会帮你赢得众人的尊敬。品性优良，容易从朋友父母得到想要的东西。会有许多艺术界的朋友。 "


west(607) = "双重人格，爱幻想，生意上的成功，似乎因家中教育而来，喜模仿动作、舞台和戏院。有坚强的意志，很好的组织和指挥能力。 "

west(608) = "具艺术家怜悯和慈悲心肠，常在大众服务方面得到认可成功。避免冲动，因为这种个性对你很不利。"

west(609) = "具直觉力，善用技巧处理职务，即使面对困难仍能坚忍果断，会带给你成功和财务上的独立。你选择朋友和同伴是有分别的，而且相当保守，会克服不合意的困难。有易感的天性和急躁的脾气。 "


west(610) = "脑筋快，喜旅行，物质上的成功，因为能结交有势利的朋友。突出的创新和幽默感，使你博得大众重视。活泼的想像力、过剩的精力，有时缺乏耐心。  "


west(611) = "羞怯，勤勉，朴素，要控制好妒的个性，爱好文学、音乐、艺术，发挥智力和商业才能，培养自信和了解别人观点的能力，感情不定，有时很难了解。文学和艺术方面的努力，会带来众人的喝采。 "


west(612) = "想像力丰富，很好的观察力，不可太相信运气，是个聪明的学生，很好的数学能力，是发明家和科学家的人才，有恻稳之心，在政府机关做事很不错。 "

west(613) = "具认真、热望的天性，强大的心智能力，很独立，喜欢帮助不幸的人，有时太宽大了，应知道措财才好。从事娱乐大众之新企业或机构，能成功。 "

west(614) = "勤勉，保守，谨慎，能替人设想，有直觉力，专心致力于目标上，亦能奋发向上。因具达成愿望的野心，勤奋而自恃，肩有野心，果决。 "

west(615) = "对自己的同胞有伟大的怜悯心，任何求助都愿意帮忙，常导致意外开销，喜受阿谀，易受利用。富灵感的构想和锐利的察觉力，会帮你克服恶运的阶段。 "

west(616) = "具沈静而忧郁的性情，富同情心，是亲友所喜欢的谘询者，有外交艺术，对向你求助的人很有怜悯心。"

west(617) = "易受感动，很有魅力的人格和上进的个性，带来许多受人尊敬而对你有利的友情。不要没考虑後果就做任何改变。除非能坚持目标，否则财务状况会不稳定。 "

west(618) = "具独断专制的人格，必须发展稳定感，以便控制好冒险的精神。有很好的构想，会在不寻常的工作上成功。能未雨绸缪的典型，许多成功的发明家、科学家，都是这天出生的。 "


west(619) = "很好的方向感，意志坚强，自行其道，能成功，好质问，渴望，有知觉力，适于职业生涯，喜研究文学和音乐，能与各类型的人相处，做一个成功的教师或公务人员。"

west(620) = "适应力强，多才多艺，有直觉力。要学习集中心志。可能会缺少与现实生活战斗的意愿。记住，要转注意力到另一件事之前，先把原先的工作做好。具迷人的个性和幽默感，吸引许多朋友，聪明地运用你的才能，始可获致成功幸福。"


west(621) = "具舒畅而温和的人格，社交场合中很受欢迎，你那艺术气质和热望的心，使你在职业生涯上成功。热心的天性和理想主义，使你参加一些很有趣而有利的工作。   "






west(622) = "具很好的个性和怜悯心，感情深邃，天性聪颖早熟，需要快乐的环境，伴侣必须是他所尊敬或景仰的人。喜户外活动，感情上的性情是有耐心的了解，而不是粗暴和专制。 "


west(623) = "严肃、热望，有很强的心智能力，诚挚而实际，很独立，帮助不幸的人，尊敬优越的人，经营新的企业和娱乐大众的大机构，可成功。"

west(624) = "害羞而坚贞，避免嫉妒，爱好文学、音乐和艺术。努力把心智训练和商业训练运用到极致。有自信了解别人的观点的能力。"

west(625) = "独到的见解，企业的头脑，和谐满足的脾气，易满足。应学习获致成功的方法。你的同伴必需是成功的人。不可好高骛远。希望和实际合乎实际才好。 "

west(626) = "有悦人的人格，很好的直觉力，极端感情用事会发展成无法控制的脾气，必需加以克制。长时独处，有忧郁和作白日梦的倾向。最重要的是避免殉道者一般的坚持。可自朋友得到许多帮助。 "


west(627) = "具柔顺快乐的性情，明敏，富想像力，为了自己的好处，可能过份大方和讨人喜欢。需要勇敢，更需要小心。喜欢模仿动作、舞台和银幕生涯。 "

west(628) = "脑筋锐利聪慧，注意力好，是他的特质。勿将生活看得过份认真。明朗快活的气氛，适当的伴侣，可除去此倾向。很有磁性的人格，独立的个性，在生意上非常成功。爱好音乐、旅行和两者相关的有利经验。"

west(629) = "很有活力，执行职务颇有技巧。紧毅、果断，即使面临困难，都能带来成功和财务上的独立，非常保守。应学习容忍、体谅别人。冲动的倾向，可能因判断不佳，导致财务上的损失。"

west(630) = "温柔、仁慈、怜悯的性情，非常好的头脑和慈悲心肠，会为人所爱。吸引力的人格，有吸引忠实朋友的能力，可能导致依赖的态度。要知道，独立自主和个人的成就，都是成功的幸福所必需的。 "


west(701) = "敏感、理想主义、友善、仁慈。为了私人利益，有太大方的倾向。要学习辨别真假。爱好水路旅行。长途海上旅行，会产生有利的商业关系。婚姻很成功。"

west(702) = "仁慈、快活。强烈的独立感和外交能力，应用于检核冒险的精神。好追根究底，领悟力甚佳，导致职业上的成功。有些人可能会去研究医药和医学。"

west(703) = "有适应环境的能力，富于才华，有直觉力，学习集中精神，有助克服困难。易厌倦，毫无目标的由一件事游移到另一年事。具迷人的人格，很好的幽默感，会吸引许多朋友。"

west(704) = "浪漫、理想主义，爱作白日梦，有时很难取悦，由于自我中心和执着的观点所致。喜欢优越和成功的伴侣。和人相处，要学习容忍和亲切的了解。 "

west(705) = "讨人喜欢彬彬有礼的脾气，脑筋机警爱沈思冥想，有直觉力，要控制一切不顾後果、只想指挥别人和领导别人的倾向。 "

west(706) = "有礼貌、讨人喜欢、善交际、个性迷人、吸引许多朋友，深得他们欢心。生活态度太认真，具敏感及感情的属性。成功的朋友会加强他的信心，帮助他克服悲观的情绪。太多的刺激和纷争，对他有不利的影响。"

west(707) = "具独创力，勇敢而专横。要控制强烈意志和自行其是不顾一切的倾向。接受各种不同教育或做生意，对你都有很大益处。商业上能成功。学习容忍和耐心，精力过剩的情形就会缓和些。家庭和个人的命运会有许多起伏。避免冲动和急进的改变。"


west(708) = "不自私、慈悲为怀、身体强健，对运动很有兴趣。避免一次做好几件事。讲话流利，是多产作家。有目标的教育和训练，是你一大资产。能控制感情便能成功。 "

west(709) = "足智多谋，富于执行能力。区别对错能助你克服许多困难。教育是克服不良特质的重要因素。应学习节制。 "

west(710) = "狡猾、聪明、实际，有直觉力。艺术的气质和明显人道主义的原则，会带来许多持久而有势力的朋友。慷慨大方，有好的仪表，和蔼可亲、慈悲为怀。逃避责任的倾向，并想以最少的代贷获致成功。 "


west(711) = "乐于行善、慈悲、有同情心，愿意帮助不幸的人，会导致意外的奢侈开销。易受感动，会因此招来麻烦。过份殷勤，有受欺的倾向。"

west(712) = "赋有聪明、科学和试验的头脑，独立守秩序的精神，以及锐利的外交谋略。此特质使你很容易得到一生中所想要的东西。 "

west(713) = "具有受人欢迎的迷人的丰采和伟大的磁性，有创造鳏明的头脑，工作和观念能创新，易着凉，避开风口和突变的气温。要抑制投机的欲望，否则一旦成为习惯，会惹来许多失望。"

west(714) = "脑筋敏锐，生活态度安静平和，迷人而虚心的态度，和人道主义的原则，使你为人锺爱。来自感情的结合和权势之友的利益。把好运和成功看得太容易，应重视自己的天赋，积极的利用，不可虚掷过一生。 "


west(715) = "具非凡的心智特质，做智识科学方面的追求，个性稳定而坚强，思想独立，有否定传统，成为怪癖的人之倾向，应抑制之。野心是可助长的，脾气大则需控制。 "

west(716) = "心智锐利、精明机警、深受较高的同情所感动，是人权斗士。喜干涉他人的倾向，但动机是慈悲的，因此会有许多误会和失望。记住，生命的最大报酬是属于那些先关心自己事情的人。 "


west(717) = "聪明、快活、仁慈、大方，无论公私都很正直。好脑力，适于职业生涯。强烈依恋家庭和家人，爱穿好，随时想保持在时尚尖端，可能因此养成奢侈的习惯，购置一些无什价值的东西。"

west(718) = "脾气冲动、有些急进、有才华、直觉好，此心智的组合，常在显赫成功之人和大众工作上得到认可的人身上发现。抑制冲动和怪癖，否则受害可能很大。 "

west(719) = "聪明、快活、多才多艺，结交并把握权势之友或对自己有利的朋友而成功。有卓越创新的新本领和特佳的幽默感，在社交界受尊崇。想像力很活泼，过剩的精力，导致耐心的缺乏。 "


west(720) = "意志坚强，具堂皇而精力充沛的人格，好直觉力，非常优越的生意眼，在生意场中进步很快，应抑制奢侈的倾向，并珍惜自己所拥有的东西。 "

west(721) = "富于才华的脑力，突出的社交能力，可获致成功或赢得大众认可。太过为自己所爱的人和朋友的利益而操心，要稍加调节才好。对父母很有亲情，易被自己的感情所蒙蔽。 "


west(722) = "聪明伶俐，快活而亲爱，非常慷慨，心地善良，乐于助人，有作白日梦和一次想做好几件事的倾向。应学习专注才好。有磁性的吸引力，会成为异性风迷的对象。   "




west(723) = " 有野心，专制，有些粗率，很独立，不需假借多少外力便能成功。在生意和财务上应避免骤下决定，能成功，并有来自上司的意外利益，有权威。"

west(724) = "有迷人的人格和亲爱的性情，喜欢音乐和艺术，欣赏美丽高贵和美好的东西，有自我纵容的倾向，喜欢受纵容。这只会增加你的困难，因为你得要面对现实，最适于艺术和职业生涯。"

west(725) = "有上进心，值得信任，有很好的表白能力，天生是领导人才，精力充沛，能控制场面，而博得朋友和同事的尊敬，有错交朋友的危险，较受男性的欢迎，要小心虚假的女性。 "


west(726) = "诚实，廉洁。强有力的意志，能抵销一些过份敏感的性情，喜欢回顾旧事，稍受激恼就变得纳闷和失望，有害羞的举止，对最高命令能做聪明的研究。多去旅行和追求自己喜受的东西，但别太浪费了。很吸引异性。 "


west(727) = "大方，有礼貌，感恩而好客，亲友和生人都很愿意帮忙你，会有成功的旅行，能获致成功的极致，具先知般的灵感，在金融界能表现相当的技巧。 "

west(728) = "大胆的想像力和敏锐的脑力，不停地追求知识和感受，亲爱而热情，透过文学和音乐方面的活动，会有利得，你所选择的事业，会给你名誉和金钱的报偿。 "

west(729) = "具豁达的胸襟，慷慨，性情豪放自由，有优越的心智和许多创新的构想，易骤下决定，有在医学或哲学方面成名的能力。 "

west(730) = "坚定而固执，对自己有信心，有直觉力，生意方面的事轻易就能做到。喜欢神秘，爱争论，增加你的知识，是智力发展的要因，并能增加安全感。对家人亲爱而豁达，但是待别人，却可能有不体谅的倾向。 "


west(731) = "眼光远大，同情心广泛，对朋友很挚爱，却可能会备 失望的滋味。要多靠直觉做事。很难和家人和睦相处。坦白而易受感动的天性，使你看来有些怪异。 "

west(801) = "热情，害羞而敏感，表现出为自己和同事赚钱的许多性向，会化烦恼为祝福。 "

west(802) = "乐观，有个需要节制的暴烈脾气，有错信别人所产生的失望，很不安静 ，经常在动，大体而言，会是幸运的，并能精通某些特定的领域。 "

west(803) = "精力充沛，勇敢，有些固执和急性，和家人有很强的连系，喜冒险，有神秘的气息，常赢得大众的醒目相看。 "

west(804) = "很优越的领悟力，有点过份忧心，能自长辈或继承方面得到利益，避免诈骗计划和投机企业。你的优越感常带来许多敌人，避免身心过劳累。"

west(805) = "喜欢学习，有艺术天份且勤学，做水路旅行时要小心，有很好的科学眼光，哲学方面会成功，是很优良的教师。肉体上能吸引异性。 "

west(806) = "能自我激励并承担重任，感情受激恼时，会冲动，坚强的意志能克服一生中的许多困难，立刻就结交到许多朋友，因你具有很好的价值感。 "

west(807) = "智力好，对知识有饥渴之情，文学、艺术或音乐方面能得到成功。豁达大度，愿帮助不幸的人，会给你带来奇妙的经验。有人道的本能，带来名誉和报偿。 "

west(808) = "有同情心，果断而坚强，小心自己的脾气。公正而有领悟力的脑子和紧张的神经系统，所以要避免心智上的过劳，而导致神经错乱。爱好户外活动。 "

west(809) = "个性温柔有力，对目标有诚意，会助你成功。有来自长亲、社交和财产方面的得利。有生意本能，具娱乐他人的潜能，喜安适轻闲的生活。 "

west(810) = "能创新，大胆而意气轩昂，坚强，果断，善于表白，言辞俱佳，有固执急躁不安的倾向。不尊敬传统，爱表现自己的急进观点，要避免强制朋友或同事，以免惹来敌视。 "


west(811) = "武断，自信，果决，热情，有时变得顽固和坚持，体质可能纤弱，要尽量放轻松点，经验会启迪你，有来自值得怀疑的朋友或同事方面的损失和麽烦。 "

west(812) = "骄傲，严肃，保守，容易发火。要培养一个较为容忍的外貌。比较悲观，要追求比较光明的地方，耐心地训练自己，养成信心，都会助你克服忧郁的倾向。不适当的刺激和太多的争吵，对神经系统有很大的不良影响。"


west(813) = "有慷慨奢侈的天性，有许多磨难和考验。信心、容忍和耐心是保持健康和财务所必需。不顺意时，会有来自异性的慰藉。赚钱能力颇佳。 "

west(814) = "勤勉而有生意头脑，很好的分析能力，但会变得自私和武断，此自我主义会导致同僚或生意伙伴的为难。要抑制任意用钱的倾向。 "

west(815) = "具有相当的先知先见，会是计划的创始人，有许多有趣和浪漫的事，有一个快乐的人生，行为急进而有生气，随时提防由于不小心所导致的意外。 "

west(816) = "创新的脑力，好的领悟力，喜作对和恶作剧。艺术和音乐天赋助你在娱乐界成功。一生中，经常有来自亲威方面的麻烦，也有来自财产、文件和旅行的麻烦。和蔼怜悯的天性和帮助亲威的愿望，会给你带来失望。"

west(817) = "聪明，好批评，有直觉力，意志坚强，脾气猛烈，对文学和艺术有强烈兴趣，稍加努力，会得到有智识的人格。要抑制好讽刺的倾向。 "

west(818) = "有冒险精神，探幽的欲望，喜欢有人陪伴，惊人的组织能力和控制别人的性向，随时准备表受职责，会得助于亲威，发展一种达到成功的实际方法。 "

west(819) = "上进心，个性很有趣，财政方面有突出的性向，应会有生意上的成功。有来自亲威方面的困难，常有意外之财，有很好的执行能力。避免匆促结婚。 "

west(820) = "很好的幽默感和分析能力，喜欢感受新的和非凡的构想，要抑制骄傲的倾向以及和比自己不幸的人交易时的专横。可能发展一种烈士般的情操，最後导致一种退隐而孤单的生活。"

west(821) = "能创新，勇敢，专横，应抑制过强的意志和专横地想自行其是的欲望。明智地利用分析能力，在生意上便很可以成功。学习和忍耐心，可制服急躁和精力过剩的脾气。有法律和税赋方面的损失。"

west(822) = "冲动而又有些急进，有聪明的脑筋和明显的直觉，此特质常在赫赫有名的成功人士身上发现。应试着抑制太冲动和不安定的举止，否则对你会有致命的伤害。很适于做政治家。   "






west(823) = "浪漫的天性，可能是理想主义者和爱作梦的人，有时很难取悦，由于固执、自我观点所致。爱与有钱人为伍，和不幸的人相处时，应要有容忍和体谅之心，不要遗忘故旧才好。 "


west(824) = "机智，富执政能力，无论做什麽都有很大的决心，是好是坏都照样做下去。对错要分明。各种不同的教育，可克服可能的不好倾向。节制和保守的原则，可以给你带来成功和幸福。"

west(825) = "敏感，易受感动，保守，有依赖的倾向。来自婚姻方面的财富，常使你生活安适多了。生意上的阻难，可能来自有恩惠于你的长者。对政治、法律、人道主义的活动，很有兴趣。为大学教育所做的任何牺牲都是值得。 "


west(826) = "认真而热望的天性，很强的心智能力，独立感强，愿意帮助不幸的人，有时为了自己的利益，会变得过份大方，要知道财务上保守主义的价值才是。会尊敬比你优越的，人应创造自己的命运才是。"

west(827) = "害羞、勤勉，有分析能力，要抑制嫉妒的癖性，喜欢文学、音乐和艺术，应尽量利用自己智力和商业上的训练，培养自信心和了解别人观点的能力。感情不稳定，有时很难了解。"

west(828) = "具非凡的心智才能，思想独立，个性坚定。有否定传统的倾向，应控制才好。要记住，成功和幸福是最可能来自经久而有价值的友谊。 "

west(829) = "心智敏锐，赋有安静和平的生活态度。有来自感情之结合和权势之友等方面的利益。会把好运视为理所当然。性欲相当强烈。 "

west(830) = "富于适应力，多才多艺，有直觉力，要能认知集中心力的价值。易疲倦，做事要有始有终才好，迷人的个性和良好的幽默感，会吸引很多朋友。一旦发展成积极的心智态度，只要稍加利用丰富的天赋，便可成功。"

west(831) = "敏感，理想主义者，为了己利，有太慷慨的倾向，要能辨真假。喜欢水上旅行和长程的海上航行。 "

west(901) = "创新的脑力，相当敏感害羞。可能缺乏克服阻难的动机，要培养自信，实际点，去追求成功。愉快的伴侣、友善的同事，会有助于你的成功和生活环境的改善。 "

west(902) = "聪明快活的个性，在生意上和社交上都很公平正直，可能从事于使你的脑子得到最大发挥的职业生涯。希望随时都穿得很好，可能形成奢侈的习性，而购买没价值的东西。要学习节制才好。"

west(903) = "机敏，实际，有直觉力。艺术气质加上人道主义的原则，会给你带来许多持久而有势力的朋友。慷慨大方，很好的外貌，慈悲为怀，有逃避责任的倾向，想花最少的精力而敷衍塞责。"

west(904) = "坚强的意志，精力充沛，专制，赋有直觉和生意眼，能使在生意界闯得成功，奢侈，要珍惜个人所拥有的才好，热情而感情用事的天性，好冒险，不着实际的狂人梦，都要稍稍节制才好。"


west(905) = "柔顺，快活，明敏，富于想像力，为了己利有太慷慨的倾向，喜模模仿别人，爱好舞台和戏剧。有散用自己才华的倾向，应学习对目标专注。 "

west(906) = "敏感，易受感动，爱追根究底，喜欢到远处旅行，有怪癖，脾气暴躁，要检讨。相当奢侈，对政治、法律和人道主义的活动，有显着的兴趣，大大的帮助你成功。"

west(907) = "头脑敏锐，精明，机警，有干涉别人事务的倾向，但动机是慈悲的，可能因此遭受误会和失望。记住，人生最大的报酬就是管自己的事。"

west(908) = "精力旺盛，善于职责之履行，即使面临困难仍能坚毅不挠和下决断，会成功，财务上也能独立，要慎选朋友和伙伴，和忍别人的观点，冲动，特别是心事方面。 "

west(909) = "明敏聪颖，很能专心，不可把人生看得太认真，明朗快活的气氛和成功之友的陪伴，都可除去这种个性。非常迷人，独立的天性，使你在生意界相当成功。 "

west(910) = "受欢迎的迷人风采，创造发明的脑子，工作上很能创新，易受寒，避免接近风口，预防气温的突变，抑制投机的冲动。 "

west(911) = "具极佳的个性和怜悯的心胸，感情深邃，能不轻易自满，便可成功。要有快乐的环境和欣赏你的才能的朋友。只要有耐心和了解才能治疗你感情上的创伤。 "

west(912) = "胸襟宽阔，个性和蔼，社交场合中颇受欢迎，有迷人的人格，很好的谈吐，喜欢旅行。感情要稳定才好。有艺术气质和吸引力，适于职业生涯，施以适当的教育，可使你的文学天才，得到大众的认可。不可太喜欢作梦，要实际点。 "


west(913) = "机智，快活，脑力甚佳，个性迷人，结交有权势和对自己有利的朋友，得到物质上的成功。突出的才能和很好的幽默感，使你在社交界取得尊崇。想像力活泼，精力丰足，有点浮躁或粗心。旅行时可能遇到危险。"

west(914) = "具有对人喜欢对个性和很好的直觉，易受感动。除非能自我驾驭，否则会发展成无法控制的脾气。长时间的独处，会有抑闷不乐和做白日梦的倾向。由于想像力太丰富所致。非常有艺术、音乐和文学的才能。避免独处。 "


west(915) = "聪明，具科学和试验的头脑，健谈，很受人喜欢，有很强的独立感，心思很有条理，个性非常吸引人。由于有这些特质，很容易自朋友那儿得到自己所要的东西。 "

west(916) = "和蔼，有怜悯心，具悦人的性情，应学习独立和外交手腕，以便运用你爱冒险的精神。好质问，热烈地渴望着，和好的领悟力，显示你在职业上可成功，能与各类型的人相处得很好，这会大大帮助你获致成功。财务上应遵守保守的原则。 "


west(917) = "乐于行善，慈悲为怀，有同情心，乐于帮助永求助之人，可能导致本意之外的浪费，喜爱受阿谀，容易为居心不良的朋友或同事所利用，可能有些自负，要提防虚情假意的朋友。 "


west(918) = "有好脑力，仁慈，能为人所喜爱，迷人的个性，能吸引忠实的朋友，可能导致依赖别人的态度，早年的家庭生活，可能鼓励你逃避职责。最好集中全力在个人的成就上，以确保成功和幸福。 "


west(919) = " 有很大的野心和敏锐的领悟力。非常吸引人的人格和上进心，坚强的个性，会为你吸引许多令人尊敬的朋友或于你有利的友谊。有时应克服过份夸张的自傲。顽皮，任性，喜欢经常的变化，这种倾向必须稍加节制才好。 "


west(920) = "有悦人而有礼貌的性情，机警，能为人设想，有直觉力，有不顾後果，想管制别人做领袖的倾向，必须小心才好，并施以适当的教育和训练。勤劳而自立，相当有野心，果决，会有贵人相助，要学习谦逊才好。 "


west(921) = "有礼貌，善意，有可与之为友的性情，个性吸引人，能带来许多朋友，而且受欢迎，比较悲观，所见皆为人生暗淡忧郁的一面，由于极端敏感和感情的天性所致。培养信心，便可以克服这种悲观的态度，不当的刺激，会导致神经系不好的反应。需要有快乐的环境和愉快的朋友。"


west(922) = "机智，快活，大方，社交上，财务上，都很公平正直。适于发挥自己心智能力的职业生涯，非常依恋家人和家庭，对衣着很敏感，希望走在时尚的尖端，有时这种小小的虚荣心，会引致不必要的花费。   "






west(923) = "具强大而进步的精神，有古道热肠和仁心，在医学和慈善事业方面有卓越的才能，经营财产和和保险事业亦能成功，可能有爱情方面的失意。 "

west(924) = "心地善良，有艺术气质，缺乏信心，仁慈怜悯的心性，使你成为公民自由之斗士，愿意帮助不幸的人，直觉和灵感配合得很好。 "

west(925) = "具灵活而好质问的脑子，公正，怜悯，平易近人，善于社交，懒散，喜欢奢侈，喜欢自己的家人和亲戚，有显着的文学才能，悦人的人格，即刻就能交到朋友，签署文件时要特别小心。 "


west(926) = "具可爱的脾气，友谊和爱情都很坚定，一定能完成自己的职责，在医学，哲学，心理学和人道主义方面可获致美名。 "

west(927) = "具好评论的脑子，购买值钱的东西时，要谨防受骗。太关心别人的利益，可能会忽略自身的利益，不要让太敏感的天性，使你对亲近的伴侣过份挑剔。 "

west(928) = "理想主义者，浪漫，有自我牺牲的倾向，非常挚爱所爱的人，在社交界非常成功，很适于艺术或职业生涯，不宣于生意方面的追求。 "

west(929) = "脑筋很快，想像力丰富，为了自身的利益，可能就会变得太冲动，要抑制纵食纵饮的倾向。散用自己的才华，一次做太多事。 "

west(930) = "机智，能干，实际，有勇气，有点宿命论，是天生的领导人，具有控制场面的能力，很博得朋友和同事的尊敬，有误信他人的危险，和年长的亲戚同做生意，可以确保成功。 "


west(1001) = "足智多谋，仁慈，骄傲，非常喜欢好的衣着，希望经常盛装打扮，具贵族般的自然举止，有过份注重外表的倾向，会有爱情和生意上的不忠所致的损失。 "

west(1002) = "能干，和蔼，很有适应力，有不寻常的复制力和模仿的性向，非常喜欢小孩，希望自己的艺术才能可以得到广泛的认同。 "

west(1003) = "温和而害羞，怜悯，了解人，是人性的一个杰出的法官，也是忠诚的友人，要特注意身体健康，可能有继承方面的利益，谨防别人的诈财。 "

west(1004) = "头脑很进步，谨乃错用精力，有爱情和婚姻方面的困难，要抑制投机的嗜好。良好的智力和对人生问题的了解，会使你成为卓越的法官。 "

west(1005) = "具富于感受性的心灵，机智，自我依赖，会由生人那儿得到许多利益，爱好神秘，喜欢争辩，这会大大地增见识，成为发展心智能力和社会背景的主要因素，非常亲爱，乐意对自己的家人宽大。 "


west(1006) = "富冒险精神，喜欢到陌生的地方旅行或访问，喜欢社交生活，很容易受感动，可能会由于受欺骗而遭受痛苦，多才多艺，脑筋敏捷。对目标要专注方能成功。 "

west(1007) = "有野心，富于冒险进取的精神，头脑清楚，要小心选择商业上的伴侣，由于想利用别人，而有某些危险的存在，你母亲对你的影响很大，是很卓越的演说者，从政很不错，要抑制投机的冲动。 "


west(1008) = "精力充沛，在推广大企业的关合上有潜能，富于科学才能，要避免在生意或私事方面做不必要的冒险。 "

west(1009) = "能干，勇敢，但对朋友和同事可能过份要求，通常可以不顾反对，得到自己所追求的，可能由于仆人或同夥人而招惹麻烦。要减少缺乏正确的理由，而做错误的改变之倾向。 "


west(1010) = "有些自我中心，顽固，夸大，要避免为了给某些人有深刻的印象，而大事花费的强烈倾向，签署重要文件和契约时，要格外小心，有因疏忽所致的财务损失。 "

west(1011) = "思想清晰，脑子开通，愿意努力达成自己的目标。多谋略，坚忍，有勇气，会使你克服大部分的困难，有继承方面的利益，对异性有吸引力。 "

west(1012) = "个性坚强，聪明，前进，有非凡待构想，在很多方面都很幸运，牵涉到重要文件时，要格外小心，可能有因疏忽所致的损失。 "

west(1013) = "思想清晰，果决，仁慈，亲爱，有迷人的人格，有时会有些顽固和任性，体质孱弱，不可身心过劳，有杰出的艺术和音乐才能，许多有利的，持久的友谊，对你的获致成功和幸福，都很有帮助。 "


west(1014) = "冲动，机警，大方，会利用机会，自我纵容和太多的爱情纠纷，会影响健康，会做远地旅行，足迹遍及各处，有来自好的直觉方面的好运气。 "

west(1015) = "勤勉，有野心，锐利的领悟力，有点专图私利，生意上易受骗，发展音乐和科学才能，可带来金钱和私人方面的报酬，有由于爱情及感情纠纷所致的健康上的危险。 "

west(1016) = "具高贵的艺术气质，脑筋活泼而进步，要经常小心自己的健康，随时会有家事和生意上的麻烦，会遭受许多考验及磨难，特别是关于小孩的方面。 "

west(1017) = "有令人激赏的音乐及艺术天才，假使能控制对享乐的沈溺，必可获致此方面的成功，追求人生美好的事物时，表现出很大的冒险精神。成功要靠自己努力，防身心过劳，以避免神经方面的疾病。 "


west(1018) = " 非常聪明，具极佳的个性，富才华，有很显着的社交本领，在大众事业方面成功和得到认可，有太过为自己所爱的人担心的倾向，对父母很有亲情。 "

west(1019) = "有恻隐之心，本性善良，可能会被人利用，财务上很幸运。但有时太贪图快乐，非常喜欢动物可能包括马在内，甚至有以之为赌注的倾向，切记：桃花运是虚假的。 "

west(1020) = "冒险而神秘的天性，科学头脑，非常喜欢海和旅行，宽大，乐于帮助人，会使你一生中体验许多奇妙的经验，科学的天性可用于增益人类的福利，并减轻世间的悲苦，而给自己带来成功和幸福。 "


west(1021) = "坦白，直爽，爱好运动和旅行，有迷人的个性，很受异性之赞赏，别被自己幸运的特质愚弄了，或因此而骄傲，要控制贪图安逸的倾向。 "

west(1022) = "勇敢，果决，有野心，希望能得到大众的喝彩，心智能力发展得很好，处理法律方面的事要特别小心，可能因书写或签文件而招惹麻烦，非常容易受感动，反应也很快，要慎选朋友和终身伴侣。   "





west(1023) = "果决，善变，能同大众做生意，并透过有势力或发达的朋友的联合而成功，适当的教育，能使其准备妥当，过一种为大众受务的生活，不肯接受违反其固有之正义感和正道不合之勾当。必须学会相信自己的判断。 "


west(1024) = "意志力强烈，具崇高而精力充沛的人格。自动发起和很好的生意眼，在生意界窜得很快，要抑制奢侈的倾向，了解并珍惜自己私人的所有物的价值，热情，太感情用事，好冒险和狂妄的梦想等倾向，均应节制。"

west(1025) = "聪明，悟力高，加上对目标之强烈和真诚，会倾向于过份苛求，和不能忍受自己的野心或计划受干涉的可能，这能藉小心分析可能之反应及後果而避免，要多做体能运动，以免太用功，而使脑子的负担太大。 "


west(1026) = "害羞，勤学，坚贞，有妒羡的癖性，且随时都可能表现得很明显，应稍加节制。爱好文学和艺术，多利用自己的心智和商业训练，培养信心和了解别人观点的能力，感情不稳定，有时很难了解。"

west(1027) = "赋聪慧，科学，试验的脑子，很卓越的健谈人物，很受欢迎，表现出独立而井然有序精神，及敏锐的外交能力，这些好的特质，使你很容易自朋友那儿，得到想要的东西。可能变得自我中心，而且很容易被惯坏。 "


west(1028) = "极佳的个性，仁慈，感情深邃，本能的就很有智力，可能表现出太保守太严肃的倾向，必须要有快乐的环境和可资敬仰的人相处才好，可多做户外活动和运动，须培养耐心和了解。"

west(1029) = "聪明，好批评，富同情心，意志强，脾气猛烈，非常爱好文学和艺术，写作能赋予他心智特质的人格，应发展天性中实际的一面，以获致终身的成功和幸福。 "

west(1030) = "不自私而慈悲的性情，身体健康，对运动和技艺有锐利的兴趣，要能专注在一个目标之上，口才流利，可能成为多产作家，表现出一种激情的天性。 "

west(1031) = "性情悦人，机警，喜欢思考，有直觉力，要抑制那种不顾後果的领袖欲，勤勉，自立，很有野心，果决，由于异性会招惹一些困难。 "

west(1101) = "具非凡心智特质，能做知识和科学方面的追求，思想独立，稳固而有力的性情，否认传统，成为一个勇敢的创新者的倾向，避免自负，刚愎的欲望。 "

west(1102) = "精明，机警，可能成为人权斗士，有干涉别人事物的倾向，用意本佳，却可能因此而失望。"

west(1103) = "冲动，急进。聪慧和很好的直觉力，这种心智特质，常可在赫赫有名的成功人士身上发现。要控制冲动和急进的倾向，否则可能是你的致命伤。 "

west(1104) = "敏感，易受感动，好追根究底，喜欢旅行，会出门做长程旅行，有时变得很难驾驭，表现出急躁的脾气。应时常自我检讨。有奢侈的倾向，宜加抑制。 "

west(1105) = "敏感，易受感动，能集中心神，避免把人生看得太认真，明朗愉快之伴侣和气氛，很有助于除去这种倾向，极具吸引力之人格和独立之天性，在生意上大获成功。特别是他愿意克服不好的倾向的话。"

west(1106) = "脑筋很锐利，对人生采安静和平的态度，迷人而坦白待风度，加上人道主义之原则，使这些幸运者为人所锺爱。能透过有势力之友的联结而发财，但有将这种好运看得太轻易的倾向，应学习欣赏自己的天赋，并积极利用才是。 "


west(1107) = "机敏，实际，有直觉力。艺术气质和人道主义的原则，带来许多持久而有势力的朋友，大方，和蔼可亲，有好的容颜，想逃避责任，想轻易就得所愿，切记，聪明地应用你的天赋，是成功所必须。"

west(1108) = "慈善，有恻稳之心，愿意帮助前来求助的人，导致意外的大笔开销，喜欢受阿谀，很容易为人所蓄意利用，星相显示有受欺的倾向。"

west(1109) = "创新，大胆，喜欢权势，要随时抑制过份强烈的意志和独断独行的倾向，接受各种不同教育和从事氐意方面的工作，有很大的益处。学习容忍和耐心，就可减轻浮躁的性情。 "


west(1110) = "个性悦人，直觉力极佳，除非能经常控制过份感情用事的个性，否则会发展成无法控制的脾气，长时的孤独，会有做白日梦和过份抑郁的倾向，应多和性情相近的人接触，最重要的是不可过份偏激。"

west(1111) = "精力充沛，做事有技巧，坚毅，果断，即使面对困难也不改变，会成功，财力能独立，选择朋友和同事有歧视待遇，并经常相当保留，抑制冲动的倾向。 "

west(1112) = "敏感，理想主义，友善，心地善良，为一己之利，有太慷慨的倾向，要明辨真假。切记，施舍是美德，但应给予值得给予的人。喜欢水路旅行，会有长程的海上旅行，能发生有利的商业联系。"

west(1113) = "野心很大，反应快，悟力强，有吸引人的人格，带来许多受人尊敬和有势力的朋友，要抑制自视过高的夸大，经常想做改变，资产会给你带来成功。 "

west(1114) = "认真而渴望的天性，聪慧能干，深沈而实际，很独立，乐意帮助不幸的人，有时为了自己的好处，可能太宽怀大度了，财务上保守些较好，尊敬上司，能创造自己的命运。 "


west(1115) = "胸怀宽大，个性和蔼，迷人，健谈，社交场合中深受欢迎，有艺术气和吸引人的人格，比较适于职业生涯，不宜经营生意。 "

west(1116) = "有浪漫的天性，会成功理想主义者和寻梦人，有时很难昨取悦，因为太自我中心了。显然是喜欢和比自己发达的人为伍，这种倾向必须修正，以免一时误信别人，而遭受财务上莫大的损失。"

west(1117) = "适应力强，富才华和直觉，应了解集中目标的价值，星相显示，很容易厌倦，由一件事到另一件事，毫无目标地漂流，个性迷人，富幽默感，能吸引许多朋友。 "

west(1118) = "机智，精力旺盛，谨慎，真诚，有直觉力，成功是可预见的，特别是在法律，新闻和生意方面，须辨清对错，以免遭受严重损失。有不顾风险，想达到成功的显着果断，不论是好是坏都坚持下去。 "


west(1119) = "能创新，有冒险进取的精神，性情和谐知足，可能很容易满足而丧失克服阻难之动机，要培养自信心和获取成功的实际方法，应结交快活的同伴和成功的伴侣。"

west(1120) = "赋有受人欢迎的魅力和伟大的吸引力，有创造发明的头脑，观念上和工作上均能创新。容易伤风，避免接近风口和气温的突变。对投机的固有欲望应加抑制，显示出有科学方面的成就。"

west(1121) = "聪明伶俐，快活，亲爱，慷慨大方，心地仁慈，愿意帮助别人，有大做白日梦的倾向，并一次想做好多事，这样地散用精力，会成为万事通，一事无成，要能集中目标，有始有终。有迷人的天性，会深受异性的风迷。 "




west(1122) = " 直觉及未卜先知的天性，赋锐利的观察力，经常追求知识，会大大地帮助你成功。生意和科学生涯都会很幸运，唯独爱情方面多挫折。哲学及研究亦有异于常人的天才。"

west(1123) = "实际的梦想家，能将构想付诸实现。与人合夥比独自经营更会成功。会经常遭受阻难，但终必能克服之。 "

west(1124) = "心胸宽大，本性善良。许多这天山生的人都与远亲结婚。要抑制敏感不定的脾气。表现出显着的艺术，音乐及心灵方面的才能，能在这些领域中得到成功。"

west(1125) = "亲爱直爽，非常喜欢自己的家庭和亲戚。电子方面能成功。在水上要小心。随时要注意身体健康。朋友大多不如表面所表现的那样。 "

west(1126) = "丰富的创造天才，艺术和灵感的先知。具克服一切阻难的能力及决心。生活不会很安适，但终究会不曾。圆滑地与长者相处，要积极审慎才行。 "

west(1127) = "有游说才华，能影响他人，运气时好时坏。喜欢运动及旅行。性欲很强。"

west(1128) = "能与敌为友，由于有悦人而和谐的天性的缘故。有尊严，很适于执政的职位。进步较慢，但一定会有。年长的亲戚会增益你的好运。 "

west(1129) = "敏锐的内察力和大的毅力会使你克服生意上的困难和财务上的阻碍。有来自亲戚的麻烦。极爱好享乐和粗鲁的态度必须克制才行。"

west(1130) = "具有相当棘手的人格，虽能集中精神，却又有点顽固和愚勇。会有许多考验和磨难，可过似乎有贵人在冥冥之中相助。成功和幸福主要靠你的教育及早年的教养。"

west(1201) = "强烈的野心，很好的执政能力，非常爱好真理及正义。凡事得来不易，但大致来说，将会幸运和成功。常在一生的某个时候有遗产的继承。偶然的厄运，足以激励你的决心。"

west(1202) = "思想深沈，侵略性的独断，大方，心地好。小心投机方面可能发生的错误。不智的投资会遭受突然的损失。要慎防意外事故，一生中会有许多事故发生。"

west(1203) = "敏感有点浮躁，坦白，谨慎而可靠，不管是有朋友或贵人相助，成功还是不很容易。精力非常充沛，心力两佳。会过着完满的生活。"

west(1204) = "会过着活泼的生活。喜欢从事大胆的企业投资和冒险会有不同的得失，但很幸运，有来自不可知的地方的帮助，一个幸运的预兆就会带来相当的财富。 "

west(1205) = "聪明睿智，能干，效率高，肯干，有些神秘的事情会发展祚有力的影响。耽情主义可能会引致沮丧的心境。要抑制反判权威的倾向。婚姻美满。"

west(1206) = "脑筋清楚，成功多靠朋友，星相显示有由于不幸的恋情所带来的心碎。 勤善良，须经常提防不小心的同事。 "

west(1207) = "独立的精神，非常爱好行动的自由。极容易获得许多知识。决断，浮躁，长久留在相同的环境会变得相当沮丧。由于兄长或姻亲会惹来一些偶然的厄运。"

west(1208) = "具哲学，批判，有适应力而灵敏的脑筋，是个能干的记者。好冒险，但很小心，非常适于政治生涯。要尽量避免伤风。这是一个幸运的日子。"

west(1209) = "灵巧而令人佳许的天性，有高尚的嗜好。有许多异性朋友。有文学和艺术的脑筋，能在文学界争得一席之地。秘密行为会有得失。"

west(1210) = "侵略性，有野心，心智能力及体力都很充沛。保守而节俭，但牵涉到异性时例外；慎防法律纠纷及过食。试着抑制你的奢侈和投机的倾向。"

west(1211) = "格外地幸运，受欢迎，心地好而真诚，生意方面的事有很杰出的先见之明，会受上司及居高位者所恩宠。别因成功而歪曲了你良好的判断力，特别是牵涉到异性时更是。"

west(1212) = "具许多优点和突出的才能。会是很好的辩论家，在律师和政治家的行业上都会很优越。要避免招致别人的嫉妒，牵涉到异性时须抑制不定的脾气。"

west(1213) = "有力的个性，亲爱的脾气，还有运气多变，会完全影响或改变你的一生。交友恋爱都不如意。有显着的社会主义者倾向。"

west(1214) = "富同情心，有灵气，伶巧，多才多艺，活泼。一生中有某些时候会有意外或其他不幸的重大危险，要特别小心。不可做不必要的投机，不只是旅行时如此，所有财务方面也是。"

west(1215) = "有无尽的活力，天性好， 勤而可靠。非常浪漫，但很忠心。舞台、音乐、文学或艺术都是很适合而能厚利的生涯。异性朋友中可能有许多柏拉图式的精神恋爱。"

west(1216) = "高贵，怜悯而好心。明朗快乐，热爱生命。有点诗人的气质，易受感动，渴望肉体上的快乐。好冒险会带来令你关心的重大责任。要抑制求变化的倾向。"

west(1217) = "个性强，贪慕名位。在文学领域非常多才多艺，想像力丰富，极可能成为作家。很能干，易遭出卖及嫉妒。要多与家人连系以防交错朋友所致之错觉。"

west(1218) = "多情而敏感。喜欢家庭，但还是会有许多旅行。易受感动，所以很容易受利用。要明智地选择你的朋友及同事。没明了对方的动机前，别轻易表露自己的感情。"

west(1219) = "一生多变化，勇气及信心可以抵消突然的意外变化。慈悲为怀，仁慈，亲爱，容易受利用。有执政能力，有一些生意或职业上的成功。"

west(1220) = "本性善良，对生命很乐观。靠近火旁，机械和爆炸品要随时小心。生意方面很有冒险的企业精神，适于政府或政治生涯。慎防生意上或爱情上受欺骗。与陌生人见面或接触要谨慎。"

west(1221) = "对无能为力的事，有忧心或小题大作的倾向。很好的人智能力，但会因感情的纠缠而导致厄运。消化力衰弱，很容易影响健康。在面临多事之秋时，多用你的直觉，不要犹豫去追求良友的忠告。   "





west(1222) = " 机智、能干、实际，有商业头脑，对生意有关的各方面都很有兴趣，天生的领导人才，能博得朋友和同事的尊敬。克服好争吵搏斗的精神，便可减少许多困难。若使不良的倾向发展成习惯，在你的社交活动和生意上，都会变得诡计多端。 "


west(1223) = "脑筋深沈，交替在原则和野心之间，非常囝决，野心勃勃、适应力强、再制力和模仿力甚佳，计划或野心受干涉时，会变得相当暴躁和激动，天生的领导人才，凡事都想照自己的方式去做。 "


west(1224) = "具野心及果决的人格，财务方面表现出明显的性向，集中目标能使你获益良多，有发展不良情结的倾向。 "

west(1225) = "好冒险、喜欢旅行、社交活动，有组织能力和控制别人的性向，要避免一次做好多件事情。"

west(1226) = "圆滑、有恻稳之心、骄傲、喜欢好看的衣着，希望随时打扮得整整齐齐，很注重外表，为了讨好外人，常和家人处得不好。"

west(1227) = "明朗、快活、亲爱、非常慷慨、乐于助人，有做白日梦的倾向。做事应该专注，很性感，对异性有吸引力。"

west(1228) = "好冒险的而神秘的天性，科学头脑，爱好旅行，喜欢海洋，乐于帮助别 人，会导致一生中许多奇妙的经验，终身之抱负是为增益全人类和减轻人们的痛苦而研究。"

west(1229) = "敏感，计划和欲望不能立即完成时，有激恼和缺乏耐心的倾向，具进步 的脑筋和活泼的想像力，可能有逃避责任的倾向。 "

west(1230) = "精力充沛，个性强而有力，富有耐心，有认真的头脑。富才华的脑子，显着的社交本能，可以成功而受到大众认可。有太为自己所爱的人耽心的倾向，要稍加调和，以免引起误会和不幸。很爱自己的母亲。 "

west(1231) = "具进步、果断和慷慨大方的天性。成功靠自己努力。巧于细节，做这方 面的事业或职业会非常成功。必须要有爱和了解的滋润，否则会形成自私和现实的个性。"

west(101) = "脑筋清楚、果断、亲爱、具迷人的人格，有时有点暴躁和顽固，体质纤弱，不可身心过劳，经验会带来新的启迪，一生颇多波折，容易受骗， 应该努力学习人生较深的一面。"

west(102) = "性温和而可亲，有悟力，学得快，具艺术气质和可爱的个性，喜欢美好 的事物，有好逸恶劳的倾向，不可以过份纵容自己。"

west(103) = "个性吸引人，身体强健，复原能力甚佳，迷人而动人心神的脾气，会带 来许多真诚的友谊，有误信别人的倾向，也会有财务上的损失。 "

west(104) = "具带头、创新和领悟的能力。多多少少带点恶作剧般的作对。好的事业 才能和商业先见，有助于达到事业上的成功。仁慈怜悯，乐善好施，但可能由于对方的不知领受而失望。爱情不是挺顺利。"

west(105) = "有直觉力、富艺术气质、敏感、实际、独立，能实现自己的目标而无须 假借外力，不可强迫自己做那一行业，应让你的直觉去做选择。 "

west(106) = "个性果断富弹性、做事讲求方法、脾气温和、谨慎、善与众人相处，能 得有钱有势之友的资助而成功，反对不平和不公正的事，不可以轻信自己骤发的判断，应该等到自己已经得到知识和实际经验之後，再下判断 。 "

west(107) = "具高贵的艺术气质和慈悲为怀的天性，温和，谨慎，敏感，可能会缺乏 耐心，具有迷人亲爱的性情，要学习专注，加强自己的信心，对家人 贴而顺服，不应承担家庭重任，因为到最後，可能会过着退隐而孤单的 生活。"

west(108) = "脑筋深沈而保守，好运和迷人的丰 会带来许多友谊，应学习体谅他人 ，言行要圆滑些才好。追求美好的事物时，有上进心，成功全靠自己努力，不可身心过劳。 "

west(109) = "机敏、聪慧、判断力锐利、个性和蔼可亲讨人喜欢，具社交上的优雅和迷人。应抑制猛烈的脾气和不屈的个性。对目标执着而有执几技巧，会成功。不可以牺牲家庭生活去换取社交生活。"

west(110) = "富大胆想像力、脑筋敏锐、经常在追求知识和改进，很有助于事业上的 成功，具奢侈天性，喜爱穿好衣着，爱交浪费的朋友，应该要节制些才好。避免投机的冲动，以免惹来困窘的局面。耐心、容忍和体贴会助长 敏感的天性。 "

west(111) = "聪明好批评、有同情心、脾气固执，对文学和艺术有强烈的兴趣，可能成为一个有智慧的人格。天性敏感，需要有公正和同情的气氛才行。应 发展天性中实际的一面，以确保成功和幸福。 "

west(112) = "伶俐而有悟力，加上对目标的真诚和热烈，使你在计划或野心受干涉时 ，变得好苛求和讥讽，要多做体能运动，以减轻因过份用力而加重的精神负担，不可过份小心。多疑而好多揭发别人的弱点。"

west(113) = "富想像力、精力旺盛、有侵略性、有试验和科学头脑，对神秘科学有兴 趣。在人道追求上富享盛名，自己选择职业可更成功，在家事方面可能会感到悲哀和失望。 "

west(114) = "赋吸引人的人格和悦人而亲爱的性情，发好音乐、艺术，并享受生命中 美丽高贵的事物，有自我纵容的倾向，不要宠坏自己，否则只有增加困难罢了，应拒绝所有不必要的纵容，保存自己的精力和资源。 "

west(115) = "具灵妙的天性和艺术气质，感情反应很快，要能自制才行。天生的才能 ，直觉和锐利之先见，都暗示着生意生涯的训练，可以确保你的成功。旅行时必须谨慎。"

west(116) = "具艺术、灵敏聪慧的脑子，除非有适当的环境，否则敏感的心性会发展 成不好的情结，喜欢回顾，稍受激惹就会抑郁不乐和失望。执政能力和艺术天才并用之生涯，会达到最伟大的成功。"

west(117) = "果断、坚强，要学习自律，以克服凶恶的脾气。思想平稳、公正、有悟 力、敏感，要避免太伤神而遭受不良後果。勤勉、有野心，当进步滞碍停顿时，可能会郁闷不乐和倔强。"

west(118) = "讨人喜欢、深思、勤勉、害羞，在户外活动和体能活动方面，有很好的 技巧，可能扬名于运动场上，意志坚强，能克服生命中许多困难。 "

west(119) = "彬彬有礼、讨人喜欢、长于交际、个性迷人，会吸引许多朋友。比较悲 观，所见多为人生黯淡之面，应培养自信心，以克服这种悲观的态度，不当的刺激和争吵，对神经系统会有不良影响。   "





west(120) = " 本性善良、 勤、长于外交、天生喜欢海洋，航海方面会成功。有来自亲戚、遗嘱和文件方面的一些麻烦，敏感而艺术的天性，会吸引许多不平常的友谊。 "


west(121) = "很重感情，会从给你阻难的人那儿得到利益。亲友的欺骗会阻碍你的进步，不过有料想不到的财务上的成功。 "


west(122) = "理想主义者、长于外交、富于才华，结识许多重要的朋友。有稍受激恼便生气的倾向，这样很容易发生意外事故和突然的厄运，尤其在婚姻、法律和财务方面。好学不倦。 "

west(123) = "一生中会有许多有势力的朋友，具创造发明的脑筋，从事科学方面的努力会成功，有些并能纯粹依靠自己的才能，而获致名誉和财富。非常喜欢旅行，会由一个地方玩到另一个地方。 "


west(124) = "会遭受许多悲哀、考验和禧难，必需要有许多的爱，助你战胜这些战役。心境奇怪，有些怪癖，思想行动互相矛盾，千万不要用金钱去买友谊，要以美德佳行去赢取他们的友情才是。 "


west(125) = "喜欢写作、阅读、变化和新奇，又因对新领域的兴趣和努力，是会成功的，而且因为这样，你的成功可能是狂极为现代化的活动方面。有法律方面的纠纷，有来自长者的帮助。"


west(126) = "赚钱容易、花费也大，并且会把钱花在享乐上。有工作过度或过份用功而伤害健康的可能，有来自异性方面的麻烦，使你花费相当的金钱，必须小心虚假的诺言和秘密的引诱。"


west(127) = "庄严、有野心、有直觉力、能宽恕别人的错，自己寻求享乐，并愿意给予他人许多欢乐，可能会因此受骗，不过一生金银丰足。随时都会遭受肿胀的疾病。婚姻有点不太乐观，因此结婚前应先清楚，别被外貌所愚弄了。 "


west(128) = "坚强的个性、富于才华，很容易因身心过荣而损害健康。聪明伶俐、喜欢辩论和政治，在市政工作方面能成功。会是优越的法官、律师和执法官。 "


west(129) = "在大部份的事情上都很幸运，有野心、意志坚强、好争论、聪明、通于世故，主宰和指挥别人，能以自己的方式享受人生。非常爱家庭，也一家会有一个家庭，来自小孩和幸福婚姻上的荣誉也是有的。  "
west(130) = "怜悯、冲动、执着、实际。会有来自亲戚的好运，当计划迷失时，有恼怒和粗卤的倾向，在艺术界能得到成功和认可。 "


west(201) = "意志坚强、顽固、勇敢，富于冒险精神、冲动而亲爱。爱好音乐、艺术和美好的东西。星相显示有航海生涯和艺术方面的成功。 "


west(202) = "大方、独立、热心、乐观、爱好自由而开放的生活、爱好宽大的地方，会使你在运动方面得到大众的认可，有很好的身体和永不疲倦的精力。"


west(203) = "创新、坚毅、聪明、喜欢研究，有杰出的科学才能，要抑制大胆和好冒险的倾向，有由于过份紧张和意外而遭受痛苦的倾向，要节制太急躁的脾气，并学习中庸之道。 "


west(204) = "感情强烈、能感动他人，必须谨防神经紧张，寒冷和着凉，身体不会很健康，有来自财产、文件和上司方面的麻烦。在法律、银行和保险方面会成功。 "


west(205) = "敏感而有反应的性情，有时太紧张而容易激动，无法预知的变化可能会影响你的财富，但是会有神秘的助力，渴望旅行和成功于异地。 "


west(206) = "个性强、有特殊才能、野心勃勃、时时追求完美、与亲朋很友善，可能会有相当的遗产继承，会有电子、现代科学方面的成功。 "


west(207) = "有点怀疑和不信任他人的天性，有阴谋和怪癖的性向，要抑制苛求朋友和同事的倾向，接受伟大的构想，但无法有效的付诸实现。"


west(208) = "生意和家事方面会有许多考验，可能会善变、犹豫而容易发火，不能同情别人的不幸，可能会过着平庸的生活，几乎没有好运气。 "


west(209) = "相当热心，对大众工作并有一些实际的才能，是创造自己命运的那种人。生意上由于别人的不小心，会导致你的突然损失，签署文件时必须小心。善用自己的判断力，事业和私人生活方面都要经常小心受骗。"


west(210) = "一生多事、变化很多，因为不好的健康、别人的出卖和欺骗，常有意外之灾害。有时坦白得近乎残忍，但本性是仁慈真诚的，天性很复杂。 "


west(211) = "天性殷勤、有些组织能力，能在写作中发现伟大的灵感，生意上容易遭受偶然的变化，慎防受骗。"

west(212) = "宽大的同情和强烈的感情，非常爱自己的母亲喜欢异性，在事业上、爱情大致都很成功，要抑制沮丧的情绪，否则会失去许多人生的乐趣。"

west(213) = "冲动、灵活的想像力、爱好新奇的变化、容易受人影响。一旦受激恼，会变得顽固而不原谅别人。在任何智识方面的职业都会成功。"


west(214) = "赋艺术、音乐和科学方面的头脑、冲动、急躁、侵略、而且很紧张，常常都很奢侈，有深沈的了解和很好的生意才能。 "

west(215) = "必须学习不要太奢侈。很聪明，但有苛求远不如你的人的倾向。性情冲动，运气时好时坏，抑制这种倾向，否则不良的健康会阻滞你的成功和幸福。 "


west(216) = "直觉锐利、性情仁慈大方，容易勃然大怒和憎恨，不可太悲观，在秘密事件和联合方面成功，可能会有一个富有的婚姻。 "

west(217) = "一般而言，生意、爱情都很成功，会来自亲友的利益，尤其是女性方面。很机智，并有相当的文学天赋，具许多好运和家庭的欢乐，强烈的社会主义倾向，并且是一位从事战争的民主政治拥护者。 "

west(218) = "能转厄运为胜利，具毁灭和创造自己命运的人格，虽会偶而因生人的闯入而失望，在爱情和生意上都会成功。智力出众，野心必能实现。 "




west(219) = "害羞、勤学而坚贞之天性，应随时克制妒羡的倾向，爱好文学、音乐和艺术，尽量利用你的智力和商业训练，应培养自信心和了解别人观点之能力。感情不稳定，有时很难了解。 "


west(220) = "野心勃勃、有领悟力，具非常吸引人的人格和上进冒险而坚强的脾气，会带来许多受人尊敬和有利的朋友，应克服过份自大的倾向，常思变换，会到许多地方旅行，应稍加调整，以达到一种比较中肯和稳定的观点。 "


west(221) = "意志坚强、有威仪、精力旺盛、直觉力和良好的生意眼，会很快在商场中辉煌腾达，有奢侈倾向，要稍加控制才好，应珍惜所拥有的。充满冒险和狂妄的梦想，要稍加调和。你必须了解：人生最大之报偿在善用自己的努力。 "


west(222) = "能创新、大胆而专横，避免太强烈的意志和独行的倾向，各种不同的教育和经商会使你得到很大益处，在商场僧很成功。应学习容忍和耐心，便可克服精力过剩的脾气。由于良好的本性和不智的投资，可能惹来金钱的麻烦。 "


west(223) = "机警而快活，会因结交有势力的朋友而大大成功。卓越的创新力和幽默感，使你在社交界中取得尊荣之地位，活泼的想像力和过剩的精力，会使你缺乏耐心、浮躁，且容易发火。 "


west(224) = "机警、好的执政能力、坚决地想要成功，不论是好是坏都照样做下去，应明辨对错。自由教育可克服许多不好的特质，要中肯些才好。 "

west(225) = "冲动而有点急进、聪明、很优越的直觉力，这种心智特质，常可在赫赫有名的成功人士身上发现，应控制冲动和过份急进之倾向，否则害处很多。 "

west(226) = "敏锐、精明、机警、富同情心，会成为人权斗士，喜欢干涉他人事物，动机是慈悲的，却会因此而有许多误会和失望。艺术、文学和科学方面会成功。 "

west(227) = "适应力强、多才多艺、有直觉力，要学习专心致志才行。很容易厌倦，常常毫无目标地由一件事游移到另一件事。个性迷人、有幽默感、会吸引许多朋友、善用良好的特质，必能助你达到成功和幸福。 "


west(228) = "悲天悯人的心性，很讨人喜欢。应学习自立和外交手腕，以克服好冒险的精神。喜欢质问、悟力佳，能成功于职业界，受科学研究之吸引。可以和形形色色的人相处很好，这会助你成功。 "


west(229) = "具悦人而彬彬有礼之气质和机警直觉的脑子，必须经常检讨自己那种不顾後果，想主宰别人和当领导人物的倾向。勤劳而自立，具很大的野心和果决力。 "

west(301) = "赋迷人的丰采和伟大的磁性，创造发明的脑子，观念上和工作上都表现出创新的本领，小心伤风，避免接近风口和预防气温之突变，要控制投机的欲望，否则会有许多失望。一生中，要安安逸逸的度过颇不易。 "


west(302) = "温和而怜悯的性情，脑力佳而慈悲，有吸引人之人格，能带来许多忠诚的友谊，也可能因此养成依赖的个性，父母和亲戚之过份关心，也会养成依赖的性情。 "

west(303) = "灵活而快乐的天性、聪明伶俐、富于想像力、为一己之利，有过份慷慨之倾向，还是谨慎点较妥，有把好运看做理所当然之倾向。 "

west(304) = "对人生有一种安静和平的态度，坦白之心胸和人道主义之原则，会使你深受朋友和同事之锺爱，会来自颇有影响力之社团的利益，有把好运视为理所当然的倾向，应加以修正才行。 "


west(305) = "乐于行善、慈悲为怀、愿意帮助前来求助的人，会导致外开销而引起财务困难，喜欢受阿谀，容易被狡诈的朋友所利用，星相显示有受骗的倾向。 "

west(306) = "精力充沛、做事很有技巧，即使面对困难，仍能坚持而果断，会带来成功和中年财务之独立，要学习谅解别人的感情和容忍，有点冲动。 "

west(307) = "敏感而理想主义的天性，友善而仁慈的性情，都是你的好运气，有太慷慨的倾向，要学习辨真假。喜欢水路旅行和长途的海上航行，会带来有利的商业关系。 "

west(308) = "具聪明、科学和试验的脑筋，很杰出的健谈人物，应该会受欢迎才是。独立而有条理的精神和迷人的个性，使你很容易由父母、朋友和同事得到所要的。 "

west(309) = "悦人的人格、极佳的直觉力、非常感情用事，可能会形成无法控制的脾气。有想像力，长久的单独会有忧郁和做白日梦的倾向，最好和同年纪的人在一起，并避免太过固执。 "


west(310) = "认真而渴望、强大的心智能力、深沈而实际、独立感很强，希望帮助不幸的人，有时为了自己的好处，会变得太豁达大方。很尊敬老人，会开创自己的命运，经营新的企业和娱乐大众之机构会成功。 "


west(311) = "创新而上进冒险的脑子，性情和谐，很容易满足，可能会缺少克服困难的动机，要培养自信心和学习获致成功的实际方法，需要有快乐的同年伙伴，应缓和太高的希望和太大的野心，并要记得现实。 "


west(312) = "极佳的个性、怜悯的心胸、感情深邃、非常聪明，会很早熟，有过份保守和用功的倾向，需要有快乐的环境，多做户外活动。你感情用事的脾气需要耐心而了解的处理才行。 "


west(313) = "舒畅和蔼的人格、风度迷人、谈吐风趣、社交生活中会很受欢迎、爱好旅行、感情稳定，具艺术气质和迷人的个性，适于职业生涯，施以适当的教育，必能在文学界得到大众认可。 "


west(314) = "精明、实际而有直觉力，艺术气质加上人道主义的原则，给来许多持久而有影响力的朋友，仪态大方而优雅、仁慈、喜爱运动，有逃避责任和好逸恶劳的倾向。善用天赋才能是成功所必须。 "


west(315) = "敏锐的智力和卓越的集中力，避免把人生看得太认真，明朗快活的气氛，和性情相近的朋友相处，可助你除去这种倾向。非常有吸引力的个性，独立的天性，使你在商场非常成功。爱好音乐和旅行。 "


west(316) = "非凡的心智特质，从事科学工作。会形成一种坚强而固定的脾气，具否定传统、大胆创新倾向。要克制爱情和生意方面做急进之改善，很容易因受骗和不智之投机而遭受痛苦。 "


west(317) = "敏感，容易受感动，喜欢质问和到处旅行。有时因脾气暴躁，显得难应付。相当奢侈，应稍加节制，以免晚景陷入严重财务困窘。对人道主义之活动有很明显的兴趣，应加以鼓励才是。 "


west(318) = "浪漫的天性，非常自我中心和难取悦，喜欢和年长的人在一起，应试着和同年的人做朋友，学习容忍和了解。喜欢人生中美好的东西，会因此而人、财两空，应加以节制。 "


west(319) = "明朗快活、慷慨大方，无论公私均正直，应追求一种能使心智能力得到最大发挥之生涯，非常眷恋家庭和父母，喜欢穿着得很好，且经常走在时尚尖端。 "

west(320) = "不自私、慈悲为怀、要学习专注，克服一次想做好多事的倾向，能成为口才流利的健谈人物和多产作家，应训练自己朝固定之目标或职业，化天赋才能为资产。一旦能控制自己的感情，就可成功并得到大众之认可，会比一般人还幸运。 "
response.write(west(day)+"<br>"+ west(100 * month + day)+"<br>")
end sub
%>
  </span></td>
</tr>
    </TBODY>
</TABLE>
