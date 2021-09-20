<!--#include file="Connections/oufuconnt.asp" -->
<link href="/css/layout0<%=(wzgl.Fields.Item("moban").Value)%>.css" rel="stylesheet" type="text/css" />
  <script>
//初始化
var def="1";
function mover(object){
  //主菜单
  var mm=document.getElementById("m_"+object);
  mm.className="m_li_a";
  //初始主菜单先隐藏效果
  if(def!=0){
    var mdef=document.getElementById("m_"+def);
    mdef.className="m_li";
  }
  //子菜单
  var ss=document.getElementById("s_"+object);
  ss.style.display="block";
  //初始子菜单先隐藏效果
  if(def!=0){
    var sdef=document.getElementById("s_"+def);
    sdef.style.display="none";
  }
}

function mout(object){
  //主菜单
  var mm=document.getElementById("m_"+object);
  mm.className="m_li";
  //初始主菜单还原效果
  if(def!=0){
    var mdef=document.getElementById("m_"+def);
    mdef.className="m_li_a";
  }
  //子菜单
  var ss=document.getElementById("s_"+object);
  ss.style.display="none";
  //初始子菜单还原效果
  if(def!=0){
    var sdef=document.getElementById("s_"+def);
    sdef.style.display="block";
  }
}
</script>

<style type="text/css">
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
</style>

<div id="top_logoad">
  <table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td><img src="/images/logo.gif" width="200" height="60" /></td>
      <td align="center"><p><%=(adgl.Fields.Item("addaima01").Value)%></p></td>
    </tr>
  </table>
</div>
<div id="menu">
  <ul>
    <li class="m_line"><img src="/images/line1.gif" /></li>
    <li id="m_1" class='m_li_a'><a href="/">首 页</a></li>
    <li class="m_line"><img src="/images/line1.gif" /></li>
    <li id="m_3" class='m_li' onmouseover='mover(3);' onmouseout='mout(3);'><a href="/ppziwei/">起卦排盘</a></li>
    <li class="m_line"><img src="/images/line1.gif" /></li>
    <li id="m_2" class='m_li' onmouseover='mover(2);' onmouseout='mout(2);'><a href="#">姓名学</a></li>
    <li class="m_line"><img src="/images/line1.gif" /></li>
    <li id="m_4" class='m_li' onmouseover='mover(4);' onmouseout='mout(4);'><a href="/sm/">八字</a></li>
    <li class="m_line"><img src="/images/line1.gif" /></li>
    <li id="m_5" class='m_li' onmouseover='mover(5);' onmouseout='mout(5);'><a href="/guanyin/">抽签</a></li>
    <li class="m_line"><img src="/images/line1.gif" /></li>
    <li id="m_6" class='m_li' onmouseover='mover(6);' onmouseout='mout(6);'><a href="/zhixiang/">相学</a></li>
    <li class="m_line"><img src="/images/line1.gif" /></li>
    <li id="m_7" class='m_li' onmouseover='mover(7);' onmouseout='mout(7);'><a href="/xmwgpd/">速配</a></li>
    <li class="m_line"><img src="/images/line1.gif" /></li>
    <li id="m_8" class='m_li' onmouseover='mover(8);' onmouseout='mout(8);'><a href="/sanshishu/">预测</a></li>
    <li class="m_line"><img src="/images/line1.gif" /></li>
    <li id="m_9" class='m_li' onmouseover='mover(9);' onmouseout='mout(9);'><a href="/#mn">其他工具</a></li>
    <li class="m_line"><img src="/images/line1.gif" /></li>
    <li id="m_10" class='m_li' onmouseover='mover(10);' onmouseout='mout(10);'><a href="#">手机版 </a></li>
    <li class="m_line"><img src="/images/line1.gif" /></li>
  </ul>
</div>
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>
<div id="menu2" style="height:32px; background-color:#FFFFFF;">
   <ul class="smenu">
     <li style="padding-left:15px;" id="s_1" class='s_li_a'><a href="/zgjm/">周公解梦</a> | <a href="#">2016年运程</a> | <a href="/pp6y/">六爻金钱卦</a> | <a href="/xingzuo/">今日星座</a> | <a href="#">车牌吉凶</a> | <a href="/shouji/">手机吉凶</a> | <a href="/ymjxcs/">域名吉凶</a> | <a href="/sfz/">身份证测算</a>
     </li>
     <li style="padding-left:80px;" id="s_3" class='s_li' onmouseover='mover(3);' onmouseout='mout(3);'><a href="/ppbazi/">四柱八字</a> | <a href="/pp6r/">六壬排盘</a> | <a href="/ppqmdj/">奇门遁甲</a> | <a href="/pp6y/">六爻起卦</a> | <a href="/ppjkj/">金口诀排盘</a> | <a href="/ppxk/">玄空飞星</a> | <a href="/ppziwei/">紫微斗数</a></li>
     <li style="padding-left:165px;" id="s_2" class='s_li' onmouseover='mover(2);' onmouseout='mout(2);'><a href="#">宝宝起名 | <a href="/sm/sm.asp?sm=5">姓名测算 | <a href="/sm/sm.asp?sm=5">姓名评分</a> | <a href="#">公司名评分</a></li>
     <li style="padding-left:250px;" id="s_4" class='s_li' onmouseover='mover(4);' onmouseout='mout(4);'><a href="/sm/sm.asp?sm=2">八字详批</a> | <a href="/ppbazi/">八字排盘</a> | <a href="/sanshishu/">三世书财运</a> | <a href="/sm/sm.asp?sm=4">称骨论命</a> | <a href="/sm/sm.asp?sm=3">日干论命</a> | <a href="/sm/sm.asp?sm=7">前世是谁</a></li>
     <li style="padding-left:325px;" id="s_5" class='s_li' onmouseover='mover(5);' onmouseout='mout(5);'><a href="/guanyin/">观音灵签</a> | <a href="/lvzu/">吕祖灵签</a> | <a href="/huangdaxian/">黄大仙灵签</a> | <a href="/guandi/">关圣帝灵签</a> | <a href="/mazu/">妈祖灵签</a> | <a href="/baifo/">在线拜佛</a></li>
     <li style="padding-left:400px;" id="s_6" class='s_li' onmouseover='mover(6);' onmouseout='mout(6);'><a href="/mianxiang/">面相图解</a> | <a href="/zhixiang/">男女脸痣</a> | <a href="/nanzhi/">男身体墨痣</a> | <a href="/nvzhi/">女身体墨痣</a> | <a href="/zwsm/">手指纹算命</a></li>
     <li style="padding-left:480px;" id="s_7" class='s_li' onmouseover='mover(7);' onmouseout='mout(7);'><a href="/xmwgpd/">姓名速配</a> | <a href="/xzpd/">星座配对</a> | <a href="/sxpd/">生肖配对</a> | <a href="/xxpd/">血型配对</a> | <a href="/qqtest/">QQ号码配对</a></li>
     <li style="padding-left:350px;" id="s_8" class='s_li' onmouseover='mover(8);' onmouseout='mout(8);'><a href="/snsn/">生男生女</a> | <a href="/ytyc/">眼跳预测</a> | <a href="/mryc/">面热预测</a> | <a href="/ptyc/">喷嚏预测</a> | <a href="/xjyc/">心惊预测</a> | <a href="/emyc/">耳鸣预测</a> | <a href="/sanshishu/">财运预测</a></li>
     <li style="padding-left:230px;" id="s_9" class='s_li' onmouseover='mover(9);' onmouseout='mout(9);'><a href="/wannianli/">万年历</a> | <a href="/xingzuo/">今日星座</a> | <a href="/baifo/">在线拜佛</a> | <a href="/zgjm/">周公解梦</a> | <a href="/tarot/">塔罗牌</a> | <a href="/taiyang/">真太阳</a> | <a href="/jingdu/">地区经度</a> | <a href="/renpin/">人品计算</a> | <a href="/snsn/">生男生女</a></li>
   </ul>
</div>
  </td>
  </tr>
  <tr>
    <td><div id="foot_ad"><%=(adgl.Fields.Item("addaima02").Value)%></div></td>
  </tr>
</table>