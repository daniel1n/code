<!--#include file="Connections/oufuconnt.asp" -->
<link href="/css/layout0<%=(wzgl.Fields.Item("moban").Value)%>.css" rel="stylesheet" type="text/css" />
  <script>
//��ʼ��
var def="1";
function mover(object){
  //���˵�
  var mm=document.getElementById("m_"+object);
  mm.className="m_li_a";
  //��ʼ���˵�������Ч��
  if(def!=0){
    var mdef=document.getElementById("m_"+def);
    mdef.className="m_li";
  }
  //�Ӳ˵�
  var ss=document.getElementById("s_"+object);
  ss.style.display="block";
  //��ʼ�Ӳ˵�������Ч��
  if(def!=0){
    var sdef=document.getElementById("s_"+def);
    sdef.style.display="none";
  }
}

function mout(object){
  //���˵�
  var mm=document.getElementById("m_"+object);
  mm.className="m_li";
  //��ʼ���˵���ԭЧ��
  if(def!=0){
    var mdef=document.getElementById("m_"+def);
    mdef.className="m_li_a";
  }
  //�Ӳ˵�
  var ss=document.getElementById("s_"+object);
  ss.style.display="none";
  //��ʼ�Ӳ˵���ԭЧ��
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
    <li id="m_1" class='m_li_a'><a href="/">�� ҳ</a></li>
    <li class="m_line"><img src="/images/line1.gif" /></li>
    <li id="m_3" class='m_li' onmouseover='mover(3);' onmouseout='mout(3);'><a href="/ppziwei/">��������</a></li>
    <li class="m_line"><img src="/images/line1.gif" /></li>
    <li id="m_2" class='m_li' onmouseover='mover(2);' onmouseout='mout(2);'><a href="#">����ѧ</a></li>
    <li class="m_line"><img src="/images/line1.gif" /></li>
    <li id="m_4" class='m_li' onmouseover='mover(4);' onmouseout='mout(4);'><a href="/sm/">����</a></li>
    <li class="m_line"><img src="/images/line1.gif" /></li>
    <li id="m_5" class='m_li' onmouseover='mover(5);' onmouseout='mout(5);'><a href="/guanyin/">��ǩ</a></li>
    <li class="m_line"><img src="/images/line1.gif" /></li>
    <li id="m_6" class='m_li' onmouseover='mover(6);' onmouseout='mout(6);'><a href="/zhixiang/">��ѧ</a></li>
    <li class="m_line"><img src="/images/line1.gif" /></li>
    <li id="m_7" class='m_li' onmouseover='mover(7);' onmouseout='mout(7);'><a href="/xmwgpd/">����</a></li>
    <li class="m_line"><img src="/images/line1.gif" /></li>
    <li id="m_8" class='m_li' onmouseover='mover(8);' onmouseout='mout(8);'><a href="/sanshishu/">Ԥ��</a></li>
    <li class="m_line"><img src="/images/line1.gif" /></li>
    <li id="m_9" class='m_li' onmouseover='mover(9);' onmouseout='mout(9);'><a href="/#mn">��������</a></li>
    <li class="m_line"><img src="/images/line1.gif" /></li>
    <li id="m_10" class='m_li' onmouseover='mover(10);' onmouseout='mout(10);'><a href="#">�ֻ��� </a></li>
    <li class="m_line"><img src="/images/line1.gif" /></li>
  </ul>
</div>
<table width="960" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td>
<div id="menu2" style="height:32px; background-color:#FFFFFF;">
   <ul class="smenu">
     <li style="padding-left:15px;" id="s_1" class='s_li_a'><a href="/zgjm/">�ܹ�����</a> | <a href="#">2016���˳�</a> | <a href="/pp6y/">��س��Ǯ��</a> | <a href="/xingzuo/">��������</a> | <a href="#">���Ƽ���</a> | <a href="/shouji/">�ֻ�����</a> | <a href="/ymjxcs/">��������</a> | <a href="/sfz/">���֤����</a>
     </li>
     <li style="padding-left:80px;" id="s_3" class='s_li' onmouseover='mover(3);' onmouseout='mout(3);'><a href="/ppbazi/">��������</a> | <a href="/pp6r/">��������</a> | <a href="/ppqmdj/">���Ŷݼ�</a> | <a href="/pp6y/">��س����</a> | <a href="/ppjkj/">��ھ�����</a> | <a href="/ppxk/">���շ���</a> | <a href="/ppziwei/">��΢����</a></li>
     <li style="padding-left:165px;" id="s_2" class='s_li' onmouseover='mover(2);' onmouseout='mout(2);'><a href="#">�������� | <a href="/sm/sm.asp?sm=5">�������� | <a href="/sm/sm.asp?sm=5">��������</a> | <a href="#">��˾������</a></li>
     <li style="padding-left:250px;" id="s_4" class='s_li' onmouseover='mover(4);' onmouseout='mout(4);'><a href="/sm/sm.asp?sm=2">��������</a> | <a href="/ppbazi/">��������</a> | <a href="/sanshishu/">���������</a> | <a href="/sm/sm.asp?sm=4">�ƹ�����</a> | <a href="/sm/sm.asp?sm=3">�ո�����</a> | <a href="/sm/sm.asp?sm=7">ǰ����˭</a></li>
     <li style="padding-left:325px;" id="s_5" class='s_li' onmouseover='mover(5);' onmouseout='mout(5);'><a href="/guanyin/">������ǩ</a> | <a href="/lvzu/">������ǩ</a> | <a href="/huangdaxian/">�ƴ�����ǩ</a> | <a href="/guandi/">��ʥ����ǩ</a> | <a href="/mazu/">������ǩ</a> | <a href="/baifo/">���߰ݷ�</a></li>
     <li style="padding-left:400px;" id="s_6" class='s_li' onmouseover='mover(6);' onmouseout='mout(6);'><a href="/mianxiang/">����ͼ��</a> | <a href="/zhixiang/">��Ů����</a> | <a href="/nanzhi/">������ī��</a> | <a href="/nvzhi/">Ů����ī��</a> | <a href="/zwsm/">��ָ������</a></li>
     <li style="padding-left:480px;" id="s_7" class='s_li' onmouseover='mover(7);' onmouseout='mout(7);'><a href="/xmwgpd/">��������</a> | <a href="/xzpd/">�������</a> | <a href="/sxpd/">��Ф���</a> | <a href="/xxpd/">Ѫ�����</a> | <a href="/qqtest/">QQ�������</a></li>
     <li style="padding-left:350px;" id="s_8" class='s_li' onmouseover='mover(8);' onmouseout='mout(8);'><a href="/snsn/">������Ů</a> | <a href="/ytyc/">����Ԥ��</a> | <a href="/mryc/">����Ԥ��</a> | <a href="/ptyc/">����Ԥ��</a> | <a href="/xjyc/">�ľ�Ԥ��</a> | <a href="/emyc/">����Ԥ��</a> | <a href="/sanshishu/">����Ԥ��</a></li>
     <li style="padding-left:230px;" id="s_9" class='s_li' onmouseover='mover(9);' onmouseout='mout(9);'><a href="/wannianli/">������</a> | <a href="/xingzuo/">��������</a> | <a href="/baifo/">���߰ݷ�</a> | <a href="/zgjm/">�ܹ�����</a> | <a href="/tarot/">������</a> | <a href="/taiyang/">��̫��</a> | <a href="/jingdu/">��������</a> | <a href="/renpin/">��Ʒ����</a> | <a href="/snsn/">������Ů</a></li>
   </ul>
</div>
  </td>
  </tr>
  <tr>
    <td><div id="foot_ad"><%=(adgl.Fields.Item("addaima02").Value)%></div></td>
  </tr>
</table>