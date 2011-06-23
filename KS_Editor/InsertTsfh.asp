<HTML>
<HEAD>
<TITLE>插入特殊符号</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<STYLE type=text/css>
table.content {background-color:#000000;width:100%;}
table.content td {background-color:#ffffff;width:18px;height:18px;text-align:center;vertical-align:middle;cursor:pointer;}
.card {cursor:pointer;background-color:#3A6EA5;text-align:center;}
</STYLE>
<SCRIPT language=JavaScript>

// 选项卡点击事件
function cardClick(cardID){
	var obj;
	for (var i=1;i<7;i++){
		obj=document.all("card"+i);
		obj.style.backgroundColor="#3A6EA5";
		obj.style.color="#FFFFFF";
	}
	obj=document.all("card"+cardID);
	obj.style.backgroundColor="#FFFFFF";
	obj.style.color="#3A6EA5";

	for (var i=1;i<7;i++){
		obj=document.all("content"+i);
		obj.style.display="none";
	}
	obj=document.all("content"+cardID);
	obj.style.display="";
}

// 预览
function SymbolOver(){
	var el=event.srcElement
	preview.innerHTML=el.innerHTML;
}

// 点击返回
function SymbolClick(){
	var el=event.srcElement;
        window.returnValue=el.innerHTML;
	window.close();
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>
<link href="Editor.css" rel="stylesheet" type="text/css">
</HEAD>

<BODY bgColor=#D4D0C8>

<table border=0 cellpadding=0 cellspacing=0><tr valign=top><td>
<fieldset><legend><b>插入特殊符号</b></legend><br><table border=0 cellpadding=3 cellspacing=0>
<tr align=center>
	<td class="card" onClick="cardClick(1)" id="card1">特殊</td>
	<td width=2></td>
	<td class="card" onClick="cardClick(2)" id="card2">标点</td>
	<td width=2></td>
	<td class="card" onClick="cardClick(3)" id="card3">数学</td>
	<td width=2></td>
	<td class="card" onClick="cardClick(4)" id="card4">单位</td>
	<td width=2></td>
	<td class="card" onClick="cardClick(5)" id="card5">数字</td>
	<td width=2></td>
	<td class="card" onClick="cardClick(6)" id="card6">拼音</td>
</tr>
<tr>
	<td bgcolor=#ffffff align=center valign=middle colspan=11>
	<table border=0 cellpadding=3 cellspacing=1 class="content" id="content1">
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">＃</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">＠</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">＆</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">＊</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">※</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">§</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">〃</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">№</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">〓</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">○</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">●</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">△</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">▲</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">◎</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">☆</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">★</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">◇</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">◆</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">□</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">■</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">▽</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">▼</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">㊣</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">℅</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ˉ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">￣</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">＿</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹉</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹊</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹍</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹎</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹋</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹌</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹟</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹠</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹡</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">♀</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">♂</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⊕</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⊙</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">↑</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">↓</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">←</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">→</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">↖</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">↗</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">↙</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">↘</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">∥</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">∣</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">／</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">＼</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">∕</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹨</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">&#65533;</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">&yen;</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">&pound;</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">&#8482;</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">&reg;</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">&copy;</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">  </td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
	</tr>
	</table>
	<table border=0 cellpadding=3 cellspacing=1 class="content" id="content2">
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">，</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">、</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">。</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">．</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">；</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">：</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">？</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">！</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">︰</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">…</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">‥</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">′</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">‵</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">々</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">～</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">‖</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ˇ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ˉ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹐</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹑</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹒</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">·</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹔</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹕</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹖</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹗</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">｜</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">-</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">︱</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">-</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">︳</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">︴</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹏</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">（</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">）</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">︵</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">︶</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">｛</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">｝</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">︷</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">︸</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">〔</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">〕</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">︹</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">︺</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">【</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">】</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">︻</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">︼</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">《</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">》</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">︽</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">︾</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">〈</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">〉</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">︿</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹀</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">「</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">」</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹁</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹂</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">『</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">』</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹃</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹄</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹙</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹚</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹛</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹜</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹝</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹞</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">'</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">'</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">"</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">"</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">〝</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">〞</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ˋ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ˊ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
	</tr>
	</table>
	<table border=0 cellpadding=3 cellspacing=1 class="content" id="content3">
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">≈</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">≡</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">≠</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">＝</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">≤</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">≥</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">＜</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">＞</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">≮</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">≯</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">∷</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">±</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">＋</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">－</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">×</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">÷</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">／</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">∫</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">∮</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">∝</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">∞</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">∧</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">∨</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">∑</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">∏</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">∪</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">∩</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">∈</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">∵</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">∴</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⊥</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">∥</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">∠</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⌒</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⊙</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">≌</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">∽</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">√</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">≦</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">≧</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">≒</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">≡</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹢</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹣</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹤</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹥</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹦</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">～</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">∟</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⊿</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">㏒</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">㏑</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
	</tr>
	</table>
	<table border=0 cellpadding=3 cellspacing=1 class="content" id="content4">
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">°</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">′</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">″</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">＄</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">￥</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">〒</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">￠</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">￡</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">％</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">＠</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">℃</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">℉</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹩</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹪</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">‰</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">﹫</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">㏕</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">㎜</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">㎝</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">㎞</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">㏎</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">㎡</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">㎎</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">㎏</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">㏄</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">°</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">○</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">¤</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
	</tr>
	</table>
	<table border=0 cellpadding=3 cellspacing=1 class="content" id="content5">
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ⅰ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ⅱ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ⅲ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ⅳ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ⅴ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ⅵ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ⅶ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ⅷ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ⅸ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ⅹ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">Ⅰ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">Ⅱ</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">Ⅲ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">Ⅳ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">Ⅴ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">Ⅵ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">Ⅶ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">Ⅷ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">Ⅸ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">Ⅹ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">Ⅺ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">Ⅻ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⒈</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⒉</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⒊</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⒋</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⒌</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⒍</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⒎</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⒏</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⒐</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⒑</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⒒</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⒓</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⒔</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⒕</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⒖</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⒗</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⒘</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⒙</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⒚</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⒛</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⑴</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⑵</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⑶</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⑷</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⑸</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⑹</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⑺</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⑻</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⑼</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⑽</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⑾</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⑿</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⒀</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⒁</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⒂</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⒃</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⒄</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⒅</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⒆</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⒇</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">①</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">②</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">③</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">④</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⑤</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⑥</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⑦</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⑧</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⑨</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">⑩</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">㈠</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">㈡</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">㈢</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">㈣</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">㈤</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">㈥</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">㈦</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">㈧</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">㈨</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">㈩</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
	</tr>
	</table>
	<table border=0 cellpadding=3 cellspacing=1 class="content" id="content6">
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ā</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">á</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ǎ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">à</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ō</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ó</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ǒ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ò</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ē</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">é</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ě</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">è</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ī</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">í</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ǐ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ì</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ū</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ú</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ǔ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ù</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ǖ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ǘ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ǚ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ǜ</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ü</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ê</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ɑ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()"></td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ń</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ň</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ǹ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">ɡ</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">　</td>
	</tr>
	</table>

	</td>
</tr>
</table>
</fieldset>
</td><td width=10></td><td>
<table border=0 cellpadding=0 cellspacing=0>
  <tr><td height=25></td></tr>
  <tr><td align=center>预览</td></tr>
  <tr><td height=10></td></tr>
  <tr>
    <td align=center valign=middle>
      <table border=0 cellpadding=0 cellspacing=1 bgcolor=#000000>
        <tr>
	  <td bgcolor=#ffffff style="font-size:32px;color:#0000ff" id=preview align=center valign=middle width=50 height=50>
	  </td>
	</tr>
      </table>
    </td>
  </tr>
  <tr><td height=52></td></tr>
  <tr><td align=center><input type=button value='  取消  ' onClick="window.close();"></td></tr>
</table>
</td></tr></table>
<script language=javascript>
cardClick(1);
</script>
</BODY>
</HTML> 
