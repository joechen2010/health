<HTML>
<HEAD>
<TITLE>�����������</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<STYLE type=text/css>
table.content {background-color:#000000;width:100%;}
table.content td {background-color:#ffffff;width:18px;height:18px;text-align:center;vertical-align:middle;cursor:pointer;}
.card {cursor:pointer;background-color:#3A6EA5;text-align:center;}
</STYLE>
<SCRIPT language=JavaScript>

// ѡ�����¼�
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

// Ԥ��
function SymbolOver(){
	var el=event.srcElement
	preview.innerHTML=el.innerHTML;
}

// �������
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
<fieldset><legend><b>�����������</b></legend><br><table border=0 cellpadding=3 cellspacing=0>
<tr align=center>
	<td class="card" onClick="cardClick(1)" id="card1">����</td>
	<td width=2></td>
	<td class="card" onClick="cardClick(2)" id="card2">���</td>
	<td width=2></td>
	<td class="card" onClick="cardClick(3)" id="card3">��ѧ</td>
	<td width=2></td>
	<td class="card" onClick="cardClick(4)" id="card4">��λ</td>
	<td width=2></td>
	<td class="card" onClick="cardClick(5)" id="card5">����</td>
	<td width=2></td>
	<td class="card" onClick="cardClick(6)" id="card6">ƴ��</td>
</tr>
<tr>
	<td bgcolor=#ffffff align=center valign=middle colspan=11>
	<table border=0 cellpadding=3 cellspacing=1 class="content" id="content1">
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�I</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�G</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�h</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�i</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�l</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�m</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�j</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�k</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�|</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�}</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�~</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�I</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�J</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�L</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�K</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�O</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�M</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">&#65533;</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">&yen;</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">&pound;</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">&#8482;</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">&reg;</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">&copy;</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">  </td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	</table>
	<table border=0 cellpadding=3 cellspacing=1 class="content" id="content2">
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�U</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�E</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�F</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�o</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�p</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�q</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�r</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�s</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�t</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�u</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">-</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">-</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�n</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�v</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�w</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�x</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�y</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�z</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�{</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">'</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">'</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">"</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">"</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�A</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�@</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	</table>
	<table border=0 cellpadding=3 cellspacing=1 class="content" id="content3">
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�Q</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�R</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�P</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�N</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�S</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�S</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�R</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	</table>
	<table border=0 cellpadding=3 cellspacing=1 class="content" id="content4">
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�H</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�T</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�L</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�M</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�N</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�Q</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�O</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�J</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�K</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">�P</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	</table>
	<table border=0 cellpadding=3 cellspacing=1 class="content" id="content5">
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	</table>
	<table border=0 cellpadding=3 cellspacing=1 class="content" id="content6">
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	<tr>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
		<td onMouseOver="SymbolOver()" onClick="SymbolClick()">��</td>
	</tr>
	</table>

	</td>
</tr>
</table>
</fieldset>
</td><td width=10></td><td>
<table border=0 cellpadding=0 cellspacing=0>
  <tr><td height=25></td></tr>
  <tr><td align=center>Ԥ��</td></tr>
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
  <tr><td align=center><input type=button value='  ȡ��  ' onClick="window.close();"></td></tr>
</table>
</td></tr></table>
<script language=javascript>
cardClick(1);
</script>
</BODY>
</HTML> 
