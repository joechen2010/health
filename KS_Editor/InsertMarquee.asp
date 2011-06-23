<HTML>
<HEAD>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<style type="text/css">
<!--
.Separator 
{
	BORDER-RIGHT: buttonhighlight 1px solid;
	FONT-SIZE: 0px;
	BORDER-LEFT: buttonshadow 1px solid;
	cursor: default;
	height: 50px;
	width: 1px;
	top: 10px;
}
-->
</style>
<script language=javascript>
var sAction = "INSERT";
var sTitle = "插入";
var el;
var sText = "";
var sBehavior = "";
document.write("<title>滚动文本（" + sTitle + "）</title>");


// 单选的点击事件
function check(){
	sBehavior = event.srcElement.value;
}

// 初始值
function InitDocument() {
	d_text.value = sText;
	switch (sBehavior) {
	case "scroll":
		document.all("d_behavior")[0].checked = true;
		break;
	case "slide":
		document.all("d_behavior")[1].checked = true;
		break;
	default:
		sBehavior = "alternate";
		document.all("d_behavior")[2].checked = true;
		break;
	}

}
</script>
<SCRIPT event=onclick for=Ok language=JavaScript>
	sText = d_text.value;
	if (sText!='')
	{
		if (sAction=="MODI") 
		{
			el.behavior=sBehavior;
			el.innerHTML=sText;
		}
		else
		{
			var str1;
			str1="<marquee behavior='"+sBehavior+"'>"+sText+"</marquee>"
		}
		window.returnValue = str1
		window.close();
	}
</script>
</HEAD>
<link rel="stylesheet" href="Editor.css">
<body bgcolor=menu onload="InitDocument()">
<div align="center">
  <table border=0 cellpadding=0 cellspacing=0>
    <tr valign=middle> 
      <td width="37">文本:</td>
      <td> 
        <textarea name="text" style="width:100%;" cols="20" rows="2" id="d_text"></textarea> 
      </td>
      <td width="80" rowspan="2" align="center" valign="top"> 
        <input name="submit" type=submit id=Ok value=' 确 定 '> 
        <br> <br> <input name="button" type=button onClick="window.close();" value=' 取 消 '></td>
    </tr>
    <tr valign=middle> 
      <td height="30">表现:</td>
      <td height="30">
<input onclick="check()" type="radio" name="d_behavior" value="scroll">
        滚动条 
        <input onclick="check()" type="radio" name="d_behavior" value="slide">
        幻灯片 
        <input onclick="check()" type="radio" name="d_behavior" value="alternate">
        交替</td>
    </tr>
  </table>
</div>
</body>
</html>
<script language="JavaScript">
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script> 
