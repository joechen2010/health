<html>
<head>
<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
<meta HTTP-EQUIV="Cache-Control" content="no-cache">
<meta HTTP-EQUIV="Expires" content="0">'
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>图文并排属性</title>
<link rel=stylesheet type=text/css href="editor.css">
<style type="text/css">
<!--
.BtnTemplet {
	cursor: default;
	border: 2px solid buttonface;
}
.BtnTempletMouseDown {
	border-top-width: 2px;
	border-right-width: 2px;
	border-bottom-width: 2px;
	border-left-width: 2px;
	border-top-style: solid;
	border-right-style: solid;
	border-bottom-style: solid;
	border-left-style: solid;
	border-top-color: #666666;
	border-right-color: #FFFFFF;
	border-bottom-color: #FFFFFF;
	border-left-color: #666666;
}
-->
</style>
<script language=javascript>
var CurrMode;
function ClickType(Param)
{
	CurrMode=Param;
	AllImgs=document.all.tags('IMG');
	var el=event.srcElement;
	for(i=0;i<AllImgs.length;i++) AllImgs[i].className='BtnTemplet';
	el.className='BtnTempletMouseDown';
}
function SubmitFun()
{
	var error = 0;
	if(error!= 1)
	{
		returnValue=new Object();
		returnValue.style=new Object();
		switch(CurrMode)
		{
			case 'left':
				returnValue.style.position='static';
				returnValue.align='left';
				break;
			case 'right':
				returnValue.style.position='static';
				returnValue.align='right';
				break;
			case 'center':
				returnValue.style.position='static';
				returnValue.align='center';
				break;
			case 'over':
				returnValue.style.position='absolute';
				returnValue.style.zIndex=100;
				break;
			case 'behind':
				returnValue.style.position='absolute';
				returnValue.style.zIndex=-100;
				break;
			default:
				returnValue.style.position='static';
				returnValue.style.align='';
		}
		window.close()
	}
	else
	{
		return false;
	}
}
</script>
</head>
<body>
<table align=center>
 <form name=imageForm onsubmit='SubmitFun();'>
  <tr>
   <td>
        <fieldset><legend>图片位置</legend>
        <img src=images/Vnone.gif width="50" height="50" hspace=10 vspace=10 class="BtnTemplet" onclick='ClickType("none");'> 
        <img src=images/Vleft.gif width="50" height="50" hspace=10 vspace=10 class="BtnTemplet" onclick='ClickType("left");'> 
        <img src=images/Vright.gif width="50" height="50" hspace=10 vspace=10 class="BtnTemplet" onclick='ClickType("right");'> 
        <img src=images/Vcenter.gif width="50" height="50" hspace=10 vspace=10 class="BtnTemplet" onclick='ClickType("center");'> 
        <img src=images/VCenterTop.gif width="50" height="50" hspace=10 vspace=10 class="BtnTemplet" onclick='ClickType("over");'> 
        <img src=images/VCenterBottom.gif width="50" height="50" hspace=10 vspace=10 class="BtnTemplet" onclick='ClickType("behind");'> 
        </fieldset> 
        <div align="center"><br>
          <input type=button onClick="SubmitFun();" value=" 确 定 ">
          &nbsp;&nbsp;&nbsp;&nbsp;
          <input type=button value=" 取 消 " onClick="self.close();">
        </div>
	</td>
  </tr>
 </form>
</table>
</body>
</html> 
