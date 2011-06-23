<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>修改属性</title>
</head>
<script language="javascript" event="onerror(msg, url, line)" for="window">return true;</script>
<style type="text/css">
<!--
.Separator 
{
	BORDER-RIGHT: buttonhighlight 1px solid;
	FONT-SIZE: 0px;
	BORDER-LEFT: buttonshadow 1px solid;
	cursor: default;
	height: 136px;
	width: 1px;
	top: 10px;
}
-->
</style>
<link href="Editor.css" rel="stylesheet" type="text/css">
<body scroll=no topmargin="0" leftmargin="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td valign="middle"> <div align="left">
        <textarea name="EditTagCode" cols="50" id="EditTagCode" style="Height:136;"></textarea>
      </div></td>
    <td width="1" valign="top">
<div align="center" class="Separator"></div></td>
    <td width="100" valign="top"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="30"> 
            <div align="center"> 
              <input type="button" onClick="SetAttribute();" name="BtnSet" value=" 设 置 ">
            </div></td>
        </tr>
        <tr> 
          <td height="30"> 
            <div align="center"> 
              <input type="button" onClick="window.close();" name="Submit" value=" 取 消 ">
            </div></td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>
<script language="JavaScript">
var EditControl=null;
var DialogSelection=dialogArguments.KS_EditArea.document.selection;
if (DialogSelection.type=='Control') EditControl=DialogSelection.createRange().item(0);
else 
{
	if (dialogArguments.EditControl!=null) EditControl=dialogArguments.EditControl;
}
if (EditControl==null)
{
	document.all.BtnSet.disabled=true;
}
else
{
	document.all.EditTagCode.value=EditControl.outerHTML;
}
function SetAttribute()
{
	EditControl.outerHTML=document.all.EditTagCode.value;
	dialogArguments.ShowTableBorders();
	window.close();
}
			window.onunload=SetReturnValue;
			function SetReturnValue()
			{
				if (typeof(window.returnValue)!='string') window.returnValue='';
			}

</script> 
