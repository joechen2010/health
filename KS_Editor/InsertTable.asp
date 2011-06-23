<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>插入表格</title>
<style type="text/css">
<!--
 SELECT  {font-family:宋体; font-size:9pt}
 -->
</style>
</head>
<link rel="stylesheet" href="Editor.css">
<body>
<div align="center">
  <table width="288" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td> <div align="center"> 
          <table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr> 
              <td height="22" colspan="2"><font color="#990000">表格大小</font></td>
            </tr>
            <tr> 
              <td width="110" height="26"><div align="center">表格行数 
                  <input name="RowNum" type="text" value="1" size="8">
                </div></td>
              <td width="140" height="26">表格列数 
                <input name="CellNum" type="text" value="1" size="8"> </td>
            </tr>
            <tr valign="middle"> 
              <td height="22" colspan="2"><font color="#990000">表格布局</font></td>
            </tr>
            <tr> 
              <td height="26"><div align="left">对齐方式 
                  <Select name="Align">
                    <option value="">默认</option>
                    <option value="left">左对齐</option>
                    <option value="center">居中对齐</option>
                    <option value="right">右对齐</option>
                  </Select>
                </div></td>
              <td height="26"> 边框粗细 
                <input name="Border" type="text" value="1" size="8"></td>
            </tr>
            <tr> 
              <td height="26"><div align="left">单元格填充 
                  <input name="cellpadding" type="text" value="0" size="7">
                </div></td>
              <td height="26">单元格间距 
                <input name="cellspacing" type="text" value="0" size="6"></td>
            </tr>
            <tr> 
              <td height="22" colspan="2"><font color="#990000">表格尺寸</font></td>
            </tr>
            <tr> 
              <td height="26" nowrap><div align="center">
                  <input name="checkboxWidth" onClick="ChangeStatusWidth()" type="checkbox" id="checkboxWidth" value="checkbox" checked>
                  指定宽度
             <input name="Width" type="text" value="100" size="8">
              </div></td>
              <td height="26"> 　
<select name="SelectWidth">
                  <option value="pt">像素</option>
                  <option value="%" selected>百分比</option>
                </select> </td>
            </tr>
            <tr> 
              <td height="26" nowrap><div align="center">
                  <input name="checkboxHeight" onClick="ChangeStatusHeight()" type="checkbox" id="checkboxHeight" value="checkbox">
                  指定高度 
                  <input name="Height" type="text" id="Height" value="" size="8" disabled=true>
              </div></td>
              <td height="26"> 　 
                <select name="SelectHeight" disabled=true>
                  <option value="pt">像素</option>
                  <option value="%" selected>百分比</option>
                </select></td>
            </tr>
          </table>
		  <hr size=1>
		  <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="30"> <div align="center"> 
                <input type="button" onClick="window.returnValue=GetTableStr();window.close();" name="Submit" value=" 确 定 ">
                  <input onClick="window.close();" type="button" name="Submit2" value=" 取 消 ">
                </div></td>
          </tr>
         
        </table>
        </div></td>
    </tr>
  </table>
</div>
</body>
</html>
<script language="JavaScript">
function ChangeStatusWidth()
{
if (document.all.checkboxWidth.checked==true)
 {
 document.all.Width.disabled=false;
 document.all.SelectWidth.disabled=false;
 }
else
 {
  document.all.Width.disabled=true;
 document.all.SelectWidth.disabled=true;
 }
}
function ChangeStatusHeight()
{  
 if (document.all.checkboxHeight.checked==true)
 {
 document.all.Height.disabled=false;
 document.all.SelectHeight.disabled=false;
 }
else
 {
  document.all.Height.disabled=true;
 document.all.SelectHeight.disabled=true;
 } 

}
function GetTableStr()
{  var Border='',Align='',Height='',Width='',cellspacing='',cellpadding='';  
  if (document.all.Align.value!='') Align=' align="'+document.all.Align.value+'"';
  if (document.all.Border.value!='') Border=' border="'+document.all.Border.value+'"';
  if (document.all.Width.value!='')  Width=' width="'+document.all.Width.value+document.all.SelectWidth.value+'"';
  if (document.all.Height.value!='') Height=' height="'+document.all.Height.value+document.all.SelectHeight.value+'"';
  if (document.all.cellspacing.value!='') cellspacing=' cellspacing="'+document.all.cellspacing.value+'"';
  if (document.all.cellpadding.value!='') cellpadding=' cellpadding="'+document.all.cellpadding.value+'"';
      
	var TempStr='<table'+Align+Border+Width+Height+cellspacing+cellpadding+'>\n';
	for (var i=0;i<document.all.RowNum.value;i++)
	{
		TempStr=TempStr+' <tr>\n';
		for (var j=0;j<document.all.CellNum.value;j++)
		{
			TempStr=TempStr+'  <td>&nbsp;</td>\n';
		}
		TempStr=TempStr+' </tr>\n';
	}
	TempStr=TempStr+'</table>';
	return TempStr;
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script> 
