<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>������</title>
<style type="text/css">
<!--
 SELECT  {font-family:����; font-size:9pt}
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
              <td height="22" colspan="2"><font color="#990000">����С</font></td>
            </tr>
            <tr> 
              <td width="110" height="26"><div align="center">������� 
                  <input name="RowNum" type="text" value="1" size="8">
                </div></td>
              <td width="140" height="26">������� 
                <input name="CellNum" type="text" value="1" size="8"> </td>
            </tr>
            <tr valign="middle"> 
              <td height="22" colspan="2"><font color="#990000">��񲼾�</font></td>
            </tr>
            <tr> 
              <td height="26"><div align="left">���뷽ʽ 
                  <Select name="Align">
                    <option value="">Ĭ��</option>
                    <option value="left">�����</option>
                    <option value="center">���ж���</option>
                    <option value="right">�Ҷ���</option>
                  </Select>
                </div></td>
              <td height="26"> �߿��ϸ 
                <input name="Border" type="text" value="1" size="8"></td>
            </tr>
            <tr> 
              <td height="26"><div align="left">��Ԫ����� 
                  <input name="cellpadding" type="text" value="0" size="7">
                </div></td>
              <td height="26">��Ԫ���� 
                <input name="cellspacing" type="text" value="0" size="6"></td>
            </tr>
            <tr> 
              <td height="22" colspan="2"><font color="#990000">���ߴ�</font></td>
            </tr>
            <tr> 
              <td height="26" nowrap><div align="center">
                  <input name="checkboxWidth" onClick="ChangeStatusWidth()" type="checkbox" id="checkboxWidth" value="checkbox" checked>
                  ָ�����
             <input name="Width" type="text" value="100" size="8">
              </div></td>
              <td height="26"> ��
<select name="SelectWidth">
                  <option value="pt">����</option>
                  <option value="%" selected>�ٷֱ�</option>
                </select> </td>
            </tr>
            <tr> 
              <td height="26" nowrap><div align="center">
                  <input name="checkboxHeight" onClick="ChangeStatusHeight()" type="checkbox" id="checkboxHeight" value="checkbox">
                  ָ���߶� 
                  <input name="Height" type="text" id="Height" value="" size="8" disabled=true>
              </div></td>
              <td height="26"> �� 
                <select name="SelectHeight" disabled=true>
                  <option value="pt">����</option>
                  <option value="%" selected>�ٷֱ�</option>
                </select></td>
            </tr>
          </table>
		  <hr size=1>
		  <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="30"> <div align="center"> 
                <input type="button" onClick="window.returnValue=GetTableStr();window.close();" name="Submit" value=" ȷ �� ">
                  <input onClick="window.close();" type="button" name="Submit2" value=" ȡ �� ">
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
