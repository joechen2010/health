<HTML><HEAD><title>������ҳ�ļ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Editor.css" rel="stylesheet" type="text/css">
<style>
.Separator 
{
	BORDER-RIGHT: buttonhighlight 1px solid;
	FONT-SIZE: 0px;
	BORDER-LEFT: buttonshadow 1px solid;
	cursor: default;
	height: 120px;
	width: 1px;
	top: 10px;
}
</style>
<script language="JavaScript">
function OK(){
  var str1="";
  var strurl=url.value;
  if (strurl==""||strurl=="http://")
  {
  	alert("����������ҳ�ļ��ĵ�ַ��");
	url.focus();
	return false;
  }
  else
  {
  str1="<iframe src='"+url.value+"'"
  str1+=" scrolling="+scrolling.value+""
  str1+=" frameborder="+frameborder.value+""
  if(marginheight.value!='')str1+=" marginheight="+marginheight.value
  if(marginwidth.value!='')str1+=" marginwidth="+marginwidth.value
  if(width.value!='')str1+=" width="+width.value
  if(height.value!='')str1+=" height="+height.value
  str1=str1+"></iframe>"
  window.returnValue = str1
  window.close();
  }
}
function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
</script>
</head>
<body bgcolor=menu topmargin=15 leftmargin=15 >
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center"> 
      <LEGEND align=left></LEGEND> <table border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td align="right">��ҳ��ַ��</td>
          <td><input name="url" id=url value='http://' size=30></td>
        </tr>
        <tr> 
          <td height="23" align=right>��������</td>
          <td><select name="scrolling" id=scrolling>
              <option value=noshade selected>�� 
              <option value=yes>�� </select> &nbsp;&nbsp;&nbsp;�߿��ߣ� 
            <select name="frameborder" id=frameborder>
              <option value=0 selected>�� 
              <option value=1>�� </select></td>
        </tr>
        <tr> 
          <td align=right>���±߾ࣺ</td>
          <td ><input name="marginheight" id=marginheight ONKEYPRESS="event.returnValue=IsDigit();" value="0" size=8 maxlength=2> 
            &nbsp;&nbsp;���ұ߾ࣺ 
            <input name="marginwidth"  id=marginwidth ONKEYPRESS="event.returnValue=IsDigit();" value="0" size=8 maxlength=2></td>
        </tr>
        <tr> 
          <td align="right">��ҳ��ȣ�</td>
          <td ><input name="width" id=width ONKEYPRESS="event.returnValue=IsDigit();" value=500 size=8 maxlength=4> 
            &nbsp;&nbsp;��ҳ�߶ȣ� 
            <input name="height" id=height ONKEYPRESS="event.returnValue=IsDigit();" value=400 size=8 maxlength=4></td>
        </tr>
      </table></td>
    <td width=1 align="center">
<div align="center" class="Separator"></div></td>
    <td width=100 align="center" valign="top"> 
      <input name="cmdOK" type="button" id="cmdOK" value="  ȷ��  " onClick="OK();"> 
      <br> <br>
      <input name="cmdCancel" type=button id="cmdCancel" onclick="window.close();" value='  ȡ��  '></td>
  </tr>
</table>
</body>
</html>
<script language="JavaScript">
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>
 
