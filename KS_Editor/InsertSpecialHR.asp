<HTML>
<HEAD>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<link href="Editor.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>����ˮƽ��</title>
<script language="JavaScript">
function OK(){
  var str1;
  str1="<hr color='"+color.value+"' size="+size.value+"' "+shadetype.value+" align="+align.value+" width="+width.value+">"
  window.returnValue = str1
  window.close();
}
function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
</script>
</head>
<BODY bgColor=menu topmargin=15 leftmargin=15 >
<table width=100% border="0" cellpadding="0" cellspacing="2">
  <tr><td>
<FIELDSET align=left>
<LEGEND align=left><strong>����ˮƽ�߲���</strong></LEGEND>
      <table border="0" cellpadding="0" cellspacing="3">
        <tr> 
          <td>������ɫ��
            <input name="color" id=color style='cursor:text' onclick='p=showModalDialog("SelectColor.asp",window,"center:yes;dialogHeight:320px;dialogWidth:300px;help:no;status:no");if(p!=null){this.value=p.split("*")[0]}else{this.value=""}' readonly>
          </td>
        </tr>
        <tr>
          <td>�����ֶȣ�
            <input name="size"  id=size onKeyPress="event.returnValue=IsDigit();" value="2" size="4" maxlength=3>
���������֣���Χ������1-100֮��</td>
        </tr>
        <tr> 
          <td> ҳ����룺
            <select name="align"  id=align>
              <option value="left" selected>Ĭ�϶���</option>
              <option value="left">����� </option>
              <option value="center">�ж��� </option>
              <option value="right">�Ҷ��� </option>
            </select>
            &nbsp;&nbsp;��ӰЧ����
            <select name="shadetype"  id=shadetype>
              <option value=noshade selected>�� 
              <option value=''>�� 
            </select>
          </td>
        </tr>
        <tr> 
          <td> ˮƽ��ȣ�
            <input name="width" id=width ONKEYPRESS="event.returnValue=IsDigit();" value="400" size="6" maxlength=3>
            ���������֣���Χ������1-999֮��</td>
        </tr>
      </table>
</fieldset></td>
    <td width=80 align="center"><input name="cmdOK" type="button" id="cmdOK" value="  ȷ��  " onClick="OK();">
      <br>
      <br>
      <input name="cmdCancel" type=button id="cmdCancel" onClick="window.close();" value='  ȡ��  '></td>
  </tr></table>
</body>
</html> 
