<HTML>
<HEAD>
<title>������Ŀ��</title>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<link rel="stylesheet" href="editor.css">
<script language="JavaScript">
function OK(){
    var str1="";
    str1="<FIELDSET align='"+align1.value+"' style='"
    if(color.value!='')str1=str1+"color:"+color.value+";"
    if(backcolor.value!='')str1=str1+"background-color:"+backcolor.value+";"
    str1=str1+"'><Legend"
    str1=str1+" align="+align2.value+">"+LegendTitle.value+"</Legend>��������������Ŀ�������</FIELDSET>"
    window.returnValue = str1;
    window.close();
}
</script>
<link href="ModeWindow.css" rel="stylesheet" type="text/css">
</head>
<style type="text/css">
<!--
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
-->
</style>
<body bgcolor=menu topmargin=15 leftmargin=15 >
<div align="center">
  <table border="0" cellpadding="0" cellspacing="0">
    <tr> 
      <td width="260"> <LEGEND align=left></LEGEND> <table border="0" align="center" cellpadding="0" cellspacing="3">
          <tr> 
            <td align="right"><div align="left">��Ŀ����뷽ʽ��</div></td>
            <td><select name="align1" id=align1>
                <option value="left" selected>����� 
                <option value="center">���� 
                <option value="right">�Ҷ��� </select></td>
          </tr>
          <tr> 
            <td align="right"><div align="left">��Ŀ���⣺ </div></td>
            <td><input name="LegendTitle" type="text" id="LegendTitle" size="20"></td>
          </tr>
          <tr> 
            <td align="right"><div align="left">������뷽ʽ�� </div></td>
            <td><select name="align2" id=select3>
                <option value="left" selected>����� 
                <option value="center">���� 
                <option value="right">�Ҷ��� </select></td>
          </tr>
          <tr> 
            <td align="right"><div align="left">�߿���ɫ�� </div></td>
            <td><input name="color" id=color2 onClick='p=showModalDialog("SelectColor.asp",window,"center:yes;dialogHeight:251px;dialogWidth:300px;help:no;status:no");if(p!=null){this.value=p.split("*")[0]}else{this.value=""}' size="20" maxlength="20" readonly></td>
          </tr>
          <tr> 
            <td align="right"><div align="left">������ɫ�� </div></td>
            <td><input name="backcolor" id=backcolor2 onClick='p=showModalDialog("SelectColor.asp",window,"center:yes;dialogHeight:251px;dialogWidth:300px;help:no;status:no");if(p!=null){this.value=p.split("*")[0]}else{this.value=""}' size="20" maxlength="20" readonly></td>
          </tr>
        </table></td>
      <td width=80 align="center"> 
        <input name="cmdOK" type="button" id="cmdOK" value="  ȷ��  " onClick="OK();"> 
        <br> <br> <input name="cmdCancel" type=button id="cmdCancel" onclick="window.close();" value='  ȡ��  '></td>
    </tr>
  </table>
</div>
</body>
</html> 
