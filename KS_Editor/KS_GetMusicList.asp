<!--#include file="../conn.asp"-->
<html>
<head>
<title>���������б��������</title>
<script language="JavaScript" src="../KS_Inc/Common.js"></script>
<script language="javascript">
function OK() {
    var TypeID,Val,ShowSelect,type,ShowMouseTX,ShowDetailTF;
	
	for (var i=0;i<document.myform.ShowSelect.length;i++){
	 var KM = document.myform.ShowSelect[i];
	if (KM.checked==true)	   
		ShowSelect = KM.value
	}
	for (var i=0;i<document.myform.type.length;i++){
	 var KM = document.myform.type[i];
	if (KM.checked==true)	   
		type = KM.value
	}
	for (var i=0;i<document.myform.ShowMouseTX.length;i++){
	 var KM = document.myform.ShowMouseTX[i];
	if (KM.checked==true)	   
		ShowMouseTX = KM.value
	}
	for (var i=0;i<document.myform.ShowDetailTF.length;i++){
	 var KM = document.myform.ShowDetailTF[i];
	if (KM.checked==true)	   
		ShowDetailTF = KM.value
	}

    Val = '{=GetMusicList('+document.myform.TypeID.value+','+ShowSelect+','+type+','+document.myform.Num.value+','+document.myform.RowHeight.value+','+ShowMouseTX+','+ShowDetailTF+','+document.myform.Row.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>

<link href="Editor.css" rel="stylesheet">
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>
</head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>���������б��������</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td align="right"><div align="center">ѡ�����</div></td>
    <td >
	<select name="TypeID">
	 <option value='0'>-��ָ���κ����-</option>
	 <option value='-1' style="color:red">-��ǰ���ͨ��-</option>
	 <%
	  dim rs
	  set rs=server.createobject("adodb.recordset")
	  rs.open "select SclassID,Sclass from KS_MSSClass",conn,1,1
	  do while not rs.eof
	    response.write "<option value=""" & rs(0) & """>" & rs(1) & "</option>"
		rs.movenext
	  loop
	  rs.close
	  set rs=nothing
	  conn.close
	  set conn=nothing
	 %>
	</select>
	</td>
  </tr>
  <tr >
    <td align="right"><div align="center">��ʾѡ���</div></td>
    <td ><input name="ShowSelect" type="radio" value="true" checked>
      ��
        <input type="radio" name="ShowSelect" value="false">
        ��</td>
  </tr>
  <tr >
    <td align="right"><div align="center">�б�����</div></td>
    <td ><input name="type" type="radio" value="0" checked>
      ���¸���
        <input type="radio" name="type" value="1">
        �Ƽ�����
        <input type="radio" name="type" value="2">
        �ȵ����</td>
  </tr>
  <tr >
    <td width="24%" align="right"><div align="center">�г������׸���</div></td>
    <td width="76%" ><input name="Num" type="text" id="Num" value="10" size="8" onBlur="CheckNumber(this,'��������');">
      �� ÿ����ʾ: 
        <input name="Row" type="text" id="Row" value="2" size="6" onBlur="CheckNumber(this,'��������');">
        ��</td>
  </tr>
  <tr >
    <td align="right"><div align="center">����֮����о�</div></td>
    <td ><input name="RowHeight" type="text" id="RowHeight" value="25" size="8" onBlur="CheckNumber(this,'��������');">
      px</td>
  </tr>
  <tr >
    <td align="right"><div align="center">��꾭���Ƿ���Ч</div></td>
    <td ><input name="ShowMouseTX" type="radio" value="true" checked>
��
  <input type="radio" name="ShowMouseTX" value="false">
��</td>
  </tr>
  <tr >
    <td align="right"><div align="center">�г��Ƿ���ʾ��ϸ</div></td>
    <td ><input name="ShowDetailTF" type="radio" value="true" checked>
��
  <input type="radio" name="ShowDetailTF" value="false">
�� ����ʾ��������ϸ�������أ��ղص�</td>
  </tr>
</table>
</FIELDSET></td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' ȷ �� ' onCLICK='OK()'></div></td></tr>
<tr>
  <td height="30"><div align="center"><span class="STYLE1">��ע���˱�ǩ����Ƶ��ͨ��</span></div></td>
</tr>
</table>
</form>
</body>
</html>
 
