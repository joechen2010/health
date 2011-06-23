<!--#include file="../conn.asp"-->
<html>
<head>
<title>歌曲播放列表参数设置</title>
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
 <LEGEND align=left>歌曲播放列表参数设置</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
  <tr >
    <td align="right"><div align="center">选择类别</div></td>
    <td >
	<select name="TypeID">
	 <option value='0'>-不指定任何类别-</option>
	 <option value='-1' style="color:red">-当前类别通用-</option>
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
    <td align="right"><div align="center">显示选择框</div></td>
    <td ><input name="ShowSelect" type="radio" value="true" checked>
      是
        <input type="radio" name="ShowSelect" value="false">
        否</td>
  </tr>
  <tr >
    <td align="right"><div align="center">列表属性</div></td>
    <td ><input name="type" type="radio" value="0" checked>
      最新歌曲
        <input type="radio" name="type" value="1">
        推荐歌曲
        <input type="radio" name="type" value="2">
        热点歌曲</td>
  </tr>
  <tr >
    <td width="24%" align="right"><div align="center">列出多少首歌曲</div></td>
    <td width="76%" ><input name="Num" type="text" id="Num" value="10" size="8" onBlur="CheckNumber(this,'歌曲首数');">
      首 每行显示: 
        <input name="Row" type="text" id="Row" value="2" size="6" onBlur="CheckNumber(this,'歌曲首数');">
        首</td>
  </tr>
  <tr >
    <td align="right"><div align="center">歌曲之间的行距</div></td>
    <td ><input name="RowHeight" type="text" id="RowHeight" value="25" size="8" onBlur="CheckNumber(this,'歌曲首数');">
      px</td>
  </tr>
  <tr >
    <td align="right"><div align="center">鼠标经过是否特效</div></td>
    <td ><input name="ShowMouseTX" type="radio" value="true" checked>
是
  <input type="radio" name="ShowMouseTX" value="false">
否</td>
  </tr>
  <tr >
    <td align="right"><div align="center">列出是否显示详细</div></td>
    <td ><input name="ShowDetailTF" type="radio" value="true" checked>
是
  <input type="radio" name="ShowDetailTF" value="false">
否 　显示歌曲的详细，如下载，收藏等</td>
  </tr>
</table>
</FIELDSET></td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' 确 定 ' onCLICK='OK()'></div></td></tr>
<tr>
  <td height="30"><div align="center"><span class="STYLE1">备注：此标签音乐频道通用</span></div></td>
</tr>
</table>
</form>
</body>
</html>
 
