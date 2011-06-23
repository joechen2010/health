<!--#include file="../../conn.asp"-->
<!--#include file="../../ks_cls/kesion.commoncls.asp"-->
<%
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
dim ks:set ks=new publiccls
%>
<html>
<head>
<title>插入首页栏目列表参数设置</title>
<META HTTP-EQUIV="pragma" CONTENT="no-cache">
<META HTTP-EQUIV="Cache-Control" CONTENT="no-cache, must-revalidate">
<META HTTP-EQUIV="expires" CONTENT="Wed, 26 Feb 1997 08:21:57 GMT">
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language='JavaScript' src='../../KS_Inc/Jquery.js'></script>
<script language="javascript">

$(document).ready(function(){
		  $("#ChannelID").change(function(){
			$.get('../../plus/ajaxs.asp',{action:'GetClassOption',channelid:$(this).val()},function(data){
			  $("#ClassList").empty();
			  $("#ClassList").append("<option value='-1' style='color:red'>-当前栏目(通用)-</option>");
			  $("#ClassList").append("<option value='0'>-不指定栏目-</option>");
			  $("#ClassList").append(unescape(data));
			 });
		    });
})

function OK() {
    var Val;
    Val = '{$GetIndexList('+document.myform.ChannelID.value+','+document.myform.ClassList.value+','+document.myform.strType.value+','+document.myform.strHead.value+','+document.myform.strTail.value+','+document.myform.strNum.value+','+document.myform.strTitleNum.value+')}';  
    window.returnValue = Val;
    window.close();
}
</script>
<link href="Editor.css" rel="stylesheet"></head>
<body>
<form name="myform">
  <br>
  <table  width='96%' border='0'  align='center' cellpadding='2' cellspacing='1'>
<tr>
<td>
<FIELDSET align=center>
 <LEGEND align=left>首页栏目列表设置</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
 
  <tr>
    <td align="right"><div align="center">调用范围：</div></td>
    <td >
	 模型:
	<select name="ChannelID" id="ChannelID">
	<%
		If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
		Dim ModelXML,Node
		Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
		For Each Node In ModelXML.documentElement.SelectNodes("channel")
		 if Node.SelectSingleNode("@ks21").text="1" and Node.SelectSingleNode("@ks0").text<>"6" and Node.SelectSingleNode("@ks0").text<>"9" and Node.SelectSingleNode("@ks0").text<>"7" and Node.SelectSingleNode("@ks0").text<>"8" and Node.SelectSingleNode("@ks0").text<>"4" and Node.SelectSingleNode("@ks0").text<>"10" Then
		  
		  KS.echo "<option value='" &Node.SelectSingleNode("@ks0").text &"'>" & Node.SelectSingleNode("@ks1").text & "</option>"
		 End If
		next
	%>
	</select>
	栏目:
	<select name="ClassList" id="ClassList">
	 <option value="-1" style="color:red">---当前栏目通用---</option>
	 <option value="0">---不指定栏目---</option>
	<% KS.Echo KS.LoadClassOption(1)%>
	 </select></td>
  </tr>
 
   <tr >
    <td width="40%" align="right"><div align="center">列表类型：</div></td>
    <td width="60%" ><select name="strType">
      <option value="1">热门</option>
      <option value="2">最新</option>
      <option value="3">推荐</option>
      <option value="4">随机</option>
    </select></td>
  </tr>
 
 
  <tr >
    <td width="40%" align="right"><div align="center">头导航类型：</div></td>
    <td width="60%" ><input name="strHead" type="text" size="30" value=""> 支持WML语言</td>
  </tr>
  <tr >
    <td width="40%" align="right"><div align="center">尾导航类型：</div></td>
    <td width="60%" ><input name="strTail" type="text" size="30" value="<br/>"> 支持WML语言</td>
  </tr>
  <tr>
    <td align="right"><div align="center">显示记录数：</div></td>
    <td ><input name="strNum" type="text" onBlur="CheckNumber(this,'显示列表数量');" size="30" value="5"></td>
  </tr>
    <tr>
    <td align="right"><div align="center">链接标题字符：</div></td>
    <td ><input name="strTitleNum" type="text" onBlur="CheckNumber(this,'链接标题字符');" size="30" value="30"> 1个中文等于2字符</td>
  </tr>

  
</table>
</FIELDSET>
</td>
</tr>
<tr><td><div align="center"><input TYPE='button' value=' 确 定 ' onCLICK='OK()'></div></td></tr>
</table>
</form>
</body>
</html>
<%
CloseConn
Set KS=Nothing
%>
 
