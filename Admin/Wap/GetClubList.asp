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
<title>插入小论坛帖子调用参数设置</title>
<META HTTP-EQUIV="pragma" CONTENT="no-cache">
<META HTTP-EQUIV="Cache-Control" CONTENT="no-cache, must-revalidate">
<META HTTP-EQUIV="expires" CONTENT="Wed, 26 Feb 1997 08:21:57 GMT">
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language='JavaScript' src='../../KS_Inc/Jquery.js'></script>
<script language="javascript">


function OK() {
    var Val;
    Val = '{$GetClubList('+document.myform.BoardID.value+','+document.myform.strType.value+','+document.myform.strHead.value+','+document.myform.strTail.value+','+document.myform.strNum.value+','+document.myform.strTitleNum.value+')}';  
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
 <LEGEND align=left>小论坛帖子调用设置</LEGEND>
 <table  width='100%' border='0'  align='center' cellpadding='2' cellspacing='1'>
 
  <tr>
    <td align="right"><div align="center">调用版面：</div></td>
    <td >

	栏目:
	<select name="BoardID" id="BoardID">
	 <option value="0">---不指定版面(全部)---</option>
	<% KS.Echo selectboard%>
	 </select></td>
  </tr>
 
   <tr >
    <td width="40%" align="right"><div align="center">列表类型：</div></td>
    <td width="60%" ><select name="strType">
      <option value="1">热门(点击数最高)</option>
      <option value="2">最新</option>
      <option value="3">精华</option>
      <option value="4">随机</option>
    </select></td>
  </tr>
 
 
  <tr >
    <td width="40%" align="right"><div align="center">头导航类型：</div></td>
    <td width="60%" ><input name="strHead" type="text" size="30" value="·"> 支持WML语言</td>
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

Function SelectBoard()
		
		 KS.LoadClubBoard()
		
		 dim str,xmls,nodes,XML,Node
         dim rs:set rs=conn.execute("select id,boardname from ks_guestboard where parentid=0 order by orderid")
		 if not rs.eof then set xml=KS.RsToXml(rs,"row",""):rs.close:set rs=nothing
		 If isobject(xml) then
		   for each node in xml.documentelement.selectnodes("row")
		   str=str & "<optgroup label=""" & node.selectsinglenode("@boardname").text &"""></optgroup>"
		        Set Xmls=Application(KS.SiteSN&"_ClubBoard")
				for each nodes in xmls.documentelement.selectnodes("row[@parentid=" & Node.SelectSingleNode("@id").text & "]")
				  if trim(request("bid"))=trim(Nodes.selectsinglenode("@id").text) then
				    str=str & "<option value=""" & Nodes.selectsinglenode("@id").text & """ selected=""selected"">--" & nodes.selectsinglenode("@boardname").text &"</option>"
				  else
				    str=str & "<option value=""" & Nodes.selectsinglenode("@id").text & """>--" & nodes.selectsinglenode("@boardname").text &"</option>"
				 end if
				next
		   next
		End If
		
		selectboard=str
	   End Function

CloseConn
Set KS=Nothing
%>
 
