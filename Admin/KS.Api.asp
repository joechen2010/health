<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<!--#include file="../api/cls_api.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
If Not KS.ReturnPowerResult(0, "KMST10002") Then          '检查是否有基本信息设置的权限
	Call KS.ReturnErr(1, "")
	Response.End
End If

Dim Action
Action = LCase(Request("action"))
Select Case Trim(Action)
	Case "save"
		Call SaveConformify
	Case Else
		Call showmain
End Select
Sub showmain()
Response.Write "<html><head><title>多系统整合接口设置</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='include/Admin_Style.css' rel='stylesheet' type='text/css'></head>" & vbCrLf
Response.Write "</head>"
Response.Write "<body bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"" scroll=no>"
Response.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
Response.Write "  <tr>"
Response.Write "    <td height=""25"" class=""topdashed"" valign='top' align='center'>"
Response.Write "      <b>多系统整合接口设置</b></td>"
Response.Write "  </tr>"
Response.Write "</TABLE>"
%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1">
<form name="myform" method="post" action="?action=save">
<tr class="tdbg">
	<td height="30" width="25%" class="clefttitle" align="right"><strong>是否开启多系统整合程序：</strong></td>
	<td>
	<input type="radio" name="API_Enable" value="false"<%
	If Not API_Enable Then Response.Write " checked"
	%>> 关闭&nbsp;&nbsp;
	<input type="radio" name="API_Enable" value="true"<%
	If API_Enable Then Response.Write " checked"
	%>> 开启
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>设置系统密钥：</strong></td>
	<td><input type="text" name="API_ConformKey" size="35" value="<%=API_ConformKey%>"> 
		<font color="red">系统整合，必须保证与其它系统设置的密钥一致。</font>
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>是否除错：</strong></td>
	<td>
	<input type="radio" name="API_Debug" value="false"<%
	If Not API_Debug Then Response.Write " checked"
	%>> 否&nbsp;&nbsp;
	<input type="radio" name="API_Debug" value="true"<%
	If API_Debug Then Response.Write " checked"
	%>> 是&nbsp;&nbsp;<font color="red">如果整合的论坛程序和科汛程序的用户数据不同步，你可以选择“是”</font>
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>整合程序的接口文件路径：</strong></td>
	<td><textarea name="API_Urls" rows="6" cols="70"><%=API_Urls%></textarea></td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>整合用户登录后转向URL：</strong></td>
	<td><input type="text" name="API_LoginUrl" size="45" value="<%=API_LoginUrl%>"> 
		<font color="red">不设置请输入“0”。</font>
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>整合用户注册后转向URL：</strong></td>
	<td><input type="text" name="API_ReguserUrl" size="45" value="<%=API_ReguserUrl%>"> 
		<font color="red">不设置请输入“0”。</font>
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>整合用户注销后转向URL：</strong></td>
	<td><input type="text" name="API_LogoutUrl" size="45" value="<%=API_LogoutUrl%>"> 
		<font color="red">不设置请输入“0”。</font>
	</td>
</tr>
</form>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>使用说明：</strong></td>
	<td><font color="blue">如果有多个程序整合，接口之间用半角"|"分隔<br />例如：http://你的论坛网址/dv_dpo.asp|http://你的网站地址/博客安装目录/oblogresponse.asp;<br />
	本系统的接口路径：<font color="red"><%=KS.GetDomain%>api/api_response.asp</font><br /></font></td>
</tr>
</table>
<script>
 function CheckForm()
 {
  document.all.myform.submit();
 }
</script>
<%
End Sub

Sub SaveConformify()
	Dim XslDoc,XslNode,Xsl_Files
	Xsl_Files = API_Path & "api.config"
	Xsl_Files = Server.MapPath(Xsl_Files)
	Set XslDoc = KS.InitialObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
	If Not XslDoc.Load(Xsl_Files) Then
		Response.Write "初始数据不存在！"
		Response.End
	Else
		Set XslNode = XslDoc.documentElement.selectSingleNode("rs:data/z:row")
		XslNode.attributes.getNamedItem("api_enable").text = Trim(Request.Form("API_Enable"))
		XslNode.attributes.getNamedItem("api_conformkey").text = ChkRequestForm("API_ConformKey")
		XslNode.attributes.getNamedItem("api_urls").text = ChkRequestForm("API_Urls")
		XslNode.attributes.getNamedItem("api_debug").text = ChkRequestForm("API_Debug")
		XslNode.attributes.getNamedItem("api_loginurl").text = ChkRequestForm("API_LoginUrl")
		XslNode.attributes.getNamedItem("api_reguserurl").text = ChkRequestForm("API_ReguserUrl")
		XslNode.attributes.getNamedItem("api_logouturl").text = ChkRequestForm("API_LogoutUrl")
		'XslNode.attributes.setNamedItem(XslDoc.createNode(2,"date","")).text = Now()
		'XslNode.appendChild(XslDoc.createNode(1,"pubDate","")).text = Now()
		XslDoc.save Xsl_Files
		Set XslNode = Nothing
	End If
	Set XslDoc = Nothing
	Response.Write ("<script>alert('恭喜您！保存设置成功。');location.href='KS.Api.asp';</script>")
End Sub
Function ChkRequestForm(reform)
	Dim strForm
	strForm = Trim(Request.Form(reform))
	If IsNull(strForm) Then
		strForm = "0"
	Else
		strForm = Replace(strForm, Chr(0), vbNullString)
		strForm = Replace(strForm, Chr(34), vbNullString)
		strForm = Replace(strForm, "'", vbNullString)
		strForm = Replace(strForm, """", vbNullString)
	End If
	If strForm = "" Then strForm = "0"
	ChkRequestForm = strForm
End Function

%>