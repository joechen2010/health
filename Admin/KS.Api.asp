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
If Not KS.ReturnPowerResult(0, "KMST10002") Then          '����Ƿ��л�����Ϣ���õ�Ȩ��
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
Response.Write "<html><head><title>��ϵͳ���Ͻӿ�����</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='include/Admin_Style.css' rel='stylesheet' type='text/css'></head>" & vbCrLf
Response.Write "</head>"
Response.Write "<body bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"" scroll=no>"
Response.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
Response.Write "  <tr>"
Response.Write "    <td height=""25"" class=""topdashed"" valign='top' align='center'>"
Response.Write "      <b>��ϵͳ���Ͻӿ�����</b></td>"
Response.Write "  </tr>"
Response.Write "</TABLE>"
%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1">
<form name="myform" method="post" action="?action=save">
<tr class="tdbg">
	<td height="30" width="25%" class="clefttitle" align="right"><strong>�Ƿ�����ϵͳ���ϳ���</strong></td>
	<td>
	<input type="radio" name="API_Enable" value="false"<%
	If Not API_Enable Then Response.Write " checked"
	%>> �ر�&nbsp;&nbsp;
	<input type="radio" name="API_Enable" value="true"<%
	If API_Enable Then Response.Write " checked"
	%>> ����
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>����ϵͳ��Կ��</strong></td>
	<td><input type="text" name="API_ConformKey" size="35" value="<%=API_ConformKey%>"> 
		<font color="red">ϵͳ���ϣ����뱣֤������ϵͳ���õ���Կһ�¡�</font>
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>�Ƿ����</strong></td>
	<td>
	<input type="radio" name="API_Debug" value="false"<%
	If Not API_Debug Then Response.Write " checked"
	%>> ��&nbsp;&nbsp;
	<input type="radio" name="API_Debug" value="true"<%
	If API_Debug Then Response.Write " checked"
	%>> ��&nbsp;&nbsp;<font color="red">������ϵ���̳����Ϳ�Ѵ������û����ݲ�ͬ���������ѡ���ǡ�</font>
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>���ϳ���Ľӿ��ļ�·����</strong></td>
	<td><textarea name="API_Urls" rows="6" cols="70"><%=API_Urls%></textarea></td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>�����û���¼��ת��URL��</strong></td>
	<td><input type="text" name="API_LoginUrl" size="45" value="<%=API_LoginUrl%>"> 
		<font color="red">�����������롰0����</font>
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>�����û�ע���ת��URL��</strong></td>
	<td><input type="text" name="API_ReguserUrl" size="45" value="<%=API_ReguserUrl%>"> 
		<font color="red">�����������롰0����</font>
	</td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>�����û�ע����ת��URL��</strong></td>
	<td><input type="text" name="API_LogoutUrl" size="45" value="<%=API_LogoutUrl%>"> 
		<font color="red">�����������롰0����</font>
	</td>
</tr>
</form>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>ʹ��˵����</strong></td>
	<td><font color="blue">����ж���������ϣ��ӿ�֮���ð��"|"�ָ�<br />���磺http://�����̳��ַ/dv_dpo.asp|http://�����վ��ַ/���Ͱ�װĿ¼/oblogresponse.asp;<br />
	��ϵͳ�Ľӿ�·����<font color="red"><%=KS.GetDomain%>api/api_response.asp</font><br /></font></td>
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
		Response.Write "��ʼ���ݲ����ڣ�"
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
	Response.Write ("<script>alert('��ϲ�����������óɹ���');location.href='KS.Api.asp';</script>")
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