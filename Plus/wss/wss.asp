<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../Plus/Session.asp"-->
<!--#include file="../../Plus/md5.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 4.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim Chk:Set Chk=New LoginCheckCls1
Chk.Run()
Set Chk=Nothing
Dim KS:Set KS=New PublicCls

Dim Wss_IsUsed,Wss_SiteID,Wss_PassWord,Wss_Domain,Wss_Key

Dim Action:Action = LCase(Request("action"))
LoadWssConfig
Select Case Trim(Action)
	Case "save"
		Call savewss
	Case "show"
	    Call show
	Case Else
		Call showmain
End Select

Sub show
 Response.Write "<script>window.open('http://intf.cnzz.com/user/companion/newasp_login.php?site_id=" & Wss_SiteID & "&password=" & Wss_password & "');history.back();</script>"
End Sub

Sub ShowMain

If Len(Wss_Domain)<3 Then Wss_Domain=KS.GetAutoDomain

Response.Write "<html><head><title>��ϵͳ���Ͻӿ�����</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'></head>" & vbCrLf
Response.Write "</head>"
Response.Write "<body bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"" scroll=no>"
Response.Write "<ul id='menu_top' style='text-align:center;padding-top:10px;font-weight:bold'>"
Response.Write "     WSS����ͳ������</ul>"
%>
<table border="0" align="center" cellpadding="3" cellspacing="1" width="100%" class="border">
<form name="myform" method="post" action="?action=save">
<tr class="tdbg">
	<td class="clefttitle" width="20%" align="right"><u>WSSͳ������</u>��</td>
	<td width="80%"><input type="text" name="Wss_Domain" size="35" value="<%=Wss_Domain%>"> 
		<font color="red">* </font>
	</td>
</tr>
<tr class="tdbg">
	<td class="clefttitle" align="right"><u>WSSͳ��վ��ID</u>��</td>
	<td><input type="text" name="Wss_SiteID" size="35" value="<%=Wss_SiteID%>"> 
		<font color="red">* ������Ѿ�ע���WSS���������վ��ID</font>
	</td>
</tr>
<tr class="tdbg">
	<td class="clefttitle" align="right"><u>WSSͳ�Ƶ�¼����</u>��</td>
	<td><input type="text" name="Wss_PassWord" size="35" value="<%=Wss_PassWord%>"> 
		<font color="red">* ������Ѿ�ע���WSS��������ĵ�¼����</font>
	</td>
</tr>
<tr class="tdbg">
	<td class="clefttitle" align="right"><u>�Ƿ���WSSͳ�ƹ���</u>��</td>
	<td>
	<input type="radio" name="wss_isused" value="0"<%
	If Wss_IsUsed=0 Then Response.Write " checked"
	%>> �ر�&nbsp;&nbsp;
	<input type="radio" name="wss_isused" value="1"<%
	If Wss_IsUsed=1 Then Response.Write " checked"
	%>> ����&nbsp;&nbsp;
	</td>
</tr>
<tr class="tdbg">
	<td class="clefttitle" align="right"><u>����WSSͳ��</u>��</td>
	<td class="clefttitle"><input type="checkbox" name="apply" value="1"/> 
		<font color="red">* ������ǵ�һ��������ѡ��</font>
	</td>
</tr>
<tr class="tdbg">
	<td colspan="2" align="center">
	<input type="submit" value="��������" name="B1" class="Button"></td>
</tr>
</form>
<tr>
	<td class="clefttitle" colspan="2"><b>˵��</b><br/>&nbsp;&nbsp;<a href="http://intf.cnzz.com/" target="_blank">WSS</a> һֱ�����ھ�ȷʱʵ����վ����ͳ�Ʒ���������ͨ�����ϵ�Ŭ��Ϊ����վ�ṩ�����١���ֱ�ۡ���׼ȷ��ͳ�Ʒ���<br/><br/>
	<b>����ʧ������´�����룺</b><br/>
&nbsp;&nbsp;-1 ��ʾkey����<a href="http://bbs.kesion.com" target="_blank">���ύ�����ǵ���̳</a>��,<br/>
&nbsp;&nbsp;-2 ��ʾ��������������1~64��,<br/>
&nbsp;&nbsp;-3 ��ʾ�����������󣨱������뺺�֣�,<br/>
&nbsp;&nbsp;-4 ��ʾ�����������ݿ�����<a href="http://bbs.kesion.com" target="_blank">���ύ�����ǵ���̳</a>����<br/>
&nbsp;&nbsp;-5 ��ʾͬһ��IP�û�����ҳ�泬����ֵ����ֵ�ݶ�Ϊ10��
</td>
</tr>

<tr>
	<td class="clefttitle" align="right"><u>ͳ�ƴ���</u>��</td>
	<td class="clefttitle">
	<textarea name="wsscode" rows="3" cols="70">&lt;script src='http://pw.cnzz.com/c.php?id=<%=Wss_SiteID%>&l=2' language='JavaScript' charset='gb2312'&gt;&lt;/script&gt;</textarea>
	<br><font color=red>�����ϴ��븴�Ƶ���Ҫͳ�Ƶ���ҳģ���Ｔ��</font>
	</td>
</tr>

</table>
<%

End Sub

Sub savewss()
	If Len(Request.Form("wss_domain")) < 3 Then
		response.write "<script>alert('�����������!');history.back();</script>"
	End If
	Dim XmlDoc,XmlNode,Xml_Files
	Dim apply : apply = KS.ChkClng(KS.G("apply"))
	Xml_Files = "wss.config"
	Xml_Files = Server.MapPath(Xml_Files)
	Set XmlDoc = KS.InitialObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
	If XmlDoc.Load(Xml_Files) Then
		Set XmlNode = XmlDoc.documentElement.selectSingleNode("rs:data/z:row[@id=0]")
		If apply = 0 Then
			XmlNode.attributes.getNamedItem("wss_siteid").text = KS.S("wss_siteid")
			XmlNode.attributes.getNamedItem("wss_password").text = KS.S("wss_password")
		Else
			If Len(Request.Form("wss_domain")) > 3 Then
				Dim strWssData
				Dim strURL,strDomain,strKey
				strDomain = KS.G("wss_domain")
				strKey = Md5(strDomain&"Ioi6pPdV",32)
				strURL = "http://intf.cnzz.com/user/companion/kesion.php?domain="&strDomain&"&key=" & strKey
				strWssData = GetWssData(strURL)
				If InStr(strWssData,"@") > 0 Then
					Dim WssArray
					WssArray = Split(strWssData, "@")
					XmlNode.attributes.getNamedItem("wss_siteid").text = Trim(WssArray(0))
					XmlNode.attributes.getNamedItem("wss_password").text = Trim(WssArray(1))
				Else
					Response.Write "<script>alert('����WSSʧ��!������룺" & strWssData & strKey &"');history.back();</script>"
					Exit Sub
				End If
			End If
		End If
		XmlNode.attributes.getNamedItem("wss_isused").text = KS.ChkCLng(KS.S("wss_isused"))
		XmlNode.attributes.getNamedItem("wss_domain").text = KS.G("wss_domain")
		XmlDoc.save Xml_Files
		Set XmlNode = Nothing
	End If
	Set XmlDoc = Nothing
	 Response.Write "<script>alert('��ϲ��������WSS���óɹ���');location.href='wss.asp';</script>"
End Sub
Function GetWssData(ByVal strURL)
	On Error Resume Next
	Dim xmlhttp,TextBody
	Set xmlhttp = KS.InitialObject("msxml2.ServerXMLHTTP")
	xmlhttp.setTimeouts 65000, 65000, 65000, 65000
	xmlhttp.Open "GET",strURL,false
	xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	xmlhttp.send()
	'TextBody = strAnsi2Unicode(xmlhttp.responseBody)
	TextBody = xmlhttp.responseText
	Set xmlhttp = Nothing
	GetWssData = TextBody
End Function
Function strAnsi2Unicode(asContents)
	Dim len1,i,varchar,varasc
	strAnsi2Unicode = ""
	len1=LenB(asContents)
	If len1=0 Then Exit Function
	  For i=1 to len1
	  	varchar=MidB(asContents,i,1)
	  	varasc=AscB(varchar)
	  	If varasc > 127  Then
	  		If MidB(asContents,i+1,1)<>"" Then
	  			strAnsi2Unicode = strAnsi2Unicode & chr(ascw(midb(asContents,i+1,1) & varchar))
	  		End If
	  		i=i+1
	     Else
	     	strAnsi2Unicode = strAnsi2Unicode & Chr(varasc)
	     End If	
	  Next
End Function
Sub LoadWssConfig()
Dim XmlDoc,XmlNode,Xml_Files
Xml_Files = "wss.config"
Xml_Files = Server.MapPath(Xml_Files)
Set XmlDoc = KS.InitialObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
If Not XmlDoc.Load(Xml_Files) Then
			Wss_IsUsed = 0
			Wss_SiteID = ""
			Wss_PassWord = ""
			Wss_Domain = KS.GetAutoDomain
			Wss_Key = ""
Else
			Set XmlNode	= XmlDoc.documentElement.selectSingleNode("rs:data/z:row[@id=0]")
			Wss_IsUsed = KS.ChkClng(XmlNode.getAttribute("wss_isused"))
			Wss_SiteID = XmlNode.getAttribute("wss_siteid")
			Wss_PassWord = XmlNode.getAttribute("wss_password")
			Wss_Domain = XmlNode.getAttribute("wss_domain")
			Wss_Key = XmlNode.getAttribute("wss_key")
			Set XmlNode = Nothing
End If
Set XmlDoc = Nothing
End Sub
%>