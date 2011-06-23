<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
'-- �������������޸��Զ�����̳ϵͳApi�ӿ�
'=========================================================
Dim API_Path,API_Enable,API_ConformKey,API_Urls
Dim API_Debug,API_LoginUrl,API_ReguserUrl,API_LogoutUrl
Dim KS:Set KS=New PublicCls
API_Path = KS.Setting(3) & "API/"
LoadXslConfig()

Class API_Conformity
	Public AppID,Status,GetData,GetAppid
	Private XmlDoc,XmlHttp
	Private MessageCode,ArrUrls,SysKey,XmlPath
	
	Private Sub Class_Initialize()
		GetAppid = ""
		AppID = "KesionCMS"
		ArrUrls = Split(Trim(API_Urls),"|")
		Status = "1"
		SysKey = API_ConformKey
		MessageCode = ""
		XmlPath = API_Path & "api_user.xml"
		XmlPath = Server.MapPath(XmlPath)
		Set XmlDoc = KS.InitialObject("msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
		Set GetData = KS.InitialObject("Scripting.Dictionary")
		XmlDoc.ASYNC = False
		LoadXmlData()
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(XmlDoc) Then Set XmlDoc = Nothing
		If IsObject(GetData) Then Set GetData = Nothing
	End Sub

	Public Sub LoadXmlData()
		If Not XmlDoc.Load(XmlPath) Then
			XmlDoc.LoadXml "<?xml version=""1.0"" encoding=""gb2312""?><root/>"
		End If
		NodeValue "appID",AppID,1,False
	End Sub
	
	Public Sub NodeValue(Byval NodeName,Byval NodeText,Byval NodeType ,Byval blnEncode)
		Dim ChildNode,CreateCDATASection
		NodeName = Lcase(NodeName)
		If XmlDoc.documentElement.selectSingleNode(NodeName) is nothing Then
			Set ChildNode = XmlDoc.documentElement.appendChild(XmlDoc.createNode(1,NodeName,""))
		Else
			Set ChildNode = XmlDoc.documentElement.selectSingleNode(NodeName)
		End If
		If blnEncode = True Then
			NodeText = AnsiToUnicode(NodeText)
		End If
		If NodeType = 1 Then
			ChildNode.Text = ""
			Set CreateCDATASection = XmlDoc.createCDATASection(Replace(NodeText,"]]>","]]&gt;"))
			ChildNode.appendChild(createCDATASection)
		Else
			ChildNode.Text = NodeText
		End If
	End Sub

	Public Property Get XmlNode(Byval Str)
		If XmlDoc.documentElement.selectSingleNode(Str) is Nothing Then
			XmlNode = "Null"
		Else
			XmlNode = XmlDoc.documentElement.selectSingleNode(Str).text
		End If
	End Property

	Public Property Get GetXmlData()
		Dim GetXmlDoc
		GetXmlData = Null
		If GetAppid <> "" Then
			GetAppid = Lcase(GetAppid)
			If GetData.Exists(GetAppid) Then
				Set GetXmlData = GetData(GetAppid)
			End If
		End If
	End Property

	Public Sub SendHttpData()
		Dim i,GetXmlDoc,LoadAppid
		'On Error Resume Next
		Set Xmlhttp = KS.InitialObject("MSXML2.ServerXMLHTTP" & MsxmlVersion)
		Set GetXmlDoc = KS.InitialObject("msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
		For i = 0 to Ubound(ArrUrls)
			XmlHttp.Open "POST", Trim(ArrUrls(i)), false
			XmlHttp.SetRequestHeader "content-type", "text/xml"
			XmlHttp.Send XmlDoc
			'Response.Write strAnsi2Unicode(xmlhttp.responseBody)
    		'response.end
			If GetXmlDoc.load(XmlHttp.responseXML) Then
				LoadAppid = Lcase(GetXmlDoc.documentElement.selectSingleNode("appid").Text)
				GetData.add LoadAppid,GetXmlDoc
				Status = GetXmlDoc.documentElement.selectSingleNode("status").Text
				MessageCode = MessageCode & LoadAppid & "(" & Status &")��" & GetXmlDoc.documentElement.selectSingleNode("body/message").Text
				If Status = "1" Then '����������ʱ�˳�
					Exit For
				End If
			Else
				Status = "1"
				MessageCode = "�������ݴ���"
				Exit For
			End If
		Next
		Set GetXmlDoc = Nothing
		Set XmlHttp = Nothing
	End Sub

	Public Property Get Message()
		Message = MessageCode
	End Property
	
	Public Function SetCookie(Byval C_Syskey,Byval C_UserName,Byval C_PassWord,Byval C_SetType)
		Dim i,TempStr
		TempStr = ""
		For i = 0 to Ubound(ArrUrls)
			TempStr = TempStr & vbNewLine & "<script language=""JavaScript"" src="""&Trim(ArrUrls(i))&"?syskey="&Server.URLEncode(C_Syskey)&"&username="&Server.URLEncode(C_UserName)&"&password="&Server.URLEncode(C_PassWord)&"&savecookie="&Server.URLEncode(C_SetType)&"""></script>"
		Next
		SetCookie = TempStr
	End Function

	Public Sub PrintGetXmlData()
		Response.Clear
		Response.ContentType = "text/xml"
		Response.CharSet = "gb2312"
		Response.Expires = 0
		Response.Write "<?xml version=""1.0"" encoding=""gb2312""?>"&vbNewLine
		Response.Write GetXmlData.documentElement.XML
	End Sub

	Private Function AnsiToUnicode(ByVal str)
		Dim i, j, c, i1, i2, u, fs, f, p
		AnsiToUnicode = ""
		p = ""
		For i = 1 To Len(str)
			c = Mid(str, i, 1)
			j = AscW(c)
			If j < 0 Then
				j = j + 65536
			End If
			If j >= 0 And j <= 128 Then
				If p = "c" Then
					AnsiToUnicode = " " & AnsiToUnicode
					p = "e"
				End If
				AnsiToUnicode = AnsiToUnicode & c
			Else
				If p = "e" Then
					AnsiToUnicode = AnsiToUnicode & " "
					p = "c"
				End If
				AnsiToUnicode = AnsiToUnicode & ("&#" & j & ";")
			End If
		Next
	End Function

	Private Function strAnsi2Unicode(asContents)
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
End Class

Sub LoadXslConfig()
	Dim XslDoc,XslNode,Xsl_Files
	Xsl_Files = API_Path & "api.config"
	Xsl_Files = Server.MapPath(Xsl_Files)
	Set XslDoc = KS.InitialObject("Msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
	If Not XslDoc.Load(Xsl_Files) Then
		Response.Write "��ʼ���ݲ����ڣ�"
		Response.End
	Else
		Set XslNode = XslDoc.documentElement.selectSingleNode("rs:data/z:row")
		API_Enable		= XslNode.getAttribute("api_enable")
		API_ConformKey		= XslNode.getAttribute("api_conformkey")
		API_Urls		= XslNode.getAttribute("api_urls")
		API_Debug		= XslNode.getAttribute("api_debug")
		API_LoginUrl		= XslNode.getAttribute("api_loginurl")
		API_ReguserUrl		= XslNode.getAttribute("api_reguserurl")
		API_LogoutUrl		= XslNode.getAttribute("api_logouturl")
		Set XslNode = Nothing
	End If
	Set XslDoc = Nothing
End Sub
%>