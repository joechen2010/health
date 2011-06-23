<!--#include file="Kesion.KeyCls.asp"-->
<% 
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
'-----------------------------------------------------------------------------------------------
'��Ѵ��վ����ϵͳ,�ɼ�ͨ����
'����:���ݿ�����Ϣ�������޹�˾ �汾 V 6.0
'-----------------------------------------------------------------------------------------------
Class CollectPublicCls
		 Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		'==================================================
		'��������GetItemLocation
		'��  �ã���ʾ����ģ���б�
		'��  ����Step ------�ڼ���,ItemID----��ĿID
		'==================================================
		Function GetItemLocation(Step, ItemID)
		 Select Case Step
		   Case 1
		   GetItemLocation="<strong>�������裺</strong>�༭��Ŀ>> <span style=""color:red"">��������</span> >> �б����� >> �������� >> �������� >> �������� >> �������� >> ���"
		   Case 2
		    GetItemLocation="<strong>�������裺</strong>�༭��Ŀ>> �������� >> <span style=""color:red"">�б�����</span> >> �������� >> �������� >> �������� >> �������� >> ���"
		   Case 3
		    GetItemLocation="<strong>�������裺</strong>�༭��Ŀ>> �������� >> �б����� >> <span style=""color:red"">��������</span> >> �������� >> �������� >> �������� >> ���"
		   Case 4
		    GetItemLocation="<strong>�������裺</strong>�༭��Ŀ>> �������� >> �б����� >> �������� >> <span style=""color:red"">��������</span> >> �������� >> �������� >> ���"
		   Case 5
		    GetItemLocation="<strong>�������裺</strong>�༭��Ŀ>> �������� >> �б����� >> �������� >> �������� >> <span style=""color:red"">��������</span> >> �������� >> ���"
		  Case 6
            GetItemLocation="<strong>�������裺</strong>�༭��Ŀ>> �������� >> �б����� >> �������� >> �������� >> �������� >> <span style=""color:red"">��������</span> >> ���"
		 End Select
		
		End Function
		'==================================================
		'��������GetTemplate
		'��  �ã���ʾ����ģ���б�
		'��  ����ChannelID ------Ƶ��ID,TemplateID----��ѡ��ģ��ID
		'==================================================
		Function GetTemplate(ChannelID, TemplateID)
			 Dim TemplateSql, TemplateRS
			TemplateSql = "Select TemplateName,TemplateID,IsDefault From KS_Template Where ChannelID=" & ChannelID & " And TemplateType=3 Order By TemplateID" ' 3��������ҳģ��
		   Set TemplateRS = conn.Execute(TemplateSql)
				If TemplateRS.EOF And TemplateRS.BOF Then
					 GetTemplate = "<option value=0>�������ģ��</option>"
				Else
					 Do While Not TemplateRS.EOF
					  If CInt(TemplateRS("TemplateID")) = CInt(TemplateID) Then
					  GetTemplate = GetTemplate & "<option value=" & TemplateRS("TemplateID") & " selected>" & TemplateRS("TemplateName") & "</option>"
					  Else
					  GetTemplate = GetTemplate & "<option value=" & TemplateRS("TemplateID") & ">" & TemplateRS("TemplateName") & "</option>"
					  End If
					 TemplateRS.MoveNext
					 Loop
			 End If
			 TemplateRS.Close:Set TemplateRS = Nothing
		 End Function
		 '==================================================
		'��������GetSpecialList
		'��  �ã���ʾƵ���µ�ר��,�������Ƶ��ʹ��
		'��  ����ChannelID ------Ƶ��ID
		'==================================================
		
		Sub GetSpecialList()
		Dim Rs, i, SpecialOpStr
		Set Rs = conn.Execute("Select * From KS_Class Where ChannelID=1 And TN='0'")
		Response.Write ("<Script language=""Javascript"">") & vbCrLf
		Response.Write "var SpecialArr = new Array();" & vbCrLf
		Do While Not Rs.EOF
		  i = i + 1
		  SpecialOpStr = "<option value='0'>---�������κ�ר��---</option>" & KS.ReturnSpecial(0)
		  Response.Write "SpecialArr[" & Rs("ID") & "] =new Array(""" & SpecialOpStr & """)" & vbCrLf
		Rs.MoveNext
		Loop
		Response.Write ("</Script>")
		Rs.Close:Set Rs = Nothing
		End Sub
		'==================================================
		'��������GetClassList
		'��  �ã���ʾƵ���µ�Ŀ¼,�������ϵͳʹ��
		'��  ����ChannelID ------Ƶ��ID
		'==================================================
		
		Sub GetClassList()
		Dim Rs:Set Rs = conn.Execute("Select * From KS_Channel Where ChannelStatus=1 And CollectTF=1")
		Response.Write ("<Script language=""Javascript"">") & vbCrLf
		Response.Write "var ClassArr = new Array();" & vbCrLf
		Do While Not Rs.EOF
		  Response.Write "ClassArr[" & Rs("ChannelID") & "] =new Array(""" & KS.LoadClassOption(Rs("ChannelID")) & """)" & vbCrLf
		Rs.MoveNext
		Loop
		Response.Write ("</Script>")
		Rs.Close
		Set Rs = Nothing
		End Sub
		
		'==================================================
		'��������Collect_ShowChannel_Option
		'��  �ã���ʾƵ��ѡ��
		'��  ����ChannelID ------Ƶ��ID
		'==================================================
		Function Collect_ShowChannel_Option(ChannelID)
		   Dim Sqlc, Rsc, ChannelName, TempStr
		   ChannelID = CLng(ChannelID)
		   Sqlc = "select ChannelID,ChannelName from KS_Channel where CollectTF=1 And ChannelStatus=1 order by ChannelID asc"
		   Set Rsc = KS.InitialObject("adodb.recordset")
		   Rsc.Open Sqlc, conn, 1, 1
		   TempStr = "<option value=""0"" selected>---��ѡ��ϵͳģ��---</option>"
		   If Rsc.EOF And Rsc.BOF Then
			  TempStr = TempStr & "<option value=""0"">-------</option>"
		   Else
			  Do While Not Rsc.EOF
				 TempStr = TempStr & "<option value=" & """" & Rsc("ChannelID") & """" & ""
				 If ChannelID = Rsc("ChannelID") Then
					TempStr = TempStr & " selected"
				 End If
				 TempStr = TempStr & ">" & Rsc("ChannelName")
				 TempStr = TempStr & "</option>"
			  Rsc.MoveNext
			  Loop
		   End If
		   Rsc.Close:Set Rsc = Nothing
		   Collect_ShowChannel_Option = TempStr
		End Function
		
		'==================================================
		'��������Collect_ShowClass_Name
		'��  �ã���ʾ��Ŀ����
		'��  ����ChannelID ------Ƶ��ID
		'��  ����ClassID ------��ĿID
		'==================================================
		Function Collect_ShowClass_Name(ChannelID, ClassID)
		   Dim Sqlc, Rsc, TempStr
		   Set Rsc = conn.execute("Select top 1 FolderName from KS_Class Where ChannelID=" & ChannelID & " and ID='" & ClassID & "'")
		   If Rsc.EOF And Rsc.BOF Then
			  TempStr = "��ָ����Ŀ"
		   Else
			  TempStr = Rsc("FolderName")
		   End If
		   Rsc.Close
		   Set Rsc = Nothing
		   Collect_ShowClass_Name = TempStr
		End Function
		
		'==================================================
		'��������Collect_ShowSpecial_Name
		'��  �ã���ʾר������
		'��  ����ChannelID ------Ƶ��ID
		'��  ����SpecialID ------ר��ID
		'==================================================
		Sub Collect_ShowSpecial_Name(ChannelID, SpecialID)
		   Dim Sqlc, Rsc, TempStr
		   ChannelID = CLng(ChannelID)
		   Sqlc = "select top 1 SpecialName from KS_Special Where ChannelID=" & ChannelID & " and ID='" & SpecialID & "'"
		   Set Rsc = KS.InitialObject("adodb.recordset")
		   Rsc.Open Sqlc, conn, 1, 1
		   If Rsc.EOF And Rsc.BOF Then
			  TempStr = "��ָ��ר��"
		   Else
			  TempStr = Rsc("SpecialName")
		   End If
		   Rsc.Close
		   Set Rsc = Nothing
		   Response.Write TempStr
		End Sub
		
		'==================================================
		'��������Collect_ShowClass_Option
		'��  �ã���ʾ��Ŀѡ��
		'��  ����ChannelID ------Ƶ��ID
		'��  ����ClassID ------��ĿID
		'==================================================
		Sub Collect_ShowClass_Option(ChannelID, ClassID)
			Dim rsClass, sqlClass, strTempC, tmpTJ, i
			Dim arrShowLine(20)
			ChannelID = CLng(ChannelID)
			ClassID = ClassID
			For i = 0 To UBound(arrShowLine)
				arrShowLine(i) = False
			Next
				strTempC = ""
			sqlClass = "Select * From KS_Class where channelid=" & ChannelID & " order by OrderID"
			Set rsClass = conn.Execute(sqlClass)
			If rsClass.BOF And rsClass.EOF Then
				strTempC = "<option value=''>���������Ŀ</option>"
			Else
						Do While Not rsClass.EOF
								tmpTJ = rsClass("TJ")
					If rsClass("NextID") > 0 Then
						arrShowLine(tmpTJ) = True
					Else
						arrShowLine(tmpTJ) = False
					End If
					strTempC = strTempC & "<option value='" & rsClass("ClassID") & "'"
						If rsClass("ID") = ClassID Then
							strTempC = strTempC & ""
						End If
					strTempC = strTempC & rsClass("FolderName")
					strTempC = strTempC & "</option>"
				rsClass.MoveNext
				Loop
			End If
			rsClass.Close
			Set rsClass = Nothing
			Response.Write strTempC
		End Sub
		
		'==================================================
		'��������Collect_ShowSpecial_Option
		'��  �ã���ʾר��ѡ��
		'��  ����ChannelID ------Ƶ��ID
		'��  ����SpecialID ------ר��ID
		'==================================================
		Sub Collect_ShowSpecial_Option(ChannelID, SpecialID)
			ChannelID = CLng(ChannelID)
			SpecialID = SpecialID
			Dim TempStr
			TempStr = "<select name='SpecialID' id='SpecialID'><option value=''"
			If SpecialID = 0 Then
				TempStr = TempStr & " selected"
			End If
			TempStr = TempStr & ">�������κ�ר��</option>"
							
			Dim sqlSpecial, rsSpecial
				sqlSpecial = "select * from KS_Special where ChannelID=" & ChannelID
			Set rsSpecial = KS.InitialObject("adodb.recordset")
			rsSpecial.Open sqlSpecial, conn, 1, 1
			Do While Not rsSpecial.EOF
				If rsSpecial("ID") = SpecialID Then
					TempStr = TempStr & "<option value='" & rsSpecial("ID") & "' selected>" & rsSpecial("SpecialName") & "</option>"
				Else
					TempStr = TempStr & "<option value='" & rsSpecial("ID") & "'>" & rsSpecial("SpecialName") & "</option>"
				End If
			rsSpecial.MoveNext
			Loop
			rsSpecial.Close
				Set rsSpecial = Nothing
				Response.Write TempStr
		End Sub
				
		
		'==================================================
		'��������Collect_ShowItem_Name
		'��  �ã���ʾ��Ŀ����
		'��  ����ItemID ------��ĿID
		'==================================================
		Function Collect_ShowItem_Name(ItemID, ConnItem)
		   Dim Sqlc, Rsc, TempStr
		   ItemID = CLng(ItemID)
		   Sqlc = "select top 1 ItemName From KS_CollectItem Where ItemID=" & ItemID
		   Set Rsc = KS.InitialObject("adodb.recordset")
		   Rsc.Open Sqlc, ConnItem, 1, 1
		   If Rsc.EOF And Rsc.BOF Then
			  TempStr = "��ָ����Ŀ"
		   Else
			  TempStr = Rsc("ItemName")
		   End If
		   Rsc.Close
		   Set Rsc = Nothing
		   Collect_ShowItem_Name = TempStr
		End Function
		
		
		'==================================================
		'��������Collect_ShowItem_Option
		'��  �ã���ʾ��Ŀѡ��
		'��  ����ItemID ------��ĿID
		'==================================================
		Function Collect_ShowItem_Option(ItemID, ConnItem)
		   Dim SqlI, RsI, TempStr
		   ItemID = CLng(ItemID)
		   SqlI = "select ItemID,ItemName From KS_CollectItem order by ItemID desc"
		   Set RsI = KS.InitialObject("adodb.recordset")
		   RsI.Open SqlI, ConnItem, 1, 1
		   TempStr = "<select Name=""ItemID"" ID=""ItemID"">"
		   If RsI.EOF And RsI.BOF Then
			  TempStr = TempStr & "<option value="""">�������Ŀ</option>"
		   Else
			  TempStr = TempStr & "<option value="""">��ѡ����Ŀ</option>"
			  Do While Not RsI.EOF
				 TempStr = TempStr & "<option value=" & """" & RsI("ItemID") & """" & ""
				 If ItemID = RsI("ItemID") Then
					TempStr = TempStr & " Selected"
				 End If
				 TempStr = TempStr & ">" & RsI("ItemName")
				 TempStr = TempStr & "</option>"
			  RsI.MoveNext
			  Loop
		   End If
		   RsI.Close
		   Set RsI = Nothing
		   TempStr = TempStr & "</select>"
		   Collect_ShowItem_Option = TempStr
		End Function
		'==================================================
		'��������SplitNewsPage
		'��  �ã���ȡ�Զ���ҳ
		'��  ����Content--���� MaxPerChar--ÿҳ����ַ�
		'==================================================
		Function SplitNewsPage(Content,MaxPerChar)
		      SplitNewsPage=KS.AutoSplitPage(Content,"[NextPage]",KS.ChkClng(MaxPerChar))
		End Function
		'==================================================
		'��������GetHttpPage
		'��  �ã���ȡ��ҳԴ��
		'��  ����HttpUrl ------��ҳ��ַ
		'==================================================
		Function GetHttpPage(HttpUrl,CharsetCode)
		   If IsNull(HttpUrl) = True Or Len(HttpUrl) < 18 Or HttpUrl = "Error" Then
			  GetHttpPage = "Error"
			  Exit Function
		   End If
		   Dim Http
		  Set Http = KS.InitialObject("MSXML2.ServerXMLHTTP") 
		   Http.Open "GET", HttpUrl, False
		  on error resume next
		   Http.Send
		   If Http.Readystate <> 4 Then
		  'If Http.Status<>200 then 
			  Set Http = Nothing
			  GetHttpPage = "Error"
			  Exit Function
		   End If
		   if CharsetCode="auto" Or CharsetCode="" then CharsetCode=GetEncodeing(HttpUrl)
		   GetHttpPage = BytesToBstr(Http.ResponseBody, CharsetCode)
		   Set Http = Nothing
		   If Err.Number <> 0 Then
			  Err.Clear
		   End If
		End Function
		'�Զ�ȡ�ñ����ʽ
		function GetEncodeing(sUrl)
		On Error Resume Next
		dim http,re,encodeing
		Set http=KS.InitialObject("Microsoft.XMLHTTP")
		http.Open "GET",sUrl,False
		http.send
		if http.status="200" then
			Set re=new RegExp
			re.IgnoreCase =True
			re.Global=True
			re.Pattern="encoding=\""gb2312"
			if re.test(http.responseText) then
				encodeing="gb2312"
			else
				encodeing="utf-8"
			end if
			set re=nothing
		end if
		If Err Then
			Err.Clear
			GetEncodeing="utf-8"
		else
			GetEncodeing=encodeing
		End If
		set http=nothing
	end function
		'==================================================
		'��������BytesToBstr
		'��  �ã�����ȡ��Դ��ת��Ϊ����
		'��  ����Body ------Ҫת���ı���
		'��  ����Cset ------Ҫת��������
		'==================================================
		Function BytesToBstr(Body, Cset)
		   Dim Objstream
		   Set Objstream = KS.InitialObject("adodb.stream")
		   Objstream.Type = 1
		   Objstream.Mode = 3
		   Objstream.Open
		   Objstream.Write Body
		   Objstream.Position = 0
		   Objstream.Type = 2
		   Objstream.Charset = Cset
		   BytesToBstr = Objstream.ReadText
		   Objstream.Close
		   Set Objstream = Nothing
		End Function
		
		'==================================================
		'��������PostHttpPage
		'��  �ã���¼
		'==================================================
		Function PostHttpPage(RefererUrl, PostUrl, PostData)
			Dim xmlHttp
			Dim RetStr
			Set xmlHttp = KS.InitialObject("Msxml2.XMLHTTP")
			xmlHttp.Open "POST", PostUrl, False
			xmlHttp.setRequestHeader "Content-Length", Len(PostData)
			xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
			xmlHttp.setRequestHeader "Referer", RefererUrl
			xmlHttp.Send PostData
			If Err.Number <> 0 Then
				Set xmlHttp = Nothing
				PostHttpPage = "Error"
				Exit Function
			End If
			PostHttpPage = BytesToBstr(xmlHttp.ResponseBody, "GB2312")
			Set xmlHttp = Nothing
		End Function
		
		'==================================================
		'��������UrlEncoding
		'��  �ã�ת������
		'==================================================
		Function UrlEncoding(DataStr)
			Dim StrReturn, Si, ThisChr, InnerCode, Hight8, Low8
			StrReturn = ""
			For Si = 1 To Len(DataStr)
				ThisChr = Mid(DataStr, Si, 1)
				If Abs(Asc(ThisChr)) < &HFF Then
					StrReturn = StrReturn & ThisChr
				Else
					InnerCode = Asc(ThisChr)
					If InnerCode < 0 Then
					   InnerCode = InnerCode + &H10000
					End If
					Hight8 = (InnerCode And &HFF00) \ &HFF
					Low8 = InnerCode And &HFF
					StrReturn = StrReturn & "%" & Hex(Hight8) & "%" & Hex(Low8)
				End If
			Next
			UrlEncoding = StrReturn
		End Function
		
		'==================================================
		'��������GetBody
		'��  �ã���ȡ�ַ���
		'��  ����ConStr ------��Ҫ��ȡ���ַ���
		'��  ����StartStr ------��ʼ�ַ���
		'��  ����OverStr ------�����ַ���
		'��  ����IncluL ------�Ƿ����StartStr
		'��  ����IncluR ------�Ƿ����OverStr
		'==================================================
		Function GetBody(Constr, StartStr, OverStr, IncluL, IncluR)
		   If Constr = "Error" Or Constr = "" Or IsNull(Constr) = True Or StartStr = "" Or IsNull(StartStr) = True Or OverStr = "" Or IsNull(OverStr) = True Then
			  GetBody = "Error"
			  Exit Function
		   End If
		   Dim ConstrTemp
		   Dim Start, Over
		   ConstrTemp = LCase(Constr)
		   StartStr = LCase(StartStr)
		   OverStr = LCase(OverStr)
		   Start = InStrB(1, ConstrTemp, StartStr, vbBinaryCompare)
		   If Start <= 0 Then
			  GetBody = "Error"
			  Exit Function
		   Else
			  If IncluL = False Then
				 Start = Start + LenB(StartStr)
			  End If
		   End If
		   Over = InStrB(Start, ConstrTemp, OverStr, vbBinaryCompare)
		   If Over <= 0 Or Over <= Start Then
			  GetBody = "Error"
			  Exit Function
		   Else
			  If IncluR = True Then
				 Over = Over + LenB(OverStr)
			  End If
		   End If
		   GetBody = MidB(Constr, Start, Over - Start)
		End Function
		
		
		'==================================================
		'��������GetArray
		'��  �ã���ȡ���ӵ�ַ����$Array$�ָ�
		'��  ����ConStr ------��ȡ��ַ��ԭ�ַ�
		'��  ����StartStr ------��ʼ�ַ���
		'��  ����OverStr ------�����ַ���
		'��  ����IncluL ------�Ƿ����StartStr
		'��  ����IncluR ------�Ƿ����OverStr
		'==================================================
		Function GetArray(Byval Constr, StartStr, OverStr, IncluL, IncluR)
		   If Constr = "Error" Or Constr = "" Or IsNull(Constr) = True Or StartStr = "" Or OverStr = "" Or IsNull(StartStr) = True Or IsNull(OverStr) = True Then
			  GetArray = "Error"
			  Exit Function
		   End If
		   Dim TempStr, TempStr2, objRegExp, Matches, Match
		   TempStr = ""
		   Set objRegExp = New RegExp
		   objRegExp.IgnoreCase = True
		   objRegExp.Global = True
		   objRegExp.Pattern = "(" & StartStr & ").+?(" & OverStr & ")"
		   Set Matches = objRegExp.Execute(Constr)
		   For Each Match In Matches
			  TempStr = TempStr & "$Array$" & Match.value
		   Next
		   Set Matches = Nothing
		
		   If TempStr = "" Then
			  GetArray = "Error"
			  Exit Function
		   End If
		   TempStr = Right(TempStr, Len(TempStr) - 7)
		   If IncluL = False Then
			  objRegExp.Pattern = StartStr
			  TempStr = objRegExp.Replace(TempStr, "")
		   End If
		   If IncluR = False Then
			  objRegExp.Pattern = OverStr
			  TempStr = objRegExp.Replace(TempStr, "")
		   End If
		   Set objRegExp = Nothing
		   Set Matches = Nothing
		   
		   'TempStr = Replace(TempStr, """", "")
		   'TempStr = Replace(TempStr, "'", "")
		  ' TempStr = Replace(TempStr, " ", "")
		   'TempStr = Replace(TempStr, "(", "")
		   'TempStr = Replace(TempStr, ")", "")
		
		   If TempStr = "" Then
			  GetArray = "Error"
		   Else
			  GetArray = TempStr
		   End If
		End Function
		
		
		'==================================================
		'��������DefiniteUrl
		'��  �ã�����Ե�ַת��Ϊ���Ե�ַ
		'��  ����PrimitiveUrlStr ------Ҫת������Ե�ַ
		'��  ����ConsultUrlStr ------��ǰ��ҳ��ַ
		'==================================================
		'Function DefiniteUrl(ByVal PrimitiveUrlStr, ByVal ConsultUrlStr)
		Function DefiniteUrl(ByVal URL,ByVal CurrentUrl)
		Dim strUrl
		If Len(URL) < 2 Or Len(URL) > 255 Or Len(CurrentUrl) < 2 Then
			DefiniteUrl = vbNullString
			Exit Function
		End If
		CurrentUrl = Trim(Replace(Replace(Replace(Replace(CurrentUrl, "'", vbNullString), """", vbNullString), vbNewLine, vbNullString), "\", "/"))
		URL = Trim(Replace(Replace(Replace(Replace(URL, "'", vbNullString), """", vbNullString), vbNewLine, vbNullString), "\", "/"))	
		If InStr(9, CurrentUrl, "/") = 0 Then
			strUrl = CurrentUrl
		Else
			strUrl = Left(CurrentUrl, InStr(9, CurrentUrl, "/") - 1)
		End If

		If strUrl = vbNullString Then strUrl = CurrentUrl
		Select Case Left(LCase(URL), 6)
			Case "http:/", "https:", "ftp://", "rtsp:/", "mms://"
				DefiniteUrl = URL
				Exit Function
		End Select

		If Left(URL, 1) = "/" Then
			DefiniteUrl = strUrl & URL
			Exit Function
		End If
		
		If Left(URL, 3) = "../" Then
			Dim ArrayUrl
			Dim ArrayCurrentUrl
			Dim ArrayTemp()
			Dim strTemp
			Dim i, n
			Dim c, l
			n = 0
			ArrayCurrentUrl = Split(CurrentUrl, "/")
			ArrayUrl = Split(URL, "../")
			c = UBound(ArrayCurrentUrl)
			l = UBound(ArrayUrl) + 1
			
			If c > l + 2 Then
				For i = 0 To c - l
					ReDim Preserve ArrayTemp(n)
					ArrayTemp(n) = ArrayCurrentUrl(i)
					n = n + 1
				Next
				strTemp = Join(ArrayTemp, "/")
			Else
				strTemp = strUrl
			End If
			URL = Replace(URL, "../", vbNullString)
			DefiniteUrl = strTemp & "/" & URL
			Exit Function
		End If
		strUrl = Left(CurrentUrl, InStrRev(CurrentUrl, "/"))
		DefiniteUrl = strUrl & Replace(URL, "./", vbNullString)
		Exit Function
		   
		End Function
		
		'==================================================
		'��������ReplaceSaveRemoteFile
		'��  �ã��滻������Զ��ͼƬ
		'��  ����ConStr ------ Ҫ�滻���ַ���
		'��  ����SaveTf ------ �Ƿ񱣴��ļ���False�����棬True����
		'��  ��: TistUrl------ ��ǰ��ҳ��ַ
		'==================================================
		Function ReplaceSaveRemoteFile(UploadFiles, Constr, strInstallDir, strChannelDir, SaveTf, TistUrl)
		   If Constr = "Error" Or Constr = "" Or strInstallDir = "" Or strChannelDir = "" Then
			  ReplaceSaveRemoteFile = Constr
			  Exit Function
		   End If
		   Dim TempStr, TempStr2, TempStr3, re, Matches, Match, Tempi, TempArray, TempArray2
		
		   Set re = New RegExp
		   re.IgnoreCase = True
		   re.Global = True
		   re.Pattern = "<img.+?[^\>]>"
		   Set Matches = re.Execute(Constr)
		   For Each Match In Matches
			  If TempStr <> "" Then
				 TempStr = TempStr & "$Array$" & Match.value
			  Else
				 TempStr = Match.value
			  End If
		   Next
		   If TempStr <> "" Then
			  TempArray = Split(TempStr, "$Array$")
			  TempStr = ""
			  For Tempi = 0 To UBound(TempArray)
				 re.Pattern = "src\s*=\s*.+?\.(gif|jpg|bmp|jpeg|psd|png|svg|dxf|wmf|tiff)"
				 Set Matches = re.Execute(TempArray(Tempi))
				 For Each Match In Matches
					If TempStr <> "" Then
					   TempStr = TempStr & "$Array$" & Match.value
					Else
					   TempStr = Match.value
					End If
				 Next
			  Next
		   End If
		   If TempStr <> "" Then
			  re.Pattern = "src\s*=\s*"
			  TempStr = re.Replace(TempStr, "")
		   End If
		   Set Matches = Nothing
		   Set re = Nothing
		   If TempStr = "" Or IsNull(TempStr) = True Then
			  ReplaceSaveRemoteFile = Constr
			  Exit Function
		   End If
		   TempStr = Replace(TempStr, """", "")
		   TempStr = Replace(TempStr, "'", "")
		   TempStr = Replace(TempStr, " ", "")
		
		   Dim RemoteFileUrl, SavePath, PathTemp, DtNow, strFileName, strFileType, ArrSaveFileName, RanNum, Arr_Path
		   DtNow = Now()
		   If SaveTf = True Then
		     '***********************************
				 SavePath = KS.GetUpFilesDir & "/"
			  'Response.Write "����·����" & savepath & "<br>"
			  Arr_Path = Split(SavePath, "/")
			  PathTemp = ""
			  For Tempi = 0 To UBound(Arr_Path)
				 If Tempi = 0 Then
					PathTemp = Arr_Path(0) & "/"
				 ElseIf Tempi = UBound(Arr_Path) Then
					Exit For
				 Else
					PathTemp = PathTemp & Arr_Path(Tempi) & "/"
				 End If
				 If CheckDir(PathTemp) = False Then
					If MakeNewsDir(PathTemp) = False Then
					   SaveTf = False
					   Exit For
					End If
				 End If
			  Next
		   End If
		
		   'ȥ���ظ�ͼƬ��ʼ
		   TempArray = Split(TempStr, "$Array$")
		   TempStr = ""
		   For Tempi = 0 To UBound(TempArray)
			  If InStr(LCase(TempStr), LCase(TempArray(Tempi))) < 1 Then
				 TempStr = TempStr & "$Array$" & TempArray(Tempi)
			  End If
		   Next
		   TempStr = Right(TempStr, Len(TempStr) - 7)
		   TempArray = Split(TempStr, "$Array$")
		   'ȥ���ظ�ͼƬ����
		
		   'ת�����ͼƬ��ַ��ʼ
		   TempStr = ""
		   For Tempi = 0 To UBound(TempArray)
			  TempStr = TempStr & "$Array$" & DefiniteUrl(TempArray(Tempi), TistUrl)
		   Next
		   TempStr = Right(TempStr, Len(TempStr) - 7)
		   TempStr = Replace(TempStr, Chr(0), "")
		   TempArray2 = Split(TempStr, "$Array$")
		   TempStr = ""
		   'ת�����ͼƬ��ַ����
		
		   'ͼƬ�滻/����
		   Set re = New RegExp
		   re.IgnoreCase = True
		   re.Global = True
		
		   For Tempi = 0 To UBound(TempArray2)
			  RemoteFileUrl = TempArray2(Tempi)
			  If RemoteFileUrl <> "Error" And SaveTf = True Then '����ͼƬ
				 ArrSaveFileName = Split(RemoteFileUrl, ".")
			 strFileType = LCase(ArrSaveFileName(UBound(ArrSaveFileName))) '�ļ�����
				 If strFileType = "asp" Or strFileType = "asa" Or strFileType = "aspx" Or strFileType = "cer" Or strFileType = "cdx" Or strFileType = "exe" Or strFileType = "rar" Or strFileType = "zip" Then
					UploadFiles = ""
					ReplaceSaveRemoteFile = Constr
					Exit Function
				 End If
		
				 Randomize
				 RanNum = Int(900 * Rnd) + 100
			 strFileName = Year(DtNow) & Right("0" & Month(DtNow), 2) & Right("0" & Day(DtNow), 2) & Right("0" & Hour(DtNow), 2) & Right("0" & Minute(DtNow), 2) & Right("0" & Second(DtNow), 2) & RanNum & "." & strFileType
				 re.Pattern = replace(replace(TempArray(Tempi),"(","\("),")","\)")
			     If SaveRemoteFile(SavePath & strFileName, RemoteFileUrl) = True Then
		         '********************************
					PathTemp =KS.Setting(2) & SavePath & strFileName
					Constr = re.Replace(Constr, PathTemp)
					re.Pattern = strInstallDir & strChannelDir & "/"
					UploadFiles = UploadFiles & "|" & re.Replace(SavePath & strFileName, "")
				 Else
					PathTemp = RemoteFileUrl
					Constr = re.Replace(Constr, PathTemp)
					'UploadFiles=UploadFiles & "|" & RemoteFileUrl
				 End If
				 
			  ElseIf RemoteFileUrl <> "Error" And SaveTf = False Then '������ͼƬ
				 re.Pattern = replace(replace(TempArray(Tempi),"(","\("),")","\)")
				 Constr = re.Replace(Constr, RemoteFileUrl)
				 UploadFiles = UploadFiles & "|" & RemoteFileUrl
			  End If
		   Next
		   Set re = Nothing
		   If UploadFiles <> "" Then
			  UploadFiles = Right(UploadFiles, Len(UploadFiles) - 1)
		   End If
		   ReplaceSaveRemoteFile = Constr
		End Function
		
		'==================================================
		'��������ReplaceSwfFile
		'��  �ã���������·��
		'��  ����ConStr ------ Ҫ�滻���ַ���
		'��  ��: TistUrl------ ��ǰ��ҳ��ַ
		'==================================================
		Function ReplaceSwfFile(Constr, TistUrl)
		   Dim RemoteFileUrl
		   If Constr = "Error" Or Constr = "" Or TistUrl = "" Or TistUrl = "Error" Then
			  ReplaceSwfFile = Constr: Exit Function
		   End If
		
		   Dim TempStr, TempStr2, TempStr3, re, Matches, Match, Tempi, TempArray, TempArray2
		
		   Set re = New RegExp
		   re.IgnoreCase = True
		   re.Global = True
		   re.Pattern = "<object.+?[^\>]>"
		   Set Matches = re.Execute(Constr)
		   For Each Match In Matches
			  If TempStr <> "" Then
				 TempStr = TempStr & "$Array$" & Match.value
			  Else
				 TempStr = Match.value
			  End If
		   Next
		   If TempStr <> "" Then
			  TempArray = Split(TempStr, "$Array$")
			  TempStr = ""
			  For Tempi = 0 To UBound(TempArray)
				 re.Pattern = "value\s*=\s*.+?\.swf"
				 Set Matches = re.Execute(TempArray(Tempi))
				 For Each Match In Matches
					If TempStr <> "" Then
					   TempStr = TempStr & "$Array$" & Match.value
					Else
					   TempStr = Match.value
					End If
				 Next
			  Next
		   End If
		   If TempStr <> "" Then
			  re.Pattern = "value\s*=\s*"
			  TempStr = re.Replace(TempStr, "")
		   End If
		   If TempStr = "" Or IsNull(TempStr) = True Then
			  ReplaceSwfFile = Constr
			  Exit Function
		   End If
		   TempStr = Replace(TempStr, """", "")
		   TempStr = Replace(TempStr, "'", "")
		   TempStr = Replace(TempStr, " ", "")
		
		   Set Matches = Nothing
		   Set re = Nothing
		
		   'ȥ���ظ��ļ���ʼ
		   TempArray = Split(TempStr, "$Array$")
		   TempStr = ""
		   For Tempi = 0 To UBound(TempArray)
			  If InStr(LCase(TempStr), LCase(TempArray(Tempi))) < 1 Then
				 TempStr = TempStr & "$Array$" & TempArray(Tempi)
			  End If
		   Next
		   TempStr = Right(TempStr, Len(TempStr) - 7)
		   TempArray = Split(TempStr, "$Array$")
		   'ȥ���ظ��ļ�����
		
		   'ת����Ե�ַ��ʼ
		   TempStr = ""
		   For Tempi = 0 To UBound(TempArray)
			  TempStr = TempStr & "$Array$" & DefiniteUrl(TempArray(Tempi), TistUrl)
		   Next
		   TempStr = Right(TempStr, Len(TempStr) - 7)
		   TempStr = Replace(TempStr, Chr(0), "")
		   TempArray2 = Split(TempStr, "$Array$")
		   TempStr = ""
		   'ת����Ե�ַ����
		
		   '�滻
		   Set re = New RegExp
		   re.IgnoreCase = True
		   re.Global = True
		   For Tempi = 0 To UBound(TempArray2)
			  RemoteFileUrl = TempArray2(Tempi)
			  re.Pattern = TempArray(Tempi)
			  Constr = re.Replace(Constr, RemoteFileUrl)
		   Next
		   Set re = Nothing
		   ReplaceSwfFile = Constr
		End Function
		
		'==================================================
		'��������SaveRemoteFile
		'��  �ã�����Զ�̵��ļ�������
		'��  ����LocalFileName ------ �����ļ���
		'��  ����RemoteFileUrl ------ Զ���ļ�URL
		'==================================================
		Function SaveRemoteFile(LocalFileName, RemoteFileUrl)
		    On Error Resume Next
	         SaveRemoteFile=True
			dim Ads,Retrieval,GetRemoteData
			Set Retrieval = KS.InitialObject("Microsoft.XMLHTTP")
			With Retrieval
				.Open "Get", RemoteFileUrl, False, "", ""
				.Send
				If .Readystate<>4 then
					SaveRemoteFile=False
					Exit Function
				End If
				GetRemoteData = .ResponseBody
			End With
			Set Retrieval = Nothing
			Set Ads = KS.InitialObject("Adodb.Stream")
			With Ads
				.Type = 1
				.Open
				.Write GetRemoteData
				.SaveToFile server.MapPath(LocalFileName),2
				.Cancel()
				.Close()
			End With
			Set Ads=nothing
		End Function
		
		'==================================================
		'��������FpHtmlEnCode
		'��  �ã��������
		'��  ����fString ------�ַ���
		'==================================================
		Function FpHtmlEnCode(fString)
		   If IsNull(fString) = False Or fString <> "" Or fString <> "Error" Then
			   fString = nohtml(fString)
			   fString = FilterJS(fString)
			   fString = Replace(fString, "&nbsp;", " ")
			   fString = Replace(fString, "&quot;", "")
			   fString = Replace(fString, "&#39;", "")
			   fString = Replace(fString, ">", "")
			   fString = Replace(fString, "<", "")
			   fString = Replace(fString, Chr(9), " ") '&nbsp;
			   fString = Replace(fString, Chr(10), "")
			   fString = Replace(fString, Chr(13), "")
			   fString = Replace(fString, Chr(34), "")
			   fString = Replace(fString, Chr(32), " ") 'space
			   fString = Replace(fString, Chr(39), "")
			   fString = Replace(fString, Chr(10) & Chr(10), "")
			   fString = Replace(fString, Chr(10) & Chr(13), "")
			   fString = Trim(fString)
			   FpHtmlEnCode = ReplaceChar(fString)
		   Else
			   FpHtmlEnCode = "Error"
		   End If
		End Function
		Function ReplaceChar(Content)
			Content=Replace(Replace(Content,"[",""),"]","")
			Content=Replace(Replace(Content,"��",""),"��","")
			Content=Replace(Replace(Content,"(",""),")","")
			Content=Replace(Replace(Content,"��",""),"��","")
			Content=Replace(Replace(Content,"��",""),"��","")
			Content=Replace(Replace(Content,"{",""),"}","")
			Content=Replace(Replace(Content,"'",""),"""","")
			Content=Replace(Replace(Content,"?",""),""="","")
			Content=Replace(Replace(Content,":",""),"��","")
			Content=Replace(Replace(Content,";",""),"��","")
			Content=Replace(Replace(Content,"/",""),"��","")
			Content=Replace(Replace(Content,"��",""),"��","")
			ReplaceChar=Content
		End Function
		
		'==================================================
		'��������GetPage
		'��  �ã���ȡ��ҳ
		'==================================================
		Function GetPage(ByVal Constr, StartStr, OverStr, IncluL, IncluR)
		If Constr = "Error" Or Constr = "" Or StartStr = "" Or OverStr = "" Or IsNull(Constr) = True Or IsNull(StartStr) = True Or IsNull(OverStr) = True Then
		   GetPage = "Error"
		   Exit Function
		End If
		
		Dim Start, Over, ConTemp, TempStr
		TempStr = LCase(Constr)
		StartStr = LCase(StartStr)
		OverStr = LCase(OverStr)
		Over = InStr(1, TempStr, OverStr)
		If Over <= 0 Then
		   GetPage = "Error"
		   Exit Function
		Else
		   If IncluR = True Then
			  Over = Over + Len(OverStr)
		   End If
		End If
		TempStr = Mid(TempStr, 1, Over)
		Start = InStrRev(TempStr, StartStr)
		If IncluL = False Then
		   Start = Start + Len(StartStr)
		End If
		
		If Start <= 0 Or Start >= Over Then
		   GetPage = "Error"
		   Exit Function
		End If
		ConTemp = Mid(Constr, Start, Over - Start)
		
		ConTemp = Trim(ConTemp)
		ConTemp = Replace(ConTemp, " ", "")
		ConTemp = Replace(ConTemp, ",", "")
		ConTemp = Replace(ConTemp, "'", "")
		ConTemp = Replace(ConTemp, """", "")
		ConTemp = Replace(ConTemp, ">", "")
		ConTemp = Replace(ConTemp, "<", "")
		ConTemp = Replace(ConTemp, "&nbsp;", "")
		GetPage = ConTemp
		End Function
		
		
		
		Function CheckDir(ByVal FolderPath)
			Dim fso
			Set fso = KS.InitialObject(KS.Setting(99))
			If fso.FolderExists(Server.MapPath(FolderPath)) Then
			'����
				CheckDir = True
			Else
			'������
				CheckDir = False
			End If
			Set fso = Nothing
		End Function
		Function MakeNewsDir(ByVal foldername)
			Dim fso
			Set fso = KS.InitialObject(KS.Setting(99))
				fso.CreateFolder (Server.MapPath(foldername))
				If fso.FolderExists(Server.MapPath(foldername)) Then
				   MakeNewsDir = True
				Else
				   MakeNewsDir = False
				End If
			Set fso = Nothing
		End Function
		
		'**************************************************
		'��������IsObjInstalled
		'��  �ã��������Ƿ��Ѿ���װ
		'��  ����strClassString ----�����
		'����ֵ��True  ----�Ѿ���װ
		'       False ----û�а�װ
		'**************************************************
		Function IsObjInstalled(strClassString)
			IsObjInstalled = False
			Err = 0
			Dim xTestObj
			Set xTestObj = KS.InitialObject(strClassString)
			If 0 = Err Then IsObjInstalled = True
			Set xTestObj = Nothing
			Err = 0
		End Function
	
		
		Sub WriteCollectSucced(ErrMsg,ChannelID)
			Dim strErr
			strErr = strErr & "<html><head><title>�ɹ���Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
			strErr = strErr & "<link href='../KS_Inc/Admin_STYLE.CSS' rel='stylesheet' type='text/css'></head><body>" & vbCrLf
			strErr = strErr & "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""sortbutton"">" & vbCrLf
			strErr = strErr & "</table><br>" & vbCrLf
			strErr = strErr & "<table cellpadding=2 cellspacing=1 border=0 width='90%' class='ctable' align=center>" & vbCrLf
			strErr = strErr & "  <tr align='center' class='sort'><td height='22'><strong>��ϲ�㣡</strong></td></tr>" & vbCrLf
			strErr = strErr & "  <tr class='tdbg'><td height='100' valign='top' align='center'>" & ErrMsg & "</td></tr>" & vbCrLf
			strErr = strErr & "  <tr  class='tdbg' align='center'><td><input type='button' onclick='location.href=""Collect_Main.asp?ChannelID=" & ChannelID & """;parent.frames[""BottomFrame""].location.href=""../KS.Split.asp?OpStr=��Ϣ�ɼ����� >> <font color=red>���ݲɼ�</font>&ButtonSymbol=DataCollect"";' value='���زɼ�����' class='button'></td></tr>" & vbCrLf
			strErr = strErr & "</table>" & vbCrLf
			strErr = strErr & "</body></html>" & vbCrLf
			Response.Write strErr
		End Sub
		Sub WriteCollectSuccedStart(ErrMsg)
			Dim strErr
			strErr = strErr & "<html><head><title>�ɹ���Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
			strErr = strErr & "<link href='../KS_Inc/Admin_STYLE.CSS' rel='stylesheet' type='text/css'></head><body>" & vbCrLf
			strErr = strErr & "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""sortbutton"">" & vbCrLf
			strErr = strErr & "</table><br>" & vbCrLf
			strErr = strErr & "<table cellpadding=2 cellspacing=1 border=0 width=90% class='border' align=center>" & vbCrLf
			strErr = strErr & "  <tr><td height='100' valign='top' align='center'>" & ErrMsg & "</td></tr>" & vbCrLf
			strErr = strErr & "</table>" & vbCrLf
			strErr = strErr & "</body></html>" & vbCrLf
			Response.Write strErr
		End Sub
		
		'**************************************************
		'��������JoinChar
		'��  �ã����ַ�м��� ? �� &
		'��  ����strUrl  ----��ַ
		'����ֵ������ ? �� & ����ַ
		'**************************************************
		Function JoinChar(strUrl)
			If strUrl = "" Then
				JoinChar = ""
				Exit Function
			End If
			If InStr(strUrl, "?") < Len(strUrl) Then
				If InStr(strUrl, "?") > 1 Then
					If InStr(strUrl, "&") < Len(strUrl) Then
						JoinChar = strUrl & "&"
					Else
						JoinChar = strUrl
					End If
				Else
					JoinChar = strUrl & "?"
				End If
			Else
				JoinChar = strUrl
			End If
		End Function
		
		'**************************************************
		'��������CreateKeyWord
		'��  �ã��ɸ������ַ������ɹؼ���
		'��  ����Constr---Ҫ���ɹؼ��ֵ�ԭ�ַ���
		'����ֵ�����ɵĹؼ���
		'**************************************************
		Function CreateKeyWord(ByVal Constr, num)
		   If Constr = "" Or IsNull(Constr) = True Or Constr = "Error" Then
			  CreateKeyWord = "Error"
			  Exit Function
		   End If
		   
		   Dim MaxLen:MaxLen=25
		   Dim WS:Set WS=New Wordsegment_Cls
			 CreateKeyWord=WS.SplitKey(KS.R(Constr),4,MaxLen)
		   Set WS=Nothing
		End Function
		
		Function CheckUrl(strUrl)
		   
		   '����,��ʱ������
		   
		   CheckUrl = strUrl
		   
		   'Dim re
		   'Set re = New RegExp
		   're.IgnoreCase = True
		   're.Global = True
		   're.Pattern = "http://([\w-]+\.)+[\w-]+(/[\w-./?%&=]*)?"
		   'If re.Test(strUrl) = True Then
		   '   CheckUrl = strUrl
		   'Else
		   '   CheckUrl = "Error"
		   'End If
		
		End Function
		
		
		
		
		
		
		
		
		
		
		
		Function UBBCode(ByVal strContent, strInstallDir, strChannelDir)
			Dim ImagePath
			Dim emotImagePath
			
			ImagePath = strInstallDir & "images/"
			emotImagePath = strInstallDir & "guestbook/images/emot/"
			strContent = FilterJS(strContent)
			Dim re
			Dim po, ii
			Dim reContent
			Set re = New RegExp
			re.IgnoreCase = True
			re.Global = True
			po = 0
			ii = 0
		
			re.Pattern = "\[IMG\](.)\[\/IMG\]"
			strContent = re.Replace(strContent, "<img src='$1' border=0>")
				
			re.Pattern = "\[IMG\](http|https|ftp):\/\/(.[^\[]*)\[\/IMG\]"
			strContent = re.Replace(strContent, "<a onfocus=this.blur() href=""$1://$2"" target=_blank><IMG SRC=""$1://$2"" border=0 alt=�������´������ͼƬ onload=""javascript:if(this.width>screen.width-333)this.width=screen.width-333""></a>")
			re.Pattern = "\[UPLOAD=(gif|jpg|jpeg|bmp|png)\](.[^\[]*)(gif|jpg|jpeg|bmp|png)\[\/UPLOAD\]"
			strContent = re.Replace(strContent, "<br><IMG SRC=""" & ImagePath & "$1.gif"" border=0>���������ͼƬ���£�<br><A HREF=""$2$1"" TARGET=_blank><IMG SRC=""$2$1"" border=0 alt=�������´������ͼƬ onload=""javascript:if(this.width>screen.width-333)this.width=screen.width-333""></A>")
		
			re.Pattern = "\[UPLOAD=(.[^\[]*)\](.[^\[]*)\[\/UPLOAD\]"
			strContent = re.Replace(strContent, "<br><IMG SRC=""" & ImagePath & "$1.gif"" border=0> <a href=""" & strInstallDir & strChannelDir & "/$2"">���������ļ�</a>")
		
			re.Pattern = "\[DIR=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/DIR]"
			strContent = re.Replace(strContent, "<object classid=clsid:166B1BCA-3F9C-11CF-8075-444553540000 codebase=http://download.macromedia.com/pub/shockwave/cabs/director/sw.cab#version=7,0,2,0 width=$1 height=$2><param name=src value=$3><embed src=$3 pluginspage=http://www.macromedia.com/shockwave/download/ width=$1 height=$2></embed></object>")
			re.Pattern = "\[QT=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/QT]"
			strContent = re.Replace(strContent, "<embed src=$3 width=$1 height=$2 autoplay=true loop=false controller=true playeveryframe=false cache=false scale=TOFIT bgcolor=#000000 kioskmode=false targetcache=false pluginspage=http://www.apple.com/quicktime/>")
			re.Pattern = "\[MP=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/MP]"
			strContent = re.Replace(strContent, "<object align=middle classid=CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95 class=OBJECT id=MediaPlayer width=$1 height=$2 ><param name=ShowStatusBar value=-1><param name=Filename value=$3><embed type=application/x-oleobject codebase=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701 flename=mp src=$3  width=$1 height=$2></embed></object>")
			re.Pattern = "\[RM=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/RM]"
			strContent = re.Replace(strContent, "<OBJECT classid=clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA class=OBJECT id=RAOCX width=$1 height=$2><PARAM NAME=SRC VALUE=$3><PARAM NAME=CONSOLE VALUE=Clip1><PARAM NAME=CONTROLS VALUE=imagewindow><PARAM NAME=AUTOSTART VALUE=true></OBJECT><br><OBJECT classid=CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA height=32 id=video2 width=$1><PARAM NAME=SRC VALUE=$3><PARAM NAME=AUTOSTART VALUE=-1><PARAM NAME=CONTROLS VALUE=controlpanel><PARAM NAME=CONSOLE VALUE=Clip1></OBJECT>")
		
			re.Pattern = "(\[FLASH\])(.[^\[]*)(\[\/FLASH\])"
			strContent = re.Replace(strContent, "<a href=""$2"" TARGET=_blank><IMG SRC=" & ImagePath & "swf.gif border=0 alt=������´������͸�FLASH����! height=16 width=16>[ȫ������]</a><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=500 height=400><PARAM NAME=movie VALUE=""$2""><PARAM NAME=quality VALUE=high><embed src=""$2"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=500 height=400>$2</embed></OBJECT>")
		
			re.Pattern = "(\[FLASH=*([0-9]*),*([0-9]*)\])(.[^\[]*)(\[\/FLASH\])"
			strContent = re.Replace(strContent, "<a href=""$4"" TARGET=_blank><IMG SRC=" & ImagePath & "swf.gif border=0 alt=������´������͸�FLASH����! height=16 width=16>[ȫ������]</a><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=$2 height=$3><PARAM NAME=movie VALUE=""$4""><PARAM NAME=quality VALUE=high><embed src=""$4"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=$2 height=$3>$4</embed></OBJECT>")
		
			re.Pattern = "(\[URL\])(.[^\[]*)(\[\/URL\])"
			strContent = re.Replace(strContent, "<A HREF=""$2"" TARGET=_blank>$2</A>")
			re.Pattern = "(\[URL=(.[^\[]*)\])(.[^\[]*)(\[\/URL\])"
			strContent = re.Replace(strContent, "<A HREF=""$2"" TARGET=_blank>$3</A>")
		
			re.Pattern = "(\[EMAIL\])(\S+\@.[^\[]*)(\[\/EMAIL\])"
			strContent = re.Replace(strContent, "<img align=absmiddle src=" & ImagePath & "email1.gif><A HREF=""mailto:$2"">$2</A>")
			re.Pattern = "(\[EMAIL=(\S+\@.[^\[]*)\])(.[^\[]*)(\[\/EMAIL\])"
			strContent = re.Replace(strContent, "<img align=absmiddle src=" & ImagePath & "email1.gif><A HREF=""mailto:$2"" TARGET=_blank>$3</A>")
		
			'�Զ�ʶ����ַ
			're.Pattern = "^((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
			'strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif border=0><a target=_blank href=$1>$1</a>")
			're.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@':+!]+)$"
			'strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif border=0><a target=_blank href=$1>$1</a>")
			're.Pattern = "([^>=""])((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
			'strContent = re.Replace(strContent,"$1<img align=absmiddle src=pic/url.gif border=0><a target=_blank href=$2>$2</a>")
		
			'�Զ�ʶ��www�ȿ�ͷ����ַ
			're.Pattern = "([^(http://|http:\\)])((www|cn)[.](\w)+[.]{1,}(net|com|cn|org|cc)(((\/[\~]*|\\[\~]*)(\w)+)|[.](\w)+)*(((([?](\w)+){1}[=]*))*((\w)+){1}([\&](\w)+[\=](\w)+)*)*)"
			'strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif border=0><a target=_blank href=http://$2>$2</a>")
		
			'�Զ�ʶ��Email��ַ����򿪱�������������ݺܶ�����ӻ����������ͣ��
			're.Pattern = "([^(=)])((\w)+[@]{1}((\w)+[.]){1,3}(\w)+)"
			'strContent = re.Replace(strContent,"<img align=absmiddle src=pic/url.gif border=0><a target=_blank href=""mailto:$2"">$2</a>")
		
			re.Pattern = "\[em(.[^\[]*)\]"
			strContent = re.Replace(strContent, "<img src=" & emotImagePath & "em$1.gif border=0 align=middle>")
		
			re.Pattern = "\[HTML\](.[^\[]*)\[\/HTML\]"
			strContent = re.Replace(strContent, "<table width='100%' border='0' cellspacing='0' cellpadding='6' class=tableborder1><td><b>��������Ϊ�������:</b><br>$1</td></table>")
			re.Pattern = "\[code\](.[^\[]*)\[\/code\]"
			strContent = re.Replace(strContent, "<table width='100%' border='0' cellspacing='0' cellpadding='6' class=tableborder1><td><b>��������Ϊ�������:</b><br>$1</td></table>")
		
			re.Pattern = "\[color=(.[^\[]*)\](.[^\[]*)\[\/color\]"
			strContent = re.Replace(strContent, "<font color=$1>$2</font>")
			re.Pattern = "\[face=(.[^\[]*)\](.[^\[]*)\[\/face\]"
			strContent = re.Replace(strContent, "<font face=$1>$2</font>")
			re.Pattern = "\[align=(center|left|right)\](.*)\[\/align\]"
			strContent = re.Replace(strContent, "<div align=$1>$2</div>")
		
			re.Pattern = "\[QUOTE\](.*)\[\/QUOTE\]"
			strContent = re.Replace(strContent, "<table style=""width:80%"" cellpadding=5 cellspacing=1 class=tableborder1><TR><TD class=tableborder1>$1</td></tr></table><br>")
			re.Pattern = "\[fly\](.*)\[\/fly\]"
			strContent = re.Replace(strContent, "<marquee width=90% behavior=alternate scrollamount=3>$1</marquee>")
			re.Pattern = "\[move\](.*)\[\/move\]"
			strContent = re.Replace(strContent, "<MARQUEE scrollamount=3>$1</marquee>")
			re.Pattern = "\[GLOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/GLOW]"
			strContent = re.Replace(strContent, "<table width=$1 style=""filter:glow(color=$2, strength=$3)"">$4</table>")
			re.Pattern = "\[SHADOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/SHADOW]"
			strContent = re.Replace(strContent, "<table width=$1 style=""filter:shadow(color=$2, strength=$3)"">$4</table>")
		
			re.Pattern = "\[i\](.[^\[]*)\[\/i\]"
			strContent = re.Replace(strContent, "<i>$1</i>")
			re.Pattern = "\[u\](.[^\[]*)(\[\/u\])"
			strContent = re.Replace(strContent, "<u>$1</u>")
			re.Pattern = "\[b\](.[^\[]*)(\[\/b\])"
			strContent = re.Replace(strContent, "<b>$1</b>")
			re.Pattern = "\[size=([1-7])\](.[^\[]*)\[\/size\]"
			strContent = re.Replace(strContent, "<font size=$1>$2</font>")
			strContent = Replace(strContent, "<I></I>", "")
			Set re = Nothing
			UBBCode = strContent
		End Function
		
		Function FilterJS(ByVal v)
			If IsNull(v) Or Trim(v) = "" Then
				FilterJS = ""
				Exit Function
			End If
		
			Dim t
			Dim re
			Dim reContent
			Set re = New RegExp
			re.IgnoreCase = True
			re.Global = True
			re.Pattern = "(javascript)"
			t = re.Replace(v, "&#106avascript")
			re.Pattern = "(jscript:)"
			t = re.Replace(t, "&#106script:")
			re.Pattern = "(js:)"
			t = re.Replace(t, "&#106s:")
			're.Pattern="(value)"
			't=re.Replace(t,"&#118alue")
			re.Pattern = "(about:)"
			t = re.Replace(t, "about&#58")
			re.Pattern = "(file:)"
			t = re.Replace(t, "file&#58")
			re.Pattern = "(document.cookie)"
			t = re.Replace(t, "documents&#46cookie")
			re.Pattern = "(vbscript:)"
			t = re.Replace(t, "&#118bscript:")
			re.Pattern = "(vbs:)"
			t = re.Replace(t, "&#118bs:")
			're.Pattern="(on(mouse|exit|error|click|key))"
			't=re.Replace(t,"&#111n$2")
			're.Pattern="(&#)"
			't=re.Replace(t,"��#")
			FilterJS = t
			Set re = Nothing
		End Function
		
		Function dvHTMLEncode(ByVal fString)
			If IsNull(fString) Or Trim(fString) = "" Then
				dvHTMLEncode = ""
				Exit Function
			End If
			fString = Replace(fString, ">", "&gt;")
			fString = Replace(fString, "<", "&lt;")
		
			fString = Replace(fString, Chr(32), "&nbsp;")
			fString = Replace(fString, Chr(9), "&nbsp;")
			fString = Replace(fString, Chr(34), "&quot;")
			fString = Replace(fString, Chr(39), "&#39;")
			fString = Replace(fString, Chr(13), "")
			fString = Replace(fString, Chr(10) & Chr(10), "</P><P> ")
			fString = Replace(fString, Chr(10), "<BR> ")
		
			dvHTMLEncode = fString
		End Function
		
		Function dvHTMLCode(ByVal fString)
			If IsNull(fString) Or Trim(fString) = "" Then
				dvHTMLCode = ""
				Exit Function
			End If
			fString = Replace(fString, "&gt;", ">")
			fString = Replace(fString, "&lt;", "<")
		
			fString = Replace(fString, "&nbsp;", " ")
			fString = Replace(fString, "&quot;", Chr(34))
			fString = Replace(fString, "&#39;", Chr(39))
			fString = Replace(fString, "</P><P> ", Chr(10) & Chr(10))
			fString = Replace(fString, "<BR> ", Chr(10))
		
			dvHTMLCode = fString
		End Function
		
		Function nohtml(ByVal str)
			If IsNull(str) Or Trim(str) = "" Then
				nohtml = ""
				Exit Function
			End If
			Dim re
			Set re = New RegExp
			re.IgnoreCase = True
			re.Global = True
			re.Pattern = "(\<.[^\<]*\>)"
			str = re.Replace(str, "")
			re.Pattern = "(\<\/[^\<]*\>)"
			str = re.Replace(str, "")
			Set re = Nothing
			str = Replace(str, Chr(34), "")
			str = Replace(str, "'", "")
			nohtml = str
		End Function
		'===============================================================================
		'������: CheckTheChar
		'�� ��: ���ĳһ�Ӵ����ֵĴ���
		'�� ��:TheChar="Ҫ�����ַ���",TheString="�������ַ���"
		'================================================================================
		Function CheckTheChar(TheChar, TheString)
		  Dim n
		  If InStr(TheString, TheChar) Then
			For n = 1 To Len(TheString)
			  If Mid(TheString, n, Len(TheChar)) = TheChar Then
			  CheckTheChar = CheckTheChar + 1
			  End If
			Next
			  CheckTheChar = CheckTheChar
		  Else
			  CheckTheChar = 0
		  End If
		End Function

End Class

'ͨ�û�����
Class ClsCache
        Private cache           '��������
        Private cacheName       '����Application����
        Private expireTime      '�������ʱ��
        Private expireTimeName  '�������ʱ��Application����
        Private vaild           'ansir���
		Private Sub Class_Initialize()
		End Sub
        Private Sub Class_Terminate()
		End Sub
		Property Get Version()
         Version = "Kesion Cache"
		End Property
		
		Property Get valid() 
		If IsEmpty(cache) Or (Not IsDate(expireTime)) Then
		vaild = False
		Else
		valid = True
		End If
		End Property
		
		Property Get value()
		If IsEmpty(cache) Or (Not IsDate(expireTime)) Then
		value = Null
		ElseIf CDate(expireTime) < Now Then
		value = Null
		Else
		value = cache
		End If
		End Property
		
		Public Property Let name(str)
		cacheName = str
		cache = Application(cacheName)
		expireTimeName = str & "_expire"
		expireTime = Application(expireTimeName)
		End Property
		
		Public Property Let expire(tm)
		expireTime = tm
		Application.Lock
		Application(expireTimeName) = expireTime
		Application.UnLock
		End Property
		
		Public Sub add(varCache, varExpireTime) 
		If IsEmpty(varCache) Or Not IsDate(varExpireTime) Then
		Exit Sub
		End If
		cache = varCache
		expireTime = varExpireTime
		Application.Lock
		Application(cacheName) = cache
		Application(expireTimeName) = expireTime
		Application.UnLock
		End Sub
		
		Public Sub clean()
		Application.Lock
		Application(cacheName) = Empty
		Application(expireTimeName) = Empty
		Application.UnLock
		cache = Empty
		expireTime = Empty
		End Sub
		 
		Public Function verify(varcache2) 
		If TypeName(cache) <> TypeName(varcache2) Then
			verify = False
		ElseIf TypeName(cache) = "Object" Then
			If cache Is varcache2 Then
				verify = True
			Else
				verify = False
			End If
		ElseIf TypeName(cache) = "Variant()" Then
			If Join(cache, "^") = Join(varcache2, "^") Then
				verify = True
			Else
				verify = False
			End If
		Else
			If cache = varcache2 Then
				verify = True
			Else
				verify = False
			End If
		End If
		End Function
End Class
%> 
