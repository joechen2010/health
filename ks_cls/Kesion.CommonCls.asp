<!--#include file="Kesion.Thumbs.asp"-->
<!--#include file="Kesion.TranPinYinCls.asp"-->
<!--#include file="Kesion.VersionCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************

Const ClassField="ID,FolderName,Folder,ClassPurview,FolderDomain,TemplateID,ClassBasicInfo,ClassDefineContent,TS,ClassID,Tj,DefaultDividePercent,ChannelID,TN,ClassType,FolderOrder,AdminPurview,AllowArrGroupID,CommentTF,Child"           '�������뻺�����Ŀ�ֶ�

Class PublicCls
		Public SiteSN,Version
		Public Setting,TbSetting,SSetting,JSetting,ASetting,WSetting
	  Private Sub Class_Initialize()
		if Not Response.IsClientConnected then die ""
		Call Initialize_Kesion_Config
      End Sub
	 Private Sub Class_Terminate()

	 End Sub
	 
	 Function InitialObject(str)
		'iis5�������󷽷�Server.CreateObject(ObjectName);
		'iis6�������󷽷�CreateObject(ObjectName);
		'Ĭ��Ϊiis6�������iis5��ʹ�ã���Ҫ��ΪServer.CreateObject(str);
		Set InitialObject=CreateObject(str)
	 End Function
	 '*******************************************************************************************************************
	 '��������Initialize_Kesion_Config
	 '��  ��: ����KesionCMS�ı�Ҫ����
	 '��  ע�����²����벻Ҫ���ġ�����ϵͳ�����޷���������
	 '*******************************************************************************************************************
	 Public Function Initialize_Kesion_Config()
		Dim KCls:Set KCls=New KesionCls
		SiteSN =KCls.SiteSN 
		Version = KCls.KSVer
        Set KCls=Nothing
		Call InitialConfig()
		Call IsIPlock()      'IP��������
	 End Function
	 

	'����ʾ,�����������,���� PreCacheName-ǰ��ƥ��
	Public Sub DelCaches(PreCacheName)
	    Dim i
		Dim CacheList:CacheList=split(GetCacheList(PreCacheName),",")
		If UBound(CacheList)>1 Then
			For i=0 to UBound(CacheList)-1
				DelCahe CacheList(i)
			Next
		End IF
	End Sub
	'ȡ�û����б� ���� PreCacheName-ǰ��ƥ��
	Public Function GetCacheList(PreCacheName)
		Dim Cacheobj
		For Each Cacheobj in Application.Contents
		If CStr(Left(Cacheobj,Len(PreCacheName)))=CStr(PreCacheName) Then GetCacheList=GetCacheList&Cacheobj&","
		Next
	End Function
	'�������,���� MyCaheName-��������
	Public Sub DelCahe(MyCaheName)
		Application.Lock
		Application.Contents.Remove(MyCaheName)
		Application.unLock
	End Sub

	 Public Sub GetSetting()
		    Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		    RSObj.Open "SELECT top 1 Setting,TbSetting,SpaceSetting,JobSetting,AskSetting,WapSetting from [KS_Config]",conn,1,1
		    Dim i,node,xml,j,DataArray,rs
			Set xml = Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			xml.appendChild(xml.createElement("xml"))
			If Not RSObj.EOF Then
						DataArray=RSObj.GetRows(1)
						For i=0 To UBound(DataArray,2)
							Set Node=xml.createNode(1,"config","")
							j=0
							For Each rs in RSObj.Fields
								node.attributes.setNamedItem(xml.createNode(2,LCase(rs.name),"")).text= Replace(DataArray(j,i),vbcrlf,"$br$")& ""
								j=j+1
							Next
							xml.documentElement.appendChild(Node)
						Next
			End If
			DataArray=Null
		   Set Application(SiteSN&"_Config")=Xml
		   RSObj.Close:Set RSObj=Nothing
	 End Sub

	 Public Sub InitialConfig()
		If not IsObject(Application(SiteSN&"_Config")) then  GetSetting
		Setting=Split(Replace(Application(SiteSN&"_Config").documentElement.selectSingleNode("config/@setting").text,"$br$",vbcrlf),"^%^")
		TbSetting=Split(Replace(Application(SiteSN&"_Config").documentElement.selectSingleNode("config/@tbsetting").text,"$br$",vbcrlf),"^%^")
        SSetting=Split(Replace(Application(SiteSN&"_Config").documentElement.selectSingleNode("config/@spacesetting").text,"$br$",vbcrlf),"^%^")
		JSetting=Split(Replace(Application(SiteSN&"_Config").documentElement.selectSingleNode("config/@jobsetting").text,"$br$",vbcrlf),"^%^")
		ASetting=Split(Replace(Application(SiteSN&"_Config").documentElement.selectSingleNode("config/@asksetting").text,"$br$",vbcrlf),"^%^")
		WSetting=Split(Replace(Application(SiteSN&"_Config").documentElement.selectSingleNode("config/@wapsetting").text,"$br$",vbcrlf),"^%^")
	 End Sub
	
	 'xmlroot���ڵ����� row��¼�нڵ�����
	 Public Function RecordsetToxml(RSObj,row,xmlroot)
	  Dim i,node,rs,j,DataArray
	  If xmlroot="" Then xmlroot="xml"
	  If row="" Then row="row"
	  Set RecordsetToxml=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	  RecordsetToxml.appendChild(RecordsetToxml.createElement(xmlroot))
	  If Not RSObj.EOF Then
	   DataArray=RSObj.GetRows(-1)
	   For i=0 To UBound(DataArray,2)
		Set Node=RecordsetToxml.createNode(1,row,"")
		j=0
		For Each rs in RSObj.Fields		   
		   node.attributes.setNamedItem(RecordsetToxml.createNode(2,"ks"&j,"")).text= DataArray(j,i)& ""
		   j=j+1
		Next
		RecordsetToxml.documentElement.appendChild(Node)
	   Next
	  End If
	  DataArray=Null
	 End Function
	 
	 'xmlroot���ڵ����� row��¼�нڵ�����
	Public Function RsToxml(RSObj,row,xmlroot)
			Dim i,node,rs,j,DataArray
			If xmlroot="" Then xmlroot="xml"
			If row="" Then row="row"
			Set RsToxml = Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			RsToxml.appendChild(RsToxml.createElement(xmlroot))
			If Not RSObj.EOF Then
						DataArray=RSObj.GetRows(-1)
						For i=0 To UBound(DataArray,2)
							Set Node=RsToxml.createNode(1,row,"")
							j=0
							For Each rs in RSObj.Fields
								node.attributes.setNamedItem(RsToxml.createNode(2,LCase(rs.name),"")).text= DataArray(j,i)& ""
								j=j+1
							Next
							RsToxml.documentElement.appendChild(Node)
						Next
			End If
			DataArray=Null
	End Function
	Public Function ArrayToxml(DataArray,Recordset,row,xmlroot)
				Dim i,node,rs,j
				If xmlroot="" Then xmlroot="xml"
				Set ArrayToxml = Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
				ArrayToxml.appendChild(ArrayToxml.createElement(xmlroot))
				If row="" Then row="row"
				For i=0 To UBound(DataArray,2)
					Set Node=ArrayToxml.createNode(1,row,"")
					j=0
					For Each rs in Recordset.Fields
							 node.attributes.setNamedItem(ArrayToxml.createNode(2,LCase(rs.name),"")).text= DataArray(j,i)& ""
							 j=j+1
					Next
					ArrayToxml.documentElement.appendChild(Node)
				Next
		End Function
	 
	 Public Function LoadChannelConfig()
	 Application.Lock
	 Dim RS:Set Rs=conn.execute("select ChannelID,ChannelName,ChannelTable,ItemName,ItemUnit,FieldBit,BasicType,FsoHtmlTF,FsoFolder,RefreshFlag,ModelEname,MaxPerPage,VerificCommentTF,CommentVF,CommentLen,CommentTemplate,UserSelectFilesTF,InfoVerificTF,UserAddMoney,UserAddPoint,UserAddScore,ChannelStatus,CollectTF,UpFilesTF,UpFilesDir,UpFilesSize,UserUpFilesTF,UserUpFilesDir,AllowUpPhotoType,AllowUpFlashType,AllowUpMediaType,AllowUpRealType,AllowUpOtherType,SearchTemplate,EditorType,FsoListNum,UserTF,DiggByVisitor,DiggByIP,DiggRepeat,DiggPerTimes,UserClassStyle,UserEditTF,FsoContentRule,FsoClassListRule,FsoClassPreTag,ThumbnailsConfig,LatestNewDay,StaticTF,PubTimeLimit From KS_Channel Order by ChannelID")
	 Set Application(SiteSN&"_ChannelConfig")=RecordsetToxml(rs,"channel","ChannelConfig")
	 Set Rs=Nothing
	 Application.unLock
	 End Function
	 
	 Function C_S(sChannelID,FieldID)
	  If IsNul(sChannelID) Then Exit Function
	  If not IsObject(Application(SiteSN&"_ChannelConfig")) Then LoadChannelConfig()
	  Dim Node:Set Node=Application(SiteSN&"_ChannelConfig").documentElement.selectSingleNode("channel[@ks0=" & sChannelID & "]/@ks" & FieldID & "")
	  If Not Node Is Nothing  Then C_S = Node.Text Else C_S=0
	  Set Node = Nothing
	 End Function
	 
	 Public Function LoadClassConfig()
		If not IsObject(Application(SiteSN&"_class")) Then
		 Application.Lock
		 Dim RS:Set Rs=conn.execute("select " & ClassField & " From KS_Class Order by root,folderorder")
		 Set Application(SiteSN&"_class")=RecordsetToxml(rs,"class","classConfig")
		 Set Rs=Nothing
		 Application.unLock
	   End If
	 End Function

	 Function C_C(ClassID,FieldID)
	   If ClassID="" Or IsNull(ClassID) Then Exit Function
	   LoadClassConfig()
	   Dim Node:Set Node=Application(SiteSN&"_class").documentElement.selectSingleNode("class[@ks0=" & classID & "]/@ks" & FieldID & "")
	   If Not Node Is Nothing Then C_C=Node.text
	   Set Node=Nothing
	 End Function
	
	 '�����û��黺��
	 Sub LoadUserGroup()
	   If Not IsObject(Application(SiteSN&"_UserGroup")) Then 
	    Application.Lock
	     Dim RS:Set Rs=conn.execute("select id,groupname,powerlist,descript,usertype,formid,templatefile,showonreg,ChargeType,GroupPoint,GroupSetting From KS_UserGroup Order by ID")
		 Set Application(SiteSN&"_UserGroup")=RsToxml(rs,"row","groupConfig")
         Set Rs=Nothing
	     Application.unLock
	   End If
	 End Sub
	 '��ȡ�û�������Ȩ��
	 Function U_S(GroupID,i)
	   If IsNul(GroupID) Then U_S=0 : Exit Function
	   Dim GroupSetting:GroupSetting=U_G(GroupID,"GroupSetting") &",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
	   Dim GroupSetArr:GroupSetArr=Split(GroupSetting,",")
	   U_S=GroupSetArr(i)
	 End Function
	 
	 Function U_G(GroupID,FieldName)
	   If IsNul(GroupID) Then Exit Function
	   LoadUserGroup
	   Dim Node:Set Node=Application(SiteSN&"_UserGroup").DocumentElement.selectSingleNode("row[@id=" & GroupID & "]/@" & Lcase(FieldName))
	   If Not Node Is Nothing Then U_G=Node.text
	   Set Node=Nothing
	 End Function
	 
	 '�������԰��滺��
	 Sub LoadClubBoard()
	   If Not IsObject(Application(SiteSN&"_ClubBoard")) Then 
	    Application.Lock
	     Dim RS:Set Rs=conn.execute("select [id],[boardname],[note],[master],[todaynum],[postnum],[topicnum],[parentid],[LastPost],[BoardRules],[Settings] From KS_GuestBoard Order by orderid,ID")
		 Set Application(SiteSN&"_ClubBoard")=RsToxml(rs,"row","clubConfig")
         Set Rs=Nothing
	     Application.unLock
	   End If
	 End Sub

	
	'**************************************************
	'��������LoadClassOption
	'��  �ã�������Ŀѡ��
	'��  ����ChannelID-----��ǰģ��ID
	'����ֵ��������
	'**************************************************
	Public Function LoadClassOption(ChannelID)
		Dim Node,K,SQL,NodeText,Pstr,TJ,SpaceStr,TreeStr
		LoadClassConfig()
		If ChannelID<>0 Then Pstr="and @ks12=" & channelid & ""
		For Each Node In Application(SiteSN&"_class").DocumentElement.SelectNodes("class[@ks14=1" & Pstr&"]")
		  SpaceStr=""
		  If (C("SuperTF")=1 or FoundInArr(Node.SelectSingleNode("@ks16").text,C("AdminName"),",") or Instr(C("ModelPower"),C_S(Node.SelectSingleNode("@ks12").text,10)&"1")>0) and (C_S(Node.SelectSingleNode("@ks12").text,21)=1 or Node.SelectSingleNode("@ks12").text=5) Then 
			  TJ=Node.SelectSingleNode("@ks10").text
			  If TJ>1 Then
				 For k = 1 To TJ - 1
					SpaceStr = SpaceStr & "����"
				 Next
				TreeStr = TreeStr & "<option value='" & Node.SelectSingleNode("@ks0").text & "'>" & SpaceStr & Node.SelectSingleNode("@ks1").text & " </option>"
			  Else
				TreeStr = TreeStr & "<option value='" & Node.SelectSingleNode("@ks0").text & "'>" & Node.SelectSingleNode("@ks1").text & " </option>"
			  End If
		 End If 
		Next
		LoadClassOption=TreeStr
	End Function
	 
	
	Sub Echo(Str)
	  Response.Write Str
	End Sub
	
	Sub Die(Str)
	  Response.Write Str : Response.End
	End Sub
	
	Function IsNul(Str)
	  If Str="" Or IsNull(Str) Then IsNul=True Else IsNul=false
	End Function
	
	Sub LoadChannelField()
	  If Not IsObject(Application(SiteSN & "_ChannelField")) then
			Dim Rs:Set Rs = Conn.Execute("Select ChannelID,Title,FieldName From KS_Field Order By FieldID")
			Set Application(SiteSN & "_ChannelField")=RsToxml(Rs,"row","root")
			Set Rs = Nothing
	  End If
	End Sub
	 
	 Sub IsIPlock()
	   On Error Resume Next
	    If Setting(100)=0 Then Exit Sub
		If session("KS_IPlock") = "" Then
			session("KS_IPlock") = CheckIPlock(Setting(100), Setting(101), GetIP)
		End If
		If session("KS_IPlock") = True Then
			die "�Բ�������IP��" &GetIP & "����ϵͳ�޶��������Ժ�վ����ϵ��"
		End If
	End Sub
	Function EncodeIP(Sip)
		Dim strIP:strIP = Split(Sip, ".")
		If UBound(strIP) < 3 Then
			EncodeIP = 0:Exit Function
		End If
		If IsNumeric(strIP(0)) = 0 Or IsNumeric(strIP(1)) = 0 Or IsNumeric(strIP(2)) = 0 Or IsNumeric(strIP(3)) = 0 Then
			Sip = 0
		Else
			Sip = CInt(strIP(0)) * 256 * 256 * 256 + CInt(strIP(1)) * 256 * 256 + CInt(strIP(2)) * 256 + CInt(strIP(3)) - 1
		End If
		EncodeIP = Sip
	End Function
	Function CStrIP(ByVal anNewIP)
	Dim lsResults ' Results To be returned
	Dim lnTemp ' Temporary value being parsed
	Dim lnIndex ' Position of number being parsed
	For lnIndex = 3 To 0 Step-1
	lnTemp = Int(anNewIP / (256 ^ lnIndex))
	lsResults = lsResults & lnTemp & "."
	anNewIP = anNewIP - (lnTemp * (256 ^ lnIndex))
	Next
	lsResults = Left(lsResults, Len(lsResults) - 1)
	lsResults=Split(lsResults,".")
	Dim IPStr,i:For I=0 To Ubound(lsResults)
	 if i=3 then 
	  IPStr=IPStr & "." &lsResults(3)+1
	 elseif i=0 then 
	   IPStr=lsResults(0) 
	 else 
	  IPStr=IPStr & "." & lsResults(i)
	 end if
	Next
	CStrIP = IPStr
	End Function 
	'�������Ķ˵���Է��ʺͺ������Ķ˵㽫��������ʡ�
	Function ChecKIPlock(ByVal sLockType, ByVal sLockList, ByVal sUserIP)
		Dim IPlock, rsLockIP
		Dim arrLockIPW, arrLockIPB, arrLockIPWCut, arrLockIPBCut
		IPlock = False
		ChecKIPlock = IPlock
		Dim i, sKillIP
		If sLockType = "" Or IsNull(sLockType) Then Exit Function
		If sLockList = "" Or IsNull(sLockList) Then Exit Function
		If sUserIP = "" Or IsNull(sUserIP) Then Exit Function
		sUserIP = CDbl(EncodeIP(sUserIP))
		rsLockIP = Split(sLockList, "|||")
		If sLockType = 4 Then
			arrLockIPB = Split(Trim(rsLockIP(1)), "$$$")
			For i = 0 To UBound(arrLockIPB)
				If arrLockIPB(i) <> "" Then
					arrLockIPBCut = Split(Trim(arrLockIPB(i)), "----")
					IPlock = True
					If CDbl(arrLockIPBCut(0)) > sUserIP Or sUserIP > CDbl(arrLockIPBCut(1)) Then IPlock = False
					If IPlock Then Exit For
				End If
			Next
			If IPlock = True Then
				arrLockIPW = Split(Trim(rsLockIP(0)), "$$$")
				For i = 0 To UBound(arrLockIPW)
					If arrLockIPW(i) <> "" Then
						arrLockIPWCut = Split(Trim(arrLockIPW(i)), "----")
						IPlock = True
						If CDbl(arrLockIPWCut(0)) <= sUserIP And sUserIP <= CDbl(arrLockIPWCut(1)) Then IPlock = False
						If IPlock Then Exit For
					End If
				Next
			End If
		Else
			If sLockType = 1 Or sLockType = 3 Then
				arrLockIPW = Split(Trim(rsLockIP(0)), "$$$")
				For i = 0 To UBound(arrLockIPW)
					If arrLockIPW(i) <> "" Then
						arrLockIPWCut = Split(Trim(arrLockIPW(i)), "----")
						IPlock = True
						If CDbl(arrLockIPWCut(0)) <= sUserIP And sUserIP <= CDbl(arrLockIPWCut(1)) Then IPlock = False
						If IPlock Then Exit For
					End If
				Next
			End If
			If IPlock = False And (sLockType = 2 Or sLockType = 3) Then
				arrLockIPB = Split(Trim(rsLockIP(1)), "$$$")
				For i = 0 To UBound(arrLockIPB)
					If arrLockIPB(i) <> "" Then
						arrLockIPBCut = Split(Trim(arrLockIPB(i)), "----")
						IPlock = True
						If CDbl(arrLockIPBCut(0)) > sUserIP Or sUserIP > CDbl(arrLockIPBCut(1)) Then IPlock = False
						If IPlock Then Exit For
					End If
				Next
			End If
		End If
		ChecKIPlock = IPlock
	End Function
    Public Function Conn()
	   On Error Resume Next
	  Dim ConnObj:Set ConnObj=Server.CreateObject("ADODB.Connection")
	  ConnObj.Open ConnStr
	  Set Conn = ConnObj
	End Function
	'�ɼ����ݿ�����
	Public Function ConnItem()
	  Dim ConnObj:Set ConnObj=Server.CreateObject("ADODB.Connection")
	  ConnObj.Open CollcetConnStr
	  Set ConnItem = ConnObj
	End Function

	
	'***************************************************************************************************************
	'��������GetDomain
	'��  �ã���ȡURL,��������Ŀ¼ ��http://www.kesion.com/ �� http://www.kesion.com/Sys/  ���� Sys/Ϊ����Ŀ¼
	'��  ����  ��
	'����ֵ����������
	'***************************************************************************************************************
	Public Function GetDomain()
	    GetDomain = Trim(Setting(2) & Setting(3))
	End Function
	'**************************************************
	'��������GetChannelDomain
	'��  �ã���ȡ����Ƶ��������Url
	'��  ����ChannelIDƵ��ID
	'����ֵ����������
	'**************************************************
	Public Function GetChannelDomain(ChannelID)
		GetChannelDomain=C_S(ChannelID,8)
		If Left(GetChannelDomain, 1) = "/" Then GetChannelDomain = Right(GetChannelDomain, Len(GetChannelDomain) - 1)
		GetChannelDomain = GetDomain() & GetChannelDomain
	End Function
	'**************************************************
	'��������GetAutoDoMain()
	'��  �ã�ȡ�õ�ǰ������IP �磺http://127.0.0.1
	'��  ������
	'**************************************************
	Public Function GetAutoDomain()
		Dim TempPath
		If Request.ServerVariables("SERVER_PORT") = "80" Then
			GetAutoDomain = Request.ServerVariables("SERVER_NAME")
		Else
			GetAutoDomain = Request.ServerVariables("SERVER_NAME") & ":" & Request.ServerVariables("SERVER_PORT")
		End If
		 If Instr(UCASE(GetAutoDomain),"/W3SVC")<>0 Then
			   GetAutoDomain=Left(GetAutoDomain,Instr(GetAutoDomain,"/W3SVC"))
		 End If
		 GetAutoDomain = "http://" & GetAutoDomain
	End Function
	
	Function CutFixContent(ByVal str, ByVal start, ByVal last, ByVal n)
		Dim strTemp
		On Error Resume Next
		If InStr(str, start) > 0 Then
			Select Case n
			Case 0  '���Ҷ���ȡ����ȡǰ�棩��ȥ���ؼ��֣�
				strTemp = Right(str, Len(str) - InStr(str, start) - Len(start) + 1)
				strTemp = Left(strTemp, InStr(strTemp, last) - 1)
			Case Else  '���Ҷ���ȡ����ȡǰ�棩�������ؼ��֣�
				strTemp = Right(str, Len(str) - InStr(str, start) + 1)
				strTemp = Left(strTemp, InStr(strTemp, last) + Len(last) - 1)
			End Select
		Else
			strTemp = ""
		End If
		CutFixContent = strTemp
	End Function
	
	'ȡ��Tag֮���ѭ����
	Function GetTagLoop(ByVal Content)
			Dim regEx, Matches, Match, LoopStr
			Set regEx = New RegExp
			regEx.Pattern = "{Tag([\s\S]*?):(.+?)}"
			regEx.IgnoreCase = True
			regEx.Global = True
			Set Matches = regEx.Execute(Content)
			For Each Match In Matches
				Content=Replace(Content,Match.Value,"")
				Content=Replace(Content,"{/Tag}","")
			Next
			GetTagLoop=Content
    End Function
	
	
	'==================================================
	'��������ScriptHtml
	'��  �ã�����html���
	'��  ����ConStr ------ Ҫ���˵��ַ���
	'==================================================
	Function ScriptHtml(ByVal Constr, TagName, FType)
			Dim re
			Set re = New RegExp
			re.IgnoreCase = True
			re.Global = True
			Select Case FType
			Case 1
			   re.Pattern = "<" & TagName & "([^>])*>"
			   Constr = re.Replace(Constr, "")
			Case 2
			   re.Pattern = "<" & TagName & "([^>])*>.*?</" & TagName & "([^>])*>"
			   Constr = re.Replace(Constr, "")
			Case 3
			   re.Pattern = "<" & TagName & "([^>])*>"
			   Constr = re.Replace(Constr, "")
			   re.Pattern = "</" & TagName & "([^>])*>"
			   Constr = re.Replace(Constr, "")
			End Select
			ScriptHtml = Constr
			Set re = Nothing
	End Function
	

	'*************************************************************************
	'��������gotTopic
	'��  �ã����ַ���������һ���������ַ���Ӣ����һ���ַ�
	'��  ����str   ----ԭ�ַ���
	'       strlen ----��ȡ����
	'����ֵ����ȡ����ַ���
	'*************************************************************************
	Public Function GotTopic(ByVal Str, ByVal strlen)
		If Str = "" OR IsNull(Str) Then GotTopic = "":Exit Function
		If strlen=0 Then GotTopic=Str:Exit Function
		Dim l, T, c, I, strTemp
		Str = Replace(Replace(Replace(Replace(Str, "&nbsp;", " "), "&quot;", Chr(34)), "&gt;", ">"), "&lt;", "<")
		l = Len(Str)
		T = 0
		strTemp = Str
		strlen = CLng(strlen)
		For I = 1 To l
			c = Abs(Ascw(Mid(Str, I, 1)))
			If c > 255 Then
				T = T + 2
			Else
				T = T + 1
			End If
			If T >= strlen Then
				strTemp = Left(Str, I)
				Exit For
			End If
		Next
		If strTemp <> Str Then	strTemp = strTemp
		GotTopic = Replace(Replace(Replace(Replace(strTemp, " ", "&nbsp;"), Chr(34), "&quot;"), ">", "&gt;"), "<", "&lt;")
	End Function
	
	'**************************************************
	'��������ListTitle
	'��  �ã�ȡ����
	'��  ����TitleStr ����, TitleNum ȡ�ַ���
	'����ֵ��������ֽ������
	'**************************************************
	Public Function ListTitle(TitleStr, TitleNum)
		  Dim LeftStr, RightStr
			ListTitle = Trim(GotTopic(Trim(TitleStr), TitleNum))
			If Len(ListTitle) > CInt(TitleNum / 2) Then
			  LeftStr = GotTopic(ListTitle, CInt(TitleNum / 2))
			  RightStr = Mid(ListTitle, Len(LeftStr) + 1)
			  ListTitle = LeftStr & "<br>" & RightStr
			End If
	 End Function
	Function ListTitle1(TitleStr, TitleNum)
		   Dim ClsTitleStr, ClsTitleNum, I, J, ClsTempNum, k, ClsTitleStrResult, LeftStr, RightStr
			   ClsTitleNum = CInt(TitleNum)
			   ClsTempNum = Len(CStr(TitleStr))
			   If ClsTitleNum > ClsTempNum Then
				   ClsTitleNum = ClsTempNum
			   End If
			   ClsTitleStr = Left(CStr(TitleStr), ClsTitleNum)
			   Dim TempStr
			   For I = 1 To ClsTitleNum - 1
				   TempStr = TempStr & Mid(ClsTitleStr, I, 1) & "<br />"
			   Next
			   TempStr = TempStr & Right(ClsTitleStr, 1)
			   ListTitle1 = TempStr
	End Function

	'**************************************************
	'��������GetIP
	'��  �ã�ȡ����ȷ��IP
	'����ֵ��IP�ַ���
	'**************************************************
	Public Function GetIP() 
		Dim strIPAddr 
		If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" Or InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then 
			strIPAddr = Request.ServerVariables("REMOTE_ADDR") 
		ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then 
			strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",")-1) 
		ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then 
			strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";")-1)
		Else 
			strIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR") 
		End If 
		getIP = Checkstr(Trim(Mid(strIPAddr, 1, 30)))
	End Function
	Public Function Checkstr(Str)
		If Isnull(Str) Then
			CheckStr = ""
			Exit Function 
		End If
		Str = Replace(Str,Chr(0),"")
		CheckStr = Replace(Str,"'","''")
	End Function
	'================================================
	'��������URLDecode
	'��  �ã�URL����
	'================================================
	Function URLDecode(ByVal urlcode)
		Dim start,final,length,char,i,butf8,pass
		Dim leftstr,rightstr,finalstr
		Dim b0,b1,bx,blength,position,u,utf8
		On Error Resume Next
	
		b0 = Array(192,224,240,248,252,254)
		urlcode = Replace(urlcode,"+"," ")
		pass = 0
		utf8 = -1
	
		length = Len(urlcode) : start = InStr(urlcode,"%") : final = InStrRev(urlcode,"%")
		If start = 0 Or length < 3 Then URLDecode = urlcode : Exit Function
		leftstr = Left(urlcode,start - 1) : rightstr = Right(urlcode,length - 2 - final)
	
		For i = start To final
			char = Mid(urlcode,i,1)
			If char = "%" Then
				bx = URLDecode_Hex(Mid(urlcode,i + 1,2))
				If bx > 31 And bx < 128 Then
					i = i + 2
					finalstr = finalstr & ChrW(bx)
				ElseIf bx > 127 Then
					i = i + 2
					If utf8 < 0 Then
						butf8 = 1 : blength = -1 : b1 = bx
						For position = 4 To 0 Step -1
							If b1 >= b0(position) And b1 < b0(position + 1) Then
								blength = position
								Exit For
							End If
						Next
						If blength > -1 Then
							For position = 0 To blength
								b1 = URLDecode_Hex(Mid(urlcode,i + position * 3 + 2,2))
								If b1 < 128 Or b1 > 191 Then butf8 = 0 : Exit For
							Next
						Else
							butf8 = 0
						End If
						If butf8 = 1 And blength = 0 Then butf8 = -2
						If butf8 > -1 And utf8 = -2 Then i = start - 1 : finalstr = "" : pass = 1
						utf8 = butf8
					End If
					If pass = 0 Then
						If utf8 = 1 Then
							b1 = bx : u = 0 : blength = -1
							For position = 4 To 0 Step -1
								If b1 >= b0(position) And b1 < b0(position + 1) Then
									blength = position
									b1 = (b1 xOr b0(position)) * 64 ^ (position + 1)
									Exit For
								End If
							Next
							If blength > -1 Then
								For position = 0 To blength
									bx = URLDecode_Hex(Mid(urlcode,i + 2,2)) : i = i + 3
									If bx < 128 Or bx > 191 Then u = 0 : Exit For
									u = u + (bx And 63) * 64 ^ (blength - position)
								Next
								If u > 0 Then finalstr = finalstr & ChrW(b1 + u)
							End If
						Else
							b1 = bx * &h100 : u = 0
							bx = URLDecode_Hex(Mid(urlcode,i + 2,2))
							If bx > 0 Then
								u = b1 + bx
								i = i + 3
							Else
								If Left(urlcode,1) = "%" Then
									u = b1 + Asc(Mid(urlcode,i + 3,1))
									i = i + 2
								Else
									u = b1 + Asc(Mid(urlcode,i + 1,1))
									i = i + 1
								End If
							End If
							finalstr = finalstr & Chr(u)
						End If
					Else
						pass = 0
					End If
				End If
			Else
				finalstr = finalstr & char
			End If
		Next
		URLDecode = leftstr & finalstr & rightstr
	End Function
	
Function URLDecode_Hex(ByVal h)
	On Error Resume Next
	h = "&h" & Trim(h) : URLDecode_Hex = -1
	If Len(h) <> 4 Then Exit Function
	If isNumeric(h) Then URLDecode_Hex = cInt(h)
End Function
	'**************************************************
	'��������R
	'��  �ã����˷Ƿ���SQL�ַ�
	'��  ����strChar-----Ҫ���˵��ַ�
	'����ֵ�����˺���ַ�
	'**************************************************
	Public Function R(strChar)
		If strChar = "" Or IsNull(strChar) Then R = "":Exit Function
		Dim strBadChar, arrBadChar, tempChar, I
		'strBadChar = "$,#,',%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & "," & Chr(0) & ""
		strBadChar = "+,',--,%,^,&,?,(,),<,>,[,],{,},/,\,;,:," & Chr(34) & "," & Chr(0) & ""
		arrBadChar = Split(strBadChar, ",")
		tempChar = strChar
		For I = 0 To UBound(arrBadChar)
			tempChar = Replace(tempChar, arrBadChar(I), "")
		Next
		tempChar = Replace(tempChar, "@@", "@")
		R = tempChar
	End Function
	'����xss
	Function CheckXSS(ByVal strCode)
		Dim Re
		Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
		re.Pattern="<.[^>]*(style).>"
		strCode = re.Replace(strCode, "")
		re.Pattern="<(a.[^>]*|\/a|li|br|B|\/li|\/B|font.[^>]*|\/font)>"
		strCode=re.Replace(strCode,"[$1]")
		strCode=Replace(Replace(strCode, "<", "&lt;"), ">", "&gt;")
		re.Pattern="\[(a.[^\]]*|\/a|li|br|B|\/li|\/B|font.[^\]]*|\/font)\]"
		strCode=re.Replace(strCode,"<$1>")
		re.Pattern="<.[^>]*(on(load|click|dbclick|mouseover|mouseout|mousedown|mouseup|mousewheel|keydown|submit|change|focus)).>"
		strCode = re.Replace(strCode, "")
		Set Re=Nothing
		CheckXSS=strCode
   End Function
	
	Function FilterIDs(byval strIDs)
	Dim arrIDs,i,strReturn
	strIDs=Trim(strIDs)
	If Len(strIDs)=0  Then Exit Function
	arrIDs=Split(strIDs,",")
	For i=0 To Ubound(arrIds)
		If ChkClng(Trim(arrIDs(i)))<>0 Then
			strReturn=strReturn & "," & Int(arrIDs(i))
		End If
	Next
	If Left(strReturn,1)="," Then strReturn=Right(strReturn,Len(strReturn)-1)
	FilterIDs=strReturn
	End Function
	'********************************************
	'��������IsValidEmail
	'��  �ã����Email��ַ�Ϸ���
	'��  ����email ----Ҫ����Email��ַ
	'����ֵ��True  ----Email��ַ�Ϸ�
	'       False ----Email��ַ���Ϸ�
	'********************************************
	Public Function IsValidEmail(Email)
		Dim names, name, I, c
		IsValidEmail = True
		names = Split(Email, "@")
		If UBound(names) <> 1 Then IsValidEmail = False: Exit Function
		For Each name In names
			If Len(name) <= 0 Then IsValidEmail = False:Exit Function
			For I = 1 To Len(name)
				c = LCase(Mid(name, I, 1))
				If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then IsValidEmail = False:Exit Function
		   Next
		   If Left(name, 1) = "." Or Right(name, 1) = "." Then IsValidEmail = False:Exit Function
		Next
		If InStr(names(1), ".") <= 0 Then IsValidEmail = False:Exit Function
		I = Len(names(1)) - InStrRev(names(1), ".")
		If I <> 2 And I <> 3 Then IsValidEmail = False:Exit Function
		If InStr(Email, "..") > 0 Then IsValidEmail = False
	End Function
	'**************************************************
	'��������strLength
	'��  �ã����ַ������ȡ������������ַ���Ӣ����һ���ַ���
	'��  ����str  ----Ҫ�󳤶ȵ��ַ���
	'����ֵ���ַ�������
	'**************************************************
	Public Function strLength(Str)
		On Error Resume Next
		Dim WINNT_CHINESE:WINNT_CHINESE = (Len("�й�") = 2)
		If WINNT_CHINESE Then
			Dim l, T, c,I
			l = Len(Str)
			T = l
			For I = 1 To l
				c = Asc(Mid(Str, I, 1))
				If c < 0 Then c = c + 65536
				If c > 255 Then
					T = T + 1
				End If
			Next
			strLength = T
		Else
			strLength = Len(Str)
		End If
		If Err.Number <> 0 Then Err.Clear
	End Function

	'**************************************************
	'������: GetFolderPath
	'�� ��:ȡ��Ŀ¼Url
	'�� ��: FolderIDĿ¼��ID
	'**************************************************
	Public Function GetFolderPath(FolderID)
			If Not IsObject(Application(SiteSN&"_classpath")) Then
		     Dim Folder,ClassPurview,ChannelFsoHtmlTF,Node,K,SQL,RS
			 Set  Application(SiteSN&"_classpath")=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		     Application(SiteSN&"_classpath").appendChild( Application(SiteSN&"_classpath").createElement("xml"))
              Set RS=Server.CreateObject("ADODB.RECORDSET")
			  RS.Open "Select C.ClassID,C.ChannelID,TN,Folder,FolderDomain,ClassPurview,FsoHtmlTF,StaticTF,C.ID,ClassType,M.FsoClassListRule,M.FsoClassPreTag,FolderFsoIndex,StaticTF From KS_Class C inner join KS_Channel M On C.ChannelID=M.ChannelID Order BY FolderOrder", Conn, 1, 1
			  If RS.Eof And RS.Bof Then RS.Close:Set RS=Nothing:Exit Function
			  SQL=RS.GetRows(-1):RS.Close:Set RS=Nothing
			  For K=0 To Ubound(SQL,2)
					       ClassPurview=SQL(5,K)
						   ChannelFsoHtmlTF=SQL(6,K)
						   If SQL(9,K)="2" Then
						    GetFolderPath=SQL(3,K)
						   Else
							   If Trim(SQL(4,K)) <> "" And SQL(2,K) = "0" Then
								   IF ClassPurview=2 Or ChannelFsoHtmlTF=0 Or ChannelFsoHtmlTF=2 Then
									 GetFolderPath= GetChannelNoHtmlUrl(SQL(7,K),SQL(0,K))
								   Else
									 GetFolderPath=Trim(SQL(4,K))
								   End If
							   ElseIf Trim(SQL(4,K)) <> "" Then
								  Folder = Trim(SQL(3,K))
								  Folder = Right(Mid(Folder, InStr(Folder, "/")), Len(Mid(Folder, InStr(Folder, "/"))) - 1)
								   IF ClassPurview=2 Or ChannelFsoHtmlTF=0 Or ChannelFsoHtmlTF=3 Then
									 GetFolderPath= Trim(SQL(4,K)) & GetChannelNoHtmlUrl(SQL(7,K),SQL(0,K))
								   Else
									 GetFolderPath= Trim(SQL(4,K)) & Folder
								   End If
							  Else
							       IF ClassPurview=2 Or ChannelFsoHtmlTF=0 Or ChannelFsoHtmlTF=2 Then
									 GetFolderPath= GetChannelNoHtmlUrl(SQL(7,K),SQL(0,K))
								   Else
									 	 GetFolderPath= GetChannelDomain(SQL(1,K)) 
										 If SQL(9,K)="3" Then
										  GetFolderPath= GetChannelDomain(SQL(1,K)) & SQL(3,K)
										 Else
											 Select Case SQL(10,K)
											   Case "1":GetFolderPath= GetChannelDomain(SQL(1,K)) & SQL(3,K)
											   Case "2":GetFolderPath= GetChannelDomain(SQL(1,K)) & SQL(11,K) &"_" & SQL(0,K) &Mid(Trim(SQL(12,K)), InStrRev(Trim(SQL(12,K)), ".")) '�������չ��
											   Case "3":
												 GetFolderPath= GetChannelDomain(SQL(1,K)) & Split(SQL(3,K),"/")(0) & "/"
												 If SQL(2,K) <> "0" Then GetFolderPath= GetFolderPath & SQL(11,K) &"_" & SQL(0,K) &Mid(Trim(SQL(12,K)), InStrRev(Trim(SQL(12,K)), ".")) '�������չ��
											 End Select
                                         End If
								   End If
							  End If
						 End If
		            Set Node=Application(SiteSN&"_classpath").documentElement.appendChild(Application(SiteSN&"_classpath").createNode(1,"classpath",""))
			        Node.attributes.setNamedItem(Application(SiteSN&"_classpath").createNode(2,"classid","")).text=SQL(8,K)
			        Node.text=GetFolderPath
               Next			
     End If
	 Dim NodeText:Set NodeText=Application(SiteSN&"_classpath").documentElement.selectSingleNode("classpath[@classid=" & FolderID & "]")
	 If Not NodeText Is Nothing Then GetFolderPath=NodeText.text
	End Function
	'************************************************************************
	'������: GetClassNP
	'�� ��: ȡ��Ŀ¼���Ʋ���������
	'�� ��: ClassIDĿ¼��ID	          
	'*************************************************************************
	Function GetClassNP(ClassID)
		If Not IsObject(Application(SiteSN&"_classnamepath")) Then
		    Dim Folder,ClassPurview,ChannelFsoHtmlTF,Node,K,SQL,RS
			Dim OpenTypeStr:OpenTypeStr=" target=""_blank"""
			Set  Application(SiteSN&"_classnamepath")=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		    Application(SiteSN&"_classnamepath").appendChild( Application(SiteSN&"_classnamepath").createElement("xml"))
              Set RS=Server.CreateObject("ADODB.RECORDSET")
			  RS.Open "Select ID,FolderName From KS_Class Order BY FolderOrder", Conn, 1, 1
			  If RS.Eof And RS.Bof Then RS.Close:Set RS=Nothing:Exit Function
			  SQL=RS.GetRows(-1):RS.Close:Set RS=Nothing
			  For K=0 To Ubound(SQL,2)
		            Set Node=Application(SiteSN&"_classnamepath").documentElement.appendChild(Application(SiteSN&"_classnamepath").createNode(1,"classnamepath",""))
			        Node.attributes.setNamedItem(Application(SiteSN&"_classnamepath").createNode(2,"classid","")).text=SQL(0,K)
			        Node.text="<a href=""" & GetFolderPath(SQL(0,K)) & """" & OpenTypeStr & ">" & Trim(SQL(1,K)) & "</a>"
              Next			
     End If
	 Dim NodeText:Set NodeText=Application(SiteSN&"_classnamepath").documentElement.selectSingleNode("classnamepath[@classid=" & ClassID & "]")
	 If Not NodeText Is Nothing Then GetClassNP=NodeText.text
	End Function
	
	'�滻����ҳ���ɹ���
	 Function LoadFsoContentRule(ChannelID,ClassID)
	    on error resume next
		Dim FsoContentRule:FsoContentRule=C_S(ChannelID,43)
        FsoContentRule=Replace(FsoContentRule,"{$ChannelEname}",Split(C_C(ClassID,2),"/")(0))
        FsoContentRule=Replace(FsoContentRule,"{$ClassDir}",C_C(ClassID,2))
        FsoContentRule=Replace(FsoContentRule,"{$ClassID}",C_C(ClassID,9))
        FsoContentRule=Replace(FsoContentRule,"{$ClassEname}",Split(C_C(ClassID,2), "/")(C_C(ClassID,10)- 1))
		FsoContentRule=Replace(Setting(3) & C_S(ChannelID,8),"//","/") & FsoContentRule
		LoadFsoContentRule=FsoContentRule
	 End Function
     Function LoadInfoUrl(ChannelID,ClassID,Fname)
	   If C_C(ClassID,4)<>"" Then
	    LoadInfoUrl=GetFolderPath(ClassID) & Fname
	   Else
	    LoadInfoUrl=Setting(2) & LoadFsoContentRule(ChannelID,ClassID) & Fname
	   End If
	 End Function
		'----------------------------------------------------------------------------------------------------------------------
		'������: GetSpecialPath
		'�� ��: ȡ��ר��Ŀ¼Url
		'�� ��: SpecialrRS
		'-----------------------------------------------------------------------------------------------------------------------
		Public Function GetSpecialPath(SpecialID,SpecialEname,FsoSpecialIndex)
		      Dim SpecialDir:SpecialDir = Setting(95)
			  If Left(SpecialDir, 1) = "/" Or Left(SpecialDir, 1) = "\" Then SpecialDir = Right(SpecialDir, Len(SpecialDir) - 1)
			  If Setting(78)="0" Then
				GetSpecialPath=GetDomain & "Special.asp?ID=" & SpecialID
			  Else
				 GetSpecialPath = GetDomain & SpecialDir & SpecialEname & "/" & FsoSpecialIndex
              End iF
		End Function
		'----------------------------------------------------------------------------------------------------------------------
		'������: GetFolderSpecialPath
		'�� ��: ȡ����Ŀר�����Url
		'�� ��: ClassIDĿ¼��ID,FullPathFlag�Ƿ�����·��(ȡ��Ŀ��ҳ���),����ר����ҳ
		'-----------------------------------------------------------------------------------------------------------------------
		Function GetFolderSpecialPath(ClassID, FullPathFlag)
		   Dim SpecialDir:SpecialDir =Setting(95)
		    If Left(SpecialDir, 1) = "/" Or Left(SpecialDir, 1) = "\" Then SpecialDir = Right(SpecialDir, Len(SpecialDir) - 1)
		     IF Setting(78)="0" Then
			     GetFolderSpecialPath = GetDomain &"SpecialList.asp?ClassID="&ClassID
			 Else
			  Dim RS:Set RS=Conn.Execute("Select ClassEname,FsoIndex From KS_SpecialClass Where ClassID=" & ChkClng(ClassID))
			  If RS.Eof Then
			   GetFolderSpecialPath = GetDomain &"SpecialList.asp?ClassID="&ClassID
			  Else
			    GetFolderSpecialPath = GetDomain & SpecialDir & RS(0) & "/"
			    If FullPathFlag = True Then
			     GetFolderSpecialPath=GetFolderSpecialPath & RS(1)
			    End If
              	RS.Close:Set RS = Nothing
			 End IF
			End If
		End Function
		'ȡ����Ŀ������URL
		Public Function GetChannelNoHtmlUrl(StaticTF,ClassID)
		     If StaticTF=0 Then
		      GetChannelNoHtmlUrl=GetDomain &"Item/list.asp?id=" & ClassID
			 ElseIf StaticTF=2 Then
		      GetChannelNoHtmlUrl=GetDomain & GCls.StaticPreList & "-" & ClassID & GCls.StaticExtension
			 Else
		      GetChannelNoHtmlUrl=GetDomain & "?" & GCls.StaticPreList & "-" & ClassID & GCls.StaticExtension
			 End If
		End Function
		
		
		Public Function GetItemURL(ByVal ChannelID,ByVal Tid,ByVal InfoID,ByVal Fname)
		  IF Not Isnumeric(ChannelID) Then GetItemURL="#":Exit Function
		  If  C_S(ChannelID,7)=0 Then 
		        if C_S(ChannelID,48)=0 Then
				 GetItemURL=GetDomain & "Item/Show.asp?m=" & ChannelID & "&d=" &InfoID
				ElseIf C_S(ChannelID,48)=2 Then
				 GetItemURL=GetDomain & GCls.StaticPreContent & "-" & InfoID & "-"& ChannelID & GCls.StaticExtension
				Else
				 GetItemURL=GetDomain & "?" & GCls.StaticPreContent & "-" & InfoID & "-"& ChannelID & GCls.StaticExtension
				End If
		  Else
				GetItemURL=LoadInfoUrl(ChannelID,TID,Fname)
		  End If
		End Function
		
		'ȡ��HTML
		Public Function LoseHtml(ByVal ContentStr)
		    On Error Resume Next
			Dim TempLoseStr, regEx
			If ContentStr="" Or ContentStr=Null Then Exit Function
			TempLoseStr = HtmlCode(ContentStr)
			Set regEx = New RegExp
			regEx.Pattern = "<\/*[^<>]*>"
			regEx.IgnoreCase = True
			regEx.Global = True
			TempLoseStr = regEx.Replace(TempLoseStr, "")
			LoseHtml = TempLoseStr
		End Function
				                 '---------------------------------------------------------------------------------------------------
		'������: G_O_T_S
		'�� ��:ȡ�ô�����
		'�� ��: OpenType ȡtrueʱ,�´��ڴ�
		'--------------------------------------------------------------------------------------------
		Function G_O_T_S(OpenType)
			  If OpenType = "" Or OpenType = False Then
				G_O_T_S = ""
			  ElseIf OpenType = True Then
				G_O_T_S = " target=""_blank"""
			  Else
				G_O_T_S = " target=""" & OpenType & """"
			  End If
		End Function
		'--------------------------------------------------------------------------------------------------
		'������: GetCss
		'�� ��:ȡ����ʽ
		'�� ��: CssName��ʽ����
		'--------------------------------------------------------------------------------------------
		Function GetCss(CssName)
			 If CssName = "" Or IsNull(CssName) Then  GetCss = "" Else GetCss = " class=""" & CssName & """"
		End Function
				
		'ȡ��CSS��ID
		Function GetCssID(ID)
		  If ID="" Then GetCssID="" Else GetCssID=" id=""" & ID & """"
		End Function  
		'-------------------------------------------------------------------------------------------------------------
		'������: G_R_H
		'�� ��:ȡ�õ�Ԫ���о�
		'�� ��: RowHeight Ĭ���о�
		'-----------------------------------------------------------------------------------------------------------
		Function G_R_H(RowHeight)
			If IsNumeric(RowHeight) Then G_R_H = RowHeight Else G_R_H = 20
		End Function
	'----------------------------------------------------------------------------------------------------------------------------
		'������:GetMenuBg
		'�� ��:ȡ�ñ�ͷ����
		'�� ��: MenuBGType ���� 1 ȡ����ͼƬ 0 ȡ������ɫ, MenuBg ������ɫ��ֵ ��#CCCCCC �� /Upfies/TITLE_BG.GIF ,ColNumber����
   '---------------------------------------------------------------------------------------------------------------------------
		Function GetMenuBg(MenuBgType, MenuBg, ColNumber)
		  If MenuBgType = 0 Then
			 If MenuBg = "" Then GetMenuBg = "" Else GetMenuBg = MenuBg
		  Else
			 If MenuBg = "" Then
			   GetMenuBg = "url(" & GetDomain & "Images/Default/MenuBg" & ColNumber & ".Gif)"
			 Else
			   If Left(MenuBg, 1) = "/" Or Left(MenuBg, 1) = "\" Then MenuBg = Right(MenuBg, Len(MenuBg) - 1)
			   If LCase(Left(MenuBg, 4)) = "http" Then MenuBg = MenuBg Else MenuBg = GetDomain & MenuBg
			   GetMenuBg = "url(" & MenuBg & ")"
			 End If
		  End If
		End Function
	'----------------------------------------------------------------------------------------------------------------------------
		'������:GetPhotoBorder
		'�� ��: ȡ��ͼƬ�ı߿�
		'�� ��: BorderType ���� 1 ȡ͸��ͼƬ�߿� 0 ȡ��ɫ�߿�, Border ������ɫ��ֵ ��#CCCCCC �� /Upfies/TITLE_BG.GIF ,ColNumber����
		'----------------------------------------------------------------------------------------------------------------------------
		Function GetPhotoBorder(LinkPhotoStr, BorderType, Border)
				   Dim bgColorStr
				   If Trim(Border) = "" Then
					 GetPhotoBorder = LinkPhotoStr:Exit Function
				   Else
					 If BorderType = 0 Then
					  bgColorStr = " bgcolor=""" & Border & """"
					   GetPhotoBorder = "<table borderColor=#ffffff cellSpacing=1 cellPadding=1 align=center " & bgColorStr & " border=0>" & vbCrLf
					  GetPhotoBorder = GetPhotoBorder & "  <tr>" & vbCrLf
					  GetPhotoBorder = GetPhotoBorder & "   <td valign=center align=middle bgColor=#ffffff>" & LinkPhotoStr & "</td>" & vbCrLf
					  GetPhotoBorder = GetPhotoBorder & "  </tr>" & vbCrLf
					  GetPhotoBorder = GetPhotoBorder & "</table>" & vbCrLf
					  Exit Function
					 Else
					   If Left(Border, 1) = "/" Or Left(Border, 1) = "\" Then Border = Right(Border, Len(Border) - 1)
					   If LCase(Left(Border, 4)) = "http" Then
						 Border = Border
					   Else
						 Border = GetDomain & Border
					   End If
						bgColorStr = " style=""background:url(" & Border & ") #FFF no-repeat;"""
					  GetPhotoBorder = "<table borderColor=#ffffff cellSpacing=0 cellPadding=0 align=center " & bgColorStr & " border=0>" & vbCrLf
					  GetPhotoBorder = GetPhotoBorder & "  <tr>" & vbCrLf
					  GetPhotoBorder = GetPhotoBorder & "   <td valign=center align=middle>" & LinkPhotoStr & "</td>" & vbCrLf
					  GetPhotoBorder = GetPhotoBorder & "  </tr>" & vbCrLf
					  GetPhotoBorder = GetPhotoBorder & "</table>" & vbCrLf
					   End If
				   End If
			  End Function
		'--------------------------------------------------------------------------------------------------------------------
		'������: GetNavi
		'�� ��: ȡ�õ���ֵ
		'�� ��: NaviType ��������, NaviStr����ֵ
		'---------------------------------------------------------------------------------------------------------------
		Function GetNavi(NaviType, NaviStr)
		 If NaviType = "0" Then
			 If NaviStr = "" Then GetNavi = "" Else GetNavi = NaviStr
		 ElseIf NaviType = "1" Then
		   If NaviStr <> "" Then  GetNavi = "<img src=""" & NaviStr & """ alt="""" border=""0""/>"
		 Else
			 GetNavi = ""
		 End If
		End Function
		'---------------------------------------------------------------
		'������:GetDateStr
		'���ã�ȡ���ڵ���ʽ
		'������AddDate,DateRule,DateAlign,DateCssStr,ByRef ColSpanNum
		'---------------------------------------------------------------
		Function GetDateStr(ChannelID,AddDate,DateRule,DateAlign,DateCssStr,ByVal ColNumber,ByRef ColSpanNum)
		       If CStr(DateRule) <> "0" And CStr("DateRule") <> "" Then
					  	Dim NowDate,NowFormatStr
						If DateDiff("d",AddDate,Now())-ChkClng(C_S(ChannelID,47))<0 Then NowFormatStr=" style=""color:red""" Else  NowFormatStr=""
						If Lcase(DateAlign)="left" Then
							GetDateStr="&nbsp;<span" & NowFormatStr & DateCssStr &">" & DateFormat(AddDate, DateRule) & "</span>"
							ColSpanNum = ColNumber+1
						Else
							GetDateStr="</td><td width=""*"" nowrap align=" & DateAlign & "><span" & NowFormatStr & DateCssStr & ">" & DateFormat(AddDate, DateRule) & "</span>"
							ColSpanNum = ColNumber+2
						End If
				Else
				GetDateStr="":ColSpanNum = ColNumber+1
				End If
		End Function
		'ȡ��������ʽ(div+css)
		Function GetDCDateStr(ChannelID,AddDate,DateRule,DateCssStr)
			 If CStr(DateRule) <> "0" And CStr("DateRule") <> "" Then
					  	Dim NowFormatStr
						If DateDiff("d",AddDate,Now())-ChkClng(C_S(ChannelID,47))<0 Then
						 NowFormatStr=" style=""color:red"""
						Else
						 NowFormatStr=""
						End If
						GetDCDateStr="&nbsp;<span" & NowFormatStr & DateCssStr &">" & DateFormat(AddDate, DateRule) & "</span>"
				Else
				GetDCDateStr=""
				End If
        End Function
		
		  '���ظ�ʽ�����ʱ��
		   Function GetTimeFormat(DateTime)
		      if DateDiff("n",DateTime,now)<5 then
			   GetTimeFormat="�ո�"
			  elseif DateDiff("n",DateTime,now)<60 then
			   GetTimeFormat=DateDiff("n",DateTime,now) & " ����ǰ"
			  elseif DateDiff("h",DateTime,now)<12 Then
			   GetTimeFormat=DateDiff("h",DateTime,now) & " Сʱǰ"
			  else
			   GetTimeFormat=formatdatetime(DateTime,2)
			  end if
		   End Function
				'----------------------------------------------------------------------------------------------------------------------------
		'������:DateFormat
		'�� ��:���ڸ�ʽ����
		'�� ��: DateStr����, Typesת������		'----------------------------------------------------------------------------------------------------------------------------
		Function DateFormat(DateStr, Types)
			Dim DateString
			If IsDate(DateStr) = False Then
				DateFormat = "":Exit Function
			End If
			Select Case CStr(Types)
			  Case "0"
				DateFormat = ""
				Exit Function
			  Case 1,21,41
			      DateString=Year(DateStr) & "-" & Right("0" & Month(DateStr), 2) & "-" & Right("0" & Day(DateStr), 2)
			      if Types=21 then
				   DateString = "(" & DateString &")"
				  elseIf Types=41 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 2,22,42
			      DateString=Year(DateStr) & "." & Right("0" & Month(DateStr), 2) & "." & Right("0" & Day(DateStr), 2)
			      if Types=22 then
				   DateString = "(" & DateString &")"
				  elseIf Types=42 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 3,23,43
			      DateString=Year(DateStr) & "/" & Right("0" & Month(DateStr), 2) & "/" & Right("0" & Day(DateStr), 2)
			      if Types=23 then
				   DateString = "(" & DateString &")"
				  elseIf Types=43 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 4,24,44
			      DateString=Right("0" & Month(DateStr), 2) & "/" & Right("0" & Day(DateStr), 2) & "/" & Year(DateStr)
			      if Types=24 then
				   DateString = "(" & DateString &")"
				  elseIf Types=44 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 5,25,45
				  DateString = Year(DateStr) & "��" & Right("0" & Month(DateStr), 2) & "��"
			      if Types=25 then
				   DateString = "(" & DateString &")"
				  elseIf Types=45 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 6,26,46
				  DateString = Year(DateStr) & "��" & Right("0" & Month(DateStr), 2) & "��" & Right("0" & Day(DateStr), 2) & "��"
			      if Types=26 then
				   DateString = "(" & DateString &")"
				  elseIf Types=46 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 7,27,47
				  DateString = Right("0" & Month(DateStr), 2) & "." & Right("0" & Day(DateStr), 2) & "." & Year(DateStr)
			      if Types=27 then
				   DateString = "(" & DateString &")"
				  elseIf Types=47 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 8,28,48
				  DateString = Right("0" & Month(DateStr), 2) & "-" & Right("0" & Day(DateStr), 2) & "-" & Year(DateStr)
				  if Types=28 then
				   DateString = "(" & DateString &")"
				  elseIf Types=48 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 9,29,49
				  DateString = Right("0" & Month(DateStr), 2) & "/" & Right("0" & Day(DateStr), 2)
				  if Types=29 then
				   DateString = "(" & DateString &")"
				  elseIf Types=49 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 10,30,50
				  DateString = Right("0" & Month(DateStr), 2) & "." & Right("0" & Day(DateStr), 2)
			      if Types=30 then
				   DateString = "(" & DateString &")"
				  elseIf Types=50 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 11,31,51
				  DateString = Right("0" & Month(DateStr), 2) & "��" & Right("0" & Day(DateStr), 2) & "��"
			      if Types=31 then
				   DateString = "(" & DateString &")"
				  elseIf Types=51 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 12,32,52
				  DateString = Right("0" & Day(DateStr), 2) & "��" & Right("0" & Hour(DateStr), 2) & "ʱ"
				  if Types=32 then
				   DateString = "(" & DateString &")"
				  elseIf Types=52 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 13,33,53
				  DateString = Right("0" & Day(DateStr), 2) & "��" & Right("0" & Hour(DateStr), 2) & "��"
			      if Types=33 then
				   DateString = "(" & DateString &")"
				  elseIf Types=53 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 14,34,54
				  DateString = Right("0" & Hour(DateStr), 2) & "ʱ" & Minute(DateStr) & "��"
				  if Types=34 then
				   DateString = "(" & DateString &")"
				  elseIf Types=54 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 15,35,55
				  DateString = Right("0" & Hour(DateStr), 2) & ":" & Right("0" & Minute(DateStr), 2)
			      if Types=35 then
				   DateString = "(" & DateString &")"
				  elseIf Types=55 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 16,36,56
				  DateString = Right("0" & Month(DateStr), 2) & "-" & Right("0" & Day(DateStr), 2)
				 if Types=36 then
				   DateString = "(" & DateString &")"
				  elseIf Types=56 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case 17,37,57
				  DateString = Right("0" & Month(DateStr), 2) & "/" & Right("0" & Day(DateStr), 2) &" " &Right("0" & Hour(DateStr), 2)&":"&Right("0" & Minute(DateStr), 2)
				  if Types=37 then
				   DateString = "(" & DateString &")"
				  elseIf Types=57 then
				  	DateString = "[" & DateString &"]"
				  end if
			  Case Else
				  DateString = DateStr
			 End Select
			 DateFormat = DateString
		 End Function
		 '----------------------------------------------------------------------------------------------------------------------------
		'������:GetOrigin
		'�� ��:ȡ��������Դ������������
		'�� ��: OriginName����
		'����ֵ: ���� <a href="http://www.xinhua.com" target="_blank">�»���</a>
'----------------------------------------------------------------------------------------------------------------------------
		Function GetOrigin(OriginName)
		  Dim RS: Set RS=Server.CreateObject("ADODB.Recordset")
		  RS.Open "select OriginName,HomePage From KS_Origin Where OriginName='" & Trim(OriginName) & "'", Conn, 1, 1
		  If RS.EOF Then
		   GetOrigin = OriginName
		  Else
		   If RS("HomePage") <> "" And UCase(Trim(RS("HomePage"))) <> "HTTP://" Then
		   GetOrigin = "<a href=""" & Trim(RS("HomePage")) & """ target=""_blank"">" & OriginName & "</a>"
		   Else
			GetOrigin = OriginName
		   End If
		 End If
		 RS.Close:Set RS = Nothing
		End Function
	'----------------------------------------------------------------------------------------------------------------------------
		'������:GetMoreLink
		'�� ��:ȡ�ø�������
		'�� ��: ColNum����, RowHeight�о�, MoreLinkType��������, LinkUrl���ӵ�ַ, OpenTypeStr�Ƿ��´��ڴ�
	'----------------------------------------------------------------------------------------------------------------------------
		Function GetMoreLink(PrintType,ColNum, RowHeight, MoreLinkType, LinkNameStr, LinkUrl, OpenTypeStr)
		   If LinkNameStr = "" Then GetMoreLink = "":Exit Function
	      If PrintType=2 Then
		   If MoreLinkType = "0" Then
			  GetMoreLink = "<li><a href=""" & LinkUrl & """" & OpenTypeStr & " > " & LinkNameStr & "</a></li>"
		   ElseIf MoreLinkType = "1" Then
			  GetMoreLink = "<li><a href=""" & LinkUrl & """" & OpenTypeStr & " > <img src=""" & LinkNameStr & """ border=""0"" align=""absmiddle""/></a></li>"
		   Else
			 GetMoreLink = ""
		   End If
		  Else
			   LinkNameStr = Trim(LinkNameStr):LinkUrl = Trim(LinkUrl)
			   If MoreLinkType = "0" Then
				  GetMoreLink = "<tr><td colspan= """ & ColNum+1 & """ height=""" & RowHeight & """ align=""right""><a href=""" & LinkUrl & """" & OpenTypeStr & " > " & LinkNameStr & "</a></td></tr>"
			   ElseIf MoreLinkType = "1" Then
						GetMoreLink = "<tr><td colspan= """ & ColNum+1 & """ height=""" & RowHeight & """ align=""right""><a href=""" & LinkUrl & """" & OpenTypeStr & " > <img src=""" & LinkNameStr & """ border=""0"" align=""absmiddle""/></a></td></tr>"
				 
			   Else
				 GetMoreLink = ""
			   End If
		  End If
		End Function			
 '----------------------------------------------------------------------------------------------------------------------------
		'������: GetSplitPic
		'�� ��:ȡ�÷ָ�ͼƬ
		'�� ��: ColSpanNum ����, SplitPic ͼƬSRC		'-------------------------------------------------------------------------------------------------------------------------------
		Function GetSplitPic(SplitPic, ColSpanNum)
		     Dim ColStr
			 If SplitPic = "" or IsNull(SplitPic) Then
			   GetSplitPic = ""
			 Else
			   If ColSpanNum>=2 Then ColStr=" colspan=""" & ColSpanNum & """"
			   GetSplitPic = "<tr><td height=""1"""  & ColStr & " background=""" & SplitPic & """ ></td></tr>" & vbcrlf
			 End If
		End Function
	'-------------------------------------------------------------------------------------------------------------------
		'������:GetFolderTid
		'�� ��:ȡ����Ŀ¼��ID����
		'�� ��:  FolderID��Ŀ¼ID
		'����ֵ: ���� 1255555,111111,4444��ID����
   '---------------------------------------------------------------------------------------------------------
		Function GetFolderTid(FolderID)
			GetFolderTid="Select ID From KS_Class Where DelTF=0 AND TS LIKE '%" & FolderID & "%'":Exit Function
		End Function
		'ȡ��ר���ѯ����,Ӧ����Sql����
		Function GetSpecialPara(ChannelID,SpecialID)
			   If SpecialID = "-1" Then
					 If FCls.RefreshType = "Special" Then
					   If ChannelID<>0 Then
						GetSpecialPara=" And ID in(select infoid from ks_specialr where ChannelID=" & ChannelID & " and  SpecialID=" & ChkClng(FCls.CurrSpecialID) & ") "
					   Else
						GetSpecialPara=" And InfoID in(select infoid from ks_specialr r where SpecialID=" & ChkClng(FCls.CurrSpecialID) & " and i.channelid=r.channelid) "
					   End If
					 Else
						 GetSpecialPara = ""
					 End If
			  ElseIf (SpecialID = "" Or SpecialID = "0" Or IsNull(SpecialID))  Then
					 GetSpecialPara = ""
			  Else
			      If ChannelID<>0 Then
			      GetSpecialPara=" And ID in(select infoid from ks_specialr where ChannelID=" & ChannelID & " and SpecialID=" & ChkClng(SpecialID) & ") "
				  Else
			      GetSpecialPara=" And InfoID in(select infoid from ks_specialr r where SpecialID=" & ChkClng(SpecialID) & " and i.channelid=r.channelid) "
				  End If
			  End If
		End Function
		
	'�����ļ����Զ����ֶ�
	Sub LoadFieldToXml()
	  If Not IsObject(Application(SiteSN & "_FeildXml")) then
			Dim Rs:Set Rs = Conn.Execute("Select ChannelID,FieldName,fieldtype From KS_Field Where FieldType=9 or FieldType=10 Order By FieldID")
			Set Application(SiteSN & "_FeildXml")=RsToxml(Rs,"row","FeildXml")
			Set Rs = Nothing
	  End If
	End Sub
		
	'����Թ������ݿ�	
	Sub FileAssociation(ByVal ChannelID,ByVal InfoID,ByVal Content,ByVal Flag)
	  If Flag<>0 Then
	  Conn.Execute("Delete From KS_UploadFiles Where ChannelID=" & ChannelID & " and InfoID=" & InfoID)
	  End If
	  If ChannelID<>0 And ChannelID<1000 Then
	     Dim Node
	     LoadFieldToXml()
		 For Each Node In Application(SiteSN & "_FeildXml").DocumentElement.SelectNodes("row[@channelid=" & ChannelID &" and @fieldtype=9 or @fieldtype=10]")
		    Content=Content & Request(Node.SelectSingleNode("@fieldname").text)
		 Next
	  End If
	  Dim FileLists,I,FileArr
	  FileLists=GetFilesList(ChannelID,Content)
	  If Not IsNul(FileLists) Then
	    FileArr=Split(FileLists,"|")
		For I=0 To Ubound(FileArr)
		 Conn.Execute("Insert Into [KS_UploadFiles](ChannelID,InfoID,FileName) values(" &ChannelID &"," & InfoID &",'" & FileArr(i) & "')")
		Next
	  End If
	End Sub
	
	'�������ݻ�ȡ�ϴ��ļ���
	Public Function GetFilesList(ChannelID,Content)
		Dim re, UpFile, BFU, FileName,SaveFileList,FileExt
		If ChannelID<1000 Then FileExt=ReturnChannelAllowUpFilesType(ChannelID,0) Else FileExt=Setting(7)
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "(\/uploadfiles\/)[^(\/uploadfiles\/)]?(.*?)[.]{1}(" & FileExt & "|wma|mp3)"
		're.Pattern = "(\/uploadfiles\/)[^(\/uploadfiles\/)](.*?)[.]{1}(" & FileExt & "|wma|mp3)"
		Set UpFile = re.Execute(Content)
		Set re = Nothing
		For Each BFU In UpFile
		  If Instr(SaveFileList,BFU)=0 Then
		     if FileName="" then
			  FileName=BFU
			 Else
		      FileName=FileName & "|" & BFU
			 End If
		  End If
		   SaveFileList=SaveFileList & "," & BFU
		Next
		GetFilesList = FileName
     End Function
	
	'**************************************************
	'��������ReturnChannelAllowUpFilesTF
	'��  �ã�����Ƶ�����Ƿ������ϴ��ļ�
	'��  ����ChannelID--Ƶ��ID
	'**************************************************
	Public Function ReturnChannelAllowUpFilesTF(ChannelID)
	  If ChannelID = "" Or Not IsNumeric(ChannelID) Then  ChannelID = 0
	   Dim CRS:Set CRS=Server.CreateObject("ADODB.RECORDSET")
	   CRS.Open "Select UpFilesTF From KS_Channel Where ChannelID=" & ChannelID, Conn, 1, 1
	  If CInt(ChannelID) = 0 Or (CRS.EOF And CRS.BOF) Then  'Ĭ�������ϴ��ļ�
		ReturnChannelAllowUpFilesTF = True
	  Else
		If CRS(0) = 1 Then ReturnChannelAllowUpFilesTF = True	Else ReturnChannelAllowUpFilesTF = False
	  End If
	CRS.Close:Set CRS = Nothing
	End Function
	
	'ȡ�ϴ�Ŀ¼6.0��Ϊ�����ڴ��
	Function GetUpFilesDir()
	   Dim DateFolder:DateFolder=Setting(3) & Setting(91) & Year(Now) & "-" & Right("0"&Month(Now),2)
	   If Setting(96) = "1" Then
		   Dim Ce:Set Ce=new CtoeCls
		   Dim UserFolder:UserFolder=Ce.CTOE(R(C("AdminName")))
		   Set Ce=Nothing
		   If UserFolder<>"" Then DateFolder=DateFolder & "/" & UserFolder
	   End If
	   CreateListFolder(DateFolder)
	   GetUpFilesDir=DateFolder
	End Function
	
	'ȡ�ú�̨���������ֵ��ϴ�Ŀ¼,һ�����ڹ��,�������õ�
	Function GetCommonUpFilesDir()
	  Dim Str
	  If C("SuperTF")="1" Then 
	    Str=Setting(3) & Setting(91)
	  Else
	    Str=GetUpFilesDir()
	  End If
	  If Right(Str,1)="/" Then Str=Left(Str,Len(Str)-1)
	  GetCommonUpFilesDir=Str
	End Function

	'**************************************************
	'��������ReturnChannelAllowUserUpFilesTF
	'��  �ã�����Ƶ���Ƿ������Ա�ϴ��ļ�
	'��  ����ChannelID--Ƶ��ID
	'**************************************************
	Public Function ReturnChannelAllowUserUpFilesTF(ChannelID)
	  If ChannelID = "" Or Not IsNumeric(ChannelID) Then 'Ĭ�������ϴ��ļ�
	  ReturnChannelAllowUserUpFilesTF = True:Exit Function
	  End If
		If C_S(ChannelID,26) = 1 Then
		 ReturnChannelAllowUserUpFilesTF = True
		Else
		 ReturnChannelAllowUserUpFilesTF = False
		End If
	End Function

	'**************************************************
	'��������ReturnChannelUserUpFilesDir
	'��  �ã�����Ƶ��ǰ̨��Ա���ϴ�Ŀ¼
	'��  ����ChannelID--Ƶ��ID,UserFolder-���û������ɵ�Ŀ¼
	'����ֵ��Ŀ¼�ַ���
	'**************************************************
	Public Function ReturnChannelUserUpFilesDir(ChannelID,UserFolder)
	   If HasChinese(UserFolder) Then
	     Dim Ce:Set Ce=new CtoeCls
	     UserFolder="[" & Ce.CTOE(R(UserFolder)) & "]"
	     Set Ce=Nothing
	   End If
	   
	   ChannelID = ChkCLng(ChannelID)
	   If UserFolder="" Then UserFolder="Temp"
	   Select Case ChannelID
	    Case 9999 '�û�ͷ��
		   ReturnChannelUserUpFilesDir=Setting(3)&Setting(91)&"User/" & UserFolder &"/upface/"
		Case 9998,9997 '���
		   ReturnChannelUserUpFilesDir=Setting(3)&Setting(91)&"User/" & UserFolder &"/xc/"
		Case 9996 'Ȧ��ͼƬ
		   ReturnChannelUserUpFilesDir=Setting(3)&Setting(91)&"User/" & UserFolder &"/team/"
		Case 9995 '����
		   ReturnChannelUserUpFilesDir=Setting(3)&Setting(91)&"User/" & UserFolder &"/music/"
		Case 9994 'С��̳
		   ReturnChannelUserUpFilesDir=Setting(3)&Setting(91)&"User/" & UserFolder &"/club/"
		Case 9993 '��־
		   ReturnChannelUserUpFilesDir=Setting(3)&Setting(91)&"User/" & UserFolder &"/blog/"
		Case 999
		   ReturnChannelUserUpFilesDir=Setting(3)&Setting(91)&"User/" & UserFolder &"/"
		Case Else
		  ReturnChannelUserUpFilesDir = Setting(3) & Setting(91)&"User/" & UserFolder &"/"
	   End Select
	End Function
	
	'�ж���û������
	function HasChinese(str) 
		HasChinese = false 
		dim i 
		for i=1 to Len(str) 
		if Asc(Mid(str,i,1)) < 0 then 
		HasChinese = true 
		exit for 
		end if 
		next 
	end function 
	
	
	'**************************************************
	'��������ReturnChannelAllowUpFilesSize
	'��  �ã�����Ƶ������������ϴ��ļ���С
	'��  ����ChannelID--Ƶ��ID
	'**************************************************
	Public Function ReturnChannelAllowUpFilesSize(ChannelID)
	   ChannelID = ChkClng(ChannelID)
	   Dim CRS:Set CRS=conn.execute("Select top 1 UpFilesSize From KS_Channel Where ChannelID=" & ChannelID)
	  If CInt(ChannelID) = 0 Or (CRS.EOF And CRS.BOF) Then
		ReturnChannelAllowUpFilesSize = Setting(6)
	  Else
		ReturnChannelAllowUpFilesSize = CRS(0)
	  End If
	CRS.Close:Set CRS = Nothing
	End Function
	'**************************************************
	'��������ReturnChannelAllowUpFilesType
	'��  �ã�����Ƶ���������ϴ����ļ�����
	'��  ����ChannelID--Ƶ��ID,TypeFlag 0-ȡȫ�� 1-ͼƬ���� 2-flash ���� 3-Windows ý������ 4-Real ���� 5-��������
	'**************************************************
	Public Function ReturnChannelAllowUpFilesType(ChannelID, TypeFlag)
	  If ChkClng(ChannelID) = 0 Then  ReturnChannelAllowUpFilesType = Setting(7):Exit Function
	  If Not IsNumeric(TypeFlag) Then TypeFlag = 0
		If TypeFlag = 0 Then   '�������������
		 ReturnChannelAllowUpFilesType = Replace(C_S(ChannelID,28) & "|" & C_S(ChannelID,29) & "|" & C_S(ChannelID,30) & "|" & C_S(ChannelID,31) & "|" & C_S(ChannelID,32),"||","|")
		Else
		 ReturnChannelAllowUpFilesType = Replace(C_S(ChannelID,27+TypeFlag),"||","|")
		End If
	End Function
	'���ظ��ʽ����,����TypeID,0���� 1�ۿ���
	Function ReturnPayment(ID,TypeID)
	  If Application(SiteSn &"Payment_" & ID&TypeID)="" Then
         Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select TypeName,Discount From KS_PaymentType Where TypeID=" & ID,conn,1,1
		 If Not RS.Eof Then
		     If TypeID=0 Then
		  	  ReturnPayment=rs(0)
			  If RS(1)<100 Then ReturnPayment=ReturnPayment & "&nbsp;&nbsp;<font color=red>�ۿ���:" & RS(1) & "%"
			 Else
			  ReturnPayment=rs(1)
			 End if
		End iF 
		Application(SiteSn &"Payment_" & ID&TypeID)=ReturnPayment
	  Else
	    ReturnPayment=Application(SiteSn &"Payment_" & ID&TypeID)
	  End If
	End Function
		'�����ջ���ʽ����,����TypeID,0���� 1����
	Function ReturnDelivery(ID,TypeID)
	  If Application(SiteSn &"Delivery_" & ID&TypeID)="" Then
         Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select TypeName,fee From KS_Delivery Where TypeID=" & ID,conn,1,1
		 If Not RS.Eof Then
		     If TypeID=0 Then
		  	  ReturnDelivery=rs(0)
			  If RS(1)=0 Then ReturnDelivery=ReturnDelivery & "&nbsp;<font color=blue>���</font>" Else ReturnDelivery=ReturnDelivery & "&nbsp;<font color=red>���� " & RS(1) & "Ԫ"
			 Else
			  ReturnDelivery=rs(1)
			 End iF
		End iF 
		Application(SiteSn &"Delivery_" & ID&TypeID)=ReturnDelivery
	  Else
	    ReturnDelivery=Application(SiteSn &"Delivery_" & ID&TypeID)
	  End If
	End Function
	'**********************************************************************
	'��������ReturnSpecial
	'��  �ã�����ר������
	'��  ����Selected-Ԥѡ����
	'����ֵ��ר������
	'**********************************************************************
	Public Function ReturnSpecial(SelectID)
	 Dim RS,ParaStr,SpecialChannelStr,SQL,K
	 Set RS=Conn.Execute("Select ClassID,ClassName From KS_SpecialClass Order By OrderID")
	 If Not RS.Eof Then SQL=RS.GetRows(-1)
     RS.Close
	 If IsArray(SQL) Then
	  For K=0 To Ubound(SQL,2)
	  ReturnSpecial = ReturnSpecial & "<optgroup label='---" & SQL(1,K) & "---'>"
	  Set RS=Conn.Execute("Select SpecialName,SpecialID From KS_Special Where ClassID=" & SQL(0,K) & " Order By SpecialID Desc")
		 If Not RS.EOF Then
		  Do While Not RS.EOF
			 If Trim(SelectID) = Trim(RS(1)) Then
				  ReturnSpecial = ReturnSpecial & "<Option value=" & RS(1) & " Selected>" & Trim(RS("SpecialName")) & SpecialChannelStr & "</Option>"
			 Else
				  ReturnSpecial = ReturnSpecial & "<Option value=" & RS(1) & ">" & Trim(RS("SpecialName")) & SpecialChannelStr & "</Option>"
			 End If
			 RS.MoveNext
		  Loop
		End If
	 Next
	  RS.Close:Set RS = Nothing
	 Else
	  Set RS = Nothing
	 End If
	End Function
	
	'**************************************************
	'������FoundInArr
	'��  �ã����һ������������Ԫ���Ƿ����ָ���ַ���
	'��  ����strArr     ----�ַ���
	'        strToFind    ----Ҫ���ҵ��ַ���
	'       strSplit    ----����ķָ���
	'����ֵ��True,False
	'**************************************************
	Public Function FoundInArr(strArr, strToFind, strSplit)
		Dim arrTemp, i
		FoundInArr = False
		If InStr(strArr, strSplit) > 0 Then
			arrTemp = Split(strArr, strSplit)
			For i = 0 To UBound(arrTemp)
			If LCase(Trim(arrTemp(i))) = LCase(Trim(strToFind)) Then
				FoundInArr = True:Exit For
			End If
			Next
		Else
			If LCase(Trim(strArr)) = LCase(Trim(strToFind)) Then FoundInArr = True
		End If
	End Function
	
	'����Ƿ������� ����ת��Ϊ������
	Public Function ChkClng(ByVal str)
	    On error resume next
		If IsNumeric(str) Then
			ChkClng = CLng(str)
		Else
			ChkClng = 0
		End If
		If Err Then ChkClng=0
	End Function
	'**************************************************
	'��������ShowPage
	'��  �ã���ʾ����һҳ ��һҳ������Ϣ
	'��  ����filename�ļ��� TotalNumber������ MaxPerPageÿҳ���� ShowTurn��ʾת�� PrintOut�������
	'**************************************************
	Function ShowPage(totalnumber, MaxPerPage, FileName, CurrPage,ShowTurn,PrintOut)
	             Dim n,j,startpage,pageStr,TotalPage,ParamStr
				 If totalnumber Mod MaxPerPage = 0 Then
						TotalPage = totalnumber \ MaxPerPage
				 Else
						TotalPage = totalnumber \ MaxPerPage + 1
				 End If
				 ParamStr=QueryParam("page") : If ParamStr<>"" Then ParamStr="&" & ParamStr	
				 n=0:startpage=1
				 pageStr=pageStr & "<div id='fenye' class='fenye'><table border=""0"" align=""right""><form action=""" & FileName & "?1=1" & ParamStr & """ name=""pageform"" method=""post""><tr><td>" & vbcrlf
				 if (CurrPage>1) then pageStr=PageStr & "<a href=""" & FileName & "?page=" & CurrPage-1 & ParamStr & """ class=""prev"">��һҳ</a>"
				 if CurrPage<>TotalPage and totalnumber>MaxPerPage then pageStr=PageStr & "<a href=""" & FileName & "?page=" & CurrPage+1 & ParamStr & """ class=""next"">��һҳ</a>"
				 pageStr=pageStr & "<a href=""" & FileName & "?page=1" & ParamStr & """ class=""prev"">�� ҳ</a>"
				 if (CurrPage>=7) then startpage=CurrPage-5
				 if TotalPage-CurrPage<5 Then startpage=TotalPage-9
				 If startpage<0 Then startpage=1
				 For J=startpage To TotalPage
				    If J= CurrPage Then
				     PageStr=PageStr & " <a href=""#"" class=""curr""><font color=red>" & J &"</font></a>"
				    Else
				     PageStr=PageStr & " <a class=""num"" href=""" & FileName & "?page=" &J& ParamStr & """>" & J &"</a>"
					End If
					n=n+1
					if n>=10 then exit for
				 Next
				 pageStr=pageStr & "<a href=""" & FileName & "?page=" & TotalPage & ParamStr & """ class=""prev"">ĩҳ</a>"
				 pageStr=PageStr & " <span>��" & TotalPage & "ҳ"
				 If ShowTurn=true Then
				 If CurrPage=TotalPage Then CurrPage=0
				 pageStr=PageStr & " ת��:<input type='text' value='" & (CurrPage + 1) &"' name='page' style='width:30px;height:18px;text-align:center;'>&nbsp;<input style='height:18px;border:1px #a7a7a7 solid;background:#fff;' type='submit' value='GO' name='sb'>"
				 End If
				 PageStr=PageStr & "</span></td></tr></form></table></div>"
				If PrintOut=true Then echo PageStr Else ShowPage=PageStr
	End Function
	'**************************************************
	'��������ShowPagePara
	'��  �ã���ʾ����һҳ ��һҳ������Ϣ
	'��  ����filename  ----���ӵ�ַ
	'       TotalNumber ----������
	'       MaxPerPage  ----ÿҳ����
	'       ShowAllPages ---�Ƿ��������б���ʾ����ҳ���Թ���ת��
	'       strUnit     ----������λ,CurrentPage--��ǰҳ,ParamterStr����
	'����ֵ���޷���ֵ
	'**************************************************
	Public Function ShowPagePara(totalnumber, MaxPerPage, FileName, ShowAllPages, strUnit, CurrentPage, ParamterStr)
		  Dim N, I, PageStr
				Const Btn_First = "<span style='font-family:webdings;font-size:14px' title='��һҳ'>9</span>" '�����һҳ��ť��ʾ��ʽ
				Const Btn_Prev = "<span style='font-family:webdings;font-size:14px' title='��һҳ'>3</span>" '����ǰһҳ��ť��ʾ��ʽ
				Const Btn_Next = "<span style='font-family:webdings;font-size:14px' title='��һҳ'>4</span>" '������һҳ��ť��ʾ��ʽ
				Const Btn_Last = "<span style='font-family:webdings;font-size:14px' title='���һҳ'>:</span>" '�������һҳ��ť��ʾ��ʽ
				  PageStr = ""
					If totalnumber Mod MaxPerPage = 0 Then
						N = totalnumber \ MaxPerPage
					Else
						N = totalnumber \ MaxPerPage + 1
					End If
				If N > 1 Then
					PageStr = PageStr & ("<div class='showpage' style='height:20px'><form action=""" & FileName & "?" & ParamterStr & """ name=""myform"" method=""post"">ҳ�Σ�<font color=red>" & CurrentPage & "</font>/" & N & "ҳ ����:" & totalnumber & strUnit & " ÿҳ:" & MaxPerPage & strUnit & " ")
					If CurrentPage < 2 Then
						PageStr = PageStr & Btn_First & " " & Btn_Prev & " "
					Else
						PageStr = PageStr & ("<a href=" & FileName & "?page=1" & "&" & ParamterStr & ">" & Btn_First & "</a> <a href=" & FileName & "?page=" & CurrentPage - 1 & "&" & ParamterStr & ">" & Btn_Prev & "</a> ")
					End If
					
					If N - CurrentPage < 1 Then
						PageStr = PageStr & " " & Btn_Next & " " & Btn_Last & " "
					Else
						PageStr = PageStr & (" <a href=" & FileName & "?page=" & (CurrentPage + 1) & "&" & ParamterStr & ">" & Btn_Next & "</a> <a href=" & FileName & "?page=" & N & "&" & ParamterStr & ">" & Btn_Last & "</a> ")
					End If
					If ShowAllPages = True Then
						PageStr = PageStr & ("ת��:<input type='text' value='" & (CurrentPage + 1) &"' name='page' style='width:30px;height:18px;text-align:center;'>&nbsp;<input style='height:18px;border:1px #a7a7a7 solid;background:#fff;' type='submit' value='GO' name='sb'>")
				  End If
				  PageStr = PageStr & "</form></div>"
			 End If
			 ShowPagePara = PageStr
	End Function
	Sub ShowPageParamter(totalnumber, MaxPerPage, FileName, ShowAllPages, strUnit, CurrentPage, ParamterStr)
		echo (ShowPagePara(totalnumber, MaxPerPage, FileName, ShowAllPages, strUnit, CurrentPage, ParamterStr))
	End Sub
	'***********************************************************************************************************
	'��������ReturnLabelFolderTree
	'��  �ã���ʾ��ǩĿ¼�б�
	'��  ����SelectID ----  Ĭ��Ŀ¼��ID��,ChannelIDƵ��ID��,FolderTypeĿ¼���� 0ϵͳ������ǩĿ¼,1���ɱ�ǩĿ¼
	'����ֵ����ǩĿ¼�б�
	'*************************************************************************************************************
	Public Function ReturnLabelFolderTree(SelectID, FolderType)
		   Dim TempStr,ID,FolderName
		   SelectID = Trim(SelectID)
		   If FolderType = "" Then FolderType = 0
		   TempStr = "<select class='textbox' style='width:200;border-style: solid; border-width: 1' name='ParentID'>"
		   
		   TempStr = TempStr & "<option value='0' Selected>��Ŀ¼</option>"
			Dim RS:Set RS=Conn.Execute("Select ID,FolderName from KS_LabelFolder Where FolderType=" & FolderType & " And ParentID='0' Order By AddDate desc")
			
			Do While Not RS.EOF
			   ID = Trim(RS(0))
			   FolderName = Trim(RS(1))
			   TempStr = TempStr & "<option  "
			   If SelectID = ID Then TempStr = TempStr & " Selected"
			   TempStr = TempStr & " value='" & ID & "'>" & FolderName & " </option>"
			   TempStr = TempStr & ReturnSubLabelFolderTree(ID, SelectID)
			RS.MoveNext
			Loop
			RS.Close:Set RS = Nothing
			TempStr = TempStr & "</select>"
			ReturnLabelFolderTree = TempStr
	End Function
	
	'************************************************************************************
	'��������ReturnSubLabelFolderTree
	'��  �ã����Ҳ����������ݡ�
	'��  ����ParentID ----���ڵ�ID,   FolderID ----ѡ����ID
	'����ֵ����ǩĿ¼�����б�
	'************************************************************************************
	Public Function ReturnSubLabelFolderTree(ParentID, FolderID)
	  Dim SubTypeList, SubRS, SpaceStr, k, Total, Num,FolderName, ID,TJ
	  
	  Set SubRS = Server.CreateObject("ADODB.RECORDSET")
	  SubRS.Open ("Select count(ID) AS total from KS_LabelFolder Where ParentID='" & ParentID & "'"), Conn, 1, 1
	  Total = SubRS("Total")
	  SubRS.Close
	  SubRS.Open ("Select ID,FolderName,TS from KS_LabelFolder Where ParentID='" & ParentID & "' Order BY AddDate Desc"), Conn, 1, 1
	  Num = 0
	  Do While Not SubRS.EOF
	   Num = Num + 1:SpaceStr = ""
		TJ = UBound(Split(SubRS(2), ","))
		For k = 1 To TJ - 1
		  If k = 1 And k <> TJ - 1 Then
		  SpaceStr = SpaceStr & "&nbsp;&nbsp;��"
		  ElseIf k = TJ - 1 Then
			If Num = Total Then
				 SpaceStr = SpaceStr & "&nbsp;&nbsp;�� "
			Else
				 SpaceStr = SpaceStr & "&nbsp;&nbsp;�� "
			End If
		  Else
		   SpaceStr = SpaceStr & "&nbsp;&nbsp;��"
		  End If
		Next
	  ID = Trim(SubRS(0))
	  FolderName = Trim(SubRS(1))
	  If FolderID = ID Then
	   SubTypeList = SubTypeList & "<option selected value='" & ID & "'>" & SpaceStr & FolderName & "</option>"
	  Else
	   SubTypeList = SubTypeList & "<option value='" & ID & "'>" & SpaceStr & FolderName & "</option>"
	  End If
	   SubTypeList = SubTypeList & ReturnSubLabelFolderTree(ID, FolderID)
	  SubRS.MoveNext
	 Loop
	  SubRS.Close:Set SubRS = Nothing:ReturnSubLabelFolderTree = SubTypeList
	End Function
	
	'***********************************************************************************************************
	'��������ReturnLabelInfo
	'��  ����LabelName ----  Ĭ�ϱ�ǩ����,FolderID---��ǩĿ¼ID��,Descript---��ǩ����
	'����ֵ����ǩ������Ϣ
	'*************************************************************************************************************
	Public Function ReturnLabelInfo(LabelName, FolderID, Descript)
	  ReturnLabelInfo = ReturnLabelInfo & ("        <table width=""98%"" border='0' align='center' cellpadding='2' cellspacing='1' class='border' style='margin-top:6px'>")
	  ReturnLabelInfo = ReturnLabelInfo & ("          <tr  height=""26"" class=title><td colspan=2 align=center><strong>")
	  If g("labelid")="" Then
	  ReturnLabelInfo = ReturnLabelInfo & ("�� �� �� �� ǩ")
	  Else
	  ReturnLabelInfo = ReturnLabelInfo & (" �� �� �� ǩ �� ��")
	  End If
	  ReturnLabelInfo = ReturnLabelInfo & ("</strong></td>")
	  ReturnLabelInfo = ReturnLabelInfo & ("          </tr>")
	  ReturnLabelInfo = ReturnLabelInfo & ("          <tr class=tdbg>")
	  ReturnLabelInfo = ReturnLabelInfo & ("      <td  colspan=2 height=""30"">��ǩ����")
	  ReturnLabelInfo = ReturnLabelInfo & ("        <input name=""LabelName"" size='35' class=""textbox"" type=""text"" id=""LabelName"" value=""" & LabelName & """>")
	  ReturnLabelInfo = ReturnLabelInfo & ("        <font color=""#FF0000""> * ���ø�ʽ""{LB_��ǩ����}""</font></td>")
	  ReturnLabelInfo = ReturnLabelInfo & ("    </tr>")
	  ReturnLabelInfo = ReturnLabelInfo & ("    <tr class=tdbg>")
	  ReturnLabelInfo = ReturnLabelInfo & ("      <td  colspan=2 height=""30"">��ǩĿ¼ " & ReturnLabelFolderTree(FolderID, 0) & "<font color=""#FF0000""> ��ѡ���ǩ����Ŀ¼���Ա��պ�����ǩ</font></td>")
	  ReturnLabelInfo = ReturnLabelInfo & ("    </tr>")
	  ReturnLabelInfo = ReturnLabelInfo & ("    <tr class=tdbg style='display:none'>")
	  ReturnLabelInfo = ReturnLabelInfo & ("      <td  colspan=2 height=""30"">��ǩ����")
	  ReturnLabelInfo = ReturnLabelInfo & ("        <input name=""Descript"" class=""textbox"" type=""text"" id=""Descript"" value=""" & Descript & """ size=""40"">")
	  ReturnLabelInfo = ReturnLabelInfo & ("        <font color=""#FF0000""> ���ڴ������ǩ��˵��,�����Ժ����</font></td>")
	  ReturnLabelInfo = ReturnLabelInfo & ("    </tr>")
	 ' ReturnLabelInfo = ReturnLabelInfo & ("    </table>")
	End Function
	
	'ģ��ѡ��
	Sub LoadChannelOption(ChannelID)
		If not IsObject(Application(SiteSN&"_ChannelConfig")) Then LoadChannelConfig
		Dim ModelXML,Node
		Set ModelXML=Application(SiteSN&"_ChannelConfig")
		For Each Node In ModelXML.documentElement.SelectNodes("channel")
		 if Node.SelectSingleNode("@ks21").text="1" and Node.SelectSingleNode("@ks0").text<>"6" and Node.SelectSingleNode("@ks0").text<>"9" and Node.SelectSingleNode("@ks0").text<>"10" Then
		  If Trim(ChannelID)=Trim(Node.SelectSingleNode("@ks0").text) Then
		  echo "<option value='" &Node.SelectSingleNode("@ks0").text &"' selected>" & Node.SelectSingleNode("@ks1").text & "</option>"
		  Else
		  echo "<option value='" &Node.SelectSingleNode("@ks0").text &"'>" & Node.SelectSingleNode("@ks1").text & "</option>"
		  End If
		 End If
		next
	End Sub
		
	'****************************************************************************************************************************
	'��������ReturnJSInfo
	'��  ����JSID--JSID��,JSName ----    Ĭ��JS����,JSFileName----JS�ļ���,FolderID---��ǩĿ¼ID��,FolderType---Ŀ¼����,Descript---��ǩ����
	'����ֵ����ǩ������Ϣ
'*******************************************************************************************************************************
	Public Function ReturnJSInfo(JSID, JSName, JSFileName, FolderID, FolderType, Descript)
		 ReturnJSInfo = "<table width=""96%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		 ReturnJSInfo = ReturnJSInfo & ("    <tr>")
		 ReturnJSInfo = ReturnJSInfo & ("       <td>")
		 ReturnJSInfo = ReturnJSInfo & ("      <FIELDSET align=center><LEGEND align=left>JS������Ϣ</LEGEND>")
		 ReturnJSInfo = ReturnJSInfo & ("        <table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">")
		 ReturnJSInfo = ReturnJSInfo & ("            <tr>")
		 ReturnJSInfo = ReturnJSInfo & ("             <td height=""22"">JS �� ��")
		 ReturnJSInfo = ReturnJSInfo & ("                &nbsp;<input name=""JSName"" type=""text"" class=""textbox"" id=""JSName"" value=""" & JSName & """>")
		 ReturnJSInfo = ReturnJSInfo & ("                <font color=""#FF0000""> *</font><font color=""#FF0000""> ����JS���ƣ�&quot;�Ƽ������б�&quot;������ģ���е��ã�&quot;{JS_�Ƽ������б�}&quot;��ע��Ӣ�Ĵ�Сд��ȫ��ǣ���</font></td>")
		 ReturnJSInfo = ReturnJSInfo & ("            </tr>")
		 ReturnJSInfo = ReturnJSInfo & ("            <tr>")
		 ReturnJSInfo = ReturnJSInfo & ("              <td height=""22"">JS�ļ���")
		 
		   If JSID <> "" Then
			  ReturnJSInfo = ReturnJSInfo & ("                <input class=""textbox"" disabled=true name=""JSFileName"" type=""text"" id=""JSFileName"" title=""JS�ļ��������ܴ�\/��*���� < > | ���������"" value=""" & JSFileName & """>")
		   Else
			  ReturnJSInfo = ReturnJSInfo & ("                <input class=""textbox"" name=""JSFileName"" type=""text"" id=""JSFileName"" title=""JS�ļ��������ܴ�\/��*���� < > | ���������"" value=""" & JSFileName & """>")
		   End If
		 ReturnJSInfo = ReturnJSInfo & ("            <font color=""#FF0000""> * ���� &quot;News.js&quot; һ��Ҫ����չ�� &quot;.js&quot;����</font></td>")
		 ReturnJSInfo = ReturnJSInfo & ("        </tr>")
		 ReturnJSInfo = ReturnJSInfo & ("        <tr>")
		 ReturnJSInfo = ReturnJSInfo & ("         <td height=""22"">���Ŀ¼ " & ReturnLabelFolderTree(FolderID, FolderType) & " </td>")
		 ReturnJSInfo = ReturnJSInfo & ("       </tr>")
		 ReturnJSInfo = ReturnJSInfo & ("            <tr>")
		 ReturnJSInfo = ReturnJSInfo & ("              <td height=""22"">JS �� ��")
		 ReturnJSInfo = ReturnJSInfo & ("                <textarea class=""textbox"" name=""Descript"" cols=""60"" rows=""4"" id=""Descript"">" & Descript & "</textarea>")
		 ReturnJSInfo = ReturnJSInfo & ("           <font color=""#FF0000""> ���ڴ�����JS��˵��,�����Ժ����</font></td>")
		 ReturnJSInfo = ReturnJSInfo & ("            </tr>")
		 ReturnJSInfo = ReturnJSInfo & ("          </table>")
		 ReturnJSInfo = ReturnJSInfo & ("        </FIELDSET></td>")
		 ReturnJSInfo = ReturnJSInfo & ("      </tr>")
		 ReturnJSInfo = ReturnJSInfo & ("   </table>")
		 
		 '�ɼ���������
		 ReturnJSInfo = ReturnJSInfo & ("<input type=""hidden"" name=""KeyWord"" value=""" & Request.QueryString("KeyWord") & """>")
		 ReturnJSInfo = ReturnJSInfo & ("<input type=""hidden"" name=""SearchType"" value=""" & Request.QueryString("SearchType") & """>")
		 ReturnJSInfo = ReturnJSInfo & ("<input type=""hidden"" name=""StartDate"" value=""" & Request.QueryString("StartDate") & """>")
		 ReturnJSInfo = ReturnJSInfo & ("<input type=""hidden"" name=""EndDate"" value=""" & Request.QueryString("EndDate") & """>")
	End Function

	'**************************************************
	'��������ReturnDateFormat
	'��  �ã�����ϵͳ֧�ֵ����ڸ�ʽ
	'��  ����SelectDate Ԥ��ѡ�е����ڸ�ʽ
	'**************************************************
	Public Function ReturnDateFormat(SelectDate)
			 Dim TempFormatDateStr, Str
			 If CStr(SelectDate) = "0" Then
				TempFormatDateStr = ("<option value=""0"" Selected>-����ʾ����-</option> ")
			  Else
				TempFormatDateStr = ("<option value=""0"">-����ʾ����-</option> ")
			  End If
			  If CStr(SelectDate) = "1" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""1""" & Str & " >2005-10-1</option>")
			  If CStr(SelectDate) = "2" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""2""" & Str & ">2005.10.1</option>")
			  If CStr(SelectDate) = "3" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""3""" & Str & ">2005/10/1</option>")
			  If CStr(SelectDate) = "4" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""4""" & Str & ">10/1/2005</option>")
			  If CStr(SelectDate) = "5" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""5""" & Str & ">2005��10��</option>")
			  If CStr(SelectDate) = "6" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""6""" & Str & ">2005��10��1��</option>")
			  If CStr(SelectDate) = "7" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""7""" & Str & ">10.1.2005</option>")
			  If CStr(SelectDate) = "8" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""8""" & Str & ">10-1-2005</option>")
			  If CStr(SelectDate) = "9" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""9""" & Str & ">10/1</option>")
			  If CStr(SelectDate) = "10" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""10""" & Str & ">10.1</option>")
			  If CStr(SelectDate) = "11" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""11""" & Str & ">10��1��</option>")
			  If CStr(SelectDate) = "12" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""12""" & Str & ">1��12ʱ</option>")
			  If CStr(SelectDate) = "13" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""13""" & Str & ">1��12��</option>")
			  If CStr(SelectDate) = "14" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""14""" & Str & ">12ʱ12��</option>")
			  If CStr(SelectDate) = "15" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""15""" & Str & ">12:12</option>")
			  If CStr(SelectDate) = "16" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""16""" & Str & ">10-1</option>")
			   If CStr(SelectDate) = "17" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""17""" & Str & ">10/1 12:00</option>")
			  
			  TempFormatDateStr = TempFormatDateStr & ("<optgroup  label=""-----�����Ÿ�ʽ-----""></optgroup>")

			  If CStr(SelectDate) = "21" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""21""" & Str & " >(2005-10-1)</option>") 
			  If CStr(SelectDate) = "22" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""22""" & Str & ">(2005.10.1)</option>")
			  If CStr(SelectDate) = "23" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""23""" & Str & ">(2005/10/1)</option>")
			  If CStr(SelectDate) = "24" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""24""" & Str & ">(10/1/2005)</option>")
			  If CStr(SelectDate) = "25" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""225""" & Str & ">(2005��10��)</option>")
			  If CStr(SelectDate) = "26" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""26""" & Str & ">(2005��10��1��)</option>")
			  If CStr(SelectDate) = "27" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""27""" & Str & ">(10.1.2005)</option>")
			  If CStr(SelectDate) = "28" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""28""" & Str & ">(10-1-2005)</option>")
			  If CStr(SelectDate) = "29" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""29""" & Str & ">(10/1)</option>")
			  If CStr(SelectDate) = "30" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""30""" & Str & ">(10.1)</option>")
			  If CStr(SelectDate) = "31" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""31""" & Str & ">(10��1��)</option>")
			  If CStr(SelectDate) = "32" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""32""" & Str & ">(1��12ʱ)</option>")
			  If CStr(SelectDate) = "33" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""33""" & Str & ">(1��12��)</option>")
			  If CStr(SelectDate) = "34" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""34""" & Str & ">(12ʱ12��)</option>")
			  If CStr(SelectDate) = "35" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""35""" & Str & ">(12:12)</option>")
			  If CStr(SelectDate) = "36" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""36""" & Str & ">(10-1)</option>")
			  If CStr(SelectDate) = "37" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""37""" & Str & ">(10/1 12:00)</option>")


			  TempFormatDateStr = TempFormatDateStr & ("<optgroup  label=""-----�������Ÿ�ʽ-----""></optgroup>")
			  If CStr(SelectDate) = "41" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""41""" & Str & ">[2005-10-1]</option>")
			  If CStr(SelectDate) = "42" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""42""" & Str & ">[2005.10.1]</option>")
			  If CStr(SelectDate) = "43" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""43""" & Str & ">[2005/10/1]</option>")
			  If CStr(SelectDate) = "44" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""44""" & Str & ">[10/1/2005]</option>")
			  If CStr(SelectDate) = "45" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""45""" & Str & ">[2005��10��]</option>")
			  If CStr(SelectDate) = "46" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""46""" & Str & ">[2005��10��1��]</option>")
			  If CStr(SelectDate) = "47" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""47""" & Str & ">[10.1.2005]</option>")
			  If CStr(SelectDate) = "48" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""48""" & Str & ">[10-1-2005]</option>")
			  If CStr(SelectDate) = "49" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""49""" & Str & ">[10/1]</option>")
			  If CStr(SelectDate) = "50" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""50""" & Str & ">[10.1]</option>")
			  If CStr(SelectDate) = "51" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""51""" & Str & ">[10��1��]</option>")
			  If CStr(SelectDate) = "52" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""52""" & Str & ">[1��12ʱ]</option>")
			  If CStr(SelectDate) = "53" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""53""" & Str & ">[1��12��]</option>")
			  If CStr(SelectDate) = "54" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""54""" & Str & ">[12ʱ12��]</option>")
			  If CStr(SelectDate) = "55" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""55""" & Str & ">[12:12]</option>")
			  If CStr(SelectDate) = "56" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""56""" & Str & ">[10-1]</option>")
			  If CStr(SelectDate) = "57" Then Str = " Selected" Else Str = ""
			   TempFormatDateStr = TempFormatDateStr & ("<option value=""57""" & Str & ">[10/1 12:00]</option>")
			ReturnDateFormat = TempFormatDateStr
	End Function
	'**************************************************
	'��������ReturnOpenTypeStr
	'��  �ã�����ϵͳ֧�ֵĴ򿪴��ڷ�ʽ(���������������)
	'��  ����SelectValue Ԥ��ѡ�е�����Ŀ��
	'**************************************************
	Public Function ReturnOpenTypeStr(SelectValue)
	  ReturnOpenTypeStr = "����Ŀ�� <select onchange=""document.getElementById('OpenType').value=this.value;"" name='sOpenType'><option value=''>-û������-</option>"
	  ReturnOpenTypeStr = ReturnOpenTypeStr & "<option value=""_blank""> �´���(_blank) </option>"
	  ReturnOpenTypeStr = ReturnOpenTypeStr & "<option value=""_parent""> ������(_parent) </option>"
	  ReturnOpenTypeStr = ReturnOpenTypeStr & "<option value=""_self""> ������(_self) </option>"
	  ReturnOpenTypeStr = ReturnOpenTypeStr & "<option value=""_top""> ��ҳ(_top) </option>"
	  ReturnOpenTypeStr = ReturnOpenTypeStr & "</select>=>"
	  ReturnOpenTypeStr = ReturnOpenTypeStr & "<input type='text' name='OpenType' id='OpenType' size='10' value='" & SelectValue &"'>"
	  Exit Function
	End Function
	'��ҳ��ʽ
	Public Function ReturnPageStyle(PageStyle)
		ReturnPageStyle = "         ��ҳ��ʽ"
		ReturnPageStyle = ReturnPageStyle & "         <select name=""PageStyle"" id=""PageStyle"" style=""width:70%;"" class=""textbox"">"
		ReturnPageStyle = ReturnPageStyle & "          <option value=1"
		If PageStyle="1" Then ReturnPageStyle = ReturnPageStyle & " Selected"
		ReturnPageStyle = ReturnPageStyle & ">����ҳ ��һҳ ��һҳ βҳ</option>"
		ReturnPageStyle = ReturnPageStyle & "          <option value=2"
		If PageStyle="2" Then ReturnPageStyle = ReturnPageStyle & " Selected"
		ReturnPageStyle = ReturnPageStyle & ">�ڹ�Nҳ/Nƪ [1] [2] [3]</option>"
		ReturnPageStyle = ReturnPageStyle & "          <option value=3"
		If PageStyle="3" Then ReturnPageStyle = ReturnPageStyle & " Selected"
		ReturnPageStyle = ReturnPageStyle & ">��<< <  > >></option>"
		ReturnPageStyle = ReturnPageStyle & "          <option value=4"
		If PageStyle="4" Then ReturnPageStyle = ReturnPageStyle & " Selected"
		ReturnPageStyle = ReturnPageStyle & " style='color:blue'>�����ֵ�����ʽ(����)</option>"
		ReturnPageStyle = ReturnPageStyle & "         </select>"
	End Function
	'ר����ʾ��ʽ
	Public Function ReturnSpecialStyle(Sel)
		ReturnSpecialStyle= "��ʾ��ʽ&nbsp;<select name=""ShowStyle"" id=""ShowStyle"" style=""width:70%"" class=""textbox"">"
		ReturnSpecialStyle=ReturnSpecialStyle &"<option value=""1"""
		If Sel="1" Then ReturnSpecialStyle=ReturnSpecialStyle &" selected"
		ReturnSpecialStyle=ReturnSpecialStyle &">�ٱ���ʽ</option>"
		ReturnSpecialStyle=ReturnSpecialStyle &"<option value=""2"""
		If Sel="2" Then ReturnSpecialStyle=ReturnSpecialStyle &" selected"
		ReturnSpecialStyle=ReturnSpecialStyle &">�ڽ���ʾͼƬ</option>"
		ReturnSpecialStyle=ReturnSpecialStyle &"<option value=""3"""
		If Sel="3" Then ReturnSpecialStyle=ReturnSpecialStyle &" Selected"
		ReturnSpecialStyle=ReturnSpecialStyle &">��ͼƬ+����:����</option>"
		ReturnSpecialStyle=ReturnSpecialStyle &"<option value=""4"""
		If Sel="4" Then ReturnSpecialStyle=ReturnSpecialStyle &" Selected"
		ReturnSpecialStyle=ReturnSpecialStyle &">��ͼƬ+����:����</option>"
		ReturnSpecialStyle=ReturnSpecialStyle &"<option value=""5"""
		If Sel="5" Then ReturnSpecialStyle=ReturnSpecialStyle &" selected"
		ReturnSpecialStyle=ReturnSpecialStyle &">��ͼƬ+(����+����:����):����</option>"
		ReturnSpecialStyle=ReturnSpecialStyle &"</select>"
	End Function
	
	'**************************************************
	'��������SaveBeyondFile
	'��  �ã�����Զ���ļ�������
	'��  ����LocalFile �����ļ�,BFUԶ���ļ�
	'����ֵ����
	'**************************************************
	Public Function ReplaceBeyondUrl(ReplaceContent, SaveFilePath)
		Dim re, BeyondFile, BFU, SaveFileName,SaveFileList
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\){1}((\w)+[.]){1,}(net|com|cn|org|cc|tv|[0-9]{1,3})(\S*\/)((\S)+[.]{1}(gif|jpg|png|bmp)))"
		Set BeyondFile = re.Execute(ReplaceContent)
		Set re = Nothing
		For Each BFU In BeyondFile
		  If Instr(SaveFileList,BFU)=0 Then
			SaveFileName = Year(Now()) & Month(Now()) & Day(Now()) & MakeRandom(10) & Mid(BFU, InStrRev(BFU, "."))
			If Instr(BFU,Setting(2))<=0 Then
			Call SaveBeyondFile(SaveFilePath&SaveFileName,BFU)
			ReplaceContent = Replace(ReplaceContent, BFU, Setting(2) & SaveFilePath & SaveFileName)
			End If
		  End If
		   SaveFileList=SaveFileList & "," & BFU
		Next
		ReplaceBeyondUrl = ReplaceContent
	End Function

	'==================================================
	'��������SaveBeyondFile
	'��  �ã�����Զ�̵��ļ�������
	'��  ����LocalFileName ------ �����ļ���
	'��  ����RemoteFileUrl ------ Զ���ļ�URL
	'==================================================
	Function SaveBeyondFile(LocalFileName,RemoteFileUrl)
	    on error resume next
		Dim SaveRemoteFile:SaveRemoteFile=True
		dim Ads,Retrieval,GetRemoteData
		Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
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
		Set Ads = Server.CreateObject("Adodb.Stream")
		With Ads
			.Type = 1
			.Open
			.Write GetRemoteData
			.SaveToFile server.MapPath(LocalFileName),2
			.Cancel()
			.Close()
		End With
		Set Ads=nothing
		SaveBeyondFile=SaveRemoteFile
		'��ˮӡ
		Dim T:Set T=New Thumb
		call T.AddWaterMark(LocalFileName)
		Set T=Nothing
	end Function
	'****************************************************
	'����˵��
	  'Subject     : �ʼ�����
	  'MailAddress : �����������ĵ�ַ,��smtp.163.com
	  'LoginName     ----��¼�û���(����Ҫ����д"")
	  'LoginPass     ----�û�����(����Ҫ����д"")
	  'Email       : �ռ����ʼ���ַ
	  'Sender      : ����������
	  'Content     : �ʼ�����
	  'Fromer      : �����˵��ʼ���ַ
	'****************************************************
	  Public Function SendMail(MailAddress, LoginName, LoginPass, Subject, Email, Sender, Content, Fromer)
	   On Error Resume Next
		Dim JMail
		  Set jmail = Server.CreateObject("JMAIL.Message") '���������ʼ��Ķ���
			jmail.silent = true '����������󣬷���FALSE��TRUE��ֵj
			jmail.Charset = "GB2312" '�ʼ������ֱ���Ϊ����
			jmail.ContentType = "text/html" '�ʼ��ĸ�ʽΪHTML��ʽ
			jmail.AddRecipient Email '�ʼ��ռ��˵ĵ�ַ
			jmail.From = Fromer '�����˵�E-MAIL��ַ
			jmail.FromName = Sender
			  If LoginName <> "" And LoginPass <> "" Then
				JMail.MailServerUserName = LoginName '�����ʼ���������¼��
				JMail.MailServerPassword = LoginPass '��¼����
			  End If

			jmail.Subject = Subject '�ʼ��ı��� 
			JMail.Body = Content
			JMail.Priority = 1'�ʼ��Ľ�������1 Ϊ��죬5 Ϊ������ 3 ΪĬ��ֵ
			jmail.Send(MailAddress) 'ִ���ʼ����ͣ�ͨ���ʼ���������ַ��
			jmail.Close() '�رն���
		Set JMail = Nothing
		If Err Then
			SendMail = Err.Description
			Err.Clear
		Else
			SendMail = "OK"
		End If
	  End Function
	
	'**************************************************
	'�������� ReplaceUserFile
	'��  �ã�����Ա�ϴ����ļ��Ƶ�ָ�����ϴ�Ŀ¼��
	'**************************************************
	Public Function ReplaceUserFile(ReplaceContent,ChannelID)
		Dim re, BeyondFile, BFU, SaveFileName
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "(" &Setting(3)&Setting(91) & "user(\S*\/)((\S)+[.]{1}(gif|jpg|png|bmp|rar|doc|xsl|zip|exe)))"
		Set BeyondFile = re.Execute(ReplaceContent)
		Set re = Nothing
		Dim Path,DateDir
		Path = GetUpFilesDir()
		DateDir = Year(Now()) & Right("0" & Month(Now()), 2) & "/"
		Path = Path & "/" & DateDir
		For Each BFU In BeyondFile
		    Dim NewPath:NewPath=Path & Split(BFU,"/")(Ubound(Split(bfu,"/")))
			Call CopyFile(BFU,NewPath)
			ReplaceContent = Replace(ReplaceContent, BFU, NewPath)
		Next
		ReplaceUserFile = ReplaceContent
	End Function
	
	'ģ������ļ�����
	Public Function CopyFile(OldPath,NewPath)
		CopyFile=false
		Call CreateListFolder(Replace(NewPath,Split(NewPath,"/")(Ubound(Split(NewPath,"/"))),""))
		on error resume next
		dim fso:set fso = Server.CreateObject(Setting(99))
	    fso.CopyFile Server.MapPath(OldPath), server.mappath(NewPath), True
		DeleteFile(OldPath)
		if err then
			CopyFile=false
		else
			CopyFile=true
		end if
		IF err Then	 CopyFile=false
	End Function
	
	'**************************************************
	'��������CreateListFolder
	'��  �ã����޷ּ�����Ŀ¼ ���� 1\2\3\ ������վ��Ŀ¼�´����ּ�Ŀ¼
	'��  ����FolderҪ������Ŀ¼
	'����ֵ���ɹ�����true ���򷵻�Flase
	'**************************************************
	Public Function CreateListFolder(ByVal Folder)
		Dim FSO, WaitCreateFolder, SplitFolder, CF, k
		 On Error Resume Next
		If Folder = "" Then
		 CreateListFolder = False:Exit Function
		End If
	   Folder = Replace(Folder, "\", "/")
	   If Right(Folder, 1) <> "/" Then Folder = Folder & "/"
	   If Left(Folder, 1) <> "/" Then Folder = "/" & Folder

		 Set FSO = CreateObject(Setting(99))
		 If Not FSO.FolderExists(Server.MapPath(Folder)) Then
		   SplitFolder = Split(Folder, "/")
		 For k = 0 To UBound(SplitFolder) - 1
		  If k = 0 Then
		   CF = SplitFolder(k) & "/"
		  Else
		   CF = CF & SplitFolder(k) & "/"
		  End If
		  If (Not FSO.FolderExists(Server.MapPath(CF))) Then
			 FSO.CreateFolder (Server.MapPath(CF))
			 CreateListFolder = True
		  End If
		 Next
	   End If
	   Set FSO = Nothing
	   If Err.Number <> 0 Then
	   Err.Clear
	   CreateListFolder = False
	   Else
	   CreateListFolder = True
	   End If
	 End Function
	
	 '**************************************************
	'��������DeleteFolder
	'��  �ã�ɾ��ָ��Ŀ¼
	'��  ����FolderStrҪɾ����Ŀ¼
	'����ֵ���ɹ�����true ���򷵻�Flase
	'**************************************************
	Public Function DeleteFolder(FolderStr)
	   Dim FSO
	   On Error Resume Next
	   FolderStr = Replace(FolderStr, "\", "/")
	   Set FSO = CreateObject(Setting(99))
		If FSO.FolderExists(Server.MapPath(FolderStr)) Then
			  FSO.DeleteFolder (Server.MapPath(FolderStr))
		Else
		DeleteFolder = True
		End If
	   Set FSO = Nothing
	   If Err.Number <> 0 Then
	   Err.Clear:DeleteFolder = False
	   Else
	   DeleteFolder = True
	   End If
	End Function
	 '**************************************************
	'��������DeleteFile
	'��  �ã�ɾ��ָ���ļ�
	'��  ����FileStrҪɾ�����ļ�
	'����ֵ���ɹ�����true ���򷵻�Flase
	'**************************************************
	Public Function DeleteFile(FileStr)
	   Dim FSO
	   On Error Resume Next
	   Set FSO = CreateObject(Setting(99))
		If FSO.FileExists(Server.MapPath(FileStr)) Then
			FSO.DeleteFile Server.MapPath(FileStr), True
		Else
		DeleteFile = True
		End If
	   Set FSO = Nothing
	   If Err.Number <> 0 Then
	   Err.Clear:DeleteFile = False
	   Else
	   DeleteFile = True
	   End If
	End Function
	'**********************************************************************
	'��������CheckFileShowOrNot
	'������AllowShowExtNameStr������ļ���չ����ExtNameʵ���ļ���չ��
	'**********************************************************************
	Public Function CheckFileShowOrNot(AllowShowExtNameStr, ExtName)
		If ExtName = "" Then
			CheckFileShowOrNot = False
		Else
			If InStr(1, AllowShowExtNameStr, ExtName) = 0 Then
				CheckFileShowOrNot = False
			Else
				CheckFileShowOrNot = True
			End If
		End If
	End Function
	'**********************************************************************
	'��������GetFieSize
	'���ã�ȡ��ָ���ļ��Ĵ�С
	'������FilePath--�ļ�λ��
	'**********************************************************************
	Public Function GetFieSize(FilePath)
			GetFieSize = 0
			Dim FSO, F
			On Error Resume Next
			Set FSO = Server.CreateObject(Setting(99))
			Set F = FSO.GetFile(FilePath)
			GetFieSize = F.size
			Set F = Nothing:Set FSO = Nothing
	End Function
    'ȡ��Ŀ¼��С
	Public Function GetFolderSize(FolderPath)
		dim fso:Set FSO = Server.CreateObject(Setting(99))
		if fso.FolderExists(Server.MapPath(FolderPath)) then
		dim userfilespace:set UserFileSpace=FSO.GetFolder(Server.MapPath(FolderPath))
        GetFolderSize=UserFileSpace.size
		else
		 GetFolderSize=0:exit function
		end if
		set userfilespace=nothing:set fso=nothing
	End Function
	'*************************************************************************************
	'�ļ����ݹ���
	'��������backupdata
	'������CurrPathԭ�ļ����������ַ��BackPathĿ�걸���ļ����������ַ
	'*************************************************************************************
	
	Public Function BackUpData(CurrPath, BackPath)
	  On Error Resume Next
	  Dim FSO:Set FSO = Server.CreateObject(Setting(99))
	 FSO.copyfile CurrPath, BackPath
	 If Err Then
	   BackUpData = False
	 Else
	   BackUpData = True
	 End If
	  FSO.Close:Set FSO = Nothing
	End Function
	'------------------���ĳһĿ¼�Ƿ����-------------------
	Public Function CheckDir(FolderPath)
	Dim fso1
	FolderPath = Server.MapPath(".") & "\" & FolderPath
	Set fso1 = CreateObject(Setting(99))
	If fso1.FolderExists(FolderPath) Then
	CheckDir = True
	Else
	CheckDir = False
	End If
	Set fso1 = Nothing
	End Function
	'------------------���ĳһ�ļ��Ƿ����-------------------
	Public Function CheckFile(FileName)
		 On Error Resume Next
		 Dim FsoObj
		 Set FsoObj = Server.CreateObject(Setting(99))
		  If Not FsoObj.FileExists(Server.MapPath(FileName)) Then
			  CheckFile = False
			  Exit Function
		  End If
		 CheckFile = True:Set FsoObj = Nothing
	End Function
	
	
	'**************************************************
	'��������WriteTOFile
	'��  �ã�д���ݵ�ָ����html�ļ�
	'��  ����Filename  ----Ŀ���ļ��� �� mb\index.htm
	'        Content   ------Ҫд��Ŀ���ļ�������
	'����ֵ���ɹ�����true ,ʧ�ܷ���false
	'**************************************************
	Public Function WriteTOFile(FileName, Content)
	    On Error Resume Next
		dim stm:set stm=server.CreateObject("adodb.stream")
		stm.Type=2 '���ı�ģʽ��ȡ
		stm.mode=3
		stm.charset="gb2312"
		stm.open
		stm.WriteText content
		stm.SaveToFile server.MapPath(FileName),2 
		stm.flush
		stm.Close
		set stm=nothing
	  
	   If Err.Number <> 0 Then
		 WriteTOFile = False
	   Else
		 WriteTOFile = True
	   End If
	End Function
	'**************************************************
	'��������ReadFromFile
	'��  �ã�д���ݵ�ָ����html�ļ�
	'��  ����Filename  ----Ŀ���ļ��� �� mb\index.htm
	'����ֵ���ɹ������ļ����� ,ʧ�ܷ���""
	'**************************************************
	Public Function ReadFromFile(FileName)
		 On Error Resume Next
		 Dim FsoObj, FileStreamObj, FileObj
		 Set FsoObj = Server.CreateObject(Setting(99))
		 If CheckFile(FileName) = False Then
			  Call Alert("������ʾ:\n\n[" & Server.MapPath(FileName) & "]�ļ�������", ""):Exit Function
		  End If
		  Set FileObj = FsoObj.GetFile(Server.MapPath(FileName))
		  Set FileStreamObj = FileObj.OpenAsTextStream(1)
		  If Not FileStreamObj.AtEndOfStream Then
				ReadFromFile = FileStreamObj.ReadAll
		 Else
				 ReadFromFile = ""
		 End If
	End Function
	'**************************************************
	'��������MakeRandom
	'��  �ã�����ָ��λ���������
	'��  ���� maxLen  ----����λ��
	'����ֵ���ɹ�:���������
	'**************************************************
	Public Function MakeRandom(ByVal maxLen)
	  Dim strNewPass,whatsNext, upper, lower, intCounter
	  Randomize
	 For intCounter = 1 To maxLen
	   upper = 57:lower = 48:strNewPass = strNewPass & Chr(Int((upper - lower + 1) * Rnd + lower))
	 Next
	   MakeRandom = strNewPass
	End Function
	'�����������
	Function GetRndPassword(PasswordLen)
		Dim Ran, i, strPassword
		strPassword = ""
		For i = 1 To PasswordLen
			Randomize
			Ran = CInt(Rnd * 2)
			Randomize
			If Ran = 0 Then
				Ran = CInt(Rnd * 25) + 97
				strPassword = strPassword & UCase(Chr(Ran))
			ElseIf Ran = 1 Then
				Ran = CInt(Rnd * 9)
				strPassword = strPassword & Ran
			ElseIf Ran = 2 Then
				Ran = CInt(Rnd * 25) + 97
				strPassword = strPassword & Chr(Ran)
			End If
		Next
		GetRndPassword = strPassword
	End Function
	'**************************************************
	'��������MakeRandomChar
	'��  �ã�����ָ��λ����������ַ��� �� "sJKD_!@KK"
	'��  ���� Length  ----����λ��
	'����ֵ���ɹ���������ַ���
	'**************************************************
	Public Function MakeRandomChar(Length)
	  Dim I, tempS, v
	  Dim c(65)
	   tempS = ""
	   c(1) = "a": c(2) = "b": c(3) = "c": c(4) = "d": c(5) = "e": c(6) = "f": c(7) = "g"
	   c(8) = "h": c(9) = "i": c(10) = "j": c(11) = "k": c(12) = "l": c(13) = "m": c(14) = "n"
	  c(15) = "o": c(16) = "p": c(17) = "q": c(18) = "r": c(19) = "s": c(20) = "t": c(21) = "u"
	  c(22) = "v": c(23) = "w": c(24) = "x": c(25) = "y": c(26) = "z": c(27) = "1": c(28) = "2"
	   c(29) = "3": c(30) = "4": c(31) = "5": c(32) = "6": c(33) = "7": c(34) = "8": c(35) = "9"
	  c(36) = "-": c(37) = "_": c(38) = "@": c(39) = "!": c(40) = "A": c(41) = "B": c(42) = "C"
	  c(43) = "D": c(44) = "E": c(45) = "F": c(46) = "G": c(47) = "H": c(48) = "I": c(49) = "J": c(50) = "K"
	  c(51) = "L": c(52) = "M": c(53) = "N": c(54) = "O": c(55) = "P": c(56) = "Q": c(57) = "R": c(58) = "S"
	  c(59) = "J": c(60) = "U": c(61) = "V": c(62) = "W": c(63) = "X": c(64) = "Y": c(65) = "Z"
	
	  If IsNumeric(Length) = False Then
		 MakeRandomChar = "":Exit Function
	  End If
	  For I = 1 To Length
		 Randomize
		 v = Int((65 * Rnd) + 1):tempS = tempS & c(v)
		 Next
		MakeRandomChar = tempS
	End Function
	'**************************************************
	'��������GetFileName
	'��  �ã������ļ�����
	'��  ����FsoType  ----��������,addDate   -----���ʱ��,GetFileNameType--��չ��
	'**************************************************
	Public Function GetFileName(FsoType, AddDate, GetFileNameType)
		Dim N
		Randomize
		N = Rnd * 10 + 5
		Dim Y,M,D
		Y=Year(AddDate):M=Right("0"&Month(AddDate),2):D=Right("0"&Day(AddDate),2)
	 Select Case FsoType
	  Case 1:GetFileName = Y & "/" & M & "-" & D & "/" & MakeRandom(N) & GetFileNameType  '��/��-��/�����+��չ��
	  Case 2:GetFileName = Y & "/" & M & "/" & D & "/" & MakeRandom(N) & GetFileNameType '��/��/��/�����+��չ��
	  Case 3:GetFileName = Y & "-" & M & "-" & D & "/" & MakeRandom(N) & GetFileNameType '��-��-��/�����+��չ��
	  Case 4:GetFileName = Y & "/" & M & "/" & MakeRandom(N) & GetFileNameType '��/��/�����+��չ��
	  Case 5:GetFileName = Y & "-" & M & "/" & MakeRandom(N) & GetFileNameType '��-��/�����+��չ��
	  Case 12:GetFileName = Y & M & "/" & MakeRandom(N) & GetFileNameType '��-��/�����+��չ��
	  Case 6:GetFileName = Y & M & D & "/" & MakeRandom(N) & GetFileNameType '������/�����+��չ��
	  Case 7:GetFileName = Y & "/" & MakeRandom(N) & GetFileNameType '��/�����+��չ��
	  Case 8:GetFileName = Y & M & D & MakeRandom(N) & GetFileNameType '��+��+��+�����+��չ��
	  Case 9:GetFileName = MakeRandom(N) & GetFileNameType
	  Case 10:GetFileName = MakeRandomChar(N) & GetFileNameType '����ַ�
	  Case 11:GetFileName ="ID"
	  Case Else
	   GetFileName = Y & M & D & GetFileNameType '12λ�����+��չ��
	End Select
	End Function
	'**************************************************
	'��������Alert
	'��  �ã������ɹ���ʾ��
	'��  ����SuccessStr  ----�ɹ���ʾ��Ϣ
	'        Url   ------�ɹ���ʾ����"ȷ��"ת������
	'����ֵ����
	'**************************************************
	Public Function Alert(SuccessStr, Url)
	 If Url <> "" Then
	  echo ("<script language=""Javascript""> alert('" & SuccessStr & "');location.href='" & Url & "';</script>")
	 Else
	  echo ("<script language=""Javascript""> alert('" & SuccessStr & "');</script>")
	 End If
	End Function
	'**************************************************
	'��������AlertHistory
	'��  �ã�����������Ϣ��,ֹͣ����ҳ���ִ��,����n����
	'��  ����SuccessStr  ----�ɹ���ʾ��Ϣ
	'        n   ------���ؼ���
	'����ֵ����
	'**************************************************
	Public Function AlertHistory(SuccessStr, N)
		echo ("<script language=""Javascript""> alert('" & SuccessStr & "');history.back(" & N & ");</script>")
		die ""
	End Function
	'��ʾ�ɹ���������
	Sub AlertHintScript(SuccessStr)
	  echo "<script language=JavaScript>" & vbCrLf
	  echo "alert('" & SuccessStr & "');"
	  echo "location.replace('" & Request.ServerVariables("HTTP_REFERER") & "')" & vbCrLf
	  echo "</script>" & vbCrLf
	  die ""
	End Sub
	'**************************************************
	'��������Confirm
	'��  �ã������ɹ���ʾ��
	'��  ����SuccessStr  ----�ɹ���ʾ��Ϣ
	'        Url   ------�ɹ���ʾ����"ȷ��"ת������
	'        Url1   ------confirm����"ȡ��"ת������
	'����ֵ����
	'**************************************************
	Public Function Confirm(SuccessStr, Url, Url1)
	  echo ("<script language=""Javascript""> if (confirm('" & SuccessStr & "')){location.href='" & Url & "';}else{location.href='" & Url1 & "';}</script>")
	End Function
    
	Public Sub ShowTips(Action,Message)
		 Response.Redirect(Setting(3) & "Plus/error.asp?action="&action &"&message="&Server.URLEncode(message))
	End Sub
	'**************************************************
	'��������ShowError
	'��  �ã���ʾ������Ϣ��
	'��  ����Errmsg  ----������Ϣ
	'����ֵ����
	'**************************************************
	Public Sub ShowError(Errmsg)
		echo ("<br><br><div align=""center"">")
		echo ("  <center>")
		echo ("  <table border=""0"" cellpadding='2' cellspacing='1' class='border' width=""75%"" style=""MARGIN-TOP: 3px"" class='border'>")
		echo ("	 <tr class=tdbg>")
		echo ("			  <td width=""100%"" height=""30"" align=""center"">")
		echo ("				<b> " & Errmsg & "&nbsp; </b>")
		echo ("				</b>")
		echo ("			  </td>")
		echo ("	 </tr>")
		echo ("	 <tr  class=tdbg>")
		echo ("			  <td width=""100%"" height=""30"" align=""center"">")
		echo ("				<p><b><a href=""javascript:history.go(-1)"">...::: �� �� �� �� ")
		echo ("				:::...</a></b>")
		echo ("			  </td>")
		echo ("			</tr>")		
		echo ("  </table>")
		echo ("  </center>")
		echo ("</div>")
		die  ("")
    end sub
	'*****************************************************************************************
	'��������ReturnPowerResult
	'��  �ã�������Ȩ�ޡ�
	'��  ����ChannelID---����ϵͳ(Ƶ��) 1����ϵͳ2ͼƬϵͳ�� PowerOpName ---��ǰ������Ȩ������
	'����ֵ��������true,���򷵻�false
	'******************************************************************************************
	Public Function ReturnPowerResult(ChannelID, PowerOpName)
		If C("AdminName") = "" Then
			 ReturnPowerResult = False
			 Exit Function
		ElseIf C("SuperTF") = "1" Then    '����������ӵ������Ȩ��
			ReturnPowerResult = True
			Exit Function
		Else
		   If Instr(C("ModelPower"),C_S(ChannelID,10)&"0")>0 then          'û���κι���Ȩ
			ReturnPowerResult = False
		   ElseIf Instr(C("ModelPower"),C_S(ChannelID,10)&"1")>0 then      'ӵ������Ȩ��
			ReturnPowerResult = True
		   ElseIf Instr(C("ModelPower"),C_S(ChannelID,10)&"2")>0 then      '������Ŀ,ӵ�в���Ȩ��
			ReturnPowerResult = CheckPower(PowerOpName)
		   Else
			ReturnPowerResult = CheckPower(PowerOpName)
		   End If
	   End If
	End Function
	'�������ReturnPowerResult������ʹ��
	Public Function CheckPower(PowerOpName)
	        Dim PowerList, ModelPower
		    PowerList = Trim(C("PowerList"))
			If (PowerList <> "") And (PowerOpName <> "") Then
				Select Case Left(PowerOpName, 4)     '����Ƿ���ģ�����Ȩ��
				  Case "KMST" 'ϵͳ
				   If Instr(C("ModelPower"),"sysset0") >0 and C("SuperTF")<>"1" Then ModelPower = false else ModelPower=true
				  Case "KMUA" '�û�
				   If Instr(C("ModelPower"),"user0") >0 and C("SuperTF")<>"1" Then ModelPower = false else ModelPower=true
				  Case "KMTL"
				  If Instr(C("ModelPower"),"lab0")>0 and C("SuperTF")<>"1" Then ModelPower = false else ModelPower=true
				  Case "KSMM"
				  If Instr(C("ModelPower"),"model0")>0 and C("SuperTF")<>"1" Then ModelPower = false else ModelPower=true
				 ' Case "KSMS"
				 ' If Instr(C("ModelPower"),"subsys0")>0 and C("SuperTF")<>"1" Then ModelPower = false else ModelPower=true
				  Case Else
				   ModelPower = true
				End Select
				   If InStr(PowerList, PowerOpName) <> 0 And ModelPower Then
					 CheckPower = True:Exit Function
				   Else
					 CheckPower = False:Exit Function
				   End If
			Else
			   CheckPower = False:Exit Function
			End If
			
	End Function
	'�������ReturnPowerResult����ʹ��,     ReturnFlag  ----���� 0�ر�,1����ǰһҳ2,ת��URL, Url    -�����ת���Url
	Sub ReturnErr(ReturnFlag, Url)
	   If ReturnFlag = 0 Then
		 echo ("<script>alert('������ʾ:\n\n��û�д��������Ȩ��,����ϵͳ����Ա��ϵ!');window.close();</script>")
	   ElseIf ReturnFlag = 1 Then
		 echo ("<script>alert('������ʾ:\n\n��û�д��������Ȩ��,����ϵͳ����Ա��ϵ!');history.back();</script>")
	  ElseIf ReturnFlag = 2 Then
	     echo ("<script>alert('������ʾ:\n\n��û�д��������Ȩ��,����ϵͳ����Ա��ϵ!');location.href='" & Url & "';</script>")
	  End If
	End Sub
	'������վ��̨��־ , UserName --- ����Ա�˺� , ResultTF ---0��¼ʧ�� 1---��¼�ɹ� ,ScriptName---��¼·�� ,Descript---������Ϣ
	Sub InsertLog(UserName, ResultTF, ScriptName, Descript)
		Dim SystemStr:SystemStr = Request.ServerVariables("HTTP_USER_AGENT")
		If InStr(SystemStr, "Windows NT 5.2") Then
		  SystemStr = "Win2003"
		ElseIf InStr(SystemStr, "Windows NT 5.0") Then
		  SystemStr = "Win2000"
		ElseIf InStr(SystemStr, "Windows NT 5.1") Then
		  SystemStr = "WinXP"
		ElseIf InStr(SystemStr, "Windows NT") Then
		  SystemStr = "WinNT"
		ElseIf InStr(SystemStr, "Windows 9") Then
		  SystemStr = "Win9x"
		ElseIf InStr(SystemStr, "unix") Or InStr(SystemStr, "linux") Or InStr(SystemStr, "SunOS") Or InStr(SystemStr, "BSD") Then
		  SystemStr = "����Unix"
		ElseIf InStr(SystemStr, "Mac") Then
		  SystemStr = "Mac"
		Else
		  SystemStr = "Other"
		End If
		Conn.Execute("Insert into KS_Log(UserName,ResultTF,LoginTime,LoginOS,LoginIP,ScriptName,Description) values('" & UserName & "'," & ResultTF & "," & SqlNowString & ",'" & replace(SystemStr,"'","""") & "','" & getip & "','" & replace(scriptname,"'","""") & "','" & replace(descript,"'","""") & "')")
	End Sub
	
	'��ʾ��ҳ��ǰ����
	'����˵��:PageStyle-��ҳ��ʽ,ItemUnit-��λ,TotalPage-��ҳ��,CurrPage-��ǰ��Nҳ,TotalInfo-����Ϣ��,PerPageNumber-ÿҳ��ʾ��
	Function  GetPrePageList(PageStyle,ItemUnit,TotalPage,CurrPage,TotalInfo,PerPageNumber)
	    Select Case  Cint(PageStyle)
		  Case 1:GetPrePageList= "<div align=""right"" height=""25"" class=""fenye"" id=""fenye"">" & "�� " & TotalInfo & " " & ItemUnit &"  ҳ��:<font color=red> " & CurrPage & "</font>/" & TotalPage & "ҳ  " & PerPageNumber & " " & ItemUnit &"/ҳ "
		 Case 2:GetPrePageList= "<div align=""right"" height=""25"" class=""fenye"" id=""fenye"">��<font color=red>" & CurrPage & "</font>ҳ ��" & TotalPage & "ҳ "
		 Case 3:GetPrePageList= "<div align=""right"" height=""25"" class=""fenye"" id=""fenye"">��<font color=red>" & CurrPage & "</font>ҳ ��" & TotalPage & "ҳ "
		 Case 4:GetPrePageList= "<div align=""right"" height=""25"" class=""fenye"" id=""fenye"">"
	   End Select
	End Function
	'��̬��ʾ��ҳ
	 Function GetPageList(FileName,PageStyle,CurrPage,TotalPage, ShowTurnToFlag)
			Dim PageStr, I, J, SelectStr
			 If ChkClng(PageStyle)=0 Then PageStyle=1
			 Select Case PageStyle
			  Case 1
			   If CurrPage = 1 And CurrPage <> TotalPage Then
				PageStr = "��ҳ  ��һҳ <a href=""" & FileName & "&Page=" & CurrPage + 1 & """>��һҳ</a>  <a href= """ & FileName & "&Page=" & TotalPage & """>βҳ</a>"
			   ElseIf CurrPage = 1 And CurrPage = TotalPage Then
				PageStr = "��ҳ  ��һҳ ��һҳ βҳ"
			   ElseIf CurrPage = TotalPage And CurrPage <> 2 Then  '�������һҳ�պ��ǵڶ�ҳ��Ҫ�����⴦��
				 PageStr = "<a href=""" & FileName & """>��ҳ</a>  <a href=""" & FileName & "&Page=" & CurrPage - 1 & """>��һҳ</a> ��һҳ  βҳ"
			   ElseIf CurrPage = TotalPage And CurrPage = 2 Then
				 PageStr = "<a href=""" & FileName & """>��ҳ</a>  <a href=""" & FileName & """>��һҳ</a> ��һҳ  βҳ"
			   ElseIf CurrPage = 2 Then
				PageStr = "<a href=""" & FileName & """>��ҳ</a>  <a href=""" & FileName & """>��һҳ</a> <a href=""" & FileName & "&Page=" & CurrPage + 1 & """>��һҳ</a>  <a href= """ & FileName & "&Page=" &TotalPage & """>βҳ</a>"
			   Else
				PageStr = "<a href=""" & FileName & """>��ҳ</a>  <a href=""" & FileName & "&Page=" & CurrPage - 1 & """>��һҳ</a> <a href=""" & FileName & "&Page=" & CurrPage + 1 & """>��һҳ</a>  <a href= """ & FileName & "&Page=" & TotalPage & """>βҳ</a>"
			   End If
			 Case 2
			 	If CurrPage=1 Then
			     PageStr="<font face=webdings>9</font> <font face=webdings>7</font>"
				ElseIf CurrPage=2 Then
			     PageStr="<a href=""" & FileName & """ title=""��ҳ""><font face=webdings>9</font></a> <a href=""" & FileName & """ title=""��һҳ""><font face=webdings>7</font></a>"
				Else
				 PageStr="<a href=""" & FileName & """ title=""��ҳ""><font face=webdings>9</font></a> <a href=""" & FileName & "&Page="&  CurrPage - 1 &""" title=""��һҳ""><font face=webdings>7</font></a> "
				End If
				 dim startpage,n
				 startpage=1
				 if (CurrPage>=10) then startpage=(CurrPage\10-1)*10+CurrPage mod 10+2
				
				 For J=startpage To TotalPage
				    If J= CurrPage Then
				     PageStr=PageStr & " <a href=""#""><font color=red>" & J &"</font></a>"
				    Else
				     PageStr=PageStr & " <a href=""" & FileName & "&Page=" & J&""">" & J &"</a>"
					End If
					n=n+1
					if n>=10 then exit for
				 Next
				 If CurrPage=TotalPage Then
				  PageStr=PageStr & " <font face=webdings>8</font> <font face=webdings>:</font>"
				 Else
				  PageStr=PageStr & " <a href=""" & FileName & "&Page=" & CurrPage + 1 & """ title=""��һҳ""><font face=webdings>8</font></a> <a href=""" & FileName & "&Page=" & TotalPage & """><font face=webdings>:</font></a> "
				 End If
			 Case 3
			 	If CurrPage=1 Then
			     PageStr="<font face=webdings>9</font> <font face=webdings>7</font>"
				ElseIf CurrPage=2 Then
			     PageStr="<a href=""" & FileName & """ title=""��ҳ""><font face=webdings>9</font></a> <a href=""" & FileName & """ title=""��һҳ""><font face=webdings>7</font></a>"
				Else
				 PageStr="<a href=""" & FileName & """ title=""��ҳ""><font face=webdings>9</font></a> <a href=""" & FileName & "&Page="&  CurrPage - 1 &""" title=""��һҳ""><font face=webdings>7</font></a> "
				End If
				 If CurrPage=TotalPage Then
				  PageStr=PageStr & " <font face=webdings>8</font> <font face=webdings>:</font>"
				 Else
				  PageStr=PageStr & " <a href=""" & FileName & "&Page=" & CurrPage + 1 & """ title=""��һҳ""><font face=webdings>8</font></a> <a href=""" & FileName & "&Page=" & TotalPage & """><font face=webdings>:</font></a> "
				 End If
			 case 4 
				 n=0:startpage=1
				 pageStr=pageStr & "<table border=""0"" align=""right""><tr><td>" & vbcrlf
				 if (CurrPage>1) then pageStr=PageStr & "<a href=""" & FileName & "&page=" & CurrPage-1 & """ class=""prev"">��һҳ</a>"
				 if (CurrPage<>TotalPage) then pageStr=PageStr & "<a href=""" & FileName & "&page=" & CurrPage+1 & """ class=""next"">��һҳ</a>"
				 pageStr=pageStr & "<a href=""" & FileName & """ class=""prev"">�� ҳ</a>"
				 if (CurrPage>=7) then startpage=CurrPage-5
				 if TotalPage-CurrPage<5 Then startpage=TotalPage-10
				 If startpage<0 Then startpage=1
				 For J=startpage To TotalPage
				    If J= CurrPage Then
				     PageStr=PageStr & " <a href=""#"" class=""curr""><font color=red>" & J &"</font></a>"
				    Else
				     PageStr=PageStr & " <a class=""num"" href=""" & FileName & "&page=" & J&""">" & J &"</a>"
					End If
					n=n+1
					if n>=10 then exit for
				 Next
				 pageStr=pageStr & "<a href=""" & FileName & "&page=" & TotalPage &""" class=""prev"">ĩҳ</a>"
				 pageStr=PageStr & " <span>�ܹ�" & TotalPage & "ҳ</span></td></tr></table>"
			 End Select
			   
			   If CBool(ShowTurnToFlag) = True and pagestyle<>4 Then
				  PageStr = PageStr & " ת����<select name=""page"" size=""1"" onchange=""javascript:window.location=this.options[this.selectedIndex].value;"">"
				  For J = 1 To TotalPage
				   If J = CurrPage Then
					 SelectStr = " selected"
				   Else
					 SelectStr = ""
				   End If
				   If J = 1 Then
					 PageStr = PageStr & "<option value=""" & FileName & """" & SelectStr & ">��" & J & "ҳ</option>"
				   Else
					 PageStr = PageStr & "<option value=""" & FileName & "&Page=" & J & """" & SelectStr & ">��" & J & "ҳ</option>"
				   End If
			   Next
				  PageStr = PageStr & "</select>"
			   End If
			   	GetPageList=PageStr	&"</div>"	   
		End Function
		
	'��ʾα��̬��ҳ
	 Function GetStaticPageList(FileName,PageStyle,CurrPage,TotalPage, ShowTurnToFlag,Extension)
			Dim PageStr, I, J, SelectStr
			 If ChkClng(PageStyle)=0 Then PageStyle=1
			 Select Case PageStyle
			  Case 1
			   If CurrPage = 1 And CurrPage <> TotalPage Then
				PageStr = "��ҳ  ��һҳ <a href=""" & FileName & CurrPage + 1 & Extension & """>��һҳ</a>  <a href= """ & FileName & TotalPage & Extension & """>βҳ</a>"
			   ElseIf CurrPage = 1 And CurrPage = TotalPage Then
				PageStr = "��ҳ  ��һҳ ��һҳ βҳ"
			   ElseIf CurrPage = TotalPage And CurrPage <> 2 Then  '�������һҳ�պ��ǵڶ�ҳ��Ҫ�����⴦��
				 PageStr = "<a href=""" & FileName & "1" & Extension & """>��ҳ</a>  <a href=""" & FileName & CurrPage - 1 & Extension & """>��һҳ</a> ��һҳ  βҳ"
			   ElseIf CurrPage = TotalPage And CurrPage = 2 Then
				 PageStr = "<a href=""" & FileName & "1" & Extension & """>��ҳ</a>  <a href=""" & FileName & "1" & Extension & """>��һҳ</a> ��һҳ  βҳ"
			   ElseIf CurrPage = 2 Then
				PageStr = "<a href=""" & FileName & "1" & Extension & """>��ҳ</a>  <a href=""" & FileName & "1" & Extension & """>��һҳ</a> <a href=""" & FileName & CurrPage + 1 & Extension & """>��һҳ</a>  <a href= """ & FileName & TotalPage & Extension & """>βҳ</a>"
			   Else
				PageStr = "<a href=""" & FileName & "1" & Extension & """>��ҳ</a>  <a href=""" & FileName & CurrPage - 1 & Extension & """>��һҳ</a> <a href=""" & FileName & CurrPage + 1 & Extension & """>��һҳ</a>  <a href= """ & FileName & TotalPage & Extension & """>βҳ</a>"
			   End If
			 Case 2
			 	If CurrPage=1 Then
			     PageStr="<font face=webdings>9</font> <font face=webdings>7</font>"
				'ElseIf CurrPage=2 Then
			   '  PageStr="<a href=""" & FileName & "1" & Extension & """ title=""��ҳ""><font face=webdings>9</font></a> <a href=""" & FileName & "-" & Extension &""" title=""��һҳ""><font face=webdings>7</font></a>"
				Else
				 PageStr="<a href=""" & FileName &"1"& Extension&""" title=""��ҳ""><font face=webdings>9</font></a> <a href=""" & FileName & CurrPage - 1 & Extension&""" title=""��һҳ""><font face=webdings>7</font></a> "
				End If
				 dim startpage,n
				 startpage=1
				 if (CurrPage>=10) then startpage=(CurrPage\10-1)*10+CurrPage mod 10+2
				
				 For J=startpage To TotalPage
				    If J= CurrPage Then
				     PageStr=PageStr & " <a href=""#""><font color=red>" & J &"</font></a>"
				    Else
				     PageStr=PageStr & " <a href=""" & FileName & J& Extension&""">" & J &"</a>"
					End If
					n=n+1
					if n>=10 then exit for
				 Next
				 If CurrPage=TotalPage Then
				  PageStr=PageStr & " <font face=webdings>8</font> <font face=webdings>:</font>"
				 Else
				  PageStr=PageStr & " <a href=""" & FileName & CurrPage + 1& Extension & """ title=""��һҳ""><font face=webdings>8</font></a> <a href=""" & FileName & TotalPage & Extension& """><font face=webdings>:</font></a> "
				 End If
			 Case 3
			 	If CurrPage=1 Then
			     PageStr="<font face=webdings>9</font> <font face=webdings>7</font>"
				ElseIf CurrPage=2 Then
			     PageStr="<a href=""" & FileName &"1" & Extension & """ title=""��ҳ""><font face=webdings>9</font></a> <a href=""" & FileName & "1"  & Extension & """ title=""��һҳ""><font face=webdings>7</font></a>"
				Else
				 PageStr="<a href=""" & FileName & "1" & Extension & """ title=""��ҳ""><font face=webdings>9</font></a> <a href=""" & FileName & CurrPage - 1 & Extension &""" title=""��һҳ""><font face=webdings>7</font></a> "
				End If
				 If CurrPage=TotalPage Then
				  PageStr=PageStr & " <font face=webdings>8</font> <font face=webdings>:</font>"
				 Else
				  PageStr=PageStr & " <a href=""" & FileName & CurrPage + 1 & Extension & """ title=""��һҳ""><font face=webdings>8</font></a> <a href=""" & FileName & TotalPage & Extension & """><font face=webdings>:</font></a> "
				 End If
			 Case 4
			     n=0:startpage=1
				 pageStr=pageStr & "<table border=""0"" align=""right""><tr><td>" & vbcrlf
				 if (CurrPage>1) then pageStr=PageStr & "<a href=""" & FileName & CurrPage - 1 & Extension & """ class=""prev"">��һҳ</a>"
				 if (CurrPage<>TotalPage) then pageStr=PageStr & "<a href=""" & FileName & CurrPage + 1 & Extension &""" class=""next"">��һҳ</a>"
				 pageStr=pageStr & "<a href=""" & FileName &"1" & Extension & """ class=""prev"">�� ҳ</a>"
				 if (CurrPage>=7) then startpage=CurrPage-5
				 if TotalPage-CurrPage<5 Then startpage=TotalPage-10
				 If startpage<0 Then startpage=1 
				 For J=startpage To TotalPage
				    If J= CurrPage Then
				     PageStr=PageStr & " <a href=""#"" class=""curr""><font color=red>" & J &"</font></a>"
				    Else
				     PageStr=PageStr & " <a class=""num"" href=""" & FileName & J& Extension&""">" & J &"</a>"
					End If
					n=n+1
					if n>=10 then exit for
				 Next
				 pageStr=pageStr & "<a href=""" & FileName & TotalPage & Extension &""" class=""prev"">ĩҳ</a>"
				 pageStr=PageStr & " <span>�ܹ�" & TotalPage & "ҳ</span></td></tr></table>"
			 End Select
			   
			   If CBool(ShowTurnToFlag) = True and pageStyle<>4 Then
				  PageStr = PageStr & " ת����<select name=""page"" size=""1"" onchange=""javascript:window.location=this.options[this.selectedIndex].value;"">"
				  For J = 1 To TotalPage
				   If J = CurrPage Then
					 SelectStr = " selected"
				   Else
					 SelectStr = ""
				   End If
				   If J = 1 Then
					 PageStr = PageStr & "<option value=""" & FileName & "1" & Extension & """" & SelectStr & ">��" & J & "ҳ</option>"
				   Else
					 PageStr = PageStr & "<option value=""" & FileName & J & Extension & """" & SelectStr & ">��" & J & "ҳ</option>"
				   End If
			   Next
				  PageStr = PageStr & "</select>"
			   End If
			   	GetStaticPageList=PageStr	&"</div>"	   
      End Function
	'*************************************************************************************
	'������:GetClassID
	'��  ��:������Ŀ¼��Ƶ����ID��,����Ŀ¼ID ��+10λ���
	'��  ��:��
	'*************************************************************************************
	Public Function GetClassID()
		Do While True
		 GetClassID = Year(Now()) & MakeRandom(10)
		 If Conn.Execute("Select ID from KS_Class Where ID='" & GetClassID & "'").Eof Then Exit Do
		Loop
	End Function
	
	'ȡר��������
	Function GetSpecialClass(ClassID,FieldName)
	  If Not IsObject(Application(SiteSN & "_SpecialClass")) then
			Dim Rs:Set Rs = Conn.Execute("Select ClassID,ClassName,ClassEname,Descript,FsoIndex From KS_SpecialClass Order By ClassID")
			Set Application(SiteSN & "_SpecialClass")=RsToxml(Rs,"row","root")
			Set Rs = Nothing
	  End If
	  Dim Node:Set Node=Application(SiteSN&"_SpecialClass").documentElement.selectSingleNode("row[@classid=" & ClassID & "]/@" & Lcase(FieldName) & "")
	  If Not Node Is Nothing  Then GetSpecialClass=Node.text
	  Set Node = Nothing
	End Function
	
	'���빩������
	Sub LoadGQTypeToXml()
	  If Not IsObject(Application(SiteSN & "_SupplyType")) then
			Dim Rs:Set Rs = Conn.Execute("Select TypeID,TypeName,TypeColor From KS_GQType Order By TypeID")
			Set Application(SiteSN & "_SupplyType")=RsToxml(Rs,"row","SupplyType")
			Set Rs = Nothing
	  End If
	End Sub
	
	'*************************************************************************************
	'������:GetGQTypeName
	'��  ��:��ù���Ľ����������
	'��  ��:TypeID
	'*************************************************************************************
	Public Function GetGQTypeName(TypeID)
	   If Not IsNumeric(TypeID) Then GetGQTypeName="":Exit Function
	   LoadGQTypeToXml()
	   Dim NodeName,NodeColor
	   Set NodeName=Application(SiteSN & "_SupplyType").documentElement.SelectSingleNode("row[@typeid=" & TypeID & "]/@typename")
	   If Not NodeName  Is Nothing Then
		 Set NodeColor=Application(SiteSN & "_SupplyType").documentElement.SelectSingleNode("row[@typeid=" & TypeID & "]/@typecolor")
		 GetGQTypeName="<span style=""color:" & NodeColor.Text & """>" & NodeName.Text & "</span>"
	   End If 
	End Function
	'���ع����������б�
	'������Flag:1-��ǩ���� 0-�����Ϣʱ����
	Public Function ReturnGQType(SelID,Flag)
	   Dim Node
	   LoadGQTypeToXml()
	    If Flag=1 Then 
	   	   ReturnGQType="<select class=""textbox"" name=""TypeID"" id=""TypeID"" style=""width:70%"">"
	        If SelID = "0" Then ReturnGQType=ReturnGQType & "<option  value=""0"" selected>- �������Ͳ��� -</option>"	Else ReturnGQType=ReturnGQType & "<option  value=""0"">- �������Ͳ��� -</option>"
	   Else
	   	   ReturnGQType="<select class=""textbox"" name=""TypeID"">"
	   End If
	   For Each Node In Application(SiteSN & "_SupplyType").DocumentElement.SelectNodes("row")
	     If trim(SelID)=trim(node.SelectSingleNode("@typeid").text) Then
			 ReturnGQType=ReturnGQType & "<option value=""" & node.SelectSingleNode("@typeid").text & """ style=""color:" & node.SelectSingleNode("@typecolor").text & """ selected>" & node.SelectSingleNode("@typename").text & "</option>"
		 else
			  ReturnGQType=ReturnGQType & "<option value=""" & node.SelectSingleNode("@typeid").text & """ style=""color:" & node.SelectSingleNode("@typecolor").text & """>" & node.SelectSingleNode("@typename").text & "</option>"
		 end if
       Next
	   ReturnGQType=ReturnGQType & "</select>"
	End Function
	
	'*************************************************************************************
	'������:GetInfoID
	'��  ��:��������,ͼƬ�����صȵ�ΨһID
	'��  ��:ChannelID--Ƶ��ID
	'*************************************************************************************
	Public Function GetInfoID(ChannelID)
	   On Error Resume Next
	   Dim RSC, TableNameStr
       Set RSC=Server.CreateObject("ADODB.RECORDSET")
	   TableNameStr = "Select ProID From " & C_S(ChannelID,2) & " Where ProID='"
	   Do While True
		 GetInfoID = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Now(), "-", ""), " ", ""), ":", ""), "PM", ""), "AM", ""), "����", ""), "����", "") & MakeRandom(3)
			RSC.Open TableNameStr & GetInfoID & "'", Conn, 1, 1
			If RSC.EOF And RSC.BOF Then Exit Do
	   Loop
		RSC.Close:Set RSC = Nothing
	End Function
	'*************************************************************************************
	'������:ReplaceInnerLink
	'��  ��:�滻վ������
	'��  ��:Content-���滻����
	'*************************************************************************************
	Public Function ReplaceInnerLink(ByVal Content)
	  'Content=HTMLCode(Content)
	  If Not IsObject(Application(SiteSN & "_InnerLink")) then
			Dim Rs:Set Rs = Conn.Execute("Select Title,Url,OpenType,CaseTF,Times,Start From KS_InnerLink Where OpenTF=1 Order By ID")
			Set Application(SiteSN & "_InnerLink")=RecordsetToxml(Rs,"InnerLink","InnerLinkList")
			Set Rs = Nothing
	  end if
		Dim Node,CaseTF,Times,Inti,DLocation,XLocation,StrReplace,CurrentTimes,SourceStr
		For Each Node In Application(SiteSN & "_InnerLink").DocumentElement.SelectNodes("InnerLink")
			 CurrentTimes=0
			 Dim OpenTypeStr:OpenTypeStr = G_O_T_S(Node.selectSingleNode("@ks2").text)
			 CaseTF=Cint(Node.selectSingleNode("@ks3").text)
			 Times=Cint(Node.selectSingleNode("@ks4").text)
			 Inti=ChkClng(Node.selectSingleNode("@ks5").text)
			 StrReplace=Node.selectSingleNode("@ks0").text
			 If Inti=0 Then Inti=1
			If InStr(Inti,Content,StrReplace,CaseTF)>0 Then
			  Do While instr(Inti,Content,StrReplace,CaseTF)<>0
			    Inti=instr(Inti,Content,StrReplace,CaseTF)
				If Inti<>0 then
				  DLocation=instr(Inti,Content,">") '���滻��><֮��Ĺؼ���
				  XLocation=instr(Inti,Content,"<")
				  If DLocation >= XLocation Then
					Content=left(Content,Inti-1) & "<a href="""&Node.selectSingleNode("@ks1").text&"""" & OpenTypeStr & " class=""innerlink"">"&Node.selectSingleNode("@ks0").text&"</a>" & mid(Content,Inti+len(StrReplace))
					Inti=Inti+len("<a href="""&Node.selectSingleNode("@ks1").text&"""" & OpenTypeStr & " class=""innerlink"">"&StrReplace&"</a>")
					CurrentTimes=CurrentTimes+1
					If Times<>-1 And CurrentTimes>= Times Then Exit Do
				 Else
				    Inti=Inti+len(StrReplace)
				 End If
			   End If
			  Loop	
			End if
		Next
	 ReplaceInnerLink = Content
	End Function
	
	'=============================================================
	'�������ã��ж���ԴURL�Ƿ������ⲿ
	'=============================================================
	Public Function CheckOuterUrl()
		On Error Resume Next
		Dim server_v1, server_v2
		server_v1 = Replace(LCase(Trim(Request.ServerVariables("HTTP_REFERER"))), "http://", "")
		server_v2 = LCase(Trim(Request.ServerVariables("SERVER_NAME")))
		CheckOuterUrl = True
		If Mid(server_v1,8,len(server_v2))=server_v2 Then CheckOuterUrl=False 
	End Function 
	
	'����
	Function Encrypt(ecode)
	dim texts,i
	for i=1 to len(ecode)
	texts=texts & chr(asc(mid(ecode,i,1))+3)
	next
	Encrypt = texts
	End Function
	'����
	Function Decrypt(dcode)
	 If IsNul(dcode) then exit function
	dim texts,i
	for i=1 to len(dcode)
	texts=texts & chr(asc(mid(dcode,i,1))-3)
	next
	Decrypt=texts
	End Function
	'ƥ�� img src,�����|���� 
	Function GetImgSrcArr(strng) 
	If strng="" Or IsNull(strng) Then GetImgSrcArr="":Exit Function
	Dim regEx,Match,Matches,values
	Set regEx = New RegExp
	regEx.Pattern = "src\=.+?\.(gif|jpg)"
	regEx.IgnoreCase = true 
	regEx.Global = True 
	Set Matches = regEx.Execute(strng)
	For Each Match in Matches
		if instr(lcase(Match.Value),"fileicon")=0 then
		 values=values&Match.Value&"|" 
		end if
	Next 
	GetImgSrcArr = Replace(Replace(Replace(Replace(values,"'",""),"""",""),"src=",""),Setting(2),"")
	If GetImgSrcArr<>"" Then GetImgSrcArr = left(GetImgSrcArr,len(GetImgSrcArr)-1)
	End Function
	

	'ȡ��Request.Querystring �� Request.Form ��ֵ
	Public Function G(Str)
	 G = Replace(Replace(Request(Str), "'", ""), """", "")
	End Function
	Function DelSql(Str)
		Dim SplitSqlStr,SplitSqlArr,I
		SplitSqlStr="dbcc|alter|drop|*|and |exec|or |insert|select|delete|update|count |master|truncate|declare|char|mid|chr|set |where|xp_cmdshell"
		SplitSqlArr = Split(SplitSqlStr,"|")
		For I=LBound(SplitSqlArr) To Ubound(SplitSqlArr)
			If Instr(LCase(Str),SplitSqlArr(I))>0 Then
				Die "<script>alert('ϵͳ���棡\n\n1�����ύ�������ж����ַ�" & SplitSqlArr(I) &";\n2�����������Ѿ�����¼;\n3������IP��"&GetIP&";\n4���������ڣ�"&Now&";\n		Powered By Kesion.Com!');window.close();</script>"
			End if
		Next
		DelSql = Str
    End Function
	'ȡ��Request.Querystring �� Request.Form ��ֵ
	Public Function S(Str)
	 S = DelSql(Replace(Replace(Request(Str), "'", ""), """", ""))
	End Function
	'��Cookiesֵ
	Public Function C(Str)
	 C=DelSql(Request.Cookies(SiteSN)(Str))
	End Function
	
	'ȡ��QueryString,��Form��������,����NoCollect��ʾ���ռ����ֶ�,�����Ӣ�Ķ��Ÿ���
	Function QueryParam(NoCollect)
		 Dim Param,R
		 For Each r In Request.QueryString
		  If FoundInArr(NoCollect,R,",")=false Then
			  If Request.QueryString(r)<>"" Then
				If Param="" Then
				 Param=r & "=" & Server.UrlEncode(Trim(Request.QueryString(r)))
				Else
				 Param=Param & "&" & r & "=" & Server.UrlEncode(Trim(Request.QueryString(r)))
				End If
			  End If
		 End If
		 Next
		' If Param<>"" Then QueryParam=Param:Exit Function
		 For Each r In Request.Form
		  If FoundInArr(NoCollect,R,",")=false Then
			  If Request.Form(r)<>"" Then
				If Param="" Then
				 Param=r & "=" & Server.UrlEncode(Trim(Request.Form(r)))
				Else
				 Param=Param & "&" & r & "=" & Server.UrlEncode(Trim(Request.Form(r)))
				End If
			  End If
		 End If
		 Next
		 QueryParam=Param
	End Function


	
   	'���нű�����
	Function CheckScript(byVal Content)
		Dim oRegExp,oMatch,spamCount
		Set oRegExp = New Regexp
		oRegExp.IgnoreCase = True
		oRegExp.Global = True
		oRegExp.pattern ="<script(.|\n)+?/script>"
		Content=oRegExp.replace(Content,"")
		Set oRegExp=Nothing
		CheckScript=Content
	End Function


	'�رղɼ����ݿ����
	Public Sub CloseConnItem()
	   On Error Resume Next
	   If IsObject(ConnItem) Then
		 ConnItem.Close:Set ConnItem = Nothing
	   End If
	End Sub
	'�����Զ���ҳ
	'������Content-�������� SplitPageStr-���·ָ��� PerPageLen-ÿҳ��Լ�ַ���
	Function AutoSplitPage(Content,SplitPageStr,maxPagesize)
	    Dim sContent,ss,i,IsCount,c,iCount,strTemp,Temp_String,Temp_Array
		sContent=Content
		If maxPagesize<100 Or Len(sContent)<maxPagesize+100 Then
			AutoSplitPage=sContent
		End If
		sContent=Replace(sContent, SplitPageStr, "")
		sContent=Replace(sContent, "&nbsp;", "<&nbsp;>")
		sContent=Replace(sContent, "&gt;", "<&gt;>")
		sContent=Replace(sContent, "&lt;", "<&lt;>")
		sContent=Replace(sContent, "&quot;", "<&quot;>")
		sContent=Replace(sContent, "&#39;", "<&#39;>")
		If sContent<>"" and maxPagesize<>0 and InStr(1,sContent,SplitPageStr)=0 then
			IsCount=True:Temp_String=""
			For i= 1 To Len(sContent)
				c=Mid(sContent,i,1)
				If c="<" Then
					IsCount=False
				ElseIf c=">" Then
					IsCount=True
				Else
					If IsCount=True Then
						'If Abs(Asc(c))>255 Then
						'	iCount=iCount+2
						'Else
							iCount=iCount+1
						'End If
						If iCount>=maxPagesize And i<Len(sContent) Then
							strTemp=Left(sContent,i)
							If CheckPagination(strTemp,"table|a|b>|i>|strong|div|span") then
								Temp_String=Temp_String & Trim(CStr(i)) & "," 
								iCount=0
							End If
						End If
					End If
				End If	
			Next
			If Len(Temp_String)>1 Then Temp_String=Left(Temp_String,Len(Temp_String)-1)
			Temp_Array=Split(Temp_String,",")
			For i = UBound(Temp_Array) To LBound(Temp_Array) Step -1
				ss = Mid(sContent,Temp_Array(i)+1)
				If Len(ss) > 100 Then
					sContent=Left(sContent,Temp_Array(i)) & SplitPageStr & ss
				Else
					sContent=Left(sContent,Temp_Array(i)) & ss
				End If
			Next
		End If
		sContent=Replace(sContent, "<&nbsp;>", "&nbsp;")
		sContent=Replace(sContent, "<&gt;>", "&gt;")
		sContent=Replace(sContent, "<&lt;>", "&lt;")
		sContent=Replace(sContent, "<&quot;>", "&quot;")
		sContent=Replace(sContent, "<&#39;>", "&#39;")
		AutoSplitPage=sContent
	End Function
    '������Ϻ���ʹ��
	Private Function CheckPagination(strTemp,strFind)
		Dim i,n,m_ingBeginNum,m_intEndNum
		Dim m_strBegin,m_strEnd,FindArray
		strTemp=LCase(strTemp)
		strFind=LCase(strFind)
		If strTemp<>"" and strFind<>"" then
			FindArray=split(strFind,"|")
			For i = 0 to Ubound(FindArray)
				m_strBegin="<"&FindArray(i)
				m_strEnd  ="</"&FindArray(i)
				n=0
				do while instr(n+1,strTemp,m_strBegin)<>0
					n=instr(n+1,strTemp,m_strBegin)
					m_ingBeginNum=m_ingBeginNum+1
				Loop
				n=0
				do while instr(n+1,strTemp,m_strEnd)<>0
					n=instr(n+1,strTemp,m_strEnd)
					m_intEndNum=m_intEndNum+1
				Loop
				If m_intEndNum=m_ingBeginNum then
					CheckPagination=True
				Else
					CheckPagination=False
					Exit Function
				End If
			Next
		Else
			CheckPagination=False
		End If
	End Function
	Public Function HTMLEncode(fString)
		If Not IsNull(fString) then
		    fString = ClearBadChr(fString)
			fString = Replace(fString, "&", "&amp;")
			fString = Replace(fString, "'", "&#39;")
			fString = Replace(fString, ">", "&gt;")
			fString = Replace(fString, "<", "&lt;")
			fString = Replace(fString, Chr(32), " ")
			fString = Replace(fString, Chr(9), " ")
			fString = Replace(fString, Chr(34), "&quot;")
			fString = Replace(fString, Chr(39), "&#39;")
			fString = Replace(fString, Chr(13), "")
			'fString = Replace(fString, " ", "&nbsp;")
			'fString = Replace(fString, Chr(10), "<br />")
		HTMLEncode = fString
		End If
	End Function
	
	Function ClearBadChr(str)
	  If Str<>"" Then
	     Dim re:Set re=new RegExp
		re.IgnoreCase =True
		re.Global=True
		re.Pattern="(on(load|click|dbclick|mouseover|mouseout|mousedown|mouseup|mousewheel|keydown|submit|change|focus)=""[^""]+"")"
		str = re.Replace(str, "")
		re.Pattern="((name|id|class)=""[^""]+"")"
		str = re.Replace(str, "")
		re.Pattern = "(<s+cript[^>]*?>([\w\W]*?)<\/s+cript>)"
		str = re.Replace(str, "")
		re.Pattern = "(<iframe[^>]*?>([\w\W]*?)<\/iframe>)"
		str = re.Replace(str, "")
		re.Pattern = "(<p>&nbsp;<\/p>)"
		str = re.Replace(str, "")
		Set re=Nothing
		ClearBadChr = str
	 End If	
	End Function

	
	Public Function HTMLCode(HtmlStr)
		If Not IsNul(HtmlStr) then
		'HtmlStr = Replace(HtmlStr, "&nbsp;", " ")
		HtmlStr = Replace(HtmlStr, "&quot;", Chr(34))
		HtmlStr = Replace(HtmlStr, "&#39;", Chr(39))
		HtmlStr = Replace(HtmlStr, "&#123;", Chr(123))
		HtmlStr = Replace(HtmlStr, "&#125;", Chr(125))
		HtmlStr = Replace(HtmlStr, "&#36;", Chr(36))
		HtmlStr = Replace(HtmlStr, "&amp;", "&")
		'HtmlStr = Replace(HtmlStr, vbCrLf, "")

		HtmlStr = Replace(HtmlStr, "&gt;", ">")
		HtmlStr = Replace(HtmlStr, "&lt;", "<")
		
		HTMLCode = HtmlStr
		End If
	End Function
	

	Public Function IsObjInstalled(strClassString)
		On Error Resume Next
		IsObjInstalled = False
		Err = 0
		Dim xTestObj:Set xTestObj = Server.CreateObject(strClassString)
		If 0 = Err Then IsObjInstalled = True
		Set xTestObj = Nothing
		Err = 0
	End Function
	Public Function IsExpired(strClassString)
		On Error Resume Next
		IsExpired = True
		Err = 0
		Dim xTestObj:Set xTestObj = Server.CreateObject(strClassString)
	
		If 0 = Err Then
			Select Case strClassString
				Case "Persits.Jpeg"
					If xTestObjResponse.Expires > Now Then
						IsExpired = False
					End If
				Case "wsImage.Resize"
					If InStr(xTestObj.errorinfo, "�Ѿ�����") = 0 Then
						IsExpired = False
					End If
				Case "SoftArtisans.ImageGen"
					xTestObj.CreateImage 500, 500, RGB(255, 255, 255)
					If Err = 0 Then
						IsExpired = False
					End If
			End Select
		End If
		Set xTestObj = Nothing
		Err = 0
	End Function
	Public Function ExpiredStr(I)
		   Dim ComponentName(3)
			ComponentName(0) = "Persits.Jpeg"
			ComponentName(1) = "wsImage.Resize"
			ComponentName(2) = "SoftArtisans.ImageGen"
			ComponentName(3) = "CreatePreviewImage.cGvbox"
			If IsObjInstalled(ComponentName(I)) Then
				If IsExpired(ComponentName(I)) Then
					ExpiredStr = "�����ѹ���"
				Else
					ExpiredStr = ""
				End If
			  ExpiredStr = " ��֧��" & ExpiredStr
			Else
			  ExpiredStr = "����֧��"
			End If
	End Function

  
  '======================================��Ա��غ���====================================
    'ȡ�û�Ա��ѡ��--�����б�  ������Selected--Ĭ��ѡ��
	Public Function GetUserGroup_Option(Selected)
	 Dim RSObj:Set RSObj=Server.CreateObject("Adodb.Recordset")
	  RSObj.Open "Select ID,GroupName From KS_UserGroup",Conn,1,1
	  	Do While Not RSObj.Eof
		   IF Selected=RSObj(0) Then
			GetUserGroup_Option=GetUserGroup_Option & "<option value=""" & RSObj(0) & """ Selected>" & RSObj(1) & "</option>"
		   Else
			GetUserGroup_Option=GetUserGroup_Option & "<option value=""" & RSObj(0) & """>" & RSObj(1) & "</option>"
		   End If
		RSObj.MoveNext
		Loop
	  RSObj.Close:Set RSObj=Nothing
	End Function
	 'ȡ�û�Ա��ѡ��--��ѡ�б� ������SelectArr--Ĭ��ѡ������","����,RowNum--ÿ����ʾѡ����
	Public Function GetUserGroup_CheckBox(OptionName,SelectArr,RowNum)
	   Dim n:n=0
	   Dim RSObj:Set RSObj=Server.CreateObject("Adodb.Recordset")
	   IF RowNum<=0 Then RowNum=3
	   RSObj.Open "Select ID,GroupName From KS_UserGroup",Conn,1,1
	   GetUserGroup_CheckBox="<table width=""100%"" align=""center"" border=""0"">"
	   Do While Not RSObj.Eof
	        GetUserGroup_CheckBox=GetUserGroup_CheckBox & "<TR>"
	     For N=1 To RowNum
		    GetUserGroup_CheckBox=GetUserGroup_CheckBox & "<TD WIDTH=""" & CInt(100 / CInt(RowNum)) & "%"">"
			If FoundInArr(SelectArr,RSObj(0),",")<>0 Then
			 GetUserGroup_CheckBox=GetUserGroup_CheckBox & "<input type=""checkbox"" checked name=""" & OptionName & """ value=""" & RSObj(0) & """>" & RSObj(1) & "&nbsp;&nbsp;"
			Else
			 GetUserGroup_CheckBox=GetUserGroup_CheckBox & "<input type=""checkbox"" name=""" & OptionName & """ value=""" & RSObj(0) & """>" & RSObj(1) & "&nbsp;&nbsp;"
			End IF
		 GetUserGroup_CheckBox=GetUserGroup_CheckBox & "</TD>"
		 		RSObj.MoveNext
				If RSObj.Eof Then Exit For
		Next
		GetUserGroup_CheckBox=GetUserGroup_CheckBox & "</TR>"
		If RSObj.Eof Then Exit Do
	   Loop
	   GetUserGroup_CheckBox=GetUserGroup_CheckBox & "</TABLE>"
	   RSObj.Close:Set RSObj=Nothing
	End Function
	 
  	'ȡ���û�������
	Public Function GetUserGroupName(GroupID)
	 On Error Resume Next
	 GetUserGroupName=Conn.Execute("Select GroupName From KS_UserGroup Where ID=" & GroupID)(0)
	 if err then GetUserGroupName=""
	End Function
    
	'��ԱͶ�����£�ͼƬ�����ص����ӻ���,����վ�ڶ��Ų���
	'����ChannelID-Ƶ��ID,UserName---�û�����,InfoTitle---Ͷ�������
	Public Sub SignUserInfoOK(ChannelID,UserName,InfoTitle,InfoID)
	    IF Not IsNumeric(ChannelID) Then Exit Sub
	    Dim ClientName,GroupID,RSObj:Set RSObj=Conn.Execute("Select top 1 RealName,GroupID From KS_User Where UserName='" & UserName & "'")
		IF Not RSObj.Eof Then
					ClientName=RSObj(0):If ClientName="" Then ClientName=UserName
					GroupID=RSObj(1)
					Dim ScoreRate:ScoreRate=ChkClng(U_S(GroupID,3))
					Dim PointRate:PointRate=ChkClng(U_S(GroupID,4))
					Dim MoneyRate:MoneyRate=ChkClng(U_S(GroupID,5))
					
					'�ɹ�����վ��֪ͨ�ż�
					Dim Sender:Sender=Setting(0)
					Dim Title:Title="��ϲ���������" & C_S(ChannelID,3) & "[" & InfoTitle & "]��ͨ����ˣ�����"
					Dim Message:Message="" & C_S(ChannelID,3) & "���⣺" & InfoTitle &" ��ͨ�����!<br>"
					
					If Conn.Execute("Select top 1 * From KS_LogMoney Where UserName='" & UserName & "' and ChannelID=" & ChannelID & " and InfoID=" & InfoID).Eof And C_S(ChannelID,18)*MoneyRate<>0 Then    'û�м�¼�Ÿ����ӽ�Ǯ
					 If C_S(ChannelID,18)>0 Then
					  Message = Message & "��ý�Ǯ��<font color=red>" & C_S(ChannelID,18)*MoneyRate & "</font> Ԫ�����<br>"
					 ElseIf C_S(ChannelID,18)<0 Then
					  Message = Message & "���Ľ�Ǯ��<font color=red>" & Abs(C_S(ChannelID,18))*MoneyRate & "</font> Ԫ�����<br>"
					 End IF
					End If
					 
					If Conn.Execute("Select top 1 * From KS_LogPoint Where UserName='" & UserName & "' and ChannelID=" & ChannelID & " and InfoID=" & InfoID & " and ContributeFlag=1").Eof And C_S(ChannelID,19)*PointRate<>0 Then
					 If C_S(ChannelID,19)>0 Then
					  Message = Message & "���" & Setting(45) & "��<font color=red>" & C_S(ChannelID,19)*PointRate & "</font> " & Setting(46) & Setting(45) & "<br>"
					 ElseIf C_S(ChannelID,19)<0 Then
					  Message = Message & "����" & Setting(45) & "��<font color=red>" & Abs(C_S(ChannelID,19))*PointRate & "</font> " & Setting(46) & Setting(45) & "<br>"
					 End If
					End If
					 
					If Conn.Execute("Select top 1 * From KS_LogScore Where UserName='" & UserName & "' and ChannelID=" & ChannelID & " and InfoID=" & InfoID).Eof And C_S(ChannelID,20)*ScoreRate<>0 Then    'û�м�¼�Ÿ����ӻ���
						 If C_S(ChannelID,20)>0 Then
						  Message = Message & "��û��֣�<font color=red>" & C_S(ChannelID,20)*ScoreRate & "</font> �ֻ���<br>"
						 ElseIf C_S(ChannelID,20)<0 Then
						  Message = Message & "���Ļ��֣�<font color=red>" & Abs(C_S(ChannelID,20))*ScoreRate & "</font> �ֻ���<br>"
						 End If
					End If
					
					Message = Message & "<br />��ע������Ϣ��ϵͳ�Զ��������벻Ҫ�ظ�������"
					If C_S(ChannelID,19)<0 Then  
					Call PointInOrOut(ChannelID,InfoID,UserName,2,-C_S(ChannelID,19)*PointRate,"ϵͳ","����" & C_S(ChannelID,3) & "[" & InfoTitle & "]����",1)            
					Else
					Call PointInOrOut(ChannelID,InfoID,UserName,1,C_S(ChannelID,19)*PointRate,"ϵͳ","����" & C_S(ChannelID,3) & "[" & InfoTitle & "]����",1)            
					End If
					
					If C_S(ChannelID,20)<0 Then
					 Call ScoreInOrOut(UserName,2,-C_S(ChannelID,20)*ScoreRate,"ϵͳ","����" & C_S(ChannelID,3) & "[" & InfoTitle & "]����",ChannelID,InfoID)            
					Else
					 Call ScoreInOrOut(UserName,1,C_S(ChannelID,20)*ScoreRate,"ϵͳ","����" & C_S(ChannelID,3) & "[" & InfoTitle & "]����",ChannelID,InfoID)            
					End If
					If C_S(ChannelID,18)<0 Then
					Call MoneyInOrOut(UserName,ClientName,-C_S(ChannelID,18)*MoneyRate,4,2,SqlNowString,"0","ϵͳ","����" & C_S(ChannelID,3) & "[" & InfoTitle & "]����",ChannelID,InfoID)
					Else
					Call MoneyInOrOut(UserName,ClientName,C_S(ChannelID,18)*MoneyRate,4,1,SqlNowString,"0","ϵͳ","����" & C_S(ChannelID,3) & "[" & InfoTitle & "]����",ChannelID,InfoID)
					End If
					If ChkClng(U_S(GroupID,10))=1 Then Call SendInfo(UserName,Sender,Title,Message)
		End IF
		RSObj.Close:Set RSObj=Nothing
	End Sub
	'����:��Ա��ȯ��ϸ���뺯��	                                                       '����:Channelid-ģ��ID,InfoID-��ϢID��UserName-�û���,InOrOutFlag-��������1����2֧��,Point-���׵���,User-����Ա,Descript-������ע
	Public Function PointInOrOut(ChannelID,InfoID,UserName,InOrOutFlag,Point,User,Descript,ContributeFlag)
	  If Not IsNumeric(InOrOutFlag) Or Not IsNumeric(Point) Or Point=0 Then PointInOrOut=false:Exit Function
	  Dim PointParam,CurrPoint
	  If InOrOutFlag=1 Then 
	     PointParam="Set Point=Point+" & Point
	  ElseIF InOrOutFlag=2 Then
	     PointParam="Set Point=Point-" & Point
	  Else
	    PointInOrOut=false:Exit Function
	  End If
	  If (Conn.Execute("Select top 1 * From KS_LogPoint Where UserName='" & UserName & "' and ChannelID=" & ChannelID & " and InfoID=" & InfoID & " And InOrOutFlag=" & InOrOutFlag).Eof) Or (ChannelID=0 And InfoID=0) or ContributeFlag=0 Then
		  on error resume next
		  Conn.Execute("Update KS_User " & PointParam & " Where UserName='" & UserName & "'")
		  CurrPoint=Conn.Execute("Select top 1 Point From KS_User Where UserName='" & UserName & "'")(0)
		  Conn.Execute("Insert into KS_LogPoint(ChannelID,InfoID,UserName,InOrOutFlag,Point,Times,[User],Descript,Adddate,IP,CurrPoint,ContributeFlag) values(" & ChannelID & "," & InfoID & ",'" & UserName & "',"& InOrOutFlag & "," & Point & ",1,'" & replace(User,"'","""") & "','" & replace(Descript,"'","""") & "'," & SqlNowString & ",'" & replace(getip,"'","""") & "'," & CurrPoint & "," & ContributeFlag & ")")
	  End If
	  IF Err Then PointInOrOut=false Else PointInOrOut=true
	End Function
	
	'����:��Ա������ϸ���뺯��	
	'����:UserName-�û���,InOrOutFlag-��������1����2֧��,Score-���׵���,User-����Ա,Descript-������ע
	Public Function ScoreInOrOut(UserName,InOrOutFlag,Score,User,Descript,ChannelID,InfoID)
	  If Not IsNumeric(InOrOutFlag) Or Not IsNumeric(Score) Or Score=0 Then ScoreInOrOut=false:Exit Function
	  Dim ScoreParam,CurrScore
	  If InOrOutFlag=1 Then 
	     ScoreParam="Set Score=Score+" & Score
		 '�ж���û�е���ÿ�����ӵ�����
		 If ChkClng(Setting(165))<>0 Then
		  Dim TodayScore:TodayScore=ChkClng(Conn.Execute("select sum(Score) from ks_logscore where InOrOutFlag=1 and year(adddate)=year(" & SQLNowString & ") and month(adddate)=month(" & SQLNowString & ") and day(adddate)=day(" & SQLNowString & ") and username='" & UserName & "'")(0))
		  If TodayScore>=ChkClng(Setting(165)) Then Exit Function
		 End If
	  ElseIF InOrOutFlag=2 Then
	     ScoreParam="Set Score=Score-" & Score
	  Else
	    ScoreInOrOut=false:Exit Function
	  End If
	  If (Conn.Execute("Select top 1 * From KS_LogScore Where UserName='" & UserName & "' and ChannelID=" & ChannelID & " and InfoID=" & InfoID & " And InOrOutFlag=" & InOrOutFlag).Eof) Or (ChannelID=0 And InfoID=0) Then
		  on error resume next
		  Conn.Execute("Update KS_User " & ScoreParam & " Where UserName='" & UserName & "'")
		  CurrScore=Conn.Execute("Select top 1 Score From KS_User Where UserName='" & UserName & "'")(0)
		  Conn.Execute("Insert into KS_LogScore(UserName,InOrOutFlag,Score,CurrScore,[User],Descript,Adddate,IP,ChannelID,InfoID) values('" & UserName & "',"& InOrOutFlag & "," & Score & ","&CurrScore & ",'" & replace(User,"'","""") & "','" & replace(Descript,"'","""") & "'," & SqlNowString & ",'" & replace(getip,"'","""") & "'," & ChannelID &"," & InfoID &")")
	  End If
	  IF Err Then ScoreInOrOut=false Else ScoreInOrOut=true
	End Function
	
	'����:�ʽ���ϸ���뺯��	                 
	'����:UserName-�û���,ClientName-�ͻ�����,Money-��Ǯ,MoneyType-����,InOrOutFlag-��������1����2֧��,PayTime-�������,OrderID-������,Inputer-����Ա,Remark-������ע
	Public Function MoneyInOrOut(UserName,ClientName,Money,MoneyType,InorOutFlag,PayTime,OrderID,Inputer,Remark,ChannelID,InfoID)
	  If Not IsNumeric(InOrOutFlag) Or Not IsNumeric(Money) Or Money="0" Then MoneyInOrOut=false:Exit Function
	  Dim MoneyParam,CurrMoney
	  If InOrOutFlag=1 Then 
	     MoneyParam="Set [Money]=[Money]+" & Money
	  ElseIF InOrOutFlag=2 Then
	     MoneyParam="Set [Money]=[Money]-" & Money
	  Else
	    MoneyInOrOut=false:Exit Function
	  End If
	  If (Conn.Execute("Select top 1 * From KS_LogMoney Where UserName='" & UserName & "' and ChannelID=" & ChannelID & " and InfoID=" & InfoID & " And IncomeOrPayOut=" & InOrOutFlag).Eof) Or (ChannelID=0 And InfoID=0) Then
		  on error resume next
		  Conn.Execute("Update KS_User " & MoneyParam & " Where UserName='" & UserName & "'")
		  CurrMoney=Conn.Execute("Select top 1 Money From KS_User Where UserName='" & UserName & "'")(0)
	      Conn.Execute("Insert into KS_LogMoney([UserName],[ClientName],[Money],[MoneyType],[IncomeOrPayOut],[OrderID],[Remark],[PayTime],[LogTime],[Inputer],[IP],[CurrMoney],[ChannelID],[InfoID]) values('" & UserName & "','" & ClientName & "'," & Money & "," & MoneyType & ","& InOrOutFlag & ",'" & OrderID & "','" & replace(Remark,"'","""") & "'," & SqlNowString & "," &SqlNowString & ",'" & replace(inputer,"'","""") & "','" & replace(getip,"'","""") & "'," & CurrMoney & "," & ChannelID & "," & InfoID & ")")
	  End If
	  IF Err Then MoneyInOrOut=false Else MoneyInOrOut=true
	End Function
	'��Ա��Ч����ϸ���뺯��
	'����:UserName,InOrOutFlag,Edays,User,Descript
	Function EdaysInOrOut(UserName,InOrOutFlag,Edays,User,Descript)
		 If Not IsNumeric(InOrOutFlag) Or Not IsNumeric(Edays) Or Edays=0 Then EdaysInOrOut=false:Exit Function
		  Conn.Execute("insert into KS_LogEdays(UserName,InOrOutFlag,Edays,[user],descript,adddate,ip) values('" & UserName & "'," & InOrOutFlag & "," & Edays & ",'" & user & "','" & replace(descript,"'","""") & "'," & SqlNowString & ",'" & getip & "')")
		  IF Err Then EdaysInOrOut=false Else EdaysInOrOut=true
	 End Function
	'����վ����Ϣ
	'����Incept--������,Sender-������,title--����,Content--�ż�����
	Public Sub SendInfo(Incept,Sender,title,Content)
	  Conn.Execute("insert Into KS_Message(Incept,Sender,Title,Content,SendTime,Flag,IsSend,DelR,DelS) values('" & Incept & "','" & Sender & "','" & replace(Title,"'","""") & "','" & replace(Content,"'","""") & "'," & SqlNowString & ",0,1,0,0)")
	End Sub
	'���˷Ƿ��ַ�
	Public Function FilterIllegalChar(ByVal Content)
	   If IsNul(Content) Then Exit Function
	   Dim SplitStrArr,K:SplitStrArr=split(Setting(55),vbCrlf)
	   For K=0 To Ubound(SplitStrArr)
		  If Not IsNul(SplitStrArr(K)) Then
		   Content=Replace(Content,Split(SplitStrArr(K),"=")(0),Split(SplitStrArr(K),"=")(1))
		  End If
		Next
        FilterIllegalChar=Content 
	End Function
	
  '======================================================================================
End Class

%> 