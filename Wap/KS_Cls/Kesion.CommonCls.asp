<!--#include file="Kesion.MemberCls.asp"-->
<!--#include file="Kesion.TranPinyinCls.asp"-->
<!--#include file="Kesion.Thumbs.asp"-->
<%
Const ClassField="ID,FolderName,Folder,ClassPurview,FolderDomain,TemplateID,ClassBasicInfo,ClassDefineContent,TS,ClassID,Tj,DefaultDividePercent,ChannelID,TN,ClassType,FolderOrder,AdminPurview,AllowArrGroupID,CommentTF"           '定义载入缓存的栏目字段

Class PublicCls
	  Private LocalCacheName,Cache_Data,CacheData
	  Public SiteSN,Version,BusinessVersion
	  Public Setting,TbSetting,SSetting,ASetting,JSetting,WSetting
	  Private Sub Class_Initialize()
		If Not Response.IsClientConnected Then Response.End()
		Call Initialize_Kesion_Config
     End Sub
	 Function InitialObject(str)
		'iis5创建对象方法Server.CreateObject(ObjectName);
		'iis6创建对象方法CreateObject(ObjectName);
		'默认为iis6，如果在iis5中使用，需要改为Server.CreateObject(str);
		Set InitialObject=CreateObject(str)
	 End Function

	 Private Sub Class_Terminate()
	 End Sub
	 '*******************************************************************************************************************
	 '函数名：KSInitialize
	 '作  用: 加载KesionCMS的必要参数
	 '备  注：以下参数请不要更改。否则系统可能无法正常运行
	 '*******************************************************************************************************************
	 Public Function Initialize_Kesion_Config()
		 Call InitialConfig()
		 SiteSN="KS6" & Replace(Replace(LCase(Request.ServerVariables("SERVER_NAME")), "/", ""), ".", "")
		 Version = "V6"
		 BusinessVersion = 0
		 If WSetting(0)<>1 Then
		    '是否关闭WAP功能
		    Dim TemplateContent
			TemplateContent = C_T("20087352214569",2)
			Response.Write TemplateContent
			Response.End
	     End If
		 Call MoniqiKaiguan()'是否开启Wap模拟器访问
		 Call GetSiteOnline()'是否启用站点计数器
		 Call GetPromotion()'推广积分
		 If G("U")<>"" Then
		    Response.Redirect GetDomain&"Space/?i="&G("U")&"&"&WapValue&""'个人空间转向
	     End If
	 End Function
	 
	'不提示,批量清除缓存,参数 PreCacheName-前段匹配
	Public Sub DelCaches(PreCacheName)
	    Dim i
		Dim CacheList:CacheList=split(GetCacheList(PreCacheName),",")
		If UBound(CacheList)>1 Then
			For i=0 to UBound(CacheList)-1
				DelCahe CacheList(i)
			Next
		End IF
	End Sub
	'取得缓存列表 参数 PreCacheName-前段匹配
	Public Function GetCacheList(PreCacheName)
		Dim Cacheobj
		For Each Cacheobj in Application.Contents
		If CStr(Left(Cacheobj,Len(PreCacheName)))=CStr(PreCacheName) Then GetCacheList=GetCacheList&Cacheobj&","
		Next
	End Function
	'清除缓存,参数 MyCaheName-缓存名称
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
	
	 'xmlroot跟节点名称 row记录行节点名称
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
	 
	 
	 	 'xmlroot跟节点名称 row记录行节点名称
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
		 Dim RS:Set Rs=Conn.Execute("select ChannelID,ChannelName,ChannelTable,ItemName,ItemUnit,FieldBit,BasicType,FsoHtmlTF,FsoFolder,RefreshFlag,ModelEname,MaxPerPage,VerificCommentTF,CommentVF,CommentLen,CommentTemplate,UserSelectFilesTF,InfoVerificTF,UserAddMoney,UserAddPoint,UserAddScore,ChannelStatus,CollectTF,UpFilesTF,UpFilesDir,UpFilesSize,UserUpFilesTF,UserUpFilesDir,AllowUpPhotoType,AllowUpFlashType,AllowUpMediaType,AllowUpRealType,AllowUpOtherType,SearchTemplate,EditorType,FsoListNum,UserTF,DiggByVisitor,DiggByIP,DiggRepeat,DiggPerTimes,UserClassStyle,UserEditTF,FsoContentRule,FsoClassListRule,FsoClassPreTag,ThumbnailsConfig,LatestNewDay,StaticTF,WapSwitch,WapSearchTemplate From KS_Channel Order by ChannelID")
		 Set Application(SiteSN&"_ChannelConfig")=RecordsetToxml(rs,"channel","ChannelConfig")
		 Set Rs=Nothing
		 Application.unLock
	 End Function
	 
	 Function C_S(sChannelID,FieldID)
	     On Error Resume Next
		 If not IsObject(Application(SiteSN&"_ChannelConfig")) Then LoadChannelConfig()
		 C_S=Application(SiteSN&"_ChannelConfig").documentElement.selectSingleNode("channel[@ks0=" & sChannelID & "]/@ks" & FieldID & "").text
		 If err Then C_S=0:err.Clear
	 End Function
	 
	 Public Function LoadClassConfig()
	     Application.Lock
		 Dim RS:Set Rs=conn.execute("select " & ClassField & " From KS_Class Order by root,folderorder")
		 Set Application(SiteSN&"_class")=RecordsetToxml(rs,"class","classConfig")
		 Set Rs=Nothing
		 Application.unLock
	 End Function

	 Function C_C(ClassID,FieldID)
	     On Error Resume Next
		 If not IsObject(Application(SiteSN&"_class")) Then LoadClassConfig()
		 C_C=Application(SiteSN&"_class").documentElement.selectSingleNode("class[@ks0=" & classID & "]/@ks" & FieldID & "").text
	 End Function
	 
	 '加载用户组缓存
	 Sub LoadUserGroup()
	   If Not IsObject(Application(SiteSN&"_UserGroup")) Then 
	    Application.Lock
	     Dim RS:Set Rs=conn.execute("select id,groupname,powerlist,descript,usertype,formid,templatefile,showonreg From KS_UserGroup Order by ID")
		 Set Application(SiteSN&"_UserGroup")=RsToxml(rs,"row","groupConfig")
         Set Rs=Nothing
	     Application.unLock
	   End If
	 End Sub
	 
	 Function U_G(GroupID,FieldName)
	   LoadUserGroup
	   Dim Node:Set Node=Application(SiteSN&"_UserGroup").DocumentElement.selectSingleNode("row[@id=" & GroupID & "]/@" & Lcase(FieldName))
	   If Not Node Is Nothing Then U_G=Node.text
	   Set Node=Nothing
	 End Function
	 
	 '加载留言版面缓存
	 Sub LoadClubBoard()
	   If Not IsObject(Application(SiteSN&"_ClubBoard")) Then 
	    Application.Lock
	     Dim RS:Set Rs=conn.execute("select [id],[boardname],[note],[master],[todaynum],[postnum],[topicnum],[parentid],[LastPost],[BoardRules],[Settings] From KS_GuestBoard Order by orderid,ID")
		 Set Application(SiteSN&"_ClubBoard")=RsToxml(rs,"row","clubConfig")
         Set Rs=Nothing
	     Application.unLock
	   End If
	 End Sub
	 
	 

	'缓存自定义页面内容
	Function C_T(LabelID,FieldID)
	    On Error Resume Next
		If not IsObject(Application(SiteSN&"_waplabellist")) Then
		   Application.Lock
		   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		   RS.Open "select ID,TemplateName,TemplateContent from KS_WapTemplate Order by AddDate",Conn,1,1
		   Set Application(SiteSN&"_waplabellist")=RecordsetToxml(RS,"waplabel","waplabellist")
		   RS.Close:Set RS=Nothing
		   Application.UnLock
		End If
		C_T=Application(SiteSN&"_waplabellist").documentElement.selectSingleNode("waplabel[@ks0='" & LabelID & "']/@ks" & FieldID & "").text
		If Err Then C_T="":Err.Clear
	End Function


     '**************************************************
	 '作  用：推广积分
	 '**************************************************
	 Sub GetPromotion()
	     If Setting(140)="1" Then
		    Dim UID:UID=S("UID")
			If UID<>"" Then
		    If Not Conn.Execute("Select Top 1 UserName From KS_User Where UserName='" & UID & "'").EOF Then
			   Dim UserIP,ComeUrl,RS,SQL
			   UserIP=GetIP()
			   ComeUrl=Request.ServerVariables("HTTP_REFERER")
			   If ComeUrl="" Then ComeUrl="★直接手机输入或书签导入★"
			   Set RS=Server.CreateObject("ADODB.RECORDSET")
			   If DataBaseType=1 Then
			      SQL="Select top 1 * from KS_PromotedPlan Where UserName='" & UID & "' And UserIP='" & UserIP & "' And DateDiff(day,AddDate," & SqlNowString & ")<1"
			   Else 
			      SQL="Select top 1 * from KS_PromotedPlan Where UserName='" & UID & "' And UserIP='" & UserIP & "' And DateDiff('d',AddDate," & SqlNowString & ")<1"
			   End If
			   RS.Open SQL ,Conn,1,3
			   If RS.Eof And RS.Bof Then
				  RS.AddNew
				  RS("UserName") = UID
				  RS("UserIP")   = UserIP
				  RS("AddDate")  = Now
				  RS("ComeUrl")  = URLDecode(ComeUrl)
				  RS("Score")    = Setting(141)
				  RS("AllianceUser")="-"
				  RS.Update
				  RS.Close
				  Conn.Execute("Update KS_User Set Score=Score+" & Setting(141) & " where UserName='" & UID & "'")
			   Else 
			      RS.Close
			   End IF
			   Set RS=Nothing
			End If
			End If
		 End If
	 End Sub
	   
	'**************************************************
	'作  用：是否开启Wap模拟器访问,1开启,0关闭
	'**************************************************
	Function MoniqiKaiguan()
	    If WSetting(1)=0 Then
		   Dim StrFilter,arrfilter,j
	       StrFilter="oper|winw|wapi|mc21|up.b|upg1|upsi|qwap|jigs|java|alca|wapj|cdr/|nec-|fetc|r380|mozi|m3ga"
		   arrfilter = Split(StrFilter,"|")
		   For j = 0 to Ubound(Arrfilter)
		       If Instr(Lcase(Request.ServerVariables("HTTP_USER_AGENT")),arrfilter(j))>0 Then
			      Dim TemplateContent
				  TemplateContent = C_T("20088977113042",2)
				  Response.Write TemplateContent
				  Response.End
			   End if
		   Next
		End If  
	End Function

	'**************************************************
	'函数名：
	'作  用：显示错误信息。
	'参  数：Errmsg  ----出错信息
	'返回值：无
	'**************************************************
	Public Sub ShowError(strTitle,strErr)
	     response.redirect getdomain & "plus/error.asp?message=" & server.URLEncode(strErr)
	    'Dim FileContent,KSRFObj
		'Set KSRFObj = New Refresh
	    'FileContent = C_T("20083987211486",2)
		'FileContent = Replace(FileContent,"{$GetTitle}",strTitle)
		'FileContent = Replace(FileContent,"{$GetContent}",strErr)
		'FileContent = KSRFObj.KSLabelReplaceAll(FileContent)
		'Response.Write FileContent
		'Set KSRFObj=Nothing
		'Response.End
    End Sub

	'================================================
	'过程名：PreventRefresh
	'作 用：防止刷新页面
	'================================================
	Public Function PreventRefresh()
	    Dim RefreshTime
		RefreshTime=5 '防止刷新时间‚单位（秒）
		If NOT IsEmpty(Session("RefreshTime")) And Isnumeric(Session("RefreshTime")) Then
		   If (Timer()-Int(Session("RefreshTime")))*1000 < RefreshTime*1000 Then
		      PreventRefresh=True
			  Session("RefreshTime")=Timer()
		   Else
		      PreventRefresh=False
			  Session("RefreshTime")=Timer()
		   End If
		Else
		   PreventRefresh=False
		   Session("RefreshTime")=Timer()
		End If
	End Function

	Public Function WapValue()
	    On Error Resume Next
		Dim Wap,ShareID
		Wap=WSetting(2)
		WapValue = ""&Wap&"="&Request.QueryString(Wap)&""
		ShareID = ChkClng(S("ShareID"))
		If ShareID <> 0 Then
	       WapValue = WapValue & "&amp;ShareID=" & ShareID
	    End If
	End Function
	
	
	
	'************************************************************************
	'函数名: GetClassNP
	'功 能: 取得目录名称并加上链接
	'参 数: ClassID目录的ID	          
	'*************************************************************************
	Function GetClassNP(ClassID)
		On Error Resume Next
		If Not IsObject(Application(SiteSN&"_classnamepath")) Then
		   Dim Folder,ClassPurview,ChannelFsoHtmlTF,Node,K,SQL,RS
		   Set  Application(SiteSN&"_classnamepath")=Server.CreateObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		   Application(SiteSN&"_classnamepath").appendChild( Application(SiteSN&"_classnamepath").createElement("xml"))
		   Set RS=Server.CreateObject("ADODB.RECORDSET")
		   RS.Open "Select ID,FolderName,ClassID,ChannelID From KS_Class Order BY FolderOrder", Conn, 1, 1
		   If RS.Eof And RS.Bof Then RS.Close:Set RS=Nothing:Exit Function
		   SQL=RS.GetRows(-1):RS.Close:Set RS=Nothing
		   For K=0 To Ubound(SQL,2)
		       Set Node=Application(SiteSN&"_classnamepath").documentElement.appendChild(Application(SiteSN&"_classnamepath").createNode(1,"classnamepath",""))
			   Node.attributes.setNamedItem(Application(SiteSN&"_classnamepath").createNode(2,"classid","")).text=SQL(0,K)
			   Node.text="<a href=""" & GetFolderPath(SQL(3,K),SQL(2,K)) & """>" & Trim(SQL(1,K)) & "</a>"
		   Next			
        End If
		GetClassNP=Application(SiteSN&"_classnamepath").documentElement.selectSingleNode("classnamepath[@classid=" & ClassID & "]").text
	End Function
	
	
	'**************************************************
	'函数名：GetGoBackIndex
	'作  用：取出还回首页地址
	'**************************************************
	Public Function GetGoBackIndex()
	    On Error Resume Next
		Dim ShareID:ShareID = ChkClng(G("ShareID"))
		If ShareID = "" Then
	       GetGoBackIndex = "" & GetDomain & "?" & WapValue & ""
	    Else
           Dim RS:Set RS = Conn.Execute("select top 1 HomePage from KS_User Where UserID=" & ShareID & "")
		   If RS.Eof Then
		      GetGoBackIndex = "" & GetDomain & "?" & WapValue & ""
		   Else
		      GetGoBackIndex = RS(0)
		   End iF
		   Set RS = nothing
	    End iF
	End Function
	
	'**************************************************
	'函数名：GetDomain
	'作  用：获取URL,包括虚拟目录 如http://www.kesion.com/ 或 http://www.kesion.com/Sys/wap/  其中 Sys/wap/为虚拟目录
	'参  数：  无
	'返回值：完整域名
	'**************************************************
	Public Function GetDomain()
	  If G_Domain<>"" Then
	   GetDomain = Trim(G_Domain) 
	  Else
	   GetDomain = Trim(Setting(2) & Setting(3) & WSetting(4)) 
	  End If
	End Function	 
	
	'**************************************************
	'函数名:GetFolderTid
	'功 能:取得子目录的ID集合
	'参 数:  FolderID父目录ID
	'返回值: 形如 1255555,111111,4444的ID集合
	'---------------------------------------------------------------------------------------------------------
	Function GetFolderTid(FolderID)
	    GetFolderTid="Select ID From KS_Class Where DelTF=0 AND WapSwitch=1 AND TS LIKE '%" & FolderID & "%'":Exit Function
	End Function
	
	Public Function GetUserRealName(UserName)
	    On Error Resume Next
		Dim UserRS:set UserRS=Conn.Execute("select top 1 RealName from KS_User where UserName='" & UserName & "'")
		If UserRS.EOF Then
		   GetUserRealName=UserName 
		Else
		   If IsNul(UserRS(0)) Then
		      GetUserRealName=UserName 
		   Else
		      GetUserRealName=UserRS(0)
		   End If
		End If
		Set UserRS=Nothing
	End Function
	
	
	'载入文件类自定义字段
	Sub LoadFieldToXml()
	  If Not IsObject(Application(SiteSN & "_FeildXml")) then
			Dim Rs:Set Rs = Conn.Execute("Select ChannelID,FieldName,fieldtype From KS_Field Where FieldType=9 or FieldType=10 Order By FieldID")
			Set Application(SiteSN & "_FeildXml")=RsToxml(Rs,"row","FeildXml")
			Set Rs = Nothing
	  End If
	End Sub
	'添加自关联数据库	
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
	
	'根据内容获取上传文件名
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
	'函数名：ReturnChannelAllowUserUpFilesTF
	'作  用：返回频道是否允许会员上传文件
	'参  数：ChannelID--频道ID
	'**************************************************
	Public Function ReturnChannelAllowUserUpFilesTF(ChannelID)
	    If ChannelID = "" Or Not IsNumeric(ChannelID) Then '默认允许上传文件
		   ReturnChannelAllowUserUpFilesTF = True:Exit Function
		End If
		If C_S(ChannelID,26) = 1 Then
		   ReturnChannelAllowUserUpFilesTF = True
		Else
		   ReturnChannelAllowUserUpFilesTF = False
		End If
	End Function

	'**************************************************
	'函数名：ReturnChannelUserUpFilesDir
	'作  用：返回频道前台会员的上传目录
	'参  数：ChannelID--频道ID,UserFolder-按用户名生成的目录
	'返回值：目录字符串
	'**************************************************
	Public Function ReturnChannelUserUpFilesDir(ChannelID,UserFolder)
	    Dim Ce:Set Ce=new CtoeCls
	    UserFolder=Ce.CTOE(R(UserFolder))
	    Set Ce=Nothing
	    ChannelID = ChkCLng(ChannelID)
	    Select Case ChannelID
	    Case 9999 '用户头像
		   ReturnChannelUserUpFilesDir=Setting(3)&Setting(91)&"User/" & UserFolder &"/upface/"
		Case 9998 '相册封面
		   ReturnChannelUserUpFilesDir=Setting(3)&Setting(91)&"User/" & UserFolder &"/xc/"
		Case 9997 '照片
		   ReturnChannelUserUpFilesDir=Setting(3)&Setting(91)&"User/" & UserFolder &"/xc/"
		Case 9996 '圈子图片
		   ReturnChannelUserUpFilesDir=Setting(3)&Setting(91)&"User/" & UserFolder &"/team/"
		Case 9995 '音乐
		   ReturnChannelUserUpFilesDir=Setting(3)&Setting(91)&"User/" & UserFolder &"/music/"
		Case 9994 '产品
		   ReturnChannelUserUpFilesDir=Setting(3)&Setting(91)&"User/" & UserFolder &"/product/"
		Case 999
		   ReturnChannelUserUpFilesDir=Setting(3)&Setting(91)&"GuestBook/"&UserFolder &"/"
		Case Else
		  ReturnChannelUserUpFilesDir = C_S(ChannelID,27)
		  ReturnChannelUserUpFilesDir = Setting(3) & Setting(91)&"User/" & UserFolder &"/"& ReturnChannelUserUpFilesDir
	    End Select
	End Function

	'**************************************************
	'函数名：ReturnChannelAllowUpFilesSize
	'作  用：返回频道的最大允许上传文件大小
	'参  数：ChannelID--频道ID
	'**************************************************
	Public Function ReturnChannelAllowUpFilesSize(ChannelID)
	    ChannelID = ChkClng(ChannelID)
	    Dim CRS:Set CRS=conn.execute("Select UpFilesSize From KS_Channel Where ChannelID=" & ChannelID)
	    If CInt(ChannelID) = 0 Or (CRS.EOF And CRS.BOF) Then
		   ReturnChannelAllowUpFilesSize = Setting(6)
	    Else
		   ReturnChannelAllowUpFilesSize = CRS(0)
	    End If
		CRS.Close:Set CRS = Nothing
	End Function

	'**************************************************
	'函数名：ReturnChannelAllowUpFilesType
	'作  用：返回频道的允许上传的文件类型
	'参  数：ChannelID--频道ID,TypeFlag 0-取全部 1-图片类型 2-flash 类型 3-Windows 媒体类型 4-Real 类型 5-其它类型
	'**************************************************
	Public Function ReturnChannelAllowUpFilesType(ChannelID, TypeFlag)
	    If ChkClng(ChannelID) = 0 Then  ReturnChannelAllowUpFilesType = Setting(7):Exit Function
	    If Not IsNumeric(TypeFlag) Then TypeFlag = 0
		If TypeFlag = 0 Then   '所有允许的类型
		   ReturnChannelAllowUpFilesType = C_S(ChannelID,28) & "|" & C_S(ChannelID,29) & "|" & C_S(ChannelID,30) & "|" & C_S(ChannelID,31) & "|" & C_S(ChannelID,32)
		Else
		   ReturnChannelAllowUpFilesType = C_S(ChannelID,27+TypeFlag)
		End If
	End Function

	'返回付款方式名称,参数TypeID,0名称 1折扣率
	Function ReturnPayment(ID,TypeID)
	    If Application(SiteSn &"Payment_" & ID&TypeID)="" Then
           Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		   RS.Open "Select TypeName,Discount From KS_PaymentType Where TypeID=" & ID,Conn,1,1
		   If Not RS.Eof Then
			  If TypeID=0 Then
			     ReturnPayment=RS(0)
				 If RS(1)<100 Then ReturnPayment=ReturnPayment & "折扣率:" & RS(1) & "%"
			  Else
			      ReturnPayment=RS(1)
			  End if
		   End iF 
		   Application(SiteSn &"Payment_" & ID&TypeID)=ReturnPayment
		Else
		   ReturnPayment=Application(SiteSn &"Payment_" & ID&TypeID)
		End If
	End Function
	'返回收货方式名称,参数TypeID,0名称 1费用
	Function ReturnDelivery(ID,TypeID)
	    If Application(SiteSn &"Delivery_" & ID&TypeID)="" Then
           Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		   RS.Open "Select TypeName,fee From KS_Delivery Where TypeID=" & ID,Conn,1,1
		   If Not RS.Eof Then
		      If TypeID=0 Then
		  	     ReturnDelivery=RS(0)
				 If RS(1)=0 Then ReturnDelivery=ReturnDelivery & "免费" Else ReturnDelivery=ReturnDelivery & "加收 " & RS(1) & "元"
			  Else
			     ReturnDelivery=RS(1)
			  End iF
		   End iF 
		   Application(SiteSn &"Delivery_" & ID&TypeID)=ReturnDelivery
		Else
	       ReturnDelivery=Application(SiteSn &"Delivery_" & ID&TypeID)
		End If
	End Function
	
	'**************************************************
	'函数：FoundInArr
	'作  用：检查一个数组中所有元素是否包含指定字符串
	'参  数：strArr     ----字符串
	'        strToFind    ----要查找的字符串
	'       strSplit    ----数组的分隔符
	'返回值：True,False
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

	
	'存入返回缓存链接
	Public Function GetWriteinReturn(str)
	    Dim ReturnLin,UserReturnID,ReturnID
	    ReturnLin = 500
	    UserReturnID = Session("UserReturnID")
	    If UserReturnID="" Then
		   Session("UserReturnID") = "OK"
		   If Session("UserReturnID") = "OK" Then
		      ReturnID = Application(SiteSn & "_ReturnID")
		      If ReturnID = "" Then ReturnID = 1
			  If ReturnID > ReturnLin Then ReturnID = 1
			  Application.Lock
			  Application(SiteSn & "_ReturnID") = ReturnID + 1
			  Application(SiteSn & "_Return" & UserReturnID) = str
			  Application.UnLock
			  Session("UserReturnID") = ReturnID + 1
		   Else
		      Session("UserReturnID") = ""
		   End If
		Else
	       Application.Lock
		   Application(SiteSn & "_Return" & UserReturnID) = str
		   Application.UnLock
		End If
	End Function
	
	'读取返回缓存链接
	Public Function GetReadReturn()
	    Dim UserReturnID,TempStr
	    UserReturnID = Session("UserReturnID")
		If UserReturnID="" Then
		   TempStr = ""
		Else
           TempStr = Application(SiteSn & "_Return" & UserReturnID)
		End If
		GetReadReturn = TempStr
	End Function

	 '**************************************************
	 '函数名:DateFormat
	 '功 能:日期格式函数
	 '参 数: DateStr日期, Types转换类型	
	 '**************************************************
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
			   If Types=21 Then
			      DateString = "(" & DateString &")"
			   ElseIf Types=41 Then
			      DateString = "[" & DateString &"]"
			   End If
		   Case 2,22,42
		       DateString=Year(DateStr) & "." & Right("0" & Month(DateStr), 2) & "." & Right("0" & Day(DateStr), 2)
			   If Types=22 Then
			      DateString = "(" & DateString &")"
			   ElseIf Types=42 Then
			      DateString = "[" & DateString &"]"
			   End If
		   Case 3,23,43
		       DateString=Year(DateStr) & "/" & Right("0" & Month(DateStr), 2) & "/" & Right("0" & Day(DateStr), 2)
			   If Types=23 Then
			      DateString = "(" & DateString &")"
			   ElseIf Types=43 Then
			      DateString = "[" & DateString &"]"
			   End If
		   Case 4,24,44
		       DateString=Right("0" & Month(DateStr), 2) & "/" & Right("0" & Day(DateStr), 2) & "/" & Year(DateStr)
			   If Types=24 Then
			      DateString = "(" & DateString &")"
			   ElseIf Types=44 Then
			      DateString = "[" & DateString &"]"
			   End If
		   Case 5,25,45
		       DateString = Year(DateStr) & "年" & Right("0" & Month(DateStr), 2) & "月"
			   If Types=25 Then
			      DateString = "(" & DateString &")"
			   ElseIf Types=45 Then
			      DateString = "[" & DateString &"]"
			   End If
		   Case 6,26,46
		       DateString = Year(DateStr) & "年" & Right("0" & Month(DateStr), 2) & "月" & Right("0" & Day(DateStr), 2) & "日"
			   If Types=26 Then
			      DateString = "(" & DateString &")"
			   ElseIf Types=46 Then
			      DateString = "[" & DateString &"]"
			   End If
		   Case 7,27,47
		       DateString = Right("0" & Month(DateStr), 2) & "." & Right("0" & Day(DateStr), 2) & "." & Year(DateStr)
			   If Types=27 Then
			      DateString = "(" & DateString &")"
			   ElseIf Types=47 Then
			      DateString = "[" & DateString &"]"
			   End If
		   Case 8,28,48
		       DateString = Right("0" & Month(DateStr), 2) & "-" & Right("0" & Day(DateStr), 2) & "-" & Year(DateStr)
			   If Types=28 Then
			      DateString = "(" & DateString &")"
			   ElseIf Types=48 Then
			      DateString = "[" & DateString &"]"
			   End If
		   Case 9,29,49
		       DateString = Right("0" & Month(DateStr), 2) & "/" & Right("0" & Day(DateStr), 2)
			   If Types=29 Then
			      DateString = "(" & DateString &")"
			   ElseIf Types=49 Then
			      DateString = "[" & DateString &"]"
			   End If
		   Case 10,30,50
		       DateString = Right("0" & Month(DateStr), 2) & "." & Right("0" & Day(DateStr), 2)
			   If Types=30 Then
			      DateString = "(" & DateString &")"
			   ElseIf Types=50 Then
			      DateString = "[" & DateString &"]"
			   End If
		   Case 11,31,51
		       DateString = Right("0" & Month(DateStr), 2) & "月" & Right("0" & Day(DateStr), 2) & "日"
			   If Types=31 Then
			      DateString = "(" & DateString &")"
			   ElseIf Types=51 Then
			      DateString = "[" & DateString &"]"
			   End If
		   Case 12,32,52
		       DateString = Right("0" & Day(DateStr), 2) & "日" & Right("0" & Hour(DateStr), 2) & "时"
			   If Types=32 Then
			      DateString = "(" & DateString &")"
			   ElseIf Types=52 Then
			      DateString = "[" & DateString &"]"
			   End If
		   Case 13,33,53
		       DateString = Right("0" & Day(DateStr), 2) & "日" & Right("0" & Hour(DateStr), 2) & "点"
			   If Types=33 Then
			      DateString = "(" & DateString &")"
			   ElseIf Types=53 Then
			      DateString = "[" & DateString &"]"
			   End If
		   Case 14,34,54
		       DateString = Right("0" & Hour(DateStr), 2) & "时" & Minute(DateStr) & "分"
			   If Types=34 Then
			      DateString = "(" & DateString &")"
			   ElseIf Types=54 Then
			      DateString = "[" & DateString &"]"
			   End If
		   Case 15,35,55
		       DateString = Right("0" & Hour(DateStr), 2) & ":" & Right("0" & Minute(DateStr), 2)
			   If Types=35 Then
			      DateString = "(" & DateString &")"
			   ElseIf Types=55 Then
			      DateString = "[" & DateString &"]"
			   End If
		   Case 16,36,56
		       DateString = Right("0" & Month(DateStr), 2) & "-" & Right("0" & Day(DateStr), 2)
			   If Types=36 Then
			      DateString = "(" & DateString &")"
			   ElseIf Types=56 Then
			      DateString = "[" & DateString &"]"
			   End If
		   Case 17,37,57
		       DateString = Right("0" & Month(DateStr), 2) & "/" & Right("0" & Day(DateStr), 2) &" " &Right("0" & Hour(DateStr), 2)&":"&Right("0" & Minute(DateStr), 2)
			   If Types=37 Then
			   DateString = "(" & DateString &")"
			   ElseIf Types=57 Then
			   DateString = "[" & DateString &"]"
			   End If
		   Case Else
		       DateString = DateStr
	    End Select
		DateFormat = DateString
    End Function 
	
	'----------------------------------------------------------------------------------------------------------------------------
	'函数名:GetOrigin
	'功 能:取得文章来源并附加上链接
	'参 数: OriginName名称
	'返回值: 形如 <a href="http://www.xinhua.com" target="_blank">新华网</a>
	'----------------------------------------------------------------------------------------------------------------------------
	Function GetOrigin(OriginName)
		 Dim RS: Set RS=Server.CreateObject("ADODB.Recordset")
		 RS.Open "select OriginName,HomePage From KS_Origin Where OriginName='" & Trim(OriginName) & "'", Conn, 1, 1
		 If RS.EOF Then
		    GetOrigin = OriginName
		 Else
		    If RS("HomePage") <> "" And UCase(Trim(RS("HomePage"))) <> "HTTP://" Then
		       GetOrigin = "<a href=""" & Trim(RS("HomePage")) & """>" & OriginName & "</a>"
		    Else
			   GetOrigin = OriginName
		    End If
		 End If
		 RS.Close:Set RS = Nothing
	 End Function
	'**************************************************
	'函数名：GetReadMessage
	'作  用：提取未读短消息
	'**************************************************
	Public Function GetReadMessage()
	    If Cbool(KSUser.UserLoginChecked)=True Then
	       On Error Resume Next
		   Dim RS:set RS=Server.CreateObject("ADODB.Recordset")
		   RS.Open "select ID from KS_Message where Incept='" & KSUser.UserName & "' And Flag=0",Conn,1,1
		   If Not(RS.BOF And RS.EOF) Then
		      GetReadMessage="<a href="""&GetDomain&"User/User_Message.asp?Action=read&amp;ID="&RS("ID")&"&amp;"&WapValue&""">你有(" & RS.RecordCount & ")未读短消息</a><img src="""&GetDomain&"Images/new_message.gif"" alt=""""/><br/>"
		   End If
		   RS.Close:set RS=nothing
		End if
    End Function
	
	Public Function ReplaceFace(c)
	   Dim k
	   For k=0 To 19
	       c=Replace(c,"[e"&k &"]","<img src=""" & Setting(3) & "images/emot/" & k & ".gif"" alt=""""/>")
	   Next
	   ReplaceFace=C
	End Function

	
	'*********************************************************************************************************
	'函数名：GetSingleFieldValue
	'作  用：取单字段值
	'参  数：SQLStr SQL语句
	'*********************************************************************************************************
	Public Function GetSingleFieldValue(SQLStr)
	    If DataBaseType=0 then
		   On Error Resume Next
		   GetSingleFieldValue=Conn.Execute(SQLStr)(0)
		   If Err Then GetSingleFieldValue=""
		Else
		   Dim RS:Set RS=Conn.Execute(SQLStr)
		   If Not RS.EOF Then
			  GetSingleFieldValue=RS(0)
		   Else
			  GetSingleFieldValue=""
		   End If
		   RS.Close:Set RS=Nothing
		End If
	End Function


	'加密
	Function Encrypt(ecode)
	    dim texts,i
		For i=1 To len(ecode)
		    texts=texts & chr(asc(mid(ecode,i,1))+3)
		next
		Encrypt = texts
	End Function
	'解密
	Function Decrypt(dcode)
	    dim texts,i
		For i=1 To len(dcode)
		    texts=texts & chr(asc(mid(dcode,i,1))-3)
		next
		Decrypt=texts
	End Function
	'匹配 img src,结果以|隔开 
	Function GetImgSrcArr(strng) 
	    If strng="" Or IsNull(strng) Then GetImgSrcArr="":Exit Function
		Dim regEx,Match,Matches,values
		Set regEx = New RegExp
		regEx.Pattern = "src\=.+?\.(gif|jpg)"
		regEx.IgnoreCase = true 
		regEx.Global = True 
		Set Matches = regEx.Execute(strng)
		For Each Match in Matches
		    If instr(lcase(Match.Value),"fileicon")=0 Then
			   values=values&Match.Value&"|" 
			End If
		Next 
		GetImgSrcArr = Replace(Replace(Replace(Replace(values,"'",""),"""",""),"src=",""),Setting(2),"")
		If GetImgSrcArr<>"" Then GetImgSrcArr = left(GetImgSrcArr,len(GetImgSrcArr)-1)
	End Function

	'**************************************************
	'函数名：GetIP
	'作  用：取得正确的IP
	'返回值：IP字符串
	'**************************************************
	Public Function GetIP() 
		Dim strIPAddr 
		strIPAddr = Request.ServerVariables("REMOTE_ADDR") 
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
	'函数名：URLDecode
	'作  用：URL解码
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
	'函数名：R
	'作  用：过滤非法的SQL字符
	'参  数：strChar-----要过滤的字符
	'返回值：过滤后的字符
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
	
	'**************************************************
	'作  用：生成4位数验证码
	'**************************************************
	Function GetVerifyCode()
        Randomize
		Dim cAmount,cCode
		cAmount = 10
		cCode = "0123456789"
		Dim i,vCode(4), vCodes
		For i = 0 To 3
	        vCode(i) = Int(Rnd * cAmount)
			vCodes = vCodes & Mid(cCode, vCode(i) + 1, 1)
		Next
		Session("Verifycode") = vCodes
		GetVerifyCode = vCodes
	End Function

    '**************************************************
	'作  用：取得Request.Querystring 或 Request.Form 的值
	'**************************************************
	Public Function G(Str)
	   G = Replace(Replace(Request(Str), "'", ""), """", "")
	End Function
	Function DelSql(Str)
		Dim SplitSqlStr,SplitSqlArr,I
		SplitSqlStr="dbcc|alter|drop|*|and|exec|or |insert|select|delete|update|count |master|truncate|declare|char|mid|chr|set |where|xp_cmdshell"
		SplitSqlArr = Split(SplitSqlStr,"|")
		For I=LBound(SplitSqlArr) To Ubound(SplitSqlArr)
			If Instr(LCase(Str),SplitSqlArr(I))>0 Then
			    Response.Redirect GetDomain & "plus/error.asp?Action=DelSql&Message=" & SplitSqlArr(I) &""
				Response.End
			End if
		Next
		DelSql = Str
    End Function
	'**************************************************
	'作  用：取得Request.Querystring 或 Request.Form 的值
	'**************************************************
	Public Function S(str)
	    S = DelSql(Replace(Replace(Request(str), "'", ""), """", ""))
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
	'读Cookies值
	Public Function C(Str)
	 C=DelSql(Request.Cookies(SiteSN)(Str))
	End Function
	
	
	'**************************************************
	'函数名：QueryParam
	'作  用：取得QueryString,或Form参数集合
	'参  数:NoCollect表示不收集的字段,多个用英文逗号隔开
	'**************************************************
	Function QueryParam(NoCollect)
		Dim Param,R
		For Each r In Request.QueryString
		    If FoundInArr(NoCollect,R,",")=false Then
			   If Request.QueryString(r)<>"" Then
			      If Param="" Then
				     Param=r & "=" & Server.UrlEncode(Trim(Request.QueryString(r)))
				  Else
				     Param=Param & "&amp;" & r & "=" & Server.UrlEncode(Trim(Request.QueryString(r)))
				  End If
			   End If
		    End If
		Next
		'If Param<>"" Then QueryParam=Param:Exit Function
		For Each r In Request.Form
		    If FoundInArr(NoCollect,R,",")=false Then
			   If Request.Form(r)<>"" Then
			      If Param="" Then
				     Param=r & "=" & Server.UrlEncode(Trim(Request.Form(r)))
				  Else
				     Param=Param & "&amp;" & r & "=" & Server.UrlEncode(Trim(Request.Form(r)))
				  End If
			   End If
			End If
		Next
		QueryParam=Param
	End Function

	'**************************************************
	'函数名：GetUrl
	'作  用：取得当前地址
	'返回值：如(/wap/index.asp?wap=wqc3qev9nmWFnCSkNjRW&)
	'**************************************************		
    Public Function GetUrl()
	    Dim ScriptAddress, M_itemUrl, Page, M_item 
		ScriptAddress = LCase(CStr(Request.ServerVariables("SCRIPT_NAME")))
		M_itemUrl = "" 
		If (Request.QueryString <> "") Then 
		   ScriptAddress = ScriptAddress & "?" 
		   For Each M_item In Request.QueryString 
		       If InStr(Page,M_item)=0 Then 
			      M_itemUrl = M_itemUrl & M_item &"="& Request.QueryString(""&M_item&"") & "&amp;" 
			   End If 
		   Next 
		End if 
		GetUrl = ScriptAddress & M_itemUrl
		'GetItemUrl = left(ItemUrl,len(ItemUrl)-5) 
	End Function 

	Public Function CutFixContent(ByVal str, ByVal start, ByVal last, ByVal n)
		Dim strTemp
		On Error Resume Next
		If InStr(str, start) > 0 Then
			Select Case n
			Case 0  '左右都截取（都取前面）（去处关键字）
				strTemp = Right(str, Len(str) - InStr(str, start) - Len(start) + 1)
				strTemp = Left(strTemp, InStr(strTemp, last) - 1)
			Case Else  '左右都截取（都取前面）（保留关键字）
				strTemp = Right(str, Len(str) - InStr(str, start) + 1)
				strTemp = Left(strTemp, InStr(strTemp, last) + Len(last) - 1)
			End Select
		Else
			strTemp = ""
		End If
		CutFixContent = strTemp
	End Function
	
	'*************************************************************************
	'函数名：gotTopic
	'作  用：截字符串，汉字一个算两个字符，英文算一个字符
	'参  数：str   ----原字符串
	'       strlen ----截取长度
	'返回值：截取后的字符串
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
			c = Abs(Asc(Mid(Str, I, 1)))
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
	'函数名：strLength
	'作  用：求字符串长度。汉字算两个字符，英文算一个字符。
	'参  数：str  ----要求长度的字符串
	'返回值：字符串长度
	'**************************************************
	Public Function strLength(Str)
		On Error Resume Next
		Dim WINNT_CHINESE:WINNT_CHINESE = (Len("中国") = 2)
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

	Public Function FilterIDs(byval strIDs)
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
	
	'**************************************************
	'取得栏目的链接URL
	'**************************************************
	Public Function GetFolderPath(ChannelID,ClassID)
	    GetFolderPath=GetDomain & "List.asp?ID=" & ClassID & "&amp;" & WapValue & ""
	End Function
	
	'***************************************************************************
	'函数名: GetInfoUrl
	'功 能: 取得每篇文章、图片等的Url链接
	'****************************************************************************
	Public Function GetInfoUrl(ByVal ChannelID,InfoID,ByVal Fname)
	    On Error Resume Next
	    IF Not Isnumeric(ChannelID) Then GetInfoUrl="#":Exit Function
		GetInfoUrl=GetDomain & "Show.asp?ID=" & InfoID & "&amp;ChannelID=" & ChannelID & "&amp;" & WapValue & ""
	End Function
	
	'**************************************************
	'函数名：ChkClng
	'作  用：检查是否是数字 ，并转换为长整型
	'**************************************************
	Public Function ChkClng(ByVal str)
	    On Error Resume Next
		If IsNumeric(str) Then
		   ChkClng = CLng(str)
		Else
		   ChkClng = 0
		End If
		If Err Then ChkClng=0
	End Function
	
	'*************************************   
	'函数名：IsValidChars
	'检测是否只包含英文和数字    
	'*************************************     
	Public Function IsValidChars(str)    
	    Dim re,chkstr    
		Set re=new RegExp    
		re.IgnoreCase =True    
		re.Global = True   
		re.Pattern="[^_\.a-zA-Z\d]"   
		IsValidChars = True   
		chkstr=re.Replace(str,"")    
		If chkstr<>str Then IsValidChars=False   
		set re=nothing    
	End Function
	
	'*************************************    
	'检测是否有效的数字    
	'*************************************    
	Public Function IsInteger(Para)     
	    IsInteger=False   
		If Not (IsNull(Para) Or Trim(Para)="" Or Not IsNumeric(Para)) Then   
		   IsInteger=True   
		End If   
	End Function
	
	
	'**************************************************
	'函数名：MakeRandom
	'作  用：生成指定位数的随机数
	'参  数： maxLen  ----生成位数
	'返回值：成功:返回随机数
	'**************************************************
	Public Function MakeRandom(ByVal maxLen)
	    Dim strNewPass,whatsNext, upper, lower, intCounter
		Randomize
		For intCounter=1 To maxLen
		    upper=57:lower = 48:strNewPass = strNewPass & Chr(Int((upper - lower + 1) * Rnd + lower))
		Next
		MakeRandom = strNewPass
	End Function
	'生成随机密码
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
	'函数名：MakeRandomChar
	'作  用：生成指定位数的随机数字符串 如 "sJKD_!@KK"
	'参  数： Length  ----生成位数
	'返回值：成功返回随机字符串
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
	'函数名：rndNum
	'作  用：生成指定位数的随机数字 如 "222222222222"
	'参  数：strLong----生成位数
	'返回值：成功返回随机字符串
	'**************************************************
	Public Function rndNum(strLong) 
        Dim temNum 
		Randomize 
		Do While Len(RndNum) < strLong 
		   TemNum=CStr(Chr((57-48)*rnd+48)) 
		   RndNum=RndNum&temNum 
		loop 
	End Function 

	'**************************************************
	'函数名：GetFileName
	'作  用：构造文件名。
	'参  数：FsoType  ----生成类型,addDate   -----添加时间,GetFileNameType--扩展名
	'**************************************************
	Public Function GetFileName(FsoType, AddDate, GetFileNameType)
		Dim N
		Randomize
		N = Rnd * 10 + 5
		Select Case FsoType
		    Case 1:GetFileName = Year(AddDate) & "/" & Month(AddDate) & "-" & Day(AddDate) & "/" & MakeRandom(N) & GetFileNameType  '年/月-日/随机数+扩展名
			Case 2:GetFileName = Year(AddDate) & "/" & Month(AddDate) & "/" & Day(AddDate) & "/" & MakeRandom(N) & GetFileNameType '年/月/日/随机数+扩展名
			Case 3:GetFileName = Year(AddDate) & "-" & Month(AddDate) & "-" & Day(AddDate) & "/" & MakeRandom(N) & GetFileNameType '年-月-日/随机数+扩展名
			Case 4:GetFileName = Year(AddDate) & "/" & Month(AddDate) & "/" & MakeRandom(N) & GetFileNameType '年/月/随机数+扩展名
			Case 5:GetFileName = Year(AddDate) & "-" & Month(AddDate) & "/" & MakeRandom(N) & GetFileNameType '年-月/随机数+扩展名
			Case 6:GetFileName = Year(AddDate) & Month(AddDate) & Day(AddDate) & "/" & MakeRandom(N) & GetFileNameType '年月日/随机数+扩展名
			Case 7:GetFileName = Year(AddDate) & "/" & MakeRandom(N) & GetFileNameType '年/随机数+扩展名
			Case 8:GetFileName = Year(AddDate) & Month(AddDate) & Day(AddDate) & MakeRandom(N) & GetFileNameType '年+月+日+随机数+扩展名
			Case 9:GetFileName = MakeRandom(N) & GetFileNameType
			Case 10:GetFileName = MakeRandomChar(N) & GetFileNameType '随机字符
			Case 11:GetFileName ="ID"
			Case Else
			GetFileName = Year(AddDate) & Month(AddDate) & Day(AddDate) & GetFileNameType '12位随机数+扩展名
		End Select
	End Function
	
	'**************************************************
	'函数名：GetFolderSize
	'作  用：取得目录大小
	'参  数：FolderPath--目录文件夹
	'**************************************************
	Public Function GetFolderSize(FolderPath)
		Dim fso:Set FSO = Server.CreateObject(Setting(99))
		If Fso.FolderExists(GetMapPath & FolderPath) Then
		   Dim UserFilespace:set UserFileSpace=FSO.GetFolder(GetMapPath & FolderPath)
		   GetFolderSize=UserFileSpace.size
		Else
		   GetFolderSize=0:exit function
		End If
		set userfilespace=nothing:set fso=nothing
	End Function	
	
	'*******************************************
	'函数作用：格式化文件的大小
	'*******************************************
	Public Function GetFileSize(ByVal size)
	    Dim FileSize
		FileSize=size/1024
		FileSize=FormatNumber(FileSize,2)
		If FileSize < 1024 and FileSize > 1 then
		   GetFileSize=""& FileSize & "KB"
		ElseIf FileSize >1024 then
		   GetFileSize=""& FormatNumber(FileSize / 1024,2) & "MB"
		Else
		   GetFileSize=""& Size & "Bytes"
		End If
	End Function
	
	'==================================================
	'过程名：JpegName
	'作  用：生成缩略图文件名
	'参  数：Str--图片的格式
	'==================================================
	Function JpegFileName(Str)
	    Dim JpegLin,JpegFilesDir,JpegID
	    JpegLin = 500
		JpegFilesDir = "" & Setting(3) & "Cookies/"
		Call CreateListFolder(JpegFilesDir)'创建目录
	    JpegID = Application(SiteSn & "_JpegID")
		If JpegID = "" Then JpegID=1
		If JpegID > JpegLin Then JpegID=1
		Application.Lock
		Application(SiteSn & "_JpegID") = JpegID+1
		Application.UnLock
		JpegFileName = JpegFilesDir & "JpegID_" & JpegID & Str
	End Function

	'==================================================
	'过程名：SaveBeyondFile
	'作  用：保存远程的文件到本地
	'参  数：LocalFileName ------ 本地文件名
	'参  数：RemoteFileUrl ------ 远程文件URL
	'==================================================
	Function SaveBeyondFile(LocalFileName,RemoteFileUrl)
	    On Error Resume Next
		SaveBeyondFile=True
		dim Ads,Retrieval,GetRemoteData
		Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
		With Retrieval
			.Open "Get", RemoteFileUrl, False, "", ""
			.Send
			If .Readystate<>4 then
				SaveBeyondFile=False
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
		If Err.Number<>0 Then
		   Err.Clear
		   SaveBeyondFile=False
		   Exit Function
		End If
		Set Ads=nothing
	End Function

	'****************************************************
	'参数说明
	  'Subject     : 邮件标题
	  'MailAddress : 发件服务器的地址,如smtp.163.com
	  'LoginName     ----登录用户名(不需要请填写"")
	  'LoginPass     ----用户密码(不需要请填写"")
	  'Email       : 收件人邮件地址
	  'Sender      : 发件人姓名
	  'Content     : 邮件内容
	  'Fromer      : 发件人的邮件地址
	'****************************************************
	  Public Function SendMail(MailAddress, LoginName, LoginPass, Subject, Email, Sender, Content, Fromer)
	      On Error Resume Next
		  Dim JMail
		  Set jmail = Server.CreateObject("JMAIL.Message") '建立发送邮件的对象
			jmail.silent = true '屏蔽例外错误，返回FALSE跟TRUE两值j
			jmail.Charset = "GB2312" '邮件的文字编码为国标
			jmail.ContentType = "text/html" '邮件的格式为HTML格式
			jmail.AddRecipient Email '邮件收件人的地址
			jmail.From = Fromer '发件人的E-MAIL地址
			jmail.FromName = Sender
			  If LoginName <> "" And LoginPass <> "" Then
				JMail.MailServerUserName = LoginName '您的邮件服务器登录名
				JMail.MailServerPassword = LoginPass '登录密码
			  End If

			jmail.Subject = Subject '邮件的标题 
			JMail.Body = Content
			JMail.Priority = 1'邮件的紧急程序，1 为最快，5 为最慢， 3 为默认值
			jmail.Send(MailAddress) '执行邮件发送（通过邮件服务器地址）
			jmail.Close() '关闭对象
		Set JMail = Nothing
		If Err Then
			SendMail = Err.Description
			Err.Clear
		Else
			SendMail = "OK"
		End If
	  End Function


	'**************************************************
	'函数名：CreateListFolder
	'作  用：不限分级创建目录 形如 1\2\3\ 则在网站根目录下创建分级目录
	'参  数：Folder要创建的目录
	'返回值：成功返回true 否则返回Flase
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
		If Not FSO.FolderExists(GetMapPath & Folder) Then
		   SplitFolder = Split(Folder, "/")
		   For k = 0 To UBound(SplitFolder) - 1
		       If k = 0 Then
			      CF = SplitFolder(k) & "/"
			   Else
			      CF = CF & SplitFolder(k) & "/"
			   End If
			   If (Not FSO.FolderExists(GetMapPath &CF)) Then
			      FSO.CreateFolder (GetMapPath &CF)
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
	'函数名：DeleteFile
	'作  用：删除指定文件
	'参  数：FileStr要删除的文件
	'返回值：成功返回true 否则返回Flase
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
	'函数名：GetFieSize
	'作用：取得指定文件的大小
	'参数：FilePath--文件位置
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

	'------------------检查某一目录是否存在-------------------
	Public Function CheckDir(FolderPath)
	    Dim Fso1
		FolderPath = Server.MapPath(".") & "\" & FolderPath
		Set Fso1 = CreateObject(Setting(99))
		If Fso1.FolderExists(FolderPath) Then
		   CheckDir = True
		Else
		   CheckDir = False
		End If
		Set Fso1 = Nothing
	End Function
	'------------------检查某一文件是否存在-------------------
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

	'********************************************
	'函数名：IsValidEmail
	'作  用：检查Email地址合法性
	'参  数：email ----要检查的Email地址
	'返回值：True  ----Email地址合法
	'       False ----Email地址不合法
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

	Public Function GetEncodeConversion(str)
		Dim Index,Index_Right,Tag,Tag1,Txt1,Del_Tag,str_Tmp
		str=Trim(str)
		str=HTMLCode(str)
		str=GetURLConversion(str)
		
		Dim str_Card
		str_Card=CutFixContent(str,"<card ",">",1)
		str=Replace(str,"&nbsp;","&#32;")
		str=Replace(str,"&#nbsp;"," ")
		Do while Trim(str)<>""
		   Index=0
		   Index=InStr(1,str,"<",1)
		   If Index=1 then
		      Index_Right=InStr(1,str,">",1)
			  Tag=left(str,Index_Right)
			  If Mid(Tag,2,1)<>"/" then
			     Tag1=Alone_Tag(Tag)
				 Txt1=Txt1+Tag1
				 Del_Tag=Len(Tag)
			  Else
			     Txt1=Txt1+LCase(Tag)
			     Del_Tag=Len(Tag)
			  End if
		   Else
		      If Index>1 then
		         str_Tmp=Left(str,Index-1)
				 Txt1=Txt1+str_Tmp
				 Del_Tag=Len(Left(str,Index-1))
		     End If
			 If Index=0 Then
				 Txt1=Txt1+str
				 Del_Tag=Len(str)
		      End If
		   End If
		   str=Right(str,Len(str)-Del_Tag)
	    Loop
		Dim strArray1,strArray2,I
		strArray1=Array(" ="""" "," >","<tbody>","</tbody>","=""/""","//>","/""/>","&ldquo;","&rdquo;","&#32;")
		strArray2=Array(" ",">","","","","/>","""/>","","","")
		For I=0 To UBound(strArray1)
		    Txt1=Replace(Txt1,strArray1(I),strArray2(I))
		Next
	    TXT1=Replace(TXT1,"&","&amp;")
	    Txt1=Replace(Txt1,"@@@","&amp;")
		TXT1=Replace(TXT1,CutFixContent(Txt1,"<card ",">",1),str_Card)
	    GetEncodeConversion=Txt1
    End Function
    Function GetURLConversion(str)
        Dim Re,URLContents,URLContent
	    Set Re = New Regexp
	    Re.IgnoreCase = True
	    Re.Global = True
	    Re.Pattern = "(href=|onpick=|src=)('|"&CHR(34)&")(.*?)('|"&CHR(34)&")"
	    Set URLContents=Re.Execute(str)
	    IF URLContents.Count<>0 Then
	       For Each URLContent In URLContents
		       str=Replace(str,URLContent,Replace(URLContent,"&","@@@"))
		   Next
	    End IF
	    set Re = Nothing
	    GetURLConversion=str
    End Function
	Function Func_Flag1(str)
	    Dim Index,str1,str2,str3
		Index=InStr(1,str,"=",1)
		str1=Left(str,Index)
		str2=""""
		str3=Mid(str,Index+1,Len(str)-Len(str1))
		Func_Flag1=str1+str2+str3+str2
	End Function
	Function Func_Flag2(str)
		Func_Flag2=str&"="""&str&""""
	End Function
	Function Func_Flag3(str)
		Func_Flag3=Replace(Cstr(str),">","/>")
	End Function
	Function Alone_Tag(Tag)
		Dim Index,Tmpattri,Attribute,Count,Flag,Attribute_Tmp,Tag1,strTag
		Tag=LCase(Tag)
		Index=InStr(1,Tag," ",1)
		Tmpattri=Right(Tag,Len(Tag)-Index)
		If Len(Tmpattri)>1 Then
		   Tmpattri=Trim(left(Tmpattri,Len(Tmpattri)-1))
		End If
		Tmpattri=Replace(Tmpattri,Chr(13)," ")
		Tmpattri=Replace(Tmpattri,Chr(10)," ")
		Tmpattri=Replace(Tmpattri,Chr(10)&Chr(13)," ")
		Tmpattri=Replace(Tmpattri,"  "," ")
		Tmpattri=Replace(Tmpattri,"  "," ")
		Tmpattri=Replace(Tmpattri,"  "," ")
		Tmpattri=Replace(Tmpattri,"  "," ")
		Tmpattri=Replace(Tmpattri,"  "," ")
		Tmpattri=Replace(Tmpattri,"  "," ")
		Tmpattri=Replace(Tmpattri,"  "," ")
		Attribute=Split(Tmpattri, " ", -1, 1)
		For Count=0 to UBound(Attribute, 1)
			If InStr(1,Attribute(Count),"=",1)=0 Then
			   Flag=2
			Else
			   IF InStr(1,Attribute(Count),"""",1)=0 Then
			      Flag=1
			   Else
			      Flag=0
				  IF InStr(1,Attribute(Count),"""",1)>0 Then
				     Flag=4
				  End IF
			   End If
			End If
			Select Case Flag
			    Case 0 Attribute(Count)=Attribute(Count)
				Case 1 Attribute(Count)=Func_Flag1(Attribute(Count))
				Case 2 Attribute(Count)=Func_Flag2(Attribute(Count))
				Case 3 Attribute(Count)=Func_Flag3(Attribute(Count))
				Case 4 Attribute(Count)=(Attribute(Count))
			End Select
		Next
		Count=0
		For Count=0 to UBound(Attribute, 1)
		    Attribute_Tmp=Attribute_Tmp&" "&Attribute(Count)
		Next
		Index=InStr(1,Tag," ",1)
		If InStr(1,Tag," ",1)=0 And Len(Tag)<>"" Then
		   Tag1=Replace(Tag,">"," >")
		Else
		   Tag1=Left(Tag,Index-1) & Attribute_Tmp & ">"
		End If
		strTag = Split("<input ,<img ,<hr ,<br ,<meta", ",", -1, 1)
		For Count=0 to UBound(strTag,1)
		    If InStr(1,Tag1,strTag(Count),1)<>0 Then
			   Tag1=Func_Flag3(Tag1)
			End If
		 Next
		 Alone_Tag=Tag1
	 End Function


	Public Function HTMLEncode(str)
		If Not IsNull(str) Then
		   str = ClearBadChr(str)
		   str = Replace(str, "&", "&amp;")
		   str = Replace(str, "'", "&#39;")
		   str = Replace(str, ">", "&gt;")
		   str = Replace(str, "<", "&lt;")
		   str = Replace(str, Chr(32), " ")
		   str = Replace(str, Chr(9), " ")
		   str = Replace(str, Chr(34), "&quot;")
		   str = Replace(str, Chr(39), "&#39;")
		   str = Replace(str, Chr(13), "")
		   'str = Replace(str, " ", "&nbsp;")
		   'str = Replace(str, Chr(10), "<br />")
		   HTMLEncode = str
		End If
	End Function
	
	Function ClearBadChr(str)
	    If str<>"" Then
		   Dim Re:Set Re=New RegExp
		   Re.IgnoreCase =True
		   Re.Global=True
		   Re.Pattern="(on(load|click|dbclick|mouseover|mouseout|mousedown|mouseup|mousewheel|keydown|submit|change|focus)=""[^""]+"")"
		   str = Re.Replace(str, "")
		   Re.Pattern="((name|id|class)=""[^""]+"")"
		   str = Re.Replace(str, "")
		   Re.Pattern = "(<s+cript[^>]*?>([\w\W]*?)<\/s+cript>)"
		   str = Re.Replace(str, "")
		   Re.Pattern = "(<iframe[^>]*?>([\w\W]*?)<\/iframe>)"
		   str = Re.Replace(str, "")
		   Re.Pattern = "(<p>&nbsp;<\/p>)"
		   str = Re.Replace(str, "")
		   Set Re=Nothing
		   ClearBadChr = str
		End If	
	End Function
	
	Public Function HTMLCode(str)
		If Not IsNull(str) Then
		   'str = Replace(str, "&nbsp;", " ")
		   str = Replace(str, "&quot;", Chr(34))
		   str = Replace(str, "&#39;", Chr(39))
		   str = Replace(str, "&#123;", Chr(123))
		   str = Replace(str, "&#125;", Chr(125))
		   str = Replace(str, "&#36;", Chr(36))
		   str = Replace(str, "&amp;", "&")
		   'str = Replace(str, vbCrLf, "")
		   str = Replace(str, "&gt;", ">")
		   str = Replace(str, "&lt;", "<")
		   HTMLCode = str
		End If
	End Function
	
	'================================================
	'函数名：ReplaceTrim
	'作  用：过滤掉字符中所有的tab和回车和换行
	'================================================
	Public Function ReplaceTrim(ByVal strContent)
	    On Error Resume Next
		Dim re
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "(" & Chr(8) & "|" & Chr(9) & "|" & Chr(10) & "|" & Chr(13) & ")"
		strContent = re.Replace(strContent, vbNullString)
		Set re = Nothing
		ReplaceTrim = strContent
		Exit Function
	End Function

	'**************************************************
	'函数名：UBBToHTML
	'作  用：Ubb转Html
	'参  数：TempStr--文件内容
	'**************************************************
	Public Function UBBToHTML(strContent)
	    Dim re
	    strContent = Trim(strContent)
	    If IsNull(strContent) Then Exit Function
	    Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
	    strContent = Replace(strContent,"[br]","<br/>")
		strContent = Replace(strContent,"<br/><br/>","<br/>")
	    re.Pattern = "(\[img\])(.[^\[]*)(\[\/img\])"
	    strContent = re.Replace(strContent,"<a href=""$2""><img src=""$2"" alt=""""/></a>")
	    re.Pattern = "(\[url\])(.[^\[]*)(\[\/url\])"
	    strContent = re.Replace(strContent,"<a href=""$2"" >$2</a>")
	    re.Pattern = "(\[url=(.[^\]]*)\])(.[^\[]*)(\[\/url\])"
	    strContent = re.Replace(strContent,"<a href=""$2"" >$3</a>")
	    re.Pattern = "^(http://[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
	    strContent = re.Replace(strContent,"<a href=""$1"">$1</a>")
	    Set re = Nothing
	    UBBToHTML = strContent
	End Function
	
	'**************************************************
	'函数名：HTMLToUBB
	'作  用：HTML转UBB
	'参  数：TempStr--文件内容
	'**************************************************
	Public Function HTMLToUBB(strContent)
	    Dim re
		strContent = Trim(strContent)
		If IsNull(strContent) Then Exit Function
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		strContent = Replace(strContent,"</p><p>","[br]")
		strContent = Replace(strContent,"<br />","[br]")
		strContent = Replace(strContent,"<br/>","[br]")
		strContent = Replace(strContent,"<br>","[br]")
		strContent = Replace(strContent,"<BR>","[br]")
		
		re.Pattern = "<a[^>]+href=('|"&CHR(34)&")([A-Za-z0-9\./=\?%\-&_&#x7E;`@':+!,;*()#]+)('|"&CHR(34)&")[^>]*><img[^>]+src=('|"&CHR(34)&")([A-Za-z0-9\./=\?%\-&_&#x7E;`@':+!,;*()#]+)('|"&CHR(34)&")[^>]*></a>"
		strContent = re.Replace(strContent,"[url=Plus/PhotoDownLoad.asp?JpegUrl=$5]浏览图片[/url]")
		re.Pattern = "<img[^>]+src=('|"&CHR(34)&")([A-Za-z0-9\./=\?%\-&_&#x7E;`@':+!,;*()#]+)('|"&CHR(34)&")[^>]*>"
		strContent = re.Replace(strContent,"[url=Plus/PhotoDownLoad.asp?JpegUrl=$2]浏览图片[/url]")
		'strContent = re.Replace(strContent,"[img]$2[/img]")
		re.Pattern = "<a[^>]+href=('|"&CHR(34)&")([A-Za-z0-9\./=\?%\-&_&#x7E;`@':+!,;*()#]+)('|"&CHR(34)&")[^>]*>(.*?)</a>"
		strContent = re.Replace(strContent,"[url=$2]$4[/url]")
		Set re = Nothing
		HTMLToUBB = strContent
	End Function
	
	'取消HTML
	Public Function LoseHtml(ContentStr)
	    On Error Resume Next
		Dim TempLoseStr, regEx
		If ContentStr="" Or ContentStr=Null Then Exit Function
		TempLoseStr = CStr(ContentStr)
		Set regEx = New RegExp
		regEx.Pattern = "<\/*[^<>]*>"
		regEx.IgnoreCase = True
		regEx.Global = True
		TempLoseStr = regEx.Replace(TempLoseStr, "")
		LoseHtml = TempLoseStr
	End Function
	
	'**************************************************
	'函数名：JoinChar
	'作  用：向地址中加入 ? 或 &amp; 
	'**************************************************
    Public Function JoinChar(Byval strUrl)
		If strUrl = "" Then JoinChar = "" : Exit Function
		If InStr(strUrl,"?")<len(strUrl) Then
		   If InStr(strUrl,"?")>1 Then
			  If InStr(strUrl,"&amp;")<len(strUrl) Then 
			     JoinChar=strUrl & "&amp;"
			  Else
			     JoinChar=strUrl
			  End If
		   Else
			  JoinChar=strUrl & "?"
	       End If
		Else
		   JoinChar=strUrl
		End If
	End Function
	
	'**************************************************
	'函数名：CreateKeyWord
	'作  用：由给定的字符串生成关键字
	'参  数：Constr---要生成关键字的原字符串
	'返回值：生成的关键字
	'**************************************************
	Public Function CreateKeyWord(byval Constr,Num)
	    If Constr="" or IsNull(Constr)=True Then
		   CreateKeyWord=""
		   Exit Function
		End If
		If Num="" or IsNumeric(Num)=False Then
		   Num=2
		End If
		Constr=Replace(Constr,CHR(32),"")
		Constr=Replace(Constr,CHR(9),"")
		Constr=Replace(Constr,"&nbsp;","")
		Constr=Replace(Constr," ","")
		Constr=Replace(Constr,"(","")
		Constr=Replace(Constr,")","")
		Constr=Replace(Constr,"<","")
		Constr=Replace(Constr,">","")
		Constr=Replace(Constr,"""","")
		Constr=Replace(Constr,"?","")
		Constr=Replace(Constr,"*","")
		Constr=Replace(Constr,"|","")
		Constr=Replace(Constr,",","")
		Constr=Replace(Constr,".","")
		Constr=Replace(Constr,"/","")
		Constr=Replace(Constr,"\","")
		Constr=Replace(Constr,"-","")
		Constr=Replace(Constr,"@","")
		Constr=Replace(Constr,"#","")
		Constr=Replace(Constr,"$","")
		Constr=Replace(Constr,"%","")
		Constr=Replace(Constr,"&","")
		Constr=Replace(Constr,"+","")
		Constr=Replace(Constr,":","")
		Constr=Replace(Constr,"：","")   
		Constr=Replace(Constr,"‘","")
		Constr=Replace(Constr,"“","")
		Constr=Replace(Constr,"”","")
		Constr=Replace(Constr,"&","")  
		Constr=Replace(Constr,"gt;","")      
		Dim i,ConstrTemp
		For i=1 To Len(Constr)
		    ConstrTemp=ConstrTemp & "|" & Mid(Constr,i,Num)
		Next
		If Len(ConstrTemp)<254 Then
		   ConstrTemp=ConstrTemp & "|"
		Else
		   ConstrTemp=Left(ConstrTemp,254) & "|"
		End If
		ConstrTemp=left(ConstrTemp,len(ConstrTemp)-1)
		ConstrTemp= Right(ConstrTemp,len(ConstrTemp)-1)
		CreateKeyWord=ConstrTemp
	End Function

	'*************************************************************************************
	'函数名:GetGQTypeName
	'作  用:获得供求的交易类别名称
	'参  数:TypeID
	'*************************************************************************************
	Public Function GetGQTypeName(TypeID)
	    If Not IsNumeric(TypeID) Then GetGQTypeName="":Exit Function
		Dim KS_RS_Obj:Set KS_RS_Obj=Server.CreateObject("ADODB.RECORDSET")
		KS_RS_Obj.Open "Select TypeName,TypeColor From KS_GQType Where TypeID=" & TypeID,Conn,1,1
		If Not KS_RS_Obj.Eof Then
	       GetGQTypeName=KS_RS_Obj(0)
	    Else 
	       GetGQTypeName=""
	    End If
		GetGQTypeName=Replace(Replace(GetGQTypeName,"【",""),"】","")
	    KS_RS_Obj.Close:Set KS_RS_Obj=Nothing
	End Function
	'返回供求交易类型列表
	'参数：Flag:1-标签调用 0-添加信息时调用
	Public Function ReturnGQType(SelID,Flag)
	    Dim SQL,K,RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
	    RS.Open "Select TypeID,TypeName,TypeColor From KS_GQType Order By TypeID",Conn,1,1
	    If Flag=1 Then 
	   	   ReturnGQType="<select name=""TypeID"">"
		   ReturnGQType=ReturnGQType & "<option value=""0"">交易类型不限</option>"
	    Else
	   	   ReturnGQType="<select name=""TypeID"">"
	    End If
	   
	    SQL=RS.GetRows(-1):RS.CLose:Set RS=Nothing
	    For K=0 To Ubound(SQL,2)
			ReturnGQType=ReturnGQType & "<option value=""" & sql(0,k) & """>" & sql(1,k) & "</option>"
        Next
	    ReturnGQType=ReturnGQType & "</select>"
	End Function
	'*************************************************************************************
	'函数名:GetInfoID
	'作  用:生成文章,图片或下载等的唯一ID
	'参  数:ChannelID--频道ID
	'*************************************************************************************
	Public Function GetInfoID(ChannelID)
	    On Error Resume Next
	    Dim RSC, TableNameStr
        Set RSC=Server.CreateObject("ADODB.RECORDSET")
	    Select Case C_S(ChannelID,6)
			Case 5:TableNameStr = "Select ProID From " & KS.C_S(ChannelID,2) & " Where ProID='"
	    End Select
		Do While True
		   GetInfoID = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Now(), "-", ""), " ", ""), ":", ""), "PM", ""), "AM", ""), "上午", ""), "下午", "") & MakeRandom(3)
		   RSC.Open TableNameStr & GetInfoID & "'", Conn, 1, 1
		   If RSC.EOF And RSC.BOF Then Exit Do
	    Loop
		RSC.Close:Set RSC = Nothing
	End Function

	'*************************************************************************************
	'函数名:ReplaceInnerLink
	'作  用:替换站内链接
	'参  数:Content-待替换内容
	'*************************************************************************************
	Public Function ReplaceInnerLink(Content)
	    If Not IsObject(Application(SiteSN & "_InnerLink")) Then
		   Dim RS:Set RS = Conn.Execute("Select Title,Url,OpenType From KS_InnerLink Where OpenTF=1 Order By ID")
		   Set Application(SiteSN & "_InnerLink")=RecordsetToxml(RS,"InnerLink","InnerLinkList")
		   Set RS = Nothing
		End If
	    Dim Node
		For Each Node In Application(SiteSN & "_InnerLink").DocumentElement.SelectNodes("InnerLink")
			If InStr(Content,Node.selectSingleNode("@ks0").text)>0 Then
			   Dim OpenTypeStr:OpenTypeStr = G_O_T_S(Node.selectSingleNode("@ks2").text)
			   Content= Replace(Content,Node.selectSingleNode("@ks0").text,"<a href="""&Node.selectSingleNode("@ks1").text&"""" & OpenTypeStr & ">"&Node.selectSingleNode("@ks0").text&"</a>")
			End if
		Next
		ReplaceInnerLink = HTMLCode(Content)
	End Function

	'=================================================
	'函数名：ContentPage
	'作  用：取出内容分页
	'参  数：sfilename--地址,sContent--文章内容,sLen--每页显示的字符,strEnable--是否显示转到该页:True否:False
	'=================================================
	Public Function ContentPage(Byval sfilename,Byval sContent,Byval sLen,Byval strEnable)
        Dim aLen, bLen
		Dim n, sTemp, strUrl
		CurrentPage=Int(Abs(Request("CurrentPage")))
		aLen = Len(sContent)
		strUrl = JoinChar(sfilename)'向地址中加入 ? 或 &amp; 
		If (aLen Mod sLen)=0 Then
		   n = aLen \ sLen
		Else
		   n = aLen \ sLen + 1
		End if
		If CurrentPage > n Then CurrentPage=n
		   If CurrentPage <= 1 Then
		   CurrentPage=1
		   sTemp = Left(sContent,sLen)
		   if n<2 then ContentPage=sTemp : Exit Function
		Else
		   bLen = sLen*(CurrentPage-1)+1
		   'eLen = sLen*CurrentPage
		   sTemp = mid(sContent,bLen,sLen)
		End if
		sTemp = sTemp & "<br /> "
		sTemp = sTemp & "第" & CurrentPage & "页,共" & n & "页"
		sTemp = sTemp & "<br /> "
		If CurrentPage < n Then
	       sTemp = sTemp & "<a href=""" & strUrl & "CurrentPage=" & (CurrentPage+1) & """>下页</a> "
		   sTemp = sTemp & "<a href=""" & strUrl & "CurrentPage=" & n & """>尾页</a>"
	    Else
	       sTemp = sTemp & "下页 尾页"
	    End if
		sTemp = sTemp & "  "
	    If CurrentPage > 1 Then
	       sTemp = sTemp & "<a href=""" & strUrl & "CurrentPage=" & (CurrentPage-1) & """>上页</a> "
		   sTemp = sTemp & "<a href=""" & strUrl & "CurrentPage=1"">首页</a> "
	    Else
	       sTemp = sTemp & "上页 首页"
	    End if
	    If strEnable = "True" Then
	       sTemp = sTemp & "<br /> <input name=""CurrentPage"&minute(now)&second(now)&""" format=""*N"" emptyok=""true"" size=""3"" type=""text"" value=""" & (CurrentPage+1) & """ title=""请输入页码"" maxlength=""9""/>"
		   sTemp = sTemp & "<anchor><go href=""" & sFileName & """ method=""post"">"
		   sTemp = sTemp & "<postfield name=""CurrentPage"" value=""$(CurrentPage"&minute(now)&second(now)&")""/>"
		   sTemp = sTemp & "</go>[转到该页]</anchor>"
	    End if
	    ContentPage = sTemp
	End Function
	
	'显示分页的前部分
	'参数说明:PageStyle-分页样式,ItemUnit-单位,TotalPage-总页数,CurrPage-当前第N页,TotalInfo-总信息数,PerPageNumber-每页显示数
	Function  GetPrePageList(PageStyle,ItemUnit,TotalPage,CurrPage,TotalInfo,PerPageNumber)
	    Select Case  Cint(PageStyle)
		    Case 1:GetPrePageList= "共"&TotalInfo&""&ItemUnit&"页次:"&CurrPage&"/"&TotalPage&"页,"&PerPageNumber&""&ItemUnit&"/页<br/>"
			Case 2:GetPrePageList= ""&CurrPage&"/"&TotalPage&"页,"&PerPageNumber&""&ItemUnit&"/页<br/>"
			'Case 3:GetPrePageList= "第"&CurrPage&"页,共"&TotalPage&"页<br/>"
		End Select
	End Function
	
	'动态显示分页
	Function GetPageList(Byval FileName,Byval PageStyle,Byval CurrPage,Byval TotalPage,Byval ShowTurnToFlag)
		Dim PageStr, I, J, SelectStr
		If ChkClng(PageStyle)=0 Then PageStyle=1
		Select Case PageStyle
		    Case 1
			If CurrPage = 1 And CurrPage <> TotalPage Then
			   PageStr = "<a href="""&FileName&"Page="&CurrPage+1&""">下一页</a> <a href="""&FileName&"Page="&TotalPage&""">尾页</a> 首页 上一页"
			ElseIf CurrPage = 1 And CurrPage = TotalPage Then
			   PageStr = "下一页 尾页 首页 上一页"
			ElseIf CurrPage = TotalPage And CurrPage <> 2 Then  '对于最后一页刚好是第二页的要做特殊处理
			   PageStr = "<a href="""&FileName&""">首页</a> <a href="""&FileName&"Page="&CurrPage-1&""">上一页</a> 下一页 尾页"
			ElseIf CurrPage = TotalPage And CurrPage = 2 Then
			   PageStr = "<a href="""&FileName&""">首页</a> <a href="""&FileName&""">上一页</a> 下一页 尾页"
			ElseIf CurrPage = 2 Then
			   PageStr = "<a href="""&FileName&"Page="&CurrPage+1&""">下一页</a> <a href="""&FileName&"Page="&TotalPage&""">尾页</a> <a href="""&FileName&""">首页</a> <a href="""&FileName&""">上一页</a>"
			Else
			   PageStr = "<a href="""&FileName&"Page="&CurrPage+1&""">下一页</a> <a href="""&FileName&"Page="&TotalPage&""">尾页</a> <a href="""&FileName&""">首页</a> <a href="""&FileName&"Page="&CurrPage-1&""">上一页</a>"
			End If
			Case 2
			If CurrPage=TotalPage Then
			   PageStr="下页 尾页"
			Else
			   PageStr="<a href="""&FileName&"Page="&CurrPage+1&""">下页</a> <a href="""&FileName&"Page="&TotalPage&""">尾页</a> "
			End If
			Dim startpage,n
			startpage=1
			If (CurrPage>=5) Then startpage=(CurrPage\5-1)*5+CurrPage mod 5+2
			For J=startpage To TotalPage
			    'If J>TotalPage Then Exit For
				If J= CurrPage Then
				   PageStr=PageStr & "["&J&"]"
				Else
				   PageStr=PageStr & "<a href="""&FileName&"Page="&J&""">" & J &"</a> "
				End If
				n=n+1
				If n>=5 Then Exit For
			Next
			If CurrPage=1 Then
			   PageStr=PageStr & "首页 上页"
			ElseIf CurrPage=2 Then
			   PageStr=PageStr & "<a href="""&FileName&""">首页</a> <a href="""&FileName&"Page="&CurrPage-1&""">上页</a> "
			Else
			   PageStr=PageStr & "<a href="""&FileName&""">首页</a> <a href="""&FileName&"Page="&CurrPage-1&""">上页</a> "
			End If
			Case 3
			If CurrPage=TotalPage Then
			   PageStr=" 下页 尾页"
			Else
			   PageStr=" <a href="""&FileName&"Page="&CurrPage+1&""">下页</a> <a href="""&FileName&"Page="&TotalPage&""">尾页</a> "
			End If
			If CurrPage=1 Then
			   PageStr=PageStr & "首页 上页"
			ElseIf CurrPage=2 Then
			   PageStr=PageStr & "<a href="""&FileName&""">首页</a> <a href="""&FileName&""">上页</a>"
			Else
			   PageStr=PageStr & "<a href="""&FileName&""">首页</a> <a href="""&FileName&"Page="&CurrPage-1&""">上页</a> "
			End If
		End Select
		If CBool(ShowTurnToFlag) = True Then
		   PageStr = PageStr & "<br /><input name=""Page"&Minute(Now)&Second(Now)&""" format=""*N"" emptyok=""true"" size=""3"" type=""text"" value=""" & (CurrPage+1) & """ title=""请输入页码"" maxlength=""9""/>"
		   PageStr = PageStr & "<anchor><go href=""" & FileName & """ method=""post"">"
		   PageStr = PageStr & "<postfield name=""Page"" value=""$(Page"&Minute(Now)&Second(Now)&")""/>"
		   PageStr = PageStr & "</go>[转到该页]</anchor>"
		End If
		GetPageList=PageStr
	End Function
	
	'**************************************************
	'函数名：ShowPagePara
	'作  用：显示“上一页 下一页”等信息
	'参  数：filename  ----链接地址
	'       TotalNumber ----总数量
	'       MaxPerPage  ----每页数量
	'       ShowAllPages ---是否用下拉列表显示所有页面以供跳转。
	'       strUnit     ----计数单位
	'       CurrentPage ----当前页
	'       ParamterStr ----参数
	'返回值：无返回值
	'**************************************************
	Public Function ShowPagePara(totalnumber, MaxPerPage, FileName, ShowAllPages, strUnit, CurrentPage, ParamterStr)
        Dim N, I, PageStr
		Const Btn_First = "首页" '定义第一页按钮显示样式
		Const Btn_Prev = "上页" '定义前一页按钮显示样式
		Const Btn_Next = "下页" '定义下一页按钮显示样式
		Const Btn_Last = "尾页" '定义最后一页按钮显示样式
		PageStr = ""
		If totalnumber Mod MaxPerPage = 0 Then
		   N = totalnumber \ MaxPerPage
		Else
		   N = totalnumber \ MaxPerPage + 1
		End If
		If N > 1 Then
		   PageStr = PageStr & ("页次" & CurrentPage & "/" & N & "页 共" & totalnumber & strUnit & " 每页" & MaxPerPage & strUnit & "<br/>")
		   If CurrentPage < 2 Then
		      PageStr = PageStr & Btn_First & " " & Btn_Prev & " "
		   Else
		      PageStr = PageStr & ("<a href=""" & FileName & "?page=1" & "&amp;" & ParamterStr & """>" & Btn_First & "</a> <a href=""" & FileName & "?page=" & CurrentPage - 1 & "&amp;" & ParamterStr & """>" & Btn_Prev & "</a> ")
		   End If
		   If N - CurrentPage < 1 Then
		      PageStr = PageStr & " " & Btn_Next & " " & Btn_Last & " "
		   Else
		      PageStr = PageStr & (" <a href=""" & FileName & "?page=" & (CurrentPage + 1) & "&amp;" & ParamterStr & """>" & Btn_Next & "</a> <a href=""" & FileName & "?page=" & N & "&amp;" & ParamterStr & """>" & Btn_Last & "</a> ")
		   End If
		   If ShowAllPages = True Then	   
		      PageStr = PageStr & ("<br /> <input name=""Page" & Minute(Now) & Second(Now) & """ format=""*N"" emptyok=""true"" size=""3"" type=""text"" value=""" & (CurrentPage + 1) & """ title=""请输入页码"" maxlength=""9""/><anchor><go href=""" & FileName & "?" & ParamterStr & """ method=""post""><postfield name=""Page"" value=""$(Page" & Minute(Now) & Second(Now) & ")""/></go>[转到该页]</anchor>")
		   End If
		End If
		ShowPagePara = PageStr
	End Function
	Sub ShowPageParamter(totalnumber, MaxPerPage, FileName, ShowAllPages, strUnit, CurrentPage, ParamterStr)
	    Response.Write (ShowPagePara(totalnumber, MaxPerPage, FileName, ShowAllPages, strUnit, CurrentPage, ParamterStr))
    End Sub
	
	'=================================================
	'过程名：ContentPagination
	'作  用：采用自动分页方式显示文章内容
	'参  数：sContent--文章内容
	'       sLen--每页显示的字数
	'       sUrl--传递过来的参数
	'       sEnable--是否显示转到该页,是True否False
	'       sPageWord--是否显示字符设置,是True否False
	'=================================================
	Public Function ContentPagination(ByVal sContent,ByVal sLen,ByVal sUrl,ByVal sEnable,ByVal sPageWord)
        If ChkClng(S("sLen"))=0 Then
		   sLen=CLng(sLen)
		Else
		   sLen=ChkClng(S("sLen"))
		End If
		Dim cLen:cLen=strLength(sContent)
	    Dim strContent,ContentLen,CurrentPage,arrContent,Paginate,ArticleContent,UserArticle,IsURLRewrite
		Dim m_strFileExt,m_strFileUrl
	    sUrl = JoinChar(sUrl)'向地址中加入 ? 或 &amp; 
	    strContent = InsertPageBreak(sContent,sLen)
	    ContentLen = Len(strContent)
	    CurrentPage=ChkClng(S("CPage"))
	    If CurrentPage="" Then CurrentPage=0
	    If InStr(strContent, "[NextPage]") <= 0 Then
	       ArticleContent = strContent
	    Else
	       arrContent = Split(strContent, "[NextPage]")
		   Paginate = UBound(arrContent) + 1
		   If CurrentPage = 0 Then
		      CurrentPage = 1
		   Else
		      CurrentPage = CLng(CurrentPage)
		   End If
		   If CurrentPage < 1 Then CurrentPage = 1
		   If CurrentPage > Paginate Then CurrentPage = Paginate
		   strContent = arrContent(CurrentPage - 1)
		   ArticleContent = ArticleContent & strContent
		   If UserArticle = True Then
		      ArticleContent = ArticleContent
		   Else
	          ArticleContent = ArticleContent
		   End If
		   If IsURLRewrite Then
	          m_strFileUrl = sUrl
		   Else
	          m_strFileExt = "&amp;sLen="&sLen&""
	          m_strFileUrl = sUrl&"CPage="
		   End If
		   ArticleContent = ArticleContent&"<br/>"
		  
		   If Paginate > 0 Then
		      If CurrentPage < Paginate Then
		         ArticleContent = ArticleContent & " <a href="""& m_strFileUrl & CurrentPage + 1 & m_strFileExt & """>下一页</a>"
			  End If
			  If CurrentPage > 1 Then
		         If IsURLRewrite And (CurrentPage-1) = 1 Then
		            ArticleContent = ArticleContent & " <a href="""& ArticleID & m_strFileExt & """>上一页</a>"
				 Else
		            ArticleContent = ArticleContent & " <a href="""& m_strFileUrl & CurrentPage - 1 & m_strFileExt & """>上一页</a>"
				 End If
	          End If
			  ArticleContent = ArticleContent & ""&CurrentPage&"/"&Paginate&"页"
			  If sEnable = "True" then
			     ArticleContent = ArticleContent & " <input name=""CPage" & Minute(Now) & Second(Now) & """ format=""*N"" emptyok=""true"" size=""3"" type=""text"" value=""" & (CurrentPage+1) & """ title=""请输入页码"" maxlength=""9""/>"
				 ArticleContent = ArticleContent & "<anchor>转到该页<go href=""" & sUrl & """ method=""post"">"
				 ArticleContent = ArticleContent & "<postfield name=""CPage"" value=""$(CPage" & Minute(Now) & Second(Now) & ")""/></go></anchor>"
			  End If
		   End If
	    End If
		If sPageWord = "True" And sLen < cLen Then
		   ArticleContent = ArticleContent & "<br/>当前设置"&sLen&"字/页<br/>"
		   If sLen <> "200" Then
		      ArticleContent = ArticleContent & "<a href="""&sUrl&"sLen=200"">200</a> "
		   End If
		   If sLen <> "400" Then
		      ArticleContent = ArticleContent & "<a href="""&sUrl&"sLen=400"">400</a> "
		   End If
		   If sLen <> "800" Then
		      ArticleContent = ArticleContent & "<a href="""&sUrl&"sLen=800"">800</a> "
		   End If
		   If sLen <> "1600" Then
		      ArticleContent = ArticleContent & "<a href="""&sUrl&"sLen=1600"">1600</a> "
		   End If
		   If sLen <> cLen Then
		      ArticleContent = ArticleContent & "<a href="""&sUrl&"sLen="&cLen&""">全文</a>"
		   End If
		End If
	    ContentPagination = ArticleContent
    End Function
	
	Function InsertPageBreak(Byval strText,Byval sLen)
        Dim strPagebreak,T,SS
		Dim i,IsCount,c,iCount,strTemp,Temp_String,Temp_Array
		strPagebreak="[NextPage]"
		T=strText
		If Len(T)<sLen Then
		   InsertPageBreak=T
		End If
		T=Replace(T, strPagebreak, "")
		T=Replace(T, "&nbsp;", "<&nbsp;>")
		T=Replace(T, "&gt;", "<&gt;>")
		T=Replace(T, "&lt;", "<&lt;>")
		T=Replace(T, "&quot;", "<&quot;>")
		T=Replace(T, "&#39;", "<&#39;>")
		If T<>"" And sLen<>0 And InStr(1,T,strPagebreak)=0 Then
	       IsCount=True
		   Temp_String=""
		   For I= 1 To Len(T)
		       C=Mid(T,I,1)
			   If C="<" Then
			      IsCount=False
			   ElseIf C=">" Then
			      IsCount=True
			   Else
			      If IsCount=True Then
				     If Abs(Asc(C))>255 Then
					    iCount=iCount+2
					 Else
					    iCount=iCount+1
					 End If
					 If iCount>=sLen And I<Len(T) Then
					    strTemp=Left(T,I)
						If CheckPagination(strTemp,"table|a|b/>|i>|strong|div|span") then
						   Temp_String=Temp_String & Trim(CStr(I)) & "," 
						   iCount=0
						End If
					 End If
				  End If
			   End If 
		   Next
		   If Len(Temp_String)>1 Then Temp_String=Left(Temp_String,Len(Temp_String)-1)
		   Temp_Array=Split(Temp_String,",")
		   For I = UBound(Temp_Array) To LBound(Temp_Array) Step -1
		       SS = Mid(T,Temp_Array(I)+1)
			   If Len(SS) > 380 Then
			      T=Left(T,Temp_Array(I)) & strPagebreak & SS
			   Else
			      T=Left(T,Temp_Array(I)) & SS
			   End If
		   Next
	    End If
		T=Replace(T, "<&nbsp;>", "&nbsp;")
		T=Replace(T, "<&gt;>", "&gt;")
		T=Replace(T, "<&lt;>", "&lt;")
		T=Replace(T, "<&quot;>", "&quot;")
		T=Replace(T, "<&#39;>", "&#39;")
		InsertPageBreak=T
	End Function
	
	Function CheckPagination(Byval strTemp,Byval strFind)
        Dim i,n,m_ingBeginNum,m_intEndNum
		Dim m_strBegin,m_strEnd,FindArray
		strTemp=LCase(strTemp)
		strFind=LCase(strFind)
		If strTemp<>"" And strFind<>"" Then
		   FindArray=split(strFind,"|")
		   For i = 0 to Ubound(FindArray)
		       m_strBegin="<"&FindArray(i)
			   m_strEnd   ="</"&FindArray(i)
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
	
	'======================================会员相关函数==================================== 
  	'取得用户组名称
	Public Function GetUserGroupName(GroupID)
	    On Error Resume Next
		GetUserGroupName=Conn.Execute("Select GroupName From KS_UserGroup Where ID=" & GroupID)(0)
		If err Then GetUserGroupName=""
	End Function
	
	
	'会员投稿文章，图片，下载等增加积分,发送站内短信操作
	'参数ChannelID-频道ID,UserName---用户名称,InfoTitle---投稿的主题
	Public Sub SignUserInfoOK(ChannelID,UserName,InfoTitle)
	    IF Not IsNumeric(ChannelID) Then Exit Sub
	    Dim ClientName,RSObj:Set RSObj=Conn.Execute("Select top 1 RealName From KS_User Where UserName='" & UserName & "'")
		IF Not RSObj.Eof Then
					ClientName=RSObj(0):If ClientName="" Then ClientName=UserName
					'成功则发送站内通知信件
					Dim Sender:Sender=Setting(0)
					Dim Title:Title="恭喜，您发表的" & C_S(ChannelID,3) & "[" & InfoTitle & "]已通过审核！！！"
					Dim Message:Message="" & C_S(ChannelID,3) & "标题：" & InfoTitle &"<br>"
					 If C_S(ChannelID,18)>0 Then
					  Message = Message & "获得金钱：<font color=red>" & C_S(ChannelID,18) & "</font> 元人民币<br>"
					 ElseIf C_S(ChannelID,18)<0 Then
					  Message = Message & "消耗金钱：<font color=red>" & Abs(C_S(ChannelID,18)) & "</font> 元人民币<br>"
					 End IF
					 If C_S(ChannelID,19)>0 Then
					  Message = Message & "获得" & Setting(45) & "：<font color=red>" & C_S(ChannelID,19) & "</font> " & Setting(46) & Setting(45) & "<br>"
					 ElseIf C_S(ChannelID,19)<0 Then
					  Message = Message & "消耗" & Setting(45) & "：<font color=red>" & Abs(C_S(ChannelID,19)) & "</font> " & Setting(46) & Setting(45) & "<br>"
					 End If
					 If C_S(ChannelID,20)>0 Then
					  Message = Message & "获得积分：<font color=red>" & C_S(ChannelID,20) & "</font> 分积分<br>"
					 ElseIf C_S(ChannelID,20)<0 Then
					  Message = Message & "消耗积分：<font color=red>" & Abs(C_S(ChannelID,20)) & "</font> 分积分<br>"
					 End If
					  Message = Message & "<br />备注：此信息由系统自动发布，请不要回复！！！"
					If C_S(ChannelID,19)<0 Then  
					Call PointInOrOut(ChannelID,0,UserName,2,-C_S(ChannelID,19),"系统","发表" & C_S(ChannelID,3) & "[" & InfoTitle & "]产生")            
					Else
					Call PointInOrOut(ChannelID,0,UserName,1,C_S(ChannelID,19),"系统","发表" & C_S(ChannelID,3) & "[" & InfoTitle & "]产生")            
					End If
					If C_S(ChannelID,20)<0 Then
					 Call ScoreInOrOut(UserName,2,-C_S(ChannelID,20),"系统","发表" & C_S(ChannelID,3) & "[" & InfoTitle & "]产生")            
					Else
					 Call ScoreInOrOut(UserName,1,C_S(ChannelID,20),"系统","发表" & C_S(ChannelID,3) & "[" & InfoTitle & "]产生")            
					End If
					If C_S(ChannelID,18)<0 Then
					Call MoneyInOrOut(UserName,ClientName,-C_S(ChannelID,18),4,2,SqlNowString,"0","系统","发表" & C_S(ChannelID,3) & "[" & InfoTitle & "]产生")
					Else
					Call MoneyInOrOut(UserName,ClientName,C_S(ChannelID,18),4,1,SqlNowString,"0","系统","发表" & C_S(ChannelID,3) & "[" & InfoTitle & "]产生")
					End If
					Call SendInfo(UserName,Sender,Title,Message)
		End IF
		RSObj.Close:Set RSObj=Nothing
	End Sub
	'功能:会员积分明细出入函数	
	'参数:UserName-用户名,InOrOutFlag-操作类型1收入2支出,Score-交易点数,User-操作员,Descript-操作备注
	Public Function ScoreInOrOut(UserName,InOrOutFlag,Score,User,Descript)
	  If Not IsNumeric(InOrOutFlag) Or Not IsNumeric(Score) Or Score=0 Then ScoreInOrOut=false:Exit Function
	  Dim ScoreParam,CurrScore
	  If InOrOutFlag=1 Then 
	     ScoreParam="Set Score=Score+" & Score
	  ElseIF InOrOutFlag=2 Then
	     ScoreParam="Set Score=Score-" & Score
	  Else
	    ScoreInOrOut=false:Exit Function
	  End If
	  on error resume next
	  Conn.Execute("Update KS_User " & ScoreParam & " Where UserName='" & UserName & "'")
	  CurrScore=Conn.Execute("Select top 1 Score From KS_User Where UserName='" & UserName & "'")(0)
	  Conn.Execute("Insert into KS_LogScore(UserName,InOrOutFlag,Score,CurrScore,[User],Descript,Adddate,IP) values('" & UserName & "',"& InOrOutFlag & "," & Score & ","&CurrScore & ",'" & replace(User,"'","""") & "','" & replace(Descript,"'","""") & "'," & SqlNowString & ",'" & replace(getip,"'","""") & "')")
	  IF Err Then ScoreInOrOut=false Else ScoreInOrOut=true
	End Function
	'功能:资金明细出入函数	                 
	'参数:UserName-用户名,ClientName-客户姓名,Money-金钱,MoneyType-类型,InOrOutFlag-操作类型1收入2支出,PayTime-汇款日期,OrderID-订单号,Inputer-操作员,Remark-操作备注
	Public Function MoneyInOrOut(UserName,ClientName,Money,MoneyType,InorOutFlag,PayTime,OrderID,Inputer,Remark)
	  If Not IsNumeric(InOrOutFlag) Or Not IsNumeric(Money) Or Money=0 Then MoneyInOrOut=false:Exit Function
	  Dim MoneyParam
	  If InOrOutFlag=1 Then 
	     MoneyParam="Set [Money]=[Money]+" & Money
	  ElseIF InOrOutFlag=2 Then
	     MoneyParam="Set [Money]=[Money]-" & Money
	  Else
	    MoneyInOrOut=false:Exit Function
	  End If
	  'on error resume next
	  Conn.Execute("Update KS_User " & MoneyParam & " Where UserName='" & UserName & "'")
	  Conn.Execute("Insert into KS_LogMoney([UserName],[ClientName],[Money],[MoneyType],[IncomeOrPayOut],[OrderID],[Remark],[PayTime],[LogTime],[Inputer],[IP]) values('" & UserName & "','" & ClientName & "'," & Money & "," & MoneyType & ","& InOrOutFlag & ",'" & OrderID & "','" & replace(Remark,"'","""") & "'," & SqlNowString & "," &SqlNowString & ",'" & replace(inputer,"'","""") & "','" & replace(getip,"'","""") & "')")
	  IF Err Then MoneyInOrOut=false Else MoneyInOrOut=true
	End Function	
	'功能:会员点券明细出入函数
	'参数:Channelid-模块ID,InfoID-信息ID，UserName-用户名,InOrOutFlag-操作类型1收入2支出,Point-交易点数,User-操作员,Descript-操作备注
	Public Function PointInOrOut(ChannelID,InfoID,UserName,InOrOutFlag,Point,User,Descript)
	    If Not IsNumeric(InOrOutFlag) Or Not IsNumeric(Point) Then PointInOrOut=false:Exit Function
		Dim PointParam
		If InOrOutFlag=1 Then 
		   PointParam="Set Point=Point+" & Point
		ElseIF InOrOutFlag=2 Then
		   PointParam="Set Point=Point-" & Point
		Else
		   PointInOrOut=false:Exit Function
		End If
		On Error Resume Next
		Conn.Execute("Update KS_User " & PointParam & " Where UserName='" & UserName & "'")
		Conn.Execute("Insert into KS_LogPoint(ChannelID,InfoID,UserName,InOrOutFlag,Point,Times,[User],Descript,Adddate,IP) values(" & ChannelID & "," & InfoID & ",'" & UserName & "',"& InOrOutFlag & "," & Point & ",1,'" & replace(User,"'","""") & "','" & replace(Descript,"'","""") & "'," & SqlNowString & ",'" & replace(getip,"'","""") & "')")
		IF Err Then PointInOrOut=false Else PointInOrOut=true
	End Function
	
	'会员有效期明细出入函数
	'参数:UserName,InOrOutFlag,Edays,User,Descript
	Function EdaysInOrOut(UserName,InOrOutFlag,Edays,User,Descript)
		 If Not IsNumeric(InOrOutFlag) Or Not IsNumeric(Edays) Then EdaysInOrOut=false:Exit Function
		 Conn.Execute("insert into KS_LogEdays(UserName,InOrOutFlag,Edays,[user],descript,adddate,ip) values('" & UserName & "'," & InOrOutFlag & "," & Edays & ",'" & user & "','" & replace(descript,"'","""") & "'," & SqlNowString & ",'" & getip & "')")
		 IF Err Then EdaysInOrOut=false Else EdaysInOrOut=true
	 End Function
	 
	'发送站内信息
	'参数Incept--接收者,Sender-发送者,title--主题,Content--信件内容
	Public Sub SendInfo(Incept,Sender,title,Content)
	     Conn.Execute("insert Into KS_Message(Incept,Sender,Title,Content,SendTime,Flag,IsSend,DelR,DelS) values('" & Incept & "','" & Sender & "','" & replace(Title,"'","""") & "','" & replace(Content,"'","""") & "'," & SqlNowString & ",0,1,0,0)")
	End Sub
	'======================================================================================
	
	
	
	'**************************************************
	'函数名：GetSiteOnline
	'作  用：显示在线人数（总在线：1人，用户：1人，游客：0人）
	'**************************************************
	Public Sub GetSiteOnline()
	    If WSetting(0)=1 Then
	       'Response.Expires = 0
		   On Error Resume Next
		   Dim strUserName,strReferer,remoteaddr,platform,BrowserType,CurrentStation
		   Dim UserSessionID,strSQL,rsOnline,OnlineSQL
		   '删除不活动的用户
		   If DataBaseType = 1 Then
		      Conn.Execute("DELETE FROM KS_Online WHERE DateDIff(s,lastTime,GetDate()) > " & CLng(Setting(8)) & " * 60")
		   Else
		      Conn.Execute("DELETE from KS_Online where DateDIff('s',lastTime,Now()) > " & CLng(Setting(8)) & " * 60")
		   End If
		   '写入用户统计
		   Application.Lock
		   IF Cbool(KSUser.UserLoginChecked)=True Then
		      strUserName=KSUser.UserName
		   Else
		      strUserName="匿名用户"
		   End if
		   strReferer = CheckInSQL(URLDecode(Request.ServerVariables("HTTP_REFERER")))'用来获取(从哪个页面转到当前页面的)
		   If strReferer = Empty Then
		      strReferer = "★直接输入或书签导入★"
		   Else
		      strReferer = Left(strReferer,220)
		   End if
		   remoteaddr=Request.ServerVariables("HTTP_X_UP_CALLING_LINE_ID")
		   If remoteaddr="" Then remoteaddr=GetIP
		   BrowserType = CheckInSQL(Request.ServerVariables("HTTP_USER_AGENT"))
		   Platform = "WAP浏览器"
		   '识别搜索引擎
		   Dim BotList, i
		   BotList = "timewe,Twiceler,roboo,google,baidu,yahoo,msn"
		   BotList = Split(BotList, ",")
		   For i = 0 To UBound(BotList)
			   If InStr(BrowserType, BotList(i)) > 0 Then
				  Platform = BotList(i) & "搜索器"
				  Exit For
			   End If
		   Next
		   'If BrowserType<>"" Then BrowserType = Split(BrowserType, " ")(0)'操作系统
		   CurrentStation = CheckInSQL(Left(Request.ServerVariables("HTTP_URL"),255))'访问的文件路径 
		   '写入访问来源详细地址
		   UserSessionID = Session.Sessionid
		   strSQL = "SELECT * FROM [KS_Online] WHERE IP='" & remoteaddr & "' And UserName='" & strUserName & "' Or ID=" & UserSessionID
		   Set rsOnline = Server.CreateObject("ADODB.Recordset")
		   rsOnline.Open strSQL,Conn,1,1
		   If rsOnline.BOF And rsOnline.EOF Then
		      OnlineSQL = "INSERT INTO KS_Online(id,UserName,station,ip,Browser,startTime,LastTime,strReferer) VALUES (" & UserSessionID & ",'" & strUserName & "','" &CurrentStation & "','" & remoteaddr & "','" & Platform & "|" & BrowserType & "|ON'," & SqlNowString & "," & SqlNowString & ",'" & strReferer & "')"
		      Call AddCountData(BrowserType)'写入流量统计
		   Else
		      OnlineSQL = "UPDATE KS_Online SET ID="&UserSessionID&",UserName='"&strUsername&"',station='"&CurrentStation&"',LastTime="&SqlNowString&" WHERE ID = "&UserSessionID
		      Call UpdateCountData(BrowserType)'写入流量统计
		   End If
		   Conn.Execute(OnlineSQL)
		   rsOnline.close:Set rsOnline = Nothing
		   Application.UnLock
		End If
	End Sub
    '**************************************************
	'函数名：AddCountData
	'作  用：写入流量统计
	'**************************************************
	Sub AddCountData(BrowserType)
	    Dim strSQL,ORS
		Dim rowname:rowname = GetSearcher(BrowserType)
        If DataBaseType = 1 Then
		   strSQL = "SELECT id FROM [KS_SiteCount] WHERE Datediff(d,CountDate,GetDate())=0"
		Else
		   strSQL = "SELECT id FROM [KS_SiteCount] WHERE Datediff('d',CountDate,Now())=0"
		End If
		Set ORS = Server.CreateObject("ADODB.Recordset")
		ORS.Open strSQL,Conn,1,1
		If ORS.BOF And ORS.EOF Then
		   strSQL = "INSERT INTO KS_SiteCount(UniqueIP,Pageview,CountDate," & rowname & ") VALUES (1,1," & SqlNowString & ",1)"
		Else
		   strSQL = "UPDATE KS_SiteCount SET UniqueIP=UniqueIP+1,Pageview=Pageview+1," & rowname & "=" & rowname & "+1 WHERE ID=" & ORS("ID")
		End If
		ORS.Close:Set ORS = Nothing
		Conn.Execute(strSQL)
		strSQL = Empty
	End Sub
    '**************************************************
	'函数名：UpdateCountData
	'作  用：写入流量统计
	'**************************************************
	Sub UpdateCountData(BrowserType)
	    Dim strSQL,ORS
		Dim rowname:rowname = GetSearcher(BrowserType)
	    If DataBaseType = 1 Then
		   strSQL = "SELECT id FROM [KS_SiteCount] WHERE Datediff(d,CountDate,GetDate())=0"
		Else
		   strSQL = "SELECT id FROM [KS_SiteCount] WHERE Datediff('d',CountDate,Now())=0"
	    End If
	    Set ORS = Server.CreateObject("ADODB.Recordset")
		ORS.Open strSQL,Conn,1,1
		If ORS.BOF And ORS.EOF Then
		   strSQL = "INSERT INTO KS_SiteCount(UniqueIP,Pageview,CountDate," & rowname & ") VALUES (1,1," & SqlNowString & ",1)"
		Else
		   strSQL = "UPDATE KS_SiteCount SET Pageview=Pageview+1 WHERE ID=" & ORS("ID")
		End If
		ORS.Close:Set ORS = Nothing
		Conn.Execute(strSQL)
		strSQL = Empty
	End Sub
	
	Function GetSearcher(ByVal strUrl)
	    'On Error Resume Next
		If Len(strUrl) < 5 Then
		   GetSearcher = "DirectInput"
		   Exit Function
		End If
		If InStr(strUrl, "http://") = 0 Then
		   GetSearcher = "DirectInput"
		   Exit Function
		End If
		Dim Searchlist,i,SearchName
		Searchlist = "google,baidu,yahoo,3721,zhongsou,sogou"
		Searchlist = Split(Searchlist, ",")
		For i = 0 To UBound(Searchlist)
			If InStr(strUrl, Searchlist(i)) > 0 Then
			   SearchName = Searchlist(i)
			   Exit For
			Else
			   SearchName = "other"
			End If
		Next
		GetSearcher = SearchName
	End Function
	
	Function CheckInSQL(str)
	    If IsNull(str) Then Exit Function
		On Error Resume Next
		Dim s,Badstring,i
		Badstring = " and | mid |exec|insert|select|delete|update|count|master|truncate|char|declare"
		str = Replace(str, Chr(0), "")
		str = Replace(str, Chr(9), " ")
		str = Replace(str, Chr(255), " ")
		str = Replace(str, "　", " ")
		str = Replace(str, "'", "''")
		str = Replace(str, "--", "－－")
		str = Replace(str, "@", "＠")
		str = Replace(str, "*", "＊")
		str = Replace(str, "%", "％")
		str = Replace(str, "^", "＾")
		Badstring = Split(Badstring, "|")
		s = LCase(str)
		s = Replace(s, Chr(10), "")
		s = Replace(s, Chr(13), "")
		For i = 0 To UBound(Badstring)
		    If InStr(s, Badstring(i))>0 Then
			   CheckInSQL = ""
			Exit Function
			End If
		Next
		CheckInSQL = str
	End Function
	
	'过滤非法字符
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
End Class
%>


