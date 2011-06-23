<!--#include file="Kesion.Label.FunctionCls.asp"-->
<!--#include file="Kesion.Label.SQLCls.asp"-->
<!--#include file="Kesion.Label.BaseFunCls.asp"-->
<!--#include file="Kesion.FsoVarCls.asp"-->
<%
Class Refresh
        Private KS,KSLabel,DomainStr,WapValue
		Private PerPageNumber
		Private Sub Class_Initialize()
            Set KS=New PublicCls
			Set KSLabel=New RefreshFunction
		    DomainStr=KS.GetDomain
			WapValue=KS.WapValue
		End Sub
        Private Sub Class_Terminate()
		    Set KS=Nothing
			Set KSLabel=Nothing
		End Sub

		'*******************************************************************************************************
		'函数名：KSLabelReplaceAll
		'作  用：替换所有标签
		'参  数：F_C 模板内容
		'返回值：替换过的模板内容
		'********************************************************************************************************
		Public Function KSLabelReplaceAll(F_C)
		    On Error Resume Next
			Dim DCls:Set Dcls = New DIYCls
			F_C = DCls.ReplaceUserFunctionLabel(F_C)'替换自定义函数标签
			Set DCls = Nothing			
			F_C = ReplaceGeneralLabelContent(F_C)
			F_C = ReplaceLableFlag(F_C)
			KSLabelReplaceAll=F_C
	    End Function	

		'**************************************************
		'函数名：LoadTemplate
		'作  用：取出模板内容
		'参  数：TemplateFname模板地址
		'返回值：模板内容
		'**************************************************
		Function LoadTemplate(TemplateFname)
			
			
			  on error resume next
			  TemplateFname=Replace(TemplateFname,"{@TemplateDir}",KS.Setting(3) & KS.Setting(90))
			  TemplateFname =GetMapPath & Replace(TemplateFname, "//", "/")
				dim str,stm
				set stm=server.CreateObject("adodb.stream")
				stm.Type=2 '以本模式读取
				stm.mode=3 
				stm.charset=WapCharset
				stm.open
				stm.loadfromfile TemplateFname
				str=stm.readtext
				stm.Close
				set stm=nothing
				if err then
				LoadTemplate="请绑定模板!"
				else
				LoadTemplate=str
				End if
			    LoadTemplate=Replace(LoadTemplate,"{$UID}",Request("Uid"))
			    LoadTemplate=LoadTemplate & Published

			
		End Function

		'**************************************************
		'函数名：ReplaceLableFlag
		'作  用：去除标签{$},并分组以将标签参数用","隔开
		'          示例: km=ReplaceLableFlag("{$Test("par1","par2","par3")}")
		'          结果     km=Test,Par1,Par2,Par3
		'参  数： Content  ----待替换内容
		'返回值：返回用","隔开的字符串
		'**************************************************
		Function ReplaceLableFlag(Content)
			Dim regEx, Matches, Match, TempStr
			Set regEx = New RegExp
			regEx.Pattern = "{\$[^{\$}]*}"
			regEx.IgnoreCase = True
			regEx.Global = True
			Set Matches = regEx.Execute(Content)
			ReplaceLableFlag = Content
			For Each Match In Matches
				On Error Resume Next
				TempStr = Match.Value
				TempStr = Replace(TempStr, Chr(13) & Chr(10), "")
				TempStr = Replace(TempStr, "{$", "")
				TempStr = Replace(TempStr, "}", "")
				TempStr = Left(TempStr, InStr(TempStr, "(") - 1) & "," & Mid(TempStr, InStr(TempStr, "(") + 1)
				TempStr = Left(TempStr, InStrRev(TempStr, ")") - 1)
				TempStr = Replace(TempStr, """", "")
			   If Err.Number = 0 Then
				ReplaceLableFlag = Replace(ReplaceLableFlag, Match.Value, ChangeLableToFunction(TempStr))
			   End If
			Next
		End Function
		
		'**************************************************
		'函数名：GetFunctionLabelParam
		'作  用：取得标签的参数，用“，”隔开 如 {=GetFlashByPlayer(100,50)},返回100,50
		'参  数：Content--查找的内容，MatchStr--前缀匹配字符串
		'返回值：返回用","隔开的字符串参数
		'**************************************************
		Public Function GetFunctionLabelParam(Content, MatchStr)
	        GetFunctionLabelParam = Replace(Content, MatchStr & "(", "")
			GetFunctionLabelParam = Replace(Replace(GetFunctionLabelParam, ")}", ""), """", "")
	    End Function
		
		'**************************************************
		'函数名：GetFunctionLabel
		'作  用：取得函数标签 如 sssssss{=GetFlashByPlayer(100,50)}sssss,返回{=GetFlashByPlayer(100,50)}
		'参数：Content--查找的内容，MatchStr--前缀匹配字符串
		'返回值：函数标签
		'**************************************************
		Public Function GetFunctionLabel(Content, MatchStr)
		    Dim regEx, Matches, Match,N
			Set regEx = New RegExp
			regEx.Pattern = MatchStr & "[^{\=}]*}"
			regEx.IgnoreCase = True
			regEx.Global = True
			Set Matches = regEx.Execute(Content)
			GetFunctionLabel = ""
			For Each Match In Matches
			    On Error Resume Next
				N=N+1
				IF N=1 Then
				   GetFunctionLabel = Match.Value
				Else
				   GetFunctionLabel=GetFunctionLabel & "@@@" & Match.Value
				End IF
			Next
		End Function

		'**************************************************
		'函数名：GetTopUserLogin
		'作  用：显示会员登录入口(横排)
		'**************************************************
		Function GetTopUserLogin()
		    Dim TempStr
		    If Cbool(KSUser.UserLoginChecked)=True Then
			   Dim MyMailTotal:MyMailTotal=Conn.Execute("Select Count(ID) From KS_Message Where Incept='"&KSUser.UserName&"'And Flag=0 and IsSend=1 and delR=0")(0)
			   TempStr = TempStr &"<a href="""&DomainStr&"User/User_Message.asp?Action=inbox&"&WapValue&""">收信箱 "&MyMailTotal&"</a> <a href="""&DomainStr&"User/Index.asp?"&WapValue&""">会员中心</a> "
			Else
			   TempStr = TempStr &"<a href="""&DomainStr&"User/Login/"">会员登陆</a> <a href="""&DomainStr&"User/Reg/"">会员注册</a>"
			End if
			GetTopUserLogin = TempStr
		End Function

		'**************************************************
		'函数名：GetUserLogin
		'作  用：显示会员登录入口(竖排)
		'**************************************************
		Function GetUserLogin()
		    Dim TempStr
		    If KSUser.UserLoginChecked()=True Then
			   Dim MyMailTotal:MyMailTotal=Conn.Execute("Select Count(ID) From KS_Message Where Incept='" & KSUser.UserName &"' And Flag=0 and IsSend=1 and delR=0")(0)
			   TempStr = TempStr & KSUser.RealName
			   If (Hour(Now) < 6) Then
			      TempStr = TempStr & "凌晨好!<br/>"
			   ElseIf (Hour(Now) < 9) Then
			      TempStr = TempStr & "早上好!<br/>"
			   ElseIf (Hour(Now) < 12) Then
			      TempStr = TempStr & "上午好!<br/>"
			   ElseIf (Hour(Now) < 14) Then
			      TempStr = TempStr &"中午好!<br/>"
			   ElseIf (Hour(Now) < 17) Then
			      TempStr = TempStr & "下午好!<br/>"
			   ElseIf (Hour(Now) < 18) Then
			      TempStr = TempStr & "傍晚好!<br/>"
			   Else
			      TempStr = TempStr & "晚上好!<br/>"
			   End If
			   TempStr = TempStr &"<a href="""&DomainStr&"User/User_Message.asp?Action=inbox&"&WapValue&""">收信箱 "&MyMailTotal&"</a> <a href="""&DomainStr&"User/Index.asp?"&WapValue&""">会员中心</a> "
			Else
			   TempStr = TempStr &"<a href="""&DomainStr&"User/Login/"">会员登陆</a> <a href="""&DomainStr&"User/Reg/"">会员注册</a>"
			End if
			GetUserLogin = TempStr
		End Function
		
		'Tags通用标签
		Function GetTags(TagType,Num)
		    If not Isnumeric(Num) Then Exit Function
			Dim sqlstr,SQL,i,n,str
			Select Case cint(TagType)
			    Case 1:sqlstr="select top 500 keytext,hits from KS_Keywords order by hits desc"
				Case 2:sqlstr="select top 500 keytext,hits from KS_Keywords order by lastusetime desc,id desc"
				Case 3:sqlstr="select top 500 keytext,hits from KS_Keywords order by Adddate desc,id desc"
				Case Else 
				GetTags="":Exit Function
		    End Select
		    Dim RS:set RS=Conn.Execute(sqlstr)
		    If RS.EOF Then RS.Close:set RS=Nothing:Exit Function
		    SQL=RS.Getrows(-1)
		    RS.Close:set RS=Nothing
			For i=0 To Ubound(sql,2)
			    If KS.FoundInArr(str,SQL(0,i),",")=False Then
				   n=n+1
				   str=str & "," & SQL(0,i)
				   GetTags=GetTags & "<a href=""" & DomainStr & "Plus/Tags.asp?n=" & server.URLEncode(SQL(0,i))& "&amp;"&WapValue&""">" & SQL(0,i) & "</a> "
				End If
				If n>=Cint(Num) Then Exit For
		    Next
		End Function
		
		'Tags通用标签
		Function ReplaceKeyTags(KeyStr)
		    On Error Resume Next
		    Dim I,K_Arr:K_Arr=Split(KeyStr," ")
		    For I=0 To Ubound(K_Arr)
		        ReplaceKeyTags=ReplaceKeyTags & "<a href=""" & DomainStr & "Plus/Tags.asp?n=" & server.URLEncode(K_Arr(i)) & "&amp;"&WapValue&""">" & K_Arr(i) & "</a> "
		    Next
		    If Err Then ReplaceKeyTags="":Err.Clear
		End Function
		
		'=================================================
		'函数名：GetWenhouyu
		'作  用：显示不同时段不同的问候语
		'参  数：无
		'=================================================
		Function GetWenhouyu(str1,str2,str3,str4,str5,str6,str7,str8)
		     Dim MyTime,MyHour
		     MyTime=Now
			 MyHour=hour(MyTime)
			 If MyHour<6 or MyHour>=22 Then GetWenhouyu = str1
			 If MyHour>=6 And MyHour<=9 Then GetWenhouyu = str2
			 If MyHour>=9 And MyHour<=12 Then GetWenhouyu = str3
			 If MyHour>=12 And MyHour<=14 Then GetWenhouyu = str4
			 If MyHour>=14 And MyHour<=17 Then GetWenhouyu = str5
			 If MyHour>=17 And MyHour<=18 Then GetWenhouyu = str6
			 If MyHour>=18 And MyHour<=19 Then GetWenhouyu = str7
			 If MyHour>=19 And MyHour<=22 Then GetWenhouyu = str8
		End Function

		'*********************************************************************************************************
		'函数名：GetSiteCountAll
		'作  用：替换网站统计标签为内容
		'参  数：Flag-0总统计，1-文章统计 2-图片统计
		'*********************************************************************************************************
		Function GetSiteCountAll()
		    Dim ChannelTotal: ChannelTotal = Conn.Execute("Select Count(*) From KS_Class Where TN='0'")(0)
			Dim MemberTotal:MemberTotal=Conn.Execute("Select Count(*) From KS_User")(0)
			Dim CommentTotal: CommentTotal = Conn.Execute("Select Count(*) From KS_Comment")(0)
			Dim GuestBookTotal:GuestBookTotal=Conn.Execute("Select Count(ID) From KS_GuestBook")(0)
			GetSiteCountAll = GetSiteCountAll & "频道总数： " & ChannelTotal & " 个<br/>"
			dim rsc:set rsc=conn.execute("select ChannelID,ItemName,Itemunit,ChannelTable from KS_Channel where ChannelStatus=1 And ChannelID<>6 And ChannelID<>9")
			dim k,sql:sql=rsc.getrows(-1)
			rsc.close:set rsc=nothing
			For k=0 To ubound(sql,2)
			    GetSiteCountAll = GetSiteCountAll & sql(1,k) & "总数： " & Conn.Execute("Select Count(id) From " & sql(3,k))(0) & " " & sql(2,k)&"<br/>"
			Next
			GetSiteCountAll = GetSiteCountAll & "注册会员： " & MemberTotal & " 位<br/>"
			GetSiteCountAll = GetSiteCountAll & "留言总数： " & GuestBookTotal &" 条<br/>"
			GetSiteCountAll = GetSiteCountAll & "评论总数： " & CommentTotal & " 条<br/>"
			GetSiteCountAll = GetSiteCountAll
		End Function
		
		'**************************************************
		'函数名：GetIndexChannel
		'作  用：取出首页频道导航
		'参  数：Num--一行栏目数量,Cent--分隔符号(支持WML语言)
		'**************************************************
		Function GetIndexChannel(Num,Cent)
		    Dim RS,SQL,TempStr,I,N
			Set RS=Conn.Execute("select a.ClassID,a.FolderName,a.ChannelID from KS_Class a inner join KS_Channel B on A.ChannelID=B.ChannelID where a.TN='0' And b.ChannelStatus=1 and B.WapSwitch=1 And a.WapSwitch=1 order by a.root,a.FolderOrder")
			If RS.EOF Then
			   TempStr = "您还没有频道栏目！<br/>"
			Else
			   N=1
			   Do while Not RS.EOF
			      TempStr = TempStr &Cent &"<a href=""" & KS.GetFolderPath(RS("ChannelID"),RS("ClassID")) & """>"&RS("FolderName")&"</a>"
				  If N Mod Num=0 or N=RS.Recordcount Then
			         TempStr = TempStr &"<br/>"
				  End If
				  N=N+1
				  RS.Movenext
			   loop
			End If
			Set RS=Nothing
			GetIndexChannel=TempStr
		End Function

		'**************************************************
		'函数名：GetIndexSearch
		'作  用：取出首页搜索
		'**************************************************
		Function GetIndexSearch()
		    Dim RS,TempStr,N
		    TempStr = "<input name=""keyword"" type=""text"" size=""20"" value=""""/><br/>"
			TempStr = TempStr & "<select name=""channelid"">"
				 If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
					Dim ModelXML,Node
					Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
					For Each Node In ModelXML.documentElement.SelectNodes("channel")
					 if Node.SelectSingleNode("@ks21").text="1" and Node.SelectSingleNode("@ks49").text="1" and Node.SelectSingleNode("@ks6").text<6 Then
					  
					   TempStr=TempStr & "<option value='" & Node.SelectSingleNode("@ks0").text & "'>" &Node.SelectSingleNode("@ks1").text & "</option>"
					End If
				next 
             TempStr = TempStr &"</select>"
			  TempStr = TempStr &"<anchor>搜索<go href="""&DomainStr&"Plus/Search.asp?searchtype=1&"&WapValue&""" method=""get""><postfield name=""keyword"" value=""$(keyword)""/><postfield name=""ChannelID"" value=""$(channelid)""/></go></anchor>"
			GetIndexSearch=TempStr
		End Function
		
		'**************************************************
		'函数名：GetIndexList
		'作  用：取出首页"热门,最新,推荐,随机"列表
		'参  数：channelid --模型ID
		'       strID--栏目ID
		'       strType--热门1,最新2,推荐3,随机4
		'       strHead--头导航类型
		'       strTail--尾导航类型
		'       strNum--显示记录数
		'       strTitleNum--链接标题字符
		'**************************************************
		Function GetIndexList(ChannelID,strID,strType,strHead,strTail,strNum,strTitleNum)
			Dim RS,SQL,I,Node,XML
			Dim SqlStr,TempStr
			If strID="" Then strID="0"
			If strID="-1" Then strID=FCls.RefreshFolderID
			ChannelID=KS.ChkCLng(ChannelID)
			
			If ChannelID=0 Then
			   SQLStr="Select Top " & StrNum & " InfoID as ID,Tid,Title,Fname,channelID From KS_ItemInfo"
			Else
			   SQLStr="Select Top " & StrNum & " ID,Tid,Title,Fname," & ChannelID & " as channelid From " & KS.C_S(ChannelID,2)
			End If
			
			Dim Param:Param=" Where Verific=1 And DelTF=0"
			If strID<>"0" Then Param=Param & " and Tid in(" & KS.GetFolderTid(strID) & ")"
			If KS.C_S(ChannelID,6)=1 Then Param=Param & " and changes=0"
			
     		Select Case strType
			    Case 1'热门
				SqlStr= SqlStr & Param & "  order by Hits Desc,ID"
				Case 2'最新
				SqlStr= SqlStr & Param & " order by AddDate desc,ID"
				Case 3'推荐
				SqlStr= SqlStr & Param & " And Recommend=1 order by ID desc"
				Case 4'随机
				If DataBaseType=0 then
				   Randomize()
				   SqlStr= SqlStr & Param & " order by Rnd("&-1*(Int(1000*Rnd)+1)&"*ID)"
				Else
				   SqlStr= SqlStr & Param & " order by NewID()"
				End if
			End Select	
			
			
			set RS=Conn.Execute(SqlStr)
			If Not RS.Eof Then
			  Set XML=KS.RsToXml(rs,"row","")
			End If
			RS.Close:Set RS=Nothing
			If Not IsObject(XML) Then
			  TempStr = "没有内容！<br/>"
			Else
			   Dim strReplace
			   For Each Node In XML.DocumentElement.SelectNodes("row")
			       i=I+1
				   strReplace = Replace(strHead,"[ClassName]",KS.GetClassNP(Node.SelectSingleNode("@tid").text))
				   strReplace = Replace(strReplace,"[AutoID]",I)
			       TempStr = TempStr&strReplace&"<a href=""" & KS.GetInfoUrl(Node.SelectSingleNode("@channelid").text,Node.SelectSingleNode("@id").text,Node.SelectSingleNode("@fname").text) & """>"&KS.GotTopic(Node.SelectSingleNode("@title").text,strTitleNum)&"</a>" & strTail & ""
			   Next
			End if
			GetIndexList=TempStr
		End Function
		
		'调用小论坛帖子列表
		Function GetClubList(boardid,strType,strHead,strTail,Num,TitleLen)
		   Dim SqlStr,Node,Xml,RS,Str
		   SqlStr="select top " & Num & " id,subject,username,AddTime from KS_GuestBook Where verific=1"
		   If KS.ChkClng(BoardID)<>0 Then
		   SqlStr=SqlStr & " and boardid=" & boardid
		   End If
     		Select Case KS.ChkClng(strType)
			    Case 1'热门
				SqlStr= SqlStr & " order by Hits Desc,ID"
				Case 2'最新
				SqlStr= SqlStr & " order by AddTime desc,ID"
				Case 3'精华
				SqlStr= SqlStr & " IsBest=1 order by ID desc"
				Case 4'随机
				If DataBaseType=0 then
				   Randomize()
				   SqlStr= SqlStr & " order by Rnd("&-1*(Int(1000*Rnd)+1)&"*ID)"
				Else
				   SqlStr= SqlStr & " order by NewID()"
				End if
			End Select	
			Set RS=Conn.Execute(SQLStr)
			If Not RS.Eof Then
			  Set XML=KS.RsToXml(RS,"row","")
			End If
			RS.Close : Set RS=Nothing
			If IsObject(XML) Then
			 str=""
			 For Each Node In XML.DocumentElement.SelectNodes("row")
			   str=str & strHead & "<a href=""" & DomainStr & "club/display.asp?id=" & Node.SelectSingleNode("@id").text & "&amp;" & KS.WapValue & """>" & KS.Gottopic(Node.SelectSingleNode("@subject").text,titlelen) & "(" & FormatDateTime(Node.SelectSingleNode("@addtime").text,2) & ")</a>" & strTail
			 Next   
			End If
			GetClubList=str
		End Function
		
		'调用日志列表
		Function GetLogList(TypeID,strType,strHead,strTail,Num,TitleLen)
		   Dim SqlStr,Node,Xml,RS,Str
		   SqlStr="select top " & Num & " id,title,username,AddDate from KS_BlogInfo Where status=0"
		   If KS.ChkClng(TypeID)<>0 Then
		   SqlStr=SqlStr & " and TypeID=" & TypeID
		   End If
     		Select Case KS.ChkClng(strType)
			    Case 1'热门
				SqlStr= SqlStr & " order by Hits Desc,ID"
				Case 2'最新
				SqlStr= SqlStr & " order by AddDate desc,ID"
				Case 3'精华
				SqlStr= SqlStr & " Best=1 order by ID desc"
				Case 4'随机
				If DataBaseType=0 then
				   Randomize()
				   SqlStr= SqlStr & " order by Rnd("&-1*(Int(1000*Rnd)+1)&"*ID)"
				Else
				   SqlStr= SqlStr & " order by NewID()"
				End if
			End Select	
			Set RS=Conn.Execute(SQLStr)
			If Not RS.Eof Then
			  Set XML=KS.RsToXml(RS,"row","")
			End If
			RS.Close : Set RS=Nothing
			If IsObject(XML) Then
			 str=""
			 For Each Node In XML.DocumentElement.SelectNodes("row")
			   str=str & strHead & "<a href=""" & DomainStr & "space/list.asp?id=" & Node.SelectSingleNode("@id").text & "&amp;username=" & Node.SelectSingleNode("@username").text & "&amp;" & KS.WapValue & """>" & KS.Gottopic(Node.SelectSingleNode("@title").text,titlelen) & "(" & FormatDateTime(Node.SelectSingleNode("@adddate").text,2) & ")</a>" & strTail
			 Next   
			End If
			GetLogList=str
		End Function
		
		
		
		

		'=================================================
		'函数名：GetPourAccount
		'作  用：显示倒计时
		'参  数：strTxt--说明，strTime--时间
		'=================================================
		Function GetPourAccount(strTxt,strTime)
		    GetPourAccount = strTxt & Datediff("d",Date(),strTime)&"天"
		End Function
		
		'=================================================
		'函数名：GetVote
		'作  用：显示网站调查
		'参  数：无
		'=================================================
		Function GetVote(VoteID)
		    On Error Resume Next
		    Dim sqlVote,rsVote
		    sqlVote="select Title from KS_Vote where ID=" & VoteID & ""
		    Set rsVote=Conn.Execute(sqlvote)
		    If rsVote.BOF And rsVote.EOF Then 
		       GetVote="没有任何调查!"
		    Else
		       GetVote="<a href="""&DomainStr&"Plus/Vote.asp?Action=Show&ID="&VoteID&"&"&WapValue&""">"&rsVote("Title")&"</a>"
		    End If
		    rsVote.close:set rsVote=nothing
	    End Function
		
		'显示会员排行
		Function GetTopUser(Num,MoreStr)
		    Dim SQL,I,RSObj
			Set RSObj=Conn.Execute("Select Top "&Num&" UserID,UserName,RealName,LoginTimes From KS_User Order BY LoginTimes Desc,UserID Desc")
			SQL = RSObj.GetRows(-1)
			RSObj.Close : Set RSObj = Nothing
			For I = 0 To UBound(SQL,2)
			    GetTopUser = GetTopUser & I+1 & ".<a href="""&DomainStr&"User/ShowUser.asp?Keyword="&SQL(1,I)&"&"&WapValue&""">" & KS.GetUserRealName(SQL(1,i)) & "</a>"&SQL(3,I)&"<br/>"
			Next
			GetTopUser=GetTopUser & "<a href="""&DomainStr&"User/UserList.asp?"&WapValue&""">"&MoreStr&"</a>"
		End Function
		
		'=================================================
		'函数名：GetAnnounce
		'作  用：显示网站公告
		'参  数：无
		'=================================================
		Function GetAnnounce(AnnounceID)
		    On Error Resume Next
			Dim sqlAnnounce,rsAnnounce
			sqlAnnounce="select Title from KS_Announce where ID=" & AnnounceID & ""
			Set rsAnnounce=Conn.Execute(sqlAnnounce)
			If rsAnnounce.bof And rsAnnounce.eof Then 
			   GetAnnounce="没有任何公告!"
			Else
			   GetAnnounce="<a href="""&DomainStr&"Plus/Announce.asp?AnnounceID="&AnnounceID&"&"&WapValue&""">"&rsAnnounce("Title")&"</a>"
			End If
			rsAnnounce.close:set rsAnnounce=nothing
		End Function
		
		'**************************************************
		'函数名：GetClassList
		'作  用：取出频道栏目列表
		'参  数：Num--一行栏目数量,Cent--分隔符号(支持WML语言)
		'**************************************************
		Function GetClassList(Num,Cent)
		    Dim RS,TempStr,N
		    set RS=server.createobject("ADODB.Recordset")
			RS.Open "select ClassID,FolderName from KS_Class where TN='"&FCls.RefreshFolderID&"' order by FolderOrder Asc",Conn,1,1
			If RS.EOF Then
			   TempStr = "没有频道分类！<br/>"
			Else
			   N=1
			   Do while not RS.EOF
			      TempStr = TempStr &"<a href=""" & KS.GetFolderPath(FCls.ChannelID,RS("ClassID")) & """>"&RS("FolderName")&"</a>"
				  If N Mod Num=0 or N=RS.Recordcount Then
				     TempStr = TempStr &"<br/>"
				  Else
				     TempStr = TempStr &Cent
				  End If
				  N=N+1
				  RS.Movenext
			   Loop
			End If
			RS.Close:set RS=nothing
			GetClassList = TempStr
	    End Function
		
		'**************************************************
		'函数名：GetShowClassList
		'作  用：取出栏目"热门,最新,推荐,随机"列表
		'参  数：strType--热门1,最新2,推荐3,随机4
		'       strHead--头导航类型
		'       strTail--尾导航类型
		'       strNum--显示记录数
		'       strTitleNum--链接标题字符
		'**************************************************
		Function GetShowClassList(strType,strHead,strTail,strNum,strTitleNum)
		    On Error Resume Next
			Dim RS,SqlStr,TempStr,I
			Dim ChannelID:ChannelID=FCls.ChannelID
			Dim FolderID:FolderID=FCls.RefreshFolderID
			Dim Param
			Select Case ChannelID
			    Case 1:Param="Select top "&strNum&" ID,Tid,Title,Fname,Changes From " & KS.C_S(ChannelID,2) & ""
				Case 2:Param="Select top "&strNum&" ID,Tid,Title,Fname,0 From " & KS.C_S(ChannelID,2) & ""
				Case 3:Param="Select top "&strNum&" ID,Tid,Title,Fname,0 From " & KS.C_S(ChannelID,2) & ""
				Case 4:Param="Select top "&strNum&" ID,Tid,Title,Fname,0 From KS_Flash"
				Case 5:Param="Select top "&strNum&" ID,Tid,Title,Fname,0 From KS_Product"
				Case 7:Param="Select top "&strNum&" ID,Tid,Title,Fname,0 From KS_Movie"
				Case 8:Param="Select top "&strNum&" ID,Tid,Title,Fname,0 From KS_GQ"
				Case Else
				Exit Function
			End Select
			Select Case strType
			    Case 1'热门
				SqlStr="" & Param & " where Tid In ("&KS.GetFolderTid(FolderID)&") And Verific=1 And DelTF=0 order by Hits Desc"
				Case 2'最新
				SqlStr="" & Param & " where Tid In ("&KS.GetFolderTid(FolderID)&") And Verific=1 order by AddDate desc"
				Case 3'推荐
				SqlStr="" & Param & " where Tid In ("&KS.GetFolderTid(FolderID)&") And Verific=1 And Recommend=1 order by ID desc"
				Case 4'随机
				If DataBaseType=0 then
				   Randomize()
				   SqlStr="" & Param & " where Tid In ("&KS.GetFolderTid(FolderID)&") And Verific=1 order by Rnd("&-1*(Int(1000*Rnd)+1)&"*ID)"
				Else
				   SqlStr="" & Param & " where Tid In ("&KS.GetFolderTid(FolderID)&") And Verific=1 order by newid()"
				End if
			End Select
			set RS=Conn.Execute(SqlStr)
			If RS.BOF And RS.EOF Then
			   TempStr = "没有内容！<br/>"
			Else
			   Dim strReplace
			   For I=1 To strNum
				   strReplace = Replace(strHead,"[ClassName]",KS.GetClassNP(RS("Tid")))
				   strReplace = Replace(strReplace,"[AutoID]",I)
			       TempStr = TempStr&strReplace&"<a href=""" & KS.GetInfoUrl(ChannelID,RS(0),RS(3)) & """>"&KS.GotTopic(RS("Title"),strTitleNum)&"</a>"&strTail&""
			       RS.MoveNext
			   Next
			End if
			RS.Close:set RS=Nothing
			GetShowClassList = TempStr
		End Function
		
		'**************************************************
		'函数名：GetRandomPhotoText
		'作  用：取出随机图文
		'参  数：width--宽度设置,height--高度设置
		'       strName--是否显示标题True显示False不显示
		'       strTitleNum--链接标题字符
		'       strNum--显示记录数
		'       strLineNum--一行显示数
		'       strAttach--显示附加内容True显示False不显示
		'**************************************************
		Function GetRandomPhotoText(width,height,strName,strTitleNum,strNum,strLineNum,strAttach)
		    'On Error Resume Next
			Dim ChannelID:ChannelID=FCls.ChannelID
			Dim FolderID:FolderID=FCls.RefreshFolderID
			Randomize()
			Dim Preface,Param,RS,SqlStr,I,TempStr
			If DataBaseType=0 then
			   Preface="order by Rnd("&-1*(Int(1000*Rnd)+1)&"*ID)"
			Else
			   Preface="order by NewID()"
			End if
			Param=" where verific=1 And Tid In ("&KS.GetFolderTid(FolderID)&") "&Preface&""
			Select Case KS.C_S(ChannelID,6)
			    Case 1:SqlStr="select top "&strNum&" ID,Title,PhotoUrl,Hits from "&KS.C_S(ChannelID,2)&" " & Param &""
				Case 2:SqlStr="select top "&strNum&" ID,Title,PhotoUrl,Hits from "&KS.C_S(ChannelID,2)&" " & Param &""
				Case 3:SqlStr="select top "&strNum&" ID,Title,PhotoUrl,Hits,DownPT from "&KS.C_S(ChannelID,2)&" " & Param &""
				Case 5:SqlStr="select top "&strNum&" ID,Title,PhotoUrl,Unit,Price_Market,Price,Price_Member,Point from "&KS.C_S(ChannelID,2)&" " & Param &""
			End Select
			set RS=Conn.Execute(SqlStr)
			If RS.BOF And RS.EOF Then
			   TempStr = "没有内容！<br/>"
			Else
			   I=1
			   do while not RS.EOF
		   Select Case KS.C_S(ChannelID,6)
		       Case 1
			      If KS.BusinessVersion = 1 Then
					 TempStr = TempStr&"<a href=""Show.asp?ID="&RS("ID")&"&ChannelID="&ChannelID&"&"&WapValue&"""><img src=""JpegMore.asp?JpegSize=" & width & "x" & height & "&amp;JpegUrl=" & Rs("PhotoUrl") & """ alt=""""/></a>"
				  Else
				     TempStr = TempStr&"<a href=""Show.asp?ID="&RS("ID")&"&ChannelID="&ChannelID&"&"&WapValue&"""><img src=""" & Rs("PhotoUrl") & """ width=""" & width & """ height=""" & height &""" alt="".""/></a>"
				  End if
				  If Cbool(strName)=True Then TempStr = TempStr & "<br/><a href=""Show.asp?ID="&RS("ID")&"&ChannelID="&ChannelID&"&"&WapValue&""">"&KS.GotTopic(Rs("Title"),strLineNum)&"</a>"
				  If Cbool(strAttach)=True Then TempStr = TempStr &"(" & Rs("Hits") & ")"
			   Case 2
			      If KS.BusinessVersion = 1 Then
				     TempStr = TempStr&"<a href=""Show.asp?ID="&RS("ID")&"&ChannelID="&ChannelID&"&"&WapValue&"""><img src=""JpegMore.asp?JpegSize=" & width & "x" & height & "&amp;JpegUrl=" & Rs("PhotoUrl") & """ alt=""""/></a>"
				  Else
				     TempStr = TempStr&"<a href=""Show.asp?ID="&RS("ID")&"&ChannelID="&ChannelID&"&"&WapValue&"""><img src=""" & Rs("PhotoUrl") & """ width=""" & width & """ height=""" & height &""" alt="".""/></a>"
				  End if
				  If Cbool(strName)=True Then TempStr = TempStr & "<br/><a href=""Show.asp?ID="&RS("ID")&"&ChannelID="&ChannelID&"&"&WapValue&""">"&KS.GotTopic(Rs("Title"),strLineNum)&"</a>"
				  If Cbool(strAttach)=True Then TempStr = TempStr &"(" & Rs("Hits") & ")"
			   Case 3
			      If KS.BusinessVersion = 1 Then
				     TempStr = TempStr&"<a href=""Show.asp?ID="&RS("ID")&"&ChannelID="&ChannelID&"&"&WapValue&"""><img src=""JpegMore.asp?JpegSize=" & width & "x" & height & "&amp;JpegUrl=" & Rs("PhotoUrl") & """ alt=""""/></a>"
				  Else
				     TempStr = TempStr&"<a href=""Show.asp?ID="&RS("ID")&"&ChannelID="&ChannelID&"&"&WapValue&"""><img src=""" & Rs("PhotoUrl") & """ width=""" & width & """ height=""" & height &""" alt="".""/></a>"
				  End if
				  If Cbool(strName)=True Then TempStr = TempStr &"<br/><a href=""Show.asp?ID="&RS("ID")&"&ChannelID="&ChannelID&"&"&WapValue&""">"&KS.GotTopic(Rs("Title"),strLineNum)&"</a>"
				  If Cbool(strAttach)=True Then
				     TempStr = TempStr&"(" & Rs("Hits") & ")"
				     If Rs("DownPT")<>"" Then TempStr = TempStr & "<br/>"&KS.GotTopic(Rs("DownPT"),strLineNum)&""
				  End if
			   Case 5
			      If KS.BusinessVersion = 1 Then
				     TempStr = TempStr&"<a href=""Show.asp?ID="&RS("ID")&"&ChannelID="&ChannelID&"&"&WapValue&"""><img src=""JpegMore.asp?JpegSize=" & width & "x" & height & "&amp;JpegUrl=" & Rs("PhotoUrl") & """ alt=""""/></a>"
				  Else
				     TempStr = TempStr&"<a href=""Show.asp?ID="&RS("ID")&"&ChannelID="&ChannelID&"&"&WapValue&"""><img src=""" & Rs("PhotoUrl") & """ width=""" & width & """ height=""" & height &""" alt="".""/></a>"
				  End if
				  If Cbool(strName)=True Then TempStr = TempStr &"<br/><a href=""Show.asp?ID="&RS("ID")&"&ChannelID="&ChannelID&"&"&WapValue&""">"&KS.GotTopic(Rs("Title"),strLineNum)&"</a>"
				  If Cbool(strAttach)=True Then TempStr = TempStr & "<br/>市场:"&Rs("Price_Market")&"/"&Rs("Unit")&" 会员:"&Rs("Price_Member")&"/"&Rs("Unit")&" 购物积分:"&Rs("Point")&"/"&Rs("Unit")&""
		 End Select
				   If I Mod strLineNum=0 Or I=RS.Recordcount Then TempStr = TempStr & "<br/>"
				   
				   RS.MoveNext
				   I=I+1
			   Loop
			End IF
			RS.Close:set RS=Nothing
			GetRandomPhotoText = TempStr
		End Function

		'******************************************************************************************************
		'函数名：GetFolderNameStr
		'作  用：返回栏目顺序列表
		'参  数：NaviStr--链接字符串,RefreshFolderIDValue--栏目ID
		'返回值：形如: 科汛网络 产品列表
		'******************************************************************************************************
		Function GetFolderNaviStr()
		    On Error Resume Next
			Dim FolderID,TSArr, I
			FolderID=FCls.RefreshFolderID
			TSArr = Split(KS.C_C(FolderID,8), ",")
			GetFolderNaviStr = "<a href="""&KS.GetGoBackIndex&""">返回首页</a>"
			For I = LBound(TSArr) To UBound(TSArr) - 1
			    GetFolderNaviStr = GetFolderNaviStr & " <a href=""" & KS.GetFolderPath(FCls.ChannelID,KS.C_C(TSArr(I),9)) & """>"&KS.C_C(TSArr(I),1)&"</a>"
			Next
		End Function
		
		'显示栏目列表页返回频道首页
		Function GetGoBackChannel()
			Dim TN:TN = Conn.Execute("select top 1 TN from ks_Class where ID='" & FCls.RefreshFolderID & "'")(0)
			GetGoBackChannel = "<a href=""" & KS.GetFolderPath(FCls.ChannelID,KS.C_C(TN,9)) & """>"&KS.C_C(TN,1)&"</a>"
		End Function
		
		'显示内容页返回栏目列表页
		Function GetGoBackClass()
		    Dim FolderID:FolderID=FCls.RefreshFolderID
			GetGoBackClass = "<a href=""" & KS.GetFolderPath(FCls.ChannelID,KS.C_C(FolderID,9)) & """>"&KS.C_C(FolderID,1)&"</a>"
		End Function
		
		'**************************************************
		'函数名：GetAdvertise
		'作  用：取出随机广告
		'参  数：getplace--广告位ID
		'**************************************************
		Public Function GetAdvertise(GetPlace)
	        On Error Resume Next
			Dim AdsRS,AdsSQL
			Dim Advertvirtualvalue
			Set AdsRS=server.createobject("adodb.recordset")
			AdsRS.Open "select place from KS_ADPlace where Show_Flag=1 And Place=" & GetPlace,Conn,1,1
			If AdsRS.EOF Then
			   GetAdvertise="":AdsRS.Close:Set AdsRS=Nothing:Exit Function
			End If
			AdsRS.Close
			'每次显示广告位前，检测其中的各广告条是否过期，并更新状态
			AdsSQL="Select * from KS_Advertise where Act=1 And Class <> 0 And  Place=" & GetPlace & " order by time"
			AdsRS.Open AdsSQL,Conn,1,3
			while not AdsRS.EOF
		      Advertvirtualvalue=0
			  If AdsRS("Class")=1 Then
			     If AdsRS("Click")>=AdsRS("Clicks") Then
				    Advertvirtualvalue=1
				 End If
			  ElseIf AdsRS("Class")=2 Then
			     If AdsRS("Show")>=AdsRS("Shows") Then
				    Advertvirtualvalue=1
				 End If
			  ElseIf AdsRS("Class")=3 Then
			     If Now()>=AdsRS("lasttime") Then
				    Advertvirtualvalue=1
				 End If
			  ElseIf AdsRS("class")=4 Then
			     If AdsRS("Click")>=AdsRS("Clicks") Then
			        Advertvirtualvalue=1
				 End If
				 If AdsRS("Show")>=AdsRS("Shows") Then
			        Advertvirtualvalue=1
				 End If
			  ElseIf AdsRS("Class")=5 Then
			     If AdsRS("Click")>=AdsRS("Clicks") Then
				    Advertvirtualvalue=1
				 End If
				 If Now()>=AdsRS("lasttime") Then
				    Advertvirtualvalue=1
				 End If
			  ElseIf AdsRS("class")=6 Then
				 If AdsRS("show")>=AdsRS("shows") Then
				    Advertvirtualvalue=1
				 End If
				 If Now()>=AdsRS("lasttime") Then
				    Advertvirtualvalue=1
				 End If
			  ElseIf AdsRS("class")=7 Then
			     If AdsRS("Click")>=AdsRS("Clicks") Then
				    Advertvirtualvalue=1
				 End If
				 If AdsRS("Show")>=AdsRS("Shows") Then
				    Advertvirtualvalue=1
				 End If
				 If Now()>=AdsRS("lasttime") Then
				    Advertvirtualvalue=1
			     End If
		      End If
			  If Advertvirtualvalue>=1 Then
			     AdsRS("Act")=2
			     AdsRS.Update
			  End If
			  AdsRS.Movenext
			wend
			AdsRS.Close
			'结束 检测、更新
			set AdsRS=server.createobject("adodb.recordset")
			AdsSQL="Select id,sitename,intro,gif_url,window,show,place,time,xslei,wid,hei,url From KS_Advertise where act=1 and place=" & GetPlace & " order by Time"
			AdsRS.Open AdsSQL,Conn,1,3 
			If not AdsRS.EOF Then
			   AdsRS("show")=AdsRS("show")+1
			   AdsRS("Time")=Now()
			   AdsRS.Update
			   Select Case AdsRS("xslei")
		           Case "txt"
				   GetAdvertise="<a href=""" & DomainStr & "Plus/Advertise.asp?ID="&AdsRS("ID")&""">"&AdsRS("sitename")&"</a><br/>"
				   Case "gif"
				   GetAdvertise="<a href=""" & DomainStr & "Plus/Advertise.asp?ID="&AdsRS("ID")&"""><img src="""&AdsRS("Gif_Url")&""" alt="""&AdsRS("sitename")&"""/></a><br/>"
				   Case Else
				   GetAdvertise="请将"&AdsRS("ID")&"广告类型改为文本或图片<br/>"
			   End Select
			End if
			AdsRS.Close:Set AdsRS=Nothing
		End Function
	
		'*************************************************
		'共享广告	
		'*************************************************
		Public Function GetShareAdvertise()
		    On Error Resume Next
			Dim ShareID,TempStr,PlaceName
			ShareID=KS.ChkClng(KS.S("ShareID"))
			If ShareID=0 Then
			   TempStr=GetAdvertise(KS.WSetting(13))
			Else
			   PlaceName=KS.GetSingleFieldValue("select UserName from KS_User where UserID=" & ShareID & "")
			   If PlaceName="" Then
			      TempStr=GetAdvertise(KS.WSetting(13))
			   Else
			      TempStr=GetAdvertise(Conn.Execute("select place from KS_ADPlace where PlaceName='" & PlaceName & "'")(0))
			   End If
			End If
			GetShareAdvertise=TempStr 
		End Function

		'*********************************************************************************************************
		'函数名：ReplaceGeneralLabelContent
		'作  用：替换通用标签为内容
		'参  数：FileContent原文件
		'*********************************************************************************************************
		Function ReplaceGeneralLabelContent(F_C)
		    On Error Resume Next
		    Dim HtmlLabel,HtmlLabelArr, Param,LabelTotal,I
			F_C = ReplaceChannelLabel(F_C)
			F_C = Replace(F_C, "{$GetCopyRight}", KS.WSetting(8))'显示版权信息
			F_C = Replace(F_C, "{$GetSiteName}", KS.WSetting(3))'显示网站名称
			F_C = Replace(F_C, "{$GetSiteTitle}", KS.Setting(1))'显示网站标题
			F_C = Replace(F_C, "{$GetSiteLogo}", "<img src=""" & KS.WSetting(5) & """ alt=""Logo...""/>")'显示网站LOGO不带参数
			F_C = Replace(F_C,"{$GetTopUserLogin}",GetTopUserLogin)'显示会员登录入口(横排)
			F_C = Replace(F_C,"{$GetUserLogin}",GetUserLogin)'显示会员登录入口(竖排)
			F_C = Replace(F_C,"{$GetInstallDir}",Trim(KS.Setting(2) & KS.Setting(3) ) )'WAP安装目录
			F_C = Replace(F_C, "{$GetWebmaster}", KS.Setting(10))'显示站长
			F_C = Replace(F_C, "{$GetWebmasterEmail}", KS.Setting(11))'显示站长EMail
			F_C = Replace(F_C, "{$GetSiteUrl}", DomainStr)'显示网站URL
			F_C = Replace(F_C, "{$GetSearch}",GetIndexSearch)'显示搜索
			
			F_C = Replace(F_C,"{$GetGoBack}", "<anchor>返回上级<prev/></anchor>")'返回上级
			If InStr(F_C, "{$GetLocation}") <> 0 Then F_C = Replace(F_C, "{$GetLocation}",GetFolderNaviStr)'位置导航
			If InStr(F_C, "{$WapValue}") <> 0 Then F_C = Replace(F_C,"{$WapValue}",WapValue)''取出WAP值
			If InStr(F_C, "{$GetUrl}") <> 0 Then F_C = Replace(F_C,"{$GetUrl}",KS.GetUrl)'取得当前地址
			If InStr(F_C, "{$GetReadReturn}") <> 0 Then F_C = Replace(F_C,"{$GetReadReturn}",KS.GetReadReturn)'读取返回地址缓存超链接
			If InStr(F_C, "{$GetReadMessage}") <> 0 Then F_C = Replace(F_C, "{$GetReadMessage}", KS.GetReadMessage)'显示未读短消息
			If InStr(F_C, "{$GetGoBackIndex}") <> 0 Then F_C = Replace(F_C,"{$GetGoBackIndex}", "<a href="""&KS.GetGoBackIndex&""">返回首页</a>")'显示返回首页
			If InStr(F_C, "{$GetGoBackChannel}") <> 0 Then F_C = Replace(F_C,"{$GetGoBackChannel}",GetGoBackChannel)'显示栏目列表页返回频道首页
			If InStr(F_C, "{$GetGoBackClass}") <> 0 Then F_C = Replace(F_C,"{$GetGoBackClass}",GetGoBackClass)'显示内容页返回栏目列表页
			'替换网站Logo(带参数)
			If InStr(F_C,"{=GetLogo")<>0 Then
			   HtmlLabel = GetFunctionLabel(F_C,"{=GetLogo")
			   HtmlLabelArr=Split(HtmlLabel,"@@@")
			   For I=0 To Ubound(HtmlLabelArr)
			       Param = Split(GetFunctionLabelParam(HtmlLabelArr(I),"{=GetLogo"),",")
				   F_C = Replace(F_C, HtmlLabelArr(I),"<img src="""&KS.WSetting(5)&""" width="""&Param(0)&""" height="""&Param(1)&""" alt=""Logo...""/>")
			   Next
			End If

			'显示会员排行
			If InStr(F_C, "{=GetTopUser") <> 0 Then
			   HtmlLabel = GetFunctionLabel(F_C, "{=GetTopUser")
			   HtmlLabelArr=Split(HtmlLabel,"@@@")
			   For I=0 To Ubound(HtmlLabelArr)
			       Param = Split(GetFunctionLabelParam(HtmlLabelArr(I),"{=GetTopUser"),",")
				   F_C = Replace(F_C, HtmlLabelArr(I),GetTopUser(Param(0),Param(1)))
			   Next
	        End If
			'替换网站广告
			If InStr(F_C, "{=GetAdvertise") <> 0 Then
			   HtmlLabel = GetFunctionLabel(F_C, "{=GetAdvertise")
			   HtmlLabelArr=Split(HtmlLabel,"@@@")
			   For I=0 To Ubound(HtmlLabelArr)
			       Param = Split(GetFunctionLabelParam(HtmlLabelArr(I),"{=GetAdvertise"),",")
				   F_C = Replace(F_C, HtmlLabelArr(I),GetAdvertise(Param(0)))
			   Next
	        End If
			'替换共享广告
			If InStr(F_C, "{=GetShareAdvertise") <> 0 Then
			   HtmlLabel = GetFunctionLabel(F_C, "{=GetShareAdvertise")
			   HtmlLabelArr=Split(HtmlLabel,"@@@")
			   For I=0 To Ubound(HtmlLabelArr)
				   F_C = Replace(F_C, HtmlLabelArr(I),GetShareAdvertise)
			   Next
	        End If
			'显示网站调查
			If InStr(F_C, "{=GetVote") <> 0 Then
			   HtmlLabel = GetFunctionLabel(F_C, "{=GetVote")
			   HtmlLabelArr=Split(HtmlLabel,"@@@")
			   For I=0 To Ubound(HtmlLabelArr)
			       Param = Split(GetFunctionLabelParam(HtmlLabelArr(I),"{=GetVote"),",")
				   F_C = Replace(F_C, HtmlLabelArr(I),GetVote(Param(0)))
			   Next
	        End If
			'Tags通用标签
			If InStr(F_C, "{=GetTags") <> 0 Then
				 HtmlLabel = GetFunctionLabel(F_C, "{=GetTags")
				 HtmlLabelArr=Split(HtmlLabel,"@@@")
				 For I=0 To Ubound(HtmlLabelArr)
					 Param = Split(GetFunctionLabelParam(HtmlLabelArr(I), "{=GetTags"),",")
					 F_C = Replace(F_C, HtmlLabelArr(I), GetTags(Param(0),Param(1)))
				 Next
			End If
			F_C = Replace(F_C, "{$GetSiteCountAll}", GetSiteCountAll())'网站统计
			'显示在线人数
			If InStr(F_C, "{$GetOnline") <> 0 Then
			   Dim OnlineMany,OnlineMember
			   OnlineMany = Conn.Execute("Select Count(*) from KS_Online")(0)
			   OnlineMember = Conn.Execute("Select Count(*) from KS_Online where UserName <> '匿名用户'")(0)
			   F_C = Replace(F_C,"{$GetOnlineTotal}",OnlineMany)'总在线人数
			   F_C = Replace(F_C,"{$GetOnlineUser}",OnlineMember)'用户人数
			   F_C = Replace(F_C,"{$GetOnlineGuest}",OnlineMany-OnlineMember)'游客人数
	        End If
			'WAP自定义页面
			If InStr(F_C, "{=GetTemplate") <> 0 Then
				 HtmlLabel = GetFunctionLabel(F_C, "{=GetTemplate")
				 HtmlLabelArr=Split(HtmlLabel,"@@@")
				 For I=0 To Ubound(HtmlLabelArr)
					 Param = Split(GetFunctionLabelParam(HtmlLabelArr(I), "{=GetTemplate"),",")
					 F_C = Replace(F_C, HtmlLabelArr(I), "<a href="""&DomainStr&"plus/Template.asp?ID="&Param(0)&"&"&WapValue&""">"&Param(1)&"</a>")
				 Next
			End If
			'显示网站公告
			If InStr(F_C, "{=GetAnnounce") <> 0 Then
			   HtmlLabel = GetFunctionLabel(F_C, "{=GetAnnounce")
			   HtmlLabelArr=Split(HtmlLabel,"@@@")
			   For I=0 To Ubound(HtmlLabelArr)
			       Param = Split(GetFunctionLabelParam(HtmlLabelArr(I),"{=GetAnnounce"),",")
				   F_C = Replace(F_C, HtmlLabelArr(I),GetAnnounce(Param(0)))
			   Next
	        End If
			'显示当前时间
			If InStr(F_C, "{=GetCurrentTime") <> 0 Then
			   HtmlLabel = GetFunctionLabel(F_C, "{=GetCurrentTime")
			   HtmlLabelArr=Split(HtmlLabel,"@@@")
			   For I=0 To Ubound(HtmlLabelArr)
				   F_C = Replace(F_C, HtmlLabelArr(I),KS.DateFormat(Now(),Split(GetFunctionLabelParam(HtmlLabelArr(I),"{=GetCurrentTime"),",")(0)))
			   Next
	        End If
			ReplaceGeneralLabelContent = F_C
		End Function

		'*********************************************************************************************************
		'函数名：ReplaceNewsContent
		'作  用：替换文章内容页标签为内容
		'参  数：RS Recordset数据集,FileContent待替换的内容,Content文章内容
		'*********************************************************************************************************
		Function ReplaceNewsContent(ChannelID,RS, F_C, Content)
			Dim TempStr, N
			Dim HtmlLabel,HtmlLabelArr,I,Param
			On Error Resume Next
			F_C=LFCls.ReplaceUserDefine(ChannelID,F_C,RS)'替换自定义字段
			F_C = Replace(F_C, "{$GetComment}", "<a href="""&DomainStr&"plus/Comment.asp?ChannelID="&ChannelID&"&InfoID="&RS("ID")&"&"&WapValue&""">发表评论</a>")
			F_C = Replace(F_C, "{$GetFavorite}", "<a href="""&DomainStr&"plus/Favorite.asp?ChannelID="&ChannelID&"&InfoID="&RS("ID")&"&"&WapValue&""">收藏此文</a>")
			F_C = Replace(F_C, "{$GetToBlogInfo}", "<a href="""&DomainStr&"plus/ToBlogInfo.asp?ChannelID="&ChannelID&"&ID="&RS("ID")&"&"&WapValue&""">此文转载日记</a>")
		 	F_C = Replace(F_C, "{$ChannelID}", ChannelID)
			F_C = Replace(F_C, "{$InfoID}", RS("ID"))'当前文章小ID
			F_C = Replace(F_C, "{$ItemName}", KS.C_S(ChannelID,3))'当前项目名称
			F_C = Replace(F_C, "{$ItemUnit}", KS.C_S(ChannelID,4))'当前项目单位
			F_C = Replace(F_C, "{$GetArticleIntro}", RS("Intro"))
			F_C = Replace(F_C, "{$GetArticleShortTitle}", RS("Title"))
			'F_C = Replace(F_C, "{$GetArticleUrl}", "")
			F_C = Replace(F_C, "{$GetArticleKeyWord}", Replace(RS("KeyWords"), " ", ","))
			F_C = Replace(F_C, "{$GetKeyTags}",ReplaceKeyTags(RS("Keywords")))'Tags通用标签	
			F_C = Replace(F_C, "{$GetArticleAuthor}", LFCls.ReplaceDBNull(RS("Author"),"佚名"))
			F_C = Replace(F_C, "{$GetArticleInput}", "<a href='" & DomainStr & "/Space/Space.asp?UserName=" & RS("Inputer")&"'>" & rs("Inputer") & "</a>" )'文章录入
		    F_C = Replace(F_C, "{$GetArticleTitle}", LFCls.ReplaceDBNull(RS("FullTitle"),RS("Title")))
			F_C = Replace(F_C, "{$GetArticleOrigin}", KS.GetOrigin(LFCls.ReplaceDBNull(RS("Origin"),"本站原创")))'取得文章来源并附加上链接
			'文章内容
			If InStr(F_C,"{=GetArticleContent")<>0 Then
			   Dim ArticleContent
			   HtmlLabel = GetFunctionLabel(F_C,"{=GetArticleContent")
			   HtmlLabelArr=Split(HtmlLabel,"@@@")
			   For I=0 To Ubound(HtmlLabelArr)
			       Param = Split(GetFunctionLabelParam(HtmlLabelArr(I),"{=GetArticleContent"),",")
				   If Param(0)="True" Then
				      'HTML处理->HTML转UBB->取消HTML->UBBToHTML
				      ArticleContent=KS.UBBToHTML(KS.LoseHtml(KS.HTMLToUBB(KS.ReplaceTrim(KS.GetEncodeConversion(RS("ArticleContent"))))))
				   Else
				      'HTML处理->取消HTML
				      ArticleContent=KS.LoseHtml(KS.ReplaceTrim(KS.HTMLCode(RS("ArticleContent"))))
				   End If
				   F_C = Replace(F_C, HtmlLabelArr(I),KS.ContentPagination(ArticleContent,Param(1),"Show.asp?ID="&RS("ID")&"&ChannelID="&ChannelID&"&"&WapValue&"",Param(2),Param(3)))

			   Next
			End If
			'内容页图片
			If InStr(F_C, "{=GetPhoto") <> 0 Then
			   HtmlLabel = GetFunctionLabel(F_C, "{=GetPhoto")
			   Param = GetFunctionLabelParam(HtmlLabel, "{=GetPhoto")
			   Dim PhotoUrl:PhotoUrl=LFCls.ReplaceDBNull(RS("PhotoUrl"), DomainStr & "Images/Nopic.gif")
			   If KS.ChkClng(KS.S("Page"))<2 Then
			      If Not (IsNull(PhotoUrl) Or PhotoUrl = "") Then
					 F_C = Replace(F_C,HtmlLabel, KSLabel.GetPhoto(RS("ArticleContent"),RS("ID"),ChannelID,PhotoUrl,Split(Param, ",")(0),Split(Param, ",")(1)))
				  Else
					 F_C = Replace(F_C, HtmlLabel, "")
				  End If
			    Else
			      F_C = Replace(F_C, HtmlLabel, "")
			    End If
			End If

			'属性
			If InStr(F_C, "{$GetArticleProperty}") <> 0 Then
			   TempStr = ""
			   If CInt(RS("Recommend")) = 1 Then TempStr = TempStr&"荐 "
			   If CInt(RS("Popular")) = 1 Then TempStr = TempStr&"热 "
			   If CInt(RS("Strip")) = 1 Then TempStr = TempStr&"头 "
			   'TempStr = TempStr & "   " & Replace(RS("Rank"),"★","<img src=""" & DomainStr & "Images/Star.gif"" border=""0"">")
			   F_C = Replace(F_C, "{$GetArticleProperty}", TempStr)
		    End If
			F_C = Replace(F_C, "{$GetArticleHits}", RS("Hits"))'点击数
			If InStr(F_C, "{$GetArticleDate}") <> 0 Then
			   F_C = Replace(F_C, "{$GetArticleDate}", KS.DateFormat(RS("AddDate"), 6))'添加日期
			End If
			'心情指数
			If InStr(F_C,"{=GetMoodContent") <> 0 Then
			   HtmlLabel = GetFunctionLabel(F_C,"{=GetMoodContent")
			   HtmlLabelArr=Split(HtmlLabel,"@@@")
			   For I=0 To Ubound(HtmlLabelArr)
			       Param = Split(GetFunctionLabelParam(HtmlLabelArr(I),"{=GetMoodContent"),",")
				   F_C = Replace(F_C, HtmlLabelArr(I),KSLabel.GetMoodContent(Param(0),Param(1),ChannelID,RS("ID")))
			   Next
			End If
			'当允许评论时,则显示评论
			If InStr(F_C, "{=GetShowComment") <> 0 Then
			   HtmlLabel = GetFunctionLabel(F_C,"{=GetShowComment")
			   HtmlLabelArr=Split(HtmlLabel,"@@@")
			   For I=0 To Ubound(HtmlLabelArr)
			       Param = Split(GetFunctionLabelParam(HtmlLabelArr(I),"{=GetShowComment"),",")
				   If RS("Comment") = 1 Then
				      F_C = Replace(F_C, HtmlLabelArr(I),KSLabel.GetShowComment(Param(0),Param(1),ChannelID,RS("ID")))
				   Else
				      F_C = Replace(F_C, HtmlLabelArr(I), "")
				   End If
			   Next
	        End If
			'发表评论
			If InStr(F_C, "{$GetWriteComment}") <> 0 And RS("Comment") = 1 Then
			   F_C = Replace(F_C, "{$GetWriteComment}", KSLabel.GetWriteComment(ChannelID,RS("ID")))
			Else
			   F_C = Replace(F_C, "{$GetWriteComment}", "")
			End If
		    If InStr(F_C, "{$GetDigg}") <> 0 Then F_C = Replace(F_C,"{$GetDigg}",KSLabel.GetDigg(ChannelID,RS("ID")))'顶一下
			If InStr(F_C, "{$GetPrevArticle}") <> 0 Then F_C = Replace(F_C, "{$GetPrevArticle}", LFCls.ReplacePrevNext(ChannelID,RS("ID"), RS("Tid"), "<"))'上一篇
			If InStr(F_C, "{$GetNextArticle}") <> 0 Then F_C = Replace(F_C, "{$GetNextArticle}", LFCls.ReplacePrevNext(ChannelID,RS("ID"), RS("Tid"), ">"))'下一篇
			ReplaceNewsContent = F_C
		End Function
				
		'*********************************************************************************************************
		'函数名：ReplacePictureContent
		'作  用：替换图片内容页标签为内容
		'参  数：RS Recordset数据集,FileContent待替换的内容,PictureContent图片内容
		'*********************************************************************************************************
		Function ReplacePictureContent(ChannelID,RS, F_C, PictureContent)
			Dim TempStr, N
			Dim HtmlLabel,HtmlLabelArr,I,Param
			On Error Resume Next			 
			If InStr(F_C, "{$GetPictureByPage}") <> 0 And PictureContent<>"" Then
			   F_C = Replace(F_C,"{$GetPictureByPage}",PictureContent)'查看图片内容（上一页、下一页方式）
			End If
			F_C = Replace(F_C, "{$GetComment}", "<a href="""&DomainStr&"plus/Comment.asp?ChannelID="&ChannelID&"&InfoID="&RS("ID")&"&"&WapValue&""">我来评论</a>")
			F_C = Replace(F_C, "{$GetFavorite}", "<a href="""&DomainStr&"plus/Favorite.asp?ChannelID="&ChannelID&"&InfoID="&RS("ID")&"&"&WapValue&""">我要收藏</a>")
			F_C = Replace(F_C, "{$GetPictureIntro}",  KS.ReplaceInnerLink(RS("PictureContent")))'图片介绍
			F_C = Replace(F_C, "{$ChannelID}", ChannelID)
			F_C = Replace(F_C, "{$InfoID}", RS("ID"))'当前文章小ID
			F_C = Replace(F_C, "{$ItemName}", KS.C_S(ChannelID,3))'当前项目名称
			F_C = Replace(F_C, "{$ItemUnit}", KS.C_S(ChannelID,4))'当前项目单位
			F_C = Replace(F_C, "{$GetPictureID}", RS("ID"))'当前图片ID
			F_C = Replace(F_C, "{$GetPictureName}", RS("Title"))'图片名称
			'F_C = Replace(F_C, "{$GetPictureUrl}", "")'URL
			F_C = Replace(F_C, "{$GetPictureKeyWord}", Replace(RS("KeyWords"), " ", ","))'关键字
			F_C = Replace(F_C, "{$GetKeyTags}",ReplaceKeyTags(RS("Keywords")))
			F_C = LFCls.ReplaceUserDefine(ChannelID,F_C,RS)'替换自定义字段
			F_C = Replace(F_C, "{$GetPictureAuthor}", RS("Author"))'图片作者
			F_C = Replace(F_C, "{$GetPictureInput}", "<a href='"&DomainStr&"/Space/Space.asp?UserName=" & RS("Inputer")&"'>" & RS("inputer") & "</a>" )'图片录入
			F_C = Replace(F_C, "{$GetPictureSrc}", RS("PhotoUrl"))
			F_C = Replace(F_C, "{$GetPictureOrigin}", KS.GetOrigin(LFCls.ReplaceDBNull(RS("Origin"),"本站原创")))'取得文章来源并附加上链接
			'图片属性
			If InStr(F_C, "{$GetPictureProperty}") <> 0 Then
			   TempStr = ""
			   If CInt(RS("Recommend")) = 1 Then TempStr = TempStr&"荐 "
			   If CInt(RS("Popular")) = 1 Then TempStr = TempStr&"热 "
			   If CInt(RS("Strip")) = 1 Then TempStr = TempStr&"头 "
			   F_C = Replace(F_C, "{$GetPictureProperty}", TempStr)
		    End If
			F_C = Replace(F_C, "{$GetPictureStar}", Replace(RS("Rank"),"★","<img src="""&DomainStr&"Images/Star.gif"" alt=""""/>"))'显示推荐星级
			F_C = Replace(F_C, "{$GetPictureVote}", "<a href=""plus/PhotoVote.asp?ID="&RS("ID")&"&ChannelID="&ChannelID&"&"&WapValue&""">投它一票("&RS("Score")&")</a>")
			F_C = Replace(F_C,"{$GetPictureDate}",RS("AddDate"))'添加日期
			F_C = Replace(F_C,"{$GetPictureHits}",RS("Hits"))'图片人气（总浏览数）
			F_C = Replace(F_C,"{$GetPictureHitsByDay}",RS("HitsByDay"))'图片本日浏览数
			F_C = Replace(F_C,"{$GetPictureHitsByWeek}",RS("HitsByWeek"))'图片本周浏览数
			F_C = Replace(F_C,"{$GetPictureHitsByMonth}",RS("HitsByMonth"))'图片本月浏览数
			If InStr(F_C, "{$GetPictureDate}") <> 0 Then
			   F_C = Replace(F_C, "{$GetPictureDate}", KS.DateFormat(RS("AddDate"), 6))'添加日期
			End If
			'心情指数
			If InStr(F_C,"{=GetMoodContent") <> 0 Then
			   HtmlLabel = GetFunctionLabel(F_C,"{=GetMoodContent")
			   HtmlLabelArr=Split(HtmlLabel,"@@@")
			   For I=0 To Ubound(HtmlLabelArr)
			       Param = Split(GetFunctionLabelParam(HtmlLabelArr(I),"{=GetMoodContent"),",")
				   F_C = Replace(F_C, HtmlLabelArr(I),KSLabel.GetMoodContent(Param(0),Param(1),ChannelID,RS("ID")))
			   Next
			End If
			'当允许评论时,则显示评论
			If InStr(F_C, "{=GetShowComment") <> 0 Then
			   HtmlLabel = GetFunctionLabel(F_C,"{=GetShowComment")
			   HtmlLabelArr=Split(HtmlLabel,"@@@")
			   For I=0 To Ubound(HtmlLabelArr)
			       Param = Split(GetFunctionLabelParam(HtmlLabelArr(I),"{=GetShowComment"),",")
				   If RS("Comment") = 1 Then
				      F_C = Replace(F_C, HtmlLabelArr(I),KSLabel.GetShowComment(Param(0),Param(1),ChannelID,RS("ID")))
				   Else
				      F_C = Replace(F_C, HtmlLabelArr(I), "")
				   End If
			   Next
	        End If
			'发表评论
			If InStr(F_C, "{$GetWriteComment}") <> 0 And RS("Comment") = 1 Then
			   F_C = Replace(F_C, "{$GetWriteComment}", KSLabel.GetWriteComment(ChannelID,RS("ID")))
			Else
			   F_C = Replace(F_C, "{$GetWriteComment}", "")
			End If
		    If InStr(F_C, "{$GetDigg}") <> 0 Then F_C = Replace(F_C,"{$GetDigg}",KSLabel.GetDigg(ChannelID,RS("ID")))'顶一下
			If InStr(F_C, "{$GetPrevPicture}") <> 0 Then F_C = Replace(F_C, "{$GetPrevPicture}", LFCls.ReplacePrevNext(ChannelID,RS("ID"), RS("Tid"), "<"))'上一篇
			If InStr(F_C, "{$GetNextPicture}") <> 0 Then F_C = Replace(F_C, "{$GetNextPicture}", LFCls.ReplacePrevNext(ChannelID,RS("ID"), RS("Tid"), ">"))'下一篇
		    ReplacePictureContent = F_C
		End Function

		'*********************************************************************************************************
		'函数名：ReplaceDownLoadContent
		'作  用：替换下载内容页标签为内容
		'参  数：RS Recordset数据集,FileContent待替换的内容,DownContent图片内容
		'*********************************************************************************************************
		Function ReplaceDownLoadContent(ChannelID,RS, F_C)
			Dim TempStr,s,YSDZ, ZCDZ,N
			Dim HtmlLabel,HtmlLabelArr,Param,I
			On Error Resume Next
			F_C = Replace(F_C, "{$GetComment}", "<a href="""&DomainStr&"plus/Comment.asp?ChannelID="&ChannelID&"&InfoID="&RS("ID")&"&"&WapValue&""">我来评论</a>")
			F_C = Replace(F_C, "{$GetFavorite}", "<a href="""&DomainStr&"plus/Favorite.asp?ChannelID="&ChannelID&"&InfoID="&RS("ID")&"&"&WapValue&""">我要收藏</a>")
			F_C = LFCls.ReplaceUserDefine(ChannelID,F_C,RS)'替换自定义字段
			F_C = Replace(F_C, "{$ChannelID}", ChannelID)
			F_C = Replace(F_C, "{$InfoID}", RS("ID"))
			F_C = Replace(F_C, "{$ItemName}", KS.C_S(ChannelID,3))'当前项目名称
			F_C = Replace(F_C, "{$ItemUnit}", KS.C_S(ChannelID,4))'当前项目单位
		    F_C = Replace(F_C, "{$GetDownID}", RS("ID"))'当前软件ID
		    F_C = Replace(F_C, "{$GetDownKeyWord}", Replace(RS("KeyWords"), " ", ","))
			F_C = Replace(F_C, "{$GetKeyTags}",ReplaceKeyTags(RS("Keywords")))
			F_C = Replace(F_C, "{$GetDownTitle}", RS("Title")&""&RS("DownVersion"))'软件名称+版本号
			'F_C = Replace(F_C, "{$GetDownUrl}", "")'URL
			F_C = Replace(F_C, "{$GetDownSystem}", RS("DownPT"))'系统平台
			F_C = Replace(F_C, "{$GetDownAuthor}", LFCls.ReplaceDBNull(RS("Author"),"未知"))'下载作者
			F_C = Replace(F_C, "{$GetDownSize}", RS("DownSize"))'文件大小+MB（KB）
			F_C = Replace(F_C, "{$GetDownType}", RS("DownLB"))'软件类别
			F_C = Replace(F_C, "{$GetDownLanguage}", RS("DownYY"))'软件语言
			F_C = Replace(F_C, "{$GetDownPower}", RS("DownSQ"))'授权方式
			F_C = Replace(F_C, "{$GetStar}", Replace(RS("Rank"),"★","<img src=""" & DomainStr & "Images/Star.gif"" alt=""""/>"))'显示推荐星级
			F_C = Replace(F_C, "{$GetDownDecPass}", RS("JYMM"))
			F_C = Replace(F_C, "{$GetDownOrigin}", KS.GetOrigin(LFCls.ReplaceDBNull(RS("Origin"),"本站原创"))) 
			'下载地址
			If InStr(F_C, "{$GetDownAddress}") <> 0 Then
			   Dim UrlArr, TotalNum, AUrl, UrlStr
			   UrlArr = Split(RS("DownUrls"), "|||")
			   TotalNum = UBound(UrlArr)
			   For I = 0 To TotalNum
			       N=N+1
				   AUrl = Split(UrlArr(I), "|")
				   If AUrl(0)=0 Then
				      UrlStr = UrlStr & "<img src="""&DomainStr&"Images/Down.gif"" alt=""""/><a href="""&DomainStr&"plus/DownLoad.asp?ChannelID="&ChannelID&"&ID="&RS("ID")&"&DownID="&N&"&"&WapValue&""">"&AUrl(1)&"</a><br/>" & vbCrLf          
				   Else
				      Dim RS_S:Set RS_S=Server.CreateObject("ADODB.RecordSet")
					  RS_S.Open "Select DownloadName,IsDisp,DownloadPath,DownID,SelFont From KS_DownSer Where ParentID=" & AUrl(0),Conn,1,1
					  If RS_S.Eof Then
					     If TotalNum=0 Then UrlStr="暂不提供下载地址<br/>"
					  Else
					     DO While Not RS_S.Eof
						    IF RS_S(1)=1 Then
						       UrlStr = UrlStr & "<img src="""&DomainStr&"Images/Down.gif"" alt=""""/><a href="""&RS_S(2)&Aurl(2)&""">" & RS_S(0) & "</a><br/>" & vbCrLf          
						    Else
						       UrlStr = UrlStr & "<img src="""& domainstr & "Images/Down.gif"" alt=""""/><a href="""&DomainStr&"plus/DownLoad.asp?ChannelID="&ChannelID&"&ID="&RS("ID")&"&DownID="&N&"&Sid="&RS_S(3)&"&"&WapValue&""">"&RS_S(0)&"</a><br/>" & vbCrLf          
						    End If
						    RS_S.MoveNext
					     Loop
					   End If
					   RS_S.Close:Set RS_S=Nothing
				    End If
			    Next
			    F_C = Replace(F_C, "{$GetDownAddress}", UrlStr)
		    End If
			 
			YSDZ = RS("YSDZ")
			ZCDZ = RS("ZCDZ")
			If InStr(F_C, "{$GetDownLink}") <> 0 Then
			   Dim LinkStr
			   If Not (LCase(YSDZ) = "http://" Or YSDZ = "") Then
				  LinkStr = "<a href=""" & YSDZ & """>作者或开发商主页</a>"
			   End If
			   If Not (LCase(ZCDZ) = "http://" Or ZCDZ = "") Then
				  LinkStr = LinkStr & " <a href=""" & ZCDZ & """>注册地址</a>"
			   End If
			   F_C = Replace(F_C, "{$GetDownLink}", LinkStr)
		    End If
			If InStr(F_C, "{$GetDownYSDZ}") <> 0 Then
			   If LCase(YSDZ) = "http://" Or YSDZ = "" Then
				  F_C = Replace(F_C, "{$GetDownYSDZ}", "无")
			   Else
				  F_C = Replace(F_C, "{$GetDownYSDZ}", "<a href=""" & RS("YSDZ") & """>" & RS("YSDZ") & "</a>")
			   End If
		    End If
			If InStr(F_C, "{$GetDownZCDZ}") <> 0 Then
			   If LCase(ZCDZ) = "http://" Or ZCDZ = "" Then
				  F_C = Replace(F_C, "{$GetDownZCDZ}", "无")
			   Else
				  F_C = Replace(F_C, "{$GetDownZCDZ}", "<a href=""" & RS("ZCDZ") & """>" & RS("ZCDZ") & "</a>")
			   End If
		    End If
			F_C = Replace(F_C, "{$GetDownInput}", "<a href='" & DomainStr & "/Space/Space.asp?UserName=" & RS("Inputer")&"&"&WapValue&"'>" & rs("Inputer") & "</a>" )
			If InStr(F_C,"{=GetContentIntro")<>0 Then
			   HtmlLabel = GetFunctionLabel(F_C,"{=GetContentIntro")
			   HtmlLabelArr=Split(HtmlLabel,"@@@")
			   For I=0 To Ubound(HtmlLabelArr)
			       Param = Split(GetFunctionLabelParam(HtmlLabelArr(I),"{=GetContentIntro"),",")
				   TempStr = KS.GotTopic(KS.LoseHtml(KS.HTMLCode(RS("DownContent"))),Param(0))
				   If KS.strLength(RS("DownContent"))<Param(0) Then
				      TempStr=TempStr&"...<a href=""plus/Content.asp?ID="&RS("ID")&"&ChannelID="&ChannelID&"&"&WapValue&""">更多>></a>"
				   End If
				   Call KS.GetWriteinReturn("<a href="""&KS.GetUrl&""">返回"&KS.C_S(ChannelID,3)&"页</a>")
				   F_C = Replace(F_C, HtmlLabelArr(I),TempStr)
			   Next
			End If
			'下载缩略图(带参数)
			If InStr(F_C, "{=GetDownPhoto") <> 0 Then
			   HtmlLabel = GetFunctionLabel(F_C, "{=GetDownPhoto")
			   HtmlLabelArr=Split(HtmlLabel,"@@@")
			   For I=0 To Ubound(HtmlLabelArr)
				   Param = GetFunctionLabelParam(HtmlLabelArr(I), "{=GetDownPhoto")
				   Dim LogoWidth: LogoWidth = Split(Param, ",")(0)
				   Dim LogoHeight: LogoHeight = Split(Param, ",")(1)
				   Dim PhotoUrl:PhotoUrl=RS("PhotoUrl")
				   If KS.IsNul(PhotoUrl) Then PhotoUrl="images/nopic.gif"
				   if left(PhotoUrl,1)="/" then PhotoUrl=right(PhotoUrl,len(PhotoUrl)-1)
				   if lcase(left(PhotoUrl,4))<>"http" then PhotoUrl=KS.Setting(2) & KS.Setting(3) & PhotoUrl
				  F_C = Replace(F_C,HtmlLabelArr(I), "<img src=""" & PhotoUrl & """  width=""" & LogoWidth & """ height=""" & LogoHeight & """ alt=""""/>")

			   Next
			End If
			'下载属性
			If InStr(F_C, "{$GetDownProperty}") <> 0 Then
			   TempStr = ""
			   If CInt(RS("Recommend")) = 1 Then TempStr = TempStr&"荐 "
			   If CInt(RS("Popular")) = 1 Then TempStr = TempStr&"热 "
			   F_C = Replace(F_C, "{$GetDownProperty}", TempStr)
		    End If
			F_C = Replace(F_C,"{$GetDownHits}",RS("Hits"))'总下载点击数
			F_C = Replace(F_C,"{$GetDownHitsByDay}",RS("HitsByDay"))'本日点击数
			F_C = Replace(F_C,"{$GetDownPower}",RS("HitsByWeek"))'本周点击数
			F_C = Replace(F_C,"{$GetDownPower}",RS("HitsByWeek"))'本周点击数
			F_C = Replace(F_C,"{$GetDownHitsByMonth}",RS("HitsByMonth"))'本月点击数
			If InStr(F_C, "{$GetDownDate}") <> 0 Then
			   F_C = Replace(F_C, "{$GetDownDate}", KS.DateFormat(RS("AddDate"), 6))'添加（更新）日期
			End If
			'心情指数
			If InStr(F_C,"{=GetMoodContent") <> 0 Then
			   HtmlLabel = GetFunctionLabel(F_C,"{=GetMoodContent")
			   HtmlLabelArr=Split(HtmlLabel,"@@@")
			   For I=0 To Ubound(HtmlLabelArr)
			       Param = Split(GetFunctionLabelParam(HtmlLabelArr(I),"{=GetMoodContent"),",")
				   F_C = Replace(F_C, HtmlLabelArr(I),KSLabel.GetMoodContent(Param(0),Param(1),ChannelID,RS("ID")))
			   Next
			End If
			'当允许评论时,则显示评论
			If InStr(F_C, "{=GetShowComment") <> 0 Then
			   HtmlLabel = GetFunctionLabel(F_C,"{=GetShowComment")
			   HtmlLabelArr=Split(HtmlLabel,"@@@")
			   For I=0 To Ubound(HtmlLabelArr)
			       Param = Split(GetFunctionLabelParam(HtmlLabelArr(I),"{=GetShowComment"),",")
				   If RS("Comment") = 1 Then
				      F_C = Replace(F_C, HtmlLabelArr(I),KSLabel.GetShowComment(Param(0),Param(1),ChannelID,RS("ID")))
				   Else
				      F_C = Replace(F_C, HtmlLabelArr(I), "")
				   End If
			   Next
	        End If
			'发表评论
			If InStr(F_C, "{$GetWriteComment}") <> 0 And RS("Comment") = 1 Then
			   F_C = Replace(F_C, "{$GetWriteComment}", KSLabel.GetWriteComment(ChannelID,RS("ID")))
			Else
			   F_C = Replace(F_C, "{$GetWriteComment}", "")
			End If
		    If InStr(F_C, "{$GetDigg}") <> 0 Then F_C = Replace(F_C,"{$GetDigg}",KSLabel.GetDigg(ChannelID,RS("ID")))'顶一下
			If InStr(F_C, "{$GetPrevDown}") <> 0 Then F_C = Replace(F_C, "{$GetPrevDown}", LFCls.ReplacePrevNext(ChannelID,RS("ID"), RS("Tid"), "<"))'上一篇
			If InStr(F_C, "{$GetNextDown}") <> 0 Then F_C = Replace(F_C, "{$GetNextDown}", LFCls.ReplacePrevNext(ChannelID,RS("ID"), RS("Tid"), ">"))'下一篇
		    ReplaceDownLoadContent = F_C
		End Function

		'*********************************************************************************************************
		'函数名：ReplaceProductContent
		'作  用：替换内容页标签为内容
		'参  数：RS Recordset数据集,FileContent待替换的内容
		'*********************************************************************************************************
		Function ReplaceProductContent(ChannelID,RS, F_C)
		    Dim TempStr,N
			On Error Resume Next 
			'商品简介
			If InStr(F_C,"{=GetContentIntro")<>0 Then
			   HtmlLabel = GetFunctionLabel(F_C,"{=GetContentIntro")
			   HtmlLabelArr=Split(HtmlLabel,"@@@")
			   For I=0 To Ubound(HtmlLabelArr)
			       Param = Split(GetFunctionLabelParam(HtmlLabelArr(I),"{=GetContentIntro"),",")
				   TempStr = KS.GotTopic(KS.LoseHtml(KS.HTMLCode(RS("ProIntro"))),Param(0))
				   If KS.strLength(RS("ProIntro"))<Param(0) Then
				      TempStr=TempStr&"...<a href=""plus/Content.asp?ID="&RS("ID")&"&ChannelID="&ChannelID&"&"&WapValue&""">更多>></a>"
				   End If
				   Call KS.GetWriteinReturn("<a href="""&KS.GetUrl&""">返回"&KS.C_S(ChannelID,3)&"页</a>")
				   F_C = Replace(F_C, HtmlLabelArr(I),TempStr)
			   Next
			End If
		 	F_C = Replace(F_C, "{$ChannelID}", ChannelID)
			F_C = Replace(F_C, "{$InfoID}", RS("ID"))'小ID
			F_C = Replace(F_C, "{$ItemName}", KS.C_S(ChannelID,3))'当前项目名称
			F_C = Replace(F_C, "{$ItemUnit}", KS.C_S(ChannelID,4))'当前项目单位
			F_C = Replace(F_C, "{$GetProductID}", RS("ProID"))'当前模型ID
			F_C = Replace(F_C, "{$GetProductName}", RS("Title"))'商品名称
			'F_C = Replace(F_C, "{$GetProductUrl}", "")'Url
			'F_C = Replace(F_C, "{$GetProductInputer}", RS("Inputer"))'显示商品录入
			F_C = Replace(F_C, "{$GetProductModel}", RS("ProModel"))'商品型号
			F_C = Replace(F_C, "{$GetProductSpecificat}", RS("ProSpecificat"))'商品规格
			F_C = Replace(F_C, "{$GetProducerName}", RS("ProducerName"))'商品生产商
			F_C = Replace(F_C, "{$GetTrademarkName}", RS("TrademarkName"))'品牌商标
			F_C = Replace(F_C, "{$GetServiceTerm}", RS("ServiceTerm"))'服务期限
			F_C = Replace(F_C, "{$GetProductType}", GetProductType(RS("ProductType")))'销售类型
			F_C = Replace(F_C, "{$GetRank}",Replace(RS("Rank"),"★","<img src=""" & DomainStr & "Images/Star.gif"" alt=""""/>"))'推荐等级
			F_C = Replace(F_C, "{$GetTotalNum}",RS("TotalNum"))'库存数量
			F_C = Replace(F_C, "{$GetProductUnit}", RS("Unit"))'商品单位
            F_C = Replace(F_C, "{$GetProductHits}", RS("Hits"))'浏览次数
			F_C = Replace(F_C, "{$GetProductDate}", KS.DateFormat(RS("AddDate"), 6))'上架时间
			F_C = Replace(F_C, "{$GetPrice_Market}", RS("Price_Market"))'显示市场价
			F_C = Replace(F_C, "{$GetPrice}", RS("Price"))'显示当前零售价
			F_C = Replace(F_C, "{$GetPrice_Member}", RS("Price_Member"))'显示会员价
			F_C = Replace(F_C, "{$GetPrice_Original}", RS("Price_Original"))'显示原始零售价
			If RS("ProductType")=3 Then
			   F_C = Replace(F_C, "{$GetDiscount}", RS("Discount"))'显示折扣率
			Else
			   F_C = Replace(F_C, "{$GetDiscount}", "")
			End If
			F_C = Replace(F_C, "{$GetScore}", RS("Point"))'显示购物积分
			F_C = Replace(F_C, "{$GetAddCar}", "<a href=""" & DomainStr & "shop/ShoppingCart.asp?ID=" & RS("ID")& "&"&WapValue&""">加入购物车</a>")'加入购物车
			'F_C = Replace(F_C, "{$GetAddFav}", "<a href=""" & DomainStr & "User/index.asp?User_Favorite.asp?Action=Add&ChannelID=5&InfoID=" & RS("ID") & """ target=""_blank""><img src=""" & DomainStr & "Images/fav.gif"" border=""0""></a>")
			F_C = Replace(F_C, "{$GetFavorite}", "<a href="""&DomainStr&"plus/Favorite.asp?ChannelID="&ChannelID&"&InfoID="&RS("ID")&"&"&WapValue&""">我要收藏</a>")
			F_C = Replace(F_C, "{$GetProductKeyWord}", Replace(RS("KeyWords"), " ", ","))
			F_C = Replace(F_C, "{$GetKeyTags}",ReplaceKeyTags(RS("Keywords")))
			F_C = Replace(F_C, "{$GetProductPhotoURL}",RS("BigPhoto"))
			
			F_C=LFCls.ReplaceUserDefine(ChannelID,F_C,RS)'替换自定义字段
			If InStr(F_C, "{=GetProductPhoto") <> 0 Then
				 Dim I,HtmlLabel: HtmlLabel = GetFunctionLabel(F_C, "{=GetProductPhoto")
				 Dim HtmlLabelArr:HtmlLabelArr=Split(HtmlLabel,"@@@")
				 Dim PhotoUrl:PhotoUrl=RS("BigPhoto")
				 For I=0 To Ubound(HtmlLabelArr)
					 Dim Param: Param = GetFunctionLabelParam(HtmlLabelArr(I), "{=GetProductPhoto")
					 Dim LogoWidth: LogoWidth = Split(Param, ",")(0)
					 Dim LogoHeight: LogoHeight = Split(Param, ",")(1)
					 If Not (IsNull(PhotoUrl) Or PhotoUrl = "") Then
					  Dim TempBigPhoto:TempBigPhoto=PhotoUrl
					  If lcase(left(TempBigPhoto,4))<>"http" Then 
					   if left(TempBigPhoto,1)="/" then TempBigPhoto=right(TempBigPhoto,len(TempBigPhoto)-1)
					   TempBigPhoto=KS.Setting(2) &KS.Setting(3) & TempBigPhoto
					  end if
					  F_C = Replace(F_C,HtmlLabelArr(I), "<div align=""center""><img src=""" & TempBigPhoto & """  width=""" & LogoWidth & """ height=""" & LogoHeight & """ border=""0""></div>") 
					 Else
					  F_C = Replace(F_C, HtmlLabelArr(I), "<div align=""center""><img src=""" & DomainStr & "images/nopic.gif""  width=""" & LogoWidth & """ height=""" & LogoHeight & """ border=""0""></div>")
					 End If
			    Next
			   End If
		    '商品属性
		    If InStr(F_C, "{$GetProductProperty}") <> 0 Then
			   TempStr = ""
			   If CInt(RS("Recommend")) = 1 Then TempStr = TempStr&"荐 "
			   If CInt(RS("Popular")) = 1 Then TempStr = TempStr&"热 "
			   If CInt(RS("IsSpecial")) = 1 Then TempStr = TempStr&"特 "
			   F_C = Replace(F_C, "{$GetProductProperty}", TempStr)
		    End If
			'心情指数
			If InStr(F_C,"{=GetMoodContent") <> 0 Then
			   HtmlLabel = GetFunctionLabel(F_C,"{=GetMoodContent")
			   HtmlLabelArr=Split(HtmlLabel,"@@@")
			   For I=0 To Ubound(HtmlLabelArr)
			       Param = Split(GetFunctionLabelParam(HtmlLabelArr(I),"{=GetMoodContent"),",")
				   F_C = Replace(F_C, HtmlLabelArr(I),KSLabel.GetMoodContent(Param(0),Param(1),ChannelID,RS("ID")))
			   Next
			End If
			'当允许评论时,则显示评论
			If InStr(F_C, "{=GetShowComment") <> 0 Then
			   HtmlLabel = GetFunctionLabel(F_C,"{=GetShowComment")
			   HtmlLabelArr=Split(HtmlLabel,"@@@")
			   For I=0 To Ubound(HtmlLabelArr)
			       Param = Split(GetFunctionLabelParam(HtmlLabelArr(I),"{=GetShowComment"),",")
				   If RS("Comment") = 1 Then
				      F_C = Replace(F_C, HtmlLabelArr(I),KSLabel.GetShowComment(Param(0),Param(1),ChannelID,RS("ID")))
				   Else
				      F_C = Replace(F_C, HtmlLabelArr(I), "")
				   End If
			   Next
	        End If
			'发表评论
			If InStr(F_C, "{$GetWriteComment}") <> 0 And RS("Comment") = 1 Then
			   F_C = Replace(F_C, "{$GetWriteComment}", KSLabel.GetWriteComment(ChannelID,RS("ID")))
			Else
			   F_C = Replace(F_C, "{$GetWriteComment}", "")
			End If
			If InStr(F_C, "{$GetPrevProduct}") <> 0 Then F_C = Replace(F_C, "{$GetPrevProduct}", LFCls.ReplacePrevNext(ChannelID,RS("ID"), RS("Tid"), "<"))
			If InStr(F_C, "{$GetNextProduct}") <> 0 Then F_C = Replace(F_C, "{$GetNextProduct}", LFCls.ReplacePrevNext(ChannelID,RS("ID"), RS("Tid"), ">"))
			ReplaceProductContent = F_C
		End Function
		Function GetProductType(TypeID)
		    Select Case TypeID
			    Case 1:GetProductType="正常销售"
				Case 2:GetProductType="涨价销售"
				Case 3:GetProductType="降价销售"
			End Select
		End Function

		'替换频道专用标签
		Function ReplaceChannelLabel(F_C)
		    On Error Resume Next
			If FCls.RefreshFolderID="0" Or FCls.RefreshFolderID="" Then ReplaceChannelLabel=F_C:Exit Function
			If KS.ChkClng(FCls.ChannelID)<>0 Then
		 	   F_C=Replace(F_C,"{$GetChannelID}",FCls.ChannelID)
			   F_C=Replace(F_C,"{$GetChannelName}",KS.C_S(FCls.ChannelID,1))
			   F_C=Replace(F_C,"{$GetItemName}",KS.C_S(FCls.ChannelID,3))
			   F_C=Replace(F_C,"{$GetItemUnit}",KS.C_S(FCls.ChannelID,4))
			End If
		 	F_C=Replace(F_C,"{$GetClassID}",FCls.RefreshFolderID)
			F_C=Replace(F_C,"{$GetClassName}",KS.C_C(FCls.RefreshFolderID,1))
		    F_C=Replace(F_C,"{$GetClassUrl}",KS.GetFolderPath(FCls.RefreshFolderID))
		    ReplaceChannelLabel=F_C
		End Function
		
        '**************************************************
		'函数名：ChangeLableToFunction
		'作  用：将标签转换为函数执行
		'参  数： LabelContent  ----标签参数
		'返回值：函数执行结果
		'**************************************************
		Function ChangeLableToFunction(LabelContent)
		    Dim L_Arr:L_Arr = Split(LabelContent, ",")
			If L_Arr(0) = "" Then
			   ChangeLableToFunction = ""
			   Exit Function
			End If
			Select Case UCase(L_Arr(0))'小写转换成大写
			    'Case "GETWRITEINRETURN"'存入返回缓存链接
				   'ChangeLableToFunction=KS.GetWriteinReturn(L_Arr(1))
				 Case "GETPOURACCOUNT"'显示倒计时
				    ChangeLableToFunction=GetPourAccount(L_Arr(1), L_Arr(2))
				 Case "GETWENHOUYU"'显示根据当前的时间不同的问候语
				    ChangeLableToFunction=GetWenhouyu(L_Arr(1), L_Arr(2), L_Arr(3), L_Arr(4), L_Arr(5), L_Arr(6), L_Arr(7), L_Arr(8))
				 Case "GETINDEXCHANNEL"'显示首页频道导航
				    ChangeLableToFunction=GetIndexChannel(L_Arr(1), L_Arr(2))
				 Case "GETINDEXLIST"'显示首页最新文章
				    ChangeLableToFunction=GetIndexList(L_Arr(1), L_Arr(2), L_Arr(3), L_Arr(4), L_Arr(5), L_Arr(6),L_Arr(7))
				 '===============栏目标签===============
				 Case "GETSHOWCLASSLIST"'栏目属性（热门、最新、推荐）
				    ChangeLableToFunction=GetShowClassList(L_Arr(1), L_Arr(2), L_Arr(3), L_Arr(4), L_Arr(5))
				 Case "GETRANDOMPHOTOTEXT"'随机图文
				    ChangeLableToFunction=GetRandomPhotoText(L_Arr(1), L_Arr(2), L_Arr(3), L_Arr(4), L_Arr(5), L_Arr(6), L_Arr(7))
				 Case "GETCLASSLIST"'栏目分类
				    ChangeLableToFunction=GetClassList(L_Arr(1), L_Arr(2))
				 Case "GETSHOWCLASSCENT"'栏目终级列表分页
				    Application("PageParam")=LabelContent
					ChangeLableToFunction=Application("PageParam")
				    'ChangeLableToFunction=GetShowClassCent(L_Arr(1), L_Arr(2), L_Arr(3), L_Arr(4), L_Arr(5), L_Arr(6))  
				 Case "GETRANDOMCONTENTSLIST"'显示内容页随机列表
				    ChangeLableToFunction=KSLabel.GetRandomContentsList(L_Arr(1), L_Arr(2), L_Arr(3), L_Arr(4))  
				 Case "GETRELATEDCONTENTSLIST"'显示内容页相关列表
				    ChangeLableToFunction=KSLabel.GetRelatedContentsList(L_Arr(1), L_Arr(2), L_Arr(3), L_Arr(4))  
					
				 Case "GETCLUBLIST"   '显示论坛帖子 
				    ChangeLableToFunction=GetClubList(L_Arr(1), L_Arr(2), L_Arr(3), L_Arr(4), L_Arr(5), L_Arr(6))
				 Case "GETLOGLIST"    '显示日志
				    ChangeLableToFunction=GetLogList(L_Arr(1), L_Arr(2), L_Arr(3), L_Arr(4), L_Arr(5), L_Arr(6))
			Case Else
			   ChangeLableToFunction = ""
			   Exit Function
			End Select
		End Function

End Class
%>
