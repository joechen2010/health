<%
Class SpaceCls
      Private KS,KSBCls,KSRFObj
	  Private Action,UserName,UserType
	  Private CurrentPage,TotalPut,MaxPerPage
	  Private Sub Class_Initialize()
	      Set KS=New PublicCls
          Set KSBCls=New BlogCls
		  Set KSRFObj=New Refresh
		  Action=KS.S("Action")
		  UserName=KS.S("i")
		  MaxPerPage = 10
		  If UserName="" Then Response.End()
		 If KS.S("page") <> "" Then
			CurrentPage = KS.ChkClng(KS.G("page"))
		 Else
			CurrentPage = 1
		 End If
      End Sub
	  Private Sub Class_Terminate()
	      Set KS=Nothing
		  Set KSBCls=Nothing
		  Set KSLabel=Nothing
	  End Sub
	 
	 Function FriendList()
	     FriendList = "【我的好友】<br/>"
		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		 RSObj.Open "select friend,u.userface,u.username from ks_friend f inner join ks_user u on f.friend=u.username where f.username='" & username & "' and flag=1",Conn,1,1
		 If RSObj.EOF and RSObj.Bof  Then
		    FriendList = FriendList & "没有加好友！<br/>"
		 Else
		    TotalPut = RSObj.RecordCount
			If CurrentPage < 1 Then	CurrentPage = 1
			If (CurrentPage - 1) * MaxPerPage > TotalPut Then
			   If (TotalPut Mod MaxPerPage) = 0 Then
			      CurrentPage = TotalPut \ MaxPerPage
			   Else
			      CurrentPage = TotalPut \ MaxPerPage + 1
			   End If
			End If
			If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < TotalPut Then
			   RSObj.Move (CurrentPage - 1) * MaxPerPage
			Else
			   CurrentPage = 1
			End If
			Dim I
			Do While Not RSObj.EOF
			   FriendList = FriendList & "<a href=""blog.asp?i=" & RSObj("username") & "&amp;"&KS.WapValue&""">" & RSObj(0) & "</a><br/>" &vbcrlf
			   RSObj.Movenext
			   I = I + 1
			   If I >= MaxPerPage Then Exit Do
			Loop
			FriendList = FriendList & KS.ShowPagePara(TotalPut, MaxPerPage, "index.asp", True, "位", CurrentPage, KS.QueryParam("page"))
			FriendList = FriendList & "<br/>"
		 End If
		 RSObj.Close:Set RSObj=Nothing
	 End Function

	 
	 Function GroupList()
	     If UserType=1 Then
		    GroupList = "【公司圈子】<br/>"
		 Else
		    GroupList = "【我的圈子】<br/>"
		 End If
		 MaxPerPage = 2 
		 If KS.S("page") <> "" Then
		    CurrentPage = KS.ChkClng(KS.G("page"))
		 Else
		    CurrentPage = 1
		 End If
		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		 RSObj.Open "select * from KS_team where UserName='" & UserName & "' And verific=1 order by id desc",Conn,1,1
		 If RSObj.EOF and RSObj.Bof  Then
		    GroupList = GroupList & "没有创建圈子！<br/>"
		 Else
		    TotalPut = RSObj.RecordCount
			If CurrentPage < 1 Then	CurrentPage = 1
			If (CurrentPage - 1) * MaxPerPage > TotalPut Then
			   If (TotalPut Mod MaxPerPage) = 0 Then
			      CurrentPage = TotalPut \ MaxPerPage
			   Else
			      CurrentPage = TotalPut \ MaxPerPage + 1
			   End If
			End If
			If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < TotalPut Then
			   RSObj.Move (CurrentPage - 1) * MaxPerPage
			Else
			   CurrentPage = 1
			End If
			Dim I
			Do While Not RSObj.EOF
					Dim PhotoUrl:PhotoUrl=RSObj("PhotoUrl")
					If KS.IsNul(PhotoUrl) Then PhotoUrl="images/nopic.gif"
					if left(PhotoUrl,1)="/" then PhotoUrl=right(PhotoUrl,len(PhotoUrl)-1)
					if lcase(left(PhotoUrl,4))<>"http" then PhotoUrl=KS.Setting(2) & KS.Setting(3) & PhotoUrl

		       GroupList = GroupList & "<a href=""Group.asp?ID=" & RSObj("ID") & "&amp;"&KS.WapValue&"""><img src=""" & PhotoUrl& """ width=""110"" height=""80"" alt=""""/></a><br/>"
			   GroupList = GroupList & "<a href=""Group.asp?ID=" & RSObj("ID") & "&amp;"&KS.WapValue&"""> " & RSObj("TeamName") & "</a><br/>"
			   GroupList = GroupList & "时间:" &RSObj("AddDate") & "<br/>"
			   GroupList = GroupList & "主题:" & Conn.Execute("select Count(ID) from KS_TeamTopic where TeamID=" & RSObj("ID") & "  And ParentID=0")(0) & ""
			   GroupList = GroupList & "回复:" & Conn.Execute("select Count(ID) from KS_TeamTopic where TeamID=" & RSObj("ID"))(0) & " "
			   GroupList = GroupList & "成员:" & Conn.Execute("select Count(UserName)  from KS_TeamUsers where status=3 And TeamID=" & RSObj("ID"))(0) & "人<br/>"
			   RSObj.Movenext
			   I = I + 1
			   If I >= MaxPerPage Then Exit Do
			Loop
			GroupList = GroupList & KS.ShowPagePara(TotalPut, MaxPerPage, "Space.asp", True, "位", CurrentPage, "UserName="&UserName&"&amp;Action=" & Action & "&amp;" & KS.WapValue & "")
			GroupList = GroupList & "<br/>"
		 End If
		 RSObj.Close:Set RSObj=Nothing
	 End Function

	 
	 Function PhotoList()
	     If UserType=1 Then
		    PhotoList = "【公司相册】<br/>"
		 Else
		    PhotoList = "【我的相册】<br/>"
		 End If

		 MaxPerPage =4
		 If KS.S("page") <> "" Then
			CurrentPage = KS.ChkClng(KS.G("page"))
		 Else
			CurrentPage = 1
		 End If
		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		 RSObj.Open "Select * from KS_Photoxc Where username='" & username & "' and status=1 order by id desc",Conn,1,1
		 If RSObj.EOF and RSObj.Bof  Then
		    PhotoList = PhotoList & "没有创建相册！<br/>"
		 Else
		    TotalPut = RSObj.RecordCount
			If CurrentPage < 1 Then	CurrentPage = 1
			If (CurrentPage - 1) * MaxPerPage > TotalPut Then
			   If (TotalPut Mod MaxPerPage) = 0 Then
			      CurrentPage = TotalPut \ MaxPerPage
			   Else
			      CurrentPage = TotalPut \ MaxPerPage + 1
			   End If
			End If
			If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < TotalPut Then
			   RSObj.Move (CurrentPage - 1) * MaxPerPage
			Else
			   CurrentPage = 1
			End If
			Dim I
			Do While Not RSObj.EOF
			   	PhotoUrl=RSObj("PhotoUrl")
				If KS.IsNul(PhotoUrl) Then PhotoUrl="images/nopic.gif"
				if left(PhotoUrl,1)="/" then PhotoUrl=right(PhotoUrl,len(PhotoUrl)-1)
				if lcase(left(PhotoUrl,4))<>"http" then PhotoUrl=KS.Setting(2) & KS.Setting(3) & PhotoUrl

			   PhotoList = PhotoList & "<a href=""ShowPhoto.asp?xcid="&RSObj("ID")&"&amp;i="&RSObj("UserName")&"&amp;"&KS.WapValue&"""><img src="""&PhotoUrl&""" width=""120"" height=""90"" alt=""""/></a><br/>"
			   PhotoList = PhotoList & "<a href=""ShowPhoto.asp?xcid="&RSObj("ID")&"&amp;i="&RSObj("UserName")&"&amp;"&KS.WapValue&""">["&GetStatusStr(RSObj("Flag"))&"]"&RSObj("xcname")&"("&RSObj("xps")&"/"&RSObj("hits")&")</a><br/>"
			   RSObj.Movenext
			   I = I + 1
			   If I >= MaxPerPage Then Exit Do
			Loop
			PhotoList = PhotoList & KS.ShowPagePara(TotalPut, MaxPerPage, "Space.asp", True, "个", CurrentPage,KS.QueryParam("page"))
			PhotoList = PhotoList & "<br/>"
		 End If
		 RSObj.Close:Set RSObj=Nothing
	 End Function
	 
	 Function GetStatusStr(val)
	     Select Case Val
		 Case 1:GetStatusStr="公开"
		 Case 2:GetStatusStr="会员"
		 Case 3:GetStatusStr="密码"
		 Case 4:GetStatusStr="隐私"
		 End Select
		 GetStatusStr="<b>" & GetStatusStr & "</b>"
	 End Function
	 
	 Function LogList()
	     If UserType=1 Then
		    LogList = "【公司日志】<br/>"
		 Else
		    LogList = "【我的日志】<br/>"
		 End If
		 
		 MaxPerPage =KSBCls.GetUserBlogParam(UserName,"ListBlogNum")'取得用户参数
		 If KS.S("page") <> "" Then
			CurrentPage = KS.ChkClng(KS.G("page"))
		 Else
			CurrentPage = 1
		 End If
		 Dim ClassID:ClassID=KS.ChkClng(KS.S("ClassID"))
		 Dim Param:Param=" UserName='" & UserName &"'"
		 Dim KeyWord:KeyWord=KS.S("Date") '日历搜索
		 Dim Key:Key=KS.R(KS.S("Key")) '表单搜索
		 Dim Tag:Tag=KS.R(KS.S("Tag")) 
		 If IsDate(KeyWord) Then
		    If CInt(DataBaseType) = 1 Then
			   Param=Param & " And AddDate>='" & KeyWord & " 00:00:00' and AddDate<='" &KeyWord & " 23:59:59'"
			Else
			   Param=Param & " And AddDate>=#" & KeyWord & " 00:00:00# and AddDate<=#" &KeyWord & " 23:59:59#"
			End If
		 End If
		 If ClassID<>0 Then Param=Param & " And ClassID=" & ClassID
		 If Key <>"" Then Param=Param & " And Title Like '%" & Key & "%'"
		 If Tag <>"" Then Param=Param & " And Tags Like '%" & Tag & "%'"

		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		 RSObj.Open "Select * from KS_BlogInfo Where " & Param & " and Status=0 Order By AddDate Desc",Conn,1,1
		 If RSObj.EOF and RSObj.Bof  Then
		    If Key<>"" Then
			   LogList = LogList & "找不到日志标题含有""<b>" & key & "</b>""的日志!<br/>"
			Else
			   If KeyWord="" And ClassID=0 Then
			      LogList = LogList & "您还没有写日志！<br/>"
			   ElseIf ClassID<>0 Then
			      LogList = LogList & "找不到该分类的日志!<br/>"
			   Else
			      LogList = LogList & "日期：<b>" & KeyWord & "</b>,您没有写日志！<br/>"
			   End If
			End if
		 Else
		    TotalPut = RSObj.RecordCount
			If CurrentPage < 1 Then	CurrentPage = 1
			If (CurrentPage - 1) * MaxPerPage > TotalPut Then
			   If (TotalPut Mod MaxPerPage) = 0 Then
			      CurrentPage = TotalPut \ MaxPerPage
			   Else
			      CurrentPage = TotalPut \ MaxPerPage + 1
			   End If
			End If
			If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < TotalPut Then
			   RSObj.Move (CurrentPage - 1) * MaxPerPage
			Else
			   CurrentPage = 1
			End If
			Dim I
			Do While Not RSObj.EOF
			   Dim JFStr:If RSObj("Best")="1" Then JFStr="  <img src=""../Images/jh.gif"" alt=""""/>" Else JFStr=""
			   LogList = LogList & "<a href=""" & KSBCls.GetLogUrl(RSObj) & """>" & RSObj("Title") & "</a>" & JFStr & ""
			   LogList = LogList & "<br/>"
			   RSObj.Movenext
			   I = I + 1
			   If I >= MaxPerPage Then Exit Do
			Loop
			LogList = LogList & KS.ShowPagePara(TotalPut, MaxPerPage, "Space.asp", True, "个", CurrentPage, "UserName="&UserName&"&amp;Action=" & Action & "&amp;" & KS.WapValue & "")
			LogList = LogList & "<br/>"
		 End If
		 RSObj.Close:Set RSObj=Nothing
	 End Function
	 
	 
	 Function GuestList()
	     GuestList = "【用户留言】 "
		 GuestList = GuestList & "<a href=""Message.asp?action=ReplayComment&amp;UserName="&UserName&"&amp;"&KS.WapValue&""">签写留言</a><br/>"

		 MaxPerPage =5
		 If KS.S("page") <> "" Then
			CurrentPage = KS.ChkClng(KS.G("page"))
		 Else
			CurrentPage = 1
		 End If
		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		 RSObj.Open "Select * from KS_BlogMessage Where UserName='" & UserName & "' Order By AddDate Desc",Conn,1,1
		 If RSObj.EOF and RSObj.Bof  Then
		    GuestList = GuestList& "找不到任何留言信息！<br/>"
		 Else
		    TotalPut = RSObj.RecordCount
			If CurrentPage < 1 Then	CurrentPage = 1
			If (CurrentPage - 1) * MaxPerPage > TotalPut Then
			   If (TotalPut Mod MaxPerPage) = 0 Then
			      CurrentPage = TotalPut \ MaxPerPage
			   Else
			      CurrentPage = TotalPut \ MaxPerPage + 1
			   End If
			End If
			If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < TotalPut Then
			   RSObj.Move (CurrentPage - 1) * MaxPerPage
			Else
			   CurrentPage = 1
			End If
			Dim I
			GuestList = GuestList& "共有" & TotalPut & "条留言信息<br/><br/>"
			Do While Not RSObj.EOF
			   GuestList = GuestList& "标题:"&RSObj("Title")&"<br/>"
			   GuestList = GuestList& "内容:"&RSObj("Content")&"<br/>"
			   If Not IsNull(RSObj("Replay")) Or RSObj("Replay")<>"" Then
			      GuestList = GuestList& "以下为主人的回复:<br/>" & RSObj("Replay") & " 时间:" & RSObj("replaydate") &"<br/>"
			   End If
			   Dim LoginTF,MoreStr
			   LoginTF=Cbool(KSUser.UserLoginChecked)
			   IF LoginTF=true And KSUser.UserName=UserName Then
			      MoreStr = "<a href='../User/User_Blog.asp?Action=MessageDel&amp;ID=" & RSObj("ID") & "&amp;"&KS.WapValue&"'>删除</a> <a href='../User/User_Blog.asp?ID=" & RSObj("ID") & "&amp;Action=ReplayMessage&amp;"&KS.WapValue&"'>回复</a>"
			   Else
			      MoreStr = ""
			   End If
			   GuestList = GuestList& "" & RSObj("AnounName") & " 发表于:" & RSObj("AddDate") &"" & MoreStr & "<br/>"
			   GuestList = GuestList& "---------<br/>"
			   RSObj.Movenext
			   I = I + 1
			   If I >= MaxPerPage Then Exit Do
			Loop
			GuestList = GuestList & KS.ShowPagePara(TotalPut, MaxPerPage, "Space.asp", True, "个", CurrentPage, "UserName="&UserName&"&amp;Action=" & Action & "&amp;" & KS.WapValue & "")
			GuestList = GuestList & "<br/>"
		 End If
		 RSObj.Close:Set RSObj=Nothing
	 End Function
	 
	 Function ListXX()
		 Dim ChannelID:ChannelID=KS.ChkClng(KS.S("ChannelID"))
		 If ChannelID=0 Then ChannelID=1
		 
		  If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
			Dim ModelXML,Node
			Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
			For Each Node In ModelXML.documentElement.SelectNodes("channel")
			 if Node.SelectSingleNode("@ks21").text="1" and (Node.SelectSingleNode("@ks6").text<=5 and Node.SelectSingleNode("@ks6").text<>4) and Node.SelectSingleNode("@ks0").text<>6  and Node.SelectSingleNode("@ks0").text<>10 and Node.SelectSingleNode("@ks0").text<>9 Then
			  Listxx = Listxx& "<a href=""index.asp?ChannelID=" &Node.SelectSingleNode("@ks0").text & "&amp;Action=Listxx&amp;i=" & UserName & "&amp;" & KS.WapValue & """>" & Node.SelectSingleNode("@ks3").text & "</a> "
			 End If
			next
		 
		 Dim SQL,K,OPStr
		 Listxx = Listxx & "<br/>"
		 Dim Sqlstr,Max
		 Max = 3
		 Select Case KS.C_S(ChannelID,6) 
		  Case 1
		   SQLStr="Select top " & Max & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,Changes,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc,id desc"
		  Case 2
  		   SQLStr="Select top " & Max & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate,PhotoUrl from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc,id desc"
          Case 3
  		   SQLStr="Select top " & Max & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc,id desc"
		  Case 4
  		   SQLStr="Select top " & Max & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc,id desc"
		  Case 5
  		   SQLStr="Select top " & Max & " ID,Title,Tid,0,0,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc,id desc"
		  Case 7
  		   SQLStr="Select top " & Max & " ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc,id desc"
		  Case 8
  		   SQLStr="Select top " & Max & " ID,Title,Tid,0,0,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc,id desc"
		 End Select
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open SqlStr,Conn,1,1
		 If RS.EOF and RS.Bof  Then
		    ListXX = ListXX & "找不到您要的信息！<br/>"
		 Else
		    Do While Not RS.EOF
			   Listxx = Listxx & "<a href=""../Show.asp?ID=" & RS(0) & "&amp;ChannelID=" & ChannelID & "&amp;" & KS.WapValue & """>" & RS(1) & "</a><br/>"
			   RS.MoveNext
			Loop
		 End If	
		 RS.Close:set RS=Nothing
	 End Function
	 
	 Function GetUserPhoto(RS,totalPut,ChannelID)
	     Dim I,Url,PhotoUrl
		 Do While Not RS.Eof 
		    PhotoUrl=RS("PhotoUrl")
			If KS.IsNul(PhotoUrl) Then PhotoUrl="images/nopic.gif"
			if left(PhotoUrl,1)="/" then PhotoUrl=right(PhotoUrl,len(PhotoUrl)-1)
			if lcase(left(PhotoUrl,4))<>"http" then PhotoUrl=KS.Setting(2) & KS.Setting(3) & PhotoUrl

		    If ChannelID=2 Then
		       Url="Show_Photo.asp?i=" & UserName & "&amp;ID=" & RS(0) & "&amp;" & KS.WapValue & ""
		    Else
		       Url="../Show.asp?ID=" & RS(0) & "&amp;ChannelID=" & ChannelID & "&amp;" & KS.WapValue & ""
		    End If
			GetUserPhoto=GetUserPhoto & "<a href=""" & Url & """><img src=""" & PhotoUrl & """ width=""120"" height=""80"" alt=""""/></a><br/>"
			GetUserPhoto=GetUserPhoto & "<a href=""" & Url & """>" & KS.Gottopic(RS(1),15) & "</a><br/>"
			RS.MoveNext
			I = I + 1
			If I >= TotalPut Then Exit Do
		 Loop
	 End Function

	 Function xxList()
		 Dim ChannelID:ChannelID=KS.ChkClng(KS.S("ChannelID"))
		 If channelid=0 then channelid=1
		  If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
			Dim ModelXML,Node
			Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
			For Each Node In ModelXML.documentElement.SelectNodes("channel")
			 if Node.SelectSingleNode("@ks21").text="1" and (Node.SelectSingleNode("@ks6").text<=5 and Node.SelectSingleNode("@ks6").text<>4) and Node.SelectSingleNode("@ks0").text<>6  and Node.SelectSingleNode("@ks0").text<>10 and Node.SelectSingleNode("@ks0").text<>9 Then
			  xxList = xxList & "<a href=""index.asp?ChannelID=" &Node.SelectSingleNode("@ks0").text & "&amp;Action=xx&amp;i=" & UserName & "&amp;" & KS.WapValue & """>" & Node.SelectSingleNode("@ks3").text & "</a> "
			 End If
			next
		 
		 Dim SQL,K,OPStr,RSC
		 xxList = xxList & "<br/>"
		 xxList = xxList & "【投稿信息】<br/>"
		 If KS.C_S(ChannelID,6) =2 Then
		    MaxPerPage =4
		 Else
		    MaxPerPage =15
		 End If
		 If KS.S("page") <> "" Then
			CurrentPage = KS.ChkClng(KS.G("page"))
		 Else
			CurrentPage = 1
		 End If
		 Dim Sqlstr
		 Select Case KS.C_S(ChannelID,6) 
		  Case 1
		   SQLStr="Select ID,Title,Tid,ReadPoint,InfoPurview,Fname,Changes,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc"
		  Case 2
  		   SQLStr="Select ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate,photourl from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc"
          Case 3
  		   SQLStr="Select ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc"
		  Case 4
  		   SQLStr="Select ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc"
		  Case 5
  		   SQLStr="Select ID,Title,Tid,0,0,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc"
		  Case 7
  		   SQLStr="Select ID,Title,Tid,ReadPoint,InfoPurview,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc"
		  Case 8
  		   SQLStr="Select ID,Title,Tid,0,0,Fname,0,AddDate from " & KS.C_S(ChannelID,2) & " Where Inputer='" & UserName & "' Order By AddDate Desc"
		 End Select

		 Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		 RSObj.Open SqlStr,Conn,1,1
		 If RSObj.EOF and RSObj.Bof  Then
		    xxList = xxList & "找不到您要的信息！<br/>"
		 Else
		    TotalPut = RSObj.RecordCount
			If CurrentPage < 1 Then	CurrentPage = 1
			If (CurrentPage - 1) * MaxPerPage > TotalPut Then
			   If (TotalPut Mod MaxPerPage) = 0 Then
			      CurrentPage = TotalPut \ MaxPerPage
			   Else
			      CurrentPage = TotalPut \ MaxPerPage + 1
			   End If
			End If
			If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < TotalPut Then
			   RSObj.Move (CurrentPage - 1) * MaxPerPage
			Else
			   CurrentPage = 1
			End If
			Dim I
			If KS.C_S(ChannelID,6) =2 Then'图片显示不同
			   xxList = xxList & GetUserPhoto(RSObj,MaxPerPage,ChannelID)
			Else
			   Do While Not RSObj.Eof
			      xxList = xxList & "<a href=""../Show.asp?ID=" & RSObj(0) & "&amp;ChannelID=" & ChannelID & "&amp;" & KS.WapValue & """>" & Replace(RSObj(1),"&","&amp;") & "</a><br/>"
				  I = I + 1
				  If I >= MaxPerPage Then Exit Do
				  RSObj.MoveNext
			   Loop
			End if
			xxList = xxList & KS.ShowPagePara(TotalPut, MaxPerPage, "Space.asp", True, "个", CurrentPage, "UserName="&UserName&"&amp;Action=" & Action & "&amp;" & KS.WapValue & "")
			xxList = xxList & "<br/>"
		 End If
		 RSObj.Close:Set RSObj=Nothing
	 End Function
	 	
	 Function UserInfo()	 
	     On Error Resume Next
		 UserInfo = "【联系档案】<br/>"
	     Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_User Where UserName='" & UserName & "'",Conn,1,1
		 If RS.Eof And RS.Bof Then
		    UserInfo = UserInfo & "参数传递出错!"
		    RS.Close:Set RS=Nothing
		 End If
		 If RS("UserType")=1 Then
		    UserInfo = UserInfo & ReplaceEnterPriseInfo(KSRFObj.LoadTemplate(KS.WSetting(25)),RS("UserName"))'企业联系我们副模板
		 Else
		    UserInfo = UserInfo & ReplaceUserInfoContent(KSRFObj.LoadTemplate(KS.WSetting(22)),RS)'个人小档案副模板
		 End If
		 RS.Close:Set RS=Nothing
	 End Function
	 
	 Function ReplaceUserInfoContent(ByVal Content,ByVal RS)
	     Dim Privacy:Privacy=RS("Privacy")
		 Content=Replace(Content,"{$GetUserName}",RS("UserName"))
		 Dim UserFaceSrc:UserFaceSrc=RS("UserFace")
		 Dim FaceWidth:FaceWidth=KS.ChkClng(RS("FaceWidth"))
		 Dim FaceHeight:FaceHeight=KS.ChkClng(RS("FaceHeight"))
		
		 if left(UserFaceSrc,1)="/" then UserFaceSrc=right(UserFaceSrc,len(UserFaceSrc)-1)
		 if lcase(left(UserFaceSrc,4))<>"http" then UserFaceSrc=KS.Setting(2) & KS.Setting(3) & UserFaceSrc
		 
		 
		 Content=Replace(Content,"{$GetUserFace}","<img src=""" & UserFaceSrc & """ alt=""""/>")
		 '联系方式
		 If Privacy=2 Then
		    Content=Replace(Content,"{$GetEmail}","保密")
		 Else
		    Dim Email:Email=RS("Email")
			If IsNull(Email) Or Email="" Then Email="暂无"
			Content=Replace(Content,"{$GetEmail}",Email)
		 End If
		 If Privacy=2 Then
		    Content=Replace(Content,"{$GetQQ}","保密")
		 Else
		    Dim QQ:QQ=RS("QQ")
		    If IsNull(QQ) Or QQ="" Then QQ="暂无"
		    Content=Replace(Content,"{$GetQQ}",QQ)
		 End If
		 If Privacy=2 Then
		    Content=Replace(Content,"{$GetUC}","保密")
		 Else
		    Dim UC:UC=RS("UC")
		    If IsNull(UC) Or UC="" Then UC="暂无"
			Content=Replace(Content,"{$GetUC}",UC)
		 End If
		 If Privacy=2 Then
		    Content=Replace(Content,"{$GetMSN}","保密")
		 Else
		    Dim MSN:MSN=RS("MSN")
			If IsNull(MSN) Or MSN="" Then MSN="暂无"
			Content=Replace(Content,"{$GetMSN}",MSN)
		 End If
		 If Privacy=2 Then
		    Content=Replace(Content,"{$GetHomePage}","保密")
		 Else
		    Dim HomePage:HomePage=RS("MSN")
			If Not IsNull(HomePage) Then
			   Content=Replace(Content,"{$GetHomePage}","<a href=""" & RS("HomePage") & """ target=""_blank"">" & RS("HomePage") & "</a>")
			Else
			   Content=Replace(Content,"{$GetHomePage}","")
			End iF
		 End If
		 
		 If Privacy=2 or Privacy=1 Then
		    Content=Replace(Content,"{$GetRealName}","保密")
		 Else
		    Dim RealName:RealName=RS("RealName")
			If IsNull(RealName) Or RealName="" Then RealName="暂无"
			Content=Replace(Content,"{$GetRealName}",RealName)
		 End If
		 If Privacy=2 or Privacy=1 Then
		    Content=Replace(Content,"{$GetSex}","保密")
		 Else
		    Dim Sex:Sex=RS("Sex")
			If IsNull(Sex) Or Sex="" Then Sex="暂无"
			Content=Replace(Content,"{$GetSex}",Sex)
		 End If
		 If Privacy=2 or Privacy=1 Then
		    Content=Replace(Content,"{$GetBirthday}","保密")
		 Else
		    Dim BirthDay:BirthDay=RS("BirthDay")
			If IsNull(BirthDay) Or BirthDay="" Then BirthDay="暂无"
			Content=Replace(Content,"{$GetBirthday}",BirthDay)
		 End If
		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetIDCard}","保密")
		Else
		 Dim IDCard:IDCard=RS("IDCard")
		 If IsNull(IDCard) Or IDCard="" Then IDCard="暂无"
		 Content=Replace(Content,"{$GetIDCard}",IDCard)
		End If
		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetOfficeTel}","保密")
		Else
		 Dim OfficeTel:OfficeTel=RS("OfficeTel")
		 If IsNull(OfficeTel) Or OfficeTel="" Then OfficeTel="暂无"
		 Content=Replace(Content,"{$GetOfficeTel}",OfficeTel)
		End If
		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetHomeTel}","保密")
		Else
		 Dim HomeTel:HomeTel=RS("HomeTel")
		 If IsNull(HomeTel) Or HomeTel="" Then HomeTel="暂无"
		 Content=Replace(Content,"{$GetHomeTel}",HomeTel)
		End If
		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetMobile}","保密")
		Else
		 Dim Mobile:Mobile=RS("Mobile")
		 If IsNull(Mobile) Or Mobile="" Then Mobile="暂无"
		 Content=Replace(Content,"{$GetMobile}",Mobile)
		End If
		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetFax}","保密")
		Else
		 Dim Fax:Fax=RS("Fax")
		 If IsNull(Fax) Or Fax="" Then Fax="暂无"
		 Content=Replace(Content,"{$GetFax}",Fax)
		End If
		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetUserArea}","保密")
		Else
		 Dim Province:Province=RS("Province")
		 If IsNull(Province) Or Province="" Then Province=""
		 Dim City:City=RS("City")
		 If IsNull(City) Or Fax="" Then City="未知"
		 Content=Replace(Content,"{$GetUserArea}",Province & City)
		End If

		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetAddress}","保密")
		Else
		 Dim AddRess:AddRess=RS("AddRess")
		 If IsNull(AddRess) Or AddRess="" Then AddRess="暂无"
		 Content=Replace(Content,"{$GetAddress}",AddRess)
		End If
		If Privacy=2 or Privacy=1 Then
		 Content=Replace(Content,"{$GetZip}","保密")
		Else
		 Dim Zip:Zip=RS("Zip")
		 If IsNull(Zip) Or Zip="" Then Zip="暂无"
		 Content=Replace(Content,"{$GetZip}",ZIP)
		End If
		 If Privacy=2 or Privacy=1 Then
		    Content=Replace(Content,"{$GetSign}","保密")
		 Else
		    Dim Sign:Sign=RS("Sign")
		    If IsNull(Sign) Or Sign="" Then Sign="暂无"
			Content=Replace(Content,"{$GetSign}",Sign)
		 End If
         ReplaceUserInfoContent=Content
	 End Function
  
     '================================================企业空间部分=============================================
	 Function EnterPriseJob()
	     On Error Resume Next
		 EnterPriseJob = "【公司招聘】<br/>"
		 EnterPriseJob = EnterPriseJob & KS.HtmlCode(Conn.Execute("Select job From KS_EnterPrise Where UserName='" & UserName & "'")(0))
	 End Function
	 '公司介绍
	 Function EnterpriseIntro()
		     Dim str,ContentStr
			 On Error Resume Next
			 str="【公司简介】<br/>"
			 ContentStr=KS.UBBToHTML(KS.LoseHtml(KS.HTMLToUBB(KS.ReplaceTrim(KS.GetEncodeConversion(Conn.Execute("Select top 1 Intro From KS_EnterPrise Where UserName='" & UserName & "'")(0))))))
			 str=str & ContentStr
		     EnterpriseIntro=str
	 End Function
	 
	 '公司动态
	 Function EnterPriseNews()
	        Dim SQL,i,param,str
			str="【公司动态】<br/>"
			Dim RS:Set RS=Conn.Execute("Select classid,classname from ks_userclass where username='" & UserName & "' and typeid=4 order by orderid")
			If Not RS.Eof Then SQL=RS.GetRows(-1)
			RS.Close:Set RS=Nothing
			If IsArray(SQL) Then
			   str=str &"按分类查看"
			   If KS.S("ClassID")="" Then
			      str=str &"<a href='?i=" & UserName & "'><font color=red>全部文章</font></a>&nbsp;&nbsp;"
			   Else
			      str=str &"<a href='?i=" & UserName & "'>全部文章</a>&nbsp;&nbsp;"
			   End If
			   For I=0 To Ubound(SQL,2)
			       If KS.ChkClng(KS.S("ClassID"))=SQL(0,I) Then
				   str=str & "<a href='?pro=" & Server.UrlEncode(SQL(1,I)) & "&i=" & UserName & "&classid=" & SQL(0,i) & "'><font color=red>" & SQL(1,i) & "(" & conn.execute("select count(id) from ks_enterprisenews where classid=" & sql(0,i))(0) &")</font></a>&nbsp;&nbsp;&nbsp;&nbsp;"
				   Else
				   str=str & "<a href='?pro=" & Server.UrlEncode(SQL(1,I)) & "&i=" & UserName & "&classid=" & SQL(0,i) & "'>" & SQL(1,i) & "(" & conn.execute("select count(id) from ks_enterprisenews where classid=" & sql(0,i))(0) &")</a> "
				   End If
			   Next
			End If
			If KS.S("ClassID")="" Then
			   str=str &"所有新闻<br/>"
			Else
			   str=str &"" & KS.S("Pro") & "<br/>"
			End If
			Param=" Where UserName='" & UserName & "'"
			If KS.ChkClng(KS.S("ClassID"))<>0 Then Param=Param & " and classid=" & KS.ChkClng(KS.S("ClassID"))
			Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select ID,Title,AddDate From KS_EnterPriseNews " & Param &" order by adddate desc",conn,1,1
			If RS.EOF and RS.Bof  Then
			   str=str & "没有发布动态文章,请<a href=""../User/User_EnterPriseNews.asp?Action=Add&amp;"& KS.WapValue &""">点此发布</a>！<br/>"
			Else
			   TotalPut = RS.RecordCount
               If CurrentPage < 1 Then	CurrentPage = 1
			   If (totalPut Mod MaxPerPage) = 0 Then
			      pagenum = totalPut \ MaxPerPage
			   Else
			      pagenum = totalPut \ MaxPerPage + 1
			   End If
			   If CurrentPage>  1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
			      RS.Move (CurrentPage - 1) * MaxPerPage
			   Else
			      CurrentPage = 1
			   End If
			   SQL=RS.GetRows(-1)
			   Dim N,Total
			   Total=Ubound(SQL,2)+1
			   For I=0 To Total
			       str=str & "<a href=""Show_News.asp?i=" & UserName & "&amp;id=" & sql(0,n) & "&amp;" & KS.WapValue & """>" & SQL(1,N) & "</a>" & sql(2,n) & "<br/>"
				   N=N+1
				   If N>=Total Or N>=MaxPerPage Then Exit For
			   Next
			   str=str & KS.ShowPagePara(TotalPut, MaxPerPage, "index.asp", True, "件", CurrentPage, KS.QueryParam("page"))
		    End If
			RS.Close:Set  RS=Nothing
			str=str &"<br/>"   
			EnterPriseNews=str
	 End Function
	 
	 '公司产品
	 Function EnterPrisePro()
	        Dim SQL,i,param,str
			str="【产品展示】<br/>"
			Dim RS:Set RS=Conn.Execute("Select ClassID,ClassName from KS_UserClass where UserName='" & UserName & "' And typeid=3 order by OrderID")
			If Not RS.Eof Then SQL=RS.GetRows(-1)
			RS.Close:Set RS=Nothing
			If IsArray(SQL) Then
			   str=str &"按分类查看"
			   If KS.S("ClassID")="" Then
			      str=str &"<a href='?action=product&amp;i=" & UserName & "'>全部产品</a>"
			   Else
			      str=str &"<a href='?action=product&amp;i=" & UserName & "'>全部产品</a>"
			   End If
			   For I=0 To Ubound(SQL,2)
			       If KS.ChkClng(KS.S("ClassID"))=SQL(0,I) Then
				   str=str & "<a href='?action=product&amp;pro=" & Server.UrlEncode(SQL(1,I)) & "&amp;i=" & UserName & "&amp;classid=" & SQL(0,i) & "'><font color=""red"">" & SQL(1,i) & "(" & conn.execute("select count(id) from ks_product where classid=" & sql(0,i))(0) &")</font></a> "
				   Else
				   str=str & "<a href='?action=product&amp;pro=" & Server.UrlEncode(SQL(1,I)) & "&amp;i=" & UserName & "&amp;classid=" & SQL(0,i) & "'>" & SQL(1,i) & "(" & conn.execute("select count(id) from ks_product where classid=" & sql(0,i))(0) &")</a> "
				   End If
			   Next
		    End If
			   str=str &"<br/>"
			If KS.S("ClassID")="" Then
			   str=str &"所有产品<br/>"
			Else
			   str=str &"" & KS.S("Pro") & "<br/>"
			End If
			Param=" Where verific=1 and Inputer='" & UserName & "'"
			If KS.ChkClng(KS.S("ClassID"))<>0 Then Param=Param & " and classid=" & KS.ChkClng(KS.S("ClassID"))
			Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select ID,Title,PhotoUrl From KS_Product " & Param &" order by adddate desc,id",conn,1,1
			If RS.EOF and RS.Bof  Then
			   str=str & "没有发布产品展示,请<a href='../User/User_myshop.asp?Action=Add&amp;" & KS.WapValue & "'>点此发布</a>！<br/>"
			Else
			   TotalPut = RS.RecordCount
			   If CurrentPage < 1 Then	CurrentPage = 1
			   If (totalPut Mod MaxPerPage) = 0 Then
				  Pagenum = totalPut \ MaxPerPage
			   Else
			      Pagenum = totalPut \ MaxPerPage + 1
			   End If
			   If CurrentPage>  1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
			      RS.Move (CurrentPage - 1) * MaxPerPage
			   Else
			      CurrentPage = 1
			   End If
			   SQL=RS.GetRows(MaxperPage)
			   Dim N,Total,PhotoUrl
			   Total=Ubound(SQL,2)+1
			   For I=0 To Total
			        PhotoUrl=SQL(2,N)
					If KS.IsNul(PhotoUrl) Then PhotoUrl="images/nopic.gif"
					if left(PhotoUrl,1)="/" then PhotoUrl=right(PhotoUrl,len(PhotoUrl)-1)
					if lcase(left(PhotoUrl,4))<>"http" then PhotoUrl=KS.Setting(2) & KS.Setting(3) & PhotoUrl
				   
				   str=str & "<a href='Show_Product.asp?i=" & UserName & "&amp;id=" & sql(0,n) & "'><img src='" & PhotoUrl & "' width='130' height='90' alt=""""/></a><br/>"
				   str=str & "<a href='Show_Product.asp?i=" & UserName & "&amp;id=" & sql(0,n) & "'>" & SQL(1,N) & "</a><br/>"
				   N=N+1
				   If N>=Total Or N>=MaxPerPage Then Exit For
			   Next
			   str=str & KS.ShowPagePara(TotalPut, MaxPerPage, "index.asp", True, "件", CurrentPage, KS.QueryParam("page"))
			End If
			RS.Close:set RS=Nothing
			str=str &"<br/>"   
			EnterPrisePro=str
	 End Function
	 
	 Function ReplaceEnterpriseInfo(ByVal Content,ByVal UserName)
	     On Error Resume Next
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_EnterPrise Where UserName='" & UserName & "'",conn,1,1
		 IF RS.Eof Then
		    RS.Close:Set RS=Nothing
			ReplaceEnterpriseInfo=""
		 End If
		 Content=Replace(Content,"{$GetCompanyName}",RS("CompanyName"))
		 If isnull(RS("BusinessLicense")) Then
		    Content=Replace(Content,"{$GetBusinessLicense}","---")
		 Else
		 Content=Replace(Content,"{$GetBusinessLicense}",RS("BusinessLicense"))
		 End If
		 If isnull(RS("profession")) Then
		    Content=Replace(Content,"{$GetProfession}","---")
		 Else
		    Content=Replace(Content,"{$GetProfession}",RS("profession"))
		 End If
		 If isnull(RS("Companyscale")) Then
		    Content=Replace(Content,"{$GetCompanyScale}","---")
		 Else
		    Content=Replace(Content,"{$GetCompanyScale}",RS("Companyscale"))
		 End If
		 If isnull(rs("province")) Then
		    Content=Replace(Content,"{$GetProvince}","---")
		 Else
		    Content=Replace(Content,"{$GetProvince}",RS("province"))
		 End If
		 If isnull(rs("city")) Then
		    Content=Replace(Content,"{$GetCity}","---")
		 Else
		    Content=Replace(Content,"{$GetCity}",RS("city"))
		 End If
		 If isnull(RS("Contactman")) Then
		    Content=Replace(Content,"{$GetContactMan}","---")
		 Else
		    Content=Replace(Content,"{$GetContactMan}",RS("Contactman"))
		 End If
		 If isnull(RS("address")) Then
		    Content=Replace(Content,"{$GetAddress}","---")
		 Else
		    Content=Replace(Content,"{$GetAddress}",RS("address"))
		 End If
		 If isnull(RS("ZipCode")) Then
		    Content=Replace(Content,"{$GetZipCode}","---")
		 Else
		    Content=Replace(Content,"{$GetZipCode}",RS("zipcode"))
		 End If
		 If Isnull(RS("telphone")) Then
		    Content=Replace(Content,"{$GetTelphone}","---")
		 Else
		    Content=Replace(Content,"{$GetTelphone}",RS("telphone"))
		 End If
		 
		 If IsNull(RS("fax")) then
		    Content=Replace(Content,"{$GetFax}","---")
		 Else
		    Content=Replace(Content,"{$GetFax}",RS("fax"))
		 End If
		 If isnull(rs("weburl")) then
		    Content=Replace(Content,"{$GetWebUrl}","---")
		 Else
		    Content=Replace(Content,"{$GetWebUrl}",RS("weburl"))
		 End If
		 If isnull(rs("bankaccount")) then
		    Content=Replace(Content,"{$GetBankAccount}","---")
		 Else
		    Content=Replace(Content,"{$GetBankAccount}",RS("bankaccount"))
		 End If
		 If isnull(RS("accountnumber")) then
		    Content=Replace(Content,"{$GetAccountNumber}","---")
		 Else
		    Content=Replace(Content,"{$GetAccountNumber}",RS("accountnumber"))
	     End If
	     ReplaceEnterpriseInfo=Content
	 End Function
	 
End Class 
%>