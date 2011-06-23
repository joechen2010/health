<%
Class BlogCls
      Private KS
	  Private Sub Class_Initialize()
	      Set KS=New PublicCls
      End Sub
	  Private Sub Class_Terminate()
	      Set KS=Nothing
	  End Sub

	  '取得用户参数
	  Function GetUserBlogParam(UserName,FieldName)
	      Dim Num:Num=0
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "Select top 1 " & FieldName & " From KS_Blog Where UserName='" & UserName & "'",conn,1,1
		  If Not RS.EOF Then
		     Num=KS.ChkClng(RS(0))
		  End if
		  RS.Close:Set RS=Nothing
		  If Num=0 Then Num=10
		  GetUserBlogParam=Num
	  End Function
	 
	 '日志链接
	 Function GetLogUrl(RS)
	     GetLogUrl=GetCurrLogUrl(RS("ID"),RS("UserName"))
	 End Function
	 
	 Function GetCurrLogUrl(ID,UserName)
	     GetCurrLogUrl="List.asp?ID=" & ID & "&amp;UserName=" & UserName & "&amp;"&KS.WapValue&""
	 End Function
	 
	 '替换用户博客标签
	 Function ReplaceBlogLabel(RS,Template)
	     On Error Resume Next
		 Template=Replace(Template,"{$ShowAnnounce}",RS("Announce"))
		 Template=Replace(Template,"{$ShowBlogName}",RS("BlogName"))
		 Template=Replace(Template,"{$ShowLogo}",ReplaceLogo(RS("Logo")))
		 Template=Replace(Template,"{$ShowNavigation}",ReplaceMenu(RS("UserName")))
		 Template=Replace(Template,"{$ShowUserLogin}",GetUserLogin(RS("UserName")))
		 ReplaceBlogLabel=Template
	 End Function
	 
	 '**************************************************
	 '函数名：GetUserLogin
	 '作  用：显示会员登录入口
	 '**************************************************
	 Function GetUserLogin(UserName)
	     Dim TempStr
		 If Cbool(KSUser.UserLoginChecked)=True Then
		    Dim MyMailTotal:MyMailTotal=Conn.Execute("Select Count(ID) From KS_Message Where Incept='"&KSUser.UserName&"'And Flag=0 and IsSend=1 And delR=0")(0)
			TempStr = TempStr &"<a href="""&KS.GetDomain&"User/User_Message.asp?Action=inbox&amp;"&KS.WapValue&""">收信箱 "&MyMailTotal&"</a> <a href="""&KS.GetDomain&"User/Index.asp?"&KS.WapValue&""">会员中心</a> "
		 Else
		 TempStr = TempStr &"<a href="""&KS.GetDomain&"User/Login/?../Space/index.asp?i="&UserName&"&amp;Action="&KS.S("Action")&""">会员登陆</a> <a href="""&KS.GetDomain&"User/Reg/?../Space/index.asp?i="&UserName&"&amp;Action="&KS.S("Action")&""">会员注册</a>"
		 End if
		 GetUserLogin = TempStr
	 End Function

	 Function ReplaceLogo(Logo)
	     If Logo="" Or IsNull(Logo) Then Logo="../images/logo.jpg"
		 ReplaceLogo="<img src=""" & Logo & """ alt=""""/>"
	 End Function
	 
	 Function ReplaceMenu(ByVal UserName)
	     If Conn.Execute("Select UserType From KS_User WHere UserName='" & UserName & "'")(0)=0 Then
		    ReplaceMenu="<a href=""index.asp?i="&UserName&"&amp;" & KS.WapValue & """>首页</a> "
			ReplaceMenu=ReplaceMenu&"<a href=""index.asp?i="&UserName&"&amp;Action=log&amp;" & KS.WapValue & """>日志</a> "
			ReplaceMenu=ReplaceMenu&"<a href=""index.asp?i="&UserName&"&amp;Action=photo&amp;" & KS.WapValue & """>相册</a> "
			ReplaceMenu=ReplaceMenu&"<a href=""index.asp?i="&UserName&"&amp;Action=group&amp;" & KS.WapValue & """>圈子</a> "
			ReplaceMenu=ReplaceMenu&"<a href=""index.asp?i="&UserName&"&amp;Action=friend&amp;" & KS.WapValue & """>好友</a> "
			ReplaceMenu=ReplaceMenu&"<a href=""index.asp?i="&UserName&"&amp;Action=xx&amp;" & KS.WapValue & """>信息</a> "
			ReplaceMenu=ReplaceMenu&"<a href=""index.asp?i="&UserName&"&amp;Action=info&amp;" & KS.WapValue & """>档案</a> "
			ReplaceMenu=ReplaceMenu&"<a href=""index.asp?i="&UserName&"&amp;Action=guest&amp;" & KS.WapValue & """>留言</a>"
		 Else
		    ReplaceMenu="<a href=""index.asp?i="&UserName&"&amp;" & KS.WapValue & """>首页</a> "
			ReplaceMenu=ReplaceMenu&"<a href=""index.asp?action=intro&amp;i="&UserName&"&amp;" & KS.WapValue & """>简介</a> "
			ReplaceMenu=ReplaceMenu&"<a href=""index.asp?action=news&amp;i="&UserName&"&amp;" & KS.WapValue & """>动态</a> "
			ReplaceMenu=ReplaceMenu&"<a href=""index.asp?action=product&amp;i="&UserName&"&amp;" & KS.WapValue & """>展示</a> "
			ReplaceMenu=ReplaceMenu&"<a href=""index.asp?i="&UserName&"&amp;Action=job&amp;" & KS.WapValue & """>招聘</a> "
			ReplaceMenu=ReplaceMenu&"<a href=""index.asp?i="&UserName&"&amp;Action=photo&amp;" & KS.WapValue & """>相册</a> "
			ReplaceMenu=ReplaceMenu&"<a href=""index.asp?i="&UserName&"&amp;Action=group&amp;" & KS.WapValue & """>圈子</a> "
			ReplaceMenu=ReplaceMenu&"<a href=""index.asp?i="&UserName&"&amp;Action=log&amp;" & KS.WapValue & """>日志</a> "
			ReplaceMenu=ReplaceMenu&"<a href=""index.asp?i="&UserName&"&amp;Action=xx&amp;" & KS.WapValue & """>文集</a> "
			ReplaceMenu=ReplaceMenu&"<a href=""index.asp?i="&UserName&"&amp;Action=info&amp;" & KS.WapValue & """>联系</a> "
			ReplaceMenu=ReplaceMenu&"<a href=""index.asp?i="&UserName&"&amp;Action=guest&amp;" & KS.WapValue & """>留言</a>"
		 End If
	 End Function

	 
	 '替换所有标签
	 Function ReplaceAllLabel(UserName,Template)
	     Template=Replace(Template,"{$ShowUserInfo}",GetUserInfo(UserName))'用户信息
		 'Template=Replace(Template,"{$ShowCalendar}",Getcalendar(UserName))'日历
		 Template=Replace(Template,"{$ShowUserClass}",GetUserClass(UserName))'专栏列表
		 Template=Replace(Template,"{$ShowComment}",GetComment(UserName))'最新评论
		 'Template=Replace(Template,"{$ShowMusicBox}",GetMusicBox(UserName))'音乐盒
		 'Template=Replace(Template,"{$GetMediaPlayer}",GetMediaPlayer(UserName))
		 Template=Replace(Template,"{$ShowMessage}",GetMessage(UserName))'最新留言
		 Template=Replace(Template,"{$ShowBlogInfo}",GetBlogInfo(UserName))'最新日志
		 Template=Replace(Template,"{$ShowBlogTotal}",GetBlogTotal(UserName))'统计
		 Template=Replace(Template,"{$ShowSearch}",GetSearch(UserName))'搜索
		 Template=Replace(Template,"{$ShowUserName}",UserName)
		 '===========企业=======
		 Template=Replace(Template,"{$ShowContact}",GetEnterpriseContact(UserName))
		 Template=Replace(Template,"{$ShowNews}",GetEnterpriseNews(UserName))
		 ReplaceAllLabel=Template
	 End Function


	 Function GetEnterPriseNews(UserName)
	     On Error Resume Next
		 Dim RS:Set RS=Conn.Execute("Select top 10 ID,Title,AddDate From KS_EnterpriseNews where UserName='" & UserName & "' order by id desc")
		 Dim I,SQL:Sql=RS.GetRows(-1)
		 RS.Close:Set RS=Nothing
		 For I=0 To Ubound(SQL,2)
		     GetEnterPriseNews =GetEnterPriseNews & "<a href=""Show_News.asp?UserName=" & Username & "&amp;ID=" & SQL(0,I) & """>" & SQL(1,I) & "(" & formatdatetime(SQL(2,I),2) & ")</a><br/>"
		 Next
		 GetEnterPriseNews=GetEnterPriseNews
	 End Function

	 Function GetEnterpriseContact(username)
	     Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_EnterPrise Where UserName='" & UserName & "'",Conn,1,1
		 IF RS.Eof Then
		    RS.Close:Set RS=Nothing
			GetEnterpriseContact=""
			Exit Function
		 End If
		 GetEnterpriseContact="联 系 人：" & RS("Contactman") & "<br/>"
		 GetEnterpriseContact=GetEnterpriseContact & "公司地址：" & RS("AddRess") & "<br/>"
		 GetEnterpriseContact=GetEnterpriseContact & "邮政编码：" & RS("ZipCode") & "<br/>"
		 GetEnterpriseContact=GetEnterpriseContact & "联系电话：" & RS("TelPhone") & "<br/>"
		 GetEnterpriseContact=GetEnterpriseContact & "传真号码：" & RS("Fax") & "<br/>"
		 GetEnterpriseContact=GetEnterpriseContact & "公司网址：" & RS("WebUrl") & "<br/>"
		 RS.Close:Set RS=Nothing
	 End Function
	 
	 
	 '用户信息
	 Function GetUserInfo(UserName)
	     Dim str,RS:Set RS=Server.CreateObject("adodb.recordset")
		 RS.Open "select  top 1 UserFace,RealName,QQ from KS_User where UserName='" & UserName & "'",Conn,1,1
		 If not RS.EOF Then
		    'Dim UserFacesrc:UserFacesrc=RS(0)
			Dim RealName:RealName=RS(1)
			If RealName="" or Isnull(RealName) Then RealName=UserName
			'If Ucase(Left(UserFacesrc,4))<>"http" Then UserFacesrc="../" & UserFacesrc
			str="<a href=""Message.asp?UserName=" & UserName & "&amp;" & KS.WapValue & """> 给我留言</a>"
			str=str&"<a href=""../User/User_Friend.asp?Action=saveF&amp;ToUser=" & UserName & "&amp;" & KS.WapValue & """> 加为好友</a>"
			str=str&"<a href=""../User/User_Message.asp?Action=new&amp;ToUser=" & UserName & "&amp;" & KS.WapValue & """> 发送消息</a>"
			str=str & "<br/>"
		 End If
		 RS.Close:set RS=Nothing
		 GetUserInfo=str
	 End Function


	 '日历
	 'Function Getcalendar(UserName)
	     'Dim CalCls:Set CalCls=New CalendarCls
		 'Call CalCls.calendar(Getcalendar,UserName)
		 'Set CalCls=Nothing
	 'End Function
	 
	 '搜索
	 Function GetSearch(UserName)
	     GetSearch="关键字:<input type=""text"" size=""10"" name=""key""/>" & vbcrlf
		 GetSearch=GetSearch &"<anchor>搜索<go href=""Blog.asp?i=" & UserName & "&amp;"&KS.WapValue&""" method=""post"">" &vbcrlf
		 GetSearch=GetSearch & "<postfield name=""key"" value=""$(key)""/>" & vbcrlf
		 GetSearch=GetSearch & "</go></anchor><br/>"
	 End Function
	 
     '统计
	 Function GetBlogTotal(UserName)
	     GetBlogTotal="日志:"&Conn.Execute("select Count(id) from KS_BlogInfo where UserName='" & UserName &"' And status=0")(0) & " 篇"
		 GetBlogTotal=GetBlogTotal&" 回复:"&Conn.Execute("select Count(id) from KS_BlogComment where UserName='" & UserName &"'")(0) & " 条"
		 GetBlogTotal=GetBlogTotal&" 留言:"&Conn.Execute("select Count(id) from KS_BlogMessage where UserName='" & UserName &"'")(0) & " 条<br/>"
		 GetBlogTotal=GetBlogTotal&"日志阅读:"&Conn.Execute("select Sum(hits) from KS_BlogInfo where UserName='" & UserName &"' and status=0")(0) & " 人次"
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "select hits from KS_Blog where UserName='" & UserName & "'",Conn,1,3
		 If RS.EOF Then
		    hits=0
		 Else 
		    If RS(0)="" or isnull(RS(0)) Then
			   RS(0)=1
			Else
			  RS(0)=RS(0)+1
			End if
			hits=RS(0)
			RS.Update
		 End If
		 RS.Close:set RS=Nothing
		 GetBlogTotal=GetBlogTotal&" 总访问数:" & hits &"人次<br/>"
	 End Function
	 
	 '专栏列表
	 Function GetUserClass(UserName)
	     Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select ClassID,ClassName From KS_UserClass Where UserName='" & UserName & "' And TypeID=2",Conn,1,1
		 Do While Not RS.Eof 
		    GetUserClass=GetUserClass & "<a href=""Blog.asp?UserName=" & UserName & "&amp;ClassID=" & RS(0) & "&amp;"&KS.WapValue&""">" & RS(1) & "</a>" & vbcrlf
			RS.MoveNext
		 Loop
		 RS.Close:Set RS=Nothing
	 End Function
	 
	 '音乐盒
	 'Function GetMusicBox(UserName)
	     'GetMusicBox=""
	 'End Function
	 
	 'Function GetMediaPlayer(UserName)
	     'On Error Resume Next
		 'GetMediaPlayer="<EMBED style=""WIDTH: 272px; HEIGHT: 29px"" src=""" & Conn.Execute("select top 1 Url from KS_BlogMusic where UserName='" & UserName & "'")(0) & """ width=299 height=10 type=audio/x-wav autostart=""true"" loop=""true""></DIV></EMBED>"
	 'End Function
	 
	 '最新日志
	 Function GetBlogInfo(UserName)
	     Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select Top " & GetUserBlogParam(UserName,"ListLogNum") & " *  From KS_BlogInfo Where UserName='" & UserName & "' And Status=0 Order By ID Desc",Conn,1,1
		 If Not RS.Eof Then
		    Do While Not RS.EOF
			   GetBlogInfo=GetBlogInfo & "<a href=""" &GetCurrLogUrl(RS("ID"),RS("UserName")) & "&amp;"&KS.WapValue&""">" & RS("Title") & "</a><br/>" & vbcrlf
			   RS.MoveNext
			Loop
		 Else
			GetBlogInfo="暂无日志!<br/>"
		 End If
		 RS.Close:Set RS=Nothing
	 End Function
	 '最新评论
	 Function GetComment(UserName)
	     Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select Top " & GetUserBlogParam(UserName,"ListReplayNum") & " *  From KS_BlogComment Where UserName='" & UserName & "' Order By AddDate Desc",Conn,1,1
		 If Not RS.Eof Then
		    Do While Not RS.EOF
			   GetComment=GetComment & "<a href=""" &GetCurrLogUrl(RS("LogID"),RS("UserName")) & "&amp;"&KS.WapValue&""">" & RS("Title") & "</a><br/>" & vbcrlf
			   RS.MoveNext
			Loop
		 Else
		    GetComment="暂无评论!<br/>"
		 End If
		 RS.Close:Set RS=Nothing
	 End Function
	 
	 '最新留言
	 Function GetMessage(UserName)
	     'GetMessage="<a href=""Message.asp?UserName=" & UserName &"#write"">签写留言</a><br>"
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select Top " & GetUserBlogParam(UserName,"ListGuestNum") & " *  From KS_BlogMessage Where UserName='" & UserName & "' Order By AddDate Desc",Conn,1,1
		 If Not RS.Eof Then
		    Do While Not RS.EOF
			   GetMessage=GetMessage & "<a href=""message.asp?username=" & UserName &"#" & RS("ID")&""">" & RS("Title") & "</a><br/>"
			   RS.MoveNext
			Loop
		 End If
		 RS.Close:Set RS=Nothing
	 End Function
	 '天气
	 Function GetWeather(RS)
	     Dim TitleStr
	     Select Case RS("Weather")
		 Case "sun.gif":TitleStr="晴天"
		 Case "sun2.gif":TitleStr="和煦"
		 Case "yin.gif":TitleStr="阴天"
		 Case "qing.gif":TitleStr="清爽"
	     Case "yun.gif":TitleStr="多云"
		 Case "wu.gif":TitleStr="有雾"
		 Case "xiaoyu.gif":TitleStr="小雨"
	     Case "yinyu.gif":TitleStr="中雨"
		 Case "leiyu.gif":TitleStr="雷雨"
		 Case "caihong.gif":TitleStr="彩虹"
		 Case "hexu.gif":TitleStr="酷热"
		 Case "feng.gif":TitleStr="寒冷"
		 Case "xue.gif":TitleStr="小雪"
		 Case "daxue.gif":TitleStr="大雪"
		 Case "moon.gif":TitleStr="月圆"
		 Case "moon2.gif":TitleStr="月缺"
		 End Select
	 	 GetWeather="<img src=""../User/images/weather/" & RS("Weather") & """ alt=""" & TitleStr &"""/>"
	 End Function
	 
	 Function ReplaceLogLabel(UserName,ByVal TP,RS)
	     Dim EmotSrc:If RS("Face")<>"0" Then EmotSrc="<img src=""../User/images/face/" & RS("Face") & ".gif"" alt=""""/>"
		 Dim MoreStr
		 MoreStr="<a href=""" & GetLogUrl(RS) & """>阅读全文("&RS("hits")&")</a>|<a href=""" & GetLogUrl(RS) & """>回复（"& Conn.Execute("Select Count(ID) From KS_BlogComment Where LogID="  &RS("ID"))(0) &"）</a>"
		 Dim ContentStr
		 If IsNull(RS("Password")) Or RS("PassWord")="" Then 
		    ContentStr=KS.GotTopic(KS.LoseHtml(RS("Content")),GetUserBlogParam(UserName,"ContentLen"))
		 Else
		   ContentStr="请输入日志的查看密码：<input type=""password"" name=""pass"" size=""15""/><anchor>查看<go href=""" & GetLogUrl(RS) & """ method=""post""><postfield name=""pass"" value=""$(pass)""/></go></anchor>"
		 End IF
		 Dim JFStr:If RS("Best")="1" Then JFStr="  <img src=""../Images/jh.gif"" alt=""""/>" Else JFStr=""
		 TP=Replace(TP,"{$ShowLogTopic}",EmotSrc&"<a href=""" & GetLogUrl(RS) & """>" & RS("Title") & "</a>" & JFStr)
		 TP=Replace(TP,"{$ShowLogInfo}","[" & RS("AddDate") & "|by:" & RS("UserName") & "]")
		 TP=Replace(TP,"{$ShowLogText}",ContentStr)
		 TP=Replace(TP,"{$ShowLogMore}",MoreStr)
		 TP=Replace(TP,"{$ShowTopic}",RS("Title"))
		 TP=Replace(TP,"{$ShowAuthor}",RS("UserName"))
		 TP=Replace(TP,"{$ShowAddDate}",RS("AddDate"))
		 TP=Replace(TP,"{$ShowEmot}",EmotSrc)
		 TP=Replace(TP,"{$ShowWeather}",GetWeather(RS))
		 ReplaceLogLabel=TP
	 End Function
	 
	 '=============================圈子相关标签替换=============================
	 '替换标签
	 Function ReplaceGroupLabel(RS,Template)
	     On Error Resume Next
		 Template=Replace(Template,"{$ShowAnnounce}",RS("Announce"))'显示最新公告
		 Template=Replace(Template,"{$ShowNewUser}",GetUserList(RS("ID"),"new"))'显示最新加入成员列表
		 Template=Replace(Template,"{$ShowActiveUser}",GetUserList(RS("ID"),"active"))'显示最近活跃会员列表
		 Template=Replace(Template,"{$ShowGroupInfo}",GetGroupInfo(RS))'显示圈子信息
		 Template=Replace(Template,"{$ShowNavigation}",GetGroupMenu(RS))'显示圈子导航条等
		 Template=Replace(Template,"{$ShowGroupName}",RS("TeamName"))'显示圈子名称
		 Template=Replace(Template,"{$ShowGroupURL}",KS.GetDomain & "Space/Group.asp?ID=" & RS("ID"))'显示圈子URL
		 Template=Replace(Template,"{$ShowUserLogin}",GetUserLogin(RS("UserName")))
		 ReplaceGroupLabel=Template
	 End Function
	 
	 '圈子导航
	 Function GetGroupMenu(RS)
	     GetGroupMenu="<a href=""Group.asp?ID=" & RS("ID") &"&amp;"&KS.WapValue&""">首页</a> "
		 GetGroupMenu=GetGroupMenu&"<a href=""Group.asp?ID=" & RS("ID") &"&amp;Isbest=1&amp;"&KS.WapValue&""">精华</a> "
		 GetGroupMenu=GetGroupMenu&"<a href=""Group.asp?ID=" & RS("ID") &"&amp;Action=users&amp;"&KS.WapValue&""">成员</a> "
		 GetGroupMenu=GetGroupMenu&"<a href=""Group.asp?ID=" & RS("ID") &"&amp;Action=join&amp;"&KS.WapValue&""">加入</a> "
		 GetGroupMenu=GetGroupMenu&"<a href=""Group.asp?ID=" & RS("ID") &"&amp;Action=post&amp;"&KS.WapValue&""">发贴</a> "
		 GetGroupMenu=GetGroupMenu&"<a href=""Group.asp?ID=" & RS("ID") &"&amp;Action=info&amp;"&KS.WapValue&""">信息</a> "
	 End Function
     
	 '成员列表
	Function GetUserList(TeamID,Flag)
	    Dim Orderstr
		If Flag="active" Then
		   Orderstr=" order by LastLoginTime desc"
		Else
		   Orderstr=" order by a.id desc"
		End If
		Dim RS:set RS=Server.Createobject("adodb.recordset")
		RS.Open "select top 9 a.username,b.userid,b.userface,b.facewidth,b.faceheight from ks_teamusers a,ks_user b where a.username=b.username and status=3 and teamid="& TeamID & Orderstr,Conn,1,1
		Do While Not RS.EOF
		   GetUserList=GetUserList & "<a href=""index.asp?i=" & RS("UserName") & """>" & RS("UserName") & "</a><br/>"
		   RS.Movenext
		Loop
	End Function

    Function GetGroupInfo(RS)
	    GetGroupInfo = "成员人数:" & Conn.Execute("select Count(ID) from KS_TeamUsers where status=3 And TeamID=" & RS("ID"))(0) & "<br/>"
		GetGroupInfo = GetGroupInfo & "主题回复:" & Conn.Execute("select Count(ID) from KS_TeamTopic where TeamID=" & RS("ID") & "And Parentid=0")(0) & "/" &Conn.Execute("select Count(ID) from KS_TeamTopic where TeamID=" & RS("ID") & "  And ParentID<>0")(0) &"<br/>"
		GetGroupInfo = GetGroupInfo & "创建时间:" & RS("AddDate") &""
	End Function
	'=============================圈子相关标签替换结束==========================

End Class
%> 
