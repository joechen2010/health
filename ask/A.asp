<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="function.asp"-->
<!--#include file="../KS_Cls/template.asp"-->
<%

Dim KSCls
Set KSCls = New Ask_A
KSCls.Kesion()
Set KSCls = Nothing

Class Ask_A
        Private KS, KSR,KSUser,UserLoginTF,AnonymScore
		Private Sub Class_Initialize()
		 If (Not Response.IsClientConnected)Then
			Response.Clear
			Response.End
		 End If
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		  Set KSUser=New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KSR=Nothing
		 Set KSUser=Nothing
		 Set KS=Nothing
		End Sub
		%>
		<!--#include file="../KS_Cls/Kesion.IFCls.asp"-->
		<%
		Public Sub Kesion()
		           UserLoginTF=Cbool(KSUser.UserLoginChecked)
		           AnonymScore=KS.ChkClng(KS.ASetting(36))
				   Select Case LCase(Request.Form("Action"))
						Case "save"
							Call saveQuestion()
						Case Else
							Call showmain()
					End Select
		End Sub
		
		Sub ShowMain()
			 Dim FileContent
			 FileContent = KSR.LoadTemplate(KS.ASetting(21))    
			 FCls.RefreshType = "question" '设置刷新类型，以便取得当前位置导航等
			 FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
			 Immediate=false
			 Scan FileContent
			 Templates=KSR.KSLabelReplaceAll(Templates)
			 KS.Echo RexHtml_IF(Templates)
		End Sub
		
		Sub ParseArea(sTokenName, sTemplate)
        End Sub 
		
		Sub ParseNode(sTokenType, sTokenName)
			Select Case lcase(sTokenType)
				Case "ask"  
				  echo ACls.ReturnAskConfig(sTokenName)
				Case "get"
				  select case lcase(sTokenName)
				    case "username" echo ksuser.username
				    case "userscore" echo KS.ChkClng(KSUser.Score)
					case "anonymscore" echo anonymscore
					case "question" echo KS.CheckXSS(request("Q"))
				  end select
		    End Select 
        End Sub
		
        Sub saveQuestion()
		 	Dim Rs,SQL
			Dim AskTopic,classid,AskContent,RewardScores,Anonymous,Broadcast,UserNowScore,NeedScore
			Dim TopicID,classname,parentid,parentstr,TextLength
			If UserLoginTF=false Then
				KS.Die "<script>parent.ShowLogin();</script>"
			End If
			AskTopic = KS.Gottopic(KS.S("topic"),255)
			classid = KS.ChkClng(Request.Form("smallerclassid"))
			If ClassID=0 Then ClassID = KS.ChkClng(Request.Form("smallclassid"))
			If ClassID=0 Then ClassID = KS.ChkClng(Request.Form("classid"))
			RewardScores = KS.ChkClng(Request.Form("Scores"))
			Anonymous = KS.ChkClng(Request.Form("anonym"))
			Broadcast = KS.ChkClng(Request.Form("broadcast"))
			AskContent = KS.CheckScript(Request.Form("askcontent"))
			AskContent = KS.FilterIllegalChar(AskContent)
			AskTopic=KS.FilterIllegalChar(AskTopic)
			TextLength = KS.strLength(AskContent)
			If KS.ASetting(3) = "0" Then
				KS.Die "<script>alert('友情提示!\n\n本问吧暂时禁止提问!');</script>"
			End If
			If KS.ASetting(6)="1" And Trim(Request.Form("Verifycode"))<>Trim(Session("Verifycode")) Then
			   	KS.Die "<script>alert('友情提示!\n\n您输入的验证码不正确,请重输!');parent.document.getElementById('Verifycode').value='';parent.document.getElementById('VerifyImg').src='../plus/verifycode.asp?n='+ Math.random();</script>"
			End If
			
			If KS.ChkCLng(KS.ASetting(4))<>0 Then
				If TextLength < KS.ChkCLng(KS.ASetting(4)) Then
					KS.Die "<script>alert('友情提示!\n\n问题描述不能小于 " & KS.ASetting(4) & " 个字节!');</script>"
				End If
			End If
			If KS.ChkCLng(KS.ASetting(5))<>0 Then
				If TextLength > KS.ChkCLng(KS.ASetting(5)) Then
					KS.Die "<script>alert('友情提示!\n\n问题描述不能大于 " & KS.ASetting(5) & " 个字节!');</script>"
				End If
			End If
			If KS.ChkClng(KS.S("ExpiredDays"))>KS.ChkClng(KS.ASetting(41)) and KS.ChkClng(KS.ASetting(41))<>0 Then
				KS.Die "<script>alert('友情提示!\n\n对不起,本问吧设置问题最多有效天数为" & KS.Asetting(41) & "天!');</script>"
			End If
			
			If classid = 0 Then
				KS.Die "<script>alert('友情提示!\n\n请选择正确的问题分类!');</script>"
			End If
			
			Set Rs = Conn.Execute("SELECT top 1 classid,classname,parentid,parentstr FROM KS_AskClass WHERE classid="&classid)
			If Rs.BOF And Rs.EOF Then
			    Rs.Close : Set RS=Nothing
				KS.Die "<script>alert('友情提示!\n\n找不到分类,请正确选择您的问题分类!');</script>"
			Else
				classname = Rs(1)
				parentid = Rs(2)
				parentstr = Rs(3)
			End If
			Rs.Close:Set Rs = Nothing
			Set Rs = Conn.Execute("SELECT TopicID FROM KS_AskTopic WHERE UserName='"&KSUser.UserName&"' And title='"&AskTopic&"'")
			If Not (Rs.BOF And Rs.EOF) Then
			    RS.Close : Set RS=Nothing
				KS.Die "<script>alert('友情提示!\n\n问题已经提交过.请不要重复提交问题!');</script>"
			End If
			Rs.Close:Set Rs = Nothing
			
			UserNowScore=KSUser.Score
			NeedScore = 0
			If RewardScores > 0 Then
				NeedScore = RewardScores
				If KS.ChkClng(RewardScores) > KS.ChkClng(UserNowScore) Then
					KS.Die "<script>alert('亲爱的用户:\n\n您的积分不够,不能设置悬赏分!');</script>"
				End If
			End If
			If Anonymous > 0 Then
				NeedScore = NeedScore + AnonymScore
				If KS.ChkClng(NeedScore) > KS.ChkClng(UserNowScore) Then
					KS.Die "<script>alert('亲爱的用户:\n\n您的积分不够,不能设置匿名问题!\n\n设置匿名问题需要 " & AnonymScore & "分');</script>"
				End If
			End If
			
			
			'主题模式(TopicMode: 0=待解决的问题,1=已解决的问题,2=投票中的问题,3=用户分享主题,4=过期问题)
			'帖子模式(PostsMode: 0=问,1=回) expiration
			Set Rs = Server.CreateObject("ADODB.Recordset")
			SQL = "SELECT Top 1 * FROM KS_AskTopic WHERE (TopicID is null)"
			Rs.Open SQL,Conn,1,3
			Rs.Addnew
				Rs("classid") = classid
				Rs("username") = KSUser.UserName
				Rs("classname") = classname
				Rs("title") = AskTopic
				Rs("Expired") = 0
				Rs("Closed") = 0
				Rs("PostTable") = "KS_AskPosts1"
				Rs("DateAndTime") = Now()
				Rs("LastPostTime") = Now()
				Rs("ExpiredTime") = Now()+KS.ChkClng(KS.S("ExpiredDays"))
				Rs("LockTopic") = 0
				Rs("Reward") = RewardScores
				Rs("Hits") = 0
				Rs("PostNum") = 0
				Rs("CommentNum") = 0
				Rs("TopicMode") = 0
				Rs("AskedMode") = 0
				Rs("Highlight") = 0
				Rs("Broadcast") = Broadcast
				Rs("Anonymous") = Anonymous
				Rs("IsTop") = 0
				Rs("supplement") = 0
			Rs.Update
			RS.MoveLast
			TopicID=RS("TopicID")
			Rs.Close:Set Rs = Nothing
			Set Rs = Server.CreateObject("ADODB.Recordset")
			SQL = "SELECT * FROM [KS_AskPosts1] WHERE (postsid is null)"
			Rs.Open SQL,Conn,1,3
			Rs.Addnew
				Rs("classid") = classid
				Rs("TopicID") = TopicID
				Rs("UserName") = KSUser.UserName
				Rs("topic") = AskTopic
				Rs("content") = AskContent
				Rs("addText") = ""
				Rs("PostTime") = Now()
				Rs("DoneTime") = Now()
				Rs("length") = TextLength
				Rs("star") = 0
				Rs("satis") = 0
				Rs("LockTopic") = 0
				Rs("PostsMode") = 0
				Rs("VoteNum") = 0
				Rs("Plus") = 0
				Rs("Minus") = 0
				Rs("PostIP") = KS.GetIP()
				Rs("Report") = 0
			Rs.Update
			Rs.MoveLast
			 Call KS.FileAssociation(1032,rs("postsid"),AskContent ,0)
			Rs.Close:Set Rs = Nothing
			
			'积分处理
			 '悬赏
			If RewardScores>0 Then
			Call KS.ScoreInOrOut(KSUser.UserName,2,RewardScores,"系统","问吧提问题[" & AskTopic & "]设置悬赏分!",0,0)
			End If
			 '发问题赠送
			If KS.ChkClng(KS.ASetting(35))>0 Then
			Call KS.ScoreInOrOut(KSUser.UserName,1,KS.ChkClng(KS.ASetting(35)),"系统","问吧提问题[" & AskTopic & "]系统奖赏分!",0,0)
			End If
			 '匿名
			If AnonymScore>0 and Anonymous<>0 Then
			Call KS.ScoreInOrOut(KSUser.UserName,2,AnonymScore,"系统","问吧设置问题[" & AskTopic & "]匿名提问!",0,0)
			End If
			

			If Not IsObject(Application(KS.SiteSN&"_askclasslist")) Then ACls.LoadCategoryList
			   Dim Catelist:Set Catelist = Application(KS.SiteSN&"_askclasslist")
			   If Not Catelist Is Nothing Then
				Dim Node:Set Node=Catelist.documentElement.selectSingleNode("row[@classid="&classid&"]")
				Dim parentarr,k
				parentarr=split(Node.selectSingleNode("@parentstr").text,",")
				for k=0 to ubound(parentarr)-1
			       Conn.Execute ("UPDATE KS_AskClass SET AskPendNum=AskPendNum+1 WHERE classid=" & KS.ChkClng(parentarr(k)))
				next
		    End If
			
			
			Dim strReturnURL,Direct
			Response.Write "<script language=""JavaScript"">"
			If Direct = 0 Then Response.Write "alert('恭喜您!问题提交成功');"
			Response.Write "try{top.location='" & KS.Setting(3) & KS.Asetting(1) & "';"
			Response.Write "}catch(e){}"
			Response.Write "</script>"

		End Sub
	
	
%>
<%	
End Class
%>
