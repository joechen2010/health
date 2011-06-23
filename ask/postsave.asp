<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
Dim KSCls
Set KSCls = New Ask_A
KSCls.Kesion()
Set KSCls = Nothing

Class Ask_A

		Private Action,TopicID,TopicUseTable,CloseTopic
		Private AskTopic,classid,classname,Quserid,PostUsername
		Private Expired,Closed,DateAndTime,Reward,TopicMode,supplement
		Private allowAnswers,islock
		Private KS, KSR,KSUser,UserLoginTF,AnonymScore
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		  Set KSUser=New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		
		Sub Kesion()
            UserLoginTF=Cbool(KSUser.UserLoginChecked)
			allowAnswers = KS.ChkClng(KS.ASetting(7))
			Action = KS.S("action")
			TopicID = KS.ChkClng(Request.Form("TopicID"))
			CloseTopic = 0 : islock = 0
			
			Select Case LCase(Action)
				Case "saveanswer"
					Call saveanswer()
			End Select
				
			End Sub
			
			Sub saveanswer()
				Dim Rs,SQL,TextContent,UserNowPoint,TextLength
				If allowAnswers = 0 Then
					KS.Die "<script>alert('友情提示!\n\n本站暂时禁止回答问题!');</script>"
				End If
				If KS.ASetting(8)="1" And Trim(Request.Form("Verifycode"))<>Trim(Session("Verifycode")) Then
			   	    KS.Die "<script>alert('友情提示!\n\n您输入的验证码不正确,请重输!');parent.document.getElementById('VerifyCode').value='';parent.document.getElementById('VerifyImg').src='../plus/verifycode.asp?n='+ Math.random();</script>"
			    End If
				
				If KS.ASetting(13)="0" And KSUser.UserName="" Then
				 KS.Die "<script>alert('友情提示!\n\n本站不允许游客回答,请登录!');parent.ShowLogin()</script>"
				End If
			
				TextContent = Request.Form("TextContent")
				If TopicID = 0 Then
					KS.Die "<script>alert('友情提示!\n\n错误的系统参数!');</script>"
				End If
				If Len(TextContent) < 2 Then
					KS.Die "<script>alert('友情提示!\n\n请填写正确的答案内容!');</script>"
				End If
				TextContent = KS.FilterIllegalChar(KS.CheckScript(TextContent))
				TextLength = KS.strLength(TextContent)
				If KS.ChkClng(KS.ASetting(4))<>0 Then
					If TextLength < CLng(KS.ASetting(4)) Then
						KS.Die "<script>alert('友情提示!\n\n问题回答不能小于 " & KS.ASetting(4) & " 个字节!');</script>"
					End If
				End If
				If KS.ChkClng(KS.ASetting(5))<>0 Then
					If TextLength > CLng(KS.ASetting(5)) Then
						KS.Die "<script>alert('友情提示!\n\n问题回答不能大于 " & KS.ASetting(5) & " 个字节!');</script>"
					End If
				End If
				
				
				Call LoadTopicInfo(0)
				If TopicMode <> 0 Then
					KS.Die "<script>alert('友情提示!\n\n错误的系统参数!');</script>"
				End If
				
				Set Rs = Server.CreateObject("ADODB.Recordset")
				SQL = "SELECT top 1 * FROM KS_AskAnswer WHERE TopicID="& TopicID &" And UserName='"& KSUser.UserName & "'"
				Rs.Open SQL,Conn,1,3
				If Rs.BOF And Rs.EOF Then
					Rs.Addnew
					Rs("TopicID") = TopicID
					Rs("classid") = classid
					Rs("classname") = classname
					Rs("Username") = KSUser.UserName
					Rs("PostUsername") = PostUsername
					Rs("title") = AskTopic
					Rs("AnswerTime") = Now()
					Rs("PostTable") = TopicUseTable
					Rs("AnswerNum") = 1
					Rs("AnswerMode") = 0
					Rs("TopicMode") = 0
					Rs.Update
				Else
				    IF KS.ASetting(9)="1" Then
						If Rs("AnswerNum") >=1 Then
							KS.Die "<script>alert('友情提示!\n\n本问吧不允许重复提交答案!');</script>"
						End If
					End If
					Rs("AnswerTime") = Now()
					Rs("AnswerNum") = Rs("AnswerNum") + 1
					Rs.Update
				End If
				Rs.Close:Set Rs = Nothing
				Set Rs = Server.CreateObject("ADODB.Recordset")
				SQL = "SELECT * FROM " & TopicUseTable & " WHERE (postsid is null)"
				Rs.Open SQL,Conn,1,3
				Rs.Addnew
					Rs("classid") = classid
					Rs("TopicID") = TopicID
					Rs("UserName") = KSUser.UserName
					Rs("topic") = AskTopic
					Rs("content") = TextContent
					Rs("addText") = ""
					Rs("PostTime") = Now()
					Rs("DoneTime") = Now()
					Rs("length") = TextLength
					Rs("star") = 0
					Rs("satis") = 0
					Rs("LockTopic") = islock
					Rs("PostsMode") = 1
					Rs("VoteNum") = 0
					Rs("Plus") = 0
					Rs("Minus") = 0
					Rs("PostIP") = KS.GetIP()
					Rs("Report") = 0
				Rs.Update
				Rs.Close:Set Rs = Nothing
				
				If KS.ChkClng(KS.ASetting(30))>0 Then
				 Call KS.ScoreInOrOut(KSUser.UserName,1,KS.ChkClng(KS.ASetting(30)),"系统","问吧回答问题[" & AskTopic & "]悬赏!",0,0)
				End If
				Conn.Execute ("UPDATE KS_AskTopic SET PostNum=PostNum+1,LastPostTime=" & SqlNowString & " WHERE TopicID="& TopicID)
				Dim Direct
				Direct = KS.ChkClng(Request.Form("direct"))
				Response.Write "<script>"
				If Direct = 0 Then Response.Write "alert('恭喜您!答案提交成功');"
				Response.Write "try{top.location.replace(document.referrer);}catch(e){}</script>"
			End Sub
			
			
			
			Sub LoadTopicInfo(iMode)
				Dim SQL,Rs
				SQL = "SELECT TopicID,classid,classname,title,Username,Expired,Closed,PostTable,DateAndTime,Reward,PostNum,CommentNum,TopicMode,supplement FROM KS_AskTopic WHERE TopicID="&TopicID&" And TopicMode="&iMode&" And LockTopic=0"
				Set Rs = Conn.Execute(SQL)
				If Rs.BOF And Rs.EOF Then
				    Rs.Close:Set RS=Nothing
					Response.Write "<script>alert('友情提示!\n\n问题已经处理,不能回答!');</script>"
					Response.End()
				Else
					TopicID = Rs("TopicID")
					classid = Rs("classid")
					classname = Rs("classname")
					AskTopic = Rs("title")
					PostUsername = Rs("Username")
					Expired = Rs("Expired")
					Closed = Rs("Closed")
					TopicUseTable = Trim(Rs("PostTable"))
					DateAndTime = Rs("DateAndTime")
					Reward = Rs("Reward")
					TopicMode = Rs("TopicMode")
					supplement = Rs("supplement")
					'If iMode=0 And Rs("PostNum") > 10 Then
					'	CloseTopic = 1
					'End If
					If iMode=1 And Rs("CommentNum") > 100 Then
						CloseTopic = 1
					End If
					If Expired=1 Then
						Rs.Close:Set Rs=Nothing
						Response.Write "<script>alert('友情提示!\n\n问题已过期不能提交答案!');</script>"
						Response.End()
					End If
					If Closed = 1 Then
						Rs.Close:Set Rs=Nothing
						Response.Write "<script>alert('友情提示!\n\n问题已关闭不能提交答案!');</script>"
						Response.End()
					End If
					If CloseTopic = 1 Then
						Rs.Close:Set Rs=Nothing
						Response.Write "<script>alert('友情提示!\n\n问题已关闭不能提交答案!');</script>"
						Response.End()
					End If
					If KS.ChkClng(KS.ASetting(11))=0 Then
						If iMode=0 And PostUsername = KSUser.UserName Then
							Rs.Close:Set Rs=Nothing
							Response.Write "<script>alert('友情提示!\n\n不能回答自己提出的问题!');</script>"
							Response.End()
						End If
				    End If
				End If
				Set Rs = Nothing
			End Sub
End Class
%>
