<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="function.asp"-->
<%
Dim KSCls
Set KSCls = New Ask_Handle
KSCls.Kesion()
Set KSCls = Nothing

Class Ask_Handle

		Private Action,TopicID,TopicUseTable,CloseTopic
		Private AskTopic,classid,classname,Quserid,PostUsername
		Private Expired,Closed,DateAndTime,Reward,TopicMode,supplement
		Private allowAnswers,islock,Title,AddText
		Private KS, KSR,KSUser,UserLoginTF,AnonymScore
		Private XMLDom,UserReward,ExpiredTime,RemainDays
		Private  LockTopic,PostNum,TopicUserID

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
			Action = KS.S("action")
			TopicID = KS.ChkClng(Request("TopicID"))
			UserLoginTF=Cbool(KSUser.UserLoginChecked)
			If UserLoginTF=false Then Call KS.AlertHintScript("对不起,没有权限!"):Response.End
			If TopicID = 0 Then Call KS.AlertHintScript("参数传递出错!"):Response.End
			If Len(Action) = 0 Then KS.AlertHintScript("非法参数!"):Response.End
			Call showmain()
			
			Select Case LCase(Action)
				Case "0"
					Call HandleQuestion()
				Case "1"
					Call AddQuestion()
				Case "2"
					Call AdvanceReward()
				Case "3"
				     Call NoSatisAnswer()
				Case "selbest","handle"
					Call HandleQuestion()
				Case "saveadd"
					Call saveadd()
				Case "nosatis"
					Call Nosatis()
				Case "reward"
					Call AddReward()
				Case "delanswer"
					Call DelAnswer()
				Case Else
					Call HandleQuestion()
			End Select
				Call main()
         End Sub
		 
		 Sub NoSatisAnswer()
		 End Sub

		Sub HandleQuestion()
			Dim bestAnswerID,AnswerIDArray
			Dim Rs,SQL,i,star
			bestAnswerID = KS.ChkClng(KS.S("pid"))
			If LockTopic > 0 Or TopicMode=1 Or TopicMode=3 Then
				KS.Die "<script>alert('友情提示!\n\n问题已处理不能进行此项操作!');top.location.replace(document.referrer);</script>"
			End If
			If TopicID = 0 Then
				KS.Die "<script>alert('友情提示!\n\n请选择正确的问题ID!');history.back();</script>"
			End If
			If bestAnswerID = "" Then
				KS.Die "<script>alert('友情提示!\n\n请选择您满意的答案!');history.back();</script>"
			End If
			Set Rs = Server.CreateObject("ADODB.Recordset")
			SQL = "SELECT postsid,TopicID,username FROM ["&TopicUseTable&"] WHERE topicid="&topicid&" and PostsMode=1 and LockTopic=0 and postsid="& bestAnswerID
			Rs.Open SQL,Conn,1,1
			If Rs.BOF And Rs.EOF Then
				Rs.Close : Set Rs = Nothing
				KS.Die "<script>alert('友情提示!\n\n请选择正确的问题ID!');history.back();</script>"
			Else
				'UserReward = UserReward + KS.ChkClng(KS.ASetting(31))
				Do While Not Rs.EOF
					star = KS.ChkClng(Request.Form("star_"&Rs(0)))
					If star=0 Then star=3
					Conn.Execute ("UPDATE ["&TopicUseTable&"] SET satis=1,star="& star &",DoneTime="& SqlNowString &" WHERE postsid="& Rs(0))
					
					If KS.ChkClng(KS.ASetting(31))>0 Then
				     Call KS.ScoreInOrOut(Rs(2),1,UserReward,"系统","问吧回答问题[" & title & "]被采纳悬赏!",0,0)
				     Call KS.ScoreInOrOut(Rs(2),1,KS.ChkClng(KS.ASetting(31)),"系统","问吧回答问题[" & title & "]被采纳额外悬赏!",0,0)
				    End If
					
					Conn.Execute ("UPDATE KS_AskAnswer SET AnswerMode=1 WHERE topicid="&topicid&" and username='"& Rs(2) & "'")
					Rs.movenext
				Loop
			End If
			Rs.Close:Set Rs = Nothing
		
			Conn.Execute ("UPDATE KS_AskTopic SET LastPostTime="& SqlNowString &",TopicMode=1 WHERE topicid="&topicid&" and username='"& KSUser.UserName &"' and Closed=0 and LockTopic=0")
			Conn.Execute ("UPDATE KS_AskAnswer SET TopicMode=1 WHERE topicid="&topicid)
			
			If  KS.ChkClng(KS.ASetting(32))>0 Then
			Call KS.ScoreInOrOut(KSUser.UserName,1,KS.ChkClng(KS.ASetting(32)),"系统","问吧处理问题[" & Title & "]奖励!",0,0)
			End If
			
			Conn.Execute ("UPDATE KS_AskClass SET AskPendNum=AskPendNum-1,AskDoneNum=AskDoneNum+1 WHERE classid="& classid)
			
			Dim strReturnURL,Direct
			Direct = KS.ChkClng(Request.Form("direct"))
			strReturnURL = "q.asp?id=" & topicid
			Response.Write "<script language=""JavaScript"">"
			If Direct = 0 Then Response.Write "alert('满意答案设置成功!');"
			Response.Write "try{top.location='" & strReturnURL & "';"
			Response.Write "}catch(e){}"
			Response.Write "</script>"
		End Sub
		
		
		Sub AddQuestion()
			If LockTopic > 0 Or TopicMode=1 Then
				Response.Write "<script>alert('友情提示!\n\n问题已处理不能进行此项操作!');top.location.replace(document.referrer);</script>"
				Response.End
			End If
			If supplement > 5 Then
				Response.Write "<script>alert('友情提示!\n\n补充问题已超过5次,不能再进行问题补充!');top.location.replace(document.referrer);</script>"
				Response.End
			End If
			Dim SQL,Rs,Postlist,Node
			SQL="SELECT postsid,classid,TopicID,UserName,topic,content,addText,PostTime,DoneTime,satis,LockTopic,PostsMode,Report FROM ["&TopicUseTable&"] WHERE topicid="&topicid&" and PostsMode=0 and LockTopic=0 and UserName='"& KSUser.UserName & "'"
			Set Rs = Conn.Execute(SQL)
			If Not Rs.EOF Then
				addtext=Rs(6)
			End If
			Set Rs=Nothing
			SQL=Empty
			
			
		End Sub
		
		Sub saveadd()
			Dim AddContent,Rs,SQL,TextLength,TextContent
			AddContent = KS.S("AddContent")
			If LockTopic > 0 Or TopicMode=1 Then
				Response.Write "<script>alert('友情提示!\n\n问题已处理不能进行此项操作!');top.location.replace(document.referrer);</script>"
				Response.End()
			End If
			If Len(AddContent) < 2 Then
				Response.Write "<script>alert('友情提示!\n\n请输入要补充的问题内容!');history.back();</script>"
				Exit Sub
			End If
			AddContent = KS.HtmlEncode(KS.CheckScript(AddContent))
			TextLength = KS.strLength(AddContent)
			If TextLength > 2000 Then
				Response.Write "<script>alert('友情提示!\n\n补充问题内容太长!');history.back();</script>"
				Exit Sub
			End If
			If supplement > 5 Then
				Response.Write "<script>alert('友情提示!\n\n补充问题已超过5次,不能再进行问题补充!');history.back();</script>"
				Exit Sub
			End If
			
		
			Set Rs = Server.CreateObject("ADODB.Recordset")
			SQL = "SELECT * FROM [" & TopicUseTable & "] WHERE topicid="&topicid&" and PostsMode=0 and LockTopic=0 and satis=0 and username='"& KSUser.UserName & "'"
			Rs.Open SQL,Conn,1,3
			If Rs.BOF And Rs.EOF Then
				Set Rs = Nothing
				Response.Write "<script>alert('友情提示!\n\n错误的系统参数!');history.back();</script>"
				Response.End
			Else
				TextLength = CLng(KS.strLength(Rs("content")) + TextLength)
				Rs("addText") = AddContent
				Rs("length") = TextLength
				Rs.Update
			End If
			Rs.Close:Set Rs = Nothing
			Conn.Execute ("UPDATE KS_AskTopic SET supplement=supplement+1 WHERE topicid="&topicid&" and username='"& KSUser.UserName & "'")
			Response.Write "<script language=""JavaScript"">"
			Response.Write "alert('恭喜您!问题补充成功.');"
			Response.Write "try{top.location='q.asp?id=" & topicid & "';"
			Response.Write "}catch(e){}"
			Response.Write "</script>"
		End Sub
		
		Sub AdvanceReward()
			If UserReward > 300 Then
				Response.Write "<script>alert('友情提示!\n\n您的问题悬赏已高于300分不能再提高悬赏!');top.location.replace(document.referrer);</script>"
				Exit Sub
			End If
		End Sub
		
		Sub DelAnswer()
			Dim allowDeletes:allowDeletes = KS.ChkClng(KS.ASetting(12))
			If allowDeletes = 0 Then
				Response.Write "<script>alert('友情提示!\n\n禁止用户删除回答!');history.back();</script>"
				Response.End
			End If
			If LockTopic > 0 Or TopicMode=1 Then
				Response.Write "<script>alert('友情提示!\n\n问题已处理不能进行此项操作!');top.location.replace(document.referrer);</script>"
				Exit Sub
			End If
			Dim Rs,SQL,postsid,AnswerUserName,MinusPoints,MinusExperience,totalnumber
			postsid = KS.ChkClng(Request("pid"))
			If postsid = 0 Then
				Response.Write "<script>alert('友情提示!\n\n错误的系统参数!');history.back();</script>"
				Response.End
			End If
			SQL = "SELECT postsid,TopicID,username FROM ["&TopicUseTable&"] WHERE topicid="&topicid&" and postsid="& postsid &" And PostsMode=1 and LockTopic=0 and satis=0"
			Set Rs = Conn.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				Set Rs = Nothing
				Response.Write "<script>alert('友情提示!\n\n错误的系统参数!');history.back();</script>"
				Response.End
			Else
				postsid = Rs("postsid")
				AnswerUserName = Rs("username")
			End If
			Set Rs = Nothing
			MinusPoints = KS.ChkClng(KS.ASetting(37))
			If AnswerUserName <> ""  Then
				totalnumber = CLng(Conn.Execute("SELECT COUNT(*) FROM KS_AskAnswer WHERE topicid="&topicid&" And username='"&AnswerUserName & "'")(0))
				If totalnumber > 1 Then
					Conn.Execute ("UPDATE [KS_AskAnswer] SET AnswerNum=" & totalnumber-1 & " WHERE topicid="&topicid&" and username='"&AnswerUserName & "'")
				Else
					Conn.Execute ("DELETE FROM [KS_AskAnswer] WHERE topicid="&topicid&" and username='"&AnswerUserName & "'")
				End If
			End If
			Conn.Execute ("DELETE FROM ["&TopicUseTable&"] WHERE topicid="&topicid&" and postsid="& postsid)
			Conn.Execute ("UPDATE KS_AskTopic SET PostNum=PostNum-1 WHERE topicid="&topicid&" and username='"& KSUser.UserName &"' and Closed=0 and LockTopic=0")
			
			If MinusPoints>0 Then
				 Call KS.ScoreInOrOut(AnswerUserName,2,MinusPoints,"系统","问吧问题[" & Title & "]的回答被删除!",0,0)
			End If
			
			Dim strReturnURL
			strReturnURL = "q.asp?id=" & topicid & ""
			Response.Write "<script language=""JavaScript"">"
			Response.Write "try{top.location='" & strReturnURL & "';"
			Response.Write "}catch(e){}"
			Response.Write "</script>"
		End Sub
		

		
		Sub AddReward()
			If LockTopic > 0 Or TopicMode=1 Then
				Response.Write "<script>alert('友情提示!\n\n问题已处理不能进行此项操作!');top.location.replace(document.referrer);</script>"
				Response.End
			End If
			Dim RewardPoints,NeedPoint
			RewardPoints = KS.ChkClng(Request.Form("points"))
			If UserReward > 300 Then
				Response.Write "<script>alert('友情提示!\n\n您的问题悬赏已高于300分不能再提高悬赏!');top.location.replace(document.referrer);</script>"
				Response.End
			End If
			If RewardPoints > KS.ChkClng(KSUser.Score) Then
				Response.Write "<script>alert('亲爱的用户:\n\n您的积分不够,不能提高悬赏分!');history.back();</script>"
				Response.End
			End If
			If RewardPoints = 0 Then
				Response.Write "<script>alert('友情提示!\n\n请选择你要增加的悬赏分!');history.back();</script>"
				Response.End
			End If
			NeedPoint = RewardPoints + UserReward
			Conn.Execute ("UPDATE KS_AskTopic SET Reward=" & NeedPoint & " WHERE topicid="&topicid&" and username='"& KSUser.UserName &"' and Closed=0 and LockTopic=0")
			If RemainDays<0 Then
				Conn.Execute ("UPDATE KS_AskTopic SET ExpiredTime=" & SQLNowString & "+5,expired=0 WHERE topicid="&topicid&" and username='"& KSUser.UserName &"' and Closed=0 and LockTopic=0")
			ElseIf RemainDays < 5 Then
				Conn.Execute ("UPDATE KS_AskTopic SET ExpiredTime=ExpiredTime+5 WHERE topicid="&topicid&" and username='"& KSUser.UserName &"' and Closed=0 and LockTopic=0")
			End If
			
			If RewardPoints>0 Then
				 Call KS.ScoreInOrOut(KSUser.UserName,2,RewardPoints,"系统","问吧问题[" & Title & "]提高悬赏分!",0,0)
			End If
			
			Dim strReturnURL,Direct
			Direct = KS.ChkClng(Request.Form("direct"))
			Response.Write "<script language=""JavaScript"">"
			If Direct = 0 Then Response.Write "alert('提高问题悬赏分成功!');"
			Response.Write "try{top.location='q.asp?id=" & topicid & "';"
			Response.Write "}catch(e){}"
			Response.Write "</script>"
		End Sub
		
		Sub Nosatis()
			If LockTopic > 0 Or TopicMode=1 Then
				Response.Write "<script>alert('友情提示!\n\n问题已处理不能进行此项操作!');top.location.replace(document.referrer);</script>"
				Response.End
			End If
			Conn.Execute ("UPDATE KS_AskTopic SET Closed=1 WHERE topicid="&topicid&" and username='"& KSUser.UserName & "'")
			Dim strReturnURL,Direct
			Direct = KS.ChkClng(Request.Form("direct"))
			Response.Write "<script language=""JavaScript"">"
			If Direct = 0 Then Response.Write "alert('问题已成功关闭!');"
			Response.Write "try{top.location='q.asp?id=" & topicid & "';"
			Response.Write "}catch(e){}"
			Response.Write "</script>"
		End Sub
		
		Sub showmain()
			Dim SQL,Rs,Node
				SQL="SELECT TopicID,classid,classname,title,Username,Expired,Closed,PostTable,DateAndTime,LastPostTime,ExpiredTime,LockTopic,Reward,Hits,PostNum,CommentNum,TopicMode,Broadcast,Anonymous,supplement FROM KS_AskTopic WHERE topicid="&topicid&" and Username='"& KSUser.UserName &"'"
			Set Rs = Conn.Execute(SQL)
			If Rs.BOF And Rs.EOF Then
				Set Rs = Nothing
				Response.Write "<script>alert('友情提示!\n\n参数传递出错!');top.location.replace(document.referrer);</script>"
				Response.End()
			End If
			Set XMLDom = KS.RsToxml(Rs,"topic","xml")
			Set Rs = Nothing
			Set Node = XMLDom.documentElement.selectSingleNode("topic")
			If Node.selectSingleNode("@closed").text="1" Then
				Response.Write "<script>alert('友情提示!\n\n对不起,该问题已关闭!');top.location.replace(document.referrer);</script>"
				Response.End()
			End if
			topicid = CLng(Node.selectSingleNode("@topicid").text)
			title = Node.selectSingleNode("@title").text
			classid = CLng(Node.selectSingleNode("@classid").text)
			TopicUseTable = Trim(Node.selectSingleNode("@posttable").text)
			UserReward = CLng(Node.selectSingleNode("@reward").text)
			ExpiredTime = CDate(Node.selectSingleNode("@expiredtime").text)
			supplement = CLng(Node.selectSingleNode("@supplement").text)
			LockTopic = CLng(Node.selectSingleNode("@locktopic").text)
			TopicMode = CLng(Node.selectSingleNode("@topicmode").text)
			PostNum = CLng(Node.selectSingleNode("@postnum").text)
			PostUserName=Node.selectSingleNode("@username").text
			RemainDays = DateDIff("d",Now(),ExpiredTime)
			If CLng(Node.selectSingleNode("@closed").text) = 1 Then
				Response.Write "<script>alert('友情提示!\n\n问题已关闭不能进行此项操作!');top.location.replace(document.referrer);</script>"
				Exit Sub
			End If
			If PostUserName <> KSUser.UserName Then
					Response.Write "<script>alert('友情提示!\n\n错误的系统参数!');top.location.replace(document.referrer);</script>"
					Exit Sub
			End If
			Set Node = Nothing
		End Sub
		
		Sub main()
			%>
			<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
			<html>
			<head>
			<meta http-equiv="Content-Type" content="text/html; charset=GB2312" />
			<style type="text/css" media="all">
			body,td,input,select,textarea,a,div{font:12px Verdana, Arial, 宋体, sans-serif;color:#000;text-decoration:none;line-height:16px;}
			body{background:#fff;margin:0 auto;}
			li{list-style:none;padding:0;line-height:20px;}
			form{margin:0;padding:0;}
			h1,h2,h3,h4,h5,h6 {font-family:Verdana;font-size:12px;font-weight:400;}
			.mainBody {background:#fff;border-top:1px solid #b5cfe8;line-height:18px;margin-bottom:8px;}
			.mainBody h2 {clear:both;margin:0;letter-spacing:2px;height:22px;line-height:22px;background:#e7f5ff;color:#004299;text-align:center;font-weight:bold;}
			</style>
			<script language="JavaScript">
			<!--
			function TB_closeWindow(){
				try{
					window.parent.Thickenbox.tb_close();
				}
				catch(e){
					window.parent.location.replace(document.referrer);
				}
			}
			//-->
			</script>
			</head>
			<body>
			<div class="mainBody">
                <%
				Select Case LCase(Action)
				  Case "1"  showaddquestion
				  Case "2" showadvancereward
				  Case "3" ShowNoSatisAnswer
				End Select
				%>
             </div>
			 </body>
			 </html>
			<%
		End Sub
		
		Sub showaddquestion()
		%>
		<form method="post" action="handle.asp?action=saveadd">
		  <input type="hidden" name="TopicID" value="<%=topicid%>" /> 
		  <input type="hidden" name="direct" id="direct" value="1" /> 
		 <table width="100%" cellpadding="5" cellspacing="3" border="0">
		 <tr>
		 <td>
		  补充问题：<font color="blue"><%=title%> </font>
		  </td>
		  </tr>
		 <tr>
		 <td>
		 <textarea name="addContent" id="addContent" wrap="PHYSICAL" style="width:480px;height:100px;padding:4px;"><%=AddText%></textarea>
		  </td>
		  </tr>
		 <tr>
		 <td>
		  <input type="submit" name="addSubmit" id="addSubmit" value="提交问题补充" class="btn2" style="margin-right:10px;" /> 
		  </td>
		  </tr>
		  </table>
		  </form>
		<%
		End Sub
		
		Sub showadvancereward()
		  if ksuser.score<0 then 
		   ks.die "对不起,您的积分不足!您当前可用积分为<font color=red>" & KSUser.Score & "</font>。"
		  end if
		%>
		<form method="post" action="handle.asp?action=reward">
		  <input type="hidden" name="TopicID" value="<%=topicid%>" /> 
		  <input type="hidden" name="direct" id="direct" value="1" /> 
		 <table width="100%" cellpadding="5" cellspacing="3" border="0">
		  <tr>
		  <td>
		  问题： 
		  <font color="blue">
		  <%=title%>
		  </font>
		  </td>
		  </tr>
		  <tr>
		  <td>
		  增加悬赏分，可以提高问题的关注度，此时：（从现在开始计时） 
		  <br /> 
		  1）如果您的问题已过期，则系统自动增加5天有效期，如果距离结束时间不足5天，则系统自动把该问题的有效期增加至5天； 
		  <br /> 
		  2）如果您的问题距离结束时间尚多于5天，则截止日期不发生变化。 
		  </td>
		  </tr>
		 <tr>
		 <td>
		  增加悬赏分： 
		 <select name="points" id="points">
		  <option value="5">5分</option> 
		  <option value="10">10分</option> 
		  <option value="15">15分</option> 
		  <option value="20">20分</option> 
		  <option value="30">30分</option> 
		  <option value="50">50分</option> 
		  <option value="80">80分</option> 
		  <option value="100">100分</option> 
		  </select>
		  您当前的积分是 
		  <font color="red"><%=ksuser.score%>分</font> 
		  ，请设置您要增加的悬赏分 
		  </td>
		  </tr>
		 <tr>
		 <td>
		  <input type="submit" name="rewardSubmit" id="rewardSubmit" value="增加悬赏分" class="btn2" style="margin-right:10px;" /> 
		  </td>
		  </tr>
		  </table>
		  </form>

		<%
		End Sub
		
		Sub ShowNoSatisAnswer()
		%>
		 <form method="post" action="handle.asp?action=nosatis">
		  <input type="hidden" name="TopicID" value="<%=topicid%>" /> 
		  <input type="hidden" name="direct" id="direct" value="1" /> 
		 <table align="center" width="100%" cellpadding="5" cellspacing="3" border="0">
		 <tr>
		 <td height="45" align="center">
		 <br /><font color="#808080" size="3">
		  <b>如果没有满意的回答，请点击下面的“关闭问题”按钮直接结束提问！</b> 
		  </font>
		  <br /> 
		  <br /> 
		  出于保护回答者利益的考虑，问题的悬赏分将不再返还。 
		  <br /> 
		  如果因此造成您的积分损失，我们希望能得到您的理解。 
		  </td>
		  </tr>
		 <tr>
		 <td height="30" align="center">
		  <input type="submit" name="closeSubmit" id="closeSubmit" value="关闭问题" class="btn2" style="margin-right:10px;" /> 
		  <input type="button" name="Submit5" onclick="TB_closeWindow();" value="还是再等等吧" class="btn2" style="margin-right:10px;" /> 
		  </td>
		  </tr>
		  </table>
		  </form>

		<%
		End Sub
End Class
%>