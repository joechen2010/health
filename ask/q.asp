<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="function.asp"-->
<!--#include file="../KS_Cls/template.asp"-->
<%

Dim KSCls
Set KSCls = New Ask_Show_List
KSCls.Kesion()
Set KSCls = Nothing

Class Ask_Show_List
        Private classid,topicid,cid,topicmode,child,classname,parentstr
		Private SqlStr,Answer,classarr,Catelist,currpage,totalPut,MaxPerPage,I,PageNum
        Private KS, KSR,KSUser,UserLoginTF,BestID,Expired,Anonymous
		Private CloseTopic,XMLDom,PostNum,ExpiredTime,CommentNum,HeadTitle,TopicUseTable,RemainDays,RemainHour,icons,DateAndTime,Reward,PostUserName,Hits
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		  Set KSUser=New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		%>
		<!--#include file="../KS_Cls/Kesion.IFCls.asp"-->
		<%
		Public Sub Kesion()
		   GetQueryParam
		   UserLoginTF=Cbool(KSUser.UserLoginChecked)
		   LoadTopic
		   GetListParam
		   LoadQuestionList
		   showmain
		   set Answer=nothing
		   set classarr=nothing
		End Sub
		
		Sub ShowMain()
			 Dim FileContent
			 FileContent = KSR.LoadTemplate(KS.ASetting(23))    
			 FCls.RefreshType = "asklist" '设置刷新类型，以便取得当前位置导航等
			 FCls.RefreshFolderID = "0"   '设置当前刷新目录ID 为"0" 以取得通用标签
			 FileContent=KSR.KSLabelReplaceAll(FileContent)
			 FileContent=FileContent
			 Immediate=false
			 Scan FileContent
			 Response.write RexHtml_IF(Templates)
		End Sub
		
		Sub GetQueryParam()
		  topicid=KS.ChkClng(KS.S("id"))
		  If topicid=0 Then 
		   Call KS.AlertHintScript("对不起,非法参数!")
		   Response.End()
		  End If
		  If KS.S("page") <> "" Then
			  currpage = CInt(Request("page"))
		  Else
			  currpage = 1
		  End If
		End Sub
		
		Sub LoadTopic()
			If topicid = 0 Then Exit Sub
			Dim SQLStr,Rs,Node
			CloseTopic = 0
		
			SQLStr="SELECT top 1 TopicID,classid,classname,title,Username,Expired,Closed,PostTable,DateAndTime,LastPostTime,ExpiredTime,LockTopic,Reward,Hits,PostNum,CommentNum,TopicMode,Broadcast,Anonymous,supplement FROM KS_AskTopic WHERE topicid="&topicid & " and LockTopic=0"
			Set Rs = Conn.Execute(SQLStr)
			If Rs.BOF And Rs.EOF Then
				Set Rs = Nothing
				KS.AlertHintScript "参数传递出错或未通过审核!"
				Response.End
			End If
			Set XMLDom = KS.RsToxml(Rs,"topic","xml")
			Set Rs = Nothing
			Set Node = XMLDom.documentElement.selectSingleNode("topic")
			classid = CLng(Node.selectSingleNode("@classid").text)
			classname = Node.selectSingleNode("@classname").text
			topicmode = CLng(Node.selectSingleNode("@topicmode").text)
			PostNum = CLng(Node.selectSingleNode("@postnum").text)
			ExpiredTime = CDate(Node.selectSingleNode("@expiredtime").text)
			CommentNum = CLng(Node.selectSingleNode("@commentnum").text)
			HeadTitle = Trim(Node.selectSingleNode("@title").text)
			TopicUseTable = Trim(Node.selectSingleNode("@posttable").text)
			DateAndTime = Node.selectSingleNode("@dateandtime").text
			Reward = Node.selectSingleNode("@reward").text
			PostUserName = Node.selectSingleNode("@username").text
			Hits=Node.selectSingleNode("@hits").text
			Expired=Node.selectSingleNode("@expired").text
			Anonymous=Node.selectSingleNode("@anonymous").text
			RemainDays = DateDIff("d",Now(),ExpiredTime)
			RemainHour = DateDIff("h",Now(),ExpiredTime)
			RemainHour = RemainHour mod 24
			If RemainHour>0 Then RemainDays = RemainDays-1
			icons = topicmode
            
			'if topicmode=0 and datediff("s",ExpiredTime,now)>0 then
			if  datediff("s",ExpiredTime,now)>0 then
			expired=1
			Conn.Execute ("UPDATE KS_AskTopic SET expired=1 WHERE topicid="&topicid)
			end if
			If CLng(Node.selectSingleNode("@closed").text) = 1 Or CLng(Node.selectSingleNode("@commentnum").text) > 100 Then
				CloseTopic = 1
				icons = 5
			Else
				CloseTopic = 0
			End If
			If topicmode = 2 Then CloseTopic = 1
			Conn.Execute ("UPDATE KS_AskTopic SET Hits=Hits+1 WHERE topicid="&topicid)
			Set Node = Nothing
		End Sub
		
		Sub GetListParam()
		   If Not IsObject(Application(KS.SiteSN&"_askclasslist")) Then ACls.LoadCategoryList
		   Set Catelist = Application(KS.SiteSN&"_askclasslist")
		   If Not Catelist Is Nothing Then
			Dim Node:Set Node=Catelist.documentElement.selectSingleNode("row[@classid="&classid&"]")
			If Not Node Is Nothing Then
				classname=Node.selectSingleNode("@classname").text
				child=Node.selectSingleNode("@child").text
				parentstr=Node.selectSingleNode("@parentstr").text
				If child>0 Then
					cid=classid
				Else
					cid=CLng(Node.selectSingleNode("@parentid").text)
				End If 
		    End If
		   End If
		   MaxPerPage=KS.ChkClng(KS.ASetting(15))
		End Sub

		
		Sub LoadQuestionList()
		    Dim Param
			SQLStr="SELECT A.postsid,A.classid,A.TopicID,A.UserName,A.topic,A.content,A.addText,A.PostTime,A.DoneTime,A.star,A.satis,A.LockTopic,A.PostsMode,A.VoteNum,A.Plus,A.Minus,A.PostIP,A.Report,B.GradeTitle,B.userface FROM ["&TopicUseTable&"] A Left Join KS_User B On A.UserName=B.UserName WHERE A.topicid="&topicid&" and A.LockTopic=0 ORDER BY a.satis desc, A.postsid ASC"
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open SQLStr,Conn,1,1
			If Not RS.Eof Then
			                TotalPut= rs.recordcount
							If currpage < 1 Then currpage = 1
							if (TotalPut mod MaxPerPage)=0 then
								PageNum = TotalPut \ MaxPerPage
							else
								PageNum = TotalPut \ MaxPerPage + 1
							end if
		
							If currpage >1 and (currpage - 1) * MaxPerPage < totalPut Then
									RS.Move (currpage - 1) * MaxPerPage
							Else
									currpage = 1
							End If
							Dim FieldNum:FieldNum=rs.fields.count
							Answer=RS.GetRows(MaxPerPage)
							
							'置换最佳答案位置
							Dim L:L=Ubound(Answer,2)
							if (L>=1) Then
							 If Answer(10,0)=1 Then
							    Dim A0,A1,I
								For I=0 To FieldNum-1
								  If i=0 Then
								    A0=Answer(0,0)
									A1=Answer(0,1)
								  Else
								    A0=A0 & "+#+" & Answer(i,0)
								    A1=A1 & "+#+" & Answer(i,1)
								  End If
								Next
								A0=Split(A0,"+#+")
								A1=Split(A1,"+#+")
								For I=0 To FieldNum-1
								  Answer(I,0)=A1(i)
								  Answer(I,1)=A0(i)
								Next
							 End If
							End If
			End If
			RS.Close:Set RS=Nothing
		End Sub
		
	
        Sub ParseArea(sTokenName, sTemplate)
			Select Case sTokenName
				Case "answerlist"
				    If IsArray(Answer) Then 
					 Dim LoopStart
					 If currpage=1 Then LoopStart=1 Else LoopStart=0
					 For i=LoopStart To Ubound(Answer,2) 
					      If Answer(10,i)=1 Then BestID=Answer(0,i)   '设定最佳答案ID
						  Scan sTemplate
					 Next 
					End If
			End Select 
        End Sub 
		
		Sub ParseNode(sTokenType, sTokenName)
			Select Case sTokenType
			    case "ask"   echo ACls.ReturnAskConfig(sTokenName)
				Case "topic"  EchoTopicItem sTokenName
				Case "class"   EchoClassItem sTokenName
				Case "answer"  EchoAnswerItem sTokenName
				Case "foot"
				     If KS.Asetting(16)="1" Then
				      Echo ShowPage()
					  Else
					  echo KS.GetPrePageList(4,"条",PageNum,CurrPage,TotalPut,MaxPerPage)& KS.GetPageList("?"&KS.QueryParam("page"),4,CurrPage,PageNum, True)
					  End If
		    End Select 
        End Sub 
		
		Sub EchoClassItem(sTokenName)
		     Dim childclasslist,k
		     Select Case lcase(sTokenName)
				case "classname" Echo classname
				case "classid" Echo classid
			    case "shownav"
				     Dim parentArr:parentArr=Split(parentstr,",")
					 If Not Catelist Is Nothing Then
						 For k=0 To Ubound(parentArr)-1
						  Dim Node:Set Node=Catelist.documentElement.selectSingleNode("row[@classid="&parentArr(k)&"]")
							  If KS.ASetting(16)="1" Then
							  echo " &gt; <a href=""list-" & parentArr(k) & KS.ASetting(17) & """>" & Node.selectSingleNode("@classname").text & "</a>"
							  Else
							  echo " &gt; <a href=""showlist.asp?id=" & parentArr(k) & """>" & Node.selectSingleNode("@classname").text & "</a>"
							  End If
						 Next
					 End If
			 End Select
		End Sub
		
		Sub EchoTopicItem(sTokenName)
		  Select Case lcase(sTokenName)
		    Case "topicid" Echo topicid
		    Case "classname" Echo classname
		    Case "title" Echo HeadTitle
			Case "content" echo Answer(5,0)
			Case "gradetitle" 
			  If Anonymous<>1 Then
				  If KS.IsNul(Answer(18,0)) Then
				   echo "游客"
				  Else
				   echo Answer(18,0)
				  End If
			  End If
			Case "userface"
			  If KS.IsNul(Answer(19,0)) Then
			   echo "../images/face/0.gif"
			  ElseIf left(Lcase(Answer(19,0)),4)="http" Then 
			   echo answer(19,0)
			  Else
			   echo KS.Setting(3) & answer(19,0)
			  End If
			Case "addtext" if answer(6,0)<>"" and not isnull(answer(6,0)) then echo ks.htmlcode(answer(6,0))
			case "remaindays" echo RemainDays
			Case "username" if Anonymous=1 then Echo "匿名" else Echo PostUserName
			Case "time" Echo DateAndTime
			Case "hits" Echo Hits
			Case "status" Echo icons
			Case "firstanswerscore" echo KS.ASetting(30)
			Case "adoptedanswerscore" echo KS.ASetting(31)
			Case "reward"
			 If KS.ChkCLng(reward) > 0 Then
			   Echo " <img src=""images/ask_xs.gif"" width=""20"" height=""17"" />悬赏: " & Reward & " 金币</span> "
			 End If
		  End Select
		End Sub
		
		Sub EchoAnswerItem (sTokenName)
		 Select Case lcase(sTokenName)
		    case "postsid" echo Answer(0,i)
		    case "content" Echo Answer(5,i)
			case "time" echo answer(7,i)
			Case "gradetitle" 
			  If Answer(18,i)="" or Isnull(Answer(18,i)) Then
			   echo "游客"
			  Else
			   echo Answer(18,i)
			  End If
			Case "userface"
			  If trim(Answer(19,i))="" or Isnull(Answer(19,i)) Then
			   echo "../images/face/0.gif"
			  ElseIf left(Lcase(Answer(19,i)),4)="http" Then 
			   echo answer(19,i)
			  Else
			   echo KS.Setting(3) & answer(19,i)
			  End If
			Case "username" Echo answer(3,i)
		 End Select
		End Sub

	    '伪静态分页
		Public Function ShowPage()
		           Dim I, pageStr
				   pageStr= ("<div id=""fenye"" class=""fenye""><table border='0' align='right'><tr><td>")
					if (CurrPage>1) then pageStr=PageStr & "<a href=""show-" & topicid & "-" & CurrPage-1 & KS.ASetting(17) & """ class=""prev"">上一页</a>"
				   if (CurrPage<>PageNum) then pageStr=PageStr & "<a href=""show-" & TopicID & "-" & CurrPage+1 & KS.ASetting(17) & """ class=""next"">下一页</a>"
				   pageStr=pageStr & "<a href=""show-" & TopicID & "-1" & KS.ASetting(17) & """ class=""prev"">首 页</a>"
				 
					Dim startpage,n,j
					 if (CurrPage>=7) then startpage=CurrPage-5
					 if PageNum-CurrPage<5 Then startpage=PageNum-10
					 If startpage<0 Then startpage=1
					 n=0
					 For J=startpage To PageNum
						If J= CurrPage Then
						 PageStr=PageStr & " <a href=""#"" class=""curr""><font color=red>" & J &"</font></a>"
						Else
						 PageStr=PageStr & " <a class=""num"" href=""show-" & TopicID & "-" & J &KS.ASetting(17)&""">" & J &"</a>"
						End If
						n=n+1 : if n>=10 then exit for
					 Next
					
					 PageStr=PageStr & " <a class=""next"" href=""show-" & TopicID & "-" & PageNum & KS.ASetting(17)&""">末页</a>"
					 pageStr=PageStr & " <span>共" & totalPut & "条记录,分" & PageNum & "页</span></td></tr></table>"
				     PageStr = PageStr & "</td></tr></table></div>"
			         ShowPage = PageStr
	     End Function

End Class
%>
