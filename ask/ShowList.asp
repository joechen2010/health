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
        Private classid,cid,topicmode,child,classname,parentstr
		Private SqlStr,Topic,classarr,Catelist,CurrPage,totalPut,MaxPerPage,I,M,PageNum
        Private KS, KSR,KSUser,UserLoginTF
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		  Set KSUser=New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()
		   GetQueryParam
		   UserLoginTF=Cbool(KSUser.UserLoginChecked)
		   GetListParam
		   GetChildList
		   GetTopicList
		   showmain
		   set topic=nothing
		   set classarr=nothing
		End Sub
		
		Sub ShowMain()
			 Dim FileContent
			 FileContent = KSR.LoadTemplate(KS.ASetting(22))    
			 FCls.RefreshType = "asklist" '设置刷新类型，以便取得当前位置导航等
			 FCls.RefreshFolderID = "0"   '设置当前刷新目录ID 为"0" 以取得通用标签
			 FileContent=KSR.KSLabelReplaceAll(FileContent)
			 Scan FileContent
		End Sub
		
		Sub GetQueryParam()
		  classid=KS.ChkClng(KS.S("id"))
		  If classid=0 Then 
		   Call KS.AlertHintScript("对不起,非法参数!")
		   Response.End()
		  End If
		  m=KS.ChkClng(KS.S("m"))
		  If M<=2 and m<>0 Then topicmode=m-1
		  If KS.S("page") <> "" Then
			  CurrPage = CInt(Request("page"))
		  Else
			  CurrPage = 1
		  End If
		End Sub
		
		Sub GetListParam()
		   If Not IsObject(Application(KS.SiteSN&"_askclasslist")) Then ACls.LoadCategoryList
		   Set Catelist = Application(KS.SiteSN&"_askclasslist")
		   If Not Catelist Is Nothing Then
			Dim Node:Set Node=Catelist.documentElement.selectSingleNode("row[@classid="&classid&"]")
			classname=Node.selectSingleNode("@classname").text
			child=Node.selectSingleNode("@child").text
			parentstr=Node.selectSingleNode("@parentstr").text
			If child>0 Then
				cid=classid
			Else
				cid=CLng(Node.selectSingleNode("@parentid").text)
			End If 
			
		   End If
		   MaxPerPage=KS.ChkClng(KS.ASetting(14))
		End Sub
		
		Sub GetChildList()
		  Dim SQLStr:SQLStr = "SELECT classid,classname,AskPendNum,AskDoneNum FROM KS_AskClass WHERE parentid="&KS.ChkClng(cid)&" ORDER BY orders,classid"
		  Dim RS:Set RS=Conn.Execute(SQLStr)
		  If Not RS.Eof Then
		   classarr=RS.GetRows(-1)
		  End If
		  RS.Close:Set RS=Nothing
		End Sub
		
		Sub GetTopicList()
		    Dim Param
			If child > 0 Then
			   Param="WHERE classid in (SELECT classid FROM KS_AskClass WHERE ','+parentstr+'' like '%,"&classid&",%') And isTop=0 And LockTopic=0"
			Else
			   Param="WHERE classid="&classid & " And isTop=0 And LockTopic=0"
			End If
			If topicmode<>"" Then Param=Param &" and topicmode=" & topicmode
			If m=3 Then Param=Param & " and reward>0"
			SQLStr="SELECT TopicID,classid,classname,title,Username,Expired,Closed,DateAndTime,LastPostTime,LockTopic,Reward,Hits,PostNum,CommentNum,TopicMode,Highlight,Broadcast,Anonymous,IsTop FROM KS_AskTopic " & Param & " ORDER BY LastPostTime DESC"

			
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open SQLStr,Conn,1,1
			If Not RS.Eof Then
			                TotalPut= rs.recordcount
							If CurrPage < 1 Then CurrPage = 1
							if (TotalPut mod MaxPerPage)=0 then
								PageNum = TotalPut \ MaxPerPage
							else
								PageNum = TotalPut \ MaxPerPage + 1
							end if
		
							If CurrPage >1 and (CurrPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrPage - 1) * MaxPerPage
							Else
									CurrPage = 1
							End If
							Topic=RS.GetRows(MaxPerPage)
			End If
			RS.Close:Set RS=Nothing
		End Sub
		
	
        Sub ParseArea(sTokenName, sTemplate)
			Select Case sTokenName
				Case "topiclist"
				    If IsArray(Topic) Then 
					 For i=0 To Ubound(Topic,2) 
						Scan sTemplate
					 Next 
				    Else
					
					End If
			End Select 
        End Sub 
		
		Sub ParseNode(sTokenType, sTokenName)
			Select Case lcase(sTokenType)
			    case "ask"
				    echo ACls.ReturnAskConfig(sTokenName)
			    Case "head"
				    Echo "<li class="
					if m=0 then echo "curr" else echo "normal"
					 If KS.ASetting(16)="1" Then
					  echo "><a href=""list-" & classid & KS.ASetting(17) &""">全部问题</a>"
					 Else
					  echo "><a href=""?id=" & classid & """>全部问题</a>"
					 End If
				    Echo "<li class="
				    if m=1 then Echo  "curr" Else Echo "normal"
					 If KS.ASetting(16)="1" Then
					  Echo "><a href=""list-" & classid & "-0-1" & KS.ASetting(17) & """>待解决问题</a>"
					 Else
					  Echo "><a href=""?id=" & classid & "&m=1"">待解决问题</a></li>"
					 End If
					Echo "<li class="
					If m=2 then Echo "curr" Else Echo "normal"
					 If KS.ASetting(16)="1" Then
					  Echo "><a href=""list-" & classid & "-0-2" & KS.ASetting(17) & """>已解决</a>"
					 Else
					  Echo "><a href=""?id=" & classid & "&m=2"">已解决</a></li>"
					 End If
					Echo "<li class="
					If m=3 then Echo "curr" Else Echo "normal"
					 If KS.ASetting(16)="1" Then
					  Echo "><a href=""list-" & classid & "-0-3" & KS.ASetting(17) & """>悬赏分</a>"
					 Else
					  Echo "><a href=""?id=" & classid & "&m=3"">悬赏分</a></li>"
					 End If
				Case "topic"
					 EchoTopicItem sTokenName
				Case "class"
				     EchoClassItem sTokenName
				Case "foot"
				     if sTokenName="showpage" then
					  If KS.Asetting(16)="1" Then
				      Echo ShowPage()
					  Else
					  echo KS.GetPrePageList(4,"条",PageNum,CurrPage,TotalPut,MaxPerPage)& KS.GetPageList("?"&KS.QueryParam("page"),4,CurrPage,PageNum, True)
					  End If
					 end if
		    End Select 
        End Sub 
		
		Sub EchoClassItem(sTokenName)
		     Dim childclasslist,k
		     Select Case lcase(sTokenName)
				case "classname" Echo classname
				case "classid" Echo classid
				case "childclasslist"
				 If IsArray(classarr) Then
				    For k=0 To Ubound(classarr,2)
					  If KS.ASetting(16)="1" Then
					  echo "<li><a href=""list-" & classarr(0,k) & KS.ASetting(17) & """>" & classarr(1,k) & "(" & classarr(2,k)+classarr(3,k) & ")</a></li>"
					  Else
					  echo "<li><a href=""showlist.asp?id=" & classarr(0,k) & """>" & classarr(1,k) & "(" & classarr(2,k)+classarr(3,k) & ")</a></li>"
					  End If
					Next
				 End If	
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
		  Select Case sTokenName
		    Case "topicid" Echo Topic(0,i)
			Case "topicurl" If KS.ASetting(16)="0" Then Echo "q.asp?id=" & Topic(0,i) Else Echo "show-" & Topic(0,i) & KS.ASetting(17)
		    Case "classname" Echo Topic(2,i)
		    Case "title" Echo KS.Gottopic(Topic(3,i),30)
			Case "username" 
			 if Topic(17,i)=1 then Echo "匿名" else Echo Topic(4,i) 
			Case "time" Echo Topic(7,i)
			Case "postnum" Echo Topic(12,i)
			Case "status" Echo Topic(14,I)
			Case "reward"
			 If KS.ChkCLng(Topic(10,I)) > 0 Then
			   Echo " <span style=""color:#ff6600;font-weight:bold""><img src='images/ask_xs.gif' align='absmiddle'> " & Topic(10,I) & "</span>"
			 End If
		  End Select
		End Sub

	    '伪静态分页
		Public Function ShowPage()
		           Dim I, pageStr
				   pageStr= ("<div id=""fenye"" class=""fenye""><table border='0' align='right'><tr><td>")
					if (CurrPage>1) then pageStr=PageStr & "<a href=""list-" & classid & "-" & CurrPage-1 & "-" & m & KS.ASetting(17) & """ class=""prev"">上一页</a>"
				   if (CurrPage<>PageNum) then pageStr=PageStr & "<a href=""list-" & classid & "-" & CurrPage+1 & "-" & m & KS.ASetting(17) & """ class=""next"">下一页</a>"
				   pageStr=pageStr & "<a href=""list-" & classid & "-1-" & m & KS.ASetting(17) & """ class=""prev"">首 页</a>"
				 
					Dim startpage,n,j
					 if (CurrPage>=7) then startpage=CurrPage-5
					 if PageNum-CurrPage<5 Then startpage=PageNum-10
					 If startpage<0 Then startpage=1
					 n=0
					 For J=startpage To PageNum
						If J= CurrPage Then
						 PageStr=PageStr & " <a href=""#"" class=""curr""><font color=red>" & J &"</font></a>"
						Else
						 PageStr=PageStr & " <a class=""num"" href=""list-" & classid & "-" & J & "-" & m & KS.ASetting(17)&""">" & J &"</a>"
						End If
						n=n+1 : if n>=10 then exit for
					 Next
					
					 PageStr=PageStr & " <a class=""next"" href=""list-" & classid & "-" & PageNum & "-" & m & KS.ASetting(17)&""">末页</a>"
					 pageStr=PageStr & " <span>共" & totalPut & "条记录,分" & PageNum & "页</span></td></tr></table>"
				     PageStr = PageStr & "</td></tr></table></div>"
			         ShowPage = PageStr
	     End Function
End Class
%>
