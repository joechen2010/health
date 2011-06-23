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
        Private keyword,cid,topicmode,child,classname,parentstr,PageNum
		Private SqlStr,Topic,classarr,Catelist,CurrPage,totalPut,MaxPerPage,I,M
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
		   MaxPerPage=KS.ChkClng(KS.ASetting(14))
		   GetTopicList
		   showmain
		   set topic=nothing
		   set classarr=nothing
		End Sub
		
		Sub ShowMain()
			 Dim FileContent
			 FileContent = KSR.LoadTemplate(KS.ASetting(24))    
			 FCls.RefreshType = "asksearch" '设置刷新类型，以便取得当前位置导航等
			 FCls.RefreshFolderID = "0"   '设置当前刷新目录ID 为"0" 以取得通用标签
			 FileContent=KSR.KSLabelReplaceAll(FileContent)
			 Scan FileContent
		End Sub
		
		Sub GetQueryParam()
		  keyword=KS.CheckXSS(KS.S("keyword"))
		  If keyword="" Then 
		   Call KS.AlertHintScript("请输入关键字!")
		   Response.End()
		  End If
		  If KS.S("page") <> "" Then
			  CurrPage = CInt(Request("page"))
		  Else
			  CurrPage = 1
		  End If
		End Sub
				
		
		Sub GetTopicList()
		    Dim Param
			
			 Param="WHERE title like '%"& keyword & "%' And isTop=0 And LockTopic=0"
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
				Case "topic"
					 EchoTopicItem sTokenName
				Case "search"
				     EchoClassItem sTokenName
				Case "foot"
				     if sTokenName="showpage" then
				      echo KS.GetPrePageList(4,"条",PageNum,CurrPage,TotalPut,MaxPerPage)& KS.GetPageList("?"&KS.QueryParam("page"),4,CurrPage,PageNum, True)
					 end if
		    End Select 
        End Sub 
		
		Sub EchoClassItem(sTokenName)
		     Dim childclasslist,k
		     Select Case lcase(sTokenName)
			    case "shownav"
				  echo " &gt; 问题搜索"
				case "keyword" echo keyword
				case "totalnum" if totalput="" then echo "0" else echo totalput
			 End Select
		End Sub
		
		Sub EchoTopicItem(sTokenName)
		  Select Case sTokenName
		    Case "topicid" Echo Topic(0,i)
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

	    

End Class
%>
