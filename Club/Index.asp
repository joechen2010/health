<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New SiteIndex
KSCls.Kesion()
Set KSCls = Nothing

Class SiteIndex
        Private KS, KSR,ListStr,Node,BSetting,KSUser
		Private ListTemplate,pLoopTemplate,LoopTemplate,LoopList,boardid
		Private MaxPerPage, TotalPut , CurrentPage, TotalPage, i, j, Loopno
	    Private KeyWord, SearchType,SqlStr
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
		 Set KS=Nothing
		 Set KSUser=Nothing
		 Set KSR=Nothing
		End Sub
		
		%>
		<!--#include file="../KS_Cls/Kesion.IfCls.asp"-->
		<%
		
		Public Sub Kesion()
			If KS.Setting(56)="0" Then KS.Die "本站已关闭留言功能!"
			If KS.Setting(59)="1" Then 
			 Dim P:P=KS.QueryParam("page")
			 If P="" Then
		    	response.Redirect("guestbook.asp")
			 Else
		    	response.Redirect("guestbook.asp?" & P)
			 End If
			End If
			KeyWord = KS.R(KS.S("keyword"))
			SearchType = KS.R(KS.S("SearchType"))
		    Dim FileContent

		          If KS.Setting(114)="" Then KS.Die "请先到""基本信息设置->模板绑定""进行模板绑定操作!":response.end
				   FileContent = KSR.LoadTemplate(KS.Setting(114))
				   If Trim(FileContent) = "" Then FileContent = "模板不存在!"
				   BoardID=KS.ChkClng(KS.S("BoardID"))

				   FCls.RefreshType = "guestindex" '设置刷新类型，以便取得当前位置导航等
				   FCls.RefreshFolderID = BoardID '设置当前刷新目录ID 为"0" 以取得通用标签
				   
				   KS.LoadClubBoard
				   
				   If BoardID<>0 Or Request("pid")<>"" Then 
				    If Request("pid")<>"" Then
				     Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & KS.ChkClng(request("pid")) &"]")
					Else
				     Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]")
					End If
					If Node Is Nothing Then
					 KS.Die "非法参数!"
					End If
					 BSetting=Node.SelectSingleNode("@settings").text
					 FileContent=RexHtml_IF(FileContent) '列表页先过滤其它标签,减少标签解释
					 FileContent=Replace(FileContent,"{$BoardRules}",Node.SelectSingleNode("@boardrules").text)
				   End If
				   If KS.IsNul(BSetting) Then BSetting="1$$$0$0$0$0$0$0$0$0$$$$$$$$$$$$$$$$"
				 	BSetting=Split(BSetting,"$")


					If KS.S("page") <> "" Then
					  CurrentPage = CInt(Request("page"))
					Else
					  CurrentPage = 1
					End If
				   	MaxPerPage=KS.ChkClng(KS.Setting(51))

				   FileContent=Replace(FileContent,"{$PostBoardID}","?bid=" & BoardID)
				
				  
				If  BSetting(0)="0" And KS.C("UserName")="" Then
					ListTemplate=LFCls.GetConfigFromXML("GuestBook","/guestbook/template","error1")
				ElseIf boardid<>0 or KS.S("KeyWord")<>"" or KS.S("Istop")="1" or KS.S("IsBest")="1" then
				       KSUser.UserLoginChecked
					   If BSetting(1)<>"" and KS.FoundInArr(Replace(BSetting(1)," ",""),KSUser.GroupID,",")=false Then
					    ListTemplate=LFCls.GetConfigFromXML("GuestBook","/guestbook/template","error2")
					   Else
						   ListTemplate=LFCls.GetConfigFromXML("GuestBook","/guestbook/template","list")
						   LoopTemplate=KS.CutFixContent(ListTemplate, "[loop]", "[/loop]", 0)
						   Call GetLoopList()
						   if boardid<>0 or request("pid")<>"" Then
						   FileContent=Replace(FileContent,"{$GuestTitle}",Node.SelectSingleNode("@boardname").text)
						   else
							if KS.S("Istop")="1" then
							 FileContent=Replace(FileContent,"{$GuestTitle}","置顶帖子")
							else
							 FileContent=Replace(FileContent,"{$GuestTitle}","精华帖子")
							end if
						   end if
						   ListTemplate = Replace(ListTemplate,"[loop]" & LoopTemplate &"[/loop]",LoopList)
					 End If
				Else
				    Dim Doc:set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
					Doc.async = false
					Doc.setProperty "ServerHTTPRequest", true 
					Doc.load(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
					Dim XMLDate:XMLDate=doc.documentElement.attributes.getNamedItem("date").text
					If DateDiff("d",xmldate,now)=0 Then
					  If KS.ChkClng(doc.documentElement.attributes.getNamedItem("todaynum").text)>KS.ChkClng(doc.documentElement.attributes.getNamedItem("maxdaynum").text) then
					   doc.documentElement.attributes.getNamedItem("maxdaynum").text=doc.documentElement.attributes.getNamedItem("todaynum").text
					   doc.save(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
					  end if
					Else
					  Conn.Execute("Update KS_GuestBoard Set TodayNum=0")
					 ' Call KS.DelCahe(KS.SiteSN & "_ClubBoard")
				      Application(KS.SiteSN&"_ClubBoard")=empty	
					  doc.documentElement.attributes.getNamedItem("date").text=now
					  doc.documentElement.attributes.getNamedItem("yesterdaynum").text=doc.documentElement.attributes.getNamedItem("todaynum").text
					  doc.documentElement.attributes.getNamedItem("todaynum").text=0
					  doc.save(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
					End If
					
					FileContent=Replace(FileContent,"{$TodayNum}",doc.documentElement.attributes.getNamedItem("todaynum").text)
					FileContent=Replace(FileContent,"{$YesterDayNum}",doc.documentElement.attributes.getNamedItem("yesterdaynum").text)
					FileContent=Replace(FileContent,"{$MaxDayNum}",doc.documentElement.attributes.getNamedItem("maxdaynum").text)
					FileContent=Replace(FileContent,"{$TopicNum}",doc.documentElement.attributes.getNamedItem("topicnum").text)
					FileContent=Replace(FileContent,"{$ReplayNum}",doc.documentElement.attributes.getNamedItem("postnum").text)
					FileContent=Replace(FileContent,"{$UserNum}",conn.execute("select count(userid) from ks_user")(0))
					FileContent=Replace(FileContent,"{$NewUser}",conn.execute("select top 1 username from ks_user order by userid desc")(0))
				    KS.LoadClubBoard
				   
					   Call GetBoardList()
					   FileContent=Replace(FileContent,"{$GuestTitle}",KS.Setting(61))
					   ListTemplate = LoopList
				 end if
	               
				   FileContent=Replace(FileContent,"{$GetGuestList}",ListTemplate)
				   FileContent=Replace(FileContent,"{$PageStr}",PageList())
				   
				    
					
				   FileContent=KSR.KSLabelReplaceAll(FileContent)
				   KS.Echo RexHtml_IF(FileContent)
		End Sub
		'列出版面
		Sub GetBoardList()
		  Dim LC,PNode,Node,Xml,Str,TStr,pid,Bparam,LastPost,LastPost_A
          Set Xml=Application(KS.SiteSN&"_ClubBoard")
		  pid=KS.ChkClng(Request("pid"))
		  If pid=0 Then Bparam="parentid=0" Else BParam="id=" & pid
		  If IsObject(xml) Then
		       PLoopTemplate=LFCls.GetConfigFromXML("GuestBook","/guestbook/template","boardclass")
		       LoopTemplate=LFCls.GetConfigFromXML("GuestBook","/guestbook/template","board")
			   For Each Pnode In Xml.DocumentElement.SelectNodes("row[@" & BParam & "]")
					 LC=PLoopTemplate
					 LC=replace(LC,"{$BoardID}",PNode.SelectSingleNode("@id").text)
					 LC=replace(LC,"{$BoardName}",PNode.SelectSingleNode("@boardname").text)
					 LC=replace(LC,"{$Intro}",PNode.SelectSingleNode("@note").text)
					 If KS.IsNul(PNode.SelectSingleNode("@master").text) then
					 LC=replace(LC,"{$Master}","暂无版主")
					 else
					 LC=replace(LC,"{$Master}",PNode.SelectSingleNode("@master").text)
					 end if
					 LC=replace(LC,"{$TotalSubject}",PNode.SelectSingleNode("@topicnum").text)
					 LC=replace(LC,"{$TotalReply}",PNode.SelectSingleNode("@postnum").text)
					 LC=replace(LC,"{$TodayNum}",PNode.SelectSingleNode("@todaynum").text)
                     
					 tstr=""
					 
				   For Each Node In Xml.DocumentElement.SelectNodes("row[@parentid=" & Pnode.SelectSingleNode("@id").text & "]")
					 str=LoopTemplate
					 str=replace(str,"{$BoardID}",Node.SelectSingleNode("@id").text)
					 str=replace(str,"{$BoardName}",Node.SelectSingleNode("@boardname").text)
					 str=replace(str,"{$Intro}",Node.SelectSingleNode("@note").text)
					 If KS.IsNul(Node.SelectSingleNode("@master").text) then
					 str=replace(str,"{$Master}","暂无版主")
					 else
					 str=replace(str,"{$Master}",Node.SelectSingleNode("@master").text)
					 end if
					 
					 LastPost=Node.SelectSingleNode("@lastpost").text
					 If KS.IsNul(LastPost) Then
					  str=replace(str,"{$NewTopic}","无")
					 Else
					  LastPost_A=Split(LastPost,"$")
					  If LastPost_A(0)="0" or LastPost_A(2)="无" then
					  str=replace(str,"{$NewTopic}","无")
					  else
					  str=replace(str,"{$NewTopic}","<a href='display.asp?id=" & LastPost_A(0) & "'>" & KS.gottopic(LastPost_A(2),30) & "</a>")
					  end if
					 End If

					 str=replace(str,"{$TotalSubject}",Node.SelectSingleNode("@topicnum").text)
					 str=replace(str,"{$TotalReply}",Node.SelectSingleNode("@postnum").text)
					 str=replace(str,"{$TodayNum}",Node.SelectSingleNode("@todaynum").text)
					 TStr=TStr&str
				  Next
					LC=Replace(LC,"<!--boardlist-->",tstr)
				  LoopList=LoopList & LC
			 Next
		  End If
		End Sub
		'列出帖子
		Sub GetLoopList()
			Dim Param:Param=" where verific<>0"
			If KS.ChkClng(KS.S("Istop"))=1 Then Param=Param & " and istop=1"
			If KS.ChkClng(KS.S("IsBest"))=1 Then Param=Param & " and isbest=1"
			If BoardID<>0 Then Param=Param &" and boardid=" & boardid
			If KS.S("KeyWord")<>"" Then
			  If KS.S("SearchType")="1" Then
			   Param=Param & " and subject like '%" & KS.S("KeyWord") & "%'"
			  Else
			   Param=Param & " and username='" & KS.S("KeyWord") & "'"
			  End If
			End If
			 SqlStr = "SELECT * From KS_GuestBook " & Param &" ORDER BY IsTop Desc, LastReplayTime Desc,ID DESC" 
			 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			 RS.Open SqlStr,conn,1,1
			  IF RS.Eof And RS.Bof Then
				  totalput=0
				  LoopList = "<tr><td colspan=5>此版面没有" & KS.Setting(62) & "!</td></tr>"
				  exit sub
			  Else
								TotalPut= RS.RecordCount
								If CurrentPage < 1 Then CurrentPage = 1
			
								If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
								Else
										CurrentPage = 1
								End If
								Call GetTopicList(RS)
				End IF
				RS.Close:Set RS=Nothing
		End Sub
		
		Sub GetTopicList(RS)
		   On Error Resume Next
		   Dim I,LC,SplitTF
		   Dim ATF:ATF=True
		 	Do While Not RS.Eof
			 If Not Response.IsClientConnected Then Response.end
			  LC=LoopTemplate
			  LC=replace(LC,"{$TopicID}",rs("ID"))
			  If RS("IsBest")=1 Then
			  LC=replace(LC,"{$Subject}",rs("subject") & "<Img src='images/jing.gif' border='0'>")
			  else
			  LC=replace(LC,"{$Subject}",rs("subject"))
			  end if
			  If KS.IsNul(rs("username")) Then
			  LC=replace(LC,"{$Author}","游客")
			  Else
			  LC=replace(LC,"{$Author}",rs("username"))
			  End If
			  LC=replace(LC,"{$Hits}",RS("hits"))
			  LC=replace(LC,"{$PubTime}",RS("AddTime"))
			  LC=replace(LC,"{$ReplayTimes}",RS("TotalReplay"))
			  If KS.IsNul(RS("LastReplayUser")) Then
			  LC=replace(LC,"{$LastReplayUser}","游客")
			  Else
			  LC=replace(LC,"{$LastReplayUser}",RS("LastReplayUser"))
			  End If
			  LC=replace(LC,"{$LastReplayTime}",KS.GetTimeFormat(RS("LastReplayTime")))
			  Dim IcoUrl
			  If RS("IsTop")=1 Then
			   IcoUrl="top.gif"
			  ElseIf RS("hits")>KS.ChkClng(KS.Setting(58)) Then
			   SplitTF=true
			   IcoUrl="hot.gif"
			  Else
			   SplitTF=true
			   IcoUrl="common.gif"
			  End If
			  LC=Replace(LC,"{$Ico}","<a href='display.asp?id=" & rs("id") &"' target='_blank'><img border='0' src='images/" & IcoUrl & "' title='点击新窗口打开'></a>")


			  
			  If CurrentPage=1 and SplitTF=true and ATF=true Then
			    ATF=false
			    LoopList=LoopList &"<tr><td colspan=10 style='border-right:1px solid #E4E7EC;background:#FAFDFF;height:25px;padding-left:15px'>==普通帖子==</td></tr>" & LC
			  Else
			    LoopList=LoopList & LC
			  End If
	        I=I+1
			If I >= MaxPerPage Then Exit Do
			RS.MoveNext
		Loop

		End Sub
		

 
 Function PageList()
    PageList= "<table width=""100%"" aling=""center""><tr><td align=right>" & KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,false) & "</td></tr></table>"
 End Function
					  
End Class
%>
