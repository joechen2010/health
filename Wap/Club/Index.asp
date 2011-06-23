<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.ContentType="text/vnd.wap.wml"
Response.Charset="utf-8"
Response.Write "<?xml version=""1.0"" encoding=""utf-8""?>" &vbcrlf
Response.Write "<!DOCTYPE wml PUBLIC ""-//WAPFORUM//DTD WML 1.1//EN"" ""http://www.wapforum.org/DTD/wml_1.1.xml"">" &vbcrlf
%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
Dim KSCls
Set KSCls = New SiteIndex
KSCls.Kesion()
Set KSCls = Nothing
%>
<%
Class SiteIndex
        Private KS, KSR,ListStr,BSetting,Node
		Private LoopTemplate,LoopList,BoardID
		Private MaxPerPage, TotalPut , CurrentPage, TotalPage, i, j
	    Private KeyWord, SearchType,SqlStr
		Private Sub Class_Initialize()
		    If (Not Response.IsClientConnected)Then
			   Response.Clear
			   Response.End
		    End If
		    Set KS=New PublicCls
		    Set KSR = New Refresh
		End Sub
        Private Sub Class_Terminate()
		    Call CloseConn()
		    Set KS=Nothing
		End Sub
		Public Sub Kesion()
			If KS.Setting(56)="0" Then Call KS.ShowError("系统提示！","本站已关闭留言功能！")
			If KS.Setting(59)="1" Then 
			   Dim P:P=KS.QueryParam("page")
			   If P="" Then
			      Response.Redirect("GuestBook.asp?" & KS.WapValue & "")
			   Else
			      Response.Redirect("GuestBook.asp?" & P & "&amp;" & KS.WapValue & "")
			   End If
			End If
			KeyWord = KS.R(KS.S("keyword"))
			SearchType = KS.R(KS.S("SearchType"))
		    Dim FileContent
			BoardID=KS.ChkClng(KS.S("BoardID"))
			'FCls.RefreshType = "guestindex" '设置刷新类型，以便取得当前位置导航等
			'FCls.RefreshFolderID = BoardID '设置当前刷新目录ID 为"0" 以取得通用标签
			If KS.S("page") <> "" Then
			   CurrentPage = CInt(Request("page"))
			Else
			   CurrentPage = 1
			End If
			KS.LoadClubBoard
				   If BoardID<>0 Or Request("pid")<>"" Then 
				    If Request("pid")<>"" Then
				     Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & request("pid") &"]")
					Else
				     Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]")
					End If
					If Node Is Nothing Then
					 KS.Die "非法参数!"
					End If
					 BSetting=Node.SelectSingleNode("@settings").text
					 FileContent=Replace(FileContent,"{$BoardRules}",Node.SelectSingleNode("@boardrules").text)
				   End If
				   If KS.IsNul(BSetting) Then BSetting="1$$$0$0$0$0$0$0$0$0$$$$$$$$$$$$$$$$"
				 	BSetting=Split(BSetting,"$")
					
					
			
			'MaxPerPage=KS.ChkClng(KS.Setting(51))'留言列表显示信息数
			MaxPerPage=5 '留言列表显示信息数
			If Conn.Execute("select id from KS_GuestBoard").EOF or BoardID<>0 or KS.S("KeyWord")<>"" or KS.S("Istop")="1" or KS.S("IsBest")="1" Then
			   '主题列表
			   If KS.WSetting(31)="" Then
			      Call KS.ShowError("系统提示！","请先到""WAP基本信息设置->模板绑定""进行模板绑定操作!")
			   End If
			   FileContent = KSR.LoadTemplate(KS.WSetting(31))'取出留言板首页模板
			   If Trim(FileContent) = "" Then FileContent = "模板不存在!"
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
			   
			   If InStr(FileContent, "{$Intro}") <> 0 Then
			      Dim RSG:Set RSG=Conn.Execute("select Note From KS_GuestBoard Where ID=" & BoardID & "")
				  If Not(RSG.BOF And RSG.EOF) Then
				     FileContent=Replace(FileContent,"{$Intro}",RSG("Note")&"<br/>")
				  Else
				     FileContent=Replace(FileContent,"{$Intro}","")
				  End If
				  Set RSG=Nothing
			   End If
			   FileContent=Replace(FileContent,"{$GetBreakUrl}","Index.asp?BoardID=" & KS.ChkClng(KS.S("BoardID")) & "&amp;IsBest=" & KS.ChkClng(KS.S("IsBest")) & "&amp;IsTop=" & KS.ChkClng(KS.S("IsTop")) & "&amp;" & KS.WapValue & "")
			Else
			   '版面分类
			   If KS.WSetting(31)="" Then
			      Call KS.ShowError("系统提示！","请先到""WAP基本信息设置->模板绑定""进行模板绑定操作!")
			   End If
			   FileContent = KSR.LoadTemplate(KS.WSetting(31))'取出留言板首页模板
			   If Trim(FileContent) = "" Then FileContent = "模板不存在!"
			   Call GetBoardList()
			   FileContent = Replace(FileContent,"{$GuestTitle}",KS.Setting(61))
			End If
			
			   FileContent=Replace(FileContent,"{$BoardID}",BoardID)
			   FileContent = Replace(FileContent,"{$GetClubList}",LoopList)
			   FileContent = Replace(FileContent,"{$PageStr}",PageList())
			'===========================================================================
			Dim Doc:set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			Doc.async = False
			Doc.setProperty "ServerHTTPRequest", True 
			Doc.load(GetMapPath&"/Config/guestbook.xml")
			Dim XMLDate:XMLDate=doc.documentElement.attributes.getNamedItem("date").text
			If DateDiff("d",xmldate,now)=0 Then
			   If KS.ChkClng(doc.documentElement.attributes.getNamedItem("todaynum").text)>KS.ChkClng(doc.documentElement.attributes.getNamedItem("maxdaynum").text) Then
			      Doc.documentElement.attributes.getNamedItem("maxdaynum").text=doc.documentElement.attributes.getNamedItem("todaynum").text
				  Doc.save(GetMapPath&"/Config/guestbook.xml")
			   End If
			Else
			   Doc.documentElement.attributes.getNamedItem("date").text=now
			   Doc.documentElement.attributes.getNamedItem("yesterdaynum").text=doc.documentElement.attributes.getNamedItem("todaynum").text
			   Doc.documentElement.attributes.getNamedItem("todaynum").text=0
			   Doc.save(GetMapPath&"/Config/guestbook.xml")
			End If
			'===========================================================================		
			FileContent = Replace(FileContent,"{$TodayNum}",doc.documentElement.attributes.getNamedItem("todaynum").text)'今日帖子
			FileContent = Replace(FileContent,"{$YesterDayNum}",doc.documentElement.attributes.getNamedItem("yesterdaynum").text)'昨日帖子
			FileContent = Replace(FileContent,"{$MaxDayNum}",doc.documentElement.attributes.getNamedItem("maxdaynum").text)'最高日帖子
			FileContent = Replace(FileContent,"{$TopicNum}",Conn.Execute("select Count(ID) from KS_GuestBook")(0))'主题
			FileContent = Replace(FileContent,"{$ReplayNum}",Conn.Execute("select Count(ID) from KS_GuestReply")(0))
			FileContent = Replace(FileContent,"{$UserNum}",Conn.Execute("select Count(UserID) from KS_User")(0))
			FileContent = Replace(FileContent,"{$NewUser}",Conn.Execute("select top 1 UserName from KS_User order by UserID desc")(0))
			FileContent = KSR.KSLabelReplaceAll(FileContent)'替换所有标签
		    Response.Write FileContent
		End Sub
		
		'列出版面
		Sub GetBoardList()
		  Dim LC,PNode,Node,Xml,Str,TStr,pid,Bparam,LastPost,LastPost_A
          Set Xml=Application(KS.SiteSN&"_ClubBoard")
		  pid=KS.ChkClng(Request("pid"))
		  If pid=0 Then Bparam="parentid=0" Else BParam="id=" & pid
		  If IsObject(xml) Then
			   For Each Pnode In Xml.DocumentElement.SelectNodes("row[@" & BParam & "]")
					 LC= "【<a href=""?pid=" & PNode.SelectSingleNode("@id").text & "&amp;" & KS.WapValue & """>" & PNode.SelectSingleNode("@boardname").text & "</a>】<br/>"
					 
					 'LC=replace(LC,"{$BoardName}",PNode.SelectSingleNode("@boardname").text)
					 'LC=replace(LC,"{$Intro}",PNode.SelectSingleNode("@note").text)
					 'If KS.IsNul(PNode.SelectSingleNode("@master").text) then
					 'LC=replace(LC,"{$Master}","暂无版主")
					 'else
					 'LC=replace(LC,"{$Master}",PNode.SelectSingleNode("@master").text)
					 'end if
					 'LC=replace(LC,"{$TotalSubject}",PNode.SelectSingleNode("@topicnum").text)
					 'LC=replace(LC,"{$TotalReply}",PNode.SelectSingleNode("@postnum").text)
					 'LC=replace(LC,"{$TodayNum}",PNode.SelectSingleNode("@todaynum").text)
                     
					 tstr=""
					 
				   For Each Node In Xml.DocumentElement.SelectNodes("row[@parentid=" & Pnode.SelectSingleNode("@id").text & "]")
					 str="[<a href=""?boardid=" &Node.SelectSingleNode("@id").text & "&amp;" & KS.WapValue & """>" &Node.SelectSingleNode("@boardname").text & "</a>]"
					 
					 If KS.IsNul(Node.SelectSingleNode("@master").text) then
					 else
					  str=str & "(版主:" & Node.SelectSingleNode("@master").text & ")"
					 end if

					 str= str & "<br/>版面介绍:" & Node.SelectSingleNode("@note").text & "<br/>"
					 
					 LastPost=Node.SelectSingleNode("@lastpost").text
					 If KS.IsNul(LastPost) Then
					 Else
					  LastPost_A=Split(LastPost,"$")
					  If LastPost_A(0)="0" or LastPost_A(2)="无" then
					  else
					  str=str & "最新帖:<a href=""display.asp?id=" & LastPost_A(0) & "&amp;" & KS.WapValue & """>" & LastPost_A(2) & "</a> 总主题:" & Node.SelectSingleNode("@topicnum").text & "帖 总帖子:" & Node.SelectSingleNode("@postnum").text & "帖 今日帖:" & Node.SelectSingleNode("@todaynum").text & "帖<br/><br/>"
					  end if
					 End If
					 LC=LC&str
				  Next
				  LoopList=LoopList & LC
			 Next
		  End If
		  
		End Sub

		'列出贴子
		Sub GetLoopList()
			Dim Param:Param=" where verific=1"
			If KS.ChkClng(KS.S("Istop"))=1 Then Param=Param & " And IsTop=1"
			If KS.ChkClng(KS.S("IsBest"))=1 Then Param=Param & " And IsBest=1"
			If BoardID<>0 Then Param=Param &" And BoardID=" & BoardID
			If KS.S("KeyWord")<>"" Then
			   If KS.S("SearchType")="1" Then
			      Param=Param & " And SubJect Like '%" & KS.S("KeyWord") & "%'"
			   Else
			      Param=Param & " And UserName='" & KS.S("KeyWord") & "'"
			   End If
			End If
			SqlStr = "SELECT * From KS_GuestBook " & Param &" ORDER BY IsTop Desc, LastReplayTime Desc,ID DESC" 
			Dim RS:Set RS=KS.InitialObject("ADODB.RECORDSET")
			RS.Open SqlStr,Conn,1,1
			IF RS.EOF And RS.BOF Then
			   Totalput=0
			   LoopList = "没有留言!<br/>"
			   Exit Sub
			Else
			   TotalPut= RS.RecordCount
			   If CurrentPage < 1 Then CurrentPage = 1
			   If (CurrentPage - 1) * MaxPerPage > TotalPut Then
			      If (TotalPut Mod MaxPerPage) = 0 Then
				     CurrentPage = TotalPut \ MaxPerPage
				  Else
				     CurrentPage = TotalPut \ MaxPerPage + 1
				  End If
			   End If
			   If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < TotalPut Then
			      RS.Move (CurrentPage - 1) * MaxPerPage
			   Else
			      CurrentPage = 1
			   End If
			   Call GetTopicList(RS)
			End IF
			RS.Close:Set RS=Nothing
		End Sub
		
		Sub GetTopicList(RS)
		   ' On Error Resume Next
		    If Request("keyword")="" Then
			 LoopList= LoopList &"您当前正在浏览【<a href=""?boardid=" & Node.SelectSingleNode("@id").text & "&amp;" & KS.WapValue & """>" & Node.SelectSingleNode("@boardname").text & "</a>】版面,版主:" & Node.SelectSingleNode("@master").text & "<br/>"
			 Else
			  LoopList= LoopList &"【搜索结果】<br/>"
			 End If
		    Dim I,LC,SplitTF
		    Dim ATF:ATF=True
		 	Do While Not RS.EOF
			   If Not Response.IsClientConnected Then Response.End
			   
			   Dim IcoUrl
			   If RS("IsTop")=1 Then
			      IcoUrl="[顶]"
			   ElseIf RS("hits")>KS.ChkClng(KS.Setting(58)) Then
			      SplitTF=True
			      IcoUrl="[热]"
			   Else
			      SplitTF=True
			      IcoUrl=""
			  End If
			   
			   LC=(((I+1)+CurrentPage*MaxPerPage)-MaxPerPage) & "." & IcoUrl
			   'LC=Replace(LC,"{$TopicID}",RS("ID"))
			   LC=LC & "<a href=""display.asp?id=" & RS("ID") & "&amp;" & KS.WapValue & """>"
			   If RS("IsBest")=1 Then
			      LC=LC & RS("SubJect") & "</a><img src=""../Images/Jing.gif"" alt="".""/>"
			   Else
			      LC=LC & RS("SubJect") & "</a>"
			   End If
			   LC=LC & "(作者:" & RS("UserName") & " 浏览/回复:" & RS("Hits") & "/" & RS("TotalReplay") & "次 发表时间:" & rs("AddTime") & ")"
			  ' LC=Replace(LC,"{$Hits}",RS("hits"))
			  ' LC=Replace(LC,"{$PubTime}",RS("AddTime"))
			  ' LC=Replace(LC,"{$ReplayTimes}",RS("TotalReplay"))
			  ' LC=Replace(LC,"{$LastReplayUser}",RS("LastReplayUser"))
			  ' LC=Replace(LC,"{$LastReplayTime}",RS("LastReplayTime"))
			    LC=LC & "<br/>"
			  If CurrentPage=1 And SplitTF=True And ATF=True Then
			     ATF=False
				 LoopList=LoopList & LC
			  Else
			     LoopList=LoopList & LC
			  End If
			  I=I+1
			  If I >= MaxPerPage Then Exit Do
			  RS.MoveNext
		    Loop
			  LoopList=LoopList & PageList
			 If Request("keyword")="" Then
			  If Node.SelectSingleNode("@boardrules").text<>"" Then
			  LoopList=LoopList & "<br/>【本版版规】<br/>" & Node.SelectSingleNode("@boardrules").text & "<br/>"
			  End If
			 End If
		End Sub
		
		Function PageList()
		    PageList= KS.ShowPagePara(TotalPut, MaxPerPage, "Index.asp", True, "条", CurrentPage, "IsBest=" & KS.S("IsBest")&"&amp;IsTop=" & KS.S("Istop") & "&amp;KeyWord=" & KeyWord &"&amp;SearchType=" & SearchType & "&amp;BoardID=" & BoardID & "&amp;" & KS.WapValue & "")
		End Function
					  
End Class
%>
