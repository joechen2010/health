<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%Response.ContentType="text/vnd.wap.wml" %><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="发帖">
<p>
<%
Dim KSCls
Set KSCls = New GuestPost
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%
Class GuestPost
        Private KS
		Private LoginTF,Prev,BSetting,Node,TopicID
		Private Name, Subject, VerifyCode, IP, Pic, TxtHead, Memo, ErrorMsg, a, ID, BoardID
		Private Sub Class_Initialize()
		    If (Not Response.IsClientConnected)Then
			Response.Clear
			Response.End
			End If
			Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		    Call CloseConn()
			Set KS=Nothing
		End Sub
		
		Public Sub Kesion()
		    ID = KS.ChkClng(KS.S("ID"))
			If KS.Setting(56)="0" Then Response.Write "本站已关闭留言功能":Exit Sub
			
			KS.LoadClubBoard
			LoginTF=KSUser.UserLoginChecked
			If KS.Setting(57)="1" And LoginTF=false Then
			Response.Write "出错啦！<br/>很抱歉，你没有发布留言的权限！<br/><a href='../User/Login/'>点此登录</a>或<a href='../User/Reg/'>点此注册</a>新会员!<br/>"
			Else
			       If KS.ChkClng(Request("bid"))<>0 Then
				      Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & Request("bid") &"]")                  
					  If Node Is Nothing Then KS.Echo "非法参数!"
					  BSetting=Node.SelectSingleNode("@settings").text
				   End If
				   If KS.IsNul(BSetting) Then BSetting="1$$$0$0$0$0$0$0$0$0$$$$$$$$$$$$$$$$"
				   BSetting=Split(BSetting,"$")
				   
				   If BSetting(2)<>"" And KS.FoundInArr(BSetting(2),KSUser.GroupID,",")=false Then
 				     Response.Write "<br/>对不起,你没有此版面发表的权限!<br/>"
                   Else
					   Select Case KS.S("Action")
						   Case "SavePost"
						   Call SavePost()
						   Case "EditPost"'编辑
						   Call EditPost()
						   Case "SaveEditPost"
						   Call SaveEditPost()
						   Case "ConnectPost"'续写
						   Call ConnectPost()
						   Case "SaveConnectPost"
						   Call SaveConnectPost()
						   Case Else
						   Call ShowPost()
					   End Select
					   
					   If Prev=True Then
						  Response.Write "<anchor>返回上一页<prev/></anchor><br/>"
					   End If
					   Response.Write "<br/>"
					   If ID<>0 Then Response.Write "<a href=""Display.asp?ID=" & ID & "&amp;" & KS.WapValue & """>返回贴子</a><br/>" &vbcrlf
					  ' If KS.ChkClng(KS.S("BoardID"))=0 Then
						  'Dim TopicID:TopicID=Conn.Execute("select TopicID from KS_GuestReply Where ID=" & ID & "")(0)
					  '    Dim BoardID:BoardID=Conn.Execute("select BoardID from KS_GuestBook Where ID=" & ID & "")(0)
					  ' Else
						  BoardID=KS.ChkClng(KS.S("BID")) 
					   'End If
					   
					   if boardid<>0 then
					   Response.Write "<a href=""Index.asp?BoardID=" & BoardID & "&amp;" & KS.WapValue & """>"&Conn.Execute("select BoardName from KS_GuestBoard Where ID=" & BoardID & "")(0)&"</a><br/>" &vbcrlf
					   end if
					   
				End If
			   
			   Response.Write "<br/>"
			   Response.Write "<a href=""Index.asp?" & KS.WapValue & """>论坛首页</a>" &vbcrlf
			   Response.Write "&gt;&gt;" &vbcrlf
			   Response.Write "<a href=""" & KS.GetGoBackIndex & """>返回首页</a>" &vbcrlf
			End If
		End Sub
		
		Sub EditPost()
		    Dim RS:Set RS=Conn.Execute("Select * From KS_GuestBook Where ID="&ID&"")
			If RS.EOF Then
			   Response.write "非法参数!<br/>"
			   Prev=True
			   Exit Sub
			End If
			Response.Write "【帖子编辑】<br/>" &vbcrlf
			Response.Write "主题:<input name=""Subject" & Minute(Now) & Second(Now) & """ maxlength=""50"" value=""" & RS("Subject") & """ emptyok=""false""/><br/>" &vbcrlf
			Response.Write SelectBoard
			Response.Write "<br/>" &vbcrlf
			Response.Write "内容:<input name=""Memo" & Minute(Now) & Second(Now) & """ type=""text"" value=""" & Server.HTMLEncode(RS("Memo")) & """ emptyok=""false""/><br/>" &vbcrlf
			IF KS.Setting(53)=1 Then
			   Response.Write "附加码:<input name=""VerifyCode" & Minute(Now) & Second(Now) & """ type=""text"" value="""" emptyok=""false"" format=""*N""/>" & KS.GetVerifyCode &"请输入左边的数字<br/>" &vbcrlf
			End If
			Response.Write "<anchor>立即提交<go href=""Post.asp?Action=SaveEditPost&amp;ID=" & ID & "&amp;" & KS.WapValue & """ method=""post"">" &vbcrlf
			Response.Write "<postfield name=""Subject"" value=""$(Subject" & Minute(Now) & Second(Now) & ")""/>" &vbcrlf
            Response.Write "<postfield name=""BoardID"" value=""$(BoardID)""/>" &vbcrlf
            Response.Write "<postfield name=""VerifyCode"" value=""$(VerifyCode" & Minute(Now) & Second(Now) & ")""/>" &vbcrlf
            Response.Write "<postfield name=""Memo"" value=""$(Memo" & Minute(Now) & Second(Now) & ")""/>" &vbcrlf
            Response.Write "</go></anchor>" &vbcrlf
            Response.Write "<br/>" &vbcrlf
		End Sub
		
		Sub SaveEditPost()	
		    Dim I,SplitStrArr
			Dim LastLoginIP:LastLoginIP = KS.GetIP
			VerifyCode = KS.S("VerifyCode")
			IP = LastLoginIP
			Subject = KS.S("Subject")
			Memo = KS.S("Memo")
			BoardID=KS.ChkClng(KS.S("BoardID"))
			Memo=KS.FilterIllegalChar(Memo)
			A = CheckEnter()
			If Len(Replace(Memo,"&nbsp;",""))<=10 Then
			   A=false
			   ErrorMsg="内容不能少于10个字符！<br/>"
			End If
			If A = True Then 
		       Dim SqlStr:SqlStr = "SELECT top 1 * From KS_GuestBook Where ID=" & ID & "" 
			   Dim RSObj:Set RSObj=KS.InitialObject("Adodb.RecordSet")
			   RSObj.Open SqlStr,Conn,1,3
			   RSObj("Subject") = KS.HTMLEncode(Subject)
			   RSObj("Memo") = KS.HTMLEncode(Memo)
			   RSObj("GuestIP") = IP  
			   If KS.Setting(52)=1 Then  
			      RSObj("Verific")=0
			   Else
			      RSObj("Verific")=1
			   End If
			   RSObj("BoardID")=BoardID
			   RSObj("LastReplayTime")=Now
			   RSObj("TotalReplay")=0
			   RSObj("LastReplayUser")=KS.HTMLEncode(Name)
			   RSObj.Update
			   RSObj.Close
			   Set RSObj = Nothing
			   Response.Write "编辑成功！<br/>"
			Else
			   Response.Write ErrorMsg
			   Prev=True
			End If
		End Sub
		
		
		Sub ConnectPost()
		    Response.Write "【帖子续写】<br/>" &vbcrlf
		    Response.Write "内容:<input name=""Memo" & Minute(Now) & Second(Now) & """ type=""text"" value="""" emptyok=""false""/><br/>" &vbcrlf
			Response.Write "<anchor>续写确定<go href=""Post.asp?Action=SaveConnectPost&amp;ID=" & ID & "&amp;" & KS.WapValue & """ method=""post"">" &vbcrlf
			Response.Write "<postfield name=""Memo"" value=""$(Memo" & Minute(Now) & Second(Now) & ")""/>" &vbcrlf
			Response.Write "</go></anchor>" &vbcrlf
			Response.Write "<br/>" &vbcrlf
		End Sub
		
		Sub SaveConnectPost()
		    Dim RS:Set RS=Conn.Execute("Select Memo From KS_GuestBook Where ID="&ID&"")
			If RS.EOF Then
			   Response.write "非法参数!<br/>"
			   Prev=True
			   Exit Sub
			End If
			Dim Memo,SplitStrArr,I
			Memo=KS.FilterIllegalChar(KS.G("Memo"))
			If Memo="" Then
			   Response.write "出错提示，你没有输入续写内容！<br/>"
			   Prev=True
			   Exit Sub
			End If
			Dim RSObj:Set RSObj=KS.InitialObject("Adodb.Recordset")
			RSObj.Open "Select top 1 Memo From KS_GuestBook Where ID=" & ID,Conn,1,3
			RSObj("Memo")=Rs("Memo")&KS.HTMLEncode(Memo)
			RSObj.Update:RSObj.Close:Set RSObj=Nothing
			Response.write "续写成功。<br/>"
			Response.Write "<a href=""Display.asp?ID=" & ID & "&amp;" & KS.WapValue & """>返回该帖子</a><br/>"
			Set RS=Nothing
		End Sub
		
		Sub ShowPost()
		    Response.Write "【发表帖子】<br/>" &vbcrlf
			Response.Write SelectFont
			Response.Write "<br/>" &vbcrlf
			Response.Write "标题:<input name=""Subject" & Minute(Now) & Second(Now) & """ maxlength=""50"" value="""" emptyok=""false""/><br/>" &vbcrlf
			If BoardID=0 Then
			   Response.Write SelectBoard
			   Response.Write "<br/>" &vbcrlf
			End If
			Response.Write "内容:<input name=""Memo" & Minute(Now) & Second(Now) & """ type=""text"" value="""" emptyok=""false""/><br/>" &vbcrlf
			IF KS.Setting(53)=1 Then
			   Response.Write "附加码:<input name=""VerifyCode" & Minute(Now) & Second(Now) & """ type=""text"" value="""" emptyok=""false"" format=""*N""/>" & KS.GetVerifyCode &"请输入左边的数字<br/>" &vbcrlf
			End If
			Response.Write "<anchor>立即发表<go href=""Post.asp?Action=SavePost&amp;" & KS.WapValue & """ method=""post"">" &vbcrlf
			Response.Write "<postfield name=""Font"" value=""$(Font)""/>" &vbcrlf
			Response.Write "<postfield name=""Subject"" value=""$(Subject" & Minute(Now) & Second(Now) & ")""/>" &vbcrlf
			If BoardID=0 Then
            Response.Write "<postfield name=""BoardID"" value=""$(Bid)""/>" &vbcrlf
			Else
			Response.Write "<postfield name=""BoardID"" value=""" & BoardID & """/>" &vbcrlf
			End If
            Response.Write "<postfield name=""VerifyCode"" value=""$(VerifyCode" & Minute(Now) & Second(Now) & ")""/>" &vbcrlf
            Response.Write "<postfield name=""Memo"" value=""$(Memo" & Minute(Now) & Second(Now) & ")""/>" &vbcrlf
            Response.Write "</go></anchor>" &vbcrlf
            Response.Write "<br/>" &vbcrlf
		End Sub

	    Function SelectFont()
			SelectFont="话题:<select name=""Font"">"
			SelectFont=SelectFont&"<option value="""">[话题]</option>"
			SelectFont=SelectFont&"<option value=""[原创]"">[原创]</option>"
			SelectFont=SelectFont&"<option value=""[转帖]"">[转帖]</option>"
			SelectFont=SelectFont&"<option value=""[灌水]"">[灌水]</option>"
			SelectFont=SelectFont&"<option value=""[讨论]"">[讨论]</option>"
			SelectFont=SelectFont&"<option value=""[求助]"">[求助]</option>"
			SelectFont=SelectFont&"<option value=""[推荐]"">[推荐]</option>"
			SelectFont=SelectFont&"<option value=""[公告]"">[公告]</option>"
			SelectFont=SelectFont&"<option value=""[注意]"">[注意]</option>"
			SelectFont=SelectFont&"<option value=""[贴图]"">[贴图]</option>"
			SelectFont=SelectFont&"<option value=""[建议]"">[建议]</option>"
			SelectFont=SelectFont&"<option value=""[下载]"">[下载]</option>"
			SelectFont=SelectFont&"<option value=""[分享]"">[分享]</option>"
			SelectFont=SelectFont&"</select>"
	    End Function

	    Function SelectBoard()
		
		 If KS.Setting(59)="1" Then Exit Function
		 KS.LoadClubBoard()
		
		 dim str,xmls,nodes,XML,Node
         dim rs:set rs=conn.execute("select id,boardname from ks_guestboard where parentid=0 order by orderid")
		 if not rs.eof then set xml=KS.RsToXml(rs,"row",""):rs.close:set rs=nothing
           str="版面:<select name=""Bid"">"
		 If isobject(xml) then
		   for each node in xml.documentelement.selectnodes("row")
		   str=str & "<optgroup title=""" & node.selectsinglenode("@boardname").text &"""></optgroup>"
		        Set Xmls=Application(KS.SiteSN&"_ClubBoard")
				for each nodes in xmls.documentelement.selectnodes("row[@parentid=" & Node.SelectSingleNode("@id").text & "]")
				  if trim(request("bid"))=trim(Nodes.selectsinglenode("@id").text) then
				    str=str & "<option value=""" & Nodes.selectsinglenode("@id").text & """ selected=""selected"">--" & nodes.selectsinglenode("@boardname").text &"</option>"
				  else
				    str=str & "<option value=""" & Nodes.selectsinglenode("@id").text & """>--" & nodes.selectsinglenode("@boardname").text &"</option>"
				 end if
				next
		   next
		End If
           str=str & " </select>"		
		
		selectboard=str
	   End Function
		
		
		Sub SavePost()
		    Dim I,SplitStrArr
			Dim RefreshTime:RefreshTime = 30  '设置防刷新时间
			Dim LastLoginIP:LastLoginIP = KS.GetIP
			'Name = KS.S("Name")
			VerifyCode = KS.S("VerifyCode")
			IP = LastLoginIP
			'Pic = KS.S("Pic")
			'TxtHead = KS.S("TxtHead")
			Subject = KS.S("Subject")
			Memo = KS.FilterIllegalChar(KS.S("Memo"))
			BoardID=KS.ChkClng(KS.S("BoardID"))
			A = CheckEnter()
			If Len(replace(Memo,"&nbsp;",""))<=10 Then
			   A=false
			   ErrorMsg="留言内容不能少于10个字符！<br/>"
			End If
			If DateDiff("s", Session("SearchTime"), Now()) < RefreshTime Then
			   A=false
			   ErrorMsg="本页面起用了防刷新机制，请不要在"&RefreshTime&"秒内连续刷新本页面<br/>"
			End If
			If A = True Then 
			   SaveData()
			   Response.Write "发帖成功！<br/>"
			   Response.Write "<a href=""display.asp?id=" & TopicID & "&amp;" & KS.WapValue & """>进入帖子</a><br/>"
			   Session("SearchTime")=Now()
			Else
			   Response.Write ErrorMsg
			   Prev=True
			End If
		End Sub
		
		Function CheckEnter()
	        If LoginTF=False Then
			   Name="游客"
			Else  
			   Name=KSUser.UserName
			End if
			IF Trim(VerifyCode)<>Trim(Session("Verifycode")) And KS.Setting(53)=1 Then 
		   	   CheckEnter=False
			   ErrorMsg="验证码有误，请重新输入！<br/>"
			Else
			   If Subject="" Then
			      CheckEnter=False
				  ErrorMsg="请填写主题！"
			   End If
			   If Name="" Then
			      CheckEnter=False
				  ErrorMsg="你好像忘了填“昵称”！"
			   Else
			      'If Email="" or InStr(2,Email,"@")=0 Then
				     'CheckEnter=False
				     'ErrorMsg="你的Email有问题请重新填写！"
				  'Else
				      'If Pic="" Then
					      'CheckEnter=False
					      'ErrorMsg="你的头像没选,选一个把！"
					  'Else
					      'If TxtHead="" Then
						     'CheckEnter=False
						     'ErrorMsg="你的表情没选,选一个把！"
						  'Else
						      If Replace(Memo,"&nbsp;","")="" Then
							     CheckEnter=False
							     ErrorMsg="留言不能为空！<br/>"
							  Else
							     CheckEnter=True
							  End If
						  'End If
				      'End If
				  'End If	   
			   End If
			End If
		End Function

		Sub SaveData()
		
		    Dim O_LastPost,N_LastPost,O_LastPost_A,BSetting
		    If BoardID<>0 Then
			 KS.LoadClubBoard()
			 Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]")
			 O_LastPost=Node.SelectSingleNode("@lastpost").text
			 BSetting=Node.SelectSingleNode("@settings").text
			End If
			If KS.IsNul(BSetting) Then BSetting="1$$$0$0$0$0$0$0$0$0$$$$$$$$$$$$$$$$"
			BSetting=Split(BSetting,"$")
		
		
		    Dim SqlStr:SqlStr = "SELECT top 1 * From KS_GuestBook WHERE ID IS NULL" 
			Dim RSObj:Set RSObj=KS.InitialObject("Adodb.RecordSet")
			RSObj.Open SqlStr,Conn,1,3
			RSObj.AddNew 
			RSObj("UserName") = KS.HTMLEncode(Name)
			RSObj("Email") = ""
			RSObj("HomePage") = "http://www.kesion.com"
			RSObj("Face") = "1.gif"
			RSObj("TxtHead") = "Face1.gif"
			RSObj("Subject") = KS.HTMLEncode(Subject)
			RSObj("Memo") = KS.HTMLEncode(Memo)
			RSObj("Oicq") = ""
			RSObj("GuestIP") = IP  
			If KS.Setting(52)=1 Then  
			RSObj("Verific")=0
			Else
			RSObj("Verific")=1
			End If
			RSObj("AddTime") = Now()
			RSObj("Hits")=0
			RSObj("IsTop")=0
			RSObj("IsBest")=0
			RSObj("BoardID")=BoardID
			RSObj("LastReplayTime")=Now
			RSObj("TotalReplay")=0
			RSObj("LastReplayUser")=KS.HTMLEncode(Name)
			RSObj.Update
			RSObj.MoveLast
			TopicID=RSObj("ID")
			N_LastPost=RSObj("ID")&"$"& now & "$" & Replace(left(subject,200),"$","") & "$$$$"
			RSObj.Close
			Set RSObj = Nothing
			
			
			If KS.ChkClng(BSetting(3))>0 and LoginTF=true Then
				 Call KS.ScoreInOrOut(KSUser.UserName,1,KS.ChkClng(BSetting(3)),"系统","在论坛发表主题[" & Subject & "]所得!",0,0)
			End If
			If LoginTF=true Then
			  Call KSUser.AddLog(KSUser.UserName,"在论坛发表了主题[<a href='{$GetSiteUrl}club/display.asp?id=" & TopicID & "' target='_blank'>" & subject &"</a>]",100)
			End If
			
			
'更新今日发帖数等
			If BoardID<>0 Then
				Conn.Execute("Update KS_GuestBoard set lastpost='" & N_LastPost & "',postnum=postnum+1,topicnum=topicnum+1 where id=" & BoardID)
				If KS.IsNul(O_LastPost) Then
				  Conn.Execute("Update KS_GuestBoard set todaynum=1 where id=" & BoardID)
				  Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@todaynum").text=1
				Else
				 O_LastPost_A=Split(O_LastPost,"$")
				 Dim LastPostTime:LastPostTime=O_LastPost_A(1)
				 If Not IsDate(LastPostTime) Then LastPostTime=now
				 If datediff("d",LastPostTime,Now())=0 Then
				  Conn.Execute("Update KS_GuestBoard set todaynum=todaynum+1 where id=" & BoardID)
				  Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@todaynum").text=KS.ChkClng(Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@todaynum").text)+1
				 Else
				  Conn.Execute("Update KS_GuestBoard set todaynum=1 where id=" & BoardID)
				  Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@todaynum").text=1
				 End If
				End If
				 Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@postnum").text=KS.ChkClng(Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@postnum").text)+1
				 Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@topicnum").text=KS.ChkClng(Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@topicnum").text)+1
				 Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@lastpost").text=N_LastPost
		   End  If
			
			
			Dim Doc:set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			Doc.async = false
			Doc.setProperty "ServerHTTPRequest", true 
			Doc.load(GetMapPath&"/Config/guestbook.xml")
			Dim XMLDate:XMLDate=doc.documentElement.attributes.getNamedItem("date").text
					If DateDiff("d",xmldate,now)=0 Then
					  doc.documentElement.attributes.getNamedItem("todaynum").text=doc.documentElement.attributes.getNamedItem("todaynum").text+1
					  If KS.ChkClng(doc.documentElement.attributes.getNamedItem("todaynum").text)>KS.ChkClng(doc.documentElement.attributes.getNamedItem("maxdaynum").text) then
					   doc.documentElement.attributes.getNamedItem("maxdaynum").text=doc.documentElement.attributes.getNamedItem("todaynum").text
					  end if
					  
					Else
					  doc.documentElement.attributes.getNamedItem("date").text=now
					  doc.documentElement.attributes.getNamedItem("yesterdaynum").text=doc.documentElement.attributes.getNamedItem("todaynum").text
					  doc.documentElement.attributes.getNamedItem("todaynum").text=0
					End If
					  doc.documentElement.attributes.getNamedItem("topicnum").text=doc.documentElement.attributes.getNamedItem("topicnum").text+1
					  doc.documentElement.attributes.getNamedItem("postnum").text=doc.documentElement.attributes.getNamedItem("postnum").text+1
					  doc.save(GetMapPath&"/Config/guestbook.xml")
	  End Sub


	   Function ImageList()
	           dim i
			   for i=1 to 56 
			   ImageList=ImageList & "<option value=" & i & ">" & i & ".gif</option>"
			   next
	   End Function
	   
	   Function EmotList()
	        Dim I
			For I=1 To 30
			 EmotList=EmotList &  "<input type=""radio"" name=""txthead"" value=""" & I & """"
			  IF I=1 Then EmotList=EmotList &  " Checked"
				EmotList=EmotList &  " ><img src=""../Images/Face1/Face" & I & ".gif"" border=""0"">"
			  IF I Mod 15=0 Then EmotList=EmotList &  "<br/>"
			Next
	   End Function


End Class
%>
