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
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="论坛回复">
<p>
<%
Dim KSCls
Set KSCls = New ReplyCls
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%
Class ReplyCls
        Private KS,Prev,Action,BSetting,BoardID,Node
        Private UserName,TxtHead,Content,TopicID,LoginTF
		Private Master
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		    Call CloseConn()
		    Set KS=Nothing
		End Sub
	   
	   Public Sub Kesion()
		   LoginTF=KSUser.UserLoginChecked
		   Action=Lcase(Request("Action"))
		   Select Case Action
		       Case "savereply"
			   Call SaveReply()
			   Case "leadreply"
			   Call LeadReply()
			   Case Else
			   Call WriteReply()
		   End Select 
		   If Prev=True Then
		      Response.Write "<anchor>返回上一页<prev/></anchor><br/>"
		   End If
		   Response.Write "<br/>"
		   'Response.Write "<a href=""Index.asp?" & KS.WapValue & """>我的地盘</a><br/>" &vbcrlf
		   Response.Write "<a href=""" & KS.GetGoBackIndex & """>返回首页</a>" &vbcrlf
	   End Sub
	   
	   
	   Sub WriteReply()
	       TopicID = KS.ChkClng(KS.S("TopicID"))
		   Dim k,str:str="惊讶|撇嘴|色色|发呆|得意|流泪|害羞|闭嘴|睡觉|大哭|尴尬|发怒|调皮|呲牙|微笑|难过|超酷|非典|抓狂|吐血|"
		   Dim strArr:strArr=Split(str,"|")
		   Response.Write "<select name=""TxtHead"">" &vbcrlf
		   Response.Write "<option value=""0"">无</option>" &vbcrlf
		   For k=0 to 19
		       Response.Write "<option value=""" & (k+1) & """>" & strArr(k) & "</option>" &vbcrlf
		   Next
		   Response.Write "</select>" &vbcrlf
		   Dim reSayArry:reSayArry = Array("好帖，要顶!","看帖回帖是美德!","你牛，我顶!","这帖不错，该顶!","支持你!","反对你!")
		   Randomize
		   Response.Write "<input name=""Content" & Minute(Now) & Second(Now) & """ value=""" & reSayArry(Int(Ubound(reSayArry)*Rnd)) & """ maxlength=""300""/>" &vbcrlf
		   Response.Write "<anchor>回复<go href=""Reply.asp?Action=SaveReply&amp;TopicID=" & TopicID & "&amp;" & KS.WapValue & """ method=""post"">"
		   Response.Write "<postfield name=""TxtHead"" value=""$(TxtHead)""/>" &vbcrlf
		   Response.Write "<postfield name=""Content"" value=""$(Content" & Minute(Now) & Second(Now) & ")""/>" &vbcrlf
		   Response.Write "<postfield name=""Page"" value=""" & KS.S("Page") & """/>" &vbcrlf
		   Response.Write "</go></anchor>" &vbcrlf
		   Response.Write "<br/>" &vbcrlf
       End Sub
	   
	   Sub LeadReply()
	       TopicID = KS.ChkClng(KS.S("TopicID"))
	       Dim LeadID:LeadID = KS.ChkClng(KS.S("LeadID"))
		   Dim SqlStr:SqlStr = "SELECT * From KS_GuestReply where TopicID=" & TopicID & " And ID=" & LeadID & "" 
		   Dim RS:Set RS=KS.InitialObject("ADODB.RECORDSET")
		   RS.Open SqlStr,Conn,1,1
		   IF RS.EOF And RS.BOF Then
		      Response.Write "非法参数！<br/>"
			  Prev=True
		      Exit Sub
		   Else
		      'If RS("TxtHead")<>0 And Isnull(RS("TxtHead"))=False Then Response.Write "<img src=""../../Images/Face1/face" & RS("TxtHead") &".gif"" alt="".""/>"
		      Response.Write "以下是引用" & RS("UserName") & "在" & RS("ReplayTime") & "的发言：<br/>" &vbcrlf
			  Response.Write KS.UBBToHTML(KS.LoseHtml(KS.HTMLToUBB(KS.ReplaceTrim(Replace(KS.GetEncodeConversion(RS("Content")),"</div>","[br]")))))
			  Response.Write "<br/>" &vbcrlf
			  Response.Write "<br/>" &vbcrlf
			  Dim k,str:str="惊讶|撇嘴|色色|发呆|得意|流泪|害羞|闭嘴|睡觉|大哭|尴尬|发怒|调皮|呲牙|微笑|难过|超酷|非典|抓狂|吐血|"
			  Dim strArr:strArr=Split(str,"|")
			  Response.Write "<select name=""TxtHead"">" &vbcrlf
			  For k=0 to 19
			      Response.Write "<option value=""" & (k+1) & """>" & strArr(k) & "</option>" &vbcrlf
			  Next
			  Response.Write "</select>" &vbcrlf
			  Dim reSayArry:reSayArry = Array("好帖，要顶!","看帖回帖是美德!","你牛，我顶!","这帖不错，该顶!","支持你!","反对你!")
			  Randomize
			  Response.Write "<input name=""Content" & Minute(Now) & Second(Now) & """ value=""" & reSayArry(Int(Ubound(reSayArry)*Rnd)) & """ maxlength=""300""/>" &vbcrlf
			  Dim LeadContent:LeadContent=Server.HTMLEncode("以下是引用" & RS("UserName") & "在" & RS("ReplayTime") & "的发言：<br/>" & RS("Content") & "")
			  Response.Write "<anchor>回复<go href=""Reply.asp?Action=SaveReply&amp;TopicID=" & TopicID & "&amp;" & KS.WapValue & """ method=""post"">"
			  Response.Write "<postfield name=""TxtHead"" value=""$(TxtHead)""/>" &vbcrlf
			  Response.Write "<postfield name=""Content"" value=""$(Content" & Minute(Now) & Second(Now) & ")""/>" &vbcrlf
			  Response.Write "<postfield name=""LeadContent"" value=""" & LeadContent & """/>" &vbcrlf
			  Response.Write "<postfield name=""Page"" value=""" & KS.S("Page") & """/>" &vbcrlf
			  Response.Write "</go></anchor>" &vbcrlf
			  Response.Write "<br/>" &vbcrlf
		   End If
       End Sub
	   
	   
	   Public Sub SaveReply()
	       Dim RefreshTime,I,SplitStrArr
	       RefreshTime = 10  '设置防刷新时间
		   If KS.Setting(54)<>3 And LoginTF=false Then
		      Response.Write "对不起，你没有发表的权限！<br/>"
			  Prev=True
		      Exit Sub
		   ElseIf KS.Setting(54)=1 And KSUser.GroupID<>4 Then
			  Response.Write "对不起，本站只允许管理人员回复!<br/>"
			  Prev=True
		      Exit Sub
		   ElseIf KS.Setting(54)=2 And LoginTF=False Then
			  Response.Write "对不起，本站至少要求是会员才可以发表回复！<br/>"
			  Prev=True
		      Exit Sub
		   End If
		   
		   
			BoardID=KS.ChkClng(Request("BoardID"))
			If BoardID<>0 Then
			 KS.LoadClubBoard()
			 Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]")
			 BSetting=Node.SelectSingleNode("@settings").text
			End If
			If KS.IsNul(BSetting) Then BSetting="1$$$0$0$0$0$0$0$0$0$$$$$$$$$$$$$$$$"
			BSetting=Split(BSetting,"$")
			
			
			If BSetting(2)<>"" And KS.FoundInArr(BSetting(2),KSUser.GroupID,",")=false Then
			     KS.AlertHintScript "你所在的用户组,没有发表权限!"
			End If		   
		   
		   
		   If LoginTF= True Then
		      UserName=KSUser.UserName
		   Else
		      UserName="游客"
		   End If
		   TopicID = KS.ChkClng(KS.S("TopicID"))
		   If Len(Replace(KS.S("Content"),"&nbsp;",""))<=3 Then
		      Response.Write "回复字数不能少于5个字符!<br/>"
			  Prev=True
		      Exit Sub
		   End If
		   Content = KS.FilterIllegalChar(KS.HTMLEncode(KS.S("Content")))
		   TxtHead = KS.S("TxtHead")
		   If TopicID=0 Then
		      Response.Write "非法参数！<br/>"
			  Prev=True
		      Exit Sub
		   End If
		   If Content="" Then
		      Response.Write "你没有输入回复内容!<br/>"
			  Prev=True
		      Exit Sub
		   End If
		   If DateDiff("s", Session("SearchTime"), Now()) < RefreshTime Then
		      Response.Write "本页面起用了防刷新机制，请不要在"&RefreshTime&"秒内连续刷新本页面<br/>"
			  Prev=True
		      Exit Sub
		   End If
		   Dim LeadContent:LeadContent = KS.S("LeadContent")
		   If LeadContent<>"" Then
		      Content=LeadContent&"<br/>"&Content
		   End If
		   SaveData
		   Response.Write "发表成功！<br/>"
		   Session("SearchTime")=Now()
		   Response.Write "<a href=""Display.asp?ID=" & TopicID&"&amp;Page=" &KS.S("Page") & "&amp;" & KS.WapValue & """>返回上级</a><br/>"
	   End Sub
	   
	   Sub SaveData()
	       Dim O_LastPost,N_LastPost,O_LastPost_A
		   Dim SqlStr:SqlStr = "SELECT top 1 * From KS_GuestReply WHERE ID IS NULL" 
		   Dim RSObj:Set RSObj=KS.InitialObject("Adodb.RecordSet")
		   RSObj.Open SqlStr,Conn,1,3
		   RSObj.AddNew 
		   RSObj("UserName") = UserName
		   RSObj("UserIP") = KS.GetIP
		   RSObj("TopicID") = TopicID
		   RSObj("Content") =Content
		   RSObj("TxtHead")=TxtHead
		   RSObj("ReplayTime") = Now
			If KS.Setting(60)="1" and Check=false Then  
			RSObj("Verific")=0
			Else
			RSObj("Verific")=1
			End If
		   RSObj.Update
		   RSObj.Close
		   Set RSObj = Nothing
		   Dim Subject:Subject=Conn.Execute("Select top 1 subject From KS_GuestBook Where ID=" & TopicID)(0)
			
			Conn.Execute("Update KS_GuestBook Set LastReplayTime=" & SqlNowString &",LastReplayUser='" & UserName &"',TotalReplay=TotalReplay+1 where id=" & TopicID)
			
			N_LastPost=topicid & "$" & now & "$" & Replace(Subject,"$","") &"$$$$"
			
			If KS.ChkClng(BSetting(4))>0 and LoginTF=true Then
				 Call KS.ScoreInOrOut(KSUser.UserName,1,KS.ChkClng(BSetting(4)),"系统","在论坛回复主题[" & Subject & "]所得!",0,0)
			End If
			
             If LoginTF=true Then
			  Call KSUser.AddLog(KSUser.UserName,"在论坛回复了主题[<a href='{$GetSiteUrl}club/display.asp?id=" & TopicID & "' target='_blank'>" & subject &"</a>]",100)
			End If			
			
			'更新版面数据
			If BoardID<>0 Then
			  KS.LoadClubBoard()
			  O_LastPost=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@lastpost").text
			  
			  Conn.Execute("Update KS_GuestBoard set lastpost='" & N_LastPost & "',postnum=postnum+1 where id=" & BoardID)
				If KS.IsNul(O_LastPost) Then
				  Conn.Execute("Update KS_GuestBoard set todaynum=1 where id=" & BoardID)
				  Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@todaynum").text=1
				Else
				 O_LastPost_A=Split(O_LastPost,"$")
				 Dim LastPostDate:LastPostDate=O_LastPost_A(1)
				 If Not IsDate(LastPostDate) Then LastPostDate=Now
				 If datediff("d",LastPostDate,Now())=0 Then
				  Conn.Execute("Update KS_GuestBoard set todaynum=todaynum+1 where id=" & BoardID)
				  Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@todaynum").text=KS.ChkClng(Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@todaynum").text)+1
				 Else
				  Conn.Execute("Update KS_GuestBoard set todaynum=1 where id=" & BoardID)
				  Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@todaynum").text=1
				 End If
				End If
				 Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@postnum").text=KS.ChkClng(Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@postnum").text)+1
				 Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]/@lastpost").text=N_LastPost
			End If
			
			'更新今日发帖数等
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
            Set Doc=Nothing
	   End Sub
	   
	   Function Check()
		    If Cbool(KSUser.UserLoginChecked)=False Then 
			   Check=False
			   Exit Function
			Else
			   If KSUser.GroupID=4 Then
			      Check=True
				  Exit Function
			   Else
			      Master=LFCls.GetSingleFieldValue("select Master from KS_GuestBoard where ID=" & KS.ChkClng(FCls.RefreshFolderID))
				  Check=KS.FoundInArr(Master, KSUser.UserName, ",")
			   End If
			End If
		End Function
End Class
%> 
