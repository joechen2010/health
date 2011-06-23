<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../Plus/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Guest_SaveData
KSCls.Kesion()
Set KSCls = Nothing

Class Guest_SaveData
        Private KS,KSUser,Node,BSetting
        Private UserName,Subject, Verifycode,TxtHead, Content, ErrorMsg,TopicID,BoardID,LoginTF,ShowIP,ShowSign
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser=New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
	   
	   Public Sub Kesion()
		Dim TmpIsSelfRefer,I,SplitStrArr
		TmpIsSelfRefer = IsSelfRefer()
			
		If TmpIsSelfRefer <> TRUE Then '�ⲿ�ύ������
			Call KS.AlertHistory("�����ύ����",-1)
		End If
		    LoginTF=KSUser.UserLoginChecked
			If KS.Setting(54)<>3 And LoginTF=false Then
			 Call KS.AlertHistory("�Բ�����û�з����Ȩ�ޣ�",-1)
			 Exit Sub
			ElseIf KS.Setting(54)=1 And KSUser.GroupID<>4 Then
			 Call KS.AlertHistory("�Բ��𣬱�վֻ���������Ա�ظ�!",-1)
			 Exit Sub
			ElseIf KS.Setting(54)=2 And LoginTF=False Then
			 Call KS.AlertHistory("�Բ��𣬱�վ����Ҫ���ǻ�Ա�ſ��Է���ظ���",-1)
			 Exit Sub
			End If
			
			If KS.Setting(57)="1" and LoginTF=false Then
				 KS.AlertHintScript "û�з���Ȩ��!"
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
			     KS.AlertHintScript "�����ڵ��û���,û�з���Ȩ��!"
			End If

			
			
			If LoginTF= True Then
			  UserName=KSUser.UserName
			Else
			  UserName="�ο�"
			End If
			TopicID = KS.ChkClng(KS.S("TopicID"))
			If len(Replace(KS.S("Content"),"&nbsp;",""))<=5 Then Call KS.AlertHistory("�ظ�������������5���ַ�!",-1)
			Content = Request.Form("Content")
			Content=KS.CheckScript(KS.HtmlCode(content))
			Content=KS.HtmlEncode(Content)
			ShowIP=KS.ChkClng(Request("showip"))
			ShowSign=KS.ChkClng(Request("showsign"))
			TxtHead = KS.S("TxtHead")
			Content=KS.FilterIllegalChar(Content)
		    If TopicID=0 Then Call KS.AlertHistory("�Ƿ�������",-1)
	        If Content="" Then Call KS.AlertHistory("��û������ظ�����!",-1)
			SaveData
			Response.Redirect "display.asp?id=" & TopicID&"&page=" &KS.S("Page")
			'Call KS.Alert("����ɹ���","display.asp?id=" & TopicID&"&page=" &KS.S("Page"))
	End Sub
		
	Sub SaveData()
			Dim O_LastPost,N_LastPost,O_LastPost_A
		    Dim SqlStr:SqlStr = "SELECT top 1 * From KS_GuestReply WHERE ID IS NULL" 
			Dim RSObj:Set RSObj=Server.CreateObject("Adodb.RecordSet")
			RSObj.Open SqlStr,Conn,1,3
			RSObj.AddNew 
			RSObj("UserName") = UserName
			RSObj("UserIP") = KS.GetIP
			RSObj("TopicID") = TopicID
			RSObj("Content") =Content
			RSObj("TxtHead")=TxtHead
			RSObj("ShowIp")=ShowIP
			RSObj("ShowSign")=ShowSign
			RSObj("ReplayTime") = Now
			If KS.Setting(60)="1" and Check=false Then  
			RSObj("Verific")=0
			Else
			RSObj("Verific")=1
			End If
			RSObj.Update
			RSObj.MoveLast
			Dim Rid:Rid=RSObj("id")
			RSObj.Close
			Set RSObj = Nothing
			'�����ϴ��ļ�
			Call KS.FileAssociation(1036,RID,Content,0)

			Dim Subject:Subject=Conn.Execute("Select top 1 subject From KS_GuestBook Where ID=" & TopicID)(0)
			
			Conn.Execute("Update KS_GuestBook Set LastReplayTime=" & SqlNowString &",LastReplayUser='" & UserName &"',TotalReplay=TotalReplay+1 where id=" & TopicID)
			
			N_LastPost=topicid & "$" & now & "$" & Replace(Subject,"$","") &"$$$$"
			
			If KS.ChkClng(BSetting(4))>0 and LoginTF=true Then
				 Call KS.ScoreInOrOut(KSUser.UserName,1,KS.ChkClng(BSetting(4)),"ϵͳ","����̳�ظ�����[" & Subject & "]����!",0,0)
			End If
			
             If LoginTF=true Then
			  Call KSUser.AddLog(KSUser.UserName,"����̳�ظ�������[<a href='{$GetSiteUrl}club/display.asp?id=" & TopicID & "' target='_blank'>" & subject &"</a>]",100)
			End If			
			
			'���°�������
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
			
			'���½��շ�������
			Dim Doc:set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			Doc.async = false
			Doc.setProperty "ServerHTTPRequest", true 
			Doc.load(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
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
			 doc.save(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
            Set Doc=Nothing
		End sub
		
		' ============================================
		' �����ҳ�Ƿ�ӱ�վ�ύ
		' ����:True,False
		' ============================================
		Function IsSelfRefer()
			Dim sHttp_Referer, sServer_Name
			sHttp_Referer = CStr(Request.ServerVariables("HTTP_REFERER"))
			sServer_Name = CStr(Request.ServerVariables("SERVER_NAME"))
			If Mid(sHttp_Referer, 8, Len(sServer_Name)) = sServer_Name Then
				IsSelfRefer = True
			Else
				IsSelfRefer = False
			End If
		End Function
		
		function check()
	 	Dim KSLoginCls,Master
		Set KSLoginCls = New LoginCheckCls1
		If KSLoginCls.Check=true Then
		  check=true
		  Exit function
		else
		    master=LFCls.GetSingleFieldValue("select master from ks_guestboard where id=" & KS.ChkClng(FCls.RefreshFolderID))
			Dim KSUser:Set KSUser=New UserCls
			If Cbool(KSUser.UserLoginChecked)=false Then 
			  check=false
			  exit function
			else
			   check=KS.FoundInArr(master, KSUser.UserName, ",")
			End If
		end if
 End function	
End Class
%> 
