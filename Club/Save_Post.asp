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
Set KSCls = New Guest_SaveData
KSCls.Kesion()
Set KSCls = Nothing

Class Guest_SaveData
        Private KS,KSUser,Node,LoginTF
        Private Name, Email, Subject, Oicq, Verifycode, IP, Pic, TxtHead, HomePage, Memo, ErrorMsg, a,BoardID,Purview,ShowIP,ShowSign,ShowScore
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
		If Request.servervariables("REQUEST_METHOD") <> "POST" Then
			Response.Write "<script>alert('�벻Ҫ�Ƿ��ύ��');</script>"
			Response.end
		End If
		If KS.IsNul(Request.ServerVariables("HTTP_REFERER")) Then
			Response.Write "<script>alert('�벻Ҫ�Ƿ��ύ��');</script>"
			Response.end
		End If
		if instr(lcase(Request.ServerVariables("HTTP_REFERER")),"post.asp")<=0 then
			Response.Write "<script>alert('�Ƿ��ύ��');</script>"
			Response.end
		end if
		
		
		LoginTF=KSUser.UserLoginChecked
		
	   If KS.Setting(57)="1" and LoginTF=false Then
	     KS.AlertHintScript "û�з���Ȩ��!"
	   End If
		
		
		Dim LastLoginIP:LastLoginIP = KS.GetIP
			Name = KS.S("Name")
			Email = KS.S("Email")
			HomePage = KS.S("HomePage")
			Oicq = KS.ChkClng(KS.S("Oicq"))
			Verifycode = KS.S("Code")
			IP = LastLoginIP
			Pic = KS.S("Pic")
			TxtHead = KS.S("txthead")
			Subject = KS.S("Subject")
			Memo = KS.CheckScript(Request.Form("Memo"))
			BoardID=KS.ChkClng(KS.S("BoardID"))
			Purview=KS.ChkClng(Request.Form("purview"))
			showip=KS.ChkClng(Request.Form("showip"))
			showsign=KS.ChkClng(Request.Form("showsign"))
			showscore=KS.ChkClng(Request.Form("showscore"))
			Memo=KS.FilterIllegalChar(memo)
		a = CheckEnter()
		If Len(replace(Memo,"&nbsp;",""))<=10 Then
		 a=false
		 ErrorMsg="�������ݲ�������10���ַ���"
		End If
		If a = True Then 
			SaveData()
			If KS.Setting(52)=1 Then   '������Ҫ���
			Response.Write("<script>alert('�����ɹ�,�������������˺�Ż���ʾ��');location.href='Index.asp?boardid=" & BoardID & "';</script>")
			Else
			Response.Redirect "Index.asp?boardid=" & BoardID
			End If
		Else
			Call KS.AlertHistory(ErrorMsg,-1)
		End If
	
	End Sub
	
	Function CheckEnter()
	        If KS.C("UserName")="" then
			  Name="�οͣ�" & Name
			Else
			  Name=KS.C("UserName")
			end if
			IF Trim(Verifycode)<>Trim(Session("Verifycode")) And KS.Setting(53)=1 then 
		   	 CheckEnter=False
			 ErrorMsg="��֤���������������룡"
			Else
			    If Subject="" Then
				   CheckEnter=False
				   ErrorMsg="����д���⣡"
				End If
				
				If KS.Setting(59)="1" Then 
					If Name="" Then
						CheckEnter=False
						ErrorMsg="�����������ǳơ���"
					Else
						If Email="" or InStr(2,Email,"@")=0 Then
							CheckEnter=False
							ErrorMsg="���Email��������������д��"
						Else
								If TxtHead="" Then
									CheckEnter=False
									ErrorMsg="��ı���ûѡ��"
								Else
									If replace(Memo,"&nbsp;","")="" Then
										CheckEnter=False
										ErrorMsg="���Բ���Ϊ�գ�"
									Else
										CheckEnter=TRUE
									End If
								End If
						End If	   
					End If
				Else
				  CheckEnter=TRUE
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
			
			if datediff("n",KSUser.RegDate,now)<KS.ChkClng(bsetting(9)) Then
			  Call KS.AlertHistory("�Բ���,����������" & bsetting(9) & "������ע��Ļ�Ա���ܷ���!",-1)
			  Response.End
			End if
			
			 Dim GroupPurview:GroupPurview= True : If Not KS.IsNul(BSetting(1)) and KS.FoundInArr(Replace(BSetting(1)," ",""),KSUser.GroupID,",")=false Then GroupPurview=false
			If (GroupPurview=false) Then
				KS.Die "<script>alert('�Բ���,��û���ڴ˰��淢����Ȩ��!');history.back();</script>"
			End If
			
		    Dim SqlStr:SqlStr = "SELECT top 1 * From KS_GuestBook WHERE ID IS NULL" 
			Dim RSObj:Set RSObj=Server.CreateObject("Adodb.RecordSet")
			RSObj.Open SqlStr,Conn,1,3
			RSObj.AddNew 
			RSObj("UserName") = KS.HTMLEncode(Name)
			RSObj("Email") = KS.HTMLEncode(Email)
			RSObj("HomePage") = KS.HTMLEncode(HomePage)
			if KS.Setting(59)="0" then
			 RSObj("Face") =Pic
			else
			 RSObj("Face") =KS.ChkClng(Pic)&".gif"
			end if
			RSObj("TxtHead") = "Face" &  TxtHead&".gif"
			RSObj("Subject") = KS.HTMLEncode(Subject)
			RSObj("Memo") = KS.HTMLEncode(Memo)
			RSObj("Oicq") = KS.HTMLEncode(Oicq)        
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
			RSObj("Purview")=Purview
			RSObj("ShowIP")=ShowIP
			RSObj("ShowSign")=ShowSign
			RSObj("ShowScore")=ShowScore
			RSObj("LastReplayTime")=Now
			RSObj("TotalReplay")=0
			RSObj("LastReplayUser")=KS.HTMLEncode(Name)
			RSObj.Update
			RSObj.MoveLast
			Dim TopicID:TopicID=RSObj("ID")
			N_LastPost=RSObj("ID")&"$"& now & "$" & Replace(left(subject,200),"$","") & "$$$$"
			RSObj.Close
			Set RSObj = Nothing
			
			'�����ϴ��ļ�
			Call KS.FileAssociation(1035,TopicID,Memo,0)
			
			If KS.ChkClng(BSetting(3))>0 and LoginTF=true Then
				 Call KS.ScoreInOrOut(KSUser.UserName,1,KS.ChkClng(BSetting(3)),"ϵͳ","����̳��������[" & Subject & "]����!",0,0)
			End If
			If LoginTF=true Then
			  Call KSUser.AddLog(KSUser.UserName,"����̳����������[<a href='{$GetSiteUrl}club/display.asp?id=" & TopicID & "' target='_blank'>" & subject &"</a>]",100)
			End If
			
			'���½��շ�������
			If BoardID<>0 Then
			    If KS.Setting(52)=1 Then   '������Ҫ���
				Conn.Execute("Update KS_GuestBoard set postnum=postnum+1,topicnum=topicnum+1 where id=" & BoardID)
				Else
				Conn.Execute("Update KS_GuestBoard set lastpost='" & N_LastPost & "',postnum=postnum+1,topicnum=topicnum+1 where id=" & BoardID)
				End If
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
End Class
%> 
