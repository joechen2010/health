<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../../Plus/md5.asp"-->
<!--#include file="../../api/cls_api.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New User_RegPost
KSCls.Kesion()
Set KSCls = Nothing

Class User_RegPost
        Private KS,KSRFObj
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSRFObj=New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSRFObj=Nothing
		End Sub
		Public Sub Kesion()
		 'ϵͳ���ò���
		 Dim Locked:Locked=0
		 Dim VerificCodeTF:VerificCodeTF=KS.Setting(27)
		 Dim EmailMultiRegTF:EmailMultiRegTF=KS.Setting(28)
		 Dim UserNameLimitChar:UserNameLimitChar=Cint(KS.Setting(29))
		 Dim UserNameMaxChar:UserNameMaxChar=Cint(KS.Setting(30))
		 Dim EnabledUserName:EnabledUserName=KS.Setting(31)
		 Dim NewRegUserMoney:NewRegUserMoney=KS.Setting(38) : If Not IsNumerIc(NewRegUserMoney) Then NewRegUserMoney=0
		 Dim NewRegUserScore:NewRegUserScore=KS.Setting(39) : If Not IsNumeric(NewRegUserScore) Then NewRegUserScore=0
		 Dim NewRegUserPoint:NewRegUserPoint=KS.Setting(40) : If Not IsNumeric(NewRegUserPoint) Then NewRegUserPoint=0
         
		  If Request.ServerVariables("HTTP_REFERER")="" Then Call KS.Alert("�벻Ҫ�Ƿ��ύ!","../"):Response.End
		  If Instr(Lcase(Request.ServerVariables("SCRIPT_NAME")),"user/reg")=0 Then Call KS.Alert("�벻Ҫ�Ƿ��ύ!","../") : Response.End
		 
		 '�ռ��û�����
		 Dim Verifycode:Verifycode=KS.S("Verifycode")
		 If KS.Setting(32)="1"  Then
		  IF Trim(Verifycode)<>Trim(Session("Verifycode")) And VerificCodeTF=1 then 
		   	 Response.Write("<script>alert('��֤���������������룡');history.back(-1);</script>")
		     Exit Sub
		  End IF
		 End If
		 
		  '���ע��ش�����
		  Dim CanReg,n
		   If Mid(KS.Setting(161),1,1)="1" Then
		        CanReg=false
				 For N=0 To Ubound(Split(KS.Setting(162),vbcrlf))
				   If Trim(Request.Form("a" & MD5(n,16)))<>"" Then
					  If trim(Lcase(Request.Form("a" & MD5(n,16))))<>trim(Lcase(Split(KS.Setting(163),vbcrlf)(n))) Then
					   Call KS.AlertHistory("�Բ���,ע������Ļش���ȷ!",-1) : Response.End
					   CanReg=false
					  Else
					   CanReg=True
					  End If
				   End If
				 Next
			 If CanReg=false Then Call KS.AlertHistory("�Բ���,ע��𰸲���Ϊ��!",-1) : Response.End
		   End If
		 
		  
		 Dim UserName:UserName=KS.R(KS.S("UserName"))
		 If UserName = "" Or KS.strLength(UserName) > UserNameMaxChar Or KS.strLength(UserName) < UserNameLimitChar Then
		   	 Response.Write("<script>alert('�������û���(���ܴ���" & UserNameMaxChar & "С��" & UserNameLimitChar & ")');history.back();</script>")
			 Exit Sub
         ElseIF KS.FoundInArr(EnabledUserName, UserName, "|") = True Then
		   	 Response.Write("<script>alert('��������û���Ϊϵͳ��ֹע����û���');history.back();</script>")
			 Exit Sub
		 ElseIF InStr(UserName, "=") > 0 Or InStr(UserName, ".") > 0 Or InStr(UserName, "%") > 0 Or InStr(UserName, Chr(32)) > 0 Or InStr(UserName, "?") > 0 Or InStr(UserName, "&") > 0 Or InStr(UserName, ";") > 0 Or InStr(UserName, ",") > 0 Or InStr(UserName, "'") > 0 Or InStr(UserName, ",") > 0 Or InStr(UserName, Chr(34)) > 0 Or InStr(UserName, Chr(9)) > 0 Or InStr(UserName, "��") > 0 Or InStr(UserName, "$") > 0 Or InStr(UserName, "*") Or InStr(UserName, "|") Or InStr(UserName, """") > 0 Then
             Response.Write("<script>alert('�û����к��зǷ��ַ�');history.back();</script>")
			 Exit Sub
        End If
		 
		 Dim PassWord,RePassWord,NoMD5_Pass
		 If Session("PassWord")<>"" Then
		   PassWord=Session("PassWord")
		 Else
			 PassWord=KS.R(KS.S("PassWord"))
			 RePassWord=KS.S("RePassWord")
			 If PassWord = "" Then
				 Response.Write("<script>alert('�������¼����!');history.back();</script>")
				 Exit Sub
			 ElseIF RePassWord="" Then
				 Response.Write("<script>alert('������ȷ������');history.back();</script>")
				 Exit Sub
			 ElseIF PassWord<>RePassWord Then
				 Response.Write("<script>alert('������������벻һ��');history.back();</script>")
				 Exit Sub
			 End If
		 End If
		 NoMD5_Pass=PassWord
		 Dim RndPassword:RndPassword=NoMD5_Pass
		 Dim Question:Question=KS.S("Question")
		 Dim Answer:Answer=KS.S("Answer")
		 If KS.Setting(148)="1" Then
			 If Question = "" Then
				 Response.Write("<script>alert('������ʾ���ⲻ��Ϊ��!');history.back();</script>")
				 Exit Sub
			 ElseIF Answer="" Then
				 Response.Write("<script>alert('����𰸲���Ϊ��');history.back();</script>")
				 Exit Sub
			 End If
		 End If
		 
		 Dim Email:Email=KS.S("Email")
		 if KS.IsValidEmail(Email)=false then
			 Response.Write("<script>alert('��������ȷ�ĵ�������!');history.back();</script>")
			 Exit Sub
		 end if
		 
		 Dim RealName:RealName=KS.S("RealName")
		 Dim Sex:Sex=KS.S("Sex")
		 Dim Birthday:Birthday=KS.S("Birthday")
		 If Not IsDate(Birthday) Then Birthday=FormatDateTime(Now,2)
		 Dim IDCard:IDCard=KS.S("IDCard")
		 Dim OfficeTel:OfficeTel=KS.S("OfficeTel")
		 Dim HomeTel:HomeTel=KS.S("HomeTel")
		 Dim Mobile:Mobile=KS.S("Mobile")
		 Dim Fax:Fax=KS.S("Fax")
		 Dim province:province=KS.S("province")
		 Dim city:city=KS.S("city")
		 Dim Address:Address=KS.S("Address")
		 Dim ZIP:ZIP=KS.S("ZIP")
		 Dim HomePage:HomePage=KS.S("HomePage")
		 Dim UserFace:UserFace=KS.S("UserFace")
		 if userface="" then 
		   if sex="��" then userface=KS.GetDomain & "Images/Face/0.gif" else userface=KS.GetDomain & "Images/face/girl.gif"	 	 
		 End If
		 Dim QQ:QQ=KS.S("QQ")		 
		 Dim ICQ:ICQ=KS.S("ICQ")		 
		 Dim MSN:MSN=KS.S("MSN")		 
		 Dim UC:UC=KS.S("UC")		 
		 Dim Sign:Sign=KS.S("Sign")	
		 Dim Privacy:Privacy=KS.ChkClng(KS.S("Privacy"))
		 Dim AllianceUser:AllianceUser=KS.S("AllianceUser")
		 
		 Dim LastLoginIP:LastLoginIP = KS.GetIP()
		 Dim CheckNum:CheckNum = KS.MakeRandomChar(6)  '����ַ���֤��
		 Dim CheckUrl:CheckUrl = Request.ServerVariables("HTTP_REFERER")
		 CheckUrl=KS.GetDomain &"User/?RegActive?UserName=" & UserName &"&CheckNum=" & CheckNum
		 	 
		 PassWord =MD5(KS.R(PassWord),16)
		 Dim RS,SQL,K
		 Dim GroupID:GroupID=KS.ChkClng(KS.S("GroupID")):If GroupID=0 Then GroupID=2
		 Dim FieldsList:FieldsList=LFCls.GetSingleFieldValue("Select FormField From KS_UserForm Where ID=(Select FormID From KS_UserGroup Where ID=" & GroupID&")")
		 If FieldsList="" Then FieldsList="0"
	     Set RS = Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select FieldName,MustFillTF,Title,FieldType From KS_Field Where ChannelID=101 and FieldID In(" & KS.FilterIDs(FieldsList) & ")",conn,1,1
		 If Not RS.Eof Then SQL=RS.GetRows(-1)
		 RS.Close
		 If KS.Setting(32)="2" Then
			 For K=0 To UBound(SQL,2)
			   If SQL(1,K)="1" Then
			     If SQL(0,K)="Province&City" Then
				  If KS.S("Province")="" and  KS.S("City")="" Then
					 Response.Write "<script>alert('" & SQL(2,K) & "������д!');history.back();</script>"
					 Response.End()
				  End If
				 ElseIf KS.S(SQL(0,K))="" Then
					 Response.Write "<script>alert('" & SQL(2,K) & "������д!');history.back();</script>"
					 Response.End()
				 End If
			   End If
			   If SQL(3,K)="4" And Not Isnumeric(KS.S(SQL(0,K))) Then 
				 Response.Write "<script>alert('" & SQL(2,K) & "������д����!');history.back();</script>"
				 Response.End()
			   End If
			   If SQL(3,K)="5" And Not IsDate(KS.S(SQL(0,K))) Then 
				 Response.Write "<script>alert('" & SQL(2,K) & "������д��ȷ������!');history.back();</script>"
				 Response.End()
			   End If
			   If SQL(3,K)="8" And Not KS.IsValidEmail(KS.S(SQL(0,K))) and SQL(1,K)="1" Then 
				Response.Write "<script>alert('" & SQL(2,K) & "������д��ȷ��Email��ʽ!');history.back();</script>"
				Response.End()
			   End If 
			 Next
		End If
		RS.Open "Select ID From KS_UserGroup Where ID=" & GroupID,conn,1,1
		If RS.Eof And RS.Bof Then
		     Rs.Close:Set RS=Nothing
			 Response.Write "<script>alert('�Բ���,�û������Ͳ���ȷ!');history.back();</script>"
			 Response.End()
		End If
		RS.Close
		RS.Open "select top 1 * from KS_User where UserName='" & UserName & "'", Conn, 1, 3
		If Not (RS.BOF And RS.EOF) Then
				 RS.Close:Set RS=Nothing
				 Response.Write("<script>alert('��ע����û��Ѿ����ڣ��뻻һ���û��������ԣ�');history.back();</script>")
				 Exit Sub
		Else
			If EmailMultiRegTF=0 Then
				Dim EmailRSCheck:Set EmailRSCheck = Conn.Execute("select UserID from KS_User where Email='" & Email & "'")
				If Not (EmailRSCheck.BOF And EmailRSCheck.EOF) Then
					EmailRSCheck.Close:Set EmailRSCheck = Nothing
					Response.Write("<script>alert('��ע���Email�Ѿ����ڣ������Email�����ԣ�');history.back();</script>")
					Exit Sub
				End If
				EmailRSCheck.Close:Set EmailRSCheck = Nothing
			 End If
		
			 If KS.ChkClng(KS.Setting(26))=1 Then
			   If Not (Conn.Execute("select top 1 UserID From KS_User Where LastLoginIP='" & KS.GetIP & "'").eof) Then
					Response.Write("<script>alert('����IP�Ѿ����ڣ�������ע�ᣡ');history.back();</script>")
					Exit Sub
			   End If
			 End If

		
		'-----------------------------------------------------------------
		'ϵͳ����
		'-----------------------------------------------------------------
		Dim API_KS,API_SaveCookie,SysKey
		If API_Enable Then
		    If Question="" Then Question="what's the date?"
			If Answer="" Then Answer=now
			Set API_KS = New API_Conformity
			API_KS.NodeValue "action","reguser",0,False
			API_KS.NodeValue "username",UserName,1,False
			'Md5OLD = 1
			SysKey = Md5(API_KS.XmlNode("username") & API_ConformKey,16)
			'Md5OLD = 0
			API_KS.NodeValue "syskey",SysKey,0,False
			API_KS.NodeValue "password",NoMD5_Pass,0,False
			API_KS.NodeValue "email",Email,1,False
			API_KS.NodeValue "question",Question,1,False
			API_KS.NodeValue "answer",Answer,1,False
			API_KS.NodeValue "gender",sex,0,False

			API_KS.SendHttpData
			If API_KS.Status = "1" Then
				Response.Write "<script>alert('" & API_KS.Message & "');history.back();</script>"
				Exit Sub
			Else
				API_SaveCookie = API_KS.SetCookie(SysKey,UserName,Password,1)
			End If
			Set API_KS = Nothing
		End If
		'-----------------------------------------------------------------
		 
		 
		 RS.AddNew
		 RS("GroupID")=GroupID
		 RS("UserName")=UserName
		 RS("PassWord")=PassWord
		 RS("Question")=Question
		 RS("Answer")=Answer
		 RS("Email")=Email
		 RS("RealName")=RealName
		 RS("Sex")=Sex
		 RS("Birthday")=Birthday
		 RS("IDCard")=IDCard
		 RS("OfficeTel")=OfficeTel
		 RS("HomeTel")=HomeTel
		 RS("Mobile")=Mobile
		 RS("Fax")=Fax
		 RS("Province")=Province
		 RS("City")=City
		 RS("Address")=Address
		 RS("Zip")=Zip
		 RS("HomePage")=HomePage
		 RS("QQ")=QQ
		 RS("ICQ")=ICQ
		 RS("MSN")=MSN
		 RS("UC")=UC
		 RS("UserFace")=UserFace
		 RS("Sign")=Sign
		 RS("Privacy")=Privacy
		 RS("RegDate")=Now
		 RS("BeginDate")=Now '��ʼ����ʱ��
		 RS("LastLoginIP")=LastLoginIP
		 RS("JoinDate")=Now
		 RS("LastLoginTime")=Now
		 RS("CheckNum")=CheckNum
		 RS("RndPassword")=RndPassword
		 RS("LoginTimes")=1
		 
		 '�Զ����ֶ�
		 Dim UpFiles
		 For K=0 To UBound(SQL,2)
		  If left(Lcase(SQL(0,K)),3)="ks_" Then
		   RS(SQL(0,K))=KS.S(SQL(0,K))
		   If SQL(3,K)="9" or SQL(3,K)="10" Then
		   UpFiles=KS.S(SQL(0,K))
		   End If
		  End If
		 Next
		 
		 RS("AllianceUser")=AllianceUser

		 '�»�Աע�ᣬ������Ӧ������
		 RS("Money")=0
		 RS("Score")=NewRegUserScore
		 
		 If KS.ChkClng(KS.U_G(GroupID,"chargetype"))=1 Then
		  NewRegUserPoint=KS.ChkClng(KS.U_G(GroupID,"grouppoint"))
		 End If
		 'If KS.ChkClng(KS.U_G(GroupID,"grouppoint"))<>0 Then NewRegUserPoint=KS.ChkClng(KS.U_G(GroupID,"grouppoint"))
		 
		 RS("Point")=NewRegUserPoint
		 
		 RS("Locked")=Locked
		 RS.Update
		 
		 RS.MoveLast
		 Dim UserID:UserID=RS("UserID")
		 RS.Close
		 If Not KS.IsNul(UpFiles) Then
		  Call KS.FileAssociation(1023,UserID,UpFiles,0)
		 End If
		 
		 
		 If NewRegUserPoint<>0 Then
		   Conn.Execute("Insert into KS_LogPoint(ChannelID,InfoID,UserName,InOrOutFlag,Point,Times,[User],Descript,Adddate,IP) values(0,0,'" & UserName & "',1," & NewRegUserPoint & ",1,'ϵͳ','ע���»�Ա,����!'," & SqlNowString & ",'" & replace(ks.getip,"'","""") & "')")
		 End If
		 IF NewRegUserScore<>0 Then
		  Conn.Execute("Insert into KS_LogScore(UserName,InOrOutFlag,Score,CurrScore,[User],Descript,Adddate,IP) values('" & UserName & "',1," & NewRegUserScore & ","&NewRegUserScore& ",'ϵͳ','ע���»�Ա,����!'," & SqlNowString & ",'" & replace(ks.getip,"'","""") & "')")
		 End If
		 If NewRegUserMoney<>0 Then 
		  Call KS.MoneyInOrOut(UserName,UserName,NewRegUserMoney,4,1,now,0,"System","ע���»�Ա,����!",0,0)
		 End If
		 
		 RS.Open "Select * From KS_UserGroup Where ID=" & GroupID,conn,1,1
		 If RS.Eof Then RS.Close : Set RS=Nothing :Response.Write "<script>location.href='../../';</script>"
		 
		 Dim EmailCheckTF:EmailCheckTF=RS("ValidType")
		 Dim UserType:UserType= RS("UserType")
		 Conn.Execute("Update KS_User Set ChargeType=" & RS("ChargeType") & ",EDays=" & RS("ValidDays") & ",UserType=" & UserType &" Where UserID=" & UserID)
		 
		 Dim MailBodyStr
		 If EmailCheckTF = 1 Then
			MailBodyStr = Replace(RS("ValidEmail"), "{$CheckNum}", CheckNum)
			MailBodyStr = Replace(MailBodyStr, "{$CheckUrl}", CheckUrl)
	
	       Dim ReturnInfo:ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), "���û�ע�ἤ����", Email,KS.Setting(0), MailBodyStr,KS.Setting(11))
			  IF ReturnInfo="OK" Then
			     Conn.Execute("Update KS_User Set Locked=3 where userid=" & UserID)        '���ô�����
				 ReturnInfo="<li>ע��ɹ���ע����֤���ѷ��͵���������<font color='#ff6600>'>" &Email &"</font>��ֻ�м����ſ�����ʽ��Ϊ��վ��Ա!</li>"
			  Else
				ReturnInfo="<li>�ż�����ʧ��!ʧ��ԭ��:" & ReturnInfo & "������ϵ��վ����Ա!</li>"
			  End if
		ElseIF EmailCheckTF=2 Then
		    Conn.Execute("Update KS_User Set Locked=2 where userid=" & UserID)        '������Ҫ��̨��֤
		    ReturnInfo="<li>ע��ɹ�!�����û���:<font color=red>" & UserName & "</font>,����Ҫͨ������Ա����֤���ܳ�Ϊ��ʽ��Ա!</li>"
		Else
		    ReturnInfo="<li>ע��ɹ�!�����û���:<font color=red>" & UserName & "</font>,���ѳ�Ϊ�˱�վ����ʽ��Ա!<br><div align=center></li>"
        End IF
		
		RS.Close
			
			
			'====================�Ƽ��ƻ�======================================
			If AllianceUser<>"" and AllianceUser<>UserName  Then
			 If Not Conn.Execute("Select Top 1 UserID From KS_User Where UserName='" & AllianceUser & "'").eof Then
			   Call KS.ScoreInOrOut(AllianceUser,1,KS.ChkClng(KS.Setting(144)),"ϵͳ","�ɹ��Ƽ�һ��ע���û�:" & UserName & "!",0,0)
			   
			   Conn.Execute("Insert InTo KS_PromotedPlan(UserName,UserIP,AddDate,ComeUrl,Score,AllianceUser) values('" & AllianceUser & "','" & KS.GetIP & "'," & SqlNowString & ",'" & KS.URLDecode(Request.ServerVariables("HTTP_REFERER")) & "'," & KS.ChkClng(KS.Setting(144)) & ",'" & UserName & "')")
			  '=================�ж��ǲ��Ǻ����ʼ��Ƽ���==================
				Dim f:f=KS.S("F")
				if f="r" Then
				 Conn.Execute("insert into KS_Friend (username,friend,addtime,flag,message,accepted) values ('"&AllianceUser&"','"& UserName &"',"&SqlNowString&",1,'',1)")
				End If
			  '============================================================

			 End If
			End If
			'====================�ƹ�ƻ�����=================================
			
			'==================ע��ɹ����ʼ�֪ͨ======================
			 If KS.Setting(146)="1" and Not KS.IsNul(KS.Setting(147)) And EmailCheckTF<>1 Then
				MailBodyStr = Replace(KS.Setting(147), "{$UserName}", UserName)
				MailBodyStr = Replace(MailBodyStr, "{$PassWord}", NoMD5_Pass)
				MailBodyStr = Replace(MailBodyStr, "{$SiteName}", KS.Setting(0))
				ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), KS.Setting(0) & "-��Աע��ɹ�", Email,UserName, MailBodyStr,KS.Setting(11))
				IF ReturnInfo="OK" Then
				  ReturnInfo="<li>ע��ɹ�!�����û���:<font color=red>" & UserName & "</font>,�ѽ��û��������뷢����������!</li>"
				End If
			 End If
			'==========================================================
			
			
			
		    '===================д����˿ռ�================
			if KS.SSetting(0)=1 And KS.SSetting(1)=1 then
			 RS.Open "Select * From KS_Blog Where 1=0",conn,1,3
			 RS.AddNew
			  RS("UserName")=UserName
			  RS("BlogName")=UserName & "�ĸ��˿ռ�"
			  RS("ClassID") = KS.ChkClng(Conn.Execute("Select Top 1 ClassID From KS_BlogClass")(0))
			  If UserType=1 Then
			  RS("TemplateID")=KS.ChkClng(Conn.Execute("Select Top 1 ID From KS_BlogTemplate Where flag=4 and IsDefault='true'")(0))
			  Else
			  RS("TemplateID")=KS.ChkClng(Conn.Execute("Select Top 1 ID From KS_BlogTemplate Where flag=2 and IsDefault='true'")(0))
			  End If
			  RS("Announce")="���޹���!"
			  RS("ContentLen")=500
			  RS("Recommend")=0
			  if KS.SSetting(2)=1 then
			  RS("Status")=0
			  else
			  RS("Status")=1
			  end if
			 RS.Update
			 RS.Close
			 '�ж�����ҵ��Ա���Զ���ͨ��ҵ�ռ�
				 On Error Resume Next
				 If UserType=1 then
				   Dim FieldsXml:Set FieldsXml=LFCls.GetXMLFromFile("SpaceFields")
				   RS.Open "Select top 1 * From KS_EnterPrise Where 1=0",conn,1,3
				   RS.AddNew
				   
	   			     RS("UserName")=UserName
					' RS("CompanyName")=KS.S("KS_company")
					 RS("Province")=Province
					 RS("City")=City
					 RS("Address")=Address
					 RS("ZipCode")=Zip
					 RS("ContactMan")=RealName
					 RS("TelPhone")=OfficeTel
					 RS("Fax")=Fax
					 RS("AddDate")=Now
					 RS("Recommend")=0
					 RS("ClassID")=0
					 RS("SmallClassID")=0
					  if KS.SSetting(2)=1 then
					  RS("Status")=0
					  else
					  RS("Status")=1
					  end if
					  
				    If IsObject(FieldsXml) Then
						 on error resume next
						 Dim objNode,i,j,objAtr
						 Set objNode=FieldsXml.documentElement 
						 For i=0 to objNode.ChildNodes.length-1 
								set objAtr=objNode.ChildNodes.item(i) 
								Execute("RS(""" & objAtr.Attributes.item(0).Text & """)=KS.S(""" & objAtr.Attributes.item(1).Text &""")")
						 Next
				
					   End If
				   RS.Update 
				   RS.Close
				 End If
			 End If
		    '==================================
			 Set RS=Nothing
		    If EmailCheckTF=0 Then
			Response.Cookies(KS.SiteSn).path = "/"
			Response.Cookies(KS.SiteSn)("UserName") = UserName
			Response.Cookies(KS.SiteSn)("PassWord") = PassWord
			Response.Cookies(KS.SiteSn)("RndPassword") = RndPassword
			End If
			'-----------------------------------------------------------------
			'ϵͳ����
			'-----------------------------------------------------------------
			If API_Enable Then
				Response.Write API_SaveCookie
				Response.Flush
				If API_ReguserUrl <> "0" Then
					Response.Write "<script language=JavaScript>"
					Response.Write "setTimeout(""window.location='"& API_ReguserUrl &"'"",1000);"
					Response.Write "</script>"
				End If

			End If
			'-----------------------------------------------------------------
			Call ShowRegResult(ReturnInfo)
    End If	 
          
		End Sub
		
		Sub ShowRegResult(ReturnInfo)
		   ReturnInfo="<table border='0' align='center' width='60%' cellspacing='1' cellpadding='3'><tr class='tdbg'><td></td></tr><tr class='tdbg'><td><img src='../images/regok.jpg' align='left'>" & ReturnInfo & "<li><span id=""countdown""> <font color='#ff6600'><strong>10</strong></font> </span>���Ӻ��Զ�ת����Ա����!</li><li><a href=""../index.asp"">���Ͻ����Ա����</a>  <a href=""" &  KS.Setting(3)& """>������վ��ҳ</a></li></td></tr></table>" & vbcrlf
		   Dim FileContent
		   If KS.Setting(119)="" Then Response.Write "���ȵ�""������Ϣ����->ģ���""����ģ��󶨲���!":response.end
		    FileContent = KSRFObj.LoadTemplate(KS.Setting(119)) &  GetTurnMember(10)
            FileContent= Replace(FileContent,"{$GetUserRegResult}",ReturnInfo)
			FileContent = KSRFObj.KSLabelReplaceAll(FileContent) '�滻������ǩ
            Response.Write FileContent
		End Sub
		
		Function GetTurnMember(Times)
		GetTurnMember="<script language=""JavaScript"">"&vbcrlf
		GetTurnMember=GetTurnMember & "var OutTimes = " & Times & "; " & vbcrlf
		GetTurnMember=GetTurnMember & "leavePage()" & vbcrlf
		GetTurnMember=GetTurnMember & "function leavePage() {" & vbcrlf
		GetTurnMember=GetTurnMember & "if (OutTimes==0)" & vbcrlf
		GetTurnMember=GetTurnMember & "	top.location.href='../';" & vbcrlf
		GetTurnMember=GetTurnMember & "else {" & vbcrlf
		GetTurnMember=GetTurnMember & "OutTimes -= 1;"& vbcrlf
		GetTurnMember=GetTurnMember & "document.getElementById('countdown').innerHTML =""<font color='#ff6600'><strong>""+ OutTimes + ""</strong></font> "";"& vbcrlf
		GetTurnMember=GetTurnMember & "setTimeout(""leavePage()"", 1000);"  & vbcrlf
		GetTurnMember=GetTurnMember & "}}" & vbcrlf
		GetTurnMember=GetTurnMember & "</script>" & vbcrlf
		End Function
End Class
%>

 
