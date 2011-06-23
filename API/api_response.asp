<!--#include file="../conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Plus/md5.asp"-->
<!--#include file="Cls_Api.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
'-- �������������޸��Զ�����̳ϵͳApi�ӿ�
'=========================================================
Dim XMLDom,XmlDoc,Node,Status,Messenge
Dim UserName,Act,appid
Status = 1
Messenge = ""

If Request.QueryString<>"" And API_Enable Then
	SaveUserCookie()
Else
	Set XmlDoc = KS.InitialObject("msxml2.FreeThreadedDOMDocument" & MsxmlVersion)
	XmlDoc.ASYNC = False
	If API_Enable Then
		If Not XmlDoc.LOAD(Request) Then
			Status = 1
			Messenge = "���ݷǷ���������ֹ��"
			appid = "δ֪"
		Else
			If CheckPost() Then
				Select Case Act
					Case "checkname"
						Checkname()
					Case "reguser"
						UserReguser()
					Case "login"
						UesrLogin()
					Case "logout"
						LogoutUser()
					Case "update"
						UpdateUser()
					Case "delete"
						Deleteuser()
					Case "lock"
						Lockuser()
					Case "getinfo"
						GetUserinfo()
				End Select
			End If
		End If
	Else
		Status = 0
		Messenge = "API�ӿڹرգ�������ֹ��"
		appid = "KesionCMS"
	End If
	ReponseData()
	Set XmlDoc = Nothing
End If

Sub ReponseData()
	If Act <> "getinfo" Then
		XmlDoc.loadxml "<root><appid>dvbbs</appid><status>0</status><body><message/></body></root>"
	End If
	XmlDoc.documentElement.selectSingleNode("appid").text = "KesionCMS"
	If API_Debug And Act <> "reguser" Then
		XmlDoc.documentElement.selectSingleNode("status").text = 0
		Messenge = ""
	Else
		XmlDoc.documentElement.selectSingleNode("status").text = status
	End If
	XmlDoc.documentElement.selectSingleNode("body/message").text = ""
	Set Node = XmlDoc.createCDATASection(Replace(Messenge,"]]>","]]&gt;"))
	XmlDoc.documentElement.selectSingleNode("body/message").appendChild(Node)
	Response.Clear
	Response.ContentType="text/xml"
	Response.CharSet="gb2312"
	Response.Write "<?xml version=""1.0"" encoding=""gb2312""?>"&vbNewLine
	Response.Write XmlDoc.documentElement.XML
End Sub

Function CheckPost()
	CheckPost = False
	Dim Syskey
	If XmlDoc.documentElement.selectSingleNode("action") is Nothing or XmlDoc.documentElement.selectSingleNode("syskey") is Nothing or XmlDoc.documentElement.selectSingleNode("username")  is Nothing Then
		Status = 1
		Messenge = Messenge & "<li>�Ƿ�����</li>"
		Exit Function
	End If
	UserName = KS.R(XmlDoc.documentElement.selectSingleNode("username").text)
	Syskey = XmlDoc.documentElement.selectSingleNode("syskey").text
	Act = XmlDoc.documentElement.selectSingleNode("action").text
	Appid = XmlDoc.documentElement.selectSingleNode("appid").text
	
	Dim NewMd5,OldMd5
	NewMd5 = Md5(UserName & API_ConformKey,16)
	Md5OLD = 1
	OldMd5 = Md5(UserName & API_ConformKey,16)
	Md5OLD = 0

	If Syskey=NewMd5 or Syskey=OldMd5 Then
		CheckPost = True
	Else
		Status = 1
		Messenge = Messenge & "<li>����������֤��ͨ�����������Ա��ϵ��</li>"
	End If
End Function

Sub GetUserinfo()
	Dim Rs,Sql
	XmlDoc.loadxml "<root><appid>KesionCMS</appid><status>0</status><body><message/><email/><question/><answer/><savecookie/><truename/><gender/><birthday/><qq/><msn/><mobile/><telephone/><address/><zipcode/><homepage/><userip/><jointime/><experience/><ticket/><valuation/><balance/><posts/><userstatus/></body></root>"
	
	Sql = "SELECT TOP 1 * FROM KS_User WHERE UserName='" & KS.R(UserName) & "'"
	Set Rs = Conn.Execute(Sql)
	If Not Rs.Eof And Not Rs.Bof Then
		XmlDoc.documentElement.selectSingleNode("body/email").text = Rs("email") & ""
		XmlDoc.documentElement.selectSingleNode("body/question").text = Rs("question") & ""
		XmlDoc.documentElement.selectSingleNode("body/answer").text = Rs("answer") & ""
		XmlDoc.documentElement.selectSingleNode("body/gender").text = Rs("sex") & ""
		XmlDoc.documentElement.selectSingleNode("body/birthday").text = ""
		XmlDoc.documentElement.selectSingleNode("body/mobile").text = RS("mobile")
		XmlDoc.documentElement.selectSingleNode("body/userip").text = Rs("LastLoginIP") & ""
		XmlDoc.documentElement.selectSingleNode("body/jointime").text = Rs("Joindate") & ""
		XmlDoc.documentElement.selectSingleNode("body/experience").text =""
		XmlDoc.documentElement.selectSingleNode("body/ticket").text = ""
		XmlDoc.documentElement.selectSingleNode("body/valuation").text = Rs("point") & ""
		XmlDoc.documentElement.selectSingleNode("body/balance").text = Rs("Money") & ""
		XmlDoc.documentElement.selectSingleNode("body/posts").text = Rs("zip") & ""
		XmlDoc.documentElement.selectSingleNode("body/userstatus").text = Rs("Locked")
		XmlDoc.documentElement.selectSingleNode("body/homepage").text = Rs("HomePage") & ""
		XmlDoc.documentElement.selectSingleNode("body/qq").text = Rs("qq")
		XmlDoc.documentElement.selectSingleNode("body/msn").text = rs("msn")
		XmlDoc.documentElement.selectSingleNode("body/truename").text = Rs("realName") & ""
		XmlDoc.documentElement.selectSingleNode("body/telephone").text = Rs("OfficeTel") & ""
		XmlDoc.documentElement.selectSingleNode("body/address").text = Rs("address") & ""
		Status = 0
		Messenge = Messenge & "<li>��ȡ�û����ϳɹ���</li>"
	Else
		Status = 1
		Messenge = Messenge & "<li>���û������ڡ�</li>"
	End If
	Rs.Close
	Set Rs = Nothing
End Sub

Sub Checkname()
	Dim Rs,SQL,UserEmail
	UserEmail = KS.R(Trim(XmlDoc.documentElement.selectSingleNode("email").text))
	If KS.IsValidEmail(UserEmail) = False Then
		Messenge = "<li>����Email�д���</li>"
		Status = 1
		Exit Sub
	End If
	If CInt(KS.Setting(28)) = 1 Then
		Set Rs = Conn.Execute("SELECT userid FROM KS_User WHERE Email='" & UserEmail & "'")
		If Not Rs.EOF Then
			Status = 1
			Messenge = "<li>������["&UserEmail&"]�Ѿ�ռ�ã�������һ��������ע��ɡ�</li>"
			Exit Sub
		End If
		Rs.Close:Set Rs = Nothing
	End If
	Set Rs = Conn.Execute("SELECT username FROM KS_User WHERE username = '" & UserName & "'")
	If Not (Rs.bof And Rs.EOF) Then
		Status = 1
		Messenge =  "<li>Sorry�����û��Ѿ�����,�뻻һ���û������ԣ�</li>"
	Else
		Status = 0
		Messenge =  "<li><font color=red><b>" & UserName & "</b></font> ��δ����ʹ�ã��Ͻ�ע��ɣ�</li>"
	End If
	Rs.Close:Set Rs = Nothing
End Sub

Sub UserReguser()
	Dim nickname,UserPass,UserEmail,Question,Answer,usercookies
	Dim strGroupName,Password,usersex,sex
	Dim Rs,SQL
	UserPass = KS.R(XmlDoc.documentElement.selectSingleNode("password").text)
	UserEmail = KS.R(Trim(XmlDoc.documentElement.selectSingleNode("email").text))
	Question = KS.R(XmlDoc.documentElement.selectSingleNode("question").text)
	Answer = KS.R(XmlDoc.documentElement.selectSingleNode("answer").text)
	sex = KS.R(XmlDoc.documentElement.selectSingleNode("gender").text)
	
	Dim NewRegUserMoney:NewRegUserMoney=KS.Setting(38)
	Dim NewRegUserScore:NewRegUserScore=KS.Setting(39)
	Dim NewRegUserPoint:NewRegUserPoint=KS.Setting(40)

	If sex = "0" Then
		usersex = "Ů"
	Else
		usersex = "��"
	End If
	usercookies = 1
	If UserName = "" Or UserPass = "" Then
		Status = 1
		Messenge = Messenge & "<li>����д�û��������롣"
		Exit Sub
	End If
	If Question = "" Then Question = KS.MakeRandomChar(20)
	If Answer = "" Then Answer = KS.MakeRandomChar(20)
	nickname = UserName
	Password = MD5(KS.R(UserPass),16)
	Answer = Answer
	If KS.IsValidEmail(UserEmail) = False Then
		Messenge = Messenge & "<li>����Email�д���</li>"
		Status = 1
		Exit Sub
	End If
	Set Rs = Conn.Execute("SELECT username FROM KS_User WHERE username='" & UserName & "'")
	If Not (Rs.BOF And Rs.EOF) Then
		Status = 1
		Messenge = Messenge & "<li>Sorry�����û��Ѿ�����,�뻻һ���û������ԣ�</li>"
		Exit Sub
	End If
	Rs.Close:Set Rs = Nothing
	If CInt(KS.Setting(28)) = 1 Then
		Set Rs = Conn.Execute("SELECT userid FROM KS_User WHERE Email='" & UserEmail & "'")
		If Not Rs.EOF Then
			Status = 1
			Messenge = Messenge & "<li>�Բ��𣡱�ϵͳ�Ѿ�����һ������ֻ��ע��һ���˺š�</li><li>������["&UserEmail&"]�Ѿ�ռ�ã�������һ��������ע��ɡ�</li>"
			Exit Sub
		End If
		Rs.Close:Set Rs = Nothing
	End If

	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM KS_User WHERE (userid is null)"
	Rs.Open SQL,Conn,1,3
	Rs.Addnew
		Rs("username") = UserName
		Rs("password") = Password
		RS("GroupID")=2    '�趨Ĭ���û�����Ϊ���˻�Ա
		Rs("answer") = Answer
		Rs("question") = Question
		Rs("UserFace") = "Images/Face/0.gif"
		Rs("RealName") = UserName
		Rs("sex") = usersex
		Rs("Email") = UserEmail
		Rs("qq") = ""
		RS("RegDate")=Now
		RS("BeginDate")=Now '��ʼ����ʱ��
		RS("LastLoginIP")=KS.GetIP
		RS("JoinDate")=Now
		RS("LastLoginTime")=Now
		
		 '�»�Աע�ᣬ������Ӧ������
		 RS("Money")=NewRegUserMoney
		 RS("Score")=NewRegUserScore
		 RS("Point")=NewRegUserPoint
		 Call KS.PointInOrOut(0,0,UserName,1,NewRegUserPoint,"ϵͳ","ע���»�Ա,���͵�" & KS.Setting(46) & KS.Setting(45),0)
		 RS("Locked")=0
	Rs.update
	RS.movelast
	Conn.Execute("Update KS_User Set ChargeType=" & Conn.Execute("Select ChargeType From KS_UserGroup Where ID=" & RS("GroupID"))(0) & " Where UserID=" & RS("UserID"))
	RS.Close
			  '===================д����˿ռ�================
			  If KS.SSetting(1)=1 Then
			 RS.Open "Select * From KS_Blog Where Blogid is null",conn,1,3
			 RS.AddNew
			  RS("UserName")=UserName
			  RS("BlogName")=UserName & "�ĸ��˿ռ�"
			  RS("ClassID") = KS.ChkClng(Conn.Execute("Select Top 1 ClassID From KS_BlogClass")(0))
			  RS("TemplateID")=KS.ChkClng(Conn.Execute("Select Top 1 ID From KS_BlogTemplate Where flag=2 and IsDefault='true'")(0))
			  RS("Announce")="���޹���!"
			  RS("ContentLen")=500
			  RS("Recommend")=0
			  RS("Status")=1
			 RS.Update
			 RS.Close
			 end if
			 Set RS=Nothing
		    '==================================

	Status = 0
	Messenge = "�û�ע��ɹ���"
End Sub

Sub UesrLogin()
	Dim UserPass
	
	UserPass = XmlDoc.documentElement.selectSingleNode("password").text
	If UserName="" or UserPass="" Then
		Status = 1
		Messenge = Messenge & "<li>����д�û��������롣</li>"
		Exit Sub
	End If
	UserPass = Md5(UserPass,16)
	
	If ChkUserLogin(username,UserPass,1) Then
		Status = 0
		Messenge = Messenge & "<li>��½�ɹ���</li>"
	Else
		Status = 1
		Messenge = Messenge & "<li>��½ʧ�ܡ�</li>"
	End If
End Sub

Sub LogoutUser()
	Response.Cookies(KS.SiteSn).path = "/"
	Response.Cookies(KS.SiteSn)("UserName") = ""
	Response.Cookies(KS.SiteSn)("Password") = ""
	Response.Cookies(KS.SiteSn)("RndPassword")=""
End Sub

Sub UpdateUser()
	Dim Rs,SQL
	Dim UserPass,UserEmail,Question,Answer
	UserPass = XmlDoc.documentElement.selectSingleNode("password").text
	UserEmail = Trim(XmlDoc.documentElement.selectSingleNode("email").text)
	Question = XmlDoc.documentElement.selectSingleNode("question").text
	Answer = XmlDoc.documentElement.selectSingleNode("answer").text
	If UserPass <> "" Then
		UserPass = Md5(UserPass,16)
	End If
	If Answer <> "" THen
		Answer = Answer
	End If
	If KS.IsValidEmail(UserEmail) = False Then
		UserEmail = ""
	End If
	Set Rs = Server.CreateObject("Adodb.RecordSet")
	SQL = "SELECT TOP 1 * FROM [KS_User] WHERE Username='" & UserName & "'"
	Rs.Open SQL,Conn,1,3
	If Not Rs.Eof And Not Rs.Bof Then
		If UserPass <> "" Then Rs("password") = UserPass
		If Answer <> "" THen Rs("answer") = Answer
		If UserEmail <> "" Then Rs("email") = UserEmail
		If Question <> "" Then Rs("question") = Question
		Rs.update
		Status = 0
		Messenge = "<li>���������޸ĳɹ���</li>"
		Response.Cookies(KS.SiteSN)("password") = UserPass
	Else
		Status = 1
		Messenge = "<li>���û������ڣ��޸�����ʧ�ܡ�</li>"
	End If
	Rs.Close:Set Rs = Nothing
End Sub

Sub Deleteuser()
	Dim Del_Users,i,AllUserID,Del_UserName
	Dim Rs
	Del_Users = Split(UserName,",")
	For i = 0 To UBound(Del_Users)
		Del_UserName = KS.R(Del_Users(i))
		Set Rs = Conn.Execute("SELECT userid,username FROM [KS_User] WHERE UserName='" & Del_UserName & "'")
		If Not (Rs.Eof And Rs.Bof) Then
			Conn.Execute ("DELETE FROM KS_User WHERE UserName='" & Del_UserName & "')")
			Conn.Execute ("DELETE FROM KS_Favorite WHERE UserName='" & Del_UserName & "')")
			Conn.Execute ("DELETE FROM KS_Comment WHERE UserName='" & Del_UserName & "')")
			Messenge = Messenge & "<li>�û���" & Del_UserName & "��ɾ���ɹ���</li>"
		End If
	Next
	Set Rs = Nothing
	Status = 0
End Sub

Sub Lockuser()
	Dim UserStatus
	If XmlDoc.documentElement.selectSingleNode("userstatus") is Nothing Then
		Messenge = "<li>�����Ƿ�����ֹ����</li>"
		Status = 1
		Exit Sub
	ElseIf Not IsNumeric(XmlDoc.documentElement.selectSingleNode("userstatus").text) Then
		Messenge = "<li>�����Ƿ�����ֹ����</li>"
		Status = 1
		Exit Sub
	Else
		UserStatus = Clng(XmlDoc.documentElement.selectSingleNode("userstatus").text)
	End If
	If UserStatus = 0 Then
		Conn.Execute ("UPDATE KS_User SET Locked=0 WHERE Username='" & UserName & "'")
	Else
		Conn.Execute ("UPDATE KS_User SET Locked=1 WHERE Username='" & UserName & "'")
	End If
	Status = 0
End Sub

Sub SaveUserCookie()
	Dim S_syskey,Password,usercookies,TruePassWord,userclass,Userhidden
	
	S_syskey = Request.QueryString("syskey")
	UserName = KS.R(Request.QueryString("UserName"))
	Password = Request.QueryString("Password")
	usercookies = Request.QueryString("savecookie")
	If UserName="" or S_syskey="" Then Exit Sub
	Dim NewMd5,OldMd5
	NewMd5 = Md5(UserName & API_ConformKey,16)
	Md5OLD = 1
	OldMd5 = Md5(UserName & API_ConformKey,16)
	Md5OLD = 0
	If Not (S_syskey=NewMd5 or S_syskey=OldMd5) Then
		Exit Sub
	End If
	If usercookies="" or Not IsNumeric(usercookies) Then usercookies = 0
	
	'�û��˳�
	If Password = "" Then
		Response.Cookies(KS.SiteSn).path = "/"
		Response.Cookies(KS.SiteSn)("UserName") = ""
		Response.Cookies(KS.SiteSn)("Password") = ""
		Response.Cookies(KS.SiteSn)("RndPassword")=""
		Exit Sub
	End If
	ChkUserLogin username,password,usercookies
End Sub

Function ChkUserLogin(username,password,usercookies)
	ChkUserLogin = False
	Dim Rs,SQL,RndPassWord
	RndPassWord=KS.R(KS.MakeRandomChar(20))
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM [KS_User] WHERE username='" & UserName & "'"
	Rs.Open SQL, Conn, 1, 3
	If Not (Rs.BOF And Rs.EOF) Then
		If password <> Rs("password") Then
			ChkUserLogin = False
			Exit Function
		End If
		If Rs("Locked") <> 0 Then
			ChkUserLogin = False
			Exit Function
		End If
		'��¼�ɹ��������û���Ӧ������
		If datediff("n",RS("LastLoginTime"),now)>=KS.Setting(36) then '�ж�ʱ��
		 RS("Score")=RS("Score")+KS.Setting(37)
		end if
		 RS("LastLoginIP") = KS.GetIP
         RS("LastLoginTime") = Now()
         RS("LoginTimes") = RS("LoginTimes") + 1
		 RS("RndPassword")= RndPassWord	
		Rs.Update
		
		Select Case usercookies
		Case 0
			Response.Cookies(KS.SiteSn)("usercookies") = usercookies
		Case 1
			Response.Cookies(KS.SiteSn)("usercookies") = usercookies
			Response.Cookies(KS.SiteSn).Expires=Date+1
		Case 2
			Response.Cookies(KS.SiteSn).Expires=Date+31
			Response.Cookies(KS.SiteSn)("usercookies") = usercookies
		Case 3
			Response.Cookies(KS.SiteSn).Expires=Date+365
			Response.Cookies(KS.SiteSn)("usercookies") = usercookies
		End Select
		Response.Cookies(KS.SiteSn).path = "/"
		Response.Cookies(KS.SiteSn)("UserName") = Rs("username")
		Response.Cookies(KS.SiteSn)("Password") = Rs("password")
		Response.Cookies(KS.SiteSn)("RndPassword")=RndPassWord
		ChkUserLogin = True
	End If
	Rs.Close:Set Rs = Nothing
End Function

%>