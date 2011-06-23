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
		 '系统配置参数
		 Dim Locked:Locked=0
		 Dim VerificCodeTF:VerificCodeTF=KS.Setting(27)
		 Dim EmailMultiRegTF:EmailMultiRegTF=KS.Setting(28)
		 Dim UserNameLimitChar:UserNameLimitChar=Cint(KS.Setting(29))
		 Dim UserNameMaxChar:UserNameMaxChar=Cint(KS.Setting(30))
		 Dim EnabledUserName:EnabledUserName=KS.Setting(31)
		 Dim NewRegUserMoney:NewRegUserMoney=KS.Setting(38) : If Not IsNumerIc(NewRegUserMoney) Then NewRegUserMoney=0
		 Dim NewRegUserScore:NewRegUserScore=KS.Setting(39) : If Not IsNumeric(NewRegUserScore) Then NewRegUserScore=0
		 Dim NewRegUserPoint:NewRegUserPoint=KS.Setting(40) : If Not IsNumeric(NewRegUserPoint) Then NewRegUserPoint=0
         
		  If Request.ServerVariables("HTTP_REFERER")="" Then Call KS.Alert("请不要非法提交!","../"):Response.End
		  If Instr(Lcase(Request.ServerVariables("SCRIPT_NAME")),"user/reg")=0 Then Call KS.Alert("请不要非法提交!","../") : Response.End
		 
		 '收集用户资料
		 Dim Verifycode:Verifycode=KS.S("Verifycode")
		 If KS.Setting(32)="1"  Then
		  IF Trim(Verifycode)<>Trim(Session("Verifycode")) And VerificCodeTF=1 then 
		   	 Response.Write("<script>alert('验证码有误，请重新输入！');history.back(-1);</script>")
		     Exit Sub
		  End IF
		 End If
		 
		  '检查注册回答问题
		  Dim CanReg,n
		   If Mid(KS.Setting(161),1,1)="1" Then
		        CanReg=false
				 For N=0 To Ubound(Split(KS.Setting(162),vbcrlf))
				   If Trim(Request.Form("a" & MD5(n,16)))<>"" Then
					  If trim(Lcase(Request.Form("a" & MD5(n,16))))<>trim(Lcase(Split(KS.Setting(163),vbcrlf)(n))) Then
					   Call KS.AlertHistory("对不起,注册问题的回答不正确!",-1) : Response.End
					   CanReg=false
					  Else
					   CanReg=True
					  End If
				   End If
				 Next
			 If CanReg=false Then Call KS.AlertHistory("对不起,注册答案不能为空!",-1) : Response.End
		   End If
		 
		  
		 Dim UserName:UserName=KS.R(KS.S("UserName"))
		 If UserName = "" Or KS.strLength(UserName) > UserNameMaxChar Or KS.strLength(UserName) < UserNameLimitChar Then
		   	 Response.Write("<script>alert('请输入用户名(不能大于" & UserNameMaxChar & "小于" & UserNameLimitChar & ")');history.back();</script>")
			 Exit Sub
         ElseIF KS.FoundInArr(EnabledUserName, UserName, "|") = True Then
		   	 Response.Write("<script>alert('您输入的用户名为系统禁止注册的用户名');history.back();</script>")
			 Exit Sub
		 ElseIF InStr(UserName, "=") > 0 Or InStr(UserName, ".") > 0 Or InStr(UserName, "%") > 0 Or InStr(UserName, Chr(32)) > 0 Or InStr(UserName, "?") > 0 Or InStr(UserName, "&") > 0 Or InStr(UserName, ";") > 0 Or InStr(UserName, ",") > 0 Or InStr(UserName, "'") > 0 Or InStr(UserName, ",") > 0 Or InStr(UserName, Chr(34)) > 0 Or InStr(UserName, Chr(9)) > 0 Or InStr(UserName, "") > 0 Or InStr(UserName, "$") > 0 Or InStr(UserName, "*") Or InStr(UserName, "|") Or InStr(UserName, """") > 0 Then
             Response.Write("<script>alert('用户名中含有非法字符');history.back();</script>")
			 Exit Sub
        End If
		 
		 Dim PassWord,RePassWord,NoMD5_Pass
		 If Session("PassWord")<>"" Then
		   PassWord=Session("PassWord")
		 Else
			 PassWord=KS.R(KS.S("PassWord"))
			 RePassWord=KS.S("RePassWord")
			 If PassWord = "" Then
				 Response.Write("<script>alert('请输入登录密码!');history.back();</script>")
				 Exit Sub
			 ElseIF RePassWord="" Then
				 Response.Write("<script>alert('请输入确认密码');history.back();</script>")
				 Exit Sub
			 ElseIF PassWord<>RePassWord Then
				 Response.Write("<script>alert('两次输入的密码不一致');history.back();</script>")
				 Exit Sub
			 End If
		 End If
		 NoMD5_Pass=PassWord
		 Dim RndPassword:RndPassword=NoMD5_Pass
		 Dim Question:Question=KS.S("Question")
		 Dim Answer:Answer=KS.S("Answer")
		 If KS.Setting(148)="1" Then
			 If Question = "" Then
				 Response.Write("<script>alert('密码提示问题不能为空!');history.back();</script>")
				 Exit Sub
			 ElseIF Answer="" Then
				 Response.Write("<script>alert('密码答案不能为空');history.back();</script>")
				 Exit Sub
			 End If
		 End If
		 
		 Dim Email:Email=KS.S("Email")
		 if KS.IsValidEmail(Email)=false then
			 Response.Write("<script>alert('请输入正确的电子邮箱!');history.back();</script>")
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
		   if sex="男" then userface=KS.GetDomain & "Images/Face/0.gif" else userface=KS.GetDomain & "Images/face/girl.gif"	 	 
		 End If
		 Dim QQ:QQ=KS.S("QQ")		 
		 Dim ICQ:ICQ=KS.S("ICQ")		 
		 Dim MSN:MSN=KS.S("MSN")		 
		 Dim UC:UC=KS.S("UC")		 
		 Dim Sign:Sign=KS.S("Sign")	
		 Dim Privacy:Privacy=KS.ChkClng(KS.S("Privacy"))
		 Dim AllianceUser:AllianceUser=KS.S("AllianceUser")
		 
		 Dim LastLoginIP:LastLoginIP = KS.GetIP()
		 Dim CheckNum:CheckNum = KS.MakeRandomChar(6)  '随机字符验证码
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
					 Response.Write "<script>alert('" & SQL(2,K) & "必须填写!');history.back();</script>"
					 Response.End()
				  End If
				 ElseIf KS.S(SQL(0,K))="" Then
					 Response.Write "<script>alert('" & SQL(2,K) & "必须填写!');history.back();</script>"
					 Response.End()
				 End If
			   End If
			   If SQL(3,K)="4" And Not Isnumeric(KS.S(SQL(0,K))) Then 
				 Response.Write "<script>alert('" & SQL(2,K) & "必须填写数字!');history.back();</script>"
				 Response.End()
			   End If
			   If SQL(3,K)="5" And Not IsDate(KS.S(SQL(0,K))) Then 
				 Response.Write "<script>alert('" & SQL(2,K) & "必须填写正确的日期!');history.back();</script>"
				 Response.End()
			   End If
			   If SQL(3,K)="8" And Not KS.IsValidEmail(KS.S(SQL(0,K))) and SQL(1,K)="1" Then 
				Response.Write "<script>alert('" & SQL(2,K) & "必须填写正确的Email格式!');history.back();</script>"
				Response.End()
			   End If 
			 Next
		End If
		RS.Open "Select ID From KS_UserGroup Where ID=" & GroupID,conn,1,1
		If RS.Eof And RS.Bof Then
		     Rs.Close:Set RS=Nothing
			 Response.Write "<script>alert('对不起,用户组类型不正确!');history.back();</script>"
			 Response.End()
		End If
		RS.Close
		RS.Open "select top 1 * from KS_User where UserName='" & UserName & "'", Conn, 1, 3
		If Not (RS.BOF And RS.EOF) Then
				 RS.Close:Set RS=Nothing
				 Response.Write("<script>alert('您注册的用户已经存在！请换一个用户名再试试！');history.back();</script>")
				 Exit Sub
		Else
			If EmailMultiRegTF=0 Then
				Dim EmailRSCheck:Set EmailRSCheck = Conn.Execute("select UserID from KS_User where Email='" & Email & "'")
				If Not (EmailRSCheck.BOF And EmailRSCheck.EOF) Then
					EmailRSCheck.Close:Set EmailRSCheck = Nothing
					Response.Write("<script>alert('您注册的Email已经存在！请更换Email再试试！');history.back();</script>")
					Exit Sub
				End If
				EmailRSCheck.Close:Set EmailRSCheck = Nothing
			 End If
		
			 If KS.ChkClng(KS.Setting(26))=1 Then
			   If Not (Conn.Execute("select top 1 UserID From KS_User Where LastLoginIP='" & KS.GetIP & "'").eof) Then
					Response.Write("<script>alert('您的IP已经存在！不能再注册！');history.back();</script>")
					Exit Sub
			   End If
			 End If

		
		'-----------------------------------------------------------------
		'系统整合
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
		 RS("BeginDate")=Now '开始计算时间
		 RS("LastLoginIP")=LastLoginIP
		 RS("JoinDate")=Now
		 RS("LastLoginTime")=Now
		 RS("CheckNum")=CheckNum
		 RS("RndPassword")=RndPassword
		 RS("LoginTimes")=1
		 
		 '自定义字段
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

		 '新会员注册，更新相应的数据
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
		   Conn.Execute("Insert into KS_LogPoint(ChannelID,InfoID,UserName,InOrOutFlag,Point,Times,[User],Descript,Adddate,IP) values(0,0,'" & UserName & "',1," & NewRegUserPoint & ",1,'系统','注册新会员,赠送!'," & SqlNowString & ",'" & replace(ks.getip,"'","""") & "')")
		 End If
		 IF NewRegUserScore<>0 Then
		  Conn.Execute("Insert into KS_LogScore(UserName,InOrOutFlag,Score,CurrScore,[User],Descript,Adddate,IP) values('" & UserName & "',1," & NewRegUserScore & ","&NewRegUserScore& ",'系统','注册新会员,赠送!'," & SqlNowString & ",'" & replace(ks.getip,"'","""") & "')")
		 End If
		 If NewRegUserMoney<>0 Then 
		  Call KS.MoneyInOrOut(UserName,UserName,NewRegUserMoney,4,1,now,0,"System","注册新会员,赠送!",0,0)
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
	
	       Dim ReturnInfo:ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), "新用户注册激活信", Email,KS.Setting(0), MailBodyStr,KS.Setting(11))
			  IF ReturnInfo="OK" Then
			     Conn.Execute("Update KS_User Set Locked=3 where userid=" & UserID)        '设置待激活
				 ReturnInfo="<li>注册成功，注册验证码已发送到您的信箱<font color='#ff6600>'>" &Email &"</font>，只有激活后才可以正式成为本站会员!</li>"
			  Else
				ReturnInfo="<li>信件发送失败!失败原因:" & ReturnInfo & "，请联系网站管理员!</li>"
			  End if
		ElseIF EmailCheckTF=2 Then
		    Conn.Execute("Update KS_User Set Locked=2 where userid=" & UserID)        '设置需要后台认证
		    ReturnInfo="<li>注册成功!您的用户名:<font color=red>" & UserName & "</font>,您需要通过管理员的认证才能成为正式会员!</li>"
		Else
		    ReturnInfo="<li>注册成功!您的用户名:<font color=red>" & UserName & "</font>,您已成为了本站的正式会员!<br><div align=center></li>"
        End IF
		
		RS.Close
			
			
			'====================推荐计划======================================
			If AllianceUser<>"" and AllianceUser<>UserName  Then
			 If Not Conn.Execute("Select Top 1 UserID From KS_User Where UserName='" & AllianceUser & "'").eof Then
			   Call KS.ScoreInOrOut(AllianceUser,1,KS.ChkClng(KS.Setting(144)),"系统","成功推荐一个注册用户:" & UserName & "!",0,0)
			   
			   Conn.Execute("Insert InTo KS_PromotedPlan(UserName,UserIP,AddDate,ComeUrl,Score,AllianceUser) values('" & AllianceUser & "','" & KS.GetIP & "'," & SqlNowString & ",'" & KS.URLDecode(Request.ServerVariables("HTTP_REFERER")) & "'," & KS.ChkClng(KS.Setting(144)) & ",'" & UserName & "')")
			  '=================判断是不是好友邮件推荐的==================
				Dim f:f=KS.S("F")
				if f="r" Then
				 Conn.Execute("insert into KS_Friend (username,friend,addtime,flag,message,accepted) values ('"&AllianceUser&"','"& UserName &"',"&SqlNowString&",1,'',1)")
				End If
			  '============================================================

			 End If
			End If
			'====================推广计划结束=================================
			
			'==================注册成功发邮件通知======================
			 If KS.Setting(146)="1" and Not KS.IsNul(KS.Setting(147)) And EmailCheckTF<>1 Then
				MailBodyStr = Replace(KS.Setting(147), "{$UserName}", UserName)
				MailBodyStr = Replace(MailBodyStr, "{$PassWord}", NoMD5_Pass)
				MailBodyStr = Replace(MailBodyStr, "{$SiteName}", KS.Setting(0))
				ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), KS.Setting(0) & "-会员注册成功", Email,UserName, MailBodyStr,KS.Setting(11))
				IF ReturnInfo="OK" Then
				  ReturnInfo="<li>注册成功!您的用户名:<font color=red>" & UserName & "</font>,已将用户名和密码发到您的信箱!</li>"
				End If
			 End If
			'==========================================================
			
			
			
		    '===================写入个人空间================
			if KS.SSetting(0)=1 And KS.SSetting(1)=1 then
			 RS.Open "Select * From KS_Blog Where 1=0",conn,1,3
			 RS.AddNew
			  RS("UserName")=UserName
			  RS("BlogName")=UserName & "的个人空间"
			  RS("ClassID") = KS.ChkClng(Conn.Execute("Select Top 1 ClassID From KS_BlogClass")(0))
			  If UserType=1 Then
			  RS("TemplateID")=KS.ChkClng(Conn.Execute("Select Top 1 ID From KS_BlogTemplate Where flag=4 and IsDefault='true'")(0))
			  Else
			  RS("TemplateID")=KS.ChkClng(Conn.Execute("Select Top 1 ID From KS_BlogTemplate Where flag=2 and IsDefault='true'")(0))
			  End If
			  RS("Announce")="暂无公告!"
			  RS("ContentLen")=500
			  RS("Recommend")=0
			  if KS.SSetting(2)=1 then
			  RS("Status")=0
			  else
			  RS("Status")=1
			  end if
			 RS.Update
			 RS.Close
			 '判断是企业会员，自动开通企业空间
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
			'系统整合
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
		   ReturnInfo="<table border='0' align='center' width='60%' cellspacing='1' cellpadding='3'><tr class='tdbg'><td></td></tr><tr class='tdbg'><td><img src='../images/regok.jpg' align='left'>" & ReturnInfo & "<li><span id=""countdown""> <font color='#ff6600'><strong>10</strong></font> </span>秒钟后自动转到会员中心!</li><li><a href=""../index.asp"">马上进入会员中心</a>  <a href=""" &  KS.Setting(3)& """>返回网站首页</a></li></td></tr></table>" & vbcrlf
		   Dim FileContent
		   If KS.Setting(119)="" Then Response.Write "请先到""基本信息设置->模板绑定""进行模板绑定操作!":response.end
		    FileContent = KSRFObj.LoadTemplate(KS.Setting(119)) &  GetTurnMember(10)
            FileContent= Replace(FileContent,"{$GetUserRegResult}",ReturnInfo)
			FileContent = KSRFObj.KSLabelReplaceAll(FileContent) '替换函数标签
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

 
