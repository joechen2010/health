<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.ContentType="text/vnd.wap.wml"
Response.Charset="utf-8"
Response.Write "<?xml version=""1.0"" encoding=""utf-8""?>" &vbcrlf
Response.Write "<!DOCTYPE wml PUBLIC ""-//WAPFORUM//DTD WML 1.1//EN"" ""http://www.wapforum.org/DTD/wml_1.1.xml"">" &vbcrlf
%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Md5.asp"-->
<%
Dim KSCls
Set KSCls = New User_RegPost
KSCls.Kesion()
Set KSCls = Nothing

Class User_RegPost
        Private KS,ToUrl
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		    Call CloseConn()
		    Set KS=Nothing
		End Sub
		Public Sub Kesion()
		    '系统配置参数
			'Dim AdminCheckTF:AdminCheckTF=KS.Setting(26)
			Dim Locked:Locked=0
			Dim VerificCodeTF:VerificCodeTF=KS.Setting(27)
			Dim EmailMultiRegTF:EmailMultiRegTF=KS.Setting(28)
			Dim UserNameLimitChar:UserNameLimitChar=Cint(KS.Setting(29))
			Dim UserNameMaxChar:UserNameMaxChar=Cint(KS.Setting(30))
			Dim EnabledUserName:EnabledUserName=KS.Setting(31)
			Dim NewRegUserMoney:NewRegUserMoney=KS.Setting(38)
			Dim NewRegUserScore:NewRegUserScore=KS.Setting(39)
			Dim NewRegUserPoint:NewRegUserPoint=KS.Setting(40)
			Dim GroupID:GroupID=KS.ChkClng(KS.S("GroupID")):If GroupID=0 Then GroupID=2
			'收集用户资料
			Dim Verifycode:Verifycode=KS.S("Verifycode")
			If KS.Setting(32)="1"  Then
			   IF Trim(Verifycode)<>Trim(Session("Verifycode")) And VerificCodeTF=1 Then
			      Call KS.ShowError("错误提示","验证码有误，请重新输入！")
			   End IF
		    End If
			
			Dim UserName:UserName=KS.R(KS.S("UserName"))
			If UserName = "" Or KS.strLength(UserName) > UserNameMaxChar Or KS.strLength(UserName) < UserNameLimitChar Then
			   Call KS.ShowError("错误提示","请输入用户名(不能大于" & UserNameMaxChar & "小于" & UserNameLimitChar & ")")
			   Exit Sub
            ElseIF KS.FoundInArr(EnabledUserName, UserName, "|") = True Then
			   Call KS.ShowError("错误提示","您输入的用户名为系统禁止注册的用户名")
		    ElseIF InStr(UserName, "=") > 0 Or InStr(UserName, ".") > 0 Or InStr(UserName, "%") > 0 Or InStr(UserName, Chr(32)) > 0 Or InStr(UserName, "?") > 0 Or InStr(UserName, "&") > 0 Or InStr(UserName, ";") > 0 Or InStr(UserName, ",") > 0 Or InStr(UserName, "'") > 0 Or InStr(UserName, ",") > 0 Or InStr(UserName, Chr(34)) > 0 Or InStr(UserName, Chr(9)) > 0 Or InStr(UserName, "") > 0 Or InStr(UserName, "$") > 0 Or InStr(UserName, "*") Or InStr(UserName, "|") Or InStr(UserName, """") > 0 Then
			   Call KS.ShowError("错误提示","用户名中含有非法字符")
			ElseIF Cbool(KS.IsValidChars(UserName))=False Then
			   Call KS.ShowError("错误提示","用户名请使用英文和数字！")
            End If
			Dim RndPassword:RndPassword=KS.R(KS.MakeRandomChar(20))
			Dim PassWord,RePassWord,NoMD5_Pass
			If Session("PassWord")<>"" Then
			   PassWord=Session("PassWord")
			Else
			   PassWord=KS.R(KS.S("PassWord"))
			   RePassWord=KS.S("RePassWord")
			   If PassWord = "" Then
			      Call KS.ShowError("错误提示","请输入登录密码!")
				  Exit Sub
			   'ElseIF RePassWord="" Then
			      'Call KS.ShowError("错误提示","请输入确认密码!")
				  'Exit Sub
			   'ElseIF PassWord<>RePassWord Then
			      'Call KS.ShowError("错误提示","两次输入的密码不一致!")
				  'Exit Sub
			   End If
		    End If
			NoMD5_Pass=PassWord
			Dim Question:Question=KS.S("Question")
			Dim Answer:Answer=KS.S("Answer")
			If KS.Setting(148)="1" Then
			   If Question = "" Then
			      Call KS.ShowError("错误提示","密码提示问题不能为空!")
				  Exit Sub
			   ElseIF Answer="" Then
			      Call KS.ShowError("错误提示","密码答案不能为空!")
				  Exit Sub
			   End If
		    End If
			Dim Email:Email=KS.S("Email")
			
			If KS.ChkClng(KS.Setting(146))=1 Then
			   If KS.IsValidEmail(Email)=false Then
			      Call KS.ShowError("错误提示","请输入正确的电子邮箱!")
		       End If
			End If
			
			Dim RealName:RealName=KS.S("RealName")
			Dim Sex:Sex=KS.S("Sex")
			Dim Birthday:Birthday=KS.S("Birthday")
			If Not IsDate(Birthday) Then Birthday=Now
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
			Dim UserFace:UserFace=KS.S("UserFace"):if userface="" then userface="Images/Face/0.gif"		 	 	 
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
			CheckUrl=KS.GetDomain &"User/reg/active.asp?UserName=" & UserName &"&CheckNum=" & CheckNum
			
			PassWord =MD5(KS.R(PassWord),16)
			Dim RS,SQL,K
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
					    ' If KS.S("Province")="" and  KS.S("City")="" Then
						'    Call KS.ShowError("错误提示","" & SQL(2,K) & "必须填写!")
						' End If
					  ElseIf KS.S(SQL(0,K))="" Then
					     Call KS.ShowError("错误提示","" & SQL(2,K) & "必须填写!")
					  End If
				   End If
				   If SQL(3,K)="4" And Not Isnumeric(KS.S(SQL(0,K))) Then 
				      Call KS.ShowError("错误提示","" & SQL(2,K) & "必须填写数字!")
				   End If
				   If SQL(3,K)="5" And Not IsDate(KS.S(SQL(0,K))) Then
				      Call KS.ShowError("错误提示","" & SQL(2,K) & "必须填写正确的日期!")
				   End If
				   If SQL(3,K)="8" And Not KS.IsValidEmail(KS.S(SQL(0,K))) and SQL(1,K)="1" Then
				      Call KS.ShowError("错误提示","" & SQL(2,K) & "必须填写正确的Email格式!")
				   End If 
			   Next
			End If
			RS.Open "select top 1 * from KS_User where UserName='" & UserName & "'", Conn, 1, 3
			If Not (RS.BOF And RS.EOF) Then
			   RS.Close:Set RS=Nothing
			   Call KS.ShowError("错误提示","您注册的用户已经存在！请换一个用户名再试试！")
			Else
			   If KS.ChkClng(KS.Setting(146))=1 And EmailMultiRegTF=0 Then
				  Dim EmailRSCheck:Set EmailRSCheck = Conn.Execute("select UserID from KS_User where Email='" & Email & "'")
				  If Not (EmailRSCheck.BOF And EmailRSCheck.EOF) Then
					 EmailRSCheck.Close:Set EmailRSCheck = Nothing
					 Call KS.ShowError("错误提示","您注册的Email已经存在！请更换Email再试试！")
				  End If
				  EmailRSCheck.Close:Set EmailRSCheck = Nothing
			   End If
			   If KS.ChkClng(KS.Setting(149))=1 Then
			      IF Len(Mobile)<>11 Then
				     Call KS.ShowError("错误提示","您注册的手机号码有误！")
				  End IF
				  If Not (Conn.Execute("select UserID from KS_User where Mobile='" & Mobile & "'").EOF) Then
				     Call KS.ShowError("错误提示","您注册的手机号码已经存在！")
				  End IF
			   End If
			   If KS.ChkClng(KS.Setting(26))=1 Then
			      If Not (Conn.Execute("select UserID From KS_User Where LastLoginIP='" & KS.GetIP & "'").EOF) Then
				     Call KS.ShowError("错误提示","您的IP已经存在！不能再注册！")
			      End If
			   End If
			   
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
			   RS("Wap")=MD5(UserName&RndPassword,32)
			   RS("LoginTimes")=1
			   '自定义字段
			   For K=0 To UBound(SQL,2)
			       If left(Lcase(SQL(0,K)),3)="ks_" Then
				      RS(SQL(0,K))=KS.S(SQL(0,K))
				   End If
			   Next
			   '=======================增加加盟号开始============
			   RS("AllianceUser")=AllianceUser
			   '=======================增加加盟号结束===============
			   '新会员注册，更新相应的数据
			   RS("Money")=NewRegUserMoney
			   RS("Score")=NewRegUserScore
			   RS("Point")=NewRegUserPoint
			   Call KS.PointInOrOut(0,0,UserName,1,NewRegUserPoint,"系统","注册新会员,赠送的" & KS.Setting(46) & KS.Setting(45))
			   RS("Locked")=Locked
			   RS.Update
			   RS.MoveLast
			 Dim UserID:UserID=RS("UserID")
			   ToUrl = Replace(Replace(KS.S("ToUrl"),"&amp;","&"),"&","&amp;")
			   If ToUrl = "" Then ToUrl = "../Index.asp"
			   ToUrl = KS.JoinChar(ToUrl)
			   ToUrl = ToUrl & KS.WSetting(2) & "=" & RS("Wap") & ""
			 RS.Close
			   
				 If NewRegUserPoint<>0 Then
				   Conn.Execute("Insert into KS_LogPoint(ChannelID,InfoID,UserName,InOrOutFlag,Point,Times,[User],Descript,Adddate,IP,CurrPoint,ContributeFlag) values(0,0,'" & UserName & "',1," & NewRegUserPoint & ",1,'系统','注册新会员,赠送!'," & SqlNowString & ",'" & replace(ks.getip,"'","""") & "'," & NewRegUserPoint &",0)")
				 End If
				 IF NewRegUserScore<>0 Then
				  Conn.Execute("Insert into KS_LogScore(UserName,InOrOutFlag,Score,CurrScore,[User],Descript,Adddate,IP) values('" & UserName & "',1," & NewRegUserScore & ","&NewRegUserScore& ",'系统','注册新会员,赠送!'," & SqlNowString & ",'" & replace(ks.getip,"'","""") & "')")
				 End If
				 If NewRegUserMoney<>0 Then 
				  Conn.Execute("Insert into KS_LogMoney([UserName],[ClientName],[Money],[MoneyType],[IncomeOrPayOut],[OrderID],[Remark],[PayTime],[LogTime],[Inputer],[IP],[CurrMoney],[ChannelID],[InfoID]) values('" & UserName & "','" & UserName & "'," & NewRegUserMoney & ",4,1,'0','注册新会员,赠送!'," & SqlNowString & "," &SqlNowString & ",'系统','" & replace(ks.getip,"'","""") & "'," & NewRegUserMoney & ",0,0)")
				 End If			   
			   
			   
			     RS.Open "Select * From KS_UserGroup Where ID=" & GroupID,conn,1,1
				 If RS.Eof Then RS.Close : Set RS=Nothing 
				 
				 Dim EmailCheckTF:EmailCheckTF=RS("ValidType")
				 Dim UserRegSendMail:UserRegSendMail=RS("ValidEmail")
				 Dim UserType:UserType= RS("UserType")
				 Conn.Execute("Update KS_User Set ChargeType=" & RS("ChargeType") & ",EDays=" & RS("ValidDays") & ",UserType=" & UserType &" Where UserID=" & UserID)


			   Dim MailBodyStr
			   If EmailCheckTF = 1 Then
			      MailBodyStr = Replace(UserRegSendMail, "{$CheckNum}", CheckNum)
			      MailBodyStr = Replace(MailBodyStr, "{$CheckUrl}", CheckUrl)
				  Dim ReturnInfo:ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), KS.Setting(0) & "-会员注册激活信", Email,UserName, MailBodyStr,KS.Setting(11))
				  IF ReturnInfo="OK" Then
				     Conn.Execute("Update KS_User Set Locked=3 where userid=" & UserID)        '设置待激活
				     ReturnInfo="注册成功，注册验证码已发送到您的信箱<b>" &Email &"</b>，只有激活后才可以正式成为本站会员!<br/>"
					 ToUrl="../../"
				  Else
				     ReturnInfo="信件发送失败!失败原因:" & ReturnInfo & "，请联系网站管理员!</li>"
				  End if
			   ElseIF EmailCheckTF=2 Then
			      Conn.Execute("Update KS_User Set Locked=2 where userid=" & UserID)        '设置需要后台认证
			      ReturnInfo="注册成功!您的用户名:<b>" & UserName & "</b>,您需要通过管理员的认证才能成为正式会员!<br/>"
				  ToUrl="../../"
			   Else
			      ReturnInfo="注册成功!您的用户名:<b>" & UserName & "</b>,您已成为了本站的正式会员!<br/>"
			   End IF
			   RS.Close
			   
			   '====================推荐计划开始======================================
			   If AllianceUser<>"" Then
			      If Not Conn.Execute("Select Top 1 UserID From KS_User Where UserName='" & AllianceUser & "'").EOF Then
				     Dim AllianceUrl:AllianceUrl=KS.S("AllianceUrl")
					 If AllianceUrl="" Then AllianceUrl="★直接手机输入或书签导入★"
				     Conn.Execute("Update KS_User Set Score=Score+" & KS.ChkClng(KS.Setting(144)) & " Where UserName='" & AllianceUser & "'")
					 Conn.Execute("Insert InTo KS_PromotedPlan(UserName,UserIP,AddDate,ComeUrl,Score,AllianceUser) values('" & AllianceUser & "','" & KS.GetIP & "'," & SqlNowString & ",'" & AllianceUrl & "'," & KS.ChkClng(KS.Setting(144)) & ",'" & UserName & "')")
				  End If
			   End If
			   '====================推广计划结束=================================
			   
			   '==================注册成功发邮件通知======================
			   If KS.Setting(146)="1" and KS.Setting(147)<>"" Then
				  MailBodyStr = Replace(KS.Setting(147), "{$UserName}", UserName)
				  MailBodyStr = Replace(MailBodyStr, "{$PassWord}", NoMD5_Pass)
				  MailBodyStr = Replace(MailBodyStr, "{$SiteName}", KS.Setting(0))
				  ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), KS.Setting(0) & "-会员注册成功", Email,UserName, MailBodyStr,KS.Setting(11))
				  IF ReturnInfo="OK" Then
				     ReturnInfo="注册成功!您的用户名:<b>" & UserName & "</b>,已将用户名和密码发到您的信箱!<br/>"
				  Else
				     ReturnInfo="注册成功!您的用户名:<b>" & UserName & "</b>,系统原因未将用户名和密码发到您的信箱!<br/>"
				  End If
			   End If
			   '==========================================================
			   
			   '===================写入个人空间================
			   If KS.SSetting(0)=1 And KS.SSetting(1)=1 Then
			      RS.Open "Select top 1 * From KS_Blog Where 1=0",Conn,1,3
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
				  If KS.SSetting(2)=1 Then
				     RS("Status")=0
				  Else
				     RS("Status")=1
				  End If
				  RS.Update
				  RS.Close
				  
				  '判断是企业会员，自动开通企业空间
				  On Error Resume Next
				  If UserType=1 then
				     RS.Open "Select * From KS_EnterPrise Where 1=0",Conn,1,3
				     RS.AddNew
				     RS("UserName")=UserName
					 RS("CompanyName")=KS.S("KS_companyname")
					 RS("Province")=Province
					 RS("City")=City
					 RS("Address")=Address
					 RS("ZipCode")=Zip
					 RS("ContactMan")=RealName
					 RS("TelPhone")=OfficeTel
					 RS("Fax")=Fax
					 RS("AddTime")=Now
					 RS("Recommend")=0
					 RS("ClassID")=0
					 RS("SmallClassID")=0
					 If KS.SSetting(2)=1 Then
					    RS("Status")=0
				     Else
					    RS("Status")=1
					 End If
					 RS.Update 
					 RS.Close
				  End If
			   End If
			   '==================================
			   Set RS=Nothing
			   Call ShowRegResult(ReturnInfo)
			End If
		End Sub
		
		Sub ShowRegResult(ReturnInfo)
		    ReturnInfo="注册结果:<br/>" & ReturnInfo & ""
			ReturnInfo=ReturnInfo&"<a href=""" & ToUrl & """>马上进入...</a><br/><br/>"
			ReturnInfo=ReturnInfo&"<a href=""" & ToUrl & """>返回网站首页</a>" & vbcrlf
			Response.Write "<wml>" &vbcrlf
			Response.Write "<head>" &vbcrlf
			Response.Write "<meta http-equiv=""Cache-Control"" content=""no-Cache""/>" &vbcrlf
			Response.Write "<meta http-equiv=""Cache-Control"" content=""max-age=0""/>" &vbcrlf
			Response.Write "</head>" &vbcrlf
			Response.Write "<card title=""正在进入.."" newcontext=""true"" ontimer=""" & ToUrl & """><timer value=""3""/>" &vbcrlf
			'Response.Write "<card title=""正在进入.."">" &vbcrlf
			Response.Write "<p align=""left"">" &vbcrlf
            Response.Write ReturnInfo
			Response.Write "</p>" &vbcrlf
			Response.Write "</card>" &vbcrlf
			Response.Write "</wml>"
		End Sub
End Class
%>

 
