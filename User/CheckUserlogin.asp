<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../Plus/md5.asp"-->
<!--#include file="../API/cls_api.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New UserLogin
KSCls.Kesion()
Set KSCls = Nothing

Class UserLogin
        Private KS
		Private KSUser
		Private UserName,PassWord,Verifycode,ExpiresDate,RndPassword
		Private LoginVerificCodeTF
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		
		
		Public Sub Kesion()
			UserName=KS.R(KS.S("UserName"))
			PassWord=KS.R(KS.S("PassWord"))
			ExpiresDate=KS.R(KS.S("ExpiresDate"))
			Verifycode=KS.R(KS.S("Verifycode"))
			LoginVerificCodeTF=KS.Setting(34)
			RndPassword=KS.R(KS.MakeRandomChar(20))
			IF UserName="" Then
		   	 KS.Die "<script>alert('用户名不能为空，请输入！');history.back();</script>"
			End IF
		    IF PassWord="" Then
		   	 KS.Die "<script>alert('登录密码不能为空，请输入！');history.back();</script>"
			End IF
			IF Trim(Verifycode)<>Trim(Session("Verifycode")) And LoginVerificCodeTF=1 then 
		   	 KS.Die "<script>alert('验证码有误，请重新输入！');history.back();</script>"
			End IF
            
			
			PassWord=MD5(PassWord,16)
			Dim UserRS:Set UserRS=Server.CreateObject("Adodb.RecordSet")
			 UserRS.Open "Select top 1 UserID,UserName,PassWord,Locked,GroupID,Score,LastLoginIP,LastLoginTime,LoginTimes,RndPassword,IsOnline,GradeTitle,UserCardID,Point,Money,Edays,BeginDate From KS_User Where UserName='" &UserName & "' And PassWord='" & PassWord & "'",Conn,1,3
			 If UserRS.Eof And UserRS.BOf Then
				  UserRS.Close:Set UserRS=Nothing
				  KS.Die "<script>alert('你输入的用户名或密码有误，请重新输入！');history.back();</script>"
			 ElseIf UserRS("Locked")=1 Then
			   UserRS.Close:Set UserRS=Nothing
			   KS.Die "<script>alert('您的账号已被管理员锁定，请与管理员联系！');history.back();</script>"
			 ElseIF UserRS("Locked")=3 Then
			   UserRS.Close:Set UserRS=Nothing
			   KS.Die "<script>alert('您的账号还没有激活，请注意查收您的邮箱并进行激活！');history.back();</script>"
			 ElseIF UserRS("Locked")=2 Then
			   UserRS.Close:Set UserRS=Nothing
			   KS.Die "<script>alert('您的账号还没有通过认证！');history.back();</script>"
			 Else
			        	'-----------------------------------------------------------------
						'系统整合
						'-----------------------------------------------------------------
						Dim API_KS,API_SaveCookie,SysKey
						If API_Enable Then
							Set API_KS = New API_Conformity
							API_KS.NodeValue "action","login",0,False
							API_KS.NodeValue "username",UserName,1,False
							Md5OLD = 1
							SysKey = Md5(API_KS.XmlNode("username") & API_ConformKey,16)
							Md5OLD = 0
							API_KS.NodeValue "syskey",SysKey,0,False
							API_KS.NodeValue "password",KS.R(KS.S("PassWord")),0,False
							API_KS.SendHttpData
							If API_KS.Status = "1" Then
								KS.Die "<script>alert('" & API_KS.Message & "');history.back();</script>"
							Else
							    Dim usercookies:usercookies=1
								API_SaveCookie = API_KS.SetCookie(SysKey,UserName,Password,usercookies)
							End If
							Set API_KS = Nothing
						End If
						'-----------------------------------------------------------------
			 
			            '登录成功，更新用户相应的数据
						Dim ScoreTF:ScoreTF=False
						If KS.ChkClng(KS.U_S(UserRS("GroupID"),8))>0 and KS.ChkClng(KS.U_S(UserRS("GroupID"),9))>0 And datediff("n",UserRS("LastLoginTime"),now)>=KS.ChkClng(KS.U_S(UserRS("GroupID"),8)) then '判断时间
						ScoreTF=true
						End if
						UserRS("LastLoginIP") = KS.GetIP
                        UserRS("LastLoginTime") = Now()
                        UserRS("LoginTimes") = UserRS("LoginTimes") + 1
						UserRS("RndPassword")= RndPassword
						UserRS("IsOnline")=1
						'判断上一次是不是通过充值卡充值
						If UserRS("UserCardID")<>0 Then
						  Dim RSCard,ValidUnit,ExpireGroupID
						  Set RSCard=Conn.Execute("Select top 1 * From KS_UserCard Where ID=" & UserRS("UserCardID"))
						  If Not RSCard.Eof Then
						     ValidUnit=RSCard("ValidUnit")
							 ExpireGroupID=RSCard("ExpireGroupID")
							 If ValidUnit=1 Then                      '点券
							   If UserRS("Point")<=0 And ExpireGroupID<>0 Then
							     UserRS("GroupID")=ExpireGroupID
								 UserRS("UserCardID")=0
							   End If
							 ElseIf ValidUnit=2 Then                   '有效天数
							   If UserRS("Edays")-DateDiff("D",UserRS("BeginDate"),now())<=0 And ExpireGroupID<>0 Then
							     UserRS("GroupID")=ExpireGroupID
								 UserRS("UserCardID")=0
							   End If 
							 ElseIf ValidUnit=3 Then                  '资金
							   If UserRS("Money")<=0 And ExpireGroupID<>0 Then
							     UserRS("GroupID")=ExpireGroupID
								 UserRS("UserCardID")=0
							   End If
							 End If
						  End If
						  RSCard.Close : Set RSCard=Nothing
						End If
                        UserRS.Update
						
						on error resume next
						UserRS("GradeTitle")=Conn.Execute("select top 1 usertitle from KS_AskGrade where score<=" & UserRS("Score") & " order by score desc")(0)
						
                        UserRS.Update
						if err then err.clear
						
						
						If ScoreTF then 
						Call KS.ScoreInOrOut(UserName,1,KS.ChkClng(KS.U_S(UserRS("GroupID"),9)),"系统",KS.ChkClng(KS.U_S(UserRS("GroupID"),8)) & "分钟后,重新登录奖励获得",0,0)
						End if
						
						'更新购物车的ID号
						If Not IsNul(KS.C("CartID")) Then
						Conn.Execute("Update KS_ShopPackageSelect Set UserName='" & UserName & "' where username='" & KS.C("CartID") & "'")
						End If
												
                            Response.Cookies(KS.SiteSn).path = "/"
						    If ExpiresDate<>"" Then Response.Cookies(KS.SiteSn).Expires = Date + 365
							Response.Cookies(KS.SiteSn)("UserName") = UserName
							Response.Cookies(KS.SiteSn)("Password") = Password
							Response.Cookies(KS.SiteSN)("RndPassword")= RndPassword
								'-----------------------------------------------------------------
								'系统整合
								'-----------------------------------------------------------------
								
								If API_Enable Then
									Response.Write API_SaveCookie
									Response.Flush
									If API_LoginUrl <> "0" Then
										Response.Write "<script language=JavaScript>"
										Response.Write "setTimeout(""top.location='"& API_LoginUrl &"'"",1000);"
										Response.Write "</script>"
										Response.End
									End If
								End If
								If KS.S("Action")="PopLogin" Then
								 response.write "<script>window.parent.location.reload(); </script>"
								Else
									Dim ToUrl
									If InStr(lcase(Request.ServerVariables("HTTP_REFERER")), "/login") > 0 Then 
									     ToUrl="index.asp"
									ElseIf InStr(lcase(Request.ServerVariables("HTTP_REFERER")), "login") > 0 Then 
										 ToUrl= KS.GetDomain & "User/userlogin.asp?action=" & KS.S("Action")
									else
										 ToUrl= Request.ServerVariables("HTTP_REFERER")
									end if
									if GCls.ComeUrl<>"" then 
									 ToUrl=GCls.ComeUrl
									 GCls.ComeUrl=""
									 response.write "<script>top.location.href='" & ToUrl & "';</script>"
									Else
									 response.write "<script>location.href='" & ToUrl & "';</script>"
									End If
								End If
			 End IF
			
        End Sub
End Class
%>

 
