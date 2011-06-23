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
<!--#include file="../KS_Cls/Md5.asp"-->
<%
Dim KSCls
Set KSCls = New UserLogin
KSCls.Kesion()
Set KSCls = Nothing

Class UserLogin
        Private KS
		Private UserName,PassWord,Verifycode,ExpiresDate,RndPassword
		Private LoginVerificCodeTF
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		    Call CloseConn()
		    Set KS=Nothing
		End Sub
		Public Sub Kesion()
			UserName=KS.R(KS.S("UserName"))
			PassWord=KS.R(KS.S("PassWord"))
			ExpiresDate=KS.R(KS.S("ExpiresDate"))
			Verifycode=KS.R(KS.S("Verifycode"))
			LoginVerificCodeTF=KS.Setting(34)
			RndPassword=KS.R(KS.MakeRandomChar(20))
			IF UserName="" Then
			   Call KS.ShowError("错误提示","用户名不能为空，请输入！")
			End IF
		    IF PassWord="" Then
			   Call KS.ShowError("错误提示","登录密码不能为空，请输入！")
			End IF
			IF Trim(Verifycode)<>Trim(Session("Verifycode")) And LoginVerificCodeTF=1 Then
			   Call KS.ShowError("错误提示","验证码有误，请重新输入！")
			End IF
			
			PassWord=MD5(PassWord,16)
			Dim UserRS:Set UserRS=Server.CreateObject("Adodb.RecordSet")
			 UserRS.Open "Select UserID,UserName,PassWord,Locked,GroupID,Score,LastLoginIP,LastLoginTime,LoginTimes,RndPassword,IsOnline,GradeTitle,wap From KS_User Where UserName='" &UserName & "' And PassWord='" & PassWord & "'",Conn,1,3
			If UserRS.Eof And UserRS.BOf Then
			   UserRS.Close:Set UserRS=Nothing
			   Call KS.ShowError("错误提示","你输入的用户名或密码有误，请重新输入！")
			ElseIf UserRS("Locked")=1 Then
			   UserRS.Close:Set UserRS=Nothing
			   Call KS.ShowError("错误提示","您的账号已被管理员锁定，请与管理员联系！")
			ElseIF UserRS("Locked")=3 Then
			   UserRS.Close:Set UserRS=Nothing
			   Call KS.ShowError("错误提示","您的账号还没有激活，请注意查收您的邮箱并进行激活！")
			ElseIF UserRS("Locked")=2 Then
			   UserRS.Close:Set UserRS=Nothing
			   Call KS.ShowError("错误提示","您的账号还没有通过认证！")
			Else
			   '登录成功，更新用户相应的数据
			   If datediff("n",UserRS("LastLoginTime"),now)>=KS.ChkClng(KS.Setting(36)) Then '判断时间
			      UserRS("Score")=UserRS("Score")+KS.ChkClng(KS.Setting(37))
			   End If
			   UserRS("LastLoginIP") = KS.GetIP
			   UserRS("LastLoginTime") = Now()
			   UserRS("LoginTimes") = UserRS("LoginTimes") + 1
			   UserRS("RndPassword")= RndPassword
			   UserRS("IsOnline")=1
			   If KS.IsNul(UserRS("wap")) or KS.strLength(UserRS("wap")) < 32 Then UserRS("wap")= MD5(KS.MakeRandomChar(20),32)
			   UserRS.Update
			   '***************************************
				on error resume next
				UserRS("GradeTitle")=Conn.Execute("select top 1 usertitle from KS_AskGrade where score<=" & UserRS("Score") & " order by score desc")(0)
						
                 UserRS.Update
						
			   
			   
			   Dim ToUrl
			   ToUrl = KS.S("ToUrl")
			   ToUrl = Replace(Replace(ToUrl,"&amp;","&"),"&","&amp;")
			   If ToUrl = "" Then ToUrl = "Index.asp"
			   ToUrl = KS.JoinChar(ToUrl)
			   ToUrl = ToUrl & KS.WSetting(2) & "=" & UserRS("wap") & ""
			   '***************************************
			   Response.Write "<wml>" &vbcrlf
			   Response.Write "<head>" &vbcrlf
			   Response.Write "<meta http-equiv=""Cache-Control"" content=""no-Cache""/>" &vbcrlf
			   Response.Write "<meta http-equiv=""Cache-Control"" content=""max-age=0""/>" &vbcrlf
			   Response.Write "</head>" &vbcrlf
			   Response.Write "<card title=""正在进入.."" newcontext=""true"" ontimer=""" & ToUrl & """><timer value=""3""/>" &vbcrlf
			   Response.Write "<p align=""left"">" &vbcrlf
			   Response.Write "登陆成功...<a href="""&ToUrl&""">马上进入</a>"
			   Response.Write "</p>" &vbcrlf
			   Response.Write "</card>" &vbcrlf
			   Response.Write "</wml>"   
			End If
        End Sub
End Class
%>