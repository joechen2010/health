<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.EscapeCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="gb2312"
Dim KSCls
Set KSCls = New Ajax_Check
KSCls.Kesion()
Set KSCls = Nothing

Class Ajax_Check
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		  
		  Select Case KS.S("Action")
		   Case "checkusername"
		    Call CheckUserName()
		   Case "checkemail"
		    Call CheckEmail()
		   Case "checkcode"
		    Call CheckCode()
		  End Select
		End Sub
		
		Sub CheckUserName()
			dim username:username=UnEscape(KS.S("username"))
			if username="" then
			 KS.Echo escape("err|请输入会员名！")
			elseif InStr(UserName, "=") > 0 Or InStr(UserName, ".") > 0 Or InStr(UserName, "%") > 0 Or InStr(UserName, Chr(32)) > 0 Or InStr(UserName, "?") > 0 Or InStr(UserName, "&") > 0 Or InStr(UserName, ";") > 0 Or InStr(UserName, ",") > 0 Or InStr(UserName, "'") > 0 Or InStr(UserName, ",") > 0 Or InStr(UserName, Chr(34)) > 0 Or InStr(UserName, Chr(9)) > 0 Or InStr(UserName, "") > 0 Or InStr(UserName, "$") > 0 Or InStr(UserName, "*") Or InStr(UserName, "|") Or InStr(UserName, """") > 0 Then
			KS.Echo escape("err|用户名中含有非法字符!")
			elseif KS.StrLength(username)<KS.ChkClng(KS.Setting(29)) or KS.StrLength(username)>KS.ChkClng(KS.Setting(30)) then
			 KS.Echo escape("err|输入的会员名长度应为<font color=#ff6600>" & KS.Setting(29) &"-" & KS.Setting(30) & "位</font>！")
			elseif KS.FoundInArr(KS.Setting(31), UserName, "|") = True Then
			 KS.Echo escape("err|您输入的用户名为系统禁止注册的用户名</font>！")
			elseif conn.Execute("Select Userid From KS_User where username='"&username&"'" ).eof Then
			 KS.Echo escape("ok|恭喜,该会员名可以正常注册！")
			else
			 KS.Echo escape("err|该会员名已经有人使用!")
			end if
		End Sub
		Sub CheckEmail()
			dim email:email=unescape(KS.S("email"))
			if email="" then
			 KS.Echo escape("err|请输入电子邮箱！")
			elseif instr(email,"@")=0 or instr(email,".")=0 then
			 KS.Echo escape("err|您输入电子邮箱有误！")
			elseif ks.setting(28)=1 or conn.Execute("Select userid From KS_User where email='"&email&"'" ).eof Then
			 KS.Echo escape("ok|该邮箱可以正常注册!")
			else
			 KS.Echo escape("err|该邮箱已经有人使用，请重新选择。")
			end if
		End Sub
		Sub CheckCode()
		  dim code:code=unescape(KS.S("code"))
		  IF Trim(code)<>Trim(Session("Verifycode")) And KS.ChkCLng(KS.Setting(27))=1 then 
		   	 KS.Echo escape("err|验证码有误，请重新输入！")
		  Else
		   	 KS.Echo escape("ok|验证码已输入！")
		  End IF
		End Sub
End Class
%> 
