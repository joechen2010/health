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
			 KS.Echo escape("err|�������Ա����")
			elseif InStr(UserName, "=") > 0 Or InStr(UserName, ".") > 0 Or InStr(UserName, "%") > 0 Or InStr(UserName, Chr(32)) > 0 Or InStr(UserName, "?") > 0 Or InStr(UserName, "&") > 0 Or InStr(UserName, ";") > 0 Or InStr(UserName, ",") > 0 Or InStr(UserName, "'") > 0 Or InStr(UserName, ",") > 0 Or InStr(UserName, Chr(34)) > 0 Or InStr(UserName, Chr(9)) > 0 Or InStr(UserName, "��") > 0 Or InStr(UserName, "$") > 0 Or InStr(UserName, "*") Or InStr(UserName, "|") Or InStr(UserName, """") > 0 Then
			KS.Echo escape("err|�û����к��зǷ��ַ�!")
			elseif KS.StrLength(username)<KS.ChkClng(KS.Setting(29)) or KS.StrLength(username)>KS.ChkClng(KS.Setting(30)) then
			 KS.Echo escape("err|����Ļ�Ա������ӦΪ<font color=#ff6600>" & KS.Setting(29) &"-" & KS.Setting(30) & "λ</font>��")
			elseif KS.FoundInArr(KS.Setting(31), UserName, "|") = True Then
			 KS.Echo escape("err|��������û���Ϊϵͳ��ֹע����û���</font>��")
			elseif conn.Execute("Select Userid From KS_User where username='"&username&"'" ).eof Then
			 KS.Echo escape("ok|��ϲ,�û�Ա����������ע�ᣡ")
			else
			 KS.Echo escape("err|�û�Ա���Ѿ�����ʹ��!")
			end if
		End Sub
		Sub CheckEmail()
			dim email:email=unescape(KS.S("email"))
			if email="" then
			 KS.Echo escape("err|������������䣡")
			elseif instr(email,"@")=0 or instr(email,".")=0 then
			 KS.Echo escape("err|�����������������")
			elseif ks.setting(28)=1 or conn.Execute("Select userid From KS_User where email='"&email&"'" ).eof Then
			 KS.Echo escape("ok|�������������ע��!")
			else
			 KS.Echo escape("err|�������Ѿ�����ʹ�ã�������ѡ��")
			end if
		End Sub
		Sub CheckCode()
		  dim code:code=unescape(KS.S("code"))
		  IF Trim(code)<>Trim(Session("Verifycode")) And KS.ChkCLng(KS.Setting(27))=1 then 
		   	 KS.Echo escape("err|��֤���������������룡")
		  Else
		   	 KS.Echo escape("ok|��֤�������룡")
		  End IF
		End Sub
End Class
%> 
