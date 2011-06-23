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
<%
Dim KSCls
Set KSCls = New Template
KSCls.Kesion()
Set KSCls = Nothing

Class Template
        Private KS,KSRFObj
		Private ID
		Private Sub Class_Initialize()
		    If (Not Response.IsClientConnected) Then
			   Response.Clear
			   Response.End
		    End If
		    Set KS=New PublicCls
			Set KSRFObj = New Refresh
		End Sub
        Private Sub Class_Terminate()
		    Call CloseConn()
		    Set KS=Nothing
			Set KSRFObj=Nothing
		End Sub
		
		Public Sub Kesion()	
		    ID=KS.S("ID")
			Dim FileContent:FileContent = KS.C_T(ID,2)
			If Trim(FileContent) = "" Then
			   Call KS.ShowError("错误信息！","模板内容为空！")
			End If
			FileContent = KSRFObj.KSLabelReplaceAll(FileContent)
			If InStr(FileContent,"[KS_Charge]")<>0 Then
			   If Cbool(KSUser.UserLoginChecked)=False Then
			      Dim LoginContent
				  LoginContent="对不起，你还没有登录，至少要求本站的注册会员才可查看!<br/>"
				  LoginContent=LoginContent&"如果你还没有注册，请<a href=""../User/reg?../plus/Template.asp?ID=" & KS.S("ID") & """>点此注册</a>吧!<br/>"
				  LoginContent=LoginContent&"如果您已是本站注册会员，赶紧<a href=""User/Login/?../plus/Template.asp?ID=" & KS.S("ID") & """>点此登录</a>吧！<br/>"
				  Dim ChargeContent:ChargeContent=KS.CutFixContent(FileContent, "[KS_Charge]", "[/KS_Charge]", 0)
				  FileContent=Replace(FileContent,"[KS_Charge]" & ChargeContent &"[/KS_Charge]",LoginContent)
			   Else
			      FileContent=Replace(FileContent,"[KS_Charge]","")
				  FileContent=Replace(FileContent,"[/KS_Charge]","")
			   End If
			Else
			   FileContent=Replace(FileContent,"[KS_Charge]","")
			   FileContent=Replace(FileContent,"[/KS_Charge]","")
			End If
			FileContent = KS.GetEncodeConversion(FileContent)
			Response.Write FileContent
		End Sub
End Class
%>