<% @LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="收益计划">
<p>
<%
Dim KSCls
Set KSCls = New User_Expansion
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%
Class User_Expansion
        Private KS
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		    Set KS=Nothing
		    Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		    IF Cbool(KSUser.UserLoginChecked)=False Then
			   Response.redirect KS.GetDomain&"User/Login/"
			   Exit Sub
			End If
			
			If KS.Setting(140)="1" Then
               Response.Write "【参与推广计划】<br/>" &vbcrlf
               Response.Write "将本站推荐给朋友:<br/>"&Replace(KS.Setting(142),"{$UID}",KSUser.UserName)&"<br/>" &vbcrlf
               Response.Write "奖励说明：成功推荐一个访问者,您就可以增加 <b>"&KS.Setting(141)&"</b> 个积分。赶快行动吧！<br/>" &vbcrlf
			End If
			If KS.Setting(143)="1" Then
			   Response.Write "<br/>" &vbcrlf
			   Response.Write "【会员注册推广】<br/>" &vbcrlf
               Response.Write "引导朋友注册:<br/>"&KS.GetDomain&"User/Reg/?UID="&KSUser.UserName&"<br/>" &vbcrlf
               Response.Write "奖励说明：成功推荐一个用户注册,您就可以增加 <b>"&KS.Setting(144)&"</b> 个积分。赶快行动吧！<br/>" &vbcrlf
			End If
			Response.Write "<br/>" &vbcrlf
			Response.Write "<a href=""Index.asp?" & KS.WapValue & """>我的地盘</a><br/>" &vbcrlf
			Response.Write "<a href=""" & KS.GetGoBackIndex & """>返回首页</a>" &vbcrlf
		End Sub
End Class
%>