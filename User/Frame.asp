<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Frame
KSCls.Kesion()
Set KSCls = Nothing

Class Frame
        Private KS,KSUser
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "<script>window.close();</script>"
		  Exit Sub
		End If
		Dim Url, FileName, PageTitle, ChannelID, Action
		Dim QueryParam
		Url = Request.QueryString("Url")
		Action = Request.QueryString("Action")
		PageTitle = Request.QueryString("PageTitle")
		ChannelID=KS.ChkClng(KS.S("ChannelID"))
		QueryParam ="?ChannelID=" & ChannelID
		If Action <> "" Then QueryParam = QueryParam & "&Action=" & Action
		
		FileName = Url & "?" & KS.QueryParam("url")
		
		
		Response.Write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">"
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		Response.Write "<title>" & PageTitle & "</title>"
		Response.Write "</head>"
		Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"" scrolling=no>"
		Response.Write "<Iframe src=" & FileName & " style=""width:100%;height:100%;"" frameborder=0 scrolling=""yes""></Iframe>"
		Response.Write "</body>"
		Response.Write "</html>"
		End Sub
End Class
%>
 
