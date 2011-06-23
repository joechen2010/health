<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Admin_Frame
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Frame
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()
		Dim ParaList,RequestItem,FileName,Url,PageTitle
		Url = KS.G("Url")
		PageTitle=KS.G("PageTitle")
		ParaList = ""
		For Each RequestItem In Request.QueryString
			If Ucase(RequestItem) <> "URL" And Ucase(RequestItem) <> "PAGETITLE" Then
				If ParaList = "" Then
					ParaList = RequestItem & "=" & Server.URLEncode(Request.QueryString(RequestItem))
				Else
					ParaList = ParaList & "&" & RequestItem & "=" & Server.URLEncode(Request.QueryString(RequestItem))
				End If
			End If
		Next
		If Url <> "" Then
			FileName = Url & "?" & ParaList
		Else
			Response.Write ("<script language=""JavaScript"">alert('文件不存在');window.close();</script>")
			Exit Sub
		End If
		Response.Write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">"
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		Response.Write "<title>" & PageTitle & "</title>"
		Response.Write "</head>"
		Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"" scrolling=no>"
		Response.Write "<Iframe src=""" & FileName & """ style=""width:100%;height:100%;"" frameborder=0 scrolling=""no""></Iframe>"
		Response.Write "</body>"
		Response.Write "</html>"
		End Sub
End Class
%>
 
