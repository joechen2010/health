<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
'********************************
'* 程序功能：更多介绍
'* 演示地址: http://wap.kesion.com/
'********************************
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
Set KSCls = New MoreContent
KSCls.Kesion()
Set KSCls = Nothing


Class MoreContent
        Private KS
		Private Sub Class_Initialize()
		    If (Not Response.IsClientConnected) Then
			   Response.Clear
			   Response.End
		    End If
		    Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		    Call CloseConn()
		    Set KS=Nothing
		End Sub
		Public Sub Kesion()
		    Dim RS,SqlStr,ID,ChannelID,Title,Content
			ID=KS.ChkClng(KS.R(KS.S("ID")))
			If ID=0 Then Exit Sub
			ChannelID=KS.ChkClng(KS.S("ChannelID"))
		    Set RS=Server.CreateObject("Adodb.Recordset")
			Select Case KS.C_S(ChannelID,6)
				Case 3
				SqlStr= "Select * from "&KS.C_S(ChannelID,2)&" Where ID="&ID
				Case 5
				SqlStr= "Select * from "&KS.C_S(ChannelID,2)&"  Where verific=1 And ID="&ID
				Case 8
				SqlStr= "Select * from KS_GQ Where ID="&ID
			End Select
			RS.Open SqlStr,Conn,1,3
			IF RS.Eof And RS.Bof Then
			   RS.Close:Set RS=Nothing
			   Select Case KS.C_S(ChannelID,6)
				   Case 3
				   Call KS.ShowError("系统提示！","系统提示！<br/>您要查看的" & KS.C_S(ChannelID,3) & "已删除。或是您非法传递注入参数！")
				   Case 5
				   Call KS.ShowError("系统提示！","系统提示！<br/>您要查看的" & KS.C_S(ChannelID,3) & "已删除或是未通过暂停销售！")
			   End Select
			Else
			   Select Case KS.C_S(ChannelID,6)
			       Case 3
				      Title=RS("Title")
					  Content=RS("DownContent")
				   Case 5
				      Title=RS("Title")
					  Content=KS.UbbToHtml(KS.LoseHtml(KS.HtmlToUbb(KS.GetEncodeConversion(RS("ProIntro")))))
				   Case 8
				      Title=RS("Title")
					  Content=KS.UbbToHtml(KS.LoseHtml(KS.HtmlToUbb(KS.GetEncodeConversion(RS("GQContent")))))
			   End Select
		       Response.Write "<wml>" &vbcrlf
			   Response.Write "<head>" &vbcrlf
			   Response.Write "<meta http-equiv=""Cache-Control"" content=""no-Cache""/>" &vbcrlf
			   Response.Write "<meta http-equiv=""Cache-Control"" content=""max-age=0""/>" &vbcrlf
			   Response.Write "</head>" &vbcrlf
			   Response.Write "<card id=""main"" title="""&Title&""">" &vbcrlf
			   Response.Write "<p align=""left"">" &vbcrlf
			   Response.Write KS.ContentPagination(Content,200,"?ID="&ID&"&amp;ChannelID="&ChannelID&"&amp;"&KS.WapValue&"",True,True)&"<br/>" &vbcrlf
			   Response.Write "---------<br/>" &vbcrlf
			   Response.Write "<a href=""../Show.asp?ID="&ID&"&amp;ChannelID="&ChannelID&"&amp;"&KS.WapValue&""">返回"&KS.C_S(ChannelID,3)&"页</a><br/>" &vbcrlf
			   Response.Write "<a href="""&KS.GetGoBackIndex&""">返回网站首页</a>" &vbcrlf
			   Response.Write "</p>" &vbcrlf
			   Response.Write "</card>" &vbcrlf
			   Response.Write "</wml>"
			End If
			RS.Close:Set RS=Nothing
		End Sub
End Class
%>