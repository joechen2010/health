<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
'********************************
'* 程序功能：百变彩图下载
'* 演示地址: http://wap.kesion.com/
'********************************
%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<% Response.ContentType="text/vnd.wap.wml" %><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="百变彩图下载">
<p>
<%
Dim KSCls
Set KSCls = New PhotoDownLoad
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%
Class PhotoDownLoad
        Private KS,JpegUrl
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		    Call CloseConn()
		    Set KS=Nothing
		End Sub
        Public Sub Kesion()
		    JpegUrl=KS.S("JpegUrl")
			If JpegUrl="" Then JpegUrl=""&KS.GetDomain&"Images/nopic.gif"
			If Left(JpegUrl,1)="/" Then JpegUrl=Right(JpegUrl,Len(JpegUrl)-1)
			If Lcase(Left(JpegUrl,4))<>"http" Then JpegUrl=KS.Setting(2) & KS.Setting(3) & JpegUrl
			Response.Write "=百变彩图=<br/>"
			Response.Write "<img src=""../JpegMore.asp?JpegSize=128x0&amp;JpegUrl="&JpegUrl&""" alt=""""/><br/>"
			If KS.BusinessVersion = 1 Then
			   Response.Write "<a href=""../JpegDown.asp?JpegUrl="&JpegUrl&"&amp;"&KS.WapValue&""">下载</a> "
			   Response.Write "<a href=""../JpegShow.asp?JpegUrl="&JpegUrl&"&amp;"&KS.WapValue&""">缩放</a> "
			   Response.Write "<a href=""../JpegEdit.asp?JpegUrl="&JpegUrl&"&amp;"&KS.WapValue&""">高级</a><br/>"	  
			Else
			   Response.Write "<a href="""&JpegUrl&""">立即下载</a><br/>"
			End if
			Response.Write "=免责声明=<br/>"
			Response.Write "你正下载的图片文件均属网友上传或网上转载，本文件只用于大家学习和交流，由于文件的可用性或版权产生纠纷，本网概不负责。如果你发现下载文件出现问题或侵害了你的权益，请第一时间通知我们客服及时处理。谢谢！ E-mail:"&KS.Setting(11)&"<br/>"
			Response.Write "---------<br/>"
			Response.Write "<anchor><prev/>还回上级</anchor><br/>"
			Response.Write "<a href="""&KS.GetGoBackIndex&""">返回首页</a><br/>"
	    End Sub
End Class
%>
