<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card title="操作提示">
<p>
<%
Dim KS:Set KS=New PublicCls
Dim Message,MessageID
Message = Request("Message")
Action = KS.S("Action")
If Action<>"" Then
	Select Case Action
		Case "DelSql"
		   Message="系统警告！您提交的数据有恶意字符“" & Message &"”，您的数据已经被记录。IP：" & KS.GetIP & " 日期:"&Now()&""
		Case Else
		   Message="非法参数！"
	End Select
End If
Response.Write Message&"<br/>"
Response.Write "<br/><anchor><prev/>还回上级</anchor><br/>"

Call CloseConn
Set KS=Nothing
Set KSUser=Nothing
%>
</p>
</card>
</wml>