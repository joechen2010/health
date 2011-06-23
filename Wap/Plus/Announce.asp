<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
'********************************
'* 程序功能：栏目列表
'* 演示地址: http://wap.kesion.com/
'********************************
%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="最新公告">
<%
Dim KS:Set KS=New PublicCls

Dim AnnounceID,RefreshRS
AnnounceID = KS.ChkClng(KS.S("AnnounceID"))

Set RefreshRS = Server.CreateObject("Adodb.Recordset")
RefreshRS.Open "Select Title,Author,AddDate,Content From KS_Announce Where ID=" & AnnounceID, Conn, 1, 1
If Not RefreshRS.EOF Then
   Response.Write "<p align=""center"">"&RefreshRS(0)&"</p>"
   Response.Write "<p align=""left"">"&KS.UBBToHTML(KS.LoseHtml(KS.HTMLToUBB(KS.GetEncodeConversion(RefreshRS(3)))))&"</p>"
   Response.Write "<p align=""right"">"&RefreshRS(1)&"<br/>"
   Response.Write ""&RefreshRS(2)&"</p>"
Else
   Response.Write "<p align=""center"">参数传递错误!</p>"
End If
RefreshRS.Close
Set RefreshRS = Nothing
%> 

<p align="center"> 
---------<br/>
<%
IF Cbool(KSUser.UserLoginChecked)=True Then Response.Write "<a href=""../User/Index.asp?" & KS.WapValue & """>我的地盘</a>"
Response.write " <a href=""" & KS.GetGoBackIndex & """>返回首页</a><br/>"
Call CloseConn
Set KSUser=Nothing
Set KS=Nothing
%>
</p>

</card>
</wml> 
