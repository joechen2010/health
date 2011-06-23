<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
'********************************
'* 程序功能：广告
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
<card id="card" title="正在进入...">
<p>
<%
Dim KS:Set KS=New PublicCls
Dim ID:ID=KS.ChkClng(KS.G("ID"))
Dim RS:set RS=Server.CreateObject("Adodb.Recordset")
RS.Open "select ID,Click,Lasttime,Url from KS_Advertise where ID="&ID,Conn,1,3
If NOT RS.EOF Then
   Rs("Click")=Rs("Click")+1
   Rs("LastTime")=Now()
   Rs.Update
   'Dim RS1:set RS1=Server.CreateObject("adODB.Recordset")
   'RS1.Open "select * from KS_Adiplist",Conn,1,3
   'RS1.AddNew
   'RS1("AdID")=RS("ID")
   'RS1("Time")=now()
   'RS1("ip")=KS.GetIP
   'RS1("Class")=2
   'RS1.Update
   'RS1.Close:set RS1=Nothing
   Response.Redirect ""&RS("url")&""
Else
   Response.Redirect KS.GetGoBackIndex
End If
RS.Close:set RS=Nothing
Call CloseConn
Set KSUser=Nothing
Set KS=Nothing
%>
</p>
</card>
</wml>
