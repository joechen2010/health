<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit
 Response.Buffer=true
%>
<!--#include file="../KS_Cls/Kesion.EditorCls.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSUser:Set KSUser=New UserCls
IF Cbool(KSUser.UserLoginChecked)=false And Request.QueryString("ChannelID")<>"10000" And Request.QueryString("ChannelID")<>"9999" Then
	  Response.Write "<script>location.href='/';</script>"
	  Response.End()
End If
Set KSUser=Nothing
Dim EditorClass:Set EditorClass = New KesionEditor
EditorClass.Kesion(0)
Set EditorClass = Nothing
%> 
