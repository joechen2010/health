<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit
 Response.Buffer=true
%>
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.EditorCls.asp"-->
<!--#include File="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim EditorClass
Set EditorClass = New KesionEditor
EditorClass.Kesion(1)
Set EditorClass = Nothing
%> 
