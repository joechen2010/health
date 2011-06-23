<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.SpaceCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KS,KSBCls
Set KS=New PublicCls
Set KSBCls=New BlogCls
dim TemplateID
TemplateID=KS.ChkClng(KS.S("TemplateID"))
KS.Echo KSBCls.GetTemplatePath(TemplateID,"TemplateMain")
Set KS=Nothing
Set KSBCls=Nothing
call closeconn()
%>