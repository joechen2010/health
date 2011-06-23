<%@ Language="VBSCRIPT" codepage="936" %>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="gb2312"

Dim KS:Set KS=New PublicCls
Dim KSCls:Set KSCls=New RefreshFunction

Dim LabelID:LabelID=KS.R(KS.S("LabelID"))   '标签ID
Dim InfoID:InfoID=KS.R(KS.S("InfoID"))      '信息ID
FCls.RefreshInfoID=InfoID      '设置信息ID，以取得相关链接
IF KS.S("labtype")="-1" Then
FCls.RefreshFolderID=KS.S("ClassID")
End IF
FCls.ChannelID=KS.ChkCLng(KS.S("Channelid"))

IF (KS.IsNul(Request.ServerVariables("HTTP_REFERER"))) Then KS.Die "error!"

If LabelID="" Then Response.Write "非法调用！":Response.End
If KS.S("Action")="SQL" Then
	Dim KSRCls:Set KSRCls=New DIYCls
	Dim LabelName:LabelName=replace(replace("{"&split(Request.QueryString("LabelID"),"ksr")(0)&")}","ksl","("),"ksu","_")
	KS.Echo KSRCls.ReplaceDIYFunctionLabel(LabelName,"ajax")
	Set KSRCls=Nothing
Else
     Dim L_P
     Dim RCls:Set RCls=New Refresh
	 Call RCls.LoadLabelToCache()    '加载标签
	 Set RCls=Nothing
     L_P=Replace(Application(KS.SiteSN&"_labellist").documentElement.selectSingleNode("labellist[@labelid='" & LabelID & "']").text,LabelID,"ajax")
	 If L_P="" Then Response.End
	 KS.Echo KSCls.GetLabel(l_p)
End If
Set KSCls=Nothing
Set KS=Nothing
CloseConn
%>
