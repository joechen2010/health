<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->

<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New UserList
KSCls.Kesion()
Set KSCls = Nothing

Class UserList
        Private KS,KSUser,LoginTF,TemplateFile
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser=New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		 Dim FileContent,MainUrl,RequestItem
		 Dim KMRFObj,ParaList
		 FCls.RefreshType = "Member"   '设置当前位置为会员中心
		 Set KMRFObj = New Refresh
         Dim LoginTF:LoginTF=Cbool(KSUser.UserLoginChecked)
		 TemplateFile=KS.Setting(116)
		 If LoginTF=True Then
		  TemplateFile=KS.U_G(KSUser.GroupID,"templatefile")
		 End If
		 If trim(TemplateFile)="" Then TemplateFile=KS.Setting(116)
         If trim(TemplateFile)="" Then Response.Write "请先到""基本信息设置->模板绑定""进行模板绑定操作!":response.end
		   FileContent = KMRFObj.LoadTemplate(TemplateFile)
		   If Trim(FileContent) = "" Then FileContent = "模板不存在!"
			ParaList = ""
			For Each RequestItem In Request.QueryString
				If Ucase(RequestItem) <> "COMEURL" Then
					If ParaList = "" Then
						ParaList = RequestItem & "=" & Server.URLEncode(Request.QueryString(RequestItem))
					Else
						ParaList = ParaList & "&" & RequestItem & "=" & Server.URLEncode(Request.QueryString(RequestItem))
					End If
				End If
			Next
			For Each RequestItem In Request.Form
				If Ucase(RequestItem) <> "COMEURL" Then
					If ParaList = "" Then
						ParaList = RequestItem & "=" & Server.URLEncode(Request.Form(RequestItem))
					Else
						ParaList = ParaList & "&" & RequestItem & "=" & Server.URLEncode(Request.Form(RequestItem))
					End If
				End If
			Next
			MainUrl=KS.S("ComeUrl")
			If MainUrl <> "" Then
				MainUrl = MainUrl & "?" & ParaList
			Else
				MainUrl="User_main.asp"
			End If
			FileContent = Replace(FileContent,"{$MainUrl}",MainUrl)
		  
			FileContent = KMRFObj.KSLabelReplaceAll(FileContent)
		   Set KMRFObj = Nothing
		   Response.Write FileContent   
	End Sub

End Class
%> 
