<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
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

Dim KSCls
Set KSCls = New Spacemore
KSCls.Kesion()
Set KSCls = Nothing

Class Spacemore
        Private KS, KSRFObj
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSRFObj = New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		      Dim FileContent
				   FileContent = KSRFObj.LoadTemplate(KS.SSetting(8))
				   FCls.RefreshType = "MoreGroup" '����ˢ�����ͣ��Ա�ȡ�õ�ǰλ�õ�����
				   Application(KS.SiteSN & "RefreshFolderID") = "0" '���õ�ǰˢ��Ŀ¼ID Ϊ"0" ��ȡ��ͨ�ñ�ǩ
				   If Trim(FileContent) = "" Then FileContent = "�ռ丱ģ�岻����!"
				   FileContent=Replace(FileContent,"{$ShowMain}",GetGroupList())
				   FileContent=KSRFObj.KSLabelReplaceAll(FileContent)
		   Response.Write FileContent  
		End Sub
		Function GetGroupList()
		  GetGroupList= "<script src=""../ks_inc/jquery.js"" language=""javascript""></script>"  & vbcrlf
		  GetGroupList=GetGroupList & "<script src=""../ks_inc/kesion.box.js"" language=""javascript""></script>"  & vbcrlf
		  GetGroupList=GetGroupList & "<script src=""js/ks.space.js"" language=""javascript""></script>"  & vbcrlf
		  GetGroupList=GetGroupList & "<script src=""js/ks.space.page.js"" language=""javascript""></script>"  & vbcrlf
		 

		  GetGroupList=GetGroupList & "<div id=""spacemain""><script language=""javascript"" defer>SpacePage(1,'group&classid=" & KS.S("ClassID") & "&recommend=" & ks.s("recommend") & "')</script></div><div id=""kspage""></div>"  & vbcrlf

		End Function
End Class
%>
