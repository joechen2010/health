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
				   FCls.RefreshType = "MoreSpace" '设置刷新类型，以便取得当前位置导航等
				   Application(KS.SiteSN & "RefreshFolderID") = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
				   If Trim(FileContent) = "" Then FileContent = "空间副模板不存在!"
				   FileContent=Replace(FileContent,"{$ShowMain}",GetSpaceList())
				   FileContent=KSRFObj.KSLabelReplaceAll(FileContent)
		   Response.Write FileContent  
		End Sub
		Function GetSpaceList()
		  GetSpaceList= "<script src=""../ks_inc/jquery.js"" language=""javascript""></script>"  & vbcrlf
		  GetSpaceList=GetSpaceList & "<script src=""../ks_inc/kesion.box.js"" language=""javascript""></script>"  & vbcrlf
		  GetSpaceList=GetSpaceList & "<script src=""js/ks.space.js"" language=""javascript""></script>"  & vbcrlf
		  GetSpaceList=GetSpaceList & "<script src=""js/ks.space.page.js"" language=""javascript""></script>"  & vbcrlf
		

		  GetSpaceList=GetSpaceList & "<div id=""spacemain""><script language=""javascript"" defer>SpacePage(1,'space&classid=" & KS.S("ClassID") & "&recommend=" & KS.S("Recommend") & "')</script></div><div id=""kspage""></div>"  & vbcrlf

		End Function
End Class
%>
