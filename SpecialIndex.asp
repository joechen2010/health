<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="Conn.asp"-->
<!--#include file="KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************

Dim KSCls
Set KSCls = New SpecialIndex
KSCls.Kesion()
Set KSCls = Nothing

Class SpecialIndex
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
		  Dim FileContent,FsoIndex:FsoIndex=KS.Setting(5)
				   FileContent = KSRFObj.LoadTemplate(KS.Setting(111))
				   FCls.RefreshType = "SpecialIndex"  '����ˢ�����ͣ��Ա�ȡ�õ�ǰλ�õ�����
				   FCls.RefreshFolderID = "0"         '���õ�ǰˢ��Ŀ¼ID Ϊ"0" ��ȡ��ͨ�ñ�ǩ
				   FCls.CurrSpecialID="" '�����ǰר��ID
				   If Trim(FileContent) = "" Then FileContent = "��ҳģ�岻����!"
				    FileContent=KSRFObj.KSLabelReplaceAll(FileContent)
		   Response.Write FileContent  
		End Sub
End Class
%>
