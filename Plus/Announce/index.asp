<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Announce
KSCls.Kesion()
Set KSCls = Nothing

Class Announce
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		Dim AnnounceID, FileContent
		Dim RefreshRS, KMRFObj
		Set KMRFObj = New Refresh
		
		 AnnounceID = KS.ChkClng(request.QueryString)
If KS.Setting(112)="" Then Response.Write "���ȵ�""������Ϣ����->ģ���""����ģ��󶨲���!":response.end
		   FileContent = KMRFObj.LoadTemplate(KS.Setting(112))
		   If Trim(FileContent) = "" Then FileContent = "ģ�岻����!"
		   Set RefreshRS = Server.CreateObject("Adodb.Recordset")
		   RefreshRS.Open "Select Title,Author,AddDate,Content From KS_Announce Where ID=" & AnnounceID, Conn, 1, 1
		   If Not RefreshRS.EOF Then
			FileContent = ReplaceAnnounceContent(RefreshRS, FileContent)     '�滻�������ݱ�ǩΪ����
		   Else
			FileContent = "�������ݴ���!"
		   End If
		   RefreshRS.Close:Set RefreshRS = Nothing
		   FileContent=KMRFObj.KSLabelReplaceAll(FileContent)
		   Set KMRFObj = Nothing
		   Response.Write FileContent   '�����������ҳ
		End Sub
		'*********************************************************************************************************
		'��������ReplaceAnnounceContent
		'��  �ã��滻��������ҳ��ǩΪ����
		'��  ����FileContent���滻������
		'*********************************************************************************************************
		Function ReplaceAnnounceContent(RefreshRS, FileContent)
			   If InStr(FileContent, "{$GetAnnounceTitle}") <> 0 Then
				  FileContent = Replace(FileContent, "{$GetAnnounceTitle}", RefreshRS(0))
			   End If
			   If InStr(FileContent, "{$GetAnnounceAuthor}") <> 0 Then
				  FileContent = Replace(FileContent, "{$GetAnnounceAuthor}", RefreshRS(1))
			   End If
			   If InStr(FileContent, "{$GetAnnounceDate}") <> 0 Then
				  FileContent = Replace(FileContent, "{$GetAnnounceDate}", RefreshRS(2))
			   End If
			   If InStr(FileContent, "{$GetAnnounceContent}") <> 0 Then
				  FileContent = Replace(FileContent, "{$GetAnnounceContent}", RefreshRS(3))
			   End If
			   ReplaceAnnounceContent = FileContent
		End Function

End Class
%>

 
