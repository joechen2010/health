<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New RefreshIndex
KSCls.Kesion()
Set KSCls = Nothing

Class RefreshIndex
        Private KS,KSRObj
		Private SaveFilePath
		Private FileContent
        Private ReturnInfo
		Private ErrFlag
		Private Domain
		Private IndexFile
		Private StartRefreshTime
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSRObj=New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KSRObj=Nothing
		 Set KS=Nothing
		End Sub
		Sub Kesion()
		  With Response
		  If Not KS.ReturnPowerResult(0, "KMTL20000") Then          '����Ƿ��з���վ����ҳ��Ȩ��
		   .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back();';</script>")
		   Call KS.ReturnErr(1, "")
		   Exit Sub
		   .End
		  End If
		 If Split(KS.Setting(5),".")(1)="asp" Then Call KS.AlertHistory("��Ѵϵͳ��������\n\n1��վ����ҳû���������ɾ�̬HTML����\n\n2���뵽ϵͳ����->������Ϣ�����������ɾ�̬Html����",-1):Exit Sub
		   StartRefreshTime = Timer()
		   FCls.RefreshType = "INDEX" '����ˢ�����ͣ��Ա�ȡ�õ�ǰλ�õ�����
		   FCls.RefreshFolderID = "0" '���õ�ǰˢ��Ŀ¼ID Ϊ"0" ��ȡ��ͨ�ñ�ǩ
		   Domain = KS.GetDomain
		   IndexFile = KS.Setting(5)
			SaveFilePath = KS.Setting(3) & IndexFile
			FileContent = KSRObj.LoadTemplate(KS.Setting(110))
			If FileContent = "" Then
			  ReturnInfo = "���ݿ����Ҳ�����ҳģ��"
			  ErrFlag = True
			  Call Main
			  .End
			Else
			 ' On Error Resume Next
			  FileContent = KSRObj.KSLabelReplaceAll(FileContent) '�滻������ǩ
			  FileContent = KSRObj.ReplaceRA(FileContent, "")
			  If Err Then
			   ReturnInfo = Err.Description
			   ErrFlag = True
				 Err.Clear
				Call Main
				.End
			  End If
			  Call KSRObj.FSOSaveFile(FileContent, SaveFilePath)
			  If Err Then
				ReturnInfo = Err.Description
				ErrFlag = True
				 Err.Clear
				Call Main
				.End
			  End If
			  ReturnInfo = "��վ��ҳ�����ɹ�"
			  ErrFlag = False
			  Call Main
			  if request("f")="task" then
			   KS.Echo "<script>setTimeout('window.close();',3000);</script>"
			  end if
			End If
		 End With
		End Sub
		Sub Main()
			With Response	  
		        .Write "<html>"
				.Write ("<head><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>")
				.Write ("<title>��վ��ҳ��������</title></head>")
				.Write ("<link rel='stylesheet' href='Admin_Style.CSS'>")
				If KS.G("RefreshFlag")<>"Info" Then
				.Write ("<body topmargin='0' leftmargin='0' oncontextmenu='return false;'>")
				Else
		        .Write ("<body oncontextmenu=""return false;"" scroll=no bgcolor='transparent'>")
				End If
				If KS.G("RefreshFlag")<>"Info" Then
				.Write ("<table width='100%' border='0' cellpadding='0' cellspacing='0'>")
				.Write ("  <tr>")
				.Write ("    <td height='25' class='sort'>")
				.Write ("      <div align='center'><strong>������վ��ҳ</strong></div></td>")
				.Write ("</tr>")
				.Write ("</table>")
				.Write ("<table width='100%' border='0' cellspacing='0' cellpadding='0'><tr><td>&nbsp;</td></tr>")
				.Write ("  <tr><td>&nbsp;</td></tr>")
				.Write ("  <tr>")
				.Write ("    <td height='50'><div align='center'><br>")
					   .Write ReturnInfo & "���ܷ�ʱ:<font color=red>" & Left((Timer() - StartRefreshTime), 4) & "</font>��</div></td>"
				.Write ("</tr><tr><td><div align='center'>")
				.Write ("        <table width='100%' border='0' cellspacing='0' cellpadding='0'>")
				.Write ("          <tr><td width='50%' height='25'><div align='right'>�������:</div></td>")
				.Write ("            <td width='50%' height='25'>")
						   
						   If ErrFlag = False Then
							.Write ("�ɹ�")
							Else
							.Write ("ʧ��")
							End If
				.Write ("            </td></tr><tr><td height='25'> <div align='right'>��ǰʱ��:</div></td><td height='25'> " & Now & "</td></tr>")
						 
						 If ErrFlag = False Then
				.Write ("          <tr><td height='25'><div align='right'>������:</div></td>")
				.Write ("           <td height='25'><font color='#FF0000'> <a href='" & Domain & IndexFile & "' target='_blank'>�����ҳ</a></font>")
				.Write ("            </td></tr>")
						  End If
				.Write ("        </table></div></td></tr></table>")
			 Else
				.Write ("<table width='67%' border='0' cellpadding='0' cellspacing='0'>")
				.Write ("  <tr>")
				.Write ("    <td height='25'>")
				.Write ("      <div ><li><strong>" & ReturnInfo & "</strong><font color='#FF0000'> <a href='" & Domain & IndexFile & "' target='_blank'>" & Domain & IndexFile &"</a></font></li></div></td>")
				.Write ("</tr>")
				.Write ("</table>")
			 End If
			 .Write    ("</body></html>")
		End With
		End Sub
End Class
%> 
