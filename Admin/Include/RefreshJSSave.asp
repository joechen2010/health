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
Set KSCls = New RefreshJSSave
KSCls.Kesion()
Set KSCls = Nothing

Class RefreshJSSave
        Private KS
		Private KSRObj
		Private ReturnInfo
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSRObj=New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSRObj=Nothing
		End Sub
		Sub Kesion()
			Dim RefreshFlag, RefreshSql, FolderID, NewsNo, RefreshRS, NewsTotalNum, StartRefreshTime
			'ˢ��ʱ��
			StartRefreshTime = KS.G("StartRefreshTime")
			If StartRefreshTime = "" Then StartRefreshTime = Timer()
			
			 RefreshFlag = KS.G("RefreshFlag")
			  RefreshSql = request("RefreshSql")
			  NewsNo = KS.G("NewsNo")
			 If NewsNo = "" Then NewsNo = 0
			 If RefreshSql = "" Then
				If RefreshFlag = "Folder" Then
				  FolderID =Replace(Replace(KS.G("FolderID")," ",""),",","','")
				  RefreshSql = "Select JSName From KS_JSFile Where JSTYPE=0 And FolderID IN ('" & FolderID & "')"
				ElseIf RefreshFlag = "All" Then
				  RefreshSql = "Select JSName From KS_JSFile Where JSTYPE=0"
			   Else
				  RefreshSql = ""
			   End If
			End If
			If RefreshSql <> "" Then
				Set RefreshRS = Server.CreateObject("ADODB.RecordSet")
				RefreshRS.Open RefreshSql, Conn, 1, 1
				NewsTotalNum = RefreshRS.RecordCount
				If RefreshRS.EOF Then
					ReturnInfo = "û��Ҫˢ�µ�ϵͳJS&nbsp;&nbsp;<br><input name=""button1"" type=""button"" onclick=""javascript:location='RefreshJS.asp';""  class='button' value="" �� �� "">"
					Set RefreshRS = Nothing
					Call Main
				Else
					RefreshRS.Move NewsNo
					If Not RefreshRS.EOF Then
					   Call KSRObj.RefreshJS(RefreshRS(0))  '������Ŀˢ�º���
						NewsNo = NewsNo + 1
						Response.Write ("<meta http-equiv=""refresh"" content=""0;url='RefreshJSSave.asp?StartRefreshTime=" & Server.URLEncode(StartRefreshTime) & "&NewsNo=" & NewsNo & "&RefreshSql=" & Server.URLEncode(RefreshSql) & "&RefreshFlag=" & RefreshFlag & "'"">")
						ReturnInfo = "�ܹ���Ҫˢ�� <font color=red><b>" & NewsTotalNum & "</b></font> ��ϵͳJS<br><br>����ˢ�µ� <font color=red><b>" & NewsNo - 1 & "</b></font> ��ϵͳJS,���Ժ�... <font color=red><b>�ڴ˹���������ˢ�´�ҳ�棡����</b></font><br>"
					Else
						ReturnInfo = "ˢ��ϵͳJS�������ܷ�ʱ:<font color=red>" & Left((Timer() - StartRefreshTime), 4) & "</font> ��<br><br>�ܹ�ˢ���� <font color=red><b>" & NewsTotalNum & "</b></font> ��ϵͳJS <br><br><input name=""button1"" type=""button"" onclick=""javascript:location='RefreshJS.ASP';""  class='button' value="" �� �� "">"
					End If
					Set RefreshRS = Nothing
					Call Main
				End If
				Set RefreshRS = Nothing
			Else
				ReturnInfo = "�Բ�����û��ѡ��Ҫ������ϵͳJSĿ¼&nbsp;&nbsp;<font color=""red""><a href=""RefreshJS.ASP"">����</a></font>"
				Call Main
			End If
			
			End Sub
			
			Sub Main()
			 Response.Write ("<html>")
			 Response.Write ("<head>")
			 Response.Write ("<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">")
			 Response.Write ("<title>ϵͳ��Ϣ</title>")
			 Response.Write ("</head>")
			 Response.Write ("<link rel=""stylesheet"" href=""Admin_Style.css"">")
			 Response.Write ("<body oncontextmenu=""return false;"">")
			 Response.Write ("<table width=""80%"" height=""50%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">")
			 Response.Write (" <tr>")
			 Response.Write ("   <td height=""50"">")
			 Response.Write ("     <div align=""center""> ")
			 Response.Write (ReturnInfo)
			 Response.Write ("       </div></td>")
			 Response.Write ("   </tr>")
			 Response.Write ("</table>")
			 Response.Write ("</body>")
			 Response.Write ("</html>")
			End Sub

End Class
%> 
