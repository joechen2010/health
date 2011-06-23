<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="include/session.asp"-->
<%
Dim KSCls
Set KSCls= New AdminUploadFileCls
KSCls.Kesion()
Set KSCls=Nothing

Class AdminUploadFileCls
	Private KS
	Private ChannelDir, fullPath, FilePath, UploadDir, ThisDir
	Private Action, rsChannel
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		End Sub

	Public Sub Kesion()
		Action = LCase(Request("action"))
       If Not KS.ReturnPowerResult(0, "KMST10018") Then                  'Ȩ�޼��
				Call KS.ReturnErr(1, "")   
				Response.End()
		End iF
				
		ChannelDir = KS.Setting(3)& KS.Setting(91)
		If Trim(Request("UploadDir")) <> "" Then
			UploadDir = Trim(Request("UploadDir")) & "/"
		End If
		If Trim(Request("ThisDir")) <> "" Then
			ThisDir = Trim(Request("ThisDir")) & "/"
		End If
		ThisDir = Replace(ThisDir, "\", "/")


		if (left(UploadDir,1)="/") Then UploadDir=Right(UploadDir,len(UploadDir)-1)
		FilePath = Replace(ChannelDir & UploadDir, "\", "/")

		fullPath = Server.MapPath(FilePath)

		Select Case Trim(Action)
		Case "clear"
			Call ClearUploadFile
		Case "delete"
			Call DelUselessFile
		Case "del"
			Call DelFile
		Case "delalldirfile"
			Call DelAllDirFile
		Case "delthisallfile"
			Call DelThisAllFile
		Case "delemptyfolder"
			Call DelEmptyFolder
		Case Else
			Call ShowUploadMain
		End Select
	End Sub

	
	'=================================================
	'��������ShowChildFolder
	'��  �ã���ʾ��Ŀ¼�˵�
	'=================================================
	Private Sub ShowChildFolder()
		Dim fso, fsoFile, DirFolder
		Dim strFolderPath
		On Error Resume Next
		strFolderPath = ChannelDir & UploadDir
		strFolderPath = Server.MapPath(strFolderPath)
		Set fso = KS.InitialObject(KS.Setting(99))
		If fso.FolderExists(strFolderPath) Then
			Set fsoFile = fso.GetFolder(strFolderPath)
			For Each DirFolder In fsoFile.SubFolders
				Response.Write "<a href=""?UploadDir=" &Request("UploadDir") & "/"& DirFolder.Name& "&ThisDir=" & DirFolder.Name & """><img src=""images/folder/folderclosed.gif"" width=20 height=20 border=0 alt=""�޸�ʱ�䣺" & DirFolder.DateLastModified & """ align=absMiddle> "
				If Replace(ThisDir, "/", "") = DirFolder.Name Then
					Response.Write "<font color=red>" & DirFolder.Name & "</font>"
				Else
					Response.Write DirFolder.Name
				End If
				Response.Write "</a> &nbsp;&nbsp;" & vbNewLine
			Next
		Else
			Response.Write "û���ҵ��ļ��У�"
		End If
		Set fsoFile = Nothing: Set fso = Nothing
	End Sub

	'=================================================
	'��������showpage
	'��  �ã���ҳ
	'=================================================
	Private Function showpage(ByVal CurrentPage, ByVal TotalNumber, ByVal maxperpage, ByVal TotleSize)
		Dim n
		Dim strTemp
		
		If (TotalNumber Mod maxperpage) = 0 Then
			n = TotalNumber \ maxperpage
		Else
			n = TotalNumber \ maxperpage + 1
		End If
		strTemp = "<table align='center'><form method='Post' action='?UploadDir=" & Request("UploadDir") & "&ThisDir=" & Request("ThisDir") & "'><tr><td>" & vbNewLine
		strTemp = strTemp & "�� <b>" & TotalNumber & "</b> ���ļ���ռ�� <b>" & TotleSize & "</b>&nbsp;&nbsp;"
		'sfilename = JoinChar(sfilename)
		If CurrentPage < 2 Then
			strTemp = strTemp & "��ҳ ��һҳ&nbsp;"
		Else
			strTemp = strTemp & "<a href='?page=1&UploadDir=" & Request("UploadDir") & "&ThisDir=" & Request("ThisDir") & "'>��ҳ</a>&nbsp;"
			strTemp = strTemp & "<a href='?page=" & (CurrentPage - 1) & "&UploadDir=" & Request("UploadDir") & "&ThisDir=" & Request("ThisDir") & "'>��һҳ</a>&nbsp;"
		End If

		If n - CurrentPage < 1 Then
			strTemp = strTemp & "��һҳ βҳ"
		Else
			strTemp = strTemp & "<a href='?page=" & (CurrentPage + 1) & "&UploadDir=" & Request("UploadDir") & "&ThisDir=" & Request("ThisDir") & "'>��һҳ</a>&nbsp;"
			strTemp = strTemp & "<a href='?page=" & n & "&UploadDir=" & Request("UploadDir") & "&ThisDir=" & Request("ThisDir") & "'>βҳ</a>"
		End If
		strTemp = strTemp & "&nbsp;ҳ�Σ�<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>ҳ "
		strTemp = strTemp & "&nbsp;ת����"
		strTemp = strTemp & "<input name=page size=3 value='" & CurrentPage & "'> <input type=submit name=Submit value='ת��' class=Button>"
		strTemp = strTemp & "</select>"
		strTemp = strTemp & "</td>"
		strTemp = strTemp & "<td></td>"
		strTemp = strTemp & "</tr></form></table>"
		showpage = strTemp
	End Function
	'=================================================
	'��������GetFilePic
	'��  �ã���ȡ�ļ�ͼƬ
	'=================================================
	Private Function GetFilePic(sName)
		Dim FileName, Icon
		FileName = LCase(GetExtensionName(sName))
		Select Case FileName
			Case "gif", "jpg", "bmp", "png"
				Icon = sName
			Case "exe"
				Icon = "Images/FileIcon/file_exe.gif"
			Case "rar"
				Icon = "Images/FileIcon/file_rar.gif"
			Case "zip"
				Icon = "Images/FileIcon/file_zip.gif"
			Case "swf"
				Icon = "Images/FileIcon/file_flash.gif"
			Case "rm", "wma"
				Icon = "Images/FileIcon/file_rm.gif"
			Case "mid"
				Icon = "Images/FileIcon/file_media.gif"
			Case Else
				Icon = "Images/FileIcon/file_other.gif"
		End Select
		GetFilePic = Icon
	End Function
	'=================================================
	'��������GetExtensionName
	'��  �ã���ȡ�ļ���չ��
	'=================================================
	Private Function GetExtensionName(ByVal sName)
		Dim FileName
		FileName = Split(sName, ".")
		GetExtensionName = FileName(UBound(FileName))
	End Function
	'=================================================
	'��������GetFileSize
	'��  �ã���ʽ���ļ��Ĵ�С
	'=================================================
	Private Function GetFileSize(ByVal n)
		Dim FileSize
		FileSize = n / 1024
		FileSize = FormatNumber(FileSize, 2)
		If FileSize < 1024 And FileSize > 1 Then
			GetFileSize = "<font color=red>" & FileSize & "</font>&nbsp;KB"
		ElseIf FileSize > 1024 Then
			GetFileSize = "<font color=red>" & FormatNumber(FileSize / 1024, 2) & "</font>&nbsp;MB"
		Else
			GetFileSize = "<font color=red>" & n & "</font>&nbsp;Bytes"
		End If
	End Function
	'=================================================
	'��������DelFile
	'��  �ã�ɾ���ļ�
	'=================================================
	Private Sub DelFile()
		Dim fso, i
		Dim strFileName, strFilePath
		Dim strFolderName, strFolderPath
		'---- ɾ���ļ�
		If Trim(Request("FileName")) <> "" Then
			strFileName = Split(Request("FileName"), ",")
			If UBound(strFileName) <> -1 Then 'ɾ���ļ�
				Set fso = KS.InitialObject(KS.Setting(99))
				For i = 0 To UBound(strFileName)
					strFilePath = Server.MapPath(FilePath & Trim(strFileName(i)))
					If fso.FileExists(strFilePath) Then
						fso.DeleteFile strFilePath, True
					End If
				Next
				Set fso = Nothing
			End If
		End If
		'---- ɾ���ļ���
		If Trim(Request("FolderName")) <> "" Then
			strFolderName = Split(Request("FolderName"), ",")
			If UBound(strFolderName) <> -1 Then 'ɾ���ļ�
				Set fso = KS.InitialObject(KS.Setting(99))
				For i = 0 To UBound(strFolderName)
					strFolderPath = Server.MapPath(FilePath & Trim(strFolderName(i)))
					If fso.FolderExists(strFolderPath) Then
						fso.DeleteFolder strFolderPath, True
					End If
				Next
				Set fso = Nothing
			End If
		End If
		Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
	End Sub
	'=================================================
	'��������DelAllDirFile
	'��  �ã�ɾ�������ļ����ļ���
	'=================================================
	Private Sub DelAllDirFile()
		Dim fso, oFolder
		Dim DirFile, DirFolder
		Dim tempPath
		
		Set fso = KS.InitialObject(KS.Setting(99))
		If fso.FolderExists(fullPath) Then
			Set oFolder = fso.GetFolder(fullPath)
			'---- ɾ�������ļ�
			For Each DirFile In oFolder.Files
				tempPath = fullPath & "\" & DirFile.Name
				fso.DeleteFile tempPath, True
			Next
			'---- ɾ��������Ŀ¼
			For Each DirFolder In oFolder.SubFolders
				tempPath = fullPath & "\" & DirFolder.Name
				fso.DeleteFolder tempPath, True
			Next
			Set oFolder = Nothing
		End If
		Set fso = Nothing
		Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
	End Sub
	'=================================================
	'��������DelThisAllFile
	'��  �ã�ɾ����ǰĿ¼�����ļ�
	'=================================================
	Private Sub DelThisAllFile()
		Dim fso, oFolder
		Dim DirFiles
		Dim tempPath
		
		Set fso = KS.InitialObject(KS.Setting(99))
		If fso.FolderExists(fullPath) Then
			Set oFolder = fso.GetFolder(fullPath)
			'---- ɾ�������ļ�
			For Each DirFiles In oFolder.Files
				tempPath = fullPath & "\" & DirFiles.Name
				fso.DeleteFile tempPath, True
			Next
			Set oFolder = Nothing
		End If
		Set fso = Nothing
		Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
	End Sub
	'=================================================
	'��������DelEmptyFolder
	'��  �ã�ɾ�����п��ļ���
	'=================================================
	Private Sub DelEmptyFolder()
		Dim fso, oFolder
		Dim DirFolder, tempPath
		
		Set fso = KS.InitialObject(KS.Setting(99))
		If fso.FolderExists(fullPath) Then
			Set oFolder = fso.GetFolder(fullPath)
			'---- ɾ�����п���Ŀ¼
			For Each DirFolder In oFolder.SubFolders
				If DirFolder.Size = 0 Then
					tempPath = fullPath & "\" & DirFolder.Name
					fso.DeleteFolder tempPath, True
				End If
			Next
			Set oFolder = Nothing
		End If
		Set fso = Nothing
		Response.Redirect (Request.ServerVariables("HTTP_REFERER"))
	End Sub
	'=================================================
	'��������ShowUploadMain
	'��  �ã���ʾ�ϴ��ļ���ҳ��
	'=================================================
	Private Sub ShowUploadMain()
		Dim maxperpage, CurrentPage, TotalNumber, Pcount
		Dim fso, FileCount, TotleSize, totalPut
		
		maxperpage = 20 '###ÿҳ��ʾ��
		
		If IsNumeric(Request("page")) And Trim(Request("page")) <> "" Then
			CurrentPage = CLng(Request("page"))
		Else
			CurrentPage = 1
		End If
		If CLng(CurrentPage) = 0 Then CurrentPage = 1
		On Error Resume Next
		If Not KS.IsObjInstalled(KS.Setting(99)) Then
			Response.Write "<b><font color=red>��ķ�������֧�� fso(Scripting.FileSystemObject)! ����ʹ�ñ�����</font></b>"
		End If
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<meta http-equiv='Content-Type' content='text/html; chaRSet=gb2312'>"
		Response.Write "<title>Digg��¼����</title>"
		Response.Write "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		Response.Write "<script>" &vbnewline
		Response.Write "function CheckAll(form) {  "&vbnewline
		Response.Write "	for (var i=0;i<form.elements.length;i++)  {  "&vbnewline
		Response.Write "		var e = form.elements[i];  "&vbnewline
		Response.Write "		if (e.name != 'chkall')  "&vbnewline
		Response.Write "		e.checked = true // form.chkall.checked;  "&vbnewline
		Response.Write "	}  "&vbnewline
		Response.Write "} "&vbnewline
		 
		Response.Write "function ContraSel(form) {"&vbnewline
		Response.Write "	for (var i=0;i<form.elements.length;i++){"&vbnewline
		Response.Write "		var e = form.elements[i];"&vbnewline
		Response.Write "		if (e.name != 'chkall')"&vbnewline
		Response.Write "		e.checked=!e.checked;"&vbnewline
		Response.Write "	}"&vbnewline
		Response.Write "}"&vbnewline
		Response.Write "</script>"&vbnewline
		Response.Write "</head>"
		
		Response.Write "<body topmargin='0' leftmargin='0'>"
		Response.Write "<ul id='mt'> <div style='font-weight:bold;margin-top:10px'><a href='?'>�ϴ��ļ�����</a> | <a href='?action=clear'>���������ļ�</a></div></ul>"
		Response.Write "<table border=0 align=center cellpadding=3 style='margin-top:5px' cellspacing=1 width='99%' class='ctable'>"
		Response.Write "<tr>"
		Response.Write "        <td class=clefttitle colspan=""2"">"
		Call ShowChildFolder
		Response.Write "</td>"
		Response.Write "</tr>"
		Response.Write "<tr>"
		Response.Write "        <td width=""50%"" class=clefttitle>��ǰĿ¼��" & FilePath & "</td>"
		Response.Write "        <td width=""50%"" align=center class=clefttitle>"
		Response.Write "<!--<a href=""?action=clear&UploadDir=" & Request("UploadDir") & """>���������ļ�</a>--> &nbsp;&nbsp;"
		If Trim(Request("ThisDir")) <> "" Then

			Response.Write "<a href=""?UploadDir=" & Left(Request("UploadDir"),Len(Request("UploadDir"))-Len(Mid(Request("UploadDir"), InStrRev(Request("UploadDir"), "/")))) & "&ThisDir=" & Request("ThisDir") & """>��������һ��Ŀ¼</a>"
		End If
		Response.Write "</td>"
		Response.Write "</tr>"
		Response.Write "</table><br>" & vbNewLine

		Set fso = KS.InitialObject(KS.Setting(99))
		If fso.FolderExists(fullPath) Then
			Dim fsoFile, fsoFileSize
			Dim DirFiles, DirFolder
			Set fsoFile = fso.GetFolder(fullPath)
			Dim c
			FileCount = fsoFile.Files.Count
			TotleSize = GetFileSize(fsoFile.Size)
			totalPut = fsoFile.Files.Count
			If CurrentPage < 1 Then
				CurrentPage = 1
			End If
			If (CurrentPage - 1) * maxperpage > totalPut Then
				If (totalPut Mod maxperpage) = 0 Then
					CurrentPage = totalPut \ maxperpage
				Else
					CurrentPage = totalPut \ maxperpage + 1
				End If
			End If
			FileCount = 0
			c = 0
			Response.Write "<table border=0 align=center cellpadding=3 cellspacing=1 class=ctable width='99%'>" & vbNewLine
			Response.Write "<tr><td colspan=4 class=clefttitle>" & vbNewLine
			Response.Write showpage(CurrentPage, totalPut, maxperpage, TotleSize)
			Response.Write "</td></tr>" & vbNewLine
			Response.Write "<form name=""myform"" method=""post"" action='KS.AdminFiles.asp'>" & vbCrLf
			Response.Write "<tr>" & vbNewLine
			Response.Write "<input type=hidden name=action value='del'>" & vbNewLine
			Response.Write "<input type=hidden name=UploadDir value='" & Request("UploadDir") & "'>" & vbNewLine
			Response.Write "<input type=hidden name=ThisDir value='" & Request("ThisDir") & "'>" & vbNewLine
			For Each DirFiles In fsoFile.Files
				c = c + 1
				If c > maxperpage * (CurrentPage - 1) Then
					Response.Write "<td class=clefttitle>"
					Response.Write "<div align=center><a href='" & FilePath & DirFiles.Name & "'target=_blank><img src='" & GetFilePic(FilePath & DirFiles.Name) & "' width=140 height=100 border=0 alt='���ͼƬ�鿴ԭʼ�ļ���'></a></div>"
					Response.Write "�ļ�����<a href='" & FilePath & DirFiles.Name & "'target=_blank>" & DirFiles.Name & "</a><br>"
					Response.Write "�ļ���С��" & GetFileSize(DirFiles.Size) & "<br>"
					Response.Write "�ļ����ͣ�" & DirFiles.Type & "<br>"
					Response.Write "�޸�ʱ�䣺" & DirFiles.DateLastModified & "<br>"
					Response.Write "���������<input type=checkbox name=FileName value='" & DirFiles.Name & "' checked> ѡ��&nbsp;&nbsp;"
					Response.Write "<a href='?action=del&UploadDir=" & Request("UploadDir") & "&ThisDir=" & Request("ThisDir") & "&FileName=" & DirFiles.Name & "' onclick=""return confirm('��ȷ��Ҫɾ�����ļ���!');"">��ɾ��</a>"
					FileCount = FileCount + 1
					Response.Write "</td>" & vbNewLine
					If (FileCount Mod 4) = 0 And FileCount < maxperpage And c < totalPut Then
						Response.Write "</tr>" & vbNewLine & "<tr>" & vbNewLine
					End If
				End If
				If FileCount >= maxperpage Then Exit For
			Next
			Response.Write "</tr>" & vbNewLine
			Response.Write "<tr><td colspan=4 class=clefttitle>" & vbNewLine
			Response.Write "<input class=Button type=button name=chkall value='ȫѡ' onClick=""CheckAll(this.form)"">&nbsp;<input class=Button type=button name=chksel value='��ѡ' onClick=""ContraSel(this.form)"">" & vbNewLine
			Response.Write "&nbsp;&nbsp;<input class=Button type=submit name=Submit2 value='ɾ��ѡ�е��ļ�' onClick=""return confirm('ȷ��Ҫɾ��ѡ�е��ļ���')"">" & vbNewLine
			Response.Write "&nbsp;&nbsp;<input class=Button type=submit name=Submit3 value='ɾ�������ļ�' onClick=""document.myform.action.value='DelThisAllFile';return confirm('ȷ��Ҫɾ����ǰĿ¼�����ļ���')"">" & vbNewLine
			Response.Write "&nbsp;&nbsp;<input class=Button type=submit name=Submit4 value='ɾ�������ļ����ļ���' onClick=""document.myform.action.value='DelAllDirFile';return confirm('ȷ��Ҫɾ����ǰĿ¼���ļ����ļ�����')"">" & vbNewLine
			Response.Write "&nbsp;&nbsp;<input class=Button type=submit name=Submit5 value='ɾ�����п��ļ���' onClick=""document.myform.action.value='DelEmptyFolder';return confirm('ȷ��Ҫɾ����ǰĿ¼���п��ļ�����')"">" & vbNewLine
			Response.Write "</tr></form>" & vbNewLine
			Response.Write "<tr><td colspan=4 class=clefttitle>" & vbNewLine
			Response.Write showpage(CurrentPage, totalPut, maxperpage, TotleSize)
			Response.Write "</td></tr>" & vbNewLine
			Response.Write "</table>"
			
			Response.Write " <div style='text-align:center;margin-top:27px'><input class=Button type=button name=Submit2 value=' һ����������δ�����������ļ� ' onclick=""if (confirm('��ȷ��Ҫһ������������õ��ļ��𣿴˲��������棬�����ȱ���UploadFilesĿ¼����ִ�У�')){location.href='?action=delete';}""></div>"

		Else
			Response.Write "��Ŀ¼û���κ��ļ���"
		End If
		Set fsoFile = Nothing: Set fso = Nothing
	End Sub
	'=================================================
	'��������ClearUploadFile
	'��  �ã��������õ��ϴ��ļ�
	'=================================================
	Private Sub ClearUploadFile()
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<meta http-equiv='Content-Type' content='text/html; chaRSet=gb2312'>"
		Response.Write "<title>����</title>"
		Response.Write "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		Response.Write "</head>"
		
		Response.Write "<body topmargin='0' leftmargin='0'>"
		Response.Write "<ul id='mt'> <div style='text-align:center;font-weight:bold;margin-top:10px'>"
		If LCase(Request("UploadDir")) = "DownUrl" Then
			Response.Write "�������õ��ϴ��ļ�"
		Else
			Response.Write "�������õ��ϴ�ͼƬ"
		End If
		Response.Write "</div></ul>"
		Response.Write "<table border=0 align=center cellpadding=3 style='margin-top:5px' cellspacing=1 width='99%' class='ctable'>"
		
		Response.Write "<form name=""myform"" method=""get"" action='KS.AdminFiles.asp'>" & vbCrLf
		Response.Write "<input type=hidden name=action value='delete'>" & vbNewLine
		Response.Write "<input type=hidden name=UploadDir value='" & Request("UploadDir") & "'>" & vbNewLine
		Response.Write "<tr><td class=clefttitle>" & vbNewLine
		Response.Write "<br>&nbsp;&nbsp;�١������վ��ʹ��һ��ʱ��󣬾ͻ�����������������ļ���������˷����Ŀռ䣬���Զ���ʹ�ñ����ܽ�������<br>"
		Response.Write "<br>&nbsp;&nbsp;�ڡ��°棨V6���Ժ�İ汾���ϴ�Ŀ¼ͳһ����ΪUploadFiles,�˹��ܽ���UpLoadFilesĿ¼ִ�������ܣ�<br>"
		Response.Write "<br>&nbsp;&nbsp;�ۡ�����ϴ��ļ��ܶ࣬�������ݿ����Ϣ���϶ִ࣬�б�������Ҫ�ķ��൱����ʱ�䣬���ڷ�������ʱִ�б�������<br>"
		Response.Write "<br></td></tr>" & vbNewLine
		Response.Write "<tr align=center><td  class=clefttitle>��ѡ��Ҫ�����Ŀ¼��"
		Call ShowFolderPath
		Response.Write "<input class=Button type=submit name=Submit2 value=' ��ʼ���������ļ� ' onclick=""return confirm('��ȷ��Ҫ����������õ��ļ���');"">"
		Response.Write " <input class=Button type=button name=Submit2 value=' һ���������������ļ� ' onclick=""if (confirm('��ȷ��Ҫһ������������õ��ļ��𣿴˲��������棬�����ȱ���UploadFilesĿ¼����ִ�У�')){location.href='?action=delete';}"">"
		Response.Write " ����<a href='?'>�����ϴ�����</a>"
		Response.Write "</td></tr></form>" & vbNewLine
		Response.Write "</table>"
	End Sub
	'=================================================
	'��������ShowFolderPath
	'��  �ã���ʾ��Ŀ¼�˵�
	'=================================================
	Private Sub ShowFolderPath()
		Dim fso, fsoFile, DirFolder
		Dim strFolderPath
		On Error Resume Next
		strFolderPath = ChannelDir & UploadDir
		strFolderPath = Server.MapPath(strFolderPath)
		Set fso = KS.InitialObject(KS.Setting(99))
		If fso.FolderExists(strFolderPath) Then
			Set fsoFile = fso.GetFolder(strFolderPath)
			Response.Write "<select name=""path"">" & vbNewLine
			For Each DirFolder In fsoFile.SubFolders
			   If IsDate(DirFolder.Name) Then
				Response.Write "	<option value=""" & DirFolder.Name & """>" & DirFolder.Name & "</option>" & vbNewLine
			   End If
			Next
			'Response.Write "	<option value="""">�ϴ���Ŀ¼</option>" & vbNewLine
			Response.Write "</select>" & vbNewLine
			Set fsoFile = Nothing
		Else
			'Response.Write "û���ҵ��ļ��У�"
		End If
		Set fso = Nothing
	End Sub
	
	Sub DeleteFile(strFolderPath,i)
		Dim fso, fsoFile, DirFiles
		Dim strFileName,ParentPath
		Dim strFilePath
		Set fso = KS.InitialObject(KS.Setting(99))
		If fso.FolderExists(strFolderPath) Then
		    Set fsoFile = fso.GetFolder(strFolderPath)
			ParentPath=strFolderPath
			For Each DirFiles In fsoFile.SubFolders
			 Call DeleteFile(ParentPath & "\" & DirFiles.Name,i)
			Next
			
			For Each DirFiles In fsoFile.Files
			    
				strFileName = DirFiles.Name
				strFilePath = strFolderPath & "\" & DirFiles.Name
				If Not CheckFileExists(strFilePath) Then
					i = i + 1
					fso.DeleteFile(strFilePath)
				End If
			Next
			Set fsoFile = Nothing
		End If
		Set fso = Nothing
	End Sub
	'=================================================
	'��������DelUselessFile
	'��  �ã�ɾ���������õ��ϴ��ļ�
	'=================================================
	Private Sub DelUselessFile()
		Dim SQL,i
		Dim fso, fsoFile, DirFiles
		Dim strFileName,strFolderPath
		Dim strFilePath,strDirName
		Server.ScriptTimeout = 9999999
		'On Error Resume Next
		i=0
		If Len(Request("path")) > 0 Then
			strDirName = Request("path") & "/"
		Else
			strDirName = vbNullString
		End If
		strFolderPath = ChannelDir & UploadDir & strDirName
		strFolderPath = Server.MapPath(strFolderPath)
		Call DeleteFile(strFolderPath,i)
	

		Response.Write "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		Response.Write " <Br><br><br><table align='center' width=""95%"" height='200' class='ctable' cellpadding=""1"" cellspacing=""1"">"
		Response.Write "	  <tr class=""sort""> "
		Response.Write "		<td  height=""28"" colspan=2>ϵͳ������ʾ��Ϣ</td>" & vbcrlf
		Response.Write "	  </tr>"
		Response.Write "    <tr class='tdbg'>"
		Response.Write "          <td align='center'><img src='images/succeed.gif'></td>"
		Response.Write "<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<li>�ļ�������ɣ�</li><li>һ��������<font color=red><b>" & i & "</b></font>�������ļ���</b><br></td></tr>"
		Response.Write "	  <tr class=""sort""> "
		Response.Write "		<td  height=""28"" colspan=2><a href='#' onclick='javascript:history.back(-1);'>������һ��</a> <a  href='?'>�����ϴ�Ŀ¼</a></td>" & vbcrlf
		Response.Write "	  </tr>"
		Response.Write "</table>"
	End Sub
	Public Function CheckFileExists(ByVal str)
	   str=Lcase(str)
	   str=replace(str,"\","/")
	   if instr(str,lcase(KS.Setting(91)))<0 then
	     CheckFileExists=false
		 exit function
	   end if
	   
	   str=Split(str,lcase(KS.Setting(91)))	   
	   If Ubound(Str)=0 Then
	     CheckFileExists=false
		 exit function
	   End If
	   Dim FileName
        
	   FileName=lcase(KS.Setting(91)) & str(1)
	  
		Dim Rs,SQL,Param
		IF INSTR(FileName,"[")<>0 and Instr(FileName,"]")<>0 then
		 FileName=Split(FileName,"[")(0) & "%" & Split(FileName,"]")(1)
		end if
		SQL = "SELECT TOP 1 ID FROM [KS_UploadFiles] WHERE FileName like '%" & FileName & "'"
		Set Rs = Conn.Execute(SQL)
		If Rs.BOF And Rs.EOF Then
			CheckFileExists = False
		Else
			CheckFileExists = True
		End If
		Set Rs = Nothing
	End Function
	
End Class
%>