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
Set KSCls = New RefreshSpecialSave
KSCls.Kesion()
Set KSCls = Nothing

Class RefreshSpecialSave
        Private KS,KSRObj
		Private RefreshFlag
		Private ReturnInfo
		Private StartRefreshTime
		Private ChannelID
		Private Types
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSRObj=New Refresh
		End Sub
		Function Kesion()
		Types = Request("Types")             'Index ����ר����ҳ���� Special ����ר��ҳ����
		RefreshFlag = Request("RefreshFlag") 'ȡ���ǰ���������ˢ��,��Folder����ָ����ר�� All��������ר��
		'ˢ��ʱ��
		StartRefreshTime = Request("StartRefreshTime")
		If StartRefreshTime = "" Then StartRefreshTime = Timer()
		  Select Case Types
			 Case "Special"          'ˢ��ר��ҳ
				 Call RefreshSpecial
			 Case "Index"            'ˢ��ר����ҳ
				 Call RefreshSpecialIndex
			 Case "ChannelSpecial"   'ˢ��Ƶ��ר���б�ҳ
				 Call RefreshChannelSpecial
		End Select
		End Function
		Sub Main()
		 Response.Write ("<html>")
		 Response.Write ("<head>")
		 Response.Write ("<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">")
		 Response.Write ("<title>ϵͳ��Ϣ</title>")
		 Response.Write ("</head>")
		 Response.Write ("<link rel=""stylesheet"" href=""Admin_Style.css"">")
		 Response.Write ("<body oncontextmenu=""return false;"">")
				Response.Write "<br><br><br><table id=""BarShowArea"" width=""400"" border=""0"" align=""center"" cellspacing=""1"" cellpadding=""1"">"
				Response.Write "<tr> "
				Response.Write "<td bgcolor=000000>"
				Response.Write " <table width=""400"" border=""0"" cellspacing=""0"" cellpadding=""1"">"
				Response.Write "<tr> "
				Response.Write "<td bgcolor=ffffff height=9><img src=""../images/114_r2_c2.jpg"" width=0 height=10 id=img2 name=img2 align=absmiddle></td></tr></table>"
				Response.Write "</td></tr></table>"
				Response.Write "<table width=""550"" border=""0"" align=""center"" cellspacing=""1"" cellpadding=""1""><tr> "
				Response.Write "<td align=center> <span id=txt2 name=txt2 style=""font-size:9pt"">0</span><span id=txt4 style=""font-size:9pt"">%</span></td></tr> "
				Response.Write "<tr><td align=center><span id=txt3 name=txt3 style=""font-size:9pt"">0</span></td></tr>"
				Response.Write "</table>"

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
		
		'=============================================================================================
		'����Ϊ��ģ����Ӧ����ĺ���
		'===============================================================================================
		
		'����ר����ҳ�Ĵ������
		Sub RefreshSpecialIndex()
		   Dim InstallDir, IndexFile, SaveFilePath
		   Dim SpecialDir, FileContent, Domain
		   FCls.RefreshType = "SpecialIndex"  '����ˢ�����ͣ��Ա�ȡ�õ�ǰλ�õ�����
		   FCls.RefreshFolderID = "0"         '���õ�ǰˢ��Ŀ¼ID Ϊ"0" ��ȡ��ͨ�ñ�ǩ
		   FCls.CurrSpecialID="" '�����ǰר��ID
		   
		   InstallDir = KS.Setting(3)
		   SpecialDir = KS.Setting(95)
		   If Left(SpecialDir, 1) = "/" Or Left(SpecialDir, 1) = "\" Then SpecialDir = Right(SpecialDir, Len(SpecialDir) - 1)
		   IndexFile = KS.Setting(5)
			SaveFilePath = InstallDir & SpecialDir
			FileContent = KSRObj.LoadTemplate(KS.Setting(111))
			If FileContent = "" Then
			  ReturnInfo = "���ݿ����Ҳ���ר����ҳģ��"
			  Call Main
			  Response.End
			Else
			  On Error Resume Next
			  FileContent = KSRObj.ReplaceLableFlag(KSRObj.ReplaceAllLabel(FileContent)) '�滻������ǩ
			  FileContent = KSRObj.ReplaceGeneralLabelContent(FileContent)  '�滻ͨ�ñ�ǩ ��{$GetWebmaster}
			  If Err Then
			   ReturnInfo = Err.Description
				 Err.Clear
				Call Main
				Response.End
			  End If
			  Call KS.CreateListFolder(SaveFilePath)
			  Call KSRObj.FSOSaveFile(FileContent, SaveFilePath & IndexFile)
			  If Err Then
				ReturnInfo = Err.Description
				 Err.Clear
				Call Main
				Response.End
			  End If
			  Domain = KS.GetDomain

			  ReturnInfo = "ר����ҳ�����ɹ����ܷ�ʱ:<font color=red>" & Left((Timer() - StartRefreshTime), 4) & "</font>��<br><br>"
			  ReturnInfo = ReturnInfo & "������: <a href=" & Domain & SpecialDir & IndexFile & " target=_blank>���ר����ҳ</a><br><br>"
			  ReturnInfo = ReturnInfo & "<input name=""button1"" type=""button"" onclick=""javascript:location='RefreshSpecial.asp';"" class=""button"" value="" �� �� "">"
			  Call Main
			    Response.Write "<script>" & vbCrlf
			  	Response.Write "img2.width=400;" & vbCrLf
				Response.Write "txt2.innerHTML=""����ר����ҳ������100"";" & vbCrLf
				Response.Write "document.getElementById('txt3').parentElement.style.display='none';" & vbCrLf
				Response.Write "</script>" & vbCrLf
			End If
		End Sub
		'����ר�����Ĵ������
		Sub RefreshChannelSpecial()
		 Dim FolderID, RefreshSql, RefreshTotalNum, RefreshRS, NewsTotalNum, NewsNo
		  RefreshSql = Trim(Request("RefreshSql"))
		  NewsNo = Request("NewsNo")
		 If NewsNo = "" Then NewsNo = 0
		 If RefreshSql = "" Then
		  Select Case RefreshFlag
			Case "Folder"
				FolderID = Replace(Request("FolderID")," ","")
				If FolderID <> "" Then
				  RefreshSql = "Select * from [KS_SpecialClass] where ClassID IN (" & FolderID & ") Order By ClassID"
				Else
				  RefreshSql = "Select * From [KS_SpecialClass] Where 1=0"
				End If
		   Case "All"
				RefreshSql = "Select * from [KS_SpecialClass] Order By ClassID"
		   Case Else
			RefreshSql = ""
			RefreshTotalNum = 0
		  End Select
		End If
		If RefreshSql <> "" Then
		    Call Main
			Set RefreshRS = Server.CreateObject("ADODB.RecordSet")
			RefreshRS.Open RefreshSql, Conn, 1, 1
			NewsTotalNum = RefreshRS.RecordCount
			If RefreshRS.EOF Then
				ReturnInfo = "û��Ҫˢ�µ�ר�����&nbsp;&nbsp;<br><input name=""button1"" type=""button"" onclick=""javascript:location='RefreshSpecial.asp';"" class=""button"" value="" �� �� "">"
				Set RefreshRS = Nothing
			Else
				For NewsNo=1 To NewsTotalNum
				   Call KSRObj.RefreshSpecialClass(RefreshRS)  '����Ƶ��ר��ˢ�º���
				   Call InnerJs(NewsNo,NewsTotalNum,"��ר�����")
				   RefreshRS.MoveNext
				Next
			End If
				Response.Write "<script>"
				Response.Write "img2.width=400;" & vbCrLf
				Response.Write "txt2.innerHTML=""����ר����������100"";" & vbCrLf
				Response.Write "txt3.innerHTML=""�ܹ������� <font color=red><b>" & NewsTotalNum & "</b></font> ��ר�����,�ܷ�ʱ:<font color=red>" & Left((Timer() - StartRefreshTime), 4) & "</font> ��<br><br><input name='button1' type='button' onclick=javascript:location='RefreshSpecial.asp'; class='button' value=' �� �� '>"";</script>" & vbCrLf

			Set RefreshRS = Nothing
		End If
		End Sub
		'����ר��ҳ�Ĵ������
		Sub RefreshSpecial()
		  Dim FolderID, RefreshSql, RefreshTotalNum, RefreshRS, NewsTotalNum, NewsNo
		  RefreshSql = Trim(Request("RefreshSql"))
		  NewsNo = Request("NewsNo")
		 If NewsNo = "" Then NewsNo = 0
		 If RefreshSql = "" Then
		  Select Case RefreshFlag
		  	Case "ID"
				RefreshSql = "Select * From KS_Special where specialid in(" & KS.G("ID") & ") Order By SpecialAddDate Desc"
			Case "New"
				Dim TotalNum
				TotalNum = Request.Form("TotalNum")
				If TotalNum = "" Then TotalNum = 20
				RefreshSql = "Select Top " & TotalNum & " * From KS_Special Order By SpecialAddDate Desc"
			Case "Folder"
				FolderID = KS.FilterIDs(Request("FolderID"))
				If FolderID <> "" Then
				RefreshSql = "Select * from [KS_Special] where  ClassID IN (" & FolderID & ") Order By SpecialAddDate Desc"
				Else
				RefreshSql = "Select * From [KS_Special] Where 1=0"
				End If
		   Case "All"
				'RefreshSql = "Select * from [KS_Special] a inner join ks_channel b on a.channelid=b.channelid where b.FsoHtmlTF=1 order by specialadddate desc"
				RefreshSql = "Select * from [KS_Special] order by specialadddate desc"
		   Case Else
			RefreshSql = ""
			RefreshTotalNum = 0
		  End Select
		End If
		If RefreshSql <> "" Then
			Call Main
			Set RefreshRS = Server.CreateObject("ADODB.RecordSet")
			RefreshRS.Open RefreshSql, Conn, 1, 1
			NewsTotalNum = RefreshRS.RecordCount
			If RefreshRS.EOF Then
				Response.Write "<script>img2.width=""0"";" & vbCrLf
				Response.Write "txt2.innerHTML=""�Բ���û�п����ɵ�ר�⣡<br><br><input name='button1' type='button' onclick=javascript:location='RefreshSpecial.asp'; class='button' value=' �� �� '>"";" & vbCrLf
				Response.Write "txt3.innerHTML="""";" & vbCrLf
				Response.Write "txt4.innerHTML="""";" & vbCrLf
				Response.Write "document.all.BarShowArea.style.display='none';" & vbCrLf
				Response.Write "</script>" & vbCrLf
				Response.Flush
				Set RefreshRS = Nothing
			Else
			   For NewsNo=1 To NewsTotalNum
				   Call KSRObj.RefreshSpecials(RefreshRS)  '����ר��ˢ�º���
                   Call InnerJS(NewsNo,NewsTotalNum,"��ר��")
				   RefreshRS.MoveNext
			  Next 
				Response.Write "<script>"
				Response.Write "img2.width=400;" & vbCrLf
				Response.Write "txt2.innerHTML=""����ר�������100"";" & vbCrLf
				Response.Write "txt3.innerHTML=""�ܹ������� <font color=red><b>" & NewsTotalNum & "</b></font> ��,�ܷ�ʱ:<font color=red>" & Left((Timer() - StartRefreshTime), 4) & "</font> ��<br><br><input name='button1' type='button' onclick=javascript:location='RefreshSpecial.asp'; class='button' value=' �� �� '>"";</script>" & vbCrLf
			End If
			Set RefreshRS = Nothing
		End If
		End Sub
        Sub InnerJS(NowNum,TotalNum,itemname)
		  With Response
				.Write "<script>"
				'.Write "fsohtml.innerHTML='" & FsoHtmlList & "';" & vbCrLf
				.Write "img2.width=" & Fix((NowNum / TotalNum) * 400) & ";" & vbCrLf
				.Write "txt2.innerHTML=""���ɽ���:" & FormatNumber(NowNum / TotalNum * 100, 2, -1) & """;" & vbCrLf
				.Write "txt3.innerHTML=""�ܹ���Ҫ���� <font color=red><b>" & TotalNum & "</b></font> " & itemname & ",<font color=red><b>�ڴ˹���������ˢ�´�ҳ�棡����</b></font> ϵͳ�������ɵ� <font color=red><b>" & NowNum & "</b></font> " & itemname & """;" & vbCrLf
				.Write "img2.title=""(" & NowNum & ")"";</script>" & vbCrLf
				.Flush
		  End With
		End Sub
End Class
%> 
