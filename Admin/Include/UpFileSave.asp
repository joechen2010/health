<% Option Explicit %>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New UpFileSave
KSCls.Kesion()
Set KSCls = Nothing

Const NoAllowExt = "asa|asax|ascs|ashx|asmx|asp|aspx|axd|cdx|cer|cfm|config|cs|csproj|idc|licx|rem|resources|resx|shtm|shtml|soap|stm|vb|vbproj|vsdisco|webinfo"    '�������ϴ�����
Const NeedCheckFileMimeExt = "gif|jpg|jpeg|jpe|bmp|png|swf|mid|mp3|wmv|asf|avi|mpg|ram|rm|ra|rar|exe|doc|zip" '������Ҫ����Ƿ�α����ļ�����

Class UpFileSave
        Private KS
		Dim FilePath,MaxFileSize,AllowFileExtStr,AutoReName,RsConfigObj,BasicType,ChannelID,UpType
		Dim FormName,Path,UpLoadFrom,TempFileStr,FormPath,ThumbFileName,ThumbPathFileName
		Dim UpFileObj,FsoObjName,AddWaterFlag,T,CurrNum,CreateThumbsFlag,FieldName,	U_FileSize
		Dim DefaultThumb    '�趨�ڼ���Ϊ����ͼ
		Dim ReturnValue
		Private Sub Class_Initialize()
		  Set T=New Thumb
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set T=Nothing
		 Set KS=Nothing
		End Sub
		
		Function IsSelfRefer() 
			Dim sHttp_Referer, sServer_Name 
			sHttp_Referer = CStr(Request.ServerVariables("HTTP_REFERER")) 
			sServer_Name = CStr(Request.ServerVariables("SERVER_NAME")) 
			If Mid(sHttp_Referer, 8, Len(sServer_Name)) = sServer_Name Then 
			IsSelfRefer = True 
			Else 
			IsSelfRefer = False 
			End If 
		End Function 

		
		Sub Kesion()
		Response.Write("<style type='text/css'>" & vbcrlf)
		Response.Write("<!--" & vbcrlf)
		Response.Write("body {background:#f0f0f0;" & vbcrlf)
		Response.Write("	margin-left: 0px;" & vbcrlf)
		Response.Write("	margin-top: 0px;" & vbcrlf)
		Response.Write("}" & vbcrlf)
		Response.Write("-->" & vbcrlf)
		Response.Write("</style>" & vbcrlf)
		
		If KS.IsNul(KS.C("AdminName")) Or KS.IsNul(KS.C("AdminPass")) Or KS.IsNul(KS.C("PowerList"))="" Or KS.IsNUL(KS.C("UserName")) Then
			Response.Write "<script>alert('û�е�¼!');history.back();</script>"
			Response.end
		End If
		
		
		 If Trim(Request.ServerVariables("HTTP_REFERER"))="" Then
			Response.Write "<script>alert('�Ƿ��ϴ�1��');history.back();</script>"
			Response.end
		 End If
		 if instr(lcase(Request.ServerVariables("HTTP_REFERER")),"ks.upfileform.asp")<=0 and instr(lcase(Request.ServerVariables("HTTP_REFERER")),"upfileform.asp")<=0 then
			Response.Write "<script>alert('�Ƿ��ϴ���');history.back();</script>"
			Response.end
		 end if
		 if IsSelfRefer=false Then
			Response.Write "<script>alert('�벻Ҫ�Ƿ��ϴ���');history.back();</script>"
			Response.end
		 End If
		 
		 

		FsoObjName=KS.Setting(99)
		Set UpFileObj = New UpFileClass
		UpFileObj.GetData
		FormPath=Replace(UpFileObj.Form("Path"),".","") 
		IF Instr(FormPath,KS.Setting(3))=0 Then	FormPath=KS.Setting(3) & FormPath
		Call KS.CreateListFolder(FormPath)       '�����ϴ��ļ����Ŀ¼
		FilePath=Server.MapPath(FormPath) & "\"
		If KS.Setting(97)=1 Then
		FormPath=KS.Setting(2) & FormPath
		End if
		AutoReName = KS.ChkClng(UpFileObj.Form("AutoRename"))
		UpLoadFrom=UpFileObj.Form("UpLoadFrom")        '0--ͨ�öԻ��� 2-- ͼƬ�����ϴ� 31--������������ͼ 32--���������ļ� 41--������������ͼ 42--�������ĵĶ����ļ�
		IF UpLoadFrom="" then  UpLoadFrom=0
		
		BasicType=KS.ChkClng(UpFileObj.Form("BasicType"))        ' 2-- ͼƬ�����ϴ� 3--������������ͼ/�ļ� 41--������������ͼ 42--�������ĵĶ����ļ�
		ChannelID=KS.ChkClng(UpFileObj.Form("ChannelID")) 
		UpType=UpFileObj.Form("UpType")
		
		CurrNum=0
		CreateThumbsFlag=false
		DefaultThumb=UpFileObj.Form("DefaultUrl")
		if DefaultThumb="" then DefaultThumb=0
		AddWaterFlag = UpFileObj.Form("AddWaterFlag")
		If AddWaterFlag <> "1" Then	'�����Ƿ�Ҫ���ˮӡ���
			AddWaterFlag = "0"
		End if
		
		'�����ļ��ϴ�����,���ͼ���С
		If UpType="Field" Then
		   Dim RS:Set RS=Conn.Execute("Select FieldName,AllowFileExt,MaxFileSize From KS_Field Where FieldID=" & KS.ChkClng(UpFileObj.Form("FieldID")))
		   If Not RS.Eof Then
		    FieldName=RS(0):MaxFileSize=RS(2):AllowFileExtStr=RS(1)
		   Else
		    Response.End()
		   End IF
		   RS.Close:Set RS=Nothing
		Else
			Select Case BasicType
			   Case 0           'Ĭ���ϴ�����
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(0)  '�趨�ļ��ϴ�����ֽ���
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(0,0)
			   Case 1     '��������
				CreateThumbsFlag=true
				If UpType="Pic" Then  '��������ͼ
					MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '�趨�ļ��ϴ�����ֽ���
					AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,1)
				Else
					MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '�趨�ļ��ϴ�����ֽ���
					AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,0)
				End If
			  Case 2     'ͼƬ����
				CreateThumbsFlag=true
				If UpType="Pic" Then  '��������ͼ
					MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '�趨�ļ��ϴ�����ֽ���
					AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,1)
				Else
					MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '�趨�ļ��ϴ�����ֽ���
					AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,0)
				End If
			  Case 21     'ͼƬ�����ϴ�ͼƬ
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(2)   '�趨�ļ��ϴ�����ֽ���
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(2,1)
			  Case 3  
				If UpType="Pic" Then  '����ͼ
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '�趨�ļ��ϴ�����ֽ���
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,1)
				Else    '���������ļ�	
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '�趨�ļ��ϴ�����ֽ���
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,0)
				End If
			  Case 4   
			   If UpType="Pic" Then  '����ͼ
				 CreateThumbsFlag=true
				 MaxFileSize = KS.ReturnChannelAllowUpFilesSize(4)   '�趨�ļ��ϴ�����ֽ���
				 AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(4,1)
			   ElseIf UpType="Flash"  Then'Flash�ļ�
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(4)   '�趨�ļ��ϴ�����ֽ���
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(4,2)  'ȡ�����ϴ��Ķ�������
			   Else
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(4)   '�趨�ļ��ϴ�����ֽ���
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(4,0)
			   End If
			 Case 5     '�̳�����
			   If UpType="Pic" Then  '����ͼ
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(5)   '�趨�ļ��ϴ�����ֽ���
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(5,1)
			   Else
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(5)   '�趨�ļ��ϴ�����ֽ���
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(5,0)
			   End If
			 Case 7    'Ӱ����������ͼ
			   If UpType="Pic" Then  '����ͼ
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(7)   '�趨�ļ��ϴ�����ֽ���
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(7,1)	
			   Else
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(7)   '�趨�ļ��ϴ�����ֽ���
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(7,0)
			   End iF
			 Case 8
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(8)   '�趨�ļ��ϴ�����ֽ���
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(8,1)	
			Case 9     '����ϵͳ
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(9)   '�趨�ļ��ϴ�����ֽ���
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(9,0)
			End Select
		End If
			
		ReturnValue = CheckUpFile(FilePath,MaxFileSize,AllowFileExtStr,AutoReName)
		if ReturnValue <> "" then
		     ReturnValue = Replace(ReturnValue,"'","\'")
		     KS.AlertHintScript ReturnValue
			 Response.End()
		else 
			If UpType="Field" Then
					  Response.Write("<script language=""JavaScript"">")
					  Response.Write("parent.document.all."& FieldName & ".value='" & replace(TempFileStr,"|","") & "';")
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>��ϲ���ϴ��ɹ���</font>');")
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?ChannelID=" & ChannelID & "&UpType=Field&FieldID=" & UpFileObj.Form("FieldID") &"\'>');")
					  Response.Write("</script>")
			Else
			    TempFileStr=replace(TempFileStr,"'","\'")
				Select Case BasicType
				   Case 1         '����
					  Response.Write("<script language=""JavaScript"">")
					   if UpType="File" Then   '�ϴ�����
					      If KS.C_S(ChannelID,34)=0 Then
						  Response.Write("parent.ArticleContent.InsertFileFromUp('" & TempFileStr &"','" & KS.Setting(3) & "');")
						  Else
						  Response.Write("parent.InsertFileFromUp('" & TempFileStr &"','" & KS.Setting(3) & "');")
						  End If
						  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>�����ϴ��ɹ���</font>');")
						Else
						  if DefaultThumb=0 then
						   Response.Write("parent.document.myform.PhotoUrl.value='" &  replace(TempFileStr,"|","") & "';")
						  else
								 If KS.CheckFile(Replace(ThumbPathFileName,KS.Setting(2),""))=true Then        '����Ƿ��������ͼ
								  Response.Write("parent.document.myform.PhotoUrl.value='" & ThumbPathFileName & "';")
								 Else
								  Response.Write("parent.document.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
								 End If
						  end if
							  If Split(Split(KS.C_S(ChannelID,5),"@@@")(0),"|")(9)=1 Then
									  If KS.C_S(ChannelID,34)=0 Then
									  Response.Write("parent.ArticleContent.InsertPictureFromUp('" & replace(TempFileStr,"|","") &"');")
									  Else
									  Response.Write ("parent.insertHTMLToEditor('<img src=" & replace(TempFileStr,"|","") &" />');")
									  End IF
									 End If
									 Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>ͼƬ�ϴ��ɹ���</font>');")
							End If
		
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?ChannelID=" & ChannelID & "&UpType=" & UpType & "\'>');")
					  Response.Write("</script>")
				   Case 2          'ͼƬ
					  Response.Write("<script language=""JavaScript"">")
					  If UpType="Single" Then
					  Response.Write("parent.document.getElementById('imgurl"&UpFileObj.Form("objid")&"').value='"& replace(TempFileStr,"|","") &"';")
					  Response.Write("parent.document.getElementById('thumb"&UpFileObj.Form("objid")&"').value='"& ThumbPathFileName &"';")
					  Response.Write("document.write('<br><div align=center><font size=2>ͼƬ�ϴ��ɹ���</font></div>');")
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'1; url=../KS.UpFileForm.asp?UpType=Single&ChannelID=" & ChannelID & "\'>');")
					  Else
		
					  Response.Write("parent.SetPicUrlByUpLoad(" & DefaultThumb & ",'" & TempFileStr &  "','" & ThumbPathFileName & "|');")
					  Response.Write("document.write('<br><br><div align=center><font size=2>ͼƬ�ϴ��ɹ���</font></div>');")
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'1; url=../KS.UpFileForm.asp?ChannelID=" & ChannelID & "\'>');")
					  End If
					  Response.Write("</script>")
				  Case 3  
					  Response.Write("<script language=""JavaScript"">")
					 If UpType="Pic" Then
					  if DefaultThumb=0 then
					   Response.Write("parent.document.getElementById('PhotoUrl').value='" & replace(TempFileStr,"|","") & "';")
					   Response.Write("parent.document.getElementById('BigPhoto').value='" & replace(TempFileStr,"|","") & "';")
					  else
					   Response.Write("parent.document.getElementById('PhotoUrl').value='" & ThumbPathFileName & "';")
					   Response.Write("parent.document.getElementById('BigPhoto').value='" & replace(TempFileStr,"|","") & "';")
					  end if
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>ͼƬ�ϴ��ɹ���</font>');")
					 Else
						  Response.Write("parent.SetDownUrlByUpLoad('" & replace(TempFileStr,"|","") & "'," & U_FileSize & ");")
						  Response.Write("document.write('<br><br><div align=center><font size=2>�ļ��ϴ��ɹ���</font></div>');")
					 End If
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?ChannelID=" & ChannelID & "&UpType=" & UpType & "\'>');")
					  Response.Write("</script>")
				  Case 4         '�������ĵ��ϴ�����ͼ
					  Response.Write("<script language=""JavaScript"">")
					  
					 If UpType="Pic" Then 
					  if DefaultThumb=0 Or KS.CheckFile(Replace(ThumbPathFileName,KS.Setting(2),""))=false then  '����Ƿ��������ͼ
					   Response.Write("parent.document.all.PhotoUrl.value='" &  replace(TempFileStr,"|","") & "';")
					  else
						  Response.Write("parent.document.all.PhotoUrl.value='" & ThumbPathFileName & "';")
						  Call KS.DeleteFile(replace(TempFileStr,"|",""))  'ɾ��ԭͼƬ
					  end if
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>ͼƬ�ϴ��ɹ���</font>');")
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?ChannelID=4&UpType=Pic\'>');")
					ElseIf UpType="Flash" Then 'Flash�ļ�
					  Response.Write("parent.document.all.FlashUrl.value='" & replace(TempFileStr,"|","") & "';")
					  Response.Write("document.write('<br><br><div align=center><font size=2>�ļ��ϴ��ɹ���</font></div>');")
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'1; url=../KS.UpFileForm.asp?ChannelID=4&UpType=Flash\'>');")
					End If
					  
					  Response.Write("</script>")
				 Case 5    '�̳���������ͼ
					  Response.Write("<script language=""JavaScript"">")
					  if UpType="File" Then   '�ϴ�����
						  Response.Write("parent.InsertFileFromUp('" & TempFileStr &"');")
						  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>�����ϴ��ɹ���</font>');")
					  ElseIf UpType="ProImage" Then
						  Response.Write("parent.SetPicUrlByUpLoad('" & TempFileStr &  "','" & ThumbPathFileName & "|');")
						  Response.Write("document.write('<br><br><div align=center><font size=2>ͼƬ�ϴ��ɹ���</font></div>');")
						  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'1; url=../KS.UpFileForm.asp?UpType=ProImage&ChannelID=" & ChannelID & "\'>');")
					  Else
						  if DefaultThumb=0 then
						   Response.Write("parent.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
						   Response.Write("parent.myform.BigPhoto.value='" & replace(TempFileStr,"|","") & "';")
						  else
						   Response.Write("parent.myform.PhotoUrl.value='" & ThumbPathFileName & "';")
						   Response.Write("parent.myform.BigPhoto.value='" & replace(TempFileStr,"|","") & "';")
						  end if
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>ͼƬ�ϴ��ɹ���</font>');")
					  End If
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?ChannelID=5&UpType=" & UpType & "\'>');")
					  Response.Write("</script>")
				  Case 7    'Ӱ����������ͼ
					  Response.Write("<script language=""JavaScript"">")
					  
					  If UpType="Pic" Then 
						  if DefaultThumb=0 then
						   Response.Write("parent.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
						  else
						   Response.Write("parent.myform.PhotoUrl.value='" & ThumbPathFileName & "';")
						  end if
						  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>ͼƬ�ϴ��ɹ���</font>');")
					  Else
						  Response.Write("parent.SetMovieUrlByUpLoad('" & replace(TempFileStr,"|","") & "');")
						  Response.Write("document.write('<br><br><div align=center><font size=2>�ļ��ϴ��ɹ���</font></div>');")
					  End If	  
						  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?ChannelID=7&upType=" & UpType & "\'>');")
					  Response.Write("</script>")
				  Case 8
					  Response.Write("<script language=""JavaScript"">")
					  if DefaultThumb=0 then
					   Response.Write("parent.document.all.PhotoUrl.value='" &  replace(TempFileStr,"|","") & "';")
					  else
						 If KS.CheckFile(Replace(ThumbPathFileName,KS.Setting(2),""))=true Then        '����Ƿ��������ͼ
						  Response.Write("parent.document.all.PhotoUrl.value='" & ThumbPathFileName & "';")
						  'Call KS.DeleteFile(replace(TempFileStr,"|",""))  'ɾ��ԭͼƬ
						 Else
						  Response.Write("parent.document.all.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
						 End If
					  end if
					  If KS.C_S(ChannelID,34)=0 Then
					  Response.Write("parent.GQContent.InsertPictureFromUp('" & replace(TempFileStr,"|","") &"');")
					  Else
					  Response.Write ("parent.insertHTMLToEditor('<img src=" & replace(TempFileStr,"|","") &" />');")
					  End If
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>ͼƬ�ϴ��ɹ���</font>');")
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?Channelid=8\'>');")
					  Response.Write("</script>")		
				  Case 9
					  Response.Write("<script language=""JavaScript"">")
					  Response.Write("parent.document.all.DownUrl.value='" &  replace(TempFileStr,"|","") & "';")
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>�Ծ��ϴ��ɹ���</font>');")
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?Channelid=9\'>');")
					  Response.Write("</script>")		
				  Case else
					 if ReturnValue <> "" then
					  Response.Write("<script language=""JavaScript"">"&vbcrlf)
					  Response.Write("alert('" & ReturnValue & "');"&vbcrlf)
					  Response.Write("dialogArguments.location.reload();"&vbcrlf)
					  Response.Write("close();"&vbcrlf)
					  Response.Write("</script>"&vbcrlf)
					 else
					  Response.Write("<script language=""JavaScript"">"&vbcrlf)
					  Response.Write("dialogArguments.location.reload();"&vbcrlf)
					  Response.Write("close();"&vbcrlf)
					  Response.Write("</script>"&vbcrlf)
					 end if
				End Select
			 End If
		  End iF
		Set UpFileObj=Nothing
		End Sub
		Function CheckUpFile(Path,FileSize,AllowExtStr,AutoReName)
			Dim ErrStr,NoUpFileTF,FsoObj,FileName,FileExtName,FileContent,SameFileExistTF
			NoUpFileTF = True
			ErrStr = ""
			Set FsoObj = KS.InitialObject(FsoObjName)
			For Each FormName in UpFileObj.File
				SameFileExistTF = False
				FileName = UpFileObj.File(FormName).FileName
				FileExtName = UpFileObj.File(FormName).FileExt
				FileContent = UpFileObj.File(FormName).FileData
				U_FileSize=UpFileObj.File(FormName).FileSize
				Dim FileType:FileType=UpFileObj.File(FormName).FileType
				'�Ƿ���������ļ�
				if U_FileSize > 1 then
					NoUpFileTF = False
					ErrStr = ""
					if U_FileSize > CLng(FileSize)*1024 then
						ErrStr = ErrStr & FileName & "�ļ��ϴ�ʧ��\n���������ƣ����ֻ���ϴ�" & FileSize & "K���ļ�\n"
					end if
					if AutoRename = "0" then
						If FsoObj.FileExists(Path & FileName) = True  then
							ErrStr = ErrStr & FileName & "�ļ��ϴ�ʧ��,����ͬ���ļ�\n"
						else
							SameFileExistTF = True
						end if
					else
						SameFileExistTF = True
					End If
					if CheckFileType(AllowExtStr,FileExtName) = False then
						ErrStr = ErrStr & FileName & "�ļ��ϴ�ʧ��,�ļ����Ͳ�����\n�����������" + AllowExtStr + "\n"
					end if
					If Left(LCase(FileType), 5) = "text/" and KS.FoundInArr(NeedCheckFileMimeExt,FileExtName,"|")=true Then
					 ErrStr = ErrStr & FileName & "�ļ��ϴ�ʧ��\nΪ��ϵͳ��ȫ���������ϴ����ı��ļ�α���ͼƬ�ļ���\n"
					End If
					If instr(FileName,";")>0 or instr(lcase(FileName),".asp")>0 or instr(lcase(FileName),".php")>0 or instr(lcase(FileName),".asa")>0 or instr(lcase(FileName),".cer")>0 then
						ErrStr = ErrStr & FileName & "�ļ��ϴ�ʧ��,�ļ������Ϸ�\n"
					end if
					
					if ErrStr = "" then
						if SameFileExistTF = True then
							SaveFile Path,FormName,AutoReName
						else
							SaveFile Path,FormName,""
						end if
					else
						CheckUpFile = CheckUpFile & ErrStr
					end if
				end if
			Next
			Set FsoObj = Nothing
			if NoUpFileTF = True then
				CheckUpFile = "û���ϴ��ļ�"
			end if
		End Function
		
		Function CheckFileType(AllowExtStr,FileExtName)
			Dim i,AllowArray
			AllowArray = Split(AllowExtStr,"|")
			FileExtName = LCase(FileExtName)
			CheckFileType = False
			For i = LBound(AllowArray) to UBound(AllowArray)
				if LCase(AllowArray(i)) = LCase(FileExtName) then
					CheckFileType = True
				end if
			Next
			If KS.FoundInArr(LCase(NoAllowExt),FileExtName,"|")=true Then
				CheckFileType = False
			end if
		End Function
		
		Function SaveFile(FilePath,FormNameItem,AutoNameType)
			Dim FileName,FileExtName,FileContent,FormName,RandomFigure,n,RndStr
			Randomize 
			n=2* Rnd+10
			RndStr=KS.MakeRandom(n)
			RandomFigure = CStr(Int((99999 * Rnd) + 1))
			FileName = UpFileObj.File(FormNameItem).FileName
			FileExtName = UpFileObj.File(FormNameItem).FileExt
			FileContent = UpFileObj.File(FormNameItem).FileData
			select case AutoNameType 
			  case "1"
				FileName= "����" & FileName
			  case "2"
				FileName= RndStr&"."&FileExtName
			  Case "3"
				FileName= RndStr & FileName
			  case "4"
				FileName= Year(Now())&Right("0"&Month(Now()),2)&Right("0"&Day(Now()),2)&Right("0"&Hour(Now()),2)&Right("0"&Minute(Now()),2)&Right("0"&Second(Now()),2)&RandomFigure&"."&FileExtName
			  case else
				FileName=FileName
			End Select
		   UpFileObj.File(FormNameItem).SaveToFile FilePath  &FileName
		   
		   
		   
		   '======================���Ӽ���ļ������Ƿ�Ϸ�===================================
			call CheckFileContent(FormPath  &FileName,UpFileObj.File(FormNameItem).FileSize /1024)
			'==================================================================================
			
		   
		   TempFileStr=TempFileStr & FormPath & FileName & "|"
		  
		  CurrNum=CurrNum+1
		  IF CreateThumbsFlag=true and  (cint(CurrNum)=cint(DefaultThumb) or BasicType=2 or (Channelid=5 and UpType="ProImage")) Then
		  	  If KS.TBSetting(0)=0 then
			   if ThumbPathFileName="" then
			   ThumbPathFileName=FormPath &FileName
			   Else
			   ThumbPathFileName=ThumbPathFileName & "|" & FormPath & FileName
			   End If
			  Else
				ThumbFileName=split(FileName,".")(0)&"_S."&FileExtName
				Dim CreateTF:CreateTF=T.CreateThumbs(FilePath & FileName,FilePath & ThumbFileName)
				if CreateTF=true Then
				 'ȡ������ͼ��ַ
				 if ThumbPathFileName="" then
				 ThumbPathFileName=FormPath & ThumbFileName
				 else
				 ThumbPathFileName=ThumbPathFileName & "|" & FormPath & ThumbFileName
				end if
			   Else
				 'ȡ������ͼ��ַ
				 if ThumbPathFileName="" then
				 ThumbPathFileName=FormPath & FileName
				 else
				 ThumbPathFileName=ThumbPathFileName & "|" & FormPath & FileName
				 end if
			   End If
			  End If
		  End if
		  If AddWaterFlag = "1" Then   '�ڱ���õ�ͼƬ�����ˮӡ
				call T.AddWaterMark(FilePath  & FileName)
		   End if
		   
		   
		End Function
		
		
		'����ļ����ݵ��Ƿ�Ϸ�
		Function  CheckFileContent(byval path,byval filesize)
			        dim kk,NoAllowExtArr
					NoAllowExtArr=split(NoAllowExt,"|")
					for kk=0 to ubound(NoAllowExtArr)
					   if instr(lcase(path),"." & NoAllowExtArr(kk))<>0 then
					    call KS.DeleteFile(path)
					    ks.die  "<script>alert('�ļ��ϴ�ʧ��,�ļ������Ϸ�\n');</script>"
					   end if
					Next

		    if filesize>50 then exit function  '����1000K�������
		    on error resume next
		    Dim findcontent,regEx,foundtf
			findcontent=KS.ReadFromFile(Replace(path,KS.Setting(2),""))
			if err then exit function:err.clear
			foundtf=false
			Set regEx = New RegExp
			regEx.IgnoreCase = True
			regEx.Global = True
			regEx.Pattern = "@\s*LANGUAGE\s*=\s*[""]?\s*(vbscript|jscript|javascript).encode\b"
			If regEx.Test(findcontent) Then
				foundtf=true
			end if	
			
			regEx.Pattern = "execute\s*request"
			If regEx.Test(findcontent) Then
				foundtf=true
			end if
			
			regEx.Pattern = "executeglobal\s*request"
			If regEx.Test(findcontent) Then
				foundtf=true
			end if
			regEx.Pattern = "<script.*runat.*server(\n|.)*execute(\n|.)*<\/script>"
			If regEx.Test(findcontent) Then
				foundtf=true
			end if
			regEx.Pattern = "\<%(.|\n)*%\>"
			If regEx.Test(findcontent) Then
				foundtf=true
			end if
			If Instr(lcase(findcontent),"scripting.filesystemobject")<>0 or instr(lcase(findcontent),"adodb.stream")<>0 Then
			foundtf=true
			End If
			
			set regEx=nothing
			
			if foundtf then
			   KS.DeleteFile(path)
			   KS.Die "<script>alert('ϵͳ��鵽���ϴ����ļ����ܴ���Σ�մ��룬�������ϴ���');history.back(-1);</script>"
			end if
			
	  End Function
		
End Class
Dim UpFileStream
Class UpFileClass
	Dim Form,File,Err 
	Private Sub Class_Initialize
		Err = -1
	End Sub
	Private Sub Class_Terminate  
		'�������������
		If Err < 0 Then
			Form.RemoveAll
			Set Form = Nothing
			File.RemoveAll
			Set File = Nothing
			UpFileStream.Close
			Set UpFileStream = Nothing
		End If
	End Sub
	
	Public Property Get ErrNum()
		ErrNum = Err
	End Property
	
	Public Sub GetData ()
		'�������
		Dim RequestBinData,sSpace,bCrLf,sObj,iObjStart,iObjEnd,tStream,iStart,oFileObj
		Dim iFileSize,sFilePath,sFileType,sFormValue,sFileName
		Dim iFindStart,iFindEnd
		Dim iFormStart,iFormEnd,sFormName
		Dim KS:Set KS=New PublicCls
		'���뿪ʼ
		If Request.TotalBytes < 1 Then  '���û�������ϴ�
			Err = 1
			Exit Sub
		End If
		Set Form = KS.InitialObject ("Scripting.Dictionary")
		Form.CompareMode = 1
		Set File = KS.InitialObject ("Scripting.Dictionary")
		File.CompareMode = 1
		Set tStream = KS.InitialObject ("ADODB.Stream")
		Set UpFileStream = KS.InitialObject ("ADODB.Stream")
		UpFileStream.Type = 1
		UpFileStream.Mode = 3
		UpFileStream.Open
		UpFileStream.Write (Request.BinaryRead(Request.TotalBytes))
		UpFileStream.Position = 0
		RequestBinData=UpFileStream.Read 
		iFormEnd = UpFileStream.Size
		bCrLf = ChrB (13) & ChrB (10)
		'ȡ��ÿ����Ŀ֮��ķָ���
		sSpace=MidB (RequestBinData,1, InStrB (1,RequestBinData,bCrLf)-1)
		iStart=LenB (sSpace)
		iFormStart = iStart+2
		'�ֽ���Ŀ
		Do
			iObjEnd=InStrB(iFormStart,RequestBinData,bCrLf & bCrLf)+3
			tStream.Type = 1
			tStream.Mode = 3
			tStream.Open
			UpFileStream.Position = iFormStart
			UpFileStream.CopyTo tStream,iObjEnd-iFormStart
			tStream.Position = 0
			tStream.Type = 2
			tStream.CharSet = "gb2312"
			sObj = tStream.ReadText      
			'ȡ�ñ���Ŀ����
			iFormStart = InStrB (iObjEnd,RequestBinData,sSpace)-1
			iFindStart = InStr (22,sObj,"name=""",1)+6
			iFindEnd = InStr (iFindStart,sObj,"""",1)
			sFormName = Mid  (sObj,iFindStart,iFindEnd-iFindStart)
			'������ļ�
			If InStr  (45,sObj,"filename=""",1) > 0 Then
				Set oFileObj = new FileObj_Class
				'ȡ���ļ�����
				iFindStart = InStr (iFindEnd,sObj,"filename=""",1)+10
				iFindEnd = InStr (iFindStart,sObj,"""",1)
				sFileName = Mid (sObj,iFindStart,iFindEnd-iFindStart)
				oFileObj.FileName = Mid (sFileName,InStrRev (sFileName, "\")+1)
				oFileObj.FilePath = Left (sFileName,InStrRev (sFileName, "\"))
				oFileObj.FileExt = Mid (sFileName,InStrRev (sFileName, ".")+1)
				iFindStart = InStr (iFindEnd,sObj,"Content-Type: ",1)+14
				iFindEnd = InStr (iFindStart,sObj,vbCr)
				oFileObj.FileType = Mid  (sObj,iFindStart,iFindEnd-iFindStart)
				oFileObj.FileStart = iObjEnd
				oFileObj.FileSize = iFormStart -iObjEnd -2
				oFileObj.FormName = sFormName
				File.add sFormName,oFileObj
			else
				'����Ǳ���Ŀ
				tStream.Close
				tStream.Type = 1
				tStream.Mode = 3
				tStream.Open
				UpFileStream.Position = iObjEnd 
				UpFileStream.CopyTo tStream,iFormStart-iObjEnd-2
				tStream.Position = 0
				tStream.Type = 2
				tStream.CharSet = "gb2312"
				sFormValue = tStream.ReadText
				If Form.Exists(sFormName)Then
					Form (sFormName) = Form (sFormName) & ", " & sFormValue
				else
					form.Add sFormName,sFormValue
				End If
			End If
			tStream.Close
			iFormStart = iFormStart+iStart+2
			'������ļ�β�˾��˳�
		Loop Until  (iFormStart+2) >= iFormEnd 
		RequestBinData = ""
		Set tStream = Nothing
		Set KS=Nothing
	End Sub
End Class

'----------------------------------------------------------------------------------------------------
'�ļ�������
Class FileObj_Class
	Dim FormName,FileName,FilePath,FileSize,FileType,FileStart,FileExt
	'�����ļ�����
	Public Function SaveToFile (Path)
		On Error Resume Next
		Dim KS:Set KS=New PublicCls
		Dim oFileStream
		Set oFileStream = KS.InitialObject ("ADODB.Stream")
		oFileStream.Type = 1
		oFileStream.Mode = 3
		oFileStream.Open
		UpFileStream.Position = FileStart
		UpFileStream.CopyTo oFileStream,FileSize
		oFileStream.SaveToFile Path,2
		oFileStream.Close
		Set oFileStream = Nothing 
		Set KS=Nothing
	End Function
	'ȡ���ļ�����
	Public Function FileData
		UpFileStream.Position = FileStart

		FileData = UpFileStream.Read (FileSize)
	End Function
End Class

%> 
