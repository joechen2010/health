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

Const NoAllowExt = "asa|asax|ascs|ashx|asmx|asp|aspx|axd|cdx|cer|cfm|config|cs|csproj|idc|licx|rem|resources|resx|shtm|shtml|soap|stm|vb|vbproj|vsdisco|webinfo"    '不允许上传类型
Const NeedCheckFileMimeExt = "gif|jpg|jpeg|jpe|bmp|png|swf|mid|mp3|wmv|asf|avi|mpg|ram|rm|ra|rar|exe|doc|zip" '定义需要检查是否伪造的文件类型

Class UpFileSave
        Private KS
		Dim FilePath,MaxFileSize,AllowFileExtStr,AutoReName,RsConfigObj,BasicType,ChannelID,UpType
		Dim FormName,Path,UpLoadFrom,TempFileStr,FormPath,ThumbFileName,ThumbPathFileName
		Dim UpFileObj,FsoObjName,AddWaterFlag,T,CurrNum,CreateThumbsFlag,FieldName,	U_FileSize
		Dim DefaultThumb    '设定第几张为缩略图
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
			Response.Write "<script>alert('没有登录!');history.back();</script>"
			Response.end
		End If
		
		
		 If Trim(Request.ServerVariables("HTTP_REFERER"))="" Then
			Response.Write "<script>alert('非法上传1！');history.back();</script>"
			Response.end
		 End If
		 if instr(lcase(Request.ServerVariables("HTTP_REFERER")),"ks.upfileform.asp")<=0 and instr(lcase(Request.ServerVariables("HTTP_REFERER")),"upfileform.asp")<=0 then
			Response.Write "<script>alert('非法上传！');history.back();</script>"
			Response.end
		 end if
		 if IsSelfRefer=false Then
			Response.Write "<script>alert('请不要非法上传！');history.back();</script>"
			Response.end
		 End If
		 
		 

		FsoObjName=KS.Setting(99)
		Set UpFileObj = New UpFileClass
		UpFileObj.GetData
		FormPath=Replace(UpFileObj.Form("Path"),".","") 
		IF Instr(FormPath,KS.Setting(3))=0 Then	FormPath=KS.Setting(3) & FormPath
		Call KS.CreateListFolder(FormPath)       '生成上传文件存放目录
		FilePath=Server.MapPath(FormPath) & "\"
		If KS.Setting(97)=1 Then
		FormPath=KS.Setting(2) & FormPath
		End if
		AutoReName = KS.ChkClng(UpFileObj.Form("AutoRename"))
		UpLoadFrom=UpFileObj.Form("UpLoadFrom")        '0--通用对话框 2-- 图片中心上传 31--下载中心缩略图 32--下载中心文件 41--动漫中心缩略图 42--动漫中心的动漫文件
		IF UpLoadFrom="" then  UpLoadFrom=0
		
		BasicType=KS.ChkClng(UpFileObj.Form("BasicType"))        ' 2-- 图片中心上传 3--下载中心缩略图/文件 41--动漫中心缩略图 42--动漫中心的动漫文件
		ChannelID=KS.ChkClng(UpFileObj.Form("ChannelID")) 
		UpType=UpFileObj.Form("UpType")
		
		CurrNum=0
		CreateThumbsFlag=false
		DefaultThumb=UpFileObj.Form("DefaultUrl")
		if DefaultThumb="" then DefaultThumb=0
		AddWaterFlag = UpFileObj.Form("AddWaterFlag")
		If AddWaterFlag <> "1" Then	'生成是否要添加水印标记
			AddWaterFlag = "0"
		End if
		
		'设置文件上传限制,类型及大小
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
			   Case 0           '默认上传参数
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(0)  '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(0,0)
			   Case 1     '文章中心
				CreateThumbsFlag=true
				If UpType="Pic" Then  '文章缩略图
					MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
					AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,1)
				Else
					MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
					AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,0)
				End If
			  Case 2     '图片中心
				CreateThumbsFlag=true
				If UpType="Pic" Then  '文章缩略图
					MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
					AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,1)
				Else
					MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
					AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,0)
				End If
			  Case 21     '图片中心上传图片
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(2)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(2,1)
			  Case 3  
				If UpType="Pic" Then  '缩略图
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,1)
				Else    '下载中心文件	
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,0)
				End If
			  Case 4   
			   If UpType="Pic" Then  '缩略图
				 CreateThumbsFlag=true
				 MaxFileSize = KS.ReturnChannelAllowUpFilesSize(4)   '设定文件上传最大字节数
				 AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(4,1)
			   ElseIf UpType="Flash"  Then'Flash文件
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(4)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(4,2)  '取允许上传的动漫类型
			   Else
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(4)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(4,0)
			   End If
			 Case 5     '商城中心
			   If UpType="Pic" Then  '缩略图
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(5)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(5,1)
			   Else
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(5)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(5,0)
			   End If
			 Case 7    '影视中心缩略图
			   If UpType="Pic" Then  '缩略图
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(7)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(7,1)	
			   Else
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(7)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(7,0)
			   End iF
			 Case 8
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(8)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(8,1)	
			Case 9     '考试系统
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(9)   '设定文件上传最大字节数
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
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>恭喜，上传成功！</font>');")
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?ChannelID=" & ChannelID & "&UpType=Field&FieldID=" & UpFileObj.Form("FieldID") &"\'>');")
					  Response.Write("</script>")
			Else
			    TempFileStr=replace(TempFileStr,"'","\'")
				Select Case BasicType
				   Case 1         '文章
					  Response.Write("<script language=""JavaScript"">")
					   if UpType="File" Then   '上传附件
					      If KS.C_S(ChannelID,34)=0 Then
						  Response.Write("parent.ArticleContent.InsertFileFromUp('" & TempFileStr &"','" & KS.Setting(3) & "');")
						  Else
						  Response.Write("parent.InsertFileFromUp('" & TempFileStr &"','" & KS.Setting(3) & "');")
						  End If
						  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>附件上传成功！</font>');")
						Else
						  if DefaultThumb=0 then
						   Response.Write("parent.document.myform.PhotoUrl.value='" &  replace(TempFileStr,"|","") & "';")
						  else
								 If KS.CheckFile(Replace(ThumbPathFileName,KS.Setting(2),""))=true Then        '检查是否存在缩略图
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
									 Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>图片上传成功！</font>');")
							End If
		
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?ChannelID=" & ChannelID & "&UpType=" & UpType & "\'>');")
					  Response.Write("</script>")
				   Case 2          '图片
					  Response.Write("<script language=""JavaScript"">")
					  If UpType="Single" Then
					  Response.Write("parent.document.getElementById('imgurl"&UpFileObj.Form("objid")&"').value='"& replace(TempFileStr,"|","") &"';")
					  Response.Write("parent.document.getElementById('thumb"&UpFileObj.Form("objid")&"').value='"& ThumbPathFileName &"';")
					  Response.Write("document.write('<br><div align=center><font size=2>图片上传成功！</font></div>');")
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'1; url=../KS.UpFileForm.asp?UpType=Single&ChannelID=" & ChannelID & "\'>');")
					  Else
		
					  Response.Write("parent.SetPicUrlByUpLoad(" & DefaultThumb & ",'" & TempFileStr &  "','" & ThumbPathFileName & "|');")
					  Response.Write("document.write('<br><br><div align=center><font size=2>图片上传成功！</font></div>');")
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
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>图片上传成功！</font>');")
					 Else
						  Response.Write("parent.SetDownUrlByUpLoad('" & replace(TempFileStr,"|","") & "'," & U_FileSize & ");")
						  Response.Write("document.write('<br><br><div align=center><font size=2>文件上传成功！</font></div>');")
					 End If
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?ChannelID=" & ChannelID & "&UpType=" & UpType & "\'>');")
					  Response.Write("</script>")
				  Case 4         '动漫中心的上传缩略图
					  Response.Write("<script language=""JavaScript"">")
					  
					 If UpType="Pic" Then 
					  if DefaultThumb=0 Or KS.CheckFile(Replace(ThumbPathFileName,KS.Setting(2),""))=false then  '检查是否存在缩略图
					   Response.Write("parent.document.all.PhotoUrl.value='" &  replace(TempFileStr,"|","") & "';")
					  else
						  Response.Write("parent.document.all.PhotoUrl.value='" & ThumbPathFileName & "';")
						  Call KS.DeleteFile(replace(TempFileStr,"|",""))  '删除原图片
					  end if
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>图片上传成功！</font>');")
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?ChannelID=4&UpType=Pic\'>');")
					ElseIf UpType="Flash" Then 'Flash文件
					  Response.Write("parent.document.all.FlashUrl.value='" & replace(TempFileStr,"|","") & "';")
					  Response.Write("document.write('<br><br><div align=center><font size=2>文件上传成功！</font></div>');")
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'1; url=../KS.UpFileForm.asp?ChannelID=4&UpType=Flash\'>');")
					End If
					  
					  Response.Write("</script>")
				 Case 5    '商城中心缩略图
					  Response.Write("<script language=""JavaScript"">")
					  if UpType="File" Then   '上传附件
						  Response.Write("parent.InsertFileFromUp('" & TempFileStr &"');")
						  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>附件上传成功！</font>');")
					  ElseIf UpType="ProImage" Then
						  Response.Write("parent.SetPicUrlByUpLoad('" & TempFileStr &  "','" & ThumbPathFileName & "|');")
						  Response.Write("document.write('<br><br><div align=center><font size=2>图片上传成功！</font></div>');")
						  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'1; url=../KS.UpFileForm.asp?UpType=ProImage&ChannelID=" & ChannelID & "\'>');")
					  Else
						  if DefaultThumb=0 then
						   Response.Write("parent.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
						   Response.Write("parent.myform.BigPhoto.value='" & replace(TempFileStr,"|","") & "';")
						  else
						   Response.Write("parent.myform.PhotoUrl.value='" & ThumbPathFileName & "';")
						   Response.Write("parent.myform.BigPhoto.value='" & replace(TempFileStr,"|","") & "';")
						  end if
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>图片上传成功！</font>');")
					  End If
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?ChannelID=5&UpType=" & UpType & "\'>');")
					  Response.Write("</script>")
				  Case 7    '影视中心缩略图
					  Response.Write("<script language=""JavaScript"">")
					  
					  If UpType="Pic" Then 
						  if DefaultThumb=0 then
						   Response.Write("parent.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
						  else
						   Response.Write("parent.myform.PhotoUrl.value='" & ThumbPathFileName & "';")
						  end if
						  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>图片上传成功！</font>');")
					  Else
						  Response.Write("parent.SetMovieUrlByUpLoad('" & replace(TempFileStr,"|","") & "');")
						  Response.Write("document.write('<br><br><div align=center><font size=2>文件上传成功！</font></div>');")
					  End If	  
						  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?ChannelID=7&upType=" & UpType & "\'>');")
					  Response.Write("</script>")
				  Case 8
					  Response.Write("<script language=""JavaScript"">")
					  if DefaultThumb=0 then
					   Response.Write("parent.document.all.PhotoUrl.value='" &  replace(TempFileStr,"|","") & "';")
					  else
						 If KS.CheckFile(Replace(ThumbPathFileName,KS.Setting(2),""))=true Then        '检查是否存在缩略图
						  Response.Write("parent.document.all.PhotoUrl.value='" & ThumbPathFileName & "';")
						  'Call KS.DeleteFile(replace(TempFileStr,"|",""))  '删除原图片
						 Else
						  Response.Write("parent.document.all.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
						 End If
					  end if
					  If KS.C_S(ChannelID,34)=0 Then
					  Response.Write("parent.GQContent.InsertPictureFromUp('" & replace(TempFileStr,"|","") &"');")
					  Else
					  Response.Write ("parent.insertHTMLToEditor('<img src=" & replace(TempFileStr,"|","") &" />');")
					  End If
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>图片上传成功！</font>');")
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=../KS.UpFileForm.asp?Channelid=8\'>');")
					  Response.Write("</script>")		
				  Case 9
					  Response.Write("<script language=""JavaScript"">")
					  Response.Write("parent.document.all.DownUrl.value='" &  replace(TempFileStr,"|","") & "';")
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>试卷上传成功！</font>');")
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
				'是否存在重名文件
				if U_FileSize > 1 then
					NoUpFileTF = False
					ErrStr = ""
					if U_FileSize > CLng(FileSize)*1024 then
						ErrStr = ErrStr & FileName & "文件上传失败\n超过了限制，最大只能上传" & FileSize & "K的文件\n"
					end if
					if AutoRename = "0" then
						If FsoObj.FileExists(Path & FileName) = True  then
							ErrStr = ErrStr & FileName & "文件上传失败,存在同名文件\n"
						else
							SameFileExistTF = True
						end if
					else
						SameFileExistTF = True
					End If
					if CheckFileType(AllowExtStr,FileExtName) = False then
						ErrStr = ErrStr & FileName & "文件上传失败,文件类型不允许\n允许的类型有" + AllowExtStr + "\n"
					end if
					If Left(LCase(FileType), 5) = "text/" and KS.FoundInArr(NeedCheckFileMimeExt,FileExtName,"|")=true Then
					 ErrStr = ErrStr & FileName & "文件上传失败\n为了系统安全，不允许上传用文本文件伪造的图片文件！\n"
					End If
					If instr(FileName,";")>0 or instr(lcase(FileName),".asp")>0 or instr(lcase(FileName),".php")>0 or instr(lcase(FileName),".asa")>0 or instr(lcase(FileName),".cer")>0 then
						ErrStr = ErrStr & FileName & "文件上传失败,文件名不合法\n"
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
				CheckUpFile = "没有上传文件"
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
				FileName= "副件" & FileName
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
		   
		   
		   
		   '======================增加检查文件内容是否合法===================================
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
				 '取得缩略图地址
				 if ThumbPathFileName="" then
				 ThumbPathFileName=FormPath & ThumbFileName
				 else
				 ThumbPathFileName=ThumbPathFileName & "|" & FormPath & ThumbFileName
				end if
			   Else
				 '取得缩略图地址
				 if ThumbPathFileName="" then
				 ThumbPathFileName=FormPath & FileName
				 else
				 ThumbPathFileName=ThumbPathFileName & "|" & FormPath & FileName
				 end if
			   End If
			  End If
		  End if
		  If AddWaterFlag = "1" Then   '在保存好的图片上添加水印
				call T.AddWaterMark(FilePath  & FileName)
		   End if
		   
		   
		End Function
		
		
		'检查文件内容的是否合法
		Function  CheckFileContent(byval path,byval filesize)
			        dim kk,NoAllowExtArr
					NoAllowExtArr=split(NoAllowExt,"|")
					for kk=0 to ubound(NoAllowExtArr)
					   if instr(lcase(path),"." & NoAllowExtArr(kk))<>0 then
					    call KS.DeleteFile(path)
					    ks.die  "<script>alert('文件上传失败,文件名不合法\n');</script>"
					   end if
					Next

		    if filesize>50 then exit function  '超过1000K跳过检测
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
			   KS.Die "<script>alert('系统检查到您上传的文件可能存在危险代码，不允许上传！');history.back(-1);</script>"
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
		'清除变量及对像
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
		'定义变量
		Dim RequestBinData,sSpace,bCrLf,sObj,iObjStart,iObjEnd,tStream,iStart,oFileObj
		Dim iFileSize,sFilePath,sFileType,sFormValue,sFileName
		Dim iFindStart,iFindEnd
		Dim iFormStart,iFormEnd,sFormName
		Dim KS:Set KS=New PublicCls
		'代码开始
		If Request.TotalBytes < 1 Then  '如果没有数据上传
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
		'取得每个项目之间的分隔符
		sSpace=MidB (RequestBinData,1, InStrB (1,RequestBinData,bCrLf)-1)
		iStart=LenB (sSpace)
		iFormStart = iStart+2
		'分解项目
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
			'取得表单项目名称
			iFormStart = InStrB (iObjEnd,RequestBinData,sSpace)-1
			iFindStart = InStr (22,sObj,"name=""",1)+6
			iFindEnd = InStr (iFindStart,sObj,"""",1)
			sFormName = Mid  (sObj,iFindStart,iFindEnd-iFindStart)
			'如果是文件
			If InStr  (45,sObj,"filename=""",1) > 0 Then
				Set oFileObj = new FileObj_Class
				'取得文件属性
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
				'如果是表单项目
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
			'如果到文件尾了就退出
		Loop Until  (iFormStart+2) >= iFormEnd 
		RequestBinData = ""
		Set tStream = Nothing
		Set KS=Nothing
	End Sub
End Class

'----------------------------------------------------------------------------------------------------
'文件属性类
Class FileObj_Class
	Dim FormName,FileName,FilePath,FileSize,FileType,FileStart,FileExt
	'保存文件方法
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
	'取得文件数据
	Public Function FileData
		UpFileStream.Position = FileStart

		FileData = UpFileStream.Read (FileSize)
	End Function
End Class

%> 
