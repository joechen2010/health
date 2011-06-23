<%
Const NeedCheckFileMimeExt = "gif|jpg|jpeg|jpe|bmp|png|swf|mid|mp3|wmv|asf|avi|mpg|ram|rm|ra|rar|exe|doc|zip" '定义需要检查是否伪造的文件类型
Class UpFileSave
        Private KS
		Dim FilePath,MaxFileSize,AllowFileExtStr,AutoReName,RsConfigObj
		Dim FormName,Path,BasicType,ChannelID,UpType,TempFileStr,FormPath,ThumbFileName,ThumbPathFileName
		Dim UpFileObj,FsoObjName,AddWaterFlag,T,CurrNum,CreateThumbsFlag,FieldName
		Dim DefaultThumb    '设定第几张为缩略图
		Dim ReturnValue
		Private Sub Class_Initialize()
		    Set T=New Thumb
		    Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		    'Call CloseConn()
		    Set T=Nothing
		    Set KS=Nothing
		    'Set KSUser=Nothing
		End Sub
		
		Public Function UpFileUrl()
		    'On Error Resume Next
		    IF Cbool(KSUser.UserLoginChecked)=false Then
			   Response.redirect KS.GetDomain&"User/Login/"
			   Exit Function
			End If
			
		 If Trim(Request.ServerVariables("HTTP_REFERER"))="" Then
			   Response.Write "上传失败,请不要非法提交！<br/>"
			   Response.Write "<anchor><prev/>返回上页</anchor><br/>"
			   Response.Write "<a href=""Index.asp?" & KS.WapValue & """>我的地盘</a>"
			   Response.Write "</p></card></wml>"
			   Response.End
		 End If

			
			IF KS.GetFolderSize(KSUser.GetUserFolder(KSUser.UserName))/1024>=KS.ChkClng(KS.Setting(50)) Then
			   Response.Write "上传失败，您的可用空间不够！<br/>"
			   Response.Write "<anchor><prev/>返回上页</anchor><br/>"
			   Response.Write "<a href=""Index.asp?" & KS.WapValue & """>我的地盘</a>"
			   Response.Write "</p></card></wml>"
			   Response.End
			End If
			
			FsoObjName=KS.Setting(99)
			Set UpFileObj = New UpFileClass
			UpFileObj.GetData
			
			AutoReName = UpFileObj.Form("AutoRename")
			BasicType=KS.ChkClng(UpFileObj.Form("BasicType"))        ' 2-- 图片中心上传 3--下载中心缩略图/文件 41--动漫中心缩略图 42--动漫中心的动漫文件
			ChannelID=KS.ChkClng(UpFileObj.Form("ChannelID")) 
			UpType=UpFileObj.Form("Type")
			IF BasicType=0 And UpType<>"Field" then 
			   Response.Write "请不要非法上传！<br/>"
			   Response.Write "<anchor><prev/>还回上页</anchor><br/>"
			   Response.Write "<a href=""Index.asp?" & KS.WapValue & """>我的地盘</a>"
			   Response.Write "</p></card></wml>"
			   Response.End
			End If
			CurrNum=0
			CreateThumbsFlag=false
			DefaultThumb=UpFileObj.Form("DefaultUrl")
			If DefaultThumb="" Then DefaultThumb=0
			AddWaterFlag = UpFileObj.Form("AddWaterFlag")
			If AddWaterFlag <> "1" Then	'生成是否要添加水印标记
			   AddWaterFlag = "0"
			End if
			
			'设置文件上传限制,类型及大小
			If UpType="Field" Then
			   Dim RS
			   If ChannelID=0 Then
			      Set RS=Conn.Execute("Select FieldName,AllowFileExt,MaxFileSize From KS_FormField Where FieldID=" & KS.ChkClng(UpFileObj.Form("FieldID")))
			   Else
			      Set RS=Conn.Execute("Select FieldName,AllowFileExt,MaxFileSize From KS_Field Where FieldID=" & KS.ChkClng(UpFileObj.Form("FieldID")))
			   End if
			   If Not RS.Eof Then
			      FieldName=RS(0):MaxFileSize=RS(2):AllowFileExtStr=RS(1)
				  FormPath = KS.ReturnChannelUserUpFilesDir(ChannelID,KSUser.UserName) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
			   Else
			      Response.Write "请不要非法上传！<br/>"
				  Response.Write "<a href=""Index.asp?" & KS.WapValue & """>我的地盘</a>"
				  Response.Write "</p></card></wml>"
				  Response.End
			   End IF
			   RS.Close:Set RS=Nothing
			Else
			   Select Case BasicType
			       Case 1     '文章中心缩略图
				      If Not KS.ReturnChannelAllowUserUpFilesTF(ChannelID) Then
					     Response.Write "对不起，系统不允许此频道上传文件,请与网站管理员联系!<br/>"
						 Exit Function
					  End IF
					  CreateThumbsFlag=True
					  MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
					  If UpType="File" Then '附件
					     AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,0)
					     FormPath = KS.ReturnChannelUserUpFilesDir(ChannelID,KSUser.UserName) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
					  Else
					     AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,1)
					     FormPath = KS.ReturnChannelUserUpFilesDir(ChannelID,KSUser.UserName) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
					  End If
				   Case 2     '图片中心上传图片
				      If Not KS.ReturnChannelAllowUserUpFilesTF(ChannelID) Then
					     Response.Write "对不起，系统不允许此频道上传文件,请与网站管理员联系!<br/>"
						 Exit Function
					  End IF
					  CreateThumbsFlag=True
					  MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
					  AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,1)
					  FormPath = KS.ReturnChannelUserUpFilesDir(ChannelID,KSUser.UserName) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
				   Case 3    
				      If Not KS.ReturnChannelAllowUserUpFilesTF(ChannelID) Then
					     Response.Write "对不起，系统不允许此频道上传文件,请与网站管理员联系!<br/>"
					     Exit Function
					  End IF
					  CreateThumbsFlag=True
					  MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
					  If UpType="Pic" Then '下载中心缩略图
					     AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,1)
						 FormPath = KS.ReturnChannelUserUpFilesDir(ChannelID,KSUser.UserName)& Year(Now()) & Right("0" & Month(Now()), 2) & "/"
					  Else
					     AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,0)
						 FormPath = KS.ReturnChannelUserUpFilesDir(ChannelID,KSUser.UserName) & "DownUrl/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
					  End If
				   Case 4
				      If Not KS.ReturnChannelAllowUserUpFilesTF(4) Then
					     Response.Write "对不起，系统不允许此频道上传文件,请与网站管理员联系!<br/>"
					     Exit Function
					  End IF
					  CreateThumbsFlag=True
					  MaxFileSize = KS.ReturnChannelAllowUpFilesSize(4)   '设定文件上传最大字节数
					  If UpType="Pic" Then '动漫中心缩略图
					     AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(4,1)
						 FormPath = KS.ReturnChannelUserUpFilesDir(4,KSUser.UserName) & "FlashPhoto/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
					  Else
					     AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(4,2)  '取允许上传的动漫类型
					     FormPath = KS.ReturnChannelUserUpFilesDir(4,KSUser.UserName) & "FlashUrl/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
					  End If
				   Case 7
				      If Not KS.ReturnChannelAllowUserUpFilesTF(7) Then
					     Response.Write "对不起，系统不允许此频道上传文件,请与网站管理员联系!<br/>"
					     Exit Function
					  End IF
					  CreateThumbsFlag=True
					  MaxFileSize = KS.ReturnChannelAllowUpFilesSize(7)   '设定文件上传最大字节数
					  If UpType="Pic" Then '影片缩略图
					     AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(7,1)
					     FormPath = KS.ReturnChannelUserUpFilesDir(7,KSUser.UserName) & "MoviePhoto/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
					  Else
					     AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(7,2)  '取允许上传的动漫类型
					     FormPath = KS.ReturnChannelUserUpFilesDir(7,KSUser.UserName) & "MovieUrl/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
					  End If
				   Case 8      '供求中心图片
				      If Not KS.ReturnChannelAllowUserUpFilesTF(8) Then
					     Response.Write "对不起，系统不允许此频道上传文件,请与网站管理员联系!<br/>"
					     Exit Function
					  End IF
					  CreateThumbsFlag=True
					  MaxFileSize = KS.ReturnChannelAllowUpFilesSize(8)   '设定文件上传最大字节数
					  AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(8,1)
					  FormPath = KS.ReturnChannelUserUpFilesDir(8,KSUser.UserName) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
				   Case 9999   '用户头像
				      MaxFileSize = 50    '设定文件上传最大字节数
					  AllowFileExtStr = "jpg|gif|png"  '取允许上传的动漫类型
					  FormPath = KS.ReturnChannelUserUpFilesDir(9999,KSUser.UserName)
				   Case 9998   '相册封面 
				      MaxFileSize = 50    '设定文件上传最大字节数
					  AllowFileExtStr = "jpg|gif|png"  '取允许上传的动漫类型
					  FormPath = KS.ReturnChannelUserUpFilesDir(9998,KSUser.UserName)
				   Case 9997 '相片　
				      MaxFileSize = 100    '设定文件上传最大字节数
					  AllowFileExtStr = "jpg|gif|png"  '取允许上传的动漫类型
					  FormPath = KS.ReturnChannelUserUpFilesDir(9997,KSUser.UserName) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
				   Case 9996 '圈子图片　
				      MaxFileSize = 50    '设定文件上传最大字节数
					  AllowFileExtStr = "jpg|gif|png"  '取允许上传的动漫类型
					  FormPath =KS.ReturnChannelUserUpFilesDir(9996,KSUser.UserName)
				   Case 9995  '音乐
				      MaxFileSize = 50000    '设定文件上传最大字节数
					  AllowFileExtStr = "mp3"  '取允许上传的动漫类型
					  FormPath =KS.ReturnChannelUserUpFilesDir(9995,KSUser.UserName)
				   Case 999  '上传中心
				      MaxFileSize = 100    '设定文件上传最大字节数
					  AllowFileExtStr = "jpg|gif|png|swf"  '取允许上传的类型
					  FormPath =KS.ReturnChannelUserUpFilesDir(999,KSUser.UserName)
				   Case Else
				   MaxFileSize=0:AllowFileExtStr=""
			       Response.Write "请不要非法上传！<br/>"
				   Response.Write "<a href=""Index.asp?" & KS.WapValue & """>我的地盘</a>"
				   Response.Write "</p></card></wml>"
				   Response.End
			   End Select
		    End If
			
			FormPath=Replace(FormPath,".","")
			IF Instr(FormPath,KS.Setting(3))=0 Then FormPath=KS.Setting(3) & FormPath
			FilePath=GetMapPath &FormPath & "\"
			
			Call KS.CreateListFolder(FormPath)       '生成上传文件存放目录
			
			ReturnValue = CheckUpFile(FilePath,MaxFileSize,AllowFileExtStr,AutoReName)
			If ReturnValue <> "" Then
			   Response.Write "" & ReturnValue & ""
			   Response.Write "<anchor><prev/>还回上页</anchor><br/>"
			   Response.Write "<a href=""Index.asp?" & KS.WapValue & """>我的地盘</a>"
			   Response.Write "</p></card></wml>"
			   Response.End()
			Else  
               If UpType="Field" Then
			      UpFileUrl = FieldName & "|" & Replace(TempFileStr,"|","")
				  Response.Write "恭喜，上传成功！<br/>"
				  Exit Function
			   End If
			   Select Case BasicType
			       Case 1         '文章中心的上传缩略图
				      If UpType="File" Then   '上传附件
						 UpFileUrl = TempFileStr
						 Response.Write "附件上传成功！<br/>"
					  Else
					     If DefaultThumb=0 Then
					        UpFileUrl = Replace(TempFileStr,"|","")
						 Else
						    If KS.CheckFile(ThumbPathFileName)=True Then        '检查是否存在缩略图
							   UpFileUrl = ThumbPathFileName
							   'Call KS.DeleteFile(replace(TempFileStr,"|",""))  '删除原图片
							Else
							   UpFileUrl = Replace(TempFileStr,"|","")
							End If
					     End If 
						 If Split(Split(KS.C_S(ChannelID,5),"@@@")(0),"|")(9)=1 Then
						    UpFileUrl = "<img src=" & Replace(TempFileStr,"|","") &" />"
						 End If
						 Response.Write "图片上传成功！<br/>"
				      End If 
					  'Response.Write "User_UpFile.asp?Channelid=" & ChannelID & "&Type=" & UpType & ""
			       Case 2          '图片中心的上传图片
				      If UPType="Single" Then
					     Response.Write "parent.document.all.imgurl"&UpFileObj.Form("objid")&".value='"& replace(TempFileStr,"|","") &"';"
					     Response.Write "parent.document.all.thumb"&UpFileObj.Form("objid")&".value='"& ThumbPathFileName &"';"
					     Response.Write "图片上传成功！<br/>"
					  Else
					     Response.Write("parent.SetPicUrlByUpLoad(" & DefaultThumb & ",'" & TempFileStr &  "','" & ThumbPathFileName & "|');")
						 Response.Write "图片上传成功！<br/>"
					  End If
				   Case 3    '下载中心缩略图
				      If UPType="Pic" Then
					     UpFileUrl = Replace(TempFileStr,"|","")
					     Response.Write "图片上传成功！<br/>"
					  Else   '下载中心的文件
					     UpFileUrl = TempFileStr
						 Response.Write "文件上传成功！<br/>"
					  End If
				   Case 4         '动漫中心的上传缩略图
				      If UpType="Pic" Then
					     UpFileUrl = Replace(TempFileStr,"|","")
						 Response.Write "图片上传成功！<br/>"
					  Else
					     UpFileUrl = Replace(TempFileStr,"|","")
						 Response.Write "文件上传成功！<br/>"
					  End If
				   Case 7         '影片中心的上传缩略图
				      If UpType="Pic" Then
					     UpFileUrl = Replace(TempFileStr,"|","")
						 Response.Write "图片上传成功！<br/>"
					  Else
					     UpFileUrl = Replace(TempFileStr,"|","")
					     Response.Write "文件上传成功！<br/>"
					  End If
			       Case 8         '供求中心的上传缩略图
				      If DefaultThumb=0 Then
					     UpFileUrl = Replace(TempFileStr,"|","")
					  Else
					     If KS.CheckFile(ThumbPathFileName)=True Then        '检查是否存在缩略图
						    UpFileUrl = ThumbPathFileName
						    'Call KS.DeleteFile(replace(TempFileStr,"|",""))  '删除原图片
						 Else
						    UpFileUrl = Replace(TempFileStr,"|","")
						 End If
					  End If
					  Response.Write "图片上传成功！<br/>"
				   Case 9999        '用户头像
					  UpFileUrl = Replace(TempFileStr,"|","")
					  Response.Write "图片上传成功！<br/>"
				   Case 9998        '相册封面
					  UpFileUrl = Replace(TempFileStr,"|","")
					  Response.Write "图片上传成功！<br/>"
				   Case 9997        '相片
				      Dim I,TempFileArr
					  TempFileStr=Left(tempfilestr,len(tempfilestr)-1)
					  Response.Write("parent.myform.PhotoUrls.value='" & TempFileStr & "';")
					  TempFileArr=split(TempFileStr,"|")
					  For I=Lbound(TempFileArr) to Ubound(TempFileArr)
					      Response.Write("parent.myform.view" & I+1 & ".src='" & TempFileArr(i) & "';")
						  Response.Write("parent.myform.view" & I+1 & ".width=83;")
						  Response.Write("parent.myform.view" & I+1 & ".height=100;")
					  Next
					  Response.write "恭喜您，照片上传成功！请按发布按钮进行保存。<br/>"
			       Case 9996        '圈子图片
					  UpFileUrl = Replace(TempFileStr,"|","")
					  Response.Write "图片上传成功！<br/>"
				   Case 9995        '用户头像
				      UpFileUrl = KS.Setting(3) & Right(Replace(TempFileStr,"|",""),Len(Replace(TempFileStr,"|",""))-1)
					  Response.Write "歌曲上传成功！<br/>"
				   Case 999
				   'Response.Write("parent.location.href='selectphoto.asp?channelid=999';"&vbcrlf)
				   Case else
				   Response.End()
			   End Select
			End If
			Set UpFileObj=Nothing
		End Function
		
		
		Function CheckUpFile(Path,FileSize,AllowExtStr,AutoReName)
			Dim ErrStr,NoUpFileTF,FsoObj,FileName,FileExtName,FileContent,SameFileExistTF
			NoUpFileTF = True
			ErrStr = ""
			Set FsoObj = KS.InitialObject(FsoObjName)
			For Each FormName In UpFileObj.File
				SameFileExistTF = False
				FileName = UpFileObj.File(FormName).FileName
				If NoIllegalStr(FileName)=False Then ErrStr=ErrStr&"文件：上传被禁止！<br/>"
				FileExtName = UpFileObj.File(FormName).FileExt
				FileContent = UpFileObj.File(FormName).FileData
				Dim FileType:FileType=UpFileObj.File(FormName).FileType
				'是否存在重名文件
				If UpFileObj.File(FormName).FileSize > 1 Then
					NoUpFileTF = False
					ErrStr = ""
					If UpFileObj.File(FormName).FileSize > CLng(FileSize)*1024 Then
						ErrStr = ErrStr & FileName & "文件上传失败<br/>超过了限制，最大只能上传" & FileSize & "K的文件<br/>"
					End If
					IF KS.ChkClng(KS.GetFolderSize(KSUser.GetUserFolder(KSUser.UserName))/1024+UpFileObj.File(FormName).FileSize/1024)>=KS.ChkClng(KS.Setting(50)) Then
					   Response.Write "上传失败1，您的可用空间不够！<br/>"
					   Response.Write "<a href=""Index.asp?" & KS.WapValue & """>我的地盘</a>"
					   Response.Write "</p></card></wml>"
					   Response.End
					End If
					If AutoRename = "0" then
					   If FsoObj.FileExists(Path & FileName) = True  Then
					      ErrStr = ErrStr & FileName & "文件上传失败,存在同名文件<br/>"
					   Else
					      SameFileExistTF = True
					   End If
					Else
					   SameFileExistTF = True
					End If
					If CheckFileType(AllowExtStr,FileExtName) = False Then
					   ErrStr = ErrStr & FileName & "文件上传失败,文件类型不允许<br/>允许的类型有" + AllowExtStr + "<br/>"
					End If
					
					If Left(LCase(FileType), 5) = "text/" and KS.FoundInArr(NeedCheckFileMimeExt,FileExtName,"|")=true Then
					 ErrStr = ErrStr & FileName & "文件上传失败\n为了系统安全，不允许上传用文本文件伪造的图片文件！\n"
					End If
					If instr(FileName,";")>0 or instr(lcase(FileName),".asp")>0 or instr(lcase(FileName),".php")>0 then
						ErrStr = ErrStr & FileName & "文件上传失败,文件名不合法\n"
					end if

					
					If ErrStr = "" Then
					   If SameFileExistTF = True Then
					      SaveFile Path,FormName,AutoReName
					   Else
					      SaveFile Path,FormName,""
					   End If
					Else
						CheckUpFile = CheckUpFile & ErrStr
					End If
				 End If
			Next
			Set FsoObj = Nothing
			If NoUpFileTF = True Then
			   CheckUpFile = "没有上传文件<br/>"
			End If
		End Function
		Function NoIllegalStr(Byval FileNameStr)
			Dim Str_Len,Str_Pos
			Str_Len=Len(FileNameStr)
			Str_Pos=InStr(FileNameStr,Chr(0))
			If Str_Pos=0 or Str_Pos=Str_Len then
				NoIllegalStr=True
			Else
				NoIllegalStr=False
			End If
		End function
		Function DealExtName(Byval UpFileExt)
			If IsEmpty(UpFileExt) Then Exit Function
			DealExtName = Lcase(UpFileExt)
			DealExtName = Replace(DealExtName,Chr(0),"")
			DealExtName = Replace(DealExtName,".","")
			DealExtName = Replace(DealExtName,"'","")
			DealExtName = Replace(DealExtName,"asp","")
			DealExtName = Replace(DealExtName,"asa","")
			DealExtName = Replace(DealExtName,"aspx","")
			DealExtName = Replace(DealExtName,"cer","")
			DealExtName = Replace(DealExtName,"cdx","")
			DealExtName = Replace(DealExtName,"htr","")
			DealExtName = Replace(DealExtName,"php","")
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
			if FileExtName="asp" or FileExtName="asa" or FileExtName="aspx" or  FileExtName="php" or  FileExtName="php3" or  FileExtName="php4"  or  FileExtName="php5" then
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
			FileExtName = DealExtName(FileExtName)
			FileContent = UpFileObj.File(FormNameItem).FileData
			Select Case AutoNameType 
			    Case "1"
				FileName= "副件" & FileName
				Case "2"
				FileName= RndStr&"."&FileExtName
				Case "3"
				FileName= RndStr & FileName
				Case "4"
				FileName= Year(Now())&Right("0"&Month(Now()),2)&Right("0"&Day(Now()),2)&Right("0"&Hour(Now()),2)&Right("0"&Minute(Now()),2)&Right("0"&Second(Now()),2)&RandomFigure&"."&FileExtName
				Case Else
				FileName=FileName
			End Select

			UpFileObj.File(FormNameItem).SaveToFile FilePath &FileName
			TempFileStr=TempFileStr & FormPath & FileName & "|"
			If AddWaterFlag = "1" Then   '在保存好的图片上添加水印
			   Call T.AddWaterMark(FilePath  & FileName)
		    End if
			CurrNum=CurrNum+1
			
			IF CreateThumbsFlag=True And (Cint(CurrNum)=Cint(DefaultThumb) or BasicType=2) Then
			   If KS.TBSetting(0)=0 Then
			      If ThumbPathFileName="" Then
				     ThumbPathFileName=FormPath &FileName
				  Else
				     ThumbPathFileName=ThumbPathFileName & "|" & FormPath & FileName
				  End If
			   Else
				  ThumbFileName=split(FileName,".")(0)&"_S."&FileExtName
				  Call T.CreateThumbs(FilePath & FileName,FilePath & ThumbFileName)
				  '取得缩略图地址
				  If ThumbPathFileName="" Then
				     ThumbPathFileName=FormPath & ThumbFileName
				  Else
				     ThumbPathFileName=ThumbPathFileName & "|" & FormPath & ThumbFileName
				  End If
			   End If
			End if
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
	End Sub
End Class

'----------------------------------------------------------------------------------------------------
'文件属性类
Class FileObj_Class
	Dim FormName,FileName,FilePath,FileSize,FileType,FileStart,FileExt
	'保存文件方法
	Public Function SaveToFile (Path)
		On Error Resume Next
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
	End Function
	'取得文件数据
	Public Function FileData
		UpFileStream.Position = FileStart

		FileData = UpFileStream.Read (FileSize)
	End Function
End Class
%> 
