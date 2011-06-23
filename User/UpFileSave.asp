<% Option Explicit %>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
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
        Private KS,KSUser
		Dim FilePath,MaxFileSize,AllowFileExtStr,AutoReName,RsConfigObj
		Dim FormName,Path,BasicType,ChannelID,UpType,TempFileStr,FormPath,ThumbFileName,ThumbPathFileName
		Dim UpFileObj,FsoObjName,AddWaterFlag,T,CurrNum,CreateThumbsFlag,FieldName,U_FileSize
		Dim DefaultThumb    '设定第几张为缩略图
		Dim ReturnValue
		Private Sub Class_Initialize()
		  Set T=New Thumb
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set T=Nothing
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Sub Kesion()
		 
		 IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "<script>top.location.href='Login';</script>"
		  Exit Sub
		 End If
		 If Trim(Request.ServerVariables("HTTP_REFERER"))="" Then
			Response.Write "<script>alert('非法上传！');history.back();</script>"
			Response.end
		 End If
		 if instr(lcase(Request.ServerVariables("HTTP_REFERER")),"user_upfile.asp")<=0 and instr(lcase(Request.ServerVariables("HTTP_REFERER")),"selectphoto.asp")<=0 then
			Response.Write "<script>alert('非法上传！');history.back();</script>"
			Response.end
		 end if
			
        If Cbool(KSUser.UserLoginChecked)=True Then
         IF KS.GetFolderSize(KSUser.GetUserFolder(ksuser.username))/1024>=KS.ChkClng(KSUser.SpaceSize) Then
		  Response.Write "<script>alert('上传失败，您的可用空间不够！');history.back();</script>"
		  response.end
		 End If
		End If
		Response.Write("<style type='text/css'>" & vbcrlf)
		Response.Write("<!--" & vbcrlf)
		Response.Write("body {" & vbcrlf)
		Response.Write("	margin-left: 0px;" & vbcrlf)
		Response.Write("	margin-top: 0px;" & vbcrlf)
		Response.Write("	font-size: 12px;" & vbcrlf)
		'Response.Write("    background:#EEF8FE;" & vbcrlf)
		Response.Write("}" & vbcrlf)
		Response.Write("-->" & vbcrlf)
		Response.Write("</style>" & vbcrlf)
		
		FsoObjName=KS.Setting(99)
		
		Set UpFileObj = New UpFileClass
		UpFileObj.GetData

		AutoReName = UpFileObj.Form("AutoRename")
		BasicType=KS.ChkClng(UpFileObj.Form("BasicType"))        ' 2-- 图片中心上传 3--下载中心缩略图/文件 41--动漫中心缩略图 42--动漫中心的动漫文件
		ChannelID=KS.ChkClng(UpFileObj.Form("ChannelID")) 
		UpType=UpFileObj.Form("Type")
		
		
		IF BasicType=0 and UpType<>"Field" then 
			Response.Write "<script>alert('请不要非法上传！');history.back();</script>"
			Response.end
		End If
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
		    Response.End()
		   End IF
		   RS.Close:Set RS=Nothing
		Else
			Select Case BasicType
			  Case 1     '文章中心缩略图
				if Not KS.ReturnChannelAllowUserUpFilesTF(ChannelID) Then
					Response.Write "<br><div align=center>对不起，系统不允许此频道上传文件,请与网站管理员联系!</div>"
					Exit Sub
				 End IF
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
				If UpType="File" Then '附件
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,0)
				FormPath = KS.ReturnChannelUserUpFilesDir(ChannelID,KSUser.UserName) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
				Else
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,1)
				FormPath = KS.ReturnChannelUserUpFilesDir(ChannelID,KSUser.UserName) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
				End If
			  Case 2     '图片中心上传图片
				 if Not KS.ReturnChannelAllowUserUpFilesTF(ChannelID) Then
					Response.Write "<br><div align=center>对不起，系统不允许此频道上传文件,请与网站管理员联系!</div>"
					Exit Sub
				 End IF
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(ChannelID)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,1)
				FormPath = KS.ReturnChannelUserUpFilesDir(ChannelID,KSUser.UserName) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
			  Case 3    
				 If Not KS.ReturnChannelAllowUserUpFilesTF(ChannelID) Then
					Response.Write "<br><div align=center>对不起，系统不允许此频道上传文件,请与网站管理员联系!</div>"
					Exit Sub
				 End IF
				CreateThumbsFlag=true
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
					Response.Write "<br><div align=center>对不起，系统不允许此频道上传文件,请与网站管理员联系!</div>"
					Exit Sub
				 End IF
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(4)   '设定文件上传最大字节数
				If UpType="Pic" Then '动漫中心缩略图
					AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(4,1)
					FormPath = KS.ReturnChannelUserUpFilesDir(4,KSUser.UserName) & "FlashPhoto/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
				Else
					AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(4,2)  '取允许上传的动漫类型
					FormPath = KS.ReturnChannelUserUpFilesDir(4,KSUser.UserName) & "FlashUrl/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
				End If
			 Case 5
			     If Not KS.ReturnChannelAllowUserUpFilesTF(5) Then
					Response.Write "<br><div align=center>对不起，系统不允许此频道上传文件,请与网站管理员联系!</div>"
					Exit Sub
				 End IF
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(5)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(5,1)
				FormPath = KS.ReturnChannelUserUpFilesDir(5,KSUser.UserName) & "Shop/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
			 Case 7   
				 If Not KS.ReturnChannelAllowUserUpFilesTF(7) Then
					Response.Write "<br><div align=center>对不起，系统不允许此频道上传文件,请与网站管理员联系!</div>"
					Exit Sub
				 End IF
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(7)   '设定文件上传最大字节数
				If UpType="Pic" Then '影片缩略图
					AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(7,1)
					FormPath = KS.ReturnChannelUserUpFilesDir(7,KSUser.UserName) & "MoviePhoto/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
				Else
					AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(ChannelID,2) &"|" & KS.ReturnChannelAllowUpFilesType(ChannelID,3) & "|"& KS.ReturnChannelAllowUpFilesType(ChannelID,4)  '取允许上传的动漫类型
					FormPath = KS.ReturnChannelUserUpFilesDir(7,KSUser.UserName) & "MovieUrl/" & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
				End If
	
			Case 8      '供求中心图片
				if Not KS.ReturnChannelAllowUserUpFilesTF(8) Then
					Response.Write "<br><div align=center>对不起，系统不允许此频道上传文件,请与网站管理员联系!</div>"
					Exit Sub
				 End IF
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(8)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(8,1)
				FormPath = KS.ReturnChannelUserUpFilesDir(8,KSUser.UserName) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
		    Case 9
				if Not KS.ReturnChannelAllowUserUpFilesTF(9) Then
					Response.Write "<br><div align=center>对不起，系统不允许此频道上传文件,请与网站管理员联系!</div>"
					Exit Sub
				 End IF
				CreateThumbsFlag=true
				MaxFileSize = KS.ReturnChannelAllowUpFilesSize(9)   '设定文件上传最大字节数
				AllowFileExtStr = KS.ReturnChannelAllowUpFilesType(9,1)
				FormPath = KS.ReturnChannelUserUpFilesDir(9,KSUser.UserName) & Year(Now()) & Right("0" & Month(Now()), 2) & "/"
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
		    Case 9994  '小论坛
			    If KS.ChkClng(KS.Setting(67))=0 Then
				    Response.Write "<script>alert('对不起，系统不允许此频道上传文件,请与网站管理员联系!');history.back();</script>"
					Exit Sub
				End If
				MaxFileSize = 1000    '设定文件上传最大字节数
				AllowFileExtStr = KS.Setting(68)  '取允许上传的类型
				FormPath =KS.ReturnChannelUserUpFilesDir(9994,KSUser.UserName)
		    Case 9993  '写日志附件
			    If KS.ChkClng(KS.SSetting(26))=0 Then
				    Response.Write "<script>alert('对不起，系统不允许此频道上传文件,请与网站管理员联系!');history.back();</script>"
					Exit Sub
				End If
				MaxFileSize = 1000    '设定文件上传最大字节数
				AllowFileExtStr = KS.SSetting(27)  '取允许上传的类型
				FormPath =KS.ReturnChannelUserUpFilesDir(9993,KSUser.UserName)
		    Case 999  '上传中心
				MaxFileSize = 100    '设定文件上传最大字节数
				AllowFileExtStr = "jpg|gif|png|swf"  '取允许上传的类型
				FormPath =KS.ReturnChannelUserUpFilesDir(999,KSUser.UserName)
			Case Else
			  MaxFileSize=0:AllowFileExtStr=""
			  Response.end
			End Select
        End If
		FormPath=Replace(FormPath,".","")
		IF Instr(FormPath,KS.Setting(3))=0 Then FormPath=KS.Setting(3) & FormPath
		FilePath=Server.MapPath(FormPath) & "\"

				
		Call KS.CreateListFolder(FormPath)       '生成上传文件存放目录
		
        If KS.Setting(97)=1 Then
		FormPath=KS.Setting(2) & FormPath
		End if
		ReturnValue = CheckUpFile(FilePath,MaxFileSize,AllowFileExtStr,AutoReName)
		
		if ReturnValue <> "" then
		       ReturnValue = Replace(ReturnValue,"'","\'")
			  Response.Write("<script language=""JavaScript"">")
			  Response.Write("alert('" & ReturnValue & "');")
			  if basictype=999 then
			  Response.Write("window.close();")
			  else
			  Response.Write("history.back(-1);")
			 end if
			  Response.Write("</script>")
		else  
            If UpType="Field" Then
					  Response.Write("<script language=""JavaScript"">")
					  Response.Write("parent.document.getElementById('"& FieldName & "').value='" & replace(TempFileStr,"|","") & "';")
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>恭喜，上传成功！</font>');")
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=user_upfile.asp?ChannelID=" & ChannelID & "&Type=Field&FieldID=" & UpFileObj.Form("FieldID") &"\'>');")
					  Response.Write("</script>")
					  Response.End()
			End If
			TempFileStr=replace(TempFileStr,"'","\'")
			Select Case BasicType
			   Case 1         '文章中心的上传缩略图
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
						  'Call KS.DeleteFile(replace(TempFileStr,"|",""))  '删除原图片
						 Else
						  Response.Write("parent.document.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
						 End If
					  end if 
						If Split(Split(KS.C_S(ChannelID,5),"@@@")(0),"|")(9)=1 Then
						  If KS.C_S(ChannelID,34)=0 Then
					       Response.Write("parent.ArticleContent.InsertPictureFromUp('" & replace(TempFileStr,"|","") &"');")
						  Else
						   Response.Write ("parent.insertHTMLToEditor('<img src=" & replace(TempFileStr,"|","") &" />');")
						  End If
						end if
						 Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>图片上传成功！</font>');")
				   End If 
				  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=user_upfile.asp?Channelid=" & ChannelID & "&type=" & UpType & "\'>');")
				  Response.Write("</script>")
			   Case 2          '图片中心的上传图片
				  Response.Write("<script language=""JavaScript"">")
				  If UPType="Single" Then
				  Response.Write("parent.document.myform.imgurl"&UpFileObj.Form("objid")&".value='"& replace(TempFileStr,"|","") &"';")
				  Response.Write("parent.document.myform.thumb"&UpFileObj.Form("objid")&".value='"& ThumbPathFileName &"';")
				  Response.Write("document.write('<br><div align=center><font size=2>图片上传成功！</font></div>');")
				  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'1; url=user_upfile.asp?Type=Single&ChannelID=" & ChannelID & "\'>');")
				  Else
				  Response.Write("parent.SetPicUrlByUpLoad(" & DefaultThumb & ",'" & TempFileStr &  "','" & ThumbPathFileName & "|');")
				  Response.Write("document.write('<br><br><div align=center>图片上传成功！</div>');")
				  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'1; url=user_upfile.asp?Channelid=" & ChannelID & "\'>');")
				  End If
				  Response.Write("</script>")
			  Case 3    '下载中心缩略图
				  Response.Write("<script language=""JavaScript"">")
				  If UPType="Pic" Then
				  Response.Write("parent.document.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
				  Response.Write("document.write('<br>&nbsp;&nbsp;&nbsp;&nbsp;图片上传成功！');")
				  Else   '下载中心的文件
				  Response.Write("parent.SetDownUrlByUpLoad('" & replace(TempFileStr,"|","") & "'," & U_FileSize & ");")
				  Response.Write("document.write('<br><br><div align=center>文件上传成功！</div>');")
				  End If
				  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'1; url=User_upfile.asp?channelid=" & ChannelID & "&Type=" & UPType &"\'>');")
				  Response.Write("</script>")
			  Case 4         '动漫中心的上传缩略图
				  Response.Write("<script language=""JavaScript"">")
				  If UpType="Pic" Then
				  Response.Write("parent.document.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
				  Response.Write("document.write('<br>&nbsp;&nbsp;&nbsp;&nbsp;图片上传成功！');")
				  Else
				  Response.Write("parent.document.myform.FlashUrl.value='" & replace(TempFileStr,"|","") & "';")
				  Response.Write("document.write('<br><br><div align=center>文件上传成功！</div>');")
				  End If
				  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'1; url=User_upfile.asp?channelid=4&Type=" & UpType & "\'>');")
				  Response.Write("</script>")
			  Case 5         '商城产品
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
						   Response.Write("parent.document.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
						   Response.Write("parent.document.myform.BigPhoto.value='" & replace(TempFileStr,"|","") & "';")
						  else
						   Response.Write("parent.document.myform.PhotoUrl.value='" & ThumbPathFileName & "';")
						   Response.Write("parent.document.myform.BigPhoto.value='" & replace(TempFileStr,"|","") & "';")
						  end if
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>图片上传成功！</font>');")
					  End If
				  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'1; url=User_upfile.asp?channelid=5&Type=" & UpType & "\'>');")
					  Response.Write("</script>")
			  Case 7         '影片中心的上传缩略图
				  Response.Write("<script language=""JavaScript"">")
				  If UpType="Pic" Then
				  Response.Write("parent.document.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
				  Response.Write("document.write('<br>&nbsp;&nbsp;&nbsp;&nbsp;图片上传成功！');")
				  Else
				  Response.Write("parent.document.myform.MovieUrl.value='" & replace(TempFileStr,"|","") & "';")
				  Response.Write("document.write('<br><br><div align=center>文件上传成功！</div>');")
				  End If
				  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'1; url=User_upfile.asp?channelid=7&Type=" & UpType & "\'>');")
				  Response.Write("</script>")
			  Case 8         '供求中心的上传缩略图
				  Response.Write("<script language=""JavaScript"">")
				  
				  if DefaultThumb=0 then
				   Response.Write("parent.document.myform.PhotoUrl.value='" &  replace(TempFileStr,"|","") & "';")
				  else
					 If KS.CheckFile(Replace(ThumbPathFileName,KS.Setting(2),""))=true Then        '检查是否存在缩略图
					  Response.Write("parent.document.myform.PhotoUrl.value='" & ThumbPathFileName & "';")
					  'Call KS.DeleteFile(replace(TempFileStr,"|",""))  '删除原图片
					 Else
					  Response.Write("parent.document.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
					 End If
				  end if
				  If KS.C_S(ChannelID,34)=0 Then
					       Response.Write("parent.GQContent.InsertPictureFromUp('" & replace(TempFileStr,"|","") &"');")
				  Else
						   Response.Write ("parent.insertHTMLToEditor('<img src=" & replace(TempFileStr,"|","") &" />');")
				  End If
				  'Response.Write("parent.GQContent.InsertPictureFromUp('" & replace(TempFileStr,"|","") &"');")
				  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>图片上传成功！</font>');")
				  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=user_upfile.asp?Channelid=8\'>');")
				  Response.Write("</script>")
				  Case 9
					  Response.Write("<script language=""JavaScript"">")
					  Response.Write("parent.document.myform.DownUrl.value='" &  replace(TempFileStr,"|","") & "';")
					  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>试卷上传成功！</font>');")
					  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=user_upfile.asp?Channelid=9\'>');")
					  Response.Write("</script>")		
			  Case 9999        '用户头像
				  Response.Write("<script language=""JavaScript"">")
				  Response.Write("parent.document.myform.UserFace.value='" & replace(TempFileStr,"|","") & "';")
				  Response.Write("parent.document.myform.showimages.src='" & replace(TempFileStr,"|","") & "';")
				  Response.Write("document.write('<br>&nbsp;&nbsp;&nbsp;&nbsp;图片上传成功！');")
				  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=User_upfile.asp?channelid=9999\'>');")
				  Response.Write("</script>")
			  Case 9998        '相册封面
				  Response.Write("<script language=""JavaScript"">")
				  Response.Write("parent.document.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
				  Response.Write("document.write('<br>&nbsp;&nbsp;&nbsp;&nbsp;图片上传成功！');")
				  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=User_upfile.asp?channelid=9998\'>');")
				  Response.Write("</script>")
			  Case 9997        '相片
				  Dim I,TempFileArr
				  TempFileStr=Left(tempfilestr,len(tempfilestr)-1)
				  Response.Write("<script language=""JavaScript"">")
				  Response.Write("parent.document.myform.PhotoUrls.value='" & TempFileStr & "';")
				  TempFileArr=split(TempFileStr,"|")
				  For I=Lbound(TempFileArr) to Ubound(TempFileArr)
				  Response.Write("try{parent.document.myform.view" & I+1 & ".src='" & TempFileArr(i) & "';}catch(e){}")
				  Next
				  Response.Write("</script>")
				  Response.write("<br><br><br><div><font color=red>恭喜您，照片上传成功！请按发布按钮进行保存。</font></div>")
				  Response.Write("<meta http-equiv='refresh' content='2; url=User_upfile.asp?channelid=9997&action=OK'>")
			  Case 9996        '圈子图片
				  Response.Write("<script language=""JavaScript"">")
				  Response.Write("parent.document.myform.showimages.src='" & replace(TempFileStr,"|","") & "';")
				  Response.Write("parent.document.myform.PhotoUrl.value='" & replace(TempFileStr,"|","") & "';")
				  Response.Write("document.write('<br>&nbsp;&nbsp;&nbsp;&nbsp;图片上传成功！');")
				  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=User_upfile.asp?channelid=9998\'>');")
				  Response.Write("</script>")
			  Case 9995        '用户头像
				  Response.Write("<script language=""JavaScript"">")
				  Response.Write("parent.document.myform.Url.value='" & replace(TempFileStr,"|","") & "';")
				  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;歌曲上传成功！');")
				  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=User_upfile.asp?channelid=9995\'>');")
				  Response.Write("</script>")
			  Case 9994,9993        '小论坛,博客
			      Response.Write("<script type=""text/JavaScript"">")
				  Response.Write("parent.InsertFileFromUp('" & TempFileStr &"','" & KS.Setting(3) & "');")
				  Response.Write("document.write('&nbsp;&nbsp;&nbsp;&nbsp;<font size=2>附件上传成功！</font>');")
				  Response.Write("document.write('<meta http-equiv=\'refresh\' content=\'2; url=user_upfile.asp?Channelid=" & ChannelID & "&type=" & UpType & "\'>');")
				  Response.Write("</script>")
			  Case 999
				  Response.Write("<script language=""JavaScript"">"&vbcrlf)
				  Response.Write("parent.location.href='selectphoto.asp?channelid=999';"&vbcrlf)
				  Response.Write("</script>"&vbcrlf)
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
				

				If NoIllegalStr(FileName)=False Then ErrStr=ErrStr&"文件：上传被禁止！\n"
				FileExtName = UpFileObj.File(FormName).FileExt
				FileContent = UpFileObj.File(FormName).FileData
				U_FileSize=UpFileObj.File(FormName).FileSize
				Dim FileType:FileType=UpFileObj.File(FormName).FileType
				

				'是否存在重名文件
				if U_FileSize > 1 then
					NoUpFileTF = False
					ErrStr = ""
					if UpFileObj.File(FormName).FileSize > CLng(FileSize)*1024 then
						ErrStr = ErrStr & FileName & "文件上传失败\n超过了限制，最大只能上传" & FileSize & "K的文件\n"
					end if
					If Cbool(KSUser.UserLoginChecked)=True Then
					 IF KS.ChkClng(KS.GetFolderSize(KSUser.GetUserFolder(ksuser.username))/1024+UpFileObj.File(FormName).FileSize/1024)>=KS.ChkClng(KSUser.SpaceSize) Then
					  Response.Write "<script>alert('上传失败1，您的可用空间不够！');history.back();</script>"
					  response.end
					End If
					End If
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
				
					
					If instr(FileName,";")>0 or instr(lcase(FileName),".asp")>0 or instr(lcase(FileName),".php")>0 or instr(lcase(FileName),".cdx")>0 or instr(lcase(FileName),".asa")>0 or instr(lcase(FileName),".cer")>0 or instr(lcase(FileName),".cfm")>0 or instr(lcase(FileName),".jsp")>0 then
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
			FileExtName = DealExtName(FileExtName)
			FileContent = UpFileObj.File(FormNameItem).FileData
			select case AutoNameType 
			  case "1"
				FileName= "副件" & FileName
			  case "2"
				FileName= RndStr&"."&FileExtName
			  Case "3"
				FileName= RndStr & FileName
			  case else
				FileName= Year(Now())&Right("0"&Month(Now()),2)&Right("0"&Day(Now()),2)&Right("0"&Hour(Now()),2)&Right("0"&Minute(Now()),2)&Right("0"&Second(Now()),2)&RandomFigure&"."&FileExtName
			End Select

			UpFileObj.File(FormNameItem).SaveToFile FilePath &FileName
			

			
			'======================增加检查文件内容是否合法===================================
			call CheckFileContent(FormPath  &FileName,UpFileObj.File(FormNameItem).FileSize /1024)
			'==================================================================================
			
		   TempFileStr=TempFileStr & FormPath & FileName & "|"
		   If AddWaterFlag = "1" Then   '在保存好的图片上添加水印
				call T.AddWaterMark(FilePath  & FileName)
		   End if
		  CurrNum=CurrNum+1
		  IF CreateThumbsFlag=true and (cint(CurrNum)=cint(DefaultThumb) or BasicType=2) Then
		      If KS.TBSetting(0)=0 then
			   if ThumbPathFileName="" then
			   ThumbPathFileName=FormPath &FileName
			   Else
			   ThumbPathFileName=ThumbPathFileName & "|" & FormPath & FileName
			   End If
			  Else
				ThumbFileName=split(FileName,".")(0)&"_S."&FileExtName
				Dim CreateTF:CreateTF=T.CreateThumbs(FilePath & FileName,FilePath & ThumbFileName)
				If CreateTF=true Then
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
		
		    if filesize>1000 then exit function  '超过1000K跳过检测
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
			   call KS.DeleteFile(path)
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
		'代码开始
		If Request.TotalBytes < 1 Then  '如果没有数据上传
			Err = 1
			Exit Sub
		End If
		Dim KS:Set KS=New PublicCls
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
