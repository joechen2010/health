<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
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
Set KSCls = New Preview
KSCls.Kesion()
Set KSCls = Nothing

Class Preview
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		 On Error Resume Next
		Dim PreviewImagePath, FileExtName, FileIconDic, FileIcon, AvaiLabelShowTypeStr, PicPara
		PreviewImagePath = KS.G("FilePath")
		AvaiLabelShowTypeStr = "jpg,gif,bmp,pst,png,ico"
		Set FileIconDic = CreateObject("Scripting.Dictionary")
		FileIconDic.Add "txt", "../images/FileIcon/txt.gif"
		FileIconDic.Add "gif", "../images/FileIcon/gif.gif"
		FileIconDic.Add "exe", "../images/FileIcon/exe.gif"
		FileIconDic.Add "asp", "../images/FileIcon/asp.gif"
		FileIconDic.Add "html", "../images/FileIcon/html.gif"
		FileIconDic.Add "htm", "../images/FileIcon/html.gif"
		FileIconDic.Add "jpg", "../images/FileIcon/jpg.gif"
		FileIconDic.Add "jpeg", "../images/FileIcon/jpg.gif"
		FileIconDic.Add "pl", "../images/FileIcon/perl.gif"
		FileIconDic.Add "perl", "../images/FileIcon/perl.gif"
		FileIconDic.Add "zip", "../images/FileIcon/zip.gif"
		FileIconDic.Add "rar", "../images/FileIcon/zip.gif"
		FileIconDic.Add "gz", "../images/FileIcon/zip.gif"
		FileIconDic.Add "doc", "../images/FileIcon/doc.gif"
		FileIconDic.Add "xml", "../images/FileIcon/xml.gif"
		FileIconDic.Add "xsl", "../images/FileIcon/xml.gif"
		FileIconDic.Add "dtd", "../images/FileIcon/xml.gif"
		FileIconDic.Add "vbs", "../images/FileIcon/vbs.gif"
		FileIconDic.Add "js", "../images/FileIcon/vbs.gif"
		FileIconDic.Add "wsh", "../images/FileIcon/vbs.gif"
		FileIconDic.Add "sql", "../images/FileIcon/script.gif"
		FileIconDic.Add "bat", "../images/FileIcon/script.gif"
		FileIconDic.Add "tcl", "../images/FileIcon/script.gif"
		FileIconDic.Add "eml", "../images/FileIcon/mail.gif"
		FileIconDic.Add "swf", "../images/FileIcon/flash.gif"
		If PreviewImagePath = "" Then
			PreviewImagePath = "../images/FileIcon/DefaultPreview.gif"
		Else
			FileExtName = Right(PreviewImagePath, Len(PreviewImagePath) - InStrRev(PreviewImagePath, "."))
			If InStr(AvaiLabelShowTypeStr, lcase(FileExtName)) = 0 Then
				FileIcon = FileIconDic.Item(LCase(FileExtName))
				If FileIcon = "" Then
					FileIcon = "../images/FileIcon/unknown.gif"
				End If
				PreviewImagePath = FileIcon
				PicPara = " width=""30"" height=""30"" "
			Else
				PicPara = ""
			End If
		End If
		Set FileIconDic = Nothing
		
		Response.Write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">"
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<link href=""Admin_Style.CSS"" rel=""stylesheet"">"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		Response.Write "<title>‘§¿¿</title>"
		Response.Write "</head>"
		Response.Write "<body topmargin=""0"" leftmargin=""0"">"
		Response.Write "<table width=""100%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		Response.Write "  <tr>"
		Response.Write "    <td align=""center"" valign=""middle""><img  " & PicPara & " src=""" & PreviewImagePath & """></td>"
		Response.Write "  </tr>"
		Response.Write "</table>"
		Response.Write "</body>"
		Response.Write "</html>"
		
		End Sub
End Class
%> 
