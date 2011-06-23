<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
Dim KSCls
Set KSCls = New SpaceMore
KSCls.Kesion()
Set KSCls = Nothing

Class SpaceMore
        Private KS,KSRFObj
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		    Set KSRFObj = New Refresh
		End Sub
        Private Sub Class_Terminate()
		    Call CloseConn()
			Set KS=Nothing
		End Sub
		
		Public Sub Kesion()
		    Dim FileContent
			FileContent = KSRFObj.LoadTemplate(KS.WSetting(30))
			FileContent = KSRFObj.KSLabelReplaceAll(FileContent)
			FileContent = Replace(FileContent,"{$ShowMain}",GetPhotoList())
			Response.Write FileContent  
	    End Sub
		
		Function GetPhotoList()
		    MaxPerPage = 4
		    Dim ClassID:ClassID=KS.Chkclng(KS.S("ClassID"))
		    Dim Recommend:Recommend=KS.Chkclng(KS.S("Recommend"))
		    If KS.S("page") <> "" Then
			   CurrentPage = KS.ChkClng(KS.G("page"))
		    Else
			   CurrentPage = 1
		    End If
		    str = "【相册查找】<br/>"
		    str = str & "<select name=""ClassID"">"
		    str = str & "<option value=""0"">所有分类</option>"
		    Dim RSC:set RSC=Conn.Execute("select ClassName,ClassID from KS_PhotoClass order by orderid")
		    If Not RSC.EOF Then
		     Do While Not RSC.EOF
			    str = str & "<option value=""" & RSC(1) & """>" & RSC(0) & "</option>"
				RSC.Movenext
			 Loop
		    End If
		    RSC.Close:set RSC=Nothing
			str = str & "</select> "
			
		    str = str & "名称:<input type=""text"" size=""12"" name=""key""/>"
		    str = str & "<anchor>查找<go href=""MorePhoto.asp?"&KS.WapValue&""" method=""post"">"
		    str = str & "<postfield name=""ClassID"" value=""$(ClassID)""/>"
		    str = str & "<postfield name=""key"" value=""$(key)""/>"
		    str = str & "</go></anchor><br/>"
		    str = str & "【相册列表】<br/>"
		  
		    
		    Dim Param:Param=" where status=1"
		    If ClassID<>0 Then Param=Param & " and  ClassID=" & ClassID
		    If Recommend<>0 Then Param=Param & " and  Recommend=1"
		    If KS.S("key")<>"" Then Param=Param & " and XCName like '%" & KS.R(KS.S("key")) &"%'"
		    Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
			RSObj.Open "Select * from KS_Photoxc " & Param & " order by id desc",Conn,1,1
		    If RSObj.EOF and RSObj.BOF  Then
		       str = str & "没有创建相册！<br/>"
		    Else
		       TotalPut = RSObj.RecordCount
			   If CurrentPage < 1 Then	CurrentPage = 1
			   If (CurrentPage - 1) * MaxPerPage > TotalPut Then
			      If (TotalPut Mod MaxPerPage) = 0 Then
				     CurrentPage = TotalPut \ MaxPerPage
				  Else
				     CurrentPage = TotalPut \ MaxPerPage + 1
				  End If
			   End If
			   If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < TotalPut Then
			      RSObj.Move (CurrentPage - 1) * MaxPerPage
			   Else
			      CurrentPage = 1
			   End If
			   Dim I
			   Do While Not RSObj.EOF
			       Dim PhotoUrl:PhotoUrl=RSObj("PhotoUrl")
				   if left(PhotoUrl,1)="/" then PhotoUrl=right(PhotoUrl,len(PhotoUrl)-1)
				   if lcase(left(PhotoUrl,4))<>"http" then PhotoUrl=KS.Setting(2) & KS.Setting(3) & PhotoUrl
		          str = str & "<img src="""&PhotoUrl&""" width=""120"" height=""90"" alt=""""/><br/>"
				  str = str & "<a href=""ShowPhoto.asp?xcid="&RSObj("ID")&"&amp;i="&RSObj("UserName")&"&amp;"&KS.WapValue&""">"&RSObj("xcname")&"</a><br/>"
				  str = str & ""&RSObj("xps")&"张/"&RSObj("hits")&"次["&GetStatusStr(RSObj("flag"))&"]<br/>"
			      RSObj.MoveNext
				  I = I + 1
				  If I >= MaxPerPage Then Exit Do
			   Loop
			   str = str & KS.ShowPagePara(TotalPut, MaxPerPage, "MorePhoto.asp", True, "个", CurrentPage, KS.QueryParam("page"))
		    End If
		    GetPhotoList = str & "<br/>"
		    RSObj.Close:Set RSObj=Nothing
		End Function
		
	    Function GetStatusStr(val)
          Select Case Val
		  Case 1:GetStatusStr="公开"
		  Case 2:GetStatusStr="会员"
		  Case 3:GetStatusStr="密码"
		  Case 4:GetStatusStr="隐私"
		  End Select
		  GetStatusStr="<b>" & GetStatusStr & "</b>"
	    End Function
End Class
%>
