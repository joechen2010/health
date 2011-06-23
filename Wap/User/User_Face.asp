<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="UpFileSave.asp"-->
<%Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="形象选择">
<p>
<%
Dim KSCls
Set KSCls = New User_Face
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%
Class User_Face
        Private KS,Prev,DomainStr
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
			DomainStr=KS.GetDomain
		End Sub
        Private Sub Class_Terminate()
		    Set KS=Nothing
		    Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		    IF Cbool(KSUser.UserLoginChecked)=False Then
			   Response.redirect DomainStr&"User/Login/"
			   Exit Sub
			End If
			%>
            <a href="User_EditInfo.asp?<%=KS.WapValue%>">基本信息</a>
            <a href="User_Face.asp?<%=KS.WapValue%>">形象选择</a>
            <a href="User_EditInfo.asp?Action=ContactInfo&amp;<%=KS.WapValue%>">详细信息</a>
            <a href="User_EditInfo.asp?Action=PassInfo&amp;<%=KS.WapValue%>">修改密码</a>
            <a href="User_EditInfo.asp?Action=PassQuestionInfo&amp;<%=KS.WapValue%>">安全设置</a>
            <br/>
			<%
			Dim Action
			Action=Trim(Request("Action"))
			Select Case Action
			    Case "Edit"
				Call Edit()
			    Case "EditSave"
				Call EditSave()
			    Case "UpSaveUrl"
				Call UpSaveUrl()
				Case Else
				Call FaceMain()
			End Select
			If Prev=True Then
			   Response.Write "<anchor>返回上一页<prev/></anchor><br/>"
			End If
			Response.Write "<br/>"
			Response.Write "<a href=""Index.asp?" & KS.WapValue & """>我的地盘</a><br/>" &vbcrlf
			Response.Write "<a href=""" & KS.GetGoBackIndex & """>返回首页</a>" &vbcrlf
		End Sub
		
		Sub FaceMain()
		    Dim UserFaceSrc:UserFaceSrc=KSUser.UserFace
			If Left(UserFaceSrc,1)="/" Then UserFaceSrc=Right(UserFacesRC,Len(UserFaceSrc)-1)
			If KS.IsNul(UserFaceSrc) Then  UserFaceSrc=KS.Setting(2) & KS.Setting(3) & "Images/Face/6.gif"
			If lcase(Left(UserFaceSrc,4))<>"http" Then UserFaceSrc=KS.Setting(2)& KS.Setting(3) & UserFaceSrc
		    %>
            【形象选择】<br/>
            <img src="<%=UserFaceSrc%>" alt="<%=KSUser.UserName%>"/><br/>
            <a href="User_Face.asp?Action=Edit&amp;<%=KS.WapValue%>">更改形象</a><br/>
            <a href="User_UpFile.asp?ChannelID=9999&amp;<%=KS.WapValue%>">上传形象</a><br/>
            <%
		End Sub
		
		Sub Edit()
		    Dim CurrPage,k,CurrPage1
		    %>
            【形象选择】<br/>
            <%
			CurrPage=Trim(Request("CurrPage"))
			If CurrPage="" Then
			   CurrPage=1
			Else
			   CurrPage=CurrPage
			End If
			If CurrPage=55 Then
			   CurrPage=CurrPage
			   CurrPage1=56
			Else
			   CurrPage=CurrPage
			   CurrPage1=CurrPage+2
			End If
			For K=CurrPage To CurrPage1
			    Response.Write "<img src=""" & KS.Setting(2) & KS.Setting(3) & "Images/Face/" & K & ".gif""/><br/>"
				Response.Write "<a href=""User_Face.asp?Action=EditSave&amp;PhotoUrl=Images/Face/"&K&".gif&amp;" & KS.WapValue & """>形象选择</a><br/>"
			Next
			If CurrPage < 55 Then
			   Response.Write "<a href=""User_Face.asp?Action=Edit&amp;CurrPage="&CurrPage+3&"&amp;" & KS.WapValue & """>下页</a> "
			   Response.Write "<a href=""User_Face.asp?Action=Edit&amp;CurrPage=55&amp;" & KS.WapValue & """>尾页</a> "
			End If
			If CurrPage > 1 Then
			   Response.Write "<a href=""User_Face.asp?Action=Edit&amp;CurrPage=1&amp;" & KS.WapValue & """>首页</a> "
			   Response.Write "<a href=""User_Face.asp?Action=Edit&amp;CurrPage="&CurrPage-3&"&amp;" & KS.WapValue & """>上页</a> "
			End If
			Response.Write "<br/>"
			Response.Write "共：56个形象<br/>"
		End Sub
		
		Sub EditSave()
		    Dim PhotoUrl
		    PhotoUrl = KS.S("PhotoUrl")
			Conn.Execute ("Update KS_User set UserFace='"&PhotoUrl&"' where UserName='"&KSUser.UserName&"'")
			Response.Write "形象修改成功。<br/>"
		End Sub
		
		Sub UpSaveUrl()
		    'On Error Resume Next
			'上传图片
			Dim KSUpFile,PhotoUrl
			Set KSUpFile = New UpFileSave
			PhotoUrl=KSUpFile.UpFileUrl
			Set KSUpFile = Nothing
			
			 Dim RS: Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select * From KS_User Where UserName='" & KSUser.UserName & "'",Conn,1,3
			  IF Not RS.Eof Then
				 RS("UserFace")=PhotoUrl
		 		 RS.Update
				 Call KS.FileAssociation(1024,rs("UserID"),PhotoUrl,1)
			  End If
              RS.Close:Set RS=Nothing
			Response.Write "形象修改成功。<br/>"
		End Sub
End Class
%>