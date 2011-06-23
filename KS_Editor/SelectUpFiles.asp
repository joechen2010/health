<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../Plus/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 4.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
Dim KSCls
Set KSCls = New InsertPicture
KSCls.Kesion()
Set KSCls = Nothing

Class InsertPicture
        Private KS,KSUser
		Private AdminDir
		Private ChannelID
        Private CurrPath,InstallDir
		Private FromUrl
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser=New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
        Public Sub Kesion()
			  If KS.C("AdminName") = "" Or KS.C("AdminPass") = "" Then
				 call checkuser()
			  Else
				 Dim ChkRS:Set ChkRS = Server.CreateObject("ADODB.RecordSet")
				 ChkRS.Open "Select * From KS_Admin Where UserName='" & KS.R(KS.C("AdminName")) & "' And PassWord='" &  KS.R(KS.C("AdminPass")) & "'",Conn, 1, 1
				 If ChkRS.EOF And ChkRS.BOF Then
					 call checkuser()
				 else
				      Response.Redirect(KS.Setting(3) & KS.Setting(89) & "Include/SelectPic.asp?Currpath="& KS.GetUpFilesDir())
				 End If
			   ChkRS.Close:Set ChkRS = Nothing
			 End If

       End Sub
	   
	   Sub CheckUser()
	     IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "<script>alert('对不起，您没有权限!');window.returnValue='';window.close();</script>"
		  Exit Sub
		 End If
		 response.write "<iframe src='" &  KS.Setting(3) & "user/SelectPhoto.asp?ChannelID=999' frameborder='0' scrolling='auto' width='100%' height='100%'></iframe>"
	   End Sub
End Class
%>
 
