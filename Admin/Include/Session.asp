<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSLCCls1
Set KSLCCls1 = New LoginCheckCls1
KSLCCls1.Run()
Set KSLCCls1 = Nothing

Class LoginCheckCls1
		Private ComeUrl
		Private TrueSiteUrl
		Private AdminDirStr
		Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		End Sub
		
		'检查后台管理认证码
		Sub CheckSiteManageCode()
			If EnableSiteManageCode = True And Trim(KS.C("AdminLoginCode")) <> SiteManageCode Then
				Response.Write ("<script>top.location.href='/';</script>")
				Response.End()
			End If
		End Sub
				
		Sub Run()
		   Call CheckSiteManageCode
			'ComeUrl = Trim(Request.ServerVariables("HTTP_REFERER"))
			'TrueSiteUrl = Trim(Request.ServerVariables("HTTP_HOST"))
		 If KS.IsNul(KS.C("AdminName")) Or KS.IsNul(KS.C("AdminPass")) Or KS.IsNul(KS.C("PowerList"))="" Or KS.IsNUL(KS.C("UserName")) Then
			Response.Write ("<script>top.location.href='/';</script>")
			Response.End()
		  Else
			 Dim ChkRS:Set ChkRS = Server.CreateObject("ADODB.RecordSet")
			 ChkRS.Open "Select top 1 * From KS_Admin Where UserName='" & KS.R(KS.C("AdminName")) & "'",Conn, 1, 1
			 If ChkRS.EOF And ChkRS.BOF Then
			     ChkRS.Close:Set ChkRS=Nothing
				 Response.Write ("<script>top.location.href='/';</script>")
				 Response.End
			 Else
			     If ChkRS("PassWord")<>KS.C("AdminPass") Then
					 ChkRS.Close:Set ChkRS=Nothing
					 Response.Write ("<script>top.location.href='/';</script>")
					 Response.End
				 End If
			 End If

		   ChkRS.Close:Set ChkRS = Nothing
		 End If
		End Sub
End Class
%> 
