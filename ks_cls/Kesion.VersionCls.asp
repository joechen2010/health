<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Class KesionCls
	  Private Sub Class_Initialize()
      End Sub
	  Private Sub Class_Terminate()
	  End Sub
	 
	  '系统版本号
	  Public Property Get KSVer
		KSVer="KesionCMS V6.5 SP2 Free"
	  End Property 
	  
	  '系统缓存名称,如果你的一个站点下安装多套科汛系统，请分别将各个目录下的系统的缓存名称设置成不同
	  Public Property Get SiteSN
		SiteSN="KS6" & Replace(Replace(LCase(Request.ServerVariables("SERVER_NAME")), "/", ""), ".", "") 
	  End Property
	   
End Class
%>