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
	 
	  'ϵͳ�汾��
	  Public Property Get KSVer
		KSVer="KesionCMS V6.5 SP2 Free"
	  End Property 
	  
	  'ϵͳ��������,������һ��վ���°�װ���׿�Ѵϵͳ����ֱ𽫸���Ŀ¼�µ�ϵͳ�Ļ����������óɲ�ͬ
	  Public Property Get SiteSN
		SiteSN="KS6" & Replace(Replace(LCase(Request.ServerVariables("SERVER_NAME")), "/", ""), ".", "") 
	  End Property
	   
End Class
%>