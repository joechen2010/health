<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 5.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New SiteIndex
KSCls.Kesion()
Set KSCls = Nothing

Class SiteIndex
        Private KS, KSR,str,c_str,ID,Template,categoryname
		Private TotalPut,CurrentPage,MaxPerPage
		Private Sub Class_Initialize()
		 If (Not Response.IsClientConnected)Then
			Response.Clear
			Response.End
		 End If
		  MaxPerPage=20
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		   Dim I
		   ID=KS.ChkClng(Request("id"))
		   If ID=0 Then 
		     ks.die "�Ƿ�����!"
		   End If
		   Template = KSR.LoadTemplate(KS.Setting(103))
		   FCls.RefreshFolderID = "0" '���õ�ǰˢ��Ŀ¼ID Ϊ"0" ��ȡ��ͨ�ñ�ǩ
		   Call GetSubject()
		   
		   Template=KSR.KSLabelReplaceAll(Template)
		   Response.Write Template  
		End Sub
		
		Sub GetSubject()
		      Dim RS:Set RS=Server.CreateObject("adodb.recordset")
			  RS.Open "select top 1 * from KS_PKZT where id=" & id,conn,1,1
			  If RS.Eof And RS.Bof Then
			    RS.Close
				Set RS=Nothing
				KS.Die "�Ҳ���PK����!"
			  End If
			  Template=replace(template,"{$GetPKID}",rs("id"))
			  Template=replace(template,"{$GetPKTitle}",rs("title"))
			  If KS.IsNul(rs("newslink")) Then
			  Template=replace(template,"{$GetBackGroundNews}","")
			  Else
			  Template=replace(template,"{$GetBackGroundNews}","<a href='" & rs("newslink") & "' target='_blank'>����鿴�������� >></a>")
			  End If
			  Template=replace(template,"{$GetZFTips}",rs("zftips"))
			  Template=replace(template,"{$GetFFTips}",rs("fftips"))
		End Sub
		
End Class
%>
