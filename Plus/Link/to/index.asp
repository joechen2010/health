<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New ToLink
KSCls.Kesion()
Set KSCls = Nothing

Class ToLink
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		Dim LinkID, ObjRS,Url
		LinkID = KS.ChkClng(request.QueryString)
		Set ObjRS = Server.CreateObject("Adodb.RecordSet")
		ObjRS.Open "Select top 1 Url,hits From KS_Link Where LinkID=" & LinkID, Conn, 1, 3
		If Not ObjRS.EOF Then
		  ObjRS(1) = ObjRS(1) + 1
		  ObjRS.Update
		  Url=ObjRS(0)
		  ObjRS.Close:Set ObjRS=Nothing
		  Response.Redirect url
		Else
		  Response.Write "������������!"
		End If
		  ObjRS.Close
		  Set ObjRS = Nothing
		End Sub

End Class
%>

 
