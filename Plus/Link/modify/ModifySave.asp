<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../Plus/md5.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New FriendLinkModifySave
KSCls.Kesion()
Set KSCls = Nothing

Class FriendLinkModifySave
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
Public Sub Kesion()
Response.Write "<html>"
Response.Write "<head>"
Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
Response.Write "<title>����������������</title>"
Response.Write "</head>"

Dim LinkID, FolderID, SiteName, WebMaster, Email, OriPassWord, PassWord, ConPassWord, Locked, Url, LinkType, Logo, Hits, Recommend, Descript, TrueIP
Dim TempObj, LinkRS, LinkSql, RSCheck

LinkID = KS.ChkClng(KS.S("LinkID"))

OriPassWord = MD5(KS.R(Request.Form("OriPassWord")),16)
If OriPassWord = "" Then
      Call KS.AlertHistory("�޸�����������Ϣ��������ԭ������!", -1)
      Set KS = Nothing
End If
Set RSCheck = Server.CreateObject("Adodb.Recordset")
   RSCheck.Open " Select LinkID From KS_Link Where PassWord='" & OriPassWord & "' and linkid=" & linkid , Conn, 1, 1
   If RSCheck.EOF And RSCheck.BOF Then
      RSCheck.Close:Set RSCheck = Nothing
      Call KS.AlertHistory("�Բ���,�������ԭ����������!", -1)
      Set KS = Nothing
      Response.End
  End If
SiteName = KS.S("SiteName")
WebMaster = KS.S("Webmaster")
Email = KS.R(Request.Form("Email"))
FolderID = KS.S("FolderID")
PassWord = Request.Form("PassWord")
ConPassWord = Request.Form("ConPassWord")

If Trim(PassWord) <> Trim(ConPassWord) Then
            Call KS.AlertHistory("��վ���벻һ��!!!", -1)
            Set KS = Nothing
            Response.End
End If
PassWord = MD5(KS.R(PassWord),16)

Url = Replace(Replace(Request.Form("Url"), """", ""), "'", "")
LinkType = KS.S("LinkType")
Logo = Replace(Replace(Request.Form("Logo"), """", ""), "'", "")
Descript = KS.R(KS.S("Description"))

If SiteName <> "" Then
        If Len(SiteName) >= 200 Then
            Call KS.AlertHistory("��վ���Ʋ��ܳ���100���ַ�!", -1)
            Set KS = Nothing
             Response.End
        End If
 Else
        Call KS.AlertHistory("��������վ����!", -1)
        Set KS = Nothing
         Response.End
 End If
      Set LinkRS = Server.CreateObject("adodb.recordset")
      LinkSql = "select * from [KS_Link] Where LinkID=" & LinkID
      LinkRS.Open LinkSql, Conn, 1, 3
      LinkRS("SiteName") = SiteName
      LinkRS("WebMaster") = WebMaster
      LinkRS("Email") = Email
      If KS.S("PassWord") <> "" Then
      LinkRS("PassWord") = PassWord
      End If
      LinkRS("FolderID") = FolderID
      LinkRS("Url") = Url
      LinkRS("LinkType") = LinkType
      LinkRS("Logo") = Logo
      LinkRS("Description") = KS.HtmlEnCode(Descript)
      LinkRS.Update
      LinkRS.Close
      Set LinkRS = Nothing
      Response.Write ("<script>alert('�޸��������ӳɹ�!');location.href='../';</script>")
End Sub
End Class
%> 
