<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<%
Option Explicit
Response.ContentType="text/vnd.wap.wml"
Response.Charset="utf-8"
Response.Write "<?xml version=""1.0"" encoding=""utf-8""?>" &vbcrlf
Response.Write "<!DOCTYPE wml PUBLIC ""-//WAPFORUM//DTD WML 1.1//EN"" ""http://www.wapforum.org/DTD/wml_1.1.xml"">" &vbcrlf
%>
<!--#include file="Conn.asp"-->
<!--#include file="KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
Dim KSCls
Set KSCls = New SiteIndex
KSCls.Kesion()
Set KSCls = Nothing

Class SiteIndex
        Private KS, KSR
		Private Sub Class_Initialize()
		    If (Not Response.IsClientConnected) Then
			   Response.Clear
			   Response.End
		    End If
		    Set KS=New PublicCls
		    Set KSR=New Refresh
		End Sub
		
        Private Sub Class_Terminate()
		    Call CloseConn()
		    Set KS=Nothing
			Set KSR=Nothing
		End Sub
		
		Public Sub Kesion()
		    Dim FileContent
			FileContent = KSR.LoadTemplate(KS.WSetting(9))
			FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
			FileContent = KSR.KSLabelReplaceAll(FileContent)'替换通用标签为内容
			FileContent = KS.GetEncodeConversion(FileContent)'对全文所有html源代码进行语法规范化
			Response.Write FileContent
		End Sub
End Class
%>