<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.SpaceCls.asp"-->
<!--#include file="SpaceCls.asp"-->
<%
Dim KSCls
Set KSCls = New Blog
KSCls.Kesion()
Set KSCls = Nothing

Class Blog
        Private KS,KSBCls,KSSCls,KSRFObj
		Private RS,UserType
		Private UserName,Template,BlogName
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
			Set KSSCls=New SpaceCls
		    Set KSBCls=New BlogCls
			Set KSRFObj=New Refresh
		End Sub
        Private Sub Class_Terminate()
		    Set KS=Nothing
			Set KSSCls=Nothing
		    Set KSBCls=Nothing
			Set KSRFObj=Nothing
		End Sub
		Public Sub Kesion()
		    If KS.SSetting(0)=0 Then
		       Call KS.ShowError("对不起！","对不起，本站点关闭空间站点功能！")
			End If
			UserName=KS.S("i")
			If UserName="" Then Response.End()
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select top 1 * From KS_Blog Where UserName='" & UserName & "'",Conn,1,1
			If RS.Eof And RS.Bof Then
			   RS.Close:Set RS=Nothing
			   Call KS.ShowError("该用户没有开通空间站点！","该用户没有开通空间站点！")
		    End If
		    If KSUser.GroupID<>4 Then
			   If RS("Status")=0 Then
			      RS.Close:Set RS=Nothing
				  Call KS.ShowError("该空间站点尚未审核！","该空间站点尚未审核！")
			   ElseIf RS("Status")=2 Then
			      RS.Close:Set RS=Nothing
				  Call KS.ShowError("该空间站点已被管理员锁定！","该空间站点已被管理员锁定！")
			   End If
	        End If
		    BlogName=RS("BlogName")
		    Template="<wml>" &vbcrlf
		    Template=Template & "<head>" &vbcrlf
		    Template=Template & "<meta http-equiv=""Cache-Control"" content=""no-Cache""/>" &vbcrlf
		    Template=Template & "<meta http-equiv=""Cache-Control"" content=""max-age=0""/>" &vbcrlf
		    Template=Template & "</head>" &vbcrlf
		    Template=Template & "<card id=""main"" title=""" & BlogName & """>" &vbcrlf
		    UserType=KS.ChkClng(Conn.Execute("Select UserType From KS_User Where UserName='" & UserName & "'")(0))
		    If UserType=1 Then
		       Template=Template & KSRFObj.LoadTemplate(KS.WSetting(23))'企业主模板
		    Else
		       Template=Template & KSRFObj.LoadTemplate(KS.WSetting(20))'个人主模板
		    End If
		    Template=KSBCls.ReplaceBlogLabel(RS,Template)
		    Template=KSBCls.ReplaceAllLabel(UserName,Template)
		    RS.Close
		    Template=Replace(Template,"{$BlogMain}",KSSCls.LogList)
			Template=Template & "</card>" &vbcrlf
			Template=Template & "</wml>" &vbcrlf
		    Response.Write Template
		    Set  RS=Nothing
		End Sub
End Class
%>