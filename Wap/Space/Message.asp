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
        Private KS,KSBCls,KSRFObj
		Private RS
		Private UserName,UserType,Template,BlogName
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		    Set KSBCls=New BlogCls
			Set KSRFObj=New Refresh
		End Sub
        Private Sub Class_Terminate()
		    Set KS=Nothing
		    Set KSBCls=Nothing
			Set KSLabel=Nothing
		End Sub
		Public Sub Kesion()
		    If KS.SSetting(0)=0 Then
			   Call KS.ShowError("对不起！","对不起，本站点关闭空间站点功能！")
			End If
			Call KSUser.UserLoginChecked()
			UserName=KS.S("UserName")
			If UserName="" Then Response.End()
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select * From KS_Blog Where UserName='" & UserName & "'",Conn,1,1
			If RS.Eof And RS.Bof Then
			   Call KS.ShowError("该用户没有开通空间站点！","该用户没有开通空间站点！")
			End If
			BlogName=RS("BlogName")
			Template="<wml>" &vbcrlf
			Template=Template & "<head>" &vbcrlf
			Template=Template & "<meta http-equiv=""Cache-Control"" content=""no-Cache""/>" &vbcrlf
			Template=Template & "<meta http-equiv=""Cache-Control"" content=""max-age=0""/>" &vbcrlf
			Template=Template & "</head>" &vbcrlf
			Template=Template & "<card id=""main"" title=""" & BlogName & "-给我留言"">" &vbcrlf
		    UserType=KS.ChkClng(Conn.Execute("Select UserType From KS_User Where UserName='" & UserName & "'")(0))
			If UserType=1 Then
			   Template=Template & KSRFObj.KSLabelReplaceAll(KSRFObj.LoadTemplate(KS.WSetting(23)))'企业主模板
			Else
		       Template=Template & KSRFObj.KSLabelReplaceAll(KSRFObj.LoadTemplate(KS.WSetting(20)))'个人主模板
			End If
			Template=KSBCls.ReplaceBlogLabel(RS,Template)
			Template=KSBCls.ReplaceAllLabel(UserName,Template)
			Dim TempStr
			Select Case KS.S("Action")
			    Case "MessageSave" 
				   TempStr=MessageSave
				Case Else
				   TempStr=MessageMain
			End Select
			Template=Replace(Template,"{$BlogMain}","【给我留言】<br/>" & TempStr & "")
			Template=Template & "</card>" &vbcrlf
			Template=Template & "</wml>" &vbcrlf
			Response.Write Template
			RS.Close:Set  RS=Nothing
		End Sub

		Function MessageMain()
		    MessageMain = "昵称：<input name=""AnounName" & Minute(now) & Second(Now) & """ type=""text"" maxlength=""500"" size=""20"" value="""&KSUser.UserName&"""/><br/>"
			MessageMain = MessageMain & "标题：<input name=""Title" & Minute(now) & Second(Now) & """ type=""text"" maxlength=""500"" size=""20"" value=""""/><br/>"
			MessageMain = MessageMain & "内容：<input name=""Content" & Minute(now) & Second(Now) & """ type=""text"" maxlength=""250"" size=""20"" value=""""/><br/>"
			If KS.Setting(53)=1 Then
			   MessageMain = MessageMain & "验证码：<input name=""VerifyCode" & Minute(now) & Second(Now) & """ type=""text"" size=""4"" maxlength=""4"" value=""""/><b>" & KS.GetVerifyCode & "</b><br/>"
			End If
			MessageMain = MessageMain & "<anchor>提交留言<go href=""Message.asp?Action=MessageSave&amp;UserName="&UserName&"&amp;"&KS.WapValue&""" method=""post"">"
			MessageMain = MessageMain & "<postfield name=""AnounName"" value=""$(AnounName" & Minute(now) & Second(Now) & ")""/>"
			MessageMain = MessageMain & "<postfield name=""Title"" value=""$(Title" & Minute(now) & Second(Now) & ")""/>"
			MessageMain = MessageMain & "<postfield name=""Content"" value=""$(Content" & Minute(now) & Second(Now) & ")""/>"
			MessageMain = MessageMain & "<postfield name=""VerifyCode"" value=""$(VerifyCode" & Minute(now) & Second(Now) & ")""/>"
            MessageMain = MessageMain & "</go></anchor><br/>"
		End Function  
		
		'保存留言
		Function MessageSave()
		    Dim Content:Content=Request("Content")
			Dim AnounName:AnounName=KS.S("AnounName")
			'Dim HomePage:HomePage=KS.S("HomePage")
			Dim Title:Title=KS.S("Title")
			If AnounName="" Then 
			   MessageSave="请填写你的昵称!<br/>"
			   Exit Function
			End if
			If Title="" Then 
			   MessageSave="请填写留言主题!<br/>"
			   Exit Function
			End if
			If Content="" Then 
			   MessageSave="请填写留言内容!<br/>"
			   Exit Function
			End if
			If KS.Setting(53)=1 Then
			   IF Trim(KS.S("Verifycode"))<>Trim(Session("Verifycode")) Then
			      MessageSave="你输入的认证码不正确!<br/>"
				  Exit Function
			   End If
			End If
			Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select * From KS_BlogMessage where 1=0",Conn,1,3
			RS.AddNew
			RS("AnounName")=AnounName
			RS("Title")=Title
			RS("UserName")=KS.S("UserName")
			RS("HomePage")="http://wap.kesion.com/"
			RS("Content")=Content
			RS("UserIP")=KS.GetIP
			RS("AddDate")=Now
			RS.Update
			RS.Close:Set RS=Nothing
			MessageSave="签写留言成功。<br/>"
		End Function
End Class
%>