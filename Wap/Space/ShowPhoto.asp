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
		Private TotalPut,RS,MaxPerPage
		Private UserName,UserType,Template,BlogName,Author,xcid
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		    Set KSBCls=New BlogCls
			Set KSRFObj=New Refresh
		End Sub
        Private Sub Class_Terminate()
		    Set KS=Nothing
		    Set KSBCls=Nothing
			Set KSRFObj=Nothing
		End Sub
		Public Sub Kesion()
		    If KS.SSetting(0)=0 Then
			   Call KS.ShowError("对不起！","对不起，本站点关闭空间站点功能！")
			End If
			
			xcid=KS.Chkclng(KS.S("xcid"))
			UserName=KS.S("i")
			If UserName="" Then Response.End()
			Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select top 1 * From KS_Blog Where UserName='" & UserName & "'",conn,1,1
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
			
			Dim RSXC:Set RSXC=Server.CreateObject("ADODB.RECORDSET")
			RSXC.OPEN "Select * from ks_photoxc where id=" & xcid,Conn,1,3
			If RSXC.EOF And RSXC.BOF Then
			   RSXC.Close:set RSXC=Nothing
			   Call KS.ShowError("参数传递出错!","参数传递出错!")
			End If
			If KSUser.GroupID<>4 Then
			   If RSXC("Status")=0 Then
			      Call KS.ShowError("该相册尚未审核!","该相册尚未审核!")
			   ElseIf RSXC("Status")=2 then
			      Call KS.ShowError("该相册已被管理员锁定!","该相册已被管理员锁定!")
			   End If
			End If
			BlogName=RS("BlogName")
			Template="<wml>" &vbcrlf
			Template=Template & "<head>" &vbcrlf
			Template=Template & "<meta http-equiv=""Cache-Control"" content=""no-Cache""/>" &vbcrlf
			Template=Template & "<meta http-equiv=""Cache-Control"" content=""max-age=0""/>" &vbcrlf
			Template=Template & "</head>" &vbcrlf
			Template=Template & "<card id=""main"" title=""" & BlogName & "-浏览相片"">" &vbcrlf
			
		    UserType=KS.ChkClng(Conn.Execute("Select UserType From KS_User Where UserName='" & UserName & "'")(0))
			If UserType=1 Then
			   Template=Template & KSRFObj.LoadTemplate(KS.WSetting(23))'企业主模板
			Else
		       Template=Template & KSRFObj.LoadTemplate(KS.WSetting(20))'个人主模板
			End If
			Template=KSRFObj.KSLabelReplaceAll(Template)
			Template=KSBCls.ReplaceBlogLabel(RS,Template)
			Template=KSBCls.ReplaceAllLabel(UserName,Template)
			RS.Close
			If KS.S("Action")="Info" Then
			   TempStr=Info
			Else
			   Select Case RSXC("flag")
			       Case 1,2
				      If RSXC("Flag")=2 And Cbool(KSUser.UserLoginChecked)=False then
					     TempStr="<br/><br/>此相册设置会员可见，请先<a href=""../User/Login/?../Space/ShowPhoto.asp?xcid=" & xcid & "&amp;i="&UserName&""">登录</a>！<br/><br/>"
					  Else
					     TempStr=GetBody
					  End If
				   Case 3
				      If KS.S("Password")=RSXC("password") or Session("xcpass")=RSXC("password") Then
					     Session("xcpass")=KS.S("Password")
						 TempStr=GetBody
					  Else
					     If KS.S("Password")<>"" Or IsNull(KS.S("Password")) Then
					        If KS.S("Password")<>RSXC("password") or Session("xcpass")<>RSXC("password") Then
						       TempStr="出错啦,您输入的密码有误！<br/>"
							End If
						 End If
						 TempStr=TempStr&"请输入查看密码：<input name=""Password""  maxlength=""30"" value="""" emptyok=""false""/><a href=""ShowPhoto.asp?xcid="&xcid&"&amp;i="&UserName&"&amp;Password=$(Password)&amp;" & KS.WapValue & """>查看</a><br/><br/>"
					  End If
				   Case 4
				      If KSUser.UserName=RSXC("UserName") Then
					     TempStr=GetBody
					  Else
					     TempStr="<br/><br/>该相册设为稳私，只有相册主人才有权利浏览!<br/>"
						 TempStr=TempStr&"如果你是相册主人，<a href=""../User/Login/?../Space/ShowPhoto.asp?xcid=" & xcid & "&amp;i="&UserName&""">登录</a>后即可查看!<br/><br/>"
					  End If
			   End Select
			End If 
			RSXC("hits")=RSXC("hits")+1
			RSXC.Update
			RSXC.Close:set RSXC=Nothing
			Template=Replace(Template,"{$BlogMain}","【浏览相片】<br/>" & TempStr & "")
			Template=Template & "</card>" &vbcrlf
			Template=Template & "</wml>" &vbcrlf
			Response.Write Template
		End Sub
		
		
		Function GetBody()
		    MaxPerPage = 4
			If KS.S("page") <> "" Then
			   CurrentPage = KS.ChkClng(KS.G("page"))
			Else
			   CurrentPage = 1
			End If
			RS.Open "Select * from KS_Photozp Where xcid=" & xcid &" Order By AddDate Desc",Conn,1,1
			If RS.EOF And RS.BOF Then
			   TempStr = "该相册下没有照片！<br/>"
			Else
			   TotalPut = RS.RecordCount
			   If CurrentPage < 1 Then	CurrentPage = 1
			   If (CurrentPage - 1) * MaxPerPage > TotalPut Then
			      If (TotalPut Mod MaxPerPage) = 0 Then
				     CurrentPage = TotalPut \ MaxPerPage
				  Else
				     CurrentPage = TotalPut \ MaxPerPage + 1
				  End If
			   End If
			   If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < TotalPut Then
			      RS.Move (CurrentPage - 1) * MaxPerPage
			   Else
			      CurrentPage = 1
			   End If
			   Dim I
			   Do While Not RS.Eof
					PhotoUrl=RS("PhotoUrl")
					If KS.IsNul(PhotoUrl) Then PhotoUrl="images/nopic.gif"
					if left(PhotoUrl,1)="/" then PhotoUrl=right(PhotoUrl,len(PhotoUrl)-1)
					if lcase(left(PhotoUrl,4))<>"http" then PhotoUrl=KS.Setting(2) & KS.Setting(3) & PhotoUrl
			      TempStr = TempStr&"<img src='" &PhotoUrl& "'  /><br/>"
				  TempStr = TempStr&"<a href='ShowPhoto.asp?xcid="&xcid&"&amp;ID="&RS("ID")&"&amp;Action=Info&amp;i="&RS("UserName")&"&amp;Password="&Password&"&amp;" & KS.WapValue & "'>"&RS("Title")&"("&RS("hits")&")</a><br/>"
				  RS.Movenext
				  I = I + 1
				  If I >= MaxPerPage Then Exit Do
			   Loop
			   TempStr = TempStr & KS.ShowPagePara(TotalPut, MaxPerPage, "ShowPhoto.asp", True, "个", CurrentPage, KS.QueryParam("page"))
			   TempStr = TempStr & "<br/>"
			End If 
			GetBody=TempStr   
			RS.Close:Set  RS=Nothing
		End Function
		
		Function Info()
		    Dim ID:ID=KS.Chkclng(KS.S("ID"))
			RS.Open "Update KS_Photozp set hits=hits+1 where ID=" & ID,Conn,1,3 
			RS.Open "Select top 1 * from KS_Photozp Where ID=" & ID,Conn,1,1
			If RS.bof And RS.eof Then
			   TempStr = "参数传递出错!<br/>"
			Else
					Dim PhotoUrl:PhotoUrl=RS("PhotoUrl")
					If KS.IsNul(PhotoUrl) Then PhotoUrl="images/nopic.gif"
					if left(PhotoUrl,1)="/" then PhotoUrl=right(PhotoUrl,len(PhotoUrl)-1)
					if lcase(left(PhotoUrl,4))<>"http" then PhotoUrl=KS.Setting(2) & KS.Setting(3) & PhotoUrl
			
			   TempStr = ""
			   TempStr = TempStr &"<img src='" &PhotoUrl& "'  /><br/>"
			   TempStr = TempStr &"<a href='"&PhotoUrl&"'>点击下载</a><br/>"
			   TempStr = TempStr &"名称:"&RS("Title")&" "&RS("hits")&"<br/>"
			 
			   If RS("PhotoSize")<>"" Or IsNull(RS("PhotoSize")) Then
			      TempStr = TempStr &"相片大小:"&RS("PhotoSize")&"<br/>"
			   End If
			   TempStr = TempStr &"所属相册:" & Conn.Execute("select xcname from KS_Photoxc where id=" & RS("xcid"))(0) & "<br/>"
			   If RS("Descript")<>"" Or IsNull(RS("Descript")) Then
			      TempStr = TempStr &"相片描述:"&RS("descript")&"<br/>"
			   End If
			End If
			Info = TempStr & "<anchor>返回上一页<prev/></anchor><br/><br/>"
			RS.Close:Set  RS=Nothing 
		End Function
		
		Function GetStatusStr(val)
            Select Case Val
		    Case 1:GetStatusStr="公开"
			Case 2:GetStatusStr="会员"
			Case 3:GetStatusStr="密码"
			Case 4:GetStatusStr="隐私"
		    End Select
			GetStatusStr="<b>" & GetStatusStr & "</b>"
		End Function
End Class
%>