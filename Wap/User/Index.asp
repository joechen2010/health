<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'Option Explicit
Response.ContentType="text/vnd.wap.wml"
Response.Charset="utf-8"
Response.Write "<?xml version=""1.0"" encoding=""utf-8""?>" &vbcrlf
Response.Write "<!DOCTYPE wml PUBLIC ""-//WAPFORUM//DTD WML 1.1//EN"" ""http://www.wapforum.org/DTD/wml_1.1.xml"">" &vbcrlf
%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
Dim KSCls
Set KSCls = New SiteIndex
KSCls.Kesion()
Set KSCls = Nothing
Class SiteIndex
        Private KS,KSRFObj
		Private FileContent
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		    Set KSRFObj=New Refresh
		End Sub
        Private Sub Class_Terminate()
		    Call CloseConn()
			Set KS=Nothing
			Set KSRFObj=Nothing
		End Sub
		
		Public Sub Kesion()
		    IF Cbool(KSUser.UserLoginChecked)=False Then
			   Response.Redirect KS.GetDomain&"User/Login/"
			   Exit Sub
			End If
		    FileContent = KSRFObj.LoadTemplate(KS.WSetting(19))
		    FileContent = KSRFObj.KSLabelReplaceAll(FileContent)
			FileContent = Replace(FileContent,"{$GetUserBasicInfo}",GetUserBasicInfo)
			FileContent = Replace(FileContent,"{$GetSpaceMenu}",GetSpaceMenu)
			Response.Write FileContent
		End Sub
		
		Function GetUserBasicInfo()
		    If KS.IsNul(KSUser.RealName) Then
			   GetUserBasicInfo = "<a href=""User_EditInfo.asp?"&KS.WapValue&""">"&KSUser.UserName&"</a>"
			Else
			   GetUserBasicInfo = "<a href=""User_EditInfo.asp?"&KS.WapValue&""">"&KSUser.RealName&"</a>"
			End If
			GetUserBasicInfo = GetUserBasicInfo & ",欢迎您!您是"
			If KSUser.ChargeType=1 Then 
			   GetUserBasicInfo = GetUserBasicInfo & "扣点数计费用户<br/>"
			ElseIf KSUser.ChargeType=2 Then
			   GetUserBasicInfo = GetUserBasicInfo & "有效期计费用户,到期时间：" & Cdate(KSUser.BeginDate)+KSUser.Edays & "<br/>"
			ElseIf KSUser.ChargeType=3 Then
			   GetUserBasicInfo = GetUserBasicInfo & "无限期计费用户<br/>"
			End If
			GetUserBasicInfo = GetUserBasicInfo & "资金" & KSUser.Money & "元," & KS.Setting(45) & KSUser.Point & "个,积分" & KSUser.Score & "分。<br/>"
		End Function
		Function GetSpaceMenu()
		  Dim Str
		  If KS.SSetting(0)=1 Then
		   If Conn.Execute("select top 1 UserName from ks_blog where username='"&KSUser.UserName&"'").eof then
			  GetSpaceMenu = "个人空间,体验交友的超IN的快感!展示自已,<a href=""User_Blog.asp?Action=BlogEdit&amp;" & KS.WapValue & """>现在开通</a>吧!<br/>" : Exit Function
		   ElseIf Conn.Execute("Select top 1 status From KS_Blog Where UserName='"&KSUser.UserName&"'")(0)<>1 Then
			  GetSpaceMenu = "你的空间还没有通过审核或被锁定！<br/>" :Exit Function
		   Else
			  If KSUser.UserType=0 Then
				 str=str & "体验交友的超IN的快感!展示自已!<br/>" &vbcrlf
				 str=str & "<a href=""User_Blog.asp?Action=BlogEdit&amp;" & KS.WapValue & """>空间设置</a> " &vbcrlf
				 str=str & "<a href=""User_Blog.asp?Action=BlogList&amp;" & KS.WapValue & """>我的日志</a><br/>" & vbcrlf
				 str=str & "<a href=""User_Music.asp?" & KS.WapValue & """>我的音乐</a> " &vbcrlf
				 str=str & "<a href=""User_Photo.asp?" & KS.WapValue & """>我的相册</a><br/>" &vbcrlf
				 str=str & "<a href=""User_Team.asp?" & KS.WapValue & """>我的圈子</a> " & vbcrlf
				 str=str & "<a href=""User_Friend.asp?" & KS.WapValue & """>我的好友</a><br/>" &vbcrlf
				 str=str & "<a href=""User_Blog.asp?Action=Message&amp;" & KS.WapValue & """>我的留言</a> " &vbcrlf
				 str=str & "<a href=""User_Class.asp?" & KS.WapValue & """>我的专栏</a><br/><br/>" &vbcrlf
				 str=str & "<a href=""user_Enterprise.asp?" & KS.WapValue & """>升级为企业空间</a><img src=""../Images/new_l.gif"" alt=""""/><br/>" &vbcrlf
			Else
				 str=str & "<a href=""user_Enterprise.asp?" & KS.WapValue & """>企业基本信息</a> " &vbcrlf
				 str=str & "<a href=""user_Enterprise.asp?action=intro&amp;" & KS.WapValue & """>企业简介</a><br/>" &vbcrlf
				 str=str & "<a href=""user_MYshop.asp?" & KS.WapValue & """>企业产品管理</a> " &vbcrlf
				 str=str & "<a href=""user_EnterpriseNews.asp?" & KS.WapValue & """>企业新闻管理</a><br/>" &vbcrlf
				 str=str & "<a href=""user_Enterprise.asp?action=job&amp;" & KS.WapValue & """>企业招聘</a> " &vbcrlf
				 str=str & "<a href=""User_Blog.asp?Action=BlogEdit&amp;" & KS.WapValue & """>企业空间设置</a><br/>" &vbcrlf
				 str=str & "<a href=""User_Blog.asp?Action=BlogList&amp;" & KS.WapValue & """>企业日志管理</a> " &vbcrlf
				 str=str & "<a href=""User_Photo.asp?" & KS.WapValue & """>企业相册管理</a><br/>" &vbcrlf
				 str=str & "<a href=""User_Team.asp?" & KS.WapValue & """>企业圈子管理</a> " &vbcrlf
				 str=str & "<a href=""User_Blog.asp?Action=Message&amp;" & KS.WapValue & """>企业留言本</a><br/>" &vbcrlf
				 str=str & "<a href=""User_Class.asp?" & KS.WapValue & """>专栏分类管理</a><br/><br/>" &vbcrlf
			End If
			  str=str & "个人空间:<a href=""" & KS.GetDomain & "?u=" & KSUser.UserName & "&amp;" & KS.WapValue & """>" & KS.GetDomain & "?u=" & KSUser.UserName & "</a><br/>" &vbcrlf
			  TopDir=KS.Setting(3)&KS.Setting(91)&"User/"&KSUser.UserName&"/"'读取用户文件夹
			   str=str & ShowTable("images/bar.gif","您的总空间",KS.GetFolderSize(TopDir)/1024,KS.Setting(50))
		   End If
		   
		    
		End If
           GetSpaceMenu=str
		End Function
		'（图片对象名称，标题对象名称，更新数，总数）
		Function ShowTable(SrcName,TxtName,str,c)
			Dim Tempstr,Src_js,Txt_js,TempPercent
			If C = 0 Then C = 99999999
			Tempstr = str/C
			TempPercent = FormatPercent(tempstr,0,-1)
			ShowTable = ""&TxtName&""&c/1024&"MB,"
			ShowTable = ShowTable + "<img src=""" + SrcName + """ width="""&FormatNumber(tempstr*300,0,-1)&""" height=""10"" alt="""+TxtName+"""/>"
			If FormatNumber(tempstr*100,0,-1) < 80 Then
			   ShowTable = ShowTable + "已使用:" & TempPercent & "<br/>"
			Else
			   ShowTable = ShowTable + "已用:" & TempPercent & ",请赶快清理！<br/>"
			End If
			ShowTable = ShowTable
		End Function		
End Class
%>