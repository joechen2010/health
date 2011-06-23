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
Set KSCls = New BlogIndex
KSCls.Kesion()
Set KSCls = Nothing

Class BlogIndex
        Private KS,KSSCls,KSBCls,KSRFObj
		Private UserName,BlogName,Template,UserType
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		    Set KSBCls=New BlogCls
			Set KSRFObj=New Refresh
		End Sub
        Private Sub Class_Terminate()
		    Call CloseConn()
		    Set KS=Nothing
			Set KSSCls=Nothing
		    Set KSBCls=Nothing
			Set KSRFObj=Nothing
		End Sub
		Public Sub Kesion()
		    If KS.SSetting(0)=0 Then
			   Call KS.ShowError("对不起！","对不起，本站点关闭空间站点功能！")
		    End If
			'If Request.ServerVariables("QUERY_STRING")<>"" Then 
			If Request.QueryString("i")<>"" Then
			 Call Show()
			Else
			    
				Dim FileContent
				FileContent = KSRFObj.LoadTemplate(KS.WSetting(29))
				FileContent = KSRFObj.KSLabelReplaceAll(FileContent)
				FileContent = KS.GetEncodeConversion(FileContent)
				Response.Write FileContent  
			End If
		End Sub
		Sub Show()
		     UserName=KS.R(KS.S("i"))
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
		    UserType=KS.ChkClng(Conn.Execute("Select top 1 UserType From KS_User Where UserName='" & UserName & "'")(0))
			If UserType=1 Then
			   Template=Template & KSRFObj.KSLabelReplaceAll(KSRFObj.LoadTemplate(KS.WSetting(23)))'企业主模板
			Else
		       Template=Template & KSRFObj.KSLabelReplaceAll(KSRFObj.LoadTemplate(KS.WSetting(20)))'个人主模板
			End If
          
		  
              
		    Template=KSBCls.ReplaceBlogLabel(RS,Template)
		    Template=KSBCls.ReplaceAllLabel(UserName,Template)
		    RS.Close:Set  RS=Nothing
			Action=KS.S("Action")
			Set KSSCls=New SpaceCls
			Select Case Action
			    Case "friend"
				Template=Replace(Template,"{$BlogMain}",KSSCls.FriendList)
				Case "group"
				Template=Replace(Template,"{$BlogMain}",KSSCls.GroupList)
				Case "photo"
				Template=Replace(Template,"{$BlogMain}",KSSCls.PhotoList)
				Case "log"
				Template=Replace(Template,"{$BlogMain}",KSSCls.LogList)
				Case "guest"
				Template=Replace(Template,"{$BlogMain}",KSSCls.GuestList)
				Case "info"
				Template=Replace(Template,"{$BlogMain}",KSSCls.UserInfo)
				Case "xx"       Template=Replace(Template,"{$BlogMain}",KSSCls.xxList)
				Case "job"      Template=Replace(Template,"{$BlogMain}",KSSCls.EnterPriseJob)
				Case "news"     Template=Replace(Template,"{$BlogMain}",KSSCls.EnterpriseNews)
				case "intro"    Template=Replace(Template,"{$BlogMain}",KSSCls.EnterpriseIntro)
				case "product"  Template=Replace(Template,"{$BlogMain}",KSSCls.EnterprisePro)
				Case Else
				Template=Replace(Template,"{$BlogMain}",GetMain)
			End Select
			Template=Template & "</card>" &vbcrlf
			Template=Template & "</wml>" &vbcrlf
		    Response.Write Template
		End Sub
		
		Function GetMain()
		    Dim I,TemplateSub
			If UserType=1 Then
			   TemplateSub=KSRFObj.LoadTemplate(KS.WSetting(24))'企业首页副模板
			Else
			   TemplateSub=KSRFObj.LoadTemplate(KS.WSetting(21))'个人首页副模板
			End If
			TemplateSub=Replace(TemplateSub,"{$ShowNewAlbum}",GetNewAlbum)'照片
			TemplateSub=Replace(TemplateSub,"{$ShowNewInfo}",KSSCls.ListXX)'10条信息集
			
			'=================企业空间替换==========================
			TemplateSub=Replace(TemplateSub,"{$ShowNews}",GetEnterPriseNews)'公司动态
			TemplateSub=Replace(TemplateSub,"{$ShowSupply}",GetSupply)'供应信息
			TemplateSub=Replace(TemplateSub,"{$ShowProduct}",GetProduct)'最新产品
			TemplateSub=Replace(TemplateSub,"{$ShowProductList}",GetProductList)'文本方式显示最新产品
			TemplateSub=Replace(TemplateSub,"{$ShowIntro}",GetEnterpriseintro)'企业简介
			TemplateSub=Replace(TemplateSub,"{$ShowShortIntro}",GetEnterpriseShortintro)'企业简介(短)
			'========================================================
			GetMain=KSBCls.ReplaceAllLabel(UserName,TemplateSub)
		End Function

		
End Class
%>