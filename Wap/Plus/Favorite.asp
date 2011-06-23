<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
'********************************
'* 程序功能：收藏夹
'* 演示地址: http://wap.kesion.com/
'********************************
%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<% Response.ContentType="text/vnd.wap.wml" %><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card title="收藏夹">
<p>
<%
Dim KSCls
Set KSCls = New User_Favorite
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%
Class User_Favorite
        Private KS
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		    Set KS=Nothing
		End Sub
		Public Sub Kesion()
			Dim InfoID,ChannelID,UserCheck,UserName,Title,AddDate
			Dim RS,RS1
			InfoID=KS.ChkClng(KS.S("InfoID"))
			ChannelID=KS.S("ChannelID")
			If InfoID="" or ChannelID="" then
			   UserCheck=false
			End If
			
			If Cbool(KSUser.UserLoginChecked)=True Then
			   UserName=KSUser.UserName
			   UserCheck=True
			End If
			'-------------如果不是注册用户---------------
			If UserCheck=False Then
			   Response.Write "对不起，只有登陆的用户才可以使用无限个私人收藏夹！所以请先<a href=""../User/Login/?../user/User_Favorite.asp?InfoID="&InfoID&"&amp;ChannelID="&ChannelID&""">注册/登陆</a>吧！"
			End If
			'-------------如果是注册用户---------------
			If UserCheck=True Then
			   set RS=Server.CreateObject("ADODB.Recordset")
			   RS.Open "select * from "&KS.C_S(ChannelID,2)&" where ID="&InfoID&"",Conn,1,1
			   IF RS.Eof And RS.Bof Then
			      RS.Close:Set RS=Nothing
				  Response.Write "对不起，收藏失败！可能这个"&KS.C_S(ChannelID,3)&"已经被删除了！<br/>"
			   Else
			      Title=RS("Title")'文章名称
				  AddDate=RS("AddDate")'添加日期
				  RS.Close:Set RS=Nothing
				  set RS1=Server.CreateObject("ADODB.Recordset")
				  RS1.Open "select * from KS_Favorite Where ChannelID="&ChannelID&" And UserName='"&UserName&"' And InfoID="&InfoID&"",Conn,1,2
				  IF RS1.Eof And RS1.Bof Then
				     set RS=Server.CreateObject("ADODB.Recordset")
					 RS.Open "select * from KS_Favorite",Conn,1,2
					 RS.Addnew
					 RS("UserName")=UserName'会员昵称
					 RS("ChannelID")=ChannelID'频道ID
					 RS("InfoID")=InfoID'文章,图片,下载等的ID
					 RS("AddDate")=Now()'收藏日期
					 RS.Update
					 Response.Write "收藏成功！<br/><br/>"
					 Response.Write "名称："&Title&"<br/>"
					 Response.Write "更新时间："&AddDate&"<br/>"
					 Response.Write "以后您就可以通过我的地盘上的<a href='../User/User_Favorite.asp?ChannelID="&ChannelID&"&amp;" & KS.WapValue & "'>我的收藏夹</a>来直接访问该"&KS.C_S(ChannelID,3)&"了！<br/>"
					 RS.Close:Set RS=Nothing
				  Else
				     Response.Write "收藏失败！<br/>"""&Title&"""己经在你收藏夹里了！<br/>"
				  End If
				  RS1.Close:Set RS1=Nothing
			  End If
           End If
		   Response.Write "<br/>"
		   Response.Write "<anchor>返回上页<prev/></anchor><br/>"
		   Response.Write " <a href=""" & KS.GetGoBackIndex & """>返回首页</a><br/>"
	   End Sub
End Class
%>

