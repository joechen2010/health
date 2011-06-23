<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="用户信息">
<p>
<%
Dim KSCls
Set KSCls = New ShowUser
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%
Class ShowUser
        Private KS,DomainStr
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
			DomainStr=KS.GetDomain
		End Sub
        Private Sub Class_Terminate()
		    Set KS=Nothing
		    Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		    Response.write "<anchor>返回上一页<prev/></anchor><br/>"
			Response.write KS.GetReadMessage
			Dim UserID:UserID=KS.S("UserID")
			Dim Keyword:Keyword=KS.S("Keyword")
			If UserID<>"" Then param="UserID="&UserID&""
			If Keyword<>"" Then param="UserName like '%"&Keyword&"%'"
			Set RS=Conn.Execute("Select * from KS_User Where "&Param&"")
			If RS.Eof And RS.Bof Then
			   Response.write "参数不正确!<br/>"
			Else
			   Dim Privacy:Privacy=RS("Privacy")
		    Dim UserFaceSrc:UserFaceSrc=RS("UserFace")
			If KS.IsNul(UserFaceSrc) Then UserFaceSrc="Images/Face/1.gif"
			If Left(UserFaceSrc,1)="/" Then UserFaceSrc=Right(UserFaceSrc,Len(UserFaceSrc)-1)
			If lcase(Left(UserFaceSrc,4))<>"http" Then UserFaceSrc=KS.Setting(2)& KS.Setting(3) & UserFaceSrc
			   Response.Write "<img src=""" & UserFaceSrc & """ width=""100"" height=""100"" alt=""""/><br/>" &vbcrlf
			   Response.Write "用户名:"&RS("UserName")&"<br/>" &vbcrlf
			   If Privacy=2 or Privacy=1 Then
			      Response.Write "姓名:保密<br/>" &vbcrlf
			   Else
			      Dim RealName:RealName=RS("RealName")
				  If IsNull(RealName) Or RealName="" Then RealName="暂无"
				  Response.Write "姓名:" & RealName & "<br/>" &vbcrlf
			   End If
			   If Privacy=2 or Privacy=1 Then
			      Response.Write "性别:保密<br/>" &vbcrlf
			   Else
			      Dim Sex:Sex=RS("Sex")
				  If IsNull(Sex) Or Sex="" Then Sex="暂无"
				  Response.Write "性别:" & Sex & "<br/>" &vbcrlf
			   End If
			   If Privacy=2 or Privacy=1 Then
			      Response.Write "生日:保密<br/>" &vbcrlf
			   Else
			      Dim BirthDay:BirthDay=RS("BirthDay")
				  If IsNull(BirthDay) Or BirthDay="" Then BirthDay="暂无"
				  Response.Write "生日:" & BirthDay & "<br/>" &vbcrlf
			   End If
			   If Privacy=2 or Privacy=1 Then
			      Response.Write "电话:保密<br/>" &vbcrlf
		       Else
			      Dim Mobile:Mobile=RS("Mobile")
				  If IsNull(Mobile) Or Mobile="" Then Mobile="暂无"
				  Response.Write "电话:<a href=""wtai://wp/mc;" & Mobile & """ >" & Mobile & "</a><br/>" &vbcrlf
			   End If
			   If Privacy=2 Then
			      Response.Write "邮箱:保密<br/>" &vbcrlf
			   Else
			      Dim Email:Email=RS("Email")
				  If IsNull(Email) Or Email="" Then Email="暂无"
				  Response.Write "邮箱:" & Email & "<br/>" &vbcrlf
			   End If
			   If Privacy=2 Then
			      Response.Write "ＱＱ:保密<br/>" &vbcrlf
			   Else
			      Dim QQ:QQ=RS("QQ")
				  If IsNull(QQ) Or QQ="" Then QQ="暂无"
				  Response.Write "ＱＱ:" & QQ & "<br/>" &vbcrlf
			   End If
			   'If Privacy=2 or Privacy=1 Then
			      'Response.Write "地区:保密<br/>" &vbcrlf
			   'Else
			      'Dim Province:Province=RS("Province")
				  'If IsNull(Province) Or Province="" Then Province=""
				  'Dim City:City=RS("City")
				  'If IsNull(City) Or Fax="" Then City="未知"
				  'Response.Write "地区:" & Province & City & "<br/>" &vbcrlf
			   'End If
			   'If Privacy=2 or Privacy=1 Then
			      'Response.Write "地址:保密<br/>" &vbcrlf
			   'Else
			      'Dim AddRess:AddRess=RS("AddRess")
				  'If IsNull(AddRess) Or AddRess="" Then AddRess="暂无"
				  'Response.Write "地址:" & AddRess & "<br/>" &vbcrlf
			   'End If
			   'If Privacy=2 or Privacy=1 Then
			      'Response.Write "邮编:保密<br/>" &vbcrlf
			   'Else
			      'Dim Zip:Zip=RS("Zip")
				  'If IsNull(Zip) Or Zip="" Then Zip="暂无"
				  'Response.Write "邮编:" & ZIP & "<br/>" &vbcrlf
			   'End If		  
			   If Privacy=2 or Privacy=1 Then
			      Response.Write "签名:保密<br/>" &vbcrlf
			   Else
			      Dim Sign:Sign=RS("Sign")
				  If IsNull(Sign) Or Sign="" Then Sign="暂无"
				  Response.Write "签名:" & Sign & "<br/>" &vbcrlf
			   End If
			   IF Cbool(KSUser.UserLoginChecked)=True Then
			      Response.write "<a href=""User_Message.asp?Action=new&amp;ToUser="&RS("UserName")&"&amp;" & KS.WapValue & """>发送短信</a> "
			      Response.Write "<a href=""User_Friend.asp?Action=saveF&amp;touser="&RS("UserName")&"&amp;" & KS.WapValue & """>加为好友</a><br/>"
			   Else
			      Dim ToUrl
				  IF UserID<>"" Then ToUrl="ShowUser.asp?UserID="&UserID&""
				  IF Keyword<>True Then ToUrl="ShowUser.asp?Keyword="&Keyword&""
				  Response.Write "<a href=""Login.asp?"&ToUrl&""">注册/登陆</a>后可以加为好友<br/>"
			   End if
			   Response.Write "<a href="""&DomainStr&"Index.asp?u="&RS("UserName")&"&amp;" & KS.WapValue & """>进入他的个人空间</a><br/>"
			   Response.Write "<br/>"
			   IF Cbool(KSUser.UserLoginChecked)=True Then Response.Write "<a href=""Index.asp?" & KS.WapValue & """>我的地盘</a><br/>"
			   Response.Write " <a href=""" & KS.GetGoBackIndex & """>返回首页</a><br/>"
		  
			End if
		End Sub
End Class
%>