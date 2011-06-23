<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Md5.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<%Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="我的资料">
<p>
<%
Dim KSCls
Set KSCls = New User_EditInfo
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%
Class User_EditInfo
        Private KS,DomainStr,Prev
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
			DomainStr=KS.GetDomain
		End Sub
        Private Sub Class_Terminate()
		    Set KS=Nothing
		    Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		    IF Cbool(KSUser.UserLoginChecked)=False Then
			   Response.redirect DomainStr&"User/Login/"
			   Exit Sub
			End If
			Select Case KS.S("Action")
			    Case "VerifyCodeInfo":Call VerifyCodeInfo()'重置安全码
				Case "VerifyCodeSave":Call VerifyCodeSave()'确认重置安全码
			    Case "ContactInfo":Call ContactInfo()'修改详细信息
				Case "PassInfo":Call PassInfo()'修改密码
				Case "PassSave":Call PassSave()
				Case "PassQuestionInfo":Call PassQuestionInfo()'找回密码设置
				Case "PassQuestionSave":Call PassQuestionSave()
				Case "BasicInfoSave":Call BasicInfoSave()
				Case "ContactInfoSave":Call ContactInfoSave()
				Case "EditBasicInfo":Call EditBasicInfo()'修改基本信息
				Case Else:Call EditInfoMain()
			End Select
			If Prev=True Then
			   Response.Write "<anchor>返回上一页<prev/></anchor><br/>"
			End If
			Response.Write "<br/>"
			If KS.S("Action")<>"" Then
			   Response.Write "<a href=""User_EditInfo.asp?" & KS.WapValue & """>个人资料</a><br/>" &vbcrlf
			End If
			Response.Write "<a href=""Index.asp?" & KS.WapValue & """>我的地盘</a><br/>" &vbcrlf
			Response.Write "<a href=""" & KS.GetGoBackIndex & """>返回首页</a>" &vbcrlf
	    End Sub
		
		Sub EditInfoMain()
		    Dim UserFaceSrc:UserFaceSrc=KSUser.UserFace
			If KS.IsNul(UserFaceSrc) Then UserFaceSrc="Images/Face/1.gif"
			If Left(UserFaceSrc,1)="/" Then UserFaceSrc=Right(UserFaceSrc,Len(UserFaceSrc)-1)
			If lcase(Left(UserFaceSrc,4))<>"http" Then UserFaceSrc=KS.Setting(2)& KS.Setting(3) & UserFaceSrc
		    %>
            【基本信息】<br/>
            <img src="<%=UserFaceSrc%>" alt="" width="85" height="85"/><br/>
            【<a href="User_Face.asp?<%=KS.WapValue%>">更改形象</a>】<br/>
            会 员 名：<%=KSUser.UserName%><br/>
            会 员 组：<%=KS.GetUserGroupName(KSUser.GroupID)%><br/>
            本次登录：<%=KSUser.LastLoginTime%><br/>
            登 录 IP：<%=KSUser.LastLoginIP%><br/>
            登录次数：<%=KSUser.LoginTimes%> 次<br/>
            <a href="User_EditInfo.asp?Action=EditBasicInfo&amp;<%=KS.WapValue%>">基本信息</a>
            <a href="User_EditInfo.asp?Action=ContactInfo&amp;<%=KS.WapValue%>">详细信息</a><br/>
            <a href="User_EditInfo.asp?Action=PassInfo&amp;<%=KS.WapValue%>">修改密码</a>
            <a href="User_EditInfo.asp?Action=PassQuestionInfo&amp;<%=KS.WapValue%>">安全设置</a><br/><br/>
            <a href="User_EditInfo.asp?action=VerifyCodeInfo&amp;<%=KS.WapValue%>">重置安全码</a>
            <br/>
            <%
		End Sub
		
		Sub VerifyCodeInfo()
		    %>
            【重置安全码】<br/>
            您确认重置安全码么?<br/>
            重设置后您以前的书签将无效!重设置安全码成功后,请进入首页后把首页加入书签,方便您下次从书签自动登陆!<br/>
            <a href="User_EditInfo.asp?Action=VerifyCodeSave&amp;<%=KS.WapValue%>">确认重置</a><br/>
            <%
		End Sub
		
		Sub VerifyCodeSave()
		    %>
            【重置安全码】<br/>
            <%
			Dim wap,UserRS
			wap=MD5(KS.MakeRandomChar(20),32)
			Set UserRS=Server.CreateObject("Adodb.RecordSet")
			UserRS.Open "Select wap from KS_User Where UserName='"&KSUser.UserName&"'",Conn,1,3
			UserRS("wap")=wap
			UserRS.Update
			UserRS.Close:Set UserRS=Nothing
			%>
            重置安全码成功,任何页面存为书签,您以后再从此书签进入,将不会再要求登陆<br/>
            <a href="Index.asp?<%=KS.WSetting(2)%>=<%=wap%>">进入我的地盘...</a><br/>
            </p>
            </card>
            </wml>
            <%
			Response.End
		End Sub
	    '基本信息
	    Sub EditBasicInfo()
		   %>
           【基本资料】<br/>
           会员名称:<%=KSUser.UserName%><br/>
           真实姓名:<input name="RealName<%=Minute(Now)%><%=Second(Now)%>" type="text" value="<%=KSUser.Realname%>"/><br/>
           真实性别:<select name="Sex">
           <option value="">请选择性别</option>
           <option value="男" <%If KSUser.sex="男" Then Response.Write "selected=""selected"""%>>男</option>
           <option value="女" <%If KSUser.sex="女" Then Response.Write "selected=""selected"""%>>女</option>
           </select><br/>
           身份证号:<input  name="IDCard" type="text" value="<%=KSUser.idcard%>"/><br/>
           出生日期:格式:0000-00-00<br/>
           <input name="Birthday<%=Minute(Now)%><%=Second(Now)%>" type="text" value="<%=KSUser.Birthday%>"/><br/>
           邮箱地址:<input name="Email<%=Minute(Now)%><%=Second(Now)%>" type="text" value="<%=KSUser.Email%>"/><br/>
           隐私设定:<select name="Privacy">
           <option value="0" <%If KSUser.Privacy=0 Then Response.Write "selected=""selected"""%>>公开全部信息</option>
           <option value="1" <%If KSUser.Privacy=1 Then Response.Write "selected=""selected"""%>>公开部分信息</option>
           <option value="2" <%If KSUser.Privacy=2 Then Response.Write "selected=""selected"""%>>完全保密信息</option>
           </select><br/>
           个人签名:<input name="Sign<%=Minute(Now)%><%=Second(Now)%>" type="text" value="<%=KSUser.Sign%>"/><br/>
           <anchor>确定修改<go href="User_EditInfo.asp?Action=BasicInfoSave&amp;<%=KS.WapValue%>" method="post">
           <postfield name="Realname" value="$(Realname<%=Minute(Now)%><%=Second(Now)%>)"/>
           <postfield name="Sex" value="$(Sex)"/>
           <postfield name="IDCard" value="$(IDCard<%=Minute(Now)%><%=Second(Now)%>)"/>
           <postfield name="Birthday" value="$(Birthday<%=Minute(Now)%><%=Second(Now)%>)"/>
           <postfield name="Privacy" value="$(Privacy)"/>
           <postfield name="Sign" value="$(Sign<%=Minute(Now)%><%=Second(Now)%>)"/>
           </go></anchor>
           <br/>
           <%
	   End Sub
	   
	   '联系信息
	   Sub ContactInfo()
		   Dim RSU:Set RSU=Server.CreateObject("ADODB.RECORDSET")
		   RSU.Open "Select * From KS_User Where UserName='" & KSUser.UserName & "'",Conn,1,1
		   If RSU.Eof Then
		      RSU.Close:Set RSU=Nothing
			  Response.Write "非法参数！<br/>":Prev=True:Exit Sub
		   End If
		   Dim Template:Template=LFCls.GetSingleFieldValue("Select WapTemplate From KS_UserForm Where ID=(Select FormID From KS_UserGroup Where id=" & KSUser.GroupID & ")")
		   Dim FieldsList:FieldsList=LFCls.GetSingleFieldValue("Select FormField From KS_UserForm Where ID=(Select FormID From KS_UserGroup Where ID=" & KSUser.GroupID&")")
		   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		   RS.Open "Select FieldID,FieldType,FieldName,DefaultValue,Width,Height,Options,EditorType from KS_Field Where ChannelID=101 Order By OrderID",Conn,1,1
		   Dim SQL,K,N,InputStr,O_Arr,O_Len,F_V,O_Value,O_Text,BRStr,FieldStr
		   Dim postfield
		   If Not RS.EOf Then SQL=RS.GetRows(-1):RS.Close():Set RS=Nothing
		   For K=0 TO Ubound(SQL,2)
		       FieldStr=FieldStr & "|" & lcase(SQL(2,K))
			   If KS.FoundInArr(FieldsList,SQL(0,k),",") Then
			      InputStr=""
				  If lcase(SQL(2,K))="province&city" Then

				  Else
				     Select Case SQL(1,K)
					     Case 2
						 InputStr="<input type=""text"" name=""" & SQL(2,K) & Minute(Now)& Second(Now) & """ value=""" &RSU(SQL(2,K)) & """/>"
						 PostField=PostField&"<postfield name=""" & lcase(SQL(2,K)) & """ value=""$(" & lcase(SQL(2,K)) & Minute(Now)& Second(Now) & ")""/>" & vbCrLf
						 Case 3
						 InputStr="<select name=""" & Lcase(SQL(2,K)) & """>"
						 O_Arr=Split(SQL(6,K),vbcrlf): O_Len=Ubound(O_Arr)
						 For N=0 To O_Len
						     F_V=Split(O_Arr(N),"|")
							 If Ubound(F_V)=1 Then
							    O_Value=F_V(0):O_Text=F_V(1)
							 Else
							    O_Value=F_V(0):O_Text=F_V(0)
							 End If						   
							 InputStr=InputStr & "<option value=""" & O_Value& """>" & O_Text & "</option>"
						 Next
						 InputStr=InputStr & "</select>"
						 PostField=PostField&"<postfield name=""" & lcase(SQL(2,K)) & """ value=""$(" & lcase(SQL(2,K)) & ")""/>" & vbCrLf
						 Case 6
						 InputStr="<select name=""" & Lcase(SQL(2,K)) & """>"
						 O_Arr=Split(SQL(6,K),vbcrlf): O_Len=Ubound(O_Arr)
						 For N=0 To O_Len
						     F_V=Split(O_Arr(N),"|")
							 If Ubound(F_V)=1 Then
							    O_Value=F_V(0):O_Text=F_V(1)
							 Else
							    O_Value=F_V(0):O_Text=F_V(0)
							 End If
							 InputStr=InputStr & "<option value=""" & O_Value& """>" & O_Text & "</option>"
						 Next
						 InputStr=InputStr & "</select>"
						 PostField=PostField&"<postfield name=""" & lcase(SQL(2,K)) & """ value=""$(" & lcase(SQL(2,K)) & ")""/>" & vbCrLf
						 Case 7
						 InputStr="<select name=""" & Lcase(SQL(2,K)) & """>"
						 O_Arr=Split(SQL(6,K),vbcrlf): O_Len=Ubound(O_Arr)
						 For N=0 To O_Len
						     F_V=Split(O_Arr(N),"|")
							 If Ubound(F_V)=1 Then
							    O_Value=F_V(0):O_Text=F_V(1)
							 Else
							    O_Value=F_V(0):O_Text=F_V(0)
							 End If
							 InputStr=InputStr & "<option value=""" & O_Value& """>" & O_Text & "</option>"
						 Next
						 InputStr=InputStr & "</select>"
						 PostField=PostField&"<postfield name=""" & Lcase(SQL(2,K)) & """ value=""$(" & Lcase(SQL(2,K)) & ")""/>" & vbCrLf
						 Case 10
						 Dim H_Value:H_Value=RSU(SQL(2,K))
						 If IsNull(H_Value) Then H_Value=" "
						 InputStr=InputStr & "<input type=""text"" name=""" & Lcase(SQL(2,K)) & Minute(Now)& Second(Now) & """ value="""& Server.HTMLEncode(H_Value) &"""/>"
						 PostField=PostField&"<postfield name=""" & Lcase(SQL(2,K)) & """ value=""$(" & Lcase(SQL(2,K)) & Minute(Now)& Second(Now) & ")""/>" & vbCrLf
						 Case Else
						 InputStr="<input type=""text"" name=""" & Lcase(SQL(2,K)) & Minute(Now)& Second(Now) & """ value=""" & RSU(SQL(2,K)) & """/>"
						 PostField=PostField&"<postfield name=""" & Lcase(SQL(2,K)) & """ value=""$(" & Lcase(SQL(2,K)) & Minute(Now)& Second(Now) & ")""/>" & vbCrLf
					 End Select
				  End If
				  'If SQL(1,K)=9 Then InputStr=InputStr & "<div><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_UpFile.asp?Type=Field&FieldID=" & SQL(0,K) & "&ChannelID=101' frameborder=0 scrolling=no width='100%' height='26'></iframe></div>"
				  Template=Replace(Template,"[@" & SQL(2,K) & "]",InputStr)
			   End If
		   Next
		   RSU.Close:Set RSU=Nothing
		   Response.Write Template
		   Response.Write "<anchor>确定修改<go href=""User_EditInfo.asp?Action=ContactInfoSave&amp;"&KS.WapValue&""" method=""post"">"
		   Response.Write postfield
		   Response.Write "</go></anchor><br/>"
	   End Sub
	

	   Sub PassQuestionInfo()
	       %>
           找回密码设置<br/>
           登录密码:<input name="Password<%=minute(now)%><%=second(now)%>" type="password" value=""/><br/>
           密码问题:<input name="Question<%=minute(now)%><%=second(now)%>" type="text" value="<%=KSUser.Question%>"/><br/>
           问题答案:<input name="Answer<%=minute(now)%><%=second(now)%>" type="text" value="<%=KSUser.Answer%>"/><br/>
           <anchor>确定修改<go href="User_EditInfo.asp?Action=PassQuestionSave&amp;<%=KS.WapValue%>" method="post">
           <postfield name="Password" value="$(Password<%=minute(now)%><%=second(now)%>)"/>
           <postfield name="Question" value="$(Question)"/>
           <postfield name="Answer" value="$(Answer<%=minute(now)%><%=second(now)%>)"/>
           </go></anchor><br/>
           <%
	   End Sub
	   
	   '设置密码
	   Sub PassInfo()
  		   %>
           【修改密码】<br/>
           旧 密 码:您的旧登录密码,必须正确填写.<br/>
           <input name="oldpassword<%=Minute(Now)%><%=Second(Now)%>" type="password" value=""/><br/>
           新 密 码:请输入您的新密码!<br/>
           <input name="newpassword<%=Minute(Now)%><%=Second(Now)%>" type="password" value=""/><br/>
           确认密码:同上.<br/>
           <input name="renewpassword<%=Minute(Now)%><%=Second(Now)%>" type="password" value=""/><br/>
           <anchor>确定修改<go href="User_EditInfo.asp?Action=PassSave&amp;<%=KS.WapValue%>" method="post" accept-charset="utf-8">
           <postfield name="oldpassword" value="$(oldpassword<%=Minute(Now)%><%=Second(Now)%>)"/>
           <postfield name="newpassword" value="$(newpassword<%=Minute(Now)%><%=Second(Now)%>)"/>
           <postfield name="renewpassword" value="$(renewpassword<%=Minute(Now)%><%=Second(Now)%>)"/>
           </go></anchor><br/>
		   <%
	   End Sub
	   
	   Sub BasicInfoSave() 
	       Dim RealName:RealName=KS.S("RealName")
		   Dim Sex:Sex=KS.S("Sex")
		   Dim Birthday:Birthday=KS.S("Birthday")
		   Dim IDCard:IDCard=KS.S("IDCard")
		   'Dim UserFace:UserFace=KS.S("UserFace")		 
		   'Dim FaceWidth:FaceWidth=KS.S("FaceWidth")		 
		   'Dim FaceHeight:FaceHeight=KS.S("FaceHeight")		 
		   Dim Sign:Sign=KS.S("Sign")	
		   Dim Privacy:Privacy=KS.S("Privacy")
		   If Not IsDate(Birthday) Then
		      Response.Write "出生日期格式有误!<br/>":Prev=True:Exit Sub
		   End If
		   Dim Email:Email=KS.S("Email")
		   'If KS.IsValidEmail(Email)=false Then
		      'Response.Write "请输入正确的电子邮箱!<br/>":Prev=True:Exit Sub
		   'End If
		   'Dim EmailMultiRegTF:EmailMultiRegTF=KS.ChkClng(KS.Setting(28))
		   'If EmailMultiRegTF=0 Then
		      'Dim EmailRSCheck:Set EmailRSCheck = Conn.Execute("select UserID from KS_User where UserName<>'" & KSUser.UserName & "' And Email='" & Email & "'")
			  'If Not (EmailRSCheck.BOF And EmailRSCheck.EOF) Then
		         'EmailRSCheck.Close:Set EmailRSCheck = Nothing
		         'Response.Write "您注册的Email已经存在！请更换Email再试试！<br/>":Prev=True:Exit Sub
			  'End If
			  'EmailRSCheck.Close:Set EmailRSCheck = Nothing
		   'End If
		   Dim RS: Set RS=Server.CreateObject("Adodb.RecordSet")
		   RS.Open "Select top 1 * From KS_User Where UserName='" & KSUser.UserName & "'",Conn,1,3
		   IF RS.Eof And RS.Bof Then
		      RS.Close:Set RS=Nothing
			  Response.Write "请指定正确的参数。<br/>":Prev=True:Exit Sub
		   Else
		     ' RS("UserFace")=UserFace
			  RS("RealName")=RealName
			  RS("Sex")=Sex
			  RS("Birthday")=Birthday
			  RS("IDCard")=IDCard
			  'RS("UserFace")=UserFace
			  'RS("FaceWidth")=FaceWidth
			  'RS("FaceHeight")=FaceHeight
			  RS("Email")=Email
			  RS("Sign")=Sign
			  RS("Privacy")=Privacy
			  RS.Update
			  Response.Write "会员基本信息资料修改成功！<br/>"
		   End if
		   RS.Close:Set RS=Nothing
	   End Sub
	   
	   '保存联系信息
	   Sub ContactInfoSave()
	       Dim SQL,K
		   Dim FieldsList:FieldsList=LFCls.GetSingleFieldValue("Select FormField From KS_UserForm Where ID=(Select FormID From KS_UserGroup Where ID=" & KSUser.GroupID&")")
		   If FieldsList="" Then FieldsList="0"
		   Set RS = Server.CreateObject("ADODB.RECORDSET")
		   RS.Open "Select FieldName,MustFillTF,Title,FieldType From KS_Field Where ChannelID=101 and FieldID In(" & KS.FilterIDs(FieldsList) & ")",conn,1,1
		   If Not RS.Eof Then SQL=RS.GetRows(-1)
		   RS.Close
		   For K=0 To UBound(SQL,2)
		         If SQL(1,K)="1" Then 
				    If SQL(0,K)="Province&City" Then
					ElseIf KS.S(SQL(0,K))="" Then
				     Response.Write "" & SQL(2,K) & "必须填写!<br/>"
					 Prev=True
					 Exit Sub
					End If
				  End If
				  If SQL(3,K)="4" And Not Isnumeric(KS.S(SQL(0,K))) Then
				     Response.Write "" & SQL(2,K) & "必须填写数字!<br/>"
					 Prev=True
					 Exit Sub
				  End If
				  If SQL(3,K)="5" And Not IsDate(KS.S(SQL(0,K))) Then 
				     Response.Write "" & SQL(2,K) & "必须填写正确的日期!<br/>"
					 Prev=True
					 Exit Sub
				  End If
				  If SQL(3,K)="8" And Not KS.IsValidEmail(KS.S(SQL(0,K))) and SQL(1,K)="1" Then 
				     Response.Write "" & SQL(2,K) & "必须填写正确的Email格式!<br/>"
					 Prev=True
					 Exit Sub
				  End If 
			  Next
			  Dim RealName:RealName=KS.S("RealName")
			  Dim Sex:Sex=KS.S("Sex")
			  Dim Birthday:Birthday=KS.S("Birthday")
			  Dim IDCard:IDCard=KS.S("IDCard")
			  Dim OfficeTel:OfficeTel=KS.S("OfficeTel")
			  Dim HomeTel:HomeTel=KS.S("HomeTel")
			  Dim Mobile:Mobile=KS.S("Mobile")
			  Dim Fax:Fax=KS.S("Fax")
			  Dim province:province=KS.S("prov")
			  Dim city:city=KS.S("city")
			  Dim Address:Address=KS.S("Address")
			  Dim ZIP:ZIP=KS.S("ZIP")
			  Dim HomePage:HomePage=KS.S("HomePage")		 	 	 
			  Dim QQ:QQ=KS.S("QQ")		 
			  Dim ICQ:ICQ=KS.S("ICQ")		 
			  Dim MSN:MSN=KS.S("MSN")		 
			  Dim UC:UC=KS.S("UC")		 
			  Dim Sign:Sign=KS.S("Sign")	
			  Dim Privacy:Privacy=KS.ChkClng(KS.S("Privacy"))
			 
              Dim RS:Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select top 1 * From KS_User Where UserName='" & KSUser.UserName & "'",Conn,1,3
			  IF RS.Eof And RS.Bof Then
			     Response.Write "非法参数！<br/>"
				 Prev=True
				 Exit Sub
			  Else
			     RS("Sex")=Sex
				 If BirthDay<>"" Then RS("Birthday")=Birthday
				 RS("RealName")=RealName
				 RS("IDCard")=IDCard
				 RS("Email")=KSUser.Email
				 RS("OfficeTel")=OfficeTel
				 RS("HomeTel")=HomeTel
				 RS("Mobile")=Mobile
				 RS("Fax")=Fax
				 RS("Province")=Province
				 RS("City")=City
				 RS("Address")=Address
				 RS("Zip")=Zip
				 RS("HomePage")=HomePage
				 RS("QQ")=QQ
				 RS("ICQ")=ICQ
				 RS("MSN")=MSN
				 RS("UC")=UC
				 '自定义字段
				 For K=0 To UBound(SQL,2)
				    If left(Lcase(SQL(0,K)),3)="ks_" Then
					   RS(SQL(0,K))=KS.S(SQL(0,K))
					End If
				 Next
		 		 RS.Update
				 Response.Write "恭喜，详细信息修改成功！<br/>"
			  End if
			  RS.Close:Set RS=Nothing
	   End Sub
  
	   
	   '保存密码设置
	   Sub PassSave()
	       Dim Oldpassword:Oldpassword=KS.R(KS.S("Oldpassword"))
		   Dim NewPassWord:NewPassWord=KS.R(KS.S("NewPassWord"))
		   Dim ReNewPassWord:ReNewPassWord=KS.S("ReNewPassWord")
		   If Oldpassword = "" Then
		      Response.Write "请输入旧登录密码!<br/>":Prev=True:Exit Sub
		   End If
		   If NewPassWord = "" Then
		      Response.Write "请输入登录密码!<br/>":Prev=True:Exit Sub
		   ElseIF ReNewPassWord="" Then
		      Response.Write "请输入确认密码!<br/>":Prev=True:Exit Sub
		   ElseIF NewPassWord<>ReNewPassWord Then
		      Response.Write "两次输入的密码不一致!<br/>":Prev=True:Exit Sub
		   End If
		   OldPassWord =MD5(OldPassWord,16)
		   NewPassWord =MD5(NewPassWord,16)
		   Dim RS:Set RS=Server.CreateObject("Adodb.RecordSet")
		   RS.Open "Select PassWord From KS_User Where UserName='" & KSUser.UserName & "' And PassWord='" & OldPassWord & "'",Conn,1,3
		   IF RS.Eof And RS.Bof Then
		      Response.Write "您输入的旧密码有误！<br/>":Prev=True:Exit Sub
		   Else
		      RS(0)=NewPassWord
			  RS.Update
		   End if
		   RS.Close:Set RS=Nothing
		   Response.Write "您的会员登录密码修改成功！新密码"&KS.R(KS.S("NewPassWord"))&"请牢记。<br/>"
		   Response.Write "<a href=""Index.asp?"&KS.WapValue&""">进入会员首页</a><br/>"
		   Response.Write "<a href=""UserLogout.asp"">退出重新登录</a><br/>"
	   End Sub
	   
	   '提示问题保存
	   Sub PassQuestionSave()
	       Dim PassWord:PassWord=KS.S("PassWord")
		   Dim Question:Question=KS.S("Question")
		   Dim Answer:Answer=KS.S("Answer")
		   
		   PassWord=MD5(PassWord,16)
		   Dim RS: Set RS=Server.CreateObject("Adodb.RecordSet")
		   RS.Open "Select * From KS_User Where UserName='" & KSUser.UserName & "' And PassWord='" & PassWord & "'",Conn,1,3
		   IF RS.Eof And RS.Bof Then
		      RS.Close:Set RS=Nothing
			  Response.Write "您输入的登录密码不正确!<br/>":Prev=True:Exit Sub
		   Else
		      RS("Question")=Question
			  RS("Answer")=Answer
			  RS.Update
			  Response.Write "你的密码找回资料修改成功！<br/>"
		   End if
		   RS.Close:Set RS=Nothing
	  End Sub
End Class
%> 
