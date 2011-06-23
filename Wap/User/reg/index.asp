<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.ContentType="text/vnd.wap.wml"
Response.Charset="utf-8"
Response.Write "<?xml version=""1.0"" encoding=""utf-8""?>" &vbcrlf
Response.Write "<!DOCTYPE wml PUBLIC ""-//WAPFORUM//DTD WML 1.1//EN"" ""http://www.wapforum.org/DTD/wml_1.1.xml"">" &vbcrlf
%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
Dim KSCls
Set KSCls = New UserReg
KSCls.Kesion()
Set KSCls = Nothing

Class UserReg
        Private KS
		Private FileContent,Prev
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		    Call CloseConn()
		    Set KS=Nothing
		End Sub
		Public Sub Kesion()
		    IF KS.Setting(21)=0 Then
			   Call KS.ShowError("暂停注册！","对不起，本站暂停新会员注册！")
		    End IF
		    Response.Write "<wml>" &vbcrlf
			Response.Write "<head>" &vbcrlf
			Response.Write "<meta http-equiv=""Cache-Control"" content=""no-Cache""/>" &vbcrlf
			Response.Write "<meta http-equiv=""Cache-Control"" content=""max-age=0""/>" &vbcrlf
			Response.Write "</head>" &vbcrlf
			Response.Write "<card id=""main"" title=""会员注册"">" &vbcrlf
			Response.Write "<p align=""left"">" &vbcrlf
			
			If KS.S("Action")="Next" Then
			   Call Step2()
			Else
			   Call Step1()
			End If
			If Prev=True Then
			   Response.Write "<anchor>返回上一页<prev/></anchor><br/>"
			End If
			Response.Write "<br/>"
			Response.write "<anchor>返回来源页<prev/></anchor><br/>" &vbcrlf
			Response.Write "</p>" &vbcrlf
			Response.Write "</card>" &vbcrlf
			Response.Write "</wml>"
		End Sub

        '注册会员第一步
		Sub Step1()
		    %>
            【会员注册】<br/>
            <%
			Dim AllianceUser:AllianceUser=KS.S("Uid")
			Dim AllianceUrl:AllianceUrl=KS.URLDecode(Request.ServerVariables("HTTP_REFERER"))
			AllianceUrl = Replace(Replace(AllianceUrl,"&amp;","&"),"&","&amp;")
			Request.ServerVariables("HTTP_REFERER")
			If AllianceUser<>"" Then
			   If Conn.Execute("Select UserName From KS_User Where UserName='" & AllianceUser & "'").EOF Then
			      Response.Write "对不起，您输入的推广Uid号不存在！<br/>"
			   Else
			      Response.Write "会员<b>"&AllianceUser&"</b>推荐你加入...<br/><br/>"
			   End If
			End If
			
			If KS.Setting(33)="0" Then
			Else
			   Response.Write "注册类型:<br/>"&UserGroupList()&"<br/>" &vbcrlf
			End If
			%>
            用户名称:限英文或数字,<%=KS.Setting(29)%>~<%=KS.Setting(30)%>个字符<br/>
            <input name="UserName" type="text" maxlength="<%=KS.Setting(30)%>" value="" emptyok="false"/><br/>
            登陆密码:限英文或数字6个字符以上<br/>
            <input name="PassWord" type="text" maxlength="30" value="" emptyok="false"/><br/>
            <%
			'手机号码
			If KS.Setting(149)="1" Then
			   Response.write "手机号码:<input name=""Mobile"" type=""text"" maxlength=""11"" value="""" emptyok=""false"" format=""*N""/><br/>"
			End If
			'密码问题
			If KS.Setting(148)="1" Then
			   Response.write "提示问题:<input name=""Question"" type=""text"" value="""" emptyok=""false""/><br/>"
			   Response.write "提示答案:<input name=""Answer"" type=""text"" value="""" emptyok=""false""/><br/>"
			End If
			
			'邮箱地址
			'If KS.Setting(146)="1" Then
			   Response.write "邮箱地址:<input name=""Email"" type=""text"" value="""" emptyok=""false""/><br/>"
			'End If
			
			'验证码
			If KS.Setting(27)="1" Then
			   Response.write "验证码:<input name=""verifycode"&Minute(Now)&Second(Now)&""" type=""text"" size=""4"" maxlength=""4"" value="""" emptyok=""false"" format=""*N""/>" & KS.GetVerifyCode & "<br/>"
			End If
			Dim UserRegUrl
			If KS.Setting(32)="1" Then 
			   UserRegUrl="UserRegResult.asp"
			Else
			   UserRegUrl="index.asp?Action=Next"
			End If
			Dim ToUrl
			ToUrl=Request.ServerVariableS("QUERY_STRING")
			ToUrl=Replace(ToUrl,"&amp;","&")
			ToUrl=Replace(ToUrl,"&","&amp;")
			ToUrl=Replace(ToUrl,"UID=" & AllianceUser & "","")
			%>        
            <anchor>快速注册<go href="<%=UserRegUrl%>" method="post">
            <postfield name="GroupID" value="$(GroupID)"/>
            <postfield name="UserName" value="$(UserName)"/>
            <postfield name="PassWord" value="$(PassWord)"/>
            <postfield name="Question" value="$(Question)"/>
            <postfield name="Answer" value="$(Answer)"/>
            <postfield name="Email" value="$(Email)"/>
            <postfield name="Mobile" value="$(Mobile)"/>
            <postfield name="AllianceUser" value="<%=AllianceUser%>"/>
            <postfield name="AllianceUrl" value="<%=AllianceUrl%>"/>
            <postfield name="verifycode" value="$(verifycode<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="ToUrl" value="<%=ToUrl%>"/>
            </go>
            </anchor><br/>
            <%
		End Sub
		
		'注册会员第二步
		Sub Step2()
		    Dim Verifycode:Verifycode=KS.S("Verifycode")
			IF Trim(Verifycode)<>Trim(Session("Verifycode")) And KS.Setting(27)=1 then 
			   Response.Write "验证码有误，请重新输入！<br/>"
			   Prev=True
			   Exit Sub
		    End IF
			Dim AllianceUser:AllianceUser=KS.S("AllianceUser")'加盟号
			Dim PassWord:PassWord=KS.R(KS.S("PassWord"))
			Dim RePassWord:RePassWord=KS.S("RePassWord")
			If PassWord = "" Then
			   Response.Write "请输入登录密码!<br/>"
			   Prev=True
			   Exit Sub
			'ElseIF RePassWord="" Then
			   'Response.Write "请输入确认密码<br/>"
			   'Prev=True
			   'Exit Sub
			'ElseIF PassWord<>RePassWord Then
			   'Response.Write "两次输入的密码不一致<br/>"
			   'Prev=True
			   'Exit Sub
			End If
			
			Dim GroupID:GroupID=KS.ChkClng(KS.S("GroupID")):If GroupID=0 Then GroupID=3
			Dim Template:Template=LFCls.GetSingleFieldValue("Select WapTemplate From KS_UserForm Where ID=(Select FormID From KS_UserGroup Where id=" & GroupID & ")")
			Dim FieldsList:FieldsList=LFCls.GetSingleFieldValue("Select FormField From KS_UserForm Where ID=(Select FormID From KS_UserGroup Where ID=" & GroupID&")")
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select FieldID,FieldType,FieldName,DefaultValue,Width,Height,Options,EditorType from KS_Field Where ChannelID=101 Order By OrderID",conn,1,1
			Dim SQL,K,N,InputStr,O_Arr,O_Len,F_V,O_Value,O_Text,BRStr
			Dim PostField
			If Not RS.EOf Then SQL=RS.GetRows(-1):RS.Close():Set RS=Nothing
			For K=0 TO Ubound(SQL,2)
		        If KS.FoundInArr(FieldsList,SQL(0,k),",") Then
				   InputStr=""
				   If lcase(SQL(2,K))="province&city" Then
				   Else
				      Select Case SQL(1,K)
					      Case 2
						  InputStr="<input type=""text"" name=""" & SQL(2,K) & Minute(Now)& Second(Now) & """ value=""" & SQL(3,K) & """/>"
						  PostField=PostField&"<postfield name=""" & SQL(2,K) & """ value=""$(" & SQL(2,K) & Minute(Now)& Second(Now) & ")""/>" & vbCrLf
						  '=====================================
						  Case 3
						  InputStr="<select name=""" & SQL(2,K) & """>"
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
						  PostField=PostField&"<postfield name=""" & SQL(2,K) & """ value=""$(" & SQL(2,K) & ")""/>" & vbCrLf
						  '=====================================
						  Case 6
						  InputStr="<select name=""" & SQL(2,K) & """>"
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
						  PostField=PostField&"<postfield name=""" & SQL(2,K) & """ value=""$(" & SQL(2,K) & ")""/>" & vbCrLf
						  '=====================================
						  Case 7
						  InputStr="<select name=""" & SQL(2,K) & """>"
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
						  PostField=PostField&"<postfield name=""" & SQL(2,K) & """ value=""$(" & SQL(2,K) & ")""/>" & vbCrLf
						  '=====================================
						  Case 10
						  InputStr=InputStr & "<input type=""text"" name=""" & SQL(2,K) & Minute(Now)& Second(Now) & """ value="""& Server.HTMLEncode(SQL(3,K)) &"""/>"
						  PostField=PostField&"<postfield name=""" & SQL(2,K) & """ value=""$(" & SQL(2,K) & Minute(Now)& Second(Now) & ")""/>" & vbCrLf
						  Case Else
						  If KS.Setting(149)="1" And Lcase(SQL(2,K))="mobile" Then
						     InputStr="<input type=""text"" name=""" & SQL(2,K) & "1"" value=""" & KS.S("Mobile") & """/>"
							 PostField=PostField&"<postfield name=""" & SQL(2,K) & """ value=""$(" & SQL(2,K) & ")""/>" & vbCrLf
						  Else
						     InputStr="<input type=""text"" name=""" & SQL(2,K) & """ value=""" & SQL(3,K) & """/>"
							 PostField=PostField&"<postfield name=""" & SQL(2,K) & """ value=""$(" & SQL(2,K) & ")""/>" & vbCrLf
						  End If
					  End Select
				   End If
				   Template=Replace(Template,"[@" & SQL(2,K) & "]",InputStr)
				End If
			Next
			'FileContent=Replace(FileContent,"{$ShowRegForm}",Template)
			Response.Write Template
		    Response.Write "<anchor>提交注册<go href=""UserRegResult.asp"" method=""post"">"
			Response.Write "<postfield name=""GroupID"" value=""" & KS.S("GroupID") &"""/>" & vbcrlf
			Response.Write "<postfield name=""UserName"" value=""" & KS.S("UserName") &"""/>" & vbcrlf
			Response.Write "<postfield name=""PassWord"" value=""" & PassWord &"""/>" & vbcrlf
			Response.Write "<postfield name=""Question"" value=""" & KS.S("Question") &"""/>" & vbcrlf
			Response.Write "<postfield name=""Answer"" value=""" & KS.S("Answer") &"""/>" & vbcrlf
			Response.Write "<postfield name=""Email"" value=""" & KS.S("Email") &"""/>" & vbcrlf
			Response.Write "<postfield name=""AllianceUser"" value=""" & AllianceUser &"""/>" & vbcrlf
			Response.Write "<postfield name=""AllianceUrl"" value=""" & KS.S("AllianceUrl") &"""/>" & vbcrlf
			Response.Write "<postfield name=""Verifycode"" value=""" & KS.S("Verifycode") &"""/>" & vbcrlf
			Response.Write "<postfield name=""ToUrl"" value=""" & Replace(KS.S("ToUrl"),"&","&amp;") &"""/>" & vbcrlf
		    Response.Write PostField
		    Response.Write "</go></anchor><br/>"
		End Sub

		'会员类型
		Function UserGroupList()
		    If  KS.Setting(33)="0" Then UserGroupList="":Exit Function
			Dim RS,Node
			 Call KS.LoadUserGroup()
               
			   UserGroupList="<select name=""GroupID"">"
			  For Each Node In Application(KS.SiteSN&"_UserGroup").DocumentElement.SelectNodes("row[@showonreg=1 && @id!=1]")
			  UserGroupList=UserGroupList & "<option value=""" & Node.SelectSingleNode("@id").text  & """>" & Node.SelectSingleNode("@groupname").text &"</option>"
			 Next
			UserGroupList=UserGroupList & "</select>"
		End Function
End Class
%>

 
