<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../../Plus/md5.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New UserReg
KSCls.Kesion()
Set KSCls = Nothing

Class UserReg
        Private KS, KSRFObj,FileContent,RegAnswerID,rndReg
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSRFObj = New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		   IF KS.Setting(21)=0 Then
		    Response.Redirect "../../plus/error.asp?action=error&message=" & Server.URLEncode("<li>对不起，本站暂停新会员注册!</li>")
		    Response.End
		   End IF
		     If KS.S("Action")="Next" Then
			  Call Step2()
			 Else
			  Call Step1()
			 End If
		End Sub

        '注册会员第一步
		Sub Step1()
		  If KS.Setting(117)="" Then Response.Write "请先到""基本信息设置->模板绑定""进行模板绑定操作!":response.end
		   FileContent = KSRFObj.LoadTemplate(KS.Setting(117))
		   FCls.RefreshType="UserRegStep1"
		   If Trim(FileContent) = "" Then FileContent = "模板不存在!"
		   FileContent = ReplaceRegLable(FileContent)        '替换通用标签 如{$GetWebmaster}
           FileContent = KSRFObj.KSLabelReplaceAll(FileContent) '替换函数标签
		   Response.Write FileContent  
		End Sub
		'注册会员第二步
		Sub Step2()
		  	Dim SQL,K,N,InputStr,O_Arr,O_Len,F_V,O_Value,O_Text,BRStr,CanReg

		  Dim Verifycode:Verifycode=KS.S("Verifycode")
		   IF Trim(Verifycode)<>Trim(Session("Verifycode")) And KS.Setting(27)=1 then 
		   	 Response.Write("<script>alert('验证码有误，请重新输入！');history.back(-1);</script>")
		     Exit Sub
		   End IF
		   
		   '检查注册回答问题
		   If Mid(KS.Setting(161),1,1)="1" Then
		     CanReg=false
		     For N=0 To Ubound(Split(KS.Setting(162),vbcrlf))
			   If Trim(Request.Form("a" & MD5(n,16)))<>"" Then
			      If trim(Lcase(Request.Form("a" & MD5(n,16))))<>trim(Lcase(Split(KS.Setting(163),vbcrlf)(n))) Then
			       Call KS.AlertHistory("对不起,注册问题的回答不正确!",-1) : Response.End
				   CanReg=false
				  Else
				   RegAnswerID=N
				   CanReg=True
				  End If
			   End If
			 Next
			 If CanReg=false Then Call KS.AlertHistory("对不起,注册答案不能为空!",-1) : Response.End
		   End If
		   
		   If KS.Setting(118)="" Then Response.Write "请先到""基本信息设置->模板绑定""进行模板绑定操作!":response.end
		   FileContent = KSRFObj.LoadTemplate(KS.Setting(118))
		   FCls.RefreshType="UserRegStep2"
		   If Trim(FileContent) = "" Then FileContent = "模板不存在!"
		   ReplaceReg2()       '替换通用标签 如{$GetWebmaster}
           FileContent = KSRFObj.KSLabelReplaceAll(FileContent) '替换函数标签
		   Dim GroupID:GroupID=KS.ChkClng(KS.S("GroupID")):If GroupID=0 Then GroupID=3
		   Dim FieldsList:FieldsList=LFCls.GetSingleFieldValue("Select FormField From KS_UserForm Where ID=" & KS.U_G(GroupID,"formid"))
		   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		   RS.Open "Select FieldID,FieldType,FieldName,DefaultValue,Width,Height,Options,EditorType from KS_Field Where ChannelID=101 Order By OrderID",conn,1,1
		   If Not RS.EOf Then SQL=RS.GetRows(-1):RS.Close():Set RS=Nothing
		   For K=0 TO Ubound(SQL,2)
		     If KS.FoundInArr(FieldsList,SQL(0,k),",") Then
			  InputStr=""
			  If lcase(SQL(2,K))="province&city" Then
				 InputStr="<script language=""javascript"" src=""../../plus/area.asp""></script>"
			  Else
			  Select Case SQL(1,K)
			    Case 2:InputStr="<textarea style=""width:" & SQL(4,K) & "px;height:" & SQL(5,K) & "px"" rows=""5"" class=""textbox"" name=""" & SQL(2,K) & """>" & SQL(3,K) & "</textarea>"
				Case 3
				  InputStr="<select style=""width:" & SQL(4,K) & """ name=""" & SQL(2,K) & """>"
				  O_Arr=Split(SQL(6,K),vbcrlf): O_Len=Ubound(O_Arr)
				  For N=0 To O_Len
					 F_V=Split(O_Arr(N),"|")
					 If Ubound(F_V)=1 Then
						O_Value=F_V(0):O_Text=F_V(1)
					 Else
						O_Value=F_V(0):O_Text=F_V(0)
					 End If						   
					 If SQL(3,K)=O_Value Then
						InputStr=InputStr & "<option value=""" & O_Value& """ selected>" & O_Text & "</option>"
					 Else
						InputStr=InputStr & "<option value=""" & O_Value& """>" & O_Text & "</option>"
					 End If
				  Next
					InputStr=InputStr & "</select>"
				Case 6
					 O_Arr=Split(SQL(6,K),vbcrlf): O_Len=Ubound(O_Arr)
					 If O_Len>1 And Len(SQL(6,K))>50 Then BrStr="<br>" Else BrStr=""
					 For N=0 To O_Len
						F_V=Split(O_Arr(N),"|")
						If Ubound(F_V)=1 Then
						 O_Value=F_V(0):O_Text=F_V(1)
						Else
						 O_Value=F_V(0):O_Text=F_V(0)
						End If						   
					    If SQL(3,K)=O_Value Then
							InputStr=InputStr & "<input type=""radio"" name=""" & SQL(2,K) & """ value=""" & O_Value& """ checked>" & O_Text & BRStr
						Else
							InputStr=InputStr & "<input type=""radio"" name=""" & SQL(2,K) & """ value=""" & O_Value& """>" & O_Text & BRStr
						 End If
			         Next
			  Case 7
					O_Arr=Split(SQL(6,K),vbcrlf): O_Len=Ubound(O_Arr)
					 For N=0 To O_Len
						  F_V=Split(O_Arr(N),"|")
						  If Ubound(F_V)=1 Then
							O_Value=F_V(0):O_Text=F_V(1)
						  Else
							O_Value=F_V(0):O_Text=F_V(0)
						  End If						   
						  If KS.FoundInArr(SQL(3,K),O_Value,",")=true Then
								 InputStr=InputStr & "<input type=""checkbox"" name=""" & SQL(2,K) & """ value=""" & O_Value& """ checked>" & O_Text
						 Else
						  InputStr=InputStr & "<input type=""checkbox"" name=""" & SQL(2,K) & """ value=""" & O_Value& """>" & O_Text
						 End If
				   Next
			  Case 10
					InputStr=InputStr & "<input type=""hidden"" id=""" & SQL(2,K) &""" name=""" & SQL(2,K) &""" value="""& Server.HTMLEncode(SQL(3,K)) &""" style=""display:none"" /><iframe id=""" & SQL(2,K) &"___Frame"" src=""../../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=" & SQL(2,K) &"&amp;Toolbar=" & SQL(7,K) & """ width=""" &SQL(4,K) &""" height=""" & SQL(5,K) & """ frameborder=""0"" scrolling=""no""></iframe>"				
			  Case Else
			    If KS.Setting(149)="1" and lcase(SQL(2,K))="mobile" Then
			  InputStr="<input type=""text"" class=""textbox"" readonly style=""width:" & SQL(4,K) & "px"" name=""" & SQL(2,K) & "1"" value=""" & KS.S("Mobile") & """>"
				Else
			  InputStr="<input type=""text"" class=""textbox"" style=""width:" & SQL(4,K) & "px"" name=""" & SQL(2,K) & """ value=""" & SQL(3,K) & """>"
			    End If
			  End Select
			  End If
			  'if SQL(1,K)=9 Then InputStr=InputStr & "<div><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_UpFile.asp?UPType=Field&FieldID=" & SQL(2,K) & "&ChannelID=101' frameborder=0 scrolling=no width='100%' height='26'></iframe></div>"
			  If Instr(FileContent,"{@NoDisplay(" & SQL(2,K) & ")}")<>0 Then
			   FileContent=Replace(FileContent,"{@NoDisplay(" & SQL(2,K) & ")}"," style='display:none'")
			  End If
			   FileContent=Replace(FileContent,"[@" & SQL(2,K) & "]",InputStr)
			  End If
		   Next
		    FileContent=Replace(FileContent,"{@NoDisplay}","")
			
           FileContent = KSRFObj.ReplaceRA(FileContent, "")
		   Response.Write KSRFObj.ReplaceGeneralLabelContent(FileContent)  
		End Sub
 

	   Function ReplaceRegLable(FileContent)
		 Dim UserRegMustFill:UserRegMustFill=KS.Setting(33)
		 Dim ShowCheckEmailTF:ShowCheckEmailTF=true
		 Dim ShowVerifyCodeTF:ShowVerifyCodeTF=false
		 
		 IF KS.Setting(28)="1" Then ShowCheckEmailTF=false
		 IF KS.Setting(27)="1" then ShowVerifyCodeTF=true
		 
		 If KS.Setting(33)="0" Then
		 FileContent = Replace(FileContent, "{$ShowUserType}", "")
		 FileContent = Replace(FileContent, "{$DisplayUserType}", " style='display:none'")
		 Else
		 FileContent = Replace(FileContent, "{$ShowUserType}", UserGroupList())
		 FileContent = Replace(FileContent, "{$DisplayUserType}", "")
		 End If
		 
		 If KS.Setting(32)="1" Then 
		 FileContent = Replace(FileContent, "{$ShowAction}", "UserRegResult.asp")
		 Else
		 FileContent = Replace(FileContent, "{$ShowAction}", "index.asp?Action=Next")
		 End If
		 
		 If KS.Setting(148)="1" Then
		 FileContent = Replace(FileContent, "{$DisplayQestion}", "")
		 Else
		 FileContent = Replace(FileContent, "{$DisplayQestion}", " style=""display:none""")
		 End If

		 If KS.Setting(149)="1" Then
		 FileContent = Replace(FileContent, "{$DisplayMobile}", "")
		 Else
		 FileContent = Replace(FileContent, "{$DisplayMobile}", " style=""display:none""")
		 End If
		 If KS.Setting(143)="1" Then
		 FileContent = Replace(FileContent, "{$DisplayAlliance}", "")
		 Else
		 FileContent = Replace(FileContent, "{$DisplayAlliance}", " style=""display:none""")
		 End If
		 
		 If Mid(KS.Setting(161),1,1)="1" Then
		 rndReg=GetRegRnd()
		 FileContent = Replace(FileContent, "{$DisplayRegQuestion}", "")
		 FileContent = Replace(FileContent, "{$RegQuestion}", GetRegQuestion)
		 FileContent = Replace(FileContent, "{$AnswerRnd}", GetRegAnswerRnd)
		 Else
		 FileContent = Replace(FileContent, "{$DisplayRegQuestion}", " style=""display:none""")
		 FileContent = Replace(FileContent, "{$RegQuestion}", "")
		 FileContent = Replace(FileContent, "{$AnswerRnd}", "")
		 End If
		 
		 FileContent = Replace(FileContent, "{$Show_Question}", KS.Setting(148))
		 FileContent = Replace(FileContent, "{$Show_Mobile}", KS.Setting(149))
		 If Request("u")<>"" Then
		 FileContent = Replace(FileContent, "{$UserName}", " value=""" & split(Request("u"),"@")(0) & """")
		 Else
		 FileContent = Replace(FileContent, "{$UserName}", "")
		 End If
		 If KS.S("Uid")<>"" Then
		  FileContent = Replace(FileContent, "{$AllianceUser}", " value=""" & KS.S("Uid") & """ readonly")
		  FileContent = Replace(FileContent, "{$Friend}", " value=""" & KS.S("F") & """")
		 Else
		  FileContent = Replace(FileContent, "{$AllianceUser}", "")
		  FileContent = Replace(FileContent, "{$Friend}", "")
		 End If

		 FileContent = Replace(FileContent, "{$GetUserRegLicense}", KS.Setting(23))
		 FileContent=Replace(FileContent,"{$Show_UserNameLimitChar}",KS.Setting(29))
		 FileContent=Replace(FileContent,"{$Show_UserNameMaxChar}",KS.Setting(30))
		 'FileContent=Replace(FileContent,"{$Show_VerifyCode}","<IMG style=""cursor:pointer"" src=""{$GetSiteUrl}plus/verifycode.asp?n=" & Timer & """ onClick=""this.src='" & KS.GetDomain & "plus/verifycode.asp?n='+ Math.random();""  align=""absmiddle"">")
		  FileContent = Replace(FileContent, "{$Show_CheckEmail}", IsShow(ShowCheckEmailTF))
		  FileContent = Replace(FileContent, "{$Show_VerifyCodeTF}", IsShow(ShowVerifyCodeTF))
	

		 ReplaceRegLable=KSRFObj.ReplaceGeneralLabelContent(FileContent)
		End Function
		
		Function GetRegRnd()
		  Dim QuestionArr:QuestionArr=Split(KS.Setting(162),vbcrlf)
		  Dim RandNum,N: N=Ubound(QuestionArr)
          Randomize
          RandNum=Int(Rnd()*N)
          GetRegRnd=RandNum
		End Function
		Function GetRegQuestion()
		  Dim QuestionArr:QuestionArr=Split(KS.Setting(162),vbcrlf)
		  GetRegQuestion=QuestionArr(rndReg)
		End Function
		Function GetRegAnswerRnd()
		  GetRegAnswerRnd=md5(rndReg,16)
		End Function
		
		
		Sub ReplaceReg2()
		  If Request.ServerVariables("HTTP_REFERER")="" Then Call KS.Alert("请不要非法提交!","../"):Response.End
		  If Instr(Lcase(Request.ServerVariables("SCRIPT_NAME")),"user/reg")=0 Then Call KS.Alert("请不要非法提交!","../") : Response.End
		
		  Dim GroupID:GroupID=KS.ChkCLng(KS.S("GroupID")):If GroupID=0 Then GroupID=3
		  Dim Template:Template=LFCls.GetSingleFieldValue("Select Template From KS_UserForm Where ID=" & KS.U_G(GroupID,"formid"))
		  Template=Template & "<input type='hidden' name='f' id='f' value='" & KS.S("f") &"'>" & vbcrlf
		  Template=Template & "<input type='hidden' name='GroupID' value='" & KS.S("GroupID") &"'>" & vbcrlf
		  Template=Template & "<input type='hidden' name='UserName' value='" & KS.S("UserName") &"'>" & vbcrlf
		  Template=Template & "<input type='hidden' name='Question' value='" & KS.S("Question") &"'>" & vbcrlf
		  Template=Template & "<input type='hidden' name='Answer' value='" & KS.S("Answer") &"'>" & vbcrlf
		  Template=Template & "<input type='hidden' name='Email' value='" & KS.S("Email") &"'>" & vbcrlf
		  Template=Template & "<input type='hidden' name='a" & MD5(RegAnswerID,16) & "' value='" & Request.Form("a" & MD5(RegAnswerID,16)) &"'>" & vbcrlf
		  
		  If KS.Setting(149)="1" Then
		  Template=Template & "<input type='hidden' name='Mobile' value='" & KS.S("Mobile") &"'>" & vbcrlf
		  End If
		  '=======================增加加盟号=================================
		  Dim AllianceUser:AllianceUser=KS.S("AllianceUser")
		  If AllianceUser<>"" Then
		     If AllianceUser=KS.S("UserName") Then
			  Call KS.AlertHistory("对不起,推荐人不能是自己!",-1)
			  Exit Sub
			 End If
		'	  If Conn.Execute("Select UserName From KS_User Where UserName='" & AllianceUser & "'").Eof Then
		'		Response.Write "<script>alert('对不起，您输入的加盟号不存在！');history.back();<//script>"
		'		Response.End
		'	  End If
		  End If
		  Template=Template & "<input type='hidden' name='AllianceUser' value='" & AllianceUser &"'>" & vbcrlf
		  '=======================加盟号结束====================================
		  
		  FileContent=Replace(FileContent,"{$ShowRegForm}",Template)
			Dim PassWord:PassWord=KS.R(KS.S("PassWord"))
			Dim RePassWord:RePassWord=KS.S("RePassWord")
			If PassWord = "" Then
				 Response.Write("<script>alert('请输入登录密码!');history.back();</script>")
				 Exit Sub
			ElseIF RePassWord="" Then
				 Response.Write("<script>alert('请输入确认密码');history.back();</script>")
				 Exit Sub
			ElseIF PassWord<>RePassWord Then
				 Response.Write("<script>alert('两次输入的密码不一致');history.back();</script>")
				 Exit Sub
			End If
		  Session("PassWord")=PassWord
		End Sub
		'会员类型
		Function UserGroupList()
			 If  KS.Setting(33)="0" Then UserGroupList="":Exit Function
			Dim RS,Node
			 Call KS.LoadUserGroup()
			 If KS.ChkClng(KS.S("GroupID"))<>0 Then
				Set Node=Application(KS.SiteSN&"_UserGroup").DocumentElement.SelectSingleNode("row[@id=" & KS.S("GroupID") & "]/@groupname")
				If Not Node Is Nothing Then
				UserGroupList="<span style='font-weight:bold;color:#ff6600'>" & Node.Text &"</span><input type='hidden' value='" & KS.S("GroupID") & "' name='GroupID'>"
			    End If 
				Set Node=Nothing
			Else
			  For Each Node In Application(KS.SiteSN&"_UserGroup").DocumentElement.SelectNodes("row[@showonreg=1 && @id!=1]")
			  If UserGroupList="" Then
			  UserGroupList="<label><input type=""radio""  value=""" & Node.SelectSingleNode("@id").text & """ name=""GroupID"" checked>" & Node.SelectSingleNode("@groupname").text  & "</label>(<span onclick=""mousepopup(event,'说明','" & Replace(Replace(Replace(Node.SelectSingleNode("@descript").text,"'","\'"),vbcrlf,"\n"),chr(10),"\n") & "',300)"" style=""cursor:default;color:red;text-decoration:underline"">说明</span>)"
			  Else
			  UserGroupList=UserGroupList & "<br /><label><input type=""radio""  value=""" & Node.SelectSingleNode("@id").text & """ name=""GroupID"">" & Node.SelectSingleNode("@groupname").text & "</label>(<span onclick=""mousepopup(event,'说明','" & Replace(Replace(Replace(Node.SelectSingleNode("@descript").text,"'","\'"),vbcrlf,"\n"),chr(10),"\n") & "',300)"" style=""cursor:default;color:red;text-decoration:underline"">说明</span>)"
			  End If
			 Next
			End If
		End Function
		
		Function IsShow(Show)
			If Show =true Then
				IsShow = ""
			Else
				IsShow = " Style='display:none'"
			End If
		End Function
End Class
%>

 
