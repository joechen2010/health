<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.MemberCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New User_MyMovie
KSCls.Kesion()
Set KSCls = Nothing

Class User_MyMovie
        Private KS,KSUser,RS
		Private CurrentOpStr,Action,ID
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		Call KSUser.Head()
		Call KSUser.InnerLocation("重发激活码")
		If KS.S("Action")="Send" Then 
		 Call Send()
		Else
		%>
		<br><br><script language = "JavaScript">
				function CheckForm()
				{
				if ($("#UserName").val()=="")
				  {
					alert("请输入用户名！");
					$("#UserName").focus();
					return false;
				  }
				if ($("#Email").val()=="")
				  {
					alert("请输入邮箱！");
					$("#Email").val();
					return false;
				  }
	              return true;
				  }
				</script>
				  <table width="96%"  border="0" align="center" cellpadding="0" cellspacing="1" class="border">
					 	<form name="myform" method="post" action="?Action=Send" onSubmit="return CheckForm();">
                        <tr class="Title">
                            <td height="24" colspan=2 align="center">重发激活码 </td>
                        </tr>
                            <tr class="tdbg">
                              <td width="40%" height="30" align="right"> 用户名：</td>
                              <td width="60%"><input name="UserName" type="text" id="UserName" size="20" /></td>
                            </tr>
                            <tr class="tdbg">
                              <td width="40%" height="30" align="right"> 您的邮箱：</td>
                              <td width="60%"><input name="Email" type="text" id="Email" size="20" /></td>
                            </tr>
                           
                            <tr class="tdbg">
                              <td colspan=2 height="42" align="center"><input class="Button" name="Submit2" type="submit" value="确定发送" />
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
                            </tr>
							</form>
                        </table>
		<%
		End If
       End Sub
	   
	   Sub Send()
	    Dim UserName:UserName=KS.R(KS.S("UserName"))
		Dim Email:Email=KS.S("Email")
		If UserName="" Then
		  Call KS.AlertHistory("请输入用户名!",-1)
		  Exit Sub
		End If
		If Email="" Then
		  Call KS.AlertHistory("请输入您的邮箱!",-1)
		  Exit Sub
		End If
		If KS.IsValidEmail(Email)=false Then
		  Call KS.AlertHistory("请正确的邮箱地址!",-1)
		  Exit Sub
		End If
		Dim RS:Set RS=KS.InitialObject("ADODB.RECORDSET")
		RS.Open "Select * From KS_User Where UserName='" & UserName & "'",conn,1,3
		If RS.Eof And RS.Bof Then
		  RS.Close:Set RS=Nothing
		  Call KS.AlertHistory("对不起,您输入的用户不存在!",-1)
		   Exit Sub
		 End If
		 If RS("Locked")<>3 Then
		   RS.Close:Set RS=Nothing
		   Call KS.AlertHistory("对不起,该项用户已经激活过了!",-1)
		   Exit Sub
		 End If
		 Dim RSG:Set RSG=Server.CreateObject("ADODB.RECORDSET")
		 RSG.Open "Select * From KS_UserGroup Where ID=" & RS("GroupID"),conn,1,1
		 If RSG.Eof Then RSG.Close : Set RSG=Nothing :Response.Write "<script>location.href='../../';</script>"
			
		 Dim UserRegSendMail:UserRegSendMail=RSG("ValidType")
		 Dim CheckNum:CheckNum = KS.MakeRandomChar(6)  '随机字符验证码
		 Dim CheckUrl:CheckUrl = Request.ServerVariables("HTTP_REFERER")
		 CheckUrl=KS.GetDomain &"User/?RegActive?UserName=" & UserName &"&CheckNum=" & CheckNum
		    Dim MailBodyStr
			MailBodyStr = Replace(RSG("ValidEmail"), "{$CheckNum}", CheckNum)
			MailBodyStr = Replace(MailBodyStr, "{$CheckUrl}", CheckUrl)
	        RSG.Close:Set RSG=Nothing
	       Dim ReturnInfo:ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), "新用户注册激活信", Email,KS.Setting(0), MailBodyStr,KS.Setting(11))
			  IF ReturnInfo="OK" Then
			     RS("CheckNum")=CheckNum
				 RS("Email")=Email
				 RS.Update
				 RS.Close:Set RS=Nothing
				 Response.Write "<script>alert('恭喜,激活码已发送到您的信箱" &Email &",请查收!');location.href='../';</script>"
			  Else
			     RS.Close:Set RS=Nothing
				 Response.Write "<script>alert('对不起,激活码发送失败!失败原因:" & ReturnInfo & "');history.back();</script>"
			  End if

	   End Sub
End Class
%> 
