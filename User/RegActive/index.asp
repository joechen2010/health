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
Set KSCls = New UserReg
KSCls.Kesion()
Set KSCls = Nothing

Class UserReg
        Private KS, KSUser
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser=New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		  Call KSUser.Head()
		  Call KSUser.InnerLocation("会员注册激活")
			Dim UserName:UserName=KS.S("UserName")
			Dim CheckNum:CheckNum=KS.S("CheckNum")
		  If KS.S("Flag")="Check" Then
		    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select UserName,RndPassWord,Email,CheckNum,locked From KS_User Where UserName='" & UserName & "'",Conn,1,3
			If RS.Eof And RS.Bof Then
			rs.close:set rs=nothing
			 Response.Write "<script>alert('对不起，您输入的用户名不存在！');history.back();</script>":response.end
			else
			  if rs("checknum")<>checknum then
			  rs.close:set rs=nothing
			   Response.Write "<script>alert('激活码有误，请重新输入！');history.back();</script>":response.end
			  else
			   rs("locked")=0
			   rs.update
			   
			    Dim MailBodyStr,ReturnInfo
			    MailBodyStr = Replace(KS.Setting(147), "{$UserName}", rs("UserName"))
				MailBodyStr = Replace(MailBodyStr, "{$PassWord}", rs("RndPassWord"))
				MailBodyStr = Replace(MailBodyStr, "{$SiteName}", KS.Setting(0))
				ReturnInfo=KS.SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), KS.Setting(0) & "-会员注册成功", RS("Email"),rs("UserName"), MailBodyStr,KS.Setting(11))

				IF ReturnInfo="OK" Then
				  ReturnInfo="<li>注册成功!您的用户名:<font color=red>" & RS("UserName") & "</font>,已将用户名和密码发到您的信箱!</li>"
				End If
				rs.close:set rs=nothing
			   Response.Write "<script>alert('恭喜您,账号激活成功,您现在可以正常登录了！');location.href='../User_main.asp';</script>":response.end
			  end if
			end if
		  End If
		   %>
		    <form name="myform" method="post" action="?Flag=Check" onSubmit="return CheckForm();">
                 <table width="96%"  border="0" align="center" cellpadding="0" cellspacing="1" class="border">
					  <tr class="Title">
							<td height="24" align="center" colspan="2">用 户 激 活</td>
					  </tr>
						  <TR class="tdbg">
						    <TD height=25 align="right">您的用户名：</TD>
						    <TD><input name="UserName" type="text" id="UserName" size="20" value="<%=UserName%>"></TD>
			              </TR>
						  <TR class="tdbg">
							<TD width="40%" height=25 align="right"> 您的激活码：</TD>
							<TD width="60%"><input name="CheckNum" type="text" id="CheckNum" size="20" value="<%=CheckNum%>"></TD>
						  </TR>
						  <TR class="tdbg">
							<TD  colspan="2" height=42 align="center"> 
							<input name="Submit" type="submit" value="确定激活" style="padding:3px">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </TD>
						  </TR>
						</TBODY>
					  </TABLE>
</form>
		   <%
		End Sub

End Class
%>

 
