<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="gb2312"
Dim KSCls
Set KSCls = New UserLogin
KSCls.Kesion()
Set KSCls = Nothing

Class UserLogin
        Private KS
		Private KSUser
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		%>
		<html>
<head>
<title>��Ա��¼</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.textbox
{
BACKGROUND-COLOR: #ffffff;
BORDER-BOTTOM: #666666 1px solid;
BORDER-LEFT: #666666 1px solid;
BORDER-RIGHT: #666666 1px solid;
BORDER-TOP: #666666 1px solid;
COLOR: #666666;
HEIGHT: 18px;
border-color: #666666 #666666 #666666 #666666; font-size: 9pt;FONT-FAMILY: verdana
}
TD
{
FONT-FAMILY:����;FONT-SIZE: 9pt;line-height: 130%;
}
a{text-decoration: none;} /* �������»���,��Ϊunderline */
a:link {color: #000000;} /* δ���ʵ����� */
a:visited {color: #333333;} /* �ѷ��ʵ����� */
a:hover{COLOR: #AE0927;} /* ����������� */
a:active {color: #0000ff;} /* ����������� */
.logintitle{font-size:14px;color:#336699;font-weight:bold}
#PopLogin td{font-size:14px;line-height:180%}
#PopLogin td a{color:#336699;text-decoration:underline}
#PopLogin td span{color:#5F5C67;font-size:12px}
#PopLogin td input{margin:2px}

-->
</style>
<script language="javascript">
//if(self==top){self.location.href="index.asp";}
function CheckForm(){
	var username=document.myform.Username.value;
	var pass=document.myform.Password.value;
	if (username=='')
	{
	  alert('�������û���');
	  document.myform.Username.focus();
	  return false;
    }
	if (pass=='')
	{
	  alert('�������¼����');
	  document.myform.Password.focus();
	  return false;
	 }
	 document.myform.submit();
}
function getCode(){
 document.getElementById('showVerify').innerHTML='<IMG style="cursor:pointer" src="../KS_Inc/verifycode.asp?n="<%=Timer%>" onClick="this.src=\'../ks_inc/verifycode.asp?n=\'+ Math.random();" align="absmiddle">';
}
</script>

</head>
<body leftmargin="0" topmargin="0" style="background-color:transparent">
		<%
		If KS.S("Action")="Top" Then
		   Call Login1()
		ElseIf KS.S("Action")="Poplogin" Then
		   Call PopLogin()
		Else
		   Call Login2()
		 End If
		End Sub
		
		Sub PopLogin()
		%>
		 <table id="PopLogin"  width="387" height="184" cellpadding="0" cellspacing="0" border="0">
		  <tr>
		   <td>
		     <table border="0" width="95%" align="center">
			     <form action="checkuserlogin.asp" method="post" name="myform" onSubmit="return(CheckForm())">
			  <tr>
			   <td style="border-right:solid 1px #cccccc">
			    û���˺ţ�<a href="reg/" target="_blank">����ע��</a><br/>
				��������, <a href="<%=KS.Setting(3)%>user/?getpass"  target="_blank">��Ҫ�һ�</a> <br />
			   </td>
			   <td>
			      <div class="logintitle">�û���¼</div>
				  <span>�û��˺ţ�</span><input type="text" name="Username" class="textbox"><br />
				  <span>��¼���룺</span><input type="password" name="Password" class="textbox"><br/>
				  <% If KS.Setting(34)="1" Then%>
				  <span>�����ַ���</span><input onfocus="getCode()" type="text" name="Verifycode" size="5" class="textbox"><span id='showVerify'></span><br/>
				  <%end if%>
				  <input type="hidden" name="Action" value="PopLogin">
				  <input type="submit" value=" �� ¼ " name="submit">
				   <input name="ExpiresDate" type="checkbox" id="ExpiresDate" value="checkbox">	<span>���õ�¼</span>
			   </td>
			  </tr>
				 </form>
			 </table>
		   </td>
		  </tr>
		 </table>
		<%
		End Sub
		
		Sub Login1()
			If KSUser.UserLoginChecked() = False Then
			%>
				<table cellspacing="0" cellpadding="0" width="99%" border="0">
				<form name="myform" action="<%=KS.GetDomain%>User/CheckUserLogin.asp?Action=Top" method="post" onSubmit="return(CheckForm())">
								<tr>
								  <td>�û�����<input class="textbox" size="10" name="Username" />�� �룺<input class="textbox" type="Password" size="10" name="Password"><%if KS.Setting(34)=1 Then%>��֤�룺<input name="Verifycode" type="text" class="textbox" id="Verifycode" size="6" /><%
				Response.Write "<IMG style=""cursor:pointer"" src=""" & KS.GetDomain & "KS_Inc/verifycode.asp?n=" & Timer & """ onClick=""this.src='" & KS.GetDomain & "ks_inc/verifycode.asp?n='+ Math.random();"" align=""absmiddle"">"
				end if%> 
								    <input name="loginsubmit" type="image" src="<%=KS.GetDomain%>images/login.gif" align="top" />
								    &nbsp;<a href="<%=KS.GetDomain%>User/reg/" target="_blank"><img src="<%=KS.GetDomain%>images/reg.gif"  border="0" align="absmiddle" twffan="done" /></a></td>
								</tr>
							</table>
			<%Else
			Dim MyMailTotal:MyMailTotal=Conn.Execute("Select Count(ID) From KS_Message Where Incept='" &KSUser.UserName &"' And Flag=0 and IsSend=1 and delR=0")(0)
			MyMailTotal=MyMailTotal+Conn.Execute("Select Count(ID) From KS_BlogMessage Where UserName='" &KSUser.UserName &"' And readtf=0")(0)
			MyMailTotal=MyMailTotal+Conn.Execute("Select Count(ID) From KS_BlogComment Where UserName='" &KSUser.UserName &"' And readtf=0")(0)
			MyMailTotal=MyMailTotal+Conn.Execute("Select Count(ID) From KS_Friend Where Friend='" &KSUser.UserName &"' And accepted=0")(0)
			
			 IF MyMailTotal>0 Then MyMailTotal="<font color=red>" & MyMailTotal & "</font>":Response.Write "<bgsound src=""" & KS.GetDomain & "User/images/mail.wav"" border=0>"
			%>
			<table cellspacing="0" cellpadding="0" width="99%" border="0">
				<tr>
			     <td height="22" align="center">����!<font color=red><%=KSUser.UserName%></font>,��ӭ������Ա����!&nbsp;��<a href="<%=KS.GetDomain%>User/index.asp?User_Message.asp?action=inbox" target="_parent">������ <%=MyMailTotal%></a>��&nbsp;��<a href="<%=KS.GetDomain%>User/index.asp" target="_parent">��Ա����</a>��&nbsp;��<a href="<%=KS.GetDomain%>User/UserLogout.asp">�˳���¼</a>��</td>
				</tr>
			</table>
<%End IF
		End Sub
		Sub Login2()
			If KSUser.UserLoginChecked() = False Then
			%>
			<table align="center" width="98%" border="0" cellspacing="0" cellpadding="0">
			 <form name="myform" action="CheckUserLogin.asp" method="post" onSubmit="return(CheckForm())">
			  <tr>
				<td height="25">�û�����
				<input name="Username" type="text" class="textbox" id="Username" size="15"></td>
			  </tr>
			  <tr>
				<td height="25">�ܡ��룺
				<input name="Password" type="password" class="textbox" id="Password" size="16"></td>
			  </tr>
			  <%if KS.Setting(34)=1 Then%>
			  <tr>
				<td height="25">��֤�룺
				<input name="Verifycode" onClick="getCode()" type="text" class="textbox" id="Verifycode" size="6">
				<span id='showVerify'></span>
				</td>
			  </tr>
			  <%end if%>
			  <tr>
				<td height="25"><div align="center"><img src="<%=KS.GetDomain%>images/losspass.gif" align="absmiddle"> <a href="<%=KS.GetDomain%>User/?GetPass" target="_parent">��������</a> <img src="<%=KS.GetDomain%>images/mas.gif" align="absmiddle"> <a href="<%=KS.GetDomain%>User/reg/" target="_parent">�»�Աע��</a>    </div></td>
			  </tr>
			  <tr>
				<td height="25"><div align="center">
				  <input type="submit" name="Submit" class="inputButton" value="��¼">

				  <input name="ExpiresDate" type="checkbox" id="ExpiresDate" value="checkbox">
			���õ�¼</div></td>
			  </tr>
			  </form>
            </table>
			<%Else
			Dim MyMailTotal:MyMailTotal=Conn.Execute("Select Count(ID) From KS_Message Where Incept='" &KSUser.UserName &"' And Flag=0 and IsSend=1 and delR=0")(0)
			MyMailTotal=MyMailTotal+Conn.Execute("Select Count(ID) From KS_BlogMessage Where UserName='" &KSUser.UserName &"' And readtf=0")(0)
			MyMailTotal=MyMailTotal+Conn.Execute("Select Count(ID) From KS_BlogComment Where UserName='" &KSUser.UserName &"' And readtf=0")(0)
			MyMailTotal=MyMailTotal+Conn.Execute("Select Count(ID) From KS_Friend Where Friend='" &KSUser.UserName &"' And accepted=0")(0)
			 IF MyMailTotal>0 Then MyMailTotal="<font color=red>" & MyMailTotal & "</font>"
			 dim  ChargeTypeStr
			 if KSUser.ChargeType=1 Then
			   ChargeTypeStr="�۵�"
			 elseif KSUser.ChargeType=2 Then
			   ChargeTypeStr="��Ч��"
			 else
			   ChargeTypeStr="������"
			 End If
			%>
			<table align="center" style="margin-top:5px" width="80%" border="0" cellspacing="0" cellpadding="0">
			<tr><td align="center"><font color=red><%=KSUser.UserName%></font>,
           <%
			If (Hour(Now) < 6) Then
            Response.Write "<font color=##0066FF>�賿��!</font>"
			ElseIf (Hour(Now) < 9) Then
				Response.Write "<font color=##000099>���Ϻ�!</font>"
			ElseIf (Hour(Now) < 12) Then
				Response.Write "<font color=##FF6699>�����!</font>"
			ElseIf (Hour(Now) < 14) Then
				Response.Write "<font color=##FF6600>�����!</font>"
			ElseIf (Hour(Now) < 17) Then
				Response.Write "<font color=##FF00FF>�����!</font>"
			ElseIf (Hour(Now) < 18) Then
				Response.Write "<font color=##0033FF>�����!</font>"
			Else
				Response.Write "<font color=##ff0000>���Ϻ�!</font>"
			End If
			%>&nbsp;&nbsp;&nbsp;</td></tr>
			<tr><td>�Ʒѷ�ʽ�� <strong><%= ChargeTypeStr%></strong> </td></tr>
			<tr><td>������֣� <strong><%=KSUser.Score%></strong> ��</td></tr>
			<%if KSUser.ChargeType=1 or KSUser.ChargeType=2 then%>
			<% if KSUser.ChargeType=1 then%>
			<tr><td>���õ�ȯ�� <strong><%=KSUser.Point%></strong> ��</td></tr>
			<%else%>
			<tr><td>ʣ�������� <strong><%=KSUser.GetEdays%></strong></td></tr>
			<%end if%>
			<%end if%>
			<tr><td>���Ķ��ţ� <strong><%=MyMailTotal%></strong> ��</td></tr>
			<tr><td>��¼������ <strong><%=KSUser.LoginTimes%></strong> ��</td></tr>
            <tr><td nowrap="nowrap">��<a href="<%=KS.GetDomain%>User/index.asp" target="_parent">��Ա����</a>����<a href="<%=KS.GetDomain%>User/UserLogout.asp">�˳���¼</a>��</td></tr>
			</table>
<%End IF
  End Sub
End Class
%>

 
