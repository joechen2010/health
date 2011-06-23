<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../Plus/md5.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%
'强制浏览器重新访问服务器下载页面，而不是从缓存读取页面
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim KS:Set KS=New PublicCls

Select Case  KS.G("Action")
 Case "LoginCheck"
  Call CheckLogin()
 Case "LoginOut"
  Call LoginOut()
 Case Else
  Call CheckSetting()
  Call Main()
End Select

Sub CheckSetting()
     dim strDir,strAdminDir,InstallDir
	 strDir=Trim(request.ServerVariables("SCRIPT_NAME"))
	 strAdminDir=split(strDir,"/")(Ubound(split(strDir,"/"))-1) & "/"
	 InstallDir=left(strDir,instr(lcase(strDir),"/"&Lcase(strAdminDir)))
			
	If Instr(UCASE(InstallDir),"/W3SVC")<>0 Then
	   InstallDir=Left(InstallDir,Instr(InstallDir,"/W3SVC"))
	End If
 If KS.Setting(2)<>KS.GetAutoDoMain or KS.Setting(3)<>InstallDir Then
	
  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
  RS.Open "Select Setting From KS_Config",conn,1,3
  Dim SetArr,SetStr,I
  SetArr=Split(RS(0),"^%^")
  For I=0 To Ubound(SetArr)
   If I=0 Then 
    SetStr=SetArr(0)
   ElseIf I=2 Then
    SetStr=SetStr & "^%^" & KS.GetAutoDomain
   ElseIf I=3 Then
    SetStr=SetStr & "^%^" & InstallDir
   Else
    SetStr=SetStr & "^%^" & SetArr(I)
   End If
  Next
  RS(0)=SetStr
  RS.Update
  RS.Close:Set RS=Nothing
  Call KS.DelCahe(KS.SiteSn & "_Config")
  Call KS.DelCahe(KS.SiteSn & "_Date")
 End If
End Sub

Sub Main()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<HEAD>
<TITLE><%=KS.Setting(0) & "---网站后台管理"%></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<meta http-equiv="X-UA-Compatible" content="IE=7" />
<script language="JavaScript" type="text/JavaScript" src="Include/SoftKeyBoard.js"></script>
<SCRIPT type="text/JavaScript">
	<!--
    nereidFadeObjects = new Object();
	nereidFadeTimers = new Object();
	function nereidFade(object, destOp, rate, delta){
				if (!document.all)
				return
					if (object != "[object]"){ 
						setTimeout("nereidFade("+object+","+destOp+","+rate+","+delta+")",0);
						return;
					}
					clearTimeout(nereidFadeTimers[object.sourceIndex]);
					diff = destOp-object.filters.alpha.opacity;
					direction = 1;
					if (object.filters.alpha.opacity > destOp){
						direction = -1;
					}
					delta=Math.min(direction*diff,delta);
					object.filters.alpha.opacity+=direction*delta;
					if (object.filters.alpha.opacity != destOp){
						nereidFadeObjects[object.sourceIndex]=object;
						nereidFadeTimers[object.sourceIndex]=setTimeout("nereidFade(nereidFadeObjects["+object.sourceIndex+"],"+destOp+","+rate+","+delta+")",rate);
					}
	}
	//-->
</SCRIPT>
<STYLE>
body{ margin:0px auto; padding:0px auto; font-size:12px; color:#555; font-family:Verdana, Arial, Helvetica, sans-serif;text-align:center;}
html {
overflow:hidden;
} 
form,div,ul,li {margin:0px auto;padding: 0; border: 0; }
img,a img{border:0; margin:0; padding:0;}
a{font-size:12px;color:#000000}
a:link,a:visited{color:#00436D;}
td {
	font-size: 12px; color: #fff; LINE-HEIGHT: 20px; FONT-FAMILY: "", Arial, Tahoma
}
.textbox{height:16px;}
.head{ background:url(images/headbg.gif) repeat-x;height:84px;}
.logo{width:400px;height:68px;float:left;margin:0px;padding:0px;}
.right{width:580px;float:right;color:#ccc;font-family:"宋体";padding-top:30px;}
.right a:link,.right a:visited{color:#666; text-decoration: none}
.right a:hover{ color:#ff0000; text-decoration:underline}
.main{margin-top:80px;width:100%;text-align:center}
.loginbg{width:541px;height:300px;background:url(images/loginbg.gif) no-repeat;text-align:center;border:1px solid #fff;}
.login{width:489px;height:213px;background:url(images/login.jpg) no-repeat;margin-top:20px;padding-top:102px !important;padding-top:98px; }
#login{text-align:left;padding-left:6px;}
#copyright {margin:0px auto;left:250px;bottom:100px;text-align:center;width:550px;PADDING: 1px 1px 1px 1px; FONT: 12px 
verdana,arial,helvetica,sans-serif; COLOR: #666; TEXT-DECORATION: none}
#copyright a:link,#copyright a:visited{color:#ff6600; font-size:12px; text-decoration:underline}
</STYLE>
</head>
<body>

<div class="head">
  <div class="logo"><img src="images/kslogo.jpg" /></div>
  <div class="right"><a href="http://www.kesion.com" target="_blank">官方首页</a> | <a href="http://help.kesion.com" target="_blank"> 帮助中心</a> |<a href="http://www.kesion.com" target="_blank"> 会员中心</a>|<a href="http://bbs.kesion.com" target="_blank"> 交流论坛</a></div>
</div>
<div class="main">
  <div class="loginbg">
    <div class="login">
	    <FORM ACTION="Login.asp?Action=LoginCheck" method="post" name="LoginForm" onSubmit="return(CheckForm(this))">
      <table width="93%" border="0" align="center" cellpadding="0" cellspacing="0" id="login">
         <tr>
          <td width="21" align="right">&nbsp;</td>
          <td width="177">
              用户名：<input name="UserName" tabindex="1" class="textbox" maxlength="30" id="UserName" type="text" size="12">              </td>
          <td width="24">&nbsp;</td>
          <td width="161">密&nbsp;&nbsp;&nbsp;码：<%IF KS.Setting(98)=1 Then%><input name="PWD" type="password" class="textbox"  style="FONT-FAMILY: verdana" tabindex="2" onFocus="this.select();" onChange="Calc.password.value=this.value;" onClick="password1=this;showkeyboard();this.readOnly=1;Calc.password.value=''" onKeyDown="Calc.password.value=this.value;" size="11" maxlength="50" readOnly>
                <%Else%><INPUT name="PWD"  type="password" class="textbox" style="FONT-FAMILY: verdana" tabindex="2" size="11" maxlength="50">
            <%End IF%>	</td>
          <td width="63" rowspan="3"><input type="image" src="images/grdl.gif" onMouseOver="nereidFade(this,100,10,5)"  style="FILTER:alpha(opacity=50)" onMouseOut="nereidFade(this,50,10,5)"></td>
        </tr>
        <tr>
          <td height="18" colspan="4"></td>
        </tr>
        <tr>
          <td align="right">&nbsp;</td>
          <td>验证码：<input type=text name="Verifycode"  maxLength=6 size="8" tabindex="3" class="textbox"> 
		<IMG style="cursor:pointer;" src="../plus/verifycode.asp?n=<%=Timer%>" onClick="this.src='../plus/verifycode.asp?n='+ Math.random();" align="absmiddle"></td>
          <td>&nbsp;</td>
          <td>
		  <%if EnableSiteManageCode = True Then%>
		  认证码：<input name="AdminLoginCode" type="password"  class="textbox" id="AdminLoginCode"style="FONT-FAMILY: verdana" tabindex="4" title="认证码初始值“8888”,为了系统的安全请及时修改admin目录下的chkcode.asp文件参数!" size="11" maxlength="20" <%if EnableSiteManageCode = True And SiteManageCode = "8888" then%><%end if%>>
			</td>
		<%else%>
		 没有启用认证码
		<%end if%>
		  
		  </td>
        </tr>
		<input type="hidden" value="1" name="skinid">
      </table>
	  		</form>
	  <%
	  if EnableSiteManageCode=true And SiteManageCode="8888" Then
	   Response.Write"<br /><span style='color:#ffffff'>系统初始认证码为<span style='color:#ff0000'>8888</span>,为了您的系统安全,建议打开conn.asp进行修改!</span>"
	  End If
	  %>
    </div>
  </div>
</div>
<br />
<div class="line"></div>
<div class="botinfo" id="copyright"> 
漳州科兴信息技术有限公司 Copyright &copy;2006-2010 <a href="http://www.kesion.com" target="_blank"> www.kesion.com</a>,All Rights Reserved. </div>


<script type="text/javascript">
<!--
function document.onreadystatechange()
{  var app=navigator.appName;
  var verstr=navigator.appVersion;
  if(app.indexOf('Netscape') != -1) {
    alert('友情提示：\n    您使用的是Netscape浏览器，可能会导致无法使用后台的部分功能。建议您使用 IE6.0 或以上版本。');
  } else if(app.indexOf('Microsoft') != -1) {
    if (verstr.indexOf('MSIE 3.0')!=-1 || verstr.indexOf('MSIE 4.0') != -1 || verstr.indexOf('MSIE 5.0') != -1 || verstr.indexOf('MSIE 5.1') != -1)
      alert('友情提示：\n    您的浏览器版本太低，可能会导致无法使用后台的部分功能。建议您使用 IE6.0 或以上版本。');
  }
  document.LoginForm.UserName.focus();
}
function CheckForm(ObjForm) {
  if(ObjForm.UserName.value == '') {
    alert('请输入管理账号！');
    ObjForm.UserName.focus();
    return false;
  }
  if(ObjForm.PWD.value == '') {
    alert('请输入授权密码！');
    ObjForm.PWD.focus();
    return false;
  }
  if (ObjForm.PWD.value.length<6)
  {
    alert('授权密码不能少于六位！');
    ObjForm.PWD.focus();
    return false;
  }
  if (ObjForm.Verifycode.value == '') {
    alert ('请输入验证码！');
    ObjForm.Verifycode.focus();
    return false;
  }
  <%if EnableSiteManageCode = True Then%>
  if (ObjForm.AdminLoginCode.value == '') {
    alert ('请输入后台管理认证码！');
    ObjForm.AdminLoginCode.focus();
    return false;
  }
  <%End If%>
}
//-->
</script>
</html>
<%End Sub
Sub CheckLogin()
  Dim PWD,UserName,LoginRS,SqlStr,RndPassword
  Dim ScriptName,AdminLoginCode
  AdminLoginCode=KS.G("AdminLoginCode")
  IF Trim(Request.Form("Verifycode"))<>Trim(Session("Verifycode")) then 
   Call KS.Alert("登录失败:\n\n验证码有误，请重新输入！","Login.asp")
   exit Sub
  end if
  If EnableSiteManageCode = True And AdminLoginCode <> SiteManageCode Then
   Call KS.Alert("登录失败:\n\n您输入的后台管理认证码不对，请重新输入！","Login.asp")
   exit Sub
  End If
  Pwd =MD5(KS.R(Request.form("pwd")),16)

  UserName = KS.R(trim(Request.form("username")))
  RndPassword=KS.R(KS.MakeRandomChar(20))
  ScriptName=KS.R(Trim(Request.ServerVariables("HTTP_REFERER")))
  
  Set LoginRS = Server.CreateObject("ADODB.RecordSet")
  SqlStr = "select * from KS_Admin where UserName='" & UserName & "'"
  LoginRS.Open SqlStr,Conn,1,3
  If LoginRS.EOF AND LoginRS.BOF Then
	  Call KS.InsertLog(UserName,0,ScriptName,"输入了错误的帐号!")
      Call KS.AlertHistory("登录失败:\n\n您输入了错误的帐号，请再次输入！",-1)
  Else
  
     IF LoginRS("PassWord")=pwd THEN
       IF Cint(LoginRS("Locked"))=1 Then
          Call KS.Alert("登录失败:\n\n您的账号已被管理员锁定，请与您的系统管理员联系！","Login.asp")	
	      Response.End
	   Else
		  	 '登录成功，进行前台验证，并更新数据
			  Dim UserRS:Set UserRS=Server.CreateObject("Adodb.Recordset")
			  UserRS.Open "Select Score,LastLoginIP,LastLoginTime,LoginTimes,UserName,Password,RndPassWord,IsOnline From KS_User Where UserName='" & LoginRS("PrUserName") & "' and GroupID=1",Conn,1,3
			  IF Not UserRS.Eof Then
			  
						If datediff("n",UserRS("LastLoginTime"),now)>=KS.Setting(36) then '判断时间
						UserRS("Score")=UserRS("Score")+KS.Setting(37)
						end if
					 UserRS("LastLoginIP") = KS.GetIP
					 UserRS("LastLoginTime") = Now()
					 UserRS("LoginTimes") = UserRS("LoginTimes") + 1
					 UserRS("RndPassWord") = RndPassWord
					 UserRS("IsOnline")=1
					 UserRS.Update		
	
					'置前台会员登录状态
					 Response.Cookies(KS.SiteSn).path = "/"
					 Response.Cookies(KS.SiteSn)("UserName") = KS.R(UserRS("UserName"))
			         Response.Cookies(KS.SiteSn)("Password") = UserRS("Password")
					 Response.Cookies(KS.SiteSn)("RndPassword") = KS.R(UserRS("RndPassword"))
					 Response.Cookies(KS.SiteSn)("AdminLoginCode") = AdminLoginCode
					 Response.Cookies(KS.SiteSn)("AdminName") =  UserName
					 Response.Cookies(KS.SiteSn)("AdminPass") = pwd
					 Response.Cookies(KS.SiteSn)("SuperTF")   = LoginRS("SuperTF")
					 Response.Cookies(KS.SiteSn)("PowerList") = LoginRS("PowerList")
					 Response.Cookies(KS.SiteSn)("ModelPower") = LoginRS("ModelPower")
					 Response.Cookies(KS.SiteSn)("SkinID")  = KS.S("SkinID")
             Else 
				   Call KS.InsertLog(UserName,0,ScriptName,"找不到前台账号!")
				   Call KS.Alert("登录失败:\n\n找不到前台账号！","Login.asp")	
				   Response.End
			 End If
			   UserRS.Close:Set UserRS=Nothing
			   
	  LoginRS("LastLoginTime")=Now
	  LoginRS("LastLoginIP")=KS.GetIP
	  LoginRS("LoginTimes")=LoginRS("LoginTimes")+1
	  LoginRS.UpDate
	  Call KS.InsertLog(UserName,1,ScriptName,"成功登录后台系统!")
	  Response.Redirect("Index.asp")
	End IF
  ELse
    Response.Cookies(KS.SiteSn).path = "/"
    Response.Cookies(KS.SiteSn)("AdminName")=""
	Response.Cookies(KS.SiteSn)("AdminPass")=""
	Response.Cookies(KS.SiteSn)("SuperTF")=""
	Response.Cookies(KS.SiteSn)("AdminLoginCode")=""
	Response.Cookies(KS.SiteSn)("PowerList")=""
	Response.Cookies(KS.SiteSn)("ModelPower")=""
	Call KS.InsertLog(UserName,0,ScriptName,"输入了错误的口令:" & Request.form("pwd"))
    Call KS.Alert("登录失败:\n\n您输入了错误的口令，请再次输入！","Login.asp")	
  END IF
 End If
END Sub
Sub LoginOut()
		   Conn.Execute("Update KS_Admin Set LastLogoutTime=" & SqlNowString & " where UserName='" & KS.R(KS.C("AdminName")) &"'")
		   Dim AdminDir:AdminDir=KS.Setting(89)
		    Response.Cookies(KS.SiteSn).path = "/"
			Response.Cookies(KS.SiteSn)("PowerList")=""
			Response.Cookies(KS.SiteSn)("AdminName")=""
			Response.Cookies(KS.SiteSn)("AdminPass")=""
			Response.Cookies(KS.SiteSn)("SuperTF")=""
			Response.Cookies(KS.SiteSn)("AdminLoginCode")=""
			Response.Cookies(KS.SiteSn)("ModelPower")=""
			session.Abandon()
			Response.Write ("<script> top.location.href='" & KS.Setting(2) & KS.Setting(3) &"';</script>")
End Sub
Set KS=Nothing
%>