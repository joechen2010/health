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
Set KSCls = New Admin_Index
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Index
        Private KS,KSUser
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		IF Cbool(KSUser.UserLoginChecked)=True Then
		 Response.Redirect("../")
		End If
		%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" lang="gb2312"> 
<head>
<title>会员登录-<%=KS.Setting(1)%></title>
<meta http-equiv=Content-Type content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css" />
<style>
	/*全局样式*/
* { margin:0; padding:0; }
body { font:12px/20px Verdana; color:#666; text-align:center; background:#fff; }
ul { list-style:none; }
img { border:none; }
img, input, select, button { vertical-align:middle; color:#666 }
input, select { font:12px Verdana; }
button { cursor:pointer; }
optgroup option { padding-left: 15px;}
a{color:#597D7D}
/*布局相关*/
.wrap_index { width:900px; margin:auto; text-align:left; }
.wrap { width:778px; margin:auto; text-align:left; }
#hd { height:50px; padding:10px; position:relative; border-bottom:1px solid #EFEFEF; }
#main { zoom:1; overflow:hidden; margin-bottom:20px; }
#ft { clear:both; text-align:center; line-height:22px; color:#C9C9C9; padding:12px; margin-bottom:20px; border-top:1px solid #EFEFEF; }


/*表单元素*/
	/*输入框*/
	.ipt_tx, .ipt_tx2 { border:1px solid #D2D2D2; background:#fff; line-height:16px; height:16px; padding:2px; margin:0; margin-left:2px}
		.ipt_tx2 { border-color:#9DB5CA; background:#F2F9FE; }
	/*按钮*/
	.btn { background:no-repeat; width:93px; height:28px; color:#333; font-size:12px; line-height:28px; border:none; }
		.bnormal { background-image:url(../images/btn_normal.gif); }


/*首页样式*/
.wrap_index .content { float:left; width:621px; overflow:hidden; }
	.wrap_index .side { float:right; width:210px; padding:18px; border:1px solid #597D7D; overflow:hidden; }
	.wrap_index .welcome { margin-top:80px;width:621px; overflow:hidden; margin-bottom:20px; }
	.wrap_index .content dl { margin-bottom:10px; padding-bottom:10px; border-bottom:1px solid #EFEFEF; line-height:18px; }
	.wrap_index .content dt { color:#F39800; margin-bottom:3px; font-weight:bold; font-size: 14px;}
	.wrap_index .content dd { padding:0 2em; }
	.wrap_index .side h2 { font-size:12px; border-bottom:1px solid #EFEFEF; margin-bottom:10px; line-height:24px; }
	.wrap_index .side .form_detail p { padding-left:13px; margin-bottom:12px; }
	.wrap_index .side .form_detail p label { width:53px; font-weight:normal; }
	
.head{height:80px;text-align:left}
#head{margin-left:-4px;}
.head #head_left{text-align:left;margin-top:-2px}
</style>
<script src="../../ks_inc/jquery.js"></script>
<script type="text/javascript">
var check={
   getCode:function(){
    $("#showVerify").html("<img align='absmiddle' src='../../plus/verifycode.asp' onClick='this.src=\"../../plus/verifycode.asp?n=\"+ Math.random();'>");
   },
   CheckForm:function(){
	 var username=$('#Username').val();
	 var pass=$('#Password').val();
	 var vycode=$('#Verifycode').val();
	 if (username==''){
		 alert('请输入用户名!');
		  $('#Username').focus();
		  return false;
	 }
	 if (pass==''){
		  alert('请输入登录密码!');
		  $('#Password').focus();
	      return false;
	 }
	 <%if KS.Setting(34)="1" then%>
	 if (vycode==''){
		  alert('请输入验证码!');
		  $('#Verifycode').focus();
		  return false;
	 }
	 <%end if%>
	}
}
</script>
</head>

<body>
<!-- head begin -->
<div class="head">
  <div id="head">
  <div id="head_left"><a href="http://www.kesion.com"><img alt="kesioncms" src="../images/logo.jpg" /></a></div>
  <div id="head_right">
	 <li><a href="../../">首页</a>┊</li><li><a href="../?user_Contributor.asp">匿名投稿</a>┊</li><li><a href="../login" target="main">登录</a>┊</li><li><a href="http://bbs.kesion.com" target="main">帮助</a></li>
  </div>
  </div>
</div>
<!-- head end -->

  <br />
  <br />
  <br />
<div class="wrap_index">
	<div id="main">
		<div id="div_act_content" class="content">
			<div class="welcome">
			   <dl>
				<dt>个人信息</dt>
				<dd>登录会员中心后，您可以完善你的个人资料，设置安全问题，实时了解个人账户情况。</dd>
			   </dl>
			   <dl>
				<dt>个人/企业空间</dt>
				<dd>加入我们您将免费拥有一个空间，个人会员将得到一个个人空间,您可以在上面写日志、上传照片、找朋友、加入圈子讨论等。企业会员将得到一个企业空间，您可以将公司的简介、公司产品、公司动态、公告招聘信息等发布到您的空间。</dd>
			   </dl>
			   <dl>
				<dt>求职招聘</dt>
				<dd>您可以在会员中心发布求职信息，招聘信息。个人简历，公司介绍等。</dd>
			   </dl>
			</div>
		</div>
		<div class="side">
			<h2>会员登录</h2>
			<form action="../CheckUserLogin.asp" id="myform" name="myform" method="post">
				<div class="form_detail">
					<p>
						<label>用户名：</label>
						<br><input type="text" name="Username" maxlength="60" id="Username" class="ipt_tx" style="width:149px;" tabindex="1" />
						
					</p>
					<p>
						<label>密码：</label>
						<br><input type="password" name="Password" maxlength="60" id="Password" class="ipt_tx" style="width:149px;" tabindex="2" autocomplete="off"/>
					</p>
					<%if KS.Setting(34)="1" then%>
					<p>
						<label>验证码：</label>
						<br><input type="text" maxlength="6" name="Verifycode" id="Verifycode" onFocus="this.value='';check.getCode()" class="ipt_tx" style="width:55px;" tabindex="3" autocomplete="off"/>
						<span id="showVerify">鼠标放到输入框将显示</span>
						
					</p>
					<%end if%>
					<p><input type="hidden" name="u1" id="u1"/>
						<input type="submit" tabindex="5"  onClick="return check.CheckForm();" class="btn bnormal" value="登  录">
					</p>
					<p><a tabindex="6" href="../?GetPass" target="_blank">忘记密码？</a>| <a tabindex="7" href="../?ActiveCode" target='_blank'>重发激活码</a></p>
				</div>
			</form>
			<h2>还没有会员帐号？</h2>
			<div class="form_detail">
				<p>
					<input type="button" tabindex="8" id="btn_regist" class="btn bnormal"  onclick="location.href='../reg/'" value="现在就注册" />
				</p>
			</div>
		</div>
	</div>
	<div id="ft">
		<p>漳州市科兴信息技术有限公司 &copy; 版权所有</p>
		<p>网址：http://www.kesion.com QQ：9537636 41904294</p>
	</div>
</div>
</body>
</html>
        <%
  End Sub
End Class
%> 
