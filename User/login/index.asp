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
<title>��Ա��¼-<%=KS.Setting(1)%></title>
<meta http-equiv=Content-Type content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css" />
<style>
	/*ȫ����ʽ*/
* { margin:0; padding:0; }
body { font:12px/20px Verdana; color:#666; text-align:center; background:#fff; }
ul { list-style:none; }
img { border:none; }
img, input, select, button { vertical-align:middle; color:#666 }
input, select { font:12px Verdana; }
button { cursor:pointer; }
optgroup option { padding-left: 15px;}
a{color:#597D7D}
/*�������*/
.wrap_index { width:900px; margin:auto; text-align:left; }
.wrap { width:778px; margin:auto; text-align:left; }
#hd { height:50px; padding:10px; position:relative; border-bottom:1px solid #EFEFEF; }
#main { zoom:1; overflow:hidden; margin-bottom:20px; }
#ft { clear:both; text-align:center; line-height:22px; color:#C9C9C9; padding:12px; margin-bottom:20px; border-top:1px solid #EFEFEF; }


/*��Ԫ��*/
	/*�����*/
	.ipt_tx, .ipt_tx2 { border:1px solid #D2D2D2; background:#fff; line-height:16px; height:16px; padding:2px; margin:0; margin-left:2px}
		.ipt_tx2 { border-color:#9DB5CA; background:#F2F9FE; }
	/*��ť*/
	.btn { background:no-repeat; width:93px; height:28px; color:#333; font-size:12px; line-height:28px; border:none; }
		.bnormal { background-image:url(../images/btn_normal.gif); }


/*��ҳ��ʽ*/
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
		 alert('�������û���!');
		  $('#Username').focus();
		  return false;
	 }
	 if (pass==''){
		  alert('�������¼����!');
		  $('#Password').focus();
	      return false;
	 }
	 <%if KS.Setting(34)="1" then%>
	 if (vycode==''){
		  alert('��������֤��!');
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
	 <li><a href="../../">��ҳ</a>��</li><li><a href="../?user_Contributor.asp">����Ͷ��</a>��</li><li><a href="../login" target="main">��¼</a>��</li><li><a href="http://bbs.kesion.com" target="main">����</a></li>
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
				<dt>������Ϣ</dt>
				<dd>��¼��Ա���ĺ�������������ĸ������ϣ����ð�ȫ���⣬ʵʱ�˽�����˻������</dd>
			   </dl>
			   <dl>
				<dt>����/��ҵ�ռ�</dt>
				<dd>���������������ӵ��һ���ռ䣬���˻�Ա���õ�һ�����˿ռ�,������������д��־���ϴ���Ƭ�������ѡ�����Ȧ�����۵ȡ���ҵ��Ա���õ�һ����ҵ�ռ䣬�����Խ���˾�ļ�顢��˾��Ʒ����˾��̬��������Ƹ��Ϣ�ȷ��������Ŀռ䡣</dd>
			   </dl>
			   <dl>
				<dt>��ְ��Ƹ</dt>
				<dd>�������ڻ�Ա���ķ�����ְ��Ϣ����Ƹ��Ϣ�����˼�������˾���ܵȡ�</dd>
			   </dl>
			</div>
		</div>
		<div class="side">
			<h2>��Ա��¼</h2>
			<form action="../CheckUserLogin.asp" id="myform" name="myform" method="post">
				<div class="form_detail">
					<p>
						<label>�û�����</label>
						<br><input type="text" name="Username" maxlength="60" id="Username" class="ipt_tx" style="width:149px;" tabindex="1" />
						
					</p>
					<p>
						<label>���룺</label>
						<br><input type="password" name="Password" maxlength="60" id="Password" class="ipt_tx" style="width:149px;" tabindex="2" autocomplete="off"/>
					</p>
					<%if KS.Setting(34)="1" then%>
					<p>
						<label>��֤�룺</label>
						<br><input type="text" maxlength="6" name="Verifycode" id="Verifycode" onFocus="this.value='';check.getCode()" class="ipt_tx" style="width:55px;" tabindex="3" autocomplete="off"/>
						<span id="showVerify">���ŵ��������ʾ</span>
						
					</p>
					<%end if%>
					<p><input type="hidden" name="u1" id="u1"/>
						<input type="submit" tabindex="5"  onClick="return check.CheckForm();" class="btn bnormal" value="��  ¼">
					</p>
					<p><a tabindex="6" href="../?GetPass" target="_blank">�������룿</a>| <a tabindex="7" href="../?ActiveCode" target='_blank'>�ط�������</a></p>
				</div>
			</form>
			<h2>��û�л�Ա�ʺţ�</h2>
			<div class="form_detail">
				<p>
					<input type="button" tabindex="8" id="btn_regist" class="btn bnormal"  onclick="location.href='../reg/'" value="���ھ�ע��" />
				</p>
			</div>
		</div>
	</div>
	<div id="ft">
		<p>�����п�����Ϣ�������޹�˾ &copy; ��Ȩ����</p>
		<p>��ַ��http://www.kesion.com QQ��9537636 41904294</p>
	</div>
</div>
</body>
</html>
        <%
  End Sub
End Class
%> 
