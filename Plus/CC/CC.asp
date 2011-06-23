<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../Plus/Session.asp"-->
<!--#include file="config.asp"-->

<%
'****************************************************
' Software name:Kesion CMS 4.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************

Dim Chk:Set Chk=New LoginCheckCls1
Chk.Run()
Set Chk=Nothing
Dim KS:Set KS=New PublicCls
Dim Action:Action = LCase(KS.S("action"))
Select Case Trim(Action)
	Case "save"
		Call SaveConformify
	Case Else
		Call showmain
End Select
Set KS=Nothing

Sub showmain()
Response.Write "<html><head><title>CC视频联盟参数设置</title>" & vbCrLf
Response.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
Response.Write "<link href='Admin_Style.css' rel='stylesheet' type='text/css'></head>" & vbCrLf
Response.Write "</head>"
Response.Write "<body bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"" scroll=no>"
Response.Write "<ul id='menu_top' style='text-align:center;padding-top:10px;font-weight:bold'> CC视频联盟参数设置</ul>"

%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1">
<form name="myform" method="post" action="?action=save">
<tr class="tdbg">
	<td height="30" width="25%" class="clefttitle" align="right"><strong>是否开启CC视频联盟：</strong></td>
	<td>
	<input type="radio" name="opentf" value="false"<%
	If Not opentf Then Response.Write " checked"
	%>> 关闭&nbsp;&nbsp;
	<input type="radio" name="opentf" value="true"<%
	If opentf Then Response.Write " checked"
	%>> 开启
	</td>
</tr>

<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>您在CC视频联盟的ID：</strong><br>(<a href="http://union.bokecc.com/signup.bo" target="_blank"><font color=red>还没有账号，点此注册</font></a>)</td>
	<td><input type="text" name="userid" value="<%=userid%>"></td>
</tr>
<tr class="tdbg">
	<td height="30" class="clefttitle" align="right"><strong>选择按钮样式：</strong></td>
	<td>
	<select name="buttonstyle" onchange="s(this.value);">
	<option value="1"<%if buttonstyle="1" then response.write " selected"%>>1(72x24)</option>
							<option value="2"<%if buttonstyle="2" then response.write " selected"%>>2(72x24)</option>
							<option value="3"<%if buttonstyle="3" then response.write " selected"%>>3(72x24)</option>
							<option value="4"<%if buttonstyle="4" then response.write " selected"%>>4(72x24)</option>
							<option value="5"<%if buttonstyle="5" then response.write " selected"%>>5(72x24)</option>
							<option value="6"<%if buttonstyle="6" then response.write " selected"%>>6(72x24)</option>
							<option value="7"<%if buttonstyle="7" then response.write " selected"%>>7(72x24)</option>
							<option value="8"<%if buttonstyle="8" then response.write " selected"%>>8(71x24)</option>
							<option value="9"<%if buttonstyle="9" then response.write " selected"%>>9(72x24)</option>
							<option value="10"<%if buttonstyle="10" then response.write " selected"%>>10(72x24)</option>
							<option value="11"<%if buttonstyle="11" then response.write " selected"%>>11(72x24)</option>
							<option value="12"<%if buttonstyle="12" then response.write " selected"%>>12(72x24)</option>
							<option value="13"<%if buttonstyle="13" then response.write " selected"%>>13(72x24)</option>
							<option value="14"<%if buttonstyle="14" then response.write " selected"%>>14(72x24)</option>
							<option value="15"<%if buttonstyle="15" then response.write " selected"%>>15(72x24)</option>
							<option value="16"<%if buttonstyle="16" then response.write " selected"%>>16(86x22)</option>
							
	 </select>
	 <img name="selectimg" id="b" width="72" height="24" border=0 src=""/>
	 	<script>
	 s(<%=buttonstyle%>);
	 function s(v)
	 {
	   if (v<=9)
	    document.all.b.src='images/0'+v+'.png';
	   else
	   document.all.b.src='images/'+v+'.gif';
	 }
	</script>

	</td>
</tr>
<tr>
  <td colspan=2 class="clefttitle" align="center"><input type="submit" value="确定设置" /></td>
</tr>

</form>

</table>
<script>
 function CheckForm()
 {
  document.all.myform.submit();
 }
</script>
<%
End Sub

Sub SaveConformify()
dim fs,ts1
Set fs=Server.CreateObject("Scripting.FileSystemObject")
Set TS1 = fs.CreateTextFile(Server.MapPath("config.asp"), True) 
TS1.writeline "<"&chr(37)
'站点设置
TS1.writeline "const opentf="&chr(34)&""&KS.R(KS.S("opentf"))&""&chr(34)&""
TS1.writeline "const userid="&chr(34)&""&KS.R(KS.S("userid"))&""&chr(34)&""
TS1.writeline "const buttonstyle="&chr(34)&""&KS.R(KS.S("buttonstyle"))&""&chr(34)&""
TS1.writeline chr(37)&">"
Set TS1 = Nothing
Set fs=nothing
	Response.Write ("<script>alert('恭喜您！保存设置成功。');location.href='cc.asp';</script>")
End Sub

%>