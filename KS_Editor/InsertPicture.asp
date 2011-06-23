<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Plus/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 4.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
Dim KSCls
Set KSCls = New InsertPicture
KSCls.Kesion()
Set KSCls = Nothing

Class InsertPicture
        Private KS
		Private AdminDir
		Private ChannelID
        Private CurrPath,InstallDir
		Private FromUrl
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		End Sub
        Public Sub Kesion()
			ChannelID = KS.S("ChannelID")
			FromUrl=KS.ChkCLng(KS.S("FromUrl"))
			
			IF FromUrl=1 Then  '后台调用，检查是否登录
				Dim KSLoginCls
				Set KSLoginCls = New LoginCheckCls1
				KSLoginCls.Run()
				Set KSLoginCls= Nothing
			End IF
			
			CurrPath=KS.GetUpFilesDir()
			If Right(CurrPath,1)="/" Then CurrPath=Left(CurrPath, Len(CurrPath) - 1)
			AdminDir=KS.GetDomain & KS.Setting(89)
			%>
			<HTML>
			<HEAD>
			<%if ks.s("action")="edit" then%>
			<TITLE>修改图片属性</TITLE>
			<%else%>
			<TITLE>插入图片文件</TITLE>
			<%end if%>
			<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
			<link href="Editor.css" rel="stylesheet" type="text/css">
			<script language="JavaScript" src="../KS_Inc/Common.js"></script>
			<script language="JavaScript">
			function OK()
			{
			  var str1="";
			  var strurl=document.PicForm.url.value;
			  if (strurl==""||strurl=="http://")
			  {
				alert("请先输入图片地址，或者上传图片！");
				document.PicForm.url.focus();
				return false;
			  }
			  else
			  {
				str1="<img src='"+document.PicForm.url.value+"' alt='"+document.PicForm.alttext.value+"' ";
				if(document.PicForm.width.value!=''&&document.PicForm.width.value!='0') str1=str1+"width='"+document.PicForm.width.value+"' ";
				if(document.PicForm.height.value!=''&&document.PicForm.height.value!='0') str1=str1+"height='"+document.PicForm.height.value+"' ";
				str1=str1+"border='"+document.PicForm.PicBorder.value+"' align='"+document.PicForm.aligntype.value+"' ";
				if(document.PicForm.vspace.value!=''&&document.PicForm.vspace.value!='0') str1=str1+"vspace='"+document.PicForm.vspace.value+"' ";
				if(document.PicForm.hspace.value!=''&&document.PicForm.hspace.value!='0') str1=str1+"hspace='"+document.PicForm.hspace.value+"' ";
				if(document.PicForm.styletype.value!='')	str1=str1+"style='"+document.PicForm.styletype.value+"'";
				str1=str1+">";
				window.returnValue=str1+"$$$"+document.PicForm.UpFileName.value;
				window.close();
			  }
			}
			function ShowPicture()
			{
			  var str1="";
			  var strurl=document.PicForm.url.value;
			  if (strurl==""||strurl=="http://")
			  {
			  strurl='images/FileType/nopic.gif';
			  str1="<br><br><br><br><br><img src='"+ strurl+"'";
			  }
			  else
			  {
			  str1="<img src='"+ strurl+"'";
			   }
				str1=str1+" alt='"+document.PicForm.alttext.value+"' ";
				if(document.PicForm.width.value!=''&&document.PicForm.width.value!='0') str1=str1+"width='"+document.PicForm.width.value+"' ";
				if(document.PicForm.height.value!=''&&document.PicForm.height.value!='0') str1=str1+"height='"+document.PicForm.height.value+"' ";
				str1=str1+"border='"+document.PicForm.PicBorder.value+"' align='"+document.PicForm.aligntype.value+"' ";
				if(document.PicForm.vspace.value!=''&&document.PicForm.vspace.value!='0') str1=str1+"vspace='"+document.PicForm.vspace.value+"' ";
				if(document.PicForm.hspace.value!=''&&document.PicForm.hspace.value!='0') str1=str1+"hspace='"+document.PicForm.hspace.value+"' ";
				if(document.PicForm.styletype.value!='')	str1=str1+"style='"+document.PicForm.styletype.value+"'";
				str1=str1+">";
			objFiles.innerHTML=str1
			}
			function IsDigit()
			{
			  return ((event.keyCode >= 48) && (event.keyCode <= 57));
			}
			window.onunload=SetReturnValue;
			function SetReturnValue()
			{
				if (typeof(window.returnValue)!='string') window.returnValue='';
			}
			</script>
			</head>
			<BODY bgColor=menu topmargin=15 leftmargin=15 onload=ShowPicture();>
			<form name="PicForm" method="post" action="">
			  <table width=100% border="0" align="center" cellpadding="0" cellspacing="2">
				<tr>
				  <td> <FIELDSET align=left>
					<LEGEND align=left>输入图片参数</LEGEND>
					<TABLE border="0" align="center" cellpadding="0" cellspacing="3">
					  <tr> 
						<td width=350 align='center' bgcolor="#FFFFFF">
						 <div style="height:270; overflow:auto; width:380;" align="center"> 
							<span id="objFiles">
							</span> </div></td>
					  </tr>
					</TABLE>
					<table border="0" align=center cellpadding="0" cellspacing="3">
					  <tr> 
						<td colspan="2">图片地址：
						<%if request("Action")="edit" then%>
						  <input name="url" id="url" value='<%=request("src")%>' size=40 maxlength="200" onChange="ShowPicture();"> 
						<%else%> 
						  <input name="url" id="url" value='http://' size=30 maxlength="200" onChange="ShowPicture();"> 
						 <%end if%>
						  <%if FromUrl=1 and request("Action")<>"edit" Then%>
						  <input type="button" name="Button" value="选择图片" onClick="var TempReturnValue=OpenWindow('<%=AdminDir%>include/SelectPic.asp?ChannelID=<% = ChannelID %>&CurrPath=<% = CurrPath %>',500,290,window);if (TempReturnValue!='') document.PicForm.url.value=TempReturnValue;ShowPicture();" class=Anbutc> 
						  <%else%>
						  <input type="button" value="选择图片..." onClick="OpenThenSetValue('selectupfiles.asp?ChannelID=<% = ChannelID %>',500,360,window,document.all.url);">
						  <%End If%>
						</td>
					  </tr>
					  <tr> 
						<td> 说明文字： 
						  <input name="alttext" id=alttext size=20 maxlength="100" value="<%=KS.S("Alt")%>" onChange="ShowPicture();"> </td>
						<td>图片边框： 
						  <input name="PicBorder" id="PicBorder" ONKEYPRESS="event.returnValue=IsDigit();"  value="<%=KS.ChkClng(KS.S("Border"))%>" size=5 maxlength="2" onChange="ShowPicture();">
						  像素 </td>
					  </tr>
					  <tr> 
						<td> 特殊效果： 
						  <select name="styletype" id=styletype onChange="ShowPicture();">
							<option selected>不应用</option>
							<option value="filter:Alpha(Opacity=50)">半透明效果</option>
							<option value="filter:Alpha(Opacity=0, FinishOpacity=100, Style=1, StartX=0, StartY=0, FinishX=100, FinishY=140)">线型透明效果</option>
							<option value="filter:Alpha(Opacity=10, FinishOpacity=100, Style=2, StartX=30, StartY=30, FinishX=200, FinishY=200)">放射透明效果</option>
							<option value="filter:blur(add=1,direction=14,strength=15)">模糊效果</option>
							<option value="filter:blur(add=true,direction=45,strength=30)">风动模糊效果</option>
							<option value="filter:Wave(Add=0, Freq=60, LightStrength=1, Phase=0, Strength=3)">正弦波纹效果</option>
							<option value="filter:gray">黑白照片效果</option>
							<option value="filter:Chroma(Color=#FFFFFF)">白色为透明</option>
							<option value="filter:DropShadow(Color=#999999, OffX=7, OffY=4, Positive=1)">投射阴影效果</option>
							<option value="filter:Shadow(Color=#999999, Direction=45)">阴影效果</option>
							<option value="filter:Glow(Color=#ff9900, Strength=5)">发光效果</option>
							<option value="filter:flipv">垂直翻转显示</option>
							<option value="filter:fliph">左右翻转显示</option>
							<option value="filter:grays">降低彩色度</option>
							<option value="filter:xray">X光照片效果</option>
							<option value="filter:invert">底片效果</option>
						  </select> </td>
						<td>图片位置： 
						  <select name="aligntype" id=aligntype onChange="ShowPicture();">
							<option selected>默认位置 
							<option value="left"<%if lcase(ks.g("align"))="left" then response.Write(" selected")%>>居左 
							<option value="right" <%if lcase(ks.g("align"))="right" then response.Write(" selected")%>>居右 
							<option value="top"<%if lcase(ks.g("align"))="top" then response.Write(" selected")%>>顶部 
							<option value="middle"<%if lcase(ks.g("align"))="middle" then response.Write(" selected")%>>中部 
							<option value="bottom"<%if lcase(ks.g("align"))="bottom" then response.Write(" selected")%>>底部 
							<option value="absmiddle"<%if lcase(ks.g("align"))="absmiddle" then response.Write(" selected")%>>绝对居中 
							<option value="absbottom"<%if lcase(ks.g("align"))="absbottom" then response.Write(" selected")%>>绝对底部 
							<option value="baseline"<%if lcase(ks.g("align"))="baseline" then response.Write(" selected")%>>基线 
							<option value="texttop"<%if lcase(ks.g("align"))="texttop" then response.Write(" selected")%>>文本顶部 
							</select>
							
							</td>
					  </tr>
					  <tr> 
						<td>图片宽度： 
						  <input name="width" value="<%=request("width")%>" id=width2  ONKEYPRESS="event.returnValue=IsDigit();" size=4 maxlength="4" onChange="ShowPicture();">
						  像素</td>
						<td>图片高度： 
						  <input name="height" value="<%=request("height")%>" id="height3" onKeyPress="event.returnValue=IsDigit();" size=4 maxlength="4" onChange="ShowPicture();">
						  像素</td>
					  </tr>
					  <tr> 
						<td>上下间距： 
						  <input name="vspace" id=vspace  ONKEYPRESS="event.returnValue=IsDigit();" value="<%=KS.ChkClng(KS.S("vspace"))%>" size=4 maxlength="2" onChange="ShowPicture();">
						  像素</td>
						<td>左右间距： 
						  <input name="hspace" id=hspace onKeyPress="event.returnValue=IsDigit();"  value="<%=KS.ChkClng(KS.S("hspace"))%>" size=4 maxlength="2" onChange="ShowPicture();">
						  像素</td>
					  </tr>
					</table>
					</fieldset></td>
				  <td width=80 align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
					  <tr> 
						<td height="40"> <div align="center"> 
							<input name="cmdOK" type="button" id="cmdOK3" value="  确定  " onClick="OK();">
							<input name="UpFileName" type="hidden" id="UpFileName3" value="None">
						  </div></td>
					  </tr>
					  <tr> 
						<td height="40"> <div align="center"> 
							<input name="cmdCancel" type=button id="cmdCancel3" onClick="window.close();" value='  取消  '>
						  </div></td>
					  </tr>
					</table></td>
				</tr>
			  </table>
			</form>
			</body>
			</html>
<%			Set KS=Nothing

  End Sub
End Class
%>
 
