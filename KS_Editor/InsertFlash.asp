<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Plus/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
Dim KSCls
Set KSCls = New InsertFlash
KSCls.Kesion()
Set KSCls = Nothing

Class InsertFlash
        Private KS
		Private AdminDir
		Private ChannelID
		Private FromUrl
        Private CurrPath,InstallDir
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		End Sub
        Public Sub Kesion()
			FromUrl=KS.ChkClng(KS.S("FromUrl"))
	
	       IF FromUrl=1 Then  '后台调用，检查是否登录
				Dim KSLoginCls
				Set KSLoginCls = New LoginCheckCls1
				KSLoginCls.Run()
				Set KSLoginCls= Nothing
			End IF

			ChannelID = KS.S("ChannelID")
			CurrPath=KS.GetUpFilesDir()
			AdminDir=KS.GetDomain & KS.Setting(89)
			%>
			<HTML>
			<HEAD>
			<TITLE>插入FLASH文件</TITLE>
			<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
			<link href="Editor.css" rel="stylesheet" type="text/css">
			<script language="JavaScript" src="../KS_Inc/Common.js"></script>
			<script language="JavaScript">
			function OK()
			{
			  var str1="";
			  var strurl=document.FlashForm.url.value;
			  if (!IsExt(strurl,'swf')&&!IsExt(strurl,'flv'))
			  {
				  //alert('文件类型不对，请重新选择！');
				//  return;
			  }
			  if (strurl==""||strurl=="http://")
			  {
				alert("请先输入FLASH文件地址，或者上传FLASH文件！");
				document.FlashForm.url.focus();
				return false;
			  }
			  else
			  {
			   var typeflag;
			   for (var i=0;i<document.FlashForm.typeflag.length;i++){
				 var KM = document.FlashForm.typeflag[i];
				if (KM.checked==true)	   
					typeflag= KM.value
				}
				
			   if (typeflag==1){
				str1="<embed src="+document.FlashForm.url.value+" width="+document.FlashForm.width.value+" height="+document.FlashForm.height.value+" type=application/x-shockwave-flash></embed>"
				}
				else
				{
				str1="<EMBED allowScriptAccess='never' allowNetworking='internal'  pluginspage=http://www.macromedia.com/go/getflashplayer src='<%=KS.Setting(3)%>ks_inc/vcastr.swf?vcastr_file="+document.FlashForm.url.value+"' width="+document.FlashForm.width.value+" height="+document.FlashForm.height.value+" type=application/x-shockwave-flash  quality='high' wmode='transparent' showMovieInfo='0'></EMBED>";
				}
				window.returnValue = str1+"$$$"+document.FlashForm.UpFileName.value;
				window.close();
			  }
			}
			function swfShowChange(){
			   if(document.FlashForm.url.value=="http://"){
			   document.FlashForm.url.value=""
			   }
			   if(document.FlashForm.url.value!='')  
			   {
			   var typeflag;
			   for (var i=0;i<document.FlashForm.typeflag.length;i++){
				 var KM = document.FlashForm.typeflag[i];
				if (KM.checked==true)	   
					typeflag= KM.value
				}
				
			   if (typeflag==1){
			   objFiles.innerHTML="<embed src="+document.FlashForm.url.value+" width="+document.FlashForm.width.value+" height="+document.FlashForm.height.value+" type=application/x-shockwave-flash></embed>"}
			   else{
				objFiles.innerHTML="<EMBED allowScriptAccess='never' allowNetworking='internal'  pluginspage=http://www.macromedia.com/go/getflashplayer src='<%=KS.Setting(3)%>ks_inc/vcastr.swf?vcastr_file="+document.FlashForm.url.value+"' width="+document.FlashForm.width.value+" height="+document.FlashForm.height.value+" type=application/x-shockwave-flash  quality='high' wmode='transparent' showMovieInfo='0'></EMBED>";
			   }
			   }
			}
			function IsExt(url,opt)
			{  
				var sTemp; 
				var b=false; 
				var s=opt.toUpperCase().split("|");  
				for (var i=0;i<s.length ;i++ ) 
				{ 
					sTemp=url.substr(url.length-s[i].length-1); 
					sTemp=sTemp.toUpperCase();
					s[i]="."+s[i];
					if (s[i]==sTemp)
					{
						b=true;
						break;
					}
				}
				return b;
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
			<BODY bgColor=menu topmargin=15 leftmargin=15 >
			<form name="FlashForm" method="post" action="">
			  <table width=100% border="0" cellpadding="0" cellspacing="2">
				<tr>
				  <td> <FIELDSET align=left>
					<LEGEND align=left>FLASH动画参数</LEGEND>
					<table height="280" border="0" align="center" cellpadding="0" cellspacing="3" >
					  <tr> 
						<td width=350 align='center'> 
						<div style="height:250; overflow:auto; width:380" align="center"> 
							<span id="objFiles"> <img src="images/FileType/Flash.gif"></span> 
						  </div>
						</td>
					  </tr>
					</table>
					<TABLE border="0" align="center" cellpadding="0" cellspacing="3" >
					  <TR>
					    <TD height="17" >类型：
						<input type="radio" value="1" name='typeflag' checked>
						flash(swf)
						<input type="radio" value="2" name='typeflag'>
						flv视频
						</td>
				      </TR>
					  <TR>
						<TD height="17" >地址： <INPUT name="url" id=url value="http://" size=30 onChange="javascript:swfShowChange()">
					  <%if FromUrl=1 Then%>
						  <input type="button" name="Button" value="选择动画" onClick="var TempReturnValue=OpenWindow('<%=AdminDir%>include/SelectPic.asp?ChannelID=<% = ChannelID %>&CurrPath=<% = CurrPath %>',500,290,window);if (TempReturnValue!='') document.all.url.value=TempReturnValue;swfShowChange()" class=Anbutc> 
						<%else%>
						  <input type="button" value="选择Flash..." onClick="OpenThenSetValue('selectupfiles.asp',500,360,window,document.all.url);">
						  <%End If%>						</td>
					  </TR>
					  <TR>
						<TD >宽度：
						  <INPUT name="width" id=width ONKEYPRESS="event.returnValue=IsDigit();" value=300 size=7 maxlength="4" onChange="javascript:swfShowChange()"> 
						  &nbsp;&nbsp;高度：
						  <INPUT name="height" id=height ONKEYPRESS="event.returnValue=IsDigit();" value=200 size=7 maxlength="4" onChange="javascript:swfShowChange()"></TD>
					  </TR>
					</TABLE>
					</fieldset></td>
				  <td width=80 align="center"><input name="cmdOK" type="button" id="cmdOK" value="  确定  " onClick="OK();"> 
					<br> <br>
					<input name="cmdCancel" type=button id="cmdCancel" onClick="window.close();" value='  取消  '> 
					<input name="UpFileName" type="hidden" id="UpFileName2" value="None"></td>
				</tr>
			  </table>
			</form>
			</body>
			</html>
<%
  End Sub
End Class
%> 
