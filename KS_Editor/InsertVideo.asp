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
'****************************************************
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
Dim KSCls
Set KSCls = New InsertVideo
KSCls.Kesion()
Set KSCls = Nothing

Class InsertVideo
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
			Set KS=Nothing
			%>
				<HTML><HEAD><TITLE>插入视频文件</TITLE>
				<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
				<link rel="stylesheet" type="text/css" href="Editor.css">
				<script language="JavaScript" src="../KS_Inc/Common.js"></script>
				<script language="JavaScript">
				function OK(){
				  var str1="";
				  var autoplayvalue;
				  var strurl=document.VideoForm.url.value;
				  if (!IsExt(strurl,'avi|wmv|asf|mpg|mp3'))
				  {
					  alert('文件类型不对，请重新选择！');
					  document.VideoForm.url.focus();
					  return;
				  }
				  if (strurl==""||strurl=="http://")
				  {
					alert("请先输入视频文件地址，或者上传视频文件！");
					document.VideoForm.url.focus();
					return false;
				  }
				  else
				  {
				  
				    if (VideoForm.autoplay.checked==true)
					 {autoplayvalue=1;}
					else
					 {autoplayvalue=0;}
					str1="<embed width=\""+document.VideoForm.width.value+"\" height=\""+document.VideoForm.height.value+"\" autostart=\""+autoplayvalue+"\" src="+document.VideoForm.url.value+">";
					window.returnValue = str1+"$$$"+document.VideoForm.UpFileName.value;
					window.close();
				  }
				}
				 function windowplay(){
				   var str1="";
				   if(document.VideoForm.url.value=="http://"){
				   document.VideoForm.url.value=""
				   }  
				str1="<embed width=\""+document.VideoForm.width.value+"\" height=\""+document.VideoForm.height.value+"\" autostart=\"1\" src="+document.VideoForm.url.value+">";
				   objFiles.innerHTML=str1
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
				<form name="VideoForm" method="post" action="">
				  <table width=100% border="0" cellpadding="0" cellspacing="2">
					<tr>
					  <td> <FIELDSET align=left>
						<LEGEND align=left>视频文件参数</LEGEND>
						<TABLE border="0" align="center" cellpadding="0" cellspacing="3">
						  <tr> 
							<td width=350 align='center'>
							<div style="height:280; overflow:auto; width:380" align="center"> 
								<span id="objFiles"> 
								
								<embed width="350" height="270" autostart="1" src="1.wmv">
								
								
							  </span>
							  </div>
							</td>
						  </tr>
						</TABLE>
						<TABLE border="0" align="center" cellpadding="0" cellspacing="3">
						  <TR>
							<TD >地址：
							  <INPUT name="url" id=url  value="http://" size=30 onChange="javascript:windowplay()">						                                <%if FromUrl=1 Then%>
							  <input type="button" name="Button" value="选择视频文件" onClick="var TempReturnValue=OpenWindow('<%=AdminDir%>include/SelectPic.asp?ChannelID=<% = ChannelID %>&CurrPath=<% = CurrPath %>',500,290,window);if (TempReturnValue!='') document.VideoForm.url.value=TempReturnValue;windowplay();" class=Anbutc> 
							 <%else%>
						  <input type="button" value="选择视频..." onClick="OpenThenSetValue('selectupfiles.asp',500,360,window,document.all.url);">
						  <%End If%>
							</td>
						  </TR>
						  <TR>
							<TD >宽度：
							  <INPUT name="width" id=width onChange="javascript:windowplay()" ONKEYPRESS="event.returnValue=IsDigit();" value=350 size=8 maxlength="4">
							  &nbsp;&nbsp;&nbsp;高度： 
							  <INPUT id=height onChange="javascript:windowplay()" ONKEYPRESS="event.returnValue=IsDigit();" value=280 size=8 maxlength="4">						      　
							  自动播放：
							  <input name="autoplay" type="checkbox" id="autoplay" value="1" checked></TD>
						  </TR>
						  <TR>
							<TD align=center>支持格式为：mp3、avi、wmv、mpg、asf</TD>
						  </TR>
						</TABLE>
						</fieldset></td>
					  <td width=80 align="center"><input name="cmdOK" type="button" id="cmdOK" value="  确定  " onClick="OK();"> 
						<br> <br> <input name="UpFileName" type="hidden" id="UpFileName2" value="None"> 
						<input name="cmdCancel" type=button id="cmdCancel" onClick="window.close();" value='  取消  '></td>
					</tr>
				  </table>
				</form>
				</body>
				</html>
<%
  End Sub
End Class
%> 
