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
Set KSCls = New InsertUpfile
KSCls.Kesion()
Set KSCls = Nothing

Class InsertUpfile
        Private KS
		Private AdminDir
		Private ChannelID
		Private FromUrl,AllowUpFilesType
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
			AllowUpFilesType=KS.ReturnChannelAllowUpFilesType(1,0)
			CurrPath=KS.GetUpFilesDir()
			InstallDir=KS.Setting(3)
			AdminDir=InstallDir & KS.Setting(89)
			Set KS=Nothing
			%>
<HTML>
<HEAD>
<TITLE>插入附件下载</TITLE>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Editor.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../KS_Inc/Common.js"></script>
<script language="JavaScript">
function OK()
{

  var str1="";
  var strurl=document.UpfileForm.url.value;
  var sFilePic = getFilePic(strurl);
  if (!IsExt(strurl,'<%=AllowUpFilesType%>'))
  {
	  alert('文件类型不对，请重新选择！');
	  return;
  }
  if (strurl==""||strurl=="http://")
  {
  	alert("请先输入附件地址，或者上传附件！");
	document.UpfileForm.url.focus();
	return false;
  }
  else
  {
    str1="<img border=0 src='<%=InstallDir%>KS_Editor/images/FileIcon/"+sFilePic+"'> <A href="+document.UpfileForm.url.value+" class='newsContent'>"+document.UpfileForm.title.value+"</A>"
    window.returnValue = str1+"$$$"+document.UpfileForm.UpFileName.value;
    window.close();
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
// 按文件扩展名取图，并产生链接
function getFilePic(url)
{
	var sExt;
	sExt=url.substr(url.lastIndexOf(".")+1);
	sExt=sExt.toUpperCase();
	var sPicName;
	switch(sExt)
	{
	case "TXT":
		sPicName = "txt.gif";
		break;
	case "CHM":
	case "HLP":
		sPicName = "hlp.gif";
		break;
	case "DOC":
		sPicName = "doc.gif";
		break;
	case "PDF":
		sPicName = "pdf.gif";
		break;
	case "MDB":
		sPicName = "mdb.gif";
		break;
	case "GIF":
		sPicName = "gif.gif";
		break;
	case "JPG":
		sPicName = "jpg.gif";
		break;
	case "BMP":
		sPicName = "bmp.gif";
		break;
	case "PNG":
		sPicName = "pic.gif";
		break;
	case "ASP":
	case "JSP":
	case "JS":
	case "PHP":
	case "PHP3":
	case "ASPX":
		sPicName = "code.gif";
		break;
	case "HTM":
	case "HTML":
	case "SHTML":
		sPicName = "htm.gif";
		break;
	case "ZIP":
	case "RAR":
		sPicName = "zip.gif";
		break;
	case "EXE":
		sPicName = "exe.gif";
		break;
	case "AVI":
		sPicName = "avi.gif";
		break;
	case "MPG":
	case "MPEG":
	case "ASF":
		sPicName = "mp.gif";
		break;
	case "RA":
	case "RM":
		sPicName = "rm.gif";
		break;
	case "MP3":
		sPicName = "mp3.gif";
		break;
	case "MID":
	case "MIDI":
		sPicName = "mid.gif";
		break;
	case "WAV":
		sPicName = "audio.gif";
		break;
	case "XLS":
		sPicName = "xls.gif";
		break;
	case "PPT":
	case "PPS":
		sPicName = "ppt.gif";
		break;
	case "SWF":
		sPicName = "swf.gif";
		break;
	default:
		sPicName = "unknow.gif";
		break;
	}
	return sPicName;

}
</script>
</head>
<BODY bgColor=menu topmargin=15 leftmargin=15 >
<form name="UpfileForm" method="post" action="">
  <table width=100% border="0" cellpadding="0" cellspacing="2">
    <tr>
      <td> <FIELDSET align=left>
        <LEGEND align=left>附件参数</LEGEND>
        <TABLE border="0" cellpadding="0" cellspacing="3" >
		  <TR>
            <TD height="17" >名称： <INPUT name="title" id=title value=" 附件下载" size=30>
            </td>
          </TR>
		<TR>
			<TD height="17" >地址： <INPUT name="url" id=url value="http://" size=30 >
				<%if FromUrl=1 Then%>
				<input type="button" name="Button" value="选择附件" onClick="var TempReturnValue=OpenWindow('<%=AdminDir%>include/SelectPic.asp?ChannelID=<% = ChannelID %>&CurrPath=<% = CurrPath %>',500,290,window);if (TempReturnValue!='') document.all.url.value=TempReturnValue" class=Anbutc> 
				<%End If%>
				</td>
		</TR>
          <TR>
            <TD align=center>支持格式为：<%=left(AllowUpFilesType,50)%>等</TD>
          </TR>
        </TABLE>
        </fieldset></td>
      <td width=80 align="center"><input name="cmdOK" type="button" id="cmdOK" value="  确定  " onClick="OK();"> 
        <br> <br>
        <input name="UpFileName" type="hidden" id="UpFileName2" value="None"> 
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
