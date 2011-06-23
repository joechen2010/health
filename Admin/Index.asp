<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
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
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		    Call CheckSetting()
			Select Case KS.G("Action")
			 Case "Head" Call KS_Head()
			 Case "Left" Call KS_Left()
			 Case "Main" Call KS_Main()
			 Case "Foot" Call KS_Foot()
			 Case "ver"  Call GetRemoteVer()
			 Case Else  Call KS_Index()
			End Select
		End Sub
		Sub KS_Index()
		With Response
		.Write "<html>"
		.Write "<head>"
		.Write "<title>" & KS.Setting(0) & "---网站后台管理</title>"
		.Write "<script language=""JavaScript"">" & vbCrLf
		.Write "<!--" & vbCrLf
		.Write "   //保存复制,移动的对象,模拟剪切板功能" & vbCrLf
		.Write "  function CommonCopyCutObj(ChannelID, PasteTypeID, SourceFolderID, FolderID, ContentID)" & vbCrLf
		.Write "  {" & vbCrLf
		.Write "   this.ChannelID=ChannelID;             //频道ID" & vbCrLf
		.Write "   this.PasteTypeID=PasteTypeID;         //操作类型 0---无任何操作,1---剪切,2---复制" & vbCrLf
		.Write "   this.SourceFolderID=SourceFolderID;   //所在的源目录" & vbCrLf
		.Write "   this.FolderID=FolderID;               //目录ID" & vbCrLf
		.Write "   this.ContentID=ContentID;             //文章或图片等ID" & vbCrLf
		.Write "  }" & vbCrLf
		.Write "  function CommonCommentBack(FromUrl)" & vbCrLf
		.Write "  {" & vbCrLf
		.Write "    this.FromUrl=FromUrl;             //保存来源页的地址" & vbCrLf
		.Write "  }" & vbCrLf
		.Write "  //初始化对象实例" & vbCrLf
		.Write " var CommonCopyCut=null;" & vbCrLf
		.Write " var CommonComment=null;" & vbCrLf
		.Write " var DocumentReadyTF=false;" & vbCrLf
		.Write " function document.onreadystatechange()" & vbCrLf
		.Write "  {" & vbCrLf
		.Write "    if (DocumentReadyTF==true) return;" & vbCrLf
		.Write "    CommonCopyCut=new CommonCopyCutObj(0,0,0,'0','0');" & vbCrLf
		.Write "    CommonComment=new CommonCommentBack(0);" & vbCrLf
		.Write "    DocumentReadyTF=true;" & vbCrLf
		.Write "  }" & vbCrLf
		.Write "//-->" & vbCrLf
		.Write "</script>" & vbCrLf
		.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbCrLf
		.Write "</head>" & vbCrLf

		.Write "<frameset rows=""40,*,35"" border=""0"" frameborder=""0"" framespacing=""0"">" & vbcrlf
		.Write "	<frame src=""Index.asp?Action=Head"" name=""FrameTop"" id=""FrameTop"" noresize scrolling=""no""  frameborder=""no""></frame>" & vbcrlf
		.Write "  <frameset cols=""205,*"" name=""FrameMain"" border=""0"" frameborder=""0"" framespacing=""0"">" & vbcrlf
		.Write "		<frame src=""Index.asp?Action=Left"" name=""LeftFrame"" noresize frameborder=""no"" scrolling=""yes"" marginwidth=""0"" marginheight=""0""></frame>" &vbcrlf
		.Write "         <frameset rows=""*,26"" border=""0"" frameborder=""0"" framespacing=""0"">" & vbCrLf
		.Write "            <frame src=""Index.asp?action=Main""  noresize name=""MainFrame"" id=""MainFrame"" frameborder=""no"" scrolling=""yes"" marginwidth=""0"" marginheight=""0""></frame>" & vbCrLf
		.Write "            <frame src=""KS.Split.asp?ButtonSymbol=Disabled&OpStr=" & Server.URLEncode("<font color=red>系统管理中心</font>") & """ name=""BottomFrame"" ID=""BottomFrame"" frameborder=""no"" scrolling=""no"" marginwidth=""0"" marginheight=""0""></frame>" & vbCrLf
		.Write "        </frameset>" & vbCrLf
		.Write "  </frameset>" & vbcrlf
		.Write "  <frame src=""Index.asp?Action=Foot"" name=""FrameBottom"" id=""FrameBottom"" noresize frameborder=""no"" scrolling=""no"" marginwidth=""0"" marginheight=""0""></frame>" & vbCrLf
		.Write "</frameset>" & vbcrlf
		.Write "<noframes>您的浏览器版本太低,请安装IE5.5或以上版本!</noframes>" & vbcrlf
		.Write "</html>" & vbCrLf
		End With
		End Sub
		
		Public Sub KS_Head()
			 On Error Resume Next
			 With Response
			 .Buffer = True
			If Trim(Request.ServerVariables("HTTP_REFERER")) = "" Then
				.Write "<br><p align=center><font color='red'>对不起，为了系统安全，不允许直接输入地址访问本系统的后台管理页面。</font></p>"
				Exit Sub
			End If
			
			.Write "<html>"
		    .Write "<script language='JavaScript' src='../KS_Inc/common.js'></script>"
		    .Write "<script language='JavaScript' src='../KS_Inc/jquery.js'></script>"
			.Write "<link href=""Skin/Style"& KS.C("SkinID") &".CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "<meta http-equiv=Content-Type content=""text/html; charset=GB2312"">"
			.Write "<script language=""javascript"">"& vbcrlf
			.Write " function out(src){"& vbcrlf
			.Write " if(confirm('确定要退出吗？')){"& vbcrlf
			.Write " return true;	"& vbcrlf
			.Write " }"& vbcrlf
			.Write "return false;"& vbcrlf
			.Write " }"& vbcrlf
			.Write " function getNewMessage()"& vbcrlf
			.Write " {"& vbcrlf
			.Write "  var url = '../user/UserAjax.asp';"   & vbCrLf      
			.Write "  jQuery.get(url,{action:'GetAdminMessage'},function(d){jQuery('#newmessage').html(d);});" & vbCrLf
			.Write " }"
			.Write "setTimeout('getNewMessage()', 2000);"
            .Write "</script>"
			
			.Write "</head>"
			.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"" scroll=""no"" class=""head"">"
			.Write "<div id='ajaxmsg' style='text-align:center;background-color: #ffffee;border: 1px #f9c943 solid;position:absolute; z-index:1; left: 200px; top: 5px;display:none;'> <img src='images/loading.gif'> 请稍候,正在执行您的请求...  </div>"
			.Write "<table width=""100%""  border=""0"" cellpadding=""0"" cellspacing=""0"">"
			.Write "<tr>"
			.Write "    <td height=""30"">"
			.Write "        <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
			.Write "          <tr>"
			.Write "            <td width='220'><font class='logo'>&nbsp;&nbsp;KesionCMS V6.5</font></td>"
			.Write "            <td><table width=""100%"" height=""100%"" border=""0"">"
			 .Write "             <tr>"
			 Dim KSAnnounceDisplayFlag
			 If Instr(KS.Setting(16),"1")=0 Then
			  .Write "                <td class=""font_text"" width=""40%""><script language=""JavaScript"" src=""../ks_inc/time/3.js""></script></td>"

			  KSAnnounceDisplayFlag=" style=""display:none"""
			 Else
			  KSAnnounceDisplayFlag=""
			 End If
			 .write "                 <td " & KSAnnounceDisplayFlag & " class=""font_text"" align=""right""><font color=#ffffff>官方公告：</font></td>"
			 .Write "                 <td " & KSAnnounceDisplayFlag & "  width=""40%"">"
			 .Write "<iframe scrolling=no src=""http://www.kesion.com/websystem/GetofficialInfo.asp"" name=""ShowAnnounce"" id=""ShowAnnounce"" height=""22"" WIDTH=""100%"" marginheight=""0"" marginwidth=""0"" frameborder=""0"" align=""middle"" allowtransparency=""true""></iframe>"
			 .Write "</td>"
			 
			.Write "                <td class=""font_text"" align=""right""> [<a href=""" & KS.GetDomain &""" target=""_blank"" class=""white"">网站首页</a>] [<a href=""" & KS.GetDomain &"User/index.asp?User_Message.asp?action=inbox"" target=""_blank"" class=""white"">查看短信</a><span id='newmessage'>(<font color=#ff0000>0</font>)</span>] "
			If KS.ReturnPowerResult(0, "KMUA10010") Then
			.Write "[<a href=""#"" onClick=""OpenWindow('KS.Frame.asp?Url=KS.Admin.asp&Action=SetPass&PageTitle=" & server.URLEncode("修改后台登录密码") & "',360,175,window);"" class=""white"">修改密码</a>] "
			End If
			If KS.ReturnPowerResult(0, "KMST20000") Then
			.Write "[<a href=""KS.CleanCache.asp"" target=""MainFrame"" class=""white"">更新缓存</a>] "
			End If
			.WRite "[<a href=""Login.asp?Action=LoginOut"" target=""_top"" onClick=""return out(this)""  class=""white"">安全退出</a>]"
			
			.Write "               </td>"

			.Write "              </tr>"
			.Write "            </table></td>"
			.Write "          </tr>"
			.Write "        </table>"
			.Write "      </td>"
			.Write "  </tr>"
			.Write "</table>"
			.Write "</body>"
			.Write "</html>"
			 If KS.S("C")="1" Then
					 On Error Resume Next
					 Dim FileContent
					 FileContent=KS.ReadFromFile("../KS_Inc/ajax.js")
					 FileContent=GetAjaxInstallDir(FileContent,installdir)
					 Call KS.WriteTOFile("../KS_Inc/ajax.js", FileContent)
					 If Err Then
					  err.clear
					 End If
			 End If		
			End With
			End Sub

			Function GetAjaxInstallDir(Content,byval installdir)
			 Dim regEx, Matches, Match
			 Set regEx = New RegExp
			 regEx.Pattern="var installdir='[\S]*';"
			 regEx.IgnoreCase = True
			 regEx.Global = True
			 Set Matches = regEx.Execute(Content)
			 If Matches.count > 0 Then
			  GetAjaxInstallDir=Replace(content,Matches.item(0),"var installdir='" & installdir & "';")
			 Else
			  GetAjaxInstallDir="var installdir='/';"
			 end if
		End Function
		
		
		Public Sub KS_Left()
		Dim SQL,I,ModelXML
		Dim RSC:Set RSC=Conn.Execute("Select ChannelID,ChannelName,ChannelTable,ItemName,BasicType,ModelEname,ChannelStatus From KS_Channel Order By ChannelID")
		If Not RSC.Eof Then
		  SQL=RSC.GetRows(-1)
		  Set ModelXML=KS.ArrayToxml(SQL,RSC,"row","ModelXML")
		End If
		RSC.Close:Set RSC=Nothing
		
		
		on error resume next

		With Response
		.Write "<script language=""javascript"">"
		.Write " var ChannelID=null;" & vbcrlf
		.Write " var BasicType=null;" & vbcrlf
		For I=0 To Ubound(SQL,2)
		 .Write " var SearchPower" & SQL(0,I) & "='" & KS.ReturnPowerResult(SQL(0,I), "M"&SQL(0,I)&"10008")&"';    //搜索权限" & vbCrLf
       Next
		.Write " var SearchSpecialPower='" & KS.ReturnPowerResult(0, "KMSP10004") & "';    //搜索专题权限" & vbCrLf
		.Write " var SearchLinkPower='" & KS.ReturnPowerResult(0, "KMCT10001") & "';       //搜索友情链接的权限" & vbCrLf
		.Write " var SearchAdminPower='" & KS.ReturnPowerResult(0, "KMUA10001") & "';      //搜索管理员权限" & vbCrLf
		.Write " var SearchSysLabelPower='" & KS.ReturnPowerResult(0, "KMTL10001") & "';   //搜索系统函数标签权限" & vbCrLf
		.Write " var SearchDIYFunctionLabelPower='" & KS.ReturnPowerResult(0, "KMTL10002") & "';   //搜索自定义函数标签权限" & vbCrLf
		.Write " var SearchFreeLabelPower='" & KS.ReturnPowerResult(0, "KMTL10003") & "';  //搜索自定义静态标签权限" & vbCrLf
		.Write " var SearchSysJSPower='" & KS.ReturnPowerResult(0, "KMTL10004") & "';      //搜索系统JS权限" & vbCrLf
		.Write " var SearchFreeJSPower='" & KS.ReturnPowerResult(0, "KMTL10005") & "';     //搜索自由JS权限" & vbCrLf
		.Write "</script>"
		.Write "<script language=""JavaScript"" src=""Include/SetFocus.js""></script>"
		.Write "<script language='JavaScript' src='../KS_Inc/common.js'></script>"
		.Write "<script language='JavaScript' src='../KS_Inc/Jquery.js'></script>"
		%>
		<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
         <html xmlns="http://www.w3.org/1999/xhtml">
		<head><title>科汛网站管理系统V6.0后台</title>
		<meta http-equiv=Content-Type content="text/html; charset=gb2312">
		<link href="Skin/Style<%= KS.C("SkinID")%>.css" type=text/css rel=stylesheet>
		</head>
		<body leftmargin="0" topmargin="0" class="leftbody">
		<script language="JavaScript">
		//Search For Kesion CMS
		//Version 6.0
		//Powered By Kesion.Com
		//var normal='slategray';   //color;
		var normal='#26517B';     //color;
		var zindex=10000;         //z-index;
		var openTF=false;
		var width=160,height=window.document.body.offsetHeight-15;
		var left=0,top=0,title='搜索小助理';
		var SearchBodyStr=''
						   +'<table width="100%" border="0" cellspacing="0" cellpadding="0">'
						   +'<form name="searchform" target="MainFrame" method="post">'
						   +'<tr> '
						   +'<td height="25"><strong>按下面任意或全部条件进行搜索</strong></td>'
						   +' </tr>'
						   +'<tr><td height="25">全部或部分关键字</td></tr>'
						   +'<tr><td height="25"><input style="width:95%" type="text" id="KeyWord" name="KeyWord"></td></tr>'
						   +'  <tr><td height="25">搜索范围</td></tr>'
						   +'  <tr><td height="25"> <select style="width:95%" id="SearchArea" name="SearchArea" onchange="SetSearchTypeOption(this.options[this.selectedIndex].text)">'
						   +'     </select></td></tr>' 
						   +'<tr><td height="25">搜索类型</td></tr>'
						   +'<tr><td height="25"><select style="width:95%" id="SearchType" name="SearchType">'
						   +'</select></td></tr>'
						   +'  <tr id="DateArea" onclick="setstatus(this)" style="cursor:pointer"><td height="25"><strong>什么时候修改的?</strong></td></tr>'
						   +'  <tr style="display:none"><td height="25">开始日期<input type="text" readonly style="width:80%" name="StartDate" id="StartDate">'
						   +'  <span style="cursor:pointer" onClick=OpenThenSetValue("Include/DateDialog.asp",160,170,window,document.all.StartDate);document.all.StartDate.focus();><img src="Images/date.gif" border="0" align="absmiddle" title="选择日期"></span></td></tr>'
						   +'  <tr style="display:none"><td height="25">结束日期<input type="text" readonly style="width:80%" name="EndDate" id="EndDate">'
						   +'  <span style="cursor:pointer" onClick=OpenThenSetValue("Include/DateDialog.asp",160,170,window,document.all.EndDate);document.all.EndDate.focus();><img src="Images/date.gif" border="0" align="absmiddle" title="选择日期"></span></td></tr>'
						   +'  <tr><td height="40" align="center"><input type="submit" name="SearchButton" value="开始搜索" onclick="return(SearchFormSubmit())"></td></tr>'
						   +'</form>'
						   +'  <tr><td><strong>使用说明:</strong></td></tr>'
						   +'  <tr><td> ① 您可以利用本搜索助理来搜索文章、图片、下载Flash、专题、标签、JS等,但不能搜索（目录）诸如频道名称、栏目名称，标签目录等</td></tr>'
						   +'  <tr><td> ② 按 <font color=red>Ctrl+F</font> 可以快速进行打开或关闭搜索小助理</td></tr>'
						   +'</table>'
				var str=""
					   +"<div id='SearchBox' style='display:none;z-index:" + zindex + ";width:" + width + ";height:" + height + ";left:" + left + ";top:" + top + ";background-color:" + normal + ";color:black;font-size:12px;font-family:Verdana, Arial, Helvetica, sans-serif;position:absolute;cursor:default;border:10px solid " + normal + ";'>"
					   + "<div style='background-color:" + normal + ";width:" + (width) + ";height:16;color:white;'>"
									   + "<span style='width:" + (width-2*12-4) + ";padding-left:3px;font-weight:bold;'>" + title + "</span>"
									   + "&nbsp;&nbsp;<span id='Close' style='padding-right:0px;width:20;border-width:0px;color:white;font-family:webdings;' onclick='CloseSearchBox(this)'>r</span>"
					   + "</div>"
					   + "<div style='width:170;overflow:auto;height:" + (height-20-4) + ";background-color:white;line-height:14px;word-break:break-all;padding:6px;'>" + SearchBodyStr + "</div>"
					   + "</div>"
					   + "<div style='display:none;width:" + width + ";height:" + height + ";top:" + top + ";left:" + left + ";z-index:" + (zindex-1) + ";position:absolute;background-color:black;filter:alpha(opacity=40);'></div>";
		//关闭;
		function CloseSearchBox(el)
		{   if (el.id=='Close'){ var twin = el.parentNode.parentNode;
				var shad = twin.nextSibling;
					twin.style.display = "none";
					shad.style.display = "none";
					openTF=false;
					SearchBodyStr=null;
					str=null;
			   }
		}
		function initial()
		{if (!openTF){
		 document.body.insertAdjacentHTML("beforeEnd",str);
		 openTF=true;}
		}
		//初始化;
		function initializeSearch(SearchArea,sChannelID,sBasicType)
		{
		 initial();
		 initialSearchAreaOption(SearchArea);
		 ChannelID=sChannelID;
		 BasicType=sBasicType;
		if (jQuery('#SearchBox')[0].style.display=='none')
		 {
		  jQuery('#SearchBox').show('fast');
		  if (document.forms[0].disabled==false) document.forms[0].focus();
		 }
		 else
		 jQuery('#SearchBox').hide('fast');
		}
		<%
		 Dim ModelList,ModelEList,ChannelIDList
		 For I=0 To Ubound(SQL,2)
		  If SQL(0,I)<>6 and SQL(6,I)=1 Then
			  ModelList=ModelList & "'" & SQL(1,I) & "',"
			  ModelElist=ModelElist & "'" & SQL(4,I) & "',"
			  ChannelIDList=ChannelIDList & "'" & SQL(0,I) &"',"
		  End If
		 Next
		%>
		var sTextArr,ChannelIDArr;
		function initialSearchAreaOption(SearchArea)
		{	 var EF=false;
			 sTextArr=new Array(<%=ModelList%>'专题中心','友情链接站点','系统函数标签','自定义函数标签','自定义静态标签','系统 JS','自由 JS','管理员')
			 ChannelIDArr=new Array(<%=ChannelIDList%>'专题中心','友情链接站点','系统函数标签','自定义函数标签','自定义静态标签','系统 JS','自由 JS','管理员')
			 var valueArr=new Array(<%=ModelElist%>'Special','Link','SysLabel','DIYFunctionLabel','FreeLabel','SysJS','FreeJS','Manager')
			  for(var i=0;i<valueArr.length;++i)
			   if (SearchArea==sTextArr[i]){ 
				  EF=true;
				  break;
				 }
			  if (!EF) return false; 
			  jQuery('#KeyWord').val('');
			  jQuery('#SearchArea').empty();
			  for (var i=0;i<sTextArr.length;++i)
				{
				   if (SearchArea==sTextArr[i]){
					jQuery('#SearchArea').append("<option value='"+valueArr[i]+"' selected>"+sTextArr[i]+"</option>");
					}else{
					jQuery('#SearchArea').append("<option value='"+valueArr[i]+"'>"+sTextArr[i]+"</option>");
					}
				} 
			//进行权限检查,对没有权限的搜索模块,进行屏蔽	
			 var n=0;
			for (var i=1000;i<sTextArr.length;++i)
			   {   var removeTF=false;
				   if (valueArr[i]!=SearchArea)
				  { 
				  
				  <%For I=0 To Ubound(SQL,2)
				    If SQL(6,I)=1 Then 
				   %>
				  if (SearchPower<%=SQL(0,i)%>=='False')
					   removeTF=true;
				  <%
				    End If
				  NEXT%>
		 
					if (valueArr[i]=='Special' && SearchSpecialPower=='False')  
					   removeTF=true;
					if (valueArr[i]=='Link' && SearchLinkPower=='False')  
					   removeTF=true;
					if (valueArr[i]=='SysLabel' && SearchSysLabelPower=='False')
					   removeTF=true;
					if (valueArr[i]=='DIYFunctionLabel' && SearchDIYFunctionLabelPower=='False')
					   removeTF=true;
					if (valueArr[i]=='FreeLabel' && SearchFreeLabelPower=='False')
					   removeTF=true;
					if (valueArr[i]=='SysJS' && SearchSysJSPower=='False')
					   removeTF=true;
					if (valueArr[i]=='FreeJS' && SearchFreeJSPower=='False')
					   removeTF=true;
					if (valueArr[i]=='Manager' && SearchAdminPower=='False')
					   removeTF=true;
				   }
				  if (removeTF==true)  
					{document.all.SearchArea.options.remove(i-n);
					 n++;
					}	
			   }
			SetSearchTypeOption(SearchArea); 
		}
		function SetSearchTypeOption(AreaType)
		{	
			  //改变选择范围时，取得正确的模型ID
			  for(var i=0;i<sTextArr.length;++i)
			   if (AreaType==sTextArr[i]) 
				{ 
				  ChannelID=ChannelIDArr[i];
				  break;
				 }

			var TextArr=new Array();
			jQuery('#SearchType').empty();
		  switch (AreaType)
		  {
		   <%For I=0 To Ubound(SQL,2)
		      If SQL(6,I)=1 Then 
			%>
			case '<%= SQL(1,I)%>':
				 if (SearchPower<%= SQL(0,I)%>=='False')          //搜索权限检查
				 {
				  DisabledSearchFluctuation(true);
				  return;
				 }
				 else
				 {
				  DisabledSearchFluctuation(false);
				  jQuery('#DateArea').show();
				  TextArr=new Array('<%=SQL(3,I)%>标题','<%=SQL(3,I)%>内容','<%=SQL(3,I)%>关键字','<%=SQL(3,I)%>作者','<%=SQL(3,I)%>录入')
				  }
				  break;
		   <% End If
		   Next%>
			case '专题中心':
				 if (SearchSpecialPower=='False')        //搜索专题权限检查
				 {
				   DisabledSearchFluctuation(true);
				   return;
				 }
				 else
				 {
				  DisabledSearchFluctuation(false);
				  jQuery('#DateArea').show();
				  TextArr=new Array('专题名称','简要说明')
				 }
				 break;
			case '友情链接站点':
				 if (SearchLinkPower=='False')       //搜索友情链接站点权限检查
				 {
				   DisabledSearchFluctuation(true);
				   return;
				 }
				 else{
				  DisabledSearchFluctuation(false);
				  jQuery('#DateArea').show();
				  TextArr=new Array('站点名称','站点描述')
				 }
				 break;
			case '系统函数标签':
				 if (SearchSysLabelPower=='False')       //搜索系统标签权限检查
				 {
				   DisabledSearchFluctuation(true);
				   return;
				 }else{
				  DisabledSearchFluctuation(false);
				 jQuery('#DateArea').show();
				 TextArr=new Array('系统标签名称','系统标签描述')
				 }
				 break;
			case '自定义函数标签':
				 if (SearchDIYFunctionLabelPower=='False')       //搜索自定义函数标签权限检查
				 {
				   DisabledSearchFluctuation(true);
				   return;
				 }else{
				  DisabledSearchFluctuation(false);
				 jQuery('#DateArea').show();
				 TextArr=new Array('自定义函数标签名称','自定义函数标签描述')
				 }
				 break;
			case '自定义静态标签':
				 if (SearchFreeLabelPower=='False')       //搜索自定义静态标签权限检查
				 {
				   DisabledSearchFluctuation(true);
				   return;
				 }else{
				 jQuery('#DateArea').show()
				 TextArr=new Array('自定义静态标签名称','自定义静态标签描述','自定义静态标签内容')
				 }
				 break;
			case '系统 JS':
				 if (SearchSysJSPower=='False')       //搜索系统JS权限检查
				 {
				   DisabledSearchFluctuation(true);
				   return;
				 }else{
				 jQuery('#DateArea').show();
				 TextArr=new Array('系统JS 名称','系统JS 描述','系统JS 文件名')
				 }
				 break;
			case '自由 JS' :
				 if (SearchFreeJSPower=='False')       //搜索自由JS权限检查
				 {
				   DisabledSearchFluctuation(true);
				   return;
				 }else{
				 jQuery('#DateArea').show();
				 TextArr=new Array('自由JS 名称','自由JS 描述','自由JS 文件名')
				 }
				 break;
			case '管理员':	 
				  if (SearchAdminPower=='False')          //搜索管理员权限检查
				 {
				  DisabledSearchFluctuation(true);
				  return;
				 }else{
				  DisabledSearchFluctuation(false);
				 jQuery('#DateArea').show();
				 TextArr=new Array('管理员名称','管理员简介')
				}
				break;
		  }
		  for (var i=0;i<TextArr.length;++i){
			jQuery('#SearchType').append("<option value='"+i+"'>"+TextArr[i]+"</option>");
			}
		}
		function setstatus(Obj)
		  {var today=new Date()
			if (Obj.nextSibling.style.display=='none')
			 {
			  Obj.nextSibling.style.display='';
			  jQuery('#StartDate').val(today.getYear()+'-'+(today.getMonth()+1)+'-'+today.getDate());
			 }
			else 
			{
			 Obj.nextSibling.style.display='none';
			 jQuery('#StartDate').val('');
			 }
			if (Obj.nextSibling.nextSibling.style.display=='none')
			{
			 Obj.nextSibling.nextSibling.style.display='';
			  jQuery('#EndDate').val(today.getYear()+'-'+(today.getMonth()+1)+'-'+today.getDate());
			}
			else 
			 {
			 Obj.nextSibling.nextSibling.style.display='none';
			 jQuery('#EndDate').val('');
			 }
		  }
		 function SearchFormSubmit()
		  { var form=document.forms[0];
			if (form.elements[0].value=='')
			 {
			   alert('请输入关键字!')
			   form.elements[0].focus();
			   return false;
			 }
		   switch (form.elements[1].value)
			{
			  case '1':
			  case '2':
			  case '3':
			  case '4':
			  case '5':
			  case '7':
			  case '8':
				   form.action="KS.ItemInfo.asp?ChannelID="+ChannelID;
				   parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape("信息搜索管理 >> <font color=red>搜索结果</font>")+'&ButtonSymbol=Search';
				   break;
			  case 'Special':
				   form.action="KS.Special.asp?Action=SpecialList";
				   parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape("专题管理 >> <font color=red>搜索专题结果</font>")+'&ButtonSymbol=SpecialSearch';
				   break;
			  case 'Link':
				   form.action="KS.FriendLink.asp";
				   parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape("常规管理 >> 友情链接管理 >> <font color=red>搜索友情链接站点结果</font>")+'&ButtonSymbol=LinkSearch';
				   break;
			  case 'SysLabel'  :
				   form.action="Include/Label_Main.asp?LabelType=0";
				   parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape("标签管理 >> <font color=red>搜索系统函数标签结果</font>")+'&ButtonSymbol=SysLabelSearch';
				   break;
			 case 'DIYFunctionLabel'  :
				   form.action="Include/Label_Main.asp?LabelType=5";
				   parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape("标签管理 >> <font color=red>搜索自定义函数标签结果</font>")+'&ButtonSymbol=DIYFunctionSearch';
				   break;
			  case 'FreeLabel'  :
				   form.action="Include/Label_Main.asp?LabelType=1";
				   parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape("标签管理 >> <font color=red>搜索自由标签结果</font>")+'&ButtonSymbol=FreeLabelSearch';
				   break;
			  case 'SysJS'     :
				   form.action="Include/JS_Main.asp?JSType=0";
				   parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape("JS管理 >> <font color=red>搜索系统JS结果</font>")+'&ButtonSymbol=SysJSSearch';
				   break;
			  case 'FreeJS'     :
				   form.action="Include/JS_Main.asp?JSType=1";
				   parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape("JS管理 >> <font color=red>搜索自由JS结果</font>")+'&ButtonSymbol=FreeJSSearch';
				   break;
			  case 'Manager'     :
				   form.action="KS.Admin.asp";
				   parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape("管理员管理 >> <font color=red>搜索管理员结果</font>")+'&ButtonSymbol=ManagerSearch';
				   break;
			}
			form.submit();
		  }
		function DisabledSearchFluctuation(Flag)
		{ if (Flag==true)
		   document.all.KeyWord.value='对不起,权限不足!'; 
		  var AllBtnArray=document.body.getElementsByTagName('INPUT'),CurrObj=null;
			for (var i=0;i<AllBtnArray.length;i++)
			{
				CurrObj=AllBtnArray[i];
				CurrObj.disabled=Flag;
			}
			AllBtnArray=document.body.getElementsByTagName('SELECT'),CurrObj=null;
			for (var i=0;i<AllBtnArray.length;i++)
			{
				CurrObj=AllBtnArray[i];
				CurrObj.disabled=Flag;
			}
		}
		</script>
		<table style="border: 0px solid red" border=0 cellPadding=0 cellSpacing=0>
		  <tr vAlign=top>
			<td valign="top" align=right>
			 <div>
			   <div class="lefttop"></div>
			   <div>
			     <ul id="TabPage">
					<li class="Selected" id="left_tab1" title="内容管理" onClick="javascript:showleft(1);" name="left_tab1">内<br>容</li>
					<li id="left_tab2" <%If Instr(Request.Cookies(KS.SiteSn)("ModelPower"),"sysset1")<=0 and Request.Cookies(KS.SiteSn)("SuperTF")<>"1" then response.Write(" style='display:none' ") %>onClick="javascript:showleft(2);" title="系统管理">设<br>置</li>		
					<li id="left_tab3" <%If Instr(Request.Cookies(KS.SiteSn)("ModelPower"),"subsys1")<=0 and Request.Cookies(KS.SiteSn)("SuperTF")<>"1" then response.Write(" style='display:none' ") %>onClick="javascript:showleft(3);" title="相关操作">相<br>关</li>
					<li id="left_tab4" <%If Instr(Request.Cookies(KS.SiteSn)("ModelPower"),"model1")<=0 and Request.Cookies(KS.SiteSn)("SuperTF")<>"1" then response.Write(" style='display:none' ") %>onClick="javascript:showleft(4);" title="模型管理">模<br>型</li>
					<li id="left_tab5" <%If Instr(Request.Cookies(KS.SiteSn)("ModelPower"),"lab1")<=0 and Request.Cookies(KS.SiteSn)("SuperTF")<>"1" then response.Write(" style='display:none' ") %>onClick="javascript:showleft(5);" title="标签">标<br>签</li>
					<li id="left_tab6" <%If Instr(Request.Cookies(KS.SiteSn)("ModelPower"),"user1")<=0 and Request.Cookies(KS.SiteSn)("SuperTF")<>"1" then response.Write(" style='display:none' ") %>onClick="javascript:showleft(6);" title="用户管理">用<br>户</li>
					<li id="left_tab7" title="插件" <%If Instr(Request.Cookies(KS.SiteSn)("ModelPower"),"other1")<=0 and Request.Cookies(KS.SiteSn)("SuperTF")<>"1" then response.Write(" style='display:none' ") %>onClick="javascript:showleft(7);" name="left_tab7">插<br>
				   件</li>
			     </ul>
			   </div>
			 </div>			
             </td>
			<td align="center" class="boxright">
			 
			    <div>
			      <div class="leftdaohang"></div>	  
				  <div id="menubox">
					<ul id="dleft_tab1">
					 <% dim n:n=0%>
					 
					 <!--------------内容管理 start-------------------->
					 <DIV  class="left_menu" id="td_<%=n+1%>" onClick="javascript:switchShow(<%=n+1%>,1);" height=26>&nbsp;&nbsp;<a href="javascript:void(0)">内容管理</a></DIV>
					  <div class="hid" id="subTable<%=n+1%>" style="WIDTH: 100%">
					     <div class="modellist">
					  <%
					   For I=0 To Ubound(SQL,2)
					      If SQL(6,I)=1 Then 
						   IF instr(KS.C("ModelPower"),sql(5,i) & "0")=0 and SQL(0,I)<>6 and SQL(0,I)<>9 And SQL(0,I)<>10 Then
						   Dim ItemManageUrl
						   Select Case  SQL(4,I)
							Case 1 :ItemManageUrl="KS.Article.asp"
							Case 2 :ItemManageUrl="KS.Picture.asp"
							Case 3 :ItemManageUrl="KS.Down.asp"
							Case 4 :ItemManageUrl="KS.Flash.asp"
							Case 5 :ItemManageUrl="KS.Shop.asp"
							Case 7 :ItemManageUrl="KS.Movie.asp"
							Case 8 :ItemManageUrl="KS.Supply.asp"
						   End Select
						  
						   %>
						   <li>
						   <a href="javascript:void(0)"  onClick="SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red><%=SQL(3,I)%>管理</font>','ViewFolder','KS.ItemInfo.asp?ChannelID=<%=SQL(0,I)%>',<%=SQL(0,I)%>);"><%=KS.Gottopic(SQL(1,I),8)%></a> <span style="cursor:pointer" onClick="SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red>添加<%=SQL(3,I)%></font>','AddInfo','<%=ItemManageUrl%>?Action=Add&ChannelID=<%=SQL(0,I)%>',<%=SQL(0,I)%>);"><img alt="添加<%=SQL(3,I)%>" src="images/add.gif" border="0" align="absmiddle"></span><%if KS.ReturnPowerResult(SQL(0,I), "M"&SQL(0,I)&"10012") then%> <span style="cursor:pointer" onClick="SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red>签收<%=SQL(3,I)%></font>','Disabled','KS.ItemInfo.asp?ComeFrom=Verify&ChannelID=<%=SQL(0,I)%>',<%=SQL(0,I)%>);"><img alt="签收<%=SQL(3,I)%>" src="images/accept.gif" border="0" align="absmiddle"></span>
						   <%end if%>
						   </li>
						   <%
						   End If
						 End If
					   Next
					   %>
					     </div> 
						    <div id='classOpen' style="margin-top:5px;"></div>
						  
                          <div class="modelxg">
						  <script type="text/javascript">
						   var toggle=getCookie("ctips");
						   if (toggle==null) toggle='show';
							$(document).ready(function(){
							TipsToggle(toggle);
							})
						   function TipsToggle(f){
						    setCookie("ctips",f);
							 if (f=='hide'){
							 jQuery("#modelxg").hide('fast');
							 jQuery("#classOpen").html("<img style='cursor:pointer' id='classOpen' onclick='TipsToggle(\"show\")' src='images/left_down.gif' align='absmiddle' title='展开'>");
							 }else{
							 jQuery("#modelxg").show('fast');
							 jQuery("#classOpen").html("<img style='cursor:pointer' id='classOpen' onclick='TipsToggle(\"hide\")' src='images/left_up.gif' title='收藏' align='absmiddle'>");						
                              	 }
						   }
						  </script>
						  
                           <div  id="modelxg">
						   <%If KS.ReturnPowerResult(0, "M010001") Then %>
						   <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'内容管理 >> <font color=red><%=SQL(3,I)%>栏目管理</font>','Disabled','KS.Class.asp');">栏目管理</a> <a href='javascript:void(0)' onClick="SelectObjItem1(this,'栏目管理 >> <font color=red>添加栏目</font>','Go','KS.Class.asp?Action=Add&FolderID=1','');">添加</a></li>
						   <%End If%>
						   <%If KS.ReturnPowerResult(0, "M010002") Then %>
						    <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'内容管理 >> <font color=red>评论管理</font>','Disabled','KS.Comment.asp');">评论管理</a> <a href='javascript:void(0)' onClick="SelectObjItem1(this,'内容管理 >> <font color=red>评论管理</font>','Disabled','KS.Comment.asp?ComeFrom=Verify');">审核</a> </li>
							<%End If%>
							<%If KS.ReturnPowerResult(0, "M010003") Then %>
						    <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'内容管理 >> <font color=red>专题管理</font>','Disabled','KS.Special.asp');">全站专题管理</a> </li>
							<%End If%>
							<%If KS.ReturnPowerResult(0, "M010004") Then %>
						    <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'内容管理 >> <font color=red>关键字Tags管理</font>','Disabled','KS.KeyWord.asp');">关键字Tags管理</a> </li>
							<%End If%>
                            <%If KS.ReturnPowerResult(0, "M010005") or KS.ReturnPowerResult(0, "M010006") Then %>
							<li>
							<%If KS.ReturnPowerResult(0, "M010005") Then%><a href='javascript:void(0)' onClick="SelectObjItem1(this,'内容管理 >> <font color=red>批量设置</font>','Disabled','KS.ItemInfo.asp?Action=SetAttribute');">批量设置</a><%end if%><%If KS.ReturnPowerResult(0, "M010006") then%> <a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'相关操作 >> <font color=red>浏览回收站</font>','ViewFolder','KS.ItemInfo.asp?ComeFrom=RecycleBin','');">回 收 站</a><%end if%></li>
							<%End If%>
						   <%If KS.ReturnPowerResult(0, "M010007") Then %>
						   <li><a href="KS.Tools.asp"  target="MainFrame" title="一键管理工具">一键管理工具</a></li>
						   <%end if%>
						   <%If KS.ReturnPowerResult(0, "M010008") Then %>
						   <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'内容管理 >> <font color=red>信息采集管理</font>','Disabled','Collect/Collect_Main.asp?ChannelID=1');">信息采集管理</a> </li>
						   <%End if%>
						   </div>
						

						   
						 </div>
					 </div>
					<!--------------内容管理 end-------------------->  
					
					
					<!--------------商城管理 start-------------------->
				  <%
				  IF instr(lcase(KS.C("ModelPower")),"shop0")=0 or KS.C("SuperTf")=1 Then
					 If Not ModelXML Is Nothing Then
					  If ModelXML.documentElement.SelectNodes("row[@channelid=5 and @channelstatus=1]").length<>0 Then
					   N=N+1
					 %>
					 <DIV  class="left_menu" id="td_<%=n+1%>" onClick="javascript:switchShow(<%=n+1%>,1);" height=26>&nbsp;&nbsp;<a href="javascript:void(0)">商城管理</a></DIV>
					 <DIV class="hid" id="subTable<%=n+1%>" style="WIDTH: 100%">
					    <%If KS.ReturnPowerResult(5, "M510012") Then %>				
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>处理24小时内订单</font>','Disabled','KS.ShopOrder.asp?searchtype=1&ChannelID=5');"><font color=red>处理24小时内订单</font></a></li>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>处理所有订单</font>','Disabled','KS.ShopOrder.asp?ChannelID=5');">处理所有订单</a></li>
						 <%end if%>
						 <%If KS.ReturnPowerResult(5, "M510014") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>资金明细查询</font>','Disabled','KS.LogMoney.asp?ChannelID=<%=SQL(0,I)%>');">资金明细查询</a></li>
						 <%end if%>
						 <%If KS.ReturnPowerResult(5, "M510015") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>发退货查询</font>','Disabled','KS.LogDeliver.asp?ChannelID=5');">发退货查询</a></li>
						 <%end if%>
						 <%If KS.ReturnPowerResult(5, "M510016") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>发退货查询</font>','Disabled','KS.LogInvoice.asp?ChannelID=5');">开发票查询</a></li>
						 <%end if%>
						 <%If KS.ReturnPowerResult(5, "M510017") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>客户统计</font>','Disabled','KS.ShopStats.asp?Action=Custom');">销售数据统计</a></li>
						 <%End If%>
						 
						 <%If KS.ReturnPowerResult(5, "M520003") Then %>
						 ====================
						  <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>厂商管理</font>','Disabled','KS.Author.asp?ChannelID=5');">厂商管理</a> </li>
						  <%end if%>
						  <%If KS.ReturnPowerResult(5, "M520004") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>送货方式管理</font>','Disabled','KS.Delivery.asp?ChannelID=5');">送货&付款方式</a></li>
						 <%end if%>
						 <%If KS.ReturnPowerResult(5, "M520001") Then %>
						  <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>看货地点管理</font>','Disabled','KS.ShopUnion.asp');">看货地点管理</a></li>
						 <%End If%>
					   	  <%If KS.ReturnPowerResult(5, "M510018") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'品牌管理 >> <font color=red>品牌管理</font>','Disabled','KS.ShopBrand.asp');">品牌管理</a> <a href='javascript:void(0)' onClick="SelectObjItem1(this,'品牌管理 >> <font color=red>添加品牌</font>','Go','KS.ShopBrand.asp?Action=Add&FolderID=0',5);">添加</a> <a href='javascript:void(0)' onClick="SelectObjItem1(this,'品牌管理 >> <font color=red>生成品牌的JS菜单</font>','Go','KS.ShopBrand.asp?Action=Create&FolderID=0',5);">生成</a></li>	
						 <%end if%>	
						 
						 ====================
						 <%If KS.ReturnPowerResult(5, "M520008") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>抢购商品管理</font>','Disabled','KS.Shop.asp?action=LimitBuy&channelid=5');">限时/限量抢购管理</a></li>
						 <%End If%>
						 <%If KS.ReturnPowerResult(5, "M520009") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>限时抢购商品管理</font>','Disabled','KS.Shop.asp?action=BundleSale&channelid=5');">捆绑销售商品管理</a></li>
						 <%End If%>
						 <%If KS.ReturnPowerResult(5, "M520010") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>换购商品管理</font>','Disabled','KS.Shop.asp?action=ChangedBuy&channelid=5');">换购商品管理</a></li>
						 <%End If%>
						 <%If KS.ReturnPowerResult(5, "M520011") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>超值礼包管理</font>','Disabled','KS.Shop.asp?action=Package&channelid=5');">超值礼包管理</a></li>
						 <%end if%>
						 <%If KS.ReturnPowerResult(5, "M510005") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'商城系统 >> <font color=red>批量调价助手</font>','Disabled','KS.ItemInfo.asp?action=SetAttribute&channelid=5');">批量调价助手</a></li>
						 <%End If%>
						 <%If KS.ReturnPowerResult(5, "M520007") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red>优惠券管理</font>','Disabled','KS.ShopCoupon.asp');">优惠券管理</a></li>
						 <%End If%>
						  <%If KS.ReturnPowerResult(5, "M530001") Then %>
						  <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'团购系统 >> <font color=red>团购管理首页</font>','Disabled','KS.GroupBuy.asp');">团购管理首页</a> <a href='javascript:void(0)' onClick="SelectObjItem1(this,'团购系统 >> <font color=red>团购管理首页</font>','Go','KS.GroupBuy.asp?Action=Add');">添加</a></li>
						  <%End If%>
					   <%If KS.ReturnPowerResult(5, "M530002") Then %>
                       <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'团购系统 >> <font color=red>兴趣活动分类管理</font>','Disabled','KS.GroupBuyInt.asp');">兴趣活动分类管理</a></li>
					   <%end if%>
					 
					 </DIV>
					 <!--------------商城管理 End-------------------->
					<% End If
					End If
				  End If
					%>
					
					   
					
					<!--------------音乐管理 start-------------------->
					 <%
					IF instr(lcase(KS.C("ModelPower")),"music0")=0 or KS.C("SuperTf")=1 Then
					 If Not ModelXML Is Nothing Then
					  If ModelXML.documentElement.SelectNodes("row[@channelid=6 and @channelstatus=1]").length<>0 Then
					   N=N+1
					  %>
					 <DIV  class="left_menu" id="td_<%=n+1%>" onClick="javascript:switchShow(<%=n+1%>,1);" height=26>&nbsp;&nbsp;<a href="javascript:void(0)">音乐管理</a></DIV>
					  <DIV class="hid" id="subTable<%=n+1%>" style="WIDTH: 100%">
					  <li><a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'音乐系统 >> <font color=red>所有歌曲管理</font>','Disabled','KS.Music.asp?url=KS.MusicSong.asp');">歌曲管理</a> <a href='javascript:void(0)' onClick="SelectObjItem1(this,'音乐系统 >> <font color=red>添加歌曲</font>','Disabled','KS.Music.asp?url=KS.MusicSong.asp?Action=Add&Classid=1',6);">添加歌曲</a></li>
						 
						 <li><a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'音乐系统 >> <font color=red>所有歌曲管理</font>','Disabled','KS.Music.asp?url=KS.MusicSpecial.asp');">专辑管理</a> <a href='javascript:void(0)' onClick="SelectObjItem1(this,'音乐系统 >> <font color=red>添加歌曲</font>','Disabled','KS.Music.asp?url=KS.MusicSpecial.asp?Action=Step1',6);">添加专辑</a></li>
						 <li><a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'音乐系统 >> <font color=red>分类管理</font>','Disabled','KS.Music.asp?url=KS.MusicType.asp');">分类管理</a> <a href='javascript:void(0)' onClick="SelectObjItem1(this,'音乐系统 >> <font color=red>歌手管理</font>','Disabled','KS.Music.asp?url=KS.MusicSinger.asp',6);">歌手管理</a></li>
						 <li><a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'音乐系统 >> <font color=red>审核歌词</font>','Disabled','KS.Music.asp?url=KS.MusicGeCi.asp');">审核歌词</a> <a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'音乐系统 >> <font color=red>专辑评论管理</font>','Disabled','KS.Music.asp?url=KS.MusicComment.asp');">专辑评论</a></li>
						 <li><a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'音乐系统 >> <font color=red>歌曲服务器管理</font>','Disabled','KS.MediaServer.asp?TypeID=1');">歌曲服务器管理</a></li>
					  </DIV>
				    <!--------------音乐管理 end--------------------> 
					<%End If
					End If
				  End If
					%>
					
					<!--------------招聘求职 start-------------------->
					 <%
				   IF instr(lcase(KS.C("ModelPower")),"job0")=0 or KS.C("SuperTf")=1 Then
					 If Not ModelXML Is Nothing Then
					  If ModelXML.documentElement.SelectNodes("row[@channelid=10 and @channelstatus=1]").length<>0 Then
					   N=N+1
					  %>
					 <DIV  class="left_menu" id="td_<%=n+1%>" onClick="javascript:switchShow(<%=n+1%>,1);" height=26>&nbsp;&nbsp;<a href="javascript:void(0)">招聘求职</a></DIV>
					 <DIV class="hid" id="subTable<%=n+1%>" style="WIDTH: 100%">
					  <%
					     Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & server.urlencode("求职招聘 >> <font color=red>招聘系统设置</font>'") & ",'SetParam','KS.JobSetting.asp');"">招聘系统设置</a></li>"
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & server.urlencode("求职招聘 >> <font color=red>行业职位设置</font>'") & ",'disabled','KS.Jobhy.asp');"">行业职位设置</a></li>"
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & server.urlencode("求职招聘 >> <font color=red>简历模板管理</font>'") & ",'disabled','KS.JobTemplate.asp');"">简历模板管理</a></li>"
						  Response.Write "&nbsp;==================="
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & server.urlencode("求职招聘 >> <font color=red>招聘单位管理</font>'") & ",'disabled','KS.JobCompany.asp');"">招聘单位管理</a></li>"
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & Server.UrlEncode("求职招聘 >> <font color=red>审核招聘单位</font>") & "','disabled','KS.JobCompany.asp?ComeFrom=Verify');"">招聘单位审核</a></li>"
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & server.urlencode("求职招聘 >> <font color=red>添加招聘单位</font>'") & ",'disabled','KS.JobCompany.asp?Action=Add');"">添加招聘单位</a></li>"
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & server.urlencode("求职招聘 >> <font color=red>招聘职位管理</font>'") & ",'disabled','KS.Jobzw.asp');"">招聘职位管理</a></li>"
						  Response.Write "&nbsp;==================="
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & server.urlencode("求职招聘 >> <font color=red>个人简历管理</font>'") & ",'disabled','KS.JobResume.asp');"">个人简历管理</a></li>"
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & server.urlencode("求职招聘 >> <font color=red>审核个人简历</font>'") & ",'disabled','KS.JobResume.asp?ComeFrom=Verify');"">个人简历审核</a></li>"
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & Server.Urlencode("求职招聘 >> <font color=red>添加个人简历</font>'") & ",'disabled','KS.JobResume.asp?Action=Add');"">添加个人简历</a></li>"
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & Server.Urlencode("求职招聘 >> <font color=red>教育背景管理</font>'") & ",'disabled','KS.JobEdu.asp');"">教育背景管理</a></li>"
					  %>
					 
					 </DIV>
                    <!--------------招聘求职 end--------------------> 
					<% End If
					End If
				   End If
					%>
					
					
					<!--------------考试系统 start-------------------->
				 <%
				IF instr(lcase(KS.C("ModelPower")),"mnkc0")=0 or KS.C("SuperTf")=1 Then
					 If Not ModelXML Is Nothing Then
					  If ModelXML.documentElement.SelectNodes("row[@channelid=9 and @channelstatus=1]").length<>0 Then
					   N=N+1
					   %>
					 <DIV  class="left_menu" id="td_<%=n+1%>" onClick="javascript:switchShow(<%=n+1%>,1);" height=26>&nbsp;&nbsp;<a href="javascript:void(0)">考试系统</a></DIV>
					 <DIV class="hid" id="subTable<%=n+1%>" style="WIDTH: 100%">
					 <%
					      Response.Write "<li><a href='mnkc/mnkc.asp' target='MainFrame'>试卷管理/添加</a></li>"
					      Response.Write "<li><a href='mnkc/mnkc_score.asp' target='MainFrame'>考试成绩管理</a></li>"
						  Response.Write "==================="
						  Response.Write "<li><a href='mnkc/refreshindex.asp' target='MainFrame'>发布频道首页</a></li>"
						  Response.Write "<li><a href='mnkc/mnkc_makesortall.asp?type=all' target='MainFrame'>发布所有分类</a></li>"
						  Response.Write "<li><a href='mnkc/mnkc_makemnkcall.asp' target='MainFrame'>发布所有试卷页</a></li>"
						  Response.Write "<li><a href='mnkc/RefreshClass.asp' target='MainFrame'>发布总分类页</a></li>"
					 %>
					 </DIV>
                    <!--------------考试系统 end--------------------> 
				    <%
					 End If
					End If
			   End If
					%>
					
					<!--------------问答系统 start-------------------->
					<%IF instr(lcase(KS.C("ModelPower")),"ask0")=0 or KS.C("SuperTf")=1 Then%>
					 <%N=N+1%>
					 <DIV  class="left_menu" id="td_<%=n+1%>" onClick="javascript:switchShow(<%=n+1%>,1);" height=26>&nbsp;&nbsp;<a href="javascript:void(0)">问答系统</a></DIV>
					 <DIV class="hid" id="subTable<%=n+1%>" style="WIDTH: 100%">
					 <%If KS.ReturnPowerResult(0, "WDXT10000") Then%>
					   <li><a href="javascript:void(0)" onClick="SelectObjItem1(this,'问答系统 >> <font color=red>问答参数设置</font>','SetParam','KS.AskSetting.asp');" title="问答参数设置">问答参数设置</a></li>
					   <%end if%>
					   <li><a href="javascript:void(0)" onClick="SelectObjItem1(this,'问答系统 >> <font color=red>问题列表管理</font>','SetParam','KS.AskList.asp');" title="问题列表管理">问题列表管理</a></li>
					   <%If KS.ReturnPowerResult(0, "WDXT10002") Then%>
					   <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'问答系统 >> <font color=red>分类管理</font>','Disabled','KS.AskClass.asp');">分类管理</a>
					   <a href='javascript:void(0)' onClick="SelectObjItem1(this,'问答系统 >> <font color=red>添加问答分类</font>','GO','KS.AskClass.asp?action=add');">添加</a></li>
					   <%end if%>
					   <%If KS.ReturnPowerResult(0, "WDXT10003") Then%>
					   <li><a href="javascript:void(0)" onClick="SelectObjItem1(this,'问答系统 >> <font color=red>等级头衔设置</font>','Disabled','KS.AskGrade.asp');" title="等级头衔设置">等级头衔设置</a></li>
					   <%end if%>
					   </li>
					 </DIV>
				   <%End If%>
                    <!--------------问答系统 end--------------------> 
					
					<!--------------空间系统 start-------------------->
				   <%IF instr(lcase(KS.C("ModelPower")),"space0")=0 or KS.C("SuperTf")=1 Then%>
					 <%N=N+1%>
					 <DIV  class="left_menu" id="td_<%=n+1%>" onClick="javascript:switchShow(<%=n+1%>,1);" height=26>&nbsp;&nbsp;<a href="javascript:void(0)">空间门户</a></DIV>
					 <DIV class="hid" id="subTable<%=n+1%>" style="WIDTH: 100%">
					    	 <div style=" border:#ff6600 1px dotted;width:125px; height:21px; line-height:21px;margin-right:6px;margin-bottom:2px; margin-top:2px;text-align:left;padding-left:5px; font-size:14px;font-weight:bold; color:#ff6600;"><img src="images/ico_friend.gif">&nbsp;个人空间</div>
						 <%If cbool(KS.ReturnPowerResult(0, "KSMS10000")) Then%>
						<li><a href="javascript:void(0)" onClick="SelectObjItem1(this,'个人空间 >> <font color=red>空间参数设置</font>','SetParam','KS.SpaceSetting.asp');" title="空间参数设置">空间参数设置</a></li>
						<%end if%>
						<%If KS.ReturnPowerResult(0, "KSMS10001") Then%>
						<li><a href="KS.Space.asp"  target="MainFrame" title="所有空间管理">所有空间管理</a></li>
						<li><a href="KS.Space.asp?showtype=1"  target="MainFrame" title="个人空间管理">个人空间管理</a></li>
						<%end if%>
						<%If KS.ReturnPowerResult(0, "KSMS10002") Then%>
						<li><a href="KS.Spacelog.asp"  target="MainFrame" title="空间日志管理">空间日志管理</a></li>
						<%end if%>
						<%If KS.ReturnPowerResult(0, "KSMS10003") Then%>
						<li><a href="KS.SpaceAlbum.asp"  target="MainFrame" title="用户相册管理">用户相册管理</a></li>
						<%end if%>
						<%If KS.ReturnPowerResult(0, "KSMS10004") Then%>
						<li><a href="KS.SpaceTeam.asp"  target="MainFrame" title="用户圈子管理">用户圈子管理</a></li>
						<%end if%>
						<%If KS.ReturnPowerResult(0, "KSMS10005") Then%>
						<li><a href="KS.SpaceMessage.asp"  target="MainFrame" title="用户留言管理">用户留言管理</a></li>
						<%end if%>
						<%If KS.ReturnPowerResult(0, "KSMS10007") Then%>
						<li><a href="KS.SpaceMusic.asp"  target="MainFrame" title="用户歌曲管理">用户歌曲管理</a></li>
						<%end if%>
						 <div style=" border:#ff6600 1px dotted;width:125px; height:21px; line-height:21px;margin-left:5px; text-align:left;padding-left:5px; font-size:14px;font-weight:bold; color:#ff6600;"><img src="images/ico_home.gif">&nbsp;企业空间</div>
						<%If KS.ReturnPowerResult(0, "KSMS10008") Then%>
					  <li><a href="KS.EnterPrise.asp"  target="MainFrame" title="企业信息管理">企业空间管理</a></li>
					 <%end if%>
					 <%If KS.ReturnPowerResult(0, "KSMS10009") Then%>
					  <li><a href="KS.EnterPriseNews.asp"  target="MainFrame" title="企业新闻管理">企业新闻管理</a></li>
					 <%end if%>
					 <%If KS.ReturnPowerResult(0, "KSMS10010") Then%>
					  <li><a href="KS.EnterPrisePro.asp"  target="MainFrame" title="企业产品管理">企业产品管理</a></li>
					 <%end if%>
					 <%If KS.ReturnPowerResult(0, "KSMS10012") Then%>
					  <li><a href="KS.EnterPriseClass.asp"  target="MainFrame" title="行业分类管理">行业分类管理</a></li>
					 <%end if%>
					 <%If KS.ReturnPowerResult(0, "KSMS10013") Then%>
					  <li><a href="KS.EnterPriseAD.asp"  target="MainFrame" title="行业广告管理">行业广告管理</a></li>
					 <%end if%>
					 <%If KS.ReturnPowerResult(0, "KSMS10011") Then%>
					  <li><a href="KS.EnterPriseZS.asp"  target="MainFrame" title="荣誉证书管理">荣誉证书管理</a></li>
					 <%end if%>
						
					 </DIV>
					<%End If%>
                    <!--------------空间系统 end--------------------> 

					
						 
					
					</ul>
					
					
					
					
					<ul id="dleft_tab2" style="display:none;">
					   <div class="dt">系统设置</div>
					   <div class="dc">
					<%If KS.ReturnPowerResult(0, "KMST10001") Then%>
					   <li><a href="javascript:void(0)" onClick="SelectObjItem1(this,'系统设置 >> <font color=red>基本信息设置</font>','SetParam','KS.System.asp');" title="基本信息设置">基本信息设置</a></li>
					 <%end if%>
					      
						 <%If Not ModelXML Is Nothing Then
					       If ModelXML.documentElement.SelectNodes("row[@basictype=3 and @channelstatus=1]").length<>0 Then
						 %>
						  <%If KS.ReturnPowerResult(0, "KMST20001") Then %>
						  <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red><%=SQL(3,I)%>参数设置</font>','SetParam','KS.DownParam.asp?ChannelID=<%=SQL(0,I)%>');">下载参数设置</a></li>
						  <%End If%>
						 
						 <%If KS.ReturnPowerResult(0, "KMST20002") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red>下载服务器管理</font>','Disabled','KS.DownServer.asp?ChannelID=<%=SQL(0,I)%>');">下载服务器管理</a>
						 <%end if%>
						 
					
						<%
						  End If
						End If
						
						If Not ModelXML Is Nothing Then
					       If ModelXML.documentElement.SelectNodes("row[@channelid=7 and @channelstatus=1]").length<>0 Then
						 %>
						  <%If KS.ReturnPowerResult(0, "KMST20003") Then %>
						 <li><a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red>参数设置</font>','SetParam','KS.MovieParam.asp?ChannelID=7');">影视参数设置</a></li>
						  <%End If%>
						
						 <%If KS.ReturnPowerResult(0, "KMST20004") Then %>
						 <li><a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red>影视服务器管理</font>','Disabled','KS.MediaServer.asp?TypeID=2&ChannelID=7');">影视服务器管理</a></li>
						 <%end if%>
					    <%
						   End If
						End If
						
						If Not ModelXML Is Nothing Then
					       If ModelXML.documentElement.SelectNodes("row[@channelid=8 and @channelstatus=1]").length<>0 Then
						 %>
						 <%If KS.ReturnPowerResult(0, "KMST20005") Then %>
						 <li><a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red>供求交易类型管理</font>','Disabled','KS.SupplyType.asp');">供求交易类型管理</a></li>
						  <%End If%>
					  <%  End If
					   End If
					   %>
					 
					 <%If KS.ReturnPowerResult(0, "KMST10003") Then%>
					   <li><a href="KS.PaymentPlat.asp"  target="MainFrame" title="在线支付平台管理">在线支付平台管理</a></li>
					 <%end if%>
					 <%If KS.ReturnPowerResult(0, "KMST10002") Then%>
					   <li><a href="javascript:void(0)" onClick="SelectObjItem1(this,'系统设置 >> <font color=red>整合系统设置</font>','SetParam','KS.API.asp');"  title="整合系统设置">API通用整合设置</a></li>
					 <%end if%>
					   </div>
					   
					   
					<%If KS.ReturnPowerResult(0, "KSO10000") Then %>
					  <div class="dt">WAP系统管理</div>
					  <div class="dc">
                       <li><a href="#" onClick="SelectObjItem1(this,'WAP系统管理 >> <font color=red>WAP基本参数设置</font>','SetParam','Wap/KS_System.asp');" title="WAP基本参数设置">WAP基本参数设置</a></li>
					   <li><a href="#"  onClick="SelectObjItem1(this,'WAP系统管理 >> <font color=red>WAP自定义页面管理</font>','Disabled','Wap/KS.Template.asp');">WAP自定义页面</a></li>
					  </div>
					<%end if%>
					   
					   
					   <div class="dt">辅助管理</div>
					   <div class="dc">
						 <%If KS.ReturnPowerResult(0, "KMST10015") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'辅助管理 >> <font color=red>来源管理</font>','Disabled','KS.Origin.asp');">来源管理</a></li>
						 <%end if%>
						 <%If KS.ReturnPowerResult(0, "KMST10016") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'辅助管理 >> <font color=red>作者管理</font>','Disabled','KS.Author.asp?ChannelID=0');">作者管理</a> </li>
						 <%end if%>

						 <%If KS.ReturnPowerResult(0, "KMST10017") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'辅助管理 >> <font color=red>省市管理</font>','Disabled','KS.Province.asp');">地区管理</a> </li>
						 <%end if%>

					  <%If KS.ReturnPowerResult(0, "KMST10004") Then%>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'辅助管理 >> <font color=red>内容关键字设置</font>','Disabled','KS.InnerLink.asp');">内容关键字设置</a></li>
                      <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMST10019") Then%>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'辅助管理 >> <font color=red>搜索关键词维护</font>','Disabled','KS.KeyWord.asp?issearch=1');">搜索关键词维护</a></li>
                      <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMST10020") Then%>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'辅助管理 >> <font color=red>定时任务管理</font>','Disabled','KS.Task.asp?action=manage');">定时任务管理</a></li>
                      <%end if%>
					  
                       <%If KS.ReturnPowerResult(0, "KMST10014") Then %>
					     <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red>投票记录管理</font>','Disabled','KS.PhotoVote.asp?ChannelID=<%=SQL(0,I)%>');">图片投票记录管理</a>	</li>
					   <%End If%>

					  <%If KS.ReturnPowerResult(0, "KMST10006") Then%>
					   	<li><a href="KS.Log.asp"  target="MainFrame" title="站点文件管理">后台日志管理</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMST10007") Then%>
					   <li><a href="KS.Database.asp?Action=BackUp"  target="MainFrame" title="数据库维护">数据库维护</a></li>
					   <%end if%>
					   <%If KS.ReturnPowerResult(0, "KMST10008") Then%>
					   <li><a href="KS.DataReplace.asp"  target="MainFrame" title="数据库字段替换">数据库字段替换</a></li>
					   <%end if%>
                       <%If KS.ReturnPowerResult(0, "KMST10018") Then%>
					   <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'辅助管理 >> <font color=red>上传文件管理</font>','Disabled','KS.AdminFiles.asp');">上传文件管理</a></li>
					   <%end if%>					   
					   <%If KS.ReturnPowerResult(0, "KMST10009") Then%>
					   <li><a href="KS.Database.asp?Action=ExecSql"  target="MainFrame" title="在线执行SQL语句">在线执行SQL语句</a></li>
					   <%end if%>
					   <%If KS.ReturnPowerResult(0, "KMST10011") Then%>
					   <li><a href="KS.System.asp?Action=CopyRight"  target="MainFrame" title="服务器参数探测">服务器参数探测</a></li>
					   <%end if%>
					   <%If KS.ReturnPowerResult(0, "KMST10012") Then%>
					   <li><a href="KS.CheckMM.asp"  target="MainFrame" title="在线检测木马">在线检测木马</a></li>
					   <%end if%>
					   </div>
					   
					</ul>
					
					
					
					
					
					<ul id="dleft_tab3" style="display:none;">
					<%If KS.ReturnPowerResult(0, "KSMS10006") Then %>
					<div class="dt">自定义表单</div>
					  <div class="dc">
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'自定义表单 >> <font color=red>表单项目管理</font>','Disabled','KS.Form.asp');">自定义表单管理</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'自定义表单 >> <font color=red>添加表单项目</font>','GO','KS.Form.asp?action=Add');">添加表单项目</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'自定义表单 >> <font color=red>表单项目调用代码</font>','Disabled','KS.Form.asp?action=total');">表单项目调用代码</a></li>
					  </div>
					<%end if%>
					<%If KS.ReturnPowerResult(0, "KSMS20004") Then%>
					<div class="dt">
					小论坛/留言
					</div>
					<div class="dc">
					<li><a href="KS.GuestBook.asp?Action=Main"  target="MainFrame" title="网站留言管理">网站留言管理</a></li>
					<li><a href="KS.GuestBoard.asp"  target="MainFrame" title="版面分类管理">版面分类管理</a></li>
					</div>
					<%end if%>
					<%If KS.ReturnPowerResult(0, "KSMS20014") Then%>
					<div class="dt">PK系统</div>
					  <div class="dc">
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'观点PK系统 >> <font color=red>PK主题管理</font>','Disabled','KS.PKZT.asp');">PK主题管理</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'观点PK系统 >> <font color=red>PK用户观点管理</font>','Disabled','KS.PKGD.asp');">PK用户观点管理</a></li>
					  </div>
					<%end if%>
					
                    <div class="dt">
					其它系统
					</div>
					<div class="dc">
						 <%If KS.ReturnPowerResult(0, "KSMS20010") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'其它系统 >> <font color=red>积分兑换商品</font>','Disabled','KS.MallScore.asp');">积分兑换商品</a></li>
						 <%End If%>
					<%If KS.ReturnPowerResult(0, "KSMS20009") Then %>
					<li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'其它系统 >> <font color=red>Digg管理</font>','Disabled','KS.DiggList.asp');">文档Digg管理</a></li>
					<%End If%>
					<%If KS.ReturnPowerResult(0, "KSMS20008") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'其它系统 >> <font color=red>心情指数管理</font>','Disabled','KS.Mood.asp');">心情指数管理</a></li>
					<%end if%>
					<%If KS.ReturnPowerResult(0, "KSMS20001") Then%>
					<li><a href="KS.FriendLink.asp"  target="MainFrame" title="友情链接管理">友情链接管理</a></li>
					<%end if%>
					<%If KS.ReturnPowerResult(0, "KSMS20002") Then%>
					<li><a href="KS.Announce.asp"  target="MainFrame" title="网站公告管理">网站公告管理</a></li>
					<%end if%>
					<%If KS.ReturnPowerResult(0, "KSMS20000") Then%>
					<li><a href="KS.FeedBack.asp"  target="MainFrame" title="投诉及反馈管理">投诉及反馈管理</a></li>
					<%end if%>
					<%If KS.ReturnPowerResult(0, "KSMS20003") Then%>
					<li><a href="KS.Vote.asp"  target="MainFrame" title="站内调查管理">站内调查管理</a></li>
					<%end if%>
					
					<%If KS.ReturnPowerResult(0, "KSMS20005") Then%>
					<li><a href="KS.Online.asp"  target="MainFrame" title="站点计数器管理">站点计 数 器</a></li>
					<%end if%>
					<%If KS.ReturnPowerResult(0, "KSMS20006") Then%>
					<li><a href="KS.Ads.asp"  target="MainFrame" title="广告系统管理">广告系统管理</a></li>
					<%end if%>
					<%If KS.ReturnPowerResult(0, "KSMS20007") Then%>
					<li><a href="KS.PromotedPlan.asp"  target="MainFrame" title="推广计划管理">推广计划管理</a></li>
					<%end if%>
					</div>
					<div class="dt">版权信息</div>
                       <div class="dc">
					    <li><a href="javascript:void(0)">开发:科兴信息技术有限公司</a></li>
						<li><a href="http://www.kesion.com" target="_blank">官方:kesion.com</a></lI>
						<li><a href="javascript:void(0)">电话:0596-2218051<br />0596-2198252</a></lI>
						<li><a href="javascript:void(0)">咨询QQ:9537636 41904294</a></lI>
					   </div>
					
					
					</ul>
					
					
					
										
					<ul id="dleft_tab4" style="display:none">
					<div class="dt">模型管理</div>
					 <div class="dc">
					 <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'模型管理 >> <font color=red>模型管理首页</font>','Disabled','KS.Model.asp');">模型管理首页</a></li>
					 <%If KS.ReturnPowerResult(0, "KSMM10000") Then%>
					 <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'模型管理 >> <font color=red>添加新模型</font>','Go','KS.Model.asp?action=Add');">添加新模型</a></li>
					 <%end if%>
					 <%If KS.ReturnPowerResult(0, "KSMM10004") Then%>
					 <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'模型管理 >> <font color=red>模型信息统计</font>','Go','KS.Model.asp?action=total');">模型信息统计</a></li>
					 <%end if%>
					 </div>
					 <%If KS.ReturnPowerResult(0, "KSMM10003") Then%>
					<div class="dt">模型字段管理</div>
					 <div class="dc">
					  <%For I=0 To UBound(SQL,2)
					   if SQL(6,I)=1 AND SQL(0,I)<>6 and SQL(0,I)<>9 and SQL(0,I)<>10 Then
					  %>
						 <li><a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'模型管理 >> <font color=red>字段管理</font>','Disabled','KS.Field.asp?ChannelID=<%=SQL(0,I)%>',<%=SQL(0,I)%>);"><%=SQL(1,I)%>字段</a></li>					  
					<%
					  End iF
					 Next%>
					</div>
					<%end if%>
					</ul>

                    <ul id="dleft_tab5" style="display:none">
					 <div class="dt">标签管理</div>
					 <div class="dc">
					<%
					If KS.ReturnPowerResult(0, "KMTL10001") Then
					  .Write ("<li><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'标签管理 >> <font color=red>系统函数标签</font>','FunctionLabel','Include/Label_Main.asp?LabelType=0');"">系统函数标签</a></li>")
					End If
					If KS.ReturnPowerResult(0, "KMTL10002") Then
					  .Write ("<li><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'标签管理 >> <font color=red>自定义SQL函数标签</font>','DiyFunctionLabel','Include/Label_Main.asp?LabelType=5');"">自定义SQL函数标签</a></li>")
					End If
					If KS.ReturnPowerResult(0, "KMTL10003") Then
					  .Write ("<li><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'标签管理 >> <font color=red>自定义静态标签</font>','FreeLabel','Include/Label_Main.asp?LabelType=1');"">自定义静态标签</a></li>")
					End If
					If KS.ReturnPowerResult(0, "KMTL10010") Then
					  .Write ("<li><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'标签管理 >> <font color=red>通用循环标签</font>','FreeLabel','Include/Label_Main.asp?LabelType=6');"">通用循环列表标签</a></li>")
					End If
					If KS.ReturnPowerResult(0, "KMTL10004") Then
					  .Write ("<li><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'标签管理 >> <font color=red>自定义JS管理</font>','SysJSList','include/JS_Main.asp?JSType=0');"">系统JS管理</a></li>")
					End If
					If KS.ReturnPowerResult(0, "KMTL10005") Then
					  .Write ("<li><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'标签管理 >> <font color=red>自定义JS管理</font>','FreeJSList','include/JS_Main.asp?JSType=1');"">自定义JS管理</a></li>")
					End If
					If KS.ReturnPowerResult(0, "KMSL10008") Then
					  .Write "<li><a href='KS.ClassMenu.asp'  target='MainFrame' title='生成顶部菜单'>生成顶部菜单</a></li>"
					end if
					If KS.ReturnPowerResult(0, "KMSL10009") Then
					  .Write "<li><a href='KS.TreeMenu.asp'  target='MainFrame' title='生成树形菜单'>生成树形菜单</a></li>"
					End If

		              .write "</div>"
					  .write "<div class='dt'>模板管理</div>"
					  .write "<div class='dc'>"
					If KS.ReturnPowerResult(0, "KMTL10006") Then
						.Write ("<li id='s_1'><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'模板标签管理 >> <font color=red>自定义页面管理</font>','Disabled','KS.DIYPage.asp');"">自定义页面管理</a></li>")
				    End If
					If KS.ReturnPowerResult(0, "KMTL10007") Then
						.Write ("<li id='s_1'><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'模板标签管理 >> <font color=red>所有模板管理</font>','Disabled','KS.Template.asp');"">所有模板管理</a></li>")
					End If
					 %>	
					 </div>
					</ul>
					
					<ul id="dleft_tab6" style="display:none">
					
					  <div class="dt">
					   用户管理					  </div>
					  <div class="dc">
					  <%If KS.ReturnPowerResult(0, "KMUA10001") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>管理员管理</font>','Disabled','KS.Admin.asp');">管理员管理</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10002") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>注册用户管理</font>','Disabled','KS.User.asp');" title="注册用户管理">注册用户管理</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>添加用户</font>','Disabled','KS.User.asp?Action=Add');" title="添加用户">添加用户</a></li>
					  
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10004") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>用户组管理</font>','Disabled','KS.UserGroup.asp');" title="用户组管理">用户组管理</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10003") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>用户短信管理</font>','Disabled','KS.UserMessage.asp');" title="用户短信管理">用户短信管理</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10009") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>发送邮件管理</font>','Disabled','KS.UserMail.asp');" title="发送邮件管理">发送邮件管理</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10012") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>会员字段管理</font>','Disabled','KS.Field.asp?ChannelID=101');" title="会员字段管理">会员字段管理</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10013") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>会员表单管理</font>','Disabled','KS.UserForm.asp');" title="会员表单管理">会员表单管理</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10014") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>会员动态管理</font>','Disabled','KS.UserLog.asp');" title="会员动态管理">会员动态管理</a></li>
					  <%end if%>
					  
					  
					  </div>
					  <div class="dt">
					   账务明细管理					 
					  </div>
					  <div class="dc">
					  <%If KS.ReturnPowerResult(0, "KMUA10005") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>会员点券明细</font>','Disabled','KS.LogPoint.asp');" title="会员点券明细">会员点券明细</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10006") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>会员有效期明细</font>','Disabled','KS.LogEdays.asp');" title="会员有效期明细">会员有效期明细</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10007") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>会员资金明细</font>','Disabled','KS.LogMoney.asp');" title="会员资金明细">会员资金明细</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>会员积分明细</font>','Disabled','KS.LogScore.asp');" title="会员积分明细">会员积分明细</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10008") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>线下充值卡管理</font>','Disabled','KS.Card.asp?cardtype=0');" title="线下充值卡管理">线下充值卡管理</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>线上充值卡管理</font>','Disabled','KS.Card.asp?cardtype=1');" title="线上充值卡管理">线上充值卡管理</a></li>
					  <%end if%>
					  </div>
					  <%If KS.ReturnPowerResult(0, "KMUA10002") Then %>
					  <div class="dt">
					   快速查找用户					  </div>
					  <div class="dc">
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>24小时内登录</font>','Disabled','KS.User.asp?UserSearch=5');"><font color=#ff6600>24小时内登录</a></font></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>24小时内登录</font>','Disabled','KS.User.asp?UserSearch=6');">24小时内注册</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>24小时内登录</font>','Disabled','KS.User.asp?UserSearch=1');"> 被锁住的用户</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>24小时内登录</font>','Disabled','KS.User.asp?UserSearch=3');">待审批会员</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>24小时内登录</font>','Disabled','KS.User.asp?UserSearch=4');">待邮件验证</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'用户系统 >> <font color=red>24小时内登录</font>','Disabled','KS.User.asp?UserSearch=2');">所有管理员用户</a></li>
                      </div>
					<%end if%>
					</ul>
					<ul id="dleft_tab7" style="display:none">
					<%If KS.ReturnPowerResult(0, "KSO10002") Then %>
					  <div class="dt">CC视频插件</div>
					  <div class="dc">
					   <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'CC视频插件 >> <font color=red>参数设置</font>','Disabled','../plus/CC/cc.asp');">CC参数设置</a></li>
					  </div>
					<%end if%>
					<%If KS.ReturnPowerResult(0, "KSO10003") Then %>
					  <div class="dt">WSS统计插件</div>
					  <div class="dc">
					   <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'WSS 统计插件 >> <font color=red>WSS 设置</font>','Disabled','../plus/wss/wss.asp');">WSS 设置</a></li>
					   <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'WSS 统计插件 >> <font color=red>WSS 设置</font>','Disabled','../plus/wss/wss.asp?action=show');">查看统计</a></li>
					  </div>
					<%end if%>
					
					  <div class="dt">数据导入插件</div>
					  <div class="dc">
					   <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'数据导入插件 >> <font color=red>数据批量导入管理</font>','Disabled','KS.Import.asp');">数据批量导入管理</a></li>
					  </div>
					
				    </ul>
				  </div>
					
					<div></div>
			</div><!--menubox-->			</td>
		  </tr>
		</table>
		<SCRIPT type="text/javascript">
		function fHideFocus(tName){
		aTag=document.getElementsByTagName(tName);
		for(i=0;i<aTag.length;i++)aTag[i].onfocus=function(){this.blur();};
		}
		fHideFocus("A");
		var id = 1;  //默认选中的ID
		document.getElementById("subTable"+id).className = "show";
		document.getElementById("td_"+id).className = "left_menu_selected";
		var cache_id = id;
		function switchShow(id,tag){
		    document.getElementById("td_"+id).className='left_menu_selected';
			for(var i=1; i<=<%=n%>; i++){
			   if (i!=id)
				document.getElementById("td_"+i).className='left_menu';
		     }
			var tObj = document.getElementById("subTable"+id);
			var	cObj = document.getElementById("subTable"+cache_id);
			if(tag){
				if(tObj) tObj.className =(tObj.className=='hid') ? "show" : "hid";
			}else{
				if(tObj) tObj.className = "show";
			}
			if(cache_id != id){
				cache_id = id;
				if(cObj)cObj.className = "hid";
			}
			event.cancelBubble = true;
		}
		function showleft(id)
		{ 
		 document.getElementById("left_tab"+id).className='Selected';
		 var oItem = document.getElementById("TabPage").getElementsByTagName("li"); 
			for(var i=1; i<=oItem.length; i++){
			   if (i!=id)
				document.getElementById("left_tab"+i).className='';
		     }
			var dvs=document.getElementById("menubox").getElementsByTagName("ul");
			for (var i=0;i<dvs.length;i++){
			  if (dvs[i].id==('dleft_tab'+id))
				dvs[i].style.display='';
			  else
			  dvs[i].style.display='none';
			}
		}
		</SCRIPT>
		</body>
<%
        If Session("ShowCount")="" Then
		.Write " <ifr" & "ame src=""http://ww" &"w.k" &"e" & "si" &"on." & "co" & "m" & "/WebS" & "ystem/Co" & "unt.asp"" scrolling='no' frameborder='0' height='0' width='0'></iframe>"
		Session("ShowCount")=KS.C("AdminName")
		End If
		.Write "</html>"
	    End With
		End Sub
		Function bytes2BSTR(vIn)
		Dim i,ThisCharCode,NextCharCode
		Dim strReturn:strReturn = ""
		For i = 1 To LenB(vIn)
			ThisCharCode = AscB(MidB(vIn,i,1))
			If ThisCharCode < &H80 Then
				strReturn = strReturn & Chr(ThisCharCode)
			Else
				NextCharCode = AscB(MidB(vIn,i+1,1))
				strReturn = strReturn & Chr(CLng(ThisCharCode) * &H100 + CInt(NextCharCode))
				i = i + 1
			End If
		Next
		bytes2BSTR = strReturn
		End Function
		Function getfile(RemoteFileUrl)
		On Error Resume Next 
		Dim Retrieval:Set Retrieval = Server.CreateObject("Microsoft.XMLHTTP")
		With Retrieval
		.Open "Get", RemoteFileUrl, false, "", ""
		.Send
		If .Readystate<>4 then
				Exit Function
		End If
		 getfile =bytes2BSTR(.responseBody)
		End With
		If Err Then
		Err.clear
		getfile="<font color='#ff0000'>error!</font>"
		End if
		Set Retrieval = Nothing
		end function

		Sub GetRemoteVer()
         response.write getfile("http://www.kes"& "ion.com/websystem/GetofficialInfo.asp?action=ver")
		End Sub

  Public Sub KS_Main()
           %>
           <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
			<html xmlns="http://www.w3.org/1999/xhtml">
			<head>
			<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
			<script src="../ks_inc/jquery.js"></script>
			<script src="../ks_inc/kesion.box.js"></script>
			<script>
			$(document).ready(function(){
			  $.get('index.asp',{action:'ver'},function(d){$('#versioninfo').html(d);});
			  
			  //检查是否存在升级文件
			  $.ajax({
			  url: "KS.Update.asp",
			  cache: false,
			  data: "action=check",
			  success: function(d){
			        d=unescape(d);
					switch (d){
					 case 'enabled':
					  $("#updateInfo").html("<font color='green'>对不起,您没有开启自动检测最新版本功能!</font>");
					  break;
					 case 'false':
					  $("#updateInfo").html("<font color='green'>当前已经是最新版本!</font>");
					  break;
					 case 'localversionerr':
					  $("#updateInfo").html("<font color='green'>加载本地xml版本文件出错,请检查<%=KS.Setting(89)%>include/version.xml文件是否存在!</font>");
					  break;
					 case 'remoteversionerr':
					  $("#updateInfo").html("<font color='green'>读取服务器文件出错,请检查<%=KS.Setting(89)%>admin_update.asp文件的配置是否正确或稍候再试!</font>");
					  break;
					 case 'unallow':
					  $("#updateInfo").html("<font color='green'>系统检查到有可更新文件,但不支持在线升级,请到官方站(<a href='http://www.kesion.com' target='_blank'>www.kesion.com</a>)下载升级文件!</font>");
					  break;
					 case 'unallowversion':
					  $("#updateInfo").html("<font color='green'>系统检查到有可更新文件,但由于您的版本号与最新版本号不对应,不支持在线升级,请根据您当前使用的版本到官方站(<a href='http://www.kesion.com' target='_blank'>www.kesion.com</a>)下载升级文件手工升级!</font>");
					  break;
					 default:
					    $("#updateInfo").html("<font color='red'>系统检查到有可升级文件!</font>");
					  	var str="<div style='height:auto;padding-top:10px;' id='updateResult'><font color=red>温馨提示:系统检测到有可升级的补丁,升级前如有自行修改请先做好备份!</font><div style='margin-top:3px'>"+d+"</div><br><div style='text-align:center'><input type='button' value='在线升级' id='openwin' style='height:25px;background:#efefef;border:1px solid #000'/>&nbsp;<input id='closewin' type='button' value='关闭取消' name='button' style='height:25px;background:#efefef;border:1px solid #000' /></div></div>"
					    popupTips('系统检查到有可升级文件',str,510,300);
						  $("#closewin").click(function(){
								closeWindow();
								$("#updateInfo").html("<font color='red'>您取消自动升级操作!</font>");
							});
							$("#openwin").click(function(){
								beginUpdate();
							});
					  break;
					}
			  }
		 	 });
			  
			 });
			 
			 function beginUpdate()
			 {
			  $("#updateInfo").html("<font color='red'>正在升级,请不要刷新本页面!</font>");
			   $.ajax({
			  url: "KS.Update.asp",
			  cache: false,
			  data: "action=update",
			  success: function(r){
			      r=unescape(r);
				  switch (r){
				    case "remoteversionerr":
					 $("#updateInfo").html("官方数据获取失败,被迫终止!");
					 alert('官方数据获取失败,被迫终止!');
					 closeWindow();
					 break;
					default :
					  $("#updateInfo").html("恭喜,在线升级成功!");
					  $("#updateResult").html(r);
					  break;
				  }
			  }
			  });
			  
			  
			 }
           </script>
			<style type="text/css">
			a{color:#555;}
			.position{ border-bottom:1px #83B5CD solid;background:url(images/titlebg.png); height:36px; font-size:13px; color:#555;line-height:36px; padding-left:10px;}
			.title{ background:#FBFDFF;border-top:2px solid #E1EEFF; line-height:28px; font-weight:bold;height:28px; color:#555; margin-left:20px;margin-right:20px;text-decoration:none;font-size:14px; margin-top:10px; padding-left:10px; padding-top:8px;}
			.title img{ padding-top:5px; padding-right:6px;}
			
			.nr{ height:auto; color:#555; text-decoration:none;font-size:12px; line-height:22px; padding-left:10px;margin-left:20px;margin-right:20px;}
			.nr ul{ padding:0px;margin:0px;}
			.nr li{text-alilgn:left;list-style-type:none;}
			.l {float:left}
			.l h2{font-size:13px;color:#ff6600}
			.box{clear:both}
			.newbox1{float:left;width:49%;}
			.newbox2{float:right;width:50%;}
			<%
			If Instr(KS.Setting(16),"2")=0 Then
			 KS.Echo ".bbs{display:none}"
			End If
			%>
			.bbs li{list-style-image:url(images/38.jpg)}
			</style>
			</head>
			
			<body scroll=no>
			
			
			<div class="position"><font color=red><%=KS.C("AdminName")%></font> 您好，欢迎进入网站后台系统！
			<%
								Dim RS:Set RS = Server.CreateObject("Adodb.Recordset")
								RS.Open "Select top 1 * From KS_Admin Where UserName='" & KS.C("AdminName") & "'", Conn, 1, 1
								If Not RS.EOF Then
								  KS.Echo "身份："
										  If KS.C("SuperTF") = 0 Then
										   KS.Echo "普通管理员"
										   Else
										   KS.Echo "超级管理员"
										   End If
									 KS.Echo " 登录次数：" & RS("LoginTimes") & "次 本次登录时间：" & RS("LastLoginTime")
								End If
								 RS.Close: Set RS=Nothing
			%></div>
			<div class="title"><img src="images/gif-0760.gif">版权声明：</div>
			<div class="nr">KesionCMS系统由漳州科兴信息技术有限公司(<a href="http://www.kesion.com" target="_blank">Kesion.Com</a>)独立开发，软件制作权登记号:<a href='http://www.kesion.com/images/v5dj.jpg' target='_blank'>2009SR00339</a>。授权给[<%=KS.Setting(0)%>]使用。任何个人或组织不得在授权允许的情况下删除、修改、拷贝本软件及其他副本上一切关于版权的信息。
			</div>
			
			<div class="title"><img src="images/gif-0760.gif">程序信息：</div>
			<div class="nr l">
			 <ul>
			   <li>当前版本：<%=KS.Version%></li>
			   <li>最新版本：<span id='versioninfo'><img src='images/loading.gif' align='absmiddle'>加载中...</span></li>
			   <li>产品开发：漳州科兴信息技术有限公司</li>
			   <li>咨询 Q Q：9537636 41904294 商业技术支持QQ：111394 54004407</li>
			   <li>公司网站：<a href='http://www.kesion.com/' target='_blank'>kesion.com</a> <a href='http://www.kesion.org/' target='_blank'>kesion.org</a> <a href='http://www.kesion.cn/' target='_blank'>kesion.cn</a></li>
			  </ul>
			  </div>
			 <div class="nr l">
			  <h2>在线升级</h2>
			  <span id='updateInfo'>正在检测最新版本信息...</span>  
			 </div>
			<div class="box">
			<div class="newbox1">
			<div class="title"><img src="images/gif-0760.gif">待审核信息：</div>
			<div class="nr">
			 <%
			 Dim Node,Num,Url,HasVerify
			 HasVerify=false
			 If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig()
			 For Each Node In Application(KS.SiteSN&"_ChannelConfig").DocumentElement.SelectNodes("channel[@ks21=1 and @ks0!=6 and @ks0!=9 and @ks0!=10]")
			   Num=Conn.Execute("Select count(id) from " & Node.SelectSingleNode("@ks2").text & " where verific=0")(0)
			   If Num=0 Then
			   'KS.Echo "待签" & Node.SelectSingleNode("@ks3").text & ":<font color=red>" & Num &" </font>" & Node.SelectSingleNode("@ks4").text & "&nbsp;"
			   Else
			    HasVerify=true
			   KS.Echo "<span style='cursor:pointer;' title='点击进入签收' onclick=""location.href='KS.ItemInfo.asp?ComeFrom=Verify&ChannelID=" & Node.SelectSingleNode("@ks0").text & "';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&opstr=" & server.URLEncode(Node.SelectSingleNode("@ks1").text & " >> <font color=red>签收" & Node.SelectSingleNode("@ks3").text & "</font>")&"';"">待签" & Node.SelectSingleNode("@ks3").text & "[<font color=red>" & Num &"</font>]" & Node.SelectSingleNode("@ks4").text & "</span>&nbsp;"
			   End If
			 Next
			 If KS.C_S(10,21)="1" Then
				Num=conn.execute("select count(id) from ks_Job_Company where status=0")(0)
				If Num>0 Then
				 HasVerify=true
				 KS.Echo "<span style='cursor:pointer;' title='点击进入审核' onclick=""location.href='KS.JobCompany.asp?ComeFrom=Verify';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&opstr=" & server.URLEncode("招聘求职管理 >> <font color=red>待审招聘单位</font>")&"';"">待审招聘单位[<font color=red>" & Num & "</font>]家</span>&nbsp;"
				End If
				Num=conn.execute("select count(id) from ks_Job_Resume where status=0")(0)
				If Num>0 Then
				 HasVerify=true
				 KS.Echo "<span style='cursor:pointer;' title='点击进入审核' onclick=""location.href='KS.JobResume.asp?ComeFrom=Verify';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&opstr=" & server.URLEncode("招聘求职管理 >> <font color=red>待审招聘单位</font>")&"';"">待审简历[<font color=red>" & Num & "</font>]份</span>&nbsp;"
				End If
				Num=conn.execute("select count(id) from KS_Job_Edu where status=0")(0)
				If Num>0 Then
				 HasVerify=true
				 KS.Echo "<span style='cursor:pointer;' title='点击进入审核' onclick=""location.href='KS.JobEdu.asp?status=0';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&opstr=" & server.URLEncode("招聘求职管理 >> <font color=red>待审招聘单位</font>")&"';"">待审教育背[<font color=red>" & Num & "</font>]份</span>&nbsp;"
				End If
			 End If
			 
			KS.Echo " <div style='height:22px;padding-top:3px;border-top:1px dashed #cccccc'>"
			Num=conn.execute("select count(id) from ks_comment where verific=0")(0)
			If Num>0 Then
			 HasVerify=true
			 KS.Echo "<span style='cursor:pointer;' title='点击进入审核' onclick=""location.href='KS.Comment.asp?ComeFrom=Verify';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&opstr=" & server.URLEncode("内容管理 >> <font color=red>审核评论</font>")&"';"">待审评论[<font color=red>" & Num & "</font>]条</span>"
			End If
			Num=conn.execute("select count(linkid) from ks_link where verific=0")(0)
			If Num>0 Then
			 HasVerify=true
			KS.Echo " <span style='cursor:pointer;' title='点击进入审核' onclick=""location.href='KS.FriendLink.asp?Action=Verific';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&opstr=" & server.URLEncode("友情链接管理 >> <font color=red>审核链接</font>")&"';"">待审链接[<font color=red>" & Num & "</font>]个</span>"
		    End If
			Num=conn.execute("select count(blogid) from ks_blog where status=0")(0)
			If Num>0 Then
			HasVerify=true
			KS.Echo " <span style='cursor:pointer;' title='点击进入审核' onclick=""location.href='KS.Space.asp?from=verify';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&opstr=" & server.URLEncode("空间门户管理 >> <font color=red>审核空间</font>")&"';"">待审空间[<font color=red>" & Num & "</font>]个</span>"
			End If
			Num=conn.execute("select count(id) from ks_bloginfo where status=2")(0)
			If Num>0 Then
			HasVerify=true
			KS.Echo " <span style='cursor:pointer;' title='点击进入审核' onclick=""location.href='KS.Spacelog.asp?from=verify';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&opstr=" & server.URLEncode("空间门户管理 >> <font color=red>审核日志</font>")&"';"">待审日志[<font color=red>" & Num & "</font>]篇</span>"
			End If
			Num=conn.execute("select count(id) from ks_photoxc where status=0")(0)
			If Num>0 Then
			HasVerify=true
			KS.Echo " <span style='cursor:pointer;' title='点击进入审核' onclick=""location.href='KS.SpaceAlbum.asp?from=verify';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&opstr=" & server.URLEncode("空间门户管理 >> <font color=red>审核相册</font>")&"';"">待审相册[<font color=red>" & Num & "</font>]本</span>"
			End If
			Num=conn.execute("select count(id) from ks_team where Verific=0")(0)
			If Num>0 Then
			HasVerify=true
			KS.Echo " <span style='cursor:pointer;' title='点击进入审核' onclick=""location.href='KS.SpaceTeam.asp?from=verify';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&opstr=" & server.URLEncode("空间门户管理 >> <font color=red>审核相册</font>")&"';"">待审圈子[<font color=red>" & Num & "</font>]个</span>"
			End If
			Num=conn.execute("select count(id) from KS_EnterpriseNews where status=0")(0)
			If Num>0 Then
			HasVerify=true
			KS.Echo " <span style='cursor:pointer;' title='点击进入审核' onclick=""location.href='KS.EnterPriseNews.asp?from=verify';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&opstr=" & server.URLEncode("空间门户管理 >> <font color=red>审核企业新闻</font>")&"';"">待审企业新闻[<font color=red>" & Num & "</font>]篇</span>"
			End If
			Num=conn.execute("select count(id) from KS_EnterPriseAD where status=0")(0)
			If Num>0 Then
			HasVerify=true
			KS.Echo " <span style='cursor:pointer;' title='点击进入审核' onclick=""location.href='KS.EnterPriseAD.asp?from=verify';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&opstr=" & server.URLEncode("空间门户管理 >> <font color=red>审核行业广告</font>")&"';"">待审行业广告[<font color=red>" & Num & "</font>]个</span>"
			End If
			'Num=conn.execute("select count(id) from KS_EnterPriseZS where status=0")(0)
			'If Num>0 Then
			'HasVerify=true
			'KS.Echo " <span style='cursor:pointer;' title='点击进入审核' onclick=""location.href='KS.EnterPriseZS.asp?from=verify';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&opstr=" & server.URLEncode("空间门户管理 >> <font color=red>审核证书</font>")&"';"">待审证书[<font color=red>" & Num & "</font>]个</span>"
			'End If
			
			
			
			
			If HasVerify=false Then
			 KS.Echo "<div style='margin:30px;text-align:center;color:red'>今天没有用户提交待审核的信息！</div>"
			End If
			KS.Echo "</div>"
			   %>
			</div>
			</div>
			<div class="newbox2">
			<div class="title"><img src="images/gif-0760.gif">技术论坛新帖：</div>
			<div class="nr">
			 <ul class="bbs"><script  id="showtopic" src="http://bbs.kesion.com/Dv_News.asp?GetName=newtopic"></script>
			  </ul>
			</div>
			</div>
			
			</div>
			</body>
			</html>

          <%
				Conn.Close:Set Conn = Nothing
			End Sub
			
			Public Sub KS_Foot()
		     With Response
				.Write "<html>"
				.Write "<script language=""JavaScript"" src=""Include/SetFocus.js""></script>"
		        .Write "<script language='JavaScript' src='../KS_Inc/common.js'></script>"
				.Write "<META http-equiv=Content-Type content=""text/html; charset=gb2312"">"
		        .Write "<link href=""Skin/Style"&KS.C("SkinID") &".CSS"" rel=""stylesheet"" type=""text/css"">"
				.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"" background="""">"
				.Write "<div id='foot'>"
				%>
				<div id='color'><a href="KS.SetSkin.asp?skinid=1" target="_top"><img style="margin:2px;" src='images/ico/skin1.gif' border="0"></a><a href="KS.SetSkin.asp?skinid=2" target="_top"><img style="margin:2px;" src='images/ico/skin2.gif' border="0"></a><a href="KS.SetSkin.asp?skinid=3" target="_top"><img style="margin:2px;" src='images/ico/skin3.gif' border="0"></a><a href="KS.SetSkin.asp?skinid=4" target="_top"><img style="margin:2px;" src='images/ico/skin4.gif' border="0"></a><a href="KS.SetSkin.asp?skinid=5" target="_top"><img style="margin:2px;" src='images/ico/skin5.gif' border="0"></a>
                </div>
				<%
				.Write "<div id='co' align=""center"" onClick=""ChangeLeftFrameStatu();"" title=""全屏/半屏"" style=""cursor:pointer;""><font color=red>×</font> 关闭左栏</div>"
				.Write "<div id='footmenu'>快速通道=>："
				If KS.ReturnPowerResult(0, "KMTL20000") Then
				.Write "<a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'发布中心 >> <font color=red>发布管理首页</font>','disabled','Include/refreshindex.asp');"">发布首页</a>"
				End If
				.Write "<a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'发布中心 >> <font color=red>发布管理首页</font>','disabled','Include/RefreshHtml.asp?ChannelID=1');"">发布管理</a>"
				
				If KS.ReturnPowerResult(0, "KMTL10007") Then
				.Write "<a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'模板标签管理 >> <font color=red>模板管理</font>','disabled','KS.Template.asp');"">模板管理</a>"
				End If
				If KS.ReturnPowerResult(0, "KMST10001") Then
				.Write "<a href='javascript:void(0)' onClick=""SelectObjItem1(this,'系统设置 >> <font color=red>基本信息设置</font>','SetParam','KS.System.asp');"" title='基本信息设置'>基本信息设置</a>"
				End If
				If Instr(KS.C("ModelPower"),"model1")>0 Or KS.C("SuperTF")="1" then
				.Write "<a href='javascript:void(0)' onClick=""SelectObjItem1(this,'模型管理 >> <font color=red>模型管理首页</font>','SetParam','KS.Model.asp');"">模型管理</a>"
				End If
				If KS.ReturnPowerResult(0, "KMUA10011") Then
				.Write "<a href='javascript:void(0)' onClick=""SelectObjItem1(this,'用户管理 >> <font color=red>检查管理员工作进度</font>','SetParam','KS.UserProgress.asp');"">查看工作进度</a>"
			    End If
				.Write "</div>"
				.Write "<div id='footcopyright'>版权所有 &copy; 2006-2010 科兴信息技术有限公司</div>"
				.Write "</div>"
				
				.Write "</body>"
				.Write "</html>"
				.Write "<SCRIPT language=javascript>"
				.Write "    var screen=false;"
				.Write "    function ChangeLeftFrameStatu()"
				.Write "    {"
				.Write "        if(screen==false)"
				.Write "        {"
				.Write "            parent.FrameMain.cols='0,*';"
				.Write "            screen=true;"
				.Write "            self.co.innerHTML = ""√ 打开左栏"""
				.Write "        }"
				.Write "        else if(screen==true)"
				.Write "        {"
				.Write "            parent.FrameMain.cols='201,*';"
				.Write "           screen=false;"
				.Write "            self.co.innerHTML = ""<font color=red>×</font> 关闭左栏"""
				.Write "        }"
				.Write "    }"
				.Write "</SCRIPT>"
			End With
		End Sub
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

End Class
%> 
