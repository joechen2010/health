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
		.Write "<title>" & KS.Setting(0) & "---��վ��̨����</title>"
		.Write "<script language=""JavaScript"">" & vbCrLf
		.Write "<!--" & vbCrLf
		.Write "   //���渴��,�ƶ��Ķ���,ģ����а幦��" & vbCrLf
		.Write "  function CommonCopyCutObj(ChannelID, PasteTypeID, SourceFolderID, FolderID, ContentID)" & vbCrLf
		.Write "  {" & vbCrLf
		.Write "   this.ChannelID=ChannelID;             //Ƶ��ID" & vbCrLf
		.Write "   this.PasteTypeID=PasteTypeID;         //�������� 0---���κβ���,1---����,2---����" & vbCrLf
		.Write "   this.SourceFolderID=SourceFolderID;   //���ڵ�ԴĿ¼" & vbCrLf
		.Write "   this.FolderID=FolderID;               //Ŀ¼ID" & vbCrLf
		.Write "   this.ContentID=ContentID;             //���»�ͼƬ��ID" & vbCrLf
		.Write "  }" & vbCrLf
		.Write "  function CommonCommentBack(FromUrl)" & vbCrLf
		.Write "  {" & vbCrLf
		.Write "    this.FromUrl=FromUrl;             //������Դҳ�ĵ�ַ" & vbCrLf
		.Write "  }" & vbCrLf
		.Write "  //��ʼ������ʵ��" & vbCrLf
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
		.Write "            <frame src=""KS.Split.asp?ButtonSymbol=Disabled&OpStr=" & Server.URLEncode("<font color=red>ϵͳ��������</font>") & """ name=""BottomFrame"" ID=""BottomFrame"" frameborder=""no"" scrolling=""no"" marginwidth=""0"" marginheight=""0""></frame>" & vbCrLf
		.Write "        </frameset>" & vbCrLf
		.Write "  </frameset>" & vbcrlf
		.Write "  <frame src=""Index.asp?Action=Foot"" name=""FrameBottom"" id=""FrameBottom"" noresize frameborder=""no"" scrolling=""no"" marginwidth=""0"" marginheight=""0""></frame>" & vbCrLf
		.Write "</frameset>" & vbcrlf
		.Write "<noframes>����������汾̫��,�밲װIE5.5�����ϰ汾!</noframes>" & vbcrlf
		.Write "</html>" & vbCrLf
		End With
		End Sub
		
		Public Sub KS_Head()
			 On Error Resume Next
			 With Response
			 .Buffer = True
			If Trim(Request.ServerVariables("HTTP_REFERER")) = "" Then
				.Write "<br><p align=center><font color='red'>�Բ���Ϊ��ϵͳ��ȫ��������ֱ�������ַ���ʱ�ϵͳ�ĺ�̨����ҳ�档</font></p>"
				Exit Sub
			End If
			
			.Write "<html>"
		    .Write "<script language='JavaScript' src='../KS_Inc/common.js'></script>"
		    .Write "<script language='JavaScript' src='../KS_Inc/jquery.js'></script>"
			.Write "<link href=""Skin/Style"& KS.C("SkinID") &".CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "<meta http-equiv=Content-Type content=""text/html; charset=GB2312"">"
			.Write "<script language=""javascript"">"& vbcrlf
			.Write " function out(src){"& vbcrlf
			.Write " if(confirm('ȷ��Ҫ�˳���')){"& vbcrlf
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
			.Write "<div id='ajaxmsg' style='text-align:center;background-color: #ffffee;border: 1px #f9c943 solid;position:absolute; z-index:1; left: 200px; top: 5px;display:none;'> <img src='images/loading.gif'> ���Ժ�,����ִ����������...  </div>"
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
			 .write "                 <td " & KSAnnounceDisplayFlag & " class=""font_text"" align=""right""><font color=#ffffff>�ٷ����棺</font></td>"
			 .Write "                 <td " & KSAnnounceDisplayFlag & "  width=""40%"">"
			 .Write "<iframe scrolling=no src=""http://www.kesion.com/websystem/GetofficialInfo.asp"" name=""ShowAnnounce"" id=""ShowAnnounce"" height=""22"" WIDTH=""100%"" marginheight=""0"" marginwidth=""0"" frameborder=""0"" align=""middle"" allowtransparency=""true""></iframe>"
			 .Write "</td>"
			 
			.Write "                <td class=""font_text"" align=""right""> [<a href=""" & KS.GetDomain &""" target=""_blank"" class=""white"">��վ��ҳ</a>] [<a href=""" & KS.GetDomain &"User/index.asp?User_Message.asp?action=inbox"" target=""_blank"" class=""white"">�鿴����</a><span id='newmessage'>(<font color=#ff0000>0</font>)</span>] "
			If KS.ReturnPowerResult(0, "KMUA10010") Then
			.Write "[<a href=""#"" onClick=""OpenWindow('KS.Frame.asp?Url=KS.Admin.asp&Action=SetPass&PageTitle=" & server.URLEncode("�޸ĺ�̨��¼����") & "',360,175,window);"" class=""white"">�޸�����</a>] "
			End If
			If KS.ReturnPowerResult(0, "KMST20000") Then
			.Write "[<a href=""KS.CleanCache.asp"" target=""MainFrame"" class=""white"">���»���</a>] "
			End If
			.WRite "[<a href=""Login.asp?Action=LoginOut"" target=""_top"" onClick=""return out(this)""  class=""white"">��ȫ�˳�</a>]"
			
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
		 .Write " var SearchPower" & SQL(0,I) & "='" & KS.ReturnPowerResult(SQL(0,I), "M"&SQL(0,I)&"10008")&"';    //����Ȩ��" & vbCrLf
       Next
		.Write " var SearchSpecialPower='" & KS.ReturnPowerResult(0, "KMSP10004") & "';    //����ר��Ȩ��" & vbCrLf
		.Write " var SearchLinkPower='" & KS.ReturnPowerResult(0, "KMCT10001") & "';       //�����������ӵ�Ȩ��" & vbCrLf
		.Write " var SearchAdminPower='" & KS.ReturnPowerResult(0, "KMUA10001") & "';      //��������ԱȨ��" & vbCrLf
		.Write " var SearchSysLabelPower='" & KS.ReturnPowerResult(0, "KMTL10001") & "';   //����ϵͳ������ǩȨ��" & vbCrLf
		.Write " var SearchDIYFunctionLabelPower='" & KS.ReturnPowerResult(0, "KMTL10002") & "';   //�����Զ��庯����ǩȨ��" & vbCrLf
		.Write " var SearchFreeLabelPower='" & KS.ReturnPowerResult(0, "KMTL10003") & "';  //�����Զ��徲̬��ǩȨ��" & vbCrLf
		.Write " var SearchSysJSPower='" & KS.ReturnPowerResult(0, "KMTL10004") & "';      //����ϵͳJSȨ��" & vbCrLf
		.Write " var SearchFreeJSPower='" & KS.ReturnPowerResult(0, "KMTL10005") & "';     //��������JSȨ��" & vbCrLf
		.Write "</script>"
		.Write "<script language=""JavaScript"" src=""Include/SetFocus.js""></script>"
		.Write "<script language='JavaScript' src='../KS_Inc/common.js'></script>"
		.Write "<script language='JavaScript' src='../KS_Inc/Jquery.js'></script>"
		%>
		<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
         <html xmlns="http://www.w3.org/1999/xhtml">
		<head><title>��Ѵ��վ����ϵͳV6.0��̨</title>
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
		var left=0,top=0,title='����С����';
		var SearchBodyStr=''
						   +'<table width="100%" border="0" cellspacing="0" cellpadding="0">'
						   +'<form name="searchform" target="MainFrame" method="post">'
						   +'<tr> '
						   +'<td height="25"><strong>�����������ȫ��������������</strong></td>'
						   +' </tr>'
						   +'<tr><td height="25">ȫ���򲿷ֹؼ���</td></tr>'
						   +'<tr><td height="25"><input style="width:95%" type="text" id="KeyWord" name="KeyWord"></td></tr>'
						   +'  <tr><td height="25">������Χ</td></tr>'
						   +'  <tr><td height="25"> <select style="width:95%" id="SearchArea" name="SearchArea" onchange="SetSearchTypeOption(this.options[this.selectedIndex].text)">'
						   +'     </select></td></tr>' 
						   +'<tr><td height="25">��������</td></tr>'
						   +'<tr><td height="25"><select style="width:95%" id="SearchType" name="SearchType">'
						   +'</select></td></tr>'
						   +'  <tr id="DateArea" onclick="setstatus(this)" style="cursor:pointer"><td height="25"><strong>ʲôʱ���޸ĵ�?</strong></td></tr>'
						   +'  <tr style="display:none"><td height="25">��ʼ����<input type="text" readonly style="width:80%" name="StartDate" id="StartDate">'
						   +'  <span style="cursor:pointer" onClick=OpenThenSetValue("Include/DateDialog.asp",160,170,window,document.all.StartDate);document.all.StartDate.focus();><img src="Images/date.gif" border="0" align="absmiddle" title="ѡ������"></span></td></tr>'
						   +'  <tr style="display:none"><td height="25">��������<input type="text" readonly style="width:80%" name="EndDate" id="EndDate">'
						   +'  <span style="cursor:pointer" onClick=OpenThenSetValue("Include/DateDialog.asp",160,170,window,document.all.EndDate);document.all.EndDate.focus();><img src="Images/date.gif" border="0" align="absmiddle" title="ѡ������"></span></td></tr>'
						   +'  <tr><td height="40" align="center"><input type="submit" name="SearchButton" value="��ʼ����" onclick="return(SearchFormSubmit())"></td></tr>'
						   +'</form>'
						   +'  <tr><td><strong>ʹ��˵��:</strong></td></tr>'
						   +'  <tr><td> �� ���������ñ������������������¡�ͼƬ������Flash��ר�⡢��ǩ��JS��,������������Ŀ¼������Ƶ�����ơ���Ŀ���ƣ���ǩĿ¼��</td></tr>'
						   +'  <tr><td> �� �� <font color=red>Ctrl+F</font> ���Կ��ٽ��д򿪻�ر�����С����</td></tr>'
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
		//�ر�;
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
		//��ʼ��;
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
			 sTextArr=new Array(<%=ModelList%>'ר������','��������վ��','ϵͳ������ǩ','�Զ��庯����ǩ','�Զ��徲̬��ǩ','ϵͳ JS','���� JS','����Ա')
			 ChannelIDArr=new Array(<%=ChannelIDList%>'ר������','��������վ��','ϵͳ������ǩ','�Զ��庯����ǩ','�Զ��徲̬��ǩ','ϵͳ JS','���� JS','����Ա')
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
			//����Ȩ�޼��,��û��Ȩ�޵�����ģ��,��������	
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
			  //�ı�ѡ��Χʱ��ȡ����ȷ��ģ��ID
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
				 if (SearchPower<%= SQL(0,I)%>=='False')          //����Ȩ�޼��
				 {
				  DisabledSearchFluctuation(true);
				  return;
				 }
				 else
				 {
				  DisabledSearchFluctuation(false);
				  jQuery('#DateArea').show();
				  TextArr=new Array('<%=SQL(3,I)%>����','<%=SQL(3,I)%>����','<%=SQL(3,I)%>�ؼ���','<%=SQL(3,I)%>����','<%=SQL(3,I)%>¼��')
				  }
				  break;
		   <% End If
		   Next%>
			case 'ר������':
				 if (SearchSpecialPower=='False')        //����ר��Ȩ�޼��
				 {
				   DisabledSearchFluctuation(true);
				   return;
				 }
				 else
				 {
				  DisabledSearchFluctuation(false);
				  jQuery('#DateArea').show();
				  TextArr=new Array('ר������','��Ҫ˵��')
				 }
				 break;
			case '��������վ��':
				 if (SearchLinkPower=='False')       //������������վ��Ȩ�޼��
				 {
				   DisabledSearchFluctuation(true);
				   return;
				 }
				 else{
				  DisabledSearchFluctuation(false);
				  jQuery('#DateArea').show();
				  TextArr=new Array('վ������','վ������')
				 }
				 break;
			case 'ϵͳ������ǩ':
				 if (SearchSysLabelPower=='False')       //����ϵͳ��ǩȨ�޼��
				 {
				   DisabledSearchFluctuation(true);
				   return;
				 }else{
				  DisabledSearchFluctuation(false);
				 jQuery('#DateArea').show();
				 TextArr=new Array('ϵͳ��ǩ����','ϵͳ��ǩ����')
				 }
				 break;
			case '�Զ��庯����ǩ':
				 if (SearchDIYFunctionLabelPower=='False')       //�����Զ��庯����ǩȨ�޼��
				 {
				   DisabledSearchFluctuation(true);
				   return;
				 }else{
				  DisabledSearchFluctuation(false);
				 jQuery('#DateArea').show();
				 TextArr=new Array('�Զ��庯����ǩ����','�Զ��庯����ǩ����')
				 }
				 break;
			case '�Զ��徲̬��ǩ':
				 if (SearchFreeLabelPower=='False')       //�����Զ��徲̬��ǩȨ�޼��
				 {
				   DisabledSearchFluctuation(true);
				   return;
				 }else{
				 jQuery('#DateArea').show()
				 TextArr=new Array('�Զ��徲̬��ǩ����','�Զ��徲̬��ǩ����','�Զ��徲̬��ǩ����')
				 }
				 break;
			case 'ϵͳ JS':
				 if (SearchSysJSPower=='False')       //����ϵͳJSȨ�޼��
				 {
				   DisabledSearchFluctuation(true);
				   return;
				 }else{
				 jQuery('#DateArea').show();
				 TextArr=new Array('ϵͳJS ����','ϵͳJS ����','ϵͳJS �ļ���')
				 }
				 break;
			case '���� JS' :
				 if (SearchFreeJSPower=='False')       //��������JSȨ�޼��
				 {
				   DisabledSearchFluctuation(true);
				   return;
				 }else{
				 jQuery('#DateArea').show();
				 TextArr=new Array('����JS ����','����JS ����','����JS �ļ���')
				 }
				 break;
			case '����Ա':	 
				  if (SearchAdminPower=='False')          //��������ԱȨ�޼��
				 {
				  DisabledSearchFluctuation(true);
				  return;
				 }else{
				  DisabledSearchFluctuation(false);
				 jQuery('#DateArea').show();
				 TextArr=new Array('����Ա����','����Ա���')
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
			   alert('������ؼ���!')
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
				   parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape("��Ϣ�������� >> <font color=red>�������</font>")+'&ButtonSymbol=Search';
				   break;
			  case 'Special':
				   form.action="KS.Special.asp?Action=SpecialList";
				   parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape("ר����� >> <font color=red>����ר����</font>")+'&ButtonSymbol=SpecialSearch';
				   break;
			  case 'Link':
				   form.action="KS.FriendLink.asp";
				   parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape("������� >> �������ӹ��� >> <font color=red>������������վ����</font>")+'&ButtonSymbol=LinkSearch';
				   break;
			  case 'SysLabel'  :
				   form.action="Include/Label_Main.asp?LabelType=0";
				   parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape("��ǩ���� >> <font color=red>����ϵͳ������ǩ���</font>")+'&ButtonSymbol=SysLabelSearch';
				   break;
			 case 'DIYFunctionLabel'  :
				   form.action="Include/Label_Main.asp?LabelType=5";
				   parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape("��ǩ���� >> <font color=red>�����Զ��庯����ǩ���</font>")+'&ButtonSymbol=DIYFunctionSearch';
				   break;
			  case 'FreeLabel'  :
				   form.action="Include/Label_Main.asp?LabelType=1";
				   parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape("��ǩ���� >> <font color=red>�������ɱ�ǩ���</font>")+'&ButtonSymbol=FreeLabelSearch';
				   break;
			  case 'SysJS'     :
				   form.action="Include/JS_Main.asp?JSType=0";
				   parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape("JS���� >> <font color=red>����ϵͳJS���</font>")+'&ButtonSymbol=SysJSSearch';
				   break;
			  case 'FreeJS'     :
				   form.action="Include/JS_Main.asp?JSType=1";
				   parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape("JS���� >> <font color=red>��������JS���</font>")+'&ButtonSymbol=FreeJSSearch';
				   break;
			  case 'Manager'     :
				   form.action="KS.Admin.asp";
				   parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape("����Ա���� >> <font color=red>��������Ա���</font>")+'&ButtonSymbol=ManagerSearch';
				   break;
			}
			form.submit();
		  }
		function DisabledSearchFluctuation(Flag)
		{ if (Flag==true)
		   document.all.KeyWord.value='�Բ���,Ȩ�޲���!'; 
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
					<li class="Selected" id="left_tab1" title="���ݹ���" onClick="javascript:showleft(1);" name="left_tab1">��<br>��</li>
					<li id="left_tab2" <%If Instr(Request.Cookies(KS.SiteSn)("ModelPower"),"sysset1")<=0 and Request.Cookies(KS.SiteSn)("SuperTF")<>"1" then response.Write(" style='display:none' ") %>onClick="javascript:showleft(2);" title="ϵͳ����">��<br>��</li>		
					<li id="left_tab3" <%If Instr(Request.Cookies(KS.SiteSn)("ModelPower"),"subsys1")<=0 and Request.Cookies(KS.SiteSn)("SuperTF")<>"1" then response.Write(" style='display:none' ") %>onClick="javascript:showleft(3);" title="��ز���">��<br>��</li>
					<li id="left_tab4" <%If Instr(Request.Cookies(KS.SiteSn)("ModelPower"),"model1")<=0 and Request.Cookies(KS.SiteSn)("SuperTF")<>"1" then response.Write(" style='display:none' ") %>onClick="javascript:showleft(4);" title="ģ�͹���">ģ<br>��</li>
					<li id="left_tab5" <%If Instr(Request.Cookies(KS.SiteSn)("ModelPower"),"lab1")<=0 and Request.Cookies(KS.SiteSn)("SuperTF")<>"1" then response.Write(" style='display:none' ") %>onClick="javascript:showleft(5);" title="��ǩ">��<br>ǩ</li>
					<li id="left_tab6" <%If Instr(Request.Cookies(KS.SiteSn)("ModelPower"),"user1")<=0 and Request.Cookies(KS.SiteSn)("SuperTF")<>"1" then response.Write(" style='display:none' ") %>onClick="javascript:showleft(6);" title="�û�����">��<br>��</li>
					<li id="left_tab7" title="���" <%If Instr(Request.Cookies(KS.SiteSn)("ModelPower"),"other1")<=0 and Request.Cookies(KS.SiteSn)("SuperTF")<>"1" then response.Write(" style='display:none' ") %>onClick="javascript:showleft(7);" name="left_tab7">��<br>
				   ��</li>
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
					 
					 <!--------------���ݹ��� start-------------------->
					 <DIV  class="left_menu" id="td_<%=n+1%>" onClick="javascript:switchShow(<%=n+1%>,1);" height=26>&nbsp;&nbsp;<a href="javascript:void(0)">���ݹ���</a></DIV>
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
						   <a href="javascript:void(0)"  onClick="SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red><%=SQL(3,I)%>����</font>','ViewFolder','KS.ItemInfo.asp?ChannelID=<%=SQL(0,I)%>',<%=SQL(0,I)%>);"><%=KS.Gottopic(SQL(1,I),8)%></a> <span style="cursor:pointer" onClick="SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red>���<%=SQL(3,I)%></font>','AddInfo','<%=ItemManageUrl%>?Action=Add&ChannelID=<%=SQL(0,I)%>',<%=SQL(0,I)%>);"><img alt="���<%=SQL(3,I)%>" src="images/add.gif" border="0" align="absmiddle"></span><%if KS.ReturnPowerResult(SQL(0,I), "M"&SQL(0,I)&"10012") then%> <span style="cursor:pointer" onClick="SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red>ǩ��<%=SQL(3,I)%></font>','Disabled','KS.ItemInfo.asp?ComeFrom=Verify&ChannelID=<%=SQL(0,I)%>',<%=SQL(0,I)%>);"><img alt="ǩ��<%=SQL(3,I)%>" src="images/accept.gif" border="0" align="absmiddle"></span>
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
							 jQuery("#classOpen").html("<img style='cursor:pointer' id='classOpen' onclick='TipsToggle(\"show\")' src='images/left_down.gif' align='absmiddle' title='չ��'>");
							 }else{
							 jQuery("#modelxg").show('fast');
							 jQuery("#classOpen").html("<img style='cursor:pointer' id='classOpen' onclick='TipsToggle(\"hide\")' src='images/left_up.gif' title='�ղ�' align='absmiddle'>");						
                              	 }
						   }
						  </script>
						  
                           <div  id="modelxg">
						   <%If KS.ReturnPowerResult(0, "M010001") Then %>
						   <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'���ݹ��� >> <font color=red><%=SQL(3,I)%>��Ŀ����</font>','Disabled','KS.Class.asp');">��Ŀ����</a> <a href='javascript:void(0)' onClick="SelectObjItem1(this,'��Ŀ���� >> <font color=red>�����Ŀ</font>','Go','KS.Class.asp?Action=Add&FolderID=1','');">���</a></li>
						   <%End If%>
						   <%If KS.ReturnPowerResult(0, "M010002") Then %>
						    <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'���ݹ��� >> <font color=red>���۹���</font>','Disabled','KS.Comment.asp');">���۹���</a> <a href='javascript:void(0)' onClick="SelectObjItem1(this,'���ݹ��� >> <font color=red>���۹���</font>','Disabled','KS.Comment.asp?ComeFrom=Verify');">���</a> </li>
							<%End If%>
							<%If KS.ReturnPowerResult(0, "M010003") Then %>
						    <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'���ݹ��� >> <font color=red>ר�����</font>','Disabled','KS.Special.asp');">ȫվר�����</a> </li>
							<%End If%>
							<%If KS.ReturnPowerResult(0, "M010004") Then %>
						    <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'���ݹ��� >> <font color=red>�ؼ���Tags����</font>','Disabled','KS.KeyWord.asp');">�ؼ���Tags����</a> </li>
							<%End If%>
                            <%If KS.ReturnPowerResult(0, "M010005") or KS.ReturnPowerResult(0, "M010006") Then %>
							<li>
							<%If KS.ReturnPowerResult(0, "M010005") Then%><a href='javascript:void(0)' onClick="SelectObjItem1(this,'���ݹ��� >> <font color=red>��������</font>','Disabled','KS.ItemInfo.asp?Action=SetAttribute');">��������</a><%end if%><%If KS.ReturnPowerResult(0, "M010006") then%> <a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'��ز��� >> <font color=red>�������վ</font>','ViewFolder','KS.ItemInfo.asp?ComeFrom=RecycleBin','');">�� �� վ</a><%end if%></li>
							<%End If%>
						   <%If KS.ReturnPowerResult(0, "M010007") Then %>
						   <li><a href="KS.Tools.asp"  target="MainFrame" title="һ��������">һ��������</a></li>
						   <%end if%>
						   <%If KS.ReturnPowerResult(0, "M010008") Then %>
						   <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'���ݹ��� >> <font color=red>��Ϣ�ɼ�����</font>','Disabled','Collect/Collect_Main.asp?ChannelID=1');">��Ϣ�ɼ�����</a> </li>
						   <%End if%>
						   </div>
						

						   
						 </div>
					 </div>
					<!--------------���ݹ��� end-------------------->  
					
					
					<!--------------�̳ǹ��� start-------------------->
				  <%
				  IF instr(lcase(KS.C("ModelPower")),"shop0")=0 or KS.C("SuperTf")=1 Then
					 If Not ModelXML Is Nothing Then
					  If ModelXML.documentElement.SelectNodes("row[@channelid=5 and @channelstatus=1]").length<>0 Then
					   N=N+1
					 %>
					 <DIV  class="left_menu" id="td_<%=n+1%>" onClick="javascript:switchShow(<%=n+1%>,1);" height=26>&nbsp;&nbsp;<a href="javascript:void(0)">�̳ǹ���</a></DIV>
					 <DIV class="hid" id="subTable<%=n+1%>" style="WIDTH: 100%">
					    <%If KS.ReturnPowerResult(5, "M510012") Then %>				
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'�̳�ϵͳ >> <font color=red>����24Сʱ�ڶ���</font>','Disabled','KS.ShopOrder.asp?searchtype=1&ChannelID=5');"><font color=red>����24Сʱ�ڶ���</font></a></li>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'�̳�ϵͳ >> <font color=red>�������ж���</font>','Disabled','KS.ShopOrder.asp?ChannelID=5');">�������ж���</a></li>
						 <%end if%>
						 <%If KS.ReturnPowerResult(5, "M510014") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'�̳�ϵͳ >> <font color=red>�ʽ���ϸ��ѯ</font>','Disabled','KS.LogMoney.asp?ChannelID=<%=SQL(0,I)%>');">�ʽ���ϸ��ѯ</a></li>
						 <%end if%>
						 <%If KS.ReturnPowerResult(5, "M510015") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'�̳�ϵͳ >> <font color=red>���˻���ѯ</font>','Disabled','KS.LogDeliver.asp?ChannelID=5');">���˻���ѯ</a></li>
						 <%end if%>
						 <%If KS.ReturnPowerResult(5, "M510016") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'�̳�ϵͳ >> <font color=red>���˻���ѯ</font>','Disabled','KS.LogInvoice.asp?ChannelID=5');">����Ʊ��ѯ</a></li>
						 <%end if%>
						 <%If KS.ReturnPowerResult(5, "M510017") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'�̳�ϵͳ >> <font color=red>�ͻ�ͳ��</font>','Disabled','KS.ShopStats.asp?Action=Custom');">��������ͳ��</a></li>
						 <%End If%>
						 
						 <%If KS.ReturnPowerResult(5, "M520003") Then %>
						 ====================
						  <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'�̳�ϵͳ >> <font color=red>���̹���</font>','Disabled','KS.Author.asp?ChannelID=5');">���̹���</a> </li>
						  <%end if%>
						  <%If KS.ReturnPowerResult(5, "M520004") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'�̳�ϵͳ >> <font color=red>�ͻ���ʽ����</font>','Disabled','KS.Delivery.asp?ChannelID=5');">�ͻ�&���ʽ</a></li>
						 <%end if%>
						 <%If KS.ReturnPowerResult(5, "M520001") Then %>
						  <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'�̳�ϵͳ >> <font color=red>�����ص����</font>','Disabled','KS.ShopUnion.asp');">�����ص����</a></li>
						 <%End If%>
					   	  <%If KS.ReturnPowerResult(5, "M510018") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'Ʒ�ƹ��� >> <font color=red>Ʒ�ƹ���</font>','Disabled','KS.ShopBrand.asp');">Ʒ�ƹ���</a> <a href='javascript:void(0)' onClick="SelectObjItem1(this,'Ʒ�ƹ��� >> <font color=red>���Ʒ��</font>','Go','KS.ShopBrand.asp?Action=Add&FolderID=0',5);">���</a> <a href='javascript:void(0)' onClick="SelectObjItem1(this,'Ʒ�ƹ��� >> <font color=red>����Ʒ�Ƶ�JS�˵�</font>','Go','KS.ShopBrand.asp?Action=Create&FolderID=0',5);">����</a></li>	
						 <%end if%>	
						 
						 ====================
						 <%If KS.ReturnPowerResult(5, "M520008") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'�̳�ϵͳ >> <font color=red>������Ʒ����</font>','Disabled','KS.Shop.asp?action=LimitBuy&channelid=5');">��ʱ/������������</a></li>
						 <%End If%>
						 <%If KS.ReturnPowerResult(5, "M520009") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'�̳�ϵͳ >> <font color=red>��ʱ������Ʒ����</font>','Disabled','KS.Shop.asp?action=BundleSale&channelid=5');">����������Ʒ����</a></li>
						 <%End If%>
						 <%If KS.ReturnPowerResult(5, "M520010") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'�̳�ϵͳ >> <font color=red>������Ʒ����</font>','Disabled','KS.Shop.asp?action=ChangedBuy&channelid=5');">������Ʒ����</a></li>
						 <%End If%>
						 <%If KS.ReturnPowerResult(5, "M520011") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'�̳�ϵͳ >> <font color=red>��ֵ�������</font>','Disabled','KS.Shop.asp?action=Package&channelid=5');">��ֵ�������</a></li>
						 <%end if%>
						 <%If KS.ReturnPowerResult(5, "M510005") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'�̳�ϵͳ >> <font color=red>������������</font>','Disabled','KS.ItemInfo.asp?action=SetAttribute&channelid=5');">������������</a></li>
						 <%End If%>
						 <%If KS.ReturnPowerResult(5, "M520007") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red>�Ż�ȯ����</font>','Disabled','KS.ShopCoupon.asp');">�Ż�ȯ����</a></li>
						 <%End If%>
						  <%If KS.ReturnPowerResult(5, "M530001") Then %>
						  <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'�Ź�ϵͳ >> <font color=red>�Ź�������ҳ</font>','Disabled','KS.GroupBuy.asp');">�Ź�������ҳ</a> <a href='javascript:void(0)' onClick="SelectObjItem1(this,'�Ź�ϵͳ >> <font color=red>�Ź�������ҳ</font>','Go','KS.GroupBuy.asp?Action=Add');">���</a></li>
						  <%End If%>
					   <%If KS.ReturnPowerResult(5, "M530002") Then %>
                       <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'�Ź�ϵͳ >> <font color=red>��Ȥ��������</font>','Disabled','KS.GroupBuyInt.asp');">��Ȥ��������</a></li>
					   <%end if%>
					 
					 </DIV>
					 <!--------------�̳ǹ��� End-------------------->
					<% End If
					End If
				  End If
					%>
					
					   
					
					<!--------------���ֹ��� start-------------------->
					 <%
					IF instr(lcase(KS.C("ModelPower")),"music0")=0 or KS.C("SuperTf")=1 Then
					 If Not ModelXML Is Nothing Then
					  If ModelXML.documentElement.SelectNodes("row[@channelid=6 and @channelstatus=1]").length<>0 Then
					   N=N+1
					  %>
					 <DIV  class="left_menu" id="td_<%=n+1%>" onClick="javascript:switchShow(<%=n+1%>,1);" height=26>&nbsp;&nbsp;<a href="javascript:void(0)">���ֹ���</a></DIV>
					  <DIV class="hid" id="subTable<%=n+1%>" style="WIDTH: 100%">
					  <li><a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'����ϵͳ >> <font color=red>���и�������</font>','Disabled','KS.Music.asp?url=KS.MusicSong.asp');">��������</a> <a href='javascript:void(0)' onClick="SelectObjItem1(this,'����ϵͳ >> <font color=red>��Ӹ���</font>','Disabled','KS.Music.asp?url=KS.MusicSong.asp?Action=Add&Classid=1',6);">��Ӹ���</a></li>
						 
						 <li><a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'����ϵͳ >> <font color=red>���и�������</font>','Disabled','KS.Music.asp?url=KS.MusicSpecial.asp');">ר������</a> <a href='javascript:void(0)' onClick="SelectObjItem1(this,'����ϵͳ >> <font color=red>��Ӹ���</font>','Disabled','KS.Music.asp?url=KS.MusicSpecial.asp?Action=Step1',6);">���ר��</a></li>
						 <li><a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'����ϵͳ >> <font color=red>�������</font>','Disabled','KS.Music.asp?url=KS.MusicType.asp');">�������</a> <a href='javascript:void(0)' onClick="SelectObjItem1(this,'����ϵͳ >> <font color=red>���ֹ���</font>','Disabled','KS.Music.asp?url=KS.MusicSinger.asp',6);">���ֹ���</a></li>
						 <li><a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'����ϵͳ >> <font color=red>��˸��</font>','Disabled','KS.Music.asp?url=KS.MusicGeCi.asp');">��˸��</a> <a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'����ϵͳ >> <font color=red>ר�����۹���</font>','Disabled','KS.Music.asp?url=KS.MusicComment.asp');">ר������</a></li>
						 <li><a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'����ϵͳ >> <font color=red>��������������</font>','Disabled','KS.MediaServer.asp?TypeID=1');">��������������</a></li>
					  </DIV>
				    <!--------------���ֹ��� end--------------------> 
					<%End If
					End If
				  End If
					%>
					
					<!--------------��Ƹ��ְ start-------------------->
					 <%
				   IF instr(lcase(KS.C("ModelPower")),"job0")=0 or KS.C("SuperTf")=1 Then
					 If Not ModelXML Is Nothing Then
					  If ModelXML.documentElement.SelectNodes("row[@channelid=10 and @channelstatus=1]").length<>0 Then
					   N=N+1
					  %>
					 <DIV  class="left_menu" id="td_<%=n+1%>" onClick="javascript:switchShow(<%=n+1%>,1);" height=26>&nbsp;&nbsp;<a href="javascript:void(0)">��Ƹ��ְ</a></DIV>
					 <DIV class="hid" id="subTable<%=n+1%>" style="WIDTH: 100%">
					  <%
					     Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & server.urlencode("��ְ��Ƹ >> <font color=red>��Ƹϵͳ����</font>'") & ",'SetParam','KS.JobSetting.asp');"">��Ƹϵͳ����</a></li>"
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & server.urlencode("��ְ��Ƹ >> <font color=red>��ҵְλ����</font>'") & ",'disabled','KS.Jobhy.asp');"">��ҵְλ����</a></li>"
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & server.urlencode("��ְ��Ƹ >> <font color=red>����ģ�����</font>'") & ",'disabled','KS.JobTemplate.asp');"">����ģ�����</a></li>"
						  Response.Write "&nbsp;==================="
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & server.urlencode("��ְ��Ƹ >> <font color=red>��Ƹ��λ����</font>'") & ",'disabled','KS.JobCompany.asp');"">��Ƹ��λ����</a></li>"
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & Server.UrlEncode("��ְ��Ƹ >> <font color=red>�����Ƹ��λ</font>") & "','disabled','KS.JobCompany.asp?ComeFrom=Verify');"">��Ƹ��λ���</a></li>"
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & server.urlencode("��ְ��Ƹ >> <font color=red>�����Ƹ��λ</font>'") & ",'disabled','KS.JobCompany.asp?Action=Add');"">�����Ƹ��λ</a></li>"
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & server.urlencode("��ְ��Ƹ >> <font color=red>��Ƹְλ����</font>'") & ",'disabled','KS.Jobzw.asp');"">��Ƹְλ����</a></li>"
						  Response.Write "&nbsp;==================="
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & server.urlencode("��ְ��Ƹ >> <font color=red>���˼�������</font>'") & ",'disabled','KS.JobResume.asp');"">���˼�������</a></li>"
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & server.urlencode("��ְ��Ƹ >> <font color=red>��˸��˼���</font>'") & ",'disabled','KS.JobResume.asp?ComeFrom=Verify');"">���˼������</a></li>"
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & Server.Urlencode("��ְ��Ƹ >> <font color=red>��Ӹ��˼���</font>'") & ",'disabled','KS.JobResume.asp?Action=Add');"">��Ӹ��˼���</a></li>"
						  Response.Write "<li><a href=""javascript:SelectObjItem1(this,'" & Server.Urlencode("��ְ��Ƹ >> <font color=red>������������</font>'") & ",'disabled','KS.JobEdu.asp');"">������������</a></li>"
					  %>
					 
					 </DIV>
                    <!--------------��Ƹ��ְ end--------------------> 
					<% End If
					End If
				   End If
					%>
					
					
					<!--------------����ϵͳ start-------------------->
				 <%
				IF instr(lcase(KS.C("ModelPower")),"mnkc0")=0 or KS.C("SuperTf")=1 Then
					 If Not ModelXML Is Nothing Then
					  If ModelXML.documentElement.SelectNodes("row[@channelid=9 and @channelstatus=1]").length<>0 Then
					   N=N+1
					   %>
					 <DIV  class="left_menu" id="td_<%=n+1%>" onClick="javascript:switchShow(<%=n+1%>,1);" height=26>&nbsp;&nbsp;<a href="javascript:void(0)">����ϵͳ</a></DIV>
					 <DIV class="hid" id="subTable<%=n+1%>" style="WIDTH: 100%">
					 <%
					      Response.Write "<li><a href='mnkc/mnkc.asp' target='MainFrame'>�Ծ����/���</a></li>"
					      Response.Write "<li><a href='mnkc/mnkc_score.asp' target='MainFrame'>���Գɼ�����</a></li>"
						  Response.Write "==================="
						  Response.Write "<li><a href='mnkc/refreshindex.asp' target='MainFrame'>����Ƶ����ҳ</a></li>"
						  Response.Write "<li><a href='mnkc/mnkc_makesortall.asp?type=all' target='MainFrame'>�������з���</a></li>"
						  Response.Write "<li><a href='mnkc/mnkc_makemnkcall.asp' target='MainFrame'>���������Ծ�ҳ</a></li>"
						  Response.Write "<li><a href='mnkc/RefreshClass.asp' target='MainFrame'>�����ܷ���ҳ</a></li>"
					 %>
					 </DIV>
                    <!--------------����ϵͳ end--------------------> 
				    <%
					 End If
					End If
			   End If
					%>
					
					<!--------------�ʴ�ϵͳ start-------------------->
					<%IF instr(lcase(KS.C("ModelPower")),"ask0")=0 or KS.C("SuperTf")=1 Then%>
					 <%N=N+1%>
					 <DIV  class="left_menu" id="td_<%=n+1%>" onClick="javascript:switchShow(<%=n+1%>,1);" height=26>&nbsp;&nbsp;<a href="javascript:void(0)">�ʴ�ϵͳ</a></DIV>
					 <DIV class="hid" id="subTable<%=n+1%>" style="WIDTH: 100%">
					 <%If KS.ReturnPowerResult(0, "WDXT10000") Then%>
					   <li><a href="javascript:void(0)" onClick="SelectObjItem1(this,'�ʴ�ϵͳ >> <font color=red>�ʴ��������</font>','SetParam','KS.AskSetting.asp');" title="�ʴ��������">�ʴ��������</a></li>
					   <%end if%>
					   <li><a href="javascript:void(0)" onClick="SelectObjItem1(this,'�ʴ�ϵͳ >> <font color=red>�����б����</font>','SetParam','KS.AskList.asp');" title="�����б����">�����б����</a></li>
					   <%If KS.ReturnPowerResult(0, "WDXT10002") Then%>
					   <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'�ʴ�ϵͳ >> <font color=red>�������</font>','Disabled','KS.AskClass.asp');">�������</a>
					   <a href='javascript:void(0)' onClick="SelectObjItem1(this,'�ʴ�ϵͳ >> <font color=red>����ʴ����</font>','GO','KS.AskClass.asp?action=add');">���</a></li>
					   <%end if%>
					   <%If KS.ReturnPowerResult(0, "WDXT10003") Then%>
					   <li><a href="javascript:void(0)" onClick="SelectObjItem1(this,'�ʴ�ϵͳ >> <font color=red>�ȼ�ͷ������</font>','Disabled','KS.AskGrade.asp');" title="�ȼ�ͷ������">�ȼ�ͷ������</a></li>
					   <%end if%>
					   </li>
					 </DIV>
				   <%End If%>
                    <!--------------�ʴ�ϵͳ end--------------------> 
					
					<!--------------�ռ�ϵͳ start-------------------->
				   <%IF instr(lcase(KS.C("ModelPower")),"space0")=0 or KS.C("SuperTf")=1 Then%>
					 <%N=N+1%>
					 <DIV  class="left_menu" id="td_<%=n+1%>" onClick="javascript:switchShow(<%=n+1%>,1);" height=26>&nbsp;&nbsp;<a href="javascript:void(0)">�ռ��Ż�</a></DIV>
					 <DIV class="hid" id="subTable<%=n+1%>" style="WIDTH: 100%">
					    	 <div style=" border:#ff6600 1px dotted;width:125px; height:21px; line-height:21px;margin-right:6px;margin-bottom:2px; margin-top:2px;text-align:left;padding-left:5px; font-size:14px;font-weight:bold; color:#ff6600;"><img src="images/ico_friend.gif">&nbsp;���˿ռ�</div>
						 <%If cbool(KS.ReturnPowerResult(0, "KSMS10000")) Then%>
						<li><a href="javascript:void(0)" onClick="SelectObjItem1(this,'���˿ռ� >> <font color=red>�ռ��������</font>','SetParam','KS.SpaceSetting.asp');" title="�ռ��������">�ռ��������</a></li>
						<%end if%>
						<%If KS.ReturnPowerResult(0, "KSMS10001") Then%>
						<li><a href="KS.Space.asp"  target="MainFrame" title="���пռ����">���пռ����</a></li>
						<li><a href="KS.Space.asp?showtype=1"  target="MainFrame" title="���˿ռ����">���˿ռ����</a></li>
						<%end if%>
						<%If KS.ReturnPowerResult(0, "KSMS10002") Then%>
						<li><a href="KS.Spacelog.asp"  target="MainFrame" title="�ռ���־����">�ռ���־����</a></li>
						<%end if%>
						<%If KS.ReturnPowerResult(0, "KSMS10003") Then%>
						<li><a href="KS.SpaceAlbum.asp"  target="MainFrame" title="�û�������">�û�������</a></li>
						<%end if%>
						<%If KS.ReturnPowerResult(0, "KSMS10004") Then%>
						<li><a href="KS.SpaceTeam.asp"  target="MainFrame" title="�û�Ȧ�ӹ���">�û�Ȧ�ӹ���</a></li>
						<%end if%>
						<%If KS.ReturnPowerResult(0, "KSMS10005") Then%>
						<li><a href="KS.SpaceMessage.asp"  target="MainFrame" title="�û����Թ���">�û����Թ���</a></li>
						<%end if%>
						<%If KS.ReturnPowerResult(0, "KSMS10007") Then%>
						<li><a href="KS.SpaceMusic.asp"  target="MainFrame" title="�û���������">�û���������</a></li>
						<%end if%>
						 <div style=" border:#ff6600 1px dotted;width:125px; height:21px; line-height:21px;margin-left:5px; text-align:left;padding-left:5px; font-size:14px;font-weight:bold; color:#ff6600;"><img src="images/ico_home.gif">&nbsp;��ҵ�ռ�</div>
						<%If KS.ReturnPowerResult(0, "KSMS10008") Then%>
					  <li><a href="KS.EnterPrise.asp"  target="MainFrame" title="��ҵ��Ϣ����">��ҵ�ռ����</a></li>
					 <%end if%>
					 <%If KS.ReturnPowerResult(0, "KSMS10009") Then%>
					  <li><a href="KS.EnterPriseNews.asp"  target="MainFrame" title="��ҵ���Ź���">��ҵ���Ź���</a></li>
					 <%end if%>
					 <%If KS.ReturnPowerResult(0, "KSMS10010") Then%>
					  <li><a href="KS.EnterPrisePro.asp"  target="MainFrame" title="��ҵ��Ʒ����">��ҵ��Ʒ����</a></li>
					 <%end if%>
					 <%If KS.ReturnPowerResult(0, "KSMS10012") Then%>
					  <li><a href="KS.EnterPriseClass.asp"  target="MainFrame" title="��ҵ�������">��ҵ�������</a></li>
					 <%end if%>
					 <%If KS.ReturnPowerResult(0, "KSMS10013") Then%>
					  <li><a href="KS.EnterPriseAD.asp"  target="MainFrame" title="��ҵ������">��ҵ������</a></li>
					 <%end if%>
					 <%If KS.ReturnPowerResult(0, "KSMS10011") Then%>
					  <li><a href="KS.EnterPriseZS.asp"  target="MainFrame" title="����֤�����">����֤�����</a></li>
					 <%end if%>
						
					 </DIV>
					<%End If%>
                    <!--------------�ռ�ϵͳ end--------------------> 

					
						 
					
					</ul>
					
					
					
					
					<ul id="dleft_tab2" style="display:none;">
					   <div class="dt">ϵͳ����</div>
					   <div class="dc">
					<%If KS.ReturnPowerResult(0, "KMST10001") Then%>
					   <li><a href="javascript:void(0)" onClick="SelectObjItem1(this,'ϵͳ���� >> <font color=red>������Ϣ����</font>','SetParam','KS.System.asp');" title="������Ϣ����">������Ϣ����</a></li>
					 <%end if%>
					      
						 <%If Not ModelXML Is Nothing Then
					       If ModelXML.documentElement.SelectNodes("row[@basictype=3 and @channelstatus=1]").length<>0 Then
						 %>
						  <%If KS.ReturnPowerResult(0, "KMST20001") Then %>
						  <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red><%=SQL(3,I)%>��������</font>','SetParam','KS.DownParam.asp?ChannelID=<%=SQL(0,I)%>');">���ز�������</a></li>
						  <%End If%>
						 
						 <%If KS.ReturnPowerResult(0, "KMST20002") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red>���ط���������</font>','Disabled','KS.DownServer.asp?ChannelID=<%=SQL(0,I)%>');">���ط���������</a>
						 <%end if%>
						 
					
						<%
						  End If
						End If
						
						If Not ModelXML Is Nothing Then
					       If ModelXML.documentElement.SelectNodes("row[@channelid=7 and @channelstatus=1]").length<>0 Then
						 %>
						  <%If KS.ReturnPowerResult(0, "KMST20003") Then %>
						 <li><a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red>��������</font>','SetParam','KS.MovieParam.asp?ChannelID=7');">Ӱ�Ӳ�������</a></li>
						  <%End If%>
						
						 <%If KS.ReturnPowerResult(0, "KMST20004") Then %>
						 <li><a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red>Ӱ�ӷ���������</font>','Disabled','KS.MediaServer.asp?TypeID=2&ChannelID=7');">Ӱ�ӷ���������</a></li>
						 <%end if%>
					    <%
						   End If
						End If
						
						If Not ModelXML Is Nothing Then
					       If ModelXML.documentElement.SelectNodes("row[@channelid=8 and @channelstatus=1]").length<>0 Then
						 %>
						 <%If KS.ReturnPowerResult(0, "KMST20005") Then %>
						 <li><a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red>���������͹���</font>','Disabled','KS.SupplyType.asp');">���������͹���</a></li>
						  <%End If%>
					  <%  End If
					   End If
					   %>
					 
					 <%If KS.ReturnPowerResult(0, "KMST10003") Then%>
					   <li><a href="KS.PaymentPlat.asp"  target="MainFrame" title="����֧��ƽ̨����">����֧��ƽ̨����</a></li>
					 <%end if%>
					 <%If KS.ReturnPowerResult(0, "KMST10002") Then%>
					   <li><a href="javascript:void(0)" onClick="SelectObjItem1(this,'ϵͳ���� >> <font color=red>����ϵͳ����</font>','SetParam','KS.API.asp');"  title="����ϵͳ����">APIͨ����������</a></li>
					 <%end if%>
					   </div>
					   
					   
					<%If KS.ReturnPowerResult(0, "KSO10000") Then %>
					  <div class="dt">WAPϵͳ����</div>
					  <div class="dc">
                       <li><a href="#" onClick="SelectObjItem1(this,'WAPϵͳ���� >> <font color=red>WAP������������</font>','SetParam','Wap/KS_System.asp');" title="WAP������������">WAP������������</a></li>
					   <li><a href="#"  onClick="SelectObjItem1(this,'WAPϵͳ���� >> <font color=red>WAP�Զ���ҳ�����</font>','Disabled','Wap/KS.Template.asp');">WAP�Զ���ҳ��</a></li>
					  </div>
					<%end if%>
					   
					   
					   <div class="dt">��������</div>
					   <div class="dc">
						 <%If KS.ReturnPowerResult(0, "KMST10015") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'�������� >> <font color=red>��Դ����</font>','Disabled','KS.Origin.asp');">��Դ����</a></li>
						 <%end if%>
						 <%If KS.ReturnPowerResult(0, "KMST10016") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'�������� >> <font color=red>���߹���</font>','Disabled','KS.Author.asp?ChannelID=0');">���߹���</a> </li>
						 <%end if%>

						 <%If KS.ReturnPowerResult(0, "KMST10017") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'�������� >> <font color=red>ʡ�й���</font>','Disabled','KS.Province.asp');">��������</a> </li>
						 <%end if%>

					  <%If KS.ReturnPowerResult(0, "KMST10004") Then%>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�������� >> <font color=red>���ݹؼ�������</font>','Disabled','KS.InnerLink.asp');">���ݹؼ�������</a></li>
                      <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMST10019") Then%>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�������� >> <font color=red>�����ؼ���ά��</font>','Disabled','KS.KeyWord.asp?issearch=1');">�����ؼ���ά��</a></li>
                      <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMST10020") Then%>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�������� >> <font color=red>��ʱ�������</font>','Disabled','KS.Task.asp?action=manage');">��ʱ�������</a></li>
                      <%end if%>
					  
                       <%If KS.ReturnPowerResult(0, "KMST10014") Then %>
					     <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'<%=SQL(1,I)%> >> <font color=red>ͶƱ��¼����</font>','Disabled','KS.PhotoVote.asp?ChannelID=<%=SQL(0,I)%>');">ͼƬͶƱ��¼����</a>	</li>
					   <%End If%>

					  <%If KS.ReturnPowerResult(0, "KMST10006") Then%>
					   	<li><a href="KS.Log.asp"  target="MainFrame" title="վ���ļ�����">��̨��־����</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMST10007") Then%>
					   <li><a href="KS.Database.asp?Action=BackUp"  target="MainFrame" title="���ݿ�ά��">���ݿ�ά��</a></li>
					   <%end if%>
					   <%If KS.ReturnPowerResult(0, "KMST10008") Then%>
					   <li><a href="KS.DataReplace.asp"  target="MainFrame" title="���ݿ��ֶ��滻">���ݿ��ֶ��滻</a></li>
					   <%end if%>
                       <%If KS.ReturnPowerResult(0, "KMST10018") Then%>
					   <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'�������� >> <font color=red>�ϴ��ļ�����</font>','Disabled','KS.AdminFiles.asp');">�ϴ��ļ�����</a></li>
					   <%end if%>					   
					   <%If KS.ReturnPowerResult(0, "KMST10009") Then%>
					   <li><a href="KS.Database.asp?Action=ExecSql"  target="MainFrame" title="����ִ��SQL���">����ִ��SQL���</a></li>
					   <%end if%>
					   <%If KS.ReturnPowerResult(0, "KMST10011") Then%>
					   <li><a href="KS.System.asp?Action=CopyRight"  target="MainFrame" title="����������̽��">����������̽��</a></li>
					   <%end if%>
					   <%If KS.ReturnPowerResult(0, "KMST10012") Then%>
					   <li><a href="KS.CheckMM.asp"  target="MainFrame" title="���߼��ľ��">���߼��ľ��</a></li>
					   <%end if%>
					   </div>
					   
					</ul>
					
					
					
					
					
					<ul id="dleft_tab3" style="display:none;">
					<%If KS.ReturnPowerResult(0, "KSMS10006") Then %>
					<div class="dt">�Զ����</div>
					  <div class="dc">
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�Զ���� >> <font color=red>����Ŀ����</font>','Disabled','KS.Form.asp');">�Զ��������</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�Զ���� >> <font color=red>��ӱ���Ŀ</font>','GO','KS.Form.asp?action=Add');">��ӱ���Ŀ</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�Զ���� >> <font color=red>����Ŀ���ô���</font>','Disabled','KS.Form.asp?action=total');">����Ŀ���ô���</a></li>
					  </div>
					<%end if%>
					<%If KS.ReturnPowerResult(0, "KSMS20004") Then%>
					<div class="dt">
					С��̳/����
					</div>
					<div class="dc">
					<li><a href="KS.GuestBook.asp?Action=Main"  target="MainFrame" title="��վ���Թ���">��վ���Թ���</a></li>
					<li><a href="KS.GuestBoard.asp"  target="MainFrame" title="����������">����������</a></li>
					</div>
					<%end if%>
					<%If KS.ReturnPowerResult(0, "KSMS20014") Then%>
					<div class="dt">PKϵͳ</div>
					  <div class="dc">
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�۵�PKϵͳ >> <font color=red>PK�������</font>','Disabled','KS.PKZT.asp');">PK�������</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�۵�PKϵͳ >> <font color=red>PK�û��۵����</font>','Disabled','KS.PKGD.asp');">PK�û��۵����</a></li>
					  </div>
					<%end if%>
					
                    <div class="dt">
					����ϵͳ
					</div>
					<div class="dc">
						 <%If KS.ReturnPowerResult(0, "KSMS20010") Then %>
						 <li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'����ϵͳ >> <font color=red>���ֶһ���Ʒ</font>','Disabled','KS.MallScore.asp');">���ֶһ���Ʒ</a></li>
						 <%End If%>
					<%If KS.ReturnPowerResult(0, "KSMS20009") Then %>
					<li><a href='javascript:void(0)' onClick="SelectObjItem1(this,'����ϵͳ >> <font color=red>Digg����</font>','Disabled','KS.DiggList.asp');">�ĵ�Digg����</a></li>
					<%End If%>
					<%If KS.ReturnPowerResult(0, "KSMS20008") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'����ϵͳ >> <font color=red>����ָ������</font>','Disabled','KS.Mood.asp');">����ָ������</a></li>
					<%end if%>
					<%If KS.ReturnPowerResult(0, "KSMS20001") Then%>
					<li><a href="KS.FriendLink.asp"  target="MainFrame" title="�������ӹ���">�������ӹ���</a></li>
					<%end if%>
					<%If KS.ReturnPowerResult(0, "KSMS20002") Then%>
					<li><a href="KS.Announce.asp"  target="MainFrame" title="��վ�������">��վ�������</a></li>
					<%end if%>
					<%If KS.ReturnPowerResult(0, "KSMS20000") Then%>
					<li><a href="KS.FeedBack.asp"  target="MainFrame" title="Ͷ�߼���������">Ͷ�߼���������</a></li>
					<%end if%>
					<%If KS.ReturnPowerResult(0, "KSMS20003") Then%>
					<li><a href="KS.Vote.asp"  target="MainFrame" title="վ�ڵ������">վ�ڵ������</a></li>
					<%end if%>
					
					<%If KS.ReturnPowerResult(0, "KSMS20005") Then%>
					<li><a href="KS.Online.asp"  target="MainFrame" title="վ�����������">վ��� �� ��</a></li>
					<%end if%>
					<%If KS.ReturnPowerResult(0, "KSMS20006") Then%>
					<li><a href="KS.Ads.asp"  target="MainFrame" title="���ϵͳ����">���ϵͳ����</a></li>
					<%end if%>
					<%If KS.ReturnPowerResult(0, "KSMS20007") Then%>
					<li><a href="KS.PromotedPlan.asp"  target="MainFrame" title="�ƹ�ƻ�����">�ƹ�ƻ�����</a></li>
					<%end if%>
					</div>
					<div class="dt">��Ȩ��Ϣ</div>
                       <div class="dc">
					    <li><a href="javascript:void(0)">����:������Ϣ�������޹�˾</a></li>
						<li><a href="http://www.kesion.com" target="_blank">�ٷ�:kesion.com</a></lI>
						<li><a href="javascript:void(0)">�绰:0596-2218051<br />0596-2198252</a></lI>
						<li><a href="javascript:void(0)">��ѯQQ:9537636 41904294</a></lI>
					   </div>
					
					
					</ul>
					
					
					
										
					<ul id="dleft_tab4" style="display:none">
					<div class="dt">ģ�͹���</div>
					 <div class="dc">
					 <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'ģ�͹��� >> <font color=red>ģ�͹�����ҳ</font>','Disabled','KS.Model.asp');">ģ�͹�����ҳ</a></li>
					 <%If KS.ReturnPowerResult(0, "KSMM10000") Then%>
					 <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'ģ�͹��� >> <font color=red>�����ģ��</font>','Go','KS.Model.asp?action=Add');">�����ģ��</a></li>
					 <%end if%>
					 <%If KS.ReturnPowerResult(0, "KSMM10004") Then%>
					 <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'ģ�͹��� >> <font color=red>ģ����Ϣͳ��</font>','Go','KS.Model.asp?action=total');">ģ����Ϣͳ��</a></li>
					 <%end if%>
					 </div>
					 <%If KS.ReturnPowerResult(0, "KSMM10003") Then%>
					<div class="dt">ģ���ֶι���</div>
					 <div class="dc">
					  <%For I=0 To UBound(SQL,2)
					   if SQL(6,I)=1 AND SQL(0,I)<>6 and SQL(0,I)<>9 and SQL(0,I)<>10 Then
					  %>
						 <li><a href="javascript:void(0)" onClick="javascript:SelectObjItem1(this,'ģ�͹��� >> <font color=red>�ֶι���</font>','Disabled','KS.Field.asp?ChannelID=<%=SQL(0,I)%>',<%=SQL(0,I)%>);"><%=SQL(1,I)%>�ֶ�</a></li>					  
					<%
					  End iF
					 Next%>
					</div>
					<%end if%>
					</ul>

                    <ul id="dleft_tab5" style="display:none">
					 <div class="dt">��ǩ����</div>
					 <div class="dc">
					<%
					If KS.ReturnPowerResult(0, "KMTL10001") Then
					  .Write ("<li><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'��ǩ���� >> <font color=red>ϵͳ������ǩ</font>','FunctionLabel','Include/Label_Main.asp?LabelType=0');"">ϵͳ������ǩ</a></li>")
					End If
					If KS.ReturnPowerResult(0, "KMTL10002") Then
					  .Write ("<li><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'��ǩ���� >> <font color=red>�Զ���SQL������ǩ</font>','DiyFunctionLabel','Include/Label_Main.asp?LabelType=5');"">�Զ���SQL������ǩ</a></li>")
					End If
					If KS.ReturnPowerResult(0, "KMTL10003") Then
					  .Write ("<li><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'��ǩ���� >> <font color=red>�Զ��徲̬��ǩ</font>','FreeLabel','Include/Label_Main.asp?LabelType=1');"">�Զ��徲̬��ǩ</a></li>")
					End If
					If KS.ReturnPowerResult(0, "KMTL10010") Then
					  .Write ("<li><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'��ǩ���� >> <font color=red>ͨ��ѭ����ǩ</font>','FreeLabel','Include/Label_Main.asp?LabelType=6');"">ͨ��ѭ���б��ǩ</a></li>")
					End If
					If KS.ReturnPowerResult(0, "KMTL10004") Then
					  .Write ("<li><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'��ǩ���� >> <font color=red>�Զ���JS����</font>','SysJSList','include/JS_Main.asp?JSType=0');"">ϵͳJS����</a></li>")
					End If
					If KS.ReturnPowerResult(0, "KMTL10005") Then
					  .Write ("<li><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'��ǩ���� >> <font color=red>�Զ���JS����</font>','FreeJSList','include/JS_Main.asp?JSType=1');"">�Զ���JS����</a></li>")
					End If
					If KS.ReturnPowerResult(0, "KMSL10008") Then
					  .Write "<li><a href='KS.ClassMenu.asp'  target='MainFrame' title='���ɶ����˵�'>���ɶ����˵�</a></li>"
					end if
					If KS.ReturnPowerResult(0, "KMSL10009") Then
					  .Write "<li><a href='KS.TreeMenu.asp'  target='MainFrame' title='�������β˵�'>�������β˵�</a></li>"
					End If

		              .write "</div>"
					  .write "<div class='dt'>ģ�����</div>"
					  .write "<div class='dc'>"
					If KS.ReturnPowerResult(0, "KMTL10006") Then
						.Write ("<li id='s_1'><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'ģ���ǩ���� >> <font color=red>�Զ���ҳ�����</font>','Disabled','KS.DIYPage.asp');"">�Զ���ҳ�����</a></li>")
				    End If
					If KS.ReturnPowerResult(0, "KMTL10007") Then
						.Write ("<li id='s_1'><a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'ģ���ǩ���� >> <font color=red>����ģ�����</font>','Disabled','KS.Template.asp');"">����ģ�����</a></li>")
					End If
					 %>	
					 </div>
					</ul>
					
					<ul id="dleft_tab6" style="display:none">
					
					  <div class="dt">
					   �û�����					  </div>
					  <div class="dc">
					  <%If KS.ReturnPowerResult(0, "KMUA10001") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�û�ϵͳ >> <font color=red>����Ա����</font>','Disabled','KS.Admin.asp');">����Ա����</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10002") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�û�ϵͳ >> <font color=red>ע���û�����</font>','Disabled','KS.User.asp');" title="ע���û�����">ע���û�����</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�û�ϵͳ >> <font color=red>����û�</font>','Disabled','KS.User.asp?Action=Add');" title="����û�">����û�</a></li>
					  
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10004") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�û�ϵͳ >> <font color=red>�û������</font>','Disabled','KS.UserGroup.asp');" title="�û������">�û������</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10003") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�û�ϵͳ >> <font color=red>�û����Ź���</font>','Disabled','KS.UserMessage.asp');" title="�û����Ź���">�û����Ź���</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10009") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�û�ϵͳ >> <font color=red>�����ʼ�����</font>','Disabled','KS.UserMail.asp');" title="�����ʼ�����">�����ʼ�����</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10012") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�û�ϵͳ >> <font color=red>��Ա�ֶι���</font>','Disabled','KS.Field.asp?ChannelID=101');" title="��Ա�ֶι���">��Ա�ֶι���</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10013") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�û�ϵͳ >> <font color=red>��Ա������</font>','Disabled','KS.UserForm.asp');" title="��Ա������">��Ա������</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10014") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�û�ϵͳ >> <font color=red>��Ա��̬����</font>','Disabled','KS.UserLog.asp');" title="��Ա��̬����">��Ա��̬����</a></li>
					  <%end if%>
					  
					  
					  </div>
					  <div class="dt">
					   ������ϸ����					 
					  </div>
					  <div class="dc">
					  <%If KS.ReturnPowerResult(0, "KMUA10005") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�û�ϵͳ >> <font color=red>��Ա��ȯ��ϸ</font>','Disabled','KS.LogPoint.asp');" title="��Ա��ȯ��ϸ">��Ա��ȯ��ϸ</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10006") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�û�ϵͳ >> <font color=red>��Ա��Ч����ϸ</font>','Disabled','KS.LogEdays.asp');" title="��Ա��Ч����ϸ">��Ա��Ч����ϸ</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10007") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�û�ϵͳ >> <font color=red>��Ա�ʽ���ϸ</font>','Disabled','KS.LogMoney.asp');" title="��Ա�ʽ���ϸ">��Ա�ʽ���ϸ</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�û�ϵͳ >> <font color=red>��Ա������ϸ</font>','Disabled','KS.LogScore.asp');" title="��Ա������ϸ">��Ա������ϸ</a></li>
					  <%end if%>
					  <%If KS.ReturnPowerResult(0, "KMUA10008") Then %>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�û�ϵͳ >> <font color=red>���³�ֵ������</font>','Disabled','KS.Card.asp?cardtype=0');" title="���³�ֵ������">���³�ֵ������</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�û�ϵͳ >> <font color=red>���ϳ�ֵ������</font>','Disabled','KS.Card.asp?cardtype=1');" title="���ϳ�ֵ������">���ϳ�ֵ������</a></li>
					  <%end if%>
					  </div>
					  <%If KS.ReturnPowerResult(0, "KMUA10002") Then %>
					  <div class="dt">
					   ���ٲ����û�					  </div>
					  <div class="dc">
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�û�ϵͳ >> <font color=red>24Сʱ�ڵ�¼</font>','Disabled','KS.User.asp?UserSearch=5');"><font color=#ff6600>24Сʱ�ڵ�¼</a></font></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�û�ϵͳ >> <font color=red>24Сʱ�ڵ�¼</font>','Disabled','KS.User.asp?UserSearch=6');">24Сʱ��ע��</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�û�ϵͳ >> <font color=red>24Сʱ�ڵ�¼</font>','Disabled','KS.User.asp?UserSearch=1');"> ����ס���û�</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�û�ϵͳ >> <font color=red>24Сʱ�ڵ�¼</font>','Disabled','KS.User.asp?UserSearch=3');">��������Ա</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�û�ϵͳ >> <font color=red>24Сʱ�ڵ�¼</font>','Disabled','KS.User.asp?UserSearch=4');">���ʼ���֤</a></li>
					  <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'�û�ϵͳ >> <font color=red>24Сʱ�ڵ�¼</font>','Disabled','KS.User.asp?UserSearch=2');">���й���Ա�û�</a></li>
                      </div>
					<%end if%>
					</ul>
					<ul id="dleft_tab7" style="display:none">
					<%If KS.ReturnPowerResult(0, "KSO10002") Then %>
					  <div class="dt">CC��Ƶ���</div>
					  <div class="dc">
					   <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'CC��Ƶ��� >> <font color=red>��������</font>','Disabled','../plus/CC/cc.asp');">CC��������</a></li>
					  </div>
					<%end if%>
					<%If KS.ReturnPowerResult(0, "KSO10003") Then %>
					  <div class="dt">WSSͳ�Ʋ��</div>
					  <div class="dc">
					   <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'WSS ͳ�Ʋ�� >> <font color=red>WSS ����</font>','Disabled','../plus/wss/wss.asp');">WSS ����</a></li>
					   <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'WSS ͳ�Ʋ�� >> <font color=red>WSS ����</font>','Disabled','../plus/wss/wss.asp?action=show');">�鿴ͳ��</a></li>
					  </div>
					<%end if%>
					
					  <div class="dt">���ݵ�����</div>
					  <div class="dc">
					   <li><a href="javascript:void(0)"  onClick="SelectObjItem1(this,'���ݵ����� >> <font color=red>���������������</font>','Disabled','KS.Import.asp');">���������������</a></li>
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
		var id = 1;  //Ĭ��ѡ�е�ID
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
			  
			  //����Ƿ���������ļ�
			  $.ajax({
			  url: "KS.Update.asp",
			  cache: false,
			  data: "action=check",
			  success: function(d){
			        d=unescape(d);
					switch (d){
					 case 'enabled':
					  $("#updateInfo").html("<font color='green'>�Բ���,��û�п����Զ�������°汾����!</font>");
					  break;
					 case 'false':
					  $("#updateInfo").html("<font color='green'>��ǰ�Ѿ������°汾!</font>");
					  break;
					 case 'localversionerr':
					  $("#updateInfo").html("<font color='green'>���ر���xml�汾�ļ�����,����<%=KS.Setting(89)%>include/version.xml�ļ��Ƿ����!</font>");
					  break;
					 case 'remoteversionerr':
					  $("#updateInfo").html("<font color='green'>��ȡ�������ļ�����,����<%=KS.Setting(89)%>admin_update.asp�ļ��������Ƿ���ȷ���Ժ�����!</font>");
					  break;
					 case 'unallow':
					  $("#updateInfo").html("<font color='green'>ϵͳ��鵽�пɸ����ļ�,����֧����������,�뵽�ٷ�վ(<a href='http://www.kesion.com' target='_blank'>www.kesion.com</a>)���������ļ�!</font>");
					  break;
					 case 'unallowversion':
					  $("#updateInfo").html("<font color='green'>ϵͳ��鵽�пɸ����ļ�,���������İ汾�������°汾�Ų���Ӧ,��֧����������,���������ǰʹ�õİ汾���ٷ�վ(<a href='http://www.kesion.com' target='_blank'>www.kesion.com</a>)���������ļ��ֹ�����!</font>");
					  break;
					 default:
					    $("#updateInfo").html("<font color='red'>ϵͳ��鵽�п������ļ�!</font>");
					  	var str="<div style='height:auto;padding-top:10px;' id='updateResult'><font color=red>��ܰ��ʾ:ϵͳ��⵽�п������Ĳ���,����ǰ���������޸��������ñ���!</font><div style='margin-top:3px'>"+d+"</div><br><div style='text-align:center'><input type='button' value='��������' id='openwin' style='height:25px;background:#efefef;border:1px solid #000'/>&nbsp;<input id='closewin' type='button' value='�ر�ȡ��' name='button' style='height:25px;background:#efefef;border:1px solid #000' /></div></div>"
					    popupTips('ϵͳ��鵽�п������ļ�',str,510,300);
						  $("#closewin").click(function(){
								closeWindow();
								$("#updateInfo").html("<font color='red'>��ȡ���Զ���������!</font>");
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
			  $("#updateInfo").html("<font color='red'>��������,�벻Ҫˢ�±�ҳ��!</font>");
			   $.ajax({
			  url: "KS.Update.asp",
			  cache: false,
			  data: "action=update",
			  success: function(r){
			      r=unescape(r);
				  switch (r){
				    case "remoteversionerr":
					 $("#updateInfo").html("�ٷ����ݻ�ȡʧ��,������ֹ!");
					 alert('�ٷ����ݻ�ȡʧ��,������ֹ!');
					 closeWindow();
					 break;
					default :
					  $("#updateInfo").html("��ϲ,���������ɹ�!");
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
			
			
			<div class="position"><font color=red><%=KS.C("AdminName")%></font> ���ã���ӭ������վ��̨ϵͳ��
			<%
								Dim RS:Set RS = Server.CreateObject("Adodb.Recordset")
								RS.Open "Select top 1 * From KS_Admin Where UserName='" & KS.C("AdminName") & "'", Conn, 1, 1
								If Not RS.EOF Then
								  KS.Echo "��ݣ�"
										  If KS.C("SuperTF") = 0 Then
										   KS.Echo "��ͨ����Ա"
										   Else
										   KS.Echo "��������Ա"
										   End If
									 KS.Echo " ��¼������" & RS("LoginTimes") & "�� ���ε�¼ʱ�䣺" & RS("LastLoginTime")
								End If
								 RS.Close: Set RS=Nothing
			%></div>
			<div class="title"><img src="images/gif-0760.gif">��Ȩ������</div>
			<div class="nr">KesionCMSϵͳ�����ݿ�����Ϣ�������޹�˾(<a href="http://www.kesion.com" target="_blank">Kesion.Com</a>)�����������������Ȩ�ǼǺ�:<a href='http://www.kesion.com/images/v5dj.jpg' target='_blank'>2009SR00339</a>����Ȩ��[<%=KS.Setting(0)%>]ʹ�á��κθ��˻���֯��������Ȩ����������ɾ�����޸ġ����������������������һ�й��ڰ�Ȩ����Ϣ��
			</div>
			
			<div class="title"><img src="images/gif-0760.gif">������Ϣ��</div>
			<div class="nr l">
			 <ul>
			   <li>��ǰ�汾��<%=KS.Version%></li>
			   <li>���°汾��<span id='versioninfo'><img src='images/loading.gif' align='absmiddle'>������...</span></li>
			   <li>��Ʒ���������ݿ�����Ϣ�������޹�˾</li>
			   <li>��ѯ Q Q��9537636 41904294 ��ҵ����֧��QQ��111394 54004407</li>
			   <li>��˾��վ��<a href='http://www.kesion.com/' target='_blank'>kesion.com</a> <a href='http://www.kesion.org/' target='_blank'>kesion.org</a> <a href='http://www.kesion.cn/' target='_blank'>kesion.cn</a></li>
			  </ul>
			  </div>
			 <div class="nr l">
			  <h2>��������</h2>
			  <span id='updateInfo'>���ڼ�����°汾��Ϣ...</span>  
			 </div>
			<div class="box">
			<div class="newbox1">
			<div class="title"><img src="images/gif-0760.gif">�������Ϣ��</div>
			<div class="nr">
			 <%
			 Dim Node,Num,Url,HasVerify
			 HasVerify=false
			 If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig()
			 For Each Node In Application(KS.SiteSN&"_ChannelConfig").DocumentElement.SelectNodes("channel[@ks21=1 and @ks0!=6 and @ks0!=9 and @ks0!=10]")
			   Num=Conn.Execute("Select count(id) from " & Node.SelectSingleNode("@ks2").text & " where verific=0")(0)
			   If Num=0 Then
			   'KS.Echo "��ǩ" & Node.SelectSingleNode("@ks3").text & ":<font color=red>" & Num &" </font>" & Node.SelectSingleNode("@ks4").text & "&nbsp;"
			   Else
			    HasVerify=true
			   KS.Echo "<span style='cursor:pointer;' title='�������ǩ��' onclick=""location.href='KS.ItemInfo.asp?ComeFrom=Verify&ChannelID=" & Node.SelectSingleNode("@ks0").text & "';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&opstr=" & server.URLEncode(Node.SelectSingleNode("@ks1").text & " >> <font color=red>ǩ��" & Node.SelectSingleNode("@ks3").text & "</font>")&"';"">��ǩ" & Node.SelectSingleNode("@ks3").text & "[<font color=red>" & Num &"</font>]" & Node.SelectSingleNode("@ks4").text & "</span>&nbsp;"
			   End If
			 Next
			 If KS.C_S(10,21)="1" Then
				Num=conn.execute("select count(id) from ks_Job_Company where status=0")(0)
				If Num>0 Then
				 HasVerify=true
				 KS.Echo "<span style='cursor:pointer;' title='����������' onclick=""location.href='KS.JobCompany.asp?ComeFrom=Verify';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&opstr=" & server.URLEncode("��Ƹ��ְ���� >> <font color=red>������Ƹ��λ</font>")&"';"">������Ƹ��λ[<font color=red>" & Num & "</font>]��</span>&nbsp;"
				End If
				Num=conn.execute("select count(id) from ks_Job_Resume where status=0")(0)
				If Num>0 Then
				 HasVerify=true
				 KS.Echo "<span style='cursor:pointer;' title='����������' onclick=""location.href='KS.JobResume.asp?ComeFrom=Verify';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&opstr=" & server.URLEncode("��Ƹ��ְ���� >> <font color=red>������Ƹ��λ</font>")&"';"">�������[<font color=red>" & Num & "</font>]��</span>&nbsp;"
				End If
				Num=conn.execute("select count(id) from KS_Job_Edu where status=0")(0)
				If Num>0 Then
				 HasVerify=true
				 KS.Echo "<span style='cursor:pointer;' title='����������' onclick=""location.href='KS.JobEdu.asp?status=0';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&opstr=" & server.URLEncode("��Ƹ��ְ���� >> <font color=red>������Ƹ��λ</font>")&"';"">���������[<font color=red>" & Num & "</font>]��</span>&nbsp;"
				End If
			 End If
			 
			KS.Echo " <div style='height:22px;padding-top:3px;border-top:1px dashed #cccccc'>"
			Num=conn.execute("select count(id) from ks_comment where verific=0")(0)
			If Num>0 Then
			 HasVerify=true
			 KS.Echo "<span style='cursor:pointer;' title='����������' onclick=""location.href='KS.Comment.asp?ComeFrom=Verify';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&opstr=" & server.URLEncode("���ݹ��� >> <font color=red>�������</font>")&"';"">��������[<font color=red>" & Num & "</font>]��</span>"
			End If
			Num=conn.execute("select count(linkid) from ks_link where verific=0")(0)
			If Num>0 Then
			 HasVerify=true
			KS.Echo " <span style='cursor:pointer;' title='����������' onclick=""location.href='KS.FriendLink.asp?Action=Verific';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&opstr=" & server.URLEncode("�������ӹ��� >> <font color=red>�������</font>")&"';"">��������[<font color=red>" & Num & "</font>]��</span>"
		    End If
			Num=conn.execute("select count(blogid) from ks_blog where status=0")(0)
			If Num>0 Then
			HasVerify=true
			KS.Echo " <span style='cursor:pointer;' title='����������' onclick=""location.href='KS.Space.asp?from=verify';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&opstr=" & server.URLEncode("�ռ��Ż����� >> <font color=red>��˿ռ�</font>")&"';"">����ռ�[<font color=red>" & Num & "</font>]��</span>"
			End If
			Num=conn.execute("select count(id) from ks_bloginfo where status=2")(0)
			If Num>0 Then
			HasVerify=true
			KS.Echo " <span style='cursor:pointer;' title='����������' onclick=""location.href='KS.Spacelog.asp?from=verify';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&opstr=" & server.URLEncode("�ռ��Ż����� >> <font color=red>�����־</font>")&"';"">������־[<font color=red>" & Num & "</font>]ƪ</span>"
			End If
			Num=conn.execute("select count(id) from ks_photoxc where status=0")(0)
			If Num>0 Then
			HasVerify=true
			KS.Echo " <span style='cursor:pointer;' title='����������' onclick=""location.href='KS.SpaceAlbum.asp?from=verify';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&opstr=" & server.URLEncode("�ռ��Ż����� >> <font color=red>������</font>")&"';"">�������[<font color=red>" & Num & "</font>]��</span>"
			End If
			Num=conn.execute("select count(id) from ks_team where Verific=0")(0)
			If Num>0 Then
			HasVerify=true
			KS.Echo " <span style='cursor:pointer;' title='����������' onclick=""location.href='KS.SpaceTeam.asp?from=verify';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&opstr=" & server.URLEncode("�ռ��Ż����� >> <font color=red>������</font>")&"';"">����Ȧ��[<font color=red>" & Num & "</font>]��</span>"
			End If
			Num=conn.execute("select count(id) from KS_EnterpriseNews where status=0")(0)
			If Num>0 Then
			HasVerify=true
			KS.Echo " <span style='cursor:pointer;' title='����������' onclick=""location.href='KS.EnterPriseNews.asp?from=verify';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&opstr=" & server.URLEncode("�ռ��Ż����� >> <font color=red>�����ҵ����</font>")&"';"">������ҵ����[<font color=red>" & Num & "</font>]ƪ</span>"
			End If
			Num=conn.execute("select count(id) from KS_EnterPriseAD where status=0")(0)
			If Num>0 Then
			HasVerify=true
			KS.Echo " <span style='cursor:pointer;' title='����������' onclick=""location.href='KS.EnterPriseAD.asp?from=verify';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&opstr=" & server.URLEncode("�ռ��Ż����� >> <font color=red>�����ҵ���</font>")&"';"">������ҵ���[<font color=red>" & Num & "</font>]��</span>"
			End If
			'Num=conn.execute("select count(id) from KS_EnterPriseZS where status=0")(0)
			'If Num>0 Then
			'HasVerify=true
			'KS.Echo " <span style='cursor:pointer;' title='����������' onclick=""location.href='KS.EnterPriseZS.asp?from=verify';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&opstr=" & server.URLEncode("�ռ��Ż����� >> <font color=red>���֤��</font>")&"';"">����֤��[<font color=red>" & Num & "</font>]��</span>"
			'End If
			
			
			
			
			If HasVerify=false Then
			 KS.Echo "<div style='margin:30px;text-align:center;color:red'>����û���û��ύ����˵���Ϣ��</div>"
			End If
			KS.Echo "</div>"
			   %>
			</div>
			</div>
			<div class="newbox2">
			<div class="title"><img src="images/gif-0760.gif">������̳������</div>
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
				.Write "<div id='co' align=""center"" onClick=""ChangeLeftFrameStatu();"" title=""ȫ��/����"" style=""cursor:pointer;""><font color=red>��</font> �ر�����</div>"
				.Write "<div id='footmenu'>����ͨ��=>��"
				If KS.ReturnPowerResult(0, "KMTL20000") Then
				.Write "<a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'�������� >> <font color=red>����������ҳ</font>','disabled','Include/refreshindex.asp');"">������ҳ</a>"
				End If
				.Write "<a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'�������� >> <font color=red>����������ҳ</font>','disabled','Include/RefreshHtml.asp?ChannelID=1');"">��������</a>"
				
				If KS.ReturnPowerResult(0, "KMTL10007") Then
				.Write "<a href='javascript:void(0)' onClick=""javascript:SelectObjItem1(this,'ģ���ǩ���� >> <font color=red>ģ�����</font>','disabled','KS.Template.asp');"">ģ�����</a>"
				End If
				If KS.ReturnPowerResult(0, "KMST10001") Then
				.Write "<a href='javascript:void(0)' onClick=""SelectObjItem1(this,'ϵͳ���� >> <font color=red>������Ϣ����</font>','SetParam','KS.System.asp');"" title='������Ϣ����'>������Ϣ����</a>"
				End If
				If Instr(KS.C("ModelPower"),"model1")>0 Or KS.C("SuperTF")="1" then
				.Write "<a href='javascript:void(0)' onClick=""SelectObjItem1(this,'ģ�͹��� >> <font color=red>ģ�͹�����ҳ</font>','SetParam','KS.Model.asp');"">ģ�͹���</a>"
				End If
				If KS.ReturnPowerResult(0, "KMUA10011") Then
				.Write "<a href='javascript:void(0)' onClick=""SelectObjItem1(this,'�û����� >> <font color=red>������Ա��������</font>','SetParam','KS.UserProgress.asp');"">�鿴��������</a>"
			    End If
				.Write "</div>"
				.Write "<div id='footcopyright'>��Ȩ���� &copy; 2006-2010 ������Ϣ�������޹�˾</div>"
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
				.Write "            self.co.innerHTML = ""�� ������"""
				.Write "        }"
				.Write "        else if(screen==true)"
				.Write "        {"
				.Write "            parent.FrameMain.cols='201,*';"
				.Write "           screen=false;"
				.Write "            self.co.innerHTML = ""<font color=red>��</font> �ر�����"""
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
