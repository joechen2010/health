<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Admin_Article
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Article
        Private KS,ComeUrl,KSCls
		'=====================================声明本页面全局变量==============================================================
        Private ID, I, totalPut, Page, RS,ComeFrom
		Private KeyWord, SearchType, StartDate, EndDate,SearchParam, MaxPerPage,T, TitleStr, VerificStr
		Private TypeStr, AttributeStr, FolderID, TemplateID,WapTemplateID,FolderName, Action
		Private UserDefineFieldArr,UserDefineFieldValueStr
		Private NewsID, TitleType, Title,Fulltitle,ShowComment, TitleFontColor, TitleFontType, PicNews, ArticleContent, PhotoUrl, Changes, Recommend,IsTop,PageTitle,IsSign,SignUser,SignDateLimit,SignDateEnd,Province,City
		Private Strip, Popular, Verific, Comment, Slide,ChangesUrl, Rolls, KeyWords, Author, Origin, AddDate, Rank,  Hits, HitsByDay, HitsByWeek, HitsByMonth, SpecialID,CurrPath,UpPowerFlag,Intro
		Private Inputer,FileName,SqlStr,Errmsg,Makehtml,Tid,Fname,KSRObj,SaveFilePath
		Private ReadPoint,ChargeType,PitchTime,ReadTimes,InfoPurview,arrGroupID,DividePercent
		Private ChannelID,F_B_Arr,F_V_Arr
		'======================================================================================================================

		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub

		
		Public Sub Kesion()
		
		ChannelID=KS.ChkClng(KS.G("ChannelID"))
		Action     = KS.G("Action")
		If Action="SelectUser" Then
		   Call SelectUser()
		   Exit Sub
		End If
		
		Session("ThumbnailsConfig")=KS.C_S(ChannelID,46)
		F_B_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(0),"|")
		F_V_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(1),"|")
		
		
		KeyWord    = KS.G("KeyWord")
		SearchType = KS.G("SearchType")
		StartDate  = KS.G("StartDate")
		EndDate    = KS.G("EndDate")
		ComeFrom   = KS.G("ComeFrom")
		SearchParam = "ChannelID=" & ChannelID
		If KeyWord<>"" Then SearchParam=SearchParam & "&KeyWord=" & KeyWord
		If SearchType<>"" Then  SearchParam=SearchParam & "&SearchType=" & SearchType
		If StartDate<>"" Then SearchParam=SearchParam & "&StartDate=" & StartDate 
		If EndDate<>"" Then SearchParam=SearchParam & "&EndDate=" & EndDate
		If KS.S("Status")<>"" Then SearchParam=SearchParam & "&Status=" & KS.S("Status")
		If ComeFrom<>"" Then SearchParam=SearchParam & "&ComeFrom=" & ComeFrom

		ComeUrl=Cstr(Request.ServerVariables("HTTP_REFERER"))
		Page = KS.ChkClng(KS.G("page"))
		Select Case Action
		 Case "CheckTitle"
		    Call CheckTitle()    
		 Case Else
		    Call CreateArticle()
		End Select
		
	 End Sub
				
		'添加文章
		Sub CreateArticle()
			Page = KS.G("page")
			Action = KS.G("Action") 'Add添加新文章 Edit编辑文章 Verify 审核前台投搞
			
			IF KS.G("Method")="Save" Then
			 Call DoSave()
			Else 
			 Call ArticleAdd()
			End If

		End Sub
		
		Sub ArticleAdd()
			With Response
			.Write "<html>"
			.Write "<head>"
			.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
			.Write "<title>添加</title>"
			.Write "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
			.Write "<script src=""images/pannel/tabpane.js"" language=""JavaScript""></script>" & vbCrlf
			.Write "<link href=""images/pannel/tabpane.CSS"" rel=""stylesheet"" type=""text/css"">" & vbCrlf
		    .Write "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>" & vbCrLf
			.Write "<script language=""javascript"" src=""../KS_Inc/popcalendar.js""></script>" & vbCrlf
		    .Write "<script language=""JavaScript"" src=""../KS_Inc/Jquery.js""></script>" & vbCrLf
			CurrPath = KS.GetUpFilesDir
						
			Set RS = Server.CreateObject("ADODB.RecordSet")
			If Action = "Add" Then
			  FolderID = Trim(KS.G("FolderID"))
			  If Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10002") Then          '检查是否有添加文章的权限
			   .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=ViewFolder&FolderID=" & FolderID & "';</script>")
			   Call KS.ReturnErr(2, "KS.ItemInfo.asp?ChannelID=" & ChannelID & "&Page=" & Page & "&ID=" & FolderID)
			   Exit Sub
			  End If
			  Hits = 0:HitsByDay = 0: HitsByWeek = 0:HitsByMonth = 0:Comment = 1 :IsTop=0:Verific=1:UserDefineFieldValueStr=0
			  ReadPoint=0:PitchTime=24:ReadTimes=10: IsSign=0 : SignDateLimit=0 : SignDateEnd=Now
			  KeyWords = Session("keywords")
			  Author = Session("Author")
			  Origin = Session("Origin")
			ElseIf Action = "Edit"  Or Action="Verify" Then
			   Set RS = Server.CreateObject("ADODB.RECORDSET")
			   RS.Open "Select top 1 * From " & KS.C_S(ChannelID,2) &" Where ID=" & KS.G("ID") , conn, 1, 1
			   If RS.EOF And RS.BOF Then	Call KS.Alert("参数传递出错!", ComeUrl):Exit Sub
				FolderID = Trim(RS("Tid"))
				
				If Action = "Edit" And Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10003") Then     '检查是否有编辑文章的权限
					RS.Close:Set RS = Nothing
					 If KeyWord = "" Then
					  .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=ViewFolder&FolderID=" & FolderID & "&channelid=" & channelid & "';</script>")
					  Call KS.ReturnErr(2, "KS.ItemInfo.asp?ChannelID=" & ChannelID &"&Page=" & Page & "&ID=" & FolderID)
					 Else
					  .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=" & server.URLEncode(KS.C_S(ChannelID,1) & " >> <font color=red>搜索" & KS.C_S(ChannelID,3) & "结果</font>")&"&ButtonSymbol=ArticleSearch';</script>")
					  Call KS.ReturnErr(1, "KS.ItemInfo.asp?ChannelID=" & ChannelID &"&Page=" & Page & "&KeyWord=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate)
					 End If
					 Exit Sub
			   End If
			   If Action="Verify" And Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10012") Then     '检查是否有审核前台会员投稿文章的权限
					  .Write ("<script>$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=ViewFolder&FolderID=" & FolderID & "&channelid=" & channelid & "';</script>")
					  Call KS.ReturnErr(2, "KS.ItemInfo.asp?ChannelID=" & ChannelID &"&Page=" & Page & "&ID=" & FolderID)
			   End If
			   
				TitleType      = Trim(RS("TitleType"))
				Title          = Trim(RS("title"))
				Fulltitle      = Trim(RS("Fulltitle"))
				TitleFontColor = Trim(RS("TitleFontColor"))
				TitleFontType  = Trim(RS("TitleFontType"))
				PhotoUrl         = Trim(RS("PhotoUrl"))
				PicNews        = CInt(RS("PicNews"))
				Rolls          = CInt(RS("Rolls"))
				Changes        = CInt(RS("Changes"))
				Recommend      = CInt(RS("Recommend"))
				Strip          = CInt(RS("Strip"))
				Popular        = CInt(RS("Popular"))
				Verific        = CInt(RS("Verific"))
				IsTop          = Cint(RS("IsTop"))
				Comment        = CInt(RS("Comment"))
				Slide          = CInt(RS("Slide"))
				AddDate        = CDate(RS("AddDate"))
				Rank           = Trim(RS("Rank"))
				FileName       = RS("Fname")

				TemplateID = RS("TemplateID")
				WapTemplateID=RS("WapTemplateID")
				Hits = Trim(RS("Hits"))
				HitsByDay = Trim(RS("HitsByDay"))
				HitsByWeek = Trim(RS("HitsByWeek"))
				HitsByMonth = Trim(RS("HitsByMonth"))
				KeyWords = Trim(RS("KeyWords"))
				Author = Trim(RS("Author"))
				Origin = Trim(RS("Origin"))
				Intro  = RS("Intro")
				IsSign = RS("IsSign")
				SignUser=RS("SignUser")
				SignDateLimit=RS("SignDateLimit")
				SignDateEnd=RS("SignDateEnd")
				Province=RS("Province")
				City=RS("City")

				ReadPoint =RS("ReadPoint")
				ChargeType=RS("ChargeType")
				PitchTime =RS("PitchTime")
				ReadTimes =RS("ReadTimes")
				InfoPurview=RS("InfoPurview")
				arrGroupID =RS("arrGroupID")
				DividePercent=RS("DividePercent")
			   If CInt(Changes) = 1 Then
				ChangesUrl = Trim(RS("ArticleContent"))
			   Else
				ArticleContent = Trim(RS("ArticleContent"))
			   End If
			   If KS.IsNul(ArticleContent) Then ArticleContent="&nbsp;"
			    PageTitle = RS("PageTitle")
				FolderID = RS("Tid")
				'自定义字段
				UserDefineFieldArr=KSCls.Get_KS_D_F_Arr(ChannelID)
				If IsArray(UserDefineFieldArr) Then
				For I=0 To Ubound(UserDefineFieldArr,2)
				  Dim UnitOption
				  If UserDefineFieldArr(12,I)="1" Then
				   UnitOption="@" & RS(UserDefineFieldArr(0,I)&"_Unit")
				  Else
				   UnitOption=""
				  End If
				  If I=0 Then
				    UserDefineFieldValueStr=RS(UserDefineFieldArr(0,I))& UnitOption & "||||"
				  Else
				    UserDefineFieldValueStr=UserDefineFieldValueStr  & RS(UserDefineFieldArr(0,I))& UnitOption & "||||"
				  End If
				Next
			  End If
				RS.Close
			End If
			If IsNULL(PageTitle) Then PageTitle=""
			'取得上传权限
			UpPowerFlag = KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10009")
						 
			%>
			<script language='JavaScript'>
			function ResumeError()
			{return true;}
			window.onerror = ResumeError;
			$(document).ready(function(){
				$(parent.frames["BottomFrame"].document).find("#Button1").attr("disabled",false);
				$(parent.frames["BottomFrame"].document).find("#Button2").attr("disabled",false);
			 <%If F_B_Arr(4)=1 Then%>
				   if ($("#Changes").attr('checked'))
				   ChangesNews();
			 <%End If%>
			 <%If F_B_Arr(5)=1 Then%>
			  $('#KeyLinkByTitle').click(function(){
			    GetKeyTags();
			  });
			 <%End If%>
			  
			});
			function GetKeyTags()
			{
			  var text=escape($('input[name=Title]').val());
			  if (text!=''){
				  $('#KeyWords').val('请稍等,系统正在自动获取tags...');
				  $("#KeyWords").attr("disabled",true);
				  $.get("../plus/ajaxs.asp", { action: "GetTags", text: text,maxlen: 20 },
				  function(data){
					$('#KeyWords').val(unescape(data));
					$('#KeyWords').attr("disabled",false);
				  });
			  }else{
			   alert('对不起,请先输入内容!');
			  }
			}
			
			function ChangesNews()
			{ 
			 if ($("#Changes").attr('checked'))
			  {
			  <%If F_B_Arr(9)=1 Then%>
			  $("#ContentArea").hide();
			  <%end if%>
			  $("#ChangesUrl").attr("disabled",false);
			  }
			  else
			   {
			   <%If F_B_Arr(9)=1 Then%>
			   $("#ContentArea").show();
			   <%end if%>
			  $("#ChangesUrl").attr("disabled",true);
			   }
			}
			function SelectAll(){
			  $("#SpecialID>option").each(function(){
			    $(this).attr("selected",true);
			  });
			}
			function UnSelectAll(){
			  $("#SpecialID>option").each(function(){
			    $(this).attr("selected",false);
			  });
			}
			function GetFileNameArea(f)
			{
			   $('#filearea').toggle(f);
			}
			function GetTemplateArea(f)
			{
			   $('#templatearea').toggle(f);
			}
			function insertHTMLToEditor(codeStr) 
			{   
				oEditor=FCKeditorAPI.GetInstance("Content");
				if(oEditor   &&   oEditor.EditorWindow){ 
					oEditor.InsertHtml(codeStr); 
				} 
			} 
			function InsertFileFromUp(FileList,InstallDir)
			{  
				Files=FileList.split("|");
				for(var i=0;i<Files.length-1;i++)
				{     var ext=getFilePic(Files[i]);
					  var files=Files[i].split('/');
					  var file=files[files.length-1];
					  var br='';
					  if (i!=Files.length-1) br='<br />';
					  var fileext = Files[i].substring(Files[i].lastIndexOf(".") + 1, Files[i].length).toLowerCase();
                      if (fileext=="gif" || fileext=="jpg" || fileext=="jpeg" || fileext=="bmp" || fileext=="png")
					  {
					   insertHTMLToEditor('<img src="'+Files[i]+'" border="0"/><br/>');	
					  }
					  else
					  {
					  var str="<img border=0 src="+InstallDir+"KS_Editor/images/FileIcon/"+ext+"> <a href='"+Files[i]+"'  target='_blank'>[点击浏览该文件:"+file+"]</a>"+br;
					  insertHTMLToEditor(str);	
					  }
				 }
			}
			function insertPage()
			{
			 <%if KS.C_S(ChannelID,34)=0 Then%>
			 ArticleContent.InsertHTMLStr("[NextPage]");
			 <%Else%>
			 insertHTMLToEditor("[NextPage]");
			 <%End If%>
			}
			
			function SubmitFun()
			{  
			   if ($("input[name=title]").val()==""){
					alert("请输入<%=KS.C_S(ChannelID,3)%>标题！");
					$("input[name=title]").focus();
					return false;}
			   if ($("#tid>option[selected=true]").val()=='0')
			   {
			       alert('请选择所属栏目!');
				   return false;
			   }
			<%If F_B_Arr(5)=1 Then%> 
			  if ($("input[name=KeyWords]").val().length>255){
			    alert('关键字不能超过255个字符!');
				$("input[name=KeyWords]").focus();
				return false;}
			<%End If%>
			
			  <%If F_B_Arr(9)=1 and KS.C_S(ChannelID,34)=0 Then%>
				if (frames["ArticleContent"].CurrMode!='EDIT') {alert('其他模式下无法保存，请切换到设计模式');return false;}
					<% IF ChannelID=1 Then%>
					frames["ArticleContent"].SaveCurrPage();
					var TempContentArray=frames["ArticleContent"].ArticleContentArray;
					$("#Content").val('');
					for (var i=0;i<TempContentArray.length;i++)
					{   
						if (TempContentArray[i]!='')
						{
							if ($("#Content").val()=='') $("#Content").val(TempContentArray[i]);
							else $("#Content").val($("#Content").val()+'[NextPage]'+TempContentArray[i]);
						} 
					}
					<%else%>
					$("#Content").val(frames["ArticleContent"].KS_EditArea.document.body.innerHTML);
					<%end if%>
				   $("#Content").val(frames["ArticleContent"].ReplaceImgToScript(frames["ArticleContent"].Resumeblank(document.myform.Content.value)));
				<%end if%>
				
				
				
				<%If F_B_Arr(4)=1 and F_B_Arr(9)=1 Then%>
				   <%if KS.C_S(ChannelID,34)=0 Then%>
					if (($("#Changes").attr('checked')==false)&&$("#Content").val()=="")
					  {　
					    frames["ArticleContent"].KS_EditArea.focus();
						alert("<%=KS.C_S(ChannelID,3)%>内容不能留空！");
						return false;
					  }
					<%else%>
					  if (($("#Changes").attr('checked')==false)&&(FCKeditorAPI.GetInstance('Content').GetXHTML(true)==""))
					 {
					  alert("<%=KS.C_S(ChannelID,3)%>内容不能留空！");
					  FCKeditorAPI.GetInstance('Content').Focus();
					  return false;
					 }
					<%end if%>
				<%end if%>
				<%If F_B_Arr(4)=1 Then%>
				if (($("#Changes").attr('checked'))&&($("input[name=ChangesUrl]").val()==""))
				  { $("#ChangesUrl").focus();
					alert("请输入外部链接的Url！");
					return false;
				  }
				 
				<%end if%>
				<%If F_B_Arr(9)=1 Then%>
				if ($("input[name=BeyondSavePic]").attr('checked')==true)
				 {
				  $('#LayerPrompt').show();
				  window.setInterval('ShowPromptMessage()',150)
				 }
				 <%end if%>
				  
				  $('#myform').submit();
				  $(parent.frames["BottomFrame"].document).find("#Button1").attr("disabled",true);
				  $(parent.frames["BottomFrame"].document).find("#Button2").attr("disabled",true);
			}
			var ForwardShow=true;
			function ShowPromptMessage()
			{
				var TempStr=ShowArticleArea.innerText;
				if (ForwardShow==true)
				{
					if (TempStr.length>4) ForwardShow=false;
					ShowArticleArea.innerText=TempStr+'.';
					
				}
				else
				{
					if (TempStr.length==1) ForwardShow=true;
					ShowArticleArea.innerText=TempStr.substr(0,TempStr.length-1);
				}
			}
			
			
			var SaveBeyondInfo=''
					   +'<div id="LayerPrompt" style="position:absolute; z-index:1; left: 200px; top: 200px; background-color: #f1efd9; layer-background-color: #f1efd9; border: 1px none #000000; width: 360px; height: 63px; display: none;"> '
					   +'<table width="100%" height="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#FF0000">'
					   +'<tr> '
					   +'<td align="center">'
					   +'<table width="80%" border="0" cellspacing="0" cellpadding="0">'
					   +'<tr>'
					   +' <td width="75%" nowrap>'
					   +'<div align="right">请稍候，系统正在保存远程图片到本地</div></td>'
					   +'   <td width="25%"><font id="ShowArticleArea">&nbsp;</font></td>'
					   +' </tr>'
					   +'</table>'
					   +'</td>'
					   +'</tr>'
					   +'</table>'
					   +'</div>'
			document.write (SaveBeyondInfo)
			function SelectUser(){
				var arr=showModalDialog('KS.Article.asp?action=SelectUser&DefaultValue='+document.myform.SignUser.value,'','dialogWidth:600px; dialogHeight:450px; help: no; scroll: yes; status: no');
				if (arr != null){
					document.myform.SignUser.value=arr;
				}
			}
			</script>
			<%
			.Write "</head>"
			.Write "<body leftmargin='0' topmargin='0' marginwidth='0' onkeydown='if (event.keyCode==83 && event.ctrlKey) SubmitFun();' marginheight='0'>"
			.Write "<div align='center'>"
			.Write "<ul id='menu_top'>"
			.Write "<li class='parent' onclick=""return(SubmitFun())""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/save.gif' border='0' align='absmiddle'>确定保存</span></li>"
			.Write "<li class='parent' onclick=""history.back();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/back.gif' border='0' align='absmiddle'>取消返回</span></li>"
		    .Write "</ul>"
			
			.Write "<div class=tab-page id=ArticlePane>"
			.Write " <SCRIPT type=text/javascript>"
			.Write "   var tabPane1 = new WebFXTabPane( document.getElementById( ""ArticlePane"" ), 1 )"
			.Write " </SCRIPT>"
				 
			.Write " <div class=tab-page id=basic-page>"
			.Write "  <H2 class=tab>基本信息</H2>"
			.Write "	<SCRIPT type=text/javascript>"
			.Write "				 tabPane1.addTabPage( document.getElementById( ""basic-page"" ) );"
			.Write "	</SCRIPT>"
			
			
			
			.Write " <TABLE width='100%' BORDER='0' cellpadding='1'  cellspacing='1' class='ctable'>"
			.Write "    <form action='?ChannelID=" & ChannelID & "&Method=Save' method='post' id='myform' name='myform'>"
			.Write "      <input type='hidden' value='" & KS.G("ID") & "' name='NewsID'>"
			.Write "      <input type='hidden' value='" & Action & "' name='Action'>"
			.Write "      <input type='hidden' name='Page' value='" & Page & "'>"
			.Write "      <input type='hidden' name='KeyWord' value='" & KeyWord & "'>"
			.Write "      <input type='hidden' name='SearchType' value='" & SearchType & "'>"
			.Write "      <Input type='hidden' name='StartDate' value='" & StartDate & "'>"
			.Write "      <input type='hidden' name='EndDate' value='" & EndDate & "'>"
			.Write "      <input type='hidden' name='ArticleID' value='" & KS.G("ID") & "'>"
			.Write "      <input type='hidden' name='Inputer' value='" &Inputer & "'>"
			
			.Write "      <TR class='tdbg'>"
			.Write "         <td height='20' width='90' class='clefttitle'><div align='right'><font color='#FF0000'><strong>" & F_V_Arr(0) & ":</strong></font></div></td>"
			.Write "          <td>"
			If F_B_Arr(19)=1 Then
			.Write "<select name='TitleType' id='TitleType' class='textbox'>"
			.Write "                    <option></option>"
			
			 Dim TitleTypeXml:Set TitleTypeXml=LFCls.GetXMLFromFile("TitleType")
			 If IsObject(TitleTypeXml) Then
				 Dim objNode,i,j,objAtr
				 Set objNode=TitleTypeXml.documentElement 
				 For i=0 to objNode.ChildNodes.length-1 
						set objAtr=objNode.ChildNodes.item(i)
						If Trim(TitleType) = Trim(objAtr.Attributes.item(0).Text) Then 
						.Write "<option selected style='color:" &objAtr.Attributes.item(1).Text & "'>" & objAtr.Attributes.item(0).Text & "</option>"
						Else
						.Write "<option style='color:" &objAtr.Attributes.item(1).Text & "'>" & objAtr.Attributes.item(0).Text & "</option>"
						End If
				 Next
			End If		
	
			
			.Write "                  </select>"
		   End If
			.Write "                  <input name='title' id='title' type='text'  style='background:url(Images/rule.gif);border:1px solid #999999; height:20px' value='" & Title & "' maxlength='160' size=40>"
			.Write "                  <font color='#FF0000'>*</font>"
			
			If F_B_Arr(19)=1 Then
			.Write "                <select name='TitleFontType' id='TitleFontType'>"
			
			.Write "                    <option value=''>字形</option>"
			If TitleFontType = "1" Then
			  .Write "                    <option value='1' selected>粗体</option>"
			Else
			  .Write "                    <option value='1'>粗体</option>"
			End If
			If TitleFontType = "2" Then
			  .Write "                    <option value='2' selected>斜体</option>"
			Else
			  .Write "                    <option value='2'>斜体</option>"
			End If
			If TitleFontType = "3" Then
			  .Write "                    <option value='3' selected>粗+斜</option>"
			Else
			  .Write "                    <option value='3'>粗+斜</option>"
			End If
			If TitleFontType = "0" Then
			  .Write "                    <option value='0' selected>规则</option>"
			Else
			  .Write "                    <option value='0'>规则</option>"
			End If
			
			.Write "                  </select><input type='hidden' name='TitleFontColor' value='" & TitleFontColor &"'>"
			Dim ColorImg:If TitleFontColor="" Then ColorImg="RectNoColor.gif" Else ColorImg="rect.gif"
			.Write " <img border=0 id=""MarkFontColorShow"" src=""images/" & ColorImg & """ style=""cursor:pointer;background-Color:" & TitleFontColor & ";"" onClick=""Getcolor(this,'../ks_editor/selectcolor.asp','TitleFontColor');this.src='images/rect.gif';"" title=""选取颜色"">&nbsp;"
			End If
			
			.Write "<input class='button' type='button' value='重名检测' onclick=""if(document.all.title.value==''){alert('请输入" & KS.C_S(ChannelID,3) & "标题!');}else OpenWindow('?ChannelID=" & ChannelID & "&Action=CheckTitle&title='+document.all.title.value,280,290,window);"">"
			
			If F_B_Arr(20)=1 Then
			.Write "<input type='checkbox' name='MakeHtml' value='1' checked>" & F_V_Arr(20)
			End IF

			.Write "</td>"
			.Write "   </tr>"
			
			If F_B_Arr(2)=1 Then
			.Write "              <tr  class='tdbg'>"
			.Write "                <td height='22' class='clefttitle'> <div align='right'><strong>" & F_V_Arr(2) & ":</strong></div></td>"
			.Write "                <td> <input name='Fulltitle' type='text' maxlength='200' id='Fulltitle' size='80' value='" & Fulltitle & "' class='textbox'></td>"
			.Write "              </tr>"
			End IF

			.Write "              <tr class='tdbg'>"
			.Write "                <td class='clefttitle' align='right'><strong>" & F_V_Arr(1) & ":</strong></td>"
			.Write "                <td><input type='hidden' name='OldClassID' value='" & FolderID & "'>"
			.Write " <select size='1' name='tid' id='tid' style=""width:150px"">"
			.Write " <option value='0'>--请选择栏目--</option>"
			.Write Replace(KS.LoadClassOption(ChannelID),"value='" & FolderID & "'","value='" & FolderID &"' selected") & " </select>"

		
		If F_B_Arr(3)=1 Then
			.Write "&nbsp;" & F_V_Arr(3) & "<input name='Recommend' type='checkbox' id='Recommend' value='1'"
			If Recommend = 1 Then .Write (" Checked")
			.Write ">"
			.Write "                                推荐"
			.Write "                                <input name='Rolls' type='checkbox' id='Rolls' value='1'"
			If Rolls = 1 Then .Write (" Checked")
			.Write ">"
			.Write "                                滚动"
			.Write "                                <input name='Strip' type='checkbox' id='Strip' value='1'"
			If Strip = 1 Then .Write (" Checked")
			.Write ">"
			.Write "                                头条"
			.Write "                                <input name='Popular' type='checkbox' id='Popular' value='1'"
			If Popular = 1 Then .Write (" Checked")
			.Write ">"
			.Write "                                热门"
			.Write "                                <input name='IsTop' type='checkbox' id='IsTop' value='1'"
			If IsTop = 1 Then .Write (" Checked")
			.Write ">"
			.Write "                                固顶"
			.Write "                                <input name='Comment' type='checkbox' id='Comment' value='1'"
			If Comment = 1 Then .Write (" Checked")			
			.Write ">"
			.Write "                                允许评论"
			.Write "                                <input name='Slide' type='checkbox' id='Slide' value='1'"
			If Slide = 1 Then
			.Write (" Checked")
			End If
			.Write ">"
			.Write "                                幻灯"
		End If
			.Write "             </td></tr>"

          If F_B_Arr(4)=1 Then
			.Write "             <tr class='tdbg' id='ContentLink'>"
			.Write "                      <td class='clefttitle'><div align='right'><strong>" & F_V_Arr(4) & ":</strong></div></td>"
			.Write "                      <td>"
			
			If ChangesUrl = "" Then
			 .Write ("<input name='ChangesUrl' type='text' id='ChangesUrl' disabled value='http://' size='60' class='textbox'>")
			Else
			 .Write ("<input name='ChangesUrl' type='text' id='ChangesUrl' value='" & ChangesUrl & "' size='60' class='textbox'>")
			End If
			
			If Changes = 1 Then
			 .Write ("<input name='Changes' type='checkbox' Checked id='Changes' value='1' onclick='ChangesNews()'><font color='#FF0000'>使用转向链接</font>")
			Else
			 .Write ("<input name='Changes' type='checkbox' id='Changes' value='1' onclick='ChangesNews()'><font color='#FF0000'> 使用转向链接</font>")
			End If
			.Write " </td>"
			.Write "              </tr>"
		  End If
		  If F_B_Arr(5)=1 Then
			.Write "              <tr class='tdbg'>"
			.Write "                <td class='clefttitle'><div align='right'><strong>" & F_V_Arr(5) & ":</strong></div></td>"
			.Write "                <td height='50'> <input name='KeyWords' type='text' id='KeyWords' class='textbox' value='" & KeyWords & "' size=40> <="
			.Write "                  <select name='SelKeyWords' style='width:150px' onChange='InsertKeyWords(document.getElementById(""KeyWords""),this.options[this.selectedIndex].value)'>"
		    .Write "<option value="""" selected> </option><option value=""Clean"" style=""color:red"">清空</option>"
			.Write KSCls.Get_O_F_D("KS_KeyWords","KeyText","IsSearch=0 Order BY AddDate Desc")
			.Write "                  </select>"
			.Write " <br />【<a href=""#"" id=""KeyLinkByTitle"" style=""color:green"">根据" & F_V_Arr(0) & "自动获取Tags</a>】<input type='checkbox' name='tagstf' value='1' checked>写入Tags表  <span class='help' onclick=""mousepopup(event,'什么是Tag','Tag(标签)是一种更为灵活、有趣的分类方式，您可以为每篇文章添加一个或多个Tag(标签)，你可以看到网站上所有和您使用了相同Tag的内容，由此和他人产生更多的联系。Tag体现了群体的力量，使得内容之间的相关性和用户之间的交互性大大增强。多个Tag请用英文逗号隔开',300)"">帮助</span>"
			.Write "                </td>"
			.Write "              </tr>"
		 End If
		 If F_B_Arr(6)=1 Then
			.Write "              <tr class='tdbg'>"
			.Write "                <td class='clefttitle'><div align='right'><strong>" & F_V_Arr(6) & ":</strong></div></td>"
			.Write "                <td> <input name='author' type='text' id='author' value='" & Author & "' size=30 class='textbox'>                 <=【<font color='blue'><font color='#993300' onclick='$(""#author"").val(""未知"")' style='cursor:pointer;'>未知</font></font>】【<font color='blue'><font color='#993300' onclick=""$('#author').val('佚名')"" style='cursor:pointer;'>佚名</font></font>】【<font color='blue'><font color='red' onclick=""$('#author').val('" & KS.C("AdminName") & "')"" style='cursor:pointer;'>" & KS.C("AdminName") & "</font></font>】"
							 If Author <> "" And Author <> "未知" And Author <> KS.C("AdminName") And Author <> "佚名" Then
							  .Write ("【<font color='blue'><font color='#993300' onclick=""$('#author').va('" & Author & "')"" style='cursor:pointer;'>" & Author & "</font></font>】")
							 End If
							  .Write ("<select name='SelAuthor' style='width:100px' onChange=""$('#author').val(this.options[this.selectedIndex].value)"">")
		    .Write "<option value="""" selected> </option><option value="""" style=""color:red"">清空</option>"
			.Write KSCls.Get_O_F_D("KS_Origin","OriginName","ChannelID=0 and OriginType=1 Order BY AddDate Desc")
			.Write "                                   </select> &nbsp; </td>"
			.Write "              </tr>"
		  End If
		  If F_B_Arr(7)=1 Then
			.Write "              <tr class='tdbg'>"
			.Write "                <td class='clefttitle'><div align='right'><strong>" & F_V_Arr(7) & ":</strong></div></td>"
			.Write "                <td> <input name='Origin' type='text' id='Origin' value='" & Origin & "' size=30 class='textbox'>                 <=【<font color='blue'><font color='#993300' onclick=""$('#Origin').val('不详');"" style='cursor:pointer;'>不详</font></font>】【<font color='blue'><font color='#993300' onclick=""$('#Origin').val('本站原创')"" style='cursor:pointer;'>本站原创</font></font>】【<font color='blue'><font color='#993300' onclick=""$('#Origin').val('互联网')"" style='cursor:pointer;'>互联网</font></font>】"
							  If Origin <> "" And Origin <> "不详" And Origin <> "本站原创" And Origin <> "互联网" Then
							  .Write ("【<font color='blue'><font color='#993300' onclick=""$('#Origin').val('" & Origin & "')"" style='cursor:pointer;'>" & Origin & "</font></font>】 ")
							   End If
							  .Write ("<select name='selOrigin' style='width:100px' onChange=""$('#Origin').val(this.options[this.selectedIndex].value)"">")
		    .Write "<option value="""" selected> </option><option value="""" style=""color:red"">清空</option>"
			.Write KSCls.Get_O_F_D("KS_Origin","OriginName","OriginType=0 Order BY AddDate Desc")
			.Write "                </select> </td>"
			.Write "              </tr>"
		 End If
		 
		 If F_B_Arr(23)="1" Then
			.Write "              <tr class='tdbg'>"
			.Write "                <td class='clefttitle'><div align='right'><strong>" & F_V_Arr(23) & ":</strong></div></td>"
			.Write "                <td> <script src=""../plus/area.asp"" type=""text/javascript""></script>  <font color='#999999'>指定文档的来源地或是指定具体的分站新闻</font>"
			.Write "                </td>"
			.Write "              </tr>"
			.Write "<script type='text/javascript'>"
			if Province<>"" then
			  .Write "$('#Province').val('" & province & "');"
			end if
			if City<>"" Then
			  .Write "$('#City')[0].options[1]=new Option('" & City & "','" & City & "');"
			  .Write "$('#City')[0].options(1).selected=true;"
			end if
			.Write "</script>"
		 End If
		 
        '自定义字段
		.Write               KSCls.Get_KS_D_F(ChannelID,UserDefineFieldValueStr)

		 
		 If F_B_Arr(8)=1 Then
			.Write "              <tr  class='tdbg' style='height:25px'>"
			.Write "                <td class='clefttitle'><div align='right'><strong>" & F_V_Arr(8) & ":</strong></div><input name='AutoIntro' type='checkbox' checked value='1'>自动截取内<br>容的200个字作<br>为导读 <span class='help' onclick=""mousepopup(event,'如何设置导读','如果您留空不填,系统将截取内容的前200个字作为导读',200)"">帮助</span></td>"
			.Write "                <td><textarea class='textbox' name=""Intro"" style='width:98%;height:80'>" & Intro & "</textarea>"
			.Write "              </td>"
			.Write "      </TR>"
         End If
		 If F_B_Arr(9)=1 Then
            .Write "      <TR class='tdbg'ID='ContentArea'>"
			.Write "            <td class='clefttitle' width='90' valign='top'><br/><div align='right'><strong>" & F_V_Arr(9) & ":</strong><br><input name='BeyondSavePic' type='checkbox' value='1'>自动下载内容里的图片<br/><br/><input type='button' onclick=""insertPage()"" class='button' value='插入分页'></div><br><b><font color=red>过滤字符设置</font></b><br><label><input type='checkbox' name='FilterIframe' value='1'>Iframe</label><br/><label><input type='checkbox' name='FilterObject' value='1'>Object</label><br/><label><input type='checkbox' name='FilterScript' value='1'>Script</label><br/><label><input type='checkbox' name='FilterDiv' value='1'>Div</label><br/><label><input type='checkbox' name='FilterClass' value='1'>Class</label><br/><label><input type='checkbox' name='FilterTable' value='1'>Table</label><br/><label><input type='checkbox' name='FilterSpan' value='1'>Span</label><br/><label><input type='checkbox' name='FilterImg' value='1'>IMG</label><br/><label><input type='checkbox' name='FilterFont' value='1'>Font</label><br/><label><input type='checkbox' name='FilterA' value='1'>A链接</label><br/><label><input type='checkbox' name='FilterHtml' value='1' onclick=""alert('所有HTML格式将被清除！');"">HTML</label><br/><label><input type='checkbox' name='FilterTd' value='1'>TD</label><br/></td><TD valign='top' nowrap height='100%'>"
			If F_B_Arr(21)=1 and CBool(UpPowerFlag) = True Then
			.Write "<table border='0' width='100%' cellspacing='0' cellpadding='0'>"
			.Write "<tr><td height='30' width=70>&nbsp;<strong>" & F_V_Arr(21) & ":</strong></td><td><iframe id='UpFileFrame' name='UpFileFrame' src='KS.UpFileForm.asp?UPType=File&ChannelID=" & ChannelID &"' frameborder=0 scrolling=no width='100%' height='100%'></iframe></td></tr>"
			.Write "</table>"
		   end if

			 
			.Write "<textarea ID='Content' name='Content' style='display:none'>" & Server.HTMLEncode(ArticleContent) & "</textarea>"
			If KS.C_S(ChannelID,34)=0 Then
			.Write "<iframe id='ArticleContent' name='ArticleContent' src='KS.Editor.asp?ID=Content&style="&channelid&"&ChannelID=" & channelid & "' frameborder=0 scrolling=no width='100%' height='350'></iframe>"
			Else
			.Write "<iframe id=""content___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=Content&amp;Toolbar=NewsTool"" width=""98%"" height=""400"" frameborder=""0"" scrolling=""no""></iframe>"
            End If

			.Write "<table border='0' width='100%' cellspacing='0' cellpadding='0'>"
			.Write "<tr><td height='30' colspan='3'>&nbsp;<strong>分页方式: </strong><select onchange=""if (this.value==2){$('#pagearea').show()}else{$('#pagearea').hide()}"" name='PaginationType'><option value='0'>不分页</option><option value='1' selected>手工分页</option><option value=2>自动分页</option></select>&nbsp;&nbsp;<strong>注：</strong><font color=blue>手工分页符标记为<font color=red>“[NextPage]”</font>，注意大小写</font> </td></tr>"
			.Write "<tr id='pagearea' style='display:none'><td colspan=3>&nbsp;自动分页时的每页大约字符数<input type='text' name='MaxCharPerPage' value='" & KS.Setting(9) & "' size=6 class='textbox'> <font color=blue>必须大于100才能生效</font></td></tr>"          
 
			.Write "<tr><td align='center' width=60><strong>&nbsp;分页标题: </strong><br><font color=red>可留空</font></td><td align='left'> <textarea name=""PageTitle"" style='width:500px;height:60px;line-height:26px;padding-left:20px;background:url(images/Rule1.gif) no-repeat 0 0px;border:1px solid #999999;' ID=""PageTitle"">" & Replace(PageTitle,"§",vbcrlf) & "</textarea></td><td align='left'><font color=green>一行一个标题</font>  <span class='help' onclick=""mousepopup(event,'分页标题作用','这里是为了设置" & KS.C_S(ChannelID,3) & "有多页时,每页想指定不同的标题名称而设计的。',300)"">帮助</span></td></tr>"
			.Write "</table>"
			
			.Write "</TD>"
			.Write "      </TR>"
		 End IF
		 If F_B_Arr(10)=1 Then
			.Write "              <tr class='tdbg' id=trpic style='width:25px;'>"
			.Write "                <td height='22' class='clefttitle'><div align='right'><strong>" & F_V_Arr(10) & ":</strong></div></td>"
			.Write "                <td height='22'><input name='PhotoUrl' type='text' id='PhotoUrl' size='50' value='" & PhotoUrl & "' class='textbox'>"
			.Write "                  <input class=""button""  type='button' name='Submit' value='选择图片...' onClick=""OpenThenSetValue('Include/SelectPic.asp?ChannelID=" & ChannelID &"&CurrPath=" & CurrPath & "',550,290,window,$('#PhotoUrl')[0]);"">  <input class=""button"" type='button' name='Submit' value='抓取图片...' onClick=""OpenThenSetValue('Include/Frame.asp?FileName=SaveBeyondfile.asp&PageTitle='+escape('抓取远程图片')+'&CurrPath=" & CurrPath & "',300,100,window,$('#PhotoUrl')[0]);if ($('#PhotoUrl').val()!='' && $('#ieditor').attr('checked')==true){"
			if KS.C_S(ChannelID,34)=0 Then
			.Write "ArticleContent.InsertHTMLStr('<img src='+$('#PhotoUrl').val()+' />');"
			Else
			.Write "insertHTMLToEditor('<img src='+$('#PhotoUrl').val()+' />');"
			End If
			.Write "}"">"
			
			.Write " <input type='checkbox' name='ieditor' id='ieditor' value='1' checked>同时插入编辑器               </td>"
			.Write "              </tr>"
		End IF
		If F_B_Arr(11)=1 Then
			If CBool(UpPowerFlag) = True Then
			.Write "              <tr  class='tdbg' style='height:25px'>"
			.Write "                <td height='25' class='clefttitle'><div align=right><strong>" & F_V_Arr(11) & ":</strong></div></td>"
			.Write "                <td height='25' align='left'><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='KS.UpFileForm.asp?UPType=Pic&ChannelID=" & ChannelID &"' frameborder=0 scrolling=no width='100%' height='100%'></iframe>"
			.Write "              </td>"
			.Write "              </tr>"
			End If
	    End If
		.Write "</table>"
		.Write "</div>"
			
		If F_B_Arr(17)=1 Then
		   .Write " <div class=tab-page id=option-page>"
		   .Write "  <H2 class=tab>" & F_V_Arr(17) & "</H2>"
		   .Write "	<SCRIPT type=text/javascript>"
		   .Write "				 tabPane1.addTabPage( document.getElementById( ""option-page"" ) );"
		   .Write "	</SCRIPT>"

            .Write "<TABLE style='margin:1px' width='100%' BORDER='0' cellpadding='1'  cellspacing='1' class='ctable'>"
			.Write "           <tr class='tdbg'>"
			.Write "              <td class='clefttitle' align='right'><strong>所属专题:</strong></td>"
			.Write "              <td>"
	        Call KSCls.Get_KS_Admin_Special(ChannelID,KS.ChkClng(KS.G("ID")))
			.write "</td>"
			.Write "           </tr>"
	    If F_B_Arr(12)=1 Then
			.Write "              <tr class='tdbg'>"
			.Write "                <td class='clefttitle'><div align='right'><strong>" & F_V_Arr(12) & ":</strong></div></td>"
			.Write "                <td>"
			If Action <> "Edit" Then
			.Write ("<input name='AddDate' type='text' onclick=""popUpCalendar(this, this, dateFormat,-1,-1)"" id='AddDate' value='" & Now() & "' size='50'  class='textbox'>")
			Else
			.Write ("<input name='AddDate' type='text' onclick=""popUpCalendar(this, this, dateFormat,-1,-1)"" id='AddDate' value='" & AddDate & "' size='50'  readonly class='textbox'>")
			End If
			.Write "                  <b><a href='#' onClick=""popUpCalendar(this, $('input[name=AddDate]').get(0), dateFormat,-1,-1)""><img src='Images/date.gif'  border='0' align='absmiddle' title='选择日期'></a>日期格式：年-月-日 时：分：秒"
			.Write "               </td>"
			.Write "             </tr>"
	  End If
	  If F_B_Arr(13)=1 Then
			.Write "              <tr class='tdbg'>"
			.Write "                <td class='clefttitle'><div align='right'><strong>" & F_V_Arr(13) & ":</strong></td>"
			.Write "                <td><select name='rank'>"
			If Rank = "★" Then
			.Write "                    <option  selected>★</option>"
			Else
			.Write "                    <option>★</option>"
			End If
			If Rank = "★★" Then
			.Write "                    <option  selected>★★</option>"
			Else
			.Write "                    <option>★★</option>"
			End If
			If Rank = "★★★" Or Action = "Add" Then
			.Write "                    <option  selected>★★★</option>"
			Else
			.Write "                    <option>★★★</option>"
			End If
			If Rank = "★★★★" Then
			.Write "                    <option  selected>★★★★</option>"
			Else
			.Write "                    <option>★★★★</option>"
			End If
			If Rank = "★★★★★" Then
			.Write "                    <option  selected>★★★★★</option>"
			Else
			.Write "                    <option>★★★★★</option>"
			End If
			.Write "                  </select>&nbsp;请为" & KS.C_S(ChannelID,3) & "评定阅读等级"
			.Write "               </td>"
			.Write "             </tr>"
	   End If
	   If F_B_Arr(14)=1 Then
			.Write "              <tr class='tdbg'>"
			.Write "                <td class='clefttitle'><div align='right'><strong>" & F_V_Arr(14) & ":</strong></td><td>本日：<input name='HitsByDay' type='text' id='HitsByDay' value='" & HitsByDay & "' size='10' class='textbox'> 本周：<input name='HitsByWeek' type='text' id='HitsByWeek' value='" & HitsByWeek & "' size='10' class='textbox'> 本月：<input name='HitsByMonth' type='text' id='HitsByMonth' value='" & HitsByMonth & "' size='10' class='textbox'> 总计：<input name='Hits' type='text' id='Hits' value='" & Hits & "' size='10' class='textbox'>&nbsp;初数点击数作弊" 
			.Write "                  </td>"
			.Write "              </tr>"
	  End If
	  If F_B_Arr(15)=1 Then
			 .Write "             <tr class='tdbg'>"
			 .Write "               <td class='clefttitle'><div align='right'><strong>" & F_V_Arr(15) & ":</strong></div></td>"
			.Write "                <td> "
			IF Action <> "Edit" and  Action<>"Verify" Then
			.Write " <input type='radio' name='templateflag' onclick='GetTemplateArea(false);' value='2' checked>继承栏目设定<input type='radio' onclick='GetTemplateArea(true);' name='templateflag' value='1'>自定义"
			.Write "<div id='templatearea' style='display:none'>"
			If KS.WSetting(0)="1" Then .Write "<strong>WEB模板</strong> "
			.Write "<input id='TemplateID' name='TemplateID' readonly size=30 class='textbox' value='" & TemplateID & "'>&nbsp;" & KSCls.Get_KS_T_C("$('#TemplateID')[0]") 
			If KS.WSetting(0)="1" Then 
			.Write "<br/><strong>WAP模板</strong> "
			.Write "<input id='WapTemplateID' name='WapTemplateID' readonly size=30 class='textbox' value='" & WapTemplateID & "'>&nbsp;" & KSCls.Get_KS_T_C("$('#WapTemplateID')[0]") 
			End If
			.Write "</div>"
			Else
			
			.Write "<div id='templatearea'>"
			If KS.WSetting(0)="1" Then .Write "<strong>WEB模板</strong> "
			.Write "<input id='TemplateID' name='TemplateID' readonly maxlength='255' size=30 class='textbox' value='" & TemplateID & "'>&nbsp;" & KSCls.Get_KS_T_C("$('#TemplateID')[0]")
			If KS.WSetting(0)="1" Then 
			.Write "<br/><strong>WAP模板</strong> "
			.Write "<input id='WapTemplateID' name='WapTemplateID' readonly size=30 class='textbox' value='" & WapTemplateID & "'>&nbsp;" & KSCls.Get_KS_T_C("$('#WapTemplateID')[0]") 
			End If
			.Write "</div>"
			End If
			.Write "                </td>"
			.Write "             </tr>"
	  End If
	  If F_B_Arr(16)=1 Then
			.Write "             <tr class='tdbg'>"
			.Write "               <td class='clefttitle'><div align='right'><strong>" & F_V_Arr(16) & ":</strong></td><td>"
			IF Action = "Edit" or Action="Verify" Then
			.Write "<input name='FileName' type='text' id='FileName' readonly  value='" & FileName & "' size='25' class='textbox'> <font color=red>不能改</font>"
			Else
			.Write "<input type='radio' value='0' name='filetype' onclick='GetFileNameArea(false);' checked>自动生成 <input type='radio' value='1' name='filetype' onclick='GetFileNameArea(true);' >自定义"
			.Write "<div id='filearea' style='display:none'><input name='FileName' type='text' id='FileName'   value='" & FileName  & "' size='25' class='textbox'> <font color=red>可带路径,如 help.html,news/news_1.shtml等</font></div>"
			End IF
			 .Write "                  </td>"
			 .Write "             </tr>"
     End If
	     If ChannelID=1 Then
	 		.Write "              <tr class='tdbg'>"
			.Write "                <td class='clefttitle'><div align='right'><strong>审核状态:</strong></td><td><input name='verific' type='radio' value='0'"
			if verific=0 then .write " checked"
			.write ">待审核"
			
			.write "<input type='radio' name='verific' value='1'"
			if verific=1 then .write "checked"
			.write ">已审核"
			
			.Write "                  </td>"
			.Write "              </tr>"
         End If
			
			.Write "</table>"
			.Write "</div>"
	   End If
	   
	   			If F_B_Arr(22)=1 Then
			   .Write " <div class=tab-page id=sign-page>"
			   .Write "  <H2 class=tab>" & F_V_Arr(22) & "</H2>"
			   .Write "	<SCRIPT type=text/javascript>"
			   .Write "				 tabPane1.addTabPage( document.getElementById( ""sign-page"" ) );"
			   .Write "	</SCRIPT>"
	
				.Write "<TABLE style='margin:1px' width='100%' BORDER='0' cellpadding='1'  cellspacing='1' class='ctable'>"
				.Write "           <tr class='tdbg'>"
				.Write "              <td class='clefttitle' width='100' align='right'><strong>是否需要签收:</strong></td>"
				.Write "              <td>" 
				.Write "<label><input type='radio' name='issign' onclick=""$('#signs').hide();"" value='0'"
				If IsSign="0" Then .Write " checked"
				.Write">不需要</label>"
				.Write "<label><input type='radio' name='issign' onclick=""$('#signs').show();"" value='1'"
				If IsSign="1" Then .Write " checked"
				.Write ">需要</label>"
				
				.Write " </td></tr>"
				If IsSign="0" Then
				.Write "           <tbody style='display:none' id='signs'>"
				Else
				.Write "           <tbody  id='signs'>"
				End If
				.Write "            <tr class='tdbg'>"
				.Write "              <td class='clefttitle' width='100' align='right'><strong>签收用户:</strong></td>"
				.Write "              <td><textarea name='SignUser' id='SignUser' cols=50 rows=5>" & SignUser & "</textarea>"
				.Write "<br/><input type='button' value='选择用户' onclick='SelectUser()' class='button'> <input type='button' value='清除用户' onclick=""$('#SignUser').val('')"" class='button'>"
				.Write "</td></tr>"
				.Write "  <tr class='tdbg'>"
				.Write "              <td class='clefttitle' width='100' align='right'><strong>时间限制:</strong></td>"
				.Write "              <td>"
				.Write "<label><input type='radio' name='SignDateLimit' onclick=""$('#signdate').hide();"" value='0'"
				If SignDateLimit="0" Then .Write " checked"
				.Write ">不启用</label>"
				.Write "<label><input type='radio' name='SignDateLimit' onclick=""$('#signdate').show();"" value='1'"
				If SignDateLimit="1" Then .Write " checked"
				.Write">启用</label>"
				.Write "</td></tr>"
				If SignDateLimit="1" then
				.Write "  <tr class='tdbg' id='signdate'>"
				else
				.Write "  <tr class='tdbg' id='signdate' style='display:none'>"
				end if
				.Write "              <td class='clefttitle' width='100' align='right'><strong>签收结束时间:</strong></td>"
				.Write "              <td><input type='text' onclick=""popUpCalendar(this, this, dateFormat,-1,-1)"" id='SignDateEnd'  name='SignDateEnd' value='" & SignDateEnd & "'> <a href='#' onClick=""popUpCalendar(this, $('input[name=SignDateEnd]').get(0), dateFormat,-1,-1)""><img src='Images/date.gif'  border='0' align='absmiddle' title='选择日期'></a> <font color=blue>签收用户必须在这个时间内完成签收。格式：期格式：年-月-日 时：分：秒</font></td></tr>"
				
				
				.Write "        </tbody>"
				
				.Write "</table>"
				.Write "</div>"
			End If   

	   
	   
	   
	   
	        If F_B_Arr(18)=1 Then
	           KSCls.LoadChargeOption ChannelID,ChargeType,InfoPurview,arrGroupID,ReadPoint,PitchTime,ReadTimes,DividePercent
		    End If
		
		       KSCls.LoadRelativeOption ChannelID,KS.ChkClng(KS.G("ID"))

			 .Write "      </form></div>"
			 .Write "</body>"
			 .Write "</html>"
			End With
		End Sub
		'保存
		Sub DoSave()
			 Dim SelectInfoList,HasInRelativeID
			 With Response
				TitleType      = KS.G("TitleType")
				Title          = KS.G("Title")
				Fulltitle      = KS.G("Fulltitle")
				TitleFontColor = KS.G("TitleFontColor")
				TitleFontType  = KS.G("TitleFontType")
                ArticleContent = Request.Form("Content")
				If KS.IsNul(ArticleContent)="" Then ArticleContent="&nbsp;"
				ArticleContent = FilterScript(ArticleContent)
				PageTitle      = Replace(Request.Form("PageTitle"),vbcrlf,"§")
				Hits        = KS.ChkClng(KS.G("Hits"))
				HitsByDay   = KS.ChkClng(KS.G("HitsByDay"))
				HitsByWeek  = KS.ChkClng(KS.G("HitsByWeek"))
				HitsByMonth = KS.ChkClng(KS.G("HitsByMonth"))

				PhotoUrl      = KS.G("PhotoUrl")
				If PhotoUrl<>"" Then PicNews=1 Else PicNews=0
				Changes     = KS.ChkClng(KS.G("Changes"))
				Recommend   = KS.ChkClng(KS.G("Recommend"))
				Rolls       = KS.ChkClng(KS.G("Rolls"))
				Strip       = KS.ChkClng(KS.G("Strip"))
				Popular     = KS.ChkClng(KS.G("Popular"))
				Slide       = KS.ChkClng(KS.G("Slide"))
				Comment     = KS.ChkClng(KS.G("Comment"))
				IsTop       = KS.ChkClng(KS.G("IsTop"))
				Makehtml    = KS.ChkClng(KS.G("Makehtml"))
				ChangesUrl  = KS.G("ChangesUrl")
				SpecialID   = Replace(KS.G("SpecialID")," ",""):SpecialID = Split(SpecialID,",")
				SelectInfoList = Replace(KS.G("SelectInfoList")," ","")
				
				Tid         = KS.G("Tid")
				KeyWords    = KS.G("KeyWords")
				Author      = KS.G("Author")
				Origin      = KS.G("Origin")
				AddDate     = KS.G("AddDate")
				If Not IsDate(AddDate) Then AddDate=Now
				Rank        = KS.G("Rank")
				Intro       = KS.G("Intro")
				if Intro="" And KS.ChkClng(KS.G("AutoIntro"))=1 Then Intro=KS.GotTopic(KS.LoseHtml(ArticleContent),200)
				ArticleContent=KS.FilterIllegalChar(ArticleContent)
				Intro=KS.FilterIllegalChar(Intro)
				Title=KS.FilterIllegalChar(Title)
				
				IsSign=KS.ChkClng(KS.G("IsSign"))
				SignUser=Replace(KS.G("SignUser")," ","")
				SignDateLimit=KS.ChkClng(KS.G("SignDateLimit"))
				SignDateEnd=KS.S("SignDateEnd")
				If Not IsDate(SignDateEnd) Then SignDateEnd=Now
				Province=KS.S("Province")
				City=KS.S("City")
				
				
				'收费选项
				ReadPoint   = KS.ChkClng(KS.G("ReadPoint"))
				ChargeType  = KS.ChkClng(KS.G("ChargeType"))
				PitchTime   = KS.ChkClng(KS.G("PitchTime"))
				ReadTimes   = KS.ChkClng(KS.G("ReadTimes"))
				InfoPurview = KS.ChkClng(KS.G("InfoPurview"))
				arrGroupID  = KS.G("GroupID")
				DividePercent=KS.G("DividePercent"):IF Not IsNumeric(DividePercent) Then DividePercent=0
			
				TemplateID  = KS.G("TemplateID")
				WapTemplateID=KS.G("WapTemplateID")
				Dim filetype:filetype=KS.ChkClng(KS.G("filetype"))
				Dim FnameType
				Dim RS_C:Set RS_C=Server.CreateObject("Adodb.RecordSet")
					RS_C.Open "Select * From KS_Class Where ID='" & Tid & "'",conn,1,1
					If Not RS_C.Eof Then
					    FnameType=RS_C("FnameType")
						If KS.ChkClng(KS.G("TemplateFlag"))=2 Or TemplateID="" Then TemplateID=RS_C("TemplateID"):WapTemplateID=RS_C("WapTemplateID")
						If FileType=0 Then
						  If Action = "Add" OR Action="Verify" Then
						   Fname=KS.GetFileName(RS_C("FsoType"), Now, "") & FnameType
						   End If
						End If
					End If
				RS_C.Close:Set RS_C=Nothing
				If filetype=1 Then Fname=KS.G("FileName")
				
				UserDefineFieldArr=KSCls.Get_KS_D_F_Arr(ChannelID)
				If IsArray(UserDefineFieldArr) Then
				For I=0 To Ubound(UserDefineFieldArr,2)
				 If UserDefineFieldArr(6,I)=1 And KS.G(UserDefineFieldArr(0,I))="" Then ErrMsg = ErrMsg & UserDefineFieldArr(1,I) & "必须填写!\n"
				 If UserDefineFieldArr(3,I)=4 And Not Isnumeric(KS.G(UserDefineFieldArr(0,I))) Then ErrMsg = ErrMsg& UserDefineFieldArr(1,I) & "必须填写数字!\n"
				 If UserDefineFieldArr(3,I)=5 And Not IsDate(KS.G(UserDefineFieldArr(0,I))) and UserDefineFieldArr(6,I)=1 Then ErrMsg = ErrMsg& UserDefineFieldArr(1,I) & "必须填写正确的日期!\n" 
				 If UserDefineFieldArr(3,I)=8 And Not KS.IsValidEmail(KS.G(UserDefineFieldArr(0,I))) and UserDefineFieldArr(6,I)=1 Then ErrMsg = ErrMsg& UserDefineFieldArr(1,I) & "必须填写正确的Email格式!\n" 

				Next
				End If

				If Changes = 1 Then	ArticleContent = ChangesUrl
				If Title = "" Then KS.die ("<script>alert('" & KS.C_S(ChannelID,3) & "标题不能为空!');history.back(-1);</script>")
				If CInt(Changes) = 1 Then
				 If ChangesUrl = "" Then KS.die ("<script>alert('请输入" & KS.C_S(ChannelID,3) & "的链接地址！');history.back(-1);</script>")
				End If
				If F_B_Arr(9)=1 Then
				 If ArticleContent = "" and CInt(Changes)<>1 Then .Write ("<script>alert('" & KS.C_S(ChannelID,3) & "内容不能为空!');history.back(-1);</script>").end
				End If
				
				Set RS = Server.CreateObject("ADODB.RecordSet")
				If Tid = "" Then ErrMsg = ErrMsg & "[" & KS.C_S(ChannelID,3) & "类别]必选! \n"
				If Title = "" Then ErrMsg = ErrMsg & "[" & KS.C_S(ChannelID,3) & "标题]不能为空! \n"
				'If Title <> "" And Tid <> "" And (Action = "Add") Then
				'  SqlStr = "select * from " & KS.C_S(ChannelID,2) &" where Title='" & Title & "' And Tid='" & Tid & "'"
				'   RS.Open SqlStr, conn, 1, 1
				'	If Not RS.EOF Then ErrMsg = ErrMsg & "该类别已存在此篇" & KS.C_S(ChannelID,3) & "! \n"
				'   RS.Close
				'End If
				If ErrMsg <> "" Then
				   .Write ("<script>alert('" & ErrMsg & "');history.back(-1);</script>")
				   .End
				Else
				        If KS.ChkClng(KS.G("TagsTF"))=1 Then Call KSCls.AddKeyTags(KeyWords)
						
						If KS.ChkClng(KS.G("PaginationType"))=2 Then
						 ArticleContent=KS.AutoSplitPage(Request.Form("Content"),"[NextPage]",KS.ChkClng(KS.G("MaxCharPerPage")))
						ElseIf KS.ChkClng(KS.G("PaginationType"))=0 Then
						 ArticleContent=Replace(ArticleContent,"[NextPage]","")
						End If
						If KS.ChkClng(KS.G("BeyondSavePic")) = 1 And CInt(Changes) <> 1 Then
							  SaveFilePath = KS.GetUpFilesDir & "/"
							  KS.CreateListFolder (SaveFilePath)
							  ArticleContent = KS.ReplaceBeyondUrl(ArticleContent, SaveFilePath)
						End If
					  If Action = "Add"  Then 
					    If ChannelID=1 and F_B_Arr(17)=1 Then
						Verific=KS.ChkClng(KS.G("Verific"))
						Else
						Verific = 1
						End if
						SqlStr = "select top 1 * from " & KS.C_S(ChannelID,2) &" where 1=0"
						RS.Open SqlStr, conn, 1, 3
						RS.AddNew
						RS("TitleType")      = TitleType
						RS("Title")          = Title
						RS("Fulltitle")      = Fulltitle
						RS("TitleFontColor") = TitleFontColor
						RS("TitleFontType")  = TitleFontType
						RS("Intro")          = Intro
						RS("ArticleContent") = ArticleContent
						RS("PageTitle")      = PageTitle
						RS("Changes")        = Changes
						RS("PicNews")        = PicNews
						RS("PhotoUrl")       = PhotoUrl
						RS("Recommend")      = Recommend
						RS("IsTop")          = IsTop
						RS("Rolls")          = Rolls
						RS("Strip")          = Strip
						RS("Popular")        = Popular
						RS("Verific")        = Verific
						RS("Tid")            = Tid
						RS("KeyWords")       = KeyWords
						RS("Author")         = Author
						RS("Origin")         = Origin
						RS("AddDate")        = AddDate
						RS("Rank")           = Rank
						RS("Slide")          = Slide
						RS("Comment")        = Comment
						RS("TemplateID")     = TemplateID
						RS("WapTemplateID")  = WapTemplateID
						RS("Hits")           = Hits
						RS("HitsByDay")      = HitsByDay
						RS("HitsByWeek")     = HitsByWeek
						RS("HitsByMonth")    = HitsByMonth
						RS("Fname")          = Fname
						RS("Inputer")        = KS.C("AdminName")
						RS("RefreshTF")      = Makehtml
						RS("DelTF")          = 0
						RS("IsSign")         = IsSign
						RS("SignUser")       = SignUser
						RS("SignDateLimit")  = SignDateLimit
						RS("SignDateEnd")    = SignDateEnd
						RS("Province")       = Province
						RS("City")           = City
						
						RS("ReadPoint")      = ReadPoint
				        RS("ChargeType")     = ChargeType
				        RS("PitchTime")      = PitchTime
				        RS("ReadTimes")      = ReadTimes
						RS("InfoPurview")    = InfoPurview
						RS("arrGroupID")     = arrGroupID
						RS("DividePercent")  = DividePercent
						If IsArray(UserDefineFieldArr) Then
						For I=0 To Ubound(UserDefineFieldArr,2)
							If UserDefineFieldArr(3,I)=10  Then   '支持HTML时
							 RS("" & UserDefineFieldArr(0,I) & "")=Request.Form(UserDefineFieldArr(0,I))
							else
							 RS("" & UserDefineFieldArr(0,I) & "")=KS.G(UserDefineFieldArr(0,I))
							end if
							If UserDefineFieldArr(12,I)="1"  Then
							RS("" & UserDefineFieldArr(0,I) & "_Unit")=KS.G(UserDefineFieldArr(0,I)&"_Unit")
							End If
						Next
						End If
						RS.Update
						
					   '写入Session,添加下一篇文章调用
					   Session("KeyWords") = KeyWords
					   Session("Author")   = Author
					   Session("Origin")   = Origin
					   RS.MoveLast
					  If Left(Ucase(Fname),2)="ID" Then
					   RS("Fname") = RS("ID") & FnameType
					   RS.Update
					  End If
					  For I=0 To Ubound(SpecialID)
						Conn.Execute("Insert Into KS_SpecialR(SpecialID,InfoID,ChannelID) values(" & SpecialID(I) & "," & RS("ID") & "," & ChannelID & ")")
					  Next
					
					
					If SelectInfoList<>"" Then
					 SelectInfoList=Split(SelectInfoList,",")
					 For I=0 To Ubound(SelectInfoList)
					  If KS.FoundInArr(HasInRelativeID,SelectInfoList(i),",")=false Then
					   Conn.Execute("Insert Into KS_ItemInfoR(ChannelID,InfoID,RelativeChannelID,RelativeID) values(" & ChannelID &"," & RS("ID") & "," & Split(SelectInfoList(i),"|")(0) & "," & Split(SelectInfoList(i),"|")(1) & ")")
					   HasInRelativeID=HasInRelativeID & SelectInfoList(i) & ","
					  End If
					 Next
					End If
					  
					  Call LFCls.AddItemInfo(ChannelID,RS("ID"),Title,Tid,Intro,KeyWords,PhotoUrl,AddDate,KS.C("AdminName"),Hits,HitsByDay,HitsByWeek,HitsByMonth,Recommend,Rolls,Strip,Popular,Slide,IsTop,Comment,Verific,RS("Fname"))
					 
	 				  '关联上传文件
					  Call KS.FileAssociation(ChannelID,Rs("ID"),ArticleContent & PhotoUrl,0)

					  Call RefreshHtml(1) 
					  RS.Close:Set RS = Nothing

					ElseIf Action = "Edit" Or Action="Verify" Then
					
					 If ChannelID=1 And Action<>"Verify" and F_B_Arr(17)=1 Then
						Verific=KS.ChkClng(KS.G("Verific"))
					 Else
						Verific = 1
					 End if
						
					If Action="Verify" Then 
					 Call KS.ReplaceUserFile(ArticleContent,ChannelID)
					 Call KS.ReplaceUserFile(PhotoUrl,ChannelID)
					End If
					NewsID = Trim(Request("NewsID"))
					SqlStr = "SELECT * FROM " & KS.C_S(ChannelID,2) &" Where ID=" & NewsID & ""
						RS.Open SqlStr, conn, 1, 3
						If RS.EOF And RS.BOF Then
						 .die ("<script>alert('参数传递出错!');history.back(-1);</script>")
						End If
						RS("TitleType") = TitleType
						RS("Title") = Title
						RS("Fulltitle")=Fulltitle
						RS("TitleFontColor") = TitleFontColor
						RS("TitleFontType") = TitleFontType
						RS("ArticleContent") = ArticleContent
						RS("PageTitle")      = PageTitle
						RS("Changes") = Changes
						RS("PicNews") = PicNews
						RS("PhotoUrl") = PhotoUrl
						RS("Recommend") = Recommend
						RS("IsTop") = IsTop
						RS("Rolls") = Rolls
						RS("Strip") = Strip
						RS("Popular") = Popular
						RS("Tid") = Tid
						RS("KeyWords") = KeyWords
						RS("Author") = Author
						RS("Origin") = Origin
						RS("AddDate") = AddDate
						RS("Rank") = Rank
						RS("Slide") = Slide
						RS("Comment") = Comment
						RS("TemplateID") = TemplateID
						RS("WapTemplateID")  = WapTemplateID
						If Action="Verify" Then
						    Inputer=RS("Inputer")
						End If
	
						RS("Verific") = Verific
						If Makehtml = 1 Then
						 RS("RefreshTF") = 1
						End If
						RS("IsSign")         = IsSign
						RS("SignUser")       = SignUser
						RS("SignDateLimit")  = SignDateLimit
						RS("SignDateEnd")    = SignDateEnd
						RS("Province")       = Province
						RS("City")           = City
						RS("Hits") = Hits
						RS("HitsByDay") = HitsByDay
						RS("HitsByWeek") = HitsByWeek
						RS("HitsByMonth") = HitsByMonth
						RS("ReadPoint")=	ReadPoint
				        RS("ChargeType")=ChargeType
				        RS("PitchTime")=PitchTime
				        RS("ReadTimes")=ReadTimes
						RS("InfoPurview")=InfoPurview
						RS("arrGroupID")=arrGroupID
						RS("DividePercent")=DividePercent
						RS("Intro")=Intro
						If IsArray(UserDefineFieldArr) Then
						For I=0 To Ubound(UserDefineFieldArr,2)
							If UserDefineFieldArr(3,I)=10  Then   '支持HTML时
							 RS("" & UserDefineFieldArr(0,I) & "")=Request(UserDefineFieldArr(0,I))
							else
							 RS("" & UserDefineFieldArr(0,I) & "")=KS.G(UserDefineFieldArr(0,I))
							end if
							If UserDefineFieldArr(12,I)="1"  Then
							RS("" & UserDefineFieldArr(0,I) & "_Unit")=KS.G(UserDefineFieldArr(0,I)&"_Unit")
							End If
						Next
						End If
						RS.Update
						RS.MoveLast
						If TID<>Request.Form("OldClassID") Then
					     Call KSCls.DelInfoFile(ChannelID,Request.Form("OldClassID"),Split(RS("ArticleContent"), "[NextPage]"),RS("Fname"))
					    End If
						Conn.Execute("Delete From KS_SpecialR Where InfoID=" & NewsID & " and channelid=" & ChannelID)
						For I=0 To Ubound(SpecialID)
						Conn.Execute("Insert Into KS_SpecialR(SpecialID,InfoID,ChannelID) values(" & SpecialID(I) & "," & RS("ID") & "," & ChannelID & ")")
						Next
						 Conn.Execute("Delete From KS_ItemInfoR Where InfoID=" & NewsID & " and channelid=" & ChannelID)
						If SelectInfoList<>"" Then
						 SelectInfoList=Split(SelectInfoList,",")
						 For I=0 To Ubound(SelectInfoList)
						  If KS.FoundInArr(HasInRelativeID,SelectInfoList(i),",")=false Then
					       Conn.Execute("Insert Into KS_ItemInfoR(ChannelID,InfoID,RelativeChannelID,RelativeID) values(" & ChannelID &"," & RS("ID") & "," & Split(SelectInfoList(i),"|")(0) & "," & Split(SelectInfoList(i),"|")(1) & ")")
						   HasInRelativeID=HasInRelativeID & SelectInfoList(i) & ","
						  End If
						 Next
						End If
						
						Call LFCls.UpdateItemInfo(ChannelID,NewsID,Title,Tid,Intro,KeyWords,PhotoUrl,AddDate,Hits,HitsByDay,HitsByWeek,HitsByMonth,Recommend,Rolls,Strip,Popular,Slide,IsTop,Comment,Verific)
	 				  '关联上传文件
					  Call KS.FileAssociation(ChannelID,NewsID,ArticleContent & PhotoUrl,1)

					   Call RefreshHtml(2)
					   
					   RS.Close:Set RS = Nothing
						IF Action="Verify" Then     '如果是审核投稿文章，对用户，进行加积分等，并返回签收文章管理
							  '对用户进行增值，及发送通知操作
							  IF Inputer<>"" And Inputer<>KS.C("AdminName") Then Call KS.SignUserInfoOK(ChannelID,Inputer,Title,NewsID)
							 KS.Echo ("<script> parent.frames['MainFrame'].focus();alert('" & KS.C_S(ChannelID,3) & "成功签收,系统已发送一封站内通知信给投稿者!');location.href='KS.ItemInfo.asp?ComeFrom=Verify&ChannelID=" & ChannelID &"&Page=" & Page & "';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&OpStr="& server.URLEncode(KS.C_S(ChannelID,1) & " >> <font color=red>签收会员" & KS.C_S(ChannelID,3)) & "</font>';</script>") 
				       End IF
					   
						If KeyWord <> "" Then
							 KS.Echo  ("<script> parent.frames['MainFrame'].focus();alert('" & KS.C_S(ChannelID,3) & "修改成功!');location.href='KS.Article.asp?ChannelID=" & ChannelID &"&Page=" & Page & "&KeyWord=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate & "';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=ArticleSearch&OpStr=" & Server.URLEncode(KS.C_S(ChannelID,1) & " >> <font color=red>搜索结果</font>") & "';</script>")
						End If
					End If
				End If
				End With
			End Sub
			
			Sub RefreshHtml(Flag)
			     Dim TempStr,EditStr,AddStr
			    If Flag=1 Then
				  TempStr="添加":EditStr="修改" & KS.C_S(ChannelID,3) & "":AddStr="继续添加" & KS.C_S(ChannelID,3) & ""
				Else
				  TempStr="修改":EditStr="继续修改" & KS.C_S(ChannelID,3) & "":AddStr="添加" & KS.C_S(ChannelID,3) & ""
				End If
			    With Response
				     .Write "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
					 .Write "<meta http-equiv=Content-Type content=""text/html; charset=gb2312"">"
					 .Write "<script language='JavaScript' src='../KS_Inc/Jquery.js'></script>"
					 .Write " <Br><br><br><table align='center' width=""95%"" height='200' class='ctable' cellpadding=""1"" cellspacing=""1"">"
					  .Write "	  <tr class=""sort""> "
					  .Write "		<td  height=""36"" colspan=2>系统操作提示信息</td>" & vbcrlf
					  .Write "	  </tr>"
                      .Write "    <tr class='tdbg'>"
					  .Write "          <td align='center'><img src='images/succeed.gif'></td>"
					  .Write "<td><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;恭喜，" & TempStr &"" & KS.C_S(ChannelID,3) & "成功！</b><br>"
			           '判断是否立即发布
					   If Makehtml = 1 Then
					      .Write "<div style=""margin-top:15px;border: #E7E7E7;height:220; overflow: auto; width:100%"">" 
						  
						  If KS.C_S(ChannelID,7)=1 Or KS.C_S(ChannelID,7)=2 Then
						  	 .Write "<div><iframe src=""Include/RefreshHtmlSave.Asp?ChannelID=" & ChannelID &"&Types=Content&RefreshFlag=ID&ID=" & RS("ID") &""" width=""100%"" height=""80"" frameborder=""0"" allowtransparency='true'></iframe></div>"
						  Else
						  .Write "<div style=""height:25px""><li>由于" & KS.C_S(ChannelID,1) & "没有启用生成HTML的功能，所以ID号为 <font color=red>" & NewsID & "</font>  的" & KS.C_S(ChannelID,3) & "没有生成!</li></div> "
						  End If
						  
							If KS.C_S(ChannelID,7)<>1 Then
							  .Write "<div style=""height:25px""><li>由于" & KS.C_S(ChannelID,1) & "的栏目页没有启用生成HTML的功能，所以ID号为 <font color=red>" & TID & "</font>  的栏目没有生成!</li></div> "
							Else
							 If KS.C_S(ChannelID,9)<>1 Then
								  Dim FolderIDArr:FolderIDArr=Split(left(KS.C_C(Tid,8),Len(KS.C_C(Tid,8))-1),",")
								  For I=0 To Ubound(FolderIDArr)
								  .Write "<div align=center><iframe src=""Include/RefreshHtmlSave.Asp?ChannelID=" & ChannelID &"&Types=Folder&RefreshFlag=ID&FolderID=" & FolderIDArr(i) &""" width=""100%"" height=""90"" frameborder=""0"" allowtransparency='true'></iframe></div>"
								   Next
							 End If
						   End If
					   If Split(KS.Setting(5),".")(1)="asp" or KS.C_S(ChannelID,9)<>3 Then
					   Else
					     .Write "<div align=center><iframe src=""Include/RefreshIndex.asp?ChannelID=" & ChannelID &"&RefreshFlag=Info"" width=""100%"" height=""80"" frameborder=""0"" allowtransparency='true'></iframe></div>"
					   End If
					   .Write "</div>"
					End If
					.Write   "</td></tr>"
					.Write "	  <tr>"
					.Write "		<td  class='tdbg' height=""25"" align=""right"" colspan=2>【<a href=""KS.Article.asp?ChannelID=" & ChannelID &"&Page=" & Page & "&Action=Edit&KeyWord=" & KeyWord &"&SearchType=" & SearchType &"&StartDate=" & StartDate & "&EndDate=" & EndDate &"&ID=" & RS("ID") & """><strong>" & EditStr &"</strong></a>】&nbsp;【<a href=""#"" onclick=""location.href='KS.Article.asp?ChannelID=" & ChannelID &"&Action=Add&FolderID=" & Tid & "';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=" & Server.URLEncode("添加" & KS.C_S(ChannelID,3)) & "&ButtonSymbol=AddInfo&FolderID=" & Tid & "';""><strong>" & AddStr & "</strong></a>】&nbsp;【<a href=""#"" onclick=""location.href='KS.ItemInfo.asp?ID=" & Tid & "&ChannelID=" & ChannelID &"&Page=" & Page&"';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=ViewFolder&FolderID=" & Tid & "';""><strong>" & KS.C_S(ChannelID,3) & "管理</strong></a>】&nbsp;【<a href=""" & KS.GetDomain & "Item/Show.asp?m=" & ChannelID & "&d=" & RS("ID") & """ target=""_blank""><strong>预览" & KS.C_S(ChannelID,3) & "内容</strong></a>】</td>"
					.Write "	  </tr>"
					.Write "	</table>"	
					.Flush			
			End With
		End Sub
	
		'名称检测
        Sub CheckTitle()
		%>
		 <html>
		<head>
		<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
		<title>重名检测</title>
		<link href="include/ModeWindow.css" rel="stylesheet">
		</head>
		<body leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
		<%
			Dim Title,rsCheck
			Title=Trim(request("Title"))
			Response.write "<font color='#0000ff'>" & KS.C_S(ChannelID,3) & "标题</font>类似于<font color='#ff0033'>" & Title & "</font>的" & KS.C_S(ChannelID,3) & ""
			Set rsCheck=conn.Execute("Select Top 20 ID,Title From " & KS.C_S(ChannelID,2) &" Where Title like '%"&Title&"%'")
			if not(rsCheck.bof and rsCheck.eof) then
				Response.write "<ol>"
				do while Not rsCheck.Eof
					Response.write "<li>" & Replace(rsCheck(1),Title,"<font color='#ff0033'>" & Title & "</font>")
					rsCheck.MoveNext
				Loop
				Response.write "</ol>"
			else
		
				Response.write "<li>无任何类似" & KS.C_S(ChannelID,3) & "</li><br />"
			end if
			rsCheck.Close : set rsCheck=Nothing
		%>
		</body>
		</html>
		<%
	End Sub
	
	Sub SelectUser()
		response.cachecontrol="no-cache"
		response.addHeader "pragma","no-cache"
		response.expires=-1
		response.expiresAbsolute=now-1
		With Response
			.Write "<html>"
			.Write "<head>"
			.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
			.Write "<META HTTP-EQUIV=""pragma"" CONTENT=""no-cache"">" 
			.Write "<META HTTP-EQUIV=""Cache-Control"" CONTENT=""no-cache, must-revalidate"">"
			.Write "<META HTTP-EQUIV=""expires"" CONTENT=""Wed, 26 Feb 1997 08:21:57 GMT"">"
            .Write "<base target='_self'>" & vbCrLf
			.Write "<title>选择用户</title>"
			.Write "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
			.Write "<body>"
			%>
			
		<form method='post' name='myform' action=''>	
		<table width='560' border='0' align='center' style="margin-top:4px" cellpadding='2' cellspacing='0' class='border'>
		  <tr class='title' height='22'>
			<td valign='top'><b>已经选定的用户名：</b></td>
			<td align='right'><a href='javascript:window.returnValue=myform.UserList.value;window.close();'>返回&gt;&gt;</a></td>
		  </tr>
		  <tr class='tdbg'>
			<td><input type='text' name='UserList' size='40' maxlength='200' readonly='readonly'></td>
			<td align='center'><input type='button' name='del1' onclick='del(1)' class='button' value='删除最后'> <input type='button' name='del2' onclick='del(0)' value='删除全部' class='button'></td>
		  </tr>
		</table>
		<br/>
		<table width='560' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>
  <tr height='22' class='title'>
    <td><b><font color=red>会员</font>列表：</b></td><td align=right><input name='Key' type='text' size='20' value=>&nbsp;&nbsp;<input type='submit' class="button" value='查找'></td>
  </tr>
  <tr>
    <td valign='top' colspan=2>
	<table width='550' border='0' cellspacing='1' cellpadding='1' bgcolor='#f9f9f9'>
	 <%
	 Page=KS.ChkClng(request("page"))
	 if Page=0 Then Page=1
	 MaxPerPage=40
	 dim sqlstr,AllUserList,TotalPages,param
	 if request("key")<>"" then
	   param=" where username like '%" & KS.G("Key") & "%'"
	 end if
	 
	 sqlstr="select username from ks_user " & Param & " order by userid"
	 dim rs:set rs=server.CreateObject("adodb.recordset")
	 RS.Open SQLStr, conn, 1, 1
	 If Not RS.EOF Then
			totalPut = Conn.Execute("Select count(userid) from [ks_user] " & Param)(0)
								If Page < 1 Then Page = 1
								If (Page - 1) * MaxPerPage < totalPut Then
										RS.Move (Page - 1) * MaxPerPage
								Else
										Page = 1
								End If
								
					Dim SQL:SQL=RS.GetRows(MaxPerPage)
					RS.Close : Set RS=Nothing
			  .write "<tr>"
			For I=0 To Ubound(SQL,2)
				If AllUserList = "" Then
					AllUserList = SQL(0,I)
				Else
					AllUserList = AllUserList & "," & SQL(0,I)
				End If
			  .write "<td align='center'><a href='#' onclick='add(""" &SQL(0,I) & """)'>" &SQL(0,I) & "</a></td>"
			  If ((i+1) Mod 8) = 0 And i > 0 Then Response.Write "</tr><tr>"
			Next
			  .Write "</tr>"
	End If
	%>
	  <tr class='tdbg'>
		<td align='center' colspan=8 height=30><a href='#' onclick='add("<%=AllUserList%>")'><b>增加以上所有用户名</b></a></td>
	  </tr>
	</table>
  </td>
  </tr>
 </table>
		</form>
		
	<table width='550' border='0' cellspacing='1' cellpadding='1'>
    <tr>
	 <td>
  <%
  Call KS.ShowPage(totalput, MaxPerPage, "", Page,true,true)
 %>	
    </td>
  </tr>
  </table>
		
		<div style="margin-top:10px;text-align:center"><input type="button" onClick="window.close()" class="button" value=" 关 闭 "></div>
		
			<script language="javascript">
myform.UserList.value='<%=request("DefaultValue")%>';
var oldUser='';
function add(obj)
{
    if(obj==''){return false;}
    if(myform.UserList.value=='')
    {
        myform.UserList.value=obj;
        window.returnValue=myform.UserList.value;
        return false;
    }
    var singleUser=obj.split(',');
    var ignoreUser='';
    for(i=0;i<singleUser.length;i++)
    {
        if(checkUser(myform.UserList.value,singleUser[i]))
        {
            ignoreUser=ignoreUser+singleUser[i]+" "
        }
        else
        {
            myform.UserList.value=myform.UserList.value+','+singleUser[i];
        }
    }
    if(ignoreUser!='')
    {
        alert(ignoreUser+'用户名已经存在，此操作已经忽略！');
    }
    window.returnValue=myform.UserList.value;
}
function del(num)
{
    if (num==0 || myform.UserList.value=='' || myform.UserList.value==',')
    {
        myform.UserList.value='';
        return false;
    }

    var strDel=myform.UserList.value;
    var s=strDel.split(',');
    myform.UserList.value=strDel.substring(0,strDel.length-s[s.length-1].length-1);
    window.returnValue=myform.UserList.value;
}
function checkUser(UserList,thisUser)
{
  if (UserList==thisUser){
        return true;
  }
  else{
    var s=UserList.split(',');
    for (j=0;j<s.length;j++){
        if(s[j]==thisUser)
            return true;
    }
    return false;
  }
}
</script>
		
		
			<%
			.Write "</body>"
			.Write "</html>"
		End With
	 End Sub
	
		'执行过滤
	Function FilterScript(ByVal Content)
		   If KS.G("FilterIframe") = "1" Then  Content = KS.ScriptHtml(Content, "Iframe", 1)
		   If KS.G("FilterObject") = "1" Then  Content = KS.ScriptHtml(Content, "Object", 2)
		   If KS.G("FilterScript") = "1" Then  Content = KS.ScriptHtml(Content, "Script", 2)
		   If KS.G("FilterDiv")    = "1" Then  Content = KS.ScriptHtml(Content, "Div", 3)
	       If KS.G("FilterTable")  = "1" Then  Content = KS.ScriptHtml(Content, "table", 3)
		   If KS.G("FilterTr")     = "1" Then  Content = KS.ScriptHtml(Content, "tr", 3)
	       If KS.G("FilterTd")     = "1" Then  Content = KS.ScriptHtml(Content, "td", 3)
		   If KS.G("FilterTd")     = "1" Then  Content = KS.ScriptHtml(Content, "Span", 3)
		   If KS.G("FilterImg")    = "1" Then  Content = KS.ScriptHtml(Content, "Img", 3)
		   If KS.G("FilterFont")   = "1" Then  Content = KS.ScriptHtml(Content, "Font", 3)
		   If KS.G("FilterA")      = "1" Then  Content = KS.ScriptHtml(Content, "A", 3)
		   If KS.G("FilterHtml")   = "1" Then  Content = KS.LoseHtml(Content)
		   FilterScript=Content
	End Function

End Class
%> 
