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
Set KSCls = New Admin_ItemInfo
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_ItemInfo
        Private KS,ComeUrl,KSCls
		'=====================================声明本页面全局变量==============================================================
        Private ID, I, totalPut, Page, RS,ComeFrom,ItemManageUrl
		Private KeyWord, SearchType, StartDate, EndDate,SearchParam, MaxPerPage,T, TitleStr, VerificStr
		Private TypeStr, AttributeStr, FolderID, TemplateID,FolderName, Action,TotalPages
		Private FileName,SqlStr,Errmsg,Makehtml,Tid,Fname,KSRObj,SaveFilePath
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
		If ChannelID=0 Then ChannelID=1
		Session("ThumbnailsConfig")=KS.C_S(ChannelID,46)
		F_B_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(0),"|")
		F_V_Arr=Split(Split(KS.C_S(ChannelID,5),"@@@")(1),"|")
		
		Select Case KS.C_S(ChannelID,6)
		 Case 2:ItemManageUrl="KS.Picture.asp"
		 Case 3:ItemManageUrl="KS.Down.asp"
		 Case 4:ItemManageUrl="KS.Flash.asp"
		 Case 5:ItemManageUrl="KS.Shop.asp"
		 Case 7:ItemManageUrl="KS.Movie.asp"
		 Case 8:ItemManageUrl="KS.Supply.asp"
		 Case Else:ItemManageUrl="KS.Article.asp"
		End Select
		
		KeyWord    = KS.G("KeyWord")
		SearchType = KS.G("SearchType")
		StartDate  = KS.G("StartDate")
		EndDate    = KS.G("EndDate")
		Action     = KS.G("Action")
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
		If Page=0 Then  Page = 1
		
		Select Case Action
		 Case "Recely"
           If Not KS.ReturnPowerResult(0, "M010006") Then 
		    Call KS.ReturnErr(1, "")
		   Else
             Call KSCls.Recely(ChannelID)
           End If
		 Case "RecelyBack"
		    Call KSCls.RecelyBack(ChannelID)
		 Case "Delete"
			If Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10004") Then  
			 Call KS.ReturnErr(1, "")
			Else
		    Call KSCls.DelBySelect(ChannelID)
			End If
		 Case "DeleteAll"
			If Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10004") Then  
			 Call KS.ReturnErr(1, "")
			Else
		    Call KSCls.DeleteAll() 
			End If
		 Case "VerifyAll"
            Call KSCls.VerificAll(ChannelID)
		 Case "Tuigao"
		    Call KSCls.Tuigao(ChannelID)
		 Case "BatchSet"
		    Call KSCls.BatchSet(ChannelID)
		 Case "JS"
		   If Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10007") Then  
			  Call KS.ReturnErr(0, "")
			Else
			  Call AddToJS()
			End If
		 Case "Special"
		  If Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10006") Then 
			 Call KS.ReturnErr(0, "")
		  Else
		     Call KSCls.AddToSpecial(ChannelID)
		  End If
		 Case "SetAttribute"
			If Not KS.ReturnPowerResult(0, "M010005") Then 
				 Call KS.ReturnErr(1, "")
			Else
		         Call SetAttribute()
			End If
		 Case "Paste"
		  	If Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10011") Then  
			   Call KS.ReturnErr(1, "")   
            Else
		       Call KSCls.Paste(ChannelID)
			End If 
		 Case Else
		       Call ItemInfoMain()
		End Select
		
	 End Sub
	 
	 Sub ItemInfoMain()
		ID = KS.G("ID"):If ID = "" Then ID = "0"
		MaxPerPage = Cint(KS.C_S(ChannelID,11))     '取得每页显示数量
		With KS
		.echo "<html>"
		.echo "<head>"
		.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		.echo "<title>管理主页面</title>"
		.echo "<link href=""include/admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
		.echo "<script language=""JavaScript"">"
		.echo " var ClassID='" & ID & "';                //目录ID" & vbCrLf
		.echo " var Page='" & Page & "';                 //当前页码" & vbCrLf
		.echo " var KeyWord='" & KeyWord & "';           //关键字" & vbCrLf
		.echo " var SearchParam='" & SearchParam & "';   //搜索参数集合" & vbCrLf
		
		.echo "</script>" & vbCrLf
		.echo "<script language=""JavaScript"" src=""../KS_Inc/Common.js""></script>" & vbCrLf
		.echo "<script language=""JavaScript"" src=""../KS_Inc/JQuery.js""></script>" & vbCrLf
		.echo "<script language=""JavaScript"" src=""../KS_Inc/kesion.box.js""></script>" & vbCrLf
		%>
		<script language="JavaScript">
		function ClassToggle(f)
		{
		  setCookie("classExtStatus",f)
		  $('#classNav').toggle('slow');
		  $('#classOpen').toggle('show');
		}
		function ProcessTuigao(ev,Id)
		{
		    var ids=get_Ids(document.myform);
			if (Id=='') Id=ids;
			if (Id=='')
			{
			  alert('对不起，您没有选中要退稿的文档!');
			  return;
			}
		 	mousepopup(ev,"<b>退稿原因</b>","<div style='height:200px;text-align:center'><form name='rform' action='KS.ItemInfo.asp?action=Tuigao&Page=<%=Page%>' method='post'><input type='hidden' name='channelid' value='<%=ChannelID%>'><input type='hidden' name='Id' value='"+Id+"'><textarea name='AnnounceContent' style='width:300px;height:130px'>您好{$UserName}，您发布的稿件“{$Title}”不符合本站要求，请修改后再重新提交！</textarea><br><br/><label><input type='checkbox' value='1' name='Email' checked>发表站内短信通知</label> <input type='submit' value='确定退稿' class='button'> <input type='submit' value='取消退稿' class='button' onclick='closeWindow();'></form></div>",400);

		}
		function CreateHtml()
		{   var ids=get_Ids(document.myform);
			if (ids!='')
		PopupCenterIframe('发布选中文档','Include/RefreshHtmlSave.Asp?ChannelID=<%=ChannelID%>&Types=Content&RefreshFlag=IDS&ID='+ids,530,110,'no')
			else 
			alert('请选择要发布的文档!');
		}		
		function CreateNews()
		{   
		   location.href='<%=ItemManageUrl%>?ChannelID=<%=ChannelID%>&Action=Add&FolderID='+ClassID;
           $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ChannelID=<%=ChannelID%>&OpStr='+escape("添加<%=KS.C_S(ChannelID,3)%>")+'&ButtonSymbol=AddInfo&FolderID='+ClassID;
		}
		function VerifyInfo()
		{
		   location.href='KS.ItemInfo.asp?ComeFrom=Verify&ChannelID=<%=ChannelID%>';
           $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ChannelID=<%=ChannelID%>&OpStr='+escape("签收<%=KS.C_S(ChannelID,3)%>")+'&ButtonSymbol=Disabled';
		}
		function Edit()
		{   var ids=get_Ids(document.myform);
			 if (ids!='')
					 if (ids.indexOf(',')==-1){
						 location.href='<%=ItemManageUrl%>?Page='+Page+'&Action=Edit&'+SearchParam+'&ID='+ids;
						 if (KeyWord=='')
							$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape("编辑<%=KS.C_S(ChannelID,3)%>")+'&ButtonSymbol=AddInfo&FolderID='+ClassID;
						 else
						   $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape("<%=KS.C_S(ChannelID,1)%> >> 搜索结果 >> <font color=red>编辑<%=KS.C_S(ChannelID,3)%></font>")+'&ButtonSymbol=AddInfo';
						 }
					   else alert('一次只能够编辑一<%=KS.C_S(ChannelID,4)%><%=KS.C_S(ChannelID,3)%>');
					 
			else 
			{
			alert('请选择要编辑的<%=KS.C_S(ChannelID,3)%>');
			}
		}
		function Recely()
		{ 
		   if (chk_idBatch(myform,'将选中的<%=KS.C_S(ChannelID,3)%>放入回收站吗')==true)
		   {
		    $('input[name=action]').val('Recely'); 
			$('form[name=myform]').submit();
		   }
		}
		function BackRecely()
		{
		   if (chk_idBatch(myform,'将选中的<%=KS.C_S(ChannelID,3)%>还原吗')==true)
		   {
		    $('input[name=action]').val('RecelyBack'); 
			$('form[name=myform]').submit();
		   }
		}
		function Delete()
		{ 
		   if (chk_idBatch(myform,'此操作不可逆,彻底删除选中的<%=KS.C_S(ChannelID,3)%>吗')==true)
		   {
		    $('input[name=action]').val('Delete'); 
			$('form[name=myform]').submit();
		   }
		}
		function DelAll()
		{
		  if (confirm('友情提示:\n\n一键清空将清除所有模型里的回收站文档,且此操作不可逆，确定清空回收站吗？')==true)
		  {
		    $('input[name=action]').val('DeleteAll');
			$('form[name=myform]').submit();
		  }
		}
		function VerificAll()
		{
		   if (chk_idBatch(myform,'确定批量审核选中的<%=KS.C_S(ChannelID,3)%>吗')==true)
		   {
		    $('input[name=action]').val('VerifyAll'); 
			$('form[name=myform]').submit();
		   }

		}
		function Tuigao()
		{
		  ProcessTuigao(event,'')
		  return;
		
		 if (chk_idBatch(myform,'确定批量退稿选中的<%=KS.C_S(ChannelID,3)%>吗')==true)
		   {
		    $('input[name=action]').val('Tuigao'); 
			$('input[name=myform]').submit();
		   }
		}
		
		function Copy()
		{
		    var ids=get_Ids(document.myform);
			if (ids!='')
			  {
			   top.CommonCopyCut.ChannelID=<%=ChannelID%>;
			   top.CommonCopyCut.PasteTypeID=2;
			   top.CommonCopyCut.SourceFolderID=ClassID;
			   top.CommonCopyCut.FolderID='0';
			   top.CommonCopyCut.ContentID=ids;
			  }
			else
			 alert('请选择要复制的<%=KS.C_S(ChannelID,3)%>!');
		}
		function Paste()
		{ 
		  if (ClassID=='0')
		   { top.CommonCopyCut.PasteTypeID=0;
			 alert('目标目录不存在!');
			}
		  if (top.CommonCopyCut.ChannelID==<%=ChannelID%> && top.CommonCopyCut.PasteTypeID!=0)
		   {  var Param='';
			  Param='?ChannelID=<%=ChannelID%>&Action=Paste&Page='+Page;
			  Param+='&PasteTypeID='+top.CommonCopyCut.PasteTypeID+'&DestFolderID='+ClassID+'&SourceFolderID='+top.CommonCopyCut.SourceFolderID+'&FolderID='+top.CommonCopyCut.FolderID+'&ContentID='+top.CommonCopyCut.ContentID;
			  if (top.CommonCopyCut.PasteTypeID==2) //复制
			 {
				location.href='KS.ItemInfo.asp'+Param;
			 }
			else
			 alert('非法操作!');
		   }
		  else
		   alert('系统剪切板没有内容!');
		}
		function AddToSpecial()
		{  var ids=get_Ids(document.myform);
			if (ids!='')
				{     
				OpenWindow('KS.Frame.asp?PageTitle='+escape('<%=KS.C_S(ChannelID,3)%>加入到专题')+'&URL=KS.ItemInfo.asp&ChannelID=<%=ChannelID%>&Action=Special&NewsID='+ids,300,350,window);
				}
			else alert('请选择要加入专题的<%=KS.C_S(ChannelID,3)%>!');
			Select(2);
		}
		function AddToJS()
		{  var ids=get_Ids(document.myform);
			if (ids!='')
				{     
				OpenWindow('KS.Frame.asp?PageTitle='+escape('<%=KS.C_S(ChannelID,3)%>加入到自由JS')+'&URL=KS.ItemInfo.asp&ChannelID=<%=ChannelID%>&Action=JS&NewsID='+ids,300,100,window);
				}
			else alert('请选择要加入JS的<%=KS.C_S(ChannelID,3)%>!');
			Select(2);
		}
		function SetAttribute()
		{   var ids=get_Ids(document.myform);
		     if (ids=='')
			 {
			  alert('请选择要设置属性的<%=KS.C_S(ChannelID,3)%>!');
			  return;
			 }
			 OpenWindow('KS.Frame.asp?PageTitle='+escape('批量设置<%=KS.C_S(ChannelID,3)%>属性')+'&URL=KS.ItemInfo.asp&ChannelID=<%=ChannelID%>&Action=SetAttribute&ID='+ids,500,420,window);
			 window.location.reload();
		}
		function MoveToClass()
		{   var ids=get_Ids(document.myform);
		     if (ids=='')
			 {
			  alert('请选择要批量移动的<%=KS.C_S(ChannelID,3)%>!');
			  return;
			 }
			 OpenWindow('KS.Frame.asp?PageTitle='+escape('<%=KS.C_S(ChannelID,3)%>批量移动<%=KS.C_S(ChannelID,3)%>')+'&URL=KS.Class.asp&ChannelID=<%=ChannelID%>&Action=MoveInfo&From=main&ID='+ids,500,400,window);
			 window.location.reload();		
		}
		function ViewArticle(ArticleID)
		{
		window.open ('../Item/Show.asp?m=<%=ChannelID%>&d='+ArticleID);
		}
		function setstatus(Obj)
		  {var today=new Date()
			if (Obj.nextSibling.style.display=='none')
			 {
			  Obj.nextSibling.style.display='';
			  $('#StartDate').val(today.getYear()+'-'+(today.getMonth()+1)+'-01');
			  $('#EndDate').val(today.getYear()+'-'+(today.getMonth()+1)+'-'+today.getDate());
			 }
			else 
			{
			 Obj.nextSibling.style.display='none';
			 $('#StartDate').val('');
			 $('#EndDate').val('');
			 }
		}
		function set(o,v)
		{
		 
		 if (parseInt(v)!=0)
		  {
		  var ids=get_Ids(document.myform);
		  if (ids!='')
		   {
					if (confirm('确定将选中的<%=KS.C_S(ChannelID,3)%>'+o.value)==true)
					{
					    $('#SetAttributeBit').val(v);
						$('input[name=action]').val('BatchSet'); 
						$('form[name=myform]').submit();

					}
			}
		   else
		    alert('请选择要设置的<%=KS.C_S(ChannelID,3)%>');
		  }
		}
		function GetKeyDown()
		{
		if (event.ctrlKey)
		  switch  (event.keyCode)
		  {      case 90 : Select(2); break;
			 case 77 : CreateNews();break;
			 case 65 : Select(0);break;
			 case 83 : AddToSpecial();break;
			 case 74 : AddToJS();break;
			 case 85 : SetAttribute();break;
			 case 67 : 
				{event.keyCode=0;event.returnValue=false;Copy();}
                 break;
			 case 86 : 
			   if (top.CommonCopyCut.ChannelID==<%=ChannelID%> && top.CommonCopyCut.PasteTypeID!=0 && ClassID!='0')
			   { event.keyCode=0;event.returnValue=false;Paste();}
			   else
			    {
				 if (top.CommonCopyCut.PasteTypeID!=0)
				alert('请转向目标栏目后再粘贴!');
				return;
				}
				break;
			 case 69 : event.keyCode=0;event.returnValue=false;Edit();break;
			 case 68 : Recely();break;
			 case 70 : event.keyCode=0;event.returnValue=false;parent.frames['LeftFrame'].initializeSearch('<%=KS.C_S(ChannelID,1)%>',<%=ChannelID%>,<%=KS.C_S(ChannelID,6)%>)
		   }	
		else if (event.keyCode==46) Delete();
		}
		</script>
		<%
		.echo "</head>"
		.echo "<body scroll=no onkeydown=""GetKeyDown();"" onselectstart=""return false;"" style='overflow:auto!important;overflow:hidden;background:#fff;'>"
		.echo "<ul id='menu_top'>"
		If ComeFrom="RecycleBin" Then
		 .echo "<li class='parent' onclick='BackRecely()'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/reb.gif' border='0' align='absmiddle'>批量还原</span></li>"
		 .echo "<li class='parent' onclick='Delete()'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>彻底删除</span></li>"
		 .echo "<li class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'"" onclick='DelAll()'><img src='images/ico/recyclebin.gif' border='0' align='absmiddle'>一键清空回收站</span></li>"
		ElseIf ComeFrom="Verify" Then
		    If Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10012") Then 
		    Call KS.ReturnErr(1, "")
			End If

		 .echo "<li class='parent' onclick='VerificAll()'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>批量审核</span></li>"
		 .echo "<li class='parent' onclick='Tuigao()'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/s.gif' border='0' align='absmiddle'>批量退稿</span></li>"
		 .echo "<li class='parent' onclick='Recely()'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/recycl.gif' border='0' align='absmiddle'>放入回收站</span></li>"
		 .echo "<li class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'"" onclick='Delete()'><img src='images/ico/del.gif' border='0' align='absmiddle'>彻底删除</span></li>"
		Else
		.echo "<li class='parent' onclick='CreateNews();'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>添加" & KS.C_S(ChannelID,3) & "</span></li>"
		.echo "<li class='parent' onclick='VerifyInfo();'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>签收" & KS.C_S(ChannelID,3) & "</span></li>"
		.echo "<li class='parent' onclick='Recely()'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/recycl.gif' border='0' align='absmiddle'>放入回收站</span></li>"
		.echo "<li class='parent' onclick='Delete()'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>彻底删除</span></li>"
		.echo "<li class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'"" title=""批量设置属性"" onclick=""SetAttribute();""><img src='images/ico/set.gif' border='0' align='absmiddle'>设置属性</span></li>"
		.echo "<li class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'"" title=""批量移动""  onClick=""MoveToClass();""><img src='images/ico/move.gif' border='0' align='absmiddle'>批量移动</span></li>"
		.echo "<li class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'"" title=""加入自由JS"" onclick=""AddToJS();""><img src='images/ico/addjs.gif' border='0' align='absmiddle'>加入JS</span></li>"
		.echo "<li class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'"" title=""加入专题""  onClick=""AddToSpecial();""><img src='images/ico/as.gif' border='0' align='absmiddle'>加入专题</span></li>"
        End If
			.echo "<li></li><div><select OnChange=""location.href='KS.ItemInfo.asp?ComeFrom=" & ComeFrom & "&ChannelID=" & ChannelID & "&id='+this.value;$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=ViewFolder&FolderID='+this.value;"" style='width:150px' name='id'>"
			.echo "<option value=''>快速跳转至...</option>"
			.echo Replace(KS.LoadClassOption(ChannelID),"value='" & ID & "'","value='" & ID &"' selected") & " </select>"

			
		   .echo "</div>"
		   .echo (" </ul>")
		
		
			.echo ("<div style=""height:94%; overflow: auto; width:100%"">")
		 If KeyWord<>"" or (StartDate <> "" And EndDate <> "") Then
		 .echo ("<img src='Images/ico/search.gif' align='absmiddle'> 搜索结果: ")
				 If StartDate <> "" And EndDate <> "" Then
					.echo (KS.C_S(ChannelID,3) & "更新日期在 <font color=red>" & StartDate & "</font> 至 <font color=red> " & EndDate & "</font>&nbsp;&nbsp;&nbsp;&nbsp;")
				 End If
				 If  KeyWord<>"" Then
				   Select Case SearchType
					Case 0:.echo ("文档标题中含有 <font color=red>" & KeyWord & "</font> 的" & KS.C_S(ChannelID,3))
					Case 1:.echo ("文档录入员中含有 <font color=red>" & KeyWord & "</font> 的" & KS.C_S(ChannelID,3))
					Case 2:.echo ("文档关键字中含有<font color=red>" & KeyWord & "</font> 的" & KS.C_S(ChannelID,3))
					Case 3:.echo ("文档作者含有<font color=red>" & KeyWord & "</font> 的" & KS.C_S(ChannelID,3))
				  End Select
			     End If
		End If
		
		 If .G("ComeFrom")="RecycleBin" Then
		  ShowChannelList 
		 Else
	      ShowClassList ChannelID,ID
		 End If
		
		 .echo ("<table border='0' width='100%' cellspacing='0' cellpadding='0'><tr><td><div style='margin:5px'><b>查看：</b>")
		 .echo ("<a href='?ChannelID=" & ChannelID & "&ComeFrom=" & ComeFrom & "'><font color=#999999>全部</font></a> - ")
		 If ComeFrom="Verify" Then
		 .echo ("<a href='?ChannelID=" & ChannelID & "&Verific=0&ComeFrom=" & ComeFrom & "'><font color=#999999>待审核" & KS.C_S(ChannelID,3) & "</font></a> - <a href='?ChannelID=" & ChannelID & "&Verific=2&ComeFrom=" & ComeFrom & "'><font color=#999999>会员草稿的" & KS.C_S(ChannelID,3) & "</font></a> - <a href='?ChannelID=" & ChannelID & "&Verific=3&ComeFrom=" & ComeFrom & "'><font color=#999999>被退稿的" & KS.C_S(ChannelID,3) & "</font></a></div></td><td align='right'>")
		 Else
		 .echo ("<a href='?ChannelID=" & ChannelID & "&status=1&ComeFrom=" & ComeFrom & "'><font color=#999999>推荐</font></a> - <a href='?ChannelID=" & ChannelID & "&status=2&ComeFrom=" & ComeFrom & "'><font color=#999999>幻灯</font></a> - <a href='?ChannelID=" & ChannelID & "&status=3&ComeFrom=" & ComeFrom & "'><font color=#999999>热门</font></a> - <a href='?ChannelID=" & ChannelID & "&status=4&ComeFrom=" & ComeFrom & "'><font color=#999999>固顶</font></a> - <a href='?ChannelID=" & ChannelID & "&status=5&ComeFrom=" & ComeFrom & "'><font color=#999999>评论</font></a> - <a href='?ChannelID=" & ChannelID & "&status=6&ComeFrom=" & ComeFrom & "'><font color=#999999>头条</font></a> - <a href='?ChannelID=" & ChannelID & "&status=7&ComeFrom=" & ComeFrom & "'><font color=#999999>滚动</font></a>")
		 If KS.C_S(ChannelID,6)=1 Then
		 .echo (" - <a href='?ChannelID=" & ChannelID & "&status=10&ComeFrom=" & ComeFrom & "'><font color=#999999>签收</font></a>")
		 End If
		  If ChannelID=5 Then 
		   .echo " - <a href='?ChannelID=" & ChannelID & "&status=8&ComeFrom=" & ComeFrom & "'><font color=green>涨价↑</font></a>"
		   .echo " - <a href='?ChannelID=" & ChannelID & "&status=9&ComeFrom=" & ComeFrom & "'><font color=#ff3300>降价↓</font></a>"
		  End If
		 .echo ("</div></td><td align='right'>")
		 End If
		 .echo("<b>" & KS.C_S(ChannelID,1) & "</b>[共有 <font color=red>" & Conn.Execute("select count(id) from " & KS.C_S(ChannelID,2) & " where verific=1")(0) & "</font> " & KS.C_S(ChannelID,4) & " 回收站 <font color=blue>" &Conn.Execute("select count(id) from " & KS.C_S(ChannelID,2) & " where verific=1 and deltf=1")(0)  &"</font> "& KS.C_S(ChannelID,4) & "]</td></tr></table>")
		 .echo ("<table width=""100%"" align='center' border=""0"" cellpadding=""0"" cellspacing=""0"">")
		 .echo ("<form name='myform' method='Post' action='?channelid="& channelid & "'>")
		 .echo ("<input type='hidden' name='action' id='action' value='" & Action & "'>")
		 .echo ("<input type='hidden' name='SetAttributeBit' id='SetAttributeBit' value='0'>")
		 .echo ("<tr align=""center"" class=""sort"">")
		 .echo ("<td width='35' align='center'>选择</td>")
		 If ChannelID=8 Then
		  .echo ("<td width='60'>类型</td>")
		 End If
		 .echo ("<td height=15>标题</td><td width=100>录 入</td><td width=80>修改日期</td><td width=60> 类 型 </td><td width=100> 属 性 </td>")
		 If ComeFrom="" Then
		 .Echo ("<td width='60'>点 击</td><td> 操 作 </td></tr>")
		 Else
		 .Echo ("<td width='60'>状 态</td><td> 操 作 </td></tr>")
		 End If

		   Dim Param
		   If ComeFrom="RecycleBin" Then
		    Param = Param & " DelTF=1"
		   ElseIf ComeFrom="Verify" Then
		    Param = Param & " DelTF=0 And Verific=" & KS.ChkClng(KS.G("Verific"))
		   Else
		    Param = Param & " DelTF=0  And Verific=1"
		   End If
		   
		   '非超级管理员，只能管理自己添加的信息
		   'If KS.C("SuperTF")<>"1" Then	 Param=Param & " and inputer='" & KS.C("AdminName") & "'"
		   
		    If KS.C("SuperTF")<>"1" and Instr(KS.C("ModelPower"),KS.C_S(ChannelID,10)&"1")=0 Then 
			 If DataBaseType=1 Then
			 Param=Param & " and tid in(select id from ks_class where ','+cast(AdminPurview as nvarchar(4000))+',' like '%," & KS.C("AdminName") & "%')"
			 Else
			 Param=Param & " and tid in(select id from ks_class where ','+AdminPurview+',' like '%," & KS.C("AdminName") & "%')"
			 End If
			End If
		  

		   If KeyWord <> "" or (StartDate <> "" And EndDate <> "") Then
		        If KeyWord<>"" Then
				Select Case SearchType
				  Case 0:Param = Param & " And (Title like '%" & KeyWord & "%')"
				  Case 1:Param = Param & " And Inputer like '%" & KeyWord & "%'"
				  Case 2:Param = Param & " And KeyWords like '%" & KeyWord & "%'"
				  Case 3:Param = Param & " And Author like '%" & KeyWord & "%'"
				End Select
				End If
				If StartDate <> "" And EndDate <> "" Then
					If CInt(DataBaseType) = 1 Then         'Sql
					   Param = Param & " And (AddDate>= '" & StartDate & "' And AddDate<= '" & DateAdd("d", 1, EndDate) & "')"
					Else                                                 'Access
					   Param = Param & " And (AddDate>=#" & StartDate & "# And AddDate<=#" & DateAdd("d", 1, EndDate) & "#)"
					End If
				End If
		  Else
		    if (ID<>"0") then Param = Param & " And Tid In (" & KS.GetFolderTid(ID) & ")" 
			select case KS.ChkClng(KS.S("Status"))
			 case 1 Param = Param & " And Recommend=1"
			 case 2 Param = Param & " And Slide=1"
			 case 3 Param = Param & " And Popular=1"
			 case 4 Param = Param & " And IsTop=1"
			 case 5 Param = Param & " And Comment=1"
			 case 6 Param = Param & " And Strip=1"
			 case 7 Param = Param & " And Rolls=1"
			 case 8 Param = Param & " And ProductType=2"
			 case 9 Param = Param & " And ProductType=3"
			 case 10 Param = Param &" And IsSign=1"
			end select
			
		  End If
		  
		 
		
		Dim FieldStr
		If ChannelID=5 Then
		 FieldStr="ID,Tid,Title,Inputer,AddDate,PhotoUrl,Verific,Recommend,Popular,Strip,Rolls,Slide,IsTop,Hits,ProductType"
		ElseIf ChannelID=8 Then
		 FieldStr="ID,Tid,Title,Inputer,AddDate,PhotoUrl,Verific,Recommend,Popular,Strip,Rolls,Slide,IsTop,Hits,TypeID"
		Else
		 FieldStr="ID,Tid,Title,Inputer,AddDate,PhotoUrl,Verific,Recommend,Popular,Strip,Rolls,Slide,IsTop,Hits"
		End If
		If KS.ChkClng(KS.S("Status"))=10 Then
		 FieldStr=FieldStr & ",SignUser"
		End If
		SQLStr=KSCls.GetPageSQL(KS.C_S(ChannelID,2),"id",MaxPerPage,Page,1,Param,FieldStr)
		
		Set RS = Server.CreateObject("AdoDb.RecordSet")
		 RS.Open SQLStr, conn, 1, 1
			 If Not RS.EOF Then
					totalPut = Conn.Execute("Select count(id) from [" & KS.C_S(ChannelID,2) & "] where " & Param)(0)
					if (TotalPut mod MaxPerPage)=0 then
						TotalPages = TotalPut \ MaxPerPage
					else
						TotalPages = TotalPut \ MaxPerPage + 1
					end if
					if Page > TotalPages then Page=TotalPages
					Dim SQL:SQL=RS.GetRows(MaxPerPage)
					Call showContent(SQL)
							
			 Else
			  .echo "<tr><td colspan=8 align='center' height='35' class='splittd'><font color=red>对不起，没有找到任何" &KS.C_S(ChannelID,3) & "!</font></td></tr>"
			 End If
			  
			  .echo ("<table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>")
			  .echo ("<tr><td width='180'><div style='margin:5px'><b>选择：</b><input type='button' value='全选' onclick='javascript:Select(0)' class='button'>  <input type='button' value='反选' onclick='javascript:Select(1)' class='button'>  <input type='button' value='不选' onclick='javascript:Select(2)' class='button'> </div>")
			  .echo ("</td>")
			  .echo ("<td><td align='right'>")
			  
		If ComeFrom="RecycleBin" Then
			  .echo ("<table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>")
			  .echo ("<tr><td style='padding-left:20px'>")
			  .echo ("<input type=""button"" value=""批量还原"" onclick=""BackRecely()"" class=""button"">")
			  .echo (" <input type=""button"" value=""彻底删除"" onclick=""Delete()"" class=""button"">")
			  .echo (" <input type=""button"" value=""一键清空"" onclick=""DelAll()"" class=""button"">")
			  .echo ("</td></tr>")
			  .echo ("</table>")
		Else
			  .echo ("<table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>")
			  .echo ("<tr><td width='49%' align='center'>")
			  .echo ("<fieldset align=center><legend>设定</legend>")
			  .echo ("<input type=""button"" value=""推荐"" onclick=""set(this,1)"" class=""button"">")
			  .echo (" <input type=""button"" value=""幻灯"" onclick=""set(this,2)"" class=""button"">")
			  .echo (" <input type=""button"" value=""热门"" onclick=""set(this,3)"" class=""button"">")
			  .echo (" <input type=""button"" value=""评论"" onclick=""set(this,4)"" class=""button"">")
			  .echo (" <input type=""button"" value=""头条"" onclick=""set(this,5)"" class=""button"">")
			  .echo (" <input type=""button"" value=""固顶"" onclick=""set(this,6)"" class=""button"">")
			  .echo (" <input type=""button"" value=""滚动"" onclick=""set(this,7)"" class=""button"">")
			  
			  .echo ("</fieldset>")
			  .echo ("</td><td width='2%'></td><td width='49%' align='center'>")
			  .echo ("<fieldset align=center><legend>取消</legend>")
			  .echo ("<input type=""button"" value=""推荐"" onclick=""set(this,8)"" class=""button"">")
			  .echo (" <input type=""button"" value=""幻灯"" onclick=""set(this,9)"" class=""button"">")
			  .echo (" <input type=""button"" value=""热门"" onclick=""set(this,10)"" class=""button"">")
			  .echo (" <input type=""button"" value=""评论"" onclick=""set(this,11)"" class=""button"">")
			  .echo (" <input type=""button"" value=""头条"" onclick=""set(this,12)"" class=""button"">")
			  .echo (" <input type=""button"" value=""固顶"" onclick=""set(this,13)"" class=""button"">")
			  .echo (" <input type=""button"" value=""滚动"" onclick=""set(this,14)"" class=""button"">")
			  .echo ("</fieldset>")
			  .echo ("</td></tr>")
			  .echo ("</table>")
		  End If
			  
			  .echo ("</td></tr></form></table>")
			  
			 
			  .echo ("<table border='0' width='100%'><tr>")
			  If KS.C_S(ChannelID,7)<>0 Then
			  .echo ("<td align='center'><input class='button' onclick='CreateHtml()' type='button' value='一键发布选中'></td>")
			  End If
			  .echo ("<td>")
			  Call KS.ShowPage(totalput, MaxPerPage, "", Page,true,true)
			  ' Call KSCLS.ShowPage(totalPut, MaxPerPage, "KS.ItemInfo.asp", True, KS.C_S(ChannelID,4), Page, KS.QueryParam("page"))
			  .echo ("</td></tr></table>")

			  
			  .echo ("<table style='clear:both;background:url(images/ico/bottombg.gif);margin-top:5px;' height='43' border='0' width='100%' align='center'>")
			  .echo ("<form action='KS.ItemInfo.asp' name='searchform' method='get'>")
			  .echo ("<tr><td><img src='images/ico/search.gif' align='absmiddle'>搜索：")
			  .echo ("&nbsp;搜索类型 <select name='searchtype'>")
			  If SearchType="0" Then .echo ("<option value=0 selected>文档标题</option>") Else .echo ("<option value=0>文档标题</option>")
			  If SearchType="1" Then .echo ("<option value=1 selected>文档录入员</option>") Else .echo("<option value=1>文档录入员</option>")
			  If SearchType="2" Then .echo ("<option value=2 selected>文档关键字</option>") Else .echo ("<option value=2>文档关键字</option>")
			  If SearchType="3" Then .echo ("<option value=3 selected>文档作者</option>") Else .echo ("<option value=3>文档作者</option>")
			  .echo ("</select> <input type='text' title='关键字可留空' class='textbox' value='" & KeyWord &"' size='12' name='keyword'>&nbsp;<span style='cursor:pointer' onclick='setstatus(this)'><b>修改日期？</b></span>")
			  If StartDate <> "" And EndDate <> "" Then
			  .echo ("<span id='SearchDate'>开始日期<input onClick=""OpenThenSetValue('Include/DateDialog.asp',160,170,window,$('#StartDate')[0]);$('#StartDate').focus();"" type='text' size='12' readonly  name='StartDate' value='" & StartDate & "' style='cursor:pointer'  id='StartDate'>&nbsp;结束日期<input type='text' readonly size=12 value='" & EndDate & "' name='EndDate' id='EndDate' style='cursor:pointer'  onClick=""OpenThenSetValue('Include/DateDialog.asp',160,170,window,$('#EndDate')[0]);$('#EndDate').focus();""></span>")
			  Else
			  .echo ("<span style='display:none' id='SearchDate'>开始日期<input onClick=""OpenThenSetValue('Include/DateDialog.asp',160,170,window,$('#StartDate')[0]);$('#StartDate').focus();"" type='text' size='12' readonly  name='StartDate' style='cursor:pointer'  id='StartDate'>&nbsp;结束日期<input type='text' readonly size=12 name='EndDate' id='EndDate' style='cursor:pointer'  onClick=""OpenThenSetValue('Include/DateDialog.asp',160,170,window,$('#EndDate')[0]);$('#EndDate').focus();""></span>")
			  End If
			  .echo ("&nbsp;<input type='submit' class='button' value='开始搜索'><input type='hidden' value='" & ChannelID & "' name='channelid'><input type='hidden' value='" & ComeFrom & "' name='ComeFrom'></td>")
			  .echo ("</tr>")
			  .echo ("</form>")
			  .echo ("</table>")
		  .echo ("</div>")
		  .echo ("</body>")
		  .echo ("</html>")
		  End With
		  Set RS = Nothing
		End Sub

      Sub ShowClassList(ChannelID,ID)
		 If KS.S("ComeFrom")<>"" Then Exit Sub
		 
		 With KS
		 '============增加记忆功能=======================================
		 Dim ExtStatus,CloseDisplayStr,ShowDisplayStr,classExtStatus
		 classExtStatus=request.cookies("classExtStatus")
		 if classExtStatus="" Then classExtStatus=1
		 If classExtStatus=1 Then 
		  ExtStatus=2 :CloseDisplayStr="display:none;":ShowDisplayStr=""
		 Else 
		  ExtStatus=1 :CloseDisplayStr="":ShowDisplayStr="display:none;"
		 End If
		 '=========================================================----
		 .echo "<div id='classOpen' onclick=""ClassToggle("& ExtStatus& ")"" style='" & CloseDisplayStr &"cursor:pointer;text-align:center;position:absolute; z-index:2; left: 0px; top: 2px;' ><img src='images/kszk.gif' align='absmiddle'></div>"
		 .echo "<div id='classNav' style='" & ShowDisplayStr &"position:relative;height:auto;_height:30px;line-height:30px;margin:5px 1px;border:1px solid #DEEFFA;background-color:#F7FBFE'>"
		 .echo "<div style='padding-top:2px;cursor:pointer;text-align:center;position:absolute; z-index:1; right: 0px; top: 2px;'  onclick=""ClassToggle(" & ExtStatus &")""> <img src='images/close.gif' align='absmiddle'></div>"
		
		Dim P,RSC,Img,j,N,I,XML,Node
		P=" where ClassType=1 and ChannelID=" & ChannelID
		If ID=0 Then
		  P=P & " And tj=1"
		 Img="domain.gif"
		Else
		 P=P & " And TN='" & ID & "'"
		 Img="Smallfolder.gif"
		End If

		 on error resume next
		 Dim ParentID:ParentID = conn.Execute("Select TN From KS_Class  Where ID='" & ID & "'")(0)

		Set RSC=Conn.Execute("select id,foldername,adminpurview from ks_class " & P& " order by root,folderorder")
		If Not RSC.Eof Then 
		 Set XML=.RsToXml(RSC,"row","xmlroot")
		 RSC.Close:Set RSC=Nothing
		 If IsObject(XML) Then
		   If ID<>"0" Then
		    .echo "<a href='?ChannelID=" & ChannelID & "&ID=" & ParentID & "'><img src='images/folder/Back.gif' border=0 align='absmiddle'></a> "
		   End if
		   For Each Node In XML.DocumentElement.SelectNodes("row")
		    If KS.C("SuperTF")=1 or KS.FoundInArr(Node.SelectSingleNode("@adminpurview").text,KS.C("AdminName"),",") or Instr(KS.C("ModelPower"),KS.C_S(ChannelID,10)&"1")>0 Then 
		    .echo "<li style='margin:5px;float:left;width:100px'><img src='images/folder/" & Img & "' align='absmiddle'><a href='?ChannelID=" & ChannelID & "&ID=" & Node.SelectSingleNode("@id").text & "' title='" & Node.SelectSingleNode("@foldername").text & "'>" & .Gottopic(Node.SelectSingleNode("@foldername").text,8) & "</a></li>"
		    End If
		   Next
		 End If
		Else
		  If err Then
		   .echo "<img src='images/folder/AddFolder.gif' align='absmiddle'>请先<a href='#' onclick=""location.href='KS.Class.asp?Action=Add&ChannelID=" & ChannelID & "';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Go&OpStr=" & Server.URLEncode("栏目管理 >> <font color=red>添加栏目</font>") & "';"">添加栏目</a>"
		  Else
		   .echo "<a href='?ChannelID=" & ChannelID & "&ID=" & ParentID & "'><img src='images/folder/Back.gif' border=0 align='absmiddle'></a> <a href='#' onclick='CreateNews()'><font color=#4169E1><strong>添加" & KS.C_S(Channelid,3) & "</strong></font></a>"
		   End If
		End If
		 .echo "</div>"
		 .echo "<div style=""clear:both""></div>"
		 End With
		End Sub
		
		Sub ShowChannelList()
		  With KS
			 '============带记忆功能=======================================
			 Dim ExtStatus,CloseDisplayStr,ShowDisplayStr,classExtStatus
			 classExtStatus=request.cookies("classExtStatus")
			 if classExtStatus="" Then classExtStatus=1
			 If classExtStatus=1 Then 
			  ExtStatus=2 :CloseDisplayStr="display:none;":ShowDisplayStr=""
			 Else 
			  ExtStatus=1 :CloseDisplayStr="":ShowDisplayStr="display:none;"
			 End If
			 '=========================================================----
			 .echo "<div id='classOpen' onclick=""ClassToggle("& ExtStatus& ")"" style='" & CloseDisplayStr &"cursor:pointer;text-align:center;position:absolute; z-index:2; left: 0px; top: 2px;' ><img src='images/kszk.gif' align='absmiddle'></div>"
			 .echo "<div id='classNav' style='" & ShowDisplayStr &"position:relative;height:auto;_height:30px;line-height:30px;margin:5px 1px;border:1px solid #DEEFFA;background:#F7FBFE'>"
			 .echo "<div style='padding-top:2px;cursor:pointer;text-align:center;position:absolute; z-index:1; right: 0px; top: 2px;'  onclick=""ClassToggle(" & ExtStatus &")""> <img src='images/close.gif' align='absmiddle'></div>"
			 If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
				Dim ModelXML,Node
				Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
				For Each Node In ModelXML.documentElement.SelectNodes("channel")
				 if Node.SelectSingleNode("@ks21").text="1" and Node.SelectSingleNode("@ks0").text<>"6" and Node.SelectSingleNode("@ks0").text<>"9" and Node.SelectSingleNode("@ks0").text<>"10" Then
				   .echo "<li style='margin:5px;float:left;width:100px'><img src='images/folder/domain.gif' align='absmiddle'><a href='?ChannelID=" & Node.SelectSingleNode("@ks0").text & "&ComeFrom=RecycleBin' title='" & Node.SelectSingleNode("@ks1").text & "'>" & .Gottopic(Node.SelectSingleNode("@ks1").text,8) & "(<span style='color:red'>" & Conn.Execute("Select Count(ID) From " & Node.SelectSingleNode("@ks2").text & " Where Deltf=1")(0) & "</span>)</a></li>"
			    End If
			next
			.echo "</div>"
			.echo "<div style=""clear:both""></div>"
         End With
		End Sub


		Sub showContent(SQL)
		    Dim ItemIcon
			With KS
			For I=0 To Ubound(SQL,2)
					If SQL(5,I) <>"" Then
						 ItemIcon="Images/ico/doc1.gif"
					Else
						 ItemIcon="Images/ico/doc0.gif"
					End If
					    AttributeStr = ""
						If SQL(7,I) = 1 Or SQL(8,I) = 1 Or SQL(9,I) = 1 Or SQL(10,I) = 1 Or SQL(11,I) = 1 Or SQL(12,I) = 1 Then
								  If SQL(7,I) = 1 Then AttributeStr = AttributeStr & (" <span title=""推荐" & KS.C_S(ChannelID,3) & """ style=""cursor:default""><font color=""green"">荐</font></span>&nbsp;")
								  If SQL(8,I) = 1 Then AttributeStr = AttributeStr & ("<span title=""热门" & KS.C_S(ChannelID,3) & """ style=""cursor:default""><font color=""red"">热</font></span>&nbsp;")
								  If SQL(9,I) = 1 Then AttributeStr = AttributeStr & ("<span title=""今日头条"" style=""cursor:default""><font color=""#0000ff"">头</font></span>&nbsp;")
								  If SQL(10,I) = 1 Then AttributeStr = AttributeStr & ("<span title=""滚动" & KS.C_S(ChannelID,3) & """ style=""cursor:default""><font color=""#F709F7"">滚</font></span>&nbsp;")
								  If SQL(11,I) = 1 Then AttributeStr = AttributeStr & ("<span title=""幻灯片" & KS.C_S(ChannelID,3) & """ style=""cursor:default""><font color=""black"">幻</font></span>")
								  IF SQL(12,I) = 1 Then AttributeStr = AttributeStr & ("<span title=""固顶" & KS.C_S(ChannelID,3) & """ style=""cursor:default""><font color=""brown"">固</font></span>")
					   Else
								AttributeStr = "---"
					   End If
					   
					If KS.ChkClng(KS.G("Status"))=10 Then
					   Dim RSS,HasSignUser,XML,Node,MustSignUserArr,SignUser,NoSignUser,S,AttrStr
					   Set RSS=Conn.Execute("Select top 500 username From KS_ItemSign Where ChannelID=" & ChannelID & " and infoid=" & SQL(0,I))
					   If Not RSS.EOf Then
						   SET xml=KS.RsToXml(RSS,"row","")
						   for each node in xml.documentelement.selectnodes("row")
							 if HasSignUser="" then 
							   HasSignUser=node.selectSingleNode("@username").text
							 else
							   HasSignUser=HasSignUser& "," & node.selectSingleNode("@username").text
							 end if
						   next
					   End If
					   RSS.Close
					   
					   SignUser=SQL(14,I) : NoSignUser="" : MustSignUserArr=Split(SignUser,",")
					   If IsArray(MustSignUserArr) Then
					   For S=0 To Ubound(MustSignUserArr)
						  If KS.FoundInArr(HasSignUser,MustSignUserArr(S),",")=false Then
							if NoSignUser="" then
							  NoSignUser=MustSignUserArr(S)
							else
							  NoSignUser=NoSignUser & "," & MustSignUserArr(S)
							end if
						  End If
					   Next
					   End If
					   If NoSignUser="" Then AttrStr="<font color=blue>签收完毕</font>" Else AttrStr="<font color=red>签收中...</font>"
					   TitleStr =" title='已签收用户:" & HasSignUser & "&#13;&#10;未签收用户:"& NoSignUser &"'"
					Else
                     TitleStr = " TITLE='名 称:" & SQL(2,I) & "&#13;&#10;日 期:" & SQL(4,I) & "&#13;&#10;录 入:" & SQL(3,I) & "'"
					End If
							 .echo ("<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" & SQL(0,I) & "' onclick=""chk_iddiv('" & SQL(0,I) & "')"">")
							 .echo ("<td class='splittd' align=center><input name='id'  onclick=""chk_iddiv('" & SQL(0,I) & "')"" type='checkbox' id='c"& SQL(0,I) & "' value='" &SQL(0,I) & "'></td>")
							 If ChannelID=8 Then
							 .echo ("<td align=""center"" class='splittd'>" & KS.GetGQTypeName(SQL(14,I)) & "</td>")
							 End If
							 .echo ("<td" & TitleStr & " class='splittd'><span onDblClick=""ViewArticle(" & SQL(0,I) &")"">")
							 .echo ("<a href='javascript:ViewArticle(" & SQL(0,I) & ");'><img src=" & ItemIcon & " border=0 align=absmiddle title='预览'></a>")
							 .echo ("<span style=""cursor:default""><a href='?ID=" & SQL(1,I) &"&channelid=" & ChannelID&"'>[" & KS.C_C(SQL(1,I),1) &"]</a> "& KS.Gottopic(SQL(2,I),27)) & AttrStr
							 If ChannelID=5 Then
							  if SQL(14,I)=2 Then .echo "<span title='涨价出售' style='color:green'> ↑</span>" Else If SQL(14,I)=3 Then .echo "<span style='color:red' title='降价出售'> ↓</span>"
							 End If
							 .echo ( "</span></span></td>")
							 .echo ("<td align=""center"" class='splittd'>" & SQL(3,I) & "</td>")
							 .echo ("<td align=""center"" class='splittd'>" & FormatDateTime(SQL(4,I),2) & "</td>")
							 .echo ("<td align=""center"" class='splittd'>" & KS.C_S(ChannelID,3) & "</td>")
							 .echo ("<td align=""center"" class='splittd'>" & AttributeStr & "</td>")
							 .echo ("<td align=""center"" class='splittd'>")
							 
							  If ComeFrom="" Then
							    .echo SQL(13,I)
							  Else
							   Select Case SQL(6,I)
								  Case 0: .echo "<span style='color:red'>待审</span>"
								  Case 1: .echo "<span style='color:blue'>已审</span>"
                                  Case 2: .echo "<span style='color:#999999'>草稿</span>"
                                  Case 3: .echo "<span style='color:green'>退稿</span>"
                               End Select
							  End If
							 .echo ("</td>")
							 
							 .echo ("<td align=""center"" class='splittd'>")
							 If ComeFrom="RecycleBin" Then
							 .echo("<a href='?Page=" & Page & "&Action=RecelyBack&" &SearchParam&"&ID=" & SQL(0,I) & "'>还原</a> | <a href=""?Action=Delete&Page=" & Page & "&" & SearchParam & "&ID=" & SQL(0,I) & """ onclick=""return (confirm('此操作不可逆，确定将该" & KS.C_S(ChannelID,3) & "彻底删除吗?'))"">彻删</a>")
							 ElseIf ComeFrom="Verify" Then
							  If SQL(6,I) =2  Then
							  .echo "<font color=#cccccc>不允许操作</font>"	  
							  Else
								 If SQL(6,I) <>3  Then   '已审核或草稿文章不允许操作
								   .echo "  <a href=""#""  onclick=""$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ChannelID=" & ChannelID & "&ComeFrom=Verify&ButtonSymbol=AddInfo&OpStr=" & server.URLEncode(KS.C_S(ChannelID,3) & "管理 >> <font color=red>签收会员" & KS.C_S(ChannelID,3)) & "</font>';location.href='" & ItemManageUrl & "?ChannelID=" & ChannelID & "&Page=" & Page & "&Action=Verify&ID="&SQL(0,I)&"';"">审核</a>"
								  If SQL(6,I)<>2 Then
								 .echo "&nbsp;<a onclick=""ProcessTuigao(event," & SQL(0,I) & ")"" href='#'>退稿</a>"
								  End IF
								 End If
								 .echo (" <a href=""?Action=Recely&Page=" & Page & "&" & SearchParam & "&ID=" & SQL(0,I) & """ onclick=""return (confirm('确定将该" & KS.C_S(ChannelID,3) & "放入回收站吗?'))"">回收站</a>")
								End If
							 Else
							 .echo (" <a href='" & ItemManageUrl & "?Page=" & Page & "&Action=Edit&" &SearchParam&"&ID=" & SQL(0,I) & "' onclick='parent.frames[""BottomFrame""].location.href=""KS.Split.asp?ChannelID=" & ChannelID &"&ComeFrom="&ComeFrom&"&OpStr="&Server.URLEncode("编辑" & KS.C_S(ChannelID,3)) & "&ButtonSymbol=AddInfo&FolderID=" & ID & """;'>修改</a> | <a href=""?Action=Recely&Page=" & Page & "&" & SearchParam & "&ID=" & SQL(0,I) & """ onclick=""return (confirm('确定将该" & KS.C_S(ChannelID,3) & "放入回收站吗?'))"">回收站</a>")
							 If ChannelID=5 Then
							 .echo (" | <a href='KS.ProImages.asp?ProID=" &SQL(0,I) &"&ChannelID=5'>图片</a>")
							 End If
							 End If
							 .echo ("</td>")
							 .echo ("</tr>")
			  Next

			  .echo ("</table>")
			End With
		End Sub
	
		
		'加入JS
		Sub AddToJS()
		    DIM JSNameList,JSObj,NewsID
			NewsID=Trim(Request("NewsID"))
			 Set JSObj=Server.CreateObject("Adodb.Recordset")
			 JSObj.Open "Select JSName,JSID From KS_JSFile Where JSType=1 And JSConfig NOT LIKE 'GetExtJS%'",Conn,1,1
			 IF NOT JSObj.EOF THEN
				 JSNameList="<Option Value='0'></Option>"
			  DO While NOT JSObj.EOF 
				 JSNameList=JSNameList & "<Option value=" & JSObj("JSID") &">" & Trim(JSObj("JSName")) & "</Option>"
				 JSObj.MoveNext
			  LOOP
			 Else
				 JSNameList=JSNameList & "<Option value=0>---您没有建自由JS---</Option>"
			 END IF
			JSObj.Close:Set JSObj=Nothing
			%>  
			<html>
			<head>
			<meta http-equiv="Content-Type" content="text/html; chaRSet=gb2312">
			<title>加入自由JS</title>
			<link href="Include/Admin_Style.css" rel="stylesheet">
			<link href="Include/ModeWindow.css" rel="stylesheet">
			<script language="JavaScript" src="../KS_Inc/common.js"></script>
			</head>
			<body topmargin="0" leftmargin="0" scroll=no>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
			  <form name="myform" action="?ChannelID=<%=ChannelID%>&Action=JS" method="post">
			  <input type="hidden" value="Add" Name="Flag">
			  <input type="hidden" name="JSName">
			  <input type="hidden" value="<%=NewsID%>" Name="NewsID"> 
			  <tr> 
				<td height="18">&nbsp;</td>
			  </tr>
			  <tr> 
				<td height="30" align="center"> <strong>请选择JS名称</strong> 
				  <select name="JSID">
					  <%=JSNameList%>
				  </select>
				</td>
			  </tr>
			  <tr align="center"> 
				<td height="30"> <input type="button" class="button" name="button1" value="加入JS" onClick="CheckForm()"> 
				  &nbsp; <input type="button" class="button" onClick="window.close();" name="button2" value=" 取消 "> 
				</td>
			  </tr>
			  </form>
			</table>
			</body>
			</html>
			<Script>
			function CheckForm()
			{
			 if (document.myform.JSID.value=='0')
			  { alert('对不起,您没有选择JS名称!');
				 document.myform.JSID.focus();
				 return false;
			  }
			  document.myform.JSName.value=document.myform.JSID.options[document.myform.JSID.selectedIndex].text
			  document.myform.submit();
			  return true
			}
			</Script> 
			<%IF Request.Form("Flag")="Add" Then
			   Dim RS,OldJSID,JSID,NewsIDArr,K
			   JSID=Trim(Request.Form("JSID"))
			   NewsIDArr=Split(NewsID,",")
			   Set RS=Server.CreateObject("Adodb.RecordSET")
			   For K=Lbound(NewsIDArr) To Ubound(NewsIDArr)
				  RS.Open "Select JSID From " & KS.C_S(ChannelID,2) &" Where ID=" & NewsIDArr(K),Conn,1,3
				  IF  Not RS.Eof THEN
						 OldJSID=Trim(RS("JSID"))
					   IF Trim(RS(0))="0" or Trim(RS(0))="" or isnull(RS(0)) Then
						  RS(0)=JSID & ","
					   Elseif InStr(OldJSID,JSID)=0 then
						  RS(0)=RS(0) & JSID & ","
					   End if
					   RS.UPDate
					   
					 End IF
                  RS.Close
			   Next
			            '刷新JS
					   Dim KSRObj,JSName
					   JSName=Trim(Request.Form("JSName"))
					   Set KSRObj=New Refresh
					   KSRObj.RefreshJS(JSName)
					   Set KSRObj=Nothing
			   Set RS=Nothing
			   KS.Echo "<script>alert('操作成功!');window.close();</script>"
			End IF
		End Sub
		
		
		'设置属性
		Sub SetAttribute()
		 Dim RS, IDArr, K
		 Dim ID:ID=Trim(Request("ID"))
		 Dim ChannelID:ChannelID=KS.ChkClng(Request("ChannelID"))
		 If ChannelID=0 Then ChannelID=1
		 %>
		 	<html>
			<head>
			<meta http-equiv="Content-Type" content="text/html; chaRSet=gb2312">
			<title>设置属性</title>
			<link href="Include/Admin_Style.css" rel="stylesheet">
			<script language="JavaScript" src="../KS_Inc/common.js"></script>
			<script language="JavaScript" src="../KS_Inc/Jquery.js"></script>
	        <script src="images/pannel/tabpane.js" language="JavaScript"></script>
	        <link href="images/pannel/tabpane.CSS" rel="stylesheet" type="text/css">
				 <script language="javascript">
				  $(document).ready(function(){
				   $("#channelids").change(function(){
					 if ($(this).val()!=0){
					  $(parent.frames["FrameTop"].document).find("#ajaxmsg").toggle();
					  $.get("../plus/ajaxs.asp",{action:"GetClassOption",channelid:$(this).val()},function(data){
						 $(parent.frames["FrameTop"].document).find("#ajaxmsg").toggle();
						 $("select[name=ClassID]").empty();
						 $("select[name=ClassID]").append(unescape(data));
						 $("input[name=ChannelID]").val($("#channelids").val());
						 if ($("input[name=ChannelID]").val()==5 || $("input[name=ChannelID]").val()==7 || $("input[name=ChannelID]").val()==8){
						  $("#showauthor").hide();
						  $("#showorigin").hide();
						  }else{
						  $("#showauthor").show();
						  $("#showorigin").show();
						  }
					   });
					 }
				   });
				  })

				function SelectAll(){
				  $("select[name=ClassID]>option").each(function(){
				   $(this).attr("selected",true);
				  })
				}
				function UnSelectAll(){
				  $("select[name=ClassID]>option").each(function(){
				   $(this).attr("selected",false);
				  })
				}
				</SCRIPT>			
           </head>
			<body topmargin="0" leftmargin="0" scroll="no">
			<div class="topdashed sort">批量设置文档属性</div>
			<div style="height:94%; overflow: auto; width:100%">
			<iframe src="about:blank" width="0" height="0" name="_hiddenframe" id="_hiddenframe" style="display:none"></iframe>
			<table width="100%" style="margin-top:10px" border="0" align="center"  cellspacing="1" class='ctable'>
			<form name="myform" action="?Action=SetAttribute" method="post" target="_hiddenframe">
			  <input type='hidden' name='ChannelID' id='ChannelID' value='<%=ChannelID%>'>
			  <input type="hidden" value="Add" Name="Flag">
			  <tr class='tdbg' id='choose2'<%if ID<>"" then response.write " style='display:none'"%>>
				<td valign='top' rowspan='100' width='200'>
				<font color=red>提示：</font>可以按住“Shift”<br />或“Ctrl”键进行多个栏目的选择<br />
				<%if ChannelID<>5 then%>
				<select id='channelids' name='channelids' style='width:200px'>
				 <option value='0'>---请选择模型---</option>
				 <%
				If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
				Dim ModelXML,Node
				Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
				For Each Node In ModelXML.documentElement.SelectNodes("channel")
				 if Node.SelectSingleNode("@ks21").text="1" and Node.SelectSingleNode("@ks0").text<>"6" and Node.SelectSingleNode("@ks0").text<>"9" and Node.SelectSingleNode("@ks0").text<>"10" Then
				  Response.write "<option value='" &Node.SelectSingleNode("@ks0").text &"'>" & Node.SelectSingleNode("@ks1").text & "</option>"
				 End If
				next
				%>
				</select>
				<%end if%>
				
			<Select style="WIDTH: 200px; HEIGHT: 380px" multiple size=2 name="ClassID">
			 <%=KS.LoadClassOption(ChannelID)%>
			</Select>
			<div align=center>
			   <Input onclick=SelectAll() type=button class="button" value="选定所有栏目" name=Submit><br />
			   <Input onclick=UnSelectAll() type=button value="取消选定栏目" class="button" name=Submit></div>
                </td>
			  </tr>
			  <tr class='tdbg'>
			     <TD valign="top">
				 
				   
				        <table border="0" width="100%" cellpadding="0" cellspacing="1" class="ctable">
				            <tr>
							 <td class='clefttitle' align='right'><strong>设置选择:</strong></td>
							 <td><input type=radio name=choose value='0'<%if ID<>"" then response.write" checked"%> onClick="choose1.style.display='';choose2.style.display='none';"> 按文档ID&nbsp;&nbsp;		<input type=radio name=choose value='1' onClick="choose2.style.display='';choose1.style.display='none';"<%if ID="" then response.write " checked" else response.write "disabled"%>> 按文档分类</td>
						  </tr>
						  <tr class='tdbg' id='choose1'<%if ID="" then response.write " style='display:none'"%>>
							 <td class='clefttitle' align='right'><strong>文档ID：</strong>多个ID请用“,”分开</td>
							 <td><input type='text' size='50' value='<%=ID%>' name='ID'></td>
						  </tr>
						</table>
						
						
				  <%if ChannelID=5 then%>
				    <script type="text/javascript">
					 function setPrice(p)
					 {
					   $("#groupprice").find("input").each(function(){
					     $(this).val(p);
					   });
					 
					   $("input[name='DiscountPriceMarket']").val(p);
					   $("input[name='DiscountPrice']").val(p);
					   $("input[name='DiscountPriceMember']").val(p);
					   $("input[name='DiscountScore']").val(p);
					 }
					function regInput(obj, reg, inputStr)
					{
						var docSel = document.selection.createRange()
						if (docSel.parentElement().tagName != "INPUT")    return false
						oSel = docSel.duplicate()
						oSel.text = ""
						var srcRange = obj.createTextRange()
						oSel.setEndPoint("StartToStart", srcRange)
						var str = oSel.text + inputStr + srcRange.text.substr(oSel.text.length)
						return reg.test(str)
					}
					</script>		
					<div class="tab-page" id="SetAttrPanel">
						<SCRIPT type=text/javascript>
							   var tabPane1 = new WebFXTabPane( document.getElementById( "SetAttrPanel" ), 1 )
						</SCRIPT>
								 
					<div class=tab-page id=price-page1>
					<H2 class=tab>批量调价</H2>
					<SCRIPT type=text/javascript>
						 tabPane1.addTabPage( document.getElementById( "price-page1" ) );
					</SCRIPT>
								
					 <table border="0" width="100%" cellpadding="0" cellspacing="1">
						  <tr class='tdbg'> 
						    
                            <td class='clefttitle' align='right' width="80"><strong>折扣标志:</strong></td>
							<td class='clefttitle' height='25'>
							<label onClick="$('#zkl').hide()"><input name='ProductType' type='radio' value='0' checked>不做设置</label>
							<br/>
							<label onClick="$('#zkl').show();setPrice(10)"><input name='ProductType' type='radio' value='1'>全部复位到原始零售价</label>
							<br/>
							<label style="color:green" onClick="$('#zkl').show();setPrice(11)"><input name='ProductType' type='radio' value='2'>设为涨价销售</label>
							<br/>
							<label style="color:red" onClick="$('#zkl').show();setPrice(9.8)"><input name='ProductType' type='radio' value='3'>设为降价销售</label>
							<div id='zkl' style="display:None">
							
							  <table border="0" width="100%" cellpadding="0" cellspacing="1">
							   <tr>
							    <td class='clefttitle' height='25' align='left' nowrap>
								<label><input type='checkbox' name='ePriceMarket' value='1'><strong>市 场 价:</strong></label>
								</td>
								<td>
							 <div>以<font color="#FF0000">“原始零售价”</font>为基准,将<font color="blue">“市场价”</font>按<input size="4" style="text-align:center" name='DiscountPriceMarket' type='text' value='9' onKeyPress= "return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))">折重置选中的所有商品</div>
							     <font color="#999999">Tips:市场价仅用于给客户参考的价格，不用于交易。</font>
							    </td>
							   </tr>
							   <tr><td colspan=2><hr color="green" size=1></td></tr>
							   <tr>
							    <td class='clefttitle' height='25' align='left' nowrap>
								<label><input type='checkbox' name='ePrice' value='1'><strong>非会员价:</strong></label>
								</td>
								<td>
							 <div>以<font color="#FF0000">“原始零售价”</font>为基准,将<font color="blue">“当前零售价”</font>按<input size="4" style="text-align:center" name='DiscountPrice' type='text' value='9' onKeyPress= "return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))">折重置选中的所有商品</div>
							     <font color="#999999">Tips:当前零售价指的是非注册会员看到的价格。</font>
							    </td>
							   </tr>
							   <tr><td colspan=2><hr color="green" size=1></td></tr>
							   <tr>
							    <td class='clefttitle' height='25' align='left'>
								 <label><input type='checkbox' name='ePriceMember' value='1'><strong>会 员 价:</strong></label>
								 </td>
								<td>
							 <div>以<font color="#FF0000">“原始零售价”</font>为基准
							 <br/>
							 1、如果商品按<font color="#FF0000">“会员统一指定价”</font>,则将<font color="blue">“会员价”</font>按<input size="4" style="text-align:center" name='DiscountPriceMember' type='text' onKeyPress= "return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))" value='9'>折重置选中的所有商品
							 <br/>
							 2、如果商品按<font color="#FF0000">“详细设置会员价”</font> ,则：
							 <br/>
							 <%
							  Response.Write "<table border='0' id='groupprice' width='80%'>"
							  Response.Write "<tr class='clefttitle'><td align='center'><b>会员组</b></td><td align='center'><b>价格</b></td></tr>"
							  Call KS.LoadUserGroup()
							  For Each Node In Application(KS.SiteSN&"_UserGroup").DocumentElement.SelectNodes("row")
							  Response.Write "<tr><td>" &Node.SelectSingleNode("@groupname").text & "</td><td>按<input onKeyPress= ""return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))"" onpaste=""return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))"" ondrop=""return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))"" class='textbox' type='text' style='text-align:center' size='4' name='price" & Node.SelectSingleNode("@id").text & "'  value='9' onKeyPress= ""return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))"" onpaste=""return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))"" ondrop=""return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))""> 折重置选中的所有商品</td></tr>"
							 Next
							  Response.Write "</table>"

							 %>
							 
							 </div>
							    </td>
							   </tr>
							   <tr><td colspan=2><hr color="green" size=1></td></tr>
							   <tr>
							    <td class='clefttitle' height='25' align='left' nowrap>
								<label><input type='checkbox' name='eScore' value='1'><strong>获赠积分:</strong></label>
								</td>
								<td>
							 <div>以<font color="#FF0000">“原始零售价”</font>为基准,将<font color="blue">“获赠积分”</font>按<input size="4" style="text-align:center" name='DiscountScore' type='text' value='9' onKeyPress= "return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))">折重置选中的所有商品</div>
							     <font color="#999999">Tips:获赠积分指客户成功购买商品可得到的赠送积分。</font>
							    </td>
							   </tr>
							   
							  </table>
							
							</div>
							
							
							</td>
						  </tr>
	                 </table>
					   </div>
					   
					   
					   <div class=tab-page id=kbxs-page1>
					<H2 class=tab>限时限量</H2>
					<SCRIPT type=text/javascript>
						 tabPane1.addTabPage( document.getElementById( "kbxs-page1" ) );
					</SCRIPT>
								
					 <table border="0" width="100%" cellpadding="0" cellspacing="1">
					     <tr class='tdbg'>
						  <td class='clefttitle' height='25' align='center'><input type='checkbox' name='eLimitBuy' value='1'></td>
						 <%
						 with response
							.Write "  <td class='clefttitle' align='right'><strong><font color=green>是否限时限量:</font></strong></td>"
							.Write "  <td style='padding:10px;margin-top:3px;border:1px solid #f9c943;background:#FFFFF6'>"
							.Write "<label onclick=""$('#LimitBuy').hide();""><input name='IsLimitbuy' type='radio'  value='0' checked> 正常销售</label> &nbsp;&nbsp;<label onclick=""$('#LimitBuy').show();$('#LimitBuyTaskID1').show();$('#LimitBuyTaskID2').hide();""><input name='IsLimitbuy' type='radio'  value='1'> 限时抢购</label>&nbsp;&nbsp;<label onclick=""$('#LimitBuy').show();$('#LimitBuyTaskID1').hide();$('#LimitBuyTaskID2').show();""><input name='IsLimitbuy' type='radio'  value='2'> 限量抢购</label>"
							.Write "<div id='LimitBuy' style='margin-tio:10px;padding:10px;display:none;border:0px solid #ff6600'>"

							
							.Write "抢购任务:"
							.Write "<select name='LimitBuyTaskID1' id='LimitBuyTaskID1' style='display:none'>"
							.Write "<option value=''>---请选择---</option>"
							
							 Dim RST:Set RST=Conn.Execute("Select ID,taskname from KS_ShopLimitBuy Where TaskType=1 and Status=1 Order by id desc")
							 Do While NOt RST.Eof
								.Write "<option value='" & RST(0) & "'>" & RST(1) & "</option>"
							 RST.MoveNext
							 Loop
							 RST.CLose 
							 .Write "</select>"
							 .Write "<select name='LimitBuyTaskID2' id='LimitBuyTaskID2' style='display:none'>"
							.Write "<option value=''>---请选择---</option>"
					
							 
							 Set RST=Conn.Execute("Select ID,taskname from KS_ShopLimitBuy Where TaskType=2 and Status=1 Order by id desc")
							 Do While NOt RST.Eof
								.Write "<option value='" & RST(0) & "'>" & RST(1) & "</option>"
							 RST.MoveNext
							 Loop
							  RST.Close: Set RST=Nothing
							  .Write "</select>"
							 
							.Write " <br/>"
							.Write "抢 购 价:<input type='text' style='text-align:center' name='LimitBuyPrice' value='100' size='6'  value='100' size='4' maxlength='4' class='textbox' onKeyPress= ""return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))"" onpaste=""return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))"" ondrop=""return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))"" class='textbox'>元<br/>"
							.Write "抢购数量:<input type='text' name='LimitBuyAmount' id='LimitBuyAmount' value='100' size='10'/>件   设置允许让抢购的商品数<br/>"
							.Write "</div>"
							.Write "</td>"
							.Write "</tr>"
		              End With
						 
						 %>
					 </table>
					</div>
					   
								
					<div class=tab-page id=att-page1>
					<H2 class=tab>属性设置</H2>
					<SCRIPT type=text/javascript>
						 tabPane1.addTabPage( document.getElementById( "att-page1" ) );
					</SCRIPT>
				<%end if%>

						
						<table border="0" width="100%" cellpadding="0" cellspacing="1">
						  <tr class='tdbg'> 
							<td class='clefttitle' height='25' align='center'><input type='checkbox' name='eTemplateID' value='1'></td>
							<td class='clefttitle' align='right'><strong>文档Web模板:</strong></td>
							<td><input type="text" size='40' name='TemplateID' id='TemplateID' class='textbox'>&nbsp;<%=KSCls.Get_KS_T_C("$('#TemplateID')[0]")%></td>
						  </tr>
						  <tr class='tdbg'> 
							<td class='clefttitle' height='25' align='center'><input type='checkbox' name='eWapTemplateID' value='1'></td>
							<td class='clefttitle' align='right'><strong>文档Wap模板:</strong></td>
							<td><input type="text" size='40' name='WapTemplateID' id='WapTemplateID' class='textbox'>&nbsp;<%=KSCls.Get_KS_T_C("$('#WapTemplateID')[0]")%></td>
						  </tr>
						  <tr class='tdbg'> 
							<td class='clefttitle'  height='25' align='center'><input type='checkbox' name='eKeyWords' value='1'></td>
							<td class='clefttitle' align='right'><strong>关 键 字:</strong></td>
							<td><input type="text" size='40' name='KeyWords' id='KeyWords' class='textbox'>&nbsp; <select name='SelKeyWords' style='width:100px' onChange='InsertKeyWords($("#KeyWords")[0],this.options[this.selectedIndex].value)'>
					<option value="" selected> </option><option value="Clean" style="color:red">清空</option>"
					<%=KSCls.Get_O_F_D("KS_KeyWords","KeyText","IsSearch=0 Order BY AddDate Desc")%>
					</select></td>
						  </tr>
						  <tr class='tdbg' id='showauthor'<%If ChannelID=5 or ChannelID=7 or ChannelID=8 Then KS.Echo " style='display:none'"%>> 
							<td class='clefttitle'  height='25' align='center'><input type='checkbox' name='eAuthor' value='1'></td>
							<td class='clefttitle' align='right'><strong>文档作者:</strong></td>
							<td> <input name='author' type='text' id='author' size=20 class='textbox'><<【<font color='blue'><font color='#993300' onclick='$("#author").val("未知")' style='cursor:pointer;'>未知</font></font>】【<font color='blue'><font color='#993300' onclick="$('#author').val('佚名')" style='cursor:pointer;'>佚名</font></font>】
							<select name='SelAuthor' style='width:100px' onChange="$('#author').val(this.options[this.selectedIndex].value)">")
						<option value="" selected> </option><option value="" style="color:red">清空</option>
						<%=KSCls.Get_O_F_D("KS_Origin","OriginName","ChannelID=1 and OriginType=1 Order BY AddDate Desc")%>
						 </select></td>
						  </tr>
						  <tr class='tdbg' id='showorigin'<%If ChannelID=5 or ChannelID=7 or ChannelID=8 Then KS.Echo " style='display:none'"%>>
							<td class='clefttitle'  height='25' align='center'><input type='checkbox' name='eOrigin' value='1'></td>
							<td class='clefttitle' align='right'><strong>文档来源:</strong></td>
							<td nowrap><input name='Origin' id='Origin' type='text' size=20 class='textbox'><<【<font color='blue'><font color='#993300' onclick="$('#Origin').val('不详')" style='cursor:pointer;'>不详</font></font>】【<font color='blue'><font color='#993300' onclick="$('#Origin').val('本站原创')" style='cursor:pointer;'>本站原创</font></font>】【<font color='blue'><font color='#993300' onclick="$('#Origin').val('互联网')" style='cursor:pointer;'>互联网</font></font>】
						<select name='selOrigin' style='width:100px' onChange="$('#Origin').val(this.options[this.selectedIndex].value)">
						<option value="" selected> </option><option value="" style="color:red">清空</option>
						<%=KSCls.Get_O_F_D("KS_Origin","OriginName","OriginType=0 Order BY AddDate Desc")%>
						</select></td>
						</tr>
						 <tr class='tdbg'> 
							<td class='clefttitle' height='25' align='center'><input type='checkbox' name='erank' value='1'></td>
							<td class='clefttitle' align='right'><strong>文档等级:</strong></td>
							<td><select name='rank'>
							 <option>★</option>
							 <option>★★</option>
							 <option selected>★★★</option>
							 <option>★★★★</option>
							 <option>★★★★★</option>
							</select>
						   </td>
						  </tr>		
						 <tr class='tdbg'> 
							<td class='clefttitle' height='25' align='center'><input type='checkbox' name='ehits' value='1'></td>
							<td class='clefttitle' align='right'><strong>点击数增加:</strong></td>
							<td><input type='text' value='0' name='hits' size='5'>次 <font color=#777777>说明在原点击数上累加</font></td>
						  </tr>		
						 <tr class='tdbg'> 
							<td class='clefttitle' height='25' align='center'><input type='checkbox' name='eAdddate' value='1'></td>
							<td class='clefttitle' align='right'><strong>添加时间:</strong></td>
							<td><input type='text' value='<%=now%>' name='AddDate' size='20'> <font color=#777777>格式:2008-12-1 10:10</font></td>
						  </tr>		
						  <tr class='tdbg'> 
							<td class='clefttitle' height='25' align='center'><input type='checkbox' name='eRecommend' value='1'></td>
							<td class='clefttitle' align='right'><strong>是否推荐:</strong></td>
							<td><input name='Recommend' type='radio' id='Recommend' value='1'> 是  <input name='Recommend' type='radio' id='Recommend' value='0' checked> 否</td>
						 </tr>
						 <tr class='tdbg'>
							<td class='clefttitle' height='25' align='center'><input type='checkbox' name='eIsTop' value='1'></td>
							<td class='clefttitle' align='right'><strong>是否固顶:</strong></td>
							<td><input name='IsTop' type='radio' value='1'> 是  <input name='IsTop' type='radio' value='0' checked> 否</td>
						</tr>
						<tr class='tdbg'>
							<td class='clefttitle' height='25' align='center'><input type='checkbox' name='eRolls' value='1'></td>
							<td class='clefttitle' align='right'><strong>是否滚动:</strong></td>
							<td><input name='Rolls' type='radio' value='1'> 是  <input name='Rolls' type='radio' value='0' checked> 否</td>
					   </tr>
					   <tr class='tdbg'>
						   <td class='clefttitle' height='25' align='center'><input type='checkbox' name='ePopular' value='1'></td>
							<td class='clefttitle' align='right'><strong>是否热门:</strong></td>
							<td><input name='Popular' type='radio' value='1'> 是  <input name='Popular' type='radio' value='0' checked> 否</td>
					  </tr>
					  <tr class='tdbg'>
						   <td class='clefttitle' height='25' align='center'><input type='checkbox' name='eStrip' value='1'></td>
							<td class='clefttitle' align='right'><strong>是否头条:</strong></td>
							<td><input name='Strip' type='radio' value='1'> 是  <input name='Strip' type='radio' value='0' checked> 否</td>
					 </tr>
					 <tr class='tdbg'>
							<td class='clefttitle'  height='25' align='center'><input type='checkbox' name='eCommentID' value='1'></td>
							<td class='clefttitle' align='right'><strong>允许评论:</strong></td>
							<td><input name='Comment' type='radio' value='1'> 是  <input name='Comment' type='radio' value='0' checked> 否</td>
					</tr>
					 <tr class='tdbg'>
						   <td class='clefttitle' height='25' align='center'><input type='checkbox' name='eSlide' value='1'></td>
							<td class='clefttitle' align='right'><strong>是否幻灯:</strong></td>
							<td><input name='Slide' type='radio' value='1'> 是  <input name='Slide' type='radio' value='0' checked> 否</td>
					 </tr>
					 <tr class='tdbg'>
						   <td class='clefttitle' height='25' align='center'><input type='checkbox' name='eVerific' value='1'></td>
							<td class='clefttitle' align='right'><strong>文档状态:</strong></td>
							<td><input name='verific' type='radio' value='1' checked> 已审  <input name='Verific' type='radio' value='0'> 未审</td>
					</tr>
			    </table>
				
			<%if ChannelID=5 then%>		
			  </div>
			 </div>	
			<%end if%>	
				
				
			</TD>
		 </tr>
		 <tr class='tdbg'>
		    <td colspan=3 height='30'><b>说明：</b>若要批量修改某个属性的值，请先选中其左侧的复选框，然后再设定属性值。<br><div align='center'> <input type="submit" class="button" name="button1" value="确定设置"> 
				  &nbsp; 
				  <%if ID<>"" then%>
				  <input type="reset" class="button" onClick="window.close()" name="button2" value=" 关闭取消 ">
				  <%else%>
				  <input type="reset" class="button" name="button2" value=" 重置 ">
				  <%end if%> </div></td>
		 </tr>

			  </form>
			</table>
			<br/>
			<br/>
			</div>
			</body>
			</html>
		 <%If Request.Form("Flag") = "Add" Then
		     If KS.G("choose")=0 Then
		      IDArr=Split(ID,",")
			 Else
			  IDArr=Split(Replace(KS.G("ClassID")," ",""),",")
			 End If
		      Set RS=Server.CreateObject("ADODB.RECORDSET")
			  For K=0 To Ubound(IDArr)
			  If KS.G("choose")=0 Then
			  RS.Open "Select * From " & KS.C_S(ChannelID,2) &" Where ID=" & IDArr(K), conn, 1, 3
			  Else
			  RS.Open "Select * From " & KS.C_S(ChannelID,2) &" Where Tid='" & IDArr(K) & "'", conn, 1, 3
			  End IF
			  If Not RS.EOF Then
			     Do While Not RS.Eof
				  If KS.ChkClng(KS.G("eTemplateID"))=1 And KS.G("TemplateID")<>"" Then RS("TemplateID") = KS.G("TemplateID")
				  If KS.ChkClng(KS.G("EWapTemplateID"))=1 And KS.G("WapTemplateID")<>"" Then RS("WapTemplateid")=KS.G("WapTemplateID")
				  If KS.ChkClng(KS.G("eKeyWords"))=1 Then
					 If InStr(" "&RS("KeyWords")&" "," "&KS.G("KeyWords")&" ") = 0 then
					  RS("KeyWords")  = RS("KeyWords")&" "&KS.G("KeyWords")
					 End If
				  End if
				  If KS.ChkClng(KS.G("eRank"))=1 And ChannelID<>8 Then       RS("Rank")      = KS.G("Rank")
				  If KS.ChkClng(KS.G("eAuthor"))=1 And ChannelID<>5 And ChannelID<>7 And ChannelID<>8 Then     RS("Author")    = KS.G("Author")
				  If KS.ChkClng(KS.G("eOrigin"))=1 And ChannelID<>5 And ChannelID<>7 And ChannelID<>8 Then     RS("Origin")    = KS.G("Origin")
				   Call SetAttributeField(RS)
				   
				   If ChannelID=5 Then '商城设置价格
				       If KS.ChkClng(KS.G("ProductType"))<>0 Then 
						   RS("ProductType")      = KS.ChkClng(KS.G("ProductType"))
						   If KS.ChkClng(KS.G("ProductType"))=3 Then RS("Discount")=Request("DiscountPrice")
						   If KS.ChkClng(KS.G("EPriceMarket"))=1 Then   RS("Price_Market") = (RS("Price_Original")*(Request("DiscountPriceMarket")/10)*100)/100
						   If KS.ChkClng(KS.G("EPrice"))=1 Then   RS("Price") = (RS("Price_Original")*(Request("DiscountPrice")/10)*100)/100
						   If KS.ChkClng(KS.G("eScore"))=1 Then RS("Point")=KS.ChkClng((RS("Price_Original")*(Request("DiscountScore")/10)*100)/100)
						   If KS.ChkClng(KS.G("EPriceMember"))=1 Then 
							  If RS("GroupPrice")=0 Then
							   RS("Price_Member") = (RS("Price_Original")*(Request("DiscountPriceMember")/10)*100)/100
							  Else
								Dim RSG:Set RSG=Server.CreateObject("ADODB.RECORDSET")
								RSG.Open "Select * From KS_ProPrice Where ProID=" & RS("ID"),Conn,1,3
								Do While Not RSG.Eof
								  RSG("Price")=(RS("Price_Original")*(Request("price"&RSG("GroupID"))/10)*100)/100
								  RSG.Update
								  RSG.MoveNext
								Loop
								RSG.Close
								Set RSG=Nothing
							  End If
						   End If
					  End If
					  
					  If KS.ChkClng(KS.G("eLimitBuy"))<>0 Then
					     If KS.ChkCLng(KS.S("LimitBuyTaskID" & KS.ChkClng(KS.G("IsLimitbuy"))))=0 Then
						   KS.AlertHintScript "请选择任务ID"
						   Response.End
						 End If
					     RS("IsLimitBuy")=KS.ChkClng(KS.G("IsLimitbuy"))
						 RS("LimitBuyPrice") = KS.S("LimitBuyPrice")
						 RS("LimitBuyAmount") = KS.ChkCLng(KS.S("LimitBuyAmount"))
						 RS("LimitBuyTaskID")=KS.ChkCLng(KS.S("LimitBuyTaskID" & KS.ChkClng(KS.G("IsLimitbuy"))))
						 
					   End If
					  
				   End If
				  
				   RS.Update
				 RS.MoveNext
				Loop
			 End If
			  RS.Close
			  
			  If KS.G("choose")=0 Then
			  RS.Open "Select * From [KS_ItemInfo] Where ChannelID=" & ChannelID &" And InfoID=" & IDArr(K), conn, 1, 3
			  Else
			  RS.Open "Select * From [KS_ItemInfo] Where Tid='" & IDArr(K) & "'", conn, 1, 3
			  End IF
			  If Not RS.EOF Then
			     Do While Not RS.Eof
				   Call SetAttributeField(RS)
				   RS.Update
				 RS.MoveNext
				Loop
			 End If
			  RS.Close
			  
			  
			 Next 
			 
			
			 
		   Set RS = Nothing
		   conn.Close:Set conn = Nothing
		   if ID<>"" then
		   KS.Echo "<script>alert('恭喜，成功设置了选中文档的属性!');window.close();</script>"
		   else
		   KS.Echo "<script>alert('恭喜，批量设置成功!');</script>"
		   end if
		End If
		End Sub
		
		Sub SetAttributeField(RS)
				  If KS.ChkClng(KS.G("eHits"))=1 Then       RS("Hits")      =RS("Hits")+KS.ChkCLng(KS.G("Hits"))
				  If KS.ChkClng(KS.G("eRecommend"))=1 Then RS("Recommend") = KS.ChkCLng(KS.G("Recommend"))
				  If KS.ChkClng(KS.G("eRolls"))=1 Then     RS("Rolls")     = KS.ChkClng(KS.G("Rolls"))
				  If KS.ChkClng(KS.G("eStrip"))=1 Then     RS("Strip")     = KS.ChkClng(KS.G("Strip"))
				  If KS.ChkClng(KS.G("ePopular"))=1 Then   RS("Popular")   = KS.ChkClng(KS.G("Popular"))
				  If KS.ChkClng(KS.G("eCommentID"))=1 Then   RS("Comment")   = KS.ChkClng(KS.G("Comment"))
				  If KS.ChkClng(KS.G("eIsTop"))=1 Then     RS("IsTop")     = KS.ChkClng(KS.G("IsTop"))
				  If KS.ChkClng(KS.G("eSlide"))=1 Then     RS("Slide")     = KS.ChkCLng(KS.G("Slide"))
				  If KS.ChkClng(KS.G("eVerific"))=1 Then   RS("Verific")   = KS.ChkCLng(KS.G("Verific"))
				  If KS.ChkClng(KS.G("eAdddate"))=1 And IsDate(KS.G("AddDate")) Then  RS("AddDate")=KS.G("AddDate")
		End Sub

	
End Class
%> 
