<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit
%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.FileIcon.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.EscapeCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="gb2312"

Dim KSCls
Set KSCls = New Template
KSCls.Kesion()
Set KSCls = Nothing

Class Template
        Private KS
		'===========================================================================
		Private I, totalPut, TemplateSql, KS_T_RS
		Private TemplateType, ChannelID,DomainStr,MaxPerPage,CurrentPage,TotalPages
		Private FileItem, CurrPath, ParentPath,InstallDir,Path
		'=============================================================================
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  DomainStr=KS.GetDomain
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
			If Not KS.ReturnPowerResult(0, "KMTL10006") Then                'ģ������Ȩ�޼��
			  Call KS.ReturnErr(1, "")
			  Exit Sub
			End If
			
			
			Select Case KS.G("Action")
			 Case "getcontent"
			   Call getcontent()
			   response.end
			 Case "Del"
			   Call TemplateDel()
			 Case "NewPage","Modify"
			   Call AddTemplate()
			 Case "TemplateSave"
			   Call TemplateSave()
			 Case Else
			   Call TemplateList()
			End Select
		End Sub
		
		Sub getcontent()
		 response.write Escape(KS.ReadFromFile(Replace(Replace(UnEscape(KS.G("TemplateFileName")),"{@TemplateDir}",KS.Setting(3) & KS.Setting(90)),"//","/")))
		End Sub
		
		Sub TemplateList()
		With Response
		InstallDir=KS.Setting(3)
        If CurrPath = "" Then
			ParentPath = ""
			CurrPath= InstallDir & KS.Setting(90)
		Else
			ParentPath = Mid(CurrPath, 1, InStrRev(CurrPath, "/") - 1)
			If ParentPath = "" Then
				ParentPath = Left(InstallDir, Len(InstallDir) - 1)
			End If
		End If
		If Right(CurrPath,1)="/" Then CurrPath=Left(CurrPath,Len(CurrPath)-1)		
		
		
		.Write "<html>"
		.Write "<head>"
		.Write "<title>ģ�����</title>"
		.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
		.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
		.Write "<script language=""JavaScript"">"
		.Write "var ParentPath='" & ParentPath & "';" & vbCrLf
		.Write "var ChannelID='" & ChannelID & "';" & vbCrLf
		.Write "var TemplateType='" & TemplateType & " ';" & vbCrLf
		.Write "</script>"
		.Write "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>"
		.Write "<script language=""JavaScript"" src=""../KS_Inc/jquery.js""></script>"
		.Write "<script language=""JavaScript"" src=""Include/ContextMenu1.js""></script>"
		.Write "<script language=""JavaScript"" src=""Include/SelectElement.js""></script>"
		.Write "<script language=""JavaScript"" src=""../KS_Inc/kesion.box.js""></script>"
		%>
		<script language="javascript">
		function CreateHtml()
		{   var ids=get_Ids(document.myform);
			if (ids!='')
				PopupCenterIframe('����ѡ�е��Զ���ҳ��','Include/RefreshCommonPageSave.asp?RefreshFlag=Folder&PageID='+ids,530,50,'no')
			else 
				alert('��ѡ��Ҫ�������Զ���ҳ��!');
        }		
		var DocElementArrInitialFlag=false;
		var DocElementArr = new Array();
		var DocMenuArr=new Array();
		var SelectedFile='',SelectedFolder='';
		function document.onreadystatechange()
		{   
			if (DocElementArrInitialFlag) return;
			InitialDocElementArr('FolderID','SelectObjID');
			InitialContextMenu();
			DocElementArrInitialFlag=true;
		}
		
		function InitialContextMenu()
		{   DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.AddDIYPage('');",'�½�ҳ��(N)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.TextEdit('text');",'�ı��༭(W)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.TextEdit('');",'���ӱ༭(E)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.TemplateControl(2);",'ɾ ��(D)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.location.reload();",'ˢ ��(Z)','disabled');
		}
		function DocDisabledContextMenu()
		{
			DisabledContextMenu('FolderID','SelectObjID','���ӱ༭(E),�ı��༭(W),ɾ ��(D)','�ı��༭(W),���ӱ༭(E)','','','','')
		}
		function GoBack()
		{
		 location.href='?';
		}
		function AddDIYPage()
		{
		$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape("�Զ��嵥ҳ�ļ� >> �����ҳ��")+'&ButtonSymbol=Go';
		location.href='?Action=NewPage&flag=text';
		}		
		function EditTemplate(id)
		{
		window.parent.parent.frames['MainFrame'].location.href='KS.DIYPage.asp?Action=Modify&TemplateID='+id;
		$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape("ģ��������� >> �༭ҳ��")+'&ButtonSymbol=TemplateAdd';
		}
		function TextEdit(Flag)
		{
			GetSelectStatus('FolderID','SelectObjID');
		 if (SelectedFile!='')
			if (SelectedFile.indexOf(',')==-1) 
			{
			 location.href='KS.DIYPage.asp?Action=Modify&Flag='+Flag+'&id='+SelectedFile;
			 $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape("ģ��������� >> �༭ҳ��")+'&ButtonSymbol=Gosave';
			}
			else alert('һ��ֻ�ܱ༭һ��ģ���ļ�!')	 
	     else
		 alert('��ѡ��Ҫһ��ģ��!');
		}
		function DelTemplate(id)
		{
		if (confirm('ɾ���󽫵����Ѱ󶨵���Ϣ�Ҳ���ģ�壬ȷ�ϲ�����?'))
		 location="KS.DIYPage.asp?Action=Del&id="+id;
		}
		function TemplateControl(op)
		{
			var alertmsg='';
			GetSelectStatus('FolderID','SelectObjID');
			if (SelectedFile!='')
			 {  switch (op)
				{           
				 case 2 :   
				   DelTemplate(SelectedFile); 
				   break;
				}   
			 }   
			else 
			 {
			   switch (op)
				{case 1 :
				  alertmsg="�༭";
				   break;
				 case 2:
				  alertmsg="ɾ��"; 
				  break;
				  case 3:
				   alertmsg="����Ĭ��"; 
				  break;
				 default:
				  alertmsg="����" 
				  break;
				 } 
			 alert('��ѡ��Ҫ'+alertmsg+'��ģ��');
			  }
		}
		function GetKeyDown()
		{ event.returnValue=false;
		if (event.ctrlKey)
		  switch  (event.keyCode)
		  {  case 90 : location.reload(); break;
			 case 65 : AddTemplate('');break;
			 case 77 : AddDIYPage();break;
			 case 69 : TextEdit('');break;
			 case 87 : TextEdit('text');break;
			 case 68 : TemplateControl(2);break;
			 case 83 : TemplateControl(3);break;
		   }	
		else	
		 if (event.keyCode==46)TemplateControl(2);
		}
		</script>
		<%
		.Write "</head>"
		.Write "<body scroll=no topmargin=""0"" leftmargin=""0"" onclick=""SelectElement();"" onkeydown=""GetKeyDown();"" onselectstart=""return false;"">"

        .Write "<ul id='menu_top'>"
		.Write "<li class='parent' onclick=""AddDIYPage();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>�½�ҳ��</span></li>"
		.Write "<li class='parent' onclick=""TextEdit('text');""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>�༭ҳ��</span></li>"
		.Write "<li class='parent' onclick=""location.href='include/refreshcommonpage.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/move.gif' border='0' align='absmiddle'>�����Զ���ҳ��</span></li>"
		.Write "</ul>"	
		
		.Write ("<div style=""height:94%; overflow: auto; width:100%"" align=""center"">")	
		.Write "<table width=""100%"" height=""25"" border=""0"" cellpadding=""0"" cellspacing=""1"">"
		.Write "  <tr align=""center"" class=""sort"">"
		.Write "    <td align=""center"" width=""40"">ѡ��</td>"
		.Write "    <td height=""25"" class=""sort""> <div align=""center""><font color=""#990000"">ҳ������</font></div></td>"
		.Write "    <td align=""center"">ģ��·��</td>"
		.Write "    <td width=""143"" align=""center"">�޸�ʱ��</td>"
		.Write "    <td width=""267"" class=""sort"">��������</td>"
		.Write "  </tr>"
		.Write "  <form name='myform' id='myform' action='KS.DiyPage.asp' method='get'>"
		.Write "  <input type='hidden' name='action' value='Del'>"
		
		call ShowContent
		  
		.Write "</table>"
		
		%>
		<div style="margin:3px;text-align:right">
		<b>ѡ��</b><a href='javascript:Select(0)'><font color=#999999>ȫѡ</font></a> - <a href='javascript:Select(1)'><font color=#999999>��ѡ</font></a> - <a href='javascript:Select(2)'><font color=#999999>��ѡ</font></a> <input type='button' onclick='CreateHtml()' class='button' value='����ѡ�е��Զ���ҳ��'> <input type='submit' onclick="return(confirm('�˲���������,ȷ��ɾ����?'))" class='button' value='ɾ��ѡ�е��Զ���ҳ��'>
		</div>
		</form>
		<%
		 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
		
		.Write "</div>"
		.Write "</body>"
		.Write "</html>"
		End With
		End Sub
		Sub showContent()
		  CurrentPage=KS.ChkClng(Request("page"))
		  if CurrentPage=0 Then CurrentPage=1
		  MaxPerPage=10   'ÿҳ��ʾ����
		With Response
           Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		   RS.Open "Select * From KS_Template",conn,1,1
		   If RS.Eof And RS.Bof Then
		   Else
		            TotalPut=rs.recordcount
					if (TotalPut mod MaxPerPage)=0 then
						TotalPages = TotalPut \ MaxPerPage
					else
						TotalPages = TotalPut \ MaxPerPage + 1
					end if
					if CurrentPage > TotalPages then CurrentPage=TotalPages
					if CurrentPage < 1 then CurrentPage=1
					rs.move (CurrentPage-1)*MaxPerPage
		            dim i:i=0
			   Do While Not rs.eof
			   
			  .Write "<tr id='u" & rs("templateid") & "' onclick=""chk_iddiv('" & rs("templateid") & "')"" class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
			  .Write "  <td class='splittd' align='center'><input type='checkbox' name='id' id='c" & rs("templateid") & "' value='" & rs("templateid") &"'></td>"
			  .Write "  <td class='splittd'><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
			  .Write "      <tr>"
			  .Write "        <td height=""20"">"
			  .Write "         <span SelectObjID=" &rs("templateid") & " onDblClick=""TextEdit('');"">"
			  .Write "         <img src=""Images/Folder/TheSmallWordNews1.gif"" align=""absmiddle"">"
			  .Write "          <span style=""cursor:default"">" & rs("TemplateName") & "</span></span></td>"
			  .Write "      </tr>"
			  .Write "    </table></td>"
			  .Write "    <td class='splittd'>" & rs("templatefilename") & "</td>"
	
			
			  .Write ("<td align='center' class='splittd'>" & rs("AddDate") & " </td>")
			  .Write ("<td align='center' class='splittd'><a href=""" & ks.setting(3) & ks.setting(94) & rs("fsofilename") & """ target=""_blank"">Ԥ��ҳ��</a> | <a href='KS.DIYPage.asp?Action=Modify&ID=" & rs("templateid") &"&Flag=text' onclick=""$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=ģ��������� >> �༭ҳ��&ButtonSymbol=GoSave'"">�ı��༭</a> | <a href='KS.DIYPage.asp?Action=Modify&id=" & rs("templateid") &"'onclick=""$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=ģ��������� >> �༭ҳ��&ButtonSymbol=GoSave'"">���ӻ��༭</a> | <a href='KS.DIYPage.asp?Action=Del&id=" & rs("templateid")&"' onclick=""return(confirm('�˲��������棬ȷ��ɾ����'))"">ɾ��</a></td>")
			  .Write "</tr>"
			  RS.MoveNext
			   i=I+1
			   if i>=MaxPerPage Then Exit Do
			 Loop
		  End If	 
			 
		  End With 
			 	   
	     End Sub
			 
			 'ɾ��ģ��
		Sub TemplateDel
			Dim IDArr:IDArr=Split(Replace(KS.G("ID")," ",""),",")
			Dim I
			For i=0 to Ubound(IDarr)
			Dim FileName,CurrPath
			Call KS.DeleteFile(KS.Setting(3) & KS.Setting(94) & conn.execute("select fsofilename from ks_template where templateid=" & IDarr(i))(0))
			conn.execute("delete from ks_template where templateid=" & IDarr(i))
			Next
			'Call KS.DeleteFolder(CurrPath & "/" & FileName)
		    Response.Write "<script>window.location.href='KS.DIYPage.asp'</script>"
       End Sub
	   
	   
	   '����ģ��
	  Sub AddTemplate()
		Dim Action, TemplateID, ChannelID, TemplateType, TemplateName, FsoFileName, FnameType, TemplateContent, TemplateFileName, TemplateFromFileContent,Action1,FileName
		Dim  InstallDir, TemplateDIr,PageName
		InstallDir  = KS.Setting(3)

		If KS.G("Action")="NewPage" Then
		PageName=""
		Else
		 Dim RSt:Set RSt=Server.CreateoBject("adodb.recordset")
		 rst.open "select * from KS_Template Where TemplateID=" & KS.ChkClng(KS.G("id")),Conn,1,1
		 If RSt.Eof Then
		  Call KS.Alert("�������ݳ���!","")
		  exit sub
		 end if
		 PageName=rst("TemplateName")
		 FileName=rst("FsoFileName")
		 TemplateFileName=rst("TemplateFileName")
         TemplateFromFileContent=KS.ReadFromFile(Replace(Replace(TemplateFileName,"{@TemplateDir}",KS.Setting(3) & KS.Setting(90)),"//","/"))
		End If
		
		Response.Write "<html><head><title>ģ�����-���ģ��</title>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		Response.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
		Response.Write "<script type=""text/JavaScript"" src=""../KS_Inc/common.js""></script>"
		Response.Write "<script src=""../ks_inc/jquery.js""></script>"
		%>
                    <script language = 'JavaScript'>
					function LoadTemplateIn()
					{ 
					    var url='KS.DIYPage.asp';
						$.get(url,{action:"getcontent",TemplateFileName:escape($("#TemplateFileName").val())},function(d){
						  $('#Content').val(unescape(d));
						})
					}	
									
		            function show_ln(txt_ln,txt_main){
			            var txt_ln  = document.getElementById(txt_ln);
			            var txt_main  = document.getElementById(txt_main);
			            txt_ln.scrollTop = txt_main.scrollTop;
			            while(txt_ln.scrollTop != txt_main.scrollTop)
			            {
				            txt_ln.value += (i++) + '\n';
				            txt_ln.scrollTop = txt_main.scrollTop;
			            }
			            return;
		            }
		            function editTab(){
			            var code, sel, tmp, r
			            var tabs=''
			            event.returnValue = false
			            sel =event.srcElement.document.selection.createRange()
			            r = event.srcElement.createTextRange()
			            switch (event.keyCode){
				            case (8) :
				            if (!(sel.getClientRects().length > 1)){
					            event.returnValue = true
					            return
				            }
				            code = sel.text
				            tmp = sel.duplicate()
				            tmp.moveToPoint(r.getBoundingClientRect().left, sel.getClientRects()[0].top)
				            sel.setEndPoint('startToStart', tmp)
				            sel.text = sel.text.replace(/\t/gm, '')
				            code = code.replace(/\t/gm, '').replace(/\r\n/g, '\r')
				            r.findText(code)
				            r.select()
				            break
			            case (9) :
				            if (sel.getClientRects().length > 1){
					            code = sel.text
					            tmp = sel.duplicate()
					            tmp.moveToPoint(r.getBoundingClientRect().left, sel.getClientRects()[0].top)
					            sel.setEndPoint('startToStart', tmp)
					            sel.text = '\t'+sel.text.replace(/\r\n/g, '\r\t')
					            code = code.replace(/\r\n/g, '\r\t')
					            r.findText(code)
					            r.select()
				            }else{
					            sel.text = '\t'
					            sel.select()
				            }
				            break
			            case (13) :
				            tmp = sel.duplicate()
				            for (var i=0; tmp.text.match(/[\t]+/g) && i<tmp.text.match(/[\t]+/g)[0].length; i++) tabs += '\t'
				            sel.text = '\r\n'+tabs
				            sel.select()
				            break
			            default  :
				            event.returnValue = true
				            break
				            }
			            }
		            //-->
		            </script>
		<%
		Response.Write "</head>"
		Response.Write "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
		Response.Write " <table width='100%' height='25' border='0' cellpadding='0' cellspacing='1' bgcolor='#efefef'>"
        Response.Write " <tr>"
        Response.Write " <td class='sort'><div align='center'><font color='#990000'>�޸�ģ��</font></div></td>"
        Response.Write "  </tr>"
        Response.Write "</table>"
		 Response.Write "<table width='100%' height=""350"" style=""background-color:#ffffff;padding-right: 2px;padding-left: 2px;padding-bottom: 0px;"" border='0' align='center' cellpadding='0' cellspacing='1' class='ctable'>"
		 Response.Write " <form name=""TemplateForm"" method=""post"" action=""KS.DIYPage.asp?Action=TemplateSave&id=" & ks.g("id") & """ onSubmit=""return(CheckForm())"">"	
		 		
		 Response.Write "   <tr class=""clefttitle"">"
		 Response.Write "     <td height=""30""><b>��ҳ���ƣ�</b><input name=""PageName"" type=""text"" id=""PageName"" size=""30"" Value=""" & PageName & """> <font color=red>*</font>�磬�������ģ���վ���ܵ�</td></tr>"
		 
		 Response.Write "   <tr class=""clefttitle"">"
		 Response.Write "     <td height=""30""><b>��ҳģ���ַ��</b><input name=""TemplateFileName"" type=""text"" id=""TemplateFileName"" size=""30"" Value=""" & TemplateFileName & """>&nbsp;"
		 	    Dim CurrPath:CurrPath=KS.Setting(3)&KS.Setting(90)
		If Right(CurrPath,1)="/" Then CurrPath=Left(CurrPath,Len(CurrPath)-1)
        response.write  "<input type='button' name=""Submit"" class=""button"" value=""ѡ��ģ��..."" onClick=""OpenThenSetValue('KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle='+escape('����ģ��')+'&CurrPath=" & Server.URLEncode(CurrPath) & "',450,350,window,TemplateFileName);LoadTemplateIn();"">"	 

		 if KS.G("Flag")<>"text" then
		 Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input class=""button"" type=""button"" onclick=""settemplatearea(0)"" value=""����ģʽ""> <input type=""button"" class=""button"" onclick=""settemplatearea(1)"" value=""���ӻ�ģʽ""> <input type=""button"" value=""���ƴ���""  class=""button"" onclick=""copy()"">&nbsp;"
		 end if
		 Response.Write "  </td></tr>"
		 
		 Response.Write "   <tr class=""clefttitle"">"
		 Response.Write "     <td height=""30""><b>�����ļ����ƣ�</b>" & KS.Setting(3) & KS.Setting(94) &"<input name=""FileName"" type=""text"" id=""FileName"" size=""24"" Value=""" & FileName & """> <font color=red>*</font> �ɴ�·������ ""html/help.html"",""Common/about/help.htm""��</td></tr>"

		 Response.Write "   <tr id=""toplabelarea"" class=""clefttitle"""
		 if KS.G("Flag")<>"text" then Response.Write " style='display:none'"
		 Response.Write ">"
		 Response.Write "	<td valign=""top""><strong>�����ǩ��</strong>"
		 Response.Write "<select name=""mylabel"" style=""width:160px"">"
		 Response.Write " <option value="""">==ѡ��ϵͳ������ǩ==</option>"
		   Dim RS:Set RS=Server.Createobject("adodb.recordset")
		   rs.open "select top 200 LabelName from KS_Label Where LabelType<>5 order by adddate desc",conn,1,1
		   If not Rs.eof then
		    Do While Not Rs.Eof
			 Response.Write "<option value=""" & RS(0) & """>" & RS(0) & "</option>"
			 RS.MoveNext
			Loop 
		   End If
		  Response.Write "</select>&nbsp;<input class='button' type='button' onclick='LabelInsertCode(document.all.mylabel.value);' value='�����ǩ'>"
		  RS.Close:Set RS=Nothing
		 Response.Write "&nbsp;<input type=""button"" class='button' onclick=""javascript:LabelInsertCode();"" value=""ѡ������ǩ"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		 Response.Write " </td>"
		 Response.Write "</tr>"
		 
		 
		 Response.Write "   <tr id=""codearea"""
		 if KS.G("Flag")<>"text" then Response.Write " style='display:none'"
		 Response.Write "   >"
		 Response.Write "   <td>"
		 Response.Write "     <table border='0' cellspacing='0' cellspadding='0'>"
		 Response.Write "	  <tr>"
		 Response.Write "       <td valign=""top"" width='20'>"
		 %>
		  <textarea name="txt_ln" id="txt_ln" cols="6" style="overflow:hidden;height:423;background-color:highlight;border-right:0px;text-align:right;font-family: tahoma;font-size:12px;font-weight:bold;color:highlighttext;cursor:default;" readonly="">
<% Dim N
For N=1 To 3000
 Response.Write N & vbcrlf
Next
%>
</textarea>
		             </td>
		             <td valign="top"><textarea name="Content" rows="2" cols="30" id="Content" onscroll="show_ln('txt_ln','Content')" onKeyDown="editTab()" onChange="TemplateContent.SetContentIni();" style="height:422px;width:750;"><%=Server.HTMLEncode(TemplateFromFileContent)%></textarea>
</td>
		             </tr>
					 </table>
					 <%

		 
		 Response.Write "   <tr id=""editorarea"""
		 if KS.G("Flag")="text" then Response.Write " style=""display:none"""
		 Response.Write ">"
		 Response.Write "    <td colspan=""2"" width=""100%"" height=""510"">"
		 Response.Write "       <iframe id='TemplateContent' name='TemplateContent' src='KS.Editor.asp?ID=Content&style=3&sChannelID=" & ChannelID &"&TemplateType=" & TemplateType &"' frameborder='0' scrolling='no' width='100%' height='100%'></iframe>"
		 Response.Write "     </td>"
		 Response.Write "   </tr>"
		 Response.Write " </form>"
		 Response.Write "</table>"
		 Response.Write "</body>"
		 Response.Write "</html>"
			 Conn.Close:Set Conn = Nothing
		Response.Write "<script language=""JavaScript"">" & vbCrLf
		Response.Write "    function copy()" & vbcrlf
		Response.Write "{" & vbcrlf
		Response.Write "document.TemplateForm.Content.value=document.TemplateForm.Content.value;" & vbcrlf
        Response.Write "    document.TemplateForm.Content.select();" & vbcrlf
        Response.Write "    textRange = document.TemplateForm.Content.createTextRange();" & vbcrlf
        Response.Write "    textRange.execCommand(""Copy"");" & vbcrlf
		Response.Write "    alert('��ϲ����ǰ�����Ѹ��Ƶ�������!');" & vbcrlf
        Response.Write "}" & vbcrlf
		Response.Write "function LabelInsertCode(Val)" & vbcrlf
		Response.Write "{"
		Response.Write " if (Val==null)" & vbcrlf
		Response.Write "  Val=OpenWindow('Include/LabelFrame.asp?sChannelID=" & ChannelID &"&TemplateType=" & TemplateType &"&url=InsertLabel.asp&pagetitle='+escape('�����ǩ'),260,350,window);"&vbcrlf
		Response.Write "if (Val!='')"
		Response.Write "{ document.TemplateForm.Content.focus();" & vbcrlf
		Response.Write "  var str = document.selection.createRange();" & vbcrlf
		Response.Write "  str.text = Val;"
		Response.Write " }" & vbcrlf
		Response.Write "}" & vbcrlf
		Response.Write "function InsertFunctionLabel(Url,Width,Height)" & vbcrlf
        Response.Write "{" & vbcrlf
        Response.Write "var Val = OpenWindow(Url,Width,Height,window);"
		Response.Write "if (Val!=''&&Val!=null)"
		Response.Write "{ document.TemplateForm.Content.focus();" & vbcrlf
		Response.Write "  var str = document.selection.createRange();" & vbcrlf
		Response.Write "  str.text = Val;"
		Response.Write " }" & vbcrlf
        Response.Write "}" & vbcrlf
		Response.Write "function CheckForm()" & vbCrLf
		Response.Write "{  if ($('#PageName').val()=="""")"
		Response.Write "     {"
		Response.Write "     alert(""������ҳ������!"");"
		Response.Write "     $('#PageName').focus();"
		Response.Write "     return false;"
		Response.Write "     }" & vbCrLf
		Response.Write " if ($('#TemplateFileName').val()=="""")"
		Response.Write "     {"
		Response.Write "     alert(""�뵼�뵥ҳģ��!"");"
		Response.Write "     return false;"
		Response.Write "     }" & vbCrLf
		Response.Write "  if ($('#FileName').val()=="""")"
		Response.Write "     {"
		Response.Write "     alert(""���������ɵ��ļ���!"");"
		Response.Write "     $('#FileName').focus();"
		Response.Write "     return false;"
		Response.Write "     }" & vbCrLf

		Response.Write "   if (frames[""TemplateContent""].CurrMode!='EDIT') {alert('��ʾ��Ϣ:\n\nҪ����ģ�壬���л������ģʽ');return false;}"
		if KS.G("Flag")<>"text" then
		Response.Write "    document.TemplateForm.Content.value=""<html>\n""+frames[""TemplateContent""].ReplaceUrl(frames[""TemplateContent""].ReplaceImgToScript(frames[""TemplateContent""].Resumeblank(frames[""TemplateContent""].KS_EditArea.document.documentElement.innerHTML)))+""\n</html>"";"
		end if
		Response.Write "    document.TemplateForm.submit();"
		Response.Write "    return true;"
		Response.Write "}" & vbCrLf
		Response.Write "</script>" & vbCrLf
	  End Sub
	  
	  Sub TemplateSave()
	  	Dim Action, ChannelID, TemplateType, TemplateName, TemplatConTent, TemplateFileName, TemplateID, FsoFileName, TemplateContent,FileName
		Dim ObjRS, SQLStr, IsDefault, TemplateFilePath, OpStr
		 TemplateName = Trim(Request("PageName"))
		 TemplateContent = Trim(Request("Content"))
		 TemplateFileName = Request("TemplateFileName")   
		 FileName=Request("FileName")
		 If Instr(FileName,".")=0 Then
			Call KS.AlertHistory("������ļ���ʽ����ȷ!", -1)
			Set KS = Nothing:Response.End
		 Else
		   Dim FileExt:FileExt=lcase(Split(FileName,".")(1))
		   If FileExt<>"html" and FileExt<>"htm" and FileExt<>"shtml" and FileExt<>"shtm" Then
			Call KS.AlertHistory("������ļ���ʽ����ȷ,ֻ����html,htm,shtml��shtmΪ��չ��!", -1)
			Set KS = Nothing:Response.End
		   End If
		 End If
		 
		 If InStr(lcase(TemplateFileName),".asp")>0 or InStr(lcase(TemplateFileName),".asa")>0 or InStr(lcase(TemplateFileName),".php")>0 or InStr(lcase(TemplateFileName),".cer")>0 Then
			Call KS.AlertHistory("ģ���ļ�����ʽ����ȷ!", -1)
			Set KS = Nothing:Response.End
		 End If

				'���������ȷ��
				If TemplateFileName = "" Then
				  Call KS.AlertHistory("����û�е���ģ��!", -1)
				  Set KS = Nothing:Response.End
				End If
				
			 TemplateContent = ReplaceBadStr(Replace(Replace(Replace(TemplateContent, "contentEditable=true", ""), KS.GetDomain, "/"), KS.Setting(2), ""))
			If (Instr(TemplateContent,"<%")<>0) or (instr(TemplateContent,"<?")<>0 and instr(TemplateContent,"?>")<>0)  or instr(lcase(TemplateContent),"createobject(""adodb.stream"")")>0 Then
				  Call KS.AlertHistory("����ģ���ʽ����ȷ,�벻Ҫ������ִ�д���!", -1)
				  Set KS = Nothing
				  Response.End
			 End If

			 
			  If KS.WriteTOFile(Replace(Replace(TemplateFileName,"{@TemplateDir}",KS.Setting(3) & KS.Setting(90)),"//","/"), TemplateContent) = True Then
			   dim rs:set rs=server.createobject("adodb.recordset")
			   rs.open "select * from ks_template where templateid=" & ks.chkclng(ks.g("id")),conn,1,3
			   if rs.eof then
			    rs.addnew
			   end if
			    rs("TemplateName")=TemplateName
				rs("TemplateFileName")=TemplateFileName
				rs("fsofilename")=FileName
				rs("adddate")=now
				rs.update
				rs.close
				set rs=nothing
				'���ɲ���
				Dim KSRCls:Set KSRCls=New Refresh
				Call KSRCls.RefreshCommonPage(TemplateFileName,FileName)
				Set KSRCls=Nothing
			  Response.Write "<script src='../ks_inc/jquery.js'></script>"
			  Response.Write ("<script> alert('�ɹ���ʾ:ģ�屣��ɹ�!');window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=ģ�����&ButtonSymbol=Disabled'; location.href='KS.DIYPage.asp';</script>")
			  Else
				Call KS.AlertHistory("������ʾ,����ԭ��:1.����ʧ�ܣ�ģ���ļ�������;\n2.Ŀ¼û��д��Ȩ��", -1)
				Set KS = Nothing
			  End If
		End Sub
		Function ReplaceBadStr(Content)
			Dim regEx, Matches, Match
			Set regEx = New RegExp
			regEx.Pattern = "/" & KS.Setting(89) & "([A-Z]|[a-z]|\.|\?|\=|&|;|[0-9])*#"
			regEx.IgnoreCase = True
			regEx.Global = True
			Set Matches = regEx.Execute(Content)
			ReplaceBadStr = Content
			For Each Match In Matches
				On Error Resume Next
				ReplaceBadStr = Replace(ReplaceBadStr, Match.Value, "#")
			Next
		End Function


 End Class
%>