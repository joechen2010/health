<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New JS_Main
KSCls.Kesion()
Set KSCls = Nothing

Class JS_Main
        Private KS
		'========================================================================
		Private JSSql, JSRS, FolderID, JSID, ChannelID, Channel, Action
		Private i, totalPut, CurrentPage, JSType
		Private KeyWord, SearchType, StartDate, EndDate
		'������������
		Private SearchParam
		Private MaxPerPage
		Private Row 
		'========================================================================
		Private Sub Class_Initialize()
		  MaxPerPage = 96
		  Row = 8
		  Set KS=New PublicCls
		   Call KS.DelCahe(KS.SiteSn & "_labellist")
		   Call KS.DelCahe(KS.SiteSn & "_ReplaceFreeLabel")
		   Call KS.DelCahe(KS.SiteSn & "_jslist")
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		'�ɼ�������Ϣ
		KeyWord = KS.G("KeyWord")
		SearchType = KS.G("SearchType")
		StartDate = KS.G("StartDate")
		EndDate = KS.G("EndDate")
		SearchParam = "KeyWord=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate
		JSType = KS.G("JSType"):If JSType = "" Then JSType = 0
		
		Select Case KS.G("JsAction")
		 Case "JSDel"
		   Call JSDel()
		 Case "JSFolderDel"
		   Call JSFolderDel()
		 Case "JSView"
		   Call JSView()
		 Case Else
		   Call JSMainList()
		End Select
		End Sub
		
		Sub JSMainList()
		   With Response
			If JSType = 0 Then
				If Not KS.ReturnPowerResult(0, "KMTL10004") Then                'ϵͳJS�����Ȩ�޼��
				  Call KS.ReturnErr(1, "")
				  .End
				End If
			ElseIf JSType = 1 Then
				If Not KS.ReturnPowerResult(0, "KMTL10005") Then                '����JS�����Ȩ�޼��
				  Call KS.ReturnErr(1, "")
				  .End
				End If
			End If
			
			If Not IsEmpty(KS.G("page")) And KS.G("page") <> "" Then
				  CurrentPage = CInt(KS.G("page"))
			Else
				  CurrentPage = 1
			End If
			Action = KS.G("Action")
			FolderID = Trim(KS.G("FolderID"))
			If FolderID = "" Then FolderID = "0"
			Dim UPFolderRS, ParentID
			Set UPFolderRS = Conn.Execute("select * from [KS_LabelFolder] where  ID ='" & FolderID & "'")
			If Not UPFolderRS.EOF Then
			 ParentID = UPFolderRS("ParentID")
			End If
			UPFolderRS.Close:Set UPFolderRS = Nothing
			.Write "<html>"
			.Write "<head>"
			.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			.Write "<title>JS�б�</title>"
			.Write "<link href=""Admin_Style.CSS"" rel=""stylesheet"">"
			.Write "<script language=""JavaScript"">"
			.Write "var FolderID='" & FolderID & "';         //Ŀ¼ID" & vbCrLf
			.Write "var ParentID='" & ParentID & "'; //����ĿID" & vbCrLf
			.Write "var Page='" & CurrentPage & "';   //��ǰҳ��" & vbCrLf
			.Write "var KeyWord='" & KeyWord & "';    //�ؼ���" & vbCrLf
			.Write "var SearchParam='" & SearchParam & "';  //������������" & vbCrLf
			.Write "var Action='" & Action & "';" & vbCrLf
			.Write "var JSID='" & JSID & "';" & vbCrLf
			.Write "var JSType=" & JSType & ";" & vbCrLf
			.Write "</script>" & vbCrLf
		    .Write "<script language=""JavaScript"" src=""../../ks_inc/Common.js""></script>"
		    .Write "<script language=""JavaScript"" src=""../../ks_inc/jQuery.js""></script>"
		    .Write "<script language=""JavaScript"" src=""../../ks_inc/Kesion.Box.js""></script>"
			.Write "<script language=""JavaScript"" src=""ContextMenu.js""></script>"
			.Write "<script language=""JavaScript"" src=""SelectElement.js""></script>"
			%>
			<script language="javascript">
			var DocElementArrInitialFlag=false;
			var DocElementArr = new Array();
			var DocMenuArr=new Array();
			var SelectedFile='',SelectedFolder='';
			function document.onreadystatechange()
			{
				if (DocElementArrInitialFlag) return;
				InitialDocElementArr('FolderID','JSID');
				InitialDocMenuArr();
				 DocElementArrInitialFlag=true;
			}
			function InitialDocMenuArr()
			{  
				if (KeyWord=='')
				{
				 DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.CreateFolder();",'�½�Ŀ¼(N)','disabled');
				 DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.AddJS('');",'�½�JS(M)','disabled');
				 DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
				}
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.JSView();",'Ԥ ��(V)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.SelectAllElement();",'ȫ ѡ(A)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.Edit('');",'�� ��(E)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.Delete('');",'ɾ ��(D)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem('parent.ChangeUp();','�� ��(B)','');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem('parent.Reload();','ˢ ��(Z)','');
			}
			function DocDisabledContextMenu()
			{
			   var TempDisabledStr=''; 
			   if (FolderID=='0') TempDisabledStr='�� ��(B),';
				DisabledContextMenu('FolderID','JSID',TempDisabledStr+'Ԥ ��(V),�� ��(E),ɾ ��(D)','Ԥ��(V),�� ��(E)','','Ԥ ��(V),�� ��(E)','','')
			}
			function ChangeUp()
			{
			 if (FolderID=='0') return;
			 location.href='JS_Main.asp?JSType='+JSType+'&FolderID='+ParentID;
			   if (JSType==0)
				  $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr=JS ���� >> ϵͳ JS&ButtonSymbol=SysJSList&LabelFolderID='+ParentID;
			   else
				 $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr=JS ���� >> ���� JS&ButtonSymbol=FreeJSList&LabelFolderID='+ParentID;
			 }
			function OpenFolder(FolderID)
			{
			 location.href='JS_Main.asp?JSType='+JSType+'&FolderID='+FolderID;
			   if (JSType==0)
				 $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr=JS ���� >> ϵͳ JS&ButtonSymbol=SysJSList&LabelFolderID='+FolderID;
			   else
				$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr=JS ���� >> ���� JS&ButtonSymbol=FreeJSList&LabelFolderID='+FolderID;
			}
			function CreateFolder()
			{  
			  if (JSType==0)
				  OpenWindow('LabelFrame.asp?Url=LabelFolder.asp&PageTitle='+escape('�½�ϵͳJSĿ¼')+'&LabelType=2&FolderID='+FolderID,450,300,window);
			  else
				 OpenWindow('LabelFrame.asp?Url=LabelFolder.asp&PageTitle='+escape('�½�����JSĿ¼')+'&LabelType=3&FolderID='+FolderID,450,300,window);
			 Reload('');
			}
			function AddJS(TempUrl)
			{
			  if (JSType==0)
				{
				 location.href=TempUrl+'JS/AddSysJS.asp?FolderID='+FolderID+'&JSType="'+JSType+'&Action=AddNew';
				  $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr=JS ���� >> <font color=red>���ϵͳ JS</font>&ButtonSymbol=JSAdd';
				 }
			  else
				{location.href=TempUrl+'JS/AddFreeJS.asp?FolderID='+FolderID+'&Action='+Action+'&JSID='+JSID+'&JSType='+JSType
				 $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr=JS ���� >> <font color=red>������� JS</font>&ButtonSymbol=JSAdd';
				 }
			}
			function EditJS(TempUrl,ID)
			{  if (KeyWord=='')
				{   if (JSType==0)
					  {
					   location.href=TempUrl+'EditJS.asp?Page='+Page+'&JSType='+JSType+'&Action=Edit&JSID='+ID;
					   $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr=JS���� >> <font color=red>�޸�ϵͳJS</font>&ButtonSymbol=JSEdit';
					  }
				   else
				   {
					 location.href=TempUrl+'EditJS.asp?Page='+Page+'&JSType=1&Action=Edit&JSID='+ID;
					 $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr=JS���� >> <font color=red>�޸�����JS</font>&ButtonSymbol=JSEdit';
					}
				}
			   else
				 {  if (JSType==0)
					 {
					  location.href=TempUrl+'EditJS.asp?'+SearchParam+'&Page='+Page+'&JSType='+JSType+'&Action=Edit&JSID='+ID;
					  $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr=JS���� >> ����ϵͳJS��� >><font color=red>�޸�ϵͳJS</font>&ButtonSymbol=JSEdit';
					  }
				   else
					{
					 location.href=TempUrl+'EditJS.asp?'+SearchParam+'&Page='+Page+'&JSType=1&Action=Edit&JSID='+ID;
					 $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr=JS���� >> ��������JS��� >> <font color=red>�޸�����JS</font>&ButtonSymbol=JSEdit';
					 }
			  }
			}
			function EditFolder(ID)
			{
			 OpenWindow('LabelFrame.asp?Url=LabelFolder.asp&PageTitle='+escape('�༭Ŀ¼')+'&Action=EditFolder&FolderID='+ID,450,300,window);
			 Reload('');
			}
			function Edit(TempUrl)
			{   GetSelectStatus('FolderID','JSID');
				if (!((SelectedFile=='')&&(SelectedFolder=='')))
					{
						if (SelectedFolder!='')
						{ 
						if (TempUrl=='Folder'||TempUrl=='')
						 if (SelectedFolder.indexOf(',')==-1) 
						  {
						   EditFolder(SelectedFolder);
						 }
						else alert('һ��ֻ�ܹ��༭һ����ǩĿ¼');
					   }
					   if (SelectedFile!='')
						 {
						 if (TempUrl!='Folder'||TempUrl=='')
						 {	if (SelectedFile.indexOf(',')==-1) 
							 EditJS(TempUrl,SelectedFile);
						 else alert('һ��ֻ�ܹ��༭һ��JS');
			
						 }
						}
					}
				else 
				{
				alert('��ѡ��Ҫ�༭�ı�ǩ��Ŀ¼');
				}
			}
			function Delete(TempUrl)
			{   GetSelectStatus('FolderID','JSID');
				if (!((SelectedFile=='')&&(SelectedFolder=='')))
					{  
						if (confirm('ɾ��ȷ��:\n\n���Ҫִ��ɾ��������?'))
						  { if (SelectedFolder!='')
						   if (TempUrl=='Folder'||TempUrl=='')
							location='JS_Main.asp?JSAction=JSFolderDel&ID='+SelectedFolder;
						  if (SelectedFile!='')  
							if (TempUrl!='Folder'||TempUrl=='')
						location=TempUrl+'JS_Main.asp?JsAction=JSDel&Page='+Page+'&JSID='+SelectedFile;
						}	
					}
				else alert('��ѡ��Ҫɾ���ı�ǩĿ¼���ǩ');
			   SelectedFile='';
			   SelectedFolder='';
			}
			
			function GetKeyDown()
			{
			if (event.ctrlKey)
			  switch  (event.keyCode)
			  {  case 90 :  Reload(); break;
				 case 78 : event.keyCode=0;event.returnValue=false; CreateFolder();break;
				 case 77 : event.keyCode=0;event.returnValue=false; AddJS('');break;
				 case 65 : SelectAllElement();break;
				 case 66 : event.keyCode=0;event.returnValue=false;ChangeUp();break;
				 case 69 : event.keyCode=0;event.returnValue=false;Edit('');break;
				 case 68 : Delete('');break;
				 case 86 : JSView();break;
				 case 70 : event.keyCode=0;event.returnValue=false;
				   if (JSType==0)
					parent.frames['LeftFrame'].initializeSearch('SysJS')
				   else
					parent.frames['LeftFrame'].initializeSearch('FreeJS')
			 }	
			else if (event.keyCode==46)
			Delete('');
			}
			function Reload()
			{
			 location.href='js_Main.asp?FolderID='+FolderID+'&JSType='+JSType+'&'+SearchParam
			}
			function JSView()
			{   GetSelectStatus('FolderID','JSID');
				if (SelectedFile!='')
				{
				 window.open('LabelFrame.asp?Url=JS_Main.asp&JSAction=JSView&JSID='+SelectedFile+'&PageTitle='+escape('Ԥ��JS��ʾЧ��'),'new','width=620,height=450');
				 SelectedFile='';
				 }
				else
				 alert('��ѡ����ҪԤ����JS!')
			}

			</script>
			<%
			.Write "</head>"
			.Write "<body scroll=no topmargin=""0"" leftmargin=""0"" onclick=""SelectElement();"" onkeydown=""GetKeyDown();"" onselectstart=""return false;"">"
		    .Write "<ul id='menu_top'>"
				 If KeyWord = "" Then
			.Write "<li class='parent' onclick=""AddJS('');""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/a.gif' border='0' align='absmiddle'>���JS</span></li>"
			.Write "<li class='parent' onclick=""CreateFolder();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/a.gif' border='0' align='absmiddle'>���Ŀ¼</span></li>"
			.Write "<li class='parent' onclick=""Edit('Folder');""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/as.gif' border='0' align='absmiddle'>�༭Ŀ¼</span></li>"
			.Write "<li class='parent' onclick=""Delete('Folder');""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/del.gif' border='0' align='absmiddle'>ɾ��Ŀ¼</span></li>"

			 .Write "<li class='parent' onclick=""parent.frames['LeftFrame'].initializeSearch("
			 If JSType = 0 Then .Write ("'ϵͳ JS',0,'SysJS'") Else .Write ("'���� JS',0,'FreeJS'")
			 .Write ");""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/s.gif' border='0' align='absmiddle'>��������</span></li>"
			 .Write "<li class='parent' onclick=""ChangeUp();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/verify.gif' border='0' align='absmiddle'>����һ��</span></li>"
				 
				 Else
					  If JSType = 0 Then
					   .Write ("<img src='../Images/home.gif' align='absmiddle'><span style='cursor:pointer' onclick=""SendFrameInfo('JS_Main.asp?JSType=0','Template_Left.asp','../KS.Split.asp?ButtonSymbol=SysJSList&OpStr=JS���� >> <font color=red>ϵͳJS����</font>')"">ϵͳJS��ҳ</span>")
					Else
					   .Write ("<img src='../Images/home.gif' align='absmiddle'><span style='cursor:pointer' onclick=""SendFrameInfo('JS_Main.asp?JSType=1','Template_Left.asp','../KS.Split.asp?ButtonSymbol=FreeJSList&OpStr=JS���� >> <font color=red>����JS����</font>')"">����JS��ҳ</span>")
					End If
				   .Write (">>> �������: ")
					 If StartDate <> "" And EndDate <> "" Then
						.Write ("JS���������� <font color=red>" & StartDate & "</font> �� <font color=red> " & EndDate & "</font>&nbsp;&nbsp;&nbsp;&nbsp;")
					 End If
					Select Case SearchType
					 Case 0
					  .Write ("���ƺ��� <font color=red>" & KeyWord & "</font> ��JS")
					 Case 1
					  .Write ("�����к��� <font color=red>" & KeyWord & "</font> ��JS")
					 Case 2
					  .Write ("�ļ����к��� <font color=red>" & KeyWord & "</font> ��JS")
					 End Select
			End If
			
			.Write "    </ul>"

			.Write "<div style="" height:98%; overflow: auto; width:100%"" align=""center"">"
			.Write "<table width=""100%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
			.Write "  <tr>"
			.Write "    <td  valign=""top"">"
			.Write "      <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
			.Write "        <tr>"
			.Write "          <td height=""8"" align=""center""></td>"
			.Write "        </tr>"
				   
					Dim FolderSql, Param
					 Param = " Where JsType=" & JSType
					If KeyWord <> "" Then
					   FolderSql = "SELECT ID,FolderName,Description,OrderID FROM [KS_LabelFolder] Where 1=0"
					  Select Case SearchType
						Case 0
						  Param = Param & " AND JSName like '%" & KeyWord & "%'"
						Case 1
						 Param = Param & " AND Description like '%" & KeyWord & "%'"
						Case 2
						 Param = Param & " AND JSFileName like '%" & KeyWord & "%'"
					  End Select
					  If StartDate <> "" And EndDate <> "" Then
						   Param = Param & " And (AddDate>=#" & StartDate & "# And AddDate<=#" & DateAdd("d", 1, EndDate) & "#)"
					  End If
					Else
					   Param = Param & " AND FolderID='" & FolderID & "'"
					   FolderSql = "SELECT ID,FolderName,Description,OrderID FROM [KS_LabelFolder] Where FolderType=" & JSType + 2 & " And ParentID='" & FolderID & "'"
					End If
					Param = Param & " ORDER BY OrderID"
			Set JSRS = Server.CreateObject("ADODB.recordset")
			JSRS.Open FolderSql & " UNION  Select JSID,JSName,Description,OrderID From KS_JSFile " & Param, Conn, 1, 1
			If JSRS.EOF And JSRS.BOF Then
					 Else
						totalPut = JSRS.RecordCount
								If CurrentPage < 1 Then
									CurrentPage = 1
								End If
			
								If (CurrentPage - 1) * MaxPerPage > totalPut Then
									If (totalPut Mod MaxPerPage) = 0 Then
										CurrentPage = totalPut \ MaxPerPage
									Else
										CurrentPage = totalPut \ MaxPerPage + 1
									End If
								End If
			
								If CurrentPage = 1 Then
									Call showContent
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										JSRS.Move (CurrentPage - 1) * MaxPerPage
										
										Call showContent
									Else
										CurrentPage = 1
										Call showContent
									End If
								End If
							   
				End If
			 .Write "  </table>"
			 .Write "     </td>"
			 .Write "   </tr>"
			.Write "  </table>"
			.Write "  </di>"
			.Write "  </body>"
			.Write "  </html>"
			
			Set JSRS = Nothing
			Set Conn = Nothing
			End With
			End Sub
			
			   Sub showContent()
				Do While Not JSRS.EOF
			Response.Write "      <tr>"
			Response.Write "    <td>"
			Response.Write "    <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
			Response.Write "        <tr>"
					 
					  Dim T, TitleStr, JSName, ShortName, LabelTypeStr
						   For T = 1 To Row
						  If Not JSRS.EOF Then
								  JSName = JSRS(1)
								  ShortName = KS.ListTitle(Replace(Replace(JSName, "{JS_", ""), "}", ""), 24)
								  If JSType = 1 Then
									 LabelTypeStr = "����JS"
								   Else
									 LabelTypeStr = "ϵͳJS"
								  End If
								  TitleStr = " TITLE='�� ��:" & JSName & "&#13;&#10;�� ��:" & LabelTypeStr & "&#13;&#10;�� ��:" & JSRS("Description") & "'"
							   Response.Write ("<td width=""" & CInt(100 / Row) & "%"" Style=""cursor:default"" align=""center""" & TitleStr & ">")
							  If JSRS(3) = 0 Then
							   Response.Write ("<span onmousedown=""mousedown(this);""  FolderID=""" & JSRS(0) & """ style=""POSITION:relative;"" onDblClick=""OpenFolder(this.FolderID);""> ")
							 Else
							   Response.Write ("<span onmousedown=""mousedown(this);""  JSID=""" & JSRS(0) & """ style=""POSITION:relative;""  onDblClick=""Edit('');""> ")
							 End If
							 If JSRS(3) = 0 Then
								 Response.Write ("<img src=""../Images/Folder/folder.gif""> ")
							 Else
								 Response.Write ("<img src=""../Images/Label/JS" & JSType & ".gif"">")
							 End If
						   Response.Write ("<span style=""display:block;height:16;padding:0px 0px 0px 0px;margin:1px;width:80%;cursor:default"">" & ShortName & "</span>")
						   Response.Write ("</span>")
						   Response.Write ("</td>")
						i = i + 1
						  If JSRS.EOF Or i >= MaxPerPage Then Exit For
						   JSRS.MoveNext
						 Else
						  Exit For
						 End If
					Next
					'����7����Ԫ��,����в���
					Do While T <= Row
					 Response.Write ("<td width=70>&nbsp;</td>")
					 T = T + 1
					 Loop
					  
			Response.Write "        </tr>"
			Response.Write "        <tr><td colspan=" & Row & " height=10></td></tr>"
			Response.Write "      </table></td>"
			Response.Write "  </tr>"
			 
					  If i >= MaxPerPage Then Exit Do
					  If JSRS.EOF Then Exit Do
					Loop
					  JSRS.Close
						 Conn.Close
			
			Response.Write "        <td   align=""right"">"
				   
					 Call KS.ShowPageParamter(totalPut, MaxPerPage, "JS_Main.asp", True, "��", CurrentPage, "JSType=" & JSType & "&" & SearchParam)
			Response.Write ("</td>")
			Response.Write "</tr>"
		End Sub
		
		'ɾ��JS
		Sub JSDel()
		 Dim K, JSID, Page,RS,ArticleRS,JSType, CurrPath, JSFileName, JSDir, FolderID
		Set RS=Server.CreateObject("ADODB.Recordset")
		Set ArticleRS=Server.CreateObject("ADODB.Recordset")
		Page = Trim(KS.G("Page"))
		JSID = Split(KS.G("JSID"), ",") '���Ҫɾ����ǩ��ID����
		For K = LBound(JSID) To UBound(JSID)
		  RS.Open "SELECT * FROM [KS_JSFile] WHERE JSID='" & JSID(K) & "'", Conn, 1, 3
		  If Not RS.EOF Then
			JSType = RS("JSType")
			FolderID = RS("FolderID")
			  'ɾ������JS�ļ�
			  JSFileName = Trim(RS("JSFileName"))
			  JSDir = Trim(KS.Setting(93))
			  If Left(JSDir, 1) = "/" Or Left(JSDir, 1) = "\" Then JSDir = Right(JSDir, Len(JSDir) - 1)
			  CurrPath = KS.Setting(3) & JSDir
			  Call KS.DeleteFile(CurrPath & JSFileName)
			  '��������ɾ����JSID
			  ArticleRS.Open "Select  JSID From KS_Article Where JSID like '%" & JSID(K) & "%'", Conn, 1, 3
			  If Not ArticleRS.EOF Then
				 While Not ArticleRS.EOF
					ArticleRS(0) = Replace(ArticleRS(0), JSID(K) & ",", "")
					ArticleRS.Update
					ArticleRS.MoveNext
				 Wend
			  End If
		  End If
		 RS.Delete:RS.Close:ArticleRS.Close
		Next
		Set RS = Nothing:Set ArticleRS = Nothing
		Response.Redirect "JS_Main.asp?Page=" & Page & "&JSType=" & JSType & "&FolderID=" & FolderID
		End Sub
		
		'ɾ��JSĿ¼
		Sub JSFolderDel()
		   Dim RS,K, ID, ParentID, FolderSql,LabelFolderID,LabelType
		   Set RS=Server.CreateObject("ADODB.Recordset")
		   ID = Split(Request("ID"), ",")     '���Ҫɾ��Ŀ¼��ID����
			For K = LBound(ID) To UBound(ID)
			  FolderSql = "select ID,ParentID,FolderType from [KS_LabelFolder] where ID='" & ID(K) & "'"
			  RS.Open FolderSql, Conn, 1, 1
			  If Not RS.EOF Then
				LabelFolderID = Trim(RS(0))
				ParentID = Trim(RS(1))
				LabelType = RS(2)
						  Dim RSJS,JSDir
						  Set RSJS=Server.CreateObject("ADODB.Recordset")
						  'ɾ��JS�����ļ�
						  RSJS.Open "Select JSFileName From KS_JSFile Where FolderID='" & LabelFolderID & "'", Conn, 1, 1
								 JSDir = Trim(KS.Setting(93))
								If Left(JSDir, 1) = "/" Or Left(JSDir, 1) = "\" Then JSDir = Right(JSDir, Len(JSDir) - 1)
						  Do While Not RSJS.EOF
								Call KS.DeleteFile(KS.Setting(3) & JSDir & RSJS(0))
								RSJS.MoveNext
						  Loop
						  RSJS.Close
						  Set RSJS = Nothing
						  Conn.Execute ("DELETE  FROM KS_JSFILE WHERE FolderID='" & LabelFolderID & "'")
						  Conn.Execute ("DELETE  FROM KS_LabelFolder WHERE ID='" & LabelFolderID & "' OR TS like '%" & LabelFolderID & "%'")
			   End If
			  RS.Close
			Next
		 Set RS = Nothing
			Response.Write "<script>location.href='JS_Main.asp?JSType=" & (LabelType - 2) & "&Folderid=" & ParentID & "'</script>"
		End Sub
		
		'Ԥ��JS
		Sub JSView()
			Dim JSObj,JSID, JSdir,JSUrlStr
			JSID=Trim(Request.QueryString("JSID"))
			JSDir = KS.Setting(93)
			If Left(JSDir, 1) = "/" Or Left(JSDir, 1) = "\" Then JSDir = Right(JSDir, Len(JSDir) - 1)
			Set JSObj=Server.CreateObject("Adodb.Recordset")
			JSObj.OPEN "Select JSConfig,JSType,JSFileName From KS_JSFile Where JSID='" & JSID & "'",Conn,1,1
			IF JSObj.EOf AND JSObj.BOF THEN
			  Response.Write("�������ݳ���!")
			  JSObj.Close
			  Set JSObj=Nothing
			  Response.End
			ELSE
			  IF (trim(Split(JSObj("JSConfig"),",")(0))="GetExtJS" Or JSObj("JSType")=0) or (Request.QueryString("CanView")="1") Then
			  JSUrlStr="<script language=""javascript"" src=""" & KS.GetDomain & JSDir & Trim(JSObj("JSFileName")) & """></script>"
			  Else
				JSObj.Close:Set JSObj=Nothing
				Response.Redirect "JSFreeView.asp?JSID=" &JSID
			  End IF
			END IF
			JSObj.close:Set JSObj=Nothing
			%>
			<html>
			<head>
			<meta http-equiv="Expires" CONTENT="0">        
			<meta http-equiv="Cache-Control" CONTENT="no-cache">        
			<meta http-equiv="Pragma" CONTENT="no-cache">      
			<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
			<link href="Admin_Style.CSS" rel="stylesheet">
			<title>JSԤ��</title>
			<script language="JavaScript" src="Common.js"></script>
			</head>
			<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  bgcolor="#F1FAFA">
			<br>
			<table width="100%" height="70%" border="0" cellpadding="0" cellspacing="0">
			  <tr> 
				<td align="center"  valign="top"><table width="90%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
					<tr>
					  <td align="center" valign="top"><%=JSUrlStr%></td>
					</tr>
				  </table></td>
			  </tr>
			</table>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
			  <tr> 
				<td height="25"><strong> ��������������˵��:</strong></td>
			  </tr>
			  <tr>
				<td height="25">���������������������������JS����������ʽ,��ô�����Ԥ��Ч�����ܻ���ʵ���е���</td>
			  </tr>
			  <tr> 
				<td height="25">���������������������������������Ч�����뵥��ˢ�°�ť <input type="button" value="ˢ��" onClick="window.location.reload()"><input type="button" value="�ر�" onClick="window.parent.close()">
				<%if Request.QueryString("CanView")="1" then
				  Response.Write("<INPUT TYPE=BUTTON value=""����"" onclick=""history.back();"">")
				  End IF
				  %></td>
			  </tr>
			</table>
			</body>
			</html>
      <%
		End Sub
End Class
%> 
