<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Admin_Field
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Field
        Private KS,Action,ChannelID,Page,ItemName,TableName,KSCls
		Private I, totalPut, CurrentPage, FieldSql, FieldRS,MaxPerPage
		Private FieldName,ID,Contact, Title, Tips, FieldType, DefaultValue, MustFillTF, ShowOnForm, ShowOnUserForm,Options,OrderID,AllowFileExt,MaxFileSize,Width,Height,EditorType,ShowUnit,UnitOptions,ParentFieldName

		Private Sub Class_Initialize()
		  MaxPerPage = 30
		  Set KSCls=New ManageCls
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub


		Public Sub Kesion()
		With Response
		.Write "<html>"
		.Write "<head>"
		.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
		.Write "<title>�ֶι���</title>"
		.Write "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
             Action=KS.G("Action")
			 ChannelID=KS.ChkClng(KS.G("ChannelID"))
			 
			 TableName=KS.C_S(ChannelID,2)
			 If ChannelID=101 Then
			  TableName="KS_User"   '��Ա��
			 End If
			 ItemName=KS.C_S(ChannelID,3)
			 Page=KS.G("Page")

			 if ChannelID=101 Then
		       If Not KS.ReturnPowerResult(0, "KMUA10012")  Then          '���Ȩ��
					 Call KS.ReturnErr(1, "")
					 .End
			   End If
			 Else
		       If Not KS.ReturnPowerResult(0, "KSMM10003") Then          '���Ȩ��
					 Call KS.ReturnErr(1, "")
					 .End
			   End If
			 End If
			 
			 Select Case Action
			  Case "Add"
			    Call FieldAddOrEdit("Add")
			  Case "Edit"
			    Call FieldAddOrEdit("Edit")
			  Case "Del"
			    Call FieldDel()
			  Case "order"
			    Call FieldOrder()
			  Case "AddSave"
			    Call FieldAddSave()
			  Case "EditSave"
			    Call FieldEditSave()
			  Case Else
			   Call FieldList()
			 End Select
			.Write "</body>"
			.Write "</html>"
		 End With
		End Sub
		
		Sub FieldList()
		 On Error Resume Next
		If Not IsEmpty(KS.G("page")) Then
			  CurrentPage = KS.G("page")
		Else
			  CurrentPage = 1
		End If
		With Response
		.Write "<script language='JavaScript'>"
		.Write "var Page='" & CurrentPage & "';"
		.Write "var ItemName='" & ItemName & "';"
		.Write "var ChannelID=" & ChannelID & ";"
		.Write "</script>"
		.Write "<script language='JavaScript' src='../KS_Inc/common.js'></script>"
		.Write "<script language='JavaScript' src='../KS_Inc/jquery.js'></script>"
		.Write "<script language='JavaScript' src='Include/ContextMenu1.js'></script>"
		.Write "<script language='JavaScript' src='Include/SelectElement.js'></script>"
		%>
		 <script language="javascript">
		 var DocElementArrInitialFlag=false;
		var DocElementArr = new Array();
		var DocMenuArr=new Array();
		var SelectedFile='',SelectedFolder='';
		$(document).ready(function()
		{  
		    if (DocElementArrInitialFlag) return;
			InitialDocElementArr('FolderID','FieldID');
			InitialContextMenu();
			DocElementArrInitialFlag=true;
		});
		function InitialContextMenu()
		{	DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.FieldAdd();",'�� ��(N)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.SelectAllElement();",'ȫ ѡ(A)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.FieldControl(1);",'�� ��(E)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.FieldControl(2);",'ɾ ��(D)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.location.reload();",'ˢ ��(Z)','disabled');
		}
		function DocDisabledContextMenu()
		{
			DisabledContextMenu('FolderID','FieldID','�� ��(E),ɾ ��(D)','�� ��(E)','','','','')
		}
		function FieldAdd()
		{
		   location.href='KS.Field.asp?ChannelID='+ChannelID+'&Action=Add';
		   window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape('ģ�͹��� >> ģ���ֶι��� >> <font color=red>����'+ItemName+'�Զ����ֶ�</font>')+'&ButtonSymbol=Go';
		}
		function EditField(id)
		{
		  location="KS.Field.asp?ChannelID="+ChannelID+"&Page="+Page+"&Action=Edit&ID="+id;
		  window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape('ģ�͹��� >> ģ���ֶι��� >> <font color=red>�༭'+ItemName+'�Զ����ֶ�</font>')+'&ButtonSymbol=GoSave';
		}
		function DelField(id)
		{
		if (confirm('���Ҫɾ�����Զ����ֶ���?'))
		 location="KS.Field.asp?ChannelID="+ChannelID+"&Action=Del&Page="+Page+"&id="+id;
		  SelectedFile='';
		}
		function FieldControl(op)
		{   var alertmsg='';
			GetSelectStatus('FolderID','FieldID');
			if (SelectedFile!='')
			 {
			   if (op==1)
				{
				if (SelectedFile.indexOf(',')==-1) 
					EditField(SelectedFile)
				  else alert('һ��ֻ�ܱ༭һ���Զ����ֶ�!')	
				SelectedFile='';
				}	
			  else if (op==2)    
			   DelField(SelectedFile);
			 }
			else 
			 {
			 if (op==1)
			  alertmsg="�༭";
			 else if(op==2)
			  alertmsg="ɾ��"; 
			 else
			  {
			  alertmsg="����" 
			  }
			 alert('��ѡ��Ҫ'+alertmsg+'���Զ����ֶ�');
			  }
		}
		function GetKeyDown()
		{ 
		if (event.ctrlKey)
		  switch  (event.keyCode)
		  {  case 90 : location.reload(); break;
			 case 65 : SelectAllElement();break;
			 case 78 : event.keyCode=0;event.returnValue=false;FieldAdd();break;
			 case 69 : event.keyCode=0;event.returnValue=false;FieldControl(1);break;
			 case 68 : FieldControl(2);break;
		   }	
		else	
		{
		 //if (event.keyCode==46)FieldControl(2);
		 }
		}
		 </script>
		<%
		.Write "</head>"
		.Write "<body scroll=no topmargin='0' leftmargin='0' onclick='SelectElement();' onkeydown='GetKeyDown();' onselectstart='return false;'>"
		.Write "<ul id='menu_top'>"
		.Write "<li class='parent' onclick=""FieldAdd();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>�����ֶ�</span></li>"
		.Write "<li class='parent' onclick=""FieldControl(1);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>�޸��ֶ�</span></li>"
		.Write "<li class='parent' onclick=""FieldControl(2);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>ɾ���ֶ�</span></li>"
		.Write "<li class='parent' onclick=""location.href='KS.Model.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/back.gif' border='0' align='absmiddle'>����һ��</span></li>"
		.Write "</ul>"
        
		.Write ("<div style=""height:94%; overflow: auto; width:100%"" align=""center"">")
		.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
		.Write "<form action='KS.Field.asp?action=order&channelid=" & ChannelID&"&page="&CurrentPage &"'' name='form1' method='post'>"
		.Write "        <tr class='sort'>"
		.Write "         <td width='80' align='center'>����</td>"
		.Write "          <td width='100' align='center'>�ֶ�����</td>"
		.Write "          <td align='center'>�ֶα���</td>"		
		.Write "          <td align='center'>����ģ��</td>"
		.Write "          <td align='center'>�ֶ�����</td>"
		.Write "          <td align='center'>��̨��ʾ</td>"
		.Write "          <td align='center'>ǰ̨��ʾ</td>"
		.Write "          <td align='center'>���������</td>"
		.Write "        </tr>"
			 Set FieldRS = Server.CreateObject("ADODB.RecordSet")
				   FieldSql = "SELECT * FROM KS_Field Where ChannelID=" & ChannelID & " order by orderid asc"
				   FieldRS.Open FieldSql, conn, 1, 1
				 If FieldRS.EOF And FieldRS.BOF Then
				 Else
					totalPut = FieldRS.RecordCount
		
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
									FieldRS.Move (CurrentPage - 1) * MaxPerPage
									Call showContent
								Else
									CurrentPage = 1
									Call showContent
								End If
							End If
			End If
		 .Write " <tr>"
		 .Write "   <td colspan='3'>&nbsp;&nbsp;<input type='submit' class='button' value='���������ֶ�����'> <font color=blue>ֵԽС����Խǰ��</font></td></form>"
		 .Write "   <td height='35' colspan='5' align='right'>"
		 Call KSCLS.ShowPage(totalPut, MaxPerPage, "KS.Field.asp", True, "��",CurrentPage, "ChannelID=" & ChannelID)
		.Write "    </td>"
		.Write " </tr>"
		.Write "</table>"
		.Write "</div>"
		End With
		End Sub
		Sub showContent()
		With Response
		Do While Not FieldRS.EOF
		 .Write "<tr>"
		 .Write "<td class='splittd'>&nbsp;&nbsp;<input type='text' name='OrderID' style='width:50px;text-align:center' value='" & FieldRS("OrderID") &"'><input type='hidden' name='FieldID' value='" & FieldRS("FieldID") & "'></td>"
		 .Write "  <td class='splittd'><span FieldID='" & FieldRS("FieldID") & "' onDblClick=""EditField(this.FieldID)""><img src='Images/Field.gif' align='absmiddle'><span  style='cursor:default;'>" & FieldRS("FieldName") & "</span></span></td>"
		 .Write "   <td align='center' class='splittd'>" & FieldRS("Title") & " </td>"
		 .Write "   <td align='center' class='splittd'><font color=red>"
		 If ChannelID=101 Then
		 .Write "��Աϵͳ"
		 Else
		  .Write KS.C_S(ChannelID,1) 
		 End If
		  .Write "</font>"
		 .Write "</td>"
		 .Write "   <td align='center' class='splittd'>"
				 Select Case FieldRS("FieldType")
				  Case 1:.Write "�����ı�(text)"
				  Case 2:.Write "�ı�(��֧��HTML)"
				  Case 10:.Write "�����ı�(֧��HTML)"
				  Case 3:.Write "�����б�(select)"
				  Case 4:.Write "����(text)"
				  Case 5:.Write "����(text)"
				  Case 6:.Write "��ѡ��(radio)"
				  Case 7:.Write "��ѡ��(checkbox)"
				  Case 8:.Write "��������(text)"
				  Case 9:.Write "�ļ�(text)"
				  Case 11:.Write "�����˵�(text)"
				 End Select
		  If Left(Lcase(FieldRS("FieldName")),3)<>"ks_" Then .Write "<font color=#cccccc>[ϵͳ]</font>"
		 .Write "</td>"
		 .Write "   <td align='center' class='splittd'>" 
		  If FieldRS("ShowOnForm")=1 Then
		   .Write "<font color=red>��</font>"
		  Else
		   .Write "<font color=green>��</font>"
		  End If
		 .Write " </td>"
		 .Write "   <td align='center' class='splittd'>" 
		  If FieldRS("ShowOnUserForm")=1 Then
		   .Write "<font color=red>��</font>"
		  Else
		   .Write "<font color=green>��</font>"
		  End If
		 .Write " </td>"
		 .Write " <td align='center' class='splittd'><a href='javascript:EditField(" & FieldRS("FieldID") &");'>�޸�</a> | "
		 If Left(Lcase(FieldRS("FieldName")),3)<>"ks_" Then
		 .Write "<font color=#cccccc title='ϵͳ�ֶβ�����ɾ��'>ɾ��</font>"
		 Else
		 .Write "<a href='javascript:DelField(" & FieldRS("FieldID") &");'>ɾ��</a>"
		 End If
		 .Write " </td></tr>"
								I = I + 1
								If I >= MaxPerPage Then Exit Do
							   FieldRS.MoveNext
							   Loop
								FieldRS.Close
						 
         End With
		 End Sub
		 
		 Sub FieldAddOrEdit(OpType)
		 With Response
		  Dim FieldRS, FieldSql,OpAction,OpTempStr
		 ID = KS.G("ID")
		.Write "<html>"
		.Write "<head>"
		.Write "<meta http-equiv='Content-Type' content='text/html; chaRSet=gb2312'>"
		.Write "<title>�ֶι���</title>"
		.Write "<link href='../Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		.Write "<script language='JavaScript' src='../KS_Inc/common.js'></script>"
		.Write "</head>"
		.Write "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
		
		 If Optype = "Edit" Then
		     OpAction="EditSave":OpTempStr="�༭"
			 Set FieldRS = Server.CreateObject("ADODB.RECORDSET")
			 FieldSql = "Select TOP 1 * From [KS_Field] Where FieldID=" & ID
			 FieldRS.Open FieldSql, conn, 1, 1
			 If Not FieldRS.EOF Then
				 FieldName = Trim(FieldRS("FieldName"))
				 ChannelID = FieldRS("ChannelID")
				 Title = Trim(FieldRS("Title"))
				 Tips = Trim(FieldRS("Tips"))
				 FieldType = Trim(FieldRS("FieldType"))
				 DefaultValue = Trim(FieldRS("DefaultValue"))
				 MustFillTF = FieldRS("MustFillTF")
				 ShowOnForm = FieldRS("ShowOnForm")
				 ShowOnUserForm=FieldRS("ShowOnUserForm")
				 Options = Trim(FieldRS("Options"))
				 OrderID= FieldRS("OrderID")
				 AllowFileExt=FieldRS("AllowFileExt")
				 MaxFileSize=FieldRS("MaxFileSize")
				 Width=FieldRS("Width")
				 Height=FieldRS("Height")
				 EditorType=FieldRS("EditorType")
				 ShowUnit=FieldRS("ShowUnit")
				 UnitOptions=FieldRS("UnitOptions")
				 ParentFieldName=FieldRS("ParentFieldName")
			 End If
	  Else
	     FieldName="KS_":FieldType=1:MustFillTF=0:ShowOnForm=1:ShowOnUserForm=1:OrderID=1:AllowFileExt="jpg|gif|png":MaxFileSize=1024:Width=200:Height=80:EditorType="Basic":ShowUnit=0
		 OpAction="AddSave":OpTempStr="���"
	  End If
		 
		.Write "<div class='topdashed sort'>" & OpTempStr &"�Զ����ֶ�</div>"
		.Write "<br>"
        .Write "        <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1""  class='ctable'>" & vbCrLf
		.Write "  <form  action='KS.Field.asp?Action=" & OpAction &"' method='post' name='OrigArticlerm' onsubmit='return(CheckForm())'>"
		
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' class='clefttitle'><strong>����ϵͳ��</strong></td>"
		.Write "      <td nowrap> &nbsp;&nbsp;<font color=#ff0000>"
		If ChannelID=101 Then
		.Write "��Աϵͳ"
		Else
		.Write KS.C_S(ChannelID,1)
		End If
		.Write "</font><input type='hidden' value='" & ChannelID & "' name='ChannelID'></td>"
		.Write "    </tr>"


		.Write "   <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td width='230' height='30' align='right' class='clefttitle'><strong>�ֶ����ƣ�</strong><br><font color=blue>Ϊ�˺�ϵͳ�ֶ����֣������ԡ�KS_����ͷ,��ģ���п���ͨ����{$KS_�ֶ�����}�����е���</font></td>"
		.Write "      <td height='45' nowrap>&nbsp;"
		.Write "        <input name='FieldName' type='text' id='FieldName' value='" & FieldName & "' size='30'"
		If Optype = "Edit" Then .Write " readonly"
		.Write " class='textbox'>"
		.Write "        * </td>"
		.Write "    </tr>"
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='45' align='right' class='clefttitle'><strong>�ֶα�����</strong><br><font color=blue>�����ڹ�����Ŀ����ʾ</font></td>"
		.Write "      <td height='45' nowrap>&nbsp;&nbsp;<input name='Title' type='text' id='Title' size='30' class='textbox' value='" & Title & "'> *</td>"
		.Write "    </tr>"
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right'' class='clefttitle'><strong>������ʾ��</strong><br><font color=blue>��������Աߵ���ʾ��Ϣ</font></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<textarea name='Tips'  id='Tips' class='textbox' style='width:300px;height:60px'>" & Tips & "</textarea><font color=green>���Լ���һЩjavascript�¼�</font></td>"
		.Write "    </tr>"
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>�ֶ����ͣ�</strong></td>"
		If Optype = "Edit" Then
		.Write "      <td nowrap>&nbsp;&nbsp;<input type='hidden' value=" & FieldType & " name='FieldType'><select name=""FieldType"" disabled>"
		else
		.Write "      <td nowrap>&nbsp;&nbsp;<select name=""FieldType"" onchange=""Setdisplay(this.value)"">"
		end if
		.Write " <option value=""1"""
		If FieldType=1 Then .Write " Selected"
		.Write ">�����ı�(text)</option>"
     	.Write " <option value=""2"""
		If FieldType=2 Then .Write " Selected"
		.Write ">�����ı�(��֧��HTML)</option>"
     	.Write " <option value=""10"""
		If FieldType=10 Then .Write " Selected"
		.Write ">�����ı�(֧��HTML)</option>"
		.Write " <option value=""3"""
		If FieldType=3 Then .Write " Selected"
		.Write ">�����б�(select)</option>"
		If ChannelID<>101 Then
			.Write " <option value=""11"""
			If FieldType=11 Then .Write " selected"
			.Write " style='color:blue'>���������б�</option>"
		End If
        .Write " <option value=""4"""
		If FieldType=4 Then .Write " Selected"
		.Write ">����(text)</option>"
		.Write " <option value=""5"""
		If FieldType=5 Then .Write " Selected"
		.Write ">����(text)</option>"
		.Write " <option value=""6"""
		If FieldType=6 Then .Write " Selected"
		.Write ">��ѡ��(radio)</option>"
		.Write " <option value=""7"""
		If FieldType=7 Then .Write " Selected"
		.Write ">��ѡ��(checkbox)</option>"
		.Write " <option value=""8"""
		If FieldType=8 Then .Write " Selected"
		.Write ">��������(text)</option>"
		.Write " <option value=""9"""
		If FieldType=9 Then .Write " Selected"
		.Write ">�ļ�(text)</option>"
		
		.Write " </select>"
		.Write "<font color=red>˵����һ���趨�����޸�</font>"
		.Write " </td>"
		.Write "    </tr>"
		
		.Write "  <tbody id='editorarea'>"
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>�༭�����ͣ�</strong></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<input name='EditorType' type='text' id='EditorType' class='textbox' value='" & EditorType & "' size='10'>&nbsp;<select onchange=""$('EditorType').value=this.value"" name='selecteditor'><option value='Default'>Default</option><option value='NewsTool'>NewsTool</option><option value='Simple'>Simple</option><option value='Basic'>Basic</option></select><span style='color:green'>�����Դ�/KS_Editor/fckeditor/fckconfig.js�Զ���༭������</span>"
		.Write "       </td>"
		.Write "    </tr>"
		.Write " </tbody>"
		
		
		.Write "<tbody id=""extarea"">"
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>�����ϴ�����չ����</strong></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<input name='AllowFileExt' type='text' id='AllowFileExt' class='textbox' value='" & AllowFileExt & "' size='40'>&nbsp;<span style='color:#ff0000'>�����չչ�������ö��š�|������</span>"
		.Write "       </td>"
		.Write "    </tr>"
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>�����ϴ����ļ���С��</strong></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<input name='MaxFileSize' type='text' id='MaxFileSize' class='textbox' value='" & MaxFileSize & "' size='8' style='width:50px'>&nbsp;KB <span style='color:#ff0000'>*</span>  <span style='color:blue'>��ʾ��1 KB = 1024 Byte��1 MB = 1024 KB<span>  "
		.Write "       </td>"
		.Write "    </tr>"
		.Write " </tbody>"
		
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>Ĭ��ֵ��</strong></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<input name='DefaultValue' type='text' id='DefaultValue' class='textbox' value='" & DefaultValue & "' size='40'>&nbsp;<span id='darea' style='color:#ff0000'>���Ĭ��ѡ����ö��š�,������</span>"
		If ChannelID<>101 Then
		 .Write "<div>&nbsp;&nbsp;<font color=green>Ϊ���ڻ�Ա��ȡĬ��ֵ���ɰ󶨱�KS_User��KS_Enterprise���ֶ�ֵ<br>&nbsp;&nbsp;��ʽ������|�ֶ��� �磺<font color=red>KS_User|RealName</font></font><br/>&nbsp;&nbsp;<font color=blue>Ҳ���Խ�Ĭ��ֵ����Ϊnow��dateȡ�õ�ǰʱ��</font></div>"
		End If
		.Write "       </td>"
		.Write "    </tr>"
		
		If ChannelID<>101 Then
		.Write "    <tr id=""ldArea"" style='display:none' class='tdbg'>"
		.Write "      <td align='right' class='clefttitle'><strong>���������ֶΣ�</strong><br><font color=blue>��ѡ���ʾһ�������ֶ�<br/>����ָ��Ϊ�¼������ֶ�</font></td>"
		.Write "      <td>&nbsp;&nbsp;"
		  Dim PRS
		  If KS.ChkClng(ID)<>0 Then
		  Set PRS=Conn.Execute("Select FieldName,Title From KS_Field Where ChannelID=" & ChannelID& " and FieldType=11 And FieldID<>" & ID & " Order BY FieldID")
		  .Write "<select name='ParentFieldName' disabled>"
		  Else
		  Set PRS=Conn.Execute("Select FieldName,Title From KS_Field Where ChannelID=" & ChannelID& " and FieldType=11 Order BY FieldID")
		  .Write "<select name='ParentFieldName'>"
		  End If
		  .Write "<option value='0'>--��Ϊһ������--</option>"
		  Do While Not PRS.Eof
		      If PRS(0)=ParentFieldName Then
		      .Write "<option value='" & PRS(0) & "' selected>" & Prs(1) & "(" & PRS(0) & ")</option>"
			  Else
		      .Write "<option value='" & PRS(0) & "'>" & Prs(1) & "(" & PRS(0) & ")</option>"
			  End If
		  PRS.MoveNext
		  Loop
		  PRS.Close: Set PRS=Nothing
		.Write "      </select> <font color=red>˵����һ���趨�����޸�</font></td>"
		.Write "    </tr>"
		End If
		
		.Write "    <tr id=""OptionsArea"" style=""display:none"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td align='right'  class='clefttitle'><strong>�б�ѡ�</strong><br><font color=blue>ÿһ��Ϊһ���б�ѡ��</font><br>���ֵ����ʾ�ͬ������<font color=red>|</font>����<br>��ȷ��ʽ�磺<font color=red>��</font> �� <font color=red>0|��</font><br></td>"
		.Write "      <td height='45' nowrap>&nbsp;&nbsp;<textarea name='Options' style='height:70px' cols='50' rows='6' id='Options' class='textbox'>" & Options & "</textarea>"
		.Write "      </td>"
		.Write "    </tr>"
		
		if channelid<>101 then
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>�Ƿ���ʾ������λ��</strong></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;"
		 If Optype = "Edit" Then
		    If ShowUnit=1 Then .Write "��" Else .Write "��"
			.Write "<input type='hidden' name='ShowUnit' value='1'>"
		 Else
			.Write  "<input onclick=""$('unitArea').style.display=''"" name='ShowUnit' type='radio' id='ShowUnit' value='1'"
			If ShowUnit=1 Then .Write " Checked"
			.Write ">��"
			.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input onclick=""$('unitArea').style.display='none'"" name='ShowUnit' type='radio' id='ShowUnit' value='0'"
			If ShowUnit=0 Then .Write " Checked"
			.WRite ">��"
		 End If
		 .Write "&nbsp;&nbsp;<font color=red>˵����һ���趨�����޸�</font>"
		.Write "       </td>"
		.Write "    </tr>"
		If ShowUnit=1 Then
		.Write "    <tr class=""tdbg"" id=""unitArea"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
	   else
		.Write "    <tr class=""tdbg"" id=""unitArea"" style=""display:none"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
	   end if
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>������λѡ�</strong><br/><font color=blue>ÿһ��Ϊһ���б�ѡ��<br/>��:�� ����</font></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<textarea name='UnitOptions' style='height:70px' cols='20' rows='6' id='UnitOptions' class='textbox'>" & UnitOptions & "</textarea> "
		.Write "       </td>"
		.Write "    </tr>"
		
		end if
		
		
		
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>�Ƿ���</strong></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<input name='MustFillTF' type='radio' id='MustFillTF' value='1'"
		If MustFillTF=1 Then .Write " Checked"
		.Write ">��"
		.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name='MustFillTF' type='radio' id='MustFillTF' value='0'"
		If MustFillTF=0 Then .Write " Checked"
		.WRite ">��"
		.Write "       </td>"
		.Write "    </tr>"
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>�Ƿ����ã�</strong></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<input name='ShowOnForm' type='radio' id='ShowOnForm' value='1'"
		If ShowOnForm=1 Then .Write " Checked"
		.Write ">��"
		.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowOnForm' type='radio' id='ShowOnForm' value='0'"
		If ShowOnForm=0 Then .Write " Checked"
		.WRite ">��"
		.Write "       </td>"
		.Write "    </tr>"
		'If ChannelID=101 Then 
		'.Write "    <tr style='display:none' "
		'Else
		.Write "    <tr "
		'End If
		.Write "class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>��Ա�����Ƿ����ã�</strong><br><font color=blue>���������ã�ǰ̨�Ļ�Ա���ĲŻ���ʾ</font></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<input name='ShowOnUserForm' type='radio' id='ShowOnUserForm' value='1'"
		If ShowOnUserForm=1 Then .Write " Checked"
		.Write ">��"
		.Write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name='ShowOnUserForm' type='radio' id='ShowOnUserForm' value='0'"
		If ShowOnUserForm=0 Then .Write " Checked"
		.WRite ">��"
		.Write "       </td>"
		.Write "    </tr>"
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>��ʾ���ã�</strong> </td>"
		.Write "      <td nowrap>&nbsp;&nbsp;���<input name='Width' type='text' site='10' class='textbox' style='width:40px' id='Width' value='" & Width & "'>px <font color=red>���磺200px</font><br><span style='display:none' id='heightarea'>&nbsp;&nbsp;�߶�<input name='Height' type='text' site='10' class='textbox' style='width:40px' id='Height' value='" & Height & "'>px <font color=red>���磺100px</font></span>"
		.Write "       </td>"
		.Write "    </tr>"		
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>������ţ�</strong><br><font color=blue>���ԽС������Խǰ��</font></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<input name='OrderID' type='text' site='35' class='textbox' id='OrderID' value='" & OrderID & "'>"
		.Write "       </td>"
		.Write "    </tr>"
		.Write "   <input type='hidden' value='" & ID & "' name='id'>"
		.Write "    <input type='hidden' value='" & Page & "' name='page'>"
		.Write "  </form>"
		.Write "</table>"
		
		
		 
		.Write "<Script Language='javascript'>"
		.Write "Setdisplay(" & FieldType & ");"
		.Write "function Setdisplay(s)"
		.Write  "{if (s==3||s==6||s==7||s==11){ $('OptionsArea').style.display='';} else $('OptionsArea').style.display='none';if (s==7)$('darea').style.display='';else $('darea').style.display='none';if(s==9)$('extarea').style.display='';else $('extarea').style.display='none'; if(s==10)$('editorarea').style.display='';else $('editorarea').style.display='none';if (s==2||s==10) $('heightarea').style.display='';else $('heightarea').style.display='none';if(s==11) $('ldArea').style.display=''; else $('ldArea').style.display='none';}"
		.Write "function CheckForm()"
		.Write "{ var form=document.OrigArticlerm;"
		.Write "   if (form.FieldName.value==''||form.FieldName.value.length<=1)"
		.Write "    {"
		.Write "     alert('�������ֶ�����!');"
		.Write "     form.FieldName.focus();"
		.Write "     return false;"
		.Write "    }"
		.Write "   if (form.Title.value=='')"
		.Write "    {"
		.Write "     alert('�������ֶα���!');"
		.Write "     form.Title.focus();"
		.Write "     return false;"
		.Write "    }"
		.Write "    form.submit();"
		.Write "    return true;"
		.Write "}"
		.Write "</Script>"
		End With
		End Sub
		 
		 Sub FieldAddSave()
		 Dim FieldRS,ColumnType
		 FieldName = Trim(KS.G("FieldName"))
		 Title = KS.G("Title")
		 Tips = Request.Form("Tips")
		 FieldType = KS.G("FieldType")
		 DefaultValue = KS.G("DefaultValue")
		 MustFillTF = KS.G("MustFillTF")
		 FieldType = KS.G("FieldType")
		 ShowOnForm = KS.G("ShowOnForm")
		 ShowOnUserForm=KS.G("ShowOnUserForm")
		 Options = KS.G("Options")
		 FieldType =KS.G("FieldType")
		 Width=KS.G("Width")
		 AllowFileExt=KS.G("AllowFileExt")
		 EditorType=KS.G("EditorType")
		 ShowUnit  =KS.ChkClng(KS.G("ShowUnit"))
		 UnitOptions=KS.G("UnitOptions")
		 ParentFieldName=KS.G("ParentFieldName")
		 If KS.IsNul(ParentFieldName) Then ParentFieldName="0"

		 If FieldName = "" Then Call KS.AlertHistory("�������ֶ�����!", -1): Exit Sub
		 If Len(FieldName)<=3 Then Call KS.AlertHistory("�ֶ����Ƴ��ȱ������3!", -1): Exit Sub
		 If Ucase(Left(FieldName,3))<>"KS_" Then Call KS.AlertHistory("�ֶ����Ƹ�ʽ���󣬱�����""KS_��ͷ""!", -1): Exit Sub
		 If Title="" Then Call KS.AlertHistory("�ֶα����������!", -1): Exit Sub
		 If FieldType=4 And Not Isnumeric(DefaultValue) And DefaultValue<>"" Then Call KS.AlertHistory("Ĭ��ֵ��ʽ����ȷ!", -1): Exit Sub
		 If FieldType=5 And Not IsDate(DefaultValue) And DefaultValue<>"" and lcase(DefaultValue)<>"now" and lcase(DefaultValue)<>"date" Then Call KS.AlertHistory("Ĭ��ֵ��ʽ����ȷ!", -1): Exit Sub
		 if FieldType=8 And Not KS.IsValidEmail(DefaultValue) And DefaultValue<>"" Then Call KS.AlertHistory("Ĭ�ϸ�ʽ����ȷ����������ȷ��Email!",-1):Exit Sub
		 Select Case FieldType
		   Case 1,3,6,7,8,9,11
		     ColumnType="nvarchar(255)"
		   Case 2,10
		     ColumnType="ntext"
		   Case 5
		     ColumnType="datetime"
		   Case 4
		     ColumnType="int"
		   Case else
		     Exit Sub
		 End Select
		 Set FieldRS = Server.CreateObject("ADODB.RECORDSET")
		 FieldSql = "Select * From [KS_Field] Where FieldName='" & FieldName & "' And ChannelID=" & KS.G("ChannelID")
		 FieldRS.Open FieldSql, conn, 3, 3
		 If FieldRS.EOF And FieldRS.BOF Then
		  FieldRS.AddNew
		  FieldRS("FieldName") = FieldName
		  FieldRS("ChannelID") = KS.G("ChannelID")
		  FieldRS("Title") = Title
		  FieldRS("Tips") = Tips
		  FieldRS("FieldType") = FieldType
		  FieldRS("DefaultValue") = DefaultValue
		  FieldRS("MustFillTF") = MustFillTF
		  FieldRS("FieldType") = FieldType
		  FieldRS("ShowUnit")=ShowUnit
		  FieldRS("UnitOptions")=UnitOptions
		  FieldRS("ShowOnForm") = ShowOnForm
		  FieldRS("ShowOnUserForm")=ShowOnUserForm
		  FieldRS("Options") = Options
		  FieldRS("OrderID")=KS.ChkClng(KS.G("OrderID"))
		  FieldRS("AllowFileExt")=KS.G("AllowFileExt")
		  FieldRS("MaxFileSize")=KS.ChkClng(KS.G("MaxFileSize"))
		  FieldRS("Width")=KS.ChkClng(KS.G("Width"))
		  FieldRS("Height")=KS.ChkClng(KS.G("Height"))
		  FieldRS("EditorType")=EditorType
		  FieldRS("ParentFieldName")=ParentFieldName
		  FieldRS.Update
		  Conn.Execute("Alter Table "&TableName&" Add "&FieldName&" "&ColumnType&"")
		  If ShowUnit=1 Then  '���ӵ�λ�ֶ�
		  Conn.Execute("Alter Table "&TableName&" Add "&FieldName&"_Unit nvarchar(200)")
		  End If
		  
		  If KS.C_S(KS.G("ChannelID"),6)=1 or KS.C_S(KS.G("ChannelID"),6)=2 or KS.C_S(KS.G("ChannelID"),6)=5 Then
		  KS.ConnItem.Execute("Alter Table "&TableName&" Add "&FieldName&" "&ColumnType&"")
		  If ShowUnit=1 Then  '���ӵ�λ�ֶ�
		  KS.ConnItem.Execute("Alter Table "&TableName&" Add "&FieldName&"_Unit nvarchar(200)")
		  End If
		  
		  End If
		 Response.Write ("<Script> if (confirm('�ֶ����ӳɹ�,���������?')) { location.href='KS.Field.asp?ChannelID=" & ChannelID& "&Action=Add';} else{location.href='KS.Field.asp?ChannelID=" & ChannelID&"&Page='"&Page &";$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=ģ�͹��� >> <font color=#ff0000>ģ���ֶι���</font>&ButtonSymbol=Disabled';}</script>")
		 Else
		   Call KS.AlertHistory("���ݿ����Ѵ��ڸ��ֶ�����!", -1)
		   Exit Sub
		 End If
		 FieldRS.Close
		 End Sub
		 
		 Sub FieldEditSave()
		 With Response
		 ID = KS.G("ID")
		 FieldName = Trim(KS.G("FieldName"))
		 Title = KS.G("Title")
		 Tips = Request.Form("Tips")
		 DefaultValue = KS.G("DefaultValue")
		 MustFillTF = KS.G("MustFillTF")
		 FieldType = KS.G("FieldType")
		 ShowOnForm = KS.G("ShowOnForm")
		 ShowOnUserForm = KS.G("ShowOnUserForm")
		 Options = KS.G("Options")
		 FieldType =KS.G("FieldType")
		 OrderID   =KS.G("OrderID")
		 EditorType=KS.G("EditorType")
		 ShowUnit  =KS.ChkClng(KS.G("ShowUnit"))
		 UnitOptions=KS.G("UnitOptions")
		 ParentFieldName=KS.G("ParentFieldName")
		 If KS.IsNul(ParentFieldName) Then ParentFieldName="0"
		 
		 If Title="" Then Call KS.AlertHistory("�ֶα����������!", -1): Exit Sub
		' If FieldType=4 And Not Isnumeric(DefaultValue) And DefaultValue<>"" Then Call KS.AlertHistory("Ĭ��ֵ��ʽ����ȷ!", -1): Exit Sub
		
		 If FieldType=5 And Not IsDate(DefaultValue) And DefaultValue<>"" and lcase(DefaultValue)<>"now" and lcase(DefaultValue)<>"date" Then Call KS.AlertHistory("Ĭ��ֵ��ʽ����ȷ!", -1): Exit Sub
		
		
		 If Not IsNumeric(OrderID) Then OrderID=0

		 Set FieldRS = Server.CreateObject("ADODB.RECORDSET")
		  FieldSql = "Select * From [KS_Field] Where FieldID=" & ID 
		  FieldRS.Open FieldSql, conn, 1, 3
		  FieldRS("ChannelID") = KS.G("ChannelID")
		  FieldRS("Title") = Title
		  FieldRS("Tips") = Tips
		  FieldRS("DefaultValue") = DefaultValue
		  FieldRS("MustFillTF") = MustFillTF
		  If FieldRS("FieldType")=4 And Not Isnumeric(DefaultValue) And DefaultValue<>"" Then Call KS.AlertHistory("Ĭ��ֵ��ʽ����ȷ!", -1): Exit Sub
		'  If FieldRS("FieldType")=5 And Not IsDate(DefaultValue) And DefaultValue<>"" Then Call KS.AlertHistory("Ĭ��ֵ��ʽ����ȷ!", -1): Exit Sub

		 ' FieldRS("FieldType") = FieldType
		 ' FieldRS("ShowUnit")=ShowUnit
		  'FieldRS("ParentFieldName")=ParentFieldName
		  FieldRS("UnitOptions")=UnitOptions
		  FieldRS("ShowOnForm") = ShowOnForm
		  FieldRS("ShowOnUserForm")=ShowOnUserForm
		  FieldRS("Options") = Options
		  FieldRS("OrderID")=OrderID
		  FieldRS("AllowFileExt")=KS.G("AllowFileExt")
		  FieldRS("MaxFileSize")=KS.ChkClng(KS.G("MaxFileSize"))
		  FieldRS("Width")=KS.ChkClng(KS.G("Width"))
		  FieldRS("Height")=KS.ChkClng(KS.G("Height"))
		  FieldRS("EditorType")=EditorType
		  FieldRS.Update
		  FieldRS.Close
		  on error resume next
	   	  If KS.C_S(KS.G("ChannelID"),6)=1 or KS.C_S(KS.G("ChannelID"),6)=2 or KS.C_S(KS.G("ChannelID"),6)=5 Then
			KS.ConnItem.Execute("Update KS_FieldItem Set FieldTitle='" & Title & "',OrderID=" & OrderID &" Where FieldID=" & ID)
          End If
		 .Write ("<form name=""split"" action=""KS.Split.asp"" method=""GET"" target=""BottomFrame"">")
		 .Write ("<input type=""hidden"" name=""OpStr"" value=""ģ�͹��� >> <font color=red>ģ���ֶι���</font>"">")
		 .Write ("<input type=""hidden"" name=""ButtonSymbol"" value=""Disabled""></form>")
		 .Write ("<script language=""JavaScript"">document.split.submit();</script>")
		 Call KS.Alert("�ֶ��޸ĳɹ�!", "KS.Field.asp?ChannelID=" & ChannelID&"&Page=" & Page)
		 End With
		 End Sub
		 
		 Sub FieldDel()
		    on error resume next
			Dim ID:ID = KS.G("ID")
			Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
			RSObj.Open "Select FieldName,FieldType,ShowUnit From KS_Field Where FieldID IN(" & ID & ")",Conn,1,3
			Do While Not RSObj.Eof 
			  If left(Lcase(RSObj(0)),3)<>"ks_" Then
			   RSObj.Close:Set RSObj=Nothing
			   Response.Write "<script>alert('�Բ���ϵͳ�ֶβ���ɾ��!');history.back(-1);</script>"
			   Response.End()
			  Else
			      Conn.Execute("Alter Table "& TableName &" Drop column "& RSObj(0) &"")
				  If RSObj("ShowUnit")="1" Then
			      Conn.Execute("Alter Table "& TableName &" Drop column "& RSObj(0) &"_Unit")
				  End if
			   	  If KS.C_S(KS.G("ChannelID"),6)=1 or KS.C_S(KS.G("ChannelID"),6)=2 or KS.C_S(KS.G("ChannelID"),6)=5  Then
					  KS.ConnItem.Execute("Delete From KS_FieldItem Where FieldID IN(" & ID & ")")
					  KS.ConnItem.Execute("Delete From KS_FieldRules Where FieldID IN(" & ID & ")")
					  KS.ConnItem.Execute("Alter Table "& TableName &" Drop column "& RSObj(0) &"")
					  If RSObj("ShowUnit")="1" Then
					  KS.ConnItem.Execute("Alter Table "& TableName &" Drop column "& RSObj(0) &"_Unit")
					  End if
				  End If

			   RSObj.Delete
			  End If
			  RSObj.MoveNext
			Loop
			RSObj.Close:Set RSObj=Nothing
			Response.Redirect "KS.Field.asp?ChannelID=" & ChannelID &"&Page=" & Page
		 End Sub
		 
		 Sub FieldOrder()
			  Dim FieldID:FieldID=KS.G("FieldID")
			  Dim OrderID:OrderID=KS.G("OrderID")
			  Dim I,FieldIDArr,OrderIDArr
			  FieldIDArr=Split(FieldID,",")
			  OrderIDArr=Split(OrderID,",")
			  For I=0 To Ubound(FieldIDArr)
			   Conn.Execute("update KS_Field Set OrderID=" & OrderIDArr(i) &" where FieldID=" & FieldIDArr(I))
			   on error resume next
			   If KS.C_S(ChannelID,6)=1 Then
				KS.ConnItem.Execute("Update KS_FieldItem Set OrderID=" & OrderIDArr(i) &" Where FieldID=" & FieldIDArr(I))
			   End If
			  Next
			  Response.Write "<script>alert('���������ֶ�����ɹ���');location.href='?ChannelID=" & ChannelID & "&Page=" & Page&"';</script>"
		 End Sub
End Class
%> 
