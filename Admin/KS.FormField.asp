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
        Private KS,KSCls,Action,ItemID,Page,ItemName,TableName
		Private I, totalPut, CurrentPage, FieldSql, RS,MaxPerPage
		Private FieldName,ID,Contact, Title, Tips, FieldType, DefaultValue, MustFillTF, ShowOnForm, ShowOnUserForm,Options,OrderID,FolderID,MaxFileSize,Width,Height,AllowFileExt,Step

		Private Sub Class_Initialize()
		  MaxPerPage = 18
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub


		Public Sub Kesion()
		With Response
		   If Not KS.ReturnPowerResult(0, "KSMS10006") Then          '���Ȩ��
					 Call KS.ReturnErr(1, "")
					 .End
		   End If
		.Write "<html>"
		.Write "<head>"
		.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
		.Write "<title>�ֶι���</title>"
		.Write "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		.Write "<script language='JavaScript' src='../KS_Inc/common.js'></script>"
		.Write "<script language='JavaScript' src='../KS_Inc/jquery.js'></script>"
             Action=KS.G("Action")
			 ItemID=KS.ChkClng(KS.G("ItemID"))
			 Page=KS.G("Page")
			 Select Case Action
			  Case "Add"  Call FieldManage("Add")
			  Case "Edit" Call FieldManage("Edit")
			  Case "Del"  Call FieldDel()
			  Case "order" Call FieldOrder()
			  Case "AddSave" Call DoSave()
			  Case "EditSave"  Call FieldEditSave()
			  Case Else Call FieldList()
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
		.Write "var ItemID=" & ItemID & ";"
		.Write "</script>"
		.Write "<script language='JavaScript' src='Include/ContextMenu1.js'></script>"
		.Write "<script language='JavaScript' src='Include/SelectElement.js'></script>"
		%>
		 <script language="javascript">
		 var DocElementArrInitialFlag=false;
		var DocElementArr = new Array();
		var DocMenuArr=new Array();
		var SelectedFile='',SelectedFolder='';
		function document.onreadystatechange()
		{   if (DocElementArrInitialFlag) return;
			InitialDocElementArr('FolderID','FieldID');
			InitialContextMenu();
			DocElementArrInitialFlag=true;
		}
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
		   location.href='KS.FormField.asp?ItemID='+ItemID+'&Action=Add';
		   window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=�Զ���� >> ������� >> <font color=red>����'+ItemName+'����</font>&ButtonSymbol=Go';
		}
		function EditField(id)
		{
		  location="KS.FormField.asp?ItemID="+ItemID+"&Page="+Page+"&Action=Edit&ID="+id;
		  window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=�Զ���� >> ������� >> <font color=red>�༭'+ItemName+'����</font>&ButtonSymbol=GoSave';
		}
		function DelField(id)
		{
		if (confirm('���Ҫɾ���ñ�����?'))
		 location="KS.FormField.asp?ItemID="+ItemID+"&Action=Del&Page="+Page+"&id="+id;
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
				  else alert('һ��ֻ�ܱ༭һ������!')	
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
			 alert('��ѡ��Ҫ'+alertmsg+'�ı���');
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
		 if (event.keyCode==46)FieldControl(2);
		}
		 </script>
		<%
		.Write "</head>"
		.Write "<body scroll=no topmargin='0' leftmargin='0' onclick='SelectElement();' onkeydown='GetKeyDown();' onselectstart='return false;'>"
		.Write "<ul id='menu_top'>"
		.Write "<li class='parent' onclick='FieldAdd();'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>��������</span></li>"
		.Write "<li class='parent' onclick='FieldControl(1);'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>�޸ı���</span></li>"
		.Write "<li class='parent' onclick='FieldControl(2)'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>ɾ������</span></li>"
		.Write "<li class='parent' onclick='location.href=""KS.Form.asp"";'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/back.gif' border='0' align='absmiddle'>����һ��</span></li>"
		.Write "</ul>"
        
		.Write ("<div style=""height:94%; overflow: auto; width:100%"" align=""center"">")
		.Write "<form action='KS.FormField.asp?action=order&ItemID=" & ItemID & "&page=" & Page & "' name='form1' method='post'>"
		.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
		.Write "        <tr class='sort'>"
		.Write "         <td width='80' align='center'>����</td>"
		.Write "         <td align='center'>��������</td>"	
		.Write "         <td align='center'>��������</td>"
		.Write "         <td align='center'>Ĭ��ֵ</td>"
		.Write "         <td align='center'>�Ƿ���ʾ</td>"
		.Write "         <td align='center'>���������</td>"
		.Write "        </tr>"
			 Set RS = Server.CreateObject("ADODB.RecordSet")
				   FieldSql = "SELECT * FROM KS_FormField Where ItemID=" & ItemID & " order by OrderID desc"
				   RS.Open FieldSql, conn, 1, 1
				 If RS.EOF And RS.BOF Then
				 Else
					totalPut = RS.RecordCount
		
							If CurrentPage < 1 Then	CurrentPage = 1
		
							If (CurrentPage - 1) * MaxPerPage > totalPut Then
								If (totalPut Mod MaxPerPage) = 0 Then
									CurrentPage = totalPut \ MaxPerPage
								Else
									CurrentPage = totalPut \ MaxPerPage + 1
								End If
							End If
		
							If CurrentPage <> 1 Then
								If (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
								Else
									CurrentPage = 1
								End If
							End If
							Call showContent

			End If
		 .Write " <tr>"
		 .Write "   <td colspan='3'>&nbsp;&nbsp;<input type='button' value='��Ԥ��' onclick=""SelectObjItem1(this,'�Զ���� >> <font color=red>��Ԥ��</font>','gosave','KS.Form.asp?ItemID=" & ItemID & "&action=view');"" class='button'>&nbsp;<input type='submit' class='button' value='������������'> <font color=blue>ԽС����Խǰ��</font></td></form>"
		 .Write "   <td height='35' colspan='4' align='right'>"
		 Call KSCLS.ShowPage(totalPut, MaxPerPage, "KS.FormField.asp", True, "��",CurrentPage, "ItemID=" & ItemID)
		.Write "    </td>"
		.Write " </tr>"
		.Write "</table>"
		.Write "</div>"
		End With
		End Sub
		Sub showContent()
		With Response
		Do While Not RS.EOF
		.Write "  <tr>"
		 .Write "<td class='splittd'>&nbsp;&nbsp;<input type='text' name='OrderID' style='width:45px;text-align:center' value='" & RS("OrderID") &"'><input type='hidden' name='FieldID' value='" & RS("FieldID") & "'></td>"
		.Write "    <td class='splittd'>"
		.Write "    <span FieldID='" & RS("FieldID") & "' onDblClick=""EditField(this.FieldID)"">"
		 .Write "     <img src='Images/Field.gif' align='absmiddle'>"
		 .Write "     <span style='cursor:default;'>" & RS("Title") & "</span>"
		 .Write "   </span>"
		 .Write "   </td>"
		 .Write "   <td align='center' class='splittd'>"
		 Select Case RS("FieldType")
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
		 End Select
		 .Write "</td>"
		 .Write "   <td align='center' class='splittd'>" & RS("DefaultValue") & "&nbsp;</td>"
		 .Write "   <td align='center' class='splittd'>" 
		  If RS("ShowOnForm")=1 Then
		   .Write "<font color=red>��</font>"
		  Else
		   .Write "��"
		  End If
		 .Write " </td>"
		 .Write " <td align='center' class='splittd'><a href='javascript:EditField(" & RS("FieldID") &");'>�޸�</a> | "
		 .Write "<a href='javascript:DelField(" & RS("FieldID") &");'>ɾ��</a>"
		 .Write " </td></tr>"
			I = I + 1
			If I >= MaxPerPage Then Exit Do
			RS.MoveNext
		   Loop
		  RS.Close
         End With
		 End Sub
		 
		 Sub FieldManage(OpType)
		 With Response
		  Dim RS, FieldSql,OpAction,OpTempStr,FormName,PostByStep,StepNum,Step,K
		 ID = KS.G("ID")
		 Set RS=Server.CreateObject("ADODB.Recordset")
		 RS.Open "Select FormName,PostByStep,StepNum From KS_Form Where ID=" & ItemID,conn,1,1
		 If RS.EOF And RS.Bof Then
		  Response.Write "<script>alert('error!');history.back();</script>"
		  Exit Sub
		 Else
		   FormName=RS(0):PostByStep=RS(1):StepNum=RS(2)
		 End If
		 RS.Close
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
			 FieldSql = "Select * From [KS_FormField] Where FieldID=" & ID
			 RS.Open FieldSql, conn, 1, 1
			 If Not RS.EOF Then
				 ItemID    = RS("ItemID")
				 Title     = Trim(RS("Title"))
				 FieldName = RS("FieldName")
				 Tips      = Trim(RS("Tips"))
				 FieldType = Trim(RS("FieldType"))
				 DefaultValue = Trim(RS("DefaultValue"))
				 MustFillTF   = RS("MustFillTF")
				 ShowOnForm   = RS("ShowOnForm")
				 Options      = Trim(RS("Options"))
				 OrderID      = RS("OrderID")
				 Width        = RS("Width")
				 Height       = RS("Height")
				 AllowFileExt = RS("AllowFileExt")
				 MaxFileSize  = RS("MaxFileSize")
				 Step         = RS("Step")
			 End If
	  Else
	     FieldName="KS_":FieldType=1:MustFillTF=0:ShowOnForm=1:ShowOnUserForm=1:OrderID=1:Width="200":Height="100":AllowFileExt="jpg|gif|doc":MaxFileSize=1024
		 OpAction="AddSave":OpTempStr="���"
	  End If
		 
		.Write "<div class='topdashed sort'>" & OpTempStr &"�Զ������</div>"
		.Write "<br>"
        .Write " <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1""  class='ctable'>" & vbCrLf
		.Write "  <form  action='KS.FormField.asp?Action=" & OpAction &"' method='post' name='OrigArticlerm' onsubmit='return(CheckForm())'>"
		
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' width='180' align='right' class='clefttitle'><strong>����Ŀ��</strong></td>"
		.Write "      <td nowrap> &nbsp;&nbsp;<font color=#ff0000>" & FormName & "</font><input type='hidden' value='" & ItemID & "' name='ItemID'></td>"
		.Write "    </tr>"


		
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='45' align='right' class='clefttitle'><strong>���������</strong></td>"
		.Write "      <td height='45' nowrap>&nbsp;&nbsp;<input name='Title' type='text' size='30' class='textbox' value='" & Title & "'> *<font color=red>�磬���������</font></td>"
		.Write "    </tr>"
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='45' align='right' class='clefttitle'><strong>�ֶ����ƣ�</strong></td>"
		If Optype = "Edit" Then
		.Write "      <td height='45' nowrap>&nbsp;&nbsp;<input disabled name='FieldName' type='text'size='30' class='textbox' value='" & FieldName & "'> <font color=red>*������KS_��ͷ���ֶ�������ĸ�����֡��»������,�Ҳ����޸�</font> </td>"
		Else
		.Write "      <td height='45' nowrap>&nbsp;&nbsp;<input name='FieldName' type='text'size='30' class='textbox' value='" & FieldName & "'> <font color=red>*������KS_��ͷ���ֶ�������ĸ�����֡��»������,�Ҳ����޸�</font> </td>"
		End If
		.Write "    </tr>"
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right'' class='clefttitle'><strong>������ʾ��</strong><br><font color=blue>�������Ե���ʾ��Ϣ</font></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<textarea name='Tips'  id='Tips' class='textbox' cols='30' rows='3' style='height:50px'>" & Tips & "</textarea></td>"
		.Write "    </tr>"
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>�ֶ����ͣ�</strong></td>"
		.Write "      <td nowrap>&nbsp;"
		If Optype = "Edit" Then
		.Write "      <input type='hidden' value=" & FieldType & " name='FieldType'><select name=""FieldTypes"" disabled>"
		else
		.Write "      <select name=""FieldType"" onchange=""Setdisplay(this.value)"">"
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
		.Write " </td>"
		.Write "    </tr>"
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
		.Write "       </td>"
		.Write "    </tr>"
		.Write "    <tr id=""OptionsArea"" style=""display:none"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td align='right' ' class='clefttitle'><strong>�б�ѡ�</strong><br><font color=blue>ÿһ��Ϊһ���б�ѡ��</font></td>"
		.Write "      <td height='45' nowrap>&nbsp;&nbsp;<textarea name='Options' cols='50' rows='6' id='Options' style='height:60px' class='textbox'>" & Options & "</textarea>"
		.Write "      </td>"
		.Write "    </tr>"
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
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>��С���ã�</strong> </td>"
		.Write "      <td nowrap>&nbsp;&nbsp;���<input name='Width' type='text' site='10' class='textbox' style='width:40px' value='" & Width & "'>px &nbsp;�߶�<input name='Height' type='text' site='10' class='textbox' style='width:40px' value='" & Height & "'>px</font>"
		.Write "       </td>"
		.Write "    </tr>"				
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>������ţ�</strong><br><font color=blue>���ԽС������Խǰ��</font></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;<input name='OrderID' type='text' site='35' class='textbox' id='OrderID' value='" & OrderID & "'>"
		.Write "       </td>"
		.Write "    </tr>"
		If PostByStep="1" and StepNum>1 Then
		.Write "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
		.Write "      <td height='30' align='right' ' class='clefttitle'><strong>�ֲ��ύ���ã�</strong></td>"
		.Write "      <td nowrap>&nbsp;&nbsp;�ñ����ڵ�<select name='step'>"
		  For K=1 To StepNum
		   If K=Step Then
		   .Write "<option selected>" & K & "</option>"
		   Else
		   .Write "<option>" & K & "</option>"
		   End If
		  Next
		.Write "       </select>������</td>"
		.Write "    </tr>"
		End If
		
		.Write "   <input type='hidden' value='" & ID & "' name='id'>"
		.Write "   <input type='hidden' value='" & Page & "' name='page'>"
		.Write "  </form>"
		.Write "</table>"
		
		
		 
		.Write "<Script Language='javascript'>"
		.Write "Setdisplay(" & FieldType & ");"
		.Write "function Setdisplay(s)"
		.Write  "{if (s==3||s==6||s==7){ document.all.OptionsArea.style.display='';} else document.all.OptionsArea.style.display='none';if (s==7)document.getElementById('darea').style.display='';else document.getElementById('darea').style.display='none';if(s==9)document.getElementById('extarea').style.display='';else document.getElementById('extarea').style.display='none'; }"
		.Write "function CheckForm()"
		.Write "{ var form=document.OrigArticlerm;"
		.Write "   if (form.Title.value=='')"
		.Write "    {"
		.Write "     alert('�������������!');"
		.Write "     form.Title.focus();"
		.Write "     return false;"
		.Write "    }"
		.Write "    form.submit();"
		.Write "    return true;"
		.Write "}"
		.Write "</Script>"
		End With
		End Sub
		 
		 Sub DoSave()
		 Dim RS,ColumnType,ItemID
		 ItemID = KS.ChkClng(KS.G("ItemID"))
		 Title = KS.G("Title")
		 FieldName = KS.G("FieldName")
		 Tips = Request.Form("Tips")
		 FieldType = KS.G("FieldType")
		 DefaultValue = KS.G("DefaultValue")
		 MustFillTF = KS.G("MustFillTF")
		 FieldType = KS.G("FieldType")
		 ShowOnForm = KS.G("ShowOnForm")
		 Options = KS.G("Options")
		 FieldType =KS.G("FieldType")
		 OrderID   =KS.G("OrderID")
		 Width     = KS.ChkClng(KS.G("Width"))
		 Height    = KS.ChkClng(KS.G("Height"))
		 AllowFileExt = KS.G("AllowFileExt")
		 MaxFileSize  = KS.ChkClng(KS.G("MaxFileSize"))
		 Step         = KS.ChkClng(KS.G("Step"))

		 If FieldName = "" Then Call KS.AlertHistory("�������ֶ�����!", -1): Exit Sub
		 If Len(FieldName)<=3 Then Call KS.AlertHistory("�ֶ����Ƴ��ȱ������3!", -1): Exit Sub
		 If Ucase(Left(FieldName,3))<>"KS_" Then Call KS.AlertHistory("�ֶ����Ƹ�ʽ���󣬱�����""KS_��ͷ""!", -1): Exit Sub
		 If Title="" Then Call KS.AlertHistory("�ֶα����������!", -1): Exit Sub
		 If FieldType=4 And Not Isnumeric(DefaultValue) And DefaultValue<>"" Then Call KS.AlertHistory("Ĭ��ֵ��ʽ����ȷ!", -1): Exit Sub
		 If FieldType=5 And Not IsDate(DefaultValue) And DefaultValue<>"" Then Call KS.AlertHistory("Ĭ��ֵ��ʽ����ȷ!", -1): Exit Sub
		 if FieldType=8 And Not KS.IsValidEmail(DefaultValue) And DefaultValue<>"" Then Call KS.AlertHistory("Ĭ�ϸ�ʽ����ȷ����������ȷ��Email!",-1):Exit Sub
	     on error resume next
		 Conn.Begintrans
		 Set RS = Server.CreateObject("ADODB.RECORDSET")
		 FieldSql = "Select * From [KS_FormField] Where FieldName='" & FieldName & "' And ItemID=" & ItemID
		 RS.Open FieldSql, conn, 3, 3
		 If RS.EOF And RS.BOF Then
		  RS.AddNew
		  RS("ItemID") = KS.G("ItemID")
		  RS("Title") = Title
		  RS("FieldName") = FieldName
		  RS("Tips") = Tips
		  RS("FieldType") = FieldType
		  RS("DefaultValue") = DefaultValue
		  RS("MustFillTF") = MustFillTF
		  RS("FieldType") = FieldType
		  RS("ShowOnForm") = ShowOnForm
		  RS("Options") = Options
		  RS("OrderID")=OrderID
		  RS("Width")  = Width
		  RS("Height") = Height
		  RS("AllowFileExt")= AllowFileExt
		  RS("MaxFileSize") = MaxFileSize
		  RS("Step") = Step
		  RS.Update
		  
		  Select Case FieldType
		   Case 1,3,6,7,8,9
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
		 Dim TableName:TableName=Conn.Execute("Select TableName From KS_Form  Where ID=" & ItemID)(0)
		 Conn.Execute("Alter Table "&TableName&" Add "&FieldName&" "&ColumnType&"")
		 If err<>0 then
			Conn.RollBackTrans
			Call KS.AlertHistory("��������������" & replace(err.description,"'","\'"),-1):response.end
		 Else
			Conn.CommitTrans
		 End IF
		 Response.Write ("<Script> if (confirm('�������ӳɹ�,���������?')) { location.href='KS.FormField.asp?ItemID=" & ItemID& "&Action=Add';} else{location.href='KS.FormField.asp?ItemID=" & ItemID&"&Page='"&Page &";$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=�Զ�������� >> <font color=#ff0000>�������</font>&ButtonSymbol=Disabled';}</script>")
		 Else
		   Call KS.AlertHistory("���ݿ����Ѵ��ڸ��ֶ�����!", -1)
		   Exit Sub
		 End If
		 RS.Close
		 End Sub
		 
		 Sub FieldEditSave()
		 With Response
		 ID = KS.G("ID")
		 Title = KS.G("Title")
		 Tips = Request.Form("Tips")
		 DefaultValue = KS.G("DefaultValue")
		 MustFillTF = KS.G("MustFillTF")
		 FieldType = KS.G("FieldType")
		 ShowOnForm = KS.G("ShowOnForm")
		 Options = KS.G("Options")
		 FieldType =KS.G("FieldType")
		 OrderID   =KS.G("OrderID")
		 Width     = KS.ChkClng(KS.G("Width"))
		 Height    = KS.ChkClng(KS.G("Height"))
		 AllowFileExt = KS.G("AllowFileExt")
		 MaxFileSize  = KS.ChkClng(KS.G("MaxFileSize"))
		 Step         = KS.ChkClng(KS.G("Step"))

		 If Title="" Then Call KS.AlertHistory("�������Ʊ�������!", -1): Exit Sub
		 If FieldType=4 And Not Isnumeric(DefaultValue) And DefaultValue<>"" Then Call KS.AlertHistory("Ĭ��ֵ��ʽ����ȷ!", -1): Exit Sub
		 If FieldType=5 And Not IsDate(DefaultValue) And DefaultValue<>"" Then Call KS.AlertHistory("Ĭ��ֵ��ʽ����ȷ!", -1): Exit Sub
		 if FieldType=8 And Not KS.IsValidEmail(DefaultValue) And DefaultValue<>"" Then Call KS.AlertHistory("Ĭ�ϸ�ʽ����ȷ����������ȷ��Email!",-1):Exit Sub

		 If Not IsNumeric(OrderID) Then OrderID=0

		 Set RS = Server.CreateObject("ADODB.RECORDSET")
		  FieldSql = "Select * From [KS_FormField] Where FieldID=" & ID 
		  RS.Open FieldSql, conn, 1, 3
		  RS("ItemID") = KS.G("ItemID")
		  RS("Title") = Title
		  RS("Tips") = Tips
		  RS("DefaultValue") = DefaultValue
		  RS("MustFillTF") = MustFillTF
		  RS("ShowOnForm") = ShowOnForm
		  RS("Options") = Options
		  RS("OrderID") = OrderID
		  RS("Width")   = Width
		  RS("Height")  = Height
		  RS("AllowFileExt")= AllowFileExt
		  RS("MaxFileSize") = MaxFileSize
		  RS("Step") = Step
		  RS.Update
		  RS.Close
		 .Write ("<form name=""split"" action=""KS.Split.asp"" method=""GET"" target=""BottomFrame"">")
		 .Write ("<input type=""hidden"" name=""OpStr"" value=""�Զ�������� >> <font color=red>�Զ���������</font>"">")
		 .Write ("<input type=""hidden"" name=""ButtonSymbol"" value=""Disabled""></form>")
		 .Write ("<script language=""JavaScript"">document.split.submit();</script>")
		 Call KS.Alert("�����޸ĳɹ�!", "KS.FormField.asp?ItemID=" & ItemID&"&Page=" & Page)
		 End With
		 End Sub
		 
		 Sub FieldDel()
			Dim TableName:TableName=Conn.Execute("Select TableName From KS_Form  Where ID=" & ItemID)(0)
			Dim ID:ID = KS.G("ID")
			Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
			RSObj.Open "Select FieldName,FieldType From KS_FormField Where FieldID IN(" & ID & ")",Conn,1,3
			Do While Not RSObj.Eof 
			  If left(Lcase(RSObj(0)),3)<>"ks_" Then
			   RSObj.Close:Set RSObj=Nothing
			   Response.Write "<script>alert('�Բ���ϵͳ�ֶβ���ɾ��!');history.back(-1);</script>"
			   Response.End()
			  Else
			   Conn.Execute("Alter Table "& TableName &" Drop column "& RSObj(0) &"")
			   RSObj.Delete
			  End If
			  RSObj.MoveNext
			Loop
			RSObj.Close:Set RSObj=Nothing
			Response.Redirect "KS.FormField.asp?ItemID=" & ItemID &"&Page=" & Page
		 End Sub
		 
		 Sub FieldOrder()
			  Dim FieldID:FieldID=KS.G("FieldID")
			  Dim OrderID:OrderID=KS.G("OrderID")
			  Dim I,FieldIDArr,OrderIDArr
			  FieldIDArr=Split(FieldID,",")
			  OrderIDArr=Split(OrderID,",")
			  For I=0 To Ubound(FieldIDArr)
			   Conn.Execute("update KS_FormField Set OrderID=" & OrderIDArr(i) &" where FieldID=" & FieldIDArr(I))
			  Next
			  Response.Write "<script>alert('���������ֶ�����ɹ���');location.href='KS.FormField.asp?ItemID=" & ItemID&"&Page=" & Page & "';</script>"
		 End Sub
End Class
%> 
