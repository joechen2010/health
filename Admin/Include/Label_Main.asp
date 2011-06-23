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
Set KSCls = New Label_Main
KSCls.Kesion()
Set KSCls = Nothing

Class Label_Main
        Private KS
		Private LabelSql, LabelRS, FolderID, LabelID, ChannelID, Channel, Action
		Private i, totalPut, CurrentPage, LabelType,UPFolderRS, ParentID
		Private KeyWord, SearchType, StartDate, EndDate
		'������������
		Private SearchParam
		Private MaxPerPage
		Private Row 
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
			FolderID = KS.G("FolderID"):If FolderID = "" Then FolderID = "0"
			LabelType = KS.G("LabelType"):If LabelType = "" Then LabelType = 0
			If LabelType = 0 Then
				If Not KS.ReturnPowerResult(0, "KMTL10001") Then                'ϵͳ������ǩ��Ȩ�޼��
				  Call KS.ReturnErr(1, "")
				  Response.End
				End If
			ElseIf LabelType = 1 Then
				If Not KS.ReturnPowerResult(0, "KMTL10003") Then                '�Զ��徲̬��ǩ��Ȩ�޼��
				  Call KS.ReturnErr(1, "")
				  Response.End
				End If
			ElseIf LabelType = 5 Then
				If Not KS.ReturnPowerResult(0, "KMTL10002") Then                '�Զ��庯����ǩ��Ȩ�޼��
				  Call KS.ReturnErr(1, "")
				  Response.End
				End If
			ElseIf LabelType = 6 Then
				If Not KS.ReturnPowerResult(0, "KMTL10010") Then                '�Զ��庯����ǩ��Ȩ�޼��
				  Call KS.ReturnErr(1, "")
				  Response.End
				End If
			End If
			
		If Not IsEmpty(KS.G("page")) And KS.G("page") <> "" Then
			  CurrentPage = CInt(KS.G("page"))
		Else
			  CurrentPage = 1
		End If
		Set UPFolderRS = Conn.Execute("select * from KS_LabelFolder where ID ='" & FolderID & "'")
		If Not UPFolderRS.EOF Then
		 ParentID = UPFolderRS("ParentID")
		End If
		UPFolderRS.Close:Set UPFolderRS = Nothing
		
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		Response.Write "<title>��ǩ�б�</title>"
		Response.Write "</head>"
		Response.Write "<link href=""Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"

			Action = KS.G("Action")
			Select Case  Action
			 Case "SetPasteParam"
			   Call SetPasteParam()
			 Case "PasteSave"
			   Call LabelPasteSave()
			 Case "LabelDel"
			   Call LabelDel()
			 Case "LabelFolderDel"
			   Call LabelFolderDel()
			 Case "LabelOut"
			   Call LabelOut()
			 Case "Doexport"
			   Call Doexport()
			 Case "LabelIn"
			   Call LabelIn()
			 Case "LabelIn2"
			   Call LabelIn2()
			 Case "Doimport"
			   Call Doimport()
			 Case Else
			  Call LabelMainList()
			End Select
        End Sub 
		
		Sub LabelMainList()	
		Response.Write "<script language=""JavaScript"">" & vbCrLf
		Response.Write "var FolderID='" & FolderID & "';         //Ŀ¼ID" & vbCrLf
		Response.Write "var ParentID='" & ParentID & "'; //����ĿID" & vbCrLf
		Response.Write "var Page='" & CurrentPage & "';   //��ǰҳ��" & vbCrLf
		Response.Write "var KeyWord='" & KeyWord & "';    //�ؼ���" & vbCrLf
		Response.Write "var SearchParam='" & SearchParam & "';  //������������" & vbCrLf
		Response.Write "var LabelType=" & LabelType & ";" & vbCrLf
		Response.Write "</script>" & vbCrLf
		Response.Write "<script language=""JavaScript"" src=""../../ks_inc/Common.js""></script>"
		Response.Write "<script language=""JavaScript"" src=""../../ks_inc/jQuery.js""></script>"
		Response.Write "<script language=""JavaScript"" src=""../../ks_inc/Kesion.Box.js""></script>"
		Response.Write "<script language=""JavaScript"" src=""ContextMenu.js""></script>"
		Response.Write "<script language=""JavaScript"" src=""SelectElement.js""></script>"
		%>
		<script language="javascript">
		var DocElementArrInitialFlag=false;
		var DocElementArr = new Array();
		var DocMenuArr=new Array();
		var SelectedFile='',SelectedFolder='';
		$(document).ready(function(){
			if (DocElementArrInitialFlag) return;
			InitialDocElementArr('FolderID','LabelID');
			InitialDocMenuArr();
			DocElementArrInitialFlag=true;
		})
		function InitialDocMenuArr()
		{  
		   if (KeyWord=='')
			{ 
			 DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.AddFolder();",'�½�Ŀ¼(N)','disabled');
			 DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.AddLabel('');",'�½���ǩ(M)','disabled');
			 DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
			}
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.SelectAllElement();",'ȫ ѡ(A)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.Edit('');",'�� ��(E)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.Delete('');",'ɾ ��(D)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.Paste();",'�� ¡(V)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem('parent.ChangeUp();','�� ��(B)','');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.Reload('');",'ˢ ��(Z) ','');
		}
		function DocDisabledContextMenu()
		{ 
		   var TempDisabledStr=''; 
		   if (FolderID=='0') TempDisabledStr='�� ��(B),';
			DisabledContextMenu('FolderID','LabelID',TempDisabledStr+'�� ��(E),ɾ ��(D),�� ¡(V)','�� ¡(V)','','�� ��(E),�� ¡(V)','�� ��(E),�� ¡(V)','�� ¡(V)')
		}
		function ChangeUp()
		{
		 if (FolderID=='0') return;
		 location.href='Label_Main.asp?LabelType='+LabelType+'&FolderID='+ParentID;
		 if (LabelType==0)
			$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('��ǩ���� >> ϵͳ������ǩ')+'&ButtonSymbol=FunctionLabel&LabelFolderID='+ParentID;
		 else if(LabelType==5)
		 	$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('��ǩ���� >> �Զ��庯����ǩ')+'&ButtonSymbol=DIYFunctionLabel&LabelFolderID='+ParentID;
		 else
		   $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('��ǩ���� >> �Զ��徲̬��ǩ')+'&ButtonSymbol=FreeLabel&LabelFolderID='+ParentID;
		}
		function OpenLabelFolder(FolderID)
		{
			location.href='Label_Main.asp?LabelType='+LabelType+'&FolderID='+FolderID;
		   if (LabelType==0)
			 $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('��ǩ���� >> ϵͳ������ǩ')+'&ButtonSymbol=FunctionLabel&LabelFolderID='+FolderID;
			else if (LabelType==5)
			 $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('��ǩ���� >> �Զ��庯����ǩ')+'&ButtonSymbol=DIYFunctionLabel&LabelFolderID='+FolderID;
			else
			 $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('��ǩ���� >> �Զ��徲̬��ǩ')+'&ButtonSymbol=FreeLabel&LabelFolderID='+FolderID;
		}
		function AddFolder()
		{
		   if (LabelType==1)
			 OpenWindow('LabelFrame.asp?Url=LabelFolder.asp&PageTitle='+escape('�½��Զ��徲̬��ǩĿ¼')+'&LabelType=1&FolderID='+FolderID,450,300,window);
		   else if(LabelType==5)
			 OpenWindow('LabelFrame.asp?Url=LabelFolder.asp&PageTitle='+escape('�½��Զ��庯����ǩĿ¼')+'&LabelType=5&FolderID='+FolderID,450,300,window);
		   else if(LabelType==6)
			 OpenWindow('LabelFrame.asp?Url=LabelFolder.asp&PageTitle='+escape('�½�ѭ���б��ǩĿ¼')+'&LabelType=6&FolderID='+FolderID,450,300,window);
		   else
			 OpenWindow('LabelFrame.asp?Url=LabelFolder.asp&PageTitle='+escape('�½�ϵͳ������ǩĿ¼')+'&LabelType=0&FolderID='+FolderID,450,300,window);
		 Reload('');
		}
		function AddLabel(TempUrl)
		{ 
		   if (LabelType==1){
				location.href=TempUrl+'LabelAdd.asp?mode=text&LabelType=1&Action=AddNew&FolderID='+FolderID;
				$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('��ǩ���� >> <font color=red>����Զ��徲̬��ǩ</font>')+'&ButtonSymbol=LabelAdd';
				}
		   else if(LabelType==5){
				location.href=TempUrl+'LabelFunctionAdd.asp?LabelType=5&FolderID='+FolderID;
				$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+('��ǩ���� >> <font color=red>����Զ��庯����ǩ</font>')+'&ButtonSymbol=DIYFunctionStep1';		
}
		   else if(LabelType==6){
				location.href=TempUrl+'CirLabel.asp?LabelType=6&FolderID='+FolderID;
				$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('��ǩ���� >> <font color=red>���ѭ����ǩ</font>')+'&ButtonSymbol=LabelAdd';}
		   else
			 { 
			    location.href=TempUrl+'AddFunctionLabel.asp?FolderID='+FolderID;
				$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('��ǩ���� >> <font color=red>���ϵͳ������ǩ</font>')+'&ButtonSymbol=Go';
			    // PopupImgDir="../";
			   // PopupCenterIframe('��Ӻ�����ǩ','AddFunctionLabel.asp?FolderID='+FolderID,700,440,'no')
			  // OpenWindow('LabelFrame.asp?Url=AddFunctionLabel.asp&FolderID='+FolderID+'&PageTitle='+escape('��Ӻ�����ǩ'),450,400,window);
			   //Reload(TempUrl);
			  }
		}
		function EditLabel(TempUrl,id)
		{ 	if (LabelType==1)
				if (KeyWord=='')
				 {	location.href=TempUrl+'LabelAdd.asp?mode=text&LabelType=1&page='+Page+'&Action=EditLabel&LabelID='+id;
					$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('��ǩ���� >> <font color=red>�޸��Զ��徲̬��ǩ</font>')+'&ButtonSymbol=LabelAdd';
				 }
				else
				   { location.href=TempUrl+'LabelAdd.asp?mode=text&LabelType=1&page='+Page+'&Action=EditLabel&'+SearchParam+'&LabelID='+id;
					$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('��ǩ���� >> �����Զ��徲̬��ǩ��� >> <font color=red>�޸��Զ��徲̬��ǩ</font>')+'&ButtonSymbol=LabelAdd';
				   }
			else if(LabelType==5)
			   if (KeyWord=='')
				 {	location.href=TempUrl+'LabelFunctionAdd.asp?LabelType=5&page='+Page+'&Action=Edit&LabelID='+id;
					$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('��ǩ���� >> <font color=red>�޸��Զ��庯����ǩ</font>')+'&ButtonSymbol=DIYFUNCTIONSTEP1';
				 }
				else
				   { location.href=TempUrl+'LabelFunctionAdd.asp?LabelType=5&page='+Page+'&Action=Edit&'+SearchParam+'&LabelID='+id;
					$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('��ǩ���� >> �����Զ��庯����ǩ��� >> <font color=red>�޸��Զ��庯����ǩ</font>')+'&ButtonSymbol=DIYFUNCTIONSTEP1';
					}
		    else if(LabelType==6){
			        location.href=TempUrl+'CirLabel.asp?mode=text&LabelType=1&page='+Page+'&Action=EditLabel&LabelID='+id;
					$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('��ǩ���� >> <font color=red>�޸�ѭ����ǩ</font>')+'&ButtonSymbol=LabelAdd';
			}
			else
			 {	
			 
			 	location.href=TempUrl+'EditFunctionLabel.asp?LabelID='+id;
				$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('��ǩ���� >> <font color=red>�޸�ϵͳ������ǩ</font>')+'&ButtonSymbol=GoSave';

			// OpenWindow('LabelFrame.asp?Url=EditFunctionLabel.asp&PageTitle='+escape('�޸ĺ�����ǩ')+'&LabelID='+id,450,350,window);
				//Reload(TempUrl);
			 }
		}
		function AddByText()
		{
		 location.href='LabelAdd.asp?LabelType=1&Action=AddNew&FolderID='+FolderID;
		 $(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('��ǩ���� >> <font color=red>����Զ��徲̬��ǩ</font>')+'&ButtonSymbol=LabelAdd';
		}
		function EditByText(id)
		{ 	GetSelectStatus('FolderID','LabelID');
			
				   if (SelectedFile!='')
					 {
					 if (SelectedFile.indexOf(',')==-1) 
					  {
					  location.href='LabelAdd.asp?LabelType=1&page='+Page+'&Action=EditLabel&'+SearchParam+'&LabelID='+SelectedFile;
			$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3)&KS.Setting(89)%>KS.Split.asp?OpStr='+escape('��ǩ���� >> �����Զ��徲̬��ǩ��� >> <font color=red>�޸��Զ��徲̬��ǩ</font>')+'&ButtonSymbol=LabelAdd';
					   }
					else alert('һ��ֻ�ܹ��༭һ����ǩ');
					}
				 else
				   alert('��ѡ����Ҫ�༭�ı�ǩ!');
			SelectedFile='';
		}
		function LabelOut()
		{
		location.href='?Action=LabelOut&LabelType='+LabelType;
		}
		function LabelIn()
		{
		location.href='?Action=LabelIn&LabelType='+LabelType;
		}
		function EditFolder(ID)
		{
		 OpenWindow('LabelFrame.asp?Url=LabelFolder.asp&PageTitle='+escape('�༭��ǩĿ¼')+'&Action=EditFolder&FolderID='+ID,450,300,window);
		 Reload('');
		}
		function Edit(TempUrl)
		{   GetSelectStatus('FolderID','LabelID');
			if (!((SelectedFile=='')&&(SelectedFolder=='')))
				{
					if (SelectedFolder!='')
					{ 
					if (TempUrl=='Folder'||TempUrl=='')
					 if (SelectedFolder.indexOf(',')==-1) 
					  { EditFolder(SelectedFolder);
					 }
					else alert('һ��ֻ�ܹ��༭һ����ǩĿ¼');
				   }
				   if (SelectedFile!='')
					 {
					 if (TempUrl!='Folder'||TempUrl=='')
					 {if (SelectedFile.indexOf(',')==-1) 
					   EditLabel(TempUrl,SelectedFile);
					else alert('һ��ֻ�ܹ��༭һ����ǩ');
					 }
					}
				}
			else 
			{
			alert('��ѡ��Ҫ�༭�ı�ǩ��Ŀ¼');
			}
			SelectedFile='';
			SelectedFolder='';
		}
		function Delete(TempUrl)
		{   GetSelectStatus('FolderID','LabelID');
			if (!((SelectedFile=='')&&(SelectedFolder=='')))
				{  
					if (confirm('ɾ��ȷ��:\n\n���Ҫִ��ɾ��������?'))
					  { if (SelectedFolder!='')
					   if (TempUrl=='Folder'||TempUrl=='')
						 location='Label_Main.asp?Action=LabelFolderDel&ID='+SelectedFolder+'&FolderID='+FolderID+'&LabelType='+LabelType;
					  if (SelectedFile!='')  
						if (TempUrl!='Folder'||TempUrl=='')
						location=TempUrl+'Label_Main.asp?Action=LabelDel&Page='+Page+'&ID='+SelectedFile+'&FolderID='+FolderID+'&LabelType='+LabelType;
					}	
				}
			else alert('��ѡ��Ҫɾ���ı�ǩĿ¼���ǩ');
		   SelectedFile='';
		   SelectedFolder='';
		}
		function Paste()
		{
		GetSelectStatus('FolderID','LabelID');
			if (SelectedFile!='')  
			  OpenWindow('LabelFrame.asp?Url=Label_Main.asp&Action=SetPasteParam&PageTitle='+escape('�������±�ǩ����')+'&LabelType=0&LabelID='+SelectedFile,350,120,window);
			else alert('��ѡ��Ҫ��¡�ı�ǩ');
			SelectedFile='';
			SelectedFolder='';
			Reload('');
		}
		function GetKeyDown()
		{
		if (event.ctrlKey)
		  switch  (event.keyCode)
		  {  case 90 :  Reload(''); break;
			 case 65 : SelectAllElement();break;
			 case 66 : event.keyCode=0;event.returnValue=false;ChangeUp();break;
			 case 78 : event.keyCode=0;event.returnValue=false; AddFolder();break;
			 case 77 : event.keyCode=0;event.returnValue=false; AddLabel('');break;
			 case 69 : event.keyCode=0;event.returnValue=false;Edit('');break;
			 case 86 : event.keyCode=0;event.returnValue=false;Paste();break;
			 case 68 : Delete('');break;
			 case 70 :event.keyCode=0;event.returnValue=false;
			 if (LabelType==0)
				parent.frames['LeftFrame'].initializeSearch('SysLabel')
			 else
			   parent.frames['LeftFrame'].initializeSearch('FreeLabel')
		 }	
		else if (event.keyCode==46)
		Delete('');
		}
		function Reload(TempUrl)
		{
		location.href=TempUrl+'Label_Main.asp?FolderID='+FolderID+'&page='+Page+'&LabelType='+LabelType+'&'+SearchParam;
		}
		</script>
		<%
		Response.Write "<body scroll=no topmargin=""0"" leftmargin=""0"" OnClick=""SelectElement();"" onkeydown=""GetKeyDown();"" onselectstart=""return false;"">"
           Response.Write "<ul id='menu_top'>"			 
			 If KeyWord = "" Then
			  Response.Write "<li class='parent' onclick=""AddLabel('')""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/a.gif' border='0' align='absmiddle'>��ӱ�ǩ</span></li>"
			  Response.Write "<li class='parent' onclick=""AddFolder();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/a.gif' border='0' align='absmiddle'>���Ŀ¼</span></li>"
			  Response.Write "<li class='parent' onclick=""Edit('Folder');""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/as.gif' border='0' align='absmiddle'>�༭Ŀ¼</span></li>"
			  Response.Write "<li class='parent' onclick=""Delete('Folder');""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/del.gif' border='0' align='absmiddle'>ɾ��Ŀ¼</span></li>"
			  Response.Write "<li class='parent' onclick=""parent.frames['LeftFrame'].initializeSearch("
			If LabelType = 0 Then Response.Write ("'ϵͳ������ǩ'") Else If LabelType=5 Then Response.Write("'�Զ��庯����ǩ'") Else Response.Write ("'�Զ��徲̬��ǩ'")
			  Response.Write ");""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/s.gif' border='0' align='absmiddle'>��������</span></li>"

			If LabelType=1 Then
			  Response.Write "<li class='parent' onclick=""AddByText();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/a.gif' border='0' align='absmiddle'>���ӻ����</span></li>"
			  Response.Write "<li class='parent' onclick=""EditByText();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/as.gif' border='0' align='absmiddle'>���ӻ��༭</span></li>"
			End If
			  Response.Write "<li class='parent' onclick=""LabelIn();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/move.gif' border='0' align='absmiddle'>��ǩ����</span></li>"
			  Response.Write "<li class='parent' onclick=""LabelOut();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/verify.gif' border='0' align='absmiddle'>��ǩ����</span></li>"
			  Response.Write "<li class='parent' onclick=""ChangeUp();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/verify.gif' border='0' align='absmiddle'>����һ��</span></li>"

			 
			 Else
				If LabelType = 0 Then
				   Response.Write ("<img src='../Images/home.gif' align='absmiddle'><span style='cursor:pointer' onclick=""SendFrameInfo('Label_Main.asp?LabelType=0','../Template_Left.asp','../KS.Split.asp?ButtonSymbol=FunctionLabel&OpStr=��ǩ���� >> <font color=red>ϵͳ������ǩ</font>')"">ϵͳ��ǩ��ҳ</span>")
				ElseIf LabelType=5 Then
				  Response.Write ("<img src='../Images/home.gif' align='absmiddle'><span style='cursor:pointer' onclick=""SendFrameInfo('Label_Main.asp?LabelType=5','../Template_Left.asp','../KS.Split.asp?ButtonSymbol=FunctionLabel&OpStr=��ǩ���� >> <font color=red>�Զ��庯����ǩ</font>')"">�Զ��庯����ǩ��ҳ</span>")
				Else
				   Response.Write ("<img src='../Images/home.gif' align='absmiddle'><span style='cursor:pointer' onclick=""SendFrameInfo('Label_Main.asp?LabelType=1','../Template_Left.asp','../KS.Split.asp?ButtonSymbol=FunctionLabel&OpStr=��ǩ���� >> <font color=red>�Զ��徲̬��ǩ</font>')"">�Զ��徲̬��ǩ��ҳ</span>")
				End If
			   Response.Write (">>> �������: ")
				 If StartDate <> "" And EndDate <> "" Then
					Response.Write ("��ǩ���������� <font color=red>" & StartDate & "</font> �� <font color=red> " & EndDate & "</font>&nbsp;&nbsp;&nbsp;&nbsp;")
				 End If
				Select Case SearchType
				 Case 0
				  Response.Write ("���ƺ��� <font color=red>" & KeyWord & "</font> �ı�ǩ")
				 Case 1
				  Response.Write ("�����к��� <font color=red>" & KeyWord & "</font> �ı�ǩ")
				 Case 2
				  Response.Write ("�����к��� <font color=red>" & KeyWord & "</font> �ı�ǩ")
				 End Select
			 End If
		Response.Write "</ul>"
	
		Response.Write "<div style="" height:98%; overflow: auto; width:100%"" align=""center"">"
		Response.Write "  <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		Response.Write "      <tr>"
		Response.Write "    <td height=""6""></td>"
		Response.Write "      </tr>"
			  
			  Dim FolderSql, Param
			  Param = " Where LabelType=" & LabelType
			If KeyWord <> "" Then
				FolderSql = "SELECT ID,FolderName,Description,OrderID,OrderID FROM [KS_LabelFolder] where 1=0"
				Select Case SearchType
					Case 0
					  Param = Param & " AND LabelName like '%" & KeyWord & "%'"
					Case 1
					 Param = Param & " AND Description like '%" & KeyWord & "%'"
					Case 2
					 Param = Param & " AND LabelContent like '%" & KeyWord & "%'"
				End Select
				If StartDate <> "" And EndDate <> "" Then
					 Param = Param & " And (AddDate>=#" & StartDate & "# And AddDate<=#" & DateAdd("d", 1, EndDate) & "#)"
			   End If
			Else
			  FolderSql = "SELECT ID,FolderName,Description,OrderID as LabelFlag,OrderID FROM [KS_LabelFolder] where  FolderType=" & LabelType & " And ParentID='" & FolderID & "'"
			  Param = Param & " AND FolderID='" & FolderID & "'"
			End If
			Param = Param & " ORDER BY OrderID "
		Set LabelRS = Server.CreateObject("ADODB.recordset")
		LabelRS.Open FolderSql & " UNION all Select ID,LabelName,Description,LabelFlag,OrderID from [KS_Label] " & Param, Conn, 1, 1
		If LabelRS.EOF And LabelRS.BOF Then
				 Else
					totalPut = LabelRS.RecordCount
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
									LabelRS.Move (CurrentPage - 1) * MaxPerPage
									
									Call showContent
								Else
									CurrentPage = 1
									Call showContent
								End If
							End If
			End If
		
		
		Response.Write "    </table>"
		Response.Write "    </div>"
		Response.Write "</body>"
		Response.Write "</html>"
		End Sub
		 Sub showContent()
		 Do While Not LabelRS.EOF
		   Response.Write "<tr>"
		   Response.Write " <td>"
		   Response.Write " <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		   Response.Write "     <tr>"
				  
				  Dim T, TitleStr, FolderName, ShortName, LabelTypeStr
					   For T = 1 To Row
					   If Not LabelRS.EOF Then
						  If LabelRS(4) = 0 Then
							  FolderName = LabelRS(1)
							  ShortName = KS.ListTitle(FolderName, 24)
							  TitleStr = " TITLE='�� ��:" & FolderName & "&#13;&#10;�� ��:��ǩĿ¼'"
						  Else
							  FolderName = LabelRS(1)
							  ShortName = KS.ListTitle(Replace(Replace(Replace(FolderName, "{LB_", ""), "}", ""),"{SQL_",""), 24)
							  If LabelType = 1 Then
								 LabelTypeStr = "�Զ��徲̬��ǩ"
							  ElseIf LabelType=5 Then
							     LabelTypeStr = "�Զ��庯����ǩ"
							  Else
								 LabelTypeStr = "ϵͳ������ǩ"
							  End If
							  TitleStr = " TITLE='�� ��:" & FolderName & "&#13;&#10;�� ��:" & LabelTypeStr & "'"
						  End If
						   Response.Write ("<td width=""" & CInt(100 / Row) & "%"" Style=""cursor:default"" align=""center""" & TitleStr & ">")
						If LabelRS(4) = 0 Then
						   Response.Write ("<span onmousedown=""mousedown(this);"" FolderID=""" & LabelRS(0) & """ style=""POSITION:relative;"" OnDblClick=""OpenLabelFolder('" & LabelRS(0) & "')"">")
						Else
						   Response.Write ("<span onmousedown=""mousedown(this);"" LabelID=""" & LabelRS(0) & """ style=""POSITION:relative;"" onDblClick=""EditLabel('','" & LabelRS(0) & "');"">")
						End If
					 If LabelRS(4) = 0 Then
					   Response.Write ("<img src=""../Images/Folder/folder.gif""> ")
					 ElseIf LabelType = 1 Then
					   Response.Write ("<img src=""../Images/Label/Label3.gif"">")
					 ElseIF LabelType=5 Then
					  Response.Write ("<img src=""../Images/Label/Label5.gif"">")
					 Else
					   Response.Write ("<img src=""../Images/Label/Label" & LabelRS(3) & ".gif"">")
					 End If
					Response.Write ("<span style=""display:block;height:16;padding:0px 0px 0px 0px;margin:1px;width:80%;cursor:default"">" & ShortName & "</span>")
					Response.Write ("</span>")
					Response.Write ("</td>")
					i = i + 1
					   If LabelRS.EOF Or i >= MaxPerPage Then Exit For
					   LabelRS.MoveNext
					 Else
					  Exit For
					 End If
				Next
				'����7����Ԫ��,����в���
				Do While T <= Row
				 Response.Write ("<td width=70>&nbsp;</td>")
				 T = T + 1
				 Loop
				   
		   Response.Write "     </tr>"
		   Response.Write "     <tr><td colspan=" & Row & " height=10></td></tr>"
		   Response.Write "   </table></td>"
		   Response.Write "</tr>"
		
				  If i >= MaxPerPage Then Exit Do
				  If LabelRS.EOF Then Exit Do
				Loop
				  LabelRS.Close
				  Conn.Close
		
		  Response.Write "        <td  align=""right"">"
			 Call KS.ShowPageParamter(totalPut, MaxPerPage, "Label_Main.asp", True, "��", CurrentPage, "LabelType=" & LabelType & "&FolderID=" & FolderID & "&" & SearchParam)
		  Response.Write " </td>"
		  Response.Write "    </tr>"
		End Sub
		
		'��¡��ǩ������
		Sub SetPasteParam()
		Dim LabelID:LabelID=KS.G("LabelID")
		Dim NewLabelName:NewLabelName="����_" & Replace(Replace(Replace(Conn.Execute("Select LabelName From KS_Label Where ID='" & LabelID & "'")(0), "{LB_", ""),"{SQL_",""), "}", "")
		Response.Write "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">"
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		Response.Write "<title>��ǩ����</title>"
		Response.Write "</head>"
		Response.Write "<link href=""ModeWindow.css"" rel=""stylesheet"">"
		Response.Write "<link href=""Admin_Style.css"" rel=""stylesheet"">"
		Response.Write "<script language=""JavaScript"" src=""Common.js""></script>"
		Response.Write "<body scroll=no topmargin=""0"" leftmargin=""0"">"
		Response.Write "  <form name=""LabelPasteForm"" method=""post"" action=""?Action=PasteSave"">"
		Response.Write "  <input type=""hidden"" value=""" & LabelID & """ name=""LabelID"">"
		Response.Write "  <table width=""96%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"
		Response.Write "    <tr>"
		Response.Write "      <td>"
		Response.Write "      <FIELDSET align=center>"
		Response.Write "      <LEGEND align=left>"
        Response.Write "         ��¡��ǩ"
		Response.Write "       </LEGEND>"

		Response.Write "  <table width=""100%"" height=""30"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		Response.Write "    <tr>"
		Response.Write "      <td height=""40"" align=""center"">"
		Response.Write "       ��¡��ǩ�����ƣ�"
		Response.Write "        <input type=""text"" name=""NewLabelName"" size='30' class='textbox' value=""" & NewLabelName & """>"
		Response.Write "        <input type=""hidden"" name=""labelType"" value=""" & Conn.Execute("Select LabelType From KS_Label Where ID='" & LabelID & "'")(0) &""">"
		Response.Write "      </td>"
		Response.Write "    </tr>"
		Response.Write "    <tr>"
		Response.Write "      <td height=""40"" align=""center"">"
		Response.Write "        <input type=""button"" name=""Submit""  class=""button"" Onclick=""CheckForm()"" value="" ȷ �� "">"
		Response.Write "        <input type=""button"" name=""Submit2""  class=""button"" onclick=""window.close()"" value="" ȡ �� "">"
		Response.Write "      </td>"
		Response.Write "    </tr>"
		Response.Write "  </table>"
		Response.Write "          </FIELDSET>"
		Response.Write "          </td></tr></table>"
		Response.Write "  </form>"

		Response.Write "</body>"
		Response.Write "</html>"
		Response.Write "<Script Language=""javascript"">" & vbCrLf
		Response.Write "function CheckForm()" & vbCrLf
		Response.Write "{ var form=document.LabelPasteForm;" & vbCrLf
		Response.Write "    if (form.NewLabelName.value.length<=0)" & vbCrLf
		Response.Write "    {"
		Response.Write "     alert(""����¿�¡�ı�ǩȡ������!"");" & vbCrLf
		Response.Write "     form.NewLabelName.focus();" & vbCrLf
		Response.Write "    return false;" & vbCrLf
		Response.Write "    }"
		Response.Write "    form.submit();" & vbCrLf
		Response.Write "    return true;" & vbCrLf
		Response.Write "}" & vbCrLf
		Response.Write "</Script>"
		End Sub
		'�����¡
		Sub LabelPasteSave()
		  Dim LabelID:LabelID=KS.G("LabelID")
		  Dim NewLabelName:NewLabelName=KS.G("NewLabelName")
		  If KS.G("LabelType")=5 Then
		  NewLabelName = "{SQL_" & NewLabelName & "}"
		  Else
		  NewLabelName = "{LB_" & NewLabelName & "}"
		  End IF
		  Dim LabelRS:Set LabelRS=Server.CreateObject("ADODB.RECORDSET")
		  LabelRS.Open "Select LabelName From KS_Label Where LabelName='" & NewLabelName & "'", Conn, 1, 1
		  If Not LabelRS.Eof Then 
		     LabelRS.Close:Set LabelRS=Nothing
		     Call KS.Alert("��ǩ�����Ѵ��ڣ���������������!","Label_Main.asp?Action=SetPasteParam&LabelID=" & LabelID)
		  End If
		    LabelRS.Close
			LabelRS.Open "Select * From KS_Label Where ID='" & LabelID & "'",Conn,1,1
			If Not LabelRS.Eof Then
			    Dim NewRS:Set NewRS=Server.CreateObject("ADODB.RECORDSET")
				NewRS.Open "Select * From KS_Label",Conn,1,3
				NewRS.AddNew
				  NewRS("ID")        = Year(Now()) & KS.MakeRandom(10)
				  NewRS("LabelName") = NewLabelName
				  NewRS("LabelContent") = LabelRS("LabelContent")
				  NewRS("Description") = LabelRS("Description")
				  NewRS("FolderID")    = LabelRS("FolderID")
				  NewRS("OrderID")     = LabelRS("OrderID")
				  NewRS("LabelType")   = LabelRS("LabelType")
				  NewRS("LabelFlag")   = LabelRS("LabelFlag")
				  NewRS("AddDate")     = Now
				  NewRS.Update
				  NewRS.Close:Set NewRS=Nothing
				  LabelRS.Close:Set LabelRS=Nothing
				  Response.Write "<script>window.close();</script>"
			Else
			  Response.Write "<script>alert('��¡ʧ��!');window.close();</script>"
			End If
		End Sub
		
		'ɾ����ǩĿ¼
		Sub LabelFolderDel()
		   Dim RS,K, ID, ParentID, FolderSql,LabelFolderID
		   Set RS=Server.CreateObject("ADODB.Recordset")
		   ID = Split(Request("ID"), ",")     '���Ҫɾ��Ŀ¼��ID����
			For K = LBound(ID) To UBound(ID)
			  FolderSql = "select ParentID,FolderType from [KS_LabelFolder] where ID='" & ID(K) & "'"
			  RS.Open FolderSql, Conn, 1, 1
			  If Not RS.EOF Then
				ParentID = Trim(RS(0))
				Dim RsObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
				RSObj.Open " Select TS From  KS_LabelFolder WHERE ID='" & ID(K) & "' OR TS like '%" & ID(K) & "%'",Conn,1,3
				Do While Not RSObj.Eof
				 Dim TS:TS=Replace(RSObj(0),",","','")
				 Conn.Execute ("DELETE  FROM KS_Label WHERE FolderID In('" & mid(ts,instr(ts,id(k)))  & "') and FolderID<>'" & ParentID & "'")
				 RSObj.Delete
				RSObj.MoveNext
				Loop
				RSObj.Close:Set RSObj=Nothing
			   End If
			  RS.Close
			Next
		  Set RS = Nothing
		  Response.Write "<script>location.href='Label_Main.asp?LabelType=" & LabelType & "&Folderid=" & ParentID & "'</script>"
		End Sub
		'ɾ����ǩ
		Sub LabelDel()
			Dim K, ID,Page
			Page = KS.G("Page")
			ID = Split(Request("id"), ",") '���Ҫɾ����ǩ��ID����
			For K = LBound(ID) To UBound(ID)
			  Conn.Execute("Delete FROM KS_Label WHERE ID='" & ID(K) & "'")
			Next
			Response.Write "<script>location.href='Label_Main.asp?Page=" & Page & "&LabelType=" & LabelType & "&FolderID=" & FolderID & "';</script>"

		End Sub
		
		Sub LabelOut()
		Response.Write "<body>"
		Response.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0""   class=""sortbutton"">"
		Response.Write "  <tr>"
		Response.Write "    <td height=""23"" align=""left"">��������<a href='?Action=LabelIn&LabelType=" & LabelType & "'>��ǩ����</a> | <a href='?Action=LabelOut&LabelType=" & LabelType & "'>��������</a></td>"
		Response.Write " </tr>"
		Response.Write "</table>"

		  LabelType=KS.G("LabelType")
		  %>
		  <Script language="Javascript">
		  var ClassArr = new Array();
		  <%
			Response.Write "ClassArr[0] =new Array(""" & GetLabelOption(0,conn) & """);" & vbcrlf
			Response.Write "ClassArr[1] =new Array(""" & GetLabelOption(1,conn) & """);" & vbcrlf
			Response.Write "ClassArr[5] =new Array(""" & GetLabelOption(5,conn) & """);" & vbcrlf
			Response.Write "ClassArr[9999] =new Array(""" & GetLabelOption(0,conn)&GetLabelOption(1,conn)&GetLabelOption(5,conn) & """);" & vbcrlf
		  %>
		  </Script>
		  <form name='myform' method='post' action='Label_main.asp'>  
		  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>    
		  <tr class='title'>       
		  <td colspan="2" height='22' align='center'><strong>��ǩ����</strong></td>    
		  </tr>    
		  <tr class='tdbg'>
			  <td width="10%" height='10' align="right">ѡ�����ͣ�</td>
			  <td width="90%">
			  <select id="LabelType" name="LabelType" onChange="SelectClass(this.value)">
			  <option value="9999">ȫ����ǩ</option>
			  <option value="0"<%IF LabelType="0" Then Response.write " selected"%>>ϵͳ������ǩ</option>
			  <option value="5"<%IF LabelType="5" Then Response.write " selected"%>>�Զ��庯����ǩ</option>
			  <option value="1"<%IF LabelType="1" Then Response.write " selected"%>>�Զ��徲̬��ǩ</option>
		    </select>
			</td>
		  </tr>    
		  <tr class='tdbg'>      
		  <td colspan="2" align='center'>        
		    <table width="100%" border='0' cellpadding='0' cellspacing='0'>          
			   <tr>           
			     <td width="10%" align="right">��ǩ�б�</td>
				 <td width="54%" ID="ClassArea"><select name='LabelID' size='2' multiple style='height:300px;width:450px;'>
				 </select></td>                 <td width="36%" align='left'>&nbsp;&nbsp;&nbsp;&nbsp;
				   <input type='button' name='Submit' value=' ѡ������ ' onclick='SelectAll()'>    <br><br>&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' name='Submit' value=' ȡ��ѡ�� ' onclick='UnSelectAll()'><br><br><br><b>&nbsp;��ʾ����ס��Ctrl����Shift�������Զ�ѡ</b></td>      
			 </tr>     
			 <tr height='30'>        <td colspan='2'>��Ŀ�����ݿ⣺
			     <input name='LabelMdb' type='text' id='LabelMdb' value='<%=KS.Setting(3)%>Label.mdb' size='20' maxlength='50'>
			 &nbsp;&nbsp;�˲��������Ŀ�����ݿ�</td>      
			 </tr>      
		    <tr height='50'>         <td colspan='2' align='center'><input type='submit' name='Submit' value='ִ�е�������' onClick="document.myform.Action.value='Doexport';">              <input name='Action' type='hidden' id='Action' value='Doexport'>         </td>        </tr>    </table>   
		    </td> </tr></table></form>
		  <script language='javascript'>
		  SelectClass(<%=LabelType%>);
function SelectClass(LabelType)
{ document.all.ClassArea.innerHTML='<select name="LabelID" size="2" multiple style="height:300px;width:450px;">'+ClassArr[LabelType]+'</select>';
}
function SelectAll(){
  for(var i=0;i<document.myform.LabelID.length;i++){
    document.myform.LabelID.options[i].selected=true;}
}
function UnSelectAll(){
  for(var i=0;i<document.myform.LabelID.length;i++){
    document.myform.LabelID.options[i].selected=false;}
}
</script>
		  <%
		End Sub
		Function GetLabelOption(LabelType,DBC)
		  Dim AllLabel,RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "Select * From KS_Label Where LabelType=" & LabelType,DBC,1,1
		  Do While Not RS.Eof 
			AllLabel=AllLabel & "<option value='" & RS("ID") & "'>" & RS("LabelName") & "</option>"
			RS.MoveNext
		  Loop
          RS.Close:Set RS=Nothing
		  GetLabelOption=AllLabel
		End Function
		'��������
		Sub Doexport()
		 Dim LabelID:LabelID="'"& Replace(Replace(KS.G("LabelID")," ",""),",","','") & "'"
		 Dim LabelMdb:LabelMdb=KS.G("LabelMdb")
		 Dim rs:set rs=server.createobject("adodb.recordset")
		 Dim sqlstr,n
		   n=0
		   sqlstr="select ID,LabelName,LabelContent,Description,FolderID,OrderID,LabelType,LabelFlag,AddDate from ks_label where id in(" & LabelID & ")"
		         'on error resume next
			     if CreateDatabase(LabelMdb)=true then
						Dim DataConn:Set DataConn = Server.CreateObject("ADODB.Connection")
	                    DataConn.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(LabelMdb)
						If not Err Then
						   If Checktable("KS_Label",DataConn)=true Then
						     DataConn.Execute("drop table KS_Label")
						   end if
				             Dataconn.execute("CREATE TABLE [KS_Label] ([LabelID] int IDENTITY (1, 1) NOT NULL CONSTRAINT PrimaryKey PRIMARY KEY,[ID] varchar(50) Not Null,[LabelName] varchar(255) Not Null,[LabelContent] text not null,[Description] text null,[FolderID] varchar(100) not null,[OrderID] int not null,[LabelType] int not null,[LabelFlag] int not null,[AddDate] date not null)")
						  rs.open sqlstr,conn,1,1
						 if not rs.eof then
						   	Dim RST:Set RST=Server.CreateObject("ADODB.RECORDSET")
						   do while not rs.eof
							  n=n+1
						      'DataConn.Execute("Insert Into KS_Label(ID,LabelName,LabelContent,Description,FolderID,OrderID,LabelType,LabelFlag,AddDate) values('" & rs(0) & "','" & rs(1) & "','" &rs(2) & "','" & rs(3) & "','" & rs(4) & "'," & rs(5) & "," & rs(6) & "," & rs(7) & ",'" & rs(8) & "')")
							  RST.Open "Select * From KS_Label where 1=0",DataConn,1,3
							  RST.AddNew
							    RST("ID")=rs(0)
								RST("LabelName")=rs(1)
								RST("LabelContent")=rs(2)
								RST("Description")=rs(3)
								RST("FolderID")=rs(4)
								RST("OrderID")=rs(5)
								RST("LabelType")=rs(6)
								RST("LabelFlag")=rs(7)
								RST("AddDate")=rs(8)
							  RST.Update
							  RST.Close
							  rs.movenext
						   loop
						   Set RST=Nothing
						 end if
                          rs.close:set rs=nothing
						End if
						DataConn.Close:Set DataConn=Nothing
				 end if
				response.write "<br><br><br><div align=center>�������!�ɹ������� <font color=red>" & n & "</font> ����ǩ��<a href=" & LabelMdb & ">������������</a>(�Ҽ�Ŀ�����Ϊ)  </div><br><br><br><br><br><br><br>"

		End Sub
		
		Sub LabelIn()
				Response.Write "<body>"
		Response.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0""   class=""sortbutton"">"
		Response.Write "  <tr>"
		Response.Write "    <td height=""23"" align=""left"">��������<a href='?Action=LabelIn&LabelType=" & LabelType & "'>��ǩ����</a> | <a href='?Action=LabelOut&LabelType=" & LabelType & "'>��������</a></td>"
		Response.Write " </tr>"
		Response.Write "</table>"

		%>
		<form name='myform' method='post' action='Label_Main.asp?LabelType=<%=KS.G("LabelType")%>'>  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>    <tr class='title'>       <td height='22' align='center'><strong>��ǩ���루��һ����</strong></td>    </tr>    <tr class='tdbg'>      <td height='100'>&nbsp;&nbsp;&nbsp;&nbsp;������Ҫ����ı�ǩ���ݿ���ļ�����         <input name='LabelMdb' type='text' id='LabelMdb' value='<%=KS.Setting(3)%>Label.mdb' size='20' maxlength='50'>        <input name='Submit' type='submit' id='Submit' value=' ��һ�� '>        <input name='Action' type='hidden' id='Action' value='LabelIn2'>      </td>    </tr>  </table></form>
		<%
		End Sub
		
		Sub LabelIn2()
				Response.Write "<body>"
		Response.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0""   class=""sortbutton"">"
		Response.Write "  <tr>"
		Response.Write "    <td height=""23"" align=""left"">��������<a href='?Action=LabelIn&LabelType=" & LabelType & "'>��ǩ����</a> | <a href='?Action=LabelOut&LabelType=" & LabelType & "'>��������</a></td>"
		Response.Write " </tr>"
		Response.Write "</table>"

		on error resume next
		LabelType=KS.G("LabelType")
		Dim LabelMdb:LabelMdb=KS.G("LabelMdb")
		Dim DataConn:Set DataConn = Server.CreateObject("ADODB.Connection")
	    DataConn.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(LabelMdb)
		%>
		<form name='myform' method='post' action='Label_Main.asp'>  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>    <tr class='title'>       <td height='22' align='center'><strong>��ǩ���루�ڶ�����</strong></td>    </tr>    <tr class='tdbg'>       <td height='100' align='center'>        <br>        <table border='0' cellspacing='0' cellpadding='0'>          
		<%
		If Err Then 
		Err.Clear:Set DataConn = Nothing:Response.Write "<tr><td>���ݿ�·������ȷ�����ӳ���</td></tr>":Response.End
		else
		 	%>
		  <Script language="Javascript">
		  var ClassArr = new Array();
		  <%
			Response.Write "ClassArr[0] =new Array(""" & GetLabelOption(0,DataConn) & """);" & vbcrlf
			Response.Write "ClassArr[1] =new Array(""" & GetLabelOption(1,DataConn) & """);" & vbcrlf
			Response.Write "ClassArr[5] =new Array(""" & GetLabelOption(5,DataConn) & """);" & vbcrlf
			Response.Write "ClassArr[9999] =new Array(""" & GetLabelOption(0,DataConn)&GetLabelOption(1,DataConn)&GetLabelOption(5,DataConn) & """);" & vbcrlf
		  %>
		  </Script>
		<tr> <td><strong>ѡ��Ҫ����ı�ǩ�ķ��ࣺ</strong><select id="LabelType" name="LabelType" onChange="SelectClass(this.value)">
			  <option value="9999">ȫ����ǩ</option>
			  <option value="0"<%IF LabelType="0" Then Response.write " selected"%>>ϵͳ������ǩ</option>
			  <option value="5"<%IF LabelType="5" Then Response.write " selected"%>>�Զ��庯����ǩ</option>
			  <option value="1"<%IF LabelType="1" Then Response.write " selected"%>>�Զ��徲̬��ǩ</option>
		    </select></td></tr>   
  		<tr>
		<td><strong>��������ʽ��</strong> 
		<input type="radio" value="0" name="cl" checked>��ǩ��������
		<input type="radio" value="1" name="cl">��ǩ��������
		</td>
		</tr>  

		<tr>
		<td id="ClassArea"> 
		<select name='LabelID' size='2' multiple style='height:300px;width:350px;'> </select>
		</td>
		</tr>  
		<%end if%>                <tr><td colspan='3' height='5'></td></tr>                  <tr>                    <td height='25' align='center'><b> ��ʾ����ס��Ctrl����Shift�������Զ�ѡ</b></td>                  </tr>    <tr><td colspan='3' height='25' align='center'>
		<input type='submit' name='Submit' value=' �����ǩ ' onClick="document.myform.Action.value='Doimport';" >                </td></tr>               </table>               <input name='LabelMdb' type='hidden' id='LabelMdb' value='<%=LabelMdb%>'>               <input name='Action' type='hidden' id='Action' value='Doimport'>               <br>            </td>          </tr>       
		</table></form>
		<script language='javascript'>
		  SelectClass(<%=LabelType%>);
		function SelectClass(LabelType)
		{ document.all.ClassArea.innerHTML='<select name="LabelID" size="2" multiple style="height:300px;width:350px;">'+ClassArr[LabelType]+'</select>';
		}
   </script>
		<%
		dataconn.close:set dataconn=nothing
		End Sub
		'�������
		Sub Doimport()
			'on error resume next
			Dim n:n=0
			Dim m:m=0
			Dim k:k=0
			Dim LabelMdb:LabelMdb=KS.G("LabelMdb")
			Dim NewLabelID,cl:cl=KS.G("cl")
			Dim LabelID:LabelID="'"& Replace(Replace(KS.G("LabelID")," ",""),",","','")& "'"
			Dim DataConn:Set DataConn = Server.CreateObject("ADODB.Connection")
			DataConn.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(LabelMdb)
			If Err Then 
			Err.Clear:Set DataConn = Nothing:Response.Write "<tr><td>���ݿ�·������ȷ�����ӳ���</td></tr>":Response.End
			else
			 Dim rs:set rs=server.createobject("adodb.recordset")
			 rs.open "select * from ks_label where ID in(" & LabelID & ")",dataconn,1,1
			 Dim rsa:set rsa=server.createobject("adodb.recordset")
			 do while not rs.eof 
			  rsa.open "select * from ks_label where labelname='" & rs("labelname") & "'",conn,1,3
			  if rsa.eof then
			     rsa.addnew
				  Do While True
					'����ID  ��+10λ���
					NewLabelID = Year(Now()) & KS.MakeRandom(10)
					Dim RSCheck:Set RSCheck = Conn.Execute("Select ID from [KS_Label] Where ID='" & NewLabelID & "'")
					 If RSCheck.EOF And RSCheck.BOF Then
					  RSCheck.Close:Set RSCheck = Nothing:Exit Do
					 End If
				  Loop
			     rsa("ID")=NewLabelID
				 rsa("LabelName")=rs("LabelName")
				 rsa("LabelContent")=rs("LabelContent")
				 rsa("Description")=rs("Description")
				 if conn.execute("select ID from ks_labelfolder where id='" & rs("folderid") & "'").eof then
				 rsa("FolderID")="0"
				 else
				 rsa("FolderID")=rs("folderid")
				 end if
				 rsa("OrderID")=rs("OrderID")
				 rsa("LabelType")=rs("LabelType")
				 rsa("LabelFlag")=rs("LabelFlag")
				 rsa("AddDate")=rs("AddDate")
				 n=n+1
				rsa.update
			  else   '��������
			   if cl="1" then
				 rsa("LabelContent")=rs("LabelContent")
				 rsa("Description")=rs("Description")
				 rsa("OrderID")=rs("OrderID")
				 rsa("LabelType")=rs("LabelType")
				 rsa("LabelFlag")=rs("LabelFlag")
				 rsa("AddDate")=rs("AddDate")
				 m=m+1
				rsa.update
			   else
			    k=K+1
			   end if
			  end if
			   rsa.close
			  rs.movenext
			 loop
			 rs.close:set rs=nothing
			 set rsa=nothing
			end if
			response.write "<br><br><br><div align=center>�������!�ɹ������� <font color=red>" & n & "</font> ����ǩ,������ <font color=red>" & m & "</font> ����ǩ,���������� <font color=red>" & k & "</font> ����ǩ��  </div><br><br><br><br><br><br><br>"
           dataconn.close:set dataconn=nothing
		End Sub
		Function CreateDatabase(dbname)
		      if KS.CheckFile(dbname) then CreateDatabase=true:exit function
				dim objcreate :set objcreate=Server.CreateObject("adox.catalog") 
				if err.number<>0 then 
					set objcreate=nothing 
					CreateDatabase=false
					exit function 
				end if 
				'�������ݿ� 
				objcreate.create("data source="+server.mappath(dbname)+";provider=microsoft.jet.oledb.4.0") 
				if err.number<>0 then 
					CreateDatabase=false
					set objcreate=nothing 
					exit function
				end if 
				CreateDatabase=true
		End Function
		'������ݱ��Ƿ����	
		Function Checktable(TableName,DataConn)
			On Error Resume Next
			DataConn.Execute("select * From " & TableName)
			If Err.Number <> 0 Then
				Err.Clear()
				Checktable = False
			Else
				Checktable = True
			End If
		End Function

End Class
%> 
