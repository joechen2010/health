<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.CollectCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Collect_Main
KSCls.Kesion()
Set KSCls = Nothing

Class Collect_Main
        Private KS
		Private KMCObj
		Private ConnItem,ChannelID
		'=================================================================================================
		Private i
		Private totalPut
		Private CurrentPage
		Private SqlStr
		Private RSObj
		Private MaxPerPage
		'=================================================================================================
		Private Sub Class_Initialize()
		  MaxPerPage = 20
		  Set KS=New PublicCls
		  Set KMCObj=New CollectPublicCls
		  Set ConnItem = KS.ConnItem()
		End Sub
        Private Sub Class_Terminate()
		 Call KS.CloseConnItem()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KMCObj=Nothing
		End Sub
		Sub Kesion()
			
			If Request("page") <> "" Then
				  CurrentPage = CInt(Request("page"))
			Else
				  CurrentPage = 1
			End If
			ChannelID=KS.ChkClng(KS.G("ChannelID"))
			If Not KS.ReturnPowerResult(0, "M010008") Then                  'Ȩ�޼��
				Call KS.ReturnErr(1, "")   
				Response.End()
			End if
			
			'response.write channelid
			Select Case  KS.G("Action")
			 Case "Del"
			    Dim ItemID:ItemID = KS.FilterIds(Replace(KS.G("ItemID"), " ", ""))
				ConnItem.Execute ("Delete From KS_CollectItem Where ItemID In(" & ItemID & ")")
				ConnItem.Execute ("Delete From KS_FieldRules Where ItemID In(" & ItemID & ")")
				ConnItem.Execute ("Delete From KS_Filters Where ItemID In(" & ItemID & ")")
				ConnItem.Execute ("Delete From KS_History Where ItemID In(" & ItemID & ")")
				Response.Write "<script>alert('��ϲ,�ɼ���Ŀɾ���ɹ�!');location.href='" & request.servervariables("http_referer") & "';</script>"
			Case "Paste"
			 Call ItemPaste()
			case "delhistory"
			    ItemID = KS.FilterIds(replace(KS.G("ItemID"), " ", ""))
				ConnItem.Execute ("Delete From KS_History Where ItemID In(" & ItemID & ")")
				Response.Write "<script>alert('��ϲ,�ɼ���ʷ��¼����ɹ�!');location.href='" & request.servervariables("http_referer") & "';</script>"
			Case else
			 Call ItemList()
			End Select
          End Sub
		  
		  Sub ItemList()
			Response.Write "<html>"
			Response.Write "<head>"
			Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			Response.Write "<title>�ɼ���Ŀ����</title>"
			Response.Write "<link href=""../Include/Admin_Style.css"" rel=""stylesheet"" type=""text/css"">"
			Response.Write "<script language=""JavaScript"">" & vbCrLf
			Response.Write "var Page='" & CurrentPage & "';" & vbCrLf
			Response.Write "</script>" & vbCrLf
			Response.Write "<script language=""JavaScript"" src=""../../KS_Inc/common.js""></script>"
			Response.Write "<script language=""JavaScript"" src=""../../KS_Inc/jquery.js""></script>"
			Response.Write "<script language=""JavaScript"" src=""../Include/ContextMenu.js""></script>"
			Response.Write "<script language=""JavaScript"" src=""../Include/SelectElement.js""></script>"
			%>
			<script>
			var DocElementArrInitialFlag=false;
			var DocElementArr = new Array();
			var DocMenuArr=new Array();
			var SelectedFile='',SelectedFolder='';
			function document.onreadystatechange()
			{   if (DocElementArrInitialFlag) return;
				InitialDocElementArr('FolderID','ItemID');
				InitialContextMenu();
				DocElementArrInitialFlag=true;
			}
			function InitialContextMenu()
			{	DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.CreateCollectItem('');",'�����Ŀ(N)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.SetCollectItemPro('');",'��������(P)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.TestCollectItem('');",'��Ŀ����(T)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.SelectAllElement();",'ȫ ѡ(A)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.EditCollectItem('');",'�� ��(E)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.DelCollectItem('');",'ɾ ��(D)','disabled');
			
					 DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
					//Ԥ������ DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.Cut();",'�� ��(X)','disabled');
					 DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.Copy();",'�� ��(C)','disabled');
					 DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.Paste();",'ճ ��(V)','disabled');
			
				
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.location.reload();",'ˢ ��(Z)','disabled');
			}
			function DocDisabledContextMenu()
			{   var PasteTFStr='';
				if (top.CommonCopyCut.PasteTypeID==0||top.CommonCopyCut.ChannelID!=100<%=channelid%>)PasteTFStr='ճ ��(V),';
				DisabledContextMenu('FolderID','ItemID',PasteTFStr+'�� ��(E),ɾ ��(D),�� ��(X),�� ��(C)',PasteTFStr+'',PasteTFStr+'',PasteTFStr+'',PasteTFStr+'')
			}
			function CreateCollectItem()
			{location.href='Collect_ItemModify.asp?channelid=<%=ChannelID%>';
			}
			function EditCollectItem()
			{
				GetSelectStatus('FolderID','ItemID');
			 if (SelectedFile!='')
			   if (SelectedFile.indexOf(',')==-1)
			   location.href='Collect_ItemModify.asp?ItemID='+SelectedFile;
			   else
			   alert('һ��ֻ�ܹ��༭һ���ɼ���Ŀ!'); 
			 else
			  alert('��ѡ��Ҫ�༭�Ĳɼ���Ŀ!');
			  SelectedFile='';
			}
			function DelCollectItem()
			{
			 GetSelectStatus('FolderID','ItemID');
			 if (SelectedFile!='')
			  {
			   if (confirm('���Ҫɾ��ѡ�еĲɼ���Ŀ��?'))
				location="?ChannelID=<%=ChannelID%>&Action=Del&Page="+Page+"&ItemID="+SelectedFile;
			  }
			 else
			  alert('��ѡ��Ҫɾ���Ĳɼ���Ŀ!');
			  SelectedFile='';
			}
			function SetCollectItemPro()
			{
				GetSelectStatus('FolderID','ItemID');
			 if (SelectedFile!='')
			   if (SelectedFile.indexOf(',')==-1)
			   location.href='Collect_ItemAttribute.asp?ItemID='+SelectedFile;
			   else
			   alert('һ��ֻ�ܹ�����һ���ɼ���Ŀ������!'); 
			 else
			  alert('��ѡ��Ҫ�������ԵĲɼ���Ŀ!');
			  SelectedFile='';
			}
			function TestCollectItem()
			{
				GetSelectStatus('FolderID','ItemID');
			 if (SelectedFile!='')
			   if (SelectedFile.indexOf(',')==-1)
			   location.href='Collect_ItemModify5.asp?ItemID='+SelectedFile;
			   else
			   alert('һ��ֻ�ܹ�����һ���ɼ���Ŀ!'); 
			 else
			  alert('��ѡ��Ҫ���ԵĲɼ���Ŀ!');
			  SelectedFile='';
			}
			function Cut()
			{  
				GetSelectStatus('FolderID','ItemID');
				if (!((SelectedFile=='')&&(SelectedFolder=='')))
				  {
				   top.CommonCopyCut.ChannelID=100<%=channelid%>;
				   top.CommonCopyCut.PasteTypeID=1;
				   top.CommonCopyCut.SourceFolderID=ClassID;
				   top.CommonCopyCut.FolderID=SelectedFolder;
				   top.CommonCopyCut.ContentID=SelectedFile;
				   SelectedFolder='';
				   SelectedFile='';
				  }
				else
				 alert('��ѡ��Ҫ���е�Ŀ¼����Ŀ!');
			}
			function Copy()
			{
				GetSelectStatus('FolderID','ItemID');
				if (!((SelectedFile=='')&&(SelectedFolder=='')))
				  {
				   top.CommonCopyCut.ChannelID=100<%=channelid%>;
				   top.CommonCopyCut.PasteTypeID=2;
				  // top.CommonCopyCut.SourceFolderID=ClassID;
				   top.CommonCopyCut.FolderID=SelectedFolder;
				   top.CommonCopyCut.ContentID=SelectedFile;
				   SelectedFolder='';
				   SelectedFile='';
				  }
				else
				 alert('��ѡ��Ҫ���Ƶ�Ŀ¼����Ŀ!');
			}
			function Paste()
			{ 
			  if (top.CommonCopyCut.ChannelID==100<%=channelid%> && top.CommonCopyCut.PasteTypeID!=0)
			   {  var Param='';
				  Param='?Action=Paste&ChannelID=<%=ChannelID%>&Page='+Page;
				 //Param+='&PasteTypeID='+top.CommonCopyCut.PasteTypeID+'&DestFolderID='+ClassID+'&SourceFolderID='+top.CommonCopyCut.SourceFolderID+'&FolderID='+top.CommonCopyCut.FolderID+'&ContentID='+top.CommonCopyCut.ContentID;
				 Param+='&PasteTypeID='+top.CommonCopyCut.PasteTypeID+'&DestFolderID=0&SourceFolderID='+top.CommonCopyCut.SourceFolderID+'&FolderID='+top.CommonCopyCut.FolderID+'&ContentID='+top.CommonCopyCut.ContentID;
				 if (top.CommonCopyCut.PasteTypeID==1)      //����
				 {  
					top.CommonCopyCut.PasteTypeID=0;       //����Ϊ0,ʹճ��������
					if (top.CommonCopyCut.SourceFolderID==ClassID) return;
					location.href='Collect_Main.asp'+Param;
				 }
				else if (top.CommonCopyCut.PasteTypeID==2) //����
				 {
					location.href='Collect_Main.asp'+Param;
				 }
				else
				 alert('�Ƿ�����!');
			   }
			  else
			   alert('ϵͳ���а�û������!');
			}
			function GetKeyDown()
			{ 
			if (event.ctrlKey)
			  switch  (event.keyCode)
			  {  case 90 : location.reload(); break;
				 case 65 : SelectAllElement();break;
				 case 78 : event.keyCode=0;event.returnValue=false;CreateCollectItem();break;
				 case 69 : event.keyCode=0;event.returnValue=false;EditCollectItem();break;
				 case 80 : event.keyCode=0;event.returnValue=false;SetCollectItemPro();break;
				 case 84 : event.keyCode=0;event.returnValue=false;TestCollectItem();break;
				 case 68 : DelCollectItem('');break;
				 case 67 : 
				   event.keyCode=0;event.returnValue=false;Copy();
					break;
				 case 86 : 
				   if (top.CommonCopyCut.ChannelID==100<%=channelid%> && top.CommonCopyCut.PasteTypeID!=0)
				   { event.keyCode=0;event.returnValue=false;Paste();}
				   else
					return;
					break;
			   }	
			else	
			 if (event.keyCode==46) DelCollectItem();
			}
			function CheckAll(form)
			{
			  for (var i=0;i<form.elements.length;i++)
				{
				var e = form.elements[i];
				if (e.Name != "chkAll")
				   e.checked = form.chkAll.checked;
				}
			}
			</script>
			<%
			Response.Write "</head>"
			Response.Write "<body scroll=no topmargin=""0"" leftmargin=""0"" onclick=""SelectElement();"" onkeydown=""GetKeyDown();"" onselectstart=""return false;"">"
		    Response.Write "<ul id='menu_top'>"
			Response.Write "<li class='parent' onclick='CreateCollectItem();'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/a.gif' border='0' align='absmiddle'>�½���Ŀ</span></li>"
			Response.Write "<li class='parent' onclick='location.href=""Collect_ItemFilters.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/move.gif' border='0' align='absmiddle'>��������</span></li>"
			Response.Write "<li class='parent' onclick='location.href=""Collect_IntoDatabase.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/save.gif' border='0' align='absmiddle'>������</span></li>"
			Response.Write "<li class='parent' onclick='location.href=""Collect_ItemHistory.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/Recycl.gif' border='0' align='absmiddle'>��ʷ��¼</span></li>"
			Response.Write "<li class='parent' onclick='location.href=""Collect_Field.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/addjs.gif' border='0' align='absmiddle'>�Զ����ֶ�</span></li>"
			Response.Write "<li disabled class='parent' onclick='location.href=""Collect_main.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/back.gif' border='0' align='absmiddle'>����һ��</span></li>"
			Response.Write ("</ul>")
			
			Response.Write ("<div style=""height:94%; overflow: auto; width:100%"" align=""center"">")
			Response.Write "<table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
			Response.Write "  <tr>"
			Response.Write "    <td height=""22"" class=""sort""><div align=""center"">��Ŀ����</div></td>"
			Response.Write "    <td width=""28%"" class=""sort""><div align=""center""><span>�ɼ�(վ��)��ַ</span></div></td>"
			Response.Write "    <td width=""10%"" align=""center"" class=""sort"">�ɻ���Ŀ</td>"
			Response.Write "    <td width=""14%"" class=""sort""><div align=""center"">�ϴβɼ�</div></td>"
			Response.Write "    <td width=""5%"" align=""center"" class=""sort"">״̬</td>"
			Response.Write "    <td align=""center"" class=""sort"">����</td>"
			Response.Write "  </tr>"
			   Set RSObj = Server.CreateObject("ADODB.RecordSet")
					   RSObj.Open "select ItemID,ItemName,WebName,ListStr,ListPageType,ListPageStr2,ListPageID1,ListPageID2,ListPageStr3,ChannelID,ClassID,SpecialID,Flag From KS_CollectItem order by ItemID DESC", ConnItem, 1, 1
					 If Not RSObj.EOF Then
						totalPut = RSObj.RecordCount
			
								If CurrentPage < 1 Then
									CurrentPage = 1
								End If
			
			
								If CurrentPage = 1 Then
									Call showContent
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrentPage - 1) * MaxPerPage
										Call showContent
									Else
										CurrentPage = 1
										Call showContent
									End If
								End If
				End If
				 
			Response.Write ("</table>")
			Response.Write ("</div>")
			Response.Write ("</body>")
			Response.Write ("</html>")
			
			End Sub
			Sub showContent()
			   Dim Rs, ItemCollecDate
			   Dim ItemID, ItemName, WebName, ChannelID, ClassID, SpecialID, ListStr, ListPageType, ListPageStr2, ListPageID1, ListPageID2, ListPageStr3, Flag, ListUrl
			     Response.Write "<form name='myform' method='Post' action='Collect_ItemCollection.asp'>"
					Do While Not RSObj.EOF
					
					 ItemID = RSObj("ItemID")
				  ItemName = RSObj("ItemName")
				  WebName = RSObj("WebName")
				  ChannelID = RSObj("ChannelID")
				  ClassID = RSObj("ClassID")
				  SpecialID = RSObj("SpecialID")
				  ListStr = RSObj("ListStr")
				  ListPageType = RSObj("ListPageType")
				  ListPageStr2 = RSObj("ListPageStr2")
				  ListPageID1 = RSObj("ListPageID1")
				  ListPageID2 = RSObj("ListPageID2")
				  ListPageStr3 = RSObj("ListPageStr3")
				  Flag = RSObj("Flag")
				  If ListPageType = 0 Or ListPageType = 1 Then
						ListUrl = ListStr
				  ElseIf ListPageType = 2 Then
						ListUrl = Replace(ListPageStr2, "{$ID}", CStr(ListPageID1))
				  ElseIf ListPageType = 3 Then
						If InStr(ListPageStr3, "|") > 0 Then
						ListUrl = Left(ListPageStr3, InStr(ListPageStr3, "|") - 1)
					 Else
						   ListUrl = ListPageStr3
					 End If
				  End If
				  
					  Response.Write "<tr>"
					  Response.Write "  <td class='splittd' height='18'><input type='checkbox' name='itemid' value='" &itemid & "'><span ondblclick='EditCollectItem()' ItemID='" & ItemID & "'><img src='../Images/arrow.gif'  align='absmiddle'>"
					  Response.Write " <span style='cursor:default;'>" & KS.Gottopic(ItemName,25) & "</span></span></td>"
					  Response.Write "  <td class='splittd' align='center'><a href='" & ListUrl & "' target='_blank'>" & WebName & "</a></td>"
					  Response.Write "  <td  class='splittd' align='center'>" & KMCObj.Collect_ShowClass_Name(ChannelID, ClassID) & "</td>"
					  Response.Write "  <td  class='splittd' align='center'>"
			
					  '�ϴβɼ�
					  Set Rs = ConnItem.Execute("select Top 1 CollecDate From KS_History Where ItemID=" & ItemID & " Order by HistoryID desc")
					  If Not Rs.EOF Then
						ItemCollecDate = Rs("CollecDate")
					  Else
						ItemCollecDate = ""
					  End If
					  Set Rs = Nothing
					 If ItemCollecDate <> "" Then
						Response.Write ItemCollecDate
					 Else
						Response.Write "���޼�¼"
					 End If
					 
					  Response.Write " </td>"
					  
					 Response.Write "  <td  class='splittd' align='center'>"
					  '״̬
					  If Flag = True Then
								Response.Write "��"
					  Else
							 Response.Write "<font color=red>��</font>"
					  End If
					  Response.Write "</td>"
					  Response.Write "<td  class='splittd'><a href='Collect_ItemCollection.asp?ChannelID=" & ChannelID&"&ItemID=" & itemid & "&Action=Start&NewsFalseNum=0&ImagesNumAll=0'>�ɼ�</a> <a href='Collect_ItemModify.asp?ItemID=" & itemid & "'>�༭</a> <a href='?ChannelID=" & ChannelID & "&Action=Del&Page=" & CurrentPage & "&ItemID=" & itemid & "' onclick=""return(confirm('ȷ��ɾ���ɼ���Ŀ��'));"">ɾ��</a> <a href='Collect_ItemModify5.asp?ItemID=" & itemid & "'>����</a> <a href='Collect_ItemAttribute.asp?ItemID=" & itemid & "'>����</a> <a href='?action=delhistory&itemid=" & itemid&"' title='��ղɼ���ʷ��¼!' onclick=""return(confirm('��ղɼ���ʷ��¼���ܵ���,�ظ��ɼ�!ȷ��ɾ����?'))"">��ղɼ���¼</a></td>"
					  Response.Write "</tr>"

					i = i + 1
					  If i >= MaxPerPage Then Exit Do
						   RSObj.MoveNext
					Loop
					  RSObj.Close
					  ConnItem.Close
					 Response.Write "<tr><td colspan=7><input name='chkAll' type='checkbox' id='chkAll' onclick=CheckAll(this.form) value='checkbox'>����ѡ����Ŀ <input type='submit' onclick='this.form.action=""Collect_ItemCollection.asp?ChannelID=" & ChannelID&"&Action=Start&CollecType=1"";' value='�����ɼ�ѡ����' class='button'></td></tr>"
					 Response.Write "</form>"
					 Response.Write "<tr><td height='26' colspan='6' align='right'>"
					 Call KS.ShowPageParamter(totalPut, MaxPerPage, "Collect_Main.asp", True, "��", CurrentPage, "ChannelID=" & ChannelID)
				 End Sub
				 
				 
				 'ճ��
				 Sub ItemPaste()
		 Dim DisplayMode, Page
		 Dim PasteTypeID, DestFolderID, SourceFolderID, FolderID, ContentID
		  DisplayMode = KS.G("DisplayMode")
		  Page = KS.G("Page")
		  PasteTypeID = KS.G("PasteTypeID")
		  DestFolderID = KS.G("DestFolderID")
		  SourceFolderID = KS.G("SourceFolderID")
		  FolderID = KS.G("FolderID")
		  ContentID = KS.G("ContentID")
		  If PasteTypeID = "" Then PasteTypeID = 0
		  If DestFolderID = "" Then DestFolderID = "0"
		  If FolderID = "" Then
			 FolderID = "0"
		  End If
		  If ContentID = "" Then
			 ContentID = "0"
		  Else
			 ContentID = "'" & Replace(ContentID, ",", "','") & "'"
		  End If
		  If ContentID = "" Then
			Call KS.AlertHistory("�������ݳ���!", 1)
			Set KS = Nothing
			Exit Sub
		  End If
		  
		  If PasteTypeID = 2 Then '���Ʋ���
			Call PasteByCopy(SourceFolderID, DestFolderID, FolderID, ContentID)
		  Else
			Call KS.AlertHistory("�Ƿ�����!", 1)
			Set KS = Nothing
			Exit Sub
		  End If
		  Response.Write "<script>location.href='Collect_main.asp?ChannelID=" & KS.G("ChannelID") & "&Page=" & Page & "';</script>"
		End Sub
		
		
		
		'����:PasteByCopy����ճ��
		'����:SourceFolderID--ԴĿ¼,DestFolderID--Ŀ��Ŀ¼,FolderID---�����Ƶ�Ŀ¼,ContentID---�����Ƶ��ļ�
		Sub PasteByCopy(SourceFolderID, DestFolderID, FolderID, ContentID)
		       Dim ItemName,RS,RSA,I,NewItemID
			   
			ContentID=Replace(Replace(ContentID,"'",""),"""","")
			if instr(contentid,",") then call KS.AlertHistory("�Բ���һ��ֻ�ܸ���һ����Ŀ!",-1):exit sub
			Set RS=Server.CreateObject("Adodb.Recordset")
			RS.Open "Select top 1 * From KS_CollectItem Where ItemID=" & ContentID,ConnItem,1,1
			IF RS.Eof And RS.Bof Then
			Call KS.AlertHistory("����ʧ��!", 1)
			 Exit Sub
			Else
			   ItemName = Trim(RS("ItemName"))
			   
			   Set RSA=Server.CreateObject("ADODB.RECORDSET")
			   RSA.Open "Select top 1 * From KS_CollectItem",ConnItem,1,3
			   RSA.AddNew
			     For I=0 To RS.Fields.count-1
				   if lcase(RS.Fields(i).name)="itemid" then
				   elseif lcase(RS.Fields(i).name)="itemname" then
				    RSA("ItemName") = GetNewTitle(RS.Fields(i).value)
				   else
				    RSA(RS.Fields(i).name) = RS.Fields(i).Value
				   end if
				 Next
			   RSA.Update
			   RSA.MoveLast
			   NewItemID=RSA("ItemID")
			   RSA.Close
			   Set RSA=Nothing
			End IF
			RS.Close
			'�����Զ����ֶ�
		    If NewItemID<>"" Then
				RS.Open "Select * from KS_FieldRules Where ItemID=" & ContentID,ConnItem,1,1
				If Not RS.Eof Then
				   Set RSA=Server.CreateObject("ADODB.RECORDSET")
				   Do While Not RS.Eof 
					   RSA.Open "Select top 1 * From KS_FieldRules where 1=0",ConnItem,1,3
					   RSA.AddNew
						 For I=0 To RS.Fields.count-1
						   if lcase(RS.Fields(i).name)="id" then
						   elseif lcase(RS.Fields(i).name)="itemid" then
							RSA("ItemID")=NewItemID
						   else
							RSA(RS.Fields(i).name) = RS.Fields(i).Value
						   end if
						 Next
					   RSA.Update
				       RSA.Close
					  RS.MoveNext
					Loop
				   Set RSA=Nothing
				End If
			 RS.Close
			 End If	
			Set RS=Nothing
		End Sub
		Function GetNewTitle(OriTitle)
			Dim RSC
			On Error Resume Next
			Set RSC = Server.CreateObject("Adodb.RecordSet")
			
				 RSC.Open "Select * From KS_CollectItem Where ItemName Like '����%" & OriTitle & "' Order By ItemID Desc", connItem, 1, 1
				 If Not RSC.EOF Then
					RSC.MoveFirst
					If RSC.RecordCount = 1 Then
					   RSC.Close
					   Set RSC = Nothing
					  GetNewTitle = "����(1) " & OriTitle
					  Exit Function
					Else
					  GetNewTitle = "����(" & CInt(Left(Split(RSC("ItemName"), "(")(1), 1)) + 1 & ") " & OriTitle
					End If
					 RSC.Close
					 Set RSC = Nothing
				 Else
				  RSC.Close
				  Set RSC = Nothing
				  GetNewTitle = "���� " & OriTitle
				  Exit Function
				 End If			  
		End Function
End Class
%> 
