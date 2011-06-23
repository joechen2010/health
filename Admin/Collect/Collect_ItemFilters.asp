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
Set KSCls = New Collect_ItemFilters
KSCls.Kesion()
Set KSCls = Nothing

Class Collect_ItemFilters
        Private KS
		Private KMCObj
		Private ConnItem,ChannelID
		Private i, totalPut, CurrentPage, SqlStr
		Private SqlItem, RSObj
		Private Action, ErrMsg,FoundErr
		Private FilterID, ItemID, FilterName, FilterObject, FilterType, Flag, PublicTf, FlagName
		
		Private FilterContent, FisString, FioString, FilterRep

		Private AllPage, iItem, ItemNum
		Private MaxPerPage
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
			If Not KS.ReturnPowerResult(0, "M010008") Then                  'Ȩ�޼��
				Call KS.ReturnErr(1, "")   
				Response.End()
			End if
		ChannelID=KS.ChkClng(KS.G("ChannelID"))
		
		If Request("page") <> "" Then
			  CurrentPage = CInt(Request("page"))
		Else
			  CurrentPage = 1
		End If
		
		Action = Request("Action")
		Select Case  Action 
		 Case  "SetFlag" 
		   Call SetFlag
		   Call Main
		 Case "Add"
		   Call FiltersAdd()
		 Case "SaveAdd"
		   Call FilterAddSave()
		 Case Else
		  Call Main
		 End Select
		End Sub
		Sub Main()
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<title>�ɼ�ϵͳ</title>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		Response.Write "<link rel=""stylesheet"" type=""text/css"" href=""../Include/Admin_Style.css"">"
		Response.Write "<script language=""JavaScript"">"
		Response.Write "var Page='" & CurrentPage & "';"
		Response.Write "</script>"
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
			InitialDocElementArr('FolderID','FilterID');
			InitialContextMenu();
			DocElementArrInitialFlag=true;
		}
		function InitialContextMenu()
		{	DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.CreateFilters('');",'��ӹ���(N)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.SelectAllElement();",'ȫ ѡ(A)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.EditFilters('');",'�� ��(E)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.DelFilters('');",'ɾ ��(D)','disabled');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
			DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.location.reload();",'ˢ ��(Z)','disabled');
		}
		function DocDisabledContextMenu()
		{
			DisabledContextMenu('FolderID','FilterID','�� ��(E),ɾ ��(D)','','','','')
		}
		function CreateFilters()
		{
			location.href='?Action=Add&ChannelID=<%=ChannelID%>';
			$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3) & KS.Setting(89)%>KS.Split.asp?OpStr=��Ϣ�ɼ����� >> �������� >> <font color=red>��ӹ���</font>&ButtonSymbol=FiltersAdd';
		}
		function EditFilters(id)
		{
		  if (id!=null&&id!='')
		  {
		  	location.href='?action=Add&channelid=<%=channelid%>&Page='+Page+'&FilterID='+id;
			$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3) & KS.Setting(89)%>KS.Split.asp?OpStr=��Ϣ�ɼ����� >> �������� >> <font color=red>�༭����</font>&ButtonSymbol=FiltersEdit';
           return ;
		  }
			GetSelectStatus('FolderID','FilterID');
		 if (SelectedFile!='')
		   if (SelectedFile.indexOf(',')==-1)
			{
			location.href='?action=Add&channelid=<%=channelid%>&Page='+Page+'&FilterID='+SelectedFile;
			$(parent.document).find('#BottomFrame')[0].src='<%=KS.Setting(3) & KS.Setting(89)%>KS.Split.asp?OpStr=��Ϣ�ɼ����� >> �������� >> <font color=red>�༭����</font>&ButtonSymbol=FiltersEdit';
			}
		   else
		   alert('һ��ֻ�ܹ��༭һ������!'); 
		 else
		  alert('��ѡ��Ҫ�༭�Ĺ���!');
		  SelectedFile='';
		}

		function DelFilters()
		{
		 GetSelectStatus('FolderID','FilterID');
		 if (SelectedFile!='')
		  {
		   if (confirm('���Ҫɾ��ѡ�еĹ�����?'))
			location="Collect_ItemFilters.asp?&channelid=<%=channelid%>&Action=SetFlag&FlagName=Del&Page="+Page+"&FilterID="+SelectedFile;
		  }
		 else
		  alert('��ѡ��Ҫɾ���Ĺ���!');
		  SelectedFile='';
		}
		function GetKeyDown()
		{ 
		if (event.ctrlKey)
		  switch  (event.keyCode)
		  {  case 90 : location.reload(); break;
			 case 65 : SelectAllElement();break;
			 case 78 : event.keyCode=0;event.returnValue=false;CreateFilters();break;
			 case 69 : event.keyCode=0;event.returnValue=false;EditFilters();break;
			 case 68 : DelFilters('');break;
		   }	
		else	
		 if (event.keyCode==46) DelFilters();
		}
		
		</script>
		<%
		Response.Write "</head>"
		Response.Write "<body scroll=no topmargin=""0"" leftmargin=""0"" onclick=""SelectElement();"" onkeydown=""GetKeyDown();"" onselectstart=""return false;"">"
		Response.Write "<ul id='menu_top'>"
		Response.Write "<li class='parent' onclick='location.href=""Collect_ItemModify.asp?channelid=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/a.gif' border='0' align='absmiddle'>�½���Ŀ</span></li>"
		Response.Write "<li class='parent' onclick='location.href=""Collect_ItemFilters.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/move.gif' border='0' align='absmiddle'>��������</span></li>"
		Response.Write "<li class='parent' onclick='location.href=""Collect_IntoDatabase.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/save.gif' border='0' align='absmiddle'>������</span></li>"
		Response.Write "<li class='parent' onclick='location.href=""Collect_ItemHistory.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/Recycl.gif' border='0' align='absmiddle'>��ʷ��¼</span></li>"
		Response.Write "<li class='parent' onclick='location.href=""Collect_Field.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/addjs.gif' border='0' align='absmiddle'>�Զ����ֶ�</span></li>"
		Response.Write "<li class='parent' onclick='location.href=""Collect_main.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/back.gif' border='0' align='absmiddle'>����һ��</span></li>"
		Response.Write ("</ul>")

		
		Response.Write "<table border=""0"" cellspacing=""0"" width=""100%"" cellpadding=""0"">"
		Response.Write "        <TBODY>"
		Response.Write "        <TR class=""sort"">"
		Response.Write "          <TD width=""23%"">��������</TD>"
		Response.Write "          <TD width=""24%"">������Ŀ</TD>"
		 Response.Write "         <TD width=""13%"">���˶���</TD>"
		 Response.Write "         <TD width=""17%"">��������</TD>"
		  Response.Write "        <TD width=""13%"">״̬</TD>"
		  Response.Write "        <TD width=""13%"">�������</TD>"
		
		If Request("page") <> "" Then
			CurrentPage = CInt(Request("Page"))
		Else
			CurrentPage = 1
		End If
		Set RSObj = Server.CreateObject("adodb.recordset")
		SqlItem = "select * From KS_Filters order by FilterID DESC"
		RSObj.Open SqlItem, ConnItem, 1, 1
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
			Else
		       Response.Write "<tr><td class='splittd' colspan=7 height='22' align='center'>û����ӹ�����</td></tr>"
			End If
		 RSObj.Close
		   Response.Write "<tr><td colspan=7><input name='chkAll' type='checkbox' id='chkAll' onclick=CheckAll(this.form) value='checkbox'>����ѡ����Ŀ <input type='submit' value='����ɾ��' class='button'>&nbsp;<input type='button' onclick=""CreateFilters();"" value='��ӹ���' class='button'></td></tr>"
		   Response.Write "</form>"
		Response.Write "</table>"
		Response.Write "</body>"
		Response.Write "</html>"
		End Sub
		
		Sub SetFlag()
		   FilterID = Trim(Request("FilterID"))
		   FlagName = Trim(Request("FlagName"))
		   If FilterID <> "" Then
			  FilterID = Replace(FilterID, " ", "")
		   Else
			  Call KS.AlertHistory("��ѡ�������Ŀ!",-1)
			  response.end
		   End If
			  Select Case FlagName
			  Case "Del"
				 SqlItem = "Delete From KS_Filters Where FilterID In(" & FilterID & ")"
			  Case "Public"
				 SqlItem = "Update KS_Filters set PublicTf=Not PublicTf Where FilterID In(" & FilterID & ")"
			  Case "Passed"
				 SqlItem = "Update KS_Filters set Flag=Not Flag Where FilterID In(" & FilterID & ")"
			  End Select
			  ConnItem.Execute (SqlItem)
		End Sub
		Sub showContent()
		 iItem = 0
		   Response.Write "<form name='myform' method='Post' action='?Action=SetFlag&FlagName=Del&Page=" & CurrentPage & "&channelid=" & channelid & "'>"
		   Do While Not RSObj.EOF
			  FilterID = RSObj("FilterID")
			  ItemID = RSObj("ItemID")
			  FilterName = RSObj("FilterName")
			  FilterObject = RSObj("FilterObject")
			  FilterType = RSObj("FilterType")
			  Flag = RSObj("Flag")
			  PublicTf = RSObj("PublicTf")
		
		   Response.Write "     <TR>"
		   Response.Write "       <TD class='splittd' ondblclick='EditFilters(" &FilterID& ");'>"
				
				   Response.Write "<input type='checkbox' name='FilterID' value='" &FilterID & "'><span  FilterID='" & FilterID & "'><img src='../Images/Filter.gif'  align='absmiddle'>"
				   Response.Write "  <span style='cursor:default;'>" & FilterName & "</span></span>"
				  
				 Response.Write " </TD>"
				 Response.Write " <TD class='splittd' align=""center"">" & KMCObj.Collect_ShowItem_Name(ItemID, ConnItem)
				 Response.Write " </TD>"
				 Response.Write " <TD class='splittd' align=""center"">"
				  
				  If FilterObject = 1 Then
					 Response.Write "�������"
				  ElseIf FilterObject = 2 Then
					 Response.Write "���Ĺ���"
				  Else
					 Response.Write "<font color=red>û��ѡ��</font>"
				  End If
				  
				 Response.Write " </TD>"
				 Response.Write " <TD class='splittd' align=""center"">"
				  
				  If FilterType = 1 Then
					 Response.Write "���滻"
				  ElseIf FilterType = 2 Then
					 Response.Write "�߼�����"
				  Else
					 Response.Write "<font color=red>û��ѡ��</font>"
				  End If
				 
				 Response.Write " </TD>"
				 Response.Write " <TD class='splittd' align=""center"">"
				 
					If Flag = False Then
						Response.Write "<span style=""color:red;cursor:pointer"" onclick=""location.href='Collect_ItemFilters.asp?Action=SetFlag&FlagName=Passed&Page=" & CurrentPage & "&FilterID=" & FilterID & "';"">����</span>"
					Else
					  Response.Write "<span style=""cursor:pointer"" onclick=""location.href='Collect_ItemFilters.asp?Action=SetFlag&FlagName=Passed&Page=" & CurrentPage & "&FilterID=" & FilterID & "';"">����</span>"
					End If
					
					 Response.Write ("&nbsp;")
		
					If PublicTf = True Then
					   Response.Write "<span style=""color:red;cursor:pointer"" onclick=""location.href='Collect_ItemFilters.asp?Action=SetFlag&FlagName=Public&Page=" & CurrentPage & "&FilterID=" & FilterID & "';"">����</span>"
					Else
					   Response.Write "<span style=""cursor:pointer"" onclick=""location.href='Collect_ItemFilters.asp?Action=SetFlag&FlagName=Public&Page=" & CurrentPage & "&FilterID=" & FilterID & "';"">˽��</span>"
					End If
				  
			   Response.Write "   </TD>"
			   Response.Write "<td class='splittd' align=center><a href='#' onclick='EditFilters(" &FilterID& ");'>�༭</a> <a href='?Action=SetFlag&FlagName=Del&Page=" & CurrentPage & "&FilterID=" & FilterID & "&channelid=" & channelid & "' onclick='return(confirm(""ȷ��ɾ����""))'>ɾ��</a></td>"
			   Response.Write " </TR>"
			iItem = iItem + 1
			  If iItem >= MaxPerPage Then Exit Do
			  RSObj.MoveNext
		   Loop
		   Response.Write "<tr><td colspan=5 align=right>"
			Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
			Response.Write "</td></tr>"
		End Sub
		
		'��ӹ�����
		Sub FiltersAdd()
		
		FilterID = Trim(Request("FilterID"))
		ItemID = 0
		FilterObject = 1
		FilterType = 1
		PublicTf = False
		Flag = True
		If FilterID <> "" Then
			 Dim RSObj
			  Set RSObj = Server.CreateObject("Adodb.Recordset")
			  RSObj.Open "Select * From KS_Filters Where FilterID=" & FilterID, ConnItem, 1, 1
			  If Not RSObj.EOF Then
				 FilterName = RSObj("FilterName")
				 ItemID = RSObj("ItemID")
				 FilterObject = RSObj("FilterObject")
				 FilterType = RSObj("FilterType")
				 FilterContent = RSObj("FilterContent")
				 FisString = RSObj("FisString")
				 FioString = RSObj("FioString")
				 FilterRep = RSObj("FilterRep")
				 Flag = RSObj("Flag")
				 PublicTf = RSObj("PublicTf")
			  End If
			  RSObj.Close
			  Set RSObj = Nothing
		End If
		If KS.IsNul(FilterContent) Then FilterContent=""
		If KS.IsNul(FisString) Then FisString=""
		If KS.IsNUL(FioString) Then FioString=""
		If KS.IsNul(FilterRep) Then FilterRep=""
		
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<title>�ɼ�ϵͳ</title>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		Response.Write "<link rel=""stylesheet"" type=""text/css"" href=""../Include/Admin_Style.css"">"
		Response.Write "</head>"
		Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
		Response.Write "<ul id='menu_top'>"
		Response.Write "<li class='parent' onclick='CreateCollectItem();'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/a.gif' border='0' align='absmiddle'>�½���Ŀ</span></li>"
		Response.Write "<li class='parent' onclick='location.href=""Collect_ItemFilters.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/move.gif' border='0' align='absmiddle'>��������</span></li>"
		Response.Write "<li class='parent' onclick='location.href=""Collect_IntoDatabase.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/save.gif' border='0' align='absmiddle'>������</span></li>"
		Response.Write "<li class='parent' onclick='location.href=""Collect_IntoDatabase.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/Recycl.gif' border='0' align='absmiddle'>��ʷ��¼</span></li>"
		Response.Write "<li class='parent' onclick='location.href=""Collect_Field.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/addjs.gif' border='0' align='absmiddle'>�Զ����ֶ�</span></li>"
		Response.Write "<li class='parent' onclick='location.href=""Collect_main.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/back.gif' border='0' align='absmiddle'>����һ��</span></li>"
		Response.Write ("</ul>")
										 
		Response.Write "<form method=""post"" action=""?"" name=""form1"">"
		Response.Write "<br>"
		Response.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""1"" class=""border"" >"
		Response.Write "    <tr class='tdbg'>"
		Response.Write "      <td width=""100"" height=""25"" align=""center"" class='clefttitle'> �������ƣ�</td>"
		Response.Write "      <td><input name=""FilterName"" type=""text"" id=""FilterName"" value=""" & FilterName & """ size=""25"" maxlength=""30"">"
		Response.Write "        &nbsp;</td>"
		Response.Write "    </tr>"
		Response.Write "    <tr class='tdbg'>"
		Response.Write "      <td width=""100"" height=""25"" align=""center"" class='clefttitle'> ������Ŀ��</td>"
		Response.Write "      <td>" & KMCObj.Collect_ShowItem_Option(ItemID, ConnItem) & "      </td>"
		Response.Write "    </tr>"
		Response.Write "    <tr class='tdbg'>"
		 Response.Write "     <td width=""100"" height=""25"" align=""center"" class='clefttitle'> ���˶���</td>"
		 Response.Write "     <td>"
		 Response.Write "        <select name=""FilterObject"" id=""FilterObject"">"
				   
				   If FilterObject = 1 Then
					 Response.Write "<option value=""1"" selected>�������</option>"
					 Response.Write "<option value=""2"">���Ĺ���</option>"
					Else
					 Response.Write "<option value=""1"">�������</option>"
					 Response.Write "<option value=""2"" selected>���Ĺ���</option>"
					End If
					
		  Response.Write "       </select>      </td>"
		  Response.Write "  </tr>"
		  Response.Write "  <tr class='tdbg'>"
		  Response.Write "    <td width=""100"" height=""25"" align=""center"" class='clefttitle'> �������ͣ�</td>"
		  Response.Write "    <td>"
				 
		   Response.Write "      <select name=""FilterType"" id=""FilterType"" onchange=showset(this.value)>"
				   If FilterType = 1 Then
					 Response.Write "<option value=""1"" selected >���滻</option>"
					 Response.Write "<option value=""2"">�߼�����</option>"
					 Else
					 Response.Write "<option value=""1"">���滻</option>"
					 Response.Write "<option value=""2"" selected >�߼�����</option>"
					 End If
					 
		  Response.Write "       </select>      </td>"
		  Response.Write "  </tr>"
		  Response.Write "  <tr class='tdbg'>"
		   Response.Write "   <td width=""100"" height=""25"" align=""center"" class='clefttitle'> ʹ��״̬��</td>"
		   Response.Write "   <td>"
				 If Flag = True Then
				  Response.Write "<input type=""radio"" name=""Flag"" value=""yes"" checked>����"
				  Response.Write "<input type=""radio"" name=""Flag"" value=""no"">����"
				  Else
				  Response.Write "<input type=""radio"" name=""Flag"" value=""yes"">����"
				  Response.Write "<input type=""radio"" name=""Flag"" value=""no"" checked>����"
				  End If
				  
		  Response.Write "    </td>"
		 Response.Write "   </tr>"
		 Response.Write "   <tr class='tdbg'>"
		 Response.Write "     <td width=""100"" height=""25"" align=""center"" class='clefttitle'> ʹ�÷�Χ��</td>"
		
		 Response.Write "     <td>"
				 If PublicTf = False Then
					Response.Write "<input type=""radio"" name=""PublicTf"" value=""no"" checked>˽��"
					Response.Write "<input type=""radio"" name=""PublicTf"" value=""yes"">����"
					Else
					Response.Write "<input type=""radio"" name=""PublicTf"" value=""no"">˽��"
					Response.Write "<input type=""radio"" name=""PublicTf"" value=""yes"" checked>����"
					End If
				   
		Response.Write "      </td>"
		Response.Write "    </tr>"
		Response.Write "    <tr class='tdbg' id=""FilterType1"" style=""display:"">"
		 Response.Write "     <td width=""100"" align=""center"" class='clefttitle'> ���ݣ�</td>"
		 Response.Write "     <td ><textarea name=""FilterContent"" cols=""49"" rows=""5"">" & server.HTMLEncode(FilterContent) & "</textarea>"
		 Response.Write "     </td>"
		Response.Write "    </tr>"
		 Response.Write "   <tr class='tdbg'  id=""FilterType2"" style=""display:none"">"
		 Response.Write "     <td width=""100"" align=""center"" class='clefttitle'> ��ʼ��ǣ�</td>"
		  Response.Write "    <td><textarea name=""FisString"" cols=""49"" rows=""5"">" & server.HTMLEncode(FisString) & "</textarea>"
				Response.Write "&nbsp;</td>"
			Response.Write "</tr>"
			Response.Write "<tr class='tdbg'  id=""FilterType22"" style=""display:none"">"
			 Response.Write " <td width=""100"" align=""center"" class='clefttitle'> ������ǣ�</td>"
			 Response.Write " <td><textarea name=""FioString"" cols=""49"" rows=""5"">" & server.HTMLEncode(FioString) & "</textarea>"
			 Response.Write "   &nbsp;</td>"
			Response.Write "</tr>"
		  Response.Write "  <tr  class='tdbg' id=""FilterRep"">"
			Response.Write "  <td width=""100"" align=""center"" class='clefttitle'> �滻��</td>"
			Response.Write "  <td><textarea name=""FilterRep"" cols=""49"" rows=""5"">" & server.HTMLEncode(FilterRep) & "</textarea>"
			Response.Write "  &nbsp;</td>"
		   Response.Write " </tr>"
		  Response.Write "  <tr class='tdbg'>"
		   Response.Write "   <td colspan=""2"" align=""center"">"
			 Response.Write "  <input type=""hidden"" value=""" & FilterID & """ name=""FilterID"">"
			 Response.Write "  <input type=""hidden"" value=""" & Request("Page") & """ name=""Page"">"
			 Response.Write "  <input type=""hidden"" value=""" & Request("Channelid") & """ name=""channelid"">"
			 Response.Write "  <input name=""Action"" type=""hidden"" id=""Action"" value=""SaveAdd""></td>"
			Response.Write "</tr>"
		Response.Write "</table>"
		Response.Write "</form>"
		Response.Write "</body>"
		Response.Write "</html>"
		Response.Write "<SCRIPT language=javascript>"
		Response.Write "showset(" & FilterType & ");"
		Response.Write "function showset(num)"
		Response.Write "{"
				Response.Write "if(num!=1)"
				Response.Write "{"
					Response.Write "document.all.FilterType1.style.display = ""none"";"
					Response.Write "document.all.FilterType2.style.display = """";"
					Response.Write "document.all.FilterType22.style.display = """";"
		Response.Write "        }"
		Response.Write "        else"
		Response.Write "        {"
		Response.Write "            document.all.FilterType1.style.display = """";"
		Response.Write "            document.all.FilterType2.style.display = ""none"";"
		Response.Write "            document.all.FilterType22.style.display = ""none"";"
		Response.Write "        }"
		
		Response.Write "}"
		Response.Write "function CheckForm()"
		Response.Write "{"
		 Response.Write " var myform=document.form1;"
		 Response.Write " if (myform.FilterName.value=='')"
		 Response.Write " {"
		 Response.Write "    alert('�������������');"
		 Response.Write "    myform.FilterName.focus();"
		 Response.Write "    return false;"
		 Response.Write " }"
		 Response.Write " if (myform.ItemID.value=='')"
		Response.Write "  {"
		Response.Write "     alert('��ѡ��һ����Ŀ');"
		 Response.Write "    myform.ItemID.focus();"
		Response.Write "     return false;"
		Response.Write "  }"
		Response.Write "   myform.submit();"
		Response.Write "  return true;"
		Response.Write "}"
		Response.Write "</script>"
	  End Sub
	  
	  Sub FilterAddSave()
		Dim SqlItem, RsItem
		FilterName = Trim(Request.Form("FilterName"))
		ItemID = Trim(Request.Form("ItemID"))
		FilterObject = Request.Form("FilterObject")
		FilterType = Request.Form("FilterType")
		FilterContent = Request.Form("FilterContent")
		FisString = Request.Form("FisString")
		FioString = Request.Form("FioString")
		FilterRep = Request.Form("FilterRep")
		Flag = Request.Form("Flag")
		PublicTf = Request.Form("PublicTf")
		FilterID=Request.Form("FilterID")
		
		If FilterName = "" Then
		   FoundErr = True
		   ErrMsg = ErrMsg & "�������Ʋ���Ϊ��"
		End If
		If ItemID = "" Then
		   FoundErr = True
		   ErrMsg = ErrMsg & "��ѡ�����������Ŀ"
		Else
		   ItemID = CLng(ItemID)
		   If ItemID = 0 Then
			  FoundErr = True
			  ErrMsg = ErrMsg & "��ѡ�����������Ŀ"
		   End If
		End If
		If FilterObject = "" Then
		   FoundErr = True
		   ErrMsg = ErrMsg & "��ѡ����˶���"
		Else
		   FilterObject = CLng(FilterObject)
		End If
									
		If FilterType = "" Then
		   FoundErr = True
		   ErrMsg = ErrMsg & "��ѡ���������"
		Else
		   FilterType = CLng(FilterType)
		   If FilterType = 1 Then
			  If FilterContent = "" Then
				 FoundErr = True
				 ErrMsg = ErrMsg & "���˵����ݲ���Ϊ��"
			  End If
		   ElseIf FilterType = 2 Then
			  If FisString = "" Or FioString = "" Then
				 FoundErr = True
				 ErrMsg = ErrMsg & "��ʼ/������ǲ���Ϊ��"
			  End If
		   Else
			  FoundErr = True
			  ErrMsg = ErrMsg & "�������������Ч���ӽ���"
		   End If
		End If
		If Flag = "yes" Then
		   Flag = True
		Else
		   Flag = False
		End If
		If PublicTf = "yes" Then
		   PublicTf = True
		Else
		   PublicTf = False
		End If
										
		If FoundErr <> True Then
		   Set RsItem = Server.CreateObject("adodb.recordset")
		   If FilterID <> "" Then
			 SqlItem = "select top 1 *  From KS_Filters Where FilterID=" & FilterID
			 RsItem.Open SqlItem, ConnItem, 1, 3
			 RsItem("FilterName") = FilterName
			 RsItem("ItemID") = ItemID
			 RsItem("FilterObject") = FilterObject
			 RsItem("FilterType") = FilterType
			 If FilterType = 1 Then
			   RsItem("FilterContent") = FilterContent
			 ElseIf FilterType = 2 Then
			   RsItem("FisString") = FisString
			   RsItem("FioString") = FioString
			 End If
			 RsItem("FilterRep") = FilterRep
			 RsItem("Flag") = Flag
			 RsItem("PublicTf") = PublicTf
			 RsItem.Update
			 RsItem.Close
			 Response.Write ("<script>alert('�����޸ĳɹ�!');location.href='Collect_ItemFilters.asp?channelid=" & channelid & "&Page=" & Request("Page") & "';$(parent.document).find('#BottomFrame')[0].src='../KS.Split.asp?OpStr=��Ϣ�ɼ����� >> <font color=red>��������</font>&ButtonSymbol=Disabled'</script>")
		   Else
		   SqlItem = "select top 1 *  From KS_Filters"
		   RsItem.Open SqlItem, ConnItem, 1, 3
		   RsItem.AddNew
		   RsItem("FilterName") = FilterName
		   RsItem("ItemID") = ItemID
		   RsItem("FilterObject") = FilterObject
		   RsItem("FilterType") = FilterType
		   If FilterType = 1 Then
			  RsItem("FilterContent") = FilterContent
		   ElseIf FilterType = 2 Then
			  RsItem("FisString") = FisString
			  RsItem("FioString") = FioString
		   End If
		   RsItem("FilterRep") = FilterRep
		   RsItem("Flag") = Flag
		   RsItem("PublicTf") = PublicTf
		   RsItem.Update
		   RsItem.Close
		   Response.Write ("<script>if (confirm('������ӳɹ������������?')){location.href='?action=Add&ChannelID=" & ChannelID& "';}else{location.href='Collect_ItemFilters.asp?channelid=" & channelid & "';$(parent.document).find('#BottomFrame')[0].src='../KS.Split.asp?OpStr=��Ϣ�ɼ����� >> <font color=red>��������</font>&ButtonSymbol=Disabled'}</script>")
		   End If
		   Set RsItem = Nothing
		Else
		   Call KS.AlertHistory(ErrMsg,-1)
		End If
	  End Sub
End Class
%> 
