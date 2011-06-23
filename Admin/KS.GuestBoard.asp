<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.FunctionCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New GuestBoard_Main
KSCls.Kesion()
Set KSCls = Nothing

Class GuestBoard_Main
        Private KS,Action
		Private I, totalPut, CurrentPage, SqlStr, RSObj
        Private MaxPerPage
		Private Sub Class_Initialize()
		  MaxPerPage = 20
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub


		Public Sub Kesion()
			If Not KS.ReturnPowerResult(0, "KSMS20004") Then
			  Call KS.ReturnErr(1, "")
			  exit sub
			End If
			Action=KS.G("Action")
						If Request("page") <> "" Then
				  CurrentPage = CInt(Request("page"))
			Else
				  CurrentPage = 1
			End If
			With Response
			.Write "<html>"
			.Write "<head>"
			.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "<script language=""JavaScript"">" & vbCrLf
			.Write "var Page='" & CurrentPage & "';" & vbCrLf
			.Write "</script>" & vbCrLf
			.Write "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>"
			.Write "<script language=""JavaScript"" src=""../KS_Inc/jquery.js""></script>"
			.Write "<script language=""JavaScript"" src=""Include/ContextMenu1.js""></script>"
			.Write "<script language=""JavaScript"" src=""Include/SelectElement.js""></script>"

			Select Case Action
			 Case "Add","Edit"
				  Call GuestBoardAddOrEdit()
			 Case "Save"
			      Call GuestBoardSave()
			 Case "Del"
			      Call GuestBoardDel()
			 Case Else
			   Call MainList()
			End Select
		  End With
	    End Sub
		
		Sub MainList()
		 With Response
			%>
			<script language="JavaScript">
			var DocElementArrInitialFlag=false;
			var DocElementArr = new Array();
			var DocMenuArr=new Array();
			var SelectedFile='',SelectedFolder='';
			function document.onreadystatechange()
			{   if (DocElementArrInitialFlag) return;
				InitialDocElementArr('FolderID','GuestBoardID');
				InitialContextMenu();
				DocElementArrInitialFlag=true;
			}
			function InitialContextMenu()
			{	DocMenuArr[DocMenuArr.length]=new ContextMenuItem("window.parent.GuestBoardAdd(0);",'�� ��(N)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.SelectAllElement();",'ȫ ѡ(A)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.GuestBoardControl(1);",'�� ��(E)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.GuestBoardControl(2);",'ɾ ��(D)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.location.reload();",'ˢ ��(Z)','disabled');
			}
			function DocDisabledContextMenu()
			{
				DisabledContextMenu('FolderID','GuestBoardID','�� ��(E),ɾ ��(D)','�� ��(E)','','','','')
			}
			function GuestBoardAdd(parentid)
			{
				location.href='KS.GuestBoard.asp?Action=Add&parentid='+parentid;
				window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=����������� >> <font color=red>����°���</font>&ButtonSymbol=GO';
			}
			function EditGuestBoard(id)
			{
				location="KS.GuestBoard.asp?Action=Edit&Page="+Page+"&Flag=Edit&GuestBoardID="+id;
				window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=����������� >> <font color=red>�༭����</font>&ButtonSymbol=GoSave';
			}
			function DelGuestBoard(id)
			{
			if (confirm('������Ӱ��潫ͬʱ��ɾ��,���Ҫִ��ɾ��������?'))
			 location="KS.GuestBoard.asp?Action=Del&Page="+Page+"&GuestBoardid="+id;
			   SelectedFile='';
			}
			function GuestBoardControl(op)
			{  var alertmsg='';
				GetSelectStatus('FolderID','GuestBoardID');
				if (SelectedFile!='')
				 {  if (op==1)
					{
					if (SelectedFile.indexOf(',')==-1) 
						EditGuestBoard(SelectedFile)
					  else alert('һ��ֻ�ܱ༭һ������!')	
					SelectedFile='';
					}
				  else if (op==2)    
					 DelGuestBoard(SelectedFile);
				 }
				else 
				 {
				 if (op==1)
				  alertmsg="�༭";
				 else if(op==2)
				  alertmsg="ɾ��"; 
				 else
				  {
				  WindowReload();
				  alertmsg="����" 
				  }
				 alert('��ѡ��Ҫ'+alertmsg+'�İ���');
				  }
			}
			function GetKeyDown()
			{ 
			if (event.ctrlKey)
			  switch  (event.keyCode)
			  {  case 90 : location.reload(); break;
				 case 65 : SelectAllElement();break;
				 case 78 : event.keyCode=0;event.returnValue=false; GuestBoardAdd(0);break;
				 case 69 : event.keyCode=0;event.returnValue=false;GuestBoardControl(1);break;
				 case 68 : GuestBoardControl(2);break;
			   }	
			else	
			 if (event.keyCode==46)GuestBoardControl(2);
			}
			</script>
			<%
			.Write "</head>"
			.Write "<body topmargin=""0"" leftmargin=""0"" onclick=""SelectElement();"" onkeydown=""GetKeyDown();"" onselectstart=""return false;"">"
			  .Write "<ul id='menu_top'>"
			  .Write "<li class='parent' onclick=""GuestBoardAdd(0);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>��Ӱ���</span></li>"
			  .Write "<li class='parent' onclick=""GuestBoardControl(1);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>�༭����</span></li>"
			  .Write "<li class='parent' onclick=""GuestBoardControl(2);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>ɾ������</span></li>"
			  .Write "</ul>"
			

			.Write "<table width=""100%""  border=""0"" cellpadding=""0"" cellspacing=""0"">"
			.Write "  <tr>"			
			.Write "          <td height=""25"" class=""sort"">"
			.Write "          <div align=""center"">��������</div></td>"
			.Write "          <td class=""sort""><div align=""center"">����</div></td>"
			.Write "          <td align=""center"" class=""sort"">������</td>"
			.Write "          <td width=""50"" class=""sort"" align=""center"">����</td>"
			.Write "          <td  class=""sort"" align=""center"">�������</td>"
			.Write "  </tr>"
			 
			 Set RSObj = Server.CreateObject("ADODB.RecordSet")
					   SqlStr = "SELECT * FROM KS_GuestBoard Where ParentID=0 order by orderID desc"
					   RSObj.Open SqlStr, Conn, 1, 1
					 If RSObj.EOF And RSObj.BOF Then
					 Else
						        totalPut = RSObj.RecordCount
								If CurrentPage < 1 Then	CurrentPage = 1
			
								
								   If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrentPage - 1) * MaxPerPage
									Else
										CurrentPage = 1
									End If
									Call showContent
				End If
				
			.Write "    </td>"
			.Write "  </tr>"
			.Write "</table>"
			.Write "</body>"
			.Write "</html>"
			End With
			End Sub
			Sub showContent()
			  Dim RS,I
			  With Response
					Do While Not RSObj.EOF
					  .Write "<tr>"
					  .Write "  <td class='splittd' width='44%' height='20'> &nbsp;&nbsp; <span GuestBoardID='" & RSObj("ID") & "' ondblclick=""EditGuestBoard(this.GuestBoardID)""><img src='Images/Field.gif' align='absmiddle'>"
					  .Write "    <span style='cursor:default;'>" & KS.GotTopic(RSObj("BoardName"), 45) & "</span></span> "
					  .Write "  </td>"
					  .Write "  <td class='splittd' align='center'>&nbsp;" & RSObj("master") & "&nbsp;</td>"
					  .Write "  <td class='splittd' align='center'>����:<font Color=red>" & RSObj("todaynum") & "</font> ����:<font Color=red>" & RSObj("topicnum") & "</font> ����:<font Color=red>" & RSObj("postnum") & "</font></td>"
					  .Write "  <td class='splittd' align='center'>" & RSOBJ("OrderID") &"</td>"
					  .Write "  <td class='splittd'> <a href='javascript:GuestBoardAdd(" & rsobj("id") & ")'>��ӷְ���</a> | <a href='javascript:EditGuestBoard(" & rsobj("id") & ")'>�޸�</a> | <a href='javascript:DelGuestBoard(" & rsobj("id") & ")'>ɾ��</a> </td>"
					  .Write "</tr>"
					  Set RS=Conn.Execute("Select ID,BoardName,master,todaynum,postnum,topicnum,orderid From KS_GuestBoard Where ParentID=" & RSObj("ID") & " Order by orderid")
					  Do While not rs.eof
					  .Write "<tr>"
					  .Write "  <td class='splittd' width='44%' height='20'> &nbsp;&nbsp; &nbsp;&nbsp; &nbsp;&nbsp;|- <span GuestBoardID='" & RS("ID") & "' ondblclick=""EditGuestBoard(this.GuestBoardID)""><img src='Images/folder/folderopen.gif' align='absmiddle'>"
					  .Write "    <span style='cursor:default;'>" & KS.GotTopic(RS("BoardName"), 45) & "</span></span> "
					  .Write "  </td>"
					  .Write "  <td class='splittd' align='center'>&nbsp;" & RS("master") & "&nbsp;</td>"
					  .Write "  <td class='splittd' align='center'>����:<font Color=red>" & RS("todaynum") & "</font> ����:<font Color=red>" & RS("topicnum") & "</font> ����:<font Color=red>" & RS("postnum") & "</font></td>"
					  .Write "  <td class='splittd' align='center'>" & RS("OrderID") &"</td>"
					  .Write "  <td class='splittd'> <a href='#' disabled>��ӷְ���</a> | <a href='javascript:EditGuestBoard(" & rs("id") & ")'>�޸�</a> | <a href='javascript:DelGuestBoard(" & rs("id") & ")'>ɾ��</a> </td>"
					  .Write "</tr>"
					  rs.movenext
					  loop
					  rs.close
					  
					 I = I + 1
					  If I >= MaxPerPage Then Exit Do
						   RSObj.MoveNext
					Loop
					  RSObj.Close
					  .Write "<tr><td height='26' colspan='5' align='right'>"
					  Call KS.ShowPageParamter(totalPut, MaxPerPage, "", True, "��", CurrentPage, "Action=" & Action)
				End With
			    Set RS=Nothing
			End Sub
			
			'����޸İ���
		  Sub GuestBoardAddOrEdit()
		  		Dim GuestBoardID, RSObj, SqlStr, Content, BoardName, Note, Master, AddDate,Flag, Page,OrderID,ParentID,BoardRules,Settings,SetArr
				Flag = KS.G("Flag")
				Page = KS.G("Page")
				If Page = "" Then Page = 1
				If Flag = "Edit" Then
					GuestBoardID = KS.G("GuestBoardID")
					Set RSObj = Server.CreateObject("Adodb.Recordset")
					SqlStr = "SELECT * FROM KS_GuestBoard Where ID=" & GuestBoardID
					RSObj.Open SqlStr, Conn, 1, 1
					  BoardName     = RSObj("BoardName")
					  Note    = RSObj("Note")
					  AddDate   = RSObj("AddDate")
					  Master  = RSObj("Master")
					  ParentID= RSObj("ParentID")
					  OrderID = RSObj("OrderID")
					  BoardRules=RSObj("BoardRules")
					  Settings=RSObj("Settings")
					RSObj.Close:Set RSObj = Nothing
				Else
				   Flag = "Add"
				   ParentID=Request("Parentid")
				End If
				If KS.IsNul(Settings) Then
				Settings="0$0$0$1$1$1$1$1$1$1$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$"
				End If
				SetArr=Split(Settings,"$")
				
				With Response
				.Write "<html>"
				.Write "<head>"
				.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
				.Write "<title>�������</title>"
				.Write "</head>"
				.Write "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
				.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
				.Write "<script src=""images/pannel/tabpane.js"" language=""JavaScript""></script>"
		        .Write "<link href=""images/pannel/tabpane.CSS"" rel=""stylesheet"" type=""text/css"">"
				.Write "<body bgcolor=""#FFFFFF"" leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
				.Write " <div class='topdashed sort'>"
				If Flag = "Edit" Then
				 .Write "�޸İ���"
				Else
				 .Write "��Ӱ���"
				End If
	            .Write "</div>"
				.Write "<br>"
				
				.write "<div class=tab-page id=boardpanel>"
				.Write "  <form name=GuestBoardForm method=post action=""?Action=Save"">"
				.Write " <SCRIPT type=text/javascript>"& _
				"   var tabPane1 = new WebFXTabPane( document.getElementById( ""boardpanel"" ), 1 )"& _
				" </SCRIPT>"& _
					 
				" <div class=tab-page id=basic-page>"& _
				"  <H2 class=tab>������Ϣ</H2>"& _
				"	<SCRIPT type=text/javascript>"& _
				"				 tabPane1.addTabPage( document.getElementById( ""basic-page"" ) );"& _
				"	</SCRIPT>" 
				
				.Write "<table width=""100%"" border=""0"" cellpadding=""1"" cellspacing=""1"" class='ctable'>"
				.Write "   <input type=""hidden"" name=""Flag"" value=""" & Flag & """>"
				.Write "   <input type=""hidden"" name=""GuestBoardID"" value=""" & GuestBoardID & """>"
				.Write "   <input type=""hidden"" name=""Page"" value=""" & Page & """>"
				.Write "     <tr  class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>�� �� ��:</strong></td>"
				.Write "             <td>"
				.Write "             <select name='parentid'>"
				.Write "               <option value=0>-��Ϊ������-</option>"
				   Dim RST:Set RST=Conn.Execute("Select ID,BoardName From KS_GuestBoard Where ParentID=0 order by orderid")
				   Do While Not RST.Eof
				     If trim(ParentID)=trim(RST(0)) Then
				     .Write "<option value='" & RST(0) & "' selected>" & RST(1) & "</option>"
					 Else
				     .Write "<option value='" & RST(0) & "'>" & RST(1) & "</option>"
					 End If
				   RST.MoveNext
				   Loop
				   RST.Close
				   Set RST=Nothing
				.Write "             </select>"           
				.Write "              </td>"
				.Write "          </tr>"
				
				
				.Write "          <tr  class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>��������:</strong></td>"
				.Write "             <td>"
				.Write "              <input name=""BoardName"" type=""text"" id=""BoardName"" value=""" & BoardName & """ class=""textbox"" style=""width:60%""> �磬����������������ѯ��</td>"
				 .Write "</tr>"
				 .Write "<tr class='tdbg'>"
				.Write "  <td height=""25"" align='right' width='125' class='clefttitle'><strong>�������:</strong></td>"
				.Write "  <td>"
				.Write "<textarea name=""Note"" cols='75' rows='6' class=""textbox"" style=""height:80px;width:70%"">" & Note &"</textarea>"
				.Write "            </td>"
				.Write "          </tr>"
				 .Write "<tr class='tdbg'>"
				.Write "  <td height=""25"" align='right' width='125' class='clefttitle'><strong>�� ��:</strong><br/><font color=blue>֧��Html�﷨</font></td>"
				.Write "  <td>"
				.Write "<textarea name=""BoardRules"" cols='75' rows='6' class=""textbox"" style=""height:180px;width:70%"">" & BoardRules &"</textarea>"
				.Write "            </td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>�������:</strong></td>"
				.Write "            <td>"
				.Write "              <input name=""Master"" type=""text"" id=""Master"" value=""" & Master &""" class=""textbox"" style=""width:50%""> �����������Ӣ�Ķ��Ÿ���"
				.Write "              </td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>�� �� ��:</strong></td>"
				.Write "            <td>"
				.Write "              <input name=""OrderID"" type=""text"" value=""" & OrderID &""" class=""textbox""> ���ԽС������Խǰ��"
				.Write "              </td>"
				.Write "          </tr>"

				.Write "</table>"
				.Write "</div>"
				.Write "<div class=tab-page id=""formset"">"
		        .Write " <H2 class=tab>Ȩ�޻���</H2>"
			    .Write "<SCRIPT type=text/javascript>"
				.Write " tabPane1.addTabPage( document.getElementById( ""formset"" ) );"
			    .Write "</SCRIPT>"
				.Write "<table width=""100%"" border=""0"" cellpadding=""1"" cellspacing=""1"" class='ctable'>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>�Ƿ������ο�����鿴:</strong></td>"
				.Write "            <td>"
				.write "<input type=""radio"" name=""setarr(0)"" value=""1"" "
				If KS.ChkClng(SetArr(0)) = 1 Then .Write (" checked")
				.Write ">"
				.Write "��"
				.Write "  <input type=""radio"" name=""setarr(0)"" value=""0"" "
				If KS.ChkClng(SetArr(0)) = 0 Then .Write (" checked")
				.Write ">"
				.Write "��"
				.Write "              </td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>��������˰���Ļ�Ա��:</strong><br/><font color=blue>�������벻Ҫ��ѡ</font></td>"
				.Write "            <td>"
				.Write KS.GetUserGroup_CheckBox("SetArr(1)",SetArr(1),5)
				.Write "              </td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>�����ڴ˰��淢��Ļ�Ա��:</strong><br/><font color=blue>�������벻Ҫ��ѡ</font></td>"
				.Write "            <td>"
				.Write KS.GetUserGroup_CheckBox("SetArr(2)",SetArr(2),5)
				.Write "              </td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>�ڴ˰��淢���ɵõ�:</strong></td>"
				.Write "            <td><input type='text' style='text-align:center' name='setarr(3)' size=5 value='" & setarr(3) & "'>������</td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>�ڴ˰�������ɵõ�:</strong></td>"
				.Write "            <td><input type='text' style='text-align:center' name='setarr(4)' size=5 value='" & setarr(4) & "'>������</td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>���ӱ��ö��ɵõ�:</strong></td>"
				.Write "            <td><input type='text' style='text-align:center' name='setarr(5)' size=5 value='" & setarr(5) & "'>������</td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>���ӱ���Ϊ�����ɵõ�:</strong></td>"
				.Write "            <td><input type='text' style='text-align:center' name='setarr(6)' size=5 value='" & setarr(6) & "'>������</td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>���ӱ�ɾ�����۳�:</strong></td>"
				.Write "            <td><input type='text' style='text-align:center' name='setarr(7)' size=5 value='" & setarr(7) & "'>������</td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>�ظ���ɾ�����۳�:</strong></td>"
				.Write "            <td><input type='text' style='text-align:center' name='setarr(8)' size=5 value='" & setarr(8) & "'>������</td>"
				.Write "          </tr>"
				.Write "          <tr class='tdbg' style='color:blue'>"
				.Write "            <td height=""25"" align='right' width='125' class='clefttitle'><strong>��ע���û�:</strong></td>"
				.Write "            <td><input type='text' style='text-align:center' name='setarr(9)' size=5 value='" & setarr(9) & "'> ���Ӻ�ſ����ڱ����淢������</td>"
				.Write "          </tr>"
                .Write "</table>"
				.Write "</div>"
                 				
				.Write "  </form>"
				.Write "</body>"
				.Write "</html>"
				.Write "<script language=""JavaScript"">" & vbCrLf
				.Write "<!--" & vbCrLf
				.Write "function CheckForm()" & vbCrLf
				.Write "{ var form=document.GuestBoardForm;" & vbCrLf
				.Write "  if (form.BoardName.value=='')" & vbCrLf
				.Write "   {" & vbCrLf
				.Write "    alert('�������������!');" & vbCrLf
				.Write "    form.BoardName.focus();" & vbCrLf
				.Write "    return false;" & vbCrLf
				.Write "   }" & vbCrLf
				.Write "   if (form.Note.value=='')" & vbCrLf
				.Write "   {" & vbCrLf
				.Write "    alert('������������!');" & vbCrLf
				.Write "    form.Note.focus();" & vbCrLf
				.Write "    return false;" & vbCrLf
				.Write "   }" & vbCrLf
				.Write "      if (form.OrderID.value=='')" & vbCrLf
				.Write "   {" & vbCrLf
				.Write "    alert('������������!');" & vbCrLf
				.Write "    form.OrderID.focus();" & vbCrLf
				.Write "    return false;" & vbCrLf
				.Write "   }" & vbCrLf
				.Write "   form.submit();"
				.Write "   return true;"
				.Write "}"
				.Write "//-->"
				.Write "</script>"
			 End With
		  End Sub
		  
		  '����
		  Sub GuestBoardSave()
			Dim GuestBoardID, RSObj, SqlStr, BoardName, Note, AddDate, Content, Master,Flag, Page, RSCheck,OrderID,ParentID,BoardRules,Settings,I
			Set RSObj = Server.CreateObject("Adodb.RecordSet")
			Flag = Request.Form("Flag")
			GuestBoardID = Request("GuestBoardID")
			BoardName = Replace(Replace(Request.Form("BoardName"), """", ""), "'", "")
			Note = Replace(Replace(Request.Form("Note"), """", ""), "'", "")
			Master = Request.Form("Master")
			BoardRules=Request.Form("BoardRules")
			OrderID = KS.ChkClng(KS.G("OrderID"))
			ParentID = KS.Chkclng(Request.Form("ParentID"))
			If BoardName = "" Then Call KS.AlertHistory("�������Ʋ���Ϊ��!", -1)
			If Note = "" Then Call KS.AlertHistory("������ܲ���Ϊ��!", -1)
			
			
			For I=0 To 20
			  If I=0 Then 
			   Settings=Request("setarr(" & i & ")") &"$"
			  Else
			   Settings=Settings  & Request("setarr(" & i & ")")& "$"
			  End If
			Next
			
			Set RSObj = Server.CreateObject("Adodb.Recordset")
			If Flag = "Add" Then
			   RSObj.Open "Select ID From KS_GuestBoard Where BoardName='" & BoardName & "'", Conn, 1, 1
			   If Not RSObj.EOF Then
				  RSObj.Close
				  Set RSObj = Nothing
				  Response.Write ("<script>alert('�Բ���,�����Ѵ���!');history.back(-1);</script>")
				  Exit Sub
			   Else
				RSObj.Close
				RSObj.Open "SELECT * FROM KS_GuestBoard Where 1=0", Conn, 1, 3
				RSObj.AddNew
				  RSObj("BoardName") = BoardName
				  RSObj("Note") = Note
				  RSObj("AddDate") = Now
				  RSObj("Master") = Master
				  RSObj("OrderID") =OrderID
				  RSObj("ParentID")=ParentID
				  RSObj("lastpost")="0$" & now & "$��$$$$$"
				  RSObj("TodayNum")=0
				  RSObj("PostNum")=0
				  RSObj("TopicNum")=0
				  RSObj("BoardRules")=BoardRules
				  RSObj("Settings")=Settings
				RSObj.Update
				 RSObj.Close
			  End If
			   Set RSObj = Nothing
			   Call KS.DelCahe(KS.SiteSN & "_ClubBoard")
			   Response.Write ("<script> if (confirm('������ӳɹ�!���������?')) {location.href='KS.GuestBoard.asp?Action=Add&parentid=" & ParentID &"';}else{location.href='KS.GuestBoard.asp';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&OpStr=������� >> <font color=red>���Ա��������</font>';}</script>")
			ElseIf Flag = "Edit" Then
			  Page = Request.Form("Page")
			  RSObj.Open "Select ID FROM KS_GuestBoard Where BoardName='" & BoardName & "' And ID<>" & GuestBoardID, Conn, 1, 1
			  If Not RSObj.EOF Then
				 RSObj.Close
				 Set RSObj = Nothing
				 Response.Write ("<script>alert('�Բ���,���������Ѵ���!');history.back(-1);</script>")
				 Exit Sub
			  Else
			   RSObj.Close
			   SqlStr = "SELECT * FROM KS_GuestBoard Where ID=" & GuestBoardID
			   RSObj.Open SqlStr, Conn, 1, 3
				 RSObj("BoardName") = BoardName
				 RSObj("Note") = Note
				 RSObj("Master") = Master
				 RSObj("OrderID") =OrderID
				 RSObj("ParentID")=ParentID
				 RSObj("BoardRules")=BoardRules
				 RSObj("Settings")=Settings
			   RSObj.Update
			   RSObj.Close
			   Set RSObj = Nothing
			  End If
			  Call KS.DelCahe(KS.SiteSN & "_ClubBoard")
			  Response.Write ("<script>alert('�����޸ĳɹ�!');location.href='KS.GuestBoard.asp?Page=" & Page & "';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&OpStr=������� >> <font color=red>���Ա��������</font>';</script>")
			End If
		  End Sub
		  
		  'ɾ��
		  Sub GuestBoardDel()
		  		 Dim K, GuestBoardID, Page
				 Page = KS.G("Page")
				 GuestBoardID = Trim(KS.G("GuestBoardID"))
				 GuestBoardID = Split(GuestBoardID, ",")
				 For k = LBound(GuestBoardID) To UBound(GuestBoardID)
					Conn.Execute ("Delete From KS_GuestBoard Where ID =" & GuestBoardID(k))
					Conn.Execute ("Delete From KS_GuestBoard Where ParentID =" & GuestBoardID(k))
				 Next
				 Call KS.DelCahe(KS.SiteSN & "_ClubBoard")
				Response.Write ("<script>location.href='KS.GuestBoard.asp?Page=" & Page & "';</script>")
		  End Sub
		  
End Class
%>
 
