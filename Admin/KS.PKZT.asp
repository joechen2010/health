<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 5.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Main
KSCls.Kesion()
Set KSCls = Nothing

Class Main
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
			If Not KS.ReturnPowerResult(0, "KSMS20014") Then
			  Call KS.ReturnErr(1, "")
			  exit sub
			End If
			Action=KS.G("Action")
			Select Case Action
			 Case "Add","Edit"
				  Call MailDepartAddOrEdit()
			 Case "Save"
			      Call DoSave()
			 Case "Del"
			      Call PKDelete()
			 Case Else
			   Call MainList()
			End Select
	    End Sub
		
		Sub MainList()
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
			.Write "<script language=""JavaScript"" src=""Include/Common.js""></script>"
			.Write "<script language=""JavaScript"" src=""Include/ContextMenu1.js""></script>"
			.Write "<script language=""JavaScript"" src=""Include/SelectElement.js""></script>"
			%>
			<script language="JavaScript">
			var DocElementArrInitialFlag=false;
			var DocElementArr = new Array();
			var DocMenuArr=new Array();
			var SelectedFile='',SelectedFolder='';
			function document.onreadystatechange()
			{   if (DocElementArrInitialFlag) return;
				InitialDocElementArr('FolderID','PKID');
				InitialContextMenu();
				DocElementArrInitialFlag=true;
			}
			function InitialContextMenu()
			{	DocMenuArr[DocMenuArr.length]=new ContextMenuItem("window.parent.MailDepartAdd();",'�� ��(N)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.SelectAllElement();",'ȫ ѡ(A)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.MailDepartControl(1);",'�� ��(E)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.MailDepartControl(2);",'ɾ ��(D)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.location.reload();",'ˢ ��(Z)','disabled');
			}
			function DocDisabledContextMenu()
			{
				DisabledContextMenu('FolderID','PKID','�� ��(E),ɾ ��(D)','�� ��(E)','','','','')
			}
			function MailDepartAdd()
			{
				location.href='KS.PKZT.asp?Action=Add';
				window.parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr=��������� >> <font color=red>����»</font>&ButtonSymbol=GO';
			}
			function EditMailDepart(id)
			{
				location="KS.PKZT.asp?Action=Edit&Page="+Page+"&Flag=Edit&PKID="+id;
				window.parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr=����������� >> <font color=red>�༭�</font>&ButtonSymbol=GoSave';
			}
			function DelMailDepart(id)
			{
			if (confirm('���Ҫɾ��ѡ�еĻ��?'))
			 location="KS.PKZT.asp?Action=Del&Page="+Page+"&PKID="+id;
			   SelectedFile='';
			}
			function MailDepartControl(op)
			{  var alertmsg='';
				GetSelectStatus('FolderID','PKID');
				if (SelectedFile!='')
				 {  if (op==1)
					{
					if (SelectedFile.indexOf(',')==-1) 
						EditMailDepart(SelectedFile)
					  else alert('һ��ֻ�ܱ༭һ���!')	
					SelectedFile='';
					}
				  else if (op==2)    
					 DelMailDepart(SelectedFile);
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
				 alert('��ѡ��Ҫ'+alertmsg+'�Ļ');
				  }
			}
			function GetKeyDown()
			{ 
			if (event.ctrlKey)
			  switch  (event.keyCode)
			  {  case 90 : location.reload(); break;
				 case 65 : SelectAllElement();break;
				 case 78 : event.keyCode=0;event.returnValue=false; MailDepartAdd();break;
				 case 69 : event.keyCode=0;event.returnValue=false;MailDepartControl(1);break;
				 case 68 : MailDepartControl(2);break;
			   }	
			else	
			 if (event.keyCode==46)MailDepartControl(2);
			}
			</script>
			<%
			.Write "</head>"
			.Write "<body scroll=no topmargin=""0"" leftmargin=""0"" onclick=""SelectElement();"" onkeydown=""GetKeyDown();"" onselectstart=""return false;"">"
			  .Write "<ul id='menu_top'>"
			  .Write "<li class='parent' onclick=""MailDepartAdd();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>���PK����</span></li>"
			  .Write "<li class='parent' onclick=""MailDepartControl(1);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>�༭PK����</span></li>"
			  .Write "<li class='parent' onclick=""MailDepartControl(2);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>ɾ��PK����</span></li>"
			  .Write "</ul>"
			

			.Write "<table width=""100%""  border=""0"" cellpadding=""0"" cellspacing=""0"">"
			.Write "  <tr>"			
			.Write "          <td width=""44%"" height=""25"" class=""sort"" align=""center"">PK��������</td>"
			.Write "          <td class=""sort"" align=""center"">��Ŀ</td>"
			.Write "          <td class=""sort"" align=""center"">����ʱ��</td>"
			.Write "          <td align=""center"" class=""sort"">��Ʊ���</td>"
			.Write "          <td align=""center"" class=""sort"">״̬</td>"
			.Write "          <td align=""center"" class=""sort"">�������</td>"
			.Write "  </tr>"
			 
			 Set RSObj = Server.CreateObject("ADODB.RecordSet")
					   SqlStr = "SELECT * FROM KS_PKZT order by ID DESC"
					   RSObj.Open SqlStr, Conn, 1, 1
					 If RSObj.EOF And RSObj.BOF Then
					 Else
						totalPut = RSObj.RecordCount
			
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
			   on error resume next
			  With Response
					Do While Not RSObj.EOF
					  .Write "<tr>"
					  .Write "  <td class='splittd' width='44%' height='20'> &nbsp;&nbsp; <span PKID='" & RSObj("ID") & "' ondblclick=""EditMailDepart(this.PKID)""><img src='Images/Field.gif' align='absmiddle'>"
					  .Write "    <span style='cursor:default;'>" & KS.GotTopic(RSObj("Title"), 45) & "</span></span> "
					  .Write "  </td>"
					  .Write "  <td class='splittd' align='center'>" 
					  .Write  KS.C_C(RSObj("ClassID"),1)
					  .Write "  </td>"
					  
					  .Write "  <td class='splittd' align='center'>" 
					  if rsobj("timelimit")=1 then
					  .write rsobj("enddate")
					  else
					  .write "<font color=#cccccc>����ʱ��</font>"
					  end if
					  .Write " </td>"
					  .Write "  <td class='splittd' align='center'><a href='KS.PKGD.asp?pkid=" & rsobj("id") & "'>��:<font Color=red>" & rsobj("zfvotes") & "</font>Ʊ ��:<font Color=red>" & rsobj("ffvotes") & "</font>Ʊ ��:<font Color=red>" & rsobj("sfvotes") & "</font>Ʊ</a></td>"
					  .Write "  <td class='splittd' align='center'>"
					   if rsobj("status")=1 then
					    .write "<Font color=green>����</font>"
					   else
					    .write "<Font color=red>����</font>"
					   end if
					  .Write "</td>"
					  .Write "  <td class='splittd' align='center'><a href='../plus/pk/pk.asp?id=" & rsobj("id") &"' target='_Blank'>�鿴</a> <a href='javascript:EditMailDepart(" & rsobj("id") &")'>�༭</a> <a href='javascript:DelMailDepart(" & rsobj("id") & ")'>ɾ��</a></td>"
					  .Write "</tr>"
					 I = I + 1
					  If I >= MaxPerPage Then Exit Do
						   RSObj.MoveNext
					Loop
					  RSObj.Close
					  .Write "<tr><td height='26' colspan='5' align='right'>"
					  Call KS.ShowPageParamter(totalPut, MaxPerPage, "", True, "��", CurrentPage, "Action=" & Action)
				End With
			End Sub
			
			'����޸Ļ
		  Sub MailDepartAddOrEdit()
		  		Dim PKID, RSObj,ClassID, TimeLimit,SqlStr, NewsLink,Title,enddate, ZFTips,FFTips, CategoryID, AddDate,Flag, Page,Status,ZFVotes,FFVotes,SFVotes,LoginTf,VerifyTF,OnceTF
				Flag = KS.G("Flag")
				Page = KS.G("Page")
				If Page = "" Then Page = 1
				If Flag = "Edit" Then
					PKID = KS.G("PKID")
					Set RSObj = Server.CreateObject("Adodb.Recordset")
					SqlStr = "SELECT * FROM KS_PKZT Where ID=" & PKID
					RSObj.Open SqlStr, Conn, 1, 1
					  Title     = RSObj("Title")
					  ZFTips    = RSObj("ZFTips")
					  FFTips    = RSObj("FFTips")
					  enddate  = RSObj("enddate")
					  NewsLink = RSObj("NewsLink")
					  Status = RSObj("Status")
					  LoginTf= RSObj("LoginTf")
					  TimeLimit=RSObj("TimeLimit")
					  enddate=RSObj("EndDate")
					  ZFVotes=RSObj("ZFVotes")
					  FFVotes=RSObj("FFVotes")
					  SFVotes=RSObj("SFVotes")
					  ClassID=RSObj("ClassID")
					  VerifyTF=RSObj("verifytf")
					  OnceTF=RSObj("oncetf")
					RSObj.Close:Set RSObj = Nothing
				Else
				  Flag = "Add"
				  status=1
				  TimeLimit=0
				  enddate=now
				  ZFVotes=0
				  FFVotes=0
				  SFVotes=0
				  LoginTf=1
				  VerifyTF=1
				  OnceTF=1
				End If
				With Response
				.Write "<html>"
				.Write "<head>"
				.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
				.Write "<title>�½�PK����</title>"
				.Write "</head>"
				.Write "<script src=""Include/Common.js"" language=""JavaScript""></script>"
				.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
				.Write "<body bgcolor=""#FFFFFF"" leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
				.Write " <div class='topdashed sort'>"
				If Flag = "Edit" Then
				 .Write "�޸�PK����"
				Else
				 .Write "���PK����"
				End If
	            .Write "</div>"
				.Write "<br>"
				.Write "<table width=""100%"" border=""0"" cellpadding=""1"" cellspacing=""1"" class='ctable'>"
				.Write "  <form name=myform method=post action=""?Action=Save"">"
				.Write "   <input type=""hidden"" name=""Flag"" value=""" & Flag & """>"
				.Write "   <input type=""hidden"" name=""PKID"" value=""" & PKID & """>"
				.Write "   <input type=""hidden"" name=""Page"" value=""" & Page & """>"
				.Write "    <tr>"
				.Write "      <td height=""25"" align='right' width='85' class='clefttitle'><strong>PK��������:</strong></td>"
				.Write "      <td><input name=""Title"" type=""text"" id=""Title"" value=""" & Title & """ class=""textbox"" style=""width:60%""> ��:��������̨���֣��Ƿ������ã�</td>"
				 .Write "  </tr>"
				.Write "    <tr>"
				.Write "      <td height=""25"" align='right' width='85' class='clefttitle'><strong>ָ��Ƶ��:</strong></td>"
				.Write "      <td><select name=""ClassID"" style=""width:200;border-style: solid; border-width: 1"">"
		  If not IsObject(Application(KS.SiteSN&"_class")) Then KS.LoadClassConfig
			Dim ClassXML,Node
			Set ClassXML=Application(KS.SiteSN&"_class")
			For Each Node In ClassXML.documentElement.SelectNodes("class[@ks10=1 and @ks14=1]")
			  If Node.SelectSingleNode("@ks0").text = ClassID Then
			    .write "<option value='" &Node.SelectSingleNode("@ks0").text &"' selected>" & Node.SelectSingleNode("@ks1").text & "</option>"
			  else
			    .write "<option value='" &Node.SelectSingleNode("@ks0").text &"'>" & Node.SelectSingleNode("@ks1").text & "</option>"
			  end if
			next
				.Write "      </select> ��Ҫ���𵽰�Ƶ����������������"
				.Write "      </td></tr>"
				 
				 
				 .Write "<tr>"
				.Write "  <td height=""25"" align='right' width='85' class='clefttitle'><strong>�����۵�:</strong></td>"
				.Write "  <td><textarea ID='ZFTips' name='ZFTips' style='width:90%;height:60px'>" & ZFTips &"</textarea><br/></td></tr>"
				.Write "<tr><td height=""25"" align='right' width='85' class='clefttitle'><strong>�����۵�:</strong></td>"
				.Write "<td><textarea ID='FFTips' name='FFTips' style='width:90%;height:60px'>" & FFTips &"</textarea></td></tr>"
				.Write "<tr><td height=""25"" align='right' width='85' class='clefttitle'><strong>������������:</strong></td>"
				.Write "<td><input type='text' name='NewsLink' id='NewsLink' size='25' value='" & NewsLink & "'> ��:http://www.kesion.com/news/1.html</td></tr>"
				.Write "<tr><td height=""25"" align='right' width='85' class='clefttitle'><strong>��Ʊ���:</strong></td>"
				.Write "<td>����:<input type='text' style='text-align:center' name='ZFVotes' value='" & ZFVotes & "' size='4'> ����:<input type='text' name='FFVotes' value='" & FFVotes & "' size='4' style='text-align:center'> ������:<input type='text' name='SFVotes' value='" & SFVotes & "' size='4' style='text-align:center'></td></tr>"
				.Write "          <tr>"
				.Write "            <td height=""25"" align='right' width='85' class='clefttitle'><strong>�Ƿ������ο�PK:</strong></td>"
				.Write "            <td>"
				.write "  <Input type='radio' name='LoginTf' value='0'"
				if LoginTf="0" then .write " checked"
				.Write ">����"
				.write "  <Input type='radio' name='LoginTf' value='1'"
				if LoginTf="1" then .write " checked"
				.Write ">������"
				.Write "              </td>"
				.Write "          </tr>"
				.Write "          <tr>"
				.Write "            <td height=""25"" align='right' width='85' class='clefttitle'><strong>ÿ���û�ֻ��PKһ��:</strong></td>"
				.Write "            <td>"
				.write "  <Input type='radio' name='OnceTF' value='0'"
				if OnceTF="0" then .write " checked"
				.Write ">����"
				.write "  <Input type='radio' name='OnceTF' value='1'"
				if OnceTF="1" then .write " checked"
				.Write ">��"
				.Write "              </td>"
				.Write "          </tr>"
				.Write "          <tr>"
				.Write "            <td height=""25"" align='right' width='85' class='clefttitle'><strong>�û��۵���Ҫ���:</strong></td>"
				.Write "            <td>"
				.write "  <Input type='radio' name='VerifyTF' value='0'"
				if VerifyTF="0" then .write " checked"
				.Write ">����Ҫ"
				.write "  <Input type='radio' name='VerifyTF' value='1'"
				if VerifyTF="1" then .write " checked"
				.Write ">��Ҫ"
				.Write "              </td>"
				.Write "          </tr>"
				
				
				.Write "          <tr>"
				.Write "            <td height=""25"" align='right' width='85' class='clefttitle'><strong>״̬:</strong></td>"
				.Write "            <td>"
				.write "  <Input type='radio' name='status' value='0'"
				if status="0" then .write " checked"
				.Write ">����"
				.write "  <Input type='radio' name='status' value='1'"
				if status="1" then .write " checked"
				.Write ">����"
				.Write "              </td>"
				.Write "          </tr>"
				
				
				.Write "          <tr>"
				.Write "            <td height=""25"" align='right' width='85' class='clefttitle'><strong>�Ƿ��޶�ʱ��:</strong></td>"
				.Write "            <td>"
				
				.write "  <Input type='radio' onclick=""document.getElementById('timea').style.display='none';"" name='TimeLimit' value='0'"
				if TimeLimit="0" then .write " checked"
				.Write ">������"
				.write "  <Input type='radio'  onclick=""document.getElementById('timea').style.display='';"" name='TimeLimit' value='1'"
				if TimeLimit="1" then .write " checked"
				.Write ">����ʱ��"


               if TimeLimit="0" then
				.Write " <div id='timea' style='display:none'>"
			  Else
				.Write " <div id='timea'>"
			  End If
				.Write "<input type='text' name='enddate' value='" & enddate& "' size='30' class='textbox'> ��ʽ:YYYY-MM-DD hh:mm:ss"
				.Write "</div>"
				.Write "              </td>"
				.Write "          </tr>"

				.Write "  </form>"
				.Write "</table>"
				.Write "</body>"
				.Write "</html>"
				.Write "<script language=""JavaScript"">" & vbCrLf
				.Write "<!--" & vbCrLf
				.Write "function CheckForm()" & vbCrLf
				.Write "{ var form=document.myform;" & vbCrLf
				.Write "  if (form.Title.value=='')" & vbCrLf
				.Write "   {" & vbCrLf
				.Write "    alert('������PK��������!');" & vbCrLf
				.Write "    form.Title.focus();" & vbCrLf
				.Write "    return false;" & vbCrLf
				.Write "   }" & vbCrLf
				.Write "   if (form.ZFTips.value=='')" & vbCrLf
				.Write "   {" & vbCrLf
				'.Write "    alert('����������!');" & vbCrLf
				'.Write "    form.ZFTips.focus();" & vbCrLf
				'.Write "    return false;" & vbCrLf
				.Write "   }" & vbCrLf

				.Write "   form.submit();"
				.Write "   return true;"
				.Write "}"
				.Write "//-->"
				.Write "</script>"
			 End With
		  End Sub
		  
		  '����
		  Sub DoSave()
			Dim PKID, RSObj, SqlStr,ClassID,Title, AddDate, ZFTips, FFTips,TimeLimit,Flag, Page, RSCheck,Status,enddate,NewsLink,ZFVotes,FFVotes,SFVotes,LoginTf,VerifyTF,OnceTF
			Set RSObj = Server.CreateObject("Adodb.RecordSet")
			Flag = Request.Form("Flag")
			PKID = Request("PKID")
			Title = KS.G("Title")
			ZFTips = Request.Form("ZFTips")
			FFTips = Request.Form("FFTips")
			NewsLink=KS.G("NewsLink")
			Status = KS.ChkClng(KS.G("Status"))
			TimeLimit=KS.ChkClng(KS.G("TimeLimit"))
			ClassID=KS.G("ClassID")
			ZFVotes=KS.ChkClng(KS.G("ZFVotes"))
			FFVotes=KS.ChkClng(KS.G("FFVotes"))
			SFVotes=KS.ChkClng(KS.G("SFVotes"))
			LoginTf=KS.ChkClng(KS.G("LoginTf"))
			VerifyTF=KS.ChkClng(KS.G("VerifyTF"))
			OnceTF=KS.ChkClng(KS.G("OnceTF"))
			enddate=request("enddate")
			if not isdate(enddate) then enddate=now
			
			If Title = "" Then Call KS.AlertHistory("PK���ⲻ��Ϊ��!", -1)
			If ZFTips = "" Then Call KS.AlertHistory("PK���ⱳ�����ϲ���Ϊ��!", -1)
			
			Set RSObj = Server.CreateObject("Adodb.Recordset")
			If Flag = "Add" Then
			   RSObj.Open "Select ID From KS_PKZT Where Title='" & Title & "'", Conn, 1, 1
			   If Not RSObj.EOF Then
				  RSObj.Close
				  Set RSObj = Nothing
				  Response.Write ("<script>alert('�Բ���,PK���������Ѵ���!');history.back(-1);</script>")
				  Exit Sub
			   Else
				RSObj.Close
				RSObj.Open "SELECT * FROM KS_PKZT Where 1=0", Conn, 1, 3
				RSObj.AddNew
				  RSObj("Title") = Title
				  RSObj("ClassID")=ClassID
				  RSObj("ZFTips") = ZFTips
				  RSObj("FFTips") = FFTips
				  RSObj("NewsLink")=NewsLink
				  RSObj("AddDate")=Now
				  RSObj("TimeLimit")=TimeLimit
				  RSObj("enddate") = enddate
				  RSObj("ZFVotes") = ZFVotes
				  RSObj("FFVotes") = FFVotes
				  RSObj("SFVotes") = SFVotes
				  RSObj("LoginTf") = LoginTf
				  RSObj("VerifyTf") = VerifyTf
				  RSObj("OnceTf") = OnceTf
				  RSObj("Status") =Status
				RSObj.Update
				 RSObj.Close
			  End If
			   Set RSObj = Nothing
			   Response.Write ("<script> if (confirm('PK������ӳɹ�!���������?')) {location.href='KS.PKZT.asp?Action=Add';}else{location.href='KS.PKZT.asp';parent.frames['BottomFrame'].location.href='KS.Split.asp?ButtonSymbol=Disabled&OpStr=PKϵͳ���� >> <font color=red>PK�������</font>';}</script>")
			ElseIf Flag = "Edit" Then
			  Page = Request.Form("Page")
			  RSObj.Open "Select ID FROM KS_PKZT Where Title='" & Title & "' And ID<>" & PKID, Conn, 1, 1
			  If Not RSObj.EOF Then
				 RSObj.Close
				 Set RSObj = Nothing
				 Response.Write ("<script>alert('�Բ���,PK���������Ѵ���!');history.back(-1);</script>")
				 Exit Sub
			  Else
			   RSObj.Close
			   SqlStr = "SELECT * FROM KS_PKZT Where ID=" & PKID
			   RSObj.Open SqlStr, Conn, 1, 3
				  RSObj("Title") = Title
				  RSObj("ClassID")=ClassID
				  RSObj("ZFTips") = ZFTips
				  RSObj("FFTips") = FFTips
				  RSObj("NewsLink")=NewsLink
				  RSObj("TimeLimit")=TimeLimit
				  RSObj("enddate") = enddate
				  RSObj("ZFVotes") = ZFVotes
				  RSObj("FFVotes") = FFVotes
				  RSObj("SFVotes") = SFVotes
				  RSObj("LoginTf") = LoginTf
				  RSObj("VerifyTf") = VerifyTf
				  RSObj("OnceTf") = OnceTf
				  RSObj("Status") =Status
			   RSObj.Update
			   RSObj.Close
			   Set RSObj = Nothing
			  End If
			  Response.Write ("<script>alert('PK�����޸ĳɹ�!');location.href='KS.PKZT.asp?Page=" & Page & "';parent.frames['BottomFrame'].location.href='KS.Split.asp?ButtonSymbol=Disabled&OpStr=PKϵͳ���� >> <font color=red>PK�������</font>';</script>")
			End If
		  End Sub
		  
		  'ɾ��
		  Sub PKDelete()
		  		 Dim K, PKID, Page
				 Page = KS.G("Page")
				 PKID = Trim(KS.G("PKID"))
				 PKID = Split(PKID, ",")
				 For k = LBound(PKID) To UBound(PKID)
					Conn.Execute ("Delete From KS_PKZT Where ID =" & PKID(k))
				 Next
				 KS.Echo "<script>alert('��ϲ,PK����ɾ���ɹ�!');location.href='KS.PKZT.Asp';</script>"
		  End Sub
		  
	

End Class
%>
 
