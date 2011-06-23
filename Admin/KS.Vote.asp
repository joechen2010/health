<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Admin_Vote
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Vote
        Private KS
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
			If Not KS.ReturnPowerResult(0, "KSMS20003") Then
			  Call KS.ReturnErr(1, "")
			  exit sub
			End If
			Select Case KS.G("Action")
			 Case "Add" Call VoteAdd()
			 Case "Edit" Call VoteEdit()
			 Case "Del" Call VoteDel()
			 Case "Set" Call VoteSet()
			 Case Else Call MainList()
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
			.Write "<title>վ�����</title>"
			.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "<script language=""JavaScript"">" & vbCrLf
			.Write "var Page='" & CurrentPage & "';" & vbCrLf
			.Write "</script>" & vbCrLf
			.Write "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>"
			.Write "<script language=""JavaScript"" src=""../KS_Inc/jquery.js""></script>"
			.Write "<script language=""JavaScript"" src=""Include/ContextMenu1.js""></script>"
			.Write "<script language=""JavaScript"" src=""Include/SelectElement.js""></script>"
			%>
			<script language="javascript">
			$(document).ready(function(){
				
		      $(parent.frames["BottomFrame"].document).find("#Button1").attr("disabled",true);
			  $(parent.frames["BottomFrame"].document).find("#Button2").attr("disabled",true);
		     })
			var DocElementArrInitialFlag=false;
			var DocElementArr = new Array();
			var DocMenuArr=new Array();
			var SelectedFile='',SelectedFolder='';
			function document.onreadystatechange()
			{   if (DocElementArrInitialFlag) return;
				InitialDocElementArr('FolderID','VoteID');
				InitialContextMenu();
				DocElementArrInitialFlag=true;
			}
			function InitialContextMenu()
			{	DocMenuArr[DocMenuArr.length]=new ContextMenuItem("window.parent.VoteAdd();",'�� ��(N)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.VoteControl(3);",'�� ��(S)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.SelectAllElement();",'ȫ ѡ(A)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.VoteControl(1);",'�� ��(E)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.VoteControl(2);",'ɾ ��(D)','disabled');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem('seperator','','');
				DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.location.reload();",'ˢ ��(Z)','disabled');
			}
			function DocDisabledContextMenu()
			{
				DisabledContextMenu('FolderID','VoteID','�� ��(S),�� ��(E),ɾ ��(D)','�� ��(S),�� ��(E)','','','','')
			}
			function VoteAdd()
			{
				location.href='KS.Vote.asp?Action=Add';
				window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=��������������� >> <font color=red>����µ�������</font>&ButtonSymbol=VoteAddSave';
			}
			function EditVote(id)
			{
				location="KS.Vote.asp?Page="+Page+"&Action=Edit&VoteID="+id;
				window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=��������������� >> <font color=red>�༭��������</font>&ButtonSymbol=VoteEdit';
			}
			function DelVote(id)
			{
			if (confirm('���Ҫɾ��ѡ�еĵ���������?'))
			 location="KS.Vote.asp?Action=Del&Page="+Page+"&Voteid="+id;
			 SelectedFile='';
			}
			function SetVoteNewest(id)
			{
				location="KS.Vote.asp?Action=Set&Page="+Page+"&Voteid="+id;
			}
			function VoteControl(op)
			{  var alertmsg='';
				GetSelectStatus('FolderID','VoteID');
				if (SelectedFile!='')
				 {  if (op==1)
					{
					if (SelectedFile.indexOf(',')==-1) 
						EditVote(SelectedFile)
					  else alert('һ��ֻ�ܱ༭һ����������!')	
					SelectedFile='';
					}
				  else if (op==2)    
					 DelVote(SelectedFile);
				  else if(op==3)
					 {
					if (SelectedFile.indexOf(',')==-1) 
					 SetVoteNewest(SelectedFile);
				   else alert('һ��ֻ�ܱ༭һ����������!')	
					SelectedFile='';	 
					}
				 }
				else 
				 {
				 if (op==1)
				  alertmsg="�༭";
				 else if(op==2)
				  alertmsg="ɾ��"; 
				 else if (op==3)
				  alertmsg="��Ϊ����"
				 else
				  {
				  WindowReload();
				  alertmsg="����" 
				  }
				 alert('��ѡ��Ҫ'+alertmsg+'�ĵ�������');
				  }
			}
			function GetKeyDown()
			{ 
			if (event.ctrlKey)
			  switch  (event.keyCode)
			  {  case 90 : location.reload(); break;
				 case 65 : SelectAllElement();break;
				 case 78 : event.keyCode=0;event.returnValue=false; VoteAdd();break;
				 case 69 : event.keyCode=0;event.returnValue=false;VoteControl(1);break;
				 case 68 : VoteControl(2);break;
				 case 83 : event.keyCode=0;event.returnValue=false;VoteControl(3);break;
			   }	
			else	
			 if (event.keyCode==46)VoteControl(2);
			}
			</script>
			<%
			.Write "</head>"
			.Write "<body scroll=no topmargin=""0"" leftmargin=""0"" onclick=""SelectElement();"" onkeydown=""GetKeyDown();"" onselectstart=""return false;"">"
		    .Write "<ul id='menu_top'>"
			.Write "<li class='parent' onclick=""VoteAdd();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>��ӵ���</span></li>"
			.Write "<li class='parent' onclick=""VoteControl(1);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>�༭����</span></li>"
			.Write "<li class='parent' onclick=""VoteControl(2);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>ɾ������</span></li>"
			.Write "</ul>"
			.Write "<table width=""100%""  border=""0"" cellpadding=""0"" cellspacing=""1"">"
			.Write "  <tr>"
			.Write "    <td height=""3"" colspan=""4"" valign=""top""> <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
			.Write "        <tr>"
			.Write "          <td width=""44%"" height=""25"" class=""sort""align=""center"">��������</td>"
			.Write "          <td width=""17%"" class=""sort"" align=""center"">������</td>"
			.Write "          <td width=""20%"" align=""center"" class=""sort"">ʱ��</td>"
			.Write "          <td width=""19%"" class=""sort"" align=""center"">�Ƿ�����</td>"
			.Write "        </tr>"
			.Write "      </table></td>"
			.Write "  </tr>"
			 
			 Set RSObj = Server.CreateObject("ADODB.RecordSet")
					   SqlStr = "SELECT * FROM KS_Vote order by NewestTF desc,AddDate desc"
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
				
			.Write "    </td>"
			.Write "  </tr>"
			.Write "</table>"
			.Write "</body>"
			.Write "</html>"
			End With
			End Sub
			Sub showContent()
			  With Response
					Do While Not RSObj.EOF
					  .Write "<tr>"
					  .Write "  <td class='splittd' width='44%' height='20'> &nbsp;&nbsp; <span VoteID='" & RSObj("ID") & "' ondblclick=""EditVote(this.VoteID)""><img src='Images/Vote.gif' align='absmiddle'>"
					  .Write "    <span style='cursor:default;'>" & KS.GotTopic(RSObj("Title"), 50) & "</span></span> "
					  .Write "  </td>"
					  .Write "  <td class='splittd' width='17%' align='center'>" & RSObj("UserName") & " </td>"
					  .Write "  <td class='splittd' width='20%' align='center'><FONT Color=red>" & RSObj("AddDate") & "</font> </td>"
					  If RSObj("NewestTF") = 1 Then
					   .Write "  <td class='splittd' width='19%' align='center'><font color=red>��</font></td>"
					  Else
					   .Write "  <td class='splittd' width='19%' align='center'>��</td>"
					  End If
					  .Write "</tr>"

					I = I + 1
					  If I >= MaxPerPage Then Exit Do
						   RSObj.MoveNext
					Loop
					  RSObj.Close
					  Conn.Close
					 .Write "<tr><td height='26' colspan='4' align='right'>"
					 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
			  End With
			End Sub
			
			Sub VoteDel()
			 Conn.Execute("delete from KS_Vote where ID="&Clng(KS.G("VoteID")))
			 Response.redirect "KS.Vote.asp?Page="&KS.G("Page")
			End Sub
			
			Sub VoteSet()
				conn.execute "Update KS_Vote set NewestTF=0 where NewestTF=1"
				conn.execute "Update KS_Vote set NewestTF=1 Where ID=" & Clng(KS.G("VoteID"))
				Response.Write "<script language='JavaScript' type='text/JavaScript'>alert('���óɹ���');location.href='KS.Vote.asp?Page=" & KS.G("Page") & "';</script>"

			End Sub
			
			Sub VoteAdd()
				With Response
				.Write "<html>"
				.Write "<head>"
				.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
				.Write "<title>�������-�������</title>"
				.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
				.Write "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
				.Write "</head>"
				.Write "<body scroll=no topmargin=""0"" leftmargin=""0"" onselectstart=""return false;"">"
	
				.Write "  <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
				.Write "        <tr>"
				.Write "          <td width=""44%"" height=""25"" class=""sort"">"
				.Write "          <div align=""center""><strong>�� �� �� �� �� ��</strong></div></td>"
				.Write "        </tr>"
				.Write "      </table>"
	
				
				dim Title,VoteTime,NewestTF
				dim rs,sql
				Title=trim(request.form("Title"))
				VoteTime=trim(request.form("VoteTime"))
				if VoteTime="" then VoteTime=now()
				NewestTF=trim(request("NewestTF"))
				
				dim i
				if Title<>"" then
					sql="select top 1 * from KS_Vote"
					Set rs= Server.CreateObject("ADODB.Recordset")
					rs.open sql,conn,1,3
					rs.addnew
					rs("Title")=Title
					for i=1 to 8
						if trim(request("select"&i))<>"" then
							rs("select"&i)=trim(request("select"&i))
							if request("answer"&i)="" then
								rs("answer"&i)=0
							else
								rs("answer"&i)=clng(request("answer"&i))
							end if
						end if
					next
					rs("AddDate")=VoteTime
					rs("VoteType")=request("VoteType")
					rs("UserName")=KS.C("AdminName")
					if NewestTF="" then NewestTF=0
					if NewestTF=1 then conn.execute "Update KS_Vote set NewestTF=0 where NewestTF=1"
					rs("NewestTF")=NewestTF
					rs.update
					rs.close
					set rs=nothing
					call CloseConn()
					.Redirect "KS.Vote.asp"
				end if
				 End With
				%>
	
				<BR>
				<table cellpadding="2" cellspacing="1" border="0" width="690" align="center" class="a2">
					
					<tr>
						<td align="center">
					  <br>
							<form method="POST" name="voteform" action="KS.Vote.asp?Action=Add">
						<table width="624" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC" Class="border">
						  <tr> 
							<td width="101" height="26" align="right" bgcolor="#EEF8FE">�������ƣ�</td>
							<td colspan="3" bgcolor="#EEF8FE">
							<input name="Title" type="text" size="40" maxlength="50">
							�磺��Ա�վ����Щ��Ŀ�ϸ���Ȥ!</td>
						  </tr>
				<% for i=1 to 8%>
									<tr> 
							<td height="25" align="right" bgcolor="#EEF8FE">ѡ��<%=i%>��</td>
							<td width="303" bgcolor="#EEF8FE">
									  <input type="text" name="select<%=i%>" size="36"></td>
							<td width="80" align="right" bgcolor="#EEF8FE">��ʼƱ����</td>
							<td width="135" bgcolor="#EEF8FE">
									  <input name="answer<%=i%>" type="text" value="0" size="5">
									  Ʊ</td>
									</tr>
				<%next%>
								  <tr> 
							<td height="25" align="right" bgcolor="#EEF8FE">�������ͣ�</td>
							<td colspan="3" bgcolor="#EEF8FE">
										<select name="VoteType" id="VoteType">
											<option value="Single" selected>��ѡ</option>
											<option value="Multi">��ѡ</option>
									</select>
										<input name="NewestTF" type="checkbox" id="NewestTF" value="1" checked />
	��Ϊ���µ���</td>
						  </tr>
									
									<tr> 
							<td height="25" colspan=4 align=center bgcolor="#EEF8FE"><BR>
							  <span style="color:red">��ע��������Ϊÿ�������趨�˸�����ѡ������˸�ѡ��Ŀ�����!</span></td>
									</tr>
							  </table>
							</form>
						</td>
					</tr>
	</table>
	<script>
	 function CheckForm()
	 { var form=document.voteform;
	  if (form.Title.value=='')
	   {
		 alert('�������������!');
		  form.Title.focus();
		 return false;
	   }
	  document.voteform.submit();
	 }
	</script>
<%
			End Sub
			'�༭
			Sub VoteEdit()
			
			Response.Write "<html>"
			Response.Write "<head>"
			Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			Response.Write "<title>�������-�޸�����</title>"
			Response.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
		    Response.Write "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
  			Response.Write "</head>"
			Response.Write "<body scroll=no topmargin=""0"" leftmargin=""0"" onselectstart=""return false;"">"

			Response.Write "  <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
			Response.Write "        <tr>"
			Response.Write "          <td width=""44%"" height=""25"" class=""sort"">"
			Response.Write "          <div align=""center""><strong>�� �� �� �� �� ��</strong></div></td>"
			Response.Write "        </tr>"
			Response.Write "      </table>"
			
			dim ID,Title,AddDate,NewestTF
			dim rs,sql
			ID=KS.G("Voteid")   
			Title=trim(KS.G("Title"))
			AddDate=trim(KS.G("AddDate"))
			if AddDate="" then AddDate=now()
			NewestTF=trim(KS.G("NewestTF"))
			if NewestTF="" Then NewestTF=0
			if NewestTF=1 then
				conn.execute "Update KS_Vote set NewestTF=0 where NewestTF=1"
			end if
			dim i
			if ID="" then
				Response.Redirect "KS.Vote.asp"
			end if
			sql="select * from KS_Vote where ID="&Cint(ID)
			Set rs= Server.CreateObject("ADODB.Recordset")
			rs.open sql,conn,1,3
			
			if not rs.eof then

				if Title<>"" then
					rs("Title")=Title
					for i=1 to 8
						if trim(KS.G("select"&i))<>"" then
							rs("select"&i)=trim(KS.G("select"&i))
							if KS.G("answer"&i)="" then
								rs("answer"&i)=0
							else
								rs("answer"&i)=clng(KS.G("answer"&i))
							end if
						else
							rs("select"&i)=""
							rs("answer"&i)=0
						end if
					next
					rs("AddDate")=AddDate
					rs("VoteType")=KS.G("VoteType")
					if NewestTF="" then NewestTF=0
					rs("NewestTF")=NewestTF
					rs.update
					Response.Write ("<script>alert('���������޸ĳɹ�!');location.href='KS.Vote.asp?Page=" & KS.G("Page") & "';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=VoteList&OpStr=������� >> <font color=red>վ�ڵ������</font>';</script>")
				end if
				%>
			<BR>
			<table cellpadding="2" cellspacing="1" border="0" width="690" align="center" class="a2">
				<tr class="a4">
					<td align="center">
				  <br>
						<form method="POST" name="voteform" action="KS.Vote.asp?Action=Edit&page=<%=KS.G("Page")%>">
					<table width="624" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#CCCCCC" Class="border">
					  <tr> 
									<td width="15%" height="25" align="right" bgcolor="#EEF8FE">�������ƣ�</td>
						<td width="85%" height="25" colspan="3" bgcolor="#EEF8FE"><input name="Title" type="text" value="<%=Trim(rs("Title"))%>" size="40">
						�磺��Ա�վ����Щ��Ŀ�ϸ���Ȥ!</td>
					  </tr>
								<% for i=1 to 8%>
								<tr> 
								  <td height="25" align="right" bgcolor="#EEF8FE">ѡ��<%=i%>��</td>
									<td height="25" bgcolor="#EEF8FE"><input name="select<%=i%>" type="text" value="<%=Trim(rs("select"& i))%>" size="36">					  </td>
									<td height="25" align="right" bgcolor="#EEF8FE">Ʊ��</td>
								  <td width="80" height="25" bgcolor="#EEF8FE"><input name="answer<%=i%>" type="text" value="<%=Trim(rs("answer"&i))%>" size="5">
								  Ʊ</td>
								</tr>
								<%next%>
								<tr> 
									<td height="25" align="right" bgcolor="#EEF8FE">�������ͣ�</td>
									<td height="25" colspan="3" bgcolor="#EEF8FE"><select name="VoteType" id="VoteType">
										<option value="Single" <% if rs("VoteType")="Single" then %> selected <% end if%>>��ѡ</option>
										<option value="Multi" <% if rs("VoteType")="Multi" then %> selected <% end if%>>��ѡ</option>
								  </select>
									  <input name="NewestTF" type="checkbox" id="NewestTF" value="1" <% if rs("NewestTF")=1 then Response.Write "checked"%> />
			��Ϊ���µ���</td>
								</tr>
								
								<tr>
									<td height="25" colspan=4 align=center bgcolor="#EEF8FE"> <BR>
										<input name="VoteID" type="hidden" id="VoteID" value="<%=rs("ID")%>"> 
								 <span style="color:red">��ע��������Ϊÿ�������趨�˸�����ѡ������˸�ѡ��Ŀ�����!</span></td>
</td>
								</tr>
						  </table>
						</form>
					</td>
				</tr>
			</table>
			<BR>
			<script>
		 function CheckForm()
		 { var form=document.voteform;
		  if (form.Title.value=='')
		   {
			 alert('�������������!');
			  form.Title.focus();
			 return false;
		   }
		  document.voteform.submit();
		 }
		</script>
			<%
			end if
			rs.close:set rs=nothing
			End Sub
End Class
%>
 
