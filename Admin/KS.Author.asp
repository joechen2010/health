<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
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
Set KSCls = New Admin_Author
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Author
        Private KS,Action,Page,KSCls
		Private I, totalPut, CurrentPage,MaxPerPage, SqlStr,ChannelID,ItemName1,FlagName,Flag1Name,RS
		Private OriginName, ID, Sex, Birthday, Telphone, UnitName, UnitAddress, Zip, Email, QQ, HomePage, Note, OriginType
		
		Private Sub Class_Initialize()
		  MaxPerPage = 20
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub


		Public Sub Kesion()
		    CurrentPage = KS.ChkClng(KS.G("page"))
		    If CurrentPage=0 Then CurrentPage=1
			   ChannelID=KS.ChkClng(KS.G("ChannelID"))
             With KS
		 	    .echo "<html>"
				.echo "<head>"
				.echo "<meta http-equiv='Content-Type' content='text/html; chaRSet=gb2312'>"
				.echo "<title>���߹���</title>"
				.echo "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
				.echo "<script language='JavaScript'>"
				.echo "var Page='" & CurrentPage & "';"
				.echo "var ChannelID=" & ChannelID & ";"
				.echo "</script>"
				.echo "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>" & vbCrLf
				.echo "<script language=""JavaScript"" src=""../KS_Inc/Jquery.js""></script>" & vbCrLf
				.echo "<script language=""javascript"" src=""../KS_Inc/popcalendar.js""></script>"
				.echo "<script language=""JavaScript"" src=""../KS_Inc/Kesion.Box.js""></script>" & vbCrLf
             Action=KS.G("Action")
			 
			 If ChannelID=0 Then
			   If Not KS.ReturnPowerResult(ChannelID, "KMST10016") Then                  'Ȩ�޼��
				Call KS.ReturnErr(1, "")   
				Response.End()
			   End iF
			 Else
				If Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "20003") Then                  'Ȩ�޼��
				Call KS.ReturnErr(1, "")   
				Response.End()
				End iF
             End if
			 
			 Page=KS.G("Page")
			 Select Case KS.C_S(ChannelID,6)
			  Case 0
			   ItemName1="����":FlagName="��������":Flag1Name="�����Ա�"
			  Case 3
			   ItemName1="������":FlagName="������":Flag1Name="�����Ա�"
			  Case 5
			   ItemName1="����":FlagName="��������":Flag1Name="��ϵ�绰"
			 End Select
			 
			 Select Case Action
			  Case "Add"
			    Call AddOrEdit("Add")
			  Case "Edit"
			    Call AddOrEdit("Edit")
			  Case "Del"
			    Call AuthorDel()
			  Case "AddSave"
			    Call AuthorAddSave()
			  Case "EditSave"
			    Call AuthorEditSave()
			  Case Else
			   Call ShowMain()
			 End Select
			.echo "</body>"
			.echo "</html>"
			End With
		End Sub
		
		Sub ShowMain()
	%>
	   <script language="javascript">
	   	function set(v)
		{
			 if (v==1)
				 AuthorControl(1);
			 else if (v==2)
				 AuthorControl(2);
		}
		function AuthorAdd()
		{
		 PopupCenterIframe('��������/������','KS.Author.asp?ChannelID='+ChannelID+'&Action=Add',630,410,'no')
		}
		function EditAuthor(id)
		{
		  PopupCenterIframe('�༭����','KS.Author.asp?ChannelID='+ChannelID+'&Action=Edit&ID='+id,630,410,'no')
		}
		function DelAuthor(id)
		{
		 if (confirm('���Ҫɾ����������?'))
		  location="KS.Author.asp?ChannelID="+ChannelID+"&Action=Del&Page="+Page+"&id="+id;
		}
		function AuthorControl(op)
		{ var alertmsg='';
			var ids=get_Ids(document.myform);
			if (ids!='')
			 {
			   if (op==1)
				{
				if (ids.indexOf(',')==-1) 
					EditAuthor(ids)
				  else alert('һ��ֻ�ܱ༭һ������!')
				}	
			  else if (op==2)    
			  DelAuthor(ids);
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
			 alert('��ѡ��Ҫ'+alertmsg+'������');
			  }
		}
		function GetKeyDown()
		{ event.returnValue=false;
		if (event.ctrlKey)
		  switch  (event.keyCode)
		  {  case 90 : location.reload(); break;
			 case 65 : Select(0);break;
			 case 78 : event.keyCode=0;event.returnValue=false;AuthorAdd();break;
			 case 69 : event.keyCode=0;event.returnValue=false;AuthorControl(1);break;
			 case 68 : AuthorControl(2);break;
		   }	
		else	
		 if (event.keyCode==46)AuthorControl(2);
		}
	   </script>
	<%
	   With KS
		.echo "</head>"
		
		.echo "<body scroll=no topmargin='0' leftmargin='0' onkeydown='GetKeyDown();' onselectstart='return false;'>"
		.echo "<ul id='menu_top'>"
		.echo "<li class='parent' onClick=""AuthorAdd();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>����"&ItemName1 &"</span></li>"
		.echo "<li class='parent' onclick='AuthorControl(1);'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>�޸�"&ItemName1 &"</span></li>"
		.echo "<li class='parent' onclick='AuthorControl(2);'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>ɾ��"&ItemName1 &"</span></li>"
		.echo "</ul>"
		.echo "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
		.echo ("<form name='myform' method='Post' action='?channelid="& channelid & "'>")
	    .echo ("<input type='hidden' name='action' id='action' value='" & Action & "'>")
		.echo "   <tr>"
		.echo "          <td class=""sort"" width='35' align='center'>ѡ��</td>"
		.echo "          <td class='sort' align='center'>" & FlagName &"</td>"
		.echo "          <td class='sort'><div align='center'>" & Flag1Name & "</td>"
		.echo "          <td align='center' class='sort'>��������</td>"
		.echo "          <td class='sort' align='center'>���ʱ��</td>"
		.echo "  </tr>"
		 Set RS = Server.CreateObject("ADODB.RecordSet")
				   SqlStr = "SELECT * FROM [KS_Origin] Where ChannelID="& ChannelID& " AND OriginType=1 order by AddDate desc"
				   RS.Open SqlStr, conn, 1, 1
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
		
							If CurrentPage = 1 Then
								Call showContent
							Else
								If (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
									
									Call showContent
								Else
									CurrentPage = 1
									Call showContent
								End If
							End If
			End If
		    .echo "</table>"
			.echo ("<table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>")
	        .echo ("<tr><td width='180'><div style='margin:5px'><b>ѡ��</b><a href='javascript:Select(0)'><font color=#999999>ȫѡ</font></a> - <a href='javascript:Select(1)'><font color=#999999>��ѡ</font></a> - <a href='javascript:Select(2)'><font color=#999999>��ѡ</font></a> </div>")
	        .echo ("</td>")
	        .echo ("<td><select style='height:18px' onchange='set(this.value)' name='setattribute'><option value=0>����ѡ��...</option><option value='1'>ִ�б༭</option><option value='2'>ִ��ɾ��</option></select></td>")
	        .echo ("</form><td align='right'>")
	         Call KSCls.ShowPage(totalPut, MaxPerPage, "KS.Author.asp", True, "λ", CurrentPage, "ChannelID=" & ChannelID)
	        .echo ("</td></tr></table>")
		End With
		End Sub
		Sub showContent()
		  With KS
			Do While Not RS.EOF
			.echo "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" & RS("ID") & "' onclick=""chk_iddiv('" & RS("ID") & "')"">"
			.echo "<td class='splittd' align=center><input name='id'  onclick=""chk_iddiv('" &RS("ID") & "')"" type='checkbox' id='c"& RS("ID") & "' value='" &RS("ID") & "'></td>"
		    .echo "  <td  class='splittd' height='19'><span AuthorID='" & RS("ID") & "'  onDblClick=""EditAuthor(this.AuthorID)"">"
		    .echo "    <img src='Images/Author.gif' align='absmiddle'><span style='cursor:default;'>"& RS("OriginName") & "</span></span> </td>"
		   
		   If ChannelID=5 Then
		   .echo " <td  class='splittd' align='center'>" & RS("Telphone") & " </td>"
		   Else
		   .echo " <td  class='splittd' align='center'>" & RS("Sex") & " </td>"
		   End If
		   .echo " <td  class='splittd' align='center'>&nbsp;" & RS("Email") & "</td>"
		   .echo " <td  class='splittd' align='center'>" & RS("AddDate") & " </td>"
		   .echo "</tr>"
							  I = I + 1
								If I >= MaxPerPage Then Exit Do
							   RS.MoveNext
							   Loop
								RS.Close
						  
		   End With
		 End Sub
		 
		 Sub AddOrEdit(OpType)
		  With KS
		   Dim RS, OriginSql
		   ID = Request("ID")
		
		  Action="AddSave"
		  Sex="��"
		  If OpType = "Edit" Then
			 Set RS = Server.CreateObject("ADODB.RECORDSET")
			 OriginSql = "Select * From [KS_Origin] Where ID='" & ID & "'"
			 RS.Open OriginSql, conn, 1, 1
			 If Not RS.EOF Then
			 OriginName = Trim(RS("OriginName"))
			 ChannelID = RS("ChannelID")
			 Sex = Trim(RS("Sex"))
			 Birthday = Trim(RS("Birthday"))
			 Telphone = Trim(RS("Telphone"))
			 UnitName = Trim(RS("UnitName"))
			 UnitAddress = Trim(RS("UnitAddress"))
			 Zip = Trim(RS("Zip"))
			 Email = Trim(RS("Email"))
			 QQ = Trim(RS("QQ"))
			 HomePage = Trim(RS("HomePage"))
			 Note = Trim(RS("Note"))
			 Action="EditSave"
		    End If
		
        End If
	
		.echo "<table width='100%' border='0' align='center' cellpadding='1' cellspacing='1' class='ctable' style='margin-top:5px;border-collapse: collapse'>"
		.echo "  <form  action='KS.Author.asp?ID=" & ID &"&page=" & Page & "' method='post' name='AuthorForm' onsubmit='return(CheckForm())'>"
		.echo "     <input type='hidden' value='" & ChannelID & "' name='ChannelID'>"
		.echo "    <tr class='tdbg'>"
		.echo "      <td class='clefttitle' width='200' height='25' align='right' nowrap>" & FlagName &"��</td>"
		.echo "      <td width='21' height='30' align='right'> <input name='OriginName' value='" & OriginName & "' type='text' id='OriginName' style='width:200;border-style: solid; border-width: 1'>"
		.echo "    </tr>"
		
		if ChannelID<>5 then
		.echo "    <tr class='tdbg'>"
		.echo "      <td height='25' class='clefttitle' align='right' nowrap>�����Ա�</td>"
		.echo "      <td height='25' nowrap>"
			 If Sex = "��" Then
			   .echo "<input name='Sex' type='radio' value='��' Checked> ��"
			  Else
			   .echo "<input name='Sex' type='radio' value='��'> ��"
			  End If
			  If Sex = "Ů" Then
			   .echo "<input name='Sex' type='radio' value='Ů' Checked> Ů"
			  Else
			   .echo "<input name='Sex' type='radio' value='Ů'> Ů"
			  End If
		.echo "       </td>"
		.echo "    <tr class='tdbg'>"
		.echo "        <td class='clefttitle' align='right'>�������ڣ�</td>"
		.echo "        <td>"
		.echo "        <input name='Birthday' type='text' id='Birthday' value='" & Birthday & "' style='border-style: solid; border-width: 1' size='15' readonly>"
		.echo "        <a href='#' onClick=""popUpCalendar(this, $('input[name=Birthday]').get(0), dateFormat,-1,-1)""><img src='Images/date.gif' border='0' align='absmiddle' title='ѡ������'></a>"
		.echo "        </td>"
		.echo "    </tr>"
	   end if
		.echo "    <tr class='tdbg'>"
		.echo "      <td height='25' align='right' class='clefttitle'>��ϵ�绰��</td>"
		.echo "      <td height='25'> <input name='Telphone' type='text' value='" & Telphone & "' id='Telphone' style='border-style: solid; border-width: 1' size='50'></td>"
		.echo "    </tr>"
		
		if ChannelID<>5 then
		.echo "    <tr class='tdbg'>"
		.echo "      <td height='25' align='right' class='clefttitle'>��λ���ƣ�</td>"
		.echo "      <td height='25'> <input name='UnitName' type='text' id='UnitName' value='" & UnitName & "' style='border-style: solid; border-width: 1' size='50'></td>"
		.echo "   </tr>"
		end if 
		.echo "    <tr class='tdbg'>"
		.echo "      <td height='25' align='right' class='clefttitle'>��λ��ַ��</td>"
		.echo "      <td height='25'> <input name='UnitAddress' type='text' id='UnitAddress' value='" & UnitAddress & "' style='border-style: solid; border-width: 1' size='50'></td>"
		.echo "    </tr>"
		.echo "    <tr class='tdbg'>"
		.echo "     <td height='25' align='right' class='clefttitle'>�������룺</td>"
		.echo "     <td height='25' nowrap> <input name='Zip' type='text' id='Zip' value='" & Zip & "' style='border-style: solid; border-width: 1'></td>"
		.echo "    </tr>"
		.echo "    <tr class='tdbg'>"
		.echo "     <td align='right' class='clefttitle'>��������:</td>"
		.echo "     <td><input name='Email' type='text' id='Email' value='" & Email & "' style='border-style: solid; border-width: 1'></td>"
		.echo "    </tr>"
		.echo "    <tr class='tdbg'>"
		.echo "      <td height='25' align='right' class='clefttitle'>��ϵQQ��</td>"
		.echo "      <td height='25'> <input name='QQ' type='text' id='QQ' value='" & QQ & "' style='border-style: solid; border-width: 1'></td>"
		.echo "    </tr>"
		.echo "    <tr class='tdbg'>"
		.echo "       <td height='25' align='right' class='clefttitle'>��ҳ��ַ:</td>"
		.echo "       <td><input name='HomePage' type='text' id='HomePage' value='" & HomePage & "' style='border-style: solid; border-width: 1' value='http://'></td>"
		.echo "    </tr>"
		.echo "    <tr class='tdbg'>"
		.echo "      <td align='right' class='clefttitle'>��ע˵����</td>"
		.echo "      <td height='25'> <textarea name='Note' cols='50' rows='6' id='Note' style='border-style: solid; border-width: 1'>" & Note & "</textarea>"
		.echo "      </td>"
		.echo "    </tr>"
		.echo "    <input type='hidden' value='" & Action & "' name='Action'>"
		.echo "    <input type='hidden' name='OriginType' value='1'>"
		.echo "  </form>"
		.echo "</table>"
		.echo "<div id='save'>"
		.echo "<li class='parent' onclick=""return(CheckForm())""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/save.gif' border='0' align='absmiddle'>ȷ������</span></li>"
		.echo "<li class='parent' onclick=""parent.closeWindow();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/back.gif' border='0' align='absmiddle'>�ر�ȡ��</span></li>"
		.echo "</div>"
	
		.echo "<Script Language='javascript'>"
		.echo "function CheckForm()"
		.echo "{ var form=document.AuthorForm;"
		.echo "   if (form.OriginName.value=='')"
		.echo "    {"
		.echo "     alert('������" & FlagName &"!');"
		.echo "     form.OriginName.focus();"
		.echo "     return false;"
		.echo "    }"
		.echo "    if ((form.Zip.value!="""")&&((form.Zip.value.length>6)||(!is_number(form.Zip.value))))"
		.echo "    {"
		.echo "     alert('�Ƿ���������!');"
		.echo "     form.Zip.focus();"
		.echo "     return false;"
		.echo "    }"
		.echo "    if (form.Email.value!="""")"
		.echo "    if(is_email(form.Email.value)==false)"
		.echo "    { alert('�Ƿ���������!');"
		.echo "     form.Email.focus();"
		.echo "     return false;"
		.echo "    }"
		.echo "    form.submit();"
		.echo "    return true;"
		.echo "}"
		.echo "</Script>"
        End With
		 End Sub
		 
		 Sub AuthorAddSave()
		 OriginName = Trim(Request.Form("OriginName"))
		 Sex = Trim(Request.Form("Sex"))
		 Birthday = Request.Form("Birthday")
		 Telphone = Trim(Request.Form("Telphone"))
		 UnitName = Trim(Request.Form("UnitName"))
		 UnitAddress = Trim(Request.Form("UnitAddress"))
		 Zip = Trim(Request.Form("Zip"))
		 Email = Trim(Request.Form("Email"))
		 QQ = Trim(Request.Form("QQ"))
		 HomePage = Trim(Request.Form("HomePage"))
		 Note = Trim(Request.Form("Note"))
		 OriginType = CInt(Request.Form("OriginType"))
		 
		 If OriginName = "" Then Call KS.AlertHistory("��������������!", -1):Set KS = Nothing
		 Dim RS:Set RS = Server.CreateObject("ADODB.RECORDSET")
		 Dim OriginSQL:OriginSql = "Select * From [KS_Origin] Where OriginName='" & OriginName & "' And ChannelID=" & KS.G("ChannelID") & " And OriginType=1"
		 RS.Open OriginSql, conn, 3, 3
		 If RS.EOF And RS.BOF Then
		  RS.AddNew
		  RS("ID") = Year(Now) & Month(Now) & Day(Now) & KS.MakeRandom(5)
		  RS("OriginName") = OriginName
		  RS("ChannelID") = KS.G("ChannelID")
		  RS("Sex") = Sex
		   If Birthday <> "" Then
		  RS("Birthday") = Birthday
		   End If
		  RS("Telphone") = Telphone
		  RS("UnitName") = UnitName
		  RS("UnitAddress") = UnitAddress
		  RS("Zip") = Zip
		  RS("Email") = Email
		  RS("QQ") = QQ
		  RS("HomePage") = HomePage
		  RS("Note") = Note
		  RS("OriginType") = OriginType
		  RS("AddDate") = Now()
		  RS.Update
		  Set conn = Nothing
		  KS.Echo ("<Script> if (confirm('�������ӳɹ�,���������?')) { location.href='KS.Author.asp?ChannelID="& ChannelID& "&Action=Add';} else{top.frames[""MainFrame""].location.reload();}</script>")
		 Else
		 Call KS.AlertHistory("���ݿ����Ѵ��ڸ�����!", -1)
		 Set KS = Nothing:.End
		 End If
		 RS.Close
		 End Sub
		 
		 Sub AuthorEditSave()
		 With Response
		 ID = Request("ID")
		 OriginName = Trim(Request.Form("OriginName"))
		 Sex = Trim(Request.Form("Sex"))
		 Birthday = Request.Form("Birthday")
		 Telphone = Trim(Request.Form("Telphone"))
		 UnitName = Trim(Request.Form("UnitName"))
		 UnitAddress = Trim(Request.Form("UnitAddress"))
		 Zip = Trim(Request.Form("Zip"))
		 Email = Trim(Request.Form("Email"))
		 QQ = Trim(Request.Form("QQ"))
		 HomePage = Trim(Request.Form("HomePage"))
		 Note = Trim(Request.Form("Note"))
		 OriginType = CInt(Request.Form("OriginType"))

		 If OriginName = "" Then Call KS.AlertHistory("��������������!", -1)
		 Dim RS:Set RS = Server.CreateObject("ADODB.RECORDSET")
		 Dim OriginSQL:OriginSql = "Select * From [KS_Origin] Where ID='" & ID & "'"
		  RS.Open OriginSql, conn, 1, 3
		  RS("OriginName") = OriginName
		  RS("ChannelID") = KS.G("ChannelID")
		  RS("Sex") = Sex
		   If Birthday <> "" Then
		  RS("Birthday") = Birthday
		   End If
		  RS("Telphone") = Telphone
		  RS("UnitName") = UnitName
		  RS("UnitAddress") = UnitAddress
		  RS("Zip") = Zip
		  RS("Email") = Email
		  RS("QQ") = QQ
		  RS("HomePage") = HomePage
		  RS("Note") = Note
		  RS.Update
		  RS.Close
		  Set RS=Nothing
		  
		   KS.Echo ("<Script> alert('�����޸ĳɹ�!');top.frames[""MainFrame""].location.reload();</script>")
		   Set conn = Nothing
		  End With
		 End Sub
		 
		 Sub AuthorDel()
			Dim ID:ID = KS.G("ID")
			ID = Replace(ID, ",", "','")
			ID = "'" & ID & "'"
			conn.Execute ("Delete From KS_Origin Where ID IN(" & ID & ")")
			Response.Redirect "KS.Author.asp?ChannelID=" & ChannelID &"&Page=" & Page
		 End Sub
End Class
%> 
