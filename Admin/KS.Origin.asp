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
Set KSCls = New Admin_Origin
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Origin
        Private KS,Action,Page,KSCls
		Private I, totalPut, CurrentPage, OriginSql, RS,MaxPerPage
		Private OriginName,ID,Contact, Telphone, UnitName, UnitAddress, Zip, Email, QQ, HomePage, Note, OriginType

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
		With KS
		 CurrentPage = KS.ChkClng(KS.G("page"))
		 If CurrentPage=0 Then CurrentPage=1

		 .echo "<html>"
		 .echo "<head>"
		 .echo "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
		 .echo "<title>��Դ����</title>"
		 .echo "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		 .echo "<script language='JavaScript'>"
		 .echo "var Page='" & CurrentPage & "';"
		 .echo "</script>"
		 .echo "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>" & vbCrLf
	     .echo "<script language=""JavaScript"" src=""../KS_Inc/jQuery.js""></script>" & vbCrLf
	     .echo "<script language=""JavaScript"" src=""../KS_Inc/Kesion.Box.js""></script>" & vbCrLf
             Action=KS.G("Action")
			 Page=KS.G("Page")
			 
			 If Not KS.ReturnPowerResult(0, "KMST10015") Then                  'Ȩ�޼��
				Call KS.ReturnErr(1, "")   
				Response.End()
			 End iF
			 
			 Select Case Action
			  Case "Add"
			    Call OriginAddOrEdit("Add")
			  Case "Edit"
			    Call OriginAddOrEdit("Edit")
			  Case "Del"
			    Call OriginDel()
			  Case "AddSave"
			    Call OriginAddSave()
			  Case "EditSave"
			    Call OriginEditSave()
			  Case Else
			   Call OriginList()
			 End Select
			 .echo "</body>"
			 .echo "</html>"
		 End With
		End Sub
		
		Sub OriginList()
		 On Error Resume Next
		With KS
		%>
		<script language="javascript">
		   function set(v)
			{
				 if (v==1)
				 KeyWordControl(1);
				 else if (v==2)
				 KeyWordControl(2);
			}
		function OriginAdd()
		{
		  PopupCenterIframe('������Դ','KS.Origin.asp?Action=Add',630,410,'no')
		}
		function EditOrigin(id)
		{ 
		PopupCenterIframe('�༭��Դ',"KS.Origin.asp?action=Edit&ID="+id,630,410,'no')
		}
		function DelOrigin(id)
		{
		if (confirm('���Ҫɾ������Դ��?'))
		 location="KS.Origin.asp?Action=Del&Page="+Page+"&id="+id;
		  SelectedFile='';
		}
		function OriginControl(op)
		{  
		    var alertmsg='';
	        var ids=get_Ids(document.myform);
			if (ids!='')
			 {
			   if (op==1)
				{
				if (ids.indexOf(',')==-1) 
					EditOrigin(ids)
				  else alert('һ��ֻ�ܱ༭һ����Դ!')	
				}	
			  else if (op==2)    
			   DelOrigin(ids);
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
			 alert('��ѡ��Ҫ'+alertmsg+'����Դ');
			  }
		}
		function GetKeyDown()
		{ 
		if (event.ctrlKey)
		  switch  (event.keyCode)
		  {  case 90 : location.reload(); break;
			 case 65 : Select(0);break;
			 case 78 : event.keyCode=0;event.returnValue=false;OriginAdd();break;
			 case 69 : event.keyCode=0;event.returnValue=false;OriginControl(1);break;
			 case 68 : OriginControl(2);break;
		   }	
		else	
		 if (event.keyCode==46)OriginControl(2);
		}
		</script>
		<%
		 .echo "</head>"
		 .echo "<body scroll=no topmargin='0' leftmargin='0' onkeydown='GetKeyDown();' onselectstart='return false;'>"
		 .echo "<ul id='menu_top'>"
		 .echo "<li class='parent' onClick=""OriginAdd();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>������Դ</span></li>"
		 .echo "<li class='parent' onclick='OriginControl(1);'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>�޸���Դ</span></li>"
		 .echo "<li class='parent' onclick='OriginControl(2);'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>ɾ����Դ</span></li>"
		 .echo "</ul>"
		 .echo "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
		 .echo ("<form name='myform' method='Post' action='?'>")
	     .echo ("<input type='hidden' name='action' id='action' value='" & Action & "'>")
		 .echo "        <tr>"
		 .echo "          <td class=""sort"" width='35' align='center'>ѡ��</td>"
		 .echo "          <td height='25' class='sort' align='center'>��Դ����</td>"
		 .echo "          <td class='sort' align='center'>��λ����</td>"
		 .echo "          <td class='sort' align='center'>���ӵ�ַ</td>"
		 .echo "          <td class='sort' align='center'>���ʱ��</td>"
		 .echo "        </tr>"
			 Set RS = Server.CreateObject("ADODB.RecordSet")
				   OriginSql = "SELECT * FROM [KS_Origin] Where OriginType=0 order by AddDate desc"
				   RS.Open OriginSql, conn, 1, 1
				 If RS.EOF And RS.BOF Then
				 Else
					       totalPut = RS.RecordCount
		
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
									RS.Move (CurrentPage - 1) * MaxPerPage
							Else
									CurrentPage = 1
							End If
							Call showContent

			End If
		     .echo "</table>"
			 .echo ("<table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>")
	         .echo ("<tr><td width='180'><div style='margin:5px'><b>ѡ��</b><a href='javascript:Select(0)'><font color=#999999>ȫѡ</font></a> - <a href='javascript:Select(1)'><font color=#999999>��ѡ</font></a> - <a href='javascript:Select(2)'><font color=#999999>��ѡ</font></a> </div>")
	         .echo ("</td>")
	         .echo ("<td><select style='height:18px' onchange='set(this.value)' name='setattribute'><option value=0>����ѡ��...</option><option value='1'>ִ�б༭</option><option value='2'>ִ��ɾ��</option></select></td>")
	         .echo ("</form><td align='right'>")
	         Call KSCls.ShowPage(totalPut, MaxPerPage, "KS.Origin.asp", True, "λ", CurrentPage, "")
	         .echo ("</td></tr></table>")
		End With
		End Sub
		Sub showContent()
		With KS
		Do While Not RS.EOF
		  .echo "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" & RS("ID") & "' onclick=""chk_iddiv('" & RS("ID") & "')"">"
		  .echo "    <td class='splittd' align=center><input name='id'  onclick=""chk_iddiv('" &RS("ID") & "')"" type='checkbox' id='c"& RS("ID") & "' value='" &RS("ID") & "'></td>"
		  .echo "    <td class='splittd' height='19'><span OriginID='" & RS("ID") & "' onDblClick=""EditOrigin(this.OriginID)"">"
		  .echo "     <img src='Images/Origin.gif' width='20' height='19' align='absmiddle'><span  style='cursor:default;'>" & RS("OriginName") & "</span></span>"
		  .echo "   </td>"
		  .echo "   <td class='splittd' align='center'>&nbsp;" & RS("UnitName") & " </td>"
		  .echo "   <td class='splittd' align='center'>" & RS("HomePage") & "</td>"
		  .echo "   <td class='splittd' align='center'>" & RS("AddDate") & " </td>"
		  .echo " </tr>"
				I = I + 1
				If I >= MaxPerPage Then Exit Do
				RS.MoveNext
			 Loop
				RS.Close
         End With
		 End Sub
		 
		 Sub OriginAddOrEdit(OpType)
		 With KS
		  Dim RS, OriginSql
		 ID = KS.G("ID")
		 .echo "<html>"
		 .echo "<head>"
		 .echo "<meta http-equiv='Content-Type' content='text/html; chaRSet=gb2312'>"
		 .echo "<title>��Դ����</title>"
		 .echo "<link href='../Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		 .echo "<script language='JavaScript' src='../KS_Inc/common.js'></script>"
		 .echo "</head>"
		 .echo "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>"
		
		 Action="AddSave"
		 HomePage="http://"
		 If Optype = "Edit" Then
			 Set RS = Server.CreateObject("ADODB.RECORDSET")
			 OriginSql = "Select * From [KS_Origin] Where ID='" & ID & "'"
			 RS.Open OriginSql, conn, 1, 1
		 If Not RS.EOF Then
			 OriginName = Trim(RS("OriginName"))
			 Contact = Trim(RS("Contact"))
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
		 .echo "  <form  action='KS.Origin.asp?ID=" & ID &"&page=" & Page &"' method='post' name='OrigArticlerm' onsubmit='return(CheckForm())'>"
		 .echo "    <input type='hidden' value='" & Action &"' name='action'>"
		 .echo "    <tr class='tdbg'>"
		 .echo "      <td width='200' height='25' align='right' class='clefttitle'>��Դ���ƣ�</td>"
		 .echo "       <td><input name='OriginName' value='" & OriginName &"' type='text' id='OriginName' class='textbox'></td>"
		 .echo "    </tr>"
		 .echo "    <tr class='tdbg'>"
		 .echo "      <td height='25' class='clefttitle' align='right' nowrap>�� ϵ �ˣ�</td>"
		 .echo "      <td height='25' nowrap><input name='Contact' value='" & Contact &"' type='text' id='Contact' class='textbox'></td>"
		 .echo "   </tr>"
		 .echo "    <tr class='tdbg'>"
		 .echo "      <td height='25' align='right' class='clefttitle' nowrap>��ϵ�绰��</td>"
		 .echo "      <td height='25' nowrap> <input name='Telphone' value='" & Telphone &"' type='text' id='Telphone' class='textbox'></td>"
		 .echo "    </tr>"
		 .echo "   <tr class='tdbg'>"
		 .echo "      <td height='25' align='right' nowrap class='clefttitle'>��λ���ƣ�</td>"
		 .echo "      <td height='25'><input name='UnitName' type='text' value='" & UnitName &"' id='UnitName' class='textbox'></td>"
		 .echo "    </tr>"
		 .echo "    <tr class='tdbg'>"
		 .echo "      <td height='25' align='right' nowrap class='clefttitle'>��λ��ַ��</td>"
		 .echo "      <td height='25'> <input name='UnitAddress' type='text' value='" & UnitAddress &"' class='textbox'></td>"
		 .echo "    </tr>"
		 .echo "    <tr class='tdbg'>"
		 .echo "      <td height='25' align='right' nowrap class='clefttitle'>�������룺</td>"
		 .echo "      <td height='25' nowrap> <input name='Zip' type='text' value='" & zip &"' id='Zip' class='textbox'></td>"
		 .echo "    </tr>"
		 .echo "    <tr class='tdbg'>"
		 .echo "      <td align='right' class='clefttitle'>�������䣺</td>"
		 .echo "      <td><input name='Email' type='text' id='Email' value='" & Email &"' class='textbox'></td>"
		 .echo "    </tr>"
		 .echo "    <tr class='tdbg'>"
		 .echo "      <td height='25' align='right' nowrap class='clefttitle'>��ϵ QQ��</td>"
		 .echo "      <td height='25' nowrap> <input name='QQ' type='text' id='QQ' value='" & QQ &"' class='textbox'></td>"
		 .echo "    </tr>"
		 .echo "    <tr class='tdbg'>"
		 .echo "       <td height='25' align='right' class='clefttitle'>��ҳ��ַ��</td>"
		 .echo "       <td><input name='HomePage' type='text' id='HomePage' class='textbox' value='" & HomePage &"'></td>"
		 .echo "    </tr>"
		 .echo "    <tr class='tdbg'>"
		 .echo "      <td align='right' class='clefttitle'>��ע˵����</td>"
		 .echo "      <td height='25' nowrap> <textarea name='Note' cols='50' style='width:250px;height:80px' rows='6' id='Note' class='textbox'>" & Note &"</textarea>"
		 .echo "     </td>"
		 .echo "    </tr>"
		 .echo "    <input type='hidden' name='OriginType' value='0'>"
		 .echo "  </form>"
		 .echo "</table>"
		 
		 .echo "<div id='save'>"
		 .echo "<li class='parent' onclick=""return(CheckForm())""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/save.gif' border='0' align='absmiddle'>ȷ������</span></li>"
		 .echo "<li class='parent' onclick=""parent.closeWindow()""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/back.gif' border='0' align='absmiddle'>�ر�ȡ��</span></li>"
		 .echo "</div>"

		 .echo "<Script Language='javascript'>"
		 .echo "function CheckForm()"
		 .echo "{ var form=document.OrigArticlerm;"
		 .echo "   if (form.OriginName.value=='')"
		 .echo "    {"
		 .echo "     alert('��������Դ����!');"
		 .echo "     form.OriginName.focus();"
		 .echo "     return false;"
		 .echo "    }"
		 .echo "   if ((form.Zip.value!='')&&((form.Zip.value.length>6)||(!is_number(form.Zip.value))))"
		 .echo "    {"
		 .echo "     alert('�Ƿ���������!');"
		 .echo "     form.Zip.focus();"
		 .echo "     return false;"
		 .echo "    }"
		 .echo "    if ((form.Email.value!='')&&(is_email(form.Email.value)==false))"
		 .echo "    {"
		 .echo "    alert('�Ƿ���������!');"
		 .echo "     form.Email.focus();"
		 .echo "     return false;"
		 .echo "    }"
		 .echo "    form.submit();"
		 .echo "    return true;"
		 .echo "}"
		 .echo "</Script>"
		End With
		End Sub
		 
		 Sub OriginAddSave()
		 Dim RS
		 OriginName = KS.G("OriginName")
		 Contact = KS.G("Contact")
		 Telphone = KS.G("Telphone")
		 UnitName = KS.G("UnitName")
		 UnitAddress = KS.G("UnitAddress")
		 Zip = KS.G("Zip")
		 Email = KS.G("Email")
		 QQ = KS.G("QQ")
		 HomePage = KS.G("HomePage")
		 Note = KS.G("Note")
		 OriginType =KS.G("OriginType")

		 If OriginName = "" Then Call KS.AlertHistory("��������Դ����!", -1): Exit Sub
		 Set RS = Server.CreateObject("ADODB.RECORDSET")
		 OriginSql = "Select * From [KS_Origin] Where OriginName='" & OriginName & "' And OriginType=0"
		 RS.Open OriginSql, conn, 3, 3
		 If RS.EOF And RS.BOF Then
		  RS.AddNew
		  RS("ID") = Year(Now) & Month(Now) & Day(Now) & KS.MakeRandom(5)
		  RS("OriginName") = OriginName
		  RS("Contact") = Contact
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
		  KS.Echo ("<Script> if (confirm('��Դ���ӳɹ�,���������?')) { location.href='KS.Origin.asp?Action=Add';} else{top.frames[""MainFrame""].location.reload();}</script>")
		 Else
		   Call KS.AlertHistory("���ݿ����Ѵ��ڸ���Դ����!", -1)
		   Exit Sub
		 End If
		 RS.Close
		 End Sub
		 
		 Sub OriginEditSave()
		 With KS
		 ID = KS.G("ID")
		 OriginName = KS.G("OriginName")
		 Contact = KS.G("Contact")
		 Telphone = KS.G("Telphone")
		 UnitName = KS.G("UnitName")
		 UnitAddress = KS.G("UnitAddress")
		 Zip = KS.G("Zip")
		 Email = KS.G("Email")
		 QQ = KS.G("QQ")
		 HomePage = KS.G("HomePage")
		 Note = KS.G("Note")
		 OriginType =KS.G("OriginType")
		  If OriginName = "" Then Call KS.AlertHistory("��������Դ����!", -1): Exit Sub
		 Set RS = Server.CreateObject("ADODB.RECORDSET")
		  OriginSql = "Select * From [KS_Origin] Where ID='" & ID & "'"
		  RS.Open OriginSql, conn, 1, 3
		  RS("OriginName") = OriginName
		  RS("Contact") = Contact
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
		  KS.Echo ("<Script> alert('��Դ�޸ĳɹ�!');top.frames[""MainFrame""].location.reload();</script>")
		 End With
		 End Sub
		 
		 Sub OriginDel()
			Dim ID:ID = KS.G("ID")
			ID = Replace(ID, ",", "','")
			ID = "'" & ID & "'"
			conn.Execute ("Delete From KS_Origin Where ID IN(" & ID & ")")
			Response.Redirect "KS.Origin.asp?Page=" & Page
		 End Sub
End Class
%> 
