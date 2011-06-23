<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.FunctionCls.asp"-->
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
Set KSCls = New Announce_Main
KSCls.Kesion()
Set KSCls = Nothing

Class Announce_Main
        Private KS,KSCls,Action
		Private I, totalPut, CurrentPage, SqlStr, RSObj
        Private MaxPerPage
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
			If Not KS.ReturnPowerResult(0, "KSMS20002") Then
			  Call KS.ReturnErr(1, "")
			  exit sub
			End If
			Action=KS.G("Action")
			If Request("page") <> "" Then
				  CurrentPage = CInt(Request("page"))
			Else
				  CurrentPage = 1
			End If
			With KS
			.echo "<html>"
			.echo "<head>"
			.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			.echo "<title>վ�㹫��</title>"
			.echo "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.echo "<script language=""JavaScript"">" & vbCrLf
			.echo "var Page='" & CurrentPage & "';" & vbCrLf
			.echo "</script>" & vbCrLf
			.echo "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>"
			.echo "<script language=""JavaScript"" src=""../KS_Inc/jquery.js""></script>"
			Select Case Action
			 Case "Add","Edit" Call AnnounceAddOrEdit()
			 Case "Save" Call AnnounceSave()
			 Case "Del" Call AnnounceDel()
			 Case Else Call MainList()
			End Select
		  End With
	    End Sub
		
		Sub MainList()
			With KS
			%>
			<script language="JavaScript">
			$(document).ready(function(){
				
		      $(parent.frames["BottomFrame"].document).find("#Button1").attr("disabled",true);
			  $(parent.frames["BottomFrame"].document).find("#Button2").attr("disabled",true);
		     })

			function AnnounceAdd()
			{
				location.href='KS.Announce.asp?Action=Add';
				window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=����������� >> <font color=red>����¹���</font>&ButtonSymbol=GO';
			}
			function EditAnnounce(id)
			{ 
			    if (id=='') id=get_Ids(document.myform);
				if (id==''){
				 alert('��ѡ��Ҫ�༭�Ĺ���!');
				}else if(id.indexOf(',')==-1){
				location="KS.Announce.asp?Action=Edit&Page="+Page+"&Flag=Edit&AnnounceID="+id;
				window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=����������� >> <font color=red>�༭����</font>&ButtonSymbol=GoSave';
				}else{
				alert('һ��ֻ�ܱ༭һ������!');
				}
			}
			function DelAnnounce(id)
			{
			 if (id=='') id=get_Ids(document.myform);
			 if (id==''){
			   alert('����ѡ��Ҫɾ���Ĺ���!')
			 }else if (confirm('���Ҫɾ��ѡ�еĹ�����?')){
				 location="KS.Announce.asp?Action=Del&Page="+Page+"&id="+id;
				}
			 }
			</script>
			<%
			.echo "</head>"
			.echo "<body topmargin=""0"" leftmargin=""0"">"
			.echo "<ul id='menu_top'>"
			.echo "<li class='parent' onclick=""AnnounceAdd();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>��ӹ���</span></li>"
			.echo "<li class='parent' onclick=""EditAnnounce('');""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>�༭����</span></li>"
			.echo "<li class='parent' onclick=""DelAnnounce('');""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>ɾ������</span></li>"
			.echo "</ul>"
			.echo "<table width=""100%""  border=""0"" cellpadding=""0"" cellspacing=""0"">"
			.echo "<form name=""myform"" id=""myform"" action=""KS.Announce.asp"" method=""post"">"
			.echo "<input type=""hidden"" value=""Del"" name=""Action"" ID=""Action"">"
			.echo "<input type=""hidden"" value="""& CurrentPage & """ name=""Page"" ID=""Page"">"
			.echo "  <tr  align=""center"">"			
			.echo "          <td width=""35"" height=""25"" class=""sort"">ѡ��</div></td>"
			.echo "          <td  height=""25"" class=""sort"">�� ��</div></td>"
			.echo "          <td class=""sort""><div align=""center"">����ģ��</div></td>"
			.echo "          <td class=""sort""><div align=""center"">������</div></td>"
			.echo "          <td align=""center"" class=""sort"">����ʱ��</td>"
			.echo "          <td class=""sort"">�Ƿ�����</td>"
			.echo "          <td class=""sort"">�������</td>"
			.echo "  </tr>"
			 
			 Set RSObj = Server.CreateObject("ADODB.RecordSet")
					   SqlStr = "SELECT * FROM KS_Announce order by AddDate desc"
					   RSObj.Open SqlStr, Conn, 1, 1
					 If RSObj.EOF And RSObj.BOF Then
					 Else
						       totalPut = RSObj.RecordCount
			
								If CurrentPage < 1 Then CurrentPage = 1
			
								If (CurrentPage - 1) * MaxPerPage > totalPut Then
									If (totalPut Mod MaxPerPage) = 0 Then
										CurrentPage = totalPut \ MaxPerPage
									Else
										CurrentPage = totalPut \ MaxPerPage + 1
									End If
								End If
			
								If CurrentPage > 1  and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrentPage - 1) * MaxPerPage
								Else
										CurrentPage = 1
								End If
								 Dim AnnounceXMl:Set AnnounceXml=KS.ArrayToXml(RSObj.GetRows(MaxPerPage),RSObj,"row","root")
							     Call showContent(AnnounceXml)
								 Set AnnounceXMl=Nothing

				End If
				RSObj.Close
				Set RSObj=Nothing
			.echo "    </td>"
			.echo "  </tr>"
             .echo " <tr>"
			 .echo " <td colspan='2'><div style='margin:5px'><b>ѡ��</b><a href='javascript:void(0)' onclick='Select(0)'>ȫѡ</a> -  <a href='javascript:void(0)' onclick='Select(1)'>��ѡ</a> - <a href='javascript:void(0)' onclick='Select(2)'>��ѡ</a> <input type='submit' class='button' value='ɾ ��' onclick=""return(confirm('ȷ��ɾ��ѡ�еĹ�����?'))""></td></form>"
			 .echo "   <td align=""right"" colspan=8>"
				 Call KSCLS.ShowPage(totalPut, MaxPerPage, "KS.FriendLink.asp", True, "��", CurrentPage, KS.QueryParam("page"))
			.echo "   </td>"
			.echo "  </tr>"			
			.echo "</table>"
			.echo "</body>"
			.echo "</html>"
			End With
			End Sub
			Sub showContent(AnnounceXML)
			  Dim ID,Node
			  With KS
			   If IsObject(AnnounceXML) Then
			    For Each Node In AnnounceXML.DocumentElement.SelectNodes("row")
				       ID=Node.SelectSingleNode("@id").text
					   .echo ("<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" &ID & "' onclick=""chk_iddiv('" & ID & "')"">")
				       .echo ("<td class='splittd' align=center><input type='hidden' value='" & ID & "' name='LinkID'><input name='id'  onclick=""chk_iddiv('" & ID & "')"" type='checkbox' id='c"& ID & "' value='" & ID & "'></td>")
				  
					   .echo "  <td class='splittd' height='20'><span AnnounceID='" & ID & "' ondblclick=""EditAnnounce(this.AnnounceID)""><img src='Images/Announce.gif' align='absmiddle'>"
					   .echo "    <span style='cursor:default;'>" & KS.GotTopic(Node.SelectSingleNode("@title").text, 45) & "</span></span> "
					   .echo "  </td>"
					   .echo "  <td class='splittd' align='center'>"
					   select case Node.SelectSingleNode("@channelid").text
					    case 0:.echo "��վ��ҳ"
						case 9999:.echo "ģ�͹��ù���"
						case 9990:.echo "��Ա����"
						case else
					      .echo .C_S(Node.SelectSingleNode("@channelid").text,1) & " </td>"
					   end select
					  .echo "  <td class='splittd' align='center'>" & Node.SelectSingleNode("@author").text & " </td>"
					  .echo "  <td class='splittd' align='center'><FONT Color=red>" & Node.SelectSingleNode("@adddate").text & "</font> </td>"
					  If Node.SelectSingleNode("@newesttf").text = 1 Then
					   .echo "  <td class='splittd' align='center'><font color=red>��</font></td>"
					  Else
					   .echo "  <td class='splittd' align='center'>��</td>"
					  End If
					   .echo "  <Td class='splittd' align='center'><a href=""javascript:EditAnnounce('');"">�޸�</a> |  <a href=""javascript:DelAnnounce(" & ID &")"">ɾ��</a> </td>"
					  .echo "</tr>"
					Next
				End If
			 End With
			End Sub
			
			'����޸Ĺ���
		  Sub AnnounceAddOrEdit()
		  		Dim AnnounceID, RSObj, SqlStr, Content, Title, Author, NewestTF, AddDate,Flag, Page,ChannelID
				NewestTF = 1
				Flag = KS.G("Flag")
				Page = KS.G("Page")
				If Page = "" Then Page = 1
				If Flag = "Edit" Then
					AnnounceID = KS.G("AnnounceID")
					Set RSObj = Server.CreateObject("Adodb.Recordset")
					SqlStr = "SELECT * FROM KS_Announce Where ID=" & AnnounceID
					RSObj.Open SqlStr, Conn, 1, 1
					  Title     = RSObj("Title")
					  Author    = RSObj("Author")
					  AddDate   = RSObj("AddDate")
					  Content   = RSObj("Content")
					  NewestTF  = RSObj("NewestTF")
					  ChannelID = RSObj("ChannelID")
					RSObj.Close:Set RSObj = Nothing
				Else
				  Flag = "Add"
				End If
				With KS
				.echo "<html>"
				.echo "<head>"
				.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
				.echo "<title>�½�JS</title>"
				.echo "</head>"
				.echo "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
				.echo "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
				.echo "<body bgcolor=""#FFFFFF"" leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
				.echo "<table width='100%' cellspacing='0' cellpadding='0' border='0'>"
				.echo " <tr><td class='sort'>"
				If Flag = "Edit" Then
				 .echo "�޸Ĺ���"
				Else
				 .echo "��ӹ���"
				End If
	            .echo "</td>"
				.echo "</tr>"
				.echo "</table>"
				.echo "<table width=""100%"" border=""0"" cellpadding=""1"" cellspacing=""1"" class='ctable'>"
				.echo "  <form name=AnnounceForm method=post action=""?Action=Save"">"
				.echo "   <input type=""hidden"" name=""Flag"" value=""" & Flag & """>"
				.echo "   <input type=""hidden"" name=""AnnounceID"" value=""" & AnnounceID & """>"
				.echo "   <input type=""hidden"" name=""Page"" value=""" & Page & """>"
				.echo "   <input type=""hidden"" name=""Content"" ID=""Content"" value=""" & Server.HTMLEncode(Content) & """>"
				.echo "          <tr>"
				.echo "           <td height=""25"" align='right' width='85' class='clefttitle'><strong>����ģ��:</strong></td>"
				.echo "           <td>"
				.echo ReturnChannelList(ChannelID)
				.echo "       </td></tr>"
				.echo "          <tr>"
				.echo "            <td height=""25"" align='right' width='85' class='clefttitle'><strong>�������:</strong></td>"
				.echo "             <td>"
				.echo "              <input name=""Title"" type=""text"" id=""Title"" value=""" & Title & """ class=""textbox"" style=""width:70%""></td>"
				 .echo "</tr>"
				 .echo "<tr>"
				.echo "  <td height=""25"" align='right' width='85' class='clefttitle'><strong>�� �� ��:</strong></td>"
				.echo "  <td>"
				.echo "<input name=""Author"" type=""text"" id=""Author""  value="""
				If Flag = "Edit" Then
				.echo (Author)
				Else: .echo (KS.C("AdminName"))
				End If
				.echo """ class=""textbox"" style=""width:70%""></td>"
				.echo "          </tr>"
				.echo "          <tr>"
				.echo "            <td height=""25"" align='right' width='85' class='clefttitle'><strong>����ʱ��:</strong></td>"
				.echo "            <td>"
				.echo "              <input name=""AddDate"" type=""text"" id=""AddDate"" value="""
				 If Flag <> "Edit" Then
				 .echo (Now)
				 Else
				 .echo (AddDate)
				 End If
				.echo """ Readonly class=""textbox"" style=""width:70%"">"
				.echo "                <a href=""#"" onClick=""OpenThenSetValue('Include/DateDialog.asp',160,170,window,document.AnnounceForm.AddDate);document.AnnounceForm.AddDate.focus();""><img src=""Images/date.gif"" border=""0"" align=""absmiddle"" title=""ѡ������""></a>"
				.echo "              </td>"
				.echo "          </tr>"
				.echo "          <tr>"
				.echo "            <td height=""25"" align='right' width='85' class='clefttitle'>"
				.echo "<strong>���¹���:</strong></td>"
				.echo "            <td>"
				.echo "              <input name=""NewestTF"" type=""checkbox"" id=""NewestTF"" value=""1"""
							  If NewestTF = 1 Then .echo (" checked")
							  
				.echo "              >"
				.echo "              �򹴱�ʾ��Ϊ���¹���</td>"
				.echo "          </tr>"
				
				.echo "    <tr>"
				.echo "      <td align='right' width='85' class='clefttitle'><strong>��������:</strong></td>"
				.echo "      <td valign=""top"">"
				.echo "<iframe id=""content___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=Content&amp;Toolbar=NewsTool"" width=""695"" height=""400"" frameborder=""0"" scrolling=""no""></iframe>"
				.echo "</td></tr>"
				.echo "  </form>"
				.echo "</table>"
				.echo "</body>"
				.echo "</html>"
				.echo "<script language=""JavaScript"">" & vbCrLf
				.echo "<!--" & vbCrLf
				.echo "function CheckForm()" & vbCrLf
				.echo "{ var form=document.AnnounceForm;" & vbCrLf
				.echo "  if (form.Title.value=='')" & vbCrLf
				.echo "   {" & vbCrLf
				.echo "    alert('�����빫�����!');" & vbCrLf
				.echo "    form.Title.focus();" & vbCrLf
				.echo "    return false;" & vbCrLf
				.echo "   }" & vbCrLf
				.echo "   if (form.Author.value=='')" & vbCrLf
				.echo "   {" & vbCrLf
				.echo "    alert('�����빫������!');" & vbCrLf
				.echo "    form.Author.focus();" & vbCrLf
				.echo "    return false;" & vbCrLf
				.echo "   }" & vbCrLf
				.echo "      if (form.AddDate.value=='')" & vbCrLf
				.echo "   {" & vbCrLf
				.echo "    alert('�����빫�淢������!');" & vbCrLf
				.echo "    form.AddDate.focus();" & vbCrLf
				.echo "    return false;" & vbCrLf
				.echo "   }" & vbCrLf
				.echo "  if (FCKeditorAPI.GetInstance('Content').GetXHTML(true)=='')" & vbCrLf
				.echo "  {" & vbCrLf
				.echo "    alert('�����빫������!');" & vbCrLf
				.echo "    FCKeditorAPI.GetInstance('Content').Focus();" & vbcrlf
				.echo "    return false;" & vbCrLf
				.echo "   }" & vbCrLf
				.echo "   form.submit();"
				.echo "   return true;"
				.echo "}"
				.echo "//-->"
				.echo "</script>"
			 End With
		  End Sub
		  
		  '����
		  Sub AnnounceSave()
			Dim AnnounceID, RSObj, SqlStr, Title, Author, AddDate, Content, NewestTF,Flag, Page, RSCheck,ChannelID
			Set RSObj = Server.CreateObject("Adodb.RecordSet")
			Flag = Request.Form("Flag")
			AnnounceID = Request("AnnounceID")
			Title = Replace(Replace(Request.Form("Title"), """", ""), "'", "")
			Author = Replace(Replace(Request.Form("Author"), """", ""), "'", "")
			AddDate = Replace(Replace(Request.Form("AddDate"), """", ""), "'", "")
			Content = Replace(Request.Form("Content"), "'", "")
			NewestTF = Request.Form("NewestTF")
			ChannelID = KS.ChkClng(KS.G("ChannelID"))
			If NewestTF = "" Then NewestTF = 0
			If Title = "" Then Call KS.AlertHistory("������ⲻ��Ϊ��!", -1)
			If Author = "" Then Call KS.AlertHistory("�������߲���Ϊ��!", -1)
			If AddDate = "" Then Call KS.AlertHistory("����������ڲ���Ϊ��!", -1)
			If Content = "" Then Call KS.AlertHistory("�������ݲ���Ϊ��!", -1)
			
			Set RSObj = Server.CreateObject("Adodb.Recordset")
			If Flag = "Add" Then
			   RSObj.Open "Select ID From KS_Announce Where Title='" & Title & "'", Conn, 1, 1
			   If Not RSObj.EOF Then
				  RSObj.Close
				  Set RSObj = Nothing
				  KS.Echo ("<script>alert('�Բ���,�����Ѵ���!');history.back(-1);</script>")
				  Exit Sub
			   Else
				RSObj.Close
				RSObj.Open "SELECT * FROM KS_Announce Where (ID is Null)", Conn, 1, 3
				RSObj.AddNew
				  RSObj("Title") = Title
				  RSObj("Author") = Author
				  RSObj("AddDate") = AddDate
				  RSObj("Content") = KS.ReplaceInnerLink(Content)
				  RSObj("NewestTF") = NewestTF
				  RSObj("ChannelID") =ChannelID
				RSObj.Update
				 RSObj.MoveLast
				 If NewestTF = 1 Then
				   Conn.Execute ("UpDate KS_Announce Set NewestTF=0 Where ID<>" & RSObj("ID"))
				 End If
				 Call KS.FileAssociation(1019,RSObj("ID"),RSObj("Content"),0)
				 RSObj.Close
			  End If
			   Set RSObj = Nothing
			   KS.Echo ("<script> if (confirm('������ӳɹ�!���������?')) {location.href='KS.Announce.asp?Action=Add';}else{location.href='KS.Announce.asp';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&OpStr=������� >> <font color=red>�����������</font>';}</script>")
			ElseIf Flag = "Edit" Then
			  Page = Request.Form("Page")
			  RSObj.Open "Select ID FROM KS_Announce Where Title='" & Title & "' And ID<>" & AnnounceID, Conn, 1, 1
			  If Not RSObj.EOF Then
				 RSObj.Close
				 Set RSObj = Nothing
				 KS.Echo ("<script>alert('�Բ���,�����Ѵ���!');history.back(-1);</script>")
				 Exit Sub
			  Else
			   RSObj.Close
			   SqlStr = "SELECT * FROM KS_Announce Where ID=" & AnnounceID
			   RSObj.Open SqlStr, Conn, 1, 3
				 RSObj("Title") = Title
				 RSObj("Author") = Author
				 RSObj("AddDate") = AddDate
				 RSObj("Content") = Content
				 RSObj("NewestTF") = NewestTF
				 RSObj("ChannelID") =ChannelID
			   RSObj.Update
				If NewestTF = 1 Then
				   Conn.Execute ("UpDate KS_Announce Set NewestTF=0 Where ID<>" & RSObj("ID"))
				End If
				Call KS.FileAssociation(1019,AnnounceID,RSObj("Content"),1)
			   RSObj.Close
			   Set RSObj = Nothing
			  End If
			  KS.Echo ("<script>alert('�����޸ĳɹ�!');location.href='KS.Announce.asp?Page=" & Page & "';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&OpStr=������� >> <font color=red>�����������</font>';</script>")
			End If
		  End Sub
		  
		  'ɾ��
	Sub AnnounceDel()
		  		 Dim AnnounceID, Page
				 Page = KS.G("Page")
				 AnnounceID = Trim(KS.G("ID"))
				 Conn.Execute("Delete From KS_UploadFiles Where ChannelID=1019 and infoid in(" & KS.FilterIds(AnnounceID) & ")")
				 Conn.Execute("Delete From KS_Announce Where ID in (" & KS.FilterIds(AnnounceID) & ")")
				 KS.Echo ("<script>location.href='KS.Announce.asp?Page=" & Page & "';</script>")
	End Sub
		  
	Public Function ReturnChannelList(SelectChannelID)
	  Dim ChannelStr:ChannelStr = ""
		  ChannelStr = "<select name=""ChannelID"" style=""width:200;border-style: solid; border-width: 1"">"
		  ChannelStr = ChannelStr & "<option value=0 style='color:blue'>��վ��ҳ����</option>"
		  If SelectChannelID=9999 Then
		  ChannelStr = ChannelStr & "<option value=9999 selected style='color:red'>ģ�͹��ù���</option>"
		  Else
		  ChannelStr = ChannelStr & "<option value=9999 style='color:red'>ģ�͹��ù���</option>"
		  End If
		  If SelectChannelID=9990 Then
		  ChannelStr = ChannelStr & "<option value=9990 selected style='color:red'>��Ա���Ĺ���</option>"
		  Else
		  ChannelStr = ChannelStr & "<option value=9990 style='color:red'>��Ա���Ĺ���</option>"
		  End If
		  
		  If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
			Dim ModelXML,Node
			Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
			For Each Node In ModelXML.documentElement.SelectNodes("channel")
			 if Node.SelectSingleNode("@ks21").text="1"  Then
			  If trim(Node.SelectSingleNode("@ks0").text) = trim(SelectChannelID) Then
			  ChannelStr = ChannelStr & "<option value='" &Node.SelectSingleNode("@ks0").text &"' selected>" & Node.SelectSingleNode("@ks1").text & "</option>"
			  else
			  ChannelStr = ChannelStr & "<option value='" &Node.SelectSingleNode("@ks0").text &"'>" & Node.SelectSingleNode("@ks1").text & "</option>"
			  end if
			 End If
			next
		 
		 ChannelStr = ChannelStr & "</Select>"
	   ReturnChannelList = ChannelStr
	End Function

End Class
%>
 
