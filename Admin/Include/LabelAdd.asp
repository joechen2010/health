<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New LabelAdd
KSCls.Kesion()
Set KSCls = Nothing

Class LabelAdd
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		Dim LabelID, LabelRS, SQLStr, LabelName, Descript, LabelContent, LabelFlag, ParentID
		Dim Action, Page, RSCheck, FolderID
		Dim KeyWord, SearchType, StartDate, EndDate
		  
		'�ռ���������
		KeyWord = Request("KeyWord")
		SearchType = Request("SearchType")
		StartDate = Request("StartDate")
		EndDate = Request("EndDate")
		
		With Response
		.Write "<script src='../../ks_inc/jquery.js'></script>"
		Set LabelRS = Server.CreateObject("Adodb.RecordSet")
		Action = Request.QueryString("Action")
		Page = Request("Page")
		If Action = "EditLabel" Then
			LabelID = Request("LabelId")
			Set LabelRS = Server.CreateObject("Adodb.Recordset")
			SQLStr = "SELECT * FROM [KS_Label] Where ID='" & LabelID & "'"
			LabelRS.Open SQLStr, Conn, 1, 1
			LabelName = Replace(Replace(LabelRS("LabelName"), "{LB_", ""), "}", "")
			Descript = LabelRS("Description")
			FolderID =LabelRS("FolderID")
			LabelContent = Server.HTMLEncode(LabelRS("LabelContent"))
			LabelRS.Close
		Else
		  LabelName=Request.QueryString("LabelName")
		  Descript=Request.QueryString("Description")
		  FolderID = Request.QueryString("FolderID")
		  LabelContent=Request.QueryString("LabelContent")
		  If LabelContent="" Then LabelContent="���������Զ����html����"
		End If
		Select Case Request.Form("Action")
		 Case "AddNewSubmit"
			ParentID = Request.Form("ParentID")
			LabelName = Replace(Replace(Trim(Request.Form("LabelName")), """", ""), "'", "")
			Descript = Replace(Trim(Request.Form("Description")), "'", "")
			LabelContent = Trim(Request.Form("LabelContent"))
			LabelFlag = Request.Form("LabelFlag")
			FolderID = Request.Form("FolderID")
			If LabelName = "" Then
			   Call KS.AlertHistory("��ǩ���Ʋ���Ϊ��!", -1)
			   Set KS = Nothing
			   .End
			End If
			If LabelContent = "" Then
			  Call KS.AlertHistory("��ǩ���ݲ���Ϊ��!", -1)
			  Set KS = Nothing
			  .End
			End If
			LabelName = "{LB_" & LabelName & "}"
			LabelRS.Open "Select LabelName From [KS_Label] Where LabelName='" & LabelName & "'", Conn, 1, 1
			If Not LabelRS.EOF Then
			  Call KS.AlertHistory("��ǩ�����Ѿ�����!", -1)
			  LabelRS.Close
			  Conn.Close
			  Set LabelRS = Nothing
			  Set Conn = Nothing
			  Set KS = Nothing
			  .End
			Else
				LabelRS.Close
				LabelRS.Open "Select * From [KS_Label] Where (ID is Null)", Conn, 1, 3
				LabelRS.AddNew
				  Do While True
					'����ID  ��+12λ���
					LabelID = Year(Now()) & KS.MakeRandom(10)
					Set RSCheck = Conn.Execute("Select ID from [KS_Label] Where ID='" & LabelID & "'")
					 If RSCheck.EOF And RSCheck.BOF Then
					  RSCheck.Close
					  Set RSCheck = Nothing
					  Exit Do
					 End If
				  Loop
				 LabelRS("ID") = LabelID
				 LabelRS("LabelName") = LabelName
				 LabelRS("LabelContent") = LabelContent
				 LabelRS("LabelFlag") = LabelFlag
				 LabelRS("Description") = Descript
				 LabelRS("FolderID") = FolderID
				 LabelRS("AddDate") = Now
				 LabelRS("LabelType") = 1
				 LabelRS("OrderID") = 1
				 LabelRS.Update
				 Call KS.FileAssociation(1021,1,LabelContent,0)
				.Write ("<script>if (confirm('�ɹ���ʾ:\n\n��ӱ�ǩ�ɹ�,������ӱ�ǩ��?')){location.href='LabelAdd.asp?Action=AddNew&mode=text&LabelType=1&FolderID=" & FolderID & "';}else{$(parent.document).find('#BottomFrame')[0].src='" & KS.Setting(3) & KS.Setting(89) & "KS.Split.asp?LabelFolderID=" & FolderID & "&OpStr=��ǩ���� >> �Զ��徲̬��ǩ&ButtonSymbol=FreeLabel';parent.frames['MainFrame'].location.href='Label_Main.asp?LabelType=1&FolderID=" & FolderID & "';}</script>")
			End If
		Case "EditSubmit"
			LabelID = Trim(Request.Form("LabelID"))
			ParentID = Request.Form("ParentID")
			LabelName = Replace(Replace(Trim(Request.Form("LabelName")), """", ""), "'", "")
			Descript = Replace(Trim(Request.Form("Description")), "'", "")
			LabelContent = Trim(Request.Form("LabelContent"))
			LabelFlag = Request.Form("LabelFlag")
			If LabelName = "" Then
			   Call KS.AlertHistory("��ǩ���Ʋ���Ϊ��!", -1)
			   Set KS = Nothing
			   .End
			End If
			If LabelContent = "" Then
			  Call KS.AlertHistory("��ǩ���ݲ���Ϊ��!", -1)
			  Set KS = Nothing
			  .End
			End If
			LabelName = "{LB_" & LabelName & "}"
			LabelRS.Open "Select LabelName From [KS_Label] Where ID <>'" & LabelID & "' AND LabelName='" & LabelName & "'", Conn, 1, 1
			If Not LabelRS.EOF Then
			  Call KS.AlertHistory("��ǩ�����Ѿ�����!", -1)
			  LabelRS.Close:Conn.Close:Set LabelRS = Nothing:Set Conn = Nothing
			  Set KS = Nothing
			  .End
			Else
				LabelRS.Close
				LabelRS.Open "Select * From [KS_Label] Where ID='" & LabelID & "'", Conn, 1, 3
				 LabelRS("LabelName") = LabelName
				 LabelRS("LabelContent") = LabelContent
				 LabelRS("LabelFlag") = LabelFlag
				 LabelRS("Description") = Descript
				 LabelRS("AddDate") = Now
				 LabelRS("FolderID") = Request.Form("ParentID")
				 LabelRS.Update
				 '�������б�ǩ���ݣ��ҳ����б�ǩ��ͼƬ
				 Dim Node,UpFiles,RCls
				 UpFiles=LabelContent
				 if Not IsObject(Application(KS.SiteSN&"_labellist")) Then
				     Set RCls=New Refresh
				     Call Rcls.LoadLabelToCache()
					 Set Rcls=Nothing
				 End If
					 For Each Node in Application(KS.SiteSN&"_labellist").DocumentElement.SelectNodes("labellist")
					   UpFiles=UpFiles & Node.Text
					 Next
				 Call KS.FileAssociation(1021,1,UpFiles,1)
				 '������������
				 
				 If KeyWord = "" Then
					.Write ("<script>alert('�ɹ���ʾ:\n\n��ǩ�޸ĳɹ�!');$(parent.document).find('#BottomFrame')[0].src='" & KS.Setting(3) & KS.Setting(89) & "KS.Split.asp?LabelFolderID=" & ParentID & "&OpStr=��ǩ����  >> �Զ��徲̬��ǩ&ButtonSymbol=FreeLabel';location.href='Label_main.asp?Page=" & Page & "&LabelType=1&FolderID=" & ParentID & "';</script>")
				 Else
					.Write ("<script>alert('�ɹ���ʾ:\n\n��ǩ�޸ĳɹ�!');$(parent.document).find('#BottomFrame')[0].src='" & KS.Setting(3) & KS.Setting(89) & "KS.Split.asp?OpStr=��ǩ���� >> <font color=red>�����Զ��徲̬��ǩ���</font>&ButtonSymbol=FreeLabelSearch';location.href='Label_main.asp?Page=" & Page & "&LabelType=1&KeyWord=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate & "';</script>")
				 End If
			End If
		End Select
		
		.Write "<html>"
		.Write "<head>"
		.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		.Write "<title>�½���ǩ</title>"
		.Write "</head>"
		.Write "<link href=""Admin_Style.CSS"" rel=""stylesheet"">"
		.Write "<script language=""JavaScript"" src=""../../ks_inc/Common.js""></script>"
		%>
				<script language = 'JavaScript'>

		function show_ln(txt_ln,txt_main){
			var txt_ln  = document.getElementById(txt_ln);
			var txt_main  = document.getElementById(txt_main);
			txt_ln.scrollTop = txt_main.scrollTop;
			while(txt_ln.scrollTop != txt_main.scrollTop)
			{
				txt_ln.value += (i++) + '\n';
				txt_ln.scrollTop = txt_main.scrollTop;
			}
			return;
		}
		function editTab(){
			var code, sel, tmp, r
			var tabs=''
			event.returnValue = false
			sel =event.srcElement.document.selection.createRange()
			r = event.srcElement.createTextRange()
			switch (event.keyCode){
				case (8) :
				if (!(sel.getClientRects().length > 1)){
					event.returnValue = true
					return
				}
				code = sel.text
				tmp = sel.duplicate()
				tmp.moveToPoint(r.getBoundingClientRect().left, sel.getClientRects()[0].top)
				sel.setEndPoint('startToStart', tmp)
				sel.text = sel.text.replace(/\t/gm, '')
				code = code.replace(/\t/gm, '').replace(/\r\n/g, '\r')
				r.findText(code)
				r.select()
				break
			case (9) :
				if (sel.getClientRects().length > 1){
					code = sel.text
					tmp = sel.duplicate()
					tmp.moveToPoint(r.getBoundingClientRect().left, sel.getClientRects()[0].top)
					sel.setEndPoint('startToStart', tmp)
					sel.text = '\t'+sel.text.replace(/\r\n/g, '\r\t')
					code = code.replace(/\r\n/g, '\r\t')
					r.findText(code)
					r.select()
				}else{
					sel.text = '\t'
					sel.select()
				}
				break
			case (13) :
				tmp = sel.duplicate()
				for (var i=0; tmp.text.match(/[\t]+/g) && i<tmp.text.match(/[\t]+/g)[0].length; i++) tabs += '\t'
				sel.text = '\r\n'+tabs
				sel.select()
				break
			default  :
				event.returnValue = true
				break
				}
			}
			
		//-->
		</script>
		<%
		.Write "<script>"
		Response.Write "function LabelInsertCode(Val)" & vbcrlf
		Response.Write "{"
		Response.Write " if (Val==null)" & vbcrlf
		Response.Write "  Val=OpenWindow('LabelFrame.asp?sChannelID=0&TemplateType=0&url=InsertLabel.asp&pagetitle='+escape('�����ǩ'),260,350,window);"&vbcrlf
		Response.Write "if (Val!='')"
		Response.Write "{ document.LabelForm.LabelContent.focus();" & vbcrlf
		Response.Write "  var str = document.selection.createRange();" & vbcrlf
		Response.Write "  str.text = Val;"
		Response.Write " }" & vbcrlf
		Response.Write "}" & vbcrlf
		Response.Write "function InsertFunctionLabel(Url,Width,Height)" & vbcrlf
        Response.Write "{" & vbcrlf
        Response.Write "var Val = OpenWindow(Url,Width,Height,window);"
		Response.Write "if (Val!=''&&Val!=null)"
		Response.Write "{ document.LabelForm.LabelContent.focus();" & vbcrlf
		Response.Write "  var str = document.selection.createRange();" & vbcrlf
		Response.Write "  str.text = Val;"
		Response.Write " }" & vbcrlf
        Response.Write "}" & vbcrlf
		.Write "</script>"
		.Write "<body scroll=no leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
		.Write "<table width=""100%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		.Write "  <form name=""LabelForm"" id=""LabelForm"" method=post action="""" onSubmit=""return(CheckForm())"">"
		.Write "    <input type=""hidden"" name=""LabelFlag"" value=""3"">"
		.Write "    <input type=""hidden"" name=""LabelID"" value=""" & LabelID & """>"
		.Write "    <input type=""hidden"" name=""FolderID"" value=""" & FolderID & """>"
		.Write "    <input type=""hidden"" name=""Page"" value=""" & Page & """>"
			
			If Action = "AddNew" Or Action = "" Then .Write "<input type='hidden' name='Action' value='AddNewSubmit'>"
			If Action = "EditLabel" Then .Write "<input type='hidden' name='Action' value='EditSubmit'>"
			
		   .Write " <tr>"
		   .Write "   <td height=""25"" colspan=""2""> "
		 .Write "<table width='100%' height='25' border='0' cellpadding='0' cellspacing='1' bgcolor='#efefef' class='sort'>"
		 .Write "<tr><td><div align='center'><font color='#990000'>"
		  If Action = "EditLabel" Then
		   .Write "�޸��Զ��徲̬��ǩ"
		   Else
		   .Write "�½��Զ��徲̬��ǩ"
		  End If
		 .Write "</font></div></td></tr>"
		.Write "</table>"
		.Write " </td>"
		.Write "    </tr>"
		.Write "    <tr class=""tableBorder1"">"
		.Write "      <td height=""19"">��ǩ����</td>"
		.Write "      <td><input value=""" & LabelName & """ name=""LabelName"" style=""width:200;"">"
		.Write "        <font color=""#FF0000""> �����ǩ���ƣ�&quot;�Ƽ������б�&quot;������ģ���е��ã�&quot;{LB_�Ƽ������б�}&quot;��ע��Ӣ�Ĵ�Сд��ȫ��ǣ���</font></td>"
		.Write "    </tr>"
		.Write "    <tr class=""tableBorder1"">"
		.Write "      <td width=""60"" height=""19""> <div align=""left"">��ǩĿ¼</div></td>"
		.Write "      <td>" & KS.ReturnLabelFolderTree(FolderID, 1) & "<font color=""#FF0000"">��ѡ���ǩ����Ŀ¼���Ա��պ�����ǩ</font></td>"
		.Write "    </tr>"
		.Write "    <tr class=""tableBorder1"">"
		.Write "      <td width=""60"" height=""16""><div align=""left"">��ǩ���</div></td>"
		.Write "      <td><textarea name=""Description"" rows=""3"" style=""width:100%;"">" & Descript & "</textarea></td>"
		.Write "    </tr>"
		.Write "    <tr><td colspan=""2"" align=""center"" height=""25"" class=""tableBorder1""><strong>�� �� �� �� ̬ �� ǩ �� ��</strong></td></tr>"

		If KS.G("Mode")="text" then
		 Response.Write "   <tr class=""tableBorder1"" height=25>"
		 Response.Write "	<td  colspan=""2"">"
		 Response.Write "    &nbsp;&nbsp;&nbsp;&nbsp;"
		 Response.Write " <select name=""mylabel"" style=""width:160px"">"
		 Response.Write " <option value="""">==ѡ��ϵͳ������ǩ==</option>"
		   Dim RS:Set RS=Server.Createobject("adodb.recordset")
		   rs.open "select LabelName from KS_Label Where LabelType<>5 order by adddate desc",conn,1,1
		   If not Rs.eof then
		    Do While Not Rs.Eof
			 Response.Write "<option value=""" & RS(0) & """>" & RS(0) & "</option>"
			 RS.MoveNext
			Loop 
		   End If
		  Response.Write "</select>&nbsp;<input type='button' onclick='LabelInsertCode(document.all.mylabel.value);' value='�����ǩ'>"
		  RS.Close:Set RS=Nothing
		 Response.Write "&nbsp;<input type=""button"" onclick=""javascript:LabelInsertCode();"" value=""ѡ������ǩ"">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</Td>"
		 Response.Write "      </Tr>"
		.Write "    <tr class=""tableBorder1""><td  height=""250"" align='right'><textarea id='txt_ln' name='rollContent' cols='6' style='overflow:hidden;height:100%;background-color:highlight;border-right:0px;text-align:right;font-family: tahoma;font-size:12px;font-weight:bold;color:highlighttext;cursor:default;' readonly>"
		 Dim N
		 For N=1 To 3000
			.Write N & "&#13;&#10;"
		 Next
		 .Write"</textarea></td><td>"
		 .Write "<textarea name='LabelContent' style='width:100%;height:100%' ROWS='15' id='txt_main' onkeydown='editTab()' onscroll=""show_ln('txt_ln','txt_main')"" wrap='on'>" & LabelContent & "</textarea>" & vbNewLine
		 .Write "	<script>for(var i=3000; i<=3000; i++) document.getElementById('txt_ln').value += i + '\n';</script></td></tr>"
		Else
		.Write "    <tr><td colspan=""2"" height=""50""><textarea style=""width:100%"" type=""hidden"" ROWS='6' onfocus='GetLabelContent();' onkeyup=""SetEditorValue();"" onblur='SetEditorValue();' COLS='108' name=""LabelContent"">" & LabelContent & "</textarea></td></tr>"
				.Write "    <tr valign=""top"">"
		.Write "      <td colspan='2'> <iframe id=""LabelEditor"" src=""../KS.Editor.asp?ID=LabelContent&style=2"" scrolling=""no"" width=""100%"" height=""300"" frameborder=""0""></iframe>"
		.Write "      </td>"
		.Write "    </tr>"
		End iF
		
		.Write "  </form>"
		.Write "</table>"
		.Write "</body>"
		.Write "</html>"
		.Write "<script language=""JavaScript"">" & vbCrLf
		.Write "<!--" & vbCrLf
		.Write "function GetLabelContent()"
		.Write "{"
		.Write "var TempLabelContent=frames[""LabelEditor""].KS_EditArea.document.body.innerHTML;"
		.Write "TempLabelContent=frames[""LabelEditor""].ReplaceUrl(TempLabelContent);"
		.Write "TempLabelContent=frames[""LabelEditor""].Resumeblank(TempLabelContent);"
		.Write "TempLabelContent=frames[""LabelEditor""].ReplaceImgToScript(TempLabelContent);"
		.Write "TempLabelContent=frames[""LabelEditor""].FormatHtml(TempLabelContent);"
		.Write "if (TempLabelContent!=document.LabelForm.LabelContent.value)document.LabelForm.LabelContent.value=TempLabelContent;"
		.Write "}"
		.Write "function SetEditorValue()"
		.Write "{var TempLabelContent=document.LabelForm.LabelContent.value;"
		.Write "TempLabelContent=frames[""LabelEditor""].ReplaceScriptToImg(TempLabelContent);"
		.Write "TempLabelContent=frames[""LabelEditor""].ReplaceRealUrl(TempLabelContent);"
		.Write  "if (TempLabelContent!=frames[""LabelEditor""].KS_EditArea.document.body.innerHTML)frames[""LabelEditor""].KS_EditArea.document.body.innerHTML=TempLabelContent;"
		.Write "}"
		.Write "function CheckForm()" & vbCrLf
		.Write "{ var form=document.LabelForm;"
		If KS.G("Mode")<>"text" then
		.Write "  if (frames[""LabelEditor""].CurrMode!='EDIT') {alert('����ģʽ���޷����棬���л������ģʽ');return false;}"
		.Write "   form.LabelContent.value=frames[""LabelEditor""].FormatHtml(frames[""LabelEditor""].ReplaceUrl(frames[""LabelEditor""].ReplaceImgToScript(frames[""LabelEditor""].Resumeblank(frames[""LabelEditor""].KS_EditArea.document.body.innerHTML))));"
		End If
		.Write "  if (form.LabelName.value=='')"
		.Write "   {"
		.Write "    alert('�������ǩ����!');"
		.Write "    form.LabelName.focus();"
		.Write "    return false;"
		.Write "   }"

		 .Write " if (form.LabelContent.value==''||form.LabelContent.value=='���������Զ����html����')"
		 .Write " {"
		 .Write "   alert('�������ǩ����!');"
		 If KS.G("Mode")<>"text" then
		 .Write "   frames[""LabelEditor""].KS_EditArea.focus();"
		 Else
		 .Write "   form.LabelContent.focus();"
		 End If
		 .Write "   return false;"
		 .Write "  }"
		 .Write "  if (form.Description.value.length>255)"
		 .Write "   {"
		 .Write "    alert(""Ŀ¼���Ʋ��ܳ���125������(255��Ӣ���ַ�)!"");"
		 .Write "    form.Description.focus();"
		 .Write "   return false;"
		 .Write "   }"
		 .Write "  form.submit();"
		 .Write "  return true;"
		.Write "}" & vbCrLf
		.Write "//-->" & vbCrLf
		.Write "</script>"
		
		Set Conn = Nothing
		
		End With
End Sub
End Class
%> 
