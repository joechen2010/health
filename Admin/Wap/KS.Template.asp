<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%'Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.FileIcon.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../include/Session.asp"-->
<%
Dim KSCls
Set KSCls = New LabelCls
KSCls.Kesion()
Set KSCls = Nothing

Class LabelCls
        Private KS,Chk,KSCls
		Private Action,MaxPerPage
		Private Sub Class_Initialize()
		    MaxPerPage = 16
			Set KS=New PublicCls
			Set Chk=New LoginCheckCls1
			Set KSCls=New ManageCls
			Call KS.DelCahe(KS.SiteSn & "_waplabellist")
		End Sub
        Private Sub Class_Terminate()
		    Call CloseConn()
		    Set KS=Nothing
		    Set Chk=Nothing
			Set KSCls=Nothing
		End Sub

		Public Sub Kesion()
		    Chk.Run()
			If Not KS.ReturnPowerResult(0, "KSO10003") Then
			   Response.Write ("<script>parent.frames['BottomFrame'].location.href='javascript:history.back();';</script>")
			   Call KS.ReturnErr(1, "")
			   Response.End()
			End If
			'On Error Resume Next
			Response.Write "<html>"
			Response.Write "<head>"
			Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			Response.Write "<title>ģ�����</title>"
			Response.Write "</head>"
			Response.Write "<link href=""../Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			
			Action=KS.G("Action")
			Select Case Action
			    Case "Del"
				Call LabelDel()
				Case "PasteSave"'��¡
				Call LabelPasteSave()
				Case "FileReName"
				Call FileReName()
				Case "AddNew","EditLabel"
				Call AddLabel()
				Case "AddSave"
				Call AddSave()
				Case "EditSave"
				Call EditSave()
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
				Call LabelList()
			End Select 
        End Sub
		
		Sub LabelList()
		    Response.Write "<script language=""JavaScript"" src=""../../" & KS.Setting(89) &"Include/Common.js""></script>"
		    Dim MaxPerPage,CurrentPage
			MaxPerPage =22
			If KS.G("page") <> "" Then
			   CurrentPage = CInt(KS.G("page"))
			Else
			   CurrentPage = 1
			End If
			%>
			<script language="javascript">
			function EditFile(ID,filename)
			   {
			   var ReturnValue='';
			   ReturnValue=prompt('Ҫ����������',filename);
			   if ((ReturnValue!='') && (ReturnValue!=null)) window.location.href='KS.Template.asp?Action=FileReName&TemplateID='+ID+'&TemplateName='+ReturnValue;
			   else if(ReturnValue!=null){alert('����дҪ����������');}
			   }
			 function Paste(ID,filename)
			   {
			   var ReturnValue='';
			   ReturnValue=prompt('����¿�¡��ҳ��ȡ������',filename);
			   if ((ReturnValue!='') && (ReturnValue!=null)) window.location.href='KS.Template.asp?Action=PasteSave&TemplateID='+ID+'&NewTemplateName='+ReturnValue;
			   else if(ReturnValue!=null){alert('����дҪ��¡ҳ�������');}
			   }
            </script>
            </head>
            <body scroll=no topmargin="0" leftmargin="0" OnClick="SelectElement();" onkeydown="GetKeyDown();">
			
<ul id='menu_top'><li class='parent' onClick="location.href='KS.Template.asp?Action=AddNew'"><span class=child onMouseOver="this.parentNode.className='parent_border'" onMouseOut="this.parentNode.className='parent'"><img src='../images/ico/a.gif' border='0' align='absmiddle'>���ҳ��</span></li><li class='parent' onClick="location.href='?Action=LabelIn&LabelType=9';"><span class=child onMouseOver="this.parentNode.className='parent_border'" onMouseOut="this.parentNode.className='parent'"><img src='../images/ico/reb.gif' border='0' align='absmiddle'>ҳ�浼��</span></li><li class='parent' onClick="location.href='?Action=LabelOut&LabelType=9';"><span class=child onMouseOver="this.parentNode.className='parent_border'" onMouseOut="this.parentNode.className='parent'"><img src='../images/ico/back.gif' border='0' align='absmiddle'>ҳ�浼��</span></li><li></li></ul>			

            <div style="height:94%; overflow: auto; width:100%" align="center">
			<%
			Dim RS,SQL
			Set RS = Server.CreateObject("ADODB.RecordSet")
			SQL = "SELECT * FROM KS_WapTemplate ORDER BY AddDate Desc" 
			RS.Open SQL, Conn, 1, 1
			%>
            <table width='100%' border='0' cellspacing='0' cellpadding='0'>
            <tr>
            <td width='110' class="sort" align='center'><font color="#990000">I D</font></td>
            <td width='513' class='sort' align='center'>�����ӵ�ַ</td>
            <td width='171' class='sort' align='center'>ҳ������</td>
            <td width='202' class='sort' align='center'>�޸�ʱ��</td>
            <td width='255' class='sort' align='center'>��������</td>
            </tr>
			<%
			If Not (RS.EOF And RS.BOF) Then
			Dim TotalPut,I
			TotalPut= RS.RecordCount
			If CurrentPage < 1 Then CurrentPage = 1
			If (CurrentPage - 1) * MaxPerPage > totalPut Then
			   If (TotalPut Mod MaxPerPage) = 0 Then
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
			Do While Not RS.EOF
			%>
            <tr onMouseOut="this.className='list'" onMouseOver="this.className='listmouseover'">
            <td height="19" align=center class='splittd'><div align="center"><span style="cursor:default"><%=RS("ID")%></span></div></td>
            <td class='splittd' align='center'>{$GetInstallDir}<%=KS.WSetting(4)%>Template.asp?ID=<%=RS("ID")%>&{$WapValue}</td>
            <td class='splittd' align='center'><%=Trim(RS("TemplateName"))%></td>
            <td class='splittd' align='center'><%=RS("AddDate")%></td>
            <td class='splittd' align='center'><a href='KS.Template.asp?Action=EditLabel&TemplateID=<%=RS("ID")%>&Flag=text' onClick="parent.frames['BottomFrame'].location.href='../../<%=KS.Setting(89)%>KS.Split.asp?OpStr=WAP�Զ���ҳ��������� >> WAPҳ��༭&ButtonSymbol=GoSave'">�༭</a> | <a href="javascript:EditFile('<%=RS("ID")%>','<%=RS("TemplateName")%>')">������</a> | <a href="javascript:Paste('<%=RS("ID")%>','����_<%=RS("TemplateName")%>')">��¡</a> | <a href="KS.Template.asp?Action=Del&TemplateID=<%=RS("ID")%>" onClick="return(confirm('�˲��������棬ȷ��ɾ����'))">ɾ��</a>
            </td>
            </tr>
			<%
			I = I + 1
			If I >= MaxPerPage Then Exit Do
			RS.MoveNext
			Loop	 
			End If
			%>
            </table>
            <table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>
            <tr>
            <td align='right'>
			<%
			Call KSCLS.ShowPage(totalPut, MaxPerPage, "KS.Template.asp", True, "��", CurrentPage, "")
			%></td></tr></table>
            <div style="text-align:center;color:#003300">-----------------------------------------------------------------------------------------------------------</div>
            <div style=""height:30px;text-align:center"">KeSion CMS V 6.5, Copyright (c) 2006-2010 <a href=""http://www.kesion.com/"" target=""_blank""><font color=#ff6600>KeSion.Com</font></a>. All Rights Reserved . </div>
            			<br/><br/>
</div>
            </body>
            </html>
			<%
        End Sub

		
		'�����¡
		Sub LabelPasteSave()
		    Dim TemplateID:TemplateID=KS.G("TemplateID")
			Dim NewTemplateName:NewTemplateName=KS.G("NewTemplateName")
			Dim LabelRS:Set LabelRS=Server.CreateObject("ADODB.RECORDSET")
			LabelRS.Open "Select TemplateName From KS_WapTemplate Where TemplateName='" & NewTemplateName & "'", Conn, 1, 1
			If Not LabelRS.Eof Then 
		       LabelRS.Close:Set LabelRS=Nothing
			   Call KS.Alert("ҳ�������Ѵ��ڣ���������������!","Label_Main.asp?Action=SetPasteParam&TemplateID=" & TemplateID)
		    End If
			LabelRS.Close
			LabelRS.Open "Select * From KS_WapTemplate Where ID='" & TemplateID & "'",Conn,1,1
			If Not LabelRS.Eof Then
			    Dim NewRS:Set NewRS=Server.CreateObject("ADODB.RECORDSET")
				NewRS.Open "Select * From KS_WapTemplate",Conn,1,3
				NewRS.AddNew
				NewRS("ID")        = Year(Now()) & KS.MakeRandom(10)
				NewRS("TemplateName") = NewTemplateName
				NewRS("TemplateContent") = LabelRS("TemplateContent")
				NewRS("AddDate")     = Now
				NewRS.Update
				NewRS.Close:Set NewRS=Nothing
				LabelRS.Close:Set LabelRS=Nothing
				Response.write "<script>window.location='KS.Template.asp';</script>"
			Else
			   Response.Write "<script>alert('��¡ʧ��!');window.close();</script>"
			End If
		End Sub
		
		'ɾ��ģ��
		Sub LabelDel()
		    Dim TemplateID:TemplateID=KS.G("TemplateID")
			Conn.Execute("Delete from KS_WapTemplate where ID='" & TemplateID & "'")
			Response.write "<script>window.alert('ɾ���ɹ�');window.location='KS.Template.asp';</script>"
		End Sub
		
		'������
		Sub FileReName()
		    Dim TemplateID,TemplateName
			TemplateID=KS.G("TemplateID")
			TemplateName = KS.G("TemplateName")
			Conn.Execute("Update KS_WapTemplate set TemplateName='" & TemplateName & "' where ID='" & TemplateID & "'")
			Response.write "<script>window.alert('�������ɹ�');window.location='KS.Template.asp';</script>"
		End Sub
		
		Sub AddSave()
		    Dim RS,RSCheck,TemplateID,TemplateName,TemplateContent
			TemplateName = Replace(Replace(Trim(Request.Form("TemplateName")), """", ""), "'", "")
			TemplateContent = Trim(Request.Form("TemplateContent"))
			If TemplateName="" Then
			   Response.write "<script>window.alert('����û����д..');window.location='javascript:history.go(-1)';</script>"
			   Response.End
			End if
			If TemplateContent="" Then
			   Response.write "<script>window.alert('����ʧ��...');window.location='javascript:history.go(-1)';</script>"
			   Response.End
			End if
			Set RS = Server.CreateObject("ADODB.RecordSet")
			RS.Open "Select * From [KS_WapTemplate] Where (ID is Null)",Conn,1,3     
			RS.Addnew
			Do While True
			   '����ID  ��+12λ���
			   TemplateID = Year(Now()) & KS.MakeRandom(10)
			   Set RSCheck = Conn.Execute("Select ID from [KS_WapTemplate] Where ID='" & TemplateID & "'")
			   If RSCheck.EOF And RSCheck.BOF Then
			      RSCheck.Close
				  Set RSCheck = Nothing
				  Exit Do
			   End If
			Loop
			RS("ID")=Year(Now()) & KS.MakeRandom(10)
			RS("TemplateName")=TemplateName
			RS("TemplateContent")=TemplateContent
			RS("AddDate")=Now
			RS.Update
			RS.Close:set RS=Nothing
			Response.Write ("<script>if (confirm('�ɹ���ʾ:\n\n����Զ���ҳ��ɹ�,������ӱ�ǩ��?')){location.href='KS.Template.asp?Action=AddNew';}else{parent.frames['BottomFrame'].location.href='../../"&KS.Setting(89)&"KS.Split.asp?OpStr=WAP�Զ���ҳ��������� >> WAPҳ�����&ButtonSymbol=FreeLabel';parent.frames['MainFrame'].location.href='KS.Template.asp';}</script>")
		End Sub
		
		Sub EditSave()
		    Dim RS
			Dim TemplateID,TemplateName,TemplateContent
			TemplateID=KS.G("TemplateID")
			TemplateName = Replace(Replace(Trim(Request.Form("TemplateName")), """", ""), "'", "")
			TemplateContent = Trim(Request.Form("TemplateContent"))
			If TemplateName="" then
			   Response.write "<script>window.alert('����û����д..');window.location='javascript:history.go(-1)';</script>"
			   Response.End
			End if
			If TemplateContent="" then
			   Response.write "<script>window.alert('����ʧ��...');window.location='javascript:history.go(-1)';</script>"
			   Response.End
			End if
			Set RS = Server.CreateObject("ADODB.RecordSet")
			RS.Open "select * from KS_WapTemplate where ID='" & TemplateID & "'",Conn,1,3
			RS("TemplateName")=TemplateName
			RS("TemplateContent")=TemplateContent
			RS("AddDate")=Now
			RS.Update
			RS.Close:set RS=Nothing
			Response.Write "<script>alert('�ɹ���ʾ:\n\n�Զ���ҳ���޸ĳɹ�!');parent.frames['BottomFrame'].location.href='../../"&KS.Setting(89)&"KS.Split.asp?OpStr=WAP�Զ���ҳ���������  >> WAPҳ�����&ButtonSymbol=FreeLabel';location.href='KS.Template.asp';</script>"
		End Sub
		
		'����ģ��
		Sub AddLabel()
		    Dim TemplateName,TemplateContent
			If KS.G("Action") = "EditLabel" Then
			   Dim RS,TemplateID
			   TemplateID=KS.G("TemplateID")
			   Set RS = Conn.Execute("select * from KS_WapTemplate where ID='" & TemplateID & "'")
			   TemplateName=RS("TemplateName")
			   TemplateContent=RS("TemplateContent")
			   RS.Close:set RS=Nothing
			Else
			   TemplateName="���ڴ˴�����WAPҳ������"
			   'TemplateContent="<?xml version=""1.0"" encoding=""utf-8""?>" & vbCrLf
			   'TemplateContent=TemplateContent&"<!DOCTYPE wml PUBLIC ""-//WAPFORUM//DTD WML 1.1//EN"" ""http://www.wapforum.org/DTD/wml1_1.1.xml"">" & vbCrLf
			   TemplateContent=TemplateContent&"<wml>" & vbCrLf
			   TemplateContent=TemplateContent&"<head>" & vbCrLf
			   TemplateContent=TemplateContent&"<meta http-equiv=""Cache-Control"" content=""max-age=0""/>" & vbCrLf
			   TemplateContent=TemplateContent&"<meta http-equiv=""Cache-Control"" content=""no-cache""/>" & vbCrLf
			   TemplateContent=TemplateContent&"</head>" & vbCrLf
			   TemplateContent=TemplateContent&"<card title=""ҳ������"">" & vbCrLf
			   TemplateContent=TemplateContent&"<p>" & vbCrLf
			   TemplateContent=TemplateContent&"���ڴ˴�����ģ�����" & vbCrLf
			   TemplateContent=TemplateContent&"--------<br/>" & vbCrLf
			   TemplateContent=TemplateContent&"{$GetGoBack}<br/><br/>" & vbCrLf
			   TemplateContent=TemplateContent&"{$GetGoBackIndex}" & vbCrLf
			   TemplateContent=TemplateContent&"</p>" & vbCrLf
			   TemplateContent=TemplateContent&"</card>" & vbCrLf
			   TemplateContent=TemplateContent&"</wml>"
			End if
			%>
<script language="JavaScript" src="../../ks_inc/Common.js"></script>
<script>
function LabelInsertCode(Val)
{
if (Val==null)
	  Val=OpenWindow('../Include/LabelFrame.asp?sChannelID=&TemplateType=&url=../Wap/InsertLabel.asp&pagetitle='+escape('�����ǩ'),260,350,window);
if (Val!=''){ document.LabelForm.TemplateContent.focus();
  var str = document.selection.createRange();
  str.text = Val; }
}
function InsertFunctionLabel(Url,Width,Height)
{
var Val = OpenWindow(Url,Width,Height,window);if (Val!=''&&Val!=null){ document.LabelForm.TemplateContent.focus();
  var str = document.selection.createRange();
  str.text = Val; }
}
</script>
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
</head>       
<body scroll=no leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<table width='100%' height='25' border='0' cellpadding='0' cellspacing='1' bgcolor='#efefef'>
<tr><td class='sort'><div align='center'><font color='#990000'>
      <%
	  If Action = "EditLabel" Then
	     Response.Write "�޸�ҳ��"
	  Else
	     Response.Write "�½�ҳ��"
	  End If
	  %></font></div></td></tr>
</table>

<table width='100%' height="350" style="background-color:#EEEEEE;padding-right: 2px;padding-left: 2px;padding-bottom: 0px;" border='0' align='center' cellpadding='0' cellspacing='0' class='ctable'>
<form name="LabelForm" method="post" Action="KS.Template.asp" onSubmit="return(CheckForm())">
<input type="hidden" name="TemplateID" value="<%=TemplateID%>">
<input type="hidden" name="Page" value="<%=KS.S("Page")%>">
<%
If KS.G("Action") = "AddNew" Or Action = "" Then Response.Write "<input type='hidden' name='Action' value='AddSave'>"
If KS.G("Action") = "EditLabel" Then Response.Write "<input type='hidden' name='Action' value='EditSave'>"
%>
<tr class="clefttitle">
<td height="30"><b>ҳ�����ƣ�</b><input name="TemplateName" type="text" id="TemplateName" size="50" Value="<%=TemplateName%>">  </td>
</tr>

<tr id="toplabelarea" class="clefttitle">
<td valign="top">
<strong>�����ǩ��</strong>
<select name="mylabel" style="width:160px">
<option value="">==ѡ���ñ�ǩ==</option>
<option value="[KS_Charge]">[KS_Charge]</option>
<option value="[/KS_Charge]">[/KS_Charge]</option>
<option value="{$WapValue}">��ʾ�û�����</option>
<option value="{$GetReadMessage}">��ʾδ������Ϣ</option>
<option value="{$GetTopUserLogin}">��ʾ��Ա��¼���(����)</option>
<option value="{$GetUserLogin}">��ʾ��Ա��¼���(����)</option>
<option value="{$GetOnlineTotal}">��ʾ����������</option>
<option value="{$GetOnlineUser}">��ʾ�����û�����</option>
<option value="{$GetOnlineGuest}">��ʾ�����ο�����</option>
<option value="{$GetCopyRight}">��ʾ��Ȩ��Ϣ</option>
<option value="{$GetWebmaster}">��ʾվ��</option>
<option value="{$GetWebmasterEmail}">��ʾվ��EMail</option>
<option value="{$GetGoBack}">��ʾ�����ϼ�</option>
<option value="{$GetGoBackIndex}">��ʾ������ҳ</option>
</select>&nbsp;
<input class='button' type='button' onclick='LabelInsertCode(document.all.mylabel.value);' value='�����ǩ'>&nbsp;
<input type="button" class='button' onClick="javascript:LabelInsertCode();" value="WAP��ǩ">&nbsp;
&nbsp;��ʹ�������WML�﷨�����ݿ���ʹ��[KS_Charge]ע���Ա�ſɲ鿴������[/KS_Charge]
</td>
</tr>

<tr id="codearea">
<td>
         <table border='0' cellspacing='0' cellspadding='0'>
         <tr>
         <td valign="top" width='20'>
         <textarea name="txt_ln" id="txt_ln" cols="6" style="overflow:hidden;height:410;background-color:highlight;border-right:0px;text-align:right;font-family: tahoma;font-size:12px;font-weight:bold;color:highlighttext;cursor:default;" readonly=""><%
		 Dim N
		 For N=1 To 500
		 Response.Write N & vbcrlf
		 Next
		 %>
         </textarea>
         </td>
         <td valign="top">
         <textarea name="TemplateContent" rows="2" cols="30" id="TemplateContent" onscroll="show_ln('txt_ln','TemplateContent')" onKeyDown="editTab()" style="height:410px;width:770;"><%=TemplateContent%></textarea>
         <script>for(var i=500; i<=500; i++) document.getElementById('txt_ln').value += i + '\n';</script>
         </td>
         </tr>
         </table>

</form>
</table>

</body>
</html>

<script language="JavaScript">
<!--
function CheckForm()
   {
   var form=document.LabelForm;
   if (form.TemplateName.value=='')
      {
	  alert('�������Զ���ҳ������!');form.TemplateName.focus();
	  return false;
	  }
   if (form.TemplateContent.value==''||form.TemplateContent.value=='���������Զ����WML����')
      {
	  alert('�������Զ���ҳ������!');
	  form.TemplateContent.focus();
	  return false;
	  }
	  form.submit();
	  return true;
   }
//-->
</script>
<%
		End Sub
		
		Sub LabelOut()
		    Response.Write "<body>"
			Response.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0""   class=""sortbutton"">"
			Response.Write "  <tr>"
			Response.Write "    <td height=""23"" align=""left"">��������<a href='?Action=LabelIn'>��ǩ����</a> | <a href='?Action=LabelOut'>��������</a></td>"
			Response.Write " </tr>"
			Response.Write "</table>"
			%>
            <form name='myform' method='post' action='KS.Template.asp'>  
            <table width='100%' border='0' align='center' cellpadding='2' cellspacing='0' class='border'>    
            <tr class='title'>       
            <td colspan="2" height='22' align='center'><strong>��ǩ����</strong></td>    
            </tr>    

            <tr class='tdbg'>      
            <td colspan="2" align='center'>        
		    <table width="100%" border='0' cellpadding='0' cellspacing='0'>          
			   <tr>           
			     <td width="10%" align="right">��ǩ�б�</td>
				 <td width="54%" ID="ClassArea">
                 <select name='TemplateID' size='2' multiple style='height:300px;width:450px;'>
                 <%=GetLabelOption(Conn)%>
				 </select>
                 </td>
                 <td width="36%" align='left'>&nbsp;&nbsp;&nbsp;&nbsp;
                 <input type='button' name='Submit' value=' ѡ������ ' onclick='SelectAll()'><br><br>&nbsp;&nbsp;&nbsp;&nbsp;<input type='button' name='Submit' value=' ȡ��ѡ�� ' onclick='UnSelectAll()'><br><br><br><b>&nbsp;��ʾ����ס��Ctrl����Shift�������Զ�ѡ</b>
                 </td>      
               </tr>     
               <tr height='30'>
                 <td colspan='2'>
                 Ŀ�����ݿ⣺
                 <input name='TemplateMdb' type='text' id='TemplateMdb' value='<%=KS.Setting(3)%>WapTemplate.mdb' size='20' maxlength='50'>
                 &nbsp;&nbsp;�˲��������Ŀ�����ݿ�</td>      
                 </tr>      
                 <tr height='50'>
                 <td colspan='2' align='center'>
                 <input type='submit' name='Submit' value='ִ�е�������' onClick="document.myform.Action.value='Doexport';">              <input name='Action' type='hidden' id='Action' value='Doexport'>         </td>        </tr>    </table>   
		    </td> </tr></table>
            </form>
			<script language='javascript'>
			function SelectAll()
			{
			   for(var i=0;i<document.myform.TemplateID.length;i++){
			   document.myform.TemplateID.options[i].selected=true;}
			   }
			function UnSelectAll()
			{
			   for(var i=0;i<document.myform.TemplateID.length;i++){
			   document.myform.TemplateID.options[i].selected=false;}
			   }
            </script>
		    <%
		End Sub
		
		Function GetLabelOption(DBC)
		    Dim AllLabel,RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select * From KS_WapTemplate",DBC,1,1
			Do While Not RS.Eof 
			   AllLabel=AllLabel & "<option value='" & RS("ID") & "'>" & RS("TemplateName") & "</option>"
			   RS.MoveNext
		    Loop
			RS.Close:Set RS=Nothing
			GetLabelOption=AllLabel
		End Function
		
		'��������
		Sub Doexport()
		    Dim TemplateID:TemplateID="'"& Replace(Replace(KS.G("TemplateID")," ",""),",","','") & "'"
			Dim TemplateMdb:TemplateMdb=KS.G("TemplateMdb")
			Dim RS:set RS=server.createobject("adodb.recordset")
			Dim sqlstr,n
			n=0
			sqlstr="select ID,TemplateName,TemplateContent,AddDate from KS_WapTemplate where id in(" & TemplateID & ")"
			'on error resume next
			If CreateDatabase(TemplateMdb)=True Then
			   Dim DataConn:Set DataConn = Server.CreateObject("ADODB.Connection")
			   DataConn.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(TemplateMdb)
			   If not Err Then
			      If Checktable("KS_WapTemplate",DataConn)=True Then
				     DataConn.Execute("drop table KS_WapTemplate")
				  End If
				  Dataconn.Execute("CREATE TABLE [KS_WapTemplate] ([TemplateID] int IDENTITY (1, 1) NOT NULL CONSTRAINT PrimaryKey PRIMARY KEY,[ID] varchar(50) Not Null,[TemplateName] varchar(255) Not Null,[TemplateContent] text not null,[AddDate] date not null)")
				  RS.Open sqlstr,Conn,1,1
				  If not RS.EOF Then
				     Dim RST:Set RST=Server.CreateObject("ADODB.RECORDSET")
					 do while not RS.EOF
					    n=n+1
						RST.Open "Select * From KS_WapTemplate where 1=0",DataConn,1,3
						RST.AddNew
						RST("ID")=RS(0)
						RST("TemplateName")=RS(1)
						RST("TemplateContent")=RS(2)
						RST("AddDate")=RS(3)
						RST.Update
						RST.Close
						RS.Movenext
					loop
					Set RST=Nothing
				  End If
				  RS.Close:set RS=Nothing
			   End if
			   DataConn.Close:Set DataConn=Nothing
			End If
			Response.Write "<br><br><br><div align=center>�������!�ɹ������� <font color=red>" & n & "</font> ����ǩ��<a href=" & TemplateMdb & ">������������</a>(�Ҽ�Ŀ�����Ϊ)  </div><br><br><br><br><br><br><br>"
		End Sub

		Sub LabelIn()
		    Response.Write "<body>"
			Response.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0""   class=""sortbutton"">"
			Response.Write "  <tr>"
			Response.Write "    <td height=""23"" align=""left"">��������<a href='?Action=LabelIn'>��ǩ����</a> | <a href='?Action=LabelOut'>��������</a></td>"
			Response.Write " </tr>"
			Response.Write "</table>"
			%>
            <form name='myform' method='post' action='KS.Template.asp'>
            <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>
            <tr class='title'>
            <td height='22' align='center'><strong>��ǩ���루��һ����</strong></td>
            </tr>
            <tr class='tdbg'>
            <td height='100'>&nbsp;&nbsp;&nbsp;&nbsp;������Ҫ����ı�ǩ���ݿ���ļ�����         <input name='TemplateMdb' type='text' id='TemplateMdb' value='<%=KS.Setting(3)%>WapTemplate.mdb' size='20' maxlength='50'>        <input name='Submit' type='submit' id='Submit' value=' ��һ�� '>        <input name='Action' type='hidden' id='Action' value='LabelIn2'>      </td>    </tr>  </table></form>
		<%
		End Sub


		Sub LabelIn2()
		    Response.Write "<body>"
			Response.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0""   class=""sortbutton"">"
			Response.Write "  <tr>"
			Response.Write "    <td height=""23"" align=""left"">��������<a href='?Action=LabelIn'>��ǩ����</a> | <a href='?Action=LabelOut'>��������</a></td>"
			Response.Write " </tr>"
			Response.Write "</table>"
			On Error Resume Next
			Dim TemplateMdb:TemplateMdb=KS.G("TemplateMdb")
			Dim DataConn:Set DataConn = Server.CreateObject("ADODB.Connection")
			DataConn.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(TemplateMdb)
			%>
            <form name='myform' method='post' action='KS.Template.asp'>
            <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>
            <tr class='title'>
            <td height='22' align='center'><strong>��ǩ���루�ڶ�����</strong></td>
            </tr>
            <tr class='tdbg'>
            <td height='100' align='center'>
            <br>
            <table border='0' cellspacing='0' cellpadding='0'>          
			<%
			If Err Then 
			   Err.Clear:Set DataConn = Nothing:Response.Write "<tr><td>���ݿ�·������ȷ�����ӳ���</td></tr>":Response.End
		    Else
		 	%>
            <tr>
            <td><strong>��������ʽ��</strong> 
            <input type="radio" value="0" name="Cl" checked>��ǩ��������
            <input type="radio" value="1" name="Cl">��ǩ��������
            </td>
            </tr>  
            <tr>
            <td> 
            <select name='TemplateID' size='2' multiple style='height:300px;width:350px;'>
			<%=GetLabelOption(DataConn)%>
            </select>
            </td>
            </tr>  
			<%
			End If
			%>
            <tr><td colspan='3' height='5'></td></tr>
            <tr><td height='25' align='center'><b> ��ʾ����ס��Ctrl����Shift�������Զ�ѡ</b></td></tr>
            <tr>
            <td colspan='3' height='25' align='center'>
            <input type='submit' name='Submit' value=' �����ǩ ' onClick="document.myform.Action.value='Doimport';" >
            </td>
            </tr>
            </table>
            <input name='TemplateMdb' type='hidden' id='TemplateMdb' value='<%=TemplateMdb%>'>               <input name='Action' type='hidden' id='Action' value='Doimport'>               <br>            </td>          </tr>       
		</table></form>

		    <%
		    DataConn.Close:set DataConn=Nothing
		End Sub
		'�������
		Sub Doimport()
			'on error resume next
			Dim n:n=0
			Dim m:m=0
			Dim k:k=0
			Dim TemplateMdb:TemplateMdb=KS.G("TemplateMdb")
			Dim Cl:Cl=KS.G("Cl")
			Dim TemplateID:TemplateID="'"& Replace(Replace(KS.G("TemplateID")," ",""),",","','")& "'"
			Dim DataConn:Set DataConn = Server.CreateObject("ADODB.Connection")
			DataConn.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(TemplateMdb)
			If Err Then 
			   Err.Clear:Set DataConn = Nothing:Response.Write "<tr><td>���ݿ�·������ȷ�����ӳ���</td></tr>":Response.End
			Else
			   Dim RS:set RS=server.createobject("adodb.recordset")
			   RS.Open "select * from KS_WapTemplate where ID in(" & TemplateID & ")",DataConn,1,1
			   Dim RSA:set RSA=server.createobject("adodb.recordset")
			   do while not RS.EOF 
			      RSA.Open "select * from KS_WapTemplate where TemplateName='" & RS("TemplateName") & "'",Conn,1,3
				  If RSA.EOF Then
				     RSA.Addnew
					 RSA("ID")=RS("ID")
					 RSA("TemplateName")=RS("TemplateName")
					 RSA("TemplateContent")=RS("TemplateContent")
					 RSA("AddDate")=RS("AddDate")
					 n=n+1
					 RSA.Update
				  Else   '��������
				     If Cl="1" Then
					    RSA("TemplateContent")=RS("TemplateContent")
						RSA("AddDate")=RS("AddDate")
						m=m+1
						RSA.Update
					 Else
					 k=K+1
					 End If
				  End If
				  RSA.Close
				  RS.Movenext
			   loop
			   RS.Close:set RS=Nothing
			   set RSA=nothing
			End If
			Response.Write "<br><br><br><div align=center>�������!�ɹ������� <font color=red>" & n & "</font> ����ǩ,������ <font color=red>" & m & "</font> ����ǩ,���������� <font color=red>" & k & "</font> ����ǩ��  </div><br><br><br><br><br><br><br>"
            DataConn.Close:set DataConn=Nothing
		End Sub

		Function CreateDatabase(DBName)
		    If KS.CheckFile(DBName) Then CreateDatabase=True:Exit Function
			Dim objcreate :set objcreate=Server.CreateObject("adox.catalog") 
			If Err.Number<>0 Then 
			   set objcreate=Nothing 
			   CreateDatabase=False
			   Exit Function 
			End If 
			'�������ݿ� 
			objcreate.create("data source="+Server.Mappath(dbname)+";provider=microsoft.jet.oledb.4.0") 
			If Err.number<>0 Then 
			   CreateDatabase=False
			   set objcreate=Nothing 
			   Exit Function
			End If 
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