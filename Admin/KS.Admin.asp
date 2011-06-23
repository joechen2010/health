<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../Plus/md5.asp"-->
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
Set KSCls = New User_AdminMain
KSCls.Kesion()
Set KSCls = Nothing

Class User_AdminMain
        Private KS,UserName
		Private GroupID, I, SqlStr, RSObj,Title, CreateDate, TempStr, GRS,KeyWord, SearchType
		Private PowerRS,RS,AdminID,PowerList,SpecialPower,CollectPower,SystemPower,RefreshPower,UserAdminPower,KMTemplatePower,ModelPower

		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		 Select Case KS.G("Action")
		   Case "Add","Edit"
			 If Not KS.ReturnPowerResult(0, "KMUA10001") Then                '������Ա����(���͸�)��Ȩ�޼��
			  Call KS.ReturnErr(0, "")
			  Exit Sub
			 Else 
			  Call AdminAdd()
		     End If
		   Case "Del"
		       If Not KS.ReturnPowerResult(0, "KMUA10001") Then                '������Ա����(���͸�)��Ȩ�޼��
				  Call KS.ReturnErr(0, "")
				  Exit Sub
			   Else
			   Call AdminDel()
		       End If
		   Case "SetPass"
		   	 If Not KS.ReturnPowerResult(0, "KMUA10010") Then           
			  Call KS.ReturnErr(0, "")
			  Exit Sub
		     Else
		      Call SetAdminPass()
			 End If
		   Case Else
		     Call AdminList()
		 End Select
		End Sub
		
		Sub AdminList()
		  
		'�ռ���������
		KeyWord = KS.G("KeyWord")
		SearchType = KS.G("SearchType")
		'������������
		Dim SearchParam:SearchParam = "KeyWord=" & KeyWord & "&SearchType=" & SearchType
		Const Row = 8 'ÿ����ʾ��
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; chaRSet=gb2312"">"
		Response.Write "<title>����Ա����</title>"
		Response.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
		Response.Write "<script language=""JavaScript"">" & vbCrLf
		Response.Write "var GroupID='0';        //����Ա��ID" & vbCrLf
		Response.Write "var KeyWord='" & KeyWord & "';         //�����ؼ���" & vbCrLf
		Response.Write "var SearchParam='" & SearchParam & "'; //������������" & vbCrLf
		Response.Write "</script>" & vbCrLf
		Response.Write "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>"
		Response.Write "<script language=""JavaScript"" src=""../KS_Inc/jquery.js""></script>"
		Response.Write "<script language=""JavaScript"" src=""Include/ContextMenu1.js""></script>"
		Response.Write "<script language=""JavaScript"" src=""Include/SelectElement.js""></script>"
		%>
		<script language="javascript">
		var DocElementArrInitialFlag=false;
		var DocElementArr = new Array();
		var DocMenuArr=new Array();
		var SelectedFile='',SelectedFolder='';
		function document.onreadystatechange()
		{   
		    parent.frames['BottomFrame'].Button1.disabled=true;
			parent.frames['BottomFrame'].Button2.disabled=true;
		    if (DocElementArrInitialFlag) return;
			InitialDocElementArr('GroupID','AdminID');
			InitialDocMenuArr();
			DocElementArrInitialFlag=true;
		}
		function InitialDocMenuArr()
		{      DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.Create();",'�� ��(N)','disabled');
			   DocMenuArr[DocMenuArr.length]=new ContextMenuItem("seperator",'','');
			   DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.Edit();",'�� ��(E)','disabled');
			   DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.Delete('');",'ɾ ��(D)','disabled');
			   DocMenuArr[DocMenuArr.length]=new ContextMenuItem("seperator",'','');
			   DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.SetAdminPassWord();",'��������(P)','disabled');
			   DocMenuArr[DocMenuArr.length]=new ContextMenuItem("seperator",'','');
			   DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.Reload();",'ˢ ��(Z)','');
		}
		function DocDisabledContextMenu()
		{
		   var TempDisabledStr=''; 
			DisabledContextMenu('GroupID','AdminID',TempDisabledStr+'����Ȩ��(S),��������(P),�� ��(E),ɾ ��(D)','����Ȩ��(S),��������(P),�� ��(E)','','����Ȩ��(S),��������(P),�� ��(E)','����Ȩ��(S),��������(P),�� ��(E)','')
		}
		function CreateAdmin()
		{
		 location.href='KS.Admin.asp?Action=Add';
		 $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Go&OpStr=�û����� >> <font color=red>��ӹ���Ա</font>';
		}
		function EditAdmin(AdminID)
		{
		 location.href='KS.Admin.asp?Action=Edit&AdminID='+AdminID;
		 $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=GoSave&OpStr=�û����� >> <font color=red>�޸Ĺ���Ա</font>';
		}
		function Create()
		{
		 CreateAdmin();  
		}
		function Edit()
		{  GetSelectStatus('GroupID','AdminID');
				if (SelectedFile!='')
				{
				 if (SelectedFile.indexOf(',')==-1) 
				  EditAdmin(SelectedFile);
				 else alert('һ��ֻ�ܹ��༭һ������Ա!'); 
				}
			else 
			   alert('��ѡ��Ҫ�༭�Ĺ���Ա!');
		}
		function Delete(id)
		{  
		 if (id==''){
		  GetSelectStatus('GroupID','AdminID');  
		 }else{
		  SelectedFile=id;
		 }
			if (SelectedFile!='')
				{ if (confirm('ȷ��ɾ��ѡ�й���Ա��?'))location='KS.Admin.asp?'+SearchParam+'&Action=Del&AdminID='+SelectedFile;}
				else alert('��ѡ��Ҫɾ���Ĺ���Ա');
		}
		function SetAdminPassWord()
		{
		 GetSelectStatus('GroupID','AdminID');
		 if (SelectedFile!='')
				   if (SelectedFile.indexOf(',')==-1) 
					 { 
					 OpenWindow('KS.Frame.asp?Url=KS.Admin.asp&Action=SetPass&PageTitle='+escape('���ù���Ա����')+'&AdminID='+SelectedFile,360,160,window);
					 SelectedFile='';
					 }
				 else alert('һ��ֻ�ܸ�һ������Ա��������!'); 
		 else
		  alert('��ѡ��Ҫ��������Ĺ���Ա!')
		}
		function GetKeyDown()
		{
		if (event.ctrlKey)
		  switch  (event.keyCode)
		  {  case 90 :  Reload(); break;
			 case 78 : event.keyCode=0;event.returnValue=false;
				 CreateAdmin('');
			 case 80 :SetAdminPassWord();break;
			 case 69 : event.keyCode=0;event.returnValue=false;Edit(); break;
			 case 68 : Delete('');break;
			 case 70 : event.keyCode=0;event.returnValue=false;
				parent.frames['LeftFrame'].initializeSearch('Manager')
		   }	
		else	
		 if (event.keyCode==46)Delete('');
		}
		function Reload()
		{
		location.href='KS.Admin.asp?'+SearchParam+'&GroupID='+GroupID;
		}
		</script>
		<%
		Response.Write "</head>"
		Response.Write "<body scroll=no topmargin=""0"" leftmargin=""0"" OnClick=""SelectElement();"" onkeydown=""GetKeyDown();"" onselectstart=""return false;"">"
		Response.Write "<ul id='menu_top'>"
			  If KeyWord = "" Then
			   Response.Write "<li class='parent' onclick=""CreateAdmin();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>��ӹ���Ա</span></li>"				 
			   Response.Write "<li class='parent' onclick=""SetAdminPassWord();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/unverify.gif' border='0' align='absmiddle'>��������</span></li>"				 
			   Response.Write "<li class='parent' onclick=""parent.frames['LeftFrame'].initializeSearch('����Ա');""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/addjs.gif' border='0' align='absmiddle'>��������Ա</span></li>"				 

			  Else
				   Response.Write ("<img src='Images/home.gif' align='absmiddle'><span style='cursor:pointer' onclick=""SendFrameInfo('KS.Admin.asp','Admin_UserLeft.asp','KS.Split.asp?ButtonSymbol=Disabled&OpStr=����Ա���� >> <font color=red>������ҳ</font>')"">����Ա��ҳ</span>")
			   Response.Write (">>> �������: ")
				Select Case SearchType
				 Case 0
				  Response.Write ("�û������� <font color=red>" & KeyWord & "</font> �Ĺ���Ա")
				 Case 1
				  Response.Write ("����Ա��麬�� <font color=red>" & KeyWord & "</font> �Ĺ���Ա")
				 End Select
			   End If

		Response.Write "</ul>"
		Response.Write ("<div style=""height:94%; overflow: auto; width:100%"" align=""center"">")
		Response.Write "<table width=""100%"" height=""25"" border=""0"" cellpadding=""0"" cellspacing=""1"">"
		
			  Dim Param:Param = " Where 1=1"
			  If KeyWord <> "" Then
				Select Case SearchType
				  Case 0
				   Param = Param & " And UserName like '%" & KeyWord & "%'"
				  Case 1
				   Param = Param & " And Description like '%" & KeyWord & "%'"
				End Select
			   Else
				 Param = Param
			   End If
			  Param = Param & " Order BY SuperTF Desc,AddDate desc"
			  SqlStr = "Select * From KS_Admin " & Param
				 Response.Write ("<tr> ")
				 Response.Write ("  <td>")
				 Response.Write ("    <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">")
				 Response.Write ("<tr align=""center""><td height=23 width=100  class=""sort"">�� �� Ա</td><td width=120  class=""sort"">�� ��</td><td class=""sort"">����¼IP</td><td class=""sort"">����¼ʱ��</td><td  class=""sort"">���ǳ�ʱ��</td><td  class=""sort"">��¼����</td><td class=""sort"">�� ��</td><td class='sort'>�������</td></tr>")
		Set RSObj = Server.CreateObject("AdoDb.RecordSet")
		 RSObj.Open SqlStr, Conn, 1, 1
		 Dim T, TitleStr, LockStr, ShortName
		Do While Not RSObj.EOF
			
				If RSObj("Locked") = 1 Then
					LockStr = "<font color=red>������</font>"
					Else
					LockStr = "<font color=green>����</font>"
				End If
				TitleStr = " TITLE='�� ��:" & RSObj("RealName") & "&#13;&#10;�� ��:" & RSObj("Sex") & "&#13;&#10;���ʱ��:" & RSObj("AddDate") & "&#13;&#10;��Ҫ����:" & RSObj("Description") & "'"
			  Response.Write ("<tr><td class='splittd' height=25" & TitleStr & ">&nbsp;<span ondblclick=""EditAdmin(this.AdminID);"" AdminID=""" & RSObj("AdminID") & """><img src=Images/Folder/Admin" & Trim(CStr(RSObj("SuperTF"))) & ".gif border=0 align=absmiddle><span style=""cursor:default"">" & RSObj("UserName") & "</span><span></td>")
			  Response.Write ("<td  class='splittd' align=""center"">")
			  IF RSOBJ("SuperTF")=1 then response.write "<font color=red>��������Ա</font>" else response.write "��ͨ����Ա"
			  Response.Write ("</td>")
			  Response.Write ("<td class='splittd' align=""center"">" & RSObj("LastLoginIP") & "</td>")
			  Response.Write ("<td class='splittd' align=""center"">" & RSObj("LastLoginTime") & "</td>")
			  Response.Write ("<td class='splittd' align=""center"">" & RSObj("LastLogoutTime") & "</td>")
			  Response.Write ("<td class='splittd' align=""center"">" & RSObj("LoginTimes") & "</td>")
			  Response.Write ("<td class='splittd' align=""center"">" & LockStr & "</td>")
			  Response.Write ("<td class='splittd' align=""center""><a href='javascript:EditAdmin(" & rsobj("AdminID") &")'>�޸�</a> | <a")
			  if rsobj("supertf")="1" then response.write " disabled" else response.write " href='javascript:Delete("&rsobj("AdminID")&")'"
			  Response.Write ">ɾ��</a></td>"
			  Response.Write ("</tr>")
			  RSObj.MoveNext
			 If RSObj.EOF Then Exit Do
			Loop
			RSObj.Close:Conn.Close:Set RSObj = Nothing:Set GRS = Nothing
		  
		Response.Write "</table>"
		Response.Write "</div>"
		Response.Write "</body>"
		Response.Write "</html>"
		End Sub
		
		Sub AdminAdd()
		 IF KS.G("Method")="save" Then
		    Call AdminSave()
		   Else
		    Call AdminAddOrEdit()
		  End IF
		End SUB
		Sub AdminAddOrEdit()
		Dim SQL,RSC:Set RSC=Server.CreateObject("ADODB.RECORDSET")
		RSC.Open "Select ChannelID,ChannelName,BasicType,ItemName,ModelEname,ChannelStatus From KS_Channel where channelstatus=1 Order By ChannelID",Conn,1,1
		If Not RSC.Eof Then
		  SQL=RSC.GetRows(-1)
		End If
		RSC.Close:Set RSC=Nothing
		
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		Response.Write "<link href=""Include/admin_style.css"" rel=""stylesheet"">"
		Response.Write "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>"
		Response.Write "<script language=""JavaScript"" src=""../KS_Inc/jquery.js""></script>"
		Response.Write "<script language=""JavaScript"" src=""../ks_Inc/CheckPassWord.js""></script>"
		Response.Write "<style>"
		Response.Write ".rank { border:none; background:url(../images/rank.gif) no-repeat; width:125px; height:22px; vertical-align:middle; cursor:default; }"
		Response.Write ".r0 { background-position:0 0; }"
		Response.Write ".r1 { background-position:0 -22px; }"
		Response.Write ".r2 { background-position:0 -44px; }"
		Response.Write ".r3 { background-position:0 -66px; }"
		Response.Write ".r4 { background-position:0 -88px; }"
		Response.Write ".r5 { background-position:0 -110px; }"
		Response.Write ".r6 { background-position:0 -132px; }"
		Response.Write ".r7 { background-position:0 -154px; }"
		Response.Write "</style>"
		Response.Write "<title>����Ա���</title>"
		Response.Write "</head>"
		
		Dim AdminID, PrUserName, PassWord, Locked, RealName, Sex, TelPhone, Email, Descript, Action, GroupID, SuperTF
		
		Action = KS.G("Action")
		AdminID = KS.G("AdminID")
		GroupID = KS.G("GroupID")
		If Action = "" Then Action = "AddAdmin"
		If AdminID <> "" Then
		   Dim RSObj:Set RSObj = Conn.Execute("Select * From KS_Admin Where AdminID=" & AdminID)
		  If Not RSObj.EOF Then
			 UserName = Trim(RSObj("UserName"))
			 PrUserName=Trim(RSObj("PrUserName"))
			 Locked = Trim(CStr(RSObj("Locked")))
			 RealName = Trim(RSObj("RealName"))
			 Sex = Trim(RSObj("Sex"))
			 TelPhone = Trim(RSObj("TelPhone"))
			 Email = Trim(RSObj("Email"))
			 Descript = Trim(RSObj("Description"))
			 SuperTF = Trim(CStr(RSObj("SuperTF")))
			 PowerList=rsobj("powerlist")
	         ModelPower=rsobj("modelpower")
		  End If
		   RSObj.Close:Set RSObj = Nothing
		Else
		 ModelPower="sysset0,user0,lab0,model0,subsys0,other0,ask0,space0"
		 For i=0 to ubound(sql,2)
		  ModelPower=Modelpower &"," & sql(4,i)&"0"
		 Next
		End If
		
		
		
		Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
		
		 If AdminID = "" Then
			Response.Write ("<div class='topdashed sort'>��ӹ���Ա</div>")
		  Else
			Response.Write ("<div class='topdashed sort'>�޸Ĺ���Ա</div>")
		  End If

		Response.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"">"
		Response.Write "  <form action=""?Method=save"" name=""AdminForm"" method=""post"">"
		Response.Write "   <input name=""Action"" type=""hidden"" id=""Action"" value=""" & Action & """>"
		Response.Write "   <input name=""AdminID"" type=""hidden"" value=""" & AdminID & """>"
		Response.write "  <tr class='sort'><td colspan='4' align='center'>====��¼ѡ��====</td></tr>"
		Response.Write "    <tr class='tdbg'>"
		Response.Write "      <td class='clefttitle' align=""right"">����Ա����</td>"
		Response.Write "      <td height=""25"" colspan='3'>"
					
					If Action = "Edit" Then
						 Response.Write ("<input name=""UserName"" Readonly value=""" & UserName & """ type=""text"" id=""UserName"" size=""30"" class=""textbox"">")
					Else
						 Response.Write ("<input name=""UserName""  type=""text"" id=""UserName"" size=""30"" class=""textbox"">")
					End If
					 
		Response.Write "              ���ڵ�¼��̨������</td>"
		Response.Write "    </tr>"
		Response.Write "    <tr class='tdbg'>"
		Response.Write "            <td height=""25"" class='clefttitle' align=""right"">ǰ̨�û�����</td>"
		Response.Write "            <td colspan='3'>"
					If Action = "Edit" Then
						 Response.Write ("<input name=""PrUserName"" readonly value=""" & PrUserName & """ type=""text"" id=""PrUserName"" size=""30"" class=""textbox"">")
					Else
						 Response.Write ("<input name=""PrUserName""  type=""text"" id=""PrUserName"" size=""30"" class=""textbox"">")
					End if
					 
		Response.Write "              ǰ̨��Ա����ע����û���(���ɸ���),<a href='KS.User.asp?Action=Add'><font color=red>������</font></a></td>"
		Response.Write "          </tr>"		
				  
				  If Action <> "Edit" Then
		Response.Write "          <tr class='tdbg'>"
		Response.Write "            <td height=""25"" class=""clefttitle"" align=""right"">��ʼ���룺</td>"
		Response.Write "            <td height=""25"" colspan='3'>"
		Response.Write "             <table border='0'><tr><td><input name=""PassWord"" type=""password"" size=""20"" class=""textbox"" onkeyup=""javascript:setPasswordLevel(this, document.getElementById('passwordLevel'));"" onmouseout=""javascript:setPasswordLevel(this, document.getElementById('passwordLevel'));"" onblur=""javascript:setPasswordLevel(this, document.getElementById('passwordLevel'));""> ����ǿ�ȣ�</td><td> "
		Response.Write "         <input name=""Input"" disabled=""disabled"" class=""rank r0"" id=""passwordLevel"" /></td>"

		Response.Write "          </tr></table></td></tr>"
				 
				 End If
		Response.Write "          <tr class='tdbg'>"
		Response.Write "            <td height=""25"" align=""right"" class='clefttitle'>�Ƿ�������</td>"
		Response.Write "            <td height=""25"" colspan='3'>"
					
					If SuperTF = "1" Then
					   Response.Write ("<input type=""hidden"" value=""0"" name=""locked""> (��)����")
					  ElseIf Locked = "1" Then
					 Response.Write ("<input type=""radio"" name=""Locked"" value=""0""> ���� ")
					 Response.Write ("<input type=""radio"" name=""Locked"" value=""1"" checked> ���� ")
					 Else
					  Response.Write ("<input type=""radio"" name=""Locked"" value=""0"" checked> ���� ")
					  Response.Write ("<input type=""radio"" name=""Locked"" value=""1""> ���� ")
					 End If
					  
		Response.Write "              ����<font color=""#FF0000"">�������û����ܵ�¼��̨����</font></td>"
		Response.Write "          </tr>"
		
		Response.Write "          <tr class='tdbg'>"
		Response.Write "            <td height=""25"" align=""right"" class='clefttitle'>��ʵ������</td>"
		Response.Write "            <td><input name=""RealName"" type=""text"" class=""textbox"" value=""" & RealName & """ id=""RealName"" size=""30""></td>"
		Response.Write "            <td align=""right"" class='clefttitle'>�� ��</td>"
		Response.Write "            <td>"
					 
					 If Trim(Sex) = "Ů" Then
						  Response.Write ("<input type=""radio"" name=""Sex"" value=""��""> �� ")
						  Response.Write ("<input type=""radio"" name=""Sex"" value=""Ů"" checked>  Ů ")
					  Else
						  Response.Write ("<input type=""radio"" name=""Sex"" value=""��"" checked> �� ")
						  Response.Write ("<input type=""radio"" name=""Sex"" value=""Ů"">  Ů ")
					  End If
				   
		Response.Write "             </td>"
		Response.Write "          </tr>"
		Response.Write "          <tr class='tdbg'>"
		Response.Write "            <td height=""25"" align=""right"" class='clefttitle'>��ϵ�绰��</td>"
		Response.Write "            <td><input name=""TelPhone"" type=""text"" class=""textbox"" value=""" & TelPhone & """ id=""TelPhone"" size=""30""></td>"
		Response.Write "            <td align=""right"" class='clefttitle'>�������䣺</td>"
		Response.Write "            <td><input name=""Email"" type=""text"" class=""textbox"" id=""Email"" value=""" & Email & """ size=""30""></td>"
		Response.Write "          </tr>"
		Response.Write "          <tr class='tdbg'>"
		Response.Write "            <td height=""25"" align=""right"" class='clefttitle'>��Ҫ˵����</td>"
		Response.Write "            <td height=""25"" colspan='3'>"
		Response.Write "              <textarea class='textbox' name=""Description"" rows=""6"" id=""Description"" style=""width:80%;height:60px;border-style: solid; border-width: 1"">" & Descript & "</textarea></td>"
		Response.Write "          </tr>"
		if SuperTF=1 then
		Response.write "          <tr class='sort'><td colspan='4' align='center'>====�˹����ǳ�������Ա��ӵ�����Ȩ��====</td></tr>"
		Response.Write "          <tr class='tdbg' style='display:none'><td colspan='4'>"
		else
		Response.write "          <tr class='sort'><td colspan='4' align='center'>====�˹������ϸȨ������====</td></tr>"
		Response.Write "          <tr class='tdbg'><td colspan='4'>"
		end if
		
		 dim i
	 
	 %>
	 <table width="99%" border="0" align="center" cellspacing="0" cellpadding="0">  
	 <tr>
	 <td height="25" class='clefttitle'><strong> һ���˹���Ա�ڡ�<font color="#FF0000">���ݹ���</font>����Ȩ��</strong></td>
	 </tr>
	 </table>
    <table width="96%" border="0" align="center" cellspacing="0" cellpadding="0">   
	 <tr>       
	 <td> 
		  
          <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
		    <tr> 
              <td height="22" align="center"> <strong><font color="#993300">ģ�͹���Ȩ��</font></strong></td>
			  <td>
			  	
				<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td width="25%"> <input name="PowerList" type="checkbox" id="PowerList" value="M010001"<%if InStr(1, PowerList,"M010001" ,1)<>0 then Response.Write(" checked") %>>
                      ��Ŀ����</td>
                    <td width="25%"><input name="PowerList" type="checkbox" id="PowerList" value="M010002"<%if InStr(1, PowerList,"M010002" ,1)<>0 then Response.Write( " checked") %>>
                      ���۹���</td>
                    <td width="25%"><input name="PowerList" type="checkbox" id="PowerList" value="M010003"<%if InStr(1, PowerList,"M010003" ,1)<>0 then Response.Write( " checked") %>>
                      ר�����</td>
                    <td width="25%"><input name="PowerList" type="checkbox" id="PowerList" value="M010004"<%if InStr(1, PowerList,"M010004" ,1)<>0 then Response.Write( " checked") %>>
                      �ؼ���tags����</td>
				  </tr>
				  <tr>
                    <td width="25%"><input name="PowerList" type="checkbox" id="PowerList" value="M010005"<%if InStr(1, PowerList,"M010005" ,1)<>0 then Response.Write(" checked") %>>
��������</td>
                    <td width="25%"><input name="PowerList" type="checkbox" id="PowerList" value="M010006"<%if InStr(1, PowerList,"M010006" ,1)<>0 then Response.Write(" checked") %>>
                      ����վ����</td>
                    <td width="25%"> <input name="PowerList" type="checkbox" id="PowerList" value="M010007"<%if InStr(1, PowerList,"M010007" ,1)<>0 then Response.Write(" checked") %>>
                      һ��������</td>
                    <td width="25%"> <input name="PowerList" type="checkbox" id="PowerList" value="M010008"<%if InStr(1, PowerList,"M010008" ,1)<>0 then Response.Write(" checked") %>>
                      �ɼ�����</td>
                 
                  </tr>
				  </table>
				
				
				
			  </td>
			</tr>
			<tr><td colspan=2><hr size=1></td></tr>
		    <tr> 
              <td height="22" align="center"> <strong><font color="#993300">��ģ��Ȩ������</font></strong></td>
			  <td>
			      <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                  
			  	 <%
				  dim m:m=1
				 for i=0 to ubound(sql,2)%>
				   <tr> 
				   <td width="20%" height="35"> <strong><%=sql(1,i)%></strong></td>
				   </tr>
				   <tr>
				    <td>
					
					<%IF instr(ModelPower,sql(4,i) & "0")>0 Then
			      Response.Write("<input name=""ModelPower" & sql(0,i) & """ type=""radio"" onclick=""M_" & sql(1,i) & ".style.display='none';"" value=""" & sql(4,I) & "0"" checked>")
			   Else
			      Response.Write("<input name=""ModelPower" & sql(0,i) & """ type=""radio"" onclick=""M_" & sql(1,i) & ".style.display='none';"" value=""" & sql(4,I) & "0"">")
			   End IF
			   %>
                ��<%=SQL(1,I)%>���κι���Ȩ��(����)
					<br/>
					<%IF instr(ModelPower,sql(4,i) & "1")>0 Then
			      Response.Write("<input type=""radio"" onclick=""M_" & sql(1,i) & ".style.display='none';"" name=""ModelPower" & sql(0,i) & """ value=""" & sql(4,I) & "1"" Checked>")
				Else
			      Response.Write("<input type=""radio"" onclick=""M_" & sql(1,i) & ".style.display='none';"" name=""ModelPower" & sql(0,i) & """ value=""" & sql(4,I) & "1"">")
				End IF
				%>
                ģ�͹���Ա��ӵ�и�ģ�͵����й���Ȩ��(�൱�ڶ�<%=sql(1,i)%>û���κ�����)
				 <br>
				 <%IF instr(ModelPower,sql(4,i) & "2")>0 Then
			     Response.Write("<input type=""radio"" onclick=""M_" & sql(1,i) & ".style.display='';"" name=""ModelPower" & sql(0,i) & """ value=""" & sql(4,I) & "2"" Checked>")
			   Else
			     Response.Write("<input type=""radio"" onclick=""M_" & sql(1,i) & ".style.display='';"" name=""ModelPower" & sql(0,i) & """ value=""" & sql(4,I) & "2"">")
			   End IF
			   %>
                ��Ŀ����Ա��ֻӵ�в�����Ŀ(Ƶ��)����Ȩ��
					
					</td>
				   </tr>
				   <tr ID="M_<%=sql(1,i)%>" <%IF instr(ModelPower,sql(4,i) & "2")=0 Then Response.Write("style=""display:none""") End IF%>>	 
			       <td height="22">
				    
					
					
					<%
	  Select Case SQL(2,I)
	   Case 1,2,3,4,7,8
	   %>  
              <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td height="25" colspan="7"><strong><font color="#993300">Ȩ������</font></strong></td>
                  </tr>
                    <%
					Call BasePurview(PowerList,SQL,I)
					%>
                  <tr> 
                    <td height="25" colspan="7"><font color="#993300"><strong>��ϸָ����Ŀ��Ƶ����Ȩ��</strong></font></td>
                  </tr>
                  <tr> 
                    <td colspan="7"> 
					   <%
                       Call ClassList(SQL(0,I))
					   %>
					</td>
                  </tr>
                </table>
	<%case 5%>
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">

            <tr>
              <td height="25" colspan="7"><strong><font color="#993300">Ȩ������</font></strong></td>
            </tr>
            <%Call BasePurview(PowerList,SQL,I)%>
            <tr>
              <td height="25" colspan="7"><strong><font color="#993300">�������</font></strong></td>
            </tr>
            <tr>
			  <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10012"<%if InStr(1, PowerList,"M"&SQL(0,I) & "10012" ,1)<>0 then Response.Write(" checked") %>>
��������</td>
              <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10014"<%if InStr(1, PowerList,"M"&SQL(0,I) & "10014" ,1)<>0 then Response.Write(" checked") %>>
                �ʽ���ϸ</td>
              <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10015"<%if InStr(1, PowerList,"M"&SQL(0,I) & "10015" ,1)<>0 then Response.Write(" checked") %>>
���˻���ѯ</td>
              <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10016"<%if InStr(1, PowerList,"M"&SQL(0,I) & "10016" ,1)<>0 then Response.Write(" checked") %>>
����Ʊ��ѯ</td>
              <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10017"<%if InStr(1, PowerList,"M"&SQL(0,I) & "10017" ,1)<>0 then Response.Write(" checked") %>>
����ͳ��</td>
              <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10018"<%if InStr(1, PowerList,"M"&SQL(0,I) & "10018" ,1)<>0 then Response.Write(" checked") %>>
Ʒ�ƹ���</td>
            </tr>
			
            <tr>
              <td nowrap="nowrap" title="������༭��ɾ�����̵Ȳ�����Ȩ��"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20003"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20003" ,1)<>0 then Response.Write( " checked") %> />
                ���̹���</td>
				<td nowrap="nowrap" title="������༭��ɾ���ͻ���ʽ�Ȳ�����Ȩ��"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20004"<%if InStr(1, PowerList,"M" & sql(0,i) & "20004" ,1)<>0 then Response.Write( " checked") %> />
                �ͻ���ʽ����</td>
				<td nowrap="nowrap" title="������༭��ɾ�����ʽ�Ȳ�����Ȩ��"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20005"<%if InStr(1, PowerList,"M" & sql(0,i) & "20005" ,1)<>0 then Response.Write( " checked") %> />
                ���ʽ����</td>               
				<td nowrap><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20001"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20001" ,1)<>0 then Response.Write( " checked") %>>
                      �����ص����</td>
                    
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20007"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20007" ,1)<>0 then Response.Write( " checked") %>> �Ż�ȯ����</td>
                    <td nowrap="nowrap"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20008"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20008" ,1)<>0 then Response.Write( " checked") %>> ��ʱ/��������</td>
            </tr>
			<tr>
               <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20009"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20009" ,1)<>0 then Response.Write( " checked") %>> �������۹���</td>
               <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20010"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20010" ,1)<>0 then Response.Write( " checked") %>> ������Ʒ����</td>
               <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20011"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20011" ,1)<>0 then Response.Write( " checked") %>> ��ֵ�������</td>
              <td nowrap="nowrap"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>30001"<%if InStr(1, PowerList,"M" & SQL(0,I) & "30001" ,1)<>0 then Response.Write( " checked") %> />
                �Ź�����</td>
				<td nowrap="nowrap" title="������༭��ɾ���ͻ���ʽ�Ȳ�����Ȩ��"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>30002"<%if InStr(1, PowerList,"M" & sql(0,i) & "30002" ,1)<>0 then Response.Write( " checked") %> />
                ��Ȥ�����</td>
			</tr>
			
			
            <tr>
              <td height="25" colspan="7"><font color="#993300"><strong>��ϸָ����Ŀ��Ƶ����Ȩ��</strong></font></td>
            </tr>
            <tr>
              <td colspan="7">                      
			        <%
                       Call ClassList(SQL(0,I))
					   %>
               </td>
            </tr>
          </table>
  <%case 6%>
  <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="25" colspan="7"><strong><font color="#993300">Ȩ������</font></strong></td>
            </tr>
            <tr>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="KSM1001"<%if InStr(1, PowerList,"KSM1001" ,1)<>0 then Response.Write(" checked") %> />
                ��������</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="KSM1002"<%if InStr(1, PowerList,"KSM1002" ,1)<>0 then Response.Write(" checked") %> />
                ר������</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="KSM1003"<%if InStr(1, PowerList,"KSM1003" ,1)<>0 then Response.Write(" checked") %> />
                ������</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="KSM1004"<%if InStr(1, PowerList,"KSM1004" ,1)<>0 then Response.Write( " checked") %> />
                ���ֹ���</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="KSM1005"<%if InStr(1, PowerList,"KSM1005" ,1)<>0 then Response.Write( " checked") %> />
                ���۹���</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="KSM1006"<%if InStr(1, PowerList,"KSM1006" ,1)<>0 then Response.Write( " checked") %> />
               ����������</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="KSM1007"<%if InStr(1, PowerList,"KSM1007" ,1)<>0 then Response.Write(" checked") %>>
               ����û����</td>
            </tr>
			<tr>
			 <td colspan=6><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20005"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20005" ,1)<>0 then Response.Write( " checked") %>>
                      ����HTML����</td>
			</tr>
            
          </table>

		   <%case 9%>
           <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="25" colspan="7"><strong><font color="#993300">��Ŀ����Ϣ</font></strong></td>
            </tr>
            <tr>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M910001"<%if InStr(1, PowerList,"M910001" ,1)<>0 then Response.Write(" checked") %> />
                ��Ŀ����</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M910002"<%if InStr(1, PowerList,"M910002" ,1)<>0 then Response.Write( " checked") %> />
����Ծ�</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M910003"<%if InStr(1, PowerList,"M910003" ,1)<>0 then Response.Write( " checked") %> />
�༭�Ծ�</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M910004"<%if InStr(1, PowerList,"M910004" ,1)<>0 then Response.Write( " checked") %> />
ɾ���Ծ�</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M910005"<%if InStr(1, PowerList,"M910005" ,1)<>0 then Response.Write(" checked") %>>
�ƶ��Ծ�</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M910007"<%if InStr(1, PowerList,"M910007" ,1)<>0 then Response.Write(" checked") %>>
��������</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M910009"<%if InStr(1, PowerList,"M910009" ,1)<>0 then Response.Write(" checked") %>>
�ϴ��ļ�</td>
            </tr>
			</table>
		   <%case 10%>
           <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="25" colspan="7"><strong><font color="#993300">��Ŀ����Ϣ</font></strong></td>
            </tr>
            <tr>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M1010001"<%if InStr(1, PowerList,"M1010001" ,1)<>0 then Response.Write(" checked") %> />
                ��Ƹϵͳ����</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M1010002"<%if InStr(1, PowerList,"M1010002" ,1)<>0 then Response.Write( " checked") %> />
��Ƹ��λ����</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M1010003"<%if InStr(1, PowerList,"M1010003" ,1)<>0 then Response.Write( " checked") %> />
���˼�������</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M1010004"<%if InStr(1, PowerList,"M1010004" ,1)<>0 then Response.Write( " checked") %> />
��Ƹְλ����</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M1010005"<%if InStr(1, PowerList,"M1010005" ,1)<>0 then Response.Write(" checked") %>>
��ҵְλ����</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M1010007"<%if InStr(1, PowerList,"M1010007" ,1)<>0 then Response.Write(" checked") %>>
����ģ�����</td>
            </tr>
			</table>
			<%
	End Select
	%>  
				   </td>
				   </tr>
				   
                  <%
				  Next%>
				  
				   <tr> 
				   <td width="20%" height="35"> <strong>�ʴ�ϵͳȨ��</strong></td>
				   </tr>
				   <tr>
				    <td>
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
			<tr> 
              <td height="25" colspan="5"> 
                <%
					IF instr(ModelPower,"ask0")>0 Then
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('Ask','none')"" name=""ask"" value=""ask0"" checked>")
                    ELSE
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('Ask','none')"" name=""ask"" value=""ask0"">")
					END IF
					%>
                ���ʴ�ϵͳ���κι���Ȩ��(����)
				  <br/>
                <%
					IF instr(ModelPower,"ask1")>0 Then
					  Response.Write("<input type=""radio"" name=""ask"" onclick=""SetPowerListValue('Ask','')"" value=""ask1"" checked>")
                     ELSE
					  Response.Write("<input type=""radio"" name=""ask"" onclick=""SetPowerListValue('Ask','')"" value=""ask1"">")
					 END IF%>
                ӵ��ָ���Ĳ��ֹ���Ȩ�� 
			 </td>
            </tr>
            <tr ID="Ask" <% IF instr(ModelPower,"ask1")="0" then Response.Write("style=""display:none""") End IF%>> 
							<td><input name="PowerList" type="checkbox" id="PowerList" value="WDXT10000"<%if InStr(1, PowerList,"WDXT10000" ,1)<>0 then Response.Write( " checked") %>> 
							�ʴ��������
		</td>
							<td><input name="PowerList" type="checkbox" id="PowerList" value="WDXT10001"<%if InStr(1, PowerList,"WDXT10001" ,1)<>0 then Response.Write( " checked") %>> 
							�༭ɾ������</td>
							<td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="WDXT10002"<%if InStr(1, PowerList,"WDXT10002" ,1)<>0 then Response.Write( " checked") %>> 
							�ʴ�������</td>
							<td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="WDXT10003"<%if InStr(1, PowerList,"WDXT10003" ,1)<>0 then Response.Write( " checked") %>>
		�ȼ�ͷ�ι���</td>
		

					</tr>
					
				   <tr> 
				   <td width="20%" height="35"> <strong>�ռ��Ż�Ȩ��</strong></td>
				   </tr>
			<tr> 
              <td height="25" colspan="5"> 
                <%
					IF instr(ModelPower,"space0")>0 Then
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('Space','none')"" name=""space"" value=""space0"" checked>")
                    ELSE
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('Space','none')"" name=""space"" value=""space0"">")
					END IF
					%>
                �ڿռ��Ż����κι���Ȩ��(����)
				  <br/>
                <%
					IF instr(ModelPower,"space1")>0 Then
					  Response.Write("<input type=""radio"" name=""space"" onclick=""SetPowerListValue('Space','')"" value=""space1"" checked>")
                     ELSE
					  Response.Write("<input type=""radio"" name=""space"" onclick=""SetPowerListValue('Space','')"" value=""space1"">")
					 END IF%>
                ӵ��ָ���Ĳ��ֹ���Ȩ�� 
			 </td>
            </tr>
            <tbody ID="Space" <% IF instr(ModelPower,"space1")="0" then Response.Write("style=""display:none""") End IF%>> 
                   <tr>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10000"<%if InStr(1, PowerList,"KSMS10000" ,1)<>0 then Response.Write( " checked") %>>
�ռ��������</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10001"<%if InStr(1, PowerList,"KSMS10001" ,1)<>0 then Response.Write( " checked") %>>
���˿ռ����</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10002"<%if InStr(1, PowerList,"KSMS10002" ,1)<>0 then Response.Write( " checked") %>>
�ռ���־����</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10003"<%if InStr(1, PowerList,"KSMS10003" ,1)<>0 then Response.Write( " checked") %>>
�û������� </td>
                  </tr>
				  <tr>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10004"<%if InStr(1, PowerList,"KSMS10004" ,1)<>0 then Response.Write( " checked") %>>
�û�Ȧ�ӹ���</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10005"<%if InStr(1, PowerList,"KSMS10005" ,1)<>0 then Response.Write( " checked") %>>
�û����Թ���</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20013"<%if InStr(1, PowerList,"KSMS20013" ,1)<>0 then Response.Write( " checked") %>>
��ҵ������</td>
				   <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10007"<%if InStr(1, PowerList,"KSMS10007" ,1)<>0 then Response.Write( " checked") %>>
�û���������</td>
                  </tr>
				  <tr>
				   <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10008"<%if InStr(1, PowerList,"KSMS10008" ,1)<>0 then Response.Write( " checked") %>>
��ҵ��Ϣ����</td>
				   <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10009"<%if InStr(1, PowerList,"KSMS10009" ,1)<>0 then Response.Write( " checked") %>>
��ҵ���Ź���</td>
				   <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10010"<%if InStr(1, PowerList,"KSMS10010" ,1)<>0 then Response.Write( " checked") %>>
��ҵ��Ʒ����</td>
				   <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20011"<%if InStr(1, PowerList,"KSMS20011" ,1)<>0 then Response.Write( " checked") %>>
����֤�����</td>
                  </tr>
				  <tr>
				   <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20012"<%if InStr(1, PowerList,"KSMS20012" ,1)<>0 then Response.Write( " checked") %>>
��ҵ�������</td>
				   <td></td>

				  </tr>					
				 </tbody>	
					
					
				  </table>
					
		            </td>
				  </tr>
				  
				 
				 </table>
				  
			  </td>
			</tr>
			
		
		 </table>
			  
            
			 
			 
	       </TD>
            </TR>
          </table>
         </td>
		  </tr>
    </table>
	
	  <br/>
	 <table width="99%" border="0" align="center" cellspacing="0" cellpadding="0">  
	 <tr>
	 <td height="25" class='clefttitle'><strong> �����˹���Ա�ڡ�<font color="#FF0000">ϵͳ����</font>����Ȩ��</strong></td>
	 </tr>
	 </table>

	  <table width="96%" align="center" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>
          <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr> 
              <td height="25" colspan="2"> 
                <%
					IF instr(ModelPower,"sysset0")>0 Then
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('System','none')"" name=""sysset"" value=""sysset0"" checked>")
                    ELSE
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('System','none')"" name=""sysset"" value=""sysset0"">")
					END IF
					%>
                ��ϵͳ�����������κι���Ȩ��(����)
				  <br/>
                <%
					IF instr(ModelPower,"sysset1")>0 Then
					  Response.Write("<input type=""radio"" name=""sysset"" onclick=""SetPowerListValue('System','')"" value=""sysset1"" checked>")
                     ELSE
					  Response.Write("<input type=""radio"" name=""sysset"" onclick=""SetPowerListValue('System','')"" value=""sysset1"">")
					 END IF%>
                ӵ��ָ���Ĳ��ֹ���Ȩ�� 
			 </td>
            </tr>
            <tr ID="System" <% IF instr(ModelPower,"sysset1")="0" then Response.Write("style=""display:none""") End IF%>> 
              <td height="25" colspan="2">
			  <table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td width="13%" height="25" align="center" nowrap><strong>ϵͳ����</strong></td>
                    <td width="13%"> <input name="PowerList" type="checkbox" id="PowerList" value="KMST10001"<%if InStr(1, PowerList,"KMST10001" ,1)<>0 then Response.Write( " checked") %>>
                      ������Ϣ����</td>
                    <td width="11%"><input name="PowerList" type="checkbox" id="PowerList" value="KMST10002"<%if InStr(1, PowerList,"KMST10002" ,1)<>0 then Response.Write( " checked") %>>
API��������</td>
                    <td width="15%"><input name="PowerList" type="checkbox" id="PowerList" value="KMST20000"<%if InStr(1, PowerList,"KMST20000" ,1)<>0 then Response.Write( " checked") %>>
���»���</td>
                    <td width="15%" nowrap><input name="PowerList" type="checkbox" id="PowerList" value="KMST20001"<%if InStr(1, PowerList,"KMST20001" ,1)<>0 then Response.Write( " checked") %>>
���ز�������</td>
                    <td width="15%" nowrap><input name="PowerList" type="checkbox" id="PowerList" value="KMST20002"<%if InStr(1, PowerList,"KMST20002" ,1)<>0 then Response.Write( " checked") %>>
���ط���������</td>
                  </tr>
                  <tr> 
                    <td height="25" align="center" nowrap></td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST20003"<%if InStr(1, PowerList,"KMST20003" ,1)<>0 then Response.Write( " checked") %>>
Ӱ�Ӳ�������</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST20004"<%if InStr(1, PowerList,"KMST20004" ,1)<>0 then Response.Write( " checked") %>>
Ӱ�ӷ�����</td>
                    <td nowrap><input name="PowerList" type="checkbox" id="PowerList" value="KMST20005"<%if InStr(1, PowerList,"KMST20005" ,1)<>0 then Response.Write( " checked") %>>
���������</td>
                  </tr>
                  
				  <tr> 
                    <td width="13%" height="25" align="center" nowrap><strong>��������</strong></td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10004"<%if InStr(1, PowerList,"KMST10004" ,1)<>0 then Response.Write( " checked") %>>
                      ���ݹؼ���</td>                    
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10006"<%if InStr(1, PowerList,"KMST10006" ,1)<>0 then Response.Write( " checked") %>>
                      ��̨��־����</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10007"<%if InStr(1, PowerList,"KMST10007" ,1)<>0 then Response.Write( " checked") %>>
                      ����/ѹ�����ݿ�</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10008"<%if InStr(1, PowerList,"KMST10008" ,1)<>0 then Response.Write( " checked") %>>
                      ���ݿ��ֶ��滻</td>
					 <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10019"<%if InStr(1, PowerList,"KMST10019" ,1)<>0 then Response.Write( " checked") %>>
                      �����ؼ���ά��</td>
                  </tr>
				  <tr>
				    <td height="25" align="center" nowrap>&nbsp;</td>
				    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10013"<%if InStr(1, PowerList,"KMST10013" ,1)<>0 then Response.Write(" checked") %> />
ר�����</td>
				    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10015"<%if InStr(1, PowerList,"KMST10015" ,1)<>0 then Response.Write(" checked") %> />
��Դ����</td>
				    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10016"<%if InStr(1, PowerList,"KMST10016" ,1)<>0 then Response.Write(" checked") %> />
���߹���</td>
				    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10017"<%if InStr(1, PowerList,"KMST10017" ,1)<>0 then Response.Write(" checked") %> />
ʡ�й���</td>
				    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10014"<%if InStr(1, PowerList,"KMST10014" ,1)<>0 then Response.Write(" checked") %> />
ͼƬͶƱ��¼</td>			    </tr>
				  <tr>
				    <td height="25" align="center" nowrap></td>
				    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10009"<%if InStr(1, PowerList,"KMST10009" ,1)<>0 then Response.Write( " checked") %>>
����ִ��SQL��� </td>
				    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10010"<%if InStr(1, PowerList,"KMST10010" ,1)<>0 then Response.Write( " checked") %>>
�ռ�ռ���� </td>

				    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10011"<%if InStr(1, PowerList,"KMST10011" ,1)<>0 then Response.Write( " checked") %>>
����������̽��</td>
			        <td nowrap><input name="PowerList" type="checkbox" id="PowerList" value="KMST10012"<%if InStr(1, PowerList,"KMST10012" ,1)<>0 then Response.Write( " checked") %>>
���߼��ľ��</td>
			        <td nowrap><input name="PowerList" type="checkbox" id="PowerList" value="KMST10018"<%if InStr(1, PowerList,"KMST10018" ,1)<>0 then Response.Write( " checked") %>>
�ϴ��ļ�����</td>
		          </tr>
				  <tr>
				    <td height="25" align="center" nowrap></td>
				    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10020"<%if InStr(1, PowerList,"KMST10020" ,1)<>0 then Response.Write( " checked") %>>
��ʱ������� </td>
				    <td> </td>

				    <td></td>
			        <td nowrap></td>
			        <td nowrap></td>
		          </tr>
                </table>
			  </td>
            </tr>
			
          </table>
	  </td></tr>
	  </table>
	  
	  
	  <br/>
	 <table width="99%" border="0" align="center" cellspacing="0" cellpadding="0">  
	 <tr>
	 <td height="25" class='clefttitle'><strong> �����˹���Ա�ڡ�<font color="#FF0000">�û�����</font>����Ȩ��</strong></td>
	 </tr>
	 </table>

	  <table width="96%" align="center" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>
          <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr> 
              <td height="25" colspan="2"> 
                <%
					IF instr(ModelPower,"user0")>0 Then
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('UserAdmin','none')"" name=""userPower"" value=""user0"" checked>")
                    ELSE
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('UserAdmin','none')"" name=""UserPower"" value=""user0"">")
					END IF
					%>
                ���û������������κι���Ȩ��(����)
				<br/>
                <%
					IF instr(ModelPower,"user1")>0 Then
					  Response.Write("<input type=""radio"" name=""UserPower"" onclick=""SetPowerListValue('UserAdmin','')"" value=""user1"" checked>")
                     ELSE
					  Response.Write("<input type=""radio"" name=""UserPower"" onclick=""SetPowerListValue('UserAdmin','')"" value=""user1"">")
					 END IF%>
                ӵ��ָ���Ĳ��ֹ���Ȩ�� </td>
            </tr>
            <tr ID="UserAdmin" <% IF instr(ModelPower,"user1")=0 then Response.Write("style=""display:none""") End IF%>> 
              <td height="25"> <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td><input disabled="disabled" name="PowerList" type="checkbox" id="PowerList" value="KMUA10001"<%if InStr(1, PowerList,"KMUA10001" ,1)<>0 then Response.Write( " checked") %>>
                      ����Ա����</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10002"<%if InStr(1, PowerList,"KMUA10002" ,1)<>0 then Response.Write( " checked") %>>
                      ע���û�����</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10003"<%if InStr(1, PowerList,"KMUA10003" ,1)<>0 then Response.Write( " checked") %>>
                     �û����Ź��� </td>
                    <td> <input name="PowerList" type="checkbox" id="PowerList" value="KMUA10004"<%if InStr(1, PowerList,"KMUA10004" ,1)<>0 then Response.Write( " checked") %>>
                      �û������</td>
				    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10008"<%if InStr(1, PowerList,"KMUA10008" ,1)<>0 then Response.Write( " checked") %>>
                    ��ֵ������</td>                 
				    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10011"<%if InStr(1, PowerList,"KMUA10011" ,1)<>0 then Response.Write( " checked") %>>
                    �鿴��������</td>					 </tr>
                  <tr>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10005"<%if InStr(1, PowerList,"KMUA10005" ,1)<>0 then Response.Write( " checked") %>>
                    ��Ա��ȯ��ϸ</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10006"<%if InStr(1, PowerList,"KMUA10006" ,1)<>0 then Response.Write( " checked") %>>
                    ��Ա��Ч����ϸ</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10007"<%if InStr(1, PowerList,"KMUA10007" ,1)<>0 then Response.Write( " checked") %>>
                    ��Ա�ʽ���ϸ</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10008"<%if InStr(1, PowerList,"KMUA10008" ,1)<>0 then Response.Write( " checked") %>>
                    ��Ա�ʽ���ϸ</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10009"<%if InStr(1, PowerList,"KMUA10009" ,1)<>0 then Response.Write( " checked") %>>
                    �����ʼ�����</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10010"<%if InStr(1, PowerList,"KMUA10010" ,1)<>0 then Response.Write( " checked") %>>
                    �޸��Լ�������</td>
                  </tr>
                  <tr>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10012"<%if InStr(1, PowerList,"KMUA10012" ,1)<>0 then Response.Write( " checked") %>>
                    ��Ա�ֶι���</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10013"<%if InStr(1, PowerList,"KMUA10013" ,1)<>0 then Response.Write( " checked") %>>
                    ��Ա������</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10014"<%if InStr(1, PowerList,"KMUA10014" ,1)<>0 then Response.Write( " checked") %>>
                    ��Ա��̬����</td>
                    <td>&nbsp;</td>
                    <td height="25">&nbsp;</td>
                    <td height="25">&nbsp;</td>
                  </tr>
			
		
                </table></td>
            </tr>
          </table>
	  </td></tr>
	  </table>
	  
	  
	  <br/>
	 <table width="99%" border="0" align="center" cellspacing="0" cellpadding="0">  
	 <tr>
	 <td height="25" class='clefttitle'><strong> �ġ��˹���Ա�ڡ�<font color="#FF0000">��ǩģ�����</font>����Ȩ��</strong></td>
	 </tr>
	 </table>

	  <table width="96%" align="center" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>
          <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr> 
              <td height="25" colspan="2"> 
                <%
					IF instr(ModelPower,"lab0")>0 Then
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('KMTemplatePower','none')"" name=""labPower"" value=""lab0"" checked>")
                    ELSE
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('KMTemplatePower','none')"" name=""labPower"" value=""lab0"">")
					END IF
					%>
                ��ģ���ǩ�������Ȩ��(����)</td>
            </tr>
            <tr> 
              <td height="25" colspan="2"> 
                <%
					IF instr(ModelPower,"lab1")>0 Then
					  Response.Write("<input type=""radio"" name=""labPower"" onclick=""SetPowerListValue('KMTemplatePower','')"" value=""lab1"" checked>")
                     ELSE
					  Response.Write("<input type=""radio"" name=""labPower"" onclick=""SetPowerListValue('KMTemplatePower','')"" value=""lab1"">")
					 END IF%>
                ӵ��ָ���Ĳ��ֹ���Ȩ�� </td>
            </tr>
            <tr ID="KMTemplatePower" <% IF instr(ModelPower,"lab1")=0 then Response.Write("style=""display:none""") End IF%>> 
              <td height="25"> <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td><strong>��������</strong></td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMTL20000"<%if InStr(1, PowerList,"KMTL20000" ,1)<>0 then Response.Write( " checked") %>>
����վ����ҳ</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMTL20001"<%if InStr(1, PowerList,"KMTL20001" ,1)<>0 then Response.Write( " checked") %>>
ר�ⷢ������</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMTL20002"<%if InStr(1, PowerList,"KMTL20002" ,1)<>0 then Response.Write( " checked") %>>
ϵͳJS�������� </td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMTL20003"<%if InStr(1, PowerList,"KMTL20003" ,1)<>0 then Response.Write( " checked") %>>
����ͨ���Զ���ҳ��</td>
                    <td></td>
                  </tr>

                  <tr>
                    <td width="10%" title="��Ȩ�ް���ϵͳ������ǩĿ¼������ǩ����������ǩ��"><strong>ģ���ǩ</strong></td> 
                    <td width="16%" title="��Ȩ�ް���ϵͳ������ǩĿ¼������ǩ����������ǩ��"> <input name="PowerList" type="checkbox" id="PowerList" value="KMTL10001"<%if InStr(1, PowerList,"KMTL10001" ,1)<>0 then Response.Write( " checked") %>>
                      ϵͳ������ǩ </td>
                    <td width="18%" height="25" title="��Ȩ�ް������ɱ�ǩĿ¼������ǩ����������ǩ��"> <input name="PowerList" type="checkbox" id="PowerList" value="KMTL10002"<%if InStr(1, PowerList,"KMTL10002" ,1)<>0 then Response.Write( " checked") %>>
                      �Զ���SQL��ǩ</td>
                    <td width="18%" height="25" title="��Ȩ�ް���ϵͳJSĿ¼������ǩ����������ǩ��"> <input name="PowerList" type="checkbox" id="PowerList" value="KMTL10003"<%if InStr(1, PowerList,"KMTL10003" ,1)<>0 then Response.Write( " checked") %>>
                    �Զ��徲̬��ǩ </td>
                    <td width="14%"  title="��Ȩ�ް�������JSĿ¼������ǩ����������ǩ��"><input name="PowerList" type="checkbox" id="PowerList" value="KMTL10004"<%if InStr(1, PowerList,"KMTL10004" ,1)<>0 then Response.Write( " checked") %>>
                    ϵͳJS����</td>
                    <td width="24%" title="��Ȩ�ް���ģ�嵼�룬ģ��༭��"><input name="PowerList" type="checkbox" id="PowerList" value="KMTL10005"<%if InStr(1, PowerList,"KMTL10005" ,1)<>0 then Response.Write( " checked") %>>
����JS����</td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMTL10006"<%if InStr(1, PowerList,"KMTL10006" ,1)<>0 then Response.Write( " checked") %>>�Զ���ҳ�����</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMTL10007"<%if InStr(1, PowerList,"KMTL10007" ,1)<>0 then Response.Write( " checked") %>>
ģ�����</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMSL10008"<%if InStr(1, PowerList,"KMSL10008" ,1)<>0 then Response.Write( " checked") %> />
���ɶ����˵�</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMSL10009"<%if InStr(1, PowerList,"KMSL10009" ,1)<>0 then Response.Write( " checked") %> />
�������Ͳ˵�</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMSL10010"<%if InStr(1, PowerList,"KMSL10010" ,1)<>0 then Response.Write( " checked") %> />
ͨ��ѭ����ǩ</td>
                  </tr>
                </table></td>
            </tr>
          </table>
	  </td></tr>
	  </table>

	  <br/>
	 <table width="99%" border="0" align="center" cellspacing="0" cellpadding="0">  
	 <tr>
	 <td height="25" class='clefttitle'><strong> �塢�˹���Ա�ڡ�<font color="#FF0000">ģ���ֶι���</font>����Ȩ��</strong></td>
	 </tr>
	 </table>
	  
	  <table width="96%" align="center" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>
          <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr> 
              <td height="25" colspan="2"> 
                <%
					IF instr(ModelPower,"model0")>0 Then
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('ModelPowers','none')"" name=""modelPower"" value=""model0"" checked>")
                    ELSE
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('ModelPowers','none')"" name=""modelPower"" value=""model0"">")
					END IF
					%>
                ��ģ���ֶι������Ȩ��(����)</td>
            </tr>
            <tr> 
              <td height="25" colspan="2"> 
                <%
					IF instr(ModelPower,"model1")>0 Then
					  Response.Write("<input type=""radio"" name=""modelPower"" onclick=""SetPowerListValue('ModelPowers','')"" value=""model1"" checked>")
                     ELSE
					  Response.Write("<input type=""radio"" name=""modelPower"" onclick=""SetPowerListValue('ModelPowers','')"" value=""model1"">")
					 END IF%>
                ӵ��ָ���Ĳ��ֹ���Ȩ�� </td>
            </tr>
            <tr ID="ModelPowers" <% IF instr(ModelPower,"model1")=0 then Response.Write("style=""display:none""") End IF%>> 
              <td height="25"> <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMM10000"<%if InStr(1, PowerList,"KSMM10000" ,1)<>0 then Response.Write( " checked") %>>
���ģ��</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMM10001"<%if InStr(1, PowerList,"KSMM10001" ,1)<>0 then Response.Write( " checked") %>>
�޸�ģ��</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KSMM10002"<%if InStr(1, PowerList,"KSMM10002" ,1)<>0 then Response.Write( " checked") %>>
ɾ��ģ��</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KSMM10003"<%if InStr(1, PowerList,"KSMM10003" ,1)<>0 then Response.Write( " checked") %>>
ģ���ֶι��� </td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMM10004"<%if InStr(1, PowerList,"KSMM10004" ,1)<>0 then Response.Write( " checked") %>>
��Ϣͳ�� </td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMM10005"<%if InStr(1, PowerList,"KSMM10005" ,1)<>0 then Response.Write( " checked") %>>
�����ر�</td>
                  </tr>
                
                </table></td>
            </tr>
          </table>
	  </td></tr>
	  </table>


	  <br/>
	 <table width="99%" border="0" align="center" cellspacing="0" cellpadding="0">  
	 <tr>
	 <td height="25" class='clefttitle'><strong> �����˹���Ա�ڡ�<font color="#FF0000">���ѡ��</font>����Ȩ��</strong></td>
	 </tr>
	 </table>
<table width="96%" align="center" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>
          <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr> 
              <td height="25" colspan="2"> 
                <%
					IF instr(ModelPower,"subsys0")>0 Then
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('SubSysPowers','none')"" name=""subsysPower"" value=""subsys0"" checked>")
                    ELSE
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('SubSysPowers','none')"" name=""subsysPower"" value=""subsys0"">")
					END IF
					%>
                ����ϵͳ����Ȩ��(����)</td>
            </tr>
            <tr> 
              <td height="25" colspan="2"> 
                <%
					IF instr(ModelPower,"subsys1")>0 Then
					  Response.Write("<input type=""radio"" name=""subsysPower"" onclick=""SetPowerListValue('SubSysPowers','')"" value=""subsys1"" checked>")
                     ELSE
					  Response.Write("<input type=""radio"" name=""subsysPower"" onclick=""SetPowerListValue('SubSysPowers','')"" value=""subsys1"">")
					 END IF%>
                ӵ��ָ���Ĳ��ֹ���Ȩ�� </td>
            </tr>
            <tr ID="SubSysPowers" <% IF instr(ModelPower,"subsys1")=0 then Response.Write("style=""display:none""") End IF%>> 
              <td height="25"> <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                  
				  
				  <tr>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20000"<%if InStr(1, PowerList,"KSMS20000" ,1)<>0 then Response.Write( " checked") %>>
Ͷ�߽������</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20001"<%if InStr(1, PowerList,"KSMS20001" ,1)<>0 then Response.Write( " checked") %>>
�������ӹ���</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20002"<%if InStr(1, PowerList,"KSMS20002" ,1)<>0 then Response.Write( " checked") %>>
��վ�������</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20003"<%if InStr(1, PowerList,"KSMS20003" ,1)<>0 then Response.Write( " checked") %>>
վ�ڵ������ </td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20004"<%if InStr(1, PowerList,"KSMS20004" ,1)<>0 then Response.Write( " checked") %>>
��վ���Թ���</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20005"<%if InStr(1, PowerList,"KSMS20005" ,1)<>0 then Response.Write( " checked") %>>
վ�������</td>                    
<td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20006"<%if InStr(1, PowerList,"KSMS20006" ,1)<>0 then Response.Write( " checked") %>>
���ϵͳ����</td>

                  </tr>
				 <tr>
				   <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20007"<%if InStr(1, PowerList,"KSMS20007" ,1)<>0 then Response.Write( " checked") %>>
�ƹ�ƻ��鿴</td>
				   <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20008"<%if InStr(1, PowerList,"KSMS20008" ,1)<>0 then Response.Write( " checked") %>>
����ָ������</td>
				   <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10006"<%if InStr(1, PowerList,"KSMS10006" ,1)<>0 then Response.Write( " checked") %>>
�Զ��������</td>
				   <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20009"<%if InStr(1, PowerList,"KSMS20009" ,1)<>0 then Response.Write( " checked") %>>
�ĵ�Digg����</td>
             <td nowrap>
					<input name="PowerList" type="checkbox" id="PowerList" value="KSMS20010"<%if InStr(1, PowerList,"KSMS20010" ,1)<>0 then Response.Write( " checked") %>>
                      ���ֶһ���Ʒ</td>
             <td nowrap>
					<input name="PowerList" type="checkbox" id="PowerList" value="KSMS20014"<%if InStr(1, PowerList,"KSMS20014" ,1)<>0 then Response.Write( " checked") %>>
                     PK��Ŀ����</td>
				 </tr>

                
                </table></td>
            </tr>
          </table>
	  </td></tr>
	  </table>
	  
	  <br/>
	 <table width="99%" border="0" align="center" cellspacing="0" cellpadding="0">  
	 <tr>
	 <td height="25" class='clefttitle'><strong> �ߡ��˹���Ա�ڡ�<font color="#FF0000">�������</font>����Ȩ��</strong></td>
	 </tr>
	 </table>
	  <table width="96%" align="center" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>
          <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr> 
              <td height="25" colspan="2"> 
                <%
					IF instr(ModelPower,"other0")>0 Then
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('otherPowers','none')"" name=""otherPower"" value=""other0"" checked>")
                    ELSE
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('otherPowers','none')"" name=""otherPower"" value=""other0"">")
					END IF
					%>
                �ڲ������Ȩ��(����)</td>
            </tr>
            <tr> 
              <td height="25" colspan="2"> 
                <%
					IF instr(ModelPower,"other1")>0 Then
					  Response.Write("<input type=""radio"" name=""otherPower"" onclick=""SetPowerListValue('otherPowers','')"" value=""other1"" checked>")
                     ELSE
					  Response.Write("<input type=""radio"" name=""otherPower"" onclick=""SetPowerListValue('otherPowers','')"" value=""other1"">")
					 END IF%>
                ӵ��ָ���Ĳ��ֹ���Ȩ�� </td>
            </tr>
            <tr ID="otherPowers" <% IF instr(ModelPower,"other1")=0 then Response.Write("style=""display:none""") End IF%>> 
              <td height="25"> <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSO10000"<%if InStr(1, PowerList,"KSO10000" ,1)<>0 then Response.Write( " checked") %>> 
                    WAP�������
</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSO10001"<%if InStr(1, PowerList,"KSO10001" ,1)<>0 then Response.Write( " checked") %>>
SK�ɼ�����</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KSO10002"<%if InStr(1, PowerList,"KSO10002" ,1)<>0 then Response.Write( " checked") %>>
CC��Ƶ�������</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KSO10003"<%if InStr(1, PowerList,"KSO10003" ,1)<>0 then Response.Write( " checked") %>>
Wssͳ�Ʋ������</td>
                    <td></td>
                    <td></td>
                  </tr>
                
                </table></td>
            </tr>
          </table>
	  </td></tr>
	  </table>

<script>
function SetPowerListValue(Module,Value)
{ $('#'+Module)[0].style.display=Value;
}
</script>
<%
		

		Response.Write "  </form>"
		Response.Write "</body>"
		Response.Write "</html>"
		Response.Write "<Script Language=""javascript"">" & vbCrLf
		Response.Write "<!--" & vbCrLf
		Response.Write "function CheckForm()" & vbCrLf
		Response.Write "{ var form=document.AdminForm;" & vbCrLf
		Response.Write "   if (form.UserName.value=='')" & vbCrLf
		Response.Write "    {"
		Response.Write "     alert(""���������Ա����!"");"
		Response.Write "     form.UserName.focus();"
		Response.Write "     return false;" & vbCrLf
		Response.Write "    }" & vbCrLf
			
			If Action <> "Edit" Then
		Response.Write "   if (form.PrUserName.value=='')" & vbCrLf
		Response.Write "    {"
		Response.Write "     alert(""������ǰ̨ע���û�����!"");"
		Response.Write "     form.PrUserName.focus();"
		Response.Write "     return false;" & vbCrLf
		Response.Write "    }" & vbCrLf

		Response.Write "    if (form.PassWord.value=='')"
		Response.Write "    {"
		Response.Write "     alert(""�������ʼ����!"");"
		Response.Write "     form.PassWord.focus();"
		Response.Write "     return false;"
		Response.Write "    }"
		Response.Write "   else if (form.PassWord.value.length<6)"
		Response.Write "    {"
		Response.Write "      alert(""��ʼ���벻������6λ!"");"
		Response.Write "     form.PassWord.focus();"
		Response.Write "     return false;"
		Response.Write "    }"

		
			End If
			
		Response.Write "   if (form.RealName.value=='')" & vbCrLf
		Response.Write "    {" & vbCrLf
		Response.Write "     alert(""��������ʵ����"");" & vbCrLf
		Response.Write "     form.RealName.focus();" & vbCrLf
		Response.Write "     return false;" & vbCrLf
		Response.Write "    }" & vbCrLf
		Response.Write "   if (form.Email.value!='')" & vbCrLf
		Response.Write "   if(is_email(form.Email.value)==false)" & vbCrLf
		Response.Write "      { alert('�Ƿ���������!');" & vbCrLf
		Response.Write "        form.Email.focus();" & vbCrLf
		Response.Write "        return false;" & vbCrLf
		Response.Write "     }"
		Response.Write "    form.submit();" & vbCrLf
		Response.Write "    return true;" & vbCrLf
		Response.Write "}" & vbCrLf
		Response.Write "//-->" & vbCrLf
		Response.Write "</Script>"
		End Sub
		
		Sub AdminSave()
			Dim AdminID, GroupID, UserName,PrUserName, PassWord, ConPassWord, Locked, RealName, Sex, TelPhone, Email, Descript, TrueIP
			Dim TempObj, AdminRS, AdminSql,ComeUrl
			ComeUrl=Request.ServerVariables("HTTP_REFERER")
			AdminID = KS.G("AdminID")
			
			UserName = KS.R(KS.G("UserName"))
			PrUserName=KS.R(KS.G("PrUserName"))
			PassWord = KS.G("PassWord")

			IF PrUserName="" Then Call KS.Alert("ǰ̨ע���û���������д!",ComeUrl)
			
			PassWord = MD5(KS.R(PassWord),16)
			Locked = KS.G("Locked")
			RealName = KS.R(KS.G("RealName"))
			Sex = KS.G("Sex")
			TelPhone = KS.R(KS.G("TelPhone"))
			Email = KS.R(KS.G("Email"))
			Descript = KS.R(KS.G("Description"))
			TrueIP = KS.GetIP
			If UserName <> "" Then
					If Len(UserName) >= 100 Then
						Call KS.AlertHistory("����Ա���Ʋ��ܳ���50���ַ�!", -1)
						Set KS = Nothing
						Response.End
					End If
			 Else
					Call KS.AlertHistory("���������Ա����!", -1)
					Set KS = Nothing
					Response.End
			 End If
			 
			Dim SQL,RSC:Set RSC=Server.CreateObject("ADODB.RECORDSET")
			RSC.Open "Select ChannelID,ChannelName,BasicType,ItemName,ModelEname,ChannelStatus From KS_Channel where channelstatus=1 Order By ChannelID",Conn,1,1
			If Not RSC.Eof Then
			  SQL=RSC.GetRows(-1)
			End If
			RSC.Close

			 For I=0 To Ubound(sql,2)
			  If I=0 Then
				 ModelPower=Replace(Request("ModelPower" & sql(0,i) &"")," ","")
			  Else
				 ModelPower=ModelPower & "," & Replace(Request("ModelPower" & sql(0,i) &"")," ","")
			  End IF
			 Next
			 ModelPower=request("otherpower") &"," & request("sysset") &"," & request("userpower") & "," & request("labpower") &"," &request("modelpower") & "," &request("subsyspower")&","&request("ask")&"," & request("space") & ","& modelpower
			 PowerList=Replace(Trim(KS.G("PowerList"))," ","")
			 IF PowerList="" THEN PowerList=0

			RSC.Open "Select AdminPurview,ID From KS_Class Order By ClassID",conn,1,3
			Do While Not RSC.Eof
			    
			  If KS.FoundInArr(Replace(Request("AdminPurview")," ",""),RSC(1),",") Then
			     If KS.IsNul(RSC(0)) Then 
				  RSC(0)=UserName
				 Else
				  RSC(0)=FilterRepeat(RSC(0) & "," & UserName,",")
				 End If
				 RSC.Update
			  Else
			     If KS.IsNul(RSC(0)) Then
				  RSC(0)=""
				 Else
					' If Instr(RSC(0),",")=0 Then
					'  RSC(0)=""
					' Else
					'  RSC(0)=Replace(Replace(RSC(0),UserName &",",""),","&UserName,"")
					' End If
					 RSC(0)=DelItemInArr(RSC(0),UserName,",")
				 End If
			   	 RSC.Update
			  End If
			     on error resume next
				 Dim ENode:Set ENode=Application(KS.SiteSN&"_class").DocumentElement.SelectSingleNode("class[@ks0='" & RSC(1) & "']")
				 ENode.SelectSingleNode("@ks16").text=RSC(0)
				 If err Then err.clear
			  
			  RSC.MoveNext
			loop
			RSC.Close
			Set RSC=nothing
		   
			   
			If Request("Action") = "Add" Then
					Set TempObj = Conn.Execute("Select UserName from [KS_Admin] where UserName='" & UserName & "'")
					If Not TempObj.EOF Then
						Call KS.Alert("���ݿ����Ѵ��ڸù���Ա����!", "AdminAdd.asp")
						 Set KS = Nothing
						Response.End
					End If
					Set TempObj = Conn.Execute("Select UserName from [KS_User] where UserName='" & PrUserName & "'")
					If TempObj.BOf And TempObj.EOF Then
						Call KS.Alert("�Ҳ�����ǰ̨ע���û�!", ComeUrl)
						 Set KS = Nothing:Response.End
					End If
					IF Conn.Execute("Select Count(adminid) From KS_Admin Where PrUserName='" & PrUserName & "'")(0)>=1 Then
						Call KS.Alert("����д��ǰ̨ע���û��Ѿ��ǹ���Ա�ˣ����������!", ComeUrl)
						 Set KS = Nothing:Response.End
					End IF
				  Set AdminRS = Server.CreateObject("adodb.recordset")
				  AdminSql = "select * from [KS_Admin] Where (AdminID IS NULL)"
				  AdminRS.Open AdminSql, Conn, 1, 3
				  AdminRS.AddNew
				  AdminRS("AddDate") = Now
				  AdminRS("UserName") = UserName
				  AdminRS("PrUserName")=PrUserName
				  AdminRS("ModelPower")= ModelPower
				  AdminRS("PowerList")=PowerList
				  AdminRS("PassWord") = PassWord
				  AdminRS("Locked") = Locked
				  AdminRS("RealName") = RealName
				  AdminRS("Sex") = Sex
				  AdminRS("TelPhone") = TelPhone
				  AdminRS("Email") = Email
				  AdminRS("Description") = Descript
				  AdminRS("SuperTF") = 0
				  AdminRS("LastLoginIP") = TrueIP
				  AdminRS("LastLoginTime") = Now
				  AdminRS("LastLogOutTime") = Now
				  AdminRS("LoginTimes") = 0
				  AdminRS.Update
				  AdminRS.Close:Set AdminRS = Nothing
				  
				  '����ǰ̨�û���ʹ֮�������Ա��
				  Conn.Execute("Update KS_User Set GroupID=1 Where UserName='" & PrUserName & "'")
				  
				  Response.Write ("<script>if (confirm('��ӹ���Ա�ɹ�,���������?')) {location.href='?Action=Add';} else { location.href='KS.Admin.asp';}</script>")
			ElseIf Request("Action") = "Edit" Then
					Set TempObj = Conn.Execute("Select UserName from [KS_Admin] where AdminID<>" & AdminID & " And UserName='" & UserName & "'")
					If Not TempObj.EOF Then
						Call KS.Alert("���ݿ����Ѵ��ڸù���Ա����!", "AdminAdd.asp?AdminID=" & AdminID & "&Action=Edit")
						 Set KS = Nothing
						Response.End
					End If
				  Set AdminRS = Server.CreateObject("adodb.recordset")
				  AdminSql = "select * from [KS_Admin] Where AdminID=" & AdminID
				  AdminRS.Open AdminSql, Conn, 1, 3
				  AdminRS("UserName") = UserName
				  AdminRS("Locked") = Locked
				  AdminRS("RealName") = RealName
				  AdminRS("Sex") = Sex
				  AdminRS("TelPhone") = TelPhone
				  AdminRS("Email") = Email
				  AdminRS("Description") = Descript
				  AdminRS("ModelPower")= ModelPower
				  AdminRS("PowerList")=PowerList
				  AdminRS.Update
				  AdminRS.Close:Set AdminRS = Nothing
				  Response.Write ("<script>alert('�޸Ĺ���Ա�ɹ�!');location.href='KS.Admin.asp';</script>")
			End If
			
			
			
		End Sub
        
		'���ù���Ա����
		Sub SetAdminPass()
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		Response.Write "<link href=""Include/ModeWindow.css"" rel=""stylesheet"">"
		Response.Write "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>"
		Response.Write "<script language=""JavaScript"" src=""../ks_inc/CheckPassWord.js""></script>"
		Response.Write "<style>"
		Response.Write ".rank { border:none; background:url(../images/rank.gif) no-repeat; width:125px; height:22px; vertical-align:middle; cursor:default; }"
		Response.Write ".r0 { background-position:0 0; }"
		Response.Write ".r1 { background-position:0 -22px; }"
		Response.Write ".r2 { background-position:0 -44px; }"
		Response.Write ".r3 { background-position:0 -66px; }"
		Response.Write ".r4 { background-position:0 -88px; }"
		Response.Write ".r5 { background-position:0 -110px; }"
		Response.Write ".r6 { background-position:0 -132px; }"
		Response.Write ".r7 { background-position:0 -154px; }"
		Response.Write "</style>"
		Response.Write "<title>���ù���Ա����</title>"
		Response.Write "</head>"
		
		Dim AdminID, UserName, PassWord,RSObj
		AdminID = Request("AdminID")
		If AdminID <> "" Then
		   Set RSObj = Server.CreateObject("AdoDb.RecordSet")
		   RSObj.Open "Select * From KS_Admin Where AdminID=" & AdminID, Conn, 1, 3
		   If Not RSObj.EOF Then UserName = Trim(RSObj("UserName"))
		   RSObj.Close:Set RSObj = Nothing
		Else
		  UserName=KS.C("AdminName")
		End If
		
		     If Request("Flag") = "SetOK" Then
			   If Trim(Request.Form("PassWord")) <> Trim(Request.Form("ConPassWord")) Then
				Response.Write ("<Script>alert('ȷ����������!');history.back(-1);</script>")
				Exit Sub
				Response.End
			   Else
			    Set RSObj = Server.CreateObject("AdoDb.RecordSet")
		         RSObj.Open "Select * From KS_Admin Where UserName='" & UserName & "'", Conn, 1, 3
				 RSObj("PASSWord") = MD5(KS.R(Trim(Request.Form("PassWord"))),16)
				 RSObj.Update
				 Dim PrUserName:PrUserName=RSObj("PrUserName")
				  RSObj.Close: Set RSObj = Nothing
				  If UserName=KS.C("UserName") Then  Response.Cookies(KS.SiteSn)("AdminPass")=MD5(KS.R(Trim(Request.Form("PassWord"))),16)
				  
				  If KS.ChkClng(request("UpdateUserPass"))=1 Then
				    Conn.Execute("Update KS_User Set [PassWord]='" &MD5(KS.R(Trim(Request.Form("PassWord"))),16) &"' Where UserName='" & PrUserName & "'")
					Response.Cookies(KS.SiteSn)("Password")=MD5(KS.R(Trim(Request.Form("PassWord"))),16)
				  End If
				  
				 Response.Write ("<Script>alert('�������óɹ�!!!');window.close();</script>")
			   End If
			 End If
			 
		Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
		Response.Write "  <form action=""?Action=SetPass"" name=""AdminForm"" method=""post"">"
		Response.Write "   <input name=""Flag"" type=""hidden"" id=""Flag"" value=""SetOK"">"
		Response.Write "   <input name=""AdminID"" type=""hidden"" value=""" & AdminID & """><br>"
		Response.Write "  <table width=""99%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"
		Response.Write "    <tr>"
		Response.Write "      <td>"
		Response.Write "      <FIELDSET align=center>" & vbCrLf
		Response.Write "        <LEGEND align=left>��������</LEGEND>" & vbCrLf
		Response.Write "        <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbCrLf
		Response.Write "          <tr>"
		Response.Write "            <td width=""179"" height=""25"" align=""center""> <div align=""center"">����Ա��</div></td>" & vbCrLf
		Response.Write "            <td width=""542"" height=""25"">"
		Response.Write ("<input name=""UserName"" Readonly value=""" & UserName & """ type=""text"" id=""UserName"" size=""30"" class=""textbox"">")
		Response.Write "              ���ڵ�¼��̨������</td>" & vbCrLf
		Response.Write "          </tr>"
		
		Response.Write "          <tr>"
		Response.Write "            <td height=""25"" align=""center""> <div align=""center"">�� �� ��</div></td>"
		Response.Write "           <td height=""25""> <input name=""PassWord"" type=""password"" size=""34"" class=""textbox"" onkeyup=""javascript:setPasswordLevel(this, document.getElementById('passwordLevel'))"" onmouseout=""javascript:setPasswordLevel(this, document.getElementById('passwordLevel'))"" onblur=""javascript:setPasswordLevel(this, document.getElementById('passwordLevel'))"">" & vbCrLf
		Response.Write "              ������6λ </td>"
		Response.Write "          </tr>"
		Response.Write "          <tr>"
		Response.Write "             <td align=""center"" height=""25"">����ǿ��</td>"
		Response.Write "                <td><input name=""Input"" disabled=""disabled"" class=""rank r0"" id=""passwordLevel"" /></td>"
		Response.Write "           </tr>"
		Response.Write "          <tr>"
		Response.Write "            <td height=""25"" align=""center"">ȷ������</td>" & vbCrLf
		Response.Write "            <td height=""25""> <input name=""ConPassWord""  type=""password"" size=""34"" class=""textbox"">" & vbCrLf
		Response.Write "              ͬ��</td>"
		Response.Write "          </tr>"
		
		Response.Write "          <tr>"
		Response.Write "            <td height=""25"" align=""center"">����ǰ̨����</td>" & vbCrLf
		Response.Write "            <td height=""25""> <label><input name=""UpdateUserPass""  type=""checkbox"" value=""1"" checked>��</label> <font color=red>���ѡ����,��ôǰ̨��Ա���ĵĵ�¼���뽫�ͺ�̨��һ��</font></td>"
		Response.Write "          </tr>"
		
		Response.Write "        </table>"
		Response.Write "         </FIELDSET>" & vbCrLf
		Response.Write "       </td>" & vbCrLf
		Response.Write "    </tr>"
		Response.Write "    </table>" & vbCrLf
		Response.Write "  <table width=""100%"" height=""30"" border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf
		Response.Write "    <tr>"
		Response.Write "      <td height=""40"" align=""center"">" & vbCrLf
		Response.Write "        <input type=""button"" class='textbox' name=""Submit"" Onclick=""CheckForm()"" value="" ȷ �� "">"
		Response.Write "        <input type=""button"" class='textbox' name=""Submit2"" onclick=""window.close()"" value="" ȡ �� "">" & vbCrLf
		Response.Write "      </td>" & vbCrLf
		Response.Write "    </tr>"
		Response.Write "  </table>"
		Response.Write "  </form>" & vbCrLf
		Response.Write "</body>"
		Response.Write "</html>" & vbCrLf
		Response.Write "<Script Language=""javascript"">" & vbCrLf
		Response.Write "<!--" & vbCrLf
		Response.Write "function CheckForm()" & vbCrLf
		Response.Write "{ var form=document.AdminForm;" & vbCrLf
		Response.Write "    if (form.PassWord.value=='')" & vbCrLf
		Response.Write "    {" & vbCrLf
		Response.Write "     alert(""������������!"");" & vbCrLf
		Response.Write "     form.PassWord.focus();" & vbCrLf
		Response.Write "     return false;" & vbCrLf
		Response.Write "    }" & vbCrLf
		Response.Write "    else if (form.PassWord.value.length<6)" & vbCrLf
		Response.Write "    {" & vbCrLf
		Response.Write "      alert(""��ʼ���벻������6λ!"");"
		Response.Write "     form.PassWord.focus();" & vbCrLf
		Response.Write "     return false;" & vbCrLf
		Response.Write "    }" & vbCrLf
		Response.Write "   if (form.ConPassWord.value=='')" & vbCrLf
		Response.Write "    {"
		Response.Write "     alert(""������ȷ������!"");" & vbCrLf
		Response.Write "     form.ConPassWord.focus();" & vbCrLf
		Response.Write "     return false;" & vbCrLf
		Response.Write "    }" & vbCrLf
		Response.Write "   else if(form.ConPassWord.value.length<6)" & vbCrLf
		Response.Write "    {"
		Response.Write "     alert(""ȷ�����벻������6λ!"");" & vbCrLf
		Response.Write "     form.ConPassWord.focus();" & vbCrLf
		Response.Write "     return false;" & vbCrLf
		Response.Write "    }"
		Response.Write "   if (form.PassWord.value!=form.ConPassWord.value)" & vbCrLf
		Response.Write "    {"
		Response.Write "    alert(""������������벻һ��!"");" & vbCrLf
		Response.Write "     form.PassWord.focus();" & vbCrLf
		Response.Write "     return false;" & vbCrLf
		Response.Write "    }" & vbCrLf
		Response.Write "    form.submit();" & vbCrLf
		Response.Write "    return true;" & vbCrLf
		Response.Write "}"
		Response.Write "//-->"
		Response.Write "</Script>"
		End Sub
		
		'ɾ������Ա
		Sub AdminDel()
		Dim k, AdminID,RSObj
		AdminID = Trim(KS.G("AdminID"))
		AdminID = Split(AdminID, ",")
		For k = LBound(AdminID) To UBound(AdminID)
			   Set RSObj = Conn.Execute("Select SuperTF,PrUserName From KS_Admin Where  AdminID=" & AdminID(k))
			   If Not RSObj.EOF Then
				 If RSObj("SuperTF") = 1 Then
				  Response.Write ("<script>alert('ϵͳĬ�Ϲ���Ա����ɾ��!');location.href='KS.Admin.asp';</script>")
				 Else
				  '����ǰ̨ע���Ա��ʹ֮��Ϊע���Ա���
				  Conn.Execute("Update KS_User Set GroupID=3 Where UserName='" & RSObj("PrUserName") & "'")
				  Conn.Execute ("Delete From KS_Admin Where AdminID =" & AdminID(k))
				 End If
			  End If
			  RSObj.Close
		Next
		Set RSObj = Nothing
		Response.Write ("<script>location.href='KS.Admin.asp';</script>")
		End Sub
		
		
 Sub ClassList(ChannelID)
 %>
 <div style="border: 5px solid #E7E7E7;height:150; overflow: auto; width:100%"> 
                        <table border="0" cellspacing="0" cellpadding="0">
                          <% 
					  Dim Node, CheckStr,SpaceStr,TJ,k  
				      KS.LoadClassConfig
					  For Each Node In Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks12=" & ChannelID&"]")                     
	                  if username<>"" and KS.FoundInArr(Node.SelectSingleNode("@ks16").text,username,",")=true then CheckStr=" checked"
					  SpaceStr="&nbsp;&nbsp;&nbsp;&nbsp;"
					  TJ=Node.SelectSingleNode("@ks10").text
					  If TJ>1 Then
						 For k = 1 To TJ - 1
							SpaceStr = SpaceStr & "&nbsp;&nbsp;&nbsp;&nbsp;"
						 Next
					    SpaceStr = SpaceStr &"<img src=""Images/Folder/folderclosed.gif"">"
					  Else
					    spacestr="<img src=""Images/Folder/domain.gif"" width=""26"" height=""24"">"
					  End If
					  %>
                          
                          <tr> 
                            <td><table border="0" cellspacing="0" cellpadding="0">
                                <tr align="left" class="TempletItem"> 
                                  <td><%=SpaceStr%></td>
                                  <td><input name="AdminPurview" type="checkbox" value="<% =Node.SelectSingleNode("@ks0").text %>"<%=CheckStr%>> 
                                    <% = Node.SelectSingleNode("@ks1").text %>                                     
								 </td>
                                </tr>
                              </table></td>
                          </tr>
                          <%
	                     CheckStr = ""
	                Next
					   %>
                        </table>
                      </div>
 <%
 End Sub
 
 Sub BasePurview(PowerList,SQL,I)
 %>
     <tr> 
                    <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10002"<%if InStr(1, PowerList,"M" & SQL(0,I)&"10002" ,1)<>0 then Response.Write( " checked") %>>
                      ���<%=sql(3,i)%></td>
                    <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10003"<%if InStr(1, PowerList,"M"&SQL(0,I)&"10003" ,1)<>0 then Response.Write( " checked") %>>
                      �༭<%=sql(3,i)%></td>
                    <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10004"<%if InStr(1, PowerList,"M" & SQL(0,I) &"10004" ,1)<>0 then Response.Write( " checked") %>>
                      ɾ��<%=sql(3,i)%></td>
                    <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10005"<%if InStr(1, PowerList,"M" & SQL(0,I) &"10005" ,1)<>0 then Response.Write(" checked") %>>
��������</td>
                    <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10006"<%if InStr(1, PowerList,"M" & SQL(0,I) & "10006" ,1)<>0 then Response.Write(" checked") %>>
                      ����ר��</td>
                    <td width="13%"> <input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10007"<%if InStr(1, PowerList,"M" & SQL(0,I) & "10007" ,1)<>0 then Response.Write(" checked") %>>
                      ����JS</td>
                  </tr>
                  <tr> 
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10008"<%if InStr(1, PowerList,"M" & SQL(0,I) &"10008" ,1)<>0 then Response.Write( " checked") %>>
                      ����<%=sql(3,i)%></td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10009"<%if InStr(1, PowerList,"M" & SQL(0,I) & "10009" ,1)<>0 then Response.Write(" checked") %>>
�ϴ��ļ�</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10011"<%if InStr(1, PowerList,"M" & SQL(0,I) & "10011" ,1)<>0 then Response.Write(" checked") %>>
����ճ��</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10012"<%if InStr(1, PowerList,"M" & SQL(0,I) & "10012" ,1)<>0 then Response.Write(" checked") %>>
ǩ��<%=sql(3,i)%></td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20005"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20005" ,1)<>0 then Response.Write(" checked") %>>
����<%=sql(3,i)%></td>
                    <td></td>
                  </tr>
<%
 End Sub
 
 'ȥ��������ظ���
 Function FilterRepeat(byval str,spliter)
   if KS.IsNul(str) Then Exit Function
   Dim strA:strA=Split(str,spliter)
   Dim I,temp,newstr
   For I=0 To Ubound(Stra)
      If KS.FoundInArr(temp,strA(i),",")=false Then
	    if newstr="" then
		 newstr=stra(i)
		else
		 newstr=newstr & spliter & stra(i)
		end if
		temp=temp & "," & stra(i)
	  End If
   Next
   FilterRepeat=newstr
 End Function
 
 '��������ɾ��ָ����
 Function DelItemInArr(byval str,byval delstr,spliter)
   if KS.IsNul(str) Then Exit Function
   Dim strA:strA=Split(str,spliter)
   Dim I,temp,newstr
   For I=0 To Ubound(Stra)
      If lcase(strA(i))<>lcase(delstr) Then
	    if newstr="" then
		 newstr=stra(i)
		else
		 newstr=newstr & spliter & stra(i)
		end if
	  End If
   Next
   DelItemInArr=newstr
 End Function
 
End Class
%> 
