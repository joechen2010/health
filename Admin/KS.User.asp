<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../Plus/md5.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<!--#include file="../API/cls_api.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Admin_User
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_User
        Private KS,KSCls
		Private MaxPerPage
		Private rsAdmin,sqlAdmin
		Private UserID,UserSearch,Keyword,strField,CurrentPage,sql,FoundErr,RS,TotalPut,TotalPages,I
		Private Action,ComeUrl,strFileName
		Private ValidDays,tmpDays,BeginID,EndID
		Private ErrMsg
		Private Sub Class_Initialize()
		  MaxPerPage=20
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub
       Sub Kesion()
	   
			If KS.G("Action")="CheckUserName" Then Call CheckUserName():Response.End()
            Response.Write "<html>"
			Response.Write"<head>"
			Response.Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			Response.Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			Response.Write"<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>" & vbCrLf
			Response.Write"<script language=""JavaScript"" src=""../KS_Inc/jquery.js""></script>" & vbCrLf
			Response.Write"</head>"
			Response.Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			Response.Write"	<ul id='mt'> "
			Response.Write "<div id='mtl'>���ٲ����û���</div><li><a href=""KS.User.asp?Action=Search"">�����û�</a></li>&nbsp;|&nbsp;<a href=""?UserSearch=12"">�����û�</a>&nbsp;|&nbsp;<a href=""?UserSearch=1"">����ס���û�</a>&nbsp;|&nbsp;<a href=""?UserSearch=2"">���й���Ա</a>&nbsp;|&nbsp;<a href=""?UserSearch=3"">��������Ա</a>&nbsp;|&nbsp;<a href=""?UserSearch=4"">���ʼ�����</a>&nbsp;|&nbsp;<a href=""?UserSearch=5"">24Сʱ�ڵ�¼</a>&nbsp;|&nbsp;<a href=""?UserSearch=6"">24Сʱ��ע��</a>"
			Response.Write	" </ul>"
		     If Not KS.ReturnPowerResult(0, "KMUA10002") Then
			  Response.Write ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back();';</script>")
			  Call KS.ReturnErr(1, "")
			End If

		ComeUrl=Cstr(Request.ServerVariables("HTTP_REFERER"))
		keyword		= Trim(request("keyword"))
		strField	= Trim(request("Field"))
		UserSearch	= KS.ChkClng(request("UserSearch"))
		Action		= Trim(request("Action"))
		UserID		= Trim(Request("UserID"))
		strFileName	= "KS.User.asp"
		CurrentPage	= KS.ChkClng(request("page"))
		if keyword<>"" then keyword=KS.R(keyword)
		%>
		<SCRIPT language=javascript>
		function unselectall()
		{
			if(document.myform.chkAll.checked){
			document.myform.chkAll.checked = document.myform.chkAll.checked&0;
			} 	
		}
		
		function CheckAll(form)
		{
		  for (var i=0;i<form.elements.length;i++)
			{
			var e = form.elements[i];
			if (e.Name != "chkAll"  && e.disabled==false)
			   e.checked = form.chkAll.checked;
			}
		}
		</SCRIPT>
		</head>
		<%
		Select Case Action
		Case "Add"            call AddUser()
		Case "SaveAdd"	      call SaveAdd()
		Case "Modify"         call Modify()
		Case "SaveModify"     call SaveModify()
		Case "Del"            call DelUser()
		Case "Lock"	          call locked()
		Case "UnLock"         call Unlocked()
		Case "Verify"	      call verify(0)
		Case "UnVerify"       call verify(2)
		Case "Active"         call Unlocked()
		Case "Move"	          call MoveUser()
		Case "AddMoney"	      call AddMoney()
		Case "SaveAddMoney"	  call SaveAddMoney()
		Case "AddZJ"          call AddZJ()
		Case "SaveAddZJ"      call SaveAddZJ
		Case "Search"         Call ShowSearch()
		Case "ShowDetail"	  Call ShowDetail()
		Case Else	call main()
		End Select
		if FoundErr=True then KS.ShowError(ErrMsg)
		If Action<>"ShowDetail" Then
		Response.Write "<div style=""text-align:center;color:#003300"">-----------------------------------------------------------------------------------------------------------</div>"
		Response.Write "<div style=""height:30px;text-align:center"">KeSion CMS V 6.5, Copyright (c) 2006-2010 <a href='http://www.kesion.com/' target=""_blank""><font color=#cc6600>KeSion.Com</font></a>. All Rights Reserved . </div>"
		End If
		End Sub
		
		Sub Main()
		    Dim GroupID:GroupID=KS.G("GroupID")
				dim strGuide ,sSQL,Param
				strGuide="<table style='margin-top:0px' width='100%' align='center' border='0' cellpadding='0' cellspacing='1'><tr class='list'><td align='center' height='25'>&nbsp;"
				sSQL = " UserID,UserName,GroupID,ChargeType,Point,BeginDate,LastLoginIP,LastLoginTime,LoginTimes ,locked,Edays"
				Select Case UserSearch
				 Case 1
				    Param="locked=1"
					strGuide=strGuide & "���б���ס���û�"
                Case 2
					Param="groupid=1"
					strGuide=strGuide & "���й���Ա��ݵ��û�"
                Case 3
					Param="locked=2"
					strGuide=strGuide & "������Ա��֤�û�"
                Case 4
					Param="locked=3"
					strGuide=strGuide & "���ʼ���֤���û�"
				Case 5
				   Param="datediff(" & DataPart_H & ",LastLoginTime," & SqlNowString & ")<25"
				   strGuide=strGuide & "���24Сʱ�ڵ�¼���û�"
				Case 6
				    Param="datediff(" & DataPart_H & ",RegDate," & SqlNowString & ")<25"
					strGuide=strGuide & "���24Сʱ��ע����û�"
				Case 10
					param="GroupID=" & GroupID
					strGuide=strGuide & KS.GetUserGroupName(GroupID)
				Case 11
					UserID = KS.ChkClng(UserID)
					if UserID>0 then
						param="UserID="&UserID&""
					else 
						Dim strsql
						strsql=""
						if request("username")<>"" then
							if request("usernamechk")="yes" then
								strsql=strsql & " username='"&request("username")&"'"
							else
								strsql=strsql &" username like '%"&request("username")&"%'"
							end if
						end if
						if cint(request("GroupID"))>0 then
							if strsql="" then
								strsql=strsql & " GroupID="&request("GroupID")&""
							else
								strsql=strsql & " and GroupID="&request("GroupID")&""
							end if
						end if
						if request("Email")<>"" then
							if strsql="" then
								strsql=strsql & " Email like '%"&request("Email")&"%'"
							else
								strsql=strsql & " and Email like '%"&request("Email")&"%'"
							end if
						end if
		            '======��������=======
						dim Tsqlstr
						if request("loginT")<>"" then
							if request("loginR")="more" then
								Tsqlstr=" LoginTimes >= "&KS.Chkclng(request("loginT"))
							else
								Tsqlstr=" LoginTimes <= "&KS.Chkclng(request("loginT"))
							end if 	
							if strsql="" then 
								strsql=Tsqlstr
							else
								strsql=strsql & " and" & Tsqlstr
							end if
						end if
		
						if request("vanishT")<>"" then
							if request("vanishR")="more" then
								if DataBaseType=1 then
									Tsqlstr=" datediff(d,LastLoginTime,"&SqlNowString&") >= "&KS.Chkclng(request("vanishT"))&""
								else
									Tsqlstr=" datediff('d',LastLoginTime,"&SqlNowString&") >= "&KS.Chkclng(request("vanishT"))&""
								end if
							else
							   if DataBaseType=1 then
								Tsqlstr=" datediff(d,LastLoginTime,"&SqlNowString&") <= "&KS.Chkclng(request("vanishT"))&""
								else
								Tsqlstr=" datediff('d',LastLoginTime,"&SqlNowString&") <= "&KS.Chkclng(request("vanishT"))&""
								end if
							end if 	
							if strsql="" then 
								strsql=Tsqlstr
							else
								strsql=strsql & " and" & Tsqlstr
							end if
						end if
		
						if request("regT")<>"" then
							if request("regR")="more" then
							  if DataBaseType=1 then
								Tsqlstr=" datediff(d,RegDate,"&SqlNowString&") >= "&KS.Chkclng(request("regT"))
							   else
								Tsqlstr=" datediff('d',RegDate,"&SqlNowString&") >= "&KS.Chkclng(request("regT"))
							   end if
							else
							  if DataBaseType=1 then
								Tsqlstr=" datediff(d,RegDate,"&SqlNowString&") <= "&KS.Chkclng(request("regT"))
							  else
								Tsqlstr=" datediff('d',RegDate,"&SqlNowString&") <= "&KS.Chkclng(request("regT"))
							  end if
							end if 	
							if strsql="" then 
								strsql=Tsqlstr
							else
								strsql=strsql & " and" & Tsqlstr
							end if
						end if
		
						if request("artcleT")<>"" then
							if request("artcleR")="more" then
								Tsqlstr=" (select count(id) from ks_iteminfo where inputer=ks_user.username) >= "&KS.Chkclng(request("artcleT"))
							else
								Tsqlstr=" (select count(id) from ks_iteminfo where inputer=ks_user.username) <= "&KS.Chkclng(request("artcleT"))
							end if 	
							if strsql="" then 
								strsql=Tsqlstr
							else
								strsql=strsql & " and" & Tsqlstr
							end if
						end if
		              '======������������======
						If strsql = "" Then
							FoundErr=True
							ErrMsg=ErrMsg & "<br><li>��ָ������������</li>"
							Exit Sub
						End If
						If Request("Searchmax") = "" Or Not Isnumeric(Request("Searchmax")) Then
							param=strsql
						Else
						    param=strsql
							'Sql = "Select top "&Request("Searchmax")&" "&sSQL&" From KS_User Where " & strsql & " order by UserID desc"
						End If
					end if '''ID
					strGuide=strGuide & "��ѯ�����"
				Case Else
					Param="1=1"
					strGuide=strGuide & "�����û�"
				End Select
				strGuide=strGuide & "</td><td width='150' align='center'>"
				if FoundErr=True then Exit Sub
				If KS.C("SuperTF")<>"1" then Param=Param & " and (groupid<>1 or username='" & KS.C("AdminName") & "')"
				
				if CurrentPage < 1 then CurrentPage=1
				sql=KSCls.GetPageSQL("KS_User","userid",MaxPerPage,CurrentPage,1,Param,sSQL)
				
				
				Set rs=Server.CreateObject("Adodb.RecordSet")
				rs.Open sql,Conn,1,1
				if rs.eof and rs.bof then
					TotalPut=0
					Response.Write strGuide & "���ҵ� <font color=#ff6600>0</font> ���û�&nbsp;&nbsp;&nbsp;&nbsp;</td></tr></table>"
					rs.Close:set rs=Nothing
				else
					'TotalPut=rs.recordcount
					TotalPut=Conn.Execute("Select count(userid) from [KS_User] where " & Param)(0)
					'Response.Write strGuide & "���ҵ� <font color=#ff6600>" & TotalPut & "</font> ���û�&nbsp;&nbsp;&nbsp;&nbsp;</td></tr></table>"
					if (TotalPut mod MaxPerPage)=0 then
						TotalPages = TotalPut \ MaxPerPage
					else
						TotalPages = TotalPut \ MaxPerPage + 1
					end if
					SQL = rs.GetRows(MaxPerPage)
					rs.Close:set rs=Nothing
					ShowContent
				end if
		End Sub
		
		Sub ShowContent()
		%>
		  <table width="100%" style="border-top:1px #CCCCCC solid" border="0" align="center" cellspacing="0" cellpadding="0">
		  		  <form name="myform" method="Post" action="KS.User.asp" onSubmit="return confirm('ȷ��Ҫִ��ѡ���Ĳ�����');">
				  <tr class='sort'>
					<td width="30" align="center">ѡ��</td>
					<td width="30" align="center">ID</td>
					<td width="80" align="center"> �û���</td>
					<td height="22" align="center">�����û���</td>
					<td align="center">ʣ�����/����</td>
					<td height="22" align="center">����¼IP</td>
					<td align="center">����¼ʱ��</td>
					<td width="60" align="center">��¼����</td>
					<td width="40" align="center">״̬</td>
					<td width="120" align="center">����</td>
				  </tr>
				  <%
				For i=0 To Ubound(SQL,2)
				 %>
				  <tr height="23" class='list' id='u<%=SQL(0,i)%>' onclick="chk_iddiv('<%=SQL(0,i)%>')" onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
					<td class='splittd' width="30" align="center"><input <%If KS.C("SuperTF")<>1 and SQL(2,i)=4 and  SQL(1,i)<>KS.C("AdminName") Then Response.Write " Disabled" %> name="UserID" type="checkbox"  onclick="chk_iddiv('<%=SQL(0,i)%>')" id='c<%=SQL(0,i)%>'  value="<%=SQL(0,i)%>"></td>
					<td class='splittd' width="30" align="center"><%=SQL(0,i)%></td>
					<td class='splittd' width="80" align="center"><%
					If KS.C("SuperTF")<>1 and SQL(2,i)=4 and  SQL(1,i)<>KS.C("AdminName") then
					 response.write "<font color=red>" & SQL(1,i) & "</font>"
					else
					Response.Write "<a href='KS.User.asp?Action=ShowDetail&UserID=" & SQL(0,i) & "'>" & SQL(1,i) & "</a>"
					end if
					%>
					</td>
					<td class='splittd' align="center"><font color=blue><%=KS.GetUserGroupName(SQL(2,i))%></font></td>
					<td class='splittd' align="center">
					<%
				if SQL(3,i)=1 then
					if SQL(4,i)<=0 then
						Response.Write "<font color=#ff6600>" & SQL(4,i) & "</font> ��"
					else
						if SQL(4,i)<=10 then
							Response.Write "<font color=blue>" & SQL(4,i) & "</font> ��"
						else
							Response.Write SQL(4,i) & " ��"
						end if
					end if
				elseif SQL(3,i)=2 then
				    ValidDays=SQL(10,i)
					tmpDays = ValidDays-DateDiff("D",SQL(5,i),now())
					if tmpDays<=10 then
						Response.Write "<font color=#ff0033>" & tmpDays & "</font> ��"
					else
						Response.Write "<font color=#0000ff>" & tmpDays & "</font> ��"
					end if
				else
				   response.write "<font color=red>������</font>"
				end if
				%></td>
					<td class='splittd' align="center"> <%
				if SQL(6,i)<>"" then
					Response.Write SQL(6,i)
				else
					Response.Write "&nbsp;"
				end if%> </td>
					<td class='splittd' align="center"> <%=SQL(7,i)%> </td>
					<td class='splittd' width="60" align="center"><%=SQL(8,i)%> </td>
					<td class='splittd' width="40" align="center"><%
				select case SQL(9,i)
				   case 1 Response.Write "<font color=#ff6600>������</font>"
				   case 2 response.write "<font color=blue>�����</font>"
				   case 3 response.write "<font color=green>������</font>"
				   case else
					Response.Write "����"
				end select%></td>
					<td class='splittd' width="120" align="center"><%
				If KS.C("SuperTF")<>1 and SQL(2,i)=4 and  SQL(1,i)<>KS.C("AdminName") then
					 response.write "---"
				else	
					Response.Write "<a href='KS.User.asp?Action=Modify&UserID=" & SQL(0,i) & "'>��</a>&nbsp;"
					if SQL(2,i)<>1 then  '����Ա�ж�
						if SQL(9,i)=0 then
							Response.Write "<a href='KS.User.asp?Action=Lock&UserID=" & SQL(0,i) & "'>��</a>&nbsp;"
						else
							Response.Write "<a href='KS.User.asp?Action=UnLock&UserID=" & SQL(0,i) & "'>��</a>&nbsp;"
						end if
						Response.Write "<a href='KS.User.asp?Action=Del&UserID=" & SQL(0,i) & "' onClick='return confirm(""ȷ��Ҫɾ�����û���"");'>ɾ</a>&nbsp;"
						 If SQL(3,i)=1 Then
						   Response.Write "<a href='KS.User.asp?Action=AddMoney&UserID=" & SQL(0,i) & "'>������</a>"
						 ElseIf SQL(3,I)=2 Then
						   Response.Write "<a href='KS.User.asp?Action=AddMoney&UserID=" & SQL(0,i) & "'>������</a>"
						 End IF
					end if
					%> <a href='KS.User.asp?Action=AddZJ&UserID=<%=SQL(0,I)%>'>����</a>
				<%end if%>
				</td>
				  </tr>
			<%Next%>

		  <tr class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
			<td colspan=10 height="30">&nbsp;<label><input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
					  ѡ��</label>&nbsp;<strong>������</strong> 
					  <input name="Action" type="radio" value="Del" checked onClick="this.form.GroupID.disabled=true">ɾ�� 
					  <input name="Action" type="radio" value="Lock" onClick="this.form.GroupID.disabled=true">���� 
					  <input name="Action" type="radio" value="UnLock" onClick="this.form.GroupID.disabled=true">���� 
					  <input name="Action" type="radio" value="Verify" onClick="this.form.GroupID.disabled=true">��� 
					  <input name="Action" type="radio" value="UnVerify" onClick="this.form.GroupID.disabled=true">���� 
					  <input name="Action" type="radio" value="Active" onClick="this.form.GroupID.disabled=true">���� 
					  <input name="Action" type="radio" value="Move" onClick="this.form.GroupID.disabled=false">�ƶ���
					  <select name="GroupID" id="GroupID" disabled>
						<%=KS.GetUserGroup_Option(3)%>
					  </select>
					  &nbsp;<input type="submit" name="Submit" class='button' value=" ִ �� " >&nbsp;<input class='button' type="button" name="Submit" value="�����ʼ�" onclick="this.form.action='KS.UserMail.asp?InceptType=2';this.form.submit()" >&nbsp;<input class='button' type="button" name="Submit" value="���Ͷ���" onclick="this.form.action='KS.UserMessage.asp?action=new';this.form.submit()" > </td>
		  </tr></form>
		<tr valign=middle class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
			<td colspan="10" align="right">
			<%
			 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
			%>
		</td>
		</tr>
		</table>
		<%
		End Sub
		
		Sub ShowSearch()
		%>
		<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1">
		<form name="form2" method="get" action="KS.User.asp">
		<tr Class="sort">
			<td height="25" colspan="2" align="center"><strong>�߼���ѯ</strong></td>
		</tr>
		<tr class="tdbg">
			<td width="100" height="25" class="clefttitle" align="right"><strong>ע������:</strong></td>
			<td>�ڼ�¼�ܶ���������������Խ���ѯԽ�����뾡�����ٲ�ѯ�����������ʾ��¼��Ҳ����ѡ�����</td>
		</tr>
		<!--
		<tr class="tdbg">
			<td class="clefttitle" align="right"><strong>����ѯ��¼��:</strong></td>
			<td><input class="textbox" size="45" name="searchMax" type="text" value="100"></td>
		</tr>
		-->
		<tr class="tdbg">
			<td class="clefttitle" align="right"><strong>�û�ID:</strong></td>
			<td><input class="textbox" size="45" name="userid" type="text"></td>
		</tr>
		<tr class="tdbg">
			<td class="clefttitle" align="right"><strong>�û���:</strong></td>
			<td><input class="textbox" size="45" name="username" type="text">&nbsp;<input type="checkbox" name="usernamechk" value="yes" checked>�û�������ƥ��</td>
		</tr>
		<tr class="tdbg">
			<td class="clefttitle" align="right"><strong>�û���:</strong></td>
			<td>
			<select size="1" name="GroupID">
			<option value="0" selected>����</option>
			<%=KS.GetUserGroup_Option(0)%>
			</select>
		  </td>
		</tr>
		<tr class="tdbg">
			<td class="clefttitle" align="right"><strong>Email����:</strong></td>
			<td><input class="textbox" size="45" name="Email" type=text></td>
		</tr>
		</table>
		<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1">
		<!--��������-->
		<tr class="sort">
			<td height="23" colspan="2" align="left">�����ѯ&nbsp;��ע�⣺ 
		  <����> �� <����> ��Ĭ�ϰ��� <����>������������ʹ�ô����� ��</td>
		</tr>
		<tr class="tdbg">
			<td>��¼����:
		  <input type=radio value=more name="loginR" checked ID="Radio1">&nbsp;����&nbsp;<input type=radio value=less name="loginR" ID="Radio2">&nbsp;����&nbsp;&nbsp;<input class="textbox" size=5 name="loginT" type=text ID="Text1"> ��&nbsp;&nbsp;</td>
			<td>��ʧ����:
		  <input type=radio value=more name="vanishR" checked ID="Radio3">&nbsp;����&nbsp;<input type=radio value=less name="vanishR" ID="Radio4">&nbsp;����&nbsp;&nbsp;<input class="textbox" size=5 name="vanishT" type=text ID="Text2"> ��&nbsp;&nbsp;</td>
		</tr>
		<tr class="tdbg">
			<td width="50%">ע������:
		  <input type=radio value=more name="regR" checked ID="Radio5">&nbsp;����&nbsp;<input type=radio value=less name="regR" ID="Radio6">&nbsp;����&nbsp;&nbsp;<input class="textbox" size=5 name="regT" type=text ID="Text3"> ��&nbsp;&nbsp;</td>
			<td width="50%">��������:
		  <input type=radio value=more name="artcleR" checked ID="Radio7">&nbsp;����&nbsp;<input type=radio value=less name="artcleR" ID="Radio8">&nbsp;����&nbsp;&nbsp;<input class="textbox" size=5 name="artcleT" type=text ID="Text4"> ƪ&nbsp;&nbsp;</td>
		</tr>
		<!--������������-->
		<tr class="tdbg">
		  <td width="100%" colspan="2" align="center"><input name="submit" class='button' type=submit value="   ��  ��   "></td>
		</tr>
		<input name="UserSearch" type="hidden" id="UserSearch" value="11">
		</table>
		</form>
		<%
		end sub
		
		sub AddUser()
		 Dim GroupID:GroupID=KS.ChkClng(KS.G("GroupID"))
		 If GroupID=0 Then GroupID=3
		%>
		<SCRIPT language=javascript>
		function CheckForm()
		{
		  if(document.myform.UserName.value=="")
			{
			  alert("�û�������Ϊ�գ�");
			  document.myform.UserName.focus();
			  return false;
			}
		  if(document.myform.Password.value=="")
			{
			  alert("�û����벻��Ϊ�գ�");
			  document.myform.Password.focus();
			  return false;
			}
		  if(document.myform.Password.value!=document.myform.PwdConfirm.value)
			{
			  alert("������������벻һ�£�");
			  document.myform.Password.focus();
			  return false;
			}
		
		  if(document.myform.Question.value=="")
			{
			  alert("�������ⲻ��Ϊ�գ�");
			  document.myform.Question.focus();
			  return false;
			}
		  if(document.myform.Answer.value=="")
			{
			  alert("����𰸲���Ϊ�գ�");
			  document.myform.Answer.focus();
			  return false;
			}
		  if(document.myform.Email.value=="")
			{
			  alert("�û�Email����Ϊ�գ�");
			  document.myform.Email.focus();
			  return false;
			}
		  if((document.myform.Password.value)!=(document.myform.PwdConfirm.value))
			{
			  alert("��ʼ������ȷ�����벻ͬ��");
			  document.myform.PwdConfirm.select();
			  document.myform.PwdConfirm.focus();	  
			  return false;
			}
		}
		 checkaccount=function(val){
		  if(val=='')
		  {
			alert('�������û�����!');
			$('input[name=UserName]').focus();
			return false;
		  }
		  if(val.length<6||val.length>10)
		  {
			alert('�û����ȱ�����ڵ���6λС�ڵ���10λ!');
			$('input[name=UserName]').focus();
			return false;
		  }
		  window.open('?action=CheckUserName&username='+val,'','width=0,height=0');
		 }
		</script>
		<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1">
		<FORM name="myform" action="KS.User.asp" method="post" onsubmit="return(CheckForm());">
			<TR class="sort">
			  <TD height="22" colspan="4" align="center"><span style="margin-top:10px;text-align:center;font-weight:bold">������û�</span></TD>
		  </TR>
			<TR class="tdbg"> 
				<TD width="80" height="25" align="right"  class="clefttitle"><strong>�û��ȼ���</strong></TD>
			  <TD height="25"><select name="GroupID" id="GroupID" onchange="location.href='KS.User.asp?action=Add&amp;GroupID='+this.value;">
                <%=KS.GetUserGroup_Option(GroupID)%>
              </select></TD>
				<TD width="80" height="25" align="right" class="clefttitle"><strong>�û�״̬��</strong></TD>
			  <TD><input type="radio" name="locked" value="0" checked="checked" />
����&nbsp;&nbsp;
<input type="radio" name="locked" value="1" />
����</TD>
			</TR>
			<TR class="tdbg"> 
				<TD width="80" height="25" align="right" class="clefttitle"><strong>�û����ƣ�</strong></TD>
			  <TD> <Input class="textbox" Name="UserName" id="UserName" type=text size=20> <font color="red">*</font><input type="button" name="Submit22" value="����ʺ�"  onClick="checkaccount($('#UserName').val())" class="button"></TD>
				<TD width="80" height="25" align="right" class="clefttitle"><strong>�������䣺</strong></TD>
				<TD>  <Input class="textbox" Name="Email" type=text size=30 Value=""></TD>
			</TR>
				<TR class="tdbg"> 
					<TD width="80" height="25" align="right" class="clefttitle"><strong>�û����룺</strong></TD>
				  <TD height="25"><INPUT class="textbox" type="password" name="Password" value="" size="30" maxLength="12"><font color="red">*</font><br><font color="#FF6600">�û���¼ʱ������</font></TD>
					<TD width="80" height="25" align="right" class="clefttitle"><strong>�ظ����룺</strong></TD>
					<TD height="25"><INPUT class="textbox" type="password" name="PwdConfirm" value="" size="30" maxLength="12"><font color="red">*</font><br></TD>
				</TR>
			<TR  class="tdbg"> 
				<TD width="80" height="25" align="right" class="clefttitle"><strong>�������⣺</strong></TD>
				<TD><Select id=Question name=Question>
                    <Option value="" selected>--����ѡ��--</Option>
                    <Option value="�ҵĳ������֣�">�ҵĳ������֣�</Option>
                    <Option value="����õ�������˭��">����õ�������˭��</Option>
                    <Option value="����ϲ������ɫ��">����ϲ������ɫ��</Option>
                    <Option value="����ϲ���ĵ�Ӱ��">����ϲ���ĵ�Ӱ��</Option>
                    <Option value="����ϲ����Ӱ�ǣ�">����ϲ����Ӱ�ǣ�</Option>
                    <Option value="����ϲ���ĸ�����">����ϲ���ĸ�����</Option>
                    <Option value="����ϲ����ʳ�">����ϲ����ʳ�</Option>
                    <Option value="�����İ��ã�">�����İ��ã�</Option>
                    <Option value="����ѧУ��ȫ����ʲô��">����ѧУ��ȫ����ʲô��</Option>
                    <Option value="�ҵ��������ǣ�">�ҵ��������ǣ�</Option>
                    <Option value="����ϲ����С˵�����֣�">����ϲ����С˵�����֣�</Option>
                    <Option value="����ϲ���Ŀ�ͨ�������֣�">����ϲ���Ŀ�ͨ�������֣�</Option>
                    <Option value="��ĸ��/���׵����գ�">��ĸ��/���׵����գ�</Option>
                    <Option value="�������͵�һλ���˵����֣�">�������͵�һλ���˵����֣�</Option>
                    <Option value="����ϲ�����˶���ȫ�ƣ�">����ϲ�����˶���ȫ�ƣ�</Option>
                    <Option value="����ϲ����һ��Ӱ��̨�ʣ�">����ϲ����һ��Ӱ��̨�ʣ�</Option>
                  </Select>
			  <font color="#FF6600">*</font> </TD>
				<TD width="80" height="25" align="right" class="clefttitle"><strong>����𰸣�</strong></TD>
				<TD><INPUT class="textbox" type=text maxLength=20 size=30 name="Answer">			</TD>
			</TR>
			<TR class="tdbg">
				<TD width="80" height="25" align="right" class="clefttitle"><strong>�Ʒѷ�ʽ��</strong></TD>
				<TD Colspan=3><input name="ChargeType" type="radio" value="1" checked>
				�۵���<font color="#0066CC">���Ƽ���</font>
				<input type="radio" name="ChargeType" value="2">
			  ��Ч��(����Ч���ڣ��û����������Ķ��շ�����)
			  <input type="radio" name="ChargeType" value="3"/>
������</TD>
			</TR>
			

			<tr class="sort">
			  <td colspan="4" align="center">====�Զ���ѡ��====</td>
			</tr>
			<tr class="tdbg">
			  <td colspan="4">
			 <%
			            
					Dim Template:Template=LFCls.GetSingleFieldValue("Select Template From KS_UserForm Where ID=" & KS.U_G(GroupID,"formid"))
							
						   Dim FieldsList:FieldsList=LFCls.GetSingleFieldValue("Select FormField From KS_UserForm Where ID=" & KS.U_G(GroupID,"formid"))
						   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
						   RS.Open "Select FieldID,FieldType,FieldName,DefaultValue,Width,Height,Options,EditorType from KS_Field Where ChannelID=101 Order By OrderID",conn,1,1
						   Dim SQL,K,N,InputStr,O_Arr,O_Len,F_V,O_Value,O_Text,BRStr
						   If Not RS.EOf Then SQL=RS.GetRows(-1):RS.Close():Set RS=Nothing
						   For K=0 TO Ubound(SQL,2)
							 If KS.FoundInArr(FieldsList,SQL(0,k),",") Then
							  InputStr=""
							  If lcase(SQL(2,K))="province&city" Then
								 InputStr="<script language=""javascript"" src=""../plus/area.asp""></script>"
							  Else
							  Select Case SQL(1,K)
								Case 2:InputStr="<textarea style=""width:" & SQL(4,K) & "px;height:" & SQL(5,K) & "px;"" rows=""5"" class=""textbox"" name=""" & SQL(2,K) & """>" &SQL(3,K) & "</textarea>"
								Case 3
								  InputStr="<select style=""width:" & SQL(4,K) & """ name=""" & SQL(2,K) & """>"
								  O_Arr=Split(SQL(6,K),vbcrlf): O_Len=Ubound(O_Arr)
								  For N=0 To O_Len
									 F_V=Split(O_Arr(N),"|")
									 If Ubound(F_V)=1 Then
										O_Value=F_V(0):O_Text=F_V(1)
									 Else
										O_Value=F_V(0):O_Text=F_V(0)
									 End If						   
									 If Trim(SQL(3,K))=O_Value Then
										InputStr=InputStr & "<option value=""" & O_Value& """ selected>" & O_Text & "</option>"
									 Else
										InputStr=InputStr & "<option value=""" & O_Value& """>" & O_Text & "</option>"
									 End If
								  Next
									InputStr=InputStr & "</select>"
								Case 6
									 O_Arr=Split(SQL(6,K),vbcrlf): O_Len=Ubound(O_Arr)
									 If O_Len>1 And Len(SQL(6,K))>50 Then BrStr="<br>" Else BrStr=""
									 For N=0 To O_Len
									  IF O_Arr(N)<>"" Then
										F_V=Split(O_Arr(N),"|")
										If Ubound(F_V)=1 Then
										 O_Value=F_V(0):O_Text=F_V(1)
										Else
										 O_Value=F_V(0):O_Text=F_V(0)
										End If
										If Trim(SQL(3,K))=O_Value Then
											InputStr=InputStr & "<input type=""radio"" name=""" & SQL(2,K) & """ value=""" & O_Value& """ checked>" & O_Text & BRStr
										Else
											InputStr=InputStr & "<input type=""radio"" name=""" & SQL(2,K) & """ value=""" & O_Value& """>" & O_Text & BRStr
										 End If
									  End If
									 Next
							  Case 7
									O_Arr=Split(SQL(6,K),vbcrlf): O_Len=Ubound(O_Arr)
									 For N=0 To O_Len
										  F_V=Split(O_Arr(N),"|")
										  If Ubound(F_V)=1 Then
											O_Value=F_V(0):O_Text=F_V(1)
										  Else
											O_Value=F_V(0):O_Text=F_V(0)
										  End If						   
										  If KS.FoundInArr(Trim(SQL(3,K)),O_Value,",")=true Then
												 InputStr=InputStr & "<input type=""checkbox"" name=""" & SQL(2,K) & """ value=""" & O_Value& """ checked>" & O_Text
										 Else
										  InputStr=InputStr & "<input type=""checkbox"" name=""" & SQL(2,K) & """ value=""" & O_Value& """>" & O_Text
										 End If
								   Next
							  Case 10
							        on error resume next
									InputStr=InputStr & "<input type=""hidden"" id=""" & SQL(2,K) &""" name=""" & SQL(2,K) &""" value="""& Server.HTMLEncode(SQL(3,K)) &""" style=""display:none"" /><input type=""hidden"" id=""" & SQL(2,K) &"___Config"" value="""" style=""display:none"" /><iframe id=""" & SQL(2,K) &"___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=" & SQL(2,K) &"&amp;Toolbar=" & SQL(7,K) & """ width=""" &SQL(4,K) &""" height=""" & SQL(5,K) & """ frameborder=""0"" scrolling=""no""></iframe>"				
							  Case Else:InputStr="<input type=""text"" class=""textbox"" style=""width:" & SQL(4,K) & """ name=""" & SQL(2,K) & """ value=""" & SQL(3,K) & """>"
							  End Select
							  End If
							  'if SQL(1,K)=9 Then InputStr=InputStr & "<div><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_UpFile.asp?Type=Field&FieldID=" & SQL(0,K) & "&ChannelID=101' frameborder=0 scrolling=no width='100%' height='26'></iframe></div>"
				              Template=Replace(Template,"[@NoDisplay(" & SQL(2,K) & ")]","")
							  Template=Replace(Template,"[@" & SQL(2,K) & "]",InputStr)
							 End If
						   Next
							
							Response.Write Template
			  
			  %>			  </td>
			</tr>
			
			<TR class="tdbg"> 
			  <TD height="30" colspan="4" align="center"><input name="Action" type="hidden" id="Action" value="SaveAdd"> 
			  <input name=Submit class='button'  type=submit id="Submit" value="&nbsp;ȷ�����&nbsp;" > </TD>
			</TR>
		</form>
	    </TABLE>
		<%
		end sub
		
		Sub CheckUserName()
		  Dim UserName:UserName=KS.G("UserName")
		  If UserName="" Then
		   Response.Write "<script>alert('�������û�����!');window.close();</script>"
		  Else
		   If Conn.Execute("select userid from ks_user where username='" & UserName & "'").eof Then
		    Response.Write "<script>alert('��ϲ���û�" & UserName & "����ʹ�ã�');window.close();</script>"
		   Else
		    Response.Write "<script>alert('�Բ����û�" & UserName & "������ʹ�ã������䣡');window.close();</script>"
		   End If
		  End If
		 End Sub

		
		sub SaveAdd()
		    dim UserID,UserName,RealName,Password,PwdConfirm,Question,Answer,Sex,Email,HomePage,QQ,MSN,GroupID,locked,DataCount,ChargeType,Point,Money,BeginDate,Edays,province
			dim rsUser,sqlUser
			dim OfficeTel,Address,Sign
			dim IDCard,BirthDay,City,Zip,ICQ,UC
			Dim SchoolAge,UserWorking,HomeTel,Mobile
			Action=Trim(request("Action"))
			Password  = Trim(request("Password"))
			PwdConfirm= Trim(request("PwdConfirm"))
			Question  = Trim(request("Question"))
			Answer    = Trim(request("Answer"))
			Sex       = Trim(Request("Sex"))
			Email     = Trim(request("Email"))
			HomePage  = Trim(request("HomePage"))
			QQ        = Trim(request("QQ"))
			MSN       = Trim(request("MSN"))
			ICQ       =KS.G("ICQ")
			UC        =KS.G("UC")
			GroupID   = Trim(request("GroupID"))
			locked    = Trim(request("locked"))
			ChargeType=Trim(request("ChargeType"))
			Point=KS.ChkClng(Trim(request("Point")))
			Money=KS.ChkClng(Trim(request("Money")))
			BeginDate=Trim(request("BeginDate"))
			Edays=    KS.ChkClng(Trim(request("Edays")))
			province= KS.G("province")
			city=     KS.G("city")
			UserName=Trim(request("UserName"))
			RealName=Trim(request("RealName"))
			OfficeTel=Trim(request("OfficeTel"))
			Address=Trim(request("Address"))
			Sign=Trim(request("Sign"))
			IDCard=Trim(request("IDCard"))
			BirthDay=Trim(request("BirthDay"))
			Zip=Trim(Request("Zip"))
			HomeTel=Trim(request("HomeTel"))
			Mobile=Trim(request("Mobile"))

			if Password<>PwdConfirm then
				founderr=true
				errmsg=errmsg & "<br><li>�����ȷ�����벻һ��</li>"
			end if
			if Question="" then
				'founderr=true
				'errmsg=errmsg & "<br><li>������ʾ���ⲻ��Ϊ��</li>"
			end if
			if Email="" then
				founderr=true
				errmsg=errmsg & "<br><li>Email����Ϊ��</li>"
			else
				if KS.IsValidEmail(Email)=false then
					errmsg=errmsg & "<br><li>����Email�д���</li>"
					founderr=true
				end if
			end if
			 Dim EmailMultiRegTF:EmailMultiRegTF=KS.ChkClng(KS.Setting(28))
				If EmailMultiRegTF=0 Then
					Dim EmailRSCheck:Set EmailRSCheck = Conn.Execute("select UserID from KS_User where Email='" & Email & "'")
					If Not (EmailRSCheck.BOF And EmailRSCheck.EOF) Then
						EmailRSCheck.Close:Set EmailRSCheck = Nothing
						Response.Write("<script>alert('��ע���Email�Ѿ����ڣ������Email�����ԣ�');history.back();</script>")
						Exit Sub
					End If
					EmailRSCheck.Close:Set EmailRSCheck = Nothing
			 End If
		
			
			if GroupID="" then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>��ָ���û�����</li>"
			else
				GroupID=CLng(GroupID)
			end if
			if locked<>0 then locked=1
			if ChargeType="" then
				ChargeType=1
			else
				ChargeType=Clng(ChargeType)
			end if
			if BeginDate="" then
				BeginDate=Now()
			else
				BeginDate=Cdate(BeginDate)
			end if
			
			if BirthDay<>"" then
			    BirthDay=Split(BirthDay," ")(0)
				if Not IsDate(BirthDay) then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>�������ڴ���</li>"
				end if
			Else
			   Birthday=now
			end if
			if IDCard<>"" then
				if len(Cstr(IDCard))<15 then
					FoundErr=True
					ErrMsg=ErrMsg & "<br><li>���֤�������</li>"
				end if
			end if
			Set rsUser=Server.CreateObject("Adodb.RecordSet")
			sqlUser="select * from KS_User where username='" & username & "'"
			rsUser.Open sqlUser,Conn,1,3
			if not rsUser.Eof Then
			  rsUser.Close:Set rsUser=nothing
					FoundErr=True
					ErrMsg=ErrMsg & "<br><li>�û����Ѵ��ڣ�</li>"
			End If
			
		 Dim SQL,K
		 Dim FieldsList:FieldsList=LFCls.GetSingleFieldValue("Select FormField From KS_UserForm Where ID=" & KS.U_G(GroupID,"formid"))
		 If FieldsList="" Then FieldsList="0"
	     Set RS = Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select FieldName,MustFillTF,Title,FieldType From KS_Field Where ChannelID=101 and FieldID In(" & KS.FilterIDs(FieldsList) & ")",conn,1,1
		 If Not RS.Eof Then SQL=RS.GetRows(-1)
		 RS.Close
		  For K=0 To UBound(SQL,2)
		  	  If SQL(1,K)="1" Then 
			     if lcase(SQL(0,K))<>"province&city" and KS.S(SQL(0,K))="" then
					 FoundErr=True
					 ErrMsg=ErrMsg & "<br><li>" & SQL(2,K) & "������д!</li>"
				 elseif KS.S("province")="" or ks.s("city")="" then
					 FoundErr=True
					 ErrMsg=ErrMsg & "<br><li>" & SQL(2,K) & "����ѡ��!</li>"
				 end if
			   End If

			   
			   If SQL(3,K)="4" And Not Isnumeric(KS.S(SQL(0,K))) Then
			     FoundErr=True
				 ErrMsg=ErrMsg & "<br><li>" & SQL(2,K) & "������д����!</li>"
			   End If
			   If SQL(3,K)="5" And Not IsDate(KS.S(SQL(0,K))) Then 
			     FoundErr=True
				 ErrMsg=ErrMsg & "<br><li>" & SQL(2,K) & "������д��ȷ������!</li>"
			   End If
			   If SQL(3,K)="8" And Not KS.IsValidEmail(KS.S(SQL(0,K))) and SQL(1,K)="1" Then 
			     FoundErr=True
				 ErrMsg=ErrMsg & "<br><li>" & SQL(2,K) & "������д��ȷ��Email��ʽ!</li>"
			   End If 
			 Next


			if founderr=true then exit sub
            rsUser.AddNew			
			rsUser("UserName")=UserName
			rsUser("RealName")=RealName
			rsUser("Password")=MD5(KS.R(Password),16)
			rsUser("Question")=Question
			rsUser("Answer")=Answer
			rsUser("Email")=Email
			rsUser("HomePage")=HomePage
			rsUser("Sex")=Sex
			rsUser("GroupID")=GroupID
			rsUser("locked")=locked
			rsUser("ChargeType")=ChargeType
			rsUser("Point")=Point
			rsUser("Money")=Money
			rsUser("BeginDate")=BeginDate
			rsUser("Edays")=Edays
			rsUser("Sign")=Sign
			rsUser("Birthday")=Birthday
			rsUser("IDCard")=IDCard
			rsUser("province")=province
			rsUser("City")=City
			rsUser("Address")=Address
			rsUser("Zip")=Zip
			rsUser("MSN")=MSN
			rsUser("QQ")=QQ
			rsUser("ICQ")=ICQ
			rsUser("UC")=UC
			rsUser("HomeTel") = HomeTel
			rsUser("Mobile") = Mobile
			rsUser("OfficeTel")=OfficeTel
			rsUser("LastLoginIP")=KS.GetIP()
			rsUser("logintimes")=1
			rsUser("UserType")=KS.ChkClng(KS.U_G(GroupID,"usertype"))
			rsUser("lastlogintime")=now()
			rsUser("RegDate")=Now

				 '�Զ����ֶ�
				 For K=0 To UBound(SQL,2)
				  If left(Lcase(SQL(0,K)),3)="ks_" Then
				   rsUser(SQL(0,K))=KS.S(SQL(0,K))
				  End If
				 Next
			rsUser.update
			rsUser.Close
			set rsUser=Nothing
			Call KS.Alert("��ϲ�����û���ӳɹ���",ComeUrl)
		End Sub
		
		
		Sub Modify()
			dim rsUser,sqlUser,sSex,ChargeType,GroupID
			UserID=KS.ChkClng(UserID)
			GroupID=KS.ChkClng(KS.S("GroupID"))
			if UserID=0 then Response.Write("<script>alert('�������㣡');history.back();</script>")
			sqlUser="select * from KS_User where UserID=" & UserID
			Set rsUser=Server.CreateObject("ADODB.RECORDSET")
			rsUser.Open sqlUser,conn,1,1
			if rsUser.bof and rsUser.eof then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ�����û���</li>"
				rsUser.close:set rsUser=Nothing
				exit sub
			end if
			If GroupID=0 Then GroupID=rsUser("GroupID")
		ChargeType=rsUser("ChargeType")
		%>
		<SCRIPT language=javascript>
		function CheckFrom()
		{
		  if(document.myform.UserName.value=="")
			{
			  alert("�û�������Ϊ�գ�");
			  document.myform.UserName.focus();
			  return false;
			}
		
		  if(document.myform.Question.value=="")
			{
			  alert("�������ⲻ��Ϊ�գ�");
			  document.myform.Question.focus();
			  return false;
			}
		  if(document.myform.Answer.value=="")
			{
			  alert("����𰸲���Ϊ�գ�");
			  document.myform.Answer.focus();
			  return false;
			}
		  if(document.myform.Email.value=="")
			{
			  alert("�û�Email����Ϊ�գ�");
			  document.myform.Email.focus();
			  return false;
			}
		  if((document.myform.Password.value)!=(document.myform.PwdConfirm.value))
			{
			  alert("��ʼ������ȷ�����벻ͬ��");
			  document.myform.PwdConfirm.select();
			  document.myform.PwdConfirm.focus();	  
			  return false;
			}
		}
		</script>
		<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1">
		<FORM name="myform" action="KS.User.asp" method="post">
			<TR class="sort">
			  <TD height="22" colspan="4" align="center"><span style="margin-top:10px;text-align:center;font-weight:bold">�޸�ע���û���Ϣ</span></TD>
		  </TR>
			<TR class="tdbg"> 
				<TD width="80" height="25" align="right" class="clefttitle"><strong>�û����ƣ�</strong></TD>
			  <TD> <Input class="textbox" Name="UserName" type=text size=30 Value="<%=rsUser("UserName")%>" readonly> <font color="red">*</font></TD>
				<TD width="80" height="25" align="right" class="clefttitle"><strong>�û����䣺</strong></TD>
				<TD><input class="textbox" name="Email" type="text" size="30" value="<%=rsUser("Email")%>" /></TD>
			</TR>
				<TR class="tdbg"> 
					<TD width="80" height="25" align="right" class="clefttitle"><strong>�û����룺</strong></TD>
				  <TD height="25"><INPUT class="textbox" type="password" name="Password" value="" size="30" maxLength="12"><br><font color="#FF6600">��������޸ģ�������</font></TD>
					<TD width="80" height="25" align="right" class="clefttitle"><strong>�ظ����룺</strong></TD>
					<TD height="25"><INPUT class="textbox" type="password" name="PwdConfirm" value="" size="30" maxLength="12"><br><font color="#FF6600">��������޸ģ�������</font></TD>
				</TR>
			<TR  class="tdbg"> 
				<TD width="80" height="25" align="right" class="clefttitle"><strong>�������⣺</strong></TD>
				<TD><INPUT class="textbox" type=text maxLength=50 size=30 name="Question" value="<%=rsUser("Question")%>"> 
			  <font color="#FF6600">*</font> </TD>
				<TD width="80" height="25" align="right" class="clefttitle"><strong>����𰸣�</strong></TD>
				<TD><INPUT class="textbox" type=text maxLength=20 size=30 name="Answer" value="<%=rsUser("Answer")%>"></TD>
			</TR>
			<TR class="tdbg"> 
				<TD width="80" height="25" align="right"  class="clefttitle"><strong>�û��ȼ���</strong></TD>
			  <TD height="25"> 
			  <select name="GroupID" id="GroupID" onchange="location.href='?Action=Modify&Groupid='+this.value+'&UserID=<%=UserID%>';">
				<%=KS.GetUserGroup_Option(GroupID)%>
			  </select></TD>
			  <TD height="25" align="right"><strong>�û�״̬��</strong></TD>
			  <TD height="25"><input type="radio" name="locked" value="0" <%if rsUser("locked")=0 then Response.Write "checked"%> />
����&nbsp;&nbsp;
<input type="radio" name="locked" value="1" <%if rsUser("locked")=1 then Response.Write "checked"%> />
����<input type="radio" name="locked" value="2" <%if rsUser("locked")=2 then Response.Write "checked"%> />
�����<input type="radio" name="locked" value="3" <%if rsUser("locked")=3 then Response.Write "checked"%> />
������</TD>
			</TR>
			
			<TR class="tdbg">
				<TD width="80" height="25" align="right" class="clefttitle"><strong>�Ʒѷ�ʽ��</strong></TD>
				<TD><input name="ChargeType" type="radio" value="1" <%if ChargeType=1 then Response.Write " checked"%>>
				�۵���<font color="#0066CC">���Ƽ���</font>
				<input type="radio" name="ChargeType" value="2" <%if ChargeType=2 then Response.Write " checked"%>>
			  ��Ч��
			  <input type="radio" name="ChargeType" value="3" <%if ChargeType=3 then Response.Write " checked"%> />
������</TD>
			    <TD align="right"><strong>ͷ���ַ��</strong></TD>
			    <TD><input class="textbox" type="text" maxlength="20" size="30" name="UserFace" value="<%=rsUser("UserFace")%>" /></TD>
			</TR>
			<TR  class="tdbg">
				<TD width="80" height="25" align="right" class="clefttitle"><strong>��Ч���ޣ�</strong></TD>
				<TD height="50" Colspan=3>��ʼ���ڣ�
				<input  class="textbox" name="BeginDate" type="text" id="BeginDate" readonly value="<%=FormatDateTime(rsUser("BeginDate"),2)%>" size="20" maxlength="20"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�� Ч �ڣ�
			  <input  class="textbox" name="EDays" readonly type="text" id="EDays" value="<%=rsUser("EDays")%>" size="10" maxlength="10">
			 ��
			  <br>
				�����������ޣ����û������Ķ��շ����ݴ˹���ֻ�е��Ʒѷ�ʽΪ����Ч���ޡ�ʱ����Ч			  </TD>
			</TR>
			<TR class="tdbg"> 
				<TD width="80" height="25" align="right" class="clefttitle"><strong>�û�������</strong></TD>
				<TD> <Input class="textbox" readonly Name="Point" type=text size=30 Value="<%=rsUser("Point")%>"></TD>
				<TD width="80" height="25" align="right" class="clefttitle"><strong>�û��ʽ�</strong></TD>
			  <TD> <Input class="textbox" Name="Money" type=text size=8 Value="<%=rsUser("Money")%>">Ԫ</TD>
			</TR>
						
			<tr class="sort">
			  <td colspan="4" align="center">====�Զ���ѡ��====</td>
			</tr>
			<tr class="tdbg">
			  <td colspan="4">
			  <%
			            
					Dim Template:Template=LFCls.GetSingleFieldValue("Select Template From KS_UserForm Where ID=" & KS.U_G(GroupID,"formid"))
							
						   Dim FieldsList:FieldsList=LFCls.GetSingleFieldValue("Select FormField From KS_UserForm Where ID=" & KS.U_G(GroupID,"formid"))
						   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
						   RS.Open "Select FieldID,FieldType,FieldName,DefaultValue,Width,Height,Options,EditorType from KS_Field Where ChannelID=101 Order By OrderID",conn,1,1
						   Dim SQL,K,N,InputStr,O_Arr,O_Len,F_V,O_Value,O_Text,BRStr
						   If Not RS.EOf Then SQL=RS.GetRows(-1):RS.Close():Set RS=Nothing
						   For K=0 TO Ubound(SQL,2)
							 If KS.FoundInArr(FieldsList,SQL(0,k),",") Then
							  InputStr=""
							  If lcase(SQL(2,K))="province&city" Then
								 InputStr="<script src='../plus/area.asp'></script><script language=""javascript"">" &vbcrlf
								 If rsUser("Province")<>"" And Not ISNull(rsUser("Province")) Then
						         InputStr=InputStr & "$('#Province').val('" & rsUser("province") &"');" &vbcrlf
								 End If
						         If rsUser("City")<>"" And Not ISNull(rsUser("City")) Then
								  InputStr=InputStr & "$('#City')[0].options[1]=new Option('" & rsUser("City") & "','" & rsUser("City") & "');" &Vbcrlf
								  InputStr=InputStr & "$('#City')[0].options(1).selected=true;" & vbcrlf
						         end if
						          InputStr=InputStr & "</script>" &vbcrlf
							  Else
							  Select Case SQL(1,K)
								Case 2:InputStr="<textarea style=""width:" & SQL(4,K) & "px;height:" & SQL(5,K) & "px;"" rows=""5"" class=""textbox"" name=""" & SQL(2,K) & """>" &rsUser(SQL(2,K)) & "</textarea>"
								Case 3
								  InputStr="<select style=""width:" & SQL(4,K) & """ name=""" & SQL(2,K) & """>"
								  O_Arr=Split(SQL(6,K),vbcrlf): O_Len=Ubound(O_Arr)
								  For N=0 To O_Len
									 F_V=Split(O_Arr(N),"|")
									 If Ubound(F_V)=1 Then
										O_Value=F_V(0):O_Text=F_V(1)
									 Else
										O_Value=F_V(0):O_Text=F_V(0)
									 End If						   
									 If Trim(rsUser(SQL(2,K)))=O_Value Then
										InputStr=InputStr & "<option value=""" & O_Value& """ selected>" & O_Text & "</option>"
									 Else
										InputStr=InputStr & "<option value=""" & O_Value& """>" & O_Text & "</option>"
									 End If
								  Next
									InputStr=InputStr & "</select>"
								Case 6
									 O_Arr=Split(SQL(6,K),vbcrlf): O_Len=Ubound(O_Arr)
									 If O_Len>1 And Len(SQL(6,K))>50 Then BrStr="<br>" Else BrStr=""
									 For N=0 To O_Len
										F_V=Split(O_Arr(N),"|")
										If Ubound(F_V)=1 Then
										 O_Value=F_V(0):O_Text=F_V(1)
										Else
										 O_Value=F_V(0):O_Text=F_V(0)
										End If
										If Trim(rsUser(SQL(2,K)))=O_Value Then
											InputStr=InputStr & "<input type=""radio"" name=""" & SQL(2,K) & """ value=""" & O_Value& """ checked>" & O_Text & BRStr
										Else
											InputStr=InputStr & "<input type=""radio"" name=""" & SQL(2,K) & """ value=""" & O_Value& """>" & O_Text & BRStr
										 End If
									 Next
							  Case 7
									O_Arr=Split(SQL(6,K),vbcrlf): O_Len=Ubound(O_Arr)
									 For N=0 To O_Len
										  F_V=Split(O_Arr(N),"|")
										  If Ubound(F_V)=1 Then
											O_Value=F_V(0):O_Text=F_V(1)
										  Else
											O_Value=F_V(0):O_Text=F_V(0)
										  End If						   
										  If KS.FoundInArr(Trim(rsUser(SQL(2,K))),O_Value,",")=true Then
												 InputStr=InputStr & "<input type=""checkbox"" name=""" & SQL(2,K) & """ value=""" & O_Value& """ checked>" & O_Text
										 Else
										  InputStr=InputStr & "<input type=""checkbox"" name=""" & SQL(2,K) & """ value=""" & O_Value& """>" & O_Text
										 End If
								   Next
							  Case 10
							        Dim H_Value:H_Value=rsUser(SQL(2,K))
									If IsNull(H_Value) Then H_Value=" "
									InputStr=InputStr & "<textarea id=""" & SQL(2,K) &""" name=""" & SQL(2,K) &""" style=""display:none"">"& KS.HTMLCode(H_Value) &"</textarea><input type=""hidden"" id=""" & SQL(2,K) &"___Config"" value="""" style=""display:none"" /><iframe id=""" & SQL(2,K) &"___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=" & SQL(2,K) &"&amp;Toolbar=" & SQL(7,K) & """ width=""" &SQL(4,K) &""" height=""" & SQL(5,K) & """ frameborder=""0"" scrolling=""no""></iframe>"				
							  Case Else
								  InputStr="<input type=""text"" class=""textbox"" style=""width:" & SQL(4,K) & """ name=""" & SQL(2,K) & """ value=""" & rsUser(SQL(2,K)) & """>"

							  End Select
							  End If
							  'if SQL(1,K)=9 Then InputStr=InputStr & "<div><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_UpFile.asp?Type=Field&FieldID=" & SQL(0,K) & "&ChannelID=101' frameborder=0 scrolling=no width='100%' height='26'></iframe></div>"
				              Template=Replace(Template,"[@NoDisplay(" & SQL(2,K) & ")]","")
							  Template=Replace(Template,"[@" & SQL(2,K) & "]",InputStr)
							  
							 End If
						   Next
							
							Response.Write Template
			  
			  %>			  </td>
			</tr>
			

			<TR class="tdbg"> 
			  <TD height="30" colspan="4" align="center"><input name="Action" type="hidden" id="Action" value="SaveModify"> 
			  <input name=Submit class='button'  type=submit id="Submit" value="&nbsp;�����޸Ľ��&nbsp;" > &nbsp;&nbsp;<input type='button' onclick="location.href='KS.User.asp?userid=<%=rsUser("UserID")%>&action=ShowDetail';" value="�鿴��ӡ" class='button'><input name="UserID" type="hidden" id="UserID" value="<%=rsUser("UserID")%>"></TD>
			</TR>
		</form>
	    </TABLE>
		<%
			rsUser.close:set rsUser=Nothing
		end sub
		
		Sub ShowDetail()
			dim rsUser,sqlUser,sSex,ChargeType,GroupID
			UserID=KS.ChkClng(UserID)
			GroupID=KS.ChkClng(KS.S("GroupID"))
			if UserID=0 then Response.Write("<script>alert('�������㣡');history.back();</script>")
			sqlUser="select * from KS_User where UserID=" & UserID
			Set rsUser=Server.CreateObject("ADODB.RECORDSET")
			rsUser.Open sqlUser,conn,1,1
			if rsUser.bof and rsUser.eof then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ�����û���</li>"
				rsUser.close:set rsUser=Nothing
				exit sub
			end if
			If GroupID=0 Then GroupID=rsUser("GroupID")
		ChargeType=rsUser("ChargeType")
		%>

		<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1">
			<TR class="sort">
			  <TD height="22" colspan="4" align="center"><span style="margin-top:10px;text-align:center;font-weight:bold">�û���ϸ����</span></TD>
		  </TR>
			<TR class="tdbg"> 
				<TD height="25" align="center" colspan="4">
				<strong>�û����ƣ�</strong><%=rsUser("UserName")%>&nbsp;&nbsp;<strong>ע��ʱ�䣺</strong><%=rsUser("RegDate")%>&nbsp;&nbsp;<strong>�Ƽ��ˣ�</strong><%=rsUser("AllianceUser")%>&nbsp;&nbsp;<strong>�����ʽ�</strong><%=rsUser("money")%>Ԫ&nbsp;&nbsp;<strong>���õ�ȯ��</strong><%=rsUser("point")%>��&nbsp;&nbsp;<strong>���֣�</strong><%=rsUser("score")%>��
				</TD>
			</TR>
	
			<tr class="tdbg">
			  <td colspan="4">
			  <%
			            
					Dim Template:Template=LFCls.GetSingleFieldValue("Select Template From KS_UserForm Where ID=" & KS.U_G(GroupID,"formid"))
						   Dim FieldsList:FieldsList=LFCls.GetSingleFieldValue("Select FormField From KS_UserForm Where ID=" & KS.U_G(GroupID,"formid"))
						   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
						   RS.Open "Select FieldID,FieldType,FieldName,DefaultValue,Width,Options from KS_Field Where ChannelID=101 Order By OrderID",conn,1,1
						   Dim SQL,K,N,InputStr,O_Arr,O_Len,F_V,O_Value,O_Text,BRStr
						   If Not RS.EOf Then SQL=RS.GetRows(-1):RS.Close():Set RS=Nothing
						   For K=0 TO Ubound(SQL,2)
							 If KS.FoundInArr(FieldsList,SQL(0,k),",") Then
							  InputStr=""
							  If lcase(SQL(2,K))="province&city" Then
								 InputStr=rsUser("Province") & "" &  rsUser("City") & ""
							  Else
							  InputStr=rsUser(SQL(2,K))
							  If InputStr="" OR IsNull(InputStr) Then InputStr=" "
							 End If
							  Template=Replace(Template,"[@NoDisplay(" & SQL(2,K) & ")]","")
							  Template=Replace(Template,"[@" & SQL(2,K) & "]",InputStr)

							End If
						   Next
							Template=Replace(Template,"{@NoDisplay}"," style='display:none'")
							Response.Write Template
			  
			  %>			  </td>
			</tr>
			

			<TR class="tdbg"> 
			  <TD height="30" colspan="4" align="center"> 
			  <input name=Submit class='button'  type="button" onclick="this.style.display='none';document.getElementById('modifybutton').style.display='none';document.getElementById('backbutton').style.display='none';document.getElementById('mt').style.display='none';window.print();" id="Submit" value=" �� ӡ " >&nbsp;&nbsp;<input type="button" class="button" name="modifybutton" id="modifybutton" value=" �� �� " onclick="location.href='KS.User.asp?action=Modify&userid=<%=rsUser("UserID")%>';">&nbsp;&nbsp;<input name='backbutton' id='backbutton' type='button' onclick='history.back();' value=' �� �� ' class='button'></TD>
			</TR>
	    </TABLE>
		<br><br>
		<%
			rsUser.close:set rsUser=Nothing
	  End Sub
		
		
		
		sub AddMoney()
			dim rsUser,sqlUser
			UserID=KS.ChkClng(UserID)
			if UserID=0 then Response.Write("<script>alert('�������㣡');history.back();</script>")
			Set rsUser=Conn.Execute("select * from KS_User where UserID=" & UserID)
			if rsUser.bof and rsUser.eof then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ�����û���</li>"
				rsUser.close:set rsUser=Nothing
				exit sub
			end if
			if rsUser("ChargeType")=3 Then
			  rsUser.Close:Set rsUser=Nothing
			  Call KS.Alert("�������û��������Ѳ���!",Request.ServerVariables("HTTP_REFERER"))
			  Exit Sub
			End if
		%>
		<table width="80%" style="margin-top:10px" border="0" align="center" cellpadding="3" cellspacing="1" class="ctable">
		<FORM name="myform" action="KS.User.asp?ComeUrl=<%=Server.URLEncode(Request.ServerVariables("HTTP_REFERER"))%>" method="post">
			<TR class="sort">
			  <TD height="28" colspan="2" align="center"><b>�� �� �� ��</b></TD>
		   </TR>
			<TR class='tdbg'> 
			  <TD width="25%" height="28" align="right" class="clefttitle"><b>�û�����</b></TD>
			  <TD width="75%"><%=rsUser("UserName")%></TD>
			</TR>
			<TR class='tdbg'> 
			  <TD width="25%" height="28" align="right" class="clefttitle"><strong>�û�����</strong></TD>
			  <TD width="75%"><%=KS.GetUserGroupName(rsUser("GroupID"))%></TD>
			</TR>
			<TR  class='tdbg'>
			  <TD height="28" align="right" class="clefttitle"><strong>�Ʒѷ�ʽ��</strong></TD>
			  <TD><%
			  if rsUser("ChargeType")=1 then
				Response.Write "�۵���"
			  else
				Response.Write "��Ч��"
			  end if
			  %>
				<input name="ChargeType" type="hidden" id="ChargeType" value="<%=rsUser("ChargeType")%>">			  </TD>
			</TR>
			<TR class='tdbg'> 
			  <TD width="25%" height="28" align="right" class="clefttitle"><strong>�����ʽ�</strong></TD>
			  <TD width="75%"><%=rsUser("Money")%>Ԫ�����</TD>
			</TR>
			<%if rsUser("ChargeType")=1 then%>
			<TR  class='tdbg'>
			  <TD height="28" align="right" class="clefttitle"><strong>Ŀǰ���û�������</strong></TD>
			  <TD><%=rsUser("Point")%> ��</TD>
			</TR>
			<TR  class='tdbg'>
			  <TD height="28" align="right" class="clefttitle"><strong>׷�ӵ�����</strong></TD>
			  <TD> <input name="Point" type="text" id="Point" value="100" size="10" maxlength="10">
			  ��</TD>
			</TR>
			<%else%>
			<TR class='tdbg'>
			  <TD height="28" align="right" class="clefttitle"><strong>Ŀǰ����Ч������Ϣ��</strong></TD>
			  <TD>
			  <%
			  Response.Write "��ʼ��������" & FormatDateTime(rsUser("BeginDate"),2) & "&nbsp;&nbsp;&nbsp;&nbsp;�� Ч �ڣ�" & rsUser("Edays")
			 
				Response.Write "��"
			 
			  Response.Write "<br>"
			  tmpDays=rsUser("Edays")-DateDiff("D",rsUser("BeginDate"),now())
			  if tmpDays>=0 then
				Response.Write "���� <font color=blue>" & tmpDays & "</font> �쵽��"
			  else
				Response.Write "�Ѿ����� <font color=#ff6600>" & abs(tmpDays) & "</font> ��"
			  end if
			  %>			  </TD>
			</TR>
			<tr class='tdbg'>
			  <td height="60" align="right" class="clefttitle"><strong>׷��������</strong><br></td>
			  <td>
			  <input name="Edays" type="text" id="Edays" value="100" size="10" maxlength="10">
			  ��<br />
			  ��Ŀǰ�û���δ���ڣ���׷����Ӧ����<br />
��Ŀǰ�û��Ѿ�������Ч�ڣ�����Ч�ڴ�����֮�������¼�����</td>
			</tr>
			<%end if%>
			<tr class='tdbg'>
			  <td height="30" align="right" class="clefttitle"><strong>ͬʱ��ȥ��</strong><br></td>
			  <td>
			  <input name="Money" type="text" id="Money" value="100" size="10" maxlength="10"> Ԫ�����
			  <font color=red>
			  <%if rsUser("ChargeType")=1 then %>
			   �ʽ����ȯ��Ĭ�ϱ��ʣ�<%=KS.Setting(43)%>:1
			  <%else%>
			  �ʽ�����Ч�ڵ�Ĭ�ϱ��ʣ�<%=KS.Setting(44)%>:1
			  <%end if%>
			  </font> ����۳��ʽ�������0
			  </td>
			</tr>
			<TR class='tdbg'>
			  <TD height="28" align="right" class="clefttitle"><strong>������ԭ��</strong></TD>
			  <TD> <input name="Reason" type="text" id="Reason" value="<%If rsUser("ChargeType")=1 Then Response.Write "����ȯ����" Else Response.Write "����Ч��������"%>" size="55"></TD>
			</TR>
			<TR class='tdbg'> 
			  <TD height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveAddMoney"> 
			  <input name=Submit   type=submit id="Submit" value="&nbsp;�������ѽ��&nbsp;" > <input name="UserID" type="hidden" id="UserID" value="<%=rsUser("UserID")%>"></TD>
			</TR>
		</form>
	    </TABLE>
		<%
			rsUser.close : set rsUser=Nothing
		end sub
		
		
		
		sub SaveModify()
			dim UserID,UserName,RealName,Password,PwdConfirm,Question,Answer,Sex,Email,HomePage,QQ,MSN,GroupID,locked,DataCount,ChargeType,Point,Money,BeginDate,Edays,province,fax,UserFace
			dim rsUser,sqlUser
			dim OfficeTel,Address,Sign
			dim IDCard,BirthDay,City,Zip,ICQ,UC
			Dim SchoolAge,UserWorking,HomeTel,Mobile
			Action=Trim(request("Action"))
			UserID=KS.ChkClng(request("UserID"))
			if UserID=0 then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
				exit sub
			end if
			Password  = Trim(request("Password"))
			PwdConfirm= Trim(request("PwdConfirm"))
			Question  = Trim(request("Question"))
			Answer    = Trim(request("Answer"))
			Sex       = Trim(Request("Sex"))
			Email     = Trim(request("Email"))
			HomePage  = Trim(request("HomePage"))
			QQ        = Trim(request("QQ"))
			MSN       = Trim(request("MSN"))
			ICQ       =KS.G("ICQ")
			UC        =KS.G("UC")
			GroupID   = Trim(request("GroupID"))
			locked    = Trim(request("locked"))
			ChargeType=Trim(request("ChargeType"))
			Point=KS.ChkClng(Trim(request("Point")))
			Money=KS.ChkClng(Trim(request("Money")))
			BeginDate=Trim(request("BeginDate"))
			Edays=    KS.ChkClng(Trim(request("Edays")))
			province= KS.G("province")
			city=     KS.G("city")
			UserName=Trim(request("UserName"))
			RealName=Trim(request("RealName"))
			OfficeTel=Trim(request("OfficeTel"))
			Fax=Trim(Request("Fax"))
			Address=Trim(request("Address"))
			Sign=Trim(request("Sign"))
			IDCard=Trim(request("IDCard"))
			BirthDay=Trim(request("BirthDay"))
			Zip=Trim(Request("Zip"))
			HomeTel=Trim(request("HomeTel"))
			Mobile=Trim(request("Mobile"))
			UserFace=Trim(Request("UserFace"))
			Set rsUser=Server.CreateObject("Adodb.RecordSet")
			sqlUser="select * from KS_User where UserID=" & UserID
			rsUser.Open sqlUser,Conn,1,3
			if rsUser.bof and rsUser.eof then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ�����û���</li>"
				rsUser.close:set rsUser=Nothing
				exit sub
			end if
			if Password<>"" then
				if Instr(Password,"=")>0 or Instr(Password,"%")>0 or Instr(Password,chr(32))>0 or Instr(Password,"?")>0 or Instr(Password,"&")>0 or Instr(Password,";")>0 or Instr(Password,",")>0 or Instr(Password,"'")>0 or Instr(Password,",")>0 or Instr(Password,chr(34))>0 or Instr(Password,chr(9))>0 or Instr(Password,"��")>0 or Instr(Password,"$")>0 then
					errmsg=errmsg+"<br><li>�����к��зǷ��ַ�������㲻���޸����룬�뱣��Ϊ�ա�</li>"
					founderr=true
				end if
			end if
			if Password<>PwdConfirm then
				founderr=true
				errmsg=errmsg & "<br><li>�����ȷ�����벻һ��</li>"
			end if
			if Question="" then
				'founderr=true
				'errmsg=errmsg & "<br><li>������ʾ���ⲻ��Ϊ��</li>"
			end if
			
			if Email="" then
				founderr=true
				errmsg=errmsg & "<br><li>Email����Ϊ��</li>"
			else
				if KS.IsValidEmail(Email)=false then
					errmsg=errmsg & "<br><li>����Email�д���</li>"
					founderr=true
				end if
			end if
			
			 Dim EmailMultiRegTF:EmailMultiRegTF=KS.ChkClng(KS.Setting(28))
				If EmailMultiRegTF=0 Then
					Dim EmailRSCheck:Set EmailRSCheck = Conn.Execute("select UserID from KS_User where UserName<>'" & rsUser("UserName") & "' And Email='" & Email & "'")
					If Not (EmailRSCheck.BOF And EmailRSCheck.EOF) Then
						EmailRSCheck.Close:Set EmailRSCheck = Nothing
						Response.Write("<script>alert('��ע���Email�Ѿ����ڣ������Email�����ԣ�');history.back();</script>")
						Exit Sub
					End If
					EmailRSCheck.Close:Set EmailRSCheck = Nothing
			 End If
		
			
			if GroupID="" then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>��ָ���û�����</li>"
			else
				GroupID=CLng(GroupID)
			end if
			if locked<>0 then locked=1
			if ChargeType="" then
				ChargeType=1
			else
				ChargeType=Clng(ChargeType)
			end if
			if BeginDate="" then
				BeginDate=Now()
			else
				BeginDate=Cdate(BeginDate)
			end if
			
			if BirthDay<>"" then
			    BirthDay=Split(BirthDay," ")(0)
				if Not IsDate(BirthDay) then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>�������ڴ���</li>"
				end if
			end if

         Dim SQL,K
		 Dim FieldsList:FieldsList=LFCls.GetSingleFieldValue("Select FormField From KS_UserForm Where ID=" & KS.U_G(GroupID,"formid"))
		 If FieldsList="" Then FieldsList="0"
	     Set RS = Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select FieldName,MustFillTF,Title,FieldType From KS_Field Where ChannelID=101 and FieldID In(" & KS.FilterIDs(FieldsList) & ")",conn,1,1
		 If Not RS.Eof Then SQL=RS.GetRows(-1)
		 RS.Close
		  For K=0 To UBound(SQL,2)
			   If SQL(1,K)="1" Then 
			     if lcase(SQL(0,K))<>"province&city" and KS.S(SQL(0,K))="" then
					 FoundErr=True
					 ErrMsg=ErrMsg & "<br><li>" & SQL(2,K) & "������д!</li>"
				 elseif lcase(SQL(0,K))="province&city" and (KS.S("province")="" or ks.s("city")="") then
					 FoundErr=True
					 ErrMsg=ErrMsg & "<br><li>" & SQL(2,K) & "����ѡ��!</li>"
				 end if
			   End If
			   If SQL(3,K)="4" And Not Isnumeric(KS.S(SQL(0,K))) Then
			     FoundErr=True
				 ErrMsg=ErrMsg & "<br><li>" & SQL(2,K) & "������д����!</li>"
			   End If
			   If SQL(3,K)="5" And Not IsDate(KS.S(SQL(0,K))) Then 
			     FoundErr=True
				 ErrMsg=ErrMsg & "<br><li>" & SQL(2,K) & "������д��ȷ������!</li>"
			   End If
			   If SQL(3,K)="8" And Not KS.IsValidEmail(KS.S(SQL(0,K))) and SQL(1,K)="1" Then 
			     FoundErr=True
				 ErrMsg=ErrMsg & "<br><li>" & SQL(2,K) & "������д��ȷ��Email��ʽ!</li>"
			   End If 
			 Next
			
			if founderr=true then exit sub
			
			rsUser("RealName")=RealName
			if Password<>"" then rsUser("Password")=MD5(KS.R(Password),16)
			rsUser("Question")=Question
			if Answer<>"" then rsUser("Answer")=Answer
			rsUser("Email")=Email
			rsUser("HomePage")=HomePage
			rsUser("Sex")=Sex
			rsUser("GroupID")=GroupID
			rsUser("locked")=locked
			rsUser("ChargeType")=ChargeType
			rsUser("Point")=Point
			rsUser("Money")=Money
			rsUser("BeginDate")=BeginDate
			rsUser("Edays")=Edays
			rsUser("Sign")=Sign
			rsUser("UserFace")=UserFace
			if not isdate(birthday) then
			rsUser("Birthday")=now
			else
			rsUser("Birthday")=Birthday
			end if
			rsUser("IDCard")=IDCard
			rsUser("province")=province
			rsUser("City")=City
			rsUser("Address")=Address
			rsUser("Zip")=Zip
			rsUser("MSN")=MSN
			rsUser("Fax")=Fax
			rsUser("QQ")=QQ
			rsUser("ICQ")=ICQ
			rsUser("UC")=UC
			rsUser("HomeTel") = HomeTel
			rsUser("Mobile") = Mobile
			rsUser("OfficeTel")=OfficeTel
			'�Զ����ֶ�
			 For K=0 To UBound(SQL,2)
				If left(Lcase(SQL(0,K)),3)="ks_" Then
				   rsUser(SQL(0,K))=KS.G(SQL(0,K))
				End If
			 Next

			rsUser.update
			rsUser("UserType")=KS.ChkClng(KS.U_G(rsUser("GroupID"),"usertype"))
			rsUser.Update
			
			
			Dim FieldsXml:Set FieldsXml=LFCls.GetXMLFromFile("SpaceFields")
			If IsObject(FieldsXml) Then
				   	 Dim objNode,i,j,objAtr
					 Set objNode=FieldsXml.documentElement 
					If objNode.Attributes.item(0).Text<>"0" Then
					   If Not Conn.Execute("Select UserName From KS_EnterPrise Where UserName='" &rsUser("UserName") & "'").Eof Then
						 For i=0 to objNode.ChildNodes.length-1 
								set objAtr=objNode.ChildNodes.item(i) 
								on error resume next
								Conn.Execute("UPDATE KS_EnterPrise Set " & objAtr.Attributes.item(0).Text & "='" & rsUser(objAtr.Attributes.item(1).Text) & "' Where UserName='" & rsUser("UserName") & "'")
						 Next
					   End If
					End If
			 End If
				
				'-----------------------------------------------------------------
				'ϵͳ����
				'-----------------------------------------------------------------
				Dim API_KS,API_SaveCookie,SysKey
				If API_Enable Then
					Set API_KS = New API_Conformity
					API_KS.NodeValue "action","update",0,False
					API_KS.NodeValue "username",rsUser("UserName"),1,False
					Md5OLD = 1
					SysKey = Md5(API_KS.XmlNode("username") & API_ConformKey,16)
					Md5OLD = 0
					API_KS.NodeValue "syskey",SysKey,0,False
					API_KS.NodeValue "email",rsUser("Email"),1,False
					API_KS.NodeValue "mobile",rsUser("Mobile"),1,False
					API_KS.NodeValue "homepage",rsUser("homepage"),1,False
					API_KS.NodeValue "address",rsUser("Address"),1,False
					API_KS.NodeValue "zipcode",rsUser("zip"),1,False
					API_KS.NodeValue "qq",rsUser("qq"),1,False
					API_KS.NodeValue "icq",rsUser("icq"),1,False
					API_KS.NodeValue "msn",rsUser("msn"),1,False

					If KS.S("PassWord")<>"" Then
					API_KS.NodeValue "password",KS.R(KS.S("PassWord")),1,False
					End If
					API_KS.SendHttpData
					If API_KS.Status = "1" Then
						Response.Write "<script>alert('" &  API_KS.Message  & "');</script>"
						Exit Sub
					End If
					Set API_KS = Nothing
				End If
				'-----------------------------------------------------------------
			
			If KS.C_S(8,21)="1" Then
				  Conn.Execute("Update KS_GQ Set ContactMan='" & RealName &"',Tel='" &OfficeTel & "',Address='" & Address & "',Zip='" & Zip & "',Fax='" & Fax & "',Homepage='" & HomePage & "' where inputer='" & rsUser("UserName") & "'")
			End If
			rsUser.Close
			set rsUser=Nothing

			
			
			Call KS.Alert("��ϲ�����޸ĳɹ����밴ȷ�����أ�",ComeUrl)
		end sub
		
		sub SaveAddMoney()
			dim UserID,ChargeType,Point,Edays,rsUser,sqlUser,Reason
			Dim Money:Money=KS.G("Money")
			If Not IsNumeric(Money) Then 
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>��ȥ���ʽ�����</li>"
				exit sub
			End if
			Action=Trim(request("Action"))
			UserID=KS.ChkClng(request("UserID"))
			if UserID=0 then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
				exit sub
			end if
			ChargeType=Trim(request("ChargeType"))
			Point=KS.ChkClng(Trim(request("Point")))
			Edays=KS.ChkClng(Trim(request("Edays")))
			Reason=KS.G("Reason")
		
			if ChargeType="" then
				ChargeType=1
			else
				ChargeType=Clng(ChargeType)
			end if
			
			if ChargeType=1 and Point=0 then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>������Ҫ׷�ӵ��û�������</li>"
			end if
			if ChargeType=2 and Edays=0 then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>������Ҫ׷�ӵ�����</li>"
			end if
		    if Reason="" Then
			  FoundErr=True
			  ErrMsg=ErrMsg & " <br><li>���������ԭ��</li>"
			end if
			if founderr=true then exit sub
			Set rsUser=Server.CreateObject("Adodb.RecordSet")
			sqlUser="select * from KS_User where UserID=" & UserID
			rsUser.Open sqlUser,Conn,1,3
			if rsUser.bof and rsUser.eof then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ�����û���</li>"
				rsUser.close:set rsUser=Nothing
				exit sub
			end if
			If Round(rsUser("Money"))<Round(Money) Then
			  FoundErr=True
			  ErrMsg=ErrMsg & " <br><li>���û��Ŀ����ʽ��㣡</li>"
			  rsUser.close:set rsUser=Nothing
			  exit sub
			End If
			'rsUser("Money")=rsUser("Money")-Money
			if ChargeType=1 then
				'rsUser("Money")=rsUser("Money")-Money
			else
				ValidDays=rsUser("Edays")
				tmpDays=ValidDays-DateDiff("D",rsUser("BeginDate"),now())
				if tmpDays>0 then
					rsUser("Edays")=rsUser("Edays")+Edays
				else
					rsUser("BeginDate")=now
					rsUser("Edays")=Edays
				end if
			end if
			rsUser.update
			
			'���Ѽ�¼
			If Money>0 Then
			 if ChargeType=2 Then
			  Call KS.MoneyInOrOut(rsUser("UserName"),rsUser("RealName"),Money,4,2,now,0,KS.C("AdminName"),"���ڶһ���Ч����",0,0)
			 else
			  Call KS.MoneyInOrOut(rsUser("UserName"),rsUser("RealName"),Money,4,2,now,0,KS.C("AdminName"),"���ڶһ���ȯ",0,0)
			 end if
			end if
			
			
			if ChargeType=1 then
			 Call KS.PointInOrOut(0,0,rsUser("UserName"),1,Point,KS.C("AdminName"),Reason,0)
			else
			 Call KS.EdaysInOrOut(rsUser("UserName"),1,Edays,KS.C("AdminName"),Reason,0)
			end if
			rsUser.Close:set rsUser=Nothing
			IF Request("ComeUrl")<>"" Then
			Call KS.Alert("�����ɹ�!",Request("ComeUrl"))
			Else
			Call KS.Alert("�����ɹ�!","KS.User.asp")
			End IF
		end sub
		
		'��ӻ�Ա�ʽ�
		Sub AddZJ()
		dim rsUser,sqlUser
			UserID=KS.ChkClng(UserID)
			if UserID=0 then Response.Write("<script>alert('�������㣡');history.back();</script>")
			Set rsUser=Conn.Execute("select * from KS_User where UserID=" & UserID)
			if rsUser.bof and rsUser.eof then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ�����û���</li>"
				rsUser.close:set rsUser=Nothing
				exit sub
			end if
		%>
		<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1">
		<FORM name="myform" action="KS.User.asp?ComeUrl=<%=Server.URLEncode(Request.ServerVariables("HTTP_REFERER"))%>" method="post">
			<TR class="sort">
			  <TD colspan="2" align="center"><b>�� �� �� ��(�����ʽ�)</b></TD>
		   </TR>
			<TR class='tdbg'> 
			  <TD width="25%" height="28" align="right" class="clefttitle"><b>�û�����</b></TD>
			  <TD width="75%"><%=rsUser("UserName")%></TD>
			</TR>
			<TR class='tdbg'> 
			  <TD width="25%" height="28" align="right" class="clefttitle"><strong>�����ʽ�</strong></TD>
			  <TD width="75%"><%=rsUser("Money")%> Ԫ</TD>
			</TR>
			<TR class='tdbg'> 
			  <TD width="25%" height="28" align="right" class="clefttitle"><strong>�û�����</strong></TD>
			  <TD width="75%"><%=KS.GetUserGroupName(rsUser("GroupID"))%></TD>
			</TR>
			<TR  class='tdbg'>
			  <TD height="28" align="right" class="clefttitle"><strong>�ʽ���Դ��</strong></TD>
			  <TD><input name="MoneyType" type="radio" id="ChargeType" checked onclick="document.all.Remark.value='���л��';" value="2">���л��
			      <input name="MoneyType" type="radio" id="ChargeType" onclick="document.all.Remark.value='�ֽ���ȡ';" value="1">�������磺�ֽ�
		      </TD>
			</TR>
			<TR  class='tdbg'>
			  <TD height="28" align="right" class="clefttitle"><strong>������ڣ�</strong></TD>
			  <TD><input name="PayTime" type="text" id="PayTime" value="<%=formatdatetime(now,2)%>" size="15" class="textbox"></TD>
			</TR>
			<TR  class='tdbg'>
			  <TD height="28" align="right" class="clefttitle"><strong>���ѽ�</strong></TD>
			  <TD> <input name="Money" type="text" id="Money" value="100" size="15" class="textbox">
			  Ԫ</TD>
			</TR>
			
			
			<TR class='tdbg'>
			  <TD height="28" align="right" class="clefttitle"><strong>��ע��</strong></TD>
			  <TD> <input name="Remark" type="text" id="Remark" value="���л��" size="55"></TD>
			</TR>
			<TR class='tdbg'> 
			  <TD height="40" colspan="2" align="center"><input name="Action" type="hidden" id="Action" value="SaveAddZJ"> 
			  <input name=Submit  class='button' type=submit id="Submit" value="&nbsp;�������ѽ��&nbsp;" > <input name="UserID" type="hidden" id="UserID" value="<%=rsUser("UserID")%>"><input class='button' type='button' value=' ���� ' onclick='javascript:history.back();'></TD>
			</TR>
		</form>
	    </TABLE>
		<%
			rsUser.close : set rsUser=Nothing
		End Sub
		
		'��������
		sub SaveAddZJ()
			dim UserID,MoneyType,Money,PayTime,Remark,sqlUser,rsUser
			Action=Trim(request("Action"))
			UserID=KS.ChkClng(request("UserID"))
			if UserID=0 then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>�������㣡</li>"
				exit sub
			end if
			MoneyType=Trim(request("MoneyType"))
			Money=KS.G("Money")
			PayTime=KS.G("PayTime")
			Remark=KS.G("Remark")
            If Not IsDate(PayTime) Then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>������ڸ�ʽ����</li>"
			end if
			if KS.ChkClng(Money)=0 then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>������Ҫ���ѽ��</li>"
			end if
		    if Remark="" Then
			  FoundErr=True
			  ErrMsg=ErrMsg & " <br><li>�����뱸ע</li>"
			end if
			if founderr=true then exit sub
			Set rsUser=Server.CreateObject("Adodb.RecordSet")
			sqlUser="select * from KS_User where UserID=" & UserID
			rsUser.Open sqlUser,Conn,1,3
			if rsUser.bof and rsUser.eof then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>�Ҳ���ָ�����û���</li>"
				rsUser.close:set rsUser=Nothing
				exit sub
			end if
				'rsUser("Money")=rsUser("Money")+Money
			rsUser.update
			
			Call KS.MoneyInOrOut(rsUser("UserName"),rsUser("RealName"),Money,MoneyType,1,PayTime,"0",KS.C("AdminName"),Remark,0,0)
			IF Request("ComeUrl")<>"" Then
			Call KS.Alert("�����ɹ�!",Request("ComeUrl"))
			Else
			Call KS.Alert("�����ɹ�!","KS.User.asp")
			End IF
		end sub
		
		
		sub DelUser()
			if UserID="" then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>��ָ��Ҫɾ�����û�</li>"
				exit sub
			end if
			UserID=replace(UserID," ","")
			
		    Dim rsUser:Set rsUser=Conn.Execute("select username from KS_User Where UserID In(" & UserID & ") and GroupID<>1")					
			Do While Not rsUser.Eof
				Conn.Execute("Delete From KS_Blog Where UserName='" & rsUser(0) & "'")
				Conn.Execute("Delete From KS_BlogInfo Where UserName='" & rsUser(0) & "'")
				Conn.Execute("Delete From KS_BlogMusic Where UserName='" & rsUser(0) & "'")
				Conn.Execute("Delete From KS_Enterprise Where UserName='" & rsUser(0) & "'")
				Conn.Execute("Delete From KS_EnterpriseNews Where UserName='" & rsUser(0) & "'")
				Conn.Execute("Delete From KS_UserLog Where UserName='" & rsUser(0) & "'")
			
			
				'-----------------------------------------------------------------
				'ϵͳ����
				'-----------------------------------------------------------------
				Dim API_KS,API_SaveCookie,SysKey
				If API_Enable Then
					Set API_KS = New API_Conformity
					API_KS.NodeValue "action","delete",0,False
					API_KS.NodeValue "username",rsUser("UserName"),1,False
					Md5OLD = 1
					SysKey = Md5(API_KS.XmlNode("username") & API_ConformKey,16)
					Md5OLD = 0
					API_KS.NodeValue "syskey",SysKey,0,False
					API_KS.SendHttpData
					If API_KS.Status = "1" Then
						Response.Write "<script>alert('" &  API_KS.Message  & "');history.back();</script>"
						Exit Sub
					End If
					Set API_KS = Nothing
				End If
				'-----------------------------------------------------------------
				
		  rsUser.MoveNext
		  Loop
		  rsUser.Close
		  Set rsUser=Nothing
			
			Conn.Execute ("Delete From KS_UploadFiles Where channelid=1023 and infoid in(" & UserID &")")
			Conn.Execute ("Delete From KS_UploadFiles Where channelid=1024 and infoid in(" & UserID &")")
			if instr(UserID,",")>0 then
				sql="delete from KS_User where UserID in (" & UserID & ") and GroupID<>1"
			else
				UserID=KS.ChkClng(UserID)
				sql="delete from KS_User where UserID=" & UserID & "  and GroupID<>1"
			end if
			Conn.Execute sql
			Response.Redirect ComeUrl
		end sub
		
		sub locked()
			if UserID="" then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>��ѡ��Ҫ�������û�</li>"
				exit sub
			end if
			if instr(UserID,",")>0 then
				UserID=replace(UserID," ","")
				sql="Update KS_User set locked=1 where UserID in (" & UserID & ")"
			else
				UserID=KS.ChkClng(UserID)
				sql="Update KS_User set locked=1 where UserID=" & UserID
			end if
			Conn.Execute sql
			Response.Redirect ComeUrl
		end sub
		
		
		sub Unlocked()
			if UserID="" then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>��ָ��Ҫ�������û�</li>"
				exit sub
			end if
			if instr(UserID,",")>0 then
				UserID=replace(UserID," ","")
				sql="Update KS_User set locked=0 where UserID in (" & UserID & ")"
			else
				UserID=KS.ChkClng(UserID)
				sql="Update KS_User set locked=0 where UserID=" & UserID
			end if
			Conn.Execute sql
			Response.Redirect ComeUrl
		end sub
		
		sub verify(v)
			if UserID="" then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>��ѡ��Ҫ��˵��û�</li>"
				exit sub
			end if
			if instr(UserID,",")>0 then
				UserID=replace(UserID," ","")
				sql="Update KS_User set locked=" & v & " where UserID in (" & UserID & ")"
			else
				UserID=KS.ChkClng(UserID)
				sql="Update KS_User set locked=" & v & " where UserID=" & UserID
			end if
			Conn.Execute sql
			Response.Redirect ComeUrl
		end sub
		

		
		sub MoveUser()
			Dim RsGroup
			Dim sGroupName,sChargeType,sValidDays,sGroupPoint
			Dim GroupID
			if UserID="" then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>��ָ��Ҫ�ƶ����û�</li>"
				Exit Sub
			end if
			GroupID=KS.ChkClng(request("GroupID"))
			if GroupID=0 then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>��ָ��Ŀ���û���</li>"
				Exit Sub
			end if
			UserID=replace(UserID," ","")
			Set RsGroup=Conn.Execute("Select GroupName,ChargeType,ValidDays,GroupPoint From KS_UserGroup Where ID="&GroupID&"")
			if Not (RsGroup.Bof and RsGroup.Eof) then
				sGroupName	= RsGroup(0)
				sChargeType	= RsGroup(1)
				sValidDays	= RsGroup(2)
				sGroupPoint	= RsGroup(3)
			else
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>��ָ��Ŀ���û���</li>"
				Exit Sub
			end if
			RsGroup.Close : Set RsGroup=Nothing
			ErrMsg = "&nbsp;&nbsp;&nbsp;&nbsp;�Ѿ��ɹ���ѡ���û���Ϊ��<font color=blue>"&sGroupName&"</font>�������Ұ����û�����Ԥ��ֵ���趨����Щ�û��ļƷѷ�ʽ����ʼ��������Ч�ڵ����ݡ�"
			ErrMsg = ErrMsg & "<br><br>�Ʒѷ�ʽ��"
			if sChargeType=1 then
				ErrMsg = ErrMsg & "�۵���<br>��ʼ������" & sGroupPoint & "��"
			else
				ErrMsg = ErrMsg & "��Ч��<br>��ʼ���ڣ�" & Formatdatetime(now(),2) & "<br>�� Ч �ڣ�" & sValidDays & "��"
			end if
			Dim UserType:UserType=KS.U_G(GroupID,"usertype")
            if DataBaseType=1 then
			    Conn.Execute("Update KS_User set UserType=" & UserType & ",GroupID=" & GroupID & ",ChargeType =" & sChargeType & ",Point =" & sGroupPoint & ",BeginDate ='" & formatdatetime(now(),2) & "',EDays=" & sValidDays & " where UserID in (" & UserID & ")")
			else
				Conn.Execute("Update KS_User set  UserType=" & UserType & ",GroupID=" & GroupID & ",ChargeType =" & sChargeType & ",Point =" & sGroupPoint & ",BeginDate =#" & formatdatetime(now(),2) & "#,EDays=" & sValidDays & " where UserID in (" & UserID & ")")
			end if
			Response.Write KS.ShowError(ErrMsg)
		end sub

		
End Class
%> 
