<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
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
'response.buffer=false
Dim KSCls
Set KSCls = New Admin_UserMessage
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_UserMessage
        Private KS
		Private Action,RSObj,MaxPerPage,TotalPut,CurrentPage
		Private Title, Message, Numc

		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		   MaxPerPage = 20
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		    If Not KS.ReturnPowerResult(0, "KMUA10003") Then
			  Response.Write ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back();';</script>")
			  Call KS.ReturnErr(1, "")
			End If
	        Response.Write "<html>"
			Response.Write"<head>"
			Response.Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			Response.Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			Response.Write "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
			Response.Write"</head>"
			Response.Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			Response.Write"<table width=""100%""  height=""25"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
			Response.Write " <tr>"
			Response.Write"	<td height=""25"" class=""sort""> "
			Response.Write " &nbsp;&nbsp;<strong>�û����Ź���</strong><a href='?action=new'>���Ͷ���</a>"
			Response.Write	" </td>"
			Response.Write " </tr>"
			Response.Write"</TABLE>"
		Action=Trim(Request("Action"))
		Select Case Action
		Case "new","edit"
		    call SendMsg()
		Case "add"
			call savemsg()
		Case "saveedit"
		    Call editsavemsg()
		Case "delall"
			call delall()
		Case "delchk"
			call delchk()
		Case "del"
		    call delbyid()
		Case else
			call main()
		end Select
		Response.Write "<div style=""text-align:center;color:#003300"">-----------------------------------------------------------------------------------------------------------</div>"
		Response.Write "<div style=""height:30px;text-align:center"">KeSion CMS V 4.5 Sp1, Copyright (c) 2006-2010 <a href=""http://www.kesion.com/"" target=""_blank""><font color=#ff6600>KeSion.Com</font></a>. All Rights Reserved . </div>"%>
		</body>
		</html>
		<%
		End Sub
		
		Sub Main()
		If Request("page") <> "" Then
			  CurrentPage = CInt(Request("page"))
		Else
			  CurrentPage = 1
		End If
         %>
		<table width="100%" style="border-top:1px #CCCCCC solid" border="0" align="center" cellspacing="0" cellpadding="0">
		  	  <form name="myform" method="Post" action="KS.UserMessage.asp">
				  <tr class='sort'>
					<td height="22" width="30" align="center">ѡ��</td>
					<td align="center">����</td>
					<td width="80" align="center">������</td>
					<td align="center" width="80">������</td>
					<td width="100" align="center">����ʱ��</td>
					<td width="40" align="center">״̬</td>
					<td width="120" align="center">����</td>
				  </tr>
			<%
		           Set RSObj = Server.CreateObject("ADODB.RecordSet")
				   Dim Param:Param=" where 1=1"
				   If KS.S("KeyWord")<>"" Then
				     select case KS.ChkClng(KS.S("condition"))
					   case 1
					    Param=Param & " and title like '%" & KS.S("KeyWord") & "%'"
					   case 2
					    Param=Param & " and Sender like '%" & KS.S("KeyWord") & "%'"
					   case 3
					    Param=Param & " and Incept like '%" & KS.S("KeyWord") & "%'"
					 end select 
				   End If
				   RSObj.Open "SELECT * FROM KS_Message " & Param & " order by id Desc", Conn, 1, 1
				 If RSObj.EOF Then
				    Response.Write "<tr><td colspan=8 height='30' align='center'>�Ҳ����κζ���Ϣ��</td></tr>"
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
				 %>	
		<tr class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
			<td colspan=8 height="30">
			<input type="hidden" value="del" name="action">
			<input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">ѡ�б�ҳ��ʾ�����м�¼&nbsp;<input type="submit" value="ɾ��ѡ�еļ�¼" onclick="return(confirm('ȷ��ɾ��ѡ�еļ�¼��'))" class="button">
			&nbsp
			<input type="button" value="���Ͷ���" onclick="location.href='?action=new';" class="button">
					 </td>
		  </tr> 
		  <%
		  Response.Write "<tr><td colspan='7' align='right'>"
		  			 Call KS.ShowPageParamter(totalPut, MaxPerPage, "KS.UserMessage.asp", True, "��", CurrentPage, "KeyWord=" & KS.S("KeyWord") &"&condition=" & ks.s("condition"))

			Response.Write "</td></tr>"

		  %> 
		</form>
		</table>
		<div>
		<form action="KS.UserMessage.asp" name="myform" method="post">
		   <div style="border:1px dashed #cccccc;margin:3px;padding:4px">
			  &nbsp;<strong>��������=></strong>
			 &nbsp;�ؼ���:<input type="text" class='textbox' name="keyword">&nbsp;����:
			 <select name="condition">
			  <option value=1>���ű���</option>
			  <option value=2>�����û�</option>
			  <option value=3>�����û�</option>
			 </select>
			  &nbsp;<input type="submit" value="��ʼ����" class="button" name="s1">
			  </div>
		</form>
		</div>
		<%
		End Sub
		
		Sub ShowContent()
		 Dim i:i=1
		 Do While Not RSObj.Eof
		 %>
		  <tr height="23" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
		    <td><input name="ID" type="checkbox" onClick="unselectall()" id="ID" value="<%=RSObj("ID")%>"></td>
			<td><img src="images/Announce.gif" align="absmiddle"><a href="?action=edit&id=<%=rsobj("id")%>"><%=KS.Gottopic(rsobj("title"),35)%></a></td>
			<td><%=rsobj("sender")%></td>
			<td><%=rsobj("Incept")%></td>
			<td><%=rsobj("sendtime")%></td>
			<td align="center">
			<%if rsobj("flag")=0 then
			   response.write "<font color=red>δ��</font>"
			  else
			   response.write "<font color=blue>�Ѷ�</font>"
			  end if
			 %>
			</td>
			<td align="center"><a href="?action=edit&id=<%=rsobj("id")%>">�޸�</a> | <a onclick="return(confirm('ɾ���󲻿ɻָ���ȷ��ɾ����?'))" href="?action=del&id=<%=rsobj("id")%>">ɾ��</a></td>
		  </tr>
		   <tr><td colspan='8' background='images/line.gif'></td></tr>
		 <%if i>=maxperpage then exit do
		   i=I+1
		  RSObj.MoveNext
		 Loop
		End Sub
		
		Sub SendMsg()
		  dim flag,display,Incept,title,content,sendtime
		  If KS.S("Action")="edit" then
		    flag="saveedit"
			display=" style='display:none'"
			dim rs:set rs=server.createobject("adodb.recordset")
			rs.open "select * from ks_message where id="& ks.chkclng(ks.s("id")),conn,1,1
			if not rs.eof then
			 Incept=rs("Incept")
			 title=rs("title")
			 content=rs("content")
			 sendtime=rs("sendtime")
			end if
			rs.close:set rs=nothing
		  else
		    flag="add"
			
			 dim userid:userid=KS.FilterIds(replace(request("userid")," ",""))
			 dim usernamelist
			 if userid<>"" then
				 set rs=KS.InitialObject("adodb.recordset")
				 rs.open "select userid,username from ks_user where userid in("& userid & ")",conn,1,1
				 do while not rs.eof
				  if usernamelist="" then
				   usernamelist=rs(1)
				  else
				   usernamelist=usernamelist &"," & rs(1)
				  end if
				  rs.movenext
				 loop
				 rs.close:set rs=nothing
			 end if

			
			
			
		  end if
		%>
		<table width="100%" border="0" style="margin-top:3px" align=center cellpadding="3" cellspacing="1" class="ctable">
		  
		  <form action="KS.UserMessage.asp?action=<%=flag%>" method=post name="myform" id="myform">
		   <input type="hidden" value="<%=KS.S("id")%>" name="id">
			<tr class="sort">
			  <td height="25" colspan="2" align="center">���Ͷ���Ϣ</td>
		    </tr>
			<tr class="tdbg"<%=display%>>
				<td height="25" align="right" class="clefttitle">�û����</td>
				<td>
				<Input type="radio" name="UserType" value="1" checked onclick="UType(this.value)">�û�����
				<Input type="radio" name="UserType" value="2" onclick="UType(this.value)">�û���
				<Input type="radio" name="UserType" value="0" onclick="UType(this.value)">�����û�				</td>
			</tr>
			<%if ks.s("action")="edit" then%>
			<tr class="tdbg" id="ToUserName">
				<td height="25" align="right" class="clefttitle">�����û���</td>
				<td> 
				<%=Incept%>
				</td>
			</tr>
            <tr class="tdbg">
			   <td height="25" align="right" class="clefttitle">����ʱ�䣺</td>
			   <td><input type="text" name="sendtime" value="<%=sendtime%>"> <font color=red>��ʽ��0000-00-00 00:00</font></td>
			</tr>
			<%else%>
			<tr class="tdbg" id="ToUserName">
				<td height="25" align="right" class="clefttitle">�� �� ����</td>
				<td> <INPUT class="textbox" TYPE="text" value="<%=usernamelist%>" NAME="UserName" size="80"><br>
				�������û�����(����û�������Ӣ�Ķ��š�,���ָ�,ע�����ִ�Сд)</td>
			</tr>
			<%end if%>
			<tr class="tdbg" id="ToGroupID" style="display:none;">
				<td height="25" align="right" class="clefttitle">�� �� �飺</td>
				<td>
					<table width="100%" border="0" cellspacing="1" cellpadding="3" align=center>
					<tr><td>
					<%=KS.GetUserGroup_CheckBox("GroupID",0,5)
					%>
					</td></tr>
					<tr><td height=20><input type="button" value="�򿪸߼�����" NAME="OPENSET" onclick="openset(this,'UpSetting')"></td></tr>
					<tr><td height=20 ID="UpSetting" style="display:NONE">
						<table width="100%" border="0" cellspacing="1" cellpadding="3" align=center>
						<tr><td height=20 colspan="4">������������(������������ѡ����û�����Ч)</td></tr>
						<tr>
							<td width="15%">����½ʱ�䣺</td>
							<td width="35%">
							<input class="textbox" type="text" name="LoginTime" onkeyup="CheckNumber(this,'����')" size=6>�� &nbsp;<INPUT TYPE="radio" NAME="LoginTimeType" checked value="0">���� <INPUT TYPE="radio" NAME="LoginTimeType" value="1">����							</td>
							<td width="15%">ע��ʱ�䣺</td>
							<td width="35%">
							<input class="textbox" type="text" name="RegTime" onkeyup="CheckNumber(this,'����')" size=6>�� &nbsp;<INPUT TYPE="radio" NAME="RegTimeType" checked value="0">���� <INPUT TYPE="radio" NAME="RegTimeType" value="1">����							</td>
						</tr>
						<tr>
							<td>��½������</td>
							<td><input class="textbox" type="text" name="Logins" size=6 onkeyup="CheckNumber(this,'����')">�� &nbsp;<INPUT TYPE="radio" NAME="LoginsType" checked value="0">���� <INPUT TYPE="radio" NAME="LoginsType" value="1">����							</td>
							<td>�������£�</td>
							<td><input class="textbox" type="text" name="UserArticle" size=6 onkeyup="CheckNumber(this,'ƪ��')">ƪ &nbsp;<INPUT TYPE="radio" NAME="UserArticleType" checked value="0">���� <INPUT TYPE="radio" NAME="UserArticleType" value="1">����</td>
						</tr></table>
					</td></tr></table>				</td>
			</tr>
			<tr class=tdbg> 
			  <td width="20%" height="25" align="right" class="clefttitle">��Ϣ���⣺</td>
			  <td width="80%"> 
				<input class="textbox" type="text"  value="<%=title%>" name="title" size="80">			  </td>
			</tr>
			<tr class=tdbg> 
			  <td width="20%" height="25" align="right" class="clefttitle">��Ϣ���ݣ�</td>
			  <td width="80%"> 
				 <textarea id="message" name="message"  style="display:none"><%=server.htmlencode(content)%></textarea>
				<iframe id="content___Frame" src="../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=message&amp;Toolbar=NewsTool" width="695" height="200" frameborder="0" scrolling="no"></iframe>		  
				</td>
			</tr>
			<tr class=tdbg> 
			  <td height="25" colspan="2" align="center"> 
				  <input type="submit" name="Submit" value="������Ϣ" class='button' onclick="return(checkform())">
				  <input type="reset" name="Submit2" value="������д" class='button'>			  </td>
		    </tr>
		  </form>
		</table>
		<script>
		 function checkform()
		 {
		   if (document.myform.title.value==''){
			 alert('վ�ڶ��ű��ⲻ��Ϊ�գ�');
			 document.myform.title.focus();
			 return false;
		  }
		  if ((FCKeditorAPI.GetInstance('message').GetXHTML(true)==""))
			{
			  alert("վ�ڶ������ݲ���Ϊ�գ�");
			  FCKeditorAPI.GetInstance('message').Focus();
			  return false;
		   } 
		  
       return true
		 }
		</script>
		<br>
		<table width="100%" border="0" align=center cellpadding="3" cellspacing="1" class="ctable">
		  <tr align="center" class="sort"> 
			<td height="25" colspan="2">����Ϣ����(����ɾ��)</td>
		  </tr>
		  <form action="KS.UserMessage.asp?action=del" method=post>
		  </form>
		  <form action="KS.UserMessage.asp?action=delall" method=post>
			<tr> 
			  <td colspan="2" bgcolor="#FFFFFF" class=tdbg> ����ɾ���û�ָ�������ڶ���Ϣ��Ĭ��Ϊɾ���Ѷ���Ϣ����<br>
				<select name="delDate" size=1>
				  <option value=7>һ������ǰ</option>
				  <option value=30>һ����ǰ</option>
				  <option value=60>������ǰ</option>
				  <option value=180>����ǰ</option>
				  <option value="all">������Ϣ</option>
				</select>
				&nbsp; 
				<input type="checkbox" name="isread" value="yes">
				����δ����Ϣ 
				<input type="submit" name="Submit" class="button" value="�� ��">
			  </td>
			</tr>
		  </form>
		  <form action="KS.UserMessage.asp?action=delchk" method=post>
			<tr> 
			  <td colspan="2" bgcolor="#FFFFFF" class=tdbg> ����ɾ������ĳ�ؼ��ֶ��ţ�ע�⣺��������ɾ�������Ѷ���δ����Ϣ����<br>
				�ؼ��֣� 
				<input class="textbox" type="text" name="keyword" size=30>
				&nbsp;�� 
				<select name="selaction" size=1>
				  <option value=1>������</option>
				  <option value=2>������</option>
				</select>
				&nbsp; 
				<input type="submit" name="Submit" value="�� ��" class='button'>
			  </td>
			</tr>
		  </form>
		</table>
		<SCRIPT LANGUAGE="JavaScript">
		<!--
		function openset(v,s){
			if (v.value=='�򿪸߼�����'){
				document.getElementById(s).style.display = "";
				v.value="�رո߼�����";
			}
			else{
				v.value="�򿪸߼�����";
				document.getElementById(s).style.display = "none";
			}
		}
		function UType(n){
			if (n==1){
				document.getElementById("ToUserName").style.display = "";
				document.getElementById("ToGroupID").style.display = "none";
			}
			else if(n==2){
				document.getElementById("ToUserName").style.display = "none";
				document.getElementById("ToGroupID").style.display = "";
			}
			else{
				document.getElementById("ToUserName").style.display = "none";
				document.getElementById("ToGroupID").style.display = "none";
			}
		}
		//-->
		</SCRIPT>
		<%
		end sub
		
		sub editsavemsg()
		   dim id:id=ks.chkclng(ks.s("id"))
		   dim title:title=ks.s("title")
		   dim content:content=ks.s("message")
		   dim sendtime:sendtime=ks.s("sendtime")
		   if not isdate(sendtime) then
		    Response.Write "<script>alert('ʱ���ʽ����ȷ!');history.back();</script>"
			Exit Sub
			end if
			dim rs:set rs=server.createobject("adodb.recordset")
			rs.open "select * from ks_message where id=" &id,conn,1,3
			if not rs.eof then
			  rs("title")=title
			  rs("content")=content
			  rs("sendtime")=sendtime
			  rs.update
			end if
			rs.close
			set rs=nothing
			response.write "<script>alert('��ϲ���޸ĳɹ�!');location.href='ks.usermessage.asp';</script>"
		   response.end
		end sub
		
		Sub delbyid()
		  If Ks.G("id")="" Then
				Response.Write("<script>alert('�������ݳ���!');history.back();</script>")
				Exit Sub
			end if
		    Conn.Execute("delete from ks_message where id in(" & KS.FilterIDs(KS.G("id")) &")")
			Response.Write Response.Write("<script>alert('��ϲ��ɾ�������ɹ���');location.href='" & Request.ServerVariables("HTTP_REFERER") & "';</script>")
		End Sub
		
		Sub del()
			if KS.G("username")="" then
				Response.Write("<script>alert('������Ҫ����ɾ�����û���!');history.back();</script>")
				Exit Sub
			end if
			sql="delete from KS_Message where sender='"&KS.G("username")&"'"

			Conn.Execute(sql)
			
			Response.Write Response.Write("<script>alert('�����ɹ����������Ĳ���!');</script>")
		End Sub
		
		sub delall()
			dim selflag,sql
			if request("isread")="yes" then
			selflag=""
			else
			selflag=" and flag=1"
			end if
				select case request("delDate")
				case "all"
				sql="delete from KS_Message where id>0 "&selflag
				case 7
				sql="delete from KS_Message where datediff('d',sendtime,Now())>7 "&selflag
				case 30
				sql="delete from KS_Message where datediff('d',sendtime,Now())>30 "&selflag
				case 60
				sql="delete from KS_Message where datediff('d',sendtime,Now())>60 "&selflag
				case 180
				sql="delete from KS_Message where datediff('d',sendtime,Now())>180 "&selflag
				end select
				Conn.Execute(sql)

			Call KS.Alert("�����ɹ����������Ĳ�����","KS.UserMessage.asp")
		end Sub
		
		Sub delchk()
			if request.form("keyword")="" then
				KS.ShowError("������ؼ��֣�")
				Exit sub
			end if
			if request.form("selaction")=1 then
					conn.Execute("delete from KS_Message where title like '%"&replace(request.form("keyword"),"'","")&"%'")
			elseif request.form("selaction")=2 then
				
					conn.Execute("delete from KS_Message where content like '%"&replace(request.form("keyword"),"'","")&"%'")
			else
				KS.ShowError("δָ����ز�����")
			end if
			Call KS.Alert("�����ɹ����������Ĳ�����","KS.UserMessage.asp")
		End Sub
		
		Sub SaveMsg()
			Server.ScriptTimeout=99999
			Dim UserType
			UserType = Trim(Request.Form("UserType"))
			Title	 = Trim(Request.Form("title"))
			Message  = Request.Form("message")
			If Title="" or Message="" Then
				KS.Showerror("����д��Ϣ�ı��������!")
				Exit Sub
			End If
			If Len(Message) > KS.Setting(48) Then
				KS.Showerror("��Ϣ���ݲ��ܶ���" & KS.Setting(48) & "�ֽ�")
				Exit Sub
			End If 
 
			Select Case UserType
			Case "0" : SaveMsg_0()	'�������û�
			Case "1" : SaveMsg_1()	'��ָ���û�
			Case "2" : SaveMsg_2()	'��ָ���û���
			Case Else
				KS.Showerror("���������ŵ��û�!") : Exit Sub
			End Select
			Call KS.Alert("�����ɹ������η���"&Numc+1&"���û����������Ĳ�����","KS.UserMessage.asp")
		End Sub
		'�������û�����
		Sub SaveMsg_0()
			Dim Rs,Sql,i
			Sql = "Select UserName From KS_User Order By UserID Desc"
			Set Rs = Conn.Execute(Sql)
			If Not Rs.eof Then
				SQL = Rs.GetRows(-1)
				Numc= Ubound(SQL,2)
				For i=0 To Numc
					Conn.Execute("insert into KS_Message (incept,sender,title,content,sendtime,flag,issend,DelR,DelS) values('"&SQL(0,i)&"','"&KS.C("AdminName")&"','"&Title&"','"&Message&"',"&SqlNowString&",0,1,0,0)")
				Next
			End If
			Rs.Close : Set Rs = Nothing
		End Sub
		'��ָ���û�
		Sub SaveMsg_1()
			Dim ToUserName,Rs,Sql,i
			ToUserName = Trim(Request.Form("UserName"))
			If ToUserName = "" Then
				KS.Showerror("����дĿ���û�����ע�����ִ�Сд��")
				Exit Sub
			End If
			ToUserName = Replace(ToUserName,"'","")
			ToUserName = Split(ToUserName,",")
			Numc= Ubound(ToUserName)
			For i=0 To Numc
				SQL = "Select UserName From KS_User Where UserName = '"&ToUserName(i)&"'"
				Set Rs = Conn.Execute(SQL)
				If Not Rs.eof Then
				Conn.Execute("insert into KS_Message (incept,sender,title,content,sendtime,flag,issend,DelR,DelS) values('"&ToUserName(i)&"','"&KS.C("AdminName")&"','"&Title&"','"&Message&"',"&SqlNowString&",0,1,0,0)")
				End If
			Next
			Rs.Close : Set Rs = Nothing
		End Sub
		'��ָ���û��鼰��������
		Sub SaveMsg_2()
			Dim GroupID,ErrMsg,i
			Dim SearchStr,TempValue,DayStr
			GroupID = Replace(Request.Form("GroupID"),chr(32),"")
			If GroupID="" Then
			    ErrMsg = "����ȷѡȡ��Ӧ���û��顣"
			ElseIf GroupID<>"" and Not Isnumeric(Replace(GroupID,",","")) Then
				ErrMsg = "����ȷѡȡ��Ӧ���û��顣"
			Else
				GroupID = KS.R(GroupID)
			End If
			DayStr = "'d'"
			If Instr(GroupID,",")>0 Then
				SearchStr = "GroupID in ("&GroupID&")"
			Else
				SearchStr = "GroupID = "&KS.R(GroupID)
			End If
			'��½����
			TempValue = Request.Form("Logins")
			If TempValue<>"" and IsNumeric(TempValue) Then
				SearchStr = GetSearchString(TempValue,SearchStr,Request.Form("LoginsType"),"LoginTimes")
			End If
			'��������
			TempValue = Request.Form("UserArticle")
			If TempValue<>"" and IsNumeric(TempValue) Then
				SearchStr = GetSearchString(TempValue,SearchStr,Request.Form("UserArticleType"),"(select count(id) from ks_iteminfo where inputer=ks_user.username)")
			End If
			'����½ʱ��
			TempValue = Request.Form("LoginTime")
			If TempValue<>"" and IsNumeric(TempValue) Then
				SearchStr = GetSearchString(TempValue,SearchStr,Request.Form("LoginTimeType"),"Datediff("&DayStr&",LastLoginTime,"&SqlNowString&")")
			End If
			'ע��ʱ��
			TempValue = Request.Form("RegTime")
			If TempValue<>"" and IsNumeric(TempValue) Then
				SearchStr = GetSearchString(TempValue,SearchStr,Request.Form("RegTimeType"),"Datediff("&DayStr&",JoinDate,"&SqlNowString&")")
			End If
			If SearchStr="" Then
				ErrMsg = "����д���͵�����ѡ�"
			End If
			If ErrMsg<>"" Then KS.Showerror(ErrMsg) : Exit Sub
			Dim Rs,Sql
			Sql = "Select UserName From KS_User Where "& SearchStr & " Order By UserID Desc"
			
			Set Rs = Conn.Execute(Sql)
			If Not Rs.eof Then
				SQL = Rs.GetRows(-1)
				Numc= Ubound(SQL,2)
				For i=0 To Numc
					Conn.Execute("insert into KS_Message (incept,sender,title,content,sendtime,flag,issend,DelR,DelS) values('"&SQL(0,i)&"','"&KS.C("AdminName")&"','"&Title&"','"&Message&"',"&SqlNowString&",0,1,0,0)")
				Next
			End If
			Rs.Close : Set Rs = Nothing
		End Sub
		
		Function GetSearchString(Get_Value,Get_SearchStr,UpType,UpColumn)
			Get_Value = Clng(Get_Value)
			If Get_SearchStr<>"" Then Get_SearchStr = Get_SearchStr & " and " 
			If UpType="1" Then
				Get_SearchStr = Get_SearchStr & UpColumn &" <= "&Get_Value
			Else
				Get_SearchStr = Get_SearchStr & UpColumn &" >= "&Get_Value
			End If
			GetSearchString = Get_SearchStr
		End Function
End Class
%> 
