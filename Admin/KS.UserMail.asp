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
Dim KSCls
Set KSCls = New Admin_UserMail
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_UserMail
        Private KS
		Private Action
		Private Title, Content,sendername,senderemail, Numc,groupid,sendfile

		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		    If Not KS.ReturnPowerResult(0, "KMUA10009") Then
			  Response.Write ("<script>parent.frames['BottomFrame'].location.href='javascript:history.back();';</script>")
			  Call KS.ReturnErr(1, "")
			End If
	        Response.Write "<html>"
			Response.Write"<head>"
			Response.Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			Response.Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			Response.Write "<script src=""Include/Common.js"" language=""JavaScript""></script>"
			Response.Write"</head>"
			Response.Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			Response.Write"<table width=""100%""  height=""25"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
			Response.Write " <tr>"
			Response.Write"	<td height=""25"" class=""sort""> "
			Response.Write " &nbsp;&nbsp;<strong>������:</strong><a href='KS.UserMail.asp'>�����ʼ�</a> | <a href='?action=MailOut'>�����ʼ�</a>"
			Response.Write	" </td>"
			Response.Write " </tr>"
			Response.Write"</TABLE>"
		Action=Trim(Request("Action"))
		Select Case Action
		Case "Send"
			call send()
		Case "MailOut"
		    call MailOut()
		Case "DoExport"  '�������ı��ļ�
		    call DoExport()
		Case else
			call sendmsg()
		end Select
		Response.Write "<div style=""text-align:center;color:#003300"">-----------------------------------------------------------------------------------------------------------</div>"
		Response.Write "<div style=""height:30px;text-align:center"">KeSion CMS V 6.5, Copyright (c) 2006-2010 <a href=""http://www.kesion.com/"" target=""_blank""><font color=#ff6600>KeSion.Com</font></a>. All Rights Reserved . </div>"%>
		</body>
		</html>
		<%
		End Sub
		Sub SendMsg()
		%>
		<SCRIPT language=JavaScript>
function CheckForm(){
  if (document.myform.title.value==''){
     alert('�ʼ����ⲻ��Ϊ�գ�');
     document.myform.title.focus();
     return false;
  }
   if ((FCKeditorAPI.GetInstance('Content').GetXHTML(true)==""))
	{
	  alert("�ʼ����ݲ���Ϊ�գ�");
	  FCKeditorAPI.GetInstance('Content').Focus();
	  return false;
   } 

  return true;  
}
</SCRIPT>
</head>
<body><br>
 <% 
 dim InceptType:InceptType=KS.g("InceptType")
 if InceptType="" then InceptType="0"
 dim userid:userid=KS.FilterIds(replace(request("userid")," ",""))
 dim usernamelist
 if userid<>"" then
	 dim rs:set rs=KS.InitialObject("adodb.recordset")
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
 %>
  <table class="ctable" cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
<FORM name=myform onSubmit="return CheckForm();" action=KS.UserMail.asp method=post>
    <tr class=sort>
      <td align=middle colSpan=2 height=22><B>�� �� �� ��</B></td>
    </tr>
    <tr class=tdbg>
      <td align=right class="clefttitle">�ռ���ѡ��</td>
      <td>
        <table>
          <tr>
            <td>
              <Input type=radio<%if InceptType="0" then response.write " CHECKED"%> value="0" name=InceptType> ���л�Ա</td>
            <td></td>
          </tr>
          <tr>
            <td vAlign=top>
              <Input type=radio value="1" name=InceptType<%if InceptType="1" then response.write " CHECKED"%>> ָ����Ա��</td>
            <td><%=KS.GetUserGroup_CheckBox("GroupID",0,4)%></td>
          </tr>
          <tr>
            <td vAlign=top>
              <Input type=radio value="2" name=InceptType<%if InceptType="2" then response.write " CHECKED"%>> ָ���û���</td>
            <td>
              <Input size=40 name=inceptUser value="<%=usernamelist%>" class="textbox">
              ����û���������<font color=#0000ff>Ӣ�ĵĶ���</font>�ָ�</td>
          </tr>
          <tr>
            <td vAlign=top>
              <Input type=radio value="3" name=InceptType<%if InceptType="3" then response.write " CHECKED"%>>              ָ��Email</td>
            <td>
              <Input size=40 name=InceptEmail class="textbox"> 
              ���Email������<font color=#0000ff>Ӣ�ĵĶ���</font>�ָ�</td>
          </tr>
        </table>      </td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%" class="clefttitle">�ʼ����⣺</td>
      <td width="85%">
        <Input size=64 name=title class="textbox"> </td>
    </tr>
    <tr class=tdbg>
      <td align=right class="clefttitle">�ʼ����ݣ�</td>
      <td><TEXTAREA id=Content style="DISPLAY: none" name=Content></TEXTAREA> 
	  <iframe id="content___Frame" src="../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=Content&amp;Toolbar=NewsTool" width="695" height="200" frameborder="0" scrolling="no"></iframe>
	  
    </tr>
    <tr class=tdbg>
      <td align=right class="clefttitle">ѡ�񸽼���</td>
      <td>����һ.<Input size=30 type="file" name="sendfile"><br />
	  ������.<Input size=30 type="file" name="sendfile"><br />
	  ������.<Input size=30 type="file" name="sendfile"><br />
	  
	  <font color=red>˵���������ļ��������������ƣ�������ܵ��·���ʧ�ܣ�</font>
	  </td>
    </tr>
	
    <tr class=tdbg>
      <td align=right width="15%" class="clefttitle">�����ˣ�</td>
      <td width="85%">
        <Input size=64 value="<%=KS.Setting(0)%>" name=sendername class="textbox"> </td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%" class="clefttitle">������Email��</td>
      <td width="85%">
        <Input size=64 value="<%=KS.Setting(11)%>" name=senderemail class="textbox"> </td>
    </tr>
    <tr class=tdbg>
      <td align=right class="clefttitle">�ʼ����ȼ���</td>
      <td>
  <Input name=Priority type=radio value=1 checked="checked"> 
  �� 
  <Input type=radio value=3 name=Priority> 
  ��ͨ 
        <Input type=radio value=5 name=Priority> �� </td>
    </tr>
    <tr class=tdbg>
      <td align=middle colSpan=2>
  <Input id=Action type=hidden value=Send name=Action> 
  <Input id=Submit type=submit value=" �� �� " name=Submit>&nbsp; 
        <Input id=Reset type=reset value=" �� �� " name=Reset> </td>
    </tr>
</FORM>
  </table>
		<%
		end sub
		
		Sub Send()
			Server.ScriptTimeout=99999
			Dim InceptType
			InceptType = Trim(Request.Form("InceptType"))
			Title	 = Trim(Request.Form("title"))
			Content  = Request.Form("Content")
			sendername =KS.G("sendername")
			senderemail=KS.G("senderemail")
			sendfile=Request.Form("sendfile")
			
			If Title="" or Content="" Then
				KS.Showerror("����д�ʼ������������!")
				Exit Sub
			End If
			Numc=0
			Select Case InceptType
			Case "0" : SaveMsg_0()	'�������û�
			Case "1" : SaveMsg_1()	'��ָ���û���
			Case "2" : SaveMsg_2()	'��ָ���û�
			Case "3" : SaveMsg_3()  'ָ������
			Case Else
				KS.Showerror("���������ŵ��û�!") : Exit Sub
			End Select
			Call KS.Alert("�����ɹ������η���"&Numc&"���û����������Ĳ�����","KS.UserMail.asp")
		End Sub
		'�������û�����
		Sub SaveMsg_0()
			Dim Rs,Sql,i
			Sql = "Select Email From KS_User Order By UserID Desc"
			Set Rs = Conn.Execute(Sql)
			If Not Rs.eof Then
				SQL = Rs.GetRows(-1)
				For i=0 To Ubound(SQL,2)
				  if Not IsNull(SQL(0,i)) and SQL(0,i)<>"" then
				     Dim ReturnInfo:ReturnInfo=SendMail(KS.Setting(12),  KS.Setting(13), KS.Setting(14), Title, SQL(0,i),sendername, Content,senderemail)
					  IF ReturnInfo="OK" Then  Numc=Numc+1
				  end if
				Next
			End If
			Rs.Close : Set Rs = Nothing
		End Sub
		'ָ���û���
		Sub SaveMsg_1()
		    GroupID = Replace(Request.Form("GroupID"),chr(32),"")
			If GroupID<>"" and Not Isnumeric(Replace(GroupID,",","")) Then
				ErrMsg = "����ȷѡȡ��Ӧ���û��顣"
			Else
				GroupID = KS.R(GroupID)
			End If
			Dim Rs,Sql,i
			Sql = "Select Email From KS_User Where GroupID in(" & GroupID & ") Order By UserID Desc"
			Set Rs = Conn.Execute(Sql)
			If Not Rs.eof Then
				SQL = Rs.GetRows(-1)
				For i=0 To Ubound(SQL,2)
				  if Not IsNull(SQL(0,i)) and SQL(0,i)<>"" then
				     Dim ReturnInfo:ReturnInfo=SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), Title, SQL(0,i),sendername, Content,senderemail)
					  IF ReturnInfo="OK" Then  Numc=Numc+1
				  end if
				Next
			End If
			Rs.Close : Set Rs = Nothing
		End Sub
		'��ָ���û�
		Sub SaveMsg_2()
			Dim inceptUser,Rs,Sql,i
			inceptUser = Trim(Request.Form("inceptUser"))
			If inceptUser = "" Then
				KS.Showerror("����дĿ���û�����ע�����ִ�Сд��")
				Exit Sub
			End If
			inceptUser = Replace(inceptUser,"'","")
			inceptUser = Split(inceptUser,",")
			For i=0 To ubound(inceptUser)
				SQL = "Select Email From KS_User Where UserName = '"&inceptUser(i)&"'"
				Set Rs = Conn.Execute(SQL)
				If Not Rs.eof Then
				  if Not IsNull(rs(0)) and rs(0)<>"" then
				     Dim ReturnInfo:ReturnInfo=SendMail(KS.Setting(12),  KS.Setting(13), KS.Setting(14), Title, rs(0),sendername, Content,senderemail)
					  IF ReturnInfo="OK" Then  Numc=Numc+1
				  end if
				End If
			Next
			Rs.Close : Set Rs = Nothing
		End Sub
		'��ָ������
		Sub SaveMsg_3()
			Dim InceptEmail,Rs,Sql,i
			InceptEmail = Trim(Request.Form("InceptEmail"))
			If InceptEmail = "" Then
				KS.Showerror("����д�����͵��ʼ���ַ!")
				Exit Sub
			End If
			InceptEmail = Replace(InceptEmail,"'","")
			InceptEmail = Split(InceptEmail,",")
			For i=0 To ubound(InceptEmail)
				Dim ReturnInfo:ReturnInfo=SendMail(KS.Setting(12), KS.Setting(13), KS.Setting(14), Title, InceptEmail(i),sendername, Content,senderemail)
				IF ReturnInfo="OK" Then  Numc=Numc+1
			Next
		End Sub
        
		'�����ʼ�
		Sub MailOut()
		%>
		<br>
  <table class=border cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
<FORM action="?Action=DoExport" method=post>
    <tr class=title>
      <td class=title align=middle colSpan=2 height=22><B>�ʼ��б��������������ݿ�</B></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="24%" height=80>�����ʼ��б����ݿ⣺</td>
      <td width="76%" height=80>
  <Input id=ExportType type=hidden value=1 name=ExportType> &nbsp;&nbsp;<font color=blue>����</font>&nbsp;&nbsp; 
<Select id=GroupID name=GroupID>
  <Option value=0 selected>ȫ����Ա</Option>
<%=KS.GetUserGroup_Option(0)%>
</Select> &nbsp;<font color=blue>��</font>&nbsp; 
  <Input id=ExportFileName maxLength=200 size=30 value=<%=KS.Setting(3)%>usermail.mdb name=ExportFileName> 
        <Input type=submit value=��ʼ name=Submit> </td>
    </tr>
</FORM>
  </table>
<br>
  <table class=border cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
<FORM action="?Action=DoExport" method=post>
    <tr class=title>
      <td class=title align=middle colSpan=2 height=22><B>�ʼ��б������������ı�</B></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="24%" height=80>�����ʼ��б��ı���</td>
      <td width="76%" height=80>
  <Input id=ExportType type=hidden value=2 name=ExportType> &nbsp;&nbsp;<font color=blue>����</font>&nbsp;&nbsp; 
<Select id=GroupID name=GroupID>
  <Option value=0 selected>ȫ����Ա</Option>
<%=KS.GetUserGroup_Option(0)%>
</Select> 
</Select>&nbsp;<font color=blue>��</font>&nbsp; 
  <Input id=ExportFileName maxLength=200 size=30 value=<%=KS.Setting(3)%>usermail.txt name=ExportFileName> 
        <Input type=submit value=��ʼ���� name=Submit2> </td>
    </tr>
</FORM>
  </table>

		<%
		End Sub
		
		'�������ı��ļ�
		Sub DoExport()
		 Dim ExportFileName:ExportFileName=KS.G("ExportFileName")
		 Dim GroupID:GroupID=KS.G("GroupID")
		 Dim ExportType:ExportType=KS.G("ExportType")
		 Dim rs:set rs=KS.InitialObject("adodb.recordset")
		 Dim sqlstr,MailList,n
		   n=0
		  if GroupID="0" then
		    sqlstr="select email from ks_user"
		  else
		    sqlstr="select email from ks_user where groupid=" & groupid
		  end if
			 If ExportType=2 Then
			    		 rs.open sqlstr,conn,1,1
						 if not rs.eof then
						   do while not rs.eof
						      if rs(0)<>"" and not isnull(rs(0)) then
							  n=n+1
							  MailList=MailList & rs(0) & vbcrlf
							  end if
							  rs.movenext
						   loop
						 end if
						  rs.close:set rs=nothing
				Dim FSO:Set FSO = KS.InitialObject(KS.Setting(99))
				Dim FileObj:Set FileObj = FSO.CreateTextFile(Server.MapPath(ExportFileName), True) '�����ļ�
				FileObj.Write MailList
				FileObj.Close     '�ͷŶ���
				Set FileObj = Nothing:Set FSO = Nothing
			 Else
			      on error resume next
			     if CreateDatabase(ExportFileName)=true then
						Dim DataConn:Set DataConn = KS.InitialObject("ADODB.Connection")
	                    DataConn.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(ExportFileName)
						If not Err Then
						   If Checktable("UserEmail",DataConn)=true Then
						     DataConn.Execute("drop table useremail")
						   end if
				             Dataconn.execute("CREATE TABLE [UserEMail] ([ID] int IDENTITY (1, 1) NOT NULL CONSTRAINT PrimaryKey PRIMARY KEY,[Email] varchar(255) Not Null)")
						  rs.open sqlstr,conn,1,1
						 if not rs.eof then
						   do while not rs.eof
						      if rs(0)<>"" and not isnull(rs(0)) then
							  n=n+1
						      DataConn.Execute("Insert Into UserEmail(email) values('" & rs(0) &"')")
							  end if
							  rs.movenext
						   loop
						 end if
                          rs.close:set rs=nothing
						End if
						DataConn.Close:Set DataConn=Nothing
				 end if
			 
			 End If
		  response.write "<br><br><br><div align=center>�������!�ɹ������� <font color=red>" & n & "</font> ���ʼ���ַ��<a href=" & ExportFileName & ">������������</a>(�Ҽ�Ŀ�����Ϊ)  </div><br><br><br><br><br><br><br>"
		End Sub
		Function CreateDatabase(dbname)
		      if KS.CheckFile(dbname) then CreateDatabase=true:exit function
				dim objcreate :set objcreate=KS.InitialObject("adox.catalog") 
				if err.number<>0 then 
					set objcreate=nothing 
					CreateDatabase=false
					exit function 
				end if 
				'�������ݿ� 
				objcreate.create("data source="+server.mappath(dbname)+";provider=microsoft.jet.oledb.4.0") 
				if err.number<>0 then 
					CreateDatabase=false
					set objcreate=nothing 
					exit function
				end if 
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
		
		Public Function SendMail(MailAddress, LoginName, LoginPass, Subject, Email, Sender, Content, Fromer)
	   'On Error Resume Next
		Dim JMail
		  Set jmail = Server.CreateObject("JMAIL.Message") '���������ʼ��Ķ���
			jmail.silent = true '����������󣬷���FALSE��TRUE��ֵj
			jmail.Charset = "GB2312" '�ʼ������ֱ���Ϊ����
			If sendfile="" Then
			' jmail.ContentType = "text/html" '�ʼ��ĸ�ʽΪHTML��ʽ,��������ʱ�ſ���
			End If
			jmail.AddRecipient Email '�ʼ��ռ��˵ĵ�ַ
			jmail.From = Fromer '�����˵�E-MAIL��ַ
			jmail.FromName = Sender
			  If LoginName <> "" And LoginPass <> "" Then
				JMail.MailServerUserName = LoginName '�����ʼ���������¼��
				JMail.MailServerPassword = LoginPass '��¼����
			  End If
			jmail.Subject = Subject '�ʼ��ı��� 
			JMail.Body = Content
			JMail.HTMLBody = Content
			Dim I,sendfileArr:SendFileArr=Split(sendfile,",")
			For I=0 To UBound(SendFileArr)
			 if trim(sendfileArr(i))<>"" Then
			  jmail.AddAttachment trim(sendfileArr(i))
			 End If
			Next
			JMail.Priority = 1'�ʼ��Ľ�������1 Ϊ��죬5 Ϊ������ 3 ΪĬ��ֵ
			jmail.Send(MailAddress) 'ִ���ʼ����ͣ�ͨ���ʼ���������ַ��
			jmail.Close() '�رն���
		Set JMail = Nothing
		If Err Then
			SendMail = Err.Description
			Err.Clear
		Else
			SendMail = "OK"
		End If
	  End Function
End Class
%> 
