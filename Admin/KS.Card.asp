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
Set KSCls = New Admin_Card
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Card
        Private KS,CardType
		Private MaxPerPage,RS,TotalPut,TotalPages,I,CurrentPage,SQL,ComeUrl
		Private Sub Class_Initialize()
		  MaxPerPage=20
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
       Sub Kesion()
	        If Not KS.ReturnPowerResult(0, "KMUA10008") Then
			  Call KS.ReturnErr(1, "")
			End If
			CardType=KS.ChkClng(KS.G("CardType"))
          Response.Write "<html>"
			Response.Write"<head>"
			Response.Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			Response.Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css""><script src='../KS_Inc/common.js'></script><script src='../KS_Inc/jQuery.js'></script>"
			Response.Write"</head>"
			Response.Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			Response.Write"	<ul id='mt'> "
			Response.Write " <div id='mtl'>��������:</div><li>&nbsp;<a href=""?cardtype=" & cardtype & """>���г�ֵ��</a> | "
			If CardType=1 Then
			Response.Write "<a href=""?action=AddMore&cardtype=1"">�������߳�ֵ��</a></li>"
			Else
			Response.Write "<a href=""?status=1&cardtype=0"">δʹ�ó�ֵ��</a> | <a href=""?status=2&cardtype=0"">��ʹ�ó�ֵ��</a> | <a href=""?status=3&cardtype=0"">��ʧЧ��ֵ��</a> | <a href=""?status=4&cardtype=0"">δʧЧ��ֵ��</a> | <a href=""?action=Add&cardtype=0"">��ӳ�ֵ��</a> | <a href=""?action=AddMore&cardtype=0"">�������ɳ�ֵ��</a></li>"
			End If
			Response.Write	" </ul>"

		ComeUrl=Cstr(Request.ServerVariables("HTTP_REFERER"))
		
		Select Case KS.G("Action")
		 Case "Add","Edit"
		  Call Add()
		 Case  "DoAdd"
		  Call DoAdd()
		 Case "AddMore"
		  Call AddMore()
		 Case "DoAddMore"
		  Call DoAddMore()
		 Case "Del"
		  Call Del()
		 Case Else
		  Call CardList()
		End Select
	End Sub
	
	'�㿨�б�
	Sub CardList()
		%>
<table width="100%" border="0" align="center" cellspacing="1" cellpadding="0">
  <tr class="sort">
   <%if cardtype="1" then%>
    <td width="38" align="center"><strong>ѡ��</strong></td>
    <td align="center"><strong>��ֵ������</strong></td>
    <td width="75" align="center"><strong>��ֵ</strong></td>
    <td width="79" align="center" nowrap="nowrap"><strong>����/����</strong></td>
    <td align="center"><strong>����ʱ��</strong></td>
    <td align="center"><strong>����</strong></td>
   <%else%>
    <td width="38" align="center"><strong>ѡ��</strong></td>
    <td width="116" align="center"><strong>����</strong></td>
    <td width="116" align="center"><strong>��ֵ����</strong></td>
    <td width="88" align="center"><strong>����</strong></td>
    <td width="75" align="center"><strong>��ֵ</strong></td>
    <td width="79" align="center" nowrap="nowrap"><strong>����/����</strong></td>
    <td width="100" align="center"><strong>����ʱ��</strong></td>
    <td width="60" align="center"><strong>����</strong></td>
    <td width="60" align="center"><strong>ʹ��</strong></td>
    <td width="100" align="center"><strong>ʹ����</strong></td>
    <td width="100" align="center"><strong>��ֵʱ��</strong></td>
    <td width="80" align="center"><strong>����</strong></td>
  <%end if%>
  </tr>
  <%
  CurrentPage	= KS.ChkClng(request("page"))
  Dim Param:Param=" where cardtype=" &cardtype
  if KS.G("groupname")<>"" Then Param=Param & " and groupname='" & KS.G("groupname") & "'"
  if KS.G("KeyWord")<>"" Then Param=Param & " and cardnum='" & KS.G("KeyWord") & "'"
  Select Case  KS.ChkClng(KS.G("Status"))
   Case 1
     Param=Param & " And IsUsed=0"
   Case 2
     Param=Param & " And IsUsed=1"
   Case 3
     Param=Param & " And datediff('d',EndDate,"&SqlNowString&")>0"
   Case 4
     Param=Param & " And datediff('d',EndDate,"&SqlNowString&")<0"
  End Select
  
  Dim SqlStr:SqlStr="Select ID,CardNum,CardPass,Money,ValidNum,ValidUnit,AddDate,EndDate,UseDate,UserName,IsUsed,IsSale,groupname From KS_UserCard " & Param & " order by ID desc"
  Set RS=Server.CreateObject("ADODB.RecordSet")
    RS.Open SqlStr,conn,1,1
	If RS.Eof And RS.Bof Then
	 Response.Write "<tr><td colspan=11 align=center height=25>û�г�ֵ����</td></tr><tr><td colspan=13 background='images/line.gif'></td></tr>"
	Else
                    TotalPut=rs.recordcount
					if (TotalPut mod MaxPerPage)=0 then
						TotalPages = TotalPut \ MaxPerPage
					else
						TotalPages = TotalPut \ MaxPerPage + 1
					end if
					if CurrentPage > TotalPages then CurrentPage=TotalPages
					if CurrentPage < 1 then CurrentPage=1
					rs.move (CurrentPage-1)*MaxPerPage
					SQL = rs.GetRows(MaxPerPage)
					rs.Close:set rs=Nothing
					ShowContent
   End If
%>		
</table>
<%if cardtype=0 then%>
<table border="0" style="margin-top:20px" width="100%" align=center>
<form action="KS.Card.asp" name="myform" method="post">
<tr><td>
<div style="border:1px dashed #cccccc;margin:3px;padding:4px"><b>��������=></b>:
<%
	           Response.Write " &nbsp;<select name='groupname'>"
			   Response.Write "<option value=''>====ѡ�����====</option>"
				 Dim ZRS:Set ZRS=Server.CreateObject("ADODB.RECORDSET")
				 ZRS.Open "select Distinct groupname from ks_usercard where groupname<>'' and groupname<>null",conn,1,1
				 If Not ZRS.Eof Then
				  Do While Not ZRS.Eof 
				   if ks.g("groupname")=zrs(0) then
				   Response.Write "<option value='" & ZRS(0) & "' selected>" & ZRS(0) & "</option>"
				   else
				   Response.Write "<option value='" & ZRS(0) & "'>" & ZRS(0) & "</option>"
				   end if
				   ZRS.MoveNext
				  Loop
				 End If
				 ZRS.Close:Set ZRS=Nothing
			    Response.Write "</select>"
	  %>
����
<input type="text" name="keyword" value="" class='textbox' size="20">&nbsp;<input type="submit" value="��ʼ����" class="button">
</div>
</td></tr>
</form>
<tr><td><br><Font color=red><strong>��ʾ��</strong>
���۳�����ʹ�õĳ�ֵ����������ɾ�����޸ĵȲ�����</font>
</td></tr>
</table>
<%
 end if
End Sub
Sub ShowContent
 Dim InPoint,OutPoint
 %>
 <form name=selform method=post action=?action=Del&cardtype=<%=cardtype%>>
 <%
For i=0 To Ubound(SQL,2)
	%>
  <tr height="22" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
    <td align="center"><input type="checkbox" name="id" value="<%=SQL(0,i)%>"></td>
    <td align="center"><%=SQL(12,i)%></td>
	<%if cardtype="1" then%>
		<td align="center"><%=formatnumber(SQL(3,i),2,-1)%>Ԫ</td>
		<td align="center"><%Response.Write SQL(4,I)
		if SQL(5,I)=1 Then 
		 Response.Write "��" 
		ELSEIf SQL(5,I)=2 Then 
		 Response.Write "��" 
		elseif SQL(5,I)=3 Then
		 response.write "Ԫ"
		end if%></td>
    <td align="center"><%Response.Write formatdatetime(SQL(7,I),2)%></td>
	<%Else%>
    <td align="center"><%=SQL(1,i)%></td>
    <td align="center"><%=KS.Decrypt(SQL(2,i))%></td>
    <td align="center"><%=formatnumber(SQL(3,i),2,-1)%>Ԫ</td>
    <td align="center"><%Response.Write SQL(4,I)
	if SQL(5,I)=1 Then 
	 Response.Write "��" 
	ELSEIf SQL(5,I)=2 Then 
	 Response.Write "��" 
	elseif SQL(5,I)=3 Then
	 response.write "Ԫ"
	end if%></td>
    <td align="center"><%Response.Write formatdatetime(SQL(7,I),2)%></td>
    <td align="center">
	<%
	IF SQL(11,I)=1 Then
	 Response.Write "���۳�"
	Else
	 Response.Write "<font color=red>δ����</font>" 
	End If
	%></td>
    <td align="center">
	<%
	IF SQL(10,I)=1 Then
	 Response.Write "<font color='#a7a7a7'>��ʹ��</font>"
	Else
	 Response.Write "<font color=red>δʹ��</font>" 
	End If
	%></td>
    <td align="center"><%Response.Write SQL(9,I)%></td>
    <td align="center">
	<%if Isdate(Sql(8,i)) then
	   response.write formatdatetime(SQL(8,i),2)
	  end if%></td>
	  
	<%end if%>
	<td align="center">
	<%if SQL(11,I)<>1 and SQL(10,I)<>1 then%>
	<a href="?action=Edit&ID=<%=SQL(0,i)%>&cardtype=<%=cardtype%>">�޸�</a> <a href="?action=Del&cardtype=<%=cardtype%>&ID=<%=SQL(0,i)%>">ɾ��</a>
	<%end if%>
	</td>
  </tr>
  <tr><td colspan=13 background='images/line.gif'></td></tr>
  <%Next
  
  Response.Write "<tr onMouseOver=""this.className='listmouseover'"" onMouseOut=""this.className='list'""><td height='30' colspan=4>"
  Response.Write "&nbsp;&nbsp;<input id=""chkAll"" onClick=""CheckAll(this.form)"" type=""checkbox"" value=""checkbox""  name=""chkAll"">ȫѡ&nbsp;&nbsp;<input class=Button type=""submit"" name=""Submit2"" value="" ɾ����ֵ�� "" onclick=""{if(confirm('�˲��������棬ȷ��Ҫɾ��ѡ�еļ�¼��?')){this.document.selform.submit();return true;}return false;}""> <input type='button' value=' �� ӡ ' onclick='window.print()' class='button'></td><td colspan=13 align=right><br>"
  Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
  Response.Write "</td></tr>  <tr><td colspan=13 background='images/line.gif'></td></tr></form>"
End Sub

 'ɾ����ֵ��
 Sub Del()
  Dim ID:ID=Replace(KS.G("ID")," ","")
  ID=KS.FilterIDs(ID)
  If ID="" Then Response.Write "<script>alert('��ѡ���ֵ��!');history.back();</script>"
  Conn.Execute("Delete From KS_UserCard Where ID In(" & ID &") and IsSale=0 and IsUsed=0")
  Response.Write "<script>alert('ɾ���ɹ���');location.href='" & Request.Servervariables("http_referer") & "';</script>"
 End Sub
		
		'������ӳ�ֵ��
  Sub AddMore()
		%>
  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1'>
		<form method='post' action='KS.Card.asp' name='myform'>
    <tr class='sort'> 
      <td height='22' colspan='2'> <div align='center'><strong>�� �� �� �� �� ֵ ��</strong></div></td>
    </tr>
    <tr class='tdbg'> 
      <td width='40%' class="clefttitle"><strong>��ֵ�����ƣ�</strong></td>
      <td width='60%'><input name='GroupName' type='text' size='20' maxlength='100'> ��:��10Ԫ��100���ֵ������</td>
    </tr>
    <tr class='tdbg'> 
      <td width='40%' class="clefttitle"><strong>��ֵ��ʽ��</strong></td>
      <td width='60%'>
	  <input type="radio" name="cardtype" value="1"  onclick="$('#showother').hide()"<%if cardtype=0 then response.write " disabled" else response.write " checked"%>>���߳�ֵ����(��Ա�����ɻ�Ա�Լ������ֵ��<br/>
	  <input type="radio" name="cardtype" value="0" onclick="$('#showother').show()"<%if cardtype=1 then response.write " disabled" else response.write " checked"%>>��������
	  </td>
    </tr>
   <tbody id="showother" style="display:<%if cardtype=0 then response.write "" else response.write "none"%>">
    <tr class='tdbg'> 
      <td width='40%' class="clefttitle"><strong>��ֵ��������</strong></td>
      <td width='60%'><input name='Nums' type='text' value='100' size='10' maxlength='10'>
        ��</td>
    </tr>
    <tr class='tdbg'>
      <td width='40%' class="clefttitle"><strong>��ֵ������ǰ׺��</strong><br>
        ���磺2006,KS2006�ȹ̶��������ĸ������</td>
      <td width='60%'><input name='CardNumPrefix' type='text' id='CardNumPrefix' value='KS2007' size='10' maxlength='10'></td>
    </tr>
    <tr class='tdbg'>
      <td width='40%' class="clefttitle"><strong>��ֵ������λ����</strong><br>���������ǰ׺�ַ����ڵ���λ��</td>
      <td width='60%'><input name='CardNumLen' type='text' id='CardNumLen' value='12' size='10' maxlength='10'>
        <font color='#0000FF'>������Ϊ10--15λ</font></td>
    </tr>
    <tr class='tdbg'>
      <td width='40%' class="clefttitle"><strong>��ֵ������λ����</strong></td>
      <td width='60%'><input name='PasswordLen' type='text' id='PasswordLen' value='6' size='10' maxlength='10'>
        <font color='#0000FF'>������Ϊ6--10λ</font></td>
    </tr>
    <tr class='tdbg'>
      <td class="clefttitle"><strong>�����빹�ɷ�ʽ��</strong><br>�����ѡ�����ݻ���ĸ�����</td>
      <td><input type="radio" name="zhtype" value="1" checked>������ <input type="radio" name="zhtype" value="2">��������ĸ������ </td>
    </tr>
  </tbody>
    <tr class='tdbg'>
      <td width='40%' class="clefttitle"><strong>��ֵ����ֵ��</strong><br>
      ����������Ҫ���ѵ�ʵ�ʽ��</td>
      <td width='60%'><input name='Money' type='text' id='Money' value='50' size='10'>
      Ԫ</td>
    </tr>
    <tr class='tdbg'> 
      <td width='40%' class="clefttitle"><strong>��ֵ���������ʽ����Ч�ڣ�</strong><br>
        �����˿��Եõ��ĵ������ʽ���Ч�ںͻ���      </td>
      <td width='60%'><input name='ValidNum' type='text' id='ValidNum' value='50' size='10' maxlength='10'>
        <select name='ValidUnit' id='ValidUnit'>
          <option value='1' selected>��</option>
          <option value='2'>��</option>
          <option value='3'>Ԫ</option>
          <option value='4'>����</option>
        </select></td>
    </tr>
	
    <tr class='tdbg'>
      <td width='40%' class="clefttitle"><strong>����ʹ�ô˳�ֵ�����û��飺</strong><br>
	  �����������ջ�ȫ��ѡ�С�
     </td>
      <td width='60%'><%=KS.GetUserGroup_CheckBox("AllowGroupID","",5)%></td>
    </tr>
    <tr class='tdbg'>
      <td width='40%' class="clefttitle"><strong>��ֵ���Զ�������û��飺</strong><br>
     </td>
      <td width='60%'><select name="GroupID" id="GroupID">
	  <option value='0'>---����ԭ���û���---</option>
	<%=KS.GetUserGroup_Option(0)%>
	 </select></td>
    </tr>
    <tr class='tdbg'>
      <td width='40%' class="clefttitle"><strong>���ں��Զ�������û��飺</strong><br>
	  <span style='color:blue'>ָ�û�ѡ���ֵ��Ϊ�˻���ֵ��,���˻���ĵ�ȯ,��Ч�������ʽ������(������ݸÿ��ǵ�ȯ��,�������������ʽ𿨶���)�������ڵ��û��Զ������һ�����û�����</span>
     </td>
      <td width='60%'><select name="ExpireGroupID" id="ExpireGroupID">
	  <option value='0'>---����ԭ���û���---</option>
	<%=KS.GetUserGroup_Option(0)%>
	 </select></td>
    </tr>
	
    <tr class='tdbg'>
      <td width='40%' class="clefttitle"><strong>��ֵ��ֹ���ޣ�</strong><br>
      �����˱����ڴ�����ǰ���г�ֵ�������Զ�ʧЧ</td>
      <td width='60%' class='tdbg'><input name='EndDate' type='text' id='EndDate' value='<%=dateadd("yyyy",2,now)%>' size='20'></td>
    </tr>
    <tr class='tdbg'> 
      <td height='40' colspan='2' align='center'><input name='Action' type='hidden' id='Action' value='DoAddMore'> 
        <input  type='submit' class='button' name='Submit' value=' ��ʼ���� ' style='cursor:pointer;'> 
        &nbsp; <input class='button' name='Cancel' type='button' id='Cancel' value=' ȡ �� ' onClick="window.location.href='KS.Card.asp'" style='cursor:pointer;'></td>
    </tr>
  </table>
</form>
		<%
		End Sub	
		'��ӳ�ֵ��
		Sub Add()
		  Dim CardNum,PassWord,IsSale,IsUsed,Money,ValidNum,ValidUnit,EndDate,action1,GroupName,AllowGroupID,GroupID,ExpireGroupID,cardtype
		  Dim ID:ID=KS.ChkClng(KS.G("ID"))
		  if KS.g("action")="Edit" then
		    Dim RS:Set RS=Server.CreateObject("adodb.recordset")
			rs.open "select top 1 * from ks_usercard where ID=" & ID,conn,1,1
			if rs.bof and rs.eof then
			  rs.close:set rs=nothing
			  Call KS.AlertHistory("�������ݳ���",-1)
			  Exit sub
			end if
			CardNum=rs("CardNum")
			PassWord=KS.Decrypt(rs("CardPass"))
			Money=rs("money")
			ValidNum=rs("ValidNum")
			ValidUnit=rs("ValidUnit")
			EndDate=rs("EndDate")
			IsSale=rs("IsSale")
			IsUsed=rs("IsUsed")
			cardtype=rs("cardtype")
			GroupName=rs("GroupName")
			AllowGroupID=rs("allowgroupid")
			GroupID=rs("groupid")
			ExpireGroupID=rs("expiregroupid")
			action1="Edit"
		  else
		   IsSale=0:IsUsed=0:Money=50:ValidNum=50:ValidUnit=1:EndDate=Now+365
		  end if
		%>
  <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1'>
		<form method='post' action='KS.Card.asp?action1=<%=action1%>&id=<%=ID%>' name='myform'>
    <tr class='sort'> 
      <td height='22' colspan='2'> <div align='center'><strong>
	  <%IF KS.g("Action")="Edit" then
	   response.write "�� �� �� ֵ ��"
	    Else
		Response.Write "�� �� �� ֵ ��"
	    End If
		%></strong></div></td>
    </tr>
    <tr class='tdbg'> 
      <td width='40%' class="clefttitle"><strong>��ֵ�����ƣ�</strong></td>
      <td width='60%'><input name='GroupName' type='text' size='20' value="<%=GroupName%>" maxlength='100'> ��:��10Ԫ��100���ֵ������</td>
    </tr>
    <tr class='tdbg'> 
      <td width='40%' class="clefttitle"><strong>��ֵ��ʽ��</strong></td>
      <td width='60%'>
	  <input type="radio" name="cardtype" value="1"  onclick="$('#showother').hide()"<%if cardtype=0 then response.write " disabled" else response.write " checked"%>>���߳�ֵ����(��Ա�����ɻ�Ա�Լ������ֵ��<br/>
	  <input type="radio" name="cardtype" value="0" onclick="$('#showother').show()"<%if cardtype=1 then response.write " disabled" else response.write " checked"%>>��������
	  </td>
    </tr>
   <tbody id="showother" style="display:<%if cardtype=0 then response.write "" else response.write "none"%>">

    <tr class='tdbg'<%if KS.g("action")="Edit" Then response.write " style='display:none'"%>> 
      <td width='40%' class="clefttitle"><strong>��ӷ�ʽ��</strong></td>
      <td width='60%'><input name='AddType' type='radio' value='0' checked onclick="trSingle1.style.display='';trSingle2.style.display='';trBatch.style.display='none';"> ���ų�ֵ��&nbsp;&nbsp;&nbsp;&nbsp;<input name='AddType' type='radio' value='1' onclick="trSingle1.style.display='none';trSingle2.style.display='none';trBatch.style.display='';">������ӳ�ֵ��</td>
    </tr>
    <tr class='tdbg' id='trSingle1'>
      <td width='40%' class="clefttitle"><b>��ֵ�����ţ�</b></td>
      <td><input name='CardNum' type='text' id='CardNum' size='20' value="<%=CardNum%>" maxlength='30'>
        <font color='#0000FF'>������Ϊ10--15λ</font></td>
    </tr>
    <tr class='tdbg' id='trSingle2'>
      <td width='40%' class="clefttitle"><b>��ֵ�����룺</b></td>
      <td><input name='Password' type='text' id='Password' size='20' value="<%=PassWord%>" maxlength='30'>
        <font color='#0000FF'>������Ϊ6--10λ </font></td>
    </tr>
    <tr class='tdbg' id='trBatch' style='display:none'>
      <td width='40%' class="clefttitle"><b>��ʽ�ı���</b><br><font color='red'>�밴��ÿ��һ�ſ���ÿ�ſ��������ţ��ָ��������롱�ĸ�ʽ¼��</font><br>����734534759|kSo94Sf4Xs���ԡ�|����Ϊ�ָ�����</td>
      <td><textarea name='CardList' rows='10' cols='50'></textarea></td>
    </tr>
	</tbody>
    <tr class='tdbg'>
      <td width='40%' class="clefttitle"><strong>��ֵ����ֵ��</strong><br>
      ����������Ҫ���ѵ�ʵ�ʽ��</td>
      <td width='60%'><input name='Money' type='text' id='Money' value='<%=formatnumber(Money,2,-1)%>' size='10'>
      Ԫ</td>
    </tr>
    <tr class='tdbg'> 
      <td width='40%' class="clefttitle"><strong>��ֵ���������ʽ����Ч�ڣ�</strong><br>
        �����˿��Եõ��ĵ������ʽ���Ч�ںͻ���      </td>
      <td width='60%'><input name='ValidNum' value="<%=ValidNum%>" type='text' id='ValidNum' size='10' maxlength='10'>
        <select name='ValidUnit' id='ValidUnit'>
          <option value='1'<%if ValidUnit="1" then response.write " selected"%>>��</option>
          <option value='2'<%if ValidUnit="2" then response.write " selected"%>>��</option>
          <option value='3'<%if ValidUnit="3" then response.write " selected"%>>Ԫ</option>
          <option value='4'<%if ValidUnit="4" then response.write " selected"%>>����</option>
        </select></td>
    </tr>
	
    <tr class='tdbg'>
      <td width='40%' class="clefttitle"><strong>����ʹ�ô˳�ֵ�����û��飺</strong><br>
	  �����������ջ�ȫ��ѡ�С�
     </td>
      <td width='60%'><%=KS.GetUserGroup_CheckBox("AllowGroupID",allowgroupid,5)%></td>
    </tr>
    <tr class='tdbg'>
      <td width='40%' class="clefttitle"><strong>��ֵ���Զ�������û��飺</strong><br>
     </td>
      <td width='60%'><select name="GroupID" id="GroupID">
	  <option value='0'>---����ԭ���û���---</option>
	<%=KS.GetUserGroup_Option(groupid)%>
	 </select></td>
    </tr>
    <tr class='tdbg'>
      <td width='40%' class="clefttitle"><strong>���ں��Զ�������û��飺</strong><br>
	  <span style='color:blue'>ָ�û�ѡ���ֵ��Ϊ�˻���ֵ��,���˻���ĵ�ȯ,��Ч�������ʽ������(������ݸÿ��ǵ�ȯ��,�������������ʽ𿨶���)�������ڵ��û��Զ������һ�����û�����</span>
     </td>
      <td width='60%'><select name="ExpireGroupID" id="ExpireGroupID">
	  <option value='0'>---����ԭ���û���---</option>
	<%=KS.GetUserGroup_Option(expiregroupid)%>
	 </select></td>
    </tr>
	
	
    <tr class='tdbg'>
      <td width='40%' class="clefttitle"><strong>��ֵ��ֹ���ޣ�</strong><br>
      �����˱����ڴ�����ǰ���г�ֵ�������Զ�ʧЧ</td>
      <td width='60%' class='tdbg'><input name='EndDate' type='text' id='EndDate' value='<%=EndDate%>' size='20'></td>
    </tr>
	<tr class='tdbg'<%if cardtype=1 then response.write " style='display:none'"%>>
      <td width='40%' class="clefttitle"><strong>�Ƿ���ۣ�</strong><br>
      ����³�ֵ������ѡ��δ����</td>
      <td width='60%' class='tdbg'><input name='issale' type='radio' id='issale' value='0'<%if issale=0 then response.write " checked"%>>δ���� <input name='issale' type='radio' id='issale' value='1'<%if issale=1 then response.write " checked"%>>�ѳ���</td>
    </tr>
	<tr class='tdbg'<%if cardtype=1 then response.write " style='display:none'"%>>
      <td width='40%' class="clefttitle"><strong>�Ƿ�ʹ�ã�</strong><br>
      ����³�ֵ������ѡ��δʹ��</td>
      <td width='60%' class='tdbg'><input name='isused' type='radio' id='isused' value='0'<%if isused=0 then response.write " checked"%>>δʹ�� <input name='isused' type='radio' id='isused' value='1'<%if isused=1 then response.write " checked"%>>��ʹ��</td>
    </tr>
    <tr class='tdbg'> 
      <td height='40' colspan='2' align='center'><input name='Action' type='hidden' id='Action' value='DoAdd'> 
        <input  type='submit' name='Submit' class='button' value=' <% if KS.g("action")="Edit" then response.write "ȷ���޸�" Else Response.write "��ʼ����" %> ' style='cursor:pointer;'> 
        &nbsp; <input name='Cancel' type='button' class='button' id='Cancel' value=' ȡ �� ' onClick="window.location.href='KS.Card.asp?cardtype=<%=cardtype%>'" style='cursor:pointer;'></td>
    </tr>
	</form>

  </table>
		<%
		End Sub
		
		'��ʼ���ɳ�ֵ��
		Sub DoAdd()
		 Dim AddType:AddType=KS.G("AddType")
		 Dim CardNum:CardNum=KS.G("CardNum")
		 Dim Password:Password=KS.G("Password")
		 Dim CardList:CardList=KS.G("CardList")
		 Dim Money:Money=KS.G("Money")
		 Dim ValidNum:ValidNum=KS.ChkClng(KS.G("ValidNum"))
		 Dim ValidUnit:ValidUnit=KS.G("ValidUnit")
		 Dim EndDate:EndDate=KS.G("EndDate")
		 Dim IsUsed:IsUsed=KS.G("IsUsed")
		 Dim ISSale:IsSale=KS.G("IsSale")
		 Dim GroupName:GroupName=KS.G("GroupName")
		 Dim CardType:CardType=KS.ChkClng(KS.G("CardType"))
		 Dim AllowGroupID:AllowGroupID=KS.G("AllowGroupID")
		 Dim GroupID:GroupID=KS.ChkClng(KS.G("GroupID"))
		 Dim ExpireGroupID:ExpireGroupID=KS.ChkClng(KS.G("expiregroupid"))
		 If GroupName="" Then Call KS.AlertHistory("�������ֵ�����ƣ�",-1):exit sub
		 IF Not IsNumeric(Money) Or money="0" Then Call KS.AlertHistory("��ֵ����ֵ���������0",-1):exit sub
		 IF ValidNum=0 Then Call KS.AlertHistory("��ֵ���������������0",-1):exit sub
		 If Not IsDate(EndDate) Then Call KS.AlertHistory("��ֵ��ֹ���޸�ʽ����ȷ!",-1):exit sub
          If AddType=0 or KS.g("action1")="Edit" then
		    if cardtype=0 Then
				if CardNum="" then call KS.AlertHistory("��û�������ֵ����!",-1):exit sub
				if PassWord=" "then call KS.AlertHistory("��û�������ֵ������",-1):exit sub
			end if
			
			   Dim RS:Set RS=Server.CreateObject("adodb.recordset")
			    if KS.g("action1")="Edit" then
				 rs.open "select top 1 * from ks_usercard where id=" & KS.chkclng(KS.g("id")),conn,1,3
				else
					if not conn.execute("select cardnum from ks_usercard where cardnum='" & cardnum & "'").eof then
					  call KS.AlertHistory("������ĳ�ֵ�����Ѵ��ڣ�������!",-1):exit sub
					end if
				   rs.open "select top 1 * from ks_usercard",conn,1,3
				   rs.addnew
				   rs("AddDate")=now
				   rs("cardtype")=cardtype
			   end if
				 rs("cardnum")=CardNum
				 rs("cardpass")=KS.Encrypt(PassWord)
				 rs("money")=money
				 rs("ValidNum")=ValidNum
				 rs("ValidUnit")=ValidUnit
				 rs("enddate")=EndDate
				 rs("isused")=isused
				 rs("isSale")=issale
				 rs("groupname")=groupname
				 rs("allowgroupid")=allowgroupid
				 rs("groupid")=groupid
				 rs("expiregroupid")=expiregroupid
			   rs.update
			   rs.close:set rs=nothing
		  else 
		    if CardList="" then call KS.AlertHistory("��û�������ֵ����!",-1):exit sub
			Dim i,j,CardAndPass,CardArr:CardArr=Split(CardList,vbcrlf)
			For I=0 to Ubound(CardArr)
			   CardAndPass=Split(CardArr(I),"|")
			   if not conn.execute("select cardnum from ks_usercard where cardnum='" & CardAndPass(0) & "'").eof then
					 ' call KS.AlertHistory("������ĳ�ֵ�����Ѵ��ڣ�������!",-1):exit sub
			   else
				   Set RS=Server.CreateObject("adodb.recordset")
				   rs.open "select top 1 * from ks_usercard",conn,1,3
				   rs.addnew
					 rs("cardnum")=CardAndPass(0)
					 rs("cardpass")=KS.Encrypt(CardAndPass(1))
					 rs("money")=money
					 rs("ValidNum")=ValidNum
					 rs("ValidUnit")=ValidUnit
					 rs("AddDate")=now
					 rs("enddate")=EndDate
					 rs("isused")=isused
					 rs("isSale")=issale
					 rs("groupname")=groupname
					 rs("cardtype")=cardtype
					 rs("allowgroupid")=allowgroupid
					 rs("groupid")=groupid
					 rs("expiregroupid")=expiregroupid
				   rs.update
				   rs.close:set rs=nothing
			  end if
			Next
		  end if
		  if KS.g("action1")="Edit" then
			   response.write "<script>alert('�޸ĳ�ֵ���ɹ���');location.href='KS.Card.asp?cardtype=" & cardtype & "';</script>"
		  else
			   response.write "<script>alert('��ӳ�ֵ���ɹ���');location.href='KS.Card.asp?cardtype=" & cardtype & "';</script>"
		  end if
		End Sub
		'�������ɳ�ֵ������
		Sub DoAddMore()
		 Dim Nums:Nums=KS.ChkClng(KS.G("Nums"))
		 Dim CardNumPrefix:CardNumPrefix=KS.G("CardNumPrefix")
		 Dim CardNumLen:CardNumLen=KS.ChkClng(KS.G("CardNumLen"))
		 Dim PasswordLen:PasswordLen=KS.ChkClng(KS.G("PasswordLen"))
		 Dim zhtype:zhtype=KS.G("zhtype")
		 Dim Money:Money=KS.ChkClng(KS.g("money"))
		 Dim ValidNum:ValidNum=KS.ChkClng(KS.G("ValidNum"))
		 Dim ValidUnit:ValidUnit=KS.G("ValidUnit")
		 Dim EndDate:EndDate=KS.G("EndDate")
		 Dim GroupName:GroupName=KS.G("GroupName")
		 Dim CardType:CardType=KS.ChkClng(KS.G("CardType"))
		 Dim AllowGroupID:AllowGroupID=KS.G("AllowGroupID")
		 Dim GroupID:GroupID=KS.ChkClng(KS.G("GroupID"))
		 Dim ExpireGroupID:ExpireGroupID=KS.ChkClng(KS.G("ExpireGroupID"))
		 
		 If GroupName="" Then Call KS.AlertHistory("�����ֵ��ȡ������!",-1):exit sub
		 If CardType=0 Then
			 IF Nums=0 Then Call KS.AlertHistory("���ɳ�ֵ���������������0",-1):exit sub
			 IF CardNumLen=0 Then Call KS.AlertHistory("��ֵ�����볤�ȣ��������0",-1):exit sub
			 IF PasswordLen=0 Then Call KS.AlertHistory("��ֵ�����볤�ȣ��������0",-1):exit sub
		 End If
		 IF Not IsNumeric(KS.G("money")) Or KS.G("money")=0 Then Call KS.AlertHistory("��ֵ����ֵ���������0",-1):exit sub
		 IF ValidNum=0 Then Call KS.AlertHistory("��ֵ���������������0",-1):exit sub
		 If Not IsDate(EndDate) Then Call KS.AlertHistory("��ֵ��ֹ���޸�ʽ����ȷ!",-1):exit sub
		 
		 If CardType=1 Then
		  	   Dim RSObj:Set RSObj=Server.CreateObject("adodb.recordset")
			   rsobj.open "select top 1 * from ks_usercard",conn,1,3
			   rsobj.addnew
				 rsobj("cardnum")=""
				 rsobj("cardpass")=""
				 rsobj("money")=KS.G("money")
				 rsobj("ValidNum")=ValidNum
				 rsobj("ValidUnit")=ValidUnit
				 rsobj("AddDate")=now
				 rsobj("enddate")=EndDate
				 rsobj("isused")=0
				 rsobj("isSale")=0
				 rsobj("groupname")=groupname
				 rsobj("groupid")=groupid
				 rsobj("expiregroupid")=expiregroupid
				 rsobj("allowgroupid")=allowgroupid
				 rsobj("cardtype")=1
			   rsobj.update
			   rsobj.close:set rsobj=nothing
			   Response.Write "<script>alert('��ϲ�����߳�ֵ�������ɣ�');location.href='KS.Card.asp?cardtype=1';</script>"

		 Else
		 %>
					   <br>
				  <table width='300'  border='0' align='center' cellpadding='2' cellspacing='1' class='ctable'>
					<tr class='sort'>
					  <td colspan='2' align='center'><strong>�������ɵĵ㿨��Ϣ���£�</strong></td>
					</tr>
					<tr class='tdbg'>
					  <td width='100'>��ֵ�����ƣ�</td>
					  <td><%=GroupName%></td>
					</tr>
					<tr class='tdbg'>
					  <td width='100'>��ֵ��������</td>
					  <td><%=nums%> ��</td>
					</tr>
					<tr class='tdbg'>
					  <td width='100'>��ֵ����ֵ��</td>
					  <td><%=money%> Ԫ</td>
					</tr>
					<tr class='tdbg'>
					  <td width='100'>
					  <% select case ValidUnit
						case 1:response.write "��ֵ��������"
						case 2:response.write "��ֵ����Ч������"
						case 3:response.write "��ֵ����"
						end select
						%></td>
					  <td>
					  <% response.write ValidNum
					  select case validunit
					   case 1:response.write " ��"
					   case 2:response.write " ��"
					   case 3:response.write " Ԫ"
					  end select
					  %>
					  </td>
					</tr>
					<tr class='tdbg'>
					  <td width='100'>��ֵ��ֹ���ڣ�</td>
					  <td><%=enddate%></td>
					</tr>
					
				</table>
				<br>
				<table width='300' border='0' align='center' cellpadding='2' cellspacing='1' class='ctable'>
			  <tr align='center' class='sort'>
				<td  width=150 height='22'><strong> �� �� </strong></td>
				<td  width=150 height='22'><strong> �� �� </strong></td>
			  </tr>
			 <%
			 Dim n,currcard,CurrCardPass
			 For N=1 To Nums
			   CurrCard=KS.MakeRandom(CardNumLen-len(CardNumPrefix))
			   CurrCard=CardNumPrefix & CurrCard
			   If ZhType=2 then
				 CurrCardPass=KS.GetRndPassword(PasswordLen)
			   Else
				 CurrCardPass=KS.MakeRandom(PasswordLen)
			   End If
			   Do While not Conn.execute("select CardNum From KS_UserCard Where CardNum='" & CurrCard & "'").eof 
				   CurrCard=KS.MakeRandom(CardNumLen-len(CardNumPrefix))
				   CurrCard=CardNumPrefix & CurrCard
			   loop
			   
			   response.write "<tr align='center' class='tdbg' onmouseout=""this.className='tdbg'"" onmouseover=""this.className='tdbgmouseover'"">" & vbcrlf
			   response.write "<td height='22'>" & CurrCard & "</td>" & vbcrlf
			   response.write "<td>" & CurrCardPass & "</td>" & vbcrlf
			   response.write "</tr>" & vbcrlf
			   Dim RS:Set RS=Server.CreateObject("adodb.recordset")
			   rs.open "select top 1 * from ks_usercard",conn,1,3
			   rs.addnew
				 rs("cardnum")=CurrCard
				 rs("cardpass")=KS.Encrypt(CurrCardPass)
				 rs("money")=money
				 rs("ValidNum")=ValidNum
				 rs("ValidUnit")=ValidUnit
				 rs("AddDate")=now
				 rs("enddate")=EndDate
				 rs("isused")=0
				 rs("isSale")=0
				 rs("groupname")=groupname
				 rs("groupid")=groupid
				 rs("expiregroupid")=expiregroupid
				 rs("allowgroupid")=allowgroupid
				 rs("cardtype")=0
			   rs.update
			   rs.close:set rs=nothing
			 Next
			  response.write "</table>"
         End If
	End SUb	
End Class
%> 
