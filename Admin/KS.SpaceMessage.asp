<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
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
Set KSCls = New Admin_BlogMessage
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_BlogMessage
        Private KS
		Private Action,i,strClass,sFileName,RS,SQL,maxperpage,CurrentPage,totalPut,TotalPageNum
        Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()
		With Response
					If Not KS.ReturnPowerResult(0, "KSMS10005") Then          '�����Ȩ��
					 Call KS.ReturnErr(1, "")
					 .End
					 End If
		.Write "<script src='../KS_Inc/common.js'></script>"
		.Write "<script language='javascript' src='../ks_inc/jquery.js'></script>"
		.Write "<script language='javascript' src='../ks_inc/kesion.box.js'></script>"
		.Write "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		.Write "<ul id='mt'>"
		.Write "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
		.Write "  <tr>"
		.Write "    <td height=""23"" align=""left"" valign=""top"">"
		.Write "	<td align=""center""><strong>�� �� �� �� �� ��</strong></td>"
		.Write "    </td>"
		.Write "  </tr>"
		.Write "</table>"
		.Write "</ul>"
		End With	
		
		
		maxperpage = 30 '###ÿҳ��ʾ��
		If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
			Response.Write ("�����ϵͳ����!����������")
			Response.End
		End If
		If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
			CurrentPage = CInt(Request("page"))
		Else
			CurrentPage = 1
		End If
		If CInt(CurrentPage) = 0 Then CurrentPage = 1
		totalPut = Conn.Execute("Select Count(ID) from KS_BlogMessage")(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '�õ���ҳ��
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		Select Case KS.G("action")
		 Case "Del"
		  Call BlogMessageDel()
		 Case "Best"
		  Call BlogMessageBest()
		 Case "CancelBest"
		  Call BlogMessageCancelBest()
		 Case "verify" verify
		 Case "unverify" unverify
		 Case Else
		  Call showmain
		End Select
End Sub

Private Sub showmain()
%>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>ѡ��</td>
	<td width="29%" nowrap>���Ա���</td>
	<td width="11%" nowrap>�� ��</td>
	<td width="11%" nowrap>������</td>
	<td width="16%" nowrap>����ʱ��</td>
	<td width="10%" nowrap>�ظ����</td>
	<td width="10%" nowrap>������</td>
  <td width="18%" nowrap>�������</td>
</tr>
<%
	sFileName = "KS.SpaceMessage.asp?"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_BlogMessage order by id desc"
	If DataBaseType = 1 Then
		If CurrentPage > 100 Then
			Rs.Open SQL, Conn, 1, 1
		Else
			Set Rs = Conn.Execute(SQL)
		End If
	Else
		Rs.Open SQL, Conn, 1, 1
	End If
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr class='list'><td height=""25"" align=center colspan=7>û���˷������ԣ�</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<form name=selform method=post action="?">
<input type="hidden" value="Del" name="action" id="action">
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		
%>
<tr height="22" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td class="splittd" align="center"><input type=checkbox name=ID value='<%=Rs("id")%>'></td>
	<td class="splittd"><a href="javascript:void(0)" title="����鿴����" onclick=""><%=Rs("title")%></a></td>
	<td class="splittd" align="center"><%=rs("username")%></td>
	<td class="splittd" align="center"><%=Rs("AnounName")%></td>
	<td class="splittd" align="center"><%=Rs("adddate")%></td>
	<td class="splittd" align="center"><%if not isnull(rs("Replay")) or rs("replay")<>"" then response.write "�ѻظ�" else response.write "<font color=red>δ�ظ�</font>"%></td>
	<td class="splittd" align="center">
	 <%if rs("status")="1" then
	    response.write "<a href='?action=unverify&id=" & rs("id") & "'><font color=blue>�����</font></a>"
	   else
	    response.write "<a href='?action=verify&id=" & rs("id") & "'><font color=red>δ���</font></a>"
	   end if
	 %>
	</td>
	<td class="splittd" align="center"><a href="../space/?<%=rs("username")%>/message/#<%=rs("id")%>" target="_blank">���</a> <a href="?Action=Del&ID=<%=RS("ID")%>" onclick="return(confirm('ȷ��ɾ����������'));">ɾ��</a> </td>
</tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td  class="splittd" height='25' colspan=8>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">ȫѡ
	<input class=Button type="submit" name="Submit2" value=" ɾ��ѡ�е����� " onclick="{if(confirm('�˲��������棬ȷ��Ҫɾ��ѡ�еļ�¼��?')){$('#action').val('Del');this.document.selform.submit();return true;}return false;}">
	<input type="submit" class="button" value=" ������� " onclick="$('#action').val('verify')">
	<input type="submit" class="button" value=" ����ȡ����� " onclick="$('#action').val('unverify')">
	</td>
</tr>

</form>
<tr>
	<td  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" colspan=8 align=right>
	<%
      Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
 	%></td>
</tr>
</table>

<%
End Sub

'ɾ������
Sub BlogMessageDel()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 Conn.execute("Delete From KS_BlogMessage Where ID In("& id & ")")
 Response.Write "<script>alert('ɾ���ɹ���');location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub

Sub Verify()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 Conn.execute("Update KS_BlogMessage set status=1 Where ID In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
Sub UnVerify()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 Conn.execute("Update KS_BlogMessage set status=0 Where ID In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub

End Class
%> 
