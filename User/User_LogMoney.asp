<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New User_LogMoney
KSCls.Kesion()
Set KSCls = Nothing

Class User_LogMoney
        Private KS,KSUser
		Private CurrentPage,totalPut
		Private RS,MaxPerPage
		Private TempStr,SqlStr
		Private Sub Class_Initialize()
			MaxPerPage =20
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "<script>top.location.href='Login';</script>"
		  Exit Sub
		End If
		Call KSUser.Head()
		Call KSUser.InnerLocation("��ѯ�ҵ��ʽ���ϸ")
		 If KS.S("page") <> "" Then
						          CurrentPage = CInt(KS.S("page"))
							Else
								  CurrentPage = 1
							End If
							%>
		<div class="tabs">	
			<ul>
				<li class="select"><a href="user_logmoney.asp">�ʽ���ϸ</a></li>
				<li><a href="user_logpoint.asp">��ȯ��ϸ</a></li>
				<li><a href="user_logedays.asp">��Ч����ϸ</a></li>
				<li><a href="user_logscore.asp">������ϸ</a></li>
			</ul>
		</div>
			<div style="text-align:right"> <a href='User_LogMoney.asp'><font color=red>�����м�¼</font></a> ��<a href='?IncomeOrPayOut=1'>�����¼[<%=conn.execute("select count(id) from ks_logmoney where IncomeOrPayOut=1 and username='" & KSUser.UserName & "'")(0)%>]</a> ��<a href='?IncomeOrPayOut=2'>֧����¼[<%=conn.execute("select count(id) from ks_logmoney where IncomeOrPayOut=2 and username='" & KSUser.UserName & "'")(0)%>]</a>
		   </div>
				<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
					<tr class=title align=middle>
					  <td width=150 height="25">����ʱ��</td>
					  <td width=80>�û���</td>
					  <td width=80>�ͻ�����</td>
					  <td width=60>���׷�ʽ</td>
					  <td width=50>����</td>
					  <td width=60>������</td>
					  <td width=60>֧�����</td>
					  <td width=40>ժҪ</td>
					  <td width=40>���</td>
					  <td>��ע/˵��</td>
					</tr>
					<%  If KS.ChkClng(KS.S("IncomeOrPayOut"))=1 Or KS.ChkClng(KS.S("IncomeOrPayOut"))=2 Then
						  SqlStr="Select * From KS_LogMoney Where IncomeOrPayOut=" & KS.ChkClng(KS.S("IncomeOrPayOut")) & " And  UserName='" & KSUser.UserName &"' order by id desc"
 					    Else
						  SqlStr="Select * From KS_LogMoney Where UserName='" & KSUser.UserName &"' order by id desc"
						End if
						 Set RS=Server.createobject("adodb.recordset")
						 RS.open SqlStr,conn,1,1

						 If RS.EOF And RS.BOF Then
								  Response.Write "<tr class='tdbg'><td align=center height=25 colspan=9 valign=top>�Ҳ�����Ҫ�ļ�¼!</td></tr>"
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
			
								If CurrentPage = 1 Then
									Call ShowContent
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
										Call ShowContent
									Else
										CurrentPage = 1
										Call ShowContent
									End If
								End If
				End If

						
						 %>
					
          </table>
		  </td>
		  </tr>
</table>
		  <%
		  Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
		  <%
  End Sub
    
  Sub ShowContent()
     on error resume next
     Dim I,intotalmoney,outtotalmoney
     Do While Not rs.eof 
	%>
    <tr class=tdbg onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
      <td  class="splittd" align=middle><%=rs("LogTime")%></td>
      <td  class="splittd" align=middle width=80><%=rs("username")%></td>
	  <td  class="splittd" align=middle width=80><%=rs("clientname")%></td>
      <td   class="splittd" align=middle width=60>
	  <% Select Case rs("MoneyType")
	      Case 1:Response.WRite "�ֽ�"
		  Case 2:Response.Write "���л��"
		  Case 3:Response.Write "����֧��"
		  Case 4:Response.Write "�ʽ����"
		 End Select
	 %>
	  </td>
      <td  class="splittd" align=middle width=50>�����</td>
      <td  class="splittd" align=right>&nbsp; 
	  <%If rs("IncomeOrPayOut")=1 Then
	     Response.Write formatnumber(rs("money"),2,-1)
		 intotalmoney=intotalmoney+rs("money")
	    End If
		%></td>
      <td  class="splittd" align=right>&nbsp;
	  <%If rs("IncomeOrPayOut")=2 Then
	     Response.Write formatnumber(rs("money"),2,-1)
		 outtotalmoney=outtotalmoney+rs("money")
	    End If
		%></td>
      <td  class="splittd" align=center width=40>
	  <% If rs("IncomeOrPayOut")=1 Then
	      Response.Write "<font color=red>����</font>"
		 Else
		  Response.Write "<font color=green>֧��</font>"
		 End If
		 %>
		 </td>
      <td  class="splittd" align=center width=40>
	  <%=formatnumber(RS("CurrMoney"),2,-1)%>
		 </td>
      <td  class="splittd" align=middle><%=rs("Remark")%></td>
    </tr>
	<%
	            
				I = I + 1
				RS.MoveNext
				If I >= MaxPerPage Then Exit Do

	 loop
	%>
    <tr class=tdbg onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
      <td class="splittd"  align=right colSpan=5>��ҳ�ϼƣ�</td>
      <td class="splittd" align=right><%=formatnumber(intotalmoney,2,-1)%></td>
      <td class="splittd" align=right><%=formatnumber(outtotalmoney,2,-1)%></td>
      <td class="splittd" colSpan=3>&nbsp;</td>
    </tr>
    <tr class=tdbg onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
      <td class="splittd" align=right colSpan=5>�ܼƽ�</td>
	  	  <%intotalmoney=Conn.execute("Select Sum(Money) From KS_Logmoney Where UserName='" & KSUser.UserName & "' And IncomeOrPayOut=1")(0)
	    outtotalmoney=Conn.execute("Select Sum(Money) From KS_Logmoney Where UserName='" & KSUser.UserName & "' And IncomeOrPayOut=2")(0)
	    if not isnumeric(intotalmoney) then intotalmoney=0
		if not isnumeric(outtotalmoney) then outtotalmoney=0
	  %>
      <td class="splittd" align=right><%=formatnumber(intotalmoney,2,-1)%></td>
      <td class="splittd" align=right><%=formatnumber(outtotalmoney,2,-1)%></td>
      <td class="splittd" align=middle colSpan=3>�ʽ���<%=formatnumber(KSUser.Money,2,-1)%></td>

    </tr>
  </table>
		<%
		End Sub
  
End Class
%> 
