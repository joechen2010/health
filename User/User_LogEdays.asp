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
Set KSCls = New User_LogEdays
KSCls.Kesion()
Set KSCls = Nothing

Class User_LogEdays
        Private KS,KSUser
		Private CurrentPage,totalPut,TotalPages,SQL
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
		Call KSUser.InnerLocation("��ѯ�ҵ���Ч����ϸ")
			   		       If KS.S("page") <> "" Then
						          CurrentPage = CInt(KS.S("page"))
							Else
								  CurrentPage = 1
							End If
								   %>
		<div class="tabs">	
			<ul>
				<li><a href="user_logmoney.asp">�ʽ���ϸ</a></li>
				<li><a href="user_logpoint.asp">��ȯ��ϸ</a></li>
				<li class="select"><a href="user_logedays.asp">��Ч����ϸ</a></li>
				<li><a href="user_logscore.asp">������ϸ</a></li>
			</ul>
		</div>
			  <div style="text-align:right"> <a href='User_LogEdays.asp'><font color=red>�����м�¼</font></a> ��<a href='?InOrOutFlag=1'>�����¼[<%=conn.execute("select count(id) from ks_logEdays where InOrOutFlag=1 and username='" & KSUser.UserName & "'")(0)%>]</a> ��<a href='?InOrOutFlag=2'>֧����¼[<%=conn.execute("select count(id) from ks_logEdays where InOrOutFlag=2 and username='" & KSUser.UserName & "'")(0)%>]</a>
			 </div>
				<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
					<tr class="title">
					<td width="80" height="25" align="center"><strong> �û���</strong></td>
					<td width="180" height="25" align="center"><strong>����ʱ��</strong></td>
					<td width="113" align="center"><strong>IP��ַ</strong></td>
					<td width="71" height="25" align="center"><strong>��������</strong></td>
					<td width="74" align="center"><strong>֧������</strong></td>
					<td width="55" height="25" align="center"><strong>ժҪ</strong></td>
					<td width="75" height="25" align="center"><strong> ����Ա</strong></td>
					<td width="239" align="center"><strong>��ע</strong></td>
				  </tr>
					<%  If KS.ChkClng(KS.S("InOrOutFlag"))=1 Or KS.ChkClng(KS.S("InOrOutFlag"))=2 Then
						  SqlStr="Select ID,UserName,AddDate,IP,Edays,InOrOutFlag,User,Descript From KS_LogEdays Where InOrOutFlag=" & KS.ChkClng(KS.S("InOrOutFlag")) & " And  UserName='" & KSUser.UserName &"' order by id desc"
 					    Else
						  SqlStr="Select ID,UserName,AddDate,IP,Edays,InOrOutFlag,User,Descript From KS_LogEdays Where UserName='" & KSUser.UserName &"' order by id desc"
						End if
						 Set RS=Server.createobject("adodb.recordset")
						 RS.open SqlStr,conn,1,1

						 If RS.EOF And RS.BOF Then
								  Response.Write "<tr class='tdbg'><td align=center height=25 colspan=9 valign=top>�Ҳ�����Ҫ�ļ�¼!</td></tr>"
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
		  </td>
		  </tr>
</table>
		 <%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
		  <%
  End Sub
    
  Sub ShowContent
 Dim i,InEdays,OutEdays
For i=0 To Ubound(SQL,2)
	%>
  <tr height="25" class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td class="splittd" width="80" align="center"><%=SQL(1,i)%></td>
    <td class="splittd" align="center"><%=SQL(2,i)%></td>
    <td class="splittd"align="center"><%=SQL(3,i)%></td>
    <td class="splittd" align="right"><%if SQL(5,I)=1 Then Response.Write SQL(4,I) & "��":InEdays=InEdays+SQL(4,I) ELSE Response.Write "-"%></td>
    <td class="splittd" align="right"><%if SQL(5,I)=2 Then Response.Write SQL(4,I) & "��":OutEdays=OutEdays+SQL(4,I) ELSE Response.Write "-"%></td>
    <td class="splittd" align="center"><%if SQL(5,I)=1 Then Response.Write "<font color=red>����</font>" Else Response.Write "֧��"%></td>
    <td class="splittd" align="center"><%=SQL(6,i)%></td>
	<td class="splittd"><%=SQL(7,i)%></td>
  </tr>
  <%Next%>
  <tr class='tdbg' onMouseOut="this.className='tdbg'" onMouseOver="this.className='tdbgmouseover'">    <td class="splittd" colspan='3' align='right'>��ҳ�ϼƣ�</td>    <td class="splittd" align='right'><%=InEdays%>��</td>    <td class="splittd" align='right'><%=KS.ChkClng(OutEdays)%>��</td>    <td class="splittd" colspan='4'>&nbsp;</td>  </tr> 
  <% Dim totalinEdays:totalinEdays=conn.execute("Select sum(Edays) From KS_LogEdays where username='" & KSUser.UserName & "'AND InOrOutFlag=1")(0)
     Dim TotalOutEdays:TotalOutEdays=conn.execute("Select sum(Edays) From KS_LogEdays where username='" & KSUser.UserName & "'AND  InOrOutFlag=2")(0)
	 If KS.ChkClng(totalInEdays)=0 Then totalInEdays=0
	 If KS.ChkClng(TotalOutEdays)=0 Then TotalOutEdays=0
  %>
    <tr class='tdbg' onMouseOut="this.className='tdbg'" onMouseOver="this.className='tdbgmouseover'">    <td class="splittd" colspan='3' align='right'>���кϼƣ�</td>    <td class="splittd" align='right'><%=KS.ChkClng(totalInEdays)%>��</td>    <td align='right'><%=KS.ChkClng(totalOutEdays)%>��</td>    <td class="splittd" colspan='4' align='center'>�ۼƻ�ʣ��<%=totalInEdays-totalOutEdays%>��</td>  </tr> 

  <%  

End Sub
  
End Class
%> 
