<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 5.0
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Main
KSCls.Kesion()
Set KSCls = Nothing

Class Main
        Private KS,Action,PKID
		Private I, totalPut, CurrentPage, SqlStr, RSObj
        Private MaxPerPage
		Private Sub Class_Initialize()
		  MaxPerPage = 20
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub


		Public Sub Kesion()
			If Not KS.ReturnPowerResult(0, "KSMS20014") Then
			  Call KS.ReturnErr(1, "")
			  exit sub
			End If
			PKID=KS.ChkClng(Request("PKID"))
			Action=KS.G("Action")
			Select Case Action
			 Case "verify"
			      Call verify()
			 Case "del"
			      Call del()
			 Case Else
			   Call MainList()
			End Select
	    End Sub
		
		Sub MainList()
			If Request("page") <> "" Then
				  CurrentPage = CInt(Request("page"))
			Else
				  CurrentPage = 1
			End If
			With Response
			.Write "<html>"
			.Write "<head>"
			.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "<script language=""JavaScript"" src=""../ks_inc/Common.js""></script>"
			.Write "<script language=""JavaScript"" src=""../ks_inc/jquery.js""></script>"
			%>

			<%
			.Write "</head>"
			.Write "<body topmargin=""0"" leftmargin=""0"">"
			  .Write "<ul id='menu_top' style='font-weight:bold;text-align:center;padding-top:14px'>"
			  .Write "����PK�۵����"
			  .Write "</ul>"
			

			.Write "<table width=""100%""  border=""0"" cellpadding=""0"" cellspacing=""0"">"
			%>
			<form name='myform' method='Post' action='KS.PKGD.asp'>
		    <input type="hidden" value="del" name="action" id="action">
		    <input type="hidden" value="1" name="v">

			<%
			.Write "  <tr>"			
			.Write "          <td width=""40"" height=""25"" class=""sort"" align=""center"">ѡ��</td>"
			.Write "          <td height=""25"" class=""sort"" align=""center"">�۵�����</td>"
			.Write "          <td class=""sort"" align=""center"">PK����</td>"
			.Write "          <td align=""center"" class=""sort"">�û�</td>"
			.Write "          <td align=""center"" class=""sort"">ʱ��</td>"
			.Write "          <td align=""center"" class=""sort"">�۵�</td>"
			.Write "          <td align=""center"" class=""sort"">״̬</td>"
			.Write "          <td align=""center"" class=""sort"">�������</td>"
			.Write "  </tr>"
			 
			  dim param
			 if PKID<>0 then
			   param=" where a.pkid=" & PKID
			 end if
			 
			 Set RSObj = Server.CreateObject("ADODB.RecordSet")
					   SqlStr = "SELECT a.*,b.title FROM KS_PKGD a inner join KS_PKZT b on a.pkid=b.id" &param&" order by a.ID DESC"
					   RSObj.Open SqlStr, Conn, 1, 1
					 If RSObj.EOF And RSObj.BOF Then
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
			
								   If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RSObj.Move (CurrentPage - 1) * MaxPerPage
									Else
										CurrentPage = 1
									End If
										Call showContent
				End If
				
			.Write "    </td>"
			.Write "  </tr>"
			.Write "</table>"
			.Write "</body>"
			.Write "</html>"
			End With
			End Sub
			Sub showContent()
			   on error resume next
			  With Response
					Do While Not RSObj.EOF
					  .Write "<tr align=""center"" onMouseOut=""this.className='list'"" onMouseOver=""this.className='listmouseover'"" id='u" & RSobj("ID") & "' onClick=""chk_iddiv('" & rsobj("ID") & "')""> "
					  .Write "<td align='center' width=""40"" height=""25"" class=""splittd""><input name=""id"" onClick=""chk_iddiv('" & rsobj("id") & "')"" type='checkbox' id='c" & rsobj("id") & "' value='" & rsobj("id") & "'></td>"
					  .Write "  <td align='left' class='splittd' height='20'>&nbsp;"
					  .Write "    <span style='cursor:default;' title='" & rsobj("content") & "'>" & KS.GotTopic(RSObj("content"), 45) & "</span> </td>"
					  .Write "  <td class='splittd' align='center'><a href='../plus/pk/pk.asp?id=" & rsobj("pkid") & "' target='_blank'>" & ks.gottopic(rsobj("title"),20) & "</a></td>"
					  .Write "  <td class='splittd' align='center'>" 
					  .write rsobj("username")
					  .Write " </td>"
					  .Write "  <td class='splittd' align='center'>" & rsobj("adddate") & "</td>"
					  .Write "  <td class='splittd' align='center'>"
					   if rsobj("role")="1" then
					    .write "<font color=blue>����</font>"
					   elseif rsobj("role")="2" then
					    .write "<font color=green>����</font>"
					   else
					    .write "<font color=red>������</font>"
					   end if
					  .Write "</td>"
					  .Write "  <td class='splittd' align='center'>"
					   if rsobj("status")=1 then
					    .write "<Font color=green>�����</font>"
					   else
					    .write "<Font color=red>δ���</font>"
					   end if
					  .Write "</td>"
					  .Write "  <td class='splittd' align='center'>"
					  if rsobj("status")="1" then
					  .Write "<a href='?action=verify&v=0&id=" & rsobj("id") &"' title='ȡ�����'>ȡ��</a>"
					  else
					  .Write "<a href='?action=verify&v=1&id=" & rsobj("id") &"' title='������'>���</a>"
					  end if
					  .Write" <a href='?action=del&id=" & rsobj("id") & "' onclick=""return(confirm('ȷ��ɾ����?'))"">ɾ��</a></td>"
					  .Write "</tr>"
					 I = I + 1
					  If I >= MaxPerPage Then Exit Do
						   RSObj.MoveNext
					Loop
					  RSObj.Close
					  
					  %>
						  <tr>
						   <td colspan=6>
						   <div style='margin:5px'><b>ѡ��</b><a href='javascript:Select(0)'><font color=#999999>ȫѡ</font></a> - <a href='javascript:Select(1)'><font color=#999999>��ѡ</font></a> - <a href='javascript:Select(2)'><font color=#999999>��ѡ</font></a>
						   <input type="submit" class="button" value="ɾ��ѡ��" onClick="return(confirm('�˲���������,ȷ��ɾ����?'))">
						   <input type="submit" class="button" value="�������" onClick="$('#action').val('verify')">
							</div>
						   </td>
									</form>  
				 <td colspan=5>
					  
					  </td>
					  </tr>
				</table>
				<%

					  .Write "<tr><td height='26' colspan='8' align='right'>"
					 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
				End With
			End Sub
			
		  
		  'ɾ��
		  Sub del()
		  		 Dim K, ZJID
				 ZJID = Trim(KS.G("ID"))
				 if zjid="" then
				   ks.alerthintscript "��ѡ��Ҫɾ�������!"
				 end if
				 ZJID = Split(ZJID, ",")
				 For k = LBound(ZJID) To UBound(ZJID)
					Conn.Execute ("Delete From KS_PKGD Where ID =" & ZJID(k))
				 Next
				 KS.AlertHintScript "��ϲ,ɾ���ɹ�!"
		  End Sub
		  
		  sub verify()
		    dim id
			id=request("id")
			if id="" then
				   ks.alerthintscript "��ѡ��Ҫ��˵����!"
			end if
			conn.execute("update KS_PKGD set status=" & ks.chkclng(request("v")) & " where id in(" & ks.filterids(id) & ")")
			KS.AlertHintScript "��ϲ,�����ɹ�!"
		  end sub
	

End Class
%>
 
