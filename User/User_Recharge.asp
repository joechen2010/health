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
Set KSCls = New User_Recharge
KSCls.Kesion()
Set KSCls = Nothing

Class User_Recharge
        Private KS,KSUser
		Private Sub Class_Initialize()
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
		Call KSUser.InnerLocation("��ֵ����ֵ")
		Response.Write "<div class=""tabs"">"
		Response.Write " <ul>"
		Response.Write " <li><a href=""User_PayOnline.asp"">����֧����ֵ</a></li>"
		Response.Write " <li class='select'><a href=""user_recharge.asp"">��ֵ����ֵ</a></li>"
		Response.Write " <li><a href=""user_exchange.asp?Action=Point"">�һ�" & KS.Setting(45) & "</a></li>"
		Response.Write " <li><a href=""user_exchange.asp?Action=Edays"">�һ���Ч��</a></li>"
		Response.Write " <li><a href=""user_exchange.asp?Action=Money"">" & KS.Setting(45) & "�һ��˻��ʽ�</a></li>"
		Response.Write "</ul>"
		Response.Write "</div>"
		Select Case KS.S("Action")
		 Case "SaveExchangeEdays"
		    Call SaveExchangeEdays()
	     Case Else
		    Call ExchangeEdays()
		End Select
       End Sub
	  
	   
	   Sub ExchangeEdays()
	    %>
	   <script>
	     function Confirm()
		 {
		  if (document.myform.CardNum.value=="")
		  {
		   alert('�������ֵ������!')
		   document.myform.CardNum.focus();
		   return false;
		  }
		  if (document.myform.CardPass.value=="")
		  {
		   alert('�������ֵ������!')
		   document.myform.CardPass.focus();
		   return false;
		  }
		  return true;
		  }
	   </script>
		<FORM name=myform action="User_ReCharge.asp" method="post">
		  <table class=border cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
			<tr class=title>
			  <td align=middle colSpan=2 height=22><B> �� ֵ �� �� ֵ</B></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>�û�����</td>
			  <td><%=KSUser.UserName%></td>
			</tr>
			<tr class=tdbg>
			  <td align=right>�Ʒѷ�ʽ��</td>
			  <td><%if KSUser.ChargeType=1 Then 
		  Response.Write "�۵���</font>�Ʒ��û�"
		  ElseIf KSUser.ChargeType=2 Then
		   Response.Write "��Ч��</font>�Ʒ��û�,����ʱ�䣺" & cdate(KSUser.BeginDate)+KSUser.Edays & ","
		  ElseIf KSUser.ChargeType=3 Then
		   Response.Write "������</font>�Ʒ��û�"
		  End If
		  %>&nbsp;</td>
		    </tr>
			<tr class=tdbg>
			  <td align=right width=120>�ʽ���</td>
			  <td><input type='hidden' value='<%=KSUser.Money%>' name='Premoney'><%=KSUser.Money%> Ԫ</td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>����<%=KS.Setting(45)%>��</td>
			  <td><%=KSUser.Point%>&nbsp;<%=KS.Setting(46)%></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>ʣ��������</td>
			  <td>
			  <%if KSUser.ChargeType=3 Then%>
			  ������
			  <%else%>
			  <%=KSUser.GetEdays%>&nbsp;��
			  <%end if%></td>
			</tr>
			<tr class=tdbg>
			  <td align=right>��ֵ�����ţ�</td>
			  <td>&nbsp;<input name="CardNum" type="text" class="textbox" size="25" maxlength="50"></td>
		    </tr>
			<tr class=tdbg>
			  <td align=right width=120>��ֵ�����룺</td>
			  <td>&nbsp;<input name="CardPass" type="text" class="textbox" size="25" maxlength="50"></td>
			</tr>
			<tr class=tdbg>
			  <td align=middle colSpan=2 height=40>
		        <Input id=Action type=hidden value="SaveExchangeEdays" name="Action"> 
				<Input class="button" id=Submit type=submit value="ȷ����ֵ" onClick="return(Confirm())" name=Submit> </td>
			</tr>
		  </table>
		</FORM>
	   <%
	   End Sub
		
	   Sub SaveExchangeEdays()
	   	 Dim ChangeType:ChangeType=KS.S("ChangeType")
		 Dim Money:Money=KS.S("Money")
		 DiM CardNum:CardNum=KS.S("CardNum")
		 Dim CardPass:CardPass=KS.S("CardPass")
		 If CardNum="" Or CardPass="" Then 
		   Call KS.AlertHistory("������ĳ�ֵ���ż����룡",-1)
		   exit sub
		 end if
		 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		 rs.open "select top 1 * from ks_usercard where cardtype=0 and cardnum='" & CardNum & "'",conn,1,1
		 if rs.bof and rs.eof then
		  rs.close:set rs=nothing
		  Call KS.AlertHistory("�Բ���������ĳ�ֵ���Ų���ȷ��",-1)
		  exit sub
		 end if
		 if rs("cardpass")<>KS.Encrypt(cardpass) then
		  rs.close:set rs=nothing
		  Call KS.AlertHistory("�Բ���������ĳ�ֵ�����벻��ȷ��",-1)
		  exit sub
		 end if
		 
		 if rs("isused")=1 then
		  rs.close:set rs=nothing
		  Call KS.AlertHistory("�Բ���������ĳ�ֵ���ѱ�ʹ�ã�",-1)
		  exit sub
		 end if
		 
		 if datediff("d",rs("enddate"),now())>0 then
		  rs.close:set rs=nothing
		  Call KS.AlertHistory("�Բ���������ĳ�ֵ���ѹ��ڣ�",-1)
		  exit sub
		 end if
		 
		 if not KS.IsNul(rs("allowgroupid")) then
		    If KS.FoundInArr(rs("allowGroupID"),KSUser.GroupID,",")=false Then
			  rs.close:set rs=nothing
			  Call KS.AlertHistory("�Բ��������ڵ��û���û��ʹ�ñ���ֵ����Ȩ��,����ϵ��վ����Ա��",-1)
			  exit sub
			End If
		 end if
		 
		  Dim ValidNum:ValidNum=rs("ValidNum")
		  Dim ValidUnit:ValidUnit=rs("ValidUnit")
		  Dim UserCardID:UserCardID=rs("id")
		  Dim GroupID:GroupID=rs("GroupID")
		  rs.close
		  rs.open "select top 1 * from ks_user Where UserName='" & KSUser.UserName & "'",conn,1,1
		  if not rs.eof then
		    if rs("ChargeType")=3 and ValidUnit<>3 then
				  rs.close:set rs=nothing
				  Call KS.AlertHistory("��������˻��������ڣ������ֵ�ʽ��빺���ʽ𿨣�",-1)
				  exit sub
			end if
			dim ValidDays,tmpdays
		    select case ValidUnit
			  case 1 '����
			   'rs("point")=rs("point")+ValidNum
			   Call KS.PointInOrOut(0,0,rs("UserName"),1,ValidNum,"System","ͨ����ֵ����õĵ���",0)
			  case 2 '����
			    ValidDays=rs("Edays")
				tmpDays=ValidDays-DateDiff("D",rs("BeginDate"),now())
				if tmpDays>0 then
				    conn.execute("update ks_user set edays=edays+" & validnum & " where username='" & ksuser.username & "'")
				else
					conn.execute("update ks_user set begindate=" & sqlnowstring & ",edays=" & validnum & " where username='" & ksuser.username & "'")
				end if
				Call KS.EdaysInOrOut(rs("UserName"),1,ValidNum,"System","ͨ����ֵ��[" & CardNum & "]��õ���Ч����")
			  case 3 '���
			    Call KS.MoneyInOrOut(rs("UserName"),RS("RealName"),ValidNum,4,1,now,0,"System","ͨ����ֵ��[" & CardNum & "]��õ��ʽ�",0,0)
			  case 4 '����
			    Call KS.ScoreInOrOut(rs("UserName"),1,ValidNum,"System","ͨ����ֵ��[" & CardNum & "]��õĻ���!",0,0)
			end select
			if GroupID<>0 then conn.execute("update ks_user set groupid=" & GroupID & " where userName='" & KSUser.UserName & "'")
			conn.execute("update ks_user set usercardid="&usercardid &" where userName='" & KSUser.UserName & "'")
		  end if
		  '�ó�ֵ����ʹ�á����۳�
		  Conn.Execute("Update KS_UserCard Set Isused=1,issale=1,username='" & KSUser.UserName & "',UseDate=" & SqlNowString & " where cardnum='" & cardnum & "'")
		 
		 if GroupID<>0 then
		 Response.Write "<script>alert('��ϲ������ֵ�ɹ�������Ϊ"""& KS.U_G(GroupID,"groupname") &"""!');location.href='user_recharge.asp';</script>"
		 else
		 Response.Write "<script>alert('��ϲ������ֵ�ɹ�!');location.href='user_recharge.asp';</script>"
		 end if
		 RS.Close:Set RS=Nothing
	   End Sub
End Class
%> 
