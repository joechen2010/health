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
Set KSCls = New User_ExChange
KSCls.Kesion()
Set KSCls = Nothing

Class User_ExChange
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
		Response.Write "<div class=""tabs"">"
		Response.Write " <ul class="""">"
		Response.Write " <li><a href=""User_PayOnline.asp"">����֧����ֵ</a></li>"
		Response.Write " <li><a href=""user_recharge.asp"">��ֵ����ֵ</a></li>"
		If KS.S("Action")="Point" Then
		Response.Write " <li class='select'><a href=""user_exchange.asp?Action=Point"">�һ�" & KS.Setting(45) & "</a></li>"
		Else
		Response.Write " <li><a href=""user_exchange.asp?Action=Point"">�һ�" & KS.Setting(45) & "</a></li>"
		End IF
		If KS.S("Action")="Edays" Then
		Response.Write " <li class='select'><a href=""user_exchange.asp?Action=Edays"">�һ���Ч��</a></li>"
		Else
		Response.Write " <li><a href=""user_exchange.asp?Action=Edays"">�һ���Ч��</a></li>"
		End If
		If KS.S("Action")="Money" Then
		Response.Write " <li class='select'><a href=""user_exchange.asp?Action=Money"">" & KS.Setting(45) & "�һ��˻��ʽ�</a></li>"
		Else
		Response.Write " <li><a href=""user_exchange.asp?Action=Money"">" & KS.Setting(45) & "�һ��˻��ʽ�</a></li>"
		End If
		
		Response.Write "</ul>"
		Response.Write "</div>"
		Select Case KS.S("Action")
		 Case "Point"
		   Call KSUser.InnerLocation("�һ�" & KS.Setting(45))
		   Call ExchangePoint()
		 Case "Money" 
		   Call KSUser.InnerLocation("�һ��˻��ʽ�")
		   Call ExchangeMoney()
		 Case "SaveExchangeMoney"
		   Call SaveExchangeMoney()
		 Case "SaveExchangePoint"
		   Call SaveExchangePoint()
		 Case "Edays"
		 	Call KSUser.InnerLocation("�һ���Ч����")
		    Call ExchangeEdays()
		 Case "SaveExchangeEdays"
		    Call SaveExchangeEdays()
		End Select
       End Sub
	   
	   Sub ExchangePoint()
	   %>
	   <script>
	     function Confirm()
		 {
		   var str='��������:\n';
		   if (document.myform.ChangeType[0].checked==true){
		     if (parseInt(document.myform.Premoney.value)<parseInt(document.myform.Money.value)){
			   alert('��Ŀǰ�ʽ����㣬���ֵ�������һ���');
			   return false;
			   }
		    str+='�һ�ǰ�ʽ�'+document.myform.Premoney.value+' Ԫ\n';
			str+='�һ����ʽ�'+(document.myform.Premoney.value-document.myform.Money.value)+' Ԫ\n';
			str+='һ���һ��ɹ��������棬ȷ���һ���';
		   }else{
		   if (parseInt(document.myform.PreScore.value)<parseInt(document.myform.Score.value)){
			   alert('��Ŀǰ���û��ֲ��㣬���ܶһ���');
			   return false;
			   }
		    str+='�һ�ǰ���֣�'+document.myform.PreScore.value+' ��\n';
			str+='�һ�����֣�'+(document.myform.PreScore.value-document.myform.Score.value)+' ��\n';
			str+='һ���һ��ɹ��������棬ȷ���һ���';
		   }
		   if (confirm(str)){
		    return true}
		   else{ return false}
		 }
	   </script>
		<FORM name=myform action="User_Exchange.asp" method="post">
		  <table class=border cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
			<tr class=title>
			  <td align=middle colSpan=2 height=22><B> �� �� �� ȯ </B></td>
			</tr> 
			<tr class=tdbg>
			  <td align=right width=120>�û�����</td>
			  <td><%=KSUser.UserName%></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>�ʽ���</td>
			  <td><input type='hidden' value='<%=KSUser.Money%>' name='Premoney'><%=KSUser.Money%> Ԫ</td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>���û��֣�</td>
			  <td><input type='hidden' value='<%=KSUser.Score%>' name='PreScore'><%=KSUser.Score%> ��</td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>����<%=KS.Setting(45)%>��</td>
			  <td><%=KSUser.Point%>&nbsp;<%=KS.Setting(46)%></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>�һ���ȯ��</td>
			  <td>
		  <Input type=radio CHECKED value="1" name="ChangeType">ʹ���ʽ��� �� 
		  <Input style="TEXT-ALIGN: center" maxLength=8 size=6 value=100 name="Money"> Ԫ�һ���<%=KS.Setting(45)%> &nbsp;&nbsp;&nbsp;&nbsp;<Font color=red>�һ����ʣ�<%=KS.Setting(43)%>Ԫ:1<%=KS.Setting(46)%></Font> <br>
		  <Input type=radio value="2" name="ChangeType">ʹ�þ�����֣� �� 
				<Input style="TEXT-ALIGN: center" maxLength=8 size=6 value=100 name="Score"> �ֶһ���<%=KS.Setting(45)%> &nbsp;&nbsp;&nbsp;&nbsp;<Font color=red>�һ����ʣ�<%=KS.Setting(41)%>��:1<%=KS.Setting(46)%> </Font></td>
			</tr>
			<tr class=tdbg>
			  <td align=middle colSpan=2 height=40>
		        <Input id=Action type=hidden value="SaveExchangePoint" name="Action"> 
				<Input class="button" id=Submit type=submit value="ִ�жһ�" onClick="return(Confirm())" name=Submit> </td>
			</tr>
		  </table>
		</FORM>
	   <%
	   End Sub
	   
	   Sub SaveExchangePoint()
	     Dim ChangeType:ChangeType=KS.S("ChangeType")
		 Dim Money:Money=KS.S("Money")
		 Dim Score:Score=KS.ChkClng(KS.S("Score"))
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select * From KS_User Where UserName='" & KSUser.UserName & "'",conn,1,1
		 If RS.Eof  Then
		   Rs.Close:Set RS=Nothing
		   Response.Write "<script>alert('������!');history.back();</script>"
		   Exit Sub
		 End If
		 If ChangeType=1 Then
		    If KS.ChkClng(Money)=0 Then
			   Rs.Close:Set RS=Nothing
			   Response.Write "<script>alert('��������ʽ���ȷ,�ʽ�������0!');history.back();</script>"
			   Exit Sub
			End iF
			If KS.ChkClng(Money)<KS.ChkClng(KS.Setting(43)) Then
			   Rs.Close:Set RS=Nothing
			   Response.Write "<script>alert('��������ʽ���ȷ,�ʽ������ڵ���" & KS.Setting(43) &"!');history.back();</script>"
			   Exit Sub
			End If
		   IF Round(RS("Money"))<Round(Money) Then
			   Rs.Close:Set RS=Nothing
			   Response.Write "<script>alert('������ʽ��㣬���ֵ�������һ�!');history.back();</script>"
			   Exit Sub
		   End If

			'  ChannelID,InfoID,UserName,InOrOutFlag,Point,User,Descript
			Call KS.PointInOrOut(0,0,rs("UserName"),1,Money/KS.Setting(43),"System","�˻��ʽ�һ�����",0)
			Call KS.MoneyInOrOut(rs("UserName"),rs("RealName"),Money,4,2,now,0,"System","���ڶһ���ȯ",0,0)
	
		 Else
		    If Score=0 Then
			   Rs.Close:Set RS=Nothing
			   Response.Write "<script>alert('������Ļ��ֲ���ȷ,���ֱ������0!');history.back();</script>"
			   Exit Sub
			End If
			If KS.ChkClng(Score)<KS.ChkClng(KS.Setting(41)) Then
			   Rs.Close:Set RS=Nothing
			   Response.Write "<script>alert('������Ļ��ֲ���ȷ,���ֱ�����ڵ���" & KS.Setting(41) &"!');history.back();</script>"
			   Exit Sub
			End If
		   IF KS.ChkClng(RS("Score"))<KS.ChkClng(Score) Then
			   Rs.Close:Set RS=Nothing
			   Response.Write "<script>alert('����û��ֲ��㣬���ܶһ�!');history.back();</script>"
			   Exit Sub
		   End If
		   
		     call KS.ScoreInOrOut(rs("UserName"),2,Score,"System","�һ���ȯ����!",0,0)
			'ChannelID,InfoID,UserName,InOrOutFlag,Point,User,Descript
			 Call KS.PointInOrOut(0,0,rs("UserName"),1,Score/KS.Setting(41),"System","���ֶһ�����",0)
		 End IF
		 Response.Write "<script>alert('��ϲ������ȯ�һ��ɹ�!');location.href='User_ExChange.asp?Action=Point';</script>"
		 RS.Close:Set RS=Nothing
 	   End Sub
	   
	   
	   Sub ExchangeMoney()
	   %>
	   		<script>
	     function checkform()
		 {
		   var str='��������:\n';
		     if (parseInt(document.myforms.Prepoint.value)<parseInt(document.myforms.Point.value)){
			   alert('�Բ������<%=KS.Setting(45)%>���㣡');
			   return false;
			   }
		    str+='�һ�ǰ<%=KS.Setting(45)%>��'+document.myforms.Prepoint.value+' <%=KS.Setting(46)%>\n';
			str+='�һ���<%=KS.Setting(45)%>��'+(document.myforms.Prepoint.value-document.myforms.Point.value)+' <%=KS.Setting(46)%>\n';
			str+='һ���һ��ɹ��������棬ȷ���һ���';

		   if (confirm(str)){
		    return true}
		   else{ return false}
		 }
	   </script>
		<FORM name=myforms action="User_Exchange.asp" method="post">
		  <table class=border cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
			<tr class=title>
			  <td align=middle colSpan=2 height=22><B> �� �� �� �� </B></td>
			</tr> 
			<tr class=tdbg>
			  <td align=right width=120>�û�����</td>
			  <td><%=KSUser.UserName%></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>�ʽ���</td>
			  <td><%=KSUser.Money%> Ԫ</td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>���û��֣�</td>
			  <td><%=KSUser.Score%> ��</td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>����<%=KS.Setting(45)%>��</td>
			  <td><%=formatnumber(KSUser.Point,2)%>&nbsp;<%=KS.Setting(46)%><input type='hidden' value='<%=KSUser.Point%>' name='Prepoint'></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>�һ��ʽ�</td>
			  <td>
		   ��
		  <Input style="TEXT-ALIGN: center" maxLength=8 size=6 value=<%=KS.ChkClng(KSUser.Point)%> name="Point"> <%=KS.Setting(46)%><%=KS.Setting(45)%>�һ����˻��ʽ� &nbsp;&nbsp;&nbsp;&nbsp;<Font color=red>�һ����ʣ�1<%=KS.Setting(46)%>:<%=KS.Setting(43)%>Ԫ</Font> <br>
		  </td>
			</tr>
			<tr class=tdbg>
			  <td align=middle colSpan=2 height=40>
		        <Input id=Action type=hidden value="SaveExchangeMoney" name="Action"> 
				<Input class="button" id=Submit type=submit value="ִ�жһ�" onClick="return(checkform())" name=Submit> </td>
			</tr>
		  </table>
		</FORM>
		<div style="padding-left:60px;color:green">˵���������Խ�Ͷ���õ�<%=KS.Setting(45)%>�һ����˻��ʽ��������ڱ�վ�̳ǽ������ѡ�</div>
		
	   <%
	   End Sub
	   
	   Sub SaveExchangeMoney()
		 Dim Point:Point=KS.S("Point")

		    If Round(Point)<=0 Then
			   Response.Write "<script>alert('�������" & KS.Setting(45) & "����ȷ,�������0!');history.back();</script>"
			   Exit Sub
			End iF
			
			
		DIM RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_User Where UserName='" & KSUser.UserName & "'",conn,1,1
		 If RS.Eof  Then
		   Rs.Close:Set RS=Nothing
		   Response.Write "<script>alert('������!');history.back();</script>"
		   Exit Sub
		 End If
		   IF Round(RS("Point"))<Round(Point) Then
			   Rs.Close:Set RS=Nothing
			   Response.Write "<script>alert('�����" & KS.Setting(45) & "����!');history.back();</script>"
			   Exit Sub
		   End If
			'  ChannelID,InfoID,UserName,InOrOutFlag,Point,User,Descript
			 Call KS.PointInOrOut(0,0,rs("UserName"),2,Round(Point),"System","�û��һ��˻��ʽ�",0)
			 Call KS.MoneyInOrOut(rs("UserName"),rs("RealName"),(point*KS.Setting(43)),4,1,now,0,"System","��ȯ�һ�����",0,0)
		 Response.Write "<script>alert('��ϲ�����˻��ʽ�һ��ɹ�!');location.href='User_ExChange.asp?Action=Money';</script>"
		 RS.Close:Set RS=Nothing
	   End Sub
	   
	   	   
	   Sub ExchangeEdays()
	    %>
	   <script>
	     function Confirm()
		 {
		   var str='��������:\n';
		   if (document.myform.ChangeType[0].checked==true){
		     if (parseInt(document.myform.Premoney.value)<parseInt(document.myform.Money.value)){
			   alert('��Ŀǰ�ʽ����㣬���ֵ�������һ���');
			   return false;
			   }
		    str+='�һ�ǰ�ʽ�'+document.myform.Premoney.value+' Ԫ\n';
			str+='�һ����ʽ�'+(document.myform.Premoney.value-document.myform.Money.value)+' Ԫ\n';
			str+='һ���һ��ɹ��������棬ȷ���һ���';
		   }else{
		   if (parseInt(document.myform.PreScore.value)<parseInt(document.myform.Score.value)){
			   alert('��Ŀǰ���û��ֲ��㣬���ܶһ���');
			   return false;
			   }
		    str+='�һ�ǰ���֣�'+document.myform.PreScore.value+' ��\n';
			str+='�һ�����֣�'+(document.myform.PreScore.value-document.myform.Score.value)+' ��\n';
			str+='һ���һ��ɹ��������棬ȷ���һ���';
		   }
		   if (confirm(str)){
		    return true}
		   else{ return false}
		 }
	   </script>
		<FORM name=myform action="User_Exchange.asp" method="post">
		  <table class=border cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
			<tr class=title>
			  <td align=middle colSpan=2 height=22><B> �� �� �� Ч ��</B></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>�û�����</td>
			  <td><%=KSUser.UserName%></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>�ʽ���</td>
			  <td><input type='hidden' value='<%=KSUser.Money%>' name='Premoney'><%=KSUser.Money%> Ԫ</td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>���û��֣�</td>
			  <td><input type='hidden' value='<%=KSUser.Score%>' name='PreScore'><%=KSUser.Score%> ��</td>
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
			  <td align=right width=120>�һ���ȯ��</td>
			  <td>
		  <Input type=radio CHECKED value="1" name="ChangeType">ʹ���ʽ��� �� 
		  <Input style="TEXT-ALIGN: center" maxLength=8 size=6 value=100 name="Money"> Ԫ�һ�����Ч���� &nbsp;&nbsp;&nbsp;&nbsp;<Font color=red>�һ����ʣ�<%=KS.Setting(44)%>Ԫ:1��</Font> <br>
		  <Input type=radio value="2" name="ChangeType">ʹ�þ�����֣� �� 
				<Input style="TEXT-ALIGN: center" maxLength=8 size=6 value=100 name="Score"> �ֶһ�����Ч���� &nbsp;&nbsp;&nbsp;&nbsp;<Font color=red>�һ����ʣ�<%=KS.Setting(42)%>��:1�� </Font></td>
			</tr>
			<tr class=tdbg>
			  <td align=middle colSpan=2 height=40>
		        <Input id=Action type=hidden value="SaveExchangeEdays" name="Action"> 
				<Input class="button" id=Submit type=submit value="ִ�жһ�" onClick="return(Confirm())" name=Submit> </td>
			</tr>
		  </table>
		</FORM>
	   <%
	   End Sub
		
	   Sub SaveExchangeEdays()
	   	 Dim ChangeType:ChangeType=KS.S("ChangeType")
		 Dim Money:Money=KS.S("Money")
		 Dim Score:Score=KS.ChkClng(KS.S("Score"))
		 Dim tmpDays,ValidDays,RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select * From KS_User Where UserName='" & KSUser.UserName & "'",conn,1,3
		 If RS.Eof  Then
		   Rs.Close:Set RS=Nothing
		   Response.Write "<script>alert('������!');history.back();</script>"
		   Exit Sub
		 End If
		 If ChangeType=1 Then
		    If KS.ChkClng(Money)=0 Then
			   Rs.Close:Set RS=Nothing
			   Response.Write "<script>alert('��������ʽ���ȷ,�ʽ�������0!');history.back();</script>"
			   Exit Sub
			End iF
			If KS.ChkClng(Money)<KS.ChkClng(KS.Setting(44)) Then
			   Rs.Close:Set RS=Nothing
			   Response.Write "<script>alert('��������ʽ���ȷ,�ʽ������ڵ���" & KS.Setting(44) &"!');history.back();</script>"
			   Exit Sub
			End If
		   IF Round(RS("Money"))<Round(Money) Then
			   Rs.Close:Set RS=Nothing
			   Response.Write "<script>alert('������ʽ��㣬���ֵ�������һ�!');history.back();</script>"
			   Exit Sub
		   End If
			    ValidDays=rs("Edays")
				tmpDays=ValidDays-DateDiff("D",rs("BeginDate"),now())
				if tmpDays>0 then
					rs("Edays")=rs("Edays")+Money/KS.Setting(44)
				else
					rs("BeginDate")=now
					rs("Edays")=Money/KS.Setting(44)
				end if
			RS.Update
			'  UserName,InOrOutFlag,Edays,User,Descript
			Call KS.EdaysInOrOut(rs("UserName"),1,Money/KS.Setting(44),"System","�˻��ʽ�һ�����")
			Call KS.MoneyInOrOut(rs("UserName"),rs("RealName"),Money,4,2,now,0,"System","���ڶһ���Ч����",0,0)
			

		 Else
		    If Score=0 Then
			   Rs.Close:Set RS=Nothing
			   Response.Write "<script>alert('������Ļ��ֲ���ȷ,���ֱ������0!');history.back();</script>"
			   Exit Sub
			End If
			If KS.ChkClng(Score)<KS.ChkClng(KS.Setting(42)) Then
			   Rs.Close:Set RS=Nothing
			   Response.Write "<script>alert('������Ļ��ֲ���ȷ,���ֱ�����ڵ���" & KS.Setting(42) &"!');history.back();</script>"
			   Exit Sub
			End If
		   IF KS.ChkClng(RS("Score"))<KS.ChkClng(Score) Then
			   Rs.Close:Set RS=Nothing
			   Response.Write "<script>alert('����û��ֲ��㣬���ܶһ�!');history.back();</script>"
			   Exit Sub
		   End If
		    RS("Score")=RS("Score")-Score
			   ValidDays=rs("Edays")
				tmpDays=ValidDays-DateDiff("D",rs("BeginDate"),now())
				if tmpDays>0 then
					rs("Edays")=rs("Edays")+Score/KS.Setting(42)
				else
					rs("BeginDate")=now
					rs("Edays")=Score/KS.Setting(42)
				end if
			RS.Update
			'  UserName,InOrOutFlag,Edays,User,Descript
			Call KS.EdaysInOrOut(rs("UserName"),1,Score/KS.Setting(42),"System","���ֶһ�����")
		 End IF
		 Response.Write "<script>alert('��ϲ������Ч�����һ��ɹ�!');location.href='User_ExChange.asp?Action=Edays';</script>"
		 RS.Close:Set RS=Nothing
	   End Sub
End Class
%> 
