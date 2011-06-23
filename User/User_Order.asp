<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Conn.asp"-->
<!--#include file="../Plus/md5.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%

Dim KSCls
Set KSCls = New User_Order
KSCls.Kesion()
Set KSCls = Nothing

Class User_Order
        Private KS,KSUser
		Private CurrentPage,totalPut
		Private RS,MaxPerPage,Action
		Private TempStr,SqlStr
		Private InfoIDArr,InfoID,DomainStr
		Private Sub Class_Initialize()
			MaxPerPage =20
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		  DomainStr=KS.GetDomain
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
		If KS.S("page") <> "" Then
		      CurrentPage = KS.ChkClng(Request("page"))
		Else
			  CurrentPage = 1
		End If
		Action=Request("action")
		Call KSUser.Head()
		Call KSUser.InnerLocation("�ҵĶ���")
		Select Case Action
		  Case "ShowOrder" Call ShowOrder
		  Case "DelOrder" Call DelOrder
		  Case "SignUp"  Call SignUp
		  Case "AddPayment"  Call AddPayment '���˻�����
		  Case "SavePayment"  Call SavePayment
		  Case "coupon"  Call CouPon
		  case "dosave"   dosave
		  Case "PaymentOnline"  '����֧��
		   Response.Redirect "User_PayOnline.asp?Action=Payonline&id=" & KS.S("ID")
		  Case Else Call OrderList
		 End Select
		End Sub
		
		Sub OrderList()
		%>
		<div class="tabs">	
			<ul>
				<li<%If action<>"coupon" then ks.echo " class='select'"%>><a href="?">�ҵĶ���</a></li>
				<li<%If action="coupon" then ks.echo " class='select'"%>><a href="?action=coupon">�ҵ��Ż�ȯ</a></li>
			</ul>
        </div>
				
				<div style="text-align:center">
				<form action="user_order.asp" method="post" name="search">
				<strong>����״̬:</strong><select name="OrderStatus">
				 <option value="">������</option>
				  <option value="0">�ȴ�ȷ��</option>
				  <option value="1">�Ѿ�ȷ��</option>
				  <option value="2">�ѽ���</option>
				</select>
				<strong>�������:</strong>
				 <input type="text" name="keyword" class="textbox">
				 <input type="submit" value="��������" class="button">
				</form>				   
				</div>

				<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
					<tr class=title align=middle>
					  <td height="25" width=90>�������</td>
					  <td width=80>�û���</td>
					  <td width=100>�µ�ʱ��</td>
					  <td width=70>�������</td>
					  <td width=70>Ӧ�����</td>
					  <td width=60>�Ѹ����</td>
					  <td width=80>��Ʊ��Ϣ</td>
					  <td width=70>����״̬</td>
					  <td width=70>����״̬</td>
					  <td width=70>����״̬</td>
					</tr>
					<%
					  Dim Param:Param=" Where UserName='" & KSUser.UserName & "'"
					  If KS.S("OrderStatus")<>"" Then 
					    Param=Param & " and status=" & KS.ChkClng(KS.S("OrderStatus"))
					  End If
					  If KS.S("KeyWord")<>"" Then  
					    Param=Param & " and OrderID like '%" & KS.S("KeyWord") & "%'"
					  End If
					  
						 SqlStr="Select * From KS_Order " & Param & " order by id desc"
						 Set RS=Server.createobject("adodb.recordset")
						 RS.open SqlStr,conn,1,1

						 If RS.EOF And RS.BOF Then
								  Response.Write "<tr class='tdbg'><td align=center height=25 colspan=9 valign=top>��û�����κζ���!</td></tr>"
								 Else
									totalPut = RS.RecordCount
						
									If CurrentPage < 1 Then	CurrentPage = 1
			
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
		  <%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
		  </td>
		  </tr>
</table>
      </div>
		  <%
  End Sub
    
  Sub ShowContent()
    Dim i,MoneyTotal,MoneyReceipt
   Do While Not RS.Eof
		%>
                <tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
                       <td class="splittd"><a href="User_Order.asp?Action=ShowOrder&ID=<%=RS("ID")%>"><%=rs("orderid")%></a></td>
					   <td class="splittd" height="22" align="center"><%=rs("username")%></td>
                       <td class="splittd" align="center"><%=KS.GetTimeFormat(rs("inputtime"))%></td>
                       <td class="splittd" align="right"><%=formatnumber(rs("NoUseCouponMoney"),2,-1)%></td>
                       <td class="splittd" align="right"><%=formatnumber(rs("MoneyTotal"),2,-1)%></td>
                       <td class="splittd" align="right"><%=formatnumber(rs("MoneyReceipt"),2,-1)%></td>
                       <td class="splittd" align="center">
											<%If RS("NeedInvoice")=1 Then
											     Response.Write "<Font color=red>��Ҫ</font>"
											  	 If RS("Invoiced")=1 Then
												   Response.Write "<font color=green>(�ѿ�)</font>"
												  Else
												   Response.Write "<font color=red>(δ��)</font>"
												  End If
                                              Else
											    Response.Write "-"
											  End If
											 
											  %>
						</td>
                        <td class="splittd" align="center">
											<%If RS("Status")=0 Then
												  Response.Write "<font color=red>�ȴ�ȷ��</font>"
												  ElseIf RS("Status")=1 Then
												  Response.WRITE "<font color=green>�Ѿ�ȷ��</font>"
												  ElseIf RS("Status")=2 Then
												  Response.Write "<font color=#a7a7a7>�ѽ���</font>"
												  ElseIf RS("Status")=3 Then
												  Response.Write "<font color=#a7a7a7>��Ч����</font>"
				                              End If%></td>
                           <td class="splittd" align="center">
											<%If RS("MoneyReceipt")<=0 Then
											   Response.Write "<font color=red>�ȴ����</font>"
											  ElseIf RS("MoneyReceipt")<RS("MoneyTotal") Then
											   Response.WRITE "<font color=blue>���ն���</font>"
											  Else
											   Response.Write "<font color=green>�Ѿ�����</font>"
											  End If
											  %></td>
                         <td class="splittd" align="center">				
											<% If RS("DeliverStatus")=0 Then
											 Response.Write "<font color=red>δ����</font>"
											 ElseIf RS("DeliverStatus")=1 Then
											  Response.Write "<font color=blue>�ѷ���</font>"
											 ElseIf RS("DeliverStatus")=2 Then
											  Response.Write "<font color=green>��ǩ��</font>"
											 ElseIf RS("DeliverStatus")=3 Then
											  Response.Write "<font color=#ff6600>�˻�</font>"
											 End If
											 %></td>

                          </tr>

                                      <%
							MoneyReceipt=RS("MoneyReceipt")+MoneyReceipt
							MoneyTotal=RS("MoneyTotal")+MoneyTotal
							RS.MoveNext
							I = I + 1
					  If I >= MaxPerPage Then Exit Do
				    Loop
%>
 <tr align='center' class='tdbg' onMouseOut="this.className='tdbg'" onMouseOver="this.className='tdbgmouseover'">           <td colspan='4' align='right'><b>��ҳ�ϼƣ�</b></td>           <td align='right'><%=formatnumber(MoneyTotal,2)%></td>           <td align='right'><%=formatnumber(MoneyReceipt,2)%></td>           <td colspan='5'>&nbsp;</td>         </tr> 
 <tr align='center' class='tdbg' onMouseOut="this.className='tdbg'" onMouseOver="this.className='tdbgmouseover'">           <td colspan='4' align='right'><b>�����ܼƣ�</b></td>           <td align='right'><%=formatnumber(Conn.execute("Select sum(moneytotal) from KS_Order Where UserName='" & KSUser.UserName & "'")(0),2)%></td>           <td align='right'><%=formatnumber(Conn.execute("Select sum(MoneyReceipt) from KS_Order Where UserName='" & KSUser.UserName & "'")(0),2)%></td>           <td colspan='5'>&nbsp;</td>         </tr> 
                               
		<%  End Sub
		
		Sub ShowOrder()
		 Dim ID:ID=KS.ChkClng(KS.S("ID"))
		 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		 RS.Open "Select * from ks_order where username='" & KSUser.UserName & "' and id=" & ID ,conn,1,1
		 IF RS.Eof And RS.Bof Then
		   RS.Close:Set RS=Nothing
		   Response.Write "<script>alert('��������!');history.back();</script>"
		   response.end
		 End If
		 
		response.write "<br>"
		response.write OrderDetailStr(RS)
         %><br>
		 <div align=center id='buttonarea'>
		 <% 
If RS("Status")=3 Then
			   response.write "��������ָ��ʱ����û�и���,������!"
			 Else
		 if rs("status")=0 and rs("DeliverStatus")=0 and rs("MoneyReceipt")=0 Then%>
		 <input class="button" type='button' name='Submit' value='ɾ������' onClick="javascript:if(confirm('ȷ��Ҫɾ���˶�����')){window.location.href='User_Order.asp?Action=DelOrder&ID=<%=rs("id")%>';}">&nbsp;&nbsp;
		 <%end if%>
		 <%If RS("MoneyReceipt")<RS("MoneyTotal") Then%>
		 <span>
		 <input class="button" type='button' name='Submit' value='����֧��' onClick="window.location.href='user_PayOnline.asp?Action=Payonline&ID=<%=rs("id")%>'">
		 </span>
		 <input class="button" type='button' name='Submit' value='������пۿ�֧��' onClick="window.location.href='User_Order.asp?Action=AddPayment&ID=<%=rs("id")%>'">&nbsp;&nbsp;
		 <%end if%>
		 <% if rs("DeliverStatus")=1 Then%>
		 <input class="button" type='button' name='Submit' value='ǩ����Ʒ' onClick="window.location.href='User_Order.asp?Action=SignUp&ID=<%=RS("ID")%>'">
		 <%end if%>
		 <%
			  end if

		 %>
		 <input class="button" type='button' name='Submit' value='��ӡ����' onClick="document.all.buttonarea.style.display='none';window.print();">
		&nbsp; <input class="button" type='button' name='Submit' value='������ҳ' onClick="location.href='User_Order.asp';">
		 </div>
		 <br />
		 <br />
		 <%
		End Sub
		
		'�Ż�ȯ
		Sub Coupon
		Call KSUser.InnerLocation("�Ż�ȯ��ѯ")
		%>
		<div class="tabs">	
			<ul>
				<li<%If action<>"coupon" then ks.echo " class='select'"%>><a href="?">�ҵĶ���</a></li>
				<li<%If action="coupon" then ks.echo " class='select'"%>><a href="?action=coupon">�ҵ��Ż�ȯ</a></li>
			</ul>
        </div>
        <script src="../ks_inc/kesion.box.js"></script>
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
			<tr height="28" class="title">
				<td height="25" align="center">�Ż�ȯ��</td>
				<td height="25" align="center">�Ż�ȯ����</td>
				<td align="center">��ֵ</td>
				<td align="center">ʣ����</td>
				<td  align="center">��С�������</td>
				<td  align="center">��ֹʹ������</td>
				<td align="center">���ֿ۶�</td>
				<td align="center">ʹ�����</td>
			</tr>
		<% 
			set rs=server.createobject("adodb.recordset")
			sql="select a.*,b.title,b.enddate,b.facevalue,b.minamount,b.maxdiscount from KS_ShopCouponUser a inner join KS_ShopCoupon b on a.couponid=b.id where a.Username='"&KSUser.UserName&"' order by a.id desc"
			rs.open sql,Conn,1,1
			if rs.eof and rs.bof then
		%>
			<tr>
				<td height="26" colspan=10align=center valign=middle  class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">�Բ���,��û���Ż�ȯ���ã�</td>
			</tr>
		<%else
		
		                       totalPut = RS.RecordCount
								If CurrentPage < 1 Then	CurrentPage = 1
								
								   If CurrentPage > 1  and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
									Else
										CurrentPage = 1
									End If
									i=0
		      do while not rs.eof
		%>
						<tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
							<td class="splittd" align="center"><div class="ContentTitle"><%=rs("couponnum")%></div></td>
							<td height="25" class="splittd">
							<%=rs("title")%>
							
							</td>
							<td class="splittd" align=center>
							<%=RS("facevalue")%> Ԫ
							</td>
							<td class="splittd" align=center>
							<font color=red><%=RS("AvailableMoney")%></font> Ԫ
							</td>
							<td class="splittd" align=center>
							<%=RS("minAmount")%> Ԫ
							</td>
							<td class="splittd" align=center>
							<%=formatdatetime(RS("EndDate"),2)%>
							</td>
							<td class="splittd" align=center>
							<%
							If rs("maxdiscount")="0" Then
							Response.Write "ʵ���Ż�ȯ��ֵ"
						   Else
							Response.Write "�������ܶ��" & formatnumber(rs("maxdiscount"),2,-1) & "%,��������ʵ���Ż�ȯ��ֵ"
						   End If
							%>
							
							</td>
							
							<td class="splittd" align=center>
							<%select case  rs("useflag")
								 case 1
								     if RS("AvailableMoney")>0 then
									  response.write "��ʹ��,δ����"
									 else
									  response.write "������"
									 end if
									 response.write "<span style='cursor:pointer' onclick=""mousepopup(event,'˵��','" & RS("note") & "',300)""><font color=blue>(����)</font></span>"
								 case else
								  response.write " <font color=#999999>δʹ��</font>"
								end select
							%>
							</td>
							
						</tr>
		<%
			  rs.movenext
			  I = I + 1
			  If I >= MaxPerPage Then Exit Do
			
			loop
			end if
			rs.close
			set rs=Nothing
		%>
						
				
</table>
   
    <div style="text-align:right">
   <%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
    </div>
	<div style="clear:both"></div>
	  <br><br><br>
	  
	  <table width="98%" border="0" align="center" cellpadding="0" cellspacing="1" class="border">
	        <form name="myform" action="?action=dosave" method="post">
	        <tr>
			   <td height="35">&nbsp;<img src="images/ico1.gif" align="absmiddle"> <strong>����Ż�ȯ</strong></td>
			<tr>
				<td  class="splittd" height="35">&nbsp;&nbsp;&nbsp;&nbsp; <strong>�������Ż�ȯ��:</strong>
				<input type="text" name="CouponNum" class="textbox">
				<input type="submit" value=" �� �� " class="button">
				</td>
			</tr>
			</form>
	   </table>	

		<%
		end sub
		Sub dosave()
		   Dim CouponNum:CouponNum=KS.S("CouponNum")
	   
		
	       If CouponNum="" Then Response.Write "<script>alert('�Ż�ȯ�ű�������!');history.back();</script>":response.end
           If KS.ChkClng(Session("CouponNum"))>=3 Then 
		    Response.Write "<script>alert('�Բ���,���Ĵ����������,�ѹر�!');history.back();</script>":response.end
		   End If
            Dim RS: Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select * From KS_ShopCouponUser Where CouponNum='" & CouponNum & "'",Conn,1,3
			  If RS.Eof And RS.Bof Then
			   Session("CouponNum")=KS.ChkClng(Session("CouponNum"))+1
			   RS.Close:Set RS=Nothing
			   Response.Write "<script>alert('�Բ���,��������Ż�ȯ�Ų���ȷ!�������" & Session("CouponNum") & "��!');history.back();</script>":response.end 
			  ElseIf RS("UserName")<>"" And Not IsNull(RS("UserName")) Then
			   RS.Close:Set RS=Nothing
			   Response.Write "<script>alert('�Բ���,��������Ż�ȯ���ѱ����!');history.back();</script>":response.end 
			  Else
				 RS("UserName")=KSUser.UserName
		 		 RS.Update
			 End If
			     RS.Close
				 Set RS=Nothing
            Response.Write "<script>alert('��ϲ,�Ż�ȯ��ӳɹ�!');location.href='" & Request.ServerVariables("HTTP_REFERER") & "';</script>"
	   End Sub
		
		'ɾ������
		Sub DelOrder()
		  Dim ID:ID=KS.ChkClng(KS.S("ID"))
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select OrderID,CouponUserID From KS_Order where status=0 and DeliverStatus=0 and MoneyReceipt=0 and id=" & ID,Conn,1,3
		 If Not rs.EOF Then
		   Conn.execute("Update KS_ShopCouponUser Set UseFlag=0,OrderID='' Where ID=" & rs(1))
		   Conn.execute("delete from ks_orderitem Where OrderID='" & rs(0) &"'")
		   rs.delete
		 End if
         Response.redirect "User_Order.asp"
		End Sub
		
		'ǩ����Ʒ
		Sub SignUp()
		 Dim OrderID,id:ID=KS.ChkClng(KS.S("ID"))
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select * From KS_Order Where ID=" & ID,Conn,1,3
		 If RS.Eof Then
		  rs.close:set rs=nothing
		  response.write "<script>alert('������!');history.back();</script>":response.end
		 End If
         rs("DeliverStatus")=2
		 rs("BeginDate")=Now
		 rs.update
		 OrderID=RS("OrderID")
		 rs.close:set rs=nothing
		 Conn.execute("Update KS_LogDeliver Set Status=1 Where OrderID='" & OrderID & "'")
		 Response.Redirect "User_Order.asp?Action=ShowOrder&ID=" & id
		End Sub
		
		Sub AddPayment()
		 Dim ID:ID=KS.ChkClng(KS.S("ID"))
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_Order Where ID="& ID,Conn,1,1
		 If RS.Eof Then
		  rs.close:set rs=nothing
		  response.write "<script>alert('������!');history.back();</script>":response.end
		 End If
		 %>
		 <FORM name=form4 onSubmit="return confirm('ȷ�����������Ϣ����ȫ��ȷ��һ��ȷ�ϾͲ��ɸ���Ŷ��')" action=User_Order.asp method=post>
  <table class=border cellSpacing=1 cellPadding=2 width="98%" align="center" border=0>
    <tr class=title align=middle>
      <td colSpan=2 height=22><B>ʹ���˻��ʽ�֧������</B></td>
    </tr>
    <tr class=tdbg>
      <td align=right><B>�� �� ����</B></td>
      <td align=left><%=KSUser.UserName%></td>
    </tr>
    <tr class=tdbg>
      <td align=right><B>�ͻ����ƣ�</B></td>
      <td align=left><%=RS("ContactMan")%></td>
    </tr>
    <tr class=tdbg>
      <td align=right><B>�ʽ���</B></td>
      <td align=left><%=formatnumber(KSUser.Money,2,-1)%> Ԫ <%if Round(KSUser.Money)<=0 then response.write "<a href=""user_payonline.asp"">�ʽ���,���˳�ֵ</a>"%></td>
    </tr>
    <tr class=tdbg>
      <td align=right><B>֧�����ݣ�</B></td>
      <td align=left>
        <table cellSpacing=2 cellPadding=0 border=0>
          <tr>
            <td align=right>������ţ�</td>
            <td align=left>
              <%=RS("OrderID")%></td>
          </tr>
          <tr>
            <td align=right>������</td>
            <td align=left><font color=red><%=formatnumber(RS("MoneyTotal"),2,-1)%></font> Ԫ</td>
          </tr>
          <tr>
            <td align=right>�� �� �</td>
            <td align=left>
             <font color=blue><%=formatnumber(RS("MoneyReceipt"),2,-1)%></font> Ԫ</td>
          </tr>
        </table>
      </td>
    </tr>
    <tr class=tdbg>
      <td align=right><B>֧����</B></td>
      <td align=left>
        <Input id="Money" readonly maxLength=20 size=10 value="<%=rs("moneytotal")-rs("MoneyReceipt")%>" name="Money"> Ԫ&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color=#0000ff>֧���ɹ��󣬽��������ʽ�����п۳���Ӧ���</font></td>
    </tr>
    <tr class=tdbg>
      <td colSpan=2 height=30><B><font color=#ff0000>ע�⣺֧����Ϣһ��¼�룬�Ͳ������޸ģ������ڱ���֮ǰȷ����������</font></B></td>
    </tr>
    <tr class=tdbg align=middle>
      <td colSpan=2 height=30>
  <Input id=Action type="hidden" value="SavePayment" name="Action"> 
  <Input id=ID type=hidden value="<%=rs("id")%>" name="ID"> 
        <Input type=submit value=" ȷ��֧�� " class="button" name=Submit></td>
    </tr>
  </table>
</FORM>
		 <%
		 rs.close:set rs=nothing
		End Sub
		
		'��ʼ���֧������
		Sub SavePayment()
		 Dim ID:ID=KS.ChkClng(KS.S("ID"))
		 Dim Money:Money=KS.S("Money")
		 If Not IsNumeric(Money) Then Response.Write "<script>alert('��������Ч�Ľ��!');history.back();</script>":Response.end
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_Order Where ID=" & ID,Conn,1,3
		 If RS.Eof Then
		   RS.Close:Set RS=Nothing
		   Response.Write "<script>alert('������!');history.back();</script>"
		 End If
		 If Round(Money)>Round(KSUser.Money) or Round(KSUser.Money)<=0  Then
		  %>
		  <br><br>
		  <table cellpadding=2 cellspacing=1 border=0 width=400 class='border' align=center>
		  <tr align='center' class='title'><td height='22'><strong>������Ϣ</strong></td></tr>
		  <tr class='tdbg'><td height='100' valign='top'><b>��������Ŀ���ԭ��</b><li>�������֧�������������ʽ�����Ч֧����</li></td></tr>
		  <tr align='center' class='tdbg'><td><a href='javascript:history.go(-1)'>&lt;&lt; ������һҳ</a></td></tr>
		</table>
		  <%
		  RS.Close:Set RS=Nothing:Exit Sub
		 End If

		   
		   RS("MoneyReceipt")=RS("MoneyReceipt")+Money
		   RS("Status")=1
		   RS("PayTime")=now   '��¼����ʱ��
		   RS.Update
		   
		   Call KS.MoneyInOrOut(RS("UserName"),RS("Contactman"),Money,4,2,now,RS("OrderID"),KSUser.UserName,"֧���������ã������ţ�" & RS("Orderid"),0,0)

	
					'====================Ϊ�û����ӹ���Ӧ�û���========================
					If RS("MoneyReceipt")>=RS("MoneyTotal") Then
						Dim rsp:set rsp=conn.execute("select point,id,title from ks_product where id in(select proid from KS_OrderItem where orderid='" & rs("orderid") & "')")
						do while not rsp.eof
						  dim amount:amount=conn.execute("select amount from ks_orderitem where orderid='" & rs("orderid") & "' and proid=" & rsp(1))(0)
						  Call KS.ScoreInOrOut(KSUser.UserName,1,KS.ChkClng(rsp(0))*amount,"ϵͳ","������Ʒ<font color=red>" & rsp("title") & "</font>����!",0,0)
						rsp.movenext
						loop
						rsp.close
						set rsp=nothing
					End If
					'================================================================
		 
		 RS.Close:Set RS=Nothing
		  Response.Redirect "User_Order.asp?Action=ShowOrder&id=" & id 
		End Sub
		
		'���ض�����ϸ��Ϣ
		Function  OrderDetailStr(RS)
		 OrderDetailStr="<table width='98%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'> "&vbcrlf
		 OrderDetailStr=OrderDetailStr & "<tr align='center' class='title'>    <td height='22'><b>�� �� �� Ϣ</b>��������ţ�" & RS("ORDERID") & "��</td>  </tr>"&vbcrlf
		 OrderDetailStr=OrderDetailStr & "<tr>" & vbcrlf
		 OrderDetailStr=OrderDetailStr & " <td height='25'>" &vbcrlf
		 OrderDetailStr=OrderDetailStr & "  <table width='100%'  border='0' cellpadding='2' cellspacing='0'> "   & vbcrlf
		 OrderDetailStr=OrderDetailStr & "    <tr class='tdbg'>"
		 OrderDetailStr=OrderDetailStr & "	         <td width='18%'>�ͻ�������<font color='red'>" & RS("Contactman") & "</td>      "
		 OrderDetailStr=OrderDetailStr & "			 <td width='20%'>�� �� ����<font color='red'>" & rs("username") & "</td> " &vbcrlf
		OrderDetailStr=OrderDetailStr & "			 <td width='20%'>�� �� �̣�</td>"
		OrderDetailStr=OrderDetailStr & "			 <td width='18%'>�������ڣ�<font color='red'>" & formatdatetime(rs("inputtime"),2) & "</font></td>" & vbcrlf
		OrderDetailStr=OrderDetailStr & "			 <td width='24%'>�µ�ʱ�䣺<font color='red'>" & rs("inputtime") & "</font></td>" & vbcrlf
		OrderDetailStr=OrderDetailStr & "	</tr>"
		OrderDetailStr=OrderDetailStr & "	<tr class='tdbg'> "      
		OrderDetailStr=OrderDetailStr & "	  <td width='18%'>��Ҫ��Ʊ��"
			    If RS("NeedInvoice")=1 Then
				  OrderDetailStr=OrderDetailStr & "<Font color=red>��</font>"
				  Else
				  OrderDetailStr=OrderDetailStr & "<font color=red>��</font>"
				  End If
		OrderDetailStr=OrderDetailStr & "		 </td> "       
		OrderDetailStr=OrderDetailStr & "	 <td width='20%'>�ѿ���Ʊ��"	
				  If RS("Invoiced")=1 Then
				   OrderDetailStr=OrderDetailStr & "<font color=green>��</font>"
				  Else
				   OrderDetailStr=OrderDetailStr & "<font color=red>��</font>"
				  End If
		OrderDetailStr=OrderDetailStr & "	</td> "
		OrderDetailStr=OrderDetailStr & "	<td width='20%'>����״̬��"	
			if RS("Status")=0 Then
				 OrderDetailStr=OrderDetailStr & "<font color=red>�ȴ�ȷ��</font>"
				  ElseIf RS("Status")=1 Then
				 OrderDetailStr=OrderDetailStr & "<font color=green>�Ѿ�ȷ��</font>"
				  ElseIf RS("Status")=2 Then
				 OrderDetailStr=OrderDetailStr & "<font color=#a7a7a7>�ѽ���</font>"
				  End If
		OrderDetailStr=OrderDetailStr & "	</td>"
		OrderDetailStr=OrderDetailStr & "	  <td width='18%'>���������"	
			     If RS("MoneyReceipt")<=0 Then
				   OrderDetailStr=OrderDetailStr & "<font color=red>�ȴ����</font>"
				  ElseIf RS("MoneyReceipt")<RS("MoneyTotal") Then
				   OrderDetailStr=OrderDetailStr & "<font color=blue>���ն���</font>"
				  Else
				  OrderDetailStr=OrderDetailStr & "<font color=green>�Ѿ�����</font>"
				  End If

       OrderDetailStr=OrderDetailStr & "</td>"
	   OrderDetailStr=OrderDetailStr & "        <td width='24%'>����״̬��"
				if RS("DeliverStatus")=0 Then
				 OrderDetailStr=OrderDetailStr & "<font color=red>δ����</font>"
				 ElseIf RS("DeliverStatus")=1 Then
				  OrderDetailStr=OrderDetailStr & "<font color=blue>�ѷ���</font>"
				 ElseIf RS("DeliverStatus")=2 Then
				  OrderDetailStr=OrderDetailStr & "<font color=blue>��ǩ��</font>"
				 ElseIf RS("DeliverStatus")=3 Then
				  OrderDetailStr=OrderDetailStr & "<font color=#ff6600>�˻�</font>"
				 End If
	OrderDetailStr=OrderDetailStr & "		</td></tr>    </table> "
    OrderDetailStr=OrderDetailStr & " </td>  </tr> " 
	OrderDetailStr=OrderDetailStr & "   <tr align='center'>"
	OrderDetailStr=OrderDetailStr & "       <td height='25'>"
	OrderDetailStr=OrderDetailStr & "	   <table width='100%' border='0' align='center' cellpadding='2' cellspacing='1'>"
	OrderDetailStr=OrderDetailStr & "	           <tr class='tdbg'>"
	OrderDetailStr=OrderDetailStr & "			             <td width='12%' align='right'>�ջ���������</td>"
	OrderDetailStr=OrderDetailStr & "						 <td width='38%'>" & rs("contactman") & "</td>"
	OrderDetailStr=OrderDetailStr & "						 <td width='12%' align='right'>��ϵ�绰��</td> "      
	OrderDetailStr=OrderDetailStr & "						 <td width='38%'>" & rs("phone") & "</td>"
	OrderDetailStr=OrderDetailStr & "				</tr>"
	OrderDetailStr=OrderDetailStr & "				<tr class='tdbg' valign='top'>"
	OrderDetailStr=OrderDetailStr & "				          <td width='12%' align='right'>�ջ��˵�ַ��</td>"    
	OrderDetailStr=OrderDetailStr & "						  <td width='38%'>" & rs("address") & "</td>"          
	OrderDetailStr=OrderDetailStr & "						  <td width='12%' align='right'>�������룺</td>"    
	OrderDetailStr=OrderDetailStr & "						  <td width='38%'>" &rs("zipcode") & "</td>"
	OrderDetailStr=OrderDetailStr & "				</tr>  "      
	OrderDetailStr=OrderDetailStr & "				<tr class='tdbg'> "         
	OrderDetailStr=OrderDetailStr & "				          <td width='12%' align='right'>�ջ������䣺</td>"    
	OrderDetailStr=OrderDetailStr & "						  <td width='38%'>" & rs("email") & "</td> "         
	OrderDetailStr=OrderDetailStr & "						  <td width='12%' align='right'>�ջ����ֻ���</td>"    
	OrderDetailStr=OrderDetailStr & "						  <td width='38%'>" & rs("mobile") & "</td>       "
	OrderDetailStr=OrderDetailStr & "			   </tr>"        
	OrderDetailStr=OrderDetailStr & "			   <tr class='tdbg'> "         
	OrderDetailStr=OrderDetailStr & "			              <td width='12%' align='right'>���ʽ��</td>"    
	OrderDetailStr=OrderDetailStr & "						  <td width='38%'>" & KS.ReturnPayMent(rs("PaymentType"),0) & "</td>       "   
	OrderDetailStr=OrderDetailStr & "						  <td width='12%' align='right'>�ͻ���ʽ��</td>"          
	OrderDetailStr=OrderDetailStr & "						  <td width='38%'>" & KS.ReturnDelivery(rs("DeliverType"),0) & "</td>        "
	OrderDetailStr=OrderDetailStr & "				</tr> "       
	OrderDetailStr=OrderDetailStr & "				<tr class='tdbg' valign='top'>  "        
	OrderDetailStr=OrderDetailStr & "				          <td width='12%' align='right'>��Ʊ��Ϣ��</td>"          
	OrderDetailStr=OrderDetailStr & "						  <td width='38%'>"
	 If RS("Invoiced")=1 Then OrderDetailStr=OrderDetailStr & rs("InvoiceContent") &"</td>"
    OrderDetailStr=OrderDetailStr & "						 <td width='12%' align='right'>��ע/���ԣ�</td>"          
	OrderDetailStr=OrderDetailStr & "							<td width='38%'>" & rs("Remark") & "</td>       "
	OrderDetailStr=OrderDetailStr & "				 </tr>  "  
	OrderDetailStr=OrderDetailStr & "				 </table>"
	OrderDetailStr=OrderDetailStr & "			</td>  "
	OrderDetailStr=OrderDetailStr & "		</tr>  "
	OrderDetailStr=OrderDetailStr & "		<tr><td>"
	OrderDetailStr=OrderDetailStr & "		<table width='100%' border='0' align='center' cellpadding='2' cellspacing='1' class='border'> "
	OrderDetailStr=OrderDetailStr & "		  <tr align='center' class='title' height='25'>  "  
	OrderDetailStr=OrderDetailStr & "		   <td><b>�� Ʒ �� ��</b></td> "   
	OrderDetailStr=OrderDetailStr & "		   <td width='45'><b>��λ</b></td>  "  
	OrderDetailStr=OrderDetailStr & "		   <td width='55'><b>����</b></td>  "  
	OrderDetailStr=OrderDetailStr & "		   <td width='65'><b>ԭ��</b></td>  "  
	OrderDetailStr=OrderDetailStr & "		   <td width='65'><b>ʵ��</b></td>  "  
	OrderDetailStr=OrderDetailStr & "		   <td width='65'><b>ָ����</b></td>  "  
	OrderDetailStr=OrderDetailStr & "		   <td width='85'><b>�� ��</b></td>   " 
	OrderDetailStr=OrderDetailStr & "		   <td width='65'><b>��������</b></td>  "  
	OrderDetailStr=OrderDetailStr & "		   <td width='45'><b>��ע</b></td>  "
	OrderDetailStr=OrderDetailStr & "		  </tr> "
			 Dim TotalPrice,attributecart,RSI:Set RSI=Server.CreateObject("ADODB.RECORDSET")
			   RSI.Open "Select * From KS_OrderItem Where SaleType<>5 and SaleType<>6 and OrderID='" & RS("OrderID") & "' order by ischangedbuy,id",conn,1,1
			   If RSI.Eof Then
			     RSI.Close:Set RSI=Nothing
				' Response.Write "<script>alert('�Ҳ��������Ʒ');history.back();<//script>"
			  Else
			  Do While Not RSI.Eof
			  attributecart=rsi("attributecart")
			  if not ks.isnul(attributecart) then attributecart="<br/><font color=#888888>" & attributecart & "</font>"
		OrderDetailStr=OrderDetailStr & "	  <tr valign='middle' class='tdbg' height='20'>"    
		OrderDetailStr=OrderDetailStr & "	   <td width='*'><a href='" & DomainStr & "item/show.asp?m=5&d=" & RSi("proid") & "' target='_blank'>" & Conn.execute("select top 1 title from ks_product where id=" & rsi("proid"))(0) 
		
		If RSI("IsChangedBuy")="1" Then OrderDetailStr=OrderDetailStr & "(����)"
		
		
			  Dim RSP:Set RSP=Server.CreateObject("ADODB.RECORDSET")
			  RSP.Open "Select top 1 I.Title,I.Unit,I.IsLimitBuy,I.LimitBuyPrice,L.LimitBuyPayTime From KS_Product I Left Join KS_ShopLimitBuy L On I.LimitBuyTaskID=L.Id  Where I.ID=" & RSI("ProID"),conn,1,1
			  dim title,unit,LimitBuyPayTime
			  If Not RSP.Eof Then
				  title=rsp("title")
				  Unit=rsp("unit")
				  If RSI("IsChangedBuy")=1 Then 
				   title=title &"(����)"
				  else
				   If LimitBuyPayTime="" Then
				   LimitBuyPayTime=RSP("LimitBuyPayTime")
				   ElseIf LimitBuyPayTime>RSP("LimitBuyPayTime") Then
				    LimitBuyPayTime=RSP("LimitBuyPayTime")
				   End If
				  end  if
				  If RSI("IsLimitBuy")="1" Then OrderDetailStr=OrderDetailStr & "<span style='color:green'>(��ʱ����)</span>"
				  If RSI("IsLimitBuy")="2" Then OrderDetailStr=OrderDetailStr & "<span style='color:blue'>(��������)</span>"
			  End If
			  RSP.Close:Set RSP=Nothing
		
		OrderDetailStr=OrderDetailStr & "</a>" & attributecart & "</td>    "
		OrderDetailStr=OrderDetailStr & "	   <td width='45' align=center>"& Conn.execute("select unit from ks_product where id=" & rsi("proid"))(0) & "</td>               <td width='55' align='center'>" & rsi("amount") &"</td>    "
		OrderDetailStr=OrderDetailStr & "	   <td width='65' align='center'>" & formatnumber(rsi("price_original"),2) & "</td>    "
		OrderDetailStr=OrderDetailStr & "	   <td width='65' align='center'>" & formatnumber(rsi("price"),2) & "</td>    "
		OrderDetailStr=OrderDetailStr & "	   <td width='65' align='center'>" & formatnumber(rsi("realprice"),2) & "</td>    "
		OrderDetailStr=OrderDetailStr & "	   <td width='85' align='right'>" & formatnumber(rsi("realprice")*rsi("amount"),2) & "</td>    "
		OrderDetailStr=OrderDetailStr & "	   <td width='65' align=center>" & rsi("ServiceTerm") & "</td>    "
		OrderDetailStr=OrderDetailStr & "	   <td align=center width='45'>" & rsi("Remark") & "</td>  "
		OrderDetailStr=OrderDetailStr & "	   </tr> " 
		OrderDetailStr=OrderDetailStr & GetBundleSalePro(TotalPrice,RSI("ProID"),RSI("OrderID"))  'ȡ������������Ʒ
		
		
			  TotalPrice=TotalPrice+ rsi("realprice")*rsi("amount")
			    rsi.movenext
			  loop
			  rsi.close:set rsi=nothing
		End If
		
		OrderDetailStr=OrderDetailStr & GetPackage(TotalPrice,RS("OrderID"))         '��ֵ���
		
		
		OrderDetailStr=OrderDetailStr & "	   <tr class='tdbg' height='30' > "   
		OrderDetailStr=OrderDetailStr & "	    <td colspan='6' align='right'><b>�ϼƣ�</b></td> "   
		OrderDetailStr=OrderDetailStr & "		<td align='right'><b>" & formatnumber(totalprice,2) & "</b></td>    "
		OrderDetailStr=OrderDetailStr & "		<td colspan='3'> </td>  "
		OrderDetailStr=OrderDetailStr & "	  </tr>    "
		OrderDetailStr=OrderDetailStr & "	  <tr class='tdbg'>"
       OrderDetailStr=OrderDetailStr & "         <td colspan='4'>���ʽ�ۿ��ʣ�" & rs("Discount_Payment") & "%&nbsp;&nbsp;&nbsp;&nbsp;�˷ѣ�" & rs("Charge_Deliver")&" Ԫ&nbsp;&nbsp;&nbsp;&nbsp;˰�ʣ�" & KS.Setting(65) &"%&nbsp;&nbsp;&nbsp;&nbsp;�۸�˰��"
				IF KS.Setting(64)=1 Then 
				   OrderDetailStr=OrderDetailStr & "��"
				  Else
				   OrderDetailStr=OrderDetailStr & "����˰"
				  End If
				  Dim TaxMoney
				  Dim TaxRate:TaxRate=KS.Setting(65)
				 If KS.Setting(64)=1 Or rs("NeedInvoice")=0 Then TaxMoney=1 Else TaxMoney=1+TaxRate/100

				OrderDetailStr=OrderDetailStr & "<br>ʵ�ʽ�(" & rs("MoneyGoods") & "��" & rs("Discount_Payment") & "%��"&rs("Charge_Deliver") & ")��"
				if KS.Setting(64)=1 Or rs("NeedInvoice")=0 Then OrderDetailStr=OrderDetailStr & "100%" Else OrderDetailStr=OrderDetailStr & "(1��" & TaxRate & "%)" 
				OrderDetailStr=OrderDetailStr & "��" & formatnumber(rs("NoUseCouponMoney"),2) & "Ԫ  </td>"
    OrderDetailStr=OrderDetailStr & "<td  colspan='3' align=right><b>������</b> ��" & formatnumber(rs("NoUseCouponMoney"),2) & " Ԫ<br>"
	If KS.ChkClng(RS("CouponUserID"))<>0 and RS("UseCouponMoney")>0 Then
	OrderDetailStr=OrderDetailStr & "<b>ʹ���Ż�ȯ��</b> <font color=#ff6600>��" & formatnumber(RS("UseCouponMoney"),2) & " Ԫ</font><br>"
	End If
	OrderDetailStr=OrderDetailStr & "<b>Ӧ����</b> ��" & formatnumber(rs("MoneyTotal"),2) & "  Ԫ</td>"
    OrderDetailStr=OrderDetailStr & "<td colspan='3' align='left'><b>�Ѹ��</b>��<font color=red>" & formatnumber(rs("MoneyReceipt"),2) & "</font></b>"
	If RS("MoneyReceipt")<RS("MoneyTotal") Then
	OrderDetailStr=OrderDetailStr & "<br><B>��Ƿ���<font color=blue>" & formatnumber(RS("MoneyTotal")-RS("MoneyReceipt"),2) &"</B>"
	End If
	OrderDetailStr=OrderDetailStr & "</td></tr></table></td>  "
	OrderDetailStr=OrderDetailStr & "</tr>"  
	OrderDetailStr=OrderDetailStr & "     <tr><td><br><b>ע��</b>��<font color='blue'>ԭ��</font>��ָ��Ʒ��ԭʼ���ۼۣ���<font color='green'>ʵ��</font>��ָϵͳ�Զ������������Ʒ���ռ۸񣬡�<font color='red'>ָ����</font>��ָ����Ա���ݲ�ͬ��Ա���ֶ�ָ�������ռ۸���Ʒ���������ۼ۸��ԡ�ָ���ۡ�Ϊ׼��<br>"
	OrderDetailStr=OrderDetailStr & "	 </td>"
	OrderDetailStr=OrderDetailStr & "	 </tr>"
	
	If not conn.execute("select top 1 * from ks_orderitem where orderid='" & RS("OrderID") &"' and islimitbuy<>0").eof Then
	OrderDetailStr=OrderDetailStr & "     <tr><td><div style='margin:10px;color:red;padding:10px;margin-top:3px;border:1px solid #f9c943;background:#FFFFF6'>��ܰ��ʾ:����������ʱ/������������,�����µ���" & LimitBuyPayTime & "Сʱ֮�ڱ��븶��,���������[" & DateAdd("h",LimitBuyPayTime,RS("InputTime")) & "]֮ǰ�û�û�и���,�������Զ����ϡ�</div>"
	OrderDetailStr=OrderDetailStr & "	 </td>"
	OrderDetailStr=OrderDetailStr & "	 </tr>"
	End If
	
	If RS("DeliverStatus")=1 Then
	 Dim RSD,DeliverStr
	 Set RSD=Conn.Execute("Select Top 1 * From KS_LogDeliver Where OrderID='" & RS("OrderID") & "'")
	 If Not RSD.Eof Then
	  DeliverStr="��ݹ�˾:" & RSD("ExpressCompany") & " ��������:" & RSD("ExpressNumber") & " ��������:" & RSD("DeliverDate") & " ����������:" & RSD("HandlerName")
	 End If
	 RSD.Close : Set RSD=Nothing
	OrderDetailStr=OrderDetailStr & "     <tr><td><div style='margin:10px;color:blue;padding:10px;margin-top:3px;border:1px solid #f9c943;background:#FFFFF6'>��ܰ��ʾ:�������ѷ�����" & DeliverStr & "</div>"
	OrderDetailStr=OrderDetailStr & "	 </td>"
	OrderDetailStr=OrderDetailStr & "	 </tr>"
	End If
	
	
	OrderDetailStr=OrderDetailStr & "	</table>"
	  End Function
	  
'ȡ������������Ʒ
Function GetBundleSalePro(ByRef TotalPrice,ProID,OrderID)
  Dim Str,RS,XML,Node
  Set RS=Server.CreateObject("adodb.recordset")
  RS.Open "Select I.Title,I.Unit,O.* From KS_OrderItem O inner join KS_Product I On O.ProID=I.ID Where O.SaleType=6 and BundleSaleProID=" & ProID & " and OrderID='" & OrderID & "' order by O.id",conn,1,1
  If Not RS.Eof Then
    Set XML=KS.RsToXml(rs,"row","")
  End If
  RS.Close:Set RS=Nothing
  If IsObject(XML) Then
	     str=str & "<tr height=""25"" align=""left""><td colspan=9 style=""color:green"">&nbsp;&nbsp;ѡ���������:</td></tr>"
       For Each Node In Xml.DocumentElement.SelectNodes("row")
         str=str & "<tr>"
		 str=str &" <td style='color:#999999'>&nbsp;" & Node.SelectSingleNode("@title").text &"</td>"
		 str=str &" <td align='center'>" & Node.SelectSingleNode("@unit").text &"</td>"
		 str=str &" <td align='center'>" & Node.SelectSingleNode("@amount").text &"</td>"
		 str=str &" <td align='center'>" & formatnumber(Node.SelectSingleNode("@price_original").text,2,-1) &"</td>"
		 str=str &" <td align='center'>" & formatnumber(Node.SelectSingleNode("@realprice").text,2,-1) &"</td>"
		 str=str &" <td align='center'>" & formatnumber(Node.SelectSingleNode("@realprice").text,2,-1) &"</td>"
		 str=str &" <td align='right'>" & formatnumber(Node.SelectSingleNode("@realprice").text*Node.SelectSingleNode("@amount").text,2,-1) &"</td>"
		 str=str &" <td align='center'>" & Node.SelectSingleNode("@serviceterm").text &"</td>"
		 str=str &" <td align='center'>" & Node.SelectSingleNode("@remark").text &"</td>"
		 str=str & "</tr>"
		 TotalPrice=TotalPrice +round(Node.SelectSingleNode("@realprice").text*Node.SelectSingleNode("@amount").text,2) 
       Next
  End If
  GetBundleSalePro=str
End Function
	  
	  
 '�õ���ֵ���
 Function GetPackage(ByRef TotalPrice,OrderID)
	    If KS.IsNul(OrderID) Then Exit Function
		Dim RS,RSB,GXML,GNode,str,n,Price
		Set RS=Conn.Execute("select packid,OrderID from KS_OrderItem Where SaleType=5 and OrderID='" & OrderID & "' group by packid,OrderID")
		If Not RS.Eof Then
		 Set GXML=KS.RsToXml(Rs,"row","")
		End If
		RS.Close : Set RS=Nothing
		If IsOBJECT(GXml) Then
		   FOR 	Each GNode In GXML.DocumentElement.SelectNodes("row")
		     Set RSB=Conn.Execute("Select top 1 * From KS_ShopPackAge Where ID=" & GNode.SelectSingleNode("@packid").text)
			 If Not RSB.Eof Then
					  
						Dim RSS:Set RSS=Server.CreateObject("adodb.recordset")
						RSS.Open "Select a.title,a.GroupPrice,a.Price_Member,a.Price,b.* From KS_Product A inner join KS_OrderItem b on a.id=b.proid Where b.SaleType=5 and b.packid=" & GNode.SelectSingleNode("@packid").text & " and  b.orderid='" & OrderID & "'",Conn,1,1
						  str=str & "<tr class='tdbg' height=""25"" align=""center""><td colspan=2><strong><a href='" & DomainStr & "shop/pack.asp?id=" & RSB("ID") & "' target='_blank'>" & RSB("PackName") & "</a></strong></td>"
						  n=1
						  Dim TotalPackPrice,tempstr,i
						  TotalPackPrice=0 : tempstr=""
						Do While Not RSS.Eof
						 
						  For I=1 To RSS("Amount") 
							  '�õ�����Ʒ�۸� 
							  IF KS.C("UserName")<>"" Then
								  If RSS("GroupPrice")=0 Then
								   Price=RSS("Price_Member")
								  Else
								   Dim RSP:Set RSP=Conn.Execute("Select Price From KS_ProPrice Where GroupID=(select groupid from ks_user where username='" & KS.C("UserName") & "') And ProID=" & RSS("ID"))
								   If RSP.Eof Then
									 Price=RSS("Price_Member")
								   Else
									 Price=RSP(0)
								   End If
								   RSP.Close:Set RSP=Nothing
								  End If
							  Else
								  Price=RSS("Price")
							  End If
							
							   TotalPackPrice=TotalPackPrice+Price
							  tempstr=tempstr & n & "." & rss("title") & " " & rss("AttributeCart") & "<br/>"
							  n=n+1
						  Next
						  RSS.MoveNext
						Loop
						
						str=str &"<td>1</td><td>��" & TotalPackPrice & "</td><td>" & rsb("discount") & "��</td><td>��" & formatnumber((TotalPackPrice*rsb("discount")/10),2,-1) & "</td><td>��" & formatnumber((TotalPackPrice*rsb("discount")/10),2,-1) & "</td><td>---</td><td>---</td>"
					   
						str=str & "</tr><tr><td align='left' colspan=9>��ѡ�����װ��ϸ����:<br/>" & tempstr & "</td></tr>" 
						
						TotalPrice=TotalPrice+round(formatnumber((TotalPackPrice*rsb("discount")/10),2,-1))   '������������ܼ�
						
						RSS.Close
						Set RSS=Nothing
					
			End If
			RSB.Close
		   Next
			
	    End If
		GetPackage=str
		
End Function


 End Class
%> 
