<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../Plus/md5.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New User_PayOnline
KSCls.Kesion()
Set KSCls = Nothing

Class User_PayOnline
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
		Call KSUser.InnerLocation("����֧��")
		Response.Write "<div class=""tabs"">"
		Response.Write " <ul class="""">"
		Response.Write " <li class='select'><a href=""User_PayOnline.asp"">����֧����ֵ</a></li>"
		Response.Write " <li><a href=""user_recharge.asp"">��ֵ����ֵ</a></li>"
		Response.Write " <li><a href=""user_exchange.asp?Action=Point"">�һ�" & KS.Setting(45) & "</a></li>"
		Response.Write " <li><a href=""user_exchange.asp?Action=Edays"">�һ���Ч��</a></li>"
		Response.Write " <li><a href=""user_exchange.asp?Action=Money"">" & KS.Setting(45) & "�һ��˻��ʽ�</a></li>"
		Response.Write "</ul>"
		Response.Write "</div>"
		Select Case KS.S("Action")
		 Case "PayStep2"
		    Call PayStep2()
		 Case "PayStep3"
		    Call PayStep3()
		 Case "Payonline"
		    Call PayShopOrder()
	     Case Else
		    Call PayOnline()
		End Select
       End Sub
	  
	   
	   Sub PayOnline()
	    %>
	   <script>
	     function Confirm()
		 {
		  if (document.myform.Money.value=="")
		  {
		   alert('��������Ҫ��ֵ�Ľ��!')
		   document.myform.Money.focus();
		   return false;
		  }
		  return true;
		  }
	   </script>
		<FORM name=myform action="User_PayOnline.asp" method="post">
		  <table class=border cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
			<tr class=title>
			  <td align=middle colSpan=2 height=22><B> �� �� �� ֵ</B></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=213>�û�����</td>
			  <td width="754"><%=KSUser.UserName%></td>
			</tr>
			<tr class=tdbg>
			  <td width="213" align=right>�Ʒѷ�ʽ��</td>
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
			  <td align=right width=213>�ʽ���</td>
			  <td><input type='hidden' value='<%=KSUser.Money%>' name='Premoney'><%=formatnumber(KSUser.Money,2,-1)%> Ԫ</td>
			</tr>
			<%If KSUser.ChargeType=1 then%>
			<tr class=tdbg>
			  <td align=right width=213>����<%=KS.Setting(45)%>��</td>
			  <td><%=KSUser.Point%>&nbsp;<%=KS.Setting(46)%></td>
			</tr>
			<%end if%>
			<%If KSUser.ChargeType=2 then%>
			<tr class=tdbg>
			  <td align=right width=213>ʣ��������</td>
			  <td>
			  <%if KSUser.ChargeType=3 Then%>
			  ������
			  <%else%>
			  <%=KSUser.GetEdays%>&nbsp;��
			  <%end if%></td>
			</tr>
		   <%end if%>
			<tr class=tdbg>
			  <td align=right>��ǰ����</td>
			  <td><%=KS.U_G(KSUser.GroupID,"groupname")%></td>
		    </tr>
			<tr class=title>
			  <td align=middle colSpan=2 height=22><B> ѡ �� �� �� �� ֵ �� ʽ</B></td>
			</tr>

			<tr class=tdbg>
			  <td colspan="2">
			  <%
			   Dim RSC,AllowGroupID:Set RSC=Conn.Execute("Select ID,GroupName,Money,AllowGroupID From KS_UserCard Where CardType=1 and DateDiff(" & DataPart_S & ",EndDate," & SqlNowString& ")<0")
			   Do While NOt RSC.Eof 
			      AllowGroupID=RSC("AllowGroupID") : If IsNull(AllowGroupID) Then AllowGroupID=" "
			     If KS.IsNul(AllowGroupID) Or KS.FoundInArr(AllowGroupID,KSUser.GroupID,",")=true Then
			    response.write "&nbsp;&nbsp; <label><input checked name=""UserCardID"" onclick=""$('#m').hide()"" type=""radio"" value=""" & rsc("ID") & """/>" & rsc(1) & " (��Ҫ���� <span style='color:red'>" & formatnumber(RSC(2),2,-1) & "</span> Ԫ)</label><br/>"
				End If
			    RSC.MoveNext
			   Loop
			   RSC.Close
			   Set RSC=Nothing
			  %>
			  &nbsp;&nbsp; <label><input onClick="$('#m').show()" type="radio" value="0" name="UserCardID">���ɳ�(��������������Ҫ��ֵ�Ľ��)</label><br/>
			  <span id='m' style="display:none"> &nbsp;&nbsp;&nbsp;&nbsp;��������Ҫ��ֵ�Ľ�&nbsp;<input style="text-align:center;line-height:22px" name="Money" type="text" class="textbox" value="100" size="10" maxlength="10"> Ԫ</span>
			  </td>
		    </tr>
			<tr class=tdbg>
			  <td align=middle colSpan=2 height=40>
		        <Input id="Action" type="hidden" value="PayStep2" name="Action"> 
				<Input class="button" id=Submit type=submit value=" ��һ�� " onClick="return(Confirm())" name=Submit> </td>
			</tr>
		  </table>
		</FORM>
		<br/><br/>
	   <%
	   End Sub
	   
	   Sub PayStep2()
	    Dim UserCardID:UserCardID=KS.ChkClng(KS.G("UserCardID"))
	   	Dim Money:Money=KS.S("Money")
		Dim Title
		
		If UserCardID<>0 Then
		   Dim RS:Set RS=Conn.Execute("Select Top 1 Money,GroupName From KS_UserCard Where ID=" & UserCardID)
		   If Not RS.Eof Then
		    Title=RS(1)
		    Money=RS(0)
			RS.Close : Set RS=Nothing
		   Else
		    RS.Close : Set RS=Nothing
		    Call KS.AlertHistory("��������",-1)
			Exit Sub 
		   End If
		Else
		   Title="Ϊ�Լ����˻���ֵ"
		End If

		If Not IsNumeric(Money) Then
		  Call KS.AlertHistory("�Բ���������ĳ�ֵ����ȷ��",-1)
		  exit sub
		End If
		
		If Money=0 Then
		  Call KS.AlertHistory("�Բ��𣬳�ֵ������Ϊ0.01Ԫ��",-1)
		  exit sub
		End If
		Dim OrderID:OrderID=KS.Setting(72) & Year(Now)&right("0"&Month(Now),2)&right("0"&Day(Now),2)&hour(Now)&minute(Now)&second(Now)
		
		%>
	   <FORM name=myform action="User_PayOnline.asp" method="post">
		  <table id="c1" class=border cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
			<tr class=title>
			  <td align=middle colSpan=2 height=22><B> ȷ �� �� ��</B></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=167>�û�����</td>
			  <td width="505"><%=KSUser.UserName%></td>
			</tr>
			<tr class=tdbg>
			  <td width="167" align=right>֧����ţ�</td>
			  <td><input type='hidden' value='<%=OrderID%>' name='OrderID'><%=OrderID%>&nbsp;</td>
		    </tr>
			<tr class=tdbg>
			  <td align=right width=167>֧����</td>
			  <td><input type='hidden' value='<%=Money%>' name='Money'><%=FormatNumber(Money,2,-1)%> Ԫ</td>
			</tr>
			<%If title<>"" then%>
			<tr class=tdbg>
			  <td align=right width=167>֧����;��</td>
			  <td style="color:red">��<%=title%>��</td>
			</tr>
			<%end if%>

			<tr class=tdbg>
			  <td align=right width=167>ѡ������֧��ƽ̨��</td>
			  <td>
			  <%
			   Dim SQL,K,Param
			   If UserCardID<>0 Then
			    Param=" and id in(1,10,7)"
			   End IF
			   Set RS=Server.CreateOBject("ADODB.RECORDSET")
			   RS.Open "Select ID,PlatName,Note,IsDefault From KS_PaymentPlat Where IsDisabled=1 " & Param & " Order By OrderID",conn,1,1
			   If Not RS.Eof Then SQL=RS.GetRows(-1)
			   RS.Close:Set RS=Nothing
			   If Not IsArray(SQL) Then
			    Response.Write "<font color='red'>�Բ��𣬱�վ�ݲ���ͨ����֧�����ܣ�</font>"
			   Else
			     For K=0 To Ubound(SQL,2)
				   Response.Write "<input type='radio' value='" & SQL(0,K) & "' name='PaymentPlat'"
				   If SQL(3,K)="1" Then Response.Write " checked"
				   Response.Write ">"& SQL(1,K) & "(" & SQL(2,K) &")<br>"
				 Next
			   End If
			  %>
			  </td>
			</tr>
			
			<tr class=tdbg>
			  <td align=middle colSpan=2 height=40>
		        <Input id=Action type=hidden value="PayStep3" name="Action"> 
		        <Input id=Action type=hidden value="<%=UserCardID%>" name="UserCardID"> 
		        <Input type=hidden value="user" name="PayFrom"> 
				<input class="button" type="button" value=" ��һ�� " onClick="javascript:history.back();"> 
				<Input class="button" id=Submit type=submit value=" ��һ�� " name=Submit>
				</td>
			</tr>
		  </table>
		</FORM>
		<%
	   End Sub
	   
	   
	   '֧���̳Ƕ���
	   Sub PayShopOrder()
	  	 Dim ID:ID=KS.ChkClng(KS.S("ID"))
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 OrderID,MoneyTotal,DeliverType From KS_Order Where ID="& ID,Conn,1,1
		 If RS.Eof Then
		  rs.close:set rs=nothing
		  KS.Die "<script>alert('������!');history.back();</script>"
		 End If 
		Dim OrderID:OrderID=RS("OrderID")
	   	Dim Money:Money=RS("MoneyTotal")
		Dim DeliverType:DeliverType=RS("DeliverType")
		RS.Close
		Dim DeliverName,ProductName
		RS.Open "Select Top 1 TypeName From KS_Delivery Where Typeid=" & DeliverType,conn,1,1
		If Not RS.Eof Then
		 DeliverName=RS(0)
		End IF
		RS.Close
		
		RS.Open "Select top 1 Title From KS_Product Where ID in(Select proid From KS_OrderItem Where OrderID='" & OrderID& "')",conn,1,1
		If RS.Eof And RS.Bof Then
		 ProductName=OrderID
		Else
			Do While Not RS.Eof
			 if ProductName="" Then
			   ProductName=rs(0)
			 Else
			   ProductName=ProductName&","&rs(0)
			 End If
			 RS.MoveNext
			Loop
		End If
		RS.Close
		
		If Not IsNumeric(Money) Then
		  Call KS.AlertHistory("�Բ��𣬶�������ȷ��",-1)
		  exit sub
		End If
		If Money=0 Then
		  Call KS.AlertHistory("�Բ��𣬶���������Ϊ0.01Ԫ��",-1)
		  exit sub
		End If
		%>
	   <FORM name=myform action="User_PayOnline.asp" method="post">
		  <table id="c1" class=border cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
			<tr class=title>
			  <td align=middle colSpan=2 height=22><B> ȷ �� �� ��</B></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=167>�û�����</td>
			  <td width="505"><%=KSUser.UserName%></td>
			</tr>
			<tr class=tdbg>
			  <td width="167" align=right>��Ʒ���ƣ�</td>
			  <td><input type='hidden' value='<%=ProductName%>' name='ProductName'><%=ProductName%>&nbsp;
			  <input type='hidden' value='<%=DeliverName%>' name='DeliverName'>
			  </td>
		    </tr>
			<tr class=tdbg>
			  <td width="167" align=right>֧����ţ�</td>
			  <td><input type='hidden' value='<%=OrderID%>' name='OrderID'><%=OrderID%>&nbsp;</td>
		    </tr>
			<tr class=tdbg>
			  <td align=right width=167>֧����</td>
			  <td><input type='hidden' value='<%=Money%>' name='Money'><%=Money%> Ԫ</td>
			</tr>
			
			<tr class=tdbg>
			  <td align=right width=167>ѡ������֧��ƽ̨��</td>
			  <td>
			  <%
			   Dim SQL,K
			   RS.Open "Select ID,PlatName,Note,IsDefault From KS_PaymentPlat Where IsDisabled=1 Order By OrderID",conn,1,1
			   If Not RS.Eof Then SQL=RS.GetRows(-1)
			   RS.Close:Set RS=Nothing
			   If Not IsArray(SQL) Then
			    Response.Write "<font color='red'>�Բ��𣬱�վ�ݲ���ͨ����֧�����ܣ�</font>"
			   Else
			     For K=0 To Ubound(SQL,2)
				   Response.Write "<input type='radio' value='" & SQL(0,K) & "' name='PaymentPlat'"
				   If SQL(3,K)="1" Then Response.Write " checked"
				   Response.Write ">"& SQL(1,K) & "(" & SQL(2,K) &")<br>"
				 Next
			   End If
			  %>
			  </td>
			</tr>
			
			<tr class=tdbg>
			  <td align=middle colSpan=2 height=40>
		        <Input id=Action type=hidden value="PayStep3" name="Action"> 
		        <Input type=hidden value="shop" name="PayFrom"> 
				<Input class="button" id=Submit type=submit value=" ��һ�� " name=Submit>
				<input class="button" type="button" value=" ��һ�� " onClick="javascript:history.back();"> </td>
			</tr>
		  </table>
		</FORM>
		<%
	   End Sub
	   
	   Sub PayStep3()
	    Dim UserCardID,Title
		UserCardID=KS.ChkClng(KS.S("UserCardID"))
	    Dim Money:Money=KS.S("Money")
		If UserCardID<>0 Then
		   Dim RS:Set RS=Conn.Execute("Select Top 1 Money,GroupName From KS_UserCard Where ID=" & UserCardID)
		   If Not RS.Eof Then
		    Title=RS(1)
		    Money=RS(0)
			RS.Close : Set RS=Nothing
		   Else
		    RS.Close : Set RS=Nothing
		    Call KS.AlertHistory("��������",-1)
			Exit Sub 
		   End If
		Else
		   Title="Ϊ�Լ����˻���ֵ"
		End If
		
		
		If Not IsNumeric(Money) Then
		  Call KS.AlertHistory("�Բ���������ĳ�ֵ����ȷ��",-1)
		  exit sub
		End If
		If Money=0 Then
		  Call KS.AlertHistory("�Բ��𣬳�ֵ������Ϊ0.01Ԫ��",-1)
		  exit sub
		End If

		Dim OrderID:OrderID=KS.S("OrderID")
		Dim PaymentPlat:PaymentPlat=KS.ChkClng(KS.S("PaymentPlat"))
		
		Dim RSP:Set RSP=Server.CreateObject("ADODB.RECORDSET")
		RSP.Open "Select * From KS_PaymentPlat where id=" & PaymentPlat,conn,1,1
		If RSP.Eof Then
		 RSP.Close:Set RSP=Nothing
		 Response.Write "Error!"
		 Response.End()
		End If
		Dim AccountID:AccountID=RSP("AccountID")
		Dim MD5Key:MD5Key=RSP("MD5Key")
		Dim PayOnlineRate:PayOnlineRate=RSP("Rate") 
		Dim RateByUser:RateByUser=KS.ChkClng(RSP("RateByUser")) 
		RSP.Close:Set RSP=Nothing
		
		Dim RealPayMoney:RealPayMoney=Money
		If RateByUser=1 Then
		  RealPayMoney=RealPayMoney+RealPayMoney*PayOnlineRate/100
		End If
		RealPayMoney=round(RealPayMoney,2)
		
		Dim PayUrl,PayMentField
		Dim v_amount,v_moneytype,v_md5info,v_oid,v_mid,v_url,remark1,remark2
		Dim ReturnUrl:ReturnUrl=KS.GetDomain & "user/User_PayReceive.asp?PaymentPlat=" & PaymentPlat &"&username=" & server.URLEncode(KSUser.userName) & "&action=" &KS.S("PayFrom")&"&usercardid=" & UserCardID   ' �̻��Զ��巵�ؽ���֧�������ҳ�� Receive.asp Ϊ����ҳ��
		remark1 = KSUser.UserName			            ' ��ע�ֶ�1
		remark2 = "���߳�ֵ��������Ϊ:" &OrderID		' ��ע�ֶ�2
		
		v_oid = OrderID
		v_amount=RealPayMoney
		v_moneytype="0"
		v_mid = AccountID
		v_url = ReturnUrl
		
		Dim v_ymd, v_hms
		v_ymd = Year(Date) & Right("0" & Month(Date), 2) & Right("0" & Day(Date), 2)
		v_hms = Right("0" & Hour(Time), 2) & Right("0" & Minute(Time), 2) & Right("0" & Second(Time), 2)
		Select Case PaymentPlat
		 Case 1 '��������
		  PayUrl="https://pay3.chinabank.com.cn/PayGate"
		  v_md5info=Ucase(trim(md5(v_amount&v_moneytype&v_oid&v_mid&v_url&MD5Key,32)))	'����֧��ƽ̨��MD5ֵֻ�ϴ�д�ַ���
	
		  PayMentField="<input type=""hidden"" name=""v_md5info"" value=""" & v_md5info &""">" & _
	                   "<input type=""hidden"" name=""v_mid""  value=""" & v_mid & """>" & _
	                   "<input type=""hidden"" name=""v_oid""  value=""" & v_oid & """>" & _
                  	   "<input type=""hidden"" name=""v_amount"" value=""" & v_amount & """>" & _
	                   "<input type=""hidden"" name=""v_moneytype"" value=""" & v_moneytype & """>" & _
                       "<input type=""hidden"" name=""v_url""  value=""" & v_url & """>" & _
                       "<!--���¼�����Ϊ����֧����ɺ���֧��������Ϣһͬ������Ϣ����ҳ���ڴ�����������ݲ���ı�,�磺Receive.asp -->" & _
                        "<input type=""hidden""  name=""remark2"" value=""" & remark2 & """>"

		 Case 2  '�й�����֧����
			PayUrl = "http://www.ipay.cn/4.0/bank.shtml"
			v_oid = cstr(Hour(Now) & Second(Now) & Minute(Now))	
			v_md5info = LCase(MD5(v_mid & v_oid & v_amount & KSUser.Email & KSUser.Mobile & MD5Key, 32))
			PayMentField = PayMentField & "<input type='hidden' name='v_mid' value='" & v_mid & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='v_oid' value='" & v_oid & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='v_amount' value='" & v_amount & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='v_email' value='" & KSUser.Email & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='v_mobile' value='" & KSUser.Mobile & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='v_md5'    value='" & v_md5info & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='v_url' value='" & v_url & "'>" & vbCrLf
		 Case 3  '�Ϻ���Ѹ
		    PayUrl = "https://www.ips.com.cn/ipay/ipayment.asp"
			v_mid = Right("000000" & v_mid, 6)
			PayMentField = PayMentField & "<input type='hidden' name='mer_code' value='" & v_mid & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='billNo' value='" & v_mid & v_hms & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='amount' value='" & v_amount & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='date' value='" & v_ymd & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='lang'  value='1'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='currency'   value='01'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='Merchanturl'   value='" & v_url & "'>" & vbCrLf
	     Case 4  '����֧��
			PayUrl = "http://www.yeepay.com/Pay/WestPayReceiveOrderFromMerchant.asp"
			PayMentField = PayMentField & "<input type='hidden' name='MerchantID' value='" & v_mid & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='OrderNumber' value='" & v_oid & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='OrderAmount' value='" & v_amount & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='PostBackURL' value='" & v_url & "'>" & vbCrLf
		 Case 5  '�׸�ͨ
			PayUrl = "http://pay.xpay.cn/Pay.aspx"
			v_md5info = LCase(MD5(MD5Key & ":" & v_amount & "," & v_oid & "," & v_mid & ",bank,,sell,,2.0", 32))
			PayMentField = PayMentField & "<input type='hidden' name='Tid' value='" & v_mid & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='Bid' value='" & v_oid & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='Prc' value='" & v_amount & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='url' value='" & v_url & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='Card' value='bank'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='Scard' value=''>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='ActionCode' value='sell'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='ActionParameter' value=''>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='Ver' value='2.0'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='Pdt' value='" & trim(KS.Setting(0)) & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='type' value=''>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='lang' value='gb2312'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='md' value='" & v_md5info & "'>" & vbCrLf
		Case 6   '����֧��
			PayUrl = "https://www.cncard.net/purchase/getorder.asp"
			v_md5info = LCase(MD5(v_mid & v_oid & v_amount & v_ymd & "01" & v_url & "00" & MD5Key, 32))
			PayMentField = PayMentField & "<input type='hidden' name='c_mid' value='" & v_mid & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='c_order' value='" & v_oid & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='c_orderamount' value='" & v_amount & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='c_ymd' value='" & v_ymd & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='c_moneytype' value='0'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='c_retflag' value='1'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='c_paygate' value=''>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='c_returl' value='" & v_url & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='c_memo1' value=''>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='c_memo2' value=''>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='c_language' value='0'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='notifytype' value='0'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='c_signstr' value='" & v_md5info & "'>" & vbCrLf
		 Case 7,9  '֧����
		    PayUrl="https://www.alipay.com/cooperate/gateway.do"
			Dim Partner
			Dim ArrMD5Key
			If InStr(MD5Key, "|") > 0 Then
				ArrMD5Key = Split(MD5Key, "|")
				If UBound(ArrMD5Key) = 1 Then
					Partner = ArrMD5Key(1)
					MD5Key = ArrMD5Key(0)
				End If
			End If
			
			Session("PayType")="ALIPAY"
			If PaymentPlat=7 Then
			    v_url=KS.GetDomain & "user/Alipay_NotifyUrl.asp?username=" & ksuser.username
				Dim myString:myString = "discount=0" & "&notify_url=" & v_url & "&out_trade_no=" & v_oid & "&partner=" & Partner & "&payment_type=1" & "&price=" & v_amount & "&quantity=1" & "&return_url=" & returnurl & "&seller_email=" & v_mid & "&service=create_direct_pay_by_user&subject=" & v_oid & MD5Key
				v_md5info = LCase(MD5(myString, 32))
				PayMentField = PayMentField & "<input type='hidden' name='discount' value='0'>" '��Ʒ�ۿ�
				PayMentField = PayMentField & "<input type='hidden' name='notify_url' value='" & v_url & "'>"
				PayMentField = PayMentField & "<input type='hidden' name='out_trade_no' value='" & v_oid & "'>"
				PayMentField = PayMentField & "<input type='hidden' name='payment_type' value='1'>"
				PayMentField = PayMentField & "<input type='hidden' name='partner' value='" & Partner & "'>"
				PayMentField = PayMentField & "<input type='hidden' name='price' value='" & v_amount & "'>"
				PayMentField = PayMentField & "<input type='hidden' name='quantity' value='1'>"
				PayMentField = PayMentField & "<input type='hidden' name='seller_email' value='" & v_mid & "'>"
				PayMentField = PayMentField & "<input type='hidden' name='service' value='create_direct_pay_by_user'>"
				PayMentField = PayMentField & "<input type='hidden' name='subject' value='" & v_oid & "'>"
				PayMentField = PayMentField & "<input type='hidden' name='sign' value='" & v_md5info & "'>"
				PayMentField = PayMentField & "<input type='hidden' name='sign_type' value='MD5'>"
				PayMentField = PayMentField & "<input type='hidden' name='return_url' value='" & returnurl & "'>"
		  Else
		        'returnurl=""
				'v_url=""
				
				Dim body
				Dim IsFabrication
				If KS.S("PayFrom")="shop" Then
				 IsFabrication = False
				 Body="֧����Ʒ����""" & V_oid & """�ķ���"
				Else
				 IsFabrication = True '�ʽ��ֵ,����������Ʒ
				 Body="""" & KS.Setting(0) & """�˻����߳�ֵ,������:" & v_oid
				End If
               ' If IsFabrication Then
               '     myString = LCase(MD5("notify_url=" & v_url & "&out_trade_no=" & v_oid & "&partner=" & Partner & "&price=" & v_amount & "&quantity=1" & "&return_url=" & returnurl & "&seller_email=" & v_mid & "&service=create_digital_goods_trade_p&subject=" & v_oid & MD5Key, 32))
              '  Else
                    myString = LCase(MD5("body=" & body & "&discount=0&logistics_fee=0&logistics_payment=BUYER_PAY&logistics_type=EXPRESS&out_trade_no=" & v_oid & "&partner=" & Partner & "&payment_type=1&price=" & v_amount & "&quantity=1&seller_email=" & v_mid & "&service=create_partner_trade_by_buyer&subject=" & v_oid & MD5Key, 32))
               ' End If
                 
				PayMentField = PayMentField & "<input type='hidden' name='body' value='" & body & "'>" & vbCrLf
				PayMentField = PayMentField & "<input type='hidden' name='discount' value='0'>" & vbCrLf
				PayMentField = PayMentField & "<input type='hidden' name='logistics_fee' value='0'>" & vbCrLf

				               
               ' If IsFabrication Then
              '      PayMentField = PayMentField & "<input type='hidden' name='service' value='create_digital_goods_trade_p'>" & vbCrLf
               ' Else
                    PayMentField = PayMentField & "<input type='hidden' name='logistics_payment' value='BUYER_PAY'>" & vbCrLf
                    PayMentField = PayMentField & "<input type='hidden' name='logistics_type' value='EXPRESS'>" & vbCrLf
                    PayMentField = PayMentField & "<input type='hidden' name='out_trade_no' value='" & v_oid & "'>" & vbCrLf
                    PayMentField = PayMentField & "<input type='hidden' name='partner' value='" & Partner & "'>" & vbCrLf
                    PayMentField = PayMentField & "<input type='hidden' name='payment_type' value='1'>" & vbCrLf
                    PayMentField = PayMentField & "<input type='hidden' name='price' value='" & v_amount & "'>" & vbCrLf
                    PayMentField = PayMentField & "<input type='hidden' name='quantity' value='1'>" & vbCrLf
                    PayMentField = PayMentField & "<input type='hidden' name='seller_email' value='" & v_mid & "'>" & vbCrLf
					PayMentField = PayMentField & "<input type='hidden' name='service' value='create_partner_trade_by_buyer'>" & vbCrLf
                    PayMentField = PayMentField & "<input type='hidden' name='subject' value='" & v_oid & "'>" & vbCrLf
                    PayMentField = PayMentField & "<input type='hidden' name='sign' value='" & myString & "'>" & vbCrLf
                    PayMentField = PayMentField & "<input type='hidden' name='sign_type' value='MD5'>" & vbCrLf
					
					
					
                    'PayMentField = PayMentField & "<input type='hidden' name='logistics_fee' value='0'>" & vbCrLf
               ' End If
                'PayMentField = PayMentField & "<input type='hidden' name='notify_url' value='" & v_url & "'>" & vbCrLf
                'PayMentField = PayMentField & "<input type='hidden' name='return_url' value='" & returnurl & "'>"

		  End If
			
		 Case 8  '��Ǯ֧��
			PayUrl = "https://www.99bill.com/gateway/recvMerchantInfoAction.htm"
			Dim OrderAmount,merchantAcctId, key, inputCharset, pageUrl, bgUrl, version, language, signType, payerName, payerContactType, payerContact
			Dim orderTime, productName, productNum, productId, productDesc, ext1, ext2, payType, bankId, redoFlag, pid, signMsgVal
			merchantAcctId = v_mid   '�����˻���
			key = MD5Key '������Կ
			inputCharset = "3" '1����UTF-8; 2����GBK; 3����gb2312
			pageUrl = v_url '����֧�������ҳ���ַ
			bgUrl = v_url '����������֧������ĺ�̨��ַ
			version = "v2.0" '���ذ汾.�̶�ֵ
			language = "1" '1�������ģ�2����Ӣ��
			signType = "1" '1����MD5ǩ��
			payerName = "" '֧��������
			payerContactType = "" '֧������ϵ��ʽ���� 1����Email��2�����ֻ���
			payerContact = "" '֧������ϵ��ʽ,ֻ��ѡ��Email���ֻ���
			orderId = v_oid '�̻�������
			OrderAmount = v_amount * 100 '�������,�Է�Ϊ��λ
			orderTime = v_ymd & v_hms '�����ύʱ��,14λ����
			productName = "" '��Ʒ����
			productNum = "" '��Ʒ����
			productId = "" '��Ʒ����
			productDesc = "" '��Ʒ����
			ext1 = "" '��չ�ֶ�1,��֧��������ԭ�����ظ��̻�
			ext2 = "" '��չ�ֶ�2
			payType = "00" '֧����ʽ,00�����֧��,��ʾ��Ǯ֧�ֵĸ���֧����ʽ,11���绰����֧��,12����Ǯ�˻�֧��,13������֧��,14��B2B֧��
			bankId = "" '���д���,ʵ��ֱ����ת������ҳ��ȥ֧��,�������μ� �ӿ��ĵ����д����б�,ֻ��payType=10ʱ�������ò���
			redoFlag = "1" 'ͬһ������ֹ�ظ��ύ��־:1����ͬһ������ֻ�����ύ1��,0��ʾͬһ��������û��֧���ɹ���ǰ���¿��ظ��ύ���
			pid = "" '��Ǯ�ĺ��������˻���
	
			signMsgVal = appendParam(signMsgVal, "inputCharset", inputCharset)
			signMsgVal = appendParam(signMsgVal, "pageUrl", pageUrl)
			signMsgVal = appendParam(signMsgVal, "bgUrl", bgUrl)
			signMsgVal = appendParam(signMsgVal, "version", version)
			signMsgVal = appendParam(signMsgVal, "language", language)
			signMsgVal = appendParam(signMsgVal, "signType", signType)
			signMsgVal = appendParam(signMsgVal, "merchantAcctId", merchantAcctId)
			signMsgVal = appendParam(signMsgVal, "payerName", payerName)
			signMsgVal = appendParam(signMsgVal, "payerContactType", payerContactType)
			signMsgVal = appendParam(signMsgVal, "payerContact", payerContact)
			signMsgVal = appendParam(signMsgVal, "orderId", v_oid)
			signMsgVal = appendParam(signMsgVal, "orderAmount", OrderAmount)
			signMsgVal = appendParam(signMsgVal, "orderTime", orderTime)
			signMsgVal = appendParam(signMsgVal, "productName", productName)
			signMsgVal = appendParam(signMsgVal, "productNum", productNum)
			signMsgVal = appendParam(signMsgVal, "productId", productId)
			signMsgVal = appendParam(signMsgVal, "productDesc", productDesc)
			signMsgVal = appendParam(signMsgVal, "ext1", ext1)
			signMsgVal = appendParam(signMsgVal, "ext2", ext2)
			signMsgVal = appendParam(signMsgVal, "payType", payType)
			signMsgVal = appendParam(signMsgVal, "bankId", bankId)
			signMsgVal = appendParam(signMsgVal, "redoFlag", redoFlag)
			signMsgVal = appendParam(signMsgVal, "pid", pid)
			signMsgVal = appendParam(signMsgVal, "key", key)
			v_md5info = UCase(MD5(signMsgVal, 32))
			PayMentField = PayMentField & "<input type='hidden' name='inputCharset' value='" & inputCharset & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='bgUrl' value='" & bgUrl & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='pageUrl' value='" & pageUrl & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='version' value='" & version & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='language' value='" & language & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='signType' value='" & signType & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='signMsg' value='" & v_md5info & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='merchantAcctId' value='" & merchantAcctId & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='payerName' value='" & payerName & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='payerContactType' value='" & payerContactType & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='payerContact' value='" & payerContact & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='orderId' value='" & orderId & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='orderAmount' value='" & OrderAmount & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='orderTime' value='" & orderTime & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='productName' value='" & productName & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='productNum' value='" & productNum & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='productId' value='" & productId & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='productDesc' value='" & productDesc & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='ext1' value='" & ext1 & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='ext2' value='" & ext2 & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='payType' value='" & payType & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='bankId' value='" & bankId & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='redoFlag' value='" & redoFlag & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='pid' value='" & pid & "'>" & vbCrLf
		Case 10 '�Ƹ�ͨ
		    Dim transaction_id
			transaction_id = v_mid & v_ymd & Right(v_oid, 10)
			PayUrl = "https://www.tenpay.com/cgi-bin/v1.0/pay_gate.cgi"
			v_md5info = UCase(MD5("cmdno=1&date=" & v_ymd & "&bargainor_id=" & v_mid & "&transaction_id=" & transaction_id & "&sp_billno=" & v_oid & "&total_fee=" & v_amount * 100 & "&fee_type=1&return_url=" & v_url & "&attach=my_magic_string&key=" & MD5Key, 32))
			PayMentField = PayMentField & "<input type='hidden' name='cmdno' value='1'>"   'ҵ�����,1��ʾ֧��
			PayMentField = PayMentField & "<input type='hidden' name='date' value='" & v_ymd & "'>"   '�̻�����
			PayMentField = PayMentField & "<input type='hidden' name='bank_type' value='0'>"  '��������:�Ƹ�ͨ,0
			PayMentField = PayMentField & "<input type='hidden' name='desc' value='" & KS.Setting(1) &"����֧����:" & v_oid & "'>"    '���׵���Ʒ����
			PayMentField = PayMentField & "<input type='hidden' name='purchaser_id' value=''>"   '�û�(��)�ĲƸ�ͨ�ʻ�,����Ϊ��
			PayMentField = PayMentField & "<input type='hidden' name='bargainor_id' value='" & v_mid & "'>"  '�̼ҵ��̻���
			PayMentField = PayMentField & "<input type='hidden' name='transaction_id' value='" & transaction_id & "'>"   '���׺�(������)
			PayMentField = PayMentField & "<input type='hidden' name='sp_billno' value='" & v_oid & "'>"  '�̻�ϵͳ�ڲ��Ķ�����
			PayMentField = PayMentField & "<input type='hidden' name='total_fee' value='" & v_amount * 100 & "'>" '�ܽ��Է�Ϊ��λ
			PayMentField = PayMentField & "<input type='hidden' name='fee_type' value='1'>"  '�ֽ�֧������,1�����
			PayMentField = PayMentField & "<input type='hidden' name='return_url' value='" & v_url & "'>" '���ղƸ�ͨ���ؽ����URL
			PayMentField = PayMentField & "<input type='hidden' name='attach' value='my_magic_string'>" '�̼����ݰ���ԭ������
			PayMentField = PayMentField & "<input type='hidden' name='sign' value='" & v_md5info & "'>" 'MD5ǩ��
		case 11 '�Ƹ�ͨ�н齻��
		    Dim mch_desc:mch_desc="���߹��򶩵���:" &v_oid
			Dim mch_name
			If Request("ProductName")<>"" Then 
			 mch_name=Request("ProductName")
		    Else
			 mch_name="���߹��򶩵���:" &v_oid
			End If
			Dim mch_price:mch_price=v_amount * 100
			Dim mch_returl:mch_returl=ReturnUrl
			Dim mch_type:mch_type=1
			Dim show_url:show_url=ReturnUrl
			Dim transport_desc:transport_desc=Request("DeliverName")
			
			PayUrl = "http://www.tenpay.com/cgi-bin/med/show_opentrans.cgi"
			dim buffer
					buffer = appendParam(buffer, "attach", 		"tencent_magichu")
					buffer = appendParam(buffer, "chnid", 			"1202640601")
					buffer = appendParam(buffer, "cmdno", 			"12")
					buffer = appendParam(buffer, "encode_type", 	"1")
					buffer = appendParam(buffer, "mch_desc", 		mch_desc)
					buffer = appendParam(buffer, "mch_name", 		mch_name)
					buffer = appendParam(buffer, "mch_price", 		mch_price)
					buffer = appendParam(buffer, "mch_returl", 	mch_returl)
					buffer = appendParam(buffer, "mch_type", 		mch_type)
					buffer = appendParam(buffer, "mch_vno", 		v_oid)
					buffer = appendParam(buffer, "need_buyerinfo", "2")
					buffer = appendParam(buffer, "seller", 		v_mid)
					buffer = appendParam(buffer, "show_url", 		show_url)
					buffer = appendParam(buffer, "transport_desc", transport_desc)
					buffer = appendParam(buffer, "transport_fee", 	0)
					buffer = appendParam(buffer, "version", 		2)
					
			        buffer = appendParam(buffer, "key", 			MD5Key)
					
			v_md5info=MD5(buffer,32)
					
			PayMentField = PayMentField & "<input type='hidden' name='attach' value='tencent_magichu'>" '�̼����ݰ���ԭ������
			PayMentField = PayMentField & "<input type='hidden' name='chnid' value='1202640601'>" 'ƽ̨�ṩ�ߵĲƸ�ͨ�˺�
			PayMentField = PayMentField & "<input type='hidden' name='cmdno' value='12'>"   'ҵ�����,1��ʾ֧��
			PayMentField = PayMentField & "<input type='hidden' name='encode_type' value='1'>"   '����
			PayMentField = PayMentField & "<input type='hidden' name='mch_desc' value='" & mch_desc&"'>"   '����˵��
			PayMentField = PayMentField & "<input type='hidden' name='mch_name' value='" & mch_name&"'>"   '��Ʒ����
			PayMentField = PayMentField & "<input type='hidden' name='mch_price' value='"&mch_price&"'>"   '��Ʒ�۸�
			PayMentField = PayMentField & "<input type='hidden' name='mch_returl' value='"&mch_returl&"'>"   '�ص�֪ͨURL,���cmdnoΪ12�Ҵ��ֶ���д��Ч�ص�����,�Ƹ�ͨ���ѽ��������Ϣ֪ͨ����URL 
			PayMentField = PayMentField & "<input type='hidden' name='mch_type' value='"&mch_type&"'>"   '�������ͣ�1��ʵ�ｻ�ף�2�����⽻��
			PayMentField = PayMentField & "<input type='hidden' name='mch_vno' value='"&v_oid&"'>"   '������
			PayMentField = PayMentField & "<input type='hidden' name='need_buyerinfo' value='2'>"   '�Ƿ���Ҫ�ڲƸ�ͨ�������Ϣ��1����Ҫ��2������Ҫ��
			PayMentField = PayMentField & "<input type='hidden' name='seller' value='" & v_mid & "'>"   '�տ�Ƹ�ͨ�˺�
			PayMentField = PayMentField & "<input type='hidden' name='show_url' value='"&show_url&"'>"   '֧������̻�֧�����չʾҳ��
			PayMentField = PayMentField & "<input type='hidden' name='transport_desc' value='"&transport_desc&"'>"   '������Ϣ
			PayMentField = PayMentField & "<input type='hidden' name='transport_fee' value='0'>"   '������֧�������������Ѱ�������Ʒ�۸��У�����д0��������Ĭ��Ϊ0����λΪ��
			PayMentField = PayMentField & "<input type='hidden' name='version' value='2'>"   
			PayMentField = PayMentField & "<input type='hidden' name='sign' value='"&v_md5info&"'>"   
		End Select  

		
		 %>
	   	  <FORM name="myform"  id="myform" action="<%=PayUrl%>" <%if PaymentPlat=11 or PaymentPlat=9 then response.write "method=""get""" else response.write "method=""post"""%>  target="_blank">
		  <table id="c1" class=border cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
			<tr class=title>
			  <td align=middle colSpan=2 height=22><B> ȷ �� �� ��</B></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=167>�û�����</td>
			  <td width="505"><%=KSUser.UserName%></td>
			</tr>
			<tr class=tdbg>
			  <td width="167" align=right>֧����ţ�</td>
			  <td><%=OrderID%>&nbsp;</td>
		    </tr>
			<tr class=tdbg>
			  <td align=right width=167>֧����</td>
			  <td><%=formatnumber(Money,2,-1)%> Ԫ</td>
			</tr>
			<%if title<>"" then%>
			<tr class=tdbg>
			  <td align=right width=167>֧����;��</td>
			  <td style="color:red">��<%=title%>��</td>
			</tr>
			<%end if%>
			<%
			if RateByUser=1 then
			%>
			<tr class=tdbg>
			  <td align=right width=167>�����ѣ�</td>
			  <td><%=PayOnlineRate%>%</td>
			</tr>
			<%end if%>
			<tr class=tdbg>
			  <td align=right width=167>ʵ��֧����</td>
			  <td>
			  <%=formatnumber(RealPayMoney,2,-1)%></td>
			</tr>
			<tr class=tdbg>
			  <td colspan=2>�����ȷ��֧������ť�󣬽���������֧�����棬�ڴ�ҳ��ѡ���������п���</td>
		    </tr>
			<tr class=tdbg>
			  <td align=middle colSpan=2 height=40>
			    <%=PayMentField%>
				<%if PaymentPlat=9 then%>
				<Input class="button" id=Submit type=button onClick="$('#myform').submit()" value=" ȷ��֧�� " onClick="document.all.c1.style.display='none';document.all.c2.style.display='';">
				<%else%>
				<Input class="button" id=Submit type=submit value=" ȷ��֧�� " onClick="document.all.c1.style.display='none';document.all.c2.style.display='';">
				<%end if%>
				<input class="button" type="button" value=" ��һ�� " onClick="javascript:history.back();"> </td>
			</tr>
		  </table>
		</FORM>
		  <table id="c2" style="display:none" class=border cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
			<tr class=title>
			  <td align=middle height=22><B> ȷ �� �� ��</B></td>
			</tr>
			<tr class=tdbg>
			  <td align=center height="150">�밴ҳ����ʾ�������ֵ��</td>
			</tr>
          </table>
	   <%
	   End Sub
	  
	  '������ֵ��Ϊ�յĲ�������ַ���(��Ǯ)
		Function appendParam(returnStr, paramId, paramValue)
			If returnStr <> "" Then
				If paramValue <> "" Then
					returnStr=returnStr&"&"&paramId&"="&paramValue
				End If
			Else
				If paramValue <> "" Then
					returnStr=paramId&"="&paramValue
				End If
			End If
			appendParam = returnStr
		End Function
		
		

		
End Class
%> 
