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
		Call KSUser.InnerLocation("在线支付")
		Response.Write "<div class=""tabs"">"
		Response.Write " <ul class="""">"
		Response.Write " <li class='select'><a href=""User_PayOnline.asp"">在线支付充值</a></li>"
		Response.Write " <li><a href=""user_recharge.asp"">充值卡充值</a></li>"
		Response.Write " <li><a href=""user_exchange.asp?Action=Point"">兑换" & KS.Setting(45) & "</a></li>"
		Response.Write " <li><a href=""user_exchange.asp?Action=Edays"">兑换有效期</a></li>"
		Response.Write " <li><a href=""user_exchange.asp?Action=Money"">" & KS.Setting(45) & "兑换账户资金</a></li>"
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
		   alert('请输入你要充值的金额!')
		   document.myform.Money.focus();
		   return false;
		  }
		  return true;
		  }
	   </script>
		<FORM name=myform action="User_PayOnline.asp" method="post">
		  <table class=border cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
			<tr class=title>
			  <td align=middle colSpan=2 height=22><B> 在 线 充 值</B></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=213>用户名：</td>
			  <td width="754"><%=KSUser.UserName%></td>
			</tr>
			<tr class=tdbg>
			  <td width="213" align=right>计费方式：</td>
			  <td><%if KSUser.ChargeType=1 Then 
		  Response.Write "扣点数</font>计费用户"
		  ElseIf KSUser.ChargeType=2 Then
		   Response.Write "有效期</font>计费用户,到期时间：" & cdate(KSUser.BeginDate)+KSUser.Edays & ","
		  ElseIf KSUser.ChargeType=3 Then
		   Response.Write "无限期</font>计费用户"
		  End If
		  %>&nbsp;</td>
		    </tr>
			<tr class=tdbg>
			  <td align=right width=213>资金余额：</td>
			  <td><input type='hidden' value='<%=KSUser.Money%>' name='Premoney'><%=formatnumber(KSUser.Money,2,-1)%> 元</td>
			</tr>
			<%If KSUser.ChargeType=1 then%>
			<tr class=tdbg>
			  <td align=right width=213>可用<%=KS.Setting(45)%>：</td>
			  <td><%=KSUser.Point%>&nbsp;<%=KS.Setting(46)%></td>
			</tr>
			<%end if%>
			<%If KSUser.ChargeType=2 then%>
			<tr class=tdbg>
			  <td align=right width=213>剩余天数：</td>
			  <td>
			  <%if KSUser.ChargeType=3 Then%>
			  无限期
			  <%else%>
			  <%=KSUser.GetEdays%>&nbsp;天
			  <%end if%></td>
			</tr>
		   <%end if%>
			<tr class=tdbg>
			  <td align=right>当前级别：</td>
			  <td><%=KS.U_G(KSUser.GroupID,"groupname")%></td>
		    </tr>
			<tr class=title>
			  <td align=middle colSpan=2 height=22><B> 选 择 在 线 充 值 方 式</B></td>
			</tr>

			<tr class=tdbg>
			  <td colspan="2">
			  <%
			   Dim RSC,AllowGroupID:Set RSC=Conn.Execute("Select ID,GroupName,Money,AllowGroupID From KS_UserCard Where CardType=1 and DateDiff(" & DataPart_S & ",EndDate," & SqlNowString& ")<0")
			   Do While NOt RSC.Eof 
			      AllowGroupID=RSC("AllowGroupID") : If IsNull(AllowGroupID) Then AllowGroupID=" "
			     If KS.IsNul(AllowGroupID) Or KS.FoundInArr(AllowGroupID,KSUser.GroupID,",")=true Then
			    response.write "&nbsp;&nbsp; <label><input checked name=""UserCardID"" onclick=""$('#m').hide()"" type=""radio"" value=""" & rsc("ID") & """/>" & rsc(1) & " (需要花费 <span style='color:red'>" & formatnumber(RSC(2),2,-1) & "</span> 元)</label><br/>"
				End If
			    RSC.MoveNext
			   Loop
			   RSC.Close
			   Set RSC=Nothing
			  %>
			  &nbsp;&nbsp; <label><input onClick="$('#m').show()" type="radio" value="0" name="UserCardID">自由充(您可以任意输入要充值的金额)</label><br/>
			  <span id='m' style="display:none"> &nbsp;&nbsp;&nbsp;&nbsp;请输入你要充值的金额：&nbsp;<input style="text-align:center;line-height:22px" name="Money" type="text" class="textbox" value="100" size="10" maxlength="10"> 元</span>
			  </td>
		    </tr>
			<tr class=tdbg>
			  <td align=middle colSpan=2 height=40>
		        <Input id="Action" type="hidden" value="PayStep2" name="Action"> 
				<Input class="button" id=Submit type=submit value=" 下一步 " onClick="return(Confirm())" name=Submit> </td>
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
		    Call KS.AlertHistory("出错啦！",-1)
			Exit Sub 
		   End If
		Else
		   Title="为自己的账户充值"
		End If

		If Not IsNumeric(Money) Then
		  Call KS.AlertHistory("对不起，您输入的充值金额不正确！",-1)
		  exit sub
		End If
		
		If Money=0 Then
		  Call KS.AlertHistory("对不起，充值金额最低为0.01元！",-1)
		  exit sub
		End If
		Dim OrderID:OrderID=KS.Setting(72) & Year(Now)&right("0"&Month(Now),2)&right("0"&Day(Now),2)&hour(Now)&minute(Now)&second(Now)
		
		%>
	   <FORM name=myform action="User_PayOnline.asp" method="post">
		  <table id="c1" class=border cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
			<tr class=title>
			  <td align=middle colSpan=2 height=22><B> 确 认 款 项</B></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=167>用户名：</td>
			  <td width="505"><%=KSUser.UserName%></td>
			</tr>
			<tr class=tdbg>
			  <td width="167" align=right>支付编号：</td>
			  <td><input type='hidden' value='<%=OrderID%>' name='OrderID'><%=OrderID%>&nbsp;</td>
		    </tr>
			<tr class=tdbg>
			  <td align=right width=167>支付金额：</td>
			  <td><input type='hidden' value='<%=Money%>' name='Money'><%=FormatNumber(Money,2,-1)%> 元</td>
			</tr>
			<%If title<>"" then%>
			<tr class=tdbg>
			  <td align=right width=167>支付用途：</td>
			  <td style="color:red">“<%=title%>”</td>
			</tr>
			<%end if%>

			<tr class=tdbg>
			  <td align=right width=167>选择在线支付平台：</td>
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
			    Response.Write "<font color='red'>对不起，本站暂不开通在线支付功能！</font>"
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
				<input class="button" type="button" value=" 上一步 " onClick="javascript:history.back();"> 
				<Input class="button" id=Submit type=submit value=" 下一步 " name=Submit>
				</td>
			</tr>
		  </table>
		</FORM>
		<%
	   End Sub
	   
	   
	   '支付商城订单
	   Sub PayShopOrder()
	  	 Dim ID:ID=KS.ChkClng(KS.S("ID"))
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 OrderID,MoneyTotal,DeliverType From KS_Order Where ID="& ID,Conn,1,1
		 If RS.Eof Then
		  rs.close:set rs=nothing
		  KS.Die "<script>alert('出错啦!');history.back();</script>"
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
		  Call KS.AlertHistory("对不起，订单金额不正确！",-1)
		  exit sub
		End If
		If Money=0 Then
		  Call KS.AlertHistory("对不起，订单金额最低为0.01元！",-1)
		  exit sub
		End If
		%>
	   <FORM name=myform action="User_PayOnline.asp" method="post">
		  <table id="c1" class=border cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
			<tr class=title>
			  <td align=middle colSpan=2 height=22><B> 确 认 款 项</B></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=167>用户名：</td>
			  <td width="505"><%=KSUser.UserName%></td>
			</tr>
			<tr class=tdbg>
			  <td width="167" align=right>商品名称：</td>
			  <td><input type='hidden' value='<%=ProductName%>' name='ProductName'><%=ProductName%>&nbsp;
			  <input type='hidden' value='<%=DeliverName%>' name='DeliverName'>
			  </td>
		    </tr>
			<tr class=tdbg>
			  <td width="167" align=right>支付编号：</td>
			  <td><input type='hidden' value='<%=OrderID%>' name='OrderID'><%=OrderID%>&nbsp;</td>
		    </tr>
			<tr class=tdbg>
			  <td align=right width=167>支付金额：</td>
			  <td><input type='hidden' value='<%=Money%>' name='Money'><%=Money%> 元</td>
			</tr>
			
			<tr class=tdbg>
			  <td align=right width=167>选择在线支付平台：</td>
			  <td>
			  <%
			   Dim SQL,K
			   RS.Open "Select ID,PlatName,Note,IsDefault From KS_PaymentPlat Where IsDisabled=1 Order By OrderID",conn,1,1
			   If Not RS.Eof Then SQL=RS.GetRows(-1)
			   RS.Close:Set RS=Nothing
			   If Not IsArray(SQL) Then
			    Response.Write "<font color='red'>对不起，本站暂不开通在线支付功能！</font>"
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
				<Input class="button" id=Submit type=submit value=" 下一步 " name=Submit>
				<input class="button" type="button" value=" 上一步 " onClick="javascript:history.back();"> </td>
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
		    Call KS.AlertHistory("出错啦！",-1)
			Exit Sub 
		   End If
		Else
		   Title="为自己的账户充值"
		End If
		
		
		If Not IsNumeric(Money) Then
		  Call KS.AlertHistory("对不起，您输入的充值金额不正确！",-1)
		  exit sub
		End If
		If Money=0 Then
		  Call KS.AlertHistory("对不起，充值金额最低为0.01元！",-1)
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
		Dim ReturnUrl:ReturnUrl=KS.GetDomain & "user/User_PayReceive.asp?PaymentPlat=" & PaymentPlat &"&username=" & server.URLEncode(KSUser.userName) & "&action=" &KS.S("PayFrom")&"&usercardid=" & UserCardID   ' 商户自定义返回接收支付结果的页面 Receive.asp 为接收页面
		remark1 = KSUser.UserName			            ' 备注字段1
		remark2 = "在线充值，订单号为:" &OrderID		' 备注字段2
		
		v_oid = OrderID
		v_amount=RealPayMoney
		v_moneytype="0"
		v_mid = AccountID
		v_url = ReturnUrl
		
		Dim v_ymd, v_hms
		v_ymd = Year(Date) & Right("0" & Month(Date), 2) & Right("0" & Day(Date), 2)
		v_hms = Right("0" & Hour(Time), 2) & Right("0" & Minute(Time), 2) & Right("0" & Second(Time), 2)
		Select Case PaymentPlat
		 Case 1 '网银在线
		  PayUrl="https://pay3.chinabank.com.cn/PayGate"
		  v_md5info=Ucase(trim(md5(v_amount&v_moneytype&v_oid&v_mid&v_url&MD5Key,32)))	'网银支付平台对MD5值只认大写字符串
	
		  PayMentField="<input type=""hidden"" name=""v_md5info"" value=""" & v_md5info &""">" & _
	                   "<input type=""hidden"" name=""v_mid""  value=""" & v_mid & """>" & _
	                   "<input type=""hidden"" name=""v_oid""  value=""" & v_oid & """>" & _
                  	   "<input type=""hidden"" name=""v_amount"" value=""" & v_amount & """>" & _
	                   "<input type=""hidden"" name=""v_moneytype"" value=""" & v_moneytype & """>" & _
                       "<input type=""hidden"" name=""v_url""  value=""" & v_url & """>" & _
                       "<!--以下几项项为网上支付完成后，随支付反馈信息一同传给信息接收页，在传输过程中内容不会改变,如：Receive.asp -->" & _
                        "<input type=""hidden""  name=""remark2"" value=""" & remark2 & """>"

		 Case 2  '中国在线支付网
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
		 Case 3  '上海环迅
		    PayUrl = "https://www.ips.com.cn/ipay/ipayment.asp"
			v_mid = Right("000000" & v_mid, 6)
			PayMentField = PayMentField & "<input type='hidden' name='mer_code' value='" & v_mid & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='billNo' value='" & v_mid & v_hms & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='amount' value='" & v_amount & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='date' value='" & v_ymd & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='lang'  value='1'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='currency'   value='01'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='Merchanturl'   value='" & v_url & "'>" & vbCrLf
	     Case 4  '西部支付
			PayUrl = "http://www.yeepay.com/Pay/WestPayReceiveOrderFromMerchant.asp"
			PayMentField = PayMentField & "<input type='hidden' name='MerchantID' value='" & v_mid & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='OrderNumber' value='" & v_oid & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='OrderAmount' value='" & v_amount & "'>" & vbCrLf
			PayMentField = PayMentField & "<input type='hidden' name='PostBackURL' value='" & v_url & "'>" & vbCrLf
		 Case 5  '易付通
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
		Case 6   '云网支付
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
		 Case 7,9  '支付宝
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
				PayMentField = PayMentField & "<input type='hidden' name='discount' value='0'>" '商品折扣
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
				 Body="支付商品订单""" & V_oid & """的费用"
				Else
				 IsFabrication = True '资金充值,当做虚拟物品
				 Body="""" & KS.Setting(0) & """账户在线充值,订单号:" & v_oid
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
			
		 Case 8  '快钱支付
			PayUrl = "https://www.99bill.com/gateway/recvMerchantInfoAction.htm"
			Dim OrderAmount,merchantAcctId, key, inputCharset, pageUrl, bgUrl, version, language, signType, payerName, payerContactType, payerContact
			Dim orderTime, productName, productNum, productId, productDesc, ext1, ext2, payType, bankId, redoFlag, pid, signMsgVal
			merchantAcctId = v_mid   '网关账户号
			key = MD5Key '网关密钥
			inputCharset = "3" '1代表UTF-8; 2代表GBK; 3代表gb2312
			pageUrl = v_url '接受支付结果的页面地址
			bgUrl = v_url '服务器接受支付结果的后台地址
			version = "v2.0" '网关版本.固定值
			language = "1" '1代表中文；2代表英文
			signType = "1" '1代表MD5签名
			payerName = "" '支付人姓名
			payerContactType = "" '支付人联系方式类型 1代表Email；2代表手机号
			payerContact = "" '支付人联系方式,只能选择Email或手机号
			orderId = v_oid '商户订单号
			OrderAmount = v_amount * 100 '订单金额,以分为单位
			orderTime = v_ymd & v_hms '订单提交时间,14位数字
			productName = "" '商品名称
			productNum = "" '商品数量
			productId = "" '商品代码
			productDesc = "" '商品描述
			ext1 = "" '扩展字段1,在支付结束后原样返回给商户
			ext2 = "" '扩展字段2
			payType = "00" '支付方式,00：组合支付,显示快钱支持的各种支付方式,11：电话银行支付,12：快钱账户支付,13：线下支付,14：B2B支付
			bankId = "" '银行代码,实现直接跳转到银行页面去支付,具体代码参见 接口文档银行代码列表,只在payType=10时才需设置参数
			redoFlag = "1" '同一订单禁止重复提交标志:1代表同一订单号只允许提交1次,0表示同一订单号在没有支付成功的前提下可重复提交多次
			pid = "" '快钱的合作伙伴的账户号
	
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
		Case 10 '财付通
		    Dim transaction_id
			transaction_id = v_mid & v_ymd & Right(v_oid, 10)
			PayUrl = "https://www.tenpay.com/cgi-bin/v1.0/pay_gate.cgi"
			v_md5info = UCase(MD5("cmdno=1&date=" & v_ymd & "&bargainor_id=" & v_mid & "&transaction_id=" & transaction_id & "&sp_billno=" & v_oid & "&total_fee=" & v_amount * 100 & "&fee_type=1&return_url=" & v_url & "&attach=my_magic_string&key=" & MD5Key, 32))
			PayMentField = PayMentField & "<input type='hidden' name='cmdno' value='1'>"   '业务代码,1表示支付
			PayMentField = PayMentField & "<input type='hidden' name='date' value='" & v_ymd & "'>"   '商户日期
			PayMentField = PayMentField & "<input type='hidden' name='bank_type' value='0'>"  '银行类型:财付通,0
			PayMentField = PayMentField & "<input type='hidden' name='desc' value='" & KS.Setting(1) &"在线支付号:" & v_oid & "'>"    '交易的商品名称
			PayMentField = PayMentField & "<input type='hidden' name='purchaser_id' value=''>"   '用户(买方)的财付通帐户,可以为空
			PayMentField = PayMentField & "<input type='hidden' name='bargainor_id' value='" & v_mid & "'>"  '商家的商户号
			PayMentField = PayMentField & "<input type='hidden' name='transaction_id' value='" & transaction_id & "'>"   '交易号(订单号)
			PayMentField = PayMentField & "<input type='hidden' name='sp_billno' value='" & v_oid & "'>"  '商户系统内部的定单号
			PayMentField = PayMentField & "<input type='hidden' name='total_fee' value='" & v_amount * 100 & "'>" '总金额，以分为单位
			PayMentField = PayMentField & "<input type='hidden' name='fee_type' value='1'>"  '现金支付币种,1人民币
			PayMentField = PayMentField & "<input type='hidden' name='return_url' value='" & v_url & "'>" '接收财付通返回结果的URL
			PayMentField = PayMentField & "<input type='hidden' name='attach' value='my_magic_string'>" '商家数据包，原样返回
			PayMentField = PayMentField & "<input type='hidden' name='sign' value='" & v_md5info & "'>" 'MD5签名
		case 11 '财付通中介交易
		    Dim mch_desc:mch_desc="在线购买订单号:" &v_oid
			Dim mch_name
			If Request("ProductName")<>"" Then 
			 mch_name=Request("ProductName")
		    Else
			 mch_name="在线购买订单号:" &v_oid
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
					
			PayMentField = PayMentField & "<input type='hidden' name='attach' value='tencent_magichu'>" '商家数据包，原样返回
			PayMentField = PayMentField & "<input type='hidden' name='chnid' value='1202640601'>" '平台提供者的财付通账号
			PayMentField = PayMentField & "<input type='hidden' name='cmdno' value='12'>"   '业务代码,1表示支付
			PayMentField = PayMentField & "<input type='hidden' name='encode_type' value='1'>"   '编码
			PayMentField = PayMentField & "<input type='hidden' name='mch_desc' value='" & mch_desc&"'>"   '交易说明
			PayMentField = PayMentField & "<input type='hidden' name='mch_name' value='" & mch_name&"'>"   '商品名称
			PayMentField = PayMentField & "<input type='hidden' name='mch_price' value='"&mch_price&"'>"   '商品价格
			PayMentField = PayMentField & "<input type='hidden' name='mch_returl' value='"&mch_returl&"'>"   '回调通知URL,如果cmdno为12且此字段填写有效回调链接,财付通将把交易相关信息通知给此URL 
			PayMentField = PayMentField & "<input type='hidden' name='mch_type' value='"&mch_type&"'>"   '交易类型：1、实物交易，2、虚拟交易
			PayMentField = PayMentField & "<input type='hidden' name='mch_vno' value='"&v_oid&"'>"   '订单号
			PayMentField = PayMentField & "<input type='hidden' name='need_buyerinfo' value='2'>"   '是否需要在财付通填定物流信息，1：需要，2：不需要。
			PayMentField = PayMentField & "<input type='hidden' name='seller' value='" & v_mid & "'>"   '收款方财付通账号
			PayMentField = PayMentField & "<input type='hidden' name='show_url' value='"&show_url&"'>"   '支付后的商户支付结果展示页面
			PayMentField = PayMentField & "<input type='hidden' name='transport_desc' value='"&transport_desc&"'>"   '物流信息
			PayMentField = PayMentField & "<input type='hidden' name='transport_fee' value='0'>"   '需买方另支付的物流费如已包含在商品价格中，请填写0。如果不填，默认为0。单位为分
			PayMentField = PayMentField & "<input type='hidden' name='version' value='2'>"   
			PayMentField = PayMentField & "<input type='hidden' name='sign' value='"&v_md5info&"'>"   
		End Select  

		
		 %>
	   	  <FORM name="myform"  id="myform" action="<%=PayUrl%>" <%if PaymentPlat=11 or PaymentPlat=9 then response.write "method=""get""" else response.write "method=""post"""%>  target="_blank">
		  <table id="c1" class=border cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
			<tr class=title>
			  <td align=middle colSpan=2 height=22><B> 确 认 款 项</B></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=167>用户名：</td>
			  <td width="505"><%=KSUser.UserName%></td>
			</tr>
			<tr class=tdbg>
			  <td width="167" align=right>支付编号：</td>
			  <td><%=OrderID%>&nbsp;</td>
		    </tr>
			<tr class=tdbg>
			  <td align=right width=167>支付金额：</td>
			  <td><%=formatnumber(Money,2,-1)%> 元</td>
			</tr>
			<%if title<>"" then%>
			<tr class=tdbg>
			  <td align=right width=167>支付用途：</td>
			  <td style="color:red">“<%=title%>”</td>
			</tr>
			<%end if%>
			<%
			if RateByUser=1 then
			%>
			<tr class=tdbg>
			  <td align=right width=167>手续费：</td>
			  <td><%=PayOnlineRate%>%</td>
			</tr>
			<%end if%>
			<tr class=tdbg>
			  <td align=right width=167>实际支付金额：</td>
			  <td>
			  <%=formatnumber(RealPayMoney,2,-1)%></td>
			</tr>
			<tr class=tdbg>
			  <td colspan=2>点击“确认支付”按钮后，将进入在线支付界面，在此页面选择您的银行卡。</td>
		    </tr>
			<tr class=tdbg>
			  <td align=middle colSpan=2 height=40>
			    <%=PayMentField%>
				<%if PaymentPlat=9 then%>
				<Input class="button" id=Submit type=button onClick="$('#myform').submit()" value=" 确定支付 " onClick="document.all.c1.style.display='none';document.all.c2.style.display='';">
				<%else%>
				<Input class="button" id=Submit type=submit value=" 确定支付 " onClick="document.all.c1.style.display='none';document.all.c2.style.display='';">
				<%end if%>
				<input class="button" type="button" value=" 上一步 " onClick="javascript:history.back();"> </td>
			</tr>
		  </table>
		</FORM>
		  <table id="c2" style="display:none" class=border cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
			<tr class=title>
			  <td align=middle height=22><B> 确 认 款 项</B></td>
			</tr>
			<tr class=tdbg>
			  <td align=center height="150">请按页面提示完成最后充值！</td>
			</tr>
          </table>
	   <%
	   End Sub
	  
	  '将变量值不为空的参数组成字符串(快钱)
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
