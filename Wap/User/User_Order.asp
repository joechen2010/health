<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<% Response.ContentType="text/vnd.wap.wml" %><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="我的订单">
<p>
<%
Dim KSCls
Set KSCls = New User_Order
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%
Class User_Order
        Private KS
		Private CurrentPage,totalPut
		Private RS,MaxPerPage
		Private TempStr,SqlStr,Prev
		Private InfoIDArr,InfoID,DomainStr
		Private Sub Class_Initialize()
			MaxPerPage =5
		    Set KS=New PublicCls
			DomainStr=KS.GetDomain
		End Sub
        Private Sub Class_Terminate()
		    Set KS=Nothing
			Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		    IF Cbool(KSUser.UserLoginChecked)=False Then
			   Response.redirect DomainStr&"User/Login/"
			   Exit Sub
			End If
			Select Case KS.S("Action")
			    Case "ShowOrder" Call ShowOrder
				Case "DelOrder" Call DelOrder
				Case "SignUp"  Call SignUp
				Case "AddPayment"  Call AddPayment '从账户余额付款
				Case "SavePayment"  Call SavePayment
				Case "PaymentOnline"  '在线支付
				'Response.Redirect "User_PayOnline.asp?Action=Payonline&id=" & KS.S("ID")
				Case Else Call OrderList
			End Select
			If Prev=True Then
			   Response.Write "<anchor>返回上一页<prev/></anchor><br/>"
			End If
			Response.Write "<br/>"
			Response.Write "<a href=""Index.asp?" & KS.WapValue & """>我的地盘</a><br/>" &vbcrlf
			Response.Write "<a href=""" & KS.GetGoBackIndex & """>返回首页</a>" &vbcrlf
		End Sub
		
		Sub OrderList()
		    If KS.S("page") <> "" Then
			   CurrentPage = CInt(KS.S("page"))
			Else
			   CurrentPage = 1
			End If
			Response.Write "【我的订单】<br/>"
			SqlStr="Select * From KS_Order Where UserName='" & KSUser.UserName &"' order by id desc"
			Set RS=Server.createobject("adodb.recordset")
			RS.open SqlStr,Conn,1,1
			If RS.EOF And RS.BOF Then
			   Response.Write "您没有下任何订单!<br/>"
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
		    Call  KS.ShowPageParamter(totalPut, MaxPerPage, "User_Order.asp", True, "条订单", CurrentPage, KS.WapValue)
		End Sub
		
		Sub ShowContent()
		    Dim i,MoneyTotal,MoneyReceipt
			Do While Not RS.Eof
			%>
            编号:<a href="User_Order.asp?Action=ShowOrder&amp;ID=<%=RS("ID")%>&amp;<%=KS.WapValue%>"><%=rs("orderid")%></a><br/>
			时间:<%=rs("inputtime")%><br/>
			金额:<%=formatnumber(rs("MoneyTotal"),2)%> 收款金额:<%=formatnumber(RS("MoneyReceipt"),2)%><br/>
            <%
			If RS("NeedInvoice")=1 Then
			   If RS("Invoiced")=1 Then
			      Response.Write "发票已开,"
			   Else
			      Response.Write "发票未开,"
			   End If
			End If
			If RS("Status")=0 Then
			   Response.Write "等待确认"
			ElseIf RS("Status")=1 Then
			   Response.Write "已经确认"
			ElseIf RS("Status")=2 Then
			   Response.Write "已结清"
			End If
			Response.Write ","
			If RS("MoneyReceipt")<=0 Then
			   Response.Write "等待汇款"
			ElseIf RS("MoneyReceipt")<RS("MoneyTotal") Then
			   Response.WRITE "已收定金"
			Else
			   Response.Write "已经付清"
			End If
			Response.Write ","
			If RS("DeliverStatus")=0 Then
			   Response.Write "未发货"
			ElseIf RS("DeliverStatus")=1 Then
			   Response.Write "已发货"
			ElseIf RS("DeliverStatus")=2 Then
			   Response.Write "已签收"
			End If
			Response.Write "<br/>"
			
			MoneyReceipt=RS("MoneyReceipt")+MoneyReceipt
			MoneyTotal=RS("MoneyTotal")+MoneyTotal
			Response.Write "<img src=""../Images/Hen.gif"" alt=""""/><br/>"
			RS.MoveNext
			I = I + 1
			If I >= MaxPerPage Then Exit Do
			Loop
			
			Response.Write "<br/>"
			Response.Write "本页-合计:"&formatnumber(MoneyTotal,2)&" "
			Response.Write "收款金额:"&formatnumber(MoneyReceipt,2)&"<br/>"
			Response.Write "所有-总计:"&formatnumber(Conn.execute("Select sum(moneytotal) from KS_Order Where UserName='" & KSUser.UserName & "'")(0),2)&" "
			Response.Write "收款金额:"&formatnumber(Conn.execute("Select sum(MoneyReceipt) from KS_Order Where UserName='" & KSUser.UserName & "'")(0),2)&"<br/>"
		End Sub
		
		Sub ShowOrder()
		    Dim ID:ID=KS.ChkClng(KS.S("ID"))
			Dim RS:Set RS=Server.CreateObject("adodb.recordset")
			RS.Open "Select * from ks_order where id=" & ID ,Conn,1,1
			IF RS.Eof And RS.Bof Then
			   RS.Close:Set RS=Nothing
			   Response.Write "参数错误!<br/>"
			   Prev=True
			   Exit Sub
		    End If
			'返回订单详细信息
			Response.Write "【订单信息】<br/>" &vbcrlf
			Response.Write "订单编号:" & RS("ORDERID") & "<br/>"&vbcrlf
			Response.Write "客户姓名:" & RS("Contactman") & "<br/>" &vbcrlf
			Response.Write "用 户 名:" & RS("UserName") & "<br/>" &vbcrlf
			Response.Write "【代 理 商】<br/>" &vbcrlf
			Response.Write "购买日期:" & formatdatetime(RS("inputtime"),2) & "<br/>" & vbcrlf
			Response.Write "下单时间:" & RS("inputtime") & "<br/>" & vbcrlf
			If RS("NeedInvoice")=1 Then
			   Response.Write "发票状态:"
			   If RS("Invoiced")=1 Then
				  Response.Write "已开"
			   Else
			      Response.Write "未开"
			   End If
			   Response.Write "<br/>" & vbcrlf
			End If
			
			Response.Write "订单状态:"	
			If RS("Status")=0 Then
			   Response.Write "等待确认"
			ElseIf RS("Status")=1 Then
			   Response.Write "已经确认"
			ElseIf RS("Status")=2 Then
			   Response.Write "已结清"
			End If
			Response.Write "<br/>" & vbcrlf
			Response.Write "付款情况:"
			If RS("MoneyReceipt")<=0 Then
			   Response.Write "等待汇款"
			ElseIf RS("MoneyReceipt")<RS("MoneyTotal") Then
			   Response.Write "已收定金"
			Else
			   Response.Write "已经付清"
			End If
			Response.Write "<br/>" & vbcrlf
			Response.Write "物流状态:"
			If RS("DeliverStatus")=0 Then
			   Response.Write "未发货"
			ElseIf RS("DeliverStatus")=1 Then
			   Response.Write "已发货"
			ElseIf RS("DeliverStatus")=2 Then
			   Response.Write "已签收"
			End If
			Response.Write "<br/>" & vbcrlf
			Response.Write "<br/>" & vbcrlf
			Response.Write "收货姓名:" & RS("contactman") & "<br/>" & vbcrlf
			Response.Write "联系电话:" & RS("phone") & "<br/>" & vbcrlf
			Response.Write "收货地址:" & RS("address") & "<br/>" & vbcrlf
			Response.Write "邮政编码:" & RS("zipcode") & "<br/>" & vbcrlf
			Response.Write "收货邮箱:" & RS("email") & "<br/>" & vbcrlf
			Response.Write "收货手机:" & RS("mobile") & "<br/>" & vbcrlf
			Response.Write "付款方式:" & KS.ReturnPayMent(rs("PaymentType"),0) & "<br/>" & vbcrlf
			Response.Write "送货方式:" & KS.ReturnDelivery(rs("DeliverType"),0) & "<br/>" & vbcrlf
			If RS("Invoiced")=1 Then
			   Response.Write "发票信息:" & RS("InvoiceContent") & "<br/>" & vbcrlf
			End If
			Response.Write "备注留言:" & RS("Remark") & "<br/>" & vbcrlf
			
			Response.Write "<img src=""../Images/Hen.gif"" alt=""""/><br/>"
			Dim TotalPrice,RSI:Set RSI=Server.CreateObject("ADODB.RECORDSET")
			RSI.Open "Select * From KS_OrderItem Where OrderID='" & RS("OrderID") & "'",conn,1,1
			If RSI.Eof Then
			   RSI.Close:Set RSI=Nothing
			   Response.Write "找不到相关商品!<br/>" & vbcrlf
			Else	
			   Do While Not RSI.EOF
			      Response.Write "商品名称:<a href='" & DomainStr & "Show.asp?ID=" & RSI("proid") & "&amp;ChannelID=5&amp;"&KS.WapValue&"'>" & Conn.Execute("select title from ks_product where id=" & RSI("proid"))(0) & "</a><br/>" & vbcrlf
				  Response.Write "数量单位:" & RSI("Amount") &""& Conn.Execute("select unit from ks_product where id=" & RSI("proid"))(0) & " "
				  Response.Write "期限:" & RSI("ServiceTerm") & "<br/>"
				  Response.Write "原价:" & formatnumber(RSI("price_original"),2) & " "
				  Response.Write "实价:" & formatnumber(RSI("price"),2) & " "
				  Response.Write "指定价:" & formatnumber(RSI("realprice"),2) & " "
				  Response.Write "金额:" & formatnumber(RSI("realprice")*RSI("amount"),2) & "<br/>"
				  'Response.Write "备注:" & RSI("Remark") & "<br/>"
				  TotalPrice=TotalPrice+ RSI("realprice")*RSI("Amount")
				  Response.Write "<img src=""../Images/Hen.gif"" alt=""""/><br/>"
				  RSI.Movenext
			   loop
			   Response.Write "合计:" & formatnumber(totalprice,2) & "<br/>"
			End If
			RSI.Close:set RSI=Nothing
			Response.Write "付款方式折扣率：" & RS("Discount_Payment") & "% 运费：" & RS("Charge_Deliver")&" 元 税率：" & KS.Setting(65) &"% 价格含税："
			IF KS.Setting(64)=1 Then 
			   Response.Write "是"
			Else
			   Response.Write "不含税"
			End If
			Response.Write "<br/>" & vbcrlf
			Dim TaxMoney
			Dim TaxRate:TaxRate=KS.Setting(65)
			If KS.Setting(64)=1 Or rs("NeedInvoice")=0 Then TaxMoney=1 Else TaxMoney=1+TaxRate/100
			Response.Write "实际金额：(" & rs("MoneyGoods") & "×" & rs("Discount_Payment") & "%＋"&rs("Charge_Deliver") & ")×"
			If KS.Setting(64)=1 Or rs("NeedInvoice")=0 Then Response.Write "100%" Else Response.Write "(1＋" & TaxRate & "%)" 
			Response.Write "＝" & formatnumber(rs("MoneyTotal"),2) & "元<br/>"
			Response.Write "实际金额：￥" & formatnumber(rs("MoneyTotal"),2) & " "
			Response.Write "已付款：￥" & formatnumber(rs("MoneyReceipt"),2) & " "
			If RS("MoneyReceipt")<RS("MoneyTotal") Then
			   Response.Write "尚欠款：￥" & formatnumber(RS("MoneyTotal")-RS("MoneyReceipt"),2) &""
			End If
			Response.Write "<br/>" & vbcrlf
			
			
			Response.Write "注：“原价”指商品的原始零售价，“实价”指系统自动计算出来的商品最终价格，“指定价”指管理员根据不同会员组手动指定的最终价格。商品的最终销售价格以“指定价”为准。<br/>"
			Response.Write "<br/>"
			If rs("status")=0 And rs("DeliverStatus")=0 And rs("MoneyReceipt")=0 Then
			   Response.Write "<a href=""User_Order.asp?Action=DelOrder&amp;ID=" & rs("id") & "&amp;" & KS.WapValue & """>删除订单</a> "
			End If
			If RS("MoneyReceipt")<RS("MoneyTotal") Then
			   Response.Write "<a href=""User_Order.asp?Action=AddPayment&amp;ID=" & rs("id") & "&amp;" & KS.WapValue & """>从余额中扣款支付</a> "
			End If
			If rs("DeliverStatus")=1 Then
			   Response.Write "<a href=""User_Order.asp?Action=SignUp&amp;ID=" & RS("ID") & "&amp;" & KS.WapValue & """>签收商品</a> "
			End If
			Response.Write "<a href=""User_Order.asp?" & KS.WapValue & """>订单首页</a><br/>"
		End Sub

		'删除订单
		Sub DelOrder()
		    Dim ID:ID=KS.ChkClng(KS.S("ID"))
			If KS.S("Checked")="ok" Then
			   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			   RS.Open "Select OrderID From KS_Order where status=0 and DeliverStatus=0 and MoneyReceipt=0 and id=" & ID,Conn,1,3
			   If Not RS.EOF Then
			      Conn.Execute("delete from ks_orderitem Where OrderID='" & rs(0) &"'")
				  RS.Delete
			   End if
			   Response.redirect DomainStr&"User/User_Order.asp?" & KS.WapValue & ""
			Else
			    Response.Write "确定要删除此订单吗？"
				Response.Write "<a href=""User_Order.asp?Action=DelOrder&amp;Checked=ok&amp;ID=" & ID & "&amp;" & KS.WapValue & """>确定</a> "
				Response.Write "<a href=""User_Order.asp?Action=ShowOrder&amp;ID=" & ID & "&amp;" & KS.WapValue & """>取消</a><br/>"
			End if
		End Sub
		
		'签收商品
		Sub SignUp()
		    Dim OrderID,id:ID=KS.ChkClng(KS.S("ID"))
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select * From KS_Order Where ID=" & ID,Conn,1,3
			If RS.EOF Then
			   RS.Close:Set RS=Nothing
			   Response.Write "出错啦!<br/>"
			   Prev=True
			   Exit Sub
		    End If
			RS("DeliverStatus")=2
			RS("BeginDate")=Now
			RS.Update
			OrderID=RS("OrderID")
			RS.Close:Set RS=Nothing
			Conn.Execute("Update KS_LogDeliver Set Status=1 Where OrderID='" & OrderID & "'")
			Response.redirect DomainStr&"User/User_Order.asp?Action=ShowOrder&ID="&ID&"&"&KS.WapValue&""
		End Sub
		
		Sub AddPayment()
		    Dim ID:ID=KS.ChkClng(KS.S("ID"))
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select * From KS_Order Where ID="& ID,Conn,1,1
			If RS.EOF Then
			   RS.Close:Set RS=Nothing
			   Response.Write "出错啦!<br/>"
			   Prev=True
			   Exit Sub
		    End If
			%>
            <b>注意：支出信息一旦录入，就不能再修改！所以在保存之前确认输入无误！</b><br/><br/>
            使用账户资金支付订单<br/>
            用 户 名：<%=KSUser.UserName%><br/>
            客户名称：<%=RS("ContactMan")%><br/>
            资金余额：<%=KSUser.Money%>元<br/>
            支付内容：<br/>
            订单编号：<%=RS("OrderID")%><br/>
            订单金额：<%=RS("MoneyTotal")%>元<br/>
            已 付 款：<%=RS("MoneyReceipt")%>元<br/>
            支付成功后，将从您的资金余额中扣除相应款项。<br/>
            支出金额：<input maxLength="20" size="10" value="<%=RS("moneytotal")-RS("MoneyReceipt")%>" name="Money<%=Minute(Now)%><%=Second(Now)%>" />元<br/>
            
            <anchor>确认支付<go href="User_Order.asp?<%=KS.WapValue%>" method="post">
            <postfield name="Action" value="SavePayment"/>
            <postfield name="ID" value="<%=RS("id")%>"/>
            <postfield name="Money" value="$(Money<%=Minute(Now)%><%=Second(Now)%>)"/>
            </go></anchor>            
            <br/>
            <%
		    RS.Close:Set RS=Nothing
		End Sub
		
		'开始余额支付操作
		Sub SavePayment()
		    Dim ID:ID=KS.ChkClng(KS.S("ID"))
			Dim Money:Money=KS.S("Money")
			If Not IsNumeric(Money) Then
			   Response.Write "请输入有效的金额!<br/>"
			   Prev=True
			   Exit Sub
			End If
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select * From KS_Order Where ID=" & ID,Conn,1,3
			If RS.EOF Then
			   RS.Close:Set RS=Nothing
			   Response.Write "出错啦!<br/>"
			   Prev=True
			   Exit Sub
		    End If
			If Round(Money)>Round(KSUser.Money) Then
			   Response.Write "错误信息<br/>"
			   Response.Write "产生错误的可能原因：您输入的支付金额超过了您的资金余额，无效支付！<br/>"
			   Prev=True
			   RS.Close:Set RS=Nothing:Exit Sub
		    End If
			'从账户资金中扣除金额
			Dim CurrMoney:CurrMoney=0
			Dim RSU:Set RSU=Server.CreateObject("adodb.recordset")
			RSU.Open "Select Money From KS_User Where UserName='" & RS("UserName") & "'",Conn,1,3
			If Not RSU.Eof Then
		       RSU(0)=RSU(0)-Money
			   RSU.Update
			   CurrMoney=rsu(0)
		    End If
			RSU.Close:Set RSU=Nothing
			
			RS("MoneyReceipt")=RS("MoneyReceipt")+Money
			RS("Status")=1
			RS.Update
			
			'写入资金明细
			Dim RSLog:Set RSLog=Server.CreateObject("ADODB.RECORDSET")
			RSLog.Open "Select * From KS_LogMoney",Conn,1,3
			RSLog.AddNew
		    RSLog("UserName")=RS("UserName")
			RSLog("ClientName")=RS("Contactman")
			RSLog("Money")=Money
			RSLog("MoneyType")=4       '余额支付
			RSLog("IncomeOrPayOut")=2  '支出
			RSLog("OrderID")=RS("OrderID") 
			RSLog("Remark")="支付订单费用，订单号：" & RS("Orderid")
			RSLog("PayTime")=Now
			RSLog("LogTime")=Now
			RSLog("Inputer")=KSUser.UserName
			RSLog("IP")=KS.GetIP
			RSLog("CurrMoney")=CurrMoney
			RSLog("ChannelID")=0
			RSLog("InfoID")=0
			RSLog.Update
			RSLog.Close:Set RSLog=Nothing
			'====================为用户增加购物应得积分========================
			If RS("MoneyReceipt")>=RS("MoneyTotal") Then
			   Dim RSP:set RSP=Conn.Execute("select point,id from ks_product where id in(select proid from KS_OrderItem where orderid='" & rs("orderid") & "')")
			   do while not rsp.eof
			      Dim Amount:Amount=Conn.Execute("select amount from ks_orderitem where orderid='" & RS("orderid") & "' and proid=" & RSP(1))(0)
				  Conn.Execute("update ks_user set score=score+" & KS.ChkClng(rsp(0))*amount & " where username='" & ksuser.username & "'")
				  RSP.MoveNext
			   Loop
			   RSP.Close:set RSP=Nothing
			End If
			'================================================================
			RS.Close:Set RS=Nothing
			Response.redirect DomainStr&"User/User_Order.asp?Action=ShowOrder&ID="&ID&"&"&KS.WapValue&""
		End Sub
End Class
%> 
