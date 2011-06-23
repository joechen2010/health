<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
'********************************
'* 程序功能：订单提交成功
'* 演示地址: http://wap.kesion.com/
'********************************
%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<% Response.ContentType="text/vnd.wap.wml" %><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="订单提交成功">
<p>
<%
Dim KSCls
Set KSCls = New ShoppingCart
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%
Class ShoppingCart
        Private KS,DomainStr
		Private ProductList
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		    DomainStr=KS.GetDomain
		End Sub
        Private Sub Class_Terminate()
		    Call CloseConn()
		    Set KS=Nothing
		End Sub
		Public Sub Kesion()
		    Dim FileContent,Products,i,RS,strsql,CartStr,OrderAutoID
			ProductList = KS.FilterIDs(Session("ProductList"))
			If ProductList="" Then
			   Response.Write "您的购物车中没有商品！<br/><br/>" &vbcrlf
			   Response.Write "<a href="""&KS.GetGoBackIndex&""">返回首页</a><br/>" &vbcrlf
			   Exit Sub
			End IF
			If ProductList="" Then
			   Response.Write "再次刷新显示该网页,需要重新发送你以前提交的信息!还回<a href="""&KS.GetGoBackIndex&""">继续购物</a><br/>"
			   Exit Sub
			End IF
			'入订单
			'生成订单号
			Dim OrderID:OrderID=KS.Setting(71) & Year(Now)&right("0"&Month(Now),2)&right("0"&Day(Now),2)&KS.MakeRandom(8)
			Dim ContactMan:ContactMan=KS.S("ContactMan")
			Dim Address:Address=KS.S("Address")
			Dim ZipCode:ZipCode=KS.S("ZipCode")
			Dim Phone:Phone=KS.S("Phone")
			Dim Email:Email=KS.S("Email")
			Dim Mobile:Mobile=KS.S("Mobile")
			Dim QQ:QQ=KS.S("QQ")
			Dim PaymentType:PaymentType=KS.S("PaymentType")
			Dim DeliverType:DeliverType=KS.S("DeliverType")
			Dim InvoiceContent:InvoiceContent=KS.S("InvoiceContent")
			Dim NeedInvoice:NeedInvoice=KS.ChkClng(KS.S("NeedInvoice"))
			Dim Remark:Remark=KS.S("Remark")
			
			Dim RSA,RealPrice,MoneyGoods,RealMoneyTotal
			Set RS=Server.CreateObject("ADODB.RecordSet") 
			RS.Open "select * from KS_Product where ID in ("&ProductList&") order by ID",Conn,1,1
			Do While Not RS.eof
			   Set RSA=Server.CreateObject("ADODB.RecordSet")
			   RSA.Open"select * from KS_OrderItem where ID is null",Conn,1,3
			   RSA.AddNew
			   RSA("OrderID")=OrderID
			   RSA("ProID")=RS("ID")
			   RSA("SaleType")=RS("ProductType")
			   RSA("Price_Original")=RS("Price_Original")
			   RSA("Price")=RS("Price")
			   IF Cbool(KSUser.UserLoginChecked)=true Then
			      If RS("GroupPrice")=0 Then
				     RealPrice=RS("Price_Member")
				  Else
				     Dim RSP:Set RSP=Conn.Execute("Select Price From KS_ProPrice Where GroupID=" & KSUser.GroupID & " And ProID=" & RS("ID"))
					 If RSP.Eof Then
						RealPrice=RS("Price_Member")
					 Else
						RealPrice=RSP(0)
					 End If
					 RSP.Close:Set RSP=Nothing
				  End If
			   Else
			      RealPrice=RS("Price")
			   End If
			   RSA("RealPrice")=RealPrice
			   RSA("Amount")=Session("Amount"&RS("ID"))
			   RSA("TotalPrice")=Round(RealPrice*Session("Amount"&RS("ID")),2)
			   RSA("BeginDate")=Now
			   RSA("ServiceTerm")=RS("ServiceTerm")
			   RSA.Update
			   RSA.Close:Set RSA=Nothing
			   MoneyGoods=MoneyGoods+Round(RealPrice*Session("Amount"&RS("ID")),2)
			   Session("Amount"&RS("ID"))=""
			   RS.MoveNext
			Loop
			RS.Close
			
			'实际支付金额。
			Dim PaymentDiscount:PayMentDiscount=KS.ReturnPayment(PaymentType,1)
			Dim DeliveryMoney:DeliveryMoney=KS.ReturnDelivery(DeliverType,1)
			Dim TaxRate:TaxRate=KS.Setting(65)
			Dim IncludeTax:IncludeTax=KS.Setting(64)
			Dim TaxMoney
			If IncludeTax=1 Or NeedInvoice=0 Then TaxMoney=1 Else TaxMoney=1+Taxrate/100
			'总金额 = (总价*付费方式折扣+运费)*(1+税率)
			RealMoneyTotal=Round((MoneyGoods*PayMentDiscount/100+DeliveryMoney)*TaxMoney,2)
			
			RS.Open "Select * From KS_Order",Conn,1,3
			RS.AddNew
			RS("OrderID")=OrderID
			If Cbool(KSUser.UserLoginChecked)=true Then
			   RS("UserName")= KSUser.UserName
			Else
			   RS("UserName") = "游客"
			End If
			RS("MoneyTotal")=RealMoneyTotal
			RS("MoneyGoods")=MoneyGoods
			RS("NeedInvoice")=NeedInvoice
			RS("InvoiceContent")=InvoiceContent
			RS("Remark")=Remark
			RS("InputTime")=Now
			RS("ContactMan")=ContactMan
			RS("Address")=Address
			RS("ZipCode")=ZipCode
			RS("Mobile")=Mobile
			RS("Phone")=Phone
			RS("QQ")=QQ
			RS("Email")=Email
			RS("PaymentType")=PaymentType
			RS("DeliverType")=DeliverType
			RS("Discount_Payment")=PaymentDiscount   '付款方式折扣率
			RS("Charge_Deliver")=DeliveryMoney     '运费
			
			'相关初始值
			RS("Invoiced")=0       '发票未开
			RS("MoneyReceipt")=0   '已收款
			RS("BeginDate")=Now    '开始服务日期
			RS("Status")=0         '订单状态
			RS("DeliverStatus")=0  '送货状态
			RS("PresentMoney")=0       '返回客户现金
			RS("PresentPoint")=0       '返回客户点券
			RS("PresentScore")=0       '返回客户积分
			RS.Update
			RS.MoveLast
			OrderAutoID=RS("id")
			RS.Close:Set RS=Nothing
			Session("ProductList")=""  '交易成功！置购物车参数为空
			
			Response.Write "恭喜您,您的订单提交成功!只有付款成功后,才能完成本次交易哦.<br/>" &vbcrlf
			Response.Write "您的订单号码是:"&OrderID&"<br/>" &vbcrlf
			Response.Write "本次交易金额为:"&RealMoneyTotal&"元<br/><br/>" &vbcrlf
			Response.Write "<a href=""../plus/Template.asp?ID=20089643230062&amp;"&KS.WapValue&""">网银支付</a><br/>" &vbcrlf
			Response.Write "<a href=""../plus/Template.asp?ID=20082668859466&amp;"&KS.WapValue&""">汇款方式</a><br/>" &vbcrlf
			If Cbool(KSUser.UserLoginChecked)=True Then
			   Response.Write "<a href=""../User/User_Order.asp?Action=AddPayment&amp;ID="&OrderAutoID&"&amp;"&KS.WapValue&""">余额支付</a><br/><br/>" &vbcrlf
            End iF
            Response.Write "<a href="""&KS.GetGoBackIndex&""">返回首页</a>" &vbcrlf
        End Sub
	   
End Class
%>