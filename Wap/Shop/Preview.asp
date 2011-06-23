<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
'********************************
'* 程序功能：订单确认
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
<card id="main" title="订单确认">
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
		    Dim FileContent,Products,i,RS,strsql,CartStr
			ProductList = KS.FilterIDs(Session("ProductList"))
			
			If Cbool(KSUser.UserLoginChecked)=False Then
			   Response.Write "温馨提示：您还没有注册或登录。享受更多会员优惠，请先<a href=""User/Login/?../PayMent.asp"">登录</a>或<a href=""User/Reg/?../PayMent.asp"">注册</a>成为商城会员！<br/>" &vbcrlf
			Else
			   Response.Write "亲爱的"&KSUser.UserName&"<br/>" &vbcrlf
			   Response.Write "【个人信息】<br/>" &vbcrlf
			   Response.Write "用户组:"&KS.GetUserGroupName(KSUser.GroupID)&"<br/>" &vbcrlf
			   Response.Write "可用资金:" & KSUser.Money & "元 " & KS.Setting(45) & ":" & KSUser.Point & "" & KS.Setting(46)&" 积分:" & KSUser.Score & "分<br/>" &vbcrlf
			   Response.Write "<br/>" &vbcrlf
			End iF
			
			Set RS=Server.CreateObject("ADODB.RecordSet") 
			If ProductList<>"" Then
			   strsql="select ID,Title,ProductType,Price_Original,Price,Price_Member,Discount,TotalNum,GroupPrice from KS_Product where ID in ("&ProductList&") order by ID"
			Else
			   strsql="select ID,Title,ProductType,Price_Original,Price,Price_Member,Discount,TotalNum,GroupPrice from KS_Product where 1=0"
			End If
			RS.Open strsql,Conn,1,1
			Dim TotalPrice,RealPrice,Price_Original,Discount,Amount,RealTotalPrice
			Amount = 1
			Do While Not RS.EOF
			   Amount = KS.ChkClng(KS.S( "Q_" & RS("ID")))
			   If Amount <= 0 Then 
			      Amount = KS.ChkClng(Session("Amount"&RS("ID")))
			      If Amount <= 0 Then Amount = 1
			   End If
			   Session("Amount"&RS("ID")) = Amount
			   IF RS("TotalNum") < Amount Then
			      Amount = 1
				  Session("Amount"&RS("ID")) = 1
				  Response.Write "对不起，"&RS("Title")&"暂时库存不足，请过段时间再来购买该商品！<br/><br/>"
				  Response.Write "<a href="""&KS.GetGoBackIndex&""">返回首页</a><br/>" &vbcrlf
				  Exit Sub
			   End IF
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
			   TotalPrice=TotalPrice+Round(RealPrice*Amount,2)
			   Response.Write "编号:"&RS("ID")&" 名称:"&RS("Title")&"<br/>" &vbcrlf
			   Response.Write "数量:"&Amount&" 原价:￥"&RS("Price_Original")&" 折扣:"&RS("Discount")&" 实价:￥"&RealPrice&" 总计:￥"&Round(RealPrice*Amount,2)&"<br/><br/>" &vbcrlf
			   RS.MoveNext
		    Loop
		    RS.close:set RS=nothing
			Dim PaymentDiscount:PayMentDiscount=KS.ReturnPayment(KS.S("PaymentType"),1)
			Dim DeliveryMoney:DeliveryMoney=KS.ReturnDelivery(KS.S("DeliverType"),1)
			Dim TaxRate:TaxRate=KS.Setting(65)
			Dim IncludeTax:IncludeTax=KS.Setting(64)
			Dim TaxMoney
			If IncludeTax=1 Or KS.ChkClng(KS.S("NeedInvoice"))=0 Then TaxMoney=1 Else TaxMoney=1+Taxrate/100
			TotalPrice=Round(TotalPrice,2)
			'总金额 = (总价*付费方式折扣+运费)*(1+税率)
			RealTotalPrice=Round((TotalPrice*PayMentDiscount/100+DeliveryMoney)*TaxMoney,2)
			Response.Write "付款方式折扣率:"&PayMentDiscount&"% 运费:"&DeliveryMoney&"元 " &vbcrlf
			Response.Write "价格含税:" &vbcrlf
			If IncludeTax=1 Then Response.Write "是 " Else Response.Write "不含税 "
			Response.Write "税率:"&TaxRate&"%<br/>" &vbcrlf
			Response.Write "实际金额：("&TotalPrice&"×"&PaymentDiscount&"%＋"&DeliveryMoney&")×" &vbcrlf
			If IncludeTax=1 Or KS.ChkClng(KS.S("NeedInvoice"))=0 Then Response.Write "100%" Else Response.Write "(1+" & TaxRate & "%)"
			Response.Write "＝" & RealTotalPrice&" 元<br/>" &vbcrlf
			Response.Write "合计:￥"&RealTotalPrice&"元<br/><br/>" &vbcrlf
			
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
			If NeedInvoice="0" Then 
			   InvoiceContent="不需要发票"
			Else
			   InvoiceContent="需要开据发票，发票信息如下:"&InvoiceContent
			End If
			
			IF ContactMan="" Then
			   Response.Write "请输入收货人姓名！<br/><br/>" &vbcrlf
			   Response.Write "<anchor><prev/>还回上级</anchor><br/>" &vbcrlf
			   Exit Sub
			End IF
			IF Address="" Then
			   Response.Write "请输入收货人地址！<br/><br/>" &vbcrlf
			   Response.Write "<anchor><prev/>还回上级</anchor><br/>" &vbcrlf
			   Exit Sub
			End IF
			IF ZipCode="" Then
			   Response.Write "请输入收货人邮编！<br/><br/>" &vbcrlf
			   Response.Write "<anchor><prev/>还回上级</anchor><br/>" &vbcrlf
			   Exit Sub
			End IF
			IF Phone="" Then
			   Response.Write "请输入收货人电话！<br/><br/>" &vbcrlf
			   Response.Write "<anchor><prev/>还回上级</anchor><br/>" &vbcrlf
			   Exit Sub
			End IF
			IF Email="" Then
			   Response.Write "请输入收货人的Email！<br/><br/>" &vbcrlf
			   Response.Write "<anchor><prev/>还回上级</anchor><br/>" &vbcrlf
			   Exit Sub
			Else
			   If KS.IsValidEmail(Email)=false Then
			      Response.Write "请输入正确的电子邮箱！<br/><br/>" &vbcrlf
				  Response.Write "<anchor><prev/>还回上级</anchor><br/>" &vbcrlf
				  Exit Sub
				End If
			End IF
	
			Response.Write "请认真核对你的收货信息<br/>" &vbcrlf
			Response.Write "收货人姓名:"&ContactMan&"<br/>" &vbcrlf
            Response.Write "收货人地址:"&Address&"<br/>" &vbcrlf
            Response.Write "收货人邮编:"&ZipCode&"<br/>" &vbcrlf
            Response.Write "收货人电话:"&Phone&"<br/>" &vbcrlf
            Response.Write "收货人邮箱:"&Email&"<br/>" &vbcrlf
            Response.Write "收货人手机:"&Mobile&"<br/>" &vbcrlf
            Response.Write "收货人ＱＱ:"&QQ&"<br/>" &vbcrlf
            Response.Write "付款方式:"&KS.ReturnPayment(PaymentType,0)&"<br/>" &vbcrlf
            Response.Write "送货方式:"&KS.ReturnDelivery(DeliverType,0)&"<br/>" &vbcrlf
            Response.Write "发票信息:"&InvoiceContent&"<br/>" &vbcrlf
            Response.Write "备注留言:"&Remark&"<br/>" &vbcrlf
            Response.Write "<anchor>确认提交订单<go href=""Order.asp?"&KS.WapValue&""" method=""post"">" &vbcrlf
            Response.Write "<postfield name=""ContactMan"" value="""&ContactMan&"""/>" &vbcrlf
            Response.Write "<postfield name=""Address"" value="""&Address&"""/>" &vbcrlf
            Response.Write "<postfield name=""ZipCode"" value="""&ZipCode&"""/>" &vbcrlf
            Response.Write "<postfield name=""Phone"" value="""&Phone&"""/>" &vbcrlf
            Response.Write "<postfield name=""Email"" value="""&Email&"""/>" &vbcrlf
            Response.Write "<postfield name=""Mobile"" value="""&Mobile&"""/>" &vbcrlf
            Response.Write "<postfield name=""QQ"" value="""&QQ&"""/>" &vbcrlf
            Response.Write "<postfield name=""PaymentType"" value="""&PaymentType&"""/>" &vbcrlf
            Response.Write "<postfield name=""DeliverType"" value="""&DeliverType&"""/>" &vbcrlf
            Response.Write "<postfield name=""InvoiceContent"" value="""&InvoiceContent&"""/>" &vbcrlf
            Response.Write "<postfield name=""Remark"" value="""&Remark&"""/>" &vbcrlf
            Response.Write "</go></anchor><br/>" &vbcrlf
			Response.Write "<anchor><prev/>返回修改订单</anchor><br/><br/>" &vbcrlf
			Response.Write "<a href="""&KS.GetGoBackIndex&""">返回首页</a><br/>" &vbcrlf
        End Sub
End Class
%>
