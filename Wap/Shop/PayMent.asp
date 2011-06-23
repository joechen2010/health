<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
'********************************
'* 程序功能：我的购物车
'* 演示地址: http://wap.kesion.com/
'********************************
%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<% Response.ContentType="text/vnd.wap.wml" %><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="收银台">
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
		    Products = Split(Replace(KS.S("ID")," ",""), ",")
			For I=0 To UBound(Products)
				PutToShopBag Products(I), ProductList,I
			Next
			ProductList=KS.FilterIDs(ProductList)
			Session("ProductList")=ProductList
			If Cint(KS.Setting(63))=0 And Cbool(KSUser.UserLoginChecked)=false Then
			   Response.Write "本商城设置注册用户才可购买，请先<a href=""../User/Login/?../shop/PayMent.asp"">登录</a>!<br/>"
			   Exit Sub
		    End If
			If Cbool(KSUser.UserLoginChecked)=False Then
			   Response.Write "温馨提示：您还没有注册或登录。享受更多会员优惠，请先<a href=""../User/Login/?../shop/PayMent.asp"">登录</a>或<a href=""../User/Reg/?../shop/PayMent.asp"">注册</a>成为商城会员！<br/>" &vbcrlf
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
			   Response.Write "您的购物车中没有商品!<br/><br/>" &vbcrlf
			   Response.Write "<anchor><prev/>还回上级</anchor><br/>" &vbcrlf
			   Response.Write "<a href="""&KS.GetGoBackIndex&""">返回首页</a><br/>" &vbcrlf
			   Exit Sub
			End If
			RS.Open strsql,Conn,1,1
			Dim TotalPrice,RealPrice,Price_Original,Discount,Amount
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
					    RealPrice=RSp(0)
					 End If
					 RSP.Close:Set RSP=Nothing
				  End If
			   Else
			   RealPrice=RS("Price")
			   End If
			   TotalPrice=TotalPrice+Round(RealPrice*Amount,2)
			   Response.Write "编 号:"&RS("ID")&" 商品名称:"&RS("Title")&"<br/>" &vbcrlf
			   Response.Write "数 量:"&Amount&" 原 价:￥"&RS("Price_Original")&" 折 扣:"&RS("Discount")&" 实 价:￥"&RealPrice&" 总 计:￥"&Round(RealPrice*Amount,2)&"<br/>" &vbcrlf
			   Response.Write "<br/>" &vbcrlf
			   RS.MoveNext
		    Loop
		    RS.close:set RS=nothing
			Response.Write "以上是您购物车中的商品信息，请核对正确无误后下单!点此<a href=""ShoppingCart.asp?"&KS.WapValue&""">修改订单</a><br/>" &vbcrlf
			Response.Write "合计：￥" & Round(TotalPrice,2) & "元<br/>" &vbcrlf
			Response.Write "=请填写收货信息=<br/>" &vbcrlf
			Response.Write "收货人姓名:<input name=""ContactMan"&Minute(Now)&Second(Now)&""" type=""text"" emptyok=""false"" maxlength=""30"" value="""&KSUser.RealName&"""/><br/>" &vbcrlf
			Response.Write "收货人地址:<input name=""Address"&Minute(Now)&Second(Now)&""" type=""text"" emptyok=""false"" maxlength=""100"" value="""&KSUser.Address&"""/><br/>" &vbcrlf
			Response.Write "收货人邮编:<input name=""ZipCode"&Minute(Now)&Second(Now)&""" type=""text"" emptyok=""false"" maxlength=""30"" value="""&KSUser.zip&"""/><br/>" &vbcrlf
            Response.Write "收货人电话:<input name=""Phone"&Minute(Now)&Second(Now)&""" type=""text"" emptyok=""false"" maxlength=""30"" value="""&KSUser.OfficeTel&"""/><br/>" &vbcrlf
            Response.Write "收货人邮箱:<input name=""Email"&Minute(Now)&Second(Now)&""" type=""text"" maxlength=""30"" value="""&KSUser.Email&"""/><br/>" &vbcrlf
            Response.Write "收货人手机:<input name=""Mobile"&Minute(Now)&Second(Now)&""" type=""text"" emptyok=""false"" maxlength=""30"" value="""&KSUser.Mobile&"""/><br/>" &vbcrlf
            Response.Write "收货人QQ:<input name=""QQ"&Minute(Now)&Second(Now)&""" type=""text"" maxlength=""30"" value="""&KSUser.QQ&"""/><br/>" &vbcrlf
            Response.Write "付款方式:"&GetPaymentTypeStr&"<br/>" &vbcrlf
            Response.Write "送货方式:"&GetDeliveryTypeStr&"<br/>" &vbcrlf
            Response.Write "是否需要发票:<select name=""NeedInvoice""><option value=""0"">不用发票</option><option value=""1"">需要发票</option></select><br/>" &vbcrlf
            Response.Write "发票信息:<input name=""InvoiceContent"&Minute(Now)&Second(Now)&""" type=""text"" maxlength=""30"" value=""""/><br/>" &vbcrlf
            Response.Write "备注留言:<input name=""Remark"&Minute(Now)&Second(Now)&""" type=""text"" maxlength=""30"" value=""""/><br/>" &vbcrlf
			
			Response.Write "<anchor>确认订单<go href=""Preview.asp?"&KS.WapValue&""" method=""post"">" &vbcrlf
            Response.Write "<postfield name=""ContactMan"" value=""$(ContactMan"&Minute(Now)&Second(Now)&")""/>" &vbcrlf
            Response.Write "<postfield name=""Address"" value=""$(Address"&Minute(Now)&Second(Now)&")""/>" &vbcrlf
            Response.Write "<postfield name=""ZipCode"" value=""$(ZipCode"&Minute(Now)&Second(Now)&")""/>" &vbcrlf
            Response.Write "<postfield name=""Phone"" value=""$(Phone"&Minute(Now)&Second(Now)&")""/>" &vbcrlf
            Response.Write "<postfield name=""Email"" value=""$(Email"&Minute(Now)&Second(Now)&")""/>" &vbcrlf
            Response.Write "<postfield name=""Mobile"" value=""$(Mobile"&Minute(Now)&Second(Now)&")""/>" &vbcrlf
            Response.Write "<postfield name=""QQ"" value=""$(QQ"&Minute(Now)&Second(Now)&")""/>" &vbcrlf
            Response.Write "<postfield name=""PaymentType"" value=""$(PaymentType)""/>" &vbcrlf
            Response.Write "<postfield name=""DeliverType"" value=""$(DeliverType)""/>" &vbcrlf
            Response.Write "<postfield name=""InvoiceContent"" value=""$(InvoiceContent"&Minute(Now)&Second(Now)&")""/>" &vbcrlf
            Response.Write "<postfield name=""Remark"" value=""$(Remark"&Minute(Now)&Second(Now)&")""/>" &vbcrlf
            Response.Write "<postfield name=""NeedInvoice"" value=""$(NeedInvoice)""/>" &vbcrlf
            Response.Write "</go></anchor>" &vbcrlf
            Response.Write "<br/><br/>" &vbcrlf


			Response.Write "<anchor><prev/>还回上级</anchor><br/>" &vbcrlf
			Response.Write "<a href="""&KS.GetGoBackIndex&""">返回首页</a><br/>" &vbcrlf
        End Sub
		
		'付款方式
		Function GetPaymentTypeStr()
		    Dim DiscountStr,SQL,I,RS
			Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "select TypeID,TypeName,IsDefault,Discount from KS_PaymentType order by orderid",conn,1,1
			If Not RS.Eof Then
			   SQL=RS.GetRows(-1)
			End IF
			RS.Close:Set RS=Nothing
			GetPaymentTypeStr="<select name=""PaymentType"">"
			For I=0 To UBound(SQL,2)
			    If SQL(3,I)<>100 Then
				   DiscountStr="折扣率 " & SQL(3,I) & "%"
				Else
				   DiscountStr=""
				End iF
				If SQL(2,I)=1 Then
				   GetPaymentTypeStr=GetPaymentTypeStr& "<option value=""" & SQL(0,I) & """>"  &SQL(1,I)&DiscountStr & "</option>"
				Else
				   GetPaymentTypeStr=GetPaymentTypeStr& "<option value=""" & SQL(0,I) & """>"  &SQL(1,I) &DiscountStr & "</option>"
				End If
			Next
			GetPaymentTypeStr=GetPaymentTypeStr & "</select>"
		End Function
		
		'发货方式
		Function GetDeliveryTypeStr()
		    Dim DiscountStr,SQL,I,RS
			Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "select TypeID,TypeName,IsDefault,fee from KS_Delivery order by orderid",conn,1,1
			If Not RS.Eof Then
			   SQL=RS.GetRows(-1)
			End IF
			RS.Close:Set RS=Nothing
			GetDeliveryTypeStr="<select name=""DeliverType"">"
			For I=0 To UBound(SQL,2)
			    If SQL(3,I)=0 Then
				   DiscountStr="免费"
				Else
				   DiscountStr="加收 ￥" & SQL(3,I) & " 元"
				End iF
				If SQL(2,I)=1 Then
				   GetDeliveryTypeStr=GetDeliveryTypeStr& "<option value=""" & SQL(0,I) & """>"  &SQL(1,I) &DiscountStr & "</option>"
				Else
				   GetDeliveryTypeStr=GetDeliveryTypeStr& "<option value=""" & SQL(0,I) & """>"  &SQL(1,I) &DiscountStr & "</option>"
				End If
			Next
			GetDeliveryTypeStr=GetDeliveryTypeStr & "</select>"
		End Function

  		Sub PutToShopBag( Prodid, ProductList ,I)
		    If KS.S("Action")="set" Then
			   If i = 0 Then
				  ProductList =Prodid
			   ElseIf InStr( ProductList, Prodid ) <= 0 Then
				  ProductList = ProductList&", "&Prodid &""
			   End If
		   Else
			   If Len(ProductList) = 0 Then
				  ProductList =Prodid
			   ElseIf InStr( ProductList, Prodid ) <= 0 Then
				  ProductList = ProductList&", "&Prodid &""
			   End If
		  End If
      End Sub
End Class
%>