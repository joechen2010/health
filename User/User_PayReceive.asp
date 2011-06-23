<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../Plus/md5.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Response.Buffer = true 
Response.Expires = 0 
Response.CacheControl = "no-cache"

Dim KSUser:Set KSUser=New UserCls
Dim KS:Set KS=New PublicCls
Dim PaymentPlat:PaymentPlat=KS.ChkClng(KS.S("PaymentPlat"))
IF PaymentPlat=0 Then Response.Write("error!"):Response.End()

Dim RSP:Set RSP=Server.CreateObject("ADODB.RECORDSET")
RSP.Open "Select * From KS_PaymentPlat where id=" & PaymentPlat,conn,1,1
If RSP.Eof Then
		 RSP.Close:Set RSP=Nothing
		 Response.Write "Error!"
		 Response.End()
End If
Dim AccountID:AccountID=RSP("AccountID")
Dim MD5Key:MD5Key=RSP("MD5Key")
Dim PayOnlineRate:PayOnlineRate=KS.ChkClng(RSP("Rate")) 
Dim RateByUser:RateByUser=KS.ChkClng(RSP("RateByUser")) 
RSP.Close:Set RSP=Nothing

Select Case PaymentPlat
		 Case 1 '网银在线
		  Call ChinaBank()
		 Case 2 '中国在线支付网
		  Call ipayBack()
		 Case 3 '上海环迅
		  Call IpsBack()
		 Case 4 '西部支付
		  Call YeepayBack()
		 Case 5 '易付通
		  Call xpayBack()
		 Case 6 '云网支付
		  Call cncardBack() 
		 Case 7 '支付宝
		  Call alipayBack()
		 Case 8 '快钱支付
		  Call billback()
		 Case 9  '支付宝非即时到账
		  Call alipayBack9()
		 Case 10 '财付通
		  Call tenpayback()
		 Case 11 '财付通中介交易
		  Call tenpayZJ()
End Select 

'网银在线返回
Sub ChinaBank() 
 Dim v_oid,v_pmode,v_pstatus,v_pstring,v_string,v_amount,v_moneytype,remark2,v_md5str,text,md5text,zhuangtai
' 取得返回参数值
	v_oid=request("v_oid")                               ' 商户发送的v_oid定单编号
	v_pmode=request("v_pmode")                           ' 支付方式（字符串） 
	v_pstatus=request("v_pstatus")                       ' 支付状态 20（支付成功）;30（支付失败）
	v_pstring=request("v_pstring")                       ' 支付结果信息 支付完成（当v_pstatus=20时）；失败原因（当v_pstatus=30时）；
	v_amount=request("v_amount")                         ' 订单实际支付金额
	v_moneytype=request("v_moneytype")                   ' 订单实际支付币种
	remark2=request("remark2")                           ' 备注字段2
	v_md5str=request("v_md5str")                         ' 网银在线拼凑的Md5校验串
	if request("v_md5str")="" then
		response.Write("v_md5str：空值")
		response.end
	end if
	text = v_oid&v_pstatus&v_amount&v_moneytype&MD5Key 'md5校验
	md5text = Ucase(trim(md5(text,32)))    '商户拼凑的Md5校验串
	if md5text<>v_md5str then		' 网银在线拼凑的Md5校验串 与 商户拼凑的Md5校验串 进行对比
	  response.write("MD5 error")
	else
	  if v_pstatus=20 then '支付成功
		Call UpdateOrder(v_amount,remark2,v_oid,v_pmode)
	  end if
	end if
	Dim message
	message="此次交易编号： " & v_oid & "<p>在线支付结果："
	if v_pstatus=20 then
		message = message & "在线支付成功"
    elseif v_pstatus=30 then
		message = message & "在线支付失败!"
   end if
    message = message & "</p><p>您所使用的卡为：" & v_pmode & "</p><p>金额：" & v_amount & "</p><p>币种：人民币</p>"
    Call ShowResult(message)
end Sub

'中国在线支付网
Sub ipayBack()
	Dim v_mid, v_oid, v_pmode, v_pstatus, v_pstring, v_amount, v_md5, v_date, v_moneytype
	Dim md5string
	v_mid = AccountID
	v_date = Trim(Request("v_date"))      '订单日期
	v_oid = Trim(Request("v_oid"))       '支付定单号
	v_amount = Trim(Request("v_amount"))   '订单金额
	v_pstatus = Trim(Request("v_status"))   '订单状态
	v_md5 = Trim(Request("v_md5"))         'MD5签名
	md5string = MD5(v_date & v_mid & v_oid & v_amount & v_pstatus & MD5Key, 32)
	v_pmode = ""
	v_pstring = ""
	If UCase(v_md5) = UCase(md5string) And v_pstatus = "00" Then
	    Call UpdateOrder(v_amount,"在线充值，订单号为:" & v_oid,v_oid,v_pmode)
		Call ShowResult("恭喜你！在线支付成功！")
	Else
		Call ShowResult("在线支付失败！")
	End If
End Sub

'上海环迅
Sub IpsBack()
	Dim v_mid, v_oid, v_pmode, v_pstatus, v_pstring, v_amount, v_md5, v_date, v_moneytype
	Dim Billno, amount, succ, InputDate, Signature, myDate, msg, sContent, pubfilename
	Dim md5string
	v_mid = AccountID
	Billno = Trim(Request.QueryString("billno"))
	amount = Trim(Request.QueryString("amount"))
	succ = Trim(Request.QueryString("succ"))
	myDate = Trim(Request.QueryString("date"))
	InputDate = Mid(myDate, 1, 4) & "-" & Mid(myDate, 5, 2) & "-" & Mid(myDate, 7, 2)
	msg = Trim(Request.QueryString("msg"))
	Signature = Trim(Request.QueryString("signature")) '密文
	
	If succ = "Y" Then
		sContent = Billno & amount & myDate & succ    '组成明文字符进行签名
		pubfilename = "c:\secre\public.key"           'pubfilename为您保存公钥文件的全路经名
		'签名认证
		Dim secre
		Set secre = Server.CreateObject("SignandVerify.RSACom")
		If secre.VerifyMessage(pubfilename, sContent, Signature) = 0 Then
			v_oid = myDate & Right(Billno, 6)
			v_amount = amount
			v_pstring = msg
			v_pmode = ""
			Call UpdateOrder(v_amount,"在线充值，订单号为:" & v_oid,v_oid,v_pmode)
			Call ShowResult("恭喜你！在线支付成功！")
		Else
			Call ShowResult("在线支付失败！")
		End If
		Set secre = Nothing
	Else
		Call ShowResult("在线支付失败！")
	End If
End Sub

'西部支付
Sub YeepayBack()
	Dim PaySuccess:PaySuccess = False
	Dim v_mid, v_oid, v_pmode, v_pstatus, v_pstring, v_amount, v_md5, v_date, v_moneytype
	Dim md5string
	v_mid = Request("MerchantID")
	'注：商户必须判断此商户ID是不是您的商户ID
	v_oid = Request("MerchantOrderNumber") '和商户支付命令中的订单号相同
	'WestPayOrderNumber = Request("WestPayOrderNumber")
	v_amount = Request("PaidAmount") 'WestPay传回的实际支付金额。用CCUR转为数字型。
	'注：商户必须根据我们传回商户原始订单号找到原始订单，比较实付金额和原始订单金额，相同才是支付成功。
	
	Dim objHttp, str
	' 准备回传支付通知表单
	str = Request.Form & "&cmd=validate"
	Set objHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
	 
	'把WestPay传来的通知内容再传回到WestPay作验证以确保通知信息的真实性
	objHttp.Open "POST", "http://www.yeepay.com/pay/ISPN.asp", False    'ISPN: Instant Secure Payment Notification
	objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
	objHttp.Send str
	If (objHttp.Status <> 200) Then
		'HTTP 错误处理
		Response.Write ("Status=" & objHttp.Status)
	ElseIf (objHttp.ResponseText = "VERIFIED") Then
		'支付通知验证成功
		If Trim(v_mid) = Trim(AccountID) Then '判断此订单是不是该商户的订单。
			PaySuccess = True
		End If
	ElseIf (objHttp.ResponseText = "INVALID") Then
		'支付通知验证失败
		Response.Write ("Invalid")
	Else
		'支付通知验证过程中出现错误
		Response.Write ("Error")
	End If
	Set objHttp = Nothing
	
	If PaySuccess = True Then
		Call UpdateOrder(v_amount,"在线充值，订单号为:" & v_oid,v_oid,"")
		Call ShowResult("恭喜你！在线支付成功！")
	Else
		Call ShowResult("在线支付失败！")
	End If
End Sub

'易付通
Sub xpayBack()
	Dim v_mid, v_oid, v_pmode, v_pstatus, v_pstring, v_amount, v_md5, v_date, v_moneytype
	Dim md5string,v_sid
	v_mid = AccountID
	v_oid = Trim(Request("bid"))       '支付定单号
	v_sid = Trim(Request("sid"))         '易付通交易成功 流水号
	v_md5 = Trim(Request("md"))       '数字签名
	v_amount = Trim(Request("prc"))       '支付金额
	v_pstatus = Trim(Request("success"))       '支付状态
	v_pmode = Trim(Request("bankcode"))       '支付银行
	v_pstring = Trim(Request("v_pstring"))       '支付结果说明
	
	md5string = MD5(MD5Key & ":" & v_oid & "," & v_sid & "," & v_amount & ",sell,," & v_mid & ",bank," & v_pstatus, 32)
	
	If UCase(v_md5) = UCase(md5string) And LCase(v_pstatus) = "true" Then
		Call UpdateOrder(v_amount,"在线充值，订单号为:" & v_oid,v_oid,v_pmode)
		Call ShowResult("恭喜你！在线支付成功！")
	Else
		Call ShowResult("在线支付失败！")
	End If
End Sub

'云网支付
Sub cncardBack
	Dim PaySuccess:PaySuccess = False
	Dim v_mid, v_oid, v_pmode, v_pstatus, v_pstring, v_amount, v_md5, v_date, v_moneytype
	Dim md5string
	Dim c_mid, c_order, c_orderamount, c_ymd, c_transnum, c_succmark, c_moneytype, c_cause, c_memo1, c_memo2, c_signstr
	c_mid = Request("c_mid")                    '商户编号，在申请商户成功后即可获得，可以在申请商户成功的邮件中获取该编号
	c_order = Request("c_order")                '商户提供的订单号
	c_orderamount = Request("c_orderamount")    '商户提供的订单总金额，以元为单位，小数点后保留两位，如：13.05
	c_ymd = Request("c_ymd")                    '商户传输过来的订单产生日期，格式为"yyyymmdd"，如20050102
	c_transnum = Request("c_transnum")          '云网支付网关提供的该笔订单的交易流水号，供日后查询、核对使用；
	c_succmark = Request("c_succmark")          '交易成功标志，Y-成功 N-失败
	c_moneytype = Request("c_moneytype")        '支付币种，0为人民币
	c_cause = Request("c_cause")                '如果订单支付失败，则该值代表失败原因
	c_memo1 = Request("c_memo1")                '商户提供的需要在支付结果通知中转发的商户参数一
	c_memo2 = Request("c_memo2")                '商户提供的需要在支付结果通知中转发的商户参数二
	c_signstr = Request("c_signstr")            '云网支付网关对已上信息进行MD5加密后的字符串
	
	md5string = MD5(c_mid & c_order & c_orderamount & c_ymd & c_transnum & c_succmark & c_moneytype & c_memo1 & c_memo2 & MD5Key, 32)
	
	If UCase(md5string) <> UCase(c_signstr) Then
		Response.Write "签名验证失败"
		Response.End
	End If
	
	If Trim(AccountID) <> c_mid Then
		Response.Write "提交的商户编号有误"
		Response.End
	End If
	
	If c_succmark <> "Y" And c_succmark <> "N" Then
		Response.Write "参数提交有误"
		Response.End
	End If
	
	PaySuccess = True
	v_oid = c_order
	v_amount = c_orderamount
	v_pstring = ""
	v_pmode = ""
	If PaySuccess = True Then
		Call UpdateOrder(v_amount,"在线充值，订单号为:" & v_oid,v_oid,v_pmode)
		Call ShowResult("恭喜你！在线支付成功！")
	Else
		Call ShowResult("在线支付失败！")
	End If
End Sub

'快钱支付
Sub billback()
Dim md5string
Dim merchantAcctId, key, version, language, signType, payType, bankId, orderId, orderTime, orderAmount, dealId, bankDealId, dealTime, payAmount
Dim fee, ext1, ext2, payResult, errCode, signMsg, merchantSignMsgVal

merchantAcctId = Trim(request("merchantAcctId")) '获取人民币网关账户号
key = MD5Key '设置人民币网关密钥
version = Trim(request("version")) '获取网关版本
language = Trim(request("language")) '获取语言种类,1代表中文；2代表英文
signType = Trim(request("signType")) '签名类型,1代表MD5签名
payType = Trim(request("payType")) '获取支付方式,00：组合支付,10：银行卡支付,11：电话银行支付,12：快钱账户支付,13：线下支付,14：B2B支付
bankId = Trim(request("bankId")) '获取银行代码
orderId = Trim(request("orderId")) '获取商户订单号
orderTime = Trim(request("orderTime")) '获取订单提交时间
orderAmount = Trim(request("orderAmount")) '获取原始订单金额
dealId = Trim(request("dealId")) '获取快钱交易号
bankDealId = Trim(request("bankDealId")) '获取银行交易号
dealTime = Trim(request("dealTime")) '获取在快钱交易时间
payAmount = Trim(request("payAmount")) '获取实际支付金额,单位为分
fee = Trim(request("fee")) '获取交易手续费
ext1 = Trim(request("ext1")) '获取扩展字段1
ext2 = Trim(request("ext2")) '获取扩展字段2

'获取处理结果
''10代表 成功11代表 失败
''00代表 下订单成功（仅对电话银行支付订单返回）;01代表 下订单失败（仅对电话银行支付订单返回）
payResult = Trim(request("payResult"))
errCode = Trim(request("errCode")) '获取错误代码,详细见文档错误代码列表
signMsg = Trim(request("signMsg")) '获取加密签名串

'生成加密串。必须保持如下顺序。
merchantSignMsgVal = appendParam(merchantSignMsgVal, "merchantAcctId", merchantAcctId)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "version", version)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "language", language)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "signType", signType)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "payType", payType)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "bankId", bankId)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "orderId", orderId)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "orderTime", orderTime)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "orderAmount", orderAmount)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "dealId", dealId)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "bankDealId", bankDealId)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "dealTime", dealTime)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "payAmount", payAmount)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "fee", fee)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "ext1", ext1)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "ext2", ext2)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "payResult", payResult)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "errCode", errCode)
merchantSignMsgVal = appendParam(merchantSignMsgVal, "key", key)

md5string = MD5(merchantSignMsgVal, 32)

Dim rtnOk, rtnUrl
rtnOk = 0
rtnUrl = ""

''首先进行签名字符串验证
If UCase(signMsg) = UCase(md5string) Then
    ''接着进行支付结果判断
    Select Case payResult
          Case "10"   '支付成功，更新订单
            rtnOk = 1
			Call UpdateOrder(orderAmount / 100,"在线充值，订单号为:" & orderId,orderId,"")
			Call ShowResult("恭喜你！在线支付成功！")
         Case Else
            rtnOk = 1
    End Select
Else
    rtnOk = 1
End If
%>
<result><%=rtnOk %></result><redirecturl><%=rtnUrl %></redirecturl>
<%
End Sub

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
'支付宝
Sub alipayBack()
	Dim v_mid, v_oid, v_pmode, v_pstatus, v_pstring, v_amount, v_md5, v_date, v_moneytype
	Dim md5string,alipayNotifyURL
    v_mid = AccountID
	Dim Partner
	Dim ArrMD5Key
	If InStr(MD5Key, "|") > 0 Then
		ArrMD5Key = Split(MD5Key, "|")
		If UBound(ArrMD5Key) = 1 Then
			MD5Key = ArrMD5Key(0)
			Partner = ArrMD5Key(1)
		End If
	End If


	Dim trade_status, sign, MySign, Retrieval,ResponseTxt
	Dim mystr, Count, i, minmax, minmaxSlot, j, mark, temp, value, md5str, notify_id
	
	v_oid = DelStr(Request("out_trade_no"))            '商户定单号
	trade_status = DelStr(Request("trade_status"))
	sign = DelStr(Request("sign"))
	v_amount = DelStr(Request("total_fee"))
	notify_id = Request("notify_id")
	

	alipayNotifyURL = "http://notify.alipay.com/trade/notify_query.do?"
	alipayNotifyURL = alipayNotifyURL & "partner=" & Partner & "&notify_id=" & notify_id
	Set Retrieval = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
    Retrieval.setOption 2, 13056 
    Retrieval.open "GET", alipayNotifyURL, False, "", "" 
    Retrieval.send()
    ResponseTxt = Retrieval.ResponseText
	Set Retrieval = Nothing
	'*****************************************
	'获取支付宝GET过来通知消息,判断消息是不是被修改过
	Dim varItem
	For Each varItem in Request.QueryString
	mystr=varItem&"="&Request(varItem)&"^"&mystr
	Next 
	If mystr<>"" Then 
	mystr=Left(mystr,Len(mystr)-1)
	End If 
	mystr = SPLIT(mystr, "^")

	Count=ubound(mystr)
	'对参数排序
	For i = Count TO 0 Step -1
	minmax = mystr( 0 )
	minmaxSlot = 0
	For j = 1 To i
	mark = (mystr( j ) > minmax)
	If mark Then 
	minmax = mystr( j )
	minmaxSlot = j
	End If 
	Next
		
	If minmaxSlot <> i Then 
	temp = mystr( minmaxSlot )
	mystr( minmaxSlot ) = mystr( i )
	mystr( i ) = temp
	End If
	Next
	'构造md5摘要字符串
	 For j = 0 To Count Step 1
	 value = SPLIT(mystr( j ), "=")
	 If  value(1)<>"" And value(0)<>"sign" And value(0)<>"sign_type"  Then
	 If j=Count Then
	 md5str= md5str&mystr( j )
	 Else 
	 md5str= md5str&mystr( j )&"&"
	 End If 
	 End If 
	 Next
	 md5str=md5str&MD5Key
	 mysign=md5(md5str,32)

    ' response.write mysign & "==" & request("sign")
	'********************************************************
	
	'If mysign=Request("sign") and ResponseTxt="true"   Then 	
	'If ResponseTxt="true" and Session("PayType")="ALIPAY" Then 	
	If ResponseTxt="true" Then 	
		Call UpdateOrder(v_amount,"在线充值，订单号为:" & v_oid,v_oid,v_pmode)
		Call ShowResult("恭喜你！在线支付成功！")
	Else
	Call ShowResult("在线支付失败！")          '这里可以指定你需要显示的内容
	End If 
	
End Sub

'支付宝非即时到账
Sub alipayBack9()
    Dim PaySuccess,ResponseTxt,returnTxt
      PaySuccess = False
	Dim v_mid, v_oid, v_pmode, v_pstatus, v_pstring, v_amount, v_md5, v_date, v_moneytype
	Dim md5string,alipayNotifyURL
    v_mid = AccountID
	Dim Partner
	Dim ArrMD5Key
	If InStr(MD5Key, "|") > 0 Then
		ArrMD5Key = Split(MD5Key, "|")
		If UBound(ArrMD5Key) = 1 Then
			MD5Key = ArrMD5Key(0)
			Partner = ArrMD5Key(1)
		End If
	End If
    Dim trade_status, sign, MySign, Retrieval
    Dim mystr, Count, i, minmax, minmaxSlot, j, mark, temp, value, md5str, notify_id
    
    v_oid = DelStr(Request("out_trade_no"))            '商户定单号
    trade_status = DelStr(Request("trade_status"))
    sign = DelStr(Request("sign"))
    v_amount = DelStr(Request("price"))
    notify_id = Request.Form("notify_id")


    alipayNotifyURL = "https://www.alipay.com/cooperate/gateway.do?"

    alipayNotifyURL = alipayNotifyURL & "service=notify_verify&partner=" & Partner & "&notify_id=" & notify_id
	Set Retrieval = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
    Retrieval.Open "GET", alipayNotifyURL, False, "", ""
    Retrieval.Send
    ResponseTxt = Retrieval.ResponseText
    Set Retrieval = Nothing

                
    '获取POST过来的参数
    mystr = Split(URLDecode(Request.Form), "&")
    Count = UBound(mystr)

    '对参数排序
    For i = Count To 0 Step -1
        minmax = mystr(0)
        minmaxSlot = 0
        For j = 1 To i
            mark = (mystr(j) > minmax)
            If mark Then
                minmax = mystr(j)
                minmaxSlot = j
            End If
        Next

        If minmaxSlot <> i Then
            temp = mystr(minmaxSlot)
            mystr(minmaxSlot) = mystr(i)
            mystr(i) = temp
        End If
    Next

    '构造md5摘要字符串
    For j = 0 To Count Step 1
        value = Split(mystr(j), "=")
        If value(1) <> "" And value(0) <> "sign" And value(0) <> "sign_type" Then
            If j = Count Then
                md5str = md5str & mystr(j)
            Else
                md5str = md5str & mystr(j) & "&"
            End If
        End If
    Next

    md5str = md5str & MD5Key
    '生成md5摘要
    MySign = MD5(md5str,32)


    '等待买家付款
    Select Case trade_status
    Case "WAIT_BUYER_PAY"
        If ResponseTxt = "true" And sign = MySign Then
            returnTxt = "success"
        Else
            returnTxt = "fail"
        End If

    '买家付款成功,等待卖家发货
    Case "WAIT_SELLER_SEND_GOODS"
        If ResponseTxt = "true" And sign = MySign Then
            returnTxt = "success"
			Conn.Execute("Update KS_Order Set Status=1 Where OrderID='" & KS.R(v_oid) & "'") '只更新订单状态，不更新发货状态和订单状态
        Else
            returnTxt = "fail"
        End If

    '等待买家确认收货
    Case "WAIT_BUYER_CONFIRM_GOODS"
        If ResponseTxt = "true" And sign = MySign Then
            returnTxt = "success"
            			Conn.Execute("Update KS_Order Set Status=1,DeliverStatus=1 Where OrderID='" & v_oid & "'") '更新支付记录状态和发货状态，不更新订单状态
        Else
            returnTxt = "fail"
        End If

    '交易成功结束
    Case "TRADE_FINISHED"
        If ResponseTxt = "true" And sign = MySign Then
            returnTxt = "success"
            PaySuccess = True                '交易成功，更新订单
        Else
            returnTxt = "fail"
        End If

    '其他交易状态通知情况
    Case Else
        If ResponseTxt = "true" And sign = MySign Then
            returnTxt = "success"
        Else
            returnTxt = "fail"
        End If
    End Select
    Response.Write returnTxt
	If PaySuccess = True Then
	 Call UpdateOrder(v_amount,"在线充值，订单号为:" & v_oid,v_oid,v_pmode)
	Else
	 '	Call ShowResult("在线支付失败！")          '这里可以指定你需要显示的内容
    End If

End Sub

Function DelStr(Str)
		If IsNull(Str) Or IsEmpty(Str) Then
			Str	= ""
		End If
		DelStr	= Replace(Str,";","")
		DelStr	= Replace(DelStr,"'","")
		DelStr	= Replace(DelStr,"&","")
		DelStr	= Replace(DelStr," ","")
		DelStr	= Replace(DelStr,"　","")
		DelStr	= Replace(DelStr,"%20","")
		DelStr	= Replace(DelStr,"--","")
		DelStr	= Replace(DelStr,"==","")
		DelStr	= Replace(DelStr,"<","")
		DelStr	= Replace(DelStr,">","")
		DelStr	= Replace(DelStr,"%","")
End Function

'对post传递过来的参数作urldecode编码处理(支付宝，新接口)
Function URLDecode(enStr)
    Dim deStr
    Dim c, i, v
    deStr = ""
    For i = 1 To Len(enStr)
        c = Mid(enStr, i, 1)
        If c = "%" Then
            v = eval("&h" + Mid(enStr, i + 1, 2))
            If v < 128 Then
                deStr = deStr & Chr(v)
                i = i + 2
            Else
                If isvalidhex(Mid(enStr, i, 3)) Then
                    If isvalidhex(Mid(enStr, i + 3, 3)) Then
                        v = eval("&h" + Mid(enStr, i + 1, 2) + Mid(enStr, i + 4, 2))
                        deStr = deStr & Chr(v)
                        i = i + 5
                    Else
                        v = eval("&h" + Mid(enStr, i + 1, 2) + CStr(Hex(Asc(Mid(enStr, i + 3, 1)))))
                        deStr = deStr & Chr(v)
                        i = i + 3
                    End If
                Else
                    deStr = deStr & c
                End If
            End If
        Else
            If c = "+" Then
                deStr = deStr & " "
            Else
                deStr = deStr & c
            End If
        End If
    Next
    URLDecode = deStr
End Function '处理完毕

'财付通
Sub tenpayback()
	Dim v_mid, v_oid, v_pmode, v_pstatus, v_pstring, v_amount, v_md5, v_date, v_moneytype
	Dim md5string
	v_mid = AccountID
	
	Dim cmdno, pay_result, pay_info, bill_date, bargainor_id, transaction_id, sp_billno, total_fee, fee_type, md5_sign, attach
	cmdno = Request("cmdno")
	pay_result = Request("pay_result")
	pay_info = Request("pay_info")
	bill_date = Request("date")
	bargainor_id = Request("bargainor_id")
	transaction_id = Request("transaction_id")
	sp_billno = Request("sp_billno")
	total_fee = Request("total_fee")
	fee_type = Request("fee_type")
	attach = Request("attach")
	md5_sign = Request("sign")
	
	md5string = MD5("cmdno=" & cmdno & "&pay_result=" & pay_result & "&date=" & bill_date & "&transaction_id=" & transaction_id & "&sp_billno=" & sp_billno & "&total_fee=" & total_fee & "&fee_type=" & fee_type & "&attach=" & attach & "&key=" & MD5Key, 32)
	
	If bargainor_id = v_mid And UCase(md5string) = md5_sign And pay_result = 0 Then
		v_oid = sp_billno
		v_amount = total_fee / 100
		v_pstring = ""
		v_pmode = ""
		Call UpdateOrder(v_amount,"在线充值，订单号为:" & v_oid,v_oid,v_pmode)
		Call ShowResult("恭喜你！在线支付成功！")
	Else
		Call ShowResult("在线支付失败！")
	End If
End Sub

'财付通中介
Sub tenpayZJ()
%>
<html>
<head>
	<meta name="TENCENT_ONLINE_PAYMENT" content="China TENCENT">
</head>
<%
'获取参数
Dim attach,buyer_id,cft_tid,chnid,mch_vno,cmdno,retcode,seller,status,total_fee,trade_price,transport_fee,version,sign,localSignText	
attach					= Request("attach")
buyer_id				= Request("buyer_id")
cft_tid					= Request("cft_tid")
chnid					= Request("chnid")
cmdno					= Request("cmdno")
mch_vno					= Request("mch_vno")

retcode					= Request("retcode")
seller					= Request("seller")
status					= Request("status")

total_fee				= Request("total_fee")
trade_price				= Request("trade_price")
transport_fee			= Request("transport_fee")
version                 =request("version")

sign					= Request("sign")

dim buffer
buffer = appendParam(buffer, "attach", 		attach)
buffer = appendParam(buffer, "buyer_id", 		buyer_id)
buffer = appendParam(buffer, "cft_tid", 		cft_tid)
buffer = appendParam(buffer, "chnid", 			chnid)
buffer = appendParam(buffer, "cmdno", 			cmdno)
buffer = appendParam(buffer, "mch_vno", 		mch_vno)
buffer = appendParam(buffer, "retcode", 		retcode)
buffer = appendParam(buffer, "seller", 		seller)
buffer = appendParam(buffer, "status", 		status)
buffer = appendParam(buffer, "total_fee", 		total_fee)
buffer = appendParam(buffer, "trade_price", 	trade_price)
buffer = appendParam(buffer, "transport_fee", 	transport_fee)
buffer = appendParam(buffer, "version", 	version)

buffer = appendParam(buffer, "key",			MD5Key)

'生成签名
localSignText = UCase(md5(buffer,32) )

dim msg
'签名判断
if localSignText = sign then
	'认证签名成功
	
	if retcode = "0" then
		msg = "OK"
		msg = msg & status 
		Select case status
			case "1":
				'交易创建
			case "2":
				'收获地址填写完毕
			case "3":
				Conn.Execute("Update KS_Order Set MoneyReceipt=MoneyTotal,Status=1 Where OrderID='" & KS.R(mch_vno) & "'") '更新订单状态及已付款
                response.write "<script>alert('恭喜，支付成功！请等待商家发货！');location.href='../user/index.asp?user_order.asp';</script>"
			case "4":
				Conn.Execute("Update KS_Order Set DeliverStatus=1 Where OrderID='" & KS.R(mch_vno) & "'")
				'卖家发货成功
			case "5":
				Conn.Execute("Update KS_Order Set DeliverStatus=2 Where OrderID='" & KS.R(mch_vno) & "'")
				'买家收货确认，交易成功
			case "6":
				'交易关闭，未完成超时关闭
			case "7":
				'修改交易价格成功
			case "8":
				'买家发起退款
			case "9":
				'退款成功
			case "10":
				'退款关闭
			case else
				'error
		end Select

	else
		'支付失败，请不要按成功处理
		msg = "支付失败"
	end if

else
	'认证签名失败
	msg = "认证签名失败"
end if

%>
<body>
	<div align="center"><%=msg%></div>
</body>
</html>
<%
End Sub



Sub ShowResult(byval message)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>用户管理中心</title>
<link href="images/css.css" type="text/css" rel="stylesheet" />
</head>
<body leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0"><br><br><br>
	<table class=border cellSpacing=1 cellPadding=2 width="60%" align=center border=0>
  <tr class="title"> 
    <td height=22 align=center><b><font color="#FF0000">提示：</font> 您网上在线支付情况反馈如下：</b></td>
 </tr>
 <tr class="tdbg"><td>
      <p>
        <%=message%>
	  </p>
     </td>
  </tr>
  <tr class="title">
   <td  height="22" align="center"><a href="<%=KS.getdomain%>user/index.asp">进入会员中心</a> | <a href="<%=KS.getdomain%>">返回首页</a>
   </td>
  </tr>
</table>
<%
End Sub

'更新订单记录
Sub UpdateOrder(v_amount,remark2,v_oid,v_pmode)
 Dim UserName,MoneyType,Money,Remark,sqlUser,rsUser,orderid,mobile
 orderid=v_oid
 IF Cbool(KSUser.UserLoginChecked) Then UserName=KSUser.UserName Else UserName=KS.S("UserName")
 Mobile=KSUser.Mobile
		 Money=v_amount
		 Remark=remark2
		 Dim RSLog,RS
		Set RSLog=Server.CreateObject("ADODB.RECORDSET")
		RSLog.Open "Select top 1 * From KS_LogMoney where orderid='" & v_oid & "'",Conn,1,1
		if RSLog.Eof And RSLog.BoF Then
			 Select Case request("action")
			 case "shop"   '商城中心购物
				 Set RS=Server.CreateObject("ADODB.RECORDSET")
				 RS.Open "Select top 1 * From KS_Order Where OrderID='" & v_oid & "'",Conn,1,3
				 If RS.Eof Then
				   RS.Close:Set RS=Nothing
				   Response.Write "<br><li>支付过程中遇到问题，请联系网站管理员！"
				 End If
				  If Mobile="" Then
				  Mobile=RS("Mobile")
				  End If
				  RS("MoneyReceipt")=Money
				  Dim OrderStatus:OrderStatus=rs("status")
				  RS("Status")=1
				  RS("PayTime")=now   '记录付款时间
				  RS.Update
                  orderid=RS("OrderID")
				  Call KS.MoneyInOrOut(rs("UserName"),RS("Contactman"),Money,2,1,now,rs("orderid"),"System","为购买订单：" &v_oid & "使用" & v_pmode & "在线充值",0,0)
		          Call KS.MoneyInOrOut(rs("UserName"),RS("Contactman"),Money,4,2,now,rs("orderid"),"System",Remark,0,0)
				  
					
					'====================为用户增加购物应得积分========================
					Dim rsp:set rsp=conn.execute("select point,id,title from ks_product where id in(select proid from KS_OrderItem where orderid='" & rs("orderid") & "')")
					do while not rsp.eof
					  dim amount:amount=conn.execute("select top 1 amount from ks_orderitem where orderid='" & rs("orderid") & "' and proid=" & rsp(1))(0)
					  If OrderStatus<>1 Then
					  conn.execute("update ks_product set totalnum=totalnum-" & amount &" where totalnum>=" & amount &" and id=" & rsp(1))
					 ' response.write rs("orderid") & "=55<br>"
					 ' response.write amount & "<br>"
					 ' response.write username & "<br>"
					  
					  Call KS.ScoreInOrOut(UserName,1,KS.ChkClng(rsp(0))*amount,"系统","购买商品<font color=red>" & rsp("title") & "</font>赠送!",0,0)
					  End if
					  
					rsp.movenext
					loop
					rsp.close
					set rsp=nothing
					'================================================================
					
					RS.Close:Set RS=Nothing
			 Case else   '会员中心充值
					Set rsUser=Server.CreateObject("Adodb.RecordSet")
					sqlUser="select top 1 * from KS_User where UserName='" & UserName & "'"
					rsUser.Open sqlUser,Conn,1,1
					if rsUser.bof and rsUser.eof then
								Response.Write "<br><li>充值过程中遇到问题，请联系网站管理员！"
								rsUser.close:set rsUser=Nothing
								exit sub
					end if
					Dim RealName:RealName=rsUser("RealName")
					Dim Edays:Edays=rsUser("Edays")
					Dim BeginDate:BeginDate=rsUser("BeginDate")
					rsUser.Close : Set rsUser=Nothing

					If KS.ChkClng(KS.S("UserCardID"))<>0 Then   '充值卡
					         Conn.Execute("Update KS_User Set UserCardID=" & KS.ChkClng(KS.S("UserCardID")) & " where username='" & userName & "'")
							 Dim RSCard:Set RSCard=conn.execute("select top 1 * From KS_UserCard Where ID="&KS.ChkClng(KS.S("UserCardID")))
							 If Not RSCard.Eof Then
							   Dim ValidNum:ValidNum=RSCard("ValidNum")
							   Dim CardTitle:CardTitle=RSCard("GroupName")
							   If RSCard("groupid")<>0 Then
							     Conn.Execute("Update KS_User Set GroupID=" & RSCard("GroupID") & ",ChargeType=" & KS.U_G(RSCard("groupid"),"chargetype") &" where username='" & userName & "'") 
							   End If
							    
							   Select Case RSCard("ValidUnit")
							      case 1
								   Call KS.PointInOrOut(0,0,UserName,1,ValidNum,"System","在线购买充值卡[" & CardTitle &"]获得的点数",0)
								  case 2
									Dim tmpDays:tmpDays=Edays-DateDiff("D",BeginDate,now())
									if tmpDays>0 then
									    Conn.Execute("Update KS_User Set Edays=Edays+" & ValidNum & " where username='" & userName & "'") 
									else
									    Conn.Execute("Update KS_User Set Edays=" & ValidNum & ",BeginDate=" & SQLNowString& " where username='" & userName & "'") 
									end if
									Call KS.EdaysInOrOut(UserName,1,ValidNum,"System","在线购买充值卡[" & CardTitle &"]获得的有效天数")
                                       
								  case 3
								   	Call KS.MoneyInOrOut(UserName,RealName,ValidNum,3,1,now,v_oid,"System",v_pmode & "在线充值,在线购买充值卡[" & CardTitle &"]获得的资金",0,0)
								  case 4
								     
			                        Call KS.ScoreInOrOut(UserName,1,ValidNum,"System","通过充值卡[" & CardTitle & "]获得的积分!",0,0)
							   End Select
							   If RSCard("ValidUnit")<>3 Then
								   	Call KS.MoneyInOrOut(UserName,RealName,RSCard("Money"),3,1,now,v_oid,"System",v_pmode & "在线充值!",0,0)
								   	Call KS.MoneyInOrOut(UserName,RealName,RSCard("Money"),3,2,now,v_oid,"System", "为购买充值卡[" & CardTitle &"]而支出!",0,0)
							   End If
							 End If 
							 RSCard.Close:Set RSCard=Nothing
							 

					Else
				  	 Call KS.MoneyInOrOut(UserName,RealName,Money,3,1,now,v_oid,"System",v_pmode & "在线充值,订单号为:" & v_oid,0,0)
					End If

					
			 End Select
			 
		End If

		RSLog.Close:Set RSLog=Nothing
End Sub
%>