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
		 Case 1 '��������
		  Call ChinaBank()
		 Case 2 '�й�����֧����
		  Call ipayBack()
		 Case 3 '�Ϻ���Ѹ
		  Call IpsBack()
		 Case 4 '����֧��
		  Call YeepayBack()
		 Case 5 '�׸�ͨ
		  Call xpayBack()
		 Case 6 '����֧��
		  Call cncardBack() 
		 Case 7 '֧����
		  Call alipayBack()
		 Case 8 '��Ǯ֧��
		  Call billback()
		 Case 9  '֧�����Ǽ�ʱ����
		  Call alipayBack9()
		 Case 10 '�Ƹ�ͨ
		  Call tenpayback()
		 Case 11 '�Ƹ�ͨ�н齻��
		  Call tenpayZJ()
End Select 

'�������߷���
Sub ChinaBank() 
 Dim v_oid,v_pmode,v_pstatus,v_pstring,v_string,v_amount,v_moneytype,remark2,v_md5str,text,md5text,zhuangtai
' ȡ�÷��ز���ֵ
	v_oid=request("v_oid")                               ' �̻����͵�v_oid�������
	v_pmode=request("v_pmode")                           ' ֧����ʽ���ַ����� 
	v_pstatus=request("v_pstatus")                       ' ֧��״̬ 20��֧���ɹ���;30��֧��ʧ�ܣ�
	v_pstring=request("v_pstring")                       ' ֧�������Ϣ ֧����ɣ���v_pstatus=20ʱ����ʧ��ԭ�򣨵�v_pstatus=30ʱ����
	v_amount=request("v_amount")                         ' ����ʵ��֧�����
	v_moneytype=request("v_moneytype")                   ' ����ʵ��֧������
	remark2=request("remark2")                           ' ��ע�ֶ�2
	v_md5str=request("v_md5str")                         ' ��������ƴ�յ�Md5У�鴮
	if request("v_md5str")="" then
		response.Write("v_md5str����ֵ")
		response.end
	end if
	text = v_oid&v_pstatus&v_amount&v_moneytype&MD5Key 'md5У��
	md5text = Ucase(trim(md5(text,32)))    '�̻�ƴ�յ�Md5У�鴮
	if md5text<>v_md5str then		' ��������ƴ�յ�Md5У�鴮 �� �̻�ƴ�յ�Md5У�鴮 ���жԱ�
	  response.write("MD5 error")
	else
	  if v_pstatus=20 then '֧���ɹ�
		Call UpdateOrder(v_amount,remark2,v_oid,v_pmode)
	  end if
	end if
	Dim message
	message="�˴ν��ױ�ţ� " & v_oid & "<p>����֧�������"
	if v_pstatus=20 then
		message = message & "����֧���ɹ�"
    elseif v_pstatus=30 then
		message = message & "����֧��ʧ��!"
   end if
    message = message & "</p><p>����ʹ�õĿ�Ϊ��" & v_pmode & "</p><p>��" & v_amount & "</p><p>���֣������</p>"
    Call ShowResult(message)
end Sub

'�й�����֧����
Sub ipayBack()
	Dim v_mid, v_oid, v_pmode, v_pstatus, v_pstring, v_amount, v_md5, v_date, v_moneytype
	Dim md5string
	v_mid = AccountID
	v_date = Trim(Request("v_date"))      '��������
	v_oid = Trim(Request("v_oid"))       '֧��������
	v_amount = Trim(Request("v_amount"))   '�������
	v_pstatus = Trim(Request("v_status"))   '����״̬
	v_md5 = Trim(Request("v_md5"))         'MD5ǩ��
	md5string = MD5(v_date & v_mid & v_oid & v_amount & v_pstatus & MD5Key, 32)
	v_pmode = ""
	v_pstring = ""
	If UCase(v_md5) = UCase(md5string) And v_pstatus = "00" Then
	    Call UpdateOrder(v_amount,"���߳�ֵ��������Ϊ:" & v_oid,v_oid,v_pmode)
		Call ShowResult("��ϲ�㣡����֧���ɹ���")
	Else
		Call ShowResult("����֧��ʧ�ܣ�")
	End If
End Sub

'�Ϻ���Ѹ
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
	Signature = Trim(Request.QueryString("signature")) '����
	
	If succ = "Y" Then
		sContent = Billno & amount & myDate & succ    '��������ַ�����ǩ��
		pubfilename = "c:\secre\public.key"           'pubfilenameΪ�����湫Կ�ļ���ȫ·����
		'ǩ����֤
		Dim secre
		Set secre = Server.CreateObject("SignandVerify.RSACom")
		If secre.VerifyMessage(pubfilename, sContent, Signature) = 0 Then
			v_oid = myDate & Right(Billno, 6)
			v_amount = amount
			v_pstring = msg
			v_pmode = ""
			Call UpdateOrder(v_amount,"���߳�ֵ��������Ϊ:" & v_oid,v_oid,v_pmode)
			Call ShowResult("��ϲ�㣡����֧���ɹ���")
		Else
			Call ShowResult("����֧��ʧ�ܣ�")
		End If
		Set secre = Nothing
	Else
		Call ShowResult("����֧��ʧ�ܣ�")
	End If
End Sub

'����֧��
Sub YeepayBack()
	Dim PaySuccess:PaySuccess = False
	Dim v_mid, v_oid, v_pmode, v_pstatus, v_pstring, v_amount, v_md5, v_date, v_moneytype
	Dim md5string
	v_mid = Request("MerchantID")
	'ע���̻������жϴ��̻�ID�ǲ��������̻�ID
	v_oid = Request("MerchantOrderNumber") '���̻�֧�������еĶ�������ͬ
	'WestPayOrderNumber = Request("WestPayOrderNumber")
	v_amount = Request("PaidAmount") 'WestPay���ص�ʵ��֧������CCURתΪ�����͡�
	'ע���̻�����������Ǵ����̻�ԭʼ�������ҵ�ԭʼ�������Ƚ�ʵ������ԭʼ��������ͬ����֧���ɹ���
	
	Dim objHttp, str
	' ׼���ش�֧��֪ͨ��
	str = Request.Form & "&cmd=validate"
	Set objHttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
	 
	'��WestPay������֪ͨ�����ٴ��ص�WestPay����֤��ȷ��֪ͨ��Ϣ����ʵ��
	objHttp.Open "POST", "http://www.yeepay.com/pay/ISPN.asp", False    'ISPN: Instant Secure Payment Notification
	objHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
	objHttp.Send str
	If (objHttp.Status <> 200) Then
		'HTTP ������
		Response.Write ("Status=" & objHttp.Status)
	ElseIf (objHttp.ResponseText = "VERIFIED") Then
		'֧��֪ͨ��֤�ɹ�
		If Trim(v_mid) = Trim(AccountID) Then '�жϴ˶����ǲ��Ǹ��̻��Ķ�����
			PaySuccess = True
		End If
	ElseIf (objHttp.ResponseText = "INVALID") Then
		'֧��֪ͨ��֤ʧ��
		Response.Write ("Invalid")
	Else
		'֧��֪ͨ��֤�����г��ִ���
		Response.Write ("Error")
	End If
	Set objHttp = Nothing
	
	If PaySuccess = True Then
		Call UpdateOrder(v_amount,"���߳�ֵ��������Ϊ:" & v_oid,v_oid,"")
		Call ShowResult("��ϲ�㣡����֧���ɹ���")
	Else
		Call ShowResult("����֧��ʧ�ܣ�")
	End If
End Sub

'�׸�ͨ
Sub xpayBack()
	Dim v_mid, v_oid, v_pmode, v_pstatus, v_pstring, v_amount, v_md5, v_date, v_moneytype
	Dim md5string,v_sid
	v_mid = AccountID
	v_oid = Trim(Request("bid"))       '֧��������
	v_sid = Trim(Request("sid"))         '�׸�ͨ���׳ɹ� ��ˮ��
	v_md5 = Trim(Request("md"))       '����ǩ��
	v_amount = Trim(Request("prc"))       '֧�����
	v_pstatus = Trim(Request("success"))       '֧��״̬
	v_pmode = Trim(Request("bankcode"))       '֧������
	v_pstring = Trim(Request("v_pstring"))       '֧�����˵��
	
	md5string = MD5(MD5Key & ":" & v_oid & "," & v_sid & "," & v_amount & ",sell,," & v_mid & ",bank," & v_pstatus, 32)
	
	If UCase(v_md5) = UCase(md5string) And LCase(v_pstatus) = "true" Then
		Call UpdateOrder(v_amount,"���߳�ֵ��������Ϊ:" & v_oid,v_oid,v_pmode)
		Call ShowResult("��ϲ�㣡����֧���ɹ���")
	Else
		Call ShowResult("����֧��ʧ�ܣ�")
	End If
End Sub

'����֧��
Sub cncardBack
	Dim PaySuccess:PaySuccess = False
	Dim v_mid, v_oid, v_pmode, v_pstatus, v_pstring, v_amount, v_md5, v_date, v_moneytype
	Dim md5string
	Dim c_mid, c_order, c_orderamount, c_ymd, c_transnum, c_succmark, c_moneytype, c_cause, c_memo1, c_memo2, c_signstr
	c_mid = Request("c_mid")                    '�̻���ţ��������̻��ɹ��󼴿ɻ�ã������������̻��ɹ����ʼ��л�ȡ�ñ��
	c_order = Request("c_order")                '�̻��ṩ�Ķ�����
	c_orderamount = Request("c_orderamount")    '�̻��ṩ�Ķ����ܽ���ԪΪ��λ��С���������λ���磺13.05
	c_ymd = Request("c_ymd")                    '�̻���������Ķ����������ڣ���ʽΪ"yyyymmdd"����20050102
	c_transnum = Request("c_transnum")          '����֧�������ṩ�ĸñʶ����Ľ�����ˮ�ţ����պ��ѯ���˶�ʹ�ã�
	c_succmark = Request("c_succmark")          '���׳ɹ���־��Y-�ɹ� N-ʧ��
	c_moneytype = Request("c_moneytype")        '֧�����֣�0Ϊ�����
	c_cause = Request("c_cause")                '�������֧��ʧ�ܣ����ֵ����ʧ��ԭ��
	c_memo1 = Request("c_memo1")                '�̻��ṩ����Ҫ��֧�����֪ͨ��ת�����̻�����һ
	c_memo2 = Request("c_memo2")                '�̻��ṩ����Ҫ��֧�����֪ͨ��ת�����̻�������
	c_signstr = Request("c_signstr")            '����֧�����ض�������Ϣ����MD5���ܺ���ַ���
	
	md5string = MD5(c_mid & c_order & c_orderamount & c_ymd & c_transnum & c_succmark & c_moneytype & c_memo1 & c_memo2 & MD5Key, 32)
	
	If UCase(md5string) <> UCase(c_signstr) Then
		Response.Write "ǩ����֤ʧ��"
		Response.End
	End If
	
	If Trim(AccountID) <> c_mid Then
		Response.Write "�ύ���̻��������"
		Response.End
	End If
	
	If c_succmark <> "Y" And c_succmark <> "N" Then
		Response.Write "�����ύ����"
		Response.End
	End If
	
	PaySuccess = True
	v_oid = c_order
	v_amount = c_orderamount
	v_pstring = ""
	v_pmode = ""
	If PaySuccess = True Then
		Call UpdateOrder(v_amount,"���߳�ֵ��������Ϊ:" & v_oid,v_oid,v_pmode)
		Call ShowResult("��ϲ�㣡����֧���ɹ���")
	Else
		Call ShowResult("����֧��ʧ�ܣ�")
	End If
End Sub

'��Ǯ֧��
Sub billback()
Dim md5string
Dim merchantAcctId, key, version, language, signType, payType, bankId, orderId, orderTime, orderAmount, dealId, bankDealId, dealTime, payAmount
Dim fee, ext1, ext2, payResult, errCode, signMsg, merchantSignMsgVal

merchantAcctId = Trim(request("merchantAcctId")) '��ȡ����������˻���
key = MD5Key '���������������Կ
version = Trim(request("version")) '��ȡ���ذ汾
language = Trim(request("language")) '��ȡ��������,1�������ģ�2����Ӣ��
signType = Trim(request("signType")) 'ǩ������,1����MD5ǩ��
payType = Trim(request("payType")) '��ȡ֧����ʽ,00�����֧��,10�����п�֧��,11���绰����֧��,12����Ǯ�˻�֧��,13������֧��,14��B2B֧��
bankId = Trim(request("bankId")) '��ȡ���д���
orderId = Trim(request("orderId")) '��ȡ�̻�������
orderTime = Trim(request("orderTime")) '��ȡ�����ύʱ��
orderAmount = Trim(request("orderAmount")) '��ȡԭʼ�������
dealId = Trim(request("dealId")) '��ȡ��Ǯ���׺�
bankDealId = Trim(request("bankDealId")) '��ȡ���н��׺�
dealTime = Trim(request("dealTime")) '��ȡ�ڿ�Ǯ����ʱ��
payAmount = Trim(request("payAmount")) '��ȡʵ��֧�����,��λΪ��
fee = Trim(request("fee")) '��ȡ����������
ext1 = Trim(request("ext1")) '��ȡ��չ�ֶ�1
ext2 = Trim(request("ext2")) '��ȡ��չ�ֶ�2

'��ȡ������
''10���� �ɹ�11���� ʧ��
''00���� �¶����ɹ������Ե绰����֧���������أ�;01���� �¶���ʧ�ܣ����Ե绰����֧���������أ�
payResult = Trim(request("payResult"))
errCode = Trim(request("errCode")) '��ȡ�������,��ϸ���ĵ���������б�
signMsg = Trim(request("signMsg")) '��ȡ����ǩ����

'���ɼ��ܴ������뱣������˳��
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

''���Ƚ���ǩ���ַ�����֤
If UCase(signMsg) = UCase(md5string) Then
    ''���Ž���֧������ж�
    Select Case payResult
          Case "10"   '֧���ɹ������¶���
            rtnOk = 1
			Call UpdateOrder(orderAmount / 100,"���߳�ֵ��������Ϊ:" & orderId,orderId,"")
			Call ShowResult("��ϲ�㣡����֧���ɹ���")
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
'֧����
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
	
	v_oid = DelStr(Request("out_trade_no"))            '�̻�������
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
	'��ȡ֧����GET����֪ͨ��Ϣ,�ж���Ϣ�ǲ��Ǳ��޸Ĺ�
	Dim varItem
	For Each varItem in Request.QueryString
	mystr=varItem&"="&Request(varItem)&"^"&mystr
	Next 
	If mystr<>"" Then 
	mystr=Left(mystr,Len(mystr)-1)
	End If 
	mystr = SPLIT(mystr, "^")

	Count=ubound(mystr)
	'�Բ�������
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
	'����md5ժҪ�ַ���
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
		Call UpdateOrder(v_amount,"���߳�ֵ��������Ϊ:" & v_oid,v_oid,v_pmode)
		Call ShowResult("��ϲ�㣡����֧���ɹ���")
	Else
	Call ShowResult("����֧��ʧ�ܣ�")          '�������ָ������Ҫ��ʾ������
	End If 
	
End Sub

'֧�����Ǽ�ʱ����
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
    
    v_oid = DelStr(Request("out_trade_no"))            '�̻�������
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

                
    '��ȡPOST�����Ĳ���
    mystr = Split(URLDecode(Request.Form), "&")
    Count = UBound(mystr)

    '�Բ�������
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

    '����md5ժҪ�ַ���
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
    '����md5ժҪ
    MySign = MD5(md5str,32)


    '�ȴ���Ҹ���
    Select Case trade_status
    Case "WAIT_BUYER_PAY"
        If ResponseTxt = "true" And sign = MySign Then
            returnTxt = "success"
        Else
            returnTxt = "fail"
        End If

    '��Ҹ���ɹ�,�ȴ����ҷ���
    Case "WAIT_SELLER_SEND_GOODS"
        If ResponseTxt = "true" And sign = MySign Then
            returnTxt = "success"
			Conn.Execute("Update KS_Order Set Status=1 Where OrderID='" & KS.R(v_oid) & "'") 'ֻ���¶���״̬�������·���״̬�Ͷ���״̬
        Else
            returnTxt = "fail"
        End If

    '�ȴ����ȷ���ջ�
    Case "WAIT_BUYER_CONFIRM_GOODS"
        If ResponseTxt = "true" And sign = MySign Then
            returnTxt = "success"
            			Conn.Execute("Update KS_Order Set Status=1,DeliverStatus=1 Where OrderID='" & v_oid & "'") '����֧����¼״̬�ͷ���״̬�������¶���״̬
        Else
            returnTxt = "fail"
        End If

    '���׳ɹ�����
    Case "TRADE_FINISHED"
        If ResponseTxt = "true" And sign = MySign Then
            returnTxt = "success"
            PaySuccess = True                '���׳ɹ������¶���
        Else
            returnTxt = "fail"
        End If

    '��������״̬֪ͨ���
    Case Else
        If ResponseTxt = "true" And sign = MySign Then
            returnTxt = "success"
        Else
            returnTxt = "fail"
        End If
    End Select
    Response.Write returnTxt
	If PaySuccess = True Then
	 Call UpdateOrder(v_amount,"���߳�ֵ��������Ϊ:" & v_oid,v_oid,v_pmode)
	Else
	 '	Call ShowResult("����֧��ʧ�ܣ�")          '�������ָ������Ҫ��ʾ������
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
		DelStr	= Replace(DelStr,"��","")
		DelStr	= Replace(DelStr,"%20","")
		DelStr	= Replace(DelStr,"--","")
		DelStr	= Replace(DelStr,"==","")
		DelStr	= Replace(DelStr,"<","")
		DelStr	= Replace(DelStr,">","")
		DelStr	= Replace(DelStr,"%","")
End Function

'��post���ݹ����Ĳ�����urldecode���봦��(֧�������½ӿ�)
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
End Function '�������

'�Ƹ�ͨ
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
		Call UpdateOrder(v_amount,"���߳�ֵ��������Ϊ:" & v_oid,v_oid,v_pmode)
		Call ShowResult("��ϲ�㣡����֧���ɹ���")
	Else
		Call ShowResult("����֧��ʧ�ܣ�")
	End If
End Sub

'�Ƹ�ͨ�н�
Sub tenpayZJ()
%>
<html>
<head>
	<meta name="TENCENT_ONLINE_PAYMENT" content="China TENCENT">
</head>
<%
'��ȡ����
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

'����ǩ��
localSignText = UCase(md5(buffer,32) )

dim msg
'ǩ���ж�
if localSignText = sign then
	'��֤ǩ���ɹ�
	
	if retcode = "0" then
		msg = "OK"
		msg = msg & status 
		Select case status
			case "1":
				'���״���
			case "2":
				'�ջ��ַ��д���
			case "3":
				Conn.Execute("Update KS_Order Set MoneyReceipt=MoneyTotal,Status=1 Where OrderID='" & KS.R(mch_vno) & "'") '���¶���״̬���Ѹ���
                response.write "<script>alert('��ϲ��֧���ɹ�����ȴ��̼ҷ�����');location.href='../user/index.asp?user_order.asp';</script>"
			case "4":
				Conn.Execute("Update KS_Order Set DeliverStatus=1 Where OrderID='" & KS.R(mch_vno) & "'")
				'���ҷ����ɹ�
			case "5":
				Conn.Execute("Update KS_Order Set DeliverStatus=2 Where OrderID='" & KS.R(mch_vno) & "'")
				'����ջ�ȷ�ϣ����׳ɹ�
			case "6":
				'���׹رգ�δ��ɳ�ʱ�ر�
			case "7":
				'�޸Ľ��׼۸�ɹ�
			case "8":
				'��ҷ����˿�
			case "9":
				'�˿�ɹ�
			case "10":
				'�˿�ر�
			case else
				'error
		end Select

	else
		'֧��ʧ�ܣ��벻Ҫ���ɹ�����
		msg = "֧��ʧ��"
	end if

else
	'��֤ǩ��ʧ��
	msg = "��֤ǩ��ʧ��"
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
<title>�û���������</title>
<link href="images/css.css" type="text/css" rel="stylesheet" />
</head>
<body leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0"><br><br><br>
	<table class=border cellSpacing=1 cellPadding=2 width="60%" align=center border=0>
  <tr class="title"> 
    <td height=22 align=center><b><font color="#FF0000">��ʾ��</font> ����������֧������������£�</b></td>
 </tr>
 <tr class="tdbg"><td>
      <p>
        <%=message%>
	  </p>
     </td>
  </tr>
  <tr class="title">
   <td  height="22" align="center"><a href="<%=KS.getdomain%>user/index.asp">�����Ա����</a> | <a href="<%=KS.getdomain%>">������ҳ</a>
   </td>
  </tr>
</table>
<%
End Sub

'���¶�����¼
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
			 case "shop"   '�̳����Ĺ���
				 Set RS=Server.CreateObject("ADODB.RECORDSET")
				 RS.Open "Select top 1 * From KS_Order Where OrderID='" & v_oid & "'",Conn,1,3
				 If RS.Eof Then
				   RS.Close:Set RS=Nothing
				   Response.Write "<br><li>֧���������������⣬����ϵ��վ����Ա��"
				 End If
				  If Mobile="" Then
				  Mobile=RS("Mobile")
				  End If
				  RS("MoneyReceipt")=Money
				  Dim OrderStatus:OrderStatus=rs("status")
				  RS("Status")=1
				  RS("PayTime")=now   '��¼����ʱ��
				  RS.Update
                  orderid=RS("OrderID")
				  Call KS.MoneyInOrOut(rs("UserName"),RS("Contactman"),Money,2,1,now,rs("orderid"),"System","Ϊ���򶩵���" &v_oid & "ʹ��" & v_pmode & "���߳�ֵ",0,0)
		          Call KS.MoneyInOrOut(rs("UserName"),RS("Contactman"),Money,4,2,now,rs("orderid"),"System",Remark,0,0)
				  
					
					'====================Ϊ�û����ӹ���Ӧ�û���========================
					Dim rsp:set rsp=conn.execute("select point,id,title from ks_product where id in(select proid from KS_OrderItem where orderid='" & rs("orderid") & "')")
					do while not rsp.eof
					  dim amount:amount=conn.execute("select top 1 amount from ks_orderitem where orderid='" & rs("orderid") & "' and proid=" & rsp(1))(0)
					  If OrderStatus<>1 Then
					  conn.execute("update ks_product set totalnum=totalnum-" & amount &" where totalnum>=" & amount &" and id=" & rsp(1))
					 ' response.write rs("orderid") & "=55<br>"
					 ' response.write amount & "<br>"
					 ' response.write username & "<br>"
					  
					  Call KS.ScoreInOrOut(UserName,1,KS.ChkClng(rsp(0))*amount,"ϵͳ","������Ʒ<font color=red>" & rsp("title") & "</font>����!",0,0)
					  End if
					  
					rsp.movenext
					loop
					rsp.close
					set rsp=nothing
					'================================================================
					
					RS.Close:Set RS=Nothing
			 Case else   '��Ա���ĳ�ֵ
					Set rsUser=Server.CreateObject("Adodb.RecordSet")
					sqlUser="select top 1 * from KS_User where UserName='" & UserName & "'"
					rsUser.Open sqlUser,Conn,1,1
					if rsUser.bof and rsUser.eof then
								Response.Write "<br><li>��ֵ�������������⣬����ϵ��վ����Ա��"
								rsUser.close:set rsUser=Nothing
								exit sub
					end if
					Dim RealName:RealName=rsUser("RealName")
					Dim Edays:Edays=rsUser("Edays")
					Dim BeginDate:BeginDate=rsUser("BeginDate")
					rsUser.Close : Set rsUser=Nothing

					If KS.ChkClng(KS.S("UserCardID"))<>0 Then   '��ֵ��
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
								   Call KS.PointInOrOut(0,0,UserName,1,ValidNum,"System","���߹����ֵ��[" & CardTitle &"]��õĵ���",0)
								  case 2
									Dim tmpDays:tmpDays=Edays-DateDiff("D",BeginDate,now())
									if tmpDays>0 then
									    Conn.Execute("Update KS_User Set Edays=Edays+" & ValidNum & " where username='" & userName & "'") 
									else
									    Conn.Execute("Update KS_User Set Edays=" & ValidNum & ",BeginDate=" & SQLNowString& " where username='" & userName & "'") 
									end if
									Call KS.EdaysInOrOut(UserName,1,ValidNum,"System","���߹����ֵ��[" & CardTitle &"]��õ���Ч����")
                                       
								  case 3
								   	Call KS.MoneyInOrOut(UserName,RealName,ValidNum,3,1,now,v_oid,"System",v_pmode & "���߳�ֵ,���߹����ֵ��[" & CardTitle &"]��õ��ʽ�",0,0)
								  case 4
								     
			                        Call KS.ScoreInOrOut(UserName,1,ValidNum,"System","ͨ����ֵ��[" & CardTitle & "]��õĻ���!",0,0)
							   End Select
							   If RSCard("ValidUnit")<>3 Then
								   	Call KS.MoneyInOrOut(UserName,RealName,RSCard("Money"),3,1,now,v_oid,"System",v_pmode & "���߳�ֵ!",0,0)
								   	Call KS.MoneyInOrOut(UserName,RealName,RSCard("Money"),3,2,now,v_oid,"System", "Ϊ�����ֵ��[" & CardTitle &"]��֧��!",0,0)
							   End If
							 End If 
							 RSCard.Close:Set RSCard=Nothing
							 

					Else
				  	 Call KS.MoneyInOrOut(UserName,RealName,Money,3,1,now,v_oid,"System",v_pmode & "���߳�ֵ,������Ϊ:" & v_oid,0,0)
					End If

					
			 End Select
			 
		End If

		RSLog.Close:Set RSLog=Nothing
End Sub
%>