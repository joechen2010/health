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
Dim PaymentPlat:PaymentPlat=7

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

Call alipayBack()

'支付宝即时到账
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
	For Each varItem in Request.Form
	mystr=varItem&"="&Request.Form(varItem)&"^"&mystr
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
	If ResponseTxt="true"  Then 	
		Call UpdateOrder(v_amount,"在线充值，订单号为:" & v_oid,v_oid,v_pmode)
		response.write "success"
		  '("恭喜你！在线支付成功！")
	Else
	    response.write "fail"
	       'errror        '这里可以指定你需要显示的内容
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
			
				 Set RS=Server.CreateObject("ADODB.RECORDSET")
				 RS.Open "Select top 1 * From KS_Order Where OrderID='" & v_oid & "'",Conn,1,3
				 If RS.Eof Then
				   RS.Close:Set RS=Nothing
				   
				   '会员中心充值
							Set rsUser=Server.CreateObject("Adodb.RecordSet")
							sqlUser="select top 1 * from KS_User where UserName='" & UserName & "'"
							rsUser.Open sqlUser,Conn,1,2
							if rsUser.bof and rsUser.eof then
								Response.Write "fail"
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
				   
				  ' Response.Write "<br><li>支付过程中遇到问题，请联系网站管理员！"
				Else
				  If Mobile="" Then
				  Mobile=RS("Mobile")
				  End If
				  RS("MoneyReceipt")=Money
				  RS("PayTime")=now   '记录付款时间
                  Dim OrderStatus:OrderStatus=rs("status")
				  RS("Status")=1
				  RS.Update
                  orderid=RS("OrderID")
				  Call KS.MoneyInOrOut(rs("UserName"),RS("Contactman"),Money,2,1,now,rs("orderid"),"System","为购买订单：" &v_oid & "使用" & v_pmode & "在线充值",0,0)
		          Call KS.MoneyInOrOut(rs("UserName"),RS("Contactman"),Money,4,2,now,rs("orderid"),"System",Remark,0,0)

					'====================为用户增加购物应得积分========================
					Dim rsp:set rsp=conn.execute("select point,id,title from ks_product where id in(select proid from KS_OrderItem where orderid='" & rs("orderid") & "')")
					do while not rsp.eof
					  dim amount:amount=conn.execute("select top 1 amount from ks_orderitem where orderid='" & rs("orderid") & "' and proid=" & rsp(1))(0)
					  if OrderStatus<>1 Then
					  conn.execute("update ks_product set totalnum=totalnum-" & amount &" where totalnum>=" & amount &" and id=" & rsp(1))         '扣库存量
					 ' response.write rs("orderid") & "=55<br>"
					 ' response.write amount & "<br>"
					 ' response.write username & "<br>"
					  
					  Call KS.ScoreInOrOut(UserName,1,KS.ChkClng(rsp(0))*amount,"系统","购买商品<font color=red>" & rsp("title") & "</font>赠送!",0,0)
					 End If

					  
					rsp.movenext
					loop
					rsp.close
					set rsp=nothing
					'================================================================
					
					RS.Close:Set RS=Nothing

			 End If
		End If

		RSLog.Close:Set RSLog=Nothing
End Sub
%>