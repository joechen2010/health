<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.Buffer = True 
Response.Expires = 0 
Response.CacheControl = "no-cache"

Response.ContentType="text/vnd.wap.wml"
Response.Charset="utf-8"
Response.Write "<?xml version=""1.0"" encoding=""utf-8""?>" &vbcrlf
Response.Write "<!DOCTYPE wml PUBLIC ""-//WAPFORUM//DTD WML 1.1//EN"" ""http://www.wapforum.org/DTD/wml_1.1.xml"">" &vbcrlf
%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Md5.asp"-->
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="神州行充值接口">
<p>
<%
Dim KS:Set KS=New PublicCls
Call SpvnowBank()
%>
</p>
</card>
</wml>
<%
'万利联盟
Sub SpvnowBank()
    Dim MD5Key,UserID,sOrder,SendState,Money,sNewString,SendTime,OrderID
	MD5Key = KS.WSetting(18)'固定密匙
	UserID = KS.S("UserID")'合作ID
	sOrder = KS.S("sOrder")'代理方订单号
	SendState = KS.S("SendState")'1--成功，2--失败
	Money = KS.S("Money")'充值金额
	sNewString = Request("sNewString")'原始签名串
	SendTime = KS.S("SendTime")'状态时间
	OrderID = KS.S("OrderID")'订单号
	Dim RemoteIP
	RemoteIP = Request.ServerVariables("REMOTE_ADDR")'取得代理方的IP地址
	If RemoteIP <> "121.11.91.149" Then
	   Response.Write "MD5 error<br/>"
	   Exit Sub
	End If
	Dim Text,Md5Text
	Text = SendState & OrderID & sOrder & Money & MD5Key'Md5校验
	Md5Text = Md5(Text,32)'商户拼凑的Md5校验串
	If Md5Text <> sNewString Then
	   Response.Write "充值失败！<br/>"
	   Exit Sub
	End If
	Dim UserName,RS
	Set RS=Server.CreateObject("Adodb.RecordSet")
	RS.Open "select UserName from KS_UserCard where CardNum='" & OrderID & "'",Conn,1,1
	If Not RS.EOF Then
	   UserName=RS("UserName")
	Else
	   Response.Write "充值失败！<br/>"
	   Exit Sub
	End If
	If SendState="2" Then
	   '置充值卡已使用
	   Conn.Execute("Update KS_UserCard Set IsUsed=0,IsSale=0,UseDate=" & SqlNowString & " where CardNum='" & OrderID & "'")
	   Call KS.SendInfo(UserName,"System","神州行充值卡,充值失败！",""&UserName&"您好!你使用神州行充值卡充值失败,时间:"&SendTime&",订单号:"&sOrder&",请查看充值卡卡号和充值卡密码是否正确。")
	   Response.Write "充值失败！<br/>"
	   Exit Sub
	End If
	Dim RSLog:Set RSLog=Server.CreateObject("ADODB.RECORDSET")
	RSLog.Open "Select * From KS_LogMoney where OrderID='" & sOrder & "'",Conn,1,3
	If RSLog.EOF And RSLog.BOF Then
	   Dim rsUser,sqlUser
	   Set rsUser=Server.CreateObject("Adodb.RecordSet")
	   sqlUser="select * from KS_User where UserName='" & UserName & "'"
	   rsUser.Open sqlUser,Conn,1,3
	   If rsUser.BOF And rsUser.EOF Then
	      rsUser.Close:set rsUser=Nothing
		  Response.Write "充值过程中遇到问题，请联系网站管理员！<br/>"
		  Exit Sub
	   End If
	   rsUser("Money")=rsUser("Money")+Money
	   rsUser.Update
	   Call KS.SendInfo(UserName,"System","神州行充值卡，充值成功！",""&UserName&"您好!你使用神州行充值卡充值成功,时间:"&SendTime&",订单号:"&sOrder&",当前资金余额:"&rsUser("Money")&"元。")
	   '置充值卡已使用
	   Conn.Execute("Update KS_UserCard Set IsUsed=1,IsSale=1,UseDate=" & SqlNowString & " where CardNum='" & OrderID & "'")
	   RSLog.AddNew
	   RSLog("UserName")=UserName'用户名
	   RSLog("ClientName")=rsUser("RealName")'客户姓名
	   RSLog("CurrMoney")=rsUser("Money")
	   RSLog("Money")=Money'收入金额
	   RSLog("MoneyType")=3 '在线支付
	   RSLog("IncomeOrPayOut")=1 '摘要
	   RSLog("OrderID")=sOrder'订单号
	   RSLog("Remark")="通过神州行充值卡获得的资金,卡号为:" & OrderID
	   RSLog("PayTime")=Now
	   RSLog("LogTime")=Now'交易时间
	   RSLog("Inputer")="System"
	   RSLog("IP")=KS.GetIP
	   RSLog.Update
	   rsUser.Close:set rsUser=Nothing
	End If
    RSLog.Close:Set RSLog=Nothing
End Sub
%>
