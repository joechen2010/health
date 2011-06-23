<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<% Response.ContentType="text/vnd.wap.wml; charset=utf-8" %><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="max-age=0"/>
<meta http-equiv="Cache-Control" content="no-cache"/>
</head>
<card id="card1" title="充值卡充值">  
<p>
<%
Dim KSCls
Set KSCls = New User_Recharge
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%
Class User_Recharge
        Private KS
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		    Set KSUser=Nothing
			Set KS=Nothing
			Call CloseConn
		End Sub
		Public Sub Kesion()
		    IF Cbool(KSUser.UserLoginChecked)=false Then
			   Response.Redirect DomainStr&"User/Login/"
			   Exit Sub
			End If
			Select Case KS.S("Action")
			    Case "SaveExchangeEdays"
				Call SaveExchangeEdays()
				Case Else
				Call ExchangeEdays()
			End Select
			Response.Write "<a href=""Index.asp?" & KS.WapValue & """>我的地盘</a><br/>" &vbcrlf
			Response.Write "<a href=""" & KS.GetGoBackIndex & """>返回首页</a>" &vbcrlf
        End Sub
		
		Sub ExchangeEdays()
		    Response.Write "充值卡充值<br/>" &vbcrlf
			Response.Write "用户名:" & KSUser.UserName & "<br/>" &vbcrlf
            Response.Write "计费方式:" &vbcrlf
            If KSUser.ChargeType=1 Then 
               Response.Write "扣点数计费用户<br/>" &vbcrlf
            ElseIf KSUser.ChargeType=2 Then
               Response.Write "有效期计费用户,到期时间：" & Cdate(KSUser.BeginDate)+KSUser.Edays & "<br/>" &vbcrlf
            ElseIf KSUser.ChargeType=3 Then
               Response.Write "无限期计费用户<br/>" &vbcrlf
            End If
			Response.Write "资金余额:" & KSUser.Money & "元<br/>" &vbcrlf
			Response.Write "可用" & KS.Setting(45) & ":" & KSUser.Point & KS.Setting(46) & "<br/>" &vbcrlf
			Response.Write "剩余天数:" &vbcrlf
			If KSUser.ChargeType=3 Then
			   Response.Write "无限期<br/>" &vbcrlf
			Else
			   Response.Write KSUser.GetEdays & "天<br/>" &vbcrlf
			End If
			Response.Write "充值卡卡号:<input name=""CardNum" & Minute(Now) & Second(Now) & """ type=""text"" value="""" maxlength=""50"" /><br/>" &vbcrlf
            Response.Write "充值卡密码:<input name=""CardPass" & Minute(Now) & Second(Now) & """ type=""text"" value="""" maxlength=""50"" /><br/>" &vbcrlf
            Response.Write "<anchor>确定充值<go href=""User_ReCharge.asp?" & KS.WapValue & """ method=""post"">" &vbcrlf
            Response.Write "<postfield name=""Action"" value=""SaveExchangeEdays""/>" &vbcrlf
            Response.Write "<postfield name=""Premoney"" value=""" & KSUser.Money & """/>" &vbcrlf
            Response.Write "<postfield name=""CardNum"" value=""$(CardNum" & Minute(Now) & Second(Now) & ")""/>" &vbcrlf
            Response.Write "<postfield name=""CardPass"" value=""$(CardPass" & Minute(Now) & Second(Now) & ")""/>" &vbcrlf
            Response.Write "</go></anchor><br/><br/>" &vbcrlf
        End Sub
		
		Sub SaveExchangeEdays()
		    Dim ChangeType:ChangeType=KS.S("ChangeType")
			Dim Money:Money=KS.S("Money")
			DiM CardNum:CardNum=KS.S("CardNum")
			Dim CardPass:CardPass=KS.S("CardPass")
			Dim RS:Set RS=Server.CreateObject("adodb.recordset")
			RS.Open "select * from ks_usercard where cardnum='" & CardNum & "'",Conn,1,1
			If RS.BOF And RS.EOF Then
			   RS.Close:set RS=Nothing
			   Response.Write "对不起，您输入的充值卡号不正确！<br/><br/>"
			   Exit Sub
			End If
			If RS("cardpass")<>KS.Encrypt(cardpass) Then
			   RS.Close:set RS=Nothing
			   Response.Write "对不起，您输入的充值卡密码不正确！<br/><br/>"
			   Exit Sub
			End If
			If RS("isused")=1 Then
			   RS.Close:set RS=Nothing
			   Response.Write "对不起，您输入的充值卡已被使用！<br/><br/>"
			   Exit Sub
			End If
			If datediff("d",RS("enddate"),now())>0 Then
			   RS.Close:set RS=Nothing
			   Response.Write "对不起，您输入的充值卡已过期！<br/><br/>"
			   Exit Sub
			End If
			Dim ValidNum:ValidNum=RS("ValidNum")
			Dim ValidUnit:ValidUnit=RS("ValidUnit")
			RS.Close
			
			RS.Open "select * from ks_user Where UserName='" & KSUser.UserName & "'",Conn,1,3
			If not RS.EOF Then
			   If RS("ChargeType")=3 And ValidUnit<>3 Then
			      RS.Close:set RS=Nothing
				  Response.Write "由于你的账户永不过期，如需充值资金，请购买资金卡！<br/><br/>"
				  Exit Sub
			   End If
			   Dim ValidDays,tmpdays
			   Select Case ValidUnit
			       Case 1 '点数
				      'RS("point")=RS("point")+ValidNum
					  Call KS.PointInOrOut(0,0,RS("UserName"),1,ValidNum,"System","通过充值卡获得的点数")
				   Case 2 '天数
				      ValidDays=RS("Edays")
					  tmpDays=ValidDays-DateDiff("D",RS("BeginDate"),now())
					  If tmpDays>0 Then
					     RS("Edays")=RS("Edays")+ValidNum
					  Else
					     RS("BeginDate")=now
					     RS("Edays")=ValidNum
					  End If
					  Call KS.EdaysInOrOut(RS("UserName"),1,ValidNum,"System","通过充值卡获得的有效天数")
				   Case 3 '金币
				      RS("money")=RS("money")+ValidNum
					  Dim RSLog:Set RSLog=Server.CreateObject("ADODB.RECORDSET")
					  RSLog.Open "Select * From KS_LogMoney",Conn,1,3
					  RSLog.AddNew
					  RSLog("UserName")=RS("UserName")
					  RSLog("ClientName")=RS("RealName")
					  RSLog("Money")=ValidNum
					  RSLog("MoneyType")=4       
					  RSLog("IncomeOrPayOut")=1  '收入
					  RSLog("OrderID")="0"
					  RSLog("Remark")="通过充值卡获得的资金"
					  RSLog("PayTime")=Now
					  RSLog("LogTime")=Now
					  RSLog("Inputer")="System"
					  RSLog("IP")=KS.GetIP
					  RSLog("CurrMoney")=RS("Money")
					  RSLog("ChannelID")=0
					  RSLog("InfoID")=0
					  RSLog.Update
					  RSLog.Close:Set RSLog=Nothing
			   End Select
			   RS.Update
			End If
			'置充值卡已使用、已售出
			Conn.Execute("Update KS_UserCard Set Isused=1,issale=1,username='" & KSUser.UserName & "',UseDate=" & SqlNowString & " where cardnum='" & cardnum & "'")
			Response.Write "恭喜您，充值成功!<br/><br/>"
			RS.Close:Set RS=Nothing
		End Sub
End Class
%>