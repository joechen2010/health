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
<card id="main" title="手机神州行充值平台">
<p>
<%
Dim KSCls
Set KSCls = New User_CardOnline
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%
Class User_CardOnline
        Private KS,Prev
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		    Set KS=Nothing
		    Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
			IF Cbool(KSUser.UserLoginChecked)=false Then
			   Response.redirect KS.GetDomain&"User/Login.asp?User_CardOnline.asp"
			   Exit Sub
			End If
			IF KS.WSetting(14)="0" Then
			   Response.Write "对不起，本站暂停神州行充值卡通道。<br/>"
			   Response.Write "<anchor>返回上页<prev/></anchor><br/>"
			   Exit Sub
			End If
			Select Case KS.S("Action")
			    Case "CardStep2"
				Call CardStep2()
				Case "CardStep3"
				Call CardStep3()
				Case "CardStep4"
				Call CardStep4()
				Case "CardStep5"
				Call CardStep5()
				Case Else
				Call CardOnline()
			End Select
			If Prev=True Then
			   Response.Write "<anchor>返回上页<prev/></anchor><br/>"
			End If
			Response.Write "<br/>"
			Response.Write "<a href=""Index.asp?" & KS.WapValue & """>我的地盘</a><br/>" &vbcrlf
			Response.Write "<a href=""../?" & KS.WapValue & """>返回首页</a>" &vbcrlf
		End Sub
		
		Sub CardOnline()
		    %>
            用 户 名:<%=KSUser.UserName%><br/>
            计费方式:<%
			If KSUser.ChargeType=1 Then 
			   Response.Write "扣点数计费用户"
			ElseIf KSUser.ChargeType=2 Then
			   Response.Write "有效期计费用户,到期时间：" & Cdate(KSUser.BeginDate)+KSUser.Edays & ","
			ElseIf KSUser.ChargeType=3 Then
			   Response.Write "无限期计费用户"
			End If
			%>
            <br/>
            资金余额:<%=KSUser.Money%>元<br/>
            可用<%=KS.Setting(45)%>:<%=KSUser.Point%><%=KS.Setting(46)%><br/>
            剩余天数:<%
			If KSUser.ChargeType=3 Then
			   Response.Write "无限期"
			Else
			   Response.Write KSUser.GetEdays&"天"
			End If
			%>
            <br/><br/>
            手机神州行充值平台<br/>
            *面额必须与充值卡的实际面额一致，否则可能导致支付不成功或余额丢失。<br/>
            神州行充值卡面额：<anchor><b>50元</b><go href="User_CardOnline.asp?Action=CardStep2&amp;<%=KS.WapValue%>" method="post"><postfield name="Money" value="50"/></go></anchor>
            <anchor><b>100元</b><go href="User_CardOnline.asp?Action=CardStep2&amp;<%=KS.WapValue%>" method="post"><postfield name="Money" value="100"/></go></anchor>
            
            <br/><br/>
			手机充值卡充值帮助<br/>
            1.请选择您要充值的手机充值卡的面值，卡面值是多少就选者多少，否者会充值不成功。<br/>
            2.手机充值卡充值页面要求用户输入充值卡序列号： 充值卡密码：输入确认无误后，点击提交既可。<br/>
			<%
		End Sub
		
		Sub CardStep2()
		    Dim Money:Money=KS.S("Money")
			If Not IsNumeric(Money) Then
			   Response.Write "对不起，您输入的充值金额不正确！<br/>"
			   Prev=True
			   Exit Sub
			End If
			%>
            用户名:<%=KSUser.UserName%><br/>
            神州行充值卡面额:<b><%=Money%></b>元<br/>
            神州行充值卡卡号:<input type="text" name="CardNo<%=Minute(Now)%><%=Second(Now)%>" value="" emptyok="false" format="*N"/><br/>
            神州行充值卡密码:<input type="text" name="CardPwd<%=Minute(Now)%><%=Second(Now)%>" value="" emptyok="false" format="*N"/><br/>
            <anchor>下一步<go href="User_CardOnline.asp?Action=CardStep3&amp;<%=KS.WapValue%>" method="post">
            <postfield name="Money" value="<%=Money%>"/>
            <postfield name="CardNo" value="$(CardNo<%=Minute(Now)%><%=Second(Now)%>)"/>
            <postfield name="CardPwd" value="$(CardPwd<%=Minute(Now)%><%=Second(Now)%>)"/>
            </go></anchor>
            <anchor>上一步<prev/></anchor>
            <br/>
            <%=KS.WSetting(16)%>
			<%
		End Sub
		
		Sub CardStep3()
	        Dim Money:Money=KS.S("Money")
			If Not IsNumeric(Money) Then
			   Response.Write "对不起，您输入的充值金额不正确！<br/>"
			   Prev=True
			   Exit Sub
			End If
			Dim CardNo:CardNo=KS.S("CardNo")
			Dim CardPwd:CardPwd=KS.S("CardPwd")
			
		    If CardNo="" Then
			   Response.Write "你没有输入充值卡号！<br/>"
			   Prev=True
			   Exit Sub
			End If
			If CardPwd=" "Then
			   Response.Write "你没有输入充值卡密码！<br/>"
			   Prev=True
			   Exit Sub
			End If
			If Not Conn.Execute("select CardNum from KS_UserCard where CardNum='" & CardNo & "' And CardPass='" & KS.Encrypt(CardPwd) & "'").EOF Then
			   Response.Write "你输入的充值卡号已存在，请重输!！<br/>"
			   Prev=True
			   Exit Sub
			End If
			%>
            【确认充值卡】<br/>
            用户名:<%=KSUser.UserName%><br/>
            充值卡面值:<%=Money%>元<br/>
            充值卡卡号:<%=CardNo%><br/>
            充值卡密码:<%=CardPwd%><br/>
            <anchor>确定充值<go href="User_CardOnline.asp?Action=CardStep4&amp;<%=KS.WapValue%>" method="post">
            <postfield name="Money" value="<%=Money%>"/>
            <postfield name="CardNo" value="<%=CardNo%>"/>
            <postfield name="CardPwd" value="<%=CardPwd%>"/>
            </go></anchor>
            <br/><br/>
            <anchor>返回上步<prev/></anchor>
			<%
		End Sub
		
		Sub CardStep4()
		    Dim UserID,Money,CardNo,CardPwd
		    UserID = KS.WSetting(17)'合作ID
			Money = KS.S("Money")'充值卡面值
			CardNo = KS.S("CardNo")
			CardPwd = KS.S("CardPwd")
		    If CardNo="" Then
			   Response.Write "非法操作！<br/>"
			   Prev=True
			   Exit Sub
			End If
			If CardPwd=" "Then
			   Response.Write "非法操作！<br/>"
			   Prev=True
			   Exit Sub
			End If
			Dim MyString
			MyString = "UserID=" & UserID & "&CardNo=" & CardNo & "&CardPwd=" & Cardpwd & "&Money=" & Money & "&OrderID=" & CardNo & ""
			
			Dim CardHttp,CardChecked
			set CardHTTP = CreateObject("Microsoft.XMLHTTP")
			CardHTTP.Open "POST", "http://121.11.91.149/szxInterFace.asp", False
			CardHTTP.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
			CardHTTP.Send MyString
			If (CardHTTP.status <> 200 ) Then
			   Response.Write "网络繁忙，请稍后再充值...<br/>"
			   Response.Write "错误提示：Status=" & CardHTTP.status & "<br/>"'HTTP 错误处理
			Else
			   CardChecked = CardHTTP.ResponseText
			   If CardChecked = "1" Then
			      '交易提交成功逻辑处理
				  Dim RS:Set RS=Server.CreateObject("adodb.recordset")
				  RS.Open "select * from KS_UserCard",Conn,1,3
				  If Conn.Execute("select CardNum from KS_UserCard where CardNum='" & CardNo & "'").EOF Then
				     RS.Addnew
				     RS("CardNum")=CardNo'充值卡卡号
				  End If
				  RS("CardPass")=KS.Encrypt(CardPwd)'充值卡密码
				  RS("Money")=Money'充值卡面值
				  RS("ValidNum")=Money'充值卡点数、资金或有效期
				  RS("ValidUnit")=3'1--点,2--天,3--元
				  RS("AddDate")=Date()
				  RS("EndDate")=Dateadd("d",2,Now())'充值截止期限2天
				  RS("IsUsed")=0'是否使用
				  RS("IsSale")=1'是否出售
				  'RS("GroupName")=GroupName'分类名称
				  RS("UserName")=KSUser.UserName'使用者
				  RS.Update
				  RS.Close:set RS=Nothing
				  
				  Response.Write "定单提交成功!<br/>正在验证充值卡，可能需要三分钟左右，请你耐心等待！<br/>"
				  Response.Write "<a href=""User_CardOnline.asp?Action=CardStep5&amp;CardNo="&CardNo&"&amp;"&KS.WapValue&""">查询状态</a><br/>"
			   ElseIf CardChecked = "2" Then
			      Response.Write "充值失败！<br/>"
			   ElseIf CardChecked = "3" Then
			      Response.Write "验证失败！<br/>"
			   Else
			      Response.Write "非法操作！<br/>"
			   End If
			End If
			set CardHTTP = Nothing
		End Sub
		
		Sub CardStep5()
		    Dim CardNo
			CardNo = KS.S("CardNo")
			Dim RS:Set RS=Server.CreateObject("adodb.recordset")
			RS.Open "select * from KS_UserCard where CardNum='" & CardNo & "'",Conn,1,1
			If RS("IsUsed")=0 And RS("IsSale")=1 Then
			   Response.Write "提示：正在验证您的充值卡，由于移动网关验证充值卡较慢，可能需要三分钟左右，请你耐心等待，不要按返回键。<br/>"
			   Response.Write "<a href=""User_CardOnline.asp?Action=CardStep5&amp;CardNo="&CardNo&"&amp;"&KS.WapValue&""">再次查询</a><br/>"
			ElseIf RS("IsUsed")=0 And RS("IsSale")=0 Then
			   Response.Write "对不起，充值失败，请您再次输入。如多次尝试无法成功，请参考“相关帮助”。<br/>"
			   Response.Write "<a href=""User_CardOnline.asp?"&KS.WapValue&""">返回重试</a><br/>"
			   Response.Write "相关帮助<br/>"
			   Response.Write "<a href=""../plus/Template.asp?id=20086806230945&amp;"&KS.WapValue&""">为什么我一直充值不成功</a><br/>"
			   Response.Write "<a href=""../plus/Template.asp?id=20084783814868&amp;"&KS.WapValue&""">购买充值卡注意事项</a><br/>"
			   Response.Write "<a href=""../plus/Template.asp?id=20086538627961&amp;"&KS.WapValue&""">如何给自己帐户充值</a><br/>"
			   Response.Write "<a href=""../plus/Template.asp?id=20089621401105&amp;"&KS.WapValue&""">联通用户如何使用神州行卡支付</a><br/>"
			End If
		End Sub
End Class
%> 
