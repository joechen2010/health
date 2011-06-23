<% @LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="查询我的资金明细">
<p>
<%
Dim KSCls
Set KSCls = New User_LogMoney
KSCls.Kesion()
Set KSCls = Nothing
%>
</p>
</card>
</wml>
<%
Class User_LogMoney
        Private KS
		Private CurrentPage,totalPut
		Private RS,MaxPerPage
		Private TempStr,SqlStr
		Private Sub Class_Initialize()
			MaxPerPage =10
			Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		    Set KS=Nothing
		    Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		    IF Cbool(KSUser.UserLoginChecked)=False Then
			   Response.redirect KS.GetDomain&"User/Login/"
			   Exit Sub
			End If
			Dim IncomeOrPayOut :IncomeOrPayOut = KS.ChkClng(KS.S("IncomeOrPayOut"))
			Response.Write KS.GetReadMessage
			IF IncomeOrPayOut="" or IncomeOrPayOut="1" or IncomeOrPayOut="2" Then
			   Response.Write "<a href=""User_LogMoney.asp?"&KS.WapValue&""">所有记录</a> "
			Else
			   Response.Write "所有记录 "
			End If
			IF IncomeOrPayOut="1" Then
               Response.Write "收入["&Conn.Execute("select count(id) from ks_logmoney where IncomeOrPayOut=1 and username='" & KSUser.UserName & "'")(0)&"] "
            Else
               Response.Write "<a href=""User_LogMoney.asp?IncomeOrPayOut=1&amp;"&KS.WapValue&""">收入["&Conn.Execute("select count(id) from ks_logmoney where IncomeOrPayOut=1 and username='" & KSUser.UserName & "'")(0)&"]</a> "
            End If
            IF IncomeOrPayOut="2" Then
               Response.Write "支出["&Conn.Execute("select count(id) from ks_logmoney where IncomeOrPayOut=2 and username='" & KSUser.UserName & "'")(0)&"]"
            Else
               Response.Write "<a href=""User_LogMoney.asp?IncomeOrPayOut=2&amp;"&KS.WapValue&""">支出["&Conn.Execute("select count(id) from ks_logmoney where IncomeOrPayOut=2 and username='" & KSUser.UserName & "'")(0)&"]</a>"
            End If
            Response.Write "<br/>【资金明细】<br/>"
			
			If KS.ChkClng(IncomeOrPayOut)=1 Or KS.ChkClng(IncomeOrPayOut)=2 Then
			   SqlStr="Select * From KS_LogMoney Where IncomeOrPayOut=" & KS.ChkClng(IncomeOrPayOut) & " And  UserName='" & KSUser.UserName &"' order by id desc"
			Else
			   SqlStr="Select * From KS_LogMoney Where UserName='" & KSUser.UserName &"' order by id desc"
			End if
			If KS.S("page") <> "" Then
			   CurrentPage = CInt(KS.S("page"))
			Else
			   CurrentPage = 1
			End If
			Set RS=Server.createobject("adodb.recordset")
			RS.Open SqlStr,Conn,1,1
			If RS.EOF And RS.BOF Then
			   Response.Write "<br/>"
			   Response.Write "找不到您要的记录!<br/>"
			Else
			   MaxPerPage =3
			   totalPut = RS.RecordCount
			   If CurrentPage < 1 Then	CurrentPage = 1
			   If (CurrentPage - 1) * MaxPerPage > totalPut Then
			      If (totalPut Mod MaxPerPage) = 0 Then
				     CurrentPage = totalPut \ MaxPerPage
				  Else
				     CurrentPage = totalPut \ MaxPerPage + 1
				  End If
			   End If
			   If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
			      RS.Move (CurrentPage - 1) * MaxPerPage
			   Else
			      CurrentPage = 1
			   End If
			   Dim I,intotalmoney,outtotalmoney
			   Do While Not RS.Eof
			   %>
               
               交易时间:<%=RS("LogTime")%><br/>
               用户名:<%=RS("username")%><br/>
               客户姓名:<%=RS("clientname")%><br/>
               交易方式:
			   <%
			   Select Case RS("MoneyType")
			       Case 1:Response.WRite "现金"
				   Case 2:Response.Write "银行汇款"
				   Case 3:Response.Write "在线支付"
				   Case 4:Response.Write "资金余额"
			   End Select
			   %><br/>

			   <%
			   If RS("IncomeOrPayOut")=1 Then
			      Response.Write "收入金额:"
				  intotalmoney=intotalmoney+RS("money")
			   Else
			      Response.Write "支出金额:"
				  outtotalmoney=outtotalmoney+RS("money")
			   End If
			   Response.Write formatnumber(RS("money"),2)&"(人民币)<br/>"
			   Response.Write "备注说明:"&RS("Remark")&"<br/>"
			   Response.Write "<img src=""../Images/Hen.gif"" alt=""""/><br/>"
			   I = I + 1
			   RS.MoveNext
			   If I >= MaxPerPage Then Exit Do
			   loop
			   Call KS.ShowPageParamter(totalPut, MaxPerPage, "User_LogMoney.asp", True, "条记录", CurrentPage, "IncomeOrPayOut=" & KS.ChkClng(KS.S("IncomeOrPayOut"))&"&amp;" & KS.WapValue & "")
			   Response.Write "<br/>【本页合计】<br/>"
			   Response.Write "收入金额::"&formatnumber(intotalmoney,2)&"<br/>"
			   Response.Write "支出金额::"&formatnumber(outtotalmoney,2)&"<br/>"
			   
			   intotalmoney=Conn.Execute("Select Sum(Money) From KS_Logmoney Where UserName='" & KSUser.UserName & "' And IncomeOrPayOut=1")(0)
			   outtotalmoney=Conn.Execute("Select Sum(Money) From KS_Logmoney Where UserName='" & KSUser.UserName & "' And IncomeOrPayOut=2")(0)
			   If not isnumeric(intotalmoney) Then intotalmoney=0
			   If not isnumeric(outtotalmoney) Then outtotalmoney=0
			   Response.Write "【所有合计】<br/>"
			   Response.Write "收入金额:"&formatnumber(intotalmoney,2)&"<br/>" &vbcrlf
			   Response.Write "支出金额:"&formatnumber(outtotalmoney,2)&"<br/>" &vbcrlf
			   Response.Write "资金累计余额："&formatnumber(intotalmoney-outtotalmoney,2)&"<br/>" &vbcrlf
			   
            End If
			Response.Write "<br/>" &vbcrlf
			Response.Write "<a href=""Index.asp?" & KS.WapValue & """>我的地盘</a><br/>" &vbcrlf
			Response.Write "<a href=""" & KS.GetGoBackIndex & """>返回首页</a>" &vbcrlf
        End Sub
End Class
%>


