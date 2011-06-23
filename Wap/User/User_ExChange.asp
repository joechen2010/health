<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml"><wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<%
Dim KSCls
Set KSCls = New User_ExChange
KSCls.Kesion()
Set KSCls = Nothing
%>
</card>
</wml>
<%
Class User_ExChange
        Private KS,Action,Prev
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		    Set KS=Nothing
		End Sub
		Public Sub Kesion()
		    IF Cbool(KSUser.UserLoginChecked)=false Then
			   Response.Redirect KS.GetDomain&"User/Login/"
			   Exit Sub
			End If
			Action=KS.S("Action")
			If Action="TypePoint" or Action="ShowPoint"  or Action="SavePoint" Then
			   Response.Write "<card id=""main"" title=""兑换" & KS.Setting(45) & """>"
			End IF
			If Action="TypeEdays" or Action="ShowEdays" or Action="SaveEdays" Then
			   Response.Write "<card id=""main"" title=""兑换有效期"">"
			End If
			If Action="TypeMoney" or Action="ShowMoney" or Action="SaveMoney" Then
			   Response.Write "<card id=""main"" title=""" & KS.Setting(45) & "兑换账户资金"">"
			End If
			Response.Write "<p>"
			Select Case KS.S("Action")
			    '兑换点券
			    Case "TypePoint":Call TypePoint()
			    Case "ShowPoint":Call ShowPoint()
				Case "SavePoint":Call SavePoint()
				'兑换账户资金
				Case "ShowMoney":Call ShowMoney()
				Case "SaveMoney":Call SaveMoney()
				'兑换有效天数
				Case "TypeEdays":Call TypeEdays()
				Case "ShowEdays":Call ShowEdays()
				Case "SaveEdays":Call SaveEdays()
			End Select
			If Prev=True Then
			   Response.Write "<anchor>返回上一页<prev/></anchor><br/>"
			End If
			Response.Write "<br/>"
			Response.Write "<a href=""Index.asp?" & KS.WapValue & """>我的地盘</a><br/>" &vbcrlf
			Response.Write "<a href=""" & KS.GetGoBackIndex & """>返回首页</a>" &vbcrlf
			Response.Write "</p>" &vbcrlf
        End Sub
		
		Sub TypePoint()
		    %>
            【兑换<%=KS.Setting(45)%>】<br/>
            用 户 名:<%=KSUser.UserName%><br/>
            资金余额:<%=KSUser.Money%> 元<br/>
            可用积分:<%=KSUser.Score%> 分<br/>
            可用<%=KS.Setting(45)%>:<%=KSUser.Point%><%=KS.Setting(46)%><br/><br/>
            使用<a href="User_ExChange.asp?Action=ShowPoint&amp;ChangeType=1&amp;<%=KS.WapValue%>">资金余额</a>兑换成<%=KS.Setting(45)%>,兑换比率:<%=KS.Setting(43)%>元:1<%=KS.Setting(46)%><br/><br/>
            使用<a href="User_ExChange.asp?Action=ShowPoint&amp;ChangeType=2&amp;<%=KS.WapValue%>">经验积分</a>兑换成<%=KS.Setting(45)%>,兑换比率:<%=KS.Setting(41)%>分:1<%=KS.Setting(46)%><br/>
            <%
		End Sub
		
		Sub ShowPoint()
		    Dim ChangeType:ChangeType=KS.S("ChangeType")
			Response.Write "【兑换" & KS.Setting(45) & "】<br/>" &vbcrlf
			Response.Write "用 户 名:"&KSUser.UserName&"<br/>" &vbcrlf
			If ChangeType=1 Then
			   Response.Write "资金余额:"&KSUser.Money&"元<br/>" &vbcrlf
			Else
			   Response.Write "可用积分:"&KSUser.Score&"分<br/>" &vbcrlf
			End If
            Response.Write "可用" & KS.Setting(45) & ":" & KSUser.Point & KS.Setting(46) & "<br/>" &vbcrlf
            Response.Write "<br/>" &vbcrlf
			If ChangeType=1 Then
			   If Round(KSUser.Money)>Round(KS.Setting(43)) Then
			      Response.Write "使用资金余额兑换成" & KS.Setting(45) & ",兑换比率:" & KS.Setting(43) & "元:1" & KS.Setting(46) & "<br/>" &vbcrlf
				  Response.Write "单位(元):<input maxlength=""8"" size=""6"" value=""100"" name=""Money" & Minute(Now) & Second(Now) & """ emptyok=""false"" format=""*N""/>" &vbcrlf
			   Else
			      Response.Write "您目前资金余额不足，请充值后再来兑换！<br/>" &vbcrlf
			   End If
			Else
			   If Round(KSUser.Score)>Round(KS.Setting(41)) Then
			      Response.Write "使用经验积分兑换成" & KS.Setting(45) & ",兑换比率:" & KS.Setting(41) & "分:1" & KS.Setting(46) & "<br/>" &vbcrlf
				  Response.Write "单位(分):<input maxlength=""8"" size=""6"" value=""100"" name=""Score" & Minute(Now) & Second(Now) & """ emptyok=""false"" format=""*N""/>" &vbcrlf
			   Else
			      Response.Write "您目前可用积分不足，不能兑换！<br/>" &vbcrlf
			   End If
			End If
            Response.Write "<anchor>执行兑换<go href=""User_Exchange.asp?" & KS.WapValue & """ method=""post"">" &vbcrlf
            Response.Write "<postfield name=""Action"" value=""SavePoint""/>" &vbcrlf
            Response.Write "<postfield name=""ChangeType"" value=""" & ChangeType & """/>" &vbcrlf
            Response.Write "<postfield name=""Money"" value=""$(Money" & Minute(Now) & Second(now) & ")""/>" &vbcrlf
            Response.Write "<postfield name=""Score"" value=""$(Score" & Minute(Now) & Second(now) & ")""/>" &vbcrlf
            Response.Write "</go></anchor><br/>" &vbcrlf
			Response.Write "<br/>" &vbcrlf
            Response.Write "一旦兑换成功即不可逆!!!" &vbcrlf
	    End Sub
		
		Sub SavePoint()
	        Dim ChangeType:ChangeType=KS.S("ChangeType")
			Dim Money:Money=KS.S("Money")
			Dim Score:Score=KS.ChkClng(KS.S("Score"))
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select * From KS_User Where UserName='" & KSUser.UserName & "'",Conn,1,3
			If RS.EOF Then
			   RS.Close:Set RS=Nothing
			   Response.Write "出错啦!<br/>":Prev=True:Exit Sub
		    End If
			If ChangeType=1 Then
			   If KS.ChkClng(Money)=0 Then
			      RS.Close:Set RS=Nothing
				  Response.Write "你输入的资金不正确,资金必须大于0!<br/>"
				  Prev=True
				  Exit Sub
			   End If
			   If KS.ChkClng(Money)<KS.ChkClng(KS.Setting(43)) Then
			      RS.Close:Set RS=Nothing
				  Response.Write "你输入的资金不正确,资金必须大于等于" & KS.Setting(43) &"!<br/>"
				  Prev=True
				  Exit Sub
			   End If
			   IF Round(RS("Money"))<Round(Money) Then
			      RS.Close:Set RS=Nothing
				  Response.Write "你可用资金不足，请充值后再来兑换!<br/>"
				  Prev=True
				  Exit Sub
		       End If
			   RS("Money")=RS("Money")-Money
			   'RS("Point")=RS("Point")+Money/KS.Setting(43)
			   RS.Update
			   'ChannelID,InfoID,UserName,InOrOutFlag,Point,User,Descript
			   Call KS.PointInOrOut(0,0,RS("UserName"),1,Money/KS.Setting(43),"System","账户资金兑换所得")
			   Dim RSLog:Set RSLog=Server.CreateObject("ADODB.RECORDSET")
			   RSLog.Open "Select * From KS_LogMoney",Conn,1,3
			   RSLog.AddNew
			   RSLog("UserName")=rs("UserName")
			   RSLog("ClientName")=rs("RealName")
			   RSLog("Money")=Money
			   RSLog("CurrMoney")=rs("money")
			   RSLog("MoneyType")=4
			   RSLog("IncomeOrPayOut")=2 
			   RSLog("OrderID")="0"
			   RSLog("Remark")= "用于兑换" & KS.Setting(45) & ""
			   RSLog("PayTime")=Now
			   RSLog("LogTime")=Now
			   RSLog("Inputer")="System"
			   RSLog("IP")=KS.GetIP
			   RSLog.Update
			   RSLog.Close:set RSLog=Nothing
		    Else
		       If Score=0 Then
			      RS.Close:Set RS=Nothing
				  Response.Write "你输入的积分不正确,积分必须大于0!<br/>"
				  Prev=True
				  Exit Sub
			   End If
			   If KS.ChkClng(Score)<KS.ChkClng(KS.Setting(41)) Then
			      RS.Close:Set RS=Nothing
				  Response.Write "你输入的积分不正确,积分必须大于等于" & KS.Setting(41) &"!<br/>"
				  Prev=True
				  Exit Sub
			   End If
			   IF KS.ChkClng(RS("Score"))<KS.ChkClng(Score) Then
			      RS.Close:Set RS=Nothing
				  Response.Write "你可用积分不足，不能兑换!<br/>"
				  Prev=True
				  Exit Sub
			   End If
			   RS("Score")=RS("Score")-Score
			   'RS("Point")=RS("Point")+Score/KS.Setting(41)
			   RS.Update
			   'ChannelID,InfoID,UserName,InOrOutFlag,Point,User,Descript
			   Call KS.PointInOrOut(0,0,RS("UserName"),1,Score/KS.Setting(41),"System","积分兑换所得")
			End IF
			Response.Write "恭喜您，兑换" & KS.Setting(45) & "成功!<br/>"
			RS.Close:Set RS=Nothing
 	    End Sub
				
		Sub ShowMoney()
		    %>
            【兑换资金】<br/>
            用 户 名:<%=KSUser.UserName%><br/>
            资金余额:<%=KSUser.Money%>元<br/>
            可用积分:<%=KSUser.Score%>分<br/>
            可用<%=KS.Setting(45)%>:<%=formatnumber(KSUser.Point,2)%><%=KS.Setting(46)%><br/>
            <br/>
            将<%=KS.Setting(45)%>兑换成账户资金,兑换比率:1<%=KS.Setting(46)%>:<%=KS.Setting(43)%>元<br/>
            单位(<%=KS.Setting(46)%>):<input maxLength="8" size="6" value="<%=KS.ChkClng(KSUser.Point)%>" name="Point<%=Minute(Now)%><%=Second(now)%>" emptyok="false" format="*N"/>
            <%
            Response.Write "<anchor>执行兑换<go href=""User_Exchange.asp?" & KS.WapValue & """ method=""post"">" &vbcrlf
            Response.Write "<postfield name=""Action"" value=""SaveMoney""/>" &vbcrlf
            Response.Write "<postfield name=""Prepoint"" value=""" & KSUser.Point & """/>" &vbcrlf
            Response.Write "<postfield name=""Point"" value=""$(Point" & Minute(Now) & Second(now) & ")""/>" &vbcrlf
            Response.Write "</go></anchor><br/>" &vbcrlf
			Response.Write "<br/>" &vbcrlf
			Response.Write "一旦兑换成功即不可逆!!!<br/>" &vbcrlf
            Response.Write "说明：您可以将投稿获得的" & KS.Setting(45) & "兑换成账户资金余额，用于在本站商城进行消费。" &vbcrlf
	   End Sub
	   
	   Sub SaveMoney()
		   Dim Point:Point=KS.S("Point")
		   If Round(Point)<=0 Then
			  Response.Write "你输入的" & KS.Setting(45) & "不正确,必须大于0!<br/>"
			  Prev=True
			  Exit Sub
		   End If
		   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		   RS.Open "Select * From KS_User Where UserName='" & KSUser.UserName & "'",Conn,1,3
		   If RS.EOF Then
		      RS.Close:Set RS=Nothing
			  Response.Write "出错啦!<br/>"
			  Prev=True
			  Exit Sub
		   End If
		   IF Round(RS("Point"))<Round(Point) Then
			   RS.Close:Set RS=Nothing
			   Response.Write "你可用" & KS.Setting(45) & "不足!<br/>"
			   Prev=True
			   Exit Sub
		   End If
		   RS("Money")=RS("Money")+(Point*KS.Setting(43))
		   RS.Update
		   'ChannelID,InfoID,UserName,InOrOutFlag,Point,User,Descript
		   Call KS.PointInOrOut(0,0,RS("UserName"),2,Round(Point),"System","用户兑换账户资金")
		   Dim RSLog:Set RSLog=Server.CreateObject("ADODB.RECORDSET")
		   RSLog.Open "Select * From KS_LogMoney",Conn,1,3
		   RSLog.AddNew
		   RSLog("UserName")=rs("UserName")
		   RSLog("ClientName")=rs("RealName")
		   RSLog("CurrMoney")=rs("money")
		   RSLog("Money")=(point*KS.Setting(43))
		   RSLog("MoneyType")=4
		   RSLog("IncomeOrPayOut")=1 
		   RSLog("OrderID")="0"
		   RSLog("Remark")= "兑换资金所得"
		   RSLog("PayTime")=Now
		   RSLog("LogTime")=Now
		   RSLog("Inputer")="System"
		   RSLog("IP")=KS.GetIP
		   RSLog.Update
		   RSLog.Close:set RSLog=Nothing
		   Response.Write "恭喜您，账户资金兑换成功!<br/>"
		   RS.Close:Set RS=Nothing
	   End Sub
	   
	   Sub TypeEdays()
		    Response.Write "【兑换有效期】<br/>" &vbcrlf
			Response.Write "用 户 名:" & KSUser.UserName & "<br/>" &vbcrlf
			Response.Write "资金余额:" & KSUser.Money & "元<br/>" &vbcrlf
            Response.Write "可用积分:" & KSUser.Score & "分<br/>" &vbcrlf
            Response.Write "剩余天数:" &vbcrlf
            If KSUser.ChargeType=3 Then
               Response.Write "无限期<br/>" &vbcrlf
            Else
			   Response.Write KSUser.GetEdays & "天<br/>" &vbcrlf
			   Response.Write "<br/>" &vbcrlf
               Response.Write "使用<a href=""User_ExChange.asp?Action=ShowEdays&amp;ChangeType=1&amp;" & KS.WapValue & """>资金余额</a>兑换成有效天数,兑换比率:" & KS.Setting(44) & "元:1天<br/><br/>" &vbcrlf
               Response.Write "使用<a href=""User_ExChange.asp?Action=ShowEdays&amp;ChangeType=2&amp;" & KS.WapValue & """>经验积分</a>兑换成有效天数,兑换比率:" & KS.Setting(42) & "分:1天<br/>" &vbcrlf
            End if
	   End Sub
	   
	   Sub ShowEdays()
		    Dim ChangeType:ChangeType=KS.S("ChangeType")
			Response.Write "【兑换有效期】<br/>" &vbcrlf
			Response.Write "用 户 名:" & KSUser.UserName & "<br/>" &vbcrlf
			If ChangeType=1 Then
			   Response.Write "资金余额:" & KSUser.Money & "元<br/>" &vbcrlf
            Else
               Response.Write "可用积分:" & KSUser.Score & "分<br/>" &vbcrlf
            End If
            Response.Write "剩余天数:" &vbcrlf
            If KSUser.ChargeType=3 Then
               Response.Write "无限期<br/>" &vbcrlf
			Else
			   Response.Write KSUser.GetEdays & "天<br/>" &vbcrlf
			   Response.Write "<br/>" &vbcrlf
			   If ChangeType=1 Then
			      If Round(KSUser.Money)<KS.Setting(44) Then
                     Response.Write "使用资金余额兑换成有效天数,兑换比率:" & KS.Setting(44) & "元:1天<br/>"&vbcrlf
                     Response.Write "单位(元):<input maxLength=""8"" size=""6"" value=""100"" name=""Money" & Minute(Now) & Second(Now) & """ emptyok=""false"" format=""*N""/>" &vbcrlf
                  Else
                     Response.Write "您目前资金余额不足，请充值后再来兑换！<br/>" &vbcrlf
                  End If
               Else
                  If Round(KSUser.Score)<KS.Setting(42) Then
                     Response.Write "使用经验积分兑换成有效天数,兑换比率:" & KS.Setting(42) & "分:1天<br/>" &vbcrlf
                     Response.Write "单位(分):<input maxLength=""8"" size=""6"" value=""100"" name=""Score" & Minute(Now) & Second(Now) & """ emptyok=""false"" format=""*N""/>" &vbcrlf
                  Else
                     Response.Write "您目前可用积分不足，不能兑换！<br/>" &vbcrlf
                  End If
               End If
               Response.Write "<anchor>执行兑换<go href=""User_Exchange.asp?" & KS.WapValue & """ method=""post"">" &vbcrlf
               Response.Write "<postfield name=""Action"" value=""SaveEdays""/>" &vbcrlf
               Response.Write "<postfield name=""ChangeType"" value=""" & ChangeType & """/>" &vbcrlf
               Response.Write "<postfield name=""Money"" value=""$(Money" & Minute(Now) & Second(Now) &")""/>" &vbcrlf
               Response.Write "<postfield name=""Score"" value=""$(Score" & Minute(Now) & Second(Now) & ")""/>" &vbcrlf
               Response.Write "</go></anchor>" &vbcrlf
               Response.Write "<br/>" &vbcrlf
			   Response.Write "<br/>" &vbcrlf
               Response.Write "一旦兑换成功即不可逆!!!<br/>" &vbcrlf
            End If
	    End Sub
		
	    Sub SaveEdays()
	   	    Dim ChangeType:ChangeType=KS.S("ChangeType")
			Dim Money:Money=KS.S("Money")
			Dim Score:Score=KS.ChkClng(KS.S("Score"))
			Dim tmpDays,ValidDays,RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select * From KS_User Where UserName='" & KSUser.UserName & "'",conn,1,3
			If RS.EOF Then
			   RS.Close:Set RS=Nothing
			   Response.Write "出错啦!<br/>"
			   Prev=True
			   Exit Sub
			End If
			If ChangeType=1 Then
			   If KS.ChkClng(Money)=0 Then
			      RS.Close:Set RS=Nothing
				  Response.Write "你输入的资金不正确,资金必须大于0!<br/>"
				  Prev=True
				  Exit Sub
			   End If
			   If KS.ChkClng(Money)<KS.ChkClng(KS.Setting(44)) Then
			      RS.Close:Set RS=Nothing
				  Response.Write "你输入的资金不正确,资金必须大于等于" & KS.Setting(44) &"!<br/>"
				  Prev=True
				  Exit Sub
			   End If
			   IF Round(RS("Money"))<Round(Money) Then
			      RS.Close:Set RS=Nothing
				  Response.Write "你可用资金不足，请充值后再来兑换!<br/>"
				  Prev=True
				  Exit Sub
		       End If
			   RS("Money")=RS("Money")-Money
			   ValidDays=RS("Edays")
			   TmpDays=ValidDays-DateDiff("D",RS("BeginDate"),now())
			   If TmpDays>0 Then
			      RS("Edays")=RS("Edays")+Money/KS.Setting(44)
			   Else
			      RS("BeginDate")=now
				  RS("Edays")=Money/KS.Setting(44)
			   End If
			   RS.Update
			   'UserName,InOrOutFlag,Edays,User,Descript
			   Call KS.EdaysInOrOut(RS("UserName"),1,Money/KS.Setting(44),"System","账户资金兑换所得")
			   Dim RSLog:Set RSLog=Server.CreateObject("ADODB.RECORDSET")
			   RSLog.Open "Select * From KS_LogMoney",Conn,1,3
			   RSLog.AddNew
			   RSLog("UserName")=rs("UserName")
			   RSLog("ClientName")=rs("RealName")
			   RSLog("Money")=Money
			   RSlog("CurrMoney")=RS("Money")
			   RSLog("MoneyType")=4
			   RSLog("IncomeOrPayOut")=2 
			   RSLog("OrderID")="0"
			   RSLog("Remark")= "用于兑换有效天数"
			   RSLog("PayTime")=Now
			   RSLog("LogTime")=Now
			   RSLog("Inputer")="System"
			   RSLog("IP")=KS.GetIP
			   RSLog.Update
			   RSLog.Close:set RSLog=Nothing
		    Else
		       If Score=0 Then
			      RS.Close:Set RS=Nothing
				  Response.Write "<script>alert('你输入的积分不正确,积分必须大于0!');history.back();</script>"
				  Prev=True
				  Exit Sub
			   End If
			   If KS.ChkClng(Score)<KS.ChkClng(KS.Setting(42)) Then
			      RS.Close:Set RS=Nothing
				  Response.Write "你输入的积分不正确,积分必须大于等于" & KS.Setting(42) &"!<br/>"
				  Prev=True
				  Exit Sub
			   End If
			   IF KS.ChkClng(RS("Score"))<KS.ChkClng(Score) Then
			      RS.Close:Set RS=Nothing
				  Response.Write "你可用积分不足，不能兑换!<br/>"
				  Prev=True
				  Exit Sub
		       End If
			   RS("Score")=RS("Score")-Score
			   ValidDays=RS("Edays")
			   TmpDays=ValidDays-DateDiff("D",RS("BeginDate"),now())
			   If TmpDays>0 Then
			      RS("Edays")=RS("Edays")+Score/KS.Setting(42)
			   Else
			      RS("BeginDate")=now
			      RS("Edays")=Score/KS.Setting(42)
			   End If
			   RS.Update
			   'UserName,InOrOutFlag,Edays,User,Descript
			   Call KS.EdaysInOrOut(RS("UserName"),1,Score/KS.Setting(42),"System","积分兑换所得")
			End IF
			Response.Write "恭喜您，有效天数兑换成功!<br/>"
			RS.Close:Set RS=Nothing
	    End Sub
End Class
%> 
