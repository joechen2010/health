<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New User_ExChange
KSCls.Kesion()
Set KSCls = Nothing

Class User_ExChange
        Private KS,KSUser
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "<script>top.location.href='Login';</script>"
		  Exit Sub
		End If
		Call KSUser.Head()
		Response.Write "<div class=""tabs"">"
		Response.Write " <ul class="""">"
		Response.Write " <li><a href=""User_PayOnline.asp"">在线支付充值</a></li>"
		Response.Write " <li><a href=""user_recharge.asp"">充值卡充值</a></li>"
		If KS.S("Action")="Point" Then
		Response.Write " <li class='select'><a href=""user_exchange.asp?Action=Point"">兑换" & KS.Setting(45) & "</a></li>"
		Else
		Response.Write " <li><a href=""user_exchange.asp?Action=Point"">兑换" & KS.Setting(45) & "</a></li>"
		End IF
		If KS.S("Action")="Edays" Then
		Response.Write " <li class='select'><a href=""user_exchange.asp?Action=Edays"">兑换有效期</a></li>"
		Else
		Response.Write " <li><a href=""user_exchange.asp?Action=Edays"">兑换有效期</a></li>"
		End If
		If KS.S("Action")="Money" Then
		Response.Write " <li class='select'><a href=""user_exchange.asp?Action=Money"">" & KS.Setting(45) & "兑换账户资金</a></li>"
		Else
		Response.Write " <li><a href=""user_exchange.asp?Action=Money"">" & KS.Setting(45) & "兑换账户资金</a></li>"
		End If
		
		Response.Write "</ul>"
		Response.Write "</div>"
		Select Case KS.S("Action")
		 Case "Point"
		   Call KSUser.InnerLocation("兑换" & KS.Setting(45))
		   Call ExchangePoint()
		 Case "Money" 
		   Call KSUser.InnerLocation("兑换账户资金")
		   Call ExchangeMoney()
		 Case "SaveExchangeMoney"
		   Call SaveExchangeMoney()
		 Case "SaveExchangePoint"
		   Call SaveExchangePoint()
		 Case "Edays"
		 	Call KSUser.InnerLocation("兑换有效天数")
		    Call ExchangeEdays()
		 Case "SaveExchangeEdays"
		    Call SaveExchangeEdays()
		End Select
       End Sub
	   
	   Sub ExchangePoint()
	   %>
	   <script>
	     function Confirm()
		 {
		   var str='友情提醒:\n';
		   if (document.myform.ChangeType[0].checked==true){
		     if (parseInt(document.myform.Premoney.value)<parseInt(document.myform.Money.value)){
			   alert('您目前资金余额不足，请充值后再来兑换！');
			   return false;
			   }
		    str+='兑换前资金：'+document.myform.Premoney.value+' 元\n';
			str+='兑换后资金：'+(document.myform.Premoney.value-document.myform.Money.value)+' 元\n';
			str+='一旦兑换成功即不可逆，确定兑换吗？';
		   }else{
		   if (parseInt(document.myform.PreScore.value)<parseInt(document.myform.Score.value)){
			   alert('您目前可用积分不足，不能兑换！');
			   return false;
			   }
		    str+='兑换前积分：'+document.myform.PreScore.value+' 分\n';
			str+='兑换后积分：'+(document.myform.PreScore.value-document.myform.Score.value)+' 分\n';
			str+='一旦兑换成功即不可逆，确定兑换吗？';
		   }
		   if (confirm(str)){
		    return true}
		   else{ return false}
		 }
	   </script>
		<FORM name=myform action="User_Exchange.asp" method="post">
		  <table class=border cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
			<tr class=title>
			  <td align=middle colSpan=2 height=22><B> 兑 换 点 券 </B></td>
			</tr> 
			<tr class=tdbg>
			  <td align=right width=120>用户名：</td>
			  <td><%=KSUser.UserName%></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>资金余额：</td>
			  <td><input type='hidden' value='<%=KSUser.Money%>' name='Premoney'><%=KSUser.Money%> 元</td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>可用积分：</td>
			  <td><input type='hidden' value='<%=KSUser.Score%>' name='PreScore'><%=KSUser.Score%> 分</td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>可用<%=KS.Setting(45)%>：</td>
			  <td><%=KSUser.Point%>&nbsp;<%=KS.Setting(46)%></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>兑换点券：</td>
			  <td>
		  <Input type=radio CHECKED value="1" name="ChangeType">使用资金余额： 将 
		  <Input style="TEXT-ALIGN: center" maxLength=8 size=6 value=100 name="Money"> 元兑换成<%=KS.Setting(45)%> &nbsp;&nbsp;&nbsp;&nbsp;<Font color=red>兑换比率：<%=KS.Setting(43)%>元:1<%=KS.Setting(46)%></Font> <br>
		  <Input type=radio value="2" name="ChangeType">使用经验积分： 将 
				<Input style="TEXT-ALIGN: center" maxLength=8 size=6 value=100 name="Score"> 分兑换成<%=KS.Setting(45)%> &nbsp;&nbsp;&nbsp;&nbsp;<Font color=red>兑换比率：<%=KS.Setting(41)%>分:1<%=KS.Setting(46)%> </Font></td>
			</tr>
			<tr class=tdbg>
			  <td align=middle colSpan=2 height=40>
		        <Input id=Action type=hidden value="SaveExchangePoint" name="Action"> 
				<Input class="button" id=Submit type=submit value="执行兑换" onClick="return(Confirm())" name=Submit> </td>
			</tr>
		  </table>
		</FORM>
	   <%
	   End Sub
	   
	   Sub SaveExchangePoint()
	     Dim ChangeType:ChangeType=KS.S("ChangeType")
		 Dim Money:Money=KS.S("Money")
		 Dim Score:Score=KS.ChkClng(KS.S("Score"))
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select * From KS_User Where UserName='" & KSUser.UserName & "'",conn,1,1
		 If RS.Eof  Then
		   Rs.Close:Set RS=Nothing
		   Response.Write "<script>alert('出错啦!');history.back();</script>"
		   Exit Sub
		 End If
		 If ChangeType=1 Then
		    If KS.ChkClng(Money)=0 Then
			   Rs.Close:Set RS=Nothing
			   Response.Write "<script>alert('你输入的资金不正确,资金必须大于0!');history.back();</script>"
			   Exit Sub
			End iF
			If KS.ChkClng(Money)<KS.ChkClng(KS.Setting(43)) Then
			   Rs.Close:Set RS=Nothing
			   Response.Write "<script>alert('你输入的资金不正确,资金必须大于等于" & KS.Setting(43) &"!');history.back();</script>"
			   Exit Sub
			End If
		   IF Round(RS("Money"))<Round(Money) Then
			   Rs.Close:Set RS=Nothing
			   Response.Write "<script>alert('你可用资金不足，请充值后再来兑换!');history.back();</script>"
			   Exit Sub
		   End If

			'  ChannelID,InfoID,UserName,InOrOutFlag,Point,User,Descript
			Call KS.PointInOrOut(0,0,rs("UserName"),1,Money/KS.Setting(43),"System","账户资金兑换所得",0)
			Call KS.MoneyInOrOut(rs("UserName"),rs("RealName"),Money,4,2,now,0,"System","用于兑换点券",0,0)
	
		 Else
		    If Score=0 Then
			   Rs.Close:Set RS=Nothing
			   Response.Write "<script>alert('你输入的积分不正确,积分必须大于0!');history.back();</script>"
			   Exit Sub
			End If
			If KS.ChkClng(Score)<KS.ChkClng(KS.Setting(41)) Then
			   Rs.Close:Set RS=Nothing
			   Response.Write "<script>alert('你输入的积分不正确,积分必须大于等于" & KS.Setting(41) &"!');history.back();</script>"
			   Exit Sub
			End If
		   IF KS.ChkClng(RS("Score"))<KS.ChkClng(Score) Then
			   Rs.Close:Set RS=Nothing
			   Response.Write "<script>alert('你可用积分不足，不能兑换!');history.back();</script>"
			   Exit Sub
		   End If
		   
		     call KS.ScoreInOrOut(rs("UserName"),2,Score,"System","兑换点券消耗!",0,0)
			'ChannelID,InfoID,UserName,InOrOutFlag,Point,User,Descript
			 Call KS.PointInOrOut(0,0,rs("UserName"),1,Score/KS.Setting(41),"System","积分兑换所得",0)
		 End IF
		 Response.Write "<script>alert('恭喜您，点券兑换成功!');location.href='User_ExChange.asp?Action=Point';</script>"
		 RS.Close:Set RS=Nothing
 	   End Sub
	   
	   
	   Sub ExchangeMoney()
	   %>
	   		<script>
	     function checkform()
		 {
		   var str='友情提醒:\n';
		     if (parseInt(document.myforms.Prepoint.value)<parseInt(document.myforms.Point.value)){
			   alert('对不起，你的<%=KS.Setting(45)%>不足！');
			   return false;
			   }
		    str+='兑换前<%=KS.Setting(45)%>：'+document.myforms.Prepoint.value+' <%=KS.Setting(46)%>\n';
			str+='兑换后<%=KS.Setting(45)%>：'+(document.myforms.Prepoint.value-document.myforms.Point.value)+' <%=KS.Setting(46)%>\n';
			str+='一旦兑换成功即不可逆，确定兑换吗？';

		   if (confirm(str)){
		    return true}
		   else{ return false}
		 }
	   </script>
		<FORM name=myforms action="User_Exchange.asp" method="post">
		  <table class=border cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
			<tr class=title>
			  <td align=middle colSpan=2 height=22><B> 兑 换 资 金 </B></td>
			</tr> 
			<tr class=tdbg>
			  <td align=right width=120>用户名：</td>
			  <td><%=KSUser.UserName%></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>资金余额：</td>
			  <td><%=KSUser.Money%> 元</td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>可用积分：</td>
			  <td><%=KSUser.Score%> 分</td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>可用<%=KS.Setting(45)%>：</td>
			  <td><%=formatnumber(KSUser.Point,2)%>&nbsp;<%=KS.Setting(46)%><input type='hidden' value='<%=KSUser.Point%>' name='Prepoint'></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>兑换资金：</td>
			  <td>
		   将
		  <Input style="TEXT-ALIGN: center" maxLength=8 size=6 value=<%=KS.ChkClng(KSUser.Point)%> name="Point"> <%=KS.Setting(46)%><%=KS.Setting(45)%>兑换成账户资金 &nbsp;&nbsp;&nbsp;&nbsp;<Font color=red>兑换比率：1<%=KS.Setting(46)%>:<%=KS.Setting(43)%>元</Font> <br>
		  </td>
			</tr>
			<tr class=tdbg>
			  <td align=middle colSpan=2 height=40>
		        <Input id=Action type=hidden value="SaveExchangeMoney" name="Action"> 
				<Input class="button" id=Submit type=submit value="执行兑换" onClick="return(checkform())" name=Submit> </td>
			</tr>
		  </table>
		</FORM>
		<div style="padding-left:60px;color:green">说明：您可以将投稿获得的<%=KS.Setting(45)%>兑换成账户资金余额，用于在本站商城进行消费。</div>
		
	   <%
	   End Sub
	   
	   Sub SaveExchangeMoney()
		 Dim Point:Point=KS.S("Point")

		    If Round(Point)<=0 Then
			   Response.Write "<script>alert('你输入的" & KS.Setting(45) & "不正确,必须大于0!');history.back();</script>"
			   Exit Sub
			End iF
			
			
		DIM RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_User Where UserName='" & KSUser.UserName & "'",conn,1,1
		 If RS.Eof  Then
		   Rs.Close:Set RS=Nothing
		   Response.Write "<script>alert('出错啦!');history.back();</script>"
		   Exit Sub
		 End If
		   IF Round(RS("Point"))<Round(Point) Then
			   Rs.Close:Set RS=Nothing
			   Response.Write "<script>alert('你可用" & KS.Setting(45) & "不足!');history.back();</script>"
			   Exit Sub
		   End If
			'  ChannelID,InfoID,UserName,InOrOutFlag,Point,User,Descript
			 Call KS.PointInOrOut(0,0,rs("UserName"),2,Round(Point),"System","用户兑换账户资金",0)
			 Call KS.MoneyInOrOut(rs("UserName"),rs("RealName"),(point*KS.Setting(43)),4,1,now,0,"System","点券兑换所得",0,0)
		 Response.Write "<script>alert('恭喜您，账户资金兑换成功!');location.href='User_ExChange.asp?Action=Money';</script>"
		 RS.Close:Set RS=Nothing
	   End Sub
	   
	   	   
	   Sub ExchangeEdays()
	    %>
	   <script>
	     function Confirm()
		 {
		   var str='友情提醒:\n';
		   if (document.myform.ChangeType[0].checked==true){
		     if (parseInt(document.myform.Premoney.value)<parseInt(document.myform.Money.value)){
			   alert('您目前资金余额不足，请充值后再来兑换！');
			   return false;
			   }
		    str+='兑换前资金：'+document.myform.Premoney.value+' 元\n';
			str+='兑换后资金：'+(document.myform.Premoney.value-document.myform.Money.value)+' 元\n';
			str+='一旦兑换成功即不可逆，确定兑换吗？';
		   }else{
		   if (parseInt(document.myform.PreScore.value)<parseInt(document.myform.Score.value)){
			   alert('您目前可用积分不足，不能兑换！');
			   return false;
			   }
		    str+='兑换前积分：'+document.myform.PreScore.value+' 分\n';
			str+='兑换后积分：'+(document.myform.PreScore.value-document.myform.Score.value)+' 分\n';
			str+='一旦兑换成功即不可逆，确定兑换吗？';
		   }
		   if (confirm(str)){
		    return true}
		   else{ return false}
		 }
	   </script>
		<FORM name=myform action="User_Exchange.asp" method="post">
		  <table class=border cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
			<tr class=title>
			  <td align=middle colSpan=2 height=22><B> 兑 换 有 效 期</B></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>用户名：</td>
			  <td><%=KSUser.UserName%></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>资金余额：</td>
			  <td><input type='hidden' value='<%=KSUser.Money%>' name='Premoney'><%=KSUser.Money%> 元</td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>可用积分：</td>
			  <td><input type='hidden' value='<%=KSUser.Score%>' name='PreScore'><%=KSUser.Score%> 分</td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>剩余天数：</td>
			  <td>
			  <%if KSUser.ChargeType=3 Then%>
			  无限期
			  <%else%>
			  <%=KSUser.GetEdays%>&nbsp;天
			  <%end if%></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>兑换点券：</td>
			  <td>
		  <Input type=radio CHECKED value="1" name="ChangeType">使用资金余额： 将 
		  <Input style="TEXT-ALIGN: center" maxLength=8 size=6 value=100 name="Money"> 元兑换成有效天数 &nbsp;&nbsp;&nbsp;&nbsp;<Font color=red>兑换比率：<%=KS.Setting(44)%>元:1天</Font> <br>
		  <Input type=radio value="2" name="ChangeType">使用经验积分： 将 
				<Input style="TEXT-ALIGN: center" maxLength=8 size=6 value=100 name="Score"> 分兑换成有效天数 &nbsp;&nbsp;&nbsp;&nbsp;<Font color=red>兑换比率：<%=KS.Setting(42)%>分:1天 </Font></td>
			</tr>
			<tr class=tdbg>
			  <td align=middle colSpan=2 height=40>
		        <Input id=Action type=hidden value="SaveExchangeEdays" name="Action"> 
				<Input class="button" id=Submit type=submit value="执行兑换" onClick="return(Confirm())" name=Submit> </td>
			</tr>
		  </table>
		</FORM>
	   <%
	   End Sub
		
	   Sub SaveExchangeEdays()
	   	 Dim ChangeType:ChangeType=KS.S("ChangeType")
		 Dim Money:Money=KS.S("Money")
		 Dim Score:Score=KS.ChkClng(KS.S("Score"))
		 Dim tmpDays,ValidDays,RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select * From KS_User Where UserName='" & KSUser.UserName & "'",conn,1,3
		 If RS.Eof  Then
		   Rs.Close:Set RS=Nothing
		   Response.Write "<script>alert('出错啦!');history.back();</script>"
		   Exit Sub
		 End If
		 If ChangeType=1 Then
		    If KS.ChkClng(Money)=0 Then
			   Rs.Close:Set RS=Nothing
			   Response.Write "<script>alert('你输入的资金不正确,资金必须大于0!');history.back();</script>"
			   Exit Sub
			End iF
			If KS.ChkClng(Money)<KS.ChkClng(KS.Setting(44)) Then
			   Rs.Close:Set RS=Nothing
			   Response.Write "<script>alert('你输入的资金不正确,资金必须大于等于" & KS.Setting(44) &"!');history.back();</script>"
			   Exit Sub
			End If
		   IF Round(RS("Money"))<Round(Money) Then
			   Rs.Close:Set RS=Nothing
			   Response.Write "<script>alert('你可用资金不足，请充值后再来兑换!');history.back();</script>"
			   Exit Sub
		   End If
			    ValidDays=rs("Edays")
				tmpDays=ValidDays-DateDiff("D",rs("BeginDate"),now())
				if tmpDays>0 then
					rs("Edays")=rs("Edays")+Money/KS.Setting(44)
				else
					rs("BeginDate")=now
					rs("Edays")=Money/KS.Setting(44)
				end if
			RS.Update
			'  UserName,InOrOutFlag,Edays,User,Descript
			Call KS.EdaysInOrOut(rs("UserName"),1,Money/KS.Setting(44),"System","账户资金兑换所得")
			Call KS.MoneyInOrOut(rs("UserName"),rs("RealName"),Money,4,2,now,0,"System","用于兑换有效天数",0,0)
			

		 Else
		    If Score=0 Then
			   Rs.Close:Set RS=Nothing
			   Response.Write "<script>alert('你输入的积分不正确,积分必须大于0!');history.back();</script>"
			   Exit Sub
			End If
			If KS.ChkClng(Score)<KS.ChkClng(KS.Setting(42)) Then
			   Rs.Close:Set RS=Nothing
			   Response.Write "<script>alert('你输入的积分不正确,积分必须大于等于" & KS.Setting(42) &"!');history.back();</script>"
			   Exit Sub
			End If
		   IF KS.ChkClng(RS("Score"))<KS.ChkClng(Score) Then
			   Rs.Close:Set RS=Nothing
			   Response.Write "<script>alert('你可用积分不足，不能兑换!');history.back();</script>"
			   Exit Sub
		   End If
		    RS("Score")=RS("Score")-Score
			   ValidDays=rs("Edays")
				tmpDays=ValidDays-DateDiff("D",rs("BeginDate"),now())
				if tmpDays>0 then
					rs("Edays")=rs("Edays")+Score/KS.Setting(42)
				else
					rs("BeginDate")=now
					rs("Edays")=Score/KS.Setting(42)
				end if
			RS.Update
			'  UserName,InOrOutFlag,Edays,User,Descript
			Call KS.EdaysInOrOut(rs("UserName"),1,Score/KS.Setting(42),"System","积分兑换所得")
		 End IF
		 Response.Write "<script>alert('恭喜您，有效天数兑换成功!');location.href='User_ExChange.asp?Action=Edays';</script>"
		 RS.Close:Set RS=Nothing
	   End Sub
End Class
%> 
