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
Set KSCls = New User_Recharge
KSCls.Kesion()
Set KSCls = Nothing

Class User_Recharge
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
		Call KSUser.InnerLocation("充值卡充值")
		Response.Write "<div class=""tabs"">"
		Response.Write " <ul>"
		Response.Write " <li><a href=""User_PayOnline.asp"">在线支付充值</a></li>"
		Response.Write " <li class='select'><a href=""user_recharge.asp"">充值卡充值</a></li>"
		Response.Write " <li><a href=""user_exchange.asp?Action=Point"">兑换" & KS.Setting(45) & "</a></li>"
		Response.Write " <li><a href=""user_exchange.asp?Action=Edays"">兑换有效期</a></li>"
		Response.Write " <li><a href=""user_exchange.asp?Action=Money"">" & KS.Setting(45) & "兑换账户资金</a></li>"
		Response.Write "</ul>"
		Response.Write "</div>"
		Select Case KS.S("Action")
		 Case "SaveExchangeEdays"
		    Call SaveExchangeEdays()
	     Case Else
		    Call ExchangeEdays()
		End Select
       End Sub
	  
	   
	   Sub ExchangeEdays()
	    %>
	   <script>
	     function Confirm()
		 {
		  if (document.myform.CardNum.value=="")
		  {
		   alert('请输入充值卡卡号!')
		   document.myform.CardNum.focus();
		   return false;
		  }
		  if (document.myform.CardPass.value=="")
		  {
		   alert('请输入充值卡密码!')
		   document.myform.CardPass.focus();
		   return false;
		  }
		  return true;
		  }
	   </script>
		<FORM name=myform action="User_ReCharge.asp" method="post">
		  <table class=border cellSpacing=1 cellPadding=2 width="100%" align=center border=0>
			<tr class=title>
			  <td align=middle colSpan=2 height=22><B> 充 值 卡 充 值</B></td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>用户名：</td>
			  <td><%=KSUser.UserName%></td>
			</tr>
			<tr class=tdbg>
			  <td align=right>计费方式：</td>
			  <td><%if KSUser.ChargeType=1 Then 
		  Response.Write "扣点数</font>计费用户"
		  ElseIf KSUser.ChargeType=2 Then
		   Response.Write "有效期</font>计费用户,到期时间：" & cdate(KSUser.BeginDate)+KSUser.Edays & ","
		  ElseIf KSUser.ChargeType=3 Then
		   Response.Write "无限期</font>计费用户"
		  End If
		  %>&nbsp;</td>
		    </tr>
			<tr class=tdbg>
			  <td align=right width=120>资金余额：</td>
			  <td><input type='hidden' value='<%=KSUser.Money%>' name='Premoney'><%=KSUser.Money%> 元</td>
			</tr>
			<tr class=tdbg>
			  <td align=right width=120>可用<%=KS.Setting(45)%>：</td>
			  <td><%=KSUser.Point%>&nbsp;<%=KS.Setting(46)%></td>
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
			  <td align=right>充值卡卡号：</td>
			  <td>&nbsp;<input name="CardNum" type="text" class="textbox" size="25" maxlength="50"></td>
		    </tr>
			<tr class=tdbg>
			  <td align=right width=120>充值卡密码：</td>
			  <td>&nbsp;<input name="CardPass" type="text" class="textbox" size="25" maxlength="50"></td>
			</tr>
			<tr class=tdbg>
			  <td align=middle colSpan=2 height=40>
		        <Input id=Action type=hidden value="SaveExchangeEdays" name="Action"> 
				<Input class="button" id=Submit type=submit value="确定充值" onClick="return(Confirm())" name=Submit> </td>
			</tr>
		  </table>
		</FORM>
	   <%
	   End Sub
		
	   Sub SaveExchangeEdays()
	   	 Dim ChangeType:ChangeType=KS.S("ChangeType")
		 Dim Money:Money=KS.S("Money")
		 DiM CardNum:CardNum=KS.S("CardNum")
		 Dim CardPass:CardPass=KS.S("CardPass")
		 If CardNum="" Or CardPass="" Then 
		   Call KS.AlertHistory("请输入的充值卡号及密码！",-1)
		   exit sub
		 end if
		 Dim RS:Set RS=Server.CreateObject("adodb.recordset")
		 rs.open "select top 1 * from ks_usercard where cardtype=0 and cardnum='" & CardNum & "'",conn,1,1
		 if rs.bof and rs.eof then
		  rs.close:set rs=nothing
		  Call KS.AlertHistory("对不起，您输入的充值卡号不正确！",-1)
		  exit sub
		 end if
		 if rs("cardpass")<>KS.Encrypt(cardpass) then
		  rs.close:set rs=nothing
		  Call KS.AlertHistory("对不起，您输入的充值卡密码不正确！",-1)
		  exit sub
		 end if
		 
		 if rs("isused")=1 then
		  rs.close:set rs=nothing
		  Call KS.AlertHistory("对不起，您输入的充值卡已被使用！",-1)
		  exit sub
		 end if
		 
		 if datediff("d",rs("enddate"),now())>0 then
		  rs.close:set rs=nothing
		  Call KS.AlertHistory("对不起，您输入的充值卡已过期！",-1)
		  exit sub
		 end if
		 
		 if not KS.IsNul(rs("allowgroupid")) then
		    If KS.FoundInArr(rs("allowGroupID"),KSUser.GroupID,",")=false Then
			  rs.close:set rs=nothing
			  Call KS.AlertHistory("对不起，您所在的用户组没有使用本充值卡的权限,请联系本站管理员！",-1)
			  exit sub
			End If
		 end if
		 
		  Dim ValidNum:ValidNum=rs("ValidNum")
		  Dim ValidUnit:ValidUnit=rs("ValidUnit")
		  Dim UserCardID:UserCardID=rs("id")
		  Dim GroupID:GroupID=rs("GroupID")
		  rs.close
		  rs.open "select top 1 * from ks_user Where UserName='" & KSUser.UserName & "'",conn,1,1
		  if not rs.eof then
		    if rs("ChargeType")=3 and ValidUnit<>3 then
				  rs.close:set rs=nothing
				  Call KS.AlertHistory("由于你的账户永不过期，如需充值资金，请购买资金卡！",-1)
				  exit sub
			end if
			dim ValidDays,tmpdays
		    select case ValidUnit
			  case 1 '点数
			   'rs("point")=rs("point")+ValidNum
			   Call KS.PointInOrOut(0,0,rs("UserName"),1,ValidNum,"System","通过充值卡获得的点数",0)
			  case 2 '天数
			    ValidDays=rs("Edays")
				tmpDays=ValidDays-DateDiff("D",rs("BeginDate"),now())
				if tmpDays>0 then
				    conn.execute("update ks_user set edays=edays+" & validnum & " where username='" & ksuser.username & "'")
				else
					conn.execute("update ks_user set begindate=" & sqlnowstring & ",edays=" & validnum & " where username='" & ksuser.username & "'")
				end if
				Call KS.EdaysInOrOut(rs("UserName"),1,ValidNum,"System","通过充值卡[" & CardNum & "]获得的有效天数")
			  case 3 '金币
			    Call KS.MoneyInOrOut(rs("UserName"),RS("RealName"),ValidNum,4,1,now,0,"System","通过充值卡[" & CardNum & "]获得的资金",0,0)
			  case 4 '积分
			    Call KS.ScoreInOrOut(rs("UserName"),1,ValidNum,"System","通过充值卡[" & CardNum & "]获得的积分!",0,0)
			end select
			if GroupID<>0 then conn.execute("update ks_user set groupid=" & GroupID & " where userName='" & KSUser.UserName & "'")
			conn.execute("update ks_user set usercardid="&usercardid &" where userName='" & KSUser.UserName & "'")
		  end if
		  '置充值卡已使用、已售出
		  Conn.Execute("Update KS_UserCard Set Isused=1,issale=1,username='" & KSUser.UserName & "',UseDate=" & SqlNowString & " where cardnum='" & cardnum & "'")
		 
		 if GroupID<>0 then
		 Response.Write "<script>alert('恭喜您，充值成功并升级为"""& KS.U_G(GroupID,"groupname") &"""!');location.href='user_recharge.asp';</script>"
		 else
		 Response.Write "<script>alert('恭喜您，充值成功!');location.href='user_recharge.asp';</script>"
		 end if
		 RS.Close:Set RS=Nothing
	   End Sub
End Class
%> 
