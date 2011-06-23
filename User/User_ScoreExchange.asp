<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.UpFileCls.asp"-->
<%

Dim KSCls
Set KSCls = New User_Blog
KSCls.Kesion()
Set KSCls = Nothing

Class User_Blog
        Private KS,KSUser
		Private CurrentPage,totalPut,RS,MaxPerPage
		Private ComeUrl,AddDate,Weather
		Private TypeID,Title,Tags,UserName,Face,Content,Status,PicUrl,Action,I,ClassID,password
		Private Sub Class_Initialize()
		  MaxPerPage =15
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		ComeUrl=Request.ServerVariables("HTTP_REFERER")
		Call KSUser.Head()
		Call KSUser.InnerLocation("积分兑换礼品")
		KSUser.CheckPowerAndDie("s15")
		
		%>
		<div class="tabs">	
			<ul>
			  <li<%If request("action")="" or request("action")="showdetail" or request("action")="exchange" or request("action")="exchangesave" Then response.write " class='select'"%>><a href="user_scoreexchange.asp">可供兑换的礼品</a></li>
			  <li<%if request("action")="order" or request("action")="showdetail1" then response.write " class='select'"%>><a href="?action=order">兑换记录查询</a></li>
			</ul>
		</div>
		<%
		
			Select Case KS.S("Action")
			 Case "showdetail"
			   Call showdetail()
			   Call KSUser.InnerLocation("查看礼品详情")
			 Case "exchange"
			   Call exchange()
			   Call KSUser.InnerLocation("确认及填写收货地址")
			 Case "exchangesave"
			   Call exchangesave()
			   Call KSUser.InnerLocation("成功兑换礼品")
			 Case "showdetail1" 
			   Call showdetail1()
			   Call KSUser.InnerLocation("查看礼品详情")
			 Case "order"
			   Call ShowOrder()
			   Call KSUser.InnerLocation("查看兑换订单")
			 Case "setok"
			   Call SetOrderOk()
			 Case "dosave"
			   Call dosave()
			 Case Else
			  Call ShowMain()
			End Select
		 
	   End Sub
	   
	   
	   Sub ShowMain()
		    MaxPerPage=8
			If KS.S("page") <> "" Then
				CurrentPage = KS.ChkClng(KS.S("page"))
			Else
				CurrentPage = 1
			End If
		%>
			    <table width="98%"  border="0" align="center" cellpadding="0" cellspacing="0">
                    
                   <%
						Set RS=Server.CreateObject("AdodB.Recordset")
							RS.open "select * from KS_MallScore where status=1 order by id desc",conn,1,1
								 If RS.EOF And RS.BOF Then
								  Response.Write "<tr><td class='tdbg' align='center' height=30 valign=top>	可供兑换的礼品<td></tr>"
								 Else
									totalPut = RS.RecordCount
						
								   If CurrentPage < 1 Then	CurrentPage = 1
									If CurrentPage >1 and  (CurrentPage - 1) * MaxPerPage < totalPut Then
											RS.Move (CurrentPage - 1) * MaxPerPage
									Else
											CurrentPage = 1
									End If
			   %>
							   <style type="text/css">
								.t .onmouseover { background: #fffff0; }
								.t .onmouseout {}
								.t ul {float:left;margin:6px;padding:5px;width:152px!important;width:165px;height:225px;overflow:hidden;border: 1px #f4f4f4 solid;background: #fcfcfc;}
								.t ul li {
								list-style-type:none;line-height:1.5;margin:0;padding:0;}
								.t ul li.l1 img {width:150px;height:90px;}
								.t ul li.l1 a {display:block;margin:auto;padding:1px;width:156px;height:96px;background:url("images/tbg.png") no-repeat left top;text-align:left;}
								.t ul li.l2 {margin: 3px 0 0 0; width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
								.t ul li.l3 {margin: 3px 0 0 0; width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}
								.t ul li.l4 {margin:10px 0 0 0;text-align:center;}
							   </style>
							   <%
								 dim i,k
								 do while not rs.eof
								   response.write "<tr>"
								   for i=1 to 4
									response.write "<td class=""t"" width=""25%"">"
									 dim pic:pic=rs("photourl")
									 if pic="" or isnull(pic) then pic="../images/nophoto.gif"
									%>
									<ul onMouseOver="this.className='onmouseover'" onMouseOut="this.className='onmouseout'" class="onmouseout">
										<li class="l1"><a href='?action=showdetail&id=<%=rs("id")%>'>
						<img src="<%=pic%>" title="点击查看详情" width="200" height="122" border="0" />
						</a></li>
										<li class="l2">名称：<strong><%=rs("productname")%></strong>
										<%if rs("recommend")=1 then response.write "<font color=red>荐</font>"%>
										</li>
										<li class="l3"><font color=#ff6600>积分：<%=rs("score")%> 分</font></li>
										<li class="l2">数量：<%=rs("Quantity")%></li>
										<li class="l2">截止时间：<%=formatdatetime(rs("enddate"),2)%></li>
										
										<li class="l4">
										<input type="submit" value=" 查看 " class="button" onClick="window.location='?action=showdetail&id=<%=RS("ID")%>'" />
										<input type="submit" value=" 兑换 " class="button" onClick="window.location='?action=exchange&id=<%=RS("ID")%>'" />
										</li>									
									</ul>
									<%
									response.write "</td>"
									rs.movenext
									k=k+1
									if rs.eof or k>=MaxPerPage then exit for 
								   next
								   for i=k+1 to 4
									response.write "<td width=""25%"">&nbsp;</td>"
								   next
								  response.write "</tr>"
								  if rs.eof or k>=MaxPerPage then exit do
								 loop
								 response.write "<tr>"
								 response.write "<td colspan=4 align=""right"">"
								 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)
								 Response.write "</td>"
								 response.write "</tr>"
										End If
     %>                     
				</table>
				<br />
				<br />
				<div style="margin:15px;line-height:20px">
				 <strong>温馨提示:</strong>
				 <br />1、会员兑换礼品，需要有足够的积分才可以兑换
				 <br />2、只要积分足够,可以兑换多件礼品，兑换成功以后，系统会发出系统消息到会员的消息中心，显示兑换成功与否!
				 <br />
				  3、兑换礼品后，我们提供快递到付或者自取两种方式
				</div>

		<%
	   End Sub
	   
	   Sub ShowDetail()
	    Dim ID:ID=KS.ChkClng(KS.S("ID"))
	    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select * From KS_MallScore Where ID=" & ID & " And Status=1",conn,1,3
		If RS.Eof And RS.Bof Then
		  Rs.Close
		  Call KS.AlertHistory("对不起,参数出错!",-1)
		  Exit Sub
		Else
		  RS("Hits")=RS("Hits")+1
		  RS.Update
		End If
		%>
		<table width="98%" border="0" align="center" cellpadding="0" cellspacing="1">
			
			<tr>
				<td  class="splittd" height="35"><strong>礼品名称:</strong><%=RS("ProductName")%> 
				<%if rs("recommend")=1 then response.write "<font color=red>推荐</font>"%>
				</td>
			</tr>
			<tr>
				<td class="splittd" height="35"><strong>添加时间:</strong><%=RS("adddate")%></td>
			</tr>
			<tr>
			    <td  class="splittd" height="35">
				   <strong>剩余数量:</strong><%=rs("Quantity")%> </td>
			</tr>
			<tr>
			    <td  class="splittd" height="35">
				   <strong> 浏 览 数:</strong><%=rs("hits")%>次</td>
			</tr>
			<tr>
			    <td  class="splittd" height="35">
				   <strong> 所需积分:</strong><%=rs("score")%></td>
			</tr>
			<tr>
			    <td  class="splittd" height="35">
				   <strong> 截止时间:</strong><%=rs("enddate")%></td>
			</tr>
			<tr>
			    <td  class="splittd" height="35">
				   <strong> 简要介绍:</strong><%=rs("intro")%></td>
			</tr>
			<form name="myform" action="?action=exchange" method="post">
			<input type="hidden" value="<%=rs("id")%>" name="id">
			<tr>
			    <td  class="splittd" align="center">
				   
				   <input type="submit" value="我要兑换" class="button">
				   <input type="button" onClick="history.back()" value="返回上一级" class="button">
		      </td>
			</tr>
			</form>
			
        </table>		    	
		
		<%
		 RS.Close:Set RS=Nothing
	   End Sub
	   
	   Sub exchange()
	    Dim ID:ID=KS.ChkClng(KS.S("ID"))
	    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select * From KS_MallScore Where ID=" & ID & " And Status=1",conn,1,3
		If RS.Eof And RS.Bof Then
		  Rs.Close
		  Call KS.AlertHistory("对不起,参数出错!",-1)
		  Exit Sub
		End If
		IF KS.ChkClng(KSUser.Score)< KS.ChkClng(RS("Score")) Then
		  Call KS.AlertHistory("对不起,您的积分不足!",-1)
		  Exit Sub
		ElseIf KS.ChkClng(RS("Quantity"))<=0 Then
		  Call KS.AlertHistory("对不起,该礼品已兑换完毕!",-1)
		  Exit Sub
		ElseIf DateDiff("s",rs("enddate"),now)>0 Then
		  Call KS.AlertHistory("对不起,该礼品已截止兑换!",-1)
		  Exit Sub
		End If
		
	   '生成订单号
	   Dim OrderID:OrderID="EX" & Year(Now)&right("0"&Month(Now),2)&right("0"&Day(Now),2)&KS.MakeRandom(8)

		
		%>
		<script language="javascript">
		 function check(){
		  if ($("input[name=RealName]").val()=="")
		  {
		    alert('请输入收货人!');
			$("input[name=RealName]").focus();
			return false;
		   }
		  if ($("input[name=Address]").val()=="")
		  {
		    alert('请输入收货地址!');
			$("input[name=Address]").focus();
			return false;
		   }
		  if ($("input[name=Tel]").val()=="")
		  {
		    alert('请输入联系电话!');
			$("input[name=Tel]").focus();
			return false;
		   }
		  if ($("input[name=ZipCode]").val()=="")
		  {
		    alert('请输入邮编!');
			$("input[name=ZipCode]").focus();
			return false;
		   }
		 }
		 
		</script>
		<table width="98%" border="0" align="center" cellpadding="0" cellspacing="1">
			<tr>
			  <td style="font-size:14px">
			  		亲爱的<font color=red><%=KSUser.UserName%></font>用户!您当前可用积分:<font color=red><%=KSUser.Score%></font>分,最多可兑换<font color=blue><%=Cint(KSUser.Score/RS("Score"))%></font>件,如果您确定兑换本礼品,请认真填写以下收货信息,兑换订单一旦提交,便不可取消!              </td>
			</tr>
			<tr>
				<td  class="splittd" height="35"><strong>礼品名称:</strong>
				<%=RS("ProductName")%>	<font color=#999999>(剩余<%=rs("Quantity")%>件)</font>			</td>
			</tr>
			<tr>
				<td class="splittd" height="35"><strong>所需积分:</strong>
				<%=RS("score")%> 分</td>
			</tr>
			<form name="myform" action="?action=exchangesave" method="post">
			<input type="hidden" value="<%=rs("id")%>" name="id">
			<input type="hidden" value="<%=orderid%>" name="orderid">
			<tr>
			    <td  class="splittd" height="35">
				   <strong>订单编号:</strong>
				   <font color=green><%=OrderID%></font> </td>
			</tr>
			<tr>
			    <td  class="splittd" height="35">
				   <strong>兑换数量:</strong>
				   <select name="amount">
				   <%dim k,endnum
				   endnum=Cint(KSUser.Score/RS("Score"))
				   if endnum>rs("Quantity") then endnum=rs("Quantity")
				   for k=1 to endnum
				    response.write "<option value=" & k & ">" & k & "</option>"
				   next
				   %>
				   </select> 件</td>
			</tr>
			<tr>
			  <td  class="splittd" height="35"><strong>收货方式:
			      <select name="DeliveryType">
                  <option value="1">快递到付</option>
                  <option value="2">自取</option>
                </select>
			  </strong></td>
			  </tr>
			<tr>
			    <td  class="splittd" height="35">
				   <strong> 收 货 人:</strong>
				   <input name="RealName" type="text" class="textbox" value="<%=KSUser.RealName%>" maxlength="30"> <font color=red>*</font></td>
			</tr>
			<tr>
			    <td  class="splittd" height="35">
				   <strong> 收货地址:</strong>
				   <input name="Address" type="text"class="textbox" value="<%=KSUser.Address%>" size="40" maxlength="255"> <font color=red>*</font></td>
			</tr>
			<tr>
			    <td  class="splittd" height="35">
				   <strong> 邮政编码:</strong>
				   <input name="ZipCode" type="text"class="textbox" value="<%=KSUser.Zip%>" id="ZipCode" size="10" maxlength="10"> <font color=red>*</font></td>
			</tr>
			<tr>
			  <td  class="splittd" height="35"><strong>联系电话:
			    <input name="Tel" type="text" class="textbox" id="Tel" value="<%=KSUser.OfficeTEL%>" maxlength="30"> <font color=red>*</font>
			  </strong></td>
			  </tr>
			<tr>
			  <td  class="splittd" height="35"><strong>电子邮箱:
			      <input name="Email" type="text" class="textbox" value="<%=KSUser.Email%>" id="Email" maxlength="50">
			  </strong></td>
			  </tr>
			<tr>
			    <td  class="splittd" height="35">
				   <strong> 备注说明:</strong>
				   <textarea name="Remark" cols="50" rows="5" class="textbox" style="height:60px"></textarea></td>
			</tr>
			<tr>
			    <td  class="splittd" align="center">
				   
				   <p>
				     <br>
				     <input type="submit" onClick="return(check())" value="我要兑换" class="button">
				     <input type="button" onClick="history.back()" value="返回上一级" class="button">
			      </p>
				   <p>&nbsp;</p>
				   <p>&nbsp;				      </p></td>
			</tr>
			</form>
        </table>		    	
		
		<%
		 RS.Close:Set RS=Nothing
	   End Sub
	   
	   Sub exchangesave()
	    Dim ID:ID=KS.ChkClng(KS.S("ID"))
	    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select * From KS_MallScore Where ID=" & ID & " And Status=1",conn,1,3
		If RS.Eof And RS.Bof Then
		  Rs.Close
		  Call KS.AlertHistory("对不起,参数出错!",-1)
		  Exit Sub
		End If
		IF KSUser.Score< RS("Score") Then
		  Call KS.AlertHistory("对不起,您的积分不足!",-1)
		  Exit Sub
		ElseIf KS.ChkClng(RS("Quantity"))<=0 Then
		  Call KS.AlertHistory("对不起,该礼品已兑换完毕!",-1)
		  Exit Sub
		ElseIf DateDiff("s",rs("enddate"),now)>0 Then
		  Call KS.AlertHistory("对不起,该礼品已截止兑换!",-1)
		  Exit Sub
		End If
		
		
		
	   '生成订单号
	   Dim OrderID:OrderID=KS.S("OrderID")
	   If OrderID="" Then 
	     Call KS.AlertHistory("对不起,参数出错啦!",-1)
		 Exit Sub
	   End If
	   Dim amount:amount=KS.ChkClng(KS.S("amount"))
	   Dim RealName:RealName=KS.S("RealName")
	   Dim Address:Address=KS.S("Address")
	   Dim Tel:Tel=KS.S("Tel")
	   Dim ZipCode:ZipCode=KS.S("ZipCode")
	   Dim Email:Email=KS.S("Email")
	   Dim Remark:Remark=KS.S("Remark")
	   Dim DeliveryType:DeliveryType=KS.ChkClng(KS.S("DeliveryType"))
	   If Amount=0 Or Amount>rs("Quantity") Then
	     Call KS.AlertHistory("对不起,兑换数量不正确!",-1)
		 Exit Sub
	   End IF
	   If RealName="" Then
	     Call KS.AlertHistory("对不起,收货人必须填写!",-1)
		 Exit Sub
	   End If
	   If Address="" Then
	     Call KS.AlertHistory("对不起,收货地址必须填写!",-1)
		 Exit Sub
	   End If
	   
	   Dim RSO:Set RSO=Server.CreateObject("ADODB.RECORDSET")
       RSO.Open "Select * From KS_MallScoreOrder Where OrderID='" & OrderID & "' And ProductID=" &ID,conn,1,3
	   If RSO.Eof Then
		   RSO.AddNew
			RSO("OrderID")=OrderID
			RSO("ProductID")=ID
			RSO("UserName")=KSUser.UserName
			RSO("Amount")=Amount
			RSO("RealName")=RealName
			RSO("Address")=Address
			RSO("ZipCode")=zipcode
			RSO("Tel")=Tel
			RSO("Email")=Email
			RSO("Remark")=Remark
			RSO("DeliveryType")=DeliveryType
			RSO("AddDate")=Now
			RSO("Status")=0
		   RSO.Update 
		   
		   '更新可用数量
		   RS("Quantity")=RS("Quantity")-Amount
		   RS.Update
		   '更新用户积分
		   Call KS.ScoreInOrOut(KSUser.UserName,2,RS("Score")*Amount,"系统","兑换订单号<font color=red>" & OrderID & "</font>的礼品!",0,0)
		   
		   Call KS.SendInfo(KSUser.UserName,"system","恭喜，成功兑换礼品[" & RS("ProductName") &"]！","亲爱的" & KSUser.UserName & "!<br />&nbsp;&nbsp;&nbsp;&nbsp;恭喜您!订单号<font color=red>" & OrderID & "</font>的礼品兑换成功，请注意查收您的礼品。<br />&nbsp;&nbsp;&nbsp;&nbsp;您本次兑换共消费 <font color=red>" & RS("Score")*Amount & "</font>分积分！")
	   End If
	   RSO.Close:Set RSO=Nothing
		   RS.Close:Set RS=Nothing
	   
		%>
		
		<table width="98%" border="0" align="center" cellpadding="0" cellspacing="1">
			<tr>
			  <td style="font-size:14px;text-align:center">
			    <br><br><br>恭喜您!订单号<font color=red><%=OrderID%></font>的礼品兑换成功，请注意查收您的礼品。              </td>
			</tr>
		
			<tr>
			    <td  class="splittd" align="center">
				   
				   <p>
				     <br>  <br>  <br>  <br>  <br>
				     <input type="button" onClick="location.href='?'" value="返回上一级" class="button">
			      </p>
				   <p>&nbsp;</p>
				   <p>&nbsp;				      </p></td>
			</tr>
			</form>
        </table>		    	
		
		<%
	   End Sub
	   
	   
	   
	   '显示订单
	 sub ShowOrder()
		%>
			
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="1">
			<tr height="28" class="title">
				<td height="25" align="center">订单号</td>
				<td height="25" align="center">礼品名称</td>
				<td align="center">兑换数量</td>
				<td  align="center">消费积分</td>
				<td  align="center">兑换时间</td>
				<td align="center">状态</td>
				<td align="center">收货方式</td>
				<td align="center">操作</td>
			</tr>
		<%  dim sql
			set rs=server.createobject("adodb.recordset")
			sql="select a.*,b.productname,b.score from KS_MallScoreOrder a inner join KS_MallScore b on a.productid=b.id where a.Username='"&KSUser.UserName&"' order by a.id desc"
			rs.open sql,Conn,1,1
			if rs.eof and rs.bof then
		%>
			<tr>
				<td height="26" colspan=7 align=center valign=middle  class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">您没有兑换记录！</td>
			</tr>
		<%else
		
		                       totalPut = RS.RecordCount
								If CurrentPage < 1 Then	CurrentPage = 1
								
								   If CurrentPage > 1  and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
									Else
										CurrentPage = 1
									End If
									i=0
		      do while not rs.eof
		%>
						<tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
							<td class="splittd" align="center"><a href="?action=showdetail1&id=<%=rs("id")%>"><font color=green><%=rs("orderid")%></font></a></td>
							<td height="25" class="splittd">
							<div class="ContentTitle"><%=KS.HTMLEncode(rs("productname"))%></div>
							</td>
							<td class="splittd" align=center>
							<%=RS("Amount")%>
							</td>
							<td class="splittd" align=center>
							<%=RS("Amount")%><font color=red>*</font><%=RS("Score")%>=<%=RS("Amount")*RS("Score")%>
							</td>
							<td class="splittd" align=center>
							<%=RS("AddDate")%>
							</td>
							<td class="splittd" align=center>
							<%select case  rs("status")
								 case 1
								  response.write "已审"
								 case 2
								  response.write "<font color=blue>配货中</font>"
								 case 3
								  response.write "<font color=#ff6600>已发货</font>"
								 case 4
								  response.write "<font color=#999999>交易完成</font>"
								 case 5
								  response.write "<font color=green>无效(积分已退回)</font>"
								 case else
								  response.write " <font color=red>未审</font>"
								end select
							%>
							</td>
							<td class="splittd" align=center>
							<%if rs("DeliveryType")=1 then response.write "快递到付" else response.write "自取"%>
							</td>
							
							<td class="splittd" align=center>
							 <%if rs("status")<>0 and rs("status")<>4 and rs("status")<>5 then%>
							<a  href="User_ScoreExchange.asp?action=setok&id=<%=rs("id")%>" onclick = "return (confirm('确定收到货了吗?'))">设置已收货</a>
							 <%else%>
							  ---
							 <%end if%>
							 
							 
							</td>
						</tr>
		<%
			  rs.movenext
			  I = I + 1
			  If I >= MaxPerPage Then Exit Do
			
			loop
			end if
			rs.close
			set rs=Nothing
		%>
						
				
</table>
   
    <div style="text-align:right">
  <%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
    </div>

		<%
		end sub
	   
	   


     Sub showdetail1()
	    Dim ID:ID=KS.ChkClng(KS.S("ID"))
	    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select a.*,b.productname,score From KS_MallScoreOrder a Left Join KS_MallScore b on a.productid=b.id Where a.ID=" & ID,conn,1,3
		If RS.Eof And RS.Bof Then
		  Rs.Close
		  Call KS.AlertHistory("对不起,参数出错!",-1)
		  Exit Sub
		End If
		%>
		<script language="javascript">
		 function check(){
		  if ($("input[name=RealName]").val()=="")
		  {
		    alert('请输入收货人!');
			$("input[name=RealName]").focus();
			return false;
		   }
		  if ($("input[name=Address]").val()=="")
		  {
		    alert('请输入收货地址!');
			$("input[name=Address]").focus();
			return false;
		   }
		  if ($("input[name=Tel]").val()=="")
		  {
		    alert('请输入联系电话!');
			$("input[name=Tel]").focus();
			return false;
		   }
		  if ($("input[name=ZipCode]").val()=="")
		  {
		    alert('请输入邮编!');
			$("input[name=ZipCode]").focus();
			return false;
		   }
		 }
		 
		</script>
		<table width="98%" border="0" align="center" cellpadding="0" cellspacing="1">
			
			<tr>
				<td  class="splittd" height="35"><strong>礼品名称:</strong>
				<%=RS("ProductName")%>			</td>
			</tr>
			<tr>
				<td  class="splittd" height="35"><strong>订 单 号:</strong>
				<%=RS("orderid")%>			</td>
			</tr>
			<tr>
				<td class="splittd" height="35"><strong>兑换时间:</strong>
				<%=RS("adddate")%></td>
			</tr>
			<form name="myform" action="?action=dosave" method="post">
			<input type="hidden" value="<%=rs("id")%>" name="id">
			
			<tr>
			    <td  class="splittd" height="35">
				   <strong>兑换数量:</strong>
				   <%=rs("amount")%> 件</td>
			</tr>
			<tr>
			  <td  class="splittd" height="35"><strong>收货方式:</strong>
			  <%if rs("DeliveryType")=1 then
			    response.write "快递到付"
			   else
			    response.write "自取"
			   end if%>
			  </td>
			  </tr>
			  <tr>
			  <td  class="splittd" height="35"><strong>订单状态:</strong>
			  <%select case  rs("status")
								 case 1
								  response.write "已审"
								 case 2
								  response.write "<font color=blue>配货中</font>"
								 case 3
								  response.write "<font color=#ff6600>已发货</font>"
								 case 4
								  response.write "<font color=#999999>交易完成</font>"
								 case else
								  response.write " <font color=red>未审</font>"
								end select
							%>
			  </td>
			  </tr>
			<tr>
			    <td  class="splittd" height="35">
				   <strong> 收 货 人:</strong>
				   <input name="RealName" type="text" class="textbox" value="<%=rs("RealName")%>" maxlength="30"> <font color=red>*</font></td>
			</tr>
			<tr>
			    <td  class="splittd" height="35">
				   <strong> 收货地址:</strong>
				   <input name="Address" type="text"class="textbox" value="<%=rs("Address")%>" size="40" maxlength="255"> <font color=red>*</font></td>
			</tr>
			<tr>
			    <td  class="splittd" height="35">
				   <strong> 邮政编码:</strong>
				   <input name="ZipCode" type="text"class="textbox" value="<%=rs("ZipCode")%>" id="ZipCode" size="10" maxlength="10"> <font color=red>*</font></td>
			</tr>
			<tr>
			  <td  class="splittd" height="35"><strong>联系电话:
			    <input name="Tel" type="text" class="textbox" id="Tel" value="<%=rs("TEL")%>" maxlength="30"> <font color=red>*</font>
			  </strong></td>
			  </tr>
			<tr>
			  <td  class="splittd" height="35"><strong>电子邮箱:
			      <input name="Email" type="text" class="textbox" value="<%=rs("Email")%>" id="Email" maxlength="50">
			  </strong></td>
			  </tr>
			<tr>
			    <td  class="splittd" height="35">
				   <strong> 备注说明:</strong>
				   <textarea name="Remark" cols="50" rows="5" class="textbox" style="height:60px"><%=rs("remark")%></textarea></td>
			</tr>
			<tr>
			    <td  class="splittd" align="center">
				   
				   <p>
				     <br>
				     <input type="submit" onClick="return(check())" value="确定修改" class="button">
				     <input type="button" onClick="history.back()" value="返回上一级" class="button">
			      </p>
				   <p>&nbsp;</p>
				   <p>&nbsp;				      </p></td>

			</tr>
			</form>
        </table>		    	
		
		<%
		 RS.Close:Set RS=Nothing
	   End Sub
		
	   Sub SetOrderOk()
		 conn.execute("update KS_MallScoreOrder Set Status=4 Where ID=" & KS.ChkClng(KS.S("ID")) & " And UserName='" & KSUser.UserName & "'")
		 Response.Redirect ComeUrl
	   End Sub
	   
	   Sub dosave()
	       Dim ID:ID=KS.ChkClng(KS.G("id"))
		   Dim Address:Address=KS.G("Address")
		   Dim RealName:RealName=KS.G("RealName")
		   Dim ZipCode:ZipCode=KS.G("ZipCode")
		   Dim Tel:Tel=KS.G("Tel")
		   Dim ComeUrl:ComeUrl=KS.G("ComeUrl")
		   Dim Remark:Remark=KS.G("Remark")
		   Dim Email:Email=KS.G("Email")
		   Dim DeliveryType:DeliveryType=KS.ChkClng(KS.G("DeliveryType"))
		
	       If RealName="" Then Response.Write "<script>alert('收货人必须输入');history.back();</script>":response.end

            Dim RS: Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select * From KS_MallScoreOrder Where ID=" & ID,Conn,1,3
				 RS("RealName")=RealName
				 RS("Address")=Address
				 RS("ZipCode")=ZipCode
				 RS("Tel")=Tel
				 RS("Remark")=Remark
				 RS("Email")=Email
		 		 RS.Update
			     RS.Close
				 Set RS=Nothing
            KS.AlertHintScript "订单收货信息修改成功!"
	   End Sub

End Class
%> 
