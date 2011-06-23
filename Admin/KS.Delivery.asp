<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Delivery
KSCls.Kesion()
Set KSCls = Nothing

Class Delivery
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
       Sub Kesion()
	     If Not KS.ReturnPowerResult(5, "M520004") Then  Call KS.ReturnErr(1, ""):Exit Sub
	     Dim RS
		 Dim TypeID:TypeID=2 '影视服务器
         With Response
		   .Write "<html>"
			.Write"<head>"
			.Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			.Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
			.Write"</head>"
			.Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			.Write "<ul id='menu_top'>"
			.Write "<li id='p7'><a href='KS.Delivery.asp'>送货方式</a></li>"
			.Write "| <li id='p8'><a href='KS.PaymentType.asp'>付款方式</a></li>"
			.Write	" </ul>"
		End With
%>		
		  
		<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
		  <tr align="center"  class="sort"> 
			<td width="87"><strong>编号</strong></td>
			<td width="217"><strong>送货方式名称</strong></td>
			<td width="197"><strong>加收金额</strong></td>
			<td width="197"><strong>排序</strong></td>
			<td width="197"><strong>是否默认</strong></td>
			<td width="196"><strong>管理操作</strong></td>
		  </tr>
		  <%dim orderid
		  set rs = conn.execute("select * from KS_Delivery order by orderid")
		    if rs.eof and rs.bof then
			  Response.Write "<tr><td colspan=""6"" height=""25"" align=""center"" class=""tdbg"">还没有添加任何的收货方式!</td></tr>"
			else
			   do while not rs.eof%>
			  <form name="form1" method="post" action="?x=a">
				<tr  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'"> 
				  <td width="87" height="25" align="center"><%=rs("typeid")%> <input name="typeid" type="hidden" id="typeid" value="<%=rs("typeid")%>"></td>
				  <td width="217" align="center"><input style="color:<%=rs("Fee")%>" name="TypeName" type="text" class="textbox" id="TypeName" value="<%=rs("TypeName")%>" size="25"></td>
				  <td width="197" align="center"><input style="text-align:center" name="Fee" type="text" class="textbox" id="Fee" value="<%=rs("Fee")%>" size="8">
				  元</td>				  
				  <td width="197" align="center"><input style="text-align:center" name="OrderID" type="text" class="textbox" id="OrderID" value="<%=rs("OrderID")%>" size="8">
				  </td>
				  <td width="197" align="center">
				  <a href="?x=d&typeid=<%=rs("typeid")%>">
				  <%If RS("IsDefault")="1" Then
				     Response.Write "<font color=red>是</font>"
					Else
					 Response.Write "否"
					End If
				  %>
				  </a>
				  </td>
				  <td align="center"><input name="Submit" class="button" type="submit"value=" 修改 ">&nbsp;<input  onclick='if (confirm("确定删除吗？")==true){window.location="?x=c&typeid=<%=rs("typeid")%>";}' name="Submit2" type="button" class="button" value=" 删除 "></td>
				</tr>
				 <tr><td colspan=9 background='images/line.gif'></td></tr>
			  </form>
		  <%orderid=rs("orderid")
		   rs.movenext
		   loop
		 End IF
		rs.close%>
				<form action="?x=b" method="post" name="myform" id="form">
		    <tr>
			<td height="25" colspan="6">&nbsp;&nbsp;<strong>&gt;&gt;新增送货方式</strong><<</td>
		    </tr>
			<tr><td colspan=9 background='images/line.gif'></td></tr>
			<tr valign="middle" class="list"> 
			  <td height="25"></td>
			  <td height="25" align="center"><input name="TypeName" type="text" class="textbox" id="TypeName" size="25"></td>
			  <td height="25" align="center"><input style="text-align:center" name="Fee1" type="text" class="textbox" id="Fee1" size="8">
元</td>
			  <td height="25" align="center"><input style="text-align:center" name="orderid" type="text" value="<%=orderid+1%>" class="textbox" id="orderid" size="8">
			  <td height="25" align="center"><input name="isdefault" type="checkbox" value="1" size="8">设为默认
</td>
			  <td height="25" align="center"><input name="Submit3" class="button" type="submit" value="OK,提交"></td>
			</tr>
			<tr><td colspan=9 background='images/line.gif'></td></tr>
		</form>
</table>

		<% Select case request("x")
		   case "a"
		   		If Not Isnumeric(KS.G("Fee")) Then Response.Write "<script>alert('折扣率必须用数字!');history.back();</script>":response.end
				conn.execute("Update KS_Delivery set TypeName='" & KS.G("TypeName") & "',Fee='" & KS.G("Fee") & "',orderid='" & KS.ChkClng(KS.G("OrderID")) &"' where Typeid="&KS.G("typeid")&"")
				Response.Redirect "?"
		   case "b"
		       If KS.G("TypeName")="" Then Response.Write "<script>alert('请输入送货方式名称!');history.back();</script>":response.end
			   If Not Isnumeric(KS.G("Fee1")) Then Response.Write "<script>alert('折扣率必须用数字!');history.back();</script>":response.end
				conn.execute("Insert into KS_Delivery(TypeName,Fee,orderid)values('" & KS.G("TypeName") & "','" & KS.G("Fee1") & "','" & KS.ChkClng(KS.G("OrderID")) &"')")
				If KS.G("isdefault")="1" Then
				 Conn.execute("update KS_Delivery Set IsDefault=0")
				 Conn.execute("update KS_Delivery Set IsDefault=1 Where TypeID=" & Conn.execute("select max(typeid) from KS_Delivery")(0))
				End If
				Response.Redirect "?"
		   case "c"
				conn.execute("Delete from KS_Delivery where Typeid="&KS.G("typeid")&"")
				Response.Redirect "?"
		   case "d"
				 Conn.execute("update KS_Delivery Set IsDefault=0")
				 Conn.execute("update KS_Delivery Set IsDefault=1 Where TypeID=" & KS.ChkClng(KS.G("TypeID")))
				Response.Redirect "?"
		End Select
		%></body>
		</html>
<%End Sub
End Class
%> 
