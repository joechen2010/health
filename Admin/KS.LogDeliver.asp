<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Admin_LogDeliver
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_LogDeliver
        Private KS,KSCls
		Private totalPut,rs, CurrentPage, MaxPerPage,DomainStr,SearchType,SQLParam
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub
		Public Sub Kesion()
		If Not KS.ReturnPowerResult(5, "M510015") Then  Call KS.ReturnErr(1, ""):Exit Sub   
		SearchType=KS.ChkClng(KS.G("SearchType"))
		%>
<html>
<head><title>发退货查询</title>
<meta http-equiv=Content-Type content="text/html; charset=gb2312">
<link href="include/Admin_Style.css" type=text/css rel=stylesheet>
</head>
<body leftMargin=2 topMargin=0 marginheight="0" marginwidth="0">
  <div class="topdashed" style="padding:4px;">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
<FORM name=form1 action=KS.LogDeliver.asp method=get>
      <td><strong>发退货查询：</strong></td>
      <td valign="top">快速查询： 
<Select onchange=javascript:submit() size=1 name=SearchType class='textbox'> 
  <Option value=0<%If SearchType="0" Then Response.write " selected"%>>所有发退货记录</Option> 
  <Option value=1<%If SearchType="1" Then Response.write " selected"%>>最近10天内的新记录</Option> 
  <Option value=2<%If SearchType="2" Then Response.write " selected"%>>最近一月内的新记录</Option> 
  <Option value=3<%If SearchType="3" Then Response.write " selected"%>>所有发货记录</Option> 
  <Option value=4<%If SearchType="4" Then Response.write " selected"%>>所有退货记录</Option>
      </Select>&nbsp;&nbsp;&nbsp;&nbsp;<a href="KS.LogDeliver.asp">发退货记录首页</a></td></FORM>
<FORM name=form2 action=KS.LogDeliver.asp method=post>
      <td>高级查询： 
<Select id=Field name=Field class='textbox'> 
  <Option value=1 selected>客户姓名</Option> 
  <Option value=2>用户名</Option> 
  <Option value=3>发退货日期</Option> 
  <Option value=4>经手人</Option> 
  <Option value=5>快递公司</Option> 
  <Option value=6>快递单号</Option> 
  <Option value=7>订单号</Option>
</Select> 
  <Input id=Keyword class='textbox' maxLength=30 name=Keyword> 
  <Input class='button' type=submit value=" 查 询 " name=Submit2> 
        <Input id=SearchType type=hidden value=5 name=SearchType> </td></FORM>
    </tr>  </table>
  </div>

  <div style="border:1px #cccccc dashed;overflow:hidden"></div>
  <table width="100%">
    <tr>
      <td align=left>您现在的位置：<a href="KS.LogDeliver.asp">发退货记录管理</a>&nbsp;&gt;&gt;&nbsp;
	  <%Dim SearchTypeStr
	    Dim KeyWord:KeyWord=KS.G("KeyWord")
	  Select Case SearchType
	     Case 0 :SearchTypeStr="所有记录"
		 Case 1 :SearchTypeStr="最近10天内的新记录"
		 Case 2 :SearchTypeStr="最近一月内的新记录"
		 Case 3 :SearchTypeStr="所有发货记录"
		 Case 4 :SearchTypeStr="所有退货记录"
		 Case 5 
		    Select Case KS.ChkClng(KS.G("Field"))
			  Case 1:SearchTypeStr="客户姓名含有<font color=red>""" & KeyWord & """</font>"
			  Case 2:SearchTypeStr="用户名含有<font color=red>""" & KeyWord & """</font>"
			  Case 3:SearchTypeStr="发退货日期含有<font color=red>""" & KeyWord & """</font>"
			  Case 4:SearchTypeStr="经手人含有<font color=red>""" & KeyWord & """</font>"
			  Case 5:SearchTypeStr="快递公司含有<font color=red>""" & KeyWord & """</font>"
			  Case 6:SearchTypeStr="快递单号含有<font color=red>""" & KeyWord & """</font>"
			  Case 7:SearchTypeStr="订单号含有<font color=red>""" & KeyWord & """</font>"
			End Select
	  End Select
	  Response.Write SearchTypeStr%></td>
    </tr>
  </table>
  <div style="border:1px #cccccc dashed;overflow:hidden"></div>
  <table cellSpacing=1 cellPadding=0 width="100%" border=0>
    <tr class=sort align=middle>
      <td width=70>日期</td>
	  <td width=80>订单编号</td>
	  <td width=70>用户名</td>
      <td width=40>方向</td>
      <td width=80>客户姓名</td>
      <td width=120>快递公司</td>
      <td width=60>快递单号</td>
      <td width=60>经手人</td>
      <td width=40>签收</td>
      <td>备注</td>
    </tr>
	<%
			MaxPerPage=20
			If KS.G("page") <> "" Then
				  CurrentPage = KS.ChkClng(KS.G("page"))
			Else
				  CurrentPage = 1
			End If
			SqlParam="1=1"
            If SearchType<>"0" Then
			  Select Case SearchType
			   Case 1
			   		if DataBaseType=1 then
					SqlParam=SqlParam &" And datediff(d,DeliverDate," & SqlNowString & ")<=10"
				   else
					SqlParam=SqlParam &" And datediff('d',DeliverDate," & SqlNowString & ")<=10"
				   end if
			   Case 2
			   		if DataBaseType=1 then
					SqlParam=SqlParam &" And datediff(d,DeliverDate," & SqlNowString & ")<=30"
				   else
					SqlParam=SqlParam &" And datediff('d',DeliverDate," & SqlNowString & ")<=30"
				   end if
			  Case 3 : SqlParam = SqlParam & "And status=1"
			  Case 4 : SqlParam = SqlParam & "And DeliverType=2"
			  Case 5
			      Select Case KS.ChkClng(KS.G("Field"))
				   Case 1:SqlParam=SqlParam &" And ClientName Like '%" & Keyword & "%'"
				   Case 2:SqlParam=SqlParam &" And UserName Like '%" & Keyword & "%'"
				   Case 3:SqlParam=SqlParam &" And DeliverDate Like '%" & Keyword & "%'"
				   Case 4:SqlParam=SqlParam &" And HandlerName Like '%" & Keyword & "%'"
				   Case 5:SqlParam=SqlParam &" And ExpressCompany Like '%" & Keyword & "%'"
				   Case 6:SqlParam=SqlParam &" And ExpressNumber Like '%" & Keyword & "%'"
				   Case 7:SqlParam=SqlParam &" And OrderID Like '%" & Keyword & "%'"
				  End Select
			  End Select
			End If
	Set RS=Server.CreateObject("ADODB.RECORDSET")
	RS.Open "Select * From KS_LogDeliver Where " & SqlParam & " Order By ID Desc",Conn,1,1
	If RS.Eof AND RS.Bof Then
	 Response.WRITE "<tr class=list onmouseover=""this.className='listmouseover'"" onmouseout=""this.className='list'""><td colspan=9 align=center height='25'>找不到" & SearchTypeStr & "!</td></tr>"
   Else
                          totalPut = RS.RecordCount
							If CurrentPage < 1 Then	CurrentPage = 1
							If (CurrentPage - 1) * MaxPerPage > totalPut Then
								If (totalPut Mod MaxPerPage) = 0 Then
									CurrentPage = totalPut \ MaxPerPage
								Else
									CurrentPage = totalPut \ MaxPerPage + 1
								End If
							End If
							If CurrentPage = 1 Then
								Call showContent()
							Else
								If (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
									Call showContent()
								Else
									CurrentPage = 1
									Call showContent()
								End If
							End If
   End If
   RS.Close:Set RS=Nothing
   %>
     <div align=right>
         <%
		   	  '显示分页信息
			  Call KS.ShowPageParamter(totalPut, MaxPerPage, "", True, "条记录", CurrentPage, "SearchType=" & SearchType & "&Field=" & KS.G("Field") & "&KeyWord=" & KeyWord)
		   %>
    </div>
	<br>

</body>
<html>
   <%
   End Sub
  
  Sub ShowContent()
     Dim I,intotalDeliver,outtotalDeliver
     Do While Not rs.eof 
	%>
    <tr class=list onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
      <td height="20" align=middle><%=rs("DeliverDate")%></td>
      <td align=middle><%=rs("orderid")%></td>
	  <td align=center><%=rs("username")%></td>
	  <td align=middle>
	  <%
	  If rs("DeliverType")=1 Then
	   response.write "发货"
	  Else
	   Response.Write "退货"
	  End If
	  %></td>
      <td align=middle><%=rs("ClientName")%></td>
      <td align=middle><%=rs("ExpressCompany")%></td>
      <td align=center><%=rs("ExpressNumber")%></td>
      <td align=center><%=rs("HandlerName")%></td>
      <td align=center>
	  <% If rs("status")=1 Then
	      Response.Write "<font color=red>√</font>"
		 End If
		 %></td>
      <td align=middle><%=rs("Remark")%></td>
    </tr>
	<tr><td height="1" colspan="10"><div style="border:1px #cccccc dashed;overflow:hidden"></div></td></tr>
	<%
	            
				I = I + 1
				RS.MoveNext
				If I >= MaxPerPage Then Exit Do

	 loop
	%>

  </table>
		<%
		End Sub
End Class
%> 
