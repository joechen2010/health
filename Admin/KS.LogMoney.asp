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
Set KSCls = New Admin_ShopOrder
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_ShopOrder
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
	    If Not (KS.ReturnPowerResult(0, "KMUA10007") or KS.ReturnPowerResult(5,"M510014"))Then
			  Call KS.ReturnErr(1, "")
			End If
		SearchType=KS.ChkClng(KS.G("SearchType"))
		 dim begindate:begindate=request("begindate")
		 dim enddate:enddate=request("enddate")

		%>
<html>
<head><title>�ʽ���ϸ��ѯ</title>
<meta http-equiv=Content-Type content="text/html; charset=gb2312">
<link href="include/Admin_Style.css" type=text/css rel=stylesheet>
</head>
<body leftMargin=2 topMargin=0 marginheight="0" marginwidth="0">
  <div class="topdashed" style="padding:4px;">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
<FORM name=form1 action=KS.LogMoney.asp method=get>
      <td>�ʽ���ϸ��ѯ��</td>
      <td valign="top">���ٲ�ѯ�� 
<Select onchange=javascript:submit() size=1 name=SearchType class='textbox'> 
  <Option value=0<%If SearchType="0" Then Response.write " selected"%>>�����ʽ���ϸ��¼</Option> 
  <Option value=1<%If SearchType="1" Then Response.write " selected"%>>���10���ڵ����ʽ���ϸ��¼</Option> 
  <Option value=2<%If SearchType="2" Then Response.write " selected"%>>���һ���ڵ����ʽ���ϸ��¼</Option> 
  <Option value=3<%If SearchType="3" Then Response.write " selected"%>>���������¼</Option> 
  <Option value=4<%If SearchType="4" Then Response.write " selected"%>>����֧����¼</Option>
      </Select>
	  </td></FORM>
<FORM name=form2 action=KS.LogMoney.asp method=post>
      <td style="border:1px #cccccc dashed">�߼���ѯ�� 
<Select id=Field name=Field class='textbox'> 
  <Option value=1 selected>�ͻ�����</Option> 
  <Option value=2>�û���</Option> 
  <Option value=3>����ʱ��</Option> 
</Select> 
  <Input id=Keyword class='textbox' maxLength=30 name=Keyword> 
  <Input class='button' type=submit value=" �� ѯ " name=Submit2> 
        <Input id=SearchType type=hidden value=5 name=SearchType> </td></FORM>
    </tr>
  </table>
  </div>

  <div style="border:1px #cccccc dashed;overflow:hidden"></div>
  
   <table width="100%" border="0">
<form action="?action=search&SearchType=100" method=post name="myform">
   
   <tr>
     <td width="12%"><strong>��ʱ��β�ѯ</strong></td>
     <td width="48%">
       
       <table width="100%"  align="center" border=0 cellPadding=0 cellSpacing=0>
         <tr>
           <td nowrap="nowrap" class=form-left>��ʼ���ڣ�
             <%if isdate(begindate) then%>
             <input type="text" name="begindate" value="<%=begindate%>" size="12" class="form-input">
             <%else%>
             <input type="text" name="begindate" value="<%=year(now)&"-"&month(now)&"-1"%>" size="12" class="form-input">
             <%end if%>
             <br>
             <font color="#999999">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�磺2008-1-1</font>            </td>    
		    <td  class=form-left>��ֹ���ڣ�  
		      <%if isdate(enddate) then%>            
		      <input type="text" name="enddate" value="<%=enddate%>" size="12" class="form-input">
		      <%else%>
		      <input type="text" name="enddate" value="<%=formatdatetime(now,2)%>" size="12" class="form-input">
		      <%end if%>
		      <br>
		      <font color="#999999">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�磺2008-1-31</font>	        </td>    
			    </tr>
        </table>	 </td>
     <td width="43%">��־��<input type="radio" name="direction" value="0"<%if request("direction")="0" or request("direction")="" then response.write " checked"%>>����<input type="radio" name="direction" value="1"<%if request("direction")="1" then response.write " checked"%>>���� <input type="radio" name="direction" value="2"<%if request("direction")="2" then response.write " checked"%>>֧��	  �ؼ���:
      <input type="text" name="keyword" size="10" value="<%=request("keyword")%>"/>
      <input name="submit2" type="submit" value="���ٲ���" />
      </td>
    </tr>
</form>
 </table>
<table width="100%">
    <tr>
      <td align=left>
	  <%
	  if begindate<>"" and enddate<>"" then
	   response.write "<br><div align=center style='font-size:14px'>"
	 response.write "��ѯʱ��� <font color=red>" & begindate & "</font> �� <font color=red>" & enddate & "</font><br></div>"
	  end if
	  %></td>
    </tr>
  </table>
  
  
  
  <table width="100%">
    <tr>
      <td align=left>�����ڵ�λ�ã�<a href="KS.LogMoney.asp">�ʽ���ϸ��¼����</a>&nbsp;&gt;&gt;&nbsp;
	  <%Dim SearchTypeStr
	    Dim KeyWord:KeyWord=KS.G("KeyWord")
	  Select Case SearchType
	     Case 0 :SearchTypeStr="�����ʽ���ϸ��¼"
		 Case 1 :SearchTypeStr="���10���ڵ����ʽ���ϸ��¼"
		 Case 2 :SearchTypeStr="���һ���ڵ����ʽ���ϸ��¼"
		 Case 3 :SearchTypeStr="���������¼"
		 Case 4 :SearchTypeStr="����֧����¼"
		 Case 5 
		    Select Case KS.ChkClng(KS.G("Field"))
			  Case 1:SearchTypeStr="�ͻ���������<font color=red>""" & KeyWord & """</font>"
			  Case 2:SearchTypeStr="�û�������<font color=red>""" & KeyWord & """</font>"
			  Case 3:SearchTypeStr="����ʱ�京��<font color=red>""" & KeyWord & """</font>"
			End Select
		Case 100:SearchTypeStr="ʱ��β�ѯ���"
	  End Select
	  Response.Write SearchTypeStr%></td>
    </tr>
  </table>
    <div style="border:1px #cccccc dashed;overflow:hidden"></div>

  <table cellSpacing=0 cellPadding=0 width="100%" border=0>
    <tr class=sort align=middle>
      <td width=120>����ʱ��</td>
      <td width=80>�û���</td>
      <td width=80>�ͻ�����</td>
      <td width=60>���׷�ʽ</td>
      <td width=50>����</td>
      <td width=80>������</td>
      <td width=80>֧�����</td>
      <td width=40>ժҪ</td>
      <td width=40>���</td>
      <td>��ע/˵��</td>
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
					SqlParam=SqlParam &" And datediff('d',Logtime," & SqlNowString & ")<=10"
			   Case 2
					SqlParam=SqlParam &" And datediff('d',Logtime," & SqlNowString & ")<=30"
			  Case 3 : SqlParam = SqlParam & "And IncomeOrPayOut=1"
			  Case 4 : SqlParam = SqlParam & "And IncomeOrPayOut=2"
			  Case 5
			      Select Case KS.ChkClng(KS.G("Field"))
				   Case 1
				     SqlParam=SqlParam &" And ClientName Like '%" & Keyword & "%'"
				   Case 2
				     SqlParam=SqlParam &" And UserName Like '%" & Keyword & "%'"
				   Case 3
				     SqlParam=SqlParam &" And logtime Like '%" & Keyword & "%'"
				  End Select
			  End Select
			End If
			If CInt(DataBaseType) = 1 Then         'Sql
				if isdate(begindate) then SqlParam=SqlParam & " and logtime>='" & begindate & "'"
				if isdate(enddate) then enddate=DateAdd("d", 1,EndDate):SqlParam=SqlParam & " and logtime<='" & enddate & "'"
			else
				if isdate(begindate) then SqlParam=SqlParam & " and logtime>=#" & begindate & "#"
				if isdate(enddate) then enddate=DateAdd("d", 1,EndDate):SqlParam=SqlParam & " and logtime<=#" & enddate & "#"
			end if
			if KS.ChkClng(KS.G("direction"))<>0 Then SqlParam=SqlParam & " and IncomeOrPayOut=" & KS.ChkClng(KS.G("Direction"))

	Set RS=Server.CreateObject("ADODB.RECORDSET")
	RS.Open "Select * From KS_Logmoney Where " & SqlParam & " Order By ID Desc",Conn,1,1
	If RS.Eof AND RS.Bof Then
	 Response.WRITE "<tr class=list onmouseover=""this.className='listmouseover'"" onmouseout=""this.className='list'""><td colspan=9 align=center height='25'>�Ҳ���" & SearchTypeStr & "!</td></tr>"
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
		   	  '��ʾ��ҳ��Ϣ
			  Call KS.ShowPageParamter(totalPut, MaxPerPage, "", True, "����¼", CurrentPage, KS.QueryParam("page"))
		   %>
    </div>
	<br>
   <table border="0" width="99%" align="center">
    <tr>
	  <td>
     <font color=red>˵����Ϊ�����𲻱�Ҫ�ľ��ף��ʽ���ϸ���ṩ��ѯ���ܣ�����ɾ��������</font>
	     </td>
	</tr>
	</table>
</body>
</html>
   <%
   End Sub
  
  Sub ShowContent()
     on error resume next
     Dim I,intotalmoney,outtotalmoney
     Do While Not rs.eof 
	%>
    <tr class=list onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
      <td class="splittd" align=middle width=120><%=rs("LogTime")%></td>
      <td class="splittd" align=middle width=80><%=rs("username")%></td>
	  <td class="splittd" align=middle width=80><%=rs("clientname")%></td>
      <td class="splittd" align=middle width=60>
	  <% Select Case rs("MoneyType")
	      Case 1:Response.WRite "�ֽ�"
		  Case 2:Response.Write "���л��"
		  Case 3:Response.Write "����֧��"
		  Case 4:Response.Write "�ʽ����"
		 End Select
	 %>
	  </td>
      <td class="splittd" align=middle width=50>�����</td>
      <td class="splittd" width=80 align=right> &nbsp;
	  <%If rs("IncomeOrPayOut")=1 Then
	     Response.Write formatnumber(rs("money"),2)
		 intotalmoney=intotalmoney+rs("money")
	    End If
		%></td>
      <td class="splittd" align=right width=80>&nbsp;
	  <%If rs("IncomeOrPayOut")=2 Then
	     Response.Write formatnumber(rs("money"),2)
		 outtotalmoney=outtotalmoney+rs("money")
	    End If
		%></td>
      <td class="splittd" align=center width=40>
	  <% If rs("IncomeOrPayOut")=1 Then
	      Response.Write "<font color=red>����</font>"
		 Else
		  Response.Write "<font color=green>֧��</font>"
		 End If
		 %></td>
      <td class="splittd" align=middle><%=formatnumber(rs("currmoney"),2)%></td>
      <td class="splittd" align=middle><%=rs("Remark")%></td>
    </tr>
	<%
	            
				I = I + 1
				RS.MoveNext
				If I >= MaxPerPage Then Exit Do

	 loop
	%>
    <tr class=list onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
      <td class="splittd" align=right colSpan=5>��ҳ�ϼƣ�</td>
      <td class="splittd" align=right><%=formatnumber(intotalmoney,2)%></td>
      <td class="splittd" align=right><%=formatnumber(outtotalmoney,2)%></td>
      <td class="splittd" colSpan=3>&nbsp;</td>
    </tr>
    <tr class=list onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
      <td class="splittd" align=right colSpan=5>�ܼƽ�</td>
	  <%intotalmoney=Conn.execute("Select Sum(Money) From KS_Logmoney Where "& SqlParam & " And IncomeOrPayOut=1")(0)
	    outtotalmoney=Conn.execute("Select Sum(Money) From KS_Logmoney Where "& SqlParam & " And IncomeOrPayOut=2")(0)
	    if not isnumeric(intotalmoney) then intotalmoney=0
		if not isnumeric(outtotalmoney) then outtotalmoney=0
	  %>
      <td class="splittd" align=right><%=formatnumber(intotalmoney,2)%></td>
      <td class="splittd" align=right><%=formatnumber(outtotalmoney,2)%></td>
      <td class="splittd" align=middle colSpan=3>�ʽ���<%=formatnumber(intotalmoney-outtotalmoney,2)%></td>
    </tr>
  </table>
		<%
		End Sub
End Class
%> 
