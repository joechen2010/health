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
Set KSCls = New Admin_LogInvoice
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_LogInvoice
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
		If Not KS.ReturnPowerResult(5, "M510016") Then  Call KS.ReturnErr(1, ""):Exit Sub   
		%>
		<html>
<head><title>����Ʊ��ѯ</title>
<meta http-equiv=Content-Type content="text/html; charset=gb2312">
<link href="include/Admin_Style.css" type=text/css rel=stylesheet>
</head>
<body leftMargin=2 topMargin=0 marginheight="0" marginwidth="0">

  <div class="topdashed" style="padding:4px;">
  <table cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td><strong>��Ʊ��¼��ѯ��</strong></td>
<FORM name=form1 action=KS.LogInvoice.asp method=get>
      <td valign="top">���ٲ�ѯ�� 
<Select onchange=javascript:submit() size=1 name=SearchType class='textbox'> 
  <Option value=0<%If SearchType="0" Then Response.write " selected"%>>���з�Ʊ��¼</Option> 
  <Option value=1<%If SearchType="1" Then Response.write " selected"%>>���10���ڵ��¼�¼</Option> 
  <Option value=2<%If SearchType="2" Then Response.write " selected"%>>���һ���ڵ��¼�¼</Option> 
  <Option value=3<%If SearchType="3" Then Response.write " selected"%>>���е�˰��ͨ��Ʊ</Option> 
  <Option value=4<%If SearchType="4" Then Response.write " selected"%>>���й�˰��ͨ��Ʊ</Option>
  <Option value=4<%If SearchType="5" Then Response.write " selected"%>>������ֵ˰��Ʊ</Option>
      </Select></td></FORM>
<FORM name=form2 action=KS.LogInvoice.asp method=post>
      <td>�߼���ѯ�� 
<Select id=Field name=Field class='textbox'> 
  <Option value=1 selected>�ͻ�����</Option> 
  <Option value=2>�û���</Option> 
  <Option value=3>����Ʊ����</Option> 
  <Option value=4>������</Option> 
  <Option value=5>��Ʊ����</Option> 
  <Option value=6>��Ʊ̧ͷ</Option> 
  <Option value=7>������</Option>
</Select> 
  <Input id=Keyword class='textbox' maxLength=30 name=Keyword> 
  <Input class='button' type=submit value=" �� ѯ " name=Submit2> 
        <Input id=SearchType type=hidden value=6 name=SearchType> </td></FORM>
    </tr>  
 </table>
  </div>


		<%
		 If KS.G("Action")="ShowInvoice" Then
		   Call ShowInvoice()
		 Else
		   Call LogList()
		 End If
		End Sub
		Sub LogList
		SearchType=KS.ChkClng(KS.G("SearchType"))
		%>
  <div style="border:1px #cccccc dashed;overflow:hidden"></div>
  <table width="100%">
    <tr>
      <td align=left>�����ڵ�λ�ã�<a href="KS.LogInvoice.asp">����Ʊ��¼����</a>&nbsp;&gt;&gt;&nbsp;
	  <%Dim SearchTypeStr
	    Dim KeyWord:KeyWord=KS.G("KeyWord")
	  Select Case SearchType
	     Case 0 :SearchTypeStr="���з�Ʊ��¼"
		 Case 1 :SearchTypeStr="���10���ڵ��¼�¼"
		 Case 2 :SearchTypeStr="���һ���ڵ��¼�¼"
		 Case 3 :SearchTypeStr="���е�˰��ͨ��Ʊ"
		 Case 4 :SearchTypeStr="���й�˰��ͨ��Ʊ"
		 Case 5 :SearchTypeStr="������ֵ˰��Ʊ"
		 Case 6 
		    Select Case KS.ChkClng(KS.G("Field"))
			  Case 1:SearchTypeStr="�ͻ���������<font color=red>""" & KeyWord & """</font>"
			  Case 2:SearchTypeStr="�û�������<font color=red>""" & KeyWord & """</font>"
			  Case 3:SearchTypeStr="����Ʊ���ں���<font color=red>""" & KeyWord & """</font>"
			  Case 4:SearchTypeStr="�����˺���<font color=red>""" & KeyWord & """</font>"
			  Case 5:SearchTypeStr="��Ʊ���뺬��<font color=red>""" & KeyWord & """</font>"
			  Case 6:SearchTypeStr="��Ʊ̧ͷ����<font color=red>""" & KeyWord & """</font>"
			  Case 7:SearchTypeStr="�����ź���<font color=red>""" & KeyWord & """</font>"
			End Select
	  End Select
	  Response.Write SearchTypeStr%></td>
    </tr>
  </table>
  <div style="border:1px #cccccc dashed;overflow:hidden"></div>
  <table  cellSpacing=1 cellPadding=2 width="100%" border=0>
    <tr class=sort align=middle>
        <td>��Ʊ̧ͷ</td>
      <td width=80>����</td>
      <td width=60>�ͻ�����</td>
      <td width=60>�û���</td>
      <td width=100>�������</td>
      <td width=80>��Ʊ����</td>
      <td width=80>��Ʊ����</td>
      <td width=60>��Ʊ���</td>
      <td width=60>��Ʊ��</td>
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
					SqlParam=SqlParam &" And datediff(d,InvoiceDate," & SqlNowString & ")<=10"
				   else
					SqlParam=SqlParam &" And datediff('d',InvoiceDate," & SqlNowString & ")<=10"
				   end if
			   Case 2
			   		if DataBaseType=1 then
					SqlParam=SqlParam &" And datediff(d,InvoiceDate," & SqlNowString & ")<=30"
				   else
					SqlParam=SqlParam &" And datediff('d',InvoiceDate," & SqlNowString & ")<=30"
				   end if
			  Case 3 : SqlParam = SqlParam & "And InvoiceType='��˰��ͨ��Ʊ'"
			  Case 4 : SqlParam = SqlParam & "And InvoiceType='��˰��ͨ��Ʊ'"
			  Case 5 : SqlParam = SqlParam & "And InvoiceType='��ֵ˰��Ʊ'"
			  Case 6
			      Select Case KS.ChkClng(KS.G("Field"))
				   Case 1:SqlParam=SqlParam &" And ClientName Like '%" & Keyword & "%'"
				   Case 2:SqlParam=SqlParam &" And UserName Like '%" & Keyword & "%'"
				   Case 3:SqlParam=SqlParam &" And InvoiceDate Like '%" & Keyword & "%'"
				   Case 4:SqlParam=SqlParam &" And HandlerName Like '%" & Keyword & "%'"
				   Case 5:SqlParam=SqlParam &" And InvoiceNum Like '%" & Keyword & "%'"
				   Case 6:SqlParam=SqlParam &" And InvoiceTitle Like '%" & Keyword & "%'"
				   Case 7:SqlParam=SqlParam &" And OrderID Like '%" & Keyword & "%'"
				  End Select
			  End Select
			End If
	Set RS=Server.CreateObject("ADODB.RECORDSET")
	RS.Open "Select * From KS_LogInvoice Where " & SqlParam & " Order By ID Desc",Conn,1,1
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
   %>  </table>

     <div align=right>
         <%
		   	  '��ʾ��ҳ��Ϣ
			  Call KS.ShowPageParamter(totalPut, MaxPerPage, "", True, "����¼", CurrentPage, "SearchType=" & SearchType & "&Field=" & KS.G("Field") & "&KeyWord=" & KeyWord)
		   %>
    </div>
</body>
</html>
   <%
   End Sub
  
  Sub ShowContent()
     Dim I,intotalInvoice,outtotalInvoice
     Do While Not rs.eof 
	%>
	<tr class=list onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
      <td><a href="KS.LogInvoice.asp?Action=ShowInvoice&ID=<%=rs("id")%>"><%=rs("InvoiceTitle")%></a></td>
      <td align=middle height="22" width=80><%=rs("InvoiceDate")%></td>
      <td align=middle width=60><%=rs("ClientName")%></td>
      <td align=middle width=60><%=rs("username")%></td>
      <td align=middle width=100><%=rs("orderid")%></td>
      <td align=middle width=80><%=rs("InvoiceType")%></td>
      <td align=middle width=80><%=rs("Invoicenum")%></td>
      <td align=right width=60><%=rs("MoneyTotal")%></td>
      <td align=middle width=60><%=rs("HandlerName")%></td>
    </tr>
	<tr><td height="1" colspan="10"><div style="border:1px #cccccc dashed;overflow:hidden"></div></td></tr>
	<%
	            
				I = I + 1
				RS.MoveNext
				If I >= MaxPerPage Then Exit Do

	 loop
		End Sub
		
		'�鿴��Ʊ��Ϣ
		Sub ShowInvoice()
		 Dim ID:ID=KS.ChkClng(KS.G("ID"))
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "SELECT * FROM KS_LogInvoice Where ID=" &ID,conn,1,1
		 IF RS.Eof Then
		    rs.close:set rs=nothing
		   Response.Write "<script>alert('��������!');history.back();</script>"
		   exit sub
		 End IF
		 %><br>
  <table class=border cellSpacing=1 cellPadding=2 width="100%" border=0>
    <tr class=title align=middle>
      <td colSpan=2 height=22><B>�� Ʊ �� Ϣ</B></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">�ͻ����ƣ�</td>
      <td><%=rs("ClientName")%></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">������ţ�</td>
      <td><%=rs("orderid")%></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">������</td>
      <td><%=rs("MoneyTotal")%>Ԫ</td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">��Ʊ���ڣ�</td>
      <td><%=rs("invoicedate")%></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">��Ʊ���ͣ�</td>
      <td><%=rs("invoicetype")%></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">��Ʊ���룺</td>
      <td><%=rs("invoicenum")%></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">��Ʊ̧ͷ��</td>
      <td><%=rs("InvoiceTitle")%></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">��Ʊ���ݣ�</td>
      <td><%=rs("invoicecontent")%></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">��Ʊ��</td>
      <td><%=rs("MoneyTotal")%></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">�� Ʊ �ˣ�</td>
      <td><%=rs("HandlerName")%></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">¼ �� Ա��</td>
      <td><%=rs("inputer")%></td>
    </tr>
    <tr class=tdbg>
      <td align=right width="15%">¼��ʱ�䣺</td>
      <td><%=rs("inputtime")%></td>
    </tr>
  </table>
  <br>
  <div align=center><input type='button' value=' �� �� ' onclick='javascript:history.back();' class='button'></div>
		 <%
		 rs.close
		 set rs=nothing
		 
		End Sub
End Class
%> 
