<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../KS_Cls/Kesion.UpFileCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_MallScoreOrder
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_MallScoreOrder
        Private KS,Param,KSCls
		Private Action,i,strClass,RS,SQL,maxperpage,CurrentPage,totalPut,TotalPageNum
        Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()
		 With Response
					If Not KS.ReturnPowerResult(0, "KSMS20010") Then          '�����Ȩ��
					 Call KS.ReturnErr(1, "")
					 .End
					 End If
			  .Write "<html>"
			  .Write"<head>"
			  .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			  .Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .Write "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  .Write "<ul id='menu_top'>"
			  .Write "<li class='parent' onclick=""window.parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape('���ֶһ�ϵͳ >> <font color=red>�����Ʒ</font>')+'&ButtonSymbol=GOSave';location.href='KS.MallScore.asp?action=Add';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>�����Ʒ</span></li>"
			  .Write "<li class='parent' onclick=""location.href='KS.MallScoreOrder.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/move.gif' border='0' align='absmiddle'>����һ�����</span></li>"
			  .Write "<li style='margin-left:30px;margin-top:10px'><strong>�鿴��ʽ:</strong><a href=""KS.MallScoreOrder.asp"">���ж���</a> <a href=""KS.MallScoreOrder.asp?flag=1"">�����</a>  <a href=""KS.MallScoreOrder.asp?flag=-1"">δ���</a> <a href=""KS.MallScoreOrder.asp?flag=2"">�����</a> <a href=""KS.MallScoreOrder.asp?flag=3"">�ѷ���</a> <a href=""KS.MallScoreOrder.asp?flag=4"">�����</a></li>"

			  .Write "</ul>"
		End With
		
		
		maxperpage = 30 '###ÿҳ��ʾ��
		If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
			Response.Write ("�����ϵͳ����!����������")
			Response.End
		End If
		If Not IsEmpty(Request("page")) And Len(Request("page")) <> 0 Then
			CurrentPage = CInt(Request("page"))
		Else
			CurrentPage = 1
		End If
		If CInt(CurrentPage) = 0 Then CurrentPage = 1
		
		Param=" where 1=1"
		If KS.G("KeyWord")<>"" Then
		  If KS.G("condition")=1 Then
		   Param= Param & " and b.ProductName like '%" & KS.G("KeyWord") & "%'"
		  ElseIf KS.G("condition")=2 Then
		   Param= Param & " and a.OrderID like '%" & KS.G("KeyWord") & "%'"
		  Else
		   Param= Param & " and a.RealName like '%" & KS.G("KeyWord") & "%'"
		  End If
		End If
		If KS.G("Flag")<>"" Then
		  If KS.G("Flag")="-1" Then 
		    Param=Param & " and a.Status=0"
		  Else
		   Param=Param & " and a.Status=" & KS.ChkClng(KS.G("Flag"))
		  End If
		End If
		If KS.S("ProductID")<>"" Then Param=Param & " and a.productid=" & KS.ChkClng(KS.G("ProductID"))

		totalPut = Conn.Execute("Select Count(id) From KS_MallScoreOrder a " & Param)(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '�õ���ҳ��
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		Select Case KS.G("action")
		 Case "Add","Edit" Call ProductNameManage()
		 Case "EditSave" Call DoSave()
		 Case "Del"  Call OrderDel()
		 Case Else
		  Call showmain
		End Select
End Sub

Private Sub showmain()
If KS.S("ProductID")<>"" Then
 %>
  <div style="height:45px;font-size:14px;line-height:45px;text-align:center;font-weight:bold">�鿴��Ʒ <font color=red>[<%=LFCls.GetSingleFieldValue("Select ProductName From KS_MallScore Where ID=" & KS.ChkClng(KS.G("ProductID")))%>]</font> �Ķһ���¼</div>
 <%end If%>
<table width="100%" border="0" align="center" style="border-top:1px solid #cccccc" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>ѡ��</th>
	<td nowrap>������</th>
	<td nowrap>��Ʒ����</th>
	<td nowrap>�һ���</th>
	<td nowrap>�һ�ʱ��</th>
	<td nowrap>�һ�����</th>
	<td nowrap>�ͻ���ʽ</th>
	<td nowrap>����״̬</th>
	<td nowrap>�������</th>
</tr>
<%
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select A.*,B.ProductName from KS_MallScoreOrder A left join KS_MallScore b on a.productid=b.id " & Param & " order by a.id desc"
	If DataBaseType = 1 Then
		If CurrentPage > 100 Then
			Rs.Open SQL, Conn, 1, 1
		Else
			Set Rs = Conn.Execute(SQL)
		End If
	Else
		Rs.Open SQL, Conn, 1, 1
	End If
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td height=""25"" align=center bgcolor=""#ffffff"" colspan=10>�Բ���,�Ҳ���������</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<form name=selform method=post action=?action=Del>
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		
%>
<tr height="25" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td align="center" class="splittd"><input type=checkbox name=ID value='<%=rs("id")%>'></td>
	<td  class="splittd"><font color=green><%=Rs("orderid")%></font></td>
	<td  class="splittd"><%if Rs("ProductName")="" or isnull(Rs("ProductName")) Then Response.write "<font color=#999999>��ɾ��</font>" Else Response.Write RS("ProductName")%></td>
	<td align="center" class="splittd"><%=Rs("username")%></td>
	<td align="center" class="splittd"><%=Rs("AddDate")%></td>
	<td align="center" class="splittd"><font color=#cccccc><%=RS("amount")%>  ��</font></td>
	<td align="center" class="splittd">
	<%
	 if rs("DeliveryType")=1 then
	  response.write "��ݵ���"
	 else
	  response.write "��ȡ"
	 end if
	%>
	</td>
	
	<td align="center" class="splittd"><%
		select case  rs("status")
		 case 1
		  response.write "����"
		 case 2
		  response.write "<font color=blue>�����</font>"
		 case 3
		  response.write "<font color=#ff6600>�ѷ���</font>"
		 case 4
		  response.write "<font color=#999999>�������</font>"
		 case 5
		  response.write "<font color=green>��Ч����(�����˻�)</font>"
		 case else
		  response.write " <font color=red>δ��</font>"
		end select
	%></td>
	<td align="center" class="splittd"><a href="?action=Edit&ID=<%=RS("ID")%>"  onclick="window.parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape('���ֶһ�ϵͳ >> <font color=red>�޸��Ź���Ϣ</font>')+'&ButtonSymbol=GOSave';">�鿴/�޸�</a> <a href="?Action=Del&ID=<%=rs("id")%>" onclick="return(confirm('�˲���������,ȷ��ɾ���ö�����'));">ɾ��</a> 
		

	</td>
</tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr>
	<td class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" height='25' colspan=7>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">ȫѡ
	<input class=Button type="submit" name="Submit2" value=" ɾ��ѡ�еĶ��� " onclick="{if(confirm('�˲��������棬ȷ��Ҫɾ��ѡ�еļ�¼��?')){this.document.selform.submit();return true;}return false;}"></td>
</tr>
</form>
<tr>
	<td  colspan=7 align=right>
	<%
	 Call KSCLS.ShowPage(totalPut, MaxPerPage, "KS.MallScoreOrder.asp", True, "��", CurrentPage, KS.QueryParam("page"))
	%></td>
</tr>
</table>
<div>
<form action="KS.MallScoreOrder.asp" name="myform" method="get">
   <div style="border:1px dashed #cccccc;margin:3px;padding:4px">
      &nbsp;<strong>��������=></strong>
	 &nbsp;�ؼ���:<input type="text" class='textbox' name="keyword">&nbsp;����:
	 <select name="condition">
	  <option value=1>����Ʒ����</option>
	  <option value=2>��������</option>
	  <option value=3>���ջ���</option>
	 </select>
	  &nbsp;<input type="submit" value="��ʼ����" class="button" name="s1">
    </div>
</form>
</div>
<%
End Sub

Sub ProductNameManage()
Dim ProductName,ActiveDate,AddDate,DeliveryType,Amount,Score,Telphone,RealName,ZipCode,Protection,BuyFlow,Notes,Tel,Status,Address,Email,Remark,UserName
Dim ID:ID=KS.ChkClng(KS.G("ID"))
Dim RS:Set RS=server.createobject("adodb.recordset")
If KS.G("Action")="Edit" Then
	RS.Open "Select a.*,b.productname,score From KS_MallScoreOrder a Left Join KS_MallScore b on a.productid=b.id Where a.ID=" & ID,conn,1,1
	 If RS.Eof And RS.Bof Then
	  RS.Close:Set RS=Nothing
	  Response.Write "<script>alert('�������ݳ���');history.back();</script>"
	  Response.End
	 Else
	   ProductName=RS("ProductName")
	   If KS.IsNul(ProductName) Then ProductName="��ɾ��"
	   AddDate=RS("AddDate")
	   DeliveryType=RS("DeliveryType")
	   Amount=RS("Amount")
	   Score=RS("Score")
	   RealName=RS("RealName")
	   Address=RS("Address")
	   ZipCode=RS("ZipCode")
	   Tel=RS("Tel")
	   Email=RS("Email")
	   Remark=RS("Remark")
	   UserName=RS("UserName")
	   Status=RS("Status")
	 End If
Else
  AddDate=Now
  DeliveryType=Now+30
  ZipCode=0:Score=10
  Tel=0:Status=1
  Amount=100
  RealName=" "
  Address="../images/nophoto.gif"
 End If
%>
<script>
function CheckForm()
{
	if ($F('RealName')=='')
	{
	 alert('�������ջ���!');
	 $Foc("RealName");
	 return false;
	}

document.myform.submit();
}
</script>
<br>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
  <form name="myform" action="?action=EditSave" method="post">
    <input type="hidden" value="<%=Score%>" name="Score"/>
    <input type="hidden" value="<%=ID%>" name="id" />
    <input type="hidden" value="<%=request.servervariables("http_referer")%>" name="comeurl" />
       <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='170' height='30' align='right' class='clefttitle'><strong>�޸Ķ���״̬��</strong></td>
      <td height='30'>&nbsp;
	  <%if Status<>5 then%>
          <input type="radio" name="Status" value="0"<%if Status=0 then response.write " checked"%> />
        δͨ�����
        <input type="radio" name="Status" value="1"<%if Status=1 then response.write " checked"%> />
        ͨ�����
        <input type="radio" name="Status" value="2"<%if Status=2 then response.write " checked"%> />
        �����
        <input type="radio" name="Status" value="3"<%if Status=3 then response.write " checked"%> />
        �ѷ���
        <input type="radio" name="Status" value="4"<%if Status=4 then response.write " checked"%> />
        �������
        <label style="color:green"><input type="radio" name="Status" onclick="alert('ע��:һ��ȷ�����óɴ�״̬��,�����������ó�����״̬!')" value="5"<%if Status=5 then response.write " checked"%> />
        ��Ч�������˻ػ���</label>
	<%else%>
	   <input type="hidden" name="status" value="-1">
	   <label style="color:green">��Ч�������˻ػ���</label>
	 <%end if%>

		
				</td>
    </tr>
 <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='170' height='30' align='right' class='clefttitle'><strong>�һ���Ʒ��</strong></td>
      <td width="781" height='30'>&nbsp;
          <%=ProductName%></td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td height='30' align='right' class='clefttitle'><strong>�һ��û���</strong></td>
      <td height='30'>&nbsp;<%=username%></td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='170' height='30' align='right' class='clefttitle'><strong>�һ�ʱ�䣺</strong></td>
      <td height='30'>&nbsp;<%=adddate%></td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='170' height='30' align='right' class='clefttitle'><strong>�һ�������</strong></td>
      <td height='30'>&nbsp;<%=amount%></td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td height='30' align='right' class='clefttitle'><strong>���ͷ�ʽ��</strong></td>
      <td height='30'>&nbsp;
<input type="radio" name="DeliveryType" value="1"<%if DeliveryType=1 then response.write " checked"%> />
��ݵ���
  <input type="radio" name="DeliveryType" value="2"<%if Status=2 then response.write " checked"%> />
��ȡ </td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='170' height='30' align='right' class='clefttitle'><strong>�� �� �ˣ�</strong></td>
      <td height='30'>&nbsp;
          <input type='text' name='RealName' value='<%=RealName%>' size="20" />
          <font color=red>*</font></td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='170' height='30' align='right' class='clefttitle'><strong>��ϵ�绰��</strong></td>
      <td height='30'>&nbsp;
        <input type='text' name='Tel' value='<%=Tel%>' size="20" />
        <font color="red">*</font></td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='170' height='30' align='right' class='clefttitle'><strong>�ջ���ַ��</strong></td>
      <td height='30'>&nbsp;
          <input type='text' name='Address' value='<%=Address%>' size="35" />        
        <font color=red>*</font></td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='170' height='30' align='right' class='clefttitle'><strong>�������룺</strong></td>
      <td height='30'>&nbsp;
          <input name='ZipCode' value="<%=ZipCode%>" size="10" /></td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='170' height='30' align='right' class='clefttitle'><strong>�������䣺</strong></td>
      <td height='30'>&nbsp;
        <input type='text' name='Email' value='<%=Email%>' size="25" /></td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td height='30' align='right' class='clefttitle'><strong>��ע˵����</strong></td>
      <td height='30'>&nbsp;
        <textarea name='Remark' style="width:400px;height:80px"><%=Remark%></textarea></td>
    </tr>
  </form>
</table>
<%
End Sub

Sub DoSave()
       Dim ID:ID=KS.ChkClng(KS.G("id"))
	   Dim Address:Address=KS.G("Address")
	   Dim RealName:RealName=KS.G("RealName")
	   Dim ZipCode:ZipCode=KS.G("ZipCode")
	   Dim Tel:Tel=KS.G("Tel")
	   Dim Status:Status=KS.ChkClng(KS.G("Status"))
	   Dim ComeUrl:ComeUrl=KS.G("ComeUrl")
	   Dim Remark:Remark=KS.G("Remark")
	   Dim Email:Email=KS.G("Email")
	   Dim DeliveryType:DeliveryType=KS.ChkClng(KS.G("DeliveryType"))
	   
	   If RealName="" Then Response.Write "<script>alert('�ջ��˱�������');history.back();</script>":response.end

            Dim RS: Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select top 1 * From KS_MallScoreOrder Where ID=" & ID,Conn,1,3
				 RS("DeliveryType")=DeliveryType
				 RS("RealName")=RealName
				 RS("Address")=Address
				 RS("ZipCode")=ZipCode
				 RS("Tel")=Tel
				 RS("Remark")=Remark
				 RS("Email")=Email
				 IF (Status<>-1) then
				 RS("Status")=Status
				 end if
		 		 RS.Update
				 RS.MoveLast
				if Status=5 then
				   '�����û�����
				   Call KS.ScoreInOrOut(RS("UserName"),1,KS.ChkClng(KS.G("Score"))*RS("Amount"),"ϵͳ","���ضһ�������<font color=red>" & RS("OrderID") & "</font>����Ʒ����!",0,0)
				end if
				 
				 
			     RS.Close
				 Set RS=Nothing
				 
  Response.Write "<script>alert('�һ������޸ĳɹ���');parent.frames['BottomFrame'].location.href='KS.Split.asp?ButtonSymbol=Disabled&OpStr=" & Server.URLEncode("���ֶһ�ϵͳ >> <font color=red>��������</font>") & "';location.href='"& ComeUrl & "';</script>"

EnD Sub

'ɾ��
Sub OrderDel()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 Conn.execute("Delete From KS_MallScoreOrder Where id In("& id & ")")
 Response.Write "<script>alert('ɾ���ɹ���');location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub



End Class
%> 
