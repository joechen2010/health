<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New PaymentPlatCls
KSCls.Kesion()
Set KSCls = Nothing

Class PaymentPlatCls
        Private KS,Action,KSCls
		Private K, SqlStr,ChannelID,SQL,RS
		
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls= New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub


		Public Sub Kesion()
             With KS
		 	     .echo "<html>"
				 .echo "<head>"
				 .echo "<meta http-equiv='Content-Type' content='text/html; chaRSet=gb2312'>"
				 .echo "<title>֧��ƽ̨����</title>"
				 .echo "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		         .echo "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>" & vbCrLf
	         	 .echo "<script language=""JavaScript"" src=""../KS_Inc/jQuery.js""></script>" & vbCrLf
					If Not KS.ReturnPowerResult(0, "KMST10001") Then          '����Ƿ��л�����Ϣ���õ�Ȩ��
					  .echo ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back()';</script>")
					 Call KS.ReturnErr(1, "")
					 .End
					 End If


			  Action=KS.G("Action")
			 Select Case Action
			  Case "Modify"
			    Call DoModify()
			  Case "DoModifySave"
			    Call DoModifySave()
			  Case "DoBatch"
			    Call DoBatch()
			  Case "Disabled"
			    Call DoDisabled()
			  Case Else
			   Call MainList()
			 End Select
			 .echo "</body>"
			 .echo "</html>"
			End With
		End Sub
		
		Sub MainList()
		With KS
		 .echo "</head>"
		
		 .echo "<body scroll=no topmargin='0' leftmargin='0'>"
		 .echo "<ul id='mt'> <div id='mtl'>������ʾ��</div><li>"
		 .echo "��ϵͳ���ɶ������֧���ӿڣ��������ڴ˹������е�֧��ƽ̨ "
		 .echo "</ul>"
		 .echo "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
		 .echo(" <form name=""myform"" method=""Post"" action=""?Action=DoBatch"">")
		 .echo "    <tr class='sort'>"
		 .echo "    <td width='30' width='50' align='center'>���</td>"
		 .echo "    <td align='center'>֧��ƽ̨</td>"
		 .echo "    <td width='100' align='center'>�̼�ID</td>"
		 .echo "    <td width='300' align='center'>��ע˵��</td>"
		 .echo "    <td width='60' align='center'>������</td>"
		 .echo "    <td width='40' align='center'>Ĭ��</td>"
		 .echo "    <td width='40' align='center'>����</td>"
		 .echo "    <td width='100' align='center'>�������</td>"
		 .echo "    <td width='60' align='center'>����</td>"
		 .echo "  </tr>"
		 Set RS = Server.CreateObject("ADODB.RecordSet")
         SqlStr = "SELECT ID,OrderID,PlatName,AccountID,Note,MD5Key,Rate,RateByUser,IsDisabled,IsDefault FROM [KS_PaymentPlat] order by orderid"
		 RS.Open SqlStr, conn, 1, 1
		 If Not RS.EOF Then SQL=RS.GetRows(-1)
		 If Not IsArray(SQL) Then
			 .echo "<tr><td  class='list' onMouseOver=""this.className='listmouseover'"" onMouseOut=""this.className='list'"" colspan=8 height='25' align='center'>û���κ�֧��ƽ̨!</td></tr>"
		 Else
			 For K=0 To Ubound(SQL,2)
		       .echo "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
			   .echo "<td class='splittd' align='center'><input type='hidden' value='" & SQL(0,K) & "' name='id'><input type='text' name='orderid' value='" &SQL(1,K) & "' style='width:36px;text-align:center'></td>"
			   .echo " <td class='splittd' height='22'><span style='cursor:default;'>"
			  If SQL(0,K)=10 Then
			   .echo "<img src=images/tenpay.gif><div style='text-align:center'>"
			  End If
			   .echo SQL(2,K)
			   .echo "</td>"
			   
			    .echo " <td class='splittd' align='center'>" & SQL(3,K) & "</td>"
			    .echo " <td class='splittd' align='center'>" & SQL(4,K) & "&nbsp;</td>"
			    .echo " <td class='splittd' align='center'>" & SQL(6,K) & "%</td>"
			    .echo " <td class='splittd' align='center'>"
			   If SQL(9,K)=1 Then
			    .echo "<input type='radio' name='IsDefault' value='" & SQL(0,K) & "' checked>"
			   Else
			    .echo "<input type='radio' name='IsDefault' value='" & SQL(0,K) & "'>"
			   End If
			    .echo " </td>"
			    .echo " <td class='splittd' align='center'>" 
			   If SQL(8,K)=1 Then
			    .echo "<input type='checkbox' name='IsDisabled' value='" & SQL(0,K) & "' checked>"
			   Else
			    .echo "<input type='checkbox' name='IsDisabled' value='" & SQL(0,K) & "'>"
			   End If
			    .echo "</td>"
			    .echo " <td class='splittd' align='center'><a href='?Action=Modify&ID=" & SQL(0,K) &"'>�޸�</a> "
			   If SQL(8,K)=1 Then
			   	 .echo " <a href='?V=0&Action=Disabled&id=" & SQL(0,K) & "'>�ر�</a>"
			   Else
			     .echo " <a href='?V=1&Action=Disabled&id=" & SQL(0,K) & "'>����</a>"
			   End If
			    .echo " </td>"
			    .echo "<td class='splittd'>"
			   	Select Case SQL(0,K)
			    Case 10  .echo "<a href='http://union.tenpay.com/mch/mch_register.shtml?sp_suggestuser=1202640601' target='_blank'>�����̻�"
				case 11  .echo "<a href='http://union.tenpay.com/mch/mch_register_1.shtml?sp_suggestuser=1202640601' target='_blank'>�����̻�"
			    Case  1  .echo "<a href='http://merchant3.chinabank.com.cn/register.do' target='_blank'>�����̻�</a>"
			    Case  5  .echo "<a href='http://new.xpay.cn/SignUp/Default.aspx' target='_blank'>�����̻�</a>"
			    Case  6  .echo "<a href='https://www.cncard.net/products/products.asp' target='_blank'>�����̻�</a>"
			    Case  7,9  .echo "<a href='https://www.alipay.com/' target='_blank'>�����̻�</a>"
			    Case  8  .echo "<a href='https://www.99bill.com/website/' target='_blank'>�����̻�</a>"
			    Case  2  .echo "<a href='http://www.ipay.cn/home/index.php' target='_blank'>�����̻�</a>"
			    Case  4  .echo "<a href='http://www.yeepay.com/' target='_blank'>�����̻�</a>"
			    Case  3  .echo "<a href='https://www.ips.com.cn/' target='_blank'>�����̻�</a>"
			   End Select 
                .echo "</td>"
			    .echo "</tr>"
			Next
								 
		 End If
          .echo "<tr>"
		  .echo " <td colspan='8' height='40'>&nbsp;&nbsp;"
		  .echo " <input type='submit' value='������������' class='button'><font color=blue>&nbsp;���ԽС��ǰ̨����Խǰ�棬ֻ���������������õ�֧��ƽ̨��ǰ̨�Ż���ʾ</font>"
		  .echo " </td>"
		  .echo "</tr>"
		  .echo "</form>"
		  .echo "</table>"
		
		End With
		End Sub

		Sub DoModify()
		 Dim ID:ID=KS.ChkClng(KS.G("ID"))
		 Dim RS:Set RS=Server.CreateOBject("ADODB.RECORDSET")
		 RS.Open "Select * From KS_PaymentPlat Where ID=" & ID,conn,1,1
		 If RS.EOf And RS.Bof Then
		 RS.Close:Set RS=Nothing
		  KS.Echo "<script>alert('�������ݴ���!');history.back();</script>"
		  Exit Sub
		 End If
		%>
		<html>
		<head>
		<title>֧��ƽ̨����</title>
		<meta http-equiv=Content-Type content="text/html; chaRSet=gb2312">
		<link href="Include/Admin_Style.CSS" type=text/css rel=stylesheet>
		</head>
		<body leftMargin=0 topMargin=0>
		<script language="javascript">
		 function CheckForm()
		 {
		   this.myform.submit();
		 }
		</script>
		<ul id=menu_top>
		<li class=parent onclick=return(CheckForm())><SPAN class=child onMouseOver="this.parentNode.className='parent_border'" onMouseOut="this.parentNode.className='parent'"><img src="images/ico/save.gif" align=absMiddle border=0>ȷ������</SPAN></li>
		<li class=parent onClick="location.href='?ChannelID=1';"><SPAN class=child onMouseOver="this.parentNode.className='parent_border'" onMouseOut="this.parentNode.className='parent'"><img src="images/ico/back.gif" align=absMiddle border=0>ȡ������</SPAN></li></ul>
		<FORM name=myform onsubmit=return(CheckForm()) action="?action=DoModifySave&ID=<%=rs("ID")%>" method=post>
		  <table class=ctable style=" BORDER-COLLAPSE: collapse" cellSpacing=1 cellPadding=1 width="100%" align=center border=0>
			<tr class=tdbg>
			  <td class=clefttitle noWrap align=right height=25><strong>ƽ̨���ƣ�</strong></td>
			  <td align=right width=21 height=30>
				<Input value="<%=rs("PlatName")%>" Class="textbox" name="PlatName"> </td>
				<tr class=tdbg>
				  <td class=clefttitle noWrap align=right height=25><strong>��ע˵����</strong></td>
				  <td noWrap height=25>
		                <textarea name="Note" cols="70" rows="5"><%=rs("note")%></textarea>
		            </td>
					<tr class=tdbg>
					  <td class=clefttitle align=right><strong>֧����ţ�</strong><br>
������������֧��ƽ̨������̻����</td>
					  <td>
						<Input id="AccountID" class="textbox" name="AccountID" value="<%=rs("AccountID")%>"> </td>
					</tr>
					<tr class=tdbg>
					  <td class=clefttitle align=right height=25><strong>֧����Կ��</strong><br>
������������������֧��ƽ̨�����õ�MD5˽Կ,��������֧��ƽ̨����Ҫ����</td>
					  <td height=25>
						<Input class="textbox" name="MD5Key" value="<%=rs("MD5Key")%>"></td>
					</tr>

					<tr class=tdbg>
					  <td class=clefttitle align=right height=25><strong>�������ʣ�</strong></td>
					  <td noWrap height=25>
						<Input class="textbox" size="6" name="Rate"  value="<%=rs("rate")%>">%
						<br>
						<input type="checkbox" name="RateByUser" value="1"<%if rs("ratebyuser")=1 Then KS.Echo " checked"%>>
						 �������ɸ����˶���֧��
						</td>
					</tr>
					<tr class=tdbg>
					  <td class=clefttitle align=right><strong>�Ƿ�����:</strong></td>
					  <td>
					    <%if rs("isdisabled")=1 Then%>
						<input type="radio" value="0" name="isdisabled">����
						<input type="radio" value="1" name="isdisabled" checked>����
						<%else%>
						<input type="radio" value="0" name="isdisabled" checked>����
						<input type="radio" value="1" name="isdisabled">����
						<%end if%>
						
					  </td>
					</tr>
				  </table>
				  <div style='margin:8px;text-align:center'>
				   <input type='button' onclick='CheckForm()' class='button' value='ȷ������'>&nbsp;
				   <input type='button' class='button' value='ȡ������' onClick="javascript:location.href='KS.PaymentPlat.asp';">
				  </div>
				   </FORM>
				</body>
				</html>
		<%
		 RS.Close:Set RS=Nothing
		End Sub
		
		Sub DoModifySave()
		  Dim ID:ID=KS.ChkClng(KS.G("ID"))
		  Dim PlatName:PlatName=KS.G("PlatName")
		  Dim Note:Note=KS.G("Note")
		  Dim AccountID:AccountID=KS.G("AccountID")
		  Dim MD5Key:MD5Key=KS.G("MD5Key")
		  Dim Rate:Rate=KS.G("Rate")
		  Dim RateByUser:RateByUser=KS.ChkClng(KS.G("RateByUser"))
		  Dim IsDisabled:IsDisabled=KS.ChkClng(KS.G("IsDisabled"))
		  Dim RS:Set RS=Server.CreateOBject("ADODB.RECORDSET")
		  RS.Open "Select * from KS_PaymentPlat where id=" & ID,conn,1,3
		  If Not RS.Eof Then
		    RS("PlatName") = PlatName
			RS("Note")     = Note
			RS("AccountID")= AccountID
			RS("MD5Key")   = MD5Key
			RS("Rate")     = Rate
			RS("RateByUser")=RateByUser
			RS("IsDisabled")= IsDisabled
			RS.Update
		  End If
		  RS.Close:Set RS=Nothing
		  KS.Alert "��ϲ���޸ĳɹ���","KS.PaymentPlat.asp" 
		End Sub
		
		Sub DoBatch()
			Dim ID:ID = KS.G("ID")
			Dim OrderID:OrderID=KS.G("OrderID")
			Dim IsDisabled:IsDisabled=KS.G("IsDisabled")
			Dim IsDefault:IsDefault=KS.G("IsDefault")
			Dim ID_Arr:ID_Arr=Split(ID,",")
			Dim OrderID_Arr:OrderID_Arr=Split(OrderID,",")
		    Dim K
			For K=0 TO Ubound(ID_Arr)
			 Conn.Execute("Update KS_PaymentPlat Set OrderID=" & OrderID_Arr(K) & " where id=" & ID_Arr(K))
			 If KS.FoundInArr(IsDisabled, ID_Arr(K), ",")=true Then
			  Conn.Execute("Update KS_PaymentPlat Set IsDisabled=1 where id=" & ID_Arr(K))
			 Else
			  Conn.Execute("Update KS_PaymentPlat Set IsDisabled=0 where id=" & ID_Arr(K))
			 End If
			 If KS.FoundInArr(IsDefault, ID_Arr(K), ",")=true Then
			  Conn.Execute("Update KS_PaymentPlat Set IsDefault=1 where id=" & ID_Arr(K))
			 Else
			  Conn.Execute("Update KS_PaymentPlat Set IsDefault=0 where id=" & ID_Arr(K))
			 End If
			Next
			KS.Alert "��ϲ���������óɹ���" , "KS.PaymentPlat.asp"
		 End Sub
		Sub DoDisabled()
		  Conn.Execute("Update KS_PaymentPlat Set IsDisabled=" & KS.ChkClng(KS.G("V")) & " where id=" & KS.ChkClng(KS.G("ID")))
		  KS.AlertHintScript "��ϲ,�����ɹ�!"
		End Sub
End Class
%> 
