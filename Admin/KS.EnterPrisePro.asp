<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
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
Set KSCls = New Admin_EnterPrisePro
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_EnterPrisePro
        Private KS
		Private Action,i,strClass,sFileName,RS,SQL,maxperpage,CurrentPage,totalPut,TotalPageNum
        Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()
		 With Response
					If Not KS.ReturnPowerResult(0, "KSMS10001") Then          '�����Ȩ��
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
			  If KS.G("Action")<>"View" then
			   .Write "<div class='topdashed sort'>��ҵ��Ʒ����</div>"
			 End If
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
		totalPut = Conn.Execute("Select Count(id) From KS_Product")(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '�õ���ҳ��
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		Select Case KS.G("action")
		 Case "Edit" Call ProEdit()
		 Case "EditSave" Call DoSave()
		 Case "Del" Call ProDel()
		 Case "verific"  Call Verify()
		 Case "unverific"  Call UnVerify()
		 Case "View" Call ShowNews()
		 Case Else
		  Call showmain
		End Select
End Sub

Private Sub showmain()
%>

<script src="../ks_inc/kesion.box.js"></script>
<script>
function ShowIframe(id)
{
    PopupCenterIframe("�鿴��Ʒ","KS.EnterPrisePro.asp?action=View&ProID="+id,600,350,"auto")
}
</script>

<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>ѡ��</th>
	<td nowrap>��Ʒ����</th>
	<td nowrap>���</th>
	<td nowrap>��Ʒ�ͺ�</th>
	<td nowrap>��Ʒ�۸�</th>
	<td nowrap>����</th>
	<td nowrap>״̬</th>
	<td nowrap>�������</th>
</tr>
<%
	sFileName = "KS.EnterprisePro.asp?"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_Product order by id desc"
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
		Response.Write "<tr><td height=""25"" align=center bgcolor=""#ffffff"" colspan=7>û����ҵ��Ʒ��</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<form name=selform method=post action="?">
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		
%>
<tr height="25" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td class="splittd" align="center"><input type=checkbox name=ID value='<%=rs("id")%>'></td>
	<td class="splittd"><a href="#" onclick="ShowIframe(<%=rs("id")%>)"><%=Rs("Title")%></a></td>
	<td class="splittd" align="center"><a href='../space/?<%=rs("inputer")%>' target='_blank'><%=Rs("inputer")%></a></td>
	<td class="splittd" align="center">&nbsp;<%=Rs("ProModel")%>&nbsp;</td>
	<td class="splittd" align="center"><%=Rs("Price")%> Ԫ</td>
	<td class="splittd" align="center">
	 &nbsp;<% 
	 if rs("recommend")="1" then
	  response.write "<font color=blue>��</font> "
	 end if
	 if rs("popular")="1" then
	  response.write "<font color=#ff6600>��</font> "
	 end if
	 if rs("istop")="1" then
	  response.write "��"
	 end if
	 
	 %>
	</td>
	<td class="splittd" align="center"><%
	select case rs("verific")
	 case 0
	  response.write "<font color=red>δ��</font>"
	 case 1
	  response.write "<font color=#999999>����</font>"
	 case 2
	  response.write "<font color=blue>����</font>"
	end select
	%></td>
	<td class="splittd" align="center"><a href="#" onclick="ShowIframe(<%=rs("id")%>)">���</a> <a href="?Action=Del&ID=<%=rs("id")%>" onclick="return(confirm('ȷ��ɾ����'));">ɾ��</a> <a href="?Action=verific&id=<%=rs("id")%>">���</a></td>
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
	<input class=Button type="submit" name="Submit2" value=" ɾ��ѡ�еļ�¼" onclick="{if(confirm('�˲��������棬ȷ��Ҫɾ��ѡ�еļ�¼��?')){this.form.Action.value='Del';this.form.submit();return true;}return false;}">
	<input type="button" value="�������" class="button" onclick="this.form.Action.value='verific';this.form.submit();">
	<input type="button" value="����ȡ�����" class="button" onclick="this.form.Action.value='unverific';this.form.submit();">
	<input type="hidden" value="Del" name="Action">
	</td>
</tr>
</form>
<tr>
	<td colspan=10>
	<%
 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)	%></td>
</tr>
</table>

<%
End Sub

Sub ProEdit()
 Dim ID:ID=KS.ChkCLng(KS.G("id"))
 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
 RS.Open "Select * From KS_Product Where ID=" & ID,conn,1,1
 If RS.Eof And RS.Bof Then
  RS.Close:Set RS=Nothing
  Response.Write "<script>alert('�������ݳ���');history.back();</script>"
  Response.End
 End If
%>
<script>
function CheckForm()
{
if (document.myform.productname=='')
{
 alert('�������Ʒ����');
 document.myform.productname.focus();
 return false;
}
document.myform.submit();
}
</script>
<br>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
 <form name="myform" action="?action=EditSave" method="post">
   <input type="hidden" value="<%=rs("id")%>" name="id">
   <input type="hidden" value="<%=request.servervariables("http_referer")%>" name="comeurl">
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>��Ʒ���ƣ�</strong></td>
           <td height='28'>&nbsp;<input type='text' name='productname' value='<%=RS("productname")%>' size="40"> <font color=red>*</font></td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>������ҵ��</strong></td>
            <td height='28'>&nbsp;<%
		dim rss,sqls,count
		set rss=server.createobject("adodb.recordset")
		sqls = "select * from KS_enterpriseClass Where parentid<>0 order by orderid"
		rss.open sqls,conn,1,1
		%>
          <script language = "JavaScript">
		var onecount;
		subcat = new Array();
				<%
				count = 0
				do while not rss.eof 
				%>
		subcat[<%=count%>] = new Array("<%= trim(rss("id"))%>","<%=trim(rss("parentid"))%>","<%= trim(rss("classname"))%>");
				<%
				count = count + 1
				rss.movenext
				loop
				rss.close
				%>
		onecount=<%=count%>;
		function changelocation(locationid)
			{
			document.myform.SmallClassID.length = 0; 
			var locationid=locationid;
			var i;
			for (i=0;i < onecount; i++)
				{
					if (subcat[i][1] == locationid)
					{ 
						document.myform.SmallClassID.options[document.myform.SmallClassID.length] = new Option(subcat[i][2], subcat[i][0]);
					}        
				}
			}    
		
		</script>
		 <select class="face" name="BigClassID" onChange="changelocation(document.myform.BigClassID.options[document.myform.BigClassID.selectedIndex].value)" size="1">
		<% 
		dim rsb,sqlb
		set rsb=server.createobject("adodb.recordset")
        sqlb = "select * from ks_enterpriseClass where parentid=0 order by orderid"
        rsb.open sqlb,conn,1,1
		if rsb.eof and rsb.bof then
		else
		    do while not rsb.eof
					  If rs("BigClassID")=rsb("id") then
					  %>
                    <option value="<%=trim(rsb("id"))%>" selected><%=trim(rsb("ClassName"))%></option>
                    <%else%>
                    <option value="<%=trim(rsb("id"))%>"><%=trim(rsb("ClassName"))%></option>
                    <%end if
		        rsb.movenext
    	    loop
		end if
        rsb.close
			%>
                  </select>
                  <font color=#ff6600>&nbsp;*</font>
                  <select class="face" name="SmallClassID">
				   <option value='0'>--��ѡ��-</option>
                    <%dim rsss,sqlss
						set rsss=server.createobject("adodb.recordset")
						sqlss="select * from ks_enterpriseclass where parentid="& rs("BigClassID")&" order by orderid"
						rsss.open sqlss,conn,1,1
						if not(rsss.eof and rsss.bof) then
						do while not rsss.eof
							  if rs("SmallClassID")=rsss("id") then%>
							<option value="<%=rsss("id")%>" selected><%=rsss("ClassName")%></option>
							<%else%>
							<option value="<%=rsss("id")%>"><%=rsss("ClassName")%></option>
							<%end if
							rsss.movenext
						loop
					end if
					rsss.close
					%>
                </select></td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>��Ʒ�۸�</strong></td>
           <td height='28'>&nbsp;<input type='text' name='price' size='8' value='<%=RS("price")%>'> Ԫ&nbsp;&nbsp;&nbsp;&nbsp;<strong>��Ʒ���ԣ�</strong><input type='checkbox' name='recommend' value='1'<%if rs("recommend")="1" then response.write " checked"%>>�Ƽ� <input type='checkbox' name='popular' value='1'<%if rs("popular")="1" then response.write " checked"%>>���� <input type='checkbox' name='istop' value='1'<%if rs("istop")="1" then response.write " checked"%>>�̶�</td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>��Ʒ���أ�</strong></td>
           <td height='28'>&nbsp;<input type='text' name='address' value='<%=RS("address")%>'></td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>��Ʒ�ͺţ�</strong></td>
           <td height='28'>&nbsp;<input type='text' name='promodel' value='<%=RS("promodel")%>'></td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>��ϸ���ܣ�</strong></td>
           <td height='28'> <%
		     Response.Write "<textarea id=""Intro"" name=""Intro"" style=""display:none"">"& KS.HTMLCode(rs("Intro")) &"</textarea><input type=""hidden"" id=""Intro___Config"" value="""" style=""display:none"" /><iframe id=""Intro___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=Intro&amp;Toolbar=NewsTool"" width=""695"" height=""360"" frameborder=""0"" scrolling=""no""></iframe>"
								%>		
		   </td>
          </tr> 
		  <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>��ƷͼƬ��</strong></td>
           <td height='28'>&nbsp;<input type='text' name='photourl' size="45" value='<%=RS("photourl")%>'></td>
          </tr>  
		  <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>�� �� �ˣ�</strong></td>
           <td height='28'>&nbsp;<input type='text' name='username' size="45" value='<%=RS("username")%>'></td>
          </tr>  
 
		 </form>  
</table>
<%
RS.Close:Set RS=Nothing
End Sub

Sub DoSave
	Dim ID:ID=KS.ChkCLng(KS.G("id"))
	Dim RS:Set RS=Server.CreateOBject("ADODB.RECORDSET")
	RS.Open "Select * From KS_Product Where Id=" & ID,conn,1,3
	If RS.Eof And RS.Bof Then
	 RS.Close:Set RS=Nothing
	 Response.Write "<script>alert('�������ݳ���');history.back();</script>"
	 Response.End
	End If
	RS("ProductName")=KS.G("ProductName")
	RS("BigClassID")=KS.ChkCLng(KS.G("BigClassID"))
	RS("SmallClassID")=KS.ChkCLng(KS.G("SmallClassID"))
	RS("Price")=KS.ChkClng(KS.G("Price"))
	RS("Address")=KS.G("Address")
	RS("Intro")=KS.HtmlEncode(Request.Form("Intro"))
	RS("PhotoUrl")=KS.G("PhotoUrl")
	RS("UserName")=KS.G("UserName")
	RS("Recommend")=KS.ChkClng(KS.G("Recommend"))
	RS("Popular")=KS.ChkCLng(KS.G("Popular"))
	RS("Istop")=KS.ChkClng(KS.G("Istop"))
	RS.Update
	RS.Close:Set RS=Nothing
	Response.Write "<script>alert('��ϲ����Ʒ�޸ĳɹ���');location.href='" & Request.Form("ComeUrl") & "';</script>"
End Sub

'ɾ��
Sub ProDel()
 on error resume next
 Dim I,ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 ID=Split(ID,",")
 For I=0 To Ubound(ID)
  KS.DeleteFile(conn.execute("select photourl from KS_Product where id=" & ID(I))(0))
  Conn.execute("Delete From KS_Product Where id="& id(I))
 Next 
 Response.Write "<script>alert('ɾ���ɹ���');location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub

'���
Sub ShowNews()
	    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select * From KS_Product where id=" &KS.ChkClng(KS.S("ProID")),conn,1,1
		If Not RS.Eof Then
		   Response.WRITE "<div style='padding:30px'><div><strong>��Ʒ���ƣ�</strong>" & rs("title") & "</div>"
		   Response.Write "<div style=""text-align:left""><strong>��Ʒ�ͺţ�</strong>" & RS("ProModel") & "</div>"
		   Response.Write "<div style=""text-align:left""><strong>��Ʒ���</strong>" & RS("ProSpecificat") & "</div>"
		   Response.Write "<div style=""text-align:left""><strong>��Ʒ�۸�</strong>" & RS("price") & " Ԫ</div>"
		   Response.Write "<div style=""text-aling:left""><strong>��Ʒ���ԣ�</strong>"
			 if rs("recommend")="1" then
			  response.write "<font color=blue>��</font> "
			 end if
			 if rs("popular")="1" then
			  response.write "<font color=#ff6600>��</font> "
			 end if
			 if rs("istop")="1" then
			  response.write "��"
			 end if
	      response.write "</div>"
	  
		   Dim PhotoUrl:PhotoUrl=RS("PhotoUrl")
		   If PhotoUrl<>"" And Not IsNull(PhotoURL) Then
		   Response.Write "<div style=""text-align:left"">��ƷͼƬ��<img src='" & RS("photourl") & "'></div>"
		   End If
		   Response.Write "<div>��Ʒ���ܣ�" & KS.HTMLCode(rs("prointro")) & "</div>"
		   Response.Write "</div>"
		End If
		RS.Close:Set RS=Nothing
End Sub
'���
Sub Verify
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 Conn.execute("Update KS_Product Set verific=1 Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'ȡ�����
Sub UnVerify
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 Conn.execute("Update KS_Product Set verific=0 Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub

End Class
%> 
