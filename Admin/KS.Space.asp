<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

Dim KSCls
Set KSCls = New Admin_Space
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Space
        Private KS,Param
		Private Action,i,strClass,RS,SQL,maxperpage,CurrentPage,totalPut,TotalPageNum
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
			  .Write "<script src=""../KS_Inc/jquery.js"" language=""JavaScript""></script>"
			  .Write"</head>"
			  .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  .Write "<ul id='menu_top'>"
			  .Write "<li class='parent' onclick=""location.href='KS.Space.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>�ռ����</span></li>"
			  .Write "<li class='parent' onclick=""location.href='KS.SpaceSkin.asp?flag=2';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/move.gif' border='0' align='absmiddle'>ģ�����</span></li>"
			  .Write "<li class='parent' onclick=""location.href='?action=class';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/addjs.gif' border='0' align='absmiddle'>�ռ����</span></li>"
			  .Write "<li></li>"
			  if request("showtype")<>"1" then
			  .Write "<div><select name='classid' onchange=""location.href='?classid='+this.value;"">"
			   Dim RSC:Set RSC=Conn.Execute("Select ClassID,ClassName From KS_BlogClass order by orderid")
			   .Write "<option value=''>---�����ͷ���鿴---</option>"
			   Do While Not RSC.Eof
			    .Write "<option value='" & RSC(0) & "'>" & rsc(1) & "</option>"
				rsc.movenext
			   Loop
			   RSC.Close
			   Set RSC=Nothing
			  .Write "</select></div>"
			  End If
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
		
		
		Select Case KS.G("action")
		 Case "Del"
		  Call BlogDel()
		 Case "lock"
		  Call BlogLock()
		 Case "unlock"
		  Call BlogUnLock()
		 Case "verific"
		  Call Blogverific()
		 Case "recommend"
		  Call Blogrecommend()
		 Case "Cancelrecommend"
		  Call BlogCancelrecommend()
		 Case "modify" Call Modify()
		 case "modifysave" call ModifySave()
		 Case "class" ShowClass
		 Case Else
		  Call showmain
		End Select
End Sub

Private Sub showmain()
    Dim classname
	if request("showtype")="1" then
		Param=" inner join ks_user u on a.username=u.username where u.usertype=0"
    ElseIf KS.S("ClassID")<>"" Then
	   classname="b.classname,"
	   Param=" left join ks_BLOGClass b on a.classid=b.classid where A.classid=" & KS.ChkClng(KS.G("ClassID"))
	Else
		Param=" where 1=1"
	End If
	
		if request("from")<>"" then
		 param=param & " and status=0"
		end if
		  
		If KS.G("KeyWord")<>"" Then
		  If KS.G("condition")=1 Then
		   Param= Param & " and blogname like '%" & KS.G("KeyWord") & "%'"
		  Else
		   Param= Param & " and username like '%" & KS.G("KeyWord") & "%'"
		  End If
		End If

		totalPut = Conn.Execute("Select Count(blogID) from KS_Blog a " & Param)(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '�õ���ҳ��
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
%>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>ѡ��</th>
	<td nowrap>վ������</th>
	<td nowrap>������</th>
	<td nowrap>����ʱ��</th>
	<td nowrap>״̬</th>
	<td nowrap>�������</th>
</tr>
<%
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select " & classname & " a.* from KS_Blog a  "& Param & " order by blogid desc"
	Rs.Open SQL, Conn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr><td height=""25"" align=center bgcolor=""#ffffff"" colspan=7>û���û�������˿ռ䣡</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<form name=selform method=post action="?">
<input type="hidden" name="action" id="action" value="Del">
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		
%>
<tr height="25" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td class="splittd" align="center"><input type=checkbox name=ID value='<%=rs("blogid")%>'></td>
	<td class="splittd">
	<%if request("showtype")="" and request("classid")<>"" then%>
	[<%=RS(0)%>]
	<%end if%>
	<a href="../space/?<%=rs("username")%>" target="_blank"><%=Rs("blogname")%></a></td>
	<td class="splittd" align="center"><%=Rs("username")%></td>
	<td class="splittd" align="center"><%=Rs("adddate")%></td>
	<td class="splittd" align="center"><%
	select case rs("status")
	 case 0
	  response.write "δ��"
	 case 1
	  response.write "<font color=red>����</font>"
	 case 2
	  response.write "<font color=blue>����</font>"
	end select
	%></td>
	<td class="splittd" align="center"><a href="../space/?<%=rs("username")%>" target="_blank">���</a> <a href="?action=modify&id=<%=rs("blogid")%>" onclick="window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape('�ռ��Ż����� >> <font color=red>�޸Ŀռ���Ϣ</font>')+'&ButtonSymbol=GOSave';">�༭</a> <a href="?Action=Del&ID=<%=rs("blogid")%>" onclick="return(confirm('ɾ��վ�㽫ɾ��վ���µ�������־��ȷ��ɾ����'));">ɾ��</a> <%IF rs("recommend")="1" then %><a href="?Action=Cancelrecommend&id=<%=rs("blogid")%>"><font color=red>ȡ���Ƽ�</font></a><%else%><a href="?Action=recommend&id=<%=rs("blogid")%>">��Ϊ�Ƽ�</a><%end if%>&nbsp;<%if rs("status")=0 then%><a href="?Action=verific&id=<%=rs("blogid")%>">���</a> <%elseif rs("status")=1 then%><a href="?Action=lock&id=<%=rs("blogid")%>">����</a><%elseif rs("status")=2 then%><a href="?Action=unlock&id=<%=rs("blogid")%>">����</a><%end if%></td>
</tr>
<%
		Rs.movenext
			i = i + 1
			If i >= maxperpage Then Exit Do
		Loop
	End If
	Rs.Close:Set Rs = Nothing
%>
<tr class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td  class="splittd" height='25' colspan=7>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">ȫѡ
	<input class=Button type="submit" name="Submit2" value=" ɾ��ѡ�еĿռ� " onclick="{if(confirm('�˲��������棬ȷ��Ҫɾ��ѡ�еļ�¼��?')){$('#action').val('Del');this.document.selform.submit();return true;}return false;}">
	<input type="submit" value="�������/����" onclick="$('#action').val('verific');" class="button">
	<input type="submit" value="��������" onclick="$('#action').val('lock');" class="button">
	</td>
</tr>
</form>
<tr>
	<td colspan=10 align=right>
	<%
	 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	%></td>
</tr>
</table>
<div>
<form action="KS.Space.asp" name="myform" method="get">
   <div style="border:1px dashed #cccccc;margin:3px;padding:4px">
      &nbsp;<strong>��������=></strong>
	 &nbsp;�ؼ���:<input type="text" class='textbox' name="keyword">&nbsp;����:
	 <select name="condition">
	  <option value=1>վ������</option>
	  <option value=2>�� �� ��</option>
	 </select>
	  &nbsp;<input type="submit" value="��ʼ����" class="button" name="s1">
	  </div>
</form>
</div>
<%
End Sub

Sub Modify()
 Dim ID:ID=KS.ChkClng(Request("id"))
 If ID=0 Then KS.AlertHintScript "�Բ��𣬷Ƿ�����!"
 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
 RS.Open "select * from ks_blog where blogid=" & id,conn,1,1
 If RS.Eof AND RS.Bof Then
   RS.Close
   Set RS=Nothing
   KS.AlertHintScript "�Բ����Ҳ�����¼��"
 End If
 %>
 <script type="text/javascript">
 function CheckForm()
 {
   if ($("input[name=BlogName]").val()=='')
   {
     alert('�ռ����Ʊ������룡');
	 $("input[name=BlogName]").focus();
	 return false;
   }
   $("#myform").submit();
 }
 </script>
 <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
 <form name="myform" id="myform" action="?action=modifysave" method="post">
   <input type="hidden" value="<%=ID%>" name="id">
   <input type="hidden" value="<%=request.servervariables("http_referer")%>" name="comeurl">
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>�ռ����ƣ�</strong></td>
           <td height='28'>&nbsp;<input type='text' name='BlogName' value='<%=RS("BlogName")%>' size="40"> <font color=red>*</font></td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>�����ˣ�</strong></td>
           <td height='28'>&nbsp;<%=RS("username")%></td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>�ռ�������</strong></td>
           <td height='28'>&nbsp;<input type='text' name='domain' value='<%=RS("domain")%>' size="20"><b>.<%response.write KS.SSetting(16)%></b></td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>�ռ�Logo��</strong></td>
           <td height='28'>&nbsp;<input type='text' name='logo' value='<%=RS("logo")%>' size="40"></td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>�ռ���ࣺ</strong></td>
           <td height='28'>&nbsp;<select class="textbox" size='1' name='ClassID' style="width:250">
                    <option value="0">-��ѡ�����-</option>
                    <% Dim RSC:Set RSC=Server.CreateObject("ADODB.RECORDSET")
							  RSC.Open "Select * From KS_BlogClass order by orderid",conn,1,1
							  If Not RSC.EOF Then
							   Do While Not RSC.Eof 
							   If rs("ClassID")=RSC("ClassID") Then
								  Response.Write "<option value=""" & RSC("ClassID") & """ selected>" & RSC("ClassName") & "</option>"
							   Else
								  Response.Write "<option value=""" & RSC("ClassID") & """>" & RSC("ClassName") & "</option>"
							   End iF
								 RSC.MoveNext
							   Loop
							  End If
							  RSC.Close:Set RSC=Nothing
							  %>
                  </select>   </td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>վ��������</strong></td>
           <td height='28'>&nbsp;<textarea class="textbox" name="Descript" id="Descript" style="width:80%;height:60px" cols=50 rows=6><%=RS("Descript")%></textarea> </td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>�ռ乫�棺</strong></td>
           <td height='28'>&nbsp;<textarea class="textbox" name="Announce" id="Announce" style="width:80%;height:80px" cols=50 rows=6><%=RS("Announce")%></textarea> </td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>��־Ĭ�ϲ�����ʾ������</strong></td>
           <td height='28'>&nbsp;<input class="textbox" name="ContentLen" type="text" id="ContentLen" style="width:250px; " value="<%=rs("ContentLen")%>" />            </td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>ÿҳ��ʾ��־ƪ����</strong></td>
           <td height='28'>&nbsp;<input class="textbox" name="ListBlogNum" type="text" id="ListBlogNum" style="width:250px; " value="<%=RS("ListBlogNum")%>" />  </td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>��ʾ���»ظ�������</strong></td>
           <td height='28'>&nbsp;<input class="textbox" name="ListReplayNum" type="text" id="ListReplayNum" style="width:250px; " value="<%=RS("ListReplayNum")%>" />  </td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>��ʾ������־ƪ����</strong></td>
           <td height='28'>&nbsp;<input class="textbox" name="ListLogNum" type="text" id="ListLogNum" style="width:250px; " value="<%=RS("ListLogNum")%>" />             </td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>��ʾ��������������</strong></td>
           <td height='28'>&nbsp;<input class="textbox" name="ListGuestNum" type="text" id="ListGuestNum" style="width:250px; " value="<%=rs("ListGuestNum")%>" />  </td>
          </tr>  
          <tr class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
            <td  width='160' height='30' align='right' class='clefttitle'><strong>״̬��</strong></td>
           <td height='28'>&nbsp;<input name="Status" type="radio" value="1"<%if rs("status")=1 then response.write " checked"%> /> ����� <input name="Status" type="radio" value="0" <%if rs("status")=0 then response.write " checked"%>/> δ���<input name="Status" type="radio" value="2" <%if rs("status")=2 then response.write " checked"%>/> ����</td>
          </tr>  
         
   </form>
 </table>
 <%
 rs.close:set rs=nothing
End Sub

Sub ModifySave()
 Dim ID:ID=KS.ChkClng(Request("id"))
 If ID=0 Then KS.AlertHintScript "�Բ��𣬷Ƿ�����!"
 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
 RS.Open "select * from ks_blog where blogid=" & id,conn,1,3
 If RS.Eof AND RS.Bof Then
   RS.Close
   Set RS=Nothing
   KS.AlertHintScript "�Բ����Ҳ�����¼��"
 End If
 RS("BlogName")=KS.G("BlogName")
 RS("Domain")=KS.G("Domain")
 RS("Logo")=KS.G("Logo")
 RS("ClassID")=KS.ChkClng(KS.G("ClassID"))
 RS("Descript")=KS.G("Descript")
 RS("Announce")=KS.G("Announce")
 RS("ContentLen")=KS.ChkClng(KS.G("ContentLen"))
 RS("ListBlogNum")=KS.ChkClng(KS.G("ListBlogNum"))
 RS("ListReplayNum")=KS.ChkClng(KS.G("ListReplayNum"))
 RS("ListLogNum")=KS.ChkClng(KS.G("ListLogNum"))
 RS("ListGuestNum")=KS.ChkClng(KS.G("ListGuestNum"))
 RS("Status")=KS.ChkClng(KS.G("Status"))
 RS.Update
 RS.Close
 Set RS=Nothing
 Response.Write "<script>alert('��ϲ���ռ��޸ĳɹ���');$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Disabled&OpStr='+escape('�ռ��Ż����� >> <font color=red>�ռ�վ�����</font>');location.href='"& Request.Form("ComeUrl") & "';</script>"
End Sub

'ɾ����־
Sub BlogDel()
 Dim ID:ID=KS.G("ID")
 Dim UserName
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 Dim RS:Set RS=Server.CreateOBject("ADODB.RECORDSET")
 RS.Open "Select * from KS_Blog Where BlogID in(" & id & ")",conn,1,1
 do while not rs.eof
  username=rs("username")
  Conn.execute("Delete From KS_BlogInfo Where username='" & username & "'")
  Conn.Execute("Delete From KS_BlogComment Where username='" & username &"'")
  Conn.execute("Delete From KS_BlogMessage Where Username='" & username & "'")
  rs.movenext
 loop
 rs.close:set rs=nothing
 Conn.Execute("Delete From KS_UploadFiles Where ChannelID=1025 and infoID in(" & ID & ")")
 Conn.execute("Delete From KS_Blog Where BlogID In("& id & ")")
 Response.Write "<script>alert('ɾ���ɹ���');location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'��Ϊ����
Sub Blogrecommend()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 Conn.execute("Update KS_Blog Set recommend=1 Where BlogID In("& id & ")")
 Conn.execute("Update KS_EnterPrise Set recommend=1 Where username In(select username from ks_blog where blogid in("& id & "))")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'ȡ������
Sub BlogCancelrecommend()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 Conn.execute("Update KS_Blog Set recommend=0 Where BlogID In("& id & ")")
 Conn.execute("Update KS_EnterPrise Set recommend=0 Where username In(select username from ks_blog where blogid in("& id & "))")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'����
Sub BlogLock()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 Conn.execute("Update KS_Blog Set status=2 Where BlogID In("& id & ")")
 conn.execute("update ks_enterprise set status=2 where username in(select username from ks_blog where blogid in("&id&"))")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'����
Sub BlogUnLock()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 Conn.execute("Update KS_Blog Set status=1 Where BlogID In("& id & ")")
 conn.execute("update ks_enterprise set status=1 where username in(select username from ks_blog where blogid in("&id&"))")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'���
Sub Blogverific
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 Conn.execute("Update KS_Blog Set status=1 Where BlogID In("& id & ")")
 conn.execute("update ks_enterprise set status=1 where username in(select username from ks_blog where blogid in("&id&"))")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub

Sub ShowClass
%>		
		<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
		  <tr align="center"  class="sort"> 
			<td width="87"><strong>���</strong></td>
			<td width="217"><strong>��������</strong></td>
			<td width="197"><strong>����</strong></td>
			<td width="196"><strong>�������</strong></td>
		  </tr>
		  <%dim orderid
		  set rs = conn.execute("select * from KS_BlogClass order by orderid")
		    if rs.eof and rs.bof then
			  Response.Write "<tr><td colspan=""6"" height=""25"" align=""center"" class=""list"">��û������κεĲ�������!</td></tr>"
			else
			   do while not rs.eof%>
			  <form name="form1" method="post" action="?action=class&x=a">
				<tr  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'"> 
				  <td class="splittd" width="87" height="22" align="center"><%=rs("ClassID")%> <input name="ClassID" type="hidden" id="ClassID" value="<%=rs("ClassID")%>"></td>
				  <td class="splittd" width="217" align="center"><input name="ClassName" type="text" class="textbox" id="ClassName" value="<%=rs("ClassName")%>" size="25"></td>
				  <td class="splittd" width="197" align="center"><input style="text-align:center" name="OrderID" type="text" class="textbox" id="OrderID" value="<%=rs("OrderID")%>" size="8">				  </td>
				  <td class="splittd" align="center"><input name="Submit" class="button" type="submit"value=" �޸� ">&nbsp;<a  onclick='return(confirm("ȷ��ɾ����"))' href="?action=class&x=c&ClassID=<%=rs("ClassID")%>">ɾ��</a></td>
				</tr>
			  </form>
		  <%orderid=rs("orderid")
		   rs.movenext
		   loop
		 End IF
		rs.close%>
				<form action="?action=class&x=b" method="post" name="myform" id="form">
		    <tr>
		      <td class="splittd" height="25" colspan="4">&nbsp;&nbsp;<strong>&gt;&gt;�����ռ����<<</strong></td>
		    </tr>

			<tr valign="middle" class="list"> 
			  <td height="25"></td>
			  <td height="25" align="center"><input name="ClassName" type="text" class="textbox" id="ClassName" size="25"></td>
			  <td height="25" align="center"><input style="text-align:center" name="orderid" type="text" value="<%=orderid+1%>" class="textbox" id="orderid" size="8">
			  <td height="25" align="center"><input name="Submit3" class="button" type="submit" value="OK,�ύ"></td>
			</tr>
		</form>
</table>

		<% Select case request("x")
		   case "a"
				conn.execute("Update KS_BlogClass set ClassName='" & KS.G("ClassName") & "',orderid='" & KS.ChkClng(KS.G("OrderID")) &"' where ClassID="&KS.G("ClassID")&"")
				KS.AlertHintScript "��ϲ,�ռ�����޸ĳɹ�"
		   case "b"
		       If KS.G("ClassName")="" Then KS.Die "<script>alert('��������������!');history.back();</script>"
			   
				conn.execute("Insert into KS_BlogClass(ClassName,orderid)values('" & KS.G("ClassName") & "','" & KS.ChkClng(KS.G("OrderID")) &"')")
				KS.AlertHintScript "��ϲ,�ռ������ӳɹ�"
		   case "c"
				conn.execute("Delete from KS_BlogClass where ClassID="&KS.G("ClassID")&"")
				KS.AlertHintScript "��ϲ,�ռ����ɾ���ɹ�"
		End Select
		%></body>
		</html>
<%End Sub
End Class
%> 
