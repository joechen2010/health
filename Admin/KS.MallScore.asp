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
Set KSCls = New Admin_EnterPrise
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_EnterPrise
        Private KS,Param,KSCls
		Private Action,i,strClass,sFileName,RS,SQL,maxperpage,CurrentPage,totalPut,TotalPageNum
        Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub

		Public Sub Kesion()
		 With KS
					If Not KS.ReturnPowerResult(0, "KSMS20010") Then          '�����Ȩ��
					 Call KS.ReturnErr(1, "")
					 .End
					 End If
			  .echo "<html>"
			  .echo"<head>"
			  .echo"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			  .echo"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			  .echo "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
			  .echo "<script src=""../KS_Inc/jquery.js"" language=""JavaScript""></script>"
			  .echo"</head>"
			  .echo"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			  .echo "<ul id='menu_top'>"
			  .echo "<li class='parent' onclick=""window.parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape('���ֶһ�ϵͳ >> <font color=red>�����Ʒ</font>')+'&ButtonSymbol=GOSave';location.href='?action=Add';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>�����Ʒ</span></li>"
			  .echo "<li class='parent' onclick=""location.href='KS.MallScoreOrder.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/move.gif' border='0' align='absmiddle'>����һ�����</span></li>"
			  .echo "<li style='margin-left:30px;margin-top:10px'><strong>�鿴��ʽ:</strong><a href=""KS.MallScore.asp"">������Ʒ</a> <a href=""KS.MallScore.asp?flag=1"">����</a>  <a href=""KS.MallScore.asp?flag=2"">δ��</a> <a href=""KS.MallScore.asp?flag=3"">�ѽ���</a></li>"

			  .echo "</ul>"
		
		
		maxperpage = 30 '###ÿҳ��ʾ��
		If Not IsNumeric(Request("page")) And Len(Request("page")) <> 0 Then
			.echo ("�����ϵͳ����!����������")
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
		   Param= Param & " and ProductName like '%" & KS.G("KeyWord") & "%'"
		  Else
		   Param= Param & " and Intro like '%" & KS.G("KeyWord") & "%'"
		  End If
		End If
		If KS.G("Flag")<>"" Then
		  If KS.G("Flag")="1" Then Param=Param & " and Status=1"
		  If KS.G("Flag")="2" Then Param=Param & " and Status=0"
		  If KS.G("Flag")="3" Then Param=Param & " and datediff(day,enddate," & SqlNowString & ")>0"
		  
		End If

		totalPut = Conn.Execute("Select Count(id) From KS_MallScore " & Param)(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '�õ���ҳ��
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
		Select Case KS.G("action")
		 Case "Add","Edit" Call ProductNameManage()
		 Case "EditSave" Call DoSave()
		 Case "Del"  Call BlogDel()
		 Case "lock"  Call BlogLock()
		 Case "unlock"  Call BlogUnLock()
		 Case "recommend"  Call Blogrecommend()
		 Case "Cancelrecommend" Call BlogCancelrecommend()
		 Case Else
		  Call showmain
		End Select
	End With	
End Sub

Private Sub showmain()
%>
<table width="100%" border="0" align="center" cellspacing="0" cellpadding="0">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>ѡ��</th>
	<td nowrap>��Ʒ����</th>
	<td nowrap>�������</th>
	<td nowrap>�������</th>
	<td nowrap>����ʱ��</th>
	<td nowrap>�������</th>
	<td nowrap>�һ�����</th>
	<td nowrap>״̬</th>
	<td nowrap>�������</th>
</tr>
<%
	sFileName = "KS.MallScore.asp?"
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_MallScore " & Param & " order by id desc"
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
		Response.Write "<tr><td height=""25"" align=center bgcolor=""#ffffff"" colspan=7>��û������κ���Ʒ��</td></tr>"
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
	<td  class="splittd"><%=Rs("ProductName")%>
	</td>
	<td align="center" class="splittd"><font color=red><%=Rs("Score")%></font> ��</td>
	<td align="center" class="splittd"><font color=red><%=Rs("Quantity")%></font></td>
	<td align="center" class="splittd"><font color=#cccccc><%=RS("EndDate")%></font></td>
	<td align="center" class="splittd"><%=RS("Hits")%> ��</td>
	<td align="center" class="splittd">
	 <span style="color:red;font-weight:bold"><%=LFCls.GetSingleFieldValue("Select Count(*) From KS_MallScoreOrder Where ProductID=" & RS("ID"))%></span>
	(<a href="#" onclick="javascript:window.parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape('���ֶһ�ϵͳ >> <font color=red>�鿴����</font>')+'&ButtonSymbol=Disabled';location.href='KS.MallScoreOrder.asp?productid=<%=RS("ID")%>'">�鿴</a>)
	
	</td>
	<td align="center" class="splittd"><%
		if rs("status")=1 then
		  response.write "<font color=#cccccc>����</font>"
		else
		  response.write " <font color=red>δ��</font>"
		end if
	%></td>
	<td align="center" class="splittd"><a href="?action=Edit&ID=<%=RS("ID")%>"  onclick="window.parent.frames['BottomFrame'].location.href='KS.Split.asp?OpStr='+escape('���ֶһ�ϵͳ >> <font color=red>�޸��Ź���Ϣ</font>')+'&ButtonSymbol=GOSave';">�޸�</a> <a href="?Action=Del&ID=<%=rs("id")%>" onclick="return(confirm('ȷ��ɾ������Ʒ��'));">ɾ��</a> 
		
		&nbsp;<%IF rs("Status")="1" then %><a href="?Action=Cancelrecommend&id=<%=rs("id")%>"><font color=red>ȡ��</font></a><%else%><a href="?Action=recommend&id=<%=rs("id")%>">���</a><%end if%>

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
	<input class=Button type="submit" name="Submit2" value=" ɾ��ѡ�е���Ʒ " onclick="{if(confirm('�˲��������棬ȷ��Ҫɾ��ѡ�еļ�¼��?')){this.document.selform.submit();return true;}return false;}"></td>
</tr>
</form>
<tr>
	<td  colspan=7 align=right>
	<%
	 Call KSCLS.ShowPage(totalPut, MaxPerPage, "KS.MallScore.asp", True, "��", CurrentPage, KS.QueryParam("page"))
	%></td>
</tr>
</table>
<div>
<form action="KS.MallScore.asp" name="myform" method="get">
   <div style="border:1px dashed #cccccc;margin:3px;padding:4px">
      &nbsp;<strong>��������=></strong>
	 &nbsp;�ؼ���:<input type="text" class='textbox' name="keyword">&nbsp;����:
	 <select name="condition">
	  <option value=1>����Ʒ����</option>
	  <option value=2>����Ʒ����</option>
	 </select>
	  &nbsp;<input type="submit" value="��ʼ����" class="button" name="s1">
    </div>
</form>
</div>
<%
End Sub

Sub ProductNameManage()
Dim ProductName,ActiveDate,AddDate,EndDate,Quantity,Score,Telphone,Intro,Hits,Protection,BuyFlow,Notes,recommend,Status,PhotoUrl
Dim ID:ID=KS.ChkClng(KS.G("ID"))
Dim RS:Set RS=server.createobject("adodb.recordset")
If KS.G("Action")="Edit" Then
	RS.Open "Select * From KS_MallScore Where ID=" & ID,conn,1,1
	 If RS.Eof And RS.Bof Then
	  RS.Close:Set RS=Nothing
	  KS.AlertHintScript "�������ݳ���"
	  Response.End
	 Else
	   ProductName=RS("ProductName")
	   AddDate=RS("AddDate")
	   EndDate=RS("EndDate")
	   Quantity=RS("Quantity")
	   Score=RS("Score")
	   Intro=RS("Intro")
	   PhotoUrl=RS("PhotoUrl")
	   If PhotoUrl="" Or IsNull(PhotoUrl) Then PhotoUrl="../images/nophoto.gif"
	   Hits=RS("Hits")
	   recommend=RS("recommend")
	   Status=RS("Status")
	 End If
Else
  AddDate=Now
  EndDate=Now+30
  Hits=0:Score=10
  recommend=0:Status=1
  Quantity=100
  Intro=" "
  PhotoUrl="../images/nophoto.gif"
 End If
%>
<script>
function CheckForm()
{
	if ($('input[name=ProductName]').val()=='')
	{
	 alert('��������Ʒ����!');
	 $("input[name=ProductName]").focus();
	 return false;
	}
	if ($('input[name=Quantity]').val()=='')
	{
	 alert('������������!');
	 $("input[name=Quantity]").focus();
	 return false;
	}
	if ($('input[name=Score]').val()=='')
	{
	 alert('�������������!');
	 $("input[name=Score]").focus();
	 return false;
	}
	if (FCKeditorAPI.GetInstance('Intro').GetXHTML(true)=="")
	{
	 alert('��������Ʒ����!');
	 return false;
	}
document.myform.submit();
}
</script>
<br>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
  <form name="myform" action="?action=EditSave" method="post" enctype="multipart/form-data">
    <input type="hidden" value="<%=ID%>" name="id" />
    <input type="hidden" value="<%=request.servervariables("http_referer")%>" name="comeurl" />
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='296' height='30' align='right' class='clefttitle'><strong>��Ʒ���ƣ�</strong></td>
      <td width="435" height='30'>&nbsp;
          <input type='text' name='ProductName' value='<%=ProductName%>' size="40" />
          <font color=red>*</font></td>
      <td width="217" rowspan="4" align="center"><div id="pic" style="filter:progid:DXImageTransform.Microsoft.AlphaImageLoader(sizingMethod=scale);height:100px;width:95px;border:1px solid #777777"> <img src="<%=PhotoUrl%>" style="height:100px;width:95px;" /> </div></td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='296' height='30' align='right' class='clefttitle'><strong>����ʱ�䣺</strong></td>
      <td height='30'>&nbsp;
          <select name="AddDate1">
            <%on error resume next
					  for i=year(now) to year(now)+1
					   if trim(split(AddDate,"-")(0))=trim(i) then
					   response.write "<option value=" & i & " selected>" & i & "��</option>"
					   else
					   response.write "<option value=" & i & ">" & i & "��</option>"
					   end if
					  next
					  %>
          </select>
          <select name="AddDate2">
            <%
					  for i=1 to 12
					   if trim(split(AddDate,"-")(1))=trim(i) then
					   response.write "<option value=" & i & " selected>" & i & "��</option>"
					   else
					   response.write "<option value=" & i & ">" & i & "��</option>"
					   end if
					  next
					  %>
          </select>
          <select name="AddDate3">
            <%
					  for i=1 to 31
					   if trim(split(split(AddDate,"-")(2)," ")(0))=trim(i) then
					   response.write "<option value=" & i & " selected>" & i & "��</option>"
					   else
					   response.write "<option value=" & i & ">" & i & "��</option>"
					   end if
					  next
					  %>
          </select>
          <select name="AddDate4">
            <%
					  for i=0 to 23
					   if trim(split(split(AddDate," ")(1),":")(0))=trim(i) then
					   response.write "<option value=" & i & " selected>" & i & "ʱ</option>"
					   else
					   response.write "<option value=" & i & ">" & i & "ʱ</option>"
					   end if
					  next
					  %>
          </select>
          <select name="AddDate5">
            <%
					  for i=0 to 59
					   if trim(split(split(AddDate," ")(1),":")(1))=trim(i) then
					   response.write "<option value=" & i & " selected>" & i & "��</option>"
					   else
					   response.write "<option value=" & i & ">" & i & "��</option>"
					   end if
					  next
					  %>
          </select>
      </td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='296' height='30' align='right' class='clefttitle'><strong>����ʱ�䣺</strong></td>
      <td height='30'>&nbsp;
          <select name="EndDate1">
            <%
					  for i=year(now) to year(now)+1
					   if trim(split(EndDate,"-")(0))=trim(i) then
					   response.write "<option value=" & i & " selected>" & i & "��</option>"
					   else
					   response.write "<option value=" & i & ">" & i & "��</option>"
					   end if
					  next
					  %>
          </select>
          <select name="EndDate2">
            <%
					  for i=1 to 12
					   if trim(split(EndDate,"-")(1))=trim(i) then
					   response.write "<option value=" & i & " selected>" & i & "��</option>"
					   else
					   response.write "<option value=" & i & ">" & i & "��</option>"
					   end if
					  next
					  %>
          </select>
          <select name="EndDate3">
            <%
					  for i=1 to 31
					   if trim(split(split(EndDate,"-")(2)," ")(0))=trim(i) then
					   response.write "<option value=" & i & " selected>" & i & "��</option>"
					   else
					   response.write "<option value=" & i & ">" & i & "��</option>"
					   end if
					  next
					  %>
          </select>
          <select name="EndDate4">
            <%
					  for i=0 to 23
					   if trim(split(split(EndDate," ")(1),":")(0))=trim(i) then
					   response.write "<option value=" & i & " selected>" & i & "ʱ</option>"
					   else
					   response.write "<option value=" & i & ">" & i & "ʱ</option>"
					   end if
					  next
					  %>
          </select>
          <select name="EndDate5">
            <%
					  for i=0 to 59
					   if trim(split(split(EndDate," ")(1),":")(1))=trim(i) then
					   response.write "<option value=" & i & " selected>" & i & "��</option>"
					   else
					   response.write "<option value=" & i & ">" & i & "��</option>"
					   end if
					  next
					  %>
          </select>
        �������ʱ�佫���ܶһ� </td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td height='30' align='right' class='clefttitle'><strong>��ƷͼƬ��</strong></td>
      <td height='30'>&nbsp;
  <input class="textbox" type="file" name="photo" size="40" onchange='document.getElementById(&quot;pic&quot;).innerHTML=&quot;&quot;;document.getElementById(&quot;pic&quot;).filters.item(&quot;DXImageTransform.Microsoft.AlphaImageLoader&quot;).src=this.value;' />
        <font color="red">*</font> <br />
        &nbsp;&nbsp;<font color="blue">���ϴ�����200K��ͼƬ,֧��jpg,gif,png��ʽ</font></td></tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='296' height='30' align='right' class='clefttitle'><strong>���������</strong></td>
      <td height='30' colspan="2">&nbsp;
          <input type='text' name='Quantity' value='<%=Quantity%>' size="10" />
          <font color=red>*</font></td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='296' height='30' align='right' class='clefttitle'><strong>������֣�</strong></td>
      <td height='30' colspan="2">&nbsp;
          <input type='text' name='Score' value='<%=Score%>' size="10" />
        �� <font color=red>*</font></td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='296' height='30' align='right' class='clefttitle'><strong>��Ʒ��飺</strong></td>
      <td height='30' colspan="2">&nbsp;
          <%			
		    Response.Write "<textarea id=""Intro"" name=""Intro"" style=""display:none"">" &  KS.HTMLCode(Intro) &"</textarea><input type=""hidden"" id=""Intro___Config"" value="""" style=""display:none"" /><iframe id=""Intro___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=Intro&amp;Toolbar=NewsTool"" width=""695"" height=""220"" frameborder=""0"" scrolling=""no""></iframe>"
								%>
      </td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='296' height='30' align='right' class='clefttitle'><strong>���������</strong></td>
      <td height='30' colspan="2">&nbsp;
          <input name='Hits' value="<%=Hits%>" size="10" /></td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='296' height='30' align='right' class='clefttitle'><strong>�Ƿ��Ƽ���</strong></td>
      <td height='30' colspan="2">&nbsp;
          <input type="radio" name="recommend" value="0"<%if recommend=0 then response.write " checked"%> />
        ��
        <input type="radio" name="recommend" value="1"<%if recommend=1 then response.write " checked"%> />
        �� </td>
    </tr>
    <tr class="tdbg" onmouseover="this.className='tdbgmouseover'" onmouseout="this.className='tdbg'">
      <td  width='296' height='30' align='right' class='clefttitle'><strong>�Ƿ���ˣ�</strong></td>
      <td height='30' colspan="2">&nbsp;
          <input type="radio" name="Status" value="0"<%if Status=0 then response.write " checked"%> />
        ��
        <input type="radio" name="Status" value="1"<%if Status=1 then response.write " checked"%> />
        �� </td>
    </tr>
  </form>
</table>
<%
End Sub

Sub DoSave()
		  	Dim Fobj:Set FObj = New UpFileClass
			on error resume next
			FObj.GetData
			if err.number<>0 then
			 call ks.alerthistory("������,�ļ�������С",-1)
			 response.End()
			end if

       Dim ID:ID=KS.ChkClng(Fobj.Form("id"))
	   Dim ProductName:ProductName=KS.LoseHtml(Fobj.Form("ProductName"))
       Dim AddDate:AddDate=Fobj.Form("AddDate1") & "-" & Fobj.Form("AddDate2") & "-" & Fobj.Form("AddDate3") & " " & Fobj.Form("AddDate4") & ":" & Fobj.Form("AddDate5")
			if not isdate(AddDate) then
			 Response.Write "<script>alert('����ʱ���ʽ����ȷ��');history.back();</script>"
			 Exit Sub
			End If	 
       Dim EndDate:EndDate=Fobj.Form("EndDate1") & "-" & Fobj.Form("EndDate2") & "-" & Fobj.Form("EndDate3") & " " & Fobj.Form("EndDate4") & ":" & Fobj.Form("EndDate5")
			if not isdate(EndDate) then
			 Response.Write "<script>alert('����ʱ���ʽ����ȷ��');history.back();</script>"
			 Exit Sub
			End If	 
			
			Dim MaxFileSize:MaxFileSize = 200   '�趨�ļ��ϴ�����ֽ���
			Dim AllowFileExtStr:AllowFileExtStr = "gif|jpg|png"
			Dim FormPath:FormPath = KS.GetUpFilesDir() & "/MallScore/"
			Call KS.CreateListFolder(FormPath) 
			Dim ReturnValue:ReturnValue = FObj.UpSave(FormPath,MaxFileSize,AllowFileExtStr,"fm" & right(Year(Now),2) & right("0" & Month(Now),2) & right("0" & Day(Now),2) & right("0"&Hour(Now),2) & right("0"&Minute(Now),2) & right("0"&Second(Now),2))
			Select Case ReturnValue
			  Case "errext" Call KS.AlertHistory("�ļ��ϴ�ʧ��,�ļ����Ͳ�����\n�����������" + AllowFileExtStr + "\n",-1):exit sub
	          Case "errsize"  Call KS.AlertHistory("�ļ��ϴ�ʧ��,�ļ����������ϴ��Ĵ�С\n�����ϴ� " & MaxFileSize & " KB���ļ�\n",-1):exit sub
			End Select

	   Dim PhotoUrl:PhotoUrl=ReturnValue


			 
	   Dim Quantity:Quantity=KS.ChkClng(Fobj.Form("Quantity"))
	   Dim Score:Score=KS.ChkClng(Fobj.Form("Score"))
	   Dim Intro:Intro=Fobj.Form("Intro")
	   Dim Hits:Hits=KS.LoseHtml(Fobj.Form("Hits"))
	   Dim recommend:recommend=KS.ChkClng(Fobj.Form("recommend"))
	   Dim Status:Status=KS.ChkClng(Fobj.Form("Status"))
	   Dim ComeUrl:ComeUrl=Fobj.Form("ComeUrl")
	   Set Fobj=Nothing
	   
		
	   If ProductName="" Then Response.Write "<script>alert('��Ʒ���Ʊ�������');history.back();</script>":response.end

            Dim RS: Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select * From KS_MallScore Where ID=" & ID,Conn,1,3
			  If RS.Eof And RS.Bof Then
			     RS.AddNEW
				 RS("Inputer")=KS.C("AdminName")
			  End If
				 RS("AddDate")=AddDate
				 RS("EndDate")=EndDate
			     RS("ProductName")=ProductName
				 RS("Quantity")=Quantity
				 RS("Score")=Score
				 RS("Intro")=Intro
				 IF PhotoUrl<>"" Then RS("PhotoUrl")=PhotoUrl
				 RS("Hits")=Hits
				 RS("recommend")=recommend
				 RS("Status")=Status
		 		 RS.Update
				 If ID=0 Then
				   RS.MoveLast
                   Call KS.FileAssociation(1004,RS("ID"),Intro&RS("PhotoUrl"),0)
				 Else
                   Call KS.FileAssociation(1004,ID,Intro&RS("PhotoUrl"),1)
				 End If
				 
			     RS.Close
				 Set RS=Nothing
				 If ID=0 Then
				  Response.Write "<script>if (confirm('��Ʒ��Ϣ�����ɹ�!')){location.href='?action=Add';}else{parent.frames['BottomFrame'].location.href='KS.Split.asp?ButtonSymbol=Disabled&OpStr=" & Server.URLEncode("���ֶһ�ϵͳ >> <font color=red>������ҳ</font>") & "';location.href='KS.MallScore.asp';}</script>"
				 Else
				  Response.Write "<script>alert('��Ʒ��Ϣ�޸ĳɹ���');parent.frames['BottomFrame'].location.href='KS.Split.asp?ButtonSymbol=Disabled&OpStr=" & Server.URLEncode("���ֶһ�ϵͳ >> <font color=red>������ҳ</font>") & "';location.href='"& ComeUrl & "';</script>"
				 End If

EnD Sub

'ɾ����־
Sub BlogDel()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 Conn.execute("Delete From KS_MallScore Where id In("& id & ")")
 Conn.execute("Delete From KS_UploadFiles Where ChannelID=1004 and InfoID In("& id & ")")
 Response.Write "<script>alert('ɾ���ɹ���');location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub

Sub Blogrecommend()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 Conn.execute("Update KS_MallScore Set Status=1 Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
Sub BlogCancelrecommend()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 Conn.execute("Update KS_MallScore Set Status=0 Where id In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub

End Class
%> 
