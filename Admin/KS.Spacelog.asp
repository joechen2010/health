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
Set KSCls = New Admin_SpaceLog
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_SpaceLog
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
					If Not KS.ReturnPowerResult(0, "KSMS10002") Then          '�����Ȩ��
					 Call KS.ReturnErr(1, "")
					 .End
					 End If
		.Write "<script src='../KS_Inc/common.js'></script>"
		.Write "<script src='../KS_Inc/jquery.js'></script>"
		.Write "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		.Write "<ul id='menu_top'>"
		.Write "<li class='parent' onclick=""location.href='KS.Spacelog.asp';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>��־����</span></li>"
		.Write "<li class='parent' onclick=""location.href='?action=comment';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/move.gif' border='0' align='absmiddle'>��־����</span></li>"
		.Write "<li class='parent' onclick=""location.href='?action=class';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/addjs.gif' border='0' align='absmiddle'>��־����</span></li>"
		.Write	" </ul>"
		End With	
		
		
		maxperpage = 30 '###ÿҳ��ʾ��

		If Not IsEmpty(Request("page")) Then
			CurrentPage = KS.ChkClng(Request("page"))
		Else
			CurrentPage = 1
		End If
		If CInt(CurrentPage) = 0 Then CurrentPage = 1
		Select Case KS.G("action")
		 Case "Del" BlogInfoDel
		 Case "Best" BlogInfoBest
		 Case "CancelBest" BlogInfoCancelBest
		 Case "verific" Verific
		 case "comment" commentshow
		 case "commentdel" commentdel
		 case "class" classshow
		 case "modify" modify
		 case "DoSave" DoSave
		 Case Else
		  Call showmain
		End Select
End Sub

Private Sub showmain()
%>
<table width="100%" border="0" align="center" cellspacing="1" cellpadding="1">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>ѡ��</th>
	<td nowrap>��־����</th>
	<td nowrap>�û���</th>
	<td nowrap>���ʱ��</th>
	<td nowrap>״ ̬</th>
	<td nowrap>�������</th>
</tr>
<%
		Param=" where 1=1"
		If KS.G("KeyWord")<>"" Then
		  If KS.G("condition")=1 Then
		   Param= Param & " and title like '%" & KS.G("KeyWord") & "%'"
		  Else
		   Param= Param & " and username like '%" & KS.G("KeyWord") & "%'"
		  End If
		End If
		If Request("from")<>"" Then
		 Param=Param & " and status=2"
		End If

		totalPut = Conn.Execute("Select Count(ID) from KS_bloginfo" & Param)(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '�õ���ҳ��
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum

	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_BlogInfo " & Param & " order by id desc"
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
		Response.Write "<tr><td height=""25"" align=center bgcolor=""#ffffff"" colspan=7>û����д��־��</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<form name=selform method=post action="ks.spacelog.asp">
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
		
%>
<tr height="22" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td class="splittd" align="center"><input type=checkbox name=ID value='<%=Rs("id")%>'></td>
	<td class="splittd"><a href="../space/?<%=rs("username")%>/log/<%=rs("id")%>" target="_blank"><%=Rs("title")%></a><% if rs("best")="1" then response.write "<img src=""../images/jh.gif"" align=""absmiddle"">"%></td>
	<td class="splittd" align="center"><%=Rs("username")%></td>
	<td class="splittd" align="center"><%=Rs("adddate")%></td>
	<td class="splittd" align="center"><%
	select case rs("status")
	 case 0
	  response.write "����"
	 case 1
	  response.write "<font color=blue>�ݸ�</font>"
	 case else
	  response.write "<font color=red>δ��</font>"
	end select
	%></td>
	<td class="splittd" align="center">
	<a href="../space/?<%=rs("username")%>/log/<%=rs("id")%>" target="_blank">���</a> <a href="?Action=Del&ID=<%=RS("ID")%>" onclick="return(confirm('ȷ��ɾ������־��'));">ɾ��</a> <%IF rs("best")="1" then %><a href="?Action=CancelBest&id=<%=rs("id")%>"><font color=red>ȡ������</font></a><%else%><a href="?Action=Best&id=<%=rs("id")%>">��Ϊ����</a><%end if%>&nbsp;
	<%if rs("status")=2 then%><a href="?Action=verific&flag=0&id=<%=rs("id")%>">���</a> <%elseif rs("status")=0 then%><a href="?Action=verific&flag=2&id=<%=rs("id")%>" title="ȡ�����">ȡ��</a><%end if%>
	
	<a href="?action=modify&id=<%=rs("id")%>" onclick="window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape('�ռ��Ż����� >> <font color=red>�޸���־</font>')+'&ButtonSymbol=GOSave';">�޸�</a>
	
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
	<td  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" height='25' colspan=7>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">ȫѡ
	<input type="hidden" name="action" value="Del">
	<input type="hidden" name="flag" value="0">
	<input class=Button type="submit" name="Submit2" value="����ɾ��" onclick="{if(confirm('�˲��������棬ȷ��Ҫɾ��ѡ�еļ�¼��?')){document.selform.action.value='Del';this.document.selform.submit();return true;}return false;}">
	<input class="button" type="submit" value="�������" onclick="document.selform.action.value='verific';document.selform.flag.value='0';this.document.selform.submit();return true;">
	<input class="button" type="submit" value="����ȡ�����" onclick="document.selform.action.value='verific';document.selform.flag.value='2';this.document.selform.submit();return true;">
	</td>
</tr>
</form>
<tr>
	<td  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" colspan=7 align=right>
	<%
	 Call KS.ShowPageParamter(totalPut, MaxPerPage, "KS.Spacelog.asp", True, "ƪ", CurrentPage, "KeyWord=" &KS.G("KeyWord") & "&Condition=" & KS.G("Condition") & "&Action=" & Action)
	%></td>
</tr>
</table>
<div>
<form action="KS.SpaceLog.asp" name="myform" method="post">
   <div style="border:1px dashed #cccccc;margin:3px;padding:4px">
      &nbsp;<strong>��������=></strong>
	 &nbsp;�ؼ���:<input type="text" class='textbox' name="keyword">&nbsp;����:
	 <select name="condition">
	  <option value=1>��־����</option>
	  <option value=2>������</option>
	 </select>
	  &nbsp;<input type="submit" value="��ʼ����" class="button" name="s1">
	  </div>
</form>
</div>
<%
End Sub

'�޸���־
sub modify()
           Dim RSObj,TypeID,ClassID,Title,Tags,UserName,PassWord,face,weather,adddate,content,status
		   Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		   RSObj.Open "Select * From KS_BlogInfo Where ID=" & KS.ChkClng(KS.S("ID")),Conn,1,1
		   If Not RSObj.Eof Then
		     TypeID  = RSObj("TypeID")
			 ClassID = RSObj("ClassID")
			 Title    = RSObj("Title")
			 Tags = RSObj("Tags")
			 UserName   = RSObj("UserName")
			 password = RSObj("password")
			 Face   = RSObj("Face")
			 weather=RSObj("Weather")
			 adddate=RSObj("adddate")
			 Content  = RSObj("Content")
			 Status  = RSObj("Status")
		   End If
		   RSObj.Close:Set RSObj=Nothing
%>
<script language = "JavaScript">
function CheckForm()
{
 document.myform.submit();
}
</script>
<table width="100%" style="margin-top:2px" border="0" align="center" cellpadding="3" cellspacing="1" class="ctable">
                  <form  action="?Action=DoSave&ID=<%=KS.S("ID")%>" method="post" name="myform" id="myform" onSubmit="return CheckForm();">

                    <tr class="tdbg">
                       <td width="12%"  height="25" align="center"><span>��־���ࣺ</span></td>
                       <td width="88%">��
                          <select class="textbox" size='1' name='TypeID' style="width:150">
                             <option value="0">-��ѡ�����-</option>
							  <% Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
							  RS.Open "Select * From KS_BlogType order by orderid",conn,1,1
							  If Not RS.EOF Then
							   Do While Not RS.Eof 
							     If TypeID=RS("TypeID") Then
								  Response.Write "<option value=""" & RS("TypeID") & """ selected>" & RS("TypeName") & "</option>"
								 Else
								  Response.Write "<option value=""" & RS("TypeID") & """>" & RS("TypeName") & "</option>"
								 End If
								 RS.MoveNext
							   Loop
							  End If
							  RS.Close:Set RS=Nothing
							  %>
                         </select>
						
					  </td>
                    </tr>
                      <tr class="tdbg">
                           <td  height="25" align="center"><span>��־���⣺</span></td>
                              <td> ��
                                        <input class="textbox" name="Title" type="text" id="Title" style="width:350px; " value="<%=Title%>" maxlength="100" />
                                          <span style="color: #FF0000">*</span></td>
                    </tr>
                              <tr class="tdbg">
                                      <td height="25" align="center"><span>��־���ڣ�</span></td>
                                      <td>��
                                        <input name="AddDate"  class="textbox" type="text" id="AddDate" value="<%=adddate%>" style="width:250px; " />
                                       ����<Select Name="Weather" Size="1" onChange="Chang(this.value,'WeatherSrc','../user/images/weather/')">
									   <Option value="sun.gif"<%if weather="sun.gif" then response.write " selected"%>>����</Option>
									   <Option value="sun2.gif"<%if weather="sun2.gif" then response.write " selected"%>>����</Option>
									   <Option value="yin.gif"<%if weather="yin.gif" then response.write " selected"%>>����</Option>
									   <Option value="qing.gif"<%if weather="qing.gif" then response.write " selected"%>>��ˬ</Option>
									   <Option value="yun.gif"<%if weather="yun.gif" then response.write " selected"%>>����</Option>
									   <Option value="wu.gif"<%if weather="wu.gif" then response.write " selected"%>>����</Option>
									   <Option value="xiaoyu.gif"<%if weather="xiaoyu.gif" then response.write " selected"%>>С��</Option>
									   <Option value="yinyu.gif"<%if weather="yinyu.gif" then response.write " selected"%>>����</Option>
									   <Option value="leiyu.gif"<%if weather="leiyu.gif" then response.write " selected"%>>����</Option>
									   <Option value="caihong.gif"<%if weather="caihong.gif" then response.write " selected"%>>�ʺ�</Option>
									   <Option value="hexu.gif"<%if weather="hexu.gif" then response.write " selected"%>>����</Option>
									   <Option value="feng.gif"<%if weather="feng.gif" then response.write " selected"%>>����</Option>
									   <Option value="xue.gif"<%if weather="xue.gif" then response.write " selected"%>>Сѩ</Option>
									   <Option value="daxue.gif"<%if weather="daxue.gif" then response.write " selected"%>>��ѩ</Option>
									   <Option value="moon.gif"<%if weather="moon.gif" then response.write " selected"%>>��Բ</Option>
									   <Option value="moon2.gif"<%if weather="moon2.gif" then response.write " selected"%>>��ȱ</Option>
									</Select>
		<img id="WeatherSrc" src="../user/images/weather/<%=weather%>" border="0"></td>
                              </tr>
                              <tr class="tdbg">
                                      <td height="25" align="center"><span>Tag�� ǩ��</span></td>
                                      <td>��
                                        <input name="Tags"  class="textbox" type="text" id="Tags" value="<%=Tags%>" style="width:250px; " />
                                        �Կո�ָ�</td>
                              </tr>
                              <tr class="tdbg">
                                      <td  height="25" align="center"><span>��־���飺</span></td>
                                <td>
									  &nbsp;&nbsp;<input type="radio" name="face" value="0"<%If face=0 Then Response.Write " checked"%>>
        ��<input name="face" type="radio" value="1"<%If face=1 Then Response.Write " checked"%>><img src="../user/images/face/1.gif" width="20" height="20"> 
        <input type="radio" name="face" value="2"<%If face=2 Then Response.Write " checked"%>><img src="../user/images/face/2.gif" width="20" height="20"><input type="radio" name="face" value="3"<%If face=3 Then Response.Write " checked"%>><img src="../user/images/face/3.gif" width="20" height="20"> 
        <input type="radio" name="face" value="4"<%If face=4 Then Response.Write " checked"%>><img src="../user/images/face/4.gif" width="20" height="20"> 
        <input type="radio" name="face" value="5"<%If face=5 Then Response.Write " checked"%>><img src="../user/images/face/5.gif" width="20" height="20"> 
        <input type="radio" name="face" value="6"<%If face=6 Then Response.Write " checked"%>><img src="../user/images/face/6.gif" width="18" height="20"> 
        <input type="radio" name="face" value="7"<%If face=7 Then Response.Write " checked"%>><img src="../user/images/face/7.gif" width="20" height="20"> 
        <input type="radio" name="face" value="8"<%If face=8 Then Response.Write " checked"%>><img src="../user/images/face/8.gif" width="20" height="20"> 
        <input type="radio" name="face" value="9"<%If face=9 Then Response.Write " checked"%>><img src="../user/images/face/9.gif" width="20" height="20">
        <input type="radio" name="face" value="10"<%If face=10 Then Response.Write " checked"%>><img src="../user/images/face/10.gif" width="20" height="20">
        <input type="radio" name="face" value="11"<%If face=11 Then Response.Write " checked"%>><img src="../user/images/face/11.gif" width="20" height="20">
        <input type="radio" name="face" value="12"<%If face=12 Then Response.Write " checked"%>><img src="../user/images/face/12.gif" width="20" height="20"></td>
                              </tr>
							 
                              <tr class="tdbg">
                                  <td align="center">��־���ݣ�</td>
								  <td><div align=center>
								  <%
								  Response.Write "<textarea ID='Content' name='Content' style='display:none'>" & Server.HTMLEncode(Content) & "</textarea>"
					               Response.Write "<iframe id=""Content___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=Content&amp;Toolbar=NewsTool"" width=""93%"" height=""320"" frameborder=""0"" scrolling=""no""></iframe>"
								  %>
								                                 </div></td>
                            </tr>
                              
                  </form>
			    </table>
<%
end sub

sub DoSave()
     dim TypeID,ClassID,Title,Tags,UserName,PassWord,face,weather,adddate,content
                 TypeID=KS.ChkClng(KS.S("TypeID"))
				 Title=Trim(KS.S("Title"))
				 Tags=Trim(KS.S("Tags"))
				 UserName=Trim(KS.S("UserName"))
				 Face=Trim(KS.S("Face"))
				 weather=KS.S("weather")
				 adddate=KS.S("adddate")
				 Content = Request.Form("Content")
				  Dim RSObj
				  if TypeID="" Then TypeID=0
				  If TypeID=0 Then
				    Response.Write "<script>alert('��û��ѡ����־����!');history.back();</script>"
				    Exit Sub
				  End IF
				  If Title="" Then
				    Response.Write "<script>alert('��û��������־����!');history.back();</script>"
				    Exit Sub
				  End IF
				  if not isdate(adddate) then
				    Response.Write "<script>alert('����������ڲ���ȷ!');history.back();</script>"
				    Exit Sub
				  End IF
				  If Content="" Then
				    Response.Write "<script>alert('��û��������־����!');history.back();</script>"
				    Exit Sub
				  End IF
				Set RSObj=Server.CreateObject("Adodb.Recordset")
				RSObj.Open "Select top 1 * From KS_BlogInfo Where ID=" & KS.ChkClng(KS.S("ID")),Conn,1,3
				  RSObj("Title")=Title
				  RSObj("TypeID")=TypeID
				  RSObj("Tags")=Tags
				  RSObj("Face")=Face
				  RSObj("Weather")=weather
				  RSObj("Adddate")=adddate
				  RSObj("Content")=Content
				RSObj.Update
				RSObj.MoveLast
				Dim InfoID:InfoID=RSObj("ID")
				 RSObj.Close:Set RSObj=Nothing
				 Call KS.FileAssociation(1026,InfoID,Content,1) 
				Response.Write "<script>alert('��־�޸ĳɹ�!');location.href='KS.Spacelog.asp';</script>"
end sub

'��־���۹���
Sub commentshow()
		totalPut = Conn.Execute("Select Count(ID) from KS_BlogComment")(0)
		TotalPageNum = CInt(totalPut / maxperpage)  '�õ���ҳ��
		If TotalPageNum < totalPut / maxperpage Then TotalPageNum = TotalPageNum + 1
		If CurrentPage < 1 Then CurrentPage = 1
		If CurrentPage > TotalPageNum Then CurrentPage = TotalPageNum
%>
<table width="100%" border="0" align="center" cellspacing="1" cellpadding="1">
<tr height="25" align="center" class='sort'>
	<td width='5%' nowrap>ѡ��</td>
	<td nowrap>�� �� �� ��</td>
	<td nowrap>�� �� ��</td>
	<td nowrap>�� �� ʱ ��</td>
	<td nowrap>�� �� �� ��</td>
	<td nowrap>�� �� �� ��</td>
</tr>
<%
	Set Rs = Server.CreateObject("ADODB.Recordset")
	SQL = "select * from KS_BlogComment order by id desc"
		Rs.Open SQL, Conn, 1, 1
	If Rs.bof And Rs.EOF Then
		Response.Write "<tr class='list'><td height=""25"" align=center colspan=7>û���˷������ۣ�</td></tr>"
	Else
		If TotalPageNum > 1 then Rs.Move (CurrentPage - 1) * maxperpage
		i = 0
%>
<form name=selform method=post action=?action=commentdel>
<%
	Do While Not Rs.EOF And i < CInt(maxperpage)
		If Not Response.IsClientConnected Then Response.End
%>
<tr height="22" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'">
	<td class="splittd" align="center"><input type=checkbox name=ID value='<%=Rs("id")%>'></td>
	<td class="splittd">
	<strong>����:</strong><a href="../space/?<%=rs("username")%>/log/<%=rs("logid")%>" target="_blank"><%=Rs("title")%></a>
	<br/><strong>����:</strong><%=KS.Gottopic(KS.LoseHtml(rs("content")),150)%></td>
	<td class="splittd" align="center"><%=Rs("AnounName")%></td>
	<td class="splittd" align="center"><%=Rs("adddate")%></td>
	<td class="splittd" align="center"><%if not isnull(rs("Replay")) or rs("replay")<>"" then response.write "�ѻظ�" else response.write "<font color=red>δ�ظ�</font>"%></td>
	<td class="splittd" align="center"><a href="../space/?<%=rs("username")%>/log/<%=rs("logid")%>" target="_blank">���</a> <a href="?Action=commentdel&ID=<%=RS("ID")%>" onclick="return(confirm('ȷ��ɾ����������'));">ɾ��</a> </td>
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
	<td  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" height='25' colspan=7>
	&nbsp;&nbsp;<input id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">ȫѡ
	<input class=Button type="submit" name="Submit2" value=" ɾ��ѡ�е����� " onclick="{if(confirm('�˲��������棬ȷ��Ҫɾ��ѡ�еļ�¼��?')){this.document.selform.submit();return true;}return false;}"></td>
</tr>
</form>
<tr>
	<td  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'" colspan=7 align=right>
	<%
	 Call KS.ShowPageParamter(totalPut, MaxPerPage, "?", True, "��", CurrentPage, KS.QueryParam("page"))
	%></td>
</tr>
</table>

<%
End Sub
'ɾ������
Sub CommentDel()
 Dim ID:ID=KS.FilterIDs(KS.G("ID"))
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 Conn.execute("Delete From KS_BlogComment Where ID In("& id & ")")
 Response.Write "<script>alert('ɾ���ɹ���');location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub


'ɾ����־
Sub BlogInfoDel()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 Conn.execute("Delete From KS_BlogInfo Where ID In("& id & ")")
 Conn.Execute("Delete From KS_UploadFiles Where channelid=1026 and InfoID In(" & ID & ")")
 KS.AlertHintScript "ɾ���ɹ���"
End Sub
'��Ϊ����
Sub BlogInfoBest()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 Conn.execute("Update KS_BlogInfo Set Best=1 Where ID In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
'ȡ������
Sub BlogInfoCancelBest()
 Dim ID:ID=KS.G("ID")
 If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 Conn.execute("Update KS_BlogInfo Set Best=0 Where ID In("& id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub
Sub Verific()
 Dim ID:ID=Replace(KS.G("ID")," ","")
  If ID="" Then Response.Write "<script>alert('�Բ�����û��ѡ��!');history.back();</script>":response.end
 Conn.Execute("Update KS_BlogInfo Set status=" & KS.ChkClng(KS.G("Flag")) & " where id in(" & id & ")")
 Response.Write "<script>location.href='" & Request.servervariables("http_referer") & "';</script>"
End Sub

'��־�������
Sub classshow()
%>
<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1">
		  <tr align="center"  class="sort"> 
			<td width="87"><strong>���</strong></td>
			<td width="217"><strong>��������</strong></td>
			<td width="197"><strong>����</strong></td>
			<td width="196"><strong>�������</strong></td>
		  </tr>
		  <%dim orderid
		  set rs = conn.execute("select * from KS_BlogType order by orderid")
		    if rs.eof and rs.bof then
			  Response.Write "<tr><td colspan=""6"" height=""25"" align=""center"" class=""list"">��û������κε���־����!</td></tr>"
			else
			   do while not rs.eof%>
			  <form name="form1" method="post" action="?action=class&x=a">
				<tr  class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'"> 
				  <td class="splittd" width="87" height="20" align="center"><%=rs("typeid")%> <input name="typeid" type="hidden" id="typeid" value="<%=rs("typeid")%>"></td>
				  <td class="splittd" width="217" align="center"><input name="TypeName" type="text" class="textbox" id="TypeName" value="<%=rs("TypeName")%>" size="25"></td>
				  <td class="splittd" width="197" align="center"><input style="text-align:center" name="OrderID" type="text" class="textbox" id="OrderID" value="<%=rs("OrderID")%>" size="8">				  </td>
				  <td class="splittd" align="center"><input name="Submit" class="button" type="submit"value=" �޸� ">&nbsp;
				  <a onclick="return(confirm('ȷ��ɾ����?'))" href="?action=class&x=c&typeid=<%=rs("typeid")%>">ɾ��</a>
				  </td>
				</tr>
			  </form>
		  <%orderid=rs("orderid")
		   rs.movenext
		   loop
		 End IF
		rs.close%>
				<form action="?action=class&x=b" method="post" name="myform" id="form">
		    <tr>
		      <td class="splittd" height="25" colspan="4">&nbsp;&nbsp;<strong>&gt;&gt;������־����<<</strong></td>
		    </tr>
			<tr valign="middle" class="list"> 
			  <td class="splittd" height="25">&nbsp;</td>
			  <td class="splittd" height="25" align="center"><input name="TypeName" type="text" class="textbox" id="TypeName" size="25"></td>
			  <td class="splittd" height="25" align="center"><input style="text-align:center" name="orderid" type="text" value="<%=orderid+1%>" class="textbox" id="orderid" size="8">
			  <td class="splittd" height="25" align="center"><input name="Submit3" class="button" type="submit" value="OK,�ύ"></td>
			</tr>
		</form>
</table>
<br/><br/>

		<% Select case request("x")
		   case "a"
				conn.execute("Update KS_BlogType set TypeName='" & KS.G("TypeName") & "',orderid='" & KS.ChkClng(KS.G("OrderID")) &"' where Typeid="&KS.G("typeid")&"")
				KS.AlertHintScript "��ϲ,�����޸ĳɹ�!"
		   case "b"
		       If KS.G("TypeName")="" Then Response.Write "<script>alert('��������������!');history.back();</script>":response.end
			   
				conn.execute("Insert into KS_BlogType(TypeName,orderid)values('" & KS.G("TypeName") & "','" & KS.ChkClng(KS.G("OrderID")) &"')")
				KS.AlertHintScript "��ϲ,������ӳɹ�!"
		   case "c"
				conn.execute("Delete from KS_BlogType where Typeid="&KS.G("typeid")&"")
				KS.AlertHintScript "��ϲ,����ɾ���ɹ�!"
		End Select
		%></body>
		</html>
<%End Sub

End Class

%> 
