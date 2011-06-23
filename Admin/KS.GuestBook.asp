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
Set KSCls = New Guest_Manage
KSCls.Kesion()
Set KSCls = Nothing

Class Guest_Manage
        Private KS,Action,KSCls
	    Private MaxPerPage, TotalPut , CurrentPage, TotalPage, i, j, Loopno
	    Private KeyWord, SearchType,SqlStr,RS
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
	KeyWord = KS.R(Trim(Request("keyword")))
	SearchType = KS.R(Trim(Request("SearchType")))
	Action=KS.G("Action")
	Select Case Action
	 Case "Main"  Call GuestMain()
	 Case "Del"  Call GuestDel()
	 Case "Reply" Call Reply()
	 Case Else  Call MainList()
	 End Select
	End Sub
	Sub GuestMain()
			 With Response
			If Not KS.ReturnPowerResult(0, "KSMS20004") Then
			  Call KS.ReturnErr(1, "")
			  exit sub
			End If
				.Write "<html>"
				.Write"<head>"
				.Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
				.Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
				.Write"</head>"
				.Write"<body scroll=no leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
				.Write"<div class=""topdashed sort"">��վ���Թ���</div>"
				.Write "<table border='0' width='100%' height='100%'>"
				.Write  "<tr>"
				.Write " <td> <iframe scrolling=""auto"" frameborder=""0"" src=""KS.GuestBook.asp"" width=""100%"" height=""100%""></iframe>"
				.Write"</td>"
				.Write " </tr>"
				.Write"</TABLE>"
			End With
	End Sub

Sub MainList()
%>
<html>
<head>
<title>�������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="Include/admin_Style.css" type="text/css">
<script language="JavaScript">
<!--
function CheckSelect()
{
	var count=0;
	for(i=0;i<document.KS_GuestBook.elements.length;i++)
	{
		if(document.KS_GuestBook.elements[i].name=="GuestID")
		{		
			if(document.KS_GuestBook.elements[i].checked==true)
			{
				count++;					
			}				
		}			
	}
		
	if(count<=0)
	{
		alert("��ѡ��һ��Ҫ��������Ϣ��");
		return false;
	}

	return true;
}

function cdel()
{
	if(CheckSelect()==false)
	{
		return false;
	}
	
	if (confirm("�����Ҫɾ���������Լ�¼�𣿲��ɻָ���")){
		document.KS_GuestBook.Flag.value = "del";
		document.KS_GuestBook.submit();
	}
}

function ccheck()
{
	if(CheckSelect()==false)
	{
		return false;
	}
	
	if (confirm("��ȷ��Ҫ�����Щ��Ϣ��")){
		document.KS_GuestBook.Flag.value = "check";
		document.KS_GuestBook.submit();
	}
}

function cuncheck()
{
	if(CheckSelect()==false)
	{
		return false;
	}
	
	if (confirm("��ȷ��Ҫ������Щ��Ϣ������߽���������Щ��Ϣ��")){
		document.KS_GuestBook.Flag.value = "uncheck";
		document.KS_GuestBook.submit();
	}
}

function SelectCheckBox()
{
	for(i=0;i<document.KS_GuestBook.elements.length;i++)
	{
		if(document.all("selectCheck").checked == true)
		{
			document.KS_GuestBook.elements[i].checked = true;					
		}
		else
		{
			document.KS_GuestBook.elements[i].checked = false;
		}
	}
}
//-->
</script>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="TableBar">
      <form action="KS.GuestBook.asp" method="post" name="search" id="search">
        <tr>
          <td height="25">�������� --&gt;&gt;&gt; �ؼ��ʣ�
            <input type="text" name="keyword" class="inputtext" size="35" value="<%=KeyWord%>" onMouseOver="this.focus()" onFocus="this.select()">
                <select name="SearchType" size="1" class="inputlist">
                  <option value="content" <%If SearchType = "content" Then Response.Write "selected"%>>��������</option>
                  <option value="author" <%If SearchType = "author" Then Response.Write "selected"%>>�� �� ��</option>
                </select>
                <input type="submit" name="imageField" value="����"></td>
        </tr>
      </form>
    </table>
<table border="0" width="100%" align="center" style='border-top:1px solid #cccccc' cellpadding="0" cellspacing="0">
	<form name="KS_GuestBook" action="KS.GuestBook.asp?Action=Del" method=post>
	<input name="Flag" type="hidden" value="" id="Flag">
		<tr class="sort">
					<td width="45">ѡ��</td>
					<td>��������</td>
					<td>������</td>
					<td>�ظ�/�鿴</td>
					<td>��󷢱�</td>
					<td>״̬</td>
					
					<td>�������</td>
				</tr>
			    <%
	Set RS=Server.CreateObject("ADODB.RECORDSET")
	If SearchType = "content" Then
		SqlStr = "SELECT * FROM KS_GuestBook WHERE Subject LIKE '%"& KeyWord &"%' ORDER BY ID DESC"  
	Else
		SqlStr = "SELECT * FROM KS_GuestBook WHERE UserName LIKE '%"& KeyWord &"%' ORDER BY ID DESC" 
	End If
	RS.Open SqlStr,Conn,1,1 

	Dim Pmcount
	Pmcount = 15
	If KS.ChkClng(Pmcount) < 1 Then Pmcount = 10

	RS.Pagesize = Pmcount
	TotalPut = RS.RecordCount	'��¼���� 
	TotalPage = RS.PageCount	'�õ���ҳ��
	MaxPerPage = RS.PageSize	'����ÿҳ��
		
	CurrentPage = KS.ChkClng(Request("Page"))
	
	If CDbl(CurrentPage) < 1 Then CurrentPage = 1
	If CDbl(CurrentPage) > CDbl(TotalPage) Then CurrentPage = TotalPage

	If RS.Eof or RS.Bof Then 
		Response.Write "<tr><td colspan='10' align='center' height='30'><font color=#FF0000>��ʱ��û���κ����ԣ�</font></td></tr>"
	Else
		RS.Absolutepage = CurrentPage	'��ָ������ָ��ҳ�ĵ�һ����¼
		Loopno = MaxPerPage
		
		i = 0
		Do While Not RS.Eof and Loopno > 0
%>
        <tr>
          <td  height="30" class='splittd' width="25" align="center" valign="middle"><input type="checkbox" name="GuestID" value="<%=Trim(RS("ID"))%>"></td>
		 <td class='splittd'><img src="../club/images/common.gif" align="absmiddle">
		  
		 <% on error resume next
		   response.write "[<a href='../club/index.asp?boardid=" & rs("boardid") & "' target='_blank'>" & conn.execute("select boardname from ks_guestboard where id=" & rs("boardid"))(0) & "</a>]"
		 if KS.Setting(59)="1" Then
		  response.write "<a href='?action=Reply&guestid=" & rs("id") & "'>"
		  else
		  %>
		 <a href="../club/display.asp?id=<%=rs("id")%>" target="_blank">
		 <%end if%><%=rs("subject")%></a></td>
		 <td class='splittd'>
		 <%
		 if ks.isnul(rs("username")) then 
		  response.write "�ο�"
		 else
		  response.write rs("username")
		 end if
		 %>
		 </td>
		 <td class='splittd' align="center">
		 <%
		 if KS.Setting(59)="1" Then
			  if conn.execute("select top 1 id from ks_guestreply where topicid=" & rs("id")).eof then
			   response.write "<font color=red>δ�ظ�</font>"
			  else
			   response.write "<font color=green>�ѻظ�</font>"
			  end if
		 else
		  response.write RS("TotalReplay") & "/" & rs("hits")
		 end if
		 %>
		 </td>
		 <td class='splittd'>
		 <%
		 if ks.isnul(RS("LastReplayUser")) then 
		  response.write "�ο�"
		 else
		  response.write RS("LastReplayUser")
		 end if
		 %>
		 </td>
		 <td class='splittd' align='center'>
		 <%
		  If rs("verific")=1 then
		   response.write "<font color=blue>�����</font>"
		  else
		   response.write "<font color=red>δ���</font>"
		  end if
		 %>
		 </td>

		 <td class='splittd' align="center">
		  <%
		 if KS.Setting(59)="1" Then
		   response.write "<a href='?action=Reply&guestid="& rs("id") & "'>�ظ�/�޸�</a>  | "
		 end if
		   %>

		 <%If rs("verific")=0 then%>
		 <a href="?Action=Del&flag=check&guestid=<%=rs("id")%>">���</a>
		 <%else%>
		 <a href="?Action=Del&flag=uncheck&guestid=<%=rs("id")%>">ȡ��</a>
		 <%end if%> | 
		 

		 <a href="?Action=Del&flag=del&guestid=<%=rs("id")%>" onClick="return(confirm('���и������µĻظ�Ҳ����ɾ����ȷ����'))">ɾ��</a> | <a href="../club/display.asp?id=<%=rs("id")%>" target="_blank">�鿴</a> 
		 
		 </td>
		</tr>
        <%
	RS.MoveNext
	Loopno = Loopno-1
	i = i+1
	Loop
	%>
</form>
	</table>
	<%
End if
RS.Close
Set RS=Nothing
%>
        <table border="1" width="100%" cellspacing="0" cellpadding="2"  align="center" bgcolor="#F5F5F5"  bordercolordark="#FFFFFF" bordercolorlight="#DDDDDD">
          <tr>
		    <td width="25" align="center"><input type="checkbox"  name='selectCheck' onClick="SelectCheckBox()"></td>
            <td width="240">ȫ��ѡ��
              <input name="delbtn" value="ɾ��"  class="button" type="button" onClick="cdel();">
			  <input name="delbtn" value="���" class="button" type="button" onClick="ccheck();">
	          <input name="delbtn" value="ȡ�����" class="button" type="button" onClick="cuncheck();">
			</td>
              <td align="center">
			<%
	       '��ʾ��ҳ��Ϣ
			  Call KSCLS.ShowPage(totalPut, MaxPerPage, "KS.GuestBook.asp", True, "��", CurrentPage, "keyword=" & keyword & "&searchtype=" & SearchType)
	      %>

          </tr>
      </table>
<br>
<br>
<br>
<%
 End Sub
 
 'ɾ������
 Sub GuestDel()
			Dim strIdList,arrIdList,iId,i,Flag,SqlStr
			strIdList = Trim(KS.G("GuestID"))
			Flag = Trim(KS.G("Flag"))
			Select Case Flag
			Case "del"
				If Not IsEmpty(strIdList) Then
					arrIdList = Split(strIdList,",")
				
					For i = 0 To UBound(arrIdList)
						iId = Clng(arrIdList(i))			
						SqlStr = "DELETE FROM KS_GuestBook WHERE ID=" & iId
						Conn.Execute SqlStr	
						Conn.Execute("Delete FROM KS_GuestReply Where TopicID=" & iId)		
					Next
					Call KS.Alert("��Ϣɾ���ɹ���ȷ�Ϸ��أ�",Request.ServerVariables("HTTP_REFERER"))
				Else
					Call KS.AlertHistory("������ѡ��һ����Ϣ��¼��",-1)
				End If
			Case "check"
				If Not IsEmpty(strIdList) Then
					arrIdList = Split(strIdList,",")
				
					For i = 0 To UBound(arrIdList)
						iId = Clng(arrIdList(i))			
						Conn.Execute("UPDATE KS_GuestBook SET Verific = 1 WHERE ID="&iId&"")			
					Next
					Call KS.Alert("��Ϣ��˳ɹ���ȷ�Ϸ��أ�",Request.ServerVariables("HTTP_REFERER"))
				Else
					Call KS.AlertHistory("������ѡ��һ����Ϣ��¼��",-1)
				End If
				Case "uncheck"
					If Not IsEmpty(strIdList) Then
						arrIdList = Split(strIdList,",")
					
						For i = 0 To UBound(arrIdList)
							iId = KS.ChkClng(arrIdList(i))			
							Conn.Execute("UPDATE KS_GuestBook SET Verific = 0 WHERE ID="&iId&"")
						Next
						Call KS.Alert("��Ϣ�����ɹ���ȷ�Ϸ��أ�",Request.ServerVariables("HTTP_REFERER"))
					Else
						Call KS.AlertHistory("������ѡ��һ����Ϣ��¼��",-1)
					End If
				End Select
	End Sub
	
	Sub Reply()
	Dim Flag, pagetxt, guestid, ssubject, sanser, sadminhead, scheckbox, sansertime,SqlStr,RSObj
			Dim DomainStr:DomainStr= KS.GetDomain
			Flag =KS.G("Flag")
			pagetxt = Request("cpage")
			guestid = KS.ChkClng(Request("guestid"))
			if Flag="ok" then
			   ssubject =KS.G("txtcontop")   
			   sanser = KS.G("txtanser")
			   sadminhead = KS.G("adminhead")
			   scheckbox = KS.G("htmlok")
			   sansertime = Now()
			   set rsobj=server.createobject("adodb.recordset")
			   rsobj.open "select top 1 [memo] from ks_guestbook where id=" & guestid,conn,1,3
			   if rsobj.eof and rsobj.bof then
			    response.write "error!"
				response.End()
			   end if
			    rsobj(0)=request.Form("content")
				rsobj.update
				rsobj.close
			   rsobj.open "select top 1 * from ks_guestreply where topicid=" & guestid,conn,1,3
			   if rsobj.eof and rsobj.bof then
			    rsobj.addnew
			   end if
			   If sanser="" Then
			    rsobj.delete
			   Else
			    rsobj("username")=KS.C("AdminName")
				rsobj("userip")=KS.GetIP()
				rsobj("TopicID")=guestid
				rsobj("content")=sanser
				rsobj("ReplayTime")=now()
				rsobj("txthead")=sadminhead
				rsobj("Verific")=1
				rsobj.update
			  End If
			    rsobj.close:set rsobj=nothing
			   Response.write "<script>alert('��ϲ�����Իظ��ɹ���');location.href='KS.Guestbook.asp?page=" &pagetxt& "';</script>"
			End If
                Set RSObj=Server.CreateObject("Adodb.Recordset")
				SqlStr="SELECT * FROM KS_GuestBook WHERE ID="&guestid
				RSObj.Open SqlStr,Conn,1,1
			%>
			<html>
			<head>
			<title>�������</title>
			<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
			<link rel="stylesheet" href="Include/admin_Style.css" type="text/css">
			<br>
			<table width="540" border="0" cellspacing="0" cellpadding="0" align="center">
			  <form method="POST" action="KS.GuestBook.asp?Action=Reply&guestid=<%Response.Write guestid%>&amp;cpage=<%Response.Write pagetxt%>" name="repleBook">
				<tr>
				  <td valign="top"> <br>
					  <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
						<tr> 
						  <td colspan="2" align="center" height="14">:::::::::::::::::::::::::::::::::::: �� �� �� �� ::::::::::::::::::::::::::::::::::::</td>
						</tr>
						<tr> 
						  <td width="18%" align="center" height="32"><img src="<%=DomainStr%>Images/face/<%=RSObj("Face")%>"><br><%=RSObj("UserName")%></td>
						  <td>
						  <%			
						  Response.Write "<textarea  id=""content"" name=""content""  style=""display:none"">"& RSObj("Memo") &"</textarea><input type=""hidden"" id=""content___Config"" value="""" style=""display:none"" /><iframe id=""content___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=content&amp;Toolbar=Basic"" width=""520"" height=""160"" frameborder=""0"" scrolling=""no""></iframe>"

						  %>
						  </td>
						</tr>
					  </table>
					<table width="100%" border="0" cellspacing="0" cellpadding="0" height="150" class="font" align="center">
					  <tr> 
						<td > </td>
					  </tr>
					  <tr> 
						<td nowrap align="center">:::::::::::::::::::::::::::::::::::: վ �� �� �� ::::::::::::::::::::::::::::::::::::</td>
					  </tr>
					  <tr> 
						<td nowrap align="center"  height="135" valign="middle" style="padding-left:60px"> 
						  <p> 
						  <%
						  dim replycontent,TxtHead
						  dim rs:set rs=server.createobject("adodb.recordset")
						  rs.open "select Content,txthead from KS_GuestReply where TopicID=" & guestid,conn,1,1
						  if rs.eof then
						   replycontent=" "
						   TxtHead=1
						  else
						   replycontent=rs(0)
						   TxtHead=rs(1)
						  end if
						  rs.close:set rs=nothing%>
							<textarea rows="8" name="txtanser" cols="70" class="inputmultiline"><%=KS.HTMLCode(replycontent)%></textarea>
						</td>
					  </tr>
					</table>
					  
					<div align="center">
					  <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
						<tr valign="bottom">
						  <td nowrap="nowrap" colspan="16" class="font"><div align="center">::::::::::::::::::::::::::::::::::::: ѡ �� 
							�� �� :::::::::::::::::::::::::::::::::::::</div></td>
						</tr>
						<tr height="25" align="center">
						  <td colspan="16"><%Dim I
							For I=1 To 30
							   Response.Write "<input type=""radio"" name=""Adminhead"" value=""" & I & """"
							   IF I =TxtHead or i=1 Then Response.Write(" Checked")
							  Response.Write" ><img src=""" & DomainStr & "Images/Face1/Face" & I & ".gif"" border=""0"">"
							  IF I Mod 15=0 Then Response.Write("<BR>")
							  
							 Next
					
					
%></td>
					    </tr>
					  </table>
					  <table width="100%" border="0" cellspacing="0" cellpadding="0" class="font">
						<tr>
						  <td align="center"><font color="#400040">......................................................................................</font></td>
						</tr>
					  </table>
					  <table width="530" border="0" cellspacing="0" cellpadding="0" class="font">
						<tr>
						  <td height="35" align="center"> 
							  <input type="submit" value=" ȷ �� " name="cmdOk" class="inputbutton">
							  &nbsp; 
							  <input type="reset" value=" �� �� " name="cmdReset" class="inputbutton">
							  &nbsp; 
							  <input type="button" value=" �� �� " name="cmdExit" class="inputbutton" onClick=" history.back()">
						  <input type="hidden" name="Flag" value="ok"></td>
						</tr>
					  </table>
					</div>
					</td>
				</tr>
			  </form>
			</table>
			<p>&nbsp;</p>
			<%
	End Sub
End Class
%>
 
