<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Admin_Form
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Form
        Private KS,KSCls,I
		Private MaxPerPage,CurrentPage,TotalPut,ID,RS
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
		  With KS
		  
		   If Not KS.ReturnPowerResult(0, "KSMS20008") Then          '检查权限
					 Call KS.ReturnErr(1, "")
					 Exit Sub
		   End If
		    .echo"<html>"
			.echo"<title>心情指数设置</title>"
			.echo"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			.echo"<script src=""../ks_inc/Common.js"" language=""JavaScript""></script>"
			.echo"<script src=""../ks_inc/jQuery.js"" language=""JavaScript""></script>"
			.echo"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.echo"</head>"
			.echo"<body bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"">"

			.echo"<ul id='menu_top'>"
			.echo"<li class='parent' onclick=""location.href='KS.Mood.asp?action=Add';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Go&OpStr=心情指数 >> <font color=red>添加心情指数项目</font>';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>添加心情指数项目</span></li>"
			.echo"<li class='parent' onclick='location.href=""KS.Mood.asp?action=GetCode""'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/s.gif' border='0' align='absmiddle'>调用代码</span></li>"
             If KS.G("Action")="" Then
			.echo"<li class='parent' disabled"
		     Else
			.echo"<li class='parent'"
			 End If
			.echo" onclick='location.href=""KS.Mood.asp"";'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/back.gif' border='0' align='absmiddle'>管理首页</span></li>"
			.echo"</ul>"

		  Select Case KS.G("Action")
		   Case "SetFormParam" Call SetFormParam() 
		   Case "Edit","Add"  Call FormManage()
		   Case "EditSave" Call FormSave()
		   Case "Del" Call ProjectDel()
		   Case "DelInfo" Call DelInfo()
		   Case "GetCode" Call GetCode()
		   Case "Show" Call SubmitShow()
		   Case Else Call Main()
		  End Select
		  End With
		End Sub
 
		Sub Main()
		   With KS
			.echo"<script>"
			.echo"function document.onreadystatechange(){"
			.echo"$(parent.frames['BottomFrame'].document).find('#Button1').attr('disabled',true);"
			.echo"$(parent.frames['BottomFrame'].document).find('#Button2').attr('disabled',true);"
			.echo"}</script>"
			.echo("<div style=""height:94%; overflow: auto; width:100%"" align=""center"">")
			 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			 RS.Open "Select * From KS_MoodProject Order By ID",conn,1,1
		    .echo"<table border='0' cellpadding='0' cellspacing='0'  width='100%' align='center'>"
			.echo"<tr height='25' class='sort'>"
			.echo"  <td width='50' align=center>ID</td><td align=center>项目名称</td><td align=center>状态</td><td align=center>↓操作</td>"
			.echo"</tr>"
		  Do While Not RS.Eof 
		    .echo"<tr height='23' class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
			.echo"<td align=center class='splittd'>" & RS("ID")&"</td>"
			.echo"<td align=center class='splittd'>" & RS("ProjectName") &"</td>"
			.echo"<td align=center class='splittd'>" 
			  If RS("Status")="1" Then .echo"正常" Else .echo"<font color=red>锁定</font>"
			.echo"</td>"
			.echo"<td align=center class='splittd'>"
			.echo"<a href='#' onClick=""SelectObjItem1(this,'子系统 >> <font color=red>查看详情</font>','Disabled','KS.Mood.asp?MoodID=" & rs("ID") & "&action=Show');"">查看详情</a>｜"
			

			.echo"<a href='?action=Edit&ID=" & rs("ID") & "' onclick=""$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=GoSave&OpStr=子系统 >> <font color=red>系统自定义表单</font>';"">修改</a>｜"
			 .echo"<a href='?action=Del&ID=" & rs("ID") & "' onclick='return(confirm(""此操作不可逆，确定删除吗？""))'>删除</a>｜"
			 			 
			 If RS("Status")="1" Then .echo"<a href='?Action=SetFormParam&Flag=FormOpenOrClose&ID=" & RS("ID") & "'>锁定</a>" Else .echo"<a href='?Action=SetFormParam&Flag=FormOpenOrClose&ID=" & RS("ID") & "'>开启</a>"
			
			.echo"</td></tr>"
			RS.MoveNext 
		  Loop
		    .echo"</table>"
			.echo"</div>"
		   RS.Close:Set RS=Nothing
		    .echo"</body>"
			.echo"</html>"
		  End With
		End Sub
		
		Sub ProjectDel()
		  on error resume next
		  Dim ID:ID=KS.ChkClng(KS.G("ID"))
		  Conn.BeginTrans
		  Conn.Execute("Delete From KS_UploadFiles Where ChannelID=1017 and infoid=" & ID)
		  Conn.Execute("Delete From KS_MoodProject Where ID=" & ID)
		  Conn.Execute("Delete From KS_MoodList Where MoodID=" & ID)
		  If Err<>0 Then
		   Conn.RollBackTrans
		  Else
		   Conn.CommitTrans
		  End If
		  KS.AlertHintScript "心情指数项目删除成功!" 
		End Sub
        		
		Sub GetCode()
		  Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "Select * From KS_MoodProject Where Status=1 order by ID asc",conn,1,1
		   With KS
		  	.echo"<table border='0' cellpadding='0' cellspacing='0'  width='100%' align='center'>"
			.echo"<tr height='25' class='sort'>"
			.echo" <td align=center colspan=6>各心情指数项目的前台调用代码</td>"
			.echo"</tr>"

		  Do While Not RS.Eof
			.echo"<tr height='25' class='list' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"

			.echo"<td width='50'></td><td width='140'><img src='images/37.gif'>&nbsp;<b>" & RS("ProjectName") & "</b></td><td><input type='text' value='&lt;script language=&quot;javascript&quot; type=&quot;text/javascript&quot; src=&quot;" & KS.Setting(2) & "/plus/mood.asp?id=" & rs("id") & "&c_id={$InfoID}&M_id={$ChannelID}&quot;&gt;&lt;/script&gt;' name='s" & rs(0) & "' size=60></td><td><input class=""button"" onClick=""jm_cc('s" & rs(0) & "')"" type=""button"" value=""复制到剪贴板"" name=""button""></td><td></td>"
			
			.echo"</tr>"
			.echo"<tr><td colspan=6 background='images/line.gif'></td></tr>"
		    RS.MoveNext
		  Loop
		   .echo"</table>"
		  End With
		  RS.Close:Set RS=Nothing
		  %>
		 
		   <script>
			function jm_cc(ob)
			{
				var obj=MM_findObj(ob); 
				if (obj) 
				{
					obj.select();js=obj.createTextRange();js.execCommand("Copy");}
					alert('复制成功，粘贴到你要调用的html代码里即可!');
				}
				function MM_findObj(n, d) { //v4.0
			  var p,i,x;
			  if(!d) d=document;
			  if((p=n.indexOf("?"))>0&&parent.frames.length)
			   {
				d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);
			   }
			  if(!(x=d[n])&&d.all) x=d.all[n];
			  for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
			  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
			  if(!x && document.getElementById) x=document.getElementById(n); return x;
			}
  </script>
		  <%
		End Sub
		
		Sub SetFormParam()
		   With Response
			   Dim ID:ID=KS.ChkClng(KS.G("ID"))
			   If ID=0 Then .Redirect "?": Exit Sub
			   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			   RS.Open "Select * From KS_MoodProject Where ID=" & ID,Conn,1,3
			   If RS.Eof Then
				 RS.Close:Set RS=Nothing
				.Redirect "?": Exit Sub
			   End If
		     If KS.G("Flag")="FormOpenOrClose" Then
			   If RS("Status")=1 Then 
					RS("Status")=0 
			   Else 
			    RS("Status")=1
			   end if
			 End If
			 RS.Update
			 RS.Close:Set RS=Nothing
			 ks.echo"<script>location.href='?';</script>"
		   End With
		End Sub
		
		Sub FormManage()
		Dim TimeLimit,AllowGroupID,useronce,onlyuser,ProjectContent
		Dim TempStr,SqlStr, RS, i
		Dim ProjectName,ExpiredDate,StartDate,Status,Descript,TableName,UpLoadDir
		

		Dim ID:ID = KS.ChkClng(KS.G("ID"))
	'	On Error Resume Next
	   If KS.G("Action")="Edit" Then
			SqlStr = "select * from KS_MoodProject Where ID=" & ID
			Set RS = Server.CreateObject("ADODB.recordset")
			RS.Open SqlStr, Conn, 1,1
			Status = RS("Status")
			ProjectName    = RS("ProjectName")
			ProjectContent = RS("ProjectContent")
			StartDate    = RS("StartDate")
			TimeLimit    = RS("TimeLimit")
			ExpiredDate  = RS("ExpiredDate")
			TimeLimit    = RS("TimeLimit")
            AllowGroupID = RS("AllowGroupID")
			useronce     = RS("useronce")
			onlyuser     = RS("onlyuser")
		Else
		      Status=1:TimeLimit = 0:StartDate=Now():ExpiredDate=Now()+10:AllowGroupID="":useronce=0:onlyuser=0
		End If
		
		With KS
		.echo"<html>"&_
		"<title>心情指数项目管理</title>" &_
		"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" &_
		"<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"&_
		"<script src=""../KS_Inc/jQuery.js"" language=""JavaScript""></script>"&_
		"<script src=""images/pannel/tabpane.js"" language=""JavaScript""></script>" & _
		"<link href=""images/pannel/tabpane.CSS"" rel=""stylesheet"" type=""text/css"">" & _
		"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"&_
		"<body>" &_
		"<table width='100%' border='0' cellspacing='0' cellpadding='0'>"&_
		"  <tr>"&_
		"	<td height='25' class='sort'>心情指数管理</td>"&_
		" </tr>"&_
		" <tr><td height=5></td></tr>"&_
		"</table>" & _
			
		"<div class=tab-page id=Formpanel>"& _
		"<form name=""myform"" method=""post"" action=""KS.Mood.asp?Action=EditSave&ID=" & ID & """ onSubmit=""return(CheckForm())"">" & _
        " <SCRIPT type=text/javascript>"& _
        "   var tabPane1 = new WebFXTabPane( document.getElementById( ""Formpanel"" ), 1 )"& _
        " </SCRIPT>"& _
             
		" <div class=tab-page id=site-page>"& _
		"  <H2 class=tab>基本信息</H2>"& _
		"	<SCRIPT type=text/javascript>"& _
		"				 tabPane1.addTabPage( document.getElementById( ""site-page"" ) );"& _
		"	</SCRIPT>" & _
		"<table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">"
		.echo"    <tr class='tdbg'>"
		.echo"      <td class='clefttitle'> <div align=""right""><strong>项目状态：</strong></div></td>"
		.echo"      <td height=""30""><input type=""radio"" name=""Status"" value=""1"" "
		If Status = 1 Then .echo(" checked")
		.echo">"
		.echo"正常"
		.echo"  <input type=""radio"" name=""Status"" value=""0"" "
		If Status = 0 Then .echo(" checked")
		.echo">"
		.echo"关闭</td>"
		.echo"    </tr>"

%>
		<script>
		 function CheckForm()
		 {
		  if ($("input[name=ProjectName]").val()=="")
		  {
		   $("input[name=ProjectName]").focus();
		   alert('请输入项目名称');
		   return false;
		  }
		  
		  $("form[name=myform]").submit();
		 }
		 
		 function changedate()
		 {
		   val=$("input[name=TimeLimit][checked=true]").val();
		   if (val==1){
		    $("#BeginDate").show();
		    $("#EndDate").show();		
		   }
		   else{
		    $("#BeginDate").hide();
		    $("#EndDate").hide();		
		   }
		 }
	
		</script>

		
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="CleftTitle" align="right"><div><strong>项目名称：</strong></div></td>      
			<td height="30"> <input name="ProjectName" class="textbox" type="text" value="<%=ProjectName%>" size="30"> 如：新闻心情指数项目等</td> 
		</tr>
		

		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="CleftTitle" align="right"><div><strong>选项：</strong></div><br><font color=green>最多只能15个选项，最少要有2个选项</font></td>      
			<td height="30"> 
			<%
			 Dim Order,ProjectArr,NameValue,PicValue
			 ProjectArr=Split(ProjectContent,"$$$")
			 For I=1 To 15
			  If I<10 Then Order="0" & I Else Order=I
			  If I<Ubound(ProjectArr) Then
			   NameValue=Split(ProjectArr(I-1),"|")(0)
			   PicValue=Split(ProjectArr(I-1),"|")(1)
			  End If
			  
			 .echo Order & "、名称 <input type=""text"" value=""" & NameValue & """ name=""Name" & I & """ class=""textbox""> 图片地址 <input type=""text"" name=""Pic" & I & """ value=""" & PicValue & """><br>"
			 Next
			 %>
			
			</td> 
		</tr>
		</table>
		</div>
		 <div class=tab-page id="formset">
		  <H2 class=tab>选项设置</H2>
			<SCRIPT type=text/javascript>
				 tabPane1.addTabPage( document.getElementById( "formset" ) );
			</SCRIPT>
			<table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="CleftTitle" align="right"><div><strong>启用时间限制：</strong></div></td>      
			<td height="30"> 
			
			<%
			.echo "<input type=""radio"" onclick=""changedate()"" name=""TimeLimit"" value=""1"" "
		If TimeLimit = 1 Then .echo(" checked")
		.echo">"
		.echo"启用"
		.echo"  <input type=""radio"" onclick=""changedate()"" name=""TimeLimit"" value=""0"" "
		If TimeLimit = 0 Then .echo(" checked")
		.echo">"
		.echo"不启用"
		
			%>
			</td> 
		</tr>

		<tr ID="BeginDate" style="display:none" valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">     
		<td height="30" class="clefttitle"align="right"><div><strong>生效时间：</strong></div></td>     
		<td height="30"><input name="StartDate" id='StartDate' class="textbox" type="text" value="<%=StartDate%>" size="24"><br><font color=#ff0000>日期格式：0000-00-00 00:00:00</font></td>   
		</tr> 
		
		<tr ID="EndDate" style="display:none" valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="CleftTitle" align="right"><div><strong>失效时间：</strong></div></td>      
			<td height="30"> <input name="ExpiredDate" id="ExpiredDate" class="textbox" type="text" value="<%=ExpiredDate%>" size="30"><br><font color=#ff0000>日期格式：0000-00-00 00:00:00</font></td> 
		</tr>
		
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="CleftTitle" align="right"><div><strong>只允许会员表态：</strong></div></td>      
			<td height="30"> 
			
			<%
			.echo "<input type=""radio"" name=""onlyuser"" value=""1"" "
		If onlyuser = 1 Then .echo(" checked")
		.echo">"
		.echo"是"
		.echo"  <input type=""radio"" name=""onlyuser"" value=""0"" "
		If onlyuser = 0 Then .echo(" checked")
		.echo">"
		.echo"不是"
		
			%>
			</td> 
		</tr>
		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="CleftTitle" align="right"><div><strong>每个会员只允许表态一次：</strong></div></td>      
			<td height="30"> 
			
			<%
			.echo "<input type=""radio"" name=""useronce"" value=""1"" "
		If useronce = 1 Then .echo(" checked")
		.echo">"
		.echo"是"
		.echo"  <input type=""radio"" name=""useronce"" value=""0"" "
		If useronce = 0 Then .echo(" checked")
		.echo">"
		.echo"不是"
		
			%>
			</td> 
		</tr>
		
		

		<tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">      
			<td width="160" height="30" class="CleftTitle" align="right"><div><strong>用户组限制：</strong></div><font color=#ff0000>不限制，请不要选</font></td>      
			<td height="30"><%=KS.GetUserGroup_CheckBox("AllowGroupID",AllowGroupID,5)%> </td> 
		</tr>
				
			</table>
        </div>
		<script>changedate();</script>
		<%
		.echo"</form>"
		.echo"</div>"
		End With
		End Sub
		
		
		
		
	

		
		Sub FormSave()
		    Dim ExpiredDate,StartDate,I,OpName,ID,ProjectContent
			ID=KS.ChkClng(KS.G("ID"))
			StartDate=KS.G("StartDate")
			ExpiredDate=KS.G("ExpiredDate")
			If Not IsDate(StartDate) Then Call KS.AlertHistory("生效日期格式不正确",-1):response.end
			If Not IsDate(ExpiredDate) Then Call KS.AlertHistory("失效日期格式不正确",-1):response.end
			If ID=0 and Not Conn.Execute("select top 1 id from ks_moodproject where projectname='" & KS.G("ProjectName") &"'").eof then Call KS.AlertHistory("项目名称已存在！",-1):response.end
			
			For I=1 To 15
			 If ProjectContent="" Then
			 ProjectContent=Request("Name" & I) & "|" & Request("Pic" & I)
			 Else
			 ProjectContent=ProjectContent & "$$$" & Request("Name" & I) & "|" & Request("Pic" & I)
			 End If
			Next
			
			on error resume next
			Conn.BeginTrans
		    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select * From KS_MoodProject Where ID=" & ID,Conn,1,3
			If  RS.Eof And RS.Bof Then
			    RS.AddNew
				OpName      = "添加"
			Else
			    OpName="修改"
			End If
				RS("ProjectName")= KS.G("ProjectName")
				RS("ProjectContent")=ProjectContent
				RS("Status") = KS.G("Status")
				RS("TimeLimit")   = KS.ChkClng(KS.G("TimeLimit"))
				RS("StartDate")     = startdate
				RS("ExpiredDate")    = ExpiredDate
				RS("useronce") =KS.ChkClng(KS.G("useronce"))
				RS("onlyuser")=KS.ChkClng(KS.G("onlyuser"))
				RS("AllowGroupID")     = KS.G("AllowGroupID")
				RS.Update
				If ID=0 Then
				 Call KS.FileAssociation(1017,RS("ID"),ProjectContent,0)
				Else
				 Call KS.FileAssociation(1017,ID,ProjectContent,1)
				End If
				RS.Close
				Set RS=Nothing
				
				
				if err<>0 then
					Conn.RollBackTrans
					Call KS.AlertHistory("出错！出错描述：" & replace(err.description,"'","\'"),-1):response.end
				else
					Conn.CommitTrans
					If ID=0 Then
					 KS.Alert OpName & "心情指数项目添加成功!","KS.Mood.asp"
					Else
					 KS.Alert OpName & "心情指数项目修改成功!","KS.Mood.asp"
					End If
				end if
		End Sub
		
		Sub SubmitShow()
		Dim MoodID:MoodID=KS.ChkClng(KS.G("MoodID"))
		MaxPerPage = 20     '取得每页显示数量
		If KS.G("page") <> "" Then
			  CurrentPage = KS.ChkClng(KS.G("page"))
		Else
			  CurrentPage = 1
		End If
		 with KS
			set rs=server.createobject("adodb.recordset")
			RS.Open "Select * From KS_MoodProject Where ID=" & MoodID,conn,1,1
			If RS.Eof Then
			 RS.Close
			 Set RS=Nothing
			 KS.AlertHintScritp "出错!"
			End If
			Dim ProjectContent,VoteArr,VoteItemArr,I,VoteTitle(15)
			ProjectContent=RS("ProjectContent")
			VoteArr=Split(ProjectContent,"$$$")
			For I=0 To Ubound(VoteArr)
			 If VoteArr(i)<>"" Then
			   VoteItemArr=Split(VoteArr(i),"|")
			   If VoteItemArr(0)<>"" And VoteItemArr(1)<>"" Then
			    VoteTitle(i)=VoteItemArr(0)
			   End If
			 End If
			Next
			RS.Close
			

		    .echo("<div style=""height:94%; overflow: auto; width:100%"" align=""center"">")
		 	.echo"<table border='0' cellpadding='0' cellspacing='0'  width='100%' align='center'>"
			.echo"<tr height='25' class='sort'>"
			.echo"  <td width='40' align='center'>ID</td><td align=center>信息标题</td>"
			For i=0 to Ubound(VoteTitle)
			 If VoteTitle(i)<>"" Then
			 .echo"<td align=""center"">" & VoteTitle(i) & "</td>"
			 End If
			Next
			.echo"</tr>"
			 rs.open "select * from KS_MoodList Where MoodID=" & MoodID & " order by ID desc" ,conn,1,1
			 If RS.EOF Then
			  .echo"<tr><td colspan='18' align='center'>没有记录!</td></tr>"
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
		
							If CurrentPage <> 1 Then
								If (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
								Else
									CurrentPage = 1
								End If
							End If
							Dim n:n=0
							.echo"<form name=selform method=post action=""KS.Mood.asp"">"
							Do While Not RS.Eof
								.echo"<tr height=""22"" class='list' onMouseOver=""this.className='listmouseover'"" onMouseOut=""this.className='list'"">"
								.echo"<td align='center'><input type='checkbox' name='id' value='" & RS("ID") & "'></td>"
								.echo"<td><a href='../item/show.asp?m=" & rs("channelid") & "&d=" & RS("InfoID") & "' target='_blank'>" & RS("Title") & "</a></td>"
								For i=0 to Ubound(VoteTitle)
								 If VoteTitle(i)<>"" Then
								 .echo"<td align=""center"">" & rs("m"&i) & "</td>"
								 End If
								Next
								.echo"</tr>"
								.echo"<tr><td colspan=18 background='images/line.gif'></td></tr>"
								n=n+1
								If N>=MaxPerPage Then Exit Do
							    RS.MoveNext
							Loop
			 End If
			
			  .echo("<tr> ")
			  .echo("<td colspan='2'>&nbsp;&nbsp;<input id=""chkAll"" onClick=""CheckAll(this.form)"" type=""checkbox"" value=""checkbox""  name=""chkAll"">全选&nbsp;&nbsp;<input type='hidden' value='DelInfo' name='Action'><input type='submit' value='删除选中' class='button' onclick=""return(confirm('确定删除选择的信息吗?'))"">")
			  .echo("</form><td height=""50"" colspan=""18""  align=""right"">")
			  Call KSCLS.ShowPage(totalPut, MaxPerPage, "KS.Mood.asp", True, "条",CurrentPage, KS.QueryParam("page"))
			  .echo("<br></td>")
			  .echo("</tr>")			
			  .echo"</table>"
			  .echo"</div>"
         end with
		End Sub
       
	    Sub DelInfo()
		  Conn.Execute("Delete From KS_MoodList Where ID in(" & KS.FilterIds(KS.G("ID")) & ")")
		  Response.Redirect Request.ServerVariables("http_referer")
		End Sub
		
		
End Class
%> 
