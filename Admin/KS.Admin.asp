<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../Plus/md5.asp"-->
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
Set KSCls = New User_AdminMain
KSCls.Kesion()
Set KSCls = Nothing

Class User_AdminMain
        Private KS,UserName
		Private GroupID, I, SqlStr, RSObj,Title, CreateDate, TempStr, GRS,KeyWord, SearchType
		Private PowerRS,RS,AdminID,PowerList,SpecialPower,CollectPower,SystemPower,RefreshPower,UserAdminPower,KMTemplatePower,ModelPower

		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		 Select Case KS.G("Action")
		   Case "Add","Edit"
			 If Not KS.ReturnPowerResult(0, "KMUA10001") Then                '检查管理员操作(增和改)的权限检查
			  Call KS.ReturnErr(0, "")
			  Exit Sub
			 Else 
			  Call AdminAdd()
		     End If
		   Case "Del"
		       If Not KS.ReturnPowerResult(0, "KMUA10001") Then                '检查管理员操作(增和改)的权限检查
				  Call KS.ReturnErr(0, "")
				  Exit Sub
			   Else
			   Call AdminDel()
		       End If
		   Case "SetPass"
		   	 If Not KS.ReturnPowerResult(0, "KMUA10010") Then           
			  Call KS.ReturnErr(0, "")
			  Exit Sub
		     Else
		      Call SetAdminPass()
			 End If
		   Case Else
		     Call AdminList()
		 End Select
		End Sub
		
		Sub AdminList()
		  
		'收集搜索参数
		KeyWord = KS.G("KeyWord")
		SearchType = KS.G("SearchType")
		'搜索参数集合
		Dim SearchParam:SearchParam = "KeyWord=" & KeyWord & "&SearchType=" & SearchType
		Const Row = 8 '每行显示数
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; chaRSet=gb2312"">"
		Response.Write "<title>管理员管理</title>"
		Response.Write "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
		Response.Write "<script language=""JavaScript"">" & vbCrLf
		Response.Write "var GroupID='0';        //管理员组ID" & vbCrLf
		Response.Write "var KeyWord='" & KeyWord & "';         //搜索关键字" & vbCrLf
		Response.Write "var SearchParam='" & SearchParam & "'; //搜索参数集合" & vbCrLf
		Response.Write "</script>" & vbCrLf
		Response.Write "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>"
		Response.Write "<script language=""JavaScript"" src=""../KS_Inc/jquery.js""></script>"
		Response.Write "<script language=""JavaScript"" src=""Include/ContextMenu1.js""></script>"
		Response.Write "<script language=""JavaScript"" src=""Include/SelectElement.js""></script>"
		%>
		<script language="javascript">
		var DocElementArrInitialFlag=false;
		var DocElementArr = new Array();
		var DocMenuArr=new Array();
		var SelectedFile='',SelectedFolder='';
		function document.onreadystatechange()
		{   
		    parent.frames['BottomFrame'].Button1.disabled=true;
			parent.frames['BottomFrame'].Button2.disabled=true;
		    if (DocElementArrInitialFlag) return;
			InitialDocElementArr('GroupID','AdminID');
			InitialDocMenuArr();
			DocElementArrInitialFlag=true;
		}
		function InitialDocMenuArr()
		{      DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.Create();",'添 加(N)','disabled');
			   DocMenuArr[DocMenuArr.length]=new ContextMenuItem("seperator",'','');
			   DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.Edit();",'编 辑(E)','disabled');
			   DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.Delete('');",'删 除(D)','disabled');
			   DocMenuArr[DocMenuArr.length]=new ContextMenuItem("seperator",'','');
			   DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.SetAdminPassWord();",'设置密码(P)','disabled');
			   DocMenuArr[DocMenuArr.length]=new ContextMenuItem("seperator",'','');
			   DocMenuArr[DocMenuArr.length]=new ContextMenuItem("parent.Reload();",'刷 新(Z)','');
		}
		function DocDisabledContextMenu()
		{
		   var TempDisabledStr=''; 
			DisabledContextMenu('GroupID','AdminID',TempDisabledStr+'设置权限(S),设置密码(P),编 辑(E),删 除(D)','设置权限(S),设置密码(P),编 辑(E)','','设置权限(S),设置密码(P),编 辑(E)','设置权限(S),设置密码(P),编 辑(E)','')
		}
		function CreateAdmin()
		{
		 location.href='KS.Admin.asp?Action=Add';
		 $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=Go&OpStr=用户管理 >> <font color=red>添加管理员</font>';
		}
		function EditAdmin(AdminID)
		{
		 location.href='KS.Admin.asp?Action=Edit&AdminID='+AdminID;
		 $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?ButtonSymbol=GoSave&OpStr=用户管理 >> <font color=red>修改管理员</font>';
		}
		function Create()
		{
		 CreateAdmin();  
		}
		function Edit()
		{  GetSelectStatus('GroupID','AdminID');
				if (SelectedFile!='')
				{
				 if (SelectedFile.indexOf(',')==-1) 
				  EditAdmin(SelectedFile);
				 else alert('一次只能够编辑一个管理员!'); 
				}
			else 
			   alert('请选择要编辑的管理员!');
		}
		function Delete(id)
		{  
		 if (id==''){
		  GetSelectStatus('GroupID','AdminID');  
		 }else{
		  SelectedFile=id;
		 }
			if (SelectedFile!='')
				{ if (confirm('确定删除选中管理员吗?'))location='KS.Admin.asp?'+SearchParam+'&Action=Del&AdminID='+SelectedFile;}
				else alert('请选择要删除的管理员');
		}
		function SetAdminPassWord()
		{
		 GetSelectStatus('GroupID','AdminID');
		 if (SelectedFile!='')
				   if (SelectedFile.indexOf(',')==-1) 
					 { 
					 OpenWindow('KS.Frame.asp?Url=KS.Admin.asp&Action=SetPass&PageTitle='+escape('设置管理员密码')+'&AdminID='+SelectedFile,360,160,window);
					 SelectedFile='';
					 }
				 else alert('一次只能给一个管理员设置密码!'); 
		 else
		  alert('请选择要设置密码的管理员!')
		}
		function GetKeyDown()
		{
		if (event.ctrlKey)
		  switch  (event.keyCode)
		  {  case 90 :  Reload(); break;
			 case 78 : event.keyCode=0;event.returnValue=false;
				 CreateAdmin('');
			 case 80 :SetAdminPassWord();break;
			 case 69 : event.keyCode=0;event.returnValue=false;Edit(); break;
			 case 68 : Delete('');break;
			 case 70 : event.keyCode=0;event.returnValue=false;
				parent.frames['LeftFrame'].initializeSearch('Manager')
		   }	
		else	
		 if (event.keyCode==46)Delete('');
		}
		function Reload()
		{
		location.href='KS.Admin.asp?'+SearchParam+'&GroupID='+GroupID;
		}
		</script>
		<%
		Response.Write "</head>"
		Response.Write "<body scroll=no topmargin=""0"" leftmargin=""0"" OnClick=""SelectElement();"" onkeydown=""GetKeyDown();"" onselectstart=""return false;"">"
		Response.Write "<ul id='menu_top'>"
			  If KeyWord = "" Then
			   Response.Write "<li class='parent' onclick=""CreateAdmin();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>添加管理员</span></li>"				 
			   Response.Write "<li class='parent' onclick=""SetAdminPassWord();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/unverify.gif' border='0' align='absmiddle'>设置密码</span></li>"				 
			   Response.Write "<li class='parent' onclick=""parent.frames['LeftFrame'].initializeSearch('管理员');""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/addjs.gif' border='0' align='absmiddle'>搜索管理员</span></li>"				 

			  Else
				   Response.Write ("<img src='Images/home.gif' align='absmiddle'><span style='cursor:pointer' onclick=""SendFrameInfo('KS.Admin.asp','Admin_UserLeft.asp','KS.Split.asp?ButtonSymbol=Disabled&OpStr=管理员管理 >> <font color=red>管理首页</font>')"">管理员首页</span>")
			   Response.Write (">>> 搜索结果: ")
				Select Case SearchType
				 Case 0
				  Response.Write ("用户名含有 <font color=red>" & KeyWord & "</font> 的管理员")
				 Case 1
				  Response.Write ("管理员简介含有 <font color=red>" & KeyWord & "</font> 的管理员")
				 End Select
			   End If

		Response.Write "</ul>"
		Response.Write ("<div style=""height:94%; overflow: auto; width:100%"" align=""center"">")
		Response.Write "<table width=""100%"" height=""25"" border=""0"" cellpadding=""0"" cellspacing=""1"">"
		
			  Dim Param:Param = " Where 1=1"
			  If KeyWord <> "" Then
				Select Case SearchType
				  Case 0
				   Param = Param & " And UserName like '%" & KeyWord & "%'"
				  Case 1
				   Param = Param & " And Description like '%" & KeyWord & "%'"
				End Select
			   Else
				 Param = Param
			   End If
			  Param = Param & " Order BY SuperTF Desc,AddDate desc"
			  SqlStr = "Select * From KS_Admin " & Param
				 Response.Write ("<tr> ")
				 Response.Write ("  <td>")
				 Response.Write ("    <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">")
				 Response.Write ("<tr align=""center""><td height=23 width=100  class=""sort"">管 理 员</td><td width=120  class=""sort"">类 型</td><td class=""sort"">最后登录IP</td><td class=""sort"">最后登录时间</td><td  class=""sort"">最后登出时间</td><td  class=""sort"">登录次数</td><td class=""sort"">锁 定</td><td class='sort'>管理操作</td></tr>")
		Set RSObj = Server.CreateObject("AdoDb.RecordSet")
		 RSObj.Open SqlStr, Conn, 1, 1
		 Dim T, TitleStr, LockStr, ShortName
		Do While Not RSObj.EOF
			
				If RSObj("Locked") = 1 Then
					LockStr = "<font color=red>已锁定</font>"
					Else
					LockStr = "<font color=green>正常</font>"
				End If
				TitleStr = " TITLE='姓 名:" & RSObj("RealName") & "&#13;&#10;性 别:" & RSObj("Sex") & "&#13;&#10;添加时间:" & RSObj("AddDate") & "&#13;&#10;简要描述:" & RSObj("Description") & "'"
			  Response.Write ("<tr><td class='splittd' height=25" & TitleStr & ">&nbsp;<span ondblclick=""EditAdmin(this.AdminID);"" AdminID=""" & RSObj("AdminID") & """><img src=Images/Folder/Admin" & Trim(CStr(RSObj("SuperTF"))) & ".gif border=0 align=absmiddle><span style=""cursor:default"">" & RSObj("UserName") & "</span><span></td>")
			  Response.Write ("<td  class='splittd' align=""center"">")
			  IF RSOBJ("SuperTF")=1 then response.write "<font color=red>超级管理员</font>" else response.write "普通管理员"
			  Response.Write ("</td>")
			  Response.Write ("<td class='splittd' align=""center"">" & RSObj("LastLoginIP") & "</td>")
			  Response.Write ("<td class='splittd' align=""center"">" & RSObj("LastLoginTime") & "</td>")
			  Response.Write ("<td class='splittd' align=""center"">" & RSObj("LastLogoutTime") & "</td>")
			  Response.Write ("<td class='splittd' align=""center"">" & RSObj("LoginTimes") & "</td>")
			  Response.Write ("<td class='splittd' align=""center"">" & LockStr & "</td>")
			  Response.Write ("<td class='splittd' align=""center""><a href='javascript:EditAdmin(" & rsobj("AdminID") &")'>修改</a> | <a")
			  if rsobj("supertf")="1" then response.write " disabled" else response.write " href='javascript:Delete("&rsobj("AdminID")&")'"
			  Response.Write ">删除</a></td>"
			  Response.Write ("</tr>")
			  RSObj.MoveNext
			 If RSObj.EOF Then Exit Do
			Loop
			RSObj.Close:Conn.Close:Set RSObj = Nothing:Set GRS = Nothing
		  
		Response.Write "</table>"
		Response.Write "</div>"
		Response.Write "</body>"
		Response.Write "</html>"
		End Sub
		
		Sub AdminAdd()
		 IF KS.G("Method")="save" Then
		    Call AdminSave()
		   Else
		    Call AdminAddOrEdit()
		  End IF
		End SUB
		Sub AdminAddOrEdit()
		Dim SQL,RSC:Set RSC=Server.CreateObject("ADODB.RECORDSET")
		RSC.Open "Select ChannelID,ChannelName,BasicType,ItemName,ModelEname,ChannelStatus From KS_Channel where channelstatus=1 Order By ChannelID",Conn,1,1
		If Not RSC.Eof Then
		  SQL=RSC.GetRows(-1)
		End If
		RSC.Close:Set RSC=Nothing
		
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		Response.Write "<link href=""Include/admin_style.css"" rel=""stylesheet"">"
		Response.Write "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>"
		Response.Write "<script language=""JavaScript"" src=""../KS_Inc/jquery.js""></script>"
		Response.Write "<script language=""JavaScript"" src=""../ks_Inc/CheckPassWord.js""></script>"
		Response.Write "<style>"
		Response.Write ".rank { border:none; background:url(../images/rank.gif) no-repeat; width:125px; height:22px; vertical-align:middle; cursor:default; }"
		Response.Write ".r0 { background-position:0 0; }"
		Response.Write ".r1 { background-position:0 -22px; }"
		Response.Write ".r2 { background-position:0 -44px; }"
		Response.Write ".r3 { background-position:0 -66px; }"
		Response.Write ".r4 { background-position:0 -88px; }"
		Response.Write ".r5 { background-position:0 -110px; }"
		Response.Write ".r6 { background-position:0 -132px; }"
		Response.Write ".r7 { background-position:0 -154px; }"
		Response.Write "</style>"
		Response.Write "<title>管理员添加</title>"
		Response.Write "</head>"
		
		Dim AdminID, PrUserName, PassWord, Locked, RealName, Sex, TelPhone, Email, Descript, Action, GroupID, SuperTF
		
		Action = KS.G("Action")
		AdminID = KS.G("AdminID")
		GroupID = KS.G("GroupID")
		If Action = "" Then Action = "AddAdmin"
		If AdminID <> "" Then
		   Dim RSObj:Set RSObj = Conn.Execute("Select * From KS_Admin Where AdminID=" & AdminID)
		  If Not RSObj.EOF Then
			 UserName = Trim(RSObj("UserName"))
			 PrUserName=Trim(RSObj("PrUserName"))
			 Locked = Trim(CStr(RSObj("Locked")))
			 RealName = Trim(RSObj("RealName"))
			 Sex = Trim(RSObj("Sex"))
			 TelPhone = Trim(RSObj("TelPhone"))
			 Email = Trim(RSObj("Email"))
			 Descript = Trim(RSObj("Description"))
			 SuperTF = Trim(CStr(RSObj("SuperTF")))
			 PowerList=rsobj("powerlist")
	         ModelPower=rsobj("modelpower")
		  End If
		   RSObj.Close:Set RSObj = Nothing
		Else
		 ModelPower="sysset0,user0,lab0,model0,subsys0,other0,ask0,space0"
		 For i=0 to ubound(sql,2)
		  ModelPower=Modelpower &"," & sql(4,i)&"0"
		 Next
		End If
		
		
		
		Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
		
		 If AdminID = "" Then
			Response.Write ("<div class='topdashed sort'>添加管理员</div>")
		  Else
			Response.Write ("<div class='topdashed sort'>修改管理员</div>")
		  End If

		Response.Write "  <table width=""100%"" border=""0"" align=""center"" cellpadding=""3"" cellspacing=""1"">"
		Response.Write "  <form action=""?Method=save"" name=""AdminForm"" method=""post"">"
		Response.Write "   <input name=""Action"" type=""hidden"" id=""Action"" value=""" & Action & """>"
		Response.Write "   <input name=""AdminID"" type=""hidden"" value=""" & AdminID & """>"
		Response.write "  <tr class='sort'><td colspan='4' align='center'>====登录选项====</td></tr>"
		Response.Write "    <tr class='tdbg'>"
		Response.Write "      <td class='clefttitle' align=""right"">管理员名：</td>"
		Response.Write "      <td height=""25"" colspan='3'>"
					
					If Action = "Edit" Then
						 Response.Write ("<input name=""UserName"" Readonly value=""" & UserName & """ type=""text"" id=""UserName"" size=""30"" class=""textbox"">")
					Else
						 Response.Write ("<input name=""UserName""  type=""text"" id=""UserName"" size=""30"" class=""textbox"">")
					End If
					 
		Response.Write "              用于登录后台的名称</td>"
		Response.Write "    </tr>"
		Response.Write "    <tr class='tdbg'>"
		Response.Write "            <td height=""25"" class='clefttitle' align=""right"">前台用户名：</td>"
		Response.Write "            <td colspan='3'>"
					If Action = "Edit" Then
						 Response.Write ("<input name=""PrUserName"" readonly value=""" & PrUserName & """ type=""text"" id=""PrUserName"" size=""30"" class=""textbox"">")
					Else
						 Response.Write ("<input name=""PrUserName""  type=""text"" id=""PrUserName"" size=""30"" class=""textbox"">")
					End if
					 
		Response.Write "              前台会员中心注册的用户名(不可更改),<a href='KS.User.asp?Action=Add'><font color=red>点此添加</font></a></td>"
		Response.Write "          </tr>"		
				  
				  If Action <> "Edit" Then
		Response.Write "          <tr class='tdbg'>"
		Response.Write "            <td height=""25"" class=""clefttitle"" align=""right"">初始密码：</td>"
		Response.Write "            <td height=""25"" colspan='3'>"
		Response.Write "             <table border='0'><tr><td><input name=""PassWord"" type=""password"" size=""20"" class=""textbox"" onkeyup=""javascript:setPasswordLevel(this, document.getElementById('passwordLevel'));"" onmouseout=""javascript:setPasswordLevel(this, document.getElementById('passwordLevel'));"" onblur=""javascript:setPasswordLevel(this, document.getElementById('passwordLevel'));""> 密码强度：</td><td> "
		Response.Write "         <input name=""Input"" disabled=""disabled"" class=""rank r0"" id=""passwordLevel"" /></td>"

		Response.Write "          </tr></table></td></tr>"
				 
				 End If
		Response.Write "          <tr class='tdbg'>"
		Response.Write "            <td height=""25"" align=""right"" class='clefttitle'>是否锁定：</td>"
		Response.Write "            <td height=""25"" colspan='3'>"
					
					If SuperTF = "1" Then
					   Response.Write ("<input type=""hidden"" value=""0"" name=""locked""> (否)正常")
					  ElseIf Locked = "1" Then
					 Response.Write ("<input type=""radio"" name=""Locked"" value=""0""> 正常 ")
					 Response.Write ("<input type=""radio"" name=""Locked"" value=""1"" checked> 锁定 ")
					 Else
					  Response.Write ("<input type=""radio"" name=""Locked"" value=""0"" checked> 正常 ")
					  Response.Write ("<input type=""radio"" name=""Locked"" value=""1""> 锁定 ")
					 End If
					  
		Response.Write "              　　<font color=""#FF0000"">锁定的用户不能登录后台管理</font></td>"
		Response.Write "          </tr>"
		
		Response.Write "          <tr class='tdbg'>"
		Response.Write "            <td height=""25"" align=""right"" class='clefttitle'>真实姓名：</td>"
		Response.Write "            <td><input name=""RealName"" type=""text"" class=""textbox"" value=""" & RealName & """ id=""RealName"" size=""30""></td>"
		Response.Write "            <td align=""right"" class='clefttitle'>性 别：</td>"
		Response.Write "            <td>"
					 
					 If Trim(Sex) = "女" Then
						  Response.Write ("<input type=""radio"" name=""Sex"" value=""男""> 男 ")
						  Response.Write ("<input type=""radio"" name=""Sex"" value=""女"" checked>  女 ")
					  Else
						  Response.Write ("<input type=""radio"" name=""Sex"" value=""男"" checked> 男 ")
						  Response.Write ("<input type=""radio"" name=""Sex"" value=""女"">  女 ")
					  End If
				   
		Response.Write "             </td>"
		Response.Write "          </tr>"
		Response.Write "          <tr class='tdbg'>"
		Response.Write "            <td height=""25"" align=""right"" class='clefttitle'>联系电话：</td>"
		Response.Write "            <td><input name=""TelPhone"" type=""text"" class=""textbox"" value=""" & TelPhone & """ id=""TelPhone"" size=""30""></td>"
		Response.Write "            <td align=""right"" class='clefttitle'>电子信箱：</td>"
		Response.Write "            <td><input name=""Email"" type=""text"" class=""textbox"" id=""Email"" value=""" & Email & """ size=""30""></td>"
		Response.Write "          </tr>"
		Response.Write "          <tr class='tdbg'>"
		Response.Write "            <td height=""25"" align=""right"" class='clefttitle'>简要说明：</td>"
		Response.Write "            <td height=""25"" colspan='3'>"
		Response.Write "              <textarea class='textbox' name=""Description"" rows=""6"" id=""Description"" style=""width:80%;height:60px;border-style: solid; border-width: 1"">" & Descript & "</textarea></td>"
		Response.Write "          </tr>"
		if SuperTF=1 then
		Response.write "          <tr class='sort'><td colspan='4' align='center'>====此管理是超级管理员，拥有最高权限====</td></tr>"
		Response.Write "          <tr class='tdbg' style='display:none'><td colspan='4'>"
		else
		Response.write "          <tr class='sort'><td colspan='4' align='center'>====此管理的详细权限设置====</td></tr>"
		Response.Write "          <tr class='tdbg'><td colspan='4'>"
		end if
		
		 dim i
	 
	 %>
	 <table width="99%" border="0" align="center" cellspacing="0" cellpadding="0">  
	 <tr>
	 <td height="25" class='clefttitle'><strong> 一、此管理员在【<font color="#FF0000">内容管理</font>】的权限</strong></td>
	 </tr>
	 </table>
    <table width="96%" border="0" align="center" cellspacing="0" cellpadding="0">   
	 <tr>       
	 <td> 
		  
          <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
		    <tr> 
              <td height="22" align="center"> <strong><font color="#993300">模型公共权限</font></strong></td>
			  <td>
			  	
				<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td width="25%"> <input name="PowerList" type="checkbox" id="PowerList" value="M010001"<%if InStr(1, PowerList,"M010001" ,1)<>0 then Response.Write(" checked") %>>
                      栏目管理</td>
                    <td width="25%"><input name="PowerList" type="checkbox" id="PowerList" value="M010002"<%if InStr(1, PowerList,"M010002" ,1)<>0 then Response.Write( " checked") %>>
                      评论管理</td>
                    <td width="25%"><input name="PowerList" type="checkbox" id="PowerList" value="M010003"<%if InStr(1, PowerList,"M010003" ,1)<>0 then Response.Write( " checked") %>>
                      专题管理</td>
                    <td width="25%"><input name="PowerList" type="checkbox" id="PowerList" value="M010004"<%if InStr(1, PowerList,"M010004" ,1)<>0 then Response.Write( " checked") %>>
                      关键字tags管理</td>
				  </tr>
				  <tr>
                    <td width="25%"><input name="PowerList" type="checkbox" id="PowerList" value="M010005"<%if InStr(1, PowerList,"M010005" ,1)<>0 then Response.Write(" checked") %>>
批量设置</td>
                    <td width="25%"><input name="PowerList" type="checkbox" id="PowerList" value="M010006"<%if InStr(1, PowerList,"M010006" ,1)<>0 then Response.Write(" checked") %>>
                      回收站管理</td>
                    <td width="25%"> <input name="PowerList" type="checkbox" id="PowerList" value="M010007"<%if InStr(1, PowerList,"M010007" ,1)<>0 then Response.Write(" checked") %>>
                      一键管理工具</td>
                    <td width="25%"> <input name="PowerList" type="checkbox" id="PowerList" value="M010008"<%if InStr(1, PowerList,"M010008" ,1)<>0 then Response.Write(" checked") %>>
                      采集管理</td>
                 
                  </tr>
				  </table>
				
				
				
			  </td>
			</tr>
			<tr><td colspan=2><hr size=1></td></tr>
		    <tr> 
              <td height="22" align="center"> <strong><font color="#993300">分模型权限设置</font></strong></td>
			  <td>
			      <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                  
			  	 <%
				  dim m:m=1
				 for i=0 to ubound(sql,2)%>
				   <tr> 
				   <td width="20%" height="35"> <strong><%=sql(1,i)%></strong></td>
				   </tr>
				   <tr>
				    <td>
					
					<%IF instr(ModelPower,sql(4,i) & "0")>0 Then
			      Response.Write("<input name=""ModelPower" & sql(0,i) & """ type=""radio"" onclick=""M_" & sql(1,i) & ".style.display='none';"" value=""" & sql(4,I) & "0"" checked>")
			   Else
			      Response.Write("<input name=""ModelPower" & sql(0,i) & """ type=""radio"" onclick=""M_" & sql(1,i) & ".style.display='none';"" value=""" & sql(4,I) & "0"">")
			   End IF
			   %>
                在<%=SQL(1,I)%>无任何管理权限(屏蔽)
					<br/>
					<%IF instr(ModelPower,sql(4,i) & "1")>0 Then
			      Response.Write("<input type=""radio"" onclick=""M_" & sql(1,i) & ".style.display='none';"" name=""ModelPower" & sql(0,i) & """ value=""" & sql(4,I) & "1"" Checked>")
				Else
			      Response.Write("<input type=""radio"" onclick=""M_" & sql(1,i) & ".style.display='none';"" name=""ModelPower" & sql(0,i) & """ value=""" & sql(4,I) & "1"">")
				End IF
				%>
                模型管理员：拥有该模型的所有管理权限(相当于对<%=sql(1,i)%>没有任何限制)
				 <br>
				 <%IF instr(ModelPower,sql(4,i) & "2")>0 Then
			     Response.Write("<input type=""radio"" onclick=""M_" & sql(1,i) & ".style.display='';"" name=""ModelPower" & sql(0,i) & """ value=""" & sql(4,I) & "2"" Checked>")
			   Else
			     Response.Write("<input type=""radio"" onclick=""M_" & sql(1,i) & ".style.display='';"" name=""ModelPower" & sql(0,i) & """ value=""" & sql(4,I) & "2"">")
			   End IF
			   %>
                栏目管理员：只拥有部分栏目(频道)管理权限
					
					</td>
				   </tr>
				   <tr ID="M_<%=sql(1,i)%>" <%IF instr(ModelPower,sql(4,i) & "2")=0 Then Response.Write("style=""display:none""") End IF%>>	 
			       <td height="22">
				    
					
					
					<%
	  Select Case SQL(2,I)
	   Case 1,2,3,4,7,8
	   %>  
              <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td height="25" colspan="7"><strong><font color="#993300">权限设置</font></strong></td>
                  </tr>
                    <%
					Call BasePurview(PowerList,SQL,I)
					%>
                  <tr> 
                    <td height="25" colspan="7"><font color="#993300"><strong>详细指定栏目（频道）权限</strong></font></td>
                  </tr>
                  <tr> 
                    <td colspan="7"> 
					   <%
                       Call ClassList(SQL(0,I))
					   %>
					</td>
                  </tr>
                </table>
	<%case 5%>
<table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">

            <tr>
              <td height="25" colspan="7"><strong><font color="#993300">权限设置</font></strong></td>
            </tr>
            <%Call BasePurview(PowerList,SQL,I)%>
            <tr>
              <td height="25" colspan="7"><strong><font color="#993300">常规管理</font></strong></td>
            </tr>
            <tr>
			  <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10012"<%if InStr(1, PowerList,"M"&SQL(0,I) & "10012" ,1)<>0 then Response.Write(" checked") %>>
订单处理</td>
              <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10014"<%if InStr(1, PowerList,"M"&SQL(0,I) & "10014" ,1)<>0 then Response.Write(" checked") %>>
                资金明细</td>
              <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10015"<%if InStr(1, PowerList,"M"&SQL(0,I) & "10015" ,1)<>0 then Response.Write(" checked") %>>
发退货查询</td>
              <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10016"<%if InStr(1, PowerList,"M"&SQL(0,I) & "10016" ,1)<>0 then Response.Write(" checked") %>>
开发票查询</td>
              <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10017"<%if InStr(1, PowerList,"M"&SQL(0,I) & "10017" ,1)<>0 then Response.Write(" checked") %>>
销售统计</td>
              <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10018"<%if InStr(1, PowerList,"M"&SQL(0,I) & "10018" ,1)<>0 then Response.Write(" checked") %>>
品牌管理</td>
            </tr>
			
            <tr>
              <td nowrap="nowrap" title="浏览、编辑、删除厂商等操作的权限"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20003"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20003" ,1)<>0 then Response.Write( " checked") %> />
                厂商管理</td>
				<td nowrap="nowrap" title="浏览、编辑、删除送货方式等操作的权限"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20004"<%if InStr(1, PowerList,"M" & sql(0,i) & "20004" ,1)<>0 then Response.Write( " checked") %> />
                送货方式管理</td>
				<td nowrap="nowrap" title="浏览、编辑、删除付款方式等操作的权限"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20005"<%if InStr(1, PowerList,"M" & sql(0,i) & "20005" ,1)<>0 then Response.Write( " checked") %> />
                付款方式管理</td>               
				<td nowrap><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20001"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20001" ,1)<>0 then Response.Write( " checked") %>>
                      看货地点管理</td>
                    
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20007"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20007" ,1)<>0 then Response.Write( " checked") %>> 优惠券管理</td>
                    <td nowrap="nowrap"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20008"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20008" ,1)<>0 then Response.Write( " checked") %>> 限时/限量管理</td>
            </tr>
			<tr>
               <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20009"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20009" ,1)<>0 then Response.Write( " checked") %>> 捆绑销售管理</td>
               <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20010"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20010" ,1)<>0 then Response.Write( " checked") %>> 换购商品管理</td>
               <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20011"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20011" ,1)<>0 then Response.Write( " checked") %>> 超值礼包管理</td>
              <td nowrap="nowrap"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>30001"<%if InStr(1, PowerList,"M" & SQL(0,I) & "30001" ,1)<>0 then Response.Write( " checked") %> />
                团购管理</td>
				<td nowrap="nowrap" title="浏览、编辑、删除送货方式等操作的权限"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>30002"<%if InStr(1, PowerList,"M" & sql(0,i) & "30002" ,1)<>0 then Response.Write( " checked") %> />
                兴趣活动管理</td>
			</tr>
			
			
            <tr>
              <td height="25" colspan="7"><font color="#993300"><strong>详细指定栏目（频道）权限</strong></font></td>
            </tr>
            <tr>
              <td colspan="7">                      
			        <%
                       Call ClassList(SQL(0,I))
					   %>
               </td>
            </tr>
          </table>
  <%case 6%>
  <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="25" colspan="7"><strong><font color="#993300">权限设置</font></strong></td>
            </tr>
            <tr>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="KSM1001"<%if InStr(1, PowerList,"KSM1001" ,1)<>0 then Response.Write(" checked") %> />
                歌曲管理</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="KSM1002"<%if InStr(1, PowerList,"KSM1002" ,1)<>0 then Response.Write(" checked") %> />
                专辑管理</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="KSM1003"<%if InStr(1, PowerList,"KSM1003" ,1)<>0 then Response.Write(" checked") %> />
                类别管理</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="KSM1004"<%if InStr(1, PowerList,"KSM1004" ,1)<>0 then Response.Write( " checked") %> />
                歌手管理</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="KSM1005"<%if InStr(1, PowerList,"KSM1005" ,1)<>0 then Response.Write( " checked") %> />
                评论管理</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="KSM1006"<%if InStr(1, PowerList,"KSM1006" ,1)<>0 then Response.Write( " checked") %> />
               服务器管理</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="KSM1007"<%if InStr(1, PowerList,"KSM1007" ,1)<>0 then Response.Write(" checked") %>>
               审核用户歌词</td>
            </tr>
			<tr>
			 <td colspan=6><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20005"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20005" ,1)<>0 then Response.Write( " checked") %>>
                      生成HTML操作</td>
			</tr>
            
          </table>

		   <%case 9%>
           <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="25" colspan="7"><strong><font color="#993300">栏目与信息</font></strong></td>
            </tr>
            <tr>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M910001"<%if InStr(1, PowerList,"M910001" ,1)<>0 then Response.Write(" checked") %> />
                栏目管理</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M910002"<%if InStr(1, PowerList,"M910002" ,1)<>0 then Response.Write( " checked") %> />
添加试卷</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M910003"<%if InStr(1, PowerList,"M910003" ,1)<>0 then Response.Write( " checked") %> />
编辑试卷</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M910004"<%if InStr(1, PowerList,"M910004" ,1)<>0 then Response.Write( " checked") %> />
删除试卷</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M910005"<%if InStr(1, PowerList,"M910005" ,1)<>0 then Response.Write(" checked") %>>
移动试卷</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M910007"<%if InStr(1, PowerList,"M910007" ,1)<>0 then Response.Write(" checked") %>>
发布管理</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M910009"<%if InStr(1, PowerList,"M910009" ,1)<>0 then Response.Write(" checked") %>>
上传文件</td>
            </tr>
			</table>
		   <%case 10%>
           <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td height="25" colspan="7"><strong><font color="#993300">栏目与信息</font></strong></td>
            </tr>
            <tr>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M1010001"<%if InStr(1, PowerList,"M1010001" ,1)<>0 then Response.Write(" checked") %> />
                招聘系统设置</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M1010002"<%if InStr(1, PowerList,"M1010002" ,1)<>0 then Response.Write( " checked") %> />
招聘单位管理</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M1010003"<%if InStr(1, PowerList,"M1010003" ,1)<>0 then Response.Write( " checked") %> />
个人简历管理</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M1010004"<%if InStr(1, PowerList,"M1010004" ,1)<>0 then Response.Write( " checked") %> />
招聘职位管理</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M1010005"<%if InStr(1, PowerList,"M1010005" ,1)<>0 then Response.Write(" checked") %>>
行业职位设置</td>
              <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M1010007"<%if InStr(1, PowerList,"M1010007" ,1)<>0 then Response.Write(" checked") %>>
简历模板管理</td>
            </tr>
			</table>
			<%
	End Select
	%>  
				   </td>
				   </tr>
				   
                  <%
				  Next%>
				  
				   <tr> 
				   <td width="20%" height="35"> <strong>问答系统权限</strong></td>
				   </tr>
				   <tr>
				    <td>
			<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
			<tr> 
              <td height="25" colspan="5"> 
                <%
					IF instr(ModelPower,"ask0")>0 Then
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('Ask','none')"" name=""ask"" value=""ask0"" checked>")
                    ELSE
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('Ask','none')"" name=""ask"" value=""ask0"">")
					END IF
					%>
                在问答系统无任何管理权限(屏蔽)
				  <br/>
                <%
					IF instr(ModelPower,"ask1")>0 Then
					  Response.Write("<input type=""radio"" name=""ask"" onclick=""SetPowerListValue('Ask','')"" value=""ask1"" checked>")
                     ELSE
					  Response.Write("<input type=""radio"" name=""ask"" onclick=""SetPowerListValue('Ask','')"" value=""ask1"">")
					 END IF%>
                拥有指定的部分管理权限 
			 </td>
            </tr>
            <tr ID="Ask" <% IF instr(ModelPower,"ask1")="0" then Response.Write("style=""display:none""") End IF%>> 
							<td><input name="PowerList" type="checkbox" id="PowerList" value="WDXT10000"<%if InStr(1, PowerList,"WDXT10000" ,1)<>0 then Response.Write( " checked") %>> 
							问答参数设置
		</td>
							<td><input name="PowerList" type="checkbox" id="PowerList" value="WDXT10001"<%if InStr(1, PowerList,"WDXT10001" ,1)<>0 then Response.Write( " checked") %>> 
							编辑删除问题</td>
							<td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="WDXT10002"<%if InStr(1, PowerList,"WDXT10002" ,1)<>0 then Response.Write( " checked") %>> 
							问答分类管理</td>
							<td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="WDXT10003"<%if InStr(1, PowerList,"WDXT10003" ,1)<>0 then Response.Write( " checked") %>>
		等级头衔管理</td>
		

					</tr>
					
				   <tr> 
				   <td width="20%" height="35"> <strong>空间门户权限</strong></td>
				   </tr>
			<tr> 
              <td height="25" colspan="5"> 
                <%
					IF instr(ModelPower,"space0")>0 Then
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('Space','none')"" name=""space"" value=""space0"" checked>")
                    ELSE
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('Space','none')"" name=""space"" value=""space0"">")
					END IF
					%>
                在空间门户无任何管理权限(屏蔽)
				  <br/>
                <%
					IF instr(ModelPower,"space1")>0 Then
					  Response.Write("<input type=""radio"" name=""space"" onclick=""SetPowerListValue('Space','')"" value=""space1"" checked>")
                     ELSE
					  Response.Write("<input type=""radio"" name=""space"" onclick=""SetPowerListValue('Space','')"" value=""space1"">")
					 END IF%>
                拥有指定的部分管理权限 
			 </td>
            </tr>
            <tbody ID="Space" <% IF instr(ModelPower,"space1")="0" then Response.Write("style=""display:none""") End IF%>> 
                   <tr>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10000"<%if InStr(1, PowerList,"KSMS10000" ,1)<>0 then Response.Write( " checked") %>>
空间参数设置</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10001"<%if InStr(1, PowerList,"KSMS10001" ,1)<>0 then Response.Write( " checked") %>>
个人空间管理</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10002"<%if InStr(1, PowerList,"KSMS10002" ,1)<>0 then Response.Write( " checked") %>>
空间日志管理</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10003"<%if InStr(1, PowerList,"KSMS10003" ,1)<>0 then Response.Write( " checked") %>>
用户相册管理 </td>
                  </tr>
				  <tr>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10004"<%if InStr(1, PowerList,"KSMS10004" ,1)<>0 then Response.Write( " checked") %>>
用户圈子管理</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10005"<%if InStr(1, PowerList,"KSMS10005" ,1)<>0 then Response.Write( " checked") %>>
用户留言管理</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20013"<%if InStr(1, PowerList,"KSMS20013" ,1)<>0 then Response.Write( " checked") %>>
行业广告管理</td>
				   <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10007"<%if InStr(1, PowerList,"KSMS10007" ,1)<>0 then Response.Write( " checked") %>>
用户歌曲管理</td>
                  </tr>
				  <tr>
				   <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10008"<%if InStr(1, PowerList,"KSMS10008" ,1)<>0 then Response.Write( " checked") %>>
企业信息管理</td>
				   <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10009"<%if InStr(1, PowerList,"KSMS10009" ,1)<>0 then Response.Write( " checked") %>>
企业新闻管理</td>
				   <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10010"<%if InStr(1, PowerList,"KSMS10010" ,1)<>0 then Response.Write( " checked") %>>
企业产品管理</td>
				   <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20011"<%if InStr(1, PowerList,"KSMS20011" ,1)<>0 then Response.Write( " checked") %>>
荣誉证书管理</td>
                  </tr>
				  <tr>
				   <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20012"<%if InStr(1, PowerList,"KSMS20012" ,1)<>0 then Response.Write( " checked") %>>
行业分类管理</td>
				   <td></td>

				  </tr>					
				 </tbody>	
					
					
				  </table>
					
		            </td>
				  </tr>
				  
				 
				 </table>
				  
			  </td>
			</tr>
			
		
		 </table>
			  
            
			 
			 
	       </TD>
            </TR>
          </table>
         </td>
		  </tr>
    </table>
	
	  <br/>
	 <table width="99%" border="0" align="center" cellspacing="0" cellpadding="0">  
	 <tr>
	 <td height="25" class='clefttitle'><strong> 二、此管理员在【<font color="#FF0000">系统设置</font>】的权限</strong></td>
	 </tr>
	 </table>

	  <table width="96%" align="center" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>
          <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr> 
              <td height="25" colspan="2"> 
                <%
					IF instr(ModelPower,"sysset0")>0 Then
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('System','none')"" name=""sysset"" value=""sysset0"" checked>")
                    ELSE
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('System','none')"" name=""sysset"" value=""sysset0"">")
					END IF
					%>
                在系统管理中心无任何管理权限(屏蔽)
				  <br/>
                <%
					IF instr(ModelPower,"sysset1")>0 Then
					  Response.Write("<input type=""radio"" name=""sysset"" onclick=""SetPowerListValue('System','')"" value=""sysset1"" checked>")
                     ELSE
					  Response.Write("<input type=""radio"" name=""sysset"" onclick=""SetPowerListValue('System','')"" value=""sysset1"">")
					 END IF%>
                拥有指定的部分管理权限 
			 </td>
            </tr>
            <tr ID="System" <% IF instr(ModelPower,"sysset1")="0" then Response.Write("style=""display:none""") End IF%>> 
              <td height="25" colspan="2">
			  <table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr> 
                    <td width="13%" height="25" align="center" nowrap><strong>系统设置</strong></td>
                    <td width="13%"> <input name="PowerList" type="checkbox" id="PowerList" value="KMST10001"<%if InStr(1, PowerList,"KMST10001" ,1)<>0 then Response.Write( " checked") %>>
                      基本信息设置</td>
                    <td width="11%"><input name="PowerList" type="checkbox" id="PowerList" value="KMST10002"<%if InStr(1, PowerList,"KMST10002" ,1)<>0 then Response.Write( " checked") %>>
API整合设置</td>
                    <td width="15%"><input name="PowerList" type="checkbox" id="PowerList" value="KMST20000"<%if InStr(1, PowerList,"KMST20000" ,1)<>0 then Response.Write( " checked") %>>
更新缓存</td>
                    <td width="15%" nowrap><input name="PowerList" type="checkbox" id="PowerList" value="KMST20001"<%if InStr(1, PowerList,"KMST20001" ,1)<>0 then Response.Write( " checked") %>>
下载参数设置</td>
                    <td width="15%" nowrap><input name="PowerList" type="checkbox" id="PowerList" value="KMST20002"<%if InStr(1, PowerList,"KMST20002" ,1)<>0 then Response.Write( " checked") %>>
下载服务器设置</td>
                  </tr>
                  <tr> 
                    <td height="25" align="center" nowrap></td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST20003"<%if InStr(1, PowerList,"KMST20003" ,1)<>0 then Response.Write( " checked") %>>
影视参数设置</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST20004"<%if InStr(1, PowerList,"KMST20004" ,1)<>0 then Response.Write( " checked") %>>
影视服务器</td>
                    <td nowrap><input name="PowerList" type="checkbox" id="PowerList" value="KMST20005"<%if InStr(1, PowerList,"KMST20005" ,1)<>0 then Response.Write( " checked") %>>
供求交易类别</td>
                  </tr>
                  
				  <tr> 
                    <td width="13%" height="25" align="center" nowrap><strong>辅助管理</strong></td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10004"<%if InStr(1, PowerList,"KMST10004" ,1)<>0 then Response.Write( " checked") %>>
                      内容关键字</td>                    
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10006"<%if InStr(1, PowerList,"KMST10006" ,1)<>0 then Response.Write( " checked") %>>
                      后台日志管理</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10007"<%if InStr(1, PowerList,"KMST10007" ,1)<>0 then Response.Write( " checked") %>>
                      备份/压缩数据库</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10008"<%if InStr(1, PowerList,"KMST10008" ,1)<>0 then Response.Write( " checked") %>>
                      数据库字段替换</td>
					 <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10019"<%if InStr(1, PowerList,"KMST10019" ,1)<>0 then Response.Write( " checked") %>>
                      搜索关键词维护</td>
                  </tr>
				  <tr>
				    <td height="25" align="center" nowrap>&nbsp;</td>
				    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10013"<%if InStr(1, PowerList,"KMST10013" ,1)<>0 then Response.Write(" checked") %> />
专题管理</td>
				    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10015"<%if InStr(1, PowerList,"KMST10015" ,1)<>0 then Response.Write(" checked") %> />
来源管理</td>
				    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10016"<%if InStr(1, PowerList,"KMST10016" ,1)<>0 then Response.Write(" checked") %> />
作者管理</td>
				    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10017"<%if InStr(1, PowerList,"KMST10017" ,1)<>0 then Response.Write(" checked") %> />
省市管理</td>
				    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10014"<%if InStr(1, PowerList,"KMST10014" ,1)<>0 then Response.Write(" checked") %> />
图片投票记录</td>			    </tr>
				  <tr>
				    <td height="25" align="center" nowrap></td>
				    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10009"<%if InStr(1, PowerList,"KMST10009" ,1)<>0 then Response.Write( " checked") %>>
在线执行SQL语句 </td>
				    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10010"<%if InStr(1, PowerList,"KMST10010" ,1)<>0 then Response.Write( " checked") %>>
空间占用量 </td>

				    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10011"<%if InStr(1, PowerList,"KMST10011" ,1)<>0 then Response.Write( " checked") %>>
服务器参数探测</td>
			        <td nowrap><input name="PowerList" type="checkbox" id="PowerList" value="KMST10012"<%if InStr(1, PowerList,"KMST10012" ,1)<>0 then Response.Write( " checked") %>>
在线检测木马</td>
			        <td nowrap><input name="PowerList" type="checkbox" id="PowerList" value="KMST10018"<%if InStr(1, PowerList,"KMST10018" ,1)<>0 then Response.Write( " checked") %>>
上传文件管理</td>
		          </tr>
				  <tr>
				    <td height="25" align="center" nowrap></td>
				    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMST10020"<%if InStr(1, PowerList,"KMST10020" ,1)<>0 then Response.Write( " checked") %>>
定时任务管理 </td>
				    <td> </td>

				    <td></td>
			        <td nowrap></td>
			        <td nowrap></td>
		          </tr>
                </table>
			  </td>
            </tr>
			
          </table>
	  </td></tr>
	  </table>
	  
	  
	  <br/>
	 <table width="99%" border="0" align="center" cellspacing="0" cellpadding="0">  
	 <tr>
	 <td height="25" class='clefttitle'><strong> 三、此管理员在【<font color="#FF0000">用户管理</font>】的权限</strong></td>
	 </tr>
	 </table>

	  <table width="96%" align="center" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>
          <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr> 
              <td height="25" colspan="2"> 
                <%
					IF instr(ModelPower,"user0")>0 Then
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('UserAdmin','none')"" name=""userPower"" value=""user0"" checked>")
                    ELSE
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('UserAdmin','none')"" name=""UserPower"" value=""user0"">")
					END IF
					%>
                在用户管理中心无任何管理权限(屏蔽)
				<br/>
                <%
					IF instr(ModelPower,"user1")>0 Then
					  Response.Write("<input type=""radio"" name=""UserPower"" onclick=""SetPowerListValue('UserAdmin','')"" value=""user1"" checked>")
                     ELSE
					  Response.Write("<input type=""radio"" name=""UserPower"" onclick=""SetPowerListValue('UserAdmin','')"" value=""user1"">")
					 END IF%>
                拥有指定的部分管理权限 </td>
            </tr>
            <tr ID="UserAdmin" <% IF instr(ModelPower,"user1")=0 then Response.Write("style=""display:none""") End IF%>> 
              <td height="25"> <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td><input disabled="disabled" name="PowerList" type="checkbox" id="PowerList" value="KMUA10001"<%if InStr(1, PowerList,"KMUA10001" ,1)<>0 then Response.Write( " checked") %>>
                      管理员管理</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10002"<%if InStr(1, PowerList,"KMUA10002" ,1)<>0 then Response.Write( " checked") %>>
                      注册用户管理</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10003"<%if InStr(1, PowerList,"KMUA10003" ,1)<>0 then Response.Write( " checked") %>>
                     用户短信管理 </td>
                    <td> <input name="PowerList" type="checkbox" id="PowerList" value="KMUA10004"<%if InStr(1, PowerList,"KMUA10004" ,1)<>0 then Response.Write( " checked") %>>
                      用户组管理</td>
				    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10008"<%if InStr(1, PowerList,"KMUA10008" ,1)<>0 then Response.Write( " checked") %>>
                    充值卡管理</td>                 
				    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10011"<%if InStr(1, PowerList,"KMUA10011" ,1)<>0 then Response.Write( " checked") %>>
                    查看工作进度</td>					 </tr>
                  <tr>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10005"<%if InStr(1, PowerList,"KMUA10005" ,1)<>0 then Response.Write( " checked") %>>
                    会员点券明细</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10006"<%if InStr(1, PowerList,"KMUA10006" ,1)<>0 then Response.Write( " checked") %>>
                    会员有效期明细</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10007"<%if InStr(1, PowerList,"KMUA10007" ,1)<>0 then Response.Write( " checked") %>>
                    会员资金明细</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10008"<%if InStr(1, PowerList,"KMUA10008" ,1)<>0 then Response.Write( " checked") %>>
                    会员资金明细</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10009"<%if InStr(1, PowerList,"KMUA10009" ,1)<>0 then Response.Write( " checked") %>>
                    发送邮件管理</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10010"<%if InStr(1, PowerList,"KMUA10010" ,1)<>0 then Response.Write( " checked") %>>
                    修改自己的密码</td>
                  </tr>
                  <tr>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10012"<%if InStr(1, PowerList,"KMUA10012" ,1)<>0 then Response.Write( " checked") %>>
                    会员字段管理</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10013"<%if InStr(1, PowerList,"KMUA10013" ,1)<>0 then Response.Write( " checked") %>>
                    会员表单管理</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMUA10014"<%if InStr(1, PowerList,"KMUA10014" ,1)<>0 then Response.Write( " checked") %>>
                    会员动态管理</td>
                    <td>&nbsp;</td>
                    <td height="25">&nbsp;</td>
                    <td height="25">&nbsp;</td>
                  </tr>
			
		
                </table></td>
            </tr>
          </table>
	  </td></tr>
	  </table>
	  
	  
	  <br/>
	 <table width="99%" border="0" align="center" cellspacing="0" cellpadding="0">  
	 <tr>
	 <td height="25" class='clefttitle'><strong> 四、此管理员在【<font color="#FF0000">标签模板管理</font>】的权限</strong></td>
	 </tr>
	 </table>

	  <table width="96%" align="center" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>
          <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr> 
              <td height="25" colspan="2"> 
                <%
					IF instr(ModelPower,"lab0")>0 Then
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('KMTemplatePower','none')"" name=""labPower"" value=""lab0"" checked>")
                    ELSE
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('KMTemplatePower','none')"" name=""labPower"" value=""lab0"">")
					END IF
					%>
                在模板标签管理管理权限(屏蔽)</td>
            </tr>
            <tr> 
              <td height="25" colspan="2"> 
                <%
					IF instr(ModelPower,"lab1")>0 Then
					  Response.Write("<input type=""radio"" name=""labPower"" onclick=""SetPowerListValue('KMTemplatePower','')"" value=""lab1"" checked>")
                     ELSE
					  Response.Write("<input type=""radio"" name=""labPower"" onclick=""SetPowerListValue('KMTemplatePower','')"" value=""lab1"">")
					 END IF%>
                拥有指定的部分管理权限 </td>
            </tr>
            <tr ID="KMTemplatePower" <% IF instr(ModelPower,"lab1")=0 then Response.Write("style=""display:none""") End IF%>> 
              <td height="25"> <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td><strong>发布管理</strong></td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMTL20000"<%if InStr(1, PowerList,"KMTL20000" ,1)<>0 then Response.Write( " checked") %>>
发布站点首页</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMTL20001"<%if InStr(1, PowerList,"KMTL20001" ,1)<>0 then Response.Write( " checked") %>>
专题发布管理</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMTL20002"<%if InStr(1, PowerList,"KMTL20002" ,1)<>0 then Response.Write( " checked") %>>
系统JS发布管理 </td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMTL20003"<%if InStr(1, PowerList,"KMTL20003" ,1)<>0 then Response.Write( " checked") %>>
发布通用自定义页面</td>
                    <td></td>
                  </tr>

                  <tr>
                    <td width="10%" title="此权限包括系统函数标签目录管理，标签管理，搜索标签等"><strong>模板标签</strong></td> 
                    <td width="16%" title="此权限包括系统函数标签目录管理，标签管理，搜索标签等"> <input name="PowerList" type="checkbox" id="PowerList" value="KMTL10001"<%if InStr(1, PowerList,"KMTL10001" ,1)<>0 then Response.Write( " checked") %>>
                      系统函数标签 </td>
                    <td width="18%" height="25" title="此权限包括自由标签目录管理，标签管理，搜索标签等"> <input name="PowerList" type="checkbox" id="PowerList" value="KMTL10002"<%if InStr(1, PowerList,"KMTL10002" ,1)<>0 then Response.Write( " checked") %>>
                      自定义SQL标签</td>
                    <td width="18%" height="25" title="此权限包括系统JS目录管理，标签管理，搜索标签等"> <input name="PowerList" type="checkbox" id="PowerList" value="KMTL10003"<%if InStr(1, PowerList,"KMTL10003" ,1)<>0 then Response.Write( " checked") %>>
                    自定义静态标签 </td>
                    <td width="14%"  title="此权限包括自由JS目录管理，标签管理，搜索标签等"><input name="PowerList" type="checkbox" id="PowerList" value="KMTL10004"<%if InStr(1, PowerList,"KMTL10004" ,1)<>0 then Response.Write( " checked") %>>
                    系统JS管理</td>
                    <td width="24%" title="此权限包括模板导入，模板编辑等"><input name="PowerList" type="checkbox" id="PowerList" value="KMTL10005"<%if InStr(1, PowerList,"KMTL10005" ,1)<>0 then Response.Write( " checked") %>>
自由JS管理</td>
                  </tr>
                  <tr>
                    <td>&nbsp;</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMTL10006"<%if InStr(1, PowerList,"KMTL10006" ,1)<>0 then Response.Write( " checked") %>>自定义页面管理</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMTL10007"<%if InStr(1, PowerList,"KMTL10007" ,1)<>0 then Response.Write( " checked") %>>
模板管理</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMSL10008"<%if InStr(1, PowerList,"KMSL10008" ,1)<>0 then Response.Write( " checked") %> />
生成顶部菜单</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KMSL10009"<%if InStr(1, PowerList,"KMSL10009" ,1)<>0 then Response.Write( " checked") %> />
生成树型菜单</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KMSL10010"<%if InStr(1, PowerList,"KMSL10010" ,1)<>0 then Response.Write( " checked") %> />
通用循环标签</td>
                  </tr>
                </table></td>
            </tr>
          </table>
	  </td></tr>
	  </table>

	  <br/>
	 <table width="99%" border="0" align="center" cellspacing="0" cellpadding="0">  
	 <tr>
	 <td height="25" class='clefttitle'><strong> 五、此管理员在【<font color="#FF0000">模型字段管理</font>】的权限</strong></td>
	 </tr>
	 </table>
	  
	  <table width="96%" align="center" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>
          <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr> 
              <td height="25" colspan="2"> 
                <%
					IF instr(ModelPower,"model0")>0 Then
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('ModelPowers','none')"" name=""modelPower"" value=""model0"" checked>")
                    ELSE
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('ModelPowers','none')"" name=""modelPower"" value=""model0"">")
					END IF
					%>
                在模型字段管理管理权限(屏蔽)</td>
            </tr>
            <tr> 
              <td height="25" colspan="2"> 
                <%
					IF instr(ModelPower,"model1")>0 Then
					  Response.Write("<input type=""radio"" name=""modelPower"" onclick=""SetPowerListValue('ModelPowers','')"" value=""model1"" checked>")
                     ELSE
					  Response.Write("<input type=""radio"" name=""modelPower"" onclick=""SetPowerListValue('ModelPowers','')"" value=""model1"">")
					 END IF%>
                拥有指定的部分管理权限 </td>
            </tr>
            <tr ID="ModelPowers" <% IF instr(ModelPower,"model1")=0 then Response.Write("style=""display:none""") End IF%>> 
              <td height="25"> <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMM10000"<%if InStr(1, PowerList,"KSMM10000" ,1)<>0 then Response.Write( " checked") %>>
添加模型</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMM10001"<%if InStr(1, PowerList,"KSMM10001" ,1)<>0 then Response.Write( " checked") %>>
修改模型</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KSMM10002"<%if InStr(1, PowerList,"KSMM10002" ,1)<>0 then Response.Write( " checked") %>>
删除模型</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KSMM10003"<%if InStr(1, PowerList,"KSMM10003" ,1)<>0 then Response.Write( " checked") %>>
模型字段管理 </td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMM10004"<%if InStr(1, PowerList,"KSMM10004" ,1)<>0 then Response.Write( " checked") %>>
信息统计 </td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMM10005"<%if InStr(1, PowerList,"KSMM10005" ,1)<>0 then Response.Write( " checked") %>>
开启关闭</td>
                  </tr>
                
                </table></td>
            </tr>
          </table>
	  </td></tr>
	  </table>


	  <br/>
	 <table width="99%" border="0" align="center" cellspacing="0" cellpadding="0">  
	 <tr>
	 <td height="25" class='clefttitle'><strong> 六、此管理员在【<font color="#FF0000">相关选项</font>】的权限</strong></td>
	 </tr>
	 </table>
<table width="96%" align="center" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>
          <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr> 
              <td height="25" colspan="2"> 
                <%
					IF instr(ModelPower,"subsys0")>0 Then
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('SubSysPowers','none')"" name=""subsysPower"" value=""subsys0"" checked>")
                    ELSE
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('SubSysPowers','none')"" name=""subsysPower"" value=""subsys0"">")
					END IF
					%>
                在子系统管理权限(屏蔽)</td>
            </tr>
            <tr> 
              <td height="25" colspan="2"> 
                <%
					IF instr(ModelPower,"subsys1")>0 Then
					  Response.Write("<input type=""radio"" name=""subsysPower"" onclick=""SetPowerListValue('SubSysPowers','')"" value=""subsys1"" checked>")
                     ELSE
					  Response.Write("<input type=""radio"" name=""subsysPower"" onclick=""SetPowerListValue('SubSysPowers','')"" value=""subsys1"">")
					 END IF%>
                拥有指定的部分管理权限 </td>
            </tr>
            <tr ID="SubSysPowers" <% IF instr(ModelPower,"subsys1")=0 then Response.Write("style=""display:none""") End IF%>> 
              <td height="25"> <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                  
				  
				  <tr>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20000"<%if InStr(1, PowerList,"KSMS20000" ,1)<>0 then Response.Write( " checked") %>>
投诉建议管理</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20001"<%if InStr(1, PowerList,"KSMS20001" ,1)<>0 then Response.Write( " checked") %>>
友情链接管理</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20002"<%if InStr(1, PowerList,"KSMS20002" ,1)<>0 then Response.Write( " checked") %>>
网站公告管理</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20003"<%if InStr(1, PowerList,"KSMS20003" ,1)<>0 then Response.Write( " checked") %>>
站内调查管理 </td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20004"<%if InStr(1, PowerList,"KSMS20004" ,1)<>0 then Response.Write( " checked") %>>
网站留言管理</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20005"<%if InStr(1, PowerList,"KSMS20005" ,1)<>0 then Response.Write( " checked") %>>
站点计数器</td>                    
<td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20006"<%if InStr(1, PowerList,"KSMS20006" ,1)<>0 then Response.Write( " checked") %>>
广告系统管理</td>

                  </tr>
				 <tr>
				   <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20007"<%if InStr(1, PowerList,"KSMS20007" ,1)<>0 then Response.Write( " checked") %>>
推广计划查看</td>
				   <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20008"<%if InStr(1, PowerList,"KSMS20008" ,1)<>0 then Response.Write( " checked") %>>
心情指数管理</td>
				   <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS10006"<%if InStr(1, PowerList,"KSMS10006" ,1)<>0 then Response.Write( " checked") %>>
自定义表单管理</td>
				   <td><input name="PowerList" type="checkbox" id="PowerList" value="KSMS20009"<%if InStr(1, PowerList,"KSMS20009" ,1)<>0 then Response.Write( " checked") %>>
文档Digg管理</td>
             <td nowrap>
					<input name="PowerList" type="checkbox" id="PowerList" value="KSMS20010"<%if InStr(1, PowerList,"KSMS20010" ,1)<>0 then Response.Write( " checked") %>>
                      积分兑换商品</td>
             <td nowrap>
					<input name="PowerList" type="checkbox" id="PowerList" value="KSMS20014"<%if InStr(1, PowerList,"KSMS20014" ,1)<>0 then Response.Write( " checked") %>>
                     PK项目管理</td>
				 </tr>

                
                </table></td>
            </tr>
          </table>
	  </td></tr>
	  </table>
	  
	  <br/>
	 <table width="99%" border="0" align="center" cellspacing="0" cellpadding="0">  
	 <tr>
	 <td height="25" class='clefttitle'><strong> 七、此管理员在【<font color="#FF0000">插件管理</font>】的权限</strong></td>
	 </tr>
	 </table>
	  <table width="96%" align="center" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>
          <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr> 
              <td height="25" colspan="2"> 
                <%
					IF instr(ModelPower,"other0")>0 Then
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('otherPowers','none')"" name=""otherPower"" value=""other0"" checked>")
                    ELSE
					  Response.Write("<input type=""radio"" onclick=""SetPowerListValue('otherPowers','none')"" name=""otherPower"" value=""other0"">")
					END IF
					%>
                在插件管理权限(屏蔽)</td>
            </tr>
            <tr> 
              <td height="25" colspan="2"> 
                <%
					IF instr(ModelPower,"other1")>0 Then
					  Response.Write("<input type=""radio"" name=""otherPower"" onclick=""SetPowerListValue('otherPowers','')"" value=""other1"" checked>")
                     ELSE
					  Response.Write("<input type=""radio"" name=""otherPower"" onclick=""SetPowerListValue('otherPowers','')"" value=""other1"">")
					 END IF%>
                拥有指定的部分管理权限 </td>
            </tr>
            <tr ID="otherPowers" <% IF instr(ModelPower,"other1")=0 then Response.Write("style=""display:none""") End IF%>> 
              <td height="25"> <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSO10000"<%if InStr(1, PowerList,"KSO10000" ,1)<>0 then Response.Write( " checked") %>> 
                    WAP插件管理
</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="KSO10001"<%if InStr(1, PowerList,"KSO10001" ,1)<>0 then Response.Write( " checked") %>>
SK采集管理</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KSO10002"<%if InStr(1, PowerList,"KSO10002" ,1)<>0 then Response.Write( " checked") %>>
CC视频插件管理</td>
                    <td height="25"><input name="PowerList" type="checkbox" id="PowerList" value="KSO10003"<%if InStr(1, PowerList,"KSO10003" ,1)<>0 then Response.Write( " checked") %>>
Wss统计插件管理</td>
                    <td></td>
                    <td></td>
                  </tr>
                
                </table></td>
            </tr>
          </table>
	  </td></tr>
	  </table>

<script>
function SetPowerListValue(Module,Value)
{ $('#'+Module)[0].style.display=Value;
}
</script>
<%
		

		Response.Write "  </form>"
		Response.Write "</body>"
		Response.Write "</html>"
		Response.Write "<Script Language=""javascript"">" & vbCrLf
		Response.Write "<!--" & vbCrLf
		Response.Write "function CheckForm()" & vbCrLf
		Response.Write "{ var form=document.AdminForm;" & vbCrLf
		Response.Write "   if (form.UserName.value=='')" & vbCrLf
		Response.Write "    {"
		Response.Write "     alert(""请输入管理员名称!"");"
		Response.Write "     form.UserName.focus();"
		Response.Write "     return false;" & vbCrLf
		Response.Write "    }" & vbCrLf
			
			If Action <> "Edit" Then
		Response.Write "   if (form.PrUserName.value=='')" & vbCrLf
		Response.Write "    {"
		Response.Write "     alert(""请输入前台注册用户名称!"");"
		Response.Write "     form.PrUserName.focus();"
		Response.Write "     return false;" & vbCrLf
		Response.Write "    }" & vbCrLf

		Response.Write "    if (form.PassWord.value=='')"
		Response.Write "    {"
		Response.Write "     alert(""请输入初始密码!"");"
		Response.Write "     form.PassWord.focus();"
		Response.Write "     return false;"
		Response.Write "    }"
		Response.Write "   else if (form.PassWord.value.length<6)"
		Response.Write "    {"
		Response.Write "      alert(""初始密码不能少于6位!"");"
		Response.Write "     form.PassWord.focus();"
		Response.Write "     return false;"
		Response.Write "    }"

		
			End If
			
		Response.Write "   if (form.RealName.value=='')" & vbCrLf
		Response.Write "    {" & vbCrLf
		Response.Write "     alert(""请输入真实姓名"");" & vbCrLf
		Response.Write "     form.RealName.focus();" & vbCrLf
		Response.Write "     return false;" & vbCrLf
		Response.Write "    }" & vbCrLf
		Response.Write "   if (form.Email.value!='')" & vbCrLf
		Response.Write "   if(is_email(form.Email.value)==false)" & vbCrLf
		Response.Write "      { alert('非法电子邮箱!');" & vbCrLf
		Response.Write "        form.Email.focus();" & vbCrLf
		Response.Write "        return false;" & vbCrLf
		Response.Write "     }"
		Response.Write "    form.submit();" & vbCrLf
		Response.Write "    return true;" & vbCrLf
		Response.Write "}" & vbCrLf
		Response.Write "//-->" & vbCrLf
		Response.Write "</Script>"
		End Sub
		
		Sub AdminSave()
			Dim AdminID, GroupID, UserName,PrUserName, PassWord, ConPassWord, Locked, RealName, Sex, TelPhone, Email, Descript, TrueIP
			Dim TempObj, AdminRS, AdminSql,ComeUrl
			ComeUrl=Request.ServerVariables("HTTP_REFERER")
			AdminID = KS.G("AdminID")
			
			UserName = KS.R(KS.G("UserName"))
			PrUserName=KS.R(KS.G("PrUserName"))
			PassWord = KS.G("PassWord")

			IF PrUserName="" Then Call KS.Alert("前台注册用户名必须填写!",ComeUrl)
			
			PassWord = MD5(KS.R(PassWord),16)
			Locked = KS.G("Locked")
			RealName = KS.R(KS.G("RealName"))
			Sex = KS.G("Sex")
			TelPhone = KS.R(KS.G("TelPhone"))
			Email = KS.R(KS.G("Email"))
			Descript = KS.R(KS.G("Description"))
			TrueIP = KS.GetIP
			If UserName <> "" Then
					If Len(UserName) >= 100 Then
						Call KS.AlertHistory("管理员名称不能超过50个字符!", -1)
						Set KS = Nothing
						Response.End
					End If
			 Else
					Call KS.AlertHistory("请输入管理员名称!", -1)
					Set KS = Nothing
					Response.End
			 End If
			 
			Dim SQL,RSC:Set RSC=Server.CreateObject("ADODB.RECORDSET")
			RSC.Open "Select ChannelID,ChannelName,BasicType,ItemName,ModelEname,ChannelStatus From KS_Channel where channelstatus=1 Order By ChannelID",Conn,1,1
			If Not RSC.Eof Then
			  SQL=RSC.GetRows(-1)
			End If
			RSC.Close

			 For I=0 To Ubound(sql,2)
			  If I=0 Then
				 ModelPower=Replace(Request("ModelPower" & sql(0,i) &"")," ","")
			  Else
				 ModelPower=ModelPower & "," & Replace(Request("ModelPower" & sql(0,i) &"")," ","")
			  End IF
			 Next
			 ModelPower=request("otherpower") &"," & request("sysset") &"," & request("userpower") & "," & request("labpower") &"," &request("modelpower") & "," &request("subsyspower")&","&request("ask")&"," & request("space") & ","& modelpower
			 PowerList=Replace(Trim(KS.G("PowerList"))," ","")
			 IF PowerList="" THEN PowerList=0

			RSC.Open "Select AdminPurview,ID From KS_Class Order By ClassID",conn,1,3
			Do While Not RSC.Eof
			    
			  If KS.FoundInArr(Replace(Request("AdminPurview")," ",""),RSC(1),",") Then
			     If KS.IsNul(RSC(0)) Then 
				  RSC(0)=UserName
				 Else
				  RSC(0)=FilterRepeat(RSC(0) & "," & UserName,",")
				 End If
				 RSC.Update
			  Else
			     If KS.IsNul(RSC(0)) Then
				  RSC(0)=""
				 Else
					' If Instr(RSC(0),",")=0 Then
					'  RSC(0)=""
					' Else
					'  RSC(0)=Replace(Replace(RSC(0),UserName &",",""),","&UserName,"")
					' End If
					 RSC(0)=DelItemInArr(RSC(0),UserName,",")
				 End If
			   	 RSC.Update
			  End If
			     on error resume next
				 Dim ENode:Set ENode=Application(KS.SiteSN&"_class").DocumentElement.SelectSingleNode("class[@ks0='" & RSC(1) & "']")
				 ENode.SelectSingleNode("@ks16").text=RSC(0)
				 If err Then err.clear
			  
			  RSC.MoveNext
			loop
			RSC.Close
			Set RSC=nothing
		   
			   
			If Request("Action") = "Add" Then
					Set TempObj = Conn.Execute("Select UserName from [KS_Admin] where UserName='" & UserName & "'")
					If Not TempObj.EOF Then
						Call KS.Alert("数据库中已存在该管理员名称!", "AdminAdd.asp")
						 Set KS = Nothing
						Response.End
					End If
					Set TempObj = Conn.Execute("Select UserName from [KS_User] where UserName='" & PrUserName & "'")
					If TempObj.BOf And TempObj.EOF Then
						Call KS.Alert("找不到此前台注册用户!", ComeUrl)
						 Set KS = Nothing:Response.End
					End If
					IF Conn.Execute("Select Count(adminid) From KS_Admin Where PrUserName='" & PrUserName & "'")(0)>=1 Then
						Call KS.Alert("您填写的前台注册用户已经是管理员了，不能再添加!", ComeUrl)
						 Set KS = Nothing:Response.End
					End IF
				  Set AdminRS = Server.CreateObject("adodb.recordset")
				  AdminSql = "select * from [KS_Admin] Where (AdminID IS NULL)"
				  AdminRS.Open AdminSql, Conn, 1, 3
				  AdminRS.AddNew
				  AdminRS("AddDate") = Now
				  AdminRS("UserName") = UserName
				  AdminRS("PrUserName")=PrUserName
				  AdminRS("ModelPower")= ModelPower
				  AdminRS("PowerList")=PowerList
				  AdminRS("PassWord") = PassWord
				  AdminRS("Locked") = Locked
				  AdminRS("RealName") = RealName
				  AdminRS("Sex") = Sex
				  AdminRS("TelPhone") = TelPhone
				  AdminRS("Email") = Email
				  AdminRS("Description") = Descript
				  AdminRS("SuperTF") = 0
				  AdminRS("LastLoginIP") = TrueIP
				  AdminRS("LastLoginTime") = Now
				  AdminRS("LastLogOutTime") = Now
				  AdminRS("LoginTimes") = 0
				  AdminRS.Update
				  AdminRS.Close:Set AdminRS = Nothing
				  
				  '更新前台用户，使之加入管理员组
				  Conn.Execute("Update KS_User Set GroupID=1 Where UserName='" & PrUserName & "'")
				  
				  Response.Write ("<script>if (confirm('添加管理员成功,继续添加吗?')) {location.href='?Action=Add';} else { location.href='KS.Admin.asp';}</script>")
			ElseIf Request("Action") = "Edit" Then
					Set TempObj = Conn.Execute("Select UserName from [KS_Admin] where AdminID<>" & AdminID & " And UserName='" & UserName & "'")
					If Not TempObj.EOF Then
						Call KS.Alert("数据库中已存在该管理员名称!", "AdminAdd.asp?AdminID=" & AdminID & "&Action=Edit")
						 Set KS = Nothing
						Response.End
					End If
				  Set AdminRS = Server.CreateObject("adodb.recordset")
				  AdminSql = "select * from [KS_Admin] Where AdminID=" & AdminID
				  AdminRS.Open AdminSql, Conn, 1, 3
				  AdminRS("UserName") = UserName
				  AdminRS("Locked") = Locked
				  AdminRS("RealName") = RealName
				  AdminRS("Sex") = Sex
				  AdminRS("TelPhone") = TelPhone
				  AdminRS("Email") = Email
				  AdminRS("Description") = Descript
				  AdminRS("ModelPower")= ModelPower
				  AdminRS("PowerList")=PowerList
				  AdminRS.Update
				  AdminRS.Close:Set AdminRS = Nothing
				  Response.Write ("<script>alert('修改管理员成功!');location.href='KS.Admin.asp';</script>")
			End If
			
			
			
		End Sub
        
		'设置管理员密码
		Sub SetAdminPass()
		Response.Write "<html>"
		Response.Write "<head>"
		Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		Response.Write "<link href=""Include/ModeWindow.css"" rel=""stylesheet"">"
		Response.Write "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>"
		Response.Write "<script language=""JavaScript"" src=""../ks_inc/CheckPassWord.js""></script>"
		Response.Write "<style>"
		Response.Write ".rank { border:none; background:url(../images/rank.gif) no-repeat; width:125px; height:22px; vertical-align:middle; cursor:default; }"
		Response.Write ".r0 { background-position:0 0; }"
		Response.Write ".r1 { background-position:0 -22px; }"
		Response.Write ".r2 { background-position:0 -44px; }"
		Response.Write ".r3 { background-position:0 -66px; }"
		Response.Write ".r4 { background-position:0 -88px; }"
		Response.Write ".r5 { background-position:0 -110px; }"
		Response.Write ".r6 { background-position:0 -132px; }"
		Response.Write ".r7 { background-position:0 -154px; }"
		Response.Write "</style>"
		Response.Write "<title>设置管理员密码</title>"
		Response.Write "</head>"
		
		Dim AdminID, UserName, PassWord,RSObj
		AdminID = Request("AdminID")
		If AdminID <> "" Then
		   Set RSObj = Server.CreateObject("AdoDb.RecordSet")
		   RSObj.Open "Select * From KS_Admin Where AdminID=" & AdminID, Conn, 1, 3
		   If Not RSObj.EOF Then UserName = Trim(RSObj("UserName"))
		   RSObj.Close:Set RSObj = Nothing
		Else
		  UserName=KS.C("AdminName")
		End If
		
		     If Request("Flag") = "SetOK" Then
			   If Trim(Request.Form("PassWord")) <> Trim(Request.Form("ConPassWord")) Then
				Response.Write ("<Script>alert('确认密码有误!');history.back(-1);</script>")
				Exit Sub
				Response.End
			   Else
			    Set RSObj = Server.CreateObject("AdoDb.RecordSet")
		         RSObj.Open "Select * From KS_Admin Where UserName='" & UserName & "'", Conn, 1, 3
				 RSObj("PASSWord") = MD5(KS.R(Trim(Request.Form("PassWord"))),16)
				 RSObj.Update
				 Dim PrUserName:PrUserName=RSObj("PrUserName")
				  RSObj.Close: Set RSObj = Nothing
				  If UserName=KS.C("UserName") Then  Response.Cookies(KS.SiteSn)("AdminPass")=MD5(KS.R(Trim(Request.Form("PassWord"))),16)
				  
				  If KS.ChkClng(request("UpdateUserPass"))=1 Then
				    Conn.Execute("Update KS_User Set [PassWord]='" &MD5(KS.R(Trim(Request.Form("PassWord"))),16) &"' Where UserName='" & PrUserName & "'")
					Response.Cookies(KS.SiteSn)("Password")=MD5(KS.R(Trim(Request.Form("PassWord"))),16)
				  End If
				  
				 Response.Write ("<Script>alert('密码设置成功!!!');window.close();</script>")
			   End If
			 End If
			 
		Response.Write "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
		Response.Write "  <form action=""?Action=SetPass"" name=""AdminForm"" method=""post"">"
		Response.Write "   <input name=""Flag"" type=""hidden"" id=""Flag"" value=""SetOK"">"
		Response.Write "   <input name=""AdminID"" type=""hidden"" value=""" & AdminID & """><br>"
		Response.Write "  <table width=""99%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">"
		Response.Write "    <tr>"
		Response.Write "      <td>"
		Response.Write "      <FIELDSET align=center>" & vbCrLf
		Response.Write "        <LEGEND align=left>设置密码</LEGEND>" & vbCrLf
		Response.Write "        <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbCrLf
		Response.Write "          <tr>"
		Response.Write "            <td width=""179"" height=""25"" align=""center""> <div align=""center"">管理员名</div></td>" & vbCrLf
		Response.Write "            <td width=""542"" height=""25"">"
		Response.Write ("<input name=""UserName"" Readonly value=""" & UserName & """ type=""text"" id=""UserName"" size=""30"" class=""textbox"">")
		Response.Write "              用于登录后台的名称</td>" & vbCrLf
		Response.Write "          </tr>"
		
		Response.Write "          <tr>"
		Response.Write "            <td height=""25"" align=""center""> <div align=""center"">新 密 码</div></td>"
		Response.Write "           <td height=""25""> <input name=""PassWord"" type=""password"" size=""34"" class=""textbox"" onkeyup=""javascript:setPasswordLevel(this, document.getElementById('passwordLevel'))"" onmouseout=""javascript:setPasswordLevel(this, document.getElementById('passwordLevel'))"" onblur=""javascript:setPasswordLevel(this, document.getElementById('passwordLevel'))"">" & vbCrLf
		Response.Write "              不少于6位 </td>"
		Response.Write "          </tr>"
		Response.Write "          <tr>"
		Response.Write "             <td align=""center"" height=""25"">密码强度</td>"
		Response.Write "                <td><input name=""Input"" disabled=""disabled"" class=""rank r0"" id=""passwordLevel"" /></td>"
		Response.Write "           </tr>"
		Response.Write "          <tr>"
		Response.Write "            <td height=""25"" align=""center"">确定密码</td>" & vbCrLf
		Response.Write "            <td height=""25""> <input name=""ConPassWord""  type=""password"" size=""34"" class=""textbox"">" & vbCrLf
		Response.Write "              同上</td>"
		Response.Write "          </tr>"
		
		Response.Write "          <tr>"
		Response.Write "            <td height=""25"" align=""center"">更新前台密码</td>" & vbCrLf
		Response.Write "            <td height=""25""> <label><input name=""UpdateUserPass""  type=""checkbox"" value=""1"" checked>是</label> <font color=red>如果选择是,那么前台会员中心的登录密码将和后台的一致</font></td>"
		Response.Write "          </tr>"
		
		Response.Write "        </table>"
		Response.Write "         </FIELDSET>" & vbCrLf
		Response.Write "       </td>" & vbCrLf
		Response.Write "    </tr>"
		Response.Write "    </table>" & vbCrLf
		Response.Write "  <table width=""100%"" height=""30"" border=""0"" cellpadding=""0"" cellspacing=""0"">" & vbCrLf
		Response.Write "    <tr>"
		Response.Write "      <td height=""40"" align=""center"">" & vbCrLf
		Response.Write "        <input type=""button"" class='textbox' name=""Submit"" Onclick=""CheckForm()"" value="" 确 定 "">"
		Response.Write "        <input type=""button"" class='textbox' name=""Submit2"" onclick=""window.close()"" value="" 取 消 "">" & vbCrLf
		Response.Write "      </td>" & vbCrLf
		Response.Write "    </tr>"
		Response.Write "  </table>"
		Response.Write "  </form>" & vbCrLf
		Response.Write "</body>"
		Response.Write "</html>" & vbCrLf
		Response.Write "<Script Language=""javascript"">" & vbCrLf
		Response.Write "<!--" & vbCrLf
		Response.Write "function CheckForm()" & vbCrLf
		Response.Write "{ var form=document.AdminForm;" & vbCrLf
		Response.Write "    if (form.PassWord.value=='')" & vbCrLf
		Response.Write "    {" & vbCrLf
		Response.Write "     alert(""请输入新密码!"");" & vbCrLf
		Response.Write "     form.PassWord.focus();" & vbCrLf
		Response.Write "     return false;" & vbCrLf
		Response.Write "    }" & vbCrLf
		Response.Write "    else if (form.PassWord.value.length<6)" & vbCrLf
		Response.Write "    {" & vbCrLf
		Response.Write "      alert(""初始密码不能少于6位!"");"
		Response.Write "     form.PassWord.focus();" & vbCrLf
		Response.Write "     return false;" & vbCrLf
		Response.Write "    }" & vbCrLf
		Response.Write "   if (form.ConPassWord.value=='')" & vbCrLf
		Response.Write "    {"
		Response.Write "     alert(""请输入确定密码!"");" & vbCrLf
		Response.Write "     form.ConPassWord.focus();" & vbCrLf
		Response.Write "     return false;" & vbCrLf
		Response.Write "    }" & vbCrLf
		Response.Write "   else if(form.ConPassWord.value.length<6)" & vbCrLf
		Response.Write "    {"
		Response.Write "     alert(""确定密码不能少于6位!"");" & vbCrLf
		Response.Write "     form.ConPassWord.focus();" & vbCrLf
		Response.Write "     return false;" & vbCrLf
		Response.Write "    }"
		Response.Write "   if (form.PassWord.value!=form.ConPassWord.value)" & vbCrLf
		Response.Write "    {"
		Response.Write "    alert(""两次输入的密码不一致!"");" & vbCrLf
		Response.Write "     form.PassWord.focus();" & vbCrLf
		Response.Write "     return false;" & vbCrLf
		Response.Write "    }" & vbCrLf
		Response.Write "    form.submit();" & vbCrLf
		Response.Write "    return true;" & vbCrLf
		Response.Write "}"
		Response.Write "//-->"
		Response.Write "</Script>"
		End Sub
		
		'删除管理员
		Sub AdminDel()
		Dim k, AdminID,RSObj
		AdminID = Trim(KS.G("AdminID"))
		AdminID = Split(AdminID, ",")
		For k = LBound(AdminID) To UBound(AdminID)
			   Set RSObj = Conn.Execute("Select SuperTF,PrUserName From KS_Admin Where  AdminID=" & AdminID(k))
			   If Not RSObj.EOF Then
				 If RSObj("SuperTF") = 1 Then
				  Response.Write ("<script>alert('系统默认管理员不能删除!');location.href='KS.Admin.asp';</script>")
				 Else
				  '更改前台注册会员，使之变为注册会员身份
				  Conn.Execute("Update KS_User Set GroupID=3 Where UserName='" & RSObj("PrUserName") & "'")
				  Conn.Execute ("Delete From KS_Admin Where AdminID =" & AdminID(k))
				 End If
			  End If
			  RSObj.Close
		Next
		Set RSObj = Nothing
		Response.Write ("<script>location.href='KS.Admin.asp';</script>")
		End Sub
		
		
 Sub ClassList(ChannelID)
 %>
 <div style="border: 5px solid #E7E7E7;height:150; overflow: auto; width:100%"> 
                        <table border="0" cellspacing="0" cellpadding="0">
                          <% 
					  Dim Node, CheckStr,SpaceStr,TJ,k  
				      KS.LoadClassConfig
					  For Each Node In Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks12=" & ChannelID&"]")                     
	                  if username<>"" and KS.FoundInArr(Node.SelectSingleNode("@ks16").text,username,",")=true then CheckStr=" checked"
					  SpaceStr="&nbsp;&nbsp;&nbsp;&nbsp;"
					  TJ=Node.SelectSingleNode("@ks10").text
					  If TJ>1 Then
						 For k = 1 To TJ - 1
							SpaceStr = SpaceStr & "&nbsp;&nbsp;&nbsp;&nbsp;"
						 Next
					    SpaceStr = SpaceStr &"<img src=""Images/Folder/folderclosed.gif"">"
					  Else
					    spacestr="<img src=""Images/Folder/domain.gif"" width=""26"" height=""24"">"
					  End If
					  %>
                          
                          <tr> 
                            <td><table border="0" cellspacing="0" cellpadding="0">
                                <tr align="left" class="TempletItem"> 
                                  <td><%=SpaceStr%></td>
                                  <td><input name="AdminPurview" type="checkbox" value="<% =Node.SelectSingleNode("@ks0").text %>"<%=CheckStr%>> 
                                    <% = Node.SelectSingleNode("@ks1").text %>                                     
								 </td>
                                </tr>
                              </table></td>
                          </tr>
                          <%
	                     CheckStr = ""
	                Next
					   %>
                        </table>
                      </div>
 <%
 End Sub
 
 Sub BasePurview(PowerList,SQL,I)
 %>
     <tr> 
                    <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10002"<%if InStr(1, PowerList,"M" & SQL(0,I)&"10002" ,1)<>0 then Response.Write( " checked") %>>
                      添加<%=sql(3,i)%></td>
                    <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10003"<%if InStr(1, PowerList,"M"&SQL(0,I)&"10003" ,1)<>0 then Response.Write( " checked") %>>
                      编辑<%=sql(3,i)%></td>
                    <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10004"<%if InStr(1, PowerList,"M" & SQL(0,I) &"10004" ,1)<>0 then Response.Write( " checked") %>>
                      删除<%=sql(3,i)%></td>
                    <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10005"<%if InStr(1, PowerList,"M" & SQL(0,I) &"10005" ,1)<>0 then Response.Write(" checked") %>>
批量设置</td>
                    <td width="13%"><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10006"<%if InStr(1, PowerList,"M" & SQL(0,I) & "10006" ,1)<>0 then Response.Write(" checked") %>>
                      加入专题</td>
                    <td width="13%"> <input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10007"<%if InStr(1, PowerList,"M" & SQL(0,I) & "10007" ,1)<>0 then Response.Write(" checked") %>>
                      加入JS</td>
                  </tr>
                  <tr> 
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10008"<%if InStr(1, PowerList,"M" & SQL(0,I) &"10008" ,1)<>0 then Response.Write( " checked") %>>
                      搜索<%=sql(3,i)%></td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10009"<%if InStr(1, PowerList,"M" & SQL(0,I) & "10009" ,1)<>0 then Response.Write(" checked") %>>
上传文件</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10011"<%if InStr(1, PowerList,"M" & SQL(0,I) & "10011" ,1)<>0 then Response.Write(" checked") %>>
复制粘贴</td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>10012"<%if InStr(1, PowerList,"M" & SQL(0,I) & "10012" ,1)<>0 then Response.Write(" checked") %>>
签收<%=sql(3,i)%></td>
                    <td><input name="PowerList" type="checkbox" id="PowerList" value="M<%=SQL(0,I)%>20005"<%if InStr(1, PowerList,"M" & SQL(0,I) & "20005" ,1)<>0 then Response.Write(" checked") %>>
发布<%=sql(3,i)%></td>
                    <td></td>
                  </tr>
<%
 End Sub
 
 '去除数组的重复项
 Function FilterRepeat(byval str,spliter)
   if KS.IsNul(str) Then Exit Function
   Dim strA:strA=Split(str,spliter)
   Dim I,temp,newstr
   For I=0 To Ubound(Stra)
      If KS.FoundInArr(temp,strA(i),",")=false Then
	    if newstr="" then
		 newstr=stra(i)
		else
		 newstr=newstr & spliter & stra(i)
		end if
		temp=temp & "," & stra(i)
	  End If
   Next
   FilterRepeat=newstr
 End Function
 
 '从数组中删除指定项
 Function DelItemInArr(byval str,byval delstr,spliter)
   if KS.IsNul(str) Then Exit Function
   Dim strA:strA=Split(str,spliter)
   Dim I,temp,newstr
   For I=0 To Ubound(Stra)
      If lcase(strA(i))<>lcase(delstr) Then
	    if newstr="" then
		 newstr=stra(i)
		else
		 newstr=newstr & spliter & stra(i)
		end if
	  End If
   Next
   DelItemInArr=newstr
 End Function
 
End Class
%> 
