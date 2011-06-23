<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
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
Set KSCls = New Admin_UserForm
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_UserForm
        Private KS,Action,ComeUrl, Page
		Private I,totalPut,CurrentPage,SqlStr,RS,MaxPerPage,KSCls
		Private Sub Class_Initialize()
		  MaxPerPage =18
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub


		Public Sub Kesion()
		
		  With Response
			.Write "<html>"
			.Write "<head>"
			.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
			.Write "<title>会员表单管理</title>"
			.Write "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
			
             Action=KS.G("Action")
			 
			 Page=KS.G("Page")
			 
			 Select Case Action
			  Case "Add"
			    Call UserFormManage("Add")
			  Case "Edit"
			   Call UserFormManage("Edit")
			  Case "Del"
			    Call UserFormDel()
			  Case "FormSave"
			    Call UserFormAddSave()
			  Case Else
			   Call UserFormList()
			 End Select
			.Write "</body>"
			.Write "</html>"
		  End With
		End Sub
			  
			 
		Sub UserFormList()			
			If KS.G("page") <> "" Then
				  CurrentPage = CInt(KS.G("page"))
			Else
				  CurrentPage = 1
			End If
			With Response
            .Write "<script language='JavaScript'>"
			.Write "var Page='" & CurrentPage & "';"
			.Write "</script>"
			.Write "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>" & vbCrLf
	        .Write "<script language=""JavaScript"" src=""../KS_Inc/jQuery.js""></script>" & vbCrLf
			%>
			<script language="javascript">
			    function set(v)
				{
				 if (v==1)
				 UserFormControl(1);
				 else if (v==2)
				 UserFormControl(2);
				}
				function UserFormAdd()
				{
					location.href='KS.UserForm.asp?Action=Add';
					window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=会员系统 >> 会员表单管理 >> <font color=red>新增会员表单</font>&ButtonSymbol=GO';
				}
				function EditUserForm(id)
				{
					location="KS.UserForm.asp?Page="+Page+"&Action=Edit&ID="+id;
					window.$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=会员系统 >> 会员表单管理 >> <font color=red>编辑会员表单</font>&ButtonSymbol=GOSave';
				}
				function DelUserForm(id)
				{
				if (confirm('真的要删除该会员表单吗?'))
				 location="KS.UserForm.asp?Action=Del&Page="+Page+"&id="+id;
				   SelectedFile='';
				}
				function UserFormControl(op)
				{  var alertmsg='';
	               var ids=get_Ids(document.myform);
					if (ids!='')
					 {  if (op==1)
						{
						if (ids.indexOf(',')==-1) 
							EditUserForm(ids)
						  else alert('一次只能编辑一个会员表单tags!')	 
						}	
					  else if (op==2)    
						 DelUserForm(ids);
					 }
					else 
					 {
					 if (op==1)
					  alertmsg="编辑";
					 else if(op==2)
					  alertmsg="删除"; 
					 else
					  {
					  alertmsg="操作" 
					  }
					 alert('请选择要'+alertmsg+'的会员表单');
					  }
				}
				function GetKeyDown()
				{ 
				if (event.ctrlKey)
				  switch  (event.keyCode)
				  {  case 90 : location.reload(); break;
					 case 65 : Select(0);break;
					 case 78 : event.keyCode=0;event.returnValue=false; UserFormAdd();break;
					 case 69 : event.keyCode=0;event.returnValue=false;UserFormControl(1);break;
					 case 68 : UserFormControl(2);break;
				   }	
				else	
				 if (event.keyCode==46)UserFormControl(2);
				}
			</script>
			<%
			.Write "<body scroll=no topmargin='0' leftmargin='0'  onkeydown='GetKeyDown();' onselectstart='return false;'>"
			.Write "<ul id='menu_top'>"
			.Write "<li class='parent' onClick=""UserFormAdd();""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>添加会员表单</span></li>"
			.Write "<li class='parent' onClick=""UserFormControl(1);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>编辑会员表单</span></li>"
			.Write "<li class='parent' onClick=""UserFormControl(2);""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>删除会员表单</span></li>"
			.Write "</ul>"
            .Write ("<div style=""height:94%; overflow: auto; width:100%"" align=""center"">") 
			.Write ("<form name='myform' method='Post' action='?'>")
	        .Write ("<input type='hidden' name='action' id='action' value='" & Action & "'>")
			.Write "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
			.Write "        <tr>"
			.Write "          <td class=""sort"" width='35' align='center'>选择</td>"
			.Write "          <td class='sort' align='center'>会员表单名称</td>"
			.Write "          <td width='19%' class='sort' align='center'>创建时间</td>"
			.Write "          <td width='26%' class='sort' align='center'>管理操作</td>"
			.Write "  </tr>"
			  
			  Set RS = Server.CreateObject("ADODB.RecordSet")
					   SqlStr = "SELECT * FROM [KS_UserForm] order by AddDate desc"
					   RS.Open SqlStr, conn, 1, 1
					 If RS.EOF And RS.BOF Then
					 Else
						totalPut = RS.RecordCount
			
								If CurrentPage < 1 Then
									CurrentPage = 1
								End If
			
								If (CurrentPage - 1) * MaxPerPage > totalPut Then
									If (totalPut Mod MaxPerPage) = 0 Then
										CurrentPage = totalPut \ MaxPerPage
									Else
										CurrentPage = totalPut \ MaxPerPage + 1
									End If
								End If
			
								If CurrentPage = 1 Then
									Call showContent
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
										
										Call showContent
									Else
										CurrentPage = 1
										Call showContent
										
									End If
								End If
				End If
			.Write "</table>"
			.Write ("<table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>")
	        .Write ("<tr><td width='180'><div style='margin:5px'><b>选择：</b><a href='javascript:Select(0)'><font color=#999999>全选</font></a> - <a href='javascript:Select(1)'><font color=#999999>反选</font></a> - <a href='javascript:Select(2)'><font color=#999999>不选</font></a> </div>")
	        .Write ("</td>")
	        .Write ("<td><select style='height:18px' onchange='set(this.value)' name='setattribute'><option value=0>快速选项...</option><option value='1'>编辑会员表单</option><option value='2'>执行删除</option></select></td>")
	        .Write ("</form><td align='right'>")
	         Call KSCLS.ShowPage(totalPut, MaxPerPage, "KS.UserForm.asp", True, "个", CurrentPage, "")
	        .Write ("</td></tr></table>")
			.Write "</div>"
            End With
			End Sub
			
			Sub showContent()
			   With Response
					Do While Not RS.EOF
			          .Write "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" & RS("ID") & "' onclick=""chk_iddiv('" & RS("ID") & "')"">"
			          .Write "<td class='splittd' align=center><input name='id'  onclick=""chk_iddiv('" &RS("ID") & "')"" type='checkbox' id='c"& RS("ID") & "' value='" &RS("ID") & "'></td>"
					  .Write "<td class='splittd' height='20'><span FormID='" & RS("ID") & "' ondblclick=""EditUserForm(this.FormID)""><img src='Images/key.gif' align='absmiddle'>"
					  .Write "  <span style='cursor:default;'>" & RS("FormName") & "</span></span></td>"
					  .Write "  <td class='splittd' align='center'>" & RS("AddDate") & " </td>"
					  .Write "  <td class='splittd' align='center'><a href='javascript:EditUserForm(" & RS("ID") & ")'>修改</a> | <a href='javascript:DelUserForm(" & RS("id") & ")'>删除</a></td>"
					  .Write "</tr>"
					  I = I + 1
					  If I >= MaxPerPage Then Exit Do
						   RS.MoveNext
					Loop
					  RS.Close
				End With
			End Sub
			
			
			Sub UserFormManage(OpType)
			With Response
			 Dim Action, FormName, ID, SqlStr,Page,Template,FormField,Note,AutoCheck,WapTemplate
			  ID = KS.G("ID"):Page = KS.G("Page"):AutoCheck=" checked"
			 If OpType = "Edit" Then
				 Set RS = Server.CreateObject("ADODB.RECORDSET")
				 SqlStr = "Select * From [KS_UserForm] Where ID=" & ID
				 RS.Open SqlStr, conn, 1, 1
				 If Not RS.EOF Then 
				  FormName = RS("FormName")
				  Template = RS("Template")
				  WapTemplate = RS("WapTemplate")
				  FormField= RS("FormField")
				  Note     = RS("Note")
				 End If
				 RS.Close:Set RS=Nothing
				 AutoCheck=""
			 End If
			.Write "<ul id='menu_top'>"
			.Write "<li class='parent' onclick=""return(CheckForm())""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/save.gif' border='0' align='absmiddle'>确定保存</span></li>"
			.Write "<li class='parent' onclick=""location.href='?';""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/back.gif' border='0' align='absmiddle'>取消返回</span></li>"
		    .Write "</ul>"
			.Write "<table width='100%' border='0' align='center' cellpadding='1'  cellspacing='1' class='ctable' style='border-collapse: collapse'>" & vbCrLf
			.Write " <form  action='?action=FormSave&ID=" & ID & "' method='post' name='myform' onsubmit='return(CheckForm())'>" & vbCrLf
			.Write "    <tr class='tdbg'>" & vbCrLf
			.Write "      <td width='160' class='clefttitle' align='right'><strong>表单名称：</strong></td>" & vbCrLf
			.Write "      <td height='30' nowrap><input name='FormName' type='text' id='FormName' value='" & FormName & "' style='width:200;border-style: solid; border-width: 1'></b>*(如：企业注册填写表单) </td>" & vbCrLf
			.Write "    </tr>"
			.Write "    <tr class='tdbg'>" & vbCrLf
			.Write "      <td class='clefttitle' align='right'><strong>选择字段：</strong><br><font color=red>(<input type='checkbox' name='autoform' value='1' " & AutoCheck& ">自动生成录入表单)</font></td>" & vbCrLf
			.Write "      <td height='30' nowrap>"
			.Write "      <table border='0' wdith='100%' cellspacing='1' cellpadding='1'>"
			.Write "        <tr>"
			.Write "         <td align='center' class='sort' width='100'>选 择</td>"
			.Write "         <td align='center' class='sort' width='150'>字段别名</td>"
			.Write "         <td align='center' class='sort' width='130'>字段名</td>"
			.Write "         <td align='center' class='sort' width='150'>字段类型</td>"
			.Write "         <td align='center' class='sort' width='50'>级 别</td>"
			.Write "         <td align='center' class='sort' width='90'>注册显示</td>"
			.Write "         <td align='center' class='sort' width='60'>排 序</td>"
			.Write "        </tr>"
			Dim RSF:Set RSF=Server.CreateObject("ADODB.Recordset")
			RSF.Open "Select * From KS_Field Where ChannelID=101 order by orderid ",conn,1,1
			Do While Not RSF.Eof
			 .Write "<tr>"
			 .Write "<td class='splittd' align='center'>"
			 If Instr(FormField,RSF("FieldID"))<>0 Then
			 .Write "<Input type='checkbox' name='FieldID' checked value='" & RSF("FieldID") & "'>"
			 Else
			 .Write "<Input type='checkbox' name='FieldID' value='" & RSF("FieldID") & "'>"
			 End If
			 .Write "</td>"
			 .Write "<td class='splittd' align='center'>" & RSF("Title") & "</td>"
			 .Write "<td class='splittd' align='center'>" & RSF("FieldName") & "</td>"
			 .Write "<td class='splittd' align='center'>" 
				 Select Case RSF("FieldType")
				  Case 1:.Write "单行文本(text)"
				  Case 2:.Write "文本(不支持HTML)"
				  Case 10:.Write "多行文本(支持HTML)"
				  Case 3:.Write "下拉列表(select)"
				  Case 4:.Write "数字(text)"
				  Case 5:.Write "日期(text)"
				  Case 6:.Write "单选框(radio)"
				  Case 7:.Write "复选框(checkbox)"
				  Case 8:.Write "电子邮箱(text)"
				  Case 9:.Write "文件(text)"
				 End Select
			 .Write "</td>"
			 .Write "<td class='splittd' align='center'>" 
			  If Left(Lcase(RSF("FieldName")),3)<>"ks_" Then
				.Write "<font color=red>系统</font>"
		      Else
				.Write "<font color=blue>自定义</font>"
			  End If
			 .Write "</td>"
			 .Write "<td class='splittd' align='center'>" 
			 If RSF("ShowOnUserForm")="1" Then
			  .Write "<font color=red>是</font>"
			 Else
			  .Write "<font color=blue>否</font>"
			 End If
			  .Write "</td>"

			 .Write "<td class='splittd' align='center'>" & RSF("orderid") & "</td>"
			 .Write "</tr>"
			 
			 RSF.MoveNext
			Loop
			RSF.Close:Set RSF=Nothing
			.Write "      </table>"
			
			.Write "     </td>" & vbCrLf
			.Write "    </tr>"
			.Write "    <tr class='tdbg'>" & vbCrLf
			.Write "      <td width='160' class='clefttitle' align='right'><strong>表单模板：</strong><br>(<font color=blue>留空将自动生成录入表单模板</font>)<br><font color=green>如果某个字段在会员注册时不让填写或只能由管理员手工填写，可以加标签<font color=red>{@NoDisplay(字段名称)}</font><br>后台打印不显示的地方可以加<font color=red>{@NoDisplay}</font></font></td>" & vbCrLf
			.Write "      <td height='30' nowrap><textarea name='Template' cols='80' rows='10'>" & server.htmlencode(Template) & "</textarea></td>" & vbCrLf
			.Write "    </tr>"
			If KS.WSetting(0)="1" Then
			.Write "    <tr class='tdbg'>" & vbCrLf
			.Write "      <td width='160' class='clefttitle' align='right'><strong>WAP表单模板：</strong></td>" & vbCrLf
			.Write "      <td height='30' nowrap><textarea name='WapTemplate' cols='80' rows='4'>" & Server.htmlencode(WapTemplate) & "</textarea></td>" & vbCrLf
			.Write "    </tr>"
			End If
			.Write "    <tr class='tdbg'>" & vbCrLf
			.Write "      <td width='160' class='clefttitle' align='right'><strong>备注说明：</strong></td>" & vbCrLf
			.Write "      <td height='30' nowrap><textarea name='Note' cols='80' rows='4'>" & Note & "</textarea> </td>" & vbCrLf
			.Write "    </tr>"
			.Write "  </form>" & vbCrLf
			.Write "</table>" & vbCrLf

			.Write "<Script Language='javascript'>" & vbCrLf
			.Write "<!--" & vbCrLf
			.Write "function CheckForm()" & vbCrLf
			.Write "{ var form=document.myform;" & vbCrLf
			.Write "   if (form.FormName.value=='')" & vbCrLf
			.Write "    {" & vbCrLf
			.Write "     alert('请输入会员表单!');" & vbCrLf
			.Write "     form.FormName.focus();" & vbCrLf
			.Write "     return false;" & vbCrLf
			.Write "    }" & vbCrLf
			.Write "    form.submit();" & vbCrLf
			.Write "    return true;" & vbCrLf
			.Write "}" & vbCrLf
			.Write "//-->" & vbCrLf
			.Write "</Script>" & vbCrLf
			.Write "</body>" & vbCrLf
			.Write "</html>" & vbCrLf
			End With
			End Sub
			
			'构造录入表单
			Function GetTemplate()
			 If KS.G("autoform")="1"  Then
				 Dim FieldID:FieldID=KS.FilterIDs(KS.G("FieldID"))
				 If FieldID="" Then FieldID=0
				 GetTemplate="<table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""1"">" & vbcrlf
				 GetTemplate=GetTemplate & KSCls.Get_KS_D_F_P(101,0," and FieldID in(" & FieldID & ")")
				 GetTemplate=GetTemplate & "</table>" & vbcrlf
			 Else
			     GetTemplate=Request.Form("Template")
			 End IF
			End Function
			
			Sub UserFormAddSave()
			    Dim FormName:FormName = KS.G("FormName")
				  If FormName = "" Then Call KS.AlertHistory("请输入会员表单名称!", -1):Response.End()
				 Dim RS:Set RS = Server.CreateObject("ADODB.RECORDSET")
				 SqlStr = "Select * From [KS_UserForm] Where ID=" & KS.ChkClng(KS.G("ID"))
				 RS.Open SqlStr, conn, 3, 3
				 If RS.EOF Then
				   	  RS.AddNew
				 End If
					  RS("FormName") = FormName
					  RS("FormField") = Replace(Request("FieldID")," ","")
					  RS("AddDate") = Now()
					  RS("Template")=GetTemplate()
					  If KS.WSetting(0)="1" Then
					     If KS.G("autoform")="1"  Then
					        Dim WapTemplate
							WapTemplate=Replace(GetTemplate,"</tr>","[br]")
							WapTemplate=KS.LoseHtml(WapTemplate)
							WapTemplate=Replace(WapTemplate,Chr(8),vbNullString)
							WapTemplate=Replace(WapTemplate,Chr(9),vbNullString)
							WapTemplate=Replace(WapTemplate,Chr(10),vbNullString)
							WapTemplate=Replace(WapTemplate,Chr(13),vbNullString)
							WapTemplate=Replace(WapTemplate,"[br]","<br/>"&vbcrlf)
						 Else
					        WapTemplate=Request.Form("WapTemplate")
						 End If
						 RS("WapTemplate")=WapTemplate
					  End If
					  RS("Note") = KS.G("Note")
					  RS.Update
				 RS.Close:Set RS=Nothing
				 If KS.ChkClng(KS.G("ID"))=0 Then
				 Response.Write ("<Script> if (confirm('会员表单增加成功,继续添加吗?')) { location.href='?Action=Add';} else{location.href='?Page=" & Page &"';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=会员系统 >> 会员表单管理&ButtonSymbol=Disabled';}</script>")
				 Else
				 Response.Write ("<Script>alert('恭喜，表单修改成功!');location.href='?Page=" & Page &"';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=会员系统 >> 会员表单管理&ButtonSymbol=Disabled';</script>")
				 End If

			End Sub
			
			Sub UserFormDel()
			  Dim Page:Page=KS.G("Page")
			  Conn.Execute("Delete from [KS_UserForm] Where ID in(" & KS.FilterIDs(KS.G("ID")) & ")")
			  Response.Redirect "?Page=" & Page
			End Sub
End Class
%>
 
