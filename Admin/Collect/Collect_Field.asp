<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%> 
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Admin_Field
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Field
        Private KS,Action,ChannelID,Page,ItemName,TableName,KSCls
		Private I, totalPut, CurrentPage, FieldSql, FieldRS,MaxPerPage
		Private FieldName,ID,Contact, Title, Tips, FieldType, DefaultValue, MustFillTF, ShowOnForm, ShowOnUserForm,Options,OrderID,AllowFileExt,MaxFileSize,Width

		Private Sub Class_Initialize()
		  MaxPerPage = 30
		  Set KSCls=New ManageCls
		  Set KS=New PublicCls
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
		.Write "<title>�ֶι���</title>"
		.Write "<link href='../Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
             Action=KS.G("Action")
			 ChannelID=KS.ChkClng(KS.G("ChannelID"))
			 
			 TableName=KS.C_S(ChannelID,2)
			 If ChannelID=101 Then TableName="KS_User"   '��Ա��
			 ItemName=KS.C_S(ChannelID,3)
			 Page=KS.G("Page")
			 
			If Not KS.ReturnPowerResult(0, "M010008") Then                  'Ȩ�޼��
				Call KS.ReturnErr(1, "")   
				Response.End()
			End if
			 
			 Select Case Action
			  Case "SetCollect"
			    Call FieldSetCollect()
			  Case Else
			   Call FieldList()
			 End Select
			.Write "</body>"
			.Write "</html>"
		 End With
		End Sub
		
		Sub FieldList()
		 On Error Resume Next
		If Not IsEmpty(KS.G("page")) Then
			  CurrentPage = KS.G("page")
		Else
			  CurrentPage = 1
		End If
		With Response
		.Write "<script language='JavaScript' src='../KS_Inc/common.js'></script>"
		.Write "</head>"
		.Write "<body scroll=no topmargin='0' leftmargin='0'>"
		.Write "<ul id='menu_top'>"
		Response.Write "<li class='parent' onclick='location.href=""Collect_ItemModify.asp?channelid=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/a.gif' border='0' align='absmiddle'>�½���Ŀ</span></li>"
		.Write "<li class='parent' onclick='location.href=""Collect_ItemFilters.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/move.gif' border='0' align='absmiddle'>��������</span></li>"
		.Write "<li class='parent' onclick='location.href=""Collect_IntoDatabase.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/save.gif' border='0' align='absmiddle'>������</span></li>"
		.Write "<li class='parent' onclick='location.href=""Collect_ItemHistory.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/Recycl.gif' border='0' align='absmiddle'>��ʷ��¼</span></li>"
		.Write "<li class='parent' onclick='location.href=""Collect_Field.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/addjs.gif' border='0' align='absmiddle'>�Զ����ֶ�</span></li>"
		.Write "<li class='parent' onclick='location.href=""Collect_main.asp?ChannelID=" & ChannelID & """;'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='../images/ico/back.gif' border='0' align='absmiddle'>����һ��</span></li>"
		.Write "</ul>"
        
		.Write ("<div style=""height:94%; overflow: auto; width:100%"" align=""center"">")
		
		.Write "<div style='text-align:right'>�밴ģ������Ҫ���õ��Զ����ֶβɼ�<select id='channelid' name='channelid' onchange=""if (this.value!=0){location.href='?channelid='+this.value;}"">"
			.Write " <option value='0'>---��ѡ��ģ��---</option>"
	
			If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
			Dim ModelXML,Node
			Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
			For Each Node In ModelXML.documentElement.SelectNodes("channel[@ks21=1][@ks6=1||@ks6=2||@ks6=5]")
			  If trim(ChannelID)=trim(Node.SelectSinglenode("@ks0").text) Then
			   .Write "<option value='" &Node.SelectSingleNode("@ks0").text &"' selected>" & Node.SelectSingleNode("@ks1").text & "</option>"

			  Else
			   .Write "<option value='" &Node.SelectSingleNode("@ks0").text &"'>" & Node.SelectSingleNode("@ks1").text & "</option>"
			  End If
			next
			.Write "</select></div>"
		
		.Write "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
		.Write "<form action='Collect_Field.asp?action=SetCollect&channelid=" & ChannelID&"&page="&CurrentPage &"'' name='form1' method='post'>"
		.Write "        <tr class='sort'>"
		.Write "         <td width='80' align='center'>���òɼ�</td>"
		.Write "          <td width='100' align='center'>�ֶ�����</td>"
		.Write "          <td align='center'>�ֶα���</td>"		
		.Write "          <td align='center'>����ģ��</td>"
		.Write "          <td align='center'>�ֶ�����</td>"
		.Write "          <td align='center'>�Ƿ����òɼ�</td>"
		.Write "          <td align='center'>����λ��</td>"
		.Write "        </tr>"
			 Set FieldRS = Server.CreateObject("ADODB.RecordSet")
				   FieldSql = "SELECT * FROM KS_Field Where ChannelID=" & ChannelID & " order by orderid asc"
				   FieldRS.Open FieldSql, conn, 1, 1
				 If FieldRS.EOF And FieldRS.BOF Then
				 Else
					totalPut = FieldRS.RecordCount
		
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
									FieldRS.Move (CurrentPage - 1) * MaxPerPage
									Call showContent
								Else
									CurrentPage = 1
									Call showContent
								End If
							End If
			End If
		 .Write " <tr>"
		 .Write "   <td colspan='3'>&nbsp;&nbsp;<input type='submit' class='button' value='���������ֶβɼ�����'> </td></form>"
		 .Write "   <td height='35' colspan='4' align='right'>"
		 Call KSCLS.ShowPage(totalPut, MaxPerPage, "Collect_Field.asp", True, "��",CurrentPage, "ChannelID=" & ChannelID)
		.Write "    </td>"
		.Write " </tr>"
		.Write "</table>"
		.Write "</div>"
		End With
		End Sub
		Sub showContent()
		Dim CollectTF,ShowType
		With Response
		Do While Not FieldRS.EOF
		  Dim RS:Set RS=KS.ConnItem.Execute("Select * From KS_FieldItem where fieldid=" & FieldRS("FieldID"))
		  IF RS.Eof Then
		   CollectTF=false
		   ShowType=0
		  Else
		   ShowType=RS("ShowType")
		   CollectTF=true
		  End If
		  
		  RS.Close:Set RS=Nothing
		 .Write "<tr>"
		 .Write "<td class='splittd' align='center'>&nbsp;&nbsp;"
		 If CollectTF=True Then
		 .Write "<input type='checkbox' name='CField"& FieldRS("FieldID")&"' value='1' checked>"
		 Else
		 .Write "<input type='checkbox' name='CField"& FieldRS("FieldID")&"' value='1'>"
		 End iF
		 .Write "<input type='hidden' name='FieldID' value='" & FieldRS("FieldID") & "'>" & FieldRS("FieldID") & "</td>"
		 .Write "  <td class='splittd'><img src='../Images/Field.gif' align='absmiddle'><span  style='cursor:default;'>" & FieldRS("FieldName") & "</td>"
		 .Write "   <td align='center' class='splittd'>" & FieldRS("Title") & " </td>"
		 .Write "   <td align='center' class='splittd'><font color=red>"
		 If ChannelID=101 Then
		 .Write "��Աϵͳ"
		 Else
		  .Write KS.C_S(ChannelID,1) 
		 End If
		  .Write "</font>"
		 .Write "</td>"
		 .Write "   <td align='center' class='splittd'>"
				 Select Case FieldRS("FieldType")
				  Case 1:.Write "�����ı�(text)"
				  Case 2:.Write "�ı�(��֧��HTML)"
				  Case 10:.Write "�����ı�(֧��HTML)"
				  Case 3:.Write "�����б�(select)"
				  Case 4:.Write "����(text)"
				  Case 5:.Write "����(text)"
				  Case 6:.Write "��ѡ��(radio)"
				  Case 7:.Write "��ѡ��(checkbox)"
				  Case 8:.Write "��������(text)"
				  Case 9:.Write "�ļ�(text)"
				 End Select
		  If Left(Lcase(FieldRS("FieldName")),3)<>"ks_" Then .Write "<font color=#cccccc>[ϵͳ]</font>"
		 .Write "</td>"
		 .Write "   <td align='center' class='splittd'>&nbsp;" 
		 If CollectTF=false Then
		  .Write "δ����"
		 Else
		  .Write "<font color=red>������</font>"
		 End If
		 .Write " </td>"
		 
		 '========================�����б�ɼ�����=================
		 .Write "   <td align='center' class='splittd'>" 
		 .Write "<input type='radio' value='1' name='ShowType"& FieldRS("FieldID")&"'"
		 If ShowType=1 Then .Write " Checked"
		 .Write ">�б�ҳ"
		 .Write "<input type='radio' value='0' name='ShowType"& FieldRS("FieldID")&"'"
		 If ShowType=0 Then .Write " Checked"
		 .Write ">����ҳ"
		 .Write " </td>"
		 '================================================================
		 
		 .Write " </tr>"
								I = I + 1
								If I >= MaxPerPage Then Exit Do
							   FieldRS.MoveNext
							   Loop
								FieldRS.Close
						 
         End With
		 End Sub
		 
		 
		 
		 Sub FieldSetCollect()
			  Dim FieldID:FieldID=KS.G("FieldID")
			  Dim I,FieldIDArr,AllowStr
			  FieldIDArr=Split(FieldID,",")
			  AllowStr=KS.G("CField")
			  For I=0 To Ubound(FieldIDArr)
			   If KS.G("CField" & trim(FieldIDArr(i)))="1" Then
			      Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
				   RS.Open "Select * From KS_FieldItem Where FieldID=" & FieldIDArr(i),KS.ConnItem,1,3
				   If RS.Eof Then 
					RS.AddNew
				   End If
					RS("FieldID")=FieldIDArr(i)
					'==============�б�ɼ�����================================
					RS("ShowType")=KS.ChkClng(KS.G("ShowType"&trim(FieldIDArr(i))))
					'==========================================================
					RS("ChannelID")=KS.ChkClng(KS.G("ChannelID"))
					RS("FieldName")=Conn.Execute("Select FieldName From KS_Field Where FieldID=" & FieldIDArr(i))(0)
					RS("FieldTitle")=Conn.Execute("Select Title From KS_Field Where FieldID=" & FieldIDArr(i))(0)
					RS("OrderID")=Conn.Execute("Select OrderID From KS_Field Where FieldID=" & FieldIDArr(i))(0)
				   RS.Update
				   KS.ConnItem.Execute("Update KS_FieldRules Set ShowType=" & rs("ShowType") & ",ChannelID=" & RS("channelid") & ",FieldName='" & RS("FieldName") & "' Where FieldID=" & FieldIDArr(i))
				   RS.Close:Set RS=Nothing
			   Else
			     KS.ConnItem.Execute("Delete From KS_FieldItem Where FieldID=" & FieldIDArr(i))
			   End If
			  
			  Next
			  Response.Write "<script>alert('���������ֶγɹ���');location.href='?ChannelID=" & ChannelID & "&Page=" & Page&"';</script>"
		 End Sub
End Class
%> 
