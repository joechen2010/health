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
Set KSCls = New Special_Main
KSCls.Kesion()
Set KSCls = Nothing

Class Special_Main
        Private KS,KSCls
		Private SpecialID, i, totalPut, CurrentPage, SqlStr, SpecialRS
		Private FolderSql, FolderRS, ArticleTid, SpecialName
		Private CreateDate, TempStr,IcoUrl
		Private ChannelID,ClassID
		Private KeyWord, SearchType, StartDate, EndDate
		  '������������
		Dim SearchParam
		  
		Private MaxPerPage
		Private Sub Class_Initialize()
		  MaxPerPage = 20
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub
		Public Sub Kesion()
		
		If Not KS.ReturnPowerResult(0, "M010003") Then                  'Ȩ�޼��
				Call KS.ReturnErr(1, "")   
				Response.End()
		End iF

		KeyWord     = KS.G("KeyWord")
		SearchType  = KS.G("SearchType")
		StartDate   = KS.G("StartDate")
		EndDate     = KS.G("EndDate")
		SearchParam = "KeyWord=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate
		ClassID    = KS.G("ClassID"):If ClassID = "" Then ClassID = "0"
		SpecialID   = KS.G("SpecialID"):If SpecialID = "" Then SpecialID = "0"
		
		  
		 Select Case KS.G("Action")
		 Case "SpecialList" GetTop : Call SpecialMainList()
		 Case "Add","Edit"  GetTop : Call SpecialAddOrEdit()
		 Case "AddSave" GetTop : Call SpecialAddSave()
		 Case "EditSave"  GetTop : Call SpecialEditSave()
		 Case "SpecialDel" GetTop : Call SpecialDel()
		 Case "SpecialInfoDel" GetTop : Call SpecialInfoDel()
		 Case "AddClass","EditClass" GetTop : Call SpecialClassAdd()
		 Case "DoClassSave" GetTop : Call DoClassSave()
		 Case "DelClass" GetTop : Call DelSpecialClass()
		 Case "ShowInfo" GetTop :  Call ShowInfo()
		 Case "SpecialClassList" GetTop : Call SpecialClassList()
		 Case "Select" Call SpecialSelect()
		 Case ELSE  GetTop : Call SpecialMainList()
		 End Select
		End Sub
		
		Sub SpecialSelect()
		  ChannelID=KS.S("channelid")
     %>
	<html>
	<META HTTP-EQUIV="pragma" CONTENT="no-cache"> 
	<META HTTP-EQUIV="Cache-Control" CONTENT="no-cache, must-revalidate"> 
	<META HTTP-EQUIV="expires" CONTENT="Wed, 26 Feb 1997 08:21:57 GMT">
	<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>
	<title>ѡ��ר��</title>
    <script language="javascript" src="../KS_Inc/common.js"></script>
    <script language="javascript" src="../KS_Inc/jquery.js"></script>
	 <style>
	  body{margin:0px;padding:0px;font-size:12px;COLOR: #454545; text-decoration: none;}
	  td{font0-size:12px;}
	  a{text-decoration: none;COLOR: #454545; }
	 </style>
	    <script language="javascript">
		function SelectFolder(TypeID){
		   $("#sub"+TypeID).toggle();
		   $("#sub"+TypeID).html("<img src='images/loading.gif'>");
		   $.get("../plus/ajaxs.asp",{action:"SpecialSubList",classid:TypeID},function(d){
		    $("#sub"+TypeID).html(unescape(d));
		   });
		}
		
      function set(specialid,specialname)
	  { 
	    top.frames["MainFrame"].UpdateSpecial(specialid+'@@@'+specialname);
		top.frames["MainFrame"].closeWindow();
	  }
    </script>
	  <body bgcolor="E9F6FE">
	    <table border="0" cellpadding="0" cellspacing="0" width="100%">
		 <tr>
		  <td>
	   <%
	  With KS 
		 Dim Node,K,SQL,ID,RS,Xml
		 Set RS=Conn.Execute("select ClassID,ClassName from KS_SpecialClass Order By OrderID ASC")
		 If Not RS.Eof Then
		   Set Xml=.RsToXml(RS,"row","xmlroot")
		 End If
		   RS.Close
		   Set RS=Nothing
		 If IsOBject(Xml) Then
		    For Each Node In Xml.DocumentElement.SelectNodes("row")
				ID=Node.SelectSingleNode("@classid").text
		          .echo "<table style=""margin:14px"" width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
				  .echo "<tr>" & vbcrlf
				  .echo " <td><img src='images/folder/folder.gif' align='absmiddle'><span onClick='SelectFolder(" & ID &");return false;'><a href='#'><strong>" & Node.SelectSingleNode("@classname").text & "</strong></a></span>"
				  .echo "</td>"&vbcrlf
				  .echo "</tr>" & vbcrlf
				  .echo "<tr>" & vbcrlf
                  .echo " <td style=""padding-left:20px"" ID=""sub"& ID &""" style=""display:none"">" & vbcrlf
                  .echo " </td>" & vbcrlf
                  .echo " </tr>" & vbcrlf
	  			  .echo "</table>"
			Next
	   Else
		     .echo "�������ר��!"
	   End If
	   End With
	   		%>
		  </td>
		  </tr>
		 </table>
		</body>
		</html>
		<%
		End Sub
		
		Sub SpecialClassList()
			With KS
			.echo "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
			.echo "        <tr>"
			.echo "          <td class=""sort"" width='65' align='center'>����ID</td>"
			.echo "          <td class='sort' align='center'>ר���������</td>"
			.echo "          <td width='19%' class='sort' align='center'>ר����</td>"
			.echo "          <td width='10%' align='center' class='sort'>�����</td>"
			.echo "          <td width='35%' align='center' class='sort'>�������</td>"
			.echo "  </tr>"
			MaxPerPage=15
			  Dim RS:Set RS = Server.CreateObject("ADODB.RecordSet")
			   RS.Open "SELECT ClassID,ClassName,OrderID FROM [KS_SpecialClass] order by OrderID", conn, 1, 1
				If Not RS.EOF Then
						totalPut = RS.RecordCount
						If CurrentPage < 1 Then CurrentPage = 1
			
						If (CurrentPage - 1) * MaxPerPage > totalPut Then
							If (totalPut Mod MaxPerPage) = 0 Then
										CurrentPage = totalPut \ MaxPerPage
							Else
										CurrentPage = totalPut \ MaxPerPage + 1
							End If
						End If
			
						If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
							RS.Move (CurrentPage - 1) * MaxPerPage
						Else
							CurrentPage = 1
						End If
						Dim SQL:SQL=RS.GetRows(MaxPerPage)
						Call showSpecialClass(SQL)
				Else
				  .echo "<tr><td class='splittd' align='center' height='25' colspan=5>����û�����ר����࣬�����!</td></tr>"
				End If
			.echo "</table>"
			.echo ("<table border='0' width='100%' cellspacing='0' cellpadding='0' align='center'>")
	        .echo ("<tr><td width='180'> </div>")
	        .echo ("</td>")
	        .echo ("<td></td>")
	        .echo ("</form><td align='right'>")
	         Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
	        .echo ("</td></tr></table>")
			.echo "</div>"
          End With
		End Sub
		
		Sub GetTop
			CurrentPage = KS.ChkClng(Request("page"))
			If CurrentPage=0 Then  CurrentPage = 1
		  With KS
			.echo "<html>"
			.echo "<head>"
			.echo "<meta http-equiv=""Content-Type"" content=""text/html; chaRSet=gb2312"">"
			.echo "<title>ר������</title>"
			.echo "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.echo "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>"
			.echo "<script language=""JavaScript"" src=""../KS_Inc/Jquery.js""></script>"
			.echo "<script src=""../ks_inc/kesion.box.js""></script>"
			.echo "<script language=""JavaScript"">" & vbCrLf
			.echo "var Page='" & CurrentPage & "';        //��ǰҳ��" & vbCrLf
			.echo "var ClassID='" & ClassID & "';       //Ƶ��ID" & vbCrLf
			.echo "var SpecialID=" & SpecialID & ";       //ר��ID" & vbCrLf
			.echo "var KeyWord='" & KeyWord & "';         //�����ؼ���" & vbCrLf
			.echo "var SearchParam='" & SearchParam & "'; //������������" & vbCrLf
			.echo "</script>" & vbCrLf
			%>
			<script language="javascript">
			$(document).ready(function(){
				 // $(parent.frames["BottomFrame"].document).find("#Button1").attr("disabled",true);
				 // $(parent.frames["BottomFrame"].document).find("#Button2").attr("disabled",true);
				});
				function CreateHtml(SpecialID)
				{   if (SpecialID=='') SpecialID=get_Ids(document.myform);
					if (SpecialID!='')
					{
					   PopupCenterIframe('����ר��','include/RefreshspecialSave.asp?Types=Special&id='+SpecialID+'&RefreshFlag=ID',530,110,'no')
					}
					else 
					alert('��ѡ��Ҫ������ר��!');
				}
				
				function ChangeUp()
				{
				 location.href='KS.Special.asp';
				 $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?Opstr='+escape("����ר��")+'&ButtonSymbol=Disabled&ClassID='+ClassID;
				}
				
		
				function View(SpecialName,SpecialID)
				{ if (SpecialID=='') SpecialID=get_Ids(document.myform);
					 if (SpecialID!=''){
						 if (SpecialID.indexOf(',')==-1){
						 PopupCenterIframe('�鿴ר��<font color=red>['+SpecialName+']</font>�µ��ĵ�','?Action=ShowInfo&SpecialID='+SpecialID+'',750,430,'auto')
							 } else alert('һ��ֻ�ܹ��༭һ��ר��');
						}
					else{
					alert('��ѡ��Ҫ�༭��ר��');
					}
				}
				function Delete(SpecialID)
				{  
				  if (SpecialID=='') SpecialID=get_Ids(document.myform);
						if (SpecialID!='')
						{ 
						if (confirm('ȷ��ɾ��ѡ�е�ר����?'))location="KS.Special.asp?Action=SpecialDel&Page="+Page+"&"+SearchParam+"&SpecialID="+SpecialID+'&ClassID='+ClassID;
						}
						else alert('��ѡ��Ҫɾ����ר��!');
					
				}
				function SpecialInfoDel(ID)
				{
				 if (confirm('ȷ����ѡ�е��ĵ���ר�����Ƴ���?')) 
				 {
				   $("input[type=checkbox][value="+ID+"]").attr("checked",true);
				   $("#myform").submit();
				 }
				}
			  function showInfo(channelid,id)
			  {
				 window.open('../item/show.asp?m='+channelid+'&d='+id);
			   }
			function CreateSpecialClass()
			{
			 location.href='KS.Special.asp?Action=AddClass';
			 $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?Opstr='+escape("ר����� >> <font color=red>���ר�����</font>")+'&ButtonSymbol=Go';
			}
			function EditSpecialClass(classid)
			{
			 location.href='KS.Special.asp?Action=EditClass&ClassID='+classid+'&Page='+Page;
			 $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?Opstr='+escape("ר����� >> <font color=red>�޸�ר�����</font>")+'&ButtonSymbol=GoSave';
			}
			function AddSpecial(ClassID)
			{
			 location.href='KS.Special.asp?Action=Add&ClassID='+ClassID;
			 $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?Opstr='+escape("ר����� >> <font color=red>���ר��</font>")+'&ButtonSymbol=Go';
			}
			function Edit(SpecialID)
			{  
			 if (SpecialID=='') SpecialID=get_Ids(document.myform);
			 if (SpecialID!=''){
				 if (SpecialID.indexOf(',')==-1){
				   location.href='KS.Special.asp?Action=Edit&SpecialID='+SpecialID;
				   $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?Opstr='+escape("ר����� >> <font color=red>�༭ר��</font>")+'&ButtonSymbol=GoSave&ClassID='+ClassID;
					 } else alert('һ��ֻ�ܹ��༭һ��ר��');
				}
			else{
			alert('��ѡ��Ҫ�༭��ר��');
			}
			
			}
			function ClassToggle(f)
			{
			  setCookie("SpecialclassExtStatus",f)
			  $('#classNav').toggle('slow');
			  $('#classOpen').toggle('show');
			}
			</script>
			<body>
			<%
		  If KS.G("Action")<>"ShowInfo" Then
		 	.echo "<ul id='menu_top'>"
			.echo "<li class='parent' onclick=""location.href='?action=SpecialClassList'"""
			.echo "><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/addjs.gif' border='0' align='absmiddle'>�������</span></li>"
			.echo "<li class='parent' onclick='javascript:CreateSpecialClass();'"
			.echo "><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/addfolder.gif' border='0' align='absmiddle'>��ӷ���</span></li>"

			
			.echo "<li class='parent' onclick='javascript:AddSpecial(" & KS.G("ClassID") & ");'"
			If SpecialID <> "0" Then .echo (" Disabled=true")
			.echo "><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/addfolder.gif' border='0' align='absmiddle'>���ר��</span></li>"
			.echo "<li class='parent' onclick=""javascript:Edit('');"""
			If SpecialID<>"0" or ClassID="0" Then .echo " Disabled"
			.echo "><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/as.gif' border='0' align='absmiddle'>�༭ר��</span></li>"
			.echo "<li class='parent' onClick=""Delete('')"""
			If SpecialID <> "0" or ClassID="0" Then .echo (" Disabled=true")
			.echo "><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/del.gif' border='0' align='absmiddle'>ɾ��ר��</span></li>"
			.echo "<li class='parent' onClick=""parent.frames['LeftFrame'].initializeSearch('ר������',0,'Special');""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/s.gif' border='0' align='absmiddle'>��������</span></li>"
			.echo "<li class='parent' onClick=""ChangeUp();"""
			If ClassID="0" Then .echo " Disabled"
			.echo "><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/back.gif' border='0' align='absmiddle'>����һ��</span></li>"
			.echo "</ul>"
		   End If
          End With
		End Sub
		
		Sub showSpecialClass(SQL)
		  Dim K
		  With KS
		  For K=0 To Ubound(SQL,2)
		     .echo "<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">"
		     .echo "<td class='splittd' height='30' align='center'>" & SQL(0,K) & "</td>"
			 .echo "<td class='splittd'><img src='images/folder/folder.gif' align='absmiddle'><a href='?Action=SpecialList&ClassID=" & SQL(0,K) & "'>" & SQL(1,K) & "</a></td>"
			 .echo "<td class='splittd' align='center'>" & conn.execute("select count(*) from ks_special where classid=" & SQL(0,K))(0) & "</td>"
			 .echo "<td class='splittd' align='center'>" & SQL(2,K) &"</td>"
			 .echo "<td class='splittd' align='center'><a href='javascript:AddSpecial(" & SQL(0,K) & ");'>���ר��</a> | <a href='?Action=SpecialList&ClassID=" & SQL(0,K) & "'>�鿴�÷����µ�ר��</a> | <a href='javascript:EditSpecialClass(" & SQL(0,K) & ");'>�޸�</a> | <a href='?Action=DelClass&ClassID=" &SQL(0,K) & "' onclick=""return(confirm('ɾ�����ཫͬʱɾ���÷����µ�����ר�⣬ȷ��ɾ����'))"">ɾ��</a></td>"
			 .echo "</tr>"
		  Next
		  End With
		End Sub
		
		
		
		Sub SpecialMainList()
			With KS
			.echo "</head>"
		 If KeyWord = "" Then
			GetChannelList()
		 Else
			.echo "<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" class=""sortbutton"">"
			.echo "  <tr>"
			.echo "    <td height=""23"" align=""left"">"
					   .echo ("<img src='Images/home.gif' align='absmiddle'><span style='cursor:pointer' onclick=""SendFrameInfo('KS.Special.asp','Special_Left.asp','KS.Split.asp?ButtonSymbol=Disabled&OpStr=ר����� >> <font color=red>������ҳ</font>')"">ר����ҳ</span>")
				   .echo (">>> �������: ")
					 If StartDate <> "" And EndDate <> "" Then
						.echo ("ר����������� <font color=red>" & StartDate & "</font> �� <font color=red> " & EndDate & "</font>&nbsp;&nbsp;&nbsp;&nbsp;")
					 End If
					Select Case SearchType
					 Case 0
					  .echo ("���ƺ��� <font color=red>" & KeyWord & "</font> ��ר��")
					 Case 1
					  .echo ("��Ҫ˵���к��� <font color=red>" & KeyWord & "</font> ��ר��")
					 End Select
		 End If
				  
			.echo "    </td>"
			.echo "  </tr>"
			.echo "</table>"
			
			 '============������ʾ,�����书��=======================================
			 Dim ExtStatus,CloseDisplayStr,ShowDisplayStr,classExtStatus
			 classExtStatus=request.cookies("SpecialclassExtStatus")
			 if classExtStatus="" Then classExtStatus=1
			 If classExtStatus=1 Then 
			  ExtStatus=2 :CloseDisplayStr="display:none;":ShowDisplayStr=""
			 Else 
			  ExtStatus=1 :CloseDisplayStr="":ShowDisplayStr="display:none;"
			 End If

			Dim RS,ClassXML,Node
			Set RS=Conn.Execute("Select ClassID,ClassName From KS_SpecialClass Order by OrderID")
			If Not RS.Eof Then Set ClassXML=KS.RsToXml(RS,"row","classxml")
			RS.Close:Set RS=Nothing
			If IsObject(ClassXML) Then
			.echo "<div id='classOpen' onclick=""ClassToggle("& ExtStatus& ")"" style='" & CloseDisplayStr &"cursor:pointer;text-align:center;position:absolute; z-index:2; left: 0px; top: 38px;' ><img src='images/kszk.gif' align='absmiddle'></div>"
		    .echo "<div id='classNav' style='" & ShowDisplayStr &"position:relative;height:auto;_height:30px;top:4px;line-height:30px;margin:8px 1px;border:1px solid #DEEFFA;background:#F7FBFE'>"
		    .echo "<div style='padding-top:2px;cursor:pointer;text-align:center;position:absolute; z-index:1; right: 0px; top: 2px;'  onclick=""ClassToggle(" & ExtStatus &")""> <img src='images/close.gif' align='absmiddle'></div>"
			 For Each Node In ClassXML.DocumentElement.SelectNodes("row")
			   .echo "<li style='margin:5px;float:left;width:100px'><img src='images/folder/folderopen.gif' align='absmiddle'><a href='?classid=" & Node.SelectSingleNode("@classid").text & "' title='" & Node.SelectSingleNode("@classid").text & "'>" & KS.Gottopic(Node.SelectSingleNode("@classname").text,10) & "</a></li>"
			 Next
			 .echo "</div>"
			End If
			 '=============================================================

			
			
		    .echo ("<table width=""100%"" align='center' border=""0"" cellpadding=""0"" cellspacing=""0"">")
			.echo ("<form name='myform' id='myform' action='KS.Special.asp' method='post'>")
		    .echo ("<tr class='sort'>")
			.echo ("<td>ѡ��</td><td>ר������</td><td>����</td><td>���ʱ��</td><td>�������</td>")
			.echo ("</tr>")
	
	
	  If KeyWord <> "" Then
		  Dim Param:Param = " Where 1=1"
		  Select Case SearchType
			Case 0
			Param = Param & " And SpecialName like '%" & KeyWord & "%'"
			Case 1
			Param = Param & " And SpecialNote like '%" & KeyWord & "%'"
		  End Select
			If StartDate <> "" And EndDate <> "" Then
				Param = Param & " And (SpecialAddDate>=#" & StartDate & "# And SpecialAddDate<=#" & DateAdd("d", 1, EndDate) & "#)"
		   End If
		  Param = Param & " Order BY SpecialAddDate desc"
		  SqlStr = "Select SpecialID,a.ClassID,b.ClassName,SpecialName,Creater,SpecialAddDate,SpecialNote from KS_Special a Inner Join KS_SpecialClass B on a.classid=b.classid " & Param
	  Else
		  If ClassID<>"0" Then
		   SqlStr = "Select SpecialID,a.ClassID,b.ClassName,SpecialName,Creater,SpecialAddDate,SpecialNote from KS_Special a Inner Join KS_SpecialClass B on a.classid=b.classid Where a.ClassID=" & ClassID & " Order BY SpecialAddDate desc"
		  Else
		   SqlStr = "Select SpecialID,a.ClassID,b.ClassName,SpecialName,Creater,SpecialAddDate,SpecialNote from KS_Special a Inner Join KS_SpecialClass B on a.classid=b.classid Order BY SpecialAddDate desc"
		  End If
	  End If
	 Set SpecialRS = Server.CreateObject("AdoDb.RecordSet")
	 SpecialRS.Open SqlStr, Conn, 1, 1
	 If SpecialRS.EOF Then
	    .echo "<tr><td class='splittd' onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" colspan='5' align='center'>�Ҳ���ר��!</td></tr>"
	 Else
				totalPut = SpecialRS.RecordCount
						If CurrentPage < 1 Then	CurrentPage = 1
	
						If (CurrentPage - 1) * MaxPerPage > totalPut Then
							If (totalPut Mod MaxPerPage) = 0 Then
								CurrentPage = totalPut \ MaxPerPage
							Else
								CurrentPage = totalPut \ MaxPerPage + 1
							End If
						End If
	
						If CurrentPage > 1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
								SpecialRS.Move (CurrentPage - 1) * MaxPerPage
							Else
								CurrentPage = 1
								
						End If
						
						Dim XML:Set XML=KS.ArrayToXml(SpecialRS.GetRows(MaxPerPage),SpecialRS,"row","xmlroot")
						showSpecialList XML
						Set XML=Nothing
						
		End If
			 .echo " <tr>"
			 .echo " <td colspan='3'><div style='margin:5px'><b>ѡ��</b><a href='javascript:void(0)' onclick='Select(0)'>ȫѡ</a> -  <a href='javascript:void(0)' onclick='Select(1)'>��ѡ</a> - <a href='javascript:void(0)' onclick='Select(2)'>��ѡ</a> <input type='button' class='button' value='ɾ ��' onclick=""Delete('');""> &nbsp;&nbsp;<input type='button' class='button' value='�� ��' onclick=""CreateHtml('');""></td></form>"
			 .echo "   <td align=""right"" colspan=5>"
			 
			 Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
			.echo "   </td>"
			.echo "  </tr>"
	.echo "</table>"
	.echo "</body>"
	.echo "</html>"
	 End With
	End Sub
	
	 Sub showSpecialList(XML)
	  Dim Node,SpecialID,SpecialName
	  If Not IsObject(XML) Then Exit Sub
	  With KS
			For Each Node In XML.DocumentElement.SelectNodes("row")
			    SpecialID=Node.SelectSingleNode("@specialid").text
				SpecialName=Node.SelectSingleNode("@specialname").text
				  .echo ("<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" &SpecialID & "' onclick=""chk_iddiv('" & SpecialID & "')"">")
				  .echo ("<td class='splittd' align=center><input name='id'  onclick=""chk_iddiv('" & SpecialID & "')"" type='checkbox' id='c"& SpecialID & "' value='" & SpecialID & "'></td>")
				  .echo ("<td class='splittd' TITLE='�� ��:" & SpecialName & "'>")
				  .echo ("<span onmousedown=""mousedown(this);"" style=""POSITION:relative;"" SpecialID=""" &SpecialID & """ SpecialName=""" & SpecialName & """>")
				  .echo ("<img src=""Images/Folder/Special.gif""> ")
				  .echo ("<a href=""javascript:View('" & SpecialName & "','" & SpecialID & "')"">" & SpecialName & "</a>")
				  .echo ("</td>")
				  .echo ("<td class='splittd' align='center'>" & Node.SelectSingleNode("@classname").text & "</td>")
				  .echo ("<td class='splittd' align='center'>" & Node.SelectSingleNode("@specialadddate").text & "</td>")
				  .echo ("<td class='splittd' align='center'><a href='javascript:Edit(""" & SpecialID & """);'>�༭</a> | <a href='javascript:Delete(" & SpecialID & ")'>ɾ��</a> | <a href='javascript:CreateHtml(""" & SpecialID & """);'>����</a> | <a href=""javascript:View('" & SpecialName & "','" & SpecialID & "')"">�鿴</a> | <a href='../special.asp?id=" &SpecialID & "' target='_blank'>���</a></td>")
			     .echo " </tr>"
			Next
					   
			End With
		End Sub
		
		'��ʾר���µ���Ϣ
		Sub ShowInfo()
		    MaxPerPage=10
		 	With KS
			 .echo ("<table width=""100%"" align='center' border=""0"" cellpadding=""0"" cellspacing=""0"">")
			 .echo ("<form name='myform' id='myform' action='KS.Special.asp' method='post'>")
			 .echo ("<input type='hidden' name='action' value='SpecialInfoDel'>")
		     .echo ("<tr class='sort'>")
			 .echo ("<td>ѡ��</td><td>�ĵ�����</td><td>����</td><td>���ʱ��</td><td>�������</td>")
			 .echo ("</tr>")

			 Dim SQLStr
			 Dim RS:Set RS=Server.CreateoBject("ADODB.RECORDSET")
			 SQLStr="Select R.ID,I.ChannelID,I.InfoID,I.Title,I.Tid,I.AddDate From KS_ItemInfo I Inner Join KS_SpecialR R On I.InfoID=R.InfoID Where R.SpecialID=" & SpecialID & " and i.channelid=r.channelid Order by i.id Desc"
			 RS.Open SQLStr,Conn,1,1
			 If RS.EOF Then
			  .echo ("<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"">")
			  .echo "<td class='splittd' colspan='6' align='center'>��ר����û������ĵ�!</td>"
			  .echo "</tr>"
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
		
							If CurrentPage > 1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
								Else
									CurrentPage = 1
							End If
							
							Dim XML,Node,InfoID,RID
							Set XML=KS.ArrayToXml(RS.GetRows(MaxPerPage),RS,"row","xmlroot")
							If IsObject(XML) Then
								For Each Node In XML.DocumentElement.SelectNodes("row")
								      RID=Node.SelectSingleNode("@id").text
									  InfoID=Node.SelectSingleNode("@infoid").text
									  .echo ("<tr onmouseout=""this.className='list'"" onmouseover=""this.className='listmouseover'"" id='u" &RID & "' onclick=""chk_iddiv('" & RID & "')"">")
									  .echo ("<td class='splittd' align=center><input name='id'  onclick=""chk_iddiv('" & RID & "')"" type='checkbox' id='c"& RID & "' value='" & RID & "'></td>")
									  .echo ("<td class='splittd' TITLE='�� ��:" & Node.SelectSingleNode("@title").text & "'>")
									  .echo ("<a href='javascript:void(0)' onclick=""showInfo('" & Node.SelectSingleNode("@channelid").text & "','" & InfoID & "')"">" & KS.Gottopic(Node.SelectSingleNode("@title").text,30) & "</a>")
									  .echo ("</td>")
									  .echo ("<td class='splittd' align='center'>" & KS.C_C(Node.SelectSingleNode("@tid").text,1) & "</td>")
									  .echo ("<td class='splittd' align='center'>" & Node.SelectSingleNode("@adddate").text & "</td>")
									  .echo ("<td class='splittd' align='center'> <a href=""javascript:SpecialInfoDel('" & RID & "')"">ɾ��</a> | <a href=""javascript:showInfo(" & Node.SelectSingleNode("@channelid").text & "," & InfoID & ")"">�鿴</a></td>")
									 .echo " </tr>"
								Next
							End If
							Set XML=Nothing
							
			End If
			RS.Close:Set RS=Nothing
			 .echo " <tr>"
			 .echo " <td colspan='3'><div style='margin:5px'><b>ѡ��</b><a href='javascript:void(0)' onclick='Select(0)'>ȫѡ</a> -  <a href='javascript:void(0)' onclick='Select(1)'>��ѡ</a> - <a href='javascript:void(0)' onclick='Select(2)'>��ѡ</a> <input type='submit' class='button' value='ɾ ��' onclick=""return(confirm('ȷ���Ƴ�ѡ�е��ĵ���?'))""> </td></form>"
			 .echo "   <td align=""right"" colspan=5>"
				Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,true,true)
			.echo "   </td>"
			.echo "  </tr>"
		  End With
		End Sub
		
		'���ר�����
		Sub SpecialClassAdd()
		   Dim ClassID,Action,ClassName,ClassEName,TemplateID,FsoIndex,AddDate,Descript,TopTitle,PhotoUrl,OrderID
		   Dim CurrPath:CurrPath = KS.GetUpFilesDir()
			If KS.G("Action")="EditClass" Then
			  ClassID=KS.G("ClassID")
			  TopTitle="�༭"
			  Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
			  RSObj.Open "Select * From KS_SpecialClass Where ClassID=" & ClassID,Conn,1,1
			  If Not RSObj.Eof Then
				ClassName   = RSObj("ClassName")
				ClassEName  = RSObj("ClassEName")
				TemplateID    = RSObj("TemplateID")
				FsoIndex=RSObj("FsoIndex")
				AddDate       = RSObj("AddDate")
				Descript   = RSObj("Descript")
				OrderID = RSOBj("Orderid")
			  End If
			Else
			  TopTitle="���":AddDate=Now:FsoIndex="Index.html"
			  OrderID=KS.ChkClng(conn.execute("select max(OrderID) from ks_specialclass")(0))+1
			End If
			With KS
			.echo "<html>" & vbCrLf
			.echo "<head>" & vbCrLf
			.echo "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
			.echo "<link href='Include/admin_style.css' rel='stylesheet'>" & vbCrLf
		    .echo "<script language='JavaScript' src='../KS_Inc/common.js'></script>"
	      	.echo "<script language='JavaScript' src='../KS_Inc/jquery.js'></script>"
			.echo "<title>ר��������</title>" & vbCrLf
			.echo "</head>" & vbCrLf
			.echo "<body  leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
			.echo "<div class='topdashed sort'>" & TopTitle &"ר�����</div>" & vbCrLf

			 
			.echo "  <form action='KS.Special.asp?Action=DoClassSave' name='SpecialForm' method='post'>" & vbCrLf
			.echo "  <input name='ClassID' type='hidden' id='ClassID' value='" & ClassID &"'>" & vbCrLf
			.echo "  <input name='Page' type='hidden' value='" & KS.G("Page") &"'>" & vbCrLf
			.echo "  <table width='99%' border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
			.echo "    <tr>" & vbCrLf
			.echo "      <td>" & vbCrLf
            .echo "        <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class='ctable'>" & vbCrLf
			.echo "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
			.echo "      <td width='179' height='35'class='clefttitle'> <div align='right'><strong>ר��������ƣ�</strong></div></td>" & vbCrLf
			.echo "      <td> <input name='ClassName' value='" & ClassName & "' type='text' size='30' class='textbox'>"
			.echo "              �ſ���˵������ </td>"
			.echo "    </tr>" & vbCrLf
			.echo "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.echo "      <td height='35' class='clefttitle'> <div align='right'><strong>ר�����Ŀ¼���ƣ�</strong></div></td>" & vbCrLf
			.echo "      <td>"
			.echo "<input"
				If KS.G("Action")="EditClass" Then .echo " Disabled"
			.echo " name='ClassEName' type='text' value='" & ClassEName & "'  size='30' class='textbox'>"
			.echo "        ֻ������ĸ�����ֻ��»��ߵ����  </td>"
			.echo "    </tr>"
			.echo "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.echo "      <td height='35' class='clefttitle'> <div align='right'><strong>ר���б�ҳģ�壺</strong></div></td>" & vbCrLf
			.echo "      <td><input type='text' size='30' name='TemplateID' id='TemplateID' value='" & TemplateID & "' class='textbox'>&nbsp;" & KSCls.Get_KS_T_C("$('#TemplateID')[0]") 
			.echo "      </td>" & vbCrLf
			.echo "    </tr>" & vbCrLf
				
			.echo "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
			.echo "      <td height='35' class='clefttitle'> <div align='right'><strong>����ר���б�ҳ���ļ���</strong></td>" & vbCrLf
			.echo "      <td><select name='FsoIndex' class='textbox'>"
			.echo "          <option value='index.html' selected>index.html</option>"
			.echo "          <option value='index.htm'>index.htm</option>"
			.echo "          <option value='index.shtm'>index.shtm</option>"
			.echo "          <option value='index.shtml'>index.shtml</option>"
			.echo "          <option value='default.html'>default.html</option>"
			.echo "          <option value='default.htm'>default.htm</option>"
			.echo "          <option value='default.shtml'>default.shtml</option>"
			.echo "          <option value='default.shtm'>default.shtm</option>"
			.echo "          <option value='index.asp'>index.asp</option>"
			.echo "         <option value='" & FsoIndex & "' selected>" & FsoIndex & "</option>"
			.echo "        </select></td>"
			.echo "    </tr>" & vbCrLf
			
			.echo "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.echo "      <td & vbCrLfheight='35' class='clefttitle'> <div align='right'><strong>���ʱ�䣺</strong></div></td>"
			.echo "      <td><input name='AddDate' type='text' value='" & AddDate & "' size='30' readonly class='textbox'>"
			.echo "      </td>" & vbCrLf
			.echo "    </tr>" & vbCrLf
			.echo "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.echo "      <td height='35' class='clefttitle'> <div align='right'><strong>��Ҫ˵����</strong></div></td>"
			.echo "      <td><textarea name='Descript' rows='8' style='width:80%;border-style: solid; border-width: 1'>" &Descript & "</textarea></td>"
			.echo "    </tr>"
			.echo "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.echo "      <td height='35' class='clefttitle'> <div align='right'><strong>������ţ�</strong></div></td>"
			.echo "      <td><input name='OrderID' size='5' value='" & OrderID & "' Class='textbox' style='text-align:center'> ����ԽС����Խǰ��</td>"
			.echo "    </tr>"
			.echo "  </table>"
			.echo "       </td>"
			.echo "    </tr>"
			.echo "    </table>"
			.echo "  </form>"
			.echo "</body>"
			.echo "</html>"
			.echo "<Script Language='javascript'>" & vbCrLf
			.echo "<!--" & vbCrLf
			.echo "function CheckForm()" & vbCrLf
			.echo "{ var form=document.SpecialForm;" & vbCrLf
			.echo "   if (form.ClassName.value=='')"
			.echo "    {" & vbCrLf
			.echo "     alert('������ר���������!');" & vbCrLf
			.echo "     form.ClassName.focus();" & vbCrLf
			.echo "     return false;" & vbCrLf
			.echo "    }" & vbCrLf
			.echo "    if (form.ClassEName.value=='')" & vbCrLf
			.echo "    {"
			.echo "     alert('������ר������Ӣ������!');" & vbCrLf
			.echo "     form.ClassEName.focus();" & vbCrLf
			.echo "    return false;" & vbCrLf
			.echo "    }"
			.echo "    if (form.TemplateID.value=='')" & vbCrLf
			.echo "    {"
			.echo "     alert('���ר���б�ҳģ��!');" & vbCrLf
			.echo "     form.TemplateID.focus();" & vbCrLf
			.echo "    return false;" & vbCrLf
			.echo "    }"

			.echo "    if (CheckEnglishStr(form.ClassEName,'Ŀ¼��Ӣ������')==false)" & vbCrLf
			.echo "     return false;" & vbCrLf
			.echo "    form.submit();" & vbCrLf
			.echo "    return true;" & vbCrLf
			.echo "}" & vbCrLf
			.echo "//-->" & vbCrLf
			.echo "</Script>"
		  End With
		End Sub
		
		Sub DoClassSave()
			Dim RS, Sql,ClassName, ClassEName, TemplateID, FsoIndex, AddDate, Descript,OrderID,ClassID
			ClassID    = KS.ChkClng(KS.G("ClassID"))
			ClassName  = KS.G("ClassName")
			ClassEName = KS.G("ClassEName")
			TemplateID = KS.G("TemplateID")
			FsoIndex   = KS.G("FsoIndex")
			AddDate    = KS.G("AddDate")
			Descript   = KS.G("Descript")
			OrderID    = KS.ChkClng(KS.G("OrderID"))
			With KS		 
				 If ClassName <> "" Then
					If Len(ClassName) >= 100 Then
						Call KS.AlertHistory("ר��������Ʋ��ܳ���50���ַ�!", -1):Exit Sub
					End If
				 Else
					Call KS.AlertHistory("������ר���������!", -1):Exit Sub
				 End If
				 If ClassEName <> "" and  ClassID=0 Then
					If Len(ClassEName) >= 50 Then
						Call KS.AlertHistory("ר�����Ӣ�����Ʋ��ܳ���50���ַ�!", -1):Exit Sub
					End If
					If Not Conn.Execute("Select ClassEName,ClassName from KS_SpecialClass where ClassID<>" & ClassID & " and ClassName='" & ClassName & "'").eof  Then Call KS.alertHistory("���ݿ����Ѵ��ڸ�ר���������!", -1)
					If Not Conn.Execute("Select ClassEName,ClassName from KS_SpecialClass where ClassID<>" &ClassID & " and ClassEName='" & ClassEName & "'").eof  Then Call KS.alertHistory("���ݿ����Ѵ��ڸ�ר�����Ӣ������!", -1)
				 ElseIf ClassID=0 Then
					Call KS.alertHistory("������ר�����Ӣ������!", -1)
					.End
				 End If
				 If ClassID=0 Then
				  Conn.Execute("Insert Into KS_SpecialClass(ClassName,ClassEname,Descript,FsoIndex,AddDate,TemplateID,OrderID) Values('" & ClassName & "','" & ClassEname & "','" & Descript & "','" & FsoIndex & "','" & AddDate & "','" & TemplateID & "'," & OrderID &")")
				  .echo ("<script>if (confirm('���ר�����ɹ�,���������?')==true){location.href='KS.Special.asp?action=AddClass';}else{location.href='KS.Special.asp?action=SpecialClassList';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?Opstr=" & server.URLEncode("���ݹ��� >> ����ר�����") & "&ButtonSymbol=Disabled&ClassID=" & ClassID & "';}</script>")     
				 Else
				  Conn.Execute("Update KS_SpecialClass Set ClassName='" & ClassName & "',Descript='" & Descript & "',FsoIndex='" & FsoIndex & "',AddDate='" & AddDate & "',TemplateID='" & TemplateID & "',OrderID=" & Orderid & " Where ClassID=" & ClassID)
				  .echo ("<script>alert('ר������޸ĳɹ�');location.href='KS.Special.asp?action=SpecialClassList&Page=" & KS.G("Page") &"';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?Opstr=" & server.URLEncode("���ݹ��� >> ����ר�����") & "&ButtonSymbol=Disabled';</script>")     
				 End If
			End With
		End Sub
		
		'ɾ��ר�����
		Sub DelSpecialClass()
		  Dim ClassID:ClassID=KS.ChkClng(KS.S("ClassID"))
		  Conn.Execute("Delete From KS_SpecialR Where SpecialID in(select specialid from ks_special where classid=" & ClassID & ")")
		  Conn.Execute("Delete From KS_Special Where ClassID=" & ClassID)
		  Conn.Execute("Delete From KS_SpecialClass Where ClassID=" & ClassID)
		  Response.Redirect Request.ServerVariables("HTTP_REFERER")
		End Sub
		
		'��ӻ�༭ר��
		Sub SpecialAddOrEdit()
		   Dim SpecialID,Action,SpecialName,SpecialEName,TemplateID,FsoSpecialIndex,AddDate,SpecialNote,TopTitle,PhotoUrl,ClassID
		   Dim CurrPath:CurrPath = KS.GetUpFilesDir()
			If KS.G("Action")="Edit" Then
			  SpecialID=KS.G("SpecialID")
			  Action="EditSave":TopTitle="�༭"
			  Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
			  RSObj.Open "Select * From KS_Special Where SpecialID=" & SpecialID,Conn,1,1
			  If Not RSObj.Eof Then
			    ClassID       = RSObj("ClassID")
				SpecialName   = RSObj("SpecialName")
				SpecialEName  = RSObj("SpecialEName")
				TemplateID    = RSObj("TemplateID")
				FsoSpecialIndex=RSObj("FsoSpecialIndex")
				AddDate       = RSObj("SpecialAddDate")
				SpecialNote   = RSObj("SpecialNote")
				PhotoUrl      = RSObj("PhotoUrl")
			  End If
			Else
			  ClassID=KS.G("ClassID"):TopTitle="���":Action="AddSave":AddDate=Now:FsoSpecialIndex="Index.html"
			End If
			If KS.IsNul(SpecialNote) Then SpecialNote=" "
			With KS
			.echo "<html>" & vbCrLf
			.echo "<head>" & vbCrLf
			.echo "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
			.echo "<link href='Include/admin_style.css' rel='stylesheet'>" & vbCrLf
		    .echo "<script language='JavaScript' src='../KS_Inc/common.js'></script>"
	      	.echo "<script language='JavaScript' src='../KS_Inc/jquery.js'></script>"
			.echo "<title>ר�����</title>" & vbCrLf
			.echo "</head>" & vbCrLf
			.echo "<body  leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
			.echo "<div class='topdashed sort'>" & TopTitle &"ר��</div>" & vbCrLf

			 
			.echo "  <form action='KS.Special.asp?Action=" & Action & "' name='SpecialForm' method='post'>" & vbCrLf
			.echo "  <input name='SpecialID' type='hidden' id='SpecialID' value='" & SpecialID &"'>" & vbCrLf
			.echo "  <table width='99%' border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
			.echo "    <tr>" & vbCrLf
			.echo "      <td>" & vbCrLf
			.echo "      <FIELDSET align=center>" & vbCrLf
			.echo "  <LEGEND align=left>" & TopTitle & "ר��</LEGEND>" & vbCrLf
            .echo "        <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class='ctable'>" & vbCrLf
			.echo "       <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
			.echo "        <td height='35' class='clefttitle'> <div align='right'><strong>����ר����ࣺ</strong></div></td>" & vbCrLf
			.echo "         <td width='542'>" & vbCrLf
			.echo "         <select name='ClassID' class='textbox'>" & vbCrLf
					  
					  Dim FolderName, TempStr, FolderRS
						Set FolderRS = Server.CreateObject("ADODB.Recordset")
						TempStr = "<option value=0>--��ѡ��ר�����--</option>"
					  FolderRS.Open "Select ClassID,ClassName From KS_SpecialClass Order BY OrderID", Conn, 1, 1
					If Not FolderRS.EOF Then
					  Do While Not FolderRS.EOF
						 FolderName = Trim(FolderRS(1))
						 If trim(ClassID) = Trim(FolderRS(0)) Then
						   TempStr = TempStr & "<option value=" & FolderRS(0) & " Selected>" & FolderName & "</option>"
						 Else
						   TempStr = TempStr & "<option value=" & FolderRS(0) & ">" & FolderName & "</option>"
						 End If
						 FolderRS.MoveNext
					  Loop
					End If
					FolderRS.Close:Set FolderRS = Nothing
					.echo TempStr
					
			.echo "        </select>" & vbCrLf
			.echo "            </td>" & vbCrLf
			.echo "    </tr>" & vbCrLf
			.echo "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
			.echo "      <td width='179' height='35'class='clefttitle'> <div align='right'><strong>ר�����ƣ�</strong></div></td>" & vbCrLf
			.echo "      <td> <input name='SpecialName' value='" & SpecialName & "' type='text' id='SpecialName' size='30' class='textbox'>"
			.echo "              �ſ���˵������ </td>"
			.echo "    </tr>" & vbCrLf
			.echo "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.echo "      <td height='35' class='clefttitle'> <div align='right'><strong>ר��Ŀ¼��</strong></div></td>" & vbCrLf
			.echo "      <td>"
			.echo "<input"
				If KS.G("Action")="Edit" Then .echo " Disabled"
			.echo " name='SpecialEName' type='text' value='" & SpecialEName & "' id='SpecialEName' size='30' class='textbox'>"
			.echo "        ���ܴ�\/��*���� < > | ���������,����һ���趨�Ͳ��ܸģ�������  </td>"
			.echo "    </tr>"
			.echo "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.echo "      <td height='35' class='clefttitle'> <div align='right'><strong>ר��ҳģ�壺</strong></div></td>" & vbCrLf
			.echo "      <td><input type='text' size='30' name='TemplateID' id='TemplateID' value='" & TemplateID & "' class='textbox'>&nbsp;" & KSCls.Get_KS_T_C("$('#TemplateID')[0]") 
			.echo "      </td>" & vbCrLf
			.echo "    </tr>" & vbCrLf
				 .echo "          <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
				.echo "           <td height='40' nowrap class='clefttitle'> <div align='right'><strong>ר��ͼƬ��ַ��</strong></td>" & vbCrLf
				.echo "            <td height='28' nowrap>" & vbCrLf
				.echo "             <INPUT NAME='PhotoUrl' value='" & PhotoUrl &"' TYPE='text' id='PhotoUrl' class='textbox' size=30>"
				.echo "                  <input class=""button""  type='button' name='Submit' value='ѡ��ͼƬ...' onClick=""OpenThenSetValue('Include/SelectPic.asp?CurrPath=" & CurrPath & "',550,290,window,document.SpecialForm.PhotoUrl);"">  <input class=""button"" type='button' name='Submit' value='Զ��ץȡͼƬ...' onClick=""OpenThenSetValue('Include/Frame.asp?FileName=SaveBeyondfile.asp&PageTitle='+escape('ץȡԶ��ͼƬ')+'&ItemName=ͼƬ&CurrPath=" & CurrPath & "',300,100,window,document.SpecialForm.PhotoUrl);"">"
				.echo "              </td>" & vbCrLf
				.echo "          </tr>" & vbCrLf
			.echo "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
			.echo "      <td height='35' class='clefttitle'> <div align='right'><strong>����ר��ҳ���ļ���</strong></td>" & vbCrLf
			.echo "      <td><select name='FsoSpecialIndex' class='textbox'>"
			.echo "          <option value='index.html' selected>index.html</option>"
			.echo "          <option value='index.htm'>index.htm</option>"
			.echo "          <option value='index.shtm'>index.shtm</option>"
			.echo "          <option value='index.shtml'>index.shtml</option>"
			.echo "          <option value='default.html'>default.html</option>"
			.echo "          <option value='default.htm'>default.htm</option>"
			.echo "          <option value='default.shtml'>default.shtml</option>"
			.echo "          <option value='default.shtm'>default.shtm</option>"
			.echo "          <option value='index.asp'>index.asp</option>"
			.echo "         <option value='" & FsoSpecialIndex & "' selected>" & FsoSpecialIndex & "</option>"
			.echo "        </select></td>"
			.echo "    </tr>" & vbCrLf
			
			.echo "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.echo "      <td & vbCrLfheight='35' class='clefttitle'> <div align='right'><strong>���ʱ�䣺</strong></div></td>"
			.echo "      <td><input name='SpecialAddDate' type='text' id='SpecialAddDate' value='" & AddDate & "' size='30' class='textbox'>"
			.echo "      </td>" & vbCrLf
			.echo "    </tr>" & vbCrLf
			.echo "    <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
			.echo "      <td height='35' class='clefttitle'> <div align='right'><strong>��Ҫ˵����</strong></div></td>"
			.echo "      <td><textarea name='SpecialNote' rows='8' id='SpecialNote' style='display:none;width:80%;border-style: solid; border-width: 1'>" & Server.HTMLEncode(SpecialNote) & "</textarea><iframe id=""content___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=SpecialNote&amp;Toolbar=Basic"" width=""100%"" height=""180"" frameborder=""0"" scrolling=""no""></iframe></td>"
			.echo "    </tr>"
			.echo "  </table>"
			.echo "          </FIELDSET>"
			.echo "       </td>"
			.echo "    </tr>"
			.echo "    </table>"
			.echo "  </form>"
			.echo "</body>"
			.echo "</html>"
			.echo "<Script Language='javascript'>" & vbCrLf
			.echo "<!--" & vbCrLf
			.echo "function CheckForm()" & vbCrLf
			.echo "{ var form=document.SpecialForm;" & vbCrLf
			.echo "    if (form.ClassID.value==0)" & vbCrLf
			.echo "    {"
			.echo "     alert('��ѡ������ר�����!');" & vbCrLf
			.echo "    form.ClassID.focus();" & vbCrLf
			.echo "    return false;" & vbCrLf
			.echo "    }"
			.echo "   if (form.SpecialName.value=='')"
			.echo "    {" & vbCrLf
			.echo "     alert('������ר������!');" & vbCrLf
			.echo "     form.SpecialName.focus();" & vbCrLf
			.echo "     return false;" & vbCrLf
			.echo "    }" & vbCrLf
			.echo "    if (form.SpecialEName.value=='')" & vbCrLf
			.echo "    {"
			.echo "     alert('������ר���Ӣ������!');" & vbCrLf
			.echo "     form.SpecialEName.focus();" & vbCrLf
			.echo "    return false;" & vbCrLf
			.echo "    }"
			.echo "    if (form.TemplateID.value=='')" & vbCrLf
			.echo "    {"
			.echo "     alert('���ר��ģ��!');" & vbCrLf
			.echo "     form.TemplateID.focus();" & vbCrLf
			.echo "    return false;" & vbCrLf
			.echo "    }"

			.echo "    if (CheckEnglishStr(form.SpecialEName,'Ŀ¼��Ӣ������')==false)" & vbCrLf
			.echo "     return false;" & vbCrLf
			.echo "    form.submit();" & vbCrLf
			.echo "}" & vbCrLf
			.echo "//-->" & vbCrLf
			.echo "</Script>"
		  End With
		End Sub
		
		'�������
		Sub SpecialAddSave()
		  Dim TemplateRS, TemplateSql,TempObj, SpecialRS, SpecialSql,SpecialName, SpecialEName, TemplateID, FsoSpecialIndex, SpecialAddDate, SpecialNote,PhotoUrl,ClassID
					 SpecialName = KS.G("SpecialName")
					 SpecialEName = KS.G("SpecialEName")
					 TemplateID = KS.G("TemplateID")
					 FsoSpecialIndex = KS.G("FsoSpecialIndex")
					 SpecialAddDate = KS.G("SpecialAddDate")
					 SpecialNote = Request.Form("SpecialNote")
					 PhotoUrl = KS.G("PhotoUrl")
					 ClassID = KS.ChkClng(KS.G("ClassID"))
			With KS 
				 If SpecialName <> "" Then
					If Len(SpecialName) >= 100 Then
						Call KS.AlertHistory("ר�����Ʋ��ܳ���50���ַ�!", -1):Exit Sub
					End If
				 Else
					Call KS.AlertHistory("������ר������!", -1):Exit Sub
				 End If
				 If SpecialEName <> "" Then
					If Len(SpecialEName) >= 50 Then
						Call KS.AlertHistory("ר��Ӣ�����Ʋ��ܳ���50���ַ�!", -1):Exit Sub
					End If
					Set TempObj = Conn.Execute("Select SpecialEName,SpecialName from KS_Special where SpecialName='" & SpecialName & "' OR SpecialEName='" & SpecialEName & "'")
					If Not TempObj.EOF Then
						 If Trim(TempObj(0)) = SpecialEName Then
						   Call KS.alertHistory("���ݿ����Ѵ��ڸ�ר��Ӣ������!", -1)
						 Else
						   Call KS.alertHistory("���ݿ����Ѵ��ڸ�ר������!", -1)
						 End If
						.End
					End If
				 Else
					Call KS.alert("������ר��Ӣ������!", "Special_Add.asp?ClassID=" & ClassID)
					.End
				 End If
				 If TemplateID = "" Then
					Call KS.alert("��ѡ��ר��ģ��", "Special_Add.asp?ClassID=" & ClassID)
					.End
				 End If
				
				  Set SpecialRS = Server.CreateObject("adodb.recordset")
				  SpecialSql = "select * from [KS_Special] Where (ID IS NULL)"
				  SpecialRS.Open SpecialSql, Conn, 1, 3
				  SpecialRS.AddNew
				  SpecialRS("ID") = Year(Now) & Month(Now) & Day(Now) & KS.MakeRandom(5)
				  SpecialRS("ClassID") = ClassID
				  SpecialRS("SpecialName") = SpecialName
				  SpecialRS("SpecialEName") = SpecialEName
				  SpecialRS("TemplateID") = TemplateID
				  SpecialRS("FsoSpecialIndex") = FsoSpecialIndex
				  SpecialRS("SpecialAddDate") = SpecialAddDate
				  SpecialRS("SpecialNote") = SpecialNote
				  SpecialRS("PhotoUrl") = PhotoUrl
				  SpecialRS("Creater") = KS.C("AdminName")
				  SpecialRS.Update
				  SpecialRS.MoveLast
				  Call KS.FileAssociation(1001,SpecialRS("SpecialID"),PhotoUrl&SpecialNote ,0)
				  SpecialRS.Close:Set SpecialRS = Nothing
				  .echo ("<script>if (confirm('���ר��ɹ�,���������?')==true){location.href='KS.Special.asp?action=Add&ClassID=" & ClassID & "';}else{location.href='KS.Special.asp?Action=SpecialList&ClassID=" & ClassID & "';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?Opstr=" & server.URLEncode("���ݹ��� >> ר�����") & "&ButtonSymbol=Disabled';}</script>")     
			End With
		End Sub
		'�����޸�
		Sub SpecialEditSave()
			Dim TemplateRS, TemplateSql,TempObj, SpecialRS, SpecialSql,SpecialName, SpecialEName, TemplateID, FsoSpecialIndex, SpecialAddDate, SpecialNote,PhotoUrl
					 SpecialName = KS.G("SpecialName")
					 TemplateID = KS.G("TemplateID")
					 FsoSpecialIndex = KS.G("FsoSpecialIndex")
					 SpecialAddDate = KS.G("SpecialAddDate")
					 SpecialNote = Request.Form("SpecialNote")
					 SpecialID   = KS.G("SpecialID")
					 PhotoUrl    = KS.G("PhotoUrl")
			With KS	 
				 If SpecialName <> "" Then
					If Len(SpecialName) >= 100 Then
						Call KS.AlertHistory("ר�����Ʋ��ܳ���50���ַ�!", -1)
						.End
					End If
				 Else
					Call KS.AlertHistory("������ר������!", -1)
					.End
				 End If
					
					Set TempObj = Conn.Execute("Select SpecialEName,SpecialName from KS_Special where SpecialName='" & SpecialName & "' And SpecialID<>" & SpecialID)
					If Not TempObj.EOF Then Call KS.alertHistory("���ݿ����Ѵ��ڸ�ר������!", -1): Exit Sub
				
				    If TemplateID = "" Then	Call KS.alertHistory("��ѡ��ר��ģ��",-1):Exit Sub

				
				  Set SpecialRS = Server.CreateObject("adodb.recordset")
				  SpecialSql = "select * from [KS_Special] Where SpecialID=" & SpecialID
				  SpecialRS.Open SpecialSql, Conn, 1, 3
				  SpecialRS("ClassID") = ClassID
				  SpecialRS("SpecialName") = SpecialName
				  SpecialRS("TemplateID") = TemplateID
				  SpecialRS("FsoSpecialIndex") = FsoSpecialIndex
				  SpecialRS("SpecialAddDate") = SpecialAddDate
				  SpecialRS("SpecialNote") = SpecialNote
				  SpecialRS("PhotoUrl") = PhotoUrl
				  SpecialRS.Update
				  Call KS.FileAssociation(1001,SpecialID,PhotoUrl&SpecialNote ,1)
				  SpecialRS.Close:Set SpecialRS = Nothing
				  .echo ("<script>alert('ר����Ϣ�޸ĳɹ�');location.href='KS.Special.asp?Action=SpecialList&ClassID=" & ClassID & "';$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?Opstr=" & Server.URLEncode("����ר��") & "&ButtonSymbol=Disabled&ClassID=" & ClassID & "';</script>")     
			End With
		End Sub
		
		Sub GetChannelList()
		  With KS
		  	.echo (" <div style=""border:1px solid #000000;position:absolute;top:10;right:8; overflow:hidden;"" >")
		    Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
		    RSObj.Open "Select ClassID,ClassName From KS_SpecialClass ",conn,1,1
		    If Not RSObj.Eof Then
			 .echo "<select style=""margin:-2px;"" OnChange=""location.href=this.value;"">"
			 .echo "<option value='KS.Special.asp?Action=SpecialList'>--���������ר��--</option>"
			 Do While Not RSObj.Eof
			   If ClassID=Trim(RSObj(0)) Then
			   	 .echo "<option value='KS.Special.asp?Action=SpecialList&ClassID=" & RSObj(0) &"' selected>" & RSObj(1) &"</option>"
			   Else
			   .echo "<option value='KS.Special.asp?Action=SpecialList&ClassID=" & RSObj(0) &"'>" & RSObj(1) &"</option>"
			   End If
			   RSObj.MoveNext
			 Loop
			  .echo "</select>"
			Else
			 .echo "<select style=""margin:-2px;"" OnChange=""location.href=this.value;"">"
			 .echo "<option value='KS.Special.asp'>--��û������κη���--</option>"
			 .echo "</select>"
			End If
			 .echo "</div>"
		  End With  
		End Sub
		
		'ɾ��ר��
		Sub SpecialDel()
			Dim K, ID, SpecialRS, FolderPath,Page
			Set SpecialRS = Server.CreateObject("Adodb.RecordSet")
			ID = Trim(KS.G("SpecialID"))
			Page = KS.G("Page")
			If ID="" Then KS.AlertHintScript "��û��ѡ��ר��" : Exit Sub
			ID = Split(ID, ",")
			For K = LBound(ID) To UBound(ID)
				 SpecialRS.Open "Select * FROM KS_Special Where SpecialID=" & Trim(ID(K)), Conn, 1, 3
			  If SpecialRS.EOF And SpecialRS.BOF Then
				Call KS.AlertHistory("�������ݳ���!", -1):Exit Sub
			  Else
				   If KS.Setting(95) = "/" Or KS.Setting(95) = "\" Then
					   FolderPath = KS.Setting(3) & SpecialRS("SpecialEName")
				   Else
					   FolderPath = KS.Setting(3) & KS.Setting(95) & SpecialRS("SpecialEName")
				   End If
			       If KS.DeleteFolder(FolderPath) = False Then  Call KS.AlertHistory("error!", -1):Exit Sub
				   Conn.Execute("Delete From KS_SpecialR Where SpecialID=" & ID(K))
				   Conn.Execute("Delete From KS_UploadFiles Where ChannelID=1001 and infoid=" & ID(K))
			  SpecialRS.Delete
			  SpecialRS.Close
			  End If
			Next
			If KeyWord = "" Then
			  Response.Write ("<script>location.href='KS.Special.asp?Action=SpecialList&Page=" & Page & "&ClassID=" & ClassID & "';</script>")
			Else
			  Response.Write ("<script>location.href='KS.Special.asp?Action=SpecialList&Page=" & Page & "&KeyWord=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate & "';</script>")
			End If
		End Sub
		
		'��ר���Ƴ����£�ͼƬ�����ص�
		Sub SpecialInfoDel()
		  Dim ID:ID = Replace(KS.G("ID")," ","")
		  ID=KS.FilterIDs(ID)
		  If ID="" Then
		   KS.AlertHintScript "����!"
		  Else
		  Conn.Execute("Delete From KS_SpecialR Where ID in (" & ID & ")")
		  KS.AlertHintScript "��ϲ,�����ɹ�!"
		  End If
		End Sub
End Class
%> 
