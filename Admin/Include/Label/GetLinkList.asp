<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New GetLinkList
KSCls.Kesion()
Set KSCls = Nothing

Class GetLinkList
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		Dim FolderID, LabelContent, L_C_A, Action, LabelID, Str, Descript
		Dim show,ClassID, LinkType, ShowStyle, LogoWidth, LogoHeight, ListNumber, TitleLen, ColNumber,RollWidth,RollHeight,RollSpeed,recommend
		FolderID = Request("FolderID")
		With KS
		'�ж��Ƿ�༭
		LabelID = Trim(Request.QueryString("LabelID"))
		If LabelID = "" Then
		  Action = "Add"
		Else
			Action = "Edit"
		  Dim LabelRS, LabelName
		  Set LabelRS = Server.CreateObject("Adodb.Recordset")
		  LabelRS.Open "Select * From KS_Label Where ID='" & LabelID & "'", Conn, 1, 1
		  If LabelRS.EOF And LabelRS.BOF Then
			 LabelRS.Close
			 Set LabelRS = Nothing
			 .echo ("<Script>alert('�������ݳ���!');window.close();</Script>")
			 .End
		  End If
			LabelName = Replace(Replace(LabelRS("LabelName"), "{LB_", ""), "}", "")
			FolderID = LabelRS("FolderID")
			Descript = LabelRS("Description")
			LabelContent = LabelRS("LabelContent")
			LabelRS.Close
			Set LabelRS = Nothing
			LabelContent       = Replace(Replace(LabelContent, "{Tag:GetLinkList", ""),"}{/Tag}", "")
			Dim XMLDoc,Node
			Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			If XMLDoc.loadxml("<label><param " & LabelContent & " /></label>") Then
			  Set Node=XMLDoc.DocumentElement.SelectSingleNode("param")
			Else
			 .echo ("<Script>alert('��ǩ���س���!');history.back();</Script>")
			 Exit Sub
			End If
			If  Not Node Is Nothing Then
				show = Node.getAttribute("show")
				ClassID = Node.getAttribute("classid")
			    LinkType = Node.getAttribute("linktype")
				ShowStyle = Node.getAttribute("showstyle")
				LogoWidth = Node.getAttribute("logowidth")
				LogoHeight = Node.getAttribute("logoheight")
				ListNumber = Node.getAttribute("num")
				TitleLen = Node.getAttribute("titlelen")
				recommend=Node.getAttribute("recommend")
				ColNumber = Node.getAttribute("col")
				RollWidth = Node.getAttribute("rollwidth")
				RollHeight= Node.getAttribute("rollheight")
				RollSpeed = Node.getAttribute("rollspeed")
		   End If
		   Set Node=Nothing
		   XMLDoc=Empty
		End If
		If Show="" Then show=0
		If LinkType = "" Then LinkType = 1
		If ShowStyle = "" Then ShowStyle = 2
		If LogoWidth = "" Then LogoWidth = 88
		If LogoHeight = "" Then LogoHeight = 31
		If RollWidth = "" Then LogoWidth = 200
		If RollHeight = "" Then LogoHeight = 150
		If ListNumber = "" Then ListNumber = 0
		If TitleLen = "" Then TitleLen = 30
		If ColNumber = "" Then ColNumber = 7
		If RollSpeed="" Then RollSpeed=5
		If recommend="" Then recommend=0
		
		.echo "<html>"
		.echo "<head>"
		.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		.echo "<link href=""../admin_style.css"" rel=""stylesheet"">"
		.echo "<script src=""../../../ks_inc/Common.js"" language=""JavaScript""></script>"
		.echo "<script src=""../../../ks_inc/Jquery.js"" language=""JavaScript""></script>"
		%>
		<script language="javascript">
		function SetLogoDisabled(Num)
		{
		if (Num==0)
		{
		 $("input[name=LogoWidth]").attr("disabled",true);
		 $("input[name=LogoHeight]").attr("disabled",true);
		}
		else
		{
		 $("input[name=LogoWidth]").attr("disabled",false);
		 $("input[name=LogoHeight]").attr("disabled",false);
		}
		}
		function SetDisabled(Num)
		{
		 if (Num==1||Num==3)
		  {
		  $("input[name=ColNumber]").attr("disabled",true);
		  }
		 else
		  {
		  $("input[name=ColNumber]").attr("disabled",false);
		  }
		 if (Num==1)
		 { $("#RollArea").show();
		 }else{
		   $("#RollArea").hide();
		 }
		}
		function CheckForm()
		{
		    if ($("input[name=LabelName]").val()=='')
			 {
			  alert('�������ǩ����');
			  $("input[name=LabelName]").focus(); 
			  return false
			  }
			var show,LinkType,ShowStyle;
			var ClassID=$("#ClassID").val();
			var LogoWidth=$("input[name=LogoWidth]").val();
			var LogoHeight=$("input[name=LogoHeight]").val();
			var RollWidth=$("input[name=RollWidth]").val();
			var RollHeight=$("input[name=RollHeight]").val();
			var RollSpeed=$("input[name=RollSpeed]").val();
			var ListNumber=$("input[name=ListNumber]").val();
			var TitleLen=$("input[name=TitleLen]").val();
			var ColNumber=$("input[name=ColNumber]").val();
			var show=$("input[name=show][checked=true]").val();
			var LinkType=$("input[name=LinkType][checked=true]").val();
			var ShowStyle=$("input[name=ShowStyle][checked=true]").val();
			var recommend=0;
			if ($("#recommend").attr("checked")==true)
			{
			  recommend=1;
			}
			var tagVal='{Tag:GetLinkList labelid="0" show="'+show+'" classid="'+ClassID+'" linktype="'+LinkType+'" showstyle="'+ShowStyle+'" logowidth="'+LogoWidth+'" logoheight="'+LogoHeight+'" rollwidth="'+RollWidth+'" rollheight="'+RollHeight+'" rollspeed="'+RollSpeed+'" num="'+ListNumber+'" titlelen="'+TitleLen+'" recommend="'+recommend+'" col="'+ColNumber+'"}{/Tag}'
			$("input[name=LabelContent]").val(tagVal);
			$("#myform").submit();
		}
		</script>
		<%
		.echo "</head>"
		.echo "<body topmargin=""0"" leftmargin=""0"" scroll=no>"
		.echo "<div align=""center"">"
		.echo "<iframe src='about:blank' name='_hiddenframe' id='_hiddenframe' width='0' height='0'></iframe>"
		.echo "<form  method=""post"" id=""myform"" name=""myform"" action=""AddLabelSave.asp"" target='_hiddenframe'>"
		.echo " <input type=""hidden"" name=""LabelContent"" id=""LabelConent"">"
		.echo "   <input type=""hidden"" name=""LabelFlag"" value=""2"">"
		.echo "  <input type=""hidden"" name=""Action"" value=""" & Action & """>"
		.echo "  <input type=""hidden"" name=""LabelID"" value=""" & LabelID & """>"
		.echo "  <input type=""hidden"" name=""FileUrl"" value=""GetLinkList.asp"">"
		.echo KS.ReturnLabelInfo(LabelName, FolderID, Descript)
		.echo "          <table width=""98%"" style='margin-top:5px' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""30"" colspan=""2"">�������&nbsp;"
		
					  .echo ("<Select Name=""ClassID"" id=""ClassID"" Class=""textbox"">")
					  .echo ("<option Value=""0"">-�г����з����վ��-</option>")
					  Dim ObjRS
					  Set ObjRS = Server.CreateObject("Adodb.Recordset")
					  ObjRS.Open "Select * From KS_LinkFolder Order BY OrderID,AddDate Desc", Conn, 1, 1
					  Do While Not ObjRS.EOF
					  If ClassID = Trim(ObjRS("FolderID")) Then
					   .echo ("<option value=" & ObjRS("FolderID") & " Selected>" & ObjRS("FolderName") & "</Option>")
					  Else
					   .echo ("<option value=" & ObjRS("FolderID") & ">" & ObjRS("FolderName") & "</Option>")
					  End If
					   ObjRS.MoveNext
					  Loop
					  ObjRS.Close
					  Set ObjRS = Nothing
					  .echo ("</Select>")
					   
		.echo "               </td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""30"" colspan=""2""> ��ʾ·��"

						If show = 0 Then
						 .echo ("<input type=""radio"" name=""show"" value=""0"" Checked>ֱ������URL ")
						Else
						 .echo ("<input type=""radio"" name=""show"" value=""0"">ֱ������URL ")
						End If
						If show = 1 Then
						 .echo ("<input type=""radio"" name=""show"" value=""1"" Checked>ͨ��ToLink.aspת��(���ۼƵ������) ")
						Else
						 .echo ("<input type=""radio"" name=""show"" value=""1"">ͨ��asp����ת��(���ۼƵ������) ")
						End If
						 
		.echo "                 </td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""30"" colspan=""2""> ��������"
						
						If LinkType = 2 Then
						 .echo ("<input type=""radio"" onclick=""SetLogoDisabled(2)"" name=""LinkType"" value=""2"" Checked>ȫ������ ")
						Else
						 .echo ("<input type=""radio"" onclick=""SetLogoDisabled(2)"" name=""LinkType"" value=""2"">ȫ������ ")
						End If
						If LinkType = 0 Then
						 .echo ("<input type=""radio"" onclick=""SetLogoDisabled(0)"" name=""LinkType"" value=""0"" Checked>�ı����� ")
						Else
						 .echo ("<input type=""radio"" onclick=""SetLogoDisabled(0)"" name=""LinkType"" value=""0"">�ı����� ")
						End If
						If LinkType = 1 Then
						 .echo ("<input type=""radio"" onclick=""SetLogoDisabled(1)"" name=""LinkType"" value=""1"" Checked>LOGO���� ")
						Else
						 .echo ("<input type=""radio"" onclick=""SetLogoDisabled(1)"" name=""LinkType"" value=""1"">LOGO���� ")
						End If
						 
		.echo "                 </td>"
		.echo "            </tr>"
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""30"" colspan=""2"">��ʾ��ʽ"
						
						If ShowStyle = 1 Then
						 .echo ("<input type=""radio"" onclick=""SetDisabled(1)"" name=""ShowStyle"" value=""1"" Checked>���Ϲ��� ")
						Else
						 .echo ("<input type=""radio"" onclick=""SetDisabled(1)"" name=""ShowStyle"" value=""1"">���Ϲ��� ")
						End If
						If ShowStyle = 2 Then
						 .echo ("<input type=""radio"" onclick=""SetDisabled(2)"" name=""ShowStyle"" value=""2"" Checked>�����б� ")
						Else
						 .echo ("<input type=""radio"" onclick=""SetDisabled(2)"" name=""ShowStyle"" value=""2"">�����б� ")
						End If
						If ShowStyle = 3 Then
						 .echo ("<input type=""radio"" onclick=""SetDisabled(3)"" name=""ShowStyle"" value=""3"" Checked>�����б� ")
						Else
						 .echo ("<input type=""radio"" onclick=""SetDisabled(3)"" name=""ShowStyle"" value=""3"">�����б� ")
						End If
		 .echo "              </td>"
		 .echo "           </tr>"
		 
		 .echo "          <tbody id=""RollArea"">"
		 .echo "           <tr class='tdbg'>"
		 .echo "             <td height=""30"" colspan='2'>���� : �������"
		 .echo "               <input name=""RollWidth"" id=""RollWidth"" class=""textbox""  onBlur=""CheckNumber(this,'�������');"" type=""text"" id=""RollWidth"" style=""width:50px;"" value=""" & RollWidth & """>�����߶�"
		 .echo "               <input name=""RollHeight""  class=""textbox""  onBlur=""CheckNumber(this,'�����߶�');"" type=""text"" id=""RollHeight"" style=""width:50px;"" value=""" & RollHeight & """> �����ٶ�<input name=""RollSpeed"" class=""textbox""  onBlur=""CheckNumber(this,'�����ٶ�');"" type=""text"" id=""RollSpeed"" style=""width:50px;"" value=""" & RollSpeed & """></td>"
		 .echo "           </tr>"
		 .echo "         </tbody>"
		 
		 .echo "           <tr class='tdbg'>"
		 .echo "             <td height=""30"">Logo���"
		 .echo "               <input name=""LogoWidth"" id=""LogoWidth"" class=""textbox""  onBlur=""CheckNumber(this,'Logo���');"" type=""text"" id=""LogoWidth"" style=""width:70%;"" value=""" & LogoWidth & """></td>"
		 .echo "             <td height=""30"">Logo�߶�"
		 .echo "               <input name=""LogoHeight"" id=""LogoHeight"" class=""textbox""  onBlur=""CheckNumber(this,'Logo�߶�');"" type=""text"" id=""LogoHeight"" style=""width:70%;"" value=""" & LogoHeight & """></td>"
		 .echo "           </tr>"
		 .echo "           <tr class='tdbg'>"
		 .echo "             <td width=""50%"" height=""30"">��ʾ��Ŀ"
		 .echo "               <input name=""ListNumber"" class=""textbox""  onBlur=""CheckNumber(this,'��������');"" type=""text"" id=""ListNumber"" style=""width:100px;"" value=""" & ListNumber & """><font color=""#FF0000"">����Ϊ0ʱ���г�������������վ��</font></td>"
		 .echo "             <td width=""50%"" height=""30""><label><input type='checkbox' value='1' id='recommend' name='recommend'"
		 if recommend="1" then .echo " checked"
		 .echo ">����ʾ�Ƽ�</label></td>"
		 .echo "           </tr>"
		 .echo "           <tr class='tdbg'>"
		 .echo "             <td height=""24"">��������"
		 .echo "               <input name=""TitleLen"" class=""textbox""  onBlur=""CheckNumber(this,'��������');"" type=""text"" id=""TitleLen"" style=""width:70%;"" value=""" & TitleLen & """></td>"
		 .echo "             <td height=""24""> ��ʾ����"
		 .echo "               <input name=""ColNumber"" class=""textbox""  onBlur=""CheckNumber(this,'��ʾ����');"" type=""text"" id=""ColNumber"" style=""width:70%;"" value=""" & ColNumber & """></td>"
		 .echo "           </tr>"
		.echo "                  </table>"	
		 .echo " </form>"
		.echo "</div>"
		.echo "</body>"
		.echo "</html>"
		.echo "<script>"
		.echo "SetLogoDisabled(" & LinkType & ");"
		.echo "SetDisabled(" & ShowStyle & ");"
		.echo "</script>"
		End With
		End Sub
End Class
%> 
