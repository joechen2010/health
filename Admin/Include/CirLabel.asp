<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New CirLabel
KSCls.Kesion()
Set KSCls = Nothing

Class CirLabel
        Private KS,TempClassList, InstallDir, FolderID, LabelContent, L_C_A, Action, LabelID, Str, Descript
		Dim ChannelID,ShowClassName, ArticleListNumber, RowHeight, TitleLen, ArticleSort, ShowPicFlag,DateRule, DateAlign,ShowNewFlag,ShowHotFlag, PrintType,XslContent
		 Dim LabelRS, LabelName,innersql,outsql

		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Call KS.DelCahe(KS.SiteSn & "_cirlabellist")
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()

		FolderID = Request("FolderID")
		ChannelID=KS.ChkCLng(Request("ChannelID"))
		If ChannelID=0 Then ChannelID=1
		
		With KS
		 		.echo "<html>"
				.echo "<head>"
				.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
				.echo "<link href=""admin_style.css"" rel=""stylesheet"">"
				.echo "<script src=""../../ks_inc/Common.js"" language=""JavaScript""></script>"
				.echo "<script src=""../../ks_inc/jquery.js"" language=""JavaScript""></script>"

		If KS.G("Action")="DoSave" Then Call DoSave()
        
		'�ж��Ƿ�༭
		LabelID = Trim(Request.QueryString("LabelID"))
		If LabelID = "" Then
		 outsql="SELECT TOP 10 ID,FolderName FROM [KS_Class] Where ChannelID=1 and ClassType=1 ORDER BY FolderOrder"
		 innersql="SELECT TOP 10 id,title,adddate FROM [KS_Article] Where Tid='{R:ID}' Order By ID Desc"
		 XslContent="<?xml version=""1.0"" encoding=""GB2312""?>"&vbcrlf & _  
                     "<xsl:stylesheet version=""1.0"" xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"">"&vbcrlf & _  
					 "<xsl:output method=""xml"" omit-xml-declaration=""yes"" indent=""yes"" version=""4.0""/>" & vbcrlf & _
					 "<xsl:template match=""/"">"&vbcrlf & _ 
					 " <div class=""class_loop"">"&vbcrlf & _ 
					 "  <xsl:for-each select=""xml/outerlist/outerrow"">"&vbcrlf & _ 
					 "   <div class=""loop_content"">"&vbcrlf & _  
					 "     <div class=""loop_title"">"&vbcrlf & _  
					 "       <span class=""classname""><a href=""{@classlink}""><xsl:value-of select=""@foldername"" disable-output-escaping=""yes"" /></a></span>"&vbcrlf & _ 
					 "       <span class=""class_more""><a href=""{@classlink}"">����</a></span>"&vbcrlf & "     </div>" &vbcrlf & _ 
					 "     <div class=""loop_list"">"&vbcrlf & _ 
					 "      <ul>"&vbcrlf & _ 
					 "      <xsl:for-each select=""innerlist/innerrow"">"&vbcrlf & _  
					 "      <li><a href=""{@linkurl}"" title=""{@title}"" target=""_blank"">{KS:CutText(<xsl:value-of select=""@title"" disable-output-escaping=""yes""/>,20,""..."")}</a></li>"&vbcrlf & _  
					 "      </xsl:for-each>"&vbcrlf & _ 
					 "      </ul>"&vbcrlf & _ 
					 "     </div>"&vbcrlf & _
					 "  </div>"&vbcrlf &vbcrlf & _ 
					 "   </xsl:for-each>"&vbcrlf & _  
					 "   </div>"&vbcrlf & _  
					 "</xsl:template>"&vbcrlf & _  
					 "</xsl:stylesheet>"
		Else
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
			Descript = Split(LabelRS("Description"),"@@@")
			LabelContent = LabelRS("LabelContent")
			LabelRS.Close
			Set LabelRS = Nothing
			LabelContent       = Replace(Replace(LabelContent, "{Tag:GetCirList", ""),"}{/Tag}", "")
			Dim XMLDoc,Node
			Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			If XMLDoc.loadxml("<label><param " & LabelContent & " /></label>") Then
			  Set Node=XMLDoc.DocumentElement.SelectSingleNode("param")
			Else
			 .echo ("<Script>alert('��ǩ���س���!');history.back();</Script>")
			 Exit Sub
			End If
			If  Not Node Is Nothing Then
			ChannelID=Node.getAttribute("channelid")
			DateRule = Node.getAttribute("daterule")
			End If 
			XmlDoc=Empty
			Set Node=Nothing
			OutSql=Descript(0)
			InnerSql=Descript(1)
			XslContent=Descript(2)
		End If
		If PrintType="" Then PrintType=1
		%>
		<script language="javascript">

		function CheckForm()
		{   if ($('input[name=LabelName]').val()=='')
			 {
			  alert('�������ǩ����');
			  $('input[name=LabelName]').focus(); 
			  return false
			  }
			var ChannelID=1;
			var DateRule=document.myform.DateRule.value;
	
			document.myform.LabelContent.value=	'{Tag:GetCirList labelid="0" channelid="'+$('#channelid').val()+'" daterule="'+DateRule+'"}{/Tag}';
			document.myform.submit();
		}
		</script>
		<%
		.echo "</head>"
		.echo "<body topmargin=""0"" leftmargin=""0"">"
		.echo "<div align=""center"">"
		.echo "<form  method=""post"" name=""myform"" action=""CirLabel.asp"">"
		.echo " <input type=""hidden"" name=""LabelContent"">"
		.echo " <input type=""hidden"" name=""LabelFlag"" value=""1"">"
		.echo " <input type=""hidden"" name=""Action"" value=""DoSave"">"
		 .echo " <table width='100%' height='25' border='0' cellpadding='0' cellspacing='1' bgcolor='#efefef' class='sort'>"
		 .echo "       <tr><td><div align='center'><font color='#990000'>"
		 .echo " ͨ��ѭ����ǩ"
		 .echo "    </font></div></td></tr>"
		 .echo "    </table>"
%>
       <table border='0' cellspacing='1' cellpadding='1' width='98%' align='center' class='ctable'>
		  <form action="?" method="post" name="myform">
		   <input name='lbtf' type='hidden'>
		   <input type='hidden' name='labelid' value='<%=labelid%>'>
		  <tr class='tdbg'>
		    <td class='clefttitle' align='right'><strong>��ǩ����:</strong></td>
		    <td><input name="LabelName" value="<%=LabelName%>" onblur='testlabelname()' style="width:200;"> <font color=red>*</font><span id='labelmessage'></span>�����ǩ���ƣ�&quot;ѭ�������б�&quot;������ģ���е��ã�<font color="#FF0000">&quot;{LB_ѭ�������б�}&quot;</font>��</td>
		  </tr>
		  <tr class='tdbg'>
		   <td width='100'  height="30" class='clefttitle' align='right'><strong>��ǩĿ¼:</strong></td>
		   <td><%=KS.ReturnLabelFolderTree(FolderID, 6)%><font color=""#FF0000"">��ѡ���ǩ����Ŀ¼���Ա��պ�����ǩ</font></td>
		  </tr>
		
<%
		
		.echo "  <tr class=tdbg>"
		.echo "    <td height=""24"" align='right' class='clefttitle'><strong>�����ʽ:</strong></td>"
		.echo "    <td><span style='display:none'><select class='textbox'  name=""PrintType"">"
        .echo "  <option value=""1"""
		If PrintType="1" Then .echo " selected"
		.echo ">��ͨ��ʽ</option>"
        .echo " <option value=""3"""
		If PrintType="3" Then .echo " selected"
		.echo ">Ajax���</option>"
        
        .echo "</select></span>"
		.echo "       �ڲ�SQL��ѯ��ģ��<select name='channelid' id='channelid'><option value='0'>--ѡ��ģ��--</option>"
		Dim RSC:Set RSC=conn.execute("select ChannelID,ChannelName From KS_Channel Where ChannelStatus=1 and channelid<>6 and channelid<>9 and channelid<>10 order by channelid")
		do while not rsc.eof
		 if trim(channelid)=trim(rsc(0)) then
		 .echo "<option value='" & rsc(0) & "' selected>" & rsc(1) & "</option>"
		 else
		 .echo "<option value='" & rsc(0) & "'>" & rsc(1) & "</option>"
		 end if
		rsc.movenext
		loop
		rsc.close
		set rsc=nothing
		.echo " </select>������ʽ <select style='width:150px'  class=""textbox"" name=""DateRule"" id=""DateRule"">"
		.echo KS.ReturnDateFormat(DateRule)
		.echo "                  </select>"
		.echo "<br><font color=green>tips:��ʹ�ñ�ǩ@linkurlʱ��������ȷѡ��ģ�ͣ�����ò�����ȷ����Ϣurl</font></td>"


		.echo "              </tr>"

		.echo "            <tr class='tdbg'>"
		.echo "            <td class='clefttitle' align='right'><b>SQL��䣺</b></td>"
		.echo "            <td> <strong>���SQL��䣺</strong><textarea name='outsql' style='width:90%;height:40px'>" & outsql & "</textarea>"
		.echo "            <br><strong>�ڲ�SQL��䣺</strong>"
		.echo "            <textarea name='innersql' style='width:90%;height:40px'>" & innersql & "</textarea>"
		.echo "<br>SQL�����ñ�ǩ����ǰ��ĿID<font color=red>{$CurrClassID}</font>;��ǰ��ĿID����Ŀ¼ID��<font color=red>{$CurrClassChildID}</font>;<br>�ڲ�SQL�����ñ�ǩ<font color=green>{R:�ֶ���}</font>�����SQL��ǩ������"
		.echo "            </td>"
		.echo "            </tr>"
		
		.echo "            <tr class='tdbg'>"
		.echo "            <td class='clefttitle' align='right'><strong>XSLT��ǩ˵����</strong></td>"
		.echo "            <td>Ԥ���ǩ��<font color=red> @classlink</font> �õ���Ŀ���� <font color=red>@linkurl</font> �õ���Ϣ����<br> ��ǩ�Ĺ�����򣺸�������ѯ���ֶ���ǰ��<font color=blue>@</font>��ɣ��������ֶ�������Сд��<br>�磺select top 10 <font color=red>id,title</font> from ks_article ������ֶ�Ϊ<font color=red>@id</font>��<font color=red>@title</font>������"
		.echo "            </td>"
		.echo "            </tr>"
		
		.echo "  <tr class=tdbg>"
		.echo "    <td height=""24"" align='right' class='clefttitle'><strong>XSLT��ʽ:</strong></td>"
		.echo "    <td><textarea name='xslContent' style='width:98%;height:250px'>" & XslContent & "</textarea>"
		.echo "    <br><font color=blue>˵����xlst��ʽ�����ϸ���xslt�﷨��д��<br>���ý�ȡ�ַ����Ⱥ���<font color=red>{KS:CutText(title,len,'...')} </font> </font><br>���ú���ʹ��˵����<br><font color=red>title</font>Ҫ��ȡ������;<br><font color=red>len</font>��ȡ�ַ�����һ�������������ַ�;<br> <font color=red>...</font>��ʾ����ȡ���ʡ�Է�</td>"
		.echo "           </tr>"

		.echo "                  </table>"	
		.echo "</form>"
		.echo "</div>"
		.echo "</body>"
		.echo "</html>"
		End With
		End Sub
		
		'����
		Sub DoSave()
					LabelName = KS.G("LabelName")
					LabelID  = KS.G("LabelID")
					Descript = Request("LabelIntro")
					LabelContent = Trim(Request.Form("LabelContent"))
					FolderID = KS.G("ParentID")
					If LabelName = "" Then
					   Call KS.AlertHistory("��ǩ���Ʋ���Ϊ��!", -1)
					   Set KS = Nothing
					   Exit Sub
					End If
					
					If LabelContent = "" Then
					  Call KS.AlertHistory("��ǩ���ݲ���Ϊ��!", -1)
					  Set KS = Nothing
					  Exit Sub
					End If
					LabelName = "{LB_" & LabelName & "}"
					Set LabelRS = Server.CreateObject("Adodb.RecordSet")
					LabelRS.Open "Select LabelName From [KS_Label] Where ID<>'" & LabelID & "' and LabelName='" & LabelName & "'", Conn, 1, 1
					If Not LabelRS.EOF Then
					  Call KS.AlertHistory("��ǩ�����Ѿ�����!", -1)
					  LabelRS.Close
					  Conn.Close
					  Set LabelRS = Nothing
					  Set Conn = Nothing
					  Set KS = Nothing
					 Exit Sub
					Else
						LabelRS.Close
						LabelRS.Open "Select * From [KS_Label] Where ID='" & LabelID & "'", Conn, 1, 3
						If LabelRS.Eof Then
						 LabelRS.AddNew
						  Do While True
							'����ID  ��+12λ���
							LabelID = Year(Now()) & KS.MakeRandom(10)
							Dim RSCheck:Set RSCheck = Conn.Execute("Select ID from [KS_Label] Where ID='" & LabelID & "'")
							 If RSCheck.EOF And RSCheck.BOF Then
							  RSCheck.Close
							  Set RSCheck = Nothing
							  Exit Do
							 End If
						  Loop
						 LabelRS("ID") = LabelID
						 LabelRS("AddDate") = Now
						 LabelRS("LabelType") = 6
						 LabelRS("OrderID") = 1
						 LabelRS("LabelFlag") = 6
						End If
						 LabelRS("LabelName") = LabelName
						 LabelRS("LabelContent") = LabelContent
						 LabelRS("Description") = Request("OutSql") & "@@@" & Request("InnerSQL") & "@@@" & Request("xslContent")
						 LabelRS("FolderID") = FolderID
						 LabelRS.Update
						 If LabelID="" Then
						  Call KS.FileAssociation(1021,1,LabelContent&Request("xslContent"),1)
						ks.echo ("<script>if (confirm('�ɹ���ʾ:\n\n��ӱ�ǩ�ɹ�,������ӱ�ǩ��?')){location.href='CirLabel.asp?Action=AddNew&LabelType=6&FolderID=" & FolderID & "';}else{$(parent.document).find('#BottomFrame')[0].src='" & KS.Setting(3) & KS.Setting(89) & "KS.Split.asp?LabelFolderID=" & FolderID & "&OpStr=��ǩ���� >> ѭ����ǩ&ButtonSymbol=FreeLabel';parent.frames['MainFrame'].location.href='Label_Main.asp?LabelType=6&FolderID=" & FolderID & "';}</script>")
						Else
						 	 '�������б�ǩ���ݣ��ҳ����б�ǩ��ͼƬ
							 Dim Node,UpFiles,RCls
							 UpFiles=LabelContent&Request("xslContent")
							 if Not IsObject(Application(KS.SiteSN&"_labellist")) Then
								 Set RCls=New Refresh
								 Call Rcls.LoadLabelToCache()
								 Set Rcls=Nothing
							 End If
							 For Each Node in Application(KS.SiteSN&"_labellist").DocumentElement.SelectNodes("labellist")
								   UpFiles=UpFiles & Node.Text
							 Next
							 Call KS.FileAssociation(1021,1,UpFiles,1)
							 '������������

						ks.echo ("<script>alert('�ɹ���ʾ:\n\n��ǩ�޸ĳɹ�!');$(parent.document).find('#BottomFrame')[0].src='" & KS.Setting(3) & KS.Setting(89) & "KS.Split.asp?LabelFolderID=" & FolderID & "&OpStr=��ǩ���� >> ѭ����ǩ&ButtonSymbol=FreeLabel';parent.frames['MainFrame'].location.href='Label_Main.asp?LabelType=6&FolderID=" & FolderID & "';</script>")
						End If
					End If
			End Sub
End Class
%> 
