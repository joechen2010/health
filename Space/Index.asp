<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.SpaceCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************

Dim KSCls
Set KSCls = New SpaceIndex
KSCls.Kesion()
Set KSCls = Nothing

Class SpaceIndex
        Private KS, KSRFObj,UserName,Action,ID,Node,CurrPage,TotalPut,MaxPerPage,PageNum
		Private Template,TemplateSub,SubStr,BlogName,KSBCls
		Private Sub Class_Initialize()
		  MaxPerPage=10
		  Set KS=New PublicCls
		  Set KSRFObj = New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSRFObj=Nothing
		 Call CloseConn()
		End Sub
		Public Sub Kesion()
		    If KS.SSetting(0)=0 Then KS.Die "<script>alert('�Բ��𣬱�վ��رտռ�վ�㹦��!');window.close();</script>"
		    Dim QueryStrings:QueryStrings=Request.ServerVariables("QUERY_STRING")
			
			'ks.die QueryStrings
			If QueryStrings<>"" Then 
			 QueryStrings=KS.UrlDecode(QueryStrings)
			 Call Show(QueryStrings)
			Else
				Dim FileContent
				FileContent = KSRFObj.LoadTemplate(KS.SSetting(7))
				FCls.RefreshType = "SpaceINDEX" '����ˢ�����ͣ��Ա�ȡ�õ�ǰλ�õ�����
				FCls.RefreshFolderID = "0" '���õ�ǰˢ��Ŀ¼ID Ϊ"0" ��ȡ��ͨ�ñ�ǩ
				If Trim(FileContent) = "" Then FileContent = "�ռ���ҳģ�岻����!"
				FileContent=KSRFObj.KSLabelReplaceAll(FileContent)
				KS.Echo FileContent 
		   End If 
		End Sub
		
		Sub Show(ByVal QueryStrings)
		Dim QSArr:QSArr=Split(QueryStrings,"/")
		UserName=KS.DelSQL(QSArr(0))
		If Ubound(QSArr)>=1 Then Action=QSArr(1)
		If Ubound(QSArr)>=2 Then ID=KS.ChkClng(QSArr(2))
		If Ubound(QSArr)>=3 Then CurrPage=KS.ChkClng(QSArr(3))
		If UserName="" Then KS.Die "error username!"
		If CurrPage=0 Then CurrPage=1
		
		Set KSBCls=New BlogCls
		Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select top 1 * From KS_Blog Where UserName='" & UserName & "'",conn,1,1
		If RS.Eof And RS.Bof Then
		 rs.close:set rs=nothing
		 KS.Die "<script>location.href='index.asp';</script>"
		End If
		If KS.C("AdminName") = "" Or KS.C("AdminPass") = "" Then
			If RS("Status")=0 Then
			 rs.close:set rs=nothing
			 KS.Die "<script>alert('�ÿռ�վ����δ���!');window.close();</script>"
			elseif RS("Status")=2 then
			 rs.close:set rs=nothing
			 KS.Die "<script>alert('�ÿռ�վ���ѱ�����Ա����!');window.close();</script>"
			end if
		End If
		If KS.FoundInArr(KS.U_G(Conn.Execute("Select GroupID From KS_User Where UserName='" & UserName & "'")(0),"powerlist"),"s01",",")=false Then 
		 RS.Close : Set RS=Nothing
		 KS.Die "<script>location.href='../company/show.asp?username=" & username & "';</script>"
		End If
		
		'============================��¼���ʴ���������ÿ�============================================
		conn.execute("update KS_Blog Set Hits=Hits+1 Where UserName='" & UserName & "'")
		If KS.C("UserName")<>"" And KS.C("UserName")<>UserName Then
		   Dim RSV:Set RSV=Server.CreateObject("adodb.recordset")
		   RSV.Open "Select top 1 * From KS_BlogVisitor Where UserName='" & UserName & "' and Visitors='" & KS.C("UserName") & "'",conn,1,3
		   If RSV.Eof And RSV.Bof Then
		     RSV.AddNew
			 RSV("UserName")=UserName
			 RSV("Visitors")=KS.C("UserName")
		   End If
		    RSV("AddDate")=Now
			RSV.Update
		    RSV.Close : Set RSV=Nothing
		 End If
		'============================������¼============================================================
		 
		 Dim Xml:Set XML=KS.RsToXml(rs,"row","")
		 If Not IsObject(xml) Then KS.Die "error xml!"
		 Set Node=XML.DocumentElement.SelectSingleNode("row")
		 Set KSBCls.Node=Node
		 KSBCls.UserName=UserName
		 RS.Close : Set RS=Nothing
		 Dim TemplateID:TemplateID=Node.SelectSingleNode("@templateid").text
		 If Action<>"" Then template=Template & KSBCls.GetTemplatePath(TemplateID,"TemplateSub")
		 select case Lcase(action)
		   case "blog"
		      KSBCls.Title="����"
			  SubStr="<script language=""javascript"" defer>Page(1,'log&classid=" & KS.ChkClng(KS.S("ClassID")) & "&date=" & KS.S("Date") & "&key=" &KS.R(KS.S("Key"))&"&tag=" & KS.R(KS.S("Tag")) &"','" & UserName & "')</script><div id=""blogmain""></div><div id=""kspage"" align=""right""></div>"
		   case "log" Call BlogLog
		   case "album" 		    
		     KSBCls.Title="���"
			 SubStr="<script language=""javascript"" defer>Page(1,'photo','" & UserName & "')</script><div id=""blogmain""></div><div id=""kspage"" align=""right""></div>"
		   case "showalbum" Call ShowAlbum
		   case "group"
		     KSBCls.Title="Ȧ��"
			 SubStr="<script language=""javascript"" defer>Page(1,'group','" & UserName & "')</script><div id=""blogmain""></div><div id=""kspage"" align=""right""></div>"
		   case "friend"
		     KSBCls.Title="����"
			 SubStr="<script language=""javascript"" defer>Page(1,'friend','" & UserName & "')</script><div id=""blogmain""></div><div id=""kspage"" align=""right""></div>"
		   case "xx"
		     KSBCls.Title="�ļ�"
			 SubStr="<script language=""javascript"" defer>Page(1,'xx','" & UserName & "')</script><div id=""blogmain""></div><div id=""kspage"" align=""right""></div>"
		   case "info"
		     KSBCls.Title="����"
			 SubStr=KSBCls.UserInfo()
		   case "message"
		     KSBCls.Title="����"
		     Call ShowMessage
		   case "intro"
		     KSBCls.Title="��˾����"
			 SubStr=KSBcls.Location("<div align=""left""><strong>��ҳ >> ��˾���</strong></div>")
			 Dim Irs:Set Irs=Conn.Execute("Select top 1 Intro From KS_EnterPrise Where UserName='" & UserName & "'")
			 if Not Irs.Eof Then
			 SubStr=SubStr & KS.HtmlCode(Irs(0))
		     Else
		       Irs.Close: Set Irs=Nothing
		       KS.AlertHintScript "�Բ��𣬸��û�������ҵ�û���"
			 End If
			 Irs.Close:Set IrS=Nothing
		   case "news" KSBCls.Title="��˾��̬" : GetNews
		   case "shownews" ShowNews
		   case "product"  ProductList
		   case "showproduct" ShowProduct
		   case "ryzs" KSBCls.Title="����֤��" : GetRyzs
		   case "job" JobList
		   case "showphoto" ShowPhoto
		   case else
		    KSBCls.Title="��ҳ"
		    template=KSBCls.GetTemplatePath(TemplateID,"TemplateMain")
		 end select
		  template=Replace(Template,"{$BlogMain}",SubStr)
		  template=KSBCls.ReplaceBlogLabel(Template)
		  KS.Echo KSBCls.LoadSpaceHead
		  KS.Echo Template
		  
		End Sub
		
		'��־
		Sub BlogLog()
		  If ID=0 Then KS.Die "error logid!"
		  Dim RS,i
		  Set RS=Server.CreateObject("ADODB.RECORDSET")
		  If KS.C("AdminName") = "" Or KS.C("AdminPass") = "" Then
		   RS.Open "Select top 1 * from KS_BlogInfo Where ID=" & ID & " and Status=0",conn,1,1
		  Else
		   RS.Open "Select top 1 * from KS_BlogInfo Where ID=" & ID,conn,1,1
		  End If
		  If RS.EOF And RS.BOF Then
			KS.Die "<script>alert('�������ݳ�������־Ϊ�ݸ壡');history.back();</script>"
		  End If
		  KSBCls.Title=rs("title")
		  
		  SubStr= LFCls.GetConfigFromXML("space","/labeltemplate/label","log")
		  conn.execute("update KS_BlogInfo Set Hits=Hits+1 Where ID=" & ID)
		  
		   Dim EmotSrc:If RS("Face")<>"0" Then EmotSrc="<img src=""../User/images/face/" & RS("Face") & ".gif"" border=""0"">"
		   Dim TagList,TagsArr:TagsArr=Split(RS("Tags")," ")
				if RS("Tags")<>"" then
				TagList="<div style='display:none'><form id='mytagform' target='_blank' action='../space/?" & username & "/blog' method='post'><input type='text' name='tag' id='tag'></form></div><strong>��ǩ��</strong>"
				 For I=0 To Ubound(TagsArr)
				  If TagsArr(i)<>"" then
				    TagList=TagList &"<a href=""javascript:void(0)"" onclick=""$('#tag').val('" & TagsArr(i) & "');$('#mytagform').submit();"">" & TagsArr(i) & "</a> "
				  end if
				 Next
				 TagList=TagList &"&nbsp;&nbsp;&nbsp;&nbsp;"
				end if

		    Dim MoreStr:MoreStr="�Ķ�����("&RS("hits")&") | �ظ���("& Conn.Execute("Select Count(ID) From KS_BlogComment Where LogID="  &id)(0) &")"
		   	Dim ContentStr
			
			Dim JFStr:If RS("Best")="1" then JFStr="  <img src=""../images/jh.gif"" align=""absmiddle"">" else JFStr=""

		    If KS.IsNul(RS("PassWord")) Then 
			
			 ContentStr=RS("Content")
			ElseIf KS.S("Pass")<>"" Then
			  If KS.S("Pass")=rs("password") then
			   ContentStr=RS("Content")
			  Else
			   SubStr="<br /><br />������,���������־��������!<a href='javascript:history.back(-1)'>����</a><br/>"
			   exit sub
			  End if
			Else
			 SubStr="<br/><br/><br/><form method='post' action='" & KSBCls.GetLogUrl(RS) & "'>��ƪ�����ѱ����˼�����,��������־�Ĳ鿴���룺<input style='border-style: solid; border-width: 1' type='password' name='pass' size='15'>&nbsp;<input type='submit' value=' �鿴 '></form>"
			  exit sub
			End IF
		   SubStr=Replace(SubStr,"{$ShowLogTopic}",EmotSrc & RS("Title") & jfstr)
		   SubStr=Replace(SubStr,"{$ShowLogInfo}","[" & RS("AddDate") & "|by:" & RS("UserName") & "]")
		   SubStr=Replace(SubStr,"{$ShowLogText}",KS.ReplaceInnerLink(ContentStr))
		   SubStr=Replace(SubStr,"{$ShowLogMore}", TagList&MoreStr)
		   
		   SubStr=Replace(SubStr,"{$ShowTopic}",RS("Title"))
		   SubStr=Replace(SubStr,"{$ShowAuthor}",RS("UserName"))
		   SubStr=Replace(SubStr,"{$ShowAddDate}",RS("AddDate"))
		   SubStr=Replace(SubStr,"{$ShowEmot}",EmotSrc)
		   SubStr=Replace(SubStr,"{$ShowWeather}",KSBCls.GetWeather(RS))
		   
           SubStr=SubStr & "��һƪ:" & ReplacePrevNextArticle(ID,"Prev")
           SubStr=SubStr & "<br>��һƪ:" & ReplacePrevNextArticle(ID,"Next") & "<br><br>"
		   
           SubStr=SubStr & "<div id=""commentmainlist""><script language=""javascript"" defer>CommentPage(1," & ID & ");</script></div><div id=""commentpagelist"" align=""right""></div><script src=""writecomment.asp?ID=" & ID & "&UserName=" & RS("UserName") & "&Title=" & RS("Title") & """></script>"
		   

		   RS.Close:Set RS=Nothing
		End Sub
		Function ReplacePrevNextArticle(NowID,TypeStr)
		    Dim SqlStr
			If Trim(TypeStr) = "Prev" Then
				   SqlStr = " SELECT Top 1 ID,Title From KS_BlogInfo Where UserName='" & UserName & "' And ID<" & NowID & " And Status=0 Order By ID Desc"
			ElseIf Trim(TypeStr) = "Next" Then
				   SqlStr = " SELECT Top 1 ID,Title From KS_BlogInfo Where UserName='" & UserName & "' And ID>" & NowID & " And Status=0 Order By ID Desc"
			Else
				ReplacePrevNextArticle = "":Exit Function
			End If
			 Dim RS:Set RS=Server.CreateObject("ADODB.Recordset")
			 RS.Open SqlStr, Conn, 1, 1
			 If RS.EOF And RS.BOF Then
				ReplacePrevNextArticle = "û����"
			 Else
			  ReplacePrevNextArticle = "<a href=""" & KSBCls.GetCurrLogUrl(RS("ID"),UserName) & """ title=""" & RS("Title") & """>" & RS("title") & "</a>"
			 End If
			 RS.Close:Set RS = Nothing
	 End Function
		
	'�鿴��Ƭ
	 Sub ShowAlbum()
	   If ID=0 Then KS.Die "error xcid!"
	    Dim RSXC:Set RSXC=Server.CreateObject("ADODB.RECORDSET")
		RSXC.OPEN "Select * from ks_photoxc where id=" & id,conn,1,3
		if rsxc.eof and rsxc.bof then
		  rsxc.close:set rsxc=nothing
		  KS.Die "<script>alert('�������ݳ���!');history.back();</script>"
		end if
	   If KS.C("AdminName") = "" Or KS.C("AdminPass") = "" Then
		If RSxc("Status")=0 Then
		 KS.Die "<script>alert('�������δ���!');window.close();</script>"
		elseif RSxc("Status")=2 then
		 KS.Die "<script>alert('������ѱ�����Ա����!');window.close();</script>"
		end if
	   End If
	   
	   Select Case rsxc("flag")
		   Case 1,2
		    If rsxc("Flag")=2 and KS.C("UserName")="" then
			  substr="<br><br>��������û�Ա�ɼ�������<a href=""../User/"" target=""_blank"">��¼</a>��"
			Else
			  GetAlbumBody
		    End If
		  Case 3
		    If KS.S("Password")=rsxc("password") or Session("xcpass")=rsxc("password") then
			   Session("xcpass")=KS.S("Password")
			   GetAlbumBody
			else
		      SubStr="<form action=""../space/?" & username &"/showalbum/" & xcid& """ method=""post"" name=""myform"" id=""myform"">������鿴���룺<input type=""password"" name=""password"" size=""12"" style='border-style: solid; border-width: 1'>&nbsp;<input type='submit' value=' �鿴 '></form>"
		   end if
		  Case 4
		    If KS.C("UserName")=rsxc("username") then
			  GetAlbumBody
			else
			  SubStr="<br><br><li>�������Ϊ��˽��ֻ��������˲���Ȩ�����!</li><li>�������������ˣ�<a href=""../User/""  target=""_blank"">��¼</a>�󼴿ɲ鿴!</li>"
			end if
		 End Select
		 rsxc("hits")=rsxc("hits")+1
		 rsxc.update
		 rsxc.close:set rsxc=nothing
	 End Sub
	 Sub GetAlbumBody()
	             Dim TotalNum,RS
				 Set RS=Server.CreateObject("ADODB.RECORDSET")
				 RS.Open "Select * from KS_Photozp Where xcid=" & id &" Order By ID Desc",conn,1,1
				 If RS.EOF And RS.BOF Then
				    RS.Close : Set RS=Nothing
					SubStr = "<p>�������û����Ƭ��</p>"
				 Else
				        TotalNum=RS.Recordcount
				        If CurrPage>TotalNum Or CurrPage<=0 Then CurrPage=1
				        RS.Move(CurrPage-1)
						Conn.Execute("Update KS_PhotoZP Set Hits=hits+1 Where id=" & rs("id"))
						SubStr="<div style='height:50px;line-height:50px;text-align:center'>���������Ҽ���ҳ��<a style='padding:3px;border:1px solid #cccccc' href='../space/?" & username & "/showalbum/" & id & "/" & CurrPage-1 & "'>��һ��</a> ��<font color=red>" & currpage & "</font>/" & TotalNum & "�� <a style='padding:3px;border:1px solid #cccccc' href='../space/?" & username & "/showalbum/" & id & "/" & CurrPage+1 & "'>��һ��</a> <a style='padding:3px;border:1px solid #cccccc' href=""" & RS("PhotoUrl") & """ target=""_blank"">�鿴ԭͼ</a></div><div style='padding-bottom:20px;text-align:center'><strong>���:</strong><font color=red>" & rs("hits") & "</font>�� <strong>��С:</strong>" & round(rs("photosize") /1024,2)  & " KB <strong>�ϴ�ʱ��:</strong>" & rs("adddate") & "</div><div style='text-align:center'><a href='../space/?" & username & "/showalbum/" & id & "/" & CurrPage+1 & "'><img src='" & RS("PhotoUrl") & "' alt=""" & rs("descript") & """ style='border:1px solid #efefef' onload=""if (this.width>450) this.width=450;""/></a></div><div style='padding-top:20px;text-align:center'>" & rs("descript") & "</div>"
		   		       RS.Close:Set  RS=Nothing
			    End If
				 SubStr=SubStr & "<script>document.onkeydown=chang_page;function chang_page(event){var e=window.event||event;var eObj=e.srcElement||e.target;var oTname=eObj.tagName.toLowerCase();if(oTname=='input' || oTname=='textarea' || oTname=='form')return;	event = event ? event : (window.event ? window.event : null);if(event.keyCode==37||event.keyCode==33){location.href='../space/?" & username & "/showalbum/" & id & "/" & currpage-1 &"'}	if (event.keyCode==39 ||event.keyCode==34){location.href='../space/?" & username & "/showalbum/" & id & "/" & currpage+1 & "'}}</script>"
		End Sub
		Function GetStatusStr(val)
           Select Case Val
		    Case 1:GetStatusStr="����"
			Case 2:GetStatusStr="��Ա"
			Case 3:GetStatusStr="����"
			Case 4:GetStatusStr="��˽"
		   End Select
			GetStatusStr="<font color=red>" & GetStatusStr & "</font>"
		End Function
		
		'����
		Sub ShowMessage()
		 SubStr="<div id=""guestmain""><script language=""javascript"" defer>GuestPage(1,'guest','" & UserName & "')</script></div><div id=""guestpage""></div>" &  GetWriteMessage() & "</div><div id=""kspage"" style=""display:none""></div>"
		End Sub
		Function GetWriteMessage()
		%>
		<script type="text/javascript">
		function success()
			{var editor = FCKeditorAPI.GetInstance("Content");
				var loading_msg='\n\n\t���Եȣ������ύ����...';
				var content=document.getElementById('Content');
				
				if (loader.readyState==1)
					{
						editor.EditorDocument.body.innerHTML=loading_msg;
					}
				if (loader.readyState==4)
					{   var s=loader.responseText;
						if (s=='ok')
						 {
						 alert('��ϲ,��������ѳɹ��ύ��');
						  if (typeof(loadDate)!="undefined") loadDate(1);
						  leavePage();
						 }
						else
						 {alert(s);
						  editor.EditorDocument.body.innerHTML=document.getElementById("scontent").value;
						 }
					}
			}
		var OutTimes =11;
		function leavePage()
		{
			var editor = FCKeditorAPI.GetInstance("Content");
		if (OutTimes==0)
		 {
		 editor.EditorDocument.body.disabled=false;
		 document.getElementById('SubmitComment').disabled=false;
		 editor.EditorDocument.body.innerHTML=''
		 document.getElementById('Title').value='';
		 document.getElementById('VerifyCode').value='';
		 OutTimes =11;
		 return;
		 }
		else {
			document.getElementById('SubmitComment').disabled=true;
			OutTimes -= 1;
			editor.EditorDocument.body.disabled=true;
            editor.EditorDocument.body.innerHTML="\n\n�������ύ���ȴ� "+ OutTimes + " ���Ӻ����ɼ�������...";
			setTimeout("leavePage()", 1000);
			}
		}
		function getCode()
		{
		 document.getElementById('showVerify').innerHTML='<IMG style="cursor:pointer" src="../plus/verifycode.asp" onClick="this.src=\'../plus/verifycode.asp?n=\'+ Math.random();" align="absmiddle">'
		}
		function CheckForm()
		{
		if (document.getElementById('Title').value=='')
		{
		 alert('���������Ա���!');
		 document.getElementById('Title').focus()
		 return false;
		}
		if (document.getElementById('AnounName').value=='')
		{
		 alert('����������ǳƣ�');
		 document.getElementById('AnounName').focus()
		 return false;
		}
		if (FCKeditorAPI.GetInstance('Content').GetXHTML(true)=='')
		{
		 alert('��������������!');
		 FCKeditorAPI.GetInstance('Content').Focus();
		 return false;
		}

		if (document.getElementById('VerifyCode').value=='')
		{
		 alert('��������֤�룡');
		 document.getElementById('VerifyCode').focus()
		 return false;
		}
		document.getElementById("content").value=document.getElementById("scontent").value=FCKeditorAPI.GetInstance('Content').GetXHTML(true);
		ksblog.ajaxFormSubmit(document.myform,'success')
		}
		function ShowLogin()
		{ 
		 popupIframe('��Ա��¼','<%=KS.Setting(3)%>user/userlogin.asp?Action=Poplogin',397,184,'no');
		}
  </script>
		<%
		 If KS.SSetting(25)="0" And KS.IsNul(KS.C("UserName")) Then
		  GetWriteMessage="<div style=""margin:20px""><strong>��ܰ��ʾ��</strong>ֻ�л�Ա�ſ�������,����ǻ�Ա����<a href=""javascript:ShowLogin()"">��¼</a>,���ǻ�Ա����<a href=""../user/reg"" target=""_blank"">ע��</a>��</div>"
		 Else
		 GetWriteMessage = "<a name=""write""></a><table width=""98%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" class=""comment_write_table"">"
		 GetWriteMessage = GetWriteMessage & "<form name=""myform"" action=""../plus/ajaxs.asp?action=MessageSave"" method=""post"">"
		 GetWriteMessage = GetWriteMessage & "<input type=""hidden"" value=""" & UserName & """ name=""UserName"">"
		 GetWriteMessage = GetWriteMessage & "<input type=""hidden"" value="""" name=""scontent"">"
		 GetWriteMessage = GetWriteMessage & "<tr><td height=""30"" class=""comment_write_title""><strong>ǩд����:</strong></td></tr>"
		GetWriteMessage = GetWriteMessage & "<tr>"
		GetWriteMessage = GetWriteMessage & "<td>���⣺"
		GetWriteMessage = GetWriteMessage & "    <input name=""Title"" maxlength=""150"" value="""" type=""text"" id=""Title"" style=""width:280"" />&nbsp;<font color=red>*</font></td>"
		GetWriteMessage = GetWriteMessage & "  </tr>"
		GetWriteMessage = GetWriteMessage & "<tr>"
		GetWriteMessage = GetWriteMessage & "<td>��ҳ��"
		GetWriteMessage = GetWriteMessage & "    <input name=""HomePage"" maxlength=""150"" value=""http://"" type=""text"" id=""HomePage"" style=""width:200"" />������ҳ,���͵�ַ��</td>"
		GetWriteMessage = GetWriteMessage & "  </tr>"
		GetWriteMessage = GetWriteMessage & "  <tr>"
		GetWriteMessage = GetWriteMessage & "    <td height=""25""><textarea name=""Content"" rows=""6"" id=""Content"" style=""width:98%"" style=""display:none""></textarea><iframe id=""content___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=Content&amp;Toolbar=Basic"" width=""98%"" height=""150"" frameborder=""0"" scrolling=""no""></iframe></td>"
		GetWriteMessage = GetWriteMessage & "  </tr>"
		GetWriteMessage = GetWriteMessage & "<tr>"
		GetWriteMessage = GetWriteMessage & "  <td height=""30"" colspan=""2"">�ǳƣ�"
		GetWriteMessage = GetWriteMessage & "   <input name=""AnounName"" maxlength=""100"" type=""text"" value=""" & KS.C("UserName") & """ id=""AnounName"" value="""" style=""width:120""/>&nbsp;<font color=red>*</font> ��֤�룺<input type=""text"" name=""VerifyCode"" onclick=""this.value='';getCode()"" style=""width:50px""><span id='showVerify'>�����������ȡ</span></td>"
		GetWriteMessage = GetWriteMessage & "</tr>"
		GetWriteMessage = GetWriteMessage & "  <tr>"
		GetWriteMessage = GetWriteMessage & "   <td height=""30""><input type=""button"" onclick=""return(CheckForm());""  name=""SubmitComment"" value=""OK�ˣ��ύ����""/>"
		GetWriteMessage = GetWriteMessage & "    </td>"
		GetWriteMessage = GetWriteMessage & "  </tr>"
		GetWriteMessage = GetWriteMessage & "  </form>"
		GetWriteMessage = GetWriteMessage & "</table>"
		End If
		End Function 
		
		Sub GetNews()
		 Dim SQL,i,param
		 SubStr=KSBcls.Location("<div align=""left""><strong>��ҳ >> ��˾��̬</strong></div>")
		 Dim RS:Set RS=Conn.Execute("Select classid,classname from ks_userclass where username='" & UserName & "' and typeid=4 order by orderid")
		 If Not RS.Eof Then SQL=RS.GetRows(-1)
		 RS.Close:Set RS=Nothing
		 If IsArray(SQL) tHEN
		     SubStr=SubStr &"<h3><div>������鿴</div></h3><img width='50' src='images/search.png' align='absmiddle'>"
			 if ID=0 Then
			  SubStr=SubStr &"<a href='../space/?" & UserName & "/" & Action & "/'><font color=red>ȫ������</font></a>&nbsp;&nbsp;"
			 else
			  SubStr=SubStr &"<a href='../space/?" & UserName & "/" & Action & "/'>ȫ������</a>&nbsp;&nbsp;"
			 end if
			 For I=0 To Ubound(SQL,2)
			   if ID=SQL(0,I) then
			   SubStr=SubStr & "<a href='../space/?" & username & "/" & action & "/" & SQL(0,i) & "'><font color=red>" & SQL(1,i) & "(" & conn.execute("select count(id) from ks_enterprisenews where classid=" & sql(0,i))(0) &")</font></a>&nbsp;&nbsp;&nbsp;&nbsp;"
			   else
			   SubStr=SubStr & "<a href='../space/?" & username & "/" & action & "/" & SQL(0,i) & "'>" & SQL(1,i) & "(" & conn.execute("select count(id) from ks_enterprisenews where classid=" & sql(0,i))(0) &")</a>&nbsp;&nbsp;&nbsp;&nbsp;"
			   end if
			 Next
		 End If
		 if ID=0 Then
		 SubStr=SubStr &"<h3><div>��������</div></h3>"
		 Else
		 SubStr=SubStr &"<h3><div>" & Conn.Execute("Select ClassName From KS_UserClass Where ClassID=" & ID)(0) & "</div></h3>"
		 End If
		 MaxPerPage=10
		 param=" Where UserName='" & UserName & "'"
		 If ID<>0 Then Param=Param & " and classid=" & id
		 SubStr=SubStr & "<table border='0' width='98%' align='center' cellspacing='2' cellpadding='2'>"
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select ID,Title,AddDate From KS_EnterPriseNews " & Param &" order by adddate desc",conn,1,1
		 If RS.EOF and RS.Bof  Then
			 SubStr=SubStr & "<tr><td style=""border: #efefef 1px dotted;text-align:center;height:80px;"" colspan=3><p>û�з�����̬����,��<a href='../User/?user_EnterPriseNews.asp?Action=Add' target='_blank'><font color=red>��˷���</font></a>��</p></td></tr>"
		Else
			  totalPut = RS.RecordCount
              If CurrPage < 1 Then	CurrPage = 1
			
				If (totalPut Mod MaxPerPage) = 0 Then
					pagenum = totalPut \ MaxPerPage
				Else
					pagenum = totalPut \ MaxPerPage + 1
				End If
				If CurrPage>  1 and (CurrPage - 1) * MaxPerPage < totalPut Then
						RS.Move (CurrPage - 1) * MaxPerPage
				Else
						CurrPage = 1
				End If
				SQL=RS.GetRows(-1)
				 Dim K,N,Total,url
				 Total=Ubound(SQL,2)+1
				 For I=0 To Total
					If KS.SSetting(21)="1" Then Url="show-news-" & username & "-" & sql(0,n) & KS.SSetting(22) Else Url="../space/?" & username & "/shownews/" & sql(0,n)
					SubStr=SubStr &"<tr>"
					SubStr=SubStr & "<td style=""border-bottom: #efefef 1px dotted;height:22""><img src='../images/arrow_r.gif' align='absmiddle'> <a href='" & url & "' target='_blank'>" & SQL(1,N) & "</a>&nbsp;" & sql(2,n)
					SubStr=SubStr & "</td>"
					N=N+1
					If N>=Total Or N>=MaxPerPage Then Exit For
				   SubStr=SubStr &"</tr>"
				 Next
		 End If
		  SubStr=SubStr &"</table>" 
		  SubStr=SubStr & "<div id=""kspage"">" & ShowPage() & "</div>"
		  
		End Sub
		
		'��ʾ��������
		Sub ShowNews()
		 Dim SQL,i,RS,PhotoUrl,url
		 SubStr=KSBcls.Location("<div align=""left""><strong>��ҳ >> ��˾��̬ >> �鿴����</strong></div>")
		 SubStr=SubStr & "<table border='0' width='98%' align='center' cellspacing='2' cellpadding='2'>"
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select top 1 * From KS_EnterPriseNews Where UserName='" & UserName & "' and ID=" & ID  ,conn,1,1
		 If RS.EOF and RS.Bof  Then
		     RS.Close:Set RS=Nothing
			 KS.Die "<script>alert('�������ݳ���');window.close();</script>"
		 Else
		   KSBCls.Title=rs("Title")
		   SubStr=SubStr &"<tr><td align='center' style='color:#ff6600;font-weight:bold;font-size:14px'><div style=""font-weight:bold;text-align:center"">" & rs("title") & "</div></td></tr>"
		   SubStr=SubStr & "<tr><td><div style=""text-align:center"">���ߣ�" & UserName & "&nbsp;&nbsp;&nbsp;&nbsp;ʱ��:" & RS("AddDate") & "</div>"
		   SubStr=SubStr & "<hr size=1><div>" & KS.HTMLCode(rs("content")) & "</div></td></tr>"
		   If KS.SSetting(21)="1" Then Url="news-" & username  Else Url="../space/?" & username & "/news"
		   SubStr=SubStr &"<tr><td><div style='text-align:center'><a href='" & Url & "'>[���ع�˾��̬]</a></div></td></tr>"
		 End If
		 SubStr=SubStr &"</table>"   
         RS.Close:Set RS=Nothing
		End Sub
		
		'��Ʒ�б�
		Function ProductList()
		 Dim SQL,i,param,classUrl
		 SubStr=KSBcls.Location("<div align=""left""><strong>��ҳ >> ��Ʒչʾ</strong></div>")
		 Dim RS:Set RS=Conn.Execute("Select classid,classname from ks_userclass where username='" & UserName & "' and typeid=3 order by orderid")
		 If Not RS.Eof Then SQL=RS.GetRows(-1)
		 RS.Close:Set RS=Nothing
		 If IsArray(SQL) tHEN
		     SubStr=SubStr &"<h3><div>������鿴</div></h3><img width='50' src='images/search.png' align='absmiddle'>"
			 If KS.SSetting(21)="1" Then classUrl="product-" & username Else classUrl="../space/?" & UserName & "/product"
			 if ID=0 Then
			  SubStr=SubStr & "<a href='" & classUrl & "'><font color=red>ȫ����Ʒ</font></a>&nbsp;&nbsp;"
			 else
			  SubStr=SubStr &"<a href='" & classUrl & "'>ȫ����Ʒ</a>&nbsp;&nbsp;"
			 end if
			 For I=0 To Ubound(SQL,2)
			   If KS.SSetting(21)="1" Then classUrl="product-" & username & "-" & SQL(0,I) & ks.SSetting(22) Else classUrl="../space/?" & UserName & "/product/" & SQL(0,i)
			   if ID=SQL(0,I) then
			   SubStr=SubStr & "<a href='" & ClassURL & "'><font color=red>" & SQL(1,i) & "(" & conn.execute("select count(id) from ks_product where verific=1 and classid=" & sql(0,i))(0) &")</font></a>&nbsp;&nbsp;&nbsp;&nbsp;"
			   else
			   SubStr=SubStr & "<a href='" & ClassURL & "'>" & SQL(1,i) & "(" & conn.execute("select count(id) from ks_product where verific=1 and classid=" & sql(0,i))(0) &")</a>&nbsp;&nbsp;&nbsp;&nbsp;"
			   end if
			 Next
		 End If
		 if ID=0 Then
		 SubStr=SubStr &"<h3><div>���в�Ʒ</div></h3>"
		 Else
		 SubStr=SubStr &"<h3><div>" & Conn.Execute("Select classname from ks_userclass where classid=" &ID)(0) & "</div></h3>"
		 End If
		 MaxPerpage=12
		 param=" Where verific=1 and Inputer='" & UserName & "'"
		 If ID<>0 Then Param=Param & " and classid=" & id
		 SubStr=SubStr & "<table border='0' width='98%' align='center' cellspacing='2' cellpadding='2'>"
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select ID,Title,PhotoUrl From KS_Product " & Param &" order by adddate desc",conn,1,1
		 If RS.EOF and RS.Bof  Then
			SubStr=SubStr & "<tr><td style=""border: #efefef 1px dotted;text-align:center;height:80px;"" colspan=3><p>û�з�����Ʒչʾ,��<a href='../User/?user_myshop.asp?Action=Add' target='_blank'><font color=red>��˷���</font></a>��</p></td></tr>"
		Else
			  totalPut = RS.RecordCount
              If CurrPage < 1 Then	CurrPage = 1
			
				If (totalPut Mod MaxPerPage) = 0 Then
					pagenum = totalPut \ MaxPerPage
				Else
					pagenum = totalPut \ MaxPerPage + 1
				End If
				If CurrPage> 1 and (CurrPage - 1) * MaxPerPage < totalPut Then
						RS.Move (CurrPage - 1) * MaxPerPage
				End If
				SQL=RS.GetRows(MaxperPage)
				Dim K,N,Total,PhotoUrl,Url
				 Total=Ubound(SQL,2)+1
				 For I=0 To Total
				   SubStr=SubStr &"<tr>"
				   For K=1 To 4
					PhotoUrl=SQL(2,N)
					If KS.SSetting(21)="1" Then Url="show-product-" &username & "-" & sql(0,n) & KS.SSetting(22) Else url="../space/?" & UserName & "/showproduct/" & sql(0,n)
					If PhotoUrl="" Or IsNull(PhotoUrl) Then PhotoUrl="../images/nophoto.gif"
					SubStr=SubStr & "<td align='center'>" 
					SubStr=SubStr & "<a href='" & Url & "' target='_blank'><Img border='0' src='" & PhotoUrl & "' alt='" & SQL(1,N) & "' width='130' height='90' /></a><div style='text-align:center'><a href='" & Url & "'>" & KS.Gottopic(SQL(1,N),20) & "</a></div>"
					SubStr=SubStr & "</td>"
					N=N+1
					If N>=Total Or N>=MaxPerPage Then Exit For
				   Next
				   SubStr=SubStr &"</tr>"
				   If N>=Total  Or N>=MaxPerPage Then Exit For
				 Next
		 End If
		 SubStr=SubStr &"</table>" 
		 SubStr=SubStr & ShowPage()  
		End Function
		
		'�鿴��Ʒ����
		Function ShowProduct()
		 Dim SQL,i,RS,PhotoUrl
		 SubStr=KSBcls.Location("<div align=""left""><strong>��ҳ >> ��Ʒչʾ >> �鿴��Ʒ����</strong></div>")
		 SubStr=SubStr & "<table border='0' width='98%' align='center' cellspacing='2' cellpadding='2'>"
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select * From KS_Product Where inputer='" & UserName & "' and ID=" & ID ,conn,1,1
		 If RS.EOF and RS.Bof  Then
		     RS.Close:Set RS=Nothing
			 KS.Die "<script>alert('�������ݳ���');window.close();</script>"
		 Else
		   KSBCls.Title=RS("title")
		   photourl=RS("BigPhoto")
		   If PhotoUrl="" Or IsNull(photourl) Then photourl="../images/nophoto.gif"
		   SubStr=SubStr &"<tr><td align='center' style='color:#ff6600;font-weight:bold;font-size:14px'>" & rs("Title") & "</td></tr>"
		   SubStr=SubStr & "<tr><td align='center'><img src='" & photourl &"' border='0'></td></tr>"
		   SubStr=SubStr & "<tr><td><h3><div>��������</div></h3></td></tr>"
		   SubStr=SubStr & "<tr><td>�� �� �̣�" & RS("ProducerName") & "</td></tr>"
		   SubStr=SubStr & "<tr><td>��Ʒ���ࣺ" & KS.C_C(RS("tid"),1) & "</td></tr>"
		   SubStr=SubStr & "<tr><td>��Ʒ�ͺţ�" & RS("ProModel") & "</td></tr>"
		   SubStr=SubStr & "<tr><td>Ʒ��/�̱꣺" & RS("TrademarkName") & "</td></tr>"
		   SubStr=SubStr & "<tr><td>�� �� �̣�" & RS("ProducerName") & "</td></tr>"
		   SubStr=SubStr & "<tr><td>�� �� �ۣ���" & RS("price_market") & " Ԫ</td></tr>"
		   SubStr=SubStr & "<tr><td>�� Ա �ۣ���" & RS("price_member") & " Ԫ</td></tr>"
		   SubStr=SubStr & "<tr><td><h3><div>��ϸ����</div></h3></td></tr>"
		   SubStr=SubStr & "<tr><td>" & KS.HtmlCode(RS("proIntro")) & "</td></tr>"
		 End If
		 SubStr=SubStr &"</table>"   
         RS.Close:Set RS=Nothing
		End Function
		
		'��Ƹ
		Sub JobList()
		   SubStr=KSBcls.Location("<div align=""left""><strong>��ҳ >> ��ҵ��Ƹ</strong></div>")
		 If KS.C_S(10,21)="0" Then 
		   Dim Jrs:set Jrs=Conn.Execute("Select Job From ks_Enterprise where username='" & UserName & "'")
		   If Not Jrs.Eof Then
		    SubStr=SubStr & KS.HTMLCode(Jrs(0))
		   Else
		    Jrs.Close: Set Jrs=Nothing
		    KS.AlertHintScript "�Բ��𣬸��û�������ҵ�û���"
		   End If
		   Jrs.Close
		   Set Jrs=Nothing
		   Exit Sub
		 End If
		 
		 SubStr=SubStr &"<h3><div>��Ƹ��Ϣ</div></h3>"
		 MaxPerPage=5
		 Dim Param,rs,sql
		 param=" Where status=1 and UserName='" & UserName & "'"
		 SubStr=SubStr & "<table border='0' width='98%' align='center' cellspacing='2' cellpadding='2'>"
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select ID,JobTitle,province,city,workexperience,num,salary,refreshtime,status,intro,sex From KS_Job_ZW " & Param &" order by refreshtime desc",conn,1,1
		 If RS.EOF and RS.Bof  Then
			 SubStr=SubStr & "<tr><td style=""border: #efefef 1px dotted;text-align:center;height:80px;"" colspan=3><p>û�з�����Ƹ��Ϣ,��<a href='../User/User_JobCompanyZW.asp?Action=Add' target='_blank'><font color=red>��˷���</font></a>��</p></td></tr>"
		Else
			  totalPut = RS.RecordCount
              If CurrPage < 1 Then	CurrPage = 1
			
				If (totalPut Mod MaxPerPage) = 0 Then
					pagenum = totalPut \ MaxPerPage
				Else
					pagenum = totalPut \ MaxPerPage + 1
				End If
				If CurrPage>  1 and (CurrPage - 1) * MaxPerPage < totalPut Then
						RS.Move (CurrPage - 1) * MaxPerPage
				Else
						CurrPage = 1
				End If
				SQL=RS.GetRows(MaxperPage)
				Dim I,K,N,Total,PhotoUrl,url
				 Total=Ubound(SQL,2)
				 For I=0 To Total
				     SubStr=SubStr &"<tr><td style='line-height:180%;padding-top:6px;padding-bottom:8px;border-bottom:1px solid #cccccc;'>"
					 SubStr=SubStr & "<font color=#ff6600>��λ���ƣ�" & sql(1,i) & "</font>&nbsp;&nbsp;<a href='../job/job_read.asp?id=" & SQL(0,I) & "' target='_blank'>�������</a><br>�����ص㣺" & SQL(2,I) & "&nbsp;" & SQL(3,I) & "&nbsp;&nbsp;��Ƹ������" & SQL(5,I) & " ��<BR>"
					 SubStr=SubStr& "�������ڣ�" & sql(7,i) & "&nbsp;&nbsp;�Ա�Ҫ��" & SQL(10,I) & "<br>��ϸ���ܣ�" & SQL(9,I) & "</td>"
				     SubStr=SubStr &"</tr>"
				 Next
		 End If
		 SubStr=SubStr &"</table>"
		 SubStr=SubStr & ShowPage
		End Sub
		
		'����֤��
		Sub GetRyzs()
		Dim SQL,i,param,RS
		 Substr=KSBcls.Location("<div align=""left""><strong>��ҳ >> ����֤��</strong></div>")
		 SubStr=SubStr &"<h3><div>����֤��</div></h3>"

		 param=" Where status=1 and UserName='" & UserName & "'"
		 SubStr=SubStr & "<table style='margin-bottom:5px' border='0' width='98%' align='center' cellspacing='1' cellpadding='0' bgcolor='#FFFFFF'>"
		 SubStr=SubStr & "<tr bgcolor='#F3F3F3' align='center'><td width='20%' height='20'>֤����Ƭ</td><td width='24%'>֤������</td><td width='21%'>��֤����</td><td width='17%'>��Ч����</td><td width='18%'>��ֹ����</td></tr>"
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select ID,Title,FZJG,sxrq,jzrq,photourl From KS_EnterPriseZS " & Param &" order by adddate desc",conn,1,1
		 If RS.EOF and RS.Bof  Then
			SubStr=SubStr & "<tr><td style=""border: #efefef 1px dotted;text-align:center;height:80px;"" colspan=6><p>û�з�������֤��,��<a href='../User/?user_EnterPriseZS.asp?Action=Add' target='_blank'><font color=red>��˷���</font></a>��</p></td></tr>"
		Else
			  totalPut = RS.RecordCount
              If CurrPage < 1 Then	CurrPage = 1
			
				If (totalPut Mod MaxPerPage) = 0 Then
					pagenum = totalPut \ MaxPerPage
				Else
					pagenum = totalPut \ MaxPerPage + 1
				End If
				If CurrPage>  1 and (CurrPage - 1) * MaxPerPage < totalPut Then
						RS.Move (CurrPage - 1) * MaxPerPage
				Else
						CurrPage = 1
				End If
				SQL=RS.GetRows(MaxperPage)
				Dim K,N,Total,PhotoUrl,url,BeginDateStr,EndDateStr
		 Total=Ubound(SQL,2)
		 For I=0 To Total
		   if i mod 2=0 then
		    SubStr=SubStr &"<tr bgcolor='#ffffff'>"
		   else
		    SubStr=SubStr & "<tr bgcolor='#f6f6f6'>"
		   end if
		    PhotoUrl=SQL(5,i)
			If PhotoUrl="" Or IsNull(PhotoUrl) Then PhotoUrl="../images/nophoto.gif"
			BeginDateStr=SQL(3,I) :	If Not IsDate(BeginDateStr) Then BeginDateStr=Now
			EndDateStr =SQL(4,I) : If Not IsDate(EndDateStr) Then EndDateStr=Now
		    SubStr=SubStr & "<td width='150' style='height:80px;text-align:center;padding-top:6px;padding-bottom:8px;'>" 
			SubStr=SubStr & "<a href='" & PhotoUrl & "' target='_blank'><Img border='0' src='" & PhotoUrl & "' width='85' height='60'></a>"
			SubStr=SubStr & "</td>"
			SubStr=SubStr & "<td style='text-align:center;line-height:150%;' >" & sql(1,i) & "</td>"
			SubStr=SubStr & "<td style='text-align:center;line-height:150%;' >" & sql(2,i) & "</td>"
			SubStr=SubStr & "<td style='text-align:center;line-height:150%;' >" & year(BeginDateStr) & "��" & month(BeginDateStr) & "��</td>"
			SubStr=SubStr & "<td style='text-align:center;line-height:150%;' >" & year(EndDateStr) & "��" & month(EndDateStr) & "��</td>"
		    SubStr=SubStr &"</tr>"
		 Next
		 End If
		 SubStr=SubStr &"</table>" 
		 SubStr=SubStr & ShowPage  
		End Sub
		
		'��ʾͼƬ
		Function ShowPhoto()
		 Dim SQL,n,RS,PhotoUrlArr,PhotoUrl,t
		 substr=KSBcls.Location("<div align=""left""><strong>��ҳ >> ��Ʒչʾ >> �鿴��Ʒ</strong></div>")
		 substr=substr & "<table border='0' width='98%' align='center' cellspacing='2' cellpadding='2'>"
		 Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open "Select * From KS_Photo Where Inputer='" & UserName & "' and ID=" & ID  ,conn,1,1
		 If RS.EOF and RS.Bof  Then
		     RS.Close:Set RS=Nothing
			 KS.Die "<script>alert('�������ݳ���');window.close();</script>"
		 Else
		   KSBCls.Title = rs("title")
		   photourlArr=Split(RS("PicUrls"),"|||")
		   n=CurrPage
		   if n<0 then n=0
		   t=Ubound(PhotoUrlArr)
		   If N>=t Then n=0
		   If t=0 Then t=1
		   PhotoUrl=Split(PhotoUrlArr(N),"|")(1)
		   substr=substr & "<tr><td align='center' class='divcenter_work_on'><div class='fpic'><a href='../space/?" & UserName & "/showphoto/" & ID & "/" & n+1 &"'><img  onload=""var myImg = document.getElementById('myImg'); if (myImg.width >580 ) {myImg.width =580 ;};"" id=""myImg"" src='" & photourl &"' title='�鿴��һ��' border='0'></A></div></td></tr>"
		   substr=substr &"<tr><td height='35' align='center'>�����<Script Src='../item/GetHits.asp?Action=Count&m=2&GetFlag=0&ID=" & ID & "'></Script> �ܵ�Ʊ��<Script Src='../item/GetVote.asp?m=2&ID=" & ID & "'></Script> ͶƱ��<a href='../item/Vote.asp?m=2&ID=" & ID & "'>Ͷ��һƱ</a></td></tr>"
           substr=substr & "<tr><td height='35' align='center'>��" & N+1 & "/" & t & "�� <a href='../space/?" & UserName & "/showphoto/" & ID &"/0'><img src='images/picindex.gif' border='0'></a>&nbsp;<a href='../space/?" & UserName & "/showphoto/" & id &"/" & N-1 & "'><img src='images/picpre.gif' border='0'></a>&nbsp;<a href='../space/?" & UserName & "/showphoto/" & id &"/" & N+1 & "'><img src='images/picnext.gif' border='0'></a>&nbsp;<a href='../space/?" & UserName & "/showphoto/" & id &"/" & t-1 & "'><img src='images/picend.gif' border='0'></a></td></tr>"
		   substr=substr & "<tr><td><span class=""writecomment""><Script Language=""Javascript"" Src=""../plus/Comment.asp?Action=Write&ChannelID=2&InfoID=" &id & """></Script></span></td></tr>"
		   substr=substr & "<tr><td>&nbsp;<Img src='images/topic.gif' align='absmiddle'> <strong>��Ʒ���ۣ�</strong><br><span class=""showcomment""><script src=""../ks_inc/Comment.page.js"" language=""javascript""></script><script language=""javascript"" defer>Page(1,2,'" & ID & "','Show','../');</script><div id=""c_" & ID & """></div><div id=""p_" & ID & """ align=""right""></div> </span></td></tr>"
		 End If
		 substr=substr &"</table>"   
		End Function
		
		
		
		'ͨ�÷�ҳ
		Public Function ShowPage()
		         Dim I, PageStr
				 PageStr = ("<div class=""fenye""><table border='0' align='right'><tr><td><div class='showpage' style='height:20px'>")
					if (CurrPage>1) then pageStr=PageStr & "<a href=""../space/?" & username & "/" &action & "/" & ID & "/" & CurrPage-1 & """ class=""prev"">��һҳ</a>"
				   if (CurrPage<>PageNum) then pageStr=PageStr & "<a href=""../space/?" & username & "/" &action & "/" & ID & "/" & CurrPage+1 & """ class=""next"">��һҳ</a>"
				   pageStr=pageStr & "<a href=""../space/?" & username & "/" &action & """ class=""prev"">�� ҳ</a>"
				 
					Dim startpage,n,j
					 if (CurrPage>=7) then startpage=CurrPage-5
					 if PageNum-CurrPage<5 Then startpage=PageNum-10
					 If startpage<0 Then startpage=1
					 n=0
					 For J=startpage To PageNum
						If J= CurrPage Then
						 PageStr=PageStr & " <a href=""#"" class=""curr""><font color=red>" & J &"</font></a>"
						Else
						 PageStr=PageStr & " <a class=""num"" href=""../space/?" & username & "/" &action & "/" &id & "/" & J&""">" & J &"</a>"
						End If
						n=n+1 : if n>=10 then exit for
					 Next
					
					 PageStr=PageStr & " <a class=""next"" href=""../space/?" & username & "/" &action & "/" &id & "/" & PageNum&""">ĩҳ</a>"
					 pageStr=PageStr & " <span>��" & totalPut & "����¼,��" & PageNum & "ҳ</span></td></tr></table>"
				     PageStr = PageStr & "</div>"
			         ShowPage = PageStr
	     End Function
End Class
%>
