<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../plus/md5.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
response.cachecontrol="no-cache"
response.addHeader "pragma","no-cache"
response.expires=-1
response.expiresAbsolute=now-1
Response.CharSet="gb2312"

Dim KS:Set KS=New PublicCls
Set KSUser = New UserCls
Call KSUser.UserLoginChecked()
Dim ChannelID,InfoID,RS,CommentStr,UserIP,Total,TitleStr,TitleLinkStr,TotalPoint,N,DomainStr
Dim totalPut, CurrentPage, MaxPerPage,PageNum,SqlStr,PrintOut,CommentXML
ChannelID=KS.Chkclng(KS.S("ChannelID"))
IF ChannelID=0 And KS.S("Action")<>"Support" And KS.S("Action")<>"QuoteSave" Then KS.Die ""
PrintOut=KS.S("PrintOut")

InfoID=KS.ChkClng(KS.S("InfoID"))
DomainStr=KS.GetDomain
Select Case KS.S("Action")
 Case "Show"  Call Show()
 Case "Write"
  If KS.C_S(ChannelID,12)=0 Then Response.end()
  Call Ajax()
  Response.Write("document.write('" & GetWriteComment(ChannelID,InfoID) & "');")
 Case "WriteSave"  Call WriteSave()
 Case "Support"  
  If PrintOut="js" Then
   Response.Write "ShowSupportMessage('" & Support() & "');"
  Else
   Response.Write Support()
  End If
 Case "QuoteSave" Call QuoteSave()
 Case Else  Call CommentMain()
 End Select
 Set KS=Nothing
 Set KSUser=Nothing
 
 Sub Ajax()
 %>
function xmlhttp()
{
	if(window.XMLHttpRequest){
		return new XMLHttpRequest();
	} else if(window.ActiveXObject){
		return new ActiveXObject("Microsoft.XMLHTTP");
	} 
	throw new Error("XMLHttp object could be created.");
}
	
var loader=new xmlhttp;
function ajaxLoadPage(url,request,method,fun)
{
	method=method.toUpperCase();
	if (method=='GET')
	{
		urls=url.split("?");
		if (urls[1]=='' || typeof urls[1]=='undefined')
		{
			url=urls[0]+"?"+request;
		}
		else
		{
			url=urls[0]+"?"+urls[1]+"&"+request;
		}
		
		request=null;
	}
	loader.open(method,url,true);
	if (method=="POST")
	{
		loader.setRequestHeader("Content-Type","application/x-www-form-urlencoded");
	}
	loader.onreadystatechange=function(){
	     eval(fun+'()');
	}
	loader.send(request);
}

function formToRequestString(form_obj)
{
    var query_string='';
    var and='';
    for (var i=0;i<form_obj.length;i++ )
    {
        e=form_obj[i];
        if (e.name) {
            if (e.type=='select-one') {
                element_value=e.options[e.selectedIndex].value;
            } else if (e.type=='select-multiple') {
                for (var n=0;n<e.length;n++) {
                    var op=e.options[n];
                    if (op.selected) {
                        query_string+=and+e.name+'='+escape(op.value);
                        and="&"
                    }
                }
                continue;
            } else if (e.type=='checkbox' || e.type=='radio') {
                if (e.checked==false) {   
                    continue;   
                }   
                element_value=e.value;
            } else if (typeof e.value != 'undefined') {
                element_value=e.value;
            } else {
                continue;
            }
            query_string+=and+e.name+'='+escape(element_value);
            and="&"
        }
    }
    return query_string;
}
function ajaxFormSubmit(form_obj,fun)
{ 
	ajaxLoadPage(form_obj.getAttributeNode("action").value,formToRequestString(form_obj),form_obj.method,fun)
}
 <%
 End Sub
 
 Sub CommentMain
	Dim KSRCls,FileContent
	Set KSRCls = New Refresh
	FCls.RefreshType = "Comment" '����ˢ�����ͣ��Ա�ȡ�õ�ǰλ�õ�����

	if KS.C_S(ChannelID,15)="" then KS.Die "���ȵ�ģ�������������ҳģ��!"
	FileContent = KSRCls.LoadTemplate(KS.C_S(ChannelID,15))
	If Trim(FileContent) = "" Then FileContent = "ģ�岻����!"
	FileContent=Replace(FileContent,"{$GetShowComment}","<script src=""" & domainstr & "ks_inc/Comment.page.js"" language=""javascript""></script><script src=""" & domainstr & "ks_inc/Kesion.box.js"" language=""javascript""></script><script language=""javascript"" defer>Page(1," & ChannelID & ",'" & InfoID & "','Show','"& domainstr & "');</script><div id=""c_" & InfoID & """></div><div id=""p_" & InfoID & """ align=""right""></div>")

	if channelid<>8 then
	 if conn.execute("select count(id) from " & KS.C_S(ChannelID,2) & " Where ID=" & InfoID).eof then 
	 KS.Die "<script>alert('�Բ�����ɾ�� ��');window.close();</script>"
	end if
	if conn.execute("select comment from " & KS.C_S(ChannelID,2) & " Where ID=" & InfoID)(0)=0 then KS.Die "<script>alert('�Բ��𣬲��������� ��');window.close();</script>"
	end if
	
   TitleStr=conn.execute("Select Title From " & KS.C_S(ChannelID,2) & " Where ID=" & InfoID)(0)

  FileContent=Replace(FileContent,"{$GetTitle}",TitleStr)
  FileContent=Replace(FileContent,"{$GetWriteComment}","<script language=""javascript"" src=""?Action=Write&ChannelID=" & ChannelID& "&InfoID=" & InfoID & """></script>")
	FileContent = KSRCls.ReplaceLableFlag(KSRCls.ReplaceAllLabel(FileContent))
	FileContent = KSRCls.ReplaceGeneralLabelContent(FileContent) '�滻ͨ�ñ�ǩ
	Set KSRCls = Nothing
   Response.Write(FileContent)
End Sub

Sub Show()
	MaxPerPage=5    'ÿҳ��ʾ��������
    SqlStr="Select top 1 ID,Title,Tid,Fname From " & KS.C_S(ChannelID,2) & " Where ID=" & InfoID
  Set RS=Server.CreateObject("ADODB.RECORDSET")
 RS.Open SqlStr,Conn,1,1
 If Not RS.Eof Then
   TitleStr=RS(1):TitleLinkStr="<a href='" & KS.GetItemUrl(ChannelID,RS(2),rs(0),rs(3)) & "' target='_blank'>" & TitleStr & "</a>"
 Else
   KS.Die ""
 End If

	If KS.S("page") <> "" Then
			  CurrentPage = KS.ChkClng(KS.S("page"))
	Else
			  CurrentPage = 1
	End If
	RS.Close
RS.Open "Select  b.userface,a.* From KS_Comment a left join KS_User b on a.username=b.username Where a.Verific=1 And a.ChannelID=" & ChannelID & " And a.InfoID=" & InfoID & " Order By ID Desc",conn,1,1

  IF Not Rs.Eof Then
		 totalPut = Conn.Execute("Select Count(ID) From KS_Comment Where Verific=1 And ChannelID=" & ChannelID & " And InfoID=" & InfoID)(0)
				If CurrentPage < 1 Then	CurrentPage = 1
						If (totalPut Mod MaxPerPage) = 0 Then
									PageNum = totalPut \ MaxPerPage
						Else
									PageNum = totalPut \ MaxPerPage + 1
						End If
		
				         If CurrentPage >1 And (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
						 Else
									CurrentPage = 1
				         End If
						 Set CommentXML=KS.ArrayToxml(Rs.GetRows(MaxPerPage),Rs,"row","xml")
						 Call showContent()

  Else
	CommentStr=""
  End If
  Rs.Close:Set Rs=Nothing
  
  If KS.C_S(ChannelID,12)=0 Then TotalPut=0
  If PrintOut="js" Then
   Response.Write "show(""" & replace(replace(CommentStr,vbcrlf,"\n"),"""","\""") & "{ks:page}" & TotalPut & "|" & MaxPerPage & "|" & PageNum & "|��||2"");"
  Else
   Response.Write CommentStr & "{ks:page}" & TotalPut & "|" & MaxPerPage & "|" & PageNum & "|��||2"
  End If
End Sub

Sub ShowContent()
   If KS.C_S(ChannelID,12)=0 Then Exit Sub
	
    CommentStr="<br /> &nbsp;�����Ƕ� <strong>[" & TitleLinkStr & "]</strong> ������,�ܹ�:<font color=red>" & totalPut & " </font>������<br />"
    CommentStr=CommentStr & "<table  width='98%' border='0' align='center' cellpadding='0' cellspacing='1'>"

  If CurrentPage=1 Then	N=TotalPut	Else N=totalPut-MaxPerPage*(CurrentPage-1)
  Dim FaceStr,Publish,QuoteContentj,Content,Node,UserFace,ID,ReplyContent
  
  If IsObject(CommentXML) Then
   For Each Node In CommentXML.DocumentElement.SelectNodes("row")
		FaceStr= KS.GetDomain &  "images/face/0.gif"
		ID=Node.SelectSingleNode("@id").text
		ReplyContent=Node.SelectSingleNode("@replycontent").text
	   IF Node.SelectSingleNode("@anonymous").text="0" Then
		Publish=Node.SelectSingleNode("@username").text
		UserFace=Node.SelectSingleNode("@userface").text

		If Not KS.IsNul(UserFace) Then
			FaceStr=UserFace
			If lcase(left(FaceStr,4))<>"http" then FaceStr=KS.GetDomain & FaceStr
		End If
		Publish="��Ա:<a href=""" & KS.GetDomain & "space/?" & Publish & """ target=""_blank"">" & publish & "</a>"
	   Else
		Publish= "�οͣ�"& Node.SelectSingleNode("@anounname").text
	   End IF
	   QuoteContent=Node.SelectSingleNode("@quotecontent").text
	   If Not KS.IsNUL(QuoteContent) Then
	   QuoteContent=Replace(QuoteContent,"[quote]","<div style='margin:2px;border:1px solid #cccccc;background:#FFFFEE;padding:4px'>")
	   QuoteContent=Replace(QuoteContent,"[/quote]","</div>")
	   QuoteContent=Replace(QuoteContent,"[dt]","<div style='padding-left:10px;color:#999999'>")
	   QuoteContent=Replace(QuoteContent,"[/dt]","</div>")
	   QuoteContent=Replace(QuoteContent,"[dd]","<div style='padding-left:10px;'>")
	   QuoteContent=Replace(QuoteContent,"[/dd]","</div>")
	   End If
	   Content = KS.HtmlCode(ReplaceFace(QuoteContent & Node.SelectSingleNode("@content").text))
		
	   CommentStr=CommentStr & "<tr>"
	   CommentStr=CommentStr & "<td width='70' rowspan='3' style='margin-top:3px;BORDER-BOTTOM: #999999 1px dotted;'><img width='60' height='60' src='" & facestr & "' border='1'></td>"
	   CommentStr=CommentStr & "<td height='25' width='*'>"
	   CommentStr=CommentStr & publish
	   CommentStr=CommentStr  & " <font color='#999999'>(����ʱ�䣺 " & Node.SelectSingleNode("@adddate").text &")</font> </td><td width='30'><font style='font-size:32px;font-family:Arial Black;color:#EEF0EE'> " & N & "</font> </td>"
	   CommentStr=CommentStr & "</tr>"
	   CommentStr=CommentStr & "<tr><td height='25' colspan='2' style='word-break:break-all;'>" & Content
	   If ReplyContent<>"" Then
	   CommentStr=CommentStr & "<div style='padding:4px;color:red;border:1px solid #ccc;background:#FFFFEE;'>""" & Node.SelectSingleNode("@replyuser").text & """�ظ�:" & ReplyContent & "</div>"
	   End If
	   
	   CommentStr=CommentStr & "</td></tr>"
	   CommentStr=CommentStr & "<tr>"
	   CommentStr=CommentStr & "<td style='margin-top:3px;BORDER-BOTTOM: #999999 1px dotted;' height='25' colspan='2' style='text-align:right'><a href='javascript:void(0)' onclick=reply("& ChannelID & "," & ID & ",'" & KS.GetDomain & "');>��¥(�ظ�)</a> <a href='javascript:void(0)' onclick=javascript:Support(" & ID & ",1,'" &KS.GetDomain & "');><span style='color:brown'>֧��</span>[" & Node.SelectSingleNode("@score").text & "]</a> <a href='javascript:void(0)' onclick=javascript:Support(" & ID & ",0,'" & KS.GetDomain & "');return false>����[" & Node.SelectSingleNode("@oscore").text & "]</a> </td>"
	   CommentStr=CommentStr & "</tr>"
	   N=N-1
   Next
 End If
   CommentStr=CommentStr & "</table>"

End Sub
 
 '��������
Function GetWriteComment(ChannelID,InfoID)
%>
	function insertface(Val)
	{ 
	  <%If KS.C_S(ChannelID,14)<>0  Then%>
	  checklength(document.getElementById('C_Content'));
	  <%end if%>
	  if (Val!=''){
	  var ubb=document.getElementById("C_Content");
		var ubbLength=ubb.value.length;
		ubb.focus();
		if(typeof document.selection !="undefined")
		{
			document.selection.createRange().text=Val;  
		}
		else
		{
			ubb.value=ubb.value.substr(0,ubb.selectionStart)+Val+ubb.value.substring(ubb.selectionStart,ubbLength);
		}
     }
  }
	
function success()
{
	var loading_msg='\n\n\t���Եȣ������ύ����...';
	var C_Content=document.getElementById('C_Content');
	
 	if (loader.readyState==1)
		{
			C_Content.value=loading_msg;
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
			  C_Content.value=document.getElementById('sC_Content').value;
			 }
		}
}
	var OutTimes =11;
	function leavePage()
	{
	if (OutTimes==0)
	 {
	 document.getElementById('C_Content').disabled=false;
	 document.getElementById('SubmitComment').disabled=false;
	 document.getElementById('C_Content').value=''
	 <%If KS.C_S(ChannelID,13)="1" Then%>
	  document.getElementById('VerifyCode').value='';
	  document.getElementById('verifyimg').src=document.getElementById('verifyimg').src;
	 <%end if%>
	 <%If KS.C_S(ChannelID,14)<>0  Then%>
	 document.getElementById('cmax').value=<%=KS.C_S(ChannelID,14)%>;
	 <%end if%>
	 OutTimes =11;
	 return;
	 }
	else {
	    document.getElementById('C_Content').disabled=true;
		document.getElementById('SubmitComment').disabled=true;
		OutTimes -= 1;
		document.getElementById('C_Content').value ="\n\n\t�������ύ���ȴ� "+ OutTimes + " ���Ӻ����ɼ�������...";
		setTimeout("leavePage()", 1000);
		}
	}
function checklength(cobj)
{ 
	var cmax=<%=KS.C_S(ChannelID,14)%>;
	if (cobj.value.length>cmax) {
	cobj.value = cobj.value.substring(0,cmax);
	alert("���۲��ܳ���"+cmax+"���ַ�!");
	}
	else {
	document.getElementById('cmax').value = cmax-cobj.value.length;
	}
}

   function checkform()
   {
	var anounname=document.getElementById('AnounName');
	var C_Content=document.getElementById('C_Content');
	var sC_Content=document.getElementById('sC_Content');
	var anonymous=document.getElementById('Anonymous');
	var pass=document.getElementById('Pass');
   if (anounname.value==''){
        alert('����д�û�����');
		anounname.focus();
        return false;
     }
	if (anonymous.checked==false && pass.value==''){
	   alert('�����������ѡ���οͷ���');
	   pass.focus();
	   return false;
	}
	<%If KS.C_S(ChannelID,13)="1" Then%>
   if (document.getElementById('VerifyCode').value==''){
	   alert('������֤��!');
	   document.getElementById('VerifyCode').focus();
	   return false;
    }
	<%end if%>
   if (C_Content.value==''){
	   alert('����д��������!');
	   C_Content.focus();
	   return false;
    }
	sC_Content.value=C_Content.value;
	try{ajaxFormSubmit(document.form1,'success');
	 }catch(e){
	  document.form1.action="<%=DomainStr%>plus/Comment.asp?Action=WriteSave&flag=NotAjax";
	  document.form1.submit();
	 }
	 
	 
	}
<%
         Dim k
		 GetWriteComment = "<table width=""100%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" class=""comment_write_table"">"
		 GetWriteComment = GetWriteComment & "<form name=""form1"" action=""" & DomainStr &"plus/Comment.asp?Action=WriteSave"" method=""post"">"
		 GetWriteComment = GetWriteComment & "<input type=""hidden"" value=""" & ChannelID & """ name=""ChannelID""><input type=""hidden"" value=""" & InfoID & """ name=""InfoID"">"
		GetWriteComment = GetWriteComment & "<tr>"
		GetWriteComment = GetWriteComment & "  <td style=""padding:10px;"">"
		GetWriteComment = GetWriteComment & "  <div style=""height:30px;line-height:30px;text-align:left;""><strong>����˵����</strong>&nbsp;&nbsp;<font color=#ff6600>�ر���������������ֻ�������Ѹ��˹۵㣬�뱾վ�����޹ء�</font></div>"
		GetWriteComment = GetWriteComment & "  <div style=""text-align:left;"">�û�����"
		If KSUser.UserName="" Then
		GetWriteComment = GetWriteComment & "   <input class=""textbox"" maxlength=15 name=""AnounName"" type=""text"" id=""AnounName"" value=""����"" style=""width:12%""/>"
		Else
		GetWriteComment = GetWriteComment & "   <input class=""textbox"" maxlength=15 name=""AnounName"" type=""text"" id=""AnounName"" value=""" & KSUser.username & """ style=""width:12%""/>"
		End If
		Dim Style,Check
		If KS.C_S(ChannelID,12)=1 or KS.C_S(ChannelID,12)=2 Then
		 style="":checked=""
		Else
		 Style=" style=""display:none""":checked=" checked"
		End If
		If KS.C("UserName")="" Then checked=" checked" else checked=""
		
		GetWriteComment = GetWriteComment & "<span id=""pp""" & style & "> ���룺<input class=""textbox"" name=""Pass"" size=""10"" type=""password"" id=""Pass"" value=""" & KSUser.PassWord & """ class=""denglu""></span>"

		If KS.C_S(ChannelID,13)="1" Then
		GetWriteComment = GetWriteComment & "&nbsp;��֤�룺<input id=""VerifyCode"" name=""VerifyCode"" type=""text"" size=""4""><img style=""cursor:pointer"" src=""" & DomainStr & "plus/verifycode.asp"" onClick=""this.src=\'" & DomainStr & "plus/verifycode.asp?n=\'+ Math.random();"" id=""verifyimg"" align=""absmiddle"">"
		End IF
		If KS.C_S(Channelid,12)=1 Or KS.C_S(Channelid,12)=2 Then
		GetWriteComment = GetWriteComment & "<span style=""display:none"">"
		Else
		GetWriteComment = GetWriteComment & "<span>"
		End iF
		GetWriteComment = GetWriteComment & "<label><input onclick=""if(this.checked==true){document.getElementById(\'Pass\').disabled=true;document.getElementById(\'pp\').style.display=\'none\';}else{document.getElementById(\'Pass\').disabled=false;document.getElementById(\'pp\').style.display=\'\';}"" type=""checkbox""" & checked & " value=""1"" name=""Anonymous"" id=""Anonymous"">��������</label></span>"
		
		GetWriteComment = GetWriteComment & "&nbsp;&nbsp;<a href=""" & DomainStr & "User/reg/""><font color=red><u>ע��</u><font></a>"
		GetWriteComment = GetWriteComment & "  </div>"
		
		GetWriteComment = GetWriteComment & "  <div style=""width:98%;background:#f8f8f8;margin-top:10px;border:1px solid #d8d8d8;height:26px;padding-top:5px;border-bottom:none;"">"
		
		 Dim str:str="����|Ʋ��|ɫ|����|����|����|����|����|˯|���|����|��ŭ|��Ƥ|����|΢Ц|�ѹ�|��|�ǵ�|ץ��|��|"
		 Dim strArr:strArr=Split(str,"|")
		 For K=0 to 19
		   GetWriteComment = GetWriteComment & "<img style=""cursor:pointer"" title=""" & strarr(k) & """ onclick=""insertface(\'[e" & k &"]\')""  src=""" & DomainStr & "images/emot/" & K & ".gif"">"
		   If (K+1) mod 5=0 Then GetWriteComment = GetWriteComment 
		 Next
	 
		GetWriteComment = GetWriteComment & "</div><div>"
		If KS.C_S(ChannelID,14)<>0  Then
		GetWriteComment = GetWriteComment & "<textarea onkeydown=""checklength(this);"" onkeyup=""checklength(this);"" name=""C_Content"" rows=""6"" id=""C_Content"" style=""width:100%""></textarea>"
		Else
		GetWriteComment = GetWriteComment & "<textarea style=""border:1px solid #cccccc;width:98%;height:90px;border-right:none;"" wrap=""PHYSICAL"" name=""C_Content"" rows=""4"" id=""C_Content""></textarea>"
		End If
		
		GetWriteComment = GetWriteComment & "</div></td>"

		GetWriteComment = GetWriteComment & "  </tr>"
		GetWriteComment = GetWriteComment & "  <tr>"
		GetWriteComment = GetWriteComment & "    <td height=""25"" align=""center"">"
		If KS.C_S(ChannelID,14)<>0  Then
		GetWriteComment = GetWriteComment & "ʣ��������<input disabled type=""text"" id=""cmax"" size=""5"" name=""cmax"" value=""" & KS.C_S(ChannelID,14) & """>&nbsp;&nbsp;&nbsp;"
		End If
		
		GetWriteComment = GetWriteComment & "<input type=""hidden"" name=""sC_Content"" id=""sC_Content""><input type=""image"" id=""SubmitComment"" name=""SubmitComment"" src=""" & DomainStr &"images/comment.gif"" onclick=""checkform();return false""/>"
		
		GetWriteComment = GetWriteComment & "    <a href=""" & DomainStr &"plus/Comment.asp?ChannelID=" & ChannelID & "&InfoID=" & InfoID & """ target=""_blank""><img src=""" & DomainStr &"images/commentimg.gif""></a></td>"
		GetWriteComment = GetWriteComment & "  </tr>"
		GetWriteComment = GetWriteComment & "  </form>"
		GetWriteComment = GetWriteComment & "</table>"
		
		End Function  
		
		Function ReplaceFace(c)
		 Dim str:str="����|Ʋ��|ɫ|����|����|����|����|����|˯|���|����|��ŭ|��Ƥ|����|΢Ц|�ѹ�|��|�ǵ�|ץ��|��|"
		 Dim strArr:strArr=Split(str,"|")
		 Dim K
		 For K=0 To 19
		  c=replace(c,"[e"&K &"]","<img title='" & strarr(k) & "' src='" & KS.Setting(2) & "/images/emot/" & K & ".gif'>")
		 Next
		 C=KS.FilterIllegalChar(C)
		 ReplaceFace=C
		End Function
		
		'���淢��
		Sub WriteSave()	
		Dim UserName,Email,C_Content,Verific,Anonymous,point,VerifyCode,Pass,Flag,ComeUrl,GroupID
		Flag=KS.S("Flag")
		ComeUrl=Request.ServerVariables("HTTP_REFERER")
		If ComeUrl="" Then ComeUrl=KS.GetDomin
		If KS.C_S(Channelid,12)=0 Then 
		 If Flag="NotAjax" Then
		  Response.Write "<script>alert('�Բ���,����Ϣ����������');location.href='" & ComeUrl & "';</script>"
		 Else
		  Response.Write "�Բ���,����Ϣ���������ۣ�"
		 End If
		End If	  

		AnounName=KS.R(KS.S("AnounName"))
		If Len(AnounName)>15 Then
		 If Flag="NotAjax" Then
		  Response.Write "<script>alert('�û���̫��');location.href='" & ComeUrl & "';</script>"
		 Else
		  Response.Write "�û���̫����"
		 End If
		 Response.End()
		End If
		Pass=KS.R(KS.G("Pass"))
		Email=KS.S("Email")
		C_Content=KS.S("C_Content")
		VerifyCode=KS.S("VerifyCode")
		
		Anonymous=KS.ChkClng(KS.S("Anonymous"))
		point=KS.ChkClng(KS.S("point"))
		If KS.C_S(ChannelID,13)="1" and Trim(Request.Form("Verifycode"))<>Trim(Session("Verifycode")) Then
		 If Flag="NotAjax" Then
		  Response.Write "<script>alert('��֤����������������!');history.back();</script>"
		 Else
		 Response.Write("��֤���������������룡")
		 Response.End
		 End If
		End IF
		  
		if KS.C_S(Channelid,12)=1 Or KS.C_S(ChannelID,12)=2 then
		  if Cbool(KSUser.UserLoginChecked)=false  then
				  If Flag="NotAjax" Then
				   Response.Write "<script>alert('�Բ���ϵͳ���ò������οͷ���');history.back();</script>"
				  Else
				   Response.Write("�Բ���ϵͳ���ò������οͷ���")
				   Response.End
				  End If
		  End If
		End If
		
		IF Anonymous=0 Then
		  if Cbool(KSUser.UserLoginChecked)=false  then
		     	if Pass="" Then 
				  If Flag="NotAjax" Then
				   Response.Write "<script>alert('����д��¼�����ѡ���οͷ���');history.back();</script>"
				  Else
				   Response.Write("����д��¼�����ѡ���οͷ���")
				   Response.End
				  End If
				End if
             Pass=Md5(Pass,16)
		     Dim UserRS:Set UserRS=Server.CreateObject("Adodb.RecordSet")
			 UserRS.Open "Select top 1 UserID,UserName,PassWord,Locked,Score,LastLoginIP,LastLoginTime,LoginTimes,RndPassword,GroupID From KS_User Where UserName='" &AnounName & "' And PassWord='" & Pass & "'",Conn,1,3
			 If UserRS.Eof And UserRS.BOf Then
				  If Flag="NotAjax" Then
				   Response.Write "<script>alert('��������û�����������������������!');history.back();</script>"
				  Else
				   Response.Write("��������û�����������������������!")
				  End If
				  UserRS.Close:Set UserRS=Nothing
				 response.end
			 ElseIf UserRS("Locked")=1 Then
				  If Flag="NotAjax" Then
				   Response.Write "<script>alert('�����˺��ѱ�����Ա�������������Ա��ϵ!');history.back();</script>"
				  Else
			       Response.Write("�����˺��ѱ�����Ա�������������Ա��ϵ!")
				  End If
			   response.end
			 Else
			            GroupID=UserRS("GroupID")
			            '��¼�ɹ��������û���Ӧ������
						Dim RndPassword:RndPassword=KS.R(KS.MakeRandomChar(20))
						If datediff("n",UserRS("LastLoginTime"),now)>=KS.Setting(36) then '�ж�ʱ��
						UserRS("Score")=UserRS("Score")+KS.Setting(37)
						end if
						UserRS("LastLoginIP") = KS.GetIP
                        UserRS("LastLoginTime") = Now()
                        UserRS("LoginTimes") = UserRS("LoginTimes") + 1
						UserRS("RndPassWord")=RndPassWord
                        UserRS.Update
						Response.Cookies(KS.SiteSn)("UserName") = AnounName
						Response.Cookies(KS.SiteSn)("Password") = Pass
						Response.Cookies(KS.SiteSN)("RndPassword")= RndPassword
			end if
		  Else
		     groupid=KSUser.GroupID
		  end if
		Else
		    Dim RSG:Set RSG=Conn.Execute("select top 1 groupid from KS_User Where UserName='" & AnounName & "'")
			If Not RSG.Eof Then
			  groupID=rsg(0)
			End If
			RSG.Close : Set RSG=Nothing
		End IF


		IF InfoID="" Then 
			 If Flag="NotAjax" Then
			  Response.Write "<script>alert('������������!');history.back();</script>"
			 Else
		      Response.Write("������������!")
			 End If
		     Response.End
		End if
		if AnounName="" Then
			 If Flag="NotAjax" Then
			  Response.Write "<script>alert('����д����ǳ�!');history.back();</script>"
			 Else
		      Response.Write("����д����ǳ�!")
			 End If
		    Response.End
		End if
		'if KS.IsValidEmail(Email)=false then
		' Response.Write("<script>alert('��������ȷ�ĵ�������!');history.back();<//script>")
		' Response.End
		'end if
		
		if C_Content="" Then 
			 If Flag="NotAjax" Then
			  Response.Write "<script>alert('����д��������!');history.back();</script>"
			 Else
		      Response.Write("����д��������!")
			 End If
		 Response.End
		End if
		If Len(C_Content)>KS.ChkClng(KS.C_S(ChannelID,14)) and KS.ChkClng(KS.C_S(ChannelID,14))<>0 Then
			 If Flag="NotAjax" Then
			  Response.Write "<script>alert('�������ݱ�����" &KS.C_S(ChannelID,14) & "���ַ�����!');history.back();</script>"
			 Else
		      Response.Write("�������ݱ�����" &KS.C_S(ChannelID,14) & "���ַ�����!")
			 End If
		 Response.End
		End if

		Set RS=Server.CreateObject("ADODB.RECORDSET")

		  if KS.C_S(Channelid,12)=1 Or KS.C_S(ChannelID,12)=3 then
			verific=0
		  else
			verific=1
		  end if
		RS.Open "Select top 1 * From KS_Comment Where 1=0",Conn,1,3
		RS.AddNew
		 RS("ChannelID")=ChannelID
		 RS("InfoID")=InfoID
		 RS("AnounName")=AnounName
		 RS("UserName")=AnounName
		 RS("Anonymous")=Anonymous
		 RS("Email")=Email
		 RS("Content")=KS.HTMLEncode(C_Content)
		 RS("UserIP")=KS.GetIP
		 RS("Point")=0
		 RS("Score")=0
		 RS("OScore")=0
		 RS("Verific")=Verific
		 RS("AddDate")=Now
		RS.UpDate
		RS.MoveLast
		Dim CommentID:CommentID=RS("ID")
		RS.Close
		
		If KS.ChkClng(groupid)<>0 and Verific=1 Then
		  If KS.ChkClng(KS.U_S(GroupID,6))>0 Then
		  	 RS.Open "Select top 1 Title,Tid,Fname From " & KS.C_S(ChannelID,2) & " Where ID=" & InfoID,conn,1,1
			 If Not RS.Eof Then
			 
			     Call  KS.ScoreInOrOut(KS.C("UserName"),1,KS.ChkClng(KS.U_S(GroupID,6)),"ϵͳ","�����ĵ�[<a href=""" & KS.GetItemUrl(channelid,rs(1),infoid,rs(2)) & """ target=""_blank"">" & RS(0) & "</a>]������!",1002,""&ChannelID&""&InfoID)
			 

             
			 End If
			 RS.Close

		  End If
		End If
		
		If Anonymous=0 Or KSUser.UserName<>"" Then
		 RS.Open "Select top 1 Title,Tid,Fname From " & KS.C_S(ChannelID,2) & " Where ID=" & InfoID,conn,1,1
		 If Not RS.Eof Then
		  Call KSUser.AddLog(AnounName,"�����ĵ�[<a href=""" & KS.GetItemUrl(channelid,rs(1),infoid,rs(2)) & """ target=""_blank"">" & RS(0) & "</a>]������! ����:" & KS.GotTopic(KS.HTMLEncode(C_Content),36),100)
		 End If
		 RS.Close
		End If
		Set RS=Nothing
		
		 If Flag="NotAjax" Then
			 Response.Write "<script>alert('���۷���ɹ�!');location.href='" & ComeUrl & "';</script>"
		 Else
		     Response.Write "ok"
		 End If
		End Sub
		
Function Support()
		  Dim ID,OpType
		  ID=KS.ChkClng(KS.S("ID"))
		  OpType=KS.ChkClng(KS.S("Type"))
		   IF Cbool(Request.Cookies(Cstr(ID))("SupportCommentID"))<>true Then
				Set RS=Server.CreateObject("Adodb.Recordset")
				RS.Open "Select top 1 * From KS_Comment Where ID=" & ID ,Conn,1,3
				 if not rs.eof then
					  if OpType=1 Then
						RS("Score")=RS("Score")+1
					  else
						RS("OScore")=RS("OScore")+1
					  end if
					 RS.UpDate
					 RS.Close:Set RS=Nothing
				 end if
				 Response.Cookies(Cstr(ID))("SupportCommentID")=true
				Else
				Support="����Ͷ��Ʊ�ˣ�"
				Exit Function
			End If
			if OpType=1 Then
				Support="good"
			Else
			 Support="bad"
			End IF
End Function

Sub QuoteSave()
 Dim quoteId:quoteId=KS.ChkClng(KS.S("quoteId"))
 Dim Content:Content=KS.S("QuoteContent")
 Dim QuoteArray,AnounName,QuoteContent,Verific,Anonymous,UserName,LoginTF
 If quoteId=0 Then Response.Write "<script>alert('�������ݳ���!');</script>":Exit Sub
 If Content="" Then Response.Write "<script>alert('�ظ����ݱ�������!');</script>":Exit Sub
 If Len(Content)>KS.ChkClng(KS.C_S(ChannelID,14)) and KS.ChkClng(KS.C_S(ChannelID,14))<>0 Then
	 Response.Write "<script>alert('�������ݱ�����" &KS.C_S(ChannelID,14) & "���ַ�����!');</script>"
	 Response.End
 End if
 Anonymous=KS.ChkClng(KS.S("Anonymous"))
 LoginTF=Cbool(KSUser.UserLoginChecked)
 IF LoginTF=false and (KS.C_S(Channelid,12)=1 or KS.C_S(Channelid,12)=2) Then
  Response.Write "<script>alert('�Բ���,��վֻ����ע���Ա����!');</script>":Exit Sub
 End If
 
 If Anonymous=1 Then
  AnounName="����"
 Elseif Anonymous=0 and LoginTF=false then
  Response.Write "<script>alert('�Բ���,���ȵ�¼!');</script>":Exit Sub
 Else
   AnounName=KSUser.UserName
 End If
 If LoginTF=True Then
  UserName=KSUser.UserName
 Else
  UserName="����"
 End If
 
 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
 RS.Open "Select top 1 channelid,infoid,username,Anonymous,adddate,content,quotecontent from ks_comment where id=" & quoteid,conn,1,1
 if RS.Eof Then
  RS.Close:Set RS=Nothing
  Response.Write "<script>alert('�������ݳ���!');</script>":Exit Sub
 End If
 QuoteArray = RS.GetRows(-1)
 RS.Close
 
 Dim Qstr:Qstr="[dt]���� " 
 If QuoteArray(3,0)=1 Then
  Qstr=Qstr & "����"
 Else
  Qstr=Qstr & "��Ա:" & QuoteArray(2,0)
 End If 
 Qstr=Qstr & " ������" & QuoteArray(4,0) & "����������[/dt][dd]" & QuoteArray(5,0) & "[/dd]"
 If QuoteArray(6,0)<>"" Then
 QuoteContent="[quote]" & QuoteArray(6,0) & Qstr & "[/quote]"
 Else
 QuoteContent="[quote]" & Qstr & "[/quote]"
 End If
 if KS.C_S(Channelid,12)=1 Or KS.C_S(ChannelID,12)=3 then
	verific=0
  else
	verific=1
  end if
 RS.Open "Select top 1 * From KS_Comment Where 1=0",conn,1,3
 RS.AddNew 
    RS("ChannelID")=QuoteArray(0,0)
	RS("InfoID")=QuoteArray(1,0)
	RS("AnounName")=AnounName
	RS("UserName")=UserName
	RS("Anonymous")=Anonymous
	RS("Email")=Email
	RS("Content")=KS.HTMLEncode(Content)
	RS("QuoteContent")=QuoteContent
	RS("UserIP")=KS.GetIP
	RS("Point")=0
	RS("Score")=0
	RS("OScore")=0
	RS("Verific")=Verific
	RS("AddDate")=Now
	RS.UpDate
 RS.Close
 
 If LoginTF=True Then
		 RS.Open "Select Title,Tid,Fname,ID From " & KS.C_S(KS.ChkClng(QuoteArray(0,0)),2) & " Where ID=" & KS.ChkClng(QuoteArray(1,0)),conn,1,1
		 If Not RS.Eof Then
		  Call KSUser.AddLog(AnounName,"�����ĵ�[<a href=""" & KS.GetItemUrl(QuoteArray(0,0),rs(1),QuoteArray(1,0),rs(2)) & """ target=""_blank"">" & RS(0) & "</a>]�ĸ�¥(�ظ�)! ����:" & KS.GotTopic(KS.HTMLEncode(Content),36),100)
		 End If
		 RS.Close
 End If
 
 Set RS=Nothing
 
 Response.Write "<script>alert('��ϲ,�������۷���ɹ�!');try{parent.loadDate(1);parent.closeWindow();}catch(e){top.location.replace(document.referrer);}</script>"
End Sub

%>
 
