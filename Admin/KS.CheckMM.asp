<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.FileIcon.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************

Const FilterFiles="Kesion.CommonCls.asp,Kesion.Label.SQLCls.asp,Kesion.IfCls.asp,Collect_ItemModify4.asp,Comment.asp,Mood.asp,User_PayReceive.asp,KS_Char.asp,Wap_FilesCls.asp,MyFunction.asp,Upload.asp,User_Photo.asp,Alipay_NotifyUrl.asp,KS.Template.asp,Upfilesave.asp,Kesion.UpFileCls.asp,ex.asp,user_files.asp"  '������˲������ļ�,����ļ��ö��Ÿ���
dim Report,delnum:delnum=0
Dim KS:Set KS=New PublicCls
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>������ľ��</title>
<link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
<script language="JavaScript" src="../KS_Inc/common.js"></script></head>
<script language="JavaScript" src="../KS_Inc/jquery.js"></script></head>
<body scroll="no" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div class='topdashed sort'>���߼��ľ��</div>
<div style="height:95%; overflow: auto; width:100%" align="center">
<%
	if KS.G("act")<>"scan" then
%>
				<form action="?act=scan" method="post">

 <table style="margin-top:4px" width="99%" align="center" class="Ctable" border="0" cellpadding="0" cellspacing="0">
				
				<tr class="tdbg">       
				 <td height="28" width="150"align="right" class="clefttitle"><strong>����·����</strong></td>              <td><input name="path" type="text" style="border:1px solid #999" value="\" size="30" />
				 * ��վ��Ŀ¼�����·�����\�������������վ
				 </td>    
				</tr>
				<tr class="tdbg">
				 <td height="28" width="150" align="right" class="clefttitle"><strong>������չ����</strong></td>              <td><input name="FileExt" type="text" style="border:1px solid #999" value="asp,asa,gif,jpg" size="30" />
				 *�����չ��,���ö��Ÿ���
				 </td>    
				</tr>
				<tr class="tdbg">
				  <td height="28" align="right" class="clefttitle"><strong>������չ�����ļ�����ֱ��ɾ����</strong></td>
				  <td><input type="text" name="delfilelist" id="delfilelist" value="cdx,asa,cer" size='30'> <font color=red>˵��KesionCMSϵͳĬ���ǲ������������͵��ļ�,һ����������������վ����������͵��ļ�,�����Ǳ��ϴ���ľ��,�����ɾ�������ա�
				  </td>
                </tr>
				<tr class="tdbg">
				  <td height="28" align="right" class="clefttitle"><strong>��Σ���ļ�����</strong></td>
				  <td>&nbsp;
				  <input type="radio" name="delfile" value="1">ֱ��ɾ��
				  <input type="radio" name="delfile" value="0" checked>��ʾ��ɾ��
				  </td>
                </tr>
				</table>
				
		     <div style="text-align:center;margin-top:20px">
				<input type="submit" value=" ��ʼɨ�� " onClick="if($('#delfilelist').val()!=''){return(confirm($('#delfilelist').val()+'���ļ�����ɾ��,ȷ����ʼɨ����?'))}" class="button" />
			 </div>
				</form>
			<div style="line-height:24px;text-align:left;padding:10px;background:#ffffee;margin:5px 2px;border:1px #f9c943 solid">
			 ʹ��˵��:<br />
			  �١�ִ�б�������Ҫ�ķѼ�����ʱ�䣬���ڷ������ٵ�ʱ��ִ�б�������
			  <br />
			  �ڡ��������ô˹��߼������վ���Ƿ����ľ���ļ���
			  <br />�ۡ������µ�ľ����־����仯����,�����߲���֤����ľ�����Բ������
			</div>
<%
	else
	%>
	<br>
	<table border="0" width="98%" align="center">
	<tr>
	 <td id='message' style="line-height:24px;padding:10px;background:#ffffee;border:1px #f9c943 solid">   
		<table border="0" width="100%">
		<tr>
		 <td width="150" height="50"><img src="images/wait.gif"></td>
		 <td id="msg"></td>
		</tr>
		</table>
	</td>
   </tr>
  </table>
<%
		server.ScriptTimeout = 90000
		DimFileExt = Request("FileExt")
		If DimFileExt="" Then DimFileExt="asp"
		delfilelist= Request("delfilelist")
		If delfilelist="" Then delfilelist="0"
		If delfilelist="0" Then DimFileExt=DimFileExt & ",cdr,asa,cer"   '���û��ֱ��ɾ�����������б�
		
		Sun = 0
		SumFiles = 0
		SumFolders = 1
		if request.Form("path")="" then
			response.Write("<script>alert('������Ҫ����·��!');history.back();</script>")
			response.End()
		end if
		timer1 = timer
		if request.Form("path")="\" then
			TmpPath = Server.MapPath("\")
		elseif request.Form("path")="." then
			TmpPath = Server.MapPath(".")
		else
			TmpPath = Server.MapPath("\")&"\"&request.Form("path")
		end if
		Call ShowAllFile(TmpPath)
		
		Dim Msg
		If Sun=0 Then
		 Msg="<img src=images/succeed.gif align=absmiddle> ɨ����ϣ���ϲ��,����ϵͳ�ܰ�ȫ,û�з���ľ��,ϣ�����ܱ��֣�"
		Else
		 Msg="<img src=images/succeed.gif align=absmiddle> ɨ����ϣ�һ������ļ���<font color=""#FF0000"">" & SumFolders & "</font>�����ļ�<font color=""#FF0000"">" & SumFiles & "</font>�������ֿ��ɵ�<font color=""#FF0000"">" & Sun & "</font>��,ɾ��Σ���ļ�<dont color=red>" & delnum & "</font>��"
		End If
		Response.Write "<script>message.innerHTML='" & MSG & "';</script>"
		Response.Flush()


%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="CContent">

  <tr>
    <td class="CPanel" style="padding:5px;line-height:170%;clear:both;font-size:12px">
       
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
	 <tr>
		 <td valign="top">
		  <% If Sun<>0 Then%>
			 <table width="99%" align="center" border="1" cellpadding="0" cellspacing="0" style="padding:5px;line-height:170%;clear:both;font-size:12px">
			 <tr>
			   <td>�ļ����·��</td>
			   <td>������</td>
			   <td width="230">����</td>
			   <td>����/�޸�ʱ��</td>
			   <td>�������</td>
			   </tr>
			   <tbody id='tablemsg'>
			   </tbody>
			 <%=Report%>
			 </table>
		 <%end if%>	 
			 </td>
	 </tr>
	</table>
</td></tr></table>
<%
	timer2 = timer
	thetime=cstr(int(((timer2-timer1)*10000 )+0.5)/10)
	response.write "<br><font size=""2"">���μ�⹲����"&thetime&"����</font>"
end if

%>
<br />
</body>
</html>
<%

'��������path������Ŀ¼�����ļ�
Sub ShowAllFile(Path)
	Set FSO = CreateObject("Scripting.FileSystemObject")
	if not fso.FolderExists(path) then exit sub
	Set f = FSO.GetFolder(Path)
	Set fc2 = f.files
	For Each myfile in fc2
	     ext=FSO.GetExtensionName(path&"\"&myfile.name)
		If CheckExt(ext,DimFileExt) or CheckExt(FSO.GetExtensionName(path&"\"&myfile.name),delfilelist) Then
			Response.Write "<script>msg.innerHTML='���ڼ���ļ�:" & myfile.name & "';</script>"
		    If KS.FoundInArr(lcase(FilterFiles),lcase(myfile.name),",")=false Then
			 Call ScanFile(Path&Temp&"\"&myfile.name, "",ext)
			End If
			SumFiles = SumFiles + 1
			Response.Flush()
		End If
	Next
	Set fc = f.SubFolders
	For Each f1 in fc
	     if instr(lcase(f1.name),".")<>0 Then
			Report = Report&"<tr><td>"&replace(path&"\"&f1.name,server.MapPath("\")&"\","",1,1,1)&"</td><td>���Ϸ����ļ�������</td><td colspan=2>Σ���ļ��У�һ������IIS���ļ���ִ��©��,��ʽΪ x.asp�µ�ͼƬľ����ܱ�ִ��</td><td><font color=red>���ֹ�ɾ�����ļ���</font></td></tr>"
		 End If
		  ShowAllFile path&"\"&f1.name
		  SumFolders = SumFolders + 1
    Next
	Set FSO = Nothing
End Sub

'����ļ�
Sub ScanFile(FilePath, InFile,ext)
	If InFile <> "" Then
		Infiles = "���ļ���<a href=""http://"&Request.Servervariables("server_name")&"\"&InFile&""" target=_blank>"& InFile & "</a>�ļ�����ִ��"
	End If
	 temp = "<a href=""http://"&Request.Servervariables("server_name")&"\"&replace(FilePath,server.MapPath("\")&"\","",1,1,1)&""" target=_blank>"&replace(FilePath,server.MapPath("\")&"\","",1,1,1)&"</a>"

	if instr(FilePath,";")<>0 then
		Report = Report&"<tr><td>"&temp&"</td><td>���Ϸ����ļ���</td><td>Σ���ļ���һ������IIS���ļ���ִ��©��,��ʽΪ *.asp;*.gif��</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td><td>" & DeleteFile(filepath) & "</td></tr>"
		Sun = Sun + 1
		exit sub
	end if
	if CheckExt(ext,delfilelist) Then
		Report = Report&"<tr><td>"&temp&"</td><td>�Ƿ���չ��</td><td>Σ���ļ�����KesionCMSϵͳ�ļ���</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td><td>" & DeleteFile(filepath) & "</td></tr>"
		Sun = Sun + 1
	  exit sub
	End If
	
    	
	Set FSOs = CreateObject("Scripting.FileSystemObject")
	on error resume next
	set ofile = fsos.OpenTextFile(FilePath)
	filetxt = Lcase(ofile.readall())
	If err Then Exit Sub end if
	if len(filetxt)>0 then
		    '��������
			'Check "WScr"&DoMyBest&"ipt.Shell"
			If instr( filetxt, Lcase("WScr"&DoMyBest&"ipt.Shell") ) or Instr( filetxt, Lcase("clsid:72C24DD5-D70A"&DoMyBest&"-438B-8A42-98424B88AFB8") ) then
				Report = Report&"<tr><td>"&temp&"</td><td>WScr"&DoMyBest&"ipt.Shell ���� clsid:72C24DD5-D70A"&DoMyBest&"-438B-8A42-98424B88AFB8</td><td>Σ�������һ�㱻ASPľ�����á�"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td><td><font color=blue>���ֹ�ȷ��</font></td></tr>"
				Sun = Sun + 1
			End if
			'Check "She"&DoMyBest&"ll.Application"
			If instr( filetxt, Lcase("She"&DoMyBest&"ll.Application") ) or Instr( filetxt, Lcase("clsid:13709620-C27"&DoMyBest&"9-11CE-A49E-444553540000") ) then
				Report = Report&"<tr><td>"&temp&"</td><td>She"&DoMyBest&"ll.Application ���� clsid:13709620-C27"&DoMyBest&"9-11CE-A49E-444553540000</td><td>Σ�������һ�㱻ASPľ�����á�"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td><td><font color=blue>���ֹ�ȷ��</font></td></tr>"
				Sun = Sun + 1
			End If
			'Check .Encode
			Set regEx = New RegExp
			regEx.IgnoreCase = True
			regEx.Global = True
			regEx.Pattern = "@\s*LANGUAGE\s*=\s*[""]?\s*(vbscript|jscript|javascript).encode\b"
			If regEx.Test(filetxt) Then
				Report = Report&"<tr><td>"&temp&"</td><td>(vbscript|jscript|javascript).Encode</td><td>�ƺ��ű��������ˣ�һ��ASP�ļ��ǲ�����ܵġ�"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td><td><font color=blue>���ֹ�ȷ��</font></td></tr>"
				Sun = Sun + 1
			End If
			'Check my ASP backdoor :(
			regEx.Pattern = "\bEv"&"al\b"
			If regEx.Test(filetxt) Then
				Report = Report&"<tr><td>"&temp&"</td><td>Ev"&"al</td><td>e"&"val()��������ִ������ASP���룬��һЩ�������á�����ʽһ���ǣ�ev"&"al(X)<br>����javascript������Ҳ����ʹ�ã��п������󱨡�"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td><td><font color=blue>���ֹ�ȷ��</font></td></tr>"
				Sun = Sun + 1
			End If
			'Check exe&cute backdoor
			regEx.Pattern = "[^.]\bExe"&"cute\b"
			If regEx.Test(filetxt) and instr(filetxt,"conn.execute")=0 Then
				Report = Report&"<tr><td>"&temp&"</td><td>Exec"&"ute</td><td>e"&"xecute()��������ִ������ASP���룬��һЩ�������á�����ʽһ���ǣ�ex"&"ecute(X)��<br>"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td><td><font color=blue>���ֹ�ȷ��</font></td></tr>"
				Sun = Sun + 1
			End If
			
			'===================10-31������=========================
			dim findcontent:findcontent=lcase(filetxt)
			if (instr(findcontent,"exec"&"utestatement")<>0 or instr(findcontent,"msscript"&"control.scriptcontr")<>0 or instr(findcontent,"clsid:72c24dd5-d70"&"a-438b-8a42-98424b88afb8")<>0 or instr(findcontent,"clsid:f935dc22-1cf0-11d0-adb9"&"-00c04fd58a0b")<>0 or instr(findcontent,"clsid:093ff999-1ea0-4079-9525-961"&"4c3504b74")<>0 or instr(findcontent,"clsid:f935dc26-1cf0-11d0-adb9-"&"00c04fd58a0b")<>0 or instr(findcontent,"clsid:0d43fe01"&"-f093-11cf-8940-00a0c9054228")<>0) then
				Report = Report&"<tr><td>"&temp&"</td><td>Execute"&"Global/execute/clsid</td><td>Execute"&"Global()��������ִ������ASP���룬��һЩ�������á�����ʽһ���ǣ�Execute"&"Global(X)��<br>"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td><td><font color=blue>���ֹ�ȷ��</font></td></tr>"
				Sun = Sun + 1
			End If
			
			
			regEx.Pattern = "execute\s*request"
			If regEx.Test(findcontent) and instr(lcase(filename),"scan.asp")=0 Then
				Report = Report&"<tr><td>"&temp&"</td><td>Execute"&"Global/execute</td><td>Execute"&"Global()��������ִ������ASP���룬��һЩ�������á�����ʽһ���ǣ�Execute"&"Global(X)��<br>"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td><td><font color=blue>���ֹ�ȷ��</font></td></tr>"
				Sun = Sun + 1
			end if
			
			regEx.Pattern = "executeglobal\s*request"
			If regEx.Test(findcontent) and instr(lcase(filename),"scan.asp")=0 Then
				Report = Report&"<tr><td>"&temp&"</td><td>Execute"&"Global/execute</td><td>Execute"&"Global()��������ִ������ASP���룬��һЩ�������á�����ʽһ���ǣ�Execute"&"Global(X)��<br>"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td><td><font color=blue>���ֹ�ȷ��</font></td></tr>"
				Sun = Sun + 1
			end if
			regEx.Pattern = "<script.*runat.*server(\n|.)*execute(\n|.)*<\/script>"
			If regEx.Test(findcontent) and instr(lcase(filename),"scan.asp")=0 Then
				Report = Report&"<tr><td>"&temp&"</td><td>Execute"&"Global/execute</td><td>Execute"&"Global()��������ִ������ASP���룬��һЩ�������á�����ʽһ���ǣ�Execute"&"Global(X)��<br>"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td><td><font color=blue>���ֹ�ȷ��</font></td></tr>"
				Sun = Sun + 1
			end if

			Set regEx = Nothing
			
			
			'===================��ǿ������============================================
			
	 
			
		'Check include file
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "<!--\s*#include\s*file\s*=\s*"".*"""
		Set Matches = regEx.Execute(filetxt)
		For Each Match in Matches
			tFile = Replace(Mid(Match.Value, Instr(Match.Value, """") + 1, Len(Match.Value) - Instr(Match.Value, """") - 1),"/","\")
			ext=FSOs.GetExtensionName(tFile)
			If Not CheckExt(ext,DimFileExt) Then
				Call ScanFile( Mid(FilePath,1,InStrRev(FilePath,"\"))&tFile, replace(FilePath,server.MapPath("\")&"\","",1,1,1),ext )
				SumFiles = SumFiles + 1
			End If
		Next
		Set Matches = Nothing
		Set regEx = Nothing
		
		'Check include virtual
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "<!--\s*#include\s*virtual\s*=\s*"".*"""
		Set Matches = regEx.Execute(filetxt)
		For Each Match in Matches
			tFile = Replace(Mid(Match.Value, Instr(Match.Value, """") + 1, Len(Match.Value) - Instr(Match.Value, """") - 1),"/","\")     
			ext=FSOs.GetExtensionName(tFile)
			If Not CheckExt(ext,DimFileExt) Then
				Call ScanFile( Server.MapPath("\")&"\"&tFile, replace(FilePath,server.MapPath("\")&"\","",1,1,1),ext )
				SumFiles = SumFiles + 1
			End If
		Next
		Set Matches = Nothing
		Set regEx = Nothing
		
		'Check Server&.Execute|Transfer
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "Server.(Exec"&"ute|Transfer)([ \t]*|\()"".*"""
		Set Matches = regEx.Execute(filetxt)
		For Each Match in Matches
			tFile = Replace(Mid(Match.Value, Instr(Match.Value, """") + 1, Len(Match.Value) - Instr(Match.Value, """") - 1),"/","\")
			ext=FSOs.GetExtensionName(tFile)
			If Not CheckExt(ext,DimFileExt) Then
				Call ScanFile( Mid(FilePath,1,InStrRev(FilePath,"\"))&tFile, replace(FilePath,server.MapPath("\")&"\","",1,1,1),ext )
				SumFiles = SumFiles + 1
			End If
		Next
		Set Matches = Nothing
		Set regEx = Nothing
		
		'Check Server&.Execute|Transfer
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "Server.(Exec"&"ute|Transfer)([ \t]*|\()[^""]\)"
		If regEx.Test(filetxt) Then
			Report = Report&"<tr><td>"&temp&"</td><td>Server.Exec"&"ute</td><td>���ܸ��ټ��Server.e"&"xecute()����ִ�е��ļ��������Ա���м�顣<br>"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td><td><font color=blue>���ֹ�ȷ��</font></td></tr>"
			Sun = Sun + 1
		End If
		Set Matches = Nothing
		Set regEx = Nothing
		
		'Check Crea"&"teObject
		Set regEx = New RegExp
		regEx.IgnoreCase = True
		regEx.Global = True
		regEx.Pattern = "CreateO"&"bject[ |\t]*\(.*\)"
		Set Matches = regEx.Execute(filetxt)
		For Each Match in Matches
			If Instr(Match.Value, "&") or Instr(Match.Value, "+") or Instr(Match.Value, """") = 0 or Instr(Match.Value, "(") <> InStrRev(Match.Value, "(") Then
				Report = Report&"<tr><td>"&temp&"</td><td>Creat"&"eObject</td><td>Crea"&"teObject����ʹ���˱��μ�������ϸ���顣"&infiles&"</td><td>"&GetDateCreate(filepath)&"<br>"&GetDateModify(filepath)&"</td><td><font color=blue>���ֹ�ȷ��</font></td></tr>"
				Sun = Sun + 1
				exit sub
			End If
		Next
		Set Matches = Nothing
		Set regEx = Nothing

	end if
	set ofile = nothing
	set fsos = nothing
End Sub

'����ļ���׺�������Ԥ����ƥ�伴����TRUE
Function CheckExt(FileExt,CheckFileExt)
	If DimFileExt = "*" Then CheckExt = True
	Ext = Split(CheckFileExt,",")
	For i = 0 To Ubound(Ext)
		If Lcase(FileExt) = Ext(i) Then 
			CheckExt = True
			Exit Function
		End If
	Next
End Function

Function GetDateModify(filepath)
	Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(filepath) 
	s = f.DateLastModified 
	set f = nothing
	set fso = nothing
	GetDateModify = s
End Function

Function GetDateCreate(filepath)
	Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(filepath) 
	s = f.DateCreated 
	set f = nothing
	set fso = nothing
	GetDateCreate = s
End Function

'ɾ���ļ�
Public Function DeleteFile(FileStr)
       if request("delfile")="1" Then
		   Dim FSO
		   On Error Resume Next
		   Set FSO = CreateObject("Scripting.FileSystemObject")
			FSO.DeleteFile FileStr, True
		   Set FSO = Nothing
		   If Err.Number <> 0 Then
			Err.Clear
			DeleteFile="<font color=green>ɾ��ʧ��,���ֹ�ɾ��</font>"
		   Else
		   delnum=delnum+1
			DeleteFile="<font color=red>��ɾ��</font>"
		   End If
	   else
	     DeleteFile="<font color=blue>��ȷ�ϲ��ֹ�ɾ��</font>"
	   end if
End Function

Set KS=Nothing
CloseConn
%>