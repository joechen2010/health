<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<%Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="文件管理">
<p>
<%
Set KS=New PublicCls
Action=Trim(Request("Action"))

IF Cbool(KSUser.UserLoginChecked)=False Then
   Response.redirect KS.GetDomain&"User/Login/"
   Response.End
End If
%>
<%
ntime=Server.Urlencode(now())
path=Trim(Request("path"))
PathCanModify=Trim("/"&KS.Setting(91)&"User/"&KSUser.UserName&"/")
if InvalidChar(path) then
	path=PathCanModify
end if
if right(path,1)<>"/" then
	path=path&"/"
end if
if left(path,1)<>"/" then
	path="/"&path
end if
if PathCanModify<>left(path,len(PathCanModify)) then
   Response.write "=出错信息=<br/>"
   Response.write "抱歉,你只有管理目录"""&PathCanModify&"""及其子目录的权限!<br/>"
else
set Fso=Server.CreateObject(KS.Setting(99))
	If Err = -2147221005 Then
		Response.write "找不到目录,可能参数配置错误!<br/>"
		Response.End
	End If

         Select Case Action
			Case "Showlist"'文件列表
				Showlist
			Case "GetDeleteFile"'删除文件
				GetDeleteFile
			Case "DeleteFolder"'删除文件夹
				DeleteFolder
			Case "delfile","delfolder"'删除文件,删除文件夹
				Delete
			Case "RenameFolder"'重命名文件夹
				RenameFolder
			Case "RenameFile"'重命名文件
				RenameFile
			Case "renfolder","renfile"'重命名文件夹,重命名文件
			    Rename
			Case Else'文件列表
			    Showlist 
			End Select
end if
%>
---------<br/>
<a href="Index.asp?<%=KS.WapValue%>">我的地盘</a><br/>
<a href="<%=KS.GetGoBackIndex%>">返回首页</a><br/>
温馨提醒:为免浪费您的保贵空间,请及时删除无用的文件!<br/>
<%
Set KSUser=Nothing
Set KS=Nothing
Call CloseConn
%>
</p>
</card>
</wml>


<%
'*******************************************
'重命名文件夹
'*******************************************
Sub RenameFolder
    path=Request("path")
	fname=Request("fname")
	epath=Request("epath")
	epage=Request("epage")
	ntime=Request("ntime")
%>
=文件夹重命名=<br/>
将重命名的文件夹:<%=fname%><br/>
新名称:<input name="newname<%=minute(now)%><%=second(now)%>" type="text" maxlength="30" value="<%=fname%>"/><br/>
<anchor>保存命名<go href="User_Files.asp?Action=renfolder&amp;path=<%=pathurl%>&amp;fname=<%=fname%>&amp;epath=<%=pathurl%>&amp;epage=<%=page%>&amp;ntime=<%=ntime%>&amp;<%=KS.WapValue%>" method="post">
<postfield name="newname" value="$(newname<%=minute(now)%><%=second(now)%>)"/>
</go></anchor>
<br/>

<%
End Sub


'*******************************************
'重命名文件
'*******************************************
Sub RenameFile
    path=Request("path")
	fname=Request("fname")
	epath=Request("epath")
	epage=Request("epage")
	ntime=Request("ntime")
%>
=文件重命名=<br/>
将重命名的文件:<%=fname%><br/>
新名称:<input name="newname<%=minute(now)%><%=second(now)%>" type="text" maxlength="30" value="<%=fname%>"/><br/>
<anchor>保存命名<go href="User_Files.asp?Action=renfile&amp;path=<%=pathurl%>&amp;fname=<%=fname%>&amp;epath=<%=pathurl%>&amp;epage=<%=page%>&amp;ntime=<%=ntime%>&amp;<%=KS.WapValue%>" method="post">
<postfield name="newname" value="$(newname<%=minute(now)%><%=second(now)%>)"/>
</go></anchor>
<br/>
<%
End Sub


Sub Rename
    fname=Request("fname")
	newname=Trim(request.form("newname"))
	epath=trim(Request("epath"))
	epage=trim(Request("epage"))
	if epath<>"" then
	   epath=Server.Urlencode(epath)
	end if
	if isnumeric(epage) then
	   epage=fix(epage)
	else
	   epage=1
	end if
	if newname="" then
	   Response.write "指定的文件（夹）名称不能为空!<br/>"
	   Response.write "<anchor>返回上级<prev/></anchor><br/>"
	   response.end
	end if
	if InvalidChar(newname) or instr(newname,"/")>0 then
	   Response.write "指定的路径或文件名称非法!<br/>"
	   Response.write "<anchor>返回上级<prev/></anchor><br/>"
	   response.end
	end if
	bpath=getmappath & path
	if Action="renfolder" then
	   folder_path=bpath&"\"&fname
	   new_path=bpath&"\"&newname

	   if Fso.folderexists(folder_path) and Fso.folderexists(new_path)=false then
	      Fso.getfolder(folder_path).name=newname
		  'Response.write "重命名文件夹 ["&fname&"] 为 ["&newname&"] 成功!<br/>"
		  Response.Redirect "User_Files.asp?path="&epath&"&page="&epage&"&ntime="&ntime&"&" & KS.WapValue & ""
	   elseif Fso.folderexists(new_path) then
		  Response.write "重命名文件夹 ["&fname&"] 失败，文件夹 ["&newname&"] 已存在!<br/>"
		  Response.write "<anchor>返回上级<prev/></anchor><br/>"
	   else
		  Response.write "重命名文件夹 ["&fname&"] 失败，文件夹 ["&fname&"] 不存在!<br/>"
		  Response.write "<anchor>返回上级<prev/></anchor><br/>"
	   end if
	   set Fso=nothing
	elseif Action="renfile" then
	    file_path=bpath&"\"&fname
	    new_path=bpath&"\"&newname
		if Fso.fileexists(file_path) and Fso.fileexists(new_path)=false then
		   Fso.getfile(file_path).name=newname
		   'Response.write "文件 ["&fname&"] 更名为 ["&newname&"] 成功!<br/>"
		   Response.Redirect "User_Files.asp?path="&epath&"&page="&epage&"&ntime="&ntime&"&" & KS.WapValue & ""
	    elseif Fso.fileexists(file_path) then
		   Response.write "文件 ["&fname&"] 更名失败，文件 ["&newname&"] 已经存在!<br/>"
		   Response.write "<anchor>返回上级<prev/></anchor><br/>"
	    else
		   Response.write "文件 ["&fname&"] 更名失败，该文件不存在!<br/>"
		   Response.write "<anchor>返回上级<prev/></anchor><br/>"
	    end if
		set Fso=nothing
	else
	    Response.Redirect "User_Files.asp?path="&epath&"&page="&epage&"&ntime="&ntime&"&wap="&wap
	end if
End Sub

'*******************************************
'删除文件夹
'*******************************************
Sub DeleteFolder
    path=Request("path")
	epath=Request("epath")
	epage=Request("epage")
	ntime=Request("ntime")
%>
=删除文件夹=<br/>
将要删除文件夹:<%=path%><br/>
删除此文件夹吗? 注意：删除后文件夹下文件将不可恢复！<br/>
<a href="User_Files.asp?Action=delfolder&amp;path=<%=path%>&amp;epath=<%=pathurl%>&amp;epage=<%=page%>&amp;ntime=<%=ntime%>&amp;<%=KS.WapValue%>">删除文件夹</a><br/>
<%
End Sub

'删除文件
Sub GetDeleteFile
    path=Request("path")
	epath=Request("epath")
	epage=Request("epage")
	ntime=Request("ntime")
%>
=删除文件=<br/>
<%
sFileType = getname(path,".") 
If sFileType = "jpg" OR sFileType = "gif" Then
			Response.write "<img src="""& path & """ /><br/>"
End If
%>
将要删除文件：<%=path%><br/>
删除此文件吗? 注意：删除后将不可恢复！<br/>
<a href="User_Files.asp?Action=delfile&amp;path=<%=path%>&amp;epath=<%=pathurl%>&amp;epage=<%=page%>&amp;ntime=<%=ntime%>&amp;<%=KS.WapValue%>">删除此文件</a><br/>
<%
End Sub

Sub Delete
    epath=trim(Request("epath"))
	epage=trim(Request("epage"))
	if epath<>"" then
	   epath=Server.Urlencode(epath)
	end if
	if isnumeric(epage) then
	   epage=fix(epage)
	else
	   epage=1
	end if
	if InvalidChar(path) then
	   Response.write "指定的文件名或路径中含有非法字符!<br/>"
	   Response.write "<anchor>返回上级<prev/></anchor><br/>"
	end if
	if len(path)>1 and right(path,1)="/" then
	   path=left(path,len(path)-1)
	end if
	full_path=getmappath &path
	select case Action
	  case "delfile"
		if Fso.fileexists(full_path) then
			Fso.DeleteFile(full_path)
			Response.Redirect "User_Files.asp?path="&epath&"&page="&epage&"&ntime="&ntime&"&" & KS.WapValue & ""
		else
			Response.write "你要删除的文件 "&getname(path,"/")&" 没有找到!<br/>"
			Response.write "<anchor>返回上级<prev/></anchor><br/>"
		end if
	  case "delfolder"
		if Fso.folderexists(full_path) then
			Fso.deletefolder(full_path)
			Response.Redirect "User_Files.asp?path="&epath&"&page="&epage&"&ntime="&ntime&"&" & KS.WapValue & ""
		else
			Response.write "你要删除的子目录 "&getname(path,"/")&" 没有找到!<br/>"
			Response.write "<anchor>返回上级<prev/></anchor><br/>"
		end if
		
	  case else
		Response.Redirect "User_Files.asp?path="&epath&"&page="&epage&"&ntime="&ntime&"&" & KS.WapValue & ""
    end select
End Sub

'*******************************************
'文件列表
'*******************************************
Sub ShowList
    page = Trim(Request("page"))
	set brow2 = Server.CreateObject("MSWC.BrowserType")
	if page<>"" and isnumeric(page) then
	   page=fix(page)
	else
	   page=1
	end if
	if path=PathCanModify then
	   'goparent="此乃根目录<br/>"
	elseif Lcase(path)=Lcase(Session("pathaccess")) then
	   'goparent="只能管理到此目录<br/>"
	else
	   goparent="<a href=""User_Files.asp?path="&Server.Urlencode(left(path,instrrev(path,"/",len(path)-1)))&"&amp;ntime="&ntime&"&amp;" & KS.WapValue & """>↑上一目录</a><br/>"
	end if
	parent_url=Server.Urlencode(parent_url)
	pathurl=Server.Urlencode(path)
	s_folderpath=getmappath &path
	
	if Fso.folderexists(s_folderpath) then
	

	   set folder=Fso.GetFolder(s_folderpath)
	   if (folder.SubFolders.count+folder.Files.count)="0" then
	      Response.write "你还没有上传任何文件!<br/>"
	   else
	       totalpage=1
		   pagesize=6'每页显示数量
	   if folder.files.count mod pagesize=0 then
	      totalpage=folder.files.count\pagesize
	   else
	      totalpage=folder.files.count\pagesize+1
	   end if
	   if page<1 then
	      page=1
	   end if
	   if page>totalpage then
	      page=totalpage
	   end if
%>
	<%=goparent%>
    <%
	Set FsoFile = Fso.GetFolder(getmappath &PathCanModify)
	AllFileSize = FsoFile.size
	Set	FsoFile = Nothing
	%>
	主目录占用空间:<%=KS.GetFileSize(AllFileSize)%><br/>
    占用空间:<%=KS.GetFileSize(folder.size)%>,文件夹:<%=folder.SubFolders.count%>个,文件:<%=folder.Files.count%> 个<br/>
    ---------<br/>
<%for each s_folder in folder.subfolders%>
<img src="../Images/Fsoimg/folder.gif" />
<a href="User_Files.asp?path=<%=pathurl%><%=Server.Urlencode(s_folder.name)%>&amp;ntime=<%=ntime%>&amp;<%=KS.WapValue%>"><%=cuted(s_folder.name,17)%></a>
<a href="User_Files.asp?Action=DeleteFolder&amp;path=<%=pathurl&s_folder.name%>&amp;epath=<%=pathurl%>&amp;epage=<%=page%>&amp;ntime=<%=ntime%>&amp;<%=KS.WapValue%>">删除</a>
<a href="User_Files.asp?Action=RenameFolder&amp;path=<%=pathurl%>&amp;fname=<%=s_folder.name%>&amp;epath=<%=pathurl%>&amp;epage=<%=page%>&amp;ntime=<%=ntime%>&amp;<%=KS.WapValue%>">更名</a>
<br/>
<%
     next
	 i=1
	 startnum=(page-1)*pagesize
	 
	 for each s_file in folder.files
		if i>startnum then
%>	
<img src="../Images/Fsoimg/<%=GetFileIcon(s_file)%>" />
<%=cuted(s_file.name,50)%>
<a href="User_Files.asp?Action=GetDeleteFile&amp;path=<%=pathurl&s_file.name%>&amp;epath=<%=pathurl%>&amp;epage=<%=page%>&amp;ntime=<%=ntime%>&amp;<%=KS.WapValue%>">删除</a>
<a href="User_Files.asp?Action=RenameFile&amp;path=<%=pathurl%>&amp;fname=<%=s_file.name%>&amp;epath=<%=pathurl%>&amp;epage=<%=page%>&amp;ntime=<%=ntime%>&amp;<%=KS.WapValue%>">更名</a>
<br/>
大小:<%=KS.GetFileSize(s_file.size)%>
时间:<%=s_file.datelastmodified%><br/>
<%

		end if
		if i>startnum+pagesize then
			exit for
		end if
		i=i+1
	next
	
	if totalpage>1 then
	   Response.write "---------<br/>"
	   Response.write "共"&folder.files.count&"个文件,每页"&pagesize&"个文件,当前第"&page&"页,共"&totalpage&"页<br/>"
	   if page>1 then
	      Response.write "<a href=""User_Files.asp?path="&pathurl&"&amp;page="&(page-1)&"&amp;ntime="&ntime&"&amp;" & KS.WapValue & """>上一页</a>"
	   end if
	   if page<totalpage then
	      Response.write "　<a href=""User_Files.asp?path="&pathurl&"&amp;page="&(page+1)&"&amp;ntime="&ntime&"&amp;" & KS.WapValue & """>下一页</a>"
	   end if
	 end if
	 Response.write "<br/>"
	 end if
	set folder=nothing
	
else
	Response.write "你还没有上传任何文件!<br/>"
end if
set Fso=nothing
set brow2=nothing
End Sub

'*******************************************
'函数作用：返回文件类型
'*******************************************
Function GetFileIcon(name)
	     Dim FileName,Icon
			FileName=getname(name,".")
			select case FileName
				case "htm","html"
					Icon = "html.gif"
				case "css","ini","inf"
					Icon = "css.gif"
				case "js","vbs","vbe"
					Icon = "js.gif"
				case "exe"
					Icon = "exe.gif"
				case "bat","cmd"
					Icon = "bat.gif"
				case "pdf"
					Icon = "pdf.gif"
				case "ppt"
					Icon = "ppt.gif"
				case "swf"
					Icon = "swf.gif"
				case "xls"
					Icon = "xls.gif"
				case "asp","asa"
					Icon = "asp.gif"
				case "mht","mhtml"
					Icon = "mht.gif"
				case "txt","inc"
					Icon = "text.gif"
				case "jpg","png"
					Icon = "jpg.gif"
				case "bmp"
					Icon = "bmp.gif"
				case "gif"
					Icon = "gif.gif"
				case "mdb"
					Icon = "mdb.gif"
				case "doc"
					Icon = "word.gif"
				case "mid","midi"
					Icon = "midi.gif"
				case "wav","ram"
					Icon = "wav.gif"
				case "mp3"
					Icon = "mp3.gif"
				case "avi","rm","mp","mpg","mpeg","mpe","rmvb"
					Icon = "wmp.gif"
				case "zip"
					Icon = "zip.gif"
				case "rar"
					Icon = "rar.gif"
				case "dll","sys"
					Icon = "dll.gif"
				case "hlp"
					Icon = "hlp.gif"
				case "reg","key"
					Icon = "reg.gif"
				case "chm"
					Icon = "chm.gif"
				case "htc"
					Icon = "htc.gif"
				case "url"
					Icon = "url.gif"
				case "lnk"
					Icon = "lnk.gif"
				case else
					Icon = "unknown.gif"
			end select
			GetFileIcon=Icon
		End Function
'*******************************************
'函数作用：取得文件的后缀名
'*******************************************
function getname(s_string,s_clipchar)
	dim n_strpos
	n_strpos=instrrev(s_string,s_clipchar)
	getname=lcase(right(s_string,len(s_string)-n_strpos))
end function

function InvalidChar(strcontent)
	dim charstr,i
	InvalidChar=false
	charstr="\:*?<>|" & chr(34)
	if strcontent="" or len(strcontent)<=0 then
		InvalidChar=true
		exit function
	end if
	for i=1 to len(charstr)
		if instr(strcontent,mid(charstr,i,1))>0 then
			InvalidChar=true
			exit function
		end if
	next
	if instr(strcontent,vbcrlf)>0 then
		InvalidChar=true
	end if
end function

function cuted(types,num)
  dim ctypes,cnum,ci,tt,tc,cc,cmod,iscuted
  cmod=3
  ctypes=types
  cnum=int(num)
  cuted=""
  tc=0
  cc=0
  iscuted=false
  for ci=1 to len(ctypes)
    if cnum<0 then
	cuted=cuted&"..."
	iscuted=true
	exit for
    end if
    tt=mid(ctypes,ci,1)
    if int(asc(tt))>=0 then
      cuted=cuted&tt
      tc=tc+1
      cc=cc+1
      if tc=2 then
	tc=0
	cnum=cnum-1
      end if
      if cc>cmod then
	cnum=cnum-1
	cc=0
      end if
    else
      cnum=cnum-1
      if cnum<=0 then
	cuted=cuted&"..."
	iscuted=true
	exit for
      end if
      cuted=cuted&tt
    end if
  next
  if iscuted then
	cuted=cuted 
  end if
end function


%>
