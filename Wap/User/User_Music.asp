<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="UpFileSave.asp"-->
<%Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="我的音乐">
<p>
<%
Set KS=New PublicCls
IF Cbool(KSUser.UserLoginChecked)=False Then
   Response.redirect KS.GetDomain&"User/Login/"
   Response.End
End If

MaxPerPage =20
Action=Trim(Request("Action"))
CurrentPage=Trim(request("page"))
If Isnumeric(CurrentPage) Then
CurrentPage=Clng(CurrentPage)
Else
CurrentPage=1
End If

If Conn.Execute("Select Count(BlogID) From KS_Blog Where UserName='" & KSUser.UserName & "'")(0)=0 Then
   Response.Write "你不对，你还没有开通空间功能！<br/>" &vbcrlf
ElseIf Conn.Execute("Select status From KS_Blog Where UserName='" & KSUser.UserName & "'")(0)<>1 Then
   Response.Write "对不起,你的空间还没有通过审核或被锁定!<br/>" &vbcrlf
Else
   Response.Write "<a href=""User_Music.asp?" & KS.WapValue & """>我的音乐</a> " &vbcrlf
   Response.Write "<a href=""User_Music.asp?Action=addlink&amp;" & KS.WapValue & """>增加音乐</a><br/>" &vbcrlf
   Select Case Action
	   Case "addlink"
	   Call AddMusicLink()
	   Case "addsave"
	   Call AddMusicLinkSave()
	   Case "play"
	   Call MusicPlay()
	   Case "del"
	   Call SongDel()
	   Case Else
	   Call info()
   End Select
End If
%>
<a href="Index.asp?<%=KS.WapValue%>">我的地盘</a>
<a href="<%=KS.GetGoBackIndex%>">返回首页</a><br/>
<%
Set KSUser=Nothing
Set KS=Nothing
Call CloseConn
%>
</p>
</card>
</wml>

		
<%
Sub info()
	set rs=server.createobject("adodb.recordset")
	sql="select * from ks_blogmusic where Username='"&KSUser.UserName&"' order by adddate desc"
	rs.open sql,Conn,1,1
	If rs.eof And rs.bof Then
       Response.Write "您没有上传音乐!<br/>" &vbcrlf
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
   If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
      Rs.Move (CurrentPage - 1) * MaxPerPage
   Else
      CurrentPage = 1
   End If
   do while not rs.eof
      Response.Write "<a href=""User_Music.asp?Action=play&amp;id=" & rs(0) & "&amp;" & KS.WapValue & """>" & KS.HTMLEncode(rs("SongName")) & "</a>" &vbcrlf
	  Response.Write "<a href=""User_Music.asp?Action=addlink&amp;id=" & rs(0) & "&amp;" & KS.WapValue & """>修改</a> <a href=""User_Music.asp?Action=del&amp;id=" & rs(0) & "&amp;" & KS.WapValue & """>删除</a><br/>" &vbcrlf
      rs.movenext
	  I = I + 1
	  If I >= MaxPerPage Then Exit Do
   loop
   Call KS.ShowPageParamter(totalPut, MaxPerPage, "User_Music.asp", True, "个", CurrentPage, "" & KS.WapValue & "")
   Response.Write "<br/>" &vbcrlf
   
   End If
   rs.close:set rs=Nothing
End Sub

Sub MusicPlay()
    Dim ID:ID=KS.ChkClng(KS.S("ID"))
	If id=0 Then
	   Response.Write "非法参数!<br/>" &vbcrlf
	Else
	   Dim RS:Set RS=Server.Createobject("adodb.recordset")
	   rs.open "select * from ks_blogmusic where id="&Id,Conn,1,1
	   If rs.eof Then
	      Response.Write "非法参数!<br/>" &vbcrlf
	   Else
	      Response.Write "名称:" & KS.HTMLEncode(rs("SongName")) & "<br/>" &vbcrlf
	      Response.Write "歌手:" & rs("singer") & "<br/>" &vbcrlf
		  Response.Write "时间:" & rs("adddate") & "<br/>" &vbcrlf
		  Response.Write "<a href=""" & rs("url") & """>下载</a><br/><br/>" &vbcrlf
		  
	   End If
	   rs.close:set rs=nothing
	End If
%>

<%
End Sub
	
Sub AddMusicLink()

    Dim ID:ID=KS.ChkClng(KS.S("ID"))
	Dim SongName,Url,Singer
	If id<>0 Then
	   Dim RS:Set RS=Server.Createobject("adodb.recordset")
	   rs.open "select * from ks_blogmusic where id="&Id,conn,1,1
	   If not rs.eof Then
	      songname=rs("songname")
		  url=rs("url")
		  singer=rs("singer")
	   End If
	   rs.close:set rs=nothing
	Else
	   If KS.S("BasicType")="9995" Then
	      On Error Resume Next
		  sTemp = Kesion
		  sTemp=split(sTemp,"|||")
		  If Cbool(sTemp(0))=False Then
	         Response.Write sTemp(1)
		  Else
	         url=KS.Setting(3) & right(replace(sTemp(1),"|",""),len(replace(sTemp(1),"|",""))-1)
		  End if
	   Else
	   Response.Write "<a href=""User_UpFile.asp?ChannelID=9995&amp;" & KS.WapValue & """>WAP2.0上传</a><br/>" &vbcrlf
	   End If
	End If
	%>
    歌曲名称:<input name="SongName<%=Minute(Now)%><%=Second(Now)%>" type="text" value="<%=songname%>" maxlength="100" /><br/>
    播放地址:<input name="Url<%=Minute(Now)%><%=Second(Now)%>" type="text" value="<%=url%>" maxlength="100" /><br/>
    歌手名称:<input name="Singer<%=Minute(Now)%><%=Second(Now)%>" type="text" value="<%=singer%>" maxlength="100" /><br/>
    <anchor>确定保存<go href="User_Music.asp?action=addsave&amp;id=<%=id%>&amp;<%=KS.WapValue%>" method="post">
    <postfield name="SongName" value="$(SongName<%=Minute(Now)%><%=Second(Now)%>)"/>
    <postfield name="Url" value="$(Url<%=Minute(Now)%><%=Second(Now)%>)"/>
    <postfield name="Singer" value="$(Singer<%=Minute(Now)%><%=Second(Now)%>)"/>
    </go></anchor><br/>
<%
End Sub
        
Sub AddMusicLinkSave()
    Dim SongName:SongName=KS.S("SongName")
	Dim Url:Url=KS.S("Url")
	Dim Singer:Singer=KS.S("Singer")
	Dim ID:ID=KS.ChkClng(KS.S("ID"))
	IF SongName="" Then
	   Response.Write "歌曲名称必须输入!<br/>" &vbcrlf
	   Exit Sub
	End If
	IF Url="" Then
	   Response.Write "歌曲番放地址必须输入!<br/>" &vbcrlf
	   Exit Sub
	End If
	If ID=0 Then
	   Conn.Execute("Insert Into KS_BlogMusic(songname,url,singer,adddate,username) values('" & SongName & "','" & Url & "','" & Singer & "'," & SqlNowString & ",'" & KSUser.UserName &"')")
	   Response.Write "恭喜,歌曲添加成功!<br/>" &vbcrlf
	Else
	   Conn.Execute("Update KS_BlogMusic set songname='" & SongName & "',url='" & Url & "',singer='" & Singer & "' where username='" & KSUser.UserName & "' and id=" & ID)
	   Response.Write "恭喜,歌曲修改成功!<br/>" &vbcrlf
	End If
End Sub
		
	
Sub SongDel()
    On Error Resume Next
	Dim i,id:id=KS.FilterIDs(KS.S("id"))
	If (id="") Then
	   Response.Write "对不起,参数传递出错!<br/>" &vbcrlf
	   Exit Sub
	Else
	   dim ids:ids=split(id,",")
	   for i=0 to ubound(ids)
		   ks.deletefile(conn.execute("select url from ks_blogmusic where id=" & ids(i) & "and username='" & ksuser.username & "'")(0))
	   next
	   Conn.Execute("delete from ks_blogmusic where id in(" & id & ")")
	   Response.Write "恭喜,删除成功!<br/>" &vbcrlf
	End If
End Sub
%> 
