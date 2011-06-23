<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.ContentType="text/vnd.wap.wml"
Response.Charset="utf-8"
Response.Write "<?xml version=""1.0"" encoding=""utf-8""?>" &vbcrlf
Response.Write "<!DOCTYPE wml PUBLIC ""-//WAPFORUM//DTD WML 1.1//EN"" ""http://www.wapforum.org/DTD/wml_1.1.xml"">" &vbcrlf
%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
Dim KSCls
Set KSCls = New Display
KSCls.Kesion()
Set KSCls = Nothing

Class Display
        Private KS,ID
		Private RST,Master,BoardName
		Private MaxPerPage, TotalPut , CurrentPage, TotalPage, I
	    Private SqlStr
		Private Sub Class_Initialize()
		    If (Not Response.IsClientConnected)Then
			   Response.Clear
			   Response.End
			End If
		    Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		    Call CloseConn()
		    Set KS=Nothing
		End Sub
		
		Public Sub Kesion()
			If KS.Setting(56)="0" Then Call KS.ShowError("系统提示！","本站已关闭留言功能！")
			If KS.Setting(59)="1" Then Response.Redirect("GuestBook.asp")
			If KS.S("page") <> "" Then
			   CurrentPage = CInt(Request("page"))
			Else
			   CurrentPage = 1
			End If
			ID=KS.ChkClng(KS.S("ID"))

			Call GetSubject()'主题
			Select Case KS.S("Action")
			    Case "SetTop" Call SetTOP'设为置顶
				Case "SetBest" Call SetBest'设为精华
				Case "CancelTop" Call CancelTop'取消置顶
				Case "CancelBest" Call CancelBest'取消精华
				Case "DelSubject" Call DelSubject'删除主题
				Case "DelReply" Call DelReply'删除回复
				Case "Verify" Call Verify'审核
				Case "DependEmpress" Call DependEmpress'沉底帖子
				Case "DependFront" Call DependFront'提升帖子
			End Select
			Call GetReplayList()'回复列表
			Response.Write "<br/>"
			Call GetWriteComment()
			Response.Write "---------<br/>" &vbcrlf
			Response.Write "<a href=""Index.asp?BoardID=" & FCls.RefreshFolderID & "&amp;" & KS.WapValue & """>" & BoardName & "</a><br/><br/>" &vbcrlf
			Response.Write "<a href=""Index.asp?" & KS.WapValue & """>论坛首页</a>&gt;&gt;<a href=""" & KS.GetGoBackIndex & """>返回首页</a><br/>" &vbcrlf
			Response.Write "</p>" &vbcrlf
			Response.Write "</card>" &vbcrlf
			Response.Write "</wml>" &vbcrlf
		End Sub
		
		Sub GetSubJect()
			Set RST=Server.CreateObject("ADODB.RECORDSET")
			RST.Open "Select  top 1 * From KS_GuestBook Where ID=" & ID,Conn,1,3
			If RST.EOF Then
			   RST.Close:Set RST=Nothing
			   Call KS.ShowError("非法参数！","非法参数！")
			End If
			RST("Hits")=RST("Hits")+1
			RST.Update
			FCls.RefreshFolderID = RST("BoardID")
			Master=LFCls.GetSingleFieldValue("select Master from KS_GuestBoard where ID=" & KS.ChkClng(FCls.RefreshFolderID))
			BoardName=LFCls.GetSingleFieldValue("select BoardName from KS_GuestBoard where ID=" & KS.ChkClng(FCls.RefreshFolderID))
			Response.Write "<wml>" &vbcrlf
			Response.Write "<head>" &vbcrlf
			Response.Write "<meta http-equiv=""Cache-Control"" content=""no-Cache""/>" &vbcrlf
			Response.Write "<meta http-equiv=""Cache-Control"" content=""max-age=0""/>" &vbcrlf
			Response.Write "</head>" &vbcrlf
			Response.Write "<card id=""main"" title=""" & KS.HtmlCode(RST("SubJect")) & """>" &vbcrlf
			Response.Write "<p>" &vbcrlf
			
			Response.Write  "您当前正在查看【<a href=""index.asp?boardid=" &FCls.RefreshFolderID&"&amp;" & KS.WapValue & """>" & BoardName & "</a>】版面下的帖子<br/>"
			If CurrentPage<>1 Then Exit Sub
			
			Response.Write "<b>"&KS.HtmlCode(RST("SubJect"))&"</b><br/>"
			Response.Write "作者:<a href=""../User/ShowUser.asp?Keyword="&RST("UserName")&"&amp;" & KS.WapValue & """>" & KS.GetUserRealName(Rst("UserName")) & "</a> ("&KS.DateFormat(RST("AddTime"),17)&")<br/>"
			
			Dim Content
			Content=KS.UBBToHTML(KS.LoseHtml(KS.HTMLToUBB(KS.ReplaceTrim(KS.GetEncodeConversion(RST("Memo"))))))
			Response.Write KS.ContentPagination(Content,200,"Display.asp?ID="&ID&"&amp;" & KS.WapValue & "",False,False)
			Response.Write "<br/>"
			
			Response.Write "已有" & RST("hits") & "人关注过本帖<br/>"
			If RST("IsBest")=1 Then
			   Response.Write "<img src=""../Images/jing.gif"" alt=""."" />本贴被认定为精华<br/>"
			End If
			'Response.Write "引用 "
			Response.Write "<a href=""Reply.asp?TopicID=" & ID &"&amp;Page=" & CurrentPage & "&amp;" & KS.WapValue & """ >回复</a> "
			If RST("UserName")=KSUser.UserName or Check=True Then
			   'Response.Write "<a href=""DownUrls.asp?ID=" & ID &"&amp;" & KS.WapValue & """ >增加下载附件</a> "
			   Response.Write "<a href=""Post.asp?Action=EditPost&amp;ID=" & ID &"&amp;" & KS.WapValue & """>编辑</a> "
			   Response.Write "<a href=""Post.asp?Action=ConnectPost&amp;ID=" & ID &"&amp;" & KS.WapValue & """ >续写</a> "
			   Response.Write "<a href=""Display.asp?ID=" & ID & "&amp;Action=DelSubject&amp;" & KS.WapValue & """>删除</a> "
			End If
			Response.Write "<br/>"
			If Check=True Then
			   If RST("IsTop")=1 Then
			      Response.Write "<a href=""Display.asp?ID=" & ID &"&amp;Action=CancelTop&amp;" & KS.WapValue & """>取消置顶</a> "
			   Else
			      Response.Write "<a href=""Display.asp?ID=" & ID &"&amp;Action=SetTop&amp;" & KS.WapValue & """>设为置顶</a> "
			   End If
			   If RST("IsBest")=1 Then
			      Response.Write "<a href=""Display.asp?ID=" & ID &"&amp;Action=CancelBest&amp;" & KS.WapValue & """>取消精华</a> "
			   Else
			      Response.Write "<a href=""Display.asp?ID=" & ID &"&amp;Action=SetBest&amp;" & KS.WapValue & """>设为精华</a> "
			   End If
			   Response.Write "<a href=""Display.asp?ID=" & ID &"&amp;Action=DependFront&amp;" & KS.WapValue & """>提升</a> "
			   Response.Write "<a href=""Display.asp?ID=" & ID &"&amp;Action=DependEmpress&amp;" & KS.WapValue & """>沉底</a> "
			End If   
			Response.Write "<br/>"
		End Sub
		
		Sub GetReplayList()
		    MaxPerPage=10
			SqlStr = "SELECT * From KS_GuestReply where TopicID=" & KS.ChkClng(KS.S("ID")) & " ORDER BY ID" 
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open SqlStr,Conn,1,1
			IF RS.EOF And RS.BOF Then
			   Totalput=0
			   Exit Sub
		    Else
			   TotalPut= RS.RecordCount
			   If CurrentPage < 1 Then CurrentPage = 1
			   If (CurrentPage - 1) * MaxPerPage > totalPut Then
			      If (TotalPut Mod MaxPerPage) = 0 Then
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
			   Call GetTopicList(RS)
			   Response.Write KS.ShowPagePara(TotalPut, MaxPerPage, "", True, "条", CurrentPage, KS.QueryParam("page"))
			End IF
			RS.Close:Set RS=Nothing
		End Sub
		
		Sub GetTopicList(RS)
		    On Error Resume Next
		    Dim I,N
			If CurrentPage=1 Then
			   N=1
			Else
			   N=MaxPerPage*(CurrentPage-1)
            End If
		 	Do While Not RS.EOF
			   If Not Response.IsClientConnected Then Response.End
			   Response.Write "" & N & "楼:"
			   If RS("TxtHead")<>0 And Isnull(RS("TxtHead"))=False Then Response.Write "<img src=""" & KS.Setting(2) & KS.Setting(3) & "Images/Face1/face" & RS("TxtHead") &".gif"" alt="".""/>"
			   Dim Content

			   Content=KS.UBBToHTML(KS.LoseHtml(KS.HTMLToUBB(KS.ReplaceTrim(Replace(KS.GetEncodeConversion(RS("Content")),"</div>","[br]")))))
			   If RS("Verific")="1" Then
			      Content=Content
			   ElseIf KSUser.GroupID=4 Then
			      Content="该信息未审核,由于您是管理员所以可以看到此信息.<br/>" & Content
			   ElseIf Check=True  Then
			      Content="该信息未审核,由于您是版主所以可以看到此信息.<br/>" & Content
			   Else
			      Content="本站启用审核机制,该信息未通过审核!<br/>"
			   End If
			   Response.Write Content
			   Response.Write "<br/>"
			   If RS("UserName")="游客" Then
			      Response.Write ""&RS("UserName")&" ("&KS.DateFormat(RS("ReplayTime"),17)&")<br/>"
			   Else
			      Response.Write "<a href=""../User/ShowUser.asp?Keyword="&RS("UserName")&"&amp;" & KS.WapValue & """>" & KS.GetUserRealName(RS("UserName")) & "</a> ("&KS.DateFormat(RS("ReplayTime"),17)&") "
			   End If
			   If RS("Verific")="1" Then
			      Response.Write "<a href=""Reply.asp?Action=LeadReply&amp;LeadID=" & RS("ID") & "&amp;TopicID=" & ID &"&amp;Page=" & CurrentPage & "&amp;" & KS.WapValue & """ >引用</a> "
			   Else
			      If Check=True Then
			         Response.Write "<a href=""Display.asp?Action=Verify&amp;ID=" & ID & "&amp;ReplyID=" & RS("ID") &"&amp;" & KS.WapValue & """>审核</a> "
				  End If
			   End If
			   'Response.Write "编辑 "
			   If Check=True Then
			      Response.Write "<a href=""Display.asp?Action=DelReply&amp;ID=" & ID & "&amp;ReplyID=" & RS("ID") &"&amp;" & KS.WapValue & """>删除</a>"
			   End If
			   Response.Write "<br/>"
			   N = N + 1
			   I = I + 1
			   If CurrentPage=1 Then
			      If I > MaxPerPage-2 Then Exit Do
			   Else
			      If I >= MaxPerPage Then Exit Do
			   End If
			   RS.MoveNext
			Loop
		End Sub

		
		Sub GetWriteComment()
		    If KS.Setting(54)<>3 And Cbool(KSUser.UserLoginChecked)=False Then
			   Response.Write "<br/>请登陆回复! 请<a href=""../User/Login.asp?.Club/Display.asp?ID="&ID&"&amp;CurrentPage=" & CurrentPage & """>登录</a>!<br/>"
			Else
			   Dim k,str:str="惊讶|撇嘴|色色|发呆|得意|流泪|害羞|闭嘴|睡觉|大哭|尴尬|发怒|调皮|呲牙|微笑|难过|酷|非典|抓狂|吐吐|"
			   Dim strArr:strArr=Split(str,"|")
			   Response.Write "---------<br/>"
			   Response.Write "<select name=""TxtHead"">"
			   For k=0 To 19
			       Response.Write "<option value=""" & (k+1) & """>" & strArr(k) & "</option>"
			   Next
			   Response.Write "</select> "
			   Dim reSayArry:reSayArry = Array("好帖，要顶!","看帖回帖是美德!","你牛，我顶!","这帖不错，该顶!","支持你!","反对你!")
			   Randomize
			   Response.Write "<input name=""Content" & Minute(Now) & Second(Now) & """ value=""" & reSayArry(Int(Ubound(reSayArry)*Rnd)) & """ maxlength=""300""/> "
			   Response.Write "<anchor>快速回复<go href=""Reply.asp?Action=SaveReply&amp;TopicID=" & ID & "&amp;" & KS.WapValue & """ method=""post"">"
			   Response.Write "<postfield name=""TxtHead"" value=""$(TxtHead)""/>"
			   Response.Write "<postfield name=""Content"" value=""$(Content" & Minute(Now) & Second(Now) & ")""/>"
			   Response.Write "<postfield name=""Page"" value=""" & CurrentPage & """/>"
			   Response.Write "<postfield name=""boardid"" value=""" & FCls.RefreshFolderID & """/>"
			   Response.Write "</go></anchor>"
			   Response.Write "<br/>"			   
			End If
		End Sub

		Sub SetBest()
		    If Cbool(Check)=False Then
			   Call KS.ShowError("系统提示！","对不起，你没有设置的权限!")
			End If
			Conn.Execute("Update KS_GuestBook set isbest=1 where id=" & ID)
			Response.Write "设为精华成功!<br/>"
		End Sub
		
		Sub SetTop()
		    If Cbool(Check)=False Then
			   Call KS.ShowError("系统提示！","对不起，你没有设置的权限!")
			End If
			Conn.Execute("Update KS_GuestBook set istop=1 where id=" & ID)
			Response.Write "设为置顶成功!<br/>"
		End Sub
		
		Sub CancelBest()
		    If Cbool(Check)=False Then
			   Call KS.ShowError("系统提示！","对不起，你没有设置的权限!")
			End If
			Conn.Execute("Update KS_GuestBook set isbest=0 where id=" & ID)
			Response.Write "取消精华成功!<br/>"
		End Sub
		
		Sub CancelTop()
		    If Cbool(Check)=False Then
			   Call KS.ShowError("系统提示！","对不起，你没有设置的权限!")
			End If
			Conn.Execute("Update KS_GuestBook set istop=0 where id=" & ID)
			Response.Write "取消置顶成功!<br/>"
	    End Sub
		
		Sub DelSubject()
		    If Cbool(Check)=False Then
			   Call KS.ShowError("系统提示！","对不起，你没有删除帖子的权限!")
			End If
			Conn.Execute("delete from  KS_Guestbook where id=" & ID)
			Conn.Execute("delete from ks_guestreply where TopicID=" & ID)
			Response.Redirect "Index.asp?BoardID=" & FCls.RefreshFolderID
		End Sub
		
		Sub DelReply()
		    If Cbool(Check)=False Then
			   Call KS.ShowError("系统提示！","对不起，你没有设置的权限!")
			End If
			Conn.Execute("delete from ks_guestreply where ID=" & KS.ChkClng(KS.S("ReplyID")))
			Response.Write "删除"&KS.ChkClng(KS.S("ReplyID"))&"回复成功!<br/>"
	    End Sub
		
		Sub Verify()
		    If Cbool(Check)=False Then
			   Call KS.ShowError("系统提示！","对不起，你没有设置的权限!")
			End If
			Conn.Execute("Update KS_GuestReply set verific=1 where ID=" & KS.ChkClng(KS.S("ReplyID")))
			'Response.Redirect request.servervariables("http_referer")
		End Sub
		
		Sub DependEmpress()
		    If Cbool(Check)=False Then
			   Call KS.ShowError("系统提示！","对不起，你没有设置的权限!")
			End If
			Conn.Execute("Update KS_Guestbook set LastReplayTime='" & Conn.Execute("SELECT LastReplayTime From KS_GuestBook ORDER BY LastReplayTime asc")(0) & "' where id=" & ID)
			Response.Write "本主题放到帖子列表较靠后位置成功!<br/>"
		End Sub
		Sub DependFront()
		    If Cbool(Check)=False Then
			   Call KS.ShowError("系统提示！","对不起，你没有设置的权限!")
			End If
			Conn.Execute("Update KS_Guestbook set LastReplayTime=" & SqlNowString & " where id=" & ID)
			Response.Write"本主题提升到帖子列表最前面成功!<br/>"
		End Sub

		Function Check()
		    If Cbool(KSUser.UserLoginChecked)=False Then 
			   Check=False
			   Exit Function
			Else
			   If KSUser.GroupID=1 Then
			      Check=True
				  Exit Function
			   Else
			      Check=KS.FoundInArr(Master, KSUser.UserName, ",")
			   End If
			End If
		End Function		  
End Class
%>
