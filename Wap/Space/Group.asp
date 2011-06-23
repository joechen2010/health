<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.SpaceCls.asp"-->
<%
Dim KSCls
Set KSCls = New Group
KSCls.Kesion()
Set KSCls = Nothing

Class Group
        Private KS,KSBCls,KSRFObj
		Private TotalPut,RS,MaxPerPage
		Private CurrentPage
		Private ID,Template,TeamName,GroupAdmin
		Private Sub Class_Initialize()
			Set KS=New PublicCls
			Set KSBCls=New BlogCls
			Set KSRFObj=New Refresh
		End Sub
        Private Sub Class_Terminate()
		    Set KS=Nothing
		    Set KSBCls=Nothing
			Set KSRFObj=Nothing
		End Sub
		Public Sub Kesion()
		    If KS.SSetting(0)=0 Then
			   Call KS.ShowError("对不起！","对不起，本站点关闭空间站点功能！")
			End If
			ID=KS.ChkClng(KS.S("ID"))
			If ID=0 Then Response.End()
			Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select * From KS_Team Where ID=" & ID,Conn,1,1
			If RS.Eof And RS.Bof Then
			   Call KS.ShowError("对不起！","参数传递出错！")
			End If
			If RS("Verific")=0 Then
			   Call KS.ShowError("对不起！","该圈子尚未审核!")
			ElseIf RS("Verific")=2 Then
			   Call KS.ShowError("对不起！","该圈子已被管理员锁定!")
			End If
			TeamName=RS("TeamName")
			GroupAdmin=RS("UserName")
			Template="<wml>" &vbcrlf
			Template=Template & "<head>" &vbcrlf
			Template=Template & "<meta http-equiv=""Cache-Control"" content=""no-Cache""/>" &vbcrlf
			Template=Template & "<meta http-equiv=""Cache-Control"" content=""max-age=0""/>" &vbcrlf
			Template=Template & "</head>" &vbcrlf
			Template=Template & "<card id=""main"" title=""" & TeamName & """>" &vbcrlf
			
			Template=Template & KSRFObj.LoadTemplate(KS.WSetting(26))
			Template=KSBCls.ReplaceGroupLabel(RS,Template)
			Select Case KS.S("Action")
			    Case "showtopic"'显示帖子列表
				   Template=Replace(Template,"{$GroupMain}",ShowTopic)
				Case "replay"'回复
				   Template=Replace(Template,"{$GroupMain}",Replay)
				Case "replaysave"'保存回复
				   Call ReplaySave()
				Case "users"'成员列表
				   Template=Replace(Template,"{$GroupMain}",ShowUser)
				Case "join"'申请加入圈子
		  		   Template=Replace(Template,"{$GroupMain}",ShowJoin)
				Case "joinsave"'保存申请加入圈子
				   Template=Replace(Template,"{$GroupMain}",JoinSave)
				Case "alldeltopic"'删除
				   Template=Replace(Template,"{$GroupMain}",AllDelTopic)
				Case "deltopic"'删除
				   Template=Replace(Template,"{$GroupMain}",DelTopic)
				Case "settop"'置顶设置
				   Call SetTop()
				Case "setbest"'精华设置
				   Call SetBest()
				Case "post"'发表新贴
				   Template=Replace(Template,"{$GroupMain}",ShowPost)
				Case "connectpost"
				   Template=Replace(Template,"{$GroupMain}",ConnectPost)
				Case "connectpostsave"
				   Template=Replace(Template,"{$GroupMain}",ConnectPostSave)   
				Case "topicsave"'保存发表
				   Template=Replace(Template,"{$GroupMain}",TopicSave)
				Case "info"'圈子信息
				   Template=Replace(Template,"{$GroupMain}",ShowInfo)
				Case Else'圈子主题列表
				   Template=Replace(Template,"{$GroupMain}",TeamTopic)
			End Select
			Template=Template & "</card>" &vbcrlf
			Template=Template & "</wml>" &vbcrlf
			Response.Write Template
			RS.Close:Set  RS=Nothing
		End Sub
		
		Function Replay()
		    Dim Tid:Tid=KS.Chkclng(KS.S("Tid"))
			Replay = "【回复话题】<br/>"
			IF Cbool(KSUser.UserLoginChecked)=false Then
			   Replay = Replay &"登录后才可以参与该话题的讨论,如要参与讨论请先<a href=""../User/Login/?../Space/Group.asp?Action=replay&amp;ID=" & ID & "&amp;Tid=" & Tid & """>登录</a>到会员中心！"
			Else
			   On Error Resume Next
			   'Replay = Replay &"Re:" & Conn.Execute("select Title from KS_TeamTopic where ID="& Tid )(0) & "<br/>"
			   Replay = Replay &"回复内容:<input name=""Content" & Minute(Now) & Second(Now) & """ type=""text"" maxlength=""300"" emptyok=""false"" value=""""/>"
			   Replay = Replay &"<anchor>回复<go href=""Group.asp?action=replaysave&amp;ID=" & ID & "&amp;Tid=" & Tid & "&amp;" & KS.WapValue & """ method=""post"">"
			   Replay = Replay &"<postfield name=""Title"" value=""Re:" & Conn.Execute("select Title from KS_TeamTopic where ID="& Tid )(0) & """/>"
			   Replay = Replay &"<postfield name=""Content"" value=""$(Content" & Minute(Now) & Second(Now) & ")""/>"
			   Replay = Replay &"</go></anchor><br/>"
			End If
		End Function
		

		'保存回复
		Function ReplaySave()
		    Dim Tid:Tid=KS.chkclng(KS.S("Tid"))
			Dim Title:Title=KS.S("Title")
			Dim Content:Content=KS.S("Content")
			If Content="" Then
			   ReplaySave="请输入回复内容!"
			   Exit Function
			End If
			If Cbool(KSUser.UserLoginChecked)=false Then
			   ReplaySave="请先登录!"
			   Exit Function
			End If
			Dim UserName:UserName=KS.R(KSUser.UserName)
			Dim RS:set RS=server.createobject("adodb.recordset")
			RS.Open "select * from KS_TeamTopic",Conn,1,3
			RS.Addnew
			RS("ParentID")=Tid
			RS("TeamID")=ID
			RS("Title")=Title
			RS("Content")=Content
			RS("Adddate")=Now
			RS("UserIP")=KS.GetIP
			RS("Status")=1
			RS("UserName")=UserName
			RS("IsBest")=0
			RS("IsTop")=0
			RS.Update
			RS.Close:set RS=Nothing
			Response.Redirect KS.GetDomain&"Space/Group.asp?Action=showtopic&ID="& ID & "&Tid=" & Tid & "&" & KS.WapValue & ""
		End Function
		
		Function ShowJoin()
		    IF Cbool(KSUser.UserLoginChecked)=false Then
			   ShowJoin = "对不起，申请加入圈子之前必须先<a href=""../User/Login/?../Space/Group.asp?Action=showpost&amp;ID=" & ID & """>登录</a>到会员中心！<br/>"
			   Exit Function
			End If
			If Not Conn.Execute("select UserName from KS_TeamUsers where UserName='" & KSUser.UserName & "' And TeamID=" & ID).EOF Then
			   ShowJoin = "您不能再申请，产生的可能原因如下：<br/>"
			   ShowJoin = ShowJoin & "您已申请过，未得到圈主的审核;<br/>"
			   ShowJoin = ShowJoin & "您已是本圈子的成员，不需要再申请;<br/>"
			   ShowJoin = ShowJoin & "您可能已被圈主邀请，但您还未在会员中心确认;<br/>"
			   ShowJoin = ShowJoin & "【申请须知】<br/>"
			   ShowJoin = ShowJoin & RS("Note")
			   ShowJoin = ShowJoin & "<br/>"
			   Exit Function
			End If
			ShowJoin = ShowJoin & "【申请加入】<br/>"
			ShowJoin = ShowJoin & "申 请 人:" & KSUser.UserName & "<br/>"
			ShowJoin = ShowJoin & "加入理由:<input name=""Reason" & Minute(Now) & Second(Now) & """ type=""text"" maxlength=""30"" value="""" emptyok=""false""/>"
			ShowJoin = ShowJoin & "<anchor>提交申请<go href=""Group.asp?ID=" & ID & "&amp;Action=joinsave&amp;" & KS.WapValue & """ method=""post"">"
			ShowJoin = ShowJoin & "<postfield name=""UserName"" value=""" & KSUser.UserName & """/>"
			ShowJoin = ShowJoin & "<postfield name=""Reason"" value=""$(Reason" & Minute(Now) & Second(Now) & ")""/>"
			ShowJoin = ShowJoin & "</go></anchor><br/><br/>"
			ShowJoin = ShowJoin & "【申请须知】<br/>"
			ShowJoin = ShowJoin & RS("Note")
			ShowJoin = ShowJoin & "<br/>"
		End Function
		
		'保存申请
		Function JoinSave()
		    Dim id:id=KS.chkclng(KS.S("id"))
			Dim UserName:UserName=KS.R(KS.S("UserName"))
			Dim Reason:Reason=KS.R(KS.S("Reason"))
			If Reason="" Then
			   JoinSave = "请输入加入圈子的理由!<br/><anchor><prev/>返回重写</anchor><br/>"
			   Exit Function
			End If
			Dim RS:set RS=server.createobject("adodb.recordset")
			RS.Open "select * from KS_TeamUsers where TeamID=" & id & " And UserName='" & UserName & "'",Conn,1,3
			If RS.EOF Then
			   RS.Addnew
			   RS("TeamID")=ID
			   RS("UserName")=UserName
			   RS("Status")=2  '申请加入
			   RS("Power")=0   '普通用户
			   RS("Reason")=Reason
			   RS("Applydate")=Now
			   RS.Update
			End If
			RS.Close:set RS=Nothing
			JoinSave = "你的申请已提交，请等待圈主的审核!<br/>"
		End Function
		
		'续写
		Function ConnectPost()
		    Dim Tid:Tid=KS.Chkclng(KS.S("Tid"))
			Set RST=Conn.Execute("select Content from ks_teamtopic where UserName='"&KSUser.UserName&"' and id="&tid&"")
			If RST.EOF Then
			   ConnectPost = "非法参数!<br/>"
			Else
			   ConnectPost = "非法参数!<br/>"
			   ConnectPost = ConnectPost & "【贴子续写】<br/>"
			   ConnectPost = ConnectPost & "尾部内容:" & Right(KS.LoseHtml(RST("Content")),20) & "<br/>"
			   ConnectPost = ConnectPost & "追加内容:<input name=""Content" & Minute(Now) & Second(Now) & """ type=""text"" maxlength=""500"" value=""""/>"
			   ConnectPost = ConnectPost & "<anchor>确定<go href=""Group.asp?Action=connectpostsave&amp;ID=" & ID & "&amp;Tid=" & Tid & "&amp;" & KS.WapValue & """ method=""post""><postfield name=""Content"" value=""$(Content" & Minute(Now) & Second(Now) & ")""/></go></anchor><br/>"
			End If
			RST.Close:Set RST=Nothing
		End Function
		
		'续写保存
		Function ConnectPostSave()
		    Dim Tid:Tid=KS.Chkclng(KS.S("Tid"))
			Set RST=Conn.Execute("select * from KS_TeamTopic where UserName='"&KSUser.UserName&"' And ID="&Tid&"")
			If RST.EOF Then
			   ConnectPostSave = "非法参数!<br/>"
			Else
			   Dim Content:Content=KS.S("Content")
			   If Content="" Then
			      ConnectPostSave = "出错提示，你没有输入续写内容！<br/><anchor><prev/>返回重写</anchor><br/>"
			   Else
			   Set RSObj=Server.CreateObject("Adodb.Recordset")
			   RSObj.Open "select * from KS_TeamTopic where UserName='"&KSUser.UserName&"' And ID="&Tid&"",Conn,1,3
			   RSObj("Content")=RST("Content") & Content
			   RSObj.Update:RSObj.Close:Set RSObj=Nothing
			   ConnectPostSave = "续写成功。<br/><a href=""Group.asp?Action=showtopic&amp;ID="&ID&"&amp;Tid="&Tid&"&amp;" & KS.WapValue & """>贴子查看</a><br/>"
			   End IF
			End If
			RST.Close:Set RST=Nothing
		End Function
		
		
		'发表新贴
		Function ShowPost()
		    IF Cbool(KSUser.UserLoginChecked)=false Then
			   ShowPost = "对不起，发表新贴之前必须先<a href=""../User/Login/?../Space/Group.asp?action=showpost&amp;ID="&ID&""">登录</a>到会员中心！<br/>"
			   Exit Function
			End If
			If Conn.Execute("select UserName from KS_TeamUsers where UserName='"& KSUser.UserName & "' And TeamID=" & ID).EOF Then
			   ShowPost = "对不起，你不是该圈子的成员，没有权利发表话题！<br/>"
			   Exit Function
			ElseIf Conn.Execute("select UserName from KS_TeamUsers where UserName='"& KSUser.UserName & "' And Status<>2 And TeamID=" & ID).EOF Then
			   ShowPost = "对不起，你提交的申请还未得到确认，没有权利发表话题！<br/>"
			   Exit Function
			End If
			ShowPost =""
			ShowPost = ShowPost & "话题:<input name=""Topic" & Minute(Now) & Second(Now) & """ type=""text"" maxlength=""50"" emptyok=""false"" value=""""/><br/>"
			ShowPost = ShowPost & "内容:<input name=""Content" & Minute(Now) & Second(Now) & """ type=""text"" emptyok=""false"" value=""""/><br/>"
			ShowPost = ShowPost & "<anchor>OK,发表<go href=""Group.asp?Action=topicsave&amp;ID=" & ID & "&amp;" & KS.WapValue & """ method=""post"">"
			ShowPost = ShowPost & "<postfield name=""UserName"" value=""" & KSUser.UserName &"""/>"
			ShowPost = ShowPost & "<postfield name=""Topic"" value=""$(Topic" & Minute(Now) & Second(Now) & ")""/>"
			ShowPost = ShowPost & "<postfield name=""Content"" value=""$(Content" & Minute(Now) & Second(Now) & ")""/>"
			ShowPost = ShowPost & "</go></anchor><br/>"
			ShowPost = ShowPost & "仅该圈子成员可以发起主题，非成员仅可以回复<br/>"
		End Function
		
		'保存发表
		Function TopicSave()
		    Dim ID:ID=KS.Chkclng(KS.S("ID"))
			Dim Topic:Topic=KS.R(KS.S("Topic"))
			Dim Content:Content=KS.S("Content")
			IF Topic="" Then
			   TopicSave = "请输入讨论话题!<br/><anchor><prev/>返回重写</anchor><br/>"
			End If
			IF Content="" Then
			   TopicSave = "请输入讨论内容!<br/><anchor><prev/>返回重写</anchor><br/>"
			End If
			Dim RS:set RS=Server.Createobject("adodb.recordset")
			RS.Open "select * from KS_TeamTopic",Conn,1,3
			RS.Addnew
			RS("Title")=Topic
			RS("Content")=Content
			RS("TeamID")=ID
			RS("ParentID")=0
			RS("UserName")=KS.S("UserName")
			RS("Adddate")=now
			RS("UserIP")=KS.GetIP
			RS("Status")=1
			RS("IsBest")=0
			RS("IsTop")=0
			RS.Update
			RS.Close:set RS=Nothing
			TopicSave = "您的讨论话题发表成功！<br/>"
		End Function
		
		'圈子信息
		Function ShowInfo()
		    ShowInfo = "【圈子信息】<br/>"
			ShowInfo = ShowInfo &"<img src=""" & RS("PhotoUrl") & """ alt=""""/><br/>"
			'ShowInfo = ShowInfo &"圈子名称:" & RS("TeamName") & "<br/>"
			ShowInfo = ShowInfo &"创 建 者:" & RS("UserName") & "<br/>"
			ShowInfo = ShowInfo &"创建时间:" & RS("Adddate") & "<br/>"
			ShowInfo = ShowInfo &"成员人数:" & Conn.Execute("select Count(UserName)  from KS_TeamUsers where status=3 And TeamID=" & RS("ID"))(0) & "<br/>"
			ShowInfo = ShowInfo &"主题回复:" & Conn.Execute("select Count(*) from KS_TeamTopic where ParentID=0 and TeamID=" & ID )(0) & "/" & Conn.Execute("select count(*) from KS_TeamTopic where ParentID<>0 and TeamID=" & ID )(0) & "<br/>"
			ShowInfo = ShowInfo &"【管 理 员】<br/>"
			Dim RSU:set RSU=Server.Createobject("adodb.recordset")
			RSU.Open "select * from KS_User where UserName='" & RS("UserName") &"'",Conn,1,1
			If Not RSU.EOF Then
			   'Dim UserFaceSrc:UserFaceSrc=RSU("UserFace")
			   'Dim FaceWidth:FaceWidth=KS.ChkClng(RSU("FaceWidth"))
			   'Dim FaceHeight:FaceHeight=KS.ChkClng(RSU("FaceHeight"))
			   'If Ucase(Left(UserFaceSrc,4))<>"http" Then UserFaceSrc="../" & UserFaceSrc
			   'ShowInfo = ShowInfo &"<img src=""" & UserFaceSrc & """ width=""" & FaceWidth & """ height=""" & FaceHeight & """ alt=""""/><br/>"
			   ShowInfo = ShowInfo &"<a href=""index.asp?u=" & RSU("UserName") & "&amp;" & KS.WapValue & """>" & RS("UserName") & "(" & RSU("Province") & RSU("City") & ")</a><br/>"
			End If
			RSU.Close:set RSU=Nothing
		End Function
        
		Function AllDelTopic()
		    IF Cbool(KSUser.UserLoginChecked)=false Then
			   AllDelTopic = "对不起，请先登录！<br/>"
			   Exit Function
			End If
			Dim Tid:Tid=KS.Chkclng(KS.S("Tid"))
			If Tid=0 Then Response.End
			Dim RST:set RST=server.createobject("adodb.recordset")
			RST.Open "select * from KS_TeamTopic where ID=" & Tid,Conn,1,3
			If Not RST.EOF Then
			   If RST("UserName")=KSUser.UserName or KSUser.UserName=GroupAdmin Then
			      Conn.Execute("delete from KS_TeamTopic where ParentID=" & Tid & "")
				  RST.Delete
		       Else
		          RST.Close:Set RST=Nothing
				  AllDelTopic = "对不起，你没有删除的权限<br/>"
		       End If
			End If
		    RST.Close:Set RST=Nothing
			Response.Redirect KS.GetDomain&"Space/Group.asp?ID="& ID & "&" & KS.WapValue & ""
		End Function

		Function DelTopic()
		    IF Cbool(KSUser.UserLoginChecked)=false Then
			   DelTopic = "对不起，请先登录！<br/>"
			   Exit Function
			End If
			Dim Pid:Pid=KS.Chkclng(KS.S("Pid"))
			If Pid=0 Then Response.End
			Dim RST:set RST=server.createobject("adodb.recordset")
			RST.Open "select * from KS_TeamTopic where ID=" & Pid,Conn,1,3
			If Not RST.EOF Then
			   If RST("UserName")=KSUser.UserName or KSUser.UserName=GroupAdmin Then
			      RST.Delete
		       Else
		          RST.Close:Set RST=Nothing
				  DelTopic = "对不起，你没有删除的权限<br/>"
		       End If
			End If
		    RST.Close:Set RST=Nothing
			Response.Redirect KS.GetDomain&"Space/Group.asp?Action=showtopic&ID="& ID & "&Tid=" & KS.Chkclng(KS.S("Tid")) & "&" & KS.WapValue & ""
		End Function

		'置顶设置
		Sub SetTop()
		    Dim Tid:Tid=KS.Chkclng(KS.S("Tid"))
		    Dim RS:set RS=Server.Createobject("adodb.recordset")
			RS.Open "select IsTop from KS_TeamTopic where ID=" & Tid,Conn,1,3
			If Not RS.EOF Then
			   If RS(0)=1 Then
			      RS(0)=0
			   Else
			      RS(0)=1
			   End If
			   RS.Update
		    End If
		    RS.Close:set RS=Nothing
		    Response.Redirect "Group.asp?Action=showtopic&ID="& ID & "&Tid=" & Tid & "&" & KS.WapValue & ""
		End Sub
		
		'精华设置
		Sub SetBest()
		    Dim Tid:Tid=KS.Chkclng(KS.S("Tid"))
			Dim RS:set RS=Server.Createobject("adodb.recordset")
			RS.Open "select IsBest from KS_TeamTopic where ID=" & Tid,Conn,1,3
			If Not RS.EOF Then
			   If RS(0)=1 Then
			      RS(0)=0
			   Else
			      RS(0)=1
			   End If
			   RS.Update
		    End If
		    RS.Close:set RS=Nothing
		    Response.Redirect "Group.asp?Action=showtopic&ID="& ID & "&Tid=" & Tid & "&" & KS.WapValue & ""
		End Sub
		
		'圈子主题列表
		Function TeamTopic()
		    MaxPerPage =10
			If KS.S("page") <> "" Then
			   CurrentPage = KS.ChkClng(KS.G("page"))
			Else
			   CurrentPage = 1
			End If
			Dim Param:Param=" where TeamID=" & ID & " And ParentID=0"
			If KS.Chkclng(KS.S("IsBest"))=1 Then Param=Param & " And IsBest=1 "
			Dim RSObj:Set RSObj=Server.CreateObject("ADODB.RECORDSET")
			RSObj.Open "select * from KS_TeamTopic "& Param & " Order by IsTop desc,Adddate desc" ,Conn,1,1
			If RSObj.EOF And RSObj.Bof  Then
			   TeamTopic = "没有任何讨论话题! <br/>"
			Else
			   TotalPut = RSObj.RecordCount
			   If CurrentPage < 1 Then	CurrentPage = 1
			   If (CurrentPage - 1) * MaxPerPage > totalPut Then
			      If (TotalPut Mod MaxPerPage) = 0 Then
				     CurrentPage = TotalPut \ MaxPerPage
				  Else
				     CurrentPage = TotalPut \ MaxPerPage + 1
				  End If
			   End If
			   If CurrentPage >1 And (CurrentPage - 1) * MaxPerPage < TotalPut Then
			      RSObj.Move (CurrentPage - 1) * MaxPerPage
			   Else
			      CurrentPage = 1
			   End If
			   Dim I
			   Do While Not RSObj.EOF
			      If RSObj("IsTop")=1 Then TeamTopic = TeamTopic & "[顶]"
				  If RSObj("isbest")=1 Then TeamTopic = TeamTopic & "[精]"
				  TeamTopic = TeamTopic & "<a href=""Group.asp?Action=showtopic&amp;ID=" & ID & "&amp;Tid=" & RSObj("ID") & "&amp;" & KS.WapValue & """>" & ((I+1)+CurrentPage*MaxPerPage)-MaxPerPage &"." & RSObj("Title") & "(" & Conn.Execute("select Count(id) from KS_TeamTopic where ParentID=" & RSObj("ID"))(0) & ")</a><br/>"
				  'TeamTopic = TeamTopic & "作者:<a href=""Space.asp?UserName=" & RSObj("UserName") & "&amp;" & KS.WapValue & """>" & RSObj("UserName") & "</a> "
				  'TeamTopic = TeamTopic & "" & KS.DateFormat(RSObj("AddDate"),17) & "<br/>"
				  RSObj.MoveNext
				  I = I + 1
				  If I >= MaxPerPage Then Exit Do
			   Loop
			   TeamTopic = TeamTopic & KS.ShowPagePara(TotalPut, MaxPerPage, "Group.asp", True, "个", CurrentPage, "ID=" & ID & "&amp;IsBest=" & IsBest & "&amp;" & KS.WapValue & "")
			   TeamTopic = TeamTopic & "<br/>"
			End If
			RSObj.Close:Set RSObj=Nothing
		End Function
		
		'会员列表
		Function ShowUser()
		    MaxPerPage =10
			If KS.S("page") <> "" Then
			   CurrentPage = KS.ChkClng(KS.G("page"))
			Else
			   CurrentPage = 1
			End If
			Dim RSObj:set RSObj=server.createobject("adodb.recordset")
			RSObj.open "select * from KS_TeamUsers where TeamID=" &ID & " and Status=3",Conn,1,1
			If Not RSObj.EOF Then
			   TotalPut = RSObj.RecordCount
			   If CurrentPage < 1 Then	CurrentPage = 1
			   If (CurrentPage - 1) * MaxPerPage > totalPut Then
			      If (TotalPut Mod MaxPerPage) = 0 Then
				     CurrentPage = TotalPut \ MaxPerPage
				  Else
				     CurrentPage = TotalPut \ MaxPerPage + 1
				  End If
			   End If
			   If CurrentPage >1 And (CurrentPage - 1) * MaxPerPage < TotalPut Then
			      RSObj.Move (CurrentPage - 1) * MaxPerPage
			   Else
			      CurrentPage = 1
			   End If
			   Dim I
			   Do While Not RSObj.EOF
			      ShowUser = ShowUser & "<a href=""Space.asp?UserName="&RSObj("UserName")&"&amp;" & KS.WapValue & """>"&RSObj("UserName")&"</a><br/>"
				  RSObj.MoveNext
				  I = I + 1
				  If I >= MaxPerPage Then Exit Do
			   Loop
			   ShowUser = ShowUser & KS.ShowPagePara(TotalPut, MaxPerPage, "Group.asp", True, "个", CurrentPage, "Action=ShowUsers&amp;ID=" & ID & "&amp;IsBest=" & IsBest & "&amp;" & KS.WapValue & "")
			   ShowUser = ShowUser & "<br/>"
			End If
			RSObj.Close:Set RSObj=Nothing
		End Function
		
		'显示帖子列表
		Function ShowTopic()
		    Dim Tid:Tid=KS.Chkclng(KS.S("Tid"))
			Dim RS:set RS=server.createobject("adodb.recordset")
			RS.Open "select b.UserName,b.UserFace,b.UserID,a.* from KS_TeamTopic a ,KS_User b where a.UserName=b.UserName And a.ID=" &Tid,Conn,1,1
			If RS.EOF And RS.BOF Then
			   RS.Close:set RS=Nothing
			   ShowTopic = "参数传递出错!<br/>"
			   Exit Function
		    End If
			ShowTopic = "<b>"&RS("Title")&"</b><br/>"
			If KS.Chkclng(KS.S("Page"))<1 Then
			   ShowTopic = ShowTopic & "作者:<a href=""Space.asp?UserName=" & RS(0) & "&amp;" & KS.WapValue & """>" & RS(0) & "</a> " & KS.DateFormat(RS("Adddate"),17) & "<br/>"
			   Dim Content
			   Content=KS.UBBToHTML(KS.LoseHtml(KS.HTMLToUBB(KS.ReplaceTrim(KS.GetEncodeConversion(RS("Content"))))))
			   ShowTopic = ShowTopic & ""&KS.ContentPagination(Content,200,"Group.asp?Action=showtopic&amp;ID="& ID &"&amp;Tid=" & RS("ID") & "&amp;" & KS.WapValue & "",False,False)&""
			   ShowTopic = ShowTopic & "<br/>"
			End If
			If Cbool(KSUser.UserLoginChecked)=False Then
			   ShowTopic = ShowTopic & "登录后才可以参与该贴子的讨论！如要参与讨论请先<a href=""../User/Login/?../Space/Group.asp?Action=showtopic&amp;ID="&ID&"&amp;Tid="&Tid&""">免费注册登陆</a>！<br/>"
			Else
			   ShowTopic = ShowTopic & "<a href=""Group.asp?Action=replay&amp;ID="&ID&"&amp;Tid="&RS("ID")&"&amp;" & KS.WapValue & """>回复(" & Conn.Execute("select Count(ID) from KS_TeamTopic where ParentID=" & Tid)(0) & ")</a> "
			End If
			If RS(0)=KSUser.UserName or KSUser.UserName=GroupAdmin Then
			   ShowTopic = ShowTopic & "<a href=""Group.asp?Action=connectpost&amp;ID="&ID&"&amp;Tid="&RS("ID")&"&amp;" & KS.WapValue & """>续写</a> "
			   ShowTopic = ShowTopic & "<a href=""Group.asp?Action=alldeltopic&amp;ID="&ID&"&amp;Tid="&RS("ID")&"&amp;" & KS.WapValue & """>删除</a> "
			   If KSUser.UserName=GroupAdmin Then
			      If RS("istop")=1 Then
				     ShowTopic = ShowTopic & "<a href=""Group.asp?Action=settop&amp;ID="&ID&"&amp;tid="&RS("ID")&"&amp;" & KS.WapValue & """>取消置顶</a> "
				  Else
				     ShowTopic = ShowTopic & "<a href=""Group.asp?Action=settop&amp;ID="&ID&"&amp;tid="&RS("ID")&"&amp;" & KS.WapValue & """>设为置顶</a> "
				  End If
				  If RS("isbest")=1 Then
				     ShowTopic = ShowTopic & "<a href=""Group.asp?Action=setbest&amp;ID="&ID&"&amp;tid="&RS("ID")&"&amp;" & KS.WapValue & """>取消精华</a>"
				  Else
				     ShowTopic = ShowTopic & "<a href=""Group.asp?Action=setbest&amp;ID="&ID&"&amp;tid="&RS("ID")&"&amp;" & KS.WapValue & """>设为精华</a>"
				  End If
			   End If
			End If
			ShowTopic = ShowTopic & "<br/>"
			
			MaxPerPage=10
			
			CurrentPage=KS.ChkClng(KS.S("Page"))
			If CurrentPage<=0 Then CurrentPage=CurrentPage+1
			Dim RSP:set RSP=Server.Createobject("adodb.recordset")
			RSP.Open "select b.UserName,b.UserID,b.UserFace,a.* from KS_TeamTopic a, KS_User b where a.UserName=b.UserName and ParentID=" & Tid & " order by Adddate desc",Conn,1,1
			If Not RSP.EOF Then
			   TotalPut = RSP.Recordcount   
			   If CurrentPage < 1 Then	CurrentPage = 1
			   If (CurrentPage - 1) * MaxPerPage > totalPut Then
			      If (TotalPut Mod MaxPerPage) = 0 Then
				     CurrentPage = TotalPut \ MaxPerPage
				  Else
				     CurrentPage = TotalPut \ MaxPerPage + 1
				  End If
			   End If
			   If CurrentPage >1 And (CurrentPage - 1) * MaxPerPage < TotalPut Then
			      RSP.Move (CurrentPage - 1) * MaxPerPage
			   Else
			      CurrentPage = 1
			   End If
			   Do While Not RSP.EOF
			      ShowTopic = ShowTopic & KS.LoseHtml(KS.HTMLCode(RSP("Content"))) & "<br/>"
				  ShowTopic = ShowTopic & "<a href=""Space.asp?UserName=" & RSP(0) & "&amp;" & KS.WapValue & """>" & RSP(0) & "</a> " & KS.DateFormat(RSP("Adddate"),17) & ""
				  If RS(0)=KSUser.UserName or KSUser.UserName=GroupAdmin Then
				     ShowTopic = ShowTopic & "<a href=""Group.asp?Action=deltopic&amp;ID=" & ID & "&amp;Tid=" & RS("ID") & "&amp;Pid=" & RSP("ID") & "&amp;" & KS.WapValue & """>删除</a>"
				  End If
				  ShowTopic = ShowTopic & "<br/>"
			      RSP.MoveNext
				  I = I + 1
				  If I >= MaxPerPage Then Exit Do
			   Loop
			   ShowTopic = ShowTopic & KS.ShowPagePara(TotalPut, MaxPerPage, "Group.asp", True, "个", CurrentPage, "Action=showtopic&amp;ID=" & ID & "&amp;Tid=" & Tid & "&amp;" & KS.WapValue & "")
			   ShowTopic = ShowTopic & "<br/>"
			End If
			RSP.Close:set RSP=Nothing			
			RS.Close:set RS=Nothing
		End Function
		
End Class
%>