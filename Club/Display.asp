<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../Plus/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Display
KSCls.Kesion()
Set KSCls = Nothing

Class Display
        Private KS, KSR,ListStr,ID,BSetting,BoardID,Node,managestr
		Private ListTemplate,LoopTemplate,LoopList,FileContent,RST,master
		Private MaxPerPage, TotalPut , CurrentPage, TotalPage, i, j, Loopno
	    Private SqlStr
		Private Sub Class_Initialize()
		 If (Not Response.IsClientConnected)Then
			Response.Clear
			Response.End
		 End If
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		%>
		<!--#include file="../KS_Cls/Kesion.IfCls.asp"-->
		<%
		
		Public Sub Kesion()
			If KS.Setting(56)="0" Then KS.Die "��վ�ѹر�" & KS.Setting(61)
			If KS.Setting(59)="1" Then response.Redirect("guestbook.asp")
			
			CurrentPage = KS.ChkClng(Request("page"))
			If CurrentPage<=0 Then CurrentPage=1
		    ID=KS.ChkClng(KS.S("ID"))

		          If KS.Setting(114)="" Then Response.Write "���ȵ�""������Ϣ����->ģ���""����ģ��󶨲���!":response.end
				   FileContent = KSR.LoadTemplate(KS.Setting(114))
				   If Trim(FileContent) = "" Then FileContent = "ģ�岻����!"
				   FCls.RefreshType = "guestdisplay" '����ˢ�����ͣ��Ա�ȡ�õ�ǰλ�õ�����
				   FCls.RefreshFolderID = "0" '���õ�ǰˢ��Ŀ¼ID Ϊ"0" ��ȡ��ͨ�ñ�ǩ
				   ListTemplate=LFCls.GetConfigFromXML("GuestBook","/guestbook/template","display")
				   LoopTemplate=KS.CutFixContent(ListTemplate, "[loop]", "[/loop]", 0)
				   
				   Call GetSubject()
				   If BoardID<>0  Then 
				    KS.LoadClubBoard()
				    Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]")
					If Node Is Nothing Then
					 KS.Die "�Ƿ�����!"
					End If
					 BSetting=Node.SelectSingleNode("@settings").text
					 FileContent=RexHtml_IF(FileContent) '�б�ҳ�ȹ���������ǩ,���ٱ�ǩ����
				   End If
				   If KS.IsNul(BSetting) Then BSetting="1$$$0$0$0$0$0$0$0$0$$$$$$$$$$$$$$$$"
				 	BSetting=Split(BSetting,"$")
				   If  BSetting(0)="0" And KS.C("UserName")="" Then
					ListTemplate=LFCls.GetConfigFromXML("GuestBook","/guestbook/template","error1")
				   End If

				   
				   
					select case KS.S("Action")
					  case "settop" Call SetTOP
					  case "setbest" Call SetBest
					  case "canceltop" Call CancelTop
					  case "cancelbest" Call CancelBest
					  case "delsubject" Call delsubject
					  case "delreply" Call delreply
					  case "verify" Call verify
					  case "locked" Call Locked
					  case "unlocked" call unlocked
					  case "replylock" call replylock
					  case "replyunlock" call replyunlock
					  case "movetopic" call movetopic
					End select
				   
				   Call GetReplayList()
				   Call GetReplayForm()
	               ListTemplate = Replace(ListTemplate,"[loop]" & LoopTemplate &"[/loop]",LoopList)
				   ListTemplate = Replace(ListTemplate,"{$ManageMenu}",managestr)

				   FileContent=Replace(FileContent,"{$GetGuestList}",ListTemplate)
				   FileContent=Replace(FileContent,"{$Subject}",RST("Subject"))
				   FileContent=Replace(FileContent,"{$GuestTitle}",RST("Subject"))
				   FileContent=Replace(FileContent,"{$TopicID}",RST("ID"))
				   FileContent=Replace(FileContent,"{$BoardID}",RST("BoardID"))
				   FileContent=Replace(FileContent,"{$Hits}",RST("Hits"))
				   RST.Close:Set RST=Nothing
				   FileContent=Replace(FileContent,"{$Page}",currentpage)
				   FileContent=Replace(FileContent,"{$PageStr}",PageList())
				   FileContent=KSR.KSLabelReplaceAll(FileContent)
				   KS.Echo RexHtml_IF(FileContent)
		End Sub
		
		Sub GetSubject()
		  Dim UserInfo,LC,Sign,UN,KSUser,LoginTF
		  Set RST=Server.CreateObject("ADODB.RECORDSET")
		  RST.Open "Select top 1 * From KS_GuestBook Where ID=" & ID,conn,1,3
		  If RST.Eof Then
		   RST.Close:Set RST=Nothing
		   Response.Write("<script>alert('�Ƿ�������');window.close();</script>")
		   Response.End
		  End If
		  If RST("Verific")=0 Then
		   RST.Close:Set RST=Nothing
		   Response.Write("<script>alert('�Բ���,�����ӻ�û����ˣ�');history.back();</script>")
		   Response.End
		  End If
		  RST("Hits")=RST("Hits")+1
		  RST.Update
		  FCls.RefreshFolderID = RST("BoardID")
		  BoardID=FCls.RefreshFolderID
		  FileContent=Replace(FileContent,"{$PostBoardID}","?bid=" & FCls.RefreshFolderID)
		  master=LFCls.GetSingleFieldValue("select master from ks_guestboard where id=" & KS.ChkClng(FCls.RefreshFolderID))
		  If CurrentPage<>1 then Exit Sub
		  LC=LoopTemplate
		  LC=Replace(LC,"{$UserName}",KS.HtmlCode(RST("UserName")))
		  LC=Replace(LC,"{$Subject}",KS.HtmlCode(RST("Subject"))&"<img src='../images/face1/" & RST("txthead") & "'>")
		  LC=replace(LC,"{$Hits}",RST("hits"))
		  LC=replace(LC,"{$PubTime}",RST("AddTime"))
		  If RST("ShowIP")="0" And KS.C("AdminName")="" and Check=false and rst("username")<>KS.C("UserName") Then
		  LC=Replace(LC,"{$PubIP}","***")
		  Else
		  LC=Replace(LC,"{$PubIP}",RST("guestip"))
		  End If
		  Dim Content,ReplyContent,rept,rsp
		  If RST("ShowScore")<=0 or KS.C("AdminName")<>"" Then
		    Content=KS.CheckScript(KS.HtmlCode(RST("memo")))
		  Else
		    Set KSUser=New UserCls
			LoginTF=KSUser.UserLoginChecked
			If LoginTF=false Then
		    Content="<div style=""margin : 10px 20px; border : 1px solid #efefef; padding : 15px;background : #ffffee; line-height : normal;"">�Բ�������û�е�¼�����ȵ�¼������Ҫ����ִﵽ<font color=red>" & RST("ShowScore") & "</font>�ֲ��ܲ鿴��</div>"
			ElseIf Cint(KSUser.Score)<Cint(RST("ShowScore")) Then
		    Content="<div style=""margin : 10px 20px; border : 1px solid #efefef; padding : 15px;background : #ffffee; line-height : normal;"">�Բ������Ļ��ֲ��㣡����Ҫ����ִﵽ<font color=red>" & RST("ShowScore") & "</font>�ֲ��ܲ鿴,����ǰ���û���Ϊ<font color=green>" & KSUser.Score &"</font>�֣�</div>"
		    Else
		    Content=KS.CheckScript(KS.HtmlCode(RST("memo")))
			End If
		  End If
		  Content=bbimg(content)
		  if rst("verific")=2 then
		   content="<div style=""margin : 10px 20px; border : 1px solid #efefef; padding : 15px;background : #ffffee; line-height : normal;"">�����ѱ�������</div>"
		  end if
		  If Instr(Content,"[post]")<>0 Then
		   rept=0
		   Set KSUser=New UserCls
		   If Cbool(KSUser.UserLoginChecked)=true Then 
		    set rsp=conn.execute("select id from ks_guestreply where topicid=" & id & " and username='" & KS.C("UserName") & "'")
			if not rsp.eof then
			  rept=1
			end if
			if check=true or ks.c("adminname")<>"" or ksuser.username=rst("username") then rept=1
		   End If
		   
		   if rept=1 then
		    ReplyContent="<div style=""margin : 10px 20px; border : 1px solid #efefef; padding : 5px;background : #ffffee; line-height : normal;""><font color=""gray"">��������ֻ��<b>�ظ�</b>��ſ������</font><hr color='#ff6600' size='1'><br/>" & KS.CutFixContent(content, "[post]", "[/post]", 0) & "</div>"
		   else
		    ReplyContent="<div style=""margin : 5px 20px; border : 1px solid #efefef; padding : 15px;background : #ffffee; line-height : normal;""><font color=""red"">��������ֻ��<b>�ظ�</b>��ſ������</font></div>"
		   end if
		   content=replace(content,KS.CutFixContent(content, "[post]", "[/post]", 1),ReplyContent)
		  End If
		  LC=Replace(LC,"{$Content}",Content)
		  UserInfo="<img src='../images/Face/" & RST("Face") &"' border='0'>"
		  Dim RSU:Set RSU=Conn.Execute("Select top 1 UserName,UserFace,Sign,Score,GradeTitle,LoginTimes,RegDate,email,qq From KS_User Where UserName='" & RST("UserName") &"'")
		  
		  If Not RSU.Eof Then
			  Dim UserXml:set UserXml=KS.RSToXml(rsU,"row","")
			  Set UN=UserXml.DocumentElement.SelectSingleNode("row")
		  Else 
		      Set UN=Nothing
		  End If
		  rsu.close
		  Set rsu=Nothing
		  
		   If UN Is Nothing  Then
			  	  UserInfo="<img src='../Images/Face/0.gif' width='82' height='90'>"
			      UserInfo=UserInfo & "<div style='height:26px;margin-top:10px;text-align:center'>�� �����ο�</div>"
			      Sign=""
		   Else
			   Dim UserFaceSrc:UserFaceSrc=UN.SelectSingleNode("@userface").text
			   if lcase(left(userfacesrc,4))<>"http" then userfacesrc="../" & userfacesrc
			   If RST("ShowSign")="1" Then
               Sign=UN.SelectSingleNode("@sign").text
			   End If
			   UserInfo="<div style='margin-top:5px;'><a href='../space/?" & UN.SelectSingleNode("@username").text & "' target='_blank' style='border:1px solid #ccc;padding:1px;'><img src='" & UserFaceSrc &"' width='82' height='90' border='0'></a></div>"
			   UserInfo=UserInfo & "<div style='height:35px;line-height:35px;text-align;center'><img src='../images/user/log/106.gif'><a href='javascript:void(0)' onclick=""addF(event,'" & UN.SelectSingleNode("@username").text & "')"">��Ϊ����</a> <img src='../images/user/mail.gif'><a href='javascript:void(0)' onclick=""sendMsg(event,'" & UN.SelectSingleNode("@username").text & "')"">������Ϣ</a></div>"
			   UserInfo=UserInfo & "<div style='margin-top:10px;height:26px;padding-left:5px;text-align:left'>����:" & UN.SelectSingleNode("@gradetitle").text
			   if  KS.FoundInArr(master, UN.SelectSingleNode("@username").text, ",")=true then UserInfo=UserInfo &"<font color=red>(����)</font>"
			   UserInfo=UserInfo &"</div>"
			   UserInfo=UserInfo & "<div style='height:26px;padding-left:5px;text-align:left'>�û�����:" & UN.SelectSingleNode("@score").text &" ��</div>"
			   UserInfo=UserInfo & "<div style='height:26px;padding-left:5px;text-align:left'>��¼����:" & UN.SelectSingleNode("@logintimes").text &" ��</div>"
			   UserInfo=UserInfo & "<div style='height:26px;padding-left:5px;text-align:left'>ע��ʱ��:" & UN.SelectSingleNode("@regdate").text &"</div>"
			   
			     ListStr = " <a href='../space/?" & UN.SelectSingleNode("@username").text & "' target='_blank'><img src='images/home.gif' width='16' height='16' border='0' align='absmiddle' alt='������ҳ'></a>��ҳ"
				 ListStr = ListStr & "  |" 
				 If UN.SelectSingleNode("@email").text <> "" Then
			   ListStr = ListStr & "  <a href='mailto:" & UN.SelectSingleNode("@email").text & "' target='_blank'><img src='images/email.gif' width='18' height='18' border='0' align='absmiddle' alt='�����ʼ�:[ " & UN.SelectSingleNode("@email").text &" ]'></a>�ʼ�" & vbcrlf
				 Else
			   ListStr = ListStr & "  <a href='#'><img src='images/email-gray.gif' width='18' height='18' border='0' align='absmiddle' alt='�����ʼ�'></a>�ʼ�" & vbcrlf
				End If
				 ListStr = ListStr & "  |" 
				If UN.SelectSingleNode("@qq").text <> "" and UN.SelectSingleNode("@qq").text <> "0" Then
				ListStr = ListStr & " <a href='#'><img src='images/qq.gif' width='16' height='16' border='0' align='absmiddle' alt='QQ����:[ " & UN.SelectSingleNode("@qq").text & " ]'></a>QQ����"
				Else
				ListStr = ListStr & "  <a href='#'><img src='images/qq-gray.gif' width='16' height='16' border='0' align='absmiddle' alt='QQ����'></a>QQ����" & vbcrlf
				End If
			  
			  End If
			  
		  LC=Replace(LC,"{$UserInfo}",UserInfo)
		  LC=Replace(LC,"{$TopicID}",ID)
		  LC=Replace(LC,"{$ShowRightAd}",GetAdByRnd(36))
		  LC=Replace(LC,"{$ShowBottomAd}",GetAdByRnd(37))
		  LC=Replace(LC,"{$UserMenu}",liststr)
		  dim setstr:setstr="<a href='#reply' onclick=""reply(1,'" & RST("UserName") & "','" & RST("AddTime") & "')"">����</a> | <a href='#reply' >�ظ�</a> | "
		  setstr=setstr & "<a href='javascript:edit(1,1," & ID & ");'>�༭����</a> | <a href='?id=" & ID & "&action=delsubject' onclick=""return(confirm('ɾ�����⣬���еĻظ���ɾ����ȷ��ִ��ɾ��������'))"">ɾ������</a>"
		  LC=Replace(LC,"{$ManageMenu}",setstr)
		  If rst("isbest")=0 Then
		   	LC=Replace(LC,"{$Jing}","")
          Else
		    LC=Replace(LC,"{$Jing}","<div style='border:1px solid #aaaaaa;color:red;width:160px;background:#f1f1f1'><img src='images/jing.gif' align='absmiddle'>�������϶�Ϊ����</div>")
		  End If
		  LC=Replace(LC,"{$N}","1")
		  LC=Replace(LC,"{$UserSign}",Sign)
		  
		  if rst("verific")=1 then
		    managestr="<dl><a href=""?action=locked&id="&id &""">��������</a></dl>"
		  else
		    managestr="<dl><a href=""?action=unlocked&id=" & id & """>�������</a></dl>"
		  end if
		  managestr=managestr & "<dl><a href=""?action=delsubject&id=" & id & """  onclick=""return(confirm('ɾ�����⣬���еĻظ���ɾ����ȷ��ִ��ɾ��������'))"">ɾ������</a></dl>"
		  managestr=managestr & "<dl><a href=""javascript:void(0)"" onclick=""movetopic(event," & id & ",'" & rst("subject") & "')"">�ƶ�����</a></dl>"
		  if rst("istop")=1 then
		  managestr=managestr & "<dl><a href='?id=" & ID &"&action=canceltop' onclick=""return(confirm('ȷ��ȡ���ö���'))"">ȡ���ö�</a></dl>"
		  else
		  managestr=managestr & "<dl><a href='?id=" & ID &"&action=settop' onclick=""return(confirm('ȷ����Ϊ�ö���'))"">��Ϊ�ö�</a></dl>"
		  end if
		  if rst("isbest")=1 then
		  managestr=managestr & "<dl><a href='?id=" & ID &"&action=cancelbest' onclick=""return(confirm('ȷ��ȡ��������'))"">ȡ������</a></dl>"
		  else
		  managestr=managestr & "<dl><a href='?id=" & ID &"&action=setbest' onclick=""return(confirm('ȷ����Ϊ������'))"">��Ϊ����</a></dl>"
		  end if

		  LoopList=LC
		End Sub
		
		Sub GetReplayForm()
		 Dim ReplayForm:ReplayForm=LFCls.GetConfigFromXML("GuestBook","/guestbook/template","replayform")
		 LoopList=LoopList & ReplayForm
		End Sub
		
		
		Sub GetReplayList()
		 MaxPerPage=10
		 SqlStr = "SELECT * From KS_GuestReply WHERE topicid=" & KS.ChkClng(KS.S("ID")) & " ORDER BY ID" 
		 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		 RS.Open SqlStr,conn,1,1
		  IF RS.Eof And RS.Bof Then 
			  RS.Close:Set RS=Nothing
			  totalput=0
			  exit sub
		  Else
							TotalPut= RS.RecordCount
							If CurrentPage < 1 Then CurrentPage = 1
		
							If CurrentPage >1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
									RS.Move (CurrentPage - 1) * MaxPerPage
							Else
									CurrentPage = 1
							End If
							Dim Xml:Set XML=KS.ArrayToXml(RS.GetRows(MaxPerPage),rs,"row","")
							RS.Close:Set RS=Nothing
							If IsObject(Xml) Then
							Call GetTopicList(XML)
							End If
			End IF
			
		End Sub
		
		Sub GetTopicList(Xml)
		     Dim I,LC,UserInfo,n,liststr,Sign,Node,UserXml,UserNames
		     If CurrentPage=1 Then N=1 Else N=MaxPerPage*(CurrentPage-1)
			 For Each Node In Xml.DocumentElement.SelectNodes("row")
			    If UserNames="" Then
				 UserNames="'" & Node.SelectSingleNode("@username").text & "'"
				Else
				 UserNames=UserNames & ",'" & Node.SelectSingleNode("@username").text & "'"
				End If
			 Next
			 Dim RS:Set RS=Conn.Execute("Select top " & MaxPerPage & " UserName,UserFace,Sign,Score,GradeTitle,LoginTimes,RegDate,email,qq From KS_User Where UserName in(" & UserNames & ")")
			 If Not RS.Eof Then Set UserXml=KS.RsToXml(RS,"row","")
			 RS.Close:Set RS=Nothing
			 
		 	 For Each Node In Xml.DocumentElement.SelectNodes("row")
			  LC=LoopTemplate
			  LC=replace(LC,"{$Subject}","<img src='../images/Face1/face" & Node.SelectSingleNode("@txthead").text &".gif' border='0'>")
			  LC=replace(LC,"{$PubTime}",Node.SelectSingleNode("@replaytime").text)
			  If Node.SelectSingleNode("@showip").text="0" And KS.C("AdminName")="" and Check=false and rst("username")<>KS.C("UserName") then
			  LC=replace(LC,"{$PubIP}","***")
			  Else
			  LC=replace(LC,"{$PubIP}",Node.SelectSingleNode("@userip").text)
			  End If
			  Dim Content,UN
			      if Node.SelectSingleNode("@verific").text="2" then
				    if check=true or ks.c("adminname")<>"" then
					 Content="<div style=""margin-left:20px; margin-top: 15px; background-color: #ffffee; border: 1px solid #f9c943; width: 350px; height: 30px;line-height:30px;"">����Ϣ������,�������ǰ���/����Ա���Կ��Կ�������Ϣ.</div>" & KS.HtmlCode(Node.SelectSingleNode("@content").text)
					else
					 Content="<div style=""margin-left:20px; margin-top: 15px; background-color: #ffffee; border: 1px solid #f9c943; width: 300px; height: 30px;line-height:30px;"">�Բ��𣬸���Ϣ������!</div>"
					end if
				  elseif Node.SelectSingleNode("@verific").text="1" then
				   Content=KS.HtmlCode(Node.SelectSingleNode("@content").text)
				  else
				   if check=true  then
					 Content="<div style=""margin-left:20px; margin-top: 15px; background-color: #ffffee; border: 1px solid #f9c943; width: 300px; height: 30px;line-height:30px;"">����Ϣδ���,�������ǰ������Կ��Կ�������Ϣ.</div>" & KS.HtmlCode(Node.SelectSingleNode("@content").text)
				   ElseIf KS.C("AdminName")<>"" Then
					 Content="<div style=""margin-left:20px; margin-top: 15px; background-color: #ffffee; border: 1px solid #f9c943; width: 300px; height: 30px;line-height:30px;"">����Ϣδ���,�������ǹ���Ա���Կ��Կ�������Ϣ.</div>" & KS.HtmlCode(Node.SelectSingleNode("@content").text)
				    Else
					Content="<div style=""margin-left:20px; margin-top: 15px; background-color: #ffffee; border: 1px solid #f9c943; width: 300px; height: 50px;line-height:50px; "">��վ������˻���,����Ϣδͨ�����!</div>"
				   End If
				 end if
			   
			  LC=replace(LC,"{$Content}",bbimg(Content))
			  LC=replace(LC,"{$UserName}",Node.SelectSingleNode("@username").text)
			  If IsObject(UserXML) Then
			   set UN=UserXml.DocumentElement.SelectSingleNode("row[@username='" & Node.SelectSingleNode("@username").text & "']")
			  Else
			   Set UN=Nothing
			  End If
			  If UN Is Nothing Then
			  	  UserInfo="<img src='../Images/Face/0.gif' width='82' height='90'>"
			      UserInfo=UserInfo & "<div style='height:26px;padding-left:5px;margin-top:10px;text-align:left'>�� �� �飺�ο�</div>"
			      Sign=""
			  Else
			  
			   Dim UserFaceSrc:UserFaceSrc=UN.SelectSingleNode("@userface").text
			   if lcase(left(userfacesrc,4))<>"http" then userfacesrc="../" & userfacesrc
			   If Node.SelectSingleNode("@showsign").text="0" Then
               Sign=UN.SelectSingleNode("@sign").text
			   End If
			   UserInfo="<div style='margin-top:5px;'><a href='../space/?" & UN.SelectSingleNode("@username").text & "' target='_blank' style='border:1px solid #ccc;padding:1px;'><img src='" & UserFaceSrc &"' width='82' height='90' border='0'></a></div>"
			   UserInfo=UserInfo & "<div style='height:35px;line-height:35px;text-align;center'><img src='../images/user/log/106.gif'><a href='javascript:void(0)' onclick=""addF(event,'" & UN.SelectSingleNode("@username").text & "')"">��Ϊ����</a> <img src='../images/user/mail.gif'><a href='javascript:void(0)' onclick=""sendMsg(event,'" & UN.SelectSingleNode("@username").text & "')"">������Ϣ</a></div>"
			   UserInfo=UserInfo & "<div style='margin-top:10px;height:26px;padding-left:5px;text-align:left'>����:" & UN.SelectSingleNode("@gradetitle").text
			   if  KS.FoundInArr(master, UN.SelectSingleNode("@username").text, ",")=true then UserInfo=UserInfo &"<font color=red>(����)</font>"
			   UserInfo=UserInfo &"</div>"
			   UserInfo=UserInfo & "<div style='height:26px;padding-left:5px;text-align:left'>�û�����:" & UN.SelectSingleNode("@score").text &" ��</div>"
			   UserInfo=UserInfo & "<div style='height:26px;padding-left:5px;text-align:left'>��¼����:" & UN.SelectSingleNode("@logintimes").text &" ��</div>"
			   UserInfo=UserInfo & "<div style='height:26px;padding-left:5px;text-align:left'>ע��ʱ��:" & UN.SelectSingleNode("@regdate").text &"</div>"
			   
			     ListStr = " <a href='../space/?" & UN.SelectSingleNode("@username").text & "' target='_blank'><img src='images/home.gif' width='16' height='16' border='0' align='absmiddle' alt='������ҳ'></a>��ҳ"
				 ListStr = ListStr & "  |" 
				 If UN.SelectSingleNode("@email").text <> "" Then
			   ListStr = ListStr & "  <a href='mailto:" & UN.SelectSingleNode("@email").text & "' target='_blank'><img src='images/email.gif' width='18' height='18' border='0' align='absmiddle' alt='�����ʼ�:[ " & UN.SelectSingleNode("@email").text &" ]'></a>�ʼ�" & vbcrlf
				 Else
			   ListStr = ListStr & "  <a href='#'><img src='images/email-gray.gif' width='18' height='18' border='0' align='absmiddle' alt='�����ʼ�'></a>�ʼ�" & vbcrlf
				End If
				 ListStr = ListStr & "  |" 
				If UN.SelectSingleNode("@qq").text <> "" and UN.SelectSingleNode("@qq").text <> "0" Then
				ListStr = ListStr & " <a href='#'><img src='images/qq.gif' width='16' height='16' border='0' align='absmiddle' alt='QQ����:[ " & UN.SelectSingleNode("@qq").text & " ]'></a>QQ����"
				Else
				ListStr = ListStr & "  <a href='#'><img src='images/qq-gray.gif' width='16' height='16' border='0' align='absmiddle' alt='QQ����'></a>QQ����" & vbcrlf
				End If
			  
			  End If
			  
			 
			  n=n+1
			  LC=Replace(LC,"{$UserInfo}",UserInfo)
			  LC=Replace(LC,"{$UserMenu}",liststr)
			  Dim ManageMenu:ManageMenu=""
			  If Node.SelectSingleNode("@verific").text="1" Then
			  ManageMenu="<a href='#reply' onclick=""reply("&n&",'" & Node.SelectSingleNode("@username").text & "','" & Node.SelectSingleNode("@replaytime").text & "')"">����</a> | "
              Else
			  ManageMenu="<a href='?action=verify&id=" & ID & "&replyid=" &Node.SelectSingleNode("@id").text &"' onclick=""return(confirm('ȷ����˸ûظ���?'));"">���</a> | "
			  End If
			  If Node.SelectSingleNode("@verific").text="1" Then
			  ManageMenu=ManageMenu & "<a href='?action=replylock&id=" & ID & "&replyid=" & Node.SelectSingleNode("@id").text & "' onclick=""return(confirm('ȷ�����θûظ���?'));"">����</a> | "
			  Else
			  ManageMenu=ManageMenu & "<a href='?action=replyunlock&id=" & ID & "&replyid=" & Node.SelectSingleNode("@id").text & "' onclick=""return(confirm('ȷ��ȡ�����θûظ���?'));"">����</a> | "
			  End If
			  ManageMenu=ManageMenu & "<a href='#reply' >�ظ�</a> | <a href='javascript:edit(0," & N & "," & Node.SelectSingleNode("@id").text & ");'>�༭</a> | <a onclick='return(confirm(""ȷ��ɾ���ûظ���""))' href='?action=delreply&id=" & ID & "&replyid=" & Node.SelectSingleNode("@id").text &"'>ɾ��</a>"
			  
		      LC=Replace(LC,"{$ManageMenu}",ManageMenu)
			  LC=Replace(LC,"{$Jing}","")
		      LC=Replace(LC,"{$N}",n)
			  LC=Replace(LC,"{$UserSign}",Sign)
		      LC=Replace(LC,"{$ShowRightAd}",GetAdByRnd(36))
		      LC=Replace(LC,"{$ShowBottomAd}",GetAdByRnd(37))
			  
			  LoopList=LoopList & LC
	         I=I+1
			Next

		End Sub
		
		
     Function PageList()
		PageList= "<table width=""100%"" aling=""center""><tr><td align=right>" & KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,false) & "</td></tr></table>"
	
	 End Function
	 
	 Private Function bbimg(strText)
		Dim s,re
        Set re=new RegExp
        re.IgnoreCase =true
        re.Global=True
		s=strText
		re.Pattern="<img(.[^>]*)([/| ])>"
		s=re.replace(s,"<img$1/>")
		re.Pattern="<img(.[^>]*)/>"
		s=re.replace(s,"<img$1 onload=""if (this.width>400) this.width=400;"" onclick=""window.open(this.src)"" style=""cursor:pointer""/>")
		bbimg=s
	End Function
	 
	 Sub SetBest()
		If cbool(check)=false Then
		  Response.Write"<script>alert('�Բ�����û�����õ�Ȩ��!');history.back();</script>"
		  Exit Sub
		End If
		Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select UserName,isbest,boardid,subject From KS_GuestBook Where ID=" & ID,conn,1,3
		If Not RS.Eof Then
		  rs(1)=1
		  rs.update
		  boardid=rs(2)
		  if boardid<>0 and not KS.ISNul(rs(0)) then
		     KS.LoadClubBoard()
			 Dim Node:Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]")
			 BSetting=Node.SelectSingleNode("@settings").text
			 If Not KS.IsNul(BSetting) Then
			   If KS.ChkClng(Split(BSetting,"$")(6))>0 Then
			    Call KS.ScoreInOrOut(rs(0),1,KS.ChkClng(Split(BSetting,"$")(6)),"ϵͳ","����̳��������[" & rs(3) & "]�����óɾ���!",0,0)
			   End If
			 End If
		  end if
		End If
		rs.close:set rs=nothing
		Response.Redirect request.servervariables("http_referer")
	 End Sub
	 Sub SetTop()
		If check=false Then
		  Response.Write"<script>alert('�Բ�����û�����õ�Ȩ��!');history.back();</script>"
		  Exit Sub
		End If
		Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select UserName,istop,boardid,subject From KS_GuestBook Where ID=" & ID,conn,1,3
		If Not RS.Eof Then
		  rs(1)=1
		  rs.update
		  boardid=rs(2)
		  if boardid<>0 and not KS.ISNul(rs(0)) then
		     KS.LoadClubBoard()
			 Dim Node:Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]")
			 BSetting=Node.SelectSingleNode("@settings").text
			 If Not KS.IsNul(BSetting) Then
			   If KS.ChkClng(Split(BSetting,"$")(5))>0 Then
			    Call KS.ScoreInOrOut(rs(0),1,KS.ChkClng(Split(BSetting,"$")(5)),"ϵͳ","����̳��������[" & rs(3) & "]�����ó��ö�!",0,0)
			   End If
			 End If
		  end if
		End If
		rs.close:set rs=nothing
		Response.Redirect request.servervariables("http_referer")
	 End Sub
	 Sub CancelBest()
		If cbool(check)=false Then
		  Response.Write"<script>alert('�Բ�����û�����õ�Ȩ��!');history.back();</script>"
		  Exit Sub
		End If
        Conn.Execute("Update KS_Guestbook set isbest=0 where id=" & ID)
		Response.Redirect request.servervariables("http_referer")
	 End Sub
	 Sub CancelTop()
		If check=false Then
		  Response.Write"<script>alert('�Բ�����û�����õ�Ȩ��!');history.back();</script>"
		  Exit Sub
		End If
        Conn.Execute("Update KS_Guestbook set istop=0 where id=" & ID)
		Response.Redirect request.servervariables("http_referer")
	 End Sub
	 
	 Sub delsubject()
		If check=false Then
		  Response.Write"<script>alert('�Բ�����û�����õ�Ȩ��!');history.back();</script>"
		  Exit Sub
		End If
		Dim TodayNum:TodayNum=0
		Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select top 1 UserName,boardid,subject,AddTime From KS_GuestBook Where ID=" & ID,conn,1,1
		If Not RS.Eof Then
		  boardid=rs(1)
		  If DateDiff("d",rs(3),Now)=0 Then
		   TodayNum=1
		  End If
		  If boardid<>0 then 
		    KS.LoadClubBoard()
			 Dim Node:Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]")
			 Dim LastPost,LastPostArr:LastPostArr=Split(Node.SelectSingleNode("@lastpost").text,"$")
			 
			 '������ҳ����������
			 If KS.ChkClng(LastPostArr(0))=ID Then
			   Dim RSNew:Set RSNew=Conn.Execute("Select top 1 ID,BoardID,Subject,AddTime From KS_GuestBook Where BoardID=" & boardid & " and verific=1 and id<>" & id & " order by id desc")
			   If Not RSNew.Eof Then
			     LastPost=RSNew(0) & "$" & RSNew(3) & "$" & Replace(left(RSNew(2),200),"$","") & "$$$$$$$$"
			   Else
			     LastPost="��$��$��$$$$$$$$"
			   End If
			   Conn.Execute("Update KS_GuestBoard Set LastPost='" & LastPost & "' Where ID=" & BoardID)
			   Node.SelectSingleNode("@lastpost").text=LastPost
			 End If
		  end if
		  
		  if not KS.ISNul(rs(0)) then
			 BSetting=Node.SelectSingleNode("@settings").text
			 If Not KS.IsNul(BSetting) Then
			   If KS.ChkClng(Split(BSetting,"$")(7))>0 Then
			    Call KS.ScoreInOrOut(rs(0),2,KS.ChkClng(Split(BSetting,"$")(7)),"ϵͳ","����̳�����������[" & rs(2) & "]��ɾ��!",0,0)
			   End If
			 End If
		  end if
		  
		  Dim Num,replyNum:replyNum=Conn.Execute("Select count(id) from ks_guestreply where topicid=" & id)(0)
		  TodayNum=TodayNum+Conn.Execute("Select count(id) from ks_guestreply where topicid=" & id &" and datediff(" & DataPart_D & ",ReplayTime," & SqlNowString&")=0")(0)
		  Dim Doc:set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		  Doc.async = false
		  Doc.setProperty "ServerHTTPRequest", true 
		  Doc.load(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
		  Dim XMLDate:XMLDate=doc.documentElement.attributes.getNamedItem("date").text
		  Num=KS.ChkClng(doc.documentElement.attributes.getNamedItem("todaynum").text)-TodayNum
		  If Num<0 Then Num=0
          doc.documentElement.attributes.getNamedItem("todaynum").text=Num
		  Num=KS.ChkClng(doc.documentElement.attributes.getNamedItem("postnum").text)-replyNum
		  If Num<0 Then Num=0
		  doc.documentElement.attributes.getNamedItem("postnum").text=Num
		  Num=KS.ChkClng(doc.documentElement.attributes.getNamedItem("topicnum").text)-1
		  If Num<0 Then Num=0
		  doc.documentElement.attributes.getNamedItem("topicnum").text=Num
		  
		  Conn.Execute("Update KS_GuestBoard Set TodayNum=TodayNum-" & TodayNum & " where id=" &boardid &" and todaynum>=" & TodayNum)
		  Conn.Execute("Update KS_GuestBoard Set PostNum=PostNum-" & replyNum -1& " where id=" &boardid &" and PostNum>=" & replyNum-1)
		  Application(KS.SiteSN&"_ClubBoard").DocumentElement.selectSingleNode("row[@id=" & boardid & "]/@postnum").text=Conn.Execute("Select PostNum From KS_GuestBoard Where id=" & boardid)(0)
		  Application(KS.SiteSN&"_ClubBoard").DocumentElement.selectSingleNode("row[@id=" & boardid & "]/@todaynum").text=Conn.Execute("Select TodayNum From KS_GuestBoard Where id=" & boardid)(0)

		  doc.save(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
		End If
		rs.close:set rs=nothing
        Conn.Execute("delete from KS_Guestbook where id=" & ID)
		Conn.Execute("delete from ks_guestreply where TopicID=" & ID)
		Response.Redirect "index.asp?boardid=" & FCls.RefreshFolderID
	 End Sub
	 
	 Sub delreply()
		If check=false Then
		  Response.Write"<script>alert('�Բ�����û�����õ�Ȩ��!');history.back();</script>"
		  Exit Sub
		End If
		Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open "Select UserName,boardid,subject,TotalReplay From KS_GuestBook Where ID=" & ID,conn,1,3
		If Not RS.Eof Then
		  if rs(3)>0 then 
		    rs(3)=rs(3)-1
			rs.update
		  end if
		  boardid=rs(1)
		  
		  
		  Dim ReplayTime:ReplayTime=Conn.Execute("Select top 1 ReplayTime From ks_guestreply where ID=" & KS.ChkClng(KS.S("ReplyID")))(0)
		  '����������
		  Dim Num
		  Dim Doc:set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
		  Doc.async = false
		  Doc.setProperty "ServerHTTPRequest", true 
		  Doc.load(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
		  Dim XMLDate:XMLDate=doc.documentElement.attributes.getNamedItem("date").text
		  If DateDiff("d",xmldate,ReplayTime)=0 Then
		    Conn.Execute("Update KS_GuestBoard Set TodayNum=TodayNum-1 where id=" &boardid &" and todaynum>0")
		    Num=KS.ChkClng(doc.documentElement.attributes.getNamedItem("todaynum").text)-1
			If Num<0 Then Num=0
		    doc.documentElement.attributes.getNamedItem("todaynum").text=Num
			
			Application(KS.SiteSN&"_ClubBoard").DocumentElement.selectSingleNode("row[@id=" & boardid & "]/@todaynum").text=Conn.Execute("Select TodayNum From KS_GuestBoard Where id=" & boardid)(0)
          End If
		    Conn.Execute("Update KS_GuestBoard Set PostNum=PostNum-1 where id=" &boardid &" and PostNum>0")
		    Num=KS.ChkClng(doc.documentElement.attributes.getNamedItem("postnum").text)-1
			If Num<0 Then Num=0
		    doc.documentElement.attributes.getNamedItem("postnum").text=Num
			doc.save(Server.MapPath(KS.Setting(3)&"Config/guestbook.xml"))
			Application(KS.SiteSN&"_ClubBoard").DocumentElement.selectSingleNode("row[@id=" & boardid & "]/@postnum").text=Conn.Execute("Select PostNum From KS_GuestBoard Where id=" & boardid)(0)

		  if boardid<>0 and not KS.ISNul(rs(0)) then
		     KS.LoadClubBoard()
			 Dim Node:Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & boardid &"]")
			 BSetting=Node.SelectSingleNode("@settings").text
			 If Not KS.IsNul(BSetting) Then
			   If KS.ChkClng(Split(BSetting,"$")(8))>0 Then
			    Call KS.ScoreInOrOut(rs(0),2,KS.ChkClng(Split(BSetting,"$")(8)),"ϵͳ","����̳������[" & rs(2) & "]�Ļظ���ɾ��!",0,0)
			   End If
			 End If
		  end if

		End If
		rs.close:set rs=nothing
		
		Conn.Execute("delete from ks_guestreply where ID=" & KS.ChkClng(KS.S("ReplyID")))
		Response.Redirect request.servervariables("http_referer")
	 End Sub
	 
	 sub verify()
		If check=false Then
		  Response.Write"<script>alert('�Բ�����û�����õ�Ȩ��!');history.back();</script>"
		  Exit Sub
		End If
		Conn.Execute("update ks_guestreply set verific=1 where ID=" & KS.ChkClng(KS.S("ReplyID")))
		Response.Redirect request.servervariables("http_referer")
	 end sub
	
	 sub Locked()
		If check=false Then
		  Response.Write"<script>alert('�Բ�����û��������Ȩ��!');history.back();</script>"
		  Exit Sub
		End If
		Conn.Execute("update ks_guestbook set verific=2 where ID=" & KS.ChkClng(KS.S("ID")))
		Response.Redirect request.servervariables("http_referer")
	 end sub
	 sub unlocked()
		If check=false Then
		  Response.Write"<script>alert('�Բ�����û��������Ȩ��!');history.back();</script>"
		  Exit Sub
		End If
		Conn.Execute("update ks_guestbook set verific=1 where ID=" & KS.ChkClng(KS.S("ID")))
		Response.Redirect request.servervariables("http_referer")
	 end sub
	 sub replyLock()
		If check=false Then
		  Response.Write"<script>alert('�Բ�����û��������Ȩ��!');history.back();</script>"
		  Exit Sub
		End If
		Conn.Execute("update ks_guestreply set verific=2 where ID=" & KS.ChkClng(KS.S("replyID")))
		Response.Redirect request.servervariables("http_referer")
	 end sub
	 sub replyunlock()
		If check=false Then
		  Response.Write"<script>alert('�Բ�����û��������Ȩ��!');history.back();</script>"
		  Exit Sub
		End If
		Conn.Execute("update ks_guestreply set verific=1 where ID=" & KS.ChkClng(KS.S("replyID")))
		Response.Redirect request.servervariables("http_referer")
	 end sub
	 sub movetopic()
		If check=false Then
		  Response.Write"<script>alert('�Բ�����û��������Ȩ��!');history.back();</script>"
		  Exit Sub
		End If
		dim rs,oldboardid,replynum,boardid
		boardid=KS.ChkClng(KS.S("Boardid"))
		if boardid=0 then
		 KS.AlertHintScript "�����������!"
		end if
		set rs=server.createobject("adodb.recordset")
		rs.open "select top 1 * from ks_guestbook where id=" & KS.ChkClng(KS.S("ID")),conn,1,1
		if not rs.eof then
		 oldboardid=rs("boardid")
		 if oldboardid=boardid then
		  rs.close
		  set rs=nothing
		   Response.Redirect request.servervariables("http_referer")
		 end if
		 replynum=conn.execute("select count(id) from ks_guestreply where topicid=" & rs("id"))(0)
		 Conn.Execute("Update KS_GuestBoard set PostNum=PostNum-" & replynum &",TopicNum=TopicNum-1 where PostNum>" & replynum & " and id=" & oldboardid)
		 Conn.Execute("Update KS_GuestBoard set PostNum=PostNum+" & replynum &",TopicNum=TopicNum+1 where id=" & boardid)
		 Conn.Execute("update ks_guestbook set BoardID=" & Boardid & " where ID=" & rs("id"))
		 rs.close
		 set rs=nothing
		  KS.AlertHintscript "��ϲ�������ƶ��ɹ�!"
		end if
		rs.close
		set rs=nothing
		Response.Redirect request.servervariables("http_referer")
	 end sub
	
	 function check()
	 	Dim KSLoginCls
		Set KSLoginCls = New LoginCheckCls1
		If KSLoginCls.Check=true Then
		  check=true
		  Exit function
		else
			Dim KSUser:Set KSUser=New UserCls
			If Cbool(KSUser.UserLoginChecked)=false Then 
			  check=false
			  exit function
			else
			   check=KS.FoundInArr(master, KSUser.UserName, ",")
			End If
		end if
	 End function	
	 
	 '�����ȡ���,AdType�������  36 �Ҳ���,37 �ײ����
	 Function GetAdByRnd(ByVal AdType)
	      Dim AdStr:AdStr=KS.Setting(AdType)
	      If KS.IsNul(AdStr) Then Exit Function
		  Dim AdArr:AdArr=Split(AdStr,"@")
		  Dim RandNum,N: N=Ubound(AdArr)
          Randomize
          RandNum=Int(Rnd()*N)
          GetAdByRnd=AdArr(RandNum)
	End Function
		
					  
End Class
%>
