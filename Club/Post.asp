<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New GuestPost
KSCls.Kesion()
Set KSCls = Nothing

Class GuestPost
        Private KS, KSR,KSUser,Templates,Node,BSetting
		Private GuestNum,GuestCheckTF,LoginTF
		Private Sub Class_Initialize()
		 If (Not Response.IsClientConnected)Then
			Response.Clear
			Response.End
		 End If
		  Set KS=New PublicCls
		  Set KSR = New Refresh
		  Set KSUser=New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSUser=Nothing
		 Set KSR=Nothing
		End Sub
		Sub EchoLn(str)
		 Templates=Templates & str & VBCrlf
		End Sub
%>
<!--#include file="../KS_Cls/Kesion.IfCls.asp"-->
<%
	Public Sub Kesion()
			If KS.Setting(56)="0" Then response.write "本站已关闭留言功能":response.end
			 GuestCheckTF=KS.Setting(52)
			 GuestNum=KS.Setting(54)
		    Dim FileContent,WriteForm
		          If KS.Setting(114)="" Then Response.Write "请先到""基本信息设置->模板绑定""进行模板绑定操作!":response.end
				   FileContent = KSR.LoadTemplate(KS.Setting(115))
				   If Trim(FileContent) = "" Then FileContent = "模板不存在!"
				   FCls.RefreshType = "guestwrite" '设置刷新类型，以便取得当前位置导航等
				   FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
				   LoginTF=KSUser.UserLoginChecked
				   WriteForm=LFCls.GetConfigFromXML("guestbook","/guestbook/template","post")
				   WriteForm=Replace(WriteForm,"{$CheckCode}",CheckCode)
				   WriteForm=Replace(WriteForm,"{$GuestNum}",GuestNum)
				   WriteForm=Replace(WriteForm,"{$CodeTF}",CodeTF)
				   WriteForm=Replace(WriteForm,"{$ImageList}",ImageList)
				   WriteForm=Replace(WriteForm,"{$EmotList}",EmotList)
				   WriteForm=Replace(WriteForm,"{$SelectBoard}",SelectBoard)
				   KS.LoadClubBoard
				  If KS.ChkClng(Request("bid"))<>0 Then
				   Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & Request("bid") &"]")                  
					  If Node Is Nothing Then KS.Die "非法参数!"
					  BSetting=Node.SelectSingleNode("@settings").text
				 End If
				   If KS.IsNul(BSetting) Then BSetting="1$$$0$0$0$0$0$0$0$0$$$$$$$$$$$$$$$$"
				   BSetting=Split(BSetting,"$")
				   If BSetting(2)<>"" And KS.FoundInArr(BSetting(2),KSUser.GroupID,",")=false Then
 				    FileContent=Replace(FileContent,"{$WriteGuestForm}",LFCls.GetConfigFromXML("GuestBook","/guestbook/template","error4"))
				   End If
				   
				   
				   If KS.Setting(57)="1" and LoginTF=false Then
					GCls.ComeUrl=GCls.GetUrl()
 				    FileContent=Replace(FileContent,"{$WriteGuestForm}",LFCls.GetConfigFromXML("GuestBook","/guestbook/template","error3"))
                   Else
				    If LoginTF=true Then
					 WriteForm=Replace(WriteForm,"{$UserName}",KSUser.UserName)
					 WriteForm=Replace(WriteForm,"{$User_Enabled}"," readonly ")
					 WriteForm=Replace(WriteForm,"{$UserEmain}",KSUser.Email)
					 WriteForm=Replace(WriteForm,"{$UserHomePage}",KSUser.HomePage)
					 WriteForm=Replace(WriteForm,"{$UserQQ}",KSUser.QQ)
					Else
					 WriteForm=Replace(WriteForm,"{$UserName}","")
					 WriteForm=Replace(WriteForm,"{$User_Enabled}","")
					 WriteForm=Replace(WriteForm,"{$UserEmain}","")
					 WriteForm=Replace(WriteForm,"{$UserHomePage}","http://")
					 WriteForm=Replace(WriteForm,"{$UserQQ}","")
					End If
 				    FileContent=Replace(FileContent,"{$WriteGuestForm}",WriteForm)
				   End If
				   FileContent=KSR.KSLabelReplaceAll(FileContent)
				   KS.Echo RexHtml_IF(FileContent)
		End Sub
		
		Function  CheckCode()
		 IF KS.Setting(53)=1 Then
  	      CheckCode="if (myform.Code.value==''){" & vbcrlf
	      CheckCode=CheckCode & "alert('请输入附加码，留言要有点耐心哦！！');" & vbcrlf
	      CheckCode=CheckCode & "myform.Code.focus();" & vbcrlf
  	      CheckCode=CheckCode & "return false;" & vbcrlf
	      CheckCode=CheckCode &  "}" & vbcrlf
	    End IF
	   End Function
					  
	   Function CodeTF()
	     if KS.Setting(53)=0 then CodeTF=" style='display:none'"
	   End Function				  
	   
	   Function ImageList()
	           dim i
			   for i=1 to 56 
			   ImageList=ImageList & "<option value=" & i & ">" & i & ".gif</option>"
			   next
	   End Function
	   
	   Function EmotList()
	        Dim I
			For I=1 To 30
			 EmotList=EmotList &  "<input type=""radio"" name=""txthead"" value=""" & I & """"
			  IF I=1 Then EmotList=EmotList &  " Checked"
				EmotList=EmotList &  " ><img src=""../Images/Face1/Face" & I & ".gif"" border=""0"">"
			  IF I Mod 15=0 Then EmotList=EmotList &  "<BR>"
			Next
	   End Function
	   
	   Function SelectBoard()
	      dim str,sql,i,node,xml,ors,n
		  If KS.Setting(59)="1" Then Exit Function
		  templates=""
		  If KS.ChkClng(request("bid"))<>0 Then
		   SelectBoard="<input type='hidden' value='" & request("bid") & "' name='boardid'>"
		   Exit Function
		  End If
		  
		  echoln "<script type=""text/javascript"">"
		  echoln " var subcity = new Array();"
          KS.LoadClubBoard()
		  if isobject(Application(KS.SiteSN&"_ClubBoard")) then
		  	  Set Xml=Application(KS.SiteSN&"_ClubBoard")
			n=0
			for each node in xml.documentelement.selectnodes("row[@parentid!=0]")
			 echoln "subcity[" & n  &"] = new Array('" & Node.selectsinglenode("@parentid").text & "','" & node.selectsinglenode("@boardname").text &"','" & Node.SelectSingleNode("@id").text & "');"
			 n=n+1
			next
		  end if
		   xml=empty
		  set node=nothing	
		  

		  echoln "function changecity(selectValue)"
		  echoln "{"
		  echoln "document.getElementById('boardid').length = 0;"
		  echoln "document.getElementById('boardid').options[0] = new Option('请选择子版面','');"
		  echoln "for (i=0; i<subcity.length; i++){"
		  echoln "if (subcity[i][0] == selectValue)"
		  echoln "{document.getElementById('boardid').options[document.getElementById('boardid').length] = new Option(subcity[i][1], subcity[i][2]);}"
		  echoln "}"
		  echoln "}"
		  echoln "</script>"
		  

		 dim rs:set rs=conn.execute("select id,boardname from ks_guestboard where parentid=0 order by orderid")
		 if not rs.eof then set xml=KS.RsToXml(rs,"row",""):rs.close:set rs=nothing
         If isobject(xml) then
		  echoln "<tr>"
		  echoln "<td height='25' align='right'><b>选择版面 ：</b></td>"
          echoln  "<td><select onChange='changecity(this.value)' name='pboardid' id='pboardid'>"
		  for each node in xml.documentelement.selectnodes("row")

		    If trim(KS.S("Bid"))=trim(node.selectsinglenode("@id").text) Then
		     echoln "<option value='" & node.selectsinglenode("@id").text &"' selected>" & node.selectsinglenode("@boardname").text & "</option>"
			Else
		     echoln "<option value='" & node.selectsinglenode("@id").text &"'>" & node.selectsinglenode("@boardname").text & "</option>"
			End If
		  next
          echoln " </select><select name='boardid' id='boardid'><option value=''>-选择子版面-</option></select></td>"
          echoln "</tr>"
		  echoln "<script>changecity(jQuery('#pboardid>option:selected').val());</script>"

		end if
		selectboard=templates
	   End Function
End Class
%>
