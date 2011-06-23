<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.MemberCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Admin_Index
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Index
        Private KS,KSUser
		Private CurrentPage,totalPut
		Private RS,MaxPerPage
		Private ChannelID
		Private TempStr,SqlStr
		Private InfoIDArr,InfoID,DomainStr
		Private Sub Class_Initialize()
			MaxPerPage =20
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		  DomainStr=KS.GetDomain
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		
		Call KSUser.Head()
		Call KSUser.InnerLocation("我收藏的歌曲")
		ChannelID=KS.S("ChannelID")
		KSUser.CheckPowerAndDie("s13")
		
		%>
		
		  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td valign="top">
				<%
				Select Case KS.S("Action")
				 Case "Add"
				   Dim RSAdd
				   InfoID=KS.ChkClng(KS.S("InfoID"))
				   Set RSAdd=Server.CreateObject("Adodb.Recordset")
				   RSADD.Open "Select * From KS_Favorite Where ChannelID=6 And InfoID=" & InfoID & " And UserName='" & KSUser.UserName & "'",Conn,1,3
				   IF RSADD.Eof And RSADD.Bof Then
				      RSADD.AddNew
					    RSAdd(1)=KSUser.UserName
						RSAdd(2)=6
						RSAdd(3)=InfoID
						RSAdd(4)=Now
					  RSAdd.Update
				   End IF
				   RSADD.Close:SET RSADD=Nothing
				 Case "Cancel"
				  InfoID=KS.ChkClng(KS.S("InfoID"))
				  Conn.Execute("Delete From KS_Favorite Where InfoID=" & InfoID & " And ChannelID=6 And UserName='" & KSUser.UserName & "'")
				End Select
			   		       If KS.S("page") <> "" Then
						          CurrentPage = CInt(KS.S("page"))
							Else
								  CurrentPage = 1
							End If
				
								   %>
			<div class="tabs">						  
			<ul>
				<li><a href="User_Favorite.asp">我收藏的信息(<span class="red"><%=Conn.Execute("Select count(id) from KS_Favorite where username='" & KSUser.UserName & "' and channelid<>6")(0)%></span>)</a></li>
				<li class='select'><a href="User_MusicBox.asp">我收藏的音乐(<span class="red"><%=Conn.Execute("Select count(id) from KS_Favorite where username='" & KSUser.UserName & "' and channelid=6")(0)%></span>)</a></li>
			</ul>	
			</div>
				<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
				<script language="jscript.encode" src="<%=DomainStr%>Music/PlayList/js.js"></script>
				<FORM name="playform" onSubmit="javascript:return lbsong();"action="<%=DomainStr%>Music/PlayList/index.asp " target=KeSionmusiclisten>
                          <tr class="Title">
                            <td width="5%" height="22" align="center">选中</td>
                            <td width="31%" height="22" align="center">歌曲名称</td>
                            <td width="12%" height="22" align="center">歌手</td>
                            <td width="25%" height="22" align="center">专辑</td>
                            <td width="17%" height="22" align="center">管理操作</td>
                          </tr>
					<%
						 SqlStr="Select ID,MusicName,Singer,SpecialID From KS_MSSongList Where ID In(" & GetInfoIDArr(6) &")"
						 Set RS=Server.createobject("adodb.recordset")
						 RS.open SqlStr,conn,1,1

						 If RS.EOF And RS.BOF Then
								  Response.Write "<tr class='tdbg'><td align=center height=25 colspan=5 valign=top>没有收藏任何歌曲!</td></tr>"
								 Else
									totalPut = RS.RecordCount
						
											If CurrentPage < 1 Then
												CurrentPage = 1
											End If
			
								If (CurrentPage - 1) * MaxPerPage > totalPut Then
									If (totalPut Mod MaxPerPage) = 0 Then
										CurrentPage = totalPut \ MaxPerPage
									Else
										CurrentPage = totalPut \ MaxPerPage + 1
									End If
								End If
			
								If CurrentPage = 1 Then
									Call ShowContent
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
										Call ShowContent
									Else
										CurrentPage = 1
										Call ShowContent
									End If
								End If
				End If

						
						 %>
					
          </table>
		  </td>
		  </tr>
</table>
		  </TD>
		 </TR>
	</TABLE> 
	</div>
		  <%
  End Sub
    
  Sub ShowContent()
    Dim i
   Do While Not RS.Eof
		%>
                <tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
                   <td height="25" align="center" class="font6">
											<INPUT id="id" onClick="unselectall()" type="checkbox" value="<%=RS(0)%>"  name="id">									</td>
                                            <td height="22" align="left">
											<a href="javascript:dqsong('<%=RS(0)%>');">
											<%=RS("MusicName")%></a><img src="images/radio.gif">
											</td>
											<td width="12%" height="22" align="center" class="font6">
											<%
											    Response.Write rs(2)
											 %>
										    </td>
                                            <td width="15%" height="22" align="center" class="font6"><%=conn.execute("select name from KS_MSSpecial Where SpecialID=" & rs(3))(0)%></td>
                                            <td height="22" align="center">
											<a href="?Action=Cancel&Page=<%=CurrentPage%>&InfoID=<%=rs(0)%>" onclick = "return (confirm('确定取消该首歌曲的收藏吗?'))" class="link3">取消收藏</a>
											</td>
                                          </tr>

                                      <%
							RS.MoveNext
							I = I + 1
					  If I >= MaxPerPage Then Exit Do
				    Loop
%>
								<tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
								  <td height="30" valign=top colspan="5">
								&nbsp;<INPUT  class="Button" id="chkAll" onClick="CheckAll(this.form)"  type="button" value=" 全 选 "  name="chkAll">&nbsp;<INPUT  class="Button" id="fxAll" onClick="CheckOthers(this.form)"  type="button" value=" 反 选 "  name="fxAll"> <INPUT  class="Button" type=submit value=连续播放 name=submit1>
								  </td>
								  </FORM>
								</tr>
								<% IF totalPut>MaxPerPage Then%>
                                <tr>
                                  <td height="30" align='right' class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'" colspan=5>
								  <%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
										
                                       
							      </td>
                                </tr>
								<%End IF
  End Sub
  
  Function GetInfoIDArr(ChannelID)
     Dim RSObj,I
	 Set RSObj=Conn.Execute("Select InfoID From KS_Favorite Where UserName='" & KSUser.UserName & "' And ChannelID=" & ChannelID)
	 IF RSObj.Eof And RSObj.Bof Then
	  GetInfoIDArr="0"
	 Else
		 I=0
		 Do While Not RSObj.Eof
		   IF I=0 Then
			 GetInfoIDArr=RSObj(0) 
		   Else
			 GetInfoIDArr= GetInfoIDArr & "," & RSObj(0)
		   End IF
		   I=I+1 
		   RSObj.MoveNext
		 Loop
	End IF
	 RSObj.Close:Set RSObj=Nothing
  End Function
End Class
%> 
