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
Set KSCls = New MyComment
KSCls.Kesion()
Set KSCls = Nothing

Class MyComment
        Private KS,KSUser
		Private CurrentPage,totalPut
		Private RS,MaxPerPage
		Private TempStr,SqlStr,flag,Action
		Private Sub Class_Initialize()
			MaxPerPage =10
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		
		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "<script>top.location.href='Login';</script>"
		  Exit Sub
		End If
		Call KSUser.Head()
		flag=KS.ChkClng(KS.S("flag"))
		Action=KS.S("Action")
		
		Select Case Action
		  Case "Edit","Save"
		   Call CommentEdit()
		   Call KSUser.InnerLocation("�޸��ҷ��������")
		  Case "Cancel"
			 Conn.Execute("Delete From KS_Comment Where ID=" & KS.ChkClng(KS.S("ID")) & " And ChannelID=" & KS.ChkClng(KS.S("ChannelID")) & " And UserName='" & KSUser.UserName & "'")
			 Response.Redirect Request.ServerVariables("HTTP_REFERER")
		  Case Else
		   Call CommentList()
		End Select
		End Sub
		
		Sub CommentList()
		Dim Param:Param=" Where UserName='" & KSUser.UserName & "'"
		%>
		<TABLE cellSpacing=0 width=100% align=center border=0>
		<TR>
		<TD vAlign=top bgColor=#FFFFFF>
				<%
			   		       If KS.S("page") <> "" Then
						          CurrentPage = CInt(KS.S("page"))
							Else
								  CurrentPage = 1
							End If
							If Action="My" Then
							Call KSUser.InnerLocation("�û����ҵĸ������")
							Else
							Call KSUser.InnerLocation("�Ҳ��������")
							End If
				If Action="" Then %>
		<div class="tabs">						  
			<ul>
				<li<%If flag=0 then response.write " class='select'"%>><a href="User_MyComment.asp">�Ҳ������������(<span class="red"><%=Conn.Execute("Select count(id) from KS_Comment" & Param & "")(0)%></span>)</a></li>
				<li<%If flag=1 then response.write " class='select'"%>><a href="User_MyComment.asp?flag=1">����˵�����(<span class="red"><%=Conn.Execute("Select count(id) from KS_Comment" & Param & " and verific=1")(0)%></span>)</a></li>
				<li<%If flag=2 then response.write " class='select'"%>><a href="User_MyComment.asp?flag=2">δ��˵�����(<span class="red"><%=Conn.Execute("Select count(id) from KS_Comment" & Param & " and verific=0")(0)%></span>)</a></li>
			</ul>
          </div>
		    <%else%>
		<div class="tabs">						  
			<ul>
				<li<%If flag=0 then response.write " class='select'"%>><a href="User_MyComment.asp?action=My">�û����Ҹ��������(<span class="red"><%=Conn.Execute("Select count(id) from KS_Comment" & Param & "")(0)%></span>)</a></li>
			</ul>
          </div>
			<%end if%>
					 <table width="98%" align="center" border="0" cellspacing="1" cellpadding="1">
                              <%
								If flag=1 then Param=Param & " and c.verific=1"
								If flag=2 then Param=Param & " and c.verific=0" 
								If Action="My" Then 
							   	SqlStr="Select c.ID,c.Content,c.AddDate,c.Point,c.Verific,c.ChannelID,c.InfoID,c.replycontent From KS_Comment c inner join KS_ItemInfo I on c.infoid=i.infoid  Where i.inputer='" & KSUser.UserName & "' order by c.adddate desc"
								Else
							   	SqlStr="Select ID,Content,AddDate,Point,Verific,ChannelID,InfoID,replycontent From KS_Comment c" & Param & " order by adddate desc"
								End If

								Set RS=Server.CreateObject("AdodB.Recordset")
								 RS.open SqlStr,conn,1,1
								 If RS.EOF And RS.BOF Then
								  Response.Write "<tr class='tdbg'><td height=50 valign=top>û���κ�����!</td></tr>"
								 Else
									totalPut = RS.RecordCount
									If CurrentPage < 1 Then	CurrentPage = 1
			
								
			
								If CurrentPage>1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
								Else
										CurrentPage = 1
								End If
								Call showContent
				End If
     %>
                            </table></td>
                          </tr>
                        </table>
			      
		  </TD>
		  </TR>
</TABLE>
   
		  
		  <%
  End Sub

  Sub ShowContent()
    Dim I
   Do While Not RS.Eof
		%>
            <tr>
			  <td style="width:30px" class="splittd"><img src="images/comment.gif"></td>
              <td height="55" class="splittd"> 
			  <div class="ContentTitle"><a href="User_MyComment.asp?from=<%=Action%>&Action=Edit&ID=<%=RS(0)%>&Page=<%=CurrentPage%>">�������ݣ�<%=KS.GotTopic(RS(1),50)%></a>
			  <%
			  if rs("replycontent")<>"" and not isnull(rs("replycontent")) then
			   response.write "<font color=red>�ѻظ�</font>"
			  end if
			  %>
			  </div>
			  <div class="Contenttips">
			  <span>����ʱ�䣺<%=KS.GetTimeFormat(rs(2))%>
			  ״̬��
			  <%
			  if RS(4)=1 Then
				 Response.Write "����"
			  else
				 Response.Write "<font color=red>δ��</font>"
			 end if
			 
			SqlStr="Select ID,Title,Tid,Fname From " & KS.C_S(RS(5),2) & " Where ID=" & RS(6)
			 Dim RSI:Set RSI=Conn.Execute(SqlStr)
			 If NoT RSI.Eof Then
			  Response.Write "&nbsp;&nbsp;&nbsp;��Ϣ��<a href='" & KS.GetItemUrl(RS(5),RSI(2),RSI(0),RSI(3)) & "' target='_blank'>" & RSI(1) & "</a>"
			 End If
			 RSI.Close:Set RSI=Nothing
			 
			  Response.Write "</span><td align=center nowrap class=""splittd"">&nbsp;"
			  If Action="My" Then
				  Response.Write "<a class=""box"" href='User_MyComment.asp?from=" &Action & "&Action=Edit&ID=" & RS(0)& "&Page=" & CurrentPage & "'>�鿴/�ظ�</a> "
			  Else
				  if rs(4)<>1 Then
				  Response.Write "<a class=""box"" href='User_MyComment.asp?from=" &Action & "&Action=Edit&ID=" & RS(0)& "&Page=" & CurrentPage & "'>�޸�</a> "
				  else
				  Response.Write "<a class=""box"" href='#' disabled>�޸�</a> "
				  End If
				  Response.Write "<span ><a class=""box"" href='User_MyComment.asp?Action=Cancel&ChannelID=" & RS(5) &"&ID="& RS(0) &"&Page=" & CurrentPage & "' onclick=""return(confirm('ȷ��ɾ����������'))"">ɾ��</a></span>"
			 End If
			  Response.Write "</td>"
			  %>
			  </td>
            </tr>
                   <%
					RS.MoveNext
					I = I + 1
					If I >= MaxPerPage Then Exit Do
				    Loop
%>
          </table>
		  </td>
</tr>

	<% IF totalPut>MaxPerPage Then%>
                                <tr>
                                  <td  align="right">
									<%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>
							      </td>
                                </tr>
		<%End IF
  End Sub
  
  Sub CommentEdit()
  %>
  		<TABLE height="380" cellSpacing=0 width=100% align=center border=0>
		<TR>
		<TD vAlign=top bgColor=#FFFFFF>
		
		  <table width="100%" height="460" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td valign="top">
		<br>
        <br>
		<br><br>
		        <%Dim ID:ID=KS.ChkClng(KS.S("ID"))
				
				  IF ID="" Or Not IsNumeric(ID) Then  ID=0
				  Dim Page:Page=KS.S("Page")
				  Dim RSE
		         If KS.S("Action")="Save" Then
				   Dim AnounName:AnounName=KS.S("AnounName")
				   Dim Email:Email=KS.S("Email")
				   Dim Content:Content=KS.S("Content")
				   Dim Anonymous:Anonymous=KS.S("Anonymous")
				   Dim Point:point=KS.S("point")
				   Dim ChannelID:ChannelID=KS.ChkClng(KS.S("ChannelID"))
				   If ChannelID=0 Then ChannelID=1
				   IF point="" Or Not IsNumeric(point) Then point=0
				   Dim From :From=KS.S("From")
				   Set RSE=Server.CreateObject("Adodb.Recordset")
				   If From="My" Then
				     Content=KS.S("ReplyContent")
					 If Content="" Then 
					   Call KS.AlertHistory("��û������ظ�����!",-1)
					   Response.End
					 End If
					 RSE.Open "Select TOP 1 * From KS_Comment Where ID=" & ID,Conn,1,3
					IF RSE.EOF AND RSE.Bof Then
					 Response.Write "<script>alert('�������ݳ���!');history.back();</script>"
					 Response.End()
					Else
					  Dim OldContent:OldContent=RSE("ReplyContent")
					  RSE("ReplyContent")=Content
					  RSE("ReplyTime")=Now
					  RSE("ReplyUser")=KSUser.UserName
                      RSE.Update
					End if
					RSE.Close
					Set RSE=Nothing
					If OldContent<> Content Then
					 Call KSUser.AddLog(KSUser.UserName,"�ظ����û�����������!",100)
					End If
					 Response.Write "<script>alert('���ۻظ��ɹ�!');location.href='User_MyComment.asp?Action=My&ChannelID=" & ChannelID & "&Page=" & Page& "';</script>"
					 response.end
				   End If
					
					
					if Content="" Then 
					 Response.Write("<script>alert('����д��������!');history.back();</script>")
					 Response.End
					End if
					
					If Len(Content)>KS.C_S(ChannelID,14) and KS.C_S(ChannelID,14)<>0 Then
					 Response.Write("�������ݱ�����" &KS.C_S(ChannelID,14) & "���ַ�����!'")
					 Response.End
					End if

				    IF ID="" Then Response.Write "<script>alert('�������ݳ���!');history.back();</script>"
					
					RSE.Open "Select TOP 1 * From KS_Comment Where ID=" & ID,Conn,1,3
					IF RSE.EOF AND RSE.Bof Then
					 Response.Write "<script>alert('�������ݳ���!');history.back();</script>"
					 Response.End()
					Else
					 ' RSE("AnounName")=AnounName
					 ' RSE("Email")=Email
					  RSE("Content")=Content
					  'RSE("point")=point
                      RSE.Update
					End if
					RSE.Close
					Set RSE=Nothing
					Response.Write "<script>alert('��������޸ĳɹ�!');location.href='User_MyComment.asp?ChannelID=" & ChannelID & "&Page=" & Page& "';</script>"
				 Else
				  Call GetWriteComment(ID,Page)
				 End IF
				%>
			  </td>
           </tr>
     
          </table>
		  </TD>
		 </TR>
</TABLE> 

  <%
  End Sub
  
 Sub GetWriteComment(ID,Page)
         Dim RS,From
		 Set RS=Conn.Execute("Select * From KS_Comment Where ID=" &ID)
		 IF RS.EOF AND RS.BOF Then
		   Response.Write "<script>alert('�������ݳ���!');history.back();</script>"
		   exit Sub
		 End IF
		 From= KS.S("From")
		 %>
		 <script>
		 	function insertface(Val)
	        { 
		  if (Val!=''){
		   Val='[e'+Val+']';
		   document.getElementById('Content').focus();
		   var str = document.selection.createRange();
		   str.text = Val; 
		  }
          }
		 </script>

		 <%
		With KS
		 .echo "<table width=""90%"" border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"" class=""border"">"
		 .echo "<form name=""form1"" action=""?from=" & From & "&ChannelID=" & RS("ChannelID") &"&Action=Save&ID=" & ID & "&Page=" & Page & """ method=""post"">"
		 
		 If From="My" Then
			 .echo "<tr class=""title"">"
			 .echo "  <td height=""26"" colspan=""3"" align=""center"">�� �� �� �� �� ��</td></tr>"
			 .echo "<tr class=""tdbg"">"
			 .echo "  <td height=""30"" colspan=""3""> �� �� �ˣ�&nbsp;&nbsp;" & RS("AnounName")
			 .echo "   </td>"
			 .echo "  </tr>"
			 .echo "<tr class=""tdbg"">"
			 .echo "  <td height=""30"" colspan=""3""> �������ݣ�&nbsp;&nbsp;" & RS("content")
			 .echo "   </td>"
			 .echo "  </tr>"
			 .echo "<tr class=""tdbg"">"
			 .echo "  <td height=""30"" colspan=""3""> �ظ����ݣ�&nbsp;&nbsp;<textarea name='ReplyContent' rows='6' id='ReplyContent' style='width:80%'>" & RS("ReplyContent") & "</textarea>"
			 .echo "   </td>"
			 .echo "  </tr>"
			 
			 
		 Else
			 .echo "<tr class=""title"">"
			 .echo "  <td height=""26"" colspan=""3"" align=""center"">�� �� �� ��</td></tr>"
			 .echo "<tr class=""tdbg"">"
			 .echo "  <td height=""30"" colspan=""3""> �� �� �ˣ�"
			 .echo "   &nbsp;&nbsp;<input name=""AnounName"" type=""text"" readonly id=""AnounName"" value=""" & RS("AnounName") & """ style=""width:55%""/>"
			 .echo "   </td>"
			 .echo "  </tr>"
			 .echo "  <tr class=""tdbg"">"
			 .echo "    <td width=""80"">�������ݣ�</td>"
			 .echo "    <td height=""25""><textarea name=""Content"" rows=""6"" id=""Content"" style=""width:100%"">" & RS("Content") & "</textarea></td>"
			 .echo "    <td height=""25"" width=130 align=""center"">"
			
			 Dim k,str:str="����|Ʋ��|ɫ|����|����|����|����|����|˯|���|����|��ŭ|��Ƥ|����|΢Ц|�ѹ�|��|�ǵ�|ץ��|��|"
			 Dim strArr:strArr=Split(str,"|")
			 For K=0 to 19
				.echo "<img style=""cursor:pointer"" title=""" & strarr(k) & """ onclick=""insertface(" & k &")""  src=""" & KS.Setting(3) & "images/emot/" & K & ".gif"">&nbsp;"
			   If (K+1) mod 5=0 Then  .echo "<br />"
			 Next
		 End If
	 
		 .echo "</td>"

		 .echo "  </tr>"
		 
		 
		 .echo "  <tr class=""tdbg"">"
		 
		 If From="" Then
			If RS("Verific")<>1 Then
			 .echo "    <td height=""25"" colspan=""2"" align=""center""><input class=""button"" type=""submit"" name=""SubmitComment"" value=""�ύ�޸�""/>&nbsp;<input class=""button"" type=""button"" onclick=""javascript:history.back();"" name=""SubmitComment"" value=""ȡ������""/>"
			Else
			 .echo "    <td height=""25"" colspan=""2"" align=""center"">&nbsp;<input class=""button"" type=""button"" onclick=""javascript:history.back();"" name=""SubmitComment"" value=""ȡ������""/>"
			End If
		Else
		  .echo "<td colspan='2' align='center'><input class='button' type='submit' value='�ύ�ظ�'>"
		End If
		
		 .echo "   </td>"
		 .echo "  </tr>"
		 .echo "  </form>"
		 .echo "</table>"
		RS.Close
		Set RS=Nothing
	  End With
	End Sub

End Class
%> 
