<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../Plus/md5.asp"-->
<!--#include file="../../KS_Cls/Kesion.MemberCls.asp"-->
<!--#include file="../../API/cls_api.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Admin_GetPass
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_GetPass
        Private KS,KSUser
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		Call KSUser.Head()
		Call KSUser.InnerLocation("�һ�����")
        
		  Dim Step:Step=KS.S("Step")
		  IF Step="" Then Step=1
		  IF Step=2 Then
		     Dim RS
			 Dim UserName:UserName=KS.R(KS.S("UserName"))
			 If UserName = "" Then
				 Response.Write("<script>alert('�������û���!');history.back();</script>")
				 Response.End
              End IF
			 
			 
             Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select Question From KS_User Where UserName='" & UserName & "'",Conn,1,1
			  IF RS.Eof And RS.Bof Then
			  	 Response.Write("<script>alert('�Բ���,��������û��������ڣ�');history.back();</script>")
				 Response.End
			  Else
		     %>
			 	<script language = "JavaScript">
				function CheckForm()
				{
				if (document.myform.Answer.value=="")
				  {
					alert("����������𰸣�");
					document.myform.Answer.focus();
					return false;
				  }
				if (document.myform.Verifycode.value=="")
				  {
					alert("��������֤�룡");
					document.myform.Verifycode.focus();
					return false;
				  }
	              return true;
				  }
				  function getCode(){
				   $("#showVerify").html('<img style="cursor:pointer" src="<%=KS.GetDomain%>plus/verifycode.asp?n=<%=Timer%>" onClick="this.src=\'<%=KS.GetDomain%>plus/verifycode.asp?n=\'+ Math.random();" align="absmiddle">')
				  }
				</script>
                  <br>
					  <table width="96%"  border="0" align="center" cellpadding="0" cellspacing="1" class="border">
					 	<form name="myform" method="post" action="?Step=3" onSubmit="return CheckForm();">
                        <input type="hidden" value="<%=UserName%>" name="UserName">
                        <tr class="Title">
                            <td height="24" colspan=2 align="center">ȡ������ڶ��� �ش��������� </td>
                        </tr>
                            <tr class="tdbg">
                              <td width="40%" height="30" align="right"> �������⣺</td>
                              <td width="60%"><%=RS(0)%></td>
                            </tr>
                            <tr class="tdbg">
                              <td width="40%" height="30" align="right"> ���Ĵ𰸣�</td>
                              <td width="60%"><input name="Answer" type="text" id="Answer" size="20" /></td>
                            </tr>
                            <tr class="tdbg">
                              <td width="40%" height="30" align="right"> ��֤�룺</td>
                              <td width="60%"><input name="Verifycode" onfocus="getCode()" type="text" id="Verifycode" size="6" />
							  <span id="showVerify"></span>
							  </td>
                            </tr>
                            <tr class="tdbg">
                              <td colspan=2 height="42" align="center"><input class="Button" name="Submit2" type="submit" value="��һ��" />
                                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
                            </tr>
							</form>
                        </table>
						
                    
		  <%   End IF
		  ElseIF Step=3 Then

             Dim Verifycode:Verifycode=	KS.R(KS.S("Verifycode"))
			 UserName=KS.R(KS.S("UserName"))

			 Dim Answer:Answer=KS.S("Answer")
			IF Trim(Verifycode)<>Trim(Session("Verifycode")) then 
		   	 Response.Write("<script>alert('��֤���������������룡');history.back();</script>")
		     Response.End
			End IF
			 If Trim(Answer)="" Then 
			   Response.Write("<script>alert('�������벻��Ϊ�գ�');history.back();</script>")
			   Response.End
			 End If

			
			Dim RSC:Set RSC=Conn.Execute("Select Answer From KS_User Where UserName='" & UserName & "' and Answer='" & Answer & "'")
			IF RSC.EOF AND RSC.Bof Then
			  	 Response.Write("<script>alert('�Բ���,������Ĵ𰸲���ȷ��');history.back();</script>")
				 Response.End
			Else
			 %>
			 
			 <script language = "JavaScript">
				function CheckForm()
				{
				if (document.myform.PassWord.value=="")
				  {
					alert("�����������룡");
					document.myform.PassWord.focus();
					return false;
				  }
				if (document.myform.RePassWord.value=="")
				  {
					alert("������ȷ�����룡");
					document.myform.RePassWord.focus();
					return false;
				  }
				if (document.myform.PassWord.value!=document.myform.RePassWord.value)
				  {
					alert("������������벻һ�£�");
					document.myform.PassWord.focus();
					return false;
				  }
	              return true;
				  }
				</script>
				<br>
                       <table width="96%"  border="0" align="center" cellpadding="0" cellspacing="1" class="border">
							<tr class="Title">
									<td height="24" align="center" colspan="2">ȡ����������� ���������� </td>
							</tr>
					 <form name="myform" method="post" action="?Step=4" onSubmit="return CheckForm();">
					    <input type="hidden" value="<%=Answer%>" name="answer">
                                <tr class="tdbg">
                                  <td width="40%" height="30" align="right"> �û�����</td>
                                  <td width="60%"><input type="text" readonly value="<%=UserName%>" name="UserName"></td>
                                </tr>
                                <tr class="tdbg">
                                  <td width="40%" height="30" align="right"> �����룺</td>
                                  <td width="60%"><input name="PassWord" type="password" id="PassWord" size="20" /></td>
                                </tr>
                                <tr class="tdbg">
                                  <td width="40%" height="30" align="right"> ȷ�����룺</td>
                                  <td width="60%"><input name="RePassWord" type="password" id="RePassWord" size="20" /></td>
                                </tr>
                                <tr class="tdbg">
                                  <td height="42" align="center" colspan=2><input  class="Button" name="Submit22" type="submit" value=" �� �� " />
                                    &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
                                </tr>
                              </tbody>
                            </table>
						  </form>
					
			<% End IF
		  ElseIF Step=4 Then
		     If Trim(Request.ServerVariables("HTTP_REFERER"))="" Then
			Response.Write "<script>alert('�Ƿ�������');history.back();</script>"
			Response.end
			 End If
			 if instr(lcase(Request.ServerVariables("HTTP_REFERER")),"step=3")<=0 then
				Response.Write "<script>alert('�Ƿ�����1��');history.back();</script>"
				Response.end
			 end if 

		    UserName=KS.DelSql(Replace(Replace(Request.Form("UserName"), "'", ""), """", ""))
			Answer=KS.S("answer")
		  	 Dim PassWord:PassWord=KS.DelSql(Replace(Replace(Request.Form("PassWord"), "'", ""), """", ""))
			 Dim RePassWord:RePassWord=KS.DelSql(Replace(Replace(Request.Form("RePassWord"), "'", ""), """", ""))
			 If UserName="" Then
				 Response.Write("<script>alert('�����Ƿ�!');history.back();</script>")
				 Response.End
			 End If
			 If PassWord = "" Then
				 Response.Write("<script>alert('�������¼����!');history.back();</script>")
				 Response.End
			 ElseIF RePassWord="" Then
				 Response.Write("<script>alert('������ȷ������');history.back();</script>")
				 Response.End
			 ElseIF PassWord<>RePassWord Then
				 Response.Write("<script>alert('������������벻һ��');history.back();</script>")
				 Response.End
			 End If

             Set RS=Server.CreateObject("Adodb.RecordSet")
			  RS.Open "Select PassWord From KS_User Where UserName='" & UserName & "' and answer='" & answer &"'",Conn,1,3
			  If Not RS.Eof Then
			     RS(0)=MD5(PassWord,16)
				 RS.Update
			  Else
				 Response.Write("<script>alert('����ʧ��!');history.back();</script>")
				 Response.End
			  End If
			 RS.Close:Set RS=Nothing
			 '-----------------------------------------------------------------
				'ϵͳ����
				'-----------------------------------------------------------------
				Dim API_KS,API_SaveCookie,SysKey
				If API_Enable Then
					Set API_KS = New API_Conformity
					API_KS.NodeValue "action","update",0,False
					API_KS.NodeValue "username",KS.S("UserName"),1,False
					Md5OLD = 1
					SysKey = Md5(API_KS.XmlNode("username") & API_ConformKey,16)
					Md5OLD = 0
					API_KS.NodeValue "syskey",SysKey,0,False
					API_KS.NodeValue "password",PassWord,1,False
					API_KS.SendHttpData
					If API_KS.Status = "1" Then
						Response.Write "<script>alert('" &  API_KS.Message  & "');</script>"
						Exit Sub
					End If
					Set API_KS = Nothing
				End If
				'-----------------------------------------------------------------
			 
		  %>
		  <br>
                  <table width="96%"  border="0" align="center" cellpadding="0" cellspacing="1" class="border">
                          <tr class="Title">
                              <td height="25" valign="bottom" align="center">ȡ������ɹ�</td>
                          </tr>
                           <tr class="tdbg">
                                  <td height="50" align="center">��ϲ��,����ȡ�سɹ�!������������:<font color=red><%=PassWord%></font>,�����������¼��</td>
                                </tr>
                            </table>
                       
		  <%
           Else
		   %>
		   <script language = "JavaScript">
				function CheckForm()
				{
				if (document.myform.UserName.value=="")
				  {
					alert("�������û�����");
					document.myform.UserName.focus();
					return false;
				  }
	              return true;
				  }
				</script>

			 <form name="myform" method="post" action="?Step=2" onSubmit="return CheckForm();">
                 <table width="96%"  border="0" align="center" cellpadding="0" cellspacing="1" class="border">
					  <tr class="Title">
							<td height="24" align="center" colspan="2">ȡ�������һ�� �����û��� </td>
					  </tr>
						  <TR class="tdbg">
							<TD width="40%" height=25 align="right"> �����û�����</TD>
							<TD width="60%"><input name="UserName" type="text" id="UserName" size="20"></TD>
						  </TR>
						  <TR class="tdbg">
							<TD  colspan="2" height=42 align="center"> 
							<input  class="Button" name="Submit" type="submit" value="��һ��">
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </TD>
						  </TR>
						</TBODY>
					  </TABLE>
				</form>
		  	 <%End IF%> 			  

		  <%
  End Sub
End Class
%> 
