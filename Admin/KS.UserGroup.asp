<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Admin_UserGroup
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_UserGroup
        Private KS
		Private MaxPerPage
		Private RS,Sql
		Private ComeUrl
		Private ValidDays,tmpDays,BeginID,EndID,FoundErr,ErrMsg,PowerList
		Private iCount,Action,sPowerType,sDescript,sUserType,ValidType,ValidEmail

		Private Sub Class_Initialize()
		  MaxPerPage=20
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
       Sub Kesion()
	        Response.Write "<html>"
			Response.Write"<head>"
			Response.Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			Response.Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			Response.Write"<script src=""../KS_Inc/common.js""></script>"
			Response.Write"<script src=""../KS_Inc/jquery.js""></script>"
			Response.Write"</head>"
			Response.Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			Response.Write"	<ul id='mt'> "
			Response.Write "<div id='mtl'>�û����������</div><li><a href=""KS.UserGroup.asp"">������ҳ</a>&nbsp;|&nbsp;<a href=""#"" onclick=""AddGroup()"">�����û���</a>"
			Response.Write	" </ul>"
            If Not KS.ReturnPowerResult(0, "KMUA10004") Then
			  response.Write ("<script>$(parent.document).find('#BottomFrame')[0].src='javascript:history.back();';</script>")
			  Call KS.ReturnErr(1, "")
			End If

		Action=Trim(request("Action"))
			Select Case Action
			Case "Add", "Modify"
				call InfoPurview()
			Case "SaveAdd"
				call SaveAdd()
			Case "SaveModify"
				call SaveModify()
			Case "Del"
				call Del()
			Case else
				call main()
			End Select
			
			if FoundErr=True then
				KS.ShowError(ErrMsg)
			end if
			response.Write "<div style=""text-align:center;color:#003300"">-----------------------------------------------------------------------------------------------------------</div>"
			response.Write "<div style=""height:30px;text-align:center"">KeSion CMS V 6.5, Copyright (c) 2006-2009 <a href=""http://www.kesion.com/"" target=""_blank""><font color=#ff6600>KeSion.Com</font></a>. All Rights Reserved . </div>"
		End Sub
		
		sub main()
			Set rs=Server.CreateObject("Adodb.RecordSet")
			sql="select * from KS_UserGroup order by ID"
			OpenConn : rs.Open sql,Conn,1,1
		%>
        <script>
		 function AddGroup()
		 { 
		 location.href='KS.UserGroup.asp?Action=Add';
		$(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape("�û������ >> <font color=red>����û���</font>")+'&ButtonSymbol=Go';
		}
		function EditGroup(ID)
		{
		 location.href='?Action=Modify&ID='+ID;
		 $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr='+escape("�û������ >> <font color=red>�޸��û���</font>")+'&ButtonSymbol=GoSave';
		}
			
		</script>
		<table border="0" align="center" width="100%" cellpadding="0" cellspacing="0">
		  <tr align="center" class="sort">
			<td  width="45">ID��</td>
			<td width="168">�û�������</td>
			<td width="390">�û�����</td>
			<td width="80">�� ��</td>
			<td width="80">����ע��</td>
			<td width="120">��Ա����</td>
			<td  width="150"> �� ��</td>
		  </tr>
		  <%do while not rs.EOF
			%>
		  <tr height="40" align="center" class='list' onMouseOver="this.className='listmouseover'" onMouseOut="this.className='list'"> 
			<td class="splittd" width="45"><%=rs("ID")%></td>
			<td class="splittd"><%=rs("GroupName")%></td>
			<td class="splittd" width="390"><%=rs("Descript")%> </td>
			<td class="splittd" width="91"><%
			if rs("Type")<>0 then
				Response.Write "<font color=blue>�Զ���</font>"
			else
				Response.Write "<font color=#ff0033>ϵͳ</font>"
			end if
			%> </td>
			<td class="splittd" width="91"><%
			if rs("ShowOnReg")=1 then
				Response.Write "<font color=#ff0033>����ע��</font>"
			else
				Response.Write "<font color=green>������</font>"
			end if
			%> </td>
			<td class="splittd" width="120"><%=Conn.Execute("Select Count(UserID) From KS_User Where GroupID=" & RS("ID"))(0)%> λ</td>
			<td class="splittd" width="150"><%
			Response.Write "<a href='#' onclick=""EditGroup(" & RS("ID") & ")"">�޸�</a>&nbsp;&nbsp;"
			if rs("Type")<>0 then Response.Write "<a href='KS.UserGroup.asp?Action=Del&ID=" & rs("ID") & "' onClick=""return confirm('ȷ��Ҫɾ�����û�����');"">ɾ��</a>"
			%>
			<a href="KS.User.asp?UserSearch=10&GroupID=<%=RS("ID")%>">�г���Ա</a></td>
		  </tr>
		  <%
			rs.MoveNext
		loop
		  %>
		</table>  
		<%
			rs.Close:set rs=Nothing
		end sub
		
		sub InfoPurview()

		Dim frmAction,sSubmit,GroupSetting,GroupSetArr
		Dim sGroupName,sGroupImg,sFormID,sShowOnReg
		Dim sChargeType,sValidDays,sGroupPoint,sTemplateFile,SpaceSize
		%>
		<SCRIPT language=javascript>
		$(document).ready(function(){
		 setmail($("input[name=ValidType][checked=true]").val());
		});
		function setmail(n)
		 { 
		   if (n==1){
			  document.getElementById('mailarea').style.display='';
		   }else
			  document.getElementById('mailarea').style.display='none';
		}
		function CheckForm()
		{
		  if(document.myform.GroupName.value=="")
			{
			  alert("�û���������Ϊ�գ�");
			  document.myform.GroupName.focus();
			  return false;
			}
		 $("#myform").submit();
		}
		</script>
		  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1"  class="ctable" >
				<form method="post" id="myform" action="KS.UserGroup.asp" name="myform" onSubmit="return CheckForm();">
<%
		if Action="Modify" then
			dim GroupID
			GroupID=KS.ChkClng(Trim(Request("ID")))
			if GroupID=0 then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>��ָ��Ҫ�޸ĵ��û���ID</li>"
				Exit Sub
			end if
			Set rs=Conn.Execute("Select * from KS_UserGroup where ID=" & GroupID)
			if rs.Bof and rs.EOF then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>�����ڴ��û��飡</li>"
				Exit Sub
			end if
			sGroupName		= rs("GroupName")
			sDescript       = rs("Descript")
			sChargeType		= rs("ChargeType")
			sUserType       = rs("UserType")
			sValidDays		= rs("ValidDays")
			sGroupPoint		= rs("GroupPoint")
			sPowerType      = rs("PowerType")
			PowerList		= rs("PowerList")
			sShowOnReg      = rs("ShowOnReg")
			sTemplateFile   = rs("TemplateFile")
			sFormID         = rs("FormID")
			SpaceSize       = rs("SpaceSize")
			ValidType       = trim(rs("ValidType"))
			ValidEmail      = rs("ValidEmail")
			GroupSetting    = rs("GroupSetting")
			frmAction		= "Modify"
			sSubmit			= "�޸�"
			rs.close
		else
			sGroupName		= ""
			sChargeType		= 1
			sValidDays		= 0
			sGroupPoint		= 0
			sShowOnReg      = 1
			sDescript       = ""
			frmAction		= "Add"
			sSubmit			= "����"
			sUserType       = 0
			sTemplateFile   = KS.Setting(116)
			SpaceSize       =1024
			ValidType       =0
			ValidEmail      ="��ӭ��ע���Ϊ[" & KS.Setting(1) & "]��վ��Ա��" & chr(13) & " ������֤�룺{$CheckNum}" & chr(13) & "��������ĵ�ַ�������������֤������ʼ���֤����֤ͨ�������Ϳ�����ʽ��Ϊ���ǵĻ�Ա�������йط����ˣ�" & chr(13) & "<a href=""{$CheckUrl}"" target=""_blank"">{$CheckUrl}</a>"
		end if
		GroupSetting=GroupSetting & ",0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0"
		GroupSetArr =Split(GroupSetting,",")
		Dim CurrPath:CurrPath=KS.Setting(3)&KS.Setting(90)
		If Right(CurrPath,1)="/" Then CurrPath=Left(CurrPath,Len(CurrPath)-1)
		%>
			<tr class="sort"> 
			  <td height="25" colspan="2" align="center"><font size="2"><strong><%=sSubmit%>�û���</strong></font></td>
			</tr>
			<tr class="tdbg"> 
			  <td width="32%"  height="30" align="right" class="clefttitle"><div align="right"><strong>�û������ƣ�</strong></div></td>
			  <td height="30">			    <input name="GroupName" type="text" size=30 value="<%=sGroupName%>">		      </td>
			</tr>
			<tr class="tdbg">
			  <td height="30" align="right" class="clefttitle"><div align="right"><strong>�û����飺</strong></div></td>
			  <td height="30"><textarea name="Descript" cols="50" rows="5" id="Descript"><%=sDescript%></textarea></TD>
		    </tr>
			<tr class="tdbg">
			  <td  height="30" align="right" class="clefttitle"><div align="right"><strong>�û������</strong></div></td>
			  <td height="30"><input name="UserType" type="radio" value="0" <%if sUserType=0 then Response.Write " checked"%>>
			    ���˻�Ա 
		        <input name="UserType" type="radio" value="1" <%if sUserType=1 then Response.Write " checked"%>>		        ��ҵ��Ա</TD>
		    </tr>
			<tr class="tdbg"> 
			  <td  height="30" align="right" class="clefttitle"><div align="right"><strong>�û���Ʒѷ�ʽ��</strong></div></td>
			  <td height="30">
			    <input name="ChargeType" type="radio" value="1" <%if sChargeType=1 then Response.Write " checked"%> >
				�۵���<br>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Ĭ�ϵ�����
<input name="GroupPoint" type="text" id="GroupPoint" value="<%=sGroupPoint%>" size="6" maxlength="5"> 
��<br>
				<input type="radio" name="ChargeType" value="2" <%if sChargeType=2 then Response.Write " checked"%> >
				��Ч��<br>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Ĭ����Ч�ڣ�
<input name="ValidDays" type="text" id="ValidDays" value="<%=sValidDays%>" size="6" maxlength="5"> 
��<br />
<input type="radio" name="ChargeType" value="3" <%if sChargeType=3 then Response.Write " checked"%>> 
<font color="red">������(��������)</font></TD>
			</tr>
			<tr class="tdbg"> 
			  <td height="30" align="right" class="clefttitle"><div align="right"><strong>��Ա����ģ�壺<br>
		      </strong></div></td>
			  <td height="30">&nbsp;
			  <input type="text" name="TemplateFile" id="TemplateFile" size="30" value="<%=sTemplateFile%>">&nbsp;<input type='button' name="Submit" class="button" value="ѡ��ģ��..." onClick="OpenThenSetValue('KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle='+escape('ѡ��ģ��')+'&CurrPath=<%=CurrPath%>',450,350,window,$('#TemplateFile')[0]);">		  </td>
			</tr>
			
			<tr class="tdbg"> 
			  <td height="30" align="right" class="clefttitle"><div align="right"><strong>ѡ��¼�����<br>
		      </strong></div></td>
			  <td height="30">&nbsp;
			  <select name="formid">
			   <%
			    Set RS=Conn.Execute("select id,formname from ks_userform")
				do while not rs.eof
				 If sFormID=rs(0) Then
				 response.write "<option value='" & rs(0) & "' selected>" & rs(1) & "</option>"
				 Else
				 response.write "<option value='" & rs(0) & "'>" & rs(1) & "</option>"
				 End If
				 rs.movenext
				loop
				rs.close
			   %>
			  </select>			  </td>
			</tr>
			<tr class="tdbg">
			  <td height="30" align="right" class="clefttitle"><div align="right"><strong>�Ƿ�����ǰ̨ע��ѡ��</strong></div></td>
			  <td height="30">&nbsp;
			  <input type='radio' name='ShowOnReg' value='1'<%if sShowOnReg="1" Then Response.Write " Checked"%>> ���� <input type='radio' name='ShowOnReg' value='0'<%if sShowOnReg="0" Then Response.Write " Checked"%>>������			  </td>
		    </tr>
			<tr class="tdbg">
			      <td width="32%" height="21" class='CleftTitle' align="right"><div><strong>�»�Աע����֤��ʽ��</strong></div><font color=red>��ѡ���ʼ���֤�������Աע���ϵͳ�ᷢһ�������֤����ʼ����˻�Ա����Ա������ͨ���ʼ���֤�����������Ϊ��ʽע���Ա</font></td>
			      <td height="21">
				  <input id='a1' onClick="setmail(this.value)" name="ValidType" type="radio"  value="0"<%If ValidType="0" Then Response.Write " checked"%>><label for='a1'>������֤</label><br>
			     <input id='a2' onClick="setmail(this.value)" name="ValidType" type="radio" value="1"<%If ValidType="1" Then Response.Write" Checked"%>><label for='a2'>�ʼ���֤</label><br>
			     <input id='a3' onClick="setmail(this.value)" name="ValidType" type="radio" value="2"<%If ValidType="2" Then   Response.Write " Checked"%> /><label for='a3'>��̨�˹���֤</label>
			 </td>
			</tr>
			<tr valign="middle" id="mailarea"  class="tdbg">
			      <td width="32%" height="21" class='CleftTitle' align="right"><div><strong>�»�Աע��ʱ���͵���֤�ʼ����ݣ�</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </div><div align=center>��ǩ˵����<br>{$CheckNum}����֤��<br>{$CheckUrl}����Աע����֤��ַ<div></td>
			      <td height="21"><textarea name="ValidEmail" id="ValidEmail" cols="70" rows="5"><%=ValidEmail%></textarea>
			</td>
			</tr>
			
			<tr class="tdbg">
			  <td height="30" align="right" class="clefttitle"><div align="right"><strong>����ռ��С��</strong></div></td>
			  <td height="30">&nbsp;
<input type="text" name="SpaceSize" size="10" value="<%=SpaceSize%>" />KB <font color="#FF0000">��ʾ��1 KB = 1024 Byte��1 MB = 1024 KB</font> </td>
		    </tr>
			
			
			<tr class="tdbg">
			  <td height="30" align="right" class="clefttitle"><div align="right"><strong>���ܷ��䣺</strong></div></td>
			  <td height="30">
			    
				  <table border="0" width="100%">
				   <tr>
				    <td><input name="PowerList" type="checkbox" value="s01"<%if InStr(1, PowerList,"s01" ,1)<>0 then Response.Write( "checked") %>>����/��ҵ�ռ�
					</td>
				    <td><input name="PowerList" type="checkbox" value="s02"<%if InStr(1, PowerList,"s02" ,1)<>0 then Response.Write( "checked") %>>��־����
					</td>
				    <td><input name="PowerList" type="checkbox" value="s03"<%if InStr(1, PowerList,"s03" ,1)<>0 then Response.Write( "checked") %>>���ѹ���
					</td>
				    <td><input name="PowerList" type="checkbox" value="s04"<%if InStr(1, PowerList,"s04" ,1)<>0 then Response.Write( "checked") %>>���ֹ���
					</td>
				    <td><input name="PowerList" type="checkbox" value="s05"<%if InStr(1, PowerList,"s05" ,1)<>0 then Response.Write( "checked") %>>��Ṧ��
					</td>
				   </tr>
				   <tr>
				    <td><input name="PowerList" type="checkbox" value="s06"<%if InStr(1, PowerList,"s06" ,1)<>0 then Response.Write( "checked") %>>Ȧ�ӹ���
					</td>
				    <td><input name="PowerList" type="checkbox" value="s07"<%if InStr(1, PowerList,"s07" ,1)<>0 then Response.Write( "checked") %>>�Զ������
					</td>
				    <td><input name="PowerList" type="checkbox" value="s08"<%if InStr(1, PowerList,"s08" ,1)<>0 then Response.Write( "checked") %>>��Ƹ����
					</td>
				    <td><input name="PowerList" type="checkbox" value="s09"<%if InStr(1, PowerList,"s09" ,1)<>0 then Response.Write( "checked") %>>�ʴ�����
					</td>
				    <td>
					
					</td>
					</tr>
				   <tr>
				    <td><input name="PowerList" type="checkbox" value="s10"<%if InStr(1, PowerList,"s10" ,1)<>0 then Response.Write( "checked") %>>��ҵ��Ʒ(����)
					</td>
				    <td><input name="PowerList" type="checkbox" value="s11"<%if InStr(1, PowerList,"s11" ,1)<>0 then Response.Write( "checked") %>>��ҵ����
					</td>
				    <td><input name="PowerList" type="checkbox" value="s12"<%if InStr(1, PowerList,"s12" ,1)<>0 then Response.Write( "checked") %>>�ؼ��ʹ��
					</td>
				    <td><input name="PowerList" type="checkbox" value="s13"<%if InStr(1, PowerList,"s13" ,1)<>0 then Response.Write( "checked") %>>����֤��
					</td>
				    <td><input name="PowerList" type="checkbox" value="s14"<%if InStr(1, PowerList,"s14" ,1)<>0 then Response.Write( "checked") %>>��ְ��Ƹ
					</td>
				   </tr>
				  
				  <tr>
				    <td><input name="PowerList" type="checkbox" value="s15"<%if InStr(1, PowerList,"s15" ,1)<>0 then Response.Write( "checked") %>>���ֶһ�
					</td>
				    <td><input name="PowerList" type="checkbox" value="s16"<%if InStr(1, PowerList,"s16" ,1)<>0 then Response.Write( "checked") %>>�ղؼ�
					</td>
				    <td><input name="PowerList" type="checkbox" value="s17"<%if InStr(1, PowerList,"s17" ,1)<>0 then Response.Write( "checked") %>>Ͷ�߽���
					</td>
				    <td><input name="PowerList" type="checkbox" value="s18"<%if InStr(1, PowerList,"s18" ,1)<>0 then Response.Write( "checked") %>>���ݷ���(Ͷ��)
					</td>
				    <td>
					</td>
				   </tr>

				   
				   </table>
				   
			   </td>
		    </tr>
			<tr><td colspan=2><hr color="green" size="1"></td></tr>
			<tr class="tdbg">
			  <td height="30" align="right" class="clefttitle"><div align="right"><strong>���⹦��ѡ�</strong></div></td>
			  <td height="30">
			    <input type='checkbox' name='groupsetting0'<%if GroupSetArr(0)="1" then response.write " checked"%> value='1'>�ڷ�����Ϣ��Ҫ��˵�ģ�ͣ��˻�Ա�鷢����Ϣ����Ҫ���<br/>
			    <input type='checkbox' name='groupsetting1'<%if GroupSetArr(1)="1" then response.write " checked"%> value='1'>�����޸ĺ�ɾ������˵ģ��Լ��ģ��ĵ�<br/>
				һ�������ֻ����ͬһ��ģ�ͷ���<input type='text' name='GroupSetting2' value="<%=GroupSetArr(2)%>" style='text-align:center;width:30px' />ƪ�ĵ�  <font color='blue'>������������"-1"</font><br/>
				������Ϣʱ��ȡ����Ϊģ�����õ�<input type='text' name='GroupSetting3' value="<%=GroupSetArr(3)%>" style='text-align:center;width:30px' />�� <font color='blue'>����"0" ��ʾ�˻�Ա�鲻�÷�</font><br/>
				������Ϣʱ��ȡ��ȯΪģ�����õ�<input type='text' name='GroupSetting4' value="<%=GroupSetArr(4)%>" style='text-align:center;width:30px' />�� <font color='blue'>����"0" ��ʾ�˻�Ա�鲻�õ�ȯ</font><br/>
				������Ϣʱ��ȡ�ʽ�Ϊģ�����õ�<input type='text' name='GroupSetting5' value="<%=GroupSetArr(5)%>" style='text-align:center;width:30px' />�� <font color='blue'>����"0" ��ʾ�˻�Ա�鲻���ʽ�</font><br/>
				�˻�Ա�鷢�����ۿɵã�<input type="text" name="GroupSetting6" size="5" value="<%=GroupSetArr(6)%>" style="text-align:center"/>�ֻ�����Ϊ����
               1���������۱�ɾ�����۳�<input type="text" name="GroupSetting7" size="5" value="<%=GroupSetArr(7)%>" style="text-align:center"/>�ֻ���
			   <font color=blue>������"0"��ʾ������Ҳ�����ٻ���</font><br/>
			   
			   �˻�Ա�鷢����Ϣ����˺��Ƿ�վ�ڶ���Ϣ֪ͨ<input type="radio" name="GroupSetting10" value="1" <%if GroupSetArr(10)="1" then response.write " checked"%>>�� <input type="radio" name="GroupSetting10" value="0" <%if GroupSetArr(10)="0" then response.write " checked"%>>�� <br/>
			   
			   �˻�Ա���Աÿ��<input type="text" name="GroupSetting8" size="5" value="<%=GroupSetArr(8)%>" style="text-align:center"/>���Ӻ�,���µ�¼����<input type="text" name="GroupSetting9" size="5" value="<%=GroupSetArr(9)%>" style="text-align:center"/>�ֻ��� <font color=blue>���뽱��������"0"</font>
			   
              </td>
			</tr>			
			
				<input name="ID" type="hidden" value="<%=GroupID%>">
				<input name="Action" type="hidden" id="Action" value="Save<%=frmAction%>">
		  </table>
		</form>
		<% 
			Set rs=Nothing
		end sub
		
		sub SaveAdd()
			Dim GroupName,ChargeType,ValidDays,GroupPoint,PowerType,PowerList,Descript,FormID,ShowOnReg,UserType,TemplateFile,SpaceSize,GroupSetting,i
			
			GroupName		= Trim(Request("GroupName"))
			ChargeType		= KS.ChkClng(Request("ChargeType"))
			PowerType       = KS.ChkClng(Request("PowerType"))
			PowerList       = Request("PowerList")
			ValidDays		= KS.ChkClng(Request("ValidDays"))
			GroupPoint		= KS.ChkClng(Request("GroupPoint"))
			FormID          = KS.ChkClng(Request("FormID"))
			ShowOnReg       = KS.ChkClng(Request("ShowOnReg"))
			Descript        = KS.G("Descript")
			UserType        = KS.ChkClng(Request("UserType"))
			TemplateFile    = Request("TemplateFile")
			SpaceSize       = KS.ChkClng(Request("SpaceSize"))
			ValidType       = KS.ChkClng(Request("ValidType"))
			ValidEmail      = Request.Form("ValidEmail")
			GroupSetting=""
			For i=0 to 30
			   If GroupSetting="" Then
			     GroupSetting=KS.ChkClng(Request("GroupSetting"&i))
			   Else
			     GroupSetting=GroupSetting &"," & KS.ChkClng(Request("GroupSetting"&i))
			   End If
			Next


			if GroupName="" then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>�û������Ʋ���Ϊ�գ�</li>"
			end if
			if Not IsNumeric(Replace(Replace(PowerType,",","")," ","")) then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>�û�Ȩ�޴���</li>"
			end if
			if FoundErr=True then Exit Sub
			if ChargeType<>1 and ChargeType<>2 and ChargeType<>3 then ChargeType=1
			
			
			sql="Select * from KS_UserGroup where GroupName='"&GroupName&"'"
			Set rs=Server.CreateObject("Adodb.RecordSet")
			rs.Open sql,Conn,1,3
			if not (rs.bof and rs.EOF) then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>���ݿ����Ѿ����ڴ��û������ƣ�</li>"
				rs.close:set rs=Nothing
				exit sub
			end if
			rs.addnew
			rs("GroupName")		= GroupName
			rs("ChargeType")	= ChargeType
			rs("ValidDays")		= ValidDays
			rs("GroupPoint")	= GroupPoint
			rs("PowerList")		= PowerList
			rs("PowerType")     = PowerType
			rs("FormID")        = FormID
			rs("ShowOnReg")     = ShowOnReg
			rs("Descript")	    = Descript
			rs("UserType")      = UserType
			rs("TemplateFile")  = TemplateFile
			rs("SpaceSize")     = SpaceSize
			rs("ValidType")     = ValidType
			rs("ValidEmail")    = ValidEmail
			rs("GroupSetting")  = GroupSetting
			Rs("Type")          = 1
			rs.update
			rs.Close:set rs=Nothing
			Application(KS.SiteSN&"_UserGroup")=empty
			Call KS.Confirm("������û���ɹ�,���������?","KS.UserGroup.asp?Action=Add","KS.UserGroup.asp")
		end sub
		
		sub SaveModify()
			Dim GroupName,GroupID,GroupSetting,I
			Dim ChargeType,ValidDays,GroupPoint,PowerType,PowerList,Descript,FormID,ShowOnReg,UserType,TemplateFile,SpaceSize
			GroupID		= Trim(Request("ID"))
			GroupName		= Trim(Request("GroupName"))
			ChargeType		= KS.ChkClng(Request("ChargeType"))
			UserType        = KS.ChkClng(Request("UserType"))
			PowerType       = KS.ChkClng(Request("PowerType"))
			PowerList       = Request("PowerList")
			ValidDays		= KS.ChkClng(Request("ValidDays"))
			GroupPoint		= KS.ChkClng(Request("GroupPoint"))
			FormID          = KS.ChkClng(Request("FormID"))
			ShowOnReg       = KS.ChkClng(Request("ShowOnReg"))
			SpaceSize       = KS.ChkClng(Request("SpaceSize"))
			TemplateFile    = Request("TemplateFile")
			Descript        =KS.G("Descript")
			ValidType       = KS.ChkClng(Request("ValidType"))
			ValidEmail      = Request.Form("ValidEmail")
			
			GroupSetting=""
			For i=0 to 30
			   If GroupSetting="" Then
			     GroupSetting=KS.ChkClng(Request("GroupSetting"&i))
			   Else
			     GroupSetting=GroupSetting &"," & KS.ChkClng(Request("GroupSetting"&i))
			   End If
			Next
			
			if GroupName="" then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>�û������Ʋ���Ϊ�գ�</li>"
			end if
			if Not IsNumeric(Replace(Replace(PowerType,",","")," ","")) then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>�û�Ȩ�޴���</li>"
			end if
			if FoundErr=True then Exit Sub
			if ChargeType<>1 and ChargeType<>2 and ChargeType<>3 then ChargeType=1
			
			
			sql="Select * from KS_UserGroup where ID="&GroupID
			Set rs=Server.CreateObject("Adodb.RecordSet")
			rs.Open sql,Conn,1,3
			if not (rs.bof and rs.EOF) then
			rs("GroupName")		= GroupName
			rs("ChargeType")	= ChargeType
			rs("ValidDays")		= ValidDays
			rs("GroupPoint")	= GroupPoint
			rs("PowerList")		= PowerList
			rs("PowerType")     = PowerType
			rs("FormID")        = FormID
			rs("ShowOnReg")     = ShowOnReg
			rs("Descript")	    = Descript
			rs("UserType")      = UserType
			rs("TemplateFile")  = TemplateFile
			rs("SpaceSize")     = SpaceSize
			rs("ValidType")     = ValidType
			rs("ValidEmail")    = ValidEmail
			rs("GroupSetting")  = GroupSetting
			rs.update
		   else
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>�޴��û��飬�������ݳ���</li>"
		   end if
			rs.Close:set rs=Nothing
			conn.execute("update ks_user set usertype=" & UserType &" where groupid=" & groupid)
			IF FoundErr=true Then 
			 Exit Sub
			else
			Application(KS.SiteSN&"_UserGroup")=empty
			Response.Write ("<script>alert('�û���Ȩ���޸ĳɹ���');location.href='KS.UserGroup.asp';</script>")
			end if
		end sub
		sub Del()
		dim GroupID
		GroupID=Trim(Request("ID"))
		if GroupID="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>��ָ��Ҫɾ���Ĺ���ԱID</li>"
			exit sub
		end if
		GroupID=Clng(GroupID)
		'����ǰ̨�û���������Ȩ��
		if GroupID=0 then KS.ShowError("<br><li>������ɾ��ϵͳ�û��飡</li>")
		Conn.Execute("Update KS_User Set GroupID=2 where GroupID=" & GroupID)
		Conn.Execute("delete from KS_UserGroup where ID=" & GroupID)
		Application(KS.SiteSN&"_UserGroup")=empty
		Call KS.Alert("ɾ���û���ɹ���","KS.UserGroup.asp")
end sub
End Class
		%>
 
