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
Set KSCls = New Admin_MyShop
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_MyShop
        Private KS,KSUser,ChannelID
		Private CurrentPage,totalPut,Status,ProducerName
		Private RS,MaxPerPage,ComeUrl,SelButton,Price_Original,Price,Price_Market,Price_Member,Point,Discount
		Private ClassID,Title,KeyWords,ProModel,ProSpecificat,ProductType,Unit,TotalNum,AlarmNum,TrademarkName,Content,Verific,PhotoUrl,RSObj,I,UserDefineFieldArr,UserDefineFieldValueStr,UserClassID,ShowONSpace
		Private CurrentOpStr,Action,ID,ErrMsg,Hits,BigPhoto,BigClassID,SmallClassID,flag
		Private Sub Class_Initialize()
			MaxPerPage =12
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()
		 IF KS.S("ComeUrl")="" Then
     		ComeUrl=Request.ServerVariables("HTTP_REFERER")
		 Else
     		ComeUrl=KS.S("ComeUrl")
		 End If

		IF Cbool(KSUser.UserLoginChecked)=false Then
		  Response.Write "<script>top.location.href='Login';</script>"
		  Exit Sub
		End If
		ChannelID=KS.ChkClng(KS.S("ChannelID"))
		If ChannelID=0 Then ChannelID=5
		If KS.C_S(ChannelID,6)<>5 Then Response.End()
		if conn.execute("select usertf from ks_channel where channelid=" & channelid)(0)=0 then
		  Response.Write "<script>alert('��Ƶ���ر�Ͷ��!');window.close();</script>"
		  Exit Sub
		end if
		'��������ͼ����
		Session("ThumbnailsConfig")=KS.C_S(ChannelID,46)
		
		Call KSUser.Head()
		%>
		<div class="tabs">	
			<ul>
				<li<%If KS.S("Status")="" then response.write " class='select'"%>><a href="User_MyShop.asp?ChannelID=<%=ChannelID%>">�ҷ�����<%=KS.C_S(ChannelID,3)%>(<span class="red"><%=Conn.Execute("Select count(id) from " & KS.C_S(ChannelID,2) &" where Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="1" then response.write " class='select'"%>><a href="User_MyShop.asp?ChannelID=<%=ChannelID%>&Status=1">�����(<span class="red"><%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=1 and Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="0" then response.write " class='select'"%>><a href="User_MyShop.asp?ChannelID=<%=ChannelID%>&Status=0">�����(<span class="red"><%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=0 and Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="2" then response.write " class='select'"%>><a href="User_MyShop.asp?ChannelID=<%=ChannelID%>&Status=2">�� ��(<span class="red"><%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=2 and Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
				<li<%If KS.S("Status")="3" then response.write " class='select'"%>><a href="User_MyShop.asp?ChannelID=<%=ChannelID%>&Status=3">���˸�(<span class="red"><%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=3 and Inputer='"& KSUser.UserName &"'")(0)%></span>)</a></li>
			</ul>
		  </div>
		<%
		Action=KS.S("Action")
		Select Case Action
		 Case "Del"
		  Call KSUser.DelItemInfo(ChannelID,ComeUrl)
		 Case "Add","Edit"
		  Call ShopAdd
		 Case "AddSave","EditSave"
          Call ShopSave()
		 Case Else
		  Call ShopList
		 End Select
       End Sub
	   Sub ShopList
			  %>
			 <SCRIPT language=javascript src="../KS_Inc/showtitle.js"></script>
			  
			   <%
			   		       If KS.S("page") <> "" Then
						          CurrentPage = KS.ChkClng(KS.S("page"))
							Else
								  CurrentPage = 1
							End If
                                    
									Dim Param:Param=" Where Inputer='"& KSUser.UserName &"'"
									Verific=KS.S("status")
									If Verific="" or not isnumeric(Verific) Then Verific=4
                                    IF Verific<>4 Then 
									   Param= Param & " and Verific=" & Verific
									End If
									IF KS.S("Flag")<>"" Then
									  IF KS.S("Flag")=0 Then Param=Param & " And Title like '%" & KS.S("KeyWord") & "%'"
									  IF KS.S("Flag")=1 Then Param=Param & " And KeyWords like '%" & KS.S("KeyWord") & "%'"
									End if
									If KS.S("ClassID")<>"" And KS.S("ClassID")<>"0" Then Param=Param & " And TID='" & KS.S("ClassID") & "'"
									Dim Sql:sql = "select a.*,foldername from KS_Product a inner join ks_class b on a.tid=b.id "& Param &" order by AddDate DESC"

								  Select Case Verific
								   Case 0 
								    Call KSUser.InnerLocation("����"& KS.C_S(ChannelID,3) & "�б�")
								   Case 1
								    Call KSUser.InnerLocation("����"& KS.C_S(ChannelID,3) & "�б�")
								   Case 2
								   Call KSUser.InnerLocation("�ݸ�"& KS.C_S(ChannelID,3) & "�б�")
								   Case 3
								   Call KSUser.InnerLocation("�˸�"& KS.C_S(ChannelID,3) & "�б�")
                                   Case Else
								    Call KSUser.InnerLocation("����"& KS.C_S(ChannelID,3) & "�б�")
								   End Select
			   %>
			    <div style="padding-left:20px;"><img src="images/ico1.gif" align="absmiddle"><a href="user_myshop.asp?ChannelID=<%=ChannelID%>&Action=Add"><span style="font-size:14px;color:#ff3300">����<%=KS.C_S(ChannelID,3)%></span></a></div>

				<table width="98%" border="0" align="center" cellpadding="1" cellspacing="1" class="border">
                    <tr class="Title">
                          <td width="6%" height="22" align="center">ѡ��</td>
                          <td width="34%" height="22" align="center"><%=KS.C_S(ChannelID,3)%>����</td>
						  <td width="10%" height="22" align="center"><%=KS.C_S(ChannelID,3)%>¼��</td>
                          <td width="18%" height="22" align="center">���ʱ��</td>
                          <td width="10%" height="22" align="center">״̬</td>
                          <td width="22%" height="22" align="center">�������</td>
                   </tr>
                                      <%
									Set RS=Server.CreateObject("AdodB.Recordset")
									RS.open sql,conn,1,1
								 If RS.EOF And RS.BOF Then
								  
								  Response.Write "<tr><td class='tdbg' colspan='6' height=30 valign=top>û����Ҫ��"& KS.C_S(ChannelID,3) & "!</td></tr>"
								 
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
									Call showContent
								Else
									If (CurrentPage - 1) * MaxPerPage < totalPut Then
										RS.Move (CurrentPage - 1) * MaxPerPage
										Call showContent
									Else
										CurrentPage = 1
										Call showContent
									End If
								End If
				End If
     %>                      <tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
                                     <form action="User_MyShop.asp" method="post" name="searchform">
								  <td colspan="6" align="center">
										<%=KS.C_S(ChannelID,3)%>������
										  <select name="Flag">
										   <option value="0">����</option>
										   <option value="1">�ؼ���</option>
									      </select>
										  
										  �ؼ���
										  <input type="text" name="KeyWord" class="textbox" value="�ؼ���" size=20>&nbsp;<input class="button" type="submit" name="submit1" value=" �� �� ">
							      </td>
								    </form>
                                </tr>
							<tr>
							 <td colspan=6>
							  <strong><%=KS.C_S(ChannelID,3)%>����˵����</strong><br>
							  1���û��ڱ�վ������Ʒ���ۣ����﷽����������֧��������վ��<br/>
							  2�����﷽�ڱ�վ֧���ɹ��󣬱�վ������Ի����������Ч�Խ�����˼�֪ͨ���۷������ȣ�<br>
							  3���ٳɽ��׺�
							  ����վ����ȡ�����ܼ۵� <font color=red><%=KS.Setting(79)%>% </font>��Ϊ���׹����,��������֧�������۷���<br>
							  3����ȷ����������Ʒ��ʵ�ԣ�һ���������ڱ�վ��������Ϣ������٣���ƭ��Ϊ,���ǽ������������ڱ�վ�Ľ����˻���
							 </td>
							</tr>
</table>
		  <%
  End Sub
  
  Sub ShowContent()
     Dim I
    Response.Write "<FORM Action=""User_MyShop.asp?Action=Del"" name=""myform"" method=""post"">"
   Do While Not RS.Eof
         %>
                                          <tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
                                            <td class="splittd" height="22" align="center">
											<INPUT id="ID" type="checkbox" value="<%=RS("ID")%>"  name="ID">
											</td>
                                            <td class="splittd" align="left">[<%=RS("FolderName")%>]
											<%Dim PhotoStr:PhotoStr=RS("PhotoUrl")
											 if PhotoStr="" Or IsNull(PhotoStr) Then PhotoStr=KS.GetDomain & "images/Nopic.gif"%>
											 <a title="<table width=80 border=0 align=center><tr><td><img src='<%=PhotoStr%>' border=0 width='130' height='80'></td></tr></table>"  href="../item/show.asp?m=<%=channelid%>&d=<%=rs("id")%>" target="_blank" class="link3"><%=KS.GotTopic(trim(RS("title")),32)%></a></td>
											<td class="splittd" align="center"><%=rs("Inputer")%></td>
                                            <td class="splittd" align="center"><%=formatdatetime(rs("AddDate"),2)%></td>
                                            <td class="splittd" align="center">
											  <%Select Case rs("Verific")
											   Case 0
											     Response.Write "<span class=""font10"">����</span>"
											   Case 1
											     Response.Write "<span class=""font11"">����</span>"
                                               Case 2
											     Response.Write "<span class=""font13"">�ݸ�</span>"
											   Case 3
											     Response.Write "<span class=""font14"">�˸�</span>"
                                              end select
											  %></td>
                                            <td class="splittd" height="22" align="center">
											<%if rs("Verific")<>1 or KS.ChkClng(KS.U_S(KSUser.GroupID,1))=1 then%>
											<a class='box' href="User_MyShop.asp?channelid=<%=channelid%>&id=<%=rs("id")%>&Action=Edit&&page=<%=CurrentPage%>">�޸�</a> <a class='box' href="User_MyShop.asp?channelid=<%=channelid%>&action=Del&ID=<%=rs("id")%>" onclick = "return (confirm('ȷ��ɾ��<%=KS.C_S(ChannelID,3)%>��?'))">ɾ��</a>
											<%else
											 If KS.C_S(ChannelID,42)=0 Then
											  Response.write "---"
											 Else
											  Response.Write "<a  class='box' href='?channelid=" & channelid & "&id=" & rs("id") &"&Action=Edit&&page=" & CurrentPage &"'>�޸�</a> <a class='box' href='#' disabled>ɾ��</a>"
											 End If
											end if%>
											</td>
			</tr>
					   <tr><td colspan=6 background='images/line.gif'></td></tr>
                                      <%
							RS.MoveNext
							I = I + 1
					  If I >= MaxPerPage Then Exit Do
				    Loop
%>
         			<tr class='tdbg' onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
					 <td colspan=2 valign=top>
							&nbsp;<INPUT id="chkAll" onClick="CheckAll(this.form)" type="checkbox" value="checkbox"  name="chkAll">&nbsp;ѡ��<INPUT class="button" onClick="return(confirm('ȷ��ɾ��ѡ�е�<%=KS.C_S(ChannelID,3)%>��?'));" type=submit value="ɾ��ѡ��" name=submit1>  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  </FORM>       
					  </td>
					  <td colspan=10>
					<%Call KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,true)%>			
					  </td>
						
        </tr>
								<%
  End Sub
  
 
  '���
  Sub ShopAdd
        Call KSUser.InnerLocation("����"& KS.C_S(ChannelID,3) & "")
				Action=KS.S("Action")
				ID=KS.ChkClng(KS.S("ID"))
                 If Action="Edit" Then
				  CurrentOpStr=" OK,�޸� "
				  Action="EditSave"
				   Dim ShopRS:Set ShopRS=Server.CreateObject("ADODB.RECORDSET")
				   ShopRS.Open "Select * From KS_Product Where Inputer='" & KSUser.UserName &"' and ID=" & ID,Conn,1,1
				   IF ShopRS.Eof And ShopRS.Bof Then
				     call KS.Alert("�������ݳ���!",ComeUrl)
					 Exit Sub
				   Else
						If KS.C_S(ChannelID,42) =0 And ShopRS("Verific")=1 and KS.ChkClng(KS.U_S(KSUser.GroupID,1))=0 Then
						   ShopRS.Close():Set ShopRS=Nothing
						   Response.Redirect "../plus/error.asp?action=error&message=" & server.urlencode("��Ƶ�����������" & KS.C_S(ChannelID,3) & "�������޸�!")
						End If
							   
				        ClassID=ShopRS("TID")
						BigClassID=ShopRS("BigClassID")
						SmallClassID=ShopRS("SmallClassID")
						Title=Trim(ShopRS("Title"))
						UserClassID=ShopRS("ClassID")
						ShowOnSpace=ShopRS("ShowOnSpace")
						KeyWords=Trim(ShopRS("KeyWords"))
						ProModel=Trim(ShopRS("ProModel"))
						ProSpecificat=Trim(ShopRS("ProSpecificat"))
						Unit=Trim(ShopRS("Unit"))
						TotalNum=Trim(ShopRS("TotalNum"))
						AlarmNum=Trim(ShopRS("AlarmNum"))
						TrademarkName=Trim(ShopRS("TrademarkName"))
						Content=ShopRS("ProIntro")
						Verific  = ShopRS("Verific")
						PhotoUrl=ShopRS("PhotoUrl")
						BigPhoto=ShopRS("BigPhoto")
						ProductType=ShopRS("ProductType")
						ProducerName=Trim(ShopRS("ProducerName"))
						UserDefineFieldArr=KSUser.KS_D_F_Arr(ChannelID)
						Discount=Trim(ShopRS("Discount"))
						Price_Original=Trim(ShopRS("Price_Original"))
						Price=Trim(ShopRS("Price"))
						Price_Market=Trim(ShopRS("Price_Market"))
						Price_Member=Trim(ShopRS("Price_Member"))
						'ProductType=1:Discount=9:Hits = 0:TotalNum = 1000: AlarmNum = 10:Comment = 1
						
						If IsArray(UserDefineFieldArr) Then
						For I=0 To Ubound(UserDefineFieldArr,2)
						 Dim UnitOption
						  If UserDefineFieldArr(11,I)="1" Then
						   UnitOption="@" & ShopRS(UserDefineFieldArr(0,I)&"_Unit")
						  Else
						   UnitOption=""
						  End If
						  If UserDefineFieldValueStr="" Then
							UserDefineFieldValueStr=ShopRS(UserDefineFieldArr(0,I)) & UnitOption
						  Else
							UserDefineFieldValueStr=UserDefineFieldValueStr &  "||||" & ShopRS(UserDefineFieldArr(0,I)) & UnitOption
						  End If
						Next
					  End If
                   End If
				   SelButton=KS.C_C(ClassID,1)
				   ShopRS.Close:Set ShopRS=Nothing
				Else
				 Call KSUser.CheckMoney(ChannelID)
				 CurrentOpStr=" OK,��� "
				 Action="AddSave"
				 ProductType=1
				 ShowOnSpace=1
				 ClassID=KS.S("ClassID")
				 If ClassID="" Then ClassID="0"
				  SelButton="ѡ����Ŀ..."
				End IF	
		%>

				<script language = "JavaScript">
				function displaydiscount(){
			 if (document.myform.ProductType[2].checked==true)
			   $("#discountarea").show();
			 else
			   $("#discountarea").hide();
			}
			function getprice(Price_Original){
			  if(Price_Original==''|| isNaN(Price_Original)){Price_Original=0;}
			  if(document.myform.ProductType[2].checked==true){
			  document.myform.Price.value=Math.round(Price_Original*Math.abs(document.myform.Discount.value/10)*100)/100;}
			//  else if(document.myform.ProductType[3].checked==true){document.myform.Price.value=Math.round(Price_Original*Math.abs(document.myform.Discount.value/10)*100)/100;}
			  else{document.myform.Price.value=Price_Original;}
			}
			function regInput(obj, reg, inputStr)
			{
				var docSel = document.selection.createRange()
				if (docSel.parentElement().tagName != "INPUT")    return false
				oSel = docSel.duplicate()
				oSel.text = ""
				var srcRange = obj.createTextRange()
				oSel.setEndPoint("StartToStart", srcRange)
				var str = oSel.text + inputStr + srcRange.text.substr(oSel.text.length)
				return reg.test(str)
			}

			function PreViewPic(ImgUrl)
			{
			if (ImgUrl!=''&&ImgUrl!=null)
			  {   if (ImgUrl==1)
				   {  if (document.myform.PicUrl.length>0&&document.myform.PicUrl.value!='')
					   document.all.PicViewArea.innerHTML='<img src='+document.myform.PicUrl.value.split('|')[1]+' border=0>'
					  else
					   return
					}
				  else
				  if (ImgUrl!='')
				 {document.all.PicViewArea.innerHTML='<img src='+ImgUrl+' border=0>';}
			  }
			}
			function GetFileNameArea(val)
			{
			  if (val==0)
			  {
			   $('filearea').style.display='none';
			  }
			  else
			  {
			   $('filearea').style.display='';
			  }
			}
			function GetTemplateArea(val)
			{
			  if (val==2)
			  {
			   $('templatearea').style.display='none';
			  }
			  else
			  {
			   $('templatearea').style.display='';
			  }
			}
				 function OpenThenSetValue(Url,Width,Height,WindowObj,SetObj)
					{
						var ReturnStr=showModalDialog(Url,WindowObj,'dialogWidth:'+Width+'pt;dialogHeight:'+Height+'pt;Verific:no;help:no;scroll:no;Verific:0;help:0;scroll:0;');
						if (ReturnStr!='') SetObj.value=ReturnStr;
					}
				function CheckClassID()
				{
				if (document.myform.ClassID.value=="0") 
				  {
					alert("��ѡ��<%=KS.C_S(ChannelID,3)%>��Ŀ��");
					return false;
				  }		
				  return true;
				}
			 
				function CheckForm()
				{
				<%If KS.C_S(ChannelID,34)=0 Then%>	
				if (frames["ShopContent"].CurrMode!='EDIT') {alert('����ģʽ���޷����棬���л������ģʽ');return false;}
                document.myform.Content.value=frames["ShopContent"].KS_EditArea.document.body.innerHTML;
				<%end if%>
				if (document.myform.ClassID.value=="0") 
				  {
					alert("��ѡ��<%=KS.C_S(ChannelID,3)%>��Ŀ��");
					document.myform.ClassID.focus();
					return false;
				  }		
				if (document.myform.Title.value=="")
				  {
					alert("������<%=KS.C_S(ChannelID,3)%>���ƣ�");
					document.myform.Title.focus();
					return false;
				  }		
				  if (document.myform.KeyWords.value=="")
				  {
					alert("������<%=KS.C_S(ChannelID,3)%>�ؼ��֣�");
					document.myform.KeyWords.focus();
					return false;
				  }	
				  <%Call KSUser.ShowUserFieldCheck(ChannelID)%>
				  if (document.myform.ProModel.value=="")
				  {
					alert("������<%=KS.C_S(ChannelID,3)%>�ͺţ�");
					document.myform.ProModel.focus();
					return false;
				  }	

				  if (document.myform.ProSpecificat.value=="")
				  {
					alert("������<%=KS.C_S(ChannelID,3)%>���");
					document.myform.ProSpecificat.focus();
					return false;
				  }
				  if (document.myform.ProducerName.value=="")
				  {
					alert("������<%=KS.C_S(ChannelID,3)%>�����̣�");
					document.myform.ProducerName.focus();
					return false;
				  }
				  if (document.myform.Unit.value=="")
				  {
					alert("������<%=KS.C_S(ChannelID,3)%>��λ��");
					document.myform.Unit.focus();
					return false;
				  }
				  if (document.myform.TotalNum.value=="")
				  {
					alert("������<%=KS.C_S(ChannelID,3)%>��棡");
					document.myform.TotalNum.focus();
					return false;
				  }
				  if (document.myform.AlarmNum.value=="")
				  {
					alert("������<%=KS.C_S(ChannelID,3)%>��棡");
					document.myform.AlarmNum.focus();
					return false;
				  }
				  if (document.myform.Price_Original.value=="")
				  {
					alert("������<%=KS.C_S(ChannelID,3)%>ԭʼ���ۼۣ�");
					document.myform.Price_Original.focus();
					return false;
				  }
				  document.myform.submit();
				 return true;  
				}
				</script>

				<table width="98%" border="0" align="center" cellpadding="3" cellspacing="1" class="border">
				  <tr class="title">
				   <td colspan=2 align=center>
							 <%IF KS.S("Action")="Edit" Then
							   response.write "�޸�" & KS.C_S(ChannelID,3)
							   Else
							    response.write "����" & KS.C_S(ChannelID,3)
							   End iF
							  %>				   </td>
				  </tr> 
                  <form  action="User_MyShop.asp?Action=<%=Action%>" method="post" name="myform" id="myform">
				    <input type="hidden" name="ID" value="<%=ID%>">
				    <input type="hidden" name="comeurl" value="<%=ComeUrl%>">
				<tr class="tdbg">
                           <td height="25" align="center">������Ŀ��</td>
                           <td>��
						
							 <% Call KSUser.GetClassByGroupID(ChannelID,ClassID,Selbutton) %>
							 				  </td>
                    </tr>
                                <tr class="tdbg">

                                        <td height="25" align="center"><span><%=KS.C_S(ChannelID,3)%>���ƣ�</span></td>
                                        <td> �� 
                                          <input name="Title" class="textbox" type="text" id="Title" value="<%=Title%>" style="width:250px; " maxlength="100" />
                                          <span style="color: #FF0000">*</span></td>
                                </tr>
								<tr class="tdbg">
								   <td width="12%"  height="25" align="center"><span>�ҵķ��ࣺ</span></td>
								   <td colspan="2">��
								    
							<select class="textbox" size='1' name='UserClassID' style="width:150">
														<option value="0">-��ָ������-</option>
														<%=KSUser.UserClassOption(3,UserClassID)%>
									 </select>		
							
									 <a href="User_Class.asp?Action=Add&typeid=3"><font color="red">���</font></a>					                      </td>
								</tr>	
                                <tr class="tdbg">
                                        <td height="25" align="center"><span>�� �� �֣�</span></td>
                                  <td>��
                                          <input name="KeyWords" class="textbox" type="text" value="<%=KeyWords%>" id="KeyWords" style="width:250px; " />
                                          ����ؼ�������Ӣ�Ķ���(&quot;<span style="color: #FF0000">,</span>&quot;)���� </td>
                                </tr>
                                <tr class="tdbg">
                                        <td height="25" align="center"><span><%=KS.C_S(ChannelID,3)%>�ͺţ�</span></td>
                                        <td height="25">��
                                        <input name="ProModel" class="textbox" type="text" value="<%=ProModel%>" id="ProModel" style="width:250px; "  maxlength="30" />
                                        <span style="color: #FF0000">*</span></td>
                                </tr>
                                <tr class="tdbg">
                                        <td height="25" align="center"><span><%=KS.C_S(ChannelID,3)%>���</span></td>
                                        <td>��
                                        <input name="ProSpecificat" class="textbox" type="text" id="ProSpecificat" value="<%=ProSpecificat%>" style="width:250px; " maxlength="100" />
                                        <span style="color: #FF0000">*</span></td>
								</tr>
								
<%
							  Response.Write KSUser.KS_D_F(ChannelID,UserDefineFieldValueStr)
							  %>     
							  
								<tr class="tdbg">
								  <td height="25" align="center"><span>Ʒ��/�̱�:</span></td>
								  <td>&nbsp;&nbsp;&nbsp;<input name="TrademarkName" class="textbox" type="text" id="TrademarkName" value="<%=TrademarkName%>" style="width:250px; " maxlength="100" /></td>
				    </tr>
								<tr class="tdbg">
								  <td height="25" align="center"><span>������:</span></td>
								  <td>&nbsp;&nbsp;&nbsp;<input name="ProducerName" class="textbox" type="text" id="ProducerName" value="<%=ProducerName%>" style="width:250px; " maxlength="100" />
							      <span style="color: #FF0000">*</span></td>
				    </tr>
								<tr class="tdbg">
								  <td height="25" align="center">��Ʒ��λ:</td>
								  <td>&nbsp;&nbsp;&nbsp;<input name="Unit" type="text" class="textbox" id="Unit" style="width:40px; " value="<%=Unit%>" size="40" maxlength="40" />&nbsp;(��:��)<span style="color: #FF0000">*</span></td>
				    </tr>
								<tr class="tdbg">
								  <td height="25" align="center">�������:</td>
								  <td>&nbsp;&nbsp;&nbsp;�������&nbsp;<input name="TotalNum" type="text" class="textbox" id="TotalNum" style="width:40px; " value="<%=TotalNum%>" size="40" maxlength="40" />&nbsp;��汨��������&nbsp;<input name="AlarmNum" type="text" class="textbox" id="AlarmNum" style="width:40px; " value="<%=AlarmNum%>" size="40" maxlength="40" />
							      <span style="color: #FF0000">*</span></td>
				    </tr>
								<tr class="tdbg">
								  <td height="25" align="center">��Ʒ�۸�:</td>
								  <td>
								  <input name='ProductType' type='radio' onclick='document.myform.Price.value=document.myform.Price_Original.value;document.myform.Price.disabled=true;displaydiscount();' value='1' <%If ProductType=1 Then Response.Write " checked"%>>��������&nbsp;<input name='ProductType' type='radio' value='2'<%If ProductType=2 Then Response.Write " checked"%>onclick='document.myform.Price.disabled=false;displaydiscount();'>
								 <font color='green'> �Ǽ�����</font>
								 <input name='ProductType' type='radio' value='3'
								 <%If ProductType=3 Then Response.Write " checked"%>
onclick='getprice(document.myform.Price_Original.value);document.myform.Price.disabled=false;displaydiscount();'>
		<font color='red'> ��������&nbsp;
		<span id="discountarea"<%If ProductType<>3 Then Response.Write " style='display:none'" %>>
		�ۿ��ʣ�<input name="Discount" type='text' onKeyUp="getprice(document.all.Price_Original.value)" value="<%=Discount%>" size="4" maxlength="4" class="textbox" onKeyPress= "return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))" style="text-align: center;"> ��&nbsp;&nbsp;���鼮����Ʒ9�۴���</font></span>
        <span style="color: #FF0000">*</span>
        <hr size=1 color=blue>
<font color=red>ԭʼ���ۼ�
<input name="Price_Original" type="text" id="Price_Original" value="<%=Price_Original%>" size="6" onChange="getprice(this.value)" onKeyUp="getprice(this.value)" style="text-align: right;" onKeyPress= "return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))" class="textbox"> *Ԫ</font>&nbsp;
<font color=blue>��ǰ���ۼ�<input name="Price" type="text" id="Price" value="<%=Price%>" size="6" class="textbox"
		<%If ProductType=1 Then Response.Write " disabled"%>onKeyPress="return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))">Ԫ</font>&nbsp;&nbsp;�г���<input name="Price_Market" type="text" id="Price_Market" value="<%=Price_Market%>" size="6" class="textbox" onKeyPress="return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))">Ԫ ��Ա��<input name="Price_Member" type="text" id="Price_Member" value="<%=Price_Member%>" size="6" class="textbox" onKeyPress= "return regInput(this,/^\d*\.?\d{0,2}$/,String.fromCharCode(event.keyCode))" onpaste="return regInput(this,/^\d*\.?\d{0,2}$/,window.clipboardData.getData('Text'))" ondrop="return regInput(this,    /^\d*\.?\d{0,2}$/,event.dataTransfer.getData('Text'))">Ԫ
		</tr>
								<tr class="tdbg">
								  <td height="25" align="center">&nbsp;</td>
								  <td>&nbsp;</td>
				    </tr>
								<tr class="tdbg">
                                        <td height="25" align="center"><span>Сͼ��ַ��</span></td>
                                        <td>��
                                          <input class="textbox" name='PhotoUrl' value="<%=PhotoUrl%>" type='text' style="width:250px;" id='PhotoUrl' maxlength="100" />
                                          &nbsp;
                                          <input class="button" type='button' name='Submit3' value='ѡ��ͼƬ��ַ...' onClick="OpenThenSetValue('Frame.asp?url=SelectPhoto.asp&pagetitle=<%=Server.URLEncode("ѡ��ͼƬ")%>&channelid=5',500,360,window,document.myform.PhotoUrl);" /></td>
							   </tr>
								<tr class="tdbg">
                                        <td height="25" align="center"><span>��ͼ��ַ��</span></td>
                                        <td>��
                                          <input class="textbox" name='BigPhoto' value="<%=BigPhoto%>" type='text' style="width:250px;" id='BigPhoto' maxlength="100" />
                                          &nbsp;
                                          <input class="button" type='button' name='Submit3' value='ѡ��ͼƬ��ַ...' onClick="OpenThenSetValue('Frame.asp?url=SelectPhoto.asp&pagetitle=<%=Server.URLEncode("ѡ��ͼƬ")%>&channelid=5',500,360,window,document.myform.BigPhoto);" /></td>
							   </tr>
								<tr class="tdbg">
                                        <td height="25" align="center"><span>�ϴ�ͼƬ��</span></td>
                                        <td>��
  <iframe id='UpPhotoFrame' name='UpPhotoFrame' src='User_upfile.asp?channelid=5&Type=Pic' frameborder=0 scrolling=no width='95%' height='25'> </iframe></td></tr>
								
  								<tr class="tdbg">
                                        <td align="center"><span><%=KS.C_S(ChannelID,3)%>��飺<br />
                                          </span></td>
                                        <td>
										<table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                            <tr>
                                              <td width="12">&nbsp;</td>
                                              <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                                  <tr>
                                                    <td height="200" align="center">
										<%If KS.C_S(ChannelID,34)=0 Then%>			
										<textarea name="Content" style="display:none"><%=Server.HtmlEncode(Content)%></textarea>
                                        <iframe id='ShopContent' name='ShopContent' src='Editor.asp?ID=Content&style=0&ChannelID=9999' frameborder=0 scrolling=no width='100%' height='200'></iframe>
										<%Else
										   Response.Write "<textarea name=""Content"" style=""display:none"">" & KS.HtmlCode(Content) & "</textarea>"
								           Response.Write "<iframe id=""content___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=Content&amp;Toolbar=Basic"" width=""98%"" height=""200"" frameborder=""0"" scrolling=""no""></iframe>"  
										End If 
											 %>													</td>
                                                  </tr>
                                              </table></td>
                                            </tr>
                                        </table></td>
                                </tr>
<tr class="tdbg">
                                        <td height="25" align="center"><span>�ռ���ҳ��ʾ��</span></td>
                                        <td>��
	<input name="ShowOnSpace" type="radio" value="1" <%If ShowOnSpace="1" Then Response.Write " checked"%> />��
	<input name="ShowOnSpace" type="radio" value="0" <%If ShowOnSpace="0" Then Response.Write " checked"%>/>��					</td>
								</tr>
					  <%if KS.S("Action")="Edit" And Verific=1 Then%>
								<input type="hidden" name="okverific" value="1">
								<input type="hidden" name="verific" value="1">
								<%else%>
						<tr class="tdbg" >
                                        <td height="25" align="center"><span><%=KS.C_S(ChannelID,3)%>״̬��</span></td>
                                        <td height="25">
										 <input name="Status" type="radio" value="0" <%If Verific=0 Then Response.Write " checked"%> />
                                          Ͷ��
                                          <input name="Status" type="radio" value="2" <%If Verific=2 Then Response.Write " checked"%>/>
                                          �ݸ�
										  </td>
                                      </tr>
							  <%end if%>
                          <tr class="tdbg">
                            <td align="center" colspan=2><input class="button" type="button" onClick="return CheckForm();" name="Submit" value=" <%=CurrentOpStr%> " />
                            ��
                            <input class="button" type="reset" name="Submit2" value=" ���� " /></td>
                          </tr>
                  </form>
</table>
				
		  <%
  End Sub
  Sub ShopSave()
        Dim ID:ID=KS.ChkClng(KS.S("ID"))
  		ClassID=KS.S("ClassID")
		BigClassID=KS.ChkClng(KS.S("BigClassID"))
		SmallClassID=KS.ChkClng(KS.S("SmallClassID"))
		Title=KS.FilterIllegalChar(KS.LoseHtml(KS.S("Title")))
		KeyWords=KS.LoseHtml(KS.S("KeyWords"))
		ProModel=KS.LoseHtml(KS.S("ProModel"))
		ProSpecificat=KS.LoseHtml(KS.S("ProSpecificat"))
		Unit=KS.LoseHtml(KS.S("Unit"))
		TotalNum=KS.ChkClng(KS.S("TotalNum"))
		AlarmNum=KS.ChkClng(KS.S("AlarmNum"))
		TrademarkName=KS.LoseHtml(KS.S("TrademarkName"))
		Content=KS.FilterIllegalChar(Request.Form("Content"))
		ProducerName=KS.LoseHtml(KS.S("ProducerName"))
		UserClassID=KS.ChkClng(KS.S("UserClassID"))
		ShowOnSpace=KS.ChkClng(KS.S("ShowOnSpace"))
		Verific=KS.ChkClng(KS.S("Status"))
        If KS.C_S(ChannelID,17)<>0 And Verific=0 Then Verific=1
		 If KS.ChkClng(KS.S("ID"))<>0 and verific=1  Then
			 If KS.C_S(ChannelID,42)=2 Then Verific=1 Else Verific=0
		 End If
		 if KS.C_S(ChannelID,42)=2 and KS.ChkClng(KS.S("okverific"))=1 Then verific=1
		 If KS.ChkClng(KS.U_S(KSUser.GroupID,0))=1 Then verific=1  '����VIP�û��������
		PhotoUrl=KS.S("PhotoUrl")
		BigPhoto=KS.S("BigPhoto")

		ProductType=KS.ChkClng(KS.S("ProductType"))
		If ProductType<>3 Then
			 Discount=10
			Else 
			 Discount=KS.G("Discount")
			End If
			Price_Original = KS.G("Price_Original")
			If ProductType=1 Then
			 Price=Price_Original
			ElseIf (ProductType=2 Or ProductType=3) And KS.G("Price")="" Then
			 Price=Price_Original
			Else
			 Price = KS.G("Price")
			End If
			Price_Market = KS.G("Price_Market"):If Price_Market="" Then Price_Market=0
			Price_Member = KS.G("Price_Member"):If Price_Member="" Then Price_Member=0
			If Discount>10 Then ErrMsg = ErrMsg & "��Ʒ���ۿ��ʱ���С��10! \n"
			If ProductType=2 And KS.ChkClng(Price)<KS.ChkClng(Price_Original) Then ErrMsg = ErrMsg & "�Ǽ�����,��Ʒ�ġ���ǰ���ۼۡ�������ڵ��ڡ�ԭʼ���ۼۡ�! \n"
			If ProductType=3 And KS.ChkClng(Price_Member)>KS.ChkClng(Price) Then ErrMsg = ErrMsg & "��������,��Ʒ�ġ���Ա�ۡ�����С�ڵ��ڡ���ǰ���ۼۡ�! \n"
			
			If Not IsNumeric(Price_Original) Then Call KS.AlertHistory("ԭʼ���ۼ۱���������!",-1) : Exit Sub
			If Not IsNumeric(Price) Then Call KS.AlertHistory("��ǰ���ۼ۱���������!",-1) : Exit Sub
			If Not IsNumeric(Price_Member) Then Call KS.AlertHistory("��Ա�۱���������!",-1) : Exit Sub
			If Not IsNumeric(Price_Market) Then Call KS.AlertHistory("�г��۱���������!",-1) : Exit Sub
			
			
			
				  if ClassID="" Then ClassID=0
				  If ClassID=0 Then
				    Response.Write "<script>alert('��û��ѡ��"& KS.C_S(ChannelID,3) & "��Ŀ!');history.back();</script>"
				    Exit Sub
				  End IF
				  If Title="" Then
				    Response.Write "<script>alert('��û������"& KS.C_S(ChannelID,3) & "����!');history.back();</script>"
				    Exit Sub
				  End IF
				  
				  
				UserDefineFieldArr=KSUser.KS_D_F_Arr(ChannelID)
				If IsArray(UserDefineFieldArr) Then
				For I=0 To Ubound(UserDefineFieldArr,2)
				If UserDefineFieldArr(6,I)=1 And KS.G(UserDefineFieldArr(0,I))="" Then ErrMsg = ErrMsg & UserDefineFieldArr(1,I) & "������д!\n"
							
				 If UserDefineFieldArr(6,I)=1 And KS.S(UserDefineFieldArr(0,I))="" Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "������д!');history.back();</script>":Exit Sub
				 
				 If UserDefineFieldArr(3,I)=4 And Not Isnumeric(KS.S(UserDefineFieldArr(0,I))) Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "������д����!');history.back();</script>":Exit Sub
				 If UserDefineFieldArr(3,I)=5 And Not IsDate(KS.S(UserDefineFieldArr(0,I))) Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "������д��ȷ������!');history.back();</script>":Exit Sub
				If UserDefineFieldArr(3,I)=8 And Not KS.IsValidEmail(KS.S(UserDefineFieldArr(0,I))) and UserDefineFieldArr(6,I)=1 Then Response.Write "<script>alert('" & UserDefineFieldArr(1,I) & "������д��ȷ��Email!');history.back();</script>":Exit Sub
				Next
				End If				  
				Set RSObj=Server.CreateObject("Adodb.Recordset")
			
				 Dim Fname,FnameType,TemplateID,WapTemplateID
			  If ID=0 Then
				 Dim RSC:Set RSC=Server.CreateObject("ADODB.RECORDSET")
				 RSC.Open "select top 1 * from KS_Class Where ID='" & ClassID & "'",conn,1,1
				 if RSC.Eof Then 
				  Response.end
				 Else
				 FnameType=RSC("FnameType")
				 Fname=KS.GetFileName(RSC("FsoType"), Now, FnameType)
				 TemplateID=RSC("TemplateID")
				 WapTemplateID=RSC("WapTemplateID")
				 End If
				 RSC.Close:Set RSC=Nothing
			 End If

				RSObj.Open "Select top 1 * From KS_Product Where Inputer='" & KSUser.UserName & "' and ID=" & ID,Conn,1,3
				If RSObj.Eof And RSObj.Bof Then
				   RSObj.AddNew
				     RSObj("ProID")=KS.GetInfoID(ChannelID)   'ȡΨһID
				     RSObj("Hits")=0
					 RSObj("Rolls")=0
					 RSObj("Recommend")=0
					 RSObj("Popular")=0
					 RSObj("Slide")=0
					 RSObj("Comment")=1
					 RSObj("IsSpecial")=0
					 RSObj("ISTop")=0
					 RSObj("Fname") = Fname
					 RSObj("AddDate")=Now
					 RSObj("Rank")="����"
					 RSObj("Point") = 0
					 RSObj("TemplateID") = TemplateID
					 RSObj("WapTemplateID")=WapTemplateID
				End If
					 RSObj("Title") = Title
					 RSObj("PhotoUrl") = PhotoUrl
					 RSObj("BigPhoto") = BigPhoto
					 RSObj("ProIntro") = Content
					 RSObj("Verific") = Verific
					 RSObj("Tid") = ClassID
					 RSObj("TotalNum") = TotalNum
					 RSObj("AlarmNum") = AlarmNum
					 RSObj("ProductType") = ProductType
					 RSObj("Discount") = Discount
					 RSObj("Unit") = Unit
					 RSObj("Price_Original") = Price_Original
					 RSObj("Price") = Price
					 RSObj("Price_Member")=Price_Member
					 RSObj("Price_Market") = Price_Market
					 RSObj("KeyWords") = KeyWords
					 RSObj("ProSpecificat")=ProSpecificat
					 RSObj("ProModel") = ProModel
					 RSObj("TrademarkName") = TrademarkName
					 RSObj("Inputer")=KSUser.UserName
					 RSObj("ProducerName")=ProducerName
					 RSObj("ClassID")=UserClassID
					 RSOBj("ShowOnSpace")=ShowOnSpace
					 RSOBj("BigClassID")=BigClassID
					 RSObj("SmallClassID")=SmallClassID
					 
				     If IsArray(UserDefineFieldArr) Then
						For I=0 To Ubound(UserDefineFieldArr,2)
						 RSObj("" & UserDefineFieldArr(0,I) & "")=KS.S(UserDefineFieldArr(0,I))
						 If UserDefineFieldArr(11,I)="1"  Then
							RSObj("" & UserDefineFieldArr(0,I) & "_Unit")=KS.G(UserDefineFieldArr(0,I)&"_Unit")
						 End If
				  		Next
				     End If
				  
				RSObj.Update
				RSObj.MoveLast
				Dim InfoID:InfoID=RSObj("ID")
				If Left(Ucase(Fname),2)="ID" and ID=0 Then
					RSObj("Fname") = InfoID & FnameType
					RSObj.Update
				End If
				Fname=RSOBj("Fname")
				
				If Verific=1 Then 
				    Call KS.SignUserInfoOK(ChannelID,KSUser.UserName,Title,InfoID)
					If KS.C_S(ChannelID,17)=2  and (KS.C_S(Channelid,7)=1 or KS.C_S(ChannelID,7)=2) Then
					 Dim KSRObj:Set KSRObj=New Refresh
					 Dim DocXML:Set DocXML=KS.RsToXml(RSObj,"row","root")
				     Set KSRObj.Node=DocXml.DocumentElement.SelectSingleNode("row")
					  KSRObj.ModelID=ChannelID
					  KSRObj.ItemID = KSRObj.Node.SelectSingleNode("@id").text 
					  Call KSRObj.RefreshContent()
					  Set KSRobj=Nothing
					End If
				End If
				 RSObj.Close:Set RSObj=Nothing
				 
               If ID=0 Then
			     Call LFCls.InserItemInfo(ChannelID,InfoID,Title,ClassId,Content,KeyWords,PhotoUrl,KSUser.UserName,Verific,Fname)
  		         Call KS.FileAssociation(ChannelID,InfoID,PhotoUrl & BigPhoto & Content ,0)
			     Call KSUser.AddLog(KSUser.UserName,"����Ŀ[<a href='" & KS.GetFolderPath(ClassID) & "' target='_blank'>" & KS.C_C(ClassID,1) & "</a>]������" & KS.C_S(ChannelID,3) & """<a href='../item/Show.asp?m=" & ChannelID & "&d=" & InfoID & "' target='_blank'>" & Title & "</a>""!",5)
				 KS.Echo "<script>if (confirm('"& KS.C_S(ChannelID,3) & "��ӳɹ������������?')){location.href='User_MyShop.asp?Action=Add&ClassID=" & ClassID &"';}else{location.href='User_MyShop.asp';}</script>"
			  Else
			     Call LFCls.ModifyItemInfo(ChannelID,InfoID,Title,classid,Content,KeyWords,PhotoUrl,Verific)
				 Call KS.FileAssociation(ChannelID,InfoID,PhotoUrl & BigPhoto & Content ,1)
			     Call KSUser.AddLog(KSUser.UserName,"��" & KS.C_S(ChannelID,3) & """<a href='../item/Show.asp?m=" & ChannelID & "&d=" & InfoID & "' target='_blank'>" & Title & "</a>""�����޸�!",5)
				KS.Echo "<script>alert('"& KS.C_S(ChannelID,3) & "�޸ĳɹ�!');location.href='" & ComeUrl & "';</script>"
			  End If
		
  End Sub
End Class
%> 
