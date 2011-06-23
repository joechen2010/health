<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Response.ContentType = "text/vnd.wap.wml; charset=utf-8"%><?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml">
<wml>
<head>
<meta http-equiv="Cache-Control" content="no-Cache"/>
<meta http-equiv="Cache-Control" content="max-age=0"/>
</head>
<card id="main" title="发布产品">
<p>
<%
Dim KSCls
Set KSCls = New Admin_MyShop
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_MyShop
        Private KS,KSUser,ChannelID,LoginTF
		Private CurrentPage,totalPut,Status,ProducerName,Prev
		Private RS,MaxPerPage,ComeUrl,SelButton,Price_Original,Price,Price_Market,Price_Member,Point,Discount
		Private ClassID,Title,KeyWords,ProModel,ProSpecificat,ProductType,Unit,TotalNum,AlarmNum,TrademarkName,Content,Verific,PhotoUrl,RSObj,I,UserDefineFieldArr,UserDefineFieldValueStr,UserClassID,ShowONSpace
		Private CurrentOpStr,Action,ID,ErrMsg,Hits,BigPhoto,BigClassID,SmallClassID,flag
		Private Sub Class_Initialize()
			MaxPerPage =5
		  Set KS=New PublicCls
		  Set KSUser = New UserCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSUser=Nothing
		End Sub
		Public Sub Kesion()

		    ChannelID=5
			LoginTF=Cbool(KSUser.UserLoginChecked)
			IF LoginTF=false  Then
			   Response.redirect KS.GetDomain&"User/Login/?User_MyShop.asp"
			   Exit Sub
			End If
			If KS.C_S(ChannelID,36)=0 Then
			   Response.Write "本频道不允许投稿!<br/>"
			   Exit Sub
			End If
			Verific=KS.S("status")
			If Verific="" or not isnumeric(Verific) Then Verific=4
			
		%>
		<a href="User_MyShop.asp?ChannelID=<%=ChannelID%>&amp;<%=KS.WapValue%>">我发布的<%=KS.C_S(ChannelID,3)%>(<%=Conn.Execute("Select count(id) from " & KS.C_S(ChannelID,2) &" where Inputer='"& KSUser.UserName &"'")(0)%>)</a>
		<%If Verific=1 Then%>
		【
		<%End If%><a href="User_MyShop.asp?ChannelID=<%=ChannelID%>&amp;Status=1&amp;<%=KS.WapValue%>">已审核(<%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=1 and Inputer='"& KSUser.UserName &"'")(0)%>)</a><%If Verific=1 Then%>
		】
		<%End If%><%If Verific=0 Then%>
		【
		<%End If%>
		<a href="User_MyShop.asp?ChannelID=<%=ChannelID%>&amp;Status=0&amp;<%=KS.WapValue%>">待审核(<%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=0 and Inputer='"& KSUser.UserName &"'")(0)%>)</a><%If Verific=0 Then%>
		】
		<%End If%><%If Verific=2 Then%>
		【
		<%End If%>
		<a href="User_MyShop.asp?ChannelID=<%=ChannelID%>&amp;Status=2&amp;<%=KS.WapValue%>">草 稿(<%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=2 and Inputer='"& KSUser.UserName &"'")(0)%>)</a><%If Verific=2 Then%>
		】
		<%End If%><%If Verific=3 Then%>
		【
		<%End If%>
		<a href="User_MyShop.asp?ChannelID=<%=ChannelID%>&amp;Status=3&amp;<%=KS.WapValue%>">被退稿(<%=conn.execute("select count(id) from " & KS.C_S(ChannelID,2) &" where Verific=3 and Inputer='"& KSUser.UserName &"'")(0)%>)</a><%If Verific=3 Then%>
		】
		<%End If%>
		<%
		Action=KS.S("Action")
		Select Case Action
		 Case "Del"
		  Call KSUser.DelItemInfo(ChannelID)
		 Case "Add","Edit"
		  Call ShopAdd
		 Case "DoSave"
          Call DoSave()
		 Case Else
		  Call ShopList
		 End Select
			Response.Write "<br/>"
			If Prev=True Then
			   Response.Write "<anchor>返回上级<prev/></anchor><br/>"
			End If
			Response.Write "<a href=""Index.asp?" & KS.WapValue & """>我的地盘</a><br/>" &vbcrlf
			Response.Write "<a href=""" & KS.GetGoBackIndex & """>返回首页</a>" &vbcrlf
		 %>
		 </p>
</card>
</wml>
		 <%
       End Sub
	   Sub ShopList
			   		       If KS.S("page") <> "" Then
						          CurrentPage = KS.ChkClng(KS.S("page"))
							Else
								  CurrentPage = 1
							End If
                                    
									Dim Param:Param=" Where Inputer='"& KSUser.UserName &"'"
                                    IF Verific<>4 Then 
									   Param= Param & " and Verific=" & Verific
									End If
									IF KS.S("Flag")<>"" Then
									  IF KS.S("Flag")=0 Then Param=Param & " And Title like '%" & KS.S("KeyWord") & "%'"
									  IF KS.S("Flag")=1 Then Param=Param & " And KeyWords like '%" & KS.S("KeyWord") & "%'"
									End if
									If KS.S("ClassID")<>"" And KS.S("ClassID")<>"0" Then Param=Param & " And TID='" & KS.S("ClassID") & "'"
									Dim Sql:sql = "select a.*,foldername from KS_Product a inner join ks_class b on a.tid=b.id "& Param &" order by AddDate DESC"

                                    Response.Write "<br/>"
							
			  
			   %>
			    【<img src="images/add.gif" align="absmiddle" /><a href="user_myshop.asp?ChannelID=<%=ChannelID%>&amp;Action=Add&amp;<%=KS.WapValue%>">发布<%=KS.C_S(ChannelID,3)%>】</a><br/>

			<%
			
			Set RS=Server.CreateObject("AdodB.Recordset")
			RS.open sql,conn,1,1
			If RS.EOF And RS.BOF Then
			   Response.Write "没有你要的" & KS.C_S(ChannelID,3) & "!<br/>"
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
			   If CurrentPage > 1  and (CurrentPage - 1) * MaxPerPage < totalPut Then
			      RS.Move (CurrentPage - 1) * MaxPerPage
			   Else
			      CurrentPage = 1
			   End If
			   Call showContent
		    End If
			%>
            【<%=KS.C_S(ChannelID,3)%>搜索】<br/>
            <select name="Flag"><option value="0">标题</option><option value="1">关键字</option></select>
            关键字<input type="text" name="KeyWord" value="关键字"/>
            <anchor>搜索<go href="User_MyShop.asp?ChannelID=<%=ChannelID%>&amp;<%=KS.WapValue%>" method="post">
            <postfield name="Flag" value="$(Flag)"/>
            <postfield name="KeyWord" value="$(KeyWord)"/>
            </go></anchor><br/>	
			<%
		End Sub
		
		Sub ShowContent()
		    Dim I
			Do While Not RS.Eof
			%>
            <%=((I+1)+CurrentPage*MaxPerPage)-MaxPerPage%>.
            <a href="../Show.asp?ID=<%=RS("ID")%>&amp;ChannelID=<%=ChannelID%>&amp;<%=KS.WapValue%>"><%=KS.GotTopic(trim(RS("title")),35)%></a>
            <%
			Select Case RS("Verific")
			    Case 0:Response.Write "待审 "
				Case 1:Response.Write "已审 "
				Case 2:Response.Write "草稿 "
				Case 3:Response.Write "退稿 "
			End Select
			If RS("Verific")<>1 Then
			   Response.Write "<a href=""User_MyShop.asp?ChannelID="&ChannelID&"&amp;ID="&RS("ID")&"&amp;Action=Edit&amp;page="&CurrentPage&"&amp;"&KS.WapValue&""">修改</a>"
			   Response.Write "<a href=""User_MyShop.asp?ChannelID="&ChannelID&"&amp;Action=Del&amp;ID="&RS("ID")&"&amp;"&KS.WapValue&""">删除</a>"
			Else
			   If KS.C_S(ChannelID,42)=0 Then
			      Response.write "---"
			   Else
			      Response.Write "<a href=""User_MyShop.asp?ChannelID=" & ChannelID & "&amp;ID=" & RS("ID") &"&amp;Action=Edit&amp;page=" & CurrentPage &"&amp;"&KS.WapValue&""">修改</a>"
			   End If
			End If
			%>
            <br/>
            分类:<%=RS("FolderName")%>
            时间:<%=formatdatetime(rs("AddDate"),2)%><br/>
  <%
			RS.MoveNext
			I = I + 1
			If I >= MaxPerPage Then Exit Do
			Loop
			Call  KS.ShowPageParamter(totalPut, MaxPerPage, "User_MyShop.asp", True, KS.C_S(ChannelID,4), CurrentPage, "ChannelID=" & ChannelID & "&amp;Status=" & Verific & "&amp;" & KS.WapValue & "")
		End Sub
  
 
  '添加
  Sub ShopAdd
				Action=KS.S("Action")
				ID=KS.ChkClng(KS.S("ID"))
                 If Action="Edit" Then
				  CurrentOpStr=" OK,修改 "
				  Action="EditSave"
				   Dim ShopRS:Set ShopRS=Server.CreateObject("ADODB.RECORDSET")
				   ShopRS.Open "Select top 1 * From KS_Product Where Inputer='" & KSUser.UserName &"' and ID=" & ID,Conn,1,1
				   IF ShopRS.Eof And ShopRS.Bof Then
					 Response.Write "本频道设置已审核" & KS.C_S(ChannelID,3) & "不允许修改!<br/>"
					 Prev=True
					 Exit Sub
				   Else
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
						  If UserDefineFieldValueStr="" Then
							UserDefineFieldValueStr=ShopRS(UserDefineFieldArr(0,I))
						  Else
							UserDefineFieldValueStr=UserDefineFieldValueStr & "||||" & ShopRS(UserDefineFieldArr(0,I))
						  End If
						Next
					  End If
                   End If
				   SelButton=KS.C_C(ClassID,1)
				   ShopRS.Close:Set ShopRS=Nothing
				Else
				 Call KSUser.CheckMoney(ChannelID)
				 CurrentOpStr=" OK,添加 "
				 Action="AddSave"
				 ProductType=1
				 ShowOnSpace=1
				 Verific=0
				 ClassID=KS.S("ClassID")
				 If ClassID="" Then ClassID="0"
				  SelButton="选择栏目..."
				End IF	
		               IF KS.S("Action")="Edit" Then
							response.write "<br/>【修改" & KS.C_S(ChannelID,3)& "】<br/>"
					   Else
							response.write "<br/>【发布" & KS.C_S(ChannelID,3)& "】<br/>"
					 End iF
					 
            Response.Write "所属栏目："
			Call KSUser.GetClassByGroupID(ChannelID,ClassID,Selbutton)
			Response.Write "<br/>"					 
			%>
			<%=KS.C_S(ChannelID,3)%>名称：<input name="Title" type="text" value="<%=Title%>" emptyok="false" maxlength="100" /> <br/>
			
			我的分类：<select name="UserClassID">
			<option value="0">-不指定分类-</option>
			<%=KSUser.UserClassOption(3,UserClassID)%>
			</select><a href="User_Class.asp?Action=Add&amp;typeid=3&amp;<%=KS.WapValue%>"><font color="red">添加</font></a>	<br/> 				            
			关 键 字：<input name="KeyWords" type="text" value="<%=KeyWords%>" id="KeyWords" maxlegnth="200" />多个关键字请用英文逗号(&quot;<span style="color: #FF0000">,</span>&quot;)隔开 <br/>
            <%=KS.C_S(ChannelID,3)%>型号：<input name="ProModel" type="text" value="<%=ProModel%>"/> <br/>
			
			<%=KS.C_S(ChannelID,3)%>规格：<input name="ProSpecificat" type="text" value="<%=ProSpecificat%>"  maxlength="100" /><br/>
            <%
			 Response.Write KSUser.KS_D_F(ChannelID,UserDefineFieldValueStr)
			 %>     
			品牌/商标:<input name="TrademarkName" type="text" value="<%=TrademarkName%>"  maxlength="100" /><br/>
			生产商:<input name="ProducerName" type="text"  value="<%=ProducerName%>" maxlength="100" /><br/>
			商品单位:<input name="Unit" type="text" value="<%=Unit%>" size="10" maxlength="40" />例:本<br/>
			库存设置:库存数量<input name="TotalNum" type="text" value="<%=TotalNum%>" size="10" format="*N" maxlength="40" />库存报警下限<input name="AlarmNum" type="text" id="AlarmNum" value="<%=AlarmNum%>" size="10" maxlength="40" format="*N"/><br/>
			商品价格:
		原始零售<input name="Price_Original" type="text" value="<%=Price_Original%>" size="6" format="*N" emptyok="false"/>元 当前零售价<input name="Price" type="text" value="<%=Price%>" size="6" emptyok="false"/>元 市场价<input name="Price_Market" type="text" emptyok="false" value="<%=Price_Market%>" size="6" />元 会员价<input name="Price_Member" type="text" value="<%=Price_Member%>" size="6" emptyok="false"/>元<br/>
		
		
		
	<%=KS.C_S(ChannelID,3)%>简介：<input name="Content" type="text" value="<%=Content%>" /><br/>
    
	空间首页显示：<input name="ShowOnSpace" type="radio" value="1" <%If ShowOnSpace="1" Then Response.Write " checked=""checked"""%> />是
	<input name="ShowOnSpace" type="radio" value="0" <%If ShowOnSpace="0" Then Response.Write " checked=""checked"""%>/>否	<br/>
					           <%if KS.S("Action")="Edit" And Verific=1 Then%>
							  <%else%>
						
										 <input name="Status" type="radio" value="0" <%If Verific=0 Then Response.Write " checked=""checked"""%> />
                                          投搞
                                          <input name="Status" type="radio" value="2" <%If Verific=2 Then Response.Write " checked=""checked"""%>/>
                                          草稿
										 
							  <%end if%>
							<br/>
            <anchor>确定保存<go href="User_MyShop.asp?ChannelID=<%=ChannelID%>&amp;Action=DoSave&amp;ID=<%=ID%>&amp;<%=KS.WapValue%>" method="post">
            <postfield name="ClassID" value="$(ClassID)"/>
			<postfield name="UserClassID" value="$(UserClassID)"/>
            <postfield name="Title" value="$(Title)"/>
            <postfield name="KeyWords" value="$(KeyWords)"/>
            <postfield name="ProModel" value="$(ProModel)"/>
            <postfield name="ProSpecificat" value="$(ProSpecificat)"/>
            <%
			'自定义字段
			If IsArray(UserDefineFieldArr) Then
			   For I=0 To Ubound(UserDefineFieldArr,2)
			       Response.Write "<postfield name=""" & UserDefineFieldArr(0,I) & """ value=""$(" & UserDefineFieldArr(0,I) & ""&Minute(Now)&Second(Now)&")""/>"
			   Next
			End If
			%>
			<postfield name="TrademarkName" value="$(TrademarkName)"/>
			<postfield name="ProducerName" value="$(ProducerName)"/>
			<postfield name="Unit" value="$(Unit)"/>
			<postfield name="TotalNum" value="$(TotalNum)"/>
			<postfield name="AlarmNum" value="$(AlarmNum)"/>
			<postfield name="Price_Original" value="$(Price_Original)"/>
			<postfield name="Price" value="$(Price)"/>
			<postfield name="Price_Market" value="$(Price_Market)"/>
			<postfield name="Price_Member" value="$(Price_Member)"/>
            <postfield name="Content" value="$(Content)"/>
            <postfield name="ShowOnSpace" value="$(ShowOnSpace)"/>
			<%if KS.S("Action")="Edit" And Verific=1 Then%>
            <postfield name="okverific" value="1"/>
            <postfield name="verific" value="1"/>
			<%else%>
            <postfield name="Status" value="$(Status)"/>
			<%end if%>
            </go></anchor>
            <br/>
				
		  <%
  End Sub
  Sub DoSave()
        Dim ID:ID=KS.ChkClng(KS.S("ID"))
  		ClassID=KS.S("ClassID")
		Title=KS.LoseHtml(KS.S("Title"))
		KeyWords=KS.LoseHtml(KS.S("KeyWords"))
		ProModel=KS.LoseHtml(KS.S("ProModel"))
		ProSpecificat=KS.LoseHtml(KS.S("ProSpecificat"))
		Unit=KS.LoseHtml(KS.S("Unit"))
		TotalNum=KS.ChkClng(KS.S("TotalNum"))
		AlarmNum=KS.ChkClng(KS.S("AlarmNum"))
		TrademarkName=KS.LoseHtml(KS.S("TrademarkName"))
		Content=Request.Form("Content")
		ProducerName=KS.LoseHtml(KS.S("ProducerName"))
		UserClassID=KS.ChkClng(KS.S("UserClassID"))
		ShowOnSpace=KS.ChkClng(KS.S("ShowOnSpace"))
		Verific=KS.ChkClng(KS.S("Status"))
        If KS.C_S(ChannelID,17)<>0 And Verific=0 Then Verific=1
		 If KS.ChkClng(KS.S("ID"))<>0 and verific=1  Then
			 If KS.C_S(ChannelID,42)=2 Then Verific=1 Else Verific=0
		 End If
		 if KS.C_S(ChannelID,42)=2 and KS.ChkClng(KS.S("okverific"))=1 Then verific=1
		PhotoUrl=KS.S("PhotoUrl")
		BigPhoto=KS.S("BigPhoto")

		ProductType=KS.ChkClng(KS.S("ProductType"))
		If ProductType=0 Then ProductType=1
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
			If Discount>10 Then 
				 Response.Write "商品的折扣率必须小于10! <br/>"
				 Prev=True
				 Exit Sub
			End If
			If ProductType=2 And KS.ChkClng(Price)<KS.ChkClng(Price_Original) Then 
			    Response.Write "涨价销售,商品的“当前零售价”必须大于等于“原始零售价”! <br/>"
			 	 Prev=True
				 Exit Sub
			End If
			If ProductType=3 And KS.ChkClng(Price_Member)>KS.ChkClng(Price) Then 
			  response.write  "降价销售,商品的“会员价”必须小于等于“当前零售价”! <br/>"
			  Prev=True
			  Exit Sub
			End If
			
			If Not IsNumeric(Price_Original) Then 
			  Response.Write "原始零售价必须填数字!<br/>"
			  Prev=True
			  Exit Sub
			End If
			If Not IsNumeric(Price) Then 
			  Response.Write "当前零售价必须填数字!<br/>"
			  Prev=True : Exit Sub
			End If
			If Not IsNumeric(Price_Member) Then 
			  Response.Write "会员价必须填数字!<br/>"
			  Prev=True:Exit Sub
			End If
			If Not IsNumeric(Price_Market) Then 
			 Response.Write "市场价必须填数字!<br/>"
			 Prev=True: Exit Sub
			End If
			
			
			
				  if ClassID="" Then ClassID=0
				  If ClassID=0 Then
				    Response.Write "你没有选择"& KS.C_S(ChannelID,3) & "栏目!<br/>"
					Prev=True
				    Exit Sub
				  End IF
				  If Title="" Then
				    Response.Write "你没有输入"& KS.C_S(ChannelID,3) & "名称!<br/>"
					Prev=True
				    Exit Sub
				  End IF
				  
				  
				UserDefineFieldArr=KSUser.KS_D_F_Arr(ChannelID)
				If IsArray(UserDefineFieldArr) Then
				For I=0 To Ubound(UserDefineFieldArr,2)
				If UserDefineFieldArr(6,I)=1 And KS.G(UserDefineFieldArr(0,I))="" Then 
				 Response.Write UserDefineFieldArr(1,I) & "必须填写!<br/>"
				 Prev=True:Exit SUb
				End If
							
				 If UserDefineFieldArr(6,I)=1 And KS.S(UserDefineFieldArr(0,I))="" Then 
				    Response.Write UserDefineFieldArr(1,I) & "必须填写!<br/>"
					Prev=True:Exit Sub
				 End If
				 If UserDefineFieldArr(3,I)=4 And Not Isnumeric(KS.S(UserDefineFieldArr(0,I))) Then 
				  Response.Write UserDefineFieldArr(1,I) & "必须填写数字<br/>"
				  Prev=True:Exit Sub
				 End If
				 If UserDefineFieldArr(3,I)=5 And Not IsDate(KS.S(UserDefineFieldArr(0,I))) Then 
				  Response.Write UserDefineFieldArr(1,I) & "必须填写正确的日期!<br/>"
				  Prev=True:Exit Sub
				 End If
					If UserDefineFieldArr(3,I)=8 And Not KS.IsValidEmail(KS.S(UserDefineFieldArr(0,I))) and UserDefineFieldArr(6,I)=1 Then 
					 Response.Write UserDefineFieldArr(1,I) & "必须填写正确的Email!<br/>"
					 Prev=True:Exit Sub
					End If
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
				     RSObj("ProID")=KS.GetInfoID(ChannelID)   '取唯一ID
				     RSObj("Hits")=0
					 RSObj("Rolls")=0
					 RSObj("Recommend")=0
					 RSObj("Popular")=0
					 RSObj("Slide")=0
					 RSObj("IsSpecial")=0
					 RSObj("ISTop")=0
					 RSObj("Fname") = Fname
					 RSObj("AddDate")=Now
					 RSObj("Rank")="★★★"
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
			     Call KSUser.AddLog(KSUser.UserName,"发布了" & KS.C_S(ChannelID,3) & """<a href='../item/Show.asp?m=" & ChannelID & "&d=" & InfoID & "' target='_blank'>" & Title & "</a>""!",5)
				 KS.Echo "<br/>"
			     KS.Echo "" & KS.C_S(ChannelID,3) & "添加成功，继续添加吗? "
			     KS.Echo "<a href=""User_myshop.asp?ChannelID=" & ChannelID & "&amp;Action=Add&amp;ClassID=" & ClassID &"&amp;"&KS.WapValue&""">确定</a> "
				 
			    ' KS.Echo "" & KS.C_S(ChannelID,3) & "添加成功，给商品添加图片吗?<br/> "
			   '  KS.Echo "<a href=""User_UpFile.asp?ChannelID=" & ChannelID & "&amp;id=" & InfoID & "&amp;Action=Add&amp;"&KS.WapValue&""">确定</a> "
				 
			     KS.Echo "<a href=""User_Myshop.asp?ChannelID=" & ChannelID & "&amp;"&KS.WapValue&""">取消</a>"
			     KS.Echo "<br/>"
			  Else
			     Call LFCls.ModifyItemInfo(ChannelID,InfoID,Title,classid,Content,KeyWords,PhotoUrl,Verific)
				 Call KS.FileAssociation(ChannelID,InfoID,PhotoUrl & BigPhoto & Content ,1)
			     Call KSUser.AddLog(KSUser.UserName,"对" & KS.C_S(ChannelID,3) & """<a href='../item/Show.asp?m=" & ChannelID & "&d=" & InfoID & "' target='_blank'>" & Title & "</a>""做了修改!",5)
				KS.Echo "<br/>" & KS.C_S(ChannelID,3) & "修改成功!<br/>"
			  End If
		
  End Sub
End Class
%> 
