<%
'科汛网站管理系统,会员系统函数类
Dim KSUser
Set KSUser = New UserCls
Class UserCls
			Private KS,I  
			'---------定义会员全局变量开始---------------
			Public ID,GroupID,UserName,PassWord,Question,Answer,Email
			Public RealName,Sex,Birthday,IDCard,OfficeTel,HomeTel,Mobile,Fax
			Public Province,City,Address,Zip,HomePage,QQ,ICQ,MSN,UC,UserFace,FaceWidth,FaceHeight,Sign,Privacy,CheckNum,RegDate
			Public JoinDate,LastLoginTime,LastLoginIP,LoginTimes,Money,Score,Point,locked,RndPassword,UserType
			Public ChargeType,Edays,BeginDate
			'---------定义会员全局变量结束---------------
			
			Private Sub Class_Initialize()
			    Set KS=New PublicCls
			End Sub
			Private Sub Class_Terminate()
			    Set KS=Nothing
				'Set KSUser=Nothing
			End Sub
		   '**************************************************
			'函数名：UserLoginChecked
			'作  用：判断用户是否登录
			'返回值：true或false
			'**************************************************
			Public Function UserLoginChecked()
                On Error Resume Next
				IF KS.G(KS.WSetting(2))="" Then
				   UserLoginChecked=False
				   Exit Function
				Else
				   Dim UserRs:Set UserRS=Server.CreateOBject("ADODB.RECORDSET")
				   UserRS.Open "Select top 1 * From KS_User Where WAP='" & KS.G(KS.WSetting(2)) & "'",Conn,2,3
				   IF UserRS.Eof And UserRS.Bof Then
					  UserLoginChecked=False
				   Else
					  UserLoginChecked=True
					  For I=0 To UserRS.fields.Count-1
						  If lcase(UserRS.Fields(i).Name)="sign" Then
						     Sign=UserRS.Fields(i).Value
						  Else
						     Execute(UserRS.Fields(i).Name&"=ForValue("""&trim(UserRS.Fields(i).Value)&""")")
						  End If
					  Next
				   End if
				   UserRS.Close:Set UserRS=Nothing
			    End IF
			End Function
			
			Public Property Get GetEdays()
				GetEdays = Edays-DateDiff("D",BeginDate,now())
			End Property

			
		    Public Function ForValue(s)
			    If trim(s)="" Then ForValue=""
				ForValue=s
			End Function
			
			'用户上传目录
			Function GetUserFolder(UserName)
			    Dim Ce:Set Ce=new CtoeCls
				UserName=Ce.CTOE(KS.R(UserName))
				Set Ce=Nothing
				GetUserFolder=KS.Setting(3)&KS.Setting(91)&"User/" & UserName & "/"
			End Function
			'返回专栏选择框
			Function UserClassOption(TypeID,Sel)
		        Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
				RS.Open "Select ClassID,ClassName From KS_UserClass Where UserName='" & UserName & "' And TypeID="&TypeID,Conn,1,1
				Do While Not RS.Eof
				   UserClassOption=UserClassOption & "<option value=""" & RS(0) & """>" & RS(1) & "</option>"
				   RS.MoveNext
				Loop
			    RS.Close:Set RS=Nothing
		    End Function
			
			'返回相应模型的自定义字段名称数组(仅限会员中心调用)
		   Function KS_D_F_Arr(ChannelID)
		       Dim KS_RS_Obj:Set KS_RS_Obj=Server.CreateObject("ADODB.RECORDSET")
			   KS_RS_Obj.Open "Select FieldName,Title,Tips,FieldType,DefaultValue,Options,MustFillTF,Width,FieldID From KS_Field Where ChannelID=" & ChannelID &" And ShowOnForm=1 And ShowOnUserForm=1 Order By OrderID Asc",Conn,1,1
			   If Not KS_RS_Obj.Eof Then
			      KS_D_F_Arr=KS_RS_Obj.GetRows(-1)
			   Else
			      KS_D_F_Arr=""
			   End If
			   KS_RS_Obj.Close:Set KS_RS_Obj=Nothing
		   End Function

		   '取得会员中心信息添加时的自定义字段
		   Function KS_D_F(ChannelID,ByVal UserDefineFieldValueStr)
			   Dim PrevUrl
		       Dim I,K,F_Arr,O_Arr,F_Value
			   Dim O_Text,O_Value,BRStr,O_Len,F_V
			   F_Arr=KS_D_F_Arr(ChannelID)
			   If UserDefineFieldValueStr<>"0" And UserDefineFieldValueStr<>""  Then UserDefineFieldValueStr=Split(UserDefineFieldValueStr,"||||")
			   If IsArray(F_Arr) Then
				  For I=0 To Ubound(F_Arr,2)
				      KS_D_F=KS_D_F & F_Arr(1,I) & "："
					  If F_Arr(6,I)=1 Then KS_D_F=KS_D_F & "*"
					  If F_Arr(2,I)<>"" Then KS_D_F=KS_D_F & F_Arr(2,I) &"<br/>"
					  
					  If IsArray(UserDefineFieldValueStr) Then
					     F_Value=UserDefineFieldValueStr(I)
					  Else
					     F_Value=F_Arr(4,I)
					  End If
					  Select Case F_Arr(3,I)
					      Case 2
						     KS_D_F=KS_D_F & "<input type=""text"" name=""" & F_Arr(0,i) & ""&Minute(Now)&Second(Now)&""" value=""" & F_Value & """/>"
						  Case 3
						     KS_D_F=KS_D_F & "<select name=""" & F_Arr(0,I) & ""&Minute(Now)&Second(Now)&""">"
							 O_Arr=Split(F_Arr(5,I),vbcrlf): O_Len=Ubound(O_Arr)
							 For K=0 To O_Len
							     F_V=Split(O_Arr(K),"|")
								 If O_Arr(K)<>"" Then
								    If Ubound(F_V)=1 Then
									   O_Value=F_V(0):O_Text=F_V(1)
									Else
									   O_Value=F_V(0):O_Text=F_V(0)
									End If						   
									KS_D_F=KS_D_F & "<option value=""" & O_Value& """>" &O_Text & "</option>"
								 End If
							 Next
							 KS_D_F=KS_D_F & "</select>"
						  Case 6
						     KS_D_F=KS_D_F & "<select name=""" & F_Arr(0,I) & ""&Minute(Now)&Second(Now)&""">"
						     O_Arr=Split(F_Arr(5,I),vbcrlf): O_Len=Ubound(O_Arr)
							 For K=0 To O_Len
							     F_V=Split(O_Arr(K),"|")
								 If O_Arr(K)<>"" Then
								    If Ubound(F_V)=1 Then
									   O_Value=F_V(0):O_Text=F_V(1)
									Else
									   O_Value=F_V(0):O_Text=F_V(0)
									End If
									KS_D_F=KS_D_F & "<option value=""" & O_Value& """>" &O_Text & "</option>"	   
								 End If
							 Next
							 KS_D_F=KS_D_F & "</select>"
						  Case 7
						     KS_D_F=KS_D_F & "<select name=""" & F_Arr(0,I) & ""&Minute(Now)&Second(Now)&""">"
						     O_Arr=Split(F_Arr(5,I),vbcrlf): O_Len=Ubound(O_Arr)
							 For K=0 To O_Len
							     F_V=Split(O_Arr(K),"|")
								 If O_Arr(K)<>"" Then
								    If Ubound(F_V)=1 Then
									   O_Value=F_V(0):O_Text=F_V(1)
									Else
									   O_Value=F_V(0):O_Text=F_V(0)
									End If						   
									KS_D_F=KS_D_F & "<option value=""" & O_Value& """>" &O_Text & "</option>"
								 End If
							 Next
							 KS_D_F=KS_D_F & "</select>"
					      Case 10
						  KS_D_F=KS_D_F & "<input type=""text"" name=""" & F_Arr(0,i) & ""&Minute(Now)&Second(Now)&""" value=""" & Server.HTMLEncode(F_Value) & """/>"
						  Case Else
						  KS_D_F=KS_D_F & "<input type=""text"" name=""" & F_Arr(0,i) & ""&Minute(Now)&Second(Now)&""" value=""" & F_Value & """/>"
					  End Select
					  If F_Arr(3,I)=9 Then
					     PrevUrl=Request.ServerVariables("SCRIPT_NAME")&"?Action="&Request.QueryString("Action")&""
					     KS_D_F=KS_D_F & " <a href=""User_UpFile.asp?Type=Field&amp;FieldID=" & F_Arr(8,I) & "&amp;ID="&Request.QueryString("ID")&"&amp;ChannelID=" & ChannelID & "&amp;"&KS.WapValue&"&amp;PrevUrl="&PrevUrl&""">上传文件</a>"
				      End If
					  KS_D_F=KS_D_F & "<br/>"
				  Next
			   End If
		   End Function


		   '返回格式化后的时间
		   Function GetTimeFormat(DateTime)
		       If DateDiff("n",DateTime,now)<5 Then
			      GetTimeFormat="刚刚"
			   ElseIf DateDiff("n",DateTime,now)<60 then
			      GetTimeFormat=DateDiff("n",DateTime,now) & "分钟前"
			   ElseIf DateDiff("h",DateTime,now)<5 Then
			      GetTimeFormat=DateDiff("h",DateTime,now) & "小时前"
			   Else
			      GetTimeFormat=formatdatetime(DateTime,2)
			   End If
		   End Function
		   
           Sub CheckMoney(ChannelID)
		     If cdbl(KS.C_S(ChannelID,18))<0 And cdbl(Money)<cdbl(abs(KS.C_S(ChannelID,18))) Then
		      response.write ("在本频道发布信息最少需要消费资金" & abs(KS.C_S(ChannelID,18)) & "元,您当前可用资金为" & Money & "元,请充值续费!")
			  response.write "</p></card></wml>"
			  response.end
		     End If
		     If cdbl(KS.C_S(ChannelID,19))<0 And cdbl(Point)<cdbl(abs(KS.C_S(ChannelID,19))) Then
		      response.write ("在本频道发布信息最少需要消费" & KS.Setting(45) & abs(KS.C_S(ChannelID,19)) & KS.Setting(46) & ",您当前可用" & KS.Setting(45) & "为" & Point & KS.Setting(46) & ",请充值续费!")
			  response.write "</p></card></wml>"
			  response.end
		     End If
		     If cint(KS.C_S(ChannelID,20))<0 And cint(Score)<abs(KS.C_S(ChannelID,20)) Then
		       response.write ("在本频道发布信息最少需要消费积分" & abs(KS.C_S(ChannelID,20)) & "分,您当前可用积分" & Score & "分,请充值续费!")
			   response.write "</p></card></wml>"
			   response.end
		     End If
		   End Sub	
		   
			'增加好友动态
			'参数 username 用户 note 备注 ico图标 1评论 2添加文章 0通用
			Sub AddLog(username,note,ico)
			  Conn.Execute("Insert Into KS_UserLog([username],[note],[adddate],[ico]) values('" & UserName & "','" & replace(note,"'","""") & "'," & SqlNowString & "," & ico & ")")
			End Sub
'删除模型信息数据
		   Sub DelItemInfo(ChannelID)
		        Dim ID:ID=KS.S("ID")
				ID=KS.FilterIDs(ID)
				If ID="" Then Call KS.Alert("你没有选中要删除的" & KS.C_S(ChannelID,3) & "!",ComeUrl):Response.End
				Dim RS,DelIDS
				Set RS=Server.CreateObject("ADODB.RECORDSET")
				RS.Open "Select id  From " & KS.C_S(ChannelID,2) &" Where Inputer='" & UserName & "' and Verific<>1 And ID In(" & ID & ")",conn,1,3
				Do While Not RS.Eof
				  If DelIds="" Then DelIDs=RS(0)   Else DelIds=DelIds & "," & RS(0)
				  Conn.Execute("Delete From KS_UploadFiles Where ChannelID=" & ChannelID &" and infoid=" & rs(0))
				  RS.Delete
				  RS.MoveNext
				Loop
				RS.Close:Set RS=Nothing
				If DelIds<>"" Then
				 Call AddLog(UserName,"删除发表的" & KS.C_S(ChannelID,3) & "操作!" & KS.C_S(ChannelID,3) & "ID:" & DelIds,KS.C_S(ChannelID,6))
				End If
				Conn.Execute("Delete From KS_ItemInfo Where Inputer='" & UserName & "' and Verific<>1 and InfoID in(" & ID & ") and channelid=" & ChannelID)
				Select Case KS.ChkClng(KS.C_S(ChannelID,6))
				 Case 1 Response.Redirect "User_Myarticle.asp?channelid=" & channelid & "&" & KS.WapValue
				 Case 2 Response.Redirect "User_MyPhoto.asp?channelid=" & channelid & "&" & KS.WapValue
				 Case 3 Response.Redirect "User_MySoftWare.asp?channelid=" & channelid & "&" & KS.WapValue
				 Case 5 Response.Redirect "User_MyShop.asp?channelid=" & channelid & "&" & KS.WapValue
				 Case Else Response.Redirect("Index.asp?" & KS.WapValue)
				End Select
			   
		   End Sub			

		   '根据用户组返回对应模型的可用栏目
		   Sub GetClassByGroupID(ByVal ChannelID,ByVal ClassID,Selbutton)
				Dim SQL,K,Node,ClassStr,Pstr,TJ,SpaceStr,Xml
				KS.LoadClassConfig()
				If ChannelID<>0 Then Pstr="and @ks12=" & channelid & ""
				Set Xml=Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks14=1" & Pstr&"]")
		            KS.Echo "<select name='ClassID' id='ClassID' style='width:250px'>"
					For Each Node In Xml
					  If (Node.SelectSingleNode("@ks18").text=0) OR ((KS.FoundInArr(Node.SelectSingleNode("@ks17").text,GroupID,",")=false and Node.SelectSingleNode("@ks18").text=3) ) Then
					  Else
							SpaceStr=""
							TJ=Node.SelectSingleNode("@ks10").text
							If TJ>1 Then
							 For k = 1 To TJ - 1
								SpaceStr = SpaceStr & "──"
							 Next
							End If
							
							If ClassID=Node.SelectSingleNode("@ks0").text Then
								KS.Echo "<option value='" & Node.SelectSingleNode("@ks0").text & "' selected=""selected"">" & SpaceStr& Node.SelectSingleNode("@ks1").text & "</option>"
							Else
								KS.Echo "<option value='" & Node.SelectSingleNode("@ks0").text & "'>" & SpaceStr & Node.SelectSingleNode("@ks1").text & "</option>"
							End If
					  End If
					Next
					KS.Echo "</select>"
					Exit Sub
		   
			End Sub
End Class
%> 
