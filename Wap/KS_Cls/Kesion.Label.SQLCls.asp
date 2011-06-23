<%
Class DIYCls
		Private KS
		Private TConn,DataSourceType,DataSourceStr
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		   Set KS=Nothing
		   If Isobject(TConn) Then
		      TConn.Close:Set TConn=Nothing
		   End If
		End Sub
		
		'替换自定义函数标签
		Function ReplaceUserFunctionLabel(Content)
			Dim regEx, Matches, SqlLabel,Match
			Dim Matchn,n
			Set regEx = New RegExp
			regEx.Pattern = "{SQL_[^{]*\)}"
			regEx.IgnoreCase = True
			regEx.Global = True
			Set Matches = regEx.Execute(Content)
			Dim Str:Str=Content
			For Each Match In Matches
			  SqlLabel=Match.value
			  Str=Replace(Str,SqlLabel,ReplaceDIYFunctionLabel(SqlLabel,"label"))
			Next
			'判断嵌套,Instr(Str,",'{SQL_")=0当含有ajax输出时，不递归
			If Instr(Str,"{SQL_")<>0 And Instr(Str,",'{SQL_")=0 Then Str=ReplaceUserFunctionLabel(Str) 
			ReplaceUserFunctionLabel=Str
		End Function
		
		'缓存数据库sql标签
		Function G_S_P(LabelName,FieldID)
		    On Error Resume Next
			If Not IsObject(Application(KS.SiteSN&"_sqllabellist")) Then
			   Application.Lock
			   Dim RS:Set Rs=Server.CreateObject("ADODB.RECORDSET")
			   RS.Open "select LabelName,Description,LabelContent From KS_Label Where LabelType=5 Order by adddate",Conn,1,1
			   Set Application(KS.SiteSN&"_sqllabellist")=KS.RecordsetToxml(rs,"sqllabel","sqllabellist")
			   RS.Close:Set Rs=Nothing
			   Dim RCls:set RCls=new Refresh
			   Dim objNode,i,j,objAtr
			   Set objNode=Application(KS.SiteSN&"_sqllabellist").documentElement 
			   For I=0 To objNode.ChildNodes.length-1 
			       set objAtr=objNode.ChildNodes.item(I) 
				   objAtr.Attributes.item(2).Text=RCls.ReplaceGeneralLabelContent(objAtr.Attributes.item(2).Text)
			   Next
			   set RCls=nothing			 
			   Application.UnLock
			End If
			G_S_P=Application(KS.SiteSN&"_sqllabellist").documentElement.selectSingleNode("sqllabel[@ks0='" & LabelName & "']/@ks" & FieldID & "").text
			If Err Then G_S_P="":Err.Clear
		End Function
		
		'返回循环次数
		Function GetLoopNum(Content)
			Dim regEx, Matches, Match
			Set regEx = New RegExp
			regEx.Pattern="\[loop=\d*]"
			regEx.IgnoreCase = True
			regEx.Global = True
			Set Matches = regEx.Execute(Content)
			If Matches.Count > 0 Then
			   GetLoopNum=Replace(Replace(Matches.item(0),"[loop=",""),"]","")
			Else
			   GetLoopNum=0
			End If
		End Function

	    '替换ACC随机列表函数
		Function ReplaceAccRnd(Content)
		    Dim regEx,Matches,Match,TempStr
		    Randomize()
			Set regEx = New RegExp
			regEx.Pattern = "{Rnd[^{]*\)}"
			regEx.IgnoreCase = True
			regEx.Global = True
			Set Matches = regEx.Execute(Content)
			For Each Match In Matches
			    TempStr=Match.value
				Content=Replace(Content,TempStr,"Rnd(" & -1*(Int(1000*Rnd)+1) & "*" & Replace(Split(TempStr,"(")(1),")}","") & ")")
			Next
			ReplaceAccRnd=Content
		End Function

		'替换Request的值,支持ReqNum和ReqStr两个标签
		Function ReplaceRequest(Content)
		    Dim regEx, Matches, Match,TempStr,QStr,ReqType
			Set regEx = New RegExp
			regEx.Pattern= "{(ReqNum|ReqStr)[^{}]*}"
			regEx.IgnoreCase = True
			regEx.Global = True
			Set Matches = regEx.Execute(Content)
			For Each Match In Matches
				On Error Resume Next
				TempStr = Match.Value
				ReqType=Split(TempStr,"(")(0)
				QStr=Replace(Split(TempStr,"(")(1),")}","")
				If ReqType="{ReqNum" Then
				   Content=Replace(Content,TempStr,KS.ChkClng(KS.S(QStr)))
				Else
				   Content=Replace(Content,TempStr,KS.S(QStr))
				End If
			Next
			ReplaceRequest=Content
		End Function
		
		'条件替换	
		Function ReplaceCondition(byval str)
		    Dim regEx, Matches, Match, TempStr,Bool
			Dim FieldParam,FieldParamArr,ReturnFieldValue,I
			On Error Resume Next 
			Set regEx = New RegExp
			regEx.Pattern = "{\$IF\([^{\$}]*}"
			regEx.IgnoreCase = True
			regEx.Global = True
			Set Matches = regEx.Execute(str)
			TempStr=str
			For Each Match In Matches
			    FieldParam    = Replace(Replace(Match.Value,"{$IF(",""),")}","")
				FieldParamArr = Split(FieldParam,"||")
				Bool=Eval(Trim(FieldParamArr(0)))
				If Bool="True" Then
				   ReturnFieldValue=FieldParamArr(1)
				Else
				   ReturnFieldValue=FieldParamArr(2)
				End If
				TempStr=Replace(TempStr,"{$IF(" &FieldParam &")}",ReturnFieldValue)
			Next
			ReplaceCondition=TempStr 
		End Function
		
		'替换自定义函数标签 
		'参数SqlLabel:{SQL_标签名称(15,0,1,...)}
		Function ReplaceDIYFunctionLabel(SqlLabel,GetFrom)
		    Dim I,KS_RS_Obj,LabelName,UserParamArr,FunctionLabelParamArr,CirLabelContent,FunctionSQL,LabelContent
			Dim FunctionLabelType,ItemName,PageStyle,PerPageNumber,TotalPut,PageNum,J,TempStr,Ajax
			LabelName    = Replace(Replace(Split(SqlLabel,"(")(0),"""",""),"'","")
			'用户函数参数
			UserParamArr = Split(Replace(Replace(Replace(Replace(SqlLabel,LabelName&"(",""),")}",""),"""",""),"'",""),",")   
			Dim L_Description:L_Description=G_S_P(LabelName &"}",1)
			If L_Description="" Then
		       ReplaceDIYFunctionLabel="":Exit Function
			Else
		       FunctionLabelParamArr = Split(L_Description,"@@@")
		       LabelContent          = Replace(G_S_P(LabelName &"}",2),Chr(10) ,"$KS:Page$")
			End If
			
			FunctionSQL=FunctionLabelParamArr(0)           '查询语句
			FunctionSQL=Replace(FunctionSQL,"{$CurrClassID}",FCls.RefreshFolderID)
			FunctionSQL=Replace(FunctionSQL,"{$CurrChannelID}",FCls.ChannelID)
			If Instr(FunctionSQL,"{$CurrClassChildID}")<>0 Then
			   FunctionSQL=Replace(FunctionSQL,"{$CurrClassChildID}",KS.GetFolderTid(FCls.RefreshFolderID))
			End If
			FunctionSQL=Replace(FunctionSQL,"{$CurrInfoID}",FCls.RefreshInfoID)
			FunctionSQL=Replace(FunctionSQL,"{$CurrSpecialID}",FCls.CurrSpecialID)
			For I=0 To Ubound(UserParamArr)
		        FunctionSQL  = Replace(FunctionSQL,"{$Param("&I&")}",UserParamArr(I))
				LabelContent = Replace(LabelContent,"{$Param("&I&")}",UserParamArr(I))
		    Next
			LabelContent = ReplaceRequest(LabelContent)    '替换Request的值
			FunctionSQL = ReplaceRequest(FunctionSQL)    '替换Request的值
			
			FunctionLabelType=FunctionLabelParamArr(2)
			If Not Isnumeric(FunctionLabelType) Then FunctionLabelType=0
			'-----------------------------------
			'-----------------------------------
			ItemName=FunctionLabelParamArr(3)'分页项目单位
			PageStyle=FunctionLabelParamArr(4)'分页样式
			DataSourceType=FunctionLabelParamArr(6)'数据源
			DataSourceStr=FunctionLabelParamArr(7)'连接字符串
			If DataSourceType=1 Or DataSourceType=5 Or DataSourceType=6 Then DataSourceStr=LFCls.GetAbsolutePath(DataSourceStr)
			'-----------------------------------
			'-----------------------------------
			If OpenExtConn=False Then ReplaceDIYFunctionLabel="外部数据库连接出错!":Exit Function
			Set KS_RS_Obj=Server.CreateObject("ADODB.RECORDSET")
			If DataSourceType=0 Then
			   KS_RS_Obj.Open FunctionSQL,Conn,1,1
			Else
		       KS_RS_Obj.Open FunctionSQL,TConn,1,1
			End IF
			If Not KS_RS_Obj.EOF Then
			   Dim regEx, Matches, Match,LoopTimes
			   Set regEx = New RegExp
			   regEx.Pattern = "\[loop=\d*].+?\[/loop]"
			   regEx.IgnoreCase = True
			   regEx.Global = True
			   Set Matches = regEx.Execute(LabelContent)
			   
			   If FunctionLabelType=1 And DataSourceType=0 Then            '分页标签
			      '分页开始---------------------------------------------------------------------
			      PerPageNumber=0
				  For Each Match In Matches
				      PerPageNumber=PerPageNumber+GetLoopNum(Match.Value)   '每页记录数
				  Next
				  If PerPageNumber=0 Then ReplaceDIYFunctionLabel="自定义函数标签的循环次数必须大于0":Exit Function
				  TotalPut = KS_RS_Obj.Recordcount
				  If (TotalPut Mod PerPageNumber)=0 Then
				     PageNum = TotalPut \ PerPageNumber
				  Else
				     PageNum = TotalPut \ PerPageNumber + 1
				  End If
				  FCls.PageStyle = PageStyle
				  
				  Dim CurrPage:CurrPage=KS.ChkClng(KS.G("Page"))
				  If CurrPage<=0 Then CurrPage=1
				  FCls.TotalPage=PageNum
				  FCls.TotalPut=TotalPut
				  Dim TempCirContent
				  TempCirContent = LabelContent
				  KS_RS_Obj.Move (CurrPage - 1) * PerPageNumber
				  For Each Match In Matches
				      LoopTimes=GetLoopNum(Match.Value)   '循环次数
					  CirLabelContent = Replace(Replace(Match.value,"[loop=" & LoopTimes&"]",""),"[/loop]","")
					  TempCirContent  = Replace(TempCirContent,"[loop="&LoopTimes&"]"&CirLabelContent&"[/loop]",GetCirLabelContent(CirLabelContent,KS_RS_Obj,LoopTimes),1,1)
					  If KS_RS_Obj.Eof Then Exit For
				  Next
				  TempStr = TempCirContent & KS.GetPrePageList(PageStyle,ItemName,PageNum,CurrPage,TotalPut,PerPageNumber)'显示分页的前部分
				  TempStr = TempStr & KS.GetPageList(Replace(KS.GetUrl,"&amp;page="&KS.S("page")&"",""),PageStyle,CurrPage,PageNum,True) '加上分页符
				  ReplaceDIYFunctionLabel=CleanLabel(TempStr)
				  '分页结束--------------------------------------------------------------------- 
			   Else
			      '列表开始---------------------------------------------------------------------
			      Do While Not KS_RS_Obj.EOF
				     For Each Match In Matches
					     LoopTimes=GetLoopNum(Match.Value)   '循环次数
						 CirLabelContent = Replace(Replace(Match.value,"[loop=" & LoopTimes&"]",""),"[/loop]","")
						 LabelContent    = Replace(LabelContent,"[loop="&LoopTimes&"]"&CirLabelContent&"[/loop]",GetCirLabelContent(CirLabelContent,KS_RS_Obj,LoopTimes),1,1)
						 If KS_RS_Obj.EOF Then Exit For
					 Next
					 If KS_RS_Obj.EOF Then
					    Exit Do
					 Else
					    KS_RS_Obj.MoveNext
					 End If
				  Loop
				  '消除多余的循环体
				  ReplaceDIYFunctionLabel=CleanLabel(LabelContent)
				  '列表结束---------------------------------------------------------------------
			   End If
			Else
			   ReplaceDIYFunctionLabel="":Exit Function
			End if
			KS_RS_Obj.Close:Set KS_RS_Obj=Nothing
		End Function

		'消除多余的循环体
		Function CleanLabel(Content)
		    Dim regEx, Matches, Match,LoopTimes
			Set regEx = New RegExp
			regEx.Pattern = "\[loop=\d*][^\[\]]*\[/loop]"
			regEx.IgnoreCase = True
			regEx.Global = True
			Set Matches = regEx.Execute(Content)
			For Each Match In Matches
			    Content=Replace(Content,Match.value,"")
			Next
			CleanLabel=ReplaceCondition(Replace(Content,"$KS:Page$",vbcrlf))
		End Function
		
		'替换循环部分内容
		Function GetCirLabelContent(CirLabelContent,ByRef KS_RS_Obj,LoopTimes)
		    Dim regEx, Matches, Match, TempStr
			Dim FieldParam,FieldParamArr,FieldName,FieldType,ReturnFieldValue
			Dim DB_FieldValue,I,N
			If Not IsNumeric(LoopTimes) Then LoopTimes=10
			If LoopTimes=0 Then LoopTimes=KS_RS_Obj.RecordCount
			For N=1 To LoopTimes
			    If Not KS_RS_Obj.Eof Then
				   Set regEx = New RegExp
				   regEx.Pattern = "{\$Field\([^{\$}]*}"
				   regEx.IgnoreCase = True
				   regEx.Global = True
				   Set Matches = regEx.Execute(CirLabelContent)
				   TempStr=Replace(CirLabelContent,"{$AutoID}",N)
				   For Each Match In Matches
					   FieldParam    = Replace(Replace(Match.Value,"{$Field(",""),")}","")
					   FieldParamArr = Split(FieldParam,",")
					   FieldName     = FieldParamArr(0)       '根据参数得到字段名称
					   FieldType     = FieldParamArr(1)       '根据参数得到字段类型
					   DB_FieldValue=KS_RS_Obj(FieldName)     '得到字段的值
					   If Lcase(FieldName)="keywords" Then
					      ReturnFieldValue=ReplaceKeyTags(1,DB_FieldValue)
					   Else
					      Select Case Lcase(FieldType)
						      Case "text"'取文本字段的值
							  ReturnFieldValue=KS.HTMLCode(Get_Text_Field(DB_FieldValue,FieldParamArr(2),FieldParamArr(3),FieldParamArr(4),FieldParamArr(5)))
							  Case "num"'取数字字段的值
							  ReturnFieldValue=Get_Num_Field(DB_FieldValue,FieldParamArr(2),FieldParamArr(3))
							  Case "date"'取日期字段的值
							  ReturnFieldValue=Get_Date_Field(DB_FieldValue,FieldParamArr(2))
							  'Case "getinfourl"'取对象的链接URL
							  'ReturnFieldValue=Get_InfoUrl_Field(FieldName,DB_FieldValue,FieldParamArr(2),FieldParamArr(3))
							  Case "getclassurl"'得到栏目的链接URL
							  ReturnFieldValue=Get_ClassUrl_Field(FieldName,DB_FieldValue,FieldParamArr(2),FieldParamArr(3))
						  End Select
					   End If
					   On Error Resume Next
					   TempStr=Replace(TempStr,"{$Field(" &FieldParam &")}",ReturnFieldValue)
				   Next
				   GetCirLabelContent=GetCirLabelContent &TempStr
				Else
				   Exit For
				End If
				KS_RS_Obj.MoveNext
			Next
		End Function
		
		'取文本字段的值
		'参数说明：字段值,截段字数,未尾输出的字符,HTML处理方式
		Function Get_Text_Field(FieldValue,CutNum,EndTag,HtmlTag,DefaultChar)
		    Dim TempStr:TempStr=FieldValue
			If FieldValue="" Or IsNull(FieldValue) Then TempStr=DefaultChar
			If Not IsNumeric(HtmlTag) Or Not IsNumeric(CutNum) Then Exit Function
			If HtmlTag=0 Then
			   TempStr=KS.HtmlCode(TempStr)
			ElseIf HtmlTag=1 Then
			   TempStr=TempStr
			ElseIF HtmlTag=2 Then
			   TempStr=Replace(KS.LoseHtml(KS.HtmlCode(TempStr))," ","")
			End If
			If EndTag="0" Then EndTag=""
			If KS.strLength(TempStr)>cint(CutNum) And CutNum<>0 Then TempStr = KS.GotTopic(TempStr, CutNum) & EndTag
			Get_Text_Field=TempStr
		End Function
		
		'取数字字段的值
		'参数说明：FieldValue-字段值,OutType-输出方式0、原数，1、小数，2百分数,XSWS-小数位数
		Function Get_Num_Field(FieldValue,OutType,XSWS)
		    If Not IsNumeric(FieldValue) Then Get_Num_Field=FieldValue:Exit Function
			If Not IsNumeric(OutType) Then OutType=0
			If Not IsNumeric(XSWS) Then XSWS=0
			If OutType=1 Then
			   Get_Num_Field=FormatNumber(FieldValue,XSWS)
			ElseIf OutType=2 Then
			   Get_Num_Field=FormatPercent(FieldValue)
			Else
			   Get_Num_Field=FieldValue
			End if  
		End Function
		
		'取日期字段的值
		'参数说明：FieldValue-字段值,DateMB-输出日期模板
		Function Get_Date_Field(FieldValue,DateMB)
		    IF Not IsDate(FieldValue) Then Get_Date_Field=FieldValue:Exit Function
		    Get_Date_Field=Replace(DateMB,"YYYY",Year(FieldValue))
		    Get_Date_Field=Replace(Get_Date_Field,"YY",Right("0" & Year(FieldValue), 2))
		    Get_Date_Field=Replace(Get_Date_Field,"MM",Right("0" & Month(FieldValue), 2))
		    Get_Date_Field=Replace(Get_Date_Field,"DD",Right("0" & Day(FieldValue), 2))
		    Get_Date_Field=Replace(Get_Date_Field,"hh",Right("0" & hour(FieldValue), 2))
		    Get_Date_Field=Replace(Get_Date_Field,"mm",Right("0" & minute(FieldValue), 2))
		    Get_Date_Field=Replace(Get_Date_Field,"ss",Right("0" & second(FieldValue), 2))
		End Function

		'得到栏目的链接URL
		'参数说明：FieldName-字段名称,FieldValue-字段值，ChannelID数据表 1、2、3、4、100等,OutType输出方式  0、混合，1、URL，2、名称
		Function Get_ClassUrl_Field(FieldName,FieldValue,ChannelID,OutType)
		    If OutType=2 Or DataSourceType<>0 Then Get_ClassUrl_Field=FieldValue:Exit Function
			Dim ClassID:ClassID=FieldValue
			If FieldName="id" Then
			   ClassID  = LFCls.GetSingleFieldValue("Select Tid From " & C_S(ChannelID,2) & " Where " & FieldName &"=" &FieldValue)
			End If
			If OutType=0 Then
			   Get_ClassUrl_Field="<a href=""" & KS.GetDomain & "List.asp?ID=" & KS.C_C(ClassID,9) &"&amp;ChannelID="&ChannelID&"&amp;" & KS.WapValue & """>" & KS.C_C(ClassID,1) &"</a>"
			ElseIf OutType=1 Then
			   Get_ClassUrl_Field="" & KS.GetDomain & "List.asp?ID=" & KS.C_C(ClassID,9) &"&amp;ChannelID="&ChannelID&"&amp;" & KS.WapValue & ""
			End If
		End Function

		Function ReplaceKeyTags(ChannelID,KeyStr)
		    Dim I,K_Arr:K_Arr=Split(KeyStr,"|")
		    For I=0 To Ubound(K_Arr)
		        ReplaceKeyTags=ReplaceKeyTags & "<a href=""" & KS.GetDomain & "Plus/Tags.asp?n=" & K_Arr(i) & "&amp;"&WapValue&""">" & K_Arr(i) & "</a> "
		    Next
		End Function
		
		Function OpenExtConn()
		    If DataSourceType=0 Then
			   OpenExtConn=True
			Else
			   On Error Resume Next
			   Set TConn = Server.CreateObject("ADODB.Connection")
			   TConn.Open DataSourceStr
			   If Err Then 
			      Err.Clear
			      Set TConn = Nothing
			      OpenExtConn=False
			   Else 
			      OpenExtConn=true
			   End If
		    End If
    	End Function

End Class
%> 
