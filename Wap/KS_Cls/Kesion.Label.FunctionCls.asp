<%
Class RefreshFunction
        Private KS,DomainStr,WapValue
		Private Sub Class_Initialize()
		    Set KS=New PublicCls
		    DomainStr=KS.GetDomain
			WapValue=KS.WapValue
		End Sub
        Private Sub Class_Terminate()
		    Set KS=Nothing
		End Sub
		
		'============================================================文章发布中心通用函数声明==============================
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'函数名:KS_A_L
		'作 用:通用栏目文章列表
		'参 数:
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Public Function KS_A_L(ChannelID, SqlStr, S_C_N, T_Len, PicTF, NewTF, HotTF)
			Dim RS:Set RS=Conn.Execute(SqlStr)
			If RS.EOF Then	  KS_A_L="":RS.Close:Set RS=Nothing:Exit Function
			Dim SQL:SQL=RS.GetRows(-1):RS.Close:Set RS = Nothing
			Dim TotalNum:TotalNum=Ubound(SQL,2)
			Dim K,C_N_Link,Title,TempTitle,NewImgStr,HotImgStr
		    For K=0 To TotalNum
			    If CBool(S_C_N) = True Then C_N_Link = "[" & KS.GetClassNP(SQL(2,K)) & "]"
			    Title = SQL(1,K)
				TempTitle = GetArticleTitle(Title, T_Len, PicTF, SQL(2,K))
				If Cbool(NewTF)=True And (Year(SQL(5,K))&Month(SQL(5,K))&Day(SQL(5,K)) =Year(Now)&Month(Now)&Day(Now)) Then NewImgStr="<img src=""" & DomainStr & "images/new.gif""/>" Else NewImgStr=""
				If Cbool(HotTF)=True And SQL(7,K)=1 Then HotImgStr="<img src=""/Images/hot.gif""/>" Else HotImgStr=""
				TempTitle = "<a href=""" & KS.GetInfoUrl(ChannelID,SQL(0,K),SQL(8,K)) & """>" & TempTitle & "</a>"
				KS_A_L = KS_A_L & ("" & TempTitle & NewImgStr & HotImgStr & "<br/>" & vbCrLf)
			Next
			KS_A_L = KS_A_L
		End Function
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'函数名: KS_PicA_L
		'作  用: 通用图片文章函数
		'参  数: SqlStr 待查询的SQL语句,OpenTypStr链接打开类型,等
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Public Function KS_PicA_L(ChannelID,SqlStr, Width, Height, ShowTitle, PicStyle, C_Len, T_Len)
		     Dim SQL,K,N,Title,TempPicStr,Url,LinkAndPicStr,TempTitleStr,ArticleContent,ReturnStr
			 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			 RS.Open SqlStr, Conn, 1, 1
			 If RS.EOF Then	  KS_PicA_L="":RS.Close:Set RS=Nothing:Exit Function
			 SQL=RS.GetRows(-1):RS.Close:Set RS = Nothing
			 Dim TotalNum:TotalNum=Ubound(SQL,2)
			 N=0
			 For K=0 To TotalNum
				 Title = SQL(1,N)
				 TempPicStr = SQL(6,N)
				 TempPicStr=GetPicUrl(TempPicStr)
				 Url = KS.GetInfoUrl(ChannelID,SQL(0,N),SQL(8,N))
				 LinkAndPicStr = "<a href=""" & Url & """><img src=""" & TempPicStr & """ width=""" & Width & """ height=""" & Height & """ align=""""/></a>"
				 TempTitleStr = GetArticleTitle(Title, T_Len, False, SQL(2,N))
				 TempTitleStr = "<a href=""" & Url & """>" & TempTitleStr & "</a>"
				 If SQL(3,N)="" Or IsNull(SQL(3,N)) Then ArticleContent=SQL(4,N) Else ArticleContent=SQL(3,N)
				 Select Case CInt(PicStyle)
					 Case 1:ReturnStr = ReturnStr & LinkAndPicStr & "<br/>" & vbCrLf
					 Case 2:ReturnStr = ReturnStr & LinkAndPicStr & "<br/>" & TempTitleStr & "<br/>" & vbCrLf
					 Case 3
					    ReturnStr = ReturnStr & LinkAndPicStr
						If Cbool(ShowTitle) = True Then	ReturnStr = ReturnStr &"<br/>"& TempTitleStr &"<br/>"
						ReturnStr = ReturnStr & KS.GotTopic(Replace(Replace(Replace(KS.LoseHtml(ArticleContent), vbCrLf, ""), "[NextPage]", ""), "&nbsp;", ""), C_Len) &"...[<a href=""" & Url & """>全文</a>]<br/>"& vbCrLf
					  Case 4
						If Cbool(ShowTitle) = True Then	ReturnStr = ReturnStr &  TempTitleStr &"<br/>"
						ReturnStr = ReturnStr &  KS.GotTopic(Replace(Replace(Replace(KS.LoseHtml(ArticleContent), vbCrLf, ""), "[NextPage]", ""), "&nbsp;", ""), C_Len) &"...[<a href=""" & Url & """>全文</a>]<br/>"& vbCrLf
				 End Select
			     N=N+1
             Next
			 KS_PicA_L = ReturnStr
		End Function
		
		
		'通用专题列表
		Public Function KS_C_Special_L(SqlStr,IntroLen,T_Len,NaviStr,ShowStyle, Width, Height)
		    Dim RS,SQL,TotalNum,K
			Dim TempTitle,SpecialUrl,TempPicStr,TempStr
			Set RS=Conn.Execute(SqlStr)
			If RS.Eof And RS.Bof Then KS_C_Special_L="":RS.Close:Set RS=Nothing:Exit Function
			SQL=RS.GetRows(-1):RS.Close:Set RS=Nothing
			TotalNum=Ubound(SQL,2)
			For K=0 To TotalNum
			    TempTitle = KS.GotTopic(SQL(1,K), T_Len)
				SpecialUrl="Special.asp?id=" & SQL(0,K) & "&amp;ChannelID=" & SQL(3,K) & "&amp;"&WapValue&""
				TempTitle = "<a href=""" & SpecialUrl & """>" & TempTitle & "</a>"
				TempPicStr=GetPicUrl(SQL(6,K))
				TempPicStr="<a href=""" & SpecialUrl & """><img src=""" & TempPicStr & """ width=""" & Width & """ height=""" & Height &""" align=""""/></a>"
				Select Case Cint(ShowStyle)
				    Case 1:TempStr = TempStr & NaviStr & TempTitle & "<br/>"& vbCrLf
					Case 2:TempStr = TempStr & TempPicStr & "<br/>"& vbCrLf
					Case 3:TempStr = TempStr & TempPicStr & "<br/>" & TempTitle &"<br/>"& vbCrLf
					Case 4:TempStr = TempStr & TempPicStr & "<br/>" & KS.GotTopic(SQL(7,K),introlen) &"<br/>"& vbCrLf
					Case 5:TempStr = TempStr & TempPicStr & "<br/>" & TempTitle &"<br/>" & KS.GotTopic(KS.LoseHtml(SQL(7,K)),introlen) &"<br/>"& vbCrLf
			    End Select
            Next
			KS_C_Special_L = TempStr
		End Function

		'==========================================================================图片发布中心通用函数声明==============================
		
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'函数名:KS_P_L
		'作 用:通用图片列表
		'参 数:SqlStr 待查询的SQL语句,
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Public Function KS_P_L(ChannelID,SqlStr, Width, Height, PicStyle, T_Len, NaviStr)
		    Dim SQL,K
			Dim Title,TempPicStr,Url,LinkAndPicStr,TempTitleStr
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open SqlStr, Conn, 1, 1
			If RS.EOF Then	 KS_P_L="":RS.Close:Set RS=Nothing:Exit Function
			SQL=RS.GetRows(-1):RS.Close:Set RS = Nothing
			Dim TotalNum:TotalNum=Ubound(SQL,2)
			For K=0 To TotalNum
			    Title = SQL(1,K)
				TempPicStr=GetPicUrl(SQL(7,K))
				Url = KS.GetInfoUrl(ChannelID,SQL(0,K),0)
				LinkAndPicStr = "<a href=""" & Url & """><img src=""" & TempPicStr & """ width=""" & Width & """ height=""" & Height & """ alt=""""/></a>"
				TempTitleStr = "<a href=""" & Url & """>" & KS.GotTopic(Title, T_Len) & "</a>"
					 Select Case CInt(PicStyle)
					  Case 1:KS_P_L = KS_P_L & LinkAndPicStr & "<br/>" & vbCrLf
					  Case 2:KS_P_L = KS_P_L & LinkAndPicStr & "<br/>" & TempTitleStr & "<br/>" & vbCrLf
					  Case 3:KS_P_L = KS_P_L & NaviStr & TempTitleStr & "<br/>" & vbcrlf
				 End Select
			Next
		End Function
		
		
		'==================================下载发布中心通用函数声明==============================
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'函数名:KS_D_L
		'作 用:通用栏目下载列表
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Public Function KS_D_L(ChannelID,SqlStr, T_Len, ShowStyle, NaviStr)
			 Dim SQL,K,N,TotalNum
			 Dim Title,Url,TempTitle
			 Dim RS:Set RS=Conn.Execute(SqlStr)
			 If RS.EOF Then	KS_D_L="":RS.Close:Set RS=Nothing:Exit Function
			 SQL=RS.GetRows(-1):RS.Close:Set RS = Nothing
			 TotalNum=Ubound(SQL,2)
			 For K=0 To TotalNum
			     Title = SQL(1,K) & SQL(3,K)
				 Url = KS.GetInfoUrl(ChannelID,SQL(0,K),0)
				 TempTitle = "<a href=""" & Url & """>" & KS.GotTopic(Title, T_Len) & "</a>"
				 KS_D_L = KS_D_L & NaviStr & TempTitle & "<br/>" & vbCrLf
			 Next
			 KS_D_L = KS_D_L
		End Function

		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'函数名: KS_C_PicD_L
		'作  用: 通用图片下载函数
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Public Function KS_C_PicD_L(ChannelID,SqlStr, Width, Height, ShowTitle, PicStyle, C_Len, T_Len)
		     Dim SQL,TotalNum,K
			 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			 RS.Open SqlStr, Conn, 1, 1
			 If RS.EOF Then	KS_C_PicD_L="":RS.Close:Set RS=Nothing:Exit Function
			 SQL=RS.GetRows(-1):RS.Close:Set RS = Nothing
			 TotalNum=Ubound(SQL,2)
			 Dim Title,TempPicStr,Url,LinkAndPicStr,TempTitleStr,ReturnStr
			 For K=0 To TotalNum
			     Title = SQL(1,K)
				 TempPicStr=GetPicUrl(SQL(4,K))
				 Url = KS.GetInfoUrl(ChannelID,SQL(0,K),0)
				 LinkAndPicStr = "<a href=""" & URL & """><img src=""" & TempPicStr & """ width=""" & Width & """ height=""" & Height & """ align=""""/></a>"
				 TempTitleStr = "<a href=""" & URL & """>" & KS.GotTopic(Title,T_Len) & "</a>"
				 Select Case CInt(PicStyle)
					 Case 1:ReturnStr = ReturnStr & ("" & LinkAndPicStr & "<br/>" & vbCrLf)
					 Case 2:ReturnStr = ReturnStr & ("" & LinkAndPicStr & "<br />" & TempTitleStr & "<br/>" & vbCrLf)
					 Case 3       
					    ReturnStr = ReturnStr & ("" & LinkAndPicStr)
						If Cbool(ShowTitle) = True Then	ReturnStr = ReturnStr & ("<br/>" & TempTitleStr &"<br/>")
						ReturnStr = ReturnStr & ("" & KS.GotTopic(Replace(Replace(KS.LoseHtml(KS.HTMLCode(SQL(5,K))), vbCrLf, ""), "&nbsp;", ""), C_Len) &"...<br/>"& vbCrLf)
					 Case 4
						If Cbool(ShowTitle) = True Then	ReturnStr = ReturnStr & TempTitleStr
						ReturnStr = ReturnStr & ("<br/>" & KS.GotTopic(Replace(Replace(KS.LoseHtml(KS.HTMLCode(SQL(5,K))), vbCrLf, ""), "&nbsp;", ""), C_Len) &"...<br/>" & LinkAndPicStr &"<br/>"& vbCrLf)
				End Select
			Next
			KS_C_PicD_L = ReturnStr
	    End Function
		
		'==========================================================================商城通用函数声明==============================
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'函数名:KS_Pro_L
		'作 用:通用商品列表
		'参 数:
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Public Function KS_Pro_L(SqlStr,ShowStyle,ButtonType,PriceType,Discount,Width,Height,T_Len)
		    'On Error Resume Next
			Dim RS,SQL,TotalNum
			Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open SqlStr, Conn, 1, 1
			If RS.Eof And RS.Bof Then KS_Pro_L="" :RS.Close:Set RS=Nothing:Exit Function
			SQL=RS.GetRows(-1):RS.Close:Set RS=Nothing
			TotalNum=Ubound(SQL,2)
			Dim K,CurrTid,Title,Url,ButtonStr,PriceStr,TempPicStr,LinkAndPicStr,TempTitleStr
			For K=0 To TotalNum
			    CurrTid = SQL(2,K)
				Title = SQL(1,K)
				Url = KS.GetInfoUrl(5,SQL(0,K),0)
				ButtonStr=GetButtonStr(ButtonType,SQL(0,K),Url)
				PriceStr=GetPriceStr(PriceType,Discount,SQL(6,K),SQL(7,K),SQL(8,K),SQL(9,K),SQL(10,K))
				TempPicStr=GetPicUrl(SQL(5,K))
				LinkAndPicStr = "<a href=""" & Url & """><img src=""" & TempPicStr & """ width=""" & Width & """ height=""" & Height & """ align=""""/></a>"
				TempTitleStr = "<a href=""" & Url & """>" & KS.GotTopic(Title, T_Len) & "</a>"
				Select Case CInt(ShowStyle)
					Case 1    
						KS_Pro_L = KS_Pro_L & TempTitleStr &"<br/>"&vbcrlf
					Case 2          
						 KS_Pro_L = KS_Pro_L & LinkAndPicStr & "<br/>" & vbCrLf
				    Case 3        
						 KS_Pro_L = KS_Pro_L & LinkAndPicStr & "<br/>" & TempTitleStr & "<br/>" & vbCrLf
					Case 4        
						 KS_Pro_L = KS_Pro_L & LinkAndPicStr & "<br/>" & ButtonStr & "<br/>" & vbCrLf
                    Case 5      
						 KS_Pro_L = KS_Pro_L & LinkAndPicStr & "<br/>" & TempTitleStr & "<br/>" & ButtonStr & "<br/>" & vbCrLf
					Case 6      
						 KS_Pro_L = KS_Pro_L & LinkAndPicStr & "<br/>" & TempTitleStr & "<br/>" & PriceStr & "<br/>" & ButtonStr & "<br/>" & vbCrLf
                    Case 7
					    KS_Pro_L = KS_Pro_L & LinkAndPicStr & "<br/>" & PriceStr & "<br/>" &ButtonStr  & "<br/>" & vbCrLf
					Case 8
					    KS_Pro_L = KS_Pro_L & LinkAndPicStr & "<br/>" &TempTitleStr & PriceStr & "<br/>" &ButtonStr  & "<br/>" & vbCrLf	
					Case 9
					    KS_Pro_L = KS_Pro_L & LinkAndPicStr & TempTitleStr & "<br/>" & PriceStr & "<br/>" &ButtonStr  & "<br/>" & vbCrLf	 
					Case 10
					    KS_Pro_L = KS_Pro_L & LinkAndPicStr & "<br/>" &TempTitleStr & TempTitleStr & "<br/>" & PriceStr & "<br/>" &ButtonStr  & "<br/>" & vbCrLf	 
				End Select
			Next
			KS_Pro_L = KS_Pro_L
		End Function
		

		'价格样式
		Function GetPriceStr(PriceType,Discount,Discount_v,Price_Original,Price,Price_Market,Price_Member)
		    If Price_Market=0 Then Price_Market="—" Else Price_Market="￥"&Price_Market
			If Price_Member=0 Then Price_Member="—" Else Price_Member="￥"&Price_Member
			If Price_Original=0 Then Price_Original="—" Else Price_Original="￥"&Price_Original
			Select Case PriceType
			  Case 0:GetPriceStr="市场价:"&Price_Market &" 商城价:￥"&Price &" 会员价:" & Price_Member
			  Case 1:GetPriceStr="原价:"&Price_Original
			  Case 2:GetPriceStr="商城价:￥"&Price
			  Case 3:GetPriceStr="原　价:"&Price_Original & " 会员价:" & Price_Member
			  Case 4:GetPriceStr="商城价:￥"&Price & " 会员价:" & Price_Member
			  Case 5:GetPriceStr="市场价:"&Price_Market & " 商城价:￥"&Price
			  Case 6:GetPriceStr="市场价:"&Price_Market &" 原　价:"&Price_Original & " 会员价:"&Price_Member
			  Case 7:GetPriceStr="市场价:"&Price_Market &" 原　价:"&Price_Original & " 商城价:￥"&Price & " 会员价:"&Price_Member
			End Select
			If Cbool(Discount)=True Then GetPriceStr=GetPriceStr & "<br/>折扣率:"&FormatPercent(Discount_v/10,0)
		End Function
		'按钮样式
		Function GetButtonStr(ButtonType,ID,Url)
		    Dim BuyButton,FavButton,XQButton
			BuyButton="<a href=""" & DomainStr & "shop/ShoppingCart.asp?ProductList=" &ID &"&amp;"&WapValue&""">购买</a>"
			FavButton="<a href=""" & DomainStr & "Plus/Favorite.asp?ChannelID=5&ID=" & ID &"&amp;"&WapValue&""">收藏</a>"
			XQButton="<a href="""&Url&""">详细</a>"
			Select Case ButtonType
			    Case 1:GetButtonStr=BuyButton
				Case 2:GetButtonStr=FavButton
				Case 3:GetButtonStr=XQButton
				Case 4:GetButtonStr=BuyButton&" "&FavButton
				Case 5:GetButtonStr=BuyButton&" "&XQButton
				Case 6:GetButtonStr=FavButton&" "&XQButton
				Case 7:GetButtonStr=BuyButton&" "&XQButton&" "&FavButton
				Case Else:GetButtonStr=""
		    End Select
	    End Function
		
		'=====================================================供求通用开始==========================================================
		'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'函数名:KS_S_L
		'作 用:通用供求信息列表
		'参 数:SqlStr--待查询的SQL语句,ShowStyle--样式,Width--宽,Height--高,C_Len--内容字符,T_Len--标题字符
		'      S_C_N--栏目类别,ShowGQType--类别,NewTF--
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Public Function KS_S_L(SqlStr,ShowStyle,Width,Height,C_Len,T_Len,NaviStr,S_C_N,ShowGQType,NewTF)
		    Dim SQL,K,N,TotalNum
			Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open SqlStr, Conn, 1, 1
			If RS.EOF Then	 KS_S_L="":RS.Close:Set RS=Nothing:Exit Function
			SQL=RS.GetRows(-1):RS.Close:Set RS = Nothing
			TotalNum=Ubound(SQL,2)
			Dim URL,LinkAndPicStr,Province,City,GQType,TempPicStr, TempTitleStr, I, CurrTid,Title,C_N_Link,NewImgStr,ProCity
			For K=0 To TotalNum
				CurrTid =SQL(2,K)
				Title = SQL(1,K)
				Url = KS.GetInfoUrl(ChannelID,SQL(0,K),0)
				If Cbool(ShowGQType)=True Then GQType = KS.GetGQTypeName(SQL(7,K)) Else GQType=""
				If CBool(S_C_N) = True Then C_N_Link = "[" & KS.GetClassNP(CurrTid) & "]"		
				If Cbool(NewTF)=True And (Year(SQL(4,K))&Month(SQL(4,K))&Day(SQL(4,K)) =Year(Now)&Month(Now)&Day(Now)) Then
				   NewImgStr="<img src=""" & DomainStr &"images/new.gif"" alt=""""/>"
				Else
				   NewImgStr=""
				End If
				Province = SQL(8,K):City= SQL(9,K)
				IF Not IsNull(Province) And Province<>"" Then ProCity=Province & "/" & City Else ProCity="地区不限"
				TempPicStr=GetPicUrl(SQL(5,K))			
				LinkAndPicStr = "<a href=""" & URL & """><img src=""" & TempPicStr & """ width=""" & Width & """ height=""" & Height & """ alt=""""/></a>"
				TempTitleStr = NaviStr & C_N_Link & "<a href=""" & URL & """>[" & GQType &"]"& KS.GotTopic(Title, T_Len) & "</a>" & NewImgStr
				Select Case CInt(ShowStyle)
				    Case 1:KS_S_L = KS_S_L & TempTitleStr &"<br/>"&vbcrlf
				    Case 2:KS_S_L = KS_S_L & LinkAndPicStr & "<br/>" & TempTitleStr & "<br/>" & vbCrLf
					Case 3:KS_S_L = KS_S_L & LinkAndPicStr & "<br/>" & TempTitleStr & "<br/>" & ProCity & " " & KS.GotTopic(Replace(Replace(KS.LoseHtml(KS.HTMLCode(SQL(6,K))), vbCrLf, ""), "&nbsp;", ""), C_Len) & "……<br/>" & vbCrLf
					Case 4 :KS_S_L = KS_S_L & TempTitleStr & "<br/>" & ProCity & " " & KS.GotTopic(Replace(Replace(KS.LoseHtml(KS.HTMLCode(SQL(6,K))), vbCrLf, ""), "&nbsp;", ""), C_Len) & "...<br/>" &  vbCrLf
				End Select
			Next
			KS_S_L = KS_S_L & vbCrLf
		End Function

		'=====================================================供求通用结束==========================================================
		
		'============================================================================================================================
		'                                                         以下为相关刷新通用函数
		'============================================================================================================================
		
		'----------------------------------------------------------------------------------------------------------------------------
		'函数名: GetArticleTitle
		'功 能:取得文章标题
		'参 数: Title原标题, T_Len保留字符长度,PicTF显示图文标志与否 TitleType图文标志
		'----------------------------------------------------------------------------------------------------------------------------
		Function GetArticleTitle(Byval Title, T_Len, PicTF, TitleType)
			Dim DecoratesTitle
			If IsNumeric(T_Len) Then
			  Title = KS.GotTopic(Title, T_Len)
			End If
			If CBool(PicTF) = True Then
			 Select Case Trim(TitleType)
			   Case "[图文]":DecoratesTitle = "" & TitleType & ""
			   Case "[组图]":DecoratesTitle = "" & TitleType & ""
			   Case "[推荐]":DecoratesTitle = "" & TitleType & ""
			   Case "[注意]":DecoratesTitle = "" & TitleType & ""
			 End Select
		    End If
		    DecoratesTitle = DecoratesTitle & Title
		    GetArticleTitle = DecoratesTitle
		End Function
		Function GetPicUrl(PicUrl)
		    PicUrl=Trim(PicUrl)
			If IsNull(PicUrl) Or Trim(PicUrl) = "" Then PicUrl = DomainStr & "/Images/Nopic.gif"	
			if Lcase(left(PicUrl,7))<>"http://" Then GetPicUrl=KS.Setting(2) &PicUrl else GetPicUrl=PicUrl
		End Function
		

		'顶一下
		Function GetDigg(ChannelID,ID)
		    On Error Resume Next
			Dim Action,DiggNum,cDiggNum
			Action=KS.S("Action")
			If Action="digghits" Then
			   Dim RS
			   Dim LoginTF,UserName,Digg
			   LoginTF=KSUser.UserLoginChecked()
			   Dim DigType:DigType=KS.ChkClng(KS.S("DigType"))
			   If LoginTF=True or KS.C_S(ChannelID,37)="1" Then
			      UserName=KSUser.UserName
				  If UserName="" Then UserName="游客"
				  Set RS=Server.CreateObject("ADODB.RECORDSET")
				  RS.Open "Select * From KS_DiggList Where ChannelID=" & ChannelID & "And InfoID="&ID,Conn,1,3
				  If RS.Eof Then
				     RS.AddNew
				     RS("ChannelID")=ChannelID
				     RS("InfoID")=ID
				     RS("LastDiggTime")=Now()
				     RS("LastDiggUser")=UserName
				     RS("DiggNum")=0
				     RS("CDiggNum")=0
				     RS.Update
				  End IF
				  RS.Close
				  Dim DiggID:DiggID=Conn.Execute("Select DiggID From KS_DiggList Where ChannelID=" & ChannelID & "And InfoID=" & ID)(0)
				  RS.Open "Select * From KS_Digg Where ChannelID=" & ChannelID &" And InfoID=" & ID & " And UserIP='" & KS.GetIP() & "'",Conn,1,3
				  If Not RS.Eof Then
				     If (KS.ChkClng(KS.C_S(ChannelID,39))=0 or (RS("UserIP")=KS.GetIP() And KS.ChkClng(KS.C_S(ChannelID,38))=1)) Then
					    Digg=False
					 Else
					    Digg=True
					 End If
				 Else
				    Digg=True
				 End If
				 If Digg=True Then
				    RS.AddNew
					RS("ChannelID")=ChannelID
					RS("InfoID")=ID
					RS("UserName")=UserName
					RS("UserIP")=KS.GetIP()
					RS("DiggID")=DiggID
					RS("DiggTime")=Now
					RS("DiggType")=DigType
					RS.Update
					If DigType=0 Then
					Conn.Execute("Update KS_DiggList set DiggNum=DiggNum+" & KS.ChkClng(KS.C_S(ChannelID,40)) &" Where ChannelID=" & ChannelID & " And InfoID="& ID) 
					Else
					Conn.Execute("Update KS_DiggList set CDiggNum=CDiggNum+" & KS.ChkClng(KS.C_S(ChannelID,40)) &" Where ChannelID=" & ChannelID & " And InfoID="& ID) 
					End If
				  End If
				  RS.Close:Set RS=Nothing
			   End IF
		    End IF
		    DiggNum=Conn.Execute("Select DiggNum From KS_DiggList Where ChannelID=" & ChannelID & " And InfoID=" & ID)(0)
		    cDiggNum=Conn.Execute("Select CDiggNum From KS_DiggList Where ChannelID=" & ChannelID & " And InfoID=" & ID)(0)
		    If Err Then DiggNum="0"
		    'GetDigg="<a href=""" & DomainStr & "Show.asp?Action=digghits&amp;ID=" & ID & "&amp;ChannelID=" & ChannelID & "&amp;sLen=" & KS.S("sLen") & "&amp;CPage=" & KS.S("CPage") & "&amp;" & WapValue & """>顶一下(" & DiggNum & ")</a>"
		    GetDigg="<anchor><img src=""" & DomainStr & "images/xh.gif"" border=""0""/>顶一下(" & DiggNum & ")<go href=""" & DomainStr & "Show.asp?Action=digghits&amp;ID=" & ID & "&amp;ChannelID=" & ChannelID & "&amp;sLen=" & KS.S("sLen") & "&amp;CPage=" & KS.S("CPage") & "&amp;" & WapValue & """ method=""get""></go></anchor> <anchor><img src=""" & DomainStr & "images/jd.gif"" border=""0""/>踩一下(" & cDiggNum & ")<go href=""" & DomainStr & "Show.asp?Action=digghits&amp;ID=" & ID & "&amp;ChannelID=" & ChannelID & "&amp;digtype=1&amp;sLen=" & KS.S("sLen") & "&amp;CPage=" & KS.S("CPage") & "&amp;" & WapValue & """ method=""get""></go></anchor>"
	    End Function
		
		Function GetShowComment(Num,LenNum,ChannelID,ID)
		    On Error Resume Next
		    Dim RS,I
			Set RS=Server.CreateObject("ADODB.RECORDSET")
			RS.Open "Select top "&Num&" Content From KS_Comment Where Verific=1 And ChannelID="&ChannelID&" And InfoID="&ID&" Order By ID Desc",Conn,1,1
			If Not RS.EOF Then
			   For I=1 To Num
			       GetShowComment = GetShowComment &""&I&"."&KS.ReplaceFace(KS.GotTopic(RS("Content"),LenNum))&"<br/>"
				   RS.MoveNext
			   Next
			   GetShowComment = GetShowComment &"<a href=""" & DomainStr & "Plus/Comment.asp?Action=CommentMain&ChannelID="&ChannelID&"&InfoID="&ID&"&"&WapValue&""">更多评论("&Conn.Execute("Select Count(ID) From KS_Comment Where Verific=1 And ChannelID="&ChannelID&" And InfoID="&ID&"")(0)&"条)</a><br/>"
			End If
		End Function
		'发表评论
		Function GetWriteComment(ChannelID,ID)
		    Dim k,str,strArr,reSayArry
			str="惊讶|撇嘴|色色|发呆|得意|流泪|害羞|闭嘴|睡觉|大哭|尴尬|发怒|调皮|呲牙|微笑|难过|酷|非典|抓狂|吐|"
			strArr=Split(str,"|")
			GetWriteComment = "<select name=""insertface"">"
			GetWriteComment = GetWriteComment &"<option value="""">无</option>"
			For k=0 to 19
			    GetWriteComment = GetWriteComment &"<option value=""[e"&k&"]"">" & strArr(k) & "</option>"
			Next
			GetWriteComment = GetWriteComment &"</select> "
			reSayArry = Array("要顶!","你牛!我顶!","这个不错!该顶!","支持你!","反对你!")
			Randomize
			GetWriteComment = GetWriteComment &"<input name=""C_Content"&Minute(Now)&Second(Now)&""" type=""text"" size=""20"" maxlength="""&KS.C_S(ChannelID,14)&""" value="""&reSayArry(Int(Ubound(reSayArry)*Rnd))&"""/> "
			If KS.C_S(ChannelID,13)="1" Then
			   GetWriteComment = GetWriteComment & "认证码：<input name=""VerifyCode"&Minute(Now)&Second(Now)&""" type=""text"" size=""4"" /><b>" & KS.GetVerifyCode & "</b>"
			End IF
			GetWriteComment = GetWriteComment &" <anchor>发表<go href=""" & DomainStr & "Plus/Comment.asp?Action=WriteSave&ChannelID="&ChannelID&"&InfoID="&ID&"&"&WapValue&""" method=""post"">"
			GetWriteComment = GetWriteComment &"<postfield name=""insertface"" value=""$(insertface)""/>"
			GetWriteComment = GetWriteComment &"<postfield name=""C_Content"" value=""$(C_Content"&Minute(Now)&Second(Now)&")""/>"
			GetWriteComment = GetWriteComment &"<postfield name=""VerifyCode"" value=""$(VerifyCode"&Minute(Now)&Second(Now)&")""/>"
			GetWriteComment = GetWriteComment &"</go></anchor><br/>"
		End Function

		'**************************************************
		'函数名：GetRandomContentsList
		'作  用：显示内容页随机列表
		'参  数：strHead--头导航类型
		'       strTail--尾导航类型
		'       strNum--显示记录数
		'       strTitleNum--链接标题字符
		'**************************************************
		Function GetRandomContentsList(strHead,strTail,strNum,strTitleNum)
			Dim ID:ID=KS.ChkClng(FCls.RefreshInfoID)
			Dim ChannelID:ChannelID=KS.ChkClng(FCls.ChannelID)
			If ChannelID=0 Then Exit Function
			Dim RS,Param,SqlStr,TempStr,I,XML,Node,strReplace
			Param="Select top "&strNum&" ID,Tid,Title,Fname From " & KS.C_S(ChannelID,2) & ""
			If DataBaseType=0 Then
			   Randomize()
			   SqlStr="" & Param & " where Tid='"&Conn.Execute("select Tid from "&KS.C_S(ChannelID,2)&" where ID="&ID&"")(0)&"' And Verific=1 order by Rnd("&-1*(Int(1000*Rnd)+1)&"*ID)"
			Else
			   SqlStr="" & Param & " where Tid='"&Conn.Execute("select Tid from "&KS.C_S(ChannelID,2)&" where ID="&ID&"")(0)&"' And Verific=1 order by newid()"
			End If
			set RS=Conn.Execute(SqlStr)
			If Not RS.Eof Then
			  Set XML=KS.RsToXml(RS,"row","")
			End If
			RS.Close : Set RS=Nothing
			
			If Not IsObject(XML) Then
			   GetRandomContentsList = "暂无相关内容!<br/>"
			   Exit Function
			End IF
			I=0
			For Each Node In XML.DocumentElement.SelectNodes("row")
			       i=i+1
				   strReplace = Replace(strHead,"[ClassName]",KS.GetClassNP(node.selectsinglenode("@tid").text))
				   strReplace = Replace(strReplace,"[AutoID]",I)
			       TempStr = TempStr&strReplace&"<a href=""" & KS.GetInfoUrl(ChannelID,node.selectsinglenode("@id").text,node.selectsinglenode("@fname").text) & """>"&KS.GotTopic(node.selectsinglenode("@title").text,strTitleNum)&"</a>"&strTail&""				   
			Next
			GetRandomContentsList = TempStr
		End Function

		'**************************************************
		'函数名：GetRelatedContentsList
		'作  用：显示内容页相关列表
		'参  数：strHead--头导航类型
		'       strTail--尾导航类型
		'       strNum--显示记录数
		'       strTitleNum--链接标题字符
		'**************************************************
		Function GetRelatedContentsList(strHead,strTail,strNum,strTitleNum)
			Dim RS,SqlStr,ChannelID,TempStr,XML,Node
			ChannelID=KS.ChkClng(FCls.ChannelID)
			If ChannelID=0 Then Exit Function
			SQLStr="Select top "&strNum&" ID,Tid,Title,Fname From " & KS.C_S(ChannelID,2) & " i Inner Join KS_ItemInfoR R On I.ID=R.RelativeID Where I.Verific=1 And I.DelTF=0 And R.InfoID=" & FCls.RefreshInfoID & " And R.RelativeChannelID=" & ChannelID & " Order By I.id desc"
			GetRelatedContentsList=sqlstr
			set RS=Conn.Execute(SqlStr)
			If Not RS.Eof Then
			  Set XML=KS.RsToXml(RS,"row","")
			End If
			RS.Close : Set RS=Nothing
			If Not IsObject(XML) Then
			   GetRelatedContentsList = "暂无相关链接!<br/>"
			   Exit Function
			End IF
			
			Dim strReplace,i:i=0
			For Each Node In XML.DocumentElement.SelectNodes("row")
			       i=i+1
				   strReplace = Replace(strHead,"[ClassName]",KS.GetClassNP(node.selectsinglenode("@tid").text))
				   strReplace = Replace(strReplace,"[AutoID]",I)
			       TempStr = TempStr & strReplace&"<a href=""" & KS.GetInfoUrl(ChannelID,node.selectsinglenode("@id").text,node.selectsinglenode("@fname").text) & """>"&KS.GotTopic(node.selectsinglenode("@title").text,strTitleNum)&"</a>"&strTail&""
			Next
			
			GetRelatedContentsList = TempStr
		End Function

		'**************************************************
		'函数名：查找内容的图片地址
		'作  用：替换通用标签为内容
		'参  数：ArticleContent原文件
		'**************************************************
		Function GetArticlePicUrl(ArticleContent)
		    On Error Resume Next
		    Dim Re,URLContents,URLContent,URLValue
		    Set Re = New Regexp
			Re.IgnoreCase = True
			Re.Global = True
			Re.Pattern = "(src=)('|"&CHR(34)&")(.*?)('|"&CHR(34)&")"
			Set URLContents=Re.Execute(ArticleContent)
			IF URLContents.Count<>0 Then
			   For Each URLContent in URLContents	
			       URLValue= URLValue & Replace(Replace(Replace(URLContent,"src=",""),CHR(34),""),"'","") & "|||"
			   Next
			End IF
			set Re=nothing
			GetArticlePicUrl = Left(Trim(URLValue), Len(Trim(URLValue)) - 3)
	    End Function
		
		Function GetPhoto(ArticleContent,ID,ChannelID,PhotoUrl,width,height)
		    Dim TempStr,ArticlePicUrl,PicUrlsArr,TotalPage
			   if left(PhotoUrl,1)="/" then PhotoUrl=right(PhotoUrl,len(PhotoUrl)-1)
			   if lcase(left(PhotoUrl,4))<>"http" then PhotoUrl=KS.Setting(2) & KS.Setting(3) & PhotoUrl
			   if instr(lcase(PhotoUrl),"nopic.gif")<>0 then exit function
			TempStr = "<img src=""" & PhotoUrl & """  width=""" & width & """ height=""" & height & """ alt="".""/><br/>"
			ArticlePicUrl = GetArticlePicUrl(KS.GetEncodeConversion(ArticleContent))
			PicUrlsArr = Split(ArticlePicUrl, "|||")
			TotalPage = Cint(UBound(PicUrlsArr) + 1)
			If TotalPage > 1 Then
			   Call KS.GetWriteinReturn("<a href=""" & KS.GetUrl & """>返回" & KS.C_S(ChannelID,3) & "页</a>")
			   TempStr=TempStr&"<a href=""" & DomainStr & "Plus/PhotoDownLoad.asp?JpegUrl=" & PhotoUrl & "&"&WapValue&""">查看</a> "
			End If
			GetPhoto=TempStr
		End Function
End Class
%>
