<%Option Explicit%>
<!--#include File="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%

'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim SearchCls
Set SearchCls = New SearchResult
SearchCls.Kesion()
Set SearchCls = Nothing

Const  FuzzySearch = 1  '设为1支持模糊查找，但会加大系统资源的开销，如比如搜索“xp 2003”，包含xp和2003两者的、只包含其中一个的，都能搜索出来。
Class SearchResult
    Private KS,KMR,F_C,LoopContent,SearchResult,photourl
	Private ChannelID,ClassID,SearchType,KeyWord,SearchForm
    Private I,TotalPut, CurrentPage,MaxPerPage,RS,KeyWordArr
   
	Private Sub Class_Initialize()
		Set KS=New PublicCls
		Set KMR=New Refresh
		MaxPerPage=10
		If KS.S("page") <> "" Then
          CurrentPage = CInt(Request("page"))
        Else
          CurrentPage = 1
        End If
		Dim RefreshTime:RefreshTime = 2  '设置防刷新时间
		If DateDiff("s", Session("SearchTime"), Now()) < RefreshTime Then
			Response.Write "<META http-equiv=Content-Type content=text/html; chaRset=gb2312><meta HTTP-EQUIV=REFRESH CONTENT="&RefreshTime&"><br>本页面起用了防刷新机制，请不要在"&RefreshTime&"秒内连续刷新本页面<BR>正在打开页面，请稍后……"
			Response.End
		End If
		Session("SearchTime")=Now()
        ChannelID=KS.ChkClng(KS.S("ChannelID"))

		If ChannelID=0 Then Call KS.AlertHintScript("你没有选择搜索类型!"):response.end
        ClassID=KS.S("ClassID"):If ClassID="" Then ClassID="0"
        SearchType=KS.ChkCLng(KS.S("SearchType"))
        KeyWord=KS.CheckXSS(KS.R(KS.S("KeyWord")))
		If KeyWord="" Then KeyWord=KS.CheckXSS(KS.S("Tags"))
		KeyWordArr=Split(KeyWord," ")
		If KeyWord="" and channelid<>8 Then Call KS.AlertHintScript("你没有输入搜索关键字!"):response.end

	End Sub

	Private Sub Class_Terminate()
        closeconn
	    Set KS=Nothing
		Set KMR=Nothing
	End Sub
  
 Sub Kesion()
           if KS.C_S(ChannelID,33)="" then
		    response.write "对不起，还没有绑定搜索模板!"
			response.end
		   else
			F_C = KMR.LoadTemplate(KS.C_S(ChannelID,33))
		   end if
		   If Trim(F_C) = "" Then F_C = "模板不存在!"
		   
		   FCls.RefreshType = "search" '设置刷新类型，以便取得当前位置导航等
		   Fcls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
		   
		    LoopContent=KS.CutFixContent(F_C, "[loop]", "[/loop]", 0)
			Call LoadSearch()
			F_C = KMR.KSLabelReplaceAll(F_C) 
			Response.Write F_C
 End Sub

 Sub LoadSearch()
   Select Case KS.C_S(ChannelID,6)
     Case 1:Call ArticleSearch()
     Case 2:Call PhotoSearch()
     Case 3:Call DownSearch()
	 Case 4:Call FlashSearch()
	 Case 5:Call ProductSearch()
	 Case 6:Call MusicSearch()
	 Case 7:Call MovieSearch()
	 Case 8:Call SupplySearch()
	 Case else
	   Call KS.AlertHintScript("你没有选择搜索类型!"):response.end
    End Select
	
	F_C = Replace(F_C,"{$GetSearchKey}",keyword)
	F_C = Replace(F_C,"{$ShowTotal}",totalput)
	F_C = Replace(F_C,"{$GetMusicSearchResult}",SearchResult)
	F_C = Replace(F_C,KS.CutFixContent(F_C, "[loop]", "[/loop]", 1),SearchResult)
	F_C = Replace(F_C,"{$ShowPage}",KS.ShowPage(totalput, MaxPerPage, "", CurrentPage,false,false))
  End Sub
  
  
  Sub ArticleSearch()         
  Dim SqlStr,Param
  Param=" Where Verific=1 And DelTF=0"
  Select Case SearchType
   Case 100
     If IsDate(KeyWord) Then
      If CInt(DataBaseType) = 1 Then
       Param=Param & " And AddDate>='" & KeyWord & " 00:00:00' and AddDate<='" &KeyWord & " 23:59:59'"
	  else
	   Param=Param & " And AddDate>=#" & KeyWord & " 00:00:00# and AddDate<=#" &KeyWord & " 23:59:59#"
	  end if
	 Else
	   Exit Sub
	 End If
   Case 1
     If (FuzzySearch=1) Then
	  For I=0 To Ubound(KeyWordArr)
	   If I=0 Then
	   Param=Param & " And (Title Like '%" & KeyWordArr(i) & "%'"
	   Else
	   Param = Param & " or Title Like '%" & KeyWordArr(i) & "%'"
	   End If
	  Next
	 Else
     Param=Param & " And (Title Like '%" & KeyWord & "%'"
	 End If
	 Param=Param & ")"
   Case 2:Param=Param & " And ArticleContent Like '%" & KeyWord & "%'"
   Case 3:Param=Param & " And Author Like '%" & KeyWord & "%'"
   Case 4:Param=Param & " And Inputer Like '%" & KeyWord & "%'"
   Case 5:Param=Param & " And KeyWords Like '%" & KeyWord & "%'"
   Case else
    if (FuzzySearch=1) Then
	  For I=0 To Ubound(KeyWordArr)
	   If I=0 Then
	   Param=Param & " And (Title Like '%" & KeyWordArr(i) & "%'"
	   Else
	   Param = Param & " or Title Like '%" & KeyWordArr(i) & "%'"
	   End If
	  Next
	 Else
     Param=Param & " And (Title Like '%" & KeyWord & "%' or Author Like '%" & KeyWord & "%'"
	 End If
	 Param=Param & ")"
  End Select
  If ClassID<>"0" Then Param=Param & " And Tid In(" & KS.GetFolderTid(ClassID) & ")"
	 SqlStr="Select * From " & KS.C_S(ChannelID,2) & Param & " Order By ID Desc"
  Set RS=Server.CreateObject("ADODB.RECORDSET")
  RS.Open SqlStr,Conn,1,3

  IF RS.Eof And RS.Bof Then
      totalput=0
      SearchResult = "没有找到任何信息!"
	  exit sub
  Else
					TotalPut= Conn.Execute("Select count(*) from " & KS.C_S(Channelid,2) & Param)(0)
                    If CurrentPage < 1 Then CurrentPage = 1

                    If (CurrentPage - 1) * MaxPerPage > totalPut Then
                        If (TotalPut Mod MaxPerPage) = 0 Then
                            CurrentPage = totalPut \ MaxPerPage
                        Else
                            CurrentPage = totalPut \ MaxPerPage + 1
                        End If
                    End If

                    If CurrentPage = 1 Then
                             Call GetSearchResult
                    Else
                        If (CurrentPage - 1) * MaxPerPage < totalPut Then
                            RS.Move (CurrentPage - 1) * MaxPerPage
                            Call GetSearchResult
                        Else
                            CurrentPage = 1
                            Call GetSearchResult
                        End If
                    End If
    End IF
	RS.Close
	Set RS=Nothing
  End Sub   

  
  Sub PhotoSearch()
   Dim SqlStr,Param
  SqlStr="Select * From " & KS.C_S(Channelid,2)
  Param=" Where Verific=1 And DelTF=0 "
  Select Case SearchType
   Case 1
     If (FuzzySearch=1) Then
	  For I=0 To Ubound(KeyWordArr)
	   If I=0 Then
	   Param=Param & " And (Title Like '%" & KeyWordArr(i) & "%'"
	   Else
	   Param = Param & " or Title Like '%" & KeyWordArr(i) & "%'"
	   End If
	  Next
	 Else
     Param=Param & " And (Title Like '%" & KeyWord & "%'"
	 End If
	 Param=Param & ")"
   Case 2: Param=Param & " And PictureContent Like '%" & KeyWord & "%'"
   Case 3: Param=Param & " And Author Like '%" & KeyWord & "%'"
   Case 4: Param=Param & " And Inputer Like '%" & KeyWord & "%'"
   Case 5: Param=Param & " And KeyWords Like '%" & KeyWord & "%'"
   Case else
    if (FuzzySearch=1) Then
	  For I=0 To Ubound(KeyWordArr)
	   If I=0 Then
	   Param=Param & " And (Title Like '%" & KeyWordArr(i) & "%'"
	   Else
	   Param = Param & " or Title Like '%" & KeyWordArr(i) & "%'"
	   End If
	  Next
	 Else
     Param=Param & " And (Title Like '%" & KeyWord & "%' or Author Like '%" & KeyWord & "%'"
	 End If
	 Param=Param & ")"
  End Select
    If ClassID<>"0" Then: Param=Param & " And Tid In(" & KS.GetFolderTid(ClassID) & ")"
    SqlStr=SqlStr & Param & " Order By AddDate Desc"

  Set RS=Server.CreateObject("ADODB.RECORDSET")
  RS.Open SqlStr,Conn,1,1
  IF RS.Eof And RS.Bof Then
      totalput=0
      SearchResult = "没有找到任何信息!"
	  exit sub
  Else
                    TotalPut = Conn.Execute("Select Count(ID) From "&KS.C_S(Channelid,2) & Param)(0)
                    If CurrentPage < 1 Then CurrentPage = 1

                    If (CurrentPage - 1) * MaxPerPage > totalPut Then
                        If (TotalPut Mod MaxPerPage) = 0 Then
                            CurrentPage = totalPut \ MaxPerPage
                        Else
                            CurrentPage = totalPut \ MaxPerPage + 1
                        End If
                    End If

                    If CurrentPage = 1 Then
                            Call GetSearchResult
                    Else
                        If (CurrentPage - 1) * MaxPerPage < totalPut Then
                            RS.Move (CurrentPage - 1) * MaxPerPage
                            Call GetSearchResult
                        Else
                            CurrentPage = 1
                            Call GetSearchResult
                        End If
                    End If
    End IF
	RS.Close:Set RS=Nothing
  End Sub
  

 Sub DownSearch() 
   Dim SqlStr,Param
  SqlStr="Select * From " & KS.C_S(Channelid,2)
  Param=" Where Verific=1 And DelTF=0 "
  Select Case SearchType
   Case 1
     If (FuzzySearch=1) Then
	  For I=0 To Ubound(KeyWordArr)
	   If I=0 Then
	   Param=Param & " And (Title Like '%" & KeyWordArr(i) & "%'"
	   Else
	   Param = Param & " or Title Like '%" & KeyWordArr(i) & "%'"
	   End If
	  Next
	 Else
     Param=Param & " And (Title Like '%" & KeyWord & "%'"
	 End If
	 Param=Param & ")"
   Case 2:Param=Param & " And DownContent Like '%" & KeyWord & "%'"
   Case 3:Param=Param & " And Author Like '%" & KeyWord & "%'"
   Case 4:Param=Param & " And Inputer Like '%" & KeyWord & "%'"
   Case 5:Param=Param & " And KeyWords Like '%" & KeyWord & "%'"
   Case else
    if (FuzzySearch=1) Then
	  For I=0 To Ubound(KeyWordArr)
	   If I=0 Then
	   Param=Param & " And (Title Like '%" & KeyWordArr(i) & "%'"
	   Else
	   Param = Param & " or Title Like '%" & KeyWordArr(i) & "%'"
	   End If
	  Next
	 Else
     Param=Param & " And (Title Like '%" & KeyWord & "%'"
	 End If
	 Param=Param & ")"
  End Select
    If ClassID<>"0" Then Param=Param & " And Tid In(" & KS.GetFolderTid(ClassID) & ")"
     SqlStr=SqlStr & Param & " Order By AddDate Desc"

  Set RS=Server.CreateObject("ADODB.RECORDSET")
  RS.Open SqlStr,Conn,1,1
  IF RS.Eof And RS.Bof Then
      totalput=0
      SearchResult = "没有找到任何信息!"
	  exit sub
  Else
                    TotalPut = Conn.Execute("Select Count(ID) From " & KS.C_S(Channelid,2)&Param)(0)
                    If CurrentPage < 1 Then CurrentPage = 1

                    If (CurrentPage - 1) * MaxPerPage > totalPut Then
                        If (TotalPut Mod MaxPerPage) = 0 Then
                            CurrentPage = totalPut \ MaxPerPage
                        Else
                            CurrentPage = totalPut \ MaxPerPage + 1
                        End If
                    End If

                    If CurrentPage = 1 Then
                            Call GetSearchResult
                    Else
                        If (CurrentPage - 1) * MaxPerPage < totalPut Then
                            RS.Move (CurrentPage - 1) * MaxPerPage
                            Call GetSearchResult
                        Else
                            CurrentPage = 1
                           Call GetSearchResult
                        End If
                    End If
    End IF
	RS.Close:Set RS=Nothing
  End Sub
  
  Sub FlashSearch() 
    Dim SqlStr,Param
    SqlStr="Select * From KS_Flash"
	Param=" Where Verific=1 And DelTF=0 "
   Select Case SearchType
    Case 1
	 If (FuzzySearch=1) Then
	  For I=0 To Ubound(KeyWordArr)
	   If I=0 Then
	   Param=Param & " And (Title Like '%" & KeyWordArr(i) & "%'"
	   Else
	   Param = Param & " or Title Like '%" & KeyWordArr(i) & "%'"
	   End If
	  Next
	 Else
     Param=Param & " And (Title Like '%" & KeyWord & "%'"
	 End If
	 Param=Param & ")"
    Case 2:Param=Param & " And FlashContent Like '%" & KeyWord & "%'"
    Case 3:Param=Param & " And Author Like '%" & KeyWord & "%'"
    Case 4:Param=Param & " And Inputer Like '%" & KeyWord & "%'"
    Case 5:Param=Param & " And KeyWords Like '%" & KeyWord & "%'"
   Case else
    if (FuzzySearch=1) Then
	  For I=0 To Ubound(KeyWordArr)
	   If I=0 Then
	   Param=Param & " And (Title Like '%" & KeyWordArr(i) & "%'"
	   Else
	   Param = Param & " or Title Like '%" & KeyWordArr(i) & "%'"
	   End If
	  Next
	 Else
     Param=Param & " And (Title Like '%" & KeyWord & "%'"
	 End If
	 Param=Param & ")"
  End Select

    If ClassID<>"0" Then Param=Param & " And Tid In(" & KS.GetFolderTid(ClassID) & ")"
     SqlStr=SqlStr & Param & " Order By AddDate Desc"

  Set RS=Server.CreateObject("ADODB.RECORDSET")
  RS.Open SqlStr,Conn,1,1
  IF RS.Eof And RS.Bof Then
      totalput=0
      SearchResult = "没有找到任何信息!"
	  exit sub
  Else
                    TotalPut = Conn.Execute("Select Count(ID) From KS_Flash"&Param)(0)
                    If CurrentPage < 1 Then CurrentPage = 1

                    If (CurrentPage - 1) * MaxPerPage > totalPut Then
                        If (TotalPut Mod MaxPerPage) = 0 Then
                            CurrentPage = totalPut \ MaxPerPage
                        Else
                            CurrentPage = totalPut \ MaxPerPage + 1
                        End If
                    End If

                    If CurrentPage = 1 Then
                           Call GetSearchResult
                    Else
                        If (CurrentPage - 1) * MaxPerPage < totalPut Then
                            RS.Move (CurrentPage - 1) * MaxPerPage
                           Call GetSearchResult
                        Else
                            CurrentPage = 1
                            Call GetSearchResult
                        End If
                    End If
    End IF
	RS.Close
	Set RS=Nothing
  End Sub  
  
  Sub ProductSearch() 'Product搜索处理
    Dim SqlStr,Param
    SqlStr="Select * From KS_Product"
	Param=" Where Verific=1 And DelTF=0 "
   Select Case SearchType
    Case 1
	 If (FuzzySearch=1) Then
	  For I=0 To Ubound(KeyWordArr)
	   If I=0 Then
	   Param=Param & " And (Title Like '%" & KeyWordArr(i) & "%'"
	   Else
	   Param = Param & " or Title Like '%" & KeyWordArr(i) & "%'"
	   End If
	  Next
	 Else
     Param=Param & " And (Title Like '%" & KeyWord & "%'"
	 End If
	 Param=Param & ")"
    Case 2:Param=Param & " And ProducerName Like '%" & KeyWord & "%'"
    Case 3:Param=Param & " And ProIntro Like '%" & KeyWord & "%'"
    Case 5:Param=Param & " And KeyWords Like '%" & KeyWord & "%'"
   Case else
    if (FuzzySearch=1) Then
	  For I=0 To Ubound(KeyWordArr)
	   If I=0 Then
	   Param=Param & " And (Title Like '%" & KeyWordArr(i) & "%'"
	   Else
	   Param = Param & " or Title Like '%" & KeyWordArr(i) & "%'"
	   End If
	  Next
	 Else
     Param=Param & " And (Title Like '%" & KeyWord & "%' or ProducerName Like '%" & KeyWord & "%'"
	 End If
	 Param=Param & ")"
  End Select
    If ClassID<>"0" Then Param=Param & " And Tid In(" & KS.GetFolderTid(ClassID) & ")"
     SqlStr=SqlStr & Param & " Order By AddDate Desc"

  Set RS=Server.CreateObject("ADODB.RECORDSET")
  RS.Open SqlStr,Conn,1,1
  IF RS.Eof And RS.Bof Then
      totalput=0
      SearchResult = "没有找到任何信息!"
	  exit sub
  Else
                    TotalPut = Conn.Execute("Select Count(ID) From KS_Product"&Param)(0)
                    If CurrentPage < 1 Then CurrentPage = 1

                    If (CurrentPage - 1) * MaxPerPage > totalPut Then
                        If (TotalPut Mod MaxPerPage) = 0 Then
                            CurrentPage = totalPut \ MaxPerPage
                        Else
                            CurrentPage = totalPut \ MaxPerPage + 1
                        End If
                    End If

                    If CurrentPage = 1 Then
                            Call GetSearchResult
                    Else
                        If (CurrentPage - 1) * MaxPerPage < totalPut Then
                            RS.Move (CurrentPage - 1) * MaxPerPage
                            Call GetSearchResult
                        Else
                            CurrentPage = 1
                            Call GetSearchResult
                        End If
                    End If
    End IF
	RS.Close
	Set RS=Nothing
  End Sub

  Sub MusicSearch()  '音乐搜索处理
   Dim SqlStr
   MaxPerPage=30          '每页显示的歌曲数量
   Select Case SearchType
     Case 0,1
	     SqlStr="Select * From KS_MSSongList Where MusicName Like '%" & KeyWord & "%' Order By ID Desc"
	 Case 2
	     SqlStr="Select * From KS_MSSongList Where Singer Like '%" & KeyWord & "%' Order By ID Desc"
	 Case 3
	     SqlStr="Select * From KS_MSSpecial Where Name Like '%" & KeyWord & "%' Order By SpecialID Desc"
   End Select
	  Set RS=Server.CreateObject("ADODB.RECORDSET")
	  RS.Open SqlStr,Conn,1,1
	  IF RS.Eof And RS.Bof Then
		  totalput=0
		  SearchResult = "没有找到任何信息!"
		  exit sub
	  Else
                    TotalPut = RS.RecordCount

                    If CurrentPage < 1 Then
                        CurrentPage = 1
                    End If

                    If (CurrentPage - 1) * MaxPerPage > totalPut Then
                        If (TotalPut Mod MaxPerPage) = 0 Then
                            CurrentPage = totalPut \ MaxPerPage
                        Else
                            CurrentPage = totalPut \ MaxPerPage + 1
                        End If
                    End If

                    If CurrentPage = 1 Then
                            Call ShowMusicContent(SqlStr)
                    Else
                        If (CurrentPage - 1) * MaxPerPage < totalPut Then
                            RS.Move (CurrentPage - 1) * MaxPerPage
                            Call ShowMusicContent(SqlStr)
                        Else
                            CurrentPage = 1
                            Call ShowMusicContent(SqlStr)
                        End If
                    End If
    End IF
  End Sub
  
  Sub ShowMusicContent(SqlStr)
      Dim ItemUnit
	  Dim KSCMUSIC:Set KSCMUSIC=New RefreshMusicCls
	  Select Case SearchType
		   Case 0,1,2
		    SearchResult="<strong>" & KSCMUSIC.GetPlayList(rs,1,MaxPerPage,20,1,1,1)
		    ItemUnit="首"
		   Case 3
		   	SearchResult= "<strong>" & KSCMUSIC.GetSpecialList(RS,MaxPerPage,3,110,80,20,1)
		     ItemUnit="张"
	   End Select
	   Set KSCMUSIC=Nothing
  End Sub


  Sub MovieSearch() 'Movie搜索处理
    Dim SqlStr,Param
    SqlStr="Select * From KS_Movie"
	Param=" Where Verific=1 And DelTF=0 "
   Select Case SearchType
    Case 1
	 If (FuzzySearch=1) Then
	  For I=0 To Ubound(KeyWordArr)
	   If I=0 Then
	   Param=Param & " And (Title Like '%" & KeyWordArr(i) & "%'"
	   Else
	   Param = Param & " or Title Like '%" & KeyWordArr(i) & "%'"
	   End If
	  Next
	 Else
     Param=Param & " And (Title Like '%" & KeyWord & "%'"
	 End If
	 Param=Param & ")"
    Case 2:Param=Param & " And MovieAct Like '%" & KeyWord & "%'"
    Case 3:Param=Param & " And MovieContent Like '%" & KeyWord & "%'"
    Case 5:Param=Param & " And KeyWords Like '%" & KeyWord & "%'"
    Case else:Param=Param & " And (Title Like '%" & KeyWord & "%' Or  MovieAct Like '%" & KeyWord & "%')" 
  End Select
    If ClassID<>"0" Then Param=Param & " And Tid In(" & KS.GetFolderTid(ClassID) & ")"
     SqlStr=SqlStr & Param & " Order By AddDate Desc"

  Set RS=Server.CreateObject("ADODB.RECORDSET")
  RS.Open SqlStr,Conn,1,1
  IF RS.Eof And RS.Bof Then
		  totalput=0
		  SearchResult = "没有找到任何信息!"
		  exit sub
  Else
                    TotalPut = Conn.Execute("Select Count(ID) From KS_Movie" & Param)(0)
                    If CurrentPage < 1 Then
                        CurrentPage = 1
                    End If

                    If (CurrentPage - 1) * MaxPerPage > totalPut Then
                        If (TotalPut Mod MaxPerPage) = 0 Then
                            CurrentPage = totalPut \ MaxPerPage
                        Else
                            CurrentPage = totalPut \ MaxPerPage + 1
                        End If
                    End If

                    If CurrentPage = 1 Then
                            Call GetSearchResult
                    Else
                        If (CurrentPage - 1) * MaxPerPage < totalPut Then
                            RS.Move (CurrentPage - 1) * MaxPerPage
                            Call GetSearchResult
                        Else
                            CurrentPage = 1
                            Call GetSearchResult
                        End If
                    End If
    End IF
	RS.Close
	Set RS=Nothing
  End Sub
  


  Sub SupplySearch() 
    Dim SqlStr,Param
    SqlStr="Select * From KS_GQ"
	Param=" Where Verific=1 And DelTF=0 "
	If Not KS.IsNul(KeyWord) Then
   Select Case SearchType
    Case 1
	 If (FuzzySearch=1) Then
	  For I=0 To Ubound(KeyWordArr)
	   If I=0 Then
	   Param=Param & " And (Title Like '%" & KeyWordArr(i) & "%'"
	   Else
	   Param = Param & " or Title Like '%" & KeyWordArr(i) & "%'"
	   End If
	  Next
	 Else
     Param=Param & " And (Title Like '%" & KeyWord & "%'"
	 End If
	 Param=Param & ")"
    Case 2:Param=Param & " And UserName Like '%" & KeyWord & "%'"
    Case 3:Param=Param & " And GQContent Like '%" & KeyWord & "%'"
    Case 5:Param=Param & " And KeyWords Like '%" & KeyWord & "%'"
   Case else
    if (FuzzySearch=1) Then
	  For I=0 To Ubound(KeyWordArr)
	   If I=0 Then
	   Param=Param & " And (Title Like '%" & KeyWordArr(i) & "%'"
	   Else
	   Param = Param & " or Title Like '%" & KeyWordArr(i) & "%'"
	   End If
	  Next
	 Else
     Param=Param & " And (Title Like '%" & KeyWord & "%' or UserName Like '%" & KeyWord & "%'"
	 End If
	 Param=Param & ")"
  End Select
  End If
    If ClassID<>"0" Then Param=Param & " And Tid In(" & KS.GetFolderTid(ClassID) & ")"
		If Not KS.IsNul(KS.S("Province")) Then Param=Param & " and province='" & KS.S("Province") & "'"
	If Not KS.IsNul(KS.S("City")) Then Param=Param & " and City='" & KS.S("City") & "'"

     SqlStr=SqlStr & Param & " Order By AddDate Desc"

  Set RS=Server.CreateObject("ADODB.RECORDSET")
  RS.Open SqlStr,Conn,1,1
  IF RS.Eof And RS.Bof Then
		  totalput=0
		  SearchResult = "没有找到任何信息!"
		  exit sub
  Else
                    TotalPut = Conn.Execute("Select count(id) from KS_GQ" &Param)(0)
                    If CurrentPage < 1 Then
                        CurrentPage = 1
                    End If

                    If (CurrentPage - 1) * MaxPerPage > totalPut Then
                        If (TotalPut Mod MaxPerPage) = 0 Then
                            CurrentPage = totalPut \ MaxPerPage
                        Else
                            CurrentPage = totalPut \ MaxPerPage + 1
                        End If
                    End If

                    If CurrentPage = 1 Then
                            Call GetSearchResult()
                    Else
                        If (CurrentPage - 1) * MaxPerPage < totalPut Then
                            RS.Move (CurrentPage - 1) * MaxPerPage
                            Call GetSearchResult()
                        Else
                            CurrentPage = 1
                            Call GetSearchResult()
                        End If
                    End If
    End IF
	RS.Close
	Set RS=Nothing
  End Sub
  
  Sub GetSearchResult()  
      on error resume next 
	  I=0
       Dim LC
	  Select Case KS.C_S(ChannelID,6)
	   Case 1
		  Do While Not RS.Eof
			 If Not Response.IsClientConnected Then Response.end
			  LC=LoopContent
			  LC=replace(LC,"{$GetArticleTitle}",ReplaceKeyWordRed(rs("title")))
			  LC=replace(LC,"{$GetArticleHits}",rs("hits"))
			  LC=replace(LC,"{$GetArticleAuthor}",rs("author"))
			  LC=replace(LC,"{$GetArticleInput}",rs("inputer"))
			  LC=replace(LC,"{$GetArticleOrigin}",rs("origin"))
			  LC=replace(LC,"{$GetArticleDate}",rs("adddate"))
			  if isnull(rs("intro")) then
			  LC=replace(LC,"{$GetArticleIntro}",ReplaceKeyWordRed(KS.GotTopic(KS.LoseHtml(rs("articlecontent")),200)))
			  else
			  LC=replace(LC,"{$GetArticleIntro}",ReplaceKeyWordRed(rs("intro")))
			  end if
			  photourl=rs("photourl")
			  if photourl="" then photourl=ks.setting(3) & "images/nopic.gif"
			  LC=replace(LC,"{$GetArticlePic}",photourl)
			  LC=replace(LC,"{$GetClassNameAndPath}",KS.GetClassNP(rs("tid")))
			  LC=replace(LC,"{$GetArticleUrl}",KS.GetItemUrl(ChannelID,rs("Tid"),rs("ID"),rs("Fname")))
			  LC=LFCls.ReplaceUserDefine(ChannelID,LC,RS)
			  SearchResult=SearchResult & LC
			I = I + 1
			If I >= MaxPerPage Then Exit Do
			RS.MoveNext
		Loop
	 Case 2
	   Do While Not RS.Eof
			 If Not Response.IsClientConnected Then Response.end
			  LC=LoopContent
			  LC=replace(LC,"{$GetPictureTitle}",ReplaceKeyWordRed(rs("title")))
			  LC=replace(LC,"{$GetPictureHits}",rs("hits"))
			  LC=replace(LC,"{$GetPictureHitsByDay}",RS("hitsbyday"))
			  LC=replace(LC,"{$GetPictureHitsByWeek}",RS("hitsbyweek"))
			  LC=replace(LC,"{$GetPictureHitsByMonth}",RS("hitsbymonth"))
			  LC=replace(LC,"{$GetPictureAuthor}",rs("author"))
			  LC=replace(LC,"{$GetPictureInput}",rs("inputer"))
			  LC=replace(LC,"{$GetPictureOrigin}",rs("origin"))
			  LC=replace(LC,"{$GetPictureDate}",rs("adddate"))
			  LC=replace(LC,"{$GetPictureIntro}",ReplaceKeyWordRed(KS.Gottopic(KS.LoseHtml(rs("picturecontent")),200)))
			  LC=replace(LC,"{$GetPhotoUrl}",rs("photourl"))
			  LC=replace(LC,"{$GetClassNameAndPath}",KS.GetClassNP(rs("tid")))
			  LC=replace(LC,"{$GetPictureUrl}",KS.GetItemUrl(ChannelID,rs("Tid"),rs("ID"),rs("Fname")))
			  LC=LFCls.ReplaceUserDefine(ChannelID,LC,RS)
			  SearchResult=SearchResult & LC
	        I=I+1
			If I >= MaxPerPage Then Exit Do
			RS.MoveNext
		Loop
	 Case 3
	 	   Do While Not RS.Eof
			 If Not Response.IsClientConnected Then Response.end
			  LC=LoopContent
			  LC=replace(LC,"{$GetDownTitle}",ReplaceKeyWordRed(rs("title")))
			  LC=replace(LC,"{$GetDownHits}",rs("hits"))
			  LC=replace(LC,"{$GetDownHitsByDay}",RS("hitsbyday"))
			  LC=replace(LC,"{$GetDownHitsByWeek}",RS("hitsbyweek"))
			  LC=replace(LC,"{$GetDownHitsByMonth}",RS("hitsbymonth"))
			  LC=replace(LC,"{$GetDownAuthor}",rs("author"))
			  LC=replace(LC,"{$GetDownInput}",rs("inputer"))
			  LC=replace(LC,"{$GetDownOrigin}",rs("origin"))
			  LC=replace(LC,"{$GetDownDate}",rs("adddate"))
			  LC= Replace(LC, "{$GetDownSystem}", RS("DownPT"))
			  LC= Replace(LC, "{$GetDownAuthor}", RS("Author"))
			  LC= Replace(LC, "{$GetDownSize}", RS("DownSize"))
			  LC= Replace(LC, "{$GetDownType}", RS("DownLB"))
			  LC= Replace(LC, "{$GetDownLanguage}", RS("DownYY"))
			  LC= Replace(LC, "{$GetDownPower}", RS("DownSQ"))	
			  LC= Replace(LC, "{$GetDownStar}", Replace(RS("Rank"),"★","<img src=""../Images/Star.gif"" border=""0"">"))		  
			  LC=replace(LC,"{$GetDownIntro}",ReplaceKeyWordRed(KS.Gottopic(KS.LoseHtml(rs("downcontent")),200)))
			  photourl=rs("photourl")
			  if photourl="" then photourl=ks.setting(3) & "images/nopic.gif"

			  LC=replace(LC,"{$GetPhotoUrl}",photourl)
			  LC=replace(LC,"{$GetClassNameAndPath}",KS.GetClassNP(rs("tid")))
			  LC=replace(LC,"{$GetDownUrl}",KS.GetItemUrl(ChannelID,rs("Tid"),rs("ID"),rs("Fname")))
			  LC=LFCls.ReplaceUserDefine(ChannelID,LC,RS)
			  SearchResult=SearchResult & LC
	        I=I+1
			If I >= MaxPerPage Then Exit Do
			RS.MoveNext
		Loop
	 case 4
	 	   Do While Not RS.Eof
			 If Not Response.IsClientConnected Then Response.end
			  LC=LoopContent
			  LC=replace(LC,"{$GetFlashName}",ReplaceKeyWordRed(rs("title")))
			  LC=replace(LC,"{$GetFlashHits}",rs("hits"))
			  LC=replace(LC,"{$GetFlashHitsByDay}",RS("hitsbyday"))
			  LC=replace(LC,"{$GetFlashHitsByWeek}",RS("hitsbyweek"))
			  LC=replace(LC,"{$GetFlashHitsByMonth}",RS("hitsbymonth"))
			  LC=replace(LC,"{$GetFlashAuthor}",rs("author"))
			  LC=replace(LC,"{$GetFlashInput}",rs("inputer"))
			  LC=replace(LC,"{$GetFlashOrigin}",rs("origin"))
			  LC=replace(LC,"{$GetFlashDate}",rs("adddate"))
			  LC=replace(LC,"{$GetFlashIntro}",ReplaceKeyWordRed(KS.Gottopic(KS.LoseHtml(KS.HTMLCode(rs("flashcontent"))),200)))
			  photourl=rs("photourl")
			  if photourl="" then photourl=ks.setting(3) & "images/nopic.gif"

			  LC=replace(LC,"{$GetPhotoUrl}",photourl)
			  LC=replace(LC,"{$GetClassNameAndPath}",KS.GetClassNP(rs("tid")))
			  LC=replace(LC,"{$GetFlashUrl}",KS.GetItemUrl(ChannelID,rs("Tid"),rs("ID"),rs("Fname")))
			  LC=LFCls.ReplaceUserDefine(ChannelID,LC,RS)
			  SearchResult=SearchResult & LC
	        I=I+1
			If I >= MaxPerPage Then Exit Do
			RS.MoveNext
		Loop
     case 5
	 	   Do While Not RS.Eof
			 If Not Response.IsClientConnected Then Response.end
			  LC=LoopContent
			  LC=replace(LC,"{$GetProductName}",ReplaceKeyWordRed(rs("title")))
			LC = Replace(LC, "{$GetProductIntro}", ReplaceKeyWordRed(KS.LoseHtml(KS.HtmlCode(RS("ProIntro")))))
			LC = Replace(LC, "{$GetProductID}", RS("ProID"))
			LC = Replace(LC, "{$GetProductModel}", RS("ProModel"))
			LC = Replace(LC, "{$GetProductSpecificat}", RS("ProSpecificat"))
			LC = Replace(LC, "{$GetProducerName}", RS("ProducerName"))
			LC = Replace(LC, "{$GetTrademarkName}", RS("TrademarkName"))
			LC = Replace(LC, "{$GetServiceTerm}", RS("ServiceTerm"))
			LC = Replace(LC, "{$GetRank}",Replace(RS("Rank"),"★","<img src=""" & ks.setting(3) & "Images/Star.gif"" border=""0"">"))
			LC = Replace(LC, "{$GetTotalNum}",RS("TotalNum"))
			LC = Replace(LC, "{$GetProductUnit}", RS("Unit"))
            LC = Replace(LC, "{$GetProductHits}", rs("hits"))
			LC = Replace(LC, "{$GetProductDate}", RS("AddDate"))
			LC = Replace(LC, "{$GetPrice_Market}", RS("Price_Market"))
			LC = Replace(LC, "{$GetPrice}", RS("Price"))
			LC = Replace(LC, "{$GetPrice_Member}", RS("Price_Member"))
			LC = Replace(LC, "{$GetPrice_Original}", RS("Price_Original"))
			If RS("ProductType")=3 Then
			LC = Replace(LC, "{$GetDiscount}", RS("Discount"))
			Else
			LC = Replace(LC, "{$GetDiscount}", "")
			End If
			LC = Replace(LC, "{$GetScore}", RS("Point"))
			
			  photourl=rs("photourl")
			  if photourl="" then photourl=ks.setting(3) & "images/nopic.gif"

			  LC=replace(LC,"{$GetPhotoUrl}",photourl)
			  LC=replace(LC,"{$GetClassNameAndPath}",KS.GetClassNP(rs("tid")))
			  LC=replace(LC,"{$GetProductUrl}",KS.GetItemUrl(ChannelID,rs("Tid"),rs("ID"),rs("Fname")))
			  LC=LFCls.ReplaceUserDefine(ChannelID,LC,RS)
			  SearchResult=SearchResult & LC
	        I=I+1
			If I >= MaxPerPage Then Exit Do
			RS.MoveNext
		Loop
	 case 7
	 	   Do While Not RS.Eof
			 If Not Response.IsClientConnected Then Response.end
			  LC=LoopContent
			LC=replace(LC,"{$GetMovieName}",ReplaceKeyWordRed(rs("title")))
			LC = Replace(LC, "{$GetMovieIntro}",ReplaceKeyWordRed(KS.LoseHtml(RS("MovieContent"))))
			LC = Replace(LC, "{$GetMovieID}", RS("MovieID"))
			LC = Replace(LC, "{$GetMovieActor}", RS("MovieAct"))
			LC = Replace(LC, "{$GetMovieDirector}", RS("MovieDY"))
			LC = Replace(LC, "{$GetMovieTime}", RS("MovieTime"))
			LC = Replace(LC, "{$GetScreenTime}", RS("ScreenTime"))
			LC = Replace(LC, "{$GetMovieStar}",Replace(RS("Rank"),"★","<img src=""" & ks.setting(3) & "Images/Star.gif"" border=""0"">"))
			LC = Replace(LC, "{$GetMovieArea}",RS("MovieDQ"))
			LC = Replace(LC, "{$GetMovieLanguage}",RS("MovieYY"))
            LC = Replace(LC, "{$GetMovieHits}", rs("hits"))
			LC = Replace(LC, "{$GetMovieDate}", RS("AddDate"))
			  
			
			  photourl=rs("photourl")
			  if photourl="" then photourl=ks.setting(3) & "images/nopic.gif"

			  LC=replace(LC,"{$GetPhotoUrl}",photourl)
			  LC=replace(LC,"{$GetClassNameAndPath}",KS.GetClassNP(rs("tid")))
			  LC=replace(LC,"{$GetMovieUrl}",KS.GetItemUrl(ChannelID,rs("Tid"),rs("ID"),rs("Fname")))
			  LC=LFCls.ReplaceUserDefine(ChannelID,LC,RS)
			  SearchResult=SearchResult & LC
	        I=I+1
			If I >= MaxPerPage Then Exit Do
			RS.MoveNext
		Loop	
	 case 8
	 	   Do While Not RS.Eof
			 If Not Response.IsClientConnected Then Response.end
			  LC=LoopContent
			LC=replace(LC,"{$GetGQTitle}",ReplaceKeyWordRed(rs("title")))
			LC = Replace(LC, "{$GetGQIntro}",ReplaceKeyWordRed(KS.LoseHtml(KS.HTMLCode(RS("GQContent")))))	
			LC = Replace(LC, "{$GetGQInfoID}", RS("GQID"))
			LC = Replace(LC, "{$GetPrice}", RS("Price"))
			LC = Replace(LC, "{$GetTransType}", KS.GetGQTypeName(RS("TypeID")))
			LC = Replace(LC, "{$GetInfoType}", KS.C_C(RS("tid"),1))
			LC = Replace(LC, "{$GetValidTime}", KMR.GetValidTime(RS("ValidDate")))
			LC = Replace(LC, "{$GetGQContent}", KS.HtmlCode(RS("GQContent")))
			LC = Replace(LC, "{$GetAddDate}",RS("AddDate"))
            LC = Replace(LC, "{$GetGQHits}", rs("hits"))
			LC = Replace(LC, "{$GetGQDate}", RS("AddDate"))
			LC = Replace(LC, "{$GetInput}", RS("UserName"))
			LC = Replace(LC, "{$GetCompanyName}", RS("CompanyName"))
			LC = Replace(LC, "{$GetContactMan}", RS("ContactMan"))
			LC = Replace(LC, "{$GetContactTel}", RS("Tel"))
			LC = Replace(LC, "{$GetFax}", RS("Fax"))
			LC = Replace(LC, "{$GetAddress}", RS("Address"))
			LC = Replace(LC, "{$GetEmail}", RS("Email"))
			LC = Replace(LC, "{$GetPostCode}", RS("zip"))
			LC = Replace(LC, "{$GetProvince}", RS("Province"))
			LC = Replace(LC, "{$GetCity}", RS("City"))
			
			  photourl=rs("photourl")
			  if photourl="" then photourl=ks.setting(3) & "images/nopic.gif"

			  LC=replace(LC,"{$GetPhotoUrl}",photourl)
			  LC=replace(LC,"{$GetClassNameAndPath}",KS.GetClassNP(rs("tid")))
			  LC=replace(LC,"{$GetGQInfoUrl}",KS.GetItemUrl(ChannelID,rs("Tid"),rs("ID"),rs("Fname")))
			  LC=LFCls.ReplaceUserDefine(ChannelID,LC,RS)
			  SearchResult=SearchResult & LC
	        I=I+1
			If I >= MaxPerPage Then Exit Do
			RS.MoveNext
		Loop
	End Select
  End Sub  
  
  Function ReplaceKeyWordRed(Content)
   Dim I
   For I=0 To Ubound(KeyWordArr)
    Content=Replace(Content,KeyWordArr(i),"<font color=""#ff0000"">" &KeyWordArr(i) & "</font>")
   Next
   ReplaceKeyWordRed=Content
  End Function

End Class
%> 