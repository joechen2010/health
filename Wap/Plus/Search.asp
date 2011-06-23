<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
'********************************
'* 程序功能：搜索结果
'* 演示地址: http://wap.kesion.com/
'********************************
Response.ContentType="text/vnd.wap.wml"
Response.Charset="utf-8"
Response.Write "<?xml version=""1.0"" encoding=""utf-8""?>" &vbcrlf
Response.Write "<!DOCTYPE wml PUBLIC ""-//WAPFORUM//DTD WML 1.1//EN"" ""http://www.wapforum.org/DTD/wml_1.1.xml"">" &vbcrlf
%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
Dim SearchCls
Set SearchCls = New SearchResult
SearchCls.Kesion()
Set SearchCls = Nothing

Const  FuzzySearch = 1  '设为1支持模糊查找，但会加大系统资源的开销，如比如搜索“xp 2003”，包含xp和2003两者的、只包含其中一个的，都能搜索出来。
Class SearchResult
    Private KS,KMR,F_C,LoopContent,SearchResult,PhotoUrl
	Private ChannelID,ClassID,SearchType,KeyWord,SearchForm
    Private I,TotalPut, CurrentPage,MaxPerPage,RS,KeyWordArr
   
	Private Sub Class_Initialize()
		Set KS=New PublicCls
		Set KMR=New Refresh

		If KS.S("page") <> "" Then
           CurrentPage = CInt(Request("page"))
        Else
           CurrentPage = 1
        End If
        ChannelID=KS.ChkClng(KS.S("ChannelID"))

		If ChannelID=0 Then Call KS.ShowError("错误信息！","你没有选择搜索类型！")
        ClassID=KS.S("ClassID"):If ClassID="" Then ClassID="0"
        SearchType=KS.ChkCLng(KS.S("SearchType"))
        KeyWord=KS.S("KeyWord")
		If KeyWord="" Then KeyWord=KS.S("Tags")
		KeyWordArr=Split(KeyWord," ")
		If KeyWord="" Then Call KS.ShowError("错误信息！","你没有输入搜索关键字!")

		Dim RefreshTime:RefreshTime = 2  '设置防刷新时间
		If DateDiff("s", Session("SearchTime"), Now()) < RefreshTime Then
		   Response.Write "<wml>" &vbcrlf
		   Response.Write "<head>" &vbcrlf
		   Response.Write "<meta http-equiv=""Cache-Control"" content=""no-Cache""/>" &vbcrlf
		   Response.Write "<meta http-equiv=""Cache-Control"" content=""max-age=0""/>" &vbcrlf
		   Response.Write "</head>" &vbcrlf
		   Response.Write "<card title=""正在打开页面,请稍后……"" ontimer=""Search.asp?page=" & CurrentPage & "&amp;SearchType=" & SearchType & "&amp;ClassID=" & ClassID & "&amp;KeyWord=" & Server.URLEncode(KeyWord) & "&amp;ChannelID=" & ChannelID &"""><timer value="""&RefreshTime+10&"""/>" &vbcrlf
		   Response.Write "<p align=""center"">" &vbcrlf
		   Response.Write "本页面起用了防刷新机制，请不要在"&RefreshTime&"秒内连续刷新本页面<br/>正在打开页面，请稍后……<br/>" &vbcrlf
		   Response.Write "</p>" &vbcrlf
		   Response.Write "</card>" &vbcrlf
		   Response.Write "</wml>"
		   Response.End
		End If
		Session("SearchTime")=Now()
	End Sub

	Private Sub Class_Terminate()
        Call CloseConn()
	    Set KS=Nothing
		Set KMR=Nothing
	End Sub
	
	Sub Kesion()
        If KS.C_S(ChannelID,50)="" Then
		   Call KS.ShowError("错误信息！","对不起，还没有绑定搜索模板!")
		Else
		   F_C = KMR.LoadTemplate(KS.C_S(ChannelID,50))
		End If
		If Trim(F_C) = "" Then F_C = "模板不存在!"
		'FCls.RefreshType = "search" '设置刷新类型，以便取得当前位置导航等
		'FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
		MaxPerPage = KS.ChkClng(GetLoopNum(F_C))'循环次数
		LoopContent = KS.CutFixContent(F_C, "[loop=" & MaxPerPage & "]", "[/loop]", 0) 
		F_C = KMR.KSLabelReplaceAll(F_C) 
		Call LoadSearch()
		F_C = KS.GetEncodeConversion(F_C)
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
			Call KS.ShowError("错误信息！","你没有选择搜索类型!")
		End Select
		F_C = Replace(F_C,"{$GetSearchKey}",keyword)
		F_C = Replace(F_C,"{$ShowTotal}",Totalput)
		F_C = Replace(F_C,"{$GetMusicSearchResult}",SearchResult)
		F_C = Replace(F_C,"[loop=" & MaxPerPage & "]" & LoopContent &"[/loop]",SearchResult)
		F_C = Replace(F_C,"{$ShowPage}",KS.ShowPagePara(TotalPut, MaxPerPage, "", True, KS.C_S(ChannelID,4), CurrentPage, "SearchType=" & SearchType & "&ClassID=" & ClassID & "&KeyWord=" & Server.URLEncode(KeyWord) & "&ChannelID=" & ChannelID ))
	End Sub
	 
	Sub ArticleSearch()         
	    Dim SqlStr,Param
		Param=" Where Verific=1 And DelTF=0"
		Select Case SearchType
		    Case 100
			    If IsDate(KeyWord) Then
				   Param=Param & " And AddDate>=#" & KeyWord & " 00:00:00# and AddDate<=#" &KeyWord & " 23:59:59#"
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
			     If (FuzzySearch=1) Then
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
		    Totalput=0
			SearchResult = "没有找到任何信息!"
			Exit Sub
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
   Select Case SearchType
     Case 0,1
	     SqlStr="Select * From KS_MSSongList Where MusicName Like '%" & KeyWord & "' Order By ID Desc"
	 Case 2
	     SqlStr="Select * From KS_MSSongList Where Singer Like '%" & KeyWord & "' Order By ID Desc"
	 Case 3
	     SqlStr="Select * From KS_MSSpecial Where Name Like '%" & KeyWord & "' Order By SpecialID Desc"
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
		    SearchResult="<strong>" & KSCMUSIC.GetPlayList(RS,1,MaxPerPage,20,1,1,1)
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
    If ClassID<>"0" Then Param=Param & " And Tid In(" & KS.GetFolderTid(ClassID) & ")"
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
        On Error Resume Next 
		I=0
		Dim LC
		Select Case KS.C_S(ChannelID,6)
		    Case 1
			   Do While Not RS.Eof
			      If Not Response.IsClientConnected Then Response.End
				  LC=LoopContent
				  LC = Replace(LC,"{$GetArticleTitle}",ReplaceKeyWordRed(RS("title")))
				  LC = Replace(LC,"{$GetArticleHits}",RS("hits"))
				  LC = Replace(LC,"{$GetArticleAuthor}",RS("author"))
				  LC = Replace(LC,"{$GetArticleInput}",RS("Inputer"))
				  LC = Replace(LC,"{$GetArticleOrigin}",RS("origin"))
				  LC = Replace(LC,"{$GetArticleDate}",RS("adddate"))
				  If KS.IsNul(RS("intro")) Then
				     LC = Replace(LC,"{$GetArticleIntro}",ReplaceKeyWordRed(KS.GotTopic(Replace(Replace(Replace(KS.LoseHtml(Rs("articlecontent")), vbCrLf, ""), "[NextPage]", ""), "　", ""),50)))
				  Else
				     LC = Replace(LC,"{$GetArticleIntro}",ReplaceKeyWordRed(KS.GotTopic(Replace(Replace(KS.LoseHtml(Rs("intro")), vbCrLf, ""), "　", ""),50)))
				  End If
				  PhotoUrl=RS("PhotoUrl")
				  If PhotoUrl="" Then PhotoUrl=KS.setting(2) & KS.Setting(3) & "images/nopic.gif"
				  LC = Replace(LC,"{$GetArticlePic}",PhotoUrl)
				  LC = Replace(LC,"{$GetClassNameAndPath}",KS.GetClassNP(RS("tid")))
				  LC = Replace(LC,"{$GetArticleUrl}","../Show.asp?id="&Rs("ID")&"&amp;ChannelID="&ChannelID&"&amp;"&KS.WapValue&"")
				  LC=LFCls.ReplaceUserDefine(ChannelID,LC,RS)
				  SearchResult=SearchResult & LC
				  I = I + 1
				  If I >= MaxPerPage Then Exit Do
				  RS.MoveNext
			   Loop
			Case 2
			   Do While Not RS.Eof
			      If Not Response.IsClientConnected Then Response.End
				  LC=LoopContent
				  LC = Replace(LC,"{$GetPictureTitle}",ReplaceKeyWordRed(RS("title")))
				  LC = Replace(LC,"{$GetPictureHits}",RS("hits"))
				  LC = Replace(LC,"{$GetPictureHitsByDay}",RS("hitsbyday"))
				  LC = Replace(LC,"{$GetPictureHitsByWeek}",RS("hitsbyweek"))
				  LC = Replace(LC,"{$GetPictureHitsByMonth}",RS("hitsbymonth"))
				  LC = Replace(LC,"{$GetPictureAuthor}",RS("author"))
				  LC = Replace(LC,"{$GetPictureInput}",RS("inputer"))
				  LC = Replace(LC,"{$GetPictureOrigin}",RS("origin"))
				  LC = Replace(LC,"{$GetPictureDate}",RS("adddate"))
				  LC = Replace(LC,"{$GetPictureIntro}",ReplaceKeyWordRed(KS.Gottopic(KS.LoseHtml(RS("picturecontent")),50)))
				  LC = Replace(LC,"{$GetPhotoUrl}",RS("PhotoUrl"))
				  LC = Replace(LC,"{$GetClassNameAndPath}",KS.GetClassNP(RS("tid")))
				  LC = Replace(LC,"{$GetPictureUrl}","../Show.asp?id="&RS("ID")&"&amp;ChannelID="&ChannelID&"&amp;"&KS.WapValue&"")
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
					LC = Replace(LC,"{$GetDownTitle}",ReplaceKeyWordRed(RS("title")))
					LC = Replace(LC,"{$GetDownHits}",RS("hits"))
					LC = Replace(LC,"{$GetDownHitsByDay}",RS("hitsbyday"))
					LC = Replace(LC,"{$GetDownHitsByWeek}",RS("hitsbyweek"))
					LC = Replace(LC,"{$GetDownHitsByMonth}",RS("hitsbymonth"))
					LC = Replace(LC,"{$GetDownAuthor}",RS("author"))
					LC = Replace(LC,"{$GetDownInput}",RS("Inputer"))
					LC = Replace(LC,"{$GetDownOrigin}",RS("origin"))
					LC = Replace(LC,"{$GetDownDate}",RS("adddate"))
					LC = Replace(LC,"{$GetDownSystem}", RS("DownPT"))
					LC = Replace(LC,"{$GetDownAuthor}", RS("Author"))
					LC = Replace(LC,"{$GetDownSize}", RS("DownSize"))
					LC = Replace(LC,"{$GetDownType}", RS("DownLB"))
					LC = Replace(LC,"{$GetDownLanguage}", RS("DownYY"))
					LC = Replace(LC,"{$GetDownPower}", RS("DownSQ"))	
					LC = Replace(LC,"{$GetDownStar}", Replace(RS("Rank"),"★","<img src=""../Images/Star.gif"" border=""0"">"))		  
					LC = Replace(LC,"{$GetDownIntro}",ReplaceKeyWordRed(KS.Gottopic(KS.LoseHtml(RS("downcontent")),50)))
					PhotoUrl=RS("PhotoUrl")
					If PhotoUrl="" Then PhotoUrl=ks.setting(3) & "images/nopic.gif"
					LC = Replace(LC,"{$GetPhotoUrl}",PhotoUrl)
					LC = Replace(LC,"{$GetClassNameAndPath}",KS.GetClassNP(RS("tid")))
					LC = Replace(LC,"{$GetDownUrl}","../Show.asp?id="&RS("ID")&"&amp;ChannelID="&ChannelID&"&amp;"&KS.WapValue&"")
					LC=LFCls.ReplaceUserDefine(ChannelID,LC,RS)
					SearchResult=SearchResult & LC
					I=I+1
					If I >= MaxPerPage Then Exit Do
					RS.MoveNext
				 Loop
			  Case 4
			     Do While Not RS.Eof
				    If Not Response.IsClientConnected Then Response.end
					LC=LoopContent
					LC = Replace(LC,"{$GetFlashName}",ReplaceKeyWordRed(RS("title")))
					LC = Replace(LC,"{$GetFlashHits}",RS("hits"))
					LC = Replace(LC,"{$GetFlashHitsByDay}",RS("hitsbyday"))
					LC = Replace(LC,"{$GetFlashHitsByWeek}",RS("hitsbyweek"))
					LC = Replace(LC,"{$GetFlashHitsByMonth}",RS("hitsbymonth"))
					LC = Replace(LC,"{$GetFlashAuthor}",RS("author"))
					LC = Replace(LC,"{$GetFlashInput}",RS("Inputer"))
					LC = Replace(LC,"{$GetFlashOrigin}",RS("origin"))
					LC = Replace(LC,"{$GetFlashDate}",RS("adddate"))
					LC = Replace(LC,"{$GetFlashIntro}",ReplaceKeyWordRed(KS.Gottopic(KS.LoseHtml(KS.HTMLCode(RS("flashcontent"))),200)))
					PhotoUrl=RS("PhotoUrl")
					if PhotoUrl="" then PhotoUrl=ks.setting(3) & "images/nopic.gif"
					LC = Replace(LC,"{$GetPhotoUrl}",PhotoUrl)
					LC = Replace(LC,"{$GetClassNameAndPath}",KS.GetClassNP(RS("tid")))
					LC = Replace(LC,"{$GetFlashUrl}","../Show.asp?id="&RS("ID")&"&amp;ChannelID="&ChannelID&"&amp;"&KS.WapValue&"")
					LC=LFCls.ReplaceUserDefine(ChannelID,LC,RS)
					SearchResult=SearchResult & LC
					I=I+1
					If I >= MaxPerPage Then Exit Do
					RS.MoveNext
				 Loop
			  Case 5
			     Do While Not RS.Eof
				    If Not Response.IsClientConnected Then Response.end
					LC=LoopContent
					LC = Replace(LC,"{$GetProductName}",ReplaceKeyWordRed(RS("title")))
					LC = Replace(LC,"{$GetProductIntro}", ReplaceKeyWordRed(KS.LoseHtml(KS.HtmlCode(RS("ProIntro")))))
					LC = Replace(LC,"{$GetProductID}", RS("ProID"))
					LC = Replace(LC,"{$GetProductModel}", RS("ProModel"))
					LC = Replace(LC,"{$GetProductSpecificat}", RS("ProSpecificat"))
					LC = Replace(LC,"{$GetProducerName}", RS("ProducerName"))
					LC = Replace(LC,"{$GetTrademarkName}", RS("TrademarkName"))
					LC = Replace(LC,"{$GetServiceTerm}", RS("ServiceTerm"))
					LC = Replace(LC,"{$GetRank}",Replace(RS("Rank"),"★","<img src=""../Images/Star.gif"" border=""0"">"))
					LC = Replace(LC,"{$GetTotalNum}",RS("TotalNum"))
					LC = Replace(LC,"{$GetProductUnit}", RS("Unit"))
					LC = Replace(LC,"{$GetProductHits}", RS("hits"))
					LC = Replace(LC,"{$GetProductDate}", RS("AddDate"))
					LC = Replace(LC,"{$GetPrice_Market}", RS("Price_Market"))
					LC = Replace(LC,"{$GetPrice}", RS("Price"))
					LC = Replace(LC,"{$GetPrice_Member}", RS("Price_Member"))
					LC = Replace(LC,"{$GetPrice_Original}", RS("Price_Original"))
					If RS("ProductType")=3 Then
					   LC = Replace(LC,"{$GetDiscount}", RS("Discount"))
					Else
					   LC = Replace(LC,"{$GetDiscount}", "")
					End If
					LC = Replace(LC,"{$GetScore}", RS("Point"))
					PhotoUrl=RS("PhotoUrl")
					If PhotoUrl="" Then PhotoUrl="../images/nopic.gif"
					LC = Replace(LC,"{$GetPhotoUrl}",PhotoUrl)
					LC = Replace(LC,"{$GetClassNameAndPath}",KS.GetClassNP(RS("tid")))
					LC = Replace(LC,"{$GetProductUrl}","../Show.asp?id="&RS("ID")&"&amp;ChannelID="&ChannelID&"&amp;"&KS.WapValue&"")
					LC=LFCls.ReplaceUserDefine(ChannelID,LC,RS)
					SearchResult=SearchResult & LC
					I=I+1
					If I >= MaxPerPage Then Exit Do
					RS.MoveNext
				 Loop
			  Case 7
			     Do While Not RS.Eof
				    If Not Response.IsClientConnected Then Response.end
					LC=LoopContent
					LC = Replace(LC,"{$GetMovieName}",ReplaceKeyWordRed(RS("title")))
					LC = Replace(LC,"{$GetMovieIntro}",ReplaceKeyWordRed(KS.LoseHtml(RS("MovieContent"))))
					LC = Replace(LC,"{$GetMovieID}", RS("ID"))
					LC = Replace(LC,"{$GetMovieActor}", RS("MovieAct"))
					LC = Replace(LC,"{$GetMovieDirector}", RS("MovieDY"))
					LC = Replace(LC,"{$GetMovieTime}", RS("MovieTime"))
					LC = Replace(LC,"{$GetScreenTime}", RS("ScreenTime"))
					LC = Replace(LC,"{$GetMovieStar}",Replace(RS("Rank"),"★","<img src=""../Images/Star.gif"" border=""0"">"))
					LC = Replace(LC,"{$GetMovieArea}",RS("MovieDQ"))
					LC = Replace(LC,"{$GetMovieLanguage}",RS("MovieYY"))
					LC = Replace(LC,"{$GetMovieHits}", RS("hits"))
					LC = Replace(LC,"{$GetMovieDate}", RS("AddDate"))
					PhotoUrl=RS("PhotoUrl")
					If PhotoUrl="" Then PhotoUrl="../images/nopic.gif"
					LC = Replace(LC,"{$GetPhotoUrl}",PhotoUrl)
					LC = Replace(LC,"{$GetClassNameAndPath}",KS.GetClassNP(RS("tid")))
					LC = Replace(LC,"{$GetMovieUrl}","../Show.asp?id="&RS("ID")&"&amp;ChannelID="&ChannelID&"&amp;"&KS.WapValue&"")
					LC=LFCls.ReplaceUserDefine(ChannelID,LC,RS)
					SearchResult=SearchResult & LC
					I=I+1
					If I >= MaxPerPage Then Exit Do
					RS.MoveNext
				 Loop	
			  Case 8
			     Do While Not RS.Eof
				    If Not Response.IsClientConnected Then Response.end
					LC=LoopContent
					LC = Replace(LC,"{$GetGQTitle}",ReplaceKeyWordRed(RS("title")))
					LC = Replace(LC,"{$GetGQIntro}",ReplaceKeyWordRed(KS.LoseHtml(KS.HTMLCode(RS("GQContent")))))	
					LC = Replace(LC,"{$GetGQInfoID}", RS("ID"))
					LC = Replace(LC,"{$GetPrice}", RS("Price"))
					LC = Replace(LC,"{$GetTransType}", KS.GetGQTypeName(RS("TypeID")))
					LC = Replace(LC,"{$GetInfoType}", KS.C_C(RS("tid"),1))
					LC = Replace(LC,"{$GetValidTime}", KMR.GetValidTime(RS("ValidDate")))
					LC = Replace(LC,"{$GetGQContent}", RS("GQContent"))
					LC = Replace(LC,"{$GetAddDate}",RS("AddDate"))
					LC = Replace(LC,"{$GetGQHits}", RS("hits"))
					LC = Replace(LC,"{$GetGQDate}", RS("AddDate"))
					LC = Replace(LC,"{$GetInput}", RS("UserName"))
					LC = Replace(LC,"{$GetCompanyName}", RS("CompanyName"))
					LC = Replace(LC,"{$GetContactMan}", RS("ContactMan"))
					LC = Replace(LC,"{$GetContactTel}", RS("Tel"))
					LC = Replace(LC,"{$GetFax}", RS("Fax"))
					LC = Replace(LC,"{$GetAddress}", RS("Address"))
					LC = Replace(LC,"{$GetEmail}", RS("Email"))
					LC = Replace(LC,"{$GetPostCode}", RS("zip"))
					LC = Replace(LC,"{$GetProvince}", RS("Province"))
					LC = Replace(LC,"{$GetCity}", RS("City"))
					PhotoUrl=RS("PhotoUrl")
					If PhotoUrl="" Then PhotoUrl="../images/nopic.gif"
					LC = Replace(LC,"{$GetPhotoUrl}",PhotoUrl)
					LC = Replace(LC,"{$GetClassNameAndPath}",KS.GetClassNP(RS("tid")))
					LC = Replace(LC,"{$GetGQInfoUrl}","../Show.asp?id="&RS("ID")&"&amp;ChannelID="&ChannelID&"&amp;"&KS.WapValue&"")
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
		    Content=Replace(Content,KeyWordArr(i),"<b>" &KeyWordArr(i) & "</b>")
		Next
		ReplaceKeyWordRed=Content
	End Function
	
	'返回循环次数
	Function GetLoopNum(Content)
	    Dim regEx, Matches, Match
		Set regEx = New RegExp
		regEx.Pattern="\[loop=\d*]"
		regEx.IgnoreCase = True
		regEx.Global = True
		Set Matches = regEx.Execute(Content)
		If Matches.count > 0 Then
		   GetLoopNum=Replace(Replace(Matches.item(0),"[loop=",""),"]","")
		Else
		   GetLoopNum=0
		End If
	End Function

End Class
%> 