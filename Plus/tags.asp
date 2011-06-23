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
Dim TCls
Set TCls = New Tags
TCls.Kesion()
Set TCls = Nothing
Const MaxPerPage=20   '每页显示条数
Const MaxTags=500     '默认显示tags个数

Class Tags
    Private KS,KMR,F_C,LoopContent,SearchResult,photourl
	Private ChannelID,ClassID,SearchType,TagsName,SearchForm
    Private I,TotalPut, CurrentPage,RS
   
	Private Sub Class_Initialize()
		Set KS=New PublicCls
		Set KMR=New Refresh
		If KS.S("page") <> "" Then
          CurrentPage = CInt(Request("page"))
        Else
          CurrentPage = 1
        End If
	End Sub

	Private Sub Class_Terminate()
        closeconn
	    Set KS=Nothing
		Set KMR=Nothing
	End Sub
  
 Sub Kesion()
		   F_C = KMR.LoadTemplate(KS.Setting(120))
		   If Trim(F_C) = "" Then F_C = "模板不存在!"
		   
		   FCls.RefreshType = "tags" '设置刷新类型，以便取得当前位置导航等
		   FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
		   
			TagsName=KS.CheckXSS(KS.S("n"))
			If TagsName="" Then 
			 Call TagsMain()
			Else
			 Call TagsList()
			End If
		   
			F_C = KMR.KSLabelReplaceAll(F_C) 
			Call TagsHits()
			Response.Write F_C
 End Sub
 

 
 Sub TagsMain()
   Dim TP:Tp=LFCls.GetConfigFromXML("tags","/labeltemplate/label","tags")
   Dim RS,SQL,K,str
   If InStr(tp,"{$ShowHotTags}")<>0 Then
	   Set RS=Conn.Execute("Select top " & MaxTags & " KeyText,hits From KS_KeyWords order by hits desc,id desc")
	   If Not RS.Eof Then SQL=RS.GetRows(-1)
	   RS.Close:Set RS=Nothing
	   If IsArray(SQL) Then
		 For k=0 to Ubound(SQL,2)
		  str=str & "<a href='?n=" & server.urlencode(SQL(0,K)) & "' title='已被使用了" & SQL(1,K) & "次'>" & SQL(0,K) & "</a>  "
		 Next
	   End If
	   Tp=Replace(Tp,"{$ShowHotTags}",str)
   End If
   
    If InStr(tp,"{$ShowNewTags}")<>0 Then
	   str=""
	   Set RS=Conn.Execute("Select top " & MaxTags & " KeyText,hits From KS_KeyWords order by adddate desc")
	   If Not RS.Eof Then SQL=RS.GetRows(-1)
	   RS.Close:Set RS=Nothing
	   If IsArray(SQL) Then
		 For k=0 to Ubound(SQL,2)
		  str=str & "<a href='?n=" & server.urlencode(SQL(0,K)) & "' title='已被使用了" & SQL(1,K) & "次'>" & SQL(0,K) & "</a>  "
		 Next
	   End If
	   Tp=Replace(Tp,"{$ShowNewTags}",str)
	End If
   
    F_C=Replace(F_C,"{$ShowTags}",Tp)
	F_C=Replace(F_C,"{$TagsName}","")
 End Sub
 

 Sub TagsList()
    SearchTags()
	TagsHits()
	F_C=Replace(F_C,"{$ShowTags}",SearchResult)
	F_C=Replace(F_C,"{$PageStr}","<div style='text-align:center'>" &  ShowPagePara(totalput, MaxPerPage, "", true,"条记录", CurrentPage,KS.QueryParam("page,submit")) & "</div>")
	F_C = Replace(F_C,"{$TagsName}",TagsName)
	F_C = Replace(F_C,"{$ShowTotal}",totalput)
  End Sub
  
  Sub TagsHits()
    If TagsName<>"" Then
	 Conn.Execute("Update KS_KeyWords set hits=hits+1,lastusetime=" & SqlNowString & " where keytext='" & TagsName & "'")
	End IF
  End Sub
  
  Sub SearchTags() 
     Dim SqlStr,Param,SQL,K
     Dim RSC:Set RSC=conn.execute("select ChannelID,ChannelTable From KS_Channel Where ChannelID<>6 And ChannelID<>8 And ChannelID<>9 and ChannelID<>10 and ChannelStatus=1 order by channelid")
	 SQL=RSC.GetRows(-1):RSC.Close:Set RSC=Nothing
	 For K=0 To Ubound(SQL,2)
		 If SqlStr<>"" Then SqlStr=SqlStr & " Union All "
								 
		 Select Case  KS.C_S(SQL(0,K),6)
		  Case 1
			SqlStr=SqlStr & "select ID,Title,Tid,ReadPoint,InfoPurview,Fname,Changes,AddDate,Popular," & SQL(0,K) & " as ChannelID,hits,Inputer As username From " & SQL(1,K)
		  case 2
			 SqlStr=SqlStr & "select ID,Title,Tid,ReadPoint,InfoPurview,Fname,0 as Changes,AddDate,Popular," & SQL(0,K) & " as ChannelID,Hits,Inputer As username From " & SQL(1,K)
		  case 3
			 SqlStr=SqlStr & "select ID,Title,Tid,ReadPoint,InfoPurview,Fname,0 as Changes,AddDate,Popular," & SQL(0,K) & " as ChannelID,Hits,Inputer As username From " & SQL(1,K)
		  case 4
			 SqlStr=SqlStr & "select ID,Title,Tid,ReadPoint,InfoPurview,Fname,0 as Changes,AddDate,Popular," & SQL(0,K) & " as ChannelID,Hits,Inputer As username From " & SQL(1,K)
		  Case 5
		  SqlStr=SqlStr & "select ID,Title,Tid,0 as ReadPoint,0 as InfoPurview,Fname,0 as Changes,AddDate,Popular," & SQL(0,K) & " as ChannelID,Hits,Inputer As username From " & SQL(1,K)
		  Case 7
		  SqlStr=SqlStr & "select ID,Title,Tid,0 as ReadPoint,0 as InfoPurview,Fname,0 as Changes,AddDate,Popular," & SQL(0,K) & " as ChannelID,Hits,Inputer As username From " & SQL(1,K)
		 End Select
		SqlStr=SqlStr & " Where DelTF=0 And Verific=1 And keywords like '%" & TagsName & "%'"
	Next
	SqlStr="Select ID,Title,Tid,ReadPoint,InfoPurview,Fname,Changes,AddDate,Popular,ChannelID,hits,username From (" & SQLStr & ") a  ORDER BY ADDDATE DESC,ID Asc"
  

  Set RS=Server.CreateObject("ADODB.RECORDSET")
  RS.Open SqlStr,Conn,1,1

  IF RS.Eof And RS.Bof Then
      totalput=0
      SearchResult = "Tags:<Font color=red>" & TagsName & "</font>,没有找到任何相关信息!"
	  exit sub
  Else
					TotalPut= RS.Recordcount
                    If CurrentPage < 1 Then CurrentPage = 1

                    If (CurrentPage - 1) * MaxPerPage > totalPut Then
                        If (TotalPut Mod MaxPerPage) = 0 Then
                            CurrentPage = totalPut \ MaxPerPage
                        Else
                            CurrentPage = totalPut \ MaxPerPage + 1
                        End If
                    End If

                    If CurrentPage > 1 and (CurrentPage - 1) * MaxPerPage < totalPut Then
                            RS.Move (CurrentPage - 1) * MaxPerPage
                    Else
                            CurrentPage = 1
                    End If
					Call GetSearchResult
    End IF
	RS.Close
	Set RS=Nothing
  End Sub   

  
  
  Sub GetSearchResult()
      On Error Resume Next
      Dim LP:LP=LFCls.GetConfigFromXML("tags","/labeltemplate/label","listtp")  
      LoopContent=KS.CutFixContent(LP, "[loop]", "[/loop]", 0)     
	  I=0
       Dim LC
		  Do While Not RS.Eof
			  LC=LoopContent
			  LC=replace(LC,"{$Title}",rs(1))
			  If IsNull(rs(11)) or rs(11)="" Then
			  LC=replace(LC,"{$UserName}","-")
			  Else
			  LC=replace(LC,"{$UserName}",rs(11))
			  End If
			  LC=replace(LC,"{$Hits}",rs(10))
			  LC=replace(LC,"{$AddDate}",formatdatetime(rs(7),2))
			  
			  LC=replace(LC,"{$ClassName}",KS.GetClassNP(rs(2)))
			  LC=replace(LC,"{$Url}",KS.GetItemUrl(rs(9),rs(2),rs(0),rs(5)))
			  SearchResult=SearchResult & LC
			I = I + 1
			If I >= MaxPerPage Then Exit Do
			RS.MoveNext
		Loop
	  LP=Replace(LP,KS.CutFixContent(LP, "[loop]", "[/loop]", 1),SearchResult)
	  SearchResult=LP    
  End Sub  
  
Public Function ShowPagePara(totalnumber, MaxPerPage, FileName, ShowAllPages, strUnit, CurrentPage, ParamterStr)
		  Dim N, I, PageStr
				Const Btn_First = "<span style='font-family:webdings;font-size:14px' title='第一页'>9</span>" '定义第一页按钮显示样式
				Const Btn_Prev = "<span style='font-family:webdings;font-size:14px' title='上一页'>3</span>" '定义前一页按钮显示样式
				Const Btn_Next = "<span style='font-family:webdings;font-size:14px' title='下一页'>4</span>" '定义下一页按钮显示样式
				Const Btn_Last = "<span style='font-family:webdings;font-size:14px' title='最后一页'>:</span>" '定义最后一页按钮显示样式
				  PageStr = ""
					If totalnumber Mod MaxPerPage = 0 Then
						N = totalnumber \ MaxPerPage
					Else
						N = totalnumber \ MaxPerPage + 1
					End If
					PageStr = PageStr & ("<div class='showpage' style='height:20px'><form action=""" & FileName & "?" & ParamterStr & """ name=""myform"" method=""post"">共 <font color=red>" & totalnumber & "</font> " & strUnit & "  分 <font color=red>" & N & "</font> 页 每页 <font color=red>" & MaxPerPage &"</font> " & strUnit &" | 当前第 <font color=red>" & CurrentPage & "</font> 页 &nbsp;&nbsp;&nbsp;")
					If CurrentPage < 2 Then
						PageStr = PageStr & Btn_First & " " & Btn_Prev & " "
					Else
						PageStr = PageStr & ("<a href=" & FileName & "?page=1" & "&" & ParamterStr & ">" & Btn_First & "</a> <a href=" & FileName & "?page=" & CurrentPage - 1 & "&" & ParamterStr & ">" & Btn_Prev & "</a> ")
					End If
					
					If N - CurrentPage < 1 Then
						PageStr = PageStr & " " & Btn_Next & " " & Btn_Last & " "
					Else
						PageStr = PageStr & (" <a href=" & FileName & "?page=" & (CurrentPage + 1) & "&" & ParamterStr & ">" & Btn_Next & "</a> <a href=" & FileName & "?page=" & N & "&" & ParamterStr & ">" & Btn_Last & "</a> ")
					End If
					If ShowAllPages = True Then
						PageStr = PageStr & ("转到:<input type='text' value='" & (CurrentPage + 1) &"' name='page' style='width:30px;height:18px;text-align:center;'>&nbsp;<input style='height:18px;border:1px #a7a7a7 solid;background:#fff;' type='submit' value='GO' name='sb'>")
				  End If
				  PageStr = PageStr & "</form></div>"
			 ShowPagePara = PageStr
	End Function
End Class
%> 