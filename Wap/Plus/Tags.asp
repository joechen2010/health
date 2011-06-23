<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
Option Explicit
Response.ContentType="text/vnd.wap.wml"
Response.Charset="utf-8"
Response.Write "<?xml version=""1.0"" encoding=""utf-8""?>" &vbcrlf
Response.Write "<!DOCTYPE wml PUBLIC ""-//WAPFORUM//DTD WML 1.1//EN"" ""http://www.wapforum.org/DTD/wml_1.1.xml"">" &vbcrlf
%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<%
Dim TCls
Set TCls = New Tags
TCls.Kesion()
Set TCls = Nothing
Const MaxPerPage=20   '每页显示条数
Const MaxTags=500     '默认显示Tags个数

Class Tags
    Private KS,KMR,F_C,LoopContent,SearchResult,PhotoUrl
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
        CloseConn
	    Set KS=Nothing
		Set KMR=Nothing
	End Sub
	
	Sub Kesion()
	    F_C = KMR.LoadTemplate(KS.WSetting(10))
		If Trim(F_C) = "" Then F_C = "模板不存在!"
		'FCls.RefreshType = "Tags" '设置刷新类型，以便取得当前位置导航等
		'FCls.RefreshFolderID = "0" '设置当前刷新目录ID 为"0" 以取得通用标签
		TagsName=KS.S("n")
		Call TagsList()
		F_C = KMR.KSLabelReplaceAll(F_C) 
		Call TagsHits()
		F_C = KS.GetEncodeConversion(F_C)
		Response.Write F_C
	End Sub
	
	Sub TagsList()
	    SearchTags()
		TagsHits()
		F_C = Replace(F_C,KS.CutFixContent(F_C, "[loop]", "[/loop]", 1),SearchResult)
		F_C = Replace(F_C,"{$PageStr}",KS.ShowPagePara(TotalPut, MaxPerPage, "Tags.asp", True, "条记录", CurrentPage, KS.QueryParam("page")))
		F_C = Replace(F_C,"{$TagsName}",TagsName)'标签名
		F_C = Replace(F_C,"{$ShowTotal}",totalput)'标签数量
    End Sub
	
	Sub TagsHits()
	    If TagsName<>"" Then
		Conn.Execute("Update KS_KeyWords set hits=hits+1,Lastusetime=" & SqlNowString & " where KeyText='" & TagsName & "'")
		End IF
	End Sub
	
	Sub SearchTags() 
        Dim SqlStr,Param,SQL,K
		Dim RSC:Set RSC=Conn.Execute("select ChannelID,ChannelTable From KS_Channel Where ChannelID<>6 And ChannelID<>8 And ChannelID<>9 and ChannelID<>10 and ChannelStatus=1 order by ChannelID")
		SQL=RSC.GetRows(-1):RSC.Close:Set RSC=Nothing
		For K=0 To Ubound(SQL,2)
		    If SqlStr<>"" Then SqlStr=SqlStr & " Union All "
			Select Case  KS.C_S(SQL(0,K),6)
			    Case 1
				SqlStr=SqlStr & "select ID,Title,Tid,ReadPoint,InfoPurview,Fname,Changes,AddDate,Popular," & SQL(0,K) & " as ChannelID,hits,Inputer As username From " & SQL(1,K)
				Case 2
				SqlStr=SqlStr & "select ID,Title,Tid,ReadPoint,InfoPurview,Fname,0 as Changes,AddDate,Popular," & SQL(0,K) & " as ChannelID,Hits,Inputer As username From " & SQL(1,K)
				Case 3
				SqlStr=SqlStr & "select ID,Title,Tid,ReadPoint,InfoPurview,Fname,0 as Changes,AddDate,Popular," & SQL(0,K) & " as ChannelID,Hits,Inputer As username From " & SQL(1,K)
				Case 4
				SqlStr=SqlStr & "select ID,Title,Tid,ReadPoint,InfoPurview,Fname,0 as Changes,AddDate,Popular," & SQL(0,K) & " as ChannelID,Hits,Inputer As username From " & SQL(1,K)
				Case 5
				SqlStr=SqlStr & "select ID,Title,Tid,0 as ReadPoint,0 as InfoPurview,Fname,0 as Changes,AddDate,Popular," & SQL(0,K) & " as ChannelID,Hits,Inputer As username From " & SQL(1,K)
				Case 7
				SqlStr=SqlStr & "select ID,Title,Tid,0 as ReadPoint,0 as InfoPurview,Fname,0 as Changes,AddDate,Popular," & SQL(0,K) & " as ChannelID,Hits,Inputer As username From " & SQL(1,K)
			End Select
			SqlStr=SqlStr & " Where DelTF=0 And Verific=1 And keywords like '%" & TagsName & "%'"
		Next
		SqlStr="Select ID,Title,Tid,ReadPoint,InfoPurview,Fname,Changes,AddDate,Popular,ChannelID,hits,username From (" & SQLStr & ")  Order By ID Desc"
		Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open SqlStr,Conn,1,3
		IF RS.Eof And RS.Bof Then
		   Totalput=0
		   SearchResult = "Tags:<b>" & TagsName & "</b>,没有找到任何相关信息!<br/>"
		   Exit Sub
		Else
		   TotalPut= RS.Recordcount
		   If CurrentPage < 1 Then CurrentPage = 1
		   If (CurrentPage - 1) * MaxPerPage > TotalPut Then
		      If (TotalPut Mod MaxPerPage) = 0 Then
			     CurrentPage = TotalPut \ MaxPerPage
			  Else
			     CurrentPage = TotalPut \ MaxPerPage + 1
			  End If
		   End If
		   If CurrentPage > 1 and (CurrentPage - 1) * MaxPerPage < TotalPut Then
		      RS.Move (CurrentPage - 1) * MaxPerPage
		   Else
		      CurrentPage = 1
		   End If
		   Call GetSearchResult
		End IF
		RS.Close:Set RS=Nothing
    End Sub   
	
	Sub GetSearchResult()
	    On Error Resume Next
		LoopContent=KS.CutFixContent(F_C, "[loop]", "[/loop]", 0)
		I=0
		Dim LC
		Do While Not RS.EOF
		   LC=LoopContent
		   LC=replace(LC,"{$Title}",RS(1))
		   LC=replace(LC,"{$UserName}",RS(11))
		   LC=replace(LC,"{$Hits}",RS(10))
		   LC=replace(LC,"{$AddDate}",formatdatetime(RS(7),2))
		   LC=replace(LC,"{$ClassName}",KS.GetClassNP(RS(2)))
		   LC=replace(LC,"{$Url}","../Show.asp?ID="&RS(0)&"&amp;ChannelID="&RS(9)&"&amp;"&KS.WapValue&"")
		   SearchResult=SearchResult & LC
		   I = I + 1
		   If I >= MaxPerPage Then Exit Do
		   RS.MoveNext
		Loop
	End Sub
End Class
%> 