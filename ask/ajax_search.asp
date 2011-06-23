<!--#include file="../conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="function.asp"-->
<!--#include file="../KS_Cls/Kesion.KeyCls.asp"-->
<!--#include file="../KS_Cls/Kesion.EscapeCls.asp"-->
<%
Response.ContentType="text/html"
Response.Expires = -9999
Response.AddHeader "pragma", "no-cache"
Response.AddHeader "cache-ctrol", "no-cache"
Response.CharSet="gb2312"
%>
<h3><span><img src="images/icon_vote.gif" align="middle" />与您的提问相关的已解决问题（看看是否能解答您的疑问）</span></h3>
<ul>
<%
Dim KS:Set KS=New PublicCls
Dim XMLDom
Dim searchmode,Keyword
searchmode = 0
Keyword = KS.S("q")
If Len(Keyword) < 2 Then Keyword = KS.S("word")
If Len(Keyword) < 2 Then Keyword = ""
If Keyword = "请输入关键字" Then Keyword = ""
KeyWord=UnEscape(keyword)
showmain()
CloseConn()
Set KS=Nothing

Sub showmain()
	Dim Rs,SQL,FoundSQL,topiclist,node
	Dim FindSolved,SolvedQuestion
	Dim KeywordArray,KeywordLike,i,n
	FindSolved = 0
	Set XMLDom = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	XMLDom.appendChild(XMLDom.createElement("xml"))
	FoundSQL = ""
	If searchmode = 0 Then
			 Dim WS:Set WS=New Wordsegment_Cls
		     KeywordArray = Split(WS.SplitKey(Keyword,4,0),",")
			 Set WS=Nothing
		n = 0
		For i = 0 To UBound(KeywordArray)
			If Len(KeywordArray(i)) > 1 Then
				If n = 0 Then
					If DataBaseType=1 Then
						FoundSQL = "title like '%"&KeywordArray(i)&"%'"
					Else
						FoundSQL = "InStr(1,LCase(Title),LCase('"&KeywordArray(i)&"'),0)>0"
					End If
				Else
					If DataBaseType=1 Then
						FoundSQL = FoundSQL & " Or title like '%"&KeywordArray(i)&"%'"
					Else
						FoundSQL = FoundSQL & " Or InStr(1,LCase(Title),LCase('"&KeywordArray(i)&"'),0)>0"
					End If
				End If
				n = n + 1
			End If
		Next
		If n = 0 Then
			FoundSQL = ""
		End If
	Else
		If ws.CheckKeyword(Keyword) Then
			If DataBaseType=1 Then
				FoundSQL = "title like '%"&Keyword&"%'"
			Else
				FoundSQL = "InStr(1,LCase(Title),LCase('"&Keyword&"'),0)>0"
			End If
		Else
			FoundSQL = ""
		End If
	End If
	If Len(FoundSQL) > 10 Then
		FoundSQL = "And ("&FoundSQL&")"
		SQL="SELECT TOP 10 TopicID,classid,UserName,classname,title,Expired,Closed,PostTable,DateAndTime,LastPostTime,LockTopic,Reward,TopicMode FROM KS_AskTopic WHERE TopicMode=1 And LockTopic=0 " & FoundSQL & " ORDER BY LastPostTime DESC"
		Set Rs = Conn.Execute(SQL)
		If Not Rs.EOF Then
			SQL=Rs.GetRows(-1)
			Set topiclist=KS.ArrayToxml(SQL,Rs,"row","topic")
			FindSolved = 1
		Else
			Set topiclist=Nothing
			FindSolved = 0
		End If
		Rs.Close : Set Rs=Nothing
		SQL=Empty
		SolvedQuestion = ""
		If Not topiclist Is Nothing Then
			For Each Node in topiclist.documentElement.SelectNodes("row")
				SolvedQuestion = SolvedQuestion & "<li>"
				If KS.ASetting(16)="1" Then
				SolvedQuestion = SolvedQuestion & "<a href=""show-" & Node.selectSingleNode("@topicid").text & KS.ASetting(17) & """ target=""_blank"">" & KS.HTMLEncode(Node.selectSingleNode("@title").text) & "</a> - "
				Else
				SolvedQuestion = SolvedQuestion & "<a href=""q.asp?id=" & Node.selectSingleNode("@topicid").text & """ target=""_blank"">" & KS.HTMLEncode(Node.selectSingleNode("@title").text) & "</a> - "
				End If
				SolvedQuestion = SolvedQuestion & Node.selectSingleNode("@dateandtime").text
				SolvedQuestion = SolvedQuestion & "</li>" & vbCrLf
			Next
		End If
		Set topiclist=Nothing
		Set XMLDom=Nothing
	Else
		FindSolved = 0
	End If
	KS.Echo Escape(SolvedQuestion)
End Sub
%>
</ul>