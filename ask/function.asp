<%
Dim ACls
Set ACls = New AskCls
Call ACls.run()
Class AskCls
        Private KS
		Private Sub Class_Initialize()
		 If (Not Response.IsClientConnected)Then
			Response.Clear
			Response.End
		 End If
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set ACls=Nothing
		End Sub
		
		Sub Run()
		 Call LoadCategoryList()
		End Sub

		
		
			
	Public Sub LoadCategoryList()
	  If Not IsObject(Application(KS.SiteSN&"_askclasslist")) Then
		Dim Rs,SQL,TempXmlDoc
		Set Rs = Conn.Execute("SELECT classid,ClassName,Readme,rootid,depth,parentid,Parentstr,child FROM KS_AskClass ORDER BY orders,classid")
		If Not (Rs.BOF And Rs.EOF) Then
			SQL=Rs.GetRows(-1)
			Set TempXmlDoc = KS.ArrayToxml(SQL,Rs,"row","classlist")
		End If
		Rs.Close
		Set Rs = Nothing
		If IsObject(TempXmlDoc) Then
			Application.Lock
				Set Application(KS.SiteSN&"_askclasslist") = TempXmlDoc
			Application.unLock
		End If
	 End If
	End Sub
	Public Function IndexMenulist()
		Dim Parentlist,Node,strTempMenu
		 If IsObject(Application(KS.SiteSN&"_askclasslist")) Then
			Set Parentlist = Application(KS.SiteSN&"_askclasslist")
			If Not Parentlist Is Nothing Then
				Dim classid,ClassName,Childs,i,depth,strLinks,rootid
				Childs = Parentlist.documentElement.SelectNodes("row").Length
				i = 0
				For Each Node in Parentlist.documentElement.SelectNodes("row[@depth=0]")
					ClassName = Node.selectSingleNode("@classname").text
					classid = Node.selectSingleNode("@classid").text
					depth = Node.selectSingleNode("@depth").text
					rootid = Node.selectSingleNode("@rootid").text
					If KS.ASetting(16)="1" Then
					strLinks = "<a href=""" & KS.Setting(3) & KS.ASetting(1) & "list-" & classid & KS.ASetting(17) & """>"
					Else
					strLinks = "<a href=""" & KS.Setting(3) & KS.ASetting(1) & "showlist.asp?id=" & classid & """>"
					End If
					strLinks = strLinks & ClassName
					strLinks = strLinks & "</a> "

					strTempMenu = strTempMenu & "<dt>" & strLinks & "</dt>" & vbCrLf
					strTempMenu = strTempMenu & GetChildList(classid,4)
				Next
				Set Parentlist = Nothing
			End If
		End If
		IndexMenulist = strTempMenu
	End Function
	Public Function GetChildList(cid,m)
		Dim Childlist,Node,strTemp,i,ParentLinks
		Dim classid,ClassName,strLinks
		If IsObject(Application(KS.SiteSN&"_askclasslist")) Then
			Set Childlist = Application(KS.SiteSN&"_askclasslist")
			If Not Childlist Is Nothing Then
				i = 0
				strTemp = "<dd>"
				For Each Node in Childlist.documentElement.SelectNodes("row[@parentid="&cid&"]")
					i = i + 1
					ClassName = Node.selectSingleNode("@classname").text
					classid = Node.selectSingleNode("@classid").text
					   If KS.ASetting(16)="1" Then
						strLinks = "<a href=""" & KS.Setting(3) & KS.ASetting(1) & "list-" & classid & KS.ASetting(17) & """>"
					   Else
						strLinks = "<a href=""" & KS.Setting(3) & KS.ASetting(1) & "showlist.asp?id=" & classid & """>"
					   End If
						strLinks = strLinks & ClassName & "</a> "
						strTemp = strTemp & strLinks
					If i mod m=0 Then strTemp =strTemp & "<br />"
				Next
				Set Childlist = Nothing
				strTemp = strTemp & ParentLinks & "</dd>" & vbCrLf
			End If
			Set Childlist = Nothing
		End If
		GetChildList = strTemp
	End Function
	
	Function ReturnAskConfig(sTokenName)
	    select case lcase(sTokenName)
		   case "sitetitle" ReturnAskConfig=KS.ASetting(2)
		   case "menulist"  ReturnAskConfig=IndexMenulist
		   case "resolvednum" ReturnAskConfig=conn.execute("select count(topicid) from KS_asktopic where topicmode=1")(0)
		   case "unresolvednum" ReturnAskConfig=conn.execute("select count(topicid) from KS_asktopic where topicmode=0")(0)
		end select
	End Function
				
End Class
%>