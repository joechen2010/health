<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Response.Buffer=true
Dim KSCls
Set KSCls = New RefreshHtmlSave
KSCls.Kesion()
Set KSCls = Nothing

Class RefreshHtmlSave
        Private KS,KSRObj
		Private RefreshFlag,f
		Private ReturnInfo,FsoHtmlList
		Private StartRefreshTime
		Private ChannelID,ItemName,Table
		Private Types
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSRObj=New Refresh
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KSRObj=Nothing
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
			With KS
			Server.ScriptTimeOut=9999999
			Types = Request("Types")             'Content ��������ҳ���� Folder ������Ŀ����
			RefreshFlag = Request("RefreshFlag") 'ȡ���ǰ���������ˢ��,��Newֻ�������µ�ָ��ƪ������
			ChannelID = Request("ChannelID")     '��Ƶ������
			FCls.ChannelID=ChannelID
			
			If RefreshFlag<>"IDS" Then
				If Not KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "20005") Then                '���Ȩ��
					  Call KS.ReturnErr(1, "")
					  response.End()
				End If
			End If
			
			
			f=Request("f")
			If KS.S("FsoType")="1" Then
			FCls.FsoListNum=0
			ElseIf KS.S("FsoType")="2" Then
			FCls.FsoListNum=KS.ChkCLng(KS.S("FsoListNum"))
			Else
			FCls.FsoListNum=KS.ChkClng(KS.C_S(ChannelID,35))
			End If
			if f="task" then 
			   FCls.FsoListNum=3
			end if



			FCls.ItemUnit = KS.C_S(ChannelID,4)
			
	
			'ˢ��ʱ��
			StartRefreshTime = Request("StartRefreshTime")
			If StartRefreshTime = "" Then StartRefreshTime = Timer()
			Table=KS.C_S(ChannelID,2)
			ItemName=KS.C_S(ChannelID,3)
			Select Case Types
			 Case "Content"
			            If KS.C_S(ChannelID,7)<>1 and KS.C_S(ChannelID,7)<>2 Then Call KS.AlertHistory("KesionCMSϵͳ��������\n\n1����ģ������ҳû���������ɾ�̬HTML����\n\n2���뵽ģ�͹���->ģ����Ϣ�����������ɾ�̬Html����",-1):Exit Sub
						Call RefreshContent
			 Case "Folder"
			          
			          If ChannelID<>0 and KS.C_S(ChannelID,7)<>1 Then Call KS.AlertHistory("KesionCMSϵͳ��������\n\n1����ģ����Ŀҳû���������ɾ�̬HTML����\n\n2���뵽ģ�͹���->ģ����Ϣ�����������ɾ�̬Html����",-1):Exit Sub
  
						Call RefreshFolder
			End Select
			End With
		End Sub
		
		Sub Main()
		  With KS
		  .echo ("<html>")
		  .echo ("<head>")
		  .echo ("<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">")
		  .echo ("<title>ϵͳ��Ϣ</title>")
		  .echo ("<script src='../../ks_inc/jquery.js'></script>")
		  .echo ("<script src='../../ks_inc/kesion.box.js'></script>")
		  .echo ("<script type='text/javascript'>")
		  .echo (" function show()")
		  .echo (" { ")
		  .echo ("  PopupImgDir='../';")
		  .echo ("  var str=""<div style='height:60px;line-height:60px' id='fsotips'>������������,���Ժ�!</div>"";")
		  .echo ("  popupTips('������ʾ',str,510,300);")
		  .echo (" }")
		  .echo ("</script>")
		  .echo ("</head>")
		  .echo ("<link rel=""stylesheet"" href=""Admin_Style.css"">")
		'  If RefreshFlag<>"ID" Then
		'  .echo ("<body oncontextmenu=""return false;"" scroll=no>")
		'  Else
		  .echo ("<body oncontextmenu=""return false;"" scroll=no style='background-color:transparent'>")
		'  End If
		  If RefreshFlag="ID" Then
              .echo "<div style=""display:none"">"
				.echo "<br><br><br><table style=""display:none"" id=""BarShowArea"" width=""400"" border=""0"" align=""center"" cellspacing=""1"" cellpadding=""1"">"
		 Else
				.echo "<br><br><br><table id=""BarShowArea"" width=""400"" border=""0"" align=""center"" cellspacing=""1"" cellpadding=""1"">"
		 End iF
				.echo "<tr> "
				.echo "<td bgcolor=000000>"
				.echo " <table width=""400"" border=""0"" cellspacing=""0"" cellpadding=""1"">"
				.echo "<tr> "
				.echo "<td bgcolor=ffffff height=9><img src=""../images/114_r2_c2.jpg"" width=0 height=10 id=img2 name=img2 align=absmiddle></td></tr></table>"
				'.echo "<td bgcolor=ffffff height=9><span width=0 height=16 id=img2 name=img2 align=absmiddle bgcolor='#000000'></span></td></tr></table>"
				.echo "</td></tr></table>"
				.echo "<table width=""550"" border=""0"" align=""center"" cellspacing=""1"" cellpadding=""1""><tr> "
				.echo "<td align=center> <span id=txt2 name=txt2 style=""font-size:9pt"">0</span><span id=txt4 style=""font-size:9pt"">%</span></td></tr> "
				.echo "<tr><td align=center><span id=txt3 name=txt3 style=""font-size:9pt"">0</span></td></tr>"
				.echo "</table>"
			
			 .echo ("<table width=""80%""  border=""0"" align=""center"" cellpadding=""0"" cellspacing=""0"">")
			 .echo (" <tr>")
			 .echo ("   <td height=""50"">")
			 .echo ("     <div align=""center""> ")
			 .echo (ReturnInfo)
			 .echo ("       </div></td>")
			 .echo ("   </tr>")
			 .echo ("</table>")
			 .echo ("</div>")
		 

		 .echo ("<table width=""100%""   border=""0"" cellpadding=""0"" cellspacing=""0"">")
		 .echo (" <tr>")
		 .echo ("   <td height=""50"" id=""fsohtml"">")
		 .echo (FsoHtmlList)
		 .echo ("      </td>")
		 .echo ("   </tr>")
		 .echo ("</table>")
		 .echo ("</body>")
		 .echo ("</html>")
		 End With
		End Sub
		
		'================================================================================================================================
'                                                     ����Ϊ��ģ����Ӧ����ĺ���		'================================================================================================================================
		'������Ŀ�Ĵ������
		Sub RefreshFolder()
		With KS
		Dim FolderID, R_Sql, RefreshTotalNum, R_RS, NewsTotalNum, NewsNo		  
		 If NewsNo = "" Then NewsNo = 0
		  Select Case RefreshFlag
		    Case "ID"
			    FolderID = Trim(Request("FolderID"))
			    R_Sql = "Select * from KS_Class where ChannelID=" & ChannelID &" and  DelTF=0 And ID ='" & FolderID & "'"
			Case "IDS"
			    FolderID = Replace(Replace(Request("ID")," ",""),",","','")
			    R_Sql = "Select * from KS_Class where ID IN('" & FolderID & "')"
			Case "Folder"
				FolderID = Trim(Request("FolderID"))
				R_Sql = "Select * from KS_Class where ChannelID=" & ChannelID &" and  DelTF=0 And ID IN (" & FolderID & ") Order By FolderOrder ASC"
		   Case "All"
				R_Sql = "Select * from KS_Class where ChannelID=" & ChannelID &" and ClassType<>2 and DelTF=0 Order By FolderOrder ASC"
		   Case Else
			R_Sql = ""
		  End Select
		
		Call Main
		If R_Sql <> "" Then
			Set R_RS = Server.CreateObject("ADODB.RecordSet")
			R_RS.Open R_Sql, Conn, 1, 1
			If R_RS.EOF Then
				.echo "<script>img2.width=""0"";" & vbCrLf
				.echo "txt2.innerHTML=""û�п����ɵ�" & ItemName & "��Ŀ��<br><br><input name='button1' type='button' onclick=javascript:location.href='RefreshHtml.asp?Action=ref&ChannelID=" & ChannelID &"'; class='button' value=' �� �� '>"";" & vbCrLf
				.echo "txt3.innerHTML="""";" & vbCrLf
				.echo "txt4.innerHTML="""";" & vbCrLf
				.echo "document.all.BarShowArea.style.display='none';" & vbCrLf
				.echo "</script>" & vbCrLf
				R_RS.Close:Set R_RS = Nothing
				.die ""
			Else
			       NewsTotalNum = R_RS.RecordCount
				   For NewsNo=1 to NewsTotalNum
				    ChannelID=R_RS("ChannelID") : FCls.ChannelID=ChannelID
				    If R_RS("ClassPurview")=2 Then
				     FsoHtmlList="<table border=""0"">"_
								& "<tr><td><li><strong>ID��Ϊ��</strong></li></td><td> <font color=red>"  & R_RS("ID") & "</font> ����Ŀû������!</td></tr>"_
								& "<tr><td><li><strong>ԭ ��</strong></li></td><td>����Ŀ����Ϊ��֤��Ŀ"_
						& "</table>"		
					Else
						Dim FsoHtmlPath:FsoHtmlPath=KS.GetFolderPath(R_RS("ID"))
						FsoHtmlList="<table border=""0"">"_
									& "<tr><td><li><strong>ID �� Ϊ��</strong></li></td><td> <font color=red>"  & R_RS("ID") & "</font> ����Ŀ������</td></tr>"_
									& "<tr><td><li><strong>��Ŀ���ƣ�</strong></li></td><td><font color=red>" & R_RS("FolderName") & "</font></li></td><tr>" _
									& "<tr><td><li><strong>����·����</strong></li></td><td><a href=""" & FsoHtmlPath & """ target=""_blank"">" & FsoHtmlPath & "</a></li></td><tr>" _
									& "</table>"				
						Call KSRObj.RefreshFolder(ChannelID,R_RS)  '������Ŀˢ�º���
					End If
				
				    If RefreshFlag="ID" Then Call InnerJS(NewsNo,NewsTotalNum,"����Ŀ"):.Die ""
					
					Call InnerJS(NewsNo,NewsTotalNum,"����Ŀ")
					R_RS.MoveNext
					if Not Response.IsClientConnected then Exit FOR
				  Next
				.echo "<script>"
				.echo "fsohtml.innerHTML='';" & vbCrLf
				.echo "img2.width=400;" & vbCrLf
				.echo "txt2.innerHTML=""����" & ItemName & "��Ŀ������100"";" & vbCrLf
				.echo "txt3.innerHTML=""�ܹ������� <font color=red><b>" & NewsTotalNum & "</b></font> ��" & ItemName & "��Ŀ,�ܷ�ʱ:<font color=red>" & Left((Timer() - StartRefreshTime), 4) & "</font> ��<br><br><input name='button1' type='button' onclick=javascript:location='RefreshHtml.asp?Action=ref&ChannelID=" & ChannelID &"'; class='button' value=' �� �� '>"";" & vbCrLf
				.echo "img2.title=""(" & NewsNo & ")"";</script>" & vbCrLf
				'��ʱ����,�ر�
				if f="task" then
				 KS.Echo "<script>setTimeout('window.close();',3000);</script>"
				end if
				
				R_RS.Close:Set R_RS = Nothing
			End If
		Else
				.echo "<script>img2.width=""0"";" & vbCrLf
				.echo "txt2.innerHTML=""û�п����ɵ���Ŀ��<br><br><input name='button1' type='button' onclick=javascript:location='RefreshHtml.asp?Action=ref&ChannelID=" & ChannelID & "'; class='button' value=' �� �� '>"";" & vbCrLf
				.echo "txt3.innerHTML="""";" & vbCrLf
				.echo "txt4.innerHTML="""";" & vbCrLf
				.echo "document.all.BarShowArea.style.display='none';" & vbCrLf
				.echo "</script>" & vbCrLf
		End If
		End With
		End Sub
		
		'��������ҳ�Ĵ������
		Sub RefreshContent()
		Dim AlreadyRefreshByID, NowNum, R_Sql, R_RS, TotalNum,ID
		Dim StartDate, EndDate, FolderID, RefreshTotalNum,StartID,EndID
		AlreadyRefreshByID = Request.QueryString("AlreadyRefreshByID")
		RefreshTotalNum = Request.QueryString("RefreshTotalNum")
		NowNum = Request.QueryString("NowNum") '����ˢ�µڼ�ƪ����
		if KS.G("refreshtf")="1" then
		R_Sql=" Where refreshtf=0 and Verific=1"
		else
		R_Sql=" Where Verific=1"
		end if
		With KS
		If NowNum = "" Then NowNum = 0
		  Select Case RefreshFlag
		   Case "ID"
			    ID=KS.G("ID")
				R_Sql="Select Top 2 * From " & Table & R_SQL&" and ID IN(Select top 2 id from " & Table & R_Sql & " And ID<=" & id & " Order By ID Desc) Order By ID"
				RefreshTotalNum=conn.execute("select count(id) from " & Table  &" where verific=1 and ID<=" & ID)(0)
				If RefreshTotalNum>2 Then RefreshTotalNum=2
		   Case "IDS"
			    ID=KS.FilterIds(KS.G("ID"))
				If ID="" Then KS.Die "err!"
				R_Sql="Select Top 200 * From " & Table & R_SQL&" and ID IN(" & ID & ") Order By ID"
				RefreshTotalNum=conn.execute("select count(id) from " & Table  &" where verific=1 and ID in(" & ID & ")")(0)
		   Case "InfoID"
				 StartID = KS.ChkClng(KS.G("StartID"))
				 EndID = KS.ChkClng(KS.G("EndID"))
				 RefreshTotalNum=conn.execute("select count(id) from " & Table  & R_Sql & " and ID>= " & StartID & " And  ID <=" & EndID)(0)
				 R_Sql = "Select * from " & Table  & R_Sql & " and ID>= " & StartID & " And  ID <=" & EndID & " order by ID desc"
			Case "New"
			  TotalNum = KS.ChkCLng(Request("TotalNum"))
			  If TotalNum >conn.execute("select count(id) from "& Table )(0) Then TotalNum = conn.execute("select count(id) from "& Table )(0)
			  RefreshTotalNum = TotalNum
			  If TotalNum=0 Then TotalNum=1
			  R_Sql="Select Top " & TotalNum & " * from " & Table  & " Order By ID Desc"
		   Case "Date"
			  StartDate = Request("StartDate"):EndDate = DateAdd("d", 1, Request("EndDate"))
			 If CInt(DataBaseType) = 1 Then         'Sql
				 RefreshTotalNum=conn.execute("select count(id) from " & Table  & R_Sql & " And AddDate>= '" & StartDate & "' And  AddDate <='" & EndDate & "'")(0)
			    R_Sql = "Select * from " & Table  & R_Sql & " And AddDate>= '" & StartDate & "' And  AddDate <='" & EndDate & "' order by ID desc"
			Else                             'Access
				 RefreshTotalNum=conn.execute("select count(id) from " & Table  & R_Sql & " and AddDate>= #" & StartDate & "# And  AddDate <=#" & EndDate & "#")(0)
				 R_Sql = "Select * from " & Table  & R_Sql & " and AddDate>= #" & StartDate & "# And  AddDate <=#" & EndDate & "# order by ID desc"
			End If
		   Case "All"
		      RefreshTotalNum=conn.execute("select count(id) from " & Table  & R_Sql)(0)
			  R_Sql = "Select * from " & Table  & R_Sql & " order by ID desc"
		  Case "Folder"
			 FolderID = Trim(Replace(Request("FolderID")," ",""))
			 TotalNum = KS.ChkCLng(Request("TotalNum"))
			  If TotalNum >conn.execute("select count(id) from " & Table  & R_Sql& " And Tid IN(" & FolderID & ")")(0) Then TotalNum = conn.execute("select count(id) from " & Table  & R_Sql& " And Tid IN(" & FolderID & ")")(0)
			  RefreshTotalNum = TotalNum
			  If TotalNum=0 Then TotalNum=1
			  If KS.ChkCLng(Request("TotalNum"))<>0 Then
			   R_Sql = "Select top " & TotalNum & " * from " & Table  & R_Sql & " And Tid IN(" & FolderID & ") order by ID desc"
			  Else
			   R_Sql = "Select * from " & Table  & R_Sql & " And Tid IN(" & FolderID & ") order by ID desc"
			  End If
		  Case "Pause"
		     R_Sql=Request.QueryString("R_Sql")
			 RefreshTotalNum=KS.ChkClng(KS.G("RefreshTotalNum"))
		Case Else
			R_Sql = ""
			RefreshTotalNum = 0
		End Select
		Call Main
		If R_Sql <> "" Then
			Set R_RS = Server.CreateObject("ADODB.RecordSet")
			R_RS.Open R_Sql, Conn, 1, 1
			If R_RS.EOF And R_RS.BOF Then
				.echo "<script>img2.width=""0"";" & vbCrLf
				.echo "txt2.innerHTML=""û�п����ɵ�����ҳ��<br><br><input name='button1' type='button' onclick=javascript:location='RefreshHtml.asp?Action=ref&ChannelID=" & ChannelID  &"'; class='button' value=' �� �� '>"";" & vbCrLf
				.echo "txt3.innerHTML="""";" & vbCrLf
				.echo "txt4.innerHTML="""";" & vbCrLf
				.echo "document.all.BarShowArea.style.display='none';" & vbCrLf
				.echo "</script>" & vbCrLf
				Response.Flush
				R_RS.Close:Set R_RS=Nothing
				Exit Sub
			Else
				'On Error Resume Next
				Dim CurrNowNum:CurrNowNum=KS.ChkClng(KS.G("CurrNowNum"))
				If CurrNowNum=0 Then CurrNowNum=1
				R_RS.Move(CurrNowNum-1)
				For NowNum=CurrNowNum To RefreshTotalNum
				     Dim DocXML:Set DocXML=KS.arrayToXml(R_RS.GetRows(1),R_RS,"row","root")
				     Set KSRObj.Node=DocXml.DocumentElement.SelectSingleNode("row")
					  KSRObj.ModelID=ChannelID
					  KSRObj.ItemID = KSRObj.Node.SelectSingleNode("@id").text 

					If KS.C_S(ChannelID,7)=0 Then
					      FsoHtmlList=GetRefreshErr(KSRObj.Node,ItemName)
					Else
						  FsoHtmlList=GetRefreshSucc(KSRObj.Node,ItemName)
						  KSRObj.RefreshContent()
				    End If
				
				If Err.Number <> 0 Then
				 FsoHtmlList = "����ʧ��!<br><font color=red>" & Err.Description & "</font>"
				 Call InnerJS(NowNum,RefreshTotalNum,KS.C_S(ChannelID,4))
				End If
				If RefreshFlag="ID" and NowNum=2 Then
				 Call InnerJS(NowNum,RefreshTotalNum,KS.C_S(ChannelID,4)):R_RS.Close:Set R_RS=Nothing:.Die ""
				Else
				 Call InnerJS(NowNum,RefreshTotalNum,KS.C_S(ChannelID,4))
				End If
				
				if Not Response.IsClientConnected then Exit FOR
				If RefreshTotalNum>1 and NowNum Mod 100=0 Then
					 .echo "<script>"
				     .echo "fsohtml.innerHTML='<div style=""text-align:cdenter""><div style=""margin:10px;height:80px;padding:8px;border:1px dashed #cccccc;text-align:left;""><img src=""../images/succeed.gif"" align=""left""><br>&nbsp;&nbsp;&nbsp;&nbsp;<b>��ܰ��ʾ��</b><br>&nbsp;&nbsp;&nbsp;&nbsp;�������ռ�÷�������Դ��ϵͳ��ͣ2������<img src=""../../images/default/wait.gif""><br>&nbsp;&nbsp;&nbsp;&nbsp;���2���û�м���������<a href=""RefreshHtmlSave.asp?CurrNowNum=" & NowNum+1 & "&ChannelID=" & ChannelID & "&RefreshFlag=Pause&Types=" & Types & "&StartRefreshTime=" & StartRefreshTime & "&R_Sql=" & Server.UrlEncode(R_Sql) & "&RefreshTotalNum=" & RefreshTotalNum & """><font color=red>����</font></a>����<a href=""refreshhtml.asp?Action=ref&channelid=" & channelid & """><font color=red>ֹͣ</font></a>!</div></div>';" & vbCrLf
					 .echo "</script>" &vbcrlf
				     .die "<meta http-equiv=""refresh"" content=""2;url=RefreshHtmlSave.asp?f=" & f & "&CurrNowNum=" & NowNum+1 & "&ChannelID=" & ChannelID & "&RefreshFlag=Pause&Types=" & Types & "&StartRefreshTime=" & StartRefreshTime & "&R_Sql=" & Server.UrlEncode(R_Sql) & "&RefreshTotalNum=" & RefreshTotalNum & """>"
				End If
			Next
				.echo "<script>"
				.echo "fsohtml.innerHTML='';" & vbCrLf
				.echo "img2.width=400;" & vbCrLf
				.echo "txt2.innerHTML=""��������ҳ������100"";" & vbCrLf
				.echo "txt3.innerHTML=""�ܹ������� <font color=red><b>" & RefreshTotalNum & "</b></font> ��,�ܷ�ʱ:<font color=red>" & Left((Timer() - StartRefreshTime), 4) & "</font> ��<br><br><input name='button1' type='button' onclick=javascript:location='RefreshHtml.asp?Action=ref&ChannelID=" & ChannelID & "'; class='button' value=' �� �� '>"";" & vbCrLf
				.echo "img2.title=""(" & NowNum & ")"";</script>" & vbCrLf
				'��ʱ����,�ر�
				if f="task" then
				 KS.Echo "<script>setTimeout('window.close();',3000);</script>"
				end if
			End If
		Else
				.echo "<script>img2.width=""0"";" & vbCrLf
				.echo "txt2.innerHTML=""û�п����ɵ�����ҳ��<br><br><input name='button1' type='button' onclick=javascript:location='RefreshHtml.asp?Action=ref&ChannelID=" & ChannelID & "'; class='button' value=' �� �� '>"";" & vbCrLf
				.echo "txt3.innerHTML="""";" & vbCrLf
				.echo "txt4.innerHTML="""";" & vbCrLf
				.echo "document.all.BarShowArea.style.display='none';" & vbCrLf
				.echo "</script>" & vbCrLf
				'��ʱ����,�ر�
				if f="task" then
				 KS.Echo "<script>setTimeout('window.close();',3000);</script>"
				end if
		End If
		End With
		End Sub
		
		Function GetRefreshErr(Node,ItemName)
		GetRefreshErr="<table border=""0"">"_
								& "<tr><td><li><strong>ID ��Ϊ��</strong></li></td><td> <font color=red>"  & Node.SelectSingleNode("@id").text & "</font> ������û������!</td></tr>"_
								& "<tr><td><li><strong>����ԭ��</strong></li></td><td>1��" & ItemName & "Ƶ��û���������ɾ�̬HTML���ɹ��ܣ�<br>2����" & ItemName & "���ڵ���ĿΪ�뿪����Ŀ������֤��Ŀ��<br>3����" & ItemName & "��������Ҫ�۵�������οͲ��������������Ϊת�����ӣ�<br>"_
						& "</table>"	
		End Function
		Function GetRefreshSucc(Node,ItemName)
		 Dim str,FsoHtmlPath:FsoHtmlPath= KS.GetItemURL(ChannelID,Node.SelectSingleNode("@tid").text,Node.SelectSingleNode("@id").text,node.SelectSingleNode("@fname").text)
		 str=""
		 if RefreshFlag<>"ID" Then str="<img src=""../images/succeed.gif"" align=""left""><table border=""0"">"
		 GetRefreshSucc=str & "<table border=""0"">"_
									& "<tr><td><li><strong>ID ��Ϊ��</strong></li></td><td> <font color=red>"  & Node.SelectSingleNode("@id").text & "</font> ��" & ItemName & "������</td></tr>"_
									& "<tr><td><li><strong>" & ItemName & "���⣺</strong></li></td><td><font color=red>" & Node.SelectSingleNode("@title").text & "</font></li></td><tr>" _
									& "<tr><td><li><strong>����·����</strong></li></td><td><a href=""" & FsoHtmlPath & """ target=""_blank"">" & FsoHtmlPath & "</a></li></td><tr>" _
									& "</table>"
				  conn.execute("update " & Table & " set refreshtf=1 where id=" & Node.SelectSingleNode("@id").text)
		End Function
		
		Sub InnerJS(NowNum,TotalNum,itemname)
		  With KS
				.echo "<script>"
				if RefreshFlag<>"ID" Then
				.echo "fsohtml.innerHTML='<div style=""margin:10px;height:80px;padding:8px;border:1px dashed #cccccc;text-align:left;"">" & FsoHtmlList & "</div>';" & vbCrLf
			    else
				.echo "fsohtml.innerHTML='" & FsoHtmlList & "';" & vbCrLf
				end if
				.echo "img2.width=" & Fix((NowNum / TotalNum) * 400) & ";" & vbCrLf
				.echo "txt2.innerHTML=""���ɽ���:" & FormatNumber(NowNum / TotalNum * 100, 2, -1) & """;" & vbCrLf
				.echo "txt3.innerHTML=""�ܹ���Ҫ���� <font color=red><b>" & TotalNum & "</b></font> " & itemname & ",<font color=red><b>�ڴ˹���������ˢ�´�ҳ�棡����</b></font> ϵͳ�������ɵ� <font color=red><b>" & NowNum & "</b></font> " & itemname & """;" & vbCrLf
				.echo "img2.title=""(" & NowNum & ")"";</script>" & vbCrLf
				Response.Flush
		  End With
		End Sub
		
End Class
%> 