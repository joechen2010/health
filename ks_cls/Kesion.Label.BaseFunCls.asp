<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim LFCls:Set LFCls=New LabelBaseFunCls
Class LabelBaseFunCls
		Private KS     
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set LFCls=Nothing
		End Sub
		
		'ǰ̨�������ݱ��������
		Sub InserItemInfo(ChannelID,InfoID,Title,Tid,Intro,KeyWords,PhotoUrl,Inputer,Verific,Fname)
		 Call AddItemInfo(ChannelID,InfoID,Title,Tid,Intro,KeyWords,PhotoUrl,now,Inputer,0,0,0,0,0,0,0,0,0,0,1,Verific,Fname)
		End Sub
		'ǰ̨�������ݱ��޸�����
		Sub ModifyItemInfo(ChannelID,InfoID,Title,Tid,Intro,KeyWords,PhotoUrl,Verific)
		 Gcls.Execute("Update [KS_ItemInfo] Set Title='" & Title & "',Tid='" & Tid & "',Intro='" & Replace(left(KS.LoseHtml(Intro),255),"'","")  & "',KeyWords='" & Replace(KeyWords,"'","") & "',PhotoUrl='" & PhotoUrl & "',AddDate='" & Now & "',Verific=" & Verific & " Where  ChannelID=" & ChannelID & " and InfoID=" & InfoID)
		End Sub		
		
        '��̨��ϵͳ�����ݱ��������
        Sub AddItemInfo(ByVal ChannelID,InfoID,Title,Tid,Intro,KeyWords,PhotoUrl,AddDate,Inputer,Hits,HitsByDay,HitsByWeek,HitsByMonth,Recommend,Rolls,Strip,Popular,Slide,IsTop,Comment,Verific,Fname)
		 Gcls.Execute("Insert Into [KS_ItemInfo](ChannelID,InfoID,Title,Tid,Intro,KeyWords,PhotoUrl,AddDate,Inputer,Hits,HitsByDay,HitsByWeek,HitsByMonth,Recommend,Rolls,Strip,Popular,Slide,IsTop,Comment,Verific,DelTF,Fname) values(" & Channelid & "," & InfoID & ",'" & Title & "','" & Tid & "' ,'" & left(Replace(KS.LoseHtml(Intro),"'",""),255) & "','" & Replace(KeyWords,"'","") & "' ,'" & PhotoUrl & "' ,'" & AddDate & "' ,'" & Inputer & "' ," & Hits & "," & HitsByDay & ", " & HitsByWeek & "," & HitsByMonth & "," & Recommend & "," & Rolls & "," & KS.ChkClng(Strip) & "," & Popular & "," & Slide & "," & IsTop & "," & Comment & "," & Verific& ",0,'" & Fname & "')")
		End Sub
		'��̨�޸����ݱ�����
		Sub UpdateItemInfo(ChannelID,InfoID,Title,Tid,Intro,KeyWords,PhotoUrl,AddDate,Hits,HitsByDay,HitsByWeek,HitsByMonth,Recommend,Rolls,Strip,Popular,Slide,IsTop,Comment,Verific)
		 Gcls.Execute("Update [KS_ItemInfo] Set Title='" & Title & "',Tid='" & Tid & "',Intro='" & Replace(left(KS.LoseHtml(Intro),255),"'","")  & "',KeyWords='" & Replace(KeyWords,"'","") & "',PhotoUrl='" & PhotoUrl & "',AddDate='" & AddDate & "',Hits=" & Hits & ",HitsByDay=" & HitsByDay & ",HitsByWeek=" & HitsByWeek & ",HitsByMonth=" & HitsByMonth & ",Recommend=" & Recommend & ",Rolls=" & Rolls & ",Strip=" & Strip & ",Popular=" & Popular & ",Slide=" & Slide & ",IsTop=" & IsTop  &",Comment=" & Comment & ",Verific=" & Verific & " Where  ChannelID=" & ChannelID & " and InfoID=" & InfoID)
		End Sub		
		'*********************************************************************************************************
		'��������GetAbsolutePath
		'��  �ã��������ݿ�ľ���·��
		'��  ����RelativePath ���ݿ������ֶδ�
		'*********************************************************************************************************
		Function GetAbsolutePath(RelativePath)
			dim Exp_Path,Matches,tempStr
			tempStr=Replace(RelativePath,"\","/")
			if instr(tempStr,":/")>0 then
				GetAbsolutePath=RelativePath
				Exit Function
			End if
			set Exp_Path=new RegExp
			Exp_Path.Pattern="(Data Source=|dbq=)(.)*"
			Exp_Path.IgnoreCase=true
			Exp_Path.Global=true
			Set Matches=Exp_Path.Execute(tempStr)
			If instr(LCase(tempStr),"*.xls")<>0 Then
			GetAbsolutePath="driver={microsoft excel driver (*.xls)};dbq="&Server.MapPath(split(Matches(0).value,"=")(1))
			ElseIf Instr(Lcase(tempstr),"*.dbf")<>0 Then
			GetAbsolutePath="driver={microsoft dbase driver (*.dbf)};dbq="&Server.MapPath(split(Matches(0).value,"=")(1))
			Else
			GetAbsolutePath="Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(split(Matches(0).value,"=")(1))
			End If
		End Function

		'*********************************************************************************************************
		'��������ReplaceDBNull
		'��  �ã��滻���ݿ��ֵͨ�ú���
		'��  ����DBField �ֶ�ֵ,DefaultValue �ռ��滻��ֵ
		'*********************************************************************************************************
		Function ReplaceDBNull(DBField,DefaultValue)
		    If IsNull(DBField) Or (DBField="") then
			ReplaceDBNull=DefaultValue
			Else
			ReplaceDBNull = DBField
			end if
		End Function
		'*********************************************************************************************************
		'��������GetSingleFieldValue
		'��  �ã�ȡ���ֶ�ֵ
		'��  ����SQLStr SQL���
		'*********************************************************************************************************
		Function GetSingleFieldValue(SQLStr)
		    If DataBaseType=0 then
			On Error Resume Next
			GetSingleFieldValue=Conn.Execute(SQLStr)(0)
			If Err Then GetSingleFieldValue=""
			Else
			 Dim RS:Set RS=Conn.Execute(SQLStr)
			 If Not RS.Eof Then
			  GetSingleFieldValue=RS(0)
			 Else
			  GetSingleFieldValue=""
			 End If
			 RS.Close:Set RS=Nothing
			end if
		End Function
		
		'*********************************************************************************************************
		'��������GetConfigFromXML
		'��  �ã�ȡxml�ڵ�������Ϣ
		'��  ����FileName xml�ļ���(������չ��),Path �ڵ�·�� ,NodeName �ڵ�Name����ֵ
		'*********************************************************************************************************
		Function GetConfigFromXML(FileName,Path,NodeName)
		  If Not IsObject(Application(KS.SiteSN&"_Config"&FileName)) Then
			  Set Application(KS.SiteSN&"_Config"&FileName)=GetXMLFromFile(FileName)
		  End If  
		  GetConfigFromXML= Application(KS.SiteSN&"_Config"&FileName).documentElement.selectSingleNode(Path & "[@name='" & NodeName & "']").text
		End Function
		'*********************************************************************************************************
		'��������GetXMLFromFile
		'��  �ã�ȡxml�ļ���Application
		'��  ����FileName xml�ļ���(������չ��)
		'*********************************************************************************************************
		Function GetXMLFromFile(FileName)
		 	If Not IsObject(Application(KS.SiteSN&"_Config"&FileName)) Then
			  Dim Doc:set Doc = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			  Doc.async = false
			  Doc.setProperty "ServerHTTPRequest", true 
			  Doc.load(Server.MapPath(KS.Setting(3)&"Config/" & FileName &".xml"))
			  Set Application(KS.SiteSN&"_Config"&FileName)=Doc
		   End If  
            Set GetXMLFromFile=Application(KS.SiteSN&"_Config"&FileName)
		End Function
		
		'*********************************************************************************************************
		'��������ReplacePrevNext
		'��  �ã���һƪ����һƪ
		'��  ����NowID ����ID,Tid Ŀ¼ID,TypeStr����
		'*********************************************************************************************************
		Function GetPrevNextURL(ChannelID,NowID, Tid, TypeStr,ByRef Title)
		     Dim SqlStr,LinkUrl
		     SqlStr="SELECT Top 1 ID,Title,Tid,Fname From " & KS.C_S(ChannelID,2) & " Where Tid='" & Tid & "' And ID" & TypeStr & NowID & " And Verific=1 and  DelTF=0 Order By ID"
			 If TypeStr=">" Then SqlStr=SqlStr & " asc" else SqlStr=SqlStr & " desc"
			 Dim RS:Set RS=Conn.Execute(SqlStr)
			 If RS.EOF And RS.BOF Then
			  GetPrevNextURL = "#" : Title = "û����"
			 Else
			  LinkUrl = KS.GetItemURL(ChannelID,RS(2),RS(0),RS(3))
			  GetPrevNextURL = LinkUrl : Title= "<a href=""" & LinkUrl & """ title=""" & RS(1) & """>" & RS(1) & "</a>"
			 End If
			 RS.Close:Set RS = Nothing
		End Function
		Function ReplacePrevNext(ChannelID,NowID, Tid, TypeStr)
		     Dim Title
			 Call GetPrevNextURL(ChannelID,NowID, Tid, TypeStr,Title)
			 ReplacePrevNext=Title
		End Function
		
		
		'�滻�Զ����ֶ�
		Function ReplaceUserDefine(ChannelID,F_C,ByVal RS)
		   If Not IsObject(Application(KS.SiteSN&"_userfiledlist"&channelid)) Then
		     Set  Application(KS.SiteSN&"_userfiledlist"&channelid)=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			 Application(KS.SiteSN&"_userfiledlist"&channelid).appendChild(Application(KS.SiteSN&"_userfiledlist"&channelid).createElement("xml"))
				Dim D_F_Arr,K,Node,FieldName
				Dim KS_RS_Obj:Set KS_RS_Obj=Conn.Execute("Select FieldName From KS_Field Where ChannelID=" & ChannelID &" And ShowOnForm=1 Order By OrderID Asc")
				If Not KS_RS_Obj.Eof Then D_F_Arr=KS_RS_Obj.GetRows(-1)
			    KS_RS_Obj.Close:Set KS_RS_Obj=Nothing
				If IsArray(D_F_Arr) Then
					  For K=0 To Ubound(D_F_Arr,2)
						Set Node=Application(KS.SiteSN&"_userfiledlist"&channelid).documentElement.appendChild(Application(KS.SiteSN&"_userfiledlist"&channelid).createNode(1,"userfiledlist"&channelid,""))
						Node.attributes.setNamedItem(Application(KS.SiteSN&"_userfiledlist"&channelid).createNode(2,"fieldname","")).text=D_F_Arr(0,K)
					 Next
				 End If
		 End If
		 
		 For Each Node in Application(KS.SiteSN&"_userfiledlist"&channelid).documentElement.SelectNodes("userfiledlist"&channelid)
		     FieldName=Node.selectSingleNode("@fieldname").text
			If Not IsNull(RS(FieldName)) Then
			  F_C=Replace(F_C,"{$" & FieldName & "}",RS(FieldName))
			Else
			  F_C=Replace(F_C,"{$" & FieldName & "}","")
			End If
		 Next
		
		ReplaceUserDefine=F_C
	End Function
End Class
%> 
