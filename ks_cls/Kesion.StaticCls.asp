<!--#include file="Kesion.MemberCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************

Dim StaticCls
Set StaticCls=New KesionStaticCls
Class KesionStaticCls
        Private KS,KSUser, KSR,QueryParams,ChannelID,ThreadType,G_P_Arr
		Private FileContent,RS,SqlStr,Content,InfoPurview,ClassPurview,ReadPoint,ChargeType,PitchTime,ReadTimes
		Private DomainStr,ID,UserLoginTF,CurrPage,PayTF,UserName,UrlsTF
        Private PreListTag,PreContentTag,Extension
		Private DocXML
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSUser=New UserCls
		  Set KSR = New Refresh
		  DomainStr=KS.GetDomain
		  PreContentTag=GCls.StaticPreContent
		  PreListTag=GCls.StaticPreList
		  Extension=GCls.StaticExtension
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing:Set KSUser=Nothing
		End Sub
		Public Sub Run()
		   ChannelID=KS.ChkClng(KS.S("M"))
		   ID=KS.ChkClng(KS.S("D")) : If ID=0 Then ID=KS.ChkClng(KS.S("ID"))
		   If ChannelID<>0 And ID<>0 Then
		     if KS.C_S(ChannelID,48)=1 Then 
			  Response.Redirect (KS.Setting(3) & "?" & PreContentTag & "-" & ID & "-" & ChannelID & Extension)
			 end if
			 CurrPage=KS.ChkClng(KS.S("P"))
			 If CurrPage<=0 Then CurrPage=1
			 Call StaticContent()
		   ElseIf ID<>0 Then
		     CurrPage=KS.ChkClng(KS.S("Page")): If CurrPage<=0 Then CurrPage=1
		     Call StaticList()
		   Else
			   QueryParams=Replace(Lcase(Request.ServerVariables("QUERY_STRING")),Extension,"")
			   G_P_Arr=Split(QueryParams,"-")
			   If Ubound(G_P_Arr)<1 Then 
				 Response.Redirect("index.asp")
				 Response.End()
			   End If
			   ThreadType=G_P_Arr(0)
		   
			   ID=KS.ChkClng(G_P_Arr(1))
			   If ID=0 Then 
				 Response.Redirect("index.asp")
				 Response.End()
			   End If
			  
			   If ThreadType=PreContentTag Then
				   ChannelID=KS.ChkClng(G_P_Arr(2))
				   If ChannelID=0 Then  Response.Redirect("index.asp"): Response.End()
	
				 If Ubound(G_P_Arr)>=3 Then  CurrPage=KS.ChkClng(G_P_Arr(3))  Else  CurrPage=1
				 If Ubound(G_P_Arr)>=4 Then  PayTF=G_P_Arr(4)
				 If CurrPage<=0 Then CurrPage=CurrPage+1
				 
				 Call StaticContent()
			   ElseIf ThreadType=PreListTag Then
				 If Ubound(G_P_Arr)>=2 Then  CurrPage=KS.ChkClng(G_P_Arr(2))  Else  CurrPage=1
				 If CurrPage<=0 Then CurrPage=CurrPage+1
				 Call StaticList()
			   End If
		  End If
		End Sub
		'静态化列表
		Sub StaticList()
		 UserLoginTF=Cbool(KSUser.UserLoginChecked)
		 Dim RSObj:Set RSObj=Server.CreateObject("Adodb.Recordset")
		 RSObj.Open "Select ID,ClassPurview,TN,FolderTemplateID,FolderDomain,DefaultArrGroupID,ChannelID From KS_Class Where ClassID=" & ID,Conn,1,1
		 IF RSObj.Eof And RSObj.Bof Then  RSObj.Close:Set RSObj=Nothing:Call KS.Alert("非法参数!",""):Exit Sub

		  If RSObj("ClassPurview")=2 and  RSObj("channelid")<>8 Then
		    If Cbool(KSUser.UserLoginChecked)=false Then 
			 Call KS.Alert("本栏目为认证栏目，至少要求本站的注册会员才能浏览!",KS.GetDomain & "user/login/"):Response.End
		    elseIF KS.FoundInArr(RSObj("DefaultArrGroupID"),KSUser.GroupID,",")=false Then
		     Call KS.Alert("对不起，你所在的用户级没有权限浏览!",Request.ServerVariables("http_referer")):Response.End
		    End If
		  End If
		  	 ChannelID=RSObj("ChannelID")
		     Call FCls.SetClassInfo(ChannelID,RSObj("ID"),RSObj("TN"))
               
			 FileContent = KSR.LoadTemplate(RSObj("FolderTemplateID"))
			 FileContent = KSR.KSLabelReplaceAll(FileContent)
			Dim LabelParamStr:LabelParamStr=Application("PageParam")
			If LabelParamStr<>"" And Not IsNull(LabelParamStr) Then
				 Dim XMLDoc,XMLSql,LabelStyle,KMRFOBJ
				 Dim ParamNode,IncludeSubClass,ModelID,OrderStr,PrintType,PageStyle,PicStyle,ShowPicFlag,FieldStr,Param
				 Dim PerPageNumber,TotalPut,PageNum,TempStr,TableName
				 Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
				 If XMLDoc.loadxml("<label><param " & LabelParamStr & " /></label>") Then
					 Set ParamNode=XMLDoc.DocumentElement.SelectSingleNode("param")
					 ModelID         = ParamNode.getAttribute("modelid") : If Not IsNumeric(ModelID) Then ModelID=1
					 IncludeSubClass = ParamNode.getAttribute("includesubclass"):If KS.IsNul(IncludeSubClass) Then IncludeSubClass=true 
					 PrintType       = ParamNode.getAttribute("printtype") : If Not IsNumeric(PrintType) Then PrintType=1
					 PageStyle       = ParamNode.getAttribute("pagestyle") : If PageStyle="" Or IsNull(PageStyle) Then PageStyle=1
					 PicStyle        = ParamNode.getAttribute("picstyle")
					 OrderStr        = ParamNode.getAttribute("orderstr") : If OrderStr="" Or IsNull(OrderStr) Then OrderStr="ID Desc"
					 ShowPicFlag     = ParamNode.getAttribute("showpicflag") : If ShowPicFlag="" Or IsNull(ShowPicFlag) Then ShowPicFlag=false
					 PerPageNumber   = ParamNode.getAttribute("num") : If Not IsNumeric(PerPageNumber) Then PerPageNumber=10
					 
					 Param = " Where I.Verific=1 And I.DelTF=0"
					 If CBool(IncludeSubClass) = True Then 
					 Param= Param & " And I.Tid In (" & KS.GetFolderTid(RSObj("ID")) & ")" 
					 Else 
					 Param= Param & " And I.Tid='" & RSObj("ID") & "'"
					 End If
					 
					 Set KMRFObj= New RefreshFunction
					 Set KMRFObj.ParamNode=ParamNode
				     Call KMRFObj.LoadField(ChannelID,PrintType,PicStyle,ShowPicFlag,FieldStr,TableName,Param)
				
					If Lcase(Left(Trim(OrderStr),2))<>"id" Then  OrderStr=OrderStr & ",I.ID Desc"			
					SqlStr = "SELECT " & FieldStr & " FROM " & KS.C_S(ChannelID,2) & " I " & Param & " ORDER BY I.IsTop Desc," & OrderStr
					'response.write sqlstr
					Set RS=Server.CreateObject("ADODB.RECORDSET")
					RS.Open SqlStr, Conn, 1, 1
					If RS.EOF And RS.BOF Then
						TempStr = "<p>此栏目下没有" & KS.C_S(ChannelID,3) & "</p>"
					Else
						PerPageNumber=cint(PerPageNumber)
						TotalPut = Conn.Execute("select Count(id) from " & KS.C_S(ChannelID,2) & " I " & Param)(0)
						if (TotalPut mod PerPageNumber)=0 then
								PageNum = TotalPut \ PerPageNumber
						else
								PageNum = TotalPut \ PerPageNumber + 1
						end if
						If CurrPage >1 and (CurrPage - 1) * PerPageNumber < totalPut Then
							RS.Move (CurrPage - 1) * PerPageNumber
						Else
							CurrPage = 1
						End If
						Set XMLSQL=KS.ArrayToXml(RS.GetRows(PerPageNumber),RS,"row","root")
						Call KMRFObj.LoadPageParam(XMLSQL,ParamNode,ChannelID)
						LabelStyle=Application("LabelStyle")
						TempStr = KMRFObj.ExplainGerericListLabelBody(LabelStyle)
						XMLSql=Empty
						
						FCls.PageStyle=PageStyle       '分页样式
						FCls.TotalPage=PageNum         '总页数
						TempStr = TempStr & KS.GetPrePageList(FCls.PageStyle,KS.C_S(ChannelID,4),FCls.TotalPage,CurrPage,TotalPut,PerPageNumber) & "{KS:PageList}" 
						
					End If
				
					RS.Close:Set RS=Nothing					
					XMLDoc= Empty : Set ParamNode=Nothing
				End If	
				
			End If
			
			FileContent=Replace(FileContent,"{Tag:Page}",TempStr)
			
			
			If Instr(FileContent,"{KS:PageList}")<>0 Then
			  If KS.C_S(ChannelID,48)=0 Then
			   FileContent=Replace(FileContent,"{KS:PageList}",KS.GetPageList("?ID=" & ID,FCls.PageStyle,CurrPage,FCls.TotalPage, True))
			  ElseIf KS.C_S(ChannelID,48)=2 Then
			   FileContent=Replace(FileContent,"{KS:PageList}",KS.GetStaticPageList(GCls.StaticPreList & "-" & ID & "-",FCls.PageStyle,CurrPage,FCls.TotalPage,true,GCls.StaticExtension))
			  Else
			   FileContent=Replace(FileContent,"{KS:PageList}",KS.GetStaticPageList("?" & GCls.StaticPreList & "-" & ID & "-",FCls.PageStyle,CurrPage,FCls.TotalPage,true,GCls.StaticExtension))
			  End If
			End If
			 

		 RSObj.Close:Set RSObj=Nothing
		 Set KMRFObj=Nothing
		 KS.Echo FileContent
		End Sub
		
		
		'静态化内容页
		Sub StaticContent()
		  UserLoginTF=Cbool(KSUser.UserLoginChecked)
		  Select Case (KS.C_S(Channelid,6))
		   Case 1 Call StaticArticleContent()
		   Case 2 Call StaticPhotoContent()
		   Case 3 Call StaticDownContent()
		   Case 4 Call StaticFlashContent()
		   Case 5 Call StaticProductContent()
		   Case 7 Call StaticMovieContent()
		   Case 8 Call StaticSupplyContent()
		  End Select
		End Sub
		
		Function GetPageStr(Page)
		 If KS.C_S(ChannelID,48)=0 Then
		  GetPageStr="?m=" & ChannelID & "&d="& ID & "&p="&Page
		 ElseIf KS.C_S(ChannelID,48)=2 Then
		  GetPageStr=KS.Setting(3) & PreContentTag & "-" & ID & "-" & ChannelID & "-" & Page & Extension
		 Else
		  GetPageStr=KS.Setting(3) & "?" & PreContentTag & "-" & ID & "-" & ChannelID & "-" & Page & Extension
		 End If
		End Function
		
		Sub StaticArticleContent()
		 Set RS=Server.CreateObject("Adodb.Recordset")
		 RS.Open "Select top 1 a.*,ClassPurview,DefaultArrGroupID,DefaultReadPoint,DefaultChargeType,DefaultPitchTime,DefaultReadTimes From " & KS.C_S(ChannelID,2) & " a inner join KS_Class b on a.tid=b.id Where a.ID=" & ID,Conn,1,1
		 IF RS.Eof And RS.Bof Then
		  RS.Close:Set RS=Nothing
		  Call KS.Alert("您要查看的" & KS.C_S(ChannelID,3) & "已删除。或是您非法传递注入参数!",""):Exit Sub
		 ElseIF Cint(RS("Changes"))=1 Then 
		   Dim ClassID:ClassID=RS("Tid")
		   Dim Fname:Fname=RS("articlecontent")
		   RS.Close:Set RS=Nothing
		   Response.Redirect Fname
		 End IF
		  Set DocXML=KS.RsToXml(RS,"row","root") : RS.Close:Set RS=Nothing
		  With KSR 
			 Set .Node=DocXml.DocumentElement.SelectSingleNode("row")
		      .Tid=.Node.SelectSingleNode("@tid").text
			 If .Node.SelectSingleNode("@verific").text<>1 And UserLoginTF=False And KSUser.UserName<>.Node.SelectSingleNode("@inputer").text Then
			   KS.Echo "<script>alert('对不起，该" & KS.C_S(ChannelID,3) & "还没有通过审核!');</script>"
			   Response.End
			 End If
			 Call FCls.SetContentInfo(ChannelID,.Tid,ID,.Node.SelectSingleNode("@title").text)

			 InfoPurview = Cint(.Node.SelectSingleNode("@infopurview").text)
			 ReadPoint   = Cint(.Node.SelectSingleNode("@readpoint").text)
			 ChargeType  = Cint(.Node.SelectSingleNode("@chargetype").text)
			 PitchTime   = Cint(.Node.SelectSingleNode("@pitchtime").text)
			 ReadTimes   = Cint(.Node.SelectSingleNode("@readtimes").text)
			 ClassPurview= Cint(.Node.SelectSingleNode("@classpurview").text)
			 UserName    = .Node.SelectSingleNode("@inputer").text
		 
          If InfoPurview=2 or ReadPoint>0 Then
			   IF UserLoginTF=false Then
				 Call GetNoLoginInfo
			   Else
					 IF KS.FoundInArr(.Node.SelectSingleNode("@arrgroupid").text,KSUser.GroupID,",")=false and readpoint=0 Then
					   Content="<div align=center>对不起，你所在的用户组没有查看本" & KS.C_S(ChannelID,3) & "的权限!</div>"
					 Else
						  Call PayPointProcess()
					 End If
			   End If
		  ElseIF InfoPurview=0 And (ClassPurview=1 or ClassPurview=2) Then 
			  If UserLoginTF=false Then
			    Call GetNoLoginInfo
			  Else     
			     '============继承栏目收费设置时,读取栏目收费配置===========
			     ReadPoint  = Cint(.Node.SelectSingleNode("@defaultreadpoint").text)   
				 ChargeType = Cint(.Node.SelectSingleNode("@defaultchargetype").text)
				 PitchTime  = Cint(.Node.SelectSingleNode("@defaultpitchtime").text)
				 ReadTimes  = Cint(.Node.SelectSingleNode("@defaultreadtimes").text)
				 '============================================================
				 If ClassPurview=2 Then
					 IF KS.FoundInArr(.Node.SelectSingleNode("@defaultarrgroupid").text,KSUser.GroupID,",")=false Then
					    Content="<div align=""center"">对不起，你所在的用户组没有查看的权限!</div>"
					 Else
						Call PayPointProcess()
					 End If
				Else    
				 Call PayPointProcess()
				End If
			  End If
		 Else
		   Call PayPointProcess()
		 End If   
			
		 FileContent = KSR.LoadTemplate(.Node.SelectSingleNode("@templateid").text)
		 If InStr(FileContent,"[KS_Charge]")=0 Then
		   FileContent = Replace(FileContent,"{$GetArticleContent}","[KS_Charge]{$GetArticleContent}[/KS_Charge]")
		 End If
		 on error resume next		   
		 Dim ContentArr:ContentArr=Split(.Node.SelectSingleNode("@articlecontent").text,"[NextPage]")
		 Dim TotalPage,N,K,PageStr,NextUrl,PrevUrl
			TotalPage = Cint(UBound(ContentArr) + 1)
			   If TotalPage > 1 Then
					   If CurrPage = 1 Then
					     PrevUrl="" : NextUrl=GetPageStr(CurrPage + 1)
					   ElseIf CurrPage = TotalPage Then
					     NextUrl = KS.GetFolderPath(.Tid) : PrevUrl = GetPageStr(CurrPage - 1)
					   Else
					     NextUrl = GetPageStr(CurrPage + 1) :PrevUrl = GetPageStr(CurrPage - 1)
					   End If
					   PageStr =  "<div id=""pageNext"" style=""text-align:center""><table align=""center""><tr><td>"
					   If CurrPage > 1 And PrevUrl<>"" Then PageStr = PageStr & "<a class=""prev"" href=""" & PrevUrl & """>上一页</a> "
					 Dim StartPage:StartPage=1
					 if (CurrPage>=10) then StartPage=(CurrPage\10-1)*10+CurrPage mod 10+2
				     For N = StartPage To TotalPage
						 If CurrPage = N Then
						  PageStr = PageStr & ("<a class=""curr"" href=""#""><span style=""color:red"">" & N & "</span></a> ")
						 Else
						  PageStr = PageStr & ("<a class=""num"" href=""" & GetPageStr(N) & """>" & N & "</a> ")
						 End If
						 K=K+1
						 If K>=10 Then Exit For
					 Next
					 PageStr = ContentArr(CurrPage-1) & PageStr 
					 If CurrPage<>TotalPage Then PageStr = PageStr & " <a class=""next"" href=""" & NextUrl & """>下一页</a>"
					 PageStr = PageStr & "</td></tr></table></div>"
					 
					 Dim PageTitleArr,PageTitle
					 PageTitle=	.Node.SelectSingleNode("@pagetitle").text
					 
					 If PageTitle<>"" And Not IsNull(PageTitle) Then
					  PageTitleArr=Split(PageTitle,"§")
					  If CurrPage-1<=Ubound(PageTitleArr) Then
					   FileContent=Replace(FileContent,"{$GetArticleTitle}",PageTitleArr(CurrPage-1))
					  End If
					 End IF
				 Else
				  NextUrl=KS.GetFolderPath(.Tid)
				  PageStr = .Node.SelectSingleNode("@articlecontent").text
				 End If
				 
				 .ModelID = ChannelID
				 .ItemID  = ID
				 .PageContent=PageStr
				 .NextUrl=NextUrl
				 .TotalPage=TotalPage
				 .Templates=""
				 .Scan FileContent
		 		 FileContent = .Templates
		  If Content<>"True" Then
		   Dim ChargeContent:ChargeContent=KS.CutFixContent(FileContent, "[KS_Charge]", "[/KS_Charge]", 0)
		   FileContent=Replace(FileContent,"[KS_Charge]" & ChargeContent &"[/KS_Charge]",Content)
		  Else
		   FileContent=Replace(Replace(FileContent,"[KS_Charge]",""),"[/KS_Charge]","")
		  End If
		  If Instr(FileContent,"[KS_ShowIntro]")<>0 Then
			  If CurrPage=1 Then
		        FileContent=Replace(Replace(FileContent,"[KS_ShowIntro]",""),"[/KS_ShowIntro]","")
			  Else
		        FileContent=Replace(FileContent,KS.CutFixContent(FileContent, "[KS_ShowIntro]", "[/KS_ShowIntro]", 1),"")
			  End If
		  End If

		  FileContent = .KSLabelReplaceAll(FileContent)
		  
		End With
          FileContent=Replace(Replace(Replace(Replace(FileContent,"{§","{$"),"{#LB","{LB"),"{#SQL","{SQL"),"{#=","{=")
		  KS.Echo FileContent
		 
	   End Sub
	   
	   Sub StaticPhotoContent()
		 SqlStr= "Select a.*,ClassPurview,ClassID,DefaultArrGroupID,DefaultReadPoint,DefaultChargeType,DefaultPitchTime,DefaultReadTimes From " & KS.C_S(ChannelID,2) & " a inner join ks_class b on a.tid=b.id Where a.ID=" & ID
		 
		 Set RS=Server.CreateObject("Adodb.Recordset")
		 RS.Open SqlStr,Conn,1,1
		 IF RS.Eof And RS.Bof Then
		  Call KS.Alert("您要查看的" & KS.C_S(ChannelID,3) & "已删除。或是您非法传递注入参数!",""):Exit Sub
		 End IF
		 Set DocXML=KS.RsToXml(RS,"row","root") : RS.Close:Set RS=Nothing
		 With KSR 
			Set .Node=DocXml.DocumentElement.SelectSingleNode("row")
		    .Tid=.Node.SelectSingleNode("@tid").text

		 
		 If .Node.SelectSingleNode("@verific").text<>1 And UserLoginTF=False And KSUser.UserName<>.Node.SelectSingleNode("@inputer").text Then
		   KS.Echo "<script>alert('对不起，该" & KS.C_S(ChannelID,3) & "还没有通过审核!');</script>"
		   Response.End
		 End If
		 Call FCls.SetContentInfo(ChannelID,.Tid,ID,.Node.SelectSingleNode("@title").text)

			 InfoPurview = Cint(.Node.SelectSingleNode("@infopurview").text)
			 ReadPoint   = Cint(.Node.SelectSingleNode("@readpoint").text)
			 ChargeType  = Cint(.Node.SelectSingleNode("@chargetype").text)
			 PitchTime   = Cint(.Node.SelectSingleNode("@pitchtime").text)
			 ReadTimes   = Cint(.Node.SelectSingleNode("@readtimes").text)
			 ClassPurview= Cint(.Node.SelectSingleNode("@classpurview").text)

		 If InfoPurview=2 or ReadPoint>0 Then
               IF UserLoginTF=false Then
				 Call GetNoLoginInfo
			   Else
					 IF KS.FoundInArr(.Node.SelectSingleNode("@arrgroupid").text,KSUser.GroupID,",")=false and readpoint=0 Then
					   Content="<div align=center>对不起，你所在的用户组没有查看本" & KS.C_S(ChannelID,3) & "的权限!</div>"
					 Else
						  Call PayPointProcess()
					 End If
			   End If
		 ElseIF InfoPurview=0 And (ClassPurview=1 or ClassPurview=2) Then 
			  If UserLoginTF=false Then
			    Call GetNoLoginInfo
			  Else  
			     '============继承栏目收费设置时,读取栏目收费配置===========
			     ReadPoint  = Cint(.Node.SelectSingleNode("@defaultreadpoint").text)   
				 ChargeType = Cint(.Node.SelectSingleNode("@defaultchargetype").text)
				 PitchTime  = Cint(.Node.SelectSingleNode("@defaultpitchtime").text)
				 ReadTimes  = Cint(.Node.SelectSingleNode("@defaultreadtimes").text)
				 '============================================================
				 If ClassPurview=2 Then
					 IF KS.FoundInArr(.Node.SelectSingleNode("@defaultarrgroupid").text,KSUser.GroupID,",")=false Then
					   Content="<div align=""center"">对不起，你所在的用户组没有查看的权限!</div>"
					 Else
						Call PayPointProcess()
					 End If
				Else    
				 Call PayPointProcess()
				End If
			  End If
		 Else
		   Call PayPointProcess()
		 End If   
		 	Dim KSLabel:Set KSLabel =New RefreshFunction
			FileContent = KSR.LoadTemplate(.Node.SelectSingleNode("@templateid").text)
			 Dim PicUrlsArr,N,PageStr,TotalPage,NextUrl
			 If Cbool(UrlsTF)=true Then
				  PicUrlsArr = Split(Content, "|||")
				  TotalPage = Cint(UBound(PicUrlsArr) + 1)
				  
				  If InStr(FileContent, "{=GetPhotoPage") <> 0 Then
					Dim HtmlLabel:HtmlLabel = KSLabel.GetFunctionLabel(FileContent, "{=GetPhotoPage")
					Dim Param:Param = split(KSLabel.GetFunctionLabelParam(HtmlLabel, "{=GetPhotoPage"),",")
					Dim Rows:Rows=Param(0)
					Dim Cols:Cols=Param(1)
					Dim Width:Width=Param(2)
					Dim Height:Height=Param(3)
					Dim r,c,str
					 str="<table cellspacing=20 cellpadding=0 align=center border=0>"
					 if CurrPage<=1 then
					  n=0
					 else
					 n=(cols*rows)*(CurrPage-1)
					end if
					For r=1 to rows
					  str=str & "<tr>"
					 For c=1 To Cols
					      dim thumbsphoto
						  if n<=ubound(PicUrlsArr) Then
						  thumbsphoto="<table cellspacing=0 cellpadding=0 width=""100%"" align=center border=0><tr><td style='border:1px #999999 solid;background:#FFFFFF;padding:10px;text-align:center'><a id="""" href=""" & Split(PicUrlsArr(n), "|")(1) & """  class=""highslide"" onclick=""return hs.expand(this)"" title=""""><img alt='" & Split(PicUrlsArr(n), "|")(0) & "' width='" & width &"' height='" & height & "' src='" & Split(PicUrlsArr(n), "|")(2)  & "' style='border:1px #999999 solid' border='0'></a><div style='text-align:center'>" & Split(PicUrlsArr(n), "|")(0) & "</div></td></tr></table>"
						  else
						   thumbsphoto=""
						  end if
						  str=str & "<td valign=top>"
						  str=str & thumbsphoto 
						  str=str & "</td>"
						  n=n+1
					 Next
					 str=str & "</tr>"
					Next
					 str=str &"</table>"
					 
					    Dim TPage
					    if ((ubound(PicUrlsArr)+1) mod (cols*rows))=0 then
							Tpage=(ubound(PicUrlsArr)+1)\(cols*rows)
						else
							Tpage=(ubound(PicUrlsArr)+1)\(cols*rows) + 1
						end if

					 
					PageStr="<table style='BORDER-BOTTOM: #8eacca 1px solid' cellSpacing=0 cellPadding=0 width='95%' align=center border=0><tr><td class=text_9 width='54%' height=25>　共 <font color=#6699ff><strong>" & TPage&" </strong></font>页 第 <font color=#6699ff><strong>" & CurrPage & "</strong></font> 页</td><td class=text_9 align=right width='33%'>"
					
					    startpage=1:k=0
						 if (CurrPage>=10) then startpage=(CurrPage\10-1)*10+CurrPage mod 10+2
						 
						        PageStr=PageStr & "<a href=""" & GetPageStr(1) & """ title=""首页"">首页</a> "
								if CurrPage<>1 then
						        PageStr=PageStr & "<a href=""" & GetPageStr(CurrPage-1) & """ title=""上一页""><<</a> "
								end if
						  For N = startpage To TPage
							   If CurrPage = N Then
								 PageStr = PageStr & "<a href=""#""><font color=red>" & N & "</font></a>&nbsp;"
							   Else
								 PageStr = PageStr & "<a href=""" & GetPageStr(N) & """>" & N & "</a>&nbsp;"
							   End If
							   k=k+1
							  If k>=10  Then exit for
						Next
						If CurrPage <>tpage Then
						PageStr=PageStr & "<a href=""" & GetPageStr(Currpage+1) & """ title=""下一页"">>></a> "
						end if
						PageStr=PageStr & "<a href=""" & GetPageStr(tpage) & """ title=""末页"">末页</a> "
					
					PageStr=PageStr & "</td></tr></table>"
					
					 FileContent=Replace(FileContent, HtmlLabel,str & LFCls.GetConfigFromXML("highslide","/labeltemplate/label","highslide"))
					 FileContent=Replace(FileContent,"{$PageStr}",PageStr)

				  ElseIf InStr(FileContent, "{$GetPictureByPage}") <> 0 Then   '按分页方式生成图片内容页
					   If TotalPage > 1 Then
					          PageStr="<div class=""kspage"">" & vbcrlf & "<div style=""text-align:center"">"
							If CurrPage = 1 Then
							  NextUrl=GetPageStr(CurrPage+1)
							  PageStr = PageStr & "<a href=""" & NextUrl & """>下一张>></a>"
							ElseIf CurrPage = TotalPage Then
							  PageStr = PageStr & "<a href=""" & GetPageStr(CurrPage - 1) & """><<上一张</a>"
							  NextUrl=GetPageStr(1)
							Else
							  NextUrl=GetPageStr(CurrPage+1)
    						  PageStr = PageStr & "<a href=""" & GetPageStr(CurrPage - 1) & """><<上一张</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=""" & NextUrl & """>下一张>></a>"
							End If
							  PageStr =PageStr & "</div>"
						   
							  PageStr = PageStr & "<br /><div style=""text-align:left"">" & Split(PicUrlsArr(CurrPage-1), "|")(0) & "</div>"
							  PageStr = PageStr & "<br /><div style=""text-align:center""><a href=""#"">共<font color=""red""> " & TotalPage & "</font> 张</a>&nbsp;&nbsp;"
							
						 dim startpage,k
						 startpage=1:k=0
						 if (CurrPage>=10) then startpage=(CurrPage\10-1)*10+CurrPage mod 10+2
						 
						   PageStr=PageStr & "<a href=""" & GetPageStr(1) & """ title=""首页"">首页</a> "
						   if CurrPage<>1 then
						        PageStr=PageStr & "<a href=""" & GetPageStr(CurrPage-1) & """ title=""上一页""><<</a> "
						   end if
						  For N = startpage To TotalPage
							   If CurrPage = N Then
								 PageStr = PageStr & "<a href=""#""><font color=red>" & N & "</font></a>&nbsp;"
							   Else
								 PageStr = PageStr & "<a href=""" & GetPageStr(N) & """>" & N & "</a>&nbsp;"
							   End If
							   k=k+1
							  If k>=10  Then exit for
						Next
						If CurrPage <>totalpage Then
						PageStr=PageStr & "<a href=""" & GetPageStr(currpage+1) & """ title=""下一页"">>></a> "
						end if
						PageStr=PageStr & "<a href=""" & GetPageStr(totalpage) & """ title=""末页"">末页</a> "
						PageStr=PageStr & "</div>"
						
						PageStr = "<div style=""text-align:center""><a href=""" & NextUrl & """><Img onmousewheel=""return bbimg(this)"" onload=""javascript:resizepic(this)"" src=""" & Split(PicUrlsArr(CurrPage-1), "|")(1) & """ border=""0""></a></div>" & PageStr & "</div>"
					  Else
						PageStr = "<div style=""text-align:center""><img onmousewheel=""return bbimg(this)"" onload=""javascript:resizepic(this)"" src=""" & Split(Content, "|")(1) & """ border=""0""></div>"& PageStr
					  End If
					  'FileContent = KSR.ReplacePictureContent(ChannelID,RS, FileContent, PageStr)
				Else               '图片播放器方式
				   PageStr = KSR.GetPicturePlayer(PicUrlsArr,ChannelID)
				End If
			Else
			    PageStr = Content
				FileContent = Replace(Replace(FileContent,KSLabel.GetFunctionLabel(FileContent, "{=GetPhotoPage"),Content),"{$PageStr}","")
			End If
			     
				 .ModelID = ChannelID
				 .ItemID  = ID
				 .PageContent=PageStr
				 .NextUrl=NextUrl
				 .TotalPage=TotalPage
				 .Templates=""
				 .Scan FileContent
		 		  FileContent = .Templates
				
			FileContent = KSR.KSLabelReplaceAll(FileContent)
		  End With
		  KS.Echo FileContent
		  Set KSLabel=Nothing
	   End Sub
	   
	   Sub StaticDownContent()
	   	 SqlStr= "Select * From " & KS.C_S(ChannelID,2) & " Where ID=" & ID
		 Set RS=Server.CreateObject("Adodb.Recordset")
		 RS.Open SqlStr,Conn,1,1
		 IF RS.Eof And RS.Bof Then
		  Call KS.Alert("您要查看的软件已删除。或是您非法传递注入参数!",""):Exit Sub
		 End IF
		 Set DocXML=KS.RsToXml(RS,"row","root") : RS.Close:Set RS=Nothing
		 With KSR 
			Set .Node=DocXml.DocumentElement.SelectSingleNode("row")
		      .Tid=.Node.SelectSingleNode("@tid").text
			  Call FCls.SetContentInfo(ChannelID,.Tid,ID,.Node.SelectSingleNode("@title").text)
			  FileContent = .LoadTemplate(.Node.SelectSingleNode("@templateid").text)
			  .ModelID = ChannelID
			  .ItemID  = ID
			  .PageContent=""
			  .NextUrl=""
			  .TotalPage=0
			  .Templates=""
			  .Scan FileContent
			  FileContent = .Templates 
			  FileContent = .KSLabelReplaceAll(FileContent)
		 End With
		 KS.Echo FileContent
		 
	   End Sub
	   Sub StaticFlashContent()
		 Set RS=Server.CreateObject("Adodb.Recordset")
		 RS.Open "Select a.*,ClassPurview,DefaultArrGroupID,DefaultReadPoint,DefaultChargeType,DefaultPitchTime,DefaultReadTimes  From KS_Flash a inner join ks_class b on a.tid=b.id Where a.ID=" & ID,Conn,1,1
		 IF RS.Eof And RS.Bof Then
		  Call KS.Alert("您要查看的动漫已删除。或是您非法传递注入参数!",""):Exit Sub
		 End IF
		 If RS("Verific")<>1 And UserLoginTF=False And KSUser.UserName<>RS("Inputer") Then
		   Response.Write "<script>alert('对不起，该动漫还没有通过审核!');</script>"
		   Response.End
		 End If
		 Set DocXML=KS.RsToXml(RS,"row","root") : RS.Close:Set RS=Nothing
		 With KSR 
			Set .Node=DocXml.DocumentElement.SelectSingleNode("row")
		      .Tid=.Node.SelectSingleNode("@tid").text
			  Call FCls.SetContentInfo(ChannelID,.Tid,ID,.Node.SelectSingleNode("@title").text)
			 InfoPurview = Cint(.Node.SelectSingleNode("@infopurview").text)
			 ReadPoint   = Cint(.Node.SelectSingleNode("@readpoint").text)
			 ChargeType  = Cint(.Node.SelectSingleNode("@chargetype").text)
			 PitchTime   = Cint(.Node.SelectSingleNode("@pitchtime").text)
			 ReadTimes   = Cint(.Node.SelectSingleNode("@readtimes").text)
			 ClassPurview= Cint(.Node.SelectSingleNode("@classpurview").text)
		 
		 If InfoPurview=2 or ReadPoint>0 Then
			   IF UserLoginTF=false Then
				 Call GetNoLoginInfo
			   Else
					 IF KS.FoundInArr(.Node.SelectSingleNode("@arrgroupid").text,KSUser.GroupID,",")=false and readpoint=0 Then
					   Content="<div align=center>对不起，你所在的用户组没有查看本" & KS.C_S(ChannelID,3) & "的权限!</div>"
					 Else
					   Call PayPointProcess()
					 End If
			   End If
		 ElseIF InfoPurview=0 And (ClassPurview=1 or ClassPurview=2) Then 
			  If UserLoginTF=false Then
			    Call GetNoLoginInfo
			  Else         
			     '============继承栏目收费设置时,读取栏目收费配置===========
			     ReadPoint  = Cint(.Node.SelectSingleNode("@defaultreadpoint").text)   
				 ChargeType = Cint(.Node.SelectSingleNode("@defaultchargetype").text)
				 PitchTime  = Cint(.Node.SelectSingleNode("@defaultpitchtime").text)
				 ReadTimes  = Cint(.Node.SelectSingleNode("@defaultreadtimes").text)
				 '============================================================
				 If ClassPurview=2 Then
					 IF KS.FoundInArr(.Node.SelectSingleNode("@defaultarrgroupid").text,KSUser.GroupID,",")=false Then
					    Content="<div align=center>对不起，你所在的用户组没有查看的权限!</div>"
					 Else
						Call PayPointProcess()
					 End If
				Else    
				 Call PayPointProcess()
				End If
			  End If
		 Else
		   Call PayPointProcess()
		 End If  
		 
		 
		  
			 FileContent =.LoadTemplate(.Node.SelectSingleNode("@templateid").text)
			 
			 
		 
		  If Content<>"True" Then
		   Dim ChargeContent:ChargeContent=KS.CutFixContent(FileContent, "{=GetFlash", "}", 1)
		   If KS.IsNul(ChargeContent) Then ChargeContent=KS.CutFixContent(FileContent, "{=GetFlashByPlayer", "}", 1)
		   FileContent=Replace(FileContent,ChargeContent,Content)
		  End If
			 
			  .ModelID = ChannelID
			  .ItemID  = ID
			  .PageContent=""
			  .NextUrl=""
			  .TotalPage=0
			  .Templates=""
			  .Scan FileContent
			  FileContent = .Templates 
			  
		      FileContent = .KSLabelReplaceAll(FileContent)
		End With
		 KS.Echo FileContent
	   End Sub
	   Sub StaticProductContent()
	     Set RS=Server.CreateObject("Adodb.Recordset")
		 RS.Open "Select top 1 * From " & KS.C_S(ChannelID,2) & "  Where verific=1 And ID=" & ID ,Conn,1,1
		 IF RS.Eof And RS.Bof Then
		   RS.Close:Set RS=Nothing
		  Call KS.Alert("您要查看的" & KS.C_S(ChannelID,3) & "已删除或是未通过暂停销售!",""):Exit Sub
		 End IF
		 Set DocXML=KS.RsToXml(RS,"row","root") : RS.Close:Set RS=Nothing
		 With KSR 
			Set .Node=DocXml.DocumentElement.SelectSingleNode("row")
		      .Tid=.Node.SelectSingleNode("@tid").text
			  Call FCls.SetContentInfo(ChannelID,.Tid,ID,.Node.SelectSingleNode("@title").text)
			  FileContent = .LoadTemplate(.Node.SelectSingleNode("@templateid").text)
			  .ModelID = ChannelID
			  .ItemID  = ID
			  .PageContent=""
			  .NextUrl=""
			  .TotalPage=0
			  .Templates=""
			  .Scan FileContent
			  FileContent = .Templates 
			  FileContent = .KSLabelReplaceAll(FileContent)
		 End With
		 KS.Echo FileContent
	   End Sub
	   Sub StaticMovieContent()
		 Set RS=Server.CreateObject("Adodb.Recordset")
		 RS.Open "Select * From KS_Movie Where ID=" & ID,Conn,1,1
		 IF RS.Eof And RS.Bof Then
		  Call KS.Alert("您要观看的影片已删除。或是没有通过审核!",""):Exit Sub
		 End IF
		 If RS("Verific")<>1 And KS.C("UserName")<>RS("Inputer") Then
		   Response.Write "<script>alert('对不起，该" & KS.C_S(7,3) & "还没有通过审核!');</script>"
		   Response.End
		 End If
		 
		 Set DocXML=KS.RsToXml(RS,"row","root") : RS.Close:Set RS=Nothing
		 With KSR 
			Set .Node=DocXml.DocumentElement.SelectSingleNode("row")
		      .Tid=.Node.SelectSingleNode("@tid").text
			  Call FCls.SetContentInfo(7,.Tid,ID,.Node.SelectSingleNode("@title").text)
			  FileContent = .LoadTemplate(.Node.SelectSingleNode("@templateid").text)
			  .ModelID = 7
			  .ItemID  = ID
			  .PageContent=""
			  .NextUrl=""
			  .TotalPage=0
			  .Templates=""
			  .Scan FileContent
			  FileContent = .Templates 
			  FileContent = .KSLabelReplaceAll(FileContent)
		 End With
		 KS.Echo FileContent
	   End Sub
	   Sub StaticSupplyContent()
	   	 Set RS=Server.CreateObject("Adodb.Recordset")
		 If Not KS.IsNul(KS.C("AdminName")) Then
		 RS.Open "Select top 1 b.TemplateID,b.classpurview,b.defaultarrgroupid,a.* From KS_GQ a inner join KS_Class b on a.Tid=b.ID where a.ID=" & ID,Conn,1,1
		 Else
		 RS.Open "Select top 1 b.TemplateID,b.classpurview,b.defaultarrgroupid,a.* From KS_GQ a inner join KS_Class b on a.Tid=b.ID where a.verific=1 and a.ID=" & ID,Conn,1,1
		 End If
		 IF RS.Eof And RS.Bof Then
		  Call KS.Alert("您要查看的信息已删除或未通过审核!",""):Exit Sub
		 End IF
		  FileContent = KSR.LoadTemplate(rs(0))
		 Set DocXML=KS.RsToXml(RS,"row","root") : RS.Close:Set RS=Nothing
		 With KSR 
			Set .Node=DocXml.DocumentElement.SelectSingleNode("row")
		      .Tid=.Node.SelectSingleNode("@tid").text
			  Call FCls.SetContentInfo(8,.Tid,ID,.Node.SelectSingleNode("@title").text)
			  .ModelID = 8
			  .ItemID  = ID
			  .PageContent=""
			  .NextUrl=""
			  .TotalPage=0
			  .Templates=""
			  .Scan FileContent
			  FileContent = .Templates 
			  
			  Dim ClassPurView:ClassPurview=.Node.SelectSingleNode("@classpurview").text
			  Dim DefaultArrGroupID:DefaultArrGroupID=.Node.SelectSingleNode("@defaultarrgroupid").text
			  If ClassPurView="2" And Not KS.IsNul(DefaultArrGroupID) And DefaultArrGroupID<>"0" Then
			  	Dim ChargeContent:ChargeContent=KS.CutFixContent(FileContent, "[KS_Charge]", "[/KS_Charge]", 1)
				IF UserLoginTF=false Then
		        FileContent=Replace(FileContent,ChargeContent,"<script src=""" & KS.Setting(3) & "ks_inc/kesion.box.js"" language=""JavaScript""></script><script>function ShowLogin(){ popupIframe('会员登录','" & KS.Setting(3) & "user/userlogin.asp?Action=Poplogin',397,184,'no');}</script><div style='padding:10px;border:1px dashed #cccccc;text-align:center'>对不起,您还没有登录，请<a href='javascript:ShowLogin()'>登录</a>后再查看联系信息。</div>")
				ElseIf KS.FoundInArr(DefaultArrGroupID,KSUser.GroupID,",")=false Then
		        FileContent=Replace(FileContent,ChargeContent,"<div style='padding:10px;border:1px dashed #cccccc;text-align:center'>对不起,您的级别不够,无法查看联系信息!得到更好服务,请联系本站管理员。</div>")
				End If
			  End If
			    FileContent=Replace(Replace(FileContent,"[KS_Charge]",""),"[/KS_Charge]","")
			  
			  FileContent = .KSLabelReplaceAll(FileContent)
		 End With
		 KS.Echo FileContent
	   End Sub
	   
	   '收费扣点处理过程
	   Sub PayPointProcess()
	       Dim UserChargeType:UserChargeType=KSUser.ChargeType
	        If (Cint(ReadPoint)>0 or InfoPurview=2 or (InfoPurview=0 And (ClassPurview=1 Or ClassPurview=2))) and KSUser.UserName<>UserName Then
					 
					     If UserChargeType=1 Then
							 Select Case ChargeType
							  Case 0:Call CheckPayTF("1=1")
							  Case 1:Call CheckPayTF("datediff(" & DataPart_H &",AddDate," & SqlNowString & ")<" & PitchTime)
							  Case 2:Call CheckPayTF("Times<" & ReadTimes)
							  Case 3:Call CheckPayTF("datediff(" & DataPart_H &",AddDate," & SqlNowString & ")<" & PitchTime & " or Times<" & ReadTimes)
							  Case 4:Call CheckPayTF("datediff(" & DataPart_H &",AddDate," & SqlNowString & ")<" & PitchTime & " and Times<" & ReadTimes)
							  Case 5:Call PayConfirm()
							  End Select
						Elseif UserChargeType=2 Then
				          If KSUser.GetEdays <=0 Then
						     Content="<div align=center>对不起，你的账户已过期 <font color=red>" & KSUser.GetEdays & "</font> 天,此文需要在有效期内才可以查看，请及时与我们联系！</div>"
						  Else
						   Call GetContent()
						  End If
						Else
						 Call GetContent()
						end if
					   Else
						  Call GetContent()
					   End IF
	   End Sub
	   '检查是否过期，如果过期要重复扣点券
	   '返回值 过期返回 true,未过期返回false
	   Sub CheckPayTF(Param)
	    Dim SqlStr:SqlStr="Select top 1 Times From KS_LogPoint Where ChannelID=" & ChannelID & " And InfoID=" & ID & " And InOrOutFlag=2 and UserName='" & KSUser.UserName & "' And (" & Param & ") Order By ID"
	    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		RS.Open SqlStr,conn,1,3


		IF RS.Eof And RS.Bof Then
			Call PayConfirm()	
		Else
		       RS.Movelast
			   RS(0)=RS(0)+1
			   RS.Update
			   Call GetContent()
		End IF
		 RS.Close:Set RS=nothing
	   End Sub
	   
	   Sub PayConfirm()
	     If UserLoginTF=false Then Call GetNoLoginInfo():Exit Sub
		 If ReadPoint<=0 Then Call GetContent():Exit Sub

			 If Cint(KSUser.Point)<ReadPoint Then
					 Content="<div style=""text-align:center"">对不起，你的可用" & KS.Setting(45) & "不足!阅读本文需要 <span style=""color:red"">" & ReadPoint & "</font> " & KS.Setting(46) & KS.Setting(45) &",你还有 <span style=""color:green"">" & KSUser.Point & "</span> " & KS.Setting(46) & KS.Setting(45) & "</div>,请及时与我们联系！" 
			 Else
					If PayTF="1" Then
						IF Cbool(KS.PointInOrOut(ChannelID,ID,KSUser.UserName,2,ReadPoint,"系统","阅读收费" & KS.C_S(ChannelID,3) & "“" & KSR.Node.SelectSingleNode("@title").text & "”",0))=True Then 
						 '支付投稿者提成
						 Dim PayPoint:PayPoint=(ReadPoint*KS.C_C(KSR.Tid,11))/100
						 If PayPoint>0 Then
						 Call KS.PointInOrOut(ChannelID,ID,KSR.Node.SelectSingleNode("@inputer").text,1,PayPoint,"系统",KS.C_S(ChannelID,3) & "“" & KSR.Node.SelectSingleNode("@title").text & "”的提成",0)
						 End If
						 Call GetContent()
						End If
					Else
						Content="<div align=center>阅读本文需要消耗 <font color=red>" & ReadPoint & "</font> " & KS.Setting(46) & KS.Setting(45) &",你目前尚有 <font color=green>" & KSUser.Point & "</font> " & KS.Setting(46) & KS.Setting(45) &"可用,阅读本文后，您将剩下 <font color=blue>" & KSUser.Point-ReadPoint & "</font> " & KS.Setting(46) & KS.Setting(45) &"</div><div align=center>你确实愿意花 <font color=red>" & ReadPoint & "</font> " & KS.Setting(46) & KS.Setting(45) & "来阅读此文吗?</div><div>&nbsp;</div><div align=center><a href=""?"& PreContentTag & "-"&ID & "-" & ChannelID & "-" & CurrPage &"-" &"1"& Extension & """>我愿意</a>    <a href=""" &DomainStr & """>我不愿意</a></div>"
					End If
			 End If
	   End Sub
	   Sub GetNoLoginInfo()
	       GCls.ComeUrl=GCls.GetUrl()
		   Content="<div style='text-align:center'><script src='../ks_inc/kesion.box.js' language=""JavaScript""></script><script>function ShowLogin(){popupIframe('会员登录','../user/userlogin.asp?Action=Poplogin',397,184,'no');}</script>对不起，你还没有登录，本文至少要求本站的注册会员才可查看!</div><div style='text-align:center'>如果你还没有注册，请<a href=""../User/reg/""><font color=red>点此注册</font></a>吧!</div><div style='text-align:center'>如果您已是本站注册会员，赶紧<a href=""javascript:ShowLogin();""><font color=red>点此登录</font></a>吧！</div>"
	   End Sub
	   Sub GetContent()
	     Select Case (KS.C_S(Channelid,6))
		  Case 1 Content="True"
		  Case 2 Content=KSR.Node.SelectSingleNode("@picurls").text
		  Case 4 Content="True"
		 End Select
		 UrlsTF=true
	   End Sub
End Class
%>