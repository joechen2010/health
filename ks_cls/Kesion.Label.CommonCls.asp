<!--#include file="Kesion.Label.FunctionCls.asp"-->
<!--#include file="Kesion.Label.LocationCls.asp"-->
<!--#include file="Kesion.Label.SearchCls.asp"-->
<!--#include file="Kesion.Label.SQLCls.asp"-->
<!--#include file="Kesion.Label.JSCls.asp"-->
<!--#include file="Kesion.Label.BaseFunCls.asp"-->
<!--#include file="Kesion.FsoVarCls.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Class Refresh
		Private KS,KSLabel,DomainStr  
		public Templates,ModelID,Tid,ItemID          rem  ModelID ģ��ID ItemID �ĵ�ID      
		public Node,PageContent,NextUrl,TotalPage    rem  Node �ڵ����,PageContent ��ҳ����
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSLabel =New RefreshFunction
		  DomainStr=KS.GetDomain
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSLabel=Nothing
		End Sub
		
		Sub Echo(sStr)
			Templates    = Templates & sStr 
		End Sub
		Sub EchoLn(sStr)
		    Templates    = Templates & sStr & VbNewLine
		End Sub

		public Sub Scan(ByVal sTemplate)
		    If Fcls.RefreshType="Content" Then Call ReplaceHits(sTemplate,ModelID,ItemId)  '����ҳ���滻�������ǩ
			Dim iPosLast, iPosCur
			iPosLast    = 1
			Do While True 
				iPosCur    = InStr(iPosLast, sTemplate, "{$") 
				If iPosCur>0 Then
					Echo    Mid(sTemplate, iPosLast, iPosCur-iPosLast)
					iPosLast  = Parse(sTemplate, iPosCur+2)
				Else 
					Echo    Mid(sTemplate, iPosLast)
					Exit Do  
				End If 
		   Loop
		   
		   rem ɨ��Ⱥű�ǩ
		   sTemplate=Templates : Templates="" : iPosLast   = 1
		   Do While True
			   iPosCur = Instr(iPosLast, sTemplate,"{=")
			   If iPosCur > 0 Then
			     Echo    Mid(sTemplate, iPosLast, iPosCur-iPosLast)
				 iPosLast  = ParseEqual(sTemplate, iPosCur+2)
			   Else
				  Echo    Mid(sTemplate, iPosLast)
				  Exit Do  
			   End IF
		   Loop
		End Sub 
		
		Function GetNodeText(NodeName)
		 Dim N
		 If IsObject(Node) Then
		  set N=node.SelectSingleNode("@" & NodeName)
		  If Not N is Nothing Then GetNodeText=N.text
		 End If
		End Function
		
		Function Parse(sTemplate, iPosBegin)
			Dim iPosCur, sToken, sTemp,MyNode
			iPosCur      = InStr(iPosBegin, sTemplate, "}")
			sTemp        = Mid(sTemplate,iPosBegin,iPosCur-iPosBegin)
			iPosBegin    = iPosCur+1
			ParseChannelLabel sTemp   rem ����Ƶ����ǩ
			ParseRssLabel     sTemp   rem ����RSS��ǩ
			select case Lcase(sTemp)
			 '================================��վͨ�ò�����ʼ===========================
			 case "getsiteurl"      echo Domainstr
			 case "getsitename"     echo KS.Setting(0)
			 case "getsitetitle"    echo KS.Setting(1)
			 case "getsitelogo"     echo "<img src=""" & KS.Setting(4) & """ border=""0"" align=""absmiddle"" alt=""logo"" />"
			 case "getsitecountall" echo GetSiteCountAll()
			 case "getsiteonline"   echo "<script src=""" & DomainStr & "plus/online.asp?Referer=""+escape(document.referrer) type=""text/javascript""></script>"
             case "gettopuserlogin" echo "<iframe width=""520"" height=""22"" id=""toplogin"" name=""toplogin"" src=""" & KS.Setting(3) & "user/userlogin.asp?action=Top"" frameborder=""0"" scrolling=""no"" allowtransparency=""true""></iframe>"
             case "getuserlogin"    echo "<iframe width=""180"" height=""122"" id=""loginframe"" name=""loginframe"" src=""" & KS.Setting(3) & "user/userlogin.asp"" frameborder=""0"" scrolling=""no"" allowtransparency=""true""></iframe>"
             case "getspecial"
			      Dim SpecialIndexUrl,SpecialDir:SpecialDir = KS.Setting(95)
				  If Split(KS.Setting(5),".")(1)<>"asp" Then SpecialIndexUrl=DomainStr & SpecialDir Else SpecialIndexUrl=DomainStr & "SpecialIndex.asp"
				  echo "<a href=""" & SpecialIndexUrl & """ target=""_blank"">ר����ҳ</a>"
             case "getfriendlink"   echo "<a href=""" & DomainStr & "plus/Link/"" target=""_blank"">��������</a>"
			 case "getinstalldir"   echo KS.Setting(3)
			 case "getmanagelogin"  echo "<a href=""" & DomainStr & KS.Setting(89) & "Login.asp"" target=""_blank"">�����¼</a>"
			 case "getcopyright"    echo KS.Setting(18)
			 case "getmetakeyword"  echo KS.Setting(19)
			 case "getmetadescript" echo KS.Setting(20)
			 case "getwebmaster"    echo "<a href=""mailto:" & KS.Setting(11) & """>" & KS.Setting(10) & "</a>"
			 case "getwebmasteremail" echo KS.Setting(11)
			 case "getsiteurl"         echo DomainStr
			 '================================��վͨ�ò�������===========================
			 
			 
			 
			 case "channelid"     echo ModelID
			 case "infoid"        echo ItemID
             case "itemname"      echo KS.C_S(ModelID,3)
			 case "itemunit"      echo KS.C_S(ModelID,4)
			 case "getusername"   echo GetNodeText("inputer")
		     case "getrank"       echo Replace(GetNodeText("rank"),"��","<img src=""" & DomainStr & "Images/Star.gif"" border=""0"">")
		     case "getdate"       echo GetNodeText("adddate")
			 case "getkeytags" echo ReplaceKeyTags(GetNodeText("keywords"))
			 case "getshowcomment" If GetNodeText("comment")="1" Then  echo "<script src=""" & DomainStr & "ks_inc/Comment.page.js"" type=""text/javascript""></script><script src=""" & DomainStr & "ks_inc/kesion.box.js"" type=""text/javascript""></script><script type=""text/javascript"" defer>Page(1," & ModelID & ",'" & ItemID & "','Show','"& DomainStr & "');</script><div id=""c_" & ItemID & """></div><div id=""p_" & ItemID & """ align=""right""></div>"
			 case "getwritecomment" If GetNodeText("comment")="1" Then echo "<script tyle=""text/Javascript"" src=""" & DomainStr & "plus/Comment.asp?Action=Write&ChannelID=" & ModelID & "&InfoID=" & ItemID & """></script>"
           case "getprevurl" echo LFCls.GetPrevNextURL(ModelID,ItemID, GetNodeText("tid"), "<","")
           case "getnexturl" echo LFCls.GetPrevNextURL(ModelID,ItemID, GetNodeText("tid"), ">","")
			 
			 '================================����ģ�Ϳ�ʼ================================
			 case "getarticletitle" echo LFCls.ReplaceDBNull(GetNodeText("fulltitle"),GetNodeText("title"))
			 case "getarticlesize"  
				  echoln "<script type=""text/javascript"" language=""javascript"">"
				  echoln  "function ContentSize(size)"
				  echoln  "{document.getElementById('MyContent').style.fontSize=size+'px';}" 
				  echoln  "</script>"
				  echoln "�����壺<a href=""javascript:ContentSize(16)"">��</a> <a href=""javascript:ContentSize(14)"">��</a> <a href=""javascript:ContentSize(12)"">С</a>��"
			case "getarticlecontent"
			      echoln "<div id=""MyContent"">"
			      echoln ReplaceAd(FormatImgLink(KS.ReplaceInnerLink(Replace(Replace(Replace(Replace(PageContent,"{$","{��"),"{LB","{#LB"),"{SQL","{#SQL"),"{=","{#=")),NextUrl,TotalPage),GetNodeText("tid"))
				  echoln "</div>"
			case "getarticleaction"
			      echo "��<A href=""" & DomainStr & "plus/Comment.asp?ChannelID=" & ModelID & "&InfoID=" & ItemID & """ target=""_blank"">��������</A>����<A href=""" & DomainStr & "item/SendMail.asp?m="&ModelID &"&ID=" & ItemID & """ target=""_blank"">���ߺ���</A>����<A href=""" & DomainStr & "item/Print.asp?m=" & ModelID &"&ID=" & ItemID & """ target=""_blank"">��ӡ����</A>����<A href=""" & DomainStr & "User/index.asp?User_Favorite.asp?Action=Add&ChannelID=" & ModelID & "&InfoID=" & ItemID & """ target=""_blank"">�ղش���</A>����<A href=""javascript:window.close();"">�رմ���</A>��"
			case "getarticleintro" echo GetNodeText("intro")
			case "getarticleshorttitle" echo GetNodeText("title")
			case "getarticleurl"   echo KS.GetItemURL(ModelID,GetNodeText("tid"),ItemID,GetNodeText("fname"))
			case "getarticlekeyword" echo Replace(GetNodeText("keywords"), "|", ",")
			case "getarticleauthor" echo LFCls.ReplaceDBNull(GetNodeText("author"),"����")
			case "getarticleinput" echo "<a href=""" & DomainStr & "Space/?" & GetNodeText("inputer") &""" target=""_blank"">" & GetNodeText("inputer") & "</a>"
			case "getarticleorigin" echo KS.GetOrigin(LFCls.ReplaceDBNull(GetNodeText("origin"),"��վԭ��"))
			case "getarticleproperty" 
			  If GetNodeText("recommend") = "1" Then echo "<span title=""�Ƽ�"" style=""cursor:default;color:green"">��</span> "
			  If GetNodeText("popular") = "1" Then  echo "<span title=""����"" style=""cursor:default;color:red"">��</span> "
			  If GetNodeText("strip")="1" Then echo "<span title=""����ͷ��"" style=""cursor:default;color:#0000ff"">ͷ</span> "
			  If GetNodeText("rolls") = "1" Then echo "<span title=""����"" style=""cursor:default;color:#F709F7"">��</span> "
			  If GetNodeText("slide") = "1" Then echo "<span title=""�õ�Ƭ"" style=""cursor:default;color:black"">��</span>"
		   case "getarticledate" echo KS.DateFormat(GetNodeText("adddate"), 6)
           case "getprevarticle" echo LFCls.ReplacePrevNext(ModelID,ItemID, GetNodeText("tid"), "<")
           case "getnextarticle" echo LFCls.ReplacePrevNext(ModelID,ItemID, GetNodeText("tid"), ">")
		   case "getpictureaction" echo "��<A href=""" & DomainStr & "plus/Comment.asp?ChannelID=" & ModelID & "&InfoID=" & ItemID & """ target=""_blank"">��������</A>����<A href=""" & DomainStr & "User/index.asp?User_Favorite.asp?Action=Add&ChannelID=" & ModelID & "&InfoID=" & ItemID & """ target=""_blank"">��Ҫ�ղ�</A>����<A href=""javascript:window.close();"">�رմ���</A>��"

		   '================================����ģ�ͽ���=================================
			 
			 
		   '================================ͼƬģ�Ϳ�ʼ================================
		   case "getpicturename" echo GetNodeText("title")
		   case "getpicturebypage","getpicturebyplayer" echo PageContent
		   case "getpictureintro" echo KS.ReplaceInnerLink(GetNodeText("picturecontent"))
		   case "getpictureurl"   echo KS.GetItemURL(ModelID,GetNodeText("tid"),ItemID,GetNodeText("fname"))
		   case "getpicturekeyword" echo Replace(GetNodeText("keywords"), "|", ",")
		   case "getpictureauthor" echo LFCls.ReplaceDBNull(GetNodeText("author"),"����")
		   case "getpictureinput"   echo "<a href=""" & DomainStr & "Space/?" & GetNodeText("inputer") &""" target=""_blank"">" & GetNodeText("inputer") & "</a>"
		   case "getpicturesrc"    echo GetNodeText("photourl")
		   case "getpictureorigin" echo KS.GetOrigin(LFCls.ReplaceDBNull(GetNodeText("origin"),"��վԭ��"))
		   case "getpictureproperty"
		     If GetNodeText("recommend") = "1" Then Echo "<span title=""�Ƽ�"" style=""cursor:default;color:green"">��</span> "
			 If GetNodeText("popular") = "1" Then  echo "<span title=""����"" style=""cursor:default;color:red"">��</span> "
			 If GetNodeText("strip")="1" Then echo "<span title=""����ͷ��"" style=""cursor:default;color:#0000ff"">ͷ</span> "
			 If GetNodeText("rolls") = "1" Then echo "<span title=""����"" style=""cursor:default;color:#F709F7"">��</span> "
			 If GetNodeText("slide") = "1" Then echo "<span title=""�õ�Ƭ"" style=""cursor:default;color:black"">��</span>"
		   case "getpicturevotescore" echo "<script type=""text/Javascript"" src=""" & DomainStr & "Item/GetVote.asp?m=" & ModelID & "&ID=" & ItemID & """></script>"
		   case "getpicturevote" echo "<a href=""" & DomainStr & "Item/Vote.asp?m=" & ModelID & "&ID=" & ItemID & """>Ͷ��һƱ</a>"
		   case "getpicturedate" echo KS.DateFormat(GetNodeText("adddate"), 6)
           case "getprevpicture" echo LFCls.ReplacePrevNext(ModelID,ItemID, GetNodeText("tid"), "<")
           case "getnextpicture" echo LFCls.ReplacePrevNext(ModelID,ItemID, GetNodeText("tid"), ">")
		   '================================ͼƬģ�ͽ���================================
		   
		   
		   '================================����ģ�Ϳ�ʼ================================
		   case "getdowntitle"   echo GetNodeText("title") & " " & GetNodeText("downversion")
		   case "getdownaction"  echo "��<A href=""" & DomainStr & "plus/Comment.asp?ChannelID=" & ModelID & "&InfoID=" & ItemID & """ target=""_blank"">��������</A>����<A href=""" & DomainStr & "User/index.asp?User_Favorite.asp?Action=Add&ChannelID=" & ModelID & "&InfoID=" & ItemID & """ target=""_blank"">��Ҫ�ղ�</A>����<A href=""javascript:window.close();"">�رմ���</A>��"
		   case "getdownkeyword" echo Replace(GetNodeText("keywords"), "|", ",")
		   case "getdownurl"   echo KS.GetItemURL(ModelID,GetNodeText("tid"),ItemID,GetNodeText("fname"))
		   case "getdownsystem"   echo GetNodeText("downpt")
		   case "getdownauthor" echo LFCls.ReplaceDBNull(GetNodeText("author"),"����")
		   case "getdownorigin" echo KS.GetOrigin(LFCls.ReplaceDBNull(GetNodeText("origin"),"��վԭ��"))
		   case "getdownsize" echo GetNodeText("downsize")
		   case "getdowntype" echo GetNodeText("downlb")
		   case "getdownlanguage" echo GetNodeText("downyy")
		   case "getdownpower" echo GetNodeText("downsq")
		   case "getdownpoint" echo GetNodeText("readpoint")
		   case "getdowndecpass" echo GetNodeText("jymm")
		   case "getdownintro" echo KS.ReplaceInnerLink(GetNodeText("downcontent"))
		   case "getdownaddress" 
		     Dim UrlArr, I,N,TotalNum, AUrl
			 UrlArr = Split(GetNodeText("downurls"), "|||")
			 TotalNum = UBound(UrlArr)
			 For I = 0 To TotalNum
			    N=N+1: AUrl = Split(UrlArr(I), "|")
				If AUrl(0)=0 Then
				 echoln "<img src="""& DomainStr & "Images/Default/down.gif"" border=""0"" alt="""" align=""absmiddle"" /><a href=""" & DomainStr & "item/downLoad.asp?m=" & ModelID & "&id=" & ItemID & "&downid=" & N & """ target=""_blank"">" & AUrl(1) & "</a>"      
				 If I<>TotalNum Then echoln "<br/>"
				Else
				  Dim RS_S:Set RS_S=Conn.Execute("Select DownloadName,IsDisp,DownloadPath,DownID,SelFont From KS_DownSer Where ParentID=" & AUrl(0))
				  If RS_S.Eof Then
				    If TotalNum=0 Then UrlStr="<li>�ݲ��ṩ���ص�ַ</li>"
				  Else
				     DO While Not RS_S.Eof
					  IF RS_S(1)=1 Then
				      echoln "<img src="""& domainstr & "Images/Default/down.gif"" border=""0"" align=""absmiddle""><a href=""" & RS_S(2) & Aurl(2) & """ " & RS_S(4)&" target=""_blank"">" & RS_S(0) & "</a>"          
					  Else
				      echoln "<img src="""& domainstr & "Images/Default/down.gif"" border=""0"" align=""absmiddle""><a href=""" & DomainStr & "item/DownLoad.asp?m=" & ModelID & "&id=" & ItemID & "&DownID=" & N & "&Sid=" & RS_S(3) & """ " & RS_S(4)&" target=""_blank"">" & RS_S(0) & "</a>"        
					  End If
					  RS_S.MoveNext
					  IF Not RS_S.Eof Or I<>TotalNum Then echoln "<br/>" 
					 Loop
				  End If
				  RS_S.Close:Set RS_S=Nothing
				End If
			 Next
		   case "getdownlink"
				If Not (LCase(Node.SelectSingleNode("@ysdz").text) = "http://" Or Node.SelectSingleNode("@ysdz").text = "") Then  echo "<a href=""" & Node.SelectSingleNode("@ysdz").text & """ target=""_blank""><u>���߻򿪷�����ҳ</u></a>"
				If Not (LCase(Node.SelectSingleNode("@zcdz").text) = "http://" Or Node.SelectSingleNode("@zcdz").text = "") Then  echo "&nbsp;&nbsp;<a href=""" & Node.SelectSingleNode("@zcdz").text & """ target=""_blank""><u>ע���ַ</u></a>"
		   case "getdownysdz"
				If LCase(Node.SelectSingleNode("@ysdz").text) = "http://" Or Node.SelectSingleNode("@ysdz").text = "" Then
				   echo "��"
				Else
				   echo "<a href=""" & Node.SelectSingleNode("@ysdz").text & """ target=""_blank"">" & Node.SelectSingleNode("@ysdz").text & "</a>"
				End If
		   case "getdownzcdz"
				If LCase(Node.SelectSingleNode("@zcdz").text) = "http://" Or Node.SelectSingleNode("@zcdz").text = "" Then
				   echo "��"
				Else
				   echo "<a href=""" & Node.SelectSingleNode("@zcdz").text & """ target=""_blank"">" & Node.SelectSingleNode("@zcdz").text & "</a>"
				End If
		   case "getdownproperty"
		     If GetNodeText("recommend") = "1" Then Echo "<span title=""�Ƽ�"" style=""cursor:default;color:green"">��</span> "
			 If GetNodeText("popular") = "1" Then  echo "<span title=""����"" style=""cursor:default;color:red"">��</span> "
			 If GetNodeText("strip")="1" Then echo "<span title=""����ͷ��"" style=""cursor:default;color:#0000ff"">ͷ</span> "
			 If GetNodeText("rolls") = "1" Then echo "<span title=""����"" style=""cursor:default;color:#F709F7"">��</span> "
			 If GetNodeText("slide") = "1" Then echo "<span title=""�õ�Ƭ"" style=""cursor:default;color:black"">��</span>"
		   case "getdowndate" echo KS.DateFormat(GetNodeText("adddate"), 6)
		   case "getdowninput"   echo "<a href=""" & DomainStr & "Space/?" & GetNodeText("inputer") &""" target=""_blank"">" & GetNodeText("inputer") & "</a>"
           case "getprevdown" echo LFCls.ReplacePrevNext(ModelID,ItemID, GetNodeText("tid"), "<")
           case "getnextdown" echo LFCls.ReplacePrevNext(ModelID,ItemID, GetNodeText("tid"), ">")
		   '================================����ģ�Ϳ�ʼ================================
		   
		   case else
		       echo ShCls.run(sTemp)
		      If lcase(left(sTemp,3))="ks_" Then
			   echo GetNodeText(Lcase(sTemp))     '����Զ����ֶ�
			  ElseIf lcase(left(sTemp,3))="fl_" Then
			   echo GetNodeText(Lcase(right(sTemp,len(sTemp)-3)))     '��������ֶ�
			  elseIf left(lcase(sTemp),3)="js_" then
			   Call JsCls.Run(sTemp,Templates)
			  End If
		 end select
			Parse    = iPosBegin
			Set MyNode=Nothing
		End Function 
		
		'���͵Ⱥű�ǩ
		Function ParseEqual(sTemplate, iPosBegin)
			Dim iPosCur, sToken, sTemp,MyNode,TagName,TagParam,Param,PosTag,I
			iPosCur      = InStr(iPosBegin, sTemplate, "}")
			sTemp        = Mid(sTemplate,iPosBegin,iPosCur-iPosBegin)
			iPosBegin    = iPosCur+1
			
			PosTag       = InStr(sTemp,"(")
			If PosTag>0 Then
				TagName      = Mid(sTemp,1,PosTag-1)
				TagParam     = Replace(Replace(sTemp,")",""),TagName&"(","")
				'response.write (sTemp & "=" & tagParam)
				'response.end
				Param        = Split(TagParam,",")
				select case Lcase(TagName)
				 case "getlogo" echo "<img src=""" & KS.Setting(4) & """ border=""0"" width=""" & Param(0) & """ height=""" & Param(1) & """ align=""absmiddle"" alt=""logo"" />"
				 case "getadvertise" echo "<script src=""" & DomainStr & "plus/ShowA.asp?I="& TagParam & """ type=""text/javascript""></script>"
				 case "gettopuser" GetTopUser Param(0),Param(1)
				 case "getvote" echo GetVote(TagParam)
				 case "gettags" echo GetTags(Param(0),Param(1))
				 case "getuserdynamic" GetUserDynamic TagParam
				 
				 case "getphoto" echo "<div align=""center""><img src=""" & LFCls.ReplaceDBNull(GetNodeText("photourl"), DomainStr & "images/nopic.gif") & """  width=""" & Param(0) & """ height=""" & Param(1) & """ border=""0"" alt=""" & GetNodeText("title") &"""/></div>"
				 
				 case "getdownphoto" ,"getmoviephoto","getsupplyphoto"
				  Dim DownPhotoUrl:DownPhotoUrl=GetNodeText("photourl") : If DownPhotoUrl="" Or IsNull(DownPhotoUrl) Then DownPhotoUrl=DomainStr & "images/nopic.gif"
				  if Lcase(left(DownPhotoUrl,7))<>"http://" then DownPhotoUrl=KS.Setting(2) &DownPhotoUrl
				  echo "<img src=""" & DownPhotoUrl & """ height=""" & Param(1) & """ width=""" & Param(0) & """ alt=""" & GetNodeText("title") & """/>"
				
				
			 
			 case else
			   If left(lcase(TagName),3)="js_" then
			    Call JSCls.Equal(TagName,Param,Templates)	  
			   end if  
			 end select
		    End If
			ParseEqual   = iPosBegin
	   End Function

	
	  '�滻Ƶ��ר�ñ�ǩ
		Sub ParseChannelLabel(ByVal sTemp)
		   on error resume next
		   If FCls.RefreshFolderID="0" Or FCls.RefreshFolderID="" Then Exit Sub
		   	Dim I,ClassBasicInfoArr,ClassDefineContentArr
			ClassBasicInfoArr    = Split(KS.C_C(FCls.RefreshFolderID,6),"||||")
			ClassDefineContentArr= Split(KS.C_C(FCls.RefreshFolderID,7),"||||")

		    sTemp = Lcase(sTemp)
		   select case sTemp
		     case "getchannelid" echo Fcls.ChannelID
			 case "getchannelname" echo KS.C_S(FCls.ChannelID,1)
			 case "getitemname" echo KS.C_S(FCls.ChannelID,3)
			 case "getitemurl" echo KS.C_S(FCls.ChannelID,4)
			 
			 case "getclassid" echo FCls.RefreshFolderID
			 case "getparentid" echo FCls.RefreshParentID
			 case "getparenturl"  If FCls.RefreshParentID="0" Then echo KS.Setting(2) else echo KS.GetFolderPath(FCls.RefreshParentID)
			 case "getparentclassname" 
			 if FCls.RefreshType="Content" Then
			 echo KS.C_C(KS.C_C(FCls.RefreshFolderID,13),1)
			 Else
			 echo KS.C_C(FCls.RefreshParentID,1)
			 End If
			 case "getclassname" echo KS.C_C(FCls.RefreshFolderID,1)
			 case "getclassurl" echo KS.GetFolderPath(FCls.RefreshFolderID)
		  end select
		  
		  If IsArray(ClassBasicInfoArr) Then
		    select case sTemp
		     case "getclasspic" echo "<img src=""" & ClassBasicInfoArr(0) & """ border=""0"" alt="""" />"
			 case "getclassintro" echo ClassBasicInfoArr(1)
			 case "getclass_meta_keyword" echo ClassBasicInfoArr(2)
			 case "getclass_meta_description" echo ClassBasicInfoArr(3)
			end select
		  End If
		    
		  If IsArray(ClassDefineContentArr) Then
		     For I=1 To Ubound(ClassDefineContentArr)+1
			   if sTemp="getclassdefinecontent" & I  then echo ClassDefineContentArr(I-1)
			 Next
		  End If
		  if err then err.clear
		End Sub
		
		'�滻RSS��ǩ
		Sub ParseRssLabel(sTemp)
		   IF KS.Setting(83)=0 Then Exit Sub
		   Dim CurrentClassID:CurrentClassID=FCls.RefreshFolderID
		   Dim ChannelID:ChannelID=FCls.ChannelID
		   select case Lcase(sTemp)
		      case "rss" 
			    select case Lcase(FCls.RefreshType)
				 case "index" echo GetRssLink("rss.asp")
				 case "folder" echo GetRssLink("Rss.asp?ChannelID=" & ChannelID & "&ClassID=" &CurrentClassID & "")
			    end select
			 case "rsselite"
			    select case Lcase(FCls.RefreshType)
				 case "index" echo GetRssLink("Rss.asp?Elite=1")
				 case "folder" echo GetRssLink("Rss.asp?ChannelID=" & ChannelID & "&ClassID=" &CurrentClassID & "&Elite=1")
			    end select
			 case "rsshot"
			    select case Lcase(FCls.RefreshType)
				 case "index" echo GetRssLink("Rss.asp?Hot=1")
				 case "folder" echo GetRssLink("Rss.asp?ChannelID=" & ChannelID & "&ClassID=" &CurrentClassID & "&Hot=1")
			    end select
		   end select
		End Sub
		'ȡ��ÿ��Ƶ����RSS���ӣ����ParseRssLabel����
		Function GetRssLink(LinkStr)
		   GetRssLink="<a href=""" & DomainStr & LinkStr & """ target=""_blank""><img src=""" & DomainStr & "Images/Rss.gif" & """ border=""0""></a>"
		End Function
		
		
		'*******************************************************************************************************
		'��������KSLabelReplaceAll
		'��  �ã��滻���б�ǩ
		'��  ����F_C ģ������
		'����ֵ���滻����ģ������
		'********************************************************************************************************
		Public Function KSLabelReplaceAll(F_C)
		          F_C = ReplaceAllLabel(F_C)                    
				  F_C = ReplaceLableFlag(F_C)                   '�滻������ǩ
				  F_C = ReplaceGeneralLabelContent(F_C)        '�滻ͨ�ñ�ǩ ��{$GetWebmaster}
				  F_C = ReplaceRA(F_C, "")
				  KSLabelReplaceAll=F_C
	    End Function
		'*******************************************************************************************************
		'��������LoadTemplate
		'��  �ã�ȡ��ģ������
		'��  ����TemplateFnameģ���ַ
		'����ֵ��ģ������
		'********************************************************************************************************
		Function LoadTemplate(TemplateFname)
		    on error resume next
			Dim FSO, FileObj, FileStreamObj 
			Set FSO = KS.InitialObject(KS.Setting(99))
			  TemplateFname=Replace(TemplateFname,"{@TemplateDir}",KS.Setting(3) & KS.Setting(90))
			  TemplateFname = Server.MapPath(Replace(TemplateFname, "//", "/"))
			  If FSO.FileExists(TemplateFname) = False Then
				LoadTemplate = "ģ�岻����,���Ȱ�!"
			  Else
				Set FileObj = FSO.GetFile(TemplateFname)
				Set FileStreamObj = FileObj.OpenAsTextStream(1)
				If Not FileStreamObj.AtEndOfStream Then
					LoadTemplate = FileStreamObj.ReadAll
				Else
					LoadTemplate = "ģ������Ϊ��"
				End If
			  End If
			  Set FSO = Nothing:Set FileObj = Nothing:Set FileStreamObj = Nothing
			  LoadTemplate=LoadTemplate & Published
		End Function
		'**************************************************
		'��������ReplaceLableFlag
		'��  �ã�ȥ����ǩ{$},�������Խ���ǩ������","����
		'          ʾ��: km=ReplaceLableFlag("{$Test("par1","par2","par3")}")
		'          ���     km=Test,Par1,Par2,Par3
		'��  ���� Content  ----���滻����
		'����ֵ��������","�������ַ���
		'**************************************************
		Function ReplaceLableFlag(Content)
			Dim regEx, Matches, Match, TempStr
			Set regEx = New RegExp
			ReplaceLableFlag = Content
			Set regEx = New RegExp
			regEx.Pattern = "{Tag([\s\S]*?):(.+?)}([\s\S]*?){/Tag\1}"
			regEx.IgnoreCase = True
			regEx.Global = True
			Set Matches = regEx.Execute(Content)
			For Each Match In Matches
				ReplaceLableFlag = Replace(ReplaceLableFlag,Match.Value,KSLabel.GetLabel(Match.Value))
			Next
		End Function
		
		
		'ɨ��ϵͳ������ǩ
		Function ScanSysLabel(Content)
		  Dim iPosLast, iPosCur,Tstr
			iPosLast    = 1
			Do While True 
				iPosCur    = InStr(iPosLast, Content, "{LB_") 
				If iPosCur>0 Then
					Tstr=tstr &  Mid(Content, iPosLast, iPosCur-iPosLast)
					iPosLast  = ParseSysLabel(Content, iPosCur+4,Tstr)
				Else 
					Tstr=tstr & Mid(Content, iPosLast)
					Exit do
				End If 
		   Loop 
		   ScanSysLabel=Tstr
		End Function
		Function ParseSysLabel(sTemplate, iPosBegin,Tstr)
			Dim iPosCur, sToken, sTemp,MyNode
			iPosCur      = InStr(iPosBegin, sTemplate, "}")
			sTemp        = Mid(sTemplate,iPosBegin,iPosCur-iPosBegin)
			iPosBegin    = iPosCur+1
			Set MyNode   = Application(KS.SiteSN&"_labellist").documentElement.SelectSingleNode("labellist[@labelname='{LB_" & sTemp & "}']")
			If Not MyNode Is Nothing Then Tstr=Tstr &  MyNode.text 
			
			ParseSysLabel= iPosBegin
		End Function
		
		
		'*********************************************************************************************************
		'��������ReplaceAllLabel
		'��  �ã�����ǩ����ת���ɶ�Ӧ��ǩ����
		'��  ���� Content��ת��������
		'*********************************************************************************************************
		Function ReplaceAllLabel(Content)
			dim Node
			Call LoadLabelToCache()    '���ر�ǩ
			
		    Content=ScanSysLabel(Content)

			Call LoadJSFileToCache()   '����JS
			For Each Node in Application(KS.SiteSN&"_jslist").documentElement.SelectNodes("jslist")
				Content=Replace(Content,Node.selectSingleNode("@jsname").text,Node.text)
			Next
			If Lcase(Fcls.RefreshType)<>"content" Then Content=ReplaceSQLLabel(Content)

			ReplaceAllLabel=Content
		End Function
		
		Function ReplaceSQLLabel(Content)
			'�滻�Զ��庯����ǩ 
			Dim DCls:Set Dcls=New DIYCls
			ReplaceSQLLabel=DCls.ReplaceUserFunctionLabel(Content)
			Set DCls=nothing
		End Function

	
		'�������ݿ�����б�ǩ������	
		 Sub LoadLabelToCache()
			If Not IsObject(Application(KS.SiteSN&"_labellist")) Then
					Set  Application(KS.SiteSN&"_labellist")=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
					Application(KS.SiteSN&"_labellist").appendChild(Application(KS.SiteSN&"_labellist").createElement("xml"))
						Dim i,SQL,Node
						Dim RS:Set RS = Server.CreateObject("ADODB.Recordset")
						RS.Open "Select ID,LabelType,LabelName,LabelContent from KS_Label Where LabelType<>5", Conn, 1, 1
						If Not RS.Eof Then SQL=RS.GetRows(-1)
						RS.Close:Set RS = Nothing
						If IsArray(SQL) Then
							for i=0 to Ubound(SQL,2)
								 Set Node=Application(KS.SiteSN&"_labellist").documentElement.appendChild(Application(KS.SiteSN&"_labellist").createNode(1,"labellist",""))
								 Node.attributes.setNamedItem(Application(KS.SiteSN&"_labellist").createNode(2,"labelname","")).text=SQL(2,I)
								 Node.attributes.setNamedItem(Application(KS.SiteSN&"_labellist").createNode(2,"labelid","")).text=SQL(0,I)
								If SQL(1,I) = 1 Then
								 Node.text=ReplaceFreeLabel(SQL(3,I))
								Else
								 Node.text=Replace(SQL(3,I),"labelid=""0""","labelid=""" & SQL(0,I) & """")
								End IF
							next
						End If
			End if
		End Sub
		
		'�������ݿ������JS������
		Sub LoadJSFileToCache()
			If Not IsObject(Application(KS.SiteSN&"_jslist")) Then
					Set  Application(KS.SiteSN&"_jslist")=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
					Application(KS.SiteSN&"_jslist").appendChild( Application(KS.SiteSN&"_jslist").createElement("xml"))
						Dim i,SQL,Node
						Dim RS:Set RS = Server.CreateObject("ADODB.Recordset")
						RS.Open "Select JSID,JSName,JSFileName from KS_JSFile", Conn, 1, 1
						If Not RS.Eof Then SQL=RS.GetRows(-1)
						RS.Close:Set RS = Nothing
						If IsArray(SQL) Then
							for i=0 to Ubound(SQL,2)
								 Set Node=Application(KS.SiteSN&"_jslist").documentElement.appendChild(Application(KS.SiteSN&"_jslist").createNode(1,"jslist",""))
								 Node.attributes.setNamedItem(Application(KS.SiteSN&"_jslist").createNode(2,"jsname","")).text=SQL(1,I)
								 Node.text="<script charset=""gb2312"" language=""javascript"" src=""" & Replace(KS.Setting(3) & KS.Setting(93),"//","/") & Trim(SQL(2,I)) & """></script>"
							next
						End If
			End if
		End Sub

	'�滻���ɱ�ǩΪ����,���滻һ��
	Function ReplaceFreeLabel(sTrC)
			dim node
			If not IsObject(Application(KS.SiteSN&"_ReplaceFreeLabel")) then
					Dim RS:Set RS = Server.CreateObject("ADODB.Recordset")
					RS.Open "Select LabelName,LabelContent,ID from KS_Label", Conn, 1, 1
					if Not RS.eof then
						'KS.Value=RS.GetString(,,"^||^","^%%%^","")
						Set Application(KS.SiteSN&"_ReplaceFreeLabel")=KS.ArrayToXml(RS.GetRows(-1),rs,"row","")
					end if
					RS.Close:Set RS = Nothing

			End if
			For Each Node In Application(KS.SiteSN&"_ReplaceFreeLabel").documentElement.SelectNodes("row")
					sTrC = Replace(sTrC,trim(Node.SelectSingleNode("@labelname").text),Replace(Node.SelectSingleNode("@labelcontent").text,")}","," & Node.SelectSingleNode("@id").text &")}"))
			next
			'ReplaceFreeLabel = ReplaceGeneralLabelContent(sTrC)
			ReplaceFreeLabel = ScanSysLabel(sTrC)
		End Function

		'*********************************************************************************************************
		'��������FSOSaveFile
		'��  �ã������ļ�
		'��  ���� Content����,·�� ע������Ŀ¼
		'*********************************************************************************************************
		Sub FSOSaveFile(Content, FileName)
			dim stm:set stm=server.CreateObject("adodb.stream")
			stm.Type=2 '���ı�ģʽ��ȡ
			stm.mode=3
			stm.charset="gb2312"
			stm.open
			stm.WriteText content
			stm.SaveToFile server.MapPath(FileName),2 
			stm.flush
			stm.Close
			set stm=nothing
		End Sub
		
		'*********************************************************************************************************
		'��������RefreshJS
		'��  �ã�����JS
		'��  ����JSName JS����
		'*********************************************************************************************************
		Sub RefreshJS(JSName)
			Dim JSRS, SqlStr, JSContent
			Set JSRS = Server.CreateObject("ADODB.Recordset")
			SqlStr = "Select * From KS_JSFile Where JSName='" & Trim(JSName) & "'"
			JSRS.Open SqlStr, Conn, 1, 1
			If JSRS.EOF And JSRS.BOF Then
			 JSRS.Close:Set JSRS = Nothing:Exit Sub
			End If
			  Dim JSConfig, JSFileName, SaveFilePath, JSDir, JSType
			  JSFileName = Trim(JSRS("JSFileName"))
			  JSDir = Trim(KS.Setting(93))
			  JSType = Trim(JSRS("JSType"))
			  If Left(JSDir, 1) = "/" Or Left(JSDir, 1) = "\" Then JSDir = Right(JSDir, Len(JSDir) - 1)
			  SaveFilePath = KS.Setting(3) & JSDir
			  Call KS.CreateListFolder(SaveFilePath)
			   
			   JSConfig = Trim(JSRS("JSConfig"))
			  If JSType = "0" Then
				JSContent=Replace(Replace(Replace(Replace(KSLabel.GetLabel(JSConfig), Chr(13)& Chr(10), ""),"'","\'"),"""","\"""),vbcrlf,"")             
				JSContent=Replace(JSContent,Chr(13) ,"")
				JSContent = "document.write('" & JSContent & "');"
			  Else
				Dim FreeType
				FreeType = Left(JSConfig, InStr(JSConfig, ",") - 1) 'ȡ������JS������
				JSConfig = Replace(JSConfig, FreeType & ",", "")
				
				Select Case FreeType      '���ݺ�������Ӧ�Ĳ���
				  Case "GetExtJS"          '��չJS
					 JSConfig = Replace(JSConfig, "'", """")
					 JSConfig = ReplaceLableFlag(ReplaceAllLabel(JSConfig))
					 JSConfig = ReplaceGeneralLabelContent(JSConfig)
					 JSConfig = Replace(Replace(Replace(JSConfig, Published, ""),"'","\'"),"""","\""")
					 JSContent = ReplaceJsBr(JSConfig)
				  Case "GetWordJS"
					 JSConfig = Replace(Trim(JSConfig), """", "")   '�滻ԭ������˫����Ϊ��
					 JSContent = RefreshWordJS(Trim(JSRS("JSID")), JSConfig)           '�滻����JS
				  Case Else
					 JSContent = ""
				End Select
			End If
			  Call FSOSaveFile(JSContent, SaveFilePath & JSFileName)
			 JSRS.Close:Set JSRS = Nothing
		End Sub
		Function ReplaceJsBr(Content)
		 Dim i
		 Dim JsArr:JSArr=Split(Content,Chr(13) & Chr(10))
		 For I=0 To Ubound(JsArr)
		   ReplaceJsBr=ReplaceJsBr & "document.writeln('" & JsArr(I) &"')" & vbcrlf 
		 Next
		End Function
		'*********************************************************************************************************
		'��������RefreshWordJS
		'��  �ã���������JS
		'��  ����JSID JSID,JSConfig JS����
		'*********************************************************************************************************
		Function RefreshWordJS(JSID, JSConfig)
		     Dim JSConfigArr:JSConfigArr = Split(JSConfig, ",")
			 If UBound(JSConfigArr) = 17 Then
					RefreshWordJS = KSLabel.RefreshCss(JSID, UCase(JSConfigArr(0)), JSConfigArr(1), JSConfigArr(2), JSConfigArr(3), JSConfigArr(4), JSConfigArr(5), JSConfigArr(6), JSConfigArr(7), JSConfigArr(8), JSConfigArr(9), JSConfigArr(10), JSConfigArr(11), JSConfigArr(12), JSConfigArr(13), JSConfigArr(14), JSConfigArr(15), JSConfigArr(16), JSConfigArr(17))
					RefreshWordJS = Replace(RefreshWordJS, "'", """")
					RefreshWordJS = "document.write('" & RefreshWordJS & "');"
			 Else
					RefreshWordJS = "document.write('��ǩ���������');"
			 End If
		End Function
		
		'=================================����Ϊ�����Ŀ,����ҳ,Ƶ����ҳ�ȵ�ˢ�º���=====================================
		
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'��������RefreshContent
		'��  �ã�ˢ������ҳ��
		'��  ���� ��
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Function RefreshContent()
			 Dim TFileContent, F_C, FilePath, FilePathAndName, FilePathAndNameTemp, sFname,Fname, FExt, TempFileContent, Content, ContentArr, I, N, CurrPage, PageStr, Flag
			 Dim TemplateID
			   TID = Trim(Node.SelectSingleNode("@tid").text)
			   Call FCls.SetContentInfo(ModelID,Tid,ItemID,Node.SelectSingleNode("@title").text)
			    If ModelID=8 Then
			     TemplateID      = KS.C_C(Tid,5)
				Else
				 TemplateID      = Node.SelectSingleNode("@templateid").text
				End If
			   
			   TempFileContent = LoadTemplate(TemplateID)
			   TempFileContent = ReplaceAllLabel(TempFileContent)
               
			   If InStr(TempFileContent, "{Tag:GetRelativeList") <> 0 Then TempFileContent = Replace(TempFileContent, "{Tag:GetRelativeList", "{UnTag:GetRelativeList"):Flag = True  Else Flag = False

			   If Flag = True Then
				TFileContent = ReplaceLableFlag(TempFileContent)
			   ElseIf (TemplateID <> FCls.RefreshTemplateID) Or (Tid <> FCls.RefreshCurrTid) Or FCls.RefreshTempFileContent = "" Then
				FCls.RefreshCurrTid = Tid
				FCls.RefreshTemplateID = TemplateID
				FCls.RefreshTempFileContent = ReplaceLableFlag(TempFileContent)  '�滻������ǩ
				TFileContent = FCls.RefreshTempFileContent
			   Else
				TFileContent = FCls.RefreshTempFileContent
			   End If
			  

			  on error resume next
			  sFname = Trim(Node.SelectSingleNode("@fname").text)
			  FExt   = Mid(sFname, InStrRev(sFname, ".")) '�������չ��
			  Fname = Replace(sFname, FExt, "")  '������ļ��� �� 2005/9-10/1254ddd
			  
			  FilePathAndNameTemp =KS.LoadFsoContentRule(ModelID,Tid)
			  Dim ShowUrl:ShowUrl=KS.LoadInfoUrl(ModelID,Tid,"")
			  FilePathAndName = FilePathAndNameTemp & sFname
			  FilePath = Replace(FilePathAndName, Mid(FilePathAndName, InStrRev(FilePathAndName, "/")), "")
			  
			  Call KS.CreateListFolder(FilePath)
			  
			  '�ж��ǲ���ת������
			  If KS.C_S(ModelID,6)=1 Then
			    if node.SelectSingleNode("@changes").text="1" then
				 Templates=""
				  echoln "<script type=""text/javascript"">"
				  echoln "<!--"
				  echoln " location.href='" & Node.SelectSingleNode("@articlecontent").text & "';"
				  echoln "//-->"
				  echoln "</script>"
				 Call FSOSaveFile(Templates, FilePathAndName)
				 Exit Function
				end If
			  End If
			  '�ж��ǲ����շ���Ϣ
			  IF KS.C_S(ModelID,6)=1 or KS.C_S(ModelID,6)=2 or KS.C_S(ModelID,6)=4 Then
			    If Node.SelectSingleNode("@readpoint").text>0 or Node.SelectSingleNode("@infopurview").text="2" Or (Node.SelectSingleNode("@infopurview").text=0 And (KS.C_C(Tid,3)=1 Or KS.C_C(Tid,3)=2)) Then
				  Templates=""
				  echoln "<script type=""text/javascript"">"
				  echoln "<!--"
				  echoln "  location.href='" & KS.Setting(3) & "item/show.asp?m=" & ModelID & "&d=" & ItemID &"';"
				  echoln "//-->"
				  echoln "</script>"
				  Call FSOSaveFile(Templates, FilePathAndName)
				 Exit Function
				End If
			  End If
			  
			  
			  
			  Dim PrevUrl,StartPage,K
			  Select Case Cint(KS.C_S(ModelID,6))
			  Case 1   '����ģ��
					  Content = Node.SelectSingleNode("@articlecontent").text
					  If IsNull(Content) or Content="" Then Content = " "
					  ContentArr = Split(Content, "[NextPage]")
					  TotalPage = UBound(ContentArr) + 1
					  
					  For I = 0 To UBound(ContentArr)
					   CurrPage = I + 1
					   If TotalPage > 1 Then
							   If I = 0 Then
								 NextUrl = ShowUrl & Fname & "_" & (CurrPage + 1) & FExt : PrevUrl=""
							   ElseIf I = 1 And I <> TotalPage - 1 Then '�������һҳ�պ��ǵڶ�ҳ��Ҫ�����⴦��
								 NextUrl = ShowUrl & Fname & "_" & (CurrPage + 1) & FExt : PrevUrl = ShowUrl & sFname
							   ElseIf I = 1 And I = TotalPage - 1 Then
								 NextUrl=KS.GetFolderPath(Tid): PrevUrl = ShowUrl & sFname
							   ElseIf I = TotalPage - 1 Then
								 NextUrl=KS.GetFolderPath(Tid): PrevUrl = ShowUrl & Fname & "_" & (CurrPage - 1) & FExt
							   Else
								 NextUrl = ShowUrl & Fname & "_" & (CurrPage + 1) & FExt : PrevUrl = ShowUrl & Fname & "_" & (CurrPage - 1) & FExt
							   End If
							   PageStr = "<div id=""pageNext"" style=""text-align:center""><table align=""center""><tr><td>"
							   If CurrPage > 1 And PrevUrl<>"" Then PageStr = PageStr & "<a class=""prev"" href=""" & PrevUrl & """>��һҳ</a> "
							   
						     startpage=1:k=0: if (CurrPage>=10) then startpage=(CurrPage\10-1)*10+CurrPage mod 10+2
							   
						    For N = startpage To TotalPage
								 If CurrPage = N Then
								   PageStr = PageStr & ("<a class=""curr"" href=""#""><span style=""color:red"">" & N & "</span></a> ")
								 Else
								   If N=1 Then
								   	PageStr = PageStr & ("<a class=""num"" href="""  & ShowUrl & sFname & """>" & N & "</a> ")
								   Else
								    PageStr = PageStr & ("<a class=""num"" href=""" &  ShowUrl & Fname & "_" & N & FExt & """>" & N & "</a> ")
								   End If
								End If
								K=K+1 : If K>=10 Then Exit For
							Next
							If CurrPage<>TotalPage Then PageStr = PageStr & "<a class=""next"" href=""" & NextUrl & """>��һҳ</a>"
							PageStr = PageStr & "</td></tr></table></div>"
						 Else
						  NextUrl=KS.GetFolderPath(Tid): PageStr = ""
						 End If
						F_C = TFileContent
					   If CurrPage <> 1 Then FilePathAndName = FilePathAndNameTemp & Fname & "_" & CurrPage & FExt
						Dim PageTitleArr,PageTitle
						PageTitle=Node.SelectSingleNode("@pagetitle").text
						If PageTitle<>"" And Not IsNull(PageTitle) Then
							  PageTitleArr=Split(PageTitle,"��")
							  If CurrPage-1<=Ubound(PageTitleArr) Then
							   F_C=Replace(F_C,"{$GetArticleTitle}",PageTitleArr(CurrPage-1))
							  End If
						End IF
						
					   If InStr(F_C, "{UnTag:GetRelativeList") <> 0 Then F_C = ReplaceLableFlag(Replace(F_C, "{UnTag:GetRelativeList", "{Tag:GetRelativeList"))
					   PageContent=ContentArr(I) & PageStr
					   Templates = ""
					   Scan F_C
					   F_C = Templates
					   F_C = Replace(Replace(F_C,"[KS_Charge]",""),"[/KS_Charge]","")
					   
					   If Instr(F_C,"[KS_ShowIntro]")<>0 Then
							  If CurrPage=1 Then
								F_C=Replace(Replace(F_C,"[KS_ShowIntro]",""),"[/KS_ShowIntro]","")
							  Else
								F_C=Replace(F_C,KS.CutFixContent(F_C, "[KS_ShowIntro]", "[/KS_ShowIntro]", 1),"")
							  End If
					   End If

					   
					   F_C = ReplaceGeneralLabelContent(F_C)
					   F_C = ReplaceRA(F_C, Trim(KS.C_C(Tid,4))) 
					   F_C = Replace(Replace(Replace(Replace(F_C,"{��","{$"),"{#LB","{LB"),"{#SQL","{SQL"),"{#=","{=")
					   Call FSOSaveFile(F_C, FilePathAndName)
					Next
			case 2  'ͼƬģ��
					  Content=Node.SelectSingleNode("@picurls").text
					  If IsNull(Content) Then Content = "" 
					  ContentArr = Split(Content, "|||") : TotalPage  = UBound(ContentArr) + 1
					  If InStr(TempFileContent, "{=GetPhotoPage") <> 0 Then
								Dim HtmlLabel:HtmlLabel = KSLabel.GetFunctionLabel(TempFileContent, "{=GetPhotoPage")
								Dim Param:Param = split(KSLabel.GetFunctionLabelParam(HtmlLabel, "{=GetPhotoPage"),",")
								Dim Rows:Rows=Param(0)
								Dim Cols:Cols=Param(1)
								Dim Width:Width=Param(2)
								Dim Height:Height=Param(3)
								Dim r,c,str
								if ((ubound(ContentArr)+1) mod (cols*rows))=0 then TotalPage=(ubound(ContentArr)+1)\(cols*rows)	else TotalPage=(ubound(ContentArr)+1)\(cols*rows) + 1
									
							 For I = 1 To TotalPage
								 str="<table cellspacing=""20"" cellpadding=""0"" align=""center"" border=""0"">"
								 if TotalPage<=1 then n=0 else n=(cols*rows)*(I-1)
								For r=1 to rows
								  str=str & "<tr>"
								 For c=1 To Cols
									  dim thumbsphoto
									  if n<=ubound(ContentArr) Then
										PicSrc=Split(ContentArr(n), "|")(2)
										If (Lcase(Left(PicSrc,4))<>"http") Then PicSrc=KS.Setting(2) & PicSrc
									   thumbsphoto="<table cellspacing=""0"" cellpadding=""0"" width=""100%"" align=""center"" border=""0""><tr><td style='border:1px #999999 solid;background:#FFFFFF;padding:10px;text-align:center'><a id="""" href=""" & Split(ContentArr(n), "|")(1) & """  class=""highslide"" onclick=""return hs.expand(this)"" title=""""><img alt='" & Split(ContentArr(n), "|")(0) & "' width='" & width &"' height='" & height & "' src='" & Split(ContentArr(n), "|")(2)  & "' style='border:1px #999999 solid' border='0'></a><div style='text-align:center'>" & Split(ContentArr(n), "|")(0) & "</div></td></tr></table>"
									  else
									   thumbsphoto=""
									  end if
									  str=str & "<td valign=""top"">" & thumbsphoto & "</td>"
									  n=n+1
								 Next
								 str=str & "</tr>"
								Next
								 str=str &"</table>"
													 
								 
								PageStr="<table style=""BORDER-BOTTOM: #8eacca 1px solid"" cellSpacing=""0"" cellPadding=""0"" width=""95%"" align=""center"" border=""0""><tr><td width=""54%"" height=""25"">���� <font color=""#6699ff""><strong>" & TotalPage &" </strong></font>ҳ �� <font color=""#6699ff""><strong>" & I & "</strong></font> ҳ</td><td align=""right"" width=""33%"">"
								
									 startpage=1:k=0
									 if (I>=10) then startpage=(I\10-1)*10+I mod 10+2
									 
									  PageStr=PageStr & "<a href=""" & ShowUrl & sFname & """ title=""��ҳ"">��ҳ</a> "
									  if I<>1 then 
										if I=2 then
										 PageStr=PageStr & "<a href=""" & ShowUrl & sFname & """ title=""��һҳ""><<</a> "
										else
										 PageStr=PageStr & "<a href=""" & ShowUrl & Fname & "_" & I-1 & FExt & """ title=""��һҳ""><<</a> "
										end if
									  end if
			
								  For N = cint(startpage) To TotalPage
									 If N = 1 Then
									   If I = N Then
										 PageStr = PageStr & "<a href=""#""><font color=""red"">" & N & "</font></a>&nbsp;"
									   Else
										PageStr = PageStr & "<a href=" & ShowUrl & sFname & ">" & N & "</a>&nbsp;"
									  End If
									Else
									   If I = N Then
										  PageStr = PageStr & "<a href=""#""><font color=""red"">" & N & "</font></a>&nbsp;"
									   Else
										 PageStr = PageStr & "<a href=" & ShowUrl & Fname & "_" & N & FExt & ">" & N & "</a>&nbsp;"
									   End If
									End If
									
									k=K+1
									If k >= 10 Then exit for
								Next
								
									If I <>totalpage Then
									PageStr=PageStr & "<a href=""" & ShowUrl & Fname & "_" & I+1 & FExt & """ title=""��һҳ"">>></a> "
									end if
									PageStr=PageStr & "<a href=""" & ShowUrl &  Fname & "_" & TotalPage & FExt & """ title=""ĩҳ"">ĩҳ</a> "
									PageStr=PageStr & "</div>"
								
								PageStr=PageStr & "</td><td align=""right"" width=""13%""><Select style=""color: #6699ff"" onchange=""javascript:window.location=this.value;"" name=""nPage"">" 
								For K=1 To TotalPage
								 if k=I then
								   if k=1 then
									PageStr=PageStr & "<Option value='" & ShowUrl & sFname & "' selected>��" & K & "ҳ</Option>"
								   else
									PageStr=PageStr & "<Option value='" & ShowUrl & Fname & "_" & k & FExt & "' selected>��" & K & "ҳ</Option>"
								   end if
								 else
								   if k=1 then
									PageStr=PageStr & "<Option value='" & ShowUrl & sFname & "'>��" & K & "ҳ</Option>"
									else
									PageStr=PageStr & "<Option value='" &  ShowUrl & Fname & "_" & k & FExt & "'>��" & K & "ҳ</Option>"
									end if
								 end if
								Next
								PageStr=PageStr & "</Select> </td></tr></table>"
								 
								 If I <> 1 Then	FilePathAndName = FilePathAndNameTemp & Fname & "_" & I & FExt
								 
							  F_C = TFileContent
							  F_C=Replace(F_C, HtmlLabel,str & Replace(LFCls.GetConfigFromXML("highslide","/labeltemplate/label","highslide"),"{$GetInstallDir}",DomainStr))
							  F_C=Replace(F_C,"{$PageStr}",PageStr)
								 
							  If InStr(F_C, "{UnTag:GetRelativeList") <> 0 Then F_C = ReplaceLableFlag(Replace(F_C, "{UnTag:GetRelativeList", "{Tag:GetRelativeList"))
							 
							 Templates = ""
							 Scan F_C
							 F_C = Templates
							 
							 F_C = ReplacePictureContent(ChannelID,RS, F_C, "")
							 F_C = ReplaceRA(F_C, Trim(KS.C_C(Tid,4)))
							 Call FSOSaveFile(F_C, FilePathAndName)
							Next
					  ElseIf InStr(TempFileContent, "{$GetPictureByPage}") <> 0 Then   '����ҳ��ʽ����ͼƬ����ҳ
					   For I = 0 To TotalPage - 1
						CurrPage = I + 1
	
						If TotalPage > 1 Then
							PageStr="<div class=""kspage"">" & vbcrlf & "<div style=""text-align:center"">"
							If I = 0 Then
							  PageStr = PageStr & "<a href=" & ShowUrl & Fname & "_" & (CurrPage + 1) & FExt & ">��һ��>></a><br>"
							  NextUrl=ShowUrl & Fname & "_" & (CurrPage + 1) & FExt
							ElseIf I = 1 And I <> TotalPage - 1 Then '�������һ�Ÿպ��ǵڶ��ŵ�Ҫ�����⴦��
							  PageStr = PageStr &"<a href=" & ShowUrl & sFname & "><<��һ��</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=" & ShowUrl & Fname & "_" & (CurrPage + 1) & FExt & ">��һ��>></a><br>"
							  NextUrl=ShowUrl & Fname & "_" & (CurrPage + 1) & FExt
							ElseIf I = 1 And I = TotalPage - 1 Then
							  PageStr = PageStr &"<a href=" & ShowUrl & sFname & "><<��һ��</a><br>"
							  NextUrl=ShowUrl & sFname
							ElseIf I = TotalPage - 1 Then
							  PageStr = PageStr &"<a href=" & ShowUrl & Fname & "_" & (CurrPage - 1) & FExt & "><<��һ��</a>"
							  NextUrl=ShowUrl & sFname
							Else
							  PageStr = PageStr &"<a href=" &ShowUrl & Fname & "_" & (CurrPage - 1) & FExt & "><<��һ��</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=" &ShowUrl & Fname & "_" & (CurrPage + 1) & FExt & ">��һ��>></a>"
							  NextUrl=ShowUrl & Fname & "_" & (CurrPage + 1) & FExt
							End If
							PageStr =PageStr & "</div>"
							   
							PageStr = PageStr & "<br /><div style=""text-align:left"">" & Split(ContentArr(CurrPage-1), "|")(0) & "</div>"
							 startpage=1:k=0 : if (CurrPage>=10) then startpage=(CurrPage\10-1)*10+CurrPage mod 10+2
							 PageStr = PageStr & "<br /><div style=""text-align:center""><a href=""#"">��<font color=""red""> " & I+1 & "/" & TotalPage & "</font> ��</a>&nbsp;&nbsp;"
	
							  PageStr=PageStr & "<a href=""" & ShowUrl & sFname & """ title=""��ҳ"">��ҳ</a> "
							  if CurrPage<>1 then 
								if currpage=2 then
								 PageStr=PageStr & "<a href=""" & ShowUrl & sFname & """ title=""��һҳ""><<</a> "
								else
								 PageStr=PageStr & "<a href=""" & ShowUrl & Fname & "_" & CurrPage-1 & FExt & """ title=""��һҳ""><<</a> "
								end if
							  end if
	
						  For N = cint(startpage) To TotalPage
							 If N = 1 Then
							   If CurrPage = N Then
								 PageStr = PageStr & "<a href=""#""><font color=""red"">" & N & "</font></a>&nbsp;"
							   Else
								PageStr = PageStr & "<a href=" & ShowUrl & sFname & ">" & N & "</a>&nbsp;"
							  End If
							Else
							   If CurrPage = N Then
								  PageStr = PageStr & "<a href=""#""><font color=""red"">" & N & "</font></a>&nbsp;"
							   Else
								 PageStr = PageStr & "<a href=" & ShowUrl & Fname & "_" & N & FExt & ">" & N & "</a>&nbsp;"
							   End If
							End If
	
							k=K+1:If k >= 10 Then exit for
						 Next
						
						 If CurrPage <>totalpage Then PageStr=PageStr & "<a href=""" & ShowUrl & Fname & "_" & CurrPage+1 & FExt & """ title=""��һҳ"">>></a> "
						PageStr=PageStr & "<a href=""" & ShowUrl &  Fname & "_" & TotalPage & FExt & """ title=""ĩҳ"">ĩҳ</a> "
						PageStr=PageStr & "</div>"
					  Else
						NextUrl="#"
						PageStr = ""
					  End If
					
					  If CurrPage <> 1 Then	FilePathAndName = FilePathAndNameTemp & Fname & "_" & CurrPage & FExt
					
						F_C = TFileContent
						If InStr(F_C, "{UnTag:GetRelativeList") <> 0 Then F_C = ReplaceLableFlag(Replace(F_C, "{UnTag:GetRelativeList", "{Tag:GetRelativeList"))
						Dim PicSrc :PicSrc=Split(ContentArr(I), "|")(1)
						If (Lcase(Left(PicSrc,4))<>"http") Then PicSrc=KS.Setting(2) & PicSrc
						If NextUrl="" Then NextUrl=KS.GetFolderPath(Tid)
						
						PageContent="<div align=""center""><a href=""" & NextUrl & """><Img onmousewheel=""return bbimg(this)"" onload=""javascript:resizepic(this)"" src="""& PicSrc & """ border=""0""></a></div>" & PageStr
						If TotalPage > 1 Then
						 PageContent=PageContent & "</div>"
						End If
						Templates = ""
						Scan F_C
						F_C = Templates
						F_C=Replace(Replace(F_C,"[KS_Charge]",""),"[/KS_Charge]","")
						F_C = ReplaceGeneralLabelContent(F_C)
						F_C = ReplaceRA(F_C, Trim(KS.C_C(Tid,4))) 
						Call FSOSaveFile(F_C, FilePathAndName)
					 Next
				  Else               'ͼƬ��������ʽ
					   F_C = TFileContent
					   If InStr(F_C, "{UnTag:GetRelativeList") <> 0 Then F_C = ReplaceLableFlag(Replace(F_C, "{UnTag:GetRelativeList", "{Tag:GetRelativeList"))
					   PageContent=GetPicturePlayer(ContentArr,ModelID)
					   Templates = ""
					   Scan F_C
					   F_C = Templates
					   F_C = ReplaceRA(F_C, Trim(KS.C_C(TID,5))) '������ø����·��,���滻����·��Ϊ�����·��
					   Call FSOSaveFile(F_C, FilePathAndName)
				  End If	  
			 case Else   
			 	F_C = TFileContent
				If InStr(F_C, "{UnTag:GetRelativeList") <> 0 Then F_C = ReplaceLableFlag(Replace(F_C, "{UnTag:GetRelativeList", "{Tag:GetRelativeList"))
				Templates = ""
								'����ϵͳ�滻Ȩ�ޱ�ǩ
				If Fcls.ChannelID=8 And Instr(F_C,"[KS_Charge]")<>0 Then
				 Dim ChargeContent:ChargeContent=KS.CutFixContent(F_C, "[KS_Charge]", "[/KS_Charge]", 1)
				 F_C=Replace(F_C,ChargeContent,LFCls.GetConfigFromXML("supply","/labeltemplate/label","divajax"))
				End If

				Scan F_C
				F_C = Templates
				F_C = ReplaceRA(F_C, Trim(KS.C_C(TID,5))) '������ø����·��,���滻����·��Ϊ�����·��
				Call FSOSaveFile(F_C, FilePathAndName)
			end select
			
		End Function
		
		
		
		Function GetPicturePlayer(PicUrlsArr,ChannelID)
			 Dim I, TotalPictureNum,PictureIDArrayStr,ImageSrcArrayStr,ThumbSrcArrayStr
			 TotalPictureNum = UBound(PicUrlsArr) + 1
			 For I = 0 To TotalPictureNum - 1
			  PictureIDArrayStr = PictureIDArrayStr & "'" & Split(PicUrlsArr(I), "|")(0) & "',"
			  ImageSrcArrayStr = ImageSrcArrayStr & "'" & Split(PicUrlsArr(I), "|")(1) & "',"
			  ThumbSrcArrayStr=ThumbSrcArrayStr & "'" & Split(PicUrlsArr(I),"|")(2) &"',"
			 Next
			 PictureIDArrayStr = Left(PictureIDArrayStr, Len(PictureIDArrayStr) - 1)
			 ImageSrcArrayStr  = Left(ImageSrcArrayStr, Len(ImageSrcArrayStr) - 1)
			 ThumbSrcArrayStr  = left(ThumbSrcArrayStr,Len(ThumbSrcArrayStr)-1)
			 GetPicturePlayer  = LFCls.GetConfigFromXML("Label","/labeltemplate/label","imageplayer")
			 GetPicturePlayer  = Replace(GetPicturePlayer,"{$WebUrl}",DomainStr)
			 GetPicturePlayer  = Replace(GetPicturePlayer,"{$PictureIDArrayStr}",PictureIDArrayStr)
			 GetPicturePlayer  = Replace(GetPicturePlayer,"{$ImageSrcArrayStr}",ImageSrcArrayStr)
			 GetPicturePlayer  = Replace(GetPicturePlayer,"{$ThumbSrcArrayStr}",ThumbSrcArrayStr)
			 GetPicturePlayer  = Replace(GetPicturePlayer,"{$TotalPictureNum}",TotalPictureNum)
			 GetPicturePlayer  = Replace(GetPicturePlayer,"{$ShowUrl}",KS.Setting(3) & KS.C_S(ChannelID,10)&"/show.asp?id=" & FCls.RefreshInfoID)
		End Function
		
		
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'��������RefreshFolder
		'��  �ã�ˢ����Ŀҳ��
		'��  ����RS Recordset���ݼ�
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Function RefreshFolder(ChannelID,RS)
			Dim F_C, FolderDir, FilePath, Index
			Call FCls.SetClassInfo(RS("ChannelID"),RS("ID"),RS("TN"))
			F_C = LoadTemplate(RS("FolderTemplateID"))
			F_C = ReplaceAllLabel(F_C)
			F_C = ReplaceLableFlag(F_C)   '�滻������ǩ
            F_C = ReplaceGeneralLabelContent(F_C)          '�滻��վͨ�ñ�ǩ
			 If KS.C_S(ChannelID,44)="1" Or (KS.C_S(ChannelID,44)="3" And Trim(RS("TN")) = "0") Then 
			 Index = RS("FolderFsoIndex")
			 Else
             Index=KS.C_S(ChannelID,45) & "_" & rs("classid")&Mid(Trim(RS("FolderFsoIndex")), InStrRev(Trim(RS("FolderFsoIndex")), ".")) '�������չ��
			 End If
			 FolderDir = KS.C_S(ChannelID,8)
			 If Left(FolderDir, 1) = "/" Or Left(FolderDir, 1) = "\" Then FolderDir = Right(FolderDir, Len(FolderDir) - 1)
			
			 If KS.C_S(ChannelID,44)="1"  Or RS("ClassType")="3" Then 
			   FilePath = KS.Setting(3) & FolderDir & RS("Folder")
			 ElseIf KS.C_S(ChannelID,44)="2" Then
			   FilePath = KS.Setting(3) & FolderDir
			 Else
			   FilePath = KS.Setting(3) & FolderDir & Split(RS("Folder"),"/")(0) & "/"
			 End If

			 If RS("ClassType")="3" Then
			  Dim FsoName:FsoName = Mid(FilePath, InStrRev(FilePath, "/")) '�������չ��
			  Call KS.CreateListFolder(Replace(FilePath,FsoName,""))
			 Else
			  Call KS.CreateListFolder(FilePath)
			 End If
			 
			If (FCls.PageList <> "") Then
			  Call GetPageStr(FCls.PageList, "", Index, F_C, FilePath, Trim(RS("FolderDomain")))
			  FCls.PageList=""
			Else
			 F_C = Replace(F_C, "{PageListStr}", "")
			 F_C = ReplaceRA(F_C, Trim(RS("FolderDomain")))
			 If RS("ClassType")="3" Then Index=""
			 Call FSOSaveFile(F_C, FilePath & Index)
		   End If
		End Function
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'��������RefreshSpecials
		'��  �ã�ˢ��ר��ҳ��
		'��  ����RS Recordset���ݼ�
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Function RefreshSpecials(RS)
			Dim F_C, SpecialDir, FilePath,Index,TempStr
			'����ˢ������,��ȡ�õ�ǰ����λ��
			Call FCls.SetSpecialInfo(RS("ClassID"),RS("SpecialID"))                       
			'����ר��ҳ��Ӧ��ģ��
			  F_C = LoadTemplate(RS("TemplateID"))
  			  F_C = ReplaceSpecialContent(F_C,RS)
			  F_C = KSLabelReplaceAll(F_C)
			  Index = Trim(RS("FsoSpecialIndex"))
			  SpecialDir = KS.Setting(95)
			  If Left(SpecialDir, 1) = "/" Or Left(SpecialDir, 1) = "\" Then SpecialDir = Right(SpecialDir, Len(SpecialDir) - 1)
			  FilePath = KS.Setting(3) & SpecialDir & RS("SpecialEName") & "/"
			  Call KS.CreateListFolder(FilePath)
			  F_C = ReplaceLableFlag(F_C)                    '�滻������ǩ
			  If (FCls.PageList <> "") Then
				Call GetPageStr(FCls.PageList, Trim(DomainStr & SpecialDir & RS("SpecialEname") & "/"), Index, F_C, FilePath, "")
				FCls.PageList = ""
			  Else
				   F_C = Replace(F_C, "{PageListStr}", "")
				   Call FSOSaveFile(F_C, FilePath & Index)
			  End If
		End Function
		Function ReplaceSpecialContent(F_C,RS)
		 F_C=Replace(F_C,"{$GetSpecialName}",RS("SpecialName"))
		 If Not Isnull(RS("PhotoUrl")) And RS("PhotoUrl")<>"" Then
		 F_C=Replace(F_C,"{$GetSpecialPic}","<img src=""" & RS("PhotoUrl") & """ border=""0"">")
		 Else
		 F_C=Replace(F_C,"{$GetSpecialPic}","<img src=""" & DomainStr & "images/nophoto.gif"" border=""0"">")
		 End If
		 F_C=Replace(F_C,"{$GetSpecialNote}",RS("SpecialNote"))
		 F_C=Replace(F_C,"{$GetSpecialDate}",RS("SpecialAddDate"))
		 ReplaceSpecialContent=ReplaceSpecialClass(F_C)
		End Function
		Function ReplaceSpecialClass(F_C)
		 If FCls.RefreshType="Special" Or FCls.RefreshType="ChannelSpecial" Then
		   F_C=Replace(F_C,"{$GetSpecialClassName}",KS.GetSpecialClass(FCls.RefreshFolderID,"classname"))
		   F_C=Replace(F_C,"{$GetSpecialClassURL}",KS.GetFolderSpecialPath(FCls.RefreshFolderID, True))
		 End If
		 ReplaceSpecialClass=F_C
		End Function
		
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'��������RefreshSpecialClass
		'��  �ã�ˢ��Ƶ��ר�����ҳ
		'��  ����RS Recordset���ݼ�
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Function RefreshSpecialClass(RS)
			 Dim F_C, SpecialDir, Index, FilePath
			  FCls.RefreshType = "ChannelSpecial"    
			  FCls.RefreshFolderID = RS("ClassID")
			  FCls.ItemUnit="��"
			 If RS("TemplateID")="" Then
			 RefreshSpecialClass="���Ȱ�ר�����ģ��!":exit function
			 Else
			 F_C = LoadTemplate(RS("TemplateID"))
			 End If
			
			 F_C = ReplaceSpecialClass(F_C)  
			 F_C = KSLabelReplaceAll(F_C)
			 
			  SpecialDir = KS.Setting(95)
			  If Left(SpecialDir, 1) = "/" Or Left(SpecialDir, 1) = "\" Then SpecialDir = Right(SpecialDir, Len(SpecialDir) - 1)
			   
			  Index = RS("FsoIndex")
			  FilePath = KS.Setting(3) & SpecialDir & RS("ClassEname") & "/"
			  Call KS.CreateListFolder(FilePath)

			  If (FCls.PageList <> "") Then
				Call GetPageStr(FCls.PageList, Trim(DomainStr & SpecialDir & RS("ClassEname") & "/"), Index, F_C, FilePath, "")
				FCls.PageList=""
			  Else
				F_C = ReplaceRA(F_C, "")
				Call FSOSaveFile(F_C, FilePath & Index)
			 End If
		End Function
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		'��������RefreshCommonPage
		'��  �ã�ˢ��ͨ��ҳ��
		'��  ����RS Recordset���ݼ�
		'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		Function RefreshCommonPage(ByVal FileName,FsoFileName)
		  Dim F_C, CommonDir, FilePath
		      F_C = LoadTemplate(FileName)
			  F_C = KSLabelReplaceAll(F_C) 
			  F_C = Replace(Replace(F_C,"{$InfoID}","0"),"{$GetClassID}","0")
			  
			  '������ø����·��,���滻����·��Ϊ�����·��
			  F_C = ReplaceRA(F_C, "")
			  CommonDir = Replace(KS.Setting(94), "\", "")
			  If Left(CommonDir, 1) = "/" Then CommonDir = Right(CommonDir, Len(CommonDir) - 1)
			  'FilePath = KS.Setting(3) & CommonDir
			   FilePath=Replace(FsoFileName,Split(FsoFileName,"/")(Ubound(Split(FsoFileName,"/"))),"")
			  
			  Call KS.CreateListFolder(KS.Setting(3) & CommonDir & FilePath)
			  Call FSOSaveFile(F_C, KS.Setting(3) & CommonDir & FsoFileName)
		End Function
		
		'*********************************************************************************************************
		'��������ReplaceRA
		'��  �ã��Զ��ж�ϵͳ�Ƿ������·�������·����ת��
		'��  ����FileContentԭ�ļ�,FolderDomain �Ƿ��а󶨶�������
		'*********************************************************************************************************
		Function ReplaceRA(F_C, FolderDomain)
		     If Lcase(Fcls.RefreshType)="content" Then  F_C=ReplaceSQLLabel(F_C)
			 If CStr(KS.Setting(97)) = "0" Then
				 If FolderDomain <> "" Then
				   F_C = Replace(F_C, FolderDomain, "/")
				 Else
					  If Trim(KS.Setting(3)) = "/" Then
						F_C = Replace(F_C, DomainStr, "/")
					  Else
						F_C = Replace(F_C, Replace(DomainStr, Trim(KS.Setting(3)), ""), "")
					  End If
				End If
			  End If
			ReplaceRA = F_C
		End Function
		'*********************************************************************************************************
		'��������GetPageStr
		'��  �ã�ȡ�÷�ҳ��ͨ�ú���
		'��  ����PageContent--��ҳ����,LinkUrl--���ӵ�ַ,Index-��ҳ����
		'        F_C--��������ļ�����,FilePath---������·��,SecondDomain --��������
		'*********************************************************************************************************
		Sub GetPageStr(PageContent, LinkUrl, Index, F_C, FilePath, SecondDomain)
			Dim PageStr, FileStr, I, PageContentArr,LoopEnd,TotalPage,Fname,FExt ,LinkUrlFname
			  FExt = Mid(Trim(Index), InStrRev(Trim(Index), ".")) '�������չ��
			  Fname = Replace(Trim(Index), FExt, "")              '������ļ���
			  LinkUrlFname = LinkUrl & Fname
			  Dim HomeLink:HomeLink=LinkUrl & Index
			  
			  PageContentArr = Split(PageContent, "{KS:PageList}")
			  TotalPage = FCls.TotalPage
			  If KS.ChkClng(FCls.FsoListNum)<>0 and KS.ChkClng(FCls.FsoListNum)<FCls.TotalPage Then LoopEnd=KS.ChkClng(FCls.FsoListNum) Else LoopEnd=FCls.TotalPage
			  I=0
			  Do While I<LoopEnd
			   I=I+1  
			   '=========����Ϊ��ҳ��̬��======================
			   Select Case FCls.PageStyle
			    Case 1
				   If I=1 and I<>TotalPage Then
				   PageStr = "��ҳ  ��һҳ <a href=""" & LinkUrlFname & "_" & TotalPage -1 & FExt & """>��һҳ</a>  <a href= """ & LinkUrlFname & "_1" & FExt & """>βҳ</a>"
				   ElseIf I=1 And I=TotalPage Then
					PageStr ="��ҳ  ��һҳ ��һҳ βҳ"
				   ElseIf (I=TotalPage And I <> 2) Then
					 PageStr="<a href=""" &  HomeLink  & """>��ҳ</a>  <a href=""" &  LinkUrlFname  & "_"  &  TotalPage-I+2 & FExt  & """>��һҳ</a> ��һҳ  βҳ"
				   ElseIf(I = TotalPage And I = 2) Then
					 PageStr="<a href=""" & HomeLink & """>��ҳ</a>  <a href=""" &  HomeLink  & """>��һҳ</a> ��һҳ  βҳ"
				   ElseIf(I = 2) Then
					 PageStr="<a href=""" &  HomeLink  & """>��ҳ</a>  <a href=""" & HomeLink & """>��һҳ</a> <a href=""" & LinkUrlFname  & "_" & TotalPage-I & FExt  & """>��һҳ</a>  <a href= """ &  LinkUrlFname  & "_1" & FExt  & """>βҳ</a>"
				   Else
					 PageStr="<a href=""" & HomeLink  & """>��ҳ</a>  <a href=""" & LinkUrlFname & "_" & TotalPage-I+2 & FExt  & """>��һҳ</a> <a href=""" & LinkUrlFname  & "_" & TotalPage -I & FExt & """>��һҳ</a>  <a href= """ & LinkUrlFname & "_1" & FExt &""">βҳ</a>"
				   End If
			       PageStr="�� <span id=""totalrecord"">" & Fcls.TotalPut & "</span> " & FCls.ItemUnit &"  ҳ��:<span id=""currpage"" style=""color:red""> " & I & "</span>/<span id=""totalpage"">" & TotalPage & "</span>ҳ  <span id=""perpagenum"">" & FCls.PerPageNum & "</span>" & FCls.ItemUnit &"/ҳ " & PageStr
				Case 2,3
						PageStr="��<span id=""currpage"" style=""color:red"">" &  I  & "</span>ҳ ��<span id=""totalpage"">" & TotalPage & "</span>ҳ "
						If I=1 Then
						 PageStr=PageStr & "<span style=""font-family:webdings;font-size:14px"">9</span> <span style=""font-family:webdings;font-size:14px"">7</span> "
						ElseIf I=2 Then
						 PageStr=PageStr & "<a href=""" & HomeLink & """ title=""��ҳ""><span style=""font-family:webdings;font-size:14px"">9</span></a> <a href=""" &  HomeLink & """ title=""��һҳ""><span style=""font-family:webdings;font-size:14px"">7</span></a> "
						Else
						 PageStr=PageStr & "<a href=""" & HomeLink  & """ title=""��ҳ""><span style=""font-family:webdings;font-size:14px"">9</span></a> <a href=""" & LinkUrlFname & "_" &  TotalPage-I+2 & FExt & """ title=""��һҳ""><span style=""font-family:webdings;font-size:14px"">7</span></a> "
						End If
						
						If FCls.PageStyle=2 Then
						     PageStr=PageStr & " <span id=""pagelist"">"
							 Dim startpage:startpage=1
							 Dim P,n:n=1
							 If I>10 Then  startpage=(I/10-1)*10+(i mod 10)+1
							 for p=startpage to TotalPage
								If P=1 Then
									 If (P=I) Then
										PageStr=PageStr & "<a href=""" & HomeLink & """><font color=""#ff0000"">[" & P & "]</font></a>&nbsp;"
									 else
										PageStr=PageStr & "<a href=""" & HomeLink & """>[" & P & "]</a>&nbsp;"
									 End If
								Else
									  if (p=i) Then
										PageStr=PageStr & "<a href=""" & LinkUrlFname & "_" &  TotalPage-P+1 & FExt & """><font color=""#ff0000"">[" & i & "]</font></a>&nbsp;"
									  else
										PageStr=PageStr & "<a href=""" & LinkUrlFname & "_" & TotalPage-p+1 & FExt  &""">[" & p & "]</a>&nbsp;"
									  end if
								End If
								  n=n+1
								  if n>10 Then Exit For
							 Next
						  PageStr=PageStr & "</span>"
					  End If
					  
					If I=TotalPage Then
					  PageStr=PageStr & "<span style=""font-family:webdings;font-size:14px"">8</span> <span style=""font-family:webdings;font-size:14px"">:</span>"
					Else
					  PageStr=PageStr & "<a href=""" & LinkUrlFname &"_" & TotalPage -I & FExt &""" title=""��һҳ""><span style=""font-family:webdings;font-size:14px"">8</span></a> <a href=""" & LinkUrlFname & "_1" & FExt  & """><span style=""font-family:webdings;font-size:14px"">:</span></a>"
					End If
			 case 4  '������ʽ
				 
				 n=0:startpage=1
				 PageStr="<table border=""0"" align=""right""><tr><td id=""pagelist"">" & vbcrlf
				 if (I>1) then pageStr=PageStr & "<a href=""" & LinkUrlFname & "_" &  TotalPage-I+2 & FExt & """ class=""prev"">��һҳ</a>"
				 if (I<>TotalPage) then pageStr=PageStr & "<a href=""" & LinkUrlFname &"_" & TotalPage -I & FExt & """ class=""next"">��һҳ</a>"
				 pageStr=pageStr & "<a href=""" & HomeLink & """ class=""prev"">�� ҳ</a>"
				 if (I>=7) then startpage=I-5
				 if TotalPage-I<5 Then startpage=TotalPage-10
				 If startpage<0 Then startpage=1
				 For p=startpage To TotalPage
				    If p= I Then
				     PageStr=PageStr & " <a href=""#"" class=""curr""><font color=red>" & p &"</font></a>"
				    Else
				     PageStr=PageStr & " <a class=""num"" href=""" & LinkUrlFname & "_" & TotalPage-p+1 & FExt &""">" & p &"</a>"
					End If
					n=n+1
					if n>=10 then exit for
				 Next
				 If TotalPage=1 Then
				 pageStr=pageStr & "<a href=""" & LinkUrlFname & FExt &""" class=""prev"">ĩҳ</a>"
				 Else
				 pageStr=pageStr & "<a href=""" & LinkUrlFname & "_1" & FExt &""" class=""prev"">ĩҳ</a>"
				 End If
				 pageStr=PageStr & " <span>�ܹ�<span id=""totalpage"">" & TotalPage & "</span>ҳ</span></td></tr></table>"
			   End Select
			   
			   
			   If FCls.PageStyle<>4 Then
			   PageStr=PageStr &" ת����<select id=""turnpage"" size=""1"" onchange=""javascript:window.location=this.options[this.selectedIndex].value""><option value=""1"">1</option></select>"
			   End If
			   PageStr=PageStr & vbcrlf & "<script src=""page" & FCls.RefreshFolderID &".html"" type=""text/javascript"" language=""javascript""></script>"&vbcrlf &"<script language=""javascript"" type=""text/javascript"">pageinfo("&FCls.PageStyle&"," &FCls.PerPageNum &",'"&FExt&"','"&Fname&"');</script>"
			   
			  FileStr = Replace(F_C, "{PageListStr}",  PageContentArr(I-1)& "<div id=""fenye"" class=""plist"" style=""margin-top:6px;text-align:right;"">" & PageStr & "</div>")
			  
			  '===============��ҳ��̬������=====================================================
			   
			   
			   			  

			   FileStr = ReplaceRA(FileStr, SecondDomain)
			   if (TotalPage-I+1>0) Then
				   Dim TempFilePath
				   If I = 1  Then
					  TempFilePath = FilePath & Index
				   Else
					 TempFilePath = FilePath & Fname & "_" & TotalPage-I+1 & FExt
				   End If
				   Call FSOSaveFile(FileStr, TempFilePath)
               End If
 			  Loop
			  
			   If FCls.RefreshType="Folder" And LoopEnd>5 Then KS.Echo "<script>closeWindow();</script>"
			   
          	   Dim JSStr
			   JSStr="var TotalPage=" & TotalPage & ";"&vbcrlf & "var TotalPut=" & KS.ChkClng(Fcls.TotalPut) & ";" &vbcrlf
               JSStr=JSStr & "document.write(""<script language='javascript' src='" & KS.Setting(2) & "/ks_inc/kesion.page.js'></script>"");"&vbcrlf
			   Call FSOSaveFile(JSStr,FilePath&"page" & FCls.RefreshFolderID & ".html")
		End Sub
			
		
		'*********************************************************************************************************
		'��������ReplaceGeneralLabelContent
		'��  �ã��滻ͨ�ñ�ǩΪ����
		'��  ����FileContentԭ�ļ�
		'*********************************************************************************************************
		Function ReplaceGeneralLabelContent(F_C)
				Templates=""
				Scan F_C
				ReplaceGeneralLabelContent = Templates
		End Function
		
		Function GetTags(TagType,Num)
		  if not isnumeric(num) then exit function
		  dim sqlstr,sql,i,n,str
		  select case cint(tagtype)
		   case 1:sqlstr="select top 500 keytext,hits from ks_keywords where IsSearch=0 order by hits desc"
		   case 2:sqlstr="select top 500 keytext,hits from ks_keywords where IsSearch=0 order by lastusetime desc,id desc"
		   case 3:sqlstr="select top 500 keytext,hits from ks_keywords where IsSearch=0 order by Adddate desc,id desc"
		   case else 
		    GetTags="":exit function
		  end select
		  
		  dim rs:set rs=conn.execute(sqlstr)
		  if rs.eof then rs.close:set rs=nothing:exit function
		  sql=rs.getrows(-1)
		  rs.close:set rs=nothing
		  for i=0 to ubound(sql,2)
		   if KS.FoundInArr(str,sql(0,i),",")=false then
		    n=n+1
		    str=str & "," & sql(0,i)
		    gettags=gettags & "<a href=""" & DomainStr & "plus/tags.asp?n=" & server.URLEncode(sql(0,i))& """ target=""_blank"" title=""TAG:" & sql(0,i) & "&#10;��������" & SQL(1,I) &"��"">" & sql(0,i) & "</a> "
		   end if
		   if n>=cint(num) then exit for
		  next
		  
		End Function
		'*********************************************************************************************************
		'��������GetSiteCountAll
		'��  �ã��滻��վͳ�Ʊ�ǩΪ����
		'��  ����Flag-0��ͳ�ƣ�1-����ͳ�� 2-ͼƬͳ��
		'*********************************************************************************************************
		Function GetSiteCountAll()
		   Dim ChannelTotal: ChannelTotal = Conn.Execute("Select Count(*) From KS_Class Where TN='0'")(0)
		   Dim MemberTotal:MemberTotal=Conn.Execute("Select Count(*) From KS_User")(0)
		   Dim CommentTotal: CommentTotal = Conn.Execute("Select Count(*) From KS_Comment")(0)
		   Dim GuestBookTotal:GuestBookTotal=Conn.Execute("Select Count(ID) From KS_GuestBook")(0)
		   GetSiteCountAll="<div class=""sitetotal"">" & vbcrlf
			  GetSiteCountAll = GetSiteCountAll & "<li>Ƶ�������� " & ChannelTotal & " ��</li>" & vbcrlf
			  dim rsc:set rsc=conn.execute("select channelid,ItemName,Itemunit,channeltable from ks_channel where channelstatus=1 and channelid<>6 And ChannelID<>9 and channelid<>10  and channelid<>11")
			  dim k,sql:sql=rsc.getrows(-1)
			  rsc.close:set rsc=nothing
			  for k=0 to ubound(sql,2)
			  GetSiteCountAll = GetSiteCountAll & "<li>" & sql(1,k) & "������ " & Conn.Execute("Select Count(id) From " & sql(3,k))(0) & " " & sql(2,k)&"</li>" & vbcrlf
			  next
			  GetSiteCountAll = GetSiteCountAll & "<li>ע���Ա�� " & MemberTotal & " λ</li>" & vbcrlf
			  GetSiteCountAll = GetSiteCountAll & "<li>���������� " & GuestBookTotal &" ��</li>" & vbcrlf
			  GetSiteCountAll = GetSiteCountAll & "<li>���������� " & CommentTotal & " ��</li>" & vbcrlf
			  GetSiteCountAll = GetSiteCountAll & "<li>���������� <script language=""javascript"" src=""" & DomainStr & "plus/online.asp?ID=1""></script> ��</li>" & vbcrlf
		   GetSiteCountAll = GetSiteCountAll & "</div>" & vbcrlf
		End Function
		

		
		
		Function ReplaceKeyTags(KeyStr)
		  On error resume next
		  Dim I,K_Arr:K_Arr=Split(KeyStr,",")
		  For I=0 To Ubound(K_Arr)
		    ReplaceKeyTags=ReplaceKeyTags & "<a href=""" & DomainStr & "plus/tags.asp?n=" & K_Arr(i) & """ target=""_blank"">" & K_Arr(i) & "</a> "
		  Next
		  If Err Then ReplaceKeyTags="":Err.Clear
		End Function
		'�滻���л����
		Function ReplaceAD(ByVal Content,ClassID)
		 Dim ShowADTF,CLen,Dir,Width,Height,AdUrl,AdLinkUrl,LC,RC,AdStr,ADType
		 Dim ClassBasicInfo:ClassBasicInfo=KS.C_C(ClassID,6)
		 If ClassBasicInfo="" Then Exit Function
		 Dim AdP:AdP = Split(Split(ClassBasicInfo,"||||")(4),"%ks%")
		 ShowADTF=KS.ChkClng(Adp(0))
		 If ShowADTF=0 Then ReplaceAD=Content:Exit Function
		 Dim Param:Param=Split(AdP(1),",")

		 CLen=KS.ChkClng(Param(0)):Dir=Param(1):Width=KS.ChkClng(Param(2)):Height=KS.ChkClng(Param(3)):AdUrl=Adp(3):AdLinkUrl=Adp(4):ADType=KS.ChkClng(ADP(2))

		 If CLen<>0 Then LC=InterceptString(Content,Clen)
		 RC=Right(Content,Len(Content)-Len(LC))		 		 
               If ADType=2 Then
			     Adstr="<table border=""0"" width="""& Width & """ height=""" & height & """ align="""&Dir&"""><tr><td>" & AdUrl & "</td></tr></table>"
			   Else
                    If Lcase(Right(AdUrl,3))="swf" Then'�ж��Ƿ�SwfͼƬ
						AdStr="<table width=""0"" border=""0"" align="""&Dir&"""><tr><td><object classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" codebase=""http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0""  height=""" & height & """ width="""&width&""" ><param name=""movie"" value="""&AdUrl&"""><param name=""quality"" value=""high""><embed src="""&AdUrl&""" quality=""high"" pluginspage=""http://www.macromedia.com/go/getflashplayer"" type=""application/x-shockwave-flash"" height=""" & height & """  width="""&Width&"""></embed></object></td></tr></table>"
					Else
						If AdLinkUrl="" Then AdLinkUrl="http://www.kesion.com"
						AdStr="<table width=""0"" border=""0"" align="""&Dir&"""><tr><td><a href="""&AdLinkUrl&"""  target=""_blank""><img border=""0"" src="""&AdUrl&""" height=""" & height & """ width="""&Width&"""></a></td></tr></table>"
					End If
				End If	

		 ReplaceAD=LC & AdStr & RC
	   End Function
	   '��ȡ�ַ���
		Function InterceptString(ByVal txt,length)
			Dim x,y,ii,c,ischines,isascii,tempStr
			length=Cint(length)
			txt=trim(txt):x = len(txt):y = 0
			if x >= 1 then
				for ii = 1 to x
					c=asc(mid(txt,ii,1))
					if  c< 0 or c >255 then
						y = y + 2:ischines=1:isascii=0
					else
						y = y + 1:ischines=0:isascii=1
					end if
					if y >= length then
						if ischines=1 and StrCount(left(trim(txt),ii),"<a")=StrCount(left(trim(txt),ii),"</a>") then
							txt = left(txt,ii) '"�ַ����޳�
							exit for
						else
							if isascii=1 then x=x+1
						end if
					end if
				next
				InterceptString = txt
			else
				InterceptString = ""
			end if
		End Function
		
		'�ж��ַ������ֵĴ���
		Public Function StrCount(Str,SubStr)        
			Dim iStrCount,iStrStart,iTemp
			iStrCount = 0:iStrStart = 1:iTemp = 0:Str=LCase(Str):SubStr=LCase(SubStr)
			Do While iStrStart < Len(Str)
				iTemp = Instr(iStrStart,Str,SubStr,vbTextCompare)
				If iTemp <=0 Then
					iStrStart = Len(Str)
				Else
					iStrStart = iTemp + Len(SubStr)
					iStrCount = iStrCount + 1
				End If
			Loop
			StrCount = iStrCount
		End Function
		
		
		Sub ReplaceHits(F_C,ChannelID,Id)
			If InStr(F_C, "{$GetHits}") <> 0 Then           '�������
				 F_C = Replace(F_C, "{$GetHits}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?Action=Count&GetFlag=0&m=" & ChannelID & "&ID=" & ID & """></Script>")
				 F_C = Replace(F_C, "{$GetHitsByDay}", "<Script Language=""Javascript"" Src=""" &DomainStr & "item/GetHits.asp?GetFlag=1&m=" & ChannelID & "&ID=" & ID & """></Script>")
				 F_C = Replace(F_C, "{$GetHitsByWeek}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?GetFlag=2&m=" & ChannelID & "&ID=" & ID & """></Script>")
				 F_C = Replace(F_C, "{$GetHitsByMonth}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?GetFlag=3&m=" & ChannelID &"&ID=" & ID & """></Script>")
			ElseIf InStr(F_C, "{$GetHitsByDay}") <> 0 Then  '���������
				 F_C = Replace(F_C, "{$GetHits}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?GetFlag=0&m=" & ChannelID &"&ID=" & ID & """></Script>")
				 F_C = Replace(F_C, "{$GetHitsByDay}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?Action=Count&GetFlag=1&m=" & ChannelID & "&ID=" & ID & """></Script>")
				 F_C = Replace(F_C, "{$GetHitsByWeek}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?GetFlag=2&m=" & ChannelID & "&ID=" & ID & """></Script>")
				 F_C = Replace(F_C, "{$GetHitsByMonth}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?GetFlag=3&m=" & ChannelID & "&ID=" & ID & """></Script>")
			ElseIf InStr(F_C, "{$GetHitsByWeek}") <> 0 Then '���������
				 F_C = Replace(F_C, "{$GetHits}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?GetFlag=0&m=" & ChannelID & "&ID=" & ID & """></Script>")
				 F_C = Replace(F_C, "{$GetHitsByDay}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?GetFlag=1&m=" & ChannelID & "&ID=" & ID & """></Script>")
				 F_C = Replace(F_C, "{$GetHitsByWeek}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?Action=Count&GetFlag=2&m=" & ChannelID & "&ID=" & ID & """></Script>")
				 F_C = Replace(F_C, "{$GetHitsByMonth}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?GetFlag=3&m=" & ChannelID & "&ID=" & ID & """></Script>")
			ElseIf InStr(F_C, "{$GetHitsByMonth}") <> 0 Then '���������
				 F_C = Replace(F_C, "{$GetHits}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?GetFlag=0&m=" & ChannelID & "&ID=" & ID & """></Script>")
				 F_C = Replace(F_C, "{$GetHitsByDay}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?GetFlag=1&m=" & ChannelID & "&ID=" & ID & """></Script>")
				 F_C = Replace(F_C, "{$GetHitsByWeek}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?GetFlag=2&m=" & ChannelID & "&ID=" & ID & """></Script>")
				 F_C = Replace(F_C, "{$GetHitsByMonth}", "<Script Language=""Javascript"" Src=""" & DomainStr & "item/GetHits.asp?Action=Count&GetFlag=3&m=" & ChannelID & "&ID=" & ID & """></Script>")
			End If
		End Sub
		
		
	
		'**************************************************
		'��������Published
		'��  �ã�ȡ�÷���ʱ�估��Ȩ��Ϣ
		'��  ������
		'**************************************************
		Function Published()
		 On Error Resume Next
		  Published=vbcrlf &"<script src=""" & domainstr & "ks_inc/ajax.js"" type=""text/javascript""></script>" & vbcrlf
		  Dim PublishInfo:PublishInfo = KS.Setting(15)
		  If PublishInfo <> "0" Then
		   Published = Published & "<!-- published at " & Now() & " " & PublishInfo & " -->" & vbCrLf
		  End If
		End Function
		
		
		'=================================================
		'��������GetVote
		'��  �ã���ʾ��վ����
		'��  ������
		'=================================================
		Function GetVote(VoteID)
			dim sqlVote,rsVote,i
			Dim Domain:Domain = KS.GetDomain
			sqlVote="select * from KS_Vote where ID=" & VoteID & " Order By NewestTF Desc"
			Set rsVote= conn.execute(sqlvote)
			if rsVote.bof and rsVote.eof then 
				GetVote= "&nbsp;û���κε���"
			else
				GetVote=GetVote & "<div class=""vote"">" & vbcrlf 
				GetVote=GetVote & "<form name='VoteForm" & VoteID &"' method='post' action='" & Domain & "plus/Vote.asp' target='_blank'>" &vbcrlf
				GetVote=GetVote & "&nbsp;&nbsp;&nbsp;&nbsp;"& rsVote("Title") &"<br>"&vbcrlf
				if rsVote("VoteType")="Single" then
					for i=1 to 8
						if trim(rsVote("Select"& i) &"")="" then exit for
						GetVote=GetVote & "<input type='radio' name='VoteOption' value='"& i &"'>" & rsVote("Select" & i) &"<br>"&vbcrlf
					next
				else
					for i=1 to 8
						if trim(rsVote("Select"& i) &"")="" then exit for
						GetVote=GetVote &  "&nbsp;&nbsp;&nbsp;&nbsp;<input type='checkbox' name='VoteOption' value='"& i &"'>&nbsp;"& rsVote("Select" & i) &"<br>"&vbcrlf
					next
				end if
				GetVote=GetVote &  "<br><input name='VoteType' type='hidden'value='"& rsVote("VoteType") &"'>"&vbcrlf
				GetVote=GetVote &  "<input name='Action' type='hidden' value='Vote'>"&vbcrlf
				GetVote=GetVote &  "<input name='ID' type='hidden' value='"& rsVote("ID") &"'>"&vbcrlf
				GetVote=GetVote &  "<div style='text-align:center'>"&vbcrlf
				GetVote=GetVote &  "<input type='image' src='" & domain & "Images/Default/voteSubmit.gif' border='0'>&nbsp;"&vbcrlf
				GetVote=GetVote &  "<a href='" & Domain & "plus/Vote.asp?Action=Show&ID=" & VoteID &"' target='_blank'><img src='" & domain & "Images/Default/voteView.gif' border='0'></a>"&vbcrlf
				GetVote=GetVote &  "</div></form>"&vbcrlf
				GetVote=GetVote & "</div>"&vbcrlf
			end if
			rsVote.close:set rsVote=nothing
		End Function
		'��ʾ��Ա��¼����
		Sub GetTopUser(Num,MoreStr)
		 Dim Sql,XML,Node,UserFace,UserName
		 Dim RSObj:Set RSObj=Conn.execute("Select Top " & Num &" UserID,UserName,UserFace,LoginTimes,sex From KS_User where groupid<>1 Order BY LoginTimes Desc,UserID Desc")
		 If Not RSObj.Eof Then Set Xml=KS.RsToXml(RSObj,"row","")
		 RSObj.Close : Set RSObj = Nothing
		 If IsObject(Xml) Then
			For each Node In Xml.DocumentElement.SelectNodes("row")
			  userface=Node.SelectSingleNode("@userface").text  : UserName=Node.SelectSingleNode("@username").text
			  if userface="" then
			   if Node.SelectSingleNode("@sex").text="��" then  userface="images/face/0.gif" else userface="images/face/girl.gif"
			  End If
			  If Left(Lcase(userface),4)<>"http" Then userface=KS.GetDomain & userface
			  echoln "<li><a href=""" & KS.GetDomain & "space/?" & Server.URLEncode(UserName) & """ target=""_blank"" class=""b""><img src=""" & userface & """ border=""0"" alt=""�û�:" & UserName & "&#13;&#10;��¼:" & Node.SelectSingleNode("@logintimes").text & "��&#13;&#10;�Ա�:" & Node.SelectSingleNode("@sex").text & """/></a><br /><a class=""u"" href=""" & KS.GetDomain & "space/?" & Server.URLEncode(UserName) & """ target=""_blank"">" & UserName & "</a></li>"
			Next
			If MoreStr<>"" Then Echo "<div style=""text-align:center""><a href=""" & KS.GetDomain & "user/?userlist.asp"" target=""_blank"">" & MoreStr & "</a></div>"
			Xml=Empty : Set Node=Nothing
		 End If
		End Sub
		
		'��ʾ��Ա��̬
		Sub GetUserDynamic(Num)
		 Dim RS,XML,Node
		 Set RS=Conn.Execute("Select Top " & Num & " id,username,Note,adddate,ico From KS_UserLog Order By Id Desc")
		 If Not RS.Eof Then Set XML=KS.RsToXml(RS,"row","")
		  RS.Close:Set RS=Nothing
		 If IsObject(XML) Then
		  for each Node In XML.DocumentElement.SelectNodes("row")
		    echoln "<li><span>" & KS.GetTimeFormat(Node.SelectSingleNode("@adddate").text) & "</span><img align=""absmiddle"" src=""" & KS.GetDomain & "images/user/log/" & Node.SelectSingleNode("@ico").text & ".gif"" alt=""""/><a href=""" & KS.GetDomain & "space/?" & Node.SelectSingleNode("@username").text & """ target=""_blank"">" & Node.SelectSingleNode("@username").text & "</a> " & Replace(Replace(Replace(Replace(Node.SelectSingleNode("@note").text,"{$GetSiteUrl}",KS.GetDomain),vbcrlf,""),"<p>",""),"</p>","") & "</li>"
		  next
		 XML=Empty : Set Node=Nothing
		 End If
		End Sub
		
		
				
		Function FormatImglink(content,url,totalpage)
           dim re:Set re=new RegExp
           re.IgnoreCase =true
           re.Global=True
		   'ȥ��onclick,onload�Ƚű� 
            're.Pattern = "\s[on].+?=([\""|\'])(.*?)\1" 
            'Content = re.Replace(Content, "") 
			Dim LinkStr
		    If TotalPage=1 Then
			 LinkStr="href=""$2"" target=""_blank"""
			Else
			 LinkStr="href=""" & Url & """"
			End If
			
		   '��SRC�������ŵ�ͼƬ��ַ�������� 
            re.Pattern = "<img.*?\ssrc=([^\""\'\s][^\""\'\s>]*).*?>" 
            Content = re.Replace(Content, "<a " & LinkStr & "><img src=""$2"" alt=""��������һҳ"" onmousewheel=""return bbimg(this)"" onload=""javascript:resizepic(this)"" border=""0""/></a>") 
		   '����ƥ��ͼƬSRC��ַ 
		   re.Pattern = "<img.*?\ssrc=([\""\'])([^\""\']+?)\1.*?>" 
           Content = re.Replace(Content, "<a " & LinkStr & "><img src=""$2"" alt=""��������һҳ"" onmousewheel=""return bbimg(this)"" onload=""javascript:resizepic(this)"" border=""0""/></a>") 

		  set re = nothing
          FormatImglink = content
		end function 
End Class
%> 
