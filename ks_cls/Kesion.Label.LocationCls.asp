<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Class RefreshLocationCls
		Private KS  
		Private KMRFObj,DomainStr,WebNameStr        
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  DomainStr=KS.GetDomain
		  WebNameStr=KS.Setting(0)
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		End Sub
		'***********************************************************************************************************
		'ȡ��λ�õ���
		'***********************************************************************************************************
		Function GetLocation(ParamNode)
		    Dim Bold, StartTag, NavType, Nav, OpenType, TitleCss,ShowTitle
			Bold       = ParamNode.GetAttribute("bold")
			StartTag   = ParamNode.GetAttribute("starttag")
			NavType    = ParamNode.getAttribute("navtype")
			Nav        = ParamNode.getAttribute("nav")
			OpenType   = ParamNode.getAttribute("opentype")
			TitleCss   = ParamNode.getAttribute("titlecss")
			ShowTitle  = ParamNode.getAttribute("showtitle")
			Dim NaviStr
			If CBool(Bold) = True Then StartTag = "<strong>" & StartTag & "</strong>"
			NaviStr = GetLocationNav(NavType, Nav)
			TitleCss=KS.GetCss(TitleCss)
			Select Case UCase(FCls.RefreshType)
			   Case "MORESPACE","MORELOG","MOREGROUP","MOREXC" GetLocation = GetMoreSpaceLocation(StartTag, NaviStr, OpenType, TitleCss)
			   Case "SPECIALINDEX" GetLocation = GetSpecialIndexLocation(StartTag, NaviStr, OpenType, TitleCss)
			   Case "FOLDER" GetLocation = GetFolderLocation(StartTag, NaviStr, OpenType, TitleCss, FCls.RefreshFolderID)
			   Case "CONTENT" GetLocation = GetContentLocation(StartTag, NaviStr, OpenType, TitleCss,FCls.RefreshFolderID,ShowTitle)
			   Case "CHANNELSPECIAL" GetLocation = GetSpecialClassLocation(StartTag, NaviStr, OpenType, TitleCss, FCls.RefreshFolderID)
			   Case "SPECIAL"  GetLocation = GetSpecialLocation(StartTag, NaviStr, OpenType, TitleCss, FCls.RefreshFolderID)
					 
	    '--------------------------------------------��Ա���ĵ���-------------------------------------------		   
			   Case "USERREGSTEP1" GetLocation = GetUserRegLocation(1,StartTag, NaviStr, OpenType, TitleCss)
			   Case "USERREGSTEP2" GetLocation = GetUserRegLocation(2,StartTag, NaviStr, OpenType, TitleCss)
			   Case "USERREGSTEP3" GetLocation = GetUserRegLocation(3,StartTag, NaviStr, OpenType, TitleCss)
			   Case "USERLIST"  GetLocation = GetUserListLocation(StartTag, NaviStr, OpenType, TitleCss)	
			   Case "SHOWUSER"  GetLocation = GetUserInfoLocation(StartTag, NaviStr, OpenType, TitleCss)	
			   Case "MEMBER"  GetLocation = GetMemberLocation(StartTag, NaviStr, OpenType, TitleCss)	
		'-------------------------------------------��Ա���ĵ�������----------------------------------------
		
	    '--------------------------------------------����Ƶ������-------------------------------------------		   
			   Case "MUSICINDEX"  GetLocation = GetMusicIndexLocation(StartTag, NaviStr, OpenType, TitleCss)
			   Case "MUSICSINGER" GetLocation = GetMusicSingerLocation(StartTag, NaviStr, OpenType, TitleCss)
			   Case "MUSICSINGERSPECIAL" GetLocation = GetMusicSingerSpecialLocation(StartTag, NaviStr, OpenType, TitleCss)
			   Case "MUSICSPECIAL" GetLocation = GetMusicSpecialLocation(StartTag, NaviStr, OpenType, TitleCss)
		'-------------------------------------------����Ƶ����������----------------------------------------
		
		'-------------------------------------------��������------------------------------------------------
		      Case "SHOPPINGCART" GetLocation = GetShoppingLocation(StartTag, NaviStr, OpenType, TitleCss,1)
			  Case "SHOPPINGPAYMENT" GetLocation = GetShoppingLocation(StartTag, NaviStr, OpenType, TitleCss,2)
			  Case "SHOPPINGPREVIEW" GetLocation = GetShoppingLocation(StartTag, NaviStr, OpenType, TitleCss,3)
			  Case "SHOPPINGSUCCESS" GetLocation = GetShoppingLocation(StartTag, NaviStr, OpenType, TitleCss,4)
			  Case Else GetLocation = GetIndexLocation(StartTag, NaviStr, OpenType, TitleCss)
		   End Select
		 
		End Function
		
		'ȡ����վ��ҳ����λ�õĺ���
		Function GetIndexLocation(StartTag, NaviStr, OpenType, TitleCss)
		   Dim str,Node
		   Select Case UCase(FCls.RefreshType)
		     case "INDEX" str="��վ��ҳ"
			 case "COMMENT"str="��������"
			 case "SEARCH" str="�������"
			 case "SPACEINDEX" str="�ռ���ҳ"
			 case "LINKINDEX" str="��������"
			 case "MAP" str="��վ��ͼ"
			 case "RSS" str="RSS���ķ���"
			 case "GUESTINDEX" 
			  str="<a href='index.asp'>" & KS.Setting(61) & "</a>"
			  If Request("pid")<>"" Then
			   KS.LoadClubBoard
			   Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & Request("pid") &"]")
			   If Not Node Is Nothing Then
			   Str=Str & Navistr & Node.SelectSingleNode("@boardname").text & Navistr & "�����б�"
			   End If
			  ElseIf Request("boardid")<>"" Then
			   KS.LoadClubBoard
			   Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & request("boardid") &"]")
			   If Not Node Is Nothing Then
			   Str=Str & Navistr & "<a href=""?pid=" & Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & Node.SelectSingleNode("@parentid").text & "]/@id").text & """>" & Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & Node.SelectSingleNode("@parentid").text & "]/@boardname").text & "</a>"
			   Str=Str & Navistr & Node.SelectSingleNode("@boardname").text & Navistr & KS.Setting(62) & "�б�"
			   End If
			  End If
			 case "GUESTWRITE" 
			  str="<a href='index.asp'>" & KS.Setting(61) & "</a>"
			  If KS.ChkClng(Request("bid"))<>0 Then
			    KS.LoadClubBoard
			    Set Node=Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & request("bid") &"]")
			   Str=Str & Navistr & "<a href=""?pid=" & Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & Node.SelectSingleNode("@parentid").text & "]/@id").text & """>" & Application(KS.SiteSN&"_ClubBoard").DocumentElement.SelectSingleNode("row[@id=" & Node.SelectSingleNode("@parentid").text & "]/@boardname").text & "</a>"
			   Str=Str & Navistr & "<a href=""index.asp?boardid=" & Node.SelectSingleNode("@id").text & """>" & Node.SelectSingleNode("@boardname").text &"</a>"
			  End If
			  str=str &  Navistr &"����" & KS.Setting(62)
			  
			 case "GUESTDISPLAY" 
			  str="<a href='index.asp'>" & KS.Setting(61) & "</a>"
			  if FCls.RefreshFolderID<>0 then str=str & Navistr & "<a href='index.asp?boardid=" &FCls.RefreshFolderID&"'>" & LFCls.GetSingleFieldValue("select boardname from ks_guestboard where id=" & FCls.RefreshFolderID) &"</a>"
			  str=str & Navistr & "�鿴" & KS.Setting(62) & ""

			 case "JOBINDEX" str="��ְ��Ƹ"
			 case "RESUMESEARCH" str= "<a " & TitleCss & " href=""index.asp""" & KS.G_O_T_S(OpenType) & ">��ְ��Ƹ</a>" & NaviStr & "�����˲�" 
			 case "SEARCHZW" str= "<a " & TitleCss & " href=""index.asp""" & KS.G_O_T_S(OpenType) & ">��ְ��Ƹ</a>" & NaviStr & "����ְλ" 
			 case "RESUMESC" str= "<a " & TitleCss & " href=""index.asp""" & KS.G_O_T_S(OpenType) & ">��ְ��Ƹ</a>" & NaviStr & "�����ղؼ�" 
			 case "COMPANYSHOW" str= "<a " & TitleCss & " href=""index.asp""" & KS.G_O_T_S(OpenType) & ">��ְ��Ƹ</a>" & NaviStr & "�鿴��˾����" 
			 case "JOBREAD" str= "<a " & TitleCss & " href=""index.asp""" & KS.G_O_T_S(OpenType) & ">��ְ��Ƹ</a>" & NaviStr & "�鿴ְλ����" 
			 case "JOBAPPLY" str= "<a " & TitleCss & " href=""index.asp""" & KS.G_O_T_S(OpenType) & ">��ְ��Ƹ</a>" & NaviStr & "����ְλ" 
			 case "LETTER" str= "<a " & TitleCss & " href=""index.asp""" & KS.G_O_T_S(OpenType) & ">��ְ��Ƹ</a>" & NaviStr & "��ְ��" 
			 case "ZHAOPIN" str= "<a " & TitleCss & " href=""index.asp""" & KS.G_O_T_S(OpenType) & ">��ְ��Ƹ</a>" & NaviStr & "��ҵ��Ƹ" 
			 case "QIUZHI" str= "<a " & TitleCss & " href=""index.asp""" & KS.G_O_T_S(OpenType) & ">��ְ��Ƹ</a>" & NaviStr & "������ְ" 
			 case "JOBLTINDEX" str= "<a " & TitleCss & " href=""index.asp""" & KS.G_O_T_S(OpenType) & ">��ְ��Ƹ</a>" & NaviStr & "��ͷ������ҳ" 
			 case "JOBLTINTRO" str= "<a " & TitleCss & " href=""index.asp""" & KS.G_O_T_S(OpenType) & ">��ְ��Ƹ</a>" & NaviStr & "��ͷ����" 
			 case "JOBLTNEWS" str= "<a " & TitleCss & " href=""index.asp""" & KS.G_O_T_S(OpenType) & ">��ְ��Ƹ</a>" & NaviStr & "������ͷְλ"
			 case "JOBJZJOB" str= "<a " & TitleCss & " href=""index.asp""" & KS.G_O_T_S(OpenType) & ">��ְ��Ƹ</a>" & NaviStr & "���¼�ְְλ"
			 case "JOBJZRESUME" str= "<a " & TitleCss & " href=""index.asp""" & KS.G_O_T_S(OpenType) & ">��ְ��Ƹ</a>" & NaviStr & "���¼�ְ�˲�"
			 case "JOBJZINDEX" str= "<a " & TitleCss & " href=""index.asp""" & KS.G_O_T_S(OpenType) & ">��ְ��Ƹ</a>" & NaviStr & "��ְ�����ҳ" 
			 case "RESUMESEARCH" str= "<a " & TitleCss & " href=""index.asp""" & KS.G_O_T_S(OpenType) & ">��ְ��Ƹ</a>" & NaviStr & "���������б�" 
			 case "JOBSEARCH" str= "<a " & TitleCss & " href=""index.asp""" & KS.G_O_T_S(OpenType) & ">��ְ��Ƹ</a>" & NaviStr & "ְλ�����б�" 
			 case "ENTERPRISE" str="��ҵ��ȫ"
			 case "ENTERPRISELIST" str="<a " & TitleCss & " href=""index.asp""" & KS.G_O_T_S(OpenType) & ">��ҵ��ȫ</a>" & NaviStr & FCls.LocationStr
			 case "ENTERPRISEPRO" str="��Ʒ��"
			 case "ENTERPRISEPROLIST" str="<a " & TitleCss & " href=""index.asp""" & KS.G_O_T_S(OpenType) & ">��Ʒ��</a>" & NaviStr & FCls.LocationStr 
			 case "ENTERPRISEZS" str="װ����ҵ��ȫ"
			 case "ENTERPRISELISTZS" str="<a " & TitleCss & " href=""index.asp""" & KS.G_O_T_S(OpenType) & ">װ����ҵ��ȫ</a>" & NaviStr & FCls.LocationStr
			 case else str=""
		   End Select
			  GetIndexLocation = StartTag & "<a " & TitleCss & " href=""" & DomainStr & """" & KS.G_O_T_S(OpenType) & ">" & WebNameStr & "</a>" & NaviStr & str
		End Function

		'ȡ�ø���ռ䵼��λ�õĺ���
		Function GetMoreSpaceLocation(StartTag, NaviStr, OpenType, TitleCss)
		   Dim MoreStr
		   Select Case UCase(FCls.RefreshType)
		    Case "MORESPACE":MoreStr="���˿ռ��б�"
			Case "MORELOG":MoreStr="��־�б�"
			Case "MOREGROUP":MoreStr="Ȧ���б�"
			Case "MOREXC":MoreStr="����б�"
		   End Select 
			  GetMoreSpaceLocation = StartTag & "<a " & TitleCss & " href=""" & DomainStr & """" & KS.G_O_T_S(OpenType) & ">" & WebNameStr & "</a>" & NaviStr & "<a " & TitleCss & " href=""index.asp""" & KS.G_O_T_S(OpenType) & ">�ռ���ҳ</a>" & NaviStr  &MoreStr
		End Function

		'���л�Ա�б�ҳ
		Function GetUserListLocation(StartTag, NaviStr, OpenType, TitleCss)
			  GetUserListLocation = StartTag & "<a " & TitleCss & " href=""" & DomainStr & """" & KS.G_O_T_S(OpenType) & ">" & WebNameStr & "</a>" & NaviStr & "����ע���Ա�б�"
		End Function
		'���л�Ա��Ϣҳ
		Function GetUserInfoLocation(StartTag, NaviStr, OpenType, TitleCss)
			  GetUserInfoLocation = StartTag & "<a " & TitleCss & " href=""" & DomainStr & """" & KS.G_O_T_S(OpenType) & ">" & WebNameStr & "</a>" & NaviStr & "<a href=""" & DomainStr & "User/UserList.asp"" " & KS.G_O_T_S(OpenType) & ">���л�Ա�б�</a>"& NaviStr & "��Ա��Ϣ"
		End Function
		'��Ա����
		Function GetMemberLocation(StartTag, NaviStr, OpenType, TitleCss)
			  GetMemberLocation = StartTag & "<a " & TitleCss & " href=""" & DomainStr & """" & KS.G_O_T_S(OpenType) & ">" & WebNameStr & "</a>" & NaviStr & "��Ա����"
		End Function
		'ȡ�û�Աע�ᵼ��
		Function GetUserRegLocation(Step,StartTag, NaviStr, OpenType, TitleCss)
		  Select Case Step
		    Case 1 GetUserRegLocation = StartTag & "<a " & TitleCss & " href=""" & DomainStr & """" & KS.G_O_T_S(OpenType) & ">" & WebNameStr & "</a>" & NaviStr & "�������������"
			Case 2 GetUserRegLocation = StartTag & "<a " & TitleCss & " href=""" & DomainStr & """" & KS.G_O_T_S(OpenType) & ">" & WebNameStr & "</a>" & NaviStr & "��дע���"
			Case 3 GetUserRegLocation = StartTag & "<a " & TitleCss & " href=""" & DomainStr & """" & KS.G_O_T_S(OpenType) & ">" & WebNameStr & "</a>" & NaviStr & "���ע��"
		  End Select
		End Function
		'ȡ��ר����ҳ����λ�õĺ���
		Function GetSpecialIndexLocation(StartTag, NaviStr, OpenType, TitleCss)
			  GetSpecialIndexLocation = StartTag & "<a " & TitleCss & " href=""" & DomainStr & """" & KS.G_O_T_S(OpenType) & ">" & WebNameStr & "</a>" & NaviStr & "ר����ҳ"
		End Function
		'ȡ��ר����ർ��
		Function GetSpecialClassLocation(StartTag, NaviStr, OpenType, TitleCss, RefreshFolderIDValue)
			 Dim SpecialIndexUrl,SpecialDir:SpecialDir = KS.Setting(95)
			 If Split(KS.Setting(5),".")(1)<>"asp" Then SpecialIndexUrl=DomainStr & SpecialDir Else SpecialIndexUrl=DomainStr & "SpecialIndex.asp"
			 If Left(SpecialDir, 1) = "/" Or Left(SpecialDir, 1) = "\" Then SpecialDir = Right(SpecialDir, Len(SpecialDir) - 1)
			 GetSpecialClassLocation = StartTag & "<a " & TitleCss & " href=""" & DomainStr & """" & KS.G_O_T_S(OpenType) & ">" & WebNameStr & "</a>" & NaviStr & "<a href=""" & SpecialIndexUrl & """" & KS.G_O_T_S(OpenType) & ">ר����ҳ</a>" & NaviStr & KS.C_C(RefreshFolderIDValue,1)  & KS.GetSpecialClass(RefreshFolderIDValue,"classname")
		
		End Function
		'ȡ��ר��ҳ��λ�õ���
		Function GetSpecialLocation(StartTag, NaviStr, OpenType, TitleCss, RefreshFolderIDValue)
			 Dim SpecialIndexUrl,SpecialDir:SpecialDir = KS.Setting(95)
			 If Split(KS.Setting(5),".")(1)<>"asp" Then SpecialIndexUrl=DomainStr & SpecialDir Else SpecialIndexUrl=DomainStr & "SpecialIndex.asp"
			 If Left(SpecialDir, 1) = "/" Or Left(SpecialDir, 1) = "\" Then SpecialDir = Right(SpecialDir, Len(SpecialDir) - 1)
			 Dim TempFolderStr
				  TempFolderStr = "<a " & TitleCss & " href=""" & KS.GetFolderSpecialPath(RefreshFolderIDValue, True) & """" & KS.G_O_T_S(OpenType) & ">" & KS.GetSpecialClass(RefreshFolderIDValue,"classname") & "</a>" & NaviStr
			 GetSpecialLocation = StartTag & "<a " & TitleCss & " href=""" & DomainStr & """" & KS.G_O_T_S(OpenType) & ">" & WebNameStr & "</a>" & NaviStr & "<a href=""" & SpecialIndexUrl & """" & KS.G_O_T_S(OpenType) & ">ר����ҳ</a>" & NaviStr & TempFolderStr & "���ר��"
		End Function
		'ȡ����Ŀ��λ�õ���
		Function GetFolderLocation(StartTag, NaviStr, OpenType, TitleCss, RefreshFolderIDValue)
			Dim FolderNaviStr:FolderNaviStr = GetFolderNaviStr(NaviStr, OpenType, TitleCss, RefreshFolderIDValue)
			If FCls.BrandName<>"" Then
				  GetFolderLocation = StartTag & "<a " & TitleCss & " href=""" & DomainStr & """" & KS.G_O_T_S(OpenType) & ">" & WebNameStr & "</a>" & FolderNaviStr & NaviStr & FCls.BrandName
			Else
				If FCls.RefreshChannelHomeFlag = True Then
				  GetFolderLocation = StartTag & "<a " & TitleCss & " href=""" & DomainStr & """" & KS.G_O_T_S(OpenType) & ">" & WebNameStr & "</a>" & FolderNaviStr & NaviStr & "��ҳ"
				Else
				  GetFolderLocation = StartTag & "<a " & TitleCss & " href=""" & DomainStr & """" & KS.G_O_T_S(OpenType) & ">" & WebNameStr & "</a>" & FolderNaviStr
				End If
		   End If
		End Function
		'ȡ����Ϣҳ����λ�õĺ���
		Function GetContentLocation(StartTag, NaviStr, OpenType, TitleCss, RefreshFolderIDValue,ShowTitle)
		    Dim Str,FolderNaviStr:FolderNaviStr = GetFolderNaviStr(NaviStr, OpenType, TitleCss, RefreshFolderIDValue)
			Str = StartTag & "<a " & TitleCss & " href=""" & DomainStr & """" & KS.G_O_T_S(OpenType) & ">" & WebNameStr & "</a>" & FolderNaviStr & NaviStr
			If Cbool(ShowTitle)=true Then Str=Str & Fcls.ItemTitle Else Str=Str & "���"& KS.C_S(FCls.Channelid,3)
			GetContentLocation = Str
		End Function
		
		'ȡ������Ƶ����ҳ����λ�õĺ���
		Function GetMusicIndexLocation(StartTag, NaviStr, OpenType, TitleCss)
			  GetMusicIndexLocation = StartTag & "<a " & TitleCss & " href=""" & DomainStr & """" & KS.G_O_T_S(OpenType) & ">" & WebNameStr & "</a>" & NaviStr & "<a " & TitleCss & " href=""" & DomainStr & "Music/" & KS.G_O_T_S(OpenType) & """>����Ƶ��</a>" & NaviStr & "��ҳ"
		End Function

		'ȡ�����ָ����б�ҳ����λ�õĺ���
		Function GetMusicSingerLocation(StartTag, NaviStr, OpenType, TitleCss)
			  GetMusicSingerLocation = StartTag & "<a " & TitleCss & " href=""" & DomainStr & """" & KS.G_O_T_S(OpenType) & ">" & WebNameStr & "</a>" & NaviStr & "<a " & TitleCss & " href=""" & DomainStr & "Music/" & KS.G_O_T_S(OpenType) & """>����Ƶ��</a>" & NaviStr & Application(KS.SiteSN & "RefreshMusicTempStr")  &NaviStr & "���ֵ����б�"
		End Function
		
		'ȡ�����ָ���ר���б�ҳ����λ�õĺ���
		Function GetMusicSingerSpecialLocation(StartTag, NaviStr, OpenType, TitleCss)
			  GetMusicSingerSpecialLocation = StartTag & "<a " & TitleCss & " href=""" & DomainStr & """" & KS.G_O_T_S(OpenType) & ">" & WebNameStr & "</a>" & NaviStr & "<a " & TitleCss & " href=""" & DomainStr & "Music/" & KS.G_O_T_S(OpenType) & """>����Ƶ��</a>" & NaviStr & KS.C("RefreshMusicClass") &NaviStr & Application(KS.SiteSN & "RefreshMusicSingerTempStr") & NaviStr & "����ר��"
		End Function
		'ȡ����������ר�������б�ҳ����λ�õĺ���
		Function GetMusicSpecialLocation(StartTag, NaviStr, OpenType, TitleCss)
			  GetMusicSpecialLocation = StartTag & "<a " & TitleCss & " href=""" & DomainStr & """" & KS.G_O_T_S(OpenType) & ">" & WebNameStr & "</a>" & NaviStr & "<a " & TitleCss & " href=""" & DomainStr & "Music/" & KS.G_O_T_S(OpenType) & """>����Ƶ��</a>" & NaviStr & Application(KS.SiteSN & "RefreshMusicTempStr")  &NaviStr & Application(KS.SiteSN & "RefreshMusicSingerTempStr") & NaviStr & Application(Cstr(KS.SiteSN & "RefreshMusicSpecialNameTempStr")) & "ר�������б�"
		End Function
        '��������
		Function GetShoppingLocation(StartTag, NaviStr, OpenType, TitleCss,TypeID)
		   GetShoppingLocation = StartTag & "<a " & TitleCss & " href=""" & DomainStr & """" & KS.G_O_T_S(OpenType) & ">" & WebNameStr & "</a>" & NaviStr & "�̳�����" & NaviStr
		   Select Case TypeID
		    Case 1: GetShoppingLocation=GetShoppingLocation & "�ҵĹ��ﳵ"
			Case 2: GetShoppingLocation=GetShoppingLocation & "����̨"
			Case 3: GetShoppingLocation=GetShoppingLocation & "Ԥ��������ȷ��"
			Case 4: GetShoppingLocation=GetShoppingLocation & "�����ύ�ɹ�"
		   End Select
		End Function
         
		'******************************************************************************************************
		'��������GetFolderNameStr
		'��  �ã�������Ŀ˳���б�
		'��  ����NaviStr--�����ַ���,RefreshFolderIDValue--��ĿID, OpenType---�´��ڴ�, TitleCss---������ʽ
		'����ֵ������: ��Ѵ���� >> ��Ʒ�б� >> ��Ѵ��վ����ϵͳ
		'******************************************************************************************************
		Function GetFolderNaviStr(NaviStr, OpenType, TitleCss, RefreshFolderIDValue)
			  Dim TSArr, I
			  TSArr = Split(KS.C_C(RefreshFolderIDValue,8), ",")
			  For I = LBound(TSArr) To UBound(TSArr) - 1
					GetFolderNaviStr = GetFolderNaviStr & NaviStr & "<a " & TitleCss & " href=""" & KS.GetFolderPath(TSArr(I)) & """" & KS.G_O_T_S(OpenType) & ">" & KS.C_C(TSArr(I),1) & "</a>"
			  Next
		End Function

		
		Function GetLocationNav(NavType, Nav)
			If CStr(NavType) = "0" Then
			  If Nav = "" Then
			   GetLocationNav = " >> "
			  Else
			   GetLocationNav = Nav
			  End If
			Else
			  If Nav = "" Then
				GetLocationNav = " >> "
			  Else
				If Left(Nav, 1) = "/" Or Left(Nav, 1) = "\" Then Nav = Right(Nav, Len(Nav) - 1)
				GetLocationNav = "<img src=""" & DomainStr & Nav & """ border=""0"" align=""absmiddle"">"
			  End If
			End If
		End Function

End Class
%> 
