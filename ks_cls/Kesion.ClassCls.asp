<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
'-----------------------------------------------------------------------------------------------
'��Ѵ��վ����ϵͳ,Ƶ����Ŀͨ����
'����:���ݿ�����Ϣ�������޹�˾ �汾 V 6.5
'-----------------------------------------------------------------------------------------------
Class ClassCls
        Private KS,KSCls
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub
		'���Ƶ����Ŀ¼�Ĺ���
			'���� channelID--Ƶ��ID,FolderID ��Ŀ¼,FormProcesPage--�������ҳ��
			Sub GetAddChannelFolder(Action,FolderID, FormProcesPage)
			 Dim WapSwitch,WapFolderTemplateID,WapTemplateID
			 Dim Folder,CurrPath,TemplateRS, TemplateSql, TypeList, NowDate, YearStr, MonthStr, DayStr,DefaultArrGroupID,ReadPoint,ChargeType,PitchTime,ReadTimes,AllowArrGroupID,DividePercent,K
			 Dim ClassBasicInfoArr,FolderName,FolderEname, ClassPic,ClassDescript,MetaKeyWord,MetaDescript,CommentTF,TopFlag,FolderTemplateID,FsoType,FolderFsoIndex,FolderDomain,TemplateID,FnameType,ClassPurview,ClassDefineContentArr,ClassContent
			 Dim TopTitle,SelStr,ClassType,ChannelID,ModelXML,Node
			 dim ShowADTF,AdParam,AdUrl,AdLinkUrl,AdP,AdType
			  CurrPath = KS.GetUpFilesDir():If Right(CurrPath,1)="/" Then CurrPath=Left(CurrPath, Len(CurrPath) - 1)
			  NowDate = Now():YearStr = Year(NowDate):MonthStr = Right("0"&Month(NowDate),2):DayStr = Right("0" & Day(NowDate),2)
			  If Action="Edit" Then
			    Dim RSE:Set RSE=Server.CreateObject("ADODB.RECORDSET")
				RSE.Open "Select * From KS_Class Where ID='" & FolderID & "'",Conn,1,1
				If Not RSE.Eof Then
				   FolderID=Rse("TN")
				   ChannelID=RSE("ChannelID")
				  FolderName       = Rse("FolderName")
				  ClassType        = Rse("ClassType")
				  If ClassType=2 Then
				  FolderEname      = Rse("Folder")
				  Else
				  FolderEname      = Split(Rse("Folder"), "/")(Rse("tj") - 1)
				  End If
				  CommentTF        = Rse("CommentTF")
				  TopFlag          = Rse("TopFlag") 
				  WapSwitch        = Rse("WapSwitch")
				  WapFolderTemplateID = Rse("WapFolderTemplateID")
				  WapTemplateID       = Rse("WapTemplateID")
				  FolderTemplateID = Rse("FolderTemplateID") 
				  TemplateID       = Rse("TemplateID")
				  FolderFsoIndex   = Rse("FolderFsoIndex")
				  FnameType        = Rse("FnameType")
				  FsoType          = Rse("FsoType")
				  FolderDomain     = Rse("FolderDomain")
				  ClassPurview     = Rse("ClassPurview")
				  DefaultArrGroupID= Rse("DefaultArrGroupID")
				  AllowArrGroupID  = Rse("AllowArrGroupID")
				  ReadPoint        = Rse("DefaultReadPoint")
				  PitchTime        = Rse("DefaultPitchTime")
				  ReadTimes        = Rse("DefaultReadTimes")
				  ChargeType       = Rse("DefaultChargeType")
				  DividePercent    = Rse("DefaultDividePercent")
				  
				  ClassBasicInfoArr=Split(Rse("ClassBasicInfo"),"||||")
				  ClassPic=ClassBasicInfoArr(0)
				  ClassDescript=ClassBasicInfoArr(1)
				  MetaKeyWord=ClassBasicInfoArr(2)
				  MetaDescript=ClassBasicInfoArr(3)
				  '���л�
				  AdP=Split(ClassBasicInfoArr(4),"%ks%")
				  ShowADTF=Adp(0)
				  AdParam=Adp(1)
				  AdType=Adp(2)
				  AdUrl=Adp(3)
				  AdLinkUrl=Adp(4)
				  ClassDefineContentArr=Split(Rse("ClassDefineContent"),"||||")
				  ClassContent=ClassDescript
				  TopTitle="�༭��Ŀ"
				End If
			  Else
			    TopTitle="��������Ŀ"
				ClassType=1
			    CommentTF=3:TopFlag=1:WapSwitch=1:FsoType=11:FolderFsoIndex="index.html":FnameType=".html"
				ClassPurview=0:DefaultArrGroupID=0:AllowArrGroupID=0:ReadPoint=0:DividePercent=0:PitchTime=12:ReadTimes=10:ShowADTF=0:AdParam="250,left,300,300":AdUrl="":AdLinkUrl="http://www.kesion.com":AdType=1
				ChannelID=KS.ChkClng(Request("ChannelID"))
				If ChannelID=0 Then ChannelID=1
				If FolderID="0" Or FolderID="" or FolderID="1"  Then
				FolderTemplateID="{@TemplateDir}/" & KS.C_S(ChannelID,1) & "/Ƶ����ҳ.html"
				WapFolderTemplateID="{@TemplateDir}/WAPר��ģ��/" & KS.C_S(ChannelID,1) & "/WAPƵ����ҳ.html"
				Else
				FolderTemplateID="{@TemplateDir}/" & KS.C_S(ChannelID,1) & "/��Ŀҳ.html"
				WapFolderTemplateID="{@TemplateDir}/WAPר��ģ��/" & KS.C_S(ChannelID,1) & "/WAP��Ŀҳ.html"
				End If
				TemplateID="{@TemplateDir}/" & KS.C_S(ChannelID,1) & "/����ҳ.html"
				WapTemplateID="{@TemplateDir}/WAPר��ģ��/" & KS.C_S(ChannelID,1) & "/WAP����ҳ.html"
			  End If
			  TypeList = Replace(KS.LoadClassOption(0),"value='" & FolderID & "'","value='" & FolderID &"' selected")
			  
			With Response
				.Write "<html>" & vbCrLf
				.Write "<head>" & vbCrLf
				.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbCrLf
				.Write "<link href='Include/admin_style.css' rel='stylesheet'>" & vbCrLf
				.Write "<script language='JavaScript' src='../KS_Inc/Common.js'></script>" & vbCrLf
				.Write "<script language='JavaScript' src='../KS_Inc/Jquery.js'></script>" & vbCrLf
				.Write "<script src=""images/pannel/tabpane.js"" language=""JavaScript""></script>" & vbCrlf
				.Write "<link href=""images/pannel/tabpane.CSS"" rel=""stylesheet"" type=""text/css"">" & vbCrlf
				.Write "<script language='Javascript'>" & vbcrlf
				
				If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
				.Write "var marr = new Array();" & vbCrlf
				K=0
				Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
				For Each Node In ModelXML.documentElement.SelectNodes("channel[@ks21=1 and @ks0!=6 and @ks0!=9 and @ks0!=10 || @ks0=5]")
				.Write "marr[" & K & "] = new Array('" & Node.SelectSingleNode("@ks0").text & "','" & Node.SelectSingleNode("@ks1").text & "');" & vbCrlf
				K=K+1
				Next
				.Write "</script>" & vbcrlf
				.Write "</head>" & vbCrLf
				.Write "<body leftmargin='0' topmargin='0' marginwidth='0' marginheight='0'>" & vbCrLf
				.Write "<div class=""topdashed sort"">" & TopTitle & "</div>" & vbCrLf
				.Write "<br>"
				
				

				.Write "  <table width='100%' style='margin-top:2px' border='0' align='center' cellpadding='0' cellspacing='0'>" & vbCrLf
				.Write " <form  action='" & FormProcesPage & "' method='post' name='CreateFolderForm'>" & vbCrLf
				.Write "    <tr>" & vbCrLf
				.Write "      <td valign=top>" & vbCrLf
				
				.Write "<div class=tab-page id=ClassPane>"
				.Write " <SCRIPT type=text/javascript>"
				.Write "   var tabPane1 = new WebFXTabPane( document.getElementById( ""ClassPane"" ), 1 )"
				.Write " </SCRIPT>"
				 
				.Write " <div class=tab-page id=site-page>"
				.Write "  <H2 class=tab>������Ϣ</H2>"
				.Write "	<SCRIPT type=text/javascript>"
				.Write "				 tabPane1.addTabPage( document.getElementById( ""site-page"" ) );"
				.Write "	</SCRIPT>"

				'������Ϣ����
				.Write "      <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">" & vbCrLf
				
						 ' If FolderID <> "0" Then
				.Write "          <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
				.Write "            <td  width='200' height='30' align='right' class='clefttitle'><strong>������Ŀ��</strong></td>" & vbCrLf
				.Write "           <td height='28'>&nbsp;"
				If KS.G("Action")="Edit" Then
				.Write "<input type='hidden' name='parentid' value='" & FolderID & "'>"
				.Write "<select name='parentID1' Disabled>" & vbCrLf
				Else
				.Write "<select onchange='setchannel(this.value)' name='parentID'>" & vbCrLf
				End If
				If ChannelID<>8 Then
				.Write "<option value='0'>�ޣ���ΪƵ��)</option>" & vbcrlf
				End If
				.Write TypeList & " </select>" & vbcrlf
				
				.Write "</td>" & vbCrLf
				.Write "          </tr>"
				         'End If
						 
				.Write "        <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
				.Write "<td height='28'   width='200' align=right class='clefttitle'><strong>��ģ�ͣ�</strong>"
				.Write "         </td>"
				.Write  "<td>"
				If Action="Edit" Then
				.Write "   &nbsp;<input type='hidden' name='ChannelID' value='" & ChannelID & "'><select Disabled name='ChannelIDs' class='upfile' onchange='changemodel(this.value)'>" & vbCrLf
				Else
				.Write "   &nbsp;<select name='ChannelID' class='upfile' onchange='changemodel(this.value)'>" & vbCrLf
				End If
				
				
				For Each Node In ModelXML.documentElement.SelectNodes("channel[@ks21=1 and @ks0!=6 and @ks0!=9 and @ks0!=10 ||@ks0=5]")
				  If trim(ChannelID)=trim(Node.SelectSingleNode("@ks0").text) Then
				  .Write "<option value='" & Node.SelectSingleNode("@ks0").text & "' selected>" & Node.SelectSingleNode("@ks1").text & "|" & Node.SelectSingleNode("@ks2").text & "</option>"
				  Else
				  .Write "<option value='" & Node.SelectSingleNode("@ks0").text & "'>" & Node.SelectSingleNode("@ks1").text & "|" & Node.SelectSingleNode("@ks2").text & "</option>"
				  End If
				Next
				
				             
				.Write "             </select> ��ѡ�����ĿҪ�󶨵�ģ��" & vbCrLf 
				.Write "            </td>" & vbCrLf
				.Write "        </tr>" & vbCrLf	
				
						  
				.Write "          <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
				.Write "            <td height='30'  width='200' align='right' class='clefttitle'><strong>��Ŀ���ƣ�</strong></td>" & vbCrLf
				.Write "            <td height='28'>" & vbCrLf
				.Write "             <label id='add1'>"
				.Write "              &nbsp;<INPUT name='FolderName' onkeyup='ctoe()' type='text' value='" & FolderName & "' id='FolderName' title='��������Ŀ����' size=30><font color=red>*</font> �����Ե�˵������</label>"
				.Write "             <div id='add2' style='display:none;color:blue'><strong>¼���ʽ:</strong>��Ŀ��������|Ӣ������,˵��ÿ��һ��<br/>"
				.Write "             <textarea id='FolderNames' name='FolderNames' style='width:300px;height:150px'>��Ŀ����1|Ӣ������1</textarea>"
				
				
				.Write "             </div>"
				
				If KS.G("Action")<>"Edit" Then
				.Write "<label><input type='checkbox' onclick='ChangeAddMode()' name='AddMore' id='AddMore' value='1'><font color=red><strong>�л����������ģʽ</strong></font></label>"
				End If
				
				.Write "</td>" & vbCrLf
				.Write "          </tr>" & vbCrLf
				.Write "          <tr id='typearea' class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
				.Write "           <td height='30' align='right'  width='200' class='clefttitle'><strong>��Ŀ���ͣ�</strong></td>" & vbCrLf
				.Write "            <td height='28'>" & vbCrLf
				If Action="Edit" Then
				 .Write "&nbsp;<font color=red>["
				  Select Case ClassType
				   Case "1": .Write "ϵͳ��Ŀ"
				   Case "2": .Write "�ⲿ����"
				   Case "3": .Write "��ҳ��"
				  End Select
				  .Write "]</font>"
				Else
				.Write "             <input type='radio' onclick='changetype(this.value)' name='classtype' value='1'"
				If ClassType="1" Then .Write " checked"
				.Write ">ϵͳ��Ŀ"
				.Write "             <input type='radio' onclick='changetype(this.value)' name='classtype' value='2'"
				If ClassType="2" Then .Write " checked"
				.Write ">�ⲿ����"
				.Write "             <input type='radio' onclick='changetype(this.value)' name='classtype' value='3'"
				If ClassType="3" Then .Write " checked"
				.Write ">��ҳ��"
				End If
				.Write "            <br><span id='classarea'>Ӣ�����ƣ�</span>" &vbcrlf
				If Action="Edit" and  ClassType<>2 Then
				.Write "             <input Disabled name='FolderEname1' class='upfile' type='text' id='FolderEname1' value='" & FolderEname & "' size=30>"
				.Write "             <input style='display:none' class='upfile' name='FolderEname' type='text' id='FolderEname' value='" & FolderEname & "' size=30>"
			    Else
				.Write "              <input name='FolderEname' class='upfile' type='text' id='FolderEname' value='" & FolderEname & "' size=30>"
				End If
				.Write "             <font color=red>*</font><span id='classtips'>���ܴ�\/��*���� < > | ���������,���趨���ܸ�</span></td>" & vbCrLf
				.Write "          </tr>" & vbCrLf
				
				.Write "          <tbody id='templatearea'>"
				.Write "          <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
				.Write "            <td height='30' align='right'  width='200' class='clefttitle'><strong>" & vbCrLf
				
					   If FolderID = "0" Then  .Write ("��Ŀ��ҳģ�壺")  Else  .Write ("��Ŀģ�壺")
					   
				.Write "</strong> </td>" & vbCrLf
				.Write "            <td height='28'><b>" & vbCrLf
				.Write "              &nbsp;<input type='text' id='FolderTemplateID' name='FolderTemplateID' value='" & FolderTemplateID & "' size=35>&nbsp;" & KSCls.Get_KS_T_C("document.getElementById('FolderTemplateID')")
				.Write "         </td></tr>" & vbCrLf
				.Write "          <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
				.Write "            <td height='30' align='right'  width='200' class='clefttitle'><strong>" & vbCrLf
				
					   If FolderID = "0" Then  .Write ("WAP��Ŀ��ҳģ�壺")  Else  .Write ("WAP��Ŀģ�壺")
					   
				.Write "</strong> </td>" & vbCrLf
				.Write "            <td height='28'><b>" & vbCrLf
				.Write "              &nbsp;<input type='text' id='WAPFolderTemplateID' name='WAPFolderTemplateID' value='" & WAPFolderTemplateID & "' size=35>&nbsp;" & KSCls.Get_KS_T_C("document.getElementById('WAPFolderTemplateID')")
				.Write "         </td></tr>" & vbCrLf
				
				
				
				.Write "         <tbody id='temparea'>" & vbcrlf

					 .Write "         <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
					 .Write "           <td height='30' align='right'  class='clefttitle' width='200'><strong>����ҳģ�壺</strong></td>" & vbCrLf
					 .Write "           <td height='28'>" & vbCrLf
					 .Write "              &nbsp;<input type='text' id='TemplateID' name='TemplateID' value='" & TemplateID & "' size='35'>&nbsp;" & KSCls.Get_KS_T_C("document.getElementById('TemplateID')")										  
					.Write "    </td></tr>"
					 .Write "         <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
					 .Write "           <td height='30' align='right'  class='clefttitle' width='200'><strong>WAP����ҳģ�壺</strong></td>" & vbCrLf
					 .Write "           <td height='28'>" & vbCrLf
					 .Write "              &nbsp;<input type='text' id='TemplateID' name='WAPTemplateID' value='" & WAPTemplateID & "' size='35'>&nbsp;" & KSCls.Get_KS_T_C("document.getElementById('WAPTemplateID')")										  
					.Write "    </td></tr>"
					
					
					
					If FolderID="0" Then
					.Write "<tbody id='channel'>"
					Else
					.Write "<tbody id='channel' style='display:none'>"
					End If
					.Write "  <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
					.Write "   <td height='30' align='right' class='clefttitle'><strong>������<font color='#FF0000'>(������)</font>��</strong></td>" & vbCrLf
					.Write "     <td><b>&nbsp;<input name='FolderDomain' TYPE='text' value='" & FolderDomain & "' id='FolderDomain' class='upfile' size=30></b>&nbsp;�磺http://news.kesion.com/��ֻ��һ����Ŀ��Ч </td>" & vbCrLf
					.Write " </tr>" & vbCrLf
					.Write "</tbody>"
					
					
                    .Write "<tbody id=""fsohtmlarea"">"
				.Write "         <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'""><td align=right  class='clefttitle' width='200'><strong>" & vbCrLf
				.Write "             ���ɵ���Ŀ��ҳ�ļ���</strong>" & vbCrLf
				.Write "</td>"
				.Write "<td>"
					 .Write "             &nbsp;<select name='FolderFsoIndex' class='upfile'>" & vbCrLf
					 .Write "               <option value='index.html'>index.html</option>" & vbCrLf
					 .Write "               <option value='index.htm' selected>index.htm</option>" & vbCrLf
					 .Write "               <option value='index.shtm'>index.shtm</option>" & vbCrLf
					 .Write "               <option value='index.shtml'>index.shtml</option>" & vbCrLf
					 .Write "               <option value='default.html'>default.html</option>" & vbCrLf
					 .Write "               <option value='default.htm'>default.htm</option>" & vbCrLf
					 .Write "               <option value='default.shtm'>default.shtm</option>" & vbCrLf
					 .Write "               <option value='default.shtml'>default.shtml</option>" & vbCrLf
					 .Write "               <option value='index.asp'>index.asp</option>" & vbCrLf
					 .Write "               <option value='default.asp'>default.asp</option>" & vbCrLf
					 .Write "               <option value=""" & FolderFsoIndex & """ selected>" & FolderFsoIndex & "</option>"
					 .Write "             </select>" & vbCrLf
					 .Write "             </td>" & vbCrLf
					 .Write "         </tr>" & vbCrLf
					 .Write "        <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
					.Write "<td height='28'   width='200' align=right class='clefttitle'><strong>"
					.Write "               ����ҳ���ɵ���չ����</strong>"
					.Write "         </td>"
					.Write  "<td>"
					 .Write "             &nbsp;<input type='text' ID='FnameType' name='FnameType' value='" & FnameType & "' size='15'> <-<select name='FnameTypes'  class='upfile' onchange=""$('#FnameType').val(this.value);"">" & vbCrLf
					 .Write "               <option value='.html' selected>.html</option>" & vbCrLf
					 .Write "               <option value='.htm'>.htm</option>" & vbCrLf
					 .Write "               <option value='.shtm'>.shtm</option>" & vbCrLf
					 .Write "               <option value='.shtml'>.shtml</option>" & vbCrLf
					 .Write "               <option value='.asp'>.asp</option>" & vbCrLf
					 .Write "             </select>" & vbCrLf
					  .Write "            </td>" & vbCrLf
					  .Write "        </tr>" & vbCrLf
					  .Write "        <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
					  .Write "          <td height='30' align='right' width='200' class='clefttitle'><strong>����ҳ����HTML��ʽ��</strong></td>" & vbCrLf
					  .Write "          <td height='28'> &nbsp;<select style='width:200;' name='FsoType' id='select5' onChange='SelectFsoType(options[selectedIndex].value);'>" & vbCrLf
							   If FsoType = 1 Then SelStr = " Selected"  Else SelStr = ""
							   .Write ("<option value=""1""" & SelStr & ">" & YearStr & "/" & MonthStr & "-" & DayStr & "/RE</option>")
							  If FsoType = 2 Then SelStr = " Selected" Else SelStr = ""
							   .Write ("<option value=""2""" & SelStr & ">" & YearStr & "/" & MonthStr & "/" & DayStr & "/RE</option>")
							  If FsoType = 3 Then SelStr = " Selected" Else SelStr = ""
							   .Write ("<option value=""3""" & SelStr & ">" & YearStr & "-" & MonthStr & "-" & DayStr & "/RE</option>")
							  If FsoType = 4 Then SelStr = " Selected" Else SelStr = ""
							   .Write ("<option value=""4""" & SelStr & ">" & YearStr & "/" & MonthStr & "/RE</option>")
							  If FsoType = 5 Then SelStr = " Selected"  Else	SelStr = ""
							  .Write ("<option value=""5""" & SelStr & ">" & YearStr & "-" & MonthStr & "/RE</option>")
							  If FsoType = 12 Then SelStr = " Selected"  Else	SelStr = ""
							  .Write ("<option value=""12""" & SelStr & ">" & YearStr & MonthStr & "/RE</option>")
							  If FsoType = 6 Then SelStr = " Selected" Else	SelStr = ""
							  .Write ("<option value=""6""" & SelStr & ">" & YearStr & MonthStr & DayStr & "/RE</option>")
							  If FsoType = 7 Then SelStr = " Selected" Else	SelStr = ""
							  .Write ("<option value=""7""" & SelStr & ">" & YearStr & "/RE</option>")
							  If FsoType = 8 Then SelStr = " Selected" Else SelStr = ""
							  .Write ("<option value=""8""" & SelStr & ">" & YearStr & MonthStr & DayStr & "RE</option>")
							  If FsoType = 9 Then SelStr = " Selected" Else SelStr = ""
							  .Write ("<Option value=""9""" & SelStr & ">RE</Option>")
							  If FsoType = 10 Then SelStr = " Selected"  Else SelStr = ""
							  .Write ("<option value=""10""" & SelStr & ">SCE</option>")
							  If FsoType = 11 Then SelStr = " Selected"  Else SelStr = ""
							  .Write ("<option value=""11""" & SelStr & ">�ĵ�IDE</option>")

					  .Write "            </select> </td>"
					  .Write "        </tr>" & vbCrLf
					  .Write "        <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
					  .Write "          <td height='30' colspan='3' align='right'> <div align='center'><strong><span id='ShowAS1'></Span></strong> </div></td>" & vbCrLf
					  .Write "        </tr>" & vbCrLf
					  .Write "</tbody>" &vbcrlf
				      .Write "     </tbody>" &vbcrlf
					  .Write "     </tbody>" & vbcrlf
					  
				 .Write "         <tr id=""editorarea"" class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
				 .Write "           <td height='50' align='right' width='200' class='clefttitle'><strong>��ҳ���ݣ�</strong><br><font color='#ff000000'>ʹ�ñ�ǩ{$GetClassIntro}��ģ�������</font></td>" & vbCrLf
				 
				 .Write "           <td height='28'> "
				 .Write "<textarea id='ClassContent' name='ClassContent' style='display:none'>"& Server.HTMLEncode(ClassContent) &"</textarea>"
				 .Write "<span id='singlepage'></span>"
				 .Write "            </td>" & vbCrLf
				 .Write "         </tr>" & vbCrLf


				 
				 .Write "       </table>" & vbCrLf
				 .Write "</div>"
				 
				.Write " <div class=tab-page id=classoption-page>"
				.Write "  <H2 class=tab>��Ŀѡ��</H2>"
				.Write "	<SCRIPT type=text/javascript>"
				.Write "				 tabPane1.addTabPage( document.getElementById( ""classoption-page"" ) );"
				.Write "	</SCRIPT>"

				 'Ƶ������Ŀ��ѡ��
				 .Write " <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">" & vbCrLf
				.Write "          <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
				.Write "            <td height='30' width='200' align='right' class='clefttitle'><strong>��Ŀ����������</strong></td>" & vbCrLf
				.Write "            <td height='28'>&nbsp;"
					       If TopFlag = 1 Then
						   .Write ("<input name=""TopFlag"" type=""radio"" value=""1"" checked>")
						   Else
						   .Write ("<input name=""TopFlag"" type=""radio"" value=""1"">")
						   End If
							.Write ("��ʾ ")
							If TopFlag = 0 Then
						   .Write ("<input name=""TopFlag"" type=""radio"" value=""0"" checked>")
						   Else
						   .Write ("<input name=""TopFlag"" type=""radio"" value=""0"">")
						   End If
						 .Write "����ʾ"
					.Write "              </td>"
					.Write "          </tr>" & vbCrLf
			        .Write "          <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
				    .Write "            <td height='30' width='200' align='right' class='clefttitle'><strong>��ĿWAP״̬��</strong></td>" & vbCrLf
				    .Write "            <td height='28'>&nbsp;"
					       If WapSwitch = 1 Then
						   .Write ("<input name=""WapSwitch"" type=""radio"" value=""1"" checked>")
						   Else
						   .Write ("<input name=""WapSwitch"" type=""radio"" value=""1"">")
						   End If
							.Write ("��ʾ ")
							If WapSwitch = 0 Then
						   .Write ("<input name=""WapSwitch"" type=""radio"" value=""0"" checked>")
						   Else
						   .Write ("<input name=""WapSwitch"" type=""radio"" value=""0"">")
						   End If
						 .Write "����ʾ"
					.Write "              </td>"
					.Write "          </tr>" & vbCrLf

				 .Write "          <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
				.Write "           <td height='40' align='right'  width='200' class='clefttitle'><strong>��ĿͼƬ��ַ��</strong><br>��������Ŀҳ��ʾָ����ͼƬ </td>" & vbCrLf
				.Write "            <td height='28'>" & vbCrLf
				.Write "              &nbsp;<INPUT NAME='ClassPic' value='" & ClassPic &"' TYPE='text' id='ClassPic' class='upfile' size=30>"
				.Write "                  <input class=""button""  type='button' name='Submit' value='ѡ��ͼƬ...' onClick=""OpenThenSetValue('Include/SelectPic.asp?ChannelID=" & ChannelID &"&CurrPath=" & CurrPath & "',550,290,window,document.CreateFolderForm.ClassPic);"">  <input class=""button"" type='button' name='Submit' value='Զ��ץȡͼƬ...' onClick=""OpenThenSetValue('Include/Frame.asp?FileName=SaveBeyondfile.asp&PageTitle='+escape('ץȡԶ��ͼƬ')+'&ItemName=ͼƬ&CurrPath=" & CurrPath & "',300,100,window,document.CreateFolderForm.ClassPic);"">"
				.Write "              </td>" & vbCrLf
				.Write "          </tr>" & vbCrLf
				If ClassType=3 and Action="Edit" then
				 .Write "         <tr style='display:none' class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
                Else
				 .Write "         <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
				End if
				 .Write "           <td height='50' align='right'  width='200' class='clefttitle'><strong>��Ŀ���ܣ�</strong><br>" & vbCrLf
				 .Write "             <font color='#0000FF'>��������Ŀҳ��ϸ������Ŀ��Ϣ��֧��HTML<br>���ڶ�Ӧ����Ŀģ��ҳʹ�ñ�ǩ<br><font color=red>""{$GetClassIntro}""</font> ���е���</font></font></td>" & vbCrLf
				 .Write "           <td height='28'>" & vbCrLf
				 .Write "             &nbsp;<textarea name='ClassDescript' id='ClassDescript' class='upfile' cols='60' rows='5'>" & Server.Htmlencode(ClassDescript) & "</textarea>"
				 .Write "             </td>" & vbCrLf
				 .Write "         </tr>" & vbCrLf		 
							  
				 .Write "         <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
				 .Write "           <td height='50' align='right' width='200' class='clefttitle'><strong>��ĿMETA�ؼ��ʣ�</strong><br>" & vbCrLf
				 .Write "             <font color='#0000FF'>�������������������Ĺؼ���<br>���ڶ�Ӧ����Ŀģ��ҳʹ�ñ�ǩ<br><font color=red>""{$GetClass_Meta_KeyWord}""</font> ���е���</font></td>" & vbCrLf
				 .Write "           <td height='28'>" & vbCrLf
				 .Write "             &nbsp;<textarea name='MetaKeyWord' id='MetaKeyWord' class='upfile' cols='60' rows='5'>" & MetaKeyWord & "</textarea>"
				 .Write "             </td>" & vbCrLf
				 .Write "         </tr>" & vbCrLf
				 
				  .Write "         <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
				 .Write "           <td height='50' align='right' width='200' class='clefttitle'><strong>��ĿMETA��ҳ������</strong><br>" & vbCrLf
				 .Write "             <font color='#0000FF'>����������������������ҳ����<br>���ڶ�Ӧ����Ŀģ��ҳʹ�ñ�ǩ<br><font color=red>""{$GetClass_Meta_Description}""</font> ���е���</font></font></td>" & vbCrLf
				 .Write "           <td height='28'>" & vbCrLf
				 .Write "             &nbsp;<textarea name='MetaDescript' id='MetaDescript' class='upfile' cols='60' rows='5'>" & MetaDescript & "</textarea>"
				 .Write "             </td>" & vbCrLf
				 .Write "         </tr>" & vbCrLf
				
				
                 .Write "<tbody id='ShowAD'>" & vbcrlf
						 .Write "         <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
						 .Write "           <td height='50' align='right'  class='clefttitle'><strong>��������ʾ���л���</strong></td>" & vbCrLf
						 .Write "           <td height='28'>" & vbCrLf
                          if ShowADTF = "1" Then
						   .Write ("<input onclick=""$('#Ad').show();"" name=""ShowADTF"" type=""radio"" value=""1"" checked>")
						   Else
						   .Write ("<input onclick=""$('#Ad').show();"" name=""ShowADTF"" type=""radio"" value=""1"">")
						   End If
							.Write ("��ʾ ")
						   If ShowADTF = "0" Then
						   .Write ("<input onclick=""$('#Ad').hide();"" name=""ShowADTF"" type=""radio"" value=""0"" checked>")
						   Else
						   .Write ("<input onclick=""$('#Ad').hide();"" name=""ShowADTF"" type=""radio"" value=""0"">")
						   End If
						 .Write "����ʾ"
						 
					
                         .Write " <table"
						 If ShowADTF="0" Then .Write "  style=""display:none"""
						 .Write " id=""Ad"" class=""border"" style=""margin:5px"" border=""0"" align=""center"" cellpadding=""5"" cellspacing=""1"">"
                         .Write "<tr class=""tdbg"">"
                         .Write "<td width=""22%""><div align=""right"">���л��������ã�</div></td>"
                         .Write "<td width=""78%""><input class=""textbox"" name=""AdParam"" type=""text"" id=""AdParam"" size=""20"" maxlength=""20"" value=""" & AdParam & """>(����λ��������ǰ������,��(left)��(right),���,�߶ȣ���500,left,300,300)</td>"
                         .Write "</tr>"
						 .Write "<tr class=""tdbg"">"
						 .Write "<td><div align=""right"">������ͣ�</div></td>"
						 .Write "<td>"
						 if ADType = "1" Then
						   .Write ("<input onclick=""$('#adcodearea').hide();$('#adimgarea').show();"" name=""ADType"" type=""radio"" value=""1"" checked>")
						   Else
						   .Write ("<input onclick=""$('#adcodearea').hide();$('#adimgarea').show();"" name=""ADType"" type=""radio"" value=""1"">")
						   End If
							.Write ("ͼƬ/Flash ")
							If ADType = "2" Then
						   .Write ("<input onclick=""$('#adimgarea').hide();$('#adcodearea').show();"" name=""ADType"" type=""radio"" value=""2"" checked>")
						   Else
						   .Write ("<input onclick=""$('#adimgarea').hide();$('#adcodearea').show();"" name=""ADType"" type=""radio"" value=""2"">")
						   End If
						 .Write "�����棨֧��Google���)"
						 .Write "</td>"
						 .Write "</tr>"
						 
						 if ADType="1" Then
						 .Write "<tbody id='adcodearea' style='display:none'>"
						 Else
						 .Write "<tbody id='adcodearea'>"
						 End IF
                         .Write "<tr class=""tdbg"">"
                         .Write "<td><div align=""right"">�����룺<br><font color=red>֧��HTML�﷨</font></div></td>"
                         .Write "<td><textarea style='height:60px' name=""AdCode"" class=""textbox"" cols='60' rows=6>" & AdUrl & "</textarea>"
			             .Write "</td></tr>"
						 .Write "</tbody>"
						 
						 if ADType="2" Then
						 .Write "<tbody id='adimgarea' style='display:none'>"
						 Else
						 .Write "<tbody id='adimgarea'>"
						 End IF
                         .Write "<tr class=""tdbg"">"
                         .Write "<td><div align=""right"">ͼƬ��ַ��</div></td>"
                         .Write "<td><input name=""AdUrl"" class=""textbox"" type=""text"" id=""AdUrl""  size=""36"" maxlength=""250"" value=""" & AdUrl & """>"
                         .Write " <input class=""button""  type='button' name='Submit' value='ѡ��ͼƬ��FLASH' onClick=""OpenThenSetValue('Include/SelectPic.asp?ChannelID=" & ChannelID &"&CurrPath=" & CurrPath & "',550,290,window,document.CreateFolderForm.AdUrl);""> "
			             .Write "</td></tr>"
						 .Write "<tr class=""tdbg"">"
						 .Write "<td><div align=""right"">���ӵ�ַ��</div></td>"
						 .Write "<td><input name=""AdLinkUrl"" type=""text"" class=""textbox"" id=""AdLinkUrl""  size=""36"" maxlength=""250"" value=""" & AdLinkUrl & """>����ͼƬ��Ч</td>"
                         .Write "</tr>"
						 .Write "</tbody>"
						 
                         .Write " </table>"
						 
						 .Write "              </td>" & vbCrLf
						 .Write "         </tr>" & vbCrLf
                  .Write "</tbody>"

				  .Write "       </table>" & vbCrLf 
				  .Write "</div>"
				
				If ChannelID<>5 Then
				.Write " <div class=tab-page id=poweroption-page>"
				.Write "  <H2 class=tab>Ȩ��ѡ��</H2>"
				.Write "	<SCRIPT type=text/javascript>"
				.Write "				 tabPane1.addTabPage( document.getElementById( ""poweroption-page"" ) );"
				.Write "	</SCRIPT>"

				 'Ȩ���շ�ѡ������
				 .Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">" & vbCrLf
				.Write "          <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
				.Write "            <td height='80'  width='200' align='right' class='clefttitle'><strong>���/�鿴Ȩ�ޣ�</strong></td>" & vbCrLf
				.Write "            <td height='28'>&nbsp;"
					 If ClassPurview=0 Then SelStr=" checked" Else SelStr=""
				.Write "<input name='ClassPurview' type='radio' value='0'" & SelStr &">"
				.Write "              ������Ŀ&nbsp;&nbsp;<font color=red>�κ��ˣ������οͣ���������Ͳ鿴����Ŀ�µ���Ϣ��</font><br>"
				If  ChannelID<>8 Then
					 If ClassPurview=1 Then SelStr=" checked" Else SelStr=""
				.Write "              &nbsp;<INPUT type='radio'  name='ClassPurview' value='1'" & SelStr &">" & vbCrLf
				.Write "              �뿪����Ŀ&nbsp;&nbsp;<font color=red>�κ��ˣ������οͣ�������������οͲ��ɲ鿴��������Ա���ݻ�Ա�����ĿȨ�����þ����Ƿ���Բ鿴��</font><br/>"
				End If
					 If ClassPurview=2 Then SelStr=" checked" Else SelStr=""
				.Write "              &nbsp;<INPUT type='radio'  name='ClassPurview' value='2'" & SelStr &">" & vbCrLf
				If  ChannelID<>8 Then
				.Write "              ��֤��Ŀ&nbsp;&nbsp;<font color=red>�οͲ�������Ͳ鿴��������Ա���ݻ�Ա�����ĿȨ�����þ����Ƿ��������Ͳ鿴��</font><br>"
				Else
				.Write "              ��֤��Ŀ&nbsp;&nbsp;<font color=red>ֻ��ָ���Ļ�Ա��ſ��Բ鿴������Ϣ����ϵ��ʽ��</font><br>"
				End If
				.Write "</td>"
				.Write "          </tr>" & vbCrLf

				.Write "          <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
				If  ChannelID<>8 Then
				.Write "            <td height='80' align='right' width='200' class='clefttitle'><div><strong>����鿴����Ŀ����Ϣ�Ļ�Ա�飺</strong></div><font color=blue>�����Ŀ�ǡ���֤��Ŀ�������ڴ���������鿴����Ŀ����Ϣ�Ļ�Ա��,�������Ϣ�������˲鿴Ȩ�ޣ�������Ϣ�е�Ȩ����������</font></td>" & vbCrLf
				Else
				.Write "            <td height='80' align='right' width='200' class='clefttitle'><div><strong>����鿴����Ŀ�¹�����Ϣ��ϵ��ʽ�Ļ�Ա�飺</strong></div><font color=blue>�����Ŀ�ǡ���֤��Ŀ�������ڴ���������鿴����Ŀ����Ϣ����ϵ��ʽ�Ļ�Ա��,��Ӧ�Ĺ�������ҳģ������Ҫ��[KS_Charge]��ϵ��ʽ[/KS_Chagrge]����Ч��</font></td>" & vbCrLf
				End If
				
				.Write "            <td height='28'>&nbsp;" & KS.GetUserGroup_CheckBox("GroupID",DefaultArrGroupID,5)
				.Write "</td>"
				.Write "          </tr>" & vbCrLf
				If  ChannelID<>8 Then
				.Write "          <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
				.Write "            <td height='60' align='right' width='200' class='clefttitle'><strong>Ĭ���Ķ���Ϣ���������</strong><br><font color=blue>�������Ϣ���������Ķ�������������Ϣ�еĵ�����������</font></td>" & vbCrLf
				.Write "            <td height='28'>&nbsp;<input name='ReadPoint' type='text' id='ReadPoint'  value='" & ReadPoint & "' size='6' class='upfile' style='text-align:center'>����Ķ�����Ϊ""<font color=red>0</font>""��������Ȩ�޵Ļ�Ա�Ķ�����Ŀ�µ���Ϣʱ��������Ӧ�������οͽ��޷��Ķ���"
				.Write "</td>"
				.Write "          </tr>" & vbCrLf
				.Write "          <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
				.Write "            <td height='60' align='right' width='200' class='clefttitle'><strong>Ĭ����Ͷ���ߵķֳɱ��ʣ�</strong></td>" & vbCrLf
				.Write "            <td height='28'>&nbsp;<input name='DividePercent' type='text' value='" & DividePercent & "' size='6' class='upfile' style='text-align:center'>% ϵͳ�������������õķֳɱ��ʽ��ճɷָ�Ͷ���ߡ��������10��������!"
				.Write "</td>"
				.Write "          </tr>" & vbCrLf
				.Write "          <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
				.Write "            <td height='30' align='right' width='200' class='clefttitle'><strong>Ĭ���Ķ���Ϣ�ظ��շѣ�</strong><br><font color=blue>�������Ϣ���������Ķ�������������Ϣ�еĵ�����������</font></td>" & vbCrLf
				.Write "            <td height='28'>&nbsp;"
				.Write "<input name='ChargeType' type='radio' value='0' "
					 IF ChargeType=0 Then .Write " checked"
				.Write" >���ظ��շ�(�����Ϣ��۵������ܲ鿴������ʹ��)<br>"
				.Write "&nbsp;<input name='ChargeType' type='radio' value='1'"
					 IF ChargeType=1 Then .Write " checked"
				.write ">�����ϴ��շ�ʱ�� <input name='PitchTime' type='text' class='upfile' value='" & PitchTime & "' size='8' maxlength='8' style='text-align:center'> Сʱ�������շ�<br>            &nbsp;<input name='ChargeType' type='radio' value='2'"
					 IF ChargeType=2 Then .Write " checked"
				.write ">��Ա�ظ�����Ϣ &nbsp;<input name='ReadTimes' type='text' class='upfile' value='" & ReadTimes & "' size='8' maxlength='8' style='text-align:center'> ҳ�κ������շ�<br>            &nbsp;<input name='ChargeType' type='radio' value='3'"
					 IF ChargeType=3 Then .Write " checked"
				.write ">�������߶�����ʱ�����շ�<br>            &nbsp;<input name='ChargeType' type='radio' value='4'"
					 IF ChargeType=4 Then .Write " checked"
				.write ">����������һ������ʱ�������շ�<br>            &nbsp;<input name='ChargeType' type='radio' value='5'"
					 IF ChargeType=5 Then .Write " checked"
				.write ">ÿ�Ķ�һҳ�ξ��ظ��շ�һ�Σ����鲻Ҫʹ��,��ҳ��Ϣ���۶�ε�����"
				.Write "</td>"
				.Write "          </tr>" & vbCrLf
                  End If
				 .Write "       </table>" & vbCrLf 
				 .Write "</div>" & vbcrlf
				 End If
				 
				.Write " <div class=tab-page id=classtg-page>"
				.Write "  <H2 class=tab>Ͷ��ѡ��</H2>"
				.Write "	<SCRIPT type=text/javascript>"
				.Write "				 tabPane1.addTabPage( document.getElementById( ""classtg-page"" ) );"
				.Write "	</SCRIPT>"

				.Write "<table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">" & vbCrLf
				.Write "          <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
				.Write "            <td height='30' align='right' width='200' class='clefttitle'><strong>��Ŀ�Ƿ�����Ͷ�壺</strong></td>" & vbCrLf
				.Write "            <td height='28'>"
						If CommentTF = 0 Then
						   .Write ("��<input name=""CommentTF"" type=""radio"" value=""0"" checked>������<br>")
						Else
						   .Write ("��<input name=""CommentTF"" type=""radio"" value=""0"">������<br>")
						End If
						if CommentTF = 1 Then
						   .Write ("��<input name=""CommentTF"" type=""radio"" value=""1"" checked>�������л�ԱͶ��<font color=blue>(�οͳ���)</font><br>")
						Else
						   .Write ("��<input name=""CommentTF"" type=""radio"" value=""1"">�������л�ԱͶ��<font color=blue>(�οͳ���)</font><br>")
						End If
						if CommentTF = 2 Then
						   .Write ("��<input name=""CommentTF"" type=""radio"" value=""2"" checked>����������Ͷ��<font color=red>(�����ο�)</font><br>")
						Else
						   .Write ("��<input name=""CommentTF"" type=""radio"" value=""2"">����������Ͷ��<font color=red>(�����ο�)</font><br>")
						End If
						if CommentTF = 3 Then
						   .Write ("��<input name=""CommentTF"" type=""radio"" value=""3"" checked>ֻ����ָ���û���Ļ�ԱͶ��<br>")
						Else
						   .Write ("��<input name=""CommentTF"" type=""radio"" value=""3"">ֻ����ָ���û���Ļ�ԱͶ��<br>")
						End If

					
				.Write "              </td>"
				.Write "          </tr>"
				.Write "          <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">" & vbCrLf
				.Write "            <td height='80' align='right' width='200' class='clefttitle'><strong>�������Ŀ��Ͷ��Ļ�Ա�飺</strong><br><font color=blue>������ѡ���ʱ�����ڴ����������ڴ���Ŀ��Ͷ��Ļ�Ա��</font></td>" & vbCrLf
				.Write "            <td height='28'>&nbsp;" & KS.GetUserGroup_CheckBox("AllowArrGroupID",AllowArrGroupID,5)
				.Write "</td>"
				.Write "          </tr>" & vbCrLf
				.Write "</table>"
				.Write "</div>" 

				 
				.Write " <div class=tab-page id=defineoption-page>"
				.Write "  <H2 class=tab>����ѡ��</H2>"
				.Write "	<SCRIPT type=text/javascript>"
				.Write "				 tabPane1.addTabPage( document.getElementById( ""defineoption-page"" ) );"
				.Write "	</SCRIPT>"

				 .Write " <table width=""100%"" border=""0"" align=""center"" cellpadding=""1"" cellspacing=""1"" class=""ctable"">" & vbCrLf
				  .Write "         <tr class=""tdbg"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
				 .Write "           <td height='30' align='right' width='210' class='clefttitle'><strong>������������</strong>" & vbCrLf
				 .Write "             </td>" & vbCrLf
				 .Write "           <td height='28'>" & vbCrLf
				 .Write "              &nbsp;<select name=""ClassDefine_Num"" onChange=""setFileFileds(this.value)"">"
				  Dim DefineNum,SelDefineNum
				  If IsArray(ClassDefineContentArr) Then SelDefineNum=Ubound(ClassDefineContentArr)+1 Else SelDefineNum=1
				  For DefineNum=1 To 20
				   If DefineNum=SelDefineNum Then
				    .Write "<option value=""" & DefineNum & """ selected>" & DefineNum & "</option>"
				   Else
				    .Write "<option value=""" & DefineNum & """>" & DefineNum & "</option>"
				   End If
                  Next
			     .Write " </select>"
				 .Write "             </td>" & vbCrLf
				 .Write "         </tr>" & vbCrLf
				
				 For DefineNum=1 To 20
				 .Write "        <tr class=""tdbg"" id='objFiles" & DefineNum & "' style=""display:none"" onMouseOver=""this.className='tdbgmouseover'"" onMouseOut=""this.className='tdbg'"">"
				 .Write "           <td height='30' align='right' width='210' class='clefttitle'><strong>��������" & DefineNum & "��</strong><br> <font color=blue>����Ŀģ��ҳ����{$GetClassDefineContent" & DefineNum & "} ����</font>" & vbCrLf
				 .Write "             </td>" & vbCrLf
				 
				  If Action="Edit" Then
				     IF DefineNum-1<=Ubound(ClassDefineContentArr) Then
				      .Write "             <td>&nbsp;<TEXTAREA class='upfile' Name='ClassDefineContent" & DefineNum &"' ROWS='' COLS=''style='width:500px;height:100px'>" &ClassDefineContentArr(DefineNum-1)& "</TEXTAREA> " & vbCrLf
					 Else
					  .Write "             <td>&nbsp;<TEXTAREA class='upfile' Name='ClassDefineContent" & DefineNum &"' ROWS='' COLS=''style='width:500px;height:100px'></TEXTAREA> " & vbCrLf
					 End If
				  Else
				    .Write "             <td>&nbsp;<TEXTAREA class='upfile' Name='ClassDefineContent" & DefineNum &"' ROWS='' COLS=''style='width:500px;height:100px'></TEXTAREA> " & vbCrLf
				  End If
				 .Write "             </td>" & vbCrLf
				 .Write "         </tr>" & vbCrLf
				 Next 
				 .Write "       </table>" & vbCrLf 
				 .Write "</div>"

				 .Write "   </td></tr>" & vbCrLf
				 .Write " </form>" & vbCrLf
				 .Write " </table>" & vbCrLf
				 .Write "</div>"
				.Write "</body>" & vbCrLf
				.Write "</html>" & vbCrLf
				.Write "<Script Language='javascript'>" & vbCrLf
				.Write "<!--" & vbCrLf
				.Write "$(document).ready(function(){" & vbcrlf
				.Write " SelectFsoType('11')" & vbcrlf
				If Action="Edit" Then .Write "showad('" & ChannelID & "');" & vbcrlf
				.Write "})"& vbcrlf
				
				.Write "changetype('" & ClassType &"');" & vbcrlf
				.Write "function ChangeAddMode(){" & vbcrlf
				.Write " if ($('#AddMore').attr('checked')==true){"
				.Write "  $('#add1').hide(); $('#add2').show(); $('#typearea').hide();"
				.Write " }else{"
				.Write "  $('#add1').show();$('#add2').hide();$('#typearea').show();"
				.Write " }"
				.Write "}" & vbcrlf
				.Write "function SelectFsoType(ObjValue)" & vbCrLf
				.Write "{ var ChannelDomain='" & KS.GetChannelDomain(ChannelID) & KS.C_S(ChannelID,43) &"';" & vbCrLf
				 
					Dim N
					Randomize
					N = Rnd * 3 + 5
				.Write "switch (ObjValue)" & vbCrLf
				.Write "  {" & vbCrLf
				.Write "   case '1' :$('#ShowAS1').html(ChannelDomain+'<font color=red>" & YearStr & "/" & MonthStr & "-" & DayStr & "/" & KS.MakeRandom(N) & "' + $('#FnameType').val() + '</font>'); break;" & vbCrLf
				.Write "   case '2' :$('#ShowAS1').html(ChannelDomain+'<font color=red>" & YearStr & "/" & MonthStr & "/" & DayStr & "/" & KS.MakeRandom(N) & "' + $('#FnameType').val() + '</font>'); break;" & vbCrLf
				.Write "   case '3' :$('#ShowAS1').html(ChannelDomain+'<font color=red>" & YearStr & "-" & MonthStr & "-" & DayStr & "/" & KS.MakeRandom(N) & "' + $('#FnameType').val() + '</font>'); break;" & vbCrLf
				.Write "   case '4' :$('#ShowAS1').html(ChannelDomain+'<font color=red>" & YearStr & "/" & MonthStr & "/" & KS.MakeRandom(N) & "' + $('#FnameType').val() + '</font>'); break;" & vbCrLf
				.Write "   case '5' :$('#ShowAS1').html(ChannelDomain+'<font color=red>" & YearStr & "-" & MonthStr & "/" & KS.MakeRandom(N) & "' + $('#FnameType').val() + '</font>'); break;" & vbCrLf
				.Write "   case '12' :$('#ShowAS1').html(ChannelDomain+'<font color=red>" & YearStr & MonthStr & "/" & KS.MakeRandom(N) & "' + $('#FnameType').val() + '</font>'); break;" & vbCrLf
				.Write "   case '6' :$('#ShowAS1').html(ChannelDomain+'<font color=red>" & YearStr & MonthStr & DayStr & "/" & KS.MakeRandom(N) & "' + $('#FnameType').val() + '</font>'); break;" & vbCrLf
				.Write "   case '7' :$('#ShowAS1').html(ChannelDomain+'<font color=red>" & YearStr & "/" & KS.MakeRandom(N) & "' + $('#FnameType').val() + '</font>'); break;" & vbCrLf
				.Write "   case '8' :$('#ShowAS1').html(ChannelDomain+'<font color=red>" & YearStr & MonthStr & DayStr & KS.MakeRandom(N) & "' + $('#FnameType').val() + '</font>'); break;" & vbCrLf
				.Write "   case '9' :$('#ShowAS1').html(ChannelDomain+'<font color=red>" & KS.MakeRandom(N) & "'+ $('#FnameType').val() + '</font>'); break;" & vbCrLf
				.Write "   case '10' :$('#ShowAS1').html(ChannelDomain+'<font color=red>" & KS.MakeRandomChar(N) & "'+ $('#FnameType').val() + '</font>'); break;" & vbCrLf
				.Write "   case '11' :$('#ShowAS1').html(ChannelDomain+'<font color=red>�ĵ�ID'+ $('#FnameType').val() + '</font>'); break;" & vbCrLf
				.Write "  }"
				.Write "}" & vbCrLf
				.Write "function changemodel(mid){" &vbcrlf
				.Write "  showad(mid);" & vbcrlf
				.Write " for(i=0;i<marr.length;i++){" & vbcrlf
				.Write "  if (mid==marr[i][0]){$('input[name=FolderTemplateID]').val('{@TemplateDir}/'+marr[i][1]+'/��Ŀҳ.html');$('input[name=TemplateID]').val('{@TemplateDir}/'+marr[i][1]+'/����ҳ.html');$('input[name=WapFolderTemplateID]').val('{@TemplateDir}/WAPר��ģ��/'+marr[i][1]+'/WAP��Ŀҳ.html');$('input[name=WapTemplateID]').val('{@TemplateDir}/WAPר��ģ��/'+marr[i][1]+'/WAP����ҳ.html');}"
				.Write "  }" & vbcrlf
				.Write "}" & vbcrlf
			
				.Write "function CheckForm()" & vbCrLf
				.Write "{ var form=document.CreateFolderForm;" & vbCrLf
				.Write "   if ($('input[name=FolderName]').val()=='' && $('#AddMore').attr('checked')==false)" & vbCrLf
				.Write "    {" & vbCrLf
				.Write "     alert('��������Ŀ����������!');" & vbCrLf
				.Write "     $('input[name=FolderName]').focus();" & vbCrLf
				.Write "    return false;" & vbCrLf
				.Write "    }" & vbCrLf
				.Write "    if ($('input[name=FolderName]').val().length>50)" & vbCrLf
				.Write "    {" & vbCrLf
				.Write "     alert('��Ŀ�������Ʋ��ܳ���25������(50��Ӣ���ַ�)!');" & vbCrLf
				.Write "     $('input[name=FolderName]').focus();" & vbCrLf
				.Write "    return false;" & vbCrLf
				.Write "   }" & vbCrLf
				.Write "    if ($('input[name=FolderEname]').val()==''&& $('#AddMore').attr('checked')==false)" & vbCrLf
				.Write "    {" & vbCrLf
				.Write "     alert('��������Ŀ��Ӣ������!');" & vbCrLf
				.Write "     $('input[name=FolderEname]').focus();" & vbCrLf
				.Write "    return false;" & vbCrLf
				.Write "    }" & vbCrLf
				If Action<>"Edit" Then
				.Write "    if (form.classtype[0].checked && CheckEnglishStr(form.FolderEname,'��Ŀ��Ӣ������')==false)" & vbCrLf
				.Write "     return false;" & vbCrLf
				End If
				.Write "    if ($('input[name=FolderTemplateID]').val()=='')" & vbcrlf
				.Write "     { alert('�����Ŀģ��!')" & vbcrlf
				.Write "       $('input[name=FolderTemplateID]').focus();"
				.Write "       return false;}" & vbcrlf 
				.Write "    if ($('input[name=TemplateID]').val()=='')" & vbcrlf
				.Write "     { alert('�������ҳҳģ��!')" & vbcrlf
				.Write "       $('input[name=TemplateID]').focus();"
				.Write "       return false;}" & vbcrlf 
				.Write "    form.submit();" & vbCrLf
				.Write "    return true;" & vbCrLf
				.Write "}"
				.Write "function ctoe()" & vbCrLf
				.Write "{" & vbCrLf
				.Write " var folderName=escape($('input[name=FolderName]').val());" & vbcrlf
				.Write "$.get('../plus/ajaxs.asp', { foldername: folderName, action: 'Ctoe' }," &vbCrlf
				.Write "	function(data){" & vbcrlf
				.Write "	$('input[name=FolderEname]').val(unescape(data));" & vbcrlf
				.Write "  });"
				.Write "}" & vbCrLf
				.Write "setFileFileds($('select[name=ClassDefine_Num]').val());" & vbcrlf
				.Write "function setFileFileds(num){    " &vbcrlf
				.Write "for(var i=1,str="""";i<=20;i++){" & vbcrlf
				.Write "	$(""#objFiles"" + i).hide();" & vbcrlf
				.Write "}" & vbcrlf
				.Write "for(var i=1,str="""";i<=num;i++){"
				.Write "	$(""#objFiles"" + i).show();" & vbcrlf
				.Write "}" & vbcrlf
			    .Write "}" & vbcrlf
				.Write "function setchannel(v)" & vbcrlf
				.Write "{ if (v=='0') {$('#channel').show();} else {$('#channel').hide()}}"
				.Write "function changetype(v)" & vbcrlf
				.Write "{"
				.Write " switch(parseInt(v))"&vbcrlf
				.Write "  {case 1:$('#editorarea').hide();$('#fsohtmlarea').show();$('#classarea').html('Ӣ�����ƣ�');$('#classtips').html('���ܴ�\/��*���� < > | ���������,���趨���ܸ�');$('#templatearea').show();$('#temparea').show();break;" & vbcrlf
				.Write "   case 2:$('#editorarea').hide();$('#fsohtmlarea').hide();$('#classarea').html('���ӵ�ַ��');$('#classtips').html('�� <font color=blue>http://www.kesion.com</font> ��');$('#templatearea').hide();$('#temparea').hide();break;" & vbcrlf
				.Write "   case 3:$('#editorarea').show();$('#fsohtmlarea').hide();$('#classarea').html('�����ļ�����');$('#classtips').html('�� <font color=blue>about.html,intro.html,help.html</font>��');$('#templatearea').show();$('#temparea').hide();$('#channel').hide();$('#singlepage').html(""<iframe id='content___Frame' src='../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=ClassContent&amp;Toolbar=NewsTool' width='600' height='400' frameborder='0' scrolling='no'></iframe>"");break;" & vbcrlf
				.Write " } }"&vbcrlf
				.Write "function showad(v){" & vbcrlf
				.Write " if (v==1){$('#ShowAD').show();}else{$('#ShowAD').hide();}"
				.Write "}" & vbcrlf
				.Write "//-->"
				.Write "</Script>"
				
			End With
			End Sub
			
			
			'���Ƶ��Ŀ¼�ı������
			'����:ChannelID--Ƶ��ID
			Sub ChannelFolderAddSave(Go)
			Dim ID, TJ, FolderName, Folder,ChannelID, ClassID, TS, FolderTemplateID, FolderFsoIndex
			Dim TemplateID, FnameType, FsoType, FolderDomain, FolderOrder, CurrPath, TopFlag,ClassType,WapSwitch,WapFolderTemplateID,WapTemplateID
			Dim RSC,FolderEName,CommentTF,ClassPurview,GroupID,ReadPoint,ChargeType,DividePercent,PitchTime,ReadTimes,AllowArrGroupID,AddMore,ParentFolder,j,Root,Child,PrevOrderID
			Dim ClassPic,ClassDescript,MetaKeyWord,MetaDescript,ClassDefine_Num,N,ClassDefineContent,Action
				
				Action=KS.G("Action")
				AddMore=Request.Form("AddMore")
				
				If AddMore="1" Then
				 FolderName=Request.Form("FolderNames")
				 ClassType=1
				 If Trim(FolderName) = "" Then Call KS.AlertHistory("���������Ŀ,�밴��ʽ������Ŀ���Ƽ���ĿӢ������!",-1):.End
				 FolderName=Split(FolderName,vbcrlf)
				Else
				 FolderName = KS.G("FolderName")
				 ClassType  = KS.ChkClng(KS.G("ClassType"))
				 FolderEName = Replace(KS.G("FolderEName")," ","")
				 If Trim(FolderName) = "" Then Call KS.AlertHistory("Ŀ¼�������Ʋ���Ϊ��!",-1):.End
				 If KS.strLength(Trim(FolderName)) > 50 Then Call KS.AlertHistory("Ŀ¼�������Ʋ��ܳ���25������(50��Ӣ���ַ�)!", -1): .End 
				 If Trim(FolderEName) = "" Then Call KS.AlertHistory("Ŀ¼Ӣ�����Ʋ���Ϊ��!",-1):.End
				End If
				
				if ClassType=1 Then
				 If Instr(FolderEName,".") <>0 Then Call KS.AlertHistory("Ŀ¼Ӣ�����Ʋ��ܺ��С�.��!",-1):.End
				Elseif ClassType=3 Then
				 If right(lcase(FolderEName),4) <>".htm" and right(lcase(FolderEName),5)<>".html" and right(lcase(FolderEName),6)<>".shtml" and right(lcase(FolderEName),5)<>".shtm" Then Call KS.AlertHistory("��ҳ����չ������ȷ��ֻ����.html,.htm,.shtm,.shtml�е�һ��!",-1):.End
				End If
				
				ID = Trim(Request("parentID")):If ID = "" Then ID = "0"
				FolderTemplateID =KS.G("FolderTemplateID")
				TemplateID = KS.G("TemplateID")
				WapFolderTemplateID=KS.G("WapFolderTemplateID")
				WapTemplateID=KS.G("WapTemplateID")
				ChannelID=KS.ChkClng(KS.G("ChannelID"))
				
			
				If FolderTemplateID = "" Or TemplateID = "" Then Call KS.AlertHistory("�Բ���,�����Ƶ��Ӧ��ѡ��ģ���!", -1): Exit Sub
				If ClassType=3 Then
				 	If Instr(FolderEName,".")=0 Then
						Call KS.AlertHistory("��ҳ�汣����ļ���ʽ����ȷ!", -1)
						Set KS = Nothing:Response.End
					 Else
					   Dim FileExt:FileExt=lcase(Split(FolderEName,".")(1))
					   If FileExt<>"html" and FileExt<>"htm" and FileExt<>"shtml" and FileExt<>"shtm" Then
						Call KS.AlertHistory("��ҳ�汣����ļ���ʽ����ȷ,ֻ����html,htm,shtml��shtmΪ��չ��!", -1)
						Set KS = Nothing:Response.End
					   End If
					 End If
				End If
				
			   If ID <> "0" And ID<>"" Then  
				     Dim FolderRS,MaxOrderID
					 Set FolderRS = Server.CreateObject("ADODB.RECORDSET")
					 FolderRS.Open"Select Folder,FolderName,FolderDomain,TS,Tj,Root,FolderOrder,Child From KS_Class Where ID='" & ID & "'",conn,1,1
					 If FolderRS.EOF Then
					    FolderRS.Close:Set FolderRS=Nothing
						KS.AlertHintScript "����Ŀ�����ڣ�"
					 Else
					    Root=FolderRS("Root")
						PrevOrderID=FolderRS("FolderOrder")
						Child=FolderRS("Child")
						TS = Trim(FolderRS("TS"))

						if (Child > 0) Then
							'�õ��뱾��Ŀͬ�������һ����Ŀ��OrderID
							PrevOrderID = Conn.Execute("select Max(FolderOrder) From KS_Class where tn='" &ID& "'")(0)
	
							'�õ�ͬһ����Ŀ���ȱ���Ŀ�����������Ŀ�����OrderID�������ǰһ��ֵ����������ֵ
							MaxOrderID =  KS.ChkClng(Conn.Execute("select Max(FolderOrder) from [KS_Class] where ts like '" & ts & "%'")(0))
							if (MaxOrderID > PrevOrderID) Then	PrevOrderID = MaxOrderID
                        end if
						
					    ParentFolder=Trim(FolderRS("Folder"))
						Folder = ParentFolder & FolderEName
						FolderDomain = Trim(FolderRS("FolderDomain"))
						TJ = FolderRS("TJ")+1
					    
					 End If
					 FolderRS.Close:Set FolderRS = Nothing
			   Else 
					Folder = FolderEName
					TJ=1
					FolderDomain = KS.G("FolderDomain")
					Root=Conn.Execute("Select Max(root) From KS_Class")(0)
					If KS.IsNul(Root) Then 
					 Root=1
					Else
					 Root=Root+1
					End If
					
			   End If
			   
			   If ClassType=1 Then Folder=trim(Folder) & "/"
				
				If Action="Add" Then
					Set RSC=Server.CreateObject("ADODB.Recordset")
					RSC.Open "Select FolderName,Folder From KS_Class Where ChannelID=" & ChannelID & " and TN='" & ID & "'", Conn, 1, 1
					If Not RSC.EOF Then
					  If AddMore="1" Then
					      '���������Ƿ���ͬ��
						  For I=0 To Ubound(FolderName)
						   For J=0 To Ubound(FolderName)
							   If Ubound(split(FolderName(j),"|"))<1 Then
								Call KS.AlertHistory("�����������Ŀ��ʽ����ȷ!�밴����Ŀ��������|��ĿӢ�����ƣ��͸�ʽ¼��!", -1):.End
							   End If
							   If Not IsAlphabet(replace(Split(FolderName(i),"|")(1)," ","")) Then
								Call KS.AlertHistory("�����������ĿӢ�����Ʋ���ȷ!����Ӣ������!", -1):.End
							   End If
							   
						       If Split(FolderName(i),"|")(0)=Split(FolderName(j),"|")(0) and i<>j Then
							     Call KS.AlertHistory("�����������Ŀ[" & Split(FolderName(i),"|")(0) & "]�����ظ�!", -1):.End
							   End If
						       If trim(Split(FolderName(i),"|")(1))=trim(Split(FolderName(j),"|")(1)) and i<>j Then
							    Call KS.AlertHistory("���������Ӣ����Ŀ[" & Split(FolderName(i),"|")(1) & "]�����ظ�!", -1):.End
							   End If
						   Next
						  Next
						  
						  Do While Not RSC.Eof
						   For I=0 To Ubound(FolderName)
						    If RSC(0) = Split(FolderName(i),"|")(0) Then  Call KS.AlertHistory("�����������Ŀ[" & Split(FolderName(i),"|")(0) & "]�Ѵ���,������������!", -1):.End
						    If RSC(1) = Split(FolderName(i),"|")(1) Then Call KS.AlertHistory("���������Ӣ������[" & Split(FolderName(i),"|")(1) & "]�Ѵ���,��������Ӣ������!",-1): .End
						   Next
						   RSC.MoveNext
						  Loop
					  Else
						  Do While Not RSC.Eof
						   If RSC(0) = FolderName Then  Call KS.AlertHistory("�����Ѵ���,������������!", -1):.End
						   If RSC(1) = Folder Then Call KS.AlertHistory("Ӣ�������Ѵ���,��������Ӣ������!",-1): .End
						   RSC.MoveNext
						  Loop
					  End If
					End If
					RSC.Close:Set RSC=Nothing
				End If
				

			   TopFlag = KS.ChkClng(KS.G("TopFlag"))
			   WapSwitch  = KS.ChkClng(KS.G("WapSwitch"))
			   FolderFsoIndex = Request("FolderFsoIndex")
			   FnameType = Request("FnameType")
			   FsoType = Request("FsoType")
			   ClassPurview= KS.ChkClng(KS.G("ClassPurview"))
			
				CommentTF=Request.Form("CommentTF")
				GroupID=KS.G("GroupID"):if GroupID="" Then GroupID=0
				AllowArrGroupID=KS.G("AllowArrGroupID"):iF AllowArrGroupID="" Then AllowArrGroupID=0
				ClassPic=Request.Form("ClassPic")
				ClassDescript=Request.Form("ClassDescript")
				If ClassDescript="" Then ClassDescript=Request.Form("ClassContent")
				
				MetaKeyWord=Request.Form("MetaKeyWord")
				MetaDescript=Request.Form("MetaDescript")
				ClassDefine_Num=KS.ChkClng(KS.G("ClassDefine_Num"))
				For N=1 To ClassDefine_Num
				  If N=1 Then
				   ClassDefineContent=Request.Form("ClassDefineContent"& N)
				  Else
				   ClassDefineContent=ClassDefineContent & "||||" & Request.Form("ClassDefineContent"& N)
				  End If
				Next
				
				ReadPoint=KS.ChkClng(KS.G("ReadPoint"))
				ChargeType=KS.ChkClng(KS.G("ChargeType"))
				PitchTime=KS.ChkClng(KS.G("PitchTime"))
				ReadTimes=KS.ChkClng(KS.G("ReadTimes"))
				DividePercent=KS.G("DividePercent")
				If Not IsNumeric(DividePercent) Then
				 DividePercent=0
				End If
				Dim AdParam,AdPa
				AdPa="0%ks%0,0,0,0%ks%0%ks%%ks%"
				If KS.C_S(ChannelID,6)=1 Then
					AdParam=KS.G("AdParam")
					if Ubound(Split(AdParam,","))<>3 Then Call KS.AlertHistory("����Ļ��л���������������!",-1).end
					if KS.ChkClng(KS.G("ShowADTF"))=1 and KS.G("ADtype")="1" and KS.G("AdUrl")="" then Call KS.AlertHistory("����Ļ��л�����ͼƬ��ַ!",-1).end
					if KS.ChkClng(KS.G("ShowADTF"))=1 and KS.G("ADtype")="2" and KS.G("AdCode")="" then Call KS.AlertHistory("����Ļ��л����Ĵ���!",-1).end
					If KS.G("ADtype")="2" then
					AdPa=KS.ChkClng(KS.G("ShowADTF")) & "%ks%" & AdParam &"%ks%" & KS.G("ADType") & "%ks%" & Request.Form("AdCode") & "%ks%"
					else
					AdPa=KS.ChkClng(KS.G("ShowADTF")) & "%ks%" & AdParam &"%ks%" & KS.G("ADType") & "%ks%"& KS.G("AdUrl") & "%ks%" & KS.G("AdLinkUrl")
					end if
				End If
				
				Dim Node,oldnode,m
				Dim Farr:Farr=Split(ClassField,",")

				Dim RST:Set RST=Server.CreateObject("ADODB.Recordset")
				If Action="Add" Then
				     If Not IsArray(FolderName) Then FolderName=Split(FolderName,vbcrlf)
				     For I=Ubound(FolderName) To Lbound(FolderName) Step -1
						RST.Open "select * from KS_Class where 1=0", Conn, 1, 3
						RST.AddNew
						ClassID = KS.GetClassID()   '���ú���ȡ�µ�Ŀ¼ID
						RST("ID") = ClassID
						RST("Creater") = KS.C("AdminName")
						RST("AdminPurView")=KS.C("AdminName")
						RST("CreateDate") = Now
						If AddMore="1" Then
							if ID<>"" Then
							 RST("folder") = ParentFolder & trim(Split(FolderName(i),"|")(1)) & "/"
							Else
							 RST("Folder")=trim(Split(FolderName(i),"|")(1)) & "/"
							End If
						Else
						    if ClassType=2 Then
							 RST("folder") = FolderEname
							Else
							 RST("Folder")=Folder
							End If
						End If
						RST("FolderName") = Split(FolderName(i),"|")(0)
						RST("ClassType")=ClassType
						If ID <> "" Then  RST("TN") = ID Else  RST("TN") = "0"  
						RST("TJ") = TJ
						RST("TS") = "" & TS & "" & ClassID & ","
						RST("FolderTemplateID") = FolderTemplateID
						RST("TopFlag") = TopFlag
						RST("WapSwitch") = WapSwitch
						RST("FolderFsoIndex") = FolderFsoIndex
						RST("TemplateID") = TemplateID
						RST("WapFolderTemplateID")=WapFolderTemplateID
						RST("WapTemplateID")=WapTemplateID
						RST("FnameType") = FnameType
						RST("FsoType") = FsoType
						RST("FolderDomain") = FolderDomain
						RST("FolderOrder") = PrevOrderID+I
						If ID="" Or ID="0" Then
						 RST("Root")=Root+i
						Else
						 RST("Root")=Root
						End If
						RST("Child")=0
						RST("ChannelID") = ChannelID
						RST("DelTF") = 0
						RST("ClassPurview")=ClassPurview
						RST("CommentTF")=CommentTF
						RST("DefaultArrGroupID")=GroupID
						RST("AllowArrGroupID")=AllowArrGroupID
						RST("DefaultReadPoint")=ReadPoint
						RST("DefaultChargeType")=ChargeType
						RST("DefaultDividePercent")=DividePercent
						RST("DefaultPitchTime")=PitchTime
						RST("DefaultReadTimes")=ReadTimes
						RST("ClassBasicInfo")=ClassPic & "||||" & ClassDescript & "||||" & MetaKeyWord   &"||||" & MetaDescript & "||||" & AdPa
						RST("ClassDefineContent")=ClassDefineContent
						RST.Update
						
						Call KS.FileAssociation(1000,RST("ClassID"),RST("ClassBasicInfo")&ClassDefineContent,0)

						
						if (ID <>"" and id<>"0") Then
                           Conn.Execute ("update ks_class set Child=Child+1 where id='" & ID & "'")
                           '���¸���Ŀ�����Լ����ڱ���Ҫ��ͬ�ڱ������µ���Ŀ�������
						   Conn.Execute ("update ks_class set FolderOrder=FolderOrder+1 where root=" & Root & " and FolderOrder>" & PrevOrderID)
						   Conn.Execute ("update ks_Class set FolderOrder=" & PrevOrderID & "+1 where ID='" & RST("ID") & "'")

                       End If
					   
					   
					   	'�������ڴ�׷�ӽڵ�
						
						If ID="" Or ID=0 Then
						   Set Node=Application(KS.SiteSN&"_class").documentElement.appendChild(Application(KS.SiteSN&"_class").createNode(1,"class",""))
						Else
							set oldnode=Application(KS.SiteSN&"_class").DocumentElement.SelectSingleNode("class[@ks0='" & id & "']").NextSibling
							Set Node=Application(KS.SiteSN&"_class").createNode(1,"class","")
							Application(KS.SiteSN&"_class").documentElement.insertBefore Node,oldnode
					    End If
						for m=0 to Ubound(Farr)
						  Node.setAttribute "ks" & m,rst(Farr(m))
						next
							
    					RST.Close
				     Next
                         
						 KSCls.ClassAction  channelid          '��������JS

						Response.Write ("<script>if (confirm('�����ɹ�,����������?')) {location.href='KS.Class.asp?ChannelID=" & ChannelID &"&Action=" & Action &"&Go=" & Go & "&FolderID=" & ID & "';}else{location.href='KS.Class.asp?ChannelID=" & ChannelID & "';}</script>")
					Else
						RST.Open "select * from KS_Class Where ID='" &KS.G("FolderID") & "'", Conn, 1, 3
						RST("FolderName") = FolderName
						If  RST("ClassType")="2" Then
						  RST("Folder")=FolderEname
						End If
						RST("FolderTemplateID") = FolderTemplateID
						RST("TopFlag") = TopFlag
						RST("WapSwitch") = WapSwitch
						
						RST("FolderFsoIndex") = FolderFsoIndex
						RST("TemplateID") = TemplateID
						RST("WapFolderTemplateID")=WapFolderTemplateID
						RST("WapTemplateID")=WapTemplateID
						RST("FnameType") = FnameType
						RST("FsoType") = FsoType
						RST("FolderDomain") = FolderDomain
						RST("ClassPurview")=ClassPurview
						RST("CommentTF")=CommentTF
						RST("DefaultArrGroupID")=GroupID
						RST("AllowArrGroupID")=AllowArrGroupID
						RST("DefaultReadPoint")=ReadPoint
						RST("DefaultChargeType")=ChargeType
						RST("DefaultDividePercent")=DividePercent
						RST("DefaultPitchTime")=PitchTime
						RST("DefaultReadTimes")=ReadTimes
						If RST("ClassType")=3 Then ClassDescript=Request.Form("ClassContent")
						RST("ClassBasicInfo")=ClassPic & "||||" & ClassDescript & "||||" & MetaKeyWord   &"||||" & MetaDescript& "||||" & AdPa
						RST("ClassDefineContent")=ClassDefineContent
						RST.Update
					 
					    Call KS.FileAssociation(1000,RST("ClassID"),RST("ClassBasicInfo")&ClassDefineContent ,1)
					  
						  If RST("TN") = "0" Then
						   Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
						   RS.Open "Select FolderDomain,ClassPurview from KS_Class where TS Like '%" & KS.G("FolderID") & "%'", Conn, 1, 3
						   Do While Not RS.EOF
							RS("FolderDomain") = FolderDomain
							RS.Update
							RS.MoveNext
						   Loop
						   RS.Close
						  End If
						  Set RS = Nothing
						  
						
						 Dim ENode:Set ENode=Application(KS.SiteSN&"_class").DocumentElement.SelectSingleNode("class[@ks0='" & KS.G("FolderID") & "']")
						 for m=1 to Ubound(Farr)
						   If lcase(Farr(m))<>"adminpurview" Then
						    on error resume next
						    ENode.SelectSingleNode("@ks"&m).text=rst(Farr(m))
							if err then err.clear
						   End If
						 next
						 
						 KSCls.ClassAction  channelid
						 
						Response.Write ("<script>alert('��Ŀ��Ϣ�޸ĳɹ�!');location.href='KS.Class.asp';</script>")
					RST.Close
					End If
			        Set RST = Nothing
                 
			End Sub
			
			Function IsAlphabet(ByVal str )
				dim re
				set re = New RegExp 
				re.Global = True 
				re.IgnoreCase = True 
				re.Pattern="^[A-Za-z\d\s\_]+$" 
				IsAlphabet = re.Test(str) 
			End Function
			
End Class
%> 
