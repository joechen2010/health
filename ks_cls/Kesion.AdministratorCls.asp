<%
Class ManageCls
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		End Sub
		
		'��ҳSQL������ɴ���
		Function GetPageSQL(tblName,fldName,PageSize,PageIndex,OrderType,strWhere,fieldIds)
			Dim strTemp,strSQL,strOrder
			
			'��������ʽ������ش���
			if OrderType=0 then
				strTemp=">(select max([" & fldName & "])"
				strOrder=" order by [" & fldName & "] asc"
			else
				strTemp="<(select min([" & fldName & "])"
				strOrder=" order by [" & fldName & "] desc"
			end if
			
			'���ǵ�1ҳ�����븴�ӵ����
			if PageIndex=1 then
			strTemp=""
			if strWhere<>"" then
			strTemp = " where " + strWhere
			end if
			strSQL = "select top " & PageSize & " " & fieldIds & " from [" & tblName & "]" & strTemp & strOrder
			else '�����ǵ�1ҳ������SQL���
			strSQL="select top " & PageSize & " " & fieldIds & " from [" & tblName & "] where [" & fldName & "]" & strTemp & _
			" from (select top " & (PageIndex-1)*PageSize & " [" & fldName & "] from [" & tblName & "]" 
			if strWhere<>"" then
			strSQL=strSQL & " where " & strWhere
			end if
			strSQL=strSQL & strOrder & ") as tblTemp)"
			if strWhere<>"" then
			strSQL=strSQL & " And " & strWhere
			end if
			strSQL=strSQL & strOrder
			end if
			GetPageSQL=strSQL 
		End Function
		
		  '������Ӧģ�͵��Զ����ֶ���������
		   Function Get_KS_D_F_P_Arr(ChannelID,Param)
		      Dim KS_RS_Obj:Set KS_RS_Obj=Server.CreateObject("ADODB.RECORDSET")
			   KS_RS_Obj.Open "Select FieldName,Title,Tips,FieldType,DefaultValue,Options,MustFillTF,Width,Height,FieldID,ShowOnUserForm,EditorType,ShowUnit,UnitOptions,ParentFieldName From KS_Field Where ChannelID=" & ChannelID &"  And ShowOnForm=1 " & Param & " Order By OrderID Asc",Conn,1,1
			 If Not KS_RS_Obj.Eof Then
			  Get_KS_D_F_P_Arr=KS_RS_Obj.GetRows(-1)
			 Else
			  Get_KS_D_F_P_Arr=""
			 End If
			 KS_RS_Obj.Close:Set KS_RS_Obj=Nothing
		   End Function
			'������Ӧģ�͵��Զ����ֶ���������
		   Function Get_KS_D_F_Arr(ChannelID)
			  Get_KS_D_F_Arr=Get_KS_D_F_P_Arr(ChannelID,"")
		   End Function

		   'ȡ�ú�̨��Ϣ���ʱ���Զ����ֶα�
		   Function Get_KS_D_F_I(F_Arr,ChannelID,ByVal UserDefineFieldValueStr,V_Tag)
		      Dim I,K,O_Arr,F_Value
			  Dim O_Text,O_Value,BRStr,O_Len,F_V,UnitValue,V_Arr
                If UserDefineFieldValueStr<>"0" And UserDefineFieldValueStr<>""  Then UserDefineFieldValueStr=Split(UserDefineFieldValueStr,"||||")
              If IsArray(F_Arr) Then
				For I=0 To Ubound(F_Arr,2)
				  If F_Arr(14,I)="0" Or KS.IsNul(F_Arr(14,I)) Then
				    If ChannelID=101 and F_Arr(10,I)="0" Then
				    Get_KS_D_F_I=Get_KS_D_F_I & "<tr class='tdbg'[@NoDisplay(" & F_Arr(0,i) & ")]>" & vbcrlf 
					Else
				    Get_KS_D_F_I=Get_KS_D_F_I & "<tr class='tdbg'>" & vbcrlf 
					End If
					Get_KS_D_F_I=Get_KS_D_F_I & " <td width=""85"" align=""right"" class='clefttitle'><strong>" & F_Arr(1,I) & ":</strong></td>" & vbcrlf
					Get_KS_D_F_I=Get_KS_D_F_I & " <td>"
					 If IsArray(UserDefineFieldValueStr) Then
					    F_Value=UserDefineFieldValueStr(I)
					    If F_Arr(12,I)="1" and instr(F_Value,"@")>0 Then
						V_Arr=Split(F_Value,"@")
					    F_Value=V_Arr(0)
					    UnitValue=V_Arr(1)
						End If
					 Else
					   if lcase(F_Arr(4,I))="now" then
					   F_Value=now
					   elseif lcase(F_Arr(4,I))="date" then
					   F_Value=date
					   else
					   F_Value=F_Arr(4,I)
					   end if
					   If Instr(F_Value,"|")<>0 Then 
					   	F_Value=LFCls.GetSingleFieldValue("select top 1 " & Split(F_Value,"|")(1) & " from " & Split(F_Value,"|")(0) & " where username='" & KS.C("UserName") & "'") 
					   End If
					 End If
					 
				   If V_Tag=1 Then	 
				    Get_KS_D_F_I=Get_KS_D_F_I & "[@" & F_Arr(0,i) &"]"
                   ElseIf lcase(F_Arr(0,i))="province&city" Then
				   	Get_KS_D_F_I=Get_KS_D_F_I & "<script language=""javascript"" src=""" & KS.Setting(2) & "/Plus/Area.asp""></script>"
				   Else
					   Select Case F_Arr(3,I)
						 Case 2
						   Get_KS_D_F_I=Get_KS_D_F_I & "<textarea style=""width:" & F_Arr(7,i) & ";height:" & F_Arr(8,i) & "px"" rows=""5"" class=""upfile"" name=""" & F_Arr(0,i) & """>" & F_Value & "</textarea>"
						 Case 3,11
							   If F_Arr(3,I)=11 Then
								 Get_KS_D_F_I=Get_KS_D_F_I & "<select class=""upfile"" style=""width:" & F_Arr(7,i) & """ name=""" & F_Arr(0,I) & """ onchange=""fill" & F_Arr(0,i) &"(this.value)""><option value=''>---��ѡ��---</option>"
	
							   Else
							  Get_KS_D_F_I=Get_KS_D_F_I & "<select class=""upfile"" style=""width:" & F_Arr(7,i) & """ name=""" & F_Arr(0,I) & """>"
							   End If
								   O_Arr=Split(F_Arr(5,I),vbcrlf): O_Len=Ubound(O_Arr)
								   For K=0 To O_Len
									If O_Arr(K)<>"" Then
									   F_V=Split(O_Arr(K),"|")
									   If Ubound(F_V)=1 Then
										O_Value=F_V(0):O_Text=F_V(1)
									   Else
										O_Value=F_V(0):O_Text=F_V(0)
									   End If						   
									 If F_Value=O_Value Then
									  Get_KS_D_F_I=Get_KS_D_F_I & "<option value=""" & O_Value& """ selected>" & O_Text & "</option>"
									 Else
									  Get_KS_D_F_I=Get_KS_D_F_I & "<option value=""" & O_Value& """>" & O_Text & "</option>"
									 End If
									End If
								   Next
							  Get_KS_D_F_I=Get_KS_D_F_I & "</select>"
							  '�����˵�
							  If F_Arr(3,I)=11  Then
								Dim JSStr
								Get_KS_D_F_I=Get_KS_D_F_I &  GetLDMenuStr(ChannelID,F_Arr,UserDefineFieldValueStr,F_Arr(0,i),JSStr) & "<script type=""text/javascript"">" &vbcrlf & JSStr& vbcrlf &"</script>"
							  End If
						 Case 6
						   O_Arr=Split(F_Arr(5,I),vbcrlf): O_Len=Ubound(O_Arr)
						   If O_Len>1 And Len(F_Arr(5,I))>50 Then BrStr="<br>" Else BrStr=""
						   For K=0 To O_Len
							   F_V=Split(O_Arr(K),"|")
							   If O_Arr(K)<>"" Then
							   If Ubound(F_V)=1 Then
								O_Value=F_V(0):O_Text=F_V(1)
							   Else
								O_Value=F_V(0):O_Text=F_V(0)
							   End If						   
							 If F_Value=O_Value Then
							  Get_KS_D_F_I=Get_KS_D_F_I & "<input type=""radio"" name=""" & F_Arr(0,I) & """ value=""" & O_Value& """ checked>" & O_Text & BRStr
							 Else
							  Get_KS_D_F_I=Get_KS_D_F_I & "<input type=""radio"" name=""" & F_Arr(0,I) & """ value=""" & O_Value& """>" & O_Text & BRStr
							 End If
							End If
						   Next
						 Case 7
						   O_Arr=Split(F_Arr(5,I),vbcrlf): O_Len=Ubound(O_Arr)
						   For K=0 To O_Len
						     If O_Arr(K)<>"" Then
							   F_V=Split(O_Arr(K),"|")
							   If Ubound(F_V)=1 Then
								O_Value=F_V(0):O_Text=F_V(1)
							   Else
								O_Value=F_V(0):O_Text=F_V(0)
							   End If						   
							 If KS.FoundInArr(F_Value,O_Value,",")=true Then
							  Get_KS_D_F_I=Get_KS_D_F_I & "<input type=""checkbox"" name=""" & F_Arr(0,I) & """ value=""" & O_Value& """ checked>" & O_Text
							 Else
							  Get_KS_D_F_I=Get_KS_D_F_I & "<input type=""checkbox"" name=""" & F_Arr(0,I) & """ value=""" & O_Value& """>" & O_Text
							 End If
							End If
						   Next
						 Case 10
							Get_KS_D_F_I=Get_KS_D_F_I & "<input type=""hidden"" id=""" & F_Arr(0,I) &""" name=""" & F_Arr(0,I) &""" value="""& Server.HTMLEncode(F_Value) &""" style=""display:none"" /><input type=""hidden"" id=""" & F_Arr(0,I) &"___Config"" value="""" style=""display:none"" /><iframe id=""" & F_Arr(0,I) &"___Frame"" src=""../KS_Editor/FCKeditor/editor/fckeditor.html?InstanceName=" & F_Arr(0,I) &"&amp;Toolbar=" & F_Arr(11,i) & """ width=""" & F_Arr(7,i) &""" height=""" & F_Arr(8,i) & """ frameborder=""0"" scrolling=""no""></iframe>"
	
						 Case Else
						   Get_KS_D_F_I=Get_KS_D_F_I & "<input type=""text"" class=""upfile"" style=""width:" & F_Arr(7,i) & """ name=""" & F_Arr(0,i) & """ id=""" & F_Arr(0,i) & """ value=""" & F_Value & """>"
					   End Select
				   End If
				   
				   If F_Arr(12,I)="1" Then 
					  Get_KS_D_F_I=Get_KS_D_F_I & " <select name=""" & F_Arr(0,i) & "_Unit"" id=""" & F_Arr(0,i) & "_Unit"">"
					  If Not KS.IsNul(F_Arr(13,i)) Then
				       Dim UnitOptionsArr:UnitOptionsArr=Split(F_Arr(13,i),vbcrlf)
					   For K=0 To Ubound(UnitOptionsArr)
					       if trim(UnitValue)=trim(UnitOptionsArr(k)) then
					       Get_KS_D_F_I=Get_KS_D_F_I & "<option value='" & UnitOptionsArr(k) & "' selected>" & UnitOptionsArr(k) & "</option>"                 
						   else
					       Get_KS_D_F_I=Get_KS_D_F_I & "<option value='" & UnitOptionsArr(k) & "'>" & UnitOptionsArr(k) & "</option>"                 
						   end if
					   Next
					  End If
					  Get_KS_D_F_I=Get_KS_D_F_I & "</select>"
				   End If
				   
				   If F_Arr(6,I)=1 Then Get_KS_D_F_I=Get_KS_D_F_I & "<font color=red> * </font>"
				   if F_Arr(3,I)=9 and V_Tag<>1 Then Get_KS_D_F_I=Get_KS_D_F_I & " <input class=""button""  type='button' name='Submit' value='ѡ��...' onClick=""OpenThenSetValue('Include/SelectPic.asp?ChannelID=" & ChannelID &"&CurrPath=" & KS.GetUpFilesDir() & "',550,290,window,$('#" & F_Arr(0,I) & "')[0]);"">"
				   If  F_Arr(2,I)<>"" Then Get_KS_D_F_I=Get_KS_D_F_I & " <span style=""margin-top:5px"">" &  F_Arr(2,I) & "</span>"
				   if F_Arr(3,I)=9 and V_Tag<>1 Then Get_KS_D_F_I=Get_KS_D_F_I & "<div><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='KS.UpFileForm.asp?UPType=Field&FieldID=" & F_Arr(9,I) & "&ChannelID=" & ChannelID &"' frameborder=0 scrolling=no width='100%' height='26'></iframe></div>"
				   Get_KS_D_F_I=Get_KS_D_F_I &" </td>" &vbcrlf
				   Get_KS_D_F_I=Get_KS_D_F_I & "</tr>" &vbcrlf
				 End If
				Next
			End If
		   End Function
		   
		   'ȡ���������˵����ֶ�ֵ
		   Function GetFieldValue(F_Arr,UserDefineFieldValueStr,FieldName)
		     Dim I
			 If IsArray(UserDefineFieldValueStr) Then
			      For I=0 To Ubound(F_Arr,2)
				     If Lcase(F_Arr(0,I))=Lcase(FieldName) Then
					   GetFieldValue=UserDefineFieldValueStr(I)
					   Exit Function
					 End If
				  Next
			 End If
		   End Function
		   'ȡ�������˵�
		   Function GetLDMenuStr(ChannelID,F_Arr,UserDefineFieldValueStr,byVal ParentFieldName,JSStr)
		     Dim OptionS,OArr,I,VArr,V,F,Str
		     Dim RSL:Set RSL=Conn.Execute("Select Top 1 FieldName,Title,Options,Width From KS_Field Where ChannelID=" & ChannelID & " and ParentFieldName='" & ParentFieldName & "'")
			 If Not RSL.Eof Then
			     Str=Str & " <select name='" & RSL(0) & "' id='" & RSL(0) & "' onchange='fill" & RSL(0) & "(this.value)' style='width:" & RSL(3) & "px'><option value=''>--��ѡ��--</option>"
				 JSStr=JSStr & "var sub" &ParentFieldName & " = new Array();"
				  Options=RSL(2)
				  OArr=Split(Options,Vbcrlf)
				  For I=0 To Ubound(OArr)
				    Varr=Split(OArr(i),"|")
					If Ubound(Varr)=1 Then 
					 V=Varr(0):F=Varr(1)
					Else
					 V=Varr(0)
					 F=Varr(0)
					End If
				    JSStr=JSStr & "sub" & ParentFieldName&"[" & I & "]=new Array('" & V & "','" & F & "')" &vbcrlf
				  Next
				 Str=Str & "</select>"
				 JSStr=JSStr & "function fill"& ParentFieldName&"(v){" &vbcrlf &_
							   "$('#"& RSL(0)&"').empty();" &vbcrlf &_
							   "$('#"& RSL(0)&"').append('<option value="""">--��ѡ��--</option>');" &vbcrlf &_
							   "for (i=0; i<sub" &ParentFieldName&".length; i++){" & vbcrlf &_
							   " if (v==sub" &ParentFieldName&"[i][0]){document.getElementById('" & RSL(0) & "').options[document.getElementById('" & RSL(0) & "').length] = new Option(sub" &ParentFieldName&"[i][1], sub" &ParentFieldName&"[i][1]);}}" & vbcrlf &_
							   "}"
				 Dim DefaultVAL:DefaultVAL=GetFieldValue(F_Arr,UserDefineFieldValueStr,RSL(0))
				 If Not KS.IsNul(DefaultVAL) Then
				  str=str & "<script>$(document).ready(function(){fill"&ParentFieldName&"($('select[name=" &ParentFieldName&"] option:selected').val()); $('#"& RSL(0)&"').val('" & DefaultVAL & "');})</script>" &vbcrlf
				 End If
				 GetLDMenuStr=str & GetLDMenuStr(ChannelID,F_Arr,UserDefineFieldValueStr,RSL(0),JSStr)
			 Else
			     JSStr=JSStr & "function fill" & ParentFieldName &"(v){}"				 
			 End If
			     
		   End Function


		   'ȡ�ú�̨��Ϣ���ʱ���Զ����ֶ�
		   Function Get_KS_D_F(ChannelID,ByVal UserDefineFieldValueStr)
		      Dim F_Arr:F_Arr=Get_KS_D_F_Arr(ChannelID)
			  Get_KS_D_F=Get_KS_D_F_I(F_Arr,ChannelID,UserDefineFieldValueStr,0)
		   End Function
		   
		   '����sql ����ȡ��
		   Function Get_KS_D_F_P(ChannelID,ByVal UserDefineFieldValueStr,Param)
		      Dim F_Arr:F_Arr=Get_KS_D_F_P_Arr(ChannelID,Param)
			  Get_KS_D_F_P=Get_KS_D_F_I(F_Arr,ChannelID,UserDefineFieldValueStr,1)
		   End Function
		   
			'����ϵͳ֧�ֵ���������(.htm,.html,.shtml.shtm��)��  ����ExtType Ԥ��ѡ�е�����
			Public Function GetFsoTypeStr(ExtType)
			  GetFsoTypeStr = "<select name='fnametype' id='fnametype'>"
			If ExtType = ".html" Then
			  GetFsoTypeStr = GetFsoTypeStr & "<option value='.html' selected>.html</option>"
			Else
			 GetFsoTypeStr = GetFsoTypeStr & "<option value='.html'>.html</option>"
			End If
			If ExtType = ".htm" Then
			 GetFsoTypeStr = GetFsoTypeStr & "<option value='.htm' selected>.htm</option>"
			Else
			 GetFsoTypeStr = GetFsoTypeStr & "<option value='.htm'>.htm</option>"
			End If
			If ExtType = ".shtm" Then
			 GetFsoTypeStr = GetFsoTypeStr & "<option value='.shtm' selected>.shtm</option>"
			Else
			 GetFsoTypeStr = GetFsoTypeStr & "<option value='.shtm'>.shtm</option>"
			End If
			If ExtType = ".shtml" Then
			 GetFsoTypeStr = GetFsoTypeStr & "<option value='.shtml' selected>.shtml</option>"
			Else
			 GetFsoTypeStr = GetFsoTypeStr & "<option value='.shtml'>.shtml</option>"
			End If
			If ExtType = ".asp" Then
			 GetFsoTypeStr = GetFsoTypeStr & "<option value='.asp' selected>.asp</option>"
			Else
			 GetFsoTypeStr = GetFsoTypeStr & "<option value='.asp'>.asp</option>"
			End If
			 GetFsoTypeStr = GetFsoTypeStr & "</select>"
			End Function
       'ȡ��ר��
		Sub Get_KS_Admin_Special(ChannelID,InfoID)
		   With KS
		     .echo "<script language='javascript' src='../ks_inc/kesion.box.js'></script>" & vbcrlf
		     .echo "<script language='javascript'>" & vbcrlf
			 .echo "  SelectSpecial=function(){" &vbcrlf
			 .echo "		PopupCenterIframe('ѡ��ר��','KS.Special.asp?action=Select',350,400,'auto')" & vbcrlf
			 .echo "	}" &vbcrlf
			 .echo "  SelectSpecial1=function(){" &vbcrlf
			 .echo "		var strUrl = 'KS.SpecialSelect.asp'; "& vbcrlf
			 .echo "		var isMSIE= (navigator.appName == 'Microsoft Internet Explorer');" & vbcrlf
			 .echo "		var ReturnStr = null;" &vbcrlf
			 .echo "		if (isMSIE){ReturnStr= window.showModalDialog(strUrl,self,'width=250,height=400,resizable=yes,scrollbars=yes');}" &vbcrlf
			 .echo "		else{ var win=window.open(strUrl,'newWin','left=150,width=350,height=400,resizable=yes,scrollbars=yes'); }"&vbcrlf
			 .echo "		if (ReturnStr != null){" & vbcrlf
			 .echo "			UpdateSpecial(ReturnStr);}" & vbcrlf
			 .echo "	}" &vbcrlf
			 .echo "    function UpdateSpecial(arrstr){" &vbcrlf
			 .echo "	  if (arrstr!=''){" &vbcrlf
			 .echo "	  $('#SpecialList').show();" & vbcrlf
			 
			 .echo "     var finder=false;" & vbcrlf
			 .echo "	  var arr=arrstr.split('@@@');" & vbcrlf
			 .echo "     $('#SpecialID>option').each(function(){" & vbcrlf
			 .echo "     if (arr[0]==this.value){" & vbcrlf
			 .echo "       $('#SpecialID>option[value='+arr[0]+']').attr('selected',true);finder=true;return false;}" &vbcrlf
			 .echo "  });" & vbcrlf
			 .echo "  if (finder==false){" & vbcrlf
			 .echo "	$('#SpecialID').append(""<option value=""+arr[0]+"">""+arr[1]+""</option>"");" & vbcrlf
			 .echo "	$('#SpecialID >option[value='+arr[0]+']').attr('selected',true);" & vbcrlf
			 .echo " }" & vbcrlf
			 .echo "	 }" & vbcrlf
			 .echo "	}" & vbcrlf
			 .echo " </script>" & vbcrlf
			.echo "<table border=0 width='100%'><tr>"
			Dim ShowSpecialStr:ShowSpecialStr=" style='display:none'"
			If InfoID<>0 Then
			   Dim OptionStr,RSB
			   Set RSB=Conn.Execute("Select a.SpecialID,SpecialName From KS_Special A inner join KS_SpecialR b on a.specialid=b.specialid Where ChannelID=" & ChannelID & " and InfoID=" & InfoID)
				If Not RSB.Eof Then
				  ShowSpecialStr=""
				  Do While Not RSB.Eof
				   OptionStr=OptionStr & "<option value='" & RSB(0) & "' selected>" &RSb(1) & "</option>"
				  RSB.MoveNext
				  Loop
				End If
				RSB.Close:Set RSB=Nothing
			End If
			.echo "<td width='200' id='SpecialList'" & ShowSpecialStr &">"
			.echo "<select name='SpecialID' id='SpecialID' multiple style='height:100px;width:200px;'>" & OptionStr & "</select><div style='text-align:center'><font color=red>X</font> <a href='javascript:UnSelectAll()'><font color='#999999'>ȡ��ѡ����ר��</font></a></div></td>"
			.echo "              <td><input class='button'  type='button' name='Submit' value='ѡ��ר��...' onClick='SelectSpecial();'></td>"
			.echo "</table>"
		  End With
		End Sub
	  '�����ݱ�������ݵ�optionѡ�� ����:����,�ֶ�,��ѯ����
	  Function Get_O_F_D(Table,FieldStr,Param)
	       Dim KS_RS_Obj,Arr,I
		      If Instr(lcase(FieldStr),"distinct")<=0 and Instr(lcase(FieldStr),"top")<=0 Then FieldStr=" top 50 " &FieldStr
			  Set KS_RS_Obj = conn.Execute("Select " & FieldStr & " FROM "  & Table & " Where " & Param)
			  If Not KS_RS_Obj.Eof Then
			    Arr=KS_RS_Obj.GetRows(-1)
				KS_RS_Obj.Close:Set KS_RS_Obj=Nothing
				For I=0 To Ubound(Arr,2)
					Get_O_F_D = Get_O_F_D & "<option value=""" & Arr(0,i) & """>" & Arr(0,i) & "</option>"
				Next
			   End If
	  End Function
	  'ȡ����Ӧ��ģ��  ���� obj����
	  Function Get_KS_T_C(obj)
	    Dim CurrPath:CurrPath=KS.Setting(3)&KS.Setting(90)
		If Right(CurrPath,1)="/" Then CurrPath=Left(CurrPath,Len(CurrPath)-1)
        Get_KS_T_C= "<input type='button' name=""Submit"" class=""button"" value=""ѡ��ģ��..."" onClick=""OpenThenSetValue('KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle="& server.URLEncode("����ģ��")&"&CurrPath=" &server.urlencode(CurrPath) & "',450,350,window," & obj & ");"">"	 
	   End Function
	   
	   '====================================================���Ʋ�����ʼ=================================
	    'ճ��
		Sub Paste(ChanneLID)
		 Dim DestFolderID, ContentID,Url
		  DestFolderID = KS.G("DestFolderID")
		  ContentID = KS.G("ContentID")
		  If DestFolderID = ""  Then Call KS.AlertHistory("�������ݳ���!", 1):Exit Sub
		  Call PasteByCopy(ChannelID,DestFolderID, ContentID)
		  KS.Echo "<script>location.href='?ChannelID=" & ChannelID &"&ID=" & DestFolderID & "&Page=" & KS.S("Page") & "';</script>"
		End Sub
	   
	    '����:PasteByCopy����ճ��
		'����:ChannelID--ģ��ID,NewClassID--Ŀ��Ŀ¼,ContentID---�����Ƶ��ļ�
		Sub PasteByCopy(ChannelID,NewClassID, ContentID)
		 If ContentID <> "0" Then 
		   Dim IDS:IDS=KS.FilterIDs(ContentID)
		   Dim Flag:Flag=true 'ȡ"����(n)"��ʽ
		  Dim RS, IRS, NewID,OriTitle, SqlStr,I,Intro,PhotoUrl
		  Set RS = Server.CreateObject("Adodb.RecordSet")
		  SqlStr = "Select * From " & KS.C_S(ChannelID,2) &" Where ID In(" & IDS & ") And DelTF=0"
		  RS.Open SqlStr, conn, 1, 1
		  If Not RS.EOF Then
		     Do While Not RS.Eof
				If Flag = True Then OriTitle = GetNewTitle(KS.C_S(ChannelID,2),NewClassID, RS("Title"))
				If OriTitle="" Then OriTitle = RS("Title")
			   Set IRS = Server.CreateObject("Adodb.RecordSet")
			   IRS.Open "Select top 1 * From " & KS.C_S(ChannelID,2) &" Where 1=0", conn, 1, 3
				IRS.AddNew
				For I=2 To RS.Fields.Count-1
				 IRS(I)=RS(I)
				Next
				If ChannelID=5 Then
				 IRS("ProID")=KS.GetInfoID(5)
				End If
				IRS("Title") = OriTitle
				IRS("Tid")   = NewClassID
				IRS("DelTF") = 0
				IRS.Update
				IRS.MoveLast
				NewID=IRS("ID")
				IRS("Fname")=NewID & Mid(Trim(RS("Fname")), InStrRev(Trim(RS("Fname")), "."))
				IRS.Update
				
				select case Cint(KS.C_S(ChannelID,6))
				 case 1 Intro=RS("Intro")
				 case 2 Intro=RS("PictureContent")
				 case 3 Intro=RS("DownContent")
				 case 4 Intro=RS("FlashContent")
				 case 5 Intro=RS("ProIntro")
				 case 7 Intro=RS("MovieContent")
				 case 8 Intro=RS("GQContent")
				end select
				Call LFCls.AddItemInfo(ChannelID,NewID,OriTitle,NewClassID,Intro,RS("KeyWords"),RS("PhotoUrl"),Now,KS.C("AdminName"),RS("Hits"),RS("HitsByDay"),RS("HitsByWeek"),RS("HitsByMonth"),RS("Recommend"),RS("Rolls"),RS("Strip"),RS("Popular"),RS("Slide"),RS("IsTop"),RS("Comment"),RS("Verific"),IRS("Fname"))
				IRS.Close
			  RS.MoveNext
			Loop
		  End If
		  RS.Close:Set RS = Nothing:Set IRS = Nothing
		 End If
		End Sub
		
		'�õ����Ƶ�����
		Function GetNewTitle(TableName,NewClassID, OriTitle)
			Dim RSC, CheckRS
			On Error Resume Next
			Set CheckRS=Conn.Execute("Select Title From " & TableName & " Where TID='" & NewClassID & "' And Title='" & OriTitle & "' And DelTF=0")
			  If Not CheckRS.EOF Then
				 Set RSC=Server.Createobject("Adodb.recordset")
				 RSC.Open "Select Title From " & TableName & " Where TID='" & NewClassID & "' And Title Like '����%" & OriTitle & "' And DelTF=0 Order By ID Desc",conn,1,1
				 If Not RSC.EOF Then
					RSC.MoveFirst
					If RSC.RecordCount = 1 Then
					   RSC.Close:Set RSC = Nothing:CheckRS.Close:Set CheckRS = Nothing
					  GetNewTitle = "����(1) " & OriTitle
					  Exit Function
					Else
					  GetNewTitle = "����(" & CInt(Left(Split(RSC("Title"), "(")(1), 1)) + 1 & ") " & OriTitle
					End If
					 CheckRS.Close:RSC.Close:Set RSC = Nothing: Set CheckRS = Nothing
				 Else
				  RSC.Close:Set RSC = Nothing:CheckRS.Close:Set CheckRS = Nothing
				  GetNewTitle = "���� " & OriTitle
				  Exit Function
				 End If
				 RSC.Close:Set RSC = Nothing
			  Else
				CheckRS.Close:Set CheckRS = Nothing
				GetNewTitle = OriTitle
				Exit Function
			  End If
		End Function
		'====================================================���Ʋ�������==================================================

		'====================================================����վ������ɾ������===========================================
		 '�������վ
		 Sub Recely(ChannelID)
			Conn.Execute("Update [KS_ItemInfo] Set DelTF=1 where ChannelID=" & ChannelID & " and Infoid in(" & KS.FilterIDs(KS.S("ID")) & ")")
			Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set DelTF=1 where id in(" & KS.FilterIDs(KS.S("ID")) & ")")
			Response.Redirect Request.ServerVariables("HTTP_REFERER")
		 End Sub
		 '����վ��ԭ
		 Sub RecelyBack(ChannelID)
			Conn.Execute("Update [KS_ItemInfo] Set DelTF=0 where ChannelID=" & ChannelID & " and Infoid in(" & KS.FilterIDs(KS.S("ID")) & ")")
			Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set DelTF=0 where id in(" & KS.FilterIDs(KS.S("ID")) & ")")
			Response.Redirect Request.ServerVariables("HTTP_REFERER")
		 End Sub
		 
		 '��ռ���վ
		 Sub DeleteAll()
		   If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
				Dim ModelXML,Node
				Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
				For Each Node In ModelXML.documentElement.SelectNodes("channel")
			     If Node.SelectSingleNode("@ks21").text="1" and Node.SelectSingleNode("@ks0").text<>"6" and Node.SelectSingleNode("@ks0").text<>"9" and Node.SelectSingleNode("@ks0").text<>"10" Then
				 Call DelModelInfo(Node.SelectSingleNode("@ks0").text,"Select ID From " & Node.SelectSingleNode("@ks2").text & " Where Deltf=1")
				 End If
			    Next
			Response.Redirect Request.ServerVariables("HTTP_REFERER")
		 End Sub
		 'ɾ��ѡ��ģ����Ϣ����
		Sub DelBySelect(ChannelID)
			Call DelModelInfo(ChannelID,Request("ID"))
			Response.Redirect Request.ServerVariables("HTTP_REFERER")
		End Sub
		 
		 'ɾ����Ϣ
		 Sub DelModelInfo(ChannelID,NewsID)
			  Dim K, CurrPath,FolderID,N,ImgSrcArr,RS
			  Dim ContentPageArr, TotalPage, I, CurrPathAndName, FExt, Fname
			  conn.Execute ("Delete From KS_ItemInfo Where ChannelID=" & ChannelID &" and InfoID in(" & NewsID & ")")
			  conn.Execute ("Delete From KS_ItemInfoR Where ChannelID=" & ChannelID &" and InfoID in(" & NewsID & ")")
			  conn.Execute ("Delete From KS_Comment Where ChannelID=" & ChannelID &" and InfoID in(" & NewsID & ")")
			  conn.Execute ("Delete From KS_SpecialR Where ChannelID=" & ChannelID &" and InfoID in(" & NewsID & ")")
			  conn.Execute ("Delete From KS_Digg Where ChannelID=" & ChannelID &" and InfoID in(" & NewsID & ")")
			  conn.Execute ("Delete From KS_DiggList Where ChannelID=" & ChannelID &" and InfoID in(" & NewsID & ")")
			  
			  Set RS=Server.CreateObject("ADODB.RECORDSET")
			  RS.Open "Select FileName From KS_UploadFiles Where ChannelID=" & ChannelID &" and InfoID in(" & NewsID & ")",Conn,1,1
			  Do While Not RS.Eof
			   Call KS.DeleteFile(RS(0))
			   RS.MoveNext
			  Loop
			  RS.Close
			  conn.Execute ("Delete From KS_UploadFiles Where ChannelID=" & ChannelID &" and InfoID in(" & NewsID & ")")
			  
			  If ChannelID=5 Then  '�̳�ɾ������
			     Conn.Execute("Delete From KS_OrderItem Where ProID in(" & NewsID & ")")
				 Conn.Execute("Delete From KS_ProPrice Where ProID in(" & NewsID & ")")
				 conn.execute("Delete From KS_ShopBundleSale Where ProID in(" &NewsID &")")
				 On error resume next
				 Set RS=Conn.Execute("Select SmallPicUrl,BigPicUrl From KS_ProImages Where ProID in(" & NewsID & ")")
				 Do While Not RS.Eof
				  Call KS.DeleteFile(RS(0))
				  Call KS.DeleteFile(RS(1))
				 RS.MoveNext
				 Loop
				 RS.Close:Set RS=Nothing
				 Conn.Execute("Delete From KS_ProImages Where ProID in(" & NewsID & ")")
			  End IF
			  
			  Set RS=Server.CreateObject("ADODB.Recordset")
			  RS.Open "Select * FROM " & KS.C_S(ChannelID,2) &" Where ID in(" & NewsID & ")", conn, 1, 1
			  Do While Not RS.EOF 
				 FolderID = Trim(RS("Tid"))
				 
				 If KS.C_S(ChannelID,6)=1 Then
				  ContentPageArr = Split(RS("ArticleContent"), "[NextPage]")
				 ElseIf KS.C_S(ChannelID,6)=2 Then
				  ContentPageArr = Split(RS("PicUrls"), "|||")
				 End If
				 Call DelInfoFile(ChannelID,FolderID,ContentPageArr,RS("Fname"))
			 RS.MoveNext
			Loop
			  RS.Close
			Set RS = Nothing
			conn.execute("delete  FROM " & KS.C_S(ChannelID,2) &" Where ID in(" & NewsID & ")")
		End Sub
		
		'����:ChannelID-ģ��id,FolderID-��ĿID,ContentPageArr-��ҳ���飬FileName-�ļ���
		Sub DelInfoFile(ChannelID,FolderID,ContentPageArr,FileName)
		        on error resume next
		 		Dim CurrPath,FExt,Fname,TotalPage,I,CurrPathAndName
				CurrPath = KS.LoadFsoContentRule(ChannelID,FolderID)		 
				FExt = Mid(Trim(FileName), InStrRev(Trim(FileName), ".")) '�������չ��
				Fname = Replace(Trim(FileName), FExt, "")                    '������ļ��� �� 2005/9-10/1254ddd
				  		 
	    		  If IsArray(ContentPageArr) Then TotalPage = UBound(ContentPageArr) + 1 Else TotalPage=1
				  If TotalPage > 1 and  KS.C_S(ChannelID,6)<=2 Then
					For I = LBound(ContentPageArr) To UBound(ContentPageArr)
					 If I = 0 Then
					  CurrPathAndName = CurrPath & FileName
					 Else
					  CurrPathAndName = CurrPath & Fname & "_" & (I + 1) & FExt
					 End If
					 Call KS.DeleteFile(CurrPathAndName)
					Next
				  Else
				   CurrPathAndName = CurrPath & FileName
				   Call KS.DeleteFile(CurrPathAndName)
				  End If
		End Sub
		 '======================================================����վ/ɾ������=========================================
		 
		 '======================================================���Ͷ�忪ʼ============================================
		  '�������
		 Sub VerificAll(ChannelID)
		  Dim InputerStr,Inputer,RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
		   InputerStr="Inputer"
		  RS.Open "Select " & InputerStr & ",Title,Verific,ID From " & KS.C_S(ChannelID,2) & " Where Verific<>2 And ID In(" & KS.FilterIDs(KS.G("ID")) & ")",Conn,1,3
		  Do While Not RS.Eof
			 Inputer=RS(0)
			 IF Inputer<>"" And Inputer<>KS.C("AdminName") Then Call KS.SignUserInfoOK(ChannelID,Inputer,RS(1),RS(3))
			 RS("Verific")=1
			 RS.Update
			 RS.MoveNext
		  Loop
		  RS.Close :Set RS=Nothing
		  Conn.Execute("Update [KS_ItemInfo] Set Verific=1 Where Verific<>2 and channelid=" & ChannelID & " And InfoID In(" & KS.FilterIDs(KS.G("ID")) & ")")
		  Response.Redirect Request.ServerVariables("HTTP_REFERER")
		 End Sub
		 '�����˸�
		 Sub Tuigao(ChannelID)
		  Dim RS,Content
		  Set RS=Server.CreateObject("ADODB.RECORDSET")
		  RS.Open "Select * From " & KS.C_S(ChannelID,2) & " Where Verific<>1 And ID In(" & KS.FilterIDs(KS.G("ID")) & ")",conn,1,3
		  Do While Not RS.Eof
		   RS("Verific")=3
		   RS.Update
		   If Request("Email")="1" Then
		   Content=Request("AnnounceContent")
		   Content=Replace(Content,"{$Title}",RS("Title"))
		   Content=Replace(Content,"{$UserName}",RS("Inputer"))
		   Call KS.SendInfo(RS("Inputer"),KS.Setting(0),"�˸�֪ͨ",Content)
		   End If
		   RS.MoveNext
		  Loop
		  RS.Close
		  Set RS=Nothing
		  Conn.Execute("Update [KS_ItemInfo] Set Verific=3 Where Verific<>1 and channelid=" & ChannelID & " And InfoID In(" & KS.FilterIDs(KS.G("ID")) & ")")
		  Response.Redirect Request.ServerVariables("HTTP_REFERER")
		 End Sub
	 '======================================================���Ͷ�����============================================
			
	Sub BatchSet(ChannelID)
		  Dim NID:NID=KS.FilterIDs(KS.G("ID"))
		  Select Case (KS.ChkClng(KS.S("SetAttributeBit")))
		    Case 1
				Conn.Execute("Update [KS_ItemInfo] Set Recommend=1 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
				Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set Recommend=1 where id in(" & NID & ")")
			Case 2
				Conn.Execute("Update [KS_ItemInfo] Set Slide=1 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
			    Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set Slide=1 where id in(" & NID & ")")
			Case 3
			    Conn.Execute("Update [KS_ItemInfo] Set Popular=1 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
			    Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set Popular=1 where id in(" & NID & ")")
			Case 4
			    Conn.Execute("Update [KS_ItemInfo] Set Comment=1 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
			    Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set Comment=1 where id in(" & NID & ")")
			Case 5
			    Conn.Execute("Update [KS_ItemInfo] Set strip=1 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
			    Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set strip=1 where id in(" & NID & ")")
			Case 6
			    Conn.Execute("Update [KS_ItemInfo] Set istop=1 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
			    Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set istop=1 where id in(" & NID & ")")
			Case 7
			    Conn.Execute("Update [KS_ItemInfo] Set rolls=1 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
			    Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set rolls=1 where id in(" & NID & ")")
		    Case 8
			    Conn.Execute("Update [KS_ItemInfo] Set Recommend=0 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
			    Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set Recommend=0 where id in(" & NID & ")")
			Case 9
			    Conn.Execute("Update [KS_ItemInfo] Set Slide=0 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
			    Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set Slide=0 where id in(" & NID & ")")
			Case 10
			    Conn.Execute("Update [KS_ItemInfo] Set Popular=0 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
			    Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set Popular=0 where id in(" & NID & ")")
			Case 11
			    Conn.Execute("Update [KS_ItemInfo] Set Comment=0 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
			    Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set Comment=0 where id in(" & NID & ")")
			Case 12
			    Conn.Execute("Update [KS_ItemInfo] Set strip=0 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
				Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set strip=0 where id in(" & NID & ")")
			Case 13
			    Conn.Execute("Update [KS_ItemInfo] Set istop=0 where id in(" & NID & ")")
			    Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set istop=0 where id in(" & NID & ")")
			Case 14
			    Conn.Execute("Update [KS_ItemInfo] Set rolls=0 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
			    Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set rolls=0 where id in(" & NID & ")")
			Case 15
			    Conn.Execute("Update [KS_ItemInfo] Set Verific=1 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
			    Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set Verific=1 where id in(" &NID& ")")
			Case 16
			    Conn.Execute("Update [KS_ItemInfo] Set Verific=0 where ChannelID=" & ChannelID & " And Infoid in(" & NID & ")")
			    Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set Verific=0 where id in(" &NID& ")")
		  End Select
		  Response.Redirect Request.ServerVariables("HTTP_REFERER")
		End Sub
		
		Public Sub AddToSpecial(ChannelID)
		Dim NewsID:NewsID = Trim(Request("NewsID"))
		With KS
		.echo "<html>"
		.echo "<head>"
		.echo "<meta http-equiv='Content-Type' content='text/html; chaRSet=gb2312'>"
		.echo "<title>���뵽ר��</title>"
		.echo "<link href='Include/Admin_Style.css' rel='stylesheet'>"
		.echo "<link href='Include/ModeWindow.css' rel='stylesheet'>"
		.echo "<script language='JavaScript' src='../KS_Inc/common.js'></script>"
		.echo "</head>"
		.echo "<body topmargin='0' leftmargin='0' scroll=no>"
		.echo "<table width='100%' border='0' cellspacing='0' cellpadding='0'>"
		.echo "  <form name='specialform' action='?ChannelID=" & ChannelID&"&Action=Special' method='post'>"
		.echo "  <input type='hidden' value='Add' Name='Flag'>"
		.echo "  <input type='hidden' name='SpecialName'>"
		.echo "  <input type='hidden' value='" & NewsID & "' Name='NewsID'>"
		.echo "  <tr>"
		.echo "    <td height='18'>&nbsp;</td>"
		.echo "  </tr>"
		.echo "  <tr>"
		.echo "    <td height='30' align='center'> <strong>��ѡ��һ������ר��</strong><br>"
		.echo "      <select name='SpecialID'  multiple style='height:340px;width:260px;'>"
		.echo KS.ReturnSpecial("")
		.echo "      </select><br><font color=blue>��ʾ����ס""CTRL""��""Shift""�����Խ��ж�ѡ</font>"
		.echo "    </td>"
		.echo "  </tr>"
		.echo "  <tr align='center'>"
		.echo "   <td height='30'> <input type='button' class='button' name='button1' value='����ר��' onclick='CheckForm()'>"
		.echo "      &nbsp; <input type='button' class='button' onclick='window.close();' name='button2' value=' ȡ�� '>"
		.echo "    </td>"
		.echo "  </tr>"
		.echo "  </form>"
		.echo "</table>"
		.echo "</body>"
		.echo "</html>"
		.echo "<Script>"
		.echo "function CheckForm()"
		.echo "{"
		'.echo " if (document.specialform.SpecialID.value=='0')"
		'.echo "  { alert('�Բ���,��û��ѡ��ר������!');"
		'.echo "     document.specialform.SpecialID.focus();"
		'.echo "     return false;"
		'.echo "  }"
		'.echo " document.specialform.SpecialName.value=document.specialform.SpecialID.options[document.specialform.SpecialID.selectedIndex].text;"
		.echo "  document.specialform.submit();"
		.echo "  return true"
		.echo "}"
		.echo "</Script>"
		
		If Request.Form("Flag") = "Add" Then
		   Dim SpecialID, NewsIDArr, K,I
		   SpecialID = Replace(Request.Form("SpecialID")," ","")
		   
		   NewsID=KS.FilterIDs(NewsID)
		  If NewsID<>"" Then 
		   Dim NArr:Narr=Split(NewsID,",")
		   SpecialID= Split(SpecialID,",")
		   For K=0 To Ubound(NArr)
		     Conn.Execute("Delete From KS_SpecialR Where InfoID=" & NArr(K) & " and channelid=" & ChannelID)
			 For I=0 To Ubound(SpecialID)
			 Conn.Execute("Insert Into KS_SpecialR(SpecialID,InfoID,ChannelID) values(" & SpecialID(I) & "," & NArr(K) & "," & ChannelID & ")")
			 Next
		   Next
		 End If  

		  .echo ("<script>alert('�����ɹ�!');window.close();</script>")
         
		End If
		 End With
		End Sub
		
		'�շ�ѡ��
		Sub LoadChargeOption(ChannelID,ChargeType,InfoPurview,arrGroupID,ReadPoint,PitchTime,ReadTimes,DividePercent)
		  With KS
		    .echo " <div class=tab-page id=poweroption-page>"
			.echo "  <H2 class=tab>Ȩ��ѡ��</H2>"
			.echo "	<SCRIPT type=text/javascript>"
			.echo "				 tabPane1.addTabPage( document.getElementById( ""poweroption-page"" ) );"
			.echo "	</SCRIPT>"
				
			 .echo "<TABLE style='margin:1px' width='100%' BORDER='0' cellpadding='1'  cellspacing='1' class='ctable'>"	
			 .echo "             <tr  class='tdbg'>"
			 .echo "               <td align='right' width='100'  class='clefttitle' height=30><strong>�Ķ�Ȩ��:</strong></td>"
			 .echo "                <td height='30' nowrap> "
			 .echo "                <input name='InfoPurview' type='radio' value='0'"
			 if InfoPurview=0 Then .echo " checked"
			 .echo ">�̳���ĿȨ�ޣ���������ĿΪ��֤��Ŀʱ������ѡ����<br>"
			 .echo "            <input name='InfoPurview' type='radio' value='1'"
			 If InfoPurview=1 Then .echo " checked"
			 .echo ">���л�Ա����������ĿΪ������Ŀ���뵥����ĳЩ" & KS.C_S(ChannelID,3) & "�����Ķ�Ȩ�����ã�����ѡ����<br>"
			 .echo "            <input name='InfoPurview' type='radio' value='2'" 
			 IF InfoPurview=2 Then .echo " Checked"
			 .echo ">ָ����Ա�飨��������ĿΪ������Ŀ���뵥����ĳЩ" & KS.C_S(ChannelID,3) & "�����Ķ�Ȩ�����ã�����ѡ����<br>"
			 .echo "<table border='0' align=center width='90%'>"
			 .echo " <tr>"
			 .echo " <td>"
			 .echo KS.GetUserGroup_CheckBox("GroupID",arrGroupID,5)
			 .echo " </td>"
			 .echo "  </tr></table>"
			 .echo "                </td>"
             .echo "               </tr>"
			 .echo "             <tr  class='tdbg'>"
			 .echo "               <td align='right' width='80'  class='clefttitle' height=30><strong>�Ķ�����: </strong></td>"
			 .echo "                <td height='30' nowrap> &nbsp;"
			 .echo "                <input style='text-align:center' name='ReadPoint' type='text' id='ReadPoint'  value='" & ReadPoint & "' size='6' class='textbox'> ������Ķ�����Ϊ ""<font color=red>0</font>""��������Ȩ�޵Ļ�Ա�Ķ���" & KS.C_S(ChannelID,3) & "ʱ��������Ӧ�������οͽ��޷��Ķ���" & KS.C_S(ChannelID,3) & ""
			 .echo "                 </td>"
             .echo "               </tr>"
			 .echo "             <tr  class='tdbg'>"
			 .echo "               <td align='right' width='80'  class='clefttitle' height=30><strong>�ظ��շ�:</strong></td>"
			 .echo "                <td height='30' nowrap> "
			 .echo "                <input name='ChargeType' type='radio' value='0' "
			 IF ChargeType=0 Then .echo " checked"
			 .echo" >���ظ��շ�(�����۵���" & KS.C_S(ChannelID,3) & "������ʹ��)<br>"
			 .echo "<input name='ChargeType' type='radio' value='1'"
			 IF ChargeType=1 Then .echo " checked"
			 .echo ">�����ϴ��շ�ʱ�� <input name='PitchTime' type='text' class='textbox' value='" & PitchTime & "' size='8' maxlength='8' style='text-align:center'> Сʱ�������շ�<br>            <input name='ChargeType' type='radio' value='2'"
			 IF ChargeType=2 Then .echo " checked"
			 .echo ">��Ա�ظ��Ķ���" & KS.C_S(ChannelID,3) & " <input name='ReadTimes' type='text' class='textbox' value='" & ReadTimes & "' size='8' maxlength='8' style='text-align:center'> ҳ�κ������շ�<br>            <input name='ChargeType' type='radio' value='3'"
			 IF ChargeType=3 Then .echo " checked"
			 .echo ">�������߶�����ʱ�����շ�<br>            <input name='ChargeType' type='radio' value='4'"
			 IF ChargeType=4 Then .echo " checked"
			 .echo ">����������һ������ʱ�������շ�<br>            <input name='ChargeType' type='radio' value='5'"
			 IF ChargeType=5 Then .echo " checked"
			 .echo ">ÿ�Ķ�һҳ�ξ��ظ��շ�һ�Σ����鲻Ҫʹ��,��ҳ" & KS.C_S(ChannelID,3) & "���۶�ε�����"
			 .echo "                 </td>"
             .echo "               </tr>"
			 .echo "             <tr  class='tdbg' style=""display:none"">"
			 .echo "               <td align='right' width='80'  class='clefttitle' height=30><strong>�ֳɱ���: </strong></td>"
			 .echo "                <td height='30' nowrap> &nbsp;"
			 .echo "                <input name='DividePercent' type='text' id='DividePercent'  value='" & DividePercent & "' size='6' class='textbox'>% �������������0���򽫰����������Ķ�����ȡ�ĵ���֧����Ͷ���� "
			 .echo "                 </td>"
             .echo "               </tr>"            
			 .echo "    </TABLE>"
			 .echo "  </div>"
		  End With
		End Sub
		
		'���ѡ��
		Sub LoadRelativeOption(ChannelID,ID)
		    %>
			<script language="javascript">
			$(document).ready(function(){
			 <!--- �����Ϣ---->
			  $('#relativeButton').click(function(){
			   GetRealtiveItem();
			  });
	          $('#RAddButton').click(function(){
			   var alloptions = $("#TempInfoList option");
			   var so = $("#TempInfoList option:selected");
			   var a = (so.get(so.length-1).index == alloptions.length-1)? so.prev().attr("selected",true):so.next().attr("selected",true);
                
				if (!$("#SelectInfoList option[value="+so.val()+"]").attr("selected")){
				 $("#SelectInfoList").append(so);
				 }else{
				 so.remove();}
			  });
			  
			  $('#RAddMoreButton').click(function(){
			     $("#TempInfoList option").each(function(){
				  if ($("#SelectInfoList option[value="+$(this).val()+"]").attr("selected")){ $(this).remove() }
				 });
			    $("#SelectInfoList").append($("#TempInfoList option").attr("selected",true));
			  });
			  $('#RDelButton').click(function(){
			     var alloptions = $("#SelectInfoList option");
				 var so = $("#SelectInfoList option:selected");
				 var a = (so.get(so.length-1).index == alloptions.length-1)? so.prev().attr("selected",true):so.next().attr("selected",true);
			   
				$("#TempInfoList").append(so);
			  });
			  $('#RDelMoreButton').click(function(){
			    $("#TempInfoList").append($("#SelectInfoList option"));
			  });
			  
			  });
			
			GetRealtiveItem=function(){
			 $(parent.frames["FrameTop"].document).find("#ajaxmsg").toggle("fast");
			 var key=escape($('input[name=RelativeKey]').val());
			 var Rtitle=$('#RelativeTypeTitle').attr("checked");
			 var Rkey=$('#RelativeTypeKey').attr("checked");
			 var ChannelID=$('#ChannelID').val();
			 $.get("../plus/ajaxs.asp", { action: "GetRelativeItem", channelid:ChannelID,key: key,rtitle:"'"+Rtitle+"'",rkey:"'"+Rkey+"'",id:"<%=KS.G("ID")%>"},
			 function(data){
					$(parent.frames["FrameTop"].document).find("#ajaxmsg").toggle("fast");
					$("#TempInfoList").empty();
					$("#TempInfoList").append(data);
			  });
			}
			 </script>
			<%
			With KS
		    .echo " <div class=tab-page id=relation-page>"
			.echo "  <H2 class=tab>�����Ϣ</H2>"
			.echo "	<SCRIPT type=text/javascript>"
			.echo "		 tabPane1.addTabPage( document.getElementById( ""relation-page"" ) );"
			.echo "	</SCRIPT>"
			.echo "    <TABLE style='margin:1px' width='100%' BORDER='0' cellpadding='1'  cellspacing='1' class='ctable'>"	
			.echo "     <tr>"
			.echo "      <td align='center'><strong>�ؼ���:</strong><input type='text' class='textbox' value='' id='RelativeKey' name='RelativeKey'> <strong>����:</strong> <label><input type='checkbox' id='RelativeTypeTitle' name='RelativeTypeTitle' value='1'>����</label> <label><input type='checkbox' name='RelativeTypeKey' id='RelativeTypeKey' value='2' checked>�ؼ���TGA "
			.echo "      <select name='ChannelID' id='ChannelID'>"
			.echo "       <option value='0' style='color:red'>--��ָ��ģ��--</option>"
			.LoadChannelOption ChannelID
			.echo "      </select>"
			.echo "  <input class='button' type='button' value=' ����������Ϣ ' id='relativeButton'></td>"
			.echo "     </tr>"
			.echo "     <tr>"
			.echo "      <td align='center'><table border='0' width='90%'>" & vbcrlf
			.echo "           <tr><td>��ѡ��Ϣ<br /><select id='TempInfoList' name='TempInfoList' multiple style='width:240px;height:250px'></select></td>" & vbcrlf
			.echo "          <td><input type='button' value=' ���ѡ�� >  ' id='RAddButton' class='button'><br /><br /><input type='button' value=' ȫ����� >> ' id='RAddMoreButton' class='button'><br /><br /><input type='button' value=' < ɾ��ѡ��  ' id='RDelButton' class='button'><br /><br /><input type='button' value=' << ȫ��ɾ�� ' id='RDelMoreButton' class='button'></td>"
			.echo "          <td>ѡ����Ϣ<br /><select id='SelectInfoList' name='SelectInfoList' multiple style='width:240px;height:250px'>"
			If ID<>0 Then
				 Dim RArray,I,RSR,SQLStr
				 SQLStr="Select TOP 200 I.ChannelID,I.InfoID,I.Title From KS_ItemInfo I Inner Join KS_ItemInfoR R On I.InfoID=R.RelativeID Where R.ChannelID=" & ChannelID &"  and R.InfoID=" & ID &" and R.RelativeChannelID=I.ChannelID"
				 
				 response.write sqlstr
				 
				 Set RSR=Conn.Execute(SQLStr)
				 If Not RSR.Eof Then
				  RArray=RSR.GetRows(-1)
				 End If
				 RSR.Close
				 Set RSR=Nothing
				 If IsArray(RArray) Then
				   For i=0 To Ubound(RArray,2)
					.echo "<option value='" & RArray(0,I) & "|" & RArray(1,i) & "' selected>" & RArray(2,i) & "</option>"
				   Next
				 End If
            End If
			.echo "</select></td></tr>"
			.echo "     </tr>"
			
			.echo "    </TABLE>"
			.echo "  </div>"
		 End With
		End Sub

		
	Sub AddKeyTags(KeyWords)
		     dim i
			 dim trs:set trs=server.createobject("adodb.recordset")
			 dim karr:karr=split(KeyWords,",")
			 for i=0 to ubound(karr)
			 trs.open "select * from ks_keywords where keytext='" & left(karr(i),100) & "'",conn,1,3
			 if trs.eof then
			   trs.addnew
			   trs("keytext")=left(karr(i),100)
			   trs("adddate")=now
			  trs.update
		   end if
			  trs.close
		  next
		   set trs=nothing
	End Sub
		
	'**************************************************
	'��������ShowPagePara
	'��  �ã���ʾ����һҳ ��һҳ������Ϣ
	'��  ����filename  ----���ӵ�ַ
	'       TotalNumber ----������
	'       MaxPerPage  ----ÿҳ����
	'       ShowAllPages ---�Ƿ��������б���ʾ����ҳ���Թ���ת��
	'       strUnit     ----������λ,CurrentPage--��ǰҳ,ParamterStr����
	'����ֵ���޷���ֵ
	'**************************************************
	Public Function ShowPage(totalnumber, MaxPerPage, FileName, ShowAllPages, strUnit, CurrentPage, ParamterStr)
		  Dim N, I, PageStr
				Const Btn_First = "[��ҳ]" '�����һҳ��ť��ʾ��ʽ
				Const Btn_Prev = "[��һҳ]" '����ǰһҳ��ť��ʾ��ʽ
				Const Btn_Next = "[��һҳ]" '������һҳ��ť��ʾ��ʽ
				Const Btn_Last = "[ĩҳ]" '�������һҳ��ť��ʾ��ʽ
					If totalnumber Mod MaxPerPage = 0 Then
						N = totalnumber \ MaxPerPage
					Else
						N = totalnumber \ MaxPerPage + 1
					End If
                    With KS
					.echo ("<table border='0'>")
					.echo ("<form action=""" & FileName & "?" & ParamterStr & """ name=""goform"" method=""post"">")
					If ParamterStr<>"" Then ParamterStr="&" & ParamterStr
					.echo ("<tr>")
					.echo (" <td>ҳ�Σ�<font color=red>" & CurrentPage & "</font>/" & N & "ҳ ����:" & totalnumber & strUnit & " ÿҳ:" & MaxPerPage & strUnit & " ")
					If CurrentPage < 2 Then
						.echo (Btn_First & " " & Btn_Prev & " ")
					Else
						.echo ("<a href=" & FileName & "?page=1" & ParamterStr & ">" & Btn_First & "</a> <a href=" & FileName & "?page=" & CurrentPage - 1 & ParamterStr & ">" & Btn_Prev & "</a> ")
					End If
					
					If N - CurrentPage < 1 Then
						.echo (" " & Btn_Next & " " & Btn_Last & " ")
					Else
						.echo (" <a href=" & FileName & "?page=" & (CurrentPage + 1) & ParamterStr & ">" & Btn_Next & "</a> <a href=" & FileName & "?page=" & N & ParamterStr & ">" & Btn_Last & "</a> ")
					End If
					If ShowAllPages = True Then
						.echo ("ת��:<input type='text' value='" & (CurrentPage + 1) &"' name='page' style='width:30px;height:18px;text-align:center;'>&nbsp;<input style='height:18px;border:1px #a7a7a7 solid;background:#fff;' type='submit' value='GO' name='sb'>")
				  End If
				  .echo ("</td></tr>")
	              .echo ("</form></table>")
				 End With
	End Function
		
		Sub ClassAction(ChannelID)
				'KS.Echo "<iframe src=""KS.ClassMenu.asp?action=Create"" frameborder=""0"" width=""0"" height=""0""></iframe>"
'exit sub
				
				 Dim KSR:Set KSR=New Refresh
                 Call KS.CreateListFolder(KS.Setting(3) & KS.Setting(93))
				 Dim SearchJS,FsoPath
				  FsoPath=KS.Setting(3) & KS.Setting(93) & "S_" & KS.C_S(ChannelID,10) & ".js"
				  SearchJS = "<table width=""98%"" border=""0"" align=""center"">" & vbCrLf
				  SearchJS = SearchJS & "<form id=""SearchForm"" name=""SearchForm"" method=""get"" action=""" & KS.Setting(3) & "plus/Search.asp"">" & vbCrLf
				  SearchJS = SearchJS & "  <tr>" & vbCrLf
				  SearchJS = SearchJS & "    <td align=""center""><select name=""SearchType"">" & vbCrLf
				  
				  select case ks.c_s(channelid,6)
				   case 1
				  SearchJS = SearchJS & "        <option value=""1"">�� ��</option>" & vbCrLf
				  SearchJS = SearchJS & "        <option value=""2"">�� ��</option>" & vbCrLf
				  SearchJS = SearchJS & "        <option value=""3"">�� ��</option>" & vbCrLf
				  SearchJS = SearchJS & "        <option value=""4"">¼����</option>" & vbCrLf
				  SearchJS = SearchJS & "        <option value=""5"">�ؼ���</option>" & vbCrLf
				   case 2
				  SearchJS = SearchJS & "          <option value=""1"">�� ��</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""2"">�� ��</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""3"">�� ��</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""4"">¼����</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""5"">�ؼ���</option>" & vbCrLf
				   case 3
				  SearchJS = SearchJS & "          <option value=""1"">�� ��</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""2"">�� ��</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""3"">������</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""4"">¼����</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""5"">�ؼ���</option>" & vbCrLf
				   case 4
				  SearchJS = SearchJS & "          <option value=""1"">�� ��</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""2"">�� ��</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""3"">�� ��</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""4"">¼����</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""5"">�ؼ���</option>" & vbCrLf
				   case 5
				  SearchJS = SearchJS & "          <option value=""1"">��Ʒ����</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""2"">�� �� ��</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""3"">��Ʒ����</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""5"">��ƷTags</option>" & vbCrLf
				   case 7
				  SearchJS = SearchJS & "          <option value=""1"">ӰƬ����</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""2"">ӰƬ����</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""3"">ӰƬ����</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""5"">ӰƬTags</option>" & vbCrLf
				   case 8
				  SearchJS = SearchJS & "          <option value=""1"">��Ϣ����</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""2"">������</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""3"">��Ϣ����</option>" & vbCrLf
				  end select
				  SearchJS = SearchJS & "      </select>" & vbCrLf
				  SearchJS = SearchJS & "        <select name=""ClassID"" style=""width:150"">" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""0"" selected=""selected"">������Ŀ</option>" & vbCrLf
				  SearchJS = SearchJS & KS.LoadClassOption(ChannelID)
				  SearchJS = SearchJS & "        </select>" & vbCrLf
				  SearchJS = SearchJS & "        <input name=""KeyWord"" type=""text"" class=""textbox""  value=""�ؼ���"" onfocus=""this.select();""/>" & vbCrLf
				  SearchJS = SearchJS & "        <input name=""ChannelID"" value=""" & channelid & """ type=""hidden"" />" & vbCrLf
				  SearchJS = SearchJS & "        <input type=""submit"" class=""inputButton"" name=""Submit"" value=""�� ��"" /></td>" & vbCrLf
				  SearchJS = SearchJS & "  </tr>" & vbCrLf
				  SearchJS = SearchJS & "</form>" & vbCrLf
				  SearchJS = SearchJS & "</table>"
				  
				  SearchJS = Replace(Replace(SearchJS,"'","\'"),"""","\""")
				  SearchJS = KSR.ReplaceJsBr(SearchJS)
				  
				  Call KSR.FsoSaveFile(SearchJS,FsoPath)
                  Set KSR=Nothing
			End Sub
End Class
%> 
