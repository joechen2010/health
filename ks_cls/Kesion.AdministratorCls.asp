<%
Class ManageCls
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Set KS=Nothing
		End Sub
		
		'分页SQL语句生成代码
		Function GetPageSQL(tblName,fldName,PageSize,PageIndex,OrderType,strWhere,fieldIds)
			Dim strTemp,strSQL,strOrder
			
			'根据排序方式生成相关代码
			if OrderType=0 then
				strTemp=">(select max([" & fldName & "])"
				strOrder=" order by [" & fldName & "] asc"
			else
				strTemp="<(select min([" & fldName & "])"
				strOrder=" order by [" & fldName & "] desc"
			end if
			
			'若是第1页则无须复杂的语句
			if PageIndex=1 then
			strTemp=""
			if strWhere<>"" then
			strTemp = " where " + strWhere
			end if
			strSQL = "select top " & PageSize & " " & fieldIds & " from [" & tblName & "]" & strTemp & strOrder
			else '若不是第1页，构造SQL语句
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
		
		  '返回相应模型的自定义字段名称数组
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
			'返回相应模型的自定义字段名称数组
		   Function Get_KS_D_F_Arr(ChannelID)
			  Get_KS_D_F_Arr=Get_KS_D_F_P_Arr(ChannelID,"")
		   End Function

		   '取得后台信息添加时的自定义字段表单
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
								 Get_KS_D_F_I=Get_KS_D_F_I & "<select class=""upfile"" style=""width:" & F_Arr(7,i) & """ name=""" & F_Arr(0,I) & """ onchange=""fill" & F_Arr(0,i) &"(this.value)""><option value=''>---请选择---</option>"
	
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
							  '联动菜单
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
				   if F_Arr(3,I)=9 and V_Tag<>1 Then Get_KS_D_F_I=Get_KS_D_F_I & " <input class=""button""  type='button' name='Submit' value='选择...' onClick=""OpenThenSetValue('Include/SelectPic.asp?ChannelID=" & ChannelID &"&CurrPath=" & KS.GetUpFilesDir() & "',550,290,window,$('#" & F_Arr(0,I) & "')[0]);"">"
				   If  F_Arr(2,I)<>"" Then Get_KS_D_F_I=Get_KS_D_F_I & " <span style=""margin-top:5px"">" &  F_Arr(2,I) & "</span>"
				   if F_Arr(3,I)=9 and V_Tag<>1 Then Get_KS_D_F_I=Get_KS_D_F_I & "<div><iframe id='UpPhotoFrame' name='UpPhotoFrame' src='KS.UpFileForm.asp?UPType=Field&FieldID=" & F_Arr(9,I) & "&ChannelID=" & ChannelID &"' frameborder=0 scrolling=no width='100%' height='26'></iframe></div>"
				   Get_KS_D_F_I=Get_KS_D_F_I &" </td>" &vbcrlf
				   Get_KS_D_F_I=Get_KS_D_F_I & "</tr>" &vbcrlf
				 End If
				Next
			End If
		   End Function
		   
		   '取得子联动菜单的字段值
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
		   '取得联动菜单
		   Function GetLDMenuStr(ChannelID,F_Arr,UserDefineFieldValueStr,byVal ParentFieldName,JSStr)
		     Dim OptionS,OArr,I,VArr,V,F,Str
		     Dim RSL:Set RSL=Conn.Execute("Select Top 1 FieldName,Title,Options,Width From KS_Field Where ChannelID=" & ChannelID & " and ParentFieldName='" & ParentFieldName & "'")
			 If Not RSL.Eof Then
			     Str=Str & " <select name='" & RSL(0) & "' id='" & RSL(0) & "' onchange='fill" & RSL(0) & "(this.value)' style='width:" & RSL(3) & "px'><option value=''>--请选择--</option>"
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
							   "$('#"& RSL(0)&"').append('<option value="""">--请选择--</option>');" &vbcrlf &_
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


		   '取得后台信息添加时的自定义字段
		   Function Get_KS_D_F(ChannelID,ByVal UserDefineFieldValueStr)
		      Dim F_Arr:F_Arr=Get_KS_D_F_Arr(ChannelID)
			  Get_KS_D_F=Get_KS_D_F_I(F_Arr,ChannelID,UserDefineFieldValueStr,0)
		   End Function
		   
		   '根据sql 参数取表单
		   Function Get_KS_D_F_P(ChannelID,ByVal UserDefineFieldValueStr,Param)
		      Dim F_Arr:F_Arr=Get_KS_D_F_P_Arr(ChannelID,Param)
			  Get_KS_D_F_P=Get_KS_D_F_I(F_Arr,ChannelID,UserDefineFieldValueStr,1)
		   End Function
		   
			'返回系统支持的生成类型(.htm,.html,.shtml.shtm等)参  数：ExtType 预定选中的类型
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
       '取得专题
		Sub Get_KS_Admin_Special(ChannelID,InfoID)
		   With KS
		     .echo "<script language='javascript' src='../ks_inc/kesion.box.js'></script>" & vbcrlf
		     .echo "<script language='javascript'>" & vbcrlf
			 .echo "  SelectSpecial=function(){" &vbcrlf
			 .echo "		PopupCenterIframe('选择专题','KS.Special.asp?action=Select',350,400,'auto')" & vbcrlf
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
			.echo "<select name='SpecialID' id='SpecialID' multiple style='height:100px;width:200px;'>" & OptionStr & "</select><div style='text-align:center'><font color=red>X</font> <a href='javascript:UnSelectAll()'><font color='#999999'>取消选定的专题</font></a></div></td>"
			.echo "              <td><input class='button'  type='button' name='Submit' value='选择专题...' onClick='SelectSpecial();'></td>"
			.echo "</table>"
		  End With
		End Sub
	  '从数据表添加数据到option选项 参数:表名,字段,查询条件
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
	  '取得相应的模板  参数 obj对象
	  Function Get_KS_T_C(obj)
	    Dim CurrPath:CurrPath=KS.Setting(3)&KS.Setting(90)
		If Right(CurrPath,1)="/" Then CurrPath=Left(CurrPath,Len(CurrPath)-1)
        Get_KS_T_C= "<input type='button' name=""Submit"" class=""button"" value=""选择模板..."" onClick=""OpenThenSetValue('KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle="& server.URLEncode("导入模板")&"&CurrPath=" &server.urlencode(CurrPath) & "',450,350,window," & obj & ");"">"	 
	   End Function
	   
	   '====================================================复制操作开始=================================
	    '粘贴
		Sub Paste(ChanneLID)
		 Dim DestFolderID, ContentID,Url
		  DestFolderID = KS.G("DestFolderID")
		  ContentID = KS.G("ContentID")
		  If DestFolderID = ""  Then Call KS.AlertHistory("参数传递出错!", 1):Exit Sub
		  Call PasteByCopy(ChannelID,DestFolderID, ContentID)
		  KS.Echo "<script>location.href='?ChannelID=" & ChannelID &"&ID=" & DestFolderID & "&Page=" & KS.S("Page") & "';</script>"
		End Sub
	   
	    '过程:PasteByCopy复制粘贴
		'参数:ChannelID--模型ID,NewClassID--目标目录,ContentID---被复制的文件
		Sub PasteByCopy(ChannelID,NewClassID, ContentID)
		 If ContentID <> "0" Then 
		   Dim IDS:IDS=KS.FilterIDs(ContentID)
		   Dim Flag:Flag=true '取"复制(n)"样式
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
		
		'得到复制的名称
		Function GetNewTitle(TableName,NewClassID, OriTitle)
			Dim RSC, CheckRS
			On Error Resume Next
			Set CheckRS=Conn.Execute("Select Title From " & TableName & " Where TID='" & NewClassID & "' And Title='" & OriTitle & "' And DelTF=0")
			  If Not CheckRS.EOF Then
				 Set RSC=Server.Createobject("Adodb.recordset")
				 RSC.Open "Select Title From " & TableName & " Where TID='" & NewClassID & "' And Title Like '复制%" & OriTitle & "' And DelTF=0 Order By ID Desc",conn,1,1
				 If Not RSC.EOF Then
					RSC.MoveFirst
					If RSC.RecordCount = 1 Then
					   RSC.Close:Set RSC = Nothing:CheckRS.Close:Set CheckRS = Nothing
					  GetNewTitle = "复制(1) " & OriTitle
					  Exit Function
					Else
					  GetNewTitle = "复制(" & CInt(Left(Split(RSC("Title"), "(")(1), 1)) + 1 & ") " & OriTitle
					End If
					 CheckRS.Close:RSC.Close:Set RSC = Nothing: Set CheckRS = Nothing
				 Else
				  RSC.Close:Set RSC = Nothing:CheckRS.Close:Set CheckRS = Nothing
				  GetNewTitle = "复制 " & OriTitle
				  Exit Function
				 End If
				 RSC.Close:Set RSC = Nothing
			  Else
				CheckRS.Close:Set CheckRS = Nothing
				GetNewTitle = OriTitle
				Exit Function
			  End If
		End Function
		'====================================================复制操作结束==================================================

		'====================================================回收站及彻底删除处理===========================================
		 '放入回收站
		 Sub Recely(ChannelID)
			Conn.Execute("Update [KS_ItemInfo] Set DelTF=1 where ChannelID=" & ChannelID & " and Infoid in(" & KS.FilterIDs(KS.S("ID")) & ")")
			Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set DelTF=1 where id in(" & KS.FilterIDs(KS.S("ID")) & ")")
			Response.Redirect Request.ServerVariables("HTTP_REFERER")
		 End Sub
		 '回收站还原
		 Sub RecelyBack(ChannelID)
			Conn.Execute("Update [KS_ItemInfo] Set DelTF=0 where ChannelID=" & ChannelID & " and Infoid in(" & KS.FilterIDs(KS.S("ID")) & ")")
			Conn.Execute("Update " & KS.C_S(ChannelID,2) & " Set DelTF=0 where id in(" & KS.FilterIDs(KS.S("ID")) & ")")
			Response.Redirect Request.ServerVariables("HTTP_REFERER")
		 End Sub
		 
		 '清空加收站
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
		 '删除选中模型信息操作
		Sub DelBySelect(ChannelID)
			Call DelModelInfo(ChannelID,Request("ID"))
			Response.Redirect Request.ServerVariables("HTTP_REFERER")
		End Sub
		 
		 '删除信息
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
			  
			  If ChannelID=5 Then  '商城删除订单
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
		
		'参数:ChannelID-模型id,FolderID-栏目ID,ContentPageArr-分页数组，FileName-文件名
		Sub DelInfoFile(ChannelID,FolderID,ContentPageArr,FileName)
		        on error resume next
		 		Dim CurrPath,FExt,Fname,TotalPage,I,CurrPathAndName
				CurrPath = KS.LoadFsoContentRule(ChannelID,FolderID)		 
				FExt = Mid(Trim(FileName), InStrRev(Trim(FileName), ".")) '分离出扩展名
				Fname = Replace(Trim(FileName), FExt, "")                    '分离出文件名 如 2005/9-10/1254ddd
				  		 
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
		 '======================================================回收站/删除结束=========================================
		 
		 '======================================================审核投稿开始============================================
		  '批量审核
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
		 '批量退稿
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
		   Call KS.SendInfo(RS("Inputer"),KS.Setting(0),"退稿通知",Content)
		   End If
		   RS.MoveNext
		  Loop
		  RS.Close
		  Set RS=Nothing
		  Conn.Execute("Update [KS_ItemInfo] Set Verific=3 Where Verific<>1 and channelid=" & ChannelID & " And InfoID In(" & KS.FilterIDs(KS.G("ID")) & ")")
		  Response.Redirect Request.ServerVariables("HTTP_REFERER")
		 End Sub
	 '======================================================审核投稿结束============================================
			
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
		.echo "<title>加入到专题</title>"
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
		.echo "    <td height='30' align='center'> <strong>请选择一个或多个专题</strong><br>"
		.echo "      <select name='SpecialID'  multiple style='height:340px;width:260px;'>"
		.echo KS.ReturnSpecial("")
		.echo "      </select><br><font color=blue>提示：按住""CTRL""或""Shift""键可以进行多选</font>"
		.echo "    </td>"
		.echo "  </tr>"
		.echo "  <tr align='center'>"
		.echo "   <td height='30'> <input type='button' class='button' name='button1' value='加入专题' onclick='CheckForm()'>"
		.echo "      &nbsp; <input type='button' class='button' onclick='window.close();' name='button2' value=' 取消 '>"
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
		'.echo "  { alert('对不起,您没有选择专题名称!');"
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

		  .echo ("<script>alert('操作成功!');window.close();</script>")
         
		End If
		 End With
		End Sub
		
		'收费选项
		Sub LoadChargeOption(ChannelID,ChargeType,InfoPurview,arrGroupID,ReadPoint,PitchTime,ReadTimes,DividePercent)
		  With KS
		    .echo " <div class=tab-page id=poweroption-page>"
			.echo "  <H2 class=tab>权限选项</H2>"
			.echo "	<SCRIPT type=text/javascript>"
			.echo "				 tabPane1.addTabPage( document.getElementById( ""poweroption-page"" ) );"
			.echo "	</SCRIPT>"
				
			 .echo "<TABLE style='margin:1px' width='100%' BORDER='0' cellpadding='1'  cellspacing='1' class='ctable'>"	
			 .echo "             <tr  class='tdbg'>"
			 .echo "               <td align='right' width='100'  class='clefttitle' height=30><strong>阅读权限:</strong></td>"
			 .echo "                <td height='30' nowrap> "
			 .echo "                <input name='InfoPurview' type='radio' value='0'"
			 if InfoPurview=0 Then .echo " checked"
			 .echo ">继承栏目权限（当所属栏目为认证栏目时，建议选择此项）<br>"
			 .echo "            <input name='InfoPurview' type='radio' value='1'"
			 If InfoPurview=1 Then .echo " checked"
			 .echo ">所有会员（当所属栏目为开放栏目，想单独对某些" & KS.C_S(ChannelID,3) & "进行阅读权限设置，可以选择此项）<br>"
			 .echo "            <input name='InfoPurview' type='radio' value='2'" 
			 IF InfoPurview=2 Then .echo " Checked"
			 .echo ">指定会员组（当所属栏目为开放栏目，想单独对某些" & KS.C_S(ChannelID,3) & "进行阅读权限设置，可以选择此项）<br>"
			 .echo "<table border='0' align=center width='90%'>"
			 .echo " <tr>"
			 .echo " <td>"
			 .echo KS.GetUserGroup_CheckBox("GroupID",arrGroupID,5)
			 .echo " </td>"
			 .echo "  </tr></table>"
			 .echo "                </td>"
             .echo "               </tr>"
			 .echo "             <tr  class='tdbg'>"
			 .echo "               <td align='right' width='80'  class='clefttitle' height=30><strong>阅读点数: </strong></td>"
			 .echo "                <td height='30' nowrap> &nbsp;"
			 .echo "                <input style='text-align:center' name='ReadPoint' type='text' id='ReadPoint'  value='" & ReadPoint & "' size='6' class='textbox'> 　免费阅读请设为 ""<font color=red>0</font>""，否则有权限的会员阅读此" & KS.C_S(ChannelID,3) & "时将消耗相应点数，游客将无法阅读此" & KS.C_S(ChannelID,3) & ""
			 .echo "                 </td>"
             .echo "               </tr>"
			 .echo "             <tr  class='tdbg'>"
			 .echo "               <td align='right' width='80'  class='clefttitle' height=30><strong>重复收费:</strong></td>"
			 .echo "                <td height='30' nowrap> "
			 .echo "                <input name='ChargeType' type='radio' value='0' "
			 IF ChargeType=0 Then .echo " checked"
			 .echo" >不重复收费(如果需扣点数" & KS.C_S(ChannelID,3) & "，建议使用)<br>"
			 .echo "<input name='ChargeType' type='radio' value='1'"
			 IF ChargeType=1 Then .echo " checked"
			 .echo ">距离上次收费时间 <input name='PitchTime' type='text' class='textbox' value='" & PitchTime & "' size='8' maxlength='8' style='text-align:center'> 小时后重新收费<br>            <input name='ChargeType' type='radio' value='2'"
			 IF ChargeType=2 Then .echo " checked"
			 .echo ">会员重复阅读此" & KS.C_S(ChannelID,3) & " <input name='ReadTimes' type='text' class='textbox' value='" & ReadTimes & "' size='8' maxlength='8' style='text-align:center'> 页次后重新收费<br>            <input name='ChargeType' type='radio' value='3'"
			 IF ChargeType=3 Then .echo " checked"
			 .echo ">上述两者都满足时重新收费<br>            <input name='ChargeType' type='radio' value='4'"
			 IF ChargeType=4 Then .echo " checked"
			 .echo ">上述两者任一个满足时就重新收费<br>            <input name='ChargeType' type='radio' value='5'"
			 IF ChargeType=5 Then .echo " checked"
			 .echo ">每阅读一页次就重复收费一次（建议不要使用,多页" & KS.C_S(ChannelID,3) & "将扣多次点数）"
			 .echo "                 </td>"
             .echo "               </tr>"
			 .echo "             <tr  class='tdbg' style=""display:none"">"
			 .echo "               <td align='right' width='80'  class='clefttitle' height=30><strong>分成比例: </strong></td>"
			 .echo "                <td height='30' nowrap> &nbsp;"
			 .echo "                <input name='DividePercent' type='text' id='DividePercent'  value='" & DividePercent & "' size='6' class='textbox'>% 　如果比例大于0，则将按比例把向阅读者收取的点数支付给投稿者 "
			 .echo "                 </td>"
             .echo "               </tr>"            
			 .echo "    </TABLE>"
			 .echo "  </div>"
		  End With
		End Sub
		
		'相关选项
		Sub LoadRelativeOption(ChannelID,ID)
		    %>
			<script language="javascript">
			$(document).ready(function(){
			 <!--- 相关信息---->
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
			.echo "  <H2 class=tab>相关信息</H2>"
			.echo "	<SCRIPT type=text/javascript>"
			.echo "		 tabPane1.addTabPage( document.getElementById( ""relation-page"" ) );"
			.echo "	</SCRIPT>"
			.echo "    <TABLE style='margin:1px' width='100%' BORDER='0' cellpadding='1'  cellspacing='1' class='ctable'>"	
			.echo "     <tr>"
			.echo "      <td align='center'><strong>关键字:</strong><input type='text' class='textbox' value='' id='RelativeKey' name='RelativeKey'> <strong>条件:</strong> <label><input type='checkbox' id='RelativeTypeTitle' name='RelativeTypeTitle' value='1'>标题</label> <label><input type='checkbox' name='RelativeTypeKey' id='RelativeTypeKey' value='2' checked>关键字TGA "
			.echo "      <select name='ChannelID' id='ChannelID'>"
			.echo "       <option value='0' style='color:red'>--不指定模型--</option>"
			.LoadChannelOption ChannelID
			.echo "      </select>"
			.echo "  <input class='button' type='button' value=' 查找相似信息 ' id='relativeButton'></td>"
			.echo "     </tr>"
			.echo "     <tr>"
			.echo "      <td align='center'><table border='0' width='90%'>" & vbcrlf
			.echo "           <tr><td>待选信息<br /><select id='TempInfoList' name='TempInfoList' multiple style='width:240px;height:250px'></select></td>" & vbcrlf
			.echo "          <td><input type='button' value=' 添加选中 >  ' id='RAddButton' class='button'><br /><br /><input type='button' value=' 全部添加 >> ' id='RAddMoreButton' class='button'><br /><br /><input type='button' value=' < 删除选中  ' id='RDelButton' class='button'><br /><br /><input type='button' value=' << 全部删除 ' id='RDelMoreButton' class='button'></td>"
			.echo "          <td>选中信息<br /><select id='SelectInfoList' name='SelectInfoList' multiple style='width:240px;height:250px'>"
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
	'函数名：ShowPagePara
	'作  用：显示“上一页 下一页”等信息
	'参  数：filename  ----链接地址
	'       TotalNumber ----总数量
	'       MaxPerPage  ----每页数量
	'       ShowAllPages ---是否用下拉列表显示所有页面以供跳转。
	'       strUnit     ----计数单位,CurrentPage--当前页,ParamterStr参数
	'返回值：无返回值
	'**************************************************
	Public Function ShowPage(totalnumber, MaxPerPage, FileName, ShowAllPages, strUnit, CurrentPage, ParamterStr)
		  Dim N, I, PageStr
				Const Btn_First = "[首页]" '定义第一页按钮显示样式
				Const Btn_Prev = "[上一页]" '定义前一页按钮显示样式
				Const Btn_Next = "[下一页]" '定义下一页按钮显示样式
				Const Btn_Last = "[末页]" '定义最后一页按钮显示样式
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
					.echo (" <td>页次：<font color=red>" & CurrentPage & "</font>/" & N & "页 共有:" & totalnumber & strUnit & " 每页:" & MaxPerPage & strUnit & " ")
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
						.echo ("转到:<input type='text' value='" & (CurrentPage + 1) &"' name='page' style='width:30px;height:18px;text-align:center;'>&nbsp;<input style='height:18px;border:1px #a7a7a7 solid;background:#fff;' type='submit' value='GO' name='sb'>")
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
				  SearchJS = SearchJS & "        <option value=""1"">标 题</option>" & vbCrLf
				  SearchJS = SearchJS & "        <option value=""2"">内 容</option>" & vbCrLf
				  SearchJS = SearchJS & "        <option value=""3"">作 者</option>" & vbCrLf
				  SearchJS = SearchJS & "        <option value=""4"">录入者</option>" & vbCrLf
				  SearchJS = SearchJS & "        <option value=""5"">关键字</option>" & vbCrLf
				   case 2
				  SearchJS = SearchJS & "          <option value=""1"">名 称</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""2"">简 介</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""3"">作 者</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""4"">录入者</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""5"">关键字</option>" & vbCrLf
				   case 3
				  SearchJS = SearchJS & "          <option value=""1"">名 称</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""2"">简 介</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""3"">开发商</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""4"">录入者</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""5"">关键字</option>" & vbCrLf
				   case 4
				  SearchJS = SearchJS & "          <option value=""1"">名 称</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""2"">简 介</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""3"">作 者</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""4"">录入者</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""5"">关键字</option>" & vbCrLf
				   case 5
				  SearchJS = SearchJS & "          <option value=""1"">商品名称</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""2"">生 产 商</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""3"">商品介绍</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""5"">商品Tags</option>" & vbCrLf
				   case 7
				  SearchJS = SearchJS & "          <option value=""1"">影片名称</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""2"">影片主演</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""3"">影片介绍</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""5"">影片Tags</option>" & vbCrLf
				   case 8
				  SearchJS = SearchJS & "          <option value=""1"">信息主题</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""2"">发布者</option>" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""3"">信息介绍</option>" & vbCrLf
				  end select
				  SearchJS = SearchJS & "      </select>" & vbCrLf
				  SearchJS = SearchJS & "        <select name=""ClassID"" style=""width:150"">" & vbCrLf
				  SearchJS = SearchJS & "          <option value=""0"" selected=""selected"">所有栏目</option>" & vbCrLf
				  SearchJS = SearchJS & KS.LoadClassOption(ChannelID)
				  SearchJS = SearchJS & "        </select>" & vbCrLf
				  SearchJS = SearchJS & "        <input name=""KeyWord"" type=""text"" class=""textbox""  value=""关键字"" onfocus=""this.select();""/>" & vbCrLf
				  SearchJS = SearchJS & "        <input name=""ChannelID"" value=""" & channelid & """ type=""hidden"" />" & vbCrLf
				  SearchJS = SearchJS & "        <input type=""submit"" class=""inputButton"" name=""Submit"" value=""搜 索"" /></td>" & vbCrLf
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
