<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%
'response.cachecontrol="no-cache"
'response.addHeader "pragma","no-cache"
'response.expires=-1
'response.expiresAbsolute=now-1
'Response.CharSet="gb2312"
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New DIYFunction
KSCls.Kesion()
Set KSCls = Nothing

Class DIYFunction
        Private KS
		Private ActionStr,LabelID, LabelRS, SQLStr, LabelName, Descript, LabelContent, LabelFlag, ParentID,Action, Page, RSCheck, FolderID,FieldParam,SQLType,ItemName,pagenum,dbname1,LabelIntro,PageStyle,note,tconn
		Private datasourcetype,datasourcestr,ajax
		Private KeyWord, SearchType, StartDate, EndDate
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Call KS.DelCahe(KS.SiteSn & "_sqllabellist")
		 Set KS=Nothing
		End Sub
		Public Sub Kesion()
		KeyWord = Request("KeyWord")
		SearchType = Request("SearchType")
		StartDate = Request("StartDate")
		EndDate = Request("EndDate")
		Action = Request.QueryString("Action")
		Page = Request("Page")
		Dbname1=KS.G("dbname1")
		FolderID = Request.QueryString("FolderID")
		LabelName=Request("LabelName")
		ItemName=Request("ItemName")
		SQLType=Request("SQLType")
		LabelID = Request("LabelId")
		PageStyle=Request("PageStyle")
		Note=Request("Note")
		
		datasourcetype=KS.ChkClng(Request("datasourcetype"))
		datasourcestr=Request("datasourcestr")
		
		IF KS.G("action")="testsource" Then
		  call testsource():exit sub
		ElseIf KS.G("action")="testlabelname" then
		  call testlabelname():exit sub
		end if
		
        Call OpenExtConn()
		With KS
		.echo "<html>"
		.echo "<head>"
		.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		.echo "<title>�½���ǩ</title>"
		.echo "</head>"
		.echo "<link href=""Admin_Style.CSS"" rel=""stylesheet"">"
		.echo "<script language=""JavaScript"" src=""../../ks_inc/Common.js""></script>"
		.echo "<script language=""JavaScript"" src=""../../ks_inc/jquery.js""></script>"
		%>
		<script>
		  function ChangeSqlType(num)
		  {
		   if (num==1) 
		   {
			 $("#pagearea").show()
			}else
		   {
			$("#pagearea").hide()
		   }
		  }
	  </script>
		<%
		.echo "<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
			Select Case KS.G("Action")
			 Case "ShowClassID" ShowClassID
			 Case "AddNewSubmit"
				 Call AddLabelSave()
			 Case "EditSubmit"
				 Call EditLabelSave()
			 Case "Step2" 
			     Call Step2()
			 Case "Step1"
			     Call Step1()
			 Case "Edit"
			     Call Step0
			 Case Else
			   Call Step0()
			End Select
		.echo "</body>"
		.echo "</html>"
		End With
	  End Sub
	  
	  sub testlabelname()
	        Dim LabelID:LabelID=request.QueryString("labelid")
			Dim RS:Set RS = Server.CreateObject("Adodb.RecordSet")
			if labelid<>"" then 
			 RS.Open "Select LabelName From [KS_Label] Where id<>'" & labelid & "' and LabelName='" & "{SQL_" & LabelName & "}" & "'", Conn, 1, 1
			else
			 RS.Open "Select LabelName From [KS_Label] Where LabelName='" & "{SQL_" & LabelName & "}" & "'", Conn, 1, 1
			end if
			If Not RS.EOF Then
			 KS.Echo "false"
			Else
			 KS.Echo "true"
			end if
			rs.close:set rs=nothing
	  end sub
	  
	  Sub testsource()
	  on error resume next
	   dim str:str=request("str")
	   If KS.G("DataType")="1" or KS.G("DataType")="5" or KS.G("DataType")="6"  Then str=LFCls.GetAbsolutePath(str)
	   dim tconn:Set tconn = Server.CreateObject("ADODB.Connection")
		tconn.open str
		If Err Then 
		  Err.Clear
		  Set tconn = Nothing
		  KS.Echo "false"
		else
		  KS.Echo "true"
		end if
	  end sub
	  Sub Step0()
	    With KS
		 .echo "<body>"
	 	 .echo " <table width='100%' height='25' border='0' cellpadding='0' cellspacing='1' bgcolor='#efefef' class='sort'>"
		 .echo "       <tr><td><div align='center'><font color='#990000'>"
		 .echo "��һ��:ΪSQL��ǩ��������Դ"
		 .echo "    </font></div></td></tr>"
		 .echo "    </table>"
		 
		If LabelID <> "" Then
		    Dim FieldParamArr
		    ActionStr="Step2"
			Set LabelRS = Server.CreateObject("Adodb.Recordset")
			SQLStr = "SELECT top 1 * FROM [KS_Label] Where ID='" & LabelID & "'"
			LabelRS.Open SQLStr, Conn, 1, 1
			If Not LabelRS.Eof Then
				LabelName = Replace(Replace(LabelRS("LabelName"), "{SQL_", ""), "}", "")
				FolderID=LabelRS("FolderID")
				LabelContent = Server.HTMLEncode(LabelRS("LabelContent"))
				FieldParamArr= Split(LabelRS("Description"),"@@@")
			End IF
			LabelIntro =FieldParamArr(0)
			If Ubound(FieldParamArr)>=1 Then
			FieldParam =FieldParamArr(1)
			SQLType= FieldParamArr(2)
			ItemName=FieldParamArr(3)
			PageStyle=FieldParamArr(4)
			Ajax=FieldParamArr(5)
			datasourcetype=FieldParamArr(6)
			datasourcestr=FieldParamArr(7)
			Note=FieldParamArr(8)
			if datasourcetype<>0 then Call OpenExtConn()

			End If
			LabelRS.Close
		Else
		  ItemName="ƪ"
		  PageStyle=1
		  ActionStr="Step1"
		  Ajax=1
		  datasourcestr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=���ݿ�.mdb"
		End If
			  %>
	    <script>
		  function CheckForm()
		  {
		   if ($('#LabelName').val()=='')
		   {
		      alert('�������ǩ����!');
			  $('#LabelName').focus();
			  return;
		   }
		   if ($('#lbtf').val()=='false')
		   {
		      alert('��ǩ���Ʋ����ã�������!');
			  $('#LabelName').focus();
			  return;
		   }
		   $('#myform').submit();
		  }
		  function changeconnstr()
		  {
		    if ($('#datasourcetype').val()==0)
			{
			  $('#datasourcestr').attr("disabled",true);
			  $('#testbutton').attr("disabled",true);
			  $('#lbt').show();
			 }
			else
			{
			  $('#testbutton').attr("disabled",false);
			  $('#datasourcestr').attr("disabled",false);
			//  $('lbt').style.display='none';
			}
		    switch (parseInt($('#datasourcetype').val()))
		    {
			 case 1:
			  $('#datasourcestr').val('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=���ݿ�.mdb');
			  break;
			 case 2:
			  $('#datasourcestr').val('Provider=Sqloledb; User ID=�û���; Password=����; Initial Catalog=���ݿ�����; Data Source =(local);');
			  break;
			 case 3:
		      $('#datasourcestr').val('DSN=����Դ��;UID=�û���;PWD=����');
			  break;
			 case 4:
		      $('#datasourcestr').val('driver={microsoft odbc for oracle};uid=�û���;pwd=����;server=������');
			  break;
			 case 5:
		      $('#datasourcestr').val('driver={microsoft excel driver (*.xls)};dbq=���ݿ�����');
			  break;
			 case 6:
		      $('#datasourcestr').val('driver={microsoft dbase driver (*.dbf)};dbq=���ݿ�����');
			  break;
			 case 7:
			  alert('����mysql����Դ,��Ҫ������֧��mysql odbc 3.51 driver����Դ');
		      $('#datasourcestr').val('driver={mysql odbc 3.51 driver};server=����������;database=���ݿ�����;user name=�û���;password=����;');
			  break;
			}
		
		  }
		  function testlabelname()
		  {
		  var LabelName = $('#LabelName').val();
		  var url = 'labelfunctionadd.asp';
  		  $.get(url,{action:"testlabelname",labelid:"<%=labelid%>",labelname:LabelName},function(d){
		    if (d=='true')
			  $('#labelmessage').html('<font color=blue>��ϲ������ʹ�ø�����!</font>');
			else
			  $('#labelmessage').html('<font color=red>�Բ��𣬸����Ʋ����ã��Ѵ���!</font>');
			  $('#lbtf').val(d);
		  });
		  }
		  
		  function testsource()
		  {
			  var str = $('#datasourcestr').val();
			  var datatype=$('#datasourcetype').val();
			  if (str=='')
			  {
				alert('�����������ַ���!');
				$('#datasourcestr').focus();
				return false;
			  }
			  var url = 'labelfunctionadd.asp';
			  $.get(url,{action:"testsource",datatype:datatype,str:str},function(d){
				if (d=='true')
				 alert('��ϲ������ͨ��!')
				else
				 alert('�Բ����ַ�����������!');
			  });
		 } 
           
		</script>
		<br><br>
	     <table border='0' cellspacing='1' cellpadding='1' width='95%' align='center' class='ctable'>
		  <form action="?action=<%=ActionStr%>" method="post" id="myform" name="myform">
		   <input name='lbtf' id='lbtf' type='hidden'>
		   <input type='hidden' name='labelid' value='<%=labelid%>'>
		  <tr class='tdbg'>
		    <td class='clefttitle' align='right'><strong>��ǩ����:</strong></td>
		    <td><input name="LabelName" id="LabelName" value="<%=LabelName%>" onblur='testlabelname()' style="width:200;"> <font color=red>*</font><span id='labelmessage'></span><br>�����ǩ���ƣ�&quot;�Ƽ������б�&quot;������ģ���е��ã�<font color="#FF0000">&quot;{SQL_�Ƽ������б�(����1,����2...)}&quot;</font>��</td>
		  </tr>
		  <tr class='tdbg'>
		   <td width="80" height="30" class='clefttitle' align='right'><strong>��ǩĿ¼:</strong></td>
		   <td><%=KS.ReturnLabelFolderTree(FolderID, 5)%><font color=""#FF0000"">��ѡ���ǩ����Ŀ¼���Ա��պ�����ǩ</font></td>
		  </tr>
		  <tr class='tdbg'>
		   <td width="80" height="30" class='clefttitle' align='right'><strong>�� �� Դ:</strong></td>
		   <td>
		     <select name="datasourcetype" id="datasourcetype" style="width:290px" onChange="changeconnstr()">
			   <option value="0"<%if datasourcetype=0 then .echo " selected"%>>KesionCMS�����ݿ�</option>
			   <option value="1"<%if datasourcetype=1 then .echo  " selected"%>>Access����Դ</option>
			   <option value="2"<%if datasourcetype=2 then .echo  " selected"%>>MS SQL����Դ</option>
			   <option value="3"<%if datasourcetype=3 then .echo  " selected"%>>ODBC����Դ</option>
			   <option value="4"<%if datasourcetype=4 then .echo  " selected"%>>Oracle����Դ</option>
			   <option value="5"<%if datasourcetype=5 then .echo  " selected"%>>Excel����Դ</option>
			   <option value="6"<%if datasourcetype=6 then .echo  " selected"%>>Dbase����Դ</option>
			   <option value="7"<%if datasourcetype=7 then .echo  " selected"%>>MYSQL����Դ(��֧��mysql odbc 3.51 driver)</option>
			 </select>
		   </td>
		  </tr>
		  <tr class='tdbg'>
		   <td width="80" height="25" class='clefttitle' align='right'><strong>�����ַ���:</strong></td>
		   <td><textarea <%if datasourcetype=0 then .echo  " disabled"%> name="datasourcestr" id="datasourcestr" cols="70" rows="3"><%=Datasourcestr%></textarea>
		     &nbsp;<input class='button' id="testbutton" name="testbutton" <%if datasourcetype=0 then .echo " disabled"%> type='button' value='����' onclick='testsource();'>
			 <br><font color=green>˵��:�ⲿAccess����Դ֧�����·��,��Provider=Microsoft.Jet.OLEDB.4.0;Data Source=/���ݿ�.mdb</font>
		   </td>
		  </tr>
		  
		  <%If LabelID <> "" Then%>
		  <tr class='tdbg'>
		   <td width="80" height="25" class='clefttitle' align='right'><strong>��ѯ���:</strong></td>
		   <td><textarea name="LabelIntro" cols="80" style="width:98%" rows="4"><%=LabelIntro%></textarea>
		   </td>
		  </tr>
		  <%End If%>
		  
		  <tr class='tdbg'>
		   <td width="80" height="30" class='clefttitle' align='right'><strong>Ajax����:</strong></td>
		   <td><input type='radio' value='1' name='ajax'<%if ajax=1 then .echo  " checked"%>>��&nbsp;<input type='radio' value='0' name='ajax'<%if ajax=0 then .echo  " checked"%>>��
		   </td>
		  </tr>
		  <tr class='tdbg' id='lbt'>
		   <td width="80" height="45" class='clefttitle' align='right'><strong>��ǩ����:</strong></td>
		   <td>

		    <input type="radio" name="SQLType" value="0" <%if sqltype=0 then .echo  " checked"%> onclick='ChangeSqlType(this.value);'>��ͨ��ǩ  
			<input type="radio" name="SQLType" value="1"<%if sqltype=1 then .echo  " checked"%> onclick='ChangeSqlType(this.value);'>�ռ���ҳ��ǩ<font color=red>(���ⲿ���ݿ�����ã�һ��ҳ��ֻ�ܷ�һ����ҳ��ǩ)</font>
			
			<table border='0' id='pagearea' <%if sqltype=0 then .echo  " style=display:none"%>>
			 <tr><td>��ҳ��Ŀ��λ��<input type="text" value="<%=itemname%>" class="textbox" name="ItemName" size="6"> �磺ƪ���顢��������</td><td width='250'>&nbsp;&nbsp;&nbsp;<%=KS.ReturnPageStyle(PageStyle)%></td>
			 </tr>
			 </table>
			</td>

		  </tr>
		  
		  <tr class='tdbg'>
		   <td width="80" height="25" class='clefttitle' align='right'><strong>��Ҫ˵��:</strong></td>
		   <td><textarea name="note" cols="80" style="width:98%" rows="9"><%=note%></textarea>
		   </td>
		  </tr>
		  
		  </form>
		 </table>
	  <%
	   End With
	  End Sub
	  
	  '�ڶ���
	  Sub Step1()
	  %>
	  <script language="javascript">
	  function checkfield(){
		var strtmpp='' ;
		strtmpp= "<table border='1' cellpadding='2' cellspacing='1'  width='98%' class='border'><tr align='center'>";
		<%if datasourcetype=0 then%>
		strtmpp = strtmpp + "<td title='ͨ�ñ�ǩ'><font color=red>ͨ�ñ�ǩ=></font></td>";
		strtmpp = strtmpp +" <td title='��ǰģ��ID' style='cursor:pointer;' onclick=AddParamToSql2('{$CurrChannelID}')>{$CurrChannelID}</td>";
		strtmpp = strtmpp + "<td style='cursor:pointer;'  onclick=AddParamToSql2('{$CurrClassID}') title='��ǰ���¡�ͼƬ�����ء������ȵ�ͨ����ĿID�����������Թ����ͨ�õ��Զ��庯����ǩ.�� Select id,title From KS_Article Where Tid=��{$CurrClassID}��'>{$CurrClassID}</td>";
		strtmpp = strtmpp + "<td style='cursor:pointer;'  onclick=AddParamToSql2('{$CurrClassChildID}') title=\"��������Ŀ��ͨ����ĿID,�ԡ������Ÿ���,�� Select ID,Title,AddDate From KS_Article Where Tid in({$CurrClassChildID})\">{$CurrClassChildID}</td>";
		strtmpp = strtmpp + "<td style='cursor:pointer' onclick=AddParamToSql2('{$CurrInfoID}') title='��ǰ��Ϣ�����£�ͼƬ�����صȣ���ID,��Select ID,Intro From KS_Article Where ID={$CurrInfoID}'>{$CurrInfoID}</td>";
		strtmpp = strtmpp + "<td style='cursor:pointer' onclick='AddParamToSql2('{$CurrSpecialID}')' title=\"��ǰר��ID��ֻ����ר��ҳʹ�ã�,��Select ID,Intro From KS_Article Where specialid like '%{$CurrInfoID}%'\">{$CurrSpecialID}</td>";
		strtmpp = strtmpp + "</tr>";
		<%End if%>
		
		strtmpp = strtmpp + "<tr align='center'>";
		var fieldtemp = document.myform.FieldParam.value.split("\n");
			for(i=0;i<fieldtemp.length;i++){
				strtmpp = strtmpp + "<td style='cursor:pointer;' onclick='AddParamToSql(" + i + ")'>" + fieldtemp[i] + "</td>";
				if(((i+1)%5) == 0){
					strtmpp = strtmpp + "</tr><tr align='center'>";
				}
			}
			strtmpp = strtmpp + "</table>";
			document.getElementById ("ParamList").innerHTML=strtmpp;
     }
	 var pos=null;
	 function setPos()
	 { if (document.all){
			document.myform.LabelIntro.focus();
		    pos = document.selection.createRange();
		  }else{
		    pos = document.getElementById("LabelIntro").selectionStart;
		  }
	 }
	 //����
	function InsertValue(Val)
	{  if (pos==null) {alert('���ȶ�λҪ�����λ��!');return false;}
		if (document.all){
			  pos.text=Val;
		}else{
			   var obj=$("#LabelIntro");
			   var lstr=obj.val().substring(0,pos);
			   var rstr=obj.val().substring(pos);
			   obj.val(lstr+Val+rstr);
		}
	 }
	 function AddParamToSql(input){
		if (input != null){
			InsertValue("{$Param(" + input + ")}");
		}
	}
	function AddParamToSql2(input){
		if (input != null){
		   if (document.all)
		   {
		      myform.LabelIntro.focus();
		      var str = document.selection.createRange();
              str.text = input;
		   }else{
			InsertValue(input);
		   }
		}
	}
	
	  function CheckForm()
		{ var form=document.myform; 
		  if (form.LabelName.value=='')   
		  { alert('�������ǩ����!');
			  form.LabelName.focus();
			  return false; 
			} 
			  form.submit(); 
			  $(parent.document).find('#BottomFrame')[0].src='KS.Split.asp?OpStr=��ǩ���� >> <font color=red>�Զ���SQL��ǩ</font>&ButtonSymbol=LabelAdd';
			  return true;
		}
		 function changedb()
		 {
		  var dbname1=$('#dbname1').val();
		  var dbname2=$('#dbname2').val();
		  var LabelName=$('#LabelName').val();
		  var ParentID=$('#ParentID').val();
		  var Page=$('#Page').val();
		  var LabelID=$('#LabelID').val();
		  var PageStyle=$('#PageStyle').val();
		  var Ajax=$('#Ajax').val();
		  var SQLType=$('#SQLType').val();
		  var ItemName=$('#ItemName').val();
		  var datasourcetype=$('#datasourcetype').val();
		  var datasourcestr=$('#datasourcestr').val();
		  var Note=$('#Note').val();
		  location.href='LabelFunctionAdd.asp?action=Step1&Flag=addfield&Ajax='+Ajax+'&Note='+Note+'&datasourcetype='+datasourcetype+'&datasourcestr='+datasourcestr+'&dbname1='+dbname1+'&dbname2='+dbname2+'&LabelName='+LabelName+'&ParentID='+ParentID+'&Page='+Page+'&LabelID='+LabelID+'&SQLType='+SQLType+'&ItemName='+ItemName+'&PageStyle='+PageStyle;
		 }
		function addfield(){
			document.myform.LabelIntro.value='';
			var select=document.myform.field;
			var select2=document.myform.field2;
			for(i=0;i<select.length;i++){
				if(document.myform.field[i].selected==true){
					if(document.myform.dbname2.value==''){
						if (document.myform.LabelIntro.value==''){
							document.myform.LabelIntro.value=document.myform.field[i].value;
						}else{
							document.myform.LabelIntro.value=document.myform.LabelIntro.value+","+document.myform.field[i].value;
						}
					}else{
						if (document.myform.LabelIntro.value==''){
							document.myform.LabelIntro.value=document.myform.dbname1.value + "." + document.myform.field[i].value;
						}else{
							document.myform.LabelIntro.value=document.myform.LabelIntro.value + "," + document.myform.dbname1.value + "." + document.myform.field[i].value;
						}
					}
				}
			}
			if(document.myform.dbname2.value==''){
				if(document.myform.pagenum.value>0){
				<% if datasourcetype=5 then%>
					document.myform.LabelIntro.value="select " + document.myform.LabelIntro.value + " from [<%=dbname1%>]";
				}else{
					document.myform.LabelIntro.value="select top 10 " + document.myform.LabelIntro.value + " from [<%=dbname1%>]";
				}
				<%else%>
					document.myform.LabelIntro.value="select " + document.myform.LabelIntro.value + " from <%=dbname1%>";
				}else{
					document.myform.LabelIntro.value="select top 10 " + document.myform.LabelIntro.value + " from <%=dbname1%>";
				}
				<%end if%>
			}else{
				for(i=0;i<select2.length;i++){
					if(document.myform.field2[i].selected==true){
						if (document.myform.LabelIntro.value==''){
							document.myform.LabelIntro.value=document.myform.dbname2.value + "." + document.myform.field2[i].value;
						}else{
							document.myform.LabelIntro.value=document.myform.LabelIntro.value + "," + document.myform.dbname2.value + "." + document.myform.field2[i].value;
						}
					}
				}
				if(document.myform.dbname1.value==''){
					if(document.myform.pagenum.value>0){
						document.myform.LabelIntro.value="select " + document.myform.LabelIntro.value + " from <%=KS.G("dbname2")%>";
					}else{
						document.myform.LabelIntro.value="select top 10 " + document.myform.LabelIntro.value + " from <%=KS.G("dbname2")%>";
					}
				}else{
					if(document.myform.bg1.value==''){
						if(document.myform.pagenum.value>0){
							document.myform.LabelIntro.value="select " + document.myform.LabelIntro.value + " from <%=dbname1%>,<%=KS.G("dbname2")%>";
						}else{
							document.myform.LabelIntro.value="select top 10 " + document.myform.LabelIntro.value + " from <%=dbname1%>,<%=KS.G("dbname2")%>";
						}
					}else{
						if(document.myform.pagenum.value>0){
							document.myform.LabelIntro.value="select " + document.myform.LabelIntro.value + " from <%=dbname1%>,<%=KS.G("dbname2")%> where ";
						}else{
							document.myform.LabelIntro.value="select top 10 " + document.myform.LabelIntro.value + " from <%=dbname1%>,<%=KS.G("dbname2")%> where ";
						}
						document.myform.LabelIntro.value=document.myform.LabelIntro.value + "<%=dbname1%>." + document.myform.bg1.value + " = " + "<%=KS.G("dbname2")%>." + document.myform.bg2.value;
					}
				}
			}
		}
		</script>
		<script src="../../ks_inc/kesion.box.js"></script>
		<script>
		function ShowIframe()
		{   PopupImgDir="../";
			PopupCenterIframe("�鿴��Ŀ<=>ID���ձ�","?action=ShowClassID",600,350,"auto")
		}
		</script>
	  <%
	  FolderID=KS.G("ParentID")
	  If SQLType="" Then SQLType=0:pagenum=0 else pagenum=1
	  IF ItemName="" Then ItemName="ƪ"
	   With KS
	    .echo "<table width=""100%""  border=""0"" cellpadding=""0"" cellspacing=""0"">"
		.echo "  <form name=""myform"" id=""myform"" method=post action=""?action=Step2"" onSubmit=""return(CheckForm())"">"
		.echo "    <input type=""hidden"" name=""LabelFlag"" id=""LabelFlag"" value=""5"">"
		.echo "    <input type=""hidden"" name=""LabelID"" id=""LabelID"" value=""" & LabelID & """>"
		.echo "    <input type=""hidden"" name=""Page"" id=""Page"" value=""" & Page & """>"
		.echo "    <input type='hidden' name='pagenum' id=""pagenum"" value='" & pagenum &"' id='pagenum'>"	
		
		.echo " <input type='hidden' value=""" & LabelName & """ id=""LabelName"" name=""LabelName"" style=""width:200;"">	"
		.echo " <input type=""hidden"" name=""ParentID"" id=""ParentID"" value=""" & FolderID & """>"
		.echo " <input type=""hidden"" name=""SQLType"" id=""SQLType"" value=""" & SQLType & """>"
		.echo " <input type=""hidden"" name=""ItemName"" id=""ItemName"" value=""" & ItemName & """ size=""6""> "
		.echo " <input type=""hidden"" name=""PageStyle"" id=""PageStyle"" value=""" & PageStyle & """ size=""6""> "
		.echo " <input type=""hidden"" name=""Note"" id=""Note"" value=""" & note & """ size=""6""> "
		
		.echo " <input type=""hidden"" name=""Ajax"" id=""Ajax"" value=""" & KS.G("Ajax") & """>"
		.echo " <input type=""hidden"" name=""datasourcetype"" id=""datasourcetype"" value=""" & datasourcetype & """ size=""6""> "
		.echo " <input type=""hidden"" name=""datasourcestr"" id=""datasourcestr"" value=""" & datasourcestr & """> "
		
		
		.echo " <tr>"
		.echo "   <td height=""25"" colspan=""2""> "
		 .echo "      <table width='100%' height='25' border='0' cellpadding='0' cellspacing='1' bgcolor='#efefef' class='sort'>"
		 .echo "       <tr><td><div align='center'><font color='#990000'>"
		  If Action = "EditLabel" Then
		   .echo "�޸��Զ��庯����ǩ"
		   Else
		   .echo "�ڶ���:����SQL��ѯ���"
		  End If
		.echo "    </font></div></td><td><a href='javascript:ShowIframe()'><u>�鿴��Ŀ<=>ID���ձ�</u></a></td></tr>"
		.echo "    </table>"
		.echo " </td>"
		.echo "    </tr>"
		.echo "    <tr><td colspan=""2"" align=""center"" height=""25"" class=""title""><strong>�� �� SQL �� ѯ �� ��</strong></td></tr>"
		.echo "    <tr>"
		.echo "      <td height=""30"" colspan=2>"
		%>
		<table style="margin-top:5px" width="100%" border=0 cellpadding='2' cellspacing='1' class='border'>
			<tr class="tdbg">
			  <td width=100 height="28" align=middle><strong>����</strong></td>
			  <td>
			  <select name='dbname1' id='dbname1' onChange='changedb()' class="textbox" style="WIDTH: 250px;" >
			  <option value=''>��ѡ��һ�����ݱ�</option>
			  <%showmain(1)%>
			  </select>     </td>
			  <td align=center width=100><strong>�ӱ�</strong></td>
			  <td>
			  <select name='dbname2' id='dbname2' class="textbox" onChange='changedb()' style="WIDTH: 250px;" >
			  <option value=''>��ѡ��һ�����ݱ�</option>
			  <%showmain(2)%>
			  </select>     </td>
			</tr>
			<tr class="tdbg" <%If dbname1<>"" and KS.G("dbname2")<>"" then .echo "" else .echo " style='display:none'"%>>
			  <td height="28" align=middle><strong>Լ���ֶΣ�</strong></td>
			  <td><select name='bg1' class="textbox" style='width:250px;'>
			  <Option value=''>ѡ�������ֶ�</Option>
               <%
				if KS.G("flag")="addfield" then
				ShowChird(dbname1)
				end if
			 %>			  </select>			  </td>
			  <td align=center><strong>&lt;&lt; ���� &gt;&gt;</strong></td>
			  <td><select class="textbox" name='bg2' style='width:250px;'><option value=''>ѡ��ӱ��ֶ�</option>
		  <%
			if KS.G("flag")="addfield" then
			ShowChird(KS.G("dbname2"))
			end if
			%>
			  </select>			  </td>
		  </tr>
			<tr class="tdbg">
			  <td align=middle width=100><strong>ѡ���ֶΣ�</strong><br><br><font color=#ff0000>��ѡ����Ҫ���õ��ֶ�����,��Ctrl��Shift����ѡ</font></td>
			  <td width=100>
			<Select class="textbox" style="WIDTH: 250px; HEIGHT: 210px" onChange='addfield()' multiple size=1 name="field">
			<%
			if dbname1="" then .echo "<Option value=0>����ѡ��һ����</Option>"
			if KS.G("flag")="addfield" then
			ShowChird(dbname1)
			end if
			%>
			  </Select></td>
			  <td align=center><strong>&gt;&gt;&gt;</strong></td>
			  <td>
		<Select class="textbox" style="WIDTH: 250px; HEIGHT: 210px" onChange='addfield()' multiple size=2 name="field2">
		  <%
		  if KS.G("dbname2")="" then .echo "<Option value=0>����ѡ��һ����</Option>"
			if KS.G("flag")="addfield" then
			ShowChird(KS.G("dbname2"))
			end if
			%>
			  </Select></td>
		</tr>
		    <tr class="tdbg">
		      <td align='center'><strong>����˵����</strong></td>
		      <td colspan=2 valign="middle"><textarea class="textbox" name='FieldParam' cols='55' rows='3' id='FieldParam' onKeyUp="checkfield();" style="height:60px"></textarea></td>
	          <td valign="middle"><font color='#FF0000'>*(���ɸ�) ���뺯���б����,ÿ��һ��,�������������ա�</font></td>
	      </tr>
	      <tr class="tdbg">
            <td width='100' align='center'><strong>��ѯ��䣺</strong></td>
            <td colspan=3><div id="ParamList">
			<%if datasourcetype=0 then%>
			<table border='1' cellpadding='2' cellspacing='1'  width='98%' class='border'>
			<tr align='center'><td title='ͨ�ñ�ǩ'><font color=red>ͨ�ñ�ǩ=></font></td><td style='cursor:pointer;' onClick="AddParamToSql2('{$CurrChannelID}')" title='��ǰģ��ID'>{$CurrChannelID}</td><td style='cursor:pointer;' onClick="AddParamToSql2('{$CurrClassID}')" title='��ǰ���¡�ͼƬ�����ء������ȵ�ͨ����ĿID�����������Թ����ͨ�õ��Զ��庯����ǩ.�� Select id,title From KS_Article Where Tid=��{$CurrClassID}��'>{$CurrClassID}</td><td style='cursor:pointer;'  onclick="AddParamToSql2('{$CurrClassChildID}')" title="��������Ŀ��ͨ����ĿID,�ԡ������Ÿ���,�� Select ID,Title,AddDate From KS_Article Where Tid in({$CurrClassChildID})">{$CurrClassChildID}</td><td style="cursor:pointer" onClick="AddParamToSql2('{$CurrInfoID}')" title="��ǰ��Ϣ�����£�ͼƬ�����صȣ���ID,��Select ID,Intro From KS_Article Where ID={$CurrInfoID}">{$CurrInfoID}</td><td style="cursor:pointer" onClick="AddParamToSql2('{$CurrSpecialID}')" title="��ǰר��ID��ֻ����ר��ҳʹ�ã�,��Select ID,Intro From KS_Article Where specialid like '%{$CurrInfoID}%'">{$CurrSpecialID}</td></tr>
			</table>
			<%end if%>
			</div><textarea name='LabelIntro' onClick="setPos()" class="textbox" cols='97' rows='5' style='width:98%;height:80px' id='LabelIntro'>select top 10 * from KS_Article</textarea>
			<br>
			<font color=red>�ر���ʾ��</font>
			 <li>1.֧��ʹ��{ReqNum(�ַ���)}��{ReqStr(�ַ���)}��ȡ��Url�Ĳ���ֵ</font><br><font color=blue>�磺http://www.kesion.com/index.asp?ClassID=100,��ôsql��䣺select top 10 foldername from ks_class where classid={ReqNum(ClassID)} ���Զ�ת��Ϊselect top 10 foldername from ks_class where classid=100</font>
			 </li>
			 <li>2.ѭ����֧��ʹ��IF�������,�꿴<a href="http://www.kesion.com/tech/v5/73.html" target="_blank">http://www.kesion.com/tech/v5/73.html</a></li>
			 <li>3.������/��ҵ�ռ�Ҫʹ��sql��ǩʱ,������<font color=red>"{$GetUserName}"</font>ȡ�õ�ǰ�ռ���û���
			 <br><font color=blue>��:select top 10 id,title from ks_article where inputer='{$GetUserName}' order by id desc</font></li>
			</td>
	   </tr>
	   
	   </table>
		<%
		.echo "      </td></tr>"
		.echo "</form></table>"
	  End With
	 End Sub
	 '**************************************************
	'��������ShowMain
	'��  �ã���ʾ���ݱ��б�
	'��  ������
	'**************************************************
	Sub ShowMain(Num)
		dim rs,tablename,temptable,modeltablestr
		With KS
		if datasourcetype=0 then
			Dim rsc:set rsc=conn.execute("select itemname,channeltable,channelname from ks_channel where channelid<>6 And ChannelID<>9 and channelstatus=1 order by channelid")
			if not rsc.eof then
				 .echo "<optgroup  style=""color:blue;"" label=""=============ģ�����ݱ�============="">"
				 do while not rsc.eof
				   modeltablestr=modeltablestr & rsc(1) & ","
				   if KS.G("dbname"&num)= rsc(1) then
					.echo "<option value='" & rsc(1) & "' selected>" & rsc(0) & "���ݱ�(" & rsc(2) &"|" & rsc(1) & ")</option>"
				   else
					.echo "<option value='" & rsc(1) & "'>" & rsc(0) & "���ݱ�(" & rsc(2) &"|" & rsc(1) & ")</option>"
				   end if
					rsc.movenext
				 loop
				   if KS.G("dbname"&num)= "KS_Class" then
					.echo "<option value='KS_Class' selected style=""color:red"">ģ����Ŀ��</option>"
				   else
					.echo "<option value='KS_Class' style=""color:red"">ģ����Ŀ��</option>"
				   end if
				 modeltablestr=modeltablestr &"ks_class,"
				 .echo "<optgroup  label=""=============������============="">"
			end if
			rsc.close:set rsc=nothing
		 end if
		 
		 if datasourcetype=0 then
		 Set rs = Conn.OpenSchema(4)
		 else
		 Set rs = tConn.OpenSchema(4)
		 end if
		tablename=""
		Do While Not rs.EOF
			'temptable=Lcase(rs("Table_name"))
			temptable=rs("Table_name")
			if temptable <> tablename and temptable <> "KS_Admin" and temptable <> "KS_NotDown" and lcase(left(temptable,4)) <> "msys" and lcase(left(temptable,3)) <> "sys" and KS.FoundInArr(modeltablestr, temptable, ",")=false then
			'if (temptable ="KS_Article" or temptable = "KS_Photo" or temptable = "KS_DownLoad" or temptable = "KS_Flash" or temptable = "KS_Movie" or temptable = "KS_GQ" or temptable = "KS_Product" or temptable = "KS_Class") and temptable <> tablename then
			    if KS.G("dbname"&num)= temptable then
				.echo "<option value='" & temptable & "' selected>" & temptable & "</option>"
				else
				.echo "<option value='" & temptable & "'>" &temptable & "</option>"
				end if
				Tablename = temptable
			end if
		rs.MoveNext
		Loop
		rs.close:set rs=nothing
	 End With
	End Sub
	 '**************************************************
	'��������ShowChird
	'��  �ã���ʾָ�����ݱ���ֶ��б�
	'��  ������
	'**************************************************
	Sub ShowChird(dbname)
		dim rs
		if dbname<>"" then	
		   if datasourcetype<>0 then
		    Set rs=Tconn.OpenSchema(4)
		   else
			Set rs = Conn.OpenSchema(4)	
		   end if
		   
			Do Until rs.EOF or rs("Table_name") = trim(dbname)
				rs.MoveNext
			Loop
			Dim UserFieldArr,CommonFieldArr,CommonField
			
			if datasourcetype=0 then
				Dim rsc:set rsc=server.createobject("adodb.recordset")
				rsc.open "select channelname,itemname,BasicType from ks_channel where channelid<>6 And ChannelID<>9 and channeltable='" & dbname & "'",conn,1,1
				if not rsc.eof then
					CommonField=GetCommonField(rsc(0),rsc(2),CommonFieldArr,rsc(1))
					KS.echo CommonField
					Do Until rs.EOF or rs("Table_name") <> trim(dbname)
					   if left(lcase(rs("column_Name")),3)="ks_" then
						 if UserFieldArr="" then
						 UserFieldArr=UserFieldArr & rs("column_Name")
						 else
						 UserFieldArr=UserFieldArr & "," & rs("column_Name")
						 end if
					   elseif KS.FoundInArr(CommonFieldArr, lcase(rs("column_Name")), ",")=false and lcase(rs("column_Name"))<>"orderid" Then
						KS.echo "<option value='"&rs("column_Name")&"'>��"&GetFieldName(rs("column_Name"),rsc(1))&"</option>"
					   end if
						rs.MoveNext
					loop
					KS.echo GetUserField(UserFieldArr)
				else
				   
					if lcase(dbname)="ks_class" then KS.echo  GetCommonField("��Ŀ",0,CommonFieldArr,"��Ŀ")
					Do Until rs.EOF or rs("Table_name") <> trim(dbname)
					  if lcase(dbname)<>"ks_class" then
					   KS.echo "<option value='"&rs("column_Name")&"'>��"&rs("column_Name")&"</option>"
					  else
						KS.echo "<option value='"&rs("column_Name")&"'>��"&GetFieldName(rs("column_Name"),"")&"</option>"
					  end if
					   rs.MoveNext
					loop
				end if
				rsc.close:set rsc=nothing
			else 
					Do Until rs.EOF or rs("Table_name") <> trim(dbname)
					   KS.echo "<option value='"&rs("column_Name")&"'>��"&rs("column_Name")&"</option>"
					   rs.MoveNext
					loop
			End If
			rs.close:set rs=nothing
		End if
	End Sub
	
	'�Զ����ֶ�
	Function GetUserField(UserFieldArr) 
	  Dim i
	  GetUserField="<optgroup  style=""color:red"" label=""=====�û��Զ����ֶ�====="">"
	  UserFieldArr=Split(UserFieldArr,",")
	  For I=0 TO Ubound(UserFieldArr)
	   GetUserField= GetUserField&"<option value=""" &UserFieldArr(i) &""">��" &UserFieldArr(i) &"</option>"
	  Next
	End Function
	
	'�����ֶ��б�
	Function GetCommonField(ChannelName,BasicType,byref CommonFieldArr,itemname)
	  select case BasicType
	     Case 0
		  CommonFieldArr="classid,id,foldername,createdate,creater"
		  GetCommonField=GetCommonField &"<optgroup  style=""color:blue"" label=""=====��Ŀ��ĳ����ֶ�====="">"
		  GetCommonField=GetCommonField &"<option value=""ClassID"">����Ŀ�Զ����ID(Url)</option>"
		  GetCommonField=GetCommonField &"<option value=""ID"">����Ŀ���ID(Url)</option>"
		  GetCommonField=GetCommonField &"<option value=""FolderName"">����Ŀ����</option>"
		  GetCommonField=GetCommonField &"<option value=""CreateDate"">����Ŀ����ʱ��</option>"
		  GetCommonField=GetCommonField &"<option value=""Creater"">����Ŀ������</option>"
		  GetCommonField=GetCommonField &"<optgroup style=""color:green"" label=""=====�����ֶβ�����ʹ��====="">"
	     Case 1
		  CommonFieldArr="id,tid,title,author,editor,origin,inputer,adddate,hits,articlecontent,photourl,rank,Intro"
		  GetCommonField=GetCommonField &"<optgroup  style=""color:blue"" label=""=====" & ChannelName &"�ĳ����ֶ�====="">"
		  GetCommonField=GetCommonField &"<option value=""ID"">��" & itemname & "�Զ����ID(Url)</option>"
		  GetCommonField=GetCommonField &"<option value=""Tid"">��" & itemname & "��ĿID(Url|����)</option>"
		  GetCommonField=GetCommonField &"<option value=""Title"">��" & itemname & "����</option>"
		  GetCommonField=GetCommonField &"<option value=""Author"">��" & itemname & "����</option>"
		  GetCommonField=GetCommonField &"<option value=""Inputer"">��" & itemname & "¼��Ա</option>"
		  GetCommonField=GetCommonField &"<option value=""Origin"">��" & itemname & "��Դ</option>"
		  GetCommonField=GetCommonField &"<option value=""Adddate"">��" & itemname & "���/����ʱ��</option>"
		  GetCommonField=GetCommonField &"<option value=""Hits"">��" & itemname & "�������</option>"
		  GetCommonField=GetCommonField &"<option value=""rank"">���Ķ��ȼ�</option>"
		  GetCommonField=GetCommonField &"<option value=""Intro"">��" & itemname & "����</option>"
		  GetCommonField=GetCommonField &"<option value=""Articlecontent"">��" & itemname & "��ϸ����</option>"
		  GetCommonField=GetCommonField &"<option value=""PhotoUrl"">��" & itemname & "ͼƬ��ַ</option>"
		  GetCommonField=GetCommonField &"<optgroup style=""color:green"" label=""=====" & ChannelName & "�������ֶ�====="">"
		 Case 2
		  CommonFieldArr="id,tid,title,author,photourl,origin,inputer,adddate,hits,hitsbyday,hitsbyweek,hitsbymonth,picturecontent,score,rank"
		  GetCommonField=GetCommonField &"<optgroup  style=""color:blue"" label=""=====" & ChannelName & "�ĳ����ֶ�====="">"
		  GetCommonField=GetCommonField &"<option value=""ID"">��" & itemname & "�Զ����ID(Url)</option>"
		  GetCommonField=GetCommonField &"<option value=""Tid"">��" & itemname & "��ĿID(Url|����)</option>"
		  GetCommonField=GetCommonField &"<option value=""Title"">��" & itemname & "����</option>"
		  GetCommonField=GetCommonField &"<option value=""PhotoUrl"">��" & itemname & "��ַ</option>"
		  GetCommonField=GetCommonField &"<option value=""Author"">��" & itemname & "����</option>"
		  GetCommonField=GetCommonField &"<option value=""Inputer"">��" & itemname & "¼��Ա</option>"
		  GetCommonField=GetCommonField &"<option value=""Origin"">��" & itemname & "��Դ</option>"
		  GetCommonField=GetCommonField &"<option value=""Adddate"">��" & itemname & "���/����ʱ��</option>"
		  GetCommonField=GetCommonField &"<option value=""Hits"">�����������</option>"
		  GetCommonField=GetCommonField &"<option value=""HitsByDay"">�����������</option>"
		  GetCommonField=GetCommonField &"<option value=""HitsByWeek"">�����������</option>"
		  GetCommonField=GetCommonField &"<option value=""HitsByMonth"">�����������</option>"
		  GetCommonField=GetCommonField &"<option value=""Score"">����Ʊ��</option>"
		  GetCommonField=GetCommonField &"<option value=""Rank"">���Ƽ��ȼ�</option>"
		  GetCommonField=GetCommonField &"<option value=""picturecontent"">��" & itemname & "����</option>"
		  GetCommonField=GetCommonField &"<optgroup style=""color:green"" label=""====" & ChannelName & "�������ֶ�====="">"
		 Case 3
		  CommonFieldArr="id,tid,title,author,downlb,downyy,downsq,downsize,ysdz,photourl,origin,inputer,adddate,hits,hitsbyday,hitsbyweek,hitsbymonth,downcontent"
		  GetCommonField=GetCommonField &"<optgroup  style=""color:blue"" label=""=====" & ChannelName & "�ĳ����ֶ�====="">"
		  GetCommonField=GetCommonField &"<option value=""ID"">��" & itemname & "�Զ����ID(Url)</option>"
		  GetCommonField=GetCommonField &"<option value=""Tid"">��" & itemname & "��ĿID(Url|����)</option>"
		  GetCommonField=GetCommonField &"<option value=""Title"">��" & itemname & "����</option>"
		  GetCommonField=GetCommonField &"<option value=""PhotoUrl"">��" & itemname & "����ͼ</option>"
		  GetCommonField=GetCommonField &"<option value=""Author"">��" & itemname & "����</option>"
		  GetCommonField=GetCommonField &"<option value=""DownLB"">��" & itemname & "���</option>"
		  GetCommonField=GetCommonField &"<option value=""DownYY"">��" & itemname & "����</option>"
		  GetCommonField=GetCommonField &"<option value=""DownSQ"">��" & itemname & "��Ȩ</option>"
		  GetCommonField=GetCommonField &"<option value=""DownSize"">��" & itemname & "��С</option>"
		  GetCommonField=GetCommonField &"<option value=""YSDZ"">����ʾ��ַ</option>"
		  GetCommonField=GetCommonField &"<option value=""Inputer"">��" & itemname & "¼��Ա</option>"
		  GetCommonField=GetCommonField &"<option value=""Origin"">��" & itemname & "��Դ</option>"
		  GetCommonField=GetCommonField &"<option value=""Adddate"">��" & itemname & "���/����ʱ��</option>"
		  GetCommonField=GetCommonField &"<option value=""Hits"">�����������</option>"
		  GetCommonField=GetCommonField &"<option value=""HitsByDay"">�����������</option>"
		  GetCommonField=GetCommonField &"<option value=""HitsByWeek"">�����������</option>"
		  GetCommonField=GetCommonField &"<option value=""HitsByMonth"">�����������</option>"
		  GetCommonField=GetCommonField &"<option value=""downcontent"">��" & itemname & "����</option>"
		  GetCommonField=GetCommonField &"<optgroup style=""color:green"" label=""=====" & ChannelName & "�������ֶ�====="">"
		 Case 4
		  CommonFieldArr="id,tid,title,author,photourl,origin,inputer,adddate,hits,hitsbyday,hitsbyweek,hitsbymonth,flashcontent,score,rank"
		  GetCommonField=GetCommonField &"<optgroup  style=""color:blue"" label=""=====" & ChannelName & "�ĳ����ֶ�====="">"
		  GetCommonField=GetCommonField &"<option value=""ID"">��" & itemname & "�Զ����ID(Url)</option>"
		  GetCommonField=GetCommonField &"<option value=""Tid"">��" & itemname & "��ĿID(Url|����)</option>"
		  GetCommonField=GetCommonField &"<option value=""Title"">��" & itemname & "����</option>"
		  GetCommonField=GetCommonField &"<option value=""PhotoUrl"">��ͼƬ��ַ</option>"
		  GetCommonField=GetCommonField &"<option value=""Author"">��" & itemname & "����</option>"
		  GetCommonField=GetCommonField &"<option value=""Inputer"">��" & itemname & "¼��Ա</option>"
		  GetCommonField=GetCommonField &"<option value=""Adddate"">��" & itemname & "���/����ʱ��</option>"
		  GetCommonField=GetCommonField &"<option value=""Hits"">�����������</option>"
		  GetCommonField=GetCommonField &"<option value=""HitsByDay"">�����������</option>"
		  GetCommonField=GetCommonField &"<option value=""HitsByWeek"">�����������</option>"
		  GetCommonField=GetCommonField &"<option value=""HitsByMonth"">�����������</option>"
		  GetCommonField=GetCommonField &"<option value=""Score"">����Ʊ��</option>"
		  GetCommonField=GetCommonField &"<option value=""Rank"">���Ƽ��ȼ�</option>"
		  GetCommonField=GetCommonField &"<option value=""flashcontent"">��" & itemname & "����</option>"
		  GetCommonField=GetCommonField &"<optgroup style=""color:green"" label=""=====" & ChannelName & "�������ֶ�====="">"
		 Case 5
		  CommonFieldArr="id,tid,title,author,photourl,bigphoto,promodel,inputer,adddate,hits,prospecificat,producername,trademarkname,ProIntro,score,rank,price,price_member,price_market,price_original,serviceterm,totalnum,unit,discount"
		  GetCommonField=GetCommonField &"<optgroup  style=""color:blue"" label=""=====" & ChannelName & "�ĳ����ֶ�====="">"
		  GetCommonField=GetCommonField &"<option value=""ID"">��" & itemname & "�Զ����ID(Url)</option>"
		  GetCommonField=GetCommonField &"<option value=""Tid"">��" & itemname & "��ĿID(Url|����)</option>"
		  GetCommonField=GetCommonField &"<option value=""Title"">��" & itemname & "����</option>"
		  GetCommonField=GetCommonField &"<option value=""PhotoUrl"">��" & itemname & "Сͼ</option>"
		  GetCommonField=GetCommonField &"<option value=""BigPhoto"">��" & itemname & "��ͼ</option>"
		  GetCommonField=GetCommonField &"<option value=""Price"">����ǰ�۸�</option>"
		  GetCommonField=GetCommonField &"<option value=""Price_Member"">����Ա�۸�</option>"
		  GetCommonField=GetCommonField &"<option value=""Price_Market"">���г��۸�</option>"
		  GetCommonField=GetCommonField &"<option value=""Price_Original"">��ԭʼ���ۼ�</option>"
		  GetCommonField=GetCommonField &"<option value=""Discount"">���ۿ�</option>"
		  GetCommonField=GetCommonField &"<option value=""ServiceTerm"">����������</option>"
		  GetCommonField=GetCommonField &"<option value=""TotalNum"">���������</option>"
		  GetCommonField=GetCommonField &"<option value=""ProModel"">��" & itemname & "�ͺ�</option>"
		  GetCommonField=GetCommonField &"<option value=""Unit"">��" & itemname & "��λ</option>"
		  GetCommonField=GetCommonField &"<option value=""Inputer"">��" & itemname & "¼��Ա</option>"
		  GetCommonField=GetCommonField &"<option value=""Adddate"">������ʱ��</option>"
		  GetCommonField=GetCommonField &"<option value=""Hits"">�����������</option>"
		  GetCommonField=GetCommonField &"<option value=""ProSpecificat"">����Ʒ���</option>"
		  GetCommonField=GetCommonField &"<option value=""ProducerName"">��������</option>"
		  GetCommonField=GetCommonField &"<option value=""TrademarkName"">��Ʒ��/�̱�</option>"
		  GetCommonField=GetCommonField &"<option value=""Rank"">���Ƽ��ȼ�</option>"
		  GetCommonField=GetCommonField &"<option value=""ProIntro"">��" & itemname & "����</option>"
		  GetCommonField=GetCommonField &"<optgroup style=""color:green"" label=""=====" & ChannelName & "�������ֶ�====="">"
		 Case 7
		  CommonFieldArr="id,tid,title,movieact,photourl,movietime,moviedy,adddate,hits,screentime,movieyy,moviedq,moviecontent,score,rank,inputer"
		  GetCommonField=GetCommonField &"<optgroup  style=""color:blue"" label=""=====" & ChannelName & "�ĳ����ֶ�====="">"
		  GetCommonField=GetCommonField &"<option value=""ID"">��" & itemname & "�Զ����ID(Url)</option>"
		  GetCommonField=GetCommonField &"<option value=""Tid"">��" & itemname & "��ĿID(Url|����)</option>"
		  GetCommonField=GetCommonField &"<option value=""Title"">��" & itemname & "����</option>"
		  GetCommonField=GetCommonField &"<option value=""PhotoUrl"">��ͼƬ��ַ</option>"
		  GetCommonField=GetCommonField &"<option value=""MovieAct"">����Ҫ��Ա</option>"
		  GetCommonField=GetCommonField &"<option value=""MovieDY"">��" & itemname & "����</option>" 
		  GetCommonField=GetCommonField &"<option value=""MovieTime"">�����ų���</option>"
		  GetCommonField=GetCommonField &"<option value=""ScreenTime"">����ӳʱ��</option>"
		  GetCommonField=GetCommonField &"<option value=""MovieYY"">��" & itemname & "����</option>"
		  GetCommonField=GetCommonField &"<option value=""MovieDQ"">����������</option>"		  
		  GetCommonField=GetCommonField &"<option value=""Adddate"">��" & itemname & "���/����ʱ��</option>"
		  GetCommonField=GetCommonField &"<option value=""Hits"">�����������</option>"
		  GetCommonField=GetCommonField &"<option value=""Score"">����Ʊ��</option>"
		  GetCommonField=GetCommonField &"<option value=""Rank"">���Ƽ��ȼ�</option>"
		  GetCommonField=GetCommonField &"<option value=""MovieContent"">��" & itemname & "����</option>"
		  GetCommonField=GetCommonField &"<option value=""Inputer"">��" & itemname & "¼��Ա</option>"
		  GetCommonField=GetCommonField &"<optgroup style=""color:green"" label=""=====" & ChannelName & "�������ֶ�====="">"
		 Case 8
		  CommonFieldArr="id,tid,title,author,photourl,address,contactman,province,adddate,hits,city,companyname,tel,gqcontent,zip,username,fax,email,homepage,validdate,price"
		  GetCommonField=GetCommonField &"<optgroup  style=""color:blue"" label=""=====" & ChannelName & "�ĳ����ֶ�====="">"
		  GetCommonField=GetCommonField &"<option value=""ID"">����Ϣ�Զ����ID(Url)</option>"
		  GetCommonField=GetCommonField &"<option value=""Tid"">��" & itemname & "��ĿID(Url|����)</option>"
		  GetCommonField=GetCommonField &"<option value=""Title"">��" & itemname & "��������</option>"
		  GetCommonField=GetCommonField &"<option value=""gqcontent"">��" & itemname & "����ϸ����</option>"
		  GetCommonField=GetCommonField &"<option value=""PhotoUrl"">��" & itemname & "ͼƬ��ַ</option>"
		  GetCommonField=GetCommonField &"<option value=""Adddate"">��" & itemname & "���/����ʱ��</option>"
		  GetCommonField=GetCommonField &"<option value=""ValidDate"">����Ч����</option>"
		  GetCommonField=GetCommonField &"<option value=""Price"">���۸�</option>"
		  GetCommonField=GetCommonField &"<option value=""Hits"">���������</option>"
		  GetCommonField=GetCommonField &"<option value=""UserName"">��������Ա��</option>"
		  GetCommonField=GetCommonField &"<option value=""ContactMan"">����ϵ��</option>"
		  GetCommonField=GetCommonField &"<option value=""Address"">����ϵ��ַ</option>"
		  GetCommonField=GetCommonField &"<option value=""Tel"">����ϵ�绰</option>"
		  GetCommonField=GetCommonField &"<option value=""Fax"">���������</option>"
		  GetCommonField=GetCommonField &"<option value=""Email"">����������</option>"
		  GetCommonField=GetCommonField &"<option value=""Zip"">����������</option>"
		  GetCommonField=GetCommonField &"<option value=""HomePage"">����ҳ��ַ</option>"
		  GetCommonField=GetCommonField &"<option value=""Province"">������ʡ��</option>"
		  GetCommonField=GetCommonField &"<option value=""City"">�����ڳ���</option>"
		  GetCommonField=GetCommonField &"<option value=""CompanyName"">����˾����</option>"
		  GetCommonField=GetCommonField &"<optgroup style=""color:green"" label=""=====" & ChannelName & "�������ֶ�====="">"
		 case else
		  GetCommonField=""
	  end Select
	End Function

	Function GetFieldName(EField,itemname)
	   if datasourcetype<>0 then GetFieldName=Efield:exit function
	  Select Case Lcase(EField)
	  case "classid"
	     GetFieldName="��Ŀ�Զ����ClassID"
	   case "fnametype"
	     GetFieldName="���ɵ���Ϣ��չ��"
	   case "creater"
	     GetFieldName="��Ŀ������"
	   case "createdate"
	     GetFieldName="��Ŀ����ʱ��"
	   case "templateid"
	     GetFieldName="��Ŀ�µ���Ϣģ��ID"
	   case "channelid"
	     GetFieldName="ģ��ID"
	   case "cirspecialshowtf"
	     GetFieldName="ѭ����Ŀר����ʾ��־"
	   case "classbasicinfo"
	     GetFieldName="��Ŀ��Ϣ���ü���"
	   case "classdefinecontent"
	     GetFieldName="��Ŀ�������ݼ���"
	   case "classpurview"
	     GetFieldName="��Ŀ���Ȩ��ID"
	   case "commenttf"
	     GetFieldName="��Ŀ����Ϣ�������۱�־"
	   case "defaultarrgroupid"
	     GetFieldName="��Ŀ��Ĭ��ָ����Ա��Ĳ鿴Ȩ��"
	   case "defaultchargetype"
	     GetFieldName="��Ŀ��Ĭ���ظ��շѷ�ʽ"
	   case "defaultdividepercent"
	     GetFieldName="��Ŀ����Ͷ���ߵ�Ĭ�Ϸֳɱ���"
	   case "defaultpitchtime"
	     GetFieldName="��Ŀ�µ�Ĭ���ظ��շѲ鿴����"
	   case "defaultreadpoint"
	     GetFieldName="��Ŀ�µ�Ĭ���շѵ���"
	   case "defaultreadtimes"
	     GetFieldName="��Ŀ�µ�Ĭ���ظ��շѲ鿴����"
	   case "folder"
	     GetFieldName="Ŀ¼Ӣ������"
	   case "folderdomain"
	     GetFieldName="��Ŀ�󶨵�����"
	   case "folderfsoindex"
	     GetFieldName="��Ŀ���ɵ���ҳ����"
	   case "folderorder"
	     GetFieldName="��Ŀ�������"
	   case "foldertemplateid"
	     GetFieldName="��Ŀ��ģ��ID"
	   case "specialtemplateid"
	     GetFieldName="Ƶ��ר���б�ҳģ��ID"
	   case "tn"
	     GetFieldName="����ĿID"
	   case "tj"
	     GetFieldName="��Ŀ���"
	   case "ts"
	     GetFieldName="��ĿID�����б�"
	   case "topflag"
	     GetFieldName="����������ʾ��־"
	   case "movieact"
	     GetFieldName= itemname & "��Ա"
	   case "moviedq"
	     GetFieldName=itemname & "����"
	   case "moviedy"
	     GetFieldName=itemname & "����"
	   case "movietime"
	     GetFieldName="���ų���"
	   case "screentime"
	     GetFieldName="��ӳʱ��"
	   case "movieyy"
	     GetFieldName=itemname & "����"
	   case "moviecontent"
	     GetFieldName=itemname & "����"
	   case "movietype"
	     GetFieldName=itemname & "���Ÿ�ʽID"
	   case "movieurls"
	     GetFieldName=itemname & "���ŵ�ַ"
	   case "serverid"
	     GetFieldName="���ŷ�����ID"
	   case "alarmnum"
	     GetFieldName="���ޱ�����"
	   case "producttype"
	     GetFieldName="��������ID"
	   case "isspecial"
	     GetFieldName="�ؼ۱�־"
	   case "price_member"
	     GetFieldName="��Ա�۸�"
	   case "price_market"
	     GetFieldName="�г��۸�"
	   case "price_original"
	     GetFieldName="ԭʼ���ۼ�"
	   case "discount"
	     GetFieldName="�ۿ�"
	  case "serviceterm"
	     GetFieldName="��������"
	   case "totalnum"
	     GetFieldName="�����"
	   case "promodel"
	     GetFieldName=itemname & "�ͺ�"
	  case "unit"
	     GetFieldName=itemname & "��λ" 
	  case "producername"
	     GetFieldName="������" 
	  case "prospecificat"
	     GetFieldName=itemname & "���" 
	  case "trademarkname"
	     GetFieldName="Ʒ��/�̱�" 
	  case "point"
	     GetFieldName="�������" 
	  case "prointro"
	     GetFieldName=itemname & "����" 
	   case "newsid","picid","downid","flashid","movieid","proid","gqid"
	     GetFieldName="ϵͳ���ɵ�ΨһID(url)"
	   case "picurls"
	     GetFieldName="ͼƬ��ַ����"
	   case "score"
	     GetFieldName="��Ʊ��"
	   Case "adddate"
	    GetFieldName="���/����ʱ��"
	   Case "tid"
	    GetFieldName="��ĿID(Url|����)"
	   case "arrgroupid"
	    GetFieldName="��Ȩ�鿴�Ļ�Ա��ID"
	   case "articlecontent"
	    GetFieldName=itemname & "��ϸ����"
	   case "inputer"
	    GetFieldName=itemname & "¼��Ա"
	   case "photourl"
	    GetFieldName="ͼƬ��ַ"
	   case "bigphoto"
	    GetFieldName="��ͼƬ��ַ"
	   case "hitsbyday"
	    GetFieldName="���������"
	   case "hitsbyweek"
	    GetFieldName="���������"
	   case "hitsbymonth"
	    GetFieldName="���������"
	   case "picturecontent"
	    GetFieldName=itemname & "����"
	   case "author"
	    GetFieldName="����"
	   case "origin"
	    GetFieldName="��Դ"
	   case "picurl"
	    GetFieldName="ͼƬ��ַ"
	   case "downlb"
	    GetFieldName="������"
	   case "downpt"
	    GetFieldName="���ƽ̨"
	   case "downsize"
	    GetFieldName="�����С"
	   case "downsq"
	    GetFieldName="��Ȩ��ʽ"
	   case "downyy"
	    GetFieldName=itemname & "����"
	   case "ysdz"
	    GetFieldName=itemname & "��ʾ��ַ"
	   case "zcdz"
	    GetFieldName=itemname & "ע���ַ"
	   case "downversion"
	    GetFieldName=itemname & "�汾"
	   case "downurls"
	    GetFieldName="���ص�ַ����"
	   case "inputer"
	    GetFieldName=itemname & "¼��Ա"
	   case "downcontent"
	    GetFieldName=itemname & "���"
	   case "flashurl"
	    GetFieldName=itemname & "��ַ"
	   case "flashcontent"
	    GetFieldName=itemname & "����"		
	   case "typeid"
	    GetFieldName="�������ID"			
	   case "gqcontent"
	   	GetFieldName="������ϸ����"		
	   case "validdate"
	   	GetFieldName="��Ч����"		
	   case "price"
	   	GetFieldName="�۸�"
	   case "username"
	    GetFieldName="�û���"
	   case "contactman"
	    GetFieldName="��ϵ��"
	   case "address"
	    GetFieldName="��ϵ��ַ"
	   case "tel"
	    GetFieldName="��ϵ�绰"
	   case "fax"
	    GetFieldName="�������"
	   case "email"
	    GetFieldName="��������"
	   case "zip"
	    GetFieldName="��������"
	   case "homepage"
	    GetFieldName="��˾��ҳ"		
	   case "province"
	    GetFieldName="ʡ��"		
	   case "city"
	    GetFieldName="����"		
	   case "companyname"
	    GetFieldName="��˾����"		
	   case "beyondsavepic"
	    GetFieldName="Զ�̴�ͼ��־"
	   case "changes"
	    GetFieldName="ת�����ӱ�־"
	   case "chargetype"
	    GetFieldName="�ظ��շѷ�ʽ"
	   case "comment"
	    GetFieldName="�������۱�־"
	   case "deltf"
	    GetFieldName="�������վ��־"
	   case "dividepercent"
	     GetFieldName="Ͷ��ֳɱ���"
	   case "fname"
	     GetFieldName="���ɵ��ļ���"
	   case "fulltitle"
	     GetFieldName="������������"
	  case "infopurview"
	     GetFieldName="�Ķ�Ȩ�޷�ʽ"
	  case "istop"
	     GetFieldName="�ö���־"
	  case "jsid"
	     GetFieldName="�����JSID�б�"
	  case "keywords"
	     GetFieldName="�ؼ���"
	  case "picnews"
	     GetFieldName="ͼƬ���ű�־"
	  case "pitchtime"
	     GetFieldName="�ظ��շ�Сʱ��"
	  case "popular"
	     GetFieldName="���ű�־"
	  case "rank"
	     GetFieldName="�ȼ�"
	  case "readpoint"
	     GetFieldName="��Ҫ���Ķ�����"
	  case "readtimes"
	     GetFieldName="�Ķ�ָ�����������շ�"
	  case "recommend"
	     GetFieldName="�Ƽ���־"
	  case "refreshtf"
	     GetFieldName="�����ɱ�־"
	  case "rolls"
	     GetFieldName="������־"
	  case "showcomment"
	     GetFieldName="��������ʾ���۱�־"
	  case "slide"
	     GetFieldName="�õ�Ƭ��־"
	  case "specialid"
	     GetFieldName="ר��ID"
	  case "strip"
	     GetFieldName="ͷ����־"
	  case "templateid"
	     GetFieldName="ģ��ID"
	  case "titlefontcolor"
	     GetFieldName="������ɫ"
	  case "titlefonttype"
	     GetFieldName="����Ӵ�+б���־"
	  case "titletype"
	     GetFieldName="ͼ�ı�־"
	   case "hits"
	     GetFieldName="�������"
	   case "verific"
	     GetFieldName="��˱�־"
	   case "id"
	     GetFieldName="�Զ����ID(Url)"
	   case "foldername"
	     GetFieldName="��Ŀ����"
	   case "lasthitstime"
	     GetFieldName="��������ʱ��"
	   case "title"
	      If InStr(lcase(LabelIntro),"ks_article") then
	       GetFieldName=itemname & "����"
		  else
		   GetFieldName=itemname & "����"
		  end if
	 Case "Intro"
	     GetFieldName="��������"
	   Case else
	    GetFieldName=efield
	  End Select
	   'GetFieldName=Efield&"(" & GetFieldName&"��"
	  ' GetFieldName=GetFieldName
	End Function
	 
	  '������
	 Sub Step2()
	    Dim FieldParam,FieldParamArr,LoopTimes
		LabelName = Request.Form("LabelName")
		FolderID=KS.G("ParentID")
		Ajax=KS.G("Ajax")
	    SQLType=KS.G("SQLType")
		PageStyle=KS.G("PageStyle")
		ItemName=KS.G("ItemName")
        if datasourcetype<>0 then Call OpenExtConn()
		With KS
		Set LabelRS = Server.CreateObject("Adodb.RecordSet")
		If LabelID <> "" Then
		    ActionStr="EditSubmit"
			Set LabelRS = Server.CreateObject("Adodb.Recordset")
			SQLStr = "SELECT top 1 * FROM [KS_Label] Where ID='" & LabelID & "'"
			LabelRS.Open SQLStr, Conn, 1, 1
			If Not LabelRS.Eof Then
				LabelContent = Server.HTMLEncode(LabelRS("LabelContent"))
				FieldParamArr= Split(LabelRS("Description"),"@@@")
			End IF
			If Not KS.IsNul(Request("LabelIntro")) Then
			LabelIntro=Request("LabelIntro")
			Else
			LabelIntro =FieldParamArr(0)
			End If
			If Ubound(FieldParamArr)>=1 Then
				FieldParam =FieldParamArr(1)
			End If
			LabelRS.Close
		Else
		  LabelIntro=request("LabelIntro")
		  FieldParam=KS.G("FieldParam")
		  ActionStr="AddNewSubmit"
		  LoopTimes=GetLoopTimes(LabelIntro)
		  LabelContent="[loop=" & LoopTimes &"]���ڴ�����ѭ������[/loop]"
		End If
		 Call SqlValid(LabelIntro)
		%>
		<script src="../../ks_inc/kesion.box.js"></script>
		<script language="javascript">
		var pos=null;
		function setPos()
		{ if (document.all){
			document.myform.LabelContent.focus();
		    pos = document.selection.createRange();
		  }else{
		    pos = document.getElementById("LabelContent").selectionStart;
		  }
		}
		function FieldInsertCode(fieldname,dbtype,dbname)
		{ 
		   if(pos==null) {alert('���ȶ�λ����λ��!');return false;}
		   var link="Admin_FieldParam.asp?fieldname=" + fieldname + "&dbtype="+ dbtype + "&dbname=" + dbname+"&datasourcetype=<%=datasourcetype%>";
		  PopupImgDir="../";
		  PopupCenterIframe('�����ֶα�ǩ',link,350,230,'no');
		}
		
		//���뵽ѭ����
		function InsertValue(Val)
		{
			 if (document.all){
			  pos.text=Val;
			 }else{
			   var obj=$("#LabelContent");
			   var lstr=obj.val().substring(0,pos);
			   var rstr=obj.val().substring(pos);
			   obj.val(lstr+Val+rstr);
			 }
		}
		function FieldInsertCode1(Val)
		{ 
		
		  if (Val!=''){
		   InsertValue(Val);
		   }
		}
		</script>
		<script language = 'JavaScript'>

		function show_ln(txt_ln,txt_main){
			var txt_ln  = document.getElementById(txt_ln);
			var txt_main  = document.getElementById(txt_main);
			txt_ln.scrollTop = txt_main.scrollTop;
			while(txt_ln.scrollTop != txt_main.scrollTop)
			{
				txt_ln.value += (i++) + '\n';
				txt_ln.scrollTop = txt_main.scrollTop;
			}
			return;
		}
		
		</script>
       <script src="../../ks_inc/kesion.box.js"></script>
		<script>
		function ShowIframe()
		{   PopupImgDir="../";
			PopupCenterIframe("�鿴��Ŀ<=>ID���ձ�","?action=ShowClassID",600,350,"auto")
		}
		</script>		
		<%
		.echo "<table width=""100%"" height=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"">"
		.echo "  <form name=""myform"" id=""myform"" method=post action=""LabelFunctionAdd.asp"">"
		.echo "    <input type='hidden' name='keyword' id='keyword' value='" & KeyWord & "'>"
		.echo "    <input type='hidden' name='Searchtype' id='Searchtype' value='" & searchtype & "'>"
		.echo "    <input type=""hidden"" name=""LabelFlag"" id='LabelFlag' value=""3"">"
		.echo "    <input type=""hidden"" name=""LabelID"" id='LabelID' value=""" & LabelID & """>"
		.echo "    <input type=""hidden"" name=""FolderID"" id='FolderID' value=""" & FolderID & """>"
		.echo "    <input type=""hidden"" name=""Page"" id='Page' value=""" & Page & """>"
		.echo "    <input type=""hidden"" name=""FieldParam"" id='FieldParam' value=""" & FieldParam & """>"
		.echo "    <input type='hidden' name='Action' id='Action' value='" & ActionStr & "'>"
		
		.echo " <input type=""hidden"" name=""LabelName"" id=""LabelName"" value=""" &LabelName & """>"
		.echo " <input type=""hidden"" name=""SQLType"" id=""SQLType"" value=""" &SQLType & """>"
		.echo " <input type=""hidden"" name=""Ajax"" id=""Ajax"" value=""" & Ajax & """>"
		.echo " <input type=""hidden"" name=""ItemName"" id=""ItemName"" value=""" & ItemName & """>"
		.echo " <input type=""hidden"" name=""PageStyle"" id=""PageStyle"" value=""" & PageStyle & """>"
		.echo " <input type=""hidden"" name=""Note"" id=""Note"" value=""" & note & """ size=""6""> "
		
		.echo " <input type=""hidden"" name=""datasourcetype"" id=""datasourcetype"" value=""" & datasourcetype & """ size=""6""> "
		.echo " <input type=""hidden"" name=""datasourcestr"" id=""datasourcestr"" value=""" & datasourcestr & """> "

		.echo " <tr>"
		.echo "   <td height=""25"" colspan=""2"" bgcolor='#efefef' class='sort'> "
		.echo "    <div align='center'><font color='#990000'>"
		.echo "��������������ǩ��ʽ��ѭ�����ݣ�</font> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href='javascript:ShowIframe()'><u>�鿴��Ŀ<=>ID���ձ�</u></a>"
		.echo " </div></td>"
		.echo "    </tr>"
		.echo "    <tr style='display:none'>"
		.echo "      <td width=""60"" height=""19"">��ǩĿ¼</td>"
		.echo "      <td>" & KS.ReturnLabelFolderTree(FolderID, 1) & "</td>"
		.echo "    </tr>"
		.echo "    <tr class=""tableBorder1"">"
		.echo "      <td height=""16""><div align=""left"">��ѯ���</div></td>"
		.echo "      <td><textarea name=""LabelIntro"" rows=""4"" style=""width:98%;"">" & LabelIntro & "</textarea></td>"
		.echo "    </tr>"
        .echo "    <tr class=""tableBorder1"">"
		.echo "      <td width=""60"" height=""30"" nowrap align=center><strong>�����ֶ�</strong></td>"
		.echo "      <td>"
		 Dim FieldName,dbtype,I,J,ClickStr,isidarr,isid
		 Dim RSField:Set RSField=Server.CreateObject("ADODB.RECORDSET")
		 Call OpenExtConn()
		 if datasourcetype<>0 then
		 RSField.Open ClearParam(LabelIntro),tConn,1,3
		 else
		 RSField.Open ClearParam(LabelIntro),Conn,1,3
		 end if
		  .echo "<table style=""table-layout:fixed"" border=1 bordercolordark=""#999999"" bordercolorlight=""#FFFFFF"" width=""710"" cellpadding='0' cellspacing='0'>"
		  .echo "<tr class='tdbg' height='20'>"
		  For I=0 To RSField.Fields.count-1
		     dbtype=RSField.Fields(i).type
			 FieldName=RSField.Fields(i).name
			 isidarr=split(FieldName,".")
				isid=false
				if ubound(isidarr)=1 then
				  if lcase(isidarr(1))="id" then
					isid=true
				  end if
				end if
			 If (Lcase(FieldName)="tid" or Lcase(FieldName)="id" or isid or Lcase(FieldName)="newsid" Or Lcase(FieldName)="picid" or Lcase(FieldName)="downid" or Lcase(FieldName)="flashid" or Lcase(FieldName)="proid" or Lcase(FieldName)="movieid" or Lcase(FieldName)="gqid" or Lcase(FieldName)="classid") and  datasourcetype=0  Then
			   
			    Dim sChannelID
			   	If DataBaseType=1 Then
				  dim rsc:set rsc=server.CreateObject("adodb.recordset")
				  rsc.open "Select ChannelID From KS_Channel Where charindex(channeltable,'" & ReplaceBC(LabelIntro) & "')>0",conn,1,1
				  if not rsc.eof then
				    sChannelID=rsc(0)
				  end if
				  rsc.close:set rsc=nothing
				Else
				  if not Conn.Execute("Select ChannelID From KS_Channel Where Instr('" & ReplaceBC(LabelIntro) & "',channeltable)>0").eof then
				   sChannelID=Conn.Execute("Select ChannelID From KS_Channel Where Instr('" & ReplaceBC(LabelIntro) & "',channeltable)>0")(0)
				  end if 
				End If
			   if sChannelID>=1 then
			     ClickStr="FieldInsertCode('" & FieldName & "',"&dbtype&"," & sChannelID & ")"
			   ElseIf Instr(Lcase(LabelIntro),"ks_class") then 
			     ClickStr="FieldInsertCode('" & FieldName & "',"&dbtype&",100)"
			   Else 
			     ClickStr="FieldInsertCode('" & FieldName & "',"&dbtype&",0)"
			   End If
			 Else
			   ClickStr="FieldInsertCode('" & FieldName & "',"&dbtype&",0)"
			 End IF
			 If j=5 Then j=0:.echo "</tr><tr class='tdbg' height='20'>"
			  J=J+1
			  if instr(FieldName,".")=0 then
			  .echo "<td  width=""20%"" align=""center"" onMouseOut=""this.className='tdbg'"" onMouseOver=""this.className='tdbgmouseover'"" style=""cursor:pointer;"" onClick=""" & ClickStr & """>" & GetFieldName(trim(FieldName),"") & "</td>"
			 else
			  .echo "<td  width=""20%"" align=""center"" onMouseOut=""this.className='tdbg'"" onMouseOver=""this.className='tdbgmouseover'"" style=""cursor:pointer;"" onClick=""" & ClickStr & """>" & split(FieldName,".")(0)&"." & GetFieldName(trim(split(FieldName,".")(1)),"") & "</td>"
			 end if
		 Next
		 For I=J+1 to 5
		  .echo "<td class='tdbg' height='25'>&nbsp;</td>"
		 Next
		 .echo "</tr>"
		 .echo "</table>"
		 .echo  "</td>"
		 .echo "  </tr>"
		 If FieldParam<>"" Then
		 .echo "<tr class=""tableBorder1""><td width=""60"" height=""30"" nowrap align=center><strong>��������</strong></td>"
		 .echo "<td>"
		 .echo "<table border=1 bordercolordark=""#999999"" bordercolorlight=""#FFFFFF"" width=""100%"" cellpadding='0' cellspacing='0'>"
		 .echo "<tr class='tdbg' height='20'>"
		 FieldParamArr=Split(FieldParam,vbcrlf)
		 J=0
		 For I=0 To Ubound(FieldParamArr)
		   If j=5 Then j=0:.echo "</tr><tr class='tdbg' height='20'>"
		   J=J+1
		 .echo "<td  width=""20%"" align=""center"" onMouseOut=""this.className='tdbg'"" onMouseOver=""this.className='tdbgmouseover'"" style=""cursor:pointer;"" onClick=""FieldInsertCode1('{$Param(" & I & ")}');"">" & FieldParamArr(I) &"</td>"
		 Next
		 For I=J+1 to 5
		  .echo "<td class='tdbg' height='25'>&nbsp;</td>"
		 Next
		 .echo "</tr>"
		.echo "</table>"
		.echo "</td></tr>"
		End If
		
		 .echo "   <tr class=""tableBorder1"">"
 
		 .echo "	<td align='center'><strong>ѭ �� ��</strong>{$AutoID}</td><td height='230' valign=""top""><textarea id='txt_ln' name='rollContent' cols='6' style='width:35px;overflow:hidden;height:100%;background-color:highlight;border-right:0px;text-align:right;font-family: tahoma;font-size:12px;font-weight:bold;color:highlighttext;cursor:default;' readonly>"
		 Dim N
		 For N=1 To 3000
			.echo N & "&#13;&#10;"
		 Next
		 .echo"</textarea>"
		 .echo "<textarea name='LabelContent'  onclick='setPos()' onkeyup='setPos()' id='LabelContent' style='width:670px;height:100%' rows='15' id='txt_main' onscroll=""show_ln('txt_ln','LabelContent')"" wrap='on'>" & LabelContent & "</textarea>" & vbNewLine
		 .echo "	<script>for(var i=3000; i<=3000; i++) document.getElementById('txt_ln').value += i + '\n';</script>"
		 .echo "   </td></tr>"
		 
		 .echo "   <tr class=""tableBorder1"">"
 
		 .echo "	<td><strong>��Ҫ����</strong></td><td><font color=red>1��SQL��ǩ�������</font><br>ѭ�����ʽ��[loop=n]ѭ����ǩ������[/loop]<br>����n��ʾѭ����������n����n>=0��loopΪѭ���ؼ��֣���ѭ��������ظ�ʹ��,���ǲ���Ƕ�ס�<font color=red><br>2��SQL��ǩ�ֶι���</font><br>�ֶθ�ʽ��{$Field(FieldName,OutType,Param,...)}<br>FieldName&nbsp;&nbsp;--���ݿ����ֶ�����<br>OutType&nbsp;&nbsp;&nbsp;&nbsp;--������� ֧�֣��ı�(Text)������(Date)������(Num)������URL(GetInfoUrl)����ĿURL(GetClassUrl) 5������<br><font color=red>3��֧��ʹ��{ReqNum(�ַ���)}��{ReqStr(�ַ���)}��ȡ��Url�Ĳ���ֵ</font><br>�磺http://www.kesion.com/index.asp?ClassID=100,��ô{ReqNum(ClassID)} ���õ�100<br/><font color=red>4.������/��ҵ�ռ�Ҫʹ��sql��ǩʱ,������<font color=red>""{$GetUserName}""</font>ȡ�õ�ǰ�ռ���û��� </font><br>��:select top 10 id,title from ks_article where inputer='{$GetUserName}' order by id desc"

		 .echo "   </td></tr>"
		

		.echo "  </form>"
		.echo "</table>"
		.echo "<script language=""JavaScript"">" & vbCrLf
		.echo "<!--" & vbCrLf
		.echo "function CheckForm()" & vbCrLf
		.echo "{ var form=document.myform;"
		.echo "  if (form.LabelName.value=='')"
		.echo "   {"
		.echo "    alert('�������ǩ����!');"
		.echo "    form.LabelName.focus();"
		.echo "    return false;"
		.echo "   }"
		 .echo " if (form.LabelContent.value==''||form.LabelContent.value=='[loop="&LoopTimes&"]���ڴ�����ѭ������[/loop]')"
		 .echo " {"
		 .echo "   alert('�������ǩѭ������!');"
		 .echo "   form.LabelContent.focus();"
		 .echo "   return false;"
		 .echo "  }"
		 .echo "  form.submit();"
		 .echo "  return true;"
		.echo "}" & vbCrLf
		.echo "//-->" & vbCrLf
		.echo "</script>"
		
		Set Conn = Nothing
		
		End With
End Sub

'����
Sub AddLabelSave()
			LabelName = KS.G("LabelName")
			Descript = Request("LabelIntro")
			FieldParam = Request("FieldParam")
			LabelContent = Trim(Request.Form("LabelContent"))
			LabelFlag = KS.G("LabelFlag")
			FolderID = KS.G("FolderID")
			SQLType =KS.G("SQLType")
			ItemName=KS.G("ItemName")
			PageStyle=KS.G("PageStyle")
			Ajax=KS.G("Ajax")
			If LabelName = "" Then
			   Call KS.AlertHistory("��ǩ���Ʋ���Ϊ��!", -1)
			   Set KS = Nothing
			   Exit Sub
			End If
			If SQLType=1 And ItemName="" Then
			  Call KS.AlertHistory("��ҳ��Ŀ����Ϊ��!", -1)
			  Set KS = Nothing
			  Exit Sub
			End IF
			If LabelContent = "" Then
			  Call KS.AlertHistory("��ǩ���ݲ���Ϊ��!", -1)
			  Set KS = Nothing
			  Exit Sub
			End If
			LabelName = "{SQL_" & LabelName & "}"
			Set LabelRS = Server.CreateObject("Adodb.RecordSet")
			LabelRS.Open "Select top 1 LabelName From [KS_Label] Where LabelName='" & LabelName & "'", Conn, 1, 1
			If Not LabelRS.EOF Then
			  Call KS.AlertHistory("��ǩ�����Ѿ�����!", -1)
			  LabelRS.Close
			  Conn.Close
			  Set LabelRS = Nothing
			  Set Conn = Nothing
			  Set KS = Nothing
			 Exit Sub
			Else
				LabelRS.Close
				LabelRS.Open "Select  top 1 * From [KS_Label] Where (ID is Null)", Conn, 1, 3
				LabelRS.AddNew
				  Do While True
					'����ID  ��+12λ���
					LabelID = Year(Now()) & KS.MakeRandom(10)
					Set RSCheck = Conn.Execute("Select ID from [KS_Label] Where ID='" & LabelID & "'")
					 If RSCheck.EOF And RSCheck.BOF Then
					  RSCheck.Close
					  Set RSCheck = Nothing
					  Exit Do
					 End If
				  Loop
				 LabelRS("ID") = LabelID
				 LabelRS("LabelName") = LabelName
				 LabelRS("LabelContent") = LabelContent
				 LabelRS("LabelFlag") = LabelFlag
				 LabelRS("Description") = Descript &"@@@"&FieldParam&"@@@"&SQLType&"@@@"&ItemName&"@@@"&PageStyle&"@@@"&Ajax&"@@@"& datasourcetype &"@@@" &datasourcestr & "@@@" & note
				 LabelRS("FolderID") = FolderID
				 LabelRS("AddDate") = Now
				 LabelRS("LabelType") = 5
				 LabelRS("OrderID") = 1
				 LabelRS.Update
				 Call KS.FileAssociation(1021,2,LabelContent,0)
				 KS.echo ("<script>if (confirm('�ɹ���ʾ:\n\n��ӱ�ǩ�ɹ�,������ӱ�ǩ��?')){location.href='LabelFunctionAdd.asp?Action=AddNew&LabelType=5&FolderID=" & FolderID & "';}else{$(parent.document).find('#BottomFrame')[0].src='" & KS.Setting(3) & KS.Setting(89) & "KS.Split.asp?LabelFolderID=" & FolderID & "&OpStr=��ǩ���� >> ���嶨������ǩ&ButtonSymbol=DIYFunctionLabel';parent.frames['MainFrame'].location.href='Label_Main.asp?LabelType=5&FolderID=" & FolderID & "';}</script>")
			End If
	End Sub
	
	'�����޸�
	Sub EditLabelSave()
			LabelID = Trim(Request.Form("LabelID"))
			FolderID = Request.Form("FolderID")
			LabelName = Replace(Replace(Trim(Request.Form("LabelName")), """", ""), "'", "")
			Descript = Request("LabelIntro")
			FieldParam = Request("FieldParam")
			SQLType =KS.G("SQLType")
			ItemName=KS.G("ItemName")
			PageStyle=KS.G("PageStyle")
			Ajax=KS.G("Ajax")
			Call SqlValid(Descript)
			LabelContent = Trim(Request.Form("LabelContent"))
			LabelFlag = Request.Form("LabelFlag")
			If LabelName = "" Then
			   Call KS.AlertHistory("��ǩ���Ʋ���Ϊ��!", -1)
			   Set KS = Nothing
			   Exit Sub
			End If
			If SQLType=1 And ItemName="" Then
			  Call KS.AlertHistory("��ҳ��Ŀ����Ϊ��!", -1)
			  Set KS = Nothing
			  Exit Sub
			End IF
			If LabelContent = "" Then
			  Call KS.AlertHistory("��ǩ���ݲ���Ϊ��!", -1)
			  Set KS = Nothing
			  Exit Sub
			End If
			LabelName = "{SQL_" & LabelName & "}"
			Set LabelRS = Server.CreateObject("Adodb.RecordSet")
			LabelRS.Open "Select LabelName From [KS_Label] Where ID <>'" & LabelID & "' AND LabelName='" & LabelName & "'", Conn, 1, 1
			If Not LabelRS.EOF Then
			  Call KS.AlertHistory("��ǩ�����Ѿ�����!", -1)
			  LabelRS.Close:Conn.Close:Set LabelRS = Nothing:Set Conn = Nothing
			  Set KS = Nothing
			  Exit Sub
			Else
				LabelRS.Close
				LabelRS.Open "Select top 1 * From [KS_Label] Where ID='" & LabelID & "'", Conn, 1, 3
				 LabelRS("LabelName") = LabelName
				 LabelRS("LabelContent") = LabelContent
				 LabelRS("LabelFlag") = LabelFlag
				 LabelRS("FolderID") = FolderID
				 LabelRS("Description") = Descript & "@@@"&FieldParam&"@@@"&SQLType&"@@@"&ItemName&"@@@"&PageStyle&"@@@"&Ajax&"@@@"& datasourcetype &"@@@" &datasourcestr& "@@@" & note
				 LabelRS("AddDate") = Now
				 LabelRS.Update
				 '�������б�ǩ���ݣ��ҳ����б�ǩ��ͼƬ
				 Dim Node,UpFiles,RCls
				 UpFiles=LabelContent
				 LabelRS.Close
				 LabelRS.Open "Select LabelContent From KS_Label Where LabelType=5",conn,1,1
                 Do While Not LabelRS.Eof
				     UpFiles=UpFiles & LabelRS(0)
				     LabelRS.MoveNext
				 Loop
				 LabelRS.Close
				 Set LabelRS=Nothing
				 Call KS.FileAssociation(1021,2,UpFiles,1)


				 '������������
				 
				 If KeyWord = "" Then
					KS.Echo ("<script>alert('�ɹ���ʾ:\n\n��ǩ�޸ĳɹ�!');$(parent.document).find('#BottomFrame')[0].src='" & KS.Setting(3) & KS.Setting(89) & "KS.Split.asp?LabelFolderID=" & FolderID & "&OpStr=��ǩ����  >> �Զ��庯����ǩ&ButtonSymbol=DIYFunctionLabel';location.href='Label_main.asp?Page=" & Page & "&LabelType=5&FolderID=" & FolderID & "';</script>")
				 Else
					KS.Echo ("<script>alert('�ɹ���ʾ:\n\n��ǩ�޸ĳɹ�!');$(parent.document).find('#BottomFrame')[0].src='" & KS.Setting(3) & KS.Setting(89) & "KS.Split.asp?OpStr=��ǩ���� >> <font color=red>�����Զ��庯����ǩ���</font>&ButtonSymbol=DIYFunctionSearch';location.href='Label_main.asp?Page=" & Page & "&LabelType=5&KeyWord=" & KeyWord & "&SearchType=" & SearchType & "&StartDate=" & StartDate & "&EndDate=" & EndDate & "';</script>")
				 End If
			End If
	End Sub
	
	Sub SqlValid(SqlStr)
	     On Error Resume Next
		 if datasourcetype<>0 then
		 tConn.Execute(ClearParam(SqlStr))
		 else
		 Conn.Execute(ClearParam(SqlStr))
		 end if
		 If Err Then 
		  KS.Echo "<script>alert('" & replace(err.description,"'","\'") & "');history.back();</script>"
		  response.end
		 End If
	End Sub
	function ClearParam(byval SqlStr)
	     Dim I
		 For I=0 To 100
		  SqlStr=Replace(SqlStr,"{$Param(" & I & ")}",1)
		 Next
		  SqlStr=Replace(SqlStr,"{$CurrClassChildID}","'1'")
		  SqlStr=Replace(SqlStr,"{$CurrChannelID}",1)
		  SqlStr=Replace(SqlStr,"{$CurrClassID}",1)
		  SqlStr=Replace(SqlStr,"{$CurrInfoID}",1)
		  SqlStr=Replace(SqlStr,"{$CurrSpecialID}",1)
		  SqlStr=Replace(SqlStr,"{$GetUserName}",1)
		  ClearParam=ReplaceRequest(SqlStr)
		 exit function
     End function
	 
'�滻request��ֵ,֧��ReqNum��ReqStr������ǩ
		Function ReplaceRequest(Content)
		     Dim regEx, Matches, Match,TempStr,QStr,ReqType
			 Set regEx = New RegExp
			 regEx.Pattern= "{(ReqNum|ReqStr)[^{}]*}"
			 regEx.IgnoreCase = True
			 regEx.Global = True
			 Set Matches = regEx.Execute(Content)
			 For Each Match In Matches
				On Error Resume Next
				TempStr = Match.Value
				ReqType=Split(TempStr,"(")(0)
				QStr=Replace(Split(TempStr,"(")(1),")}","")
				If ReqType="{ReqNum" Then
				 Content=Replace(Content,TempStr,1)
				Else
				 Content=Replace(Content,TempStr,"1")
				End If
			Next
			ReplaceRequest=Content
		End Function

  
 Function GetLoopTimes(SqlStr)
		 Dim regEx, Matches, Match
		 Set regEx = New RegExp
		 regEx.Pattern = "top\s?[\d]*\d"
		 regEx.IgnoreCase = True
		 regEx.Global = True
		 Set Matches = regEx.Execute(SqlStr)
		 If Matches.count > 0 Then 
		  GetLoopTimes=Trim(Split(lcase(Matches.item(0)),"top")(1))
         End If
		 regEx.Pattern = "top\s?{\$Param\([^}]*}"
		 'regEx.Pattern = "top[^}]*}"
		 regEx.IgnoreCase = True
		 regEx.Global = True
		 Set Matches = regEx.Execute(SqlStr)
		 If Matches.count > 0 Then 
		  GetLoopTimes=Trim(Split(Matches.item(0),"top")(1))
         End If
		If GetLoopTimes="" Then GetLoopTimes=10
  End Function
  

	
	Sub OpenExtConn()
		if datasourcetype<>0 then 
		   '�ⲿaccess�Զ�ת�����·��Ϊ����·��
		   Dim connstr:connstr=datasourcestr
		   if datasourcetype=1 or datasourcetype=5 or datasourcetype=6 Then connstr=LFCls.GetAbsolutePath(connstr)  
		   if not isobject(tconn) then
			on error resume next
		    Set tconn = Server.CreateObject("ADODB.Connection")
			tconn.open connstr
			If Err Then 
			  Err.Clear
			  Set tconn = Nothing
			  KS.Echo "<script>alert('�ⲿ���ݿ�����ʧ��!');history.back();</script>"
			  response.end
			end if
		   end if
		end if
	End Sub
	
	Function ReplaceBC(ByVal C)
	 C=Replace(C,"'","")
	 C=Replace(C,"(","")
	 C=Replace(C,")","")
	 ReplaceBC=C
	End Function
	
	Sub ShowClassID()
	%>
		 <script type="text/javascript">
						  function copyToClipboard(txt) {
							 if(window.clipboardData) {
									 window.clipboardData.clearData();
									 window.clipboardData.setData("Text", txt);
							 } else if(navigator.userAgent.indexOf("Opera") != -1) {
								  window.location = txt;
							 } else if (window.netscape) {
								  try {
									   netscape.security.PrivilegeManager.enablePrivilege("UniversalXPConnect");
								  } catch (e) {
									   alert("��������ܾ���\n�����������ַ������'about:config'���س�\nȻ��'signed.applets.codebase_principal_support'����Ϊ'true'");
								  }
								  var clip = Components.classes['@mozilla.org/widget/clipboard;1'].createInstance(Components.interfaces.nsIClipboard);
								  if (!clip)
									   return;
								  var trans = Components.classes['@mozilla.org/widget/transferable;1'].createInstance(Components.interfaces.nsITransferable);
								  if (!trans)
									   return;
								  trans.addDataFlavor('text/unicode');
								  var str = new Object();
								  var len = new Object();
								  var str = Components.classes["@mozilla.org/supports-string;1"].createInstance(Components.interfaces.nsISupportsString);
								  var copytext = txt;
								  str.data = copytext;
								  trans.setTransferData("text/unicode",str,copytext.length*2);
								  var clipid = Components.interfaces.nsIClipboard;
								  if (!clip)
									   return false;
								  clip.setData(trans,null,clipid.kGlobalClipboard);
							 }
								  alert("���Ƴɹ���")
						}
		 </script>
	 <body class="tdbg">
	 <table width="100%" cellpadding="0" cellspacing="0">
	   <tr><td colspan="4" align="center" height="25" class="title"><strong>(�� Ŀ <=> ID)�� �� ��</strong></td></tr>
	   <tr class="tdbg">
		<td colspan=4>
		  <table border=0>
		   <tr>
		   <td width="30"></td>
		   <td><%
		   GetClassIDTable()
		   %></td>
		   </tr>
		   </table>
		</td>
	   </tr>
	 <table>
	 </body>
	<%
	End Sub
	
  Function GetClassIDTable()
  
		Dim Node,K,SQL,NodeText,Pstr,TJ,SpaceStr
		KS.LoadClassConfig()
		For Each Node In Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks14=1]")
		      SpaceStr=""
			  TJ=Node.SelectSingleNode("@ks10").text
			  If TJ>1 Then
				 For k = 1 To TJ - 1
					SpaceStr = SpaceStr & "����"
				 Next
				KS.Echo "<li>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & SpaceStr & Node.SelectSingleNode("@ks1").text & "&nbsp;&nbsp;&nbsp;"  & Node.SelectSingleNode("@ks0").text & " <input type='button' value='����' class='button' onclick=""copyToClipboard('"&Node.SelectSingleNode("@ks0").text&"')""></li>"
			  Else
				KS.Echo "<li><img src='../images/folder/domain.gif' align='absmiddle'>" & Node.SelectSingleNode("@ks1").text & "&nbsp;&nbsp;&nbsp;&nbsp;" & Node.SelectSingleNode("@ks0").text & " <input type='button' value='����' class='button' onclick=""copyToClipboard('"&Node.SelectSingleNode("@ks0").text&"')""></li>"
			  End If
		Next
	End Function
End Class
%> 
