<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../KS_Cls/Kesion.EscapeCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%

'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Down_Param
KSCls.Kesion()
Set KSCls = Nothing

Class Down_Param
        Private KS,KSCls,ChannelID,IConnStr,Iconn,tempField
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		 Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub
		Sub Kesion()
		 If KS.S("Action")="testsource" Then
		   Call testsource()
		   Exit Sub
		 End If
		 With KS
			.echo "<html>"
			.echo "<title>���ػ�����������</title>"
			.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			.echo "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.echo "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
			.echo "<script src=""../KS_Inc/jquery.js"" language=""JavaScript""></script>"
           %>
		    <script type="text/javascript">
			function datachanage(){
				  switch (parseInt($('#datasourcetype').val()))
					{
					 case 1:
					  $('#datasourcestr').val('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=/���ݿ�.mdb');
					  break;
					 case 3:
					  $('#datasourcestr').val('Provider=Sqloledb; User ID=�û���; Password=����; Initial Catalog=���ݿ�����; Data Source =(local);');
					  break;
					 case 2:
					  $('#datasourcestr').val('driver={microsoft excel driver (*.xls)};dbq=/���ݿ�.xls');
					  break;
					}
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
			  var url = 'KS.Import.asp';
			  $.get(url,{action:"testsource",datatype:datatype,str:escape(str)},function(d){
				if (d=='true')
				 alert('��ϲ������ͨ��!')
				else
				 alert('�Բ����ַ�����������!');
			  });
		    } 
			function checkNext()
			{
			  if ($("#channelid>option:selected").val()==0){
			     alert('��ѡ��Ҫ�����ģ��!');
				 return false;
			  }
			  if ($("#datasourcestr").val()=='')
			  {
			    alert('����������Դ�����ִ�!');
				$("#datasourcestr").focus();
				return false;
			  }
			  if ($("#tablename").val()=='')
			  {
			    alert('���������ݱ���!');
				$("#tablename").focus();
				return false;
			  }
			   return true;
			}
		   function getClass(v){
		      if (v==1){
			   $("#stid1").show();
			   $("#stid2").hide();
			  }else{
			   $("#stid1").hide();
			   $("#stid2").show();
			  }
		   }
		   function getTemplate(v){
		      if (v==1){
			   $("#stemplate1").show();
			   $("#stemplate2").hide();
			  }else{
			   $("#stemplate1").hide();
			   $("#stemplate2").show();
			  }
		   }
		   function getFname(v){
		      if (v==1){
			   $("#sfname").hide();
			  }else{
			   $("#sfname").show();
			  }
		   }
			</script>
		   <%
			.echo "</head>"
			.echo "<body bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"">"
		  
		  select case request("action")
		   case "Step2" Step2
		   case "Step3" Step3
		   case else
		     call step1
		  end select
		  	.echo "</body>"
			.echo "</html>"
        End With
	   End Sub
	   
	   
	   '
	   Sub Step1()
		 With KS
			.echo "      <div class='topdashed sort'>"
			.echo "      ��һ�� ����������������������"
			.echo "      </div>"
			.echo "<form action=""?Action=Step2"" method=""post"" name=""DownParamForm"">"
			.echo "  <table width=""100%"" border=""0"" align=""center"" cellspacing=""1"" class=""ctable"">"
			.echo "    <tr class='tdbg'>"
			.echo "      <td width=""150"" height=""30"" class='clefttitle' align='right'><strong>Ҫ�����ģ��</strong></td>"
			.echo "      <td><select id='channelid' name='channelid'>"
			.echo " <option value='0'>---��ѡ��Ŀ��ģ��---</option>"
	
			If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
			Dim ModelXML,Node
			Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
			For Each Node In ModelXML.documentElement.SelectNodes("channel[@ks21=1 and (@ks6=3 or @ks6=1)]")
			   .echo "<option value='" &Node.SelectSingleNode("@ks0").text &"'>" & Node.SelectSingleNode("@ks1").text & "(" & Node.SelectSingleNode("@ks2").text & ")</option>"
			next
			.echo "</select>"			
			.echo "     </td>"
			.echo "    </tr>"
			.echo "    <tr class='tdbg'>"
			.echo "      <td width=""150"" height=""30"" class='clefttitle' align='right'><strong>����Դ����</strong></td>"
            .echo "      <td><select name=""datasourcetype"" id=""datasourcetype"" onchange=""datachanage()""><option value='1'>access</option><option value='2'>Excel</option><option value='3'>MS SQL</option></select></td>"
			.echo "    </tr>"
			.echo "    <tr class='tdbg'>"
			.echo "      <td width=""150"" height=""30"" class='clefttitle' align='right'><strong>�����ַ���</strong></td>"
            .echo "      <td><textarea name='datasourcestr' id='datasourcestr' cols='70' rows='3'>Provider=Microsoft.Jet.OLEDB.4.0;Data Source=/���ݿ�.mdb</textarea>"
			.echo "     &nbsp;<input class='button' id='testbutton' name='testbutton' type='button' value='����' onclick='testsource();'><br><font color=green>˵��:Access/Excel����Դ֧�����·��,��Provider=Microsoft.Jet.OLEDB.4.0;Data Source=/1.mdb,��ʾ���Ӹ�Ŀ¼�µ�1.mdb���ݿ�</font></td>"
			.echo "    </tr>"
			.echo "    <tr class='tdbg'>"
			.echo "      <td width=""150"" height=""30"" class='clefttitle' align='right'><strong>���ݱ�����</strong></td>"
            .echo "      <td><input type='text' name='tablename' id='tablename' value='Table1' /></td>"
			.echo "    </tr>"
			.echo "  </table>"
			.echo " <div style='text-align:center;padding:20px'><input type='submit' onclick='return(checkNext())' value=' ��һ�� ' class='button' name='button1'></div>"
			.echo "</form>"
			End With
		End Sub
		
		Sub testsource()
			response.cachecontrol="no-cache"
			response.addHeader "pragma","no-cache"
			response.expires=-1
			response.expiresAbsolute=now-1
			Response.CharSet="gb2312"
		   on error resume next
		   dim str:str=unescape(request("str"))
		   If KS.G("DataType")="1" or KS.G("DataType")="2" Then str=LFCls.GetAbsolutePath(str)
		   dim tconn:Set tconn = Server.CreateObject("ADODB.Connection")
			tconn.open str
			If Err Then 
			  Err.Clear
			  Set tconn = Nothing
			  KS.Echo "false"
			else
			  KS.Echo "true"
			end if
		End Sub
		
		Sub OpenImporIConn()
				   if not isobject(IConn) then
					on error resume next
					Set IConn = Server.CreateObject("ADODB.Connection")
					IConn.open IConnStr
					If Err Then 
					  Err.Clear
					  Set IConn = Nothing
					  Response.Write "<script>alert('����Դ����ʧ��,�������ݿ�����!');history.back();</script>"
					  response.end
					end if
				   end if		
		End Sub
       '**************************************************
		'��������ShowChird
		'��  �ã���ʾָ�����ݱ���ֶ��б�
		'��  ������
		'**************************************************
		Function ShowField(fieldname)
				if request("tablename")="" then
				 response.write "<script>alert('�����Ʊ������룡');history.back();</script>"
				 response.end
				end if
				dim dbname:dbname=request("tablename")
				if tempField="" Then
					dim rs:Set rs=Iconn.OpenSchema(4)
					if request("datasourcetype")<>"2" then
					Do Until rs.EOF or rs("Table_name") = trim(dbname)
						rs.MoveNext
					Loop
					end if
					'Do Until rs.EOF or rs("Table_name") <> trim(dbname)
					Do Until rs.EOF
					  tempField=tempField & "<option value='"&lcase(rs("column_Name"))&"'>��"&rs("column_Name")&"</option>"
					  rs.MoveNext
					loop
				    rs.close:set rs=nothing
			   End If
			   ShowField=replace(tempField,"value='" & lcase(fieldname) & "'","value='" & lcase(fieldname) & "' selected")
		End Function	
		
		
		Sub Step2()
		   ChannelID=KS.ChkClng(Request("ChannelID"))
		   If ChannelID=0 Then 
		     KS.AlertHintScript "��ѡ��Ҫ�����ģ��!"
		   End If
		   With KS
			.echo "      <div class='topdashed sort'>"
			.echo "      �ڶ��� �������������ֶ�����"
			.echo "      </div>"
			IConnStr=Request("datasourcestr")
			If KS.G("datasourcetype")="1" or KS.G("datasourcetype")="2" Then IConnStr=LFCls.GetAbsolutePath(IConnStr)
			if IConnStr="" Then
			  KS.AlertHintScript "�����������ַ���!"
			End If
			
			
			OpenImporIConn()
			.echo "<table width='100%' style='margin-top:10px' border='0' align='center'  cellspacing='1' class='ctable'>"
			.echo "<form name='myform' id='myform' action='KS.Import.asp?action=Step3' method='post'>"
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='Title'><option value='0'>-�������-</option>"
			.echo ShowField("title")
			.echo "	</select> =>	</td>"
			.echo "	<td>" & KS.C_S(ChannelID,3) & "����(Title)*</td></tr>"
			
			If KS.C_S(ChanneliD,6)=1 Then
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='FullTitle'><option value='0'>-�������-</option>"
			.echo ShowField("fulltitle")
			.echo "	</select> =>	</td>"
			.echo "	<td>��������(FullTitle)*</td></tr>"
			End If
			
			'===================================��ĿID=====================================
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>������Ŀ:</td><td><label><input type='radio' value='1' name='tidtype' onclick=""getClass(1)"" checked/>ֱ�ӵ���ָ������Ŀ</label> <br/><label><input type='radio' onclick=""getClass(2)"" name='tidtype' value='2'>��ȡ����Դ����ĿID</label>"
			.echo "	</td></tr>"
			
			.echo "<tr class='tdbg' id='stid1'><td height='25' align='right' class='clefttitle'></td><td><select size='1' name='tid1' id='tid1' style='width:160px'>"
			.echo " <option value='0'>--��ѡ����Ŀ--</option>"
			.echo KS.LoadClassOption(ChannelID)& " </select> =>��ĿID(Tid)*</td></tr>"
			
			.echo "<tr class='tdbg' id='stid2' style='display:none'><td height='25' align='right' class='clefttitle'></td><td><select name='tid2'><option value='0'>-�������-</option>"
			.echo ShowField("tid")
			.echo "	</select> =>��ĿID(Tid)*</td></tr>"
			'=================================================================================
			
			'==================================ģ��ID=======================================================
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>��ģ��:</td><td><label><input type='radio' value='1' name='templatetype' onclick=""getTemplate(1)"" checked/>ѡ��ģ�岢��</label> <br/><label><input type='radio' onclick=""getTemplate(2)"" name='templatetype' value='2'>��ȡ����Դ��ģ��</label>"
			.echo "	</td></tr>"
			.echo "<tr class='tdbg' id='stemplate1'><td height='25' align='right' class='clefttitle'></td><td><input id='TemplateID' name='TemplateID' readonly size=20 class='textbox' value='{@TemplateDir}/"& KS.C_S(ChannelID,1) & "/����ҳ.html'>&nbsp;" & KSCls.Get_KS_T_C("$('#TemplateID')[0]") &"  =>ģ��(TemplateID)*</td></tr>"
			
			.echo "<tr class='tdbg' id='stemplate2' style='display:none'><td height='25' align='right' class='clefttitle'></td><td><select name='templateid2'><option value='0'>-�������-</option>"
			.echo ShowField("templateid")
			.echo "	</select> =>ģ��(TemplateID)*</td></tr>"
			'=================================================================================================
			
			'==================================�ļ���=======================================================
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>�ļ���:</td><td><label><input type='radio' value='1' name='Fnametype' onclick=""getFname(1)"" checked/>�Զ�����</label> <br/><label><input type='radio' onclick=""getFname(2)"" name='Fnametype' value='2'>��ȡ����Դ���ļ���</label>"
			.echo "	</td></tr>"
			.echo "<tr class='tdbg' id='sfname' style='display:none'><td height='25' align='right' class='clefttitle'></td><td><select name='Fname'><option value='0'>-�������-</option>"
			.echo ShowField("Fname")
			.echo "	</select> =>�ļ���(Fname)*</td></tr>"
			'=================================================================================================
			
			If KS.C_S(ChanneliD,6)=3 Then
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='downlb'><option value='0'>-�������-</option>"
			.echo ShowField("downlb")
			.echo "	</select> =>	</td>"
			.echo "	<td>" &KS.C_S(ChannelID,3) & "���(DownLB)</td></tr>"
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='downyy'><option value='0'>-�������-</option>"
			.echo ShowField("downyy")
			.echo "	</select> =>	</td>"
			.echo "	<td>" &KS.C_S(ChannelID,3) & "����(DownYY)</td></tr>"
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='downsq'><option value='0'>-�������-</option>"
			.echo ShowField("downsq")
			.echo "	</select> =>	</td>"
			.echo "	<td>" &KS.C_S(ChannelID,3) & "��Ȩ(DownSQ)</td></tr>"
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='downsize'><option value='0'>-�������-</option>"
			.echo ShowField("downsize")
			.echo "	</select> =>	</td>"
			.echo "	<td>" &KS.C_S(ChannelID,3) & "��С(DownSize)</td></tr>"
			
			
			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='downpt'><option value='0'>-�������-</option>"
			.echo ShowField("downpt")
			.echo "	</select> =>	</td>"
			.echo "	<td>ϵͳƽ̨(DownPT)</td></tr>"
			
			End If
			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='photourl'><option value='0'>-�������-</option>"
			.echo ShowField("photourl")
			.echo "	</select> =>	</td>"
			.echo "	<td>" &KS.C_S(ChannelID,3) & "ͼƬ(PhotoUrl)</td></tr>"
			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='keywords'><option value='0'>-�������-</option>"
			.echo ShowField("keywords")
			.echo "	</select> =>	</td>"
			.echo "	<td>�ؼ���(KeyWords)</td></tr>"
			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='author'><option value='0'>-�������-</option>"
			.echo ShowField("author")
			.echo "	</select> =>	</td>"
			If KS.C_S(ChannelID,6)=3 Then
			.echo "	<td>���߿�����(Author)</td></tr>"
			Else
			.echo " <td>" &KS.C_S(ChannelID,3) & "����(Author)</td></tr>"
			End If
			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='Origin'><option value='0'>-�������-</option>"
			.echo ShowField("Origin")
			.echo "	</select> =>	</td>"
			.echo "	<td>" &KS.C_S(ChannelID,3) & "��Դ(Origin)</td></tr>"
		If KS.C_S(ChannelID,6)=3 Then
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='downurls'><option value='0'>-�������-</option>"
			.echo ShowField("downurls")
			.echo "	</select> =>	</td>"
			.echo "	<td>���ص�ַ(DownUrls)</td></tr>"
			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='downcontent'><option value='0'>-�������-</option>"
			.echo ShowField("downcontent")
			.echo "	</select> =>	</td>"
			.echo "	<td>�������(DownContent)</td></tr>"
			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='ysdz'><option value='0'>-�������-</option>"
			.echo ShowField("ysdz")
			.echo "	</select> =>	</td>"
			.echo "	<td>��ʾ��ַ(YSDZ)</td></tr>"
			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='zcdz'><option value='0'>-�������-</option>"
			.echo ShowField("zcdz")
			.echo "	</select> =>	</td>"
			.echo "	<td>ע���ַ(ZCDZ)</td></tr>"
			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='jymm'><option value='0'>-�������-</option>"
			.echo ShowField("JYMM")
			.echo "	</select> =>	</td>"
			.echo "	<td>��ѹ����(JYMM)</td></tr>"
		  ElseIf KS.C_S(ChannelID,6)=1 Then
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='intro'><option value='0'>-�������-</option>"
			.echo ShowField("intro")
			.echo "	</select> =>	</td>"
			.echo "	<td>" & KS.C_S(ChannelID,3) & "���(Intro)</td></tr>"
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='articlecontent'><option value='0'>-�������-</option>"
			.echo ShowField("articlecontent")
			.echo "	</select> =>	</td>"
			.echo "	<td>" & KS.C_S(ChannelID,3) & "����(ArticleContent)</td></tr>"
		  End If			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='adddate'><option value='0'>-�������-</option>"
			.echo ShowField("adddate")
			.echo "	</select> =>	</td>"
			.echo "	<td>�������(AddDate)</td></tr>"
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='inputer'><option value='0'>-�������-</option>"
			.echo ShowField("inputer")
			.echo "	</select> =>	</td>"
			.echo "	<td>" & KS.C_S(ChannelID,3) & "¼��(Inputer)</td></tr>"
			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='rank'><option value='0'>-�������-</option>"
			.echo ShowField("rank")
			.echo "	</select> =>	</td>"
			.echo "	<td>" & KS.C_S(ChannelID,3) & "�ȼ�(Rank)</td></tr>"
			
			Dim FieldArr,K
			FieldArr=KSCls.Get_KS_D_F_P_Arr(ChannelID,"")
			If IsArray(FieldArr) Then
			  For K=0 TO Ubound(FieldArr,2)
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='" & FieldArr(0,k) & "'><option value='0'>-�������-</option>"
			.echo ShowField(FieldArr(0,k))
			.echo "	</select> =>	</td>"
			.echo "	<td>" & FieldArr(1,k) & "(" & FieldArr(0,k) & ")</td></tr>"
			  Next
			End If
			
			.echo "<tr class='tdbg'><td height='35' colspan='2' class='clefttitle'><strong>˵��:</strong><br/>1.���鰴����ģ�͵������ݱ�ṹ��������Դ���ݿ��ļ���������ģ�Ͷ�Ӧ�����ݱ�ṹ��ο�KS_Article;<br/>2.������Ŀ�����ѡ�񡰶�ȡ����Դ����ĿID������ô��ȷ�����������ĿID���Ѵ��ڵ���Ŀ�����ں�̨����Ŀ������Կ�������Ŀ���� </td></tr>"


			.echo "</table>"
			.echo "<input type='hidden' name='channelid' value='" & channelid & "'/>"
			.echo "<input type='hidden' name='datasourcetype' value='" & ks.g("datasourcetype") & "'/>"
			.echo "<input type='hidden' name='datasourcestr' value='" & ks.g("datasourcestr") & "'/>"
			.echo "<input type='hidden' name='tablename' value='" & ks.g("tablename") & "'/>"
			
			.echo "<div style='padding:10px;text-align:center'><input type='submit' onclick=""return(confirm('��������������ȷ��������ٵ��ȷ�ϣ�'))"" value=' �� һ �� ' class='button'</div>"
			.echo "</form>"
           End With
		End Sub
		
		'������
		Sub Step3()
		  %>
		  <div class='topdashed sort'>������ ������������ִ��ҳ��</div>
		
		<div style="text-align:center">			 
			 <div style="margin-top:50px;border:1px dashed #cccccc;width:500px;height:80px">
			 <br>
			<div id="message">
			  <br>������ʾ����
			</div>
			</div>
	    </div>
		<br/><br/><br/>
		  <%
		   ChannelID=KS.ChkClng(Request("ChannelID"))
		   If ChannelID=0 Then 
		     KS.AlertHintScript "��ѡ��Ҫ�����ģ��!"
		   End If

		  IF KS.G("Title")="0" Then
		   Call KS.AlertHintscript("�������ѡ�����ѡ��")
		   response.end
		  End If
		  
		  If KS.G("tidtype")="2" And KS.G("Tid2")="0" Then
		   Call KS.AlertHistory("��Ŀѡ�����ѡ��",-1)
		   response.end
		  ElseIf KS.G("Tidtype")="1" and KS.G("Tid1")="0" Then
		   Call KS.AlertHistory("������Ŀ����ѡ��",-1)
		   response.end
		  End If
		  
		  If KS.G("templatetype")="2" And KS.G("templateid2")="0" Then
		   Call KS.AlertHistory("ģ��ѡ�����ѡ��",-1)
		   response.end
		  End If
		  
		  Server.ScriptTimeOut=999999
			IConnStr=Request("datasourcestr")
			If KS.G("datasourcetype")="1" or KS.G("datasourcetype")="2" Then IConnStr=LFCls.GetAbsolutePath(IConnStr)
			if IConnStr="" Then
			  KS.AlertHintScript "�����������ַ���!"
			End If
			OpenImporIConn()
		 Dim TableName:TableName=Request("TableName")
		 Dim Total,n,msg,errnum,t,Intro
		 Dim IRS:Set IRS=Server.CreateOBject("ADODB.RECORDSET")
    	 Dim RS:Set RS=Server.CreateObject("ADODB.RecordSet")
		 IRS.Open "Select * From [" & TableName & "]",iConn,3,3
		 Total=IRS.RecordCount
		 n=0:t=0:errnum=0
		 Do While Not IRS.Eof
             t=t+1
          if IRS(KS.G("Title"))<>"" Then
			   
				 Dim Tid,TemplateID
				 If KS.G("tidtype")="1" Then 
				   Tid=KS.G("Tid1")
				 Else
				   Tid=IRS(KS.G("Tid2"))
				 End If
				 If KS.G("templatetype")="1" Then
				    TemplateID=KS.G("TemplateID")
				 Else
				    TemplateID=IRS(KS.G("templateid2"))
				 End If
				 
				 
				 RS.Open "Select top 1 * From [" & KS.C_S(ChannelID,2) & "] Where [Title]='" &IRS(KS.G("Title")) & "' and tid='" & tid & "'",conn,3,3
				 If RS.Eof and RS.Bof Then
				   RS.AddNew
				   RS("Title")=IRS(KS.G("Title"))
				   RS("Tid")=Tid
				   RS("TemplateID")=TemplateID
				   If KS.G("Fnametype")="2" Then
				    RS("Fname")=IRS(KS.G("Fname"))
				   End If
				   
				If KS.C_S(Channelid,6)=1 Then
					   If KS.G("Intro")<>"0" Then
						RS("Intro")=IRS(KS.G("Intro"))
					   End If
					   If KS.G("ArticleContent")<>"0" Then
					    If KS.IsNUL(IRS(KS.G("ArticleContent"))) Then
						 RS("ArticleContent")=" "
						Else
					     RS("ArticleContent")=IRS(KS.G("ArticleContent"))
						End If
					   Else
					     RS("ArticleContent")=" "
					   End If
					   
					   If KS.G("FullTitle")<>"0" Then
					    RS("FullTitle")=IRS(KS.G("FullTitle"))
					   End If
					   Intro=RS("Intro")
				ElseIf KS.C_S(Channelid,6)=3 Then   '����
				   If KS.G("DownPT")<>"0" Then
				    RS("DownPT")=IRS(KS.G("DownPT"))
				   End If
				   If KS.G("DownUrls")<>"0" Then
				    RS("DownUrls")=IRS(KS.G("DownUrls"))
				   End If
				   If KS.G("DownContent")<>"0" Then
				    RS("DownContent")=IRS(KS.G("DownContent"))
				   Else
				    RS("DownContent")=" "
				   End If
				   If KS.G("YSDZ")<>"0" Then
				    RS("YSDZ")=IRS(KS.G("YSDZ"))
				   End If
				   If KS.G("DownLB")<>"0" Then
				    RS("DownLB")=IRS(KS.G("DownLB"))
				   End If
				   If KS.G("DownYY")<>"0" Then
				    RS("DownYY")=IRS(KS.G("DownYY"))
				   End If
				   If KS.G("DownSQ")<>"0" Then
				    RS("DownSQ")=IRS(KS.G("DownSQ"))
				   End If
				   If KS.G("DownSize")<>"0" Then
				    RS("DownSize")=IRS(KS.G("DownSize"))
				   End If
				   If KS.G("ZCDZ")<>"0" Then
				    RS("ZCDZ")=IRS(KS.G("ZCDZ"))
				   End If
				   If KS.G("JYMM")<>"0" Then
				    RS("JYMM")=IRS(KS.G("JYMM"))
				   End If
				    Intro=RS("DownContent")
				End If


				   If KS.G("PhotoUrl")<>"0" Then
				    RS("PhotoUrl")=IRS(KS.G("PhotoUrl"))
				   End If
				   If KS.G("KeyWords")<>"0" Then
				    RS("KeyWords")=IRS(KS.G("KeyWords"))
				   End If
				   If KS.G("Author")<>"0" Then
				    RS("Author")=IRS(KS.G("Author"))
				   End If
				   If KS.G("Origin")<>"0" Then
				    RS("Origin")=IRS(KS.G("Origin"))
				   End If

				   If KS.G("Inputer")<>"0" Then
				    RS("inputer")=IRS(KS.G("Inputer"))
				   End If
				   If KS.G("AddDate")<>"0" Then
				    RS("AddDate")=IRS(KS.G("AddDate"))
				   Else
				    RS("AddDate")=Now
				   End If
				   If KS.G("Rank")<>"0" Then
				    RS("Rank")=IRS(KS.G("Rank"))
				   End If
				   
				   
				    Dim FieldArr,K
					FieldArr=KSCls.Get_KS_D_F_P_Arr(ChannelID,"")
					If IsArray(FieldArr) Then
					  For K=0 TO Ubound(FieldArr,2)
					   IF KS.G(FieldArr(0,k))<>"0" Then
					   RS(FieldArr(0,k))=IRS(KS.G(FieldArr(0,k)))
					   End If
					  Next
					End If
				   
				   RS("verific")=1
				   RS.Update
				   RS.MoveLast
				   Dim InfoID:InfoID=RS("ID")
				   If KS.G("Fnametype")="1" Then
				     RS("Fname")=RS("ID") & ".html"
					 RS.Update
				   End If
				   
				   Call LFCls.InserItemInfo(ChannelID,InfoID,RS("Title"),RS("Tid"),Intro,RS("KeyWords"),RS("PhotoUrl"),RS("Inputer"),RS("Verific"),RS("Fname"))

				   N=N+1
				Else
				 msg=msg & "����:" & IRS(KS.G("Title")) & "<br/>"
				 ErrNum=ErrNum+1
				End If
				RS.Close
		    'Else
			'   ErrNum=ErrNum+1
			'End If
		 End If
		  	Response.Write "<script>document.all.message.innerHTML='<br>��<font color=red>" & Total & "</font> �����ݣ����ڵ����<font color=red>" & n & "</font>������������<font color=blue>" & ErrNum & "</font>��!';</script>"
			Response.Flush

		  IRS.MoveNext
		  If t>=Total Then Exit Do
		 Loop
		 IRS.Close:Set IRS=Nothing:Set RS=Nothing
		 Response.Write "<script>document.all.message.innerHTML='<br>��ϲ���ɹ����� <font color=red>" & N & "</font> �����ݣ����� <font color=blue>" & errnum &"</font> ��';</script>"
		 
		 if msg<>"" then
		   response.write "<strong>���¼�¼�ظ�û���ٵ���:</strong><br/><font color=red>" & msg & "</font>"
		 end if
		End Sub

End Class
%> 
