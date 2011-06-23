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
			.echo "<title>下载基本参数设置</title>"
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
					  $('#datasourcestr').val('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=/数据库.mdb');
					  break;
					 case 3:
					  $('#datasourcestr').val('Provider=Sqloledb; User ID=用户名; Password=密码; Initial Catalog=数据库名称; Data Source =(local);');
					  break;
					 case 2:
					  $('#datasourcestr').val('driver={microsoft excel driver (*.xls)};dbq=/数据库.xls');
					  break;
					}
			}
			function testsource()
		    {
			  var str = $('#datasourcestr').val();
			  var datatype=$('#datasourcetype').val();
			  if (str=='')
			  {
				alert('请输入连接字符串!');
				$('#datasourcestr').focus();
				return false;
			  }
			  var url = 'KS.Import.asp';
			  $.get(url,{action:"testsource",datatype:datatype,str:escape(str)},function(d){
				if (d=='true')
				 alert('恭喜，测试通过!')
				else
				 alert('对不起，字符串连接有误!');
			  });
		    } 
			function checkNext()
			{
			  if ($("#channelid>option:selected").val()==0){
			     alert('请选择要导入的模型!');
				 return false;
			  }
			  if ($("#datasourcestr").val()=='')
			  {
			    alert('请输入数据源连接字串!');
				$("#datasourcestr").focus();
				return false;
			  }
			  if ($("#tablename").val()=='')
			  {
			    alert('请输入数据表名!');
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
			.echo "      第一步 数据批量导入主数据设置"
			.echo "      </div>"
			.echo "<form action=""?Action=Step2"" method=""post"" name=""DownParamForm"">"
			.echo "  <table width=""100%"" border=""0"" align=""center"" cellspacing=""1"" class=""ctable"">"
			.echo "    <tr class='tdbg'>"
			.echo "      <td width=""150"" height=""30"" class='clefttitle' align='right'><strong>要导入的模型</strong></td>"
			.echo "      <td><select id='channelid' name='channelid'>"
			.echo " <option value='0'>---请选择目标模型---</option>"
	
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
			.echo "      <td width=""150"" height=""30"" class='clefttitle' align='right'><strong>数据源类型</strong></td>"
            .echo "      <td><select name=""datasourcetype"" id=""datasourcetype"" onchange=""datachanage()""><option value='1'>access</option><option value='2'>Excel</option><option value='3'>MS SQL</option></select></td>"
			.echo "    </tr>"
			.echo "    <tr class='tdbg'>"
			.echo "      <td width=""150"" height=""30"" class='clefttitle' align='right'><strong>连接字符串</strong></td>"
            .echo "      <td><textarea name='datasourcestr' id='datasourcestr' cols='70' rows='3'>Provider=Microsoft.Jet.OLEDB.4.0;Data Source=/数据库.mdb</textarea>"
			.echo "     &nbsp;<input class='button' id='testbutton' name='testbutton' type='button' value='测试' onclick='testsource();'><br><font color=green>说明:Access/Excel数据源支持相对路径,如Provider=Microsoft.Jet.OLEDB.4.0;Data Source=/1.mdb,表示连接根目录下的1.mdb数据库</font></td>"
			.echo "    </tr>"
			.echo "    <tr class='tdbg'>"
			.echo "      <td width=""150"" height=""30"" class='clefttitle' align='right'><strong>数据表名称</strong></td>"
            .echo "      <td><input type='text' name='tablename' id='tablename' value='Table1' /></td>"
			.echo "    </tr>"
			.echo "  </table>"
			.echo " <div style='text-align:center;padding:20px'><input type='submit' onclick='return(checkNext())' value=' 下一步 ' class='button' name='button1'></div>"
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
					  Response.Write "<script>alert('数据源连接失败,请检查数据库连接!');history.back();</script>"
					  response.end
					end if
				   end if		
		End Sub
       '**************************************************
		'过程名：ShowChird
		'作  用：显示指定数据表的字段列表
		'参  数：无
		'**************************************************
		Function ShowField(fieldname)
				if request("tablename")="" then
				 response.write "<script>alert('表名称必须输入！');history.back();</script>"
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
					  tempField=tempField & "<option value='"&lcase(rs("column_Name"))&"'>·"&rs("column_Name")&"</option>"
					  rs.MoveNext
					loop
				    rs.close:set rs=nothing
			   End If
			   ShowField=replace(tempField,"value='" & lcase(fieldname) & "'","value='" & lcase(fieldname) & "' selected")
		End Function	
		
		
		Sub Step2()
		   ChannelID=KS.ChkClng(Request("ChannelID"))
		   If ChannelID=0 Then 
		     KS.AlertHintScript "请选择要导入的模型!"
		   End If
		   With KS
			.echo "      <div class='topdashed sort'>"
			.echo "      第二步 数据批量导入字段设置"
			.echo "      </div>"
			IConnStr=Request("datasourcestr")
			If KS.G("datasourcetype")="1" or KS.G("datasourcetype")="2" Then IConnStr=LFCls.GetAbsolutePath(IConnStr)
			if IConnStr="" Then
			  KS.AlertHintScript "请输入连接字符串!"
			End If
			
			
			OpenImporIConn()
			.echo "<table width='100%' style='margin-top:10px' border='0' align='center'  cellspacing='1' class='ctable'>"
			.echo "<form name='myform' id='myform' action='KS.Import.asp?action=Step3' method='post'>"
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='Title'><option value='0'>-此项不导入-</option>"
			.echo ShowField("title")
			.echo "	</select> =>	</td>"
			.echo "	<td>" & KS.C_S(ChannelID,3) & "标题(Title)*</td></tr>"
			
			If KS.C_S(ChanneliD,6)=1 Then
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='FullTitle'><option value='0'>-此项不导入-</option>"
			.echo ShowField("fulltitle")
			.echo "	</select> =>	</td>"
			.echo "	<td>完整标题(FullTitle)*</td></tr>"
			End If
			
			'===================================栏目ID=====================================
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>所属栏目:</td><td><label><input type='radio' value='1' name='tidtype' onclick=""getClass(1)"" checked/>直接导入指定的栏目</label> <br/><label><input type='radio' onclick=""getClass(2)"" name='tidtype' value='2'>读取数据源的栏目ID</label>"
			.echo "	</td></tr>"
			
			.echo "<tr class='tdbg' id='stid1'><td height='25' align='right' class='clefttitle'></td><td><select size='1' name='tid1' id='tid1' style='width:160px'>"
			.echo " <option value='0'>--请选择栏目--</option>"
			.echo KS.LoadClassOption(ChannelID)& " </select> =>栏目ID(Tid)*</td></tr>"
			
			.echo "<tr class='tdbg' id='stid2' style='display:none'><td height='25' align='right' class='clefttitle'></td><td><select name='tid2'><option value='0'>-此项不导入-</option>"
			.echo ShowField("tid")
			.echo "	</select> =>栏目ID(Tid)*</td></tr>"
			'=================================================================================
			
			'==================================模板ID=======================================================
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>绑定模板:</td><td><label><input type='radio' value='1' name='templatetype' onclick=""getTemplate(1)"" checked/>选择模板并绑定</label> <br/><label><input type='radio' onclick=""getTemplate(2)"" name='templatetype' value='2'>读取数据源的模板</label>"
			.echo "	</td></tr>"
			.echo "<tr class='tdbg' id='stemplate1'><td height='25' align='right' class='clefttitle'></td><td><input id='TemplateID' name='TemplateID' readonly size=20 class='textbox' value='{@TemplateDir}/"& KS.C_S(ChannelID,1) & "/内容页.html'>&nbsp;" & KSCls.Get_KS_T_C("$('#TemplateID')[0]") &"  =>模板(TemplateID)*</td></tr>"
			
			.echo "<tr class='tdbg' id='stemplate2' style='display:none'><td height='25' align='right' class='clefttitle'></td><td><select name='templateid2'><option value='0'>-此项不导入-</option>"
			.echo ShowField("templateid")
			.echo "	</select> =>模板(TemplateID)*</td></tr>"
			'=================================================================================================
			
			'==================================文件名=======================================================
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>文件名:</td><td><label><input type='radio' value='1' name='Fnametype' onclick=""getFname(1)"" checked/>自动生成</label> <br/><label><input type='radio' onclick=""getFname(2)"" name='Fnametype' value='2'>读取数据源的文件名</label>"
			.echo "	</td></tr>"
			.echo "<tr class='tdbg' id='sfname' style='display:none'><td height='25' align='right' class='clefttitle'></td><td><select name='Fname'><option value='0'>-此项不导入-</option>"
			.echo ShowField("Fname")
			.echo "	</select> =>文件名(Fname)*</td></tr>"
			'=================================================================================================
			
			If KS.C_S(ChanneliD,6)=3 Then
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='downlb'><option value='0'>-此项不导入-</option>"
			.echo ShowField("downlb")
			.echo "	</select> =>	</td>"
			.echo "	<td>" &KS.C_S(ChannelID,3) & "类别(DownLB)</td></tr>"
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='downyy'><option value='0'>-此项不导入-</option>"
			.echo ShowField("downyy")
			.echo "	</select> =>	</td>"
			.echo "	<td>" &KS.C_S(ChannelID,3) & "语言(DownYY)</td></tr>"
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='downsq'><option value='0'>-此项不导入-</option>"
			.echo ShowField("downsq")
			.echo "	</select> =>	</td>"
			.echo "	<td>" &KS.C_S(ChannelID,3) & "授权(DownSQ)</td></tr>"
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='downsize'><option value='0'>-此项不导入-</option>"
			.echo ShowField("downsize")
			.echo "	</select> =>	</td>"
			.echo "	<td>" &KS.C_S(ChannelID,3) & "大小(DownSize)</td></tr>"
			
			
			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='downpt'><option value='0'>-此项不导入-</option>"
			.echo ShowField("downpt")
			.echo "	</select> =>	</td>"
			.echo "	<td>系统平台(DownPT)</td></tr>"
			
			End If
			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='photourl'><option value='0'>-此项不导入-</option>"
			.echo ShowField("photourl")
			.echo "	</select> =>	</td>"
			.echo "	<td>" &KS.C_S(ChannelID,3) & "图片(PhotoUrl)</td></tr>"
			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='keywords'><option value='0'>-此项不导入-</option>"
			.echo ShowField("keywords")
			.echo "	</select> =>	</td>"
			.echo "	<td>关键字(KeyWords)</td></tr>"
			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='author'><option value='0'>-此项不导入-</option>"
			.echo ShowField("author")
			.echo "	</select> =>	</td>"
			If KS.C_S(ChannelID,6)=3 Then
			.echo "	<td>作者开发商(Author)</td></tr>"
			Else
			.echo " <td>" &KS.C_S(ChannelID,3) & "作者(Author)</td></tr>"
			End If
			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='Origin'><option value='0'>-此项不导入-</option>"
			.echo ShowField("Origin")
			.echo "	</select> =>	</td>"
			.echo "	<td>" &KS.C_S(ChannelID,3) & "来源(Origin)</td></tr>"
		If KS.C_S(ChannelID,6)=3 Then
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='downurls'><option value='0'>-此项不导入-</option>"
			.echo ShowField("downurls")
			.echo "	</select> =>	</td>"
			.echo "	<td>下载地址(DownUrls)</td></tr>"
			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='downcontent'><option value='0'>-此项不导入-</option>"
			.echo ShowField("downcontent")
			.echo "	</select> =>	</td>"
			.echo "	<td>软件介绍(DownContent)</td></tr>"
			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='ysdz'><option value='0'>-此项不导入-</option>"
			.echo ShowField("ysdz")
			.echo "	</select> =>	</td>"
			.echo "	<td>演示地址(YSDZ)</td></tr>"
			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='zcdz'><option value='0'>-此项不导入-</option>"
			.echo ShowField("zcdz")
			.echo "	</select> =>	</td>"
			.echo "	<td>注册地址(ZCDZ)</td></tr>"
			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='jymm'><option value='0'>-此项不导入-</option>"
			.echo ShowField("JYMM")
			.echo "	</select> =>	</td>"
			.echo "	<td>解压密码(JYMM)</td></tr>"
		  ElseIf KS.C_S(ChannelID,6)=1 Then
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='intro'><option value='0'>-此项不导入-</option>"
			.echo ShowField("intro")
			.echo "	</select> =>	</td>"
			.echo "	<td>" & KS.C_S(ChannelID,3) & "简介(Intro)</td></tr>"
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='articlecontent'><option value='0'>-此项不导入-</option>"
			.echo ShowField("articlecontent")
			.echo "	</select> =>	</td>"
			.echo "	<td>" & KS.C_S(ChannelID,3) & "内容(ArticleContent)</td></tr>"
		  End If			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='adddate'><option value='0'>-此项不导入-</option>"
			.echo ShowField("adddate")
			.echo "	</select> =>	</td>"
			.echo "	<td>添加日期(AddDate)</td></tr>"
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='inputer'><option value='0'>-此项不导入-</option>"
			.echo ShowField("inputer")
			.echo "	</select> =>	</td>"
			.echo "	<td>" & KS.C_S(ChannelID,3) & "录入(Inputer)</td></tr>"
			
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='rank'><option value='0'>-此项不导入-</option>"
			.echo ShowField("rank")
			.echo "	</select> =>	</td>"
			.echo "	<td>" & KS.C_S(ChannelID,3) & "等级(Rank)</td></tr>"
			
			Dim FieldArr,K
			FieldArr=KSCls.Get_KS_D_F_P_Arr(ChannelID,"")
			If IsArray(FieldArr) Then
			  For K=0 TO Ubound(FieldArr,2)
			.echo "<tr class='tdbg'><td height='25' align='right' class='clefttitle'>"
		    .echo "		<select name='" & FieldArr(0,k) & "'><option value='0'>-此项不导入-</option>"
			.echo ShowField(FieldArr(0,k))
			.echo "	</select> =>	</td>"
			.echo "	<td>" & FieldArr(1,k) & "(" & FieldArr(0,k) & ")</td></tr>"
			  Next
			End If
			
			.echo "<tr class='tdbg'><td height='35' colspan='2' class='clefttitle'><strong>说明:</strong><br/>1.建议按各个模型的主数据表结构制作数据源数据库文件，如文章模型对应的数据表结构请参考KS_Article;<br/>2.所属栏目如果是选择“读取数据源的栏目ID”，那么请确保您整理的栏目ID是已存在的栏目（即在后台的栏目管理可以看到的栏目）； </td></tr>"


			.echo "</table>"
			.echo "<input type='hidden' name='channelid' value='" & channelid & "'/>"
			.echo "<input type='hidden' name='datasourcetype' value='" & ks.g("datasourcetype") & "'/>"
			.echo "<input type='hidden' name='datasourcestr' value='" & ks.g("datasourcestr") & "'/>"
			.echo "<input type='hidden' name='tablename' value='" & ks.g("tablename") & "'/>"
			
			.echo "<div style='padding:10px;text-align:center'><input type='submit' onclick=""return(confirm('请认真检查各导入项，确定无误后再点击确认！'))"" value=' 下 一 步 ' class='button'</div>"
			.echo "</form>"
           End With
		End Sub
		
		'步骤三
		Sub Step3()
		  %>
		  <div class='topdashed sort'>第三步 数据批量导入执行页面</div>
		
		<div style="text-align:center">			 
			 <div style="margin-top:50px;border:1px dashed #cccccc;width:500px;height:80px">
			 <br>
			<div id="message">
			  <br>操作提示栏！
			</div>
			</div>
	    </div>
		<br/><br/><br/>
		  <%
		   ChannelID=KS.ChkClng(Request("ChannelID"))
		   If ChannelID=0 Then 
		     KS.AlertHintScript "请选择要导入的模型!"
		   End If

		  IF KS.G("Title")="0" Then
		   Call KS.AlertHintscript("软件名称选项必须选择")
		   response.end
		  End If
		  
		  If KS.G("tidtype")="2" And KS.G("Tid2")="0" Then
		   Call KS.AlertHistory("栏目选项必须选择",-1)
		   response.end
		  ElseIf KS.G("Tidtype")="1" and KS.G("Tid1")="0" Then
		   Call KS.AlertHistory("所属栏目必须选择",-1)
		   response.end
		  End If
		  
		  If KS.G("templatetype")="2" And KS.G("templateid2")="0" Then
		   Call KS.AlertHistory("模板选项必须选择",-1)
		   response.end
		  End If
		  
		  Server.ScriptTimeOut=999999
			IConnStr=Request("datasourcestr")
			If KS.G("datasourcetype")="1" or KS.G("datasourcetype")="2" Then IConnStr=LFCls.GetAbsolutePath(IConnStr)
			if IConnStr="" Then
			  KS.AlertHintScript "请输入连接字符串!"
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
				ElseIf KS.C_S(Channelid,6)=3 Then   '下载
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
				 msg=msg & "名称:" & IRS(KS.G("Title")) & "<br/>"
				 ErrNum=ErrNum+1
				End If
				RS.Close
		    'Else
			'   ErrNum=ErrNum+1
			'End If
		 End If
		  	Response.Write "<script>document.all.message.innerHTML='<br>共<font color=red>" & Total & "</font> 条数据，正在导入第<font color=red>" & n & "</font>条！出错跳过<font color=blue>" & ErrNum & "</font>条!';</script>"
			Response.Flush

		  IRS.MoveNext
		  If t>=Total Then Exit Do
		 Loop
		 IRS.Close:Set IRS=Nothing:Set RS=Nothing
		 Response.Write "<script>document.all.message.innerHTML='<br>恭喜！成功导入 <font color=red>" & N & "</font> 条数据！出错 <font color=blue>" & errnum &"</font> 条';</script>"
		 
		 if msg<>"" then
		   response.write "<strong>以下记录重复没有再导入:</strong><br/><font color=red>" & msg & "</font>"
		 end if
		End Sub

End Class
%> 
