<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Admin_Picture
KSCls.Kesion()
Set KSCls = Nothing

Class Admin_Picture
        Private KS,KSCls
		'=====================================定义本页面全局变量=====================================
		Private ProID, I, totalPut, Page, RS,ComeFrom,Action
		Private KeyWord, SearchType, StartDate, EndDate, ParentRs, SearchParam,MaxPerPage
		Private CurrPath,PreViewObj, UpPowerFlag,SaveFilePath
		Private ComeUrl,ChannelID,picnum,sqlstr

		'=============================================================================================
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		  Set KSCls=New ManageCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		 Set KSCls=Nothing
		End Sub
		Public Sub Kesion()
		ChannelID=KS.ChkClng(KS.G("ChannelID"))
		'收集搜索参数
		KeyWord   = KS.G("KeyWord")
		SearchType= KS.G("SearchType")
		StartDate = KS.G("StartDate")
		EndDate   = KS.G("EndDate")
		Action     = KS.G("Action")
		ComeFrom   = KS.G("ComeFrom")
		SearchParam = "ChannelID=" & ChannelID
		If KeyWord<>"" Then SearchParam=SearchParam & "&KeyWord=" & KeyWord
		If SearchType<>"" Then  SearchParam=SearchParam & "&SearchType=" & SearchType
		If StartDate<>"" Then SearchParam=SearchParam & "&StartDate=" & StartDate 
		If EndDate<>"" Then SearchParam=SearchParam & "&EndDate=" & EndDate
		If KS.S("Status")<>"" Then SearchParam=SearchParam & "&Status=" & KS.S("Status")
		If ComeFrom<>"" Then SearchParam=SearchParam & "&ComeFrom=" & ComeFrom
		
		ComeUrl=Cstr(Request.ServerVariables("HTTP_REFERER"))

		ProID = KS.G("ProID"):If ProID = "" Then ProID = "0"
		If Action="Del" Then
		 Call DelImages()
		ElseIf Action="DelEmpty" Then
		 Call DelEmpty()
		End If
		
		With Response
			.Write "<html>"
			.Write "<head>"
			.Write "<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>"
			.Write "<title>添加</title>"
			.Write "<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>"
		    .Write "<script language=""JavaScript"" src=""../KS_Inc/common.js""></script>" & vbCrLf
		    .Write "<script language=""JavaScript"" src=""../KS_Inc/Jquery.js""></script>" & vbCrLf
			.Write "</head>"
			.Write "<body leftmargin='0' topmargin='0' marginwidth='0' onkeydown='if (event.keyCode==83 && event.ctrlKey) SubmitFun();' marginheight='0'>"
        End With
		ComeUrl=Request.ServerVariables("HTTP_REFERER")
			IF KS.G("Method")="Save" Then
				 Call DoSave()
			Else 
				 Call PictureAdd()
			End If
		End Sub

        '添加
        Sub PictureAdd() 
			With Response
			CurrPath = KS.GetUpFilesDir()
			Set RS = Server.CreateObject("ADODB.RecordSet")
			'取得上传权限
			UpPowerFlag = KS.ReturnPowerResult(ChannelID, "M" & ChannelID & "10009")
			
			.Write "<div align='center'>"
			.Write "<ul id='menu_top'>"
			.Write "<li onclick=""return(SubmitFun())"" class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/save.gif' border='0' align='absmiddle'>确定保存</span></li>"
			.Write "<li onclick=""location.href='KS.Shop.asp?page=" & page & "&" & SearchParam  & "';"" class='parent'><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/back.gif' border='0' align='absmiddle'>取消返回</span></li>"
		    .Write "</ul>"			
			
            .Write " <TABLE style='margin:1px' width='100%' BORDER='0' cellpadding='1'  cellspacing='1' class='ctable'>"

			.Write "    <form action='?ChannelID=" & ChannelID & "&Method=Save' method='post' name='PictureForm' onsubmit='return(SubmitFun())'>"
			.Write "      <input type='hidden' value='" & ProID & "' name='ProID'>"
			.Write "      <input type='hidden' value='" & Action & "' name='Action'>"
			.Write "      <input type='hidden' name='Page' value='" & Page & "'>"
			.Write "      <input type='hidden' name='KeyWord' value='" & KeyWord & "'>"
			.Write "      <input type='hidden' name='SearchType' value='" & SearchType & "'>"
			.Write "      <Input type='hidden' name='StartDate' value='" & StartDate & "'>"
			.Write "      <input type='hidden' name='EndDate' value='" & EndDate & "'>"
			
			
			
			.Write "<tr class='tdbg'>" 
		    .Write " <td align='right' class='clefttitle'><div align='right'><strong>图片数:</strong></div></td>"
            .Write " <td height='24'>"
			.Write "<input name='picnum' type='text' id='picnum' size='4' value='2' style='text-align:center'> 分组名称<input type='text' name='groupname' id='groupname' size='10'>"  
			.Write " <select name='sgroupname' onchange=""$('#groupname').val(this.value);"">"
			.Write "<option value=''>--选择分组名称--</option>"
			.write KSCls.Get_O_F_D("KS_ProImages","distinct GroupName","1=1")
			.write "</select>"
			.Write "&nbsp;<input name='kkkup' type='button' id='kkkup2' value='设定' onClick=""MakeUpload($('#picnum').val(),'click');"" class='button'>注：最多<font color='red'>99</font>张，远程图片地址必须以<font color='red'>http://</font>开头"
	        .Write "<input type='hidden' name='PicUrls'>"
		    .Write " </td>"
            .Write "</tr>"
            .Write "<tr class='tdbg'>" 
		    .Write " <td width='85' class='CLeftTitle'><div align='right'><b>图片内容:</b><br><input type='checkbox' value='1' name='BeyondSavePic' id='BeyondSavePic' checked>自动存图<br><br><font color=#ff0000>如果想删除某张" & KS.C_S(ChannelID,3) & "，请在" & KS.C_S(ChannelID,3) & "地址里输入 ""del""或留空。</font></div></td>"
            .Write "<td height='24'>"
			.Write "	<span id='uploadfield'></span>"
			.Write "	<iframe id='UpPhotoFrame' name='UpPhotoFrame' src='KS.UpFileForm.asp?UpType=ProImage&ChannelID="& ChannelID & "' frameborder=0 scrolling=no width='100%' height='90'></iframe>"
		    .Write "</td>"
            .Write "</tr>"

			
			 .Write "    </TABLE>"
			 .Write "<div style='margin-top:4px;text-align:center'><input type='button' class='button' onclick='return SubmitFun()' value='保存上传图片'></div>"
			 .Write "</form>"
			 .Write " </div>"
			 
	  		Dim picnum
			 picnum=2
			%>
			 <script>
			 <%if Action<>"Edit" Then%>
			 var LastNum=1;
			 <%else%>
			 var LastNum=$('#picnum').val();
			 <%end if%>
			 var tempup='';
			 var picnum=<%=Picnum%>;
			 
		 	 function document.onreadystatechange()
			  {   
				 MakeUpload(<%=Picnum%>);
				 tempup=$("#uploadfield").html();
			  }
			
			function MakeUpload(mnum,str)
			{ 
			   if (parseInt(mnum)>=100){
			   alert('最多只能同上传99张!');
			   return false;}
			   var startNum=1;
			   var endNum = mnum;
			   var fhtml = "";
			   
			   if (str=='click') startNum=LastNum;
			   
			   for(startNum;startNum <= endNum;startNum++){
				   fhtml += "<table width=\"99%\" style='margin:2px' class='ctable' align=center border=\"0\" id=\"seltb"+startNum+"\" cellpadding=\"3\" cellspacing=\"1\">";
				   fhtml += "<tr class='tdbg'> "
				   fhtml +="  <td height=\"25\" width=20 align=center class=clefttitle rowspan=\"3\"><strong>第"+startNum+"张</strong></td>";
				   fhtml += " <td width=\"124\" rowspan=\"3\" align=\"center\"><div id=\"picview"+startNum+"\" name=\"picview"+startNum+"\" style=\"filter:progid:DXImageTransform.Microsoft.AlphaImageLoader(sizingMethod=scale);height:80px;width:120px;border:1px solid #777777\"><img src=\"images/pview.gif\" width=\"120\" height=\"80\"></div></td>";
				   fhtml += "</tr>";
				   fhtml += "<tr class='tdbg'> ";
				   fhtml += "<td height=\"30\">　分组名称：";
				   fhtml += "<input type=\'text\' name='groupname"+startNum+"' value='"+document.getElementById("groupname").value+"'> ";
				   fhtml += "</td></tr>";
				   fhtml += "<tr class='tdbg'> ";
				   fhtml += "<td height=\"25\"> 　小图地址：";
				   fhtml += "<input type=\"text\" onblur='view("+startNum+");' name='thumb"+startNum+"' size=\"32\" value=\"\"> ";
				   fhtml += "<input type=\"button\" name='selpic"+startNum+"' value=\"选取图片\" onclick=\"OpenThenSetValue('Include/SelectPic.asp?ChannelID=<%=ChannelID%>&CurrPath=<%=CurrPath%>',550,290,window,document.PictureForm.thumb"+startNum+");view("+startNum+");\" class=\"button\">&nbsp;<input type=\"button\" name='tgetpic"+startNum+"' value=\"远程抓图\" onclick=\"OpenThenSetValue('Include/Frame.asp?FileName=SaveBeyondfile.asp&PageTitle=抓取远程图片&ItemName=<%=KS.C_S(ChannelID,3)%>&CurrPath=<%=CurrPath%>',300,100,window,document.PictureForm.thumb"+startNum+");view("+startNum+");\" class=\"button\">";
				   fhtml += "<br>&nbsp;&nbsp;&nbsp;大图地址：<input type=\"text\" onblur='view("+startNum+");' name='imgurl"+startNum+"' size=\"32\" value=\"\"> ";
				   fhtml += "<input type=\"hidden\" name='pimgurl"+startNum+"' value=\"\">";
				   fhtml += "<input type=\"button\" name='selpic"+startNum+"' value=\"选取图片\" onclick=\"OpenThenSetValue('Include/SelectPic.asp?ChannelID=<%=ChannelID%>&CurrPath=<%=CurrPath%>',550,290,window,document.PictureForm.imgurl"+startNum+");view("+startNum+");\" class=\"button\">&nbsp;<input type=\"button\" name='getpic"+startNum+"' value=\"远程抓图\" onclick=\"OpenThenSetValue('Include/Frame.asp?FileName=SaveBeyondfile.asp&PageTitle=抓取远程图片&ItemName=<%=KS.C_S(ChannelID,3)%>&CurrPath=<%=CurrPath%>',300,100,window,document.PictureForm.imgurl"+startNum+");view("+startNum+");\" class=\"button\">";
				    <%If Action="Edit" Then%>
					if (startNum>picnum)
					{
				   fhtml += "<iframe id='UpPhotoFrame"+startNum+"' name='UpPhotoFrame"+startNum+"' src='KS.UpFileForm.asp?ChannelID=<%=ChannelID%>&uptype=Single&objid="+startNum+"' frameborder=0 scrolling=no width='100%' height='22'></iframe>"
				   }
				   <%end if%>
				   
				   fhtml += "</td></tr>";
				   fhtml += "</table>\r\n";
			  }
			  <%If Action="Edit" Then%>
			  //LastNum=Number(endNum)+1;
			  $("#uploadfield").html(tempup+fhtml);
			  <%Else%>
			  $("#uploadfield").html(fhtml);
			  frames['UpPhotoFrame'].ChooseOption(mnum);
			  $('#UpPhotoFrame').height(80+26*(mnum/2)); 
			  <%End If%>
			}
			 function view(num)
			 {
			  if ($("input[name=thumb"+num+"]").val()!=''){
			  $("#picview"+num).html("");
			  $("#picview"+num)[0].filters.item("DXImageTransform.Microsoft.AlphaImageLoader").src=$("input[name=thumb"+num+"]").val();
			  }
			  else if($("input[name=imgurl"+num+"]").val()!=''){
			  $("#picview"+num).html("");
			  $("#picview"+num)[0].filters.item("DXImageTransform.Microsoft.AlphaImageLoader").src=$("input[name=imgurl"+num+"]").val();
			  }
			 }

			 function SetPicUrlByUpLoad(PicUrlStr,ThumbPathFileName)
			{  var UrlStrArr;
			   UrlStrArr=PicUrlStr.split('|');
			   for (var i=1;i<UrlStrArr.length;i++)
			   {
			   var url=UrlStrArr[i-1]; 
			   if(url!=null&&url!=''){
				 $('input[name=imgurl'+i+']').val(url);
			   } 
			  }
			  var ThumbsArr=ThumbPathFileName.split("|")
			  for(var i=1;i<ThumbsArr.length;i++)
			  {
			   var url=ThumbsArr[i-1]; 
			   if(url!=null&&url!=''){
				 $('input[name=thumb'+i+']').val(url);
			   } 
			  }
			  

			}

            var spic=null;
			function SubmitFun()
				{ 
				var PicUrls='';
				for(var i=1;i<=$("#picnum").val();i++){
				  if ($('input[name=imgurl'+i+']').val()!=''&&$('input[name=imgurl'+i+']').val()!='del') 
				   {
				   var groupname=$('input[name=groupname'+i+']').val();
				   spic=$('input[name=imgurl'+i+']').val();
				   tpic=$('input[name=thumb'+i+']').val();
				   if (spic!='' && tpic!='' && groupname=='')
				   {
				    alert('第'+i+'张图片，请输入分组名称!');
					$('input[name=groupname'+i+']').focus();
					return false;
				   }
				   if (tpic=='') tpic=spic;
				   if (spic.substring(0,4).toLowerCase()=='http'&&$("#BeyondSavePic").attr("checked")==true)
				   {
					 $('#LayerPrompt').show();
					 window.setInterval('ShowPromptMessage()',150)
				   }
				   if (PicUrls=='')                 
				   PicUrls=groupname+'|'+spic+'|'+tpic;
				   else 
				   PicUrls+='|||'+groupname+'|'+spic+'|'+tpic;
				  }
				}
				if (PicUrls=='')
				{
				  alert('请先上传图片!');
				  $('input[name=thumb1]').focus();
				  return false;
				}
				$('form[name=PictureForm]').submit();
				return true;
			}

		var ForwardShow=true;
		function ShowPromptMessage()
			{
				var TempStr=ShowArticleArea.innerText;
				if (ForwardShow==true)
				{
					if (TempStr.length>4) ForwardShow=false;
					ShowArticleArea.innerText=TempStr+'.';
					
				}
				else
				{
					if (TempStr.length==1) ForwardShow=true;
					ShowArticleArea.innerText=TempStr.substr(0,TempStr.length-1);
				}
			}
		var SaveBeyondInfo=''
					   +'<div id="LayerPrompt" style="position:absolute;left: 200px; top: 200px; background-color: #f1efd9; layer-background-color: #f1efd9; border: 1px none #000000; width: 360px; height: 63px; display: none; "> '
					   +'<table width="100%" height="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#FF0000">'
					   +'<tr> '
					   +'<td align="center">'
					   +'<table width="80%" border="0" cellspacing="0" cellpadding="0">'
					   +'<tr>'
					   +' <td width="75%" nowrap>'
					   +'<div align="right">请稍候，系统正在保存远程图片到本地</div></td>'
					   +'   <td width="25%"><font id="ShowArticleArea">&nbsp;</font></td>'
					   +' </tr>'
					   +'</table>'
					   +'</td>'
					   +'</tr>'
					   +'</table>'
					   +'</div>'
			document.write (SaveBeyondInfo)
		</script>
		<script>
		function CheckAll(form) {  
		  $("input[type=checkbox]").each(function(){
		    if ($(this).attr("name")!="BeyondSavePic"){
		    $(this).attr("checked",true);
			}
		  });
		} 
		function ContraSel(form) {
		   $("input[type=checkbox]").each(function(){
		    if ($(this).attr("name")!="BeyondSavePic"){
			$(this).attr("checked",!($(this).attr("checked")));
			}
		  });
		}
		</script>
    
<%
             Call ImageList()
			 .Write "</body>"
			 .Write "</html>"
			 End With
		End Sub
		
		
		Sub ImageList()
		    on error resume next
			If KS.G("page") <> "" Then
				  Page = KS.ChkClng(KS.G("page"))
			Else
				  Page = 1
			End If
			MaxPerPage=12
		  With Response
		     .Write " <TABLE style='margin:1px' width='100%' BORDER='0' cellpadding='1' bgcolor='#cccccc' cellspacing='1' >"
			 .Write "<tr><td height='28' colspan='4' class='clefttitle'>"
			 .Write " <table border='0' width='100%'><tr><td>以下是关于商品 <font color=red>""" & Conn.Execute("Select Title From KS_Product Where ID=" & ProID)(0) & """</font> 的图片</td><td align='right'>"
			 .Write "<select name='showgrouplist' onchange=""location.href='?proid=" &proid & "&" & searchparam & "&groupname='+this.value;"">"
			.Write "<option value=''>--按图片组名称查看--</option>"
			.write  Replace(KSCls.Get_O_F_D("KS_ProImages","distinct GroupName","proid=" & proid),"value=""" & ks.g("groupname") &"""","value=""" & ks.g("groupname") & """ selected")
			 .Write "</select></td></tr></table>"
			 .Write "</td></tr>"
			 Dim Param:Param=" where ProID=" & ProID
			 if KS.G("GroupName")<>"" then Param=Param & " and groupname='" & KS.G("GroupName") & "'"
			 Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
			 RS.Open "Select id,proid,smallpicurl,bigpicurl,groupname From KS_ProImages" & Param & " Order By ID Desc",conn,1,1
			 If Not RS.EOF Then
					totalPut = Conn.Execute("Select count(id) from KS_ProImages" & Param)(0)
		
							If Page < 1 Then Page = 1
							If (Page - 1) * MaxPerPage > totalPut Then
								If (totalPut Mod MaxPerPage) = 0 Then
									Page = totalPut \ MaxPerPage
								Else
									Page = totalPut \ MaxPerPage + 1
								End If
							End If
		
							If Page > 1 Then
								If (Page - 1) * MaxPerPage < totalPut Then
									RS.Move (Page - 1) * MaxPerPage
								Else
									Page = 1
								End If
							End If
							Dim SQL:SQL=RS.GetRows(MaxPerPage)
							Call showContent(SQL)
			 Else
			  .Write "<tr><td colspan=4 align='center' height='35' bgcolor='#ffffff' style='border-bottom:1px solid #cccccc'><font color=red>对不起，没有找到任何图片!</font></td></tr>"
			 End If
			 .Write "</table>"
          End With
		End Sub
		
		Sub showContent(SQL)
		  Dim k:K=1
		  With Response
		    .Write "<form action='KS.ProImages.asp' method='get' name='myform'>"
			.Write "      <input type='hidden' value='" & ProID & "' name='ProID'>"
			.Write "      <input type='hidden' value='Del' name='action'>"
			.Write "      <input type='hidden' name='Page' value='" & Page & "'>"
			.Write "      <input type='hidden' name='KeyWord' value='" & KeyWord & "'>"
			.Write "      <input type='hidden' name='SearchType' value='" & SearchType & "'>"
			.Write "      <Input type='hidden' name='StartDate' value='" & StartDate & "'>"
			.Write "      <input type='hidden' name='EndDate' value='" & EndDate & "'>"
		    .Write "<tr>"&vbcrlf
		  For I=0 To Ubound(SQL,2)
		    .Write "<td bgcolor='#ffffff' width='25%' align='center'><a href='" & sql(3,i) & "' target='_blank'><img width='100' height='100' src='" & SQL(2,i) & "' border='0'></a><br><strong>分组名称：</strong>" & sql(4,i) & "<br>管理操作：<input type='checkbox' name='id' value='" & sql(0,i) & "'> <a href='?proid=" & proid& "&action=Del&id=" &sql(0,i) &"&"&SearchParam & "' onclick='return(confirm(""确定删除吗？""));'>×删除</a></td>" & vbcrlf
			If (i+1) Mod 4 =0 then 
			 k=1
			.write "</tr><tr>" &vbcrlf
			else 
			 k=k+1
			end if
		  Next
		  do while (k<=4 and k>1)
		  .Write "<td width='25%' bgcolor='#ffffff'>&nbsp;</td>"
		  k=k+1
		  loop
		  .Write "</tr>"&vbcrlf
		  .Write "<tr><td colspan='4' bgcolor='#ffffff'><input type='button' onClick=""CheckAll(this.form)"" value='全选' class='button'>&nbsp;&nbsp;<input type='button' value='反选' class='button' onClick=""ContraSel(this.form)"">&nbsp;&nbsp;<input class=Button type=submit name=Submit2 value='删除选中的图片' onClick=""document.myform.action.value='Del';return confirm('确定要删除选中的图片吗？')"">&nbsp;&nbsp;<input class=Button type=submit name=Submit2 value='删除本商品的所有图片' onClick=""document.myform.action.value='DelEmpty';return confirm('确定要删除本商品的所有图片吗？')""></td></tr>"
		  .Write "</form>"
		  .Write "<tr><td class='clefttitle' align='center' colspan='4'>"
		     Call KSCLS.ShowPage(totalPut, MaxPerPage, "KS.ProImages.asp", True, "张", Page, "ProID=" & ProID & "&GroupName=" & KS.G("GroupName") & "&" & SearchParam)

		  .Write "</td></tr>"
		  End With
		  
		End Sub
		
		'保存
		Sub DoSave()
		   With Response
			'On Error Resume Next
			'图片ID
            ProID = KS.ChkClng(KS.G("ProID"))
			picnum=KS.ChkClng(KS.G("picnum"))
			If ProID = 0 Or picnum=0 Then .Write ("<script>alert('参数传递出错!');history.back(-1);</script>")
			   
			   If KS.ChkClng(KS.G("BeyondSavePic"))=1 Then
				  	SaveFilePath = KS.GetUpFilesDir & "/"
					KS.CreateListFolder (SaveFilePath)
				End If
			
	        		Dim RS,Images
					Set RS=Server.CreateObject("ADODB.RECORDSET")
			        SqlStr = "SELECT * FROM KS_ProImages Where 1=0"
					For I=1 To PicNum
					   Dim BigPicUrl:BigPicUrl=KS.G("ImgUrl"& I)
					   Dim SmallPicUrl:SmallPicUrl= KS.G("Thumb" & I)
  					   If Left(Lcase(BigPicurl),4)="http" and KS.ChkClng(KS.G("BeyondSavePic"))=1 Then
					     BigPicurl=Replace(KS.ReplaceBeyondUrl(BigPicurl, SaveFilePath),KS.GetDomain,KS.Setting(3)) 
					   End If
  					   If Left(Lcase(SmallPicUrl),4)="http" and KS.ChkClng(KS.G("BeyondSavePic"))=1 Then
					     SmallPicUrl=Replace(KS.ReplaceBeyondUrl(SmallPicUrl, SaveFilePath),KS.GetDomain,KS.Setting(3))
					   End If
					   
					  If KS.G("Thumb" & I)<>"" and KS.G("ImgUrl" & i)<>"" Then
						RS.Open SqlStr, conn, 1, 3
						RS.AddNew			
						RS("ProID") = ProID
						RS("SmallPicUrl") = SmallPicUrl
						RS("BigPicUrl")=BigPicUrl
						RS("GroupName")=KS.G("GroupName" & I)
						RS.Update
						RS.Close
					  End If
					  Images=Images&smallpicurl&bigpicurl
			        Next

					'关联上传文件
				   Call KS.FileAssociation(5,ProID,images ,0)
				   Set RS = Nothing
					
				  .Write "<script>alert('恭喜，图片添加成功！');location.href='" & ComeUrl & "';</script>"
		 End With		
		End Sub
		
		Sub DelImages()
		 Dim RS,SmallPicURL,BigPicUrl
		 Set RS=Conn.execute("select * from ks_proimages where id in(" &KS.FilterIDs(KS.G("ID")) & ")")
		 Do While Not RS.Eof
		  SmallPicUrl=RS("SmallPicUrl")
		  BigPicURL=RS("BigPicUrl")
		  Call KS.DeleteFile(SmallPicUrl) 
		  Call KS.DeleteFile(BigPicUrl)
		  Conn.Execute("Delete From [KS_UploadFiles] where infoid=" & rs("proid") & " and (filename='" & rs("SmallPicUrl") & "' or filename ='" & rs("BigPicUrl") & "')")
		 RS.MoveNext
		 Loop
		 RS.Close:Set RS=Nothing
		 Conn.Execute("Delete From KS_ProImages where id in(" & KS.FilterIDs(KS.G("ID")) & ")")
		 Response.Redirect ComeURL
		End Sub
        
		Sub DelEmpty()
		 Dim RS,SmallPicURL,BigPicUrl
		 Set RS=Conn.execute("select * from ks_proimages where proid=" &KS.ChkClng(KS.G("ProID")))
		 Do While Not RS.Eof
		  SmallPicUrl=RS("SmallPicUrl")
		  BigPicURL=RS("BigPicUrl")
		  Call KS.DeleteFile(SmallPicUrl) 
		  Call KS.DeleteFile(BigPicUrl)
		  Conn.Execute("Delete From [KS_UploadFiles] where infoid=" & rs("proid") & " and (filename='" & rs("SmallPicUrl") & "' or filename ='" & rs("BigPicUrl") & "')")

		 RS.MoveNext
	 Loop
		 RS.Close:Set RS=Nothing
		 Conn.Execute("Delete From KS_ProImages where proid=" & KS.ChkClng(KS.G("ProID")))
		 Response.Redirect ComeURL
		End Sub
End Class
%> 
