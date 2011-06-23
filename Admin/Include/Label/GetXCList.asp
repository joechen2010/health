<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../../Conn.asp"-->
<!--#include file="../../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New GetXCList
KSCls.Kesion()
Set KSCls = Nothing

Class GetXCList
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
'主体部分
Public Sub Kesion()
Dim InstallDir, CurrPath, FolderID, LabelContent, L_C_A, Action, LabelID, Str, Descript,UserName
Dim ClassID,ShowStyle, OpenType, XCWidth, XCHeight, TitleCss, Num, TitleLen,ColNumber, ChannelID,PrintType,DivID,DivClass,UlID,UlClass,LiID,LiClass,recommend,morestr,AjaxOut,LabelStyle,OrderStr
FolderID = Request("FolderID")
CurrPath = KS.GetCommonUpFilesDir()
With KS
'判断是否编辑
LabelID = Trim(Request.QueryString("LabelID"))
If LabelID = "" Then
  Action = "Add"
Else
    Action = "Edit"
  Dim LabelRS, LabelName
  Set LabelRS = Server.CreateObject("Adodb.Recordset")
  LabelRS.Open "Select * From KS_Label Where ID='" & LabelID & "'", Conn, 1, 1
  If LabelRS.EOF And LabelRS.BOF Then
     LabelRS.Close
     Set LabelRS = Nothing
     .echo ("<Script>alert('参数传递出错!');window.close();</Script>")
     .End
  End If
    LabelName = Replace(Replace(LabelRS("LabelName"), "{LB_", ""), "}", "")
    FolderID = LabelRS("FolderID")
    Descript = LabelRS("Description")
    LabelContent = LabelRS("LabelContent")
    LabelRS.Close
    Set LabelRS = Nothing
	
	   LabelStyle         = KS.GetTagLoop(LabelContent)
	   LabelContent       = Replace(Replace(LabelContent, "{Tag:GetXCList", ""),"}" & LabelStyle&"{/Tag}", "")
			' response.write LabelContent
			Dim XMLDoc,Node
			Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			If XMLDoc.loadxml("<label><param " & LabelContent & " /></label>") Then
			  Set Node=XMLDoc.DocumentElement.SelectSingleNode("param")
			Else
			 .echo ("<Script>alert('标签加载出错!');history.back();</Script>")
			 Exit Sub
			End If
			If  Not Node Is Nothing Then
			    AjaxOut         = Node.getAttribute("ajaxout")
			    ClassID         = Node.getAttribute("classid")
				Num             = Node.getAttribute("num")
				ColNumber       = Node.getAttribute("col")
				UserName        = Node.getAttribute("username")
				TitleLen        = Node.getAttribute("titlelen")
				ShowStyle       = Node.getAttribute("showstyle")
				XCWidth         = Node.getAttribute("xcwidth")
				XCHeight        = Node.getAttribute("xcheight")
				OpenType        = Node.getAttribute("opentype")
				TitleCss        = Node.getAttribute("titlecss")
				PrintType       = Node.getAttribute("printtype")
				recommend       = Node.getAttribute("recommend")
				morestr         = Node.getAttribute("morestr")
				OrderStr        = Node.getAttribute("orderstr")
			End If
			XMLDoc=Empty
			Set Node=Nothing
    
End If
		If XCHeight="" Then XCHeight=100
		If XCWidth="" Then XCWidth=85
		If ColNumber="" Then ColNumber=3
		If PrintType="" Then PrintType=1
		If ShowStyle = "" Then ShowStyle = 1
		If TitleLen="" Then TitleLen=0
		If Num = "" Then Num = 10
		IF recommend="" Then recommend=false
		If AjaxOut="" Then AjaxOut=false
		If LabelStyle="" Then LabelStyle="[loop={@num}] " & vbcrlf & "<li><a href=""{@albumsurl}"" target=""_blank"">{@xcname}</a></li>" & vbcrlf & "[/loop]"

		.echo "<html>"
		.echo "<head>"
		.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		.echo "<link href=""../admin_style.css"" rel=""stylesheet"">"
		.echo "<script src=""../../../ks_inc/Common.js"" language=""JavaScript""></script>"
		.echo "<script src=""../../../ks_inc/Jquery.js"" language=""JavaScript""></script>"
%>
        <style type="text/css">
		 .field{width:720px;}
		 .field li{cursor:pointer;float:left;border:1px solid #DEEFFA;background-color:#F7FBFE;height:18px;line-height:18px;margin:3px 1px 0px;padding:2px}
		 .field li.diyfield{border:1px solid #f9c943;background:#FFFFF6}
		</style>
        <script type="text/javascript">
		$(document).ready(function(){
	        ChangeOutArea();
	   })
		
       function InsertLabel(label)
		{
		  InsertValue(label);
		}
		var pos=null;
		 function setPos()
		 { if (document.all){
				$("#LabelStyle").focus();
				pos = document.selection.createRange();
			  }else{
				pos = document.getElementById("LabelStyle").selectionStart;
			  }
		 }
		 //插入
		function InsertValue(Val)
		{  if (pos==null) {alert('请先定位要插入的位置!');return false;}
			if (document.all){
				  pos.text=Val;
			}else{
				   var obj=$("#LabelStyle");
				   var lstr=obj.val().substring(0,pos);
				   var rstr=obj.val().substring(pos);
				   obj.val(lstr+Val+rstr);
			}
		 }
		function ChangeOutArea()
		{
		 var Val=$("#PrintType").val();
		  if (Val==2){
		   $("#DiyArea").show();
		  }else{
		  $("#DiyArea").hide();
		  }
		}
	function SetNavStatus()
	{
		  if ($("select[name=NavType]").val()==0)
			   {$("#NavWord").show();
				$("#NavPic").hide();
				}else{
			   $("#NavWord").hide();
			   $("#NavPic").show();}
	}
	function CheckForm()
		{
		    if ($("input[name=LabelName]").val()=='')
			 {
			  alert('请输入标签名称');
			  $("input[name=LabelName]").focus(); 
			  return false
			  }
	var recommendFlag;
	var ClassID=document.myform.ClassID.value;
	var ShowStyle=document.myform.ShowStyle.value;
	var OpenType=document.myform.OpenType.value;
	var XCWidth=document.myform.XCWidth.value;
	var XCHeight=document.myform.XCHeight.value;
	var ColNumber=document.myform.ColNumber.value;
	var Num=document.myform.Num.value;
	var TitleLen=document.myform.TitleLen.value;
	var TitleCss=document.myform.TitleCss.value;
	var UserName=document.myform.UserName.value;
	var PrintType=document.myform.PrintType.value;
	var AjaxOut=false;
	if ($("#AjaxOut").attr("checked")==true){AjaxOut=true}
	var OrderStr=$("#OrderStr").val();

	if (document.myform.recommend.checked)
	   recommendFlag= true
	else
	   recommendFlag=false;
    var MoreStr=document.myform.MoreStr.value;
	
	var tagVal='{Tag:GetXCList labelid="0" printtype="'+PrintType+'" ajaxout="'+AjaxOut+'" classid="'+ClassID+'" opentype="'+OpenType+'" num="'+Num+'"  col="'+ColNumber+'" orderstr="'+OrderStr+'" titlelen="'+TitleLen+'" username="'+UserName+'" showstyle="'+ShowStyle+'" xcwidth="'+XCWidth+'" xcheight="'+XCHeight+'"  morestr="'+MoreStr+'" titlecss="'+TitleCss+'" recommend="'+recommendFlag+'"}'+$("#LabelStyle").val()+'{/Tag}';
	$("input[name=LabelContent]").val(tagVal);
	$("#myform").submit(); 
}
</script>
<%
.echo "</head>"
.echo "<body topmargin=""0"" leftmargin=""0"">"
.echo "<div align=""center"">"
.echo "<iframe src='about:blank' name='_hiddenframe' id='_hiddenframe' width='0' height='0'></iframe>"
.echo "<form  method=""post"" id=""myform"" name=""myform"" action=""AddLabelSave.asp"" target='_hiddenframe'>"
.echo " <input type=""hidden"" name=""LabelContent"" id=""LabelContent"">"
.echo "   <input type=""hidden"" name=""LabelFlag"" value=""2"">"
.echo " <input type=""hidden"" name=""Action"" value=""" & Action & """>"
.echo "  <input type=""hidden"" name=""LabelID"" value=""" & LabelID & """>"
.echo " <input type=""hidden"" name=""FileUrl"" value=""GetXCList.asp"">"
.echo KS.ReturnLabelInfo(LabelName, FolderID, Descript)
		.echo "          <table width=""98%"" style='margin-top:5px' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"

		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">输出格式"
		.echo " <select class='textbox' style='width:70%' name=""PrintType"" id=""PrintType"" onChange=""ChangeOutArea();"">"
        .echo "  <option value=""1"""
		If PrintType=1 Then .echo " selected"
		.echo ">普通Table格式</option>"
        .echo "  <option value=""2"""
		If PrintType=2 Then .echo " selected"
		.echo ">自定义输出样式</option>"
        .echo "</select>"
		.echo "              </td>"
		.echo "              <td width=""50%"" height=""24"">"
		.echo "            <label><input type='checkbox' name='AjaxOut' id='AjaxOut' value='1'"
		If AjaxOut="true" Then .echo " checked"
		.echo ">采用Ajax输出</label></td>"
		.echo "            </tr>"
		
		
		.echo "            <tbody id=""DiyArea"">"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' id='ShowFieldArea' class='field'><li onclick=""InsertLabel('{@autoid}')"">行 号</li><li onclick=""InsertLabel('{@albumsurl}')"">相册URL</li> <li onclick=""InsertLabel('{@xcname}')"">相册名称</li><li onclick=""InsertLabel('{@xps}')"">相片数</li><li onclick=""InsertLabel('{@username}')"">用户名</li><li onclick=""InsertLabel('{@photourl}')"">相册封面</li><li onclick=""InsertLabel('{@adddate}')"">上传时间</li><li onclick=""InsertLabel('{@hits}')"">浏览数</li><li onclick=""InsertLabel('{@descript}')"">简要介绍</li></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2'><textarea name='LabelStyle' onkeyup='setPos()' onclick='setPos()' id='LabelStyle' style='width:95%;height:150px'>" & LabelStyle & "</textarea></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' class='attention'><strong><font color=red>使用说明 :</font></strong><br />循环标签[loop=n][/loop]对可以省略,也可以平行出现多对；</td>"
		.echo "            </tr>"
		.echo "           </tbody>"
		
.echo "            <tr class='tdbg'>"
.echo "              <td height=""30"">相册分类"
.echo "                <select class=""textbox"" size='1' name='ClassID' style=""width:70%"">"
.echo "                   <option value=""0"">-不指定类别-</option>"
                    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
							  RS.Open "Select * From KS_PhotoClass order by orderid",conn,1,1
							  If Not RS.EOF Then
							   Do While Not RS.Eof 
							   If Trim(ClassID)=Trim(RS("ClassID")) Then
								  .echo "<option value=""" & RS("ClassID") & """ selected>" & RS("ClassName") & "</option>"
							   Else
								  .echo "<option value=""" & RS("ClassID") & """>" & RS("ClassName") & "</option>"
							   End iF
								 RS.MoveNext
							   Loop
							  End If
							  RS.Close:Set RS=Nothing
							  
.echo "                 </select> "
.echo "               </td>"
.echo "              <td height=""30"">"
    If cbool(recommend) = True Then
		 .echo ("<input type=""checkbox"" value=""true"" name=""recommend"" checked>仅显示推荐的相册")
	Else
		 .echo ("<input type=""checkbox"" value=""true"" name=""recommend"">仅显示推荐的相册")
	End If

.echo "</td>"
.echo "            </tr>"
.echo "            <tr class='tdbg'>"
.echo "              <td height=""30"">显示数量"
.echo "                <input name=""Num"" class=""textbox"" type=""text"" id=""Num"" style=""width:70%"" value=""" & Num & """></td>"
.echo "              <td height=""30"">每行数量<input name=""ColNumber"" class=""textbox"" type=""text"" id=""ColNumber"" style=""width:70%;"" value=""" & ColNumber & """>"
    
.echo "              </td>"
.echo "            </tr>"
.echo "            <tr class='tdbg'>"
.echo "              <td height=""30"">显示长度"
.echo "                <input name=""TitleLen"" class=""textbox"" type=""text"" id=""TitleLen"" style=""width:70%"" value=""" & TitleLen & """></td>"
.echo "              <td height=""30""><font color=red>如果不想控制，请设置为“0”</font>"
.echo "              </td>"
.echo "            </tr>"
.echo "            <tr class='tdbg'>"
.echo "              <td height=""30"">输出样式"
          
.echo "              <select name=""ShowStyle"" class=""textbox"" style=""width:70%"">"
              .echo "<option value=""1"" name=""ShowStyle"""
			  If ShowStyle = "1" Then .echo " selected"
			  .echo ">样式一</option>"
              .echo "<option value=""2"" name=""ShowStyle"""
			  If ShowStyle="2" Then .echo " selected"
			  .echo ">样式二</option>"
.echo "     </select>"          
.echo "              </td>"
.echo "              <td><font color=red>请选择系统支持的样式类型</font></td>"
.echo "            </tr>"


.echo "            <tr class='tdbg'>"
.echo "              <td width=""50%"" height=""30"">图片宽度"
.echo "                <input type=""text"" class=""textbox"" name=""XCWidth"" value=""" & XCwidth & """ style=""width:70%;"">"
           
.echo "                 </td>"
.echo "              <td>图片高度"
 .echo ("<input type=""text"" class=""textbox"" name=""XCHeight"" style=""width:70%;"" value=""" & XCHeight & """>")
.echo "              </td>"
.echo "            </tr>"
.echo "            <tr class='tdbg'>"
.echo "             <td height=""30"">"
.echo KS.ReturnOpenTypeStr(OpenType)
.echo "              </td>"
.echo "              <td>排序方式:"
.echo "                <select style=""width:70%;"" class='textbox' name=""OrderStr"" id=""OrderStr"">"
					If OrderStr = "ID Desc" Then
					.echo ("<option value='ID Desc' selected>相册ID(降序)</option>")
					Else
					.echo ("<option value='ID Desc'>相册ID(降序)</option>")
					End If
					If OrderStr = "ID Asc" Then
					.echo ("<option value='ID Asc' selected>相册ID(升序)</option>")
					Else
					.echo ("<option value='ID Asc'>相册ID(升序)</option>")
					End If

					If OrderStr = "AddDate Desc" Then
					.echo ("<option value='AddDate Desc' selected>相册添加时间(降序)</option>")
					Else
					.echo ("<option value='AddDate Desc'>相册添加时间(降序)</option>")
					End If
					If OrderStr = "AddDate Asc" Then
					.echo ("<option value='AddDate Asc' selected>相册添加时间(升序)</option>")
					Else
					.echo ("<option value='AddDate Asc'>相册添加时间(升序)</option>")
					End If
					
					If OrderStr = "Hits Asc" Then
					 .echo ("<option value='Hits Asc' selected>点击数(升序)</option>")
					Else
					 .echo ("<option value='Hits Asc'>点击数(升序)</option>")
					End If
					If OrderStr = "Hits Desc" Then
					  .echo ("<option value='Hits Desc' selected>点击数(降序)</option>")
					Else
					  .echo ("<option value='Hits Desc'>点击数(降序)</option>")
					End If

		.echo "         </select></td>"
.echo "            </tr>"
.echo "            <tr class='tdbg'>"
.echo "             <td height=""30"">指定用户"
.echo "                <input type=""text"" class=""textbox"" name=""UserName"" style=""width:70%;"" value=""" & UserName & """>"
.echo "              </td>"
.echo "             <td><font color=red>仅显示指定用户的相册，否则请留空</font></td>"
.echo "            </tr>"
.echo "           <tr class='tdbg'>"
.echo "              <td height=""30"">更多标志"
.echo "                <input name=""MoreStr"" class=""textbox"" type=""text"" id=""MoreStr"" style=""width:70%;"" value=""" & MoreStr & """></td>"
.echo "              <td height=""30""><font color=""#FF0000"">如果要显示更多，请输入标志如""更多..."",""more""</font></td>"
.echo "            </tr>"

.echo "           <tr class='tdbg'>"
.echo "              <td height=""30"">Css 样式"
.echo "                <input name=""TitleCss"" class=""textbox"" type=""text"" id=""TitleCss"" style=""width:70%;"" value=""" & TitleCss & """></td>"
.echo "              <td height=""30""><font color=""#FF0000"">已定义的CSS ,要有一定的网页设计基础</font></td>"
.echo "            </tr>"
.echo "                  </table>"	
.echo "  </form>"
  
.echo "</div>"
.echo "</body>"
.echo "</html>"
End With

End Sub
End Class
%> 
