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
Set KSCls = New GetSpaceList
KSCls.Kesion()
Set KSCls = Nothing

Class GetSpaceList
        Private KS
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
'���岿��
Public Sub Kesion()
Dim InstallDir, CurrPath, FolderID, LabelContent, Action, LabelID, Str, Descript
Dim TypeFlag, OpenType, NavType, Navi, TitleCss, Num, TitleLen,SplitPic, ChannelID,PrintType,AjaxOut,LabelStyle,recommend,MoreStr,ClassID,ShowType,OrderStr,RowHeight,logo,banner
FolderID = Request("FolderID")
CurrPath = KS.GetCommonUpFilesDir()
With KS
'�ж��Ƿ�༭
LabelID = Trim(Request.QueryString("LabelID"))
If LabelID = "" Then
  Action = "Add"
Else
    Action = "Edit"
  Dim LabelRS, LabelName
  Set LabelRS = Server.CreateObject("Adodb.Recordset")
  LabelRS.Open "Select top 1 * From KS_Label Where ID='" & LabelID & "'", Conn, 1, 1
  If LabelRS.EOF And LabelRS.BOF Then
     LabelRS.Close
     Set LabelRS = Nothing
     .echo ("<Script>alert('�������ݳ���!');window.close();</Script>")
     .End
  End If
    LabelName = Replace(Replace(LabelRS("LabelName"), "{LB_", ""), "}", "")
    FolderID = LabelRS("FolderID")
    Descript = LabelRS("Description")
    LabelContent = LabelRS("LabelContent")
    LabelRS.Close
    Set LabelRS = Nothing
            LabelStyle         = KS.GetTagLoop(LabelContent)
			LabelContent       = Replace(Replace(LabelContent, "{Tag:GetSpaceList", ""),"}" & LabelStyle&"{/Tag}", "")
			' response.write LabelContent
			Dim XMLDoc,Node
			Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
			If XMLDoc.loadxml("<label><param " & LabelContent & " /></label>") Then
			  Set Node=XMLDoc.DocumentElement.SelectSingleNode("param")
			Else
			 .echo ("<Script>alert('��ǩ���س���!');history.back();</Script>")
			 Exit Sub
			End If
			If  Not Node Is Nothing Then
			    ClassID          = Node.getAttribute("classid")
				Num              = Node.getAttribute("num")
				TitleLen         = Node.getAttribute("titlelen")
				NavType          = Node.getAttribute("navtype")
				Navi             = Node.getAttribute("nav")
				SplitPic         = Node.getAttribute("splitpic")
				OpenType         = Node.getAttribute("opentype")
				TitleCss         = Node.getAttribute("titlecss")
				AjaxOut          = Node.getAttribute("ajaxout")
				PrintType        = Node.getAttribute("printtype")
				ShowType         = Node.getAttribute("showtype")		
				recommend        = Node.getAttribute("recommend")
				logo             = Node.getAttribute("logo")
				banner           = Node.getAttribute("banner")
				MoreStr          = Node.getAttribute("morestr")
				OrderStr         = Node.getAttribute("orderstr")
				RowHeight        = Node.getAttribute("rowheight")
			End If
			XMLDoc=Empty
			Set Node=Nothing
    
End If
		If PrintType="" Then PrintType=1
		If TitleLen="" Then TitleLen=0
		If Num = "" Then Num = 10
		If recommend="" then recommend=false
		If banner="" or isnull(banner) Then banner=false
		If logo="" or isnull(logo) Then logo=false
		If ShowType="" Then ShowType=0
		If RowHeight="" Then RowHeight=20
		If LabelStyle="" Then LabelStyle="[loop={@num}] " & vbcrlf & "<li><a href=""{@spaceurl}"" target=""_blank"">{@blogname}</a></li>" & vbcrlf & "[/loop]"
		
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
			$("input[name=ShowType]").click(function(){
			  if ($(this).val()==1)
			  { $("#spaceclass").show();
			   }else{
			    $("#spaceclass").hide();
			   }
			});
			$("input[name=ShowType][value=<%=ShowType%>]").attr("checked",true);
			if ($("input[name=ShowType][checked=true]").attr("value")==1){
			  $("#spaceclass").show();
			}else{
			  $("#spaceclass").hide();
			}
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
		 //����
		function InsertValue(Val)
		{  if (pos==null) {alert('���ȶ�λҪ�����λ��!');return false;}
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
			  alert('�������ǩ����');
			  $("input[name=LabelName]").focus(); 
			  return false
			  }
	var recommendFlag,logo,banner;
	var ClassID=document.myform.ClassID.value;
	var OpenType=document.myform.OpenType.value;
	var Nav,NavType=document.myform.NavType.value;
	var SplitPic=document.myform.SplitPic.value;
	var Num=document.myform.Num.value;
	var RowHeight=$("#RowHeight").val();
	var TitleLen=document.myform.TitleLen.value;
	var TitleCss=document.myform.TitleCss.value;
	var ShowType=$("input[name=ShowType][checked=true]").val();
	var PrintType=$("#PrintType").val();
	var OrderStr=$("#OrderStr").val();
	var AjaxOut=false;
	if ($("#AjaxOut").attr("checked")==true){AjaxOut=true}
			
	
	if  (NavType==0) Nav=document.myform.TxtNavi.value
	 else  Nav=document.myform.NaviPic.value;
	if (document.myform.recommend.checked)
	   recommendFlag= true
	else
	   recommendFlag=false;
	if (document.myform.logo.checked)
	  logo= true
	else
	  logo=false;
	if (document.myform.banner.checked)
	   banner= true
	else
	   banner=false;
	   
    var MoreStr=document.myform.MoreStr.value;
	if (Num=='') Num=10
	
	var tagVal='{Tag:GetSpaceList labelid="0" printtype="'+PrintType+'" ajaxout="'+AjaxOut+'" classid="'+ClassID+'" opentype="'+OpenType+'" num="'+Num+'" rowheight="'+RowHeight+'" orderstr="'+OrderStr+'" showtype="'+ShowType+'" titlelen="'+TitleLen+'" navtype="'+NavType+'" nav="'+Nav+'"  morestr="'+MoreStr+'" splitpic="'+SplitPic+'" titlecss="'+TitleCss+'" recommend="'+recommendFlag+'" logo="'+logo+'" banner="'+banner+'"}'+$("#LabelStyle").val()+'{/Tag}';
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
		.echo " <input type=""hidden"" name=""LabelContent"">"
		.echo "   <input type=""hidden"" name=""LabelFlag"" id=""LabelFlag"" value=""2"">"
		.echo " <input type=""hidden"" name=""Action"" id=""Action"" value=""" & Action & """>"
		.echo "  <input type=""hidden"" name=""LabelID"" value=""" & LabelID & """>"
		.echo " <input type=""hidden"" name=""FileUrl"" value=""GetSpaceList.asp"">"
		.echo KS.ReturnLabelInfo(LabelName, FolderID, Descript)
		.echo "          <table width=""98%"" style='margin-top:5px' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"

		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">�����ʽ"
		.echo " <select class='textbox' style='width:70%' name=""PrintType"" id=""PrintType"" onChange=""ChangeOutArea();"">"
        .echo "  <option value=""1"""
		If PrintType=1 Then .echo " selected"
		.echo ">��ͨTable��ʽ</option>"
        .echo "  <option value=""2"""
		If PrintType=2 Then .echo " selected"
		.echo ">�Զ��������ʽ</option>"
        .echo "</select>"
		.echo "              </td>"
		.echo "              <td width=""50%"" height=""24"">"
		.echo "            <label><input type='checkbox' name='AjaxOut' id='AjaxOut' value='1'"
		If AjaxOut="true" Then .echo " checked"
		.echo ">����Ajax���</label></td>"
		.echo "            </tr>"
		
		
		.echo "            <tbody id=""DiyArea"">"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' id='ShowFieldArea' class='field'><li onclick=""InsertLabel('{@autoid}')"">�� ��</li><li onclick=""InsertLabel('{@spaceurl}')"">�ռ�URL</li> <li onclick=""InsertLabel('{@blogname}')"">�ռ�����</li><li onclick=""InsertLabel('{@logo}')"">�ռ�logo</li><li onclick=""InsertLabel('{@banner}')"">�ռ�banner</li><li onclick=""InsertLabel('{@hits}')"">�����</li></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2'><textarea name='LabelStyle' onkeyup='setPos()' onclick='setPos()' id='LabelStyle' style='width:95%;height:150px'>" & LabelStyle & "</textarea></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' class='attention'><strong><font color=red>ʹ��˵�� :</font></strong><br />ѭ����ǩ[loop=n][/loop]�Կ���ʡ��,Ҳ����ƽ�г��ֶ�ԣ�</td>"
		.echo "            </tr>"
		.echo "           </tbody>"
		
		.echo "           <tr class='tdbg'>"
		.echo "            <td>��ʾ����:"
		.echo "             <input type='radio' name='ShowType' value='0'>����"
		.echo "             <input type='radio' name='ShowType' value='1'>���˿ռ�"
		.echo "             <input type='radio' name='ShowType' value='2'>��ҵ�ռ�"
		.echo "            </td>"
		.echo "              <td height=""30"">"
		
       If cbool(recommend) = True Then
		 .echo ("<input type=""checkbox"" value=""true"" name=""recommend"" checked>����ʾ�Ƽ��Ŀռ�")
	   Else
		 .echo ("<input type=""checkbox"" value=""true"" name=""recommend"">����ʾ�Ƽ��Ŀռ�")
	   End If
	   
       If cbool(logo) = True Then
		 .echo ("<input type=""checkbox"" value=""true"" name=""logo"" checked>����ʾ��Logo")
	   Else
		 .echo ("<input type=""checkbox"" value=""true"" name=""logo"">����ʾ��Logo")
	   End If
       If cbool(banner) = True Then
		 .echo ("<input type=""checkbox"" value=""true"" name=""banner"" checked>����ʾ��banner")
	   Else
		 .echo ("<input type=""checkbox"" value=""true"" name=""banner"">����ʾ��banner")
	   End If
	  
	  
        .echo "</td>"
		.echo "           </tr>"
		
		.echo "            <tr class='tdbg' id='spaceclass'>"
		.echo "              <td height=""30"" colspan='2'>վ�����"
		.echo "                  <select class=""textbox"" size='1' name='ClassID' style=""width:270px"">"
		.echo "                    <option value=""0"">-��ָ��Spaceվ�����-</option>"
                              Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
							  RS.Open "Select * From KS_BlogClass order by orderid",conn,1,1
							  If Not RS.EOF Then
							   Do While Not RS.Eof 
							   If TRIM(ClassID)=TRIM(RS("ClassID")) Then
								  .echo "<option value=""" & RS("ClassID") & """ selected>" & RS("ClassName") & "</option>"
							   Else
								  .echo "<option value=""" & RS("ClassID") & """>" & RS("ClassName") & "</option>"
							   End iF
								 RS.MoveNext
							   Loop
							  End If
							  RS.Close:Set RS=Nothing
.echo "                </select>"							  
.echo "                </td>"

.echo "            </tr>"
.echo "            <tr class='tdbg'>"
.echo "              <td height=""30"">��ʾ����"
.echo "                <input name=""Num"" class=""textbox"" type=""text"" id=""Num"" style=""width:50px"" value=""" & Num & """>�� �и�<input name=""RowHeight"" class=""textbox"" type=""text"" id=""RowHeight"" style=""width:50px"" value=""" & RowHeight & """></td>"
.echo "              <td height=""30"">"
          
 .echo KS.ReturnOpenTypeStr(OpenType)
    
.echo "              </td>"
.echo "            </tr>"
.echo "            <tr class='tdbg'>"
.echo "              <td height=""30"">��ʾ����"
.echo "                <input name=""TitleLen"" class=""textbox"" type=""text"" id=""TitleLen"" style=""width:50px"" value=""" & TitleLen & """><font color=red>���������ƣ�������Ϊ��0��</font></td>"
.echo "              <td height=""30"">����ʽ"
.echo "                <select style=""width:70%;"" class='textbox' name=""OrderStr"" id=""OrderStr"">"
					If OrderStr = "BlogID Desc" Then
					.echo ("<option value='BlogID Desc' selected>�ռ�ID(����)</option>")
					Else
					.echo ("<option value='BlogID Desc'>�ռ�ID(����)</option>")
					End If
					If OrderStr = "BlogID Asc" Then
					.echo ("<option value='BlogID Asc' selected>�ռ�ID(����)</option>")
					Else
					.echo ("<option value='BlogID Asc'>�ռ�ID(����)</option>")
					End If

					
					
					If OrderStr = "B.Hits Asc" Then
					 .echo ("<option value='B.Hits Asc' selected>�����(����)</option>")
					Else
					 .echo ("<option value='B.Hits Asc'>�����(����)</option>")
					End If
					If OrderStr = "Hits Desc" Then
					  .echo ("<option value='B.Hits Desc' selected>�����(����)</option>")
					Else
					  .echo ("<option value='B.Hits Desc'>�����(����)</option>")
					End If

		.echo "         </select></td>"
.echo "            </tr>"



.echo "            <tr class='tdbg'>"
.echo "              <td width=""50%"" height=""30"">��������"
.echo "                <select class=""textbox"" name=""NavType"" id=""NavType"" style=""width:70%;"" onchange=""SetNavStatus()"">"
            
            If LabelID = "" Or CStr(NavType) = "0" Then
            .echo ("<option value=""0"" selected>���ֵ���</option>")
            .echo ("<option value=""1"">ͼƬ����</option>")
           Else
            .echo ("<option value=""0"">���ֵ���</option>")
            .echo ("<option value=""1"" selected>ͼƬ����</option>")
           End If
           
.echo "                </select> </td>"
.echo "              <td>"
        
        If LabelID = "" Or CStr(NavType) = "0" Then
          .echo ("<div align=""left"" id=""NavWord""> ")
          .echo ("<input type=""text"" class=""textbox"" name=""TxtNavi"" id='TxtNavi' style=""width:70%;"" value=""" & Navi & """>")
          .echo ("</div>")
          .echo ("<div align=""left"" id=NavPic style=""display:none""> ")
          .echo ("<input type=""text"" class=""textbox"" readonly style=""width:55%;"" id=""NaviPic"" name=""NaviPic"">")
          .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.NaviPic);"" name=""Submit3"" value=""ѡ��ͼƬ"">")
          .echo ("</div>")
        Else
          .echo ("<div align=""left"" id=""NavWord"" style=""display:none""> ")
          .echo ("<input type=""text"" class=""textbox"" id='TxtNavi' name=""TxtNavi"" style=""width:70%;"">")
          .echo ("</div>")
          .echo ("<div align=""left"" id=NavPic> ")
          .echo ("<input type=""text"" class=""textbox"" readonly style=""width:55%;"" id=""NaviPic"" name=""NaviPic"" value=""" & Navi & """>")
          .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.NaviPic);"" name=""Submit3"" value=""ѡ��ͼƬ"">")
          .echo ("</div>")
        End If
        
.echo "              </td>"
.echo "            </tr>"
.echo "            <tr class='tdbg'>"
.echo "             <td height=""30"" colspan=""2"">�ָ�ͼƬ"
.echo "                <input name=""SplitPic"" class=""textbox"" type=""text"" id=""SplitPic"" style=""width:61%;"" value=""" & SplitPic & """ readonly>"
.echo "                <input  class='button' name=""SubmitPic"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.SplitPic);"" type=""button"" id=""SubmitPic2"" value=""ѡ��ͼƬ..."">"
.echo "                <span style=""cursor:pointer;color:green;"" onclick=""javascript:document.myform.SplitPic.value='';"" onmouseover=""this.style.color='red'"" onMouseOut=""this.style.color='green'"">���</span>"
.echo "              </td>"
.echo "            </tr>"

.echo "           <tr class='tdbg'>"
.echo "              <td height=""30"">�����־"
.echo "                <input name=""MoreStr"" class=""textbox"" type=""text"" id=""MoreStr"" style=""width:70%;"" value=""" & MoreStr & """></td>"
.echo "              <td height=""30""><font color=""#FF0000"">���Ҫ��ʾ���࣬�������־��""����..."",""more""</font></td>"
.echo "            </tr>"

.echo "           <tr class='tdbg'>"
.echo "              <td height=""30"">Css ��ʽ"
.echo "                <input name=""TitleCss"" class=""textbox"" type=""text"" id=""TitleCss"" style=""width:70%;"" value=""" & TitleCss & """></td>"
.echo "              <td height=""30""><font color=""#FF0000"">�Ѷ����CSS ,Ҫ��һ������ҳ��ƻ���</font></td>"
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
