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
Set KSCls = New GetBlogInfoList
KSCls.Kesion()
Set KSCls = Nothing

Class GetBlogInfoList
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
Dim TypeID,InstallDir, CurrPath, FolderID, LabelContent, L_C_A, Action, LabelID, Str, Descript,DateRule,DateAlign,UserName
Dim TypeFlag, OpenType, NavType, Navi, TitleCss, Num, TitleLen,SplitPic, ChannelID,PrintType,isbest,morestr,ShowType,AjaxOut,LabelStyle,OrderStr,RowHeight
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
  LabelRS.Open "Select * From KS_Label Where ID='" & LabelID & "'", Conn, 1, 1
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
			LabelContent       = Replace(Replace(LabelContent, "{Tag:GetBlogInfoList", ""),"}" & LabelStyle&"{/Tag}", "")
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
			    AjaxOut         = Node.getAttribute("ajaxout")
			    TypeID          = Node.getAttribute("typeid")
				Num             = Node.getAttribute("num")
				TitleLen        = Node.getAttribute("titlelen")
				UserName        = Node.getAttribute("username")
				NavType         = Node.getAttribute("navtype")
				Navi            = Node.getAttribute("nav")
				SplitPic        = Node.getAttribute("splitpic")
				OpenType        = Node.getAttribute("opentype")
				DateRule        = Node.getAttribute("daterule")
				DateAlign       = Node.getAttribute("datealign")
				TitleCss        = Node.getAttribute("titlecss")
			    PrintType       = Node.getAttribute("printtype")
				isbest          = Node.getAttribute("isbest")
				morestr         = Node.getAttribute("morestr")
				OrderStr        = Node.getAttribute("orderstr")
				RowHeight       = Node.getAttribute("rowheight")
			End If
			XMLDoc=Empty
			Set Node=Nothing
    
End If
		If PrintType="" Then PrintType=1
		If TitleLen="" Then TitleLen=0
		If Num = "" Then Num = 10
		If isbest="" Then isbest=false
		If RowHeight="" Then RowHeight=20
		If LabelStyle="" Then LabelStyle="[loop={@num}] " & vbcrlf & "<li><a href=""{@logurl}"" target=""_blank"">{@title}</a></li>" & vbcrlf & "[/loop]"
		
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
	var isbest;
	var TypeID=document.myform.TypeID.value;
	var OpenType=document.myform.OpenType.value;
	var Nav,NavType=document.myform.NavType.value;
	var SplitPic=document.myform.SplitPic.value;
	var Num=document.myform.Num.value;
	var TitleLen=document.myform.TitleLen.value;
	var UserName=document.myform.UserName.value;
	var DateRule=document.myform.DateRule.value;
	var DateAlign=document.myform.DateAlign.value;
	var TitleCss=document.myform.TitleCss.value;
	var PrintType=$("#PrintType").val();
	var OrderStr=$("#OrderStr").val();
	var RowHeight=$("#RowHeight").val();
	var AjaxOut=false;
	if ($("#AjaxOut").attr("checked")==true){AjaxOut=true}
			
	
	if (document.myform.isbest.checked)
	   isbest= true
	else
	   isbest=false;
	if  (NavType==0) Nav=document.myform.TxtNavi.value
	 else  Nav=document.myform.NaviPic.value;
	var MoreStr=document.myform.MoreStr.value;
		 
	var tagVal='{Tag:GetBlogInfoList labelid="0" printtype="'+PrintType+'" ajaxout="'+AjaxOut+'" typeid="'+TypeID+'" opentype="'+OpenType+'" num="'+Num+'" titlelen="'+TitleLen+'" rowheight="'+RowHeight+'" username="'+UserName+'" navtype="'+NavType+'" nav="'+Nav+'" orderstr="'+OrderStr+'"   morestr="'+MoreStr+'" splitpic="'+SplitPic+'" daterule="'+DateRule+'" datealign="'+DateAlign+'" titlecss="'+TitleCss+'" isbest="'+isbest+'"}'+$("#LabelStyle").val()+'{/Tag}';
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
		.echo " <input type=""hidden"" name=""LabelContent"" id='LabelContent'>"
		.echo " <input type=""hidden"" name=""LabelFlag"" id='LabelFlag' value=""2"">"
		.echo " <input type=""hidden"" name=""Action"" id='Action' value=""" & Action & """>"
		.echo "  <input type=""hidden"" name=""LabelID"" id='LabelID' value=""" & LabelID & """>"
		.echo " <input type=""hidden"" name=""FileUrl"" value=""GetBlogInfoList.asp"">"
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
		.echo "              <td colspan='2' id='ShowFieldArea' class='field'><li onclick=""InsertLabel('{@autoid}')"">�� ��</li><li onclick=""InsertLabel('{@logurl}')"">��־URL</li> <li onclick=""InsertLabel('{@title}')"">��־����</li><li onclick=""InsertLabel('{@adddate}')"">���ʱ��</li><li onclick=""InsertLabel('{@username}')"">�û���</li><li onclick=""InsertLabel('{@hits}')"">�����</li><li onclick=""InsertLabel('{@typeid}')"">����ID</li><li onclick=""InsertLabel('{@logclassname}')"">��������</li></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2'><textarea name='LabelStyle' onkeyup='setPos()' onclick='setPos()' id='LabelStyle' style='width:95%;height:150px'>" & LabelStyle & "</textarea></td>"
		.echo "            </tr>"
		.echo "            <tr class=tdbg>"
		.echo "              <td colspan='2' class='attention'><strong><font color=red>ʹ��˵�� :</font></strong><br />ѭ����ǩ[loop=n][/loop]�Կ���ʡ��,Ҳ����ƽ�г��ֶ�ԣ�</td>"
		.echo "            </tr>"
		.echo "           </tbody>"
		
		.echo "            <tr class='tdbg'>"
		.echo "              <td height=""30"">��־����"
		.echo "                <select class=""textbox"" size='1' name='TypeID' style=""width:70%"">"
		.echo "                   <option value=""0"">-��ָ�����-</option>"
                    Dim RS:Set RS=Server.CreateObject("ADODB.RECORDSET")
							  RS.Open "Select * From KS_BlogType order by orderid",conn,1,1
							  If Not RS.EOF Then
							   Do While Not RS.Eof 
							   If Trim(TypeID)=Trim(RS("TypeID")) Then
								  .echo "<option value=""" & RS("TypeID") & """ selected>" & RS("TypeName") & "</option>"
							   Else
								  .echo "<option value=""" & RS("TypeID") & """>" & RS("TypeName") & "</option>"
							   End iF
								 RS.MoveNext
							   Loop
							  End If
							  RS.Close:Set RS=Nothing
							  
.echo "                 </select> "
.echo "               </td>"
.echo "              <td height=""30"">"
    If cbool(isbest) = True Then
		 .echo ("<input type=""checkbox"" value=""true"" name=""isbest"" checked>����ʾ��������־")
	Else
		 .echo ("<input type=""checkbox"" value=""true"" name=""isbest"">����ʾ��������־")
	End If

.echo "</td>"
.echo "            </tr>"
.echo "            <tr class='tdbg'>"
.echo "              <td height=""30"">��ʾƪ��"
.echo "                <input name=""Num"" class=""textbox"" type=""text"" id=""Num"" style=""width:50px"" value=""" & Num & """> �о�<input type=""text"" name=""RowHeight"" id=""RowHeight"" style=""width:50px"" value=""" & RowHeight & """></td>"
.echo "              <td height=""30"">"
          
 .echo KS.ReturnOpenTypeStr(OpenType)
    
.echo "              </td>"
.echo "            </tr>"
.echo "            <tr class='tdbg'>"
.echo "              <td height=""30"">���ⳤ��"
.echo "                <input name=""TitleLen"" class=""textbox"" type=""text"" id=""TitleLen"" style=""width:50px"" value=""" & TitleLen & """><font color=red>���������ƣ�������Ϊ��0��</font></td>"
.echo "              <td height=""30"">����ʽ"
.echo "                <select style=""width:70%;"" class='textbox' name=""OrderStr"" id=""OrderStr"">"
					If OrderStr = "ID Desc" Then
					.echo ("<option value='ID Desc' selected>��־ID(����)</option>")
					Else
					.echo ("<option value='ID Desc'>��־ID(����)</option>")
					End If
					If OrderStr = "ID Asc" Then
					.echo ("<option value='ID Asc' selected>��־ID(����)</option>")
					Else
					.echo ("<option value='ID Asc'>��־ID(����)</option>")
					End If

					
					
					If OrderStr = "Hits Asc" Then
					 .echo ("<option value='Hits Asc' selected>�����(����)</option>")
					Else
					 .echo ("<option value='Hits Asc'>�����(����)</option>")
					End If
					If OrderStr = "Hits Desc" Then
					  .echo ("<option value='Hits Desc' selected>�����(����)</option>")
					Else
					  .echo ("<option value='Hits Desc'>�����(����)</option>")
					End If

.echo "         </select>"
.echo "              </td>"
.echo "            </tr>"
.echo "            <tr class='tdbg'>"
.echo "              <td height=""30"">ָ���û�"
.echo "                <input name=""UserName"" class=""textbox"" type=""text"" id=""UserName"" style=""width:70%"" value=""" & UserName & """></td>"
.echo "              <td height=""30""><font color=red>����ʾָ���û�����־,����������</font>"
.echo "              </td>"
.echo "            </tr>"



.echo "            <tr class='tdbg'>"
.echo "              <td width=""50%"" height=""30"">��������"
.echo "                <select class=""textbox"" name=""NavType"" style=""width:70%;"" onchange=""SetNavStatus()"">"
            
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
          .echo ("<input type=""text"" class=""textbox"" name=""TxtNavi"" id=""TxtNavi"" style=""width:70%;"" value=""" & Navi & """>")
          .echo ("</div>")
          .echo ("<div align=""left"" id=NavPic style=""display:none""> ")
          .echo ("<input type=""text"" class=""textbox"" readonly style=""width:55%;"" id=""NaviPic"" name=""NaviPic"">")
          .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.NaviPic);"" name=""Submit3"" value=""ѡ��ͼƬ"">")
          .echo ("</div>")
        Else
          .echo ("<div align=""left"" id=""NavWord"" style=""display:none""> ")
          .echo ("<input type=""text"" class=""textbox"" name=""TxtNavi"" id=""TxtNavi"" style=""width:70%;"">")
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
		.echo "            <tr class=tdbg>"
		.echo "              <td height=""24"">���ڸ�ʽ"
		.echo "                <select class='textbox' style=""width:70%;"" name=""DateRule"" id=""select2"">"
		.echo KS.ReturnDateFormat(DateRule)
		.echo "                </select> </td>"
		.echo "              <td height=""24"">"
		.echo "                <div align=""left"">���ڶ���"
		.echo "                  <select class=""textbox"" name=""DateAlign"" id=""select3"" style=""width:70%;"">"
							
					If LabelID = "" Or CStr(DateAlign) = "left" Then
					 Str = " selected"
					Else
					 Str = ""
					End If
					 .echo ("<option value=""left""" & Str & ">�����</option>")
					If CStr(DateAlign) = "center" Then
					 Str = " selected"
					Else
					 Str = ""
					End If
					 .echo ("<option value=""center""" & Str & ">���ж���</option>")
					If CStr(DateAlign) = "right" Then
					 Str = " selected"
					Else
					 Str = ""
					End If
					 .echo ("<option value=""right""" & Str & ">�Ҷ���</option>")
					 
		.echo "                  </select>"
		.echo "                </div></td>"
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
