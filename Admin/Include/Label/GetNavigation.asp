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
Set KSCls = New GetNavigation
KSCls.Kesion()
Set KSCls = Nothing

Class GetNavigation
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
Dim InstallDir, CurrPath, FolderID, LabelContent, L_C_A, Action, LabelID, Str, Descript
Dim TypeFlag, OpenType, NavType, Navi, TitleCss, ColNumber, SplitPic, ChannelID,PrintType,DivID,DivClass,UlID,UlClass,LiID,LiClass
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
	LabelContent       = Replace(Replace(LabelContent, "{Tag:GetNavigation", ""),"}{/Tag}", "")
	Dim XMLDoc,Node
	Set XMLDoc=KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
	 If XMLDoc.loadxml("<label><param " & LabelContent & " /></label>") Then
			  Set Node=XMLDoc.DocumentElement.SelectSingleNode("param")
	 Else
			 .echo ("<Script>alert('��ǩ���س���!');history.back();</Script>")
			 Exit Sub
	  End If
	If  Not Node Is Nothing Then
			ChannelID = Node.getAttribute("channelid")
			NavType = Node.getAttribute("navtype")
			Navi = Node.getAttribute("nav")
			SplitPic = Node.getAttribute("splitpic")
			ColNumber = Node.getAttribute("col")
			OpenType =  Node.getAttribute("opentype")
			TitleCss =  Node.getAttribute("titlecss")
			PrintType=  Node.getAttribute("printtype")
			DivID    =  Node.getAttribute("divid")
			divclass =  Node.getAttribute("divclass")
			ulid     =  Node.getAttribute("ulid")
			ulclass  =  Node.getAttribute("ulclass")
			LIID     =  Node.getAttribute("liid")
			LIClass  =  Node.getAttribute("liclass")
	End If
	Set Node=Nothing
	XMLDoc=Empty
End If
If PrintType="" Then PrintType=2
If Navi = "" Then Navi = " | "
If ColNumber = "" Then ColNumber = 10
		.echo "<html>"
		.echo "<head>"
		.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		.echo "<link href=""../admin_style.css"" rel=""stylesheet"">"
		.echo "<script src=""../../../ks_inc/Common.js"" language=""JavaScript""></script>"
		.echo "<script src=""../../../ks_inc/Jquery.js"" language=""JavaScript""></script>"
%>
<script type="text/javascript">
        $(document).ready(function(){
		  ChangeOutArea($("#PrintType>option[selected=true]").val());
		})
		function ChangeOutArea(Val)
		{
		 if (Val==2){
		  $("#TableArea").hide();
		  $("#DivArea").show();
		  $("#TableShow").hide();
		 }
		 else{
		  $("#TableArea").show();
		  $("#DivArea").hide();
		  $("#TableShow").show();
		 }
		}
		function SetNavStatus()
		{
		  if ($("select[name=NavType]").val()==0)
		   { $("#NavWord").show();
			 $("#NavPic").hide();
		  }else{
		     $("#NavWord").hide();
		     $("#NavPic").show();
		 }
		}
		function CheckForm()
		{
		    if ($("input[name=LabelName]").val()=='')
			 {
			  alert('�������ǩ����');
			  $("input[name=LabelName]").focus(); 
			  return false
			  }
			var ChannelID=$("#ChannelID").val();
			var OpenType=$("#OpenType").val();
			var Nav,NavType=$("#NavType").val();
			var SplitPic=$("input[name=SplitPic]").val();
			var ColNumber=$("input[name=ColNumber]").val();
			var TitleCss=$("input[name=TitleCss]").val();
					var PrintType=$("select[name=PrintType]").val();
					var divid=$("input[name=divid]").val();
					var divclass=$("input[name=divclass]").val();
					var ulid=$("input[name=ulid]").val();
					var ulclass=$("input[name=ulclass]").val();
					var liid=$("input[name=liid]").val();
					var liclass=$("input[name=liclass]").val();
			if  (NavType==0) Nav=$("#TxtNavi").val();
			 else  Nav=$("#NaviPic").val();
			var tagVal='{Tag:GetNavigation labelid="0" channelid="'+ChannelID+'" navtype="'+NavType+'" nav="'+Nav+'" splitpic="'+SplitPic+'" col="'+ColNumber+'" opentype="'+OpenType+'" titlecss="'+TitleCss+'" printtype="'+PrintType+'" divid="'+divid+'" divclass="'+divclass+'" ulid="'+ulid+'" ulclass="'+ulclass+'" liid="'+liid+'" liclass="'+liclass+'"}{/Tag}'
			$("input[name=LabelContent]").val(tagVal);
			$("#myform").submit();
		}
		</script>
<%
		.echo "</head>"
		.echo "<body topmargin=""0"" leftmargin=""0"" scroll=no>"
		.echo "<div align=""center"">"
		.echo "<iframe src='about:blank' name='_hiddenframe' id='_hiddenframe' width='0' height='0'></iframe>"
		.echo "<form  method=""post"" id=""myform"" name=""myform"" action=""AddLabelSave.asp"" target='_hiddenframe'>"
		.echo " <input type=""hidden"" name=""LabelContent"" id=""LabelContent"">"
		.echo "   <input type=""hidden"" name=""LabelFlag"" id=""LabelFlag"" value=""2"">"
		.echo " <input type=""hidden"" name=""Action"" value=""" & Action & """>"
		.echo "  <input type=""hidden"" name=""LabelID"" value=""" & LabelID & """>"
		.echo " <input type=""hidden"" name=""FileUrl"" value=""GetNavigation.asp"">"
		.echo KS.ReturnLabelInfo(LabelName, FolderID, Descript)
		.echo "          <table width=""98%"" style='margin-top:5px' border='0' align='center' cellpadding='2' cellspacing='1' class='border'>"

		.echo "            <tr class=tdbg>"
		.echo "              <td width=""50%"" height=""24"">�����ʽ"
		.echo " <select class='textbox' style='width:70%' name=""PrintType"" id=""PrintType"" onChange=""ChangeOutArea(this.options[this.selectedIndex].value);"">"
        .echo "  <option value=""1"""
		If PrintType=1 Then .echo " selected"
		.echo ">��ͨTable��ʽ</option>"
        .echo "  <option value=""2"""
		If PrintType=2 Then .echo " selected"
		.echo ">DIV+CSS��ʽ</option>"
        .echo "</select>"
		.echo "              </td>"
		.echo "              <td width=""50%"" height=""24"">"
		.echo "<div id=""TableArea"""
		If PrintType=2 Then .echo "style=""display:none"""
		.echo "><font color=blue>��ѡ��ϵͳ֧�ֵ������ʽ</font></div><span id=""DivArea"""
		If PrintType<>2 Then .echo "style=""display:none"""
		.echo ">&lt;div id=&quot; <input name=""divid"" type=""text"" value=""" & Divid &""" id=""divid"" size=""6""  style=""border-top-width: 0px;border-right-width: 0px;border-bottom-width: 1px;border-left-width:0px;border-bottom-color: #000000"" class='textbox' title=""DIV���õ�ID�ţ�������CSS��Ԥ�ȶ����Ҳ���Ϊ��!"">&quot; class=&quot; <input name=""divclass"" class='textbox' type=""text"" value=""" & Divclass &""" id=""divclass"" size=""6"" style=""border-top-width: 0px;	border-right-width: 0px;border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"" title=""DIV���õ�Class���ƣ�����CSS��Ԥ�ȶ���,����Ϊ��!""> &quot;&gt;<span style=""color:blue"">�˴�������������</span><br> &lt;ul  id=&quot; <input value=""" & ulid &""" name=""ulid"" type=""text"" id=""ulid"" class='textbox' size=""6"" style=""border-top-width: 0px;border-right-width: 0px;border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000""  title=""����ul���õ�ID������CSS��Ԥ�ȶ���,����Ϊ��!""> &quot; class=&quot; <input class='textbox' value=""" & ulclass &""" name=""ulclass""  type=""text"" id=""ulclass"" size=""6"" style=""border-top-width: 0px;border-right-width: 0px;border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"" title=""����ul���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!"">&quot;&gt;<span style=""color:blue"">�˴�������������</span><br>&lt;li id=&quot; <input value=""" & liid &""" name=""liid"" type=""text"" id=""liid"" size=""6"" class='textbox' style=""border-top-width: 0px;	border-right-width: 0px;border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000"" title=""����li���õ�ID������CSS��Ԥ�ȶ��塣����Ϊ��!"">&quot; class=&quot; <input value=""" & liclass &""" name=""liclass"" class='textbox' type=""text"" id=""liclass"" size=""6"" style=""border-top-width: 0px;border-right-width: 0px;border-bottom-width: 1px;border-left-width: 0px;border-bottom-color: #000000""  title=""����li���õ�class���ƣ�����CSS��Ԥ�ȶ��塣����Ϊ��!""> &quot;&gt;</div></td>"
		.echo "            </tr>"
.echo "            <tr class='tdbg'>"
.echo "              <td height=""30"" colspan=""2"">ѡ��Χ"
.echo " " & ReturnAllChannel(ChannelID)    
.echo "             <font color=red>��ѡ��ǰƵ��ͨ�ã�����Ӧ��Ƶ�������ø��Ե�����Ŀ</font></td>"
.echo "            </tr>"

.echo "      <tbody id=""TableShow"">"
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
          .echo ("<input type=""text"" class=""textbox"" id=""TxtNavi"" name=""TxtNavi"" style=""width:70%;"" value=""" & Navi & """>")
          .echo ("</div>")
          .echo ("<div align=""left"" id=NavPic style=""display:none""> ")
          .echo ("<input type=""text"" class=""textbox"" readonly style=""width:55%;"" id=""NaviPic"" name=""NaviPic"">")
          .echo ("<input class='button' type=""button"" onClick=""OpenThenSetValue('../SelectPic.asp?CurrPath=" & CurrPath & "&ShowVirtualPath=true',550,290,window,document.myform.NaviPic);"" name=""Submit3"" value=""ѡ��ͼƬ"">")
          .echo ("</div>")
        Else
          .echo ("<div align=""left"" id=""NavWord"" style=""display:none""> ")
          .echo ("<input type=""text"" class=""textbox"" id=""TxtNavi"" name=""TxtNavi"" style=""width:70%;"">")
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
.echo "          </tbody>"
.echo "            <tr class='tdbg'>"
.echo "              <td height=""30"">��ʾ����"
.echo "                <input name=""ColNumber"" class=""textbox"" type=""text"" id=""ColNumber"" style=""width:70%"" value=""" & ColNumber & """></td>"
.echo "              <td height=""30"">"
          
 .echo KS.ReturnOpenTypeStr(OpenType)
    
.echo "              </td>"
.echo "            </tr>"
.echo "           <tr class='tdbg'>"
.echo "              <td height=""30"">������ʽ"
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

'ȡ����վ������Ƶ����������Ŀ
Function ReturnAllChannel(FolderID)
  Dim ChannelStr:ChannelStr = ""
      ChannelStr = "<select class='textbox' name=""ChannelID"" id=""ChannelID"" style=""width:200;border-style: solid; border-width: 1"">"
      ChannelStr = ChannelStr & "<option value=""0"">    -��վ����-  </option>"
	  if FolderID="9999" then
	  ChannelStr = ChannelStr & "<option value=""9999"" style=""color:red"" selected>-��ǰƵ��ͨ��-</option>"
	  else
	  ChannelStr = ChannelStr & "<option value=""9999"" style=""color:red"">-��ǰƵ��ͨ��-</option>"
	  end if
	 if FolderID="9998" then
	   ChannelStr = ChannelStr & "<option value=""9998"" style=""color:blue"" selected>-ͬ��Ƶ��ͨ��-</option>"
	   else
	   ChannelStr = ChannelStr & "<option value=""9998"" style=""color:blue"">-ͬ��Ƶ��ͨ��-</option>"
	   end if

		ChannelStr = ChannelStr & "<optgroup  label=""-----���˿ռ���ص���-----"">"
	  if FolderID="9997" then
	  ChannelStr = ChannelStr & "<option value=""9997"" selected>-�ռ����-</option>"
	  else
	  ChannelStr = ChannelStr & "<option value=""9997"">-�ռ����-</option>"
	  end if
	  if FolderID="9996" then
	  ChannelStr = ChannelStr & "<option value=""9996"" selected>-��־����-</option>"
	  else
	  ChannelStr = ChannelStr & "<option value=""9996"">-��־����-</option>"
	  end if
	  if FolderID="9995" then
	  ChannelStr = ChannelStr & "<option value=""9995"" selected>-Ȧ�ӷ���-</option>"
	  else
	  ChannelStr = ChannelStr & "<option value=""9995"">-Ȧ�ӷ���-</option>"
	  end if
	  if FolderID="9994" then
	  ChannelStr = ChannelStr & "<option value=""9994"" selected>-������-</option>"
	  else
	  ChannelStr = ChannelStr & "<option value=""9994"">-������-</option>"
	  end if

		ChannelStr = ChannelStr & "<optgroup  label=""-----ָ����ģ��-----"">"
		ChannelStr = ChannelStr & ReturnChannelOption(FolderID)
   ChannelStr = ChannelStr & "</Select>"
   ReturnAllChannel = ChannelStr
End Function
	'**************************************************
	'��������ReturnChannelOption
	'��  �ã���ʾƵ���б�
	'��  ����SelectChannelID ----ѡ��Ƶ��ID��
	'����ֵ��Ƶ���б�
	'**************************************************
	Public Function ReturnChannelOption(SelectChannelID)
	  Dim RS:Set RS=Server.CreateObject("ADODB.Recordset")
	  Dim SQL,K,ChannelStr:ChannelStr = ""
	   RS.Open "Select channelid,channelname From [KS_Channel] Where ChannelStatus=1 And ChannelID<>10 and channelid<>9", Conn, 1, 1
	   If RS.EOF And RS.BOF Then
		  RS.Close:Set RS = Nothing:Exit Function
	   Else
	     SQL=RS.GetRows(-1):rs.close:set rs=nothing
	   End iF
		
	    For K=0 To ubound(sql,2)
		  If Cstr(sql(0,k)) = Cstr(SelectChannelID) Then
		  ChannelStr = ChannelStr & "<option selected value=" & sql(0,k) & ">" & sql(1,k) & "</option>"
		 Else
		   ChannelStr = ChannelStr & "<option value=" & sql(0,k) & ">" & sql(1,k) & "</option>"
		 End If
		Next 
		ChannelStr = ChannelStr & "<optgroup  label=""-----ָ�����������Ŀ(�����г�����վ�ĵ�����)----"">"  
	   For K=0 To Ubound(sql,2)
	        ChannelStr=ChannelStr & Replace(KS.LoadClassOption(sql(0,k)),"value='" & SelectChannelID & "'","value='" & SelectChannelID &"' selected")
	    Next
	   ReturnChannelOption = ChannelStr
	End Function

End Class
%> 
