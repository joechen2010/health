<%option explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../KS_Cls/Kesion.Label.BaseFunCls.asp"-->
<!--#include file="../KS_Cls/Kesion.EscapeCls.asp"-->
<%
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"
Dim KS:Set KS=New PublicCls

Dim TaskXML,TaskNode,Node,N,TaskUrl,Taskid,Action
'Set TaskXML=LFCls.GetXMLFromFile("task")
set TaskXML = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
TaskXML.async = false
TaskXML.setProperty "ServerHTTPRequest", true 
TaskXML.load(Server.MapPath(KS.Setting(3)&"Config/task.xml"))
Set TaskNode=TaskXML.DocumentElement.SelectNodes("item[@isenable=1]")


Action=Request.QueryString("Action")
Select Case Action
  case "manage"Manage
  case "add","modify" add
  case "DoSave" DoSave
  case "ModifySave" ModifySave
  case "del" del
  case "taskitem" taskitem
  case else
    Call Task()
End Select

Sub taskitem()
  Dim tasktype:tasktype=KS.ChkClng(KS.G("tasktype"))
  Dim SQLStr,RS,selectid
  selectid=Request("selectid")
  Select Case TaskType
     case 1
	  SQLStr="Select ItemID,ItemName From KS_CollectItem Order By ItemID desc"
	  Set RS=Server.CreateObject("ADODB.RECORDSET")
	  RS.Open SQLStr,KS.ConnItem,1,1
	  KS.Echo Escape("<br/><strong>ѡ��Ҫ��ʱ�ɼ�����Ŀ</strong><br/>")
	  KS.Echo "<select name=""taskid"" id=""taskid"" size=10 multiple style=""width:240px"">"
	  Do While Not RS.EOF
		  If KS.FoundInArr(selectid,RS(0),",") Then
		   KS.Echo Escape("<option value=""" & RS(0) & """ selected>" & RS(1) & "</option>")
		  Else
		   KS.Echo Escape("<option value=""" & RS(0) & """>" & RS(1) & "</option>")
		  End If
	   RS.MoveNext
	  Loop
	  KS.Echo "</select>"
	  KS.Echo Escape("<br/><font color=red>���԰�סctrl�����ж�ѡ</font><br/>")
	  RS.Close
	  Set RS=Nothing
	 Case 2,3
	  KS.Echo Escape("<br/><strong>ѡ��Ҫ��ʱ��������Ŀ</strong><br/>")
	  KS.Echo "<select name=""taskid"" id=""taskid"" size=10 multiple style=""width:240px"">"
	   Dim i,Str,IDArr:IDArr=Split(selectid,",")
	   Str=KS.LoadClassOption(0)
	   For I=0 To Ubound(IDArr)
	    str=Replace(str,"value='" & IDArr(i) & "'","value='" & IDArr(i) &"' selected")
	   Next
	  KS.Echo Escape(str)
	  KS.Echo "</select>"
	  KS.Echo Escape("<br/><font color=red>���԰�סctrl�����ж�ѡ</font><br/>")
      If TaskType=3 Then
	   KS.Echo Escape("�޶�������ӵ�<input type=""text"" name=""limitnum"" size=""4"" value=""50"" style=""text-align:center"">ƪ�ĵ�")
	  End If	 
  End Select
End Sub


Sub manage()
%>
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>
<title>��ʱ�������</title>
<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>
<script language="JavaScript" src="../KS_Inc/Jquery.js"></script>
<script LANGUAGE="JavaScript"> 
<!-- 
function openwin() { 
window.open ("KS.Task.asp", "newwindow", "height=450, width=550, top=0, left=0, toolbar=no, menubar=no, scrollbars=yes, resizable=no, location=no, status=no")
} 
--> 
</script>
</head>
<body>
<ul id='mt'> <div id='mtl'>��ʱ�������</div><li><a href="?action=add"><img src="images/ico/as.gif" border='0' align='absmiddle'>�������</a></li></ul>
	  <table width="100%" align='center' border="0" cellpadding="0" cellspacing="0">
      <tr class="sort">
	    <td>����ID</td>
	    <td>��������</td>
		<td>��������</td>
		<td>ִ������</td>
		<td>ִ��ʱ��</td>
		<td>״ ̬</td>
		<td>�������</td>
	  </tr>
<%
  If TaskXML.DocumentElement.SelectNodes("item").length=0 Then
      Response.Write "<tr class='list'><td colspan=7 height='25' class='splittd' align='center'>��û����Ӷ�ʱ����!</td></tr>"
  Else
	  N=0
	  For Each Node In TaskXML.DocumentElement.SelectNodes("item")
	  %>
			  <tr  onmouseout="this.className='list'" onMouseOver="this.className='listmouseover'">               
			   <td class='splittd' height="30" align="center"><%=Node.SelectSingleNode("@id").text%></td>
			   <td class='splittd' height="30"><%=Node.SelectSingleNode("name").text%></td>
			   <td class='splittd' align="center">
			   <%
			   Select Case Node.SelectSingleNode("tasktype").text
				case "1" Response.write "�ɼ�"
				case "0" Response.Write "������ҳ"
				case "2" Response.Write "������Ŀҳ"
				case "3" Response.Write "��������ҳ"
			   End Select
			   %>
			   </td>
			   <td class='splittd' align="center">
				<font color=red>
			   <%
				if Node.SelectSingleNode("starttype").text="1" then
				 response.write "ÿ��"
				ElseIf Node.SelectSingleNode("starttype").text="3" then
				 response.write "ʱ���"
				Else
				 response.write "ÿ�� "
				 Select Case  Node.SelectSingleNode("week").text
				  case 0 response.write "������"
				  case 1 response.write "����һ"
				  case 2 response.write "���ڶ�"
				  case 3 response.write "������"
				  case 4 response.write "������"
				  case 5 response.write "������"
				  case 6 response.write "������"
				 End Select
				End If
				%>
				</font>
			   </td>
			   <td class='splittd' align="center">
			   <%
			    If Node.SelectSingleNode("starttype").text="3" then
			     response.write " " & KS.GotTopic(Node.SelectSingleNode("time").text,45) &"..."
			    else
			     response.write " " & Node.SelectSingleNode("time").text
				End If
			   %></td>
			   <td align="center" class="splittd">
				<%
				 if node.selectSingleNode("@isenable").text="1" then
				  response.write "<font color=blue>����</font>"
				 else
				  response.write "<font color=green>�ر�</font>"
				 end if
				%>
			   </td>
			   <td class='splittd' align="center">
				 <a href="?action=modify&itemid=<%=Node.SelectSingleNode("@id").text%>">�޸�</a> | <a href="?action=del&itemid=<%=Node.SelectSingleNode("@id").text%>" onClick="return(confirm('ȷ��ɾ����������?'))">ɾ��</a>
			   </td>
			  </tr>
	  <%
		n=n+1
	  Next
  End If
  %>
		
	  </table>
       <br/>
	   <div style="text-align:center">
	    <input name="Submit" type="button" onClick="openwin()" class="button" value="��ʼ����">
	   </div>
</body>
</html>
<%
End Sub

Sub Add()
 Dim ItemID:ItemID=KS.ChkClng(Request("ItemID"))
 Dim Node,TaskName,TaskType,StartType,time,week,taskid,limitnum,remark,Isenable,act,ChannelID
 Isenable=1
 starttype=1
 ChannelID=1
 limitnum=50
 act="DoSave"
 If ItemID<>0 Then
   Set Node=TaskXML.DocumentElement.SelectSingleNode("item[@id=" & ItemID & "]")
   If Not Node Is Nothing Then
    Isenable=Node.getAttribute("isenable")
    TaskName=Node.childNodes(0).text
	StartType=Node.childNodes(1).text
	week=Node.childNodes(2).text
	time=Node.childNodes(3).text
    TaskType=Node.childNodes(4).text
	taskid=Node.childNodes(5).text
	limitnum=Node.childNodes(6).text
	remark=Node.childNodes(7).text
	channelid=Node.childNodes(8).text
	Act="ModifySave"
   End If
 End If
%>
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>
<title>��ʱ�������</title>
<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>
<script language="JavaScript" src="../KS_Inc/Jquery.js"></script>
<script>
 $(document).ready(function()
 {
   $("#starttype").change(function()
   {
     if($(this).val()==2)
	  {
	   $("#weekarea").show();
	  }else{
	   $("#weekarea").hide();
	  }
	  if($(this).val()==3)
	  {
	   $("#time").attr("multiple","multiple");
	   $("#time").attr("style","width:200px;height:150px");
	  }else{
	   $("#time").removeAttr("multiple");
	   $("#time").removeAttr("style");
	  }
	  
	  
   });
   
   $("#tasktype").change(function(){
     getTaskID()
   });
   
   $("#channelid").change(function(){
     getClass();
   });
   
   <%if itemid<>0 and tasktype="1" then%>
    getTaskID();
   <%end if%>

   
 });
 
 function getClass()
 {
      $.get('../plus/ajaxs.asp',{action:'GetClassOption',channelid:$("#channelid>option[selected=true]").val()},function(data){
	     $("#typearea").html('<br/><b>ѡ��Ҫ��ʱ��������Ŀ</b><br/><select name="taskid" id="taskid" size=10 multiple style="width:240px"></select><br/><font color=red>���԰�סctrl�����ж�ѡ</font><br/>�޶�������ӵ�<input type="text" name="limitnum" size="4" value="50" style="text-align:center">ƪ�ĵ�');
	     $("#taskid").empty();
		 $('#taskid').append(unescape(data));
	  })
 
 }
 
 function getTaskID()
 {
    if ($("#tasktype>option[selected=true]").val()==3)
	{
	 $("#channelarea").show();
	 <%If itemid=0 then%>
	  $("#channelid>option[value=0]").attr("selected",true);
	 <%end if%>
	}else{
	 $("#channelarea").hide();
	}
	

 
    if ($("#tasktype>option[selected=true]").val()!=undefined && $("#tasktype>option[selected=true]").val()!=0&& $("#tasktype>option[selected=true]").val()!=3 && $("#tasktype>option[selected=true]").val()!='')
	{
	   $.get("KS.Task.asp",{action:"taskitem",tasktype:$("#tasktype>option[selected=true]").val(),selectid:"<%=taskid%>"},function(r){
	     $("#typearea").html(unescape(r));
	   });
	 }else{
	 $("#typearea").html('');
	 }
 }
 
 function CheckForm()
 {
 
   if ($("#TaskName").val()=='')
   {
     alert('��������������!');
	 $("#TaskName").focus();
	 return false
   }
   if ($("#tasktype").val()=='')
   {
     alert('��ѡ���ִ�е�����!');
	 $("#tasktype").focus();
	 return false;
   }
   if ($("#tasktype").val()!=0)
   { 
      if ($("#taskid>option[selected=true]").val()=='' || $("#taskid>option[selected=true]").val()==undefined)
	  {
	   if ($("#tasktype").val()==1)
	   alert('��ѡ��ɼ���Ŀ!');
	   else
	   alert('��ѡ����Ŀ!');
	   return false;
	  }
   }
   return true;
 }
</script>
<body>
<div class='topdashed sort'>���/�༭��ʱ����</div>
<br/>
   <form name="myform" action="KS.Task.asp?action=<%=Act%>" method="post" id="myform">
	  <table width='90%' style="margin:4px" align="center" BORDER='0' cellpadding='1'  cellspacing='1' class='ctable'>
		  <tr class='tdbg'>               
		   <td class='clefttitle' height="30" width="100" align='right'><strong>��������:</strong></td>
		   <td><input type="text" name="TaskName" id="TaskName" value="<%=TaskName%>"> ��:��ʱ������ҳ</td>
		  </tr>
		  <tr class='tdbg'>               
		   <td class='clefttitle' height="30" width="100" align='right'><strong>��ִ������:</strong></td>
		   <td>
		   <select name="tasktype" id="tasktype">
		     <option value="">--ѡ������--</option>
		     <option value="1"<%if tasktype="1" then response.write " selected"%>>��ʱ�ɼ�</option>
		     <option value="0"<%if tasktype="0" then response.write " selected"%>>������ҳ</option>
		     <option value="2"<%if tasktype="2" then response.write " selected"%>>������Ŀҳ</option>
		     <option value="3"<%if tasktype="3" then response.write " selected"%>>��������ҳ</option>
		   </select>
		   
		    <span id="channelarea"<%if tasktype<>"3" then%> style="display:None"<%end if%>>
		    <strong>ѡ��ģ��</strong><select id='channelid' name='channelid'>
			<option value='0'>---��ѡ��ģ��---</option>
			<%
			If not IsObject(Application(KS.SiteSN&"_ChannelConfig")) Then KS.LoadChannelConfig
			Dim ModelXML,MNode
			Set ModelXML=Application(KS.SiteSN&"_ChannelConfig")
			For Each MNode In ModelXML.documentElement.SelectNodes("channel")
			 if MNode.SelectSingleNode("@ks21").text="1" and MNode.SelectSingleNode("@ks0").text<>"6" and MNode.SelectSingleNode("@ks0").text<>"9" and MNode.SelectSingleNode("@ks0").text<>"10" And MNode.SelectSingleNode("@ks7").text<>"0" Then
			  If Trim(ChannelID)=Trim(MNode.SelectSingleNode("@ks0").text) Then
			  KS.echo "<option value='" &MNode.SelectSingleNode("@ks0").text &"' selected>" & MNode.SelectSingleNode("@ks1").text & "</option>"
			  Else
			  KS.echo "<option value='" &MNode.SelectSingleNode("@ks0").text &"'>" & MNode.SelectSingleNode("@ks1").text & "</option>"
			  End If
			 End If
			next
			
			%>
			</select>
			</span>
			
		   <div id="typearea">
		   <%if tasktype="3" or tasktype="2" then%>
		    <br/><b>ѡ��Ҫ��ʱ��������Ŀ</b><br/><select name="taskid" id="taskid" size=10 multiple style="width:240px">
			<%
			   Dim Str,IDArr:IDArr=Split(taskid,",")
			   if tasktype="3" then
			   Str=KS.LoadClassOption(ChannelID)
			   else
			   Str=KS.LoadClassOption(0)
			   end if
			   For I=0 To Ubound(IDArr)
				str=Replace(str,"value='" & IDArr(i) & "'","value='" & IDArr(i) &"' selected")
			   Next
			   KS.Echo str
			
			%>
			</select><br/><font color=red>���԰�סctrl�����ж�ѡ</font>
			<%if tasktype="3" then%>
			<br/>�޶�������ӵ�<input type="text" name="limitnum" size="4" value="<%=limitnum%>" style="text-align:center">ƪ�ĵ�
			<%end if%>
		   <%end if%>
		   
		   </div>
		   
		   </td>
		  </tr>
		  <tr class='tdbg'>               
		   <td class='clefttitle' height="30" width="100" align='right'><strong>ִ������:</strong></td>
		   <td>
		   <select name="starttype" id="starttype">
		     <option value="1"<%if starttype="1" then response.write " selected"%>>ÿ��</option>
		     <option value="2"<%if starttype="2" then response.write " selected"%>>ÿ��</option>
		     <option value="3"<%if starttype="3" then response.write " selected"%>>��ʱ���</option>
		   </select>
		   <span id="weekarea"<%if starttype<>"2" then response.write " style='display:none'"%>>
		    <select name="week" id="week">
			 <option value="0"<%if week="0" then response.write " selected"%>>������</option>
			 <option value="1"<%if week="1" then response.write " selected"%>>����һ</option>
			 <option value="2"<%if week="2" then response.write " selected"%>>���ڶ�</option>
			 <option value="3"<%if week="3" then response.write " selected"%>>������</option>
			 <option value="4"<%if week="4" then response.write " selected"%>>������</option>
			 <option value="5"<%if week="5" then response.write " selected"%>>������</option>
			 <option value="6"<%if week="6" then response.write " selected"%>>������</option>
			</select> 
		   </span>
		   </td>
		  </tr>
		  <tr class='tdbg'>               
		   <td class='clefttitle' height="30" width="100" align='right'><strong>ִ��ʱ��:</strong></td>
		   <td>
		   <%if starttype="3" then%>
		    <select name="time" id="time" style="width:200px;height:150px" multiple>
		   <%else%>
		    <select name="time" id="time">
		   <%end if%>
			<%dim i,Ta,Time_S : Time_S=CDate("00:00")
			 for i=1 to 144
			    Ta=Split(Time_S,":")
				If KS.FoundInArr(Time,Ta(0) & ":" & Ta(1),",") Then
				 Response.Write "<option value="""& Ta(0) & ":" & Ta(1) &""" selected>"& Ta(0) & "��" & Ta(1) &"��</option>"
				Else
				 Response.Write "<option value="""& Ta(0) & ":" & Ta(1) &""">"& Ta(0) & "��" & Ta(1) &"��</option>"
				End If
			    Time_S = CDate(Time_S) + CDate("00:10")
			 next  
			 %>
			</select>
			
		   </td>
		  </tr>
		  <tr class='tdbg'>               
		   <td class='clefttitle' height="30" width="100" align='right'><strong>�Ƿ�����:</strong></td>
		   <td>
		    <input type="radio" name="Isenable" value="0"<%if Isenable="0" then response.write " checked"%>>������
		    <input type="radio" name="Isenable" value="1"<%if Isenable="1" then response.write " checked"%>>����
		   </td>
		  </tr>
		  <tr class='tdbg'>               
		   <td class='clefttitle' height="30" width="100" align='right'><strong>��Ҫ˵��:</strong></td>
		   <td>
		    <textarea name="remark" style="width:350px;height:80px" class="textbox"><%=Remark%></textarea>
		   </td>
		  </tr>
      </table>

        <br/>
		<div style="text-align:center">
		 <Input type="hidden" value="<%=itemid%>" name="itemid">
		 <input type="submit" value="��������" class="button" onClick="return(CheckForm())">
		</div>
		</form>
</body>
</html>
<%
End Sub

'����
Sub DoSave()
 Dim TaskName:TaskName=KS.G("TaskName")
 Dim tasktype:tasktype=KS.ChkClng(Request.Form("tasktype"))
 Dim starttype:starttype=KS.ChkClng(Request.Form("starttype"))
 Dim week:week=KS.ChkClng(Request.Form("week"))
 Dim time:time=replace(request.form("time")," ","")
 Dim Isenable:Isenable=KS.ChkClng(Request.Form("Isenable"))
 Dim ChannelID:ChannelID=KS.ChkClng(Request.Form("channelid"))
 Dim TaskID:TaskID=Replace(Request.Form("TaskID")," ","")
 Dim limitnum:limitnum=KS.ChkClng(Request.Form("limitnum"))
 Dim Remark:Remark=Request.Form("Remark")
 
 Dim ItemID
 'ȡ��Ψһ����ID��
 If TaskXML.DocumentElement.SelectNodes("item").length<>0 Then
   ItemID=TaskXML.DocumentElement.SelectNodes("item").length+1
 Else
   ItemID=1
 End If
 
 Dim NodeStr,brstr
     brstr=chr(13)&chr(10)&chr(9)
     NodeStr="<item isenable=""" & IsEnable & """ id=""" & ItemID &""">" &brstr
	 NodeStr=NodeStr & "<name>" & TaskName & "</name>"&brstr
	 NodeStr=NodeStr & "<starttype>" & StartType & "</starttype>" &brstr
	 NodeStr=NodeStr & "<week>" & Week & "</week>"&brstr
	 NodeStr=NodeStr & "<time>" & Time & "</time>"&brstr
	 NodeStr=NodeStr & "<tasktype>" & TaskType & "</tasktype>"&brstr
	 NodeStr=NodeStr & "<taskid>" & TaskID & "</taskid>"&brstr
	 NodeStr=NodeStr & "<limitnum>" & limitnum & "</limitnum>" & brstr
	 NodeStr=NodeStr & "<remark><![CDATA[ " & Remark & "]]></remark>" & brstr
	 NodeStr=NodeStr & "<channelid>" & ChannelID &"</channelid>" & brstr
	 NodeStr=NodeStr & " </item>"&brstr
	 Dim XML2:set XML2 = KS.InitialObject("msxml2.FreeThreadedDOMDocument"& MsxmlVersion)
     XML2.LoadXml(NodeStr)
	 Dim NewNode:set NewNode=XML2.documentElement
	 
	 Dim TN:Set TN=TaskXML.DocumentElement
	 TN.appendChild(NewNode)
	 TaskXML.Save(Server.MapPath(KS.Setting(3)&"Config/task.xml"))
	 Response.Write "<script>if (confirm('��ϲ,��ʱ������ӳɹ�!')){location.href='?action=add'}else{location.href='?action=manage'}</script>"
End Sub

'�����޸�
Sub ModifySave()
 Dim TaskName:TaskName=KS.G("TaskName")
 Dim tasktype:tasktype=KS.ChkClng(Request.Form("tasktype"))
 Dim starttype:starttype=KS.ChkClng(Request.Form("starttype"))
 Dim week:week=KS.ChkClng(Request.Form("week"))
 Dim time:time=replace(request.form("time")," ","")
 Dim Isenable:Isenable=KS.ChkClng(Request.Form("Isenable"))
 Dim TaskID:TaskID=Replace(Request.Form("TaskID")," ","")
 Dim limitnum:limitnum=KS.ChkClng(Request.Form("limitnum"))
 Dim Remark:Remark=Request.Form("Remark")
 Dim ItemID:ItemID=KS.ChkClng(Request.Form("ItemID"))
 Dim ChannelID:ChannelID=KS.ChkClng(Request.Form("channelid"))
 Dim Node
 Set Node=TaskXML.DocumentElement.SelectSingleNode("item[@id=" & ItemID & "]")
 Node.Attributes.getNamedItem("isenable").text=isenable
 Node.childnodes(0).text=TaskName
 Node.childNodes(1).text=StartType
 Node.childNodes(2).text=week
 Node.childNodes(3).text=time
 Node.childNodes(4).text=TaskType
 Node.childNodes(5).text=taskid
 Node.childNodes(6).text=limitnum
 Node.childNodes(7).text=remark
 Node.childNodes(8).text=channelid
	 TaskXML.Save(Server.MapPath(KS.Setting(3)&"Config/task.xml"))
	 Response.Write "<script>alert('��ϲ,��ʱ�����޸ĳɹ�!');location.href='?action=manage'</script>"
End Sub

Sub Del()
  Dim ItemID:ItemID=KS.ChkClng(Request("itemid"))
  If ItemID=0 Then KS.AlertHintScript "�Բ���,��������!"
  Dim DelNode,Node,ID
  Set DelNode=TaskXML.DocumentElement.SelectSingleNode("item[@id=" & ItemID & "]")
  If DelNode Is Nothing  Then
   KS.AlertHintScript "�Բ���,��������!"
  End If
  TaskXML.DocumentElement.RemoveChild(DelNode)
  
  '���±ȵ�ǰ����ID���ID��,���μ�һ
  For Each Node In TaskXML.DocumentElement.SelectNodes("item")
     ID=KS.ChkClng(Node.SelectSingleNode("@id").text)
	 If ID>ItemID Then
	    Node.SelectSingleNode("@id").text=ID-1
	 End If
  Next
  '����
  TaskXML.Save(Server.MapPath(KS.Setting(3)&"Config/task.xml"))
  KS.AlertHintScript "��ϲ,��ʱ������ɾ��!"
End Sub

Sub Task()
%>
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>
<title>��ʱ��������...�벻Ҫ�رձ�����!!!</title>
<link href='Include/Admin_Style.CSS' rel='stylesheet' type='text/css'>
<script language="JavaScript" src="../KS_Inc/Jquery.js"></script>
<script> 
$(document).ready(function(){ 
 Timer()
});
var itemLen=<%=TaskNode.length%>;
var taskItem = new Array();
var taskUrl = new Array();
var taskTime = new Array();
var taskStartType = new Array();
var taskWeek = new Array();
<%
N=0
For Each Node In TaskNode
  taskid=Node.SelectSingleNode("taskid").text
  Select Case Node.SelectSingleNode("tasktype").text
    case "0" TaskUrl="Include/RefreshIndex.asp?f=task" 
    case "1" TaskUrl="Collect/Collect_ItemCollection.asp?f=task&Action=Start&CollecType=1&itemid=" & taskid
    case "2" TaskUrl="Include/RefreshHtmlSave.Asp?f=task&Types=Folder&RefreshFlag=IDS&ID=" & taskid
    case "3" TaskUrl="Include/RefreshHtmlSave.asp?f=task&Types=Content&RefreshFlag=Folder&ChannelID="& Node.SelectSingleNode("channelid").text & "&FolderID=" & "'" & replace(taskid,",","','") & "'"
  End Select
  
  Response.Write "taskItem[" & n & "]='" & Node.SelectSingleNode("name").text &"';" &vbcrlf
  Response.Write "taskUrl[" & n & "]=""" & TaskUrl &""";" &vbcrlf
  Response.Write "taskTime[" & n & "]='" & Node.SelectSingleNode("time").text &"';" &vbcrlf
  Response.Write "taskStartType[" & n & "]='" & Node.SelectSingleNode("starttype").text &"';" &vbcrlf
  Response.Write "taskWeek[" & n & "]='" & Node.SelectSingleNode("week").text &"';" &vbcrlf
  N=N+1
Next

%>
function timeClock(){ 
	var today=new Date();
	var year =today.getYear();
	var month=today.getMonth()+1;
	var day=today.getDate();
	var h = today.getHours();
	var m = today.getMinutes();
	var s = today.getSeconds();
	var endTime=year+'-'+month+'-'+day+' '+h+":"+m+":"+s;
	$("#currTime").html(endTime);
	
	
	//���ʱ��
	for(var i=0;i<taskItem.length;i++)
	{
	   //����ʱ
	    var djs;
	    if (taskStartType[i]==1)
		{
		  djs=year+"/"+month+"/"+day+" "+taskTime[i];
		}else{
		  djs=year+"/"+month+"/"+day+" "+taskTime[i];
		}
	   
	    BirthDay=new Date(djs);//�ĳ���ļ�ʱ����
		today=new Date();
		timeold=(BirthDay.getTime()-today.getTime());
		sectimeold=timeold/1000
		secondsold=Math.floor(sectimeold);
		msPerDay=24*60*60*1000
		e_daysold=timeold/msPerDay
		daysold=Math.floor(e_daysold);
		e_hrsold=(e_daysold-daysold)*24;
		hrsold=Math.floor(e_hrsold);
		e_minsold=(e_hrsold-hrsold)*60;
		minsold=Math.floor((e_hrsold-hrsold)*60);
		seconds=Math.floor((e_minsold-minsold)*60);
		
	   var beginTime=today.getYear()+'-'+(today.getMonth()+1)+'-'+today.getDate()+' '+taskTime[i];
	   var ct= comptime(beginTime,endTime)
		if (taskStartType[i]==1){
		$("#lea"+i).html(hrsold+"Сʱ"+minsold+"��"+seconds+"��");
		}else if(taskStartType[i]==3){
		
		}
		else{
		  var leaday=taskWeek[i]-today.getDay();
		  if (ct>=0)
		  { if (leaday==0)
		     leaday=6;
			else
			 leaday=leaday-1;
		  }
		  if (leaday<0) leaday=leaday+6;
		$("#lea"+i).html(leaday+"��"+hrsold+"Сʱ"+minsold+"��"+seconds+"��");
		}

	   
	   
	   //���ִ��
	 if(taskStartType[i]==3){
	     var harr=taskTime[i].split(',');
		 for(var k=0;k<harr.length;k++)
		 {
		    var beginTime=today.getYear()+'-'+(today.getMonth()+1)+'-'+today.getDate()+' '+harr[k];
			var ct= comptime(beginTime,endTime)
			if (ct==0)
		    { 
			 window.open(taskUrl[i]); 
		    }
		 }
	  }else{
		   var beginTime=today.getYear()+'-'+(today.getMonth()+1)+'-'+today.getDate()+' '+taskTime[i];
		   var ct= comptime(beginTime,endTime)
		   if (ct==0)
		   { 
			 if (parseInt(taskStartType[i])==1 || (parseInt(taskStartType[i])==2 && today.getDay()==parseInt(taskWeek[i]))){
			 window.open(taskUrl[i]);
			 }
		   }
	  }
	}
} 

//ע�⣺��js��ʵ�ְ�ʱ���ñ��������ַ�ʽ������ʱ�����Լ�
function Timer()
{
timeClock();
setTimeout("Timer()", 1000); // ѭ����ʱ���� 
}

//�Ƚ�ʱ�� ��ʽ yyyy-mm-dd hh:mi:ss
function comptime(beginTime,endTime){ 
var beginTimes=beginTime.split(' ')[0].split('-');
var endTimes=endTime.split(' ')[0].split('-');
beginTime=beginTimes[1]+'-'+beginTimes[2]+'-'+beginTimes[0]+' '+beginTime.split(' ')[1];
endTime=endTimes[1]+'-'+endTimes[2]+'-'+endTimes[0]+' '+endTime.split(' ')[1];
var a =(Date.parse(endTime)-Date.parse(beginTime))/3600/1000;
return a;
}

</script>
</head>
<body>
<br/>
<br/>

  <table width='98%' align="center" BORDER='0' cellpadding='1'  cellspacing='1' class='ctable'>
		  <tr class='tdbg'>               
		   <td class='clefttitle' height="30" align='center'><strong>���� <font color=red><%=TaskNode.Length%></font> ����ʱ����,��ǰʱ����:<span id='currTime' style="color:green"></span></strong></td>
		  </tr>
  <table>
  <%
  N=0
  For Each Node In TaskNode
  %>
	  <table width='98%' style="table-layout:fixed;margin:4px" align="center" BORDER='0' cellpadding='1'  cellspacing='1' class='ctable'>
		  <tr class='tdbg'>               
		   <td class='clefttitle' height="30" width="100" align='right'><strong>��������:</strong></td>
		   <td><%=Node.SelectSingleNode("name").text%></td>
		  </tr>
		  <tr class='tdbg'>               
		   <td class='clefttitle'  width="100" align='right'><strong>ִ��ʱ��:</strong></td>
		   <td style="word-wrap:break-word;">
		   <font color=red>
		   <%
		    if Node.SelectSingleNode("starttype").text="1" then
			 response.write "ÿ��"
			ElseIf Node.SelectSingleNode("starttype").text="3" then
			 response.write "ָ������ʱ���"
			Else
		     response.write "ÿ�� "
			 Select Case  Node.SelectSingleNode("week").text
			  case 0 response.write "������"
			  case 1 response.write "����һ"
			  case 2 response.write "���ڶ�"
			  case 3 response.write "������"
			  case 4 response.write "������"
			  case 5 response.write "������"
			  case 6 response.write "������"
			 End Select
			End If
			
			 response.write " " & Node.SelectSingleNode("time").text
			%>
			ִ��
			</font>
			<%if Node.SelectSingleNode("starttype").text="3" then%>
			
			<%else%>
			 ��ִ��ʱ�仹ʣ:
			<%end if%><span id="lea<%=N%>" style='color:blue'></span>
			</td>
		  </tr>
	  </table>
  <%
    n=n+1
  Next
  %>
</body>
</html>
<%
End Sub


Set KS=Nothing
CloseConn
%>