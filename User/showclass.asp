<!--#include file="../conn.asp"-->
<!--#include file="../ks_cls/kesion.membercls.asp"-->
var subsmallclassid = new Array();
<%
const tj=1    '从第几级算起
Dim KS:Set KS=new PublicCls
Dim KSUser:Set KSUser=New UserCls
If KSUser.UserLoginChecked=false Then   KS.Die ""
Dim SQL,K,Node,Pstr,Xml,ChannelID
ChannelID=KS.ChkClng(KS.S("ChannelID"))
KS.LoadClassConfig()
dim n:n=0
dim classid:classid=ks.s("classid")
if classid="" then classid="0"

If ChannelID<>0 Then Pstr="and @ks12=" & channelid & ""
 Set Xml=Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks14=1" & Pstr&"]")
 For Each Node In Xml
        If (Node.SelectSingleNode("@ks18").text=0) OR ((KS.FoundInArr(Node.SelectSingleNode("@ks17").text,KSUser.GroupID,",")=false and Node.SelectSingleNode("@ks18").text=3)) Then
		%>
		subsmallclassid[<%=n%>] = new Array('<%=Node.SelectSingleNode("@ks13").text%>','<%=Node.SelectSingleNode("@ks0").text%>','<%=Node.SelectSingleNode("@ks1").text%>',0,<%=Node.SelectSingleNode("@ks19").text%>)
		<%
	    Else
		%>
		subsmallclassid[<%=n%>] = new Array('<%=Node.SelectSingleNode("@ks13").text%>','<%=Node.SelectSingleNode("@ks0").text%>','<%=Node.SelectSingleNode("@ks1").text%>',1,<%=Node.SelectSingleNode("@ks19").text%>)
		<%
		End IF
         n=n+1
 Next

%>
function changesmallclassid(selectValue)
{
if (selectValue==0) return;
document.getElementById('smallerclassid').length = 0;   //点击一级栏目时,置三级下拉为空
document.getElementById('smallerclassid').options[0] = new Option('-选择-','0');

document.getElementById('smallclassid').length = 0;
document.getElementById('smallclassid').options[0] = new Option('-选择-','0');

document.getElementById('ClassID').value='0';

	  document.getElementById('smallclassid').style.display='';
	  document.getElementById('smallerclassid').style.display='';


for (i=0; i<subsmallclassid.length; i++)
{
    if (subsmallclassid[i][1] == selectValue && subsmallclassid[i][4]==0){  //只有一级的情况
	  document.getElementById('ClassID').value=selectValue; 
	  document.getElementById('smallclassid').style.display='none';
	  document.getElementById('smallerclassid').style.display='none';
	  return;
	}else if (subsmallclassid[i][0] == selectValue)
	{
	     //判断有没有下级允许投稿
		 var xjtk=false;
		 for(j=0;j< subsmallclassid.length; j++)
		 {
		    if (subsmallclassid[j][0]==subsmallclassid[i][1]){
			  if (subsmallclassid[j][3]==1){
			    xjtk=true;
				break;
			  }
			}
		 }
	     if (subsmallclassid[i][3] == 1 || xjtk ){
			document.getElementById('smallclassid').options[document.getElementById('smallclassid').length] = new Option(subsmallclassid[i][2], subsmallclassid[i][1]);
		 }
		 
		 //判断是否显示三级下拉列表
		 var showxj=false;
		 for(j=0;j< subsmallclassid.length; j++){
		    if (subsmallclassid[j][0]==selectValue){
			   if (parseInt(subsmallclassid[j][4])>0){
			    showxj=true;
				break;
			   }
			}
		 }
		 if (showxj==true){
		 document.getElementById('smallerclassid').style.display='';
		 }else{
		 document.getElementById('smallerclassid').style.display='none';
		 }
		 
	}
}
}
function changesmallerclassid(selectValue)
{
if (selectValue=='0') document.getElementById('ClassID').value='0';
//判断是否显示三级下拉列表
for (i=0; i<subsmallclassid.length; i++){
     if (subsmallclassid[i][1]==selectValue){
	  	  if (subsmallclassid[i][4]==0){
		     document.getElementById('ClassID').value=selectValue;
		     document.getElementById('smallerclassid').style.display='none';
		  }else{
		     document.getElementById('ClassID').value='0';
			 document.getElementById('smallerclassid').style.display='';
		}
	 }
}

document.getElementById('smallerclassid').length = 0;
document.getElementById('smallerclassid').options[0] = new Option('-选择-','0');
for (i=0; i<subsmallclassid.length; i++)
{

	if (subsmallclassid[i][0] == selectValue)
	{

		if (subsmallclassid[i][3] == 1){
		document.getElementById('smallerclassid').options[document.getElementById('smallerclassid').length] = new Option(subsmallclassid[i][2], subsmallclassid[i][1]);
		}
		
	}
}
}
function getclassid(selectValue){
document.getElementById('ClassID').value=selectValue;
}

document.write ("<select name='bigclassid' id='bigclassid' style='width:120px' size='1' onChange='changesmallclassid(this.value)'>");
document.write ("<option value='0' selected>-选择-</option>");
<%
 Set Xml=Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks14=1" & Pstr&" and @ks10=" & tj & "]")
 For Each Node In Xml
  If ((Node.SelectSingleNode("@ks18").text=0) OR ((KS.FoundInArr(Node.SelectSingleNode("@ks17").text,KSUser.GroupID,",")=false and Node.SelectSingleNode("@ks18").text=3))) and checkxjtk(Node.SelectSingleNode("@ks0").text)=false Then
  Else%>
document.write ("<option value=<%=Node.SelectSingleNode("@ks0").text%>><%=Node.SelectSingleNode("@ks1").text%></option>");
<%
  End If
 Next
%>
document.write ("</select>")

document.write ("  <select name='smallclassid' size='1' onChange='changesmallerclassid(this.value)' style='width:120px' id='smallclassid'>");
document.write ("<option value='0' selected>-选择-</option>");
document.write ("</select>")
document.write ("  <select name='smallerclassid' size='1' style='display:none;width:120px' id='smallerclassid' onChange='getclassid(this.value)'>");
document.write ("<option value='0' selected>-选择-</option>");
document.write ("</select>");
document.write ("<input type='hidden' name='ClassID' value='<%=classid%>' id='ClassID'/>");
<%

'默认值
If ClassID<>"0" Then
 If KS.C_C(ClassID,10)-tj=0 Then   '一级
 %>
 	document.getElementById('bigclassid').value='<%=ClassID%>';
    document.getElementById('smallclassid').style.display='none';
    document.getElementById('smallerclassid').style.display='none';
 <%
 ElseIf KS.C_C(ClassID,10)-tj=1 Then   '二级
 %>
    document.getElementById('smallclassid').style.display='';
	setSecoundOption('<%=KS.C_C(ClassID,13)%>','<%=ClassID%>');
	document.getElementById('bigclassid').value='<%=KS.C_C(ClassID,13)%>';
 <%
 ElseIf KS.C_C(ClassID,10)-tj=2 Then   '三级
   %>
    document.getElementById('smallerclassid').style.display='';
	for (i=0; i<subsmallclassid.length; i++){
	   //给三级下拉指定值
      if (subsmallclassid[i][0]=='<%=KS.C_C(ClassID,13)%>'){
		if (subsmallclassid[i][3] == 1){
		document.getElementById('smallerclassid').options[document.getElementById('smallerclassid').length] = new Option(subsmallclassid[i][2], subsmallclassid[i][1]);
		}
	   }
	   //得二级下拉的ParentID
	   if (subsmallclassid[i][1]=='<%=KS.C_C(ClassID,13)%>'){
	    pid=subsmallclassid[i][0];
	   }
    }
	document.getElementById('smallerclassid').value='<%=ClassID%>';
	
	//给二级下拉指定值
	setSecoundOption(pid,'<%=KS.C_C(ClassID,13)%>');
	document.getElementById('bigclassid').value=pid;
   <%
 End If
 %>
<%End If
Set KS=Nothing
Set KSUser=Nothing
CloseConn
%>

//给二级下拉填充值 参数pid 父栏目ID, sid 选中的栏目ID
function setSecoundOption(pid,sid)
{
	//给二级下拉指定值
	for (i=0; i<subsmallclassid.length; i++){
	if (subsmallclassid[i][0] == pid)
	{
	     //判断有没有下级允许投稿
		 var xjtk=false;
		 for(j=0;j< subsmallclassid.length; j++)
		 {
		    if (subsmallclassid[j][0]==subsmallclassid[i][1]){
			  if (subsmallclassid[j][3]==1){
			    xjtk=true;
				break;
			  }
			}
		 }
	     if (subsmallclassid[i][3] == 1 || xjtk ){
			document.getElementById('smallclassid').options[document.getElementById('smallclassid').length] = new Option(subsmallclassid[i][2], subsmallclassid[i][1]);
		 }
		 
	 }
	}
	document.getElementById('smallclassid').value=sid;

}
<%
'检查栏目ID检查下级有没有允许投稿的栏目
function checkxjtk(id)
     Dim Xml,Node
	 Set Xml=Application(KS.SiteSN&"_class").DocumentElement.SelectNodes("class[@ks14=1" & Pstr&" and @ks10>" & tj & "]")
	 For Each Node In Xml
	   If KS.FoundInArr(Node.SelectSingleNode("@ks8").text,id,",")=true Then  '如果是他的下级
		  If ((Node.SelectSingleNode("@ks18").text=0) OR ((KS.FoundInArr(Node.SelectSingleNode("@ks17").text,KSUser.GroupID,",")=false and Node.SelectSingleNode("@ks18").text=3)))Then
		  Else
		   checkxjtk=true
		   exit function
		  End If
	   End If
	Next

  checkxjtk=false
end function
%>