<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Session.asp"-->
<%
Dim KS:Set KS=New PublicCls
Dim FolderID:FolderID=Request.QueryString("FolderID")
Dim SQL,I,RSC:Set RSC=Server.CreateObject("ADODB.RECORDSET")
 RSC.Open "Select ChannelID,ChannelName,ChannelTable,ItemName,BasicType From KS_Channel Where ChannelStatus=1 Order By ChannelID",Conn,1,1
 If Not RSC.Eof Then
	SQL=RSC.GetRows(-1)
 End If
RSC.Close:Set RSC=Nothing

%>
<html>
<head>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<script language="JavaScript" src="../../KS_Inc/jQuery.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�½�������ǩ</title></head>
<link href="admin_style.css" rel="stylesheet">
<script language="javascript">
function CheckForm(){
 frames["LabelShow"].CheckForm();
}
</script>
<style>
li{margin:0px;padding:0px;list-style-type:none}
.list{border-bottom:1px #83B5CD solid;background:url(../images/titlebg.png); height:36px; font-size:13px; color:#555;padding-left:10px;}
.list li{display:block;float:left;border:1px solid #DEEFFA;background-color:#F7FBFE;height:20px;line-height:20px;margin:1px;padding:2px}

.submenu {z-index:999;position:absolute;white-space : nowrap; margin:0 ;background:#fff; border:1px solid #DEEFFA;display:none;background:url(../images/portalbox_bg.gif) no-repeat;left:-10px;top:22px}
.submunu_popup {line-height:18px;text-align:left;padding:8px;}
.submunu_popup a{line-height:18px;}
.rl{position:relative;}


</style>
<body topmargin="0" leftmargin="0" scroll=no>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="25">
	  <div class="list">
	   <li><a href="Label/GetGenericList.asp?FolderID=<%=FolderID%>" Target="LabelShow">�����б�</a></li>
	   <li><a href="Label/GetSlide.asp?FolderID=<%=FolderID%>" Target="LabelShow">ͨ�ûõ�</a></li>
	   <li><a href="Label/GetRolls.asp?FolderID=<%=FolderID%>" Target="LabelShow">����ͼƬ</a></li>
	   <li><a href="Label/GetMarquee.asp?FolderID=<%=FolderID%>" Target="LabelShow">��������</a></li>
	   <li><a href="Label/GetNotRuleList.asp?FolderID=<%=FolderID%>" Target="LabelShow">������</a></li>
	   <li><a href="Label/GetCirClassList.asp?FolderID=<%=FolderID%>" Target="LabelShow">ѭ���б�</a></li>
	   <li><a href="Label/GetRelativeList.asp?FolderID=<%=FolderID%>" Target="LabelShow">�������</a></li>
	   <li><a href="Label/GetPageList.asp?FolderID=<%=FolderID%>" Target="LabelShow">�ռ���ҳ</a></li>
	  
	       
	      <span onMouseOut="$('#Menu_special').hide();">
	   	   <li class="rl" style="height:25px" onMouseOver="$('#Menu_special').show();"><a href="#">ר ��</a> <img src="../images/d.gif" align="absmiddle" /></li>
		   
		    <div class="submenu submunu_popup" id="Menu_special">
			<a href="Label/GetSpecialList.asp?FolderID=<%=FolderID%>" title="ר���б��ǩ" Target="LabelShow">ר���б��ǩ</a><br />
			<a href="Label/GetCirSpecialList.asp?FolderID=<%=FolderID%>" title="ѭ����ʾ����ר���ǩ" Target="LabelShow">ѭ����ʾ����ר���ǩ</a><br />
			<a href="Label/GetLastSpecialList.asp?FolderID=<%=FolderID%>" title="ѭ����ʾ����ר���ǩ" Target="LabelShow">��ҳ��ʾ�����µ�����ר���ǩ</a><br />
			</div>
		 </span>
		 
	      <span onMouseOut="$('#Menu_space').hide();">
	   	   <li class="rl" style="height:25px" onMouseOver="$('#Menu_space').show();"><a href="#">�� ��</a> <img src="../images/d.gif" align="absmiddle" /></li>
		   
		    <div class="submenu submunu_popup" id="Menu_space">
			<a href="Label/GetSpaceList.asp?FolderID=<%=FolderID%>" title="�ռ��Ż��б��ǩ" Target="LabelShow">�ռ��Ż��б��ǩ</a><br />
			<a href="Label/GetBlogInfoList.asp?FolderID=<%=FolderID%>" title="������־�б��ǩ" Target="LabelShow">������־�б��ǩ</a><br />
			<a href="Label/Getxclist.asp?FolderID=<%=FolderID%>" title="��������б��ǩ" Target="LabelShow">��������б��ǩ</a><br />
			<a href="Label/GetGrouplist.asp?FolderID=<%=FolderID%>" title="����Ȧ���б��ǩ" Target="LabelShow">����Ȧ���б��ǩ</a><br />
			</div>
		 </span>
		 
	      <span  onMouseOut="$('#Menu_other').hide();">
	   	   <li class="rl"  style="height:25px" onMouseOver="$('#Menu_other').show();"><a href="#">�� ��</a> <img src="../images/d.gif" align="absmiddle" /></li> 
		   
		    <div class="submenu submunu_popup" id="Menu_other">
			<a href="Label/GetLocation.asp?FolderID=<%=FolderID%>" title="��վλ�õ�����ǩ" Target="LabelShow">��վλ�õ�����ǩ</a><br />
			<a href="Label/GetAnnounceList.asp?FolderID=<%=FolderID%>" title="��վ�����б��ǩ" Target="LabelShow">��վ�����б��ǩ</a><br />
			<a href="Label/GetNavigation.asp?FolderID=<%=FolderID%>" title="��Ŀ(Ƶ��)�ܵ�����ǩ" Target="LabelShow">��Ŀ(Ƶ��)�ܵ�����ǩ</a><br />
			<a href="Label/GetLinkList.asp?FolderID=<%=FolderID%>" title="���������б��ǩ" Target="LabelShow">���������б��ǩ</a><br />
			<a href="Label/GetClubList.asp?FolderID=<%=FolderID%>" title="��̳���ӵ��ñ�ǩ" Target="LabelShow">��̳���ӵ��ñ�ǩ</a><br />
			</div>
		 </span>
		 
		 <span  onMouseOut="$('#Menu_ask').hide();">
	   	   <li class="rl" style="height:25px" onMouseOver="$('#Menu_ask').show();"><a href="#">�� ��</a> <img src="../images/d.gif" align="absmiddle" /></li>
		   
		    <div class="submenu submunu_popup" id="Menu_ask">
			<a href="Label/GetAQList.asp?FolderID=<%=FolderID%>" title="�������ʱ�ǩ" Target="LabelShow">�������ʱ�ǩ</a><br />
			<a href="Label/GetAAList.asp?FolderID=<%=FolderID%>" title="���»ش��ǩ" Target="LabelShow">���»ش��ǩ</a><br />
			</div>
		 </span>
		 
		 <%If KS.C_S(10,21)="1" Then%>
	      <span  onMouseOut="$('#Menu_job').hide();">
	   	   <li class="rl" style="height:25px" onMouseOver="$('#Menu_job').show();"><a href="#">�� Ƹ</a> <img src="../images/d.gif" align="absmiddle" /></li>
		   
		    <div class="submenu submunu_popup" id="Menu_job">
			<a href="Label/GetJobList.asp?FolderID=<%=FolderID%>" title="��Ƹְλ�б��ǩ" Target="LabelShow">��Ƹְλ�б��ǩ</a><br />
			<a href="Label/GetJobZWList.asp?FolderID=<%=FolderID%>" title="��ְλ�б��ǩ" Target="LabelShow">��ְλ�б��ǩ</a><br />
			<a href="Label/GetJobResumeList.asp?FolderID=<%=FolderID%>" title="�����б��ǩ" Target="LabelShow">�����б��ǩ</a><br />
			</div>
		 </span>
		 <%End If%>
		 
		 
	  </div>
	  <div style="clear:both"></div>
</td>
  </tr>
  <tr>
    <td valign="top">
	 <iframe name="LabelShow" ID="LabelShow" src="Label/GetGenericList.asp?PageTitle=�½�ͨ���б��ǩ&FolderID=<%=FolderID%>" style="width:100%;height:100%" frameborder="0" scrolling="auto"></iframe>	</td>
  </tr>
</table>
</body>
</html> 
