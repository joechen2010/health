<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="../include/Session.asp"-->
<%
Dim Chk:Set Chk=New LoginCheckCls1
Chk.Run()
Set Chk=Nothing


Dim FolderRS,SqlStr
Dim SChannelID:SChannelID=Request("SChannelID")   'SchannelID=9999��������ɱ�ǩ/JS����
Dim TemplateType:TemplateType=Request("TemplateType")
Dim KS,KSCls,SQL,K,i,DIYFieldArr,F_B,F_V
On Error Resume Next
Set KS=New PublicCls
Set KSCls=New ManageCls
Dim DomainStr:DomainStr=KS.GetDomain
Dim RS:Set RS=Conn.Execute("Select ChannelID,BasicType,ChannelName,ItemName,ItemUnit,FieldBit,ModelEname From KS_Channel Where ChannelStatus=1 and basictype in(1,2,3,5) Order By ChannelID")
SQL=RS.GetRows(-1)
RS.Close:Set RS=Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<script language="JavaScript" src="../../KS_Inc/Common.js"></script>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��ǩ����</title>
<style type="text/css">
a{text-decoration: none;} /* �������»���,��Ϊunderline */ 
a:link {color: #000000;} /* δ���ʵ����� */
a:visited {color: #000000;} /* �ѷ��ʵ����� */
a:hover{color: #FF0000;text-decoration: underline;} /* ����������� */ 
a:active {color: #FF0000;} /* ����������� */
td	{font-family:  "Verdana, Arial, Helvetica, sans-serif"; font-size: 11.5px; color: #000000; text-decoration:none ; text-decoration:none ; }
BODY {
font-family:  "Verdana, Arial, Helvetica, sans-serif"; font-size: 11.5px;
FONT-SIZE: 9pt;
color: #000000;
text-decoration: none;
}
li{
list-style:none;
list-style-image:url(Images/label0.gif);
margin-left:20px;
margin-bottom:2px;
}
</style>
</head>
<body topmargin="0" leftmargin="0">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td height="25"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="22" align="center" bgcolor="#0000FF"><strong><font color="#FFFFFF">��վϵͳ---��ǩ�б�</font></strong></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td valign="top"> 

	 
              
                    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr ParentID=""> 
              <td> 
			   <table width="100%" border="0" cellpadding="0" cellspacing="0">
               <tr onClick="ShowLabelTree('wapchangyongfenlei')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                <td width="24"><img src="Images/folderclosed.gif" width="24" height="22"></td>
                <td width="1227"><a href="#">WAP���ñ�ǩ����</a></td>
               </tr>
               <tr> 
                <td colspan="2">
			      <table width="85%" align='center' border="0" cellspacing="0" cellpadding="0" id="wapchangyongfenlei" style="display:none">
                    <tr> 
					 <td>
                    <div onClick="ShowLabelTree('wapchangyong')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'">
                    <img src="Images/folderclosed.gif" align="absmiddle" width="24" height="22"><a href="#">WAP���ñ�ǩ</a></div></div>
                    <div  id="wapchangyong" style="display:none">
			  		<li><a href="#" onClick="InsertLabel('{$GetSiteTitle}');">��ʾ��վ����</a></li>
                    <li><a href="#" onClick="InsertLabel('{$GetSiteName}');">��ʾ��վ����</a></li>
                    <li><a href="#" onClick="InsertFunctionLabel('AddDate.html',250,105);">��ʾ��ǰʱ��</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetSiteLogo}');">��ʾ��վLogo(��������)</a></li>
					<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>KS_Editor/KS_Logo.html',250,130);">��ʾ��վLogo(������)</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetTopUserLogin}');" class="LabelItem">��ʾ��Ա��¼���(����)</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetUserLogin}');" class="LabelItem">��ʾ��Ա��¼���(����)</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetSiteCountAll}');">��ʾ��վ��Ϣͳ��</a></li>
                    <li><a href="#" onClick="InsertLabel('{$GetOnlineTotal}');">��ʾ����������</a></li>
                    <li><a href="#" onClick="InsertLabel('{$GetOnlineUser}');">��ʾ�����û�����</a></li>
                    <li><a href="#" onClick="InsertLabel('{$GetOnlineGuest}');">��ʾ�����ο�����</a></li>
					<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>KS_Editor/KS_TopUser.html',250,130);">��ʾ�û���¼����</a></li>
                    <li><a href="#" onClick="InsertLabel('{$GetSearch}');">��ʾ����</a></li>
                    <li><a href="#" onClick="InsertLabel('{$GetFriendLink}');">��ʾ�����������</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetSiteUrl}');">��ʾ��վURL</a></li>
                    <li><a href="#" onClick="InsertLabel('{$GetCopyRight}');">��ʾ��Ȩ��Ϣ</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetWebmaster}');">��ʾվ��</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetWebmasterEmail}');">��ʾվ��EMail</a></li>
                    <li><a href="#" onClick="InsertLabel('{$GetGoBack}');">��ʾ�����ϼ�</a></li>
                    <li><a href="#" onClick="InsertFunctionLabel('GetPourAccount.html',250,170);">��ʾ�ض�����ʱ</a></li>
                    <li><a href="#" onClick="InsertFunctionLabel('GetWenhouyu.html',280,270);">��ʾ���ݵ�ǰ��ʱ�䲻ͬ���ʺ���</a></li>
                    <li><a href="#" onClick="InsertLabel('{$GetGoBackIndex}');">��ʾ������ҳ</a></li>
                    <li><a href="#" onClick="InsertLabel('{$GetLocation}');">��ʾ��Ŀλ�õ���</a></li>
                    <li><a href="#" onClick="InsertLabel('{$GetGoBackChannel}');">��ʾ��Ŀ�б�ҳ�����ϼ�</a></li>
                    <li><a href="#" onClick="InsertLabel('{$GetGoBackClass}');">��ʾ����ҳ������Ŀ�б�</a></li>
                    <li><a href="#" onClick="InsertFunctionLabel('Wap_GetWriteinReturn.html',250,130);">���뷵�ص�ַ��������</a></li>
                    <li><a href="#" onClick="InsertLabel('{$GetReadReturn}');">'��ʾ���ص�ַ���泬����</a></li>
                    <li><a href="#" onClick="InsertFunctionLabel('GetHTTPCollect.html',450,450);">��ʾԶ����ҳ����(С͵)</a></li>
                    </div>
                    
                    
                    <div onClick="ShowLabelTree('wapshouye')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'">
                    <img src="Images/folderclosed.gif" align="absmiddle" width="24" height="22"><a href="#">WAPȫվͨ�ñ�ǩ</a></div>
                    <div  id="wapshouye" style="display:none">
                    <li><a href="#" onClick="InsertFunctionLabel('GetIndexChannel.html',450,140);">��ʾƵ������</a></li>
                    <li><a href="#" onClick="InsertFunctionLabel('GetIndexList.asp',450,230);">��ʾ�����б�(����,����,�Ƽ�,���)</a></li>
                    <li><a href="#" onClick="InsertFunctionLabel('GetClubList.asp',450,190);">��ʾ��̳�����б�</a></li>
                    <li><a href="#" onClick="InsertFunctionLabel('GetLogList.asp',450,190);">��ʾ��־�б�</a></li>
                    </div>
                
                    <div onClick="ShowLabelTree('waplanmu')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'">
                    <img src="Images/folderclosed.gif" align="absmiddle" width="24" height="22"><a href="#">WAP��ǰ��Ŀͨ�ñ�ǩ</a></div>
                    <div  id="waplanmu" style="display:none">
                    <li><a href="#" onClick="InsertFunctionLabel('GetShowClassList.html',450,230);">��ʾ��Ŀ�����б�(����,����,�Ƽ�,���)</a></li>
                    <li><a href="#" onClick="InsertFunctionLabel('GetRandomPhotoText.html',250,250);">��ʾ��Ŀ���ͼ��</a></li>
                    <li><a href="#" onClick="InsertFunctionLabel('GetClassList.html',450,170);">��ʾƵ����Ŀ����</a></li>
                    <li><a href="#" onClick="InsertFunctionLabel('GetShowClassCent.html',450,240);">��ʾ�ռ���ҳ�б�</a></li>
                    </div>
                    
					 </td>
				    </tr>
				  </table>
				</td>
			  </tr>
			  </table>
			  </td>
		   </tr>
	  </table>
      
          
            <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr ParentID=""> 
              <td> 
			   <table width="100%" border="0" cellpadding="0" cellspacing="0">
               <tr onClick="ShowLabelTree('wapchangyongLabel')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                <td width="24"><img src="Images/folderclosed.gif" width="24" height="22"></td>
                <td width="1227"><a href="#">WAP����ҳ��ǩ</a></td>
               </tr>
               <tr> 
                <td colspan="2">
			      <table width="85%" align='center' border="0" cellspacing="0" cellpadding="0" id="wapchangyongLabel" style="display:none">
                    <tr> 
					 <td>
					 <%
					  For K=0 To Ubound(SQL,2)
					   F_B=Split(Split(SQL(5,K),"@@@")(0),"|")
					   F_V=Split(Split(SQL(5,K),"@@@")(1),"|")
					  %>
					 	<div onClick="ShowLabelTree('wapneirong<%=SQL(6,K)%>')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'">
					<img src="Images/folderclosed.gif" align="absmiddle" width="24" height="22"><a href="#">Wap<%=SQL(3,K)%>����ҳ��ǩ</a>
				     </div>	
					 <%Select Case SQL(1,K)%>
					  <%case 1%>
					  <div  id="wapneirong<%=SQL(6,K)%>" style="display:none">
						<li><a href="#" onClick="InsertLabel('{$GetArticleID}');" class="LabelItem">��ǰ<%=SQL(3,K)%>ID</a></li>
						<li><a href="#" onClick="InsertLabel('{$InfoID}');" class="LabelItem">��ǰ<%=SQL(3,K)%>СID</a></li>
						<li><a href="#" onClick="InsertLabel('{$ItemName}');" class="LabelItem">��ǰ��Ŀ����</a></li>
						<li><a href="#" onClick="InsertLabel('{$ItemUnit}');" class="LabelItem">��ǰ��Ŀ��λ</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetKeyTags}');" class="LabelItem">ȡ��<%=SQL(3,K)%>Tags</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetArticleShortTitle}');" class="LabelItem"><%=SQL(3,K)%>��̱���</a></li>
						<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>KS_Editor/KS_ArtPhoto.html',250,130);" class="LabelItem">����ҳͼƬ</a></li>

						<%if F_B(8)=1 Then%>
						<li><a href="#" onClick="InsertLabel('{$GetArticleIntro}');" class="LabelItem"><%=F_V(8)%></a></li>
						<%end if%>
						<%if F_B(9)=1 Then%>
                        <li><a href="#" onClick="InsertFunctionLabel('GetArticleContent.html',250,205);"><%=F_V(9)%></a></li>
						<%end if%>
						<%if F_B(6)=1 Then%>
						<li><a href="#" onClick="InsertLabel('{$GetArticleAuthor}');"><%=F_V(6)%></a></li>
						<%end if%>
						<%if F_B(7)=1 Then%>
						<li><a href="#" onClick="InsertLabel('{$GetArticleOrigin}');"><%=F_V(7)%></a></li>
						<%end if%>
						<%if F_B(12)=1 Then%>
						<li><a href="#" onClick="InsertLabel('{$GetArticleDate}');"><%=F_V(12)%></a></li>
						<%end if%>
						<%if F_B(14)=1 Then%>
						<li><a href="#" onClick="InsertLabel('{$GetArticleHits}');"><%=F_V(14)%></a></li>
						<%end if%>
						<li><a href="#" onClick="InsertLabel('{$GetArticleInput}');"><%=SQL(3,K)%>¼��</a></li>
						<%if F_B(3)=1 Then%>
						<li><a href="#" onClick="InsertLabel('{$GetArticleProperty}');">��ʾ<%=SQL(3,K)%>������(���š��Ƽ���...)</a></li>
						<%end if%>
                        <li><a href="#" onClick="InsertLabel('{$GetDigg}');" class="LabelItem">��ʾ��һ��</a> </li>
                        <li><a href="#" onClick="InsertLabel('{$GetFavorite}');">��ʾ��Ҫ�ղ�</a></li>
                        <li><a href="#" onClick="InsertLabel('{$GetComment}');">��ʾ������������</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetPrevArticle}');">��ʾ��һ<%=SQL(4,K)%><%=SQL(3,K)%></a></li>
						<li><a href="#" onClick="InsertLabel('{$GetNextArticle}');">��ʾ��һ<%=SQL(4,K)%><%=SQL(3,K)%></a></li>
						<li><a href="#" onClick="InsertFunctionLabel('GetShowComment.html',450,140);">��ʾ����</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetWriteComment}');">��������</a></li>
                        <li><a href="#" onClick="InsertFunctionLabel('GetRandomContentsList.html',450,170);">��ʾ<%=SQL(3,K)%>����б�</a></li>
						<li><a href="#" onClick="InsertFunctionLabel('GetRelatedContentsList.html',450,170);">��ʾ<%=SQL(3,K)%>����б�</a></li>
                     <%Case 2%>					  
					  <div id="wapneirong<%=SQL(6,K)%>" style="display:none">
						<li><a href="#" onClick="InsertLabel('{$GetPictureID}');" class="LabelItem">��ǰ<%=SQL(3,K)%>ID</a></li>
						<li><a href="#" onClick="InsertLabel('{$InfoID}');" class="LabelItem">��ǰ<%=SQL(3,K)%>СID</a></li>
						<li><a href="#" onClick="InsertLabel('{$ItemName}');" class="LabelItem">��ǰ��Ŀ����</a></li>
						<li><a href="#" onClick="InsertLabel('{$ItemUnit}');" class="LabelItem">��ǰ��Ŀ��λ</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetPictureName}');" class="LabelItem"><%=F_V(0)%></a></li>
                        <li><a href="#" onClick="InsertLabel('{$GetKeyTags}');" class="LabelItem">ȡ��<%=SQL(3,K)%>Tags</a></li>
                        <li><a href="#" onClick="InsertLabel('{$GetPictureByPage}');" class="LabelItem">�鿴<%=SQL(3,K)%>����(��һҳ����һҳ��ʽ)</a></li>
						
						 <%if F_B(9)=1 Then%>
						 <li><a href="#" onClick="InsertLabel('{$GetPictureIntro}');" class="LabelItem"><%=F_V(9)%></a></li>
						 <%end if%>
						 <%if F_B(7)=1 Then%>
						 <li><a href="#" onClick="InsertLabel('{$GetPictureAuthor}');" class="LabelItem"><%=F_V(7)%></a></li>
						 <%end if%>
						 <%if F_B(8)=1 Then%>
						 <li><a href="#" onClick="InsertLabel('{$GetPictureOrigin}');" class="LabelItem"><%=F_V(8)%></a></li>
						 <%end if%>
						 <%if F_B(10)=1 Then%>
						 <li><a href="#" onClick="InsertLabel('{$GetPictureDate}');" class="LabelItem"><%=F_V(10)%></a></li>
						 <%end if%>
						 <li><a href="#" onClick="InsertLabel('{$GetPictureHits}');" class="LabelItem"><%=SQL(3,K)%>����(�������)</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetPictureHitsByDay}');" class="LabelItem"><%=SQL(3,K)%>���������</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetPictureHitsByWeek}');" class="LabelItem"><%=SQL(3,K)%>���������</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetPictureHitsByMonth}');" class="LabelItem"><%=SQL(3,K)%>���������</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetPictureInput}');" class="LabelItem"><%=SQL(3,K)%>¼��</a></li>
						 <%if F_B(5)=1 Then%>
						 <li><a href="#" onClick="InsertLabel('{$GetPictureProperty}');" class="LabelItem">��ʾ<%=SQL(3,K)%>����(���š��������Ƽ�...</a></li>
						 <%end if%>
                         <li><a href="#" onClick="InsertLabel('{$GetFavorite}');">��ʾ��Ҫ�ղ�</a></li>
                         <li><a href="#" onClick="InsertLabel('{$GetComment}');">��ʾ������������</a></li>
						 <li><a href="#" onClick="InsertFunctionLabel('GetShowComment.html',450,140);">��ʾ����</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetWriteComment}');">��ʾ��������</a></li>
                         <li><a href="#" onClick="InsertFunctionLabel('GetRandomContentsList.html',450,170);">��ʾ����ҳ����б�</a></li>
						 <li><a href="#" onClick="InsertFunctionLabel('GetRelatedContentsList.html',450,170);">��ʾ����ҳ����б�</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetPictureVote}');" class="LabelItem">��ʾͶ��һƱ</a> </li>
                         <li><a href="#" onClick="InsertLabel('{$GetDigg}');" class="LabelItem">��ʾ��һ��</a> </li>
						 <!--<li><a href="#" onClick="InsertLabel('{$GetPictureVoteScore}');" class="LabelItem">��ʾ<%=SQL(3,K)%>��Ʊ��</a></li>-->
  						 <li><a href="#" onClick="InsertLabel('{$GetPrevPicture}');" class="LabelItem">��ʾ��һ��<%=SQL(3,K)%></a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetNextPicture}');" class="LabelItem">��ʾ��һ��<%=SQL(3,K)%></a></li>
                         <li><a href="#" onClick="InsertFunctionLabel('Wap_GetRelatedRowform.html',450,170);">��ʾ<%=SQL(3,K)%>����б�</a></li>
                         <li><a href="#" onClick="InsertFunctionLabel('Wap_GetContentsRandom.html',450,170);">��ʾ<%=SQL(3,K)%>����б�</a></li>

				<%Case 3%>
				 <div id="wapneirong<%=SQL(6,K)%>" style="display:none">
                     <li><a href="#" onClick="InsertLabel('{$GetDownID}');" class="LabelItem">��ǰ<%=SQL(3,K)%>ID</a></li>
					 <li><a href="#" onClick="InsertLabel('{$InfoID}');" class="LabelItem">��ǰ<%=SQL(3,K)%>СID</a></li>
					 <li><a href="#" onClick="InsertLabel('{$ItemName}');" class="LabelItem">��ǰ��Ŀ����</a></li>
					 <li><a href="#" onClick="InsertLabel('{$ItemUnit}');" class="LabelItem">��ǰ��Ŀ��λ</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetDownTitle}');" class="LabelItem"><%=SQL(3,K)%>����+�汾��</a></li>
                     <li><a href="#" onClick="InsertLabel('{$GetKeyTags}');" class="LabelItem">ȡ��<%=SQL(3,K)%>Tags</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetDownAddress}');" class="LabelItem">��ͨ���ص�ַ</a></li>
                     <li><a href="#" onClick="InsertLabel('{$GetDownFenji}');" class="LabelItem">�ֻ������ص�ַ</a></li>
					 <%if F_B(8)=1 Then%>
					 <li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>KS_Editor/KS_DownPhoto.html',250,130);" class="LabelItem"><%=F_V(8)%></a></li>
					 <%end if%>
					 
					 <li><a href="#" onClick="InsertLabel('{$GetDownSize}');" class="LabelItem">�ļ���С+MB(KB)</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetDownLanguage}');" class="LabelItem"><%=SQL(3,K)%>����</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetDownType}');" class="LabelItem"><%=SQL(3,K)%>���</a></li>
					 <%if F_B(7)=1 Then%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownSystem}');" class="LabelItem"><%=F_V(7)%></a></li>
					 <%end if%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownPower}');" class="LabelItem">��Ȩ��ʽ</a></li>
					 <%if F_B(14)=1 Then%>
                     <li><a href="#" onClick="InsertFunctionLabel('GetContentIntro.html',250,105);"><%=F_V(14)%></a></li>
					 <%end if%>
					 <%if F_B(11)=1 Then%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownAuthor}');" class="LabelItem"><%=F_V(11)%></a></li>
					 <%end if%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownInput}');" class="LabelItem"><%=SQL(3,K)%>¼��</a></li>
					 <%if F_B(12)=1 Then%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownOrigin}');" class="LabelItem">�� Դ</a></li>
					 <%end if%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownDate}');" class="LabelItem">���(����)����</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetDownHits}');" class="LabelItem">�����ص����</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetDownHitsByDay}');" class="LabelItem">���յ����</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetDownHitsByWeek}');" class="LabelItem">���ܵ����</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetDownHitsByMonth}');" class="LabelItem">���µ����</a></li>
					 <%if F_B(15)=1 Then%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownYSDZ}');" class="LabelItem"><%=F_V(15)%></a></li>
					 <%end if%>
					 <%if F_B(16)=1 Then%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownZCDZ}');" class="LabelItem"><%=F_V(106)%></a></li>
					 <%end if%>
					 <%if F_B(17)=1 Then%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownDecPass}');" class="LabelItem"><%=F_V(17)%></a></li>
					 <%end if%>
                     <li><a href="#" onClick="InsertLabel('{$GetDownProperty}');" class="LabelItem">��ʾ<%=SQL(3,K)%>����(���š��Ƽ��ȣ�</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetDownStar}');" class="LabelItem">��ʾ�Ƽ��Ǽ�</a></li>
                     <li><a href="#" onClick="InsertLabel('{$GetDigg}');" class="LabelItem">��ʾ��һ��</a> </li>
                     <li><a href="#" onClick="InsertLabel('{$GetFavorite}');">��ʾ��Ҫ�ղ�</a></li>
                     <li><a href="#" onClick="InsertLabel('{$GetComment}');">��ʾ������������</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetPrevDown}');" class="LabelItem">��ʾ��һ��<%=SQL(3,K)%></a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetNextDown}');" class="LabelItem">��ʾ��һ��<%=SQL(3,K)%></a><?
                     <li><a href="#" onClick="InsertFunctionLabel('GetShowComment.html',450,140);">��ʾ����</a></li>
                     <li><a href="#" onClick="InsertLabel('{$GetWriteComment}');">��������</a></li>
                     <li><a href="#" onClick="InsertFunctionLabel('GetRandomContentsList.html',450,170);">��ʾ<%=SQL(3,K)%>����б�</a></li>
                     <li><a href="#" onClick="InsertFunctionLabel('GetRelatedContentsList.html',450,170);">��ʾ<%=SQL(3,K)%>����б�</a></li>
			       <%Case 5%>
				<div id="wapneirong<%=SQL(6,K)%>" style="display:none">
					<li><a href="#" onClick="InsertLabel('{$GetProductID}');" class="LabelItem">��ǰ<%=SQL(3,K)%>���(ID)</a></li>
                    <li><a href="#" onClick="InsertLabel('{$ChannelID}');" class="LabelItem">��ǰģ��ID</a></li>
                    <li><a href="#" onClick="InsertLabel('{$InfoID}');" class="LabelItem">��ǰ<%=SQL(3,K)%>СID</a></li>
                    <li><a href="#" onClick="InsertLabel('{$ItemName}');" class="LabelItem">��ǰ��Ŀ����</a></li>
                    <li><a href="#" onClick="InsertLabel('{$ItemUnit}');" class="LabelItem">��ǰ��Ŀ��λ</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetProductName}');" class="LabelItem"><%=SQL(3,K)%>����</a></li>
                    <li><a href="#" onClick="InsertLabel('{$GetKeyTags}');" class="LabelItem">ȡ��<%=SQL(3,K)%>Tags</a></li>
					<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>KS_Editor/KS_ProductPhoto.html',250,130);" class="LabelItem">��ƷͼƬ</a> </li>
                    <li><a href="#" onClick="InsertFunctionLabel('GetContentIntro.html',250,105);"><%=SQL(3,K)%>���</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetProducerName}');" class="LabelItem">�� �� ��</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetTrademarkName}');" class="LabelItem">Ʒ���̱�</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetProductModel}');" class="LabelItem"><%=SQL(3,K)%>�ͺ�</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetProductSpecificat}');" class="LabelItem"><%=SQL(3,K)%>���</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetProductDate}');" class="LabelItem">�ϼ�ʱ��</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetServiceTerm}');" class="LabelItem">��������</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetTotalNum}');" class="LabelItem">�������</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetProductUnit}');" class="LabelItem"><%=SQL(3,K)%>��λ</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetProductHits}');" class="LabelItem">�������</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetProductType}');" class="LabelItem">��������</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetRank}');" class="LabelItem">�Ƽ��ȼ�</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetProductProperty}');" class="LabelItem">��ʾ<%=SQL(3,K)%>����(�������ؼۡ��Ƽ��ȣ�</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetPrice_Original}');">&nbsp;��ʾԭʼ���ۼ�</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetPrice}');" class="LabelItem">��ʾ��ǰ���ۼ�</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetPrice_Member}');" class="LabelItem">��ʾ��Ա��</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetPrice_Market}');" class="LabelItem">��ʾ�г���</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetDiscount}');" class="LabelItem">��ʾ�ۿ���</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetScore}');" class="LabelItem">��ʾ�������</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetAddCar}');" class="LabelItem">���빺�ﳵ</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetFavorite}');">��ʾ��Ҫ�ղ�</a></li>
                    <li><a href="#" onClick="InsertLabel('{$GetComment}');">��ʾ������������</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetPrevProduct}');" class="LabelItem">��ʾ��һ��<%=SQL(3,K)%></a></li>
					<li><a href="#" onClick="InsertLabel('{$GetNextProduct}');" class="LabelItem">��ʾ��һ��<%=SQL(3,K)%></a></li>
                     <li><a href="#" onClick="InsertFunctionLabel('GetShowComment.html',450,140);">��ʾ����</a></li>
                     <li><a href="#" onClick="InsertLabel('{$GetWriteComment}');">��������</a></li>
                     <li><a href="#" onClick="InsertFunctionLabel('GetRandomContentsList.html',450,170);">��ʾ<%=SQL(3,K)%>����б�</a></li>
                     <li><a href="#" onClick="InsertFunctionLabel('GetRelatedContentsList.html',450,170);">��ʾ<%=SQL(3,K)%>����б�</a></li>
         <%Case 8%>
				<div id="wapneirong<%=SQL(6,K)%>" style="display:none">
				  <!--<li><a href="#" onClick="InsertLabel('{$GetGQInfoUrl}');" class="LabelItem">��ǰ<%=SQL(3,K)%> URL</a></li>-->
				  <li><a href="#" onClick="InsertLabel('{$GetGQInfoID}');" class="LabelItem">��ǰ<%=SQL(3,K)%>ID</a></li>
                  <li><a href="#" onClick="InsertLabel('{$ChannelID}');" class="LabelItem">��ǰģ��ID</a></li>
                  <li><a href="#" onClick="InsertLabel('{$InfoID}');" class="LabelItem">��ǰ<%=SQL(3,K)%>СID</a></li>
                  <li><a href="#" onClick="InsertLabel('{$ItemName}');" class="LabelItem">��ǰ��Ŀ����</a></li>
                  <li><a href="#" onClick="InsertLabel('{$ItemUnit}');" class="LabelItem">��ǰ��Ŀ��λ</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetGQTitle}');" class="LabelItem"><%=SQL(3,K)%>����</a></li>
				  <li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>KS_Editor/KS_SupplyPhoto.html',250,130);" class="LabelItem"><%=SQL(3,K)%>����ͼ(������)</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetGQKeyWords}');" class="LabelItem">ȡ�ùؼ���</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetKeyTags}');" class="LabelItem">ȡ��<%=SQL(3,K)%>Tags</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetPrice}');" class="LabelItem">�۸�˵��</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetInfoType}');" class="LabelItem">��Ϣ���</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetTransType}');" class="LabelItem">�������</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetValidTime}');" class="LabelItem">�� Ч ��</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetGQContent}');" class="LabelItem"><%=SQL(3,K)%>���ݽ���</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetGQHits}');" class="LabelItem"><%=SQL(3,K)%>�������</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetAddDate}');" class="LabelItem">����ʱ��</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetInput}');" class="LabelItem"><%=SQL(3,K)%>������(��Ա����)</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetCompanyName}');" class="LabelItem">��˾����</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetContactMan}');" class="LabelItem">��ϵ��</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetContactTel}');" class="LabelItem">��ϵ�绰</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetFax}');" class="LabelItem">�������</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetAddress}');" class="LabelItem">��ϸ��ַ</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetEmail}');" class="LabelItem">��������</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetPostCode}');" class="LabelItem">��������</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetProvince}');" class="LabelItem">��������ʡ��</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetCity}');" class="LabelItem">�������ڳ���</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetHomePage}');" class="LabelItem">��˾��ַ</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetPrevInfo}');" class="LabelItem">��ʾ��һ��<%=SQL(3,K)%></a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetNextInfo}');" class="LabelItem">��ʾ��һ��<%=SQL(3,K)%></a></li>
                  <li><a href="#" onClick="InsertFunctionLabel('GetShowComment.html',450,140);">��ʾ������Ϣ</a></li>
                  <li><a href="#" onClick="InsertLabel('{$GetWriteComment}');">��������</a></li>
                  <li><a href="#" onClick="InsertFunctionLabel('GetRandomContentsList.html',450,170);">��ʾ<%=SQL(3,K)%>����б�</a></li>
                  <li><a href="#" onClick="InsertFunctionLabel('GetRelatedContentsList.html',450,170);">��ʾ<%=SQL(3,K)%>����б�</a></li>

				<%End Select%>
					<div>============================</div>
					<div align='center'>�Զ����ֶα�ǩ</div>
					<div>============================</div>
                          <%
						  DIYFieldArr=KSCls.Get_KS_D_F_Arr(sql(0,k))
						  If IsArray(DIYFieldArr) Then
							  For i=0 To UBound(DIYFieldArr,2)
							  %>
					 <li><a href="#" onClick="InsertLabel('{$<%=DIYFieldArr(0,i)%>}');"><%=DIYFieldArr(1,i)%>-{$<%=DIYFieldArr(0,i)%>}</a></li>
							  <%
						      Next
                          End If
                           %>
                
                
                
              </div>		
			   <%Next%>				  
					 </td>
				    </tr>
				  </table>
				</td>
			  </tr>
			  </table>
			  </td>
		   </tr>
	  </table>
              
    
	
	 

	
		 <div onClick="ShowLabelTree('AdwLabel');" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                <img src="Images/folderclosed.gif" width="24" height="22" align="absmiddle"><a href="#">���λͨ�ñ�ǩ</a></div>
             <div id="AdwLabel" style="display:none">
				<%  
				
				Dim RSObj:Set RSObj=server.createobject("adodb.recordset")
					SqlStr="select Place,PlaceName From KS_ADPlace"
					RSObj.open SqlStr,Conn,1,1
					do while not RSObj.eof 
                %>
                    <li><a href="#" onClick="InsertLabel('{=GetAdvertise(<%=RSObj(0)%>)}');" class="LabelItem"> <%=RSObj(1)%></a></li>
				<%RSOBj.MoveNext
				 Loop
				 RSObj.Close:SET RSObj=Nothing
				 %>
                 
		     </div>
   
	    <div onClick="ShowLabelTree('VoteLabel');" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                <img src="Images/folderclosed.gif" width="24" height="22" align="absmiddle"><a href="#">��վ����ͨ�ñ�ǩ</a></div>
             <div id="VoteLabel" style="display:none">
				<%  Set RSObj=server.createobject("adodb.recordset")
					SqlStr="select ID,Title From KS_Vote"
					RSObj.open SqlStr,Conn,1,1
					do while not RSObj.eof 
                %>
                    <li><a href="#" onClick="InsertLabel('{=GetVote(<%=RSObj(0)%>)}');" class="LabelItem"><%=RSObj(1)%></a></li>
				<%RSOBj.MoveNext
				 Loop
				 RSObj.Close:SET RSObj=Nothing
				 %>
             </div>
			 
			<div onClick="ShowLabelTree('FreeLabel');" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                <img src="Images/folderclosed.gif" width="24" height="22" align="absmiddle"><a href="#">��վ����ͨ�ñ�ǩ</a></div>
             <div id="FreeLabel" style="display:none">
				<%  Set RSObj=server.createobject("adodb.recordset")
					SqlStr="select ID,Title From KS_Announce"
					RSObj.open SqlStr,Conn,1,1
					do while not RSObj.eof 
                %>
                    <li><a href="#" onClick="InsertLabel('{=GetAnnounce(<%=RSObj(0)%>)}');" class="LabelItem"><%=RSObj(1)%></a></li>
				<%RSOBj.MoveNext
				 Loop
				 RSObj.Close:SET RSObj=Nothing
				 %>
             </div> 
             

	  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr ParentID=""> 
          <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr onClick="ShowLabelTree('DIYFunctionLabel')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                <td width="24"><img src="Images/folderclosed.gif" width="24" height="22"></td>
                <td width="1227"><a href="#">�û��Զ��庯����ǩ</a></td>
              </tr>
              <tr> 
                <td colspan="2">
				 <table width="100%" border="0" cellspacing="0" cellpadding="0" id="DIYFunctionLabel" style="display:none">
                    <tr> 
                      <td width="8%" align="right">&nbsp;</td>
                      <td height="20"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                          <%
                          SqlStr = "Select * From KS_LabelFolder where FolderType=5 And ParentID='0'"
                         Set FolderRS = Conn.Execute(SqlStr)
                           if Not FolderRS.Eof then
	                    do while Not FolderRS.Eof
                           %>
                          <tr ParentID="<% = FolderRS("ParentID") %>" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                            <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                <tr> 
                                  <td width="3%"><img src="Images/folderclosed.gif" width="24" height="22"></td>
                                  <td width="97%"><span ShowFlag="False" TypeID="<% = FolderRS("ID") %>" onClick="SelectFolder(this)"><A href="#">
                                    <% = FolderRS("FolderName") %>
                                    </A></span></td>
                                </tr>
                              </table></td>
                          </tr>
                          <%
	 		        Response.Write(GetChildFolderList(0,5,FolderRS("ID"),""," style=""display:none;"" "))
                    Response.Write(GetLabelList(5,trim(FolderRS("ID")),"&nbsp;&nbsp;&nbsp;&nbsp;"," style=""display:none;"" "))
		        FolderRS.MoveNext
	            loop
              end if
               Response.Write(GetLabelList(5,"0","",""))
              %>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
            </table> 
          </td>
        </tr>
      </table>


	 
</td>
  </tr>

</table>
</body>
</html>
<%
Set Conn = Nothing
Set KS=Nothing
Set KSCls=Nothing
Function GetLabelList(LabelType,TypeID,CompatStr,ShowStr)
	Dim ListSql,LabelRS
	ListSql = "Select * from KS_Label where LabelType=" & LabelType &" And FolderID='" & Trim(TypeID) & "' ORDER BY LabelFlag Desc"
	Set LabelRS = Conn.Execute(ListSql)
	IF LabelRS.EOF AND LabelRS.BOF THEN
       GetLabelList=""	 
	   LabelRS.close:Set LabelRS=Nothing
	  EXIT Function
	END IF
	do while Not LabelRS.Eof
	  	GetLabelList = GetLabelList & "<tr onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#EEFFFF'"" ParentID=""" & LabelRS("FolderID") & """ " & ShowStr & ">" & vbcrlf
		GetLabelList = GetLabelList & "<td height=22>" & vbcrlf
		GetLabelList = GetLabelList & "<table border=""0"" cellspacing=""0"" cellpadding=""0""><tr><td>" & CompatStr &  "<img src=""Images/Label" & trim(LabelRS("LabelFlag")) & ".gif""></td>"
		If LabelType=5 Then
		 GetLabelList = GetLabelList & "<td><A href=""#"" onclick=""InsertFunctionLabel('"&DomainStr&"KS_Editor/InsertFunctionfield.asp?ID=" & Trim(LabelRS("ID")) & "',300,350)"">" & LabelRS("LabelName") & "</A></td></tr></table>"
		Else
		GetLabelList = GetLabelList & "<td><A href=""#"" onclick=""InsertLabel('" & Trim(LabelRS("LabelName")) & "')"">" & LabelRS("LabelName") & "</A></td></tr></table>"
		End If
		GetLabelList = GetLabelList & "</td>" & vbcrlf
		GetLabelList = GetLabelList & "</tr>" & vbcrlf
		LabelRS.MoveNext
	Loop
	Set LabelRS = Nothing
End Function
Function GetJSList(JSType,TypeID,CompatStr,ShowStr)
	Dim ListSql,JSRS
	ListSql = "Select * from KS_JSFile where JSType=" & JSType &" And FolderID='" & Trim(TypeID) & "' ORDER BY AddDate Desc"
	Set JSRS = Conn.Execute(ListSql)
	IF JSRS.EOF AND JSRS.BOF THEN
       GetJSList=""	 
	   JSRS.close
	   Set JSRS=Nothing
	  EXIT Function
	END IF
	do while Not JSRS.Eof
	  	GetJSList = GetJSList & "<tr onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#EEFFFF'"" ParentID=""" & JSRS("FolderID") & """ " & ShowStr & ">" & vbcrlf
		GetJSList = GetJSList & "<td height=22>" & vbcrlf
		GetJSList = GetJSList & "<table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf & "<tr>"  & vbcrlf & "<td>" & CompatStr &  "<img src=""Images/JS" & trim(JSType) & ".gif""></td>"
		GetJSList = GetJSList & "<td><A href=""#"" onclick=""InsertLabel('" & Trim(JSRS("JSName")) & "')"">" & JSRS("JSName") & "</A></td>" & vbcrlf & "</tr>" & vbcrlf & "</table>"
		GetJSList = GetJSList & "</td>" & vbcrlf
		GetJSList = GetJSList & "</tr>" & vbcrlf
		JSRS.MoveNext
	Loop
	Set JSRS = Nothing
End Function
Function GetChildFolderList(GetType,LabelType,TypeID,CompatStr,ShowStr)
	Dim ChildFolderRS,ChildTypeListStr,TempStr
	Set ChildFolderRS = Conn.Execute("Select * FROM KS_LabelFolder where ParentID='" & TypeID & "'")
	TempStr = CompatStr & "&nbsp;&nbsp;&nbsp;&nbsp;"
	do while Not ChildFolderRS.Eof
	  	GetChildFolderList = GetChildFolderList & "<tr onmouseout=""this.style.backgroundColor=''"" onmouseover=""this.style.backgroundColor='#EEEEEE'"" TypeFlag=""Class"" ParentID=""" & ChildFolderRS("ParentID") & """ " & ShowStr & ">" & vbcrlf
		GetChildFolderList = GetChildFolderList & "<td>" & vbcrlf
		GetChildFolderList = GetChildFolderList & "<table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf & "<tr>"  & vbcrlf & "<td>" & TempStr & "<img src=""Images/folderclosed.gif""></td>"
		GetChildFolderList = GetChildFolderList & "<td><span TypeID=""" & ChildFolderRS("ID") & """ ShowFlag=""False"" onClick=""SelectFolder(this)""><a href=""#"">" & ChildFolderRS("FolderName") & "</a></span></td>" & vbcrlf & "</tr>" & vbcrlf & "</table>"
		GetChildFolderList = GetChildFolderList & "</td>" & vbcrlf
		GetChildFolderList = GetChildFolderList & "</tr>" & vbcrlf
		IF GetType=0 Then
		  GetChildFolderList = GetChildFolderList & vbcrlf & GetLabelList(LabelType,trim(ChildFolderRS("ID")),"&nbsp;&nbsp;&nbsp;&nbsp;" & TempStr,ShowStr) 
		Else
		  GetChildFolderList = GetChildFolderList & vbcrlf & GetJSList(LabelType,trim(ChildFolderRS("ID")),"&nbsp;&nbsp;&nbsp;&nbsp;" & TempStr,ShowStr) 
		End IF
		GetChildFolderList = GetChildFolderList & GetChildFolderList(GetType,LabelType,ChildFolderRS("ID"),TempStr,ShowStr)
		ChildFolderRS.MoveNext
	loop
	ChildFolderRS.Close
	Set ChildFolderRS = Nothing
End Function
%>
<script language="JavaScript">
function ShowLabelTree(Obj)
{
 switch (Obj)
  {
	case 'CommonJSLabel':
     if (document.all.CommonJSLabel.style.display!='')
       {document.all.CommonJSLabel.style.display='';}
     else
      {document.all.CommonJSLabel.style.display='none';} 
	  break;
  <%For K=0 To Ubound(SQL,2)%>
   case '<%=SQL(3,K)%>Content':
     if (document.all.<%=SQL(3,K)%>Content.style.display!='')
       {document.all.<%=SQL(3,K)%>Content.style.display='';}
     else
      {document.all.<%=SQL(3,K)%>Content.style.display='none';} 
	  break;
   <%Next%>

   case 'AnnounceContent':
     if (document.all.AnnounceContent.style.display!='')
       {document.all.AnnounceContent.style.display='';}
     else
      {document.all.AnnounceContent.style.display='none';} 
	  break;


  case 'FreeLabel' :
      if (document.all.FreeLabel.style.display!='')
      {document.all.FreeLabel.style.display='';}
     else
      {document.all.FreeLabel.style.display='none';} 
	  break;
  case 'DIYFunctionLabel' :
      if (document.all.DIYFunctionLabel.style.display!='')
      {document.all.DIYFunctionLabel.style.display='';}
     else
      {document.all.DIYFunctionLabel.style.display='none';} 
	  break;  


 case 'AdwLabel':
      if (document.all.AdwLabel.style.display!='')
      {document.all.AdwLabel.style.display='';}
     else
      {document.all.AdwLabel.style.display='none';} 
	  break; 
 case 'VoteLabel':
      if (document.all.VoteLabel.style.display!='')
      {document.all.VoteLabel.style.display='';}
     else
      {document.all.VoteLabel.style.display='none';} 
	  break; 


 case 'wapshouye':
      if (document.all.wapshouye.style.display!='')
      {document.all.wapshouye.style.display='';}
     else
      {document.all.wapshouye.style.display='none';} 
	  break;
 case 'wapkongjian':
      if (document.all.wapkongjian.style.display!='')
      {document.all.wapkongjian.style.display='';}
     else
      {document.all.wapkongjian.style.display='none';} 
	  break;
 case 'waplanmu':
      if (document.all.waplanmu.style.display!='')
      {document.all.waplanmu.style.display='';}
     else
      {document.all.waplanmu.style.display='none';} 
	  break;
 case 'wapchangyongfenlei':
     if (document.all.wapchangyongfenlei.style.display!='')
       {document.all.wapchangyongfenlei.style.display='';}
     else
      {document.all.wapchangyongfenlei.style.display='none';} 
	  break;
 case 'wapneirong':
     if (document.all.wapneirong.style.display!='')
       {document.all.wapneirong.style.display='';}
     else
      {document.all.wapneirong.style.display='none';} 
	  break;
<%For K=0 To Ubound(SQL,2)%>
 case 'wapneirong<%=SQL(6,K)%>':
     if (document.all.wapneirong<%=SQL(6,K)%>.style.display!='')
       {document.all.wapneirong<%=SQL(6,K)%>.style.display='';}
     else
      {document.all.wapneirong<%=SQL(6,K)%>.style.display='none';} 
	  break;
   <%Next%>
 case 'wapchangyongLabel':
      if (document.all.wapchangyongLabel.style.display!='')
      {document.all.wapchangyongLabel.style.display='';}
     else
      {document.all.wapchangyongLabel.style.display='none';} 
	  break; 
 case 'wapchangyong':
      if (document.all.wapchangyong.style.display!='')
      {document.all.wapchangyong.style.display='';}
     else
      {document.all.wapchangyong.style.display='none';} 
	  break; 
 }
}
function InsertLabel(LabelContent)
{
	window.returnValue=LabelContent;
	window.close();
}
function InsertFunctionLabel(Url,Width,Height)
{
window.returnValue = OpenWindow(Url,Width,Height,window);
window.close();
}
window.onunload=SetReturnValue;
function SetReturnValue()
{
	if (typeof(window.returnValue)!='string') window.returnValue='';
}
function SelectFolder(Obj)
{
	var CurrObj;
	if (Obj.ShowFlag=='True')
	{
		ShowOrDisplay(Obj,'none',true);
		Obj.ShowFlag='False';
	}
	else
	{
		ShowOrDisplay(Obj,'',false);
		Obj.ShowFlag='True';
	}
}
function ShowOrDisplay(Obj,Flag,Tag)
{
	for (var i=0;i<document.all.length;i++)
	{
		CurrObj=document.all(i);
		if (CurrObj.ParentID==Obj.TypeID)
		{
			CurrObj.style.display=Flag;
			if (Tag) 
			if (CurrObj.TypeFlag=='Class') ShowOrDisplay(CurrObj.children(0).children(0).children(0).children(0).children(1).children(0),Flag,Tag);
		}
	}
}
</script> 