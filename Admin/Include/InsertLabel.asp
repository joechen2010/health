<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../../KS_Cls/Kesion.AdministratorCls.asp"-->
<!--#include file="Session.asp"-->
<%
Dim SChannelID:SchannelID=request("schannelid")   'SchannelID=9999��������ɱ�ǩ/JS����
Dim TemplateType:TemplateType=request("templateType")
Dim KS,KSCls,SQL,K,i,DIYFieldArr,F_B,F_V
On Error Resume Next
Set KS=New PublicCls
Set KSCls=New ManageCls
Dim DomainStr:DomainStr=KS.GetDomain
Dim RS:Set RS=Conn.Execute("Select ChannelID,BasicType,ChannelName,ItemName,ItemUnit,FieldBit,ModelEname From KS_Channel Where ChannelStatus=1 and channelid<>6  And ChannelID<>9 And ChannelID<>10 Order By ChannelID")
SQL=RS.GetRows(-1)
RS.Close:Set RS=Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<script language="JavaScript" src="../../ks_inc/Common.js"></script>
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
list-style-image:url(../Images/label0.gif);
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
          <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr  onmouseout="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'">
                <td><img src="../Images/home.gif" width="18" height="18"></td>
                <td height="20">��ǩ����</td>
              </tr>
              <tr onClick="ShowLabelTree('TY')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                <td width="24"><img src="../Images/Folder/folderclosed.gif" width="24" height="22"></td>
                <td width="672"><a href="#">��վͨ�ñ�ǩ</a></td>
              </tr>
              <tr> 
                <td colspan="2"> 
				   <div id="TY" style="display:none">
                    <li><a href="#" onClick="InsertLabel('{$GetSiteTitle}');">��ʾ��վ����</a></li>
                    <li><a href="#" onClick="InsertLabel('{$GetSiteName}');">��ʾ��վ����</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetSiteLogo}');">��ʾ��վLogo(��������)</a></li>
					<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>KS_Editor/KS_Logo.html',250,130);">��ʾ��վLogo(������)</a></li>
					<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>KS_Editor/KS_Tags.html',250,130);">��ʾ����Tags/����Tags</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetSiteCountAll}');">��ʾ��վ��Ϣͳ��</a></li>
                    <li><a href="#" onClick="InsertLabel('{$GetSiteOnline}');">��ʾ��������(�����ߣ�1�� �û���1�� �οͣ�0��)</a></li>
					<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>KS_Editor/KS_TopUser.html',250,130);">��ʾ��Ծ����</a></li>
					<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>KS_Editor/KS_UserDynamic.html',250,130);">��ʾ�û���̬</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetSpecial}');">��ʾר�����</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetFriendLink}');">��ʾ�����������</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetSiteUrl}');">��ʾ��վURL</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetInstallDir}');">��ʾ��վ��װ·��</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetManageLogin}');">��ʾ�������</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetCopyRight}');">��ʾ��Ȩ��Ϣ</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetMetaKeyWord}');">��ʾ�����������Ĺؼ���</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetMetaDescript}');">��ʾ����������������</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetWebmaster}');">��ʾվ��</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetWebmasterEmail}');">��ʾվ��EMail</a></li>
				 </div>
				 </td>
              </tr>
            </table></td>
        </tr>
      </table>
	  
	   <div onClick="ShowLabelTree('CommonJSLabel')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'">
               <img src="../Images/Folder/folderclosed.gif" width="24" height="22" align="absmiddle"><a href="#">���ýű���Ч��ǩ</a></div>
              
				 <div id="CommonJSLabel" style="display:none">
				     <li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>KS_Editor/KS_Ad.html',550,180);" class="LabelItem">�������</a></li>
					 <li><a href="#" onClick="InsertLabel('{$JS_Time1}');" class="LabelItem">ʱ����Ч(��ʽ:2006��4��8��)</a></li>
					 <li><a href="#" onClick="InsertLabel('{$JS_Time2}');" class="LabelItem">ʱ����Ч(��ʽ:2006��4��8�� ������)</a></li>
					 <li><a href="#" onClick="InsertLabel('{$JS_Time3}');" class="LabelItem">ʱ����Ч(��ʽ:2007��6��1�� �����塾ũ�� 4��...)</a></li>
					 <li><a href="#" onClick="InsertLabel('{$JS_Time4}');" class="LabelItem">ʱ����Ч(��ʽ:2006��4��8�� 11:50:46 ������)</a></li>
					 <li><a href="#" onClick="InsertLabel('{$JS_Language}');" class="LabelItem">��ת��</a></li>
					 <li><a href="#" onClick="InsertLabel('{$JS_HomePage}');" class="LabelItem">��Ϊ��ҳ</a></li>
					 <li><a href="#" onClick="InsertLabel('{$JS_Collection}');" class="LabelItem">�����ղ�</a></li>
					 <li><a href="#" onClick="InsertLabel('{$JS_ContactWebMaster}');" class="LabelItem">��ϵվ��</a></li>
					 <li><a href="#" onClick="InsertLabel('{$JS_GoBack}');" class="LabelItem">������һҳ</a></li>
					 <li><a href="#" onClick="InsertLabel('{$JS_WindowClose}');" class="LabelItem">�رմ���</a></li>
					 <li><a href="#" onClick="InsertLabel('{$JS_NoSave}');" class="LabelItem">ҳ�治������"���Ϊ"</a></li>
					 <li><a href="#" onClick="InsertLabel('{$JS_NoIframe}');" class="LabelItem">ҳ�治�����˷��ڿ����</a></li>
					 <li><a href="#" onClick="InsertLabel('{$JS_NoCopy}');" class="LabelItem">��ֹ��ҳ��Ϣ������</a></li>
					 <li><a href="#" onClick="InsertLabel('{$JS_DCRoll}');" class="LabelItem">˫��������Ч</a></li>
					 <li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>KS_Editor/KS_Status1.html',550,150);" class="LabelItem">״̬������Ч��</a></li>
					 <li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>KS_Editor/KS_Status2.html',550,150);" class="LabelItem">������״̬���ϴ�������ѭ����ʾ</a></li>
					 <li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>KS_Editor/KS_Status3.html',550,150);" class="LabelItem">������״̬���ϴ���֮���ƶ���ʧ</a></li>
					</div>
               
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr ParentID=""> 
          <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr onClick="ShowLabelTree('SysFLabel')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                <td width="24"><img src="../Images/Folder/folderclosed.gif" width="24" height="22"></td>
                <td width="672"><a href="#"><font color="blue">ϵͳ������ǩ(KesionCMS���ű�ǩ)</font></a></td>
              </tr>
              <tr> 
                <td colspan="2"> <table width="100%" border="0" cellspacing="0" cellpadding="0" id="SysFLabel" style="display:none">
                    <tr> 
                      <td width="8%" align="right">&nbsp;</td>
                      <td height="20"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                          <%dim FolderRS,SqlStr
                          SqlStr = "Select * From KS_LabelFolder where FolderType=0 And ParentID='0'"
                         Set FolderRS = Conn.Execute(SqlStr)
                           if Not FolderRS.Eof then
	                    do while Not FolderRS.Eof
                           %>
                          <tr ParentID="<% = FolderRS("ParentID") %>" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                            <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                <tr> 
                                  <td width="3%"><img src="../Images/Folder/folderclosed.gif" width="24" height="22"></td>
                                  <td width="97%"><span ShowFlag="False" TypeID="<% = FolderRS("ID") %>" onClick="SelectFolder(this)"><A href="#">
                                    <% = FolderRS("FolderName") %>
                                    </A></span></td>
                                </tr>
                              </table></td>
                          </tr>
                          <%
	 		        Response.Write(GetChildFolderList(0,0,FolderRS("ID"),""," style=""display:none;"" "))
                    Response.Write(GetLabelList(0,trim(FolderRS("ID")),"&nbsp;&nbsp;&nbsp;&nbsp;"," style=""display:none;"" "))
		        FolderRS.MoveNext
	            loop
              end if
               Response.Write(GetLabelList(0,"0","",""))
              %>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
      </table>
	  
               
         <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr ParentID=""> 
              <td> 
			   <table width="100%" border="0" cellpadding="0" cellspacing="0">
               <tr onClick="ShowLabelTree('ContentLabel')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                <td width="24"><img src="../Images/Folder/folderclosed.gif" width="24" height="22"></td>
                <td width="672"><a href="#"><font color="red">����ҳ��ǩ</font></a></td>
               </tr>
               <tr> 
                <td colspan="2">
				    <table width="85%" align='center' border="0" cellspacing="0" cellpadding="0" id="ContentLabel" style="display:none">
                    <tr> 
					 <td>
					 <%
					  For K=0 To Ubound(SQL,2)
					   F_B=Split(Split(SQL(5,K),"@@@")(0),"|")
					   F_V=Split(Split(SQL(5,K),"@@@")(1),"|")
					  %>
					 	<div onClick="ShowLabelTree('Content<%=SQL(6,K)%>')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'">
					<img src="../Images/Folder/folderclosed.gif" align="absmiddle" width="24" height="22"><a href="#"><%=SQL(3,K)%>����ҳ��ǩ(<%=sql(2,k)%>)</a>
				     </div>	
					 <%Select Case SQL(1,K)%>
					  <%case 1%>
					  <div  id="Content<%=SQL(6,K)%>" style="display:none">
						<li><a href="#" onClick="InsertLabel('{$GetArticleUrl}');" class="LabelItem">��ǰ<%=SQL(3,K)%>URL</a></li>
						<li><a href="#" onClick="InsertLabel('{$ChannelID}');" class="LabelItem">��ǰģ��ID</a></li>
						<li><a href="#" onClick="InsertLabel('{$InfoID}');" class="LabelItem">��ǰ<%=SQL(3,K)%>СID</a></li>
						<li><a href="#" onClick="InsertLabel('{$ItemName}');" class="LabelItem">��ǰ��Ŀ����</a></li>
						<li><a href="#" onClick="InsertLabel('{$ItemUnit}');" class="LabelItem">��ǰ��Ŀ��λ</a></li>
						
						<li><a href="#" onClick="InsertLabel('{$GetArticleShortTitle}');" class="LabelItem"><%=SQL(3,K)%>��̱���</a></li>
						<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>KS_Editor/KS_ArtPhoto.html',250,130);" class="LabelItem">����ҳͼƬ</a></li>
						<%if F_B(2)=1 Then%>
						<li><a href="#" onClick="InsertLabel('{$GetArticleTitle}');" class="LabelItem"><%=F_V(2)%></a></li>
						<%end if%>
						<%if F_B(5)=1 Then%>
						<li><a href="#" onClick="InsertLabel('{$GetArticleKeyWord}');" class="LabelItem"><%=F_V(5)%></a></li>
						<%end if%>
						<li><a href="#" onClick="InsertLabel('{$GetKeyTags}');" class="LabelItem">ȡ��<%=SQL(3,K)%>Tags</a></li>
						<%if F_B(8)=1 Then%>
						<li><a href="#" onClick="InsertLabel('{$GetArticleIntro}');" class="LabelItem"><%=F_V(8)%></a></li>
						<%end if%>
						<%if F_B(9)=1 Then%>
						<li><a href="#" onClick="InsertLabel('{$GetArticleContent}');" class="LabelItem"><%=F_V(9)%></a></li>
						<%end if%>
						<%if F_B(6)=1 Then%>
						<li><a href="#" onClick="InsertLabel('{$GetArticleAuthor}');"><%=F_V(6)%></a></li>
						<%end if%>
						<%if F_B(7)=1 Then%>
						<li><a href="#" onClick="InsertLabel('{$GetArticleOrigin}');"><%=F_V(7)%></a></li>
						<%end if%>
						<%if F_B(12)=1 Then%>
						<li><a href="#" onClick="InsertLabel('{$GetArticleDate}');"><%=F_V(12)%>(��ʽ:2009��5��1��)</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetDate}');"><%=F_V(12)%>(ֱ�����)</a></li>
						<%end if%>
						<%if F_B(14)=1 Then%>
						 <li><a href="#" onClick="InsertLabel('{$GetHits}');" class="LabelItem"><%=SQL(3,K)%>����(�������)</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetHitsByDay}');" class="LabelItem"><%=SQL(3,K)%>���������</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetHitsByWeek}');" class="LabelItem"><%=SQL(3,K)%>���������</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetHitsByMonth}');" class="LabelItem"><%=SQL(3,K)%>���������</a></li>
						<%end if%>
						<li><a href="#" onClick="InsertLabel('{$GetArticleInput}');"><%=SQL(3,K)%>¼��(������)</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetUserName}');"><%=SQL(3,K)%>¼��(��������)</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetRank}');"><%=SQL(3,K)%>�Ƽ��ȼ�</a></li>
						<%if F_B(3)=1 Then%>
						<li><a href="#" onClick="InsertLabel('{$GetArticleProperty}');">��ʾ<%=SQL(3,K)%>������(���š��Ƽ���������...)</a></li>
						<%end if%>
						<li><a href="#" onClick="InsertLabel('{$GetArticleSize}');">��ʾ<%=SQL(3,K)%>������:�� �� С��</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetArticleAction}');">��ʾ���������ۡ������ߺ��ѡ�����ӡ...</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetPrevArticle}');">��ʾ��һ<%=SQL(4,K)%><%=SQL(3,K)%></a></li>
						<li><a href="#" onClick="InsertLabel('{$GetNextArticle}');">��ʾ��һ<%=SQL(4,K)%><%=SQL(3,K)%></a></li>
						<li><a href="#" onClick="InsertLabel('{$GetPrevUrl}');">��ʾ��һ<%=SQL(4,K)%><%=SQL(3,K)%>URL</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetNextUrl}');">��ʾ��һ<%=SQL(4,K)%><%=SQL(3,K)%>URL</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetShowComment}');">��ʾ����</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetWriteComment}');">��������</a></li>
						
                     <%Case 2%>					  
					  <div id="Content<%=SQL(6,K)%>" style="display:none">
						 <li><a href="#" onClick="InsertLabel('{$GetPictureUrl}');" class="LabelItem">��ǰ<%=SQL(3,K)%>URL</a></li>
						<li><a href="#" onClick="InsertLabel('{$ChannelID}');" class="LabelItem">��ǰģ��ID</a></li>
						<li><a href="#" onClick="InsertLabel('{$InfoID}');" class="LabelItem">��ǰ<%=SQL(3,K)%>СID</a></li>
						<li><a href="#" onClick="InsertLabel('{$ItemName}');" class="LabelItem">��ǰ��Ŀ����</a></li>
						<li><a href="#" onClick="InsertLabel('{$ItemUnit}');" class="LabelItem">��ǰ��Ŀ��λ</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetPictureName}');" class="LabelItem"><%=F_V(0)%></a></li>
						 <%if F_B(6)=1 Then%>
						 <li><a href="#" onClick="InsertLabel('{$GetPictureKeyWord}');" class="LabelItem"><%=F_V(6)%></a></li>
						 <%end if%>
						<li><a href="#" onClick="InsertLabel('{$GetKeyTags}');" class="LabelItem">ȡ��<%=SQL(3,K)%>Tags</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetPictureSrc}');" class="LabelItem">ȡ��<%=SQL(3,K)%>����ͼSrc</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetPictureByPlayer}');" class="LabelItem">�鿴<%=SQL(3,K)%>����(��������ʽ)</a>
							 </li>
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
						 <li><a href="#" onClick="InsertLabel('{$GetPictureDate}');"><%=F_V(10)%>(��ʽ:2009��5��1��)</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetDate}');"><%=F_V(10)%>(ֱ�����)</a></li>
						 <%end if%>
						 <li><a href="#" onClick="InsertLabel('{$GetHits}');" class="LabelItem"><%=SQL(3,K)%>����(�������)</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetHitsByDay}');" class="LabelItem"><%=SQL(3,K)%>���������</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetHitsByWeek}');" class="LabelItem"><%=SQL(3,K)%>���������</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetHitsByMonth}');" class="LabelItem"><%=SQL(3,K)%>���������</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetPictureInput}');"><%=SQL(3,K)%>¼��(������)</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetUserName}');"><%=SQL(3,K)%>¼��(��������)</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetRank}');"><%=SQL(3,K)%>�Ƽ��ȼ�</a></li>
						 <%if F_B(5)=1 Then%>
						 <li><a href="#" onClick="InsertLabel('{$GetPictureProperty}');" class="LabelItem">��ʾ<%=SQL(3,K)%>����(���š��������Ƽ�...</a></li>
						 <%end if%>
						 <li><a href="#" onClick="InsertLabel('{$GetPictureAction}');">&nbsp;��ʾ���������ۡ�����Ҫ...��</a><</li>
						 <li><a href="#" onClick="InsertLabel('{$GetPictureVote}');" class="LabelItem">��ʾͶ��һƱ</a> </li>
						 <li><a href="#" onClick="InsertLabel('{$GetPictureVoteScore}');" class="LabelItem">��ʾ<%=SQL(3,K)%>��Ʊ��</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetPrevPicture}');" class="LabelItem">��ʾ��һ��<%=SQL(3,K)%></a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetNextPicture}');" class="LabelItem">��ʾ��һ��<%=SQL(3,K)%></a></li>
						<li><a href="#" onClick="InsertLabel('{$GetPrevUrl}');">��ʾ��һ��<%=SQL(3,K)%>URL</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetNextUrl}');">��ʾ��һ��<%=SQL(3,K)%>URL</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetShowComment}');">��ʾ����</a></li>
						 <li><a href="#" onClick="InsertLabel('{$GetWriteComment}');">��������</a></li>
				<%Case 3%>
				 <div id="Content<%=SQL(6,K)%>" style="display:none">
					 <li><a href="#" onClick="InsertLabel('{$GetDownUrl}');" class="LabelItem">��ǰ<%=SQL(3,K)%>URL</a></li>
						<li><a href="#" onClick="InsertLabel('{$ChannelID}');" class="LabelItem">��ǰģ��ID</a></li>
						<li><a href="#" onClick="InsertLabel('{$InfoID}');" class="LabelItem">��ǰ<%=SQL(3,K)%>СID</a></li>
						<li><a href="#" onClick="InsertLabel('{$ItemName}');" class="LabelItem">��ǰ��Ŀ����</a></li>
						<li><a href="#" onClick="InsertLabel('{$ItemUnit}');" class="LabelItem">��ǰ��Ŀ��λ</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetDownTitle}');" class="LabelItem"><%=SQL(3,K)%>����+�汾��</a></li>
					 <%if F_B(10)=1 Then%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownKeyWord}');" class="LabelItem"><%=F_V(10)%></a></li>
					 <%end if%>
						<li><a href="#" onClick="InsertLabel('{$GetKeyTags}');" class="LabelItem">ȡ��<%=SQL(3,K)%>Tags</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetDownAddress}');" class="LabelItem">���ص�ַ</a></li>
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
					 <li><a href="#" onClick="InsertLabel('{$GetDownIntro}');" class="LabelItem"><%=F_V(14)%></a></li>
					 <%end if%>
					 <%if F_B(11)=1 Then%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownAuthor}');" class="LabelItem"><%=F_V(11)%></a></li>
					 <%end if%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownInput}');" class="LabelItem"><%=SQL(3,K)%>¼��(������)</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetUserName}');" class="LabelItem"><%=SQL(3,K)%>¼��(��������)</a></li>
					 <%if F_B(12)=1 Then%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownOrigin}');" class="LabelItem">�� Դ</a></li>
					 <%end if%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownDate}');" class="LabelItem">���(����)����</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetHits}');" class="LabelItem">�����ص����</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetHitsByDay}');" class="LabelItem">���յ����</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetHitsByWeek}');" class="LabelItem">���ܵ����</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetHitsByMonth}');" class="LabelItem">���µ����</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetDownLink}');" class="LabelItem">������ӣ���ʾ��ַ+ע���ַ��</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetDownPoint}');" class="LabelItem">����������</a></li>
					 <%if F_B(15)=1 Then%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownYSDZ}');" class="LabelItem"><%=F_V(15)%></a></li>
					 <%end if%>
					 <%if F_B(16)=1 Then%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownZCDZ}');" class="LabelItem"><%=F_V(16)%></a></li>
					 <%end if%>
					 <%if F_B(17)=1 Then%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownDecPass}');" class="LabelItem"><%=F_V(17)%></a></li>
					 <%end if%>
					 <li><a href="#" onClick="InsertLabel('{$GetDownProperty}');" class="LabelItem">��ʾ<%=SQL(3,K)%>����(���š��Ƽ��ȣ�</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetDownAction}');">��ʾ���������ۡ�����Ҫ...</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetRank}');" class="LabelItem">��ʾ�Ƽ��Ǽ�</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetPrevDown}');" class="LabelItem">��ʾ��һ��<%=SQL(3,K)%></a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetNextDown}');" class="LabelItem">��ʾ��һ��<%=SQL(3,K)%></a></li>
						<li><a href="#" onClick="InsertLabel('{$GetPrevUrl}');">��ʾ��һ��<%=SQL(3,K)%>URL</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetNextUrl}');">��ʾ��һ��<%=SQL(3,K)%>URL</a></li>
					 <Li><a href="#" onClick="InsertLabel('{$GetShowComment}');">��ʾ����</a></li>
					 <li><a href="#" onClick="InsertLabel('{$GetWriteComment}');">��������</a></li>
				<%Case 4%>
				 <div id="Content<%=SQL(6,K)%>" style="display:none">
						<li><a href="#" onClick="InsertLabel('{$GetFlashUrl}');" class="LabelItem">��ǰ<%=SQL(3,K)%> URL</a></li>
						<li><a href="#" onClick="InsertLabel('{$ChannelID}');" class="LabelItem">��ǰģ��ID</a></li>
						<li><a href="#" onClick="InsertLabel('{$InfoID}');" class="LabelItem">��ǰ<%=SQL(3,K)%>СID</a></li>
						<li><a href="#" onClick="InsertLabel('{$ItemName}');" class="LabelItem">��ǰ��Ŀ����</a></li>
						<li><a href="#" onClick="InsertLabel('{$ItemUnit}');" class="LabelItem">��ǰ��Ŀ��λ</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetFlashName}');" class="LabelItem"><%=SQL(3,K)%>����</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetFlashKeyWord}');" class="LabelItem">��ǰ<%=SQL(3,K)%>�ؼ���</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetKeyTags}');" class="LabelItem">ȡ��<%=SQL(3,K)%>Tags</a></li>
						<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>KS_Editor/KS_FlashPlayer.html',250,130);" class="LabelItem">�鿴<%=SQL(3,K)%>����(��������ʽ����)</a></li>
						<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>KS_Editor/KS_Flash.html',250,130);" class="LabelItem">�鿴<%=SQL(3,K)%>����(��ͨ��ʽ����)</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetFlashIntro}');" class="LabelItem"><%=SQL(3,K)%>���</a> </li>
						<li><a href="#" onClick="InsertLabel('{$GetFlashAuthor}');" class="LabelItem"><%=SQL(3,K)%>����</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetFlashOrigin}');" class="LabelItem">�� Դ</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetFlashDate}');" class="LabelItem">���(����)����</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetFlashSrc}');" class="LabelItem"><%=SQL(3,K)%>��ַ</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetFlashFullScreen}');" class="LabelItem">ȫ���ۿ�</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetHits}');" class="LabelItem"><%=SQL(3,K)%>����(�������)</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetHitsByDay}');" class="LabelItem"><%=SQL(3,K)%>���������</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetHitsByWeek}');" class="LabelItem"><%=SQL(3,K)%>���������</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetHitsByMonth}');" class="LabelItem"><%=SQL(3,K)%>���������</a></li>
		
						<li><a href="#" onClick="InsertLabel('{$GetFlashInput}');" class="LabelItem">����¼��</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetFlashProperty}');" class="LabelItem">��ʾ<%=SQL(3,K)%>����(���š��������Ƽ����ȼ��ȣ�</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetFlashAction}');">&nbsp;��ʾ���������ۡ�����Ҫ�ղء����رմ��ڡ�</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetFlashVote}');" class="LabelItem">��ʾͶ��һƱ</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetFlashVoteScore}');" class="LabelItem">��ʾ<%=SQL(3,K)%>��Ʊ��</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetPrevFlash}');" class="LabelItem">��ʾ��һ��<%=SQL(3,K)%></a></li>
						<li><a href="#" onClick="InsertLabel('{$GetNextFlash}');" class="LabelItem">��ʾ��һ��<%=SQL(3,K)%></a></li>
						<li><a href="#" onClick="InsertLabel('{$GetPrevUrl}');">��ʾ��һ��<%=SQL(3,K)%>URL</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetNextUrl}');">��ʾ��һ��<%=SQL(3,K)%>URL</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetShowComment}');">&nbsp;��ʾ����</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetWriteComment}');">&nbsp;��������</a></li>
			       <%Case 5%>
				<div id="Content<%=SQL(6,K)%>" style="display:none">
					<li><a href="#" onClick="InsertLabel('{$GetProductUrl}');" class="LabelItem">��ǰ<%=SQL(3,K)%> URL</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetProductID}');" class="LabelItem">��ǰ<%=SQL(3,K)%>���(ID)</a></li>
						<li><a href="#" onClick="InsertLabel('{$ChannelID}');" class="LabelItem">��ǰģ��ID</a></li>
						<li><a href="#" onClick="InsertLabel('{$InfoID}');" class="LabelItem">��ǰ<%=SQL(3,K)%>СID</a></li>
						<li><a href="#" onClick="InsertLabel('{$ItemName}');" class="LabelItem">��ǰ��Ŀ����</a></li>
						<li><a href="#" onClick="InsertLabel('{$ItemUnit}');" class="LabelItem">��ǰ��Ŀ��λ</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetProductName}');" class="LabelItem"><%=SQL(3,K)%>����</a></li>
					<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>KS_Editor/KS_ProductPhoto.html',250,130);" class="LabelItem">��ƷͼƬ</a> </li>
					<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>KS_Editor/KS_ProductGroupPhoto.html',250,130);" class="LabelItem">��ʾ��ƷͼƬ�� <font color=red>new</font></a> </li>
					
					<li><a href="#" onClick="InsertLabel('{$GetProductKeyWord}');" class="LabelItem">��ǰ<%=SQL(3,K)%>�ؼ���</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetKeyTags}');" class="LabelItem">ȡ��<%=SQL(3,K)%>Tags</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetProductIntro}');" class="LabelItem"><%=SQL(3,K)%>���</a> </li>
					<li><a href="#" onClick="InsertLabel('{$GetProducerName}');" class="LabelItem">�� �� ��</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetTrademarkName}');" class="LabelItem">Ʒ���̱�</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetProductModel}');" class="LabelItem"><%=SQL(3,K)%>�ͺ�</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetProductSpecificat}');" class="LabelItem"><%=SQL(3,K)%>���</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetProductDate}');" class="LabelItem">�ϼ�ʱ��</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetServiceTerm}');" class="LabelItem">��������</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetTotalNum}');" class="LabelItem">�������</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetProductUnit}');" class="LabelItem"><%=SQL(3,K)%>��λ</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetHits}');" class="LabelItem"><%=SQL(3,K)%>����(�������)</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetHitsByDay}');" class="LabelItem"><%=SQL(3,K)%>���������</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetHitsByWeek}');" class="LabelItem"><%=SQL(3,K)%>���������</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetHitsByMonth}');" class="LabelItem"><%=SQL(3,K)%>���������</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetProductType}');" class="LabelItem">��������</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetRank}');" class="LabelItem">�Ƽ��ȼ�</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetProductProperty}');" class="LabelItem">��ʾ<%=SQL(3,K)%>����(�������ؼۡ��Ƽ��ȣ�</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetProductInput}');">&nbsp;��ʾ��Ʒ¼��</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetPrice_Original}');">&nbsp;��ʾԭʼ���ۼ�</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetPrice}');" class="LabelItem">��ʾ��ǰ���ۼ�</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetPrice_Member}');" class="LabelItem">��ʾ��Ա��</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetPrice_Market}');" class="LabelItem">��ʾ�г���</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetGroupPrice}');" class="LabelItem">�Զ�ȡ�û���۸� <font color=red>new</font></a></li>
					<li><a href="#" onClick="InsertLabel('{$GetDiscount}');" class="LabelItem">��ʾ�ۿ���</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetScore}');" class="LabelItem">��ʾ�������</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetAddCar}');" class="LabelItem">���빺�ﳵ</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetAddFav}');" class="LabelItem">�����ղؼ�</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetPrevProduct}');" class="LabelItem">��ʾ��һ��<%=SQL(3,K)%></a></li>
					<li><a href="#" onClick="InsertLabel('{$GetNextProduct}');" class="LabelItem">��ʾ��һ��<%=SQL(3,K)%></a></li>
						<li><a href="#" onClick="InsertLabel('{$GetPrevUrl}');">��ʾ��һ��<%=SQL(3,K)%>URL</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetNextUrl}');">��ʾ��һ��<%=SQL(3,K)%>URL</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetShowComment}');">��ʾ����</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetWriteComment}');">��������</a></li>
			<%Case 7%>
			<div id="Content<%=SQL(6,K)%>" style="display:none">
				<li><a href="#" onClick="InsertLabel('{$GetMovieUrl}');" class="LabelItem">��ǰ<%=SQL(3,K)%> URL</a></li>
				<li><a href="#" onClick="InsertLabel('{$ChannelID}');" class="LabelItem">��ǰģ��ID</a></li>
				<li><a href="#" onClick="InsertLabel('{$InfoID}');" class="LabelItem">��ǰ<%=SQL(3,K)%>ID</a></li>
				<li><a href="#" onClick="InsertLabel('{$ItemName}');" class="LabelItem">��ǰ��Ŀ����</a></li>
				<li><a href="#" onClick="InsertLabel('{$ItemUnit}');" class="LabelItem">��ǰ��Ŀ��λ</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetMovieName}');" class="LabelItem"><%=SQL(3,K)%>����</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetMovieActor}');" class="LabelItem">��Ҫ��Ա</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetMovieDirector}');" class="LabelItem"><%=SQL(3,K)%>����</a> </li>
				<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>KS_Editor/KS_MoviePhoto.html',250,130);" class="LabelItem"><%=SQL(3,K)%>ͼƬ</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetMovieKeyWord}');" class="LabelItem">��ǰ<%=SQL(3,K)%>�ؼ���</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetKeyTags}');" class="LabelItem">ȡ��<%=SQL(3,K)%>Tags</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetMovieLanguage}');" class="LabelItem"><%=SQL(3,K)%>����</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetMovieArea}');" class="LabelItem">��������</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetMovieIntro}');" class="LabelItem">�鿴ӰƬ����</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetMovieTime}');" class="LabelItem"><%=SQL(3,K)%>���ȣ�����ʱ�䣩</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetScreenTime}');" class="LabelItem">��ӳʱ��</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetMovieDate}');" class="LabelItem">���(����)����</a></li>
				<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>KS_Editor/KS_MoviePlay.html',250,130);" class="LabelItem">�����б�</a></li>
				<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>KS_Editor/KS_MoviePage.html',250,130);" class="LabelItem">����ҳ������(�ʺ���flv,mtv��Ƶ������վ��)</a></li>
				<li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>KS_Editor/KS_MovieDown.html',250,130);" class="LabelItem">�����б�</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetHits}');" class="LabelItem"><%=SQL(3,K)%>����(�������)</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetHitsByDay}');" class="LabelItem"><%=SQL(3,K)%>���������</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetHitsByWeek}');" class="LabelItem"><%=SQL(3,K)%>���������</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetHitsByMonth}');" class="LabelItem"><%=SQL(3,K)%>���������</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetRank}');" class="LabelItem">��ʾ�Ƽ��Ǽ�</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetMovieInput}');" class="LabelItem"><%=SQL(3,K)%>¼��</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetPoint}');" class="LabelItem">ȡ�ùۿ�/���ص��</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetMovieProperty}');" class="LabelItem">��ʾӰ������(���š��������Ƽ��ȣ�</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetMovieVote}');" class="LabelItem">��ʾͶ��һƱ</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetMovieVoteScore}');" class="LabelItem">��ʾ<%=SQL(3,K)%>��Ʊ��</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetPrevMovie}');" class="LabelItem">��ʾ��һ��<%=SQL(3,K)%></a></li>
				<li><a href="#" onClick="InsertLabel('{$GetNextMovie}');" class="LabelItem">��ʾ��һ��<%=SQL(3,K)%></a></li>
						<li><a href="#" onClick="InsertLabel('{$GetPrevUrl}');">��ʾ��һ��<%=SQL(3,K)%>URL</a></li>
						<li><a href="#" onClick="InsertLabel('{$GetNextUrl}');">��ʾ��һ��<%=SQL(3,K)%>URL</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetShowComment}');">��ʾ����</a></li>
				<li><a href="#" onClick="InsertLabel('{$GetWriteComment}');">��������</a></li>
         <%Case 8%>
				<div id="Content<%=SQL(6,K)%>" style="display:none">
				  <li><a href="#" onClick="InsertLabel('{$GetGQInfoUrl}');" class="LabelItem">��ǰ<%=SQL(3,K)%> URL</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetGQInfoID}');" class="LabelItem">��ǰ<%=SQL(3,K)%>ID</a></li>
						<li><a href="#" onClick="InsertLabel('{$ChannelID}');" class="LabelItem">��ǰģ��ID</a></li>
						<li><a href="#" onClick="InsertLabel('{$InfoID}');" class="LabelItem">��ǰ<%=SQL(3,K)%>СID</a></li>
						<li><a href="#" onClick="InsertLabel('{$ItemName}');" class="LabelItem">��ǰ��Ŀ����</a></li>
						<li><a href="#" onClick="InsertLabel('{$ItemUnit}');" class="LabelItem">��ǰ��Ŀ��λ</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetGQTitle}');" class="LabelItem"><%=SQL(3,K)%>����</a></li>
				  <li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>KS_Editor/KS_SupplyPhoto.html',250,130);" class="LabelItem"><%=SQL(3,K)%>����ͼ(������)</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetGQKeyWord}');" class="LabelItem">ȡ�ùؼ���</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetKeyTags}');" class="LabelItem">ȡ��<%=SQL(3,K)%>Tags</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetPrice}');" class="LabelItem">�۸�˵��</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetInfoType}');" class="LabelItem">��Ϣ���</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetTransType}');" class="LabelItem">�������</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetValidTime}');" class="LabelItem">�� Ч ��</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetGQContent}');" class="LabelItem"><%=SQL(3,K)%>���ݽ���</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetHits}');" class="LabelItem"><%=SQL(3,K)%>����(�������)</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetHitsByDay}');" class="LabelItem"><%=SQL(3,K)%>���������</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetHitsByWeek}');" class="LabelItem"><%=SQL(3,K)%>���������</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetHitsByMonth}');" class="LabelItem"><%=SQL(3,K)%>���������</a></li>
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
				  <li><a href="#" onClick="InsertLabel('{$GetPrevUrl}');">��ʾ��һ��<%=SQL(3,K)%>URL</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetNextUrl}');">��ʾ��һ��<%=SQL(3,K)%>URL</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetShowComment}');">��ʾ������Ϣ</a></li>
				  <li><a href="#" onClick="InsertLabel('{$GetWriteComment}');">��������</a></li>
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
			  
	  <div onClick="ShowLabelTree('ChannelClassLabel')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'">
                <img src="../Images/Folder/folderclosed.gif" width="24" height="22" align="absmiddle"><a href="#">Ƶ������Ŀ��ר�ñ�ǩ</a>
	  </div>
				  <div id="ChannelClassLabel" style="display:none">  
				    <li><a href="#" onClick="InsertLabel('{$GetChannelID}');">��ʾ��ǰģ��ID</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetChannelName}');">��ʾ��ǰģ������</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetItemName}');" class="LabelItem">��ʾ��ǰģ�͵���Ŀ����</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetItemUnit}');" class="LabelItem">��ʾ��ǰģ�͵���Ŀ��λ</a></li>
				    =======================<br>    
				    <li><a href="#" onClick="InsertLabel('{$GetClassID}');">��ʾ��ǰ��ĿID</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetClassName}');">��ʾ��ǰ��Ŀ����</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetClassUrl}');" class="LabelItem">��ʾ��ǰ��Ŀ���ӵ�ַ</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetClassPic}');" class="LabelItem">��ʾ��ǰ��ĿͼƬ</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetClassIntro}');" class="LabelItem">��ʾ��ǰ��Ŀ����</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetClass_Meta_KeyWord}');" class="LabelItem">�����������Ĺؼ���</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetClass_Meta_Description}');" class="LabelItem">����������������</a></li>
				    <li><a href="#" onClick="InsertLabel('{$GetParentID}');">��ʾ����ĿID</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetParentUrl}');">��ʾ����Ŀ���ӵ�ַ</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetParentClassName}');">��ʾ����Ŀ����</a></li>
				 </div>
               
	  
       <div onClick="ShowLabelTree('SearchLabel')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'">
                <img src="../Images/Folder/folderclosed.gif" width="24" height="22" align="absmiddle"><a href="#">����ר�ñ�ǩ</a>
		</div>
				  <div id="SearchLabel" style="display:none">
				   <li><a href="#" onClick="InsertLabel('{$GetSearchByDate}');" class="LabelItem">�߼���������(С���)</a></li>
				   <li><a href="#" onClick="InsertLabel('{$GetSearch}');" class="LabelItem">��վ����</a></li>
				   <%
				   For K=0 To Ubound(SQL,2)
				    response.write "<li><a href=""#"" onClick=""InsertLabel('{$Get"  & SQL(6,K) & "Search}');"" class=""LabelItem"">" & SQL(2,K) & "����</a></li>"
				   Next
				   %>
				  </div>
			  

		
		 <div onClick="ShowLabelTree('MusicLabel')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'">
               <img src="../Images/Folder/folderclosed.gif" width="24" height="22" align="absmiddle"><a href="#">����Ƶ��ר�ñ�ǩ</a>
		</div>
             
			  <div id="MusicLabel" style="display:none">
                 <li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>KS_Editor/KS_GetMusicList.asp',500,230);" class="LabelItem">ȡ�ø��������б�(���¡��Ƽ����ȵ�-��վͨ��)</a></li>
				 <li><a href="#" onClick="InsertFunctionLabel('<%=DomainStr%>KS_Editor/KS_GetSpecialList.html',500,250);" class="LabelItem">ȡ��ר���б�(���¡��Ƽ����ȵ�-��վͨ��)</a></li>
				 <li><a href="#" onClick="InsertLabel('{$GetMusicNavi}');" class="LabelItem">ȡ�����ֶ�������(��վͨ��)</a></li>
				 <li><a href="#" onClick="InsertLabel('{$GetSingerType}');" class="LabelItem">ȡ�õ�ǰ������ƣ��绪���и��ֵ�(����ģ��ҳ����)</a></li>
				 <li><a href="#" onClick="InsertLabel('{$GetSingerList}');" class="LabelItem">ȡ�õ�ǰ����µ����и����б�(����ģ��ҳ����)</a></li>
				 <li><a href="#" onClick="InsertLabel('{$GetMusicSpecialList}');" class="LabelItem">ȡ�õ�ǰ���ֵ�ר���б�(����ר��ģ��ҳ����)</a></li>
				 <li><a href="#" onClick="InsertLabel('{$GetPagelist}');" class="LabelItem">ȡ�õ�ǰ���ֵ�ר����ҳ(����ר��ģ��ҳ����)</a></li>
				 <br><br>
				 <strong>===���±�ǩ������������ר������ҳģ��===</strong>
				 <br>
                 <li><a href="#" onClick="InsertLabel('{$GetSpecialID}');" class="LabelItem">ȡ��ר��ID</a></li>
				 <li><a href="#" onClick="InsertLabel('{$GetSpecialName}');" class="LabelItem">ȡ��ר������</a></li>
				 <li><a href="#" onClick="InsertLabel('{$GetSingerName}');" class="LabelItem">�ݳ�����</a></li>
				 <li><a href="#" onClick="InsertLabel('{$GetSpecialCompany}');" class="LabelItem">���й�˾</a></li>
				 <li><a href="#" onClick="InsertLabel('{$GetSpecialDate}');" class="LabelItem">��������</a></li>
				 <li><a href="#" onClick="InsertLabel('{$GetSpecialLanguage}');" class="LabelItem">��������</a></li>
				 <li><a href="#" onClick="InsertLabel('{$GetSpecialComment}');" class="LabelItem">ר������</a></li>
				 <li><a href="#" onClick="InsertLabel('{$GetSpecialSave}');" class="LabelItem">ר���ղ�</a></li>
				 <li><a href="#" onClick="InsertLabel('{$GetSpecialContent}');" class="LabelItem">ר������</a></li>
				 <li><a href="#" onClick="InsertLabel('{$GetSpecialPhoto}');" class="LabelItem">ר��ͼƬ��ַ</a></li>
				 <li><a href="#" onClick="InsertLabel('{$GetMusicPlayList}');" class="LabelItem">ר�����������б�</a></li>
                  <br><strong>========����ר������ҳ��ǩ========</strong>
           </div>

		
	    <div onClick="ShowLabelTree('AnnounceContent')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
               <img src="../Images/Folder/folderclosed.gif" width="24" height="22" align="absmiddle"><a href="#">��������ҳ��ǩ</a>
		</div>
		
		 <div id="AnnounceContent" style="display:none">
               <li><a href="#" onClick="InsertLabel('{$GetAnnounceTitle}');" class="LabelItem">�������</a></li>
			   <li><a href="#" onClick="InsertLabel('{$GetAnnounceAuthor}');" class="LabelItem">��������</a></li>
			   <li><a href="#" onClick="InsertLabel('{$GetAnnounceDate}');" class="LabelItem">���淢��(����)ʱ��</a></li>
			   <li><a href="#" onClick="InsertLabel('{$GetAnnounceContent}');" class="LabelItem">����ľ�������</a></li>
		 </div>
			   
	  	  <div onClick="ShowLabelTree('LinkContent')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'">
                    <img src="../Images/Folder/folderclosed.gif" width="24" height="22" align="absmiddle"><a href="#">��������ҳ��ǩ</a>
		 </div>
             <div id="LinkContent" style="display:none">
                   <li><a href="#" onClick="InsertLabel('{$GetLinkCommonInfo}');" class="LabelItem">��ʾ�鿴��ʽ�������������ӵ�</a></li>
				   <li><a href="#" onClick="InsertLabel('{$GetClassLink}');" class="LabelItem">��ʾ���༰��������վ������</a></li>
				   <li><a href="#" onClick="InsertLabel('{$GetLinkDetail}');" class="LabelItem">��ҳ��ʾ����������ϸ�б�</a></li>
			</div>

	  	  <div onClick="ShowLabelTree('Special')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'">
                    <img src="../Images/Folder/folderclosed.gif" width="24" height="22" align="absmiddle"><a href="#">ר��ҳ��ǩ</a>
		 </div>
             <div id="Special" style="display:none">
                   <li><a href="#" onClick="InsertLabel('{$GetSpecialName}');" class="LabelItem">��ǰר������</a></li>
				   <li><a href="#" onClick="InsertLabel('{$GetSpecialPic}');" class="LabelItem">��ǰר��ͼƬ</a></li>
				   <li><a href="#" onClick="InsertLabel('{$GetSpecialNote}');" class="LabelItem">��ǰר�����</a></li>
				   <li><a href="#" onClick="InsertLabel('{$GetSpecialDate}');" class="LabelItem">��ǰר�����ʱ��</a></li>
				   <li><a href="#" onClick="InsertLabel('{$GetSpecialClassName}');" class="LabelItem">��ǰר���������</a></li>
				   <li><a href="#" onClick="InsertLabel('{$GetSpecialClassURL}');" class="LabelItem">��ǰר�����URL</a></li>
			</div>

		
	  	  <div onClick="ShowLabelTree('UserSystem');" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'">
                    <img src="../Images/Folder/folderclosed.gif" width="24" height="22" align="absmiddle"><a href="#">��Աϵͳר�ñ�ǩ</a>
		  </div>
		  
          <div id="UserSystem" style="display:none">
					<li><a href="#" onClick="InsertLabel('{$GetTopUserLogin}');" class="LabelItem">��ʾ��Ա��¼���(����)</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetUserLogin}');" class="LabelItem">��ʾ��Ա��¼���(����)</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetAllUserList}');" class="LabelItem">��ʾ����ע���Ա�б�(�˱�ǩ����ʹ���ڻ�Ա�б�ҳģ��)</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetUserRegLicense}');" class="LabelItem">��ʾ�»�Աע��������������</a></li>
					<li><a href="#" onClick="InsertLabel('{$Show_UserNameLimitChar}');" class="LabelItem">��ʾ�»�Աע��ʱ�û��������ַ���</a></li>
					<li><a href="#" onClick="InsertLabel('{$Show_UserNameMaxChar}');" class="LabelItem">��ʾ�»�Աע��ʱ�û�������ַ���</a></li>
					<li><a href="#" onClick="InsertLabel('{$Show_VerifyCode}');" class="LabelItem">��ʾ�»�Աע��ʱ��֤��</a></li>
					<li><a href="#" onClick="InsertLabel('{$GetUserRegResult}');" class="LabelItem">�»�Աע��ɹ���Ϣ</a>
               </div>
		  
		 <div onClick="ShowLabelTree('AdwLabel');" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                <img src="../Images/Folder/folderclosed.gif" width="24" height="22" align="absmiddle"><a href="#">���λͨ�ñ�ǩ</a></div>
             <div id="AdwLabel" style="display:none">
				<%  Dim RSObj:Set RSObj=server.createobject("adodb.recordset")
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
                <img src="../Images/Folder/folderclosed.gif" width="24" height="22" align="absmiddle"><a href="#">��վ����ͨ�ñ�ǩ</a></div>
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
			 
			 
		 <div onClick="ShowLabelTree('RssLabel');" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                <img src="../Images/Folder/folderclosed.gif" width="24" height="22" align="absmiddle"><a href="#">RSS��ǩ</a>
		 </div>
            <div id="RssLabel" style="display:none">
				<li><a href="#" onClick="InsertLabel('{$Rss}');" class="LabelItem">Rss��ǩ��ʾ</a></li>
				<li><a href="#" onClick="InsertLabel('{$RssElite}');" class="LabelItem">Rss�Ƽ���ǩ��ʾ</a></li>
				<li><a href="#" onClick="InsertLabel('{$RssHot}');" class="LabelItem">Rss���ű�ǩ��ʾ</a></li>
			</div>

	  <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr ParentID=""> 
          <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr onClick="ShowLabelTree('DIYFunctionLabel')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                <td width="24"><img src="../Images/Folder/folderclosed.gif" width="24" height="22"></td>
                <td width="672"><a href="#">�û��Զ��庯����ǩ</a></td>
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
                                  <td width="3%"><img src="../Images/Folder/folderclosed.gif" width="24" height="22"></td>
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
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr ParentID=""> 
          <td><table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr onClick="ShowLabelTree('FreeLabel')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                <td width="24"><img src="../Images/Folder/folderclosed.gif" width="24" height="22"></td>
                <td width="672"><a href="#">�û��Զ��徲̬��ǩ</a></td>
              </tr>
              <tr> 
                <td colspan="2">
				 <table width="100%" border="0" cellspacing="0" cellpadding="0" id="FreeLabel" style="display:none">
                    <tr> 
                      <td width="8%" align="right">&nbsp;</td>
                      <td height="20"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                          <%
                          SqlStr = "Select * From KS_LabelFolder where FolderType=1 And ParentID='0'"
                         Set FolderRS = Conn.Execute(SqlStr)
                           if Not FolderRS.Eof then
	                    do while Not FolderRS.Eof
                           %>
                          <tr ParentID="<% = FolderRS("ParentID") %>" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                            <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                <tr> 
                                  <td width="3%"><img src="../Images/Folder/folderclosed.gif" width="24" height="22"></td>
                                  <td width="97%"><span ShowFlag="False" TypeID="<% = FolderRS("ID") %>" onClick="SelectFolder(this)"><A href="#">
                                    <% = FolderRS("FolderName") %>
                                    </A></span></td>
                                </tr>
                              </table></td>
                          </tr>
                          <%
	 		        Response.Write(GetChildFolderList(0,1,FolderRS("ID"),""," style=""display:none;"" "))
                    Response.Write(GetLabelList(1,trim(FolderRS("ID")),"&nbsp;&nbsp;&nbsp;&nbsp;"," style=""display:none;"" "))
		        FolderRS.MoveNext
	            loop
              end if
               Response.Write(GetLabelList(1,"0","",""))
              %>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
            </table> 
          </td>
        </tr>
      </table>

	   <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr ParentID=""> 
          <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr onClick="ShowLabelTree('SysJS')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                <td width="24"><img src="../Images/Folder/folderclosed.gif" width="24" height="22"></td>
                <td width="672"><a href="#">ϵͳJS��ǩ</a></td>
              </tr>
              <tr> 
                <td colspan="2"> <table width="100%" border="0" cellspacing="0" cellpadding="0" id="SysJS" style="display:none">
                    <tr> 
                      <td width="8%" align="right">&nbsp;</td>
                      <td height="20"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                          <%
                          SqlStr = "Select * From KS_LabelFolder where FolderType=2 And ParentID='0'"
                         Set FolderRS = Conn.Execute(SqlStr)
                           if Not FolderRS.Eof then
	                    do while Not FolderRS.Eof
                           %>
                          <tr ParentID="<% = FolderRS("ParentID") %>" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                            <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                <tr> 
                                  <td width="3%"><img src="../Images/Folder/folderclosed.gif" width="24" height="22"></td>
                                  <td width="97%"><span ShowFlag="False" TypeID="<% = FolderRS("ID") %>" onClick="SelectFolder(this)"><A href="#"> 
                                    <% = FolderRS("FolderName") %>
                                    </A></span></td>
                                </tr>
                              </table></td>
                          </tr>
                          <%
	 		        Response.Write(GetChildFolderList(1,0,FolderRS("ID"),""," style=""display:none;"" "))
                    Response.Write(GetJSList(0,trim(FolderRS("ID")),"&nbsp;&nbsp;&nbsp;&nbsp;"," style=""display:none;"" "))
		        FolderRS.MoveNext
	            loop
              end if
               Response.Write(GetJSList(0,"0","",""))
              %>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
      </table>
      <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr ParentID=""> 
          <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
              <tr onClick="ShowLabelTree('JSLabel')" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                <td width="24"><img src="../Images/Folder/folderclosed.gif" width="24" height="22"></td>
                <td width="672"><a href="#">����JS��ǩ</a></td>
              </tr>
              <tr> 
                <td colspan="2">
				  <table width="100%" border="0" cellspacing="0" cellpadding="0" id="JSLabel" style="display:none">
                    <tr> 
                      <td width="8%" align="right">&nbsp;</td>
                      <td height="20"> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                          <%
                          SqlStr = "Select * From KS_LabelFolder where FolderType=3 And ParentID='0'"
                         Set FolderRS = Conn.Execute(SqlStr)
                           if Not FolderRS.Eof then
	                    do while Not FolderRS.Eof
                           %>
                          <tr ParentID="<% = FolderRS("ParentID") %>" onMouseOut="this.style.backgroundColor=''" onMouseOver="this.style.backgroundColor='#EEEEEE'"> 
                            <td> <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                <tr> 
                                  <td width="3%"><img src="../Images/Folder/folderclosed.gif" width="24" height="22"></td>
                                  <td width="97%"><span ShowFlag="False" TypeID="<% = FolderRS("ID") %>" onClick="SelectFolder(this)"><A href="#"> 
                                    <% = FolderRS("FolderName") %>
                                    </A></span></td>
                                </tr>
                              </table></td>
                          </tr>
                          <%
	 		        Response.Write(GetChildFolderList(1,1,FolderRS("ID"),""," style=""display:none;"" "))
                    Response.Write(GetJSList(1,trim(FolderRS("ID")),"&nbsp;&nbsp;&nbsp;&nbsp;"," style=""display:none;"" "))
		        FolderRS.MoveNext
	            loop
              end if
               Response.Write(GetJSList(1,"0","",""))
              %>
                        </table></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
      </table>
	 
</td>
  </tr>
  <tr>
    <td height="90" valign="top">
	<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td height="22" bgcolor="#0000FF"> 
            <div align="center"><font color="#FFFFFF"><strong>��ǩ˵��</strong></font></div></td>
        </tr>
        <tr> 
          <td valign="top" bgcolor="#efefef"> 
            <table width="272" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="143" height="25"> <img src="../Images/label2.gif" width="17" height="15"> 
                  ��վͨ�ñ�ǩ</td>
                <td width="129" height="25"><img src="../Images/label1.gif" width="17" height="15"> 
                  Ƶ����ͨ�ñ�ǩ</td>
              </tr>
              <tr> 
                <td><img src="../Images/label0.gif"> ��ͨ����������ǩ</td>
                <td><img src="../Images/label3.gif"> �Զ��徲̬��ǩ </td>
              </tr>
              <tr> 
                <td><img src="../Images/JS0.gif" align="absmiddle"> ϵͳJS��ǩ</td>
                <td><img src="../Images/JS1.gif" align="absmiddle"> ����JS��ǩ</td>
              </tr>
            </table></td>
        </tr>
      </table></td>
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
		GetLabelList = GetLabelList & "<table border=""0"" cellspacing=""0"" cellpadding=""0""><tr><td>" & CompatStr &  "<img src=""../Images/Label" & trim(LabelRS("LabelFlag")) & ".gif""></td>"
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
		GetJSList = GetJSList & "<table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf & "<tr>"  & vbcrlf & "<td>" & CompatStr &  "<img src=""../Images/JS" & trim(JSType) & ".gif""></td>"
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
		GetChildFolderList = GetChildFolderList & "<table border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbcrlf & "<tr>"  & vbcrlf & "<td>" & TempStr & "<img src=""../Images/Folder/folderclosed.gif""></td>"
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
  {case 'TY':
     if (document.all.TY.style.display!='')
       {document.all.TY.style.display='';}
     else
      {document.all.TY.style.display='none';} 
	  break;
	case 'CommonJSLabel':
     if (document.all.CommonJSLabel.style.display!='')
       {document.all.CommonJSLabel.style.display='';}
     else
      {document.all.CommonJSLabel.style.display='none';} 
	  break;
    case 'ChannelClassLabel':
     if (document.all.ChannelClassLabel.style.display!='')
       {document.all.ChannelClassLabel.style.display='';}
     else
      {document.all.ChannelClassLabel.style.display='none';} 
	  break;
   case 'SearchLabel':
        if (document.all.SearchLabel.style.display!='')
       {document.all.SearchLabel.style.display='';}
     else
      {document.all.SearchLabel.style.display='none';} 
	  break;
  <%For K=0 To Ubound(SQL,2)%>
   case 'Content<%=SQL(6,K)%>':
     if (document.all.Content<%=SQL(6,K)%>.style.display!='')
       {document.all.Content<%=SQL(6,K)%>.style.display='';}
     else
      {document.all.Content<%=SQL(6,K)%>.style.display='none';} 
	  break;
   <%Next%>
  case 'MusicLabel':
     if (document.all.MusicLabel.style.display!='')
       {document.all.MusicLabel.style.display='';}
     else
      {document.all.MusicLabel.style.display='none';} 
	  break;
   case 'AnnounceContent':
     if (document.all.AnnounceContent.style.display!='')
       {document.all.AnnounceContent.style.display='';}
     else
      {document.all.AnnounceContent.style.display='none';} 
	  break;
   case 'SysFLabel' :
   	  if (document.all.SysFLabel.style.display!='')
       {document.all.SysFLabel.style.display='';}
     else
      {document.all.SysFLabel.style.display='none';} 
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
  case 'DIYFieldLabel' :
      if (document.all.DIYFieldLabel.style.display!='')
      {document.all.DIYFieldLabel.style.display='';}
     else
      {document.all.DIYFieldLabel.style.display='none';} 
	  break; 
  case 'JSLabel' :
       if (document.all.JSLabel.style.display!='')
      {document.all.JSLabel.style.display='';}
     else
      {document.all.JSLabel.style.display='none';} 
	  break; 	   	  
  case 'SysJS' :
        if (document.all.SysJS.style.display!='')
      {document.all.SysJS.style.display='';}
     else
      {document.all.SysJS.style.display='none';} 
	  break; 
  case 'LinkContent':	   
        if (document.all.LinkContent.style.display!='')
      {document.all.LinkContent.style.display='';}
     else
      {document.all.LinkContent.style.display='none';} 
	  break; 
 case 'UserSystem':
      if (document.all.UserSystem.style.display!='')
      {document.all.UserSystem.style.display='';}
     else
      {document.all.UserSystem.style.display='none';} 
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
 case 'RssLabel':
      if (document.all.RssLabel.style.display!='')
      {document.all.RssLabel.style.display='';}
     else
      {document.all.RssLabel.style.display='none';} 
	  break; 
 case 'ContentLabel':
      if (document.all.ContentLabel.style.display!='')
      {document.all.ContentLabel.style.display='';}
     else
      {document.all.ContentLabel.style.display='none';} 
	  break; 
 case 'Special':
      if (document.all.Special.style.display!='')
      {document.all.Special.style.display='';}
     else
      {document.all.Special.style.display='none';} 
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
