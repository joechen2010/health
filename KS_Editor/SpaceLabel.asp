<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML><HEAD><TITLE>�ռ��ǩ</TITLE>
<META content="text/html; charset=gb2312" http-equiv=Content-Type>
<link href="editor.css" rel="stylesheet">
<style>
td{font-size:12px;}
body{background:#FFFFFF}
a{text-decoration:none;font-size:12px;color:#000000}
li{list-style-type:circle}
</style>
</HEAD>
<body>
<br>
<table align="center" width="95%" border="0" cellspacing="0" cellpadding="0">
		 <tr>
		  <td width="150" colspan=2><font color=red>���ñ�ǩ˵��</font></td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$BlogMain}');"><strong>{$BlogMain}</strong></a></td><td colspan=3>---��ʾ��־���岿��(��������������ҳģ��)��</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowUserInfo}');">{$ShowUserInfo}</a></td><td>---��ʾ�û���Ϣ��</td>
		     <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowComment}');">{$ShowComment}</a></td><td>---��ʾ���������б�</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowUserClass}');">{$ShowUserClass}</a></td><td>---��ʾר�������б�</td>
		  <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowMessage}');">{$ShowMessage}</a></td><td>---��ʾ���������б�</td>
		  </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowBlogInfo}');">{$ShowBlogInfo}</a></td><td>---��ʾ������־�б�</td>
		   <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowAnnounce}');">{$ShowAnnounce}</a></td><td>---��ʾ���¹��档</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowBlogName}');">{$ShowBlogName}</a></td><td>---��ʾ����վ�����ơ�</td>
		   <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowCalendar}');">{$ShowCalendar}</a></td><td>---��ʾ����������</td>
		 </tr>
		
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowNavigation}');">{$ShowNavigation}</a></td><td>---��ʾ��ҳ�����ȡ�</td>
		   <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowBlogTotal}');">{$ShowBlogTotal}</a></td><td>---��ʾͳ����Ϣ�ȡ�</td>
		 </tr>
		 
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowMusicBox}');">{$ShowMusicBox}</a></td><td>---��ʾ���ֲ�������</td>
		    <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowUserLogin}');">{$ShowUserLogin}</a></td><td>---��ʾ��Ա��¼��</td>
		 </tr>
		
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowSearch}');">{$ShowSearch}</a></td><td>---��ʾ������־��</td>
		   <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowXML}');">{$ShowXML}</a></td><td>---��ʾRSS���ġ�</td>
		 </tr>
		
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowUserName}');">{$ShowUserName}</a></td><td>---��ʾ�û�����</td>
		   <td><li><a href='#' onClick="parent.InsertLabel('TemplateMain','{$ShowLogo}');">{$ShowLogo}</a></td><td>---��ʾLogo��</td>
		 </tr>
		
		 <tr>
	       <td colspan=4><li>{$ShowBannerSrc1},{$ShowBannerSrc2},{$ShowBannerSrc3}    ---��ʾBannerͼƬ��ַ��</td>
		 </tr>
		
		 <tr>
		  <td width="150"><li><a href='#' onClick="parent.InsertLabel('TemplateSub0','{$ShowNewLog}');">{$ShowNewLog}</a></td><td>---����1ƪ��־</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub0','{$ShowNewAlbum}');">{$ShowNewAlbum}</a></td><td>---����3�����</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub0','{$ShowNewInfo}');">{$ShowNewInfo}</a></td><td>---10����Ϣ��</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub0','{$ShowVisitor}');">{$ShowVisitor}</a></td><td>---���·ÿ�</td>
		 </tr>

		<%if request("flag")="4" then%> 
		 
		 <tr>
	       <td colspan=4><br>================��ҵ�ռ�ר��=======================</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub0','{$ShowShortIntro}');" title="��ʾ580���ֵ���ҵ����">{$ShowShortIntro}</a></td><td>---��ҵ���(��)</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub0','{$ShowIntro}');">{$ShowIntro}</a></td><td>---��ҵ���</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub0','{$ShowNews}');">{$ShowNews}</a></td><td>---��ҵ��̬</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub0','{$ShowSupply}');">{$ShowSupply}</a></td><td>---��Ӧ��Ϣ</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub0','{$ShowProduct}');" title='һ����ʾ4������������ʾ���²�Ʒ'>{$ShowProduct}</a></td><td>---���²�Ʒ</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub0','{$ShowProductList}');" title='��������ʾ���²�Ʒ'>{$ShowProductList}</a></td><td>---�ı���ʽ��ʾ���²�Ʒ</td>
		 </tr>
		 <%end if%>
</table>

        <%response.end%>
		 <tr>
		  <td colspan=2><font color=red>��ģ��(��־)���ñ�ǩ˵��</font></td>
		 </tr>
		 <tr>
		  <td width="150"><li><a href='#' onClick="parent.InsertLabel('TemplateSub1','{$ShowLogTopic}');">{$ShowLogTopic}</a></td><td>---��ʾ���鼰��־����</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub1','{$ShowLogInfo}');">{$ShowLogInfo}</a></td><td>---��ʾ����ʱ�䡢���ߵ���Ϣ</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub1','{$ShowLogText}');">{$ShowLogText}</a></td><td>---��ʾ��־����</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub1','{$ShowLogMore}');">{$ShowLogMore}</a></td><td>---��ʾ�Ķ�ȫ��(����)���ظ�(����)����������</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub1','{$ShowTopic}');">{$ShowTopic}</a></td><td>---����ʾ��־����</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub1','{$ShowAuthor}');">{$ShowAuthor}</a></td><td>---����ʾ��־����</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub1','{$ShowAddDate}');">{$ShowAddDate}</a></td><td>---����ʾ��־����ʱ��</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub1','{$ShowEmot}');">{$ShowEmot}</a></td><td>---����ʾ��־����</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub1','{$ShowWeather}');">{$ShowWeather}</a></td><td>---����ʾ��־����</td>
		 </tr>
		  <tr>
		  <td colspan=2><font color=red>��ģ��(���˵���)���ñ�ǩ˵��</font></td>
		 </tr>
		 <tr>
		  <td width="150"><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetUserName}');">{$GetUserName}</a></td><td>--�û������ǳ�)</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetRealName}');">{$GetRealName}</a></td><td>---��ʵ����</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetSex}');">{$GetSex}</a></td><td>---�Ա�</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetBirthday}');">{$GetBirthday}</a></td><td>---��������</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetIDCard}');">{$GetIDCard}</a></td><td>---���֤��</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetOfficeTel}');">{$GetOfficeTel}</a></td><td>---�칫�绰</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetHomeTel}');">{$GetHomeTel}</a></td><td>---��ͥ�绰</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetMobile}');">{$GetMobile}</a></td><td>---�ֻ�����</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetFax}');">{$GetFax}</a></td><td>---�������</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetUserArea}');">{$GetUserArea}</a></td><td>---���ڵ���</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetAddress}');">{$GetAddress}</a></td><td>---��ϵ��ַ</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetZip}');">{$GetZip}</a></td><td>---��������</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetHomePage}');">{$GetHomePage}</a></td><td>---������ҳ</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetUserFace}');">{$GetUserFace}</a></td><td>---�û�ͷ��</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetEmail}');">{$GetEmail}</a></td><td>---��������</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetQQ}');">{$GetQQ}</a></td><td>---QQ ����</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetICQ}');">{$GetICQ}</a></td><td>---ICQ����</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetMSN}');">{$GetMSN}</a></td><td>---MSN�˺�</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetUC}');">{$GetUC}</a></td><td>---UC����</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetSign}');">{$GetSign}</a></td><td>---����ǩ��</td>
		 </tr>
		 <tr>
		  <td colspan=2><font color=red>��ģ��(��ϵ����)���ñ�ǩ˵��</font></td>
		 </tr>
		 <tr>
		  <td width="150"><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetCompanyName}');">{$GetCompanyName}</a></td><td>--��˾����</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetBusinessLicense}');">{$GetBusinessLicense}</a></td><td>---Ӫҵִ��</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetProfession}');">{$GetProfession}</a></td><td>---��˾��ҵ</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetOfficeTel}');">{$GetLegalPeople}</a></td><td>---��ҵ����</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetCompanyScale}');">{$GetCompanyScale}</a></td><td>---��˾��ģ</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetRegisteredCapital}');">{$GetRegisteredCapital}</a></td><td>---ע���ʽ�</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetProvince}');">{$GetProvince}</a></td><td>---����ʡ��</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetCity}');">{$GetCity}</a></td><td>---���ڳ���</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetContactMan}');">{$GetContactMan}</a></td><td>---�� ϵ ��</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetAddress}');">{$GetAddress}</a></td><td>---��˾��ַ</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetZipCode}');">{$GetZipCode}</a></td><td>---��������</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetTelphone}');">{$GetTelphone}</a></td><td>---��ϵ�绰</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetFax}');">{$GetFax}</a></td><td>---�������</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetWebUrl}');">{$GetWebUrl}</a></td><td>---��˾��ַ</td>
		 </tr>
		 <tr>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetBankAccount}');">{$GetBankAccount}</a></td><td>---��������</td>
	       <td><li><a href='#' onClick="parent.InsertLabel('TemplateSub2','{$GetAccountNumber}');">{$GetAccountNumber}</a></td><td>---�����˺�</td>
		 </tr>

		 </table>

</BODY>
</HTML>
 
