<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%option explicit%>
<!--#include file="../../Conn.asp"-->
<!--#include file="../../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="../include/Session.asp"-->
<!--#include file="../../Plus/md5.asp"-->
<%
Dim Chk:Set Chk=New LoginCheckCls1
Chk.Run()
Set Chk=Nothing

Dim KS:Set KS=New PublicCls

If Not KS.ReturnPowerResult(0, "KSO10003") Then
   Response.Write ("<script>parent.frames['BottomFrame'].location.href='javascript:history.back();';</script>")
   Call KS.ReturnErr(1, "")
   Response.End()
End If



Call SetSystem()

Call CloseConn()
Set KS=Nothing
	

Sub SetSystem()
    on error resume next
	Dim CurrPath:CurrPath=KS.Setting(3)&KS.Setting(90)
	If Right(CurrPath,1)="/" Then CurrPath=Left(CurrPath,Len(CurrPath)-1)
	Dim SqlStr, RS
	SqlStr = "select WapSetting from KS_Config"
	Set RS = Server.CreateObject("ADODB.recordset")
	RS.Open SqlStr, Conn, 1, 3
	Dim WapSetting:WapSetting=Split(RS(0),"^%^")
	If KS.G("Flag") = "Edit" Then
	   Dim N					
	   Dim WebSetting
	   For N=0 To 50
	       WebSetting=WebSetting & Replace(KS.G("WapSetting(" & N &")"),"^%^","") & "^%^"
	   Next
	   RS("WapSetting")=WebSetting
	   RS.Update
	   Call KS.DelCahe(KS.SiteSn & "_Config")
	   Call KS.DelCahe(KS.SiteSn & "_Date")
	   Response.Write "<script>alert('WAP���������޸ĳɹ���');location.href='KS_System.asp';</script>"				
	End If
	%>
    <html>
    <title>WAP������������</title>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<script src="Include/Common.js" language="JavaScript"></script>
    <script src="../../ks_inc/Common.js" language="JavaScript"></script>
	<script src="../../ks_inc/jquery.js" language="JavaScript"></script>
	<script src="../Images/pannel/tabpane.js" language="JavaScript"></script>
    <link href="../Images/pannel/tabpane.CSS" rel="stylesheet" type="text/css">
    <link href="../Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
	<style type="text/css">
	<!--
	.STYLE1 {color: #FF0000}
	.STYLE2 {color: #FF6600}
	-->
    </style>
    </head>

    <body bgcolor="#FFFFFF" topmargin="0" leftmargin="0">
    
    <div class="topdashed sort">WAP������������</div>
    <br/>
    <div class="tab-page" id="spaceconfig">
    <form name="myform" id="myform" method="post" Action="" onSubmit="return(CheckForm())">
	<script type=text/javascript>var tabPane1 = new WebFXTabPane( document.getElementById( "spaceconfig" ), 1 )</script>
    
    <div class=tab-page id=site-page>
    <H2 class=tab>��������</H2>
	<script type=text/javascript>
	tabPane1.addTabPage( document.getElementById( "site-page" ) );
    </script>
    <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
    <input type="hidden" value="Edit" name="Flag">
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>����״̬��</strong></div></td>
    <td width="63%" height="30"><input  name="WapSetting(0)" type="radio" value="1" <%If WapSetting(0)="1" Then Response.Write" Checked"%>>����
    <input name="WapSetting(0)" type="radio" value="0" <%If WapSetting(0)="0" Then Response.Write" Checked"%>>�ر�</td>
    </tr>

    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>�Ƿ�����ģ�������ʣ�</strong></div></td>
    <td width="63%" height="30"><input  name="WapSetting(1)" type="radio" value="1"<%If WapSetting(1)="1" Then Response.Write" Checked"%>>����
    <input name="WapSetting(1)" type="radio" value="0"<%If WapSetting(1)="0" Then Response.Write" Checked"%>>�ر�</td>
    </tr>

    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>�Ƿ�����վ���������</strong></div></td>
    <td width="63%" height="30"><input  name="WapSetting(6)" type="radio" value="1"<%If WapSetting(6)="1" Then Response.Write" Checked"%>>����
    <input name="WapSetting(6)" type="radio" value="0"<%If WapSetting(6)="0" Then Response.Write" Checked"%>>�ر�</td>
    </tr>

    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>�û���½ʶ�������</strong></div></td>
    <td width="63%" height="30"><input name="WapSetting(2)" type="text" value="<%=WapSetting(2)%>" size="30"><font color="red">* ������"wap.asp?wap=f88246e4150aef17ab1176fc27af920b��wap��</font></td>
    </tr>

    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>��վ���ƣ�</strong></div></td>
    <td width="63%" height="30"><input name="WapSetting(3)" type="text" value="<%=WapSetting(3)%>" size="30"></td>
    </tr>
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>��װĿ¼��</strong></div></td>
    <td width="63%" height="30"><input name="WapSetting(4)" type="text" value="<%=WapSetting(4)%>" size="30"><font color="red">* WAP�����װ������Ŀ¼���硰wap/��</font></td>
    </tr>    

    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>��վLogo��ַ��</strong></div></td>
    <td width="63%" height="30"><input name="WapSetting(5)" type="text" value="<%=WapSetting(5)%>" size="30"></td>
    </tr>

    <tr style="display:none" valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>��ҳ�ļ�����</strong></div></td>
    <td width="63%" height="30"><input name="WapSetting(7)" type="text" value="<%=WapSetting(7)%>" size="30"><font color="red">* ��չ��Ϊ.asp���硰wap.asp��</font></td>
    </tr>

    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>�ײ���Ȩ��Ϣ��</strong></div></td>
    <td width="63%" height="30"><textarea name="WapSetting(8)" cols="40" rows="4"><%=WapSetting(8)%></textarea><font color=red>* ֧��WAP��������,��ʾͼƬ,���У��﷨</font></td>
    </tr>
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>���������ã�</strong></div></td>
    <td width="63%" height="30">�����û���ID <input name="WapSetting(12)" type="text" value="<%=WapSetting(12)%>" size="4"><font color="red">* ���ȵ�""�û�->�û������""�������Ӳ鿴����</font><br/>
    ������λID <input name="WapSetting(13)" type="text" value="<%=WapSetting(13)%>" size="4"><font color="red">* ���ȵ�""��ϵͳ->���ϵͳ����""�������Ӳ鿴����</font></td>
    </tr>


    </table>
    </div>
    
    <div class="tab-page" id="template-page">
    <H2 class="tab">ģ������</H2>
	<script type=text/javascript>tabPane1.addTabPage( document.getElementById( "template-page" ) );</script>
    <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>��վ��ҳģ�壺</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(9)' id='WapSetting9' value="<%=WapSetting(9)%>">&nbsp;<input type='button' name="Submit" class="button" value="ѡ��ģ��..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("����ģ��")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting9').get(0));">

    
    </td>
    </tr>

    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>ȫվTagsģ�壺</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(10)' id='WapSetting10' value="<%=WapSetting(10)%>">&nbsp;<input type='button' name="Submit" class="button" value="ѡ��ģ��..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("����ģ��")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting10')[0]);"></td>
    </tr>
    
    <tr style="display:none" valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>ר����ҳģ�壺</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(11)' id='WapSetting11' value="<%=WapSetting(11)%>">&nbsp;<input type='button' name="Submit" class="button" value="ѡ��ģ��..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("����ģ��")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting11')[0]);"></td>
    </tr>  


    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>С��̳��ҳģ�壺</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(31)' id='WapSetting31' value="<%=WapSetting(31)%>">&nbsp;<input type='button' name="Submit" class="button" value="ѡ��ģ��..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("����ģ��")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting31')[0]);"></td>
    </tr>  


    <tr valign="middle" style="display:none" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>С��̳����ģ�壺</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(32)' id='WapSetting32' value="<%=WapSetting(32)%>">&nbsp;<input type='button' name="Submit" class="button" value="ѡ��ģ��..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("����ģ��")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting32')[0]);"></td>
    </tr>  

    <tr><td colspan=2 height='1' bgcolor='green'></td></tr>
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>��Ա��ҳģ�壺</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(19)' id='WapSetting19' value="<%=WapSetting(19)%>">&nbsp;<input type='button' name="Submit" class="button" value="ѡ��ģ��..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("����ģ��")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting19')[0]);"></td>
    </tr> 
    <tr><td colspan=2 height='1' bgcolor='green'></td></tr>
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>�ռ���ҳģ�壺</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(29)' id='TemplateID' value="<%=WapSetting(29)%>">&nbsp;<input type='button' name="Submit" class="button" value="ѡ��ģ��..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("����ģ��")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('WapSetting(29)'));"></td>
    </tr> 
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>�ռ丱ģ�壺</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(30)' id='TemplateID' value="<%=WapSetting(30)%>">&nbsp;<input type='button' name="Submit" class="button" value="ѡ��ģ��..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("����ģ��")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('WapSetting(30)'));"></td>
    </tr> 
    
    <tr><td colspan=2 height='1' bgcolor='green'></td></tr>

    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>���˿ռ���ģ�壺</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(20)' id='WapSetting20' value="<%=WapSetting(20)%>">&nbsp;<input type='button' name="Submit" class="button" value="ѡ��ģ��..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("����ģ��")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting20')[0]);"></td>
    </tr>
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>���˿ռ���ҳ��ģ�壺</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(21)' id='WapSetting21' value="<%=WapSetting(21)%>">&nbsp;<input type='button' name="Submit" class="button" value="ѡ��ģ��..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("����ģ��")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting21')[0]);"></td>
    </tr>
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>���˿ռ�С������ģ�壺</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(22)' id='WapSetting22' value="<%=WapSetting(22)%>">&nbsp;<input type='button' name="Submit" class="button" value="ѡ��ģ��..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("����ģ��")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting22')[0]);"></td>
    </tr>  

    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>���˿ռ���־��ģ�壺</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(28)' id='WapSetting28' value="<%=WapSetting(28)%>">&nbsp;<input type='button' name="Submit" class="button" value="ѡ��ģ��..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("����ģ��")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting28')[0]);"></td>
    </tr>  

    <tr><td colspan=2 height='1' bgcolor='green'></td></tr>

    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>��ҵ�ռ���ģ�壺</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(23)' id='WapSetting23' value="<%=WapSetting(23)%>">&nbsp;<input type='button' name="Submit" class="button" value="ѡ��ģ��..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("����ģ��")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting23')[0]);"></td>
    </tr>
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>��ҵ�ռ���ҳ��ģ�壺</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(24)' id='WapSetting24' value="<%=WapSetting(24)%>">&nbsp;<input type='button' name="Submit" class="button" value="ѡ��ģ��..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("����ģ��")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting24')[0]);"></td>
    </tr>
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>��ҵ�ռ�С������ģ�壺</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(25)' id='WapSetting25' value="<%=WapSetting(25)%>">&nbsp;<input type='button' name="Submit" class="button" value="ѡ��ģ��..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("����ģ��")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting25')[0]);"></td>
    </tr>  
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>��ҵ�ռ���־��ģ�壺</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(27)' id='WapSetting27' value="<%=WapSetting(27)%>">&nbsp;<input type='button' name="Submit" class="button" value="ѡ��ģ��..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("����ģ��")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting27')[0]);"></td>
    </tr>  
    
    <tr><td colspan=2 height='1' bgcolor='green'></td></tr>
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>Ȧ����ģ�壺</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(26)' id='WapSetting26' value="<%=WapSetting(26)%>">&nbsp;<input type='button' name="Submit" class="button" value="ѡ��ģ��..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("����ģ��")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting26')[0]);"></td>
    </tr>  
    </table>
    </div>
    
    <div class="tab-page" id="cardonline-page">
    <H2 class="tab">�ӿ�����</H2>
	<script type=text/javascript>tabPane1.addTabPage( document.getElementById( "cardonline-page" ) );</script>
    <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">

    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>ƽ̨���ƣ�</strong></div></td>
    <td width="63%" height="30"><Input name="WapSetting(15)" value="<%=WapSetting(15)%>"></td>
    </tr>
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>��ע˵����</strong></div></td>
    <td width="63%" height="30"><textarea name="WapSetting(16)" cols="40" rows="4"><%=WapSetting(16)%></textarea></td>
    </tr>

    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>֧����ţ�</strong></div>���������۷�ƽ̨������̻����</td>
    <td width="63%" height="30"><Input name="WapSetting(17)" value="<%=WapSetting(17)%>">&nbsp;<a href="http://www.spvnow.com/zhuce.asp" target="_blank"><strong><font color="red">�����ʺ�</font></strong></a></td>
    </tr>

    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>֧����Կ��</strong></div>���������������۷�ƽ̨�����õ�MD5˽Կ</td>
    <td width="63%" height="30"><Input name="WapSetting(18)" value="<%=WapSetting(18)%>"></td>
    </tr>
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>�Ƿ�����:</strong></div></td>
    <td width="63%" height="30"><input type="radio" value="0" name="WapSetting(14)">����
						<input type="radio" value="1" name="WapSetting(14)" checked>����</td>
    </tr>
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>ͬ����ַ:</strong></div>����ַ���Ƶ��۷�ƽ̨������ͬ���ش�</td>
    <td width="63%" height="30"><input name="CardReceive" type="text" value="<%=KS.Setting(2)&KS.Setting(3)&WapSetting(4)%>User/User_CardReceive.asp" size="60" DISABLED></td>
    </tr>  
    
    </table>
    </div>
    
    
    
    </body>
    </html>
	<script Language="javascript">
	<!--
	function CheckForm(){
		if ($('#WapSetting9').val()=='')
		   { alert('��ѡ��WAP��վ��ҳģ��!');
		   $('#WapSetting9').focus();
		   return false;
		   }
		if ($('#WapSetting10').val()=='')
		   { alert('��ѡ��WAP��������ҳģ��!');
		   $('#WapSetting10').focus();
		   return false;
			}
		if ($('#WapSetting11').val()=='')
		   { alert('��ѡ��WAP���԰���ҳģ��!');
		    $('#WapSetting11').focus();
		   return false;
			}
			$('#myform').submit();
	     }
	//-->
    </script>
    <%
	RS.Close:Set RS = Nothing
End Sub
%> 
