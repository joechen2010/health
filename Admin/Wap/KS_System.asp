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
	   Response.Write "<script>alert('WAP基本参数修改成功！');location.href='KS_System.asp';</script>"				
	End If
	%>
    <html>
    <title>WAP基本参数设置</title>
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
    
    <div class="topdashed sort">WAP基本参数设置</div>
    <br/>
    <div class="tab-page" id="spaceconfig">
    <form name="myform" id="myform" method="post" Action="" onSubmit="return(CheckForm())">
	<script type=text/javascript>var tabPane1 = new WebFXTabPane( document.getElementById( "spaceconfig" ), 1 )</script>
    
    <div class=tab-page id=site-page>
    <H2 class=tab>参数设置</H2>
	<script type=text/javascript>
	tabPane1.addTabPage( document.getElementById( "site-page" ) );
    </script>
    <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
    <input type="hidden" value="Edit" name="Flag">
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>功能状态：</strong></div></td>
    <td width="63%" height="30"><input  name="WapSetting(0)" type="radio" value="1" <%If WapSetting(0)="1" Then Response.Write" Checked"%>>开启
    <input name="WapSetting(0)" type="radio" value="0" <%If WapSetting(0)="0" Then Response.Write" Checked"%>>关闭</td>
    </tr>

    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>是否启用模拟器访问：</strong></div></td>
    <td width="63%" height="30"><input  name="WapSetting(1)" type="radio" value="1"<%If WapSetting(1)="1" Then Response.Write" Checked"%>>开启
    <input name="WapSetting(1)" type="radio" value="0"<%If WapSetting(1)="0" Then Response.Write" Checked"%>>关闭</td>
    </tr>

    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>是否启用站点计数器：</strong></div></td>
    <td width="63%" height="30"><input  name="WapSetting(6)" type="radio" value="1"<%If WapSetting(6)="1" Then Response.Write" Checked"%>>开启
    <input name="WapSetting(6)" type="radio" value="0"<%If WapSetting(6)="0" Then Response.Write" Checked"%>>关闭</td>
    </tr>

    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>用户登陆识别变量：</strong></div></td>
    <td width="63%" height="30"><input name="WapSetting(2)" type="text" value="<%=WapSetting(2)%>" size="30"><font color="red">* 变量如"wap.asp?wap=f88246e4150aef17ab1176fc27af920b中wap”</font></td>
    </tr>

    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>网站名称：</strong></div></td>
    <td width="63%" height="30"><input name="WapSetting(3)" type="text" value="<%=WapSetting(3)%>" size="30"></td>
    </tr>
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>安装目录：</strong></div></td>
    <td width="63%" height="30"><input name="WapSetting(4)" type="text" value="<%=WapSetting(4)%>" size="30"><font color="red">* WAP插件安装的虚拟目录，如“wap/”</font></td>
    </tr>    

    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>网站Logo地址：</strong></div></td>
    <td width="63%" height="30"><input name="WapSetting(5)" type="text" value="<%=WapSetting(5)%>" size="30"></td>
    </tr>

    <tr style="display:none" valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>首页文件名：</strong></div></td>
    <td width="63%" height="30"><input name="WapSetting(7)" type="text" value="<%=WapSetting(7)%>" size="30"><font color="red">* 扩展名为.asp，如“wap.asp”</font></td>
    </tr>

    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>底部版权信息：</strong></div></td>
    <td width="63%" height="30"><textarea name="WapSetting(8)" cols="40" rows="4"><%=WapSetting(8)%></textarea><font color=red>* 支持WAP（超链接,显示图片,换行）语法</font></td>
    </tr>
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>共享功能设置：</strong></div></td>
    <td width="63%" height="30">共享用户组ID <input name="WapSetting(12)" type="text" value="<%=WapSetting(12)%>" size="4"><font color="red">* 请先到""用户->用户组管理""进行增加查看操作</font><br/>
    共享广告位ID <input name="WapSetting(13)" type="text" value="<%=WapSetting(13)%>" size="4"><font color="red">* 请先到""子系统->广告系统管理""进行增加查看操作</font></td>
    </tr>


    </table>
    </div>
    
    <div class="tab-page" id="template-page">
    <H2 class="tab">模板设置</H2>
	<script type=text/javascript>tabPane1.addTabPage( document.getElementById( "template-page" ) );</script>
    <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>网站首页模板：</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(9)' id='WapSetting9' value="<%=WapSetting(9)%>">&nbsp;<input type='button' name="Submit" class="button" value="选择模板..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("导入模板")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting9').get(0));">

    
    </td>
    </tr>

    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>全站Tags模板：</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(10)' id='WapSetting10' value="<%=WapSetting(10)%>">&nbsp;<input type='button' name="Submit" class="button" value="选择模板..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("导入模板")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting10')[0]);"></td>
    </tr>
    
    <tr style="display:none" valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>专题首页模板：</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(11)' id='WapSetting11' value="<%=WapSetting(11)%>">&nbsp;<input type='button' name="Submit" class="button" value="选择模板..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("导入模板")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting11')[0]);"></td>
    </tr>  


    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>小论坛首页模板：</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(31)' id='WapSetting31' value="<%=WapSetting(31)%>">&nbsp;<input type='button' name="Submit" class="button" value="选择模板..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("导入模板")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting31')[0]);"></td>
    </tr>  


    <tr valign="middle" style="display:none" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>小论坛发帖模板：</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(32)' id='WapSetting32' value="<%=WapSetting(32)%>">&nbsp;<input type='button' name="Submit" class="button" value="选择模板..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("导入模板")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting32')[0]);"></td>
    </tr>  

    <tr><td colspan=2 height='1' bgcolor='green'></td></tr>
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>会员首页模板：</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(19)' id='WapSetting19' value="<%=WapSetting(19)%>">&nbsp;<input type='button' name="Submit" class="button" value="选择模板..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("导入模板")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting19')[0]);"></td>
    </tr> 
    <tr><td colspan=2 height='1' bgcolor='green'></td></tr>
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>空间首页模板：</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(29)' id='TemplateID' value="<%=WapSetting(29)%>">&nbsp;<input type='button' name="Submit" class="button" value="选择模板..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("导入模板")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('WapSetting(29)'));"></td>
    </tr> 
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>空间副模板：</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(30)' id='TemplateID' value="<%=WapSetting(30)%>">&nbsp;<input type='button' name="Submit" class="button" value="选择模板..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("导入模板")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('WapSetting(30)'));"></td>
    </tr> 
    
    <tr><td colspan=2 height='1' bgcolor='green'></td></tr>

    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>个人空间主模板：</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(20)' id='WapSetting20' value="<%=WapSetting(20)%>">&nbsp;<input type='button' name="Submit" class="button" value="选择模板..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("导入模板")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting20')[0]);"></td>
    </tr>
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>个人空间首页副模板：</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(21)' id='WapSetting21' value="<%=WapSetting(21)%>">&nbsp;<input type='button' name="Submit" class="button" value="选择模板..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("导入模板")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting21')[0]);"></td>
    </tr>
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>个人空间小档案副模板：</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(22)' id='WapSetting22' value="<%=WapSetting(22)%>">&nbsp;<input type='button' name="Submit" class="button" value="选择模板..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("导入模板")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting22')[0]);"></td>
    </tr>  

    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>个人空间日志副模板：</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(28)' id='WapSetting28' value="<%=WapSetting(28)%>">&nbsp;<input type='button' name="Submit" class="button" value="选择模板..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("导入模板")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting28')[0]);"></td>
    </tr>  

    <tr><td colspan=2 height='1' bgcolor='green'></td></tr>

    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>企业空间主模板：</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(23)' id='WapSetting23' value="<%=WapSetting(23)%>">&nbsp;<input type='button' name="Submit" class="button" value="选择模板..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("导入模板")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting23')[0]);"></td>
    </tr>
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>企业空间首页副模板：</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(24)' id='WapSetting24' value="<%=WapSetting(24)%>">&nbsp;<input type='button' name="Submit" class="button" value="选择模板..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("导入模板")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting24')[0]);"></td>
    </tr>
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>企业空间小档案副模板：</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(25)' id='WapSetting25' value="<%=WapSetting(25)%>">&nbsp;<input type='button' name="Submit" class="button" value="选择模板..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("导入模板")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting25')[0]);"></td>
    </tr>  
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>企业空间日志副模板：</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(27)' id='WapSetting27' value="<%=WapSetting(27)%>">&nbsp;<input type='button' name="Submit" class="button" value="选择模板..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("导入模板")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting27')[0]);"></td>
    </tr>  
    
    <tr><td colspan=2 height='1' bgcolor='green'></td></tr>
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>圈子主模板：</strong></div></td>
    <td width="63%" height="30"><input type="text" size='40' name='WapSetting(26)' id='WapSetting26' value="<%=WapSetting(26)%>">&nbsp;<input type='button' name="Submit" class="button" value="选择模板..." onClick="OpenThenSetValue('../../<%=KS.Setting(89)%>KS.Frame.asp?URL=KS.Template.asp&Action=SelectTemplate&PageTitle=<%=Server.URLEncode("导入模板")%>&CurrPath=<%=Server.Urlencode(CurrPath)%>',450,350,window,$('#WapSetting26')[0]);"></td>
    </tr>  
    </table>
    </div>
    
    <div class="tab-page" id="cardonline-page">
    <H2 class="tab">接口设置</H2>
	<script type=text/javascript>tabPane1.addTabPage( document.getElementById( "cardonline-page" ) );</script>
    <table width="100%" border="0" align="center" cellpadding="1" cellspacing="1" class="ctable">

    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>平台名称：</strong></div></td>
    <td width="63%" height="30"><Input name="WapSetting(15)" value="<%=WapSetting(15)%>"></td>
    </tr>
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>备注说明：</strong></div></td>
    <td width="63%" height="30"><textarea name="WapSetting(16)" cols="40" rows="4"><%=WapSetting(16)%></textarea></td>
    </tr>

    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>支付编号：</strong></div>请填入您扣费平台申请的商户编号</td>
    <td width="63%" height="30"><Input name="WapSetting(17)" value="<%=WapSetting(17)%>">&nbsp;<a href="http://www.spvnow.com/zhuce.asp" target="_blank"><strong><font color="red">申请帐号</font></strong></a></td>
    </tr>

    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>支付密钥：</strong></div>请填入您在上述扣费平台中设置的MD5私钥</td>
    <td width="63%" height="30"><Input name="WapSetting(18)" value="<%=WapSetting(18)%>"></td>
    </tr>
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>是否启用:</strong></div></td>
    <td width="63%" height="30"><input type="radio" value="0" name="WapSetting(14)">禁用
						<input type="radio" value="1" name="WapSetting(14)" checked>启用</td>
    </tr>
    
    <tr valign="middle" class="tdbg" onMouseOver="this.className='tdbgmouseover'" onMouseOut="this.className='tdbg'">
    <td width="32%" height="30" class="CleftTitle" align="right"><div><strong>同步地址:</strong></div>将地址复制到扣费平台中设置同步地处</td>
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
		   { alert('请选择WAP网站首页模板!');
		   $('#WapSetting9').focus();
		   return false;
		   }
		if ($('#WapSetting10').val()=='')
		   { alert('请选择WAP友情链接页模板!');
		   $('#WapSetting10').focus();
		   return false;
			}
		if ($('#WapSetting11').val()=='')
		   { alert('请选择WAP留言板首页模板!');
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
