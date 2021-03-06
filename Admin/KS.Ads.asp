<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<!--#include file="../Conn.asp"-->
<!--#include file="../KS_Cls/Kesion.CommonCls.asp"-->
<!--#include file="Include/Session.asp"-->
<%
'****************************************************
' Software name:Kesion CMS 6.5
' Email: service@kesion.com . QQ:111394,9537636
' Web: http://www.kesion.com http://www.kesion.cn
' Copyright (C) Kesion Network All Rights Reserved.
'****************************************************
Dim KSCls
Set KSCls = New Convention_Main
KSCls.Kesion()
Set KSCls = Nothing

Class Convention_Main
        Private KS
		Private TotalPage,MaxPerPage,adssql,RSObj,totalPut,CurrentPage,TotalPages,i,advlistact,px,adsrs
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
       Sub Kesion()
	   	    If Not KS.ReturnPowerResult(0, "KSMS20006") Then
			  Call KS.ReturnErr(1, "")
			  exit sub
			End If

	    Select Case KS.G("Action")
		 Case "Adw"
		   Call AdsAdw()
		 Case "Addads"
		   Call AdsAddads()
		 Case "Help"
		   Call AdsHelp()
		 Case "Adslist"
		   Call Adslist()
		 Case "Listip"
		   Call AdsListip()
		 Case "IPDel"
		   Call IPDel()
		 Case "Manage"
		   Call AdsManage()
		 Case Else
		  Call AdsMain()
		End Select
	   End Sub
	   Sub AdsMain()
         With Response
		 
		    .Write "<html>"
			.Write"<head>"
			.Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			.Write"<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.Write"</head>"
			.Write"<body scroll=no leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
            .Write "<table width='100%' height='100%' border='0' cellspacing='0' cellpadding='0'><tr><td height='25'>"
		    .Write "<ul id='menu_top'>"
			.Write "<li class='parent' onclick=""Ads.location.href='KS.Ads.asp?Action=Adw'""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/unverify.gif' border='0' align='absmiddle'>广告位</span></li>"
			.Write "<li class='parent' onclick=""Ads.location.href='KS.Ads.asp?Action=Addads'""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/a.gif' border='0' align='absmiddle'>增加广告</span></li>"
			.Write "<li class='parent' onclick=""Ads.location.href='KS.Ads.asp?Action=Help'""><span class=child onmouseover=""this.parentNode.className='parent_border'"" onMouseOut=""this.parentNode.className='parent'""><img src='images/ico/s.gif' border='0' align='absmiddle'>查看说明</span></li><li></li>"
			.Write "<div>&nbsp;查看选项："
			.Write "<input onclick=""Ads.location.href='?Action=Adslist'"" name=""Option1"" title=""查看正常广告"" type=""radio"">正常广告"
			.Write "<input onclick=""Ads.location.href='?type=img&Action=Adslist'"" name=""Option1"" title=""查看所有图片广告"" type=""radio"">图片广告"
            .Write "<input onclick=""Ads.location.href='?type=txt&Action=Adslist'"" name=""Option1"" title=""查看所有文本广告"" type=""radio"">文本广告"	
            .Write "<input onclick=""Ads.location.href='?type=click&Action=Adslist'"" name=""Option1"" title=""按点击排行查看所有广告"" type=""radio"">点击排行"	
            .Write "<input onclick=""Ads.location.href='?type=show&Action=Adslist'"" name=""Option1"" title=""按显示排行查看所有广告"" type=""radio"">显示排行"	
            .Write "<input onclick=""Ads.location.href='?type=close&Action=Adslist'"" name=""Option1"" title=""查看所有暂停的广告"" type=""radio"">暂停广告"	
            .Write "<input onclick=""Ads.location.href='?type=lose&Action=Adslist'"" name=""Option1"" title=""看所有失效的广告"" type=""radio"">失效广告"	
			.write "</ul>"
			.write "</tr><tr><td>"
			.Write " <iframe name=""Ads"" scrolling=""auto"" frameborder=""0"" src=""KS.Ads.asp?Action=Adw"" width=""100%"" height=""100%""></iframe>"
            .Write " </td></tr></table>"
		End With
  End Sub
  
  '查看帮助
  Sub AdsHelp()
  	    With Response
		 .Write "<html>"
		 .Write"<head>"
		 .Write"<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
		 .Write"<link href=""Include/admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
		 .Write"</head>"
		 .Write"<body leftmargin=""0"" topmargin=""0"" marginwidth=""0"" marginheight=""0"">"
      End With %>
		<br>
		<div align="center">
		  <center>
		  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="95%" id="AutoNumber1">
			<tr>
			  <td width="100%"><b>一、系统特点：</b><ol>
				<li>通过本系统可以设置并管理无数个广告位</li>
				<li>各广告位中可添加无数个循环播放的广告条</li>
				<li>
				广告位中的广告条已有7种显示方式,即&quot;页内嵌入循环&quot;、&quot;上下排列置入&quot;、&quot;左右排列置入&quot;、&quot;向上滚动置入&quot;、&quot;向左滚动置入&quot;、&quot;弹出多个窗口&quot;、&quot;循环弹出窗口&quot;，具体说明请参阅 
				<a href="addadw.asp#说明">广告位栏目中广告位显示方式说明</a></li>
				<li>广告条可以是GIF、SWF（Flash）或纯文本三种显示类型</li>
				<li>广告位上的广告条为循环播放，每次显示的是该广告位中等待时间最长、且处于正常状态的广告条</li>
				<li>可对任意广告条，随时执行暂停、激活、修改、删除等操作</li>
				<li>删除某一条广告时，与其相关的显示、点击记录也将随之删除</li>
				<li>轻松实现广告位的页面发布,具体参阅《<a href="#三">广告位发布说明</a>》</li>
				<li>多种广告播放条件控制广告播放状态，可设点击最高限制、显示最高限制、最后时间限制等</li>
				<li>完善的广告访问记录，可显示广告浏览者、点击者的IP地址</li>
				<li>当有大量广告条存在时，可通过多种条件查询广告条以对其进行操作</li>
			  </ol>
			  <p><b>二、使用说明：</b></p>
			  <ol>
				<li>在 <font color="#FF0000">广 告 位</font> 一栏内可添加新广告位或修改、删除已有广告位标识，查询广告位ID</li>
				<li>在 <font color="#FF0000">添加广告 </font>一栏内可为某广告位添加一个新广告条</li>
				<li>在 <font color="#FF0000">正常广告 </font>
				一栏内显示当前所有处于正常播放状态的广告条，并可执行修改、删除、暂停、预览操作</li>
				<li>在 <font color="#FF0000">图片广告 </font>
				一栏内显示当前所有处于正常播放状态的非文本广告条，并可执行修改、删除、暂停、预览操作</li>
				<li>在 <font color="#FF0000">文本广告 </font>
				一栏内显示当前所有处于正常播放状态的纯文本广告条，并可执行修改、删除、暂停、预览操作</li>
				<li>在 <font color="#FF0000">点击排行 </font>内 
				按点击次数的不同顺序显示各广告条的点击次数，并可执行修改、删除、暂停、激活、预览操作</li>
				<li>在 <font color="#FF0000">显示排行 </font>内 
				按显示次数的不同顺序显示各广告条的显示次数，并可执行修改、删除、暂停、激活、预览操作</li>
				<li>在 <font color="#FF0000">暂停列表 </font>内 
				显示当前所有处于暂停播放状态的广告条，并可执行修改、删除、激活、预览操作</li>
				<li>在 <font color="#FF0000">失效列表 </font>内 
				显示当前所有已经失效的广告条，并可执行修改、删除、激活、预览操作</li>
				<li>在 <font color="#FF0000">广 告 位 </font>内 
				通过某广告位连接，可显示该广告位下的所有广告条，并可执行修改、删除、暂停、预览操作</li>
			  </ol>
			  <p><b><a name="三">三</a>、广告位发布说明：</b></p>
			  <ol>
				<li>确定 <font color="#FF0000">实际页面中的预定广告位置</font> 应放置哪个 
				<font color="#FF0000">通过本系统设置的广告位</font> </li><br><br>
				<li>通过 <font color="#FF0000">广 告 位</font> 一栏，得到所需 <font color="#FF0000">
				广告位ID</font></li><br><br>
				<li>然后将下表的内容拷贝到预定广告位置，注意将其中的 <font color="#FF0000">广告位ID</font> 对应正确</li><br><br>
			   
		
				<input type="text" name="T1" size="100" value="<script language=javascript src=<%KS.Setting(2)%>ad.asp?i=广告位ID></script>"></li>
			  </ol>
		
			  <p><b>四、注意事项：</b></p>
			  <ol>
				<li>每个广告位中的所有广告条显示图片宽度、高度应尽量保持一致，并应注意跟广告位预定的实际页面位置风格一致</li><br><br>
				<li>在实际页面预定的不同广告位中尽量放置使用本系统设置的不同广告位，这样可尽可能多地投放广告</li><br><br>
				<li>同一广告位中,文字广告条与图片广告条尽量不要混合使用</li>
			  </ol>
			  <p><font color="#FF0000"><b>备注：实际页面中的预定广告位置 </b></font>
			  是指“已有网站页面中要放置广告的位置，用来放置通过本系统设置的广告位”。</p>
			  <p>　</td>
			</tr>
		  </table>
		  </center>
		</div>
<%
  End Sub
  
  '增加广告位
  Sub AdsAdw()
  %>
	<html>
	<link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
	<script language="javascript">
	<!--
	function isok(theform)
	{
		if (theform.placename.value=="")
		{
			alert("请填写广告位标识！");
			theform.placename.focus();
			return (false);
		}
	}
	-->
	</script>
	<body>
  <table border="0" align="center"  width=100%  cellpadding="0" cellspacing="0" height="1">
    <tr> 
           <td valign="top">
        <table border=0 width=98% cellspacing=0 cellpadding=0>
          <tr bgcolor=#ffffff align=center valign=top> 
            <td> 
              <table border=0 width=100% cellspacing=0 cellpadding=2 style="border-collapse: collapse" bordercolor="#111111">
				<form method="POST"  action="?job=add&Action=Adw" onSubmit="return isok(this)">

              <%
		if KS.G("job")="add" then
		 
		set RSObj=server.createobject("adodb.recordset")
		
		if  request("place")="0" then
		SqlStr="select * From KS_ADPlace "
		RSObj.open SqlStr,Conn,1,3
		RSObj.AddNew
		else
		SqlStr="select * From KS_ADPlace where place="&trim(request("place"))
		RSObj.open SqlStr,Conn,1,3
		end if
		RSObj(1) = trim(request("placename"))
		RSObj(2)= trim(request("placelei"))
		RSObj(3)= trim(request("placehei"))
		RSObj(4)= trim(request("placewid"))
		RSObj(5)=trim(request("show_flag"))
		RSObj.update
		RSObj.close
		set RSObj=nothing
		Conn.close
		set Conn=nothing
		Response.Redirect "?Action=Adw"
		
		end if

		if KS.G("job")="del" then
		if  isnumeric(request("place"))=true then
		set RSObj=server.createobject("adodb.recordset")
		SqlStr="select * From KS_ADPlace where place="&trim(request("place"))
		RSObj.open SqlStr,Conn,3,3
		RSObj.delete
		RSObj.close
		set RSObj=nothing
		Conn.Execute("Delete From KS_Advertise Where Place="&KS.ChkClng(request("place")))
		Conn.close
		set Conn=nothing
		Response.Redirect "?Action=Adw"
		end if
		end if
%>
              <tr bgcolor=#ffffff> 
                <td>　</td>
                <td> <input type=hidden name=place value="0" >
                  <p align="center">广告位名称&nbsp;
                              <input type=text name=placename class='textbox' size=20 maxlength=30>
                              <font color="#FF0000">15字以内</font>&nbsp;打开与否
							  <select class='textbox' name="show_flag">
							   <option value="1" selected>打开</option>
							   <option value="0">关闭</option>
				      </select>
							   
                  宽度
                  <input type=text class='textbox' name=placewid value="468" size=3 maxlength=30>
高度
<input class='textbox' type=text name=placehei value="60" size=3 maxlength=30>
                  &#12288;类型 
 <%Call Ggwlei(1)%>&nbsp; 
                  <input class="button" type=submit value=新增广告位 name=B1></p>
                <p align="center">
                  <font color="#808080">*** 请尽量不要使用相同的广告位标识，高度、宽度主要应用于弹出窗口大小、滚动区域设置，请适当设置，不可为空</font></td>
              </tr>
            </form>
          </table>
        </td>
      </tr>
    </table>
  </td>
    </tr>
  </table>
  
  <hr color="#808080" size="1">
<p align="center"> <font color=red><b>已有广告位列表</b></font></p>
<p align=left><font color="#808080">*** 请在高度、宽度中输入合适的<b>数字或百分比或为空自动</b></font></p>
<div align="center">
  <center>
            <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#C0C0C0" width="99%" id="AutoNumber1">
              <tr> 
                <td width="63" align="center" height="23" bgcolor="#F5F5F5"> <font color="#FF0000">广告位ID</font></td>
                <td width="192" height="20" align="center" bgcolor="#F5F5F5">广告位名称</td>
                <td width="62" height="20" align="center" bgcolor="#F5F5F5">宽度</td>
                <td width="58" height="20" align="center" bgcolor="#F5F5F5">高度</td>
                <td width="119" align="center" bgcolor="#F5F5F5">广告位显示方式</td>
                <td width="234" align="center" bgcolor="#F5F5F5">显示与否</td>
                <td width="219" align="center" bgcolor="#F5F5F5">操 作</td>
              </tr>
              <%
	Dim RSObj:Set RSObj=server.createobject("adodb.recordset")
	Dim SqlStr:SqlStr="select * From KS_ADPlace"
	RSObj.open SqlStr,Conn,1,1
	do while not RSObj.eof 
%>
              <form method="POST" action="?job=add&Action=Adw"  onSubmit="return isok(this)">
                <tr height=25> 
                  <td width="63" align="center" nowrap><font color=red><%=RSObj(0)%></font> <input type=hidden name=place value="<%=RSObj(0)%>" >
                  　</td>
                  <td align="center" nowrap> 
                    <input type=text name=placename class='textbox' value="<%=RSObj(1)%>" size=22 maxlength=30>
                  </td>
                  <td width="62" align="center" nowrap> 
                    <input name=placewid type=text class='textbox' id="placewid" value="<%=RSObj(4)%>" size=3 maxlength=30></td>
                  <td width="58" align="center" nowrap><input name=placehei type=text class='textbox' id="placehei" value="<%=RSObj(3)%>" size=3 maxlength=30></td>
                  <td width="119" align="center" nowrap>
                      <%Call Ggwlei(RSObj("placelei"))%>
                  </td>
                  <td width="234" align="center"> 
                    <%if RSObj(5)=1 then%>
                    <input type="radio" name="show_flag" value="1" checked>
                    <font color="#FF0000">打开</font> 
                    <input type="radio" name="show_flag" value="0">
                    关闭 
                    <%else%>
					 <input type="radio" name="show_flag" value="1">
                    打开 
                    <input type="radio" name="show_flag" value="0" checked>
                    <font color="#FF0000">关闭</font> 
                    <%end  if%>
                  </td>
                  <td width="219" align="center" nowrap> 
                    <input class="button" type="submit" value=" 修改" name="B1">
                    <a href="?job=del&Action=Adw&place=<%=RSObj(0)%>" onclick="return(confirm('确定删除该广告位吗?'))">删除</a>&nbsp; <a href=?Action=Adslist&type=place&place=<%=RSObj(0)%>>已有广告条</a> 
                  <a href=KS.Ads.asp?Action=Manage&id=<%=RSObj(0)%>&job=yulanggw>预览</a></td>
                </tr>
              </form>
              <%RSObj.movenext
      loop
      RSObj.close:set RSObj=nothing
	  Conn.close:set Conn=nothing
      %>
            </table>
  </center>
</div>
  <p align="left">
  <p align="left"><hr color="#808080" size="1">
<p align="left"><font color="#FF0000"><a name="说明">广告位显示方式说明</a>：</font></p>
<center>
  </p>
  <ul>
    <li><p align="left">
    页内嵌入循环：就是将广告位直接置入某页面一固定位置，并在同一位置循环显示广告位中的所有正常广告条，这样，每刷新一次就会更替显示一个新的广告条</p>
    </li>
    <li><p align="left">上下排列置入：从上到下竖排广告位中的所有正常广告条</p>
    </li>
    <li><p align="left">左右排列置入：从左到右横排广告位中的所有正常广告条</p>
    </li>
    <li><p align="left">向上滚动置入：向上滚动显示广告位中的所有正常广告条</p>
    </li>
    <li><p align="left">向左滚动置入：向左滚动显示广告位中的所有正常广告条</p>
    </li>
    <li><p align="left">弹出多个窗口：页面打开时同时弹出多个窗口，每个窗口内显示一个广告条，弹出数量跟该广告位中的正常广告条数一致</p>
    </li>
    <li><p align="left">
    循环弹出窗口：页面打开时同时弹出一个窗口，在同一窗口内循环显示广告位中的正常广告，这样，每刷新一次就会在弹出窗口中更替显示一个新的广告条</p>
    </li>
  </ul>
  <p align="left"><font color=red> 广告插入方法：</font>
  <div align=left>
  <li><font color="#FF0000">方法1、</font>在模板编辑器中插入相应的广告位标签；
  <li><font color="#FF0000">方法2、</font>将下表内容放到预定广告位置，并将其中的<font color="#FF0000">广告位ID</font>对应正确 
   <font color="#808080">请在广告位列表中查看</font><font color="#FF0000">广告位ID</font>
  </div>
  <input type="text" name="T1" size="100" value='<script type="text/javascript" src="<%=KS.GetDomain%>plus/ShowA.asp?i=广告位ID"></script>'>
</p>
</body>
</html>
<%End Sub
'调用常用广告位类型下拉菜单
Sub Ggwlei(shu) '用于表示类型的数
%>
 <select size=1 name=placelei>
                    <option value=1 <% if shu=1 then%>selected<%end if%>>页内嵌入循环</option>
                    <option value=2 <% if shu=2 then%>selected<%end if%>>上下排列置入</option>
                    <option value=3 <% if shu=3 then%>selected<%end if%>>左右排列置入</option>
                    <option value=4 <% if shu=4 then%>selected<%end if%>>向上滚动置入</option>
                    <option value=5 <% if shu=5 then%>selected<%end if%>>向左滚动置入</option>
                    <option value=6 <% if shu=6 then%>selected<%end if%>>弹出多个窗口</option>
                    <option value=7 <% if shu=7 then%>selected<%end if%>>循环弹出窗口</option>
</select>
<%
  End Sub
  
  '增加广告
Sub AdsAddads()
        Dim CurrPath:CurrPath = KS.GetCommonUpFilesDir()
%>
<html>
<link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
<script src="../KS_Inc/common.js" language="javascript"></script>
<script language="javascript">
<!--
function isok(theform)
{
    if (theform.name.value=="")
    {
        alert("请填写广告名称！");
        theform.name.focus();
        return (false);
    }
    if (theform.url.value=="")
    {
        alert("请填写链接URL！");
        theform.url.focus();
        return (false);
    }
    return (true);
}
-->
</script>
<%
Dim Ggw,sitename,url,intro,xslei,gif_url,wid,hei,window,classs,clicks,shows,lasttime,flag,AdorderID
Ggw=1:URL="http://":xslei="gif":gif_url="http://":wid="":hei="":window=0:classs=0:flag="Add":AdorderID=1
if KS.G("job")="add" then
	Call  addrk()
ElseIf KS.G("job")="edit" then
 Dim KS_RS_Obj:Set KS_RS_Obj=Server.CreateObject("Adodb.Recordset")
 KS_RS_Obj.Open "Select * From KS_Advertise where id="&KS.ChkClng(KS.G("id")),Conn,1,1
  If Not KS_RS_Obj.Eof Then
  Ggw      = KS_RS_Obj("Place")
  sitename = KS_RS_Obj("sitename")
  url      = KS_RS_Obj("url")
  intro    = KS_RS_Obj("intro")
  xslei    = KS_RS_Obj("xslei")
  gif_url  = KS_RS_Obj("gif_url")
  wid      = KS_RS_Obj("wid")
  Hei      = KS_RS_Obj("Hei")
  window   = KS_RS_Obj("window")
  classs   = KS_RS_Obj("class")
  clicks   = KS_RS_Obj("clicks")
  shows    = KS_RS_Obj("Shows")
  lasttime = KS_RS_Obj("lasttime")
  AdorderID = KS_RS_Obj("AdorderID")
  End If
  KS_RS_Obj.Close:Set KS_RS_Obj=Nothing
  flag="Edit"
end if
%>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
}
.STYLE3 {color: #3300FF}
-->
</style>
              <table border=0 width=100% cellspacing=0 cellpadding=0 >
            <tr> 
              <td align=center> 
                <%
if KS.G("job")="suc" then
  if KS.G("flag")="Add" then%>
  <font size="2" color=red><b>添加新广告条成功，请继续添加...</b></font> 
 <%
  else
  %>
    <font size="2" color=red><b>广告条修改成功...</b></font> 

  <%
  end if
elseif KS.G("job")="edit" then
%>
<font size="2" color=red><b>修改广告条</b></font> 
<%else%>
                <font size="2" color=red><b>添加新广告条</b></font> 
                <%
end if
%>
     <hr color="#808080" size="1"> 
	        </td>
            </tr>
          </table>
              <table border=0 width=100% cellspacing=1 cellpadding=2>
				<form method="POST"  name="myform"  action="?flag=<%=Flag%>&job=add&Action=Addads&id=<%=KS.G("id")%>" onSubmit="return isok(this)">
				 <input type="hidden" value="<%=request.ServerVariables("http_referer")%>" name="comeurl">
              <tr bgcolor=#ffffff> 
                <td>所属广告位</td>
                <td colspan="2"> 
                <%
                Call  Ggwxlxx(Ggw) 
				%>              </td>
              </tr>
			  <tr bgcolor=#ffffff> 
                <td width=85>广告名称</td>
                <td colspan="2"> 
                  <input type="text" class='textbox' name="name" value="<%=sitename%>" size=30 maxlength=30>
                  不超过15个中文或30个字母数字</td>
              </tr>
              <tr bgcolor=#ffffff> 
                <td width=85>链接URL</td>
                <td colspan="2"> 
                  <input type=text class='textbox' name=url size=40 value="<%=url%>">
			    </td>
              </tr>
               <tr bgcolor=#ffffff> 
                <td width=85>简介/内容</td>
                <td width="200"> 
                  <textarea rows="5" class='textbox' name="intro" cols="48" style="height:60"><%=intro%></textarea></td>
                <td> <font color="#FF0000">提示：</font><br>
                  <font color="#808080">如果是嵌入代码请将代码内容填入此处 链接URL无效<br>
                  如果显示纯文本，则显示为广告名称<br>
                  只有GIF图片时URL填写有效</font></font>                  </td>
              </tr>
              <tr bgcolor=#ffffff> 
                <td width=85>广告类型</td>
                <td colspan="2"> 
                  <input name="xslei" type="radio" value="gif" <%if xslei="gif" then response.write " checked"%>>GIF图片 
                  <input type="radio" name="xslei" value="swf" <%if xslei="swf" then response.write " checked"%>><font siz=3 >Flash动画 </font>
                  <input type="radio" name="xslei" value="txt" <%if xslei="txt" then response.write " checked"%>><font siz=3 >纯文本 </font>    
                  <input type="radio" name="xslei" value="dai" <%if xslei="dai" then response.write " checked"%>>嵌入代码                </td>
              </tr>
              <tr bgcolor=#ffffff> 
                <td>&#22270;&#29255;URL</td>
                <td colspan="2"> <input type=text class='textbox' name="gif_url"  size=40 value="<%=gif_url%>">&nbsp;<input type='button' class='button' name='Submit' value='选择地址...' onClick="OpenThenSetValue('Include/SelectPic.asp?Currpath=<%=CurrPath%>',550,290,window,document.myform.gif_url);">
                <font siz=3 > 宽度 </font>
                <input type=text name="wid" value="<%=wid%>" size=3 class='textbox' maxlength="4">
                <font siz=3 >高度 </font> 
                  <input type=text name=hei value="<%=hei%>" size=3 class='textbox'  maxlength="4"><font siz=3 >&nbsp;</font><font color=red siz=3 > 可以是百分比或空默认</font> </td>
              </tr>
              <tr bgcolor=#ffffff> 
                <td>链接&#25171;&#24320;&#26041;&#24335;</td>
                <td colspan="2"> 
                  <select size=1 name=window>
                    <option value=0<%if window=0 then response.write " selected"%>>新窗口打开</option>
                    <option value=1<%if window=1 then response.write " selected"%>>原窗口打开</option>
                  </select>                </td>
              </tr>
              <tr bgcolor=#ffffff> 
                <td>顺序ID</td>
                <td colspan="2"> 
				<input type=text name="AdorderID" value="<%=AdorderID%>" size=10 class='textbox' maxlength="4">&nbsp;(数值小的靠前)
                 </td>
              </tr>
              <tr bgcolor=#c0d0e0> 
                <td bgcolor="#FFFFEE" height="30"><strong>&#25773;&#25918;&#26465;&#20214;组</strong></td>
                <td bgcolor="#FFFFEE" height="30" colspan="2">&#12288;<span class="STYLE3">在八个条件组中，任选其中一组&#65292;用于限制&#35813;&#24191;&#21578;&#33258;&#21160;&#36827;&#20837;&#20241;&#30496;&#29366;&#24577;的条件&#65292;以后&#21487;随时&#20462;&#25913;。</span></td>
              </tr>
              <tr bgcolor=#ffffff> 
                <td align=right><font color=red>（1）</font>
                <input type=radio value=0 name=class<%if classs=0 then response.write " checked"%>></td>
                <td colspan="2"> 
                  <table border=0>
                    <tr> 
                      <td> 
                      　</td>
                      <td>&#26080;&#38480;&#21046;&#24490;&#29615;</td>
                    </tr>
                  </table>                </td>
              </tr>
              <tr> 
                <td align=right><font color=red>（2）</font>
                  <input type=radio value=1 name=class<%if classs=1 then response.write " checked"%>></td>
                <td colspan="2"> 
                  <table border=0>
                    <tr> 
                      <td> 
                      　</td>
                      <td>&#28857;&#20987;&#25968;&#19981;&#36229;&#36807;</td>
                      <td> 
                        <input type=text name=clicks1 class='textbox' size=8<%if classs=1 then response.write " value=""" & clicks &""""%>>                      </td>
                    </tr>
                </table>                </td>
              </tr>
              <tr> 
                <td align=right><font color=red>（3）</font>
                <input type=radio value=2 name=class<%if classs=2 then response.write " checked"%>></td>
                <td bgcolor=#ffffff colspan="2"> 
                  <table border="0">
                    <tr> 
                      <td> 
                      　</td>
                      <td>&#26174;&#31034;&#25968;&#19981;&#36229;&#36807;</td>
                      <td> 
                        <input type=text name=shows2 class='textbox' size=8<%if classs=2 then response.write " value=""" & shows &""""%>>                      </td>
                    </tr>
                  </table>                </td>
              </tr>
              <tr> 
               <td align=right><font color=red>（4）</font>
                 <input type=radio value=3 name=class<%if classs=3 then response.write " checked"%>></td>
                <td colspan="2"> 
                  <table border="0">
                    <tr> 
                      <td> 
                      　</td>
                      <td>&#26174;&#31034;&#25130;&#27490;&#26399;&#20026;</td>
                      <td> 
                        <input type=text name=time3 class='textbox' size=20<%if classs=3 then response.write " value=""" & lasttime &""""%>>
                        &#12288;<font color=#FF0000>&#26684;&#24335;&#20026;yyyy-mm-dd hh:mm:ss&#25110;&#32773;yyyy-mm-dd</font></td>
                    </tr>
                </table>
				</td>
              </tr>
              <tr> 
               <td align=right><font color=red>（5）</font>
                <input type=radio value=4 name=class<%if classs=4 then response.write " checked"%>></td>
                <td bgcolor=#ffffff colspan="2"> 
                  <table border="0">
                    <tr> 
                      <td> 
                      　</td>
                      <td>&#28857;&#20987;&#25968;&#19981;&#36229;&#36807;</td>
                      <td> 
                        <input type=text name=clicks4 class='textbox' size=8<%if classs=4 then response.write " value=""" & chicks &""""%>>                      </td>
                    </tr>
                    <tr> 
                      <td align=center> </td>
                      <td>&#26174;&#31034;&#25968;&#19981;&#36229;&#36807;</td>
                      <td> 
                        <input type=text name=shows4 class='textbox' size=8<%if classs=4 then response.write " value=""" & shows &""""%>>                      </td>
                    </tr>
                  </table>                </td>
              </tr>
              <tr> 
               <td align=right><font color=red>（6）</font>
                <input type=radio value=5 name=class<%if classs=5 then response.write " checked"%>></td>

                <td colspan="2"> 
                  <table border="0">
                    <tr> 
                      <td> 
                      　</td>
                      <td>&#28857;&#20987;&#25968;&#19981;&#36229;&#36807;</td>
                      <td> 
                        <input type=text name=clicks5 class='textbox' size=8<%if classs=5 then response.write " value=""" & chicks &""""%>>                      </td>
                    </tr>
                    <tr> 
                      <td align=center> </td>
                      <td>&#26174;&#31034;&#25130;&#27490;&#26399;&#20026;</td>
                      <td> 
                        <input type=text name=time5 class='textbox' size=20<%if classs=5 then response.write " value=""" & lasttime &""""%>>
                        &#12288;<font color=#FF0000>&#26684;&#24335;&#20026;yyyy-mm-dd hh:mm:ss&#25110;&#32773;yyyy-mm-dd</font></td>
                    </tr>
                </table>                </td>
              </tr>
              <tr> 
                <td align=right><font color=red>（7）</font>
                <input type=radio value=6 name=class<%if classs=6 then response.write " checked"%>></td>

                <td bgcolor=#ffffff colspan="2"> 
                  <table border="0">
                    <tr> 
                      <td> 
                      　</td>
                      <td>&#26174;&#31034;&#25968;&#19981;&#36229;&#36807;</td>
                      <td> 
                        <input type=text name=shows6 class='textbox' size=8<%if classs=6 then response.write " value=""" & shows &""""%>>                      </td>
                    </tr>
                    <tr> 
                      <td></td>
                      <td>&#26174;&#31034;&#25130;&#27490;&#26399;&#20026;</td>
                      <td> 
                        <input type=text name=time6 class='textbox' size=20<%if classs=6 then response.write " value=""" & lasttime &""""%>>
                        &#12288;<font color=#FF0000>&#26684;&#24335;&#20026;yyyy-mm-dd hh:mm:ss&#25110;&#32773;yyyy-mm-dd</font></td>
                    </tr>
                  </table>                </td>
              </tr>
              <tr> 
              <td align=right><font color=red>（8）</font>
                <input type=radio value=7 name=class<%if classs=7 then response.write " checked"%>></td>

                <td colspan="2"> 
                  <table border="0">
                    <tr> 
                      <td> 
                      　</td>
                      <td>&#28857;&#20987;&#25968;&#19981;&#36229;&#36807;</td>
                      <td> 
                        <input type=text name=clicks7 class='textbox' size=8<%if classs=7 then response.write " value=""" & clicks &""""%>>                      </td>
                    </tr>
                    <tr> 
                      <td></td>
                      <td>&#26174;&#31034;&#25968;&#19981;&#36229;&#36807;</td>
                      <td> 
                        <input type=text name=shows7 class='textbox' size=8<%if classs=7 then response.write " value=""" & shows &""""%>>                      </td>
                    </tr>
                    <tr> 
                      <td></td>
                      <td>&#26174;&#31034;&#25130;&#27490;&#26399;&#20026;</td>
                      <td> 
                        <input type=text name=time7 class='textbox' size=20<%if classs=7 then response.write " value=""" & lasttime &""""%>>
                        &#12288;<font color=#FF0000>&#26684;&#24335;&#20026;yyyy-mm-dd hh:mm:ss&#25110;&#32773;yyyy-mm-dd</font></td>
                    </tr>
                </table>                </td>
              </tr>
              <tr bgcolor=#ffffff> 
                <td colspan=3 align=center> 
                  <input type=submit class='button' value=' 提交' name=B1>
                  <input type=reset class='button' value=' 重写' name=B2>                </td>
              </tr>
            </form>
          </table>
 </body>
</html>
<%End Sub
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''广告条信息入库函数（包含修改、添加两种）'''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Sub addrk()
	if KS.G("job")="add" then
	
	dim getname,geturl,getgif,getplace,getwin,getxslei,RSObj,adssql,getclass,getclicks,getshows,gettime,getintro,gethei,getwid,getAdorderID
	
	getname = Trim(Request("name"))
	geturl = Trim(Request("url"))
	getgif = Trim(Request("gif_url"))
	getplace =trim(Request("place"))
	getwin =trim(Request("window"))
	getxslei = trim(Request("xslei"))
	getclass=trim(Request("class"))
	getintro=trim(Request("intro"))
	getwid=trim(Request("wid"))
	gethei=trim(Request("hei"))
	getAdorderID=KS.ChkClng(Request("AdorderID"))
	
	if getxslei="txt" then
	getwid=0
	gethei=0
	end if
	
	if getclass=0 then
	getclicks=0
	getshows=0
	gettime=now()
	
	elseif getclass=1 then
	getclicks=KS.ChkClng(KS.G("clicks1"))
	getshows=0
	gettime=now()
	
	elseif getclass=2 then
	getclicks=0
	getshows=KS.ChkClng(KS.G("shows2"))
	gettime=now()
	
	elseif getclass=3 then
	getclicks=0
	getshows=0
	gettime=trim(request("time3"))
	elseif getclass=4 then
	getclicks=KS.ChkClng(KS.G("clicks4"))
	getshows=KS.ChkClng(KS.G("shows4"))
	gettime=now()
	
	elseif getclass=5 then
	getclicks=KS.ChkClng(KS.G("clicks5"))
	getshows=0
	gettime=trim(request("time5"))
	
	elseif getclass=6 then
	getclicks=0
	getshows=KS.ChkClng(KS.G("shows6"))
	gettime=trim(request("time6"))
	
	elseif getclass=7 then
	getclicks=KS.ChkClng(KS.G("clicks7"))
	getshows=KS.ChkClng(KS.G("shows7"))
	gettime=trim(request("time7"))
	end if
	 if not isdate(gettime) then response.write "<script>alert('显示截止日期，格式有误!');history.back();</script>"
	
	set RSObj=server.createobject("adodb.recordset")
	if  trim(KS.G("id"))="" then '如果是新增广告条
	adssql="select * from KS_Advertise"
	RSObj.open adssql,Conn,1,3
	RSObj.AddNew
	else                                                '如果是修改广告条
	adssql="select * from KS_Advertise where id="&KS.ChkClng(KS.G("id"))
	RSObj.open adssql,Conn,1,3
	end if
	RSObj("act") = 1
	RSObj("sitename") = getname
	RSObj("url") = geturl
	RSObj("gif_url") = getgif
	RSObj("place") = getplace
	RSObj("xslei") = getxslei
	RSObj("hei") = gethei
	RSObj("wid") = getwid
	RSObj("window") = getwin
	RSObj("class") = getclass
	RSObj("clicks") = getclicks
	RSObj("shows") = getshows
	RSObj("lasttime") = gettime
	RSObj("regtime") = Now()
	RSObj("time") = now()
	RSObj("intro")=getintro
	RSObj("AdorderID")=getAdorderID
	RSObj.update
	If KS.G("ID")="" Then
	 RSObj.MoveLast
	 Call KS.FileAssociation(1020,RSObj("ID"),getgif,0)
	Else
	 Call KS.FileAssociation(1020,RSObj("ID"),getgif,1)
	End If
	RSObj.close
	set RSObj=nothing
	Conn.close
	set Conn=nothing
	if KS.g("id")<>"" then
	     %>
		 <script>alert('广告条修改成功!');location.href='<%=KS.g("comeurl")%>';</script>"
		 <%
		 response.end
    else
	Response.Redirect "?job=suc&flag=" & KS.G("Flag")& "&Action=Addads&id="&KS.G("id")
	end if
	end if
	End Sub
	'调出广告位下拉选项
	
	Sub Ggwxlxx(place) 'place 用于判断默认选项
	%>
	  <select size=1 name=place>
	<%
	Dim PRSObj:Set PRSObj=server.createobject("adodb.recordset")
	PRSObj.open "select * From KS_ADPlace",Conn,1,1
	do while not PRSObj.eof
	%>
	<option value="<%=PRSObj(0)%>" <% if PRSObj(0)=place then :Response.Write "selected":end if%>><%=PRSObj(1)%></option>
	 <%PRSObj.movenext
	   loop
	   PRSObj.close
	   Set PRSObj=nothing%>              
	  </select> 
<%
  End Sub
  
  Sub Adslist()
%>
<html>
<link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
body {
	margin-left: 5px;
	margin-top: 2px;
}
-->
</style>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="1"  width="100%" class=tableBorder >
   <form method=post action="?type=search&Action=Adslist">
    <tr>
      <td width="100%" align="right">快速搜索=&gt;&gt;
      <select size="1" name="adorder" >
<option value="id">广告ID</option>
<option value="name">名称关键字</option>
</select> <input type="text" name="nr" size="20">
<input type="submit" value="查 询" name="B1" class=button></td>
    </tr></form>
  </table>
  </center>
</div>
          <table border=0 width=100% cellspacing=3 cellpadding=3>
            <tr> 
              <td align=center> 
                <%
                  if request("px")="" then
                  px="desc"
                  else
                  px=""
                  end if
                  
                   Select Case KS.G("type")
                   
                          Case "img"
                           adssql="select * from KS_Advertise where act=1 and (xslei='gif' or xslei='swf') order by regtime "&px
                %>
                <b>正常播放的图片类广告条列表</b>  <a href=?Action=Adslist&type=<%=KS.G("type")%>&px=x>升</a>  <a href=?Action=Adslist&type=<%=KS.G("type")%>>降</a>
               
			    <%        Case "txt"
                           adssql="select * from KS_Advertise where act=1 and xslei='txt' order by regtime "&px
                %>
                <b>正常播放的纯文本广告条列表</b>  <a href=?Action=Adslist&type=<%=KS.G("type")%>&px=x>升</a>  <a href=?Action=Adslist&type=<%=KS.G("type")%>>降</a>
                <%
                          Case "close"
                           adssql="select * from KS_Advertise where act=0 order by regtime "&px

                %>
                <b>处于暂停而未失效的广告条列表</b>  <a href=?Action=Adslist&type=<%=KS.G("type")%>&px=x>升</a>  <a href=?Action=Adslist&type=<%=KS.G("type")%>>降</a>
                <%
                          Case "lose"
                           adssql="select * from KS_Advertise where act=2 order by regtime "&px
                %>
                <b>已经失效的的广告条列表</b> <a href=?Action=Adslist&type=<%=KS.G("type")%>&px=x>升</a>  <a href=?Action=Adslist&type=<%=KS.G("type")%>>降</a> 
                <%
                          Case "click"
                           adssql="select * from KS_Advertise where act<>2 order by click "&px
                %>
                <b>按点击次数<%if px="desc" then: Response.Write "降序":else:Response.Write "升序":end if%>排列未失效广告条</b>  <a href=?Action=Adslist&type=<%=KS.G("type")%>&px=x>升</a>  <a href=?Action=Adslist&type=<%=KS.G("type")%>>降</a>
               <%
                          Case "show"
                           adssql="select * from KS_Advertise where act<>2 order by show "&px
                %>
                <b>按显示次数<%if px="desc" then: Response.Write "降序":else:Response.Write "升序":end if%>排列未失效广告条</b> <a href=?Action=Adslist&type=<%=KS.G("type")%>&px=x>升</a>  <a href=?Action=Adslist&type=<%=KS.G("type")%>>降</a>
               <%
                          Case "place"
                          
                          if isnumeric(request("place"))=true then
                           adssql="select * from KS_Advertise where act=1 and place="&trim(request("place"))&" order by regtime "&px
						 
		%>
                <b>ID为<%=request("place")%>的广告位中正常播放的广告条</b>  <a href=?Action=Adslist&type=<%=KS.G("type")%>&place=<%=request("place")%>&px=x>升</a>  <a href=?Action=Adslist&type=<%=KS.G("type")%>&place=<%=request("place")%>>降</a>
				 
                <%else
                  adssql="select * from KS_Advertise where act=1 order by regtime "&px
                %>
                <b>所有正常播放的广告条列表</b>  <a href=?Action=Adslist&type=<%=KS.G("type")%>&px=x>升</a>  <a href=?Action=Adslist&type=<%=KS.G("type")%>>降</a>
                        
                <%end if%>
               <%
                          Case "search"
                          if request("adorder")="id" and isnumeric(request("nr"))=true then
                           adssql="select * from KS_Advertise where id="&trim(request("nr"))
                          
                %>
                <b>查询 ID为<%=request("nr")%> 的广告条信息</b>
                <%        else
                  adssql="select * from KS_Advertise where sitename like '%"&request("nr")&"%' order by regtime "&px
                %>
                <b>查询名称含有关键字“<%=request("nr")%>”广告条</b> <a href=?Action=Adslist&type=<%=KS.G("type")%>&px=x>升</a>  <a href=?Action=Adslist&type=<%=KS.G("type")%>>降</a>
                        
                <%end if%>

                <%       
                          Case else
                          adssql="select * from KS_Advertise where act=1 order by regtime "&px
                %>
                <b>所有正常播放的广告条列表</b>  <a href=?Action=Adslist&type=<%=KS.G("type")%>&px=x>升</a>  <a href=?Action=Adslist&type=<%=KS.G("type")%>>降</a>
                <%
                    end Select
                %>
              </td>
            </tr>
          </table>
		   </body>
</html>
<%

if isnumeric(request("page")) then
currentPage=cint(request("page"))
else
currentPage=1
end if
set RSObj=server.createobject("adodb.recordset")

RSObj.open adssql,Conn,1,1
if RSObj.eof and RSObj.bof then
Response.Write "<tr><td bgcolor=#ffffff align=center><BR><BR>没有任何相关记录<BR><BR><BR><BR>"
else
RSObj.pagesize=10  '每页显示的记录数
totalPut=RSObj.recordcount '记录总数
totalPage=RSObj.pagecount
MaxPerPage=RSObj.pagesize
if currentpage<1 then
currentpage=1
end if
if currentpage>totalPage then
currentpage=totalPage
end if
if currentPage=1 then
showContent
showpages
else
if (currentPage-1)*MaxPerPage<totalPut then
RSObj.move  (currentPage-1)*MaxPerPage
dim bookmark
bookmark=RSObj.bookmark '移动到开始显示的记录位置
showContent
showpages
end if
end if
RSObj.close:set RSObj=nothing
end if
Conn.close:set Conn=nothing
End Sub

sub showContent
i=0
do while not (RSObj.eof or err)
%>
  <div align="center">
    <center>
    <table border="0" cellpadding="0" cellspacing="1"  width="100%" class=tableBorder >
		  <input type="hidden" name="id" value="<%=RSObj("id")%>">
     <tr>
        <td width="175" class="forumRowHighlight"><font color="#FF0000">&nbsp;广告ID：<%=RSObj("id")%> </font></td>
        <td width="370" class=forumRow>&nbsp;名称：<%=RSObj("sitename")%></td>
        <td class=forumRow width="275">
       &nbsp;URL： 
       <%=RSObj("url")%></td>
        <td  width="105" align="center" class="forumRowHighlight">
        <%if RSObj("xslei")="txt" then%>
           <a href=KS.Ads.asp?Action=Manage&id=<%=RSObj("id")%>&job=yulan>预览广告</a>
        <%else
        
        %>
            <a href=KS.Ads.asp?Action=Manage&id=<%=RSObj("id")%>&job=yulan>预览广告</a>
       <%end if%>
　</td>
      </tr>
      <tr>
        <td width="175" height="60" class="forumRowHighlight">&nbsp;打开：<%= Ggdklx(RSObj("window"))%><br>&nbsp;显示：<%= Ggxslx(RSObj("xslei"))%><br>
        &nbsp;类型：<%= Ggwlx(RSObj("place"))%></td>
        <td height="60" class="forumRowHighlight">&nbsp;加入时间：<font color=red><%=RSObj("regtime")%></font><br>&nbsp;最新显示：<font color=red><%=RSObj("time")%></font><br>
        &nbsp;最新点击：<font color=red><%=RSObj("lasttime")%></font></td>
        <td height="60" width="272"class="forumRowHighlight" >&nbsp;显示次数：<%call  Xscs()%><br>&nbsp;点击次数：<%call  Djcs()%><br>
        &nbsp;广 告 位：<%= Ggwm(RSObj("place"))%>  ID=<font color=red><%=RSObj("place")%></font></td>
        <td height="60" width="104" align="center" class="forumRowHighlight">              <%
if RSObj("act")=1 then
%>                <a href=?Action=Addads&job=edit&id=<%=RSObj("id")%>>修改</a>
              <a href=?Action=Manage&id=<%=RSObj("id")%>&job=close>暂停</a> 
              <%
else
%>
              <a href=?Action=Manage&id=<%=RSObj("id")%>&job=open>激活</a> 
              <%end if%><a href=?Action=Manage&id=<%=RSObj("id")%>&job=delit>删除</a></td>
      </tr>
      <tr>
        <td colspan="3" height="20">&nbsp;失效条件：<%call  Sxtj()%></td>
                <td height="20" width="104" align="center"></td>
      </tr>
      </table>
    </center>
</div>
  <%
i=i+1
if i>=MaxPerPage then exit do '循环时如果到尾部则先退出，如果记录达到页最大显示数，也退出
RSObj.movenext
loop
end sub 

sub Showpages()
%>
    
        <table border=0 width=100% cellpadding=2>
            <tr bgcolor=#ffffff> 
              <td align=right colspan=4>
			   <%'显示分页信息
			  Call KS.ShowPageParamter(totalPut, MaxPerPage, "KS.Ads.asp", True, "条", CurrentPage, KS.QueryParam("page"))
			  %>
              </td>
            </tr>
        </table>
     
<%
end sub

Sub Sxtj()

if RSObj("class")=1 then
%>
              点击<font color=red><%=RSObj("clicks")%></font>次 
              <%
elseif RSObj("class")=2 then
%>
              显示<font color=red><%=RSObj("shows")%></font>次 
              <%
elseif RSObj("class")=3 then
%>
              截止期<font color=red><%=RSObj("lasttime")%></font> 
              <%
elseif RSObj("class")=4 then
%>
              点击<font color=red><%=RSObj("clicks")%></font>次，显示<font color=red><%=RSObj("shows")%></font>次 
              <%
elseif RSObj("class")=5 then
%>
              点击<font color=red><%=RSObj("clicks")%></font>次，截止期<font color=red><%=RSObj("lasttime")%></font> 
              <%
elseif RSObj("class")=6 then
%>
              显示<font color=red><%=RSObj("shows")%></font>次，截止期<font color=red><%=RSObj("lasttime")%></font> 
              <%
elseif RSObj("class")=7 then
%>
              点击<font color=red><%=RSObj("clicks")%></font>次，显示<font color=red><%=RSObj("shows")%></font>次，截止期<font color=red><%=RSObj("lasttime")%></font> 
              <%
else
%>
              无限制条件 
<%
end if
%>
<%end sub

Sub Xscs()%>
 <font color=red><%=RSObj("show")%></font> (<a href=?Action=Listip&id=<%=RSObj("id")%>&ip=sip>显示记录</a>)
<%end sub

Sub Djcs()%>
 <font color=red><%=RSObj("click")%></font> (<a href=?Action=Listip&id=<%=RSObj("id")%>&ip=cip>点击记录</a>)
<%end sub
	'广告显示类型名
	Function Ggxslx(lx)
	Select Case lx
		  Case "txt":Ggxslx="纯文本"
		  Case "gif":Ggxslx="GIF图片"
		  Case "swf":Ggxslx="Flash动画"
		  Case "dai":Ggxslx="嵌入代码"
	End select
	End Function
	'广告打开类型名
	Function Ggdklx(lx)
	Select Case lx
		  Case 0:Ggdklx="新窗口"
		  Case else:Ggdklx="本窗口"
	End select
	End Function
	'广告位类型标示数字调用
	Function Ggwlxsz(place1)
	set PRSObj=server.createobject("adodb.recordset")
	PRSObj.open "select * From KS_ADPlace where place="&place1,Conn,1,1
	if not PRSObj.eof then
	Ggwlxsz=PRSObj(2)
	else
	Ggwlxsz=0
	end if
	PRSObj.close
	Set PRSObj=nothing
	End Function
	'广告位类型名称调用
	Function Ggwlx(place)
	Dim  PRSObj:Set PRSObj=server.createobject("adodb.recordset")
	PRSObj.open "select * From KS_ADPlace where place="&place,Conn,1,1
	if not PRSObj.eof then
	Ggwlx=PRSObj(2)
	Select Case Ggwlx
		   Case 1:Ggwlx="页内嵌入循环"
		   Case 2:Ggwlx="上下排列置入"
		   Case 3:Ggwlx="左右排列置入"
		   Case 4:Ggwlx="向上滚动置入"
		   Case 5:Ggwlx="向左滚动置入"
		   Case 6:Ggwlx="弹出多个窗口"
		   Case 7:Ggwlx="循环弹出窗口"
	End select
	else
	Ggwlx="广告位被删除"
	end if
	PRSObj.close
	Set PRSObj=nothing
	
	End Function
	'广告位名称调用
	Function Ggwm(place)
	Dim  PRSObj:Set PRSObj=server.createobject("adodb.recordset")
	PRSObj.open "select * From KS_ADPlace where place="&place,Conn,1,1
	if not PRSObj.eof then
	Ggwm=PRSObj(1)
	else
	Ggwm=""
	end if
	PRSObj.close:Set PRSObj=nothing
	End Function
	
	'显示IP
	Sub AdsListIP()
	    Dim getadid
	   %>
	    <html>
		<link href="Include/Admin_Style.CSS" rel="stylesheet" type="text/css">
		<table border=0 width=100% cellspacing=0 cellpadding=0>
		<tr>
		<td >
		<table border="0" align=center cellpadding="5" cellspacing="1" class=tableBorder width="100%">
		<tr><td align=center class="forumRowHighlight">
		<%
		if KS.G("ip")="sip" then
		%>
		ID为 <%=KS.G("id")%> 的广告条显示记录
		<%
		elseif KS.G("ip")="cip" then
		%>
		ID为 <%=KS.G("id")%> 的广告条点击记录
		<%
		end if
		%>
		</td>
		<td class="forumRowHighlight" align="right"><input class="button" type="button" name="button1" value="清除所有的IP记录" onClick="if (confirm('此操作不可逆,确定删除所有记录吗？')){location.href='?action=IPDel&AdID=<%=KS.G("ID")%>&ip=<%=KS.G("ip")%>';}"></td>
		</tr></table>
		
		
		<table border="0" align=center cellpadding="5" cellspacing="1" class=tableBorder width="100%">
		<tr bgcolor=#ffffff><td align="center" class="forumRowHighlight" height="20">
		记录ID
		</td><td align=center class="forumRowHighlight" height="20">IP 地址</td>
		  <td align=center class="forumRowHighlight" height="20">时　间</td></tr>
		<%
		if not isempty(request("page")) then
		 currentPage=cint(request("page"))
		else
		 currentPage=1
		end if
		set adsrs=server.createobject("adodb.recordset")
		
		if KS.G("ip")="sip" then
		getadid=cint(request("id"))
		adssql="select * From KS_Adiplist where adid="&getadid&" and class=1 order by id desc"
		
		elseif KS.G("ip")="cip" then
		getadid=cint(request("id"))
		adssql="select * From KS_Adiplist where adid="&getadid&" and class=2 order by id desc"
		end if
		
		adsrs.open adssql,Conn,1,1
		if adsrs.eof and adsrs.bof then
		Response.Write "<tr align=center><td bgcolor=#ffffff colspan=3>没有记录</td></tr></table>"
		else
		adsrs.pagesize=25 '每页显示的记录数
		totalPut=adsrs.recordcount '记录总数
		totalPage=adsrs.pagecount
		MaxPerPage=adsrs.pagesize
		if currentpage<1 then
		currentpage=1
		end if
		if currentpage>totalPage then
		currentpage=totalPage
		end if
		if currentPage=1 then
		showIpContent
		else
		if (currentPage-1)*MaxPerPage<totalPut then
		adsrs.move  (currentPage-1)*MaxPerPage
		dim bookmark
		bookmark=adsrs.bookmark '移动到开始显示的记录位置
		showIpContent
		end if
		end if
		adsrs.close:set adsrs=nothing
		end if
		Conn.close:set Conn=nothing
		
		End Sub
		
		sub showIpContent
		i=0
		do while not (adsrs.eof or err)
		%>
		<tr   align=center><td class="forumRowHighlight"><font color=red><%=adsrs("id")%></font>　</td><td align=center class="forumRowHighlight"><%=adsrs("ip")%>　</td><td align=center class="forumRowHighlight"><%=adsrs("time")%>　</td></tr>
		<%
		i=i+1
		if i>=MaxPerPage then exit do 
		adsrs.movenext
		loop
		showippages
		end sub 
		
		sub showippages()
		dim n
		n=totalPage
		%>
		</table>
		
		<table border="0" align=center cellpadding="5" cellspacing="1" class=tableBorder width="100%">
		<tr><td align=right colspan=4 class="forumRowHighlight">
	
		<%
  Call KS.ShowPageParamter(totalPut, MaxPerPage, "KS.Ads.asp", True, "条", CurrentPage, KS.QueryParam("page"))
       %>
		
		</td></tr>
		</table>
		<%
	End Sub
	'删除ip记录
	Sub IPDel()
	 Conn.Execute("Delete From KS_Adiplist Where Adid=" & KS.ChkClng(KS.G("ADID")))
	 Response.Redirect "?Action=Listip&id=" & KS.G("adid") & "&ip=" & KS.G("IP")
	End Sub
	
	Sub AdsManage()
	    Dim ttarg
		Dim ComeUrl:ComeUrl=Cstr(Request.ServerVariables("HTTP_REFERER"))
		IF ComeUrl="" Then ComeUrl="Ads_List.asp"
	   %>
		<html>
		<link href="Include/admin_Style.CSS" rel="stylesheet" type="text/css">
		<div align=center>
		<center><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
		  <tr><td align=center>
		<%
		dim getid,RSObj,adssql
		getid=cint(KS.G("id"))
		
		
		Select Case KS.G("job")
			case "close"
		
		   set RSObj=server.createobject("adodb.recordset")
		   adssql="Select id,sitename,act From KS_Advertise where id="&getid
		   RSObj.open adssql,Conn,1,3
		   RSObj("act")=0
		   RSObj.Update
		   Call KS.Alert("广告条[" & RSObj("sitename") & "]被暂停！", ComeUrl)
		  RSObj.close
		
			case "delit"
		    Call KS.Confirm("删除此广告，则其 IP 记录也将被删除！广告及其IP记录被删除后不能恢复！", "?Action=Manage&ComeUrl1=" & Server.URLEncode(ComeUrl) &"&id=" & getid &"&job=del",ComeUrl)		
			case "del"
			conn.execute("delete from KS_UploadFiles Where ChannelID=1020 And InfoID=" & GetID)
			adssql="delete From KS_Advertise where id="&getid
			Conn.execute(adssql)
			dim adssqldelip
			adssqldelip="delete From KS_Adiplist where adid="&getid
			Conn.execute(adssqldelip)
		     Call KS.Alert("广告条删除成功！", KS.G("ComeUrl1"))
         
			case "yulan"
			set RSObj=server.createobject("adodb.recordset")
			adssql="Select id,sitename,intro,gif_url,window,show,place,time,xslei,wid,hei From KS_Advertise where id="&getid
			RSObj.open adssql,Conn,3,3
			
			RSObj("show")=RSObj("show")+1
			RSObj("time")=now()
			RSObj.Update
			if RSObj("window")=0 then
			ttarg = "_blank"
			else
			ttarg="_top"
			end if
			
			Dim GaoAndKuan
			GaoAndKuan=""
			
			if isnumeric(RSObj("hei")) then
			GaoAndKuan=" height="&RSObj("hei")&" "
			else
			
			if right(RSObj("hei"),1)="%" then
				if isnumeric(Left(len(RSObj("hei"))-1))=true then
				 GaoAndKuan=" height="&RSObj("hei")&" "
				end if
			end if
			
		  end if
		
		
		if isnumeric(RSObj("wid")) then
		GaoAndKuan=GaoAndKuan&" width="&RSObj("wid")&" "
		else
			if right(RSObj("wid"),1)="%" then
				if isnumeric(Left(len(RSObj("wid"))-1))=true then 
				GaoAndKuan=GaoAndKuan&" width="&RSObj("wid")&" "
				end if
			end if
		end if
		Select Case RSObj("xslei")
			
					Case "txt"%><a  title="<%=RSObj("sitename")%>"  href="url.asp?id=<%=RSObj("id")%>" target="<%=ttarg%>"><%=RSObj("intro")%></a>
		<%          Case "gif"%><a href="url.asp?id=<%=RSObj("id")%>" target="<%=ttarg%>"><img art="<%=RSObj("sitename")%>" border=0  <%=GaoAndKuan%> src="<%=RSObj("gif_url")%>"></a> 
		<%          Case "swf"%><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http:/download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0"; <%=GaoAndKuan%>><param name=movie value="<%=RSObj("gif_url")%>"><param name=quality value=high>
		
		  <embed src="<%=RSObj("gif_url")%>" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash"  width="<%=RSObj("wid")%>" height="<%=RSObj("hei")%>"></embed></object>
		<%           Case "dai"%><iframe marginwidth=0 marginheight=0  frameborder=0 bordercolor=000000 scrolling=no  name="忠网WEB广告管理系统 zon.cn" src="daima.asp?id=<%=RSObj("id")%>"  <%=GaoAndKuan%>></iframe>
		
		  <%          Case else%><a href="url.asp?id=<%=RSObj("id")%>" target="<%=ttarg%>"><img art="<%=RSObj("sitename")%>"  border=0  <%=GaoAndKuan%> src="<%=RSObj("gif_url")%>"></a>
		<%
				   End Select
		RSObj.close

		case "yulanggw"
			
		%>
		<script language="javascript" src="<%=KS.GetDomain%>plus/ShowA.asp?i=<%=getid%>"></script>
			
		<%
		case "open"
			set RSObj=server.createobject("adodb.recordset")
				adssql="Select id,sitename,act From KS_Advertise where id="&getid
				RSObj.open adssql,Conn,1,3
				RSObj("act")=1
				RSObj.Update
				Call KS.Alert("广告条[" & RSObj("sitename") & "]被激活！", ComeUrl)
				RSObj.close
			
			End Select
			set RSObj=nothing 
			Conn.close:set Conn=nothing
		%>
		</td></tr><tr height=10 align=center>
		  <td><a href="javascript:this.history.go(-1)">返回</a></td>
		</tr></table>
		</center></div>
<%	End Sub
End Class
%> 
