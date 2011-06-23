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
Set KSCls = New Down_Param
KSCls.Kesion()
Set KSCls = Nothing

Class Down_Param
        Private KS,ChannelID
		Private Sub Class_Initialize()
		  Set KS=New PublicCls
		End Sub
        Private Sub Class_Terminate()
		 Call CloseConn()
		 Set KS=Nothing
		End Sub
		Sub Kesion()
		 With KS
			If Not KS.ReturnPowerResult(0, "M010007") Then                  '权限检查
				Call KS.ReturnErr(1, "")   
				Response.End()
			End if
		    
			.echo "<html>"
			.echo "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">"
			.echo "<link href=""Include/Admin_Style.CSS"" rel=""stylesheet"" type=""text/css"">"
			.echo "<script src=""../KS_Inc/common.js"" language=""JavaScript""></script>"
			.echo "<script src=""../KS_Inc/jquery.js"" language=""JavaScript""></script>"
			.echo "<script src='../ks_inc/kesion.box.js'></script>"
			%>
			<script type="text/javascript">
			 var t;
			 function relativeDoc(tag,act,op)
			 {
			  if (tag!='') {
			   popupTips('正在执行'+op+'操作','<div style="height:100px;padding-top:30px;text-align:center"><table style="border:1px solid #000000" width="400" border="0" cellspacing="0" cellpadding="1"><tr><td bgcolor=ffffff height=9><img src="images/114_r2_c2.jpg" width=0 height=10 id=img2 name=img2 align=absmiddle></td></tr></table><div style="margin-top:10px" id="result"><img src="images/loading.gif" align="absmiddle">正在执行您所操作的任务...</div></div>',510,300);
			   $("#closeWindow").attr("disabled",true);
			   $("#closeWindow").removeAttr("onclick");
			  }
			  
			  $.get("../plus/ajaxs.asp",{action:act,flag:tag,num:$("#docNum").val(),channelid:$("#channelid").val()},function(r){
			    var rtn=r.split('|');
				var total=parseInt(rtn[0]);
				var nowid=parseInt(rtn[1]);
				
				if (act=='checkDocFname')
				{ 
				 window.clearTimeout(t);
				  if (total==0){
				   alert('恭喜,没有找到不合法文件名!');
				  }else{
				  alert('共修复了'+total+'篇文档!!!');
				  }
				  closeWindow();
				}
				 
				$("#img2").width((nowid / total) * 400)
				var p=formatNum(nowid/total * 100,2);
				var str='进度'+p+'%<br/>共有<font color=red>'+total+'</font>篇文档,当前正在执行第<font color=red>'+nowid+'</font>篇文档，请不要在执行期间刷新此页面！！！';
				
			    $("#result").html(unescape(str))
				if (total==nowid){
				 window.clearTimeout(t);
				 if (act=='getDocImage')
				 {
				  alert('本次操作累计成功设置了'+rtn[2]+'篇文档!');
				 }
				 closeWindow();
				}
			  });
			  t=window.setTimeout("relativeDoc('','"+act+"','"+op+"')",300);
			 }
			 function checkModelType()
			 {
			   var channelid=parseInt($("#channelid").val());
			   if (channelid==0) {alert('请选择文章类模型!');return false;}
			   $.get("../plus/ajaxs.asp",{action:"getModelType",channelid:channelid},function(t){
			     if (t!=1)
				 {
				  alert('对不起,你选择的基类型不是文章!');
				  return false;
				 }
			   });
			   return true;
			 }
			 function formatNum(Num1,Num2){
				 if(isNaN(Num1)||isNaN(Num2)){
					   return(0);
				 }else{
					   Num1=Num1.toString();
					   Num2=parseInt(Num2);
					   if(Num1.indexOf('.')==-1){
							 return(Num1);
					   }else{
							 var b=Num1.substring(0,Num1.indexOf('.')+Num2+1);
							 var c=Num1.substring(Num1.indexOf('.')+Num2+1,Num1.indexOf('.')+Num2+2);
							 if(c==""){
								   return(b);
							 }else{
								   if(parseInt(c)<5){
										 return(b);
								   }else{
										 return((Math.round(parseFloat(b)*Math.pow(10,Num2))+Math.round(parseFloat(Math.pow(0.1,Num2).toString().substring(0,Math.pow(0.1,Num2).toString().indexOf('.')+Num2+1))*Math.pow(10,Num2)))/Math.pow(10,Num2));
								   }
							 }
					   }
				 }
			}

			</script>
			<%
			.echo "</head>"
			
			.echo "<body bgcolor=""#FFFFFF"" topmargin=""0"" leftmargin=""0"">"
			.echo "      <div class='topdashed sort'>"
			.echo "      一键相关管理工具"
			.echo "      </div>"
			

			.echo "<form action=""?"" method=""post"" name=""myform"">"
			.echo "&nbsp;&nbsp;<strong>通用选项：</strong>仅执行最新添加的<input type='text' value='500' size='4' style='text-align:center' name='docNum' id='docNum'>篇文档 模型限制:<select id='channelid' name='channelid'>"
			.echo " <option value='0'>---请选择模型---</option>"
			.LoadChannelOption 0
			
			.echo "</select>"
			
			.echo "<div class='attention'><strong>操作说明：</strong><br>"
			.echo "       1、文章条数可以输入""0"" 表示全部执行<br /> "
			.echo "      2、当你的文档较多时，运行此功能可能需要较长时间并在执行期间会占用一些服务器资源，建议选择夜间访问人数少时执行。</div>"
			
			.echo " <table width=""99%"" border=""0"" align=""center"" cellspacing=""1"" bgcolor=""#CDCDCD"">"
			.echo "    <tr>"
			.echo "      <td width=""100%"" height=""30"" class='clefttitle'>&nbsp;<font color=""#000080""><b>自动关联相关文档</b></font></td>"
			.echo "    </tr>"
			.echo "    <tr>"
			.echo "      <td width=""25%"" class='tdbg' style='padding:10px'><strong>功能说明：</strong>"
			.echo "       本操作能将文档与文档之间通过各自设定的关键词Tags进行自动关联,以方便供相关文档标签调用时，将关联文档调用出来<br/><br/>"
			.echo "      <table border='0' width='98%' align'center'>"
			.echo "       <tr><td><input type='button' onclick=""relativeDoc('begin','relativeDoc','文档关联')"" class='button' value='一键自动关联'></td></tr>"
			.echo "      </table>"
			.echo "        </td>"
			.echo "    </tr>"
			.echo "  </table>"
			
			.echo " <table width=""99%"" style=""margin-top:5px"" border=""0"" align=""center"" cellspacing=""1"" bgcolor=""#CDCDCD"">"
			.echo "    <tr>"
			.echo "      <td width=""100%"" height=""30"" class='clefttitle'>&nbsp;<font color=""#000080""><b>自动提取内容第一张图片为文档首页图片</b></font></td>"
			.echo "    </tr>"
			.echo "    <tr>"
			.echo "      <td width=""25%"" class='tdbg' style='padding:10px'><strong>功能说明：</strong>"
			.echo "       本操作能从基类型为""文章类""且没有设置缩略图的文档内容中的自动提取第一张图片为做为文档的图片,从而自动转为图片文档,供前台标签调用<br/><br/>"
			.echo "      <table border='0' width='98%' align'center'>"
			.echo "       <tr><td><input type='button' onclick=""if(checkModelType()){relativeDoc('begin','getDocImage','提取文档图片')}"" class='button' value='一键自动提取'></td></tr>"
			.echo "      </table>"
			.echo "        </td>"
			.echo "    </tr>"
			.echo "  </table>"

			.echo " <table width=""99%"" style=""margin-top:5px"" border=""0"" align=""center"" cellspacing=""1"" bgcolor=""#CDCDCD"">"
			.echo "    <tr>"
			.echo "      <td width=""100%"" height=""30"" class='clefttitle'>&nbsp;<font color=""#000080""><b>自动修正文档文件名</b></font></td>"
			.echo "    </tr>"
			.echo "    <tr>"
			.echo "      <td width=""25%"" class='tdbg' style='padding:10px'><strong>功能说明：</strong>"
			.echo "       本操作能自动检测文档生成静态Html的文件名是否合法,如果不合法将自动修正,以免在生成静态操作时出错.<br/><br/>"
			.echo "      <table border='0' width='98%' align'center'>"
			.echo "       <tr><td><input type='button' onclick=""relativeDoc('begin','checkDocFname','自动修正文件名')"" class='button' value='一键修正文件名'></td></tr>"
			.echo "      </table>"
			.echo "        </td>"
			.echo "    </tr>"
			.echo "  </table>"
			
			.echo "</form>"
			.echo "</body>"
			.echo "</html>"
			
			End With
		End Sub

End Class
%> 
