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
			If Not KS.ReturnPowerResult(0, "M010007") Then                  'Ȩ�޼��
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
			   popupTips('����ִ��'+op+'����','<div style="height:100px;padding-top:30px;text-align:center"><table style="border:1px solid #000000" width="400" border="0" cellspacing="0" cellpadding="1"><tr><td bgcolor=ffffff height=9><img src="images/114_r2_c2.jpg" width=0 height=10 id=img2 name=img2 align=absmiddle></td></tr></table><div style="margin-top:10px" id="result"><img src="images/loading.gif" align="absmiddle">����ִ����������������...</div></div>',510,300);
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
				   alert('��ϲ,û���ҵ����Ϸ��ļ���!');
				  }else{
				  alert('���޸���'+total+'ƪ�ĵ�!!!');
				  }
				  closeWindow();
				}
				 
				$("#img2").width((nowid / total) * 400)
				var p=formatNum(nowid/total * 100,2);
				var str='����'+p+'%<br/>����<font color=red>'+total+'</font>ƪ�ĵ�,��ǰ����ִ�е�<font color=red>'+nowid+'</font>ƪ�ĵ����벻Ҫ��ִ���ڼ�ˢ�´�ҳ�棡����';
				
			    $("#result").html(unescape(str))
				if (total==nowid){
				 window.clearTimeout(t);
				 if (act=='getDocImage')
				 {
				  alert('���β����ۼƳɹ�������'+rtn[2]+'ƪ�ĵ�!');
				 }
				 closeWindow();
				}
			  });
			  t=window.setTimeout("relativeDoc('','"+act+"','"+op+"')",300);
			 }
			 function checkModelType()
			 {
			   var channelid=parseInt($("#channelid").val());
			   if (channelid==0) {alert('��ѡ��������ģ��!');return false;}
			   $.get("../plus/ajaxs.asp",{action:"getModelType",channelid:channelid},function(t){
			     if (t!=1)
				 {
				  alert('�Բ���,��ѡ��Ļ����Ͳ�������!');
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
			.echo "      һ����ع�����"
			.echo "      </div>"
			

			.echo "<form action=""?"" method=""post"" name=""myform"">"
			.echo "&nbsp;&nbsp;<strong>ͨ��ѡ�</strong>��ִ��������ӵ�<input type='text' value='500' size='4' style='text-align:center' name='docNum' id='docNum'>ƪ�ĵ� ģ������:<select id='channelid' name='channelid'>"
			.echo " <option value='0'>---��ѡ��ģ��---</option>"
			.LoadChannelOption 0
			
			.echo "</select>"
			
			.echo "<div class='attention'><strong>����˵����</strong><br>"
			.echo "       1������������������""0"" ��ʾȫ��ִ��<br /> "
			.echo "      2��������ĵ��϶�ʱ�����д˹��ܿ�����Ҫ�ϳ�ʱ�䲢��ִ���ڼ��ռ��һЩ��������Դ������ѡ��ҹ�����������ʱִ�С�</div>"
			
			.echo " <table width=""99%"" border=""0"" align=""center"" cellspacing=""1"" bgcolor=""#CDCDCD"">"
			.echo "    <tr>"
			.echo "      <td width=""100%"" height=""30"" class='clefttitle'>&nbsp;<font color=""#000080""><b>�Զ���������ĵ�</b></font></td>"
			.echo "    </tr>"
			.echo "    <tr>"
			.echo "      <td width=""25%"" class='tdbg' style='padding:10px'><strong>����˵����</strong>"
			.echo "       �������ܽ��ĵ����ĵ�֮��ͨ�������趨�Ĺؼ���Tags�����Զ�����,�Է��㹩����ĵ���ǩ����ʱ���������ĵ����ó���<br/><br/>"
			.echo "      <table border='0' width='98%' align'center'>"
			.echo "       <tr><td><input type='button' onclick=""relativeDoc('begin','relativeDoc','�ĵ�����')"" class='button' value='һ���Զ�����'></td></tr>"
			.echo "      </table>"
			.echo "        </td>"
			.echo "    </tr>"
			.echo "  </table>"
			
			.echo " <table width=""99%"" style=""margin-top:5px"" border=""0"" align=""center"" cellspacing=""1"" bgcolor=""#CDCDCD"">"
			.echo "    <tr>"
			.echo "      <td width=""100%"" height=""30"" class='clefttitle'>&nbsp;<font color=""#000080""><b>�Զ���ȡ���ݵ�һ��ͼƬΪ�ĵ���ҳͼƬ</b></font></td>"
			.echo "    </tr>"
			.echo "    <tr>"
			.echo "      <td width=""25%"" class='tdbg' style='padding:10px'><strong>����˵����</strong>"
			.echo "       �������ܴӻ�����Ϊ""������""��û����������ͼ���ĵ������е��Զ���ȡ��һ��ͼƬΪ��Ϊ�ĵ���ͼƬ,�Ӷ��Զ�תΪͼƬ�ĵ�,��ǰ̨��ǩ����<br/><br/>"
			.echo "      <table border='0' width='98%' align'center'>"
			.echo "       <tr><td><input type='button' onclick=""if(checkModelType()){relativeDoc('begin','getDocImage','��ȡ�ĵ�ͼƬ')}"" class='button' value='һ���Զ���ȡ'></td></tr>"
			.echo "      </table>"
			.echo "        </td>"
			.echo "    </tr>"
			.echo "  </table>"

			.echo " <table width=""99%"" style=""margin-top:5px"" border=""0"" align=""center"" cellspacing=""1"" bgcolor=""#CDCDCD"">"
			.echo "    <tr>"
			.echo "      <td width=""100%"" height=""30"" class='clefttitle'>&nbsp;<font color=""#000080""><b>�Զ������ĵ��ļ���</b></font></td>"
			.echo "    </tr>"
			.echo "    <tr>"
			.echo "      <td width=""25%"" class='tdbg' style='padding:10px'><strong>����˵����</strong>"
			.echo "       ���������Զ�����ĵ����ɾ�̬Html���ļ����Ƿ�Ϸ�,������Ϸ����Զ�����,���������ɾ�̬����ʱ����.<br/><br/>"
			.echo "      <table border='0' width='98%' align'center'>"
			.echo "       <tr><td><input type='button' onclick=""relativeDoc('begin','checkDocFname','�Զ������ļ���')"" class='button' value='һ�������ļ���'></td></tr>"
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
