$(document).ready(function(){
 init_reg();
})

function getlicense()
{
  if ($("#viewlicense").attr("checked")==true)
  {
    $("#license").show();
  }
  else
  {
   $("#license").hide();
  }
}
function getCode()
{
 $("#showVerify").html("<img style='cursor:pointer' src='../../plus/verifycode.asp?n='+Math.random() onClick='this.src=\"../../plus/verifycode.asp?n=\"+ Math.random();'  align='absmiddle'>");	
}
var msg	;
var bname_m=false;
var ajaxchk=null;
var ajaxstr=null;
function init_reg(){
	msg=new Array(
	"������"+minlen+"-"+maxlen+"λ�ַ���Ӣ�ġ����֡��»��ߵ���ϡ�",
	"������4-14λ�ַ���Ӣ�ġ����ֵ���ϡ�",
	"������6λ�����ַ���������ո�",
	"���ظ�������������롣",
	"��ѡ��������ʾ���⡣",
	"6���ַ������ֻ�3���������ϣ�����6������",
	"�����������õĵ��������ַ��",
	"��������壬���Ե������ˢ����֤�롣",
	"������Ϸ����ֻ����롣",
	"ֻ����ȷ�ش�ע������ſ��Լ�����"
	)
	document.getElementById("usernamemsg").innerHTML=msg[0];
	document.getElementById("passwordmsg1").innerHTML=msg[2];
	document.getElementById("passwordmsg2").innerHTML=msg[3];
	document.getElementById("questionmsg").innerHTML=msg[4];
	document.getElementById("answermsg").innerHTML=msg[5];
	document.getElementById("emailmsg").innerHTML=msg[6];
	document.getElementById("chkcodemsg").innerHTML=msg[7];
	document.getElementById("mobilemsg").innerHTML=msg[8];
	document.getElementById("reganswermsg").innerHTML=msg[9];
}

function on_input(objname){
	var strtxt;
	var obj=document.getElementById(objname);
	obj.className="d_on";
	//alert(objname);
	switch (objname){
		case "usernamemsg":
			strtxt=msg[0];
			break;
		case "passwordmsg1":
			strtxt=msg[2];
			break;
		case "passwordmsg2":
			strtxt=msg[3];
			break;
		case "answermsg":
			strtxt=msg[5];
			break;
		case "emailmsg":
			strtxt=msg[6];
			break;
		case "chkcodemsg":
		    strtxt=msg[7];
			break;	
		case "mobilemsg":
		    strtxt=msg[8];
			break;
		case "reganswermsg":
		    strtxt=msg[9];
			break;
	}
	obj.innerHTML=strtxt;
}
function out_username(){
	var obj=document.getElementById("usernamemsg");
	var str=sl(document.getElementById("UserName").value);
	var chk=true;
	if (str<minlen || str>maxlen){chk=false;}
	if (!chk){
		obj.className="d_err";
		obj.innerHTML=msg[0];
		return;
	}
	$.get("ajax_check.asp",{action:"checkusername",username:escape(document.getElementById("UserName").value)},function(d){
	     var s=unescape(d);
		 ajaxchk=s.split('|')[0];
		 ajaxstr=s.split('|')[1];
	});
	if (ajaxstr!=null){
		if (ajaxchk=='ok'){
		  obj.className="d_ok";
		  obj.innerHTML=ajaxstr;
		 }else{
			obj.className="d_err";
			obj.innerHTML=ajaxstr;
		 }
	}
}
function out_password1(){
	var obj=document.getElementById("passwordmsg1");
	var str=document.getElementById("PassWord").value;
	var chk=true;
	if (str=='' || str.length<6 || str.length>14){chk=false;}
	if (chk){
		obj.className="d_ok";
		obj.innerHTML='�����Ѿ����롣';
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[2];
	}
	return chk;
}
function out_password2(){
	var obj=document.getElementById("passwordmsg2");
	var str=document.getElementById("RePassWord").value;
	var chk=true;
	if (str!=document.getElementById("PassWord").value||str==''){chk=false;}
	if (chk){
		obj.className="d_ok";
		obj.innerHTML='�ظ�����������ȷ��';
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[3];
	}
	return chk;
}
function out_question(){
	var obj=document.getElementById("questionmsg");
	var str=document.getElementById("Question").value;
	var chk=true;
	if (question==0) return true;
	if (str==''){chk=false}
	if (chk){
		obj.className="d_ok";
		obj.innerHTML='������ʾ�����Ѿ�ѡ��';
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[4];
	}
	return chk;
}
function out_answer(){
	var obj=document.getElementById("answermsg");
	var str=sl(document.getElementById("Answer").value);
	var chk=true;
	if (question==0) return true;
	if (str<6 || str>40){chk=false}
	if (chk){
		obj.className="d_ok";
		obj.innerHTML='������ʾ������Ѿ����롣';
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[5];
	}
	return chk;
}
function out_mobile(){
	var obj=document.getElementById("mobilemsg");
	var str=document.getElementById("Mobile").value;
	if (mobile==0) return true;
	var chk=ismobile(str);
	if (chk){
		obj.className="d_ok";
		obj.innerHTML='�ֻ����������롣';
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[8];
	}
	return chk;	
}
function ismobile(s)
{
   var p = /^(013|015|13|15|018|18)\d{9}$/;
   if(s.match(p) != null){
  return true;
  }
  return false;
}
function out_email(){
	var obj=document.getElementById("emailmsg");
	var str=document.getElementById("Email").value;
	var chk=true;
	if (str==''|| !str.match(/^[\w\.\-]+@([\w\-]+\.)+[a-z]{2,4}$/ig)){chk=false}
	if (chk){
		obj.className="d_ok";
		obj.innerHTML='���������ַ�Ѿ����롣';
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[6];
		return chk;
	}
	$.get("ajax_check.asp",{action:"checkemail",email:escape(str)},function(d){
	     var s=unescape(d);
		 ajaxchk=s.split('|')[0];
		 ajaxstr=s.split('|')[1];
		if (ajaxstr!=null){
		if (ajaxchk=='ok'){
		  obj.className="d_ok";
		  obj.innerHTML=ajaxstr;
		 }else{
			obj.className="d_err";
			obj.innerHTML=ajaxstr;
		 }
		}

	});
			

}

function out_chkcode()
{	var obj=document.getElementById("chkcodemsg");
	var str=sl(document.getElementById("Verifycode").value);
	var chk=true;
	if (str<4 || str>6){chk=false}
	if (chk){
		obj.className="d_ok";
		obj.innerHTML='��֤���Ѿ����롣';
	}else{
		obj.className="d_err";
		obj.innerHTML=msg[7];
	return chk;
	}
	$.get("ajax_check.asp",{action:"checkcode",code:escape(document.getElementById("Verifycode").value)},function(d){
	     var s=unescape(d);
		 ajaxchk=s.split('|')[0];
		 ajaxstr=s.split('|')[1];
	})
	if (ajaxstr!=null){
		if (ajaxchk=='ok'){
		  obj.className="d_ok";
		  obj.innerHTML=ajaxstr;
		 }else{
			obj.className="d_err";
			obj.innerHTML=ajaxstr;
		 }
	 }
}
function sl(st){
	sl1=st.length;
	strLen=0;
	for(i=0;i<sl1;i++){
		if(st.charCodeAt(i)>255) strLen+=2;
	 else strLen++;
	}
	return strLen;
}

	 
      function CheckForm() 
		{ 
		   if ($("#viewlicense").attr("checked")!=true)
		   {
			  alert("ֻ���Ķ�����ȫ���ܻ�Ա��������ſ��Լ���ע��!")
			  return false;
			}
		
			if (document.myform.UserName.value =="")
			{
			alert("����д���Ļ�Ա����");
			document.myform.UserName.focus();
			return false;
			}
			//var filter=/^\s*[.A-Za-z0-9_-]{{$Show_UserNameLimitChar},{$Show_UserNameMaxChar}}\s*$/;
			//if (!filter.test(document.myform.UserName.value)) { 
			//alert("��Ա����д����ȷ,��������д����ʹ�õ��ַ�Ϊ��A-Z a-z 0-9 _ - .)���Ȳ�С��{$Show_UserNameLimitChar}���ַ���������{$Show_UserNameMaxChar}���ַ���ע�ⲻҪʹ�ÿո�"); 
			//document.myform.UserName.focus();
			//return false; 
			//} 
			if (document.myform.PassWord.value =="") 
			{
			alert("����д�������룡");
			document.myform.PassWord.focus();
			return false; 
			}
			if(document.myform.RePassWord.value==""){
			alert("����������ȷ�����룡");
			document.myform.RePassWord.focus();
			return false;
			}
			var filter=/^\s*[.A-Za-z0-9_-]{6,15}\s*$/;
			if (!filter.test(document.myform.PassWord.value)) { 
			alert("������д����ȷ,��������д����ʹ�õ��ַ�Ϊ��A-Z a-z 0-9 _ - .)���Ȳ�С��6���ַ���������15���ַ���ע�ⲻҪʹ�ÿո�"); 
			document.myform.PassWord.focus();
			return false; 
			} 
			if (document.myform.PassWord.value!=document.myform.RePassWord.value ){
			alert("������д�����벻һ�£���������д��"); 
			document.myform.PassWord.focus();
			return false; 
			} 
			if (document.myform.Question.value ==""&&question==1)
			{
			alert("����д�����������⣡");
			document.myform.Question.focus();
			return false;
			}
			if (document.myform.Answer.value ==""&&question==1)
			{
			alert("����д��������𰸣�");
			document.myform.Answer.focus();
			return false;
			}
			if (document.myform.Mobile.value ==""&&mobile==1)
			{
			alert("����д�����ֻ����룡");
			document.myform.Mobile.focus();
			return false;
			}
			else if(ismobile(document.myform.Mobile.value)==false&&mobile==1)
			{
			alert("�����ֻ����벻��ȷ��");
			document.myform.Mobile.focus();
			return false;
			}
			
			if (document.myform.Email.value =="")
			{
			alert("���������ĵ����ʼ���ַ��");
			document.myform.Email.focus();
			return false;
			}
			if((document.myform.Email.value.indexOf("@")==-1)||(document.myform.Email.value.indexOf(".")==-1))
			{
				alert("������ĵ����ʼ���ַ����");
				document.myform.Email.focus();
				return false;
				}
				return true;
		}
