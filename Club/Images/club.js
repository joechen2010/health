 function movetopic(ev,topicid,title)
 {
		  mousepopup(ev,"<img src='images/p_up.gif' align='absmiddle'>�����ƶ�","<form name='moveform' action='?' method='get'><br/><b>�ƶ����ӣ�</b>"+title+"<br/><br/><b>�Ƶ����棺</b><span id='showboardselect'></span><div style='text-align:center;margin:20px'><input type='submit' value='ȷ���ƶ�' class='button'><input type='hidden' value="+topicid+" name='id' id='id'><input type='hidden' value='movetopic' name='action'><input type='button' value=' ȡ �� ' onclick='closeWindow()' class='button'></div></form>",350);
		  $.get("../plus/ajaxs.asp",{action:"GetClubBoard"},function(r){
		    $("#showboardselect").html(unescape(r));
		   });
 }
 
 function checkmsg()
 {   var message=escape($("#message").val());
	 var username=escape($("#username").val());
	 if (username==''){
			  alert('�������ݳ���!');
			  closeWindow();
	 }
	 if (message==''){
			   alert('��������Ϣ����!');
			   $("#message").focus();
			   return false;
	 }
	 $.get("../plus/ajaxs.asp",{action:"SendMsg",username:username,message:message},function(r){
			   r=unescape(r);
			   if (r!='success'){
				alert(r);
			   }else{
				 alert('��ϲ��������Ϣ�ѷ���!');
				 closeWindow();
			   }
			 });
 }
 function sendMsg(ev,username)
		 {
		  mousepopup(ev,"<img src='/images/user/mail.gif' align='absmiddle'>������Ϣ","�Է���¼����Կ���������Ϣ(������255���ַ�)<br /><textarea name='message' id='message' style='width:340px;height:80px'></textarea><div style='text-align:center;margin:10px'><input type='button' onclick='return(checkmsg())' value=' ȷ �� ' class='button'><input type='hidden' value="+username+" name='username' id='username'> <input type='button' value=' ȡ �� ' onclick='closeWindow()' class='button'></div>",350);
		  $.get("/plus/ajaxs.asp",{action:"CheckLogin"},function(r){
		   if (r!='true'){
			 ShowLogin();
			}
		   });
		 }
        function check()
		{
		 var message=escape($("#message").val());
		 var username=escape($("#username").val());
		 if (username==''){
		  alert('�������ݳ���!');
		  closeWindow();
		 }
		 if (message==''){
		   alert('�����븽��!');
		   $("#message").focus();
		   return false;
		 }
		 $.get("/plus/ajaxs.asp",{action:"AddFriend",username:username,message:message},function(r){
		   r=unescape(r);
		   if (r!='success'){
		    alert(r);
		   }else{
		     alert('���������ѷ���,��ȴ��Է���ȷ��!');
			 closeWindow();
		   }
		 });
		}
		function addF(ev,username)
		{ 
		 show(ev,username);
		 var isMyFriend=false;
		 $.get("/plus/ajaxs.asp",{action:"CheckMyFriend",username:escape(username)},function(b){
		    if (b=='nologin'){
			  closeWindow();
			  ShowLogin();
			}else if (b=='true'){
			  closeWindow();
			  alert('�û�['+username+']�Ѿ������ĺ����ˣ�');
			  return false;
			 }else if(b=='verify'){
			  closeWindow();
			  alert('���������['+username+'],��ȴ��Է�����֤!');
			  return false;
			 }else{
			 }
		 })
		 
		}
		function show(ev,username)
		{
		 mousepopup(ev,"<img src='/images/user/log/106.gif'>��Ӻ���","ͨ���Է���֤���ܳ�Ϊ����(������255���ַ�)<br /><textarea name='message' id='message' style='width:340px;height:80px'></textarea><div style='text-align:center;margin:10px'><input type='button' onclick='return(check())' value=' ȷ �� ' class='button'><input type='hidden' value="+username+" name='username' id='username'> <input type='button' value=' ȡ �� ' onclick='closeWindow()' class='button'></div>",350);
		}
		function ShowLogin()
		{ 
		 popupIframe('��Ա��¼','/user/userlogin.asp?Action=Poplogin',397,184,'no');
		}

function checksearch()
{
     if ($("#keyword").val()=="")
	 {
	  alert('������ؼ���!');
	  $('#keyword').focus();
	  return false;
	 }
}


function scrollDoor(){
}
scrollDoor.prototype = {
	sd : function(menus,divs,openClass,closeClass){
		var _this = this;
		if(menus.length != divs.length)
		{
			alert("�˵������������ݲ�������һ��!");
			return false;
		}				
		for(var i = 0 ; i < menus.length ; i++)
		{	
			_this.$(menus[i]).value = i;				
			_this.$(menus[i]).onmouseover = function(){
					
				for(var j = 0 ; j < menus.length ; j++)
				{						
					_this.$(menus[j]).className = closeClass;
					_this.$(divs[j]).style.display = "none";
				}
				_this.$(menus[this.value]).className = openClass;	
				_this.$(divs[this.value]).style.display = "block";				
			}
		}
		},
	$ : function(oid){
		if(typeof(oid) == "string")
		return document.getElementById(oid);
		return oid;
	}
}
window.onload = function(){
	var SDmodel = new scrollDoor();
	try
   {
	SDmodel.sd(["tb_1","tb_2","tb_3"],["tbc_01","tbc_02","tbc_03"],"open","close");
   }catch(e){
   }
}
