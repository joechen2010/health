 function movetopic(ev,topicid,title)
 {
		  mousepopup(ev,"<img src='images/p_up.gif' align='absmiddle'>帖子移动","<form name='moveform' action='?' method='get'><br/><b>移动帖子：</b>"+title+"<br/><br/><b>移到版面：</b><span id='showboardselect'></span><div style='text-align:center;margin:20px'><input type='submit' value='确定移动' class='button'><input type='hidden' value="+topicid+" name='id' id='id'><input type='hidden' value='movetopic' name='action'><input type='button' value=' 取 消 ' onclick='closeWindow()' class='button'></div></form>",350);
		  $.get("../plus/ajaxs.asp",{action:"GetClubBoard"},function(r){
		    $("#showboardselect").html(unescape(r));
		   });
 }
 
 function checkmsg()
 {   var message=escape($("#message").val());
	 var username=escape($("#username").val());
	 if (username==''){
			  alert('参数传递出错!');
			  closeWindow();
	 }
	 if (message==''){
			   alert('请输入消息内容!');
			   $("#message").focus();
			   return false;
	 }
	 $.get("../plus/ajaxs.asp",{action:"SendMsg",username:username,message:message},function(r){
			   r=unescape(r);
			   if (r!='success'){
				alert(r);
			   }else{
				 alert('恭喜，您的消息已发送!');
				 closeWindow();
			   }
			 });
 }
 function sendMsg(ev,username)
		 {
		  mousepopup(ev,"<img src='/images/user/mail.gif' align='absmiddle'>发送消息","对方登录后可以看到您的消息(可输入255个字符)<br /><textarea name='message' id='message' style='width:340px;height:80px'></textarea><div style='text-align:center;margin:10px'><input type='button' onclick='return(checkmsg())' value=' 确 定 ' class='button'><input type='hidden' value="+username+" name='username' id='username'> <input type='button' value=' 取 消 ' onclick='closeWindow()' class='button'></div>",350);
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
		  alert('参数传递出错!');
		  closeWindow();
		 }
		 if (message==''){
		   alert('请输入附言!');
		   $("#message").focus();
		   return false;
		 }
		 $.get("/plus/ajaxs.asp",{action:"AddFriend",username:username,message:message},function(r){
		   r=unescape(r);
		   if (r!='success'){
		    alert(r);
		   }else{
		     alert('您的请求已发送,请等待对方的确认!');
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
			  alert('用户['+username+']已经是您的好友了！');
			  return false;
			 }else if(b=='verify'){
			  closeWindow();
			  alert('您已邀请过['+username+'],请等待对方的认证!');
			  return false;
			 }else{
			 }
		 })
		 
		}
		function show(ev,username)
		{
		 mousepopup(ev,"<img src='/images/user/log/106.gif'>添加好友","通过对方验证才能成为好友(可输入255个字符)<br /><textarea name='message' id='message' style='width:340px;height:80px'></textarea><div style='text-align:center;margin:10px'><input type='button' onclick='return(check())' value=' 确 定 ' class='button'><input type='hidden' value="+username+" name='username' id='username'> <input type='button' value=' 取 消 ' onclick='closeWindow()' class='button'></div>",350);
		}
		function ShowLogin()
		{ 
		 popupIframe('会员登录','/user/userlogin.asp?Action=Poplogin',397,184,'no');
		}

function checksearch()
{
     if ($("#keyword").val()=="")
	 {
	  alert('请输入关键字!');
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
			alert("菜单层数量和内容层数量不一样!");
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
