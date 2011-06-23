var sAgent=navigator.userAgent.toLowerCase();
var IsIE=sAgent.indexOf("msie")!=-1;
function $$(_sId){return document.getElementById(_sId)}

function dialog(appurl){
	var width = 300;
	var height = 150;
	var src = "";
	var path = appurl+"ks_inc/dialog/";
	//alert(path);
	var sFunc = '<input id="dialogOk" type="button" style="font-size:12px;width:71px;height:22px;line-height:24px;border-style:none;background:transparent url('+path+'button4.bmp);width: 71px;height: 22px;"  onmouseover=BtnOver(this,"'+path+'") onmouseout=BtnOut(this,"'+path+'") value="确 认" onclick="new dialog(\''+appurl+'\').reset(1);" /> <input id="dialogCancel" type="button" style="font-size:12px;width:71px;height:22px;line-height:24px;border-style:none;background:transparent url('+path+'button4.bmp);width: 71px;height: 22px;" value="取 消" onclick="new dialog(\''+appurl+'\').reset(0);" />';
	var sClose = '<span id="dialogBoxClose" onclick="new dialog(\''+appurl+'\').reset(0);" style="color: #fff; cursor:pointer; ">关闭</span>';
	var sBody = '\
		<table id="dialogBodyBox" border="0" align="center" cellpadding="0" cellspacing="0" width="100%" height="100%" >\
			<tr height="10"><td colspan="4" align="center"></td></tr>\
			<tr>\
				<td width="10"></td>\
				<td width="80" align="center" valign="middle" id="ob_boxface"><img id="dialogBoxFace" valign="absmiddle" /></td>\
				<td id="dialogMsg" style="font-size:12px;color:#000;"></td>\
				<td width="10"></td>\
			</tr>\
			<tr height="10"><td colspan="4" align="center"></td></tr>\
			<tr><td id="dialogFunc" colspan="4" align="center">' + sFunc + '</td></tr>\
			<tr height="10"><td colspan="4" align="center"></td></tr>\
		</table>\
	';
	var sBox = '\
		<table id="dialogBox" width="' + width + '" border="0" cellpadding="0" cellspacing="0" style=" border: 1px solid #1B76B7; display: none; z-index: 1000; ">\
			<tr height="24" bgcolor="#1B76B7" >\
				<td>\
					<table onselectstart="return false;" style="-moz-user-select:none;" width="100%" border="0" cellpadding="0" cellspacing="0">\
						<tr>\
							<td width="6" ></td>\
							<td id="dialogBoxTitle" onmousedown="new dialog().moveStart(event, \'dialogBox\')" style="color:#fff;cursor:move;font-size:12px;font-weight:bold;">提示信息</td>\
							<td id="dialogClose" width="27" align="right" valign="middle">\
								' + sClose + '\
							</td>\
							<td width="6"></td>\
						</tr>\
					</table>\
				</td>\
			</tr>\
			<tr id="dialogHeight" height="' + height + '">\
				<td id="dialogBody" style="background:#fff;color:#000;">' + sBody + '</td>\
			</tr>\
		</table>\
		<div id="dialogBoxShadow" style=" display:none; z-index:9; "></div>\
		<iframe id="dialogBoxDivShim" scrolling="no" frameborder="0" style="position: absolute; top: 0px; left: 0px; display: none; ">\
	';
	this.show = function(){$$('dialogBodyBox') ? function(){} : this.init();this.middle('dialogBox');this.shadow();}
	this.reset = function(_int){
	if (_int == 1 && "登录窗口,主持人及嘉宾登录窗口".indexOf($$('dialogBoxTitle').innerHTML) > -1) {
	        return;
	    }
	    $$('dialogBox').style.display='none';
	    $$('dialogBoxShadow').style.display = "none";
	    $$('dialogBoxDivShim').style.display = "none";
	    $$('dialogBody').innerHTML = sBody;
	}
	this.html = function(_sHtml){$$("dialogBody").innerHTML = _sHtml;this.show();}
	this.init = function(){
		$$('dialogCase') ? $$('dialogCase').parentNode.removeChild($$('dialogCase')) : function(){};
		var oDiv = document.createElement('span');
		oDiv.id = "dialogCase";
		oDiv.innerHTML = sBox;
		document.body.appendChild(oDiv);
	}
	this.button = function(_sId, _sFuc){
			if($$(_sId)){
				$$(_sId).style.display = '';
				if($$(_sId).addEventListener){
					if($$(_sId).act){$$(_sId).removeEventListener('click', function(){eval($$(_sId).act)}, false);}
					$$(_sId).act = _sFuc;
					$$(_sId).addEventListener('click', function(){eval(_sFuc)}, false);
				}else{
					if($$(_sId).act){$$(_sId).detachEvent('onclick', function(){eval($$(_sId).act)});}
					$$(_sId).act = _sFuc;
					$$(_sId).attachEvent('onclick', function(){eval(_sFuc)});
				}
			}
	}
	this.shadow = function(){
		var oShadow = $$('dialogBoxShadow');
		var oDialog = $$('dialogBox');
		var IfrRef = $$('dialogBoxDivShim');
		oShadow.style.position = "absolute";
		oShadow.style.background	= "#000";
		oShadow.style.display	= "";
		oShadow.style.opacity	= "0.2";
		oShadow.style.filter = "alpha(opacity=0)";
		oShadow.style.top = oDialog.offsetTop + 0+"px";
		oShadow.style.left = oDialog.offsetLeft + 0+"px";
		oShadow.style.width = oDialog.offsetWidth+"px";
		oShadow.style.height = oDialog.offsetHeight+"px";
		
		IfrRef.style.width = oDialog.offsetWidth+0+"px";
		IfrRef.style.height = oDialog.offsetHeight+0+"px";
		IfrRef.style.top = oDialog.offsetTop+"px";
		IfrRef.style.left = oDialog.offsetLeft+"px";
		IfrRef.style.zIndex = oDialog.style.zIndex - 1;
		IfrRef.style.display = "block";
	}
	this.open = function(_sUrl, _sMode){
		this.show();
		if(!_sMode || _sMode == "no" || _sMode == "yes"){
			$$("dialogBody").innerHTML = "<iframe id='dialogFrame' width='100%' height='100%' frameborder='0' scrolling='" + _sMode + "'></iframe>";
			$$("dialogFrame").src = _sUrl;
		}
	}
	this.event = function(_sMsg, _sOk, _sCancel, _sClose){
		$$('dialogFunc').innerHTML = sFunc;
		$$('dialogClose').innerHTML = sClose;
		$$('dialogBodyBox') == null ? $$('dialogBody').innerHTML = sBody : function(){};
		$$('dialogMsg') ? $$('dialogMsg').innerHTML = _sMsg  : function(){};
		this.show();
		_sOk ? this.button('dialogOk', _sOk) | $$('dialogOk').focus() : $$('dialogOk').style.display = 'none';
		_sCancel ? this.button('dialogCancel', _sCancel) : $$('dialogCancel').style.display = 'none';
		_sClose ? this.button('dialogBoxClose', _sClose) : function(){};
		//_sOk ? this.button('dialogOk', _sOk) : _sOk == "" ? function(){} : $$('dialogOk').style.display = 'none';
		//_sCancel ? this.button('dialogCancel', _sCancel) : _sCancel == "" ? function(){} : $$('dialogCancel').style.display = 'none';
	}
	this.set = function(_oAttr, _sVal){
		var oShadow = $$('dialogBoxShadow');
		var oDialog = $$('dialogBox');
		var oHeight = $$('dialogHeight');

		if(_sVal != ''){
			switch(_oAttr){
				case 'title':
					$$('dialogBoxTitle').innerHTML = _sVal;
					//title = _sVal;
					break;
				case 'width':
					oDialog.style.width = _sVal;
					width = _sVal;
					break;
				case 'height':
					oHeight.style.height = _sVal;
					height = _sVal;
					break;
				case 'src':
					if(parseInt(_sVal) >= 0){
						$$('dialogBoxFace') ? $$('dialogBoxFace').src = path + _sVal + '.png' : function(){};
					}else{
						$$('dialogBoxFace') ? $$('dialogBoxFace').src = _sVal : function(){};
					}
					src = _sVal;
					break;
				case "html":
					$$("ob_boxface").innerHTML=_sVal;
					break;
			}
		}
		this.middle('dialogBox');
		oShadow.style.top = oDialog.offsetTop + 0+"px";
		oShadow.style.left = oDialog.offsetLeft + 0+"px";
		oShadow.style.width = oDialog.offsetWidth+"px";
		oShadow.style.height = oDialog.offsetHeight+"px";
	}
	this.moveStart = function (event, _sId){
		var oObj = $$(_sId);
		oObj.onmousemove = mousemove;
		oObj.onmouseup = mouseup;
		oObj.setCapture ? oObj.setCapture() : function(){};
		oEvent = window.event ? window.event : event;
		var dragData = {x : oEvent.clientX, y : oEvent.clientY};
		var backData = {x : parseInt(oObj.style.top), y : parseInt(oObj.style.left)};
		function mousemove(){
			var oEvent = window.event ? window.event : event;
			var iLeft = oEvent.clientX - dragData["x"] + parseInt(oObj.style.left);
			var iTop = oEvent.clientY - dragData["y"] + parseInt(oObj.style.top);
			oObj.style.left = iLeft+"px";
			oObj.style.top = iTop+"px";
			$$('dialogBoxShadow').style.left = iLeft + 0+"px";
			$$('dialogBoxShadow').style.top = iTop + 0+"px";
			
			$$('dialogBoxDivShim').style.left = iLeft+"px";
			$$('dialogBoxDivShim').style.top = iTop +"px";
			
			dragData = {x: oEvent.clientX, y: oEvent.clientY};
			

		}
		function mouseup(){
			var oEvent = window.event ? window.event : event;
			oObj.onmousemove = null;
			oObj.onmouseup = null;
			if(oEvent.clientX < 1 || oEvent.clientY < 1 || oEvent.clientX > document.body.clientWidth || oEvent.clientY > document.body.clientHeight){
				oObj.style.left = backData.y +"px";
				oObj.style.top = backData.x +"px";
				$$('dialogBoxShadow').style.left = backData.y + 0+"px";
				$$('dialogBoxShadow').style.top = backData.x + 0+"px";
				
				$$('dialogBoxDivShim').style.left = backData.y +"px";
				$$('dialogBoxDivShim').style.top = backData.x +"px";
			}
			oObj.releaseCapture ? oObj.releaseCapture() : function(){};
		}
	}
	this.middle = function(_sId){	
		var theWidth;
		var theHeight;
		if (document.documentElement && document.documentElement.clientWidth) { 
			theWidth = document.documentElement.clientWidth+document.documentElement.scrollLeft*2;;
			theHeight = document.documentElement.clientHeight+document.documentElement.scrollTop*2;; 
		} else if (document.body) { 
			theWidth = document.body.clientWidth;
			theHeight = document.body.clientHeight; 
		}else if(window.innerWidth){
			theWidth = window.innerWidth;
			theHeight = window.innerHeight;
		}
		document.getElementById(_sId).style.display = '';
		document.getElementById(_sId).style.position = "absolute";
		document.getElementById(_sId).style.left = (theWidth / 2) - (document.getElementById(_sId).offsetWidth / 2)+"px";
		//alert(in_ob_useradmin);
		if(document.all||document.getElementById("user_page_top")){
			document.getElementById(_sId).style.top = (theHeight / 2 + document.body.scrollTop) - (document.getElementById(_sId).offsetHeight / 2)+"px";
		}else{
			var sClientHeight = parent ? parent.document.body.clientHeight : document.body.clientHeight;
			var sScrollTop = parent ? parent.document.body.scrollTop : document.body.scrollTop;
			var sTop = -80 + (sClientHeight / 2 + sScrollTop) - (document.getElementById(_sId).offsetHeight / 2);
			//document.getElementById(_sId).style.top = sTop > 0 ? sTop : (sClientHeight / 2 + sScrollTop) - (document.getElementById(_sId).offsetHeight / 2)+"px";
			document.getElementById(_sId).style.top = (theHeight / 2 + document.body.scrollTop) - (document.getElementById(_sId).offsetHeight / 2)+"px";
		}
	}
		
}
BtnOver=function(obj,path){obj.style.backgroundImage = "url("+path+"button3.bmp)";}
BtnOut=function(obj,path){	obj.style.backgroundImage = "url("+path+"button4.bmp)";}

function getBodySize(){
    var bodySize = [];
    with(document.documentElement) {
         bodySize[0] = (scrollWidth>clientWidth)?scrollWidth:clientWidth;
         bodySize[1] = (scrollHeight>clientHeight)?scrollHeight:clientHeight;
    }
       return bodySize;
}


function popCoverDiv(){
   if ($$("cover_div")) {
    $$("cover_div").style.display = '';
   } else {
    var coverDiv = document.createElement('div');
    document.body.appendChild(coverDiv);
    coverDiv.id = 'cover_div';
    
    with(coverDiv.style) {
     position = 'absolute';
     background = '#999999';
     left = '0px';
     top = '0px';
     var bodySize = getBodySize();
     width = bodySize[0] + 'px'
     height = bodySize[1] + 'px';
     zIndex = 98;
     if (IsIE) {
      filter = "Alpha(Opacity=60)";
     } else {
      opacity = 0.6;
     }
    }
   }
}
