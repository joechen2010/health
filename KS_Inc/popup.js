/** 
 * 为兼容旧版本的浏览器增加 Array 的 push 方法，就是向数组末尾追加一个对象
 */
if(!Array.prototype.push){
	Array.prototype.push=function(){
		var startLength=this.length;
		for(var i=0;i<arguments.length;i++)
			this[startLength+i]=arguments[i];
		return this.length;
	};
};
/** 
 * 对函数G的参数进行处理,用的比较多的是模拟getElementById
 */
function G(){
	//定义一个数组用来保存G的参数
	var elements=new Array();
	//循环分析函数G的参数内容
	for(var i=0;i<arguments.length;i++){
		var element=arguments[i];
		//如果参数的类型为字符串则以字符串的值为Id获得对象
		if(typeof element=='string')
			element=document.getElementById(element);
		//如果参数的个数为1，则返回那个参数对象
		if(arguments.length==1)
			return element;
		elements.push(element);
	};
	return elements;
};
/** 
 * 此方法封装一个javascript函数对象，返回一个新函数对象，新函数对象的主体和原对象相同，但是bind()方法参数将被用作当前对象的对象。 
 * 新函数中的 this 引用被改变为参数提供的对象。 
 * 比如： 
 *    <input type="text" name="number1" id="number1" value="number1 value"> 
 *    <input type="text" name="number2" id="number2" value="number2 value">  
 *    <script> 
 *        var n1 = document.getElementById("number1"); 
 *        var n2 = document.getElementById("number2"); 
 *        n1.showValue1 = function() {alert(this.value);} 
 *        n2.showValue2 = n1.showValue.bind(n2); 
 *    </script> 
 * 那么，调用n1.showValue1 将打印"number1 value", 而调用n1.showValue2 将返回"number2 value"。 
 * 
 * apply 是ie5.5后才出现的新方法(Netscape很早就支持了),可构造回调函数，可用于ajax中 xmlHttpRequest的多个实例化。 
 * bind 的应用很广泛，该方法和 Object.prototype.extend 一样是 Prototype 的核心。 
 * 还有一个 call 方法，应用与 apply 类似。 
 */ 
Function.prototype.bind=function(object){
	var __method=this;
	return function(){
		__method.apply(object,arguments);
	};
};
/** 
 * 该方法与bind方法类似，只不过此方法一般用做html控件对象的事件处理。所以要传递event对象 
 * 用到了 Function.call，它与 Function.apply 的不同仅仅是对参数形式的定义。 
 */ 
Function.prototype.bindAsEventListener=function(object){
	var __method=this;
	return function(event){
		__method.call(object,event||window.event);
	};
};
/**
 *扩展一个javascript object
 */
Object.extend=function(destination,source){
	for(property in source){
		destination[property]=source[property];
	};
	return destination;
};
/**
 *为Object添加扩展属性。统一NetScape和Internet Explorer支持的属性命名
 */
if(!window.Event){
	var Event=new Object();
};
Object.extend(
	Event,
	{
	observers:false,
	element:function(event){
		return event.target||event.srcElement;
	},
	isLeftClick:function(event){
		return(((event.which)&&(event.which==1))||((event.button)&&(event.button==1)));
	},
	pointerX:function(event){
		return event.pageX||(event.clientX+(document.documentElement.scrollLeft||document.body.scrollLeft));
	},
	pointerY:function(event){
		return event.pageY||(event.clientY+(document.documentElement.scrollTop||document.body.scrollTop));
	},
	stop:function(event){
		if(event.preventDefault){
			event.preventDefault();
			event.stopPropagation();
		}
		else{
			event.returnValue=false;
			event.cancelBubble=true;
		};
	},
	findElement:function(event,tagName){
		var element=Event.element(event);
		while(element.parentNode&&(!element.tagName||(element.tagName.toUpperCase()!=tagName.toUpperCase())))
			element=element.parentNode;
		return element;
	},
	_observeAndCache:function(element,name,observer,useCapture){
		if(!this.observers)
			this.observers=[];
		if(element.addEventListener){
			this.observers.push([element,name,observer,useCapture]);
			element.addEventListener(name,observer,useCapture);
		}
		else if(element.attachEvent){
			this.observers.push([element,name,observer,useCapture]);
			element.attachEvent('on'+name,observer);
		};
	},
	unloadCache:function(){
		if(!Event.observers)
			return;
		for(var i=0;i<Event.observers.length;i++){
			Event.stopObserving.apply(this,Event.observers[i]);
			Event.observers[i][0]=null;
		};
		Event.observers=false;
	},
	observe:function(element,name,observer,useCapture){
		var element=G(element);
		useCapture=useCapture||false;
		if(name=='keypress'&&(navigator.appVersion.match(/Konqueror|Safari|KHTML/)||element.attachEvent))
			name='keydown';
		this._observeAndCache(element,name,observer,useCapture);
	},
	stopObserving:function(element,name,observer,useCapture){
		var element=G(element);
		useCapture=useCapture||false;
		if(name=='keypress'&&(navigator.appVersion.match(/Konqueror|Safari|KHTML/)||element.detachEvent))
			name='keydown';
		if(element.removeEventListener){
			element.removeEventListener(name,observer,useCapture);
		}
		else if(element.detachEvent){
			element.detachEvent('on'+name,observer);
		};
	}
});
/*
 *如果页面unload或者close，注册的event都要release
 */
Event.observe(window,'unload',Event.unloadCache,false);
/*
 *定义Class，使其能够具备extensibility
 */
var Class=function(){
	var _class=function(){
		this.initialize.apply(this,arguments);
	};
	for(i=0;i<arguments.length;i++){
		superClass=arguments[i];
		for(member in superClass.prototype){
			_class.prototype[member]=superClass.prototype[member];
		};
	};
	_class.child=function(){
		return new Class(this);
	};
	_class.extend=function(f){
		for(property in f){
			_class.prototype[property]=f[property];
		};
	};
	return _class;
};
/*
 *在baidu space中用于改变最顶端或者最低端的div的id值，比如版权区域
 *如果使用必须保证该Id(ft或者ft_popup)存在，并且相应的Css存在
 *如果flag为begin，则为弹出状态的id值
 *如果flag为end，则为非弹出状态的id值，即初始的id值
 */
function space(flag){
	if(flag=="begin"){
		var ele=document.getElementById("ft");
		if(typeof(ele)!="undefined"&&ele!=null)
			ele.id="ft_popup";
		ele=document.getElementById("usrbar");
		if(typeof(ele)!="undefined"&&ele!=null)
			ele.id="usrbar_popup";
	}
	else if(flag=="end"){
		var ele=document.getElementById("ft_popup");
		if(typeof(ele)!="undefined"&&ele!=null)
			ele.id="ft";
		ele=document.getElementById("usrbar_popup");
		if(typeof(ele)!="undefined"&&ele!=null)
			ele.id="usrbar";
	};
};
/*
 *弹出窗口类，主要用于改善User Experience
 */
var Popup=new Class();
Popup.prototype={
//弹出窗口中框架的name名称，一般使用iframe包含弹出窗口的内容 iframe popup
	iframeIdName:'ifr_popup',
	initialize:function(config){
/**
 *以下为弹出对话框的初始化配置信息，用于初始调用
 *contentType         : 设置Body区域的内容：1为另外一个html文件，往往用iframe | 2为自定义html字符串 | 3为confirm对话框 | 4为alert警告对话框
 *isHaveTitle         : 是否显示标题栏
 *scrollType          ：设置或获取弹出窗口的框架是否scrollable
 *isBackgroundCanClick：弹出窗口后，是否允许父窗口蒙布后的所有元素被点击。也就是如果为false的话，就会有全屏蒙布，如果为true的话，就会去掉全屏蒙布
 *isSupportDraging    ：是否支持拖拽
 *isShowShadow        ：是否显示阴影
 *isReloadOnClose     ：关闭弹出窗口后是否刷新父窗口
 *width               ：弹出窗口宽度
 *height              ：弹出窗口高度
 */		
		this.config=Object.extend(
			{
				contentType:1,isHaveTitle:true,scrollType:'no',isBackgroundCanClick:false,isSupportDraging:true,isShowShadow:true,isReloadOnClose:true,width:400,height:300
			},
			config||{}
			);
/**		
 *以下为弹出窗口的参数值信息，用于改变弹出窗口的外观
 *shadowWidth  ：阴影的宽度
 *contentUrl   ：html链接页面，即iframe的src
 *contentHtml  ：html内容，适用于contentType为2的情况
 *callBack     ：回调函数的名称
 *parameter    ：回调函数中传的参数
 *confirmCon   ：对话框内容，适用于contentType为3的情况
 *alertCon     ：警告框内容，适用于contentType为4的情况
 *someHiddenTag：页面中需要隐藏的元素列表，以逗号分割
 *someHiddenEle：需要隐藏的元素的ID列表(和someToHidden的区别是：someHiddenEle是通过getElementById，而someToHidden是通过getElementByTagName，里面放的是对象)
 *overlay      ：控制蒙布所处的层，z-index等
 *coverOpacity ：蒙布的透明值
*/	   
		this.info=
			{
				shadowWidth:4,title:"",contentUrl:"",contentHtml:"",callBack:null,parameter:null,confirmCon:"",alertCon:"",someHiddenTag:"select,object,embed",someDisabledBtn:"",someHiddenEle:"",overlay:0,coverOpacity:40
			};
/**			
 *设置颜色属性
 *cColor:蒙布的背景 coverage color
 *bColor:内容区域的背景 body backgroud color
 *tColor:标题栏和边框的颜色 title and border color
 *wColor:字体的背景色 font backgroud color
 */
		this.color=
			{
				cColor:"#EEEEEE",bColor:"#FFFFFF",tColor:"#709CD2",wColor:"#FFFFFF"
			};
		this.dropClass=null;
		//存储隐藏对象列表，在hiddenTag方法中第一次调用
		this.someToHidden=[];
		//存储禁用对象列表
		this.someToDisabled=[];
		//如果没有标题栏则不支持拖拽
		if(!this.config.isHaveTitle)
			this.config.isSupportDraging=false;
		this.iniBuild();
	},
	setContent:function(arrt,val){
		if(val!=''){
			switch(arrt){
				case 'width':
					this.config.width=val;
					break;
				case 'height':
					this.config.height=val;
					break;
				case 'title':
					this.info.title=val;
					break;
				case 'contentUrl':
					this.info.contentUrl=val;
					break;
				case 'contentHtml':
					this.info.contentHtml=val;
					break;
				case 'callBack':
					this.info.callBack=val;
					break;
				case 'parameter':
					this.info.parameter=val;
					break;
				case 'confirmCon':
					this.info.confirmCon=val;
					break;
				case 'alertCon':
					this.info.alertCon=val;
					break;
				case 'someHiddenTag':
					this.info.someHiddenTag=val;
					break;
				case 'someHiddenEle':
					this.info.someHiddenEle=val;
					break;
				case 'someDisabledBtn':
					this.info.someDisabledBtn=val;
					break;
				case 'overlay':
					this.info.overlay=val;
			};
		};
	},
	iniBuild:function(){
		G('dialogCase')?G('dialogCase').parentNode.removeChild(G('dialogCase')):function(){};
		var oDiv=document.createElement('span');
		oDiv.id='dialogCase';
		document.body.appendChild(oDiv);
	},
	build:function(){
		 //设置全屏蒙布的z-index
		var baseZIndex=10001+this.info.overlay*10;
		//设置蒙布上的弹出窗口的z-index(比蒙布的z-index高2个值)
		var showZIndex=baseZIndex+2;
		//定义框架名称,如果弹出多个窗口，通过设置overlay可以区分
		this.iframeIdName='ifr_popup'+this.info.overlay;
		//注意，这里可以设置弹出窗口用到的图标的路径
		var path="/images/";
		//渲染关闭按钮
		var close='<input type="image"  src="'+path+'dialogclose.gif" border="0" width="16" height="16" align="absmiddle" title="close"/>';
		//使用滤镜设置对象的透明度
		var cB='filter: alpha(opacity='+this.info.coverOpacity+');opacity:'+this.info.coverOpacity/100+';';
		//设置全屏的蒙布
		var cover='<div id="dialogBoxBG" style="position:absolute;top:0px;left:0px;width:100%;height:100%;z-index:'+baseZIndex+';'+cB+'background-color:'+this.color.cColor+';display:none;"></div>';
		//设置弹出的主窗口
		var mainBox='<div id="dialogBox" style="border:3px solid '+this.color.tColor+';display:none;z-index:'+showZIndex+';position:relative;width:'+this.config.width+'px;"><table width="100%" align="center" border="0" cellpadding="0" cellspacing="0" bgcolor="'+this.color.bColor+'">';
		//如果有标题栏，则设置窗口标题栏
		if(this.config.isHaveTitle){
			mainBox+='<tr height="24" bgcolor="'+this.color.tColor+'"><td><table style="-moz-user-select:none;height:24px;" width="100%" border="0" cellpadding="0" cellspacing="0" ><tr>'+'<td width="6" height="24"></td><td id="dialogBoxTitle" style="color:'+this.color.wColor+';font-size:14px;font-weight:bold;">'+this.info.title+'&nbsp;</td>'+'<td id="dialogClose" width="60" align="right" valign="middle"><span style="cursor:pointer" id="dialogBoxClose">'+close+'关闭</span></td><td width="6"></td></tr></table></td></tr>';
		}
		else{
			mainBox+='<tr height="10"><td align="right">'+close+'</td></tr>'};
			mainBox+='<tr style="height:'+this.config.height+'px" valign="top"><td id="dialogBody" style="position:relative;"></td></tr></table></div>'+'<div id="dialogBoxShadow" style="display:none;z-index:'+baseZIndex+';"></div>';
			//如果有蒙布，先渲染蒙布，再渲染蒙布上的弹出窗口
			if(!this.config.isBackgroundCanClick){
				G('dialogCase').innerHTML=cover+mainBox;
				G('dialogBoxBG').style.height=document.body.scrollHeight;
			}
			else{
				G('dialogCase').innerHTML=mainBox;
			}
			//注册关闭事件，以便关闭弹出窗口
			Event.observe(G('dialogBoxClose'),"click",this.reset.bindAsEventListener(this),false);
			//如果允许拖拽，设置拖拽
			if(this.config.isSupportDraging){
				dropClass=new Dragdrop(this.config.width,this.config.height,this.info.shadowWidth,this.config.isSupportDraging,this.config.contentType);
				G("dialogBoxTitle").style.cursor="move";
			};
			this.lastBuild();
		},
		lastBuild:function(){
			//设置confirm对话框的具体内容
			var confirm='<div style="width:100%;height:100%;text-align:center;"><div style="margin:20px 20px 0 20px;font-size:14px;line-height:16px;color:#000000;">'+this.info.confirmCon+'</div><div style="margin:20px;"><input id="dialogOk" type="button" value="  确认  "/>&nbsp;<input id="dialogCancel" type="button" value="  取消  "/></div></div>';
			//设置alert对话框的具体内容
			var alert='<div style="width:100%;height:100%;text-align:center;"><div style="margin:20px 20px 0 20px;font-size:14px;line-height:16px;color:#000000;">'+this.info.alertCon+'</div><div style="margin:20px;"><input id="dialogYES" type="button" value="  确认  "/></div></div>';
			//蒙布层
			var baseZIndex=10001+this.info.overlay*10;
			//弹出窗口层
			var coverIfZIndex=baseZIndex+4;
			//根据内容类型flag的值决定窗口的主内容区域应该显示的内容
			if(this.config.contentType==1){
				var openIframe="<iframe width='100%' style='height:"+this.config.height+"px' name='"+this.iframeIdName+"' id='"+this.iframeIdName+"' src='"+this.info.contentUrl+"' frameborder='0' scrolling='"+this.config.scrollType+"'></iframe>";
				var coverIframe="<div id='iframeBG' style='position:absolute;top:0px;left:0px;width:1px;height:1px;z-index:"+coverIfZIndex+";filter: alpha(opacity=00);opacity:0.00;background-color:#ffffff;'><div>";
				G("dialogBody").innerHTML=openIframe+coverIframe;
			}
			else if(this.config.contentType==2){
				G("dialogBody").innerHTML=this.info.contentHtml;
			}
			else if(this.config.contentType==3){
				G("dialogBody").innerHTML=confirm;
				Event.observe(G('dialogOk'),"click",this.forCallback.bindAsEventListener(this),false);
				Event.observe(G('dialogCancel'),"click",this.close.bindAsEventListener(this),false);
			}
			else if(this.config.contentType==4){
				G("dialogBody").innerHTML=alert;
				Event.observe(G('dialogYES'),"click",this.close.bindAsEventListener(this),false);
			};
		},
		//重新加载弹出窗口的高度和内容
		reBuild:function(){
			G('dialogBody').height=G('dialogBody').clientHeight;
			this.lastBuild();
		},
		show:function(){
			//隐藏指定的元素
			this.hiddenSome();
			//居中弹出窗口
			this.middle();
			//如果有阴影则渲染阴影
			if(this.config.isShowShadow)
				this.shadow();
		},
		//设置回调函数
		forCallback:function(){
			return this.info.callBack(this.info.parameter);
		},
		//设置弹出窗口的阴影
		shadow:function(){
			var oShadow=G('dialogBoxShadow');
			var oDialog=G('dialogBox');
			oShadow['style']['position']="absolute";
			oShadow['style']['background']="#000";
			oShadow['style']['display']="";
			oShadow['style']['opacity']="0.2";
			oShadow['style']['filter']="alpha(opacity=20)";
			oShadow['style']['top']=oDialog.offsetTop+this.info.shadowWidth;
			oShadow['style']['left']=oDialog.offsetLeft+this.info.shadowWidth;
			oShadow['style']['width']=oDialog.offsetWidth;
			oShadow['style']['height']=oDialog.offsetHeight;
		},
		//居中弹出窗口
		middle:function(){
			if(!this.config.isBackgroundCanClick)
				G('dialogBoxBG').style.display='';
			var oDialog=G('dialogBox');
			oDialog['style']['position']="absolute";
			oDialog['style']['display']='';
			var sClientWidth=document.body.clientWidth;
			var sClientHeight=document.body.clientHeight;
			var sScrollTop=document.body.scrollTop;
			var sleft=(document.body.clientWidth/2)-(oDialog.offsetWidth/2);
			var iTop=-40+(sClientHeight/2+sScrollTop)-(oDialog.offsetHeight/2);
			var sTop=iTop>0?iTop:(sClientHeight/2+sScrollTop)-(oDialog.offsetHeight/2);
			if(sTop<1)
				sTop="20";
			else if(sTop>300)
			   sTop=300;
			if(sleft<1)
				sleft="20";
			oDialog['style']['left']=sleft;
			oDialog['style']['top']=sTop;
			//oDialog['style']['top']=150;
		},
		//刷新父页面，并关闭当前弹出窗口
		reset:function(){
			if(this.config.isReloadOnClose){
				location.reload();
			};
			this.close();
		},
		//关闭弹出窗口
		close:function(){
			G('dialogBox').style.display='none';
			if(!this.config.isBackgroundCanClick)
				G('dialogBoxBG').style.display='none';
			if(this.config.isShowShadow)
				G('dialogBoxShadow').style.display='none';
			G('dialogBody').innerHTML='';
			this.showSome();
		},
		hiddenSome:function(){
			var tag=this.info.someHiddenTag.split(",");
			if(tag.length==1&&tag[0]=="")tag.length=0;
			for(var i=0;i<tag.length;i++){
				this.hiddenTag(tag[i]);
			};
			var ids=this.info.someHiddenEle.split(",");
			if(ids.length==1&&ids[0]=="")
				ids.length=0;
			for(var i=0;i<ids.length;i++){
				this.hiddenEle(ids[i]);
			};
			var ids=this.info.someDisabledBtn.split(",");
			if(ids.length==1&&ids[0]=="")
				ids.length=0;
			for(var i=0;i<ids.length;i++){
				this.disabledBtn(ids[i]);
			};
			space("begin");
		},
		disabledBtn:function(id){
			var ele=document.getElementById(id);
			if(typeof(ele)!="undefined"&&ele!=null&&ele.disabled==false){
				ele.disabled=true;
				this.someToDisabled.push(ele);
			};
		},
		hiddenTag:function(tagName){
			var ele=document.getElementsByTagName(tagName);
			if(ele!=null){
				for(var i=0;i<ele.length;i++){
					if(ele[i].style.display!="none"&&ele[i].style.visibility!='hidden'){
						ele[i].style.visibility='hidden';
						this.someToHidden.push(ele[i]);
					};
				};
			};
		},
		hiddenEle:function(id){
			var ele=document.getElementById(id);
			if(typeof(ele)!="undefined"&&ele!=null){
				ele.style.visibility='hidden';
				this.someToHidden.push(ele);
			};
		},
		showSome:function(){
			for(var i=0;i<this.someToHidden.length;i++){
				this.someToHidden[i].style.visibility='visible';
			};
			for(var i=0;i<this.someToDisabled.length;i++){
				this.someToDisabled[i].disabled=false;
			};
			space("end");
		}
};
/**
 *Dragdrop类,用于弹出窗口的拖拽动作
 */
var Dragdrop=new Class();
Dragdrop.prototype={
	initialize:function(width,height,shadowWidth,showShadow,contentType){
		this.dragData=null;
		this.dragDataIn=null;
		this.backData=null;
		this.width=width;
		this.height=height;
		this.shadowWidth=shadowWidth;
		this.showShadow=showShadow;
		this.contentType=contentType;
		this.IsDraging=false;
		this.oObj=G('dialogBox');
		Event.observe(G('dialogBoxTitle'),"mousedown",this.moveStart.bindAsEventListener(this),false);
	},
	moveStart:function(event){
		this.IsDraging=true;
		if(this.contentType==1){
			G("iframeBG").style.display="";
			G("iframeBG").style.width=this.width;
			G("iframeBG").style.height=this.height;
		};
		Event.observe(document,"mousemove",this.mousemove.bindAsEventListener(this),false);
		Event.observe(document,"mouseup",this.mouseup.bindAsEventListener(this),false);
		Event.observe(document,"selectstart",this.returnFalse,false);
		this.dragData={
			x:Event.pointerX(event),
			y:Event.pointerY(event)
		};
		this.backData={
			x:parseInt(this.oObj.style.left),
			y:parseInt(this.oObj.style.top)
		};
	},
	mousemove:function(event){
		if(!this.IsDraging)return;
		var iLeft=Event.pointerX(event)-this.dragData["x"]+parseInt(this.oObj.style.left);
		var iTop=Event.pointerY(event)-this.dragData["y"]+parseInt(this.oObj.style.top);
		if(this.dragData["y"]<parseInt(this.oObj.style.top))
			iTop=iTop-12;
		else if(this.dragData["y"]>parseInt(this.oObj.style.top)+25)
			iTop=iTop+12;
		this.oObj.style.left=iLeft;
		this.oObj.style.top=iTop;
		if(this.showShadow){
			G('dialogBoxShadow').style.left=iLeft+this.shadowWidth;
			G('dialogBoxShadow').style.top=iTop+this.shadowWidth;
		};
		this.dragData={
			x:Event.pointerX(event),
			y:Event.pointerY(event)
		};
		document.body.style.cursor="move";
	},
	mouseup:function(event){
		if(!this.IsDraging)
			return;
		if(this.contentType==1)
			G("iframeBG").style.display="none";
		document.onmousemove=null;
		document.onmouseup=null;
		var mousX=Event.pointerX(event)-(document.documentElement.scrollLeft||document.body.scrollLeft);
		var mousY=Event.pointerY(event)-(document.documentElement.scrollTop||document.body.scrollTop);
		if(mousX<1||mousY<1||mousX>document.body.clientWidth||mousY>document.body.clientHeight){
			this.oObj.style.left=this.backData["x"];
			this.oObj.style.top=this.backData["y"];
			if(this.showShadow){
				G('dialogBoxShadow').style.left=this.backData.x+this.shadowWidth;
				G('dialogBoxShadow').style.top=this.backData.y+this.shadowWidth;
			};
		};
		this.IsDraging=false;
		document.body.style.cursor="";
		Event.stopObserving(document,"selectstart",this.returnFalse,false);
	},
	returnFalse:function(){
		return false;
	}
};