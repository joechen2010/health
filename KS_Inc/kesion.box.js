var isIe=(document.all)?true:false;
var onscrolls=true;
var QuoteByAdmin=false;
var PopupImgDir='';
//����괦����������ʾ
function mousepopup(ev,title,content,width)
{
   var objPos = mousePosition(ev);
   showMessageBox(title,content,objPos,width);
}
//����괦����iframe��ʾ
function mousePopupIframe(ev,title,url,width,height,scrollType)
{ 
  var objPos = mousePosition(ev);
  var text="<iframe marginwidth='0' marginheight='0' width='99%' style='height:"+height+"px'  src='"+url+"' frameborder='0' scrolling='"+scrollType+"'></iframe>";
  showMessageBox(title,text,objPos,width);
}


//���е���������ʾ
function popupTips(title,content,width,height)
{  
   QuoteByAdmin=true;
   var objPos = Position(width,height);
   showMessageBox(title,content,objPos,width);
}

//��̨���е���
function PopupCenterIframe(title,url,width,height,scrollType)
{ QuoteByAdmin=true;
  onscrolls=false;
  var objPos = Position(width,height);
  var text="<iframe marginwidth='0' marginheight='0' width='99%' style='height:"+height+"px'  src='"+url+"' frameborder='0' scrolling='"+scrollType+"'></iframe>";
  showMessageBox(title,text,objPos,width);
}
function Position(width,height)
{
	if (isIe){
   return {x:document.body.offsetWidth/2-width/2, y:document.body.offsetHeight/2-height/2-20};
	}else{
   return {x:document.documentElement.scrollWidth/2-width/2, y:window.screen.height/2-height/2-150};
	}
	
}

//���г�ʼ������
//����:title ���� content���� width ���
function popupIframe(title,url,width,height,scrollType)
{ var objPos = middlePosition(width);
  var text="<iframe width='100%' style='height:"+height+"px'  src='"+url+"' frameborder='0' scrolling='"+scrollType+"'></iframe>";
   showMessageBox(title,text,objPos,width);
}
//���е�����ͨ����
//����:title ���� content���� width ���
function popup(title,content,width)
{ var objPos = middlePosition(width);
  showMessageBox(title,content,objPos,width);
}

//����select�Ŀɼ�״̬
function setSelectState(state)
{
   var objl=document.getElementsByTagName('select');
   for(var i=0;i<objl.length;i++)
    {
     objl[i].style.visibility=state;
    }
}
function mousePosition(ev)
{
   if(ev.pageX || ev.pageY)
   {
    return {x:ev.pageX, y:ev.pageY};
   }
  
   return {
    x:ev.clientX + document.body.scrollLeft - document.body.clientLeft,y:ev.clientY + document.body.scrollTop - document.body.clientTop
   };
}
function middlePosition(width)
{ 
    var left = parseInt((screen.availWidth/2+width/2));//��Ļ����
    var top = document.body.scrollTop;
	if (top==0) top=document.documentElement.scrollTop;
    top=top+200;
    return {x:left, y:top};
}
window.onscroll = function(){
	 if (!onscrolls) return;
	 try{
	 var top=document.body.scrollTop;
	 if (top==0) top=document.documentElement.scrollTop;
	 document.getElementById("mesWindow").style.top=top+200;}
	 catch(e){}
}
var Obj=''
document.onmouseup=MUp
document.onmousemove=MMove

function MDown(Object){
Obj=Object.id
document.all(Obj).setCapture()
pX=event.x-document.all(Obj).style.pixelLeft;
pY=event.y-document.all(Obj).style.pixelTop;
}

function MMove(){
if(Obj!=''){
document.all(Obj).style.left=event.x-pX;
document.all(Obj).style.top=event.y-pY;
}
}

function MUp(){
if(Obj!=''){
document.all(Obj).releaseCapture();
Obj='';
}
}
//��������
function showMessageBox(wTitle,content,pos,wWidth)
{
   closeWindow();
   
   var mesWindowCss="border:#666 4px solid;background:#fff;" //�����߿�
   if (QuoteByAdmin==true){
	   mesWindowCss="border:#000000 1px solid;background:#F1F6F9;" //�����߿�
	   }
   
   var bWidth=parseInt(document.documentElement.scrollWidth);
   var bHeight=parseInt(document.body.offsetHeight);
   bWidth=parseInt(document.body.scrollWidth);
  // if (bHeight<parseInt(document.body.scrollHeight)) bHeight=parseInt(document.body.scrollHeight);
   
   if(isIe){ setSelectState('hidden');}
   var back=document.createElement("div");
   back.id="back";
   var styleStr="top:0px;left:0px;position:absolute;background:#666;width:"+bWidth+"px;height:"+bHeight+"px;";
   styleStr+=(isIe)?"filter:alpha(opacity=8);":"opacity:0.10;";
   back.style.cssText=styleStr;
   document.body.appendChild(back);
   var mesW=document.createElement("div");
   mesW.id="mesWindow";
   //mesW.className="mesWindow";
   if (QuoteByAdmin==true){
	      mesW.innerHTML="<div style='border-bottom:#eee 1px solid;font-weight:bold;text-align:left;font-size:12px;'><table cellpadding='0' cellspacing='0' bgcolor='#CFE0EA' background='"+PopupImgDir+"images/menu_bg.gif' width='100%' height='100%'><tr onmousedown=MDown(mesWindow)><td align='center' height='28'><b>"+wTitle+"</b></td><td align='center' width='80'><span style='cursor:pointer' id='closeWindow' onclick='closeWindow();' title='�رմ���' class='close'><img src='"+PopupImgDir+"../images/default/close.gif' border='0'> �ر�</span> </td></tr></table></div><div style='_margin:4px;font-size:12px;' id='mesWindowContent'>"+content+"</div><div class='mesWindowBottom'></div>";

   }else{
	    mesW.innerHTML="<div style='border-bottom:#eee 1px solid;margin-left:4px;padding:3px;font-weight:bold;text-align:left;font-size:12px;'><table cellpadding='0' cellspacing='0' width='100%' height='100%'><tr><td><b>"+wTitle+"</b></td><td style='width:1px;'><input type='button' onclick='closeWindow();' title='�رմ���' class='close' value='�ر�' /></td></tr></table></div><div style='margin:4px;font-size:12px;' id='mesWindowContent'>"+content+"</div><div class='mesWindowBottom'></div>";

   }
   styleStr=mesWindowCss+"left:"+(((pos.x-wWidth)>0)?(pos.x-wWidth):pos.x)+"px;top:"+(pos.y)+"px;position:absolute;width:"+wWidth+"px;";
   mesW.style.cssText=styleStr;
   document.body.appendChild(mesW);
}

function showBackground(obj,endInt)
{
   obj.filters.alpha.opacity+=1;
   if(obj.filters.alpha.opacity<endInt)
   {
    setTimeout(function(){showBackground(obj,endInt)},8);
   }
}
//�رմ���
function closeWindow()
{
   if(document.getElementById('back')!=null)
   {
    document.getElementById('back').parentNode.removeChild(document.getElementById('back'));
   }
   if(document.getElementById('mesWindow')!=null)
   {
    document.getElementById('mesWindow').parentNode.removeChild(document.getElementById('mesWindow'));
   }
   if(isIe){ setSelectState('');}
}