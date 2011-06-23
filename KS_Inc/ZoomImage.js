// JScript File
var ZoomImage_ImageHeight = 313;
var ZoomImage_ImageWidth = 250;
var ZoomImage_iDivHeight = 250;
var ZoomImage_iDivWidth = 250;
var ZoomImage_iDragHeight;
var ZoomImage_iDragWidth;
var ZoomImage_iMultiple;
var ZoomImage_iPrevImgWidth = 800;
var ZoomImage_iPrevImgHeight = 1000;
var ZoomImage_imageurl = '';
var ZoomImage_bigimageurl = '';
var ZoomImage_nedref = false;
var ZoomImage_img;
var ZoomImage_Pic_bg;
var ZoomImage_Drag_Box;
var ZoomImage_Prev_Win;
var ZoomImage_isFF = false;
function ZoomImage_ShowPicBG(ZoomImage_img,bgid,dragid,previd,imgheight,imgwidth,divheight,divwidth,prevheight,prevwidth,event)
{    
    var sAgent = navigator.userAgent.toLowerCase();    
    ZoomImage_isFF = (sAgent.indexOf("firefox")!=-1);
    
    ZoomImage_Pic_bg = bgid; 
    ZoomImage_Drag_Box = dragid;    
    ZoomImage_Prev_Win = previd;
    ZoomImage_ImageHeight = imgheight;
    ZoomImage_ImageWidth = imgwidth;
    ZoomImage_iDivHeight = divheight;
    ZoomImage_iDivWidth = divwidth;
    ZoomImage_iPrevImgWidth = prevwidth;
    ZoomImage_iPrevImgHeight = prevheight;
    if (ZoomImage_Pic_bg.style.visibility != "visible")
    {         
        ZoomImage_Pic_bg.style.width = ZoomImage_ImageWidth + "px";
        ZoomImage_Pic_bg.style.height = ZoomImage_ImageHeight + "px";
    }
    ZoomImage_Pic_bg.style.visibility = "visible";
    ZoomImage_Drag_Box.style.visibility = "visible";
    ZoomImage_Prev_Win.style.visibility = "visible";        
    if(ZoomImage_imageurl != ZoomImage_img.src)
    {
        ZoomImage_nedref = true;
        ZoomImage_imageurl = ZoomImage_img.src;          
//        if (ZoomImage_isFF) 
//            ZoomImage_bigimageurl = ZoomImage_imageurl.replace("4.jpg","6.jpg");        
//        else
            ZoomImage_bigimageurl = ZoomImage_img.getAttribute("bigimg");                            
    }    
    ZoomImage_DragImg(ZoomImage_Pic_bg,event,ZoomImage_nedref);         
}
String.prototype.trim = function()
{
var reExtraSpace = /^\s*(.*?)\s+$/;
return this.replace(reExtraSpace,"$1");
}

function ZoomImage_DragImg(src,event,refs)
{                               		
	var iPosX, iPosY;
	var idragX, idragY;	
	var sAgent = navigator.userAgent.toLowerCase();    
    var ZoomImage_isFF = (sAgent.indexOf("firefox")!=-1);//firefox  
    var iMouseX;
    var iMouseY;
    if (ZoomImage_isFF)
    {
        iMouseX = event.layerX;
        iMouseY = event.layerY;
    }
    else
    {
	    iMouseX = window.event.offsetX;
	    iMouseY = window.event.offsetY;		
	}
	ZoomImage_iMultiple = ZoomImage_iPrevImgWidth / ZoomImage_ImageWidth;
	ZoomImage_iDragHeight = ZoomImage_iDivHeight / ZoomImage_iMultiple;
	ZoomImage_iDragWidth = ZoomImage_iDivWidth / ZoomImage_iMultiple;
	if (iMouseX <= (ZoomImage_iDragWidth /2))
	{
	    idragX = 0;
	}
	else
	{
	    if ((ZoomImage_ImageWidth - iMouseX) <= (ZoomImage_iDragWidth / 2))
	        idragX = -(ZoomImage_ImageWidth - ZoomImage_iDragWidth);
	    else
	        idragX = -(iMouseX - ZoomImage_iDragWidth / 2);
	}
	
	if (iMouseY <= (ZoomImage_iDragHeight /2))
	{
	    idragY = 0;
	}
	else
	{
	    if ((ZoomImage_ImageHeight - iMouseY) <= (ZoomImage_iDragHeight / 2))
	        idragY = -(ZoomImage_ImageHeight - ZoomImage_iDragHeight);
	    else
	        idragY = -(iMouseY - ZoomImage_iDragHeight / 2);
	}
	iPosY = idragY * ZoomImage_iMultiple;
	iPosX = idragX * ZoomImage_iMultiple;
	
	if (ZoomImage_iPrevImgHeight + iPosY < ZoomImage_iDivHeight)
	    iPosY = ZoomImage_iDivHeight - ZoomImage_iPrevImgHeight;
	if (ZoomImage_iPrevImgWidth + iPosX < ZoomImage_iDivWidth)
	    iPosX = ZoomImage_iDivWidth - ZoomImage_iPrevImgWidth;
	
	var top;
	if (iMouseY - ZoomImage_iDragHeight / 2 <= 0)
	    top = 0;
	else if (ZoomImage_ImageHeight - iMouseY <= ZoomImage_iDragHeight / 2)
	    top = ZoomImage_ImageHeight - ZoomImage_iDragHeight;
	else
	    top = iMouseY - ZoomImage_iDragHeight / 2;
	    	
	var left;
	if (iMouseX - ZoomImage_iDragWidth / 2 <= 0)
	    left = 0;
	else if (ZoomImage_ImageWidth - iMouseX <= ZoomImage_iDragWidth / 2)
	    left = ZoomImage_ImageWidth - ZoomImage_iDragWidth;
	else
	    left = iMouseX - ZoomImage_iDragWidth / 2;	   	
	
	ZoomImage_Drag_Box.style.top = top + "px";
	ZoomImage_Drag_Box.style.left = left + "px";	
	
	if (ZoomImage_Drag_Box.innerHTML.toString().trim() == "" || refs)
	{	    
	    ZoomImage_Drag_Box.style.width = ZoomImage_iDragWidth + "px";
	    ZoomImage_Drag_Box.style.height = ZoomImage_iDragHeight + "px";
		ZoomImage_Drag_Box.innerHTML = "<img id='BigImg' style='position:absolute'>";		
		BigImg = document.getElementById("BigImg");
		BigImg.src = ZoomImage_imageurl;
		BigImg.style.height = ZoomImage_ImageHeight + "px";
		BigImg.style.width = ZoomImage_ImageWidth + "px";
		ZoomImage_nedref = false;
	}
	else
	    BigImg = document.getElementById("BigImg");	 	
	    
    var DivImg = document.getElementById("Div_Img");        	    	
    var PrevImg = document.getElementById("PrevImg");
    //alert(DivImg.offsetTop + "/" + ZoomImage_ImageHeight + "/" + ZoomImage_iDivHeight);
    
	if (refs || ZoomImage_Prev_Win.innerHTML.toString().trim() == "")
	{	    	    
	    ZoomImage_Prev_Win.style.width = ZoomImage_iDivWidth + "px";
	    ZoomImage_Prev_Win.style.height = ZoomImage_iDivHeight + "px";	    	    	    
	    if (ZoomImage_ImageHeight >= 313)
	        ZoomImage_Prev_Win.style.top = String(ZoomImage_ImageHeight - ZoomImage_iDivHeight) + "px";
	    else
	        ZoomImage_Prev_Win.style.top = String((ZoomImage_ImageHeight - ZoomImage_iDivHeight) / 2) + "px";	    
	    ZoomImage_Prev_Win.style.left = String(ZoomImage_ImageWidth) + "px";	    	    	    	    
	    ZoomImage_Prev_Win.innerHTML = "<img id='PrevImg' style='position:absolute;'/>";
	    PrevImg = document.getElementById("PrevImg");	    
	    PrevImg.src = ZoomImage_bigimageurl;	    
	    PrevImg.style.width = ZoomImage_iPrevImgWidth + "px";
	    PrevImg.style.height = ZoomImage_iPrevImgHeight + "px";
	    ZoomImage_nedref = false;
	}		
	else
	    PrevImg = document.getElementById("PrevImg");	
		    	    		    
	BigImg.style.top = idragY + "px";
    BigImg.style.left = idragX + "px";
	PrevImg.style.top = iPosY + "px";
	PrevImg.style.left = iPosX + "px";	
}
function ZoomImage_HidePicBG(event)
{        
    var x;
    var y;
    if (ZoomImage_isFF)
    {
        x = event.layerX;
        y = event.layerY;
    }
    else
    {
	    x = window.event.offsetX;
	    y = window.event.offsetY;
	}        	
    if (x <= 0 || x >= ZoomImage_ImageWidth || y <= 0 || y >= ZoomImage_ImageHeight)
    {
        ZoomImage_Pic_bg.style.visibility = "hidden";        
        ZoomImage_Drag_Box.style.visibility = "hidden";        
        ZoomImage_Prev_Win.style.visibility = "hidden";        
    }
}
