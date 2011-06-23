    var speedControl = 4000; 
    var agt=navigator.userAgent.toLowerCase(); 
  
    var is_major = parseInt(navigator.appVersion); 
    var is_minor = parseFloat(navigator.appVersion); 

   var is_nav  = ((agt.indexOf('mozilla')!=-1) && (agt.indexOf('spoofer')==-1) 
                && (agt.indexOf('compatible') == -1) && (agt.indexOf('opera')==-1) 
                && (agt.indexOf('webtv')==-1) && (agt.indexOf('hotjava')==-1)); 
    var is_nav2 = (is_nav && (is_major == 2)); 
    var is_nav3 = (is_nav && (is_major == 3)); 
    var is_nav4 = (is_nav && (is_major == 4)); 
    var is_nav4up = (is_nav && (is_major >= 4)); 
    var is_navonly      = (is_nav && ((agt.indexOf(";nav") != -1) || 
                          (agt.indexOf("; nav") != -1)) ); 
    var is_nav6 = (is_nav && (is_major == 5)); 
    var is_nav6up = (is_nav && (is_major >= 5)); 
    var is_gecko = (agt.indexOf('gecko') != -1); 
  
  
    var is_ie     = ((agt.indexOf("msie") != -1) && (agt.indexOf("opera") == -1)); 
    var is_ie3    = (is_ie && (is_major < 4)); 
    var is_ie4    = (is_ie && (is_major == 4) && (agt.indexOf("msie 4")!=-1) ); 
    var is_ie4up  = (is_ie && (is_major >= 4)); 
    var is_ie5    = (is_ie && (is_major == 4) && (agt.indexOf("msie 5.0")!=-1) ); 
    var is_ie5_5  = (is_ie && (is_major == 4) && (agt.indexOf("msie 5.5") !=-1)); 
    var is_ie5up  = (is_ie && !is_ie3 && !is_ie4); 
    var is_ie5_5up =(is_ie && !is_ie3 && !is_ie4 && !is_ie5); 
    var is_ie6    = (is_ie && (is_major == 4) && (agt.indexOf("msie 6.")!=-1) ); 
    var is_ie6up  = (is_ie && !is_ie3 && !is_ie4 && !is_ie5 && !is_ie5_5); 
    var is_aol   = (agt.indexOf("aol") != -1); 
    var is_aol3  = (is_aol && is_ie3); 
    var is_aol4  = (is_aol && is_ie4); 
    var is_aol5  = (agt.indexOf("aol 5") != -1); 
    var is_aol6  = (agt.indexOf("aol 6") != -1); 
  
    var is_opera = (agt.indexOf("opera") != -1); 
    var is_opera2 = (agt.indexOf("opera 2") != -1 || agt.indexOf("opera/2") != -1); 
    var is_opera3 = (agt.indexOf("opera 3") != -1 || agt.indexOf("opera/3") != -1); 
    var is_opera4 = (agt.indexOf("opera 4") != -1 || agt.indexOf("opera/4") != -1); 
    var is_opera5 = (agt.indexOf("opera 5") != -1 || agt.indexOf("opera/5") != -1); 
    var is_opera5up = (is_opera && !is_opera2 && !is_opera3 && !is_opera4); 
  
    var is_webtv = (agt.indexOf("webtv") != -1); 
  
    var is_TVNavigator = ((agt.indexOf("navio") != -1) || (agt.indexOf("navio_aoltv") != -1)); 
    var is_AOLTV = is_TVNavigator; 
  
    var is_hotjava = (agt.indexOf("hotjava") != -1); 
    var is_hotjava3 = (is_hotjava && (is_major == 3)); 
    var is_hotjava3up = (is_hotjava && (is_major >= 3)); 
  
    // *** JAVASCRIPT VERSION CHECK *** 
    var is_js; 
    if (is_nav2 || is_ie3) is_js = 1.0; 
    else if (is_nav3) is_js = 1.1; 
    else if (is_opera5up) is_js = 1.3; 
    else if (is_opera) is_js = 1.1; 
    else if ((is_nav4 && (is_minor <= 4.05)) || is_ie4) is_js = 1.2; 
    else if ((is_nav4 && (is_minor > 4.05)) || is_ie5) is_js = 1.3; 
    else if (is_hotjava3up) is_js = 1.4; 
    else if (is_nav6 || is_gecko) is_js = 1.5; 
    else if (is_nav6up) is_js = 1.5; 
    else if (is_ie5up) is_js = 1.3 
    else is_js = 0.0; 
  
    // *** PLATFORM *** 
    var is_win   = ( (agt.indexOf("win")!=-1) || (agt.indexOf("16bit")!=-1) ); 
    var is_win95 = ((agt.indexOf("win95")!=-1) || (agt.indexOf("windows 95")!=-1)); 
  
    // is this a 16 bit compiled version? 
    var is_win16 = ((agt.indexOf("win16")!=-1) || 
               (agt.indexOf("16bit")!=-1) || (agt.indexOf("windows 3.1")!=-1) || 
               (agt.indexOf("windows 16-bit")!=-1) );   
  
    var is_win31 = ((agt.indexOf("windows 3.1")!=-1) || (agt.indexOf("win16")!=-1) || 
                    (agt.indexOf("windows 16-bit")!=-1)); 
  
    var is_winme = ((agt.indexOf("win 9x 4.90")!=-1)); 
    var is_win2k = ((agt.indexOf("windows nt 5.0")!=-1)); 
    var is_win98 = ((agt.indexOf("win98")!=-1) || (agt.indexOf("windows 98")!=-1)); 
    var is_winnt = ((agt.indexOf("winnt")!=-1) || (agt.indexOf("windows nt")!=-1)); 
    var is_win32 = (is_win95 || is_winnt || is_win98 || 
                    ((is_major >= 4) && (navigator.platform == "Win32")) || 
                    (agt.indexOf("win32")!=-1) || (agt.indexOf("32bit")!=-1)); 
  
    var is_os2   = ((agt.indexOf("os/2")!=-1) || 
                    (navigator.appVersion.indexOf("OS/2")!=-1) ||   
                    (agt.indexOf("ibm-webexplorer")!=-1)); 
  
    var is_mac    = (agt.indexOf("mac")!=-1); 
    // hack ie5 js version for mac 
    if (is_mac && is_ie5up) is_js = 1.4; 
    var is_mac68k = (is_mac && ((agt.indexOf("68k")!=-1) || 
                               (agt.indexOf("68000")!=-1))); 
    var is_macppc = (is_mac && ((agt.indexOf("ppc")!=-1) || 
                                (agt.indexOf("powerpc")!=-1))); 
  
    var is_sun   = (agt.indexOf("sunos")!=-1); 
    var is_sun4  = (agt.indexOf("sunos 4")!=-1); 
    var is_sun5  = (agt.indexOf("sunos 5")!=-1); 
    var is_suni86= (is_sun && (agt.indexOf("i86")!=-1)); 
    var is_irix  = (agt.indexOf("irix") !=-1);    // SGI 
    var is_irix5 = (agt.indexOf("irix 5") !=-1); 
    var is_irix6 = ((agt.indexOf("irix 6") !=-1) || (agt.indexOf("irix6") !=-1)); 
    var is_hpux  = (agt.indexOf("hp-ux")!=-1); 
    var is_hpux9 = (is_hpux && (agt.indexOf("09.")!=-1)); 
    var is_hpux10= (is_hpux && (agt.indexOf("10.")!=-1)); 
    var is_aix   = (agt.indexOf("aix") !=-1);      // IBM 
    var is_aix1  = (agt.indexOf("aix 1") !=-1);     
    var is_aix2  = (agt.indexOf("aix 2") !=-1);     
    var is_aix3  = (agt.indexOf("aix 3") !=-1);     
    var is_aix4  = (agt.indexOf("aix 4") !=-1);     
    var is_linux = (agt.indexOf("inux")!=-1); 
    var is_sco   = (agt.indexOf("sco")!=-1) || (agt.indexOf("unix_sv")!=-1); 
    var is_unixware = (agt.indexOf("unix_system_v")!=-1); 
    var is_mpras    = (agt.indexOf("ncr")!=-1); 
    var is_reliant  = (agt.indexOf("reliantunix")!=-1); 
    var is_dec   = ((agt.indexOf("dec")!=-1) || (agt.indexOf("osf1")!=-1) || 
           (agt.indexOf("dec_alpha")!=-1) || (agt.indexOf("alphaserver")!=-1) || 
           (agt.indexOf("ultrix")!=-1) || (agt.indexOf("alphastation")!=-1)); 
    var is_sinix = (agt.indexOf("sinix")!=-1); 
    var is_freebsd = (agt.indexOf("freebsd")!=-1); 
    var is_bsd = (agt.indexOf("bsd")!=-1); 
    var is_unix  = ((agt.indexOf("x11")!=-1) || is_sun || is_irix || is_hpux || 
                 is_sco ||is_unixware || is_mpras || is_reliant || 
                 is_dec || is_sinix || is_aix || is_linux || is_bsd || is_freebsd); 
  
    var is_vms   = ((agt.indexOf("vax")!=-1) || (agt.indexOf("openvms")!=-1)); 
 // - IE5.5 and up can do the blending transition. 
var browserCanBlend = (is_ie5_5up); 
 /* show the blend select if appropriate */ 
  
var curImg = 0; 
var timerId = -1; 
var intervalf = 3000; 
var intervalm = 4500; 
var intervals = 6000; 
var interval = intervalm; 
var imgIsLoaded = false; 
  
var current_transition = 15; 
var flag = true; 
var bFirst = false; 
var transitions = new Array; 
transitions[0] = "progid:DXImageTransform.Microsoft.Fade(duration=1)"; 
transitions[1] = "progid:DXImageTransform.Microsoft.Blinds(Duration=1,bands=20)"; 
transitions[2] = "progid:DXImageTransform.Microsoft.Checkerboard(Duration=1,squaresX=20,squaresY=20)"; 
transitions[3] = "progid:DXImageTransform.Microsoft.Strips(Duration=1,motion=rightdown)"; 
transitions[4] = "progid:DXImageTransform.Microsoft.Barn(Duration=1,orientation=vertical)"; 
transitions[5] = "progid:DXImageTransform.Microsoft.GradientWipe(duration=1)"; 
transitions[6] = "progid:DXImageTransform.Microsoft.Iris(Duration=1,motion=out)"; 
transitions[7] = "progid:DXImageTransform.Microsoft.Wheel(Duration=1,spokes=12)"; 
transitions[8] = "progid:DXImageTransform.Microsoft.Pixelate(maxSquare=10,duration=1)"; 
transitions[9] = "progid:DXImageTransform.Microsoft.RadialWipe(Duration=1,wipeStyle=clock)"; 
transitions[10] = "progid:DXImageTransform.Microsoft.RandomBars(Duration=1,orientation=vertical)"; 
transitions[11] = "progid:DXImageTransform.Microsoft.Slide(Duration=1,slideStyle=push)"; 
transitions[12] = "progid:DXImageTransform.Microsoft.RandomDissolve(Duration=1,orientation=vertical)"; 
transitions[13] = "progid:DXImageTransform.Microsoft.Spiral(Duration=1,gridSizeX=40,gridSizeY=40)"; 
transitions[14] = "progid:DXImageTransform.Microsoft.Stretch(Duration=1,stretchStyle=push)"; 
transitions[15] = "special case"; 
var transition_count = 15; 
  
  
var arrPreload = new Array(); 
 function document.onreadystatechange() 
  { 
    preloadRange(0,_PRELOADRANGE-1); 
    window.onResize = function() 
    { 
        if (isMinNS4) { 
            var urlstr = ""; 
            window.open(urlstr, "_top"); 
        } 
    } 
    curImg = begImg; 
    if (curImg < 0 || curImg > numImgs - 1) 
    curImg = numImgs - 1; 
    changeSlide(); 
    interval = intervalm; 
    setTimeout("play()", 4000) 
  } 
  
function replaceNum(myinput, token, newstr) 
{ 
    var input = myinput; 
    var output = input; 
    var idx = output.indexOf(token); 
    if (idx > -1) 
    { 
        output = input.substring(0, idx); 
        output += newstr; 
        output += input.substr(idx+token.length); 
    } 
    return output; 
} 
  
function changeSpeed(sidx) 
{ 
     switch (sidx) 
    { 
        case 0: interval = 2000; break; 
        case 1: interval = 4000; break; 
        case 2: interval = 6000; break; 
        default: interval = 4000; 
    } 
    if (timerId != -1) 
    { 
        window.clearInterval(timerId); 
        timerId = window.setInterval("forward();", interval); 
    } 
} 
  
function preloadRange(intPic,intRange) { 
    var divStr = ""; 
    var iRange = 8; 
     
    for (var i=intPic; i<intPic+intRange; i++) { 
        arrPreload[i] = new Image(); 
		arrPreload[i].src = ImageSrcArray[i];     

    } 
     
    var strIn = ""; 
    var k = 0; 
    for (i=intPic; i<intPic+iRange; i++) { 
        j = i; 
        var strA = ""; 
        var s = ""; 
        var nextImg = curImg; 
        if(j>=numImgs || j<0)     
        { 
             
        } 
        else 
        { 
          strA = "<a href='javascript:changeSlide("+j+")'><img width=80 height=65 "+" id="+j+ GetBorder(j,nextImg)+" src='"+ImageSrcArray[j]+"' "+s+ " ></a>";         
        } 
         
        if(bFirst) 
        { 
            nextImg = curImg+1; 
            if(curImg==(numImgs-1)) 
            { 
                nextImg = 0; 
            } 
            if(curImg+1>numImgs) 
            { 
                nextImg = nextImg-numImgs; 
            } 
        } 
        strIn = "<span style='WIDTH: 75px; HEIGHT: 75px'>"+strA+"</span>&nbsp;&nbsp;"; 
        //strIn = strIn.replace("/b_","/s_"); 
        divStr += strIn; 
    } 
    document.getElementById("div1").innerHTML = divStr; 
    if(!bFirst) 
    { 
        if (browserCanBlend) 
        { 
            var divTrans = '&nbsp;选择效果：&nbsp;<select name="transitionType" size=1  onchange="change_transition()" ><option value=15 selected>随机显示</option> <option value=0 >混合效果</option> <option value=1 >百叶窗</option> <option value=2 >棋盘效果</option> <option value=3 >斜式效果</option> <option value=4 >开门见山</option> <option value=5 >擦除效果</option> <option value=6 >虹式效果</option> <option value=7 >风车效果</option> <option value=8 >波纹荡漾</option> <option value=9 >时钟转动</option> <option value=10 >雨丝效果</option> <option value=11 >滑动方式</option> <option value=12 >雪花显示</option> <option value=13 >盘旋收缩</option> <option value=14 >伸展方式</option> </select> '; 
            document.getElementById("div2").innerHTML = divTrans; 
        } 
        bFirst = true; 
    } 
     
    return true; 
} 
  
function GetBorder(i,j) 
{ 
    if(i==j) 
        return "  class=uimgSlide "; 
    return " border=0 "; 
  
} 
  
function ScaleImage(i) 
{ 
    // calculate the remaining avaialble width and height 
    var iAvailableWidth = 60;//span.clientWidth;; 
    var iAvailableHeight = 60;//span.clientHeight; 
    // calculate the scale to use for the image. Note, the image 
    // will never be stretched beyond it's actual width and height, 
    // and thus the scale max is 1 
    //iFullImageWidth = imageWidthArray[i]; 
    //iFullImageHeight = imageHeightArray[i]; 
    var flScale = Math.min(    1,Math.min((iAvailableWidth /iFullImageWidth),(iAvailableHeight /iFullImageHeight))); 
    // set the width of the image 
    var width = iScaledImageWidth = Math.round(iFullImageWidth * flScale); 
    // set the height of the image 
    var height = iScaledImageHeight = Math.round(iFullImageHeight * flScale); 
    // center the image by adjusting it's left and top values 
    var iLeft = Math.floor((iAvailableWidth - iScaledImageWidth) / 2); 
    var iTop = Math.floor((iAvailableHeight - iScaledImageHeight) / 2); 
    var style="style='HEIGHT: "+height+"; WIDTH: "+width+"'"; 
    return style; 
  
} 
  
function change_transition() 
{ 
    current_transition = document.all.transitionType.selectedIndex; 
    //alert(current_transition); 
    if(current_transition == 0) 
        current_transition = 15; 
    else 
        current_transition--; 
} 
  
function imgLoadNotify() 
{ 
    imgIsLoaded = true; 
} 
  
function GetChecked() 
{ 
    flag = !flag; 
} 
  
function changeSlide(n) 
{ 
    if (document.all) 
    {         
        /* transistion effects */ 
        if (browserCanBlend) 
        { 
            var do_transition; 
            //alert(current_transition); 
            if (current_transition == (transition_count)) 
            { 
                do_transition = Math.floor(Math.random() * transition_count); 
            } 
            else 
            { 
                do_transition = current_transition; 
            } 
            document.all.imgp.style.filter=transitions[do_transition]; 
            document.all.imgp.filters[0].Apply(); 
        } 
        else 
        { 
            document.all.imgp.style.filter="blendTrans(duration=1)"; 
            document.all.imgp.filters.blendTrans.Apply(); 
        } 
         
    } 
    if(n>=0) 
    { 
        curImg = n;         
        var s = document.getElementById("div1").innerHTML; 
        var sID = "id="+n; 
        var astrItems = s.split("&nbsp;"); 
        var iItemSep; 
        var str; 
        var strInner = ""; 
        for (var i=0; i < astrItems.length; i++) 
        { 
            str = astrItems[i]; 
             
            if(str.indexOf("class=uimgSlide") != -1) 
            { 
                if(str.indexOf(sID) == -1) 
                { 
                    str = str.replace("class=uimgSlide", "border=0"); 
                } 
            }     
            if(str.indexOf(sID) != -1) 
            { 
                str = str.replace("border=0", "class=uimgSlide"); 
            } 
            strInner += str+"&nbsp;"; 
        } 
        document.getElementById("div1").innerHTML = strInner; 
         
    } 
    imgIsLoaded = false; 
     var htmlCont = "<center>" + "<a href=\""+ImageSrcArray[curImg]+"\" target='_blank' title='点击查看图片详细信息'><img src=\"" + ImageSrcArray[curImg] + "\" alt=\"" + imageIDsArray[curImg] + "\""; 
     
   //if (imageWidthArray[curImg] > 0 && imageHeightArray[curImg] > 0) 
   // { 
   //     htmlCont += (imageHeightArray[curImg]>330)?" height=330":" width=" + imageWidthArray[curImg] + " height=" + imageHeightArray[curImg];         
   // }   
    htmlCont += " border=0 hspace=10 vspace=10   class=uimg  onload=\"imgLoadNotify();\">" + "</a>"; 
     
    var a = flag?("<br><span class=yphsmhdr>" + imageIDsArray[curImg] + "</span></center>"):""; 
    htmlCont +=a; 
    var pnumLine = "<span >"; 
    pnumLine += replaceNum(SHOWINGSTRING, "@ImageNum@", eval(curImg+1)); 
    pnumLine += "</span >"; 
    document.getElementById("pnumDiv").innerHTML = pnumLine; 
    document.getElementById("imgp").innerHTML = htmlCont; 
    if (document.all) 
    {             
        document.all.imgp.filters[0].Play();         
    } 
     
    preloadRange(curImg-4,_PRELOADRANGE); 
} 
  
function forward() 
{ 
    if ((curImg+1) >= numImgs) 
    {             
        curImg = 0; 
        changeSlide(); 
        stop(); 
        //if(isLastPage == "0") 
        //{ 
        //    if(confirm("播放下一组图片？")) 
        //    { 
        //        document.location.href=slideshow_end_url; 
        //    } 
       // }     
        return;         
    } 
     
    imgIsLoaded = false; 
    imgIsLoaded = preloadRange(curImg-4,_PRELOADRANGE); 
     
    if (!arrPreload[curImg+1]) 
    { 
        curImg++; 
         
    } 
    else 
    { 
        curImg++;         
    } 
    changeSlide(); 
} 
  
function rewind() 
{ 
    curImg--; 
    if (curImg < 0) 
    { 
        curImg = numImgs-1; 
        changeSlide(); 
    } 
    else 
        changeSlide(); 
} 
  
function stop() 
{ 
    window.clearInterval(timerId); 
    timerId = -1; 
    document.playbtn.src = buttonImgPfx + buttonOffArray[0]; 
    document.stopbtn.src = buttonImgPfx + buttonOnArray[1]; 
    imgIsLoaded = true; 
} 
  
function play() 
{ 
    interval = speedControl; 
     
    forward(); 
    if (timerId == -1) 
        timerId = window.setInterval('forward();', interval); 
  
    document.playbtn.src = buttonImgPfx + buttonOnArray[0]; 
    document.stopbtn.src = buttonImgPfx + buttonOffArray[1]; 
} 
  
function setButton(direction) 
{ 
    interval = speedControl; 
    if (timerId != -1) { window.clearInterval(timerId); timerId = window.setInterval("forward();", interval); } 
    imgIsLoaded = true; 
    if (direction == 0) 
    { 
        document.prevbtn.src = buttonImgPfx + buttonOnArray[2]; 
        window.setTimeout("document.prevbtn.src = buttonImgPfx + buttonOffArray[2];", 1500); 
    } 
    else 
    { 
        document.fowdbtn.src = buttonImgPfx + buttonOnArray[3]; 
        window.setTimeout("document.fowdbtn.src = buttonImgPfx + buttonOffArray[3];", 1500); 
    } 
} 