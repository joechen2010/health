﻿// JScript 文件
function NewsDivChange(id,idx,allNum)
{
    style1="news_menu1";
    style2="news_menu2";
    href_style1="link_black";
    href_style2="link_white";
    for(i=1;i<=allNum;i++)
    {
        if(document.getElementById("News"+id+"_title_"+i)!=null)
        {
            document.getElementById("News"+id+"_title_"+i).className=style2;
        }
        if(document.getElementById("News"+id+"_div_"+i)!=null)
        {
            document.getElementById("News"+id+"_div_"+i).style.display="none";
        }
        if(document.getElementById("News"+id+"_href_"+i)!=null)
        {
            document.getElementById("News"+id+"_href_"+i).className=href_style2;
        }    
    }
        if(document.getElementById("News"+id+"_title_"+idx)!=null)
        {
            document.getElementById("News"+id+"_title_"+idx).className=style1;
        }
        if(document.getElementById("News"+id+"_div_"+idx)!=null)
        {
            document.getElementById("News"+id+"_div_"+idx).style.display="block";
        }
        if(document.getElementById("News"+id+"_href_"+idx)!=null)
        {
            document.getElementById("News"+id+"_href_"+idx).className=href_style1;
        }
    
}

function QuoteDivChange(id,idx,allNum)
{
    style1="li_now";
    style2="";
    for(i=1;i<=allNum;i++)
    {
        if(document.getElementById("Quote"+id+"_title_"+i)!=null)
        {
            document.getElementById("Quote"+id+"_title_"+i).className=style2;
        }
        if(document.getElementById("Quote"+id+"_div_"+i+"_1")!=null)
        {
            document.getElementById("Quote"+id+"_div_"+i+"_1").style.display="none";
        }
        if(document.getElementById("Quote"+id+"_div_"+i+"_2")!=null)
        {
            document.getElementById("Quote"+id+"_div_"+i+"_2").style.display="none";
        }
        if(document.getElementById("Quote"+id+"_div_"+i+"_3")!=null)
        {
            document.getElementById("Quote"+id+"_div_"+i+"_3").style.display="none";
        }
        if(document.getElementById("Quote"+id+"_div_"+i+"_4")!=null)
        {
            document.getElementById("Quote"+id+"_div_"+i+"_4").style.display="none";
        }
    }
        if(document.getElementById("Quote"+id+"_title_"+idx)!=null)
        {
            document.getElementById("Quote"+id+"_title_"+idx).className=style1;
        }
        if(document.getElementById("Quote"+id+"_div_"+idx+"_1")!=null)
        {
            document.getElementById("Quote"+id+"_div_"+idx+"_1").style.display="block";
        }
        if(document.getElementById("Quote"+id+"_div_"+idx+"_2")!=null)
        {
            document.getElementById("Quote"+id+"_div_"+idx+"_2").style.display="block";
        }
        if(document.getElementById("Quote"+id+"_div_"+idx+"_3")!=null)
        {
            document.getElementById("Quote"+id+"_div_"+idx+"_3").style.display="block";
        }
        if(document.getElementById("Quote"+id+"_div_"+idx+"_4")!=null)
        {
            document.getElementById("Quote"+id+"_div_"+idx+"_4").style.display="block";
        }
    
}