///////// JS theme file for FlatCalendarXP 6.0 /////////
// This file is totally configurable. You may remove all the comments in this file to shrink the download size.
////////////////////////////////////////////////////////
var gMonths=["01","02","03","04","05","06","07","08","09","10","11","12"];
//var gMonths=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
var gWeekDay=["Su","Mo","Tu","We","Th","Fr","Sa"];	// weekday caption from Sunday to Saturday
//var gWeekDay=["日","一","二","三","四","五","六"];	// weekday caption from Sunday to Saturday
var gBegin=[1980,1,1];	// calendar date range begin from [Year,Month,Date]
var gEnd=[2030,12,31];	// calendar date range end at [Year,Month,Date]
var gsOutOfRange="Sorry, you may not go beyond the designated range!";	// out-of-date-range error message
var guOutOfRange=null;	// the background image url for the out-range dates.

var gbEuroCal=false;	// true: ISO-8601 calendar layout - Monday is the 1st day of week; false: US layout - Sunday is the 1st day of week.

var gcCalBG="#FFFFFF";	// the background color of the outer calendar panel.
var guCalBG=null;	//  the background image url for the inner table.
var gcCalFrame="white";	// the background color of the inner table, showing as a frame.
var gsInnerTable="border=0 cellpadding=2 cellspacing=1";	// properties of the inner <table> tag, which holds all the calendar cells.
var gsOuterTable=NN4?"border=1 cellpadding=3 cellspacing=0":"border=0 cellpadding=3 cellspacing=0";	// properties of the outmost container <table> tag, which holds the top, middle and bottom sections.

var gbHideTop=false;	// true: hide the top section; false: show it according to the following settings
var giDCStyle=0;	// the style of month-controls in top section.	0: 3D; 1: flat; 2: text-only;
var gsCalTitle="gMonths[gCurMonth[1]-1]+' '+gCurMonth[0]";	// dynamic statement to be eval-ed as the title when giDCStyle>0.
var gbDCSeq=false;	// (effective only when giDCStyle is 0) true: show month box before year box; false: vice-versa;
var gsYearInBox="i";	// dynamic statement to be eval-ed as the text shown in the year box. e.g. "'A.D.'+i" will show "A.D.2001"
var gsNavPrev="<INPUT type='button' value='&lt;' class='MonthNav' onclick='fPrevMonth();this.blur();'>";	// the content of the left month navigator
var gsNavNext="<INPUT type='button' value='&gt;' class='MonthNav' onclick='fNextMonth();this.blur();'>";	// the content of the right month navigator

var gbHideBottom=true;	// true: hide the bottom section; false: show it with gsBottom.
//var gsBottom="<A href='javascript:void(0)' class='Today' onclick='if(!NN4)this.blur();if(!fSetDate(gToday[0],gToday[1],gToday[2]))alert(\"You may not pick this day!\");return false;' onmouseover='return true;' title='Today'>Today : "+gToday[2]+" "+gMonths[gToday[1]-1]+" "+gToday[0]+"</A>";	// the content of the bottom section.
var gsBottom="<A href='xxx.Asp' class='Today' onclick='if(!NN4)this.blur();if(!fSetDate(gToday[0],gToday[1],gToday[2]))alert(\"You may not pick this day!\");return false;' onmouseover='return true;' title=''>Today : "+gToday[0]+" "+gMonths[gToday[1]-1]+" "+gToday[2]+"</A>";	// the content of the bottom section.
var giCellWidth=17;	// calendar cell width;
var giCellHeight=14;	// calendar cell height;
//var giHeadHeight=giCellHeight;
var giHeadHeight=10;	// calendar head row height;
var giWeekWidth=22;	// calendar week-number-column width;
var giHeadTop=0;	// calendar head row top offset;
var giWeekTop=0;	// calendar week-number-column top offset;

var gcCellBG="#e5e5e5";	// default background color of the cells. Use "" for transparent!!!
var gsCellHTML="";	// default HTML contents for days without any agenda, usually an image tag.
var guCellBGImg="";	// url of default background image for each calendar cell.
var gsAction=" ";	// default action to be eval-ed on everyday except the days with agendas, which have their own actions defined in agendas.
var gsDays="dayNo";	// the dynamic statement to be eval-ed into each day cell.

var giWeekCol=0;	// -1: disable week-number-column;  0~7: show week numbers at the designated column.
var gsWeekHead="#";	// the text shown in the table head of week-number-column.
var gsWeeks="weekNo";	// the dynamic statement to be eval-ed into the week-number-column. e.g. "'week '+weekNo" will show "week 1", "week 2" ...

var gcWorkday="black";	// Workday font color
var gcSat="black";	// Saturday font color
var gcSatBG="#dae6eb";	// Saturday background color
var gcSun="black";	// Sunday font color
var gcSunBG="#dae6eb";	// Sunday background color

var gcOtherDay="silver";	// the font color of days in other months; when hiding, it's also the background color.
var giShowOther=2;	// control the look of days in OTHER months. 1: show date & agendas effects; 2: show selected & today effects; 4: hide days in previous month; 8: hide days in next month. NOTE: values can be added up to create mix effects.

var gbFocus=true;	// whether to enable the gcToggle highlight whenever mouse pointer focuses over a calendar cell.
var gcToggle="red";	// the highlight color for the focused cell

var gcFGToday="red";	// the font color for today 
var gcBGToday="white";	// the background color for today 
var guTodayBGImg="";	// url of image as today's background
var giMarkToday=1; // Effects for today - 0: nothing; 1: set background color with gcBGToday; 2: draw a box with gcBGToday; 4: bold the font; 8: set font color with gcFGToday; 16: set background image with guTodayBGImg; - they can be added up to create mixed effects.

var gcFGSelected="white";	// the font color for the selected date
var gcBGSelected="red";	// the background color for the selected date
var guSelectedBGImg="";	// url of image as background of the selected date
var giMarkSelected=2;	// Effects for selected date - 0: nothing; 1: set background color with gcBGSelected; 2: draw a box with gcBGSelected; 4: bold the font; 8: set font color with gcFGSelected; 16: set background image with guSelectedBGImg; - they can be added up to create mixed effects.

var gbBoldAgenda=true;	// whether to boldface the dates with agendas.
var gbInvertBold=false;	// true: invert the boldface effect set by gbBoldAgenda; false: no inverts.
var gbShrink2fit=true;	// whether to hide the week line if none of its day belongs to the current month.
var gdSelect=gToday;	// default selected date in format of [year, month, day]; [0,0,0] means no default date selected.
var giFreeDiv=0;	// The number of absolutely positioned layers you want to customize, they will be named as "freeDiv0", "freeDiv1"...
var gAgendaMask=[-1,-1,-1,null,null,-1,null];	// Set the relevant bit to -1 to keep the original agenda info of that bit unchanged, otherwise the new value will substitute the one defined in agenda.js.
