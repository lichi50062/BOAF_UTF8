
/*
 * Bazillyo's Spiffy DHTML Popup Calendar v. 1.0 ?2000 S. Ousta
 *   - freeware with this comment
 *   - for download size, you can strip all spaces & comments except the ?notices
 *   - Thanks to Chris for the domlay() function
 *   - this requires calendarcode.js, calendar.css, and calendarTest.htm
 *   - works in IE4.x, IE5.x, NS4.75 possibly 4.x, NS6 (with slight cosmetic issues)
 *   - Netscape does display some controls overtop of the layer so layout is important
 *
 */

// initialiZe variables...
var ppcIE=((navigator.appName == "Microsoft Internet Explorer") || ((navigator.appName == "Netscape") && (parseInt(navigator.appVersion)==5)));
var ppcNN6=((navigator.appName == "Netscape") && (parseInt(navigator.appVersion)==5));
//var ppcIE=(navigator.appName == "Microsoft Internet Explorer");
var ppcNN=((navigator.appName == "Netscape")&&(document.layers));
var ppcX = 4;
var ppcY = 4;

var IsCalendarVisible;
var calfrmName;
var maxYearList;
var minYearList;
var todayDate = new Date;
var curDate = new Date;
var curImg;
var curDateBox;
var curYear;
var curMonth;
var curDay;
var minDate = new Date;
var maxDate = new Date;
var hideDropDowns;
var IsUsingMinMax;
var FuncsToRun;
var calRowId;
var calRowFlag;

var img_del;
var img_close;
img_del=new Image();
img_del.src="../images/btn_del_small.gif";
img_close=new Image();
img_close.src="../images/btn_close_small.gif";

// minYearList=todayDate.getFullYear()-;
minYearList= 2001;

maxYearList=todayDate.getFullYear()+5;
//minYearList=todayDate.getFullYear()-1901;
//maxYearList=todayDate.getFullYear()+1921;
IsCalendarVisible=false;



function calSwapImg(whatID, NewImg,override) {
    if (document.images) {
     if (!( IsCalendarVisible && override )) {
        document.images[whatID].src = eval(NewImg + ".src");
     }
    }
    window.status=' ';
    return true;
}
/*
deprecated in the case
function getOffsetLeft (el) {
    var ol = el.offsetLeft;
    while ((el = el.offsetParent) != null)
        ol += el.offsetLeft;
    return ol;
}

function getOffsetTop (el) {
    var ot = el.offsetTop;
    while((el = el.offsetParent) != null)
        ot += el.offsetTop;
    return ot;
}
*/

function setRowInfo(id, flag) {
	if(id==null) id="";
	if(flag==null) flag="";
	calRowId=id;
  calRowFlag=flag;
}

//modify from calendarcode.js showCalendar() method
function showCalendar(frmName, dteBox,btnImg, hideDrops, MnDt, MnMo, MnYr, MxDt, MxMo, MxYr,runFuncs) {
//	alert(frmName);
	hideDropDowns = hideDrops;
    FuncsToRun = runFuncs;
    calfrmName = frmName;
    if (IsCalendarVisible) {
        hideCalendar();
    }
    else {

        if (hideDropDowns) {toggleDropDowns('hidden');}
        if ((MnDt!=null) && (MnMo!=null) && (MnYr!=null) && (MxDt!=null) && (MxMo!=null) && (MxYr!=null)) {
            IsUsingMinMax = true;
            minDate.setDate(MnDt);
            minDate.setMonth(MnMo-1);
            minDate.setFullYear(MnYr);
            maxDate.setDate(MxDt);
            maxDate.setMonth(MxMo-1);
            maxDate.setFullYear(MxYr);
        }
        else {
            IsUsingMinMax = false;
        }

        curImg = btnImg;
        curDateBox = dteBox;

		ppcX=10;
		ppcY=10;
        domlay('popupcalendar',1,ppcX,ppcY,Calendar(todayDate.getMonth(),todayDate.getFullYear()));
        IsCalendarVisible = true;
    }
}

function toggleDropDowns(showHow){
    var i; var j;
    for (i=0;i<document.forms.length;i++) {
        for (j=0;j<document.forms[i].elements.length;j++) {
            if (document.forms[i].elements[j].tagName == "SELECT") {
                if (document.forms[i].name != "Cal")
                    document.forms[i].elements[j].style.visibility=showHow;
            }
        }
    }
}

function hideCalendar(){
	self.close();
}

function domlay(id,trigger,lax,lay,content) {
    /*
     * Cross browser Layer visibility / Placement Routine
     * Done by Chris Heilmann
     * Feel free to use with these lines included!
     * Created with help from Scott Andrews.
     * The marked part of the content change routine is taken
     * from a script by Reyn posted in the DHTML
     * Forum at Website Attraction and changed to work with
     * any layername. Cheers to that!
     * Welcome DOM-1, about time you got included... :)
     */
    // Layer visible
    if (trigger=="1"){
        if (document.layers) document.layers[''+id+''].visibility = "show"
        else if (document.all) document.all[''+id+''].style.visibility = "visible"
        else if (document.getElementById) document.getElementById(''+id+'').style.visibility = "visible"
        }
    // Layer hidden
    else if (trigger=="0"){
        if (document.layers) document.layers[''+id+''].visibility = "hide"
        else if (document.all) document.all[''+id+''].style.visibility = "hidden"
        else if (document.getElementById) document.getElementById(''+id+'').style.visibility = "hidden"
        }
    // Set horizontal position
    if (lax){
        if (document.layers){document.layers[''+id+''].left = lax}
        else if (document.all){document.all[''+id+''].style.left=lax}
        else if (document.getElementById){document.getElementById(''+id+'').style.left=lax+"px"}
        }
    // Set vertical position
    if (lay){
        if (document.layers){document.layers[''+id+''].top = lay}
        else if (document.all){document.all[''+id+''].style.top=lay}
        else if (document.getElementById){document.getElementById(''+id+'').style.top=lay+"px"}
        }
    // change content

    if (content){
    if (document.layers){
        sprite=document.layers[''+id+''].document;
        // add father layers if needed! document.layers[''+father+'']...
        sprite.open();
        sprite.write(content);
        sprite.close();
        }
    else if (document.all) document.all[''+id+''].innerHTML = content;
    else if (document.getElementById){
        //Thanx Reyn!
        rng = document.createRange();
        el = document.getElementById(''+id+'');
        rng.setStartBefore(el);
        htmlFrag = rng.createContextualFragment(content)
        while(el.hasChildNodes()) el.removeChild(el.lastChild);
        el.appendChild(htmlFrag);
        // end of Reyn ;)
        }
    }
}

//localized
function Calendar(whatMonth,whatYear) {
    var output = '';
    var datecolwidth;
    var startMonth;
    var startYear;
    startMonth=whatMonth;
    startYear=whatYear;

    curDate.setMonth(whatMonth);
    curDate.setFullYear(whatYear);
    curDate.setDate(todayDate.getDate());
    if (ppcNN6) {
        output += '<form name="Cal"><table width="185" border="3" class="cal-Table" cellspacing="0" cellpadding="0"><tr>';
    }
    else {
        output += '<table width="185" border="3" class="cal-Table" cellspacing="0" cellpadding="0"><form name="Cal"><tr>';
    }

    output += '<td class="cal-HeadCell" align="center" width="100%"><a href="javascript:clearDay();"><img name="calbtn1" src="./images/btn_del_small.gif" border="0" width="12" height="10"></a><a href="javascript:scrollMonth(-1);" class="cal-DayLink">&lt;</a><SELECT class="cal-TextBox" NAME="cboYear" onChange="changeYear();">';

    for (year=(minYearList); year< (maxYearList); year++) {
        if (year == (whatYear)) output += '<OPTION VALUE="' + year + '" SELECTED>' + (year-1911) + '<\/OPTION>';
        else              output += '<OPTION VALUE="' + year + '">'          + (year-1911) + '<\/OPTION>';
    }
	output += '<\/SELECT>年<SELECT class="cal-TextBox" NAME="cboMonth" onChange="changeMonth();">';
    for (month=0; month<12; month++) {
        if (month == whatMonth) output += '<OPTION VALUE="' + month + '" SELECTED>' + names[month] + '<\/OPTION>';
        else                output += '<OPTION VALUE="' + month + '">'          + names[month] + '<\/OPTION>';
    }
    output += '<\/SELECT>月<a href="javascript:scrollMonth(1);" class="cal-DayLink">&gt;</a><a href="javascript:hideCalendar();"><img name="calbtn2" src="./images/btn_close_small.gif" border="0" width="12" height="10"></a><\/td><\/tr><tr><td width="100%" align="center">';

    firstDay = new Date(whatYear,whatMonth,1);
    startDay = firstDay.getDay();

    if (((whatYear % 4 == 0) && (whatYear % 100 != 0)) || (whatYear % 400 == 0))
         days[1] = 29;
    else
         days[1] = 28;

    output += '<table width="185" cellspacing="1" cellpadding="2" border="0"><tr>';

    for (i=0; i<7; i++) {
        if (i==0 || i==6) {
            datecolwidth="15%"
        }
        else
        {
            datecolwidth="14%"
        }
        output += '<td class="cal-HeadCell" width="' + datecolwidth + '" align="center" valign="middle">'+ dow[i] +'<\/td>';
    }

    output += '<\/tr><tr>';

    var column = 0;
    var lastMonth = whatMonth - 1;
    var lastYear = whatYear;
    if (lastMonth == -1) { lastMonth = 11; lastYear=lastYear-1;}

    for (i=0; i<startDay; i++, column++) {
        output += getDayLink((days[lastMonth]-startDay+i+1),true,lastMonth,lastYear);
    }

    for (i=1; i<=days[whatMonth]; i++, column++) {
        output += getDayLink(i,false,whatMonth,whatYear);
        if (column == 6) {
            output += '<\/tr><tr>';
            column = -1;
        }
    }

    var nextMonth = whatMonth+1;
    var nextYear = whatYear;
    if (nextMonth==12) { nextMonth=0; nextYear=nextYear+1;}

    if (column > 0) {
        for (i=1; column<7; i++, column++) {
            output +=  getDayLink(i,true,nextMonth,nextYear);
        }
        output += '<\/tr><\/table><\/td><\/tr>';
    }
    else {
        output = output.substr(0,output.length-4); // remove the <tr> from the end if there's no last row
        output += '<\/table><\/td><\/tr>';
    }

    if (ppcNN6) {
        output += '<\/table><\/form>';
    }
    else {
        output += '<\/form><\/table>';
    }
    curDate.setDate(1);
    curDate.setMonth(startMonth);
    curDate.setFullYear(startYear);
    return output;
}

function getDayLink(linkDay,isGreyDate,linkMonth,linkYear) {
    var templink;
    if (!(IsUsingMinMax)) {
        if (isGreyDate) {
            templink='<td align="center" class="cal-GreyDate">' + linkDay + '<\/td>';
        }
        else {
            if (isDayToday(linkDay)) {
                templink='<td align="center" class="cal-DayCell">' + '<a class="cal-TodayLink" onmouseover="self.status=\' \';return true" href="javascript:changeDay(' + linkDay + ');">' + linkDay + '<\/a>' +'<\/td>';
            }
            else {
                templink='<td align="center" class="cal-DayCell">' + '<a class="cal-DayLink" onmouseover="self.status=\' \';return true" href="javascript:changeDay(' + linkDay + ');">' + linkDay + '<\/a>' +'<\/td>';
            }
        }
    }
    else {
        if (isDayValid(linkDay,linkMonth,linkYear)) {

            if (isGreyDate){
                templink='<td align="center" class="cal-GreyDate">' + linkDay + '<\/td>';
            }
            else {
                if (isDayToday(linkDay)) {
                    templink='<td align="center" class="cal-DayCell">' + '<a class="cal-TodayLink" onmouseover="self.status=\' \';return true" href="javascript:changeDay(' + linkDay + ');">' + linkDay + '<\/a>' +'<\/td>';
                }
                else {
                    templink='<td align="center" class="cal-DayCell">' + '<a class="cal-DayLink" onmouseover="self.status=\' \';return true" href="javascript:changeDay(' + linkDay + ');">' + linkDay + '<\/a>' +'<\/td>';
                }
            }
        }
        else {
            templink='<td align="center" class="cal-GreyInvalidDate">'+ linkDay + '<\/td>';
        }
    }
    return templink;
}

function isDayToday(isDay) {
    if ((curDate.getFullYear() == todayDate.getFullYear()) && (curDate.getMonth() == todayDate.getMonth()) && (isDay == todayDate.getDate())) {
        return true;
    }
    else {
        return false;
    }
}

function isDayValid(validDay, validMonth, validYear){

    curDate.setDate(validDay);
    curDate.setMonth(validMonth);
    curDate.setFullYear(validYear);

    if ((curDate>=minDate) && (curDate<=maxDate)) {
        return true;
    }
    else {
        return false;
    }
}

function padout(number) { return (number < 10) ? '0' + number : number; }

function splitCurDateBox(val)
{
	var count=0;
	var pos1=0;
	var pos=val.indexOf(",");
	while(pos>-1)
	{
		if(count==0)
			curYear=curDateBox.substring(pos1,pos);
		if(count==1)
			curMonth=curDateBox.substring(pos1,pos);

		pos1=pos+1;
		pos=val.indexOf(",",pos1);
		count=count+1;
	}
	if(pos1<val.length)
		curDay=curDateBox.substring(pos1);

}

function clearDay() {
	if(curDateBox.indexOf(",")>-1)
	{
		splitCurDateBox(curDateBox);

		eval('opener.document.' + calfrmName + '.' + curYear + '.value = \'\'');
		eval('opener.document.' + calfrmName + '.' + curMonth + '.value = \'\'');
		eval('opener.document.' + calfrmName + '.' + curDay + '.value = \'\'');

	}
	else
    eval('opener.document.' + calfrmName + '.' + curDateBox + '.value = \'\'');
//    hideCalendar();
    if (FuncsToRun!=null)
        eval(FuncsToRun);
}

//return selected value to TextField傳值如下
function changeDay(whatDay) {
	var y="";
    curDate.setDate(whatDay);
 //   eval('document.' + calfrmName + '.' + curDateBox + '.value = "'+ padout(curDate.getDate()) + '-' + names[curDate.getMonth()] + '-' + curDate.getFullYear() + '"');
 //   eval('document.' + calfrmName + '.' + curDateBox + '.value = "' + curDate.getFullYear() + '/' +  names[curDate.getMonth()] + '/'+ padout(curDate.getDate()) + '"');
    a = parseInt(curDate.getFullYear())-1911;
    //a = parseInt(curDate.getFullYear());
    //eval('document.' + calfrmName + '.' + curDateBox + '.value = "' + a + '/' +  names[curDate.getMonth()] + '/'+ padout(curDate.getDate()) + '"');
    if(curDateBox.indexOf(",")>-1)
    {
    	splitCurDateBox(curDateBox);

    	if(a<100)
    		y="0"+a;
    	else
    		y=""+a;

    	if( eval('opener.document.' + calfrmName + '.' + curYear + '.length' )) {
    	  id = eval('opener.document.' + calfrmName + '.' + calRowId );
    	  for(var i=0; i<id.length; i++) {
    	    if(id[i].value == calRowFlag ) {
    	    	eval('opener.document.' + calfrmName + '.' + curYear + '['+i+'].value = "' + y + '"');
    	      eval('opener.document.' + calfrmName + '.' + curMonth + '['+i+'].value = "' + names[curDate.getMonth()] + '"');
    	      eval('opener.document.' + calfrmName + '.' + curDay + '['+i+'].value = "' + padout(curDate.getDate()) + '"');

    	    }
    	  }
      } else {

    	  eval('opener.document.' + calfrmName + '.' + curYear + '.value = "' + y + '"');
    	  eval('opener.document.' + calfrmName + '.' + curMonth + '.value = "' + names[curDate.getMonth()] + '"');
    	  eval('opener.document.' + calfrmName + '.' + curDay + '.value = "' + padout(curDate.getDate()) + '"');
      }
    }
    else
		eval('opener.document.' + calfrmName + '.' + curDateBox + '.value = "' + a + names[curDate.getMonth()] + padout(curDate.getDate()) + '"');
    hideCalendar();
    if (FuncsToRun!=null)
        eval(FuncsToRun);
}

function scrollMonth(amount) {
    var monthCheck;
    var yearCheck;

    if (ppcIE) {
        monthCheck = document.forms["Cal"].cboMonth.selectedIndex + amount;
    }
    else if (ppcNN) {
        monthCheck = document.popupcalendar.document.forms["Cal"].cboMonth.selectedIndex + amount;
    }
    if (monthCheck < 0) {
        yearCheck = curDate.getFullYear() - 1;
        if ( yearCheck < minYearList ) {
            yearCheck = minYearList;
            monthCheck = 0;
        }
        else {
            monthCheck = 11;
        }
        curDate.setFullYear(yearCheck);
    }
    else if (monthCheck >11) {
        yearCheck = curDate.getFullYear() + 1;
        if ( yearCheck > maxYearList-1 ) {
            yearCheck = maxYearList-1;
            monthCheck = 11;
        }
        else {
            monthCheck = 0;
        }
        curDate.setFullYear(yearCheck);
    }

    if (ppcIE) {
        curDate.setMonth(document.forms["Cal"].cboMonth.options[monthCheck].value);
    }
    else if (ppcNN) {
        curDate.setMonth(document.popupcalendar.document.forms["Cal"].cboMonth.options[monthCheck].value );
    }
    domlay('popupcalendar',1,ppcX,ppcY,Calendar(curDate.getMonth(),curDate.getFullYear()));
}

function changeMonth() {

    if (ppcIE) {
        curDate.setMonth(document.forms["Cal"].cboMonth.options[document.forms["Cal"].cboMonth.selectedIndex].value);
        domlay('popupcalendar',1,ppcX,ppcY,Calendar(curDate.getMonth(),curDate.getFullYear()));
    }
    else if (ppcNN) {

        curDate.setMonth(document.popupcalendar.document.forms["Cal"].cboMonth.options[document.popupcalendar.document.forms["Cal"].cboMonth.selectedIndex].value);
        domlay('popupcalendar',1,ppcX,ppcY,Calendar(curDate.getMonth(),curDate.getFullYear()));
    }

}

function changeYear() {
    if (ppcIE) {

        curDate.setFullYear(document.forms["Cal"].cboYear.options[document.forms["Cal"].cboYear.selectedIndex].value);
        domlay('popupcalendar',1,ppcX,ppcY,Calendar(curDate.getMonth(),curDate.getFullYear()));


    }
    else if (ppcNN) {

        curDate.setFullYear(document.popupcalendar.document.forms["Cal"].cboYear.options[document.popupcalendar.document.forms["Cal"].cboYear.selectedIndex].value);
        domlay('popupcalendar',1,ppcX,ppcY,Calendar(curDate.getMonth(),curDate.getFullYear()));
    }

}

function makeArray0() {
    for (i = 0; i<makeArray0.arguments.length; i++)
        this[i] = makeArray0.arguments[i];
}

var names     = new makeArray0('01','02','03','04','05','06','07','08','09','10','11','12');
var days      = new makeArray0(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);
var dow       = new makeArray0('日','一','二','三','四','五','六');
//ycm style
var loadOK = false;
