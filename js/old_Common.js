function AskReset(form) {
  if(confirm("確定要清除所有資料嗎？")){  	
  	form.reset();    	 	
  }  
  return;
}

function AskDelete() {
  if(confirm("確定要刪除嗎？"))
    return true;
  else
    return false;
}
function AskUpdate() {
  if(confirm("確定要修改嗎？"))
    return true;
  else
    return false;
}
function AskInsert() {
  if(confirm("確定要新增嗎？"))
    return true;
  else
    return false;
}

function AskAbdicate() {
  if(confirm("確定要卸任嗎？")) {
    return true;
  }else {
    return false;
  }  
}
//93.11.26 add
function AskRevoke() {
  if(confirm("確定要裁撤嗎？")) {
    return true;
  }else {
    return false;
  }  
}
//93.11.26
function Add_Manager() {
	alert("請先輸入主檔資料");
	return false;
}

//=====================================================================
function changeVal(T1) {
//	alert('changeval:->' + T1.value);
    pos = 0
    var oring = T1.value;
    pos = oring.indexOf(",");

    while (pos != -1) {
        oring = (oring).replace(",", "");
        pos = oring.indexOf(",");
    }
//    alert('oring:->' + oring);
    return oring;

}
//================================================================
function changeStr(T1) {
//	alert(T1);
    T1.value = changeVal(T1);
    //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    c="";
    var oring = T1.value
    var t1v1  = T1.value
    var t1v2  = "";
   //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    if (isNaN(Math.abs(T1.value)))
    {

   	    alert("請輸入數字");
   	    return oring;
    }else{
        if (eval(T1.value) < 0 )
            c="-";
    }

    if(T1.value.length == 0)
        return oring;
    if(Math.abs(T1.value) == 0)
        return oring;

    T1.value= Math.abs(T1.value);
    t1v1  = T1.value

    if((pos=t1v1.indexOf(".")) != -1)
    {
        t1v2 = t1v1.substring(pos,((T1.value).length));
        t1v1 = t1v1.substring(0,pos);
    }
  //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    pos=oring.indexOf(",")
    if (pos==-1)
    {
        var len   = t1v1.length;
        a=Math.floor(len % 3)
        b=Math.floor(len / 3 -1)
        if(a !=0 && b >= 0)
            oring = t1v1.substring(0,a) + ",";
        else if(b<0)
                oring = t1v1
             else
                oring = "";
        for(i=0;i<b;i++)
        {
            oring += t1v1.substring(a,a+3) + ",";
            a += 3;
        } // end of for

        oring += t1v1.substring(a,len);

        if (c == "-" )
        {
            oring="-"+oring;
        }
        if(t1v2 != "")
            oring += t1v2;

        return oring
     }
}
//===維護人員畫面欄位CHECK=============
function Check_Maintain(form1) {
   if (trimString(form1.maintain_dept.value)=="" || trimString(form1.maintain_tel.value)==""){
        alert("維護部門名稱.承辦員電話欄位不可空白！");
        return false;
   }
   return true
}
//==================================================
function CheckYear(cyear) {
    if(cyear.value.indexOf(".") != -1)
    {
        alert("年份不可為小數");
        return false;
    }
    if(isNaN(Math.abs(cyear.value)))
    {
        alert("年份不可為文字");
        return false;
    }
    if(trimString(cyear.value) != '')
        cyear.value = Math.abs(cyear.value);
    return true;
}
//================================================================
function trimString(inString) {

	var outString;
	var startPos;
	var endPos;
	var ch;

	// where do we start?
	startPos = 0;
	test = 0;
	ch = inString.charAt(startPos);
	while ((ch == " ") || (ch == "\b") || (ch == "\f") || (ch == "\n") || (ch == "\r") || (ch == "\n")) {
		startPos++;
		if ( ch==" " ) {
			test++;
		}
		ch = inString.charAt(startPos);
	}
     if  ( test==inString.length )
     	flag = true;
     else
     	flag = false;
	endPos = inString.length - 1;
	ch = inString.charAt(endPos);
	while ((ch == " ") || (ch == "\b") || (ch == "\f") || (ch == "\n") || (ch == "\r") || (ch == "\n")) {
		endPos--;
		ch = inString.charAt(endPos);
	}

	// get the string
	outString = inString.substring(startPos, endPos + 1);
	if ( flag==true )
		return "";
	else
		return outString;
}
function ChgAction(form, actionURL) {

    form.action = actionURL;
	form.submit();
}
//=================================================================
function checkSingleYM(checkyear, checkmonth) {

	if (!CheckYear(checkyear))
    	return false;

//    if (trimString(checkyear.value) == "" || checkmonth.options[checkmonth.selectedIndex].value == ""){
    if (trimString(checkyear.value) == "" || checkmonth.value == ""){
		alert("請輸入日期:");
        return false;
    }
    return true;
}

//=================================================================
function CheckQueryDate2(checkyear, checkmonth, checkdate, szWarning) {

    if(!CheckYear(checkyear))
    	return false;
    if(trimString(checkyear.value)!="" && checkmonth.options[checkmonth.selectedIndex].value!="") {
        checkdate.value = Math.abs(checkyear.value) * 100 + Math.abs(checkmonth.options[checkmonth.selectedIndex].value);
    }
    else if(trimString(checkyear.value)=="" && checkmonth.options[checkmonth.selectedIndex].value==""){}
         else {
            alert("請輸入完整的"+ szWarning+ "或者都不輸入");
            return false;
         }
    return true
}
//=========================================================
function Check_QueryDate(form) {

    if(!CheckQueryDate2(form.S_YEAR,form.S_MONTH,form.S_DATE,"起始日期"))
    	return false;
    if(!CheckQueryDate2(form.E_YEAR,form.E_MONTH,form.E_DATE,"結束日期"))
        return false;
	if(trimString(form.S_DATE.value)!="" && trimString(form.E_DATE.value)!="") {
		if(Math.abs(form.S_DATE.value) > Math.abs(form.E_DATE.value)) {
    		alert("起始日期不可以大於終止日期");
    		return false;
    	}
    }
    return true;
}
//=========================================================
function checkNumber(Rate1){

	Rate1.value = changeVal(Rate1);

    if (isNaN(Math.abs(Rate1.value))) {
   	    alert("請輸入數字");
   	    return false;
    }
    if ((Rate1.value).indexOf(".") != -1 ) {
    	alert("不可輸入小數點");
        return false;
    }
    return true;
}
function fnValidDate(dateStr) {

    var leap = 28;
    if(leapYear(parseInt(dateStr.substring(0,4))) == true)
        leap = 29;
    var tmp = parseInt(dateStr.substring(5, 7))
    if(dateStr.substring(5, 7) == '08')
        tmp = 8;
    if(dateStr.substring(5, 7) == '09')
        tmp = 9
    if(tmp < 1 || tmp > 12){
        return (false)
    }
    var monthTable = new Array(12);
    monthTable[1] = 31;
    monthTable[2] = leap;
    monthTable[3] = 31;
    monthTable[4] = 30;
    monthTable[5] = 31;
    monthTable[6] = 30;
    monthTable[7] = 31;
    monthTable[8] = 31;
    monthTable[9] = 30;
    monthTable[10] = 31;
    monthTable[11] = 30;
    monthTable[12] = 31;

    var dtmp = parseInt(dateStr.substring(8))

    if(dateStr.substring(8) == '08')
        dtmp = 8;
    if(dateStr.substring(8) == '09')
        dtmp = 9

    if(dtmp < 1 || dtmp > monthTable[tmp]){
        return false
    }

    return true
}

function leapYear (Year) {
        if (((Year % 4)==0) && ((Year % 100)!=0) || ((Year % 400)==0))
                return (true);
        else
                return (false);
}

//==================
//opt : 0 必須輸入
//opt : 1 可不輸入
function checkDate(yr, mn, dt, txt, opt) {

	var sDate, ckDate;
	if (opt == 0) {
		if (trimString(yr.value) == "" ){
			alert("[" + txt + "] (年)欄不可空白");
			yr.focus();
			return false;
		}
		else {
	    	sDate = Math.abs(yr.value);
	        if (isNaN(sDate)) {
	            alert("[" + txt +"] (年)欄 不可輸入文字");
	            yr.focus();
	            return false;
	        }
	    }
		if (trimString(mn.value) == "" ){
			alert("[" + txt + "] (月)欄不可空白");
			mn.focus();
			return false;
		}
		if (trimString(dt.value) == "" ){
			alert("[" + txt + "] (日)欄不可空白");
			dt.focus();
			return false;
		}
	    ckDate = '' + (parseInt(yr.value) + 1911) + '/' + mn.value + '/' + dt.value;

	    if (fnValidDate(ckDate) != true){
	    	alert('所輸入的時間為無效日期!');
	        dt.focus();
	        return false;
	   	}
	}
	else {
		if (trimString(yr.value) == "" )	{
			if (trimString(mn.value) != "" ) {
				alert("[" + txt +"](年,月,日)欄 必須全填或全不填");
	        	yr.focus();
				return false;
			}
			else {
				if (trimString(dt.value) != "" ) 	{
					alert("[" + txt + "](年,月,日)欄 必須全填或全不填");
	        		dt.focus();
					return false;
				}
			}
		}
		else {
	        sDate = Math.abs(yr.value);
	        if (isNaN(sDate)) {
	            alert("[" + txt + "](年)欄 不可輸入文字");
	        	yr.focus();
	            return false;
	        }

			if (trimString(mn.value) == "" ) {
				alert("[" + txt + "](月)欄不可空白");
	        	mn.focus();
				return false;
			}
			else {
				if (trimString(dt.value) == "" ) {
					alert("[" + txt + "](日)欄不可空白");
	    	    	dt.focus();
					return false;
				}
	    		ckDate = '' + (parseInt(yr.value) + 1911) + '/' + mn.value + '/' + dt.value;

		    	if (fnValidDate(ckDate) != true) {
		    	    alert('所輸入的時間為無效日期!');
	    	        dt.focus();
	        	    return false;
	   			}
	   		}
	   	}
	}
	return true;
}

//在網頁上alert訊息,按確定後,回到上頁
function AlertMsg(msg) {
		if(msg != ''){
		   alert(msg);		   
		   history.back();		
	    }	    
	    return true;		
}
//在網頁一alert confirm的訊息,"Y"-->URL ,"N"-->回上一頁/或到URL
function ConfirmMsg(form,msg,weburl_Y,weburl_N){		
		if(msg != ''){			
	    	if (confirm(msg)) {	    	    		
	    		form.action=weburl_Y;		    		
      			form.submit();
    		}else{    			
    		    if(weburl_N != ''){    		     
    		       form.action=weburl_N;		    		
      			   form.submit();
    			}else{
    				history.back();		    		    
    			}
    		}		   		
	    }	    
	    return true;	    
}	
  
  