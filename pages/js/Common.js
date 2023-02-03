// 94.04.01 add 檢核身份証字號
// 94.11.01 add 檢查E-Mail是否錯誤
// 95.02.08 add myConfirm對話框,預設在取消
// 97.01.10 add 檢查輸入的字串長度
// 97.09.26 fix 修正日期取得格式 by 2295
//103.02.21 fix 檢查E-Mail,取消檢核第二個dot . by 2295 
function check_identity_no(identity_no) {
   tab = "ABCDEFGHJKLMNPQRSTUVWXYZIO"
   A1 = new Array (1,1,1,1,1,1,1,1,1,1,2,2,2,2,2,2,2,2,2,2,3,3,3,3,3,3 );
   A2 = new Array (0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5 );
   Mx = new Array (9,8,7,6,5,4,3,2,1,1);

   if ( identity_no.length != 10 ){        
        return false;
   }
   i = tab.indexOf( identity_no.charAt(0) );
   if ( i == -1 ){        
        return false;
   }
   sum = A1[i] + A2[i]*9;

   for ( i=1; i<10; i++ ) {
      v = parseInt( identity_no.charAt(i) );
      if ( isNaN(v) ){          
           return false;
      }
      sum = sum + v * Mx[i];
   }
   if ( sum % 10 != 0 ){       
       return false;
   }

   return true;
}


/*************************************************
功能：可以自己定制的confirm提示框，在IE6下測試通過

這個confirm用還可以做好多擴展，比如修改背景的顏色，修改那個顯示的圖片，修改按扭的樣式，

**************************************************/

/*
參數說明： strTitle   confirm框的標題
   		   strMessage confirm框要顯示的消息
   		   intType    confirm框的選中的類型1為選中YES，2為選中NO
           intWidth   confirm框的寬度
           intHeight  confirm框的高度
*/
//95.02.08 add myConfirm對話框,預設在取消
function myConfirm(strTitle,strMessage,intType,intWidth,intHeight)
{
 var strDialogFeatures = "dialogHeight:"+intHeight+";dialogWidth:"+intWidth+";center:yes;help:no;resizable:yes;scroll:no;status:no;defaultStatus:no;";
 
 var args = new Array();
 args[args.length] = strTitle;
 args[args.length] = strMessage;
 args[args.length] = intType;
 var result = showModalDialog("myConfirm.htm",args,strDialogFeatures);
 //var result = showModelessDialog("myConfirm.htm",args,strDialogFeatures);
 
 return result;
}



function AskReset(form) {
  if(confirm("確定要回復原有資料嗎？")){  	
  	form.reset();    	 	
  }  
  return;
}

function AskDelete() {
	
  //if(confirm("確定要刪除嗎？"))
  var myConfirmResult = myConfirm("Microsoft Internet Explorer","確定要刪除嗎？",2,15,8)
  if(myConfirmResult)  
    return true;
  else
    return false;
}
function AskUpdate() {
  var myConfirmResult = myConfirm("Microsoft Internet Explorer","確定要修改嗎？",2,15,8)
  if(myConfirmResult)
    return true;
  else  
  	return false;
  /*	
  if(confirm("確定要修改嗎？"))
    return true;
  else
    return false;
  */  
}
function AskInsert() {
  //if(confirm("確定要新增嗎？"))
  var myConfirmResult = myConfirm("Microsoft Internet Explorer","確定要新增嗎？",2,15,8)
  if(myConfirmResult)
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

//93.12.10 add
function AskResetPwd() {
  if(confirm("確定要重設密碼嗎？")) {
    return true;
  }else {
    return false;
  }  
}
//94.01.03 add by 2295
function AskLock() {
  if(confirm("確定要鎖定嗎？")) {
    return true;
  }else {
    return false;
  }  
}
//94.01.03 add by 2295
function AskUnLock() {
  if(confirm("確定要解除鎖定嗎？")) {
    return true;
  }else {
    return false;
  }  
}

function changeVal(T1) {
	//alert('changeval:->' + T1.value);
    pos = 0
    var oring = T1.value;
    pos = oring.indexOf(",");

    while (pos != -1) {
        oring = (oring).replace(",", "");
        pos = oring.indexOf(",");
    }
    //alert('oring:->' + oring);
    return oring;

}
//================================================================
function changeStr(T1) {
	//alert(T1);
    T1.value = changeVal(T1);
    //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    c="";
    var oring = T1.value
    var t1v1  = T1.value
    var t1v2  = "";
   //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    if (isNaN(Math.abs(T1.value)))
    {
		//alert(T1.value);
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

	if (!CheckYear(checkyear)){		
    	return false;
    }	

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
//add by 2295 若不可有小數點時,則停在該text field
function checkPoint_focus(Rate1){

	//93.03.16拿掉Rate1.value = changeVal(Rate1);
    if ((changeVal(Rate1)).indexOf(".") != -1 ) {
    	alert("不可輸入小數點");
    	Rate1.focus();
        return false;
    }
    return true;
}
//97.09.26 fix 修正日期取得格式 by 2295
function fnValidDate(dateStr) {

    var leap = 28;
     var tmpStr = dateStr.substring(5,dateStr.length);
    //alert(tmpStr);   
    //alert('年='+dateStr.substring(0,4));
    //alert('月='+tmpStr.substring(0, tmpStr.indexOf("/")));
    if(leapYear(parseInt(dateStr.substring(0,4))) == true)
        leap = 29;
    var tmp = parseInt(tmpStr.substring(0, tmpStr.indexOf("/")))
    if(tmpStr.substring(0, tmpStr.indexOf("/")) == '08')
        tmp = 8;
    if(tmpStr.substring(0, tmpStr.indexOf("/")) == '09')
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

    tmpStr = tmpStr.substring(tmpStr.indexOf("/")+1,tmpStr.length);
    //alert('日='+tmpStr);
    var dtmp = parseInt(tmpStr)

    if(tmpStr == '08')
        dtmp = 8;
    if(tmpStr == '09')
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
		//alert(ckDate);
	    if (fnValidDate(ckDate) != true) {
		    		  
		    		  if(dt.value==20)
		    		  {
		    		     return true;
		    		  }
		    		  else
		    		  {   
		    	    	alert('所輸入的時間為無效日期!');
	    	        dt.focus();
	        	    return false;
	        	  }
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

//在網頁上alert訊息,按確定後,回到若有link到該link..無則回上頁
function AlertMsg(form,msg,weburl_Y) {
		if(msg != ''){
		   alert(msg);		   
		   if(weburl_Y != ''){
		   	  form.action=weburl_Y;
		   	  form.submit();
		   }else{	
		     history.back();		
		   }  
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
  
//=========================================================
// Add by Winnin 2004.11.22
function checkFloatNumber(Rate1){

	Rate1.value = changeVal(Rate1);    
    if (isNaN(Math.abs(Rate1.value))) {    	    	
   	    alert("請輸入數字");
   	    return false;
    }
    return true;
}


/* 去除掉指定字串的前、後空白字元
 * @param   strText     想要去除前後空白字元的字串
 * @return  string
 */
function trim(strText) {
    strText = ltrim(strText);
    strText = rtrim(strText);

    return strText;
}

/* 去除掉指定字串的左側空白字元
 * @param   strText     想要去除左側空白字元的字串
 * @return  string
 */
function ltrim(strText) {
    while (strText.substring(0,1) == ' ')
        strText = strText.substring(1, strText.length);

    return strText;
}

/* 去除掉指定字串的右側空白字元
 * @param   strText     想要去除右側空白字元的字串
 * @return  string
 */
function rtrim(strText) {
    while (strText.substring(strText.length-1,strText.length) == ' ')
        strText = strText.substring(0, strText.length-1);

    return strText;
}

////================== add by 2354 2005.1.19 only到月
//opt : 0 必須輸入
//opt : 1 可不輸入
function checkDateS(yr, mn, txt, opt) {

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
	    ckDate = '' + (parseInt(yr.value) + 1911) + '/' + mn.value + '/' + '01';

	    if (fnValidDate(ckDate) != true){
	    	alert('所輸入的時間為無效日期!');
	    	
	        dt.focus();
	        return false;
	   	}
	}
	else {
		if (trimString(yr.value) == "" )	{
			if (trimString(mn.value) != "" ) {
				alert("[" + txt +"](年,月)欄 必須全填");	//modify by 2354 2005.1.19
	        	yr.focus();
				return false;
			}
			else {
				if (trimString(dt.value) != "" ) 	{
					alert("[" + txt + "](年,月)欄 必須全填");//modify by 2354 2005.1.19
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
	    		ckDate = '' + (parseInt(yr.value) + 1911) + '/' + mn.value + '/' + '01';

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

function CheckMaxLength(String,Length)
{
  if (Length == null || Length == "" || String == null || String == "") return false;
  var actualLen = 0;
  for (var i=0; i<String.length; i++)
    if (String.charCodeAt(i) < 127)
      actualLen++;
    else
      actualLen+=2;

  if (actualLen <= Length)
    return false;
  else
    return true;
}

//檢查E-Mail是否錯誤
/*103.02.21 add 取消檢核第二個dot . by 2295 */
function CheckEmailAddress(EmailString){
      if (EmailString.length <= 5) return false;      
      if (EmailString.indexOf('@', 0) == -1){
      	  return false;
      }else{      	
          EmailData = EmailString.substring(EmailString.indexOf('@', 0)+1,EmailString.length);
          //alert(EmailData);         
          if (EmailData.length <= 2) return false;               
          if (EmailData.indexOf('.', 0) == -1) return false;
          EmailData1 = EmailData.substring(EmailData.indexOf('.', 0)+1,EmailData.length);         
          //alert(EmailData1);
          /*103.02.21 add 取消檢核第二個dot . by 2295 */
          /*
          if (EmailData1.length <= 1) return false;               
          if (EmailData1.indexOf('.', 0) == -1) return false;
          EmailData2 = EmailData1.substring(EmailData1.indexOf('.', 0)+1,EmailData1.length);
          if (EmailData2.length <= 0) return false;               
          */
          //alert(EmailData2);
      }
      if (CheckMaxLength(EmailString,50)){
  		  alert("E-Mail長度輸入有誤，請重新輸入");
   	      return false;
      }
      return true;
}

//97,01.10 add 檢查輸入的字串長度
function getLength(t1)
{
         var cnt=0;
        for(var i=0;i<t1.length;i++)
        {
            if(escape(t1.charAt(i)).length>=4) cnt+=2;
            else
            	cnt++;
        }
	return cnt;
}