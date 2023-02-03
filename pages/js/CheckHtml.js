var thisBrowser = null;
if(navigator.appName == "Netscape")
    thisBrowser="Netscape";
else
    thisBrowser="IE";
function fnValidDate(dateStr)
{
    var leap = 28;
    if (leapYear(parseInt(dateStr.substring(0,4))) == 1)
        leap = 29;
    var tmp = parseInt(dateStr.substring(4,6))    
    if (dateStr.substring(4,6) == '08')
        tmp = 8;
    if (dateStr.substring(4,6) == '09')
        tmp = 9
    
    if (tmp < 1 || tmp > 12)
    {
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

    var dtmp = parseInt(dateStr.substring(6))    
    if(dateStr.substring(6) == '08')
        dtmp = 8;
    if(dateStr.substring(6) == '09')
        dtmp = 9

    if(dtmp < 1 || dtmp > monthTable[tmp])
    {
        return (false)
    }
    return (true)
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function leapYear (Year)
{
    if (((Year % 4)==0) && ((Year % 100)!=0) || ((Year % 400)==0))
        return (1);
    else
	return (0);
}
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function checkDateFormat(form)
{
    if(fnValidDate(form.S_YEAR.value + form.S_MONTH.value + form.S_DAY.value))
    {
        if(fnValidDate(form.E_YEAR.value + form.E_MONTH.value + form.E_DAY.value))
            return true
        else
        {
            alert('終止日期有誤');
            return false;
        }
    }
    else
    {
        alert('起始日期有誤');
        return false;
    }
}
//=============================================================================

function AskDelete() {
  if(confirm("確定要刪除嗎？")) {
    return true;
  }else {
    return false;
  }  
}

function DeleteAction(form1) {
  if(AskDelete()) {
    form1.Function.value = 'delete';
    form1.submit();  
  }
}

function ResetAction(form1) {
  form1.Function.value = 'passwd';
  form1.submit();
}

function Check_ZZ001W_ShowUpdate(form1) {
    if(form1.IorO!=null) //login為局內人
    {
        if(form1.IorO.value=="IN")//被update者為局內人
        {
            if(form1.group.options[form1.group.selectedIndex].value=="")
            {
                alert("請選擇組室");
                return false;    
            }
            if(form1.USER_NAME.value.length==0)    
            {
                alert("使用者名稱不可空白");
                return false;    
            }
        }
    }
    else //login為局外人
    {
        if(form1.USER_NAME.value.length==0)    
            {
                alert("使用者名稱不可空白");
                return false;    
            }
    }
    return true;
}

function Check_ZZ001W_ShowInsert(form1)
{
    form1.BANK_NO.value=trimString(form1.BANK_NO.value);
    form1.USER_NAME.value=trimString(form1.USER_NAME.value);
    if(!checkBankNo(form1.BANK_NO))
    {
        alert("使用者代碼不可為文字");
        return false;    
    }
    
    if(form1.mytype.value=="I")//Login者為局內人
    {
        if(form1.IorO[0].checked==true)//要新增局內人
        {
            if(form1.BANK_NO.value.length==0)
            {
                alert("使用者代碼不可空白");    
                return false;
            }
            if(form1.BANK_NO.value.length!=5)
            {
                alert("使用者代碼須輸入五碼");
                return false;
            }
            if(form1.group.options[form1.group.selectedIndex].value=="")
            {
                alert("請選擇組室");
                return false;    
            }
            if(form1.USER_NAME.value.length==0)
            {
                alert("使用者名稱不可空白");
                return false;    
            }
        }
        else // 要新增局外人
        {
            if(form1.BANK_NO.value.length==0)
            {
                alert("使用者代碼不可空白");    
                return false;
            }
            if(form1.BANK_NO.value.length!=12)
            {
                alert("使用者代碼須輸入12碼");
                return false;
            }
        }
    }
    else //Login者為局外人
    {
        if(form1.BANK_NO2.value.length==0)
        {
            alert("使用者代碼不可空白");    
            return false;
        }
        if(form1.BANK_NO2.value.length!=5)
        {
            alert("使用者代碼需填滿後五碼");    
            return false;
        }
        if(form1.USER_NAME.value.length==0)
        {
            alert("使用者名稱不可空白");
            return false;    
        }
        form1.BANK_NO.value=form1.BANK_NO1.value+form1.BANK_NO2.value;
    }    
    
    return true;
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

function Check_ZZ002W_ShowInsert(form1) {
  if(form1.USER_ID.options.length==0) {     
    alert("使用者不可空白");
    return false;
  }else {    
    return true;   
  }  
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

function Check_ZZ003W_ShowUpdate(form1) {
  if(form1.PROGRAM_NAME.value.length == 0) {
    alert("程式名稱不可空白");
    return false;
  }else {
    return true;
  }  
}

function Check_ZZ003W_ShowInsert(form1) {
  if(form1.PROGRAM_ID.value.length==0 ||
     form1.PROGRAM_NAME.value.length==0) {
    alert("程式代碼和名稱不可空白");
    return false;
  }else {
    if(form1.PROGRAM_ID.value.length == 6) {
      return true;
    }else {
      alert("程式代碼填滿6碼");
      return false;
    }
  }  
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

function add_value(src_name,dest_name) {
  if(src_name.options[src_name.selectedIndex].text != "") {
    if(IsExist(src_name.options[src_name.selectedIndex].text,dest_name) == true) {
      alert(' 選項已經存在，無法再新增！');
    }else {
      dest_name.options[dest_name.options.length] = 
      new Option(src_name.options[src_name.selectedIndex].
                 text,src_name.options[src_name.selectedIndex].value);
    }
  }    
}

function IsExist(Item,dest_name) {
  for(var i = dest_name.options.length-1; i >= 1; i--) {
    if(Item == dest_name.options[i].text) {    
      return true;
    }
  }    
  return false;
}

function del_value(dest_name) {
  for(var i = dest_name.options.length-1; i >= 1; i--) {
    dest_name.options[i] = null;
  }
}

function Check_ZZ004W(form1) {    
  if(form1.USER_ID.options.length > 1 &&
     form1.PROGRAM_ID.options.length > 1) {
    for(var i = 1; i <= form1.USER_ID.options.length-1; i++) {
      form1.USER_ID.options[i].selected = true;
    }
    for(var i = 1; i <= form1.PROGRAM_ID.options.length-1; i++) {
      form1.PROGRAM_ID.options[i].selected = true;
    }
    return true;
  }else {
    alert('使用者清單及程式清單不可為空白');
    return false;
  }
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

function Check_ZZ005W_ShowInsert(form1) {
  if(form1.USER_ID.options.length==0 ||
     form1.PROGRAM_ID.options.length==0) {     
    alert("使用者和程式名稱不可空白");
    return false;
  }else {    
    return true;   
  }  
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

function Check_ZZ006W(form1) {
  if(form1.OLDPASSWORD.value.length == 0 ||
     form1.NEWPASSWORD.value.length == 0 ||
     form1.AGAINPASSWORD.value.length == 0) {
    alert("使用者密碼不可空白");
    return false;
  }
  if(form1.NEWPASSWORD.value != form1.AGAINPASSWORD.value) {    
      alert("新密碼與新密碼確認不相同");
      return false;
  }
  if(form1.OLDPASSWORD.value == form1.NEWPASSWORD.value)
  {
  	alert("新密碼不可與舊密碼相同");
  	return false;
  } 
  return true;
    
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

function Check_MORE_YEAR(form1) {
  if(form1.E_YEAR.value > form1.S_YEAR.value) {
    return true;
  }else {
    return false;
  }
}

function Check_MORE_MONTH(form1) {
  if(form1.E_MONTH.value > form1.S_MONTH.value) {
    return true;
  }else {
    return false;
  }
}

function Check_MORE_DAY(form1) {
  if(form1.E_DAY.value > form1.S_DAY.value) {
    return true;
  }else {
    return false;
  }
}

function Check_EQUAL_YEAR(form1) {
  if(form1.E_YEAR.value == form1.S_YEAR.value) {
    return true;
  }else {
    return false;
  }
}

function Check_EQUAL_MONTH(form1) {
  if(form1.E_MONTH.value == form1.S_MONTH.value) {
    return true;
  }else {
    return false;
  }
}

function Check_EQUAL_DAY(form1) {
  if(form1.E_DAY.value == form1.S_DAY.value) {
    return true;
  }else {
    return false;
  }
}

function checkYear(MyYear)
{   
    
    if(isNaN(Math.abs(MyYear.value)))
    {
        alert("請輸入數字");
        return false;
    }
    else if(MyYear.value.length!=0)
        MyYear.value = Math.abs(MyYear.value);
    if(MyYear.value.length != 0 && MyYear.value.length != 4)
    {
        alert("請輸入西元 XXXX 年");
        return false;
    }
    return true;
}
//-----------------------------------------------------
function Check_ZZ007W(form1) {
  if(form1.S_YEAR.value.length  !=0 && form1.E_YEAR.value.length  !=0 &&
     form1.S_MONTH.value.length !=0 && form1.E_MONTH.value.length !=0 &&
     form1.S_DAY.value.length   !=0 && form1.E_DAY.value.length   !=0) {
     
     if(!checkDateFormat(form1)) {       
       return false;
     }
     if(!checkYear(form1.S_YEAR)) {       
       return false;
     }
     if(!checkYear(form1.E_YEAR)) {       
       return false;
     }
     if(Check_MORE_YEAR(form1)) {       
       return true;
     }
     if(Check_EQUAL_YEAR(form1) && Check_MORE_MONTH(form1)) {         
       return true;
     }
     if(Check_EQUAL_YEAR(form1) && Check_EQUAL_MONTH(form1) &&
        (Check_MORE_DAY(form1) || Check_EQUAL_DAY(form1))) {       
       return true;
     }else {
       alert("終止年月日必須大於或等於起始年月日");
       return false;
     }
  }else {
    if(form1.S_YEAR.value.length  !=0 && form1.E_YEAR.value.length  !=0 &&
       form1.S_MONTH.value.length !=0 && form1.E_MONTH.value.length !=0 &&
       form1.S_DAY.value.length   ==0 && form1.E_DAY.value.length   ==0) {
       
       if(!checkYear(form1.S_YEAR)) {       
         return false;
       }
       if(!checkYear(form1.E_YEAR)) {       
         return false;
       }
       if(Check_MORE_YEAR(form1)) {     
         return true;
       }       
       if(Check_EQUAL_YEAR(form1) && (Check_MORE_MONTH(form1) || Check_EQUAL_MONTH(form1))) {         
         return true;
       }else {
         alert("終止年月必須大於或等於起始年月");
         return false;
       }
    }else {
      if(form1.S_YEAR.value.length  ==0 && form1.E_YEAR.value.length  ==0 &&
         form1.S_MONTH.value.length ==0 && form1.E_MONTH.value.length ==0 &&
         form1.S_DAY.value.length   ==0 && form1.E_DAY.value.length   ==0) {
                  
         return true;
      }else {      
        alert("執行日期須完整填入起始及終止的年月日，不可空白");             
        return false;
      }
    }
  }
}

//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function exist1(form1) {  
  for(var i=0; i< form1.R1.length; i++) {
    if(form1.R1[i].checked == true) {      
      return true;
    }
  }
  return false;
}

function exist2(form1) { 
  if(form1.R1.checked == true) {
    return true;
  }else {
    return false;
  }
}

function exist3(form1) {  
  for(var i=0; i< form1.R2.length; i++) {
    if(form1.R2[i].checked == true) {      
      return true;
    }
  }
  return false;
}

function exist4(form1) { 
  if(form1.R2.checked == true) {
    return true;
  }else {
    return false;
  }
}
function Check_ZZ009W_LIST02(form1) {
  if ((form1.R1 == null) && (form1.R2 == null)) {
    alert("沒有要刪除或還原的使用者");
    return false;
  }else {
    if((form1.R1.length > 1) || (form1.R2.length > 1)) {      
      if((exist1(form1)==true) || (exist3(form1)==true)) {        
        return true;;
      }else { 
        alert("請選擇要刪除或還原的使用者");       
        return false;
      }      
    }else {            
      if((exist2(form1)==true) || (exist4(form1)==true)) {        
        return true;
      }else {
        alert("請選擇要刪除或還原的使用者");
        return false;        
      }
      return false;
    }
  }
}

function changeVal(T1){
     
   pos=0
   var oring=T1.value
   pos=oring.indexOf(",")

      while (pos !=-1){
         oring=(oring).replace(",","")
         pos=oring.indexOf(",")
      }   

   return(oring)     
}

function changeStr(T1){
     
   pos=0
   c="";
   var oring=T1.value
  
   if (isNaN(Math.abs(T1.value))) {
   	alert("請輸入數字");
   	return(oring);
   }else{	
     if (eval(T1.value) < 0 ) {
      c="-";
     }   	

   T1.value= Math.abs(T1.value); 
   pos=oring.indexOf(",")
   if (pos==-1) {
      a=((T1.value).length) % 3
      b=(((T1.value).length)-a) / 3
      if (b==4){
            if (a==0){
               oring=(T1.value).substring(0,3)+","+(T1.value).substring(3,6)+","+(T1.value).substring(6,9)+","+(T1.value).substring(7);            
            }
            if (a==1){
               oring=(T1.value).substring(0,1)+","+ (T1.value).substring(1,4)+","+(T1.value).substring(4,7)+","+(T1.value).substring(7,10)+","+(T1.value).substring(10);
            }                           
            if (a==2){
               oring=(T1.value).substring(0,2)+","+ (T1.value).substring(2,5)+","+(T1.value).substring(5,8)+","+(T1.value).substring(8,11)+","+(T1.value).substring(11,14);
            }                                       
      }      
      if (b==3){
         if (a==0){
            oring=(T1.value).substring(0,3)+","+(T1.value).substring(3,6)+","+(T1.value).substring(6)            
         }
         if (a==1){
            oring=(T1.value).substring(0,1)+","+ (T1.value).substring(1,4)+","+(T1.value).substring(4,7)+","+(T1.value).substring(7)
         }                           
      }
      if (b==2 ){
         if (a==0){
            oring=(T1.value).substring(0,3)+","+(T1.value).substring(3,6)
         }
         if (a==1){
            oring=(T1.value).substring(0,1)+","+ (T1.value).substring(1,4)+","+(T1.value).substring(4,7)
         }                           
         if (a==2){
            oring=(T1.value).substring(0,2)+","+ (T1.value).substring(2,5)+","+(T1.value).substring(5,8)
         }                                       
      }         
      if (b==1){
         if (a==0){
            oring=T1.value;
         }                           
         if (a==1){
            oring=(T1.value).substring(0,1)+","+ (T1.value).substring(1,4)
         }                           
         if (a==2){
            oring=(T1.value).substring(0,2)+","+ (T1.value).substring(2,5)
         }                                       
      }         
      if (b<0){
         oring=T1.value;
      }         
      
      
   }  
   if (c == "-" ) {
   	oring="-"+oring;
   }	
   return(oring)

}

}

//=========================
function checkZZ008W(form)
{
    if(form.YEAR.value.length==0)
    {
        alert("請輸入年份");
        return false;    
    }
    if(!checkYear(form.YEAR))
        return false;
    if(!fnValidDate(form.YEAR.value + form.MONTH.value + form.DAY.value))
    {
        alert('日期有誤');
        return false;
    }

    return true;
}    
//=======================================
function changeSelect(form)    
{
    if(form.IorO != null)
    {
   	 if(form.IorO[0].checked==true)
	 {
        	ChangeLength(form);
	        form.USER_NAME.disabled=false;
        	form.BANK_NO.maxLength=5;
	        if(form.selectname.value==null)
        	{    
	            for(i=0;i<form.selectname.length;i++)
        	        form.group.options[form.group.options.length] = new Option(form.selectname[i].value,form.selectid[i].value);    
	        }
        	else
	            form.group.options[form.group.options.length] = new Option(form.selectname.value,form.selectid.value);
        	if(thisBrowser=="Netscape")    
	            window.history.go(0);    
	  }     
	  else
	  {
        	ChangeUserName(form);
	        form.BANK_NO.maxLength=12;
        	for(i = form.group.options.length-1;i>=0;i--)
	            form.group.options[i] = null;
    	   }
      }
}

//===============================================================
function ChangeLength(form)
{
    if(!checkBankNo(form.BANK_NO))
        return;
    if(form.IorO[0].checked==true)
    {
        if(form.BANK_NO.value.length>=5)
            form.BANK_NO.value=(form.BANK_NO.value).substring(0,5);
    }
} 
//===============================================
function ChangeUserName(form)
{
    if(form.IorO[1].checked==true)
    {
        form.USER_NAME.disabled=true;
        form.USER_NAME.value="";
    }
}
//=====================================
function trimUserName(textfield)
{
    textfield.value=trimString(textfield.value);    
}
//=======================================
function checkBankNo(textfield)
{
    textfield.value=trimString(textfield.value);
    /*if(isNaN(Math.abs(textfield.value)))
    {
        alert("請輸入數字");
        return false;
    }
    if(eval(textfield.value)<0)
    {
        alert("不可為負值");
        return false;
    }*/
    return true;        
}
function NewcheckRate(Rate1)
{
    if(isNaN(Math.abs(Rate1.value))) 
    {
   	    alert("請輸入數字");
   	    return false;
    }
    if(eval(Rate1.value) > 99.99)
    {
        alert("利率不可大於 99.99")    
        return false;
    }
    if((Rate1.value).indexOf(".") != -1 )
    {
        len = (Rate1.value).substring((Rate1.value).indexOf(".")+1,Rate1.value.length);
        if(len.length > 2)
        {
            alert("小數點後只能有兩個位數");  
            return false;  
        }
        Rate1.value = eval(Rate1.value);
    }
    return true;
}
//================================================================
function trimString(inString) 
{
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
//======================================
function Check_WFX06W(form1) {
  if(form1.prv_yy.value.length  !=0 && form1.prv_mm.value.length  !=0 &&
     form1.prv_dd.length !=0 ) {
     if(!fnValidDate((parseInt(form1.prv_yy.value)+1911) + form1.prv_mm.value + form1.prv_dd.value)) {
        alert('核准設立日期有誤');
        return false;
     }
  }

  if(form1.setup_yy.value.length  !=0 && form1.setup_mm.value.length  !=0 &&
     form1.setup_date_d.length !=0 ) {
     if(!fnValidDate((parseInt(form1.setup_yy.value)+1911) + form1.setup_mm.value + form1.setup_dd.value)) {
        alert('設立日期有誤');
        return false;
     }
  }
  if (form1.addr.value.length !=0 ) {
     if (form1.hsiend_id.value.length ==0 || form1.area_id.value.length ==0) {	
        alert('縣市名稱及郵遞區號不可空白');
        return false;
     }   
  }

}
//====================================================================
function LinkToInsert1(form,linkaddr)
{
    linkaddr = "/servlet/" + linkaddr;
    if(form.Function.value=='Add')
    {
        alert("請先新增外國銀行在華辦事處資料");
        return;
    }
    else
        window.open(linkaddr,'_self');
}
//====================================================================

function Check_WFX06W_ShowSubR(form)
{
    if(form.corp_id.value=="")
    {
    	alert("統一編號不可空白");
        return false;
    }        
    return true;
}
function doWFX06WSubmit(form,myfun)
{
    form.Function.value = myfun;
    if(myfun=="delete")
    {
        if(AskDelete())
            form.submit();
    }        
    else    
    {
        if(Check_WFX06W_ShowSubR(form))
            form.submit();
    }
}
//==================================//
function Check_WFX04W(form1) {
  if (form1.branch_no.value.length ==0){
        alert('分支機構名稱不可空白');
        return false;  	
  }      
  if(form1.prv_yy.value.length  !=0 && form1.prv_mm.value.length  !=0 &&
     form1.prv_dd.length !=0 ) {
     if(!fnValidDate((parseInt(form1.prv_yy.value)+1911) + form1.prv_mm.value + form1.prv_dd.value)) {
        alert('核准設立日期有誤');
        return false;
     }
  }

  if(form1.open_yy.value.length  !=0 && form1.open_mm.value.length  !=0 &&
     form1.open_date_d.length !=0 ) {
     if(!fnValidDate((parseInt(form1.open_yy.value)+1911) + form1.open_mm.value + form1.open_dd.value)) {
        alert('設立日期有誤');
        return false;
     }
  }
  if (form1.addr.value.length !=0 ) {
     if (form1.hsiend_id.value.length ==0 || form1.area_id.value.length ==0) {	
        alert('縣市名稱及郵遞區號不可空白');
        return false;
     }   
  }

}
//====================================//
function LinkToInsert(form,linkaddr)
{
    linkaddr = "/servlet/" + linkaddr;
    if(form.Function.value=='Add')
    {
        alert("請先新增外國銀行在華分支機構資料");
        return;
    }
    else
        window.open(linkaddr,'_self');
}
function doWFX04WSubmit(form,myfun)
{
    form.Function.value = myfun;
    if(myfun=="delete")
    {
        if(AskDelete())
            form.submit();
    }        
    else    
    {
        if(form.code_name.value==""){
      	   alert("營業項目名稱不可空白");
           return false;
        }
            form.submit();
    }
}
//====================================================
function ChgAction(form,actionURL)
{
    form.action = actionURL;
	form.submit();
}