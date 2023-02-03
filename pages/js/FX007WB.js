//95.05.24 fix 戶數,金額不可為負數 by 2495
var act;
var checkLock = new Array()

function pushArray(lock)
{
checkLock.push(lock)
}
function popArray(form)
{
	for(i=0;i<checkLock.length;i+=2)
	{
	 if(form.S_YEAR.value==checkLock[i] && form.S_MONTH.value==checkLock[i+1])
	 {
	 	alert("你鍵入的申報年月巳被鎖定!");
	 	return false;
	 }
	}
	return true;
}

function AskDelete_Permission() {
  if(confirm("確定要刪除此筆資料嗎？"))
    return true;
  else
    return false;
}

function AskInsert_Permission() {
  if(confirm("確定要新增此筆資料嗎？"))
    return true;
  else
    return false;
}
function AskEdit_Permission() {
  if(confirm("確定要修改此筆資料嗎？"))
    return true;
  else
    return false;
}


function newSubmit(form,cnd,lockdate){	     	  
if(cnd == 'new'){
var checkDate=new Date()
var myDate=new Date()

var a = form.S_YEAR.value;
var b = form.S_MONTH.value;
	if (trimString(a) =="" ){
			alert("年份不可空白");
			form.S_YEAR.focus();
			return;
	}
	if (trimString(b) =="" ){
		alert("月份不可空白");
		form.S_MONTH.focus();
		return;
	}
if(!popArray(form)) return;		

//remeber to cast type 2005.12.23 by lilic0c0 4183
var selectDate = parseInt(a,10)*12+parseInt(b,10); 

if(!popArray(form)) return;	

if(lockdate > selectDate)
{
	alert("你鍵入的申報年月超出申報項目起始年月");	
	return;	
}

var year=parseInt(a,10)+1911;
var month=b-1;
	checkDate.setFullYear(checkDate.getYear(),checkDate.getMonth()-1,31)    
	myDate = new Date()
	myDate.setFullYear(year,month,myDate.getDate())

	if(myDate>checkDate)
	{
		alert("申報年月超過可申報年月!");	
	return;
	}

	form.action="/pages/FX007WB.jsp?act=new&checkyear="+a+"&checkmonth="+b;	    
	form.submit();
	
}
}






//add by 2295 若不可有小數點時,則停在該text field
//add by 2495 本月新增戶數與上月累積戶數相加
function checkTotal_house(Rate1,ini,flag_house){
    
	//93.03.16拿掉Rate1.value = changeVal(Rate1);  
	  if(Rate1.creditmonth_cnt.value=="")
	  	 Rate1.creditmonth_cnt.value=0;
    if(!isFinite(Rate1.creditmonth_cnt.value))
    {
    	 alert("請輸入數字!!");
	  	 Rate1.creditmonth_cnt.value=0;
    }
    var num1=parseInt(Rate1.creditmonth_cnt.value);
    var total_house=ini+num1;        
    
    if(flag_house==1)
    	Rate1.credityear_cnt_acc.value=Rate1.credityear_cnt_acc.value;
    else	
    	Rate1.credityear_cnt_acc.value=total_house.toString(10);
    /*
    if(Rate1.credityear_cnt_acc.value==0)
    {
    	Rate1.creditmonth_cnt.value=0;
    	Rate1.credityear_cnt_acc.value=0;
    } 
    */   
    return true;
}

function checkTotal_cash(Rate1,ini,flag_cash){

	//93.03.16拿掉Rate1.value = changeVal(Rate1);
    if(Rate1.creditmonth_amt.value=="")
	  	 Rate1.creditmonth_amt.value=0;
    if(!isFinite(Rate1.creditmonth_amt.value))
    {
    	 alert("請輸入數字!!");
	  	 Rate1.creditmonth_amt.value=0;
    }
    var num1=parseInt(Rate1.creditmonth_amt.value);
    var total_cash = num1+ini;     
    if(flag_cash==1)
    	Rate1.credityear_amt_acc.value=Rate1.credityear_amt_acc.value;
    else      
    	Rate1.credityear_amt_acc.value=total_cash.toString(10);
    /*
    if(Rate1.credityear_amt_acc.value==0)
    {
    	Rate1.creditmonth_amt.value=0
    	Rate1.credityear_amt_acc.value=0;
    }
    */
    return true;
}




function doSubmit(form,cnd,inihouse,inicash,flag_house,flag_cash){	     	  
if(cnd =='add'){
	
	if(!checkData(form)) return;	
	if(!submitform(form)) return;
      
	if(AskInsert_Permission()){
		var num1=parseInt(form.creditmonth_cnt.value);
    var total_house=inihouse+num1;
    if(flag_house==1)
    	form.credityear_cnt_acc.value=form.credityear_cnt_acc.value;
    else
    	form.credityear_cnt_acc.value=total_house.toString(10);
    
    var num2=parseInt(form.creditmonth_amt.value);
    var total_cash = inicash+num2;
    if(flag_cash==1)
    	form.credityear_amt_acc.value=form.credityear_amt_acc.value;
    else	
    	form.credityear_amt_acc.value=total_cash.toString(10);
      
		form.action="/pages/FX007WB.jsp?act=Insert";	    
		form.submit();
  }
}
else if(cnd =='modify'){
	
	if(!checkData(form)) return;	
	if(!submitform(form)) return;	
	
	if(AskEdit_Permission()){
	
	form.action="/pages/FX007WB.jsp?act=Update&creditmonth_cnt="+form.creditmonth_cnt.value+"&creditmonth_amt="+form.creditmonth_amt.value+"&credit_bal="+form.credit_bal.value+"&credit_cnt="+form.credit_cnt.value+"&overcreditmonth_cnt="+form.overcreditmonth_cnt.value+"&overcreditmonth_amt="+form.overcreditmonth_amt.value+"&overcredit_cnt="+form.overcredit_cnt.value+"&overcredit_bal="+form.overcredit_bal.value+"&creditmonth_cnt="+form.creditmonth_cnt.value+"&creditmonth_amt="+form.creditmonth_amt.value;	    
	form.submit();
	}
}
else if(cnd =='delete'){
	if(AskDelete_Permission()){
	form.action="/pages/FX007WB.jsp?act=Delete";	    
	form.submit();
	}
}
else if(cnd =='load'){
if( confirm("確定載入上月資料？")){
	form.action="/pages/FX007WB.jsp?act=Load&LoadFlag=TRUE";	    
	form.submit();
	}
}
else if(cnd =='returnList'){
if( confirm("確定回查詢頁？")){
	form.action="/pages/FX007WB.jsp?act=List";	    
	form.submit();
	}
}


else return;
	  	  	    
}






function checkData(form) 
{
    if (form.creditmonth_cnt.value =="" ){
		alert("本月新增貸放戶數不可空白");
		form.creditmonth_cnt.focus();
		return false;
	}
	
	else	if (form.creditmonth_amt.value =="" ){
		alert("本月新增貸放金額不可空白");
		form.creditmonth_amt.focus();
		return false;
	}
	
	
	else if (form.credit_bal.value=="" ){
	
		alert("貸放餘額餘額不可空白");
		form.credit_bal.focus();
		return false;
	}
	else if (form.credit_cnt.value=="" ){
		alert("貸放餘額戶數不可空白");
		form.credit_cnt.focus();
		return false;
	}
	
	else if (form.overcreditmonth_cnt.value=="" ){
	
		alert("本月新增逾放戶數不可空白");
		form.overcreditmonth_cnt.focus();
		return false;
	}	
	else if (form.overcreditmonth_amt.value =="" ){
	
		alert("本月新增逾放金額不可空白");
		form.overcreditmonth_amt.focus();
		return false;
	}
	
		else if (form.overcredit_cnt.value=="" ){
	
		alert("逾放餘額戶數不可空白");
		form.overcredit_cnt.focus();
		return false;
	}
	
	else if (form.overcredit_bal.value=="" ){
		alert("逾放餘額不可空白");
		form.overcredit_bal.focus();
		return false;
	}
	
	else if (form.credityear_cnt_acc.value=="" ){
	
		alert("本年累計戶數不可空白");
		form.credityear_cnt_acc.focus();
		return false;
	}
	
	else if (form.credityear_amt_acc.value=="" ){
		alert("本年累計餘額不可空白");
		form.credityear_amt_acc.focus();
		return false;
	}
	else if (form.creditmonth_avgrate.value=="" ){
		alert("本月新增平均利率不可空白");
		form.creditmonth_avgrate.focus();
		return false;
	}	
	else if (form.credityear_avgrate.value=="" ){
		alert("本年累計平均利率不可空白");
		form.credityear_avgrate.focus();
		return false;	
	}
	if(!isFinite(form.creditmonth_avgrate.value))
    {
    	alert("請輸入數字!!");
	  	form.creditmonth_avgrate.value="";
	  	form.creditmonth_avgrate.focus();
	  	return false;
	  	 
    }
	
	if(!isFinite(form.credityear_avgrate.value))
    {
    	alert("請輸入數字!!");
	  	form.credityear_avgrate.value="";
	  	form.credityear_avgrate.focus();
	  	return false;
	  	 
    }
    return true;
}

function submitform(form)
{			
    form.creditmonth_cnt.value = changeVal( form.creditmonth_cnt)   
	form.creditmonth_amt.value = changeVal(form.creditmonth_amt)	
	form.credityear_cnt_acc.value = changeVal(form.credityear_cnt_acc) 	      
	form.credityear_amt_acc.value = changeVal(form.credityear_amt_acc) 	      
					 	   			
	form.credit_bal.value = changeVal(form.credit_bal) 							
	form.credit_cnt.value = changeVal(form.credit_cnt) 				 
	form.overcreditmonth_cnt.value = changeVal( form.overcreditmonth_cnt)
	form.overcreditmonth_amt.value = changeVal(form.overcreditmonth_amt) 		                    
	form.overcredit_cnt.value = changeVal(form.overcredit_cnt)
	form.overcredit_bal.value = changeVal(form.overcredit_bal) 		
	
	form.creditmonth_avgrate.value = changeVal(form.creditmonth_avgrate)
	form.credityear_avgrate.value = changeVal(form.credityear_avgrate) 		        		        						
	        		        						
	return true;                   
                                       
}