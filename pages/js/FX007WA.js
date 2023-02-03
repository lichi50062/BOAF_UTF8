//95.05.24 fix 戶數,金額不可為負數 by 2295
//95.06.02 add 回查詢頁
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

	form.action="/pages/FX007WA.jsp?act=new&checkyear="+a+"&checkmonth="+b;	    
	form.submit();
	
}
}


function doSubmit(form,cnd){	     	  
if(cnd =='add'){
	if(!checkData(form)) return;	
	if(!submitform(form)) return;
	if(AskInsert_Permission()){
	form.action="/pages/FX007WA.jsp?act=Insert";	    
	form.submit();
  }
}
else if(cnd =='modify'){
	if(!checkData(form)) return;	
	if(!submitform(form)) return;
	if(AskEdit_Permission()){
	form.action="/pages/FX007WA.jsp?act=Update";	    
	form.submit();
	}
}
else if(cnd =='delete'){
	if(AskDelete_Permission()){
	form.action="/pages/FX007WA.jsp?act=Delete";	    
	form.submit();
	}
}
else if(cnd =='load'){
if( confirm("確定載入上月資料？")){
	form.action="/pages/FX007WA.jsp?act=Load";	    
	form.submit();
	}
}
else if(cnd =='returnList'){
if( confirm("確定回查詢頁？")){//95.06.02 add 回查詢頁
	form.action="/pages/FX007WA.jsp?act=List";	    
	form.submit();
	}
}

else return;
	  	  	    
}


function checkData(form) 
{
    if (trimString(form.bank_cnt.value) =="" ){
		alert("支票存款正會員戶數不可空白");
		form.bank_cnt.focus();
		return false;
	}else if (!checkNumber(form.bank_cnt)){
		form.bank_cnt.focus();
		return false;
	}else if(form.bank_cnt.value < 0){
	   	alert("支票存款正會員戶數不可為負數");
		form.bank_cnt.focus();
		return false;	    
    }	
    
	if (trimString(form.bank_bal.value) =="" ){
		alert("支票存款正會員餘額不可空白");
		form.bank_bal.focus();
		return false;
	}else if (!checkNumber(form.bank_bal)){
		form.bank_bal.focus();
		return false;
	}else if(form.bank_bal.value < 0){
	   	alert("支票存款正會員餘額不可為負數");
		form.bank_bal.focus();
		return false;	    
    }
    		
	if (trimString(form.bank_cnt_s.value) =="" ){
		alert("支票存款贊助會員戶數不可空白");
		form.bank_cnt_s.focus();
		return false;
	}else if (!checkNumber(form.bank_cnt_s)){
		form.bank_cnt_s.focus();
		return false;
	}else if(form.bank_cnt_s.value < 0){
	   	alert("支票存款贊助會員戶數不可為負數");
		form.bank_cnt_s.focus();
		return false;	    
    }	
	if (trimString(form.bank_bal_s.value) =="" ){
		alert("支票存款贊助會員餘額不可空白");
		form.bank_bal_s.focus();
		return false;
	}else if (!checkNumber(form.bank_bal_s)){
		form.bank_bal_s.focus();
		return false;
	}else if(form.bank_bal_s.value < 0){
	   	alert("支票存款贊助會員餘額不可為負數");
		form.bank_bal_s.focus();
		return false;	    
    }	
	if (trimString(form.bank_cnt_n.value) =="" ){
		alert("支票存款非會員戶數不可空白");
		form.bank_cnt_n.focus();
		return false;
	}else if (!checkNumber(form.bank_cnt_n)){
		form.bank_cnt_n.focus();
		return false;
	}else if(form.bank_cnt_n.value < 0){
	   	alert("支票存款非會員戶數不可為負數");
		form.bank_cnt_n.focus();
		return false;	    
    }	
	if (trimString(form.bank_bal_n.value) =="" ){
		alert("支票存款非會員餘額不可空白");
		form.bank_bal_n.focus();
		return false;
	}else if (!checkNumber(form.bank_bal_n)){
		form.bank_bal_n.focus();
		return false;
	}else if(form.bank_bal_n.value < 0){
	   	alert("支票存款非會員餘額不可為負數");
		form.bank_bal_n.focus();
		return false;	    
    }	
	  
    return true;
}

function submitform(form)
{	
    	form.bank_cnt.value = changeVal( form.bank_cnt)                    
		form.bank_bal.value = changeVal(form.bank_bal)    
		form.bank_cnt_s.value = changeVal(form.bank_cnt_s)   
		form.bank_bal_s.value = changeVal(form.bank_bal_s)                     
		form.bank_cnt_n.value = changeVal(form.bank_cnt_n)                      
		form.bank_bal_n.value = changeVal(form.bank_bal_n)                    			
		
	return true;                   
                                       
}