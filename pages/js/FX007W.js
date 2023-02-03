var checkData1; //貸放資料次數累計
var checkData2; //貸放資料金額本年累計
var checkData3; //逾放資料次數累計
var checkData4; //逾放資料金額本年累計
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
function pushData(v1,v2,v3,v4,v5)
{
	checkData1=v1;
	checkData2=v2;
	checkData3=v3;
	checkData4=v4;
	act=v5; //紀錄是否為首次新增
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

	form.action="/pages/FX007W.jsp?act=new&checkyear="+a+"&checkmonth="+b;	    
	form.submit();
	
}
}


function doSubmit(form,cnd){	     	  
if(cnd =='add'){
	if(!checkData(form)) return;	
	if(!submitform(form)) return;
	if(AskInsert_Permission()){
	form.action="/pages/FX007W.jsp?act=Insert";	    
	form.submit();
  }
}
else if(cnd =='modify'){
	if(!checkData(form)) return;	
	if(!submitform(form)) return;
	if(AskEdit_Permission()){
	form.action="/pages/FX007W.jsp?act=Update";	    
	form.submit();
	}
}
else if(cnd =='delete'){
	if(AskDelete_Permission()){
	form.action="/pages/FX007W.jsp?act=Delete";	    
	form.submit();
	}
}
else if(cnd =='load'){
if( confirm("確定載入上月資料？")){
	form.action="/pages/FX007W.jsp?act=Load";	    
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
	}
	else if (trimString(form.bank_bal.value) =="" ){
		alert("支票存款正會員餘額不可空白");
		form.bank_bal.focus();
		return false;
	}
	else if (trimString(form.bank_cnt_s.value) =="" ){
		alert("支票存款贊助會員戶數不可空白");
		form.bank_cnt_s.focus();
		return false;
	}
	else if (trimString(form.bank_bal_s.value) =="" ){
		alert("支票存款贊助會員餘額不可空白");
		form.bank_bal_s.focus();
		return false;
	}
	else if (trimString(form.bank_cnt_n.value) =="" ){
		alert("支票存款非會員戶數不可空白");
		form.bank_cnt_n.focus();
		return false;
	}
	else if (trimString(form.bank_bal_n.value) =="" ){
		alert("支票存款非會員餘額不可空白");
		form.bank_bal_n.focus();
		return false;
	}
	else	if (trimString(form.month_cnt.value) =="" ){
		alert("本月新增貸放戶數不可空白");
		form.month_cnt.focus();
		return false;
	}
	else	if (trimString(form.month_cnt_acc.value) =="" ){
		alert("累積貸放戶數不可空白");
		form.month_cnt_acc.focus();
		return false;
	}
	else if (trimString(form.month_amt.value) =="" ){
		alert("本月新增貸放金額不可空白");
		form.month_amt.focus();
		return false;
	}
	else if (trimString(form.month_amt_acc.value)=="" ){
	
		alert("累積貸放總額不可空白");
		form.month_amt_acc.focus();
		return false;
	}
	else if (trimString(form.credit_bal.value)=="" ){
	
		alert("貸放餘額不可空白");
		form.credit_bal.focus();
		return false;
	}
	else if (trimString(form.overcreditmonth_cnt.value)=="" ){
	
		alert("本月新增逾放戶數不可空白");
		form.overcreditmonth_cnt.focus();
		return false;
	}
		else if (trimString(form.overcreditmonth_cnt_acc.value)=="" ){
	
		alert("累積逾放戶數不可空白");
		form.overcreditmonth_cnt_acc.focus();
		return false;
	}
		else if (trimString(form.overcreditmonth_amt.value )=="" ){
	
		alert("本月新增逾放總額不可空白");
		form.overcreditmonth_amt.focus();
		return false;
	}else if (trimString(form.overcreditmonth_amt_acc.value)=="" ){
	
		alert("累積逾放總額不可空白");
		form.overcreditmonth_amt_acc.focus();
		return false;
	}else if (trimString(form.overcredit_bal.value)=="" ){
		alert("逾放餘額不可空白");
		form.overcredit_bal.focus();
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
		form.month_cnt.value = changeVal(form.month_cnt)                   
		form.month_cnt_acc.value = changeVal(form.month_cnt_acc)                  
		form.month_amt.value = changeVal(form.month_amt)                   
		form.month_amt_acc.value = changeVal(form.month_amt_acc)                         
       
        form.credit_bal.value = changeVal(form.credit_bal)
		form.overcreditmonth_cnt.value = changeVal(form.overcreditmonth_cnt)
		form.overcreditmonth_cnt_acc.value = changeVal(form.overcreditmonth_cnt_acc)
		form.overcreditmonth_amt.value = changeVal(form.overcreditmonth_amt)
		form.overcreditmonth_amt_acc.value = changeVal(form.overcreditmonth_amt_acc)
		form.overcredit_bal.value = changeVal(form.overcredit_bal)       
		 
			
		var mcnt = parseInt(form.month_cnt.value,10);
		var mamt= parseInt(form.month_amt.value,10);
		var ycnt = parseInt(form.month_cnt_acc.value,10);
		var yamt= parseInt(form.month_amt_acc.value,10);
		
		var overmcnt = parseInt(form.overcreditmonth_cnt.value,10);
		var overmamt= parseInt(form.overcreditmonth_amt.value,10);
		var overycnt = parseInt(form.overcreditmonth_cnt_acc.value,10);
		var overyamt= parseInt(form.overcreditmonth_amt_acc.value,10);
		
		
//檢查本月新增及上月累積是否為累積
		if(trimString(act) != "false")//非首次新增要做判斷
		{
			if (parseInt(checkData1)+mcnt!=ycnt){
				alert("『本月新增貸放戶數』加上『上一個月的累積貸放戶數』\n不等於『累積貸放戶數』");
				form.month_cnt.focus();
				return false;
			}
	    	if (parseInt(checkData2)+mamt!=yamt){
				alert("『本月新增貸放金額』加上『上一個月的累積貸放總額』\n不等於『累積貸放總額』");
				form.month_amt.focus();
				return false;
			}
			if (parseInt(checkData3)+overmcnt!=overycnt){
				alert("『本月新增逾放戶數』加上『上一個月的累積逾放戶數』\n不等於『累積逾放戶數』");
				form.overcreditmonth_cnt.focus();
				return false;
			}
			if (parseInt(checkData4)+overmamt!=overyamt){
				alert("『本月新增逾放金額』加上『上一個月的累積逾放總額』\n不等於『累積逾放總額』");
				form.overcreditmonth_amt.focus();
				return false;
			}
		}
                     
	return true;                   
                                       
}