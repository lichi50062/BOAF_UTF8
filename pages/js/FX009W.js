//951002 FIX BY 2495 不能昇報
//103.01.10 fix 申報年季超過可申報年季檢核規則 by 2295
var checkLock = new Array()

function pushArray(lock)
{
checkLock.push(lock)
}

function popArray(form)
{
	for(i=0;i<checkLock.length;i+=2)
	{
	 if(form.S_YEAR.value==checkLock[i] && form.S_QUARTER.value==checkLock[i+1])
	 {
	 	alert("你鍵入的申報年季巳被鎖定!");
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

function printSubmit(form,cnd,banktype,now_bank_no,muser_id){
	  
		
	if(muser_id=='A111111111' || banktype=='B'){
		
		var a = form.S_YEAR.value;		
		var b = form.S_QUARTER.value;	 	
		form.action="/pages/FX009W_Print_B.jsp?act=new&S_YEAR="+a+"&S_MONTH="+b+"&bank_no="+now_bank_no;	 	    
		form.submit();	
		return;	
	}
	if(banktype=='2'){
			 
		var a = form.S_YEAR.value;		
		var b = form.S_QUARTER.value;	 	
		form.action="/pages/FX009W_Print_B.jsp?act=new&S_YEAR="+a+"&S_MONTH="+b+"&bank_no="+now_bank_no;	    
		form.submit();	
		return;		
	}
	
	
	if( banktype=='6')
	{	
			
		var a = form.S_YEAR.value;
		var b = form.S_QUARTER.value;	
		form.action="/pages/FX009W_Print.jsp?act=new&S_YEAR="+a+"&S_MONTH="+b+"&bank_type="+banktype+"&bank_no="+now_bank_no;	    
		form.submit();
		return;	
	}
	if(banktype=='7')
	{		
		
		var a = form.S_YEAR.value;
		var b = form.S_QUARTER.value;	 
		form.action="/pages/FX009W_Print.jsp?act=new&S_YEAR="+a+"&S_MONTH="+b+"&bank_type="+banktype+"&bank_no="+now_bank_no;	    
		form.submit();
		return;	
	}
	
	
	
}

function printSubmitFX009WAll(form,cnd,banktype,now_bank_no){
	 	
	 	var a = form.S_YEAR.value;
		var b = form.S_QUARTER.value;
		
		if(banktype=='6'){
			 form.action="/pages/FX009W_Print_All.jsp?act=download&S_YEAR="+a+"&S_MONTH="+b+"&bank_type="+banktype+"&unit=1";	    			 
			 form.submit();
		}	
		if(banktype=='7'){
			 form.action="/pages/FX009W_Print_All.jsp?act=download&S_YEAR="+a+"&S_MONTH="+b+"&bank_type="+banktype+"&unit=1";	    			 
			 form.submit();
		}		
}

function printSubmitB(form,cnd,banktype,now_bank_no){
	 	  
	 	var a = form.S_YEAR.value;
		var b = form.S_QUARTER.value;
		if(banktype=='B'){
			form.action="/pages/FX008W_Print_B.jsp?act=new&S_YEAR="+a+"&S_MONTH="+b;	    
			form.submit();
		}	
		if(banktype=='2'||banktype=='6'){ 
			form.action="/pages/FX008W_Print_2.jsp?act=new&S_YEAR="+a+"&S_MONTH="+b+"&bank_no="+now_bank_no;	    
			form.submit();
		}	
	
		
}



//103.01.10 fix 申報年季超過可申報年季檢核規則
function newSubmit(form,cnd,lockdate){	     	  
   if(cnd == 'new'){
      var checkDate=new Date()
      var myDate=new Date()
      var checkyear;
      var checkquarter;
      var a = form.S_YEAR.value;
      var b = form.S_QUARTER.value;
      if (trimString(a) =="" ){
      		alert("年份不可空白");
      		form.S_YEAR.focus();
      		return;
      }
      if (trimString(b) =="" ){
      	alert("季不可空白");
      	form.S_QUARTER.focus();
      	return;
      }
      if(!popArray(form)) return;	
      
      //remeber to cast type 2005.12.23 by lilic0c0 4183
      var selectDate = parseInt(a,10)*12+parseInt(b,10); 
      
      
      //alert(lockdate+" "+selectDate);
      
      if(!popArray(form)) return;	
      
      if(lockdate > selectDate)
      {
      	
      	alert("你鍵入的申報年季超出申報項目起始年季");	
      	return;	
      }
      
      var year=parseInt(a,10)+1911;
      var quarter=parseInt(b,10);
      
      //alert('目前己選的.year='+year);
      //alert('目前己選的.quarter='+quarter);
      //alert('系統日期.checkDate.getYear()='+checkDate.getYear());
      //alert('系統日期.checkDate.getMonth()='+checkDate.getMonth());
      
      checkyear =checkDate.getYear();
      
      //951002 FIX BY 2495 不能昇報
      if(checkDate.getMonth()<=2){
      	checkquarter=4;
      	checkyear = checkyear - 1;
      }	
      else if(checkDate.getMonth()<=5)checkquarter=1;
      else if(checkDate.getMonth()<=8)checkquarter=2;
      else if(checkDate.getMonth()<=11)checkquarter=3;
      
      //alert('year.quarter='+(year * 100 + quarter));
      //alert('checkyear.quarter='+(checkyear * 100 + checkquarter));
      if((year * 100 + quarter) > (checkyear * 100 + checkquarter)){//103.01.10 fix 申報年季超過可申報年季檢核規則 
      	//alert('year.quarter='+(year * 100 + quarter));
      	//alert('checkyear.quarter='+(checkyear * 100 + checkquarter));
      	alert("申報年季超過可申報年季!");	
      	return;
      }
      
      form.action="/pages/FX009W.jsp?act=new&checkyear="+a+"&checkquarter="+quarter;	    
      form.submit();
      	
   }
}


function doSubmit(form,cnd){	     	  
if(cnd =='add'){
	if(!checkData(form)) return;	
	if(!submitform(form)) return;
	if(AskInsert_Permission()){
	form.action="/pages/FX009W.jsp?act=Insert";	    
	form.submit();
  }
}
else if(cnd =='modify'){
	if(!checkData(form)) return;	
	if(!submitform(form)) return;
	if(AskEdit_Permission()){
	form.action="/pages/FX009W.jsp?act=Update";	    
	form.submit();
	}
}
else if(cnd =='delete'){
	if(AskDelete_Permission()){
	form.action="/pages/FX009W.jsp?act=Delete";	    
	form.submit();
	}
}
else if(cnd =='load'){
if( confirm("確定載入上季資料？")){
	form.action="/pages/FX009W.jsp?act=Load";	    
	form.submit();
	}
}
else return;
	  	  	    
}

function checkData(form) 
{
if (trimString(form.warnaccount_tcnt.value) =="" ){
		alert("警示帳戶總戶數不可空白");
		form.warnaccount_tcnt.focus();
		return false;
	}
	else if (trimString(form.warnaccount_tbal.value) =="" ){
		alert("警示帳戶總餘額不可空白");
		form.warnaccount_tbal.focus();
		return false;
	}
	else if (trimString(form.warnaccount_remit_tcnt.value) =="" ){
		alert("警示帳戶內所匯（轉）入總筆數不可空白");
		form.warnaccount_remit_tcnt.focus();
		return false;
	}
	else if (trimString(form.warnaccount_refund_apply_cnt.value) =="" ){
		alert("警示帳戶內申請退還戶數不可空白");
		form.warnaccount_refund_apply_cnt.focus();
		return false;
	}
	else if (trimString(form.warnaccount_refund_apply_amt.value) =="" ){
		alert("警示帳戶內申請退還金額不可空白");
		form.warnaccount_refund_apply_amt.focus();
		return false;
	}
	else if (trimString(form.warnaccount_refund_cnt.value) =="" ){
		alert("警示帳戶內已辦理退還戶數不可空白");
		form.warnaccount_refund_cnt.focus();
		return false;
	}
	else	if (trimString(form.warnaccount_refund_amt.value) =="" ){
		alert("警示帳戶內已辦理退還金額不可空白");
		form.warnaccount_refund_amt.focus();
		return false;
	}


    return true;
}

function submitform(form)
{
    form.warnaccount_tcnt.value= changeVal( form.warnaccount_tcnt)                    
		form.warnaccount_tbal.value = changeVal(form.warnaccount_tbal)    
		form.warnaccount_remit_tcnt.value = changeVal(form.warnaccount_remit_tcnt)   
		form.warnaccount_refund_apply_cnt.value = changeVal(form.warnaccount_refund_apply_cnt)                     
		form.warnaccount_refund_apply_amt.value = changeVal(form.warnaccount_refund_apply_amt)                      
		form.warnaccount_refund_cnt.value = changeVal(form.warnaccount_refund_cnt)                    
		form.warnaccount_refund_amt.value = changeVal(form.warnaccount_refund_amt)                   
	  return true;                   
                                       
}