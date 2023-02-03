function doSubmit(form,cnd,muser_id){	     	    	    
	    form.action="/pages/ZZ001W.jsp?act="+cnd+"&test=nothing";	    
	    if(((cnd == "Insert") || (cnd == "Update") || (cnd == "Delete") || (cnd == "ResetPwd"))
	    && (!checkData(form,cnd))) return;	    	    	    
	    if(cnd == "new" || cnd == "Qry") form.submit();
	    if(cnd == "Edit"){
	       form.action="/pages/ZZ001W.jsp?act="+cnd+"&muser_id="+muser_id+"&test=nothing";	    	
	       form.submit();
	    }	
	    if((cnd == "Insert") && AskInsert(form)) form.submit();	    
	    if((cnd == "Update") && AskUpdate(form)) form.submit();	    
	    if((cnd == "Delete") && AskDelete(form)) form.submit();	    
	    if((cnd == "ResetPwd") && AskResetPwd(form)) form.submit();	    
	    
	    
}	

function getData(form,cnd,item){	     	    
	    if(item == 'bank_type'){
	       form.TBANK_NO.value="";	 
		}	
		if(item == 'tbank_no'){
	       form.BANK_NO.value="";	 
		}
	    form.action="/pages/ZZ001W.jsp?act=getData&nowact="+cnd+"&test=nothing";
	    form.submit();	    
}	
	
function checkData(form,cnd) 
{	
	if (trimString(form.MUSER_ID.value) =="" ){
		alert("使用者帳號不可空白");
		form.MUSER_ID.focus();
		return false;
	}else{	   
	    if((form.BANK_TYPE.value == "2") && (form.MUSER_ID.value.length != 10)){
	       alert("農金局局內使用者帳號，須填入10碼");	
	       form.MUSER_ID.focus();
		   return false;
	    }	
	    if((form.BANK_TYPE.value != "2") && (form.MUSER_ID.value.length != 3) && (cnd != 'Update' && cnd != 'Delete') ){
	       alert("農金局局外使用者帳號，須填入任意3碼");	
	       form.MUSER_ID.focus();
		   return false;
	    }	
	    
	}	
	
	if (trimString(form.MUSER_NAME.value) =="" ){
		alert("使用者姓名不可空白");
		form.MUSER_NAME.focus();
		return false;
	}		
		
   return true;
}

