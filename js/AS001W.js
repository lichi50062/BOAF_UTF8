function doSubmit(form,cnd,bank_no){	     	    	    
	    form.action="/pages/AS001W.jsp?act="+cnd+"&test=nothing";	    
	    if(((cnd == "Insert") || (cnd == "Update") || (cnd == "Delete"))
	    && (!checkData(form,cnd))) return;	    	    	    
	    if(cnd == "new" || cnd == "Qry") form.submit();
	    if(cnd == "Edit"){
	       form.action="/pages/AS001W.jsp?act="+cnd+"&TBANK_NO="+bank_no+"&test=nothing";	    	
	       form.submit();
	    }	
	    if((cnd == "Insert") && AskInsert(form)) form.submit();	    
	    if((cnd == "Update") && AskUpdate(form)) form.submit();	    
	    if((cnd == "Delete") && AskDelete(form)) form.submit();	    
}	
	
function checkData(form,cnd) 
{	
	if (trimString(form.BANK_NO.value) =="" ){
		alert("���c�N�X���i�ť�");
		form.BANK_NO.focus();
		return false;
	}else{	   
	    if((form.BANK_TYPE.value == "6" || form.BANK_TYPE.value == "7") && (form.BANK_NO.value.length != 7)){
	       alert("�A(��)�|���c�N�X�A����J7�X");	
	       form.BANK_NO.focus();
		   return false;
	    }	
	    if(((form.BANK_TYPE.value != "6") && (form.BANK_TYPE.value != "7")) && (form.BANK_NO.value.length < 3)){
	       alert("���c�N�X�A�ܤֶ���J3�X");	
	       form.BANK_NO.focus();
		   return false;
	    }	
	    
	}	
	
	if (trimString(form.BANK_NAME.value) =="" ){
		alert("���c�W�٤��i�ť�");
		form.BANK_NAME.focus();
		return false;
	}		
	
	if (trimString(form.BANK_B_NAME.value) =="" ){
		alert("���c²�٤��i�ť�");
		form.BANK_B_NAME.focus();
		return false;
	}
		
	if((form.BANK_TYPE.value == "6" || form.BANK_TYPE.value == "7") && (form.EXCHANGE_NO.value.length != 7)){
	    alert("�A(��)�q�ץN�X�A����J7�X");	
	    form.EXCHANGE_NO.focus();
		return false;
	}		
		
   return true;
}
function sameBank_Name(form){
  form.BANK_B_NAME.value = form.BANK_NAME.value;
}