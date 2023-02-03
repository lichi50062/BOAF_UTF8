function doSubmit(form,cnd){
	    if(!checkData(form)) return;	    
	    form.action="/pages/FX004W.jsp?act="+cnd+"&test=nothing";	    
	    if((cnd == "Update") && AskUpdate(form)) form.submit();	    	    
}	

	
function checkData(form) 
{	
	if (trimString(form.BUSINESS_PERSON.value) =="" ){
		alert("實際辦理農漁業金融業務人員不可空白");
		form.BUSINESS_PERSON.focus();
		return false;
	}else if(isNaN(Math.abs(form.BUSINESS_PERSON.value))){
           alert("實際辦理農漁業金融業務人員不可為文字");
           form.BUSINESS_PERSON.focus();
           return false;    
	}	
	
	if (trimString(form.M_NAME.value) =="" ){
		alert("主要連絡人-姓名不可空白");
		form.M_NAME.focus();
		return false;
	}	
	
	if (trimString(form.M_TELNO.value) =="" ){
		alert("主要連絡人-電話不可空白");
		form.M_TELNO.focus();
		return false;
	}	
	
	if (trimString(form.M_EMAIL.value) =="" ){
		alert("主要連絡人-電子郵件帳號不可空白");
		form.M_EMAIL.focus();
		return false;
	}	
	
   return true;
}