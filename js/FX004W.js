function doSubmit(form,cnd){
	    if(!checkData(form)) return;	    
	    form.action="/pages/FX004W.jsp?act="+cnd+"&test=nothing";	    
	    if((cnd == "Update") && AskUpdate(form)) form.submit();	    	    
}	

	
function checkData(form) 
{	
	if (trimString(form.BUSINESS_PERSON.value) =="" ){
		alert("��ڿ�z�A���~���ķ~�ȤH�����i�ť�");
		form.BUSINESS_PERSON.focus();
		return false;
	}else if(isNaN(Math.abs(form.BUSINESS_PERSON.value))){
           alert("��ڿ�z�A���~���ķ~�ȤH�����i����r");
           form.BUSINESS_PERSON.focus();
           return false;    
	}	
	
	if (trimString(form.M_NAME.value) =="" ){
		alert("�D�n�s���H-�m�W���i�ť�");
		form.M_NAME.focus();
		return false;
	}	
	
	if (trimString(form.M_TELNO.value) =="" ){
		alert("�D�n�s���H-�q�ܤ��i�ť�");
		form.M_TELNO.focus();
		return false;
	}	
	
	if (trimString(form.M_EMAIL.value) =="" ){
		alert("�D�n�s���H-�q�l�l��b�����i�ť�");
		form.M_EMAIL.focus();
		return false;
	}	
	
   return true;
}