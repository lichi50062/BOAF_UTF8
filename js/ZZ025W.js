//94.11.01 add e-mail���ˮ�
function doSubmit(form,cnd){
	    if(!checkData(form)) return;	    
	    form.action="/pages/ZZ025W.jsp?act="+cnd+"&test=nothing";
	    form.submit();	    
}	

	
function checkData(form) 
{	
	
	if (trimString(form.MUSER_NAME.value) =="" ){
		alert("�ӤH�򥻸�ƪ��m�W���i�ť�");
		form.MUSER_NAME.focus();
		return false;
	}	
	if (trimString(form.M_EMAIL.value) =="" ){
		alert("�ӤH�򥻸�ƪ��q�l�l��b�����i�ť�");
		form.M_EMAIL.focus();
		return false;
	}
	if(!CheckEmailAddress(form.M_EMAIL.value)){
		alert('�q�l�l��b����J���~');
		form.M_EMAIL.focus();
		return false;
     }	
   return true;
}