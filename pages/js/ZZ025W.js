//94.11.01 add e-mail的檢核
function doSubmit(form,cnd){
	    if(!checkData(form)) return;	    
	    form.action="/pages/ZZ025W.jsp?act="+cnd+"&test=nothing";
	    form.submit();	    
}	

	
function checkData(form) 
{	
	
	if (trimString(form.MUSER_NAME.value) =="" ){
		alert("個人基本資料的姓名不可空白");
		form.MUSER_NAME.focus();
		return false;
	}	
	if (trimString(form.M_EMAIL.value) =="" ){
		alert("個人基本資料的電子郵件帳號不可空白");
		form.M_EMAIL.focus();
		return false;
	}
	
	if(!CheckEmailAddress(form.M_EMAIL.value)){
		alert('電子郵件帳號輸入錯誤');
		form.M_EMAIL.focus();
		return false;
     }
    
   return true;
}