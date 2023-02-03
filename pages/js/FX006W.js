//97.07.09 fix 結束日期.可不輸入 by 2295
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




function doSubmit(form,cnd){	     	  
if(cnd =='add'){
	 if(!checkData(form)) return;	
	if(AskInsert_Permission()){
	form.action="/pages/FX006W.jsp?act=Insert";	    
	form.submit();
  }
}
else if(cnd =='modify'){
	 if(!checkData(form)) return;	
	if(AskEdit_Permission()){
	form.action="/pages/FX006W.jsp?act=Update";	    
	form.submit();
	}
}
else if(cnd =='delete'){
	if(AskDelete_Permission()){
	form.action="/pages/FX006W.jsp?act=Delete";	    
	form.submit();
	}
}

else return;
	  	  	    
}

//97.07.09 fix 結束日期,可為空白 by 2295
function checkData(form) 
{
	 var ckDate;
if (trimString(form.companyname.value) =="" ){
		alert("受託債權催收機構名稱不可空白");
		form.companyname.focus();
		return false;
	}
	else if (trimString(form.contractname.value) =="" ){
		alert("受託債權催收機構聯絡人不可空白");
		form.contractname.focus();
		return false;
	}
	else if (trimString(form.contracttel.value) =="" ){
		alert("受託債權催收機構聯絡電話不可空白");
		form.usedebitcard.focus();
		return false;
	}
	else if (trimString(form.complainname.value) =="" ){
		alert("信用部申訴窗口聯絡人不可空白");
		form.complainname.focus();
		return false;
	}
	else if (trimString(form.complaintel.value) =="" ){
		alert("信用部申訴窗口專線電話不可空白");
		form.complaintel.focus();
		return false;
	}

	if(trimString(form.BEG_DATE_Y.value)  != "" ){        
    	if(isNaN(Math.abs(form.BEG_DATE_Y.value))){
           alert("起始日期(年)不可輸入文字");    
		   form.BEG_DATE_Y.focus();            
           return false;
        }
    }else{
		alert("起始日期(年)不可空白");
		form.BEG_DATE_Y.focus();
		return false;   
	}   
    if (trimString(form.BEG_DATE_M.value) == "" ){
		alert("起始日期(月)不可空白");
		form.BEG_DATE_M.focus();
		return false;
	}			
	if (trimString(form.BEG_DATE_D.value) == "" ){
		alert("起始日期(日)不可空白");
		form.BEG_DATE_D.focus();		
		return false;
	}	
    
    ckDate = '' + (parseInt(form.BEG_DATE_Y.value)+1911) + '/' + form.BEG_DATE_M.value + '/' + form.BEG_DATE_D.value;
    var sDate=new Date()   
    var eDate=new Date()     
    sDate.setFullYear(parseInt(form.BEG_DATE_Y.value)+1911,parseInt(form.BEG_DATE_M.value)-1,parseInt(form.BEG_DATE_D.value));  
    if( fnValidDate(ckDate) != true){
    	alert('起始日期為無效日期!!');
    	form.BEG_DATE_D.focus();
    	return false;
    }    
    form.BEG_DATE.value = ckDate;   	
    
    
    if (trimString(form.END_DATE_Y.value)  != "" ){        
    	if(isNaN(Math.abs(form.END_DATE_Y.value))){
           alert("結束日期(年)不可輸入文字");    
		   form.END_DATE_Y.focus();            
           return false;
        }
    }/*else{
		alert("結束日期(年)不可空白");
		form.END_DATE_Y.focus();
		return false;   
	}   
    if (trimString(form.END_DATE_M.value) == "" ){
		alert("結束日期(月)不可空白");
		form.END_DATE_M.focus();
		return false;
	}			
	if (trimString(form.END_DATE_D.value) == "" ){
		alert("結束日期(日)不可空白");
		form.END_DATE_D.focus();		
		return false;
	}	
	*/
	if(trimString(form.END_DATE_Y.value)  != "" && trimString(form.END_DATE_M.value) != "" && trimString(form.END_DATE_D.value) != ""){
      ckDate = '' + (parseInt(form.END_DATE_Y.value)+1911) + '/' + form.END_DATE_M.value + '/' + form.END_DATE_D.value;
      eDate.setFullYear(parseInt(form.END_DATE_Y.value)+1911,parseInt(form.END_DATE_M.value)-1,parseInt(form.END_DATE_D.value)) ;    
    	if( sDate>eDate){
        	alert('起迄時間有誤!');
        	form.END_DATE_D.focus();
        	return false;
    	}  
    	if( fnValidDate(ckDate) != true){
        	alert('結束日期為無效日期!!');
        	form.END_DATE_D.focus();
        	return false;
    	}    
    	form.END_DATE.value = ckDate;   	
    }
    return true;
}

