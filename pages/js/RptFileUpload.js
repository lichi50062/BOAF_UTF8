function doSubmit(form,cnd) {			 
	
	if (!checkSingleYM(form.S_YEAR, form.S_MONTH)) {
		form.S_YEAR.focus();
		return;
	}
	
	if (cnd == 'Upload' && trimString(form.UpFileName.value) == "") {
		alert("上傳檔案位置必須輸入")
		form.UpFileName.focus();
		return;
	}
		
	if (cnd == 'Upload' && confirm('確定上傳該報表檔??')) {
		form.act.value="Upload";	   
		form.action="/pages/RptFileUpload.jsp?act="+form.act.value+"&FileName="+form.UpFileName.value+"&RPT_CODE="+form.RPT_CODE.value+"&S_YEAR="+form.S_YEAR.value+"&S_MONTH="+form.S_MONTH.value+"&test=nothing";
		form.submit();
	}else if(cnd == 'UploadFeb'){
		form.act.value="UploadFeb";
		var flag = false;
		var febstr;  
  	    for (var i = 0 ; i < form.elements.length; i++) {    
    	   if ( form.elements[i].checked == true ) {
             flag = true;
             febstr = form.elements[i].value;
           }    
        }
  		if (flag == false) {       	
      		alert('請至少選擇一筆欲上傳至檢查局的報表!');        	   
      		return;
  		}
        //alert("febstr="+febstr);		
		form.action="/pages/RptFileUpload.jsp?act="+form.act.value+"&rptLine="+febstr;
		//alert("/pages/RptFileUpload.jsp?act="+form.act.value+"&rptLine="+febstr);
		form.submit(); 		 			 
	}else if (cnd == 'List'){
		form.act.value="List";	   
		form.action="/pages/RptFileUpload.jsp?act="+form.act.value+"&RPT_CODE="+form.RPT_CODE.value+"&S_YEAR="+form.S_YEAR.value+"&S_MONTH="+form.S_MONTH.value+"&test=nothing";
		form.submit();
    }else{	
		return;
	}	
}

function checkSelect(form) 
{	
  var flag = false;  
  for (var i = 0 ; i < form.elements.length; i++) {    
    if ( form.elements[i].checked == true ) {
        flag = true;
        alert(form.elements[i].value);
    }    
  }
  if (flag == false) {       	
      alert('請至少選擇一筆欲上傳至檢查局的報表!');        	   
      return false;
  }
  return true;
}



function setSelect(S1, bankid) {	
    if(S1 == null)
    	return;
    for(i=0;i<S1.length;i++) {
        
      	if(S1.options[i].value==bankid)    	{      				
        	S1.options[i].selected=true;
        	break;
    	}
    }
}
