//94.05.20 add 加上bank_type為參數
//101.07.19 add M106/M201/M206上傳檔名檢核 ex:M10610107.csv(共13碼)
function doSubmit(form) {			 
	
	if (!checkSingleYM(form.S_YEAR, form.S_MONTH)) {
		form.S_YEAR.focus();
		return;
	}
	
	//93.12.06 上傳檔名檢核 
	var FILE_NAME_LENGTH = 15; 	//A01002000009105
	if (form.Report_no.value.substring(0,1)=='M') FILE_NAME_LENGTH = 8;
	//101.07.19 add M106/M201/M206上傳檔名檢核 ex:M10610107.csv(共13碼)
	if (form.Report_no.value =='M106' || form.Report_no.value=='M201' || form.Report_no.value=='M206'){
		  FILE_NAME_LENGTH = 13;		  
	}	  
	
	var YM = ('000' + form.S_YEAR.value).substring(form.S_YEAR.value.length) + form.S_MONTH.value;	
	var fileName;    
	if      (form.Report_no.value.substring(0,1)=='A') fileName = form.Report_no.value + form.bank_code.value + YM;	
  else if (form.Report_no.value =='M106' || form.Report_no.value=='M201' || form.Report_no.value=='M206'){
  	 fileName = form.Report_no.value + YM + ".csv";	  	
  }else if (form.Report_no.value.substring(0,1)=='M') fileName = form.Report_no.value + YM;	
  	
	var i = form.UpFileName.value.length - FILE_NAME_LENGTH;
	
	if (form.UpFileName.value.substring(i).toUpperCase() != fileName.toUpperCase()) {
		alert('檔名必須為' + fileName)
		form.UpFileName.focus();
		return;
	}
	
	if (form.Report_no.value == 'B05') {
		if ((form.S_MONTH.value != '06') && (form.S_MONTH.value != '12')) {
			alert("半年報基準日必須為6月或12月!")
			return;
		}
	}	
	
	if (trimString(form.UpFileName.value) == "") {
		alert("上傳檔案位置必須輸入")
		form.UpFileName.focus();
		return;
	}
	
	form.act.value="Upload";	
	if (confirm('確定上傳該資料檔??')) {	   
		form.action="/pages/WMFileUpload.jsp?act="+form.act.value+"&FileName="+form.UpFileName.value+"&Report_no="+form.Report_no.value+"&S_YEAR="+form.S_YEAR.value+"&S_MONTH="+form.S_MONTH.value+"&bank_type="+form.bank_type.value+"&test=nothing";
		form.submit();
	}else{
		return;
	}	
}

function ConfirmOverWrite(form,msg){		
		if(msg != ''){			
	    	if (confirm(msg)) {	    		
	    		form.action="/pages/WMFileUpload.jsp?act=OverWrite&FileName="+form.FileName.value+"&test=nothing";	    		  
	    		form.submit();	    		
    		}else{
    			history.back();		
    		}		   		
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
