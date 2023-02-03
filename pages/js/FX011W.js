//96.03.22 add 總機構代碼使用XML
function changeOption(form){
    var myXML,bankType,bankNo, bankName;
    myXML = document.all("BankNoXML").XMLDocument;
    form.examine.length = 0;    
	bankType = myXML.getElementsByTagName("bankType");//總機構
	bankNo = myXML.getElementsByTagName("bankValue");//分支機構
	bankName = myXML.getElementsByTagName("bankName");//分支機構名稱
	var oOption;	
	
	//alert(form.tbank_no.value);
	for(var i=0;i<bankType.length ;i++){		
		//alert(bankType.item(i).firstChild.nodeValue);
  		if ((bankType.item(i).firstChild.nodeValue == form.tbank_no_examine.value)) {
  			oOption = document.createElement("OPTION");
  			//alert(bankName.item(i).firstChild.nodeValue);
  			//alert(bankNo.item(i).firstChild.nodeValue);
			oOption.text=bankName.item(i).firstChild.nodeValue;
  			oOption.value=bankNo.item(i).firstChild.nodeValue;
  			form.examine.add(oOption);
    	}
    }
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




function doSubmit(form,cnd,bank_no){	     	  
if(cnd =='add'){
	 if(!checkData(form)) return;	
	if(AskInsert_Permission()){
	form.action="/pages/FX011W.jsp?act=Insert";	    
	form.submit();
  }
}
else if(cnd =='modify'){
	 if(!checkData(form)) return;	
	if(AskEdit_Permission()){
	form.action="/pages/FX011W.jsp?act=Update";	    
	form.submit();
	}
}
else if(cnd =='delete'){
	if(AskDelete_Permission()){
	form.action="/pages/FX011W.jsp?act=Delete";	    
	form.submit();
	}
}
else if(cnd =='Print'){	
	if(form.m_year.value == ''){
	   alert("年度必須輸入!!");
	}else{   
		form.action="/pages/FX011W.jsp?act=Print&m2_name="+bank_no;	    
		form.submit();
	}
}

else return;
	  	  	    
}

function checkData(form) 
{
	var ckDate;
	if(trimString(form.CHECK_DATE_Y.value)  != "" ){        
    	if(isNaN(Math.abs(form.CHECK_DATE_Y.value))){
           alert("查核日期(年)不可輸入文字");    
		   form.BEG_DATE_Y.focus();            
           return false;
        }
    }else{
		alert("查核日期(年)不可空白");
		form.CHECK_DATE_Y.focus();
		return false;   
	}   
    if (trimString(form.CHECK_DATE_M.value) == "" ){
		alert("查核日期(月)不可空白");
		form.CHECK_DATE_M.focus();
		return false;
	}			
	if (trimString(form.CHECK_DATE_D.value) == "" ){
		alert("查核日期(日)不可空白");
		form.CHECK_DATE_D.focus();		
		return false;
	}	
    
    ckDate = '' + (parseInt(form.CHECK_DATE_Y.value)+1911) + '/' + form.CHECK_DATE_M.value + '/' + form.CHECK_DATE_D.value;
    form.CHECK_DATE.value = ckDate;   
    
	/*
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
    }
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
    */
    return true;
}

