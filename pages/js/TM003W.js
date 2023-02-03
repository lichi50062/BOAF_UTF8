

var act;

function AskInsert_Permission() {
  if(confirm("確定要存檔此筆資料嗎？"))
    return true;
  else
    return false;
}

/*
function AskReset(form){
	if(confirm("確定要存檔此筆資料嗎？")){
		for(var i=0; i<document.getElementsByName("loan_Cnt01").length; i++){
			document.getElementsByName("loan_Cnt01")[i].value="0";
		}
		for(var i=0; i<document.getElementsByName("loan_Cnt02").length; i++){
			document.getElementsByName("loan_Cnt02")[i].value="0";
		}
		for(var i=0; i<document.getElementsByName("loan_Amt01").length; i++){
			document.getElementsByName("loan_Amt01")[i].value="0";
		}
		for(var i=0; i<document.getElementsByName("loan_Amt02").length; i++){
			document.getElementsByName("loan_Amt02")[i].value="0";
		}
	}
}*/

function doSubmit(form,cnd){
	if(cnd =='Update'){
		if(checkData(form)){
			if(AskInsert_Permission()){
				form.action="/pages/TM003W.jsp?act=Update";	    
				form.submit();
			}
		}
	}else if(cnd =='Delete'){
		if( confirm("確定要刪除此筆資料嗎？")){
			form.action="/pages/TM003W.jsp?act=Delete";	    
			form.submit();
		}
	}else if(cnd =='List'){
	if( confirm("確定回清單頁？")){
		form.action="/pages/TM003W.jsp?act=Qry";	    
		form.submit();
		}
	}
	
	else return;
	  	  	    
}


function checkData(form) {
	form.loan_Sbm.value = "";
	var tmpStr = "";
	for(var i=0; i<document.getElementsByName("acc_Code01").length; i++){
		if(document.getElementsByName("acc_Code01")[i].value=="")document.getElementsByName("acc_Code01")[i].value="0";
		if(document.getElementsByName("loan_Cnt01")[i].value=="")document.getElementsByName("loan_Cnt01")[i].value="0";
		if(document.getElementsByName("loan_Amt01")[i].value=="")document.getElementsByName("loan_Amt01")[i].value="0";
		if(tmpStr!="") tmpStr+=";";
		tmpStr = tmpStr 
		            +"01" 
		            +","+document.getElementsByName("acc_Code01")[i].value
		            +","+document.getElementsByName("loan_Cnt01")[i].value
		            +","+document.getElementsByName("loan_Amt01")[i].value;
	}
	for(var i=0; i<document.getElementsByName("acc_Code02").length; i++){
		if(document.getElementsByName("acc_Code02")[i].value=="")document.getElementsByName("acc_Code02")[i].value="0";
		if(document.getElementsByName("loan_Cnt02")[i].value=="")document.getElementsByName("loan_Cnt02")[i].value="0";
		if(document.getElementsByName("loan_Amt02")[i].value=="")document.getElementsByName("loan_Amt02")[i].value="0";
		if(tmpStr!="") tmpStr+=";";
		tmpStr = tmpStr 
		            +"02" 
		            +","+document.getElementsByName("acc_Code02")[i].value
		            +","+document.getElementsByName("loan_Cnt02")[i].value
		            +","+document.getElementsByName("loan_Amt02")[i].value;
	}
	form.loan_Sbm.value = tmpStr;
	return true;
}

function submitform(form)
{	
    	form.bank_cnt.value = changeVal( form.bank_cnt)                    
		form.bank_bal.value = changeVal(form.bank_bal)    
		form.bank_cnt_s.value = changeVal(form.bank_cnt_s)   
		form.bank_bal_s.value = changeVal(form.bank_bal_s)                     
		form.bank_cnt_n.value = changeVal(form.bank_cnt_n)                      
		form.bank_bal_n.value = changeVal(form.bank_bal_n)                    			
		
	return true;                   
                                       
}