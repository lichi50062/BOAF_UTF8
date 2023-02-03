

var act;

function AskInsert_Permission() {
  if(confirm("確定要存檔此筆資料嗎？"))
    return true;
  else
    return false;
}


function doSubmit(form,cnd){	     	  
	if(cnd =='Update'){
		if(checkData(form)){
			if(AskInsert_Permission()){
				form.action="/pages/TM006W.jsp?act=Update";	    
				form.submit();
			}
		}
	}else if(cnd =='Delete'){
		if( confirm("確定要刪除此筆資料嗎？")){
			form.action="/pages/TM006W.jsp?act=Delete";	    
			form.submit();
		}
	}else if(cnd =='List'){
		if( confirm("確定回清單頁？")){
		form.action="/pages/TM006W.jsp?act=ApplyList";	    
		form.submit();
		}
	}
	
	else return;
	  	  	    
}


function checkData(form) {
	form.loan_Sbm.value = "";
	var tmpStr = "";
	var reason = "";
	for(var i=0; i<document.getElementsByName("acc_Code01").length; i++){
		if(document.getElementsByName("acc_Code01")[i].value=="")document.getElementsByName("acc_Code01")[i].value="0";
		if(document.getElementsByName("apply_Cnt01")[i].value=="")document.getElementsByName("apply_Cnt01")[i].value="0";
		if(document.getElementsByName("apply_Amt01")[i].value=="")document.getElementsByName("apply_Amt01")[i].value="0";
		if(document.getElementsByName("apply_Bal01")[i].value=="")document.getElementsByName("apply_Bal01")[i].value="0";
		if(document.getElementsByName("apply_Cnt_Sum01_P")[i].value=="")document.getElementsByName("apply_Cnt_Sum01_P")[i].value="0";
		if(document.getElementsByName("apply_Amt_Sum01_P")[i].value=="")document.getElementsByName("apply_Amt_Sum01_P")[i].value="0";
		if(document.getElementsByName("apply_Bal_Sum01_P")[i].value=="")document.getElementsByName("apply_Bal_Sum01_P")[i].value="0";
		if(document.getElementsByName("appr_Cnt01")[i].value=="")document.getElementsByName("appr_Cnt01")[i].value="0";
		if(document.getElementsByName("appr_Amt01")[i].value=="")document.getElementsByName("appr_Amt01")[i].value="0";
		if(document.getElementsByName("appr_Bal01")[i].value=="")document.getElementsByName("appr_Bal01")[i].value="0";
		if(document.getElementsByName("appr_Cnt_Sum01_P")[i].value=="")document.getElementsByName("appr_Cnt_Sum01_P")[i].value="0";
		if(document.getElementsByName("appr_Amt_Sum01_P")[i].value=="")document.getElementsByName("appr_Amt_Sum01_P")[i].value="0";
		if(document.getElementsByName("appr_Bal_Sum01_P")[i].value=="")document.getElementsByName("appr_Bal_Sum01_P")[i].value="0";
		if(document.getElementsByName("nonappr_Cnt01")[i].value=="")document.getElementsByName("nonappr_Cnt01")[i].value="0";
		reason = "null";
		if(document.getElementsByName("nonappr_Reason01")[i].value!="")reason=document.getElementsByName("nonappr_Reason01")[i].value;
		if(tmpStr!="") tmpStr+=";";
		tmpStr = tmpStr 
		            +"01" 
		            +","+document.getElementsByName("acc_Code01")[i].value
		            +","+document.getElementsByName("apply_Cnt01")[i].value
		            +","+document.getElementsByName("apply_Amt01")[i].value
		            +","+document.getElementsByName("apply_Bal01")[i].value
		            +","+(parseInt(document.getElementsByName("apply_Cnt_Sum01_P")[i].value)+parseInt(document.getElementsByName("apply_Cnt01")[i].value))
		            +","+(parseInt(document.getElementsByName("apply_Amt_Sum01_P")[i].value)+parseInt(document.getElementsByName("apply_Amt01")[i].value))
		            +","+(parseInt(document.getElementsByName("apply_Bal_Sum01_P")[i].value)+parseInt(document.getElementsByName("apply_Bal01")[i].value))
		            +","+document.getElementsByName("appr_Cnt01")[i].value
		            +","+document.getElementsByName("appr_Amt01")[i].value
		            +","+document.getElementsByName("appr_Bal01")[i].value
		            +","+(parseInt(document.getElementsByName("appr_Cnt_Sum01_P")[i].value)+parseInt(document.getElementsByName("appr_Cnt01")[i].value))
		            +","+(parseInt(document.getElementsByName("appr_Amt_Sum01_P")[i].value)+parseInt(document.getElementsByName("appr_Amt01")[i].value))
		            +","+(parseInt(document.getElementsByName("appr_Bal_Sum01_P")[i].value)+parseInt(document.getElementsByName("appr_Bal01")[i].value))
		            +","+document.getElementsByName("nonappr_Cnt01")[i].value
		            +","+reason
		            ;
	}
	for(var i=0; i<document.getElementsByName("acc_Code02").length; i++){
		if(document.getElementsByName("acc_Code02")[i].value=="")document.getElementsByName("acc_Code02")[i].value="0";
		if(document.getElementsByName("apply_Cnt02")[i].value=="")document.getElementsByName("apply_Cnt02")[i].value="0";
		if(document.getElementsByName("apply_Amt02")[i].value=="")document.getElementsByName("apply_Amt02")[i].value="0";
		if(document.getElementsByName("apply_Bal02")[i].value=="")document.getElementsByName("apply_Bal02")[i].value="0";
		if(document.getElementsByName("apply_Cnt_Sum02_P")[i].value=="")document.getElementsByName("apply_Cnt_Sum02_P")[i].value="0";
		if(document.getElementsByName("apply_Amt_Sum02_P")[i].value=="")document.getElementsByName("apply_Amt_Sum02_P")[i].value="0";
		if(document.getElementsByName("apply_Bal_Sum02_P")[i].value=="")document.getElementsByName("apply_Bal_Sum02_P")[i].value="0";
		if(document.getElementsByName("appr_Cnt02")[i].value=="")document.getElementsByName("appr_Cnt02")[i].value="0";
		if(document.getElementsByName("appr_Amt02")[i].value=="")document.getElementsByName("appr_Amt02")[i].value="0";
		if(document.getElementsByName("appr_Bal02")[i].value=="")document.getElementsByName("appr_Bal02")[i].value="0";
		if(document.getElementsByName("appr_Cnt_Sum02_P")[i].value=="")document.getElementsByName("appr_Cnt_Sum02_P")[i].value="0";
		if(document.getElementsByName("appr_Amt_Sum02_P")[i].value=="")document.getElementsByName("appr_Amt_Sum02_P")[i].value="0";
		if(document.getElementsByName("appr_Bal_Sum02_P")[i].value=="")document.getElementsByName("appr_Bal_Sum02_P")[i].value="0";
		if(document.getElementsByName("nonappr_Cnt02")[i].value=="")document.getElementsByName("nonappr_Cnt02")[i].value="0";
		reason = "null";
		if(document.getElementsByName("nonappr_Reason02")[i].value!="")reason=document.getElementsByName("nonappr_Reason02")[i].value;
		if(tmpStr!="") tmpStr+=";";
		tmpStr = tmpStr 
		            +"02" 
		            +","+document.getElementsByName("acc_Code02")[i].value
		            +","+document.getElementsByName("apply_Cnt02")[i].value
		            +","+document.getElementsByName("apply_Amt02")[i].value
		            +","+document.getElementsByName("apply_Bal02")[i].value
		            +","+(parseInt(document.getElementsByName("apply_Cnt_Sum02_P")[i].value)+parseInt(document.getElementsByName("apply_Cnt02")[i].value))
		            +","+(parseInt(document.getElementsByName("apply_Amt_Sum02_P")[i].value)+parseInt(document.getElementsByName("apply_Amt02")[i].value))
		            +","+(parseInt(document.getElementsByName("apply_Bal_Sum02_P")[i].value)+parseInt(document.getElementsByName("apply_Bal02")[i].value))
		            +","+document.getElementsByName("appr_Cnt02")[i].value
		            +","+document.getElementsByName("appr_Amt02")[i].value
		            +","+document.getElementsByName("appr_Bal02")[i].value
		            +","+(parseInt(document.getElementsByName("appr_Cnt_Sum02_P")[i].value)+parseInt(document.getElementsByName("appr_Cnt02")[i].value))
		            +","+(parseInt(document.getElementsByName("appr_Amt_Sum02_P")[i].value)+parseInt(document.getElementsByName("appr_Amt02")[i].value))
		            +","+(parseInt(document.getElementsByName("appr_Bal_Sum02_P")[i].value)+parseInt(document.getElementsByName("appr_Bal02")[i].value))
		            +","+document.getElementsByName("nonappr_Cnt02")[i].value
		            +","+reason
		            ;
	}
	form.loan_Sbm.value = tmpStr;
	return true;
}

