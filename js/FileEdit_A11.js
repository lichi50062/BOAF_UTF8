//globe var 
var isEmptyArr = new Array(); //存放此年月有沒有值
 
//do request to WMFileEdit_A11=========================================================	
function printSubmit(form,bank_no,bank_name){
	form.action="/pages/WMFileEdit_A11_Print.jsp?bank_name="+bank_name+"&bank_no="+bank_no+"&s_year="+form.s_year.value+"&s_month="+form.s_month.value+"&Unit="+form.Unit.value;	
	form.submit();
}
function doSubmit(form,cnd,parameter,A11_Lock,isLastMonth){
	form.action="/pages/WMFileEdit_A11.jsp?act="+cnd+"&bank_no="+parameter+"&test=nothing"; 
	
	if(cnd == "New") {
		form.submit(); 
	}
	if(cnd == "No_Apply" ){
		if(A11_Lock=="Y"){
			alert("該年月申報資料已鎖定,無法再異動該申報資料");
			return;
		}else{
			//若上月有申報資料時,但點選[本月無聯合貸款案件資料]時,則檢核[實際授信餘額]必須輸入
			if(isLastMonth=="Y"){
				if(checkData(form,cnd) && check_No_Apply(form)) form.submit();
			}else{
				if(check_No_Apply(form)) form.submit();
			}
		}
	}
	if(((cnd == "Insert") || (cnd == "Update"))){
		if(A11_Lock=="Y"){
			alert("該年月申報資料已鎖定,無法再異動該申報資料");
			return;
		}else{
			if(checkData(form,cnd))	form.submit();

		}
	}
	if((cnd == "Delete")){
		if(A11_Lock=="Y"){
			alert("該年月申報資料已鎖定,無法再異動該申報資料");
			return;
		}else{
			if(AskDelete(form)){
				form.submit();
			}
		}
	}
}
//=========================================================	
//arrpush store year quarter data to array
function arrpush(inputData){
	isEmptyArr[isEmptyArr.length]=trimString(inputData);
}

//再次確認有沒有資料=========================================================
function check_No_Apply(form){
	if(!confirm("確定本月無聯合貸款案件申報資料!!")){
			return false;
	}
	return true;
}
//再次確認是否刪除資料=========================================================
function AskDelete(form) {
	return confirm("確定要刪除嗎");
}
//欄位檢核=========================================================	
function checkData(form,cnd){	
	if(cnd != 'No_Apply'){
		if (trimString(form.m_year.value) =="" ){
			alert("申報年份不可空白");
			form.m_year.focus();
			return false;
		}
		if (trimString(form.m_month.value) =="" ){
			alert("申報月份不可空白");
			form.m_month.focus();
			return false;
		}		
		if (trimString(form.loan_idn.value) =="" ){
			alert("借款人統一編號不可空白");
			form.loan_idn.focus();
			return false;
		}else{
			if(!check_identity_no(trimString(form.loan_idn.value))){
				alert("借款人統一編號格式不正確");
				form.loan_idn.focus();
				return false;
			}
		}
		if (trimString(form.loan_name.value) =="" ){
			alert("借款人名稱不可空白");
			form.loan_name.focus();
			return false;
		}
		if (trimString(form.loan_amt_sum.value) =="" ){
			alert("授信案總金額不可空白");
			form.loan_amt_sum.focus();
			return false;
		}
		if (trimString(form.case_begin_year.value) =="" ){
			alert("授信案期間-起始(年)不可空白");
			form.case_begin_year.focus();
			return false;
		}
		if (trimString(form.case_begin_month.value) =="" ){
			alert("授信案期間-起始(月)不可空白");
			form.case_begin_month.focus();
			return false;
		}
		if (trimString(form.case_end_year.value) =="" ){
			alert("授信案期間-到期(年)不可空白");
			form.case_end_year.focus();
			return false;
		}
		if (trimString(form.case_end_month.value) =="" ){
			alert("授信案期間-到期(月)不可空白");
			form.case_end_month.focus();
			return false;
		}
		if (trimString(form.bank_no_max.value) =="" ){
			alert("主辦行不可空白");
			form.bank_no_max.focus();
			return false;
		}
		if (trimString(form.manabank_name.value) =="" ){
			alert("管理行不可空白");
			form.manabank_name.focus();
			return false;
		}
		if (trimString(form.loan_kind.value) =="" ){
			alert("參貸型式不可空白");
			form.loan_kind.focus();
			return false;
		}
		if (trimString(form.loan_amt.value) =="" ){
			alert("參貸額度不可空白");
			form.loan_amt.focus();
			return false;
		}
		if (trimString(form.loan_bal_amt.value) =="" ){
			alert("實際授信餘額不可空白");
			form.loan_bal_amt.focus();
			return false;
		}
		if (trimString(form.loan_type.value) =="" ){
			alert("信用部參貸部分之授信用途不可空白");
			form.loan_type.focus();
			return false;
		}
		if (trimString(form.pay_state.value) =="" ){
			alert("目前放款繳息情形不可空白");
			form.pay_state.focus();
			return false;
		}
		if (trimString(form.violate_type.value) =="" ){
			alert("無違反契約承諾條款不可空白");
			form.violate_type.focus();
			return false;
		}
		if (trimString(form.loan_rate.value) =="" ){
			alert("目前放款利率不可空白");
			form.loan_rate.focus();
			return false;
		}
		if (trimString(form.new_case.value) =="" ){
			alert("是否為本月新增案件不可空白");
			form.new_case.focus();
			return false;
		}
	}else{
		if (trimString(form.loan_bal_amt.value) =="" ){
			alert("實際授信餘額不可空白");
			form.loan_bal_amt.focus();
			return false;
		}
	}
	if(form.loan_amt_sum.value!='' && !checkNumber(form.loan_amt_sum) ){
		return false;
	}
	if(form.case_begin_year.value!='' && !checkNumber(form.case_begin_year) ){
			return false;
	}
	if(form.case_end_year.value!='' && !checkNumber(form.case_end_year) ){
			return false;
	}
	if(form.loan_amt.value!='' && !checkNumber(form.loan_amt) ){
			return false;
	}
	if(form.loan_bal_amt.value!='' && !checkNumber(form.loan_bal_amt) ){
			return false;
	}
	
	if(form.loan_rate.value!=''){
		if(checkRate0000(form.loan_rate)==false){
			return false;
		}
	}
		
   return true;
}


//=========================================================
function checkRate0000(Rate1) {
	
	var rate = changeVal(Rate1);
    if (isNaN(Math.abs(rate))) {
    	alert("請輸入數字");
   	    return false;
    }else if (rate > 100.0000 || rate < -99.9999) {
        alert("利率不可大於 100.0000, 也不可小於 -99.9999")
        return false;
    }else if (rate.indexOf(".") != -1 ) {
        len = (rate.substring(rate.indexOf(".") + 1, rate.length));
        if (len.length > 4) {
            alert("小數點後只能有四個位數");
            return false;
        }
    }
    
}

