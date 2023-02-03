//103.01.02 fix 若現在月份為0[1月份],則申報年月為前一年12月份 by 2295
//104.05.12 add 增加A111111111可調整案件編號及分項項數 by 2295
//104.05.12 fix 補舊案件資料時,新增下一筆參貸項目時,該案件編號會是空值 by 2295
//globe var 
var isEmptyArr = new Array(); //存放此年月有沒有值
 
//do request to WMFileEdit_A11=========================================================	
function printSubmit(form,bank_no,bank_name){
	form.action="/pages/WMFileEdit_A11_Print.jsp?bank_name="+bank_name+"&bank_no="+bank_no+"&s_year="+form.s_year.value+"&s_month="+form.s_month.value+"&Unit="+form.Unit.value;	
	form.submit();
}
//104.05.12 add 修改案件編號/分項項數
function updateCase_No(form,m_year,m_month,bank_no,seq_no){	
	//alert(this.document.forms[0].upd_case_no.value);
	//alert(this.document.forms[0].upd_case_cnt.value);
	form.action="/pages/WMFileEdit_A11.jsp?act=updateCase_NO&m_year="+m_year+"&m_month="+m_month+"&bank_no="+bank_no+"&seq_no="+seq_no+"&upd_case_no="+form.upd_case_no.value+"&upd_case_cnt="+form.upd_case_cnt.value;	
	form.submit();
}
//104.05.12 add 新增下一筆參貸項目
function continueEditCnt(form,m_year,m_month,bank_no,case_no){	
	//alert(this.document.forms[0].upd_case_no.value);
	//alert(this.document.forms[0].upd_case_cnt.value);
	form.action="/pages/WMFileEdit_A11.jsp?act=continueEditCnt&m_year="+m_year+"&m_month="+m_month+"&bank_no="+bank_no+"&case_no="+case_no;	
	form.submit();
}
function doSubmit(form,cnd,parameter,A11_Lock,isLastMonth,isHaveApplyData,isHaveNoApplyData,isHaveNoApplyLoanData){
	
	form.action="/pages/WMFileEdit_A11.jsp?act="+cnd+"&bank_no="+parameter+"&test=nothing"; 
	if(cnd=="nextPg"){
		if(form.radio1[1].checked){
			if(A11_Lock=="Y"){
				alert("該年月申報資料已鎖定,無法再異動該申報資料");
				return false;
			}else if(isHaveNoApplyData =="Y"){
				alert("本月已申報[本月無聯合貸款案件申報資料],若需新增資料,請刪除原資料後,再新增");
				return false;
			}
			if(!confirm("確定本月無聯合貸款案件申報資料!!")){
				return false;
			}else{
				cnd = "No_Apply";
				form.action="/pages/WMFileEdit_A11.jsp?act="+cnd+"&bank_no="+parameter+"&test=nothing";
			}
		}else{
			if(A11_Lock=="Y"){
				alert("該年月申報資料已鎖定,無法再異動該申報資料");
				return false;
			}else if(isHaveNoApplyData =="Y"){
				alert("本月已申報[本月無聯合貸款案件申報資料],若需新增資料,請刪除原資料後,再新增");
				return false;
			}
			if(form.radio2[0].checked){
				if(document.getElementById('case_cnt').value==""){
					alert("請輸入項數!");
					return false;
				}else if(document.getElementById('case_cnt').value=="1"){
					alert("項數至少2筆以上!");
					return false;
				}
				if(!checkNumber_A11(document.getElementById('case_cnt'))){
			    	alert("項數 須為數字!");
			   	    return false;
				}
				
			}else if(form.radio2[1].checked){
				document.getElementById('case_cnt').value="1";
			}
			alert("開始新增"+document.getElementById('s_year').value+"年"+document.getElementById('s_month').value+"月聯貸款案件資料!!");
		}
		form.submit();
	}
	if(cnd == "New"){
		today = new Date();
		today_year = today.getYear()-1911;
		today_month = today.getMonth();
		if(today_month == 0){//若現在月份為0[1月份],則申報年月為前一年12月份 103.01.02 add
	       today_year = today_year - 1;
	       today_month = 12;
	    }	
		//alert('today_year='+today_year);
		//alert('today_month='+today_month);
		//alert('s_year'+document.getElementById('s_year').value);
		//alert('s_month'+document.getElementById('s_month').value);
		if(document.getElementById('s_year').value!= today_year || 
				document.getElementById('s_month').value != today_month ){
			alert("聯合貸款案件資料只開放新增上月份資料,尚未開放新增其他月份資料!!");
		}else{
			form.submit();
		}
	}
	/*
	if(cnd == "continueEditCnt") {
		form.submit();
	}
	*/
	if(((cnd == "Insert") || (cnd == "Update"))){
		if(A11_Lock=="Y"){
			alert("該年月申報資料已鎖定,無法再異動該申報資料");
			return;
		}else if((isHaveNoApplyData =="Y" ||isHaveNoApplyLoanData == "Y") && cnd == "Insert"){
			alert("本月已申報[本月無聯合貸款案件申報資料],若需新增資料,已刪除原資料後,再新增");
			return;
		}else{
			if(checkData(form,cnd)){
				form.submit();
			}
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
	
	if (trimString(form.loan_idn.value) =="" ){
		alert("借款人統一編號 不可空白");
		form.loan_idn.focus();
		return false;
	}
	if (trimString(form.loan_name.value) =="" ){
		alert("借款人名稱 不可空白");
		form.loan_name.focus();
		return false;
	}
	if (trimString(form.loan_amt_sum.value) =="" ){
		alert("授信案總金額 不可空白");
		form.loan_amt_sum.focus();
		return false;
	}else{
		if(!checkNumber_A11(form.loan_amt_sum)){
	    	alert("授信案總金額 需為數字");
	    	form.loan_amt_sum.focus();
	   	    return false;
		}
	}
	
	if (trimString(form.case_begin_year.value) =="" && trimString(form.case_begin_month.value) !=""){
		alert("授信案期間-起始(年) 不可空白");
		form.case_begin_year.focus();
		return false;
	}
	if (trimString(form.case_begin_year.value) !="" && trimString(form.case_begin_month.value) ==""){
		alert("授信案期間-起始(月) 不可空白");
		form.case_begin_month.focus();
		return false;
	}
	if (trimString(form.case_end_year.value) =="" && trimString(form.case_end_month.value) !=""){
		alert("授信案期間-到期(年) 不可空白");
		form.case_end_year.focus();
		return false;
	}
	if (trimString(form.case_end_year.value) !="" && trimString(form.case_end_month.value) ==""){
		alert("授信案期間-到期(月) 不可空白");
		form.case_end_month.focus();
		return false;
	}
	if (trimString(form.case_begin_year.value) =="" && trimString(form.case_end_year.value) !=""){
		alert("授信案期間-起始年月 不可空白");
		form.case_begin_month.focus();
		return false;
	} 
	if (trimString(form.case_begin_year.value) !="" && trimString(form.case_end_year.value) ==""){
		alert("授信案期間-到期年月 不可空白");
		form.case_end_month.focus();
		return false;
	}
	if(form.case_begin_year.value!='' && !checkNumber_A11(form.case_begin_year)){
    	alert("授信案期間-起始年月 需為數字");
    	form.case_begin_year.focus();
   	    return false;
	}
	if(form.case_end_year.value!='' && !checkNumber_A11(form.case_end_year)){
	    	alert("授信案期間-到期年月 需為數字");
	    	form.case_end_year.focus();
	   	    return false;
	}
	if( trimString(form.case_begin_year.value) !="" && trimString(form.case_begin_month.value) !="" 
		&& trimString(form.case_end_year.value) !="" && trimString(form.case_end_month.value) !="" ){
		if(parseInt(trimString(form.case_begin_year.value)+padLeft(trimString(form.case_begin_month.value),2))>
			parseInt(trimString(form.case_end_year.value)+padLeft(trimString(form.case_end_month.value),2))){
				alert("授信案期間 起始年月 不可大於 到期年月");
				return false;
		}
	}
	if (trimString(form.bank_no_max.value) ==""){
		alert("主辦行 不可空白");
		form.bank_no_max.focus();
		return false;
	}
	if (trimString(form.manabank_name.value) ==""){
		alert("管理行 不可空白");
		form.manabank_name.focus();
		return false;
	}
	if (trimString(form.loan_kind.value) ==""){
		alert("參貸型式 不可空白");
		form.loan_kind.focus();
		return false;
	}
	if (trimString(form.loan_amt.value) ==""){
		alert("參貸額度 不可空白");
		form.loan_amt.focus();
		return false;
	}else{
		if(!checkNumber_A11(form.loan_amt)){
	    	alert("參貸額度 需為數字");
	    	form.loan_amt.focus();
	   	    return false;
		}
	}
	if(form.loan_bal_amt.value!='' && !checkNumber_A11(form.loan_bal_amt) ){
    	alert("實際授信餘額 需為數字");
    	form.loan_bal_amt.focus();
   	    return false;
	}
	if (trimString(form.loan_type.value) ==""){
		alert("信用部參貸部分之授信用途 不可空白");
		form.loan_type.focus();
		return false;
	}
	if(form.loan_rate.value!=''){
		if(checkRate0000(form.loan_rate)==false){
			return false;
		}
	}
	if (trimString(form.new_case.value) ==""){
		alert("是否為本月新增案件不可空白");
		form.new_case.focus();
		return false;
	}
	if(form.loan_amt_sum.value!='' && form.loan_amt.value!=''){
		if(parseInt(form.loan_amt.value) > parseInt(form.loan_amt_sum.value)){
			alert("授信案總金額 須大於或等於 參貸額度");
			form.loan_amt_sum.focus();
			return false;
		}
	}
	if(form.loan_bal_amt.value!='' && form.loan_amt.value!=''){
		if(parseInt(form.loan_bal_amt.value) > parseInt(form.loan_amt.value)){
			alert("實際授信餘額 須小於或等於 參貸額度");
			form.loan_bal_amt.focus();
			return false;
		}
	}
	var case_no = form.case_no.value;
	if((case_no.substring(7,9))!="01"){
		if(form.loan_amt_sum.value!='' && form.firstLoan_amt_sum.value !=null && form.firstLoan_amt_sum.value !="" && form.firstLoan_amt_sum.value!=form.loan_amt_sum.value){
			alert("授信案總金額 與前一分項授信案總金額 不一致");
			form.loan_amt_sum.focus();
			return false;
		}
	}
   return true;
}
function checkNumber_A11(Rate1){

	Rate1.value = changeVal(Rate1);
	if (isNaN(Math.abs(Rate1.value))) {
   	    return false;
    }
    if ((Rate1.value).indexOf(".") != -1 ) {
    	alert("不可輸入小數點");
        return false;
    }
    return true;
}

//=========================================================
function checkRate0000(Rate1) {
	
	var rate = changeVal(Rate1);
    if (isNaN(Math.abs(rate))) {
    	alert("目前放款利率 需為數字");
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

// 左邊補0======================================================
function padLeft(str,lenght){  
	if(str.length >= lenght)  
		return str;  
	else 
		return padLeft("0" +str,lenght);  

}  


