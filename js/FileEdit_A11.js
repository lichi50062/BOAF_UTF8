//globe var 
var isEmptyArr = new Array(); //�s�񦹦~�릳�S����
 
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
			alert("�Ӧ~��ӳ���Ƥw��w,�L�k�A���ʸӥӳ����");
			return;
		}else{
			//�Y�W�릳�ӳ���Ʈ�,���I��[����L�p�X�U�ڮץ���]��,�h�ˮ�[��ڱ«H�l�B]������J
			if(isLastMonth=="Y"){
				if(checkData(form,cnd) && check_No_Apply(form)) form.submit();
			}else{
				if(check_No_Apply(form)) form.submit();
			}
		}
	}
	if(((cnd == "Insert") || (cnd == "Update"))){
		if(A11_Lock=="Y"){
			alert("�Ӧ~��ӳ���Ƥw��w,�L�k�A���ʸӥӳ����");
			return;
		}else{
			if(checkData(form,cnd))	form.submit();

		}
	}
	if((cnd == "Delete")){
		if(A11_Lock=="Y"){
			alert("�Ӧ~��ӳ���Ƥw��w,�L�k�A���ʸӥӳ����");
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

//�A���T�{���S�����=========================================================
function check_No_Apply(form){
	if(!confirm("�T�w����L�p�X�U�ڮץ�ӳ����!!")){
			return false;
	}
	return true;
}
//�A���T�{�O�_�R�����=========================================================
function AskDelete(form) {
	return confirm("�T�w�n�R����");
}
//����ˮ�=========================================================	
function checkData(form,cnd){	
	if(cnd != 'No_Apply'){
		if (trimString(form.m_year.value) =="" ){
			alert("�ӳ��~�����i�ť�");
			form.m_year.focus();
			return false;
		}
		if (trimString(form.m_month.value) =="" ){
			alert("�ӳ�������i�ť�");
			form.m_month.focus();
			return false;
		}		
		if (trimString(form.loan_idn.value) =="" ){
			alert("�ɴڤH�Τ@�s�����i�ť�");
			form.loan_idn.focus();
			return false;
		}else{
			if(!check_identity_no(trimString(form.loan_idn.value))){
				alert("�ɴڤH�Τ@�s���榡�����T");
				form.loan_idn.focus();
				return false;
			}
		}
		if (trimString(form.loan_name.value) =="" ){
			alert("�ɴڤH�W�٤��i�ť�");
			form.loan_name.focus();
			return false;
		}
		if (trimString(form.loan_amt_sum.value) =="" ){
			alert("�«H���`���B���i�ť�");
			form.loan_amt_sum.focus();
			return false;
		}
		if (trimString(form.case_begin_year.value) =="" ){
			alert("�«H�״���-�_�l(�~)���i�ť�");
			form.case_begin_year.focus();
			return false;
		}
		if (trimString(form.case_begin_month.value) =="" ){
			alert("�«H�״���-�_�l(��)���i�ť�");
			form.case_begin_month.focus();
			return false;
		}
		if (trimString(form.case_end_year.value) =="" ){
			alert("�«H�״���-���(�~)���i�ť�");
			form.case_end_year.focus();
			return false;
		}
		if (trimString(form.case_end_month.value) =="" ){
			alert("�«H�״���-���(��)���i�ť�");
			form.case_end_month.focus();
			return false;
		}
		if (trimString(form.bank_no_max.value) =="" ){
			alert("�D��椣�i�ť�");
			form.bank_no_max.focus();
			return false;
		}
		if (trimString(form.manabank_name.value) =="" ){
			alert("�޲z�椣�i�ť�");
			form.manabank_name.focus();
			return false;
		}
		if (trimString(form.loan_kind.value) =="" ){
			alert("�ѶU�������i�ť�");
			form.loan_kind.focus();
			return false;
		}
		if (trimString(form.loan_amt.value) =="" ){
			alert("�ѶU�B�פ��i�ť�");
			form.loan_amt.focus();
			return false;
		}
		if (trimString(form.loan_bal_amt.value) =="" ){
			alert("��ڱ«H�l�B���i�ť�");
			form.loan_bal_amt.focus();
			return false;
		}
		if (trimString(form.loan_type.value) =="" ){
			alert("�H�γ��ѶU�������«H�γ~���i�ť�");
			form.loan_type.focus();
			return false;
		}
		if (trimString(form.pay_state.value) =="" ){
			alert("�ثe���ú�����Τ��i�ť�");
			form.pay_state.focus();
			return false;
		}
		if (trimString(form.violate_type.value) =="" ){
			alert("�L�H�ϫ����ӿձ��ڤ��i�ť�");
			form.violate_type.focus();
			return false;
		}
		if (trimString(form.loan_rate.value) =="" ){
			alert("�ثe��ڧQ�v���i�ť�");
			form.loan_rate.focus();
			return false;
		}
		if (trimString(form.new_case.value) =="" ){
			alert("�O�_������s�W�ץ󤣥i�ť�");
			form.new_case.focus();
			return false;
		}
	}else{
		if (trimString(form.loan_bal_amt.value) =="" ){
			alert("��ڱ«H�l�B���i�ť�");
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
    	alert("�п�J�Ʀr");
   	    return false;
    }else if (rate > 100.0000 || rate < -99.9999) {
        alert("�Q�v���i�j�� 100.0000, �]���i�p�� -99.9999")
        return false;
    }else if (rate.indexOf(".") != -1 ) {
        len = (rate.substring(rate.indexOf(".") + 1, rate.length));
        if (len.length > 4) {
            alert("�p���I��u�঳�|�Ӧ��");
            return false;
        }
    }
    
}

