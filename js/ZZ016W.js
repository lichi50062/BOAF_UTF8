function Check_ZZ016W(form) {
    if (trimString(form.acc_code.value) ==''){
       alert("�п�J��إN�X");
       form.acc_code.focus();
       return false;
    }
    if (trimString(form.acc_name.value) ==''){
       alert("�п�J��ئW��");
       form.acc_name.focus();
       return false;
    }
    if (trimString(form.acc_b_name.value) ==''){
       alert("�п�J���²��");
       form.acc_b_name.focus();
       return false;
    }
    return true;
}

function doSubmitZZ016W(form, myfun) {
    form.Function.value = myfun;
    if (form.Function.value == 'delete') {
        if (AskDelete())
            form.submit();
    }
    else if (form.Function.value == 'update') {
        if (Check_ZZ016W(form))
            form.submit();
    }
}

//===========================
function checkFinance() {

	var frm = document.frmZZ016W;
	var i = frm.bank_type.selectedIndex;

	if (frm.bank_type.options[i].value == 'F')
		frm.RptNo.style.visibility = 'visible';
	else
		frm.RptNo.style.visibility = 'hidden';
}