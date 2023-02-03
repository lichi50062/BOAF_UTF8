function Check_ZZ018W(frm) {
	if (frm.acc_code.value == "") {
		alert("請選擇科目代碼");
       	frm.acc_code.focus();
       	return false;
	}
	return true;
}
//===========================
function doSubmitZZ018W(form, myfun) {
    form.Function.value = myfun;
    if (form.Function.value == 'delete') {
        if (AskDelete())
            form.submit();
    }
}
//===========================
function checkFinance() {

	var frm = document.frmZZ018W;
	var i = frm.bank_type.selectedIndex;

	if (frm.bank_type.options[i].value == 'F')
		frm.RptNo.style.visibility = 'visible';
	else
		frm.RptNo.style.visibility = 'hidden';
}
//===========================
function loadRptNo(frm, frm_1) {

	var arr, i;
	var sltname, elemt;

	i = frm.RptNo.selectedIndex;
	sltname = frm.RptNo.options[i].value;

	if (frm.bank_type.options[frm.bank_type.selectedIndex].value != 'F') {
		if (frm_1.pre_type.value != 'F') return;
		frm_1.pre_type.value = '';
		frm.acc_code.options.length = 0;
		elemt = frm_1.acc_code;
		for (var k = 0; k < elemt.length; k++) {
			frm.acc_code.options[k] = new Option(elemt.options[k].text, elemt.options[k].value);
		}
	}
	else {
		frm.acc_code.options.length = 0;
		frm_1.pre_type.value = 'F';
		for (var j =0; j < frm_1.length; j++){
			if ((frm_1.elements[j].type == "select-one") && (frm_1.elements[j].name == sltname)){
				elemt = frm_1.elements[j]
				for (var k = 0; k < elemt.length; k++){
					frm.acc_code.options[k] = new Option(elemt.options[k].text, elemt.options[k].value);
				}
				break;
			}
		}
	}
} //End of loadRptNo()
