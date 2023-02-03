function Check_ZZ017W(frm) {

	if (frm.bank_type.value == 'F') {
		if (frm.report_no.value == ''){
			alert("請輸入報表編號");
       		return false;
       	}
	}
	if (trimString(frm.cano.value) == '') {
		alert("請輸入公式編號");
       	return false;
    }
    if (trimString(frm.remark.value) == '') {
		alert("請輸入公式說明");
       	return false;
    }
	if (frm.L_acc_code.options.length == 0) {
		alert("請選擇左式科目代碼");
       	return false;
	}
	if (frm.R_acc_code.options.length == 0) {
		alert("請選擇右式科目代碼");
       	return false;
	}
	var i = frm.L_operator.options.length;
	var j = frm.R_operator.options.length;

	if (frm.L_operator.options[i - 1].value != '') {
		alert("左式最後一個運算式不可有值");
       	return false;
	}
	if (frm.R_operator.options[j - 1].value != '') {
		alert("右式最後一個運算式不可有值");
       	return false;
	}
	//selected set to true;
	for (i = 0; i < frm.L_acc_code.options.length; i++){
		frm.L_acc_code.options[i].selected = true;
		frm.L_operator.options[i].selected = true;
	}
	for (i = 0; i < frm.R_acc_code.options.length; i++){
		frm.R_acc_code.options[i].selected = true;
		frm.R_operator.options[i].selected = true;
	}
	return true;
}
//===========================
function doSubmitZZ017W(form, myfun) {
    form.Function.value = myfun;
    if (form.Function.value == 'delete') {
        if (AskDelete())
            form.submit();
        return;
    }

    if (form.Function.value == 'update') {
    	if (!Check_ZZ017W(form))
    		return false;
	}
	var j;
	for (j =0; j < form.L_acc_code.options.length - 1; j++){
		if (form.L_operator.options[j].value == ''){
			alert('左式第' + (j + 1) + '個項目名稱的運算式必須有值');
			return false;
		}
	}
	for (j =0; j < form.R_acc_code.options.length - 1 ; j++){
		if (form.R_operator.options[j].value == ''){
			alert('右式第' + (j + 1) + '個項目名稱的運算式必須有值');
			return false;
		}
	}

	for (i = 0; i < form.L_acc_code.options.length; i++){
		form.L_acc_code.options[i].selected = true;
		form.L_operator.options[i].selected = true;
	}
	for (i = 0; i < form.R_acc_code.options.length; i++){
		form.R_acc_code.options[i].selected = true;
		form.R_operator.options[i].selected = true;
	}

	form.submit();
}
//===========================
function addAccCode() {

	var frm = frmZZ017W;
	var i, l, LRname;
	var LR_acc_code, LR_operator;

	if (frm.LR[0].checked) {
		LR_acc_code = frm.L_acc_code;
		LR_operator = frm.L_operator;
		LRname = '左式';
	}
	else {
		LR_acc_code = frm.R_acc_code;
		LR_operator = frm.R_operator;
		LRname = '右式';
	}

	i = frm.acc_code.selectedIndex;
	l = frm.operator.selectedIndex;

	for (var j =0; j < LR_acc_code.options.length; j++){
		if (LR_acc_code.options[j].value == frm.acc_code.options[frm.acc_code.options.selectedIndex].value)
			return;
		if (LR_operator.options[j].value == ''){
			alert(LRname + '第' + (j + 1) + '個項目名稱的運算式必須有值');
			return;
		}
	}

	LR_acc_code.options[LR_acc_code.options.length] = new Option(frm.acc_code.options[i].text, frm.acc_code.options[i].value);
	LR_operator.options[LR_operator.options.length] = new Option(frm.operator.options[l].text, frm.operator.options[l].value);

} //addAccCode()

//===========================
function addFixCode() {

	var frm = frmZZ017W;
	var i, l, LRname;
	var LR_acc_code, LR_operator;

	if (frm.LR[0].checked) {
		LR_acc_code = frm.L_acc_code;
		LR_operator = frm.L_operator;
		LRname = '左式';
	}
	else {
		LR_acc_code = frm.R_acc_code;
		LR_operator = frm.R_operator;
		LRname = '右式';
	}

	if (frm.CONST.value == '') {
		alert('請選擇常數');
		return false;
	}
	i = frm.CONST.selectedIndex;
	l = frm.operator.selectedIndex;

	for (var j =0; j < LR_acc_code.options.length; j++){
		if (LR_acc_code.options[j].value == frm.CONST.options[frm.CONST.options.selectedIndex].value)
			return;
		if (LR_operator.options[j].value == ''){
			alert(LRname + '第' + (j + 1) + '個項目名稱的運算式必須有值');
			return;
		}
	}

	LR_acc_code.options[LR_acc_code.options.length] = new Option('常數 ' + frm.CONST.options[i].text, frm.CONST.options[i].value);
	LR_operator.options[LR_operator.options.length] = new Option(frm.operator.options[l].text, frm.operator.options[l].value);

} //addFixCode()

function deleteOne(sel) {

	var LR_operator;
	if (sel.name.substring(0, 1) == 'R')
		LR_operator = frmZZ017W.R_operator;
	else
		LR_operator = frmZZ017W.L_operator;

	for (var i = sel.options.length -1; i >= 0 ; i--){
		if ((sel.options[i] != null) && (sel.options[i].selected) == true) {
			sel.options[i] = null;
			LR_operator.options[i] = null;
		}
	}

}

function chk(sel){

	var frm = frmZZ017W;
	var LR_operator;

	if (sel.name.substring(0, 1) == 'R')
		LR_operator = frmZZ017W.R_operator;
	else
		LR_operator = frmZZ017W.L_operator;

	for (var i =0; i < LR_operator.length; i++)
		LR_operator.options[i].selected = false;

	LR_operator.options[sel.selectedIndex].selected = true;

}