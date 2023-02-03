function Check_ZZ017W(frm) {

	if (frm.bank_type.value == 'F') {
		if (frm.report_no.value == ''){
			alert("�п�J����s��");
       		return false;
       	}
	}
	if (trimString(frm.cano.value) == '') {
		alert("�п�J�����s��");
       	return false;
    }
    if (trimString(frm.remark.value) == '') {
		alert("�п�J��������");
       	return false;
    }
	if (frm.L_acc_code.options.length == 0) {
		alert("�п�ܥ�����إN�X");
       	return false;
	}
	if (frm.R_acc_code.options.length == 0) {
		alert("�п�ܥk����إN�X");
       	return false;
	}
	var i = frm.L_operator.options.length;
	var j = frm.R_operator.options.length;

	if (frm.L_operator.options[i - 1].value != '') {
		alert("�����̫�@�ӹB�⦡���i����");
       	return false;
	}
	if (frm.R_operator.options[j - 1].value != '') {
		alert("�k���̫�@�ӹB�⦡���i����");
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
			alert('������' + (j + 1) + '�Ӷ��ئW�٪��B�⦡��������');
			return false;
		}
	}
	for (j =0; j < form.R_acc_code.options.length - 1 ; j++){
		if (form.R_operator.options[j].value == ''){
			alert('�k����' + (j + 1) + '�Ӷ��ئW�٪��B�⦡��������');
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
		LRname = '����';
	}
	else {
		LR_acc_code = frm.R_acc_code;
		LR_operator = frm.R_operator;
		LRname = '�k��';
	}

	i = frm.acc_code.selectedIndex;
	l = frm.operator.selectedIndex;

	for (var j =0; j < LR_acc_code.options.length; j++){
		if (LR_acc_code.options[j].value == frm.acc_code.options[frm.acc_code.options.selectedIndex].value)
			return;
		if (LR_operator.options[j].value == ''){
			alert(LRname + '��' + (j + 1) + '�Ӷ��ئW�٪��B�⦡��������');
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
		LRname = '����';
	}
	else {
		LR_acc_code = frm.R_acc_code;
		LR_operator = frm.R_operator;
		LRname = '�k��';
	}

	if (frm.CONST.value == '') {
		alert('�п�ܱ`��');
		return false;
	}
	i = frm.CONST.selectedIndex;
	l = frm.operator.selectedIndex;

	for (var j =0; j < LR_acc_code.options.length; j++){
		if (LR_acc_code.options[j].value == frm.CONST.options[frm.CONST.options.selectedIndex].value)
			return;
		if (LR_operator.options[j].value == ''){
			alert(LRname + '��' + (j + 1) + '�Ӷ��ئW�٪��B�⦡��������');
			return;
		}
	}

	LR_acc_code.options[LR_acc_code.options.length] = new Option('�`�� ' + frm.CONST.options[i].text, frm.CONST.options[i].value);
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