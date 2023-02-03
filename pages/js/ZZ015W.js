function doSubmitZZ015W(form) {

	if (form.YM[0].checked) {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	if (form.DLId.value == "")	return false;

	return true;
}

function checkYM() {
	var frm = document.frmZZ015W_LIST02;

	if (frm.YM[0].checked) {
		frm.S_YEAR.disabled = false;
		frm.S_MONTH.disabled = false;
	}
	else {
		frm.S_YEAR.disabled = true;
		frm.S_MONTH.disabled = true;
	}
}