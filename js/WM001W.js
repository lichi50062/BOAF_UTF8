function doSubmitWM001W(form) {

	if (!Check_Maintain(form))
		return false;
	if (trimString(form.maintain_email.value) == "") {
		alert("�ӿ��E_MAIL������J")
		form.maintain_email.focus();
		return false;
	}
	if (!checkSingleYM(form.S_YEAR, form.S_MONTH)) {
		form.S_YEAR.focus();
		return false;
	}
	var FILE_NAME_LENGTH = 15; 	//A01002000009105
	var YM = ('000' + form.S_YEAR.value).substring(form.S_YEAR.value.length) + form.S_MONTH.value;
	var bank_code, fileName;

	if (form.BANK_CODE.value.length != 7)
		bank_code = form.BANK_CODE.value.substring(0, 3) + '0000';
	else
		bank_code = form.BANK_CODE.value;

	fileName = form.DLId.value + bank_code + YM;
	var i = form.UpFileName.value.length - FILE_NAME_LENGTH;

	if (form.UpFileName.value.substring(i) != fileName) {
		alert('�ɦW������' + fileName)
		form.UpFileName.focus();
		return false;
	}
	if (form.DLId.value == 'A05' || form.DLId.value == 'B05') {
		if ((form.S_MONTH.value != '06') && (form.S_MONTH.value != '12')) {
			alert("�b�~����Ǥ饲����6���12��!")
			return false;
		}
	}
	if (form.DLId.value == 'F12' || form.DLId.value == 'F13' || form.DLId.value == 'F14' || form.DLId.value == 'F15' ||
		form.DLId.value == 'F16' || form.DLId.value == 'F17') {
		if ((form.S_MONTH.value != '03') && (form.S_MONTH.value != '06') && (form.S_MONTH.value != '09') && (form.S_MONTH.value != '12')) {
			alert("�u����Ǥ饲����3,6,9,12��!")
			return false;
		}
	}
	if (trimString(form.INPUT_NAME.value) == "") {
		alert("�ӳ��̩m�W������J")
		form.INPUT_NAME.focus();
		return false;
	}
	if (trimString(form.INPUT_TEL.value) == "") {
		alert("�ӳ��̹q�ܥ�����J")
		form.INPUT_TEL.focus();
		return false;
	}

	return true;
}
