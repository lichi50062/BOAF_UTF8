function doSubmit(form) {

	if (!Check_Maintain(form))
		return false;
	if (trimString(form.maintain_email.value) == "") {
		alert("承辦員E_MAIL必須輸入")
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
    /*
	if (form.BANK_CODE.value.length != 7)
		bank_code = form.BANK_CODE.value.substring(0, 3) + '0000';
	else
		bank_code = form.BANK_CODE.value;

	fileName = form.DLId.value + bank_code + YM;
	var i = form.UpFileName.value.length - FILE_NAME_LENGTH;

	if (form.UpFileName.value.substring(i) != fileName) {
		alert('檔名必須為' + fileName)
		form.UpFileName.focus();
		return false;
	}
	*/
	if (form.DLId.value == 'A05' || form.DLId.value == 'B05') {
		if ((form.S_MONTH.value != '06') && (form.S_MONTH.value != '12')) {
			alert("半年報基準日必須為6月或12月!")
			return false;
		}
	}
	if (form.DLId.value == 'F12' || form.DLId.value == 'F13' || form.DLId.value == 'F14' || form.DLId.value == 'F15' ||
		form.DLId.value == 'F16' || form.DLId.value == 'F17') {
		if ((form.S_MONTH.value != '03') && (form.S_MONTH.value != '06') && (form.S_MONTH.value != '09') && (form.S_MONTH.value != '12')) {
			alert("季報基準日必須為3,6,9,12月!")
			return false;
		}
	}
	if (trimString(form.INPUT_NAME.value) == "") {
		alert("申報者姓名必須輸入")
		form.INPUT_NAME.focus();
		return false;
	}
	if (trimString(form.INPUT_TEL.value) == "") {
		alert("申報者電話必須輸入")
		form.INPUT_TEL.focus();
		return false;
	}
	if (trimString(form.UpFileName.value) == "") {
		alert("上傳檔案位置必須輸入")
		form.UpFileName.focus();
		return false;
	}
	
	form.act.value="Upload";	
	
	form.action="/pages/FileUpload.jsp?act="+form.act.value+"&FileName="+form.UpFileName.value+"&test=nothing";
	return true;
}
