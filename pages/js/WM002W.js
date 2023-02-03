function doSubmitWM002W(DLId, DLIdName) {

	var form = document.frmWM002W;
	form.DLId.value = DLId;
	form.DLIdName.value = DLIdName;
	form.action = 'WM002W_LIST';
	form.submit();
}

function doSubmitWM002W_LIST(form, fun) {

	var IDname;

	if (form.DLId.value == 'A02' || form.DLId.value == 'B02')
		IDname = 'A02';
	else if (form.DLId.value == 'A03' || form.DLId.value == 'B03' || form.DLId.value == 'G03')
		IDname = 'A03';
	else if (form.DLId.value == 'A04' || form.DLId.value == 'B04' || form.DLId.value == 'G04')
		IDname = 'A04';
	else if (form.DLId.value == 'A05' || form.DLId.value == 'B05')
		IDname = 'A05';
	else if (form.DLId.value == 'F05' || form.DLId.value == 'F06' || form.DLId.value == 'F10' || form.DLId.value == 'F11' ||
			 form.DLId.value == 'F12' || form.DLId.value == 'F13' ||
			 form.DLId.value == 'F15' || form.DLId.value == 'F16' || form.DLId.value == 'F17')
		IDname = 'F06';
	else
		IDname = form.DLId.value;

	if (fun == 'insert')
		form.action = 'WM002W_ShowInsert_' + IDname;
	else {
		var i = fun.length;
		var j = fun.indexOf("/");
		if (j != -1) {
			form.S_YEAR.value = fun.substring(0, j);
			form.S_MONTH.value = fun.substring(j + 1);
		}
		if (form.DLId.value == 'G02')
			form.action = 'WM002W_LIST01_G02';
		else if (form.DLId.value == 'G05')
			form.action = 'WM002W_LIST01_G05';
		else
			form.action = 'WM002W_ShowUpdate_' + IDname;
	}
	form.submit();
}
function doSubmitWM002W_LIST01_G02(form, cust_form, beg_date, over_date) {

	form.cust_form.value = cust_form;
	form.beg_date.value = beg_date;
	form.over_date.value = over_date;
	form.action = 'WM002W_ShowUpdate_G02';
	form.submit();
}
function doSubmitWM002W_LIST01_G05(form, cust_form) {

	form.cust_form.value = cust_form;
	form.action = 'WM002W_ShowUpdate_G05';
	form.submit();
}

function checkShowInsert_A01(form, fun) {

    if (fun == 'delete') {
    	if (AskDelete()) {
    		form.Function.value = 'delete';
            form.submit();
            return true;
        } else
        	return false;

    }

	if (!Check_Maintain(form))
		return false;
	if (trimString(form.maintain_email.value) == "") {
		alert("承辦員E_MAIL必須輸入")
		form.maintain_email.focus();
		return false;
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

	if (fun == 'insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

//	for (var i =0; i < form.length; i++){
//		if (form.elements[i].name == 'amt')	{
//			if (!checkNumber(form.elements[i]))
//				return false;
//		}
//	}
	for (var i = 0; i < form.amt.length; i++) {
		if (!checkNumber(form.amt[i]))
			return false;
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

	return true;
}

function checkShowInsert_A02(form, fun) {

    if (fun == 'delete') {
    	if (AskDelete()) {
    		form.Function.value = 'delete';
            form.submit();
            return true;
        } else
        	return false;

    }

	if (!Check_Maintain(form))
		return false;
	if (trimString(form.maintain_email.value) == "") {
		alert("承辦員E_MAIL必須輸入")
		form.maintain_email.focus();
		return false;
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

	if (fun == 'insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	for (var i =0; i < form.length; i++){
		if (form.elements[i].name == 'm_rate')	{
			var m_rate = changeVal(form.elements[i]);
		    if (isNaN(Math.abs(m_rate))) {
		   	    alert("請輸入數字");
		   	    return false;
		    }
			if (!checkRate000(form.elements[i]))
				return false;
		}
	}
	return true;
}
function checkShowInsert_A03(form, fun) {

    if (fun == 'delete') {
    	if (AskDelete()) {
    		form.Function.value = 'delete';
            form.submit();
            return true;
        } else
        	return false;

    }

	if (!Check_Maintain(form))
		return false;
	if (trimString(form.maintain_email.value) == "") {
		alert("承辦員E_MAIL必須輸入")
		form.maintain_email.focus();
		return false;
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

	if (fun == 'insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	for (var i =0; i < form.length; i++){
		if (form.elements[i].name == 'rate') {
			var m_rate = changeVal(form.elements[i]);
		    if (isNaN(Math.abs(m_rate))) {
		   	    alert("請輸入數字");
		   	    return false;
		    }
			if (!checkRate00(form.elements[i]))
				return false;
		}
	}
	return true;
}
function checkShowInsert_A04(form, fun) {

    if (fun == 'delete') {
    	if (AskDelete()) {
    		form.Function.value = 'delete';
            form.submit();
            return true;
        } else
        	return false;

    }

	if (!Check_Maintain(form))
		return false;

	if (trimString(form.maintain_email.value) == "") {
		alert("承辦員E_MAIL必須輸入")
		form.maintain_email.focus();
		return false;
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

	if (fun == 'insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	for (var i =0; i < form.length; i++){
		if ((form.elements[i].name == 'amt') || (form.elements[i].name == 'amt1')) {
			if (!checkNumber(form.elements[i]))
		   	    return false;
		}
	}
	return true;
}
function checkShowInsert_A05(form, fun) {

    if (fun == 'delete') {
    	if (AskDelete()) {
    		form.Function.value = 'delete';
            form.submit();
            return true;
        } else
        	return false;

    }

	if (!Check_Maintain(form))
		return false;

	if (trimString(form.maintain_email.value) == "") {
		alert("承辦員E_MAIL必須輸入")
		form.maintain_email.focus();
		return false;
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

	if (fun == 'insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

//	for (var i =0; i < form.length; i++){
//		if (form.elements[i].name == 'amt') {
//			if (!checkNumber(form.elements[i]))
//		   	    return false;
//		}
//	}
	for (var i = 0; i < form.amt.length; i++) {
		if (!checkNumber(form.amt[i]))
			return false;
	}
	return true;
}
function checkShowInsert_G02(form, fun) {

    if (fun == 'delete') {
    	if (AskDelete()) {
    		form.Function.value = 'delete';
            form.submit();
            return true;
        } else
        	return false;
    }

	if (!Check_Maintain(form))
		return false;

	if (trimString(form.maintain_email.value) == "") {
		alert("承辦員E_MAIL必須輸入")
		form.maintain_email.focus();
		return false;
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

	if (fun == 'insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
		if (!CheckYear(form.beg_y))
			return false;
		if (!CheckYear(form.over_y))
			return false;
		if (trimString(form.cust_form.value) == "") {
			alert("統一編號或身分證字號必須輸入")
			return false;
		}
		if (trimString(form.cust_name.value) == "") {
			alert("大額授信客戶名稱必須輸入")
			return false;
		}
		if (trimString(form.beg_y.value) == "") {
			alert("開始動用日(年)必須輸入")
			return false;
		}
		if (trimString(form.beg_m.value) == "") {
			alert("開始動用日(月)必須輸入")
			return false;
		}
		if (trimString(form.beg_d.value) == "") {
			alert("開始動用日(日)必須輸入")
			return false;
		}
		if (trimString(form.over_y.value) == "") {
			alert("開始積欠日(年)必須輸入")
			return false;
		}
		if (trimString(form.over_m.value) == "") {
			alert("開始積欠日(月)必須輸入")
			return false;
		}
		if (trimString(form.over_d.value) == "") {
			alert("開始積欠日(日)必須輸入")
			return false;
		}
		var chkDate;
		chkDate =  '' + (parseInt(form.beg_y.value) + 1911) + '/' + form.beg_m.value + '/' + form.beg_d.value;
		if (fnValidDate(chkDate) != true) {
			alert('所輸入的開始動用日為無效日期!!');
    	    form.beg_d.focus();
        	return false;
   		}
		chkDate =  '' + (parseInt(form.over_y.value) + 1911) + '/' + form.over_m.value + '/' + form.over_d.value;
		if (fnValidDate(chkDate) != true) {
			alert('所輸入的開始積欠日為無效日期!!');
    	    form.over_d.focus();
        	return false;
   		}
	} //End of if

	for (var i =0; i < form.length; i++){
		if (form.elements[i].name == 'amt')	{
			if (!checkNumber(form.elements[i]))
				return false;
		}
	}
	form.Function.value = fun;
	if (fun == 'update')
		form.submit();

	return true;
}
function checkShowInsert_G05(form, fun) {

    if (fun == 'delete') {
    	if (AskDelete()) {
    		form.Function.value = 'delete';
            form.submit();
            return true;
        } else
        	return false;
    }

	if (!Check_Maintain(form))
		return false;

	if (trimString(form.maintain_email.value) == "") {
		alert("承辦員E_MAIL必須輸入")
		form.maintain_email.focus();
		return false;
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

	if (fun == 'insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
		if (trimString(form.cust_form.value) == "") {
			alert("統一編號或身分證字號必須輸入")
			return false;
		}
		if (trimString(form.cust_name.value) == "") {
			alert("大額授信客戶名稱必須輸入")
			return false;
		}
	} //End of if

	for (var i =0; i < form.length; i++){
		if (form.elements[i].name == 'amt')	{
			if (!checkNumber(form.elements[i]))
				return false;
		}
	}
	form.Function.value = fun;
	if (fun == 'update')
		form.submit();
	return true;
}
function checkShowInsert_H02(form, fun) {

    if (fun == 'delete') {
    	if (AskDelete()) {
    		form.Function.value = 'delete';
            form.submit();
            return true;
        } else
        	return false;
    }

	if (!Check_Maintain(form))
		return false;

	if (trimString(form.maintain_email.value) == "") {
		alert("承辦員E_MAIL必須輸入")
		form.maintain_email.focus();
		return false;
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

	if (fun == 'insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	for (var i =0; i < form.length; i++){
		if (form.elements[i].name == 'amt') {
			if (!checkNumber(form.elements[i]))
		   	    return false;
		}
	}
	return true;
}
function checkShowInsert_H03(form, fun) {

    if (fun == 'delete') {
    	if (AskDelete()) {
    		form.Function.value = 'delete';
            form.submit();
            return true;
        } else
        	return false;
    }

	if (!Check_Maintain(form))
		return false;

	if (trimString(form.maintain_email.value) == "") {
		alert("承辦員E_MAIL必須輸入")
		form.maintain_email.focus();
		return false;
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

	if (fun == 'insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	for (var i =0; i < form.length; i++){
//		if ((form.elements[i].name.substr(7) == 'amt') || (form.elements[i].name.substr(7) == 'amt1')) {
		if ((form.elements[i].name == 'amt') || (form.elements[i].name == 'amt1')) {
			if (!checkNumber(form.elements[i]))
		   	    return false;
		}
//		if (form.elements[i].name.substr(7) == 'rate') {
		if (form.elements[i].name == 'rate') {
			var m_rate = changeVal(form.elements[i]);
		    if (isNaN(Math.abs(m_rate))) {
		   	    alert("請輸入數字");
		   	    return false;
		    }
			if (!checkRate00(form.elements[i]))
				return false;
		}

	}
	return true;
}

function checkShowInsert_H04(form, fun) {

    if (fun == 'delete') {
    	if (AskDelete()) {
    		form.Function.value = 'delete';
            form.submit();
            return true;
        } else
        	return false;
    }

	if (!Check_Maintain(form))
		return false;

	if (trimString(form.maintain_email.value) == "") {
		alert("承辦員E_MAIL必須輸入")
		form.maintain_email.focus();
		return false;
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

	if (fun == 'insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	for (var i =0; i < form.length; i++){
		if (form.elements[i].name == 'amt') {
			if (!checkNumber(form.elements[i]))
		   	    return false;
		}
	}
	return true;
}
function checkShowInsert_H05(form, fun) {

    if (fun == 'delete') {
    	if (AskDelete()) {
    		form.Function.value = 'delete';
            form.submit();
            return true;
        } else
        	return false;
    }

	if (!Check_Maintain(form))
		return false;

	if (trimString(form.maintain_email.value) == "") {
		alert("承辦員E_MAIL必須輸入")
		form.maintain_email.focus();
		return false;
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

	if (fun == 'insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	for (var i =0; i < form.length; i++){
		if (form.elements[i].name == 'amt') {
			if (!checkNumber(form.elements[i]))
		   	    return false;
		}
	}
	return true;
}
function checkShowInsert_H06(form, fun) {

    if (fun == 'delete') {
    	if (AskDelete()) {
    		form.Function.value = 'delete';
            form.submit();
            return true;
        } else
        	return false;
    }

	if (!Check_Maintain(form))
		return false;

	if (trimString(form.maintain_email.value) == "") {
		alert("承辦員E_MAIL必須輸入")
		form.maintain_email.focus();
		return false;
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

	if (fun == 'insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	for (var i =0; i < form.length; i++){
		if (form.elements[i].name == 'amt') {
			if (!checkNumber(form.elements[i]))
		   	    return false;
		}
	}
	return true;
}
function checkShowInsert_H07(form, fun) {

    if (fun == 'delete') {
    	if (AskDelete()) {
    		form.Function.value = 'delete';
            form.submit();
            return true;
        } else
        	return false;
    }

	if (!Check_Maintain(form))
		return false;

	if (trimString(form.maintain_email.value) == "") {
		alert("承辦員E_MAIL必須輸入")
		form.maintain_email.focus();
		return false;
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

	if (fun == 'insert') {
		if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
			return false;
	}

	for (var i =0; i < form.length; i++){
		if (form.elements[i].name == 'amt') {
			if (!checkNumber(form.elements[i]))
		   	    return false;
		}
		if (form.elements[i].name == 'rate') {
			var m_rate = changeVal(form.elements[i]);
		    if (isNaN(Math.abs(m_rate))) {
		   	    alert("請輸入數字");
		   	    return false;
		    }
			if (!checkRate00(form.elements[i]))
				return false;
		}
	}
	return true;
}
//=========================================================
function checkRate000(Rate1) {

	var rate = changeVal(Rate1)

    if (isNaN(Math.abs(rate))) {
// 	    alert("請輸入數字");
   	    return;
    }
    if (rate > 100.000 || rate < -99.999) {
        alert("利率不可大於 100.000, 也不可小於 -99.999")
        return;
    }
    if (rate.indexOf(".") != -1 ) {
        len = (rate.substring(rate.indexOf(".") + 1, rate.length));
        if (len.length > 3) {
            alert("小數點後只能有三個位數");
            return;
        }
    }
    return true;
}
//=========================================================
function checkRate00(Rate1) {

	var rate = changeVal(Rate1)

    if (isNaN(Math.abs(rate))) {
// 	    alert("請輸入數字");
   	    return;
    }
    if (rate > 100.00 || rate < -99.99) {
        alert("利率不可大於 100.00, 也不可小於 -99.99")
        return;
    }
    if (rate.indexOf(".") != -1 ) {
        len = (rate.substring(rate.indexOf(".") + 1, rate.length));
        if (len.length > 2) {
            alert("小數點後只能有二個位數");
            return;
        }
    }
    return true;
}
