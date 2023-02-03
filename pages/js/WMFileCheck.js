function doSubmit(form,fun,bank_type) {
	if(fun == 'Query'){
		form.action="/pages/WMFileCheck.jsp?act="+fun+"&bank_type="+bank_type+"&test=nothing";
		form.submit();
	}	
	if(fun == 'Check'){
		if (form.YM[0].checked) {
			if (!checkSingleYM(form.S_YEAR, form.S_MONTH))
				return;
		}
		
		if (form.Report_no.value == "")	return false;
		if (confirm('此更新作業須較長的時間(每10個單位約40秒,依此類推)\n請勿重複點選[執行],確定執行檢核作業嗎??')) {	   
		     form.action="/pages/WMFileCheck.jsp?act="+fun+"&bank_type="+bank_type+"&test=nothing";
		     form.submit();
		}     
    }
	
}


function checkYM(form) {	
	if (form.YM[0].checked) {
		form.S_YEAR.disabled = false;
		form.S_MONTH.disabled = false;
	}else {
		form.S_YEAR.disabled = true;
		form.S_MONTH.disabled = true;
	}
}