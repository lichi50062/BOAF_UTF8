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
		if (confirm('����s�@�~���������ɶ�(�C10�ӳ���40��,�̦�����)\n�Фŭ����I��[����],�T�w�����ˮ֧@�~��??')) {	   
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