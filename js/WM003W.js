function checkWM003W(form){
    if(!CheckQueryDate2(form.S_YEAR,form.S_MONTH,form.S_DATE,"�_�l���"))
            return false;
    if(!CheckQueryDate2(form.E_YEAR,form.E_MONTH,form.E_DATE,"�������"))
            return false;
	if(trimString(form.S_DATE.value)!="" && trimString(form.E_DATE.value)!="")
	{
		if(Math.abs(form.S_DATE.value) > Math.abs(form.E_DATE.value))
    	{
    		alert("�_�l������i�H�j��פ���");
    		return false;
    	}
    }
    return true;
}

