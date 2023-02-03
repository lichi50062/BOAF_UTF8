function doSubmit(form) {	
	if(checkQuery(form)){		
	    form.action="/pages/WMFileQuery.jsp?act=Query&test=nothing";	
	    form.submit();
    }
}

function checkQuery(form){
    if(!CheckQueryDate2(form.S_YEAR,form.S_MONTH,form.S_DATE,"起始日期"))
            return false;
    if(!CheckQueryDate2(form.E_YEAR,form.E_MONTH,form.E_DATE,"結束日期"))
            return false;
	if(trimString(form.S_DATE.value)!="" && trimString(form.E_DATE.value)!="")
	{
		if(Math.abs(form.S_DATE.value) > Math.abs(form.E_DATE.value))
    	{
    		alert("起始日期不可以大於終止日期");
    		return false;
    	}
    }
    return true;
}

