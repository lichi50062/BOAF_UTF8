/*
 *  created by lilic0c0 4183
 *  95.11.01 ADD 列印 by 2495 
 *  95.11.15 FIX 承受日期不檢查 by 2495 
 *  96.04.12 fix 當為第1季時,上一季資料為年度-1,季別為第4季 by 2295
 */
 
 //globe var 
 var isEmptyArr = new Array(); //存放此年季有沒有值
 
 //=========================================================	
 //do request to FX008W
function printSubmit(form,bank_name,dureassure_no,debtname,year,month,day,dureassuresite,accountamt,applydelayyear_month,applydelayreason,damage_yn,disposal_fact_yn,disposal_plan_yn){
/*
	alert(bank_name);
	alert(dureassure_no);
	alert(debtname);
	alert(year);
	alert(month);
	alert(day);
	alert(dureassuresite);
	alert(accountamt);
	alert(applydelayyear_month);
	alert(applydelayreason);
	alert(damage_yn);
	alert(disposal_fact_yn);
	alert(disposal_plan_yn);
*/	
	form.action="/pages/FX008W_Print.jsp?bank_name="+bank_name+"&dureassure_no="+dureassure_no+"&debtname="+debtname+"&year="+year+"&month="+month+"&day="+day+"&dureassuresite="+dureassuresite+"&accountamt="+accountamt+"&applydelayyear_month="+applydelayyear_month+"&applydelayreason="+applydelayreason+"&damage_yn="+damage_yn+"&disposal_fact_yn="+disposal_fact_yn+"&disposal_plan_yn="+disposal_plan_yn;	
	form.submit();
	
	
	
	}
function doSubmit(form,cnd,parameter){
	
	    form.action="/pages/FX008W.jsp?act="+cnd+"&bank_no="+parameter+"&test=nothing"; 
	       
	    if( cnd == "Insert" && (!checkData(form)) )return;	   	    
	    if( cnd == "No_Apply" && check_No_Apply(form) ) form.submit();	   	    
	    if( cnd == "Load" &&check_Apply_Ini(form)&& checkload(form)) form.submit();
	    
	    if( cnd == "Load_To"){
	    	form.action="/pages/FX008W.jsp?act="+cnd+"&form_size="+form.elements.length+"&test=nothing";
	    	form.submit();
	    }
	    
	    if( cnd == "New" && check_Apply_Ini(form)) form.submit();	 
	       	
	    if((cnd == "Insert") && (checkData(form)) &&AskInsert(form)){
	    	form.account.value = changeVal(form.account)
	    	form.submit();
	    }	    	    
	    if((cnd == "Update") && AskUpdate(form)&& checkData(form)){
	    	form.account.value = changeVal(form.account)
	    	form.submit();
	    }	   
	    if((cnd == "Delete") && AskDelete(form)) form.submit();
	    
	    if(cnd == "load_history") { 
    	//alert(form.dure_no.value);    	
			//form.action="/pages/FX008AW.jsp?act="+cnd+"&debt_name="+form.debtname.value; 
			//form.submit();	
    	window.open("/pages/FX008W.jsp?act="+cnd+"&debt_name="+form.debtname.value+"&dure_no="+form.dure_no.value,"LOAD","width=750,height=400,scrollbars =yes")
    }
	    	       	    
}
//=========================================================	
//arrpush store year quarter data to array
function arrpush(inputData){
	isEmptyArr[isEmptyArr.length]=trimString(inputData);
}

//=========================================================
//re-checkload
//96.04.12 fix 當為第1季時,上一季資料為年度-1,季別為第4季 by 2295
function checkload(form){
	
	var year = parseInt(form.s_year.value,10);
	var quarter = parseInt(form.s_quarter.value,10);
	var l_year = year;
	var l_quarter = quarter-1;
	var flag = "true";
	//alert(year);
	//alert(quarter);
	if(l_quarter == 0){
		l_year--;
		l_quarter=4;
	}
	//alert(l_year);
	//alert(l_quarter);
	year=year+"";
	quarter=quarter+"";
	l_year=l_year+"";
	l_quarter=l_quarter+"";
	
	if(!confirm("確定載入上季資料?") ){
		form.m_year.focus();
		return false;	
	}
	
	for(var i=isEmptyArr.length-1; i >= 0 ;i-- ){

		if(trimString(isEmptyArr[i])== year+quarter){
			alert("本季已有資料,不能再載入上季資料");	
			return false;
		}
		else if(trimString(isEmptyArr[i])== l_year+l_quarter){
			flag = "false";
		}
	}
	
	if(trimString(flag)=="true"){
		alert("上季沒有資料,無法載入上季資料");
		return false;
	}
	
	return true;
}

//=========================================================
//check this year and quarter is not allowed to insert or new
function check_Apply_Ini(form){
	var apply_ini = parseInt(form.apply_ini.value,10);
	var year = parseInt(form.s_year.value,10);
	var month = parseInt(form.s_quarter.value,10);

	if(apply_ini > year*12+month){
			alert("你鍵入的申報年季超出申報項目起始年季");		
			return false;
	}
	
	return true;
}

//=========================================================
//再次確認有沒有資料
function check_No_Apply(form){
	if(!confirm("確定本季無資料?")){
			form.no_apply_yn.focus();
			return false;
	}
	return true;
}

//=========================================================	
function checkData(form) 
{	
	if (trimString(form.m_year.value) =="" ){
		alert("申報年份不可空白");
		form.m_year.focus();
		return false;
	}
	
	if (trimString(form.m_quarter.value) =="" ){
		alert("申報季不可空白");
		form.m_quarter.focus();
		return false;
	}		
	
	if (trimString(form.dure_no.value) =="" ){
		alert("承受擔保品編號不可空白");
		form.dure_no.focus();
		return false;
	}
	
	if( isNaN(Math.abs(form.dure_no.value)) ){
		alert("承受擔保品編號不可為文字");
		form.dure_no.focus();
		return false;
	}
	
	if (trimString(form.debtname.value) =="" ){
		alert("機構名稱(借款人)不可空白");
		form.debtname.focus();
		return false;
	}
	
	var accept_Date= "承受日期";
	if(!checkDate(form.accept_year,form.accept_month,form.accept_day,accept_Date,0)){
		return false;
	}
	
	// 95.11.15 FIX 承受日期不檢查 by 2495 
	/*
	if(checkCurrentDate(form.accept_year,form.accept_month,form.accept_day)){
		form.accept_year.focus();
		return false;
	}
*/
	if (trimString(form.duresite.value) =="" ){
		alert("承受擔保品座落不可空白");
		form.duresite.focus();
		return false;
	}
	
	if (trimString(form.account.value) =="" ){
		alert("帳列金額不可空白");
		form.account.focus();
		return false;
	}
	
	if( isNaN(Math.abs(changeVal(form.account)))){
		alert("帳列金額不可為文字");
		form.account.focus();
		return false;
	}
	
	if (trimString(form.apply_year.value) =="" ){
		alert("申請延長期間不可空白");
		form.apply_year.focus();
		return false;
	}
	
	if(isNaN(Math.abs(form.apply_year.value))){
		alert("申請延長期間不可為文字");
		form.apply_year.focus();
		return false;
	}
	
	if (parseInt(form.apply_year.value)>4 ){
		form.apply_year.focus();
		return confirm("申請延長期間超過4年，確定嗎？");
	}
	
	if (trimString(form.apply_month.value) =="" ){
		alert("申請延長期間不可空白");
		form.apply_month.focus();
		return false;
	}
		
	if (trimString(form.apply_reason.value) =="" ){
		alert("申請延長理由不可空白");
		form.apply_reason.focus();
		return false;
	
    }
		
   return true;
}

//=========================================================
function AskDelete(form) {
	return confirm("確定要刪除嗎");
}

//=========================================================
function checknan(S1){
	
    if (isNaN(Math.abs(S1.value)))
   	    alert("<承受擔保品編號>請輸入數字");
	
	return S1.value;
}
//=========================================================

function checkCurrentDate(year,month,day){
	
	var today = new Date();
	var nowYear = today.getYear()-1911; //change Anno Domini(A.D.) to ROC date
	var nowMonth = today.getMonth()+1; //begin with 0
	var nowDate = today.getDate();
	var totalNowDate = nowYear*12*30+nowMonth*30+nowDate+0;
	var totalTestDate = parseInt(year.value,10)*12*30+parseInt(month.value,10)*30+parseInt(day.value,10)+0;
	
	//alert(totalNowDate+" "+ nowYear*12*30+" "+nowMonth*30+" "+nowDate+" |"+" "+totalTestDate+parseInt(year.value,10)*12*30+" "+parseInt(month.value,10)*30+" "+parseInt(day.value,10));
	if(totalTestDate > totalNowDate){
		alert("輸入之承受日期大於本日日期，請再次確認");
		return true;
	}
	
	return false;
}
