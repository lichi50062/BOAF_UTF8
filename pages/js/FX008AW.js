/*
 * Created on 94.11.24 by lilic0c0 4183
 * FIX 函報文號 不檢查 on 95.11.03 by 2495
 * 98.01.07 fix 不自動計算核准申請延長期限(A)+(B) by 2295
 */
//==========================================================
function doSubmit(form,cnd,parameter){
  
	form.action="/pages/FX008AW.jsp?act="+cnd+"&bank_no="+parameter+"&test=nothing"; 
    //98.01.07 fix 不自動計算核准申請延長期限(A)+(B)calculusDate(form);    
	if(cnd == "Update") {
		if(!checkUpdate(form)) 
	  		return;
	  	
	  	if(!AskUpdate(form))
	    	return;
	    	
		if(!confirm("所有資料都確定了嗎"))
    		return;  	
    	form.submit();
    }
    if(cnd == "load") { 
    	//alert(form.dure_no.value);    	
			//form.action="/pages/FX008AW.jsp?act="+cnd+"&debt_name="+form.debtname.value; 
			//form.submit();	
    	window.open("/pages/FX008AW.jsp?act="+cnd+"&debt_name="+form.debtname.value+"&dure_no="+form.dure_no.value,"LOAD","width=750,height=400,scrollbars =yes")
    }
}	

//==========================================================
function checkUpdate(form) 
{	
	var str1 = "核准日期";
	var str2 = "函報日期";
	
	if( trimString(form.DocNo.value) == "" ){
		alert("核准文號不可空白");
		form.DocNo.focus();
		return false;
	}
	else if (parseInt(form.Audit_Delay_y.value)>4 ){
		form.Audit_Delay_y.focus();
		return confirm("核准日期超過4年，確定嗎？");
	}
	else if( isNaN(Math.abs(form.ApplyOK_y.value)) ){
		alert("核准日期不可為文字");
		form.ApplyOK_y.focus();
		return false;
	}
	else if(!checkDate(form.ApplyOK_y,form.ApplyOK_m,form.ApplyOK_d,str1,0)){
		return false;
	}
	else if(checkCurrentDate(form.ApplyOK_y,form.ApplyOK_m,form.ApplyOK_d,str1)){
		return false;
	}
	// FIX 函報文號 不檢查 on 95.11.03 by 2495
	/*
	else if( trimString(form.BOAF_DocNo.value) == "" ){
		alert("函報文號不可空白");
		form.BOAF_DocNo.focus();
		return false;
	}
	*/
	else if( isNaN(Math.abs(form.BOAF_y.value)) ){
		alert("函報日期不可為文字");
		form.BOAF_y.focus();
		return false;
	}
	else if(!checkDate(form.BOAF_y,form.BOAF_m,form.BOAF_d,str2,0)){
		return false;
	}
	else if(checkCurrentDate(form.BOAF_y,form.BOAF_m,form.BOAF_d,str2)){
		return false;
	}
	else if( trimString(form.damage_yn.value) == "" &&
		trimString(form.disposal_fact_yn.value) == "" &&
		trimString(form.disposal_plan_yn.value) == "" &&
		trimString(form.auditresult_yn.value) == ""){
			
		form.damage_yn.focus();
		return confirm("審核選項欄均為空白，各機關將可以繼續修改資料，確定嗎？");
	}
	else if( trimString(form.auditresult_yn.value) != "Y" &&
		trimString(form.auditresult_yn.value) != "N"){
		alert("審核結果不可空白");
		form.auditresult_yn.focus();
		return false;
	}
	else if( trimString(form.damage_yn.value) != "Y" &&
		trimString(form.damage_yn.value) != "N"){
		alert("是否提足備抵趺價損失不可空白");
		form.damage_yn.focus();
		return false;
	}
	else if( trimString(form.disposal_fact_yn.value) != "Y" &&
		trimString(form.disposal_fact_yn.value) != "N"){
		alert("是否有積極處分之事實不可空白");
		form.disposal_fact_yn.focus();
		return false;
	}
	
	else if( trimString(form.disposal_plan_yn.value) != "Y" &&
		trimString(form.disposal_plan_yn.value) != "N"){
		alert("未來處分計劃是否合理可行不可空白");
		form.disposal_plan_yn.focus();
		return false;
	}

	else if((trimString(form.damage_yn.value) == "N" ||
			 trimString(form.disposal_fact_yn.value) == "N" ||
			 trimString(form.disposal_plan_yn.value) == "N" )&&
			 trimString(form.auditresult_yn.value) == "Y"      ){
		form.auditresult_yn.focus();
		return confirm("審核選項欄有選項為「否」，而「審核結果」卻為「合格」，確定(不合格/合格)？");
	}
	
	else if( trimString(form.damage_yn.value) == "Y" &&
			 trimString(form.disposal_fact_yn.value) == "Y" &&
			 trimString(form.disposal_plan_yn.value) == "Y" &&
			 trimString(form.auditresult_yn.value) == "N"){
		form.auditresult_yn.focus();
		return confirm("審核選項欄均為「是」，而「審核結果」卻為「不合格」，確定(合格/不合格)？");
	}
	
	return true;

}

//==========================================================
function loadUpperDate(form){
	form.Audit_Delay_y.value = form.apply_year.value;
	form.Audit_Delay_m.value = form.apply_month.value;
}

//==========================================================
function calculusDate(form){
	
	var accept_y = parseInt(form.accept_year.value);
	var accept_m = parseInt(form.accept_month.value);
	var accept_d = parseInt(form.accept_day.value);
	var delay_y = parseInt(form.Audit_Delay_y.value);
	var delay_m = parseInt(form.Audit_Delay_m.value);
	var result_y = 0; 
	var result_m = 0;
	var result_d = 0;
	
	result_y = accept_y + delay_y;
	result_m = accept_m + delay_m;
	result_d = accept_d;
	 
	if(result_m > 12){//月份超過一年
		result_m-=12;
		result_y++;
	}
	
	//如果算出來是31號卻又遇到月份是小月 
	if(result_d==31){
		
		switch(result_m){
			case 2:
				result_d=28;
				if(leapYear(parseInt(result_y,10)+1911))
					result_d=29;
				break;
			case 4:
			case 6:
			case 9:
			case 11:
				result_d=30;
				break;
		}//end of switch
	}//end of if
	
	form.Cnt_year.value	= result_y;
	form.Cnt_month.value= result_m;
	form.Cnt_date.value	= result_d;
}

//====抄錄最近一期的函報資料======================================================
function LoadRecentData(form,xml){
	
	var xmlDoc = document.all(xml).XMLDocument; //取得xml data
	var flag =  xmlDoc.getElementsByTagName("flag");
	
	if(flag.item(0) == null){
		alert("「無最近一筆資料，無從抄錄請自行輸入。」");
		form.DocNo.focus();
		return;
	}
	else{//進行抄錄
		//先去取得xml中的值
		var docno =  xmlDoc.getElementsByTagName("docno");
		var boafdocno =  xmlDoc.getElementsByTagName("boafdocno");
		var apoky =  xmlDoc.getElementsByTagName("apoky");
		var apokm =  xmlDoc.getElementsByTagName("apokm");
		var apokd =  xmlDoc.getElementsByTagName("apokd");
		var boafy =  xmlDoc.getElementsByTagName("boafy");
		var boafm =  xmlDoc.getElementsByTagName("boafm");
		var boafd =  xmlDoc.getElementsByTagName("boafd");
		
		//做抄錄的動作
		form.DocNo.value = docno.item(0).firstChild.nodeValue;
		form.BOAF_DocNo.value = boafdocno.item(0).firstChild.nodeValue;
		
		form.ApplyOK_y.value = apoky.item(0).firstChild.nodeValue;
		setSelect(form.ApplyOK_m,apokm.item(0).firstChild.nodeValue);
		setSelect(form.ApplyOK_d,apokd.item(0).firstChild.nodeValue);
		form.BOAF_y.value = boafy.item(0).firstChild.nodeValue;
		setSelect(form.BOAF_m,boafm.item(0).firstChild.nodeValue);
		setSelect(form.BOAF_d,boafd.item(0).firstChild.nodeValue);
	}

}
//=========================================================
function checkMonth(month){
	if(month.value < 0 ){		
		alert("月份不可小於零");
		return false;
	}
	else if(month.value > 12){
		alert("月份不可大於12個月");
		return false;
	}

	return true;

}
//=========================================================
function setSelect(S1,bankid) {

    for(i=0;i<S1.length;i++) {
      	if(S1.options[i].value==bankid)    	{
        	S1.options[i].selected=true;
        	break;
    	}
    }
}

//=========================================================

function checkCurrentDate(year,month,day,txt){
	var today = new Date();
	var nowYear = today.getYear()-1911; //change Anno Domini(A.D.) to ROC date
	var nowMonth = today.getMonth()+1;//begin with 0
	var nowDate = today.getDate();
	var totalNowDate = nowYear*12*30+nowMonth*30+nowDate+0;
	var totalTestDate = parseInt(year.value,10)*12*30+parseInt(month.value,10)*30+parseInt(day.value,10)+0;
	
	//alert( totalTestDate);
	if(totalTestDate > totalNowDate){
		alert("輸入之『"+txt+"』大於本日日期，請再次確認");
		return true;
	}
	
	return false;
}

	