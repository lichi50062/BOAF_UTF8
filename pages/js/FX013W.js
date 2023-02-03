/*
 *  created by 2660
 */

function AskInsert() {
  if(confirm("確定要新增此筆資料嗎？"))
    return true;
  else
    return false;
}

function AskUpdate() {
  if(confirm("確定要修改此筆資料嗎？"))
    return true;
  else
    return false;
}

function AskDelete() {
  if(confirm("確定要刪除此筆資料嗎？"))
    return true;
  else
    return false;
}

function doSubmit(form, cnd){
  form.action="/pages/FX013W.jsp?act="+cnd;

  if ( (cnd == "Insert") && ( checkData(form)) && AskInsert() ) {
    form.submit();
  }
  else if ( (cnd == "Update") && ( checkData(form)) && AskUpdate() ) {
    form.submit();
  }
  else if ( (cnd == "Delete") && AskDelete() ) {
    form.submit();
  }
  else if (cnd == "New") {
    form.submit();
  }
}

function checkData(form) 
{
  if (trimString(form.m_year.value) != "" ) {
	  if (isNaN(Math.abs(form.m_year.value))) {
		  alert("申報年份不可為文字");
		  form.m_year.focus();
		  return false;
		}
  } else {
		alert("申報年份不可空白");
		form.m_year.focus();
		return false;   
	}
	
  if (trimString(form.m_month.value) != "" ) {
	  if (isNaN(Math.abs(form.m_month.value))) {
		  alert("申報月份不可為文字");
		  form.m_month.focus();
		  return false;
		}
  } else {
		alert("申報月份不可空白");
		form.m_month.focus();
		return false;   
	}
	
	if (trimString(form.tbank_no.value) == "" ){
		alert("總機構代號不可空白");
		form.tbank_no.focus();
		return false;
	}
	
	if (trimString(form.bank_no.value) == "" ){
		alert("總分支機構代號不可空白");
		form.bank_no.focus();
		return false;
	}
	
	if (trimString(form.name.value) == "" ){
		alert("學員姓名不可空白");
		form.name.focus();
		return false;
	}
	
	if (trimString(form.position_code.value) == "" ){
		alert("職稱不可空白");
		form.position_code.focus();
		return false;
	} else if (trimString(form.tbank_no.value) == trimString(form.bank_no.value)){
	  if (trimString(form.position_code.value) == "2"){
		  alert("輸入資料為總機構時,職稱不可為信用部分部主任");
		  form.position_code.focus();
		  return false;
	  }
	} else {
	  if (trimString(form.position_code.value) == "1"){
		  alert("輸入資料為分支機構時,職稱不可為信用部主任");
		  form.position_code.focus();
		  return false;
	  }
	}
	
	if (trimString(form.train_place.value) == "" ){
		alert("訓練機構不可空白");
		form.train_place.focus();
		return false;
	} else if (trimString(form.train_place.value) == "3" ){
	  if (trimString(form.other_trainplace.value) == "" ){
		  alert("當訓練機構為其他時,則其他訓練機構為必填欄位");
		  form.other_trainplace.focus();
		  return false;
	  }
	}
	
	if (trimString(form.course_name.value) == "" ){
		alert("課程名稱不可空白");
		form.course_name.focus();
		return false;
	}
	
  if (trimString(form.course_hour.value) != "" ){
	  if( isNaN(Math.abs(form.course_hour.value)) ){
		  alert("課程時數不可為文字");
		  form.course_hour.focus();
		  return false;
		}
  } else {
		alert("課程時數不可空白");
		form.course_hour.focus();
		return false;   
	}
	
	if (trimString(form.licno.value) == "" ){
		alert("證書字號不可空白");
		form.licno.focus();
		return false;
	}
   
  var ckDate;
  var sDate = new Date();
  var eDate = new Date();

  ckDate = '' + (parseInt(form.begin_date_y.value)+1911) + '/' + form.begin_date_m.value + '/' + form.begin_date_d.value;
  sDate.setFullYear(parseInt(form.begin_date_y.value)+1911, parseInt(form.begin_date_m.value)-1, parseInt(form.begin_date_d.value));  

  if( fnValidDate(ckDate) != true){
    alert('上課期間-起始日期為無效日期!!');
    form.begin_date_y.focus();
    return false;
  }    

  if (trimString(form.begin_date_y.value) != "" ){
    if(isNaN(Math.abs(form.begin_date_y.value))){
      alert("上課期間-起始日期(年)不可輸入文字");    
      form.begin_date_y.focus();            
      return false;
    }
  } else {
		alert("上課期間-起始日期(年)不可空白");
		form.begin_date_y.focus();
		return false;   
	}

  if (trimString(form.begin_date_m.value) == "" ){
    alert("上課期間-起始日期(月)不可空白");
		form.begin_date_m.focus();
		return false;
	}

	if (trimString(form.begin_date_d.value) == "" ){
		alert("上課期間-起始日期(日)不可空白");
		form.begin_date_d.focus();		
		return false;
	}

  if (trimString(form.end_date_y.value) != "" ){        
    if(isNaN(Math.abs(form.end_date_y.value))){
      alert("上課期間-結束日期(年)不可輸入文字");    
      form.end_date_y.focus();            
      return false;
    }
  } else {
		alert("上課期間-結束日期(年)不可空白");
		form.end_date_y.focus();
		return false;   
	}

  if (trimString(form.end_date_m.value) == "" ){
    alert("上課期間-結束日期(月)不可空白");
		form.end_date_m.focus();
		return false;
	}

	if (trimString(form.end_date_d.value) == "" ){
		alert("上課期間-結束日期(日)不可空白");
		form.end_date_d.focus();		
		return false;
	}

	if(trimString(form.end_date_y.value) != "" && trimString(form.end_date_m.value) != "" && trimString(form.end_date_d.value) != ""){
      ckDate = '' + (parseInt(form.end_date_y.value)+1911) + '/' + form.end_date_m.value + '/' + form.end_date_d.value;
      eDate.setFullYear(parseInt(form.end_date_y.value)+1911,parseInt(form.end_date_m.value)-1,parseInt(form.end_date_d.value)) ;    
    	if( sDate>eDate){
        	alert('上課期間-起迄時間有誤!');
        	form.end_date_d.focus();
        	return false;
    	}
    	if( fnValidDate(ckDate) != true){
        	alert('上課期間-結束日期為無效日期!!');
        	form.end_date_d.focus();
        	return false;
    	}
    }
   return true;
}