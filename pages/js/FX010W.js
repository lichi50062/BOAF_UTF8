var checkLock = new Array()

function pushArray(lock)
{
checkLock.push(lock)
}

function popArray(form)
{
	for(i=0;i<checkLock.length;i+=2)
	{
	 if(form.S_YEAR.value==checkLock[i] && form.S_QUARTER.value==checkLock[i+1])
	 {
	 	alert("你鍵入的申報年季巳被鎖定!");
	 	return false;
	 }
	}
	return true;
}

function AskDelete_Permission() {
  if(confirm("確定要刪除此筆資料嗎？"))
    return true;
  else
    return false;
}

function AskInsert_Permission() {
  if(confirm("確定要新增此筆資料嗎？"))
    return true;
  else
    return false;
}
function AskEdit_Permission() {
  if(confirm("確定要修改此筆資料嗎？"))
    return true;
  else
    return false;
}


function newSubmit(form,cnd,lockdate){	     	  
if(cnd == 'new'){
var checkDate=new Date()
var myDate=new Date()
var checkquarter;
var a = form.S_YEAR.value;
var b = form.S_QUARTER.value;
	if (trimString(a) =="" ){
			alert("年份不可空白");
			form.S_YEAR.focus();
			return;
	}
	if (trimString(b) =="" ){
		alert("季不可空白");
		form.S_QUARTER.focus();
		return;
	}
	
//remeber to cast type 2005.12.23 by lilic0c0 4183
var selectDate = parseInt(a,10)*12+parseInt(b,10); 

//alert(lockdate+" "+selectDate);

if(!popArray(form)) return;	

if(lockdate > selectDate)
{
	
	alert("你鍵入的申報年季超出申報項目起始年季");	
	return;	
}

var year=parseInt(a,10)+1911;
var quarter=parseInt(b,10);
if(checkDate.getMonth()<=3)checkquarter=4;
else if(checkDate.getMonth()<=6)checkquarter=1;
else if(checkDate.getMonth()<=9)checkquarter=2;
else if(checkDate.getMonth()<=12)checkquarter=3;

if(checkDate.getYear()*12+checkquarter < year*12+quarter)
{
		alert("申報年月超過可申報年季!");	
	return;
}

	form.action="/pages/FX010W.jsp?act=new&checkyear="+a+"&checkquarter="+quarter;	    
	form.submit();
	
}
}


function doSubmit(form,cnd){	     	  
if(cnd =='add'){
	if(!checkData(form)) return;	
	if(AskInsert_Permission()){
	form.action="/pages/FX010W.jsp?act=Insert";	    
	form.submit();
  }
}
else if(cnd =='modify'){
	if(!checkData(form)) return;	
	if(AskEdit_Permission()){
	form.action="/pages/FX010W.jsp?act=Update";	    
	form.submit();
	}
}
else if(cnd =='delete'){
	if(AskDelete_Permission()){
	form.action="/pages/FX010W.jsp?act=Delete";	    
	form.submit();
	}
}
else if(cnd =='load'){
	if( confirm("確定載入上季資料？")){
	form.action="/pages/FX010W.jsp?act=Load";	    
	form.submit();
	}
}
else return;
	  	  	    
}

function checkPoint(val) 
{
	 if(parseFloat(val)>99)
	 	return false;
	 else if((val.length - val.indexOf("."))>=5 )
	 	return false;
   	 else
		return true;
	
}

function checkData(form) 
{
if (trimString(form.period_1_fix_rate.value) =="" ){
		alert("定期存款一個月固定不可空白");
		form.period_1_fix_rate.focus();
		return false;
	}
else if (trimString(form.period_1_var_rate.value) =="" ){
		alert("定期存款一個月機動不可空白");
		form.period_1_var_rate.focus();
		return false;
	}
else if (trimString(form.period_3_fix_rate.value) =="" ){
		alert("定期存款三個月固定不可空白");
		form.period_3_fix_rate.focus();
		return false;
	}
else if (trimString(form.period_3_var_rate.value) =="" ){
		alert("定期存款三個月機動不可空白");
		form.period_3_var_rate.focus();
		return false;
	}
else if (trimString(form.period_6_fix_rate.value) =="" ){
		alert("定期存款六個月固定不可空白");
		form.period_6_fix_rate.focus();
		return false;
	}
else if (trimString(form.period_6_var_rate.value) =="" ){
		alert("定期存款六個月機動不可空白");
		form.period_6_var_rate.focus();
		return false;
	}
else if (trimString(form.period_9_fix_rate.value) =="" ){
		alert("定期存款九個月固定不可空白");
		form.period_9_fix_rate.focus();
		return false;
	}
else if (trimString(form.period_9_var_rate.value) =="" ){
		alert("定期存款九個月機動不可空白");
		form.period_9_var_rate.focus();
		return false;
	} 
else if (trimString(form.period_12_fix_rate.value) =="" ){
		alert("定期存款一年固定不可空白");
		form.period_12_fix_rate.focus();
		return false;
	}
else if (trimString(form.period_12_var_rate.value) =="" ){
		alert("定期存款一年機動不可空白");
		form.period_12_var_rate.focus();
		return false;
	}
else if (trimString(form.basic_pay_var_rate.value) =="" ){
		alert("基本放款利率(機動)不可空白");
		form.basic_pay_var_rate.focus();
		return false;
	}
else if (trimString(form.period_house_var_rate.value) =="" ){
		alert("指數型房貸指標利率不可空白");
		form.period_house_var_rate.focus();
		return false;
	}
else if (trimString(form.base_mark_rate.value) =="" ){
		alert("指標利率不可空白");
		form.base_mark_rate.focus();
		return false;
	}
else if (trimString(form.base_fix_rate.value) =="" ){
		alert("一定比率不可空白");
		form.base_fix_rate.focus();
		return false;
	}
else if (trimString(form.base_base_rate.value) =="" ){
		alert("基準利率不可空白");
		form.base_base_rate.focus();
		return false;
	}

else if (!checkPoint(form.period_1_fix_rate.value)){
		alert("定期存款一個月固定格式錯誤");
		form.period_1_fix_rate.focus();
		return false;
	}
else if (!checkPoint(form.period_1_var_rate.value)){
		alert("定期存款一個月機動格式錯誤");
		form.period_1_var_rate.focus();
		return false;
	}
else if (!checkPoint(form.period_3_fix_rate.value)){
		alert("定期存款三個月固定格式錯誤");
		form.period_3_fix_rate.focus();
		return false;
	}
else if (!checkPoint(form.period_3_var_rate.value)){
		alert("定期存款三個月機動格式錯誤");
		form.period_3_var_rate.focus();
		return false;
	}
else if (!checkPoint(form.period_6_fix_rate.value)){
		alert("定期存款六個月固定格式錯誤");
		form.period_6_fix_rate.focus();
		return false;
	}
else if (!checkPoint(form.period_6_var_rate.value)){
		alert("定期存款六個月機動格式錯誤");
		form.period_6_var_rate.focus();
		return false;
	}
else if (!checkPoint(form.period_9_fix_rate.value)){
		alert("定期存款九個月固定格式錯誤");
		form.period_9_fix_rate.focus();
		return false;
	}
else if (!checkPoint(form.period_9_var_rate.value)){
		alert("定期存款九個月機動格式錯誤");
		form.period_9_var_rate.focus();
		return false;
	} 
else if (!checkPoint(form.period_12_fix_rate.value)){
		alert("定期存款一年固定格式錯誤");
		form.period_12_fix_rate.focus();
		return false;
	}
else if (!checkPoint(form.period_12_var_rate.value)){
		alert("定期存款一年機動格式錯誤");
		form.period_12_var_rate.focus();
		return false;
	}
else if (!checkPoint(form.basic_pay_var_rate.value)){
		alert("基本放款利率(機動)格式錯誤");
		form.basic_pay_var_rate.focus();
		return false;
	}
else if (!checkPoint(form.period_house_var_rate.value)){
		alert("指數型房貸指標利率格式錯誤");
		form.period_house_var_rate.focus();
		return false;
	}
else if (!checkPoint(form.base_mark_rate.value)){
		alert("指標利率格式錯誤");
		form.base_mark_rate.focus();
		return false;
	}
else if (!checkPoint(form.base_fix_rate.value)){
		alert("一定比率格式錯誤");
		form.base_fix_rate.focus();
		return false;
	}
else if (!checkPoint(form.base_base_rate.value)){
		alert("基準利率格式錯誤");
		form.base_base_rate.focus();
		return false;
	}

var base_mark = parseFloat(form.base_mark_rate.value);
var base_fix = parseFloat(form.base_fix_rate.value);
var base_base = parseFloat(form.base_base_rate.value);
var temp = parseFloat('0.000001');
if(Math.abs(base_base-(base_fix+base_mark))>temp)
{
	alert("指標利率與一定比率之和不為基準利率");
	return false;
}
    return true;
    
 
}
