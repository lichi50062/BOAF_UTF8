var checkData1; //交易次數累計
var checkData2; //交易金額本年累計
var checkLock = new Array()
var loadData;	//是否為首次新增
function pushArray(lock)
{
checkLock.push(lock)
}

function popArray(form)
{
	for(i=0;i<checkLock.length;i+=2)
	{
			 
	 if(form.S_YEAR.value==checkLock[i] && form.S_MONTH.value==checkLock[i+1])
	 {
	 	alert("你鍵入的申報年月巳被鎖定!");
	 	return false;
	 }
	 
	}
	return true;
}

function pushData(v1,v2,v3)
{
	checkData1=v1;
	checkData2=v2;
	loadData=v3;
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

var a = form.S_YEAR.value;
var b = form.S_MONTH.value;
	if (trimString(a) =="" ){
			alert("年份不可空白");
			form.S_YEAR.focus();
			return;
	}
	if (trimString(b) =="" ){
		alert("月份不可空白");
		form.S_MONTH.focus();
		return;
	}

//remeber to cast type 2005.12.23 by lilic0c0 4183
var selectDate = parseInt(a,10)*12+parseInt(b,10); 

if(!popArray(form)) return;	

if(lockdate > selectDate)
{
	alert("你鍵入的申報年月超出申報項目起始年月");	
	return;	
}

var year=parseInt(a,10)+1911;
var month=b-1;
	checkDate.setFullYear(checkDate.getYear(),checkDate.getMonth()-1,31)    
	myDate = new Date()
	myDate.setFullYear(year,month,myDate.getDate())

	if(myDate>checkDate)
	{
		alert("申報年月超過可申報年月!");	
	return;
	}
  
	form.action="/pages/FX005W.jsp?act=new&checkyear="+a+"&checkmonth="+b;	    
	form.submit();
	
}
}


function doSubmit(form,cnd){
		     	  
if(cnd =='add'){
	if(!checkData(form)) return;	
	if(!submitform(form)) return;
	if(true){	
	if(!checkData(form)) return;	
	var barcard = form.debitcard.value;
  var cancel_barcard = form.cancdebitcard.value;
  var count_barcard = barcard-cancel_barcard;
  var bincard = form.bincard.value;
  var cancbincard = form.cancbincard.value;
  var count_bincard = bincard-cancbincard;
  if(count_barcard!=form.usedebitcard.value){
		alert("流通張數應是發行張數減停卡張數");
		form.usedebitcard.focus();
		return;
	}
	if(count_bincard!=form.usebincard.value){
		alert("流通張數應是發行張數減停卡張數");
		form.usebincard.focus();
		return;
	}
	
	if(!submitform(form)) return;
	if(AskInsert_Permission()){
	
	form.action="/pages/FX005W.jsp?act=Insert";	    
	form.submit();
  }
}
}
else if(cnd =='modify'){
	if(!checkData(form)) return;	
	if(!submitform(form)) return;
	if(true){
		
	var barcard = form.debitcard.value;
  var cancel_barcard = form.cancdebitcard.value;
  var count_barcard = barcard-cancel_barcard;
  var bincard = form.bincard.value;
  var cancbincard = form.cancbincard.value;
  var count_bincard = bincard-cancbincard;
  if(count_barcard!=form.usedebitcard.value){
		alert("流通張數應是發行張數減停卡張數");
		form.usedebitcard.focus();
		return;
	}
	if(count_bincard!=form.usebincard.value){
		alert("流通張數應是發行張數減停卡張數");
		form.usebincard.focus();
		return;
	}	
	AskEdit_Permission();
	form.action="/pages/FX005W.jsp?act=Update";	    
	form.submit();
	}
}
else if(cnd =='delete'){
	if(AskDelete_Permission()){
	form.action="/pages/FX005W.jsp?act=Delete";	    
	form.submit();
	}
}
else if(cnd =='load'){
	if( confirm("確定載入上月資料？")){
	form.action="/pages/FX005W.jsp?act=Load";	    
	form.submit();
	}
}
else if(cnd =='returnList'){
if( confirm("確定回查詢頁？")){
	form.action="/pages/FX005W.jsp?act=List";	    
	form.submit();
	}
}
else return;
	  	  	    
}


function checkData(form) 
{
if (trimString(form.debitcard.value) =="" ){
		alert("磁條金融卡發行張數不可空白");
		form.debitcard.focus();		
		return false;
	}
else if (form.debitcard.value < 0 ){
		alert("磁條金融卡發行張數不可為負數");
		form.debitcard.focus();
		return false;
	}	
	else if (trimString(form.cancdebitcard.value) =="" ){
		alert("磁條金融卡停卡張數不可空白");
		form.cancdebitcard.focus();
		return false;
	}
else if (form.cancdebitcard.value < 0 ){
		alert("磁條金融卡停卡張數不可為負數");
		form.cancdebitcard.focus();
		return false;
	}		
	else if (trimString(form.usedebitcard.value) =="" ){
		alert("磁條金融卡流通張數不可空白");
		form.usedebitcard.focus();
		return false;
	}
	else if (form.usedebitcard.value < 0 ){
		alert("磁條金融卡流通張數不可為負數");
		form.usedebitcard.focus();
		return false;
	}		
	else if (trimString(form.bincard.value) =="" ){
		alert("晶片卡發行張數不可空白");
		form.bincard.focus();
		return false;
	}
	else if (form.bincard.value < 0 ){
		alert("晶片卡發行張數不可為負數");
		form.bincard.focus();
		return false;
	}		
	else if (trimString(form.cancbincard.value) =="" ){
		alert("晶片卡停卡張數不可空白");
		form.cancbincard.focus();
		return false;
	}
	else if (form.cancbincard.value < 0 ){
		alert("晶片卡停卡張數不可為負數");
		form.cancbincard.focus();
		return false;
	}		
	else if (trimString(form.usebincard.value) =="" ){
		alert("晶片卡流通張數不可空白");
		form.usebincard.focus();
		return false;
	}
	else if (form.usebincard.value < 0 ){
		alert("晶片卡流通張數不可為負數");
		form.usebincard.focus();
		return false;
	}	
	else if (trimString(form.setup_atm.value) =="" ){
		alert("ATM 裝設台數不可空白");
		form.setup_atm.focus();
		return false;
	}
	else if (form.setup_atm.value < 0 ){
		alert("ATM 裝設台數不可為負數");
		form.setup_atm.focus();
		return false;
	}	
	else	if (trimString(form.monthtran_cnt.value) =="" ){
		alert("本月交易次數不可空白");
		form.monthtran_cnt.focus();
		return false;
	}
	else if (form.monthtran_cnt.value < 0 ){
		alert("本月交易次數不可為負數");
		form.monthtran_cnt.focus();
		return false;
	}
	else	if (trimString(form.yeartran_cnt.value) =="" ){
		alert("本年累計交易次數不可空白");
		form.yeartran_cnt.focus();
		return false;
	}
	else if (form.yeartran_cnt.value < 0 ){
		alert("本年累計交易次數不可為負數");
		form.yeartran_cnt.focus();
		return false;
	}
	else if (trimString(form.monthtran_amt.value) =="" ){
		alert("本月交易金額不可空白");
		form.monthtran_amt.focus();
		return false;
	}
	else if (form.monthtran_amt.value < 0 ){
		alert("本月交易金額不可為負數");
		form.monthtran_amt.focus();
		return false;
	}
	else if (trimString(form.yeartran_amt.value)=="" ){
		alert("本年累計交易金額不可空白");
		form.yeartran_amt.focus();
		return false;
	}
	else if (form.yeartran_amt.value < 0 ){
		alert("本年累計交易金額不可為負數");
		form.yeartran_amt.focus();
		return false;
	}

    return true;
}

function submitform(form)
{
 	form.debitcard.value = changeVal(form.debitcard)
 	form.usedebitcard.value = changeVal(form.usedebitcard)
 	form.cancdebitcard.value = changeVal(form.cancdebitcard)
 	form.bincard.value = changeVal(form.bincard)
 	form.usebincard.value = changeVal(form.usebincard)
 	form.cancbincard.value = changeVal(form.cancbincard) 
 	form.setup_atm.value = changeVal(form.setup_atm) 
 	form.monthtran_cnt.value = changeVal(form.monthtran_cnt)
 	form.monthtran_amt.value = changeVal(form.monthtran_amt)
 	form.yeartran_cnt.value = changeVal(form.yeartran_cnt)
 	form.yeartran_amt.value = changeVal(form.yeartran_amt)


    if( parseInt(form.monthtran_amt.value,10)>parseInt(form.yeartran_amt.value,10))
	{
	    if(!confirm("本年累計交易金額小於本月交易金額，是否確定？"))
			    return false;
	}
	
	if(parseInt(form.monthtran_cnt.value,10) > parseInt(form.yeartran_cnt.value,10))
	{
	    if(!confirm("本年累計交易次數小於本月交易次數，是否確定？"))
	        return false;
	}
	
var mcnt = parseInt(form.monthtran_cnt.value,10);
var mamt= parseInt(form.monthtran_amt.value,10);
var ycnt = parseInt(form.yeartran_cnt.value,10);
var yamt= parseInt(form.yeartran_amt.value,10);

	if(form.hmonth.value=="1"){
		
		if (parseInt(checkData1)+mcnt!=ycnt){
			alert("元月份，「本年累計交易次數」不等於「本月交易次數」");
			form.monthtran_cnt.focus();
			return false;
		}
		if (parseInt(checkData2)+mamt!=yamt){
			alert("元月份，「本年累計交易金額」不等於「本月交易金額」");
			form.monthtran_amt.focus();
			return false;
		}

	} else if( trimString(loadData) != "false"){
		
		if (parseInt(checkData1)+mcnt!=ycnt){
			alert("「本月交易次數」加上「上一個月的本年累計交易次數」\n不等於「本月的本年累計交易次數」");
			form.monthtran_cnt.focus();
			return false;
		}
		if (parseInt(checkData2)+mamt!=yamt){
			alert("「本月交易金額」加上「上一個月的本年累計交易金額」\n不等於「本月的本年累計交易金額」");
			form.monthtran_amt.focus();
			return false;
		}
	}
	
	return true;
  
}