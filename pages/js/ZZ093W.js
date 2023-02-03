function doSubmit(form,cnd){
    if(cnd == "Qry"){
    	form.action="/pages/ZZ093W.jsp?act="+cnd;
    	if(!checkData(form)) return;
    	form.submit();
    }
    if(cnd == "Report"){
    	if(checkData(form)){
    		form.act.value = cnd;
    		form.action="/pages/ZZ093W.jsp?act="+cnd;
           	form.submit();
    	}
    }
    return;
}
//輸入資料檢核處理
function checkData(form){
	form.begDate.value = mergeDate(form.begY.value, form.begM.value, form.begD.value); 
	form.endDate.value = mergeDate(form.endY.value, form.endM.value, form.endD.value);
	
	if(form.begDate.value>form.endDate.value) {
	    alert("使用日期之起始日期不得小於結束日期"); 
	    return false;
	}
	if(isNaN(Math.abs(form.begY.value))) {
        alert("使用日期(起)一定要輸入數字");
        return false;
    }
    if(isNaN(Math.abs(form.endY.value))) {
        alert("使用日期(迄)一定要輸入數字");
        return false;
    }
	
	d = new Date();
	dateString  = mergeDate(d.getFullYear()-1911, d.getMonth()+1, d.getDate());
	return true;
}

function setSelect(S1, date) {
    if(S1 == null)
    	return;
    for(i=0;i<S1.length;i++) {
      	if(S1.options[i].value==date){
        	S1.options[i].selected=true;
        	break;
    	}
    }
}
function addZero(id, num) {
    var temp = "";
    for(var i = 0; i < num; i++) {
        temp += "0";
    }
    temp = temp + id;
    end = temp.length;
    start = end - num;

    return temp.substring(start,end);   
}
function mergeDate(yy, mm, dd) {
    dateY = eval(yy)+1911;
    dateM = addZero(mm, 2);
    dateD = addZero(dd, 2);
    return dateY+dateM+dateD;  
}
