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

function doSubmit(form,cnd,bank_no){	     	  
if(cnd =='add'){
	 if(!checkData(form)) return;	
	if(AskInsert_Permission()){
	form.action="/pages/FX012W.jsp?act=Insert";	    
	form.submit();
  }
}
else if(cnd =='modify'){
	 if(!checkData(form)) return;	
	if(AskEdit_Permission()){
	form.action="/pages/FX012W.jsp?act=Update";	    
	form.submit();
	}
}
else if(cnd =='delete'){
	if(AskDelete_Permission()){
	form.action="/pages/FX012W.jsp?act=Delete";	    
	form.submit();
	}
}
else if(cnd =='Print'){	
	form.action="/pages/FX012W.jsp?act=Print&bank_no="+bank_no;	    
	form.submit();
}

else return;
	  	  	    
}

function checkData(form) 
{
	var ckDate;
	if(trimString(form.m_year.value)  != "" ){        
    	if(isNaN(Math.abs(form.m_year.value))){
           alert("年度不可輸入文字");    
		   form.m_year.focus();            
           return false;
        }
    }else{
		alert("年度不可空白");
		form.m_year.focus();
		return false;   
	}
	
    return true;
}

