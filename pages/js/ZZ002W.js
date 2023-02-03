function doSubmit(form,cnd){	     	    	  	     
	    form.action="/pages/ZZ002W.jsp?act="+cnd+"&test=nothing";	    
	    if(((cnd == "Insert") || (cnd == "Delete"))
	    && (!checkData(form))) return;	   	    
	    
	    if(cnd == "new" || cnd == "del" || cnd == "Qry" || cnd == "newQry" || cnd == "delQry") form.submit();	    	
	    if((cnd == "Insert") && AskInsert(form)) form.submit();	    	    	     	 
	    if((cnd == "Delete") && AskDelete(form)) form.submit();	    
}	

function getData(form,cnd){	     	    	    
	    form.action="/pages/ZZ002W.jsp?act=getData&nowact="+cnd+"&test=nothing";
	    form.submit();	    
}	
	
function checkData(form) 
{	
  var flag = false;  
  for (var i = 0 ; i < form.elements.length; i++) {    
    if ( form.elements[i].checked == true ) {
        flag = true;
    }    
  }
  if (flag == false) {         
    alert('請至少選擇一筆資料!');
    return false;
  }
  return true;
}

function selectAll(form) {  
  for ( var i = 0; i < form.elements.length; i++) {
      if (form.elements[i].type=='checkbox') {	
      	form.elements[i].checked = true;
      }	    
  }
  return;
}

function selectNo(form) {  
  for ( var i = 0; i < form.elements.length; i++) {
      if (form.elements[i].type=='checkbox') {	
      	form.elements[i].checked = false;
      }	    
  }
  return;
}