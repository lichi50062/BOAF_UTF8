function doSubmit(form,cnd,program_id){	     	    	  	     
	    form.action="/pages/ZZ005W.jsp?act="+cnd+"&test=nothing";	    
	    if( cnd == "Insert" && (!checkData(form)) ) return;	   	    
	    if( cnd == "Delete" && (!CheckData_Delete(form))) return;	   	    
	    if( cnd == "Edit"){
	    	form.action="/pages/ZZ005W.jsp?act="+cnd+"&program_id="+program_id+"&test=nothing";	    
	    	form.submit();
	    }
	    if(cnd == "new" || cnd == "del" || cnd == "Qry") form.submit();	    	
	    if((cnd == "Insert") && AskInsert(form)) form.submit();	    	    
	    if((cnd == "Update") && AskUpdate(form)) form.submit();	   
	    if((cnd == "Delete") && AskDelete(form)) form.submit();	    
}	

function checkData(form) 
{	
	if (trimString(form.PROGRAM_ID.value) =="" ){
		alert("�{���N�X���i�ť�");
		form.PROGRAM_ID.focus();
		return false;
	}
	
	if (trimString(form.PROGRAM_NAME.value) =="" ){
		alert("�{���W�٤��i�ť�");
		form.PROGRAM_NAME.focus();
		return false;
	}		
	
	if (trimString(form.URL_ID.value) =="" ){
		alert("�{��URL���i�ť�");
		form.URL_ID.focus();
		return false;
	}
	
	if (trimString(form.INPUT_ORDER.value) =="" ){
		alert("��ƶ��Ǥ��i�ť�");
		form.INPUT_ORDER.focus();
		return false;
	}else{
	    if(isNaN(Math.abs(form.INPUT_ORDER.value))){
            alert("��ƶ��Ǥ��i��J��r");    
            form.INPUT_ORDER.focus();
            return false;
        }	
    }
		
		
   return true;
}


function CheckData_Delete(form) 
{
  var flag = false;  
  for (var i = 0 ; i < form.elements.length; i++) {    
    if ( form.elements[i].checked == true ) {
        flag = true;
    }    
  }
  if (flag == false) {         
    alert('�Цܤֿ�ܤ@�����R�������!');
    return false;
  }
  return true;
}


function selectAll(form) {  
  for ( var i = 0; i < form.elements.length; i++) {
      if(form.elements[i].type=='checkbox') {	
      	form.elements[i].checked = true;
      }	    
       
  }
  return;
}

function selectNo(form) {  
  for ( var i = 0; i < form.elements.length; i++) {
       if(form.elements[i].type=='checkbox') {	
      	 form.elements[i].checked = false;
       }	    
           
  }
  return;
}