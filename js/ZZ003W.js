function doSubmit(form,cnd){	     	    	  	     
	    form.action="/pages/ZZ003W.jsp?act="+cnd+"&test=nothing";	    
	    if( cnd == "Insert" && (!checkData_systype(form)) ) return;	   	    
	    if( cnd == "Delete" && (!checkData(form))) return;	   	    
	    if(cnd == "new" || cnd == "del" || cnd == "Qry" || cnd == "DelQry") form.submit();	    	
	    if((cnd == "Insert") && AskInsert(form)) form.submit();	    	    
	    if((cnd == "Delete") && AskDelete(form)) form.submit();	    
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
    alert('�Цܤֿ�ܤ@�����R�������!');
    return false;
  }
  return true;
}

	
function checkData_systype(form) 
{	
  var flag = false;  
  for (var i = 0 ; i < form.elements.length; i++) {    
    if ( form.elements[i].checked == true && form.elements[i].name.substr(0,16) == 'Systype_isModify') {
        flag = true;
    }    
  }
  if (flag == false) {         
    alert('�Цܤֿ�ܤ@���t�����O���!');
    return false;
  }
  if(!checkData_programtype(form)){
  	 return false;
  }
  return true;
}


function checkData_programtype(form) 
{	
  var flag = false;  
  for (var i = 0 ; i < form.elements.length; i++) {    
    if ( form.elements[i].checked == true && form.elements[i].name.substr(0,20) == 'Programtype_isModify') {
        flag = true;
    }    
  }
  
  if (flag == false) {         
    alert('�Цܤֿ�ܤ@���{���k�ݾ��c���!');
    return false;
  }
  return true;
}
function selectAll(form,kind) {  
  for ( var i = 0; i < form.elements.length; i++) {
  	  if((kind == "del") && (form.elements[i].type=='checkbox')) {	
      	form.elements[i].checked = true;
      }
      if((kind == "sys_type") && (form.elements[i].type=='checkbox') && form.elements[i].name.substr(0,16) == 'Systype_isModify') {	
      	form.elements[i].checked = true;
      }	    
      if((kind == "program_type") && (form.elements[i].type=='checkbox') && form.elements[i].name.substr(0,20) == 'Programtype_isModify') {	
      	form.elements[i].checked = true;
      }	  
  }
  return;
}

function selectNo(form,kind) {  
  for ( var i = 0; i < form.elements.length; i++) {
  	   if((kind == "del") && (form.elements[i].type=='checkbox')) {	
      	 form.elements[i].checked = false;
       }
       if((kind == "sys_type") && (form.elements[i].type=='checkbox') && form.elements[i].name.substr(0,16) == 'Systype_isModify' ) {	
      	 form.elements[i].checked = false;
       }	    
       if((kind == "program_type") && (form.elements[i].type=='checkbox') && form.elements[i].name.substr(0,20) == 'Programtype_isModify') {	
      	 form.elements[i].checked = false;
       }	    
  }
  return;
}