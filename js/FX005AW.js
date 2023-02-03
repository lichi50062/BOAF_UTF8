function AskDelete_Permission() {
  if(confirm("�T�w�n�R��������ƶܡH"))
    return true;
  else
    return false;
}

function AskInsert_Permission() {
  if(confirm("�T�w�n�s�W������ƶܡH"))
    return true;
  else
    return false;
}
function AskEdit_Permission() {
  if(confirm("�T�w�n�ק惡����ƶܡH"))
    return true;
  else
    return false;
}



function doSubmit(form,cnd,bankno,seq_no){	     	  
if(cnd =='add'){
	if(!checkData(form)) return;	
	
	if(!(form.cancel_type.value==1) || !(form.cancel_type.value==2) )
	{		
			form.action="/pages/FX005AW.jsp?act=Add_Check_Property_No&property_no="+form.property_no.value+"&bank_no="+bankno+"&cancel_type="+form.cancel_type.value+"&site_name="+form.site_name.value+"&addr="+form.addr.value+"&setup_date_y="+form.SETUP_DATE_Y.value+"&setup_date_m="+form.SETUP_DATE_M.value+"&setup_date_d="+form.SETUP_DATE_D.value+"&machine_name="+form.machine_name.value;	    				    			
			form.submit();
  }else{ 
  if(AskInsert_Permission()){
  		form.action="/pages/FX005AW.jsp?act=Insert";	        
	    form.submit();
  }
  }
}
else if(cnd =='modify'){
	if(!checkData(form)) return;	
  
  if(!(form.cancel_type.value==1) || !(form.cancel_type.value==2) )
		{				
			form.action="/pages/FX005AW.jsp?act=Modify_Check_Property_No&property_no="+form.property_no.value+"&bank_no="+bankno+"&cancel_type="+form.cancel_type.value+"&site_name="+form.site_name.value+"&addr="+form.addr.value+"&setup_date_y="+form.SETUP_DATE_Y.value+"&setup_date_m="+form.SETUP_DATE_M.value+"&setup_date_d="+form.SETUP_DATE_D.value+"&machine_name="+form.machine_name.value+"&seq_no="+seq_no;	    				    			
			form.submit();
  	}else 
  if(AskEdit_Permission()){
	form.action="/pages/FX005AW.jsp?act=Update";	    
	form.submit();
	}
}
else if(cnd =='delete'){
	if(AskDelete_Permission()){
	form.action="/pages/FX005AW.jsp?act=Delete";	    
	form.submit();
	}
}
else if(cnd =='load'){
	//form.action="/pages/FX007W.jsp?act=Load";	    
	//form.submit();
window.open("/pages/FX005AW.jsp?act=Load","LOAD","width=750,height=400,scrollbars =yes ","")
}
else if(cnd =='checkload'){
var loadseq;
if(form.CHECKSEQ == null)
{
		window.close();
		return;  
}
if(form.CHECKSEQ.length>0){
	for (i=0;i<form.CHECKSEQ.length;i++)
	{
		if (form.CHECKSEQ[i].checked)
		{
			loadseq = form.CHECKSEQ[i].value
			window.opener.location="/pages/FX005AW.jsp?act=Load&CHECKSEQ="+loadseq;	
			window.close();
		}
	}
}
else{
	if (form.CHECKSEQ.checked)
		{
			loadseq = form.CHECKSEQ.value
			window.opener.location="/pages/FX005AW.jsp?act=Load&CHECKSEQ="+loadseq;	
			window.close();
		}
}
	window.close();
}
else return;
	  	  	    
}



function checkData(form) 
{
  var ckDate;
	if(trimString(form.site_name.value) == "" ){
           alert("�˳]�a�I�W�٤��i���ť�");
           form.site_name.focus();
           return false;
    } 
    else if(trimString(form.addr.value) == "" ){
           alert("�a�}���i���ť�");
           form.addr.focus();
           return false;
    }  
	   	    
        if (trimString(form.SETUP_DATE_Y.value)  != "" ){        
        	if(isNaN(Math.abs(form.SETUP_DATE_Y.value))){
               alert("�˳]���(�~)���i��J��r");    
			   form.SETUP_DATE_Y.focus();            
               return false;
            }
        }else{
			alert("�˳]���(�~)���i�ť�");
			form.SETUP_DATE_Y.focus();
			return false;   
		}   
        if (trimString(form.SETUP_DATE_M.value) == "" ){
			alert("�˳]���(��)���i�ť�");
			form.SETUP_DATE_M.focus();
			return false;
		}			
		if (trimString(form.SETUP_DATE_D.value) == "" ){
			alert("�˳]���(��)���i�ť�");
			form.SETUP_DATE_D.focus();		
			return false;
		}	
    
    if (trimString(form.machine_name.value) == ""){
    	alert("�����~�W���i�ť�");
			form.machine_name.focus();		
			return false;
    }
   
    if (trimString(form.property_no.value) == ""){
    	alert("�����s�����i�ť�");
			form.property_no.focus();		
			return false;
    } 
     
    

    ckDate = '' + (parseInt(form.SETUP_DATE_Y.value)+1911) + '/' + form.SETUP_DATE_M.value + '/' + form.SETUP_DATE_D.value;
       
    	if( fnValidDate(ckDate) != true){
        	alert('�˳]������L�Ĥ��!!');
        	form.SETUP_DATE_D.focus();
        	return false;
    	} 
    	   
    	form.SETUP_DATE.value = ckDate;   	
    	    	
 var sDate=new Date()   
 var eDate=new Date()
   
sDate.setFullYear(parseInt(form.SETUP_DATE_Y.value)+1911,parseInt(form.SETUP_DATE_M.value)-1,parseInt(form.SETUP_DATE_D.value));   

if(form.cancel_type.value == '1' || form.cancel_type.value == '2'){
		if (trimString(form.CANCEL_DATE_Y.value)  != "" ){        
        	if(isNaN(Math.abs(form.CANCEL_DATE_Y.value))){
            	alert("�E��/�������(�~)���i��J��r");    
            	form.CANCEL_DATE_Y.focus();
            	return false;
        	}	
        }else{
			alert("�E��/�������(�~)���i�ť�");
			form.CANCEL_DATE_Y.focus();
			return false;   
		}
			
		if (trimString(form.CANCEL_DATE_M.value) == "" ){
			alert("�E��/�������(��)���i�ť�");
			form.CANCEL_DATE_M.focus();
			return false;
		}			
		if (trimString(form.CANCEL_DATE_D.value) == "" ){
			alert("�E��/�������(��)���i�ť�");
			form.CANCEL_DATE_D.focus();
			return false;
		}	       
		
		 
        
    	ckDate = '' + (parseInt(form.CANCEL_DATE_Y.value)+1911) + '/' + form.CANCEL_DATE_M.value + '/' + form.CANCEL_DATE_D.value;
       
    	if( fnValidDate(ckDate) != true){
           	alert('�E��/����������L�Ĥ��!!');
           	form.CANCEL_DATE_D.focus();
           	return false;
   		}	          		
   		
   		form.CANCEL_DATE.value = ckDate;   	    	    
   		eDate.setFullYear(parseInt(form.CANCEL_DATE_Y.value)+1911,parseInt(form.CANCEL_DATE_M.value)-1,parseInt(form.CANCEL_DATE_D.value));   
		if( sDate>eDate){
        	alert('�E��/���� ������~!');
        	form.CANCEL_DATE_Y.focus();
        	return false;
    	}  
}


    return true;
}

