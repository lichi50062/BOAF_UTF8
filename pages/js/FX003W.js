//94.04.01 add 檢核身份証號 by 2295
//94.04.06 add 卸任日期不可大於當日日期或小於就任日期 by 2295
function doSubmit(form,cnd){
	    if(!checkData(form)) return;	    
	    if(cnd == 'Abdicate'){
	       form.ABDICATE_CODE[0].selected=true;	   
	       form.ABDICATE_DATE_Y.disabled=false;
	       form.ABDICATE_DATE_M.disabled=false;
	       form.ABDICATE_DATE_D.disabled=false;    
	       if(AskAbdicate()){
	       	  if(!checkData(form)) return;
	       	  form.action="/pages/FX003W.jsp?act="+cnd+"&position_code="+form.POSITION_CODE.value+"&id="+form.ID.value+"&test=nothing";
	          form.submit();	     
	       }	
	    }else{		    
	       form.action="/pages/FX003W.jsp?act="+cnd+"&position_code="+form.POSITION_CODE.value+"&id="+form.ID.value+"&test=nothing";
	       if((cnd == "Insert") && AskInsert(form)) form.submit();	    
	       if((cnd == "Update") && AskUpdate(form)) form.submit();	    
	       if((cnd == "Delete") && AskDelete(form)) form.submit();	 	       
	    }
}	

function setAbdicateDate(form){	
	    if(form.ABDICATE_CODE.value == 'Y'){
	       form.ABDICATE_DATE_Y.disabled=false;
	       form.ABDICATE_DATE_M.disabled=false;
	       form.ABDICATE_DATE_D.disabled=false;
	    }else{
	       form.ABDICATE_DATE_Y.value="";
	       form.ABDICATE_DATE_M.value="";
	       form.ABDICATE_DATE_D.value="";
	       form.ABDICATE_DATE_Y.disabled=true;
	       form.ABDICATE_DATE_M.disabled=true;
	       form.ABDICATE_DATE_D.disabled=true;
	    }
}	
function checkData(form) 
{
	var ckDate;
	
	if (trimString(form.ID.value) =="" ){
		alert("身分證字號不可空白");
		form.ID.focus();
		return false;
	}	
	
	if (trimString(form.RANK.value) != "" ){
        if(isNaN(Math.abs(form.RANK.value))){
           alert("順位不可為文字");
           form.RANK.focus();
           return false;
        }           		
	}
	
	if (trimString(form.APPOINTED_NUM.value) !="" ){		
        if(isNaN(Math.abs(form.APPOINTED_NUM.value))){
            alert("本屆屆期不可輸入文字");    
            form.APPOINTED_NUM.focus();
            return false;
        }	
    }
    
   if((trimString(form.BIRTH_DATE_Y.value) != "" ) 
   || (trimString(form.BIRTH_DATE_M.value) != "" )
   || (trimString(form.BIRTH_DATE_D.value) != "" ))
   {				
   	    
        if (trimString(form.BIRTH_DATE_Y.value)  != "" ){        
        	if(isNaN(Math.abs(form.BIRTH_DATE_Y.value))){
               alert("出生年月日(年)不可輸入文字");    
			   form.BIRTH_DATE_Y.focus();            
               return false;
            }
        }else{
			alert("出生年月日(年)不可空白");
			form.BIRTH_DATE_Y.focus();
			return false;   
		}       
            
        if (trimString(form.BIRTH_DATE_M.value) == "" ){
			alert("出生年月日(月)不可空白");
			form.BIRTH_DATE_M.focus();
			return false;
		}			
		if (trimString(form.BIRTH_DATE_D.value) == "" ){
			alert("出生年月日(日)不可空白");
			form.BIRTH_DATE_D.focus();		
			return false;
		}	    

    	ckDate = '' + (parseInt(form.BIRTH_DATE_Y.value)+1911) + '/' + form.BIRTH_DATE_M.value + '/' + form.BIRTH_DATE_D.value;
        
    	if( fnValidDate(ckDate) != true){
        	alert('出生年月日為無效日期!!');
        	form.BIRTH_DATE_D.focus();
        	return false;
    	}    
    	form.BIRTH_DATE.value = ckDate;   	    	
    }

	if (trimString(form.INDUCT_DATE_Y.value) == "" ){
		alert("原始就任日期(年)不可空白");
		form.INDUCT_DATE_Y.focus();
		return false;
	} else {
        if(isNaN(Math.abs(form.INDUCT_DATE_Y.value))){
            alert("原始就任日期(年)不可輸入文字");    
			form.INDUCT_DATE_Y.focus();            
            return false;
        }	
        if (trimString(form.INDUCT_DATE_M.value) == "" ){
			alert("原始就任日期(月)不可空白");
			form.INDUCT_DATE_M.focus();
			return false;
		}			
		if (trimString(form.INDUCT_DATE_D.value) == "" ){
			alert("原始就任日期(日)不可空白");
			form.INDUCT_DATE_D.focus();		
			return false;
		}	    

    	ckDate = '' + (parseInt(form.INDUCT_DATE_Y.value)+1911) + '/' + form.INDUCT_DATE_M.value + '/' + form.INDUCT_DATE_D.value;
       
    	if( fnValidDate(ckDate) != true){
        	 alert('原始就任日期為無效日期!!');
         	form.INDUCT_DATE_D.focus();
         	return false;
    	}	  
    	form.INDUCT_DATE.value = ckDate;   	    	
    }	
    
   if((trimString(form.PERIOD_START_Y.value) != "" ) 
   || (trimString(form.PERIOD_START_M.value) != "" )
   || (trimString(form.PERIOD_START_D.value) != "" ))
   {		
        if (trimString(form.PERIOD_START_Y.value)  != "" ){        
            if(isNaN(Math.abs(form.PERIOD_START_Y.value))){
               alert("本屆任期(年)不可輸入文字");    
			   form.PERIOD_START_Y.focus();            
               return false;
            }
        }else{
			alert("本屆任期(年)不可空白");
			form.PERIOD_START_Y.focus();
			return false;   
		}   
        if (trimString(form.PERIOD_START_M.value) == "" ){
		    alert("本屆任期(月)不可空白");
		    form.PERIOD_START_M.focus();
		    return false;
	    }
	    
	    if (trimString(form.PERIOD_START_D.value) == "" ){
			alert("本屆任期(日)不可空白");
			form.PERIOD_START_D.focus();		
			return false;
		}	    

    	ckDate = '' + (parseInt(form.PERIOD_START_Y.value)+1911) + '/' + form.PERIOD_START_M.value + '/' + form.PERIOD_START_D.value;
       
	    if( fnValidDate(ckDate) != true){
            alert('本屆任期為無效日期!!');
            form.PERIOD_START_D.focus();
            return false;
   		}   		
   		form.PERIOD_START.value = ckDate;   	    	
    }
				
	
   if((trimString(form.PERIOD_END_Y.value) != "" ) 
   || (trimString(form.PERIOD_END_M.value) != "" )
   || (trimString(form.PERIOD_END_D.value) != "" ))
   {		
        if (trimString(form.PERIOD_END_Y.value)  != "" ){        
            if(isNaN(Math.abs(form.PERIOD_END_Y.value))){
               alert("本屆任期(年)不可輸入文字");    
			   form.PERIOD_END_Y.focus();            
               return false;
            }
        }else{
			alert("本屆任期(年)不可空白");
			form.PERIOD_END_Y.focus();
			return false;   
		}    
        if (trimString(form.PERIOD_END_M.value) == "" ){
			alert("本屆任期(月)不可空白");
			form.PERIOD_END_M.focus();
			return false;
		}			
		if (trimString(form.PERIOD_END_D.value) == "" ){
			alert("本屆任期(日)不可空白");
			form.PERIOD_END_D.focus();		
			return false;
		}	    

    	ckDate = '' + (parseInt(form.PERIOD_END_Y.value)+1911) + '/' + form.PERIOD_END_M.value + '/' + form.PERIOD_END_D.value;
       
    	if( fnValidDate(ckDate) != true){
        	alert('本屆任期為無效日期!!');
        	form.PERIOD_END_D.focus();
        	return false;
    	}       	    
    	form.PERIOD_END.value = ckDate;   	    		
    }
	
	if((form.PERIOD_START_Y.value != "")	&& 	(form.PERIOD_START_M.value != "") && (form.PERIOD_START_D.value != "")
	  &&(form.PERIOD_END_Y.value != "")	&& 	(form.PERIOD_END_M.value != "") && (form.PERIOD_END_D.value != "")){
    	if ( (Math.abs(form.PERIOD_START_Y.value) * 1000 + Math.abs(form.PERIOD_START_M.value) * 10  + Math.abs(form.PERIOD_START_D.value)) >
    		 (Math.abs(form.PERIOD_END_Y.value) * 1000 + Math.abs(form.PERIOD_END_M.value) * 10  + Math.abs(form.PERIOD_START_D.value)) )
    	{
    		alert("本屆任期起始日期不可以大於終止日期");
     		form.PERIOD_START_Y.focus();   	
    		return false;
   		}
    }
    	
	if(form.ABDICATE_CODE.value == 'Y'){
		if (trimString(form.ABDICATE_DATE_Y.value)  != "" ){        
        	if(isNaN(Math.abs(form.ABDICATE_DATE_Y.value))){
            	alert("卸任日期(年)不可輸入文字");    
            	form.ABDICATE_DATE_Y.focus();
            	return false;
        	}	
        }else{
			alert("卸任日期(年)不可空白");
			form.ABDICATE_DATE_Y.focus();
			return false;   
		}
			
		if (trimString(form.ABDICATE_DATE_M.value) == "" ){
			alert("卸任日期(月)不可空白");
			form.ABDICATE_DATE_M.focus();
			return false;
		}			
		if (trimString(form.ABDICATE_DATE_D.value) == "" ){
			alert("卸任日期(日)不可空白");
			form.ABDICATE_DATE_D.focus();
			return false;
		}	        
    	ckDate = '' + (parseInt(form.ABDICATE_DATE_Y.value)+1911) + '/' + form.ABDICATE_DATE_M.value + '/' + form.ABDICATE_DATE_D.value;
       
    	if( fnValidDate(ckDate) != true){
           	alert('卸任日期為無效日期!!');
           	form.ABDICATE_DATE_D.focus();
           	return false;
   		}	       
   		
   		form.ABDICATE_DATE.value = ckDate;   	    	    		
   		
   		//94.04.06 add 卸任日期不可大於當日日期或小於就任日期 by 2295
        if(( form.ABDICATE_DATE.value < form.INDUCT_DATE.value)
        || ( form.ABDICATE_DATE.value > form.nowDay.value))
        {
          alert('卸任日期不可大於當日日期或小於就任日期');
          return false;			
        }
   }
   //94.04.01 add 檢核身份証號 by 2295
   if(form.ID_CODE.value != 'Y' && !check_identity_no(form.ID.value)){
	    alert('身份証號檢核有誤');
	    return false;
   }
   
   return true;
}