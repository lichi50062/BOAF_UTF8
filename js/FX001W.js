//94.04.01 add ���Jwlx04����� by 2295
//94.04.01 add �ˮ֨������� by 2295
//94.04.06 add ����������i�j�������Τp��N����� by 2295
//94.04.07 add �`���c���M������i�p����@������c���M��� by 2295

function doSubmit(form,cnd,table){
	    	
	    if((table == 'Master') && (!checkWLX01Data(form))){	        	      
	       return;
	    }	
	    if((table == 'Manager') && (!checkManager(form))){
	       return;
	    }	    
	    if((table == 'WM') && (!checkWLX01_WMData(form))){
	       return;
	    }
	    if(cnd == 'AbdicateM'){//�����D�ި���
	       form.ABDICATE_CODE[0].selected=true;	   
	       form.ABDICATE_DATE_Y.disabled=false;
	       form.ABDICATE_DATE_M.disabled=false;
	       form.ABDICATE_DATE_D.disabled=false;    
	       if(AskAbdicate()){
	       	  if(!checkManager(form)) return;
	       	  form.action="/pages/FX001W.jsp?act="+cnd+"&position_code="+form.POSITION_CODE.value+"&id="+form.ID.value+"&test=nothing";
	          form.submit();	     
	       }else{
	          return;
	       }	
	    } 
		
	    if(cnd == 'Revoke'){//�`���c���M
	       form.CANCEL_NO[0].selected=true;	   
	       form.CANCEL_DATE_Y.disabled=false;
	       form.CANCEL_DATE_M.disabled=false;
	       form.CANCEL_DATE_D.disabled=false;    
	       if(AskRevoke()){
	       	  if(!checkWLX01Data(form)) return;
	       	  form.action="/pages/FX001W.jsp?act="+cnd+"&test=nothing";
	          form.submit();	     
	       }else{
	          return;
	       }	
	    }
	    
	    if(table == 'Master'){	    
	       if((cnd == "Update") && AskUpdate(form)){ 
	           form.action="/pages/FX001W.jsp?act="+cnd+"&test=nothing";
	           form.submit();	    
	       }
		}	
	    if(table == 'Manager'){	   
	       form.action="/pages/FX001W.jsp?act="+cnd+"&position_code="+form.POSITION_CODE.value+"&id="+form.ID.value+"&test=nothing";
	       if((cnd == "InsertM") && AskInsert(form)) form.submit();	    
	       if((cnd == "UpdateM") && AskUpdate(form)) form.submit();	    
	       if((cnd == "DeleteM") && AskDelete(form)) form.submit();	    	       	       
	    }
	    if(table == 'WM'){	   
	       form.action="/pages/FX001W.jsp?act="+cnd+"&S_YEAR="+form.S_YEAR.value+"&S_MONTH="+form.S_MONTH.value+"&test=nothing";
	       if((cnd == "InsertWM") && AskInsert(form)) form.submit();	    
	       if((cnd == "UpdateWM") && AskUpdate(form)) form.submit();	    
	       if((cnd == "DeleteWM") && AskDelete(form)) form.submit();	    	       	         
	    }
}	

function AddManager(form,WLX01_size){
	   if(WLX01_size == '0'){
	   	  alert("�Х���J�D�ɸ��");
	   	  return;
	   }else{	
	      form.action="/pages/FX001W.jsp?act=newM&test=nothing";
	      form.submit();
	   }   
}	

function AddWM(form,WLX01_size){
	   if(WLX01_size == '0'){
	   	  alert("�Х���J�D�ɸ��");
	   	  return;
	   }else{	
	      form.action="/pages/FX001W.jsp?act=newWM&test=nothing";
	      form.submit();
	   }   
}
//94.04.01 add ���Jwlx04����� by 2295
function loadData(form,cnd,cnd1){	 		
	      form.action="/pages/FX001W.jsp?act="+cnd+"&nowact="+cnd1+"&test=nothing";	      
	      form.submit();	      
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

function setCancelDate(form){	
	    if(form.CANCEL_NO.value == 'Y'){
	       form.CANCEL_DATE_Y.disabled=false;
	       form.CANCEL_DATE_M.disabled=false;
	       form.CANCEL_DATE_D.disabled=false;
	    }else{
	       form.CANCEL_DATE_Y.value="";
	       form.CANCEL_DATE_M.value="";
	       form.CANCEL_DATE_D.value="";
	       form.CANCEL_DATE_Y.disabled=true;
	       form.CANCEL_DATE_M.disabled=true;
	       form.CANCEL_DATE_D.disabled=true;
	    }
}


function setCenterNO(form){	
	    if(form.CENTER_FLAG.value == 'Y'){
	       form.CENTER_NO.disabled=false;	       
	    }else{
	       form.CENTER_NO.value="";
	       form.CENTER_NO.disabled=true;	       
	    }
}

function setAddr(form,cnd){	
	    if(cnd == 'IT_ADDR'){
	    	form.IT_ADDR.value=form.ADDR.value;	    	
		}	
		if(cnd == 'AUDIT_ADDR'){
	    	form.AUDIT_ADDR.value=form.ADDR.value;
		}
}

function checkWLX01Data(form) 
{
	var ckDate;
	
	if (trimString(form.STAFF_NUM.value) != "" ){
        if(isNaN(Math.abs(form.STAFF_NUM.value))){
           alert("�����c���u�`�H�Ƥ��i����r");
           form.STAFF_NUM.focus();
           return false;
        }           		
	}else{
	    alert("�����c���u�`�H�Ƥ��i���ť�");
	    return false;
	}	
	
    
	if((trimString(form.SETUP_DATE_Y.value) != "" ) 
    || (trimString(form.SETUP_DATE_M.value) != "" )
    || (trimString(form.SETUP_DATE_D.value) != "" ))
    {				   	    
        if (trimString(form.SETUP_DATE_Y.value)  != "" ){        
        	if(isNaN(Math.abs(form.SETUP_DATE_Y.value))){
               alert("��l�֭�]�ߤ��(�~)���i��J��r");    
			   form.SETUP_DATE_Y.focus();            
               return false;
            }
        }else{
			alert("��l�֭�]�ߤ��(�~)���i�ť�");
			form.SETUP_DATE_Y.focus();
			return false;   
		}   
        if (trimString(form.SETUP_DATE_M.value) == "" ){
			alert("��l�֭�]�ߤ��(��)���i�ť�");
			form.SETUP_DATE_M.focus();
			return false;
		}			
		if (trimString(form.SETUP_DATE_D.value) == "" ){
			alert("��l�֭�]�ߤ��(��)���i�ť�");
			form.SETUP_DATE_D.focus();		
			return false;
		}	    

    	ckDate = '' + (parseInt(form.SETUP_DATE_Y.value)+1911) + '/' + form.SETUP_DATE_M.value + '/' + form.SETUP_DATE_D.value;
       
    	if( fnValidDate(ckDate) != true){
        	alert('��l�֭�]�ߤ�����L�Ĥ��!!');
        	form.SETUP_DATE_D.focus();
        	return false;
    	}    
    	form.SETUP_DATE.value = ckDate;   	    	
    }
    
    if((trimString(form.CHG_LICENSE_DATE_Y.value) != "" ) 
    || (trimString(form.CHG_LICENSE_DATE_M.value) != "" )
    || (trimString(form.CHG_LICENSE_DATE_D.value) != "" ))
    {				   	    
        if (trimString(form.CHG_LICENSE_DATE_Y.value)  != "" ){        
        	if(isNaN(Math.abs(form.CHG_LICENSE_DATE_Y.value))){
               alert("�̪񴫷Ӥ��(�~)���i��J��r");    
			   form.CHG_LICENSE_DATE_Y.focus();            
               return false;
            }
        }else{
			alert("�̪񴫷Ӥ��(�~)���i�ť�");
			form.CHG_LICENSE_DATE_Y.focus();
			return false;   
		}   
        if (trimString(form.CHG_LICENSE_DATE_M.value) == "" ){
			alert("�̪񴫷Ӥ��(��)���i�ť�");
			form.CHG_LICENSE_DATE_M.focus();
			return false;
		}			
		if (trimString(form.CHG_LICENSE_DATE_D.value) == "" ){
			alert("�̪񴫷Ӥ��(��)���i�ť�");
			form.CHG_LICENSE_DATE_D.focus();		
			return false;
		}	    

    	ckDate = '' + (parseInt(form.CHG_LICENSE_DATE_Y.value)+1911) + '/' + form.CHG_LICENSE_DATE_M.value + '/' + form.CHG_LICENSE_DATE_D.value;
       
    	if( fnValidDate(ckDate) != true){
        	alert('�̪񴫷Ӥ�����L�Ĥ��!!');
        	form.CHG_LICENSE_DATE_D.focus();
        	return false;
    	}    
    	form.CHG_LICENSE_DATE.value = ckDate;   	    	
    }
   
    if((trimString(form.OPEN_DATE_Y.value) != "" ) 
    || (trimString(form.OPEN_DATE_M.value) != "" )
    || (trimString(form.OPEN_DATE_D.value) != "" ))
    {				   	    
        if (trimString(form.OPEN_DATE_Y.value)  != "" ){        
        	if(isNaN(Math.abs(form.OPEN_DATE_Y.value))){
               alert("��l�}�~���(�~)���i��J��r");    
			   form.OPEN_DATEE_Y.focus();            
               return false;
            }
        }else{
			alert("��l�}�~���(�~)���i�ť�");
			form.OPEN_DATE_Y.focus();
			return false;   
		}   
        if (trimString(form.OPEN_DATE_M.value) == "" ){
			alert("��l�}�~���(��)���i�ť�");
			form.OPEN_DATE_M.focus();
			return false;
		}			
		if (trimString(form.OPEN_DATE_D.value) == "" ){
			alert("��l�}�~���(��)���i�ť�");
			form.OPEN_DATE_D.focus();		
			return false;
		}	    

    	ckDate = '' + (parseInt(form.OPEN_DATE_Y.value)+1911) + '/' + form.OPEN_DATE_M.value + '/' + form.OPEN_DATE_D.value;
       
    	if( fnValidDate(ckDate) != true){
        	alert('��l�}�~������L�Ĥ��!!');
        	form.OPEN_DATE_D.focus();
        	return false;
    	}    
    	form.OPEN_DATE.value = ckDate;   	    	
    }
    
    if((trimString(form.START_DATE_Y.value) != "" ) 
    || (trimString(form.START_DATE_M.value) != "" )
    || (trimString(form.START_DATE_D.value) != "" ))
    {				   	    
        if (trimString(form.START_DATE_Y.value)  != "" ){        
        	if(isNaN(Math.abs(form.START_DATE_Y.value))){
               alert("�}�l��~��(�~)���i��J��r");    
			   form.START_DATEE_Y.focus();            
               return false;
            }
        }else{
			alert("�}�l��~��(�~)���i�ť�");
			form.START_DATE_Y.focus();
			return false;   
		}   
        if (trimString(form.START_DATE_M.value) == "" ){
			alert("�}�l��~��(��)���i�ť�");
			form.START_DATE_M.focus();
			return false;
		}			
		if (trimString(form.START_DATE_D.value) == "" ){
			alert("�}�l��~��(��)���i�ť�");
			form.START_DATE_D.focus();		
			return false;
		}	    

    	ckDate = '' + (parseInt(form.START_DATE_Y.value)+1911) + '/' + form.START_DATE_M.value + '/' + form.START_DATE_D.value;
       
    	if( fnValidDate(ckDate) != true){
        	alert('�}�l��~�鬰�L�Ĥ��!!');
        	form.START_DATE_D.focus();
        	return false;
    	}    
    	form.START_DATE.value = ckDate;   	    	
    }
    
    if(form.CANCEL_NO.value == 'Y'){
		if (trimString(form.CANCEL_DATE_Y.value)  != "" ){        
        	if(isNaN(Math.abs(form.CANCEL_DATE_Y.value))){
            	alert("���M�ͮĤ��(�~)���i��J��r");    
            	form.CANCEL_DATE_Y.focus();
            	return false;
        	}	
        }else{
			alert("���M�ͮĤ��(�~)���i�ť�");
			form.CANCEL_DATE_Y.focus();
			return false;   
		}
			
		if (trimString(form.CANCEL_DATE_M.value) == "" ){
			alert("���M�ͮĤ��(��)���i�ť�");
			form.CANCEL_DATE_M.focus();
			return false;
		}			
		if (trimString(form.CANCEL_DATE_D.value) == "" ){
			alert("���M�ͮĤ��(��)���i�ť�");
			form.CANCEL_DATE_D.focus();
			return false;
		}	       
		
		 
        
    	ckDate = '' + (parseInt(form.CANCEL_DATE_Y.value)+1911) + '/' + form.CANCEL_DATE_M.value + '/' + form.CANCEL_DATE_D.value;
       
    	if( fnValidDate(ckDate) != true){
           	alert('���M�ͮĤ�����L�Ĥ��!!');
           	form.CANCEL_DATE_D.focus();
           	return false;
   		}	          		
   		
   		form.CANCEL_DATE.value = ckDate;   	    	    		
   		
   		//94.04.07 add �`���c���M������i�p����@������c���M���
   		var a = form.wlx02date.value.split(',');
        for(var i =0; i < a.length; i ++){	        
        	//alert('a[i]'+a[i]);
        	//alert('form.cancel_date='+form.CANCEL_DATE.value);        	
	        if(form.CANCEL_DATE.value < a[i]){
	           alert('���M������i�p����@������c���M���');	 
	           return false;
	        }	
        }        
   }
   
   if(form.CENTER_FLAG.value == 'Y'){
   	  if (trimString(form.CENTER_NO.value) == "" ){
			alert("�q���@�Τ��ߥN�X���i�ť�");
			form.CENTER_NO.focus();
			return false;
		}
   }	
   
   
   if (trimString(form.M2_NAME.value) == "" ){
	   alert("�a��D�޾����N�����i�ť�");
	   form.M2_NAME.focus();
	   return false;
   }
   	
   
   return true;
}
//94.04.01 add �ˮ֨������� by 2295
function checkManager(form) 
{
	var ckDate;
	
	if (trimString(form.ID.value) =="" ){
		alert("�����Ҧr�����i�ť�");
		form.ID.focus();
		return false;
	}	
	
	if (trimString(form.NAME.value) =="" ){
		alert("�m�W���i�ť�");
		form.NAME.focus();
		return false;
	}
	
	if (trimString(form.RANK.value) != "" ){
        if(isNaN(Math.abs(form.RANK.value))){
           alert("���줣�i����r");
           form.RANK.focus();
           return false;
        }           		
	}
    
    if((trimString(form.BIRTH_DATE_Y.value) != "" ) 
    || (trimString(form.BIRTH_DATE_M.value) != "" )
    || (trimString(form.BIRTH_DATE_D.value) != "" ))
    {				   	    
        if (trimString(form.BIRTH_DATE_Y.value)  != "" ){        
        	if(isNaN(Math.abs(form.BIRTH_DATE_Y.value))){
               alert("�X�ͦ~���(�~)���i��J��r");    
			   form.BIRTH_DATE_Y.focus();            
               return false;
            }
        }else{
			alert("�X�ͦ~���(�~)���i�ť�");
			form.BIRTH_DATE_Y.focus();
			return false;   
		}   
        if (trimString(form.BIRTH_DATE_M.value) == "" ){
			alert("�X�ͦ~���(��)���i�ť�");
			form.BIRTH_DATE_M.focus();
			return false;
		}			
		if (trimString(form.BIRTH_DATE_D.value) == "" ){
			alert("�X�ͦ~���(��)���i�ť�");
			form.BIRTH_DATE_D.focus();		
			return false;
		}	    

    	ckDate = '' + (parseInt(form.BIRTH_DATE_Y.value)+1911) + '/' + form.BIRTH_DATE_M.value + '/' + form.BIRTH_DATE_D.value;
       
    	if( fnValidDate(ckDate) != true){
        	alert('�X�ͦ~��鬰�L�Ĥ��!!');
        	form.BIRTH_DATE_D.focus();
        	return false;
    	}    
    	form.BIRTH_DATE.value = ckDate;   	    	
    }

	if (trimString(form.INDUCT_DATE_Y.value) == "" ){
		alert("�N�����(�~)���i�ť�");
		form.INDUCT_DATE_Y.focus();
		return false;
	} else {
        if(isNaN(Math.abs(form.INDUCT_DATE_Y.value))){
            alert("�N�����(�~)���i��J��r");    
			form.INDUCT_DATE_Y.focus();            
            return false;
        }	
        if (trimString(form.INDUCT_DATE_M.value) == "" ){
			alert("�N�����(��)���i�ť�");
			form.INDUCT_DATE_M.focus();
			return false;
		}			
		if (trimString(form.INDUCT_DATE_D.value) == "" ){
			alert("�N�����(��)���i�ť�");
			form.INDUCT_DATE_D.focus();		
			return false;
		}	    

    	ckDate = '' + (parseInt(form.INDUCT_DATE_Y.value)+1911) + '/' + form.INDUCT_DATE_M.value + '/' + form.INDUCT_DATE_D.value;
       
    	if( fnValidDate(ckDate) != true){
        	 alert('�N��������L�Ĥ��!!');
         	form.INDUCT_DATE_D.focus();
         	return false;
    	}	  
    	form.INDUCT_DATE.value = ckDate;   	    	
    }	
    	
	if(form.ABDICATE_CODE.value == 'Y'){
		if (trimString(form.ABDICATE_DATE_Y.value)  != "" ){        
        	if(isNaN(Math.abs(form.ABDICATE_DATE_Y.value))){
            	alert("�������(�~)���i��J��r");    
            	form.ABDICATE_DATE_Y.focus();
            	return false;
        	}	
        }else{
			alert("�������(�~)���i�ť�");
			form.ABDICATE_DATE_Y.focus();
			return false;   
		}
			
		if (trimString(form.ABDICATE_DATE_M.value) == "" ){
			alert("�������(��)���i�ť�");
			form.ABDICATE_DATE_M.focus();
			return false;
		}			
		if (trimString(form.ABDICATE_DATE_D.value) == "" ){
			alert("�������(��)���i�ť�");
			form.ABDICATE_DATE_D.focus();
			return false;
		}	        
    	ckDate = '' + (parseInt(form.ABDICATE_DATE_Y.value)+1911) + '/' + form.ABDICATE_DATE_M.value + '/' + form.ABDICATE_DATE_D.value;
       
    	if( fnValidDate(ckDate) != true){
           	alert('����������L�Ĥ��!!');
           	form.ABDICATE_DATE_D.focus();
           	return false;
   		}	       
   		
   		form.ABDICATE_DATE.value = ckDate;   	    	    		   		
   		
   		//94.04.06 add ����������i�j�������Τp��N����� by 2295
        if(( form.ABDICATE_DATE.value < form.INDUCT_DATE.value)
        || ( form.ABDICATE_DATE.value > form.nowDay.value))
        {
            alert('����������i�j�������Τp��N�����');
            return false;			
        }
   }
   //94.04.01 add �ˮ֨������� by 2295
   if(form.ID_CODE.value != 'Y' && !check_identity_no(form.ID.value)){
	    alert('���������ˮ֦��~');
	    return false;
   }
  
   return true;
}


function checkWLX01_WMData(form) 
{	
	if (trimString(form.PUSH_DebitCard_CNT.value) != "" ){
        if(isNaN(Math.abs(form.PUSH_DebitCard_CNT.value))){
           alert("�o����ĥd��(�t������c)���i����r");
           form.PUSH_DebitCard_CNT.focus();
           return false;
        }           		
	}
	if (trimString(form.TRAN_DebitCard_CNT.value) != "" ){
        if(isNaN(Math.abs(form.TRAN_DebitCard_CNT.value))){
           alert("�y�q���ĥd��(�t������c)���i����r");
           form.TRAN_DebitCard_CNT.focus();
           return false;
        }           		
	}
	if (trimString(form.ATM_CNT.value) != "" ){
        if(isNaN(Math.abs(form.ATM_CNT.value))){
           alert("ATM�˳]�x��(�t������c)���i����r");
           form.ATM_CNT.focus();
           return false;
        }           		
	}
	if (trimString(form.TRAN_CNT.value) != "" ){
        if(isNaN(Math.abs(form.TRAN_CNT.value))){
           alert("�������(����) (�t������c)���i����r");
           form.TRAN_CNT.focus();
           return false;
        }           		
	}
	if (trimString(form.TRAN_AMT.value) != "" ){
        if(isNaN(Math.abs(form.TRAN_AMT.value))){
           alert("������B(���~�֭p) (�t������c)���i����r");
           form.TRAN_AMT.focus();
           return false;
        }           		
	}
	if (trimString(form.CHECK_DEPOSIT_CNT.value) != "" ){
        if(isNaN(Math.abs(form.CHECK_DEPOSIT_CNT.value))){
           alert("�䲼�s�ڤ�(�t������c)���i����r");
           form.CHECK_DEPOSIT_CNT.focus();
           return false;
        }           		
	}
	
	if (trimString(form.CHECK_DEPOSIT_AMT.value) != "" ){
        if(isNaN(Math.abs(form.CHECK_DEPOSIT_AMT.value))){
           alert("�䲼�s�ھl�B(�t������c)���i����r");
           form.CHECK_DEPOSIT_AMT.focus();
           return false;
        }           		
	}
	
	return true;
}