//94.03.25 add ���]�K�X��,�Y�Ȱ�='Y',�n���Ѱ��Ȱ��~�୫�]�K�X
//94.04.08 add �A�����t�κ޲z�̫e7�X�����P�`���c�N�X�ۦP
function AskDelete_Permission() {
  if(confirm("�T�w�n�R����(�@�֧R�����b�����v��)�H"))
    return true;
  else
    return false;
}

function doSubmit(form,cnd,muser_id){	     	    	    
	    form.action="/pages/ZZ001W.jsp?act="+cnd+"&test=nothing";	    
	    if(((cnd == "Insert") || (cnd == "Update") || (cnd == "Delete") )
	    && (!checkData(form,cnd))) return;	    	    	    	    
	    
	    if(cnd == "new" || cnd == "Qry") form.submit();
	    if(cnd == "ResetPwd"){
	       //94.03.25 add ���]�K�X��,�Y�Ȱ�='Y',�n���Ѱ��Ȱ��~�୫�]�K�X
	       if(form.LOCK_MARK.value == 'Y'){
	       	  alert('�Х��Ѱ��Ȱ�,�~����歫�]�K�X!!');	       	  
	       	  return;
	       }else if(AskResetPwd(form)){	       	  
	       	  form.submit();	    
	       }       	    
	       return;
	    } 
	    
	    if(cnd == "Edit"){
	       form.action="/pages/ZZ001W.jsp?act="+cnd+"&muser_id="+muser_id+"&test=nothing";	    	
	       form.submit();
	    }	
	    if((cnd == "Insert") && AskInsert(form)) form.submit();	    
	    if((cnd == "Update") && AskUpdate(form)) form.submit();	    
	    if((cnd == "Delete") && AskDelete_Permission(form)) form.submit();	    
	    
}	

function getData(form,cnd,item){	     	    
	    if(item == 'bank_type'){
	       form.TBANK_NO.value="";	 
		}	
		if(item == 'tbank_no'){
	       form.BANK_NO.value="";	 
		}
	    form.action="/pages/ZZ001W.jsp?act=getData&nowact="+cnd+"&test=nothing";
	    form.submit();	    
}	
	
function checkData(form,cnd) 
{	
	if (trimString(form.MUSER_ID.value) =="" ){
		alert("�ϥΪ̱b�����i�ť�");
		form.MUSER_ID.focus();
		return false;
	}else{	   
	    if((form.BANK_TYPE.value == "2") && (form.MUSER_ID.value.length != 10)){
	       alert("�A���������ϥΪ̱b���A����J10�X");	
	       form.MUSER_ID.focus();
		   return false;
	    }	
	    if((form.BANK_TYPE.value != "2") && (form.MUSER_ID.value.length != 3) && (cnd != 'Update' && cnd != 'Delete') ){
	       alert("�A�������~�ϥΪ̱b���A����J���N3�X");	
	       form.MUSER_ID.focus();
		   return false;
	    }	
	    
	}	
	//alert(form.MUSER_ID.value.substring(0,7));
	//94.04.08 add �A�����t�κ޲z�̫e7�X�����P�`���c�N�X�ۦP
	if(form.lguser_id.value == 'A111111111' && 
	   form.lguser_type.value == 'S' && 
	   form.BANK_TYPE.value == '2' && 
	   form.TBANK_NO.value != form.MUSER_ID.value.substring(0,7))
	{
	  	alert("�A�����t�κ޲z�̫e7�X�����P�`���c�N�X�ۦP!!");
	  	return false;	  	
	}  	
	if (trimString(form.MUSER_NAME.value) =="" ){
		alert("�ϥΪ̩m�W���i�ť�");
		form.MUSER_NAME.focus();
		return false;
	}		
	
	
		
   return true;
}

