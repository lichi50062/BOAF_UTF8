function doSubmit(form,cnd,tbank_no,bank_no){	     	    	    
	    form.action="/pages/AS002W.jsp?act="+cnd+"&test=nothing";		     
	    if(((cnd == "Insert") || (cnd == "Update") || (cnd == "Delete"))
	    && (!checkData(form,cnd))) return;	    	    	    
	    if(cnd == "new" || cnd == "Qry") form.submit();
	    if(cnd == "Edit"){
	       form.action="/pages/AS002W.jsp?act="+cnd+"&TBANK_NO="+tbank_no+"&BANK_NO="+bank_no+"&test=nothing";	    	
	       form.submit();
	    }	
	    if((cnd == "Insert") && AskInsert(form)) form.submit();	    
	    if((cnd == "Update") && AskUpdate(form)){
	    	form.action="/pages/AS002W.jsp?act="+cnd+"&TBANK_NO="+form.TBANK_NO.value+"&test=nothing";	    	
	    	form.submit();	    
	    }	 
	    if((cnd == "Delete") && AskDelete(form)){
	    	form.action="/pages/AS002W.jsp?act="+cnd+"&TBANK_NO="+form.TBANK_NO.value+"&test=nothing";	    	
	        form.submit();	    
	    }	 
}	
	
function getData(form,cnd,item){	     	    
	    if(item == 'bank_type'){
	       form.TBANK_NO.value="";	 
		}			
	    form.action="/pages/AS002W.jsp?act=getData&nowact="+cnd+"&test=nothing";
	    form.submit();	    
}	
	
function checkData(form,cnd) 
{	
	if (trimString(form.BANK_NO.value) =="" ){
		alert("������c�N�X���i�ť�");
		form.BANK_NO.focus();
		return false;
	}else{	   
	    if(form.BANK_TYPE.value == "6" || form.BANK_TYPE.value == "7"){
	       if(form.BANK_NO.value.length != 7){
	          alert("�A(��)�|������c�N�X�A����J7�X");	
	          form.BANK_NO.focus();
	          return false;
	       }
	       if(form.BANK_NO.value.substring(0,3) != form.TBANK_NO.value.substring(0,3)){
	       	  alert("�A(��)�|������c�N�X�e�T�X���P�`���c�e�T�X�ۦP");
	       	  form.BANK_NO.focus();
	       	  return false;
	       }	
	    }	
	    if(((form.BANK_TYPE.value != "6") && (form.BANK_TYPE.value != "7")) && (form.BANK_NO.value.length < 3)){
	       alert("������c�N�X�A�ܤֶ���J3�X");	
	       form.BANK_NO.focus();
		   return false;
	    }	
	    
	}	
	
	if (trimString(form.BANK_NAME.value) =="" ){
		alert("���c�W�٤��i�ť�");
		form.BANK_NAME.focus();
		return false;
	}		
	
	if (trimString(form.BANK_B_NAME.value) =="" ){
		alert("���c²�٤��i�ť�");
		form.BANK_B_NAME.focus();
		return false;
	}
		
	if((form.BANK_TYPE.value == "6" || form.BANK_TYPE.value == "7") && (form.EXCHANGE_NO.value.length != 7)){
	    alert("�A(��)�q�ץN�X�A����J7�X");	
	    form.EXCHANGE_NO.focus();
		return false;
	}		
		
   return true;
}
function sameBank_Name(form){
  form.BANK_B_NAME.value = form.BANK_NAME.value;
}

function changeOption(form){
    var myXML,nodeType,nodeValue, nodeName;

    myXML = document.all("TBankXML").XMLDocument;
    form.TBANK_NO.length = 0;
	nodeType = myXML.getElementsByTagName("bankType");
	nodeValue = myXML.getElementsByTagName("bankValue");
	nodeName = myXML.getElementsByTagName("bankName");
	var oOption;

	for(var i=0;i<nodeType.length ;i++)
	{
  		if ((nodeType.item(i).firstChild.nodeValue == form.BANK_TYPE.value) ||
  		    (form.BANK_TYPE.value == "Z"))  {
  			oOption = document.createElement("OPTION");
			oOption.text=nodeName.item(i).firstChild.nodeValue;
  			oOption.value=nodeValue.item(i).firstChild.nodeValue;
  			form.TBANK_NO.add(oOption);
    	}
    }
}
function setSelect(S1, bankid) {
    if(S1 == null)
    	return;
    for(i=0;i<S1.length;i++) {
      	if(S1.options[i].value==bankid)    	{
        	S1.options[i].selected=true;
        	break;
    	}
    }
}




