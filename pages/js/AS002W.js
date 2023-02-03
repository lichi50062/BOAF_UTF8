// 99.04.21 fix 刪除分支機構代號時,一併刪除分支機構基本資料 by 2295
//101.06.29 add 總機構資料,增加第一筆為空白 by 2295    
function doSubmit(form,cnd,tbank_no,bank_no){	     	    	    
	    form.action="/pages/AS002W.jsp?act="+cnd+"&test=nothing";		     
	    if(((cnd == "Insert") || (cnd == "Update"))
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
	    if((cnd == "Delete") && AskDeleteWLX02(form)){
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
		alert("分支機構代碼不可空白");
		form.BANK_NO.focus();
		return false;
	}else{	   
	    if(form.BANK_TYPE.value == "6" || form.BANK_TYPE.value == "7"){
	       if(form.BANK_NO.value.length != 7){
	          alert("農(漁)會分支機構代碼，須填入7碼");	
	          form.BANK_NO.focus();
	          return false;
	       }
	       if(form.BANK_NO.value.substring(0,3) != form.TBANK_NO.value.substring(0,3)){
	       	  alert("農(漁)會分支機構代碼前三碼須與總機構前三碼相同");
	       	  form.BANK_NO.focus();
	       	  return false;
	       }	
	    }	
	    if(((form.BANK_TYPE.value != "6") && (form.BANK_TYPE.value != "7")) && (form.BANK_NO.value.length < 3)){
	       alert("分支機構代碼，至少須填入3碼");	
	       form.BANK_NO.focus();
		   return false;
	    }	
	    
	}	
	
	if (trimString(form.BANK_NAME.value) =="" ){
		alert("機構名稱不可空白");
		form.BANK_NAME.focus();
		return false;
	}		
	
	if (trimString(form.BANK_B_NAME.value) =="" ){
		alert("機構簡稱不可空白");
		form.BANK_B_NAME.focus();
		return false;
	}
		
	if((form.BANK_TYPE.value == "6" || form.BANK_TYPE.value == "7") && (form.EXCHANGE_NO.value.length != 7)){
	    alert("農(漁)通匯代碼，須填入7碼");	
	    form.EXCHANGE_NO.focus();
		return false;
	}		
		
   return true;
}
function sameBank_Name(form){
  form.BANK_B_NAME.value = form.BANK_NAME.value;
}
//101.06.29 add 總機構資料,增加第一筆為空白     
function changeOption(form){
    var myXML,nodeType,nodeValue, nodeName;

    myXML = document.all("TBankXML").XMLDocument;
    form.TBANK_NO.length = 0;
	nodeType = myXML.getElementsByTagName("bankType");
	nodeValue = myXML.getElementsByTagName("bankValue");
	nodeName = myXML.getElementsByTagName("bankName");
	var oOption;
	//101.06.29 add 總機構資料,增加第一筆為空白
  oOption = document.createElement("OPTION");
	oOption.text=" ";
  oOption.value=" ";
  form.TBANK_NO.add(oOption);  
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

function AskDeleteWLX02() {
	
  //if(confirm("確定要刪除嗎？"))
  var myConfirmResult = myConfirm("Microsoft Internet Explorer","確定要刪除嗎(一併刪除此分支機構代號之機構基本資料)？",2,18,9)
  if(myConfirmResult)  
    return true;
  else
    return false;
}



