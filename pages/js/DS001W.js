//95.09.04 add 農漁會.預設全部放在已選欄位 by 2295
//         add 農/漁會.全部 by 2295
function changeOption(form,cnd){	
    var myXML,nodeType,nodeType1,nodeValue, nodeName,bankType;
    myXML = document.all("TBankXML").XMLDocument;
    form.BankListSrc.length = 0;
    if(cnd == 'change') form.BankListDst.length = 0;
    BnType = myXML.getElementsByTagName("BnType");    
    bankType = myXML.getElementsByTagName("BankType");    
	nodeType = myXML.getElementsByTagName("HsienId");
	nodeValue = myXML.getElementsByTagName("bankValue");
	nodeName = myXML.getElementsByTagName("bankName");
	var oOption;
    var checkAdd = false;	
    fn_changeBankType(form);	
    //alert(form.bank_type.value);    
  	
	for(var i=0;i<nodeType.length ;i++)
	{
		//alert(bankType.item(i).firstChild.nodeValue);		
		if(bankType.item(i).firstChild.nodeValue != form.bank_type.value) continue;
			
		if(form.HSIEN_ID.value == 'ALL'){		
			oOption = document.createElement("OPTION");
			if(form.CANCEL_NO.value == 'N'){//營運中				
			   if(BnType.item(i).firstChild.nodeValue != '2'){			   	  			   	
			      oOption.text=nodeName.item(i).firstChild.nodeValue;
  			      oOption.value=nodeValue.item(i).firstChild.nodeValue;   	    			     
			   }		
		    }else{//已裁撤		    
		       if(BnType.item(i).firstChild.nodeValue == '2'){		       
			      oOption.text=nodeName.item(i).firstChild.nodeValue;
  			      oOption.value=nodeValue.item(i).firstChild.nodeValue;   	  
			   }
		    }				  		
  			checkAdd=false;
			for(var k =0;k<form.BankListDst.length;k++){			
				if(form.BankListDst.options[k].text == oOption.text){		    
			   	   checkAdd = true;			       
		    	}   
	    	}
	    	if(checkAdd == false && oOption.text != '' && oOption.value != ''){	  
	    		//alert('add '+oOption.text);
	    		//alert('add '+oOption.value);	    		
	    		if(form.bank_type.value == 'ALL'){//農漁會.預設全部放在已選欄位
  				   form.BankListDst.add(oOption); 
  				}else{
  				   form.BankListSrc.add(oOption); 
  				}
  			}	
	    }else if (nodeType.item(i).firstChild.nodeValue == form.HSIEN_ID.value){	    	  		
  			oOption = document.createElement("OPTION");
  			if(form.CANCEL_NO.value == 'N'){//營運中  				
			   if(BnType.item(i).firstChild.nodeValue != '2'){			   	  
			      oOption.text=nodeName.item(i).firstChild.nodeValue;
  			      oOption.value=nodeValue.item(i).firstChild.nodeValue;   	  
			   }		
		    }else{//已裁撤		    
		       if(BnType.item(i).firstChild.nodeValue == '2'){
			      oOption.text=nodeName.item(i).firstChild.nodeValue;
  			      oOption.value=nodeValue.item(i).firstChild.nodeValue;   	  
			   }
		    }			      		
			checkAdd=false;
			for(var k =0;k<form.BankListDst.length;k++){			
				if(form.BankListDst.options[k].text == oOption.text){		    
			   	   checkAdd = true;			       
		    	}   
	    	}
	    	if(checkAdd == false && oOption.text != '' && oOption.value != ''){	       
  				form.BankListSrc.add(oOption); 
  			}  			
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



function fn_ShowPanel(cnd){
	//act=BankList/RptColumn/RptOrder/RptType		
	this.document.forms[0].action = "/pages/DS001W.jsp?act="+cnd;	
	//if(cnd == 'createRpt'){
	//   this.document.forms[0].target = '_blank';
	//}else{
	   this.document.forms[0].target = '_self';
    //}
	this.document.forms[0].submit();
}