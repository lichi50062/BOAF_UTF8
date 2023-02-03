//94.04.07 add 營運中/已裁撤 by 2295
//99.10.08 add 根據查詢年度.縣市別.改變總機構名稱 by 2295
function changeOption(form,cnd){
    var myXML,nodeType,nodeValue, nodeName;
	
	var m_year = form.S_YEAR.value;
    //alert(m_year);    
    if(m_year >= 100){
       m_year = 100;
    }else{
       m_year = 99;
    }	
    
    myXML = document.all("TBankXML").XMLDocument;
    form.BankListSrc.length = 0;
    if(cnd == 'change') form.BankListDst.length = 0;
    BnType = myXML.getElementsByTagName("BnType");    
	nodeType = myXML.getElementsByTagName("HsienId");
	nodeValue = myXML.getElementsByTagName("bankValue");
	nodeName = myXML.getElementsByTagName("bankName");
	nodeYear = myXML.getElementsByTagName("BankYear");//m_year
	
	var oOption;
    var checkAdd = false;	
	
	for(var i=0;i<nodeType.length ;i++)
	{
		if(form.HSIEN_ID.value == 'ALL'){
			if(nodeYear.item(i).firstChild.nodeValue != m_year){
			   continue;//顯示查詢年度的新機構名稱
			}
			oOption = document.createElement("OPTION");
			oOption.text=nodeName.item(i).firstChild.nodeValue;
  			oOption.value=nodeValue.item(i).firstChild.nodeValue;   	  
			 
  			checkAdd=false;
			for(var k =0;k<form.BankListDst.length;k++){			
				if(form.BankListDst.options[k].text == oOption.text){		    
			   	   checkAdd = true;			       
		    	}   
	    	}
	    	if(checkAdd == false && oOption.text != '' && oOption.value != ''){	  
	    		//alert('add '+oOption.text);
	    		//alert('add '+oOption.value);
  				form.BankListSrc.add(oOption); 
  			}	
	    }else if (nodeType.item(i).firstChild.nodeValue == form.HSIEN_ID.value){
	    	if(nodeYear.item(i).firstChild.nodeValue != m_year){
			   continue;//顯示查詢年度的新機構名稱
			}
  			oOption = document.createElement("OPTION");		
			oOption.text=nodeName.item(i).firstChild.nodeValue;
  			oOption.value=nodeValue.item(i).firstChild.nodeValue;   	  
			  
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

//99.10.08 add 根據查詢年月.改變縣市別名稱
function changeCity(xml, target, source, form) {	
      var myXML,nodeType,nodeValue, nodeName,nodeYear,m_year;      
      m_year = source.value;
      if(m_year >= 100){
         m_year = 100;
      }else{
         m_year = 99;
      }	
      
      target.length = 0;      
      var oOption;     
      myXML = document.all(xml).XMLDocument;
      nodeType = myXML.getElementsByTagName("cityType");//hsien_id
      nodeYear = myXML.getElementsByTagName("cityYear");//m_year
	  nodeValue = myXML.getElementsByTagName("cityValue");//hsien_id
	  nodeName = myXML.getElementsByTagName("cityName");//hsien_name
		
	  oOption = document.createElement("OPTION");
	  oOption.text='全部';
  	  oOption.value='ALL';
  	  target.add(oOption);
  	  
	  for(var i=0;i<nodeType.length ;i++)	{	  	
  	     if (nodeYear.item(i).firstChild.nodeValue == m_year)  {
  		    oOption = document.createElement("OPTION");
		    oOption.text=nodeName.item(i).firstChild.nodeValue;
  		    oOption.value=nodeValue.item(i).firstChild.nodeValue;
  		    target.add(oOption);
   	     }
      }
      form.HSIEN_ID[0].selected=true;
      changeOption(this.document.forms[0],'');
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

function doSubmit(form, myfun) {
	if (!checkSingleYM(form.S_YEAR, form.S_MONTH)) return;
	if(form.BankListDst.length == 0){      	 
       alert('金融機構代碼必須選擇');
       return;
    }
    MoveSelectToBtn(this.document.forms[0].BankList, this.document.forms[0].BankListDst);	
	var _id = form.Report_no.value;	
	var DLId = form.Report_no.value.substr(0, 1);	
	form.action="/pages/WMFileDeleteBatch.jsp?act="+myfun+"&test=nothing";
	if(AskDelete(form)) form.submit();	    	       	       	
}
