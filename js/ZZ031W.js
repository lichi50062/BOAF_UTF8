//95.02.21 add ����A01��,�ˮֵ��G�~"A01_�O����B��0" 
//         add ����A04��,�ˮֵ��G�~"A04_�O����B��0" 
function doSubmit(form,cnd){	   
	     if (!checkSingleYM(form.S_YEAR, form.S_MONTH)) {
		      form.S_YEAR.focus();
		      return;
	     }  	    	  	     
	    form.action="/pages/ZZ031W.jsp?act="+cnd+"&test=nothing";	    
	    if( cnd == "Lock" && (!checkData(form,'lock')) ) return;	   	    
	    if( cnd == "unLock" && (!checkData(form,'unlock'))) return;	   	    
	    if( cnd == "Qry") form.submit();	    	
	    if((cnd == "Lock") && AskLock(form)) form.submit();	    	    
	    if((cnd == "unLock") && AskUnLock(form)) form.submit();	    
}	

function getData(form,cnd){	     	    	    
	    form.action="/pages/ZZ031W.jsp?act=getData&nowact="+cnd+"&test=nothing";
	    form.submit();	    
}	
function checkData(form,cnd) 
{	
  var flag = false;  
  for (var i = 0 ; i < form.elements.length; i++) {    
    if ( form.elements[i].checked == true ) {
        flag = true;
    }    
  }
  if (flag == false) {     
  	if(cnd == 'lock'){    
       alert('�Цܤֿ�ܤ@������w�����!');
    }
    if(cnd == 'unlock'){
       alert('�Цܤֿ�ܤ@�����Ѱ���w�����!');
    }	   
    return false;
  }
  return true;
}	

function selectAll(form) {  
  for ( var i = 0; i < form.elements.length; i++) {
      if((form.elements[i].type=='checkbox') && form.elements[i].disabled == false) {	
      	form.elements[i].checked = true;
      }	          	  
  }
  return;
}

function selectNo(form) {  
  for ( var i = 0; i < form.elements.length; i++) {
       if((form.elements[i].type=='checkbox') && form.elements[i].disabled == false) {	
      	 form.elements[i].checked = false;
       }	           
  }
  return;
}

function changeOption(form){
    var myXML,nodeType,nodeValue, nodeName;

    myXML = document.all("ReportListXML").XMLDocument;
    form.REPORT_NO.length = 0;
	nodeType = myXML.getElementsByTagName("transType");
	nodeValue = myXML.getElementsByTagName("reportValue");
	nodeName = myXML.getElementsByTagName("reportName");
	var oOption;	
	for(var i=0;i<nodeType.length ;i++)
	{		
  		if ((nodeType.item(i).firstChild.nodeValue == form.TRANS_TYPE.value))  {
  			oOption = document.createElement("OPTION");
			oOption.text=nodeName.item(i).firstChild.nodeValue;
  			oOption.value=nodeValue.item(i).firstChild.nodeValue;
  			form.REPORT_NO.add(oOption);
    	}
    }
}


//95.02.21 add ����A01��,�ˮֵ��G�~"A01_�O����B��0" 
//         add ����A04��,�ˮֵ��G�~"A04_�O����B��0" 
function changeResult(form){
	var oOption;
	//alert(form.REPORT_NO.value);
	if(form.TRANS_TYPE.value != '1'){
	   form.UPD_CODE.remove(5);	 
    }	
	if(form.REPORT_NO.value=='A01'){
		form.UPD_CODE.remove(5);
		oOption = document.createElement("OPTION");
		oOption.text="A01_�ӳ��O����B��0";
  		oOption.value="A01_990000_0";
	    form.UPD_CODE.add(oOption);	 	    	    
	}else if(form.REPORT_NO.value=='A04'){
		form.UPD_CODE.remove(5);
		oOption = document.createElement("OPTION");
		oOption.text="A04_�ӳ��O����B��0";
  		oOption.value="A04_840740_0";
	    form.UPD_CODE.add(oOption);	 	    	    
	}else  form.UPD_CODE.remove(5);
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
