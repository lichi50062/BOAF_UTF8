function doSubmit(form,cnd){
	chgInputState(form);
	if (trimString(form.orgcate.value)==1) {
		if (trimString(form.TBANK_NO.value)=="" || trimString(form.BANK_NO.value)=="") {
			alert('總分支機構代碼皆須輸入！');
			return;
		}
	}
		else if (trimString(form.orgcate.value)==0 && trimString(form.TBANK_NO.value)=="") {
				alert('總機構代碼須輸入！');
				return;
			}
	if (trimString(form.user_pw.value)=="") {
		alert('請輸入「密碼」！');
		return;
	}
	form.action="/pages/ZZ099W.jsp?act="+cnd;
	if(cnd=="Delete" && (!checkData(form,'delete'))) return;
	if(cnd=="Qry") form.submit();
	if(cnd=="Delete" && AskDelete(form)) form.submit();
}



function getData(form,cnd){
	    form.action="/pages/ZZ099W.jsp?act=getData&nowact="+cnd+"&test=nothing";
	    form.submit();
}



function checkData(form,cnd) {
	var flag=false;
	for (var i=0 ; i < form.elements.length; i++) {
		if ( form.elements[i].checked==true ) {
			flag=true;
    	}
  	}
	if (flag==false) {
	  	if (cnd=='delete') {
			alert('請至少選擇一筆欲刪除的資料！');
	    }
	    return false;
	}
	return true;
}



function checkChoice(form,cnd) {
	if (trimString(form.orgcate.value)==1 && cnd=="isEpt") {
		if (trimString(form.TBANK_NO.value)=="" || trimString(form.bank_no.value)=="") {
			alert('總分支機構代碼皆須輸入！');
			return false;
		}
	}
		else if (trimString(form.orgcate.value)==0 && cnd=="isEpt") {
			if (trimString(form.TBANK_NO.value)=="") {
				alert('總機構代碼須輸入！');
				return false;
			}
		}
		else return true;
}



function selectAll(form) {
  for ( var i=0; i < form.elements.length; i++) {
      if((form.elements[i].type=='checkbox') && form.elements[i].disabled==false) {
      	form.elements[i].checked=true;
      }
  }
  return;
}

function selectNo(form) {
  for ( var i=0; i < form.elements.length; i++) {
       if((form.elements[i].type=='checkbox') && form.elements[i].disabled==false) {
      	 form.elements[i].checked=false;
       }
  }
  return;
}

function changeOption(form){
    var myXML,nodeType,nodeValue, nodeName;

    myXML=document.all("ReportListXML").XMLDocument;
    form.REPORT_NO.length=0;
	nodeType=myXML.getElementsByTagName("transType");
	nodeValue=myXML.getElementsByTagName("reportValue");
	nodeName=myXML.getElementsByTagName("reportName");
	var oOption;
	for(var i=0;i<nodeType.length ;i++)
	{
  		if ((nodeType.item(i).firstChild.nodeValue==form.TRANS_TYPE.value))  {
  			oOption=document.createElement("OPTION");
			oOption.text=nodeName.item(i).firstChild.nodeValue;
  			oOption.value=nodeValue.item(i).firstChild.nodeValue;
  			form.REPORT_NO.add(oOption);
    	}
    }
}
function setSelect(S1, bankid) {
    if(S1==null)
    	return;
    for(i=0;i<S1.length;i++) {
      	if(S1.options[i].value==bankid)    	{
        	S1.options[i].selected=true;
        	break;
    	}
    }
}



function chgInputState(form) {
	val=form.orgcate.value
	if (val==0) {
		form.BANK_NO.value="";
		form.BANK_NO.disabled=true;
	}
	if (val==1) form.BANK_NO.disabled=false;
	if (val=="") {
		alert("請選擇「機構種類」！");
		form.orgcate.focus();
		return;
	}
}