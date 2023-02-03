function doSubmit(form,cnd){	   
	     if (!checkSingleYM(form.S_YEAR, form.S_MONTH)) {
		      form.S_YEAR.focus();
		      return;
	     } 
    	 if(cnd == 'GenerateRpt'){
    	 	if(!checkSelect(form)) return;	     	 	
    	 	if(!confirm("���@�~������X����(�`�N�G����ɽФűj��_�A�H�K�|���͵��G������)�A�A�T�w�n����H")) return;    	 		 	
	     }	 	  	     
	     form.action="/pages/ZZ041W.jsp?act="+cnd+"&test=nothing";	    
	     form.submit();	    		    
}	

	
function checkSelect(form) 
{	
  var flag = false;  
  for (var i = 0 ; i < form.elements.length; i++) {    
    if ( form.elements[i].checked == true ) {
        flag = true;
    }    
  }
  if (flag == false) {       	
      alert('�Цܤֿ�ܤ@�������͵o�G������!');        	   
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

/*
�A(��)�|�O -->rpt_output_type == 'X',��ܹA�|/���|/�A���| 
              rpt_output_type == 'T',��ܹA���| 
�����O -->rpt_output_type == 'T',���"��������"          
          rpt_include <> 'X',���"��������"
          ��L,���"��������"�ΦU�ӿ����O
*/
function changeOption(form){	
    var myXML,hsienXML,rptCode,rptOutputType, rptInclude;
    //alert(form.RPT_CODE.value);
    myXML = document.all("ReportListXML").XMLDocument;    
    form.szRPT_OUTPUT_TYPE.length = 0;
    form.HSIEN_ID.length = 0;    
	rptCode = myXML.getElementsByTagName("rptCode");
	rptOutputType = myXML.getElementsByTagName("rptOutputType");
	rptInclude = myXML.getElementsByTagName("rptInclude");
	hsienXML = document.all("hsienListXML").XMLDocument;
	hsien_id = hsienXML.getElementsByTagName("hsien_id");
	hsien_name = hsienXML.getElementsByTagName("hsien_name");
	var oOption;	
	for(var i=0;i<rptCode.length ;i++)
	{		
		//alert(rptCode.item(i).firstChild.nodeValue);  
  		if ((rptCode.item(i).firstChild.nodeValue == form.RPT_CODE.value))  {  			
  			//�A(��)�|�O
  			if ((rptOutputType.item(i).firstChild.nodeValue == 'X'))  {//��ܹA�|/���|/�A���|  				  				 
  				 oOption = document.createElement("OPTION");
			     oOption.text="�A�|";
  			     oOption.value="6";  			     
  			     form.szRPT_OUTPUT_TYPE.add(oOption);
  			     oOption = document.createElement("OPTION");
  			     oOption.text="���|";
  			     oOption.value="7";
  			     form.szRPT_OUTPUT_TYPE.add(oOption);
  			     oOption = document.createElement("OPTION");
  			     oOption.text="�A���|";
  			     oOption.value="ALL";
  			     form.szRPT_OUTPUT_TYPE.add(oOption);
  		    }else if ((rptOutputType.item(i).firstChild.nodeValue == 'T'))  {//��ܹA���|  				  		    	 
  		    	 oOption = document.createElement("OPTION");
  		    	 oOption.text="�A���|";
  			     oOption.value="ALL";
  			     form.szRPT_OUTPUT_TYPE.add(oOption);
  			     
  		    }   		    
  		    //�����O
  		    if ((rptOutputType.item(i).firstChild.nodeValue == 'T'))  {
  		    	 //alert('rptouttype==T');
  			     //rpt_output_type = 'T',���"��������"
  			     oOption = document.createElement("OPTION");
  		    	 oOption.text="��������";
  			     oOption.value="ALL";
  			     form.HSIEN_ID.add(oOption);
  		    }else  if ((rptInclude.item(i).firstChild.nodeValue != 'X'))  {
  		    	 //alert('rptInclude !=X');
  		    	 //rpt_include <> 'X',���"��������"
  			     oOption = document.createElement("OPTION");
  		    	 oOption.text="��������";
  			     oOption.value="ALL";
  			     form.HSIEN_ID.add(oOption);
  		    }else{
  		    	 //alert('rptInclude ==X rptouttype!=T');
  		    	 oOption = document.createElement("OPTION");
  		    	 oOption.text="��������";
  			     oOption.value="ALL";
  			     form.HSIEN_ID.add(oOption);
  			     //�U�ӿ����O
  		    	 for(var j=0;j<hsien_id.length ;j++){
  		    	     oOption = document.createElement("OPTION");
  		    	     oOption.text=hsien_name.item(j).firstChild.nodeValue;
  			         oOption.value=hsien_id.item(j).firstChild.nodeValue;
  			         form.HSIEN_ID.add(oOption);
  		    	 }
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
