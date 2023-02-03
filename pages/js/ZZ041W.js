function doSubmit(form,cnd){	   
	     if (!checkSingleYM(form.S_YEAR, form.S_MONTH)) {
		      form.S_YEAR.focus();
		      return;
	     } 
    	 if(cnd == 'GenerateRpt'){
    	 	if(!checkSelect(form)) return;	     	 	
    	 	if(!confirm("本作業須執行幾分鐘(注意：執行時請勿強制中斷，以免會產生結果不完整)，你確定要執行？")) return;    	 		 	
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
      alert('請至少選擇一筆欲產生發佈的報表!');        	   
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
農(漁)會別 -->rpt_output_type == 'X',顯示農會/漁會/農漁會 
              rpt_output_type == 'T',顯示農漁會 
縣市別 -->rpt_output_type == 'T',顯示"全部縣市"          
          rpt_include <> 'X',顯示"全部縣市"
          其他,顯示"全部縣市"及各個縣市別
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
  			//農(漁)會別
  			if ((rptOutputType.item(i).firstChild.nodeValue == 'X'))  {//顯示農會/漁會/農漁會  				  				 
  				 oOption = document.createElement("OPTION");
			     oOption.text="農會";
  			     oOption.value="6";  			     
  			     form.szRPT_OUTPUT_TYPE.add(oOption);
  			     oOption = document.createElement("OPTION");
  			     oOption.text="漁會";
  			     oOption.value="7";
  			     form.szRPT_OUTPUT_TYPE.add(oOption);
  			     oOption = document.createElement("OPTION");
  			     oOption.text="農漁會";
  			     oOption.value="ALL";
  			     form.szRPT_OUTPUT_TYPE.add(oOption);
  		    }else if ((rptOutputType.item(i).firstChild.nodeValue == 'T'))  {//顯示農漁會  				  		    	 
  		    	 oOption = document.createElement("OPTION");
  		    	 oOption.text="農漁會";
  			     oOption.value="ALL";
  			     form.szRPT_OUTPUT_TYPE.add(oOption);
  			     
  		    }   		    
  		    //縣市別
  		    if ((rptOutputType.item(i).firstChild.nodeValue == 'T'))  {
  		    	 //alert('rptouttype==T');
  			     //rpt_output_type = 'T',顯示"全部縣市"
  			     oOption = document.createElement("OPTION");
  		    	 oOption.text="全部縣市";
  			     oOption.value="ALL";
  			     form.HSIEN_ID.add(oOption);
  		    }else  if ((rptInclude.item(i).firstChild.nodeValue != 'X'))  {
  		    	 //alert('rptInclude !=X');
  		    	 //rpt_include <> 'X',顯示"全部縣市"
  			     oOption = document.createElement("OPTION");
  		    	 oOption.text="全部縣市";
  			     oOption.value="ALL";
  			     form.HSIEN_ID.add(oOption);
  		    }else{
  		    	 //alert('rptInclude ==X rptouttype!=T');
  		    	 oOption = document.createElement("OPTION");
  		    	 oOption.text="全部縣市";
  			     oOption.value="ALL";
  			     form.HSIEN_ID.add(oOption);
  			     //各個縣市別
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
