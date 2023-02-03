function doSubmit(form,cnd){	   
	     if(!CheckQueryDate2(form.S_YEAR,form.S_MONTH,form.S_DATE,"起始日期")) return false;
         if(!CheckQueryDate2(form.E_YEAR,form.E_MONTH,form.E_DATE,"結束日期")) return false;     
	     if(trimString(form.S_DATE.value)!="" && trimString(form.E_DATE.value)!=""){
		    if(Math.abs(form.S_DATE.value) > Math.abs(form.E_DATE.value)){
    		   alert("起始查詢年月不可大於結束查詢年月");
    		   return false;
    	    }
    	 }   
    	 	  	     
	    form.action="/pages/ZZ042W.jsp?act="+cnd+"&test=nothing";	    
	    if((cnd == 'MultiFiles') && !checkSelect(form)){
	       return;	 
	    }	
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
      alert('請至少選擇一筆欲下載的報表!');        	   
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
    form.RPT_OUTPUT_TYPE.length = 0;
    form.HSIEN_ID.length = 0;
    form.RPT_SORT.length = 0;
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
  			     form.RPT_OUTPUT_TYPE.add(oOption);
  			     oOption = document.createElement("OPTION");
  			     oOption.text="漁會";
  			     oOption.value="7";
  			     form.RPT_OUTPUT_TYPE.add(oOption);
  			     oOption = document.createElement("OPTION");
  			     oOption.text="農漁會";
  			     oOption.value="ALL";
  			     form.RPT_OUTPUT_TYPE.add(oOption);
  		    }else if ((rptOutputType.item(i).firstChild.nodeValue == 'T'))  {//顯示農漁會  				  		    	 
  		    	 oOption = document.createElement("OPTION");
  		    	 oOption.text="農漁會";
  			     oOption.value="ALL";
  			     form.RPT_OUTPUT_TYPE.add(oOption);
  			     
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
  		    //列表順序
  		    if ((rptInclude.item(i).firstChild.nodeValue == 'T'))  {  		    	 
  			     oOption = document.createElement("OPTION");
  		    	 oOption.text="按報表年月、農(漁)會別";
  			     oOption.value="4";
  			     form.RPT_SORT.add(oOption);
  		    }else{
  		    	 oOption = document.createElement("OPTION");
  		    	 oOption.text="按縣市別、機構別、農(漁)會別、報表年月";
  			     oOption.value="1";
  			     form.RPT_SORT.add(oOption);
  			     oOption = document.createElement("OPTION");
  		    	 oOption.text="按縣市別、機構別、報表年月、農(漁)會別";
  			     oOption.value="2";
  			     form.RPT_SORT.add(oOption);
  			     oOption = document.createElement("OPTION");
  		    	 oOption.text="按報表年月、縣市別、機構別、農(漁)會別";
  			     oOption.value="3";
  			     form.RPT_SORT.add(oOption);
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
