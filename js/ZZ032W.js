function doSubmit(form,cnd){		     
	     if(cnd == "List"){
	        if (!checkSingleYM(form.S_YEAR, form.S_MONTH)) {	     	
		        form.S_YEAR.focus();		      
		        return;
	        }	 
	     }
	    form.action="/pages/ZZ032W.jsp?act="+cnd+"&test=nothing";	    
	    if( cnd == "Lock" && (!checkData(form,'lock')) ) return;	   	    
	    if( cnd == "unLock" && (!checkData(form,'unlock'))) return;	   	    
	    if( cnd == "List" || cnd =="goQry") form.submit();	    	
	    if((cnd == "Lock")){
	    	if(!confirm("���@�~������X����(�`�N�G����ɽФűj��_�A�H�K�|���͵��G������)�A�A�T�w�n����H")) return;    	 		 	
	    	form.submit();	    	    
	    }	 
	    if((cnd == "unLock")){
	    	if(!confirm("���@�~������X����(�`�N�G����ɽФűj��_�A�H�K�|���͵��G������)�A�A�T�w�n����H")) return;    	 		 	
	    	form.submit();	    
	    }	
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
	var item = form.TRANS_TYPE.value;
	var oOption;	
	var mydate= new Date();
    var month = mydate.getMonth()+1;
    var year = mydate.getYear();
    var now_month;
    
	//alert(item);
	//alert(item.substring(item.indexOf(":")+1,item.length));
	//alert('year='+year);
	//alert('month='+month);
	//�ӳ��~��(�u)=========================================================================================
	form.S_MONTH.length = 0;
	for(i=1;i<=12;i++){
		oOption = document.createElement("OPTION");
		if(i<10){
		   oOption.text="0"+i;
	    }else{
	       oOption.text=i;
	    }
  		oOption.value=i;
  		form.S_MONTH.add(oOption);  		
  		if(i == 4 && item.substring(item.indexOf(":")+1,item.length) == 'Q'){//�u��		
  			break;  			
  	    }	
    }	
    if(item.substring(item.indexOf(":")+1,item.length) == 'M'){//���		  		  			
       if(month == '1'){
       	  form.S_YEAR.value=year-1911-1;
       	  now_month="12";
       }else{
       	  form.S_YEAR.value=year-1911;
       	  now_month = now_month-1;
       }
    }else if(item.substring(item.indexOf(":")+1,item.length) == 'Q'){//�u��		
       if(month == '1' ||  month == '2' || month == '3'){		   
  		  form.S_YEAR.value=year-1911-1;
  		  now_month="4";	
  	   }else if(month == '4' ||  month == '5' || month == '6'){  	   
  	      form.S_YEAR.value=year-1911;
  		  now_month="1";		      	           	
  	   }else if(month == '7' ||  month == '8' || month == '9'){  	       
  		  form.S_YEAR.value=year-1911;
  		  now_month="2";		      	           	
  	   }else if(month == '10' ||  month == '11' || month == '12'){  	       
  	      form.S_YEAR.value=year-1911;
  	      now_month="3";		      	           	
  	   }
    }	
  	   
    setSelect(form.S_MONTH,now_month);//set selected    	
  	
  	//�ˮֵ��G===========================================================================================  	
  	form.UPD_CODE.length = 0;
  	if(item.substring(0,item.indexOf(":")) == 'C04' || item.substring(0,item.indexOf(":")) == 'C05'){  		
  		oOption = document.createElement("OPTION");
		oOption.text="����";
  		oOption.value="ALL";
  		form.UPD_CODE.add(oOption);  
  		oOption = document.createElement("OPTION");
		oOption.text="�w�ӳ�";
  		oOption.value="0";
  		form.UPD_CODE.add(oOption);  
  		oOption = document.createElement("OPTION");
		oOption.text="���ӳ�";
  		oOption.value="1";
  		form.UPD_CODE.add(oOption);  
  		oOption = document.createElement("OPTION");
		oOption.text="���f��";
  		oOption.value="2";
  		form.UPD_CODE.add(oOption);  
    }else{
    	oOption = document.createElement("OPTION");
		oOption.text="����";
  		oOption.value="ALL";
  		form.UPD_CODE.add(oOption);  
  		oOption = document.createElement("OPTION");
		oOption.text="�w�ӳ�";
  		oOption.value="0";
  		form.UPD_CODE.add(oOption);  
  		oOption = document.createElement("OPTION");
		oOption.text="���ӳ�";
  		oOption.value="1";
  		form.UPD_CODE.add(oOption);  
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
