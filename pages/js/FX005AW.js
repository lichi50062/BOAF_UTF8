function AskDelete_Permission() {
  if(confirm("確定要刪除此筆資料嗎？"))
    return true;
  else
    return false;
}

function AskInsert_Permission() {
  if(confirm("確定要新增此筆資料嗎？"))
    return true;
  else
    return false;
}
function AskEdit_Permission() {
  if(confirm("確定要修改此筆資料嗎？"))
    return true;
  else
    return false;
}


function changeSort(form){
	
	
	form.action="/pages/FX005AW.jsp?act=List&sort_type="+form.Sort.value;	        
	form.submit();
	
	
		
		
	
	}


function doSubmit(form,cnd,bankno,seq_no){
     	  
if(cnd =='add'){
	if(!checkData(form)) return;	
	
	if(!(form.cancel_type.value==1) || !(form.cancel_type.value==2) )
	{		
		  
			form.action="/pages/FX005AW.jsp?act=Add_Check_Property_No&property_no="+form.property_no.value+"&bank_no="+bankno+"&cancel_type="+form.cancel_type.value+"&site_name="+form.site_name.value+"&addr="+form.addr.value+"&setup_date_y="+form.SETUP_DATE_Y.value+"&setup_date_m="+form.SETUP_DATE_M.value+"&setup_date_d="+form.SETUP_DATE_D.value+"&machine_name="+form.machine_name.value+"&cancel_type="+form.cancel_type.value+"&CANCEL_DATE="+form.CANCEL_DATE.value+"&CANCEL_DATE_Y="+form.CANCEL_DATE_Y.value+"&CANCEL_DATE_M="+form.CANCEL_DATE_M.value+"&CANCEL_DATE_D="+form.CANCEL_DATE_D.value;	    				    			
			form.submit();
  }else{ 
  if(AskInsert_Permission()){
  		form.action="/pages/FX005AW.jsp?act=Insert";	        
	    form.submit();
  }
  }
}
else if(cnd =='modify'){
	
	if(!checkData(form)) return;	
	
  if(!(form.cancel_type.value==1 || form.cancel_type.value==2) )
		{	
			if(AskEdit_Permission())
			{									
				form.action="/pages/FX005AW.jsp?act=Modify_Check_Property_No&property_no="+form.property_no.value+"&bank_no="+bankno+"&cancel_type="+form.cancel_type.value+"&site_name="+form.site_name.value+"&addr="+form.addr.value+"&setup_date_y="+form.SETUP_DATE_Y.value+"&setup_date_m="+form.SETUP_DATE_M.value+"&setup_date_d="+form.SETUP_DATE_D.value+"&machine_name="+form.machine_name.value+"&seq_no="+seq_no;	    				    						
				form.submit();
  		}else{
  			return;
  		}
  		
  	}else{
  				if(AskEdit_Permission())
					{				
							form.action="/pages/FX005AW.jsp?act=Update";	    
							form.submit();
					}else{
  						return;
  				}
	}
	
}
else if(cnd =='delete'){
	if(AskDelete_Permission()){
	form.action="/pages/FX005AW.jsp?act=Delete";	    
	form.submit();
	}
}
else if(cnd =='load'){
	//form.action="/pages/FX007W.jsp?act=Load";	    
	//form.submit();
window.open("/pages/FX005AW.jsp?act=Load","LOAD","width=750,height=400,scrollbars =yes ","")
}
else if(cnd =='checkload'){
var loadseq;
if(form.CHECKSEQ == null)
{
		window.close();
		return;  
}
if(form.CHECKSEQ.length>0){
	for (i=0;i<form.CHECKSEQ.length;i++)
	{
		if (form.CHECKSEQ[i].checked)
		{
			loadseq = form.CHECKSEQ[i].value
			window.opener.location="/pages/FX005AW.jsp?act=Load&CHECKSEQ="+loadseq;	
			window.close();
		}
	}
}
else{
	if (form.CHECKSEQ.checked)
		{
			loadseq = form.CHECKSEQ.value
			window.opener.location="/pages/FX005AW.jsp?act=Load&CHECKSEQ="+loadseq;	
			window.close();
		}
}
	window.close();
}

else if(cnd =='returnList'){
if( confirm("確定回查詢頁？")){
	form.action="/pages/FX005AW.jsp?act=List";	    
	form.submit();
	}
}


else return;
	  	  	    
}



function checkData(form) 
{
  var ckDate;
	if(trimString(form.site_name.value) == "" ){
           alert("裝設地點名稱不可為空白");
           form.site_name.focus();
           return false;
    } 
    else if(trimString(form.addr.value) == "" ){
           alert("地址不可為空白");
           form.addr.focus();
           return false;
    }  
	   	    
        if (trimString(form.SETUP_DATE_Y.value)  != "" ){        
        	if(isNaN(Math.abs(form.SETUP_DATE_Y.value))){
               alert("裝設日期(年)不可輸入文字");    
			   form.SETUP_DATE_Y.focus();            
               return false;
            }
        }else{
			alert("裝設日期(年)不可空白");
			form.SETUP_DATE_Y.focus();
			return false;   
		}   
        if (trimString(form.SETUP_DATE_M.value) == "" ){
			alert("裝設日期(月)不可空白");
			form.SETUP_DATE_M.focus();
			return false;
		}			
		if (trimString(form.SETUP_DATE_D.value) == "" ){
			alert("裝設日期(日)不可空白");
			form.SETUP_DATE_D.focus();		
			return false;
		}	
    
    if (trimString(form.machine_name.value) == ""){
    	alert("機器品名不可空白");
			form.machine_name.focus();		
			return false;
    }
   
    if (trimString(form.property_no.value) == ""){
    	alert("機器編號不可空白");
			form.property_no.focus();		
			return false;
    } 
     
    

    ckDate = '' + (parseInt(form.SETUP_DATE_Y.value)+1911) + '/' + form.SETUP_DATE_M.value + '/' + form.SETUP_DATE_D.value;
       
    	if( fnValidDate(ckDate) != true){
        	alert('裝設日期為無效日期!!');
        	form.SETUP_DATE_D.focus();
        	return false;
    	} 
    	   
    	form.SETUP_DATE.value = ckDate;   	
    	    	
 var sDate=new Date()   
 var eDate=new Date()
   
sDate.setFullYear(parseInt(form.SETUP_DATE_Y.value)+1911,parseInt(form.SETUP_DATE_M.value)-1,parseInt(form.SETUP_DATE_D.value));   

if(form.cancel_type.value == '1' || form.cancel_type.value == '2'){
		if (trimString(form.CANCEL_DATE_Y.value)  != "" ){        
        	if(isNaN(Math.abs(form.CANCEL_DATE_Y.value))){
            	alert("遷移/裁撤日期(年)不可輸入文字");    
            	form.CANCEL_DATE_Y.focus();
            	return false;
        	}	
        }else{
			alert("遷移/裁撤日期(年)不可空白");
			form.CANCEL_DATE_Y.focus();
			return false;   
		}
			
		if (trimString(form.CANCEL_DATE_M.value) == "" ){
			alert("遷移/裁撤日期(月)不可空白");
			form.CANCEL_DATE_M.focus();
			return false;
		}			
		if (trimString(form.CANCEL_DATE_D.value) == "" ){
			alert("遷移/裁撤日期(日)不可空白");
			form.CANCEL_DATE_D.focus();
			return false;
		}	       
		
		 
        
    	ckDate = '' + (parseInt(form.CANCEL_DATE_Y.value)+1911) + '/' + form.CANCEL_DATE_M.value + '/' + form.CANCEL_DATE_D.value;
       
    	if( fnValidDate(ckDate) != true){
           	alert('遷移/裁撤日期為無效日期!!');
           	form.CANCEL_DATE_D.focus();
           	return false;
   		}	          		
   		
   		form.CANCEL_DATE.value = ckDate;   	    	    
   		eDate.setFullYear(parseInt(form.CANCEL_DATE_Y.value)+1911,parseInt(form.CANCEL_DATE_M.value)-1,parseInt(form.CANCEL_DATE_D.value));   
		
		if( sDate>eDate){
			    
        	alert('遷移/裁撤 日期有誤!');
        	form.CANCEL_DATE_Y.focus();
        	return false;
    	}
   
    
}


  return true;
}

