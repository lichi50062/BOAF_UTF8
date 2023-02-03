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

function doSubmit(form,cnd){	     	    	  	     
   form.action="/pages/ZZ100W.jsp?act="+cnd;		 	   	         	     
   if(cnd =="Qry")
   {  				   		
          window.open("OpenHeadPerson.jsp");      	  	  	
   }
   if(cnd =="Delete")
   {  				   		
        var ret = confirm("是否要刪除歷程記錄(Y/N)?");		   		
        if(ret == true)               	 
         form.submit();         
        else
         return;       	  	  	
   }  
   if(cnd =="Clear")
   {  	
	var ret = confirm("是否要刪除理監事的附件檔案(Y/N)?");		   		
        if(ret == true)
	{
       	  form.submit(); 
        }
        else
        {         
          form.action="/pages/ZZ100W.jsp";
       	  form.submit(); 
        }  	  	  	
   }
   if(cnd =="New")
   {
	if(!checkID(form)) return;	
	var ret = confirm("是否要重新建立身份識別資料(Y/N)?");	
	if(ret == true)
	{
       	  form.submit(); 
        }
        else
        {         
          form.action="/pages/ZZ100W.jsp";
       	  form.submit(); 
        }  	
   }
   if(cnd =="CreateHeadText")
   {
	//95.02.06 拿掉 by 2495
	//if(!checkData(form)) return;			
	var ret = confirm("本作業約需3-10秒,是否作業(Y/N)?");	
	if(ret == true)
	{
       	  form.submit(); 
        }
        else
        {         
          form.action="/pages/ZZ100W.jsp";
       	  form.submit(); 
        }  	
   }		    
}


function checkID(form) 
{		
	if (trimString(form.user_id.value) =="" )
	{
		alert("請填身分識別帳號");
		form.user_id.focus();
		return false;
        }
	if (trimString(form.user_pwd.value) =="" )
	{
		alert("請填身分識別密碼");
		form.user_pwd.focus();
		return false;
        }
	if (trimString(form.ip_Address.value) =="" )
	{
		alert("請填身分識別來源IP");
		form.user_pwd.focus();
		return false;
        }	            	
  return true;
}


function checkData(form) 
{	
	var today = new Date();
	var today_year = today.getYear();
	    today_year = today_year-1911;
	var today_month = today.getMonth();
	    today_month++; 
	var today_day = today.getDate();
	
	 
  if(trimString(form.StartYear.value) == "" || trimString(form.EndYear.value) == "")
  {
    		alert("年份不可空白");
        return false;
  }
  
  if(form.StartYear.value.indexOf(".") != -1 || form.EndYear.value.indexOf(".") != -1)
  {
        alert("年份不可為小數");        
        return false;
  }
  
  if(isNaN(Math.abs(form.StartYear.value)) || isNaN(Math.abs(form.EndYear.value)))
  {
        alert("年份不可為文字");
        form.StartYear.focus();
        return false;
  }

  if(form.StartYear.value > today_year)
  {
  			alert("公告開始年份不可大於當年份");
  			form.StartYear.focus();
  			return false;
  }   
	
  if(form.EndYear.value > today_year)
  {
  			alert("公告結束年份不可大於當年份");
  			form.EndYear.focus();
  			return false;
  }  


  if(form.StartYear.value > form.EndYear.value)
  {
  			alert("公告結束年份不可小於當年份");
  			form.StartYear.focus();
  			return false;
  }
   
  if(form.StartYear.value == form.EndYear.value && (form.StartMonth.options[form.StartMonth.selectedIndex].value > form.EndMonth.options[form.EndMonth.selectedIndex].value))
  {
  			alert("結束月份不可小於當月份");
  			form.EndMonth.focus(); 
  			return false; 	 	
  }
  
  
  if(form.StartYear.value == form.EndYear.value && form.StartMonth.options[form.StartMonth.selectedIndex].value == form.EndMonth.options[form.EndMonth.selectedIndex].value && form.StartDay.options[form.StartDay.selectedIndex].value > form.EndDay.options[form.EndDay.selectedIndex].value)
  {
  			alert("結束日不可小於公告起始日");
  			form.EndDay.focus(); 
  			return false; 	 	
  }
 
  if(form.StartMonth.value == 2 )
  {
			
		if(leapYear (form.StartYear.value)== true) 
		{
			if(form.StartDay.value>29)
			{
				alert("二月份輸入結束日錯誤!!");
  				form.StartDay.focus(); 
  				return false;
			}
		}
		if(leapYear (form.StartYear.value)== false) 
		{
			if(form.StartDay.value>28)
			{
				alert("二月份輸入結束日錯誤!!");
  				form.StartDay.focus(); 
  				return false;
			}
		} 		
  }

  if(form.EndMonth.value == 2 )
  {
     		if(leapYear (form.EndYear.value)== true) 
		{
			if(form.EndDay.value>29)
			{
				alert("二月份輸入結束日錯誤!!");
  				form.EndDay.focus(); 
  				return false;
			}
		}
		if(leapYear (form.EndYear.value)== false) 
		{
			if(form.EndDay.value>28)
			{
				alert("二月份輸入結束日錯誤!!");
  				form.EndDay.focus(); 
  				return false;
			}
		} 
 		
  }     

  if((form.StartMonth.value == 2 ||　form.StartMonth.value == 4 ||form.StartMonth.value == 6 ||form.StartMonth.value == 9||form.StartMonth.value == 11) && form.StartDay.value>30)
  {
		alert("輸入起始日錯誤!!");
  		form.StartDay.focus(); 
  		return false; 		
  }   

  if((form.EndMonth.value == 2 ||　form.EndMonth.value == 4 ||form.EndMonth.value == 6 ||form.EndMonth.value == 9||form.EndMonth.value == 11) && form.EndDay.value>30)
  {
		alert("輸入結束日錯誤!!");
  		form.EndDay.focus(); 
  		return false; 		
  } 

          	
  return true;
}