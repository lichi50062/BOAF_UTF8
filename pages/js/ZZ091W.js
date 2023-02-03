function doSubmit(form,cnd,seq_no){	     	    	  	     
	 form.action="/pages/ZZ091W.jsp?act="+cnd;	    
	   	         	    
	 if( cnd == "Back")
	 {  	 
	 	form.action="/pages/ZZ091W.jsp?act=List";	    
	    	form.submit();
	 }
         if( cnd == "Edit")
	 {  	 
	 			form.action="/pages/ZZ091W.jsp?act="+cnd+"&seq_no="+seq_no;	    
	    	form.submit();
	 }
	 if(cnd == "new" ) form.submit();
	 
	 if(cnd == "insert")
	 {  	 	  
	 	if(!checkData(form)) return;		   		  		
                var ret = confirm("是否要新增公告(Y/N)?");		   		
   	        if(ret == true)
		{
		   form.action="/pages/ZZ091W.jsp?act="+cnd+"&headmark="+seq_no+"&notifyUrl="+form.notifyUrl.value+"&head_title="+form.head_title.value+"&StartYear="+form.StartYear.value+"&StartMonth="+form.StartMonth.value+"&StartDay="+form.StartDay.value+"&StartHour="+form.StartHour.value+"&EndYear="+form.EndYear.value+"&EndMonth="+form.EndMonth.value+"&EndDay="+form.EndDay.value;	    
	           form.submit();
		}
                else
                   return; 
   }
         
   if(cnd =="del") 
   {
     var ret = confirm("是否要刪除勾選的公告(Y/N)?");		   		
       if(ret == true)
       {      	 
         form.submit(); 
       }
       else
         return; 
   }
   
   if(cnd =="Update")
   {  	　
   	  if(!checkData(form)) return;　
   	  var ret = confirm("是否更新本公告相關資料?");
          if(ret == true)
	  {   			   		
   		form.action="/pages/ZZ091W.jsp?act="+cnd+"&seq_no="+seq_no+"&head_title="+form.head_title.value+"&notifyUrl="+form.notifyUrl.value+"&StartYear="+form.StartYear.value+"&StartMonth="+form.StartMonth.value+"&StartDay="+form.StartDay.value+"&StartHour="+form.StartHour.value+"&EndYear="+form.EndYear.value+"&EndMonth="+form.EndMonth.value+"&EndDay="+form.EndDay.value;	    
	        form.submit();
          }
          else
          {
                form.action="/pages/ZZ091W.jsp?act=Edit&seq_no="+seq_no;
       	        form.submit(); 
          }
   }
   
   if(cnd =="Clear")
   {  	
	var ret = confirm("是否要刪除本公告的原始附件檔案(Y/N)?");		   		
        if(ret == true)
	{
          form.action="/pages/ZZ091W.jsp?act="+cnd+"&seq_no="+seq_no;
       	  form.submit(); 
        }
        else
        {
          form.action="/pages/ZZ091W.jsp?act=Edit&seq_no="+seq_no;
       	  form.submit(); 
        }  	  	  	
   } 
   
   if(cnd =="Open")
   {  				   		
          form.action="/pages/OpenNotify.jsp?seq_no="+seq_no;
       	  form.submit();         	  	  	
   }
   
   
      		    
}
       		    
function checkData(form) 
{	
	var today = new Date();
	var today_year = today.getYear();
	    today_year = today_year-1911;
	var today_month = today.getMonth();
	    today_month++; 
	var today_day = today.getDate();
	
	
	if (trimString(form.head_title.value) =="" ){
				alert("公告標題不可空白");
				form.head_title.focus();
				return false;
	}	
	 
  if(trimString(form.StartYear.value) == "" || trimString(form.EndYear.value) == "")
  {
    		alert("公告年份不可空白");
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
  
  if(form.StartYear.value < today_year)
  {
        alert("開始公告年份不可小於當年份");  
        form.StartYear.focus();     
        return false;
  }
 
 
  if(form.StartYear.value > form.EndYear.value)
  {
  			alert("公告結束年份不可小於當年份");
  			form.EndYear.focus();
  			return false;
  }
   
  if(form.StartYear.value == form.EndYear.value && (form.StartMonth.options[form.StartMonth.selectedIndex].value > form.EndMonth.options[form.EndMonth.selectedIndex].value))
  {
  			alert("公告結束月份不可小於當月份");
  			form.EndMonth.focus(); 
  			return false; 	 	
  }
  
  
  if(form.StartYear.value == form.EndYear.value && form.StartMonth.options[form.StartMonth.selectedIndex].value == form.EndMonth.options[form.EndMonth.selectedIndex].value && form.StartDay.options[form.StartDay.selectedIndex].value > form.EndDay.options[form.EndDay.selectedIndex].value)
  {
  			alert("公告結束日不可小於公告起始日");
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
