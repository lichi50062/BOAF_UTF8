function changeVal(T1){
     
   pos=0
   var oring=T1.value
   pos=oring.indexOf(",")

      while (pos !=-1){
         oring=(oring).replace(",","")
         pos=oring.indexOf(",")
      }   

   return(oring)     
}

function changeMat(T1){
     
   
  
   if (isNaN(Math.abs(T1.value))) {
   	alert("請輸入數字");
   	return(T1.value);
   } else {
   	if ((Math.abs(T1.value)) >=100) {
   	   alert("逾放比不可大於100");
   	   return(T1.value);	
   	} else {
   	   return(T1.value);	
   	}	
   	
   }		
}

function changeStr(T1){
     
   pos=0
   c="";
   var oring=T1.value
  
   if (isNaN(Math.abs(T1.value))) {
   	alert("請輸入數字");
   	return(oring);
   }else{	
     if (eval(T1.value) < 0 ) {
      c="-";
     }   	

   T1.value= Math.abs(T1.value); 
   pos=oring.indexOf(",")
   if (pos==-1) {
      a=((T1.value).length) % 3
      b=(((T1.value).length)-a) / 3
      if (b==3){
         if (a==0){
            oring=(T1.value).substring(0,3)+","+(T1.value).substring(3,6)+","+(T1.value).substring(6)            
         }
         if (a==1){
            oring=(T1.value).substring(0,1)+","+ (T1.value).substring(1,4)+","+(T1.value).substring(4,7)+","+(T1.value).substring(7)
         }                           
      }
      if (b==2 ){
         if (a==0){
            oring=(T1.value).substring(0,3)+","+(T1.value).substring(3,6)
         }
         if (a==1){
            oring=(T1.value).substring(0,1)+","+ (T1.value).substring(1,4)+","+(T1.value).substring(4,7)
         }                           
         if (a==2){
            oring=(T1.value).substring(0,2)+","+ (T1.value).substring(2,5)+","+(T1.value).substring(5,8)
         }                                       
      }         
      if (b==1){
         if (a==0){
            oring=T1.value;
         }                           
         if (a==1){
            oring=(T1.value).substring(0,1)+","+ (T1.value).substring(1,4)
         }                           
         if (a==2){
            oring=(T1.value).substring(0,2)+","+ (T1.value).substring(2,5)
         }                                       
      }         
      if (b<0){
         oring=T1.value;
      }         
      
      
   }  
   if (c == "-" ) {
   	oring="-"+oring;
   }	
   return(oring)

}
	 

}