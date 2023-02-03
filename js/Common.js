//94.04.01 add �ˮ֨������r��
//94.11.01 add �ˬdE-Mail�O�_���~
//95.02.08 add myConfirm��ܮ�,�w�]�b����
//97.01.10 add �ˬd��J���r�����
//97.09.26 fix �ץ�������o�榡 by 2295
function check_identity_no(identity_no) {
   tab = "ABCDEFGHJKLMNPQRSTUVWXYZIO"
   A1 = new Array (1,1,1,1,1,1,1,1,1,1,2,2,2,2,2,2,2,2,2,2,3,3,3,3,3,3 );
   A2 = new Array (0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5,6,7,8,9,0,1,2,3,4,5 );
   Mx = new Array (9,8,7,6,5,4,3,2,1,1);

   if ( identity_no.length != 10 ){        
        return false;
   }
   i = tab.indexOf( identity_no.charAt(0) );
   if ( i == -1 ){        
        return false;
   }
   sum = A1[i] + A2[i]*9;

   for ( i=1; i<10; i++ ) {
      v = parseInt( identity_no.charAt(i) );
      if ( isNaN(v) ){          
           return false;
      }
      sum = sum + v * Mx[i];
   }
   if ( sum % 10 != 0 ){       
       return false;
   }

   return true;
}


/*************************************************
�\��G�i�H�ۤv�w�confirm���ܮءA�bIE6�U���ճq�L

�o��confirm���٥i�H���n�h�X�i�A��p�ק�I�����C��A�ק墨����ܪ��Ϥ��A�ק���᪺�˦��A

**************************************************/

/*
�Ѽƻ����G strTitle   confirm�ت����D
   		   strMessage confirm�حn��ܪ�����
   		   intType    confirm�ت��襤������1���襤YES�A2���襤NO
           intWidth   confirm�ت��e��
           intHeight  confirm�ت�����
*/
//95.02.08 add myConfirm��ܮ�,�w�]�b����
function myConfirm(strTitle,strMessage,intType,intWidth,intHeight)
{
 var strDialogFeatures = "dialogHeight:"+intHeight+";dialogWidth:"+intWidth+";center:yes;help:no;resizable:yes;scroll:no;status:no;defaultStatus:no;";
 
 var args = new Array();
 args[args.length] = strTitle;
 args[args.length] = strMessage;
 args[args.length] = intType;
 var result = showModalDialog("myConfirm.htm",args,strDialogFeatures);
 //var result = showModelessDialog("myConfirm.htm",args,strDialogFeatures);
 
 return result;
}



function AskReset(form) {
  if(confirm("�T�w�n�^�_�즳��ƶܡH")){  	
  	form.reset();    	 	
  }  
  return;
}

function AskDelete() {
	
  //if(confirm("�T�w�n�R���ܡH"))
  var myConfirmResult = myConfirm("Microsoft Internet Explorer","�T�w�n�R���ܡH",2,15,8)
  if(myConfirmResult)  
    return true;
  else
    return false;
}
function AskUpdate() {
  var myConfirmResult = myConfirm("Microsoft Internet Explorer","�T�w�n�ק�ܡH",2,15,8)
  if(myConfirmResult)
    return true;
  else  
  	return false;
  /*	
  if(confirm("�T�w�n�ק�ܡH"))
    return true;
  else
    return false;
  */  
}
function AskInsert() {
  //if(confirm("�T�w�n�s�W�ܡH"))
  var myConfirmResult = myConfirm("Microsoft Internet Explorer","�T�w�n�s�W�ܡH",2,15,8)
  if(myConfirmResult)
    return true;
  else
    return false;
}

function AskAbdicate() {
  if(confirm("�T�w�n�����ܡH")) {
    return true;
  }else {
    return false;
  }  
}
//93.11.26 add
function AskRevoke() {
  if(confirm("�T�w�n���M�ܡH")) {
    return true;
  }else {
    return false;
  }  
}
//93.11.26
function Add_Manager() {
	alert("�Х���J�D�ɸ��");
	return false;
}

//93.12.10 add
function AskResetPwd() {
  if(confirm("�T�w�n���]�K�X�ܡH")) {
    return true;
  }else {
    return false;
  }  
}
//94.01.03 add by 2295
function AskLock() {
  if(confirm("�T�w�n��w�ܡH")) {
    return true;
  }else {
    return false;
  }  
}
//94.01.03 add by 2295
function AskUnLock() {
  if(confirm("�T�w�n�Ѱ���w�ܡH")) {
    return true;
  }else {
    return false;
  }  
}

function changeVal(T1) {
	//alert('changeval:->' + T1.value);
    pos = 0
    var oring = T1.value;
    pos = oring.indexOf(",");

    while (pos != -1) {
        oring = (oring).replace(",", "");
        pos = oring.indexOf(",");
    }
    //alert('oring:->' + oring);
    return oring;

}
//================================================================
function changeStr(T1) {
	//alert(T1);
    T1.value = changeVal(T1);
    //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    c="";
    var oring = T1.value
    var t1v1  = T1.value
    var t1v2  = "";
   //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    if (isNaN(Math.abs(T1.value)))
    {
		//alert(T1.value);
   	    alert("�п�J�Ʀr");
   	    return oring;
    }else{
        if (eval(T1.value) < 0 )
            c="-";
    }

    if(T1.value.length == 0)
        return oring;
    if(Math.abs(T1.value) == 0)
        return oring;

    T1.value= Math.abs(T1.value);
    t1v1  = T1.value

    if((pos=t1v1.indexOf(".")) != -1)
    {
        t1v2 = t1v1.substring(pos,((T1.value).length));
        t1v1 = t1v1.substring(0,pos);
    }
  //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    pos=oring.indexOf(",")
    if (pos==-1)
    {
        var len   = t1v1.length;
        a=Math.floor(len % 3)
        b=Math.floor(len / 3 -1)
        if(a !=0 && b >= 0)
            oring = t1v1.substring(0,a) + ",";
        else if(b<0)
                oring = t1v1
             else
                oring = "";
        for(i=0;i<b;i++)
        {
            oring += t1v1.substring(a,a+3) + ",";
            a += 3;
        } // end of for

        oring += t1v1.substring(a,len);

        if (c == "-" )
        {
            oring="-"+oring;
        }
        if(t1v2 != "")
            oring += t1v2;

        return oring
     }
}
//===���@�H���e�����CHECK=============
function Check_Maintain(form1) {
   if (trimString(form1.maintain_dept.value)=="" || trimString(form1.maintain_tel.value)==""){
        alert("���@�����W��.�ӿ���q����줣�i�ťաI");
        return false;
   }
   return true
}
//==================================================
function CheckYear(cyear) {
    if(cyear.value.indexOf(".") != -1)
    {
        alert("�~�����i���p��");
        return false;
    }
    if(isNaN(Math.abs(cyear.value)))
    {
        alert("�~�����i����r");
        return false;
    }
    if(trimString(cyear.value) != '')
        cyear.value = Math.abs(cyear.value);
    return true;
}
//================================================================
function trimString(inString) {

	var outString;
	var startPos;
	var endPos;
	var ch;

	// where do we start?
	startPos = 0;
	test = 0;
	ch = inString.charAt(startPos);
	while ((ch == " ") || (ch == "\b") || (ch == "\f") || (ch == "\n") || (ch == "\r") || (ch == "\n")) {
		startPos++;
		if ( ch==" " ) {
			test++;
		}
		ch = inString.charAt(startPos);
	}
     if  ( test==inString.length )
     	flag = true;
     else
     	flag = false;
	endPos = inString.length - 1;
	ch = inString.charAt(endPos);
	while ((ch == " ") || (ch == "\b") || (ch == "\f") || (ch == "\n") || (ch == "\r") || (ch == "\n")) {
		endPos--;
		ch = inString.charAt(endPos);
	}

	// get the string
	outString = inString.substring(startPos, endPos + 1);
	if ( flag==true )
		return "";
	else
		return outString;
}
function ChgAction(form, actionURL) {

    form.action = actionURL;
	form.submit();
}
//=================================================================
function checkSingleYM(checkyear, checkmonth) {

	if (!CheckYear(checkyear)){		
    	return false;
    }	

//    if (trimString(checkyear.value) == "" || checkmonth.options[checkmonth.selectedIndex].value == ""){
    if (trimString(checkyear.value) == "" || checkmonth.value == ""){
		alert("�п�J���:");
        return false;
    }
    return true;
}

//=================================================================
function CheckQueryDate2(checkyear, checkmonth, checkdate, szWarning) {

    if(!CheckYear(checkyear))
    	return false;
    if(trimString(checkyear.value)!="" && checkmonth.options[checkmonth.selectedIndex].value!="") {
        checkdate.value = Math.abs(checkyear.value) * 100 + Math.abs(checkmonth.options[checkmonth.selectedIndex].value);
    }
    else if(trimString(checkyear.value)=="" && checkmonth.options[checkmonth.selectedIndex].value==""){}
         else {
            alert("�п�J���㪺"+ szWarning+ "�Ϊ̳�����J");
            return false;
         }
    return true
}
//=========================================================
function Check_QueryDate(form) {

    if(!CheckQueryDate2(form.S_YEAR,form.S_MONTH,form.S_DATE,"�_�l���"))
    	return false;
    if(!CheckQueryDate2(form.E_YEAR,form.E_MONTH,form.E_DATE,"�������"))
        return false;
	if(trimString(form.S_DATE.value)!="" && trimString(form.E_DATE.value)!="") {
		if(Math.abs(form.S_DATE.value) > Math.abs(form.E_DATE.value)) {
    		alert("�_�l������i�H�j��פ���");
    		return false;
    	}
    }
    return true;
}
//=========================================================
function checkNumber(Rate1){

	Rate1.value = changeVal(Rate1);

    if (isNaN(Math.abs(Rate1.value))) {
   	    alert("�п�J�Ʀr");
   	    return false;
    }
    if ((Rate1.value).indexOf(".") != -1 ) {
    	alert("���i��J�p���I");
        return false;
    }
    return true;
}
//add by 2295 �Y���i���p���I��,�h���b��text field
function checkPoint_focus(Rate1){

	//93.03.16����Rate1.value = changeVal(Rate1);
    if ((changeVal(Rate1)).indexOf(".") != -1 ) {
    	alert("���i��J�p���I");
    	Rate1.focus();
        return false;
    }
    return true;
}
//97.09.26 fix �ץ�������o�榡 by 2295
function fnValidDate(dateStr) {

    var leap = 28;
     var tmpStr = dateStr.substring(5,dateStr.length);
    //alert(tmpStr);   
    //alert('�~='+dateStr.substring(0,4));
    //alert('��='+tmpStr.substring(0, tmpStr.indexOf("/")));
    if(leapYear(parseInt(dateStr.substring(0,4))) == true)
        leap = 29;
    var tmp = parseInt(tmpStr.substring(0, tmpStr.indexOf("/")))
    if(tmpStr.substring(0, tmpStr.indexOf("/")) == '08')
        tmp = 8;
    if(tmpStr.substring(0, tmpStr.indexOf("/")) == '09')
        tmp = 9
    if(tmp < 1 || tmp > 12){
        return (false)
    }
    var monthTable = new Array(12);
    monthTable[1] = 31;
    monthTable[2] = leap;
    monthTable[3] = 31;
    monthTable[4] = 30;
    monthTable[5] = 31;
    monthTable[6] = 30;
    monthTable[7] = 31;
    monthTable[8] = 31;
    monthTable[9] = 30;
    monthTable[10] = 31;
    monthTable[11] = 30;
    monthTable[12] = 31;

    tmpStr = tmpStr.substring(tmpStr.indexOf("/")+1,tmpStr.length);
    //alert('��='+tmpStr);
    var dtmp = parseInt(tmpStr)

    if(tmpStr == '08')
        dtmp = 8;
    if(tmpStr == '09')
        dtmp = 9

    if(dtmp < 1 || dtmp > monthTable[tmp]){
        return false
    }

    return true
}

function leapYear (Year) {
        if (((Year % 4)==0) && ((Year % 100)!=0) || ((Year % 400)==0))
                return (true);
        else
                return (false);
}

//==================
//opt : 0 ������J
//opt : 1 �i����J
function checkDate(yr, mn, dt, txt, opt) {

	var sDate, ckDate;
	if (opt == 0) {
		if (trimString(yr.value) == "" ){
			alert("[" + txt + "] (�~)�椣�i�ť�");
			yr.focus();
			return false;
		}
		else {
	    	sDate = Math.abs(yr.value);
	        if (isNaN(sDate)) {
	            alert("[" + txt +"] (�~)�� ���i��J��r");
	            yr.focus();
	            return false;
	        }
	    }
		if (trimString(mn.value) == "" ){
			alert("[" + txt + "] (��)�椣�i�ť�");
			mn.focus();
			return false;
		}
		if (trimString(dt.value) == "" ){
			alert("[" + txt + "] (��)�椣�i�ť�");
			dt.focus();
			return false;
		}
	    ckDate = '' + (parseInt(yr.value) + 1911) + '/' + mn.value + '/' + dt.value;
		//alert(ckDate);
	    if (fnValidDate(ckDate) != true) {
		    		  
		    		  if(dt.value==20)
		    		  {
		    		     return true;
		    		  }
		    		  else
		    		  {   
		    	    	alert('�ҿ�J���ɶ����L�Ĥ��!');
	    	        dt.focus();
	        	    return false;
	        	  }
	   			}
	}
	else {
		if (trimString(yr.value) == "" )	{
			if (trimString(mn.value) != "" ) {
				alert("[" + txt +"](�~,��,��)�� ��������Υ�����");
	        	yr.focus();
				return false;
			}
			else {
				if (trimString(dt.value) != "" ) 	{
					alert("[" + txt + "](�~,��,��)�� ��������Υ�����");
	        		dt.focus();
					return false;
				}
			}
		}
		else {
	        sDate = Math.abs(yr.value);
	        if (isNaN(sDate)) {
	            alert("[" + txt + "](�~)�� ���i��J��r");
	        	yr.focus();
	            return false;
	        }

			if (trimString(mn.value) == "" ) {
				alert("[" + txt + "](��)�椣�i�ť�");
	        	mn.focus();
				return false;
			}
			else {
				if (trimString(dt.value) == "" ) {
					alert("[" + txt + "](��)�椣�i�ť�");
	    	    	dt.focus();
					return false;
				}
	    		ckDate = '' + (parseInt(yr.value) + 1911) + '/' + mn.value + '/' + dt.value;

		    	if (fnValidDate(ckDate) != true) {
		    		  
		    	    	alert('�ҿ�J���ɶ����L�Ĥ��!');
		    	    	
	    	        dt.focus();
	        	    return false;
	        	  
	   			}
	   		}
	   	}
	}
	return true;
}

//�b�����Walert�T��,���T�w��,�^��Y��link���link..�L�h�^�W��
function AlertMsg(form,msg,weburl_Y) {
		if(msg != ''){
		   alert(msg);		   
		   if(weburl_Y != ''){
		   	  form.action=weburl_Y;
		   	  form.submit();
		   }else{	
		     history.back();		
		   }  
	    }	    
	    return true;		
}
//�b�����@alert confirm���T��,"Y"-->URL ,"N"-->�^�W�@��/�Ψ�URL
function ConfirmMsg(form,msg,weburl_Y,weburl_N){		
		if(msg != ''){			
	    	if (confirm(msg)) {	    	    		
	    		form.action=weburl_Y;		    		
      			form.submit();
    		}else{    			
    		    if(weburl_N != ''){    		     
    		       form.action=weburl_N;		    		
      			   form.submit();
    			}else{
    				history.back();		    		    
    			}
    		}		   		
	    }	    
	    return true;	    
}	
  
//=========================================================
// Add by Winnin 2004.11.22
function checkFloatNumber(Rate1){

	Rate1.value = changeVal(Rate1);    
    if (isNaN(Math.abs(Rate1.value))) {    	    	
   	    alert("�п�J�Ʀr");
   	    return false;
    }
    return true;
}


/* �h�������w�r�ꪺ�e�B��ťզr��
 * @param   strText     �Q�n�h���e��ťզr�����r��
 * @return  string
 */
function trim(strText) {
    strText = ltrim(strText);
    strText = rtrim(strText);

    return strText;
}

/* �h�������w�r�ꪺ�����ťզr��
 * @param   strText     �Q�n�h�������ťզr�����r��
 * @return  string
 */
function ltrim(strText) {
    while (strText.substring(0,1) == ' ')
        strText = strText.substring(1, strText.length);

    return strText;
}

/* �h�������w�r�ꪺ�k���ťզr��
 * @param   strText     �Q�n�h���k���ťզr�����r��
 * @return  string
 */
function rtrim(strText) {
    while (strText.substring(strText.length-1,strText.length) == ' ')
        strText = strText.substring(0, strText.length-1);

    return strText;
}

////================== add by 2354 2005.1.19 only���
//opt : 0 ������J
//opt : 1 �i����J
function checkDateS(yr, mn, txt, opt) {

	var sDate, ckDate;
	if (opt == 0) {
		if (trimString(yr.value) == "" ){
			alert("[" + txt + "] (�~)�椣�i�ť�");
			yr.focus();
			return false;
		}
		else {
	    	sDate = Math.abs(yr.value);
	        if (isNaN(sDate)) {
	            alert("[" + txt +"] (�~)�� ���i��J��r");
	            yr.focus();
	            return false;
	        }
	    }
		if (trimString(mn.value) == "" ){
			alert("[" + txt + "] (��)�椣�i�ť�");
			mn.focus();
			return false;
		}
	    ckDate = '' + (parseInt(yr.value) + 1911) + '/' + mn.value + '/' + '01';

	    if (fnValidDate(ckDate) != true){
	    	alert('�ҿ�J���ɶ����L�Ĥ��!');
	    	
	        dt.focus();
	        return false;
	   	}
	}
	else {
		if (trimString(yr.value) == "" )	{
			if (trimString(mn.value) != "" ) {
				alert("[" + txt +"](�~,��)�� ��������");	//modify by 2354 2005.1.19
	        	yr.focus();
				return false;
			}
			else {
				if (trimString(dt.value) != "" ) 	{
					alert("[" + txt + "](�~,��)�� ��������");//modify by 2354 2005.1.19
	        		dt.focus();
					return false;
				}
			}
		}
		else {
	        sDate = Math.abs(yr.value);
	        if (isNaN(sDate)) {
	            alert("[" + txt + "](�~)�� ���i��J��r");
	        	yr.focus();
	            return false;
	        }

			if (trimString(mn.value) == "" ) {
				alert("[" + txt + "](��)�椣�i�ť�");
	        	mn.focus();
				return false;
			}
			else {
	    		ckDate = '' + (parseInt(yr.value) + 1911) + '/' + mn.value + '/' + '01';

		    	if (fnValidDate(ckDate) != true) {
		    	    alert('�ҿ�J���ɶ����L�Ĥ��!');
		    	    
	    	        dt.focus();
	        	    return false;
	   			}
	   		}
	   	}
	}
	return true;
}

function CheckMaxLength(String,Length)
{
  if (Length == null || Length == "" || String == null || String == "") return false;
  var actualLen = 0;
  for (var i=0; i<String.length; i++)
    if (String.charCodeAt(i) < 127)
      actualLen++;
    else
      actualLen+=2;

  if (actualLen <= Length)
    return false;
  else
    return true;
}

//�ˬdE-Mail�O�_���~
function CheckEmailAddress(EmailString){
      if (EmailString.length <= 5) return false;      
      if (EmailString.indexOf('@', 0) == -1){
      	  return false;
      }else{
          EmailData = EmailString.substring(EmailString.indexOf('@', 0)+1,EmailString.length);
          //alert(EmailData);
          if (EmailData.length <= 2) return false;               
          if (EmailData.indexOf('.', 0) == -1) return false;
          EmailData1 = EmailData.substring(EmailData.indexOf('.', 0)+1,EmailData.length);
          //alert(EmailData1);
          if (EmailData1.length <= 1) return false;               
          if (EmailData1.indexOf('.', 0) == -1) return false;
          EmailData2 = EmailData1.substring(EmailData1.indexOf('.', 0)+1,EmailData1.length);
          if (EmailData2.length <= 0) return false;               
          //alert(EmailData2);
      }
      if (CheckMaxLength(EmailString,50)){
  		  alert("E-Mail���׿�J���~�A�Э��s��J");
   	      return false;
      }
      return true;
}

//97,01.10 add �ˬd��J���r�����
function getLength(t1)
{
         var cnt=0;
        for(var i=0;i<t1.length;i++)
        {
            if(escape(t1.charAt(i)).length>=4) cnt+=2;
            else
            	cnt++;
        }
	return cnt;
}