function doSubmit(cnd){  
  form.rtf_bank_type.value=  form.bankType.value;
  
  if(CheckYear(form.S_YEAR)){
    if((this.document.forms[0].tbank.value!='')){    	
       	this.document.forms[0].action = "/pages/FR0066W_Rtf.jsp";
   	this.document.forms[0].target = '_self';
   	this.document.forms[0].submit();

    }else{
  	alert("請選擇一家農漁會");
    }
  }
}

function checkCity() {
  if(document.getElementById("bankType").value == "6" ||
     document.getElementById("bankType").value == "7" ) {
    document.getElementById("cityType").disabled = false;
  } else {
    document.getElementById("cityType").disabled = true;
  }
}

function CheckYear(cyear) {
    if(cyear.value.indexOf(".") != -1)
    {
        alert("年份不可為小數");
        return false;
    }
    if(isNaN(Math.abs(cyear.value)))
    {
        alert("年份不可為文字");
        return false;
    }
    if(trimString(cyear.value) != '')
        cyear.value = Math.abs(cyear.value);
    return true;
}
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

function resetOption() {
    var s1 = document.getElementById("tbank");    
    s1.length = 0;  

    oOption = document.createElement("OPTION");
	oOption.text="全部";
  	oOption.value="";
  	s1.add(oOption); 	
}
