function trimString(inString)
{
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




//==============================================
function AskDelete() {
  if(confirm("�T�w�n�R���ܡH")) {
    return true;
  }else {
    return false;
  }
}



function Add_Manager()
{
	if (trimString(form.FIND.value)=='N' )
	{
		alert("�D�ɸ�Ƥ��s�b,�Х���J�D�ɸ�ƫ�,�A���T�w ");
		return false;
	}
	else
	{
		form.Step.value='1';
		form.Operation.value='Add_Manager';
		form.submit();
	}
}

function goBack() {
	form.Step.value='1';
	form.Operation.value='Back';
  	form.submit();
}

function Delete_Item(form) {
  if(AskDelete()) {
    form.DeleteType.value = 'Y';
  } else {
  	form.DeleteType.value = 'N';
  	return false;
  }
  form.submit();
}

function Abdicate_Item(form)
{

	var sDate;
	var DateStr;
    var sNum;
    var DeleteType='';



	if (trimString(form.NAME.value)=="" )
	{
		alert("[�m�W] �椣�i�ť�");
		form.NAME.focus();
		return false;
	}

	if (trimString(form.RANK.value) == "" )
	{
		form.RANK.value=0;
	} else {
        if(isNaN(Math.abs(form.RANK.value)))
        {
            alert("[����] �� ���i��J��r");
            form.RANK.focus();
            return false;
        }

	}
	if (trimString(form.NEW_ID.value) =="" )
	{
		alert("[�����Ҧr��] �椣�i�ť�");
		form.NEW_ID.focus();
		return false;
	}


	if (trimString(form.BIRTH_DATE_Y.value) == "" )
	{
		if (trimString(form.BIRTH_DATE_M.value) != "" )
		{
			alert("[�X�ͦ~���](�~,��,��)�� ��������Υ�����");
        	form.BIRTH_DATE_Y.focus();
			return false;
		}
		else
		{
			if (trimString(form.BIRTH_DATE_D.value)  != "" )
			{
				alert("[�X�ͦ~���](�~,��,��)�� ��������Υ�����");
        		form.BIRTH_DATE_D.focus();
				return false;
			}
		}
	}
	else
	{
        sDate = Math.abs(form.BIRTH_DATE_Y.value);
        if(isNaN(sDate))
        {
            alert("[�X�ͦ~���](�~)�� ���i��J��r");
        	form.BIRTH_DATE_Y.focus();
            return false;
        }
		if (trimString(form.BIRTH_DATE_M.value) == "" )
		{
			alert("[�X�ͦ~���](��)�椣�i�ť�");
        	form.BIRTH_DATE_M.focus();
			return false;
		}
		else
		{
			if (trimString(form.BIRTH_DATE_D.value) == "" )
			{
				alert("[�X�ͦ~���](��)�椣�i�ť�");
    	    	form.BIRTH_DATE_D.focus();
				return false;
			}
    		ckDate = '' + (parseInt(form.BIRTH_DATE_Y.value)+1911) + '/' + form.BIRTH_DATE_M.value + '/' + form.BIRTH_DATE_D.value;

	    	if( fnValidDate(ckDate) != true){
	    	    alert('�� �� �J �� �� �� �� �L �� �� ��!!');
    	        form.BIRTH_DATE_D.focus();
        	    return false;
   			}
   		}
   	}

	if (trimString(form.INDUCT_DATE_Y.value) == "" )
	{
		alert("[�N�����] (�~)�椣�i�ť�");
		return false;
	} else {
        sDate = Math.abs(form.INDUCT_DATE_Y.value);
        if(isNaN(sDate))
        {
            alert("[�N�����] (�~)�� ���i��J��r");
            return false;
        }
    }
	if (trimString(form.INDUCT_DATE_M.value) == "" )
	{
		alert("[�N�����] (��)�椣�i�ť�");
		return false;
	}
	if (trimString(form.INDUCT_DATE_D.value) == "" )
	{
		alert("[�N�����] (��)�椣�i�ť�");
		return false;
	}
    ckDate = '' + (parseInt(form.INDUCT_DATE_Y.value)+1911) + '/' + form.INDUCT_DATE_M.value + '/' + form.INDUCT_DATE_D.value;

    if( fnValidDate(ckDate) != true){
            alert('�� �� �J �� �� �� �� �L �� �� ��!!');
            form.INDUCT_DATE_D.focus();
            return false;
   }

	if (trimString(form.ABDICATE_DATE_Y.value)  != "" )
	{
        sDate = Math.abs(form.ABDICATE_DATE_Y.value);
        if(isNaN(sDate))
        {
            alert("������� (�~)�� ���i��J��r");
            return false;
        }
		if (trimString(form.ABDICATE_DATE_M.value) == "" )
		{
			alert("[�������] (��)�椣�i�ť�");
			return false;
		}
		if (trimString(form.ABDICATE_DATE_D.value) == "" )
		{
			alert("[�������] (��)�椣�i�ť�");
			return false;
		}
    	ckDate = '' + (parseInt(form.ABDICATE_DATE_Y.value)+1911) + '/' + form.ABDICATE_DATE_M.value + '/' + form.ABDICATE_DATE_D.value;

    	if( fnValidDate(ckDate) != true){
            alert('�� �� �J �� �� �� �� �L �� �� ��!!');
            form.ABDICATE_DATE_D.focus();
            return false;
   		}
   }
   else
   {
			alert("[�������] (�~)�椣�i�ť�");
			return false;
   }

  if(AskAbdicate())
  {
    form.AbdicateType.value = 'Y';
  }
  else
  {
  	form.AbdicateType.value = 'N';
  	return false;
  }

  form.submit();
}


//==============================================
function AskAbdicate() {

	var sDate;
	var DateStr;
    var sNum;
    var DeleteType='';


	if (trimString(form.NAME.value)=="" )
	{
		alert("[�m�W] �椣�i�ť�");
		form.NAME.focus();
		return false;
	}

	if (trimString(form.RANK.value) == "" )
	{
		form.RANK.value=0;
	} else {
        if(isNaN(Math.abs(form.RANK.value)))
        {
            alert("[����] �� ���i��J��r");
            form.RANK.focus();
            return false;
        }

	}
	if (trimString(form.ID.value) =="" )
	{
		alert("[�����Ҧr��] �椣�i�ť�");
		form.ID.focus();
		return false;
	}


	if (trimString(form.BIRTH_DATE_Y.value) == "" )
	{
		if (trimString(form.BIRTH_DATE_M.value) != "" )
		{
			alert("[�X�ͦ~���](�~,��,��)�� ��������Υ�����");
        	form.BIRTH_DATE_Y.focus();
			return false;
		}
		else
		{
			if (trimString(form.BIRTH_DATE_D.value)  != "" )
			{
				alert("[�X�ͦ~���](�~,��,��)�� ��������Υ�����");
        		form.BIRTH_DATE_D.focus();
				return false;
			}
		}
	}
	else
	{
        sDate = Math.abs(form.BIRTH_DATE_Y.value);
        if(isNaN(sDate))
        {
            alert("[�X�ͦ~���](�~)�� ���i��J��r");
        	form.BIRTH_DATE_Y.focus();
            return false;
        }
		if (trimString(form.BIRTH_DATE_M.value) == "" )
		{
			alert("[�X�ͦ~���](��)�椣�i�ť�");
        	form.BIRTH_DATE_M.focus();
			return false;
		}
		else
		{
			if (trimString(form.BIRTH_DATE_D.value) == "" )
			{
				alert("[�X�ͦ~���](��)�椣�i�ť�");
    	    	form.BIRTH_DATE_D.focus();
				return false;
			}
    		ckDate = '' + (parseInt(form.BIRTH_DATE_Y.value)+1911) + '/' + form.BIRTH_DATE_M.value + '/' + form.BIRTH_DATE_D.value;

	    	if( fnValidDate(ckDate) != true){
	    	    alert('�� �� �J �� �� �� �� �L �� �� ��!!');
    	        form.BIRTH_DATE_D.focus();
        	    return false;
   			}
   		}
   	}

	if (trimString(form.INDUCT_DATE_Y.value) == "" )
	{
		alert("[�N�����] (�~)�椣�i�ť�");
		return false;
	} else {
        sDate = Math.abs(form.INDUCT_DATE_Y.value);
        if(isNaN(sDate))
        {
            alert("[�N�����] (�~)�� ���i��J��r");
            return false;
        }
    }
	if (trimString(form.INDUCT_DATE_M.value) == "" )
	{
		alert("[�N�����] (��)�椣�i�ť�");
		return false;
	}
	if (trimString(form.INDUCT_DATE_D.value) == "" )
	{
		alert("[�N�����] (��)�椣�i�ť�");
		return false;
	}
    ckDate = '' + (parseInt(form.INDUCT_DATE_Y.value)+1911) + '/' + form.INDUCT_DATE_M.value + '/' + form.INDUCT_DATE_D.value;

    if( fnValidDate(ckDate) != true){
            alert('�� �� �J �� �� �� �� �L �� �� ��!!');
            form.INDUCT_DATE_D.focus();
            return false;
   }

  if(confirm("�T�w�n�����ܡH")) {
    return true;
  }else {
    return false;
  }
}

function Delete_Business(form) {
  if(AskDelete()) {
    form.DeleteType.value = 'Y';
  } else {
  	form.DeleteType.value = 'N';
  	return false;
  }
  form.submit();
}


function CheckInput(form)
{
	confirm("�T�w�R�� ??");
    form.submit();
}

function Ask()
{
	OrdString="�T�w�R�� ??"
	return OrdString;
}

function fnValidDate(dateStr){
    var leap = 28;
    if(leapYear(parseInt(dateStr.substring(0,4))) == true)
        leap = 29;
    var tmp = parseInt(dateStr.substring(5, 7))
    if(dateStr.substring(5, 7) == '08')
        tmp = 8;
    if(dateStr.substring(5, 7) == '09')
        tmp = 9
    if(tmp < 1 || tmp > 12){
        alert("aaa");
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

    var dtmp = parseInt(dateStr.substring(8))

    if(dateStr.substring(8) == '08')
        dtmp = 8;
    if(dateStr.substring(8) == '09')
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
//=====================================================================
function changeVal(T1)
{
    pos=0
    var oring=T1.value
    pos=oring.indexOf(",")

    while (pos !=-1)
    {
        oring=(oring).replace(",","")
        pos=oring.indexOf(",")
    }
    return oring;

}

//================================================================
function changeStr(T1)
{
    c="";
    var oring = T1.value
    var t1v1  = T1.value
    var t1v2  = "";
   //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    if (isNaN(Math.abs(T1.value)))
    {
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


//===========================================================================
function checkRate(Rate1)
{
    if(isNaN(Math.abs(Rate1.value)))
    {
   	    alert("�п�J�Ʀr");
   	    return false;
    }
    if(eval(Rate1.value) > 99.99 || eval(Rate1.value) < -99.99)
    {
        alert("�Q�v���i�j�� 99.99, �]���i�p�� -99.99")
        return Rate1.value;
    }
    if((Rate1.value).indexOf(".") != -1 )
    {
        len = (Rate1.value).substring((Rate1.value).indexOf(".")+1,Rate1.value.length);
        if(len.length > 2)
        {
            alert("�p���I��u�঳��Ӧ��");

            return Rate1.value;
        }
    }
    return Rate1.value
}



function  Check_CC001W_Show_Data(form)
{

	if (trimString(form.SETUP_DATE_Y.value) == "" )
	{
		if (trimString(form.SETUP_DATE_M.value) != "" )
		{
			alert("[�]�F����l�֭�]�ߤ��](�~,��,��)�� ��������Υ�����");
        	form.SETUP_DATE_Y.focus();
			return false;
		}
		else
		{
			if (trimString(form.SETUP_DATE_D.value) != "" )
			{
				alert("[�]�F����l�֭�]�ߤ��](�~,��,��)�� ��������Υ�����");
        		form.SETUP_DATE_D.focus();
				return false;
			}
		}
	}
	else
	{
        sDate = Math.abs(form.SETUP_DATE_Y.value);
        if(isNaN(sDate))
        {
            alert("[�]�F����l�֭�]�ߤ��](�~)�� ���i��J��r");
        	form.SETUP_DATE_Y.focus();
            return false;
        }

		if (trimString (form.SETUP_DATE_M.value) == "" )
		{
			alert("[�]�F����l�֭�]�ߤ��](��)�椣�i�ť�");
        	form.SETUP_DATE_M.focus();
			return false;
		}
		else
		{
			if (trimString(form.SETUP_DATE_D.value) == "" )
			{
				alert("[�]�F����l�֭�]�ߤ��](��)�椣�i�ť�");
    	    	form.SETUP_DATE_D.focus();
				return false;
			}
    		ckDate = '' + (parseInt(form.SETUP_DATE_Y.value)+1911) + '/' + form.SETUP_DATE_M.value + '/' + form.SETUP_DATE_D.value;

	    	if( fnValidDate(ckDate) != true){
	    	    alert('�� �� �J �� �� �� �� �L �� �� ��!!');
    	        form.SETUP_DATE_D.focus();
        	    return false;
   			}
   		}
   	}

	if (trimString(form.CHG_LICENSE_DATE_Y.value) =="" )
	{
		if (trimString(form.CHG_LICENSE_DATE_M.value) != "" )
		{
			alert("[�̪񴫷Ӥ��](�~,��,��)�� ��������Υ�����");
        	form.CHG_LICENSE_DATE_Y.focus();
			return false;
		}
		else
		{
			if (trimString(form.CHG_LICENSE_DATE_D.value) != "" )
			{
				alert("[�̪񴫷Ӥ��](�~,��,��)�� ��������Υ�����");
        		form.CHG_LICENSE_DATE_D.focus();
				return false;
			}
		}
	}
	else
	{
        sDate = Math.abs(form.CHG_LICENSE_DATE_Y.value);
        if(isNaN(sDate))
        {
            alert("[�̪񴫷Ӥ��](�~)�� ���i��J��r");
        	form.CHG_LICENSE_DATE_Y.focus();
            return false;
        }
		if (trimString(form.CHG_LICENSE_DATE_M.value) == "" )
		{
			alert("[�̪񴫷Ӥ��](��)�椣�i�ť�");
        	form.CHG_LICENSE_DATE_M.focus();
			return false;
		}
		else
		{
			if (trimString(form.CHG_LICENSE_DATE_D.value) == "" )
			{
				alert("[�̪񴫷Ӥ��](��)�椣�i�ť�");
    	    	form.CHG_LICENSE_DATE_D.focus();
				return false;
			}
    		ckDate = '' + (parseInt(form.CHG_LICENSE_DATE_Y.value)+1911) + '/' + form.CHG_LICENSE_DATE_M.value + '/' + form.CHG_LICENSE_DATE_D.value;

	    	if( fnValidDate(ckDate) != true){
	    	    alert('�� �� �J �� �� �� �� �L �� �� ��!!');
    	        form.CHG_LICENSE_DATE_D.focus();
        	    return false;
   			}
   		}
   	}
   	//serissa
   	if(!CheckMyDate(form.open_yy,form.open_mm,form.open_dd,form.open_date,"��l�}�~���"))
      return false;


	if (trimString(form.START_DATE_Y.value) == "" )
	{
		if (trimString(form.START_DATE_M.value) != "" )
		{
			alert("[�}�l��~��](�~,��,��)�� ��������Υ�����");
        	form.START_DATE_Y.focus();
			return false;
		}
		else
		{
			if (trimString(form.START_DATE_D.value) != "" )
			{
				alert("[�}�l��~��](�~,��,��)�� ��������Υ�����");
        		form.START_DATE_D.focus();
				return false;
			}
		}
	}
	else
	{
        sDate = Math.abs(form.START_DATE_Y.value);
        if(isNaN(sDate))
        {
            alert("[�}�l��~��](�~)�� ���i��J��r");
        	form.START_DATE_Y.focus();
            return false;
        }
		if (trimString(form.START_DATE_M.value) == "" )
		{
			alert("[�}�l��~��](��)�椣�i�ť�");
        	form.START_DATE_M.focus();
			return false;
		}
		else
		{
			if (trimString(form.START_DATE_D.value) == "" )
			{
				alert("[�}�l��~��](��)�椣�i�ť�");
    	    	form.START_DATE_D.focus();
				return false;
			}
    		ckDate = '' + (parseInt(form.START_DATE_Y.value)+1911) + '/' + form.START_DATE_M.value + '/' + form.START_DATE_D.value;

	    	if( fnValidDate(ckDate) != true){
	    	    alert('�� �� �J �� �� �� �� �L �� �� ��!!');
    	        form.START_DATE_D.focus();
        	    return false;
   			}
   		}
   	}

	camount = Math.abs(changeVal(form.CAPITAL_AMOUNT));
    if(isNaN(camount))
    {
    	alert("[�ꦬ�ꥻ�`�B]�� ���i��J��r");
        form.CAPITAL_AMOUNT.focus();
        return false;
    }

	if (trimString(form.ADDR.value) == "" )
	{
		alert("[�`���c�a�}] ���i�ť�");
    	form.ADDR.focus();
		return false;
   	}

    if(trimString(form.MAINTAIN_DEPT.value) == "" )
    {
		alert("[���@�����W��] ���i�ť�");
        form.MAINTAIN_DEPT.focus();
		return false;
	}
    if(trimString(form.DIRECTOR_NAME.value) == "" )
    {
		alert("[�D�ީm�W] ���i�ť�");
        form.DIRECTOR_NAME.focus();
		return false;
	}
    if(trimString(form.DIRECTOR_TEL.value) == "" )
    {
		alert("[�D�޹q��] ���i�ť�");
        form.DIRECTOR_TEL.focus();
		return false;
	}

    if(trimString(form.MAINTAIN_TEL.value) == "" )
    {
		alert("[�ӿ���p���q��] ���i�ť�");
        form.MAINTAIN_TEL.focus();
		return false;
	}

}

function  Check_CC001W_Manager_ShowInsert(form)
{
	var sDate;
	var DateStr;
    var sNum;
    var DeleteType='';



	if (trimString(form.NAME.value)=="" )
	{
		alert("[�m�W] �椣�i�ť�");
		form.NAME.focus();
		return false;
	}

	if (trimString(form.RANK.value) == "" )
	{
		form.RANK.value=0;
	} else {
        if(isNaN(Math.abs(form.RANK.value)))
        {
            alert("[����] �� ���i��J��r");
            form.RANK.focus();
            return false;
        }

	}
	if (trimString(form.ID.value) =="" )
	{
		alert("[�����Ҧr��] �椣�i�ť�");
		form.ID.focus();
		return false;
	}


	if (trimString(form.BIRTH_DATE_Y.value) == "" )
	{
		if (trimString(form.BIRTH_DATE_M.value) != "" )
		{
			alert("[�X�ͦ~���](�~,��,��)�� ��������Υ�����");
        	form.BIRTH_DATE_Y.focus();
			return false;
		}
		else
		{
			if (trimString(form.BIRTH_DATE_D.value)  != "" )
			{
				alert("[�X�ͦ~���](�~,��,��)�� ��������Υ�����");
        		form.BIRTH_DATE_D.focus();
				return false;
			}
		}
	}
	else
	{
        sDate = Math.abs(form.BIRTH_DATE_Y.value);
        if(isNaN(sDate))
        {
            alert("[�X�ͦ~���](�~)�� ���i��J��r");
        	form.BIRTH_DATE_Y.focus();
            return false;
        }
		if (trimString(form.BIRTH_DATE_M.value) == "" )
		{
			alert("[�X�ͦ~���](��)�椣�i�ť�");
        	form.BIRTH_DATE_M.focus();
			return false;
		}
		else
		{
			if (trimString(form.BIRTH_DATE_D.value) == "" )
			{
				alert("[�X�ͦ~���](��)�椣�i�ť�");
    	    	form.BIRTH_DATE_D.focus();
				return false;
			}
    		ckDate = '' + (parseInt(form.BIRTH_DATE_Y.value)+1911) + '/' + form.BIRTH_DATE_M.value + '/' + form.BIRTH_DATE_D.value;

	    	if( fnValidDate(ckDate) != true){
	    	    alert('�� �� �J �� �� �� �� �L �� �� ��!!');
    	        form.BIRTH_DATE_D.focus();
        	    return false;
   			}
   		}
   	}

	if (trimString(form.INDUCT_DATE_Y.value) == "" )
	{
		alert("[�N�����] (�~)�椣�i�ť�");
		return false;
	} else {
        sDate = Math.abs(form.INDUCT_DATE_Y.value);
        if(isNaN(sDate))
        {
            alert("[�N�����] (�~)�� ���i��J��r");
            return false;
        }
    }
	if (trimString(form.INDUCT_DATE_M.value) == "" )
	{
		alert("[�N�����] (��)�椣�i�ť�");
		return false;
	}
	if (trimString(form.INDUCT_DATE_D.value) == "" )
	{
		alert("[�N�����] (��)�椣�i�ť�");
		return false;
	}
    ckDate = '' + (parseInt(form.INDUCT_DATE_Y.value)+1911) + '/' + form.INDUCT_DATE_M.value + '/' + form.INDUCT_DATE_D.value;

    if( fnValidDate(ckDate) != true){
            alert('�� �� �J �� �� �� �� �L �� �� ��!!');
            form.INDUCT_DATE_D.focus();
            return false;
   }
   //serissa
   if(trimString(form.BACKGROUND.value).length>50)
    {
         alert("��J���̪�g�����i�W�L50�Ӧr");
         return false;
    }

    if(trimString(form.SPECIALITY.value).length>40)
    {
         alert("��J���M�����i�W�L40�Ӧr");
         return false;
    }
    if(trimString(form.INCHARGE.value).length>50)
    {
         alert("��J�����ɷ~�ȶ��ؤ��i�W�L50�Ӧr");
         return false;
    }

}




function  Check_CC001W_Manager_ShowUpdate(form)
{
	var sDate;
	var DateStr;
    var sNum;
    var DeleteType='';



	if (trimString(form.NAME.value)=="" )
	{
		alert("[�m�W] �椣�i�ť�");
		form.NAME.focus();
		return false;
	}

	if (trimString(form.RANK.value) == "" )
	{
		form.RANK.value=0;
	} else {
        if(isNaN(Math.abs(form.RANK.value)))
        {
            alert("[����] �� ���i��J��r");
            form.RANK.focus();
            return false;
        }

	}
	if (trimString(form.NEW_ID.value) =="" )
	{
		alert("[�����Ҧr��] �椣�i�ť�");
		form.NEW_ID.focus();
		return false;
	}




	if (trimString(form.BIRTH_DATE_Y.value) == "" )
	{
		if (trimString(form.BIRTH_DATE_M.value) != "" )
		{
			alert("[�X�ͦ~���](�~,��,��)�� ��������Υ�����");
        	form.BIRTH_DATE_Y.focus();
			return false;
		}
		else
		{
			if (trimString(form.BIRTH_DATE_D.value)  != "" )
			{
				alert("[�X�ͦ~���](�~,��,��)�� ��������Υ�����");
        		form.BIRTH_DATE_D.focus();
				return false;
			}
		}
	}
	else
	{
        sDate = Math.abs(form.BIRTH_DATE_Y.value);
        if(isNaN(sDate))
        {
            alert("[�X�ͦ~���](�~)�� ���i��J��r");
        	form.BIRTH_DATE_Y.focus();
            return false;
        }
		if (trimString(form.BIRTH_DATE_M.value) == "" )
		{
			alert("[�X�ͦ~���](��)�椣�i�ť�");
        	form.BIRTH_DATE_M.focus();
			return false;
		}
		else
		{
			if (trimString(form.BIRTH_DATE_D.value) == "" )
			{
				alert("[�X�ͦ~���](��)�椣�i�ť�");
    	    	form.BIRTH_DATE_D.focus();
				return false;
			}
    		ckDate = '' + (parseInt(form.BIRTH_DATE_Y.value)+1911) + '/' + form.BIRTH_DATE_M.value + '/' + form.BIRTH_DATE_D.value;

	    	if( fnValidDate(ckDate) != true){
	    	    alert('�� �� �J �� �� �� �� �L �� �� ��!!');
    	        form.BIRTH_DATE_D.focus();
        	    return false;
   			}
   		}
   	}

	if (trimString(form.INDUCT_DATE_Y.value) == "" )
	{
		alert("[�N�����] (�~)�椣�i�ť�");
		return false;
	} else {
        sDate = Math.abs(form.INDUCT_DATE_Y.value);
        if(isNaN(sDate))
        {
            alert("[�N�����] (�~)�� ���i��J��r");
            return false;
        }
    }
	if (trimString(form.INDUCT_DATE_M.value) == "" )
	{
		alert("[�N�����] (��)�椣�i�ť�");
		return false;
	}
	if (trimString(form.INDUCT_DATE_D.value) == "" )
	{
		alert("[�N�����] (��)�椣�i�ť�");
		return false;
	}
    ckDate = '' + (parseInt(form.INDUCT_DATE_Y.value)+1911) + '/' + form.INDUCT_DATE_M.value + '/' + form.INDUCT_DATE_D.value;

    if( fnValidDate(ckDate) != true){
            alert('�� �� �J �� �� �� �� �L �� �� ��!!');
            form.INDUCT_DATE_D.focus();
            return false;
    }

    //serissa
    if(trimString(form.BACKGROUND.value).length>50)
    {
         alert("��J���̪�g�����i�W�L50�Ӧr");
         return false;
    }

    if(trimString(form.SPECIALITY.value).length>40)
    {
         alert("��J���M�����i�W�L40�Ӧr");
         return false;
    }

    if(trimString(form.INCHARGE.value).length>50)
    {
         alert("��J�����ɷ~�ȶ��ؤ��i�W�L50�Ӧr");
         return false;
    }
}

//=====================================================================================================================
function CheckMyDate(checkyear,checkmonth,checkday,checkdate,szWarning)
{
    if(!CheckYear(checkyear))
            return false;
    if(trimString(checkyear.value)!="" && trimString(checkmonth.value)!="" && trimString(checkday.value)!="")
    {

        var setyear = (Math.abs(checkyear.value)) + 1911;
        if(!fnValidDate1(setyear + checkmonth.value + checkday.value))
        {
            alert(szWarning+"���~");
            return false;
        }
        checkdate.value = setyear + "/"+checkmonth.value + "/"+checkday.value;
    }
    else if(trimString(checkyear.value)=="" && trimString(checkmonth.value)=="" && trimString(checkday.value)==""){}
         else
         {
            alert("�п�J���㪺"+ szWarning+ "�Ϊ̳�����J");
            return false;
         }
    return true
}

//==============================
function CheckYear(cyear)
{
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
//==============================================
function fnValidDate1(dateStr)
{
    var leap = 28;
    if (leapYear(parseInt(dateStr.substring(0,4))) == 1)
        leap = 29;
    var tmp = parseInt(dateStr.substring(4,6))
    if (dateStr.substring(4,6) == '08')
        tmp = 8;
    if (dateStr.substring(4,6) == '09')
        tmp = 9

    if (tmp < 1 || tmp > 12)
    {
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

    var dtmp = parseInt(dateStr.substring(6))
    if(dateStr.substring(6) == '08')
        dtmp = 8;
    if(dateStr.substring(6) == '09')
        dtmp = 9

    if(dtmp < 1 || dtmp > monthTable[tmp])
    {
        return (false)
    }
    return (true)
}