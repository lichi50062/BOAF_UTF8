function CheckMyDate(checkyear,checkmonth,checkday,checkdate,szWarning)
{
    if(!CheckYear(checkyear))
            return false;
    if(trimString(checkyear.value)!="" && trimString(checkmonth.value)!="" && trimString(checkday.value)!="")
    {
        var setyear = (Math.abs(checkyear.value)) + 1911;
        if(!fnValidDate(setyear + checkmonth.value + checkday.value))
        {
            alert(szWarning+"���~");
            return false;
        }
        checkdate.value = setyear + "/"+checkmonth.value + "/"+checkday.value;
//        checkdate.value = setyear + checkmonth.value +checkday.value;
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

//================================================================
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
function fnValidDate(dateStr)
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
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function leapYear (Year)
{
    if (((Year % 4)==0) && ((Year % 100)!=0) || ((Year % 400)==0))
        return (1);
    else
	return (0);
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
//================================================================
function changeStr1(T1)
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
  //===============================================
    if(eval(T1.value) > 99999999999999.99 || eval(T1.value) < -99999999999999.99)
    {
        alert("���B���i�j�� 99999999999999.99, �]���i�p�� -99999999999999.99")
        return oring;
    }
    if((T1.value).indexOf(".") != -1 )
    {
        len = (T1.value).substring((T1.value).indexOf(".")+1,T1.value.length);
        if(len.length > 2)
        {
            alert("�p���I��u�঳��Ӧ��");
            return oring;
        }
        T1.value = eval(T1.value);
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
//================================================================
//================================================================
function changeStr9(T1,code)
{
    c="";
    var oring = changeVal(T1);
    var t1v1  = changeVal(T1);
    var t1v2  = "";
   //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
       if (isNaN(Math.abs(oring)))
    {
   	    alert("�п�J�Ʀr");
   	    return false;
    }else{
        if (eval(oring) < 0 )
            c="-";
    }
    oring= Math.abs(oring);
    t1v1  = oring

    if (code=="1") {
  //===============================================
    if(eval(oring) > 99999999999999.99 || eval(oring) < -99999999999999.99)
    {
        alert("���B���i�j�� 99999999999999.99, �]���i�p�� -99999999999999.99")
        return false;
    }
    }
    if (code=="3") {
  //===============================================
    if(eval(oring) > 99999999.99 || eval(oring) < -99999999.99)
    {
        alert("���B���i�j�� 99999999.99, �]���i�p�� -99999999.99")
        return false;
    }
    }
    if (code=="2") {
  //===============================================
    if(eval(oring) > 999999.99 || eval(oring) < -999999.99)
    {
        alert("���B���i�j�� 999999.99, �]���i�p�� -999999.99")
        return false;
    }
    }
	if (code=="4") {
  //===============================================
    if(eval(oring) > 999.99 || eval(oring) < -999.99)
    {
        alert("���B���i�j�� 999.99, �]���i�p�� -999.99")
        return false;
    }
    }
    if((T1.value).indexOf(".") != -1 )
    {
        len = (T1.value).substring((T1.value).indexOf(".")+1,T1.value.length);
        if(len.length > 2)
        {
            alert("�p���I��u�঳��Ӧ��");
            return false;
        }
     }
    return true;
}
//================================================================
function changeStr2(T1)
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
  //===============================================
    if(eval(T1.value) > 999999.99 || eval(T1.value) < -999999.99)
    {
        alert("���B���i�j�� 999999.99, �]���i�p�� -999999.99")
        return oring;
    }
    if((T1.value).indexOf(".") != -1 )
    {
        len = (T1.value).substring((T1.value).indexOf(".")+1,T1.value.length);
        if(len.length > 2)
        {
            alert("�p���I��u�঳��Ӧ��");
            return oring;
        }
        T1.value = eval(T1.value);
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
//================================================================
function changeStr3(T1)
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
  //===============================================
    if(eval(T1.value) > 99999999.99 )
    {
        alert("���B���i�j�� 99999999.99")
        return oring;
    }
    if((T1.value).indexOf(".") != -1 )
    {
        len = (T1.value).substring((T1.value).indexOf(".")+1,T1.value.length);
        if(len.length > 2)
        {
            alert("�p���I��u�঳��Ӧ��");
            return oring;
        }
        T1.value = eval(T1.value);
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
//=============================================================

function changeVal(T1)
{
    pos=0;
    var oring=T1.value;
    pos=oring.indexOf(",");
    while (pos !=-1)
    {
        oring=(oring).replace(",","");
        pos=oring.indexOf(",");
    }

    return oring;
}
//====================================================
function ChgAction(form,actionURL)
{
    form.action = actionURL;
	form.submit();
}
//===============================================================
//**�P�_��J��ƬO�_����99.99��(-99.99)
//===========================================================================
function NewcheckRate1(Rate1)
{
    if(isNaN(Math.abs(Rate1.value)))
    {
   	    alert("�п�J�Ʀr");
   	    return false;
    }
    if(eval(Rate1.value) > 99.99 || eval(Rate1.value) < -99.99)
    {
        alert("�Q�v���i�j�� 99.99, �]���i�p�� -99.99")
        return false;
    }
    if((Rate1.value).indexOf(".") != -1 )
    {
        len = (Rate1.value).substring((Rate1.value).indexOf(".")+1,Rate1.value.length);
        if(len.length > 2)
        {
            alert("�p���I��u�঳��Ӧ��");
            return false;
        }
        Rate1.value = eval(Rate1.value);
    }
    return true;
}
//===========================================================================
function NewcheckRate(Rate1)
{
    if(isNaN(Math.abs(Rate1.value)))
    {
   	    alert("�п�J�Ʀr");
   	    return false;
    }
    if(eval(Rate1.value) > 99.99 || eval(Rate1.value) < 0)
    {
        alert("�Q�v���i�j�� 99.99, �]���i�p�� 0")
        return false;
    }
    if((Rate1.value).indexOf(".") != -1 )
    {
        len = (Rate1.value).substring((Rate1.value).indexOf(".")+1,Rate1.value.length);
        if(len.length > 2)
        {
            alert("�p���I��u�঳��Ӧ��");
            return false;
        }
        Rate1.value = eval(Rate1.value);
    }
    return true;
}
//===========================================================================
function NewcheckRate2(Rate1)
{
    if(isNaN(Math.abs(Rate1.value)))
    {
   	    alert("�п�J�Ʀr");
   	    return false;
    }
    if(eval(Rate1.value) > 100 || eval(Rate1.value) < 0)
    {
        alert("�Q�v���i�j�� 100, �]���i�p�� 0")
        return false;
    }
    if((Rate1.value).indexOf(".") != -1 )
    {
        len = (Rate1.value).substring((Rate1.value).indexOf(".")+1,Rate1.value.length);
        if(len.length > 2)
        {
            alert("�p���I��u�঳��Ӧ��");
            return false;
        }
        Rate1.value = eval(Rate1.value);
    }
    return true;
}
// ========== Serissa 2001/11/05 ====================
function NewcheckRate3(Rate1)
{
    if(isNaN(Math.abs(Rate1.value)))
    {
   	    alert("�п�J�Ʀr");
   	    return false;
    }
    if(eval(Rate1.value) > 99.999 || eval(Rate1.value) < -99.999)
    {
        alert("�Q�v���i�j�� 99.999, �]���i�p�� -99.999")
        return false;
    }
    if((Rate1.value).indexOf(".") != -1 )
    {
        len = (Rate1.value).substring((Rate1.value).indexOf(".")+1,Rate1.value.length);
        if(len.length > 3)
        {
            alert("�p���I��u�঳�T�Ӧ��");
            return false;
        }
        Rate1.value = eval(Rate1.value);
    }
    return true;
}
// ========== Serissa End 2001/11/05 ====================

//anson 2004/08/17
function NewcheckRate4(Rate1)
{
    if(isNaN(Math.abs(Rate1.value)))
    {
   	    alert("�п�J�Ʀr");
   	    return false;
    }
    if(eval(Rate1.value) > 9999.99 || eval(Rate1.value) < -9999.99)
    {
        alert("�Q�v���i�j�� 9999.99, �]���i�p�� -9999.99")
        return false;
    }
    if((Rate1.value).indexOf(".") != -1 )
    {
        len = (Rate1.value).substring((Rate1.value).indexOf(".")+1,Rate1.value.length);
        if(len.length > 2)
        {
            alert("�p���I��u�঳��Ӧ��");
            return false;
        }
        Rate1.value = eval(Rate1.value);
    }
    return true;
}

function changeSelect(form)
{

    if (form.way.value==2)
	form.way_name.disabled=true;
     else
        form.way_name.disabled=false;
}
//=============================================================================

function AskDelete() {
  if(confirm("�T�w�n�R���ܡH")) {
    return true;
  }else {
    return false;
  }
}

function DeleteAction(form1) {
  if(AskDelete()) {
    form1.Function.value = 'delete';
    form1.submit();
  }
}
//===============�z�ʨƨ����P�_��=============//
function MoveAction(form1) {

    if (form1.abdicate_date_yy.value=='' || form1.abdicate_date_mm.value=='' || form1.abdicate_date_dd.value==''){
       alert("�п�J�������");
       return false;
    }
   if(!CheckMyDate(form1.abdicate_date_yy,form1.abdicate_date_mm,form1.abdicate_date_dd,form1.abdicate_date,"�������"))
       return false;

   form1.Function.value = 'delete';
   //form1.Function.value = 'move';
   form1.submit();
}

function checkNum(Rate1)
{
    if(isNaN(Math.abs(Rate1.value)))
    {
   	    alert("�п�J�Ʀr");
   	    return false;
    }
    return true;
}
function checkNum1(T1)
{
    pos=0;
    var oring=T1.value;
    pos=oring.indexOf(",");
    while (pos !=-1)
    {
        oring=(oring).replace(",","");
        pos=oring.indexOf(",");
    }

    if(isNaN(Math.abs(oring)))
    {
   	    alert("�п�J�Ʀr");
   	    return false;
    }
    return true;
}
//
function doOY004WSubmit(form,myfun,code)
{
    form.Function.value = myfun;
    if(myfun=="delete")
    {
        if(AskDelete())
            form.submit();
    }
    else
    {
    	if (code=="1") { //*�����֭�
           if(Check_OY004W_ShowSub1(form))
              form.submit();
        }
    	if (code=="2") {  //**�����D��
           if(Check_OY004W_ShowSubM(form))
              form.submit();
        }
    	if (code=="3") {  //**���ʨ�
           if(Check_OY004W_ShowSubM1(form))
              form.submit();
        }
    	if (code=="4") {  //**��������Ʒ~
           if(Check_OY004W_ShowSub2(form))
              form.submit();
        }

    }
}
//
function doOX004WSubmit(form,myfun,code)
{
    form.Function.value = myfun;
    if(myfun=="delete")
    {
        if(AskDelete())
            form.submit();
    }
    else
    {
    	if (code=="1") { //*�����֭�
           if(Check_OX004W_ShowSub1(form))
              form.submit();
        }
    	if (code=="2") {  //**�����D��
           if(Check_OX004W_ShowSubM(form))
              form.submit();
        }
    	if (code=="3") {  //**���ʨ�
           if(Check_OX004W_ShowSubM1(form))
              form.submit();
        }
    	if (code=="4") {  //**��������Ʒ~
           if(Check_OX004W_ShowSub2(form))
              form.submit();
        }

    }
}
//
function doIX016WSubmit(form,myfun,code)
{
    form.Function.value = myfun;
    if(myfun=="delete")
    {
        if(AskDelete())
            form.submit();
    }
    else
    {
    	if (code=="6") { //*�ѶU�Ȧ�
           if(Check_IX016W_ShowSub6(form))
              form.submit();
        }
    	if (code=="5") {  //**��z���Ω���
           if(Check_IX016W_ShowSub5(form))
              form.submit();
        }
    	if (code=="4") {  //**����ӷ���
           if(Check_IX016W_ShowSub4(form))
              form.submit();
        }
    	if (code=="3") {  //**�U�ڹ�H
           if(Check_IX016W_ShowSub3(form))
              form.submit();
        }
    	if (code=="2") {  //**�U�ڽd��
           if(Check_IX016W_ShowSub2(form))
              form.submit();
        }

    }
}
//==============================�H�X�����ʨư򥻸�ƺ��@�e��CHECK=============
function LinkToInsert(form,linkaddr,code)
{
    linkaddr = "/servlet/" + linkaddr;
    if(form.Function.value=='Add')
    {
    	if (code==1) {
           alert("�Х��s�W����Ʒ~�򥻸��");
           return;
        }
      	if (code==2) {
           alert("�Х��s�W�F���ʶU�ڲέp���");
        return;
        }
      	if (code==3) {
           alert("�Х��s�W�~��Ȧ�b�ؤ�����c�򥻸��");
        return;
        }
        if (code==4) {
           alert("�Х��s�W�~��Ȧ�b�ؿ�ƳB���");
           return;
        }
        //serissa
        if (code==5) {
            if (form.period_year.value == ""){
                alert("�Х��s�W�~�ר���B�׸��");
                return;
            } else {
                window.open(linkaddr,'_self');
            }
        }
    }
    else {
        window.open(linkaddr,'_self');
    }
}

//==============================�H�X�����ʨư򥻸�ƺ��@�e��CHECK=============
function Check_BX003W(form1)
{
    if (form1.id.value=='' || form1.position_code.value=='' || form1.name.value==''){
       alert("������.¾��.�m�W���i�ť�");
       return false;
    }
    if (form1.induct_date_yy.value=='' || form1.induct_date_mm.value=='' || form1.induct_date_dd.value==''){
       alert("�N��������i�ť�");
       return false;
    }

    if(!CheckMyDate(form1.birth_yy,form1.birth_mm,form1.birth_dd,form1.birth_date,"�X�ͤ��"))
            return false;
    if(!CheckMyDate(form1.induct_date_yy,form1.induct_date_mm,form1.induct_date_dd,form1.induct_date,"�N�����"))
            return false;
    if(!CheckMyDate(form1.period_start_yy,form1.period_start_mm,form1.period_start_dd,form1.period_start,"���������_��"))
            return false;
    if(!CheckMyDate(form1.period_end_yy,form1.period_end_mm,form1.period_end_dd,form1.period_end,"������������"))
            return false;
    if(!checkNum1(form1.own_stock_amount))
       return false;
    if(!NewcheckRate(form1.own_stock_rate))
       return false;
    return true;

}

//==============================�H�X�����Ѫ��B�e�Q�j�����򥻸�ƺ��@CHECK=============
function Check_BX004W(form1)
{
    if (form1.id.value=='' || form1.g_or_c.value=='' || form1.name.value==''){
       alert("������.�m�W.�����������i�ť�");
       return false;
    }

    return true;

}

//==============================�A���|���ʨư򥻸�ƺ��@�e��CHECK=============
function Check_FX003W(form1)
{
    if (form1.id.value=='' || form1.position_code.value=='' || form1.name.value==''){
       alert("������.¾��.�m�W���i�ť�");
       return false;
    }
    if (form1.induct_date_yy.value=='' || form1.induct_date_mm.value=='' || form1.induct_date_dd.value==''){
       alert("�N��������i�ť�");
       return false;
    }

    if(!CheckMyDate(form1.birth_yy,form1.birth_mm,form1.birth_dd,form1.birth_date,"�X�ͤ��"))
            return false;
    if(!CheckMyDate(form1.induct_date_yy,form1.induct_date_mm,form1.induct_date_dd,form1.induct_date,"�N�����"))
            return false;
    if(!CheckMyDate(form1.period_start_yy,form1.period_start_mm,form1.period_start_dd,form1.period_start,"���������_��"))
            return false;
    if(!CheckMyDate(form1.period_end_yy,form1.period_end_mm,form1.period_end_dd,form1.period_end,"������������"))
            return false;
    return true;

}
//==============================���ʨư򥻸�ƺ��@�e��CHECK=============
function Check_OY003W(form1)
{
    if (form1.id.value=='' || form1.position_code.value=='' || form1.name.value==''){
       alert("������.¾��.�m�W���i�ť�");
       return false;
    }
    if (form1.induct_date_yy.value=='' || form1.induct_date_mm.value=='' || form1.induct_date_dd.value==''){
       alert("�N��������i�ť�");
       return false;
    }

    if(!CheckMyDate(form1.birth_yy,form1.birth_mm,form1.birth_dd,form1.birth_date,"�X�ͤ��"))
            return false;
    if(!CheckMyDate(form1.induct_date_yy,form1.induct_date_mm,form1.induct_date_dd,form1.induct_date,"�N�����"))
            return false;
    if(!checkNum1(form1.appointed_num))
       return false;
    if(!CheckMyDate(form1.period_start_yy,form1.period_start_mm,form1.period_start_dd,form1.period_start,"���������_��"))
            return false;
    if(!CheckMyDate(form1.period_end_yy,form1.period_end_mm,form1.period_end_dd,form1.period_end,"������������"))
            return false;
    if(!checkNum1(form1.own_stock_amount))
       return false;
    if(!NewcheckRate(form1.own_stock_rate))
       return false;
    if(!checkNum1(form1.rep_own_stock))
       return false;

    return true;

}

//==============================�~�ר���B��CHECK=============
function Check_ZZ010W(form1)
{
    if (form1.period_year.value==''){
       alert("���O�~�פ��i�ť�");
       return false;
    }
    if (form1.beg_date_y.value=='' || form1.beg_date_m.value=='' || form1.beg_date_d.value==''){
       alert("�_��");
       return false;
    }
    if (form1.end_date_y.value=='' || form1.end_date_m.value=='' || form1.end_date_d.value==''){
       alert("����");
       return false;
    }
    if (!checkNum(form1.period_year))
       return false;
    if(!CheckMyDate(form1.beg_date_y,form1.beg_date_m,form1.beg_date_d,form1.beg_date,"�_��"))
        return false;
    if(!CheckMyDate(form1.end_date_y,form1.end_date_m,form1.end_date_d,form1.end_date,"����"))
        return false;

    if((Math.abs(form1.beg_date_y.value) * 10000 + Math.abs(form1.beg_date_m.value) * 100 + Math.abs(form1.beg_date_d.value)) > (Math.abs(form1.end_date_y.value) * 10000 + Math.abs(form1.end_date_m.value) * 100 + Math.abs(form1.end_date_d.value)))
    	{
    		alert("�_�l������i�H�j��פ���");
    		return false;
    	}
//    if(!checkNum1(form1.negotiate_amount))
//       return false;
    return true;
}
//==============================����Ʒ~�򥻸�ƺ��@CHECK=============
function Check_OY004W(form1)
{

    if (form1.const_type.value=='' || form1.business_id.value=='' || form1.business_name.value==''){
       alert("�ݩ�.����Ʒ~�W��.�Τ@�s�����i�ť�");
       return false;
    }

    if(!CheckMyDate(form1.exinvest_date_y,form1.exinvest_date_m,form1.exinvest_date_d,form1.exinvest_date,"�֭��l�����"))
            return false;
    if(!CheckMyDate(form1.book_amount_datey,form1.book_amount_datem,form1.book_amount_dated,form1.book_amount_date,"�b�����B��Ǥ�"))
            return false;
    if(!NewcheckRate2(form1.org_ownstock_rate))
       return false;
    if(!checkNum1(form1.book_amount))
       return false;
    if(!checkNum1(form1.reset_amount))
       return false;
    if(!checkNum(form1.o_year))
       return false;
    if(!NewcheckRate4(form1.o_netvalue_rate))
       return false;
    if(!changeStr9(form1.o_earn_per_share,"2"))
       return false;
    if(!changeStr9(form1.o_cash,"2"))
       return false;
    if(!changeStr9(form1.o_stock,"2"))
       return false;

    if(!checkNum(form1.t_year))
       return false;
    if(!NewcheckRate4(form1.t_netvalue_rate))
       return false;

    if(!changeStr9(form1.t_earn_per_share,"2"))
       return false;
    if(!changeStr9(form1.t_cash,"2"))
       return false;
    if(!changeStr9(form1.t_stock,"2"))
       return false;

    if(!checkNum(form1.stock_year))
       return false;

    if(!changeStr9(form1.highest,"3"))
       return false;
    if(!changeStr9(form1.lowest,"3"))
       return false;
    if(!changeStr9(form1.average,"3"))
       return false;
    if (!Check_Maintain(form1)){
         return false;
    }

    return true;
}
//==============================����Ʒ~-�]�F�������֭���λȦ��ڧ����B��Ƶe�����CHECK=============
function Check_OY004W_ShowSub1(form1)
{

    if(!checkNum1(form1.apply_exinvest))
       return false;
    if(!checkNum1(form1.approve_exinvest_amount))
       return false;
    if(!CheckMyDate(form1.approve_exinvest_datey,form1.approve_exinvest_datem,form1.approve_exinvest_dated,form1.approve_exinvest_date,"�֭�W�[�����"))
        return false;
    if(!checkNum1(form1.exinvest_amount))
       return false;
    if(!NewcheckRate(form1.ownstock_rate))
       return false;
    if(trimString(form1.approve_exinvest_no.value).length>20)
    {
         alert("��J���]�F���֭�W�[���帹���i�W�L20�Ӧr");
         return false;
    }
    if(trimString(form1.reason.value).length>50)
    {
         alert("��J���]�F���֭�W�[�����B�P�Ȧ��ڧ����B�۲��ɤ���]���i�W�L50�Ӧr");
         return false;
    }

    return true;

}
//==============================����Ʒ~-�D�ޤH����ƺ��@�e�����CHECK=============
function Check_OY004W_ShowSubM(form1)
{
    if (form1.position_code.value=='' || form1.name.value=='' ){
       alert("�m�W.¾�٤��i�ť�");
       return false;
    }
    if (form1.induct_date_y.value=='' || form1.induct_date_m.value=='' || form1.induct_date_d.value=='') {
       alert("�N��������i�ť�");
       return false;
    }

//    if(!CheckMyDate(form1.birth_date_y,form1.birth_date_m,form1.birth_date_d,form1.birth_date,"�X�ͤ��"))
//        return false;
    if(!CheckMyDate(form1.induct_date_y,form1.induct_date_m,form1.induct_date_d,form1.induct_date,"�����N�����"))
        return false;

    return true;

}
//==============================����Ʒ~-���ʨƸ�ƺ��@�e�����CHECK=============
function Check_OY004W_ShowSubM1(form1)
{
    if (form1.position_code.value=='' || form1.name.value=='' ){
       alert("�m�W.¾�٤��i�ť�");
       return false;
    }
    if (form1.induct_date_y.value=='' || form1.induct_date_m.value=='' || form1.induct_date_d.value=='') {
       alert("�N��������i�ť�");
       return false;
    }

//    if(!CheckMyDate(form1.birth_date_y,form1.birth_date_m,form1.birth_date_d,form1.birth_date,"�X�ͤ��"))
//        return false;
    if(!CheckMyDate(form1.induct_date_y,form1.induct_date_m,form1.induct_date_d,form1.induct_date,"�N�����"))
        return false;
    if(!CheckMyDate(form1.period_start_y,form1.period_start_m,form1.period_start_d,form1.period_start,"������������"))
        return false;
    if(!CheckMyDate(form1.period_end_y,form1.period_end_m,form1.period_end_d,form1.period_end,"���������_��"))
        return false;

    if(trimString(form1.period_start.value)!="" && trimString(form1.period_end.value)!=""){

       if((Math.abs(form1.period_start_y.value) * 10000 + Math.abs(form1.period_start_m.value) * 100 + Math.abs(form1.period_start_d.value)) > (Math.abs(form1.period_end_y.value) * 10000 + Math.abs(form1.period_end_m.value) * 100 + Math.abs(form1.period_end_d.value)))
    	{
    		alert("���������_�l������i�H�j��פ���");
    		return false;
    	}
    }
    if(!checkNum(form1.rank))
       return false;
    if(!checkNum(form1.appointed_num))
       return false;

    return true;

}
//==============================����Ʒ~-��������Ʒ~�H����ƺ��@�e�����CHECK=============
function Check_OY004W_ShowSub2(form1)
{
    if (form1.position.value=='' || form1.name.value=='' ){
       alert("�m�W.¾�٤��i�ť�");
       return false;
    }
//    if(!CheckMyDate(form1.birth_date_y,form1.birth_date_m,form1.birth_date_d,form1.birth_date,"�X�ͤ��"))
//        return false;

    return true;

}
//==============================�F���ʶU�ڲέp��ƺ��@CHECK=============
function Check_IX016W(form1)
{
    if (form1.period_year.value=='' || form1.policy_loan_no.value=='' || form1.loan_type.value=='' || form1.manage_y_or_n.value==''){
       alert("�F���ʶU�ڥN��.���O�Φ~��.�f�����O.�O�_���D��g�z�Ȧ椣�i�ť�");
       return false;
    }
    // ========== Serissa 2001/11/05 ====================
    if (form1.share_limited_yn.value == '') {
        alert("�O�_�P��L�F���ʶU�ڦ@���B�פ��i�ť�");
        return false;
    } else if (form1.share_limited_yn.value == '1' && form1.sh_policy_loan_no1.value == '') {
        alert("�п�ܦ@���B�׬F���ʶU�ڥN�� �� �O�_�P��L�F���ʶU�ڦ@���B�׽п��<�L>");
        return false;
    } else if (form1.share_limited_yn.value == '2' && (form1.sh_policy_loan_no1.value != ''
                 || form1.sh_policy_loan_no2.value != '' || form1.sh_policy_loan_no3.value != '')) {
//                 || form1.sh_policy_loan_no4.value != '' || form1.sh_policy_loan_no5.value != '')) {
        alert("�Ш����@���B�׬F���ʶU�ڥN�� �� �O�_�P��L�F���ʶU�ڦ@���B�׽п��<��>");
        return false;
    }
    // ========== Serissa End 2001/11/05 ====================

    if(!checkNum(form1.period_year))
       return false;
    if(!checkNum1(form1.total_limited))
       return false;
    if(!checkNum1(form1.deposit_limited))
       return false;
    if(!checkNum(form1.loan_period))
       return false;
    if(!checkNum(form1.extend_time_limit))
       return false;
    if(trimString(form1.loan_purpose.value).length>50)
    {
         alert("��J���U�ڥت����i�W�L100�Ӧr");
         return false;
    }
    if(trimString(form1.terms_of_loan.value).length>100)

    {
         alert("��J���U�ڧQ�v���ڤ��i�W�L100�Ӧr");
         return false;
    }
    if(trimString(form1.guarantee.value).length>60)
    {
         alert("��J����O���󤣥i�W�L60�Ӧr");
         return false;
    }

   // ========== Serissa 2001/11/05 ====================
    if (form1.apply_range.value == '1')
    {
        if(trimString(form1.start_audit_y.value) == "")
        {

            alert("[���z����](�}�l���z���)�椣�i�ť�");
    	    	form1.start_audit_y.focus();
				return false;
        }
    	if(!CheckMyDate(form1.start_audit_y,form1.start_audit_m,form1.start_audit_d,
    	                form1.start_audit_date,"�}�l���z���")) {
       	    return false;
       	}

       	if(trimString(form1.end_apply_y.value) == "")
        {
            alert("[���z����](�I����z���)�椣�i�ť�");
    	    	form1.end_apply_y.focus();
				return false;
        }
   	    if (!CheckMyDate(form1.end_apply_y,form1.end_apply_m,form1.end_apply_d,
   	                 form1.end_apply_date,"�I����z���")) {
       	    return false;
       	}

       	if ((Math.abs(form1.start_audit_y.value) * 1000 + Math.abs(form1.start_audit_m.value) * 10  + Math.abs(form1.start_audit_d.value)) >
         		 (Math.abs(form1.end_apply_y.value) * 1000 + Math.abs(form1.end_apply_m.value) * 10  + Math.abs(form1.end_apply_d.value)) )
    		{
    			alert("[���z����] �_�l������i�H�j��פ���");
         		form1.start_audit_y.focus();
    			return false;
   			}
    }
    else //�����w���z������,�N[�}�l���z����],[�I����z����]�M��
    {
        form1.start_audit_y.value = "";
        form1.start_audit_m.value = "";
        form1.start_audit_d.value = "";
        form1.end_apply_y.value   = "";
        form1.end_apply_m.value   = "";
        form1.end_apply_d.value   = "";
    }

    if(!NewcheckRate3(form1.gov_help_rate)) {
        form1.gov_help_rate[i].focus();
        return false;
    }
    for(i = 0; i < 4; i++) {
        if(isNaN(Math.abs(changeVal(form1.start_year[i]))) ||
           isNaN(Math.abs(changeVal(form1.end_year[i]))) )
        {
            alert("[�~�װ϶�]�п�J�Ʀr");
            return false;
        }
        else
        {
            if (Math.abs(changeVal(form1.start_year[i])) > Math.abs(changeVal(form1.end_year[i])))
            {
    			alert("[�~�װ϶�] �_�l�~�פ��i�H�j��פ�~��");
         		form1.start_year[i].focus();
    			return false;
   			}
        }
        if(!NewcheckRate3(form1.bank_rate[i])) {
            form1.bank_rate[i].focus();
    	    return false;
    	}
    	if(!NewcheckRate3(form1.real_loan_rate[i])) {
            form1.real_loan_rate[i].focus();
    	    return false;
    	}
    }
    // ========== Serissa End 2001/11/05 ====================


   if(!Check_Maintain(form1))
       	    return false;
   return true;
}
//==============================�F���ʶU��-�F���ʶU�ڦW�ٸ�Ƶe�����CHECK=============
function Check_IX016W_StockM(form1)
{

    if(form1.policy_loan_no.value == '' || form1.policy_loan_name.value=='') {
       alert("��줣�i�ť�");
       return false;
    }

    if(trimString(form1.policy_loan_name.value).length>40)
    {
         alert("��J���F���ʶU�ڦW�٤��i�W�L40�Ӧr");
         return false;
    }

    return true;

}
//==============================�F���ʶU��-�U�ڽd���Ƶe�����CHECK=============
function Check_IX016W_ShowSub2(form1)
{
    if (form1.loan_range.value==''){
       alert("�f�ڽd�򤣥i�ť�");
       return false;
    }

    if(trimString(form1.loan_range.value).length>100)
    {
         alert("��J���U�ڽd�򤣥i�W�L100�Ӧr");
         return false;
    }

    return true;

}
//==============================�F���ʶU��-�U�ڹ�H��Ƶe�����CHECK=============
function Check_IX016W_ShowSub3(form1)
{
    if (form1.loan_poniter.value==''){
       alert("�f�ڹ�H���i�ť�");
       return false;
    }

    if(trimString(form1.loan_poniter.value).length>100)
    {
         alert("��J���U�ڹ�H���i�W�L100�Ӧr");
         return false;
    }

    return true;

}
//==============================�F���ʶU��-����ӷ���Ƶe�����CHECK=============
function Check_IX016W_ShowSub4(form1)
{
    if (form1.way.value==''){
       alert("����ӷ��覡���i�ť�");
       return false;
    }
    if ( (form1.way.value!='2') && (form1.way_name.value=='')){
       alert("����ӷ��W�٤��i�ť�");
       return false;
    }
    if(!checkNum1(form1.way_rate))
       return false;
    if(!NewcheckRate2(form1.way_rate))
    	return false;
    return true;

}
//==============================�F���ʶU��-��z���Ω����ɵe�����CHECK=============
function Check_IX016W_ShowSub5(form1)
{
    if (form1.m_year.value=='' || form1.m_month.value==''){
       alert("���餣�i�ť�");
       return false;
    }
    if(!checkNum(form1.m_year))
       return false;
    if(!checkNum1(form1.limited))
       return false;
    if(!checkNum(form1.thismonth_apply))
       return false;
    if(!checkNum(form1.thismonth_aduit))
       return false;
    if(!checkNum(form1.a_aduit))
       return false;
    if(!checkNum1(form1.a_aduit_amount))
       return false;
    if(!checkNum(form1.a_appropriate))
       return false;
    if(!checkNum1(form1.a_appropriate_amount))
       return false;
    if(!checkNum1(form1.loan_money))
       return false;
    if(!checkNum(form1.behind_loan_amount))
       return false;
    if(!checkNum1(form1.behind_loan_money))
       return false;

    return true;

}
//==============================�F���ʶU��-�ѶU�Ȧ�e�����CHECK=============
function Check_IX016W_ShowSub6(form1)
{
    if (form1.bank_code.value==''){
       alert("�ѶU�Ȧ椣�i�ť�");
       return false;
    }
    return true;

}
//=================���w�γ~�H�U������ꤺ�~�����Ҩ�򥻸�ƺ��@�e�����CHECK=============
function Check_WB001W(form1)
{
    if(form1.data_date.value =="null") {
       if((form1.S_Year.value =="") || (form1.S_Month.value =="")) {
               alert("��Ǥ餣�i�ťաI");
               return(false);
       }
       if(!checkNum(form1.S_Year))
       return false;

    }
    if (form1.found_type.value =="") {
        alert("����������i�ťաI");
        return(false);
    }

    if(!changeStr9(form1.lmonth_amt,"1"))
       return false;
    if(!changeStr9(form1.lseason_amt,"1"))
       return false;
    if(!changeStr9(form1.mseason_o_amt,"1"))
       return false;
    if(!changeStr9(form1.mseason_r_amt,"1"))
       return false;
    if(!changeStr9(form1.mseason_n_amt,"1"))
       return false;
    if (!Check_Maintain(form1)){
         return false;
    }
    if (trimString(form1.lmonth_amt.value)=="")
        form1.lmonth_amt.value="0";
    if (trimString(form1.lseason_amt.value)=="")
        form1.lseason_amt.value="0";
    if (trimString(form1.mseason_o_amt.value)=="")
        form1.mseason_o_amt.value="0";
    if (trimString(form1.mseason_n_amt.value)=="")
        form1.mseason_n_amt.value="0";

    form1.mseason_amt.value = eval(changeVal(form1.lmonth_amt)) + eval(changeVal(form1.lseason_amt)) - eval(changeVal(form1.mseason_o_amt));
    form1.n_amt.value = eval(changeVal(form1.mseason_n_amt))-eval(form1.mseason_amt.value);
    return true;

}
//=================�H�U��ꤽ�q�]�ȷ��p�����ƺ��@�e�����CHECK=============
function Check_WB002W(form1)
{
    if(form1.data_date.value =="null") {
       if((form1.S_Year.value =="") || (form1.S_Month.value =="")) {
               alert('��Ǥ餣�i�ťաI');
               return(false);
       }
       if(!checkNum(form1.S_Year))
       return false;
    }

    if(!checkNum1(form1.capital_amount))
       return false;
    if(!checkNum1(form1.member_amount))
       return false;
    if(!checkNum1(form1.bank_branch))
       return false;
    if(!checkNum1(form1.asset_amount))
       return false;
    if(!checkNum1(form1.debt_amount))
       return false;
    if(!checkNum1(form1.stockholder_value))
       return false;
    if(!checkNum1(form1.profit_or_loss))
       return false;
    if(!checkNum1(form1.sure_trust_fund))
       return false;
    if(!checkNum1(form1.gain_or_loss))
       return false;
    if(!checkNum1(form1.ix11_010))
       return false;
    if(!checkNum1(form1.stay_shrink))
       return false;
    if(!checkNum1(form1.stay_had_shrink))
       return false;
    if(!checkNum1(form1.ix11_013))
       return false;
    if(!checkNum1(form1.loan))
       return false;
    if(!checkNum1(form1.guarantee))
       return false;
    if(!checkNum1(form1.guarantee_amount))
       return false;
    if(!checkNum1(form1.commercial_check))
       return false;
    if(!checkNum1(form1.own_stock))
       return false;
    if(!checkNum1(form1.trust_stock))
       return false;
    if(!checkNum1(form1.ix11_020))
       return false;
    if(!checkNum1(form1.direct_shrink))
       return false;
    if(!checkNum1(form1.direct_had_shrink))
       return false;
   if(!checkNum1(form1.ix11_023))
       return false;
    if(!checkNum1(form1.realty_shrink))
       return false;
    if(!checkNum1(form1.realty_had_shrink))
       return false;
    if(!checkNum1(form1.real_estate))
       return false;
    if(!checkNum1(form1.agregate_value))
       return false;
    if(!NewcheckRate(form1.over_loan_rate))
       return false;
    if (!Check_Maintain(form1)){
         return false;
    }

    return true;

}
//===�Ȧ�.�H�U��ꤽ�q���Ҩ���Ĥ��q���Ҩ�ӿ�z����ĳq���Ӱ򥻸�ƺ��@�e�����CHECK=============
function Check_WB003W(form1)
{
    if(form1.data_date.value =='null') {
       if((form1.S_Year.value =='') || (form1.S_Month.value =='')) {
               alert('��Ǥ餣�i�ťաI');
               return(false);
       }
       if(!checkNum(form1.S_Year))
       return false;

    }
    if (form1.loan_kind.value =='') {
        alert('�ɴڤ��q�ݩʤ��i�ťաI');
        return(false);
    }
     if (form1.load_date_y.value=='' || form1.load_date_m.value=='' || form1.load_date_d.value=='') {
       alert("����ĳq������i�ť�");
       return false;
    }
    if(!CheckMyDate(form1.load_date_y,form1.load_date_m,form1.load_date_d,form1.load_date,"����ĳq���"))
       return false;

    if (!Check_Maintain(form1)){
         return false;
    }
    if(!checkNum1(form1.loan_amt))
       return false;
    if(!checkNum1(form1.rt_amt))
       return false;
    if(!checkNum1(form1.tm_net))
       return false;
    if(!checkNum1(form1.net_value))
       return false;
    return true;

}
//===���@�H���e�����CHECK=============
function Check_Maintain(form1)
{
   if (trimString(form1.director_name.value)=="" ||  trimString(form1.maintain_dept.value)==""
       || trimString(form1.director_tel.value)=="" || trimString(form1.maintain_tel.value)==""){
        alert("���@�����W��.�D�ީm�W.�D�޹q��.�ӿ���q����줣�i�ťաI");
        return(false);
   }
   return true;
}

//===�Ȧ���Ҩ�ӿ�z�����Ҩ�ĸ�~�ȩҩ�ɴڶ���ƺ��@�e�����CHECK=============
function Check_WB004W(form1)
{
    /*if(form1.data_date.value =='null') {
       if((form1.S_Year.value =='') || (form1.S_Month.value =='')) {
               alert('��Ǥ餣�i�ťաI');
               return(false);
       }
       if(!checkNum(form1.S_Year))
       return false;

    }*/
     if (form1.bank_date_y.value=='' || form1.bank_date_m.value=='' || form1.bank_date_d.value=='') {
       alert("������i�ť�");
       return false;
    }
    if(!CheckMyDate(form1.bank_date_y,form1.bank_date_m,form1.bank_date_d,form1.bank_date,"���"))
       return false;

    if (!Check_Maintain(form1)){
         return false;
    }
    if(!checkNum1(form1.loan_amount))
       return false;
    if(!checkNum1(form1.working_capital))
       return false;
    form1.S_Year.value = form1.bank_date_y.value;
    form1.S_Month.value = form1.bank_date_m.value;

    return true;

}

//===���ľ��c���U�N���ʳf�d���O�b�ڲέp��ƺ��@�e�����CHECK=============
function Check_WB007W(form1)
{
    if(form1.data_date.value =='null') {
       if((form1.S_Year.value =='') || (form1.S_Month.value =='')) {
               alert('��Ǥ餣�i�ťաI');
               return(false);
       }
       if(!checkNum(form1.S_Year))
       return false;

    }
    if (!Check_Maintain(form1)){
         return false;
    }
    if(!checkNum1(form1.rec_cnt))
       return false;
    if(!checkNum1(form1.rec_amt))
       return false;
    return true;

}
//===���ľ��c���U�N���ʳf�d���O�b�ڲέp��ƺ��@�e�����CHECK=============
function IX015R(form,actionURL)
{
    if(form.s_year.value =="" || form.s_month.value == "" ) {
       alert("������i�ťաI");
       return false;
    }
    if(!checkNum(form.s_year))
        return false;
    form.action = actionURL;
    form.submit();


}
//====================
function doIX005WSubmit(form,myfun,code)
{
    form.Function.value = myfun;
    if(myfun=="delete")
    {
        if(AskDelete())
            form.submit();
    }
    else
    {
    	if (code=="1") { //*�����֭�
           if(Check_IX005W_ShowSub1(form))
              form.submit();
        }
    	if (code=="2") {  //**�����D��
           if(Check_IX005W_ShowSubM(form))
              form.submit();
        }
    	if (code=="3") {  //**���ʨ�
           if(Check_IX005W_ShowSubM1(form))
              form.submit();
        }
    	if (code=="4") {  //**��������Ʒ~
           if(Check_IX005W_ShowSub2(form))
              form.submit();
        }

    }
}
//==============================����Ʒ~�򥻸�ƺ��@CHECK=============
function Check_IX005W(form1)
{

    if (form1.const_type.value=='' || form1.business_id.value=='' || form1.business_name.value==''){
       alert("�ݩ�.����Ʒ~�W��.�Τ@�s�����i�ť�");
       return false;
    }

    if(!CheckMyDate(form1.exinvest_date_y,form1.exinvest_date_m,form1.exinvest_date_d,form1.exinvest_date,"�֭��l�����"))
            return false;
    if(!CheckMyDate(form1.book_amount_datey,form1.book_amount_datem,form1.book_amount_dated,form1.book_amount_date,"�b�����B��Ǥ�"))
            return false;
    if(!NewcheckRate2(form1.org_ownstock_rate))
       return false;
    if(!checkNum1(form1.book_amount))
       return false;
    if(!checkNum1(form1.reset_amount))
       return false;
    if(!checkNum(form1.o_year))
       return false;
    if(!NewcheckRate4(form1.o_netvalue_rate))
       return false;
    if(!changeStr9(form1.o_earn_per_share,"2"))
       return false;
    if(!changeStr9(form1.o_cash,"2"))
       return false;
    if(!changeStr9(form1.o_stock,"2"))
       return false;

    if(!checkNum(form1.t_year))
       return false;
    if(!NewcheckRate4(form1.t_netvalue_rate))
       return false;

    if(!changeStr9(form1.t_earn_per_share,"2"))
       return false;
    if(!changeStr9(form1.t_cash,"2"))
       return false;
    if(!changeStr9(form1.t_stock,"2"))
       return false;

    if(!checkNum(form1.stock_year))
       return false;

    if(!changeStr9(form1.highest,"3"))
       return false;
    if(!changeStr9(form1.lowest,"3"))
       return false;
    if(!changeStr9(form1.average,"3"))
       return false;
    if (!Check_Maintain(form1)){
         return false;
    }

    return true;
}
//==============================����Ʒ~-�]�F�������֭���λȦ��ڧ����B��Ƶe�����CHECK=============
function Check_IX005W_ShowSub1(form1)
{

    if(!checkNum1(form1.apply_exinvest))
       return false;
    if(!checkNum1(form1.approve_exinvest_amount))
       return false;
    if(!CheckMyDate(form1.approve_exinvest_datey,form1.approve_exinvest_datem,form1.approve_exinvest_dated,form1.approve_exinvest_date,"�֭�W�[�����"))
        return false;
    if(!checkNum1(form1.exinvest_amount))
       return false;
    if(!NewcheckRate(form1.ownstock_rate))
       return false;
    if(trimString(form1.approve_exinvest_no.value).length>20)
    {
         alert("��J���]�F���֭�W�[���帹���i�W�L20�Ӧr");
         return false;
    }
    if(trimString(form1.reason.value).length>50)
    {
         alert("��J���]�F���֭�W�[�����B�P�Ȧ��ڧ����B�۲��ɤ���]���i�W�L50�Ӧr");
         return false;
    }

    return true;

}
//==============================����Ʒ~-�D�ޤH����ƺ��@�e�����CHECK=============
function Check_IX005W_ShowSubM(form1)
{
    if (form1.position_code.value=='' || form1.name.value=='' ){
       alert("�m�W.¾�٤��i�ť�");
       return false;
    }
    if (form1.induct_date_y.value=='' || form1.induct_date_m.value=='' || form1.induct_date_d.value=='') {
       alert("�N��������i�ť�");
       return false;
    }
//
//    if(!CheckMyDate(form1.birth_date_y,form1.birth_date_m,form1.birth_date_d,form1.birth_date,"�X�ͤ��"))
//        return false;
    if(!CheckMyDate(form1.induct_date_y,form1.induct_date_m,form1.induct_date_d,form1.induct_date,"�����N�����"))
        return false;

    return true;

}
//==============================����Ʒ~-���ʨƸ�ƺ��@�e�����CHECK=============
function Check_IX005W_ShowSubM1(form1)
{
    if (form1.position_code.value=='' || form1.name.value=='' ){
       alert("�m�W.¾�٤��i�ť�");
       return false;
    }
    if (form1.induct_date_y.value=='' || form1.induct_date_m.value=='' || form1.induct_date_d.value=='') {
       alert("�N��������i�ť�");
       return false;
    }

//    if(!CheckMyDate(form1.birth_date_y,form1.birth_date_m,form1.birth_date_d,form1.birth_date,"�X�ͤ��"))
//        return false;
    if(!CheckMyDate(form1.induct_date_y,form1.induct_date_m,form1.induct_date_d,form1.induct_date,"�N�����"))
        return false;
    if(!CheckMyDate(form1.period_start_y,form1.period_start_m,form1.period_start_d,form1.period_start,"������������"))
        return false;
    if(!CheckMyDate(form1.period_end_y,form1.period_end_m,form1.period_end_d,form1.period_end,"���������_��"))
        return false;

    if(trimString(form1.period_start.value)!="" && trimString(form1.period_end.value)!=""){

       if((Math.abs(form1.period_start_y.value) * 10000 + Math.abs(form1.period_start_m.value) * 100 + Math.abs(form1.period_start_d.value)) > (Math.abs(form1.period_end_y.value) * 10000 + Math.abs(form1.period_end_m.value) * 100 + Math.abs(form1.period_end_d.value)))
    	{
    		alert("���������_�l������i�H�j��פ���");
    		return false;
    	}
    }
    if(!checkNum(form1.rank))
       return false;
    if(!checkNum(form1.appointed_num))
       return false;

    return true;

}
//==============================����Ʒ~-��������Ʒ~�H����ƺ��@�e�����CHECK=============
function Check_IX005W_ShowSub2(form1)
{
    if (form1.position.value=='' || form1.name.value=='' ){
       alert("�m�W.¾�٤��i�ť�");
       return false;
    }
//    if(!CheckMyDate(form1.birth_date_y,form1.birth_date_m,form1.birth_date_d,form1.birth_date,"�X�ͤ��"))
//        return false;

    return true;

}
//==================================�~��Ȧ�b�ؤ�����c�򥻸�ƺ��@�e�����check=========//
function Check_FY004W(form1) {
  /*if (form1.branch_no.value.length ==0){
        alert('������c�W�٤��i�ť�');
        return false;
  }
  */
  if(form1.MyOpType.value == 'Abdicate_Confirm')
  {
        if(!confirm("�T�w�n���M�ܡH"))
            return false;
        if (trimString(form1.BN_DATE_Y.value) == "" )
		{
    	    alert("[���M�ͮĤ�� ](�~)�椣�i�ť�");
       	    form1.BN_DATE_Y.focus();
		    return false;
	    }

	    if (trimString(form1.BN_DATE_M.value) == "" )
	    {

		    alert("[���M�ͮĤ�� ](��)�椣�i�ť�");
       	    form1.BN_DATE_M.focus();
		    return false;
	    }
	    if (trimString(form1.BN_DATE_D.value) == "" )
		{
			    alert("[���M�ͮĤ�� ](��)�椣�i�ť�");
       	    	form1.BN_DATE_D.focus();
	   			return false;
	    }

        if(!CheckMyDate(form1.BN_DATE_Y,form1.BN_DATE_M,form1.BN_DATE_D,form1.BnDate,"���M�ͮĤ��"))
            return false;
   	    return true;
  }
  if(form1.MyOpType.value == 'Delete_Confirm')
  {
    if(!confirm("�T�w�n�R���ܡH"))
        return false;
    else
        return true
  }
      if(form1.MyOpType.value != 'Delete_Confirm' && form1.MyOpType.value != 'Abdicate_Confirm' && form1.MyOpType.value != 'Update_Confirm')
    {
	    if(form1.gserial[1].checked == true)
        {
	        if (trimString(form1.branch_no.value) == "" )
	        {
	    	    alert("�L������c�N�X,�L�k�s�W");
	    	    form1.branch_no.focus();
	    	    return false;
	        }
	        if (trimString(form1.branch_no.value).length != 3 ||
	            isNaN(Math.abs(trimString(form1.branch_no.value))) )
	        {
	    	    alert("�п�J�T��Ʀr");
	    	    form1.branch_no.focus();
	    	    return false;
	        }
	    }
    }
  if (trimString(form1.bankname.value) == "" )
	{
		alert("����W�٤��i���ť�");
		form1.bankname.focus();
		return false;
	}
	if(trimString(form1.bank_type.value) == "")
	{
		alert("�������O���i���ť�");
		form1.bank_type.focus();
		return false;
	}

  if(!CheckMyDate(form1.prv_yy,form1.prv_mm,form1.prv_dd,form1.prv_date,"�֭�]�ߤ��"))
      return false;

  if(!CheckMyDate(form1.open_yy,form1.open_mm,form1.open_dd,form1.open_date,"�]�ߤ��"))
      return false;

  if(!CheckMyDate(form1.CHG_LICENSE_DATE_Y,form1.CHG_LICENSE_DATE_M,form1.CHG_LICENSE_DATE_D,form1.CHG_LICENSE_DATE,"�̪񴫷Ӥ��"))
      return false;

  if(!CheckMyDate(form1.oper_yy,form1.oper_mm,form1.oper_dd,form1.oper_date,"�}�l��~��"))
      return false;

  if(trimString(form1.hsiend_id.value) == "")
	{
		alert("�����W�٤ζl���ϸ����i���ť�");
		form1.hsiend_id.focus();
		return false;
	}
	if (trimString(form1.addr.value) == "" )
	{
		alert("[�a�}] ���i�ť�");
    	form1.addr.focus();
		return false;
   	}
  if (!Check_Maintain(form1)){
         return false;
    }
}
//====================================
function doFY004WSubmit(form,myfun)
{
    form.Function.value = myfun;
    if(myfun=="delete")
    {
        if(AskDelete())
            form.submit();
    }
    else
    {
        if(form.code_name.value==""){
      	   alert("��~���ئW�٤��i�ť�");
           return false;
        }
            form.submit();
    }
}
//======================================�~��Ȧ�b�ؿ�ƳB��������CHECK===
function Check_FY006W(form1) {

  if (trimString(form1.corp_id.value).length==0){
      alert('�Τ@�s�����i�ť�');
         return false;
    }
  if(!CheckMyDate(form1.prv_yy,form1.prv_mm,form1.prv_dd,form1.prv_date,"�֭�]�ߤ��"))
      return false;
  if(!CheckMyDate(form1.setup_yy,form1.setup_mm,form1.setup_dd,form1.setup_date,"�]�ߤ��"))
      return false;
  if(!CheckMyDate(form1.CHG_LICENSE_DATE_Y,form1.CHG_LICENSE_DATE_M,form1.CHG_LICENSE_DATE_D,form1.CHG_LICENSE_DATE,"�̪񴫷Ӥ��"))
      return false;

  if (form1.addr.value.length !=0 ) {
     if (form1.hsiend_id.value.length ==0) {
       alert('�����W�٤ζl���ϸ����i�ť�');
        return false;
     }
  }

  if (!Check_Maintain(form1)){
         return false;
    }

}
//====================================================================

function Check_FY006W_ShowSubR(form)
{
    if(form.corp_id.value=="")
    {
    	alert("�Τ@�s�����i�ť�");
        return false;
    }
     if(!NewcheckRate2(form.rate))
       return false;

    return true;
}

function doFY006WSubmit(form,myfun)
{
    form.Function.value = myfun;
    if(myfun=="delete")
    {
        if(AskDelete())
            form.submit();
    }
    else
    {
        if(Check_FY006W_ShowSubR(form))
            form.submit();
    }
}

//==============================����Ʒ~�򥻸�ƺ��@CHECK=============
function Check_OX004W(form1)
{

    if (form1.const_type.value=='' || form1.business_id.value=='' || form1.business_name.value==''){
       alert("�ݩ�.����Ʒ~�W��.�Τ@�s�����i�ť�");
       return false;
    }

    if(!CheckMyDate(form1.exinvest_date_y,form1.exinvest_date_m,form1.exinvest_date_d,form1.exinvest_date,"�֭��l�����"))
            return false;
    if(!CheckMyDate(form1.book_amount_datey,form1.book_amount_datem,form1.book_amount_dated,form1.book_amount_date,"�b�����B��Ǥ�"))
            return false;
    if(!NewcheckRate2(form1.org_ownstock_rate))
       return false;
    if(!checkNum1(form1.book_amount))
       return false;
    if(!checkNum1(form1.reset_amount))
       return false;
    if(!checkNum(form1.o_year))
       return false;
    if(!NewcheckRate4(form1.o_netvalue_rate))
       return false;
    if(!changeStr9(form1.o_earn_per_share,"2"))
       return false;
    if(!changeStr9(form1.o_cash,"2"))
       return false;
    if(!changeStr9(form1.o_stock,"2"))
       return false;

    if(!checkNum(form1.t_year))
       return false;
    if(!NewcheckRate4(form1.t_netvalue_rate))
       return false;

    if(!changeStr9(form1.t_earn_per_share,"2"))
       return false;
    if(!changeStr9(form1.t_cash,"2"))
       return false;
    if(!changeStr9(form1.t_stock,"2"))
       return false;

    if(!checkNum(form1.stock_year))
       return false;

    if(!changeStr9(form1.highest,"3"))
       return false;
    if(!changeStr9(form1.lowest,"3"))
       return false;
    if(!changeStr9(form1.average,"3"))
       return false;
    if (!Check_Maintain(form1)){
         return false;
    }

    return true;
}
//==============================����Ʒ~-�]�F�������֭���λȦ��ڧ����B��Ƶe�����CHECK=============
function Check_OX004W_ShowSub1(form1)
{

    if(!checkNum1(form1.apply_exinvest))
       return false;
    if(!checkNum1(form1.approve_exinvest_amount))
       return false;
    if(!CheckMyDate(form1.approve_exinvest_datey,form1.approve_exinvest_datem,form1.approve_exinvest_dated,form1.approve_exinvest_date,"�֭�W�[�����"))
        return false;
    if(!checkNum1(form1.exinvest_amount))
       return false;
    if(!NewcheckRate(form1.ownstock_rate))
       return false;
    if(trimString(form1.approve_exinvest_no.value).length>20)
    {
         alert("��J���]�F���֭�W�[���帹���i�W�L20�Ӧr");
         return false;
    }
    if(trimString(form1.reason.value).length>50)
    {
         alert("��J���]�F���֭�W�[�����B�P�Ȧ��ڧ����B�۲��ɤ���]���i�W�L50�Ӧr");
         return false;
    }

    return true;

}
//==============================����Ʒ~-�D�ޤH����ƺ��@�e�����CHECK=============
function Check_OX004W_ShowSubM(form1)
{
    if (form1.position_code.value=='' || form1.name.value=='' ){
       alert("�m�W.¾�٤��i�ť�");
       return false;
    }
    if (form1.induct_date_y.value=='' || form1.induct_date_m.value=='' || form1.induct_date_d.value=='') {
       alert("�N��������i�ť�");
       return false;
    }

//    if(!CheckMyDate(form1.birth_date_y,form1.birth_date_m,form1.birth_date_d,form1.birth_date,"�X�ͤ��"))
//        return false;
    if(!CheckMyDate(form1.induct_date_y,form1.induct_date_m,form1.induct_date_d,form1.induct_date,"�����N�����"))
        return false;

    return true;

}
//==============================����Ʒ~-���ʨƸ�ƺ��@�e�����CHECK=============
function Check_OX004W_ShowSubM1(form1)
{
    if (form1.position_code.value=='' || form1.name.value=='' ){
       alert("�m�W.¾�٤��i�ť�");
       return false;
    }
    if (form1.induct_date_y.value=='' || form1.induct_date_m.value=='' || form1.induct_date_d.value=='') {
       alert("�N��������i�ť�");
       return false;
    }

//    if(!CheckMyDate(form1.birth_date_y,form1.birth_date_m,form1.birth_date_d,form1.birth_date,"�X�ͤ��"))
//        return false;
    if(!CheckMyDate(form1.induct_date_y,form1.induct_date_m,form1.induct_date_d,form1.induct_date,"�N�����"))
        return false;
    if(!CheckMyDate(form1.period_start_y,form1.period_start_m,form1.period_start_d,form1.period_start,"������������"))
        return false;
    if(!CheckMyDate(form1.period_end_y,form1.period_end_m,form1.period_end_d,form1.period_end,"���������_��"))
        return false;

    if(trimString(form1.period_start.value)!="" && trimString(form1.period_end.value)!=""){

       if((Math.abs(form1.period_start_y.value) * 10000 + Math.abs(form1.period_start_m.value) * 100 + Math.abs(form1.period_start_d.value)) > (Math.abs(form1.period_end_y.value) * 10000 + Math.abs(form1.period_end_m.value) * 100 + Math.abs(form1.period_end_d.value)))
    	{
    		alert("���������_�l������i�H�j��פ���");
    		return false;
    	}
    }
    if(!checkNum(form1.rank))
       return false;
    if(!checkNum(form1.appointed_num))
       return false;

    return true;

}
//==============================����Ʒ~-��������Ʒ~�H����ƺ��@�e�����CHECK=============
function Check_OX004W_ShowSub2(form1)
{
    if (form1.position.value=='' || form1.name.value=='' ){
       alert("�m�W.¾�٤��i�ť�");
       return false;
    }
//    if(!CheckMyDate(form1.birth_date_y,form1.birth_date_m,form1.birth_date_d,form1.birth_date,"�X�ͤ��"))
//        return false;

    return true;

}
//==============================���P���ľ��c���@�e�����CHECK=============
function Check_ZZ013W(form1)
{
    if (form1.bank_no.value=='' || form1.d_date_y.value=='' || form1.d_date_m.value=='' ||
       form1.d_date_d.value=='' || form1.setup_no.value=='' || form1.setup_date_y.value=='' ||
       form1.setup_date_m.value=='' || form1.setup_date_d.value=='' || form1.setup_reason.value==''){
       alert("�e�����������i�ť�");
       return false;
    }
    if(!CheckMyDate(form1.d_date_y,form1.d_date_m,form1.d_date_d,form1.d_date,"���P���"))
        return false;
    if(!CheckMyDate(form1.setup_date_y,form1.setup_date_m,form1.setup_date_d,form1.setup_date,"�ַǤ��"))
        return false;

    return true;

}
//==============================��~���إN�X���@�e�����CHECK=============
function Check_ZZ011W(form1)
{
    if (form1.code.value=='' || form1.item.value==''){
       alert("�e�����������i�ť�");
       return false;
    }
    if(trimString(form1.item.value).length>70)
    {
         alert("��J����~�W�٤��i�W�L70�Ӧr");
         return false;
    }
    if(!checkNum(form1.code))
        return false;
    return true;

}
//serissa
//==============================�F���ʶU��-�ѶU�Ȧ�e�����CHECK=============
function Check_ZZ010W_ShowSub(form)
{
    if (form.Function.value == "insert"){
        if (form.bank_no.value==''){
            alert("���c�W�٤��i�ť�");
            form.bank_no.focus();
            return false;
        }
    }
    if (form.negotiate_amount.value==''){
       alert("�п�J����B��");
       form.negotiate_amount.focus();
       return false;
    }
    if(isNaN(Math.abs(changeVal(form.negotiate_amount))) )
    {
        alert("�п�J�Ʀr");
        return false;
    }
    return true;
}

function doZZ010WSubmit(form,myfun)
{
    form.Function.value = myfun;
    if(myfun=="delete")
    {
        if(AskDelete())
            form.submit();
    }
    else
    {   if(Check_ZZ010W_ShowSub(form)){

    	    form.submit();
        }
    }
}

function ChangeAction(form,actionURL)
{
	form.action = actionURL;
	if(form.S_YEAR.value == "" || form.E_YEAR.value == ""){
        alert("�п�J���O�~��");
        return false;
    }
	form.submit();
}

function Check_ZZ010W_Submit(form)
{
    if(!CheckYear(form.S_YEAR))
            return false;
    if(!CheckYear(form.E_YEAR))
            return false;


	if(trimString(form.S_YEAR.value)!="" && trimString(form.E_YEAR.value)!="")
	{
		if(Math.abs(form.S_YEAR.value) > Math.abs(form.E_YEAR.value))
    	{
    		alert("�_�l�~�פ��i�H�j��פ�~��");
    		form.S_YEAR.focus()
    		return false;
    	}
    }
    return true;
}

