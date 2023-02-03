function CheckMyDate(checkyear,checkmonth,checkday,checkdate,szWarning)
{
    if(!CheckYear(checkyear))
            return false;
    if(trimString(checkyear.value)!="" && trimString(checkmonth.value)!="" && trimString(checkday.value)!="")
    {
        var setyear = (Math.abs(checkyear.value)) + 1911;
        if(!fnValidDate(setyear + checkmonth.value + checkday.value))
        {
            alert(szWarning+"有誤");
            return false;
        }
        checkdate.value = setyear + "/"+checkmonth.value + "/"+checkday.value;
//        checkdate.value = setyear + checkmonth.value +checkday.value;
    }
    else if(trimString(checkyear.value)=="" && trimString(checkmonth.value)=="" && trimString(checkday.value)==""){}
         else
         {
            alert("請輸入完整的"+ szWarning+ "或者都不輸入");
            return false;
         }
    return true
}


//==============================
function CheckYear(cyear)
{
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
   	    alert("請輸入數字");
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
   	    alert("請輸入數字");
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
        alert("金額不可大於 99999999999999.99, 也不可小於 -99999999999999.99")
        return oring;
    }
    if((T1.value).indexOf(".") != -1 )
    {
        len = (T1.value).substring((T1.value).indexOf(".")+1,T1.value.length);
        if(len.length > 2)
        {
            alert("小數點後只能有兩個位數");
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
   	    alert("請輸入數字");
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
        alert("金額不可大於 99999999999999.99, 也不可小於 -99999999999999.99")
        return false;
    }
    }
    if (code=="3") {
  //===============================================
    if(eval(oring) > 99999999.99 || eval(oring) < -99999999.99)
    {
        alert("金額不可大於 99999999.99, 也不可小於 -99999999.99")
        return false;
    }
    }
    if (code=="2") {
  //===============================================
    if(eval(oring) > 999999.99 || eval(oring) < -999999.99)
    {
        alert("金額不可大於 999999.99, 也不可小於 -999999.99")
        return false;
    }
    }
	if (code=="4") {
  //===============================================
    if(eval(oring) > 999.99 || eval(oring) < -999.99)
    {
        alert("金額不可大於 999.99, 也不可小於 -999.99")
        return false;
    }
    }
    if((T1.value).indexOf(".") != -1 )
    {
        len = (T1.value).substring((T1.value).indexOf(".")+1,T1.value.length);
        if(len.length > 2)
        {
            alert("小數點後只能有兩個位數");
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
   	    alert("請輸入數字");
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
        alert("金額不可大於 999999.99, 也不可小於 -999999.99")
        return oring;
    }
    if((T1.value).indexOf(".") != -1 )
    {
        len = (T1.value).substring((T1.value).indexOf(".")+1,T1.value.length);
        if(len.length > 2)
        {
            alert("小數點後只能有兩個位數");
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
   	    alert("請輸入數字");
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
        alert("金額不可大於 99999999.99")
        return oring;
    }
    if((T1.value).indexOf(".") != -1 )
    {
        len = (T1.value).substring((T1.value).indexOf(".")+1,T1.value.length);
        if(len.length > 2)
        {
            alert("小數點後只能有兩個位數");
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
//**判斷輸入資料是否介於99.99－(-99.99)
//===========================================================================
function NewcheckRate1(Rate1)
{
    if(isNaN(Math.abs(Rate1.value)))
    {
   	    alert("請輸入數字");
   	    return false;
    }
    if(eval(Rate1.value) > 99.99 || eval(Rate1.value) < -99.99)
    {
        alert("利率不可大於 99.99, 也不可小於 -99.99")
        return false;
    }
    if((Rate1.value).indexOf(".") != -1 )
    {
        len = (Rate1.value).substring((Rate1.value).indexOf(".")+1,Rate1.value.length);
        if(len.length > 2)
        {
            alert("小數點後只能有兩個位數");
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
   	    alert("請輸入數字");
   	    return false;
    }
    if(eval(Rate1.value) > 99.99 || eval(Rate1.value) < 0)
    {
        alert("利率不可大於 99.99, 也不可小於 0")
        return false;
    }
    if((Rate1.value).indexOf(".") != -1 )
    {
        len = (Rate1.value).substring((Rate1.value).indexOf(".")+1,Rate1.value.length);
        if(len.length > 2)
        {
            alert("小數點後只能有兩個位數");
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
   	    alert("請輸入數字");
   	    return false;
    }
    if(eval(Rate1.value) > 100 || eval(Rate1.value) < 0)
    {
        alert("利率不可大於 100, 也不可小於 0")
        return false;
    }
    if((Rate1.value).indexOf(".") != -1 )
    {
        len = (Rate1.value).substring((Rate1.value).indexOf(".")+1,Rate1.value.length);
        if(len.length > 2)
        {
            alert("小數點後只能有兩個位數");
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
   	    alert("請輸入數字");
   	    return false;
    }
    if(eval(Rate1.value) > 99.999 || eval(Rate1.value) < -99.999)
    {
        alert("利率不可大於 99.999, 也不可小於 -99.999")
        return false;
    }
    if((Rate1.value).indexOf(".") != -1 )
    {
        len = (Rate1.value).substring((Rate1.value).indexOf(".")+1,Rate1.value.length);
        if(len.length > 3)
        {
            alert("小數點後只能有三個位數");
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
   	    alert("請輸入數字");
   	    return false;
    }
    if(eval(Rate1.value) > 9999.99 || eval(Rate1.value) < -9999.99)
    {
        alert("利率不可大於 9999.99, 也不可小於 -9999.99")
        return false;
    }
    if((Rate1.value).indexOf(".") != -1 )
    {
        len = (Rate1.value).substring((Rate1.value).indexOf(".")+1,Rate1.value.length);
        if(len.length > 2)
        {
            alert("小數點後只能有兩個位數");
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
  if(confirm("確定要刪除嗎？")) {
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
//===============理監事卸任判斷紐=============//
function MoveAction(form1) {

    if (form1.abdicate_date_yy.value=='' || form1.abdicate_date_mm.value=='' || form1.abdicate_date_dd.value==''){
       alert("請輸入卸任日期");
       return false;
    }
   if(!CheckMyDate(form1.abdicate_date_yy,form1.abdicate_date_mm,form1.abdicate_date_dd,form1.abdicate_date,"卸任日期"))
       return false;

   form1.Function.value = 'delete';
   //form1.Function.value = 'move';
   form1.submit();
}

function checkNum(Rate1)
{
    if(isNaN(Math.abs(Rate1.value)))
    {
   	    alert("請輸入數字");
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
   	    alert("請輸入數字");
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
    	if (code=="1") { //*歷次核准
           if(Check_OY004W_ShowSub1(form))
              form.submit();
        }
    	if (code=="2") {  //**高階主管
           if(Check_OY004W_ShowSubM(form))
              form.submit();
        }
    	if (code=="3") {  //**董監事
           if(Check_OY004W_ShowSubM1(form))
              form.submit();
        }
    	if (code=="4") {  //**派兼轉投資事業
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
    	if (code=="1") { //*歷次核准
           if(Check_OX004W_ShowSub1(form))
              form.submit();
        }
    	if (code=="2") {  //**高階主管
           if(Check_OX004W_ShowSubM(form))
              form.submit();
        }
    	if (code=="3") {  //**董監事
           if(Check_OX004W_ShowSubM1(form))
              form.submit();
        }
    	if (code=="4") {  //**派兼轉投資事業
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
    	if (code=="6") { //*參貸銀行
           if(Check_IX016W_ShowSub6(form))
              form.submit();
        }
    	if (code=="5") {  //**辦理情形明細
           if(Check_IX016W_ShowSub5(form))
              form.submit();
        }
    	if (code=="4") {  //**資金來源檔
           if(Check_IX016W_ShowSub4(form))
              form.submit();
        }
    	if (code=="3") {  //**貸款對象
           if(Check_IX016W_ShowSub3(form))
              form.submit();
        }
    	if (code=="2") {  //**貸款範圍
           if(Check_IX016W_ShowSub2(form))
              form.submit();
        }

    }
}
//==============================信合社董監事基本資料維護畫面CHECK=============
function LinkToInsert(form,linkaddr,code)
{
    linkaddr = "/servlet/" + linkaddr;
    if(form.Function.value=='Add')
    {
    	if (code==1) {
           alert("請先新增轉投資事業基本資料");
           return;
        }
      	if (code==2) {
           alert("請先新增政策性貸款統計資料");
        return;
        }
      	if (code==3) {
           alert("請先新增外國銀行在華分支機構基本資料");
        return;
        }
        if (code==4) {
           alert("請先新增外國銀行在華辦事處資料");
           return;
        }
        //serissa
        if (code==5) {
            if (form.period_year.value == ""){
                alert("請先新增年度協商額度資料");
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

//==============================信合社董監事基本資料維護畫面CHECK=============
function Check_BX003W(form1)
{
    if (form1.id.value=='' || form1.position_code.value=='' || form1.name.value==''){
       alert("身分證.職稱.姓名不可空白");
       return false;
    }
    if (form1.induct_date_yy.value=='' || form1.induct_date_mm.value=='' || form1.induct_date_dd.value==''){
       alert("就任日期不可空白");
       return false;
    }

    if(!CheckMyDate(form1.birth_yy,form1.birth_mm,form1.birth_dd,form1.birth_date,"出生日期"))
            return false;
    if(!CheckMyDate(form1.induct_date_yy,form1.induct_date_mm,form1.induct_date_dd,form1.induct_date,"就任日期"))
            return false;
    if(!CheckMyDate(form1.period_start_yy,form1.period_start_mm,form1.period_start_dd,form1.period_start,"本屆任期起日"))
            return false;
    if(!CheckMyDate(form1.period_end_yy,form1.period_end_mm,form1.period_end_dd,form1.period_end,"本屆任期迄日"))
            return false;
    if(!checkNum1(form1.own_stock_amount))
       return false;
    if(!NewcheckRate(form1.own_stock_rate))
       return false;
    return true;

}

//==============================信合社持股金額前十大社員基本資料維護CHECK=============
function Check_BX004W(form1)
{
    if (form1.id.value=='' || form1.g_or_c.value=='' || form1.name.value==''){
       alert("身分證.姓名.持股類型不可空白");
       return false;
    }

    return true;

}

//==============================農漁會董監事基本資料維護畫面CHECK=============
function Check_FX003W(form1)
{
    if (form1.id.value=='' || form1.position_code.value=='' || form1.name.value==''){
       alert("身分證.職稱.姓名不可空白");
       return false;
    }
    if (form1.induct_date_yy.value=='' || form1.induct_date_mm.value=='' || form1.induct_date_dd.value==''){
       alert("就任日期不可空白");
       return false;
    }

    if(!CheckMyDate(form1.birth_yy,form1.birth_mm,form1.birth_dd,form1.birth_date,"出生日期"))
            return false;
    if(!CheckMyDate(form1.induct_date_yy,form1.induct_date_mm,form1.induct_date_dd,form1.induct_date,"就任日期"))
            return false;
    if(!CheckMyDate(form1.period_start_yy,form1.period_start_mm,form1.period_start_dd,form1.period_start,"本屆任期起日"))
            return false;
    if(!CheckMyDate(form1.period_end_yy,form1.period_end_mm,form1.period_end_dd,form1.period_end,"本屆任期迄日"))
            return false;
    return true;

}
//==============================董監事基本資料維護畫面CHECK=============
function Check_OY003W(form1)
{
    if (form1.id.value=='' || form1.position_code.value=='' || form1.name.value==''){
       alert("身分證.職稱.姓名不可空白");
       return false;
    }
    if (form1.induct_date_yy.value=='' || form1.induct_date_mm.value=='' || form1.induct_date_dd.value==''){
       alert("就任日期不可空白");
       return false;
    }

    if(!CheckMyDate(form1.birth_yy,form1.birth_mm,form1.birth_dd,form1.birth_date,"出生日期"))
            return false;
    if(!CheckMyDate(form1.induct_date_yy,form1.induct_date_mm,form1.induct_date_dd,form1.induct_date,"就任日期"))
            return false;
    if(!checkNum1(form1.appointed_num))
       return false;
    if(!CheckMyDate(form1.period_start_yy,form1.period_start_mm,form1.period_start_dd,form1.period_start,"本屆任期起日"))
            return false;
    if(!CheckMyDate(form1.period_end_yy,form1.period_end_mm,form1.period_end_dd,form1.period_end,"本屆任期迄日"))
            return false;
    if(!checkNum1(form1.own_stock_amount))
       return false;
    if(!NewcheckRate(form1.own_stock_rate))
       return false;
    if(!checkNum1(form1.rep_own_stock))
       return false;

    return true;

}

//==============================年度協商額度CHECK=============
function Check_ZZ010W(form1)
{
    if (form1.period_year.value==''){
       alert("期別年度不可空白");
       return false;
    }
    if (form1.beg_date_y.value=='' || form1.beg_date_m.value=='' || form1.beg_date_d.value==''){
       alert("起日");
       return false;
    }
    if (form1.end_date_y.value=='' || form1.end_date_m.value=='' || form1.end_date_d.value==''){
       alert("迄日");
       return false;
    }
    if (!checkNum(form1.period_year))
       return false;
    if(!CheckMyDate(form1.beg_date_y,form1.beg_date_m,form1.beg_date_d,form1.beg_date,"起日"))
        return false;
    if(!CheckMyDate(form1.end_date_y,form1.end_date_m,form1.end_date_d,form1.end_date,"迄日"))
        return false;

    if((Math.abs(form1.beg_date_y.value) * 10000 + Math.abs(form1.beg_date_m.value) * 100 + Math.abs(form1.beg_date_d.value)) > (Math.abs(form1.end_date_y.value) * 10000 + Math.abs(form1.end_date_m.value) * 100 + Math.abs(form1.end_date_d.value)))
    	{
    		alert("起始日期不可以大於終止日期");
    		return false;
    	}
//    if(!checkNum1(form1.negotiate_amount))
//       return false;
    return true;
}
//==============================轉投資事業基本資料維護CHECK=============
function Check_OY004W(form1)
{

    if (form1.const_type.value=='' || form1.business_id.value=='' || form1.business_name.value==''){
       alert("屬性.轉投資事業名稱.統一編號不可空白");
       return false;
    }

    if(!CheckMyDate(form1.exinvest_date_y,form1.exinvest_date_m,form1.exinvest_date_d,form1.exinvest_date,"核准原始投資日期"))
            return false;
    if(!CheckMyDate(form1.book_amount_datey,form1.book_amount_datem,form1.book_amount_dated,form1.book_amount_date,"帳面金額基準日"))
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
//==============================轉投資事業-財政部歷次核准投資及銀行實際投資金額資料畫面欄位CHECK=============
function Check_OY004W_ShowSub1(form1)
{

    if(!checkNum1(form1.apply_exinvest))
       return false;
    if(!checkNum1(form1.approve_exinvest_amount))
       return false;
    if(!CheckMyDate(form1.approve_exinvest_datey,form1.approve_exinvest_datem,form1.approve_exinvest_dated,form1.approve_exinvest_date,"核准增加投資日期"))
        return false;
    if(!checkNum1(form1.exinvest_amount))
       return false;
    if(!NewcheckRate(form1.ownstock_rate))
       return false;
    if(trimString(form1.approve_exinvest_no.value).length>20)
    {
         alert("輸入的財政部核准增加投資文號不可超過20個字");
         return false;
    }
    if(trimString(form1.reason.value).length>50)
    {
         alert("輸入的財政部核准增加投資金額與銀行實際投資金額相異時之原因不可超過50個字");
         return false;
    }

    return true;

}
//==============================轉投資事業-主管人員資料維護畫面欄位CHECK=============
function Check_OY004W_ShowSubM(form1)
{
    if (form1.position_code.value=='' || form1.name.value=='' ){
       alert("姓名.職稱不可空白");
       return false;
    }
    if (form1.induct_date_y.value=='' || form1.induct_date_m.value=='' || form1.induct_date_d.value=='') {
       alert("就任日期不可空白");
       return false;
    }

//    if(!CheckMyDate(form1.birth_date_y,form1.birth_date_m,form1.birth_date_d,form1.birth_date,"出生日期"))
//        return false;
    if(!CheckMyDate(form1.induct_date_y,form1.induct_date_m,form1.induct_date_d,form1.induct_date,"本屆就任日期"))
        return false;

    return true;

}
//==============================轉投資事業-董監事資料維護畫面欄位CHECK=============
function Check_OY004W_ShowSubM1(form1)
{
    if (form1.position_code.value=='' || form1.name.value=='' ){
       alert("姓名.職稱不可空白");
       return false;
    }
    if (form1.induct_date_y.value=='' || form1.induct_date_m.value=='' || form1.induct_date_d.value=='') {
       alert("就任日期不可空白");
       return false;
    }

//    if(!CheckMyDate(form1.birth_date_y,form1.birth_date_m,form1.birth_date_d,form1.birth_date,"出生日期"))
//        return false;
    if(!CheckMyDate(form1.induct_date_y,form1.induct_date_m,form1.induct_date_d,form1.induct_date,"就任日期"))
        return false;
    if(!CheckMyDate(form1.period_start_y,form1.period_start_m,form1.period_start_d,form1.period_start,"本屆任期迄日"))
        return false;
    if(!CheckMyDate(form1.period_end_y,form1.period_end_m,form1.period_end_d,form1.period_end,"本屆任期起日"))
        return false;

    if(trimString(form1.period_start.value)!="" && trimString(form1.period_end.value)!=""){

       if((Math.abs(form1.period_start_y.value) * 10000 + Math.abs(form1.period_start_m.value) * 100 + Math.abs(form1.period_start_d.value)) > (Math.abs(form1.period_end_y.value) * 10000 + Math.abs(form1.period_end_m.value) * 100 + Math.abs(form1.period_end_d.value)))
    	{
    		alert("本屆任期起始日期不可以大於終止日期");
    		return false;
    	}
    }
    if(!checkNum(form1.rank))
       return false;
    if(!checkNum(form1.appointed_num))
       return false;

    return true;

}
//==============================轉投資事業-派兼轉投資事業人員資料維護畫面欄位CHECK=============
function Check_OY004W_ShowSub2(form1)
{
    if (form1.position.value=='' || form1.name.value=='' ){
       alert("姓名.職稱不可空白");
       return false;
    }
//    if(!CheckMyDate(form1.birth_date_y,form1.birth_date_m,form1.birth_date_d,form1.birth_date,"出生日期"))
//        return false;

    return true;

}
//==============================政策性貸款統計資料維護CHECK=============
function Check_IX016W(form1)
{
    if (form1.period_year.value=='' || form1.policy_loan_no.value=='' || form1.loan_type.value=='' || form1.manage_y_or_n.value==''){
       alert("政策性貸款代號.期別或年度.貨款類別.是否有主辦經理銀行不可空白");
       return false;
    }
    // ========== Serissa 2001/11/05 ====================
    if (form1.share_limited_yn.value == '') {
        alert("是否與其他政策性貸款共用額度不可空白");
        return false;
    } else if (form1.share_limited_yn.value == '1' && form1.sh_policy_loan_no1.value == '') {
        alert("請選擇共用額度政策性貸款代號 或 是否與其他政策性貸款共用額度請選擇<無>");
        return false;
    } else if (form1.share_limited_yn.value == '2' && (form1.sh_policy_loan_no1.value != ''
                 || form1.sh_policy_loan_no2.value != '' || form1.sh_policy_loan_no3.value != '')) {
//                 || form1.sh_policy_loan_no4.value != '' || form1.sh_policy_loan_no5.value != '')) {
        alert("請取消共用額度政策性貸款代號 或 是否與其他政策性貸款共用額度請選擇<有>");
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
         alert("輸入的貸款目的不可超過100個字");
         return false;
    }
    if(trimString(form1.terms_of_loan.value).length>100)

    {
         alert("輸入的貸款利率條款不可超過100個字");
         return false;
    }
    if(trimString(form1.guarantee.value).length>60)
    {
         alert("輸入的擔保條件不可超過60個字");
         return false;
    }

   // ========== Serissa 2001/11/05 ====================
    if (form1.apply_range.value == '1')
    {
        if(trimString(form1.start_audit_y.value) == "")
        {

            alert("[受理期間](開始受理日期)欄不可空白");
    	    	form1.start_audit_y.focus();
				return false;
        }
    	if(!CheckMyDate(form1.start_audit_y,form1.start_audit_m,form1.start_audit_d,
    	                form1.start_audit_date,"開始受理日期")) {
       	    return false;
       	}

       	if(trimString(form1.end_apply_y.value) == "")
        {
            alert("[受理期間](截止受理日期)欄不可空白");
    	    	form1.end_apply_y.focus();
				return false;
        }
   	    if (!CheckMyDate(form1.end_apply_y,form1.end_apply_m,form1.end_apply_d,
   	                 form1.end_apply_date,"截止受理日期")) {
       	    return false;
       	}

       	if ((Math.abs(form1.start_audit_y.value) * 1000 + Math.abs(form1.start_audit_m.value) * 10  + Math.abs(form1.start_audit_d.value)) >
         		 (Math.abs(form1.end_apply_y.value) * 1000 + Math.abs(form1.end_apply_m.value) * 10  + Math.abs(form1.end_apply_d.value)) )
    		{
    			alert("[受理期間] 起始日期不可以大於終止日期");
         		form1.start_audit_y.focus();
    			return false;
   			}
    }
    else //不指定受理期間時,將[開始受理期間],[截止受理期間]清空
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
            alert("[年度區間]請輸入數字");
            return false;
        }
        else
        {
            if (Math.abs(changeVal(form1.start_year[i])) > Math.abs(changeVal(form1.end_year[i])))
            {
    			alert("[年度區間] 起始年度不可以大於終止年度");
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
//==============================政策性貸款-政策性貸款名稱資料畫面欄位CHECK=============
function Check_IX016W_StockM(form1)
{

    if(form1.policy_loan_no.value == '' || form1.policy_loan_name.value=='') {
       alert("欄位不可空白");
       return false;
    }

    if(trimString(form1.policy_loan_name.value).length>40)
    {
         alert("輸入的政策性貸款名稱不可超過40個字");
         return false;
    }

    return true;

}
//==============================政策性貸款-貸款範圍資料畫面欄位CHECK=============
function Check_IX016W_ShowSub2(form1)
{
    if (form1.loan_range.value==''){
       alert("貨款範圍不可空白");
       return false;
    }

    if(trimString(form1.loan_range.value).length>100)
    {
         alert("輸入的貸款範圍不可超過100個字");
         return false;
    }

    return true;

}
//==============================政策性貸款-貸款對象資料畫面欄位CHECK=============
function Check_IX016W_ShowSub3(form1)
{
    if (form1.loan_poniter.value==''){
       alert("貨款對象不可空白");
       return false;
    }

    if(trimString(form1.loan_poniter.value).length>100)
    {
         alert("輸入的貸款對象不可超過100個字");
         return false;
    }

    return true;

}
//==============================政策性貸款-資金來源資料畫面欄位CHECK=============
function Check_IX016W_ShowSub4(form1)
{
    if (form1.way.value==''){
       alert("資金來源方式不可空白");
       return false;
    }
    if ( (form1.way.value!='2') && (form1.way_name.value=='')){
       alert("資金來源名稱不可空白");
       return false;
    }
    if(!checkNum1(form1.way_rate))
       return false;
    if(!NewcheckRate2(form1.way_rate))
    	return false;
    return true;

}
//==============================政策性貸款-辦理情形明細檔畫面欄位CHECK=============
function Check_IX016W_ShowSub5(form1)
{
    if (form1.m_year.value=='' || form1.m_month.value==''){
       alert("基准日不可空白");
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
//==============================政策性貸款-參貸銀行畫面欄位CHECK=============
function Check_IX016W_ShowSub6(form1)
{
    if (form1.bank_code.value==''){
       alert("參貸銀行不可空白");
       return false;
    }
    return true;

}
//=================指定用途信託資金投資國內外有價證券基本資料維護畫面欄位CHECK=============
function Check_WB001W(form1)
{
    if(form1.data_date.value =="null") {
       if((form1.S_Year.value =="") || (form1.S_Month.value =="")) {
               alert("基準日不可空白！");
               return(false);
       }
       if(!checkNum(form1.S_Year))
       return false;

    }
    if (form1.found_type.value =="") {
        alert("基金種類不可空白！");
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
//=================信託投資公司財務概況月報資料維護畫面欄位CHECK=============
function Check_WB002W(form1)
{
    if(form1.data_date.value =="null") {
       if((form1.S_Year.value =="") || (form1.S_Month.value =="")) {
               alert('基準日不可空白！');
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
//===銀行.信託投資公司對證券金融公司或證券商辦理資金融通明細基本資料維護畫面欄位CHECK=============
function Check_WB003W(form1)
{
    if(form1.data_date.value =='null') {
       if((form1.S_Year.value =='') || (form1.S_Month.value =='')) {
               alert('基準日不可空白！');
               return(false);
       }
       if(!checkNum(form1.S_Year))
       return false;

    }
    if (form1.loan_kind.value =='') {
        alert('借款公司屬性不可空白！');
        return(false);
    }
     if (form1.load_date_y.value=='' || form1.load_date_m.value=='' || form1.load_date_d.value=='') {
       alert("資金融通日期不可空白");
       return false;
    }
    if(!CheckMyDate(form1.load_date_y,form1.load_date_m,form1.load_date_d,form1.load_date,"資金融通日期"))
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
//===維護人員畫面欄位CHECK=============
function Check_Maintain(form1)
{
   if (trimString(form1.director_name.value)=="" ||  trimString(form1.maintain_dept.value)==""
       || trimString(form1.director_tel.value)=="" || trimString(form1.maintain_tel.value)==""){
        alert("維護部門名稱.主管姓名.主管電話.承辦員電話欄位不可空白！");
        return(false);
   }
   return true;
}

//===銀行兼證券商辦理有價證券融資業務所拆借款項資料維護畫面欄位CHECK=============
function Check_WB004W(form1)
{
    /*if(form1.data_date.value =='null') {
       if((form1.S_Year.value =='') || (form1.S_Month.value =='')) {
               alert('基準日不可空白！');
               return(false);
       }
       if(!checkNum(form1.S_Year))
       return false;

    }*/
     if (form1.bank_date_y.value=='' || form1.bank_date_m.value=='' || form1.bank_date_d.value=='') {
       alert("日期不可空白");
       return false;
    }
    if(!CheckMyDate(form1.bank_date_y,form1.bank_date_m,form1.bank_date_d,form1.bank_date,"日期"))
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

//===金融機構受託代收百貨卡消費帳款統計資料維護畫面欄位CHECK=============
function Check_WB007W(form1)
{
    if(form1.data_date.value =='null') {
       if((form1.S_Year.value =='') || (form1.S_Month.value =='')) {
               alert('基準日不可空白！');
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
//===金融機構受託代收百貨卡消費帳款統計資料維護畫面欄位CHECK=============
function IX015R(form,actionURL)
{
    if(form.s_year.value =="" || form.s_month.value == "" ) {
       alert("日期不可空白！");
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
    	if (code=="1") { //*歷次核准
           if(Check_IX005W_ShowSub1(form))
              form.submit();
        }
    	if (code=="2") {  //**高階主管
           if(Check_IX005W_ShowSubM(form))
              form.submit();
        }
    	if (code=="3") {  //**董監事
           if(Check_IX005W_ShowSubM1(form))
              form.submit();
        }
    	if (code=="4") {  //**派兼轉投資事業
           if(Check_IX005W_ShowSub2(form))
              form.submit();
        }

    }
}
//==============================轉投資事業基本資料維護CHECK=============
function Check_IX005W(form1)
{

    if (form1.const_type.value=='' || form1.business_id.value=='' || form1.business_name.value==''){
       alert("屬性.轉投資事業名稱.統一編號不可空白");
       return false;
    }

    if(!CheckMyDate(form1.exinvest_date_y,form1.exinvest_date_m,form1.exinvest_date_d,form1.exinvest_date,"核准原始投資日期"))
            return false;
    if(!CheckMyDate(form1.book_amount_datey,form1.book_amount_datem,form1.book_amount_dated,form1.book_amount_date,"帳面金額基準日"))
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
//==============================轉投資事業-財政部歷次核准投資及銀行實際投資金額資料畫面欄位CHECK=============
function Check_IX005W_ShowSub1(form1)
{

    if(!checkNum1(form1.apply_exinvest))
       return false;
    if(!checkNum1(form1.approve_exinvest_amount))
       return false;
    if(!CheckMyDate(form1.approve_exinvest_datey,form1.approve_exinvest_datem,form1.approve_exinvest_dated,form1.approve_exinvest_date,"核准增加投資日期"))
        return false;
    if(!checkNum1(form1.exinvest_amount))
       return false;
    if(!NewcheckRate(form1.ownstock_rate))
       return false;
    if(trimString(form1.approve_exinvest_no.value).length>20)
    {
         alert("輸入的財政部核准增加投資文號不可超過20個字");
         return false;
    }
    if(trimString(form1.reason.value).length>50)
    {
         alert("輸入的財政部核准增加投資金額與銀行實際投資金額相異時之原因不可超過50個字");
         return false;
    }

    return true;

}
//==============================轉投資事業-主管人員資料維護畫面欄位CHECK=============
function Check_IX005W_ShowSubM(form1)
{
    if (form1.position_code.value=='' || form1.name.value=='' ){
       alert("姓名.職稱不可空白");
       return false;
    }
    if (form1.induct_date_y.value=='' || form1.induct_date_m.value=='' || form1.induct_date_d.value=='') {
       alert("就任日期不可空白");
       return false;
    }
//
//    if(!CheckMyDate(form1.birth_date_y,form1.birth_date_m,form1.birth_date_d,form1.birth_date,"出生日期"))
//        return false;
    if(!CheckMyDate(form1.induct_date_y,form1.induct_date_m,form1.induct_date_d,form1.induct_date,"本屆就任日期"))
        return false;

    return true;

}
//==============================轉投資事業-董監事資料維護畫面欄位CHECK=============
function Check_IX005W_ShowSubM1(form1)
{
    if (form1.position_code.value=='' || form1.name.value=='' ){
       alert("姓名.職稱不可空白");
       return false;
    }
    if (form1.induct_date_y.value=='' || form1.induct_date_m.value=='' || form1.induct_date_d.value=='') {
       alert("就任日期不可空白");
       return false;
    }

//    if(!CheckMyDate(form1.birth_date_y,form1.birth_date_m,form1.birth_date_d,form1.birth_date,"出生日期"))
//        return false;
    if(!CheckMyDate(form1.induct_date_y,form1.induct_date_m,form1.induct_date_d,form1.induct_date,"就任日期"))
        return false;
    if(!CheckMyDate(form1.period_start_y,form1.period_start_m,form1.period_start_d,form1.period_start,"本屆任期迄日"))
        return false;
    if(!CheckMyDate(form1.period_end_y,form1.period_end_m,form1.period_end_d,form1.period_end,"本屆任期起日"))
        return false;

    if(trimString(form1.period_start.value)!="" && trimString(form1.period_end.value)!=""){

       if((Math.abs(form1.period_start_y.value) * 10000 + Math.abs(form1.period_start_m.value) * 100 + Math.abs(form1.period_start_d.value)) > (Math.abs(form1.period_end_y.value) * 10000 + Math.abs(form1.period_end_m.value) * 100 + Math.abs(form1.period_end_d.value)))
    	{
    		alert("本屆任期起始日期不可以大於終止日期");
    		return false;
    	}
    }
    if(!checkNum(form1.rank))
       return false;
    if(!checkNum(form1.appointed_num))
       return false;

    return true;

}
//==============================轉投資事業-派兼轉投資事業人員資料維護畫面欄位CHECK=============
function Check_IX005W_ShowSub2(form1)
{
    if (form1.position.value=='' || form1.name.value=='' ){
       alert("姓名.職稱不可空白");
       return false;
    }
//    if(!CheckMyDate(form1.birth_date_y,form1.birth_date_m,form1.birth_date_d,form1.birth_date,"出生日期"))
//        return false;

    return true;

}
//==================================外國銀行在華分支機構基本資料維護畫面欄位check=========//
function Check_FY004W(form1) {
  /*if (form1.branch_no.value.length ==0){
        alert('分支機構名稱不可空白');
        return false;
  }
  */
  if(form1.MyOpType.value == 'Abdicate_Confirm')
  {
        if(!confirm("確定要裁撤嗎？"))
            return false;
        if (trimString(form1.BN_DATE_Y.value) == "" )
		{
    	    alert("[裁撤生效日期 ](年)欄不可空白");
       	    form1.BN_DATE_Y.focus();
		    return false;
	    }

	    if (trimString(form1.BN_DATE_M.value) == "" )
	    {

		    alert("[裁撤生效日期 ](月)欄不可空白");
       	    form1.BN_DATE_M.focus();
		    return false;
	    }
	    if (trimString(form1.BN_DATE_D.value) == "" )
		{
			    alert("[裁撤生效日期 ](日)欄不可空白");
       	    	form1.BN_DATE_D.focus();
	   			return false;
	    }

        if(!CheckMyDate(form1.BN_DATE_Y,form1.BN_DATE_M,form1.BN_DATE_D,form1.BnDate,"裁撤生效日期"))
            return false;
   	    return true;
  }
  if(form1.MyOpType.value == 'Delete_Confirm')
  {
    if(!confirm("確定要刪除嗎？"))
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
	    	    alert("無分支機構代碼,無法新增");
	    	    form1.branch_no.focus();
	    	    return false;
	        }
	        if (trimString(form1.branch_no.value).length != 3 ||
	            isNaN(Math.abs(trimString(form1.branch_no.value))) )
	        {
	    	    alert("請輸入三位數字");
	    	    form1.branch_no.focus();
	    	    return false;
	        }
	    }
    }
  if (trimString(form1.bankname.value) == "" )
	{
		alert("分行名稱不可為空白");
		form1.bankname.focus();
		return false;
	}
	if(trimString(form1.bank_type.value) == "")
	{
		alert("分行類別不可為空白");
		form1.bank_type.focus();
		return false;
	}

  if(!CheckMyDate(form1.prv_yy,form1.prv_mm,form1.prv_dd,form1.prv_date,"核准設立日期"))
      return false;

  if(!CheckMyDate(form1.open_yy,form1.open_mm,form1.open_dd,form1.open_date,"設立日期"))
      return false;

  if(!CheckMyDate(form1.CHG_LICENSE_DATE_Y,form1.CHG_LICENSE_DATE_M,form1.CHG_LICENSE_DATE_D,form1.CHG_LICENSE_DATE,"最近換照日期"))
      return false;

  if(!CheckMyDate(form1.oper_yy,form1.oper_mm,form1.oper_dd,form1.oper_date,"開始營業日"))
      return false;

  if(trimString(form1.hsiend_id.value) == "")
	{
		alert("縣市名稱及郵遞區號不可為空白");
		form1.hsiend_id.focus();
		return false;
	}
	if (trimString(form1.addr.value) == "" )
	{
		alert("[地址] 不可空白");
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
      	   alert("營業項目名稱不可空白");
           return false;
        }
            form.submit();
    }
}
//======================================外國銀行在華辦事處資料檔欄位CHECK===
function Check_FY006W(form1) {

  if (trimString(form1.corp_id.value).length==0){
      alert('統一編號不可空白');
         return false;
    }
  if(!CheckMyDate(form1.prv_yy,form1.prv_mm,form1.prv_dd,form1.prv_date,"核准設立日期"))
      return false;
  if(!CheckMyDate(form1.setup_yy,form1.setup_mm,form1.setup_dd,form1.setup_date,"設立日期"))
      return false;
  if(!CheckMyDate(form1.CHG_LICENSE_DATE_Y,form1.CHG_LICENSE_DATE_M,form1.CHG_LICENSE_DATE_D,form1.CHG_LICENSE_DATE,"最近換照日期"))
      return false;

  if (form1.addr.value.length !=0 ) {
     if (form1.hsiend_id.value.length ==0) {
       alert('縣市名稱及郵遞區號不可空白');
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
    	alert("統一編號不可空白");
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

//==============================轉投資事業基本資料維護CHECK=============
function Check_OX004W(form1)
{

    if (form1.const_type.value=='' || form1.business_id.value=='' || form1.business_name.value==''){
       alert("屬性.轉投資事業名稱.統一編號不可空白");
       return false;
    }

    if(!CheckMyDate(form1.exinvest_date_y,form1.exinvest_date_m,form1.exinvest_date_d,form1.exinvest_date,"核准原始投資日期"))
            return false;
    if(!CheckMyDate(form1.book_amount_datey,form1.book_amount_datem,form1.book_amount_dated,form1.book_amount_date,"帳面金額基準日"))
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
//==============================轉投資事業-財政部歷次核准投資及銀行實際投資金額資料畫面欄位CHECK=============
function Check_OX004W_ShowSub1(form1)
{

    if(!checkNum1(form1.apply_exinvest))
       return false;
    if(!checkNum1(form1.approve_exinvest_amount))
       return false;
    if(!CheckMyDate(form1.approve_exinvest_datey,form1.approve_exinvest_datem,form1.approve_exinvest_dated,form1.approve_exinvest_date,"核准增加投資日期"))
        return false;
    if(!checkNum1(form1.exinvest_amount))
       return false;
    if(!NewcheckRate(form1.ownstock_rate))
       return false;
    if(trimString(form1.approve_exinvest_no.value).length>20)
    {
         alert("輸入的財政部核准增加投資文號不可超過20個字");
         return false;
    }
    if(trimString(form1.reason.value).length>50)
    {
         alert("輸入的財政部核准增加投資金額與銀行實際投資金額相異時之原因不可超過50個字");
         return false;
    }

    return true;

}
//==============================轉投資事業-主管人員資料維護畫面欄位CHECK=============
function Check_OX004W_ShowSubM(form1)
{
    if (form1.position_code.value=='' || form1.name.value=='' ){
       alert("姓名.職稱不可空白");
       return false;
    }
    if (form1.induct_date_y.value=='' || form1.induct_date_m.value=='' || form1.induct_date_d.value=='') {
       alert("就任日期不可空白");
       return false;
    }

//    if(!CheckMyDate(form1.birth_date_y,form1.birth_date_m,form1.birth_date_d,form1.birth_date,"出生日期"))
//        return false;
    if(!CheckMyDate(form1.induct_date_y,form1.induct_date_m,form1.induct_date_d,form1.induct_date,"本屆就任日期"))
        return false;

    return true;

}
//==============================轉投資事業-董監事資料維護畫面欄位CHECK=============
function Check_OX004W_ShowSubM1(form1)
{
    if (form1.position_code.value=='' || form1.name.value=='' ){
       alert("姓名.職稱不可空白");
       return false;
    }
    if (form1.induct_date_y.value=='' || form1.induct_date_m.value=='' || form1.induct_date_d.value=='') {
       alert("就任日期不可空白");
       return false;
    }

//    if(!CheckMyDate(form1.birth_date_y,form1.birth_date_m,form1.birth_date_d,form1.birth_date,"出生日期"))
//        return false;
    if(!CheckMyDate(form1.induct_date_y,form1.induct_date_m,form1.induct_date_d,form1.induct_date,"就任日期"))
        return false;
    if(!CheckMyDate(form1.period_start_y,form1.period_start_m,form1.period_start_d,form1.period_start,"本屆任期迄日"))
        return false;
    if(!CheckMyDate(form1.period_end_y,form1.period_end_m,form1.period_end_d,form1.period_end,"本屆任期起日"))
        return false;

    if(trimString(form1.period_start.value)!="" && trimString(form1.period_end.value)!=""){

       if((Math.abs(form1.period_start_y.value) * 10000 + Math.abs(form1.period_start_m.value) * 100 + Math.abs(form1.period_start_d.value)) > (Math.abs(form1.period_end_y.value) * 10000 + Math.abs(form1.period_end_m.value) * 100 + Math.abs(form1.period_end_d.value)))
    	{
    		alert("本屆任期起始日期不可以大於終止日期");
    		return false;
    	}
    }
    if(!checkNum(form1.rank))
       return false;
    if(!checkNum(form1.appointed_num))
       return false;

    return true;

}
//==============================轉投資事業-派兼轉投資事業人員資料維護畫面欄位CHECK=============
function Check_OX004W_ShowSub2(form1)
{
    if (form1.position.value=='' || form1.name.value=='' ){
       alert("姓名.職稱不可空白");
       return false;
    }
//    if(!CheckMyDate(form1.birth_date_y,form1.birth_date_m,form1.birth_date_d,form1.birth_date,"出生日期"))
//        return false;

    return true;

}
//==============================註銷金融機構維護畫面欄位CHECK=============
function Check_ZZ013W(form1)
{
    if (form1.bank_no.value=='' || form1.d_date_y.value=='' || form1.d_date_m.value=='' ||
       form1.d_date_d.value=='' || form1.setup_no.value=='' || form1.setup_date_y.value=='' ||
       form1.setup_date_m.value=='' || form1.setup_date_d.value=='' || form1.setup_reason.value==''){
       alert("畫面欄位全部不可空白");
       return false;
    }
    if(!CheckMyDate(form1.d_date_y,form1.d_date_m,form1.d_date_d,form1.d_date,"註銷日期"))
        return false;
    if(!CheckMyDate(form1.setup_date_y,form1.setup_date_m,form1.setup_date_d,form1.setup_date,"核準日期"))
        return false;

    return true;

}
//==============================營業項目代碼維護畫面欄位CHECK=============
function Check_ZZ011W(form1)
{
    if (form1.code.value=='' || form1.item.value==''){
       alert("畫面欄位全部不可空白");
       return false;
    }
    if(trimString(form1.item.value).length>70)
    {
         alert("輸入的營業名稱不可超過70個字");
         return false;
    }
    if(!checkNum(form1.code))
        return false;
    return true;

}
//serissa
//==============================政策性貸款-參貸銀行畫面欄位CHECK=============
function Check_ZZ010W_ShowSub(form)
{
    if (form.Function.value == "insert"){
        if (form.bank_no.value==''){
            alert("機構名稱不可空白");
            form.bank_no.focus();
            return false;
        }
    }
    if (form.negotiate_amount.value==''){
       alert("請輸入協商額度");
       form.negotiate_amount.focus();
       return false;
    }
    if(isNaN(Math.abs(changeVal(form.negotiate_amount))) )
    {
        alert("請輸入數字");
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
        alert("請輸入期別年度");
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
    		alert("起始年度不可以大於終止年度");
    		form.S_YEAR.focus()
    		return false;
    	}
    }
    return true;
}

