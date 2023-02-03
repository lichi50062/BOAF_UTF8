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
function DeleteAction(form1)
{
  if(AskDelete())
  {
    form1.Function.value = 'delete';
    form1.submit();
  }
}
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function leapYear (Year)
{
    if (((Year % 4)==0) && ((Year % 100)!=0) || ((Year % 400)==0))
        return (1);
    else
	return (0);
}
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
function ChangeAction(form,actionURL)
{
	form.action = actionURL;
	form.submit();
}

function Check_WB006W_Range(form)
{
    var sDate;
    var eDate;
    if(form.S_YEAR.value.length != 0 && form.S_MONTH.value.length != 0)
    {
        sDate = Math.abs(form.S_YEAR.value);
        if(isNaN(sDate))
        {
            alert("起始年不可輸入文字");
            return false;
        }
        else
        {
            if(sDate != 0)
                sDate += form.S_MONTH.value;
            else
            {
                alert("起始年份不可為0，也不可為空白");
                return false;
            }
        }
    }
    else if(form.S_YEAR.value.length == 0 && form.S_MONTH.value.length == 0)
            sDate = "";
         else
         {
            alert("請輸入完整的查詢起始時間");
            return false;
          }

    //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    if(form.E_YEAR.value.length != 0 && form.E_MONTH.value.length != 0)
    {
        eDate = Math.abs(form.E_YEAR.value);
        if(isNaN(eDate))
        {
            alert("終止年不可輸入文字");
            return false;
        }
        else
        {
            if(eDate != 0)
                eDate += form.E_MONTH.value;
            else
            {
                alert("終止年份不可為0，也不可為空白");
                return false;
            }
        }
    }
    else if(form.E_YEAR.value.length == 0 && form.E_MONTH.value.length == 0)
            eDate = "";
         else
         {
            alert("請輸入完整的查詢終止時間");
            return false;
          }

    //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

        if(sDate != "" && eDate != "")
        {
            if(Math.abs(sDate) > Math.abs(eDate))
            {
                alert("終止年月必須大於或等於起始年月");
                return false;
            }
        }
    //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        if(sDate != "")
        {
            if(sDate.length == 3)
                sDate = "00" + sDate;
            else if(sDate.length == 4)
                    sDate = "0" + sDate;
                  else if(sDate.length < 3)
                    sDate = "000" + sDate;
        }
        if(eDate != "")
        {
            if(eDate.length == 3)
                eDate = "00" + eDate;
            else if(eDate.length == 4)
                    eDate = "0" + eDate;
                 else if(eDate.length < 3)
                        eDate = "000" + eDate;
        }
        form.S_DATE.value = sDate;
        form.E_DATE.value = eDate;
        return true;
}
//=====================================================================
function changeVal(T1)
{
    pos=0
    var oring=T1.value;
    pos=oring.indexOf(",");

    while (pos !=-1)
    {
        oring=(oring).replace(",","");
        pos=oring.indexOf(",");
    }

    return oring;
}

//================================================================
function changeStr(T1)
{
    T1.value = changeVal(T1);
    //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
//===========================================================================
function checkRate(Rate1)
{
    if(isNaN(Math.abs(Rate1.value)))
    {
   	    alert("請輸入數字");
   	    return;
    }
    if(eval(Rate1.value) > 99.99 || eval(Rate1.value) < -99.99)
    {
        alert("利率不可大於 99.99, 也不可小於 -99.99")
        return;
    }
    if((Rate1.value).indexOf(".") != -1 )
    {
        len = (Rate1.value).substring((Rate1.value).indexOf(".")+1,Rate1.value.length);
        if(len.length > 2)
        {
            alert("小數點後只能有兩個位數");
            return;
        }
        Rate1.value = eval(Rate1.value);
    }
}
//===========================================================================
function checkRate1(Rate1)
{
    if(isNaN(Math.abs(Rate1.value)))
    {
   	    alert("請輸入數字");
   	    return;
    }
    if(eval(Rate1.value) > 100.00 || eval(Rate1.value) < -99.99)
    {
        alert("利率不可大於 100.00, 也不可小於 -99.99")
        return;
    }
    if((Rate1.value).indexOf(".") != -1 )
    {
        len = (Rate1.value).substring((Rate1.value).indexOf(".")+1,Rate1.value.length);
        if(len.length > 2)
        {
            alert("小數點後只能有兩個位數");
            return;
        }
        Rate1.value = eval(Rate1.value);
    }
    return true;
}
//===========================================================================
function checkRate3(Rate1)
{
    if(isNaN(Math.abs(Rate1.value)))
    {
   	    alert("請輸入數字");
   	    return;
    }
    if(eval(Rate1.value) > 99.999 || eval(Rate1.value) < -99.999)
    {
        alert("利率不可大於 99.999, 也不可小於 -99.999")
        return;
    }
    if((Rate1.value).indexOf(".") != -1 )
    {
        len = (Rate1.value).substring((Rate1.value).indexOf(".")+1,Rate1.value.length);
        if(len.length > 3)
        {
            alert("小數點後只能有三個位數");
            return;
        }
        Rate1.value = eval(Rate1.value);
    }
}
//===========================================================================
function checkRate4(Rate1)
{
    if(isNaN(Math.abs(Rate1.value)))
    {
   	    alert("請輸入數字");
   	    return;
    }
    if(eval(Rate1.value) > 999.999 || eval(Rate1.value) < -999.999)
    {
        alert("利率不可大於 999.999, 也不可小於 -999.999")
        return;
    }
    if((Rate1.value).indexOf(".") != -1 )
    {
        len = (Rate1.value).substring((Rate1.value).indexOf(".")+1,Rate1.value.length);
        if(len.length > 3)
        {
            alert("小數點後只能有三個位數");
            return;
        }
        Rate1.value = eval(Rate1.value);
    }
    return true;
}
//============================================================================
function writeRate(form)
{
    if(form.neg_amt.value!='' && form.neg_amt.value!='0' && form.emonth_amt.value!='')
       form.fin_rate.value = changeVal(form.emonth_amt) / changeVal(form.neg_amt);
}
//==============================================================================
function writeAmt(form)
{
    if(form.neg_amt.value != '' && form.emonth_amt.value!='')
       form.les_amt.value =   changeVal(form.neg_amt) - changeVal(form.emonth_amt);
}
//============================================================================
function Check_WB006W_ShowInsert(form)
{
    if(form.Function.value == 'delete')
            return true;

    if(form.Function.value == 'insert')
    {
        if(form.periodyear.selectedIndex < 0)
        {
        	alert("沒有期別年度");
        	return false;
        }

        if(trimString(form.S_Year.value) == '')
        {
            alert("基準日不可空白");
            return false;
        }
        if(!CheckYear(form.S_Year))
            return false;

        bdate = (form.hidebegdate[form.periodyear.selectedIndex].value).substring(0,6);
        edate = (form.hideenddate[form.periodyear.selectedIndex].value).substring(0,6);

         mdate = (eval(form.S_Year.value)+1911) + form.S_Month.value;

        if(eval(bdate)>eval(mdate) || eval(edate)<eval(mdate))
        {
            alert("基準日不在期別年度的範圍內");
            return false;
        }
    }

  //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    if(isNaN(Math.abs(changeVal(form.lmonth_cnt))) || isNaN(Math.abs(changeVal(form.lmonth_amt))) ||
       isNaN(Math.abs(changeVal(form.tmonth_cnt))) || isNaN(Math.abs(changeVal(form.tmonth_amt))) ||
       isNaN(Math.abs(changeVal(form.emonth_cnt))) || isNaN(Math.abs(changeVal(form.emonth_amt))) )
    {
        alert("貸款資料表格中請輸入數字");
        return false;
    }
//serissa
//    if(isNaN(Math.abs(form.basic_rate.value)))
//    {
//        alert("基本放款率中請輸入數字");
//        return false;
//    }

//    for(i=0;i<4;i++)
//    {
//
//
//        if(isNaN(Math.abs(form.fin_rate11.value))   || isNaN(Math.abs(form.fin_rate2[i].value)) ||
//           isNaN(Math.abs(form.house_rate1[i].value)) || isNaN(Math.abs(form.house_rate2[i].value)))
//        {
//            alert("本年度貸款利率政策表格中請輸入數字");
//            return false;
//        }
//    }

//    for(i=0;i<4;i++)
//    {
//        if(eval(form.fin_rate1[i].value) > 99.99 || eval(form.fin_rate1[i].value) < -99.99)
//        {
//            alert("利率不可大於 99.99, 也不可小於 -99.99")
//            return false;
//        }
//        if(eval(form.fin_rate2[i].value) > 99.99 || eval(form.fin_rate2[i].value) < -99.99)
//        {
//            alert("利率不可大於 99.99, 也不可小於 -99.99")
//            return false;
//        }
//        if(eval(form.house_rate1[i].value) > 99.99 || eval(form.house_rate1[i].value) < -99.99)
//        {
//            alert("利率不可大於 99.99, 也不可小於 -99.99")
//            return false;
//        }
//        if(eval(form.house_rate1[i].value) > 99.99 || eval(form.house_rate1[i].value) < -99.99)
//        {
//            alert("利率不可大於 99.99, 也不可小於 -99.99")
//            return false;
//        }
//    }
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

//    for(i=0;i<4;i++)
//    {
//        if((form.fin_rate1[i].value).indexOf(".") != -1 )
//        {
//            len = (form.fin_rate1[i].value).substring((form.fin_rate1[i].value).indexOf(".")+1,form.fin_rate1[i].value.length);
//            if(len.length > 2)
//            {
//                alert("利率的小數點後只能有兩個位數");
//                return false;
//            }
//        }
//        if((form.fin_rate2[i].value).indexOf(".") != -1 )
//        {
//            len = (form.fin_rate2[i].value).substring((form.fin_rate2[i].value).indexOf(".")+1,form.fin_rate2[i].value.length);
//            if(len.length > 2)
//            {
//                alert("利率的小數點後只能有兩個位數");
//                return false;
//            }
//         }
//         if((form.house_rate1[i].value).indexOf(".") != -1 )
//         {
//            len = (form.house_rate1[i].value).substring((form.house_rate1[i].value).indexOf(".")+1,form.house_rate1[i].value.length);
//            if(len.length > 2)
//            {
//                alert("利率的小數點後只能有兩個位數");
//                return false;
//            }
//         }
//         if((form.house_rate2[i].value).indexOf(".") != -1 )
//         {
//            len = (form.house_rate2[i].value).substring((form.house_rate2[i].value).indexOf(".")+1,form.house_rate2[i].value.length);
//            if(len.length > 2)
//            {
//                alert("利率的小數點後只能有兩個位數");
//                return false;
//            }
//         }
//    }

  //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    if(isNaN(Math.abs(form.basic_rate.value)) || isNaN(Math.abs(form.pm_rate1.value)) ||
       isNaN(Math.abs(form.pm_rate2.value))   || isNaN(Math.abs(form.pm_rate3.value)) ||
       isNaN(Math.abs(form.n_pm_rate4.value))   || isNaN(Math.abs(form.n_pm_rate1.value)) ||
       isNaN(Math.abs(form.n_pm_rate2.value))   || isNaN(Math.abs(form.n_pm_rate3.value)) ||
       isNaN(Math.abs(form.pm_rate4.value)))
    {
            alert("本年度貸款利率政策表格中請輸入數字");
            return false;
    }
    if(eval(form.basic_rate.value) > 100.00 || eval(form.basic_rate.value) < -99.99)
    {
        alert("基本放款率不可大於 100.00, 也不可小於 -99.99")
        return false;
    }
    if(eval(form.pm_rate1.value) > 100.00 || eval(form.pm_rate1.value) < -99.99 ||
       eval(form.pm_rate2.value) > 100.00 || eval(form.pm_rate2.value) < -99.99 ||
       eval(form.pm_rate3.value) > 100.00 || eval(form.pm_rate3.value) < -99.99 ||
       eval(form.pm_rate4.value) > 100.00 || eval(form.pm_rate4.value) < -99.99 )
    {
        alert("基本放款率加減百分比不可大於 100.00, 也不可小於 -99.99")
        return false;
    }


    if((form.basic_rate.value).indexOf(".") != -1 )
    {
        len = (form.basic_rate.value).substring((form.basic_rate.value).indexOf(".")+1,form.basic_rate.value.length);
        if(len.length > 3)
        {
            alert("基本放款率的小數點後只能有三個位數");
            return false;
        }
    }

    if((form.pm_rate1.value).indexOf(".") != -1)
    {
        len = (form.pm_rate1.value).substring((form.pm_rate1.value).indexOf(".")+1,form.pm_rate1.value.length);
        if(len.length > 3)
        {
            alert("第一年基本放款率加減百分比的小數點後只能有三個位數");
            return false;
        }
    }

    if((form.pm_rate2.value).indexOf(".") != -1)
    {
        len = (form.pm_rate2.value).substring((form.pm_rate2.value).indexOf(".")+1,form.pm_rate2.value.length);
        if(len.length > 3)
        {
            alert("第二年基本放款率加減百分比的小數點後只能有三個位數");
            return false;
        }
    }

     if((form.pm_rate3.value).indexOf(".") != -1)
    {
        len = (form.pm_rate3.value).substring((form.pm_rate3.value).indexOf(".")+1,form.pm_rate3.value.length);
        if(len.length > 3)
        {
            alert("第三年基本放款率加減百分比的小數點後只能有三個位數");
            return false;
        }
    }

     if((form.pm_rate4.value).indexOf(".") != -1)
    {
        len = (form.pm_rate4.value).substring((form.pm_rate4.value).indexOf(".")+1,form.pm_rate4.value.length);
        if(len.length > 3)
        {
            alert("第四年基本放款率加減百分比的小數點後只能有三個位數");
            return false;
        }
    }

    if(!Check_Maintain(form))
    	return false;
    return true;

}
//=====================================================
function doSubmit(form,myfun)
{
    form.Function.value = myfun;
    if(form.Function.value=='delete')
    {
        if(AskDelete())
            form.submit();
    }
    else if(form.Function.value=='update')
    {
        if(Check_WB006W_ShowInsert(form))
            form.submit();
    }
}
//==================================================
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
//============================================================
function Check_IX017W_ShowInsert(form)
{
    if(form.Function.value == 'delete')
            return true;
    if(form.Function.value == 'insert')
    {
        if(trimString(form.S_Year.value)=="")
        {
            alert("年份不可為空");
            return false;
        }
        if(!CheckYear(form.S_Year))
            return false;
    }
    if(isNaN(Math.abs(changeVal(form.guaranteemoney))))
    {
        alert("保證金請輸入數字");
        return false;
    }
    if(trimString(form.proposerno.value)=="")
    {
        alert("開狀(證)申請人統一編號不可為空");
        return false;
    }
    if(trimString(form.subsidiary.value).length>40)
    {
         alert("輸入的所擔保之大陸子公司不可超過40個中文字");
         return false;
    }
    if(trimString(form.investmentno.value).length>25)
    {
         alert("輸入的經濟部許可赴大陸地區投資文號不可超過25個全形中文字長度");
         return false;
    }
    if(!Check_Maintain(form))
    	return false;
    return true;
}
//=====================================================
function doSubmitIX017W(form,myfun)
{
    form.Function.value = myfun;
    if(form.Function.value=='delete')
    {
        if(AskDelete())
            form.submit();
    }
    else if(form.Function.value=='update')
    {
        if(Check_IX017W_ShowInsert(form))
            form.submit();
    }
}
//==================================================
function check_IX011R(form)
{
    if(Check_WB006W_Range(form))
    {
        if(trimString(form.S_YEAR.value)!="" && form.S_MONTH.value.length!=0 &&
           trimString(form.E_YEAR.value)!="" && form.E_MONTH.value.length!=0)
        {
            //return true;
        }
        else
        {
            alert("請完整輸入起迄年月");
            return false;
        }
		var iYear = Math.abs(form.E_YEAR.value) - Math.abs(form.S_YEAR.value)-1;
		iYear = (iYear<=0)?0:iYear*12;
		var iMonth = (13-Math.abs(form.S_MONTH.value)) + Math.abs(form.E_MONTH.value);
		if(iYear+iMonth > 24)
		{
			alert("起迄明細區間不可大於24個月");
			return false;
		}
    }
    else
    	return false;
    return true;
}
//====================================================
function ChangeAction(form,actionURL)
{
    form.action = actionURL;
    if(check_IX011R(form))
	    form.submit();
}
//===================================================
function AskDelete()
{
  if(confirm("確定要刪除嗎？"))
    return true;
  else
    return false;
}
//====================================================
function ChangeNegAmt(form)
{
    var bd = form.hidebegdate[form.periodyear.selectedIndex].value;
    var ed = form.hideenddate[form.periodyear.selectedIndex].value;
    form.shownegamt.value = form.hidenegamt[form.periodyear.selectedIndex].value;
    form.neg_amt.value = form.hidenegamt[form.periodyear.selectedIndex].value;
    writeRate(form);
    writeAmt(form);
    form.showbegyear.value = eval(bd.substring(0,4))-1911;
    form.showbegmonth.value = eval(bd.substring(4,6));
    form.showbegday.value = eval(bd.substring(6));
    form.showendyear.value = eval(ed.substring(0,4))-1911;
    form.showendmonth.value = eval(ed.substring(4,6));
    form.showendday.value = eval(ed.substring(6));
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
//========================================================
function Check_FY001W(form)
{
    if(!NewcheckRate(form.captrate))
        return false;
    if(!CheckYear(form.captdateyear))
            return false;
    if((trimString(form.captdateyear.value)!="" && trimString(form.captdatemonth.value)=="") ||
       (trimString(form.captdateyear.value)=="" && trimString(form.captdatemonth.value)!=""))
    {
        alert("請輸入完整的年月或者年月都不輸入");
        return false;
    }
    else if(trimString(form.captdateyear.value)!="" && trimString(form.captdatemonth.value)!="")
    {
        if(form.captdateyear.value.length==1)
            form.captdate.value = "00"+form.captdateyear.value + form.captdatemonth.value;
        else if(form.captdateyear.value.length==2)
            form.captdate.value = "0" + form.captdateyear.value + form.captdatemonth.value;
            else
                form.captdate.value = form.captdateyear.value + form.captdatemonth.value;
    }
    if(trimString(form.corpename.value).length>120)
    {
         alert("輸入的所屬集團英文名稱不可超過120個字");
         return false;
    }
    if(!Check_Maintain(form))
    	return false;
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
function NewcheckRate100(Rate1)
{
    if(isNaN(Math.abs(Rate1.value)))
    {
   	    alert("請輸入數字");
   	    return false;
    }
    if(eval(Rate1.value) > 999.99 || eval(Rate1.value) < -999.99)
    {
        alert("利率不可大於 999.99, 也不可小於 -999.99")
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
//====================================================================
function LinkToInsert(form,linkaddr)
{
    linkaddr = "/servlet/" + linkaddr;
    if(form.Function.value=='insert')
    {
        alert("請先新增相關資料");
        return;
    }
    else
        window.open(linkaddr,'_self');
}
//========================================================
function Check_FY001W_ShowSubR(form)
{
//    if(!NewcheckRate(form.revrate))
    if(!checkRate1(form.revrate))
        return false;
    return true;
}
//===========================================================================
function doFY001WSubmit(form,myfun)
{
    form.Function.value = myfun;
    if(myfun=="delete")
    {
        if(AskDelete())
            form.submit();
    }
    else
    {
        if(Check_FY001W_ShowSubR(form))
            form.submit();
    }
}
//========================================================
function Check_FY002W(form)
{
// jakted 2001/10/31
//    if(!NewcheckRate(form.captrate))
//        return false;
//    if(!NewcheckRate(form.roerate))
//        return false;
//    if(!NewcheckRate100(form.cirate))
//        return false;

      if(!checkRate1(form.captrate))
          return false;
      if(!checkRate1(form.roerate))
          return false;
      if(!checkRate1(form.cirate))
          return false;

    if(!CheckYear(form.captdateyear))
            return false;
    if((trimString(form.captdateyear.value)!="" && trimString(form.captdatemonth.value)=="") ||
       (trimString(form.captdateyear.value)=="" && trimString(form.captdatemonth.value)!=""))
    {
        alert("請輸入完整的年月或者年月都不輸入");
        return false;
    }
    else if(trimString(form.captdateyear.value)!="" && trimString(form.captdatemonth.value)!="")
    {
        if(form.captdateyear.value.length==1)
            form.captdate.value = "00"+form.captdateyear.value + form.captdatemonth.value;
        else if(form.captdateyear.value.length==2)
            form.captdate.value = "0" + form.captdateyear.value + form.captdatemonth.value;
            else
                form.captdate.value = form.captdateyear.value + form.captdatemonth.value;
    }

    if(!CheckYear(form.setupdateyear))
            return false;
    if(trimString(form.setupdateyear.value)!="" && trimString(form.setupdatemonth.value)!="" && trimString(form.setupdateday.value)!="")
    {
        //var setyear = (Math.abs(form.setupdateyear.value)) + 1911;
        var setyear = form.setupdateyear.value;
        if(!fnValidDate(setyear + form.setupdatemonth.value + form.setupdateday.value))
        {
            alert('創立日期有誤');
            return false;
        }
        form.setupdate.value = setyear + form.setupdatemonth.value + form.setupdateday.value;
    }
    else if(trimString(form.setupdateyear.value)=="" && trimString(form.setupdatemonth.value)=="" && trimString(form.setupdateday.value)==""){}
         else
         {
            alert("請輸入完整的年月或者年月都不輸入");
            return false;
         }

    if(trimString(form.policymemo.value).length>125)
    {
         alert("輸入的重要政策概述不可超過125個字");
         return false;
    }
    if(trimString(form.impmemo.value).length>125)
    {
         alert("輸入的重大事項概述不可超過125個字");
         return false;
    }

    if(!Check_Maintain(form))
    	return false;
    //alert(form.setupdate.value);
    return true;
}
//===========================================================================
function doFY002WSubmit(form,myfun)
{
    form.Function.value = myfun;
    if(myfun=="delete")
    {
        if(AskDelete())
            form.submit();
    }
    else
    {
        if(!NewcheckRate(form.salrate))
            return false;
        if(form.mytype.value == "2")
        {
            if(!NewcheckRate(form.mytype))
                return false;
        }

        form.submit();
    }
}
//========================================================
function Check_FY003W(form)
{
    if(trimString(form.opfund.value)=="")
    {
        alert("營運所用資金不可為空");
        return false;
    }
    if(!Check_Maintain(form))
    	return false;
    return true;
}
//===========================================================================
function doFY003WSubmit(form,myfun)
{
    form.Function.value = myfun;
    if(myfun=="delete")
    {
        if(AskDelete())
            form.submit();
    }
    else
    {
        if(trimString(form.businessname.value)=="")
        {
            alert("營業項目不可為空");
            return false;
        }
        form.submit();
    }
}
//=============================================================================
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
        checkdate.value = setyear + checkmonth.value + checkday.value;
    }
    else if(trimString(checkyear.value)=="" && trimString(checkmonth.value)=="" && trimString(checkday.value)==""){}
         else
         {
            alert("請輸入完整的"+ szWarning+ "或者都不輸入");
            return false;
         }
    return true
}
//=============================================================================
function check_Addr(checkhsien,checkarea,checkaddr,szWarning)
{
    if((trimString(checkhsien.value)=="" && trimString(checkarea.value)=="" && trimString(checkaddr.value)=="")  || (trimString(checkhsien.value)!="" && trimString(checkarea.value)!="" && trimString(checkaddr.value)!=""))
        return true;
    else
    {
        alert("請輸入完整的" + szWarning + "地址或者都不輸入");
        return false;
    }

}
//==============================================================================
function Check_WG001W(form)
{
    if(!CheckMyDate(form.setupyear,form.setupmonth,form.setupday,form.setupdate,"設立日期"))
            return false;
    if(!CheckMyDate(form.chglicenseyear,form.chglicensemonth,form.chglicenseday,form.chglicensedate,"最近換照日期"))
            return false;
    if(!CheckMyDate(form.startyear,form.startmonth,form.startday,form.startdate,"開始營業日期"))
            return false;

    if(!check_Addr(form.hsienid,form.areaid,form.addr,"總機構"))
            return false;
    if(!check_Addr(form.ithsienid,form.itareaid,form.itaddr,"資訊單位"))
            return false;
    if(!check_Addr(form.audithsienid,form.auditareaid,form.auditaddr,"稽核單位"))
            return false;
    return true;

}
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
//==============================================================================
function Check_BX002W_ShowInsert(form)
{
    if(form.branchno.value=="")
    {
        alert("請輸入分支機構代號及名稱");
        return false;
    }
    if(!CheckMyDate(form.setupyear,form.setupmonth,form.setupday,form.setupdate,"設立日期"))
            return false;
    if(!CheckMyDate(form.setupnoyear,form.setupnomonth,form.setupnoday,form.setupnodate,"原始核准設立日期"))
            return false;
    if(!CheckMyDate(form.chglicenseyear,form.chglicensemonth,form.chglicenseday,form.chglicensedate,"最近換照日期"))
            return false;
    if(!CheckMyDate(form.startyear,form.startmonth,form.startday,form.startdate,"開始營業日期"))
            return false;

    if(!check_Addr(form.hsienid,form.areaid,form.addr,""))
            return false;
    return true;

}
//=============================================================================
function CheckQueryDate(checkyear,checkmonth,checkday,checkdate,szWarning)
{
    if(!CheckYear(checkyear))
            return false;
    if(trimString(checkyear.value)!="" && checkmonth.options[checkmonth.selectedIndex].value!="" && checkday.options[checkday.selectedIndex].value!="")
    {
        checkdate.value = Math.abs(checkyear.value) * 1000 + Math.abs(checkmonth.options[checkmonth.selectedIndex].value) * 10 + Math.abs(checkday.options[checkday.selectedIndex].value);
    }
    else if(trimString(checkyear.value)=="" && checkmonth.options[checkmonth.selectedIndex].value=="" && checkday.options[checkday.selectedIndex].value==""){}
         else
         {
            alert("請輸入完整的"+ szWarning+ "或者都不輸入");
            return false;
         }
    return true
}
//=========================================================
function Check_OT001W(form)
{
    if(!CheckQueryDate(form.S_YEAR,form.S_MONTH,form.S_WEEK,form.S_DATE,"起始日期"))
            return false;
    if(!CheckQueryDate(form.E_YEAR,form.E_MONTH,form.E_WEEK,form.E_DATE,"結束日期"))
            return false;
	if(trimString(form.S_DATE.value)!="" && trimString(form.E_DATE.value)!="")
	{
		if(Math.abs(form.S_DATE.value) > Math.abs(form.E_DATE.value))
    	{
    		alert("起始日期不可以大於終止日期");
    		return false;
    	}
    	return true;
    }
}
//===============================================================================
function Check_OT001W_ShowInsert(form)
{
    if(form.Function.value == 'delete')
            return true;
    if(form.Function.value == 'insert')
    {
        if(trimString(form.S_Year.value)=="")
        {
            alert("年份不可為空");
            return false;
        }
        if(!CheckYear(form.S_Year))
            return false;
    }
    if(isNaN(Math.abs(changeVal(form.apply5num1))) || isNaN(Math.abs(changeVal(form.apply5amount1))) ||
       isNaN(Math.abs(changeVal(form.apply5num2))) || isNaN(Math.abs(changeVal(form.apply5amount2))) ||
       isNaN(Math.abs(changeVal(form.apply5num3))) || isNaN(Math.abs(changeVal(form.apply5amount3))) ||
       isNaN(Math.abs(changeVal(form.apply56numla))) || isNaN(Math.abs(changeVal(form.apply56amountla))) ||
       isNaN(Math.abs(changeVal(form.apply56numlh))) || isNaN(Math.abs(changeVal(form.apply56amountlh))) ||
       isNaN(Math.abs(changeVal(form.apply56num2a))) || isNaN(Math.abs(changeVal(form.apply56amount2a))) ||
       isNaN(Math.abs(changeVal(form.apply56num2h))) || isNaN(Math.abs(changeVal(form.apply56amount2h))) ||
       isNaN(Math.abs(changeVal(form.apply56num3a))) || isNaN(Math.abs(changeVal(form.apply56amount3a))) ||
       isNaN(Math.abs(changeVal(form.apply56num3h))) || isNaN(Math.abs(changeVal(form.apply56amount3h))) ||
       isNaN(Math.abs(changeVal(form.creditnum1))) || isNaN(Math.abs(changeVal(form.creditamount1))) ||
       isNaN(Math.abs(changeVal(form.creditnum2))) || isNaN(Math.abs(changeVal(form.creditamount2))) ||
       isNaN(Math.abs(changeVal(form.creditnum3))) || isNaN(Math.abs(changeVal(form.creditamount3))) ||
       isNaN(Math.abs(changeVal(form.capply6num))) || isNaN(Math.abs(changeVal(form.capply6amt))) ||
       isNaN(Math.abs(changeVal(form.ccheck6num))) || isNaN(Math.abs(changeVal(form.ccheck6amt))))
    {
        alert("請輸入數字");
        return false;
    }
    if(!Check_Maintain(form))
    	return false;
    return true;
}
//====================================================
function doSubmitOT001W(form,myfun)
{
    form.Function.value = myfun;
    if(form.Function.value=='delete')
    {
        if(AskDelete())
            form.submit();
    }
    else if(form.Function.value=='update')
    {
        if(Check_OT001W_ShowInsert(form))
            form.submit();
    }
}
//=================================================================
function CheckQueryDate2(checkyear,checkmonth,checkdate,szWarning)
{
    if(!CheckYear(checkyear))
            return false;
    if(trimString(checkyear.value)!="" && checkmonth.options[checkmonth.selectedIndex].value!="")
    {
        checkdate.value = Math.abs(checkyear.value) * 100 + Math.abs(checkmonth.options[checkmonth.selectedIndex].value);
    }
    else if(trimString(checkyear.value)=="" && checkmonth.options[checkmonth.selectedIndex].value==""){}
         else
         {
            alert("請輸入完整的"+ szWarning+ "或者都不輸入");
            return false;
         }
    return true
}
//=========================================================
function Check_OT002W(form)
{
    if(!CheckQueryDate2(form.S_YEAR,form.S_MONTH,form.S_DATE,"起始日期"))
            return false;
    if(!CheckQueryDate2(form.E_YEAR,form.E_MONTH,form.E_DATE,"結束日期"))
            return false;
	if(trimString(form.S_DATE.value)!="" && trimString(form.E_DATE.value)!="")
	{
		if(Math.abs(form.S_DATE.value) > Math.abs(form.E_DATE.value))
    	{
    		alert("起始日期不可以大於終止日期");
    		return false;
    	}
    }
    return true;
}
//===============================================================================
function Check_OT002W_ShowInsert(form)
{
    if(form.Function.value == 'delete')
            return true;
    if(form.Function.value == 'insert')
    {
        if(trimString(form.S_Year.value)=="")
        {
            alert("年份不可為空");
            return false;
        }
        if(!CheckYear(form.S_Year))
            return false;
    }
    if(isNaN(Math.abs(changeVal(form.ix141))) || isNaN(Math.abs(changeVal(form.ix142))) ||
       isNaN(Math.abs(changeVal(form.ix143))) || isNaN(Math.abs(changeVal(form.ix144))) ||
       isNaN(Math.abs(changeVal(form.ix145))) || isNaN(Math.abs(changeVal(form.ix146))) ||
       isNaN(Math.abs(changeVal(form.ix147))) || isNaN(Math.abs(changeVal(form.ix1471))) ||
       isNaN(Math.abs(changeVal(form.ix1472))) || isNaN(Math.abs(changeVal(form.ix148))) ||
       isNaN(Math.abs(changeVal(form.ix149))) || isNaN(Math.abs(changeVal(form.ix1410))) ||
       isNaN(Math.abs(changeVal(form.ix1412))) ||
       isNaN(Math.abs(changeVal(form.ix1413))) || isNaN(Math.abs(changeVal(form.ix1414))))
    {
        alert("請輸入數字");
        return false;
    }
    if(isNaN(Math.abs(form.ix1411.value)))
    {
        alert("本月底逾放比率中請輸入數字");
        return false;
    }
    if(eval(form.ix1411.value) > 99.99 || eval(form.ix1411.value) < -99.99)
    {
        alert("本月底逾放比率不可大於 99.99, 也不可小於 -99.99")
        return false;
    }

    if((form.ix1411.value).indexOf(".") != -1 )
    {
        len = (form.ix1411.value).substring((form.ix1411.value).indexOf(".")+1,form.ix1411.value.length);
        if(len.length > 3)
        {
            alert("本月底逾放比率小數點後只能有三個位數");
            return false;
        }
    }

    if(!Check_Maintain(form))
    	return false;
    return true;
}
//====================================================
function doSubmitOT002W(form,myfun)
{
    form.Function.value = myfun;
    if(form.Function.value=='delete')
    {
        if(AskDelete())
            form.submit();
    }
    else if(form.Function.value=='update')
    {
        if(Check_OT002W_ShowInsert(form))
            form.submit();
    }
}
//===============================================================================
function AddNumber(form,indexp)
{
    var myamount = 0
    if(indexp==1)
    {
        if(trimString(form.ix08002.value)!="" && !isNaN(Math.abs(changeVal(form.ix08002))))
            myamount += eval(changeVal(form.ix08002));
        if(trimString(form.ix08003.value)!="" && !isNaN(Math.abs(changeVal(form.ix08003))))
            myamount += eval(changeVal(form.ix08003));
        if(trimString(form.ix08004.value)!="" && !isNaN(Math.abs(changeVal(form.ix08004))))
            myamount += eval(changeVal(form.ix08004));
        if(trimString(form.ix08005.value)!="" && !isNaN(Math.abs(changeVal(form.ix08005))))
            myamount += eval(changeVal(form.ix08005));
        if(trimString(form.ix08006.value)!="" && !isNaN(Math.abs(changeVal(form.ix08006))))
            myamount += eval(changeVal(form.ix08006));
		if(form.ix08a1 != null)
			if(trimString(form.ix08a1.value)!="" && !isNaN(Math.abs(changeVal(form.ix08a1))))
            	myamount += eval(changeVal(form.ix08a1));
        form.ix08001s.value = myamount;
        form.ix08001.value = myamount;
    }else if(indexp==2){
        if(trimString(form.ix08008.value)!="" && !isNaN(Math.abs(changeVal(form.ix08008))))
            myamount += eval(changeVal(form.ix08008));
        if(trimString(form.ix08009.value)!="" && !isNaN(Math.abs(changeVal(form.ix08009))))
            myamount += eval(changeVal(form.ix08009));
        if(trimString(form.ix08010.value)!="" && !isNaN(Math.abs(changeVal(form.ix08010))))
            myamount += eval(changeVal(form.ix08010));
        if(trimString(form.ix08011.value)!="" && !isNaN(Math.abs(changeVal(form.ix08011))))
            myamount += eval(changeVal(form.ix08011));
        if(trimString(form.ix08012.value)!="" && !isNaN(Math.abs(changeVal(form.ix08012))))
            myamount += eval(changeVal(form.ix08012));
		if(form.ix08a2 != null)
			if(trimString(form.ix08a2.value)!="" && !isNaN(Math.abs(changeVal(form.ix08a2))))
            	myamount += eval(changeVal(form.ix08a2));
        form.ix08007s.value = myamount;
        form.ix08007.value = myamount;
    }else if(indexp==3){
        if(trimString(form.ix08014.value)!="" && !isNaN(Math.abs(changeVal(form.ix08014))))
            myamount += eval(changeVal(form.ix08014));
        if(trimString(form.ix08015.value)!="" && !isNaN(Math.abs(changeVal(form.ix08015))))
            myamount += eval(changeVal(form.ix08015));
        if(trimString(form.ix08016.value)!="" && !isNaN(Math.abs(changeVal(form.ix08016))))
            myamount += eval(changeVal(form.ix08016));
        if(trimString(form.ix08017.value)!="" && !isNaN(Math.abs(changeVal(form.ix08017))))
            myamount += eval(changeVal(form.ix08017));
        if(trimString(form.ix08018.value)!="" && !isNaN(Math.abs(changeVal(form.ix08018))))
            myamount += eval(changeVal(form.ix08018));
        if(form.ix08a3 != null)
			if(trimString(form.ix08a3.value)!="" && !isNaN(Math.abs(changeVal(form.ix08a3))))
            	myamount += eval(changeVal(form.ix08a3));
        form.ix08013s.value = myamount;
        form.ix08013.value = myamount;
    }
}
//===============================================================================
function Check_OT003W_ShowInsert(form)
{
    if(form.S_Year != null)
    {
        if(trimString(form.S_Year.value)=="")
        {
            alert("年份不可為空");
            return false;
        }
        if(!CheckYear(form.S_Year))
            return false;
        if(!Check_Maintain(form))
    		return false;
    }
    else{
        if(isNaN(Math.abs(changeVal(form.ix08001))) || isNaN(Math.abs(changeVal(form.ix08002))) ||
           isNaN(Math.abs(changeVal(form.ix08003))) || isNaN(Math.abs(changeVal(form.ix08004))) ||
           isNaN(Math.abs(changeVal(form.ix08005))) || isNaN(Math.abs(changeVal(form.ix08006))) ||
           isNaN(Math.abs(changeVal(form.ix08007))) || isNaN(Math.abs(changeVal(form.ix08008))) ||
           isNaN(Math.abs(changeVal(form.ix08009))) || isNaN(Math.abs(changeVal(form.ix08010))) ||
           isNaN(Math.abs(changeVal(form.ix08011))) || isNaN(Math.abs(changeVal(form.ix08012))) ||
           isNaN(Math.abs(changeVal(form.ix08013))) || isNaN(Math.abs(changeVal(form.ix08014))) ||
           isNaN(Math.abs(changeVal(form.ix08015))) || isNaN(Math.abs(changeVal(form.ix08016))) ||
           isNaN(Math.abs(changeVal(form.ix08017))) || isNaN(Math.abs(changeVal(form.ix08018))))
        {
            alert("請輸入數字");
            return false;
        }
        if(form.ix08a1 != null)
        {
        	if(isNaN(Math.abs(changeVal(form.ix08a1))) || isNaN(Math.abs(changeVal(form.ix08a2))) ||
               isNaN(Math.abs(changeVal(form.ix08a3))))
        	{
            	alert("請輸入數字");
            	return false;
        	}
        }
        AddNumber(form,1);
        AddNumber(form,2);
        AddNumber(form,3);

    }

    return true;
}
//====================================================
function doSubmitOT003W(form,myfun)
{
    form.Function.value = myfun;
    if(form.Function.value=='delete')
    {
        if(AskDelete())
            form.submit();
    }
    else if(form.Function.value=='update')
    {
        if(Check_OT003W_ShowInsert(form))
            form.submit();
    }
}
//====================================================
function ChgAction(form,actionURL)
{
    form.action = actionURL;
	form.submit();
}
//===============================================================================
function Check_IX011W_ShowInsert(form)
{
    if(form.Function.value == 'delete')
            return true;
    if(form.Function.value == 'insert')
    {
        if(trimString(form.S_Year.value)=="")
        {
            alert("年份不可為空");
            return false;
        }
        if(!CheckYear(form.S_Year))
            return false;
    }
    if(isNaN(Math.abs(changeVal(form.num1))) || isNaN(Math.abs(changeVal(form.amount1))) ||
       isNaN(Math.abs(changeVal(form.num2))) || isNaN(Math.abs(changeVal(form.amount2))) ||
       isNaN(Math.abs(changeVal(form.num3))) || isNaN(Math.abs(changeVal(form.amount3))) ||
       isNaN(Math.abs(changeVal(form.num4))) || isNaN(Math.abs(changeVal(form.amount4))) ||
       isNaN(Math.abs(changeVal(form.num5))) || isNaN(Math.abs(changeVal(form.amount5))))
    {
        alert("請輸入數字");
        return false;
    }
    if(!Check_Maintain(form))
    	return false;
    return true;
}
//====================================================
function doSubmitIX011W(form,myfun)
{
    form.Function.value = myfun;
    if(form.Function.value=='delete')
    {
        if(AskDelete())
            form.submit();
    }
    else if(form.Function.value=='update')
    {
        if(Check_IX011W_ShowInsert(form))
            form.submit();
    }
}
//===============================================================================
function Check_IX012W_ShowInsert(form)
{
    if(form.Function.value == 'delete')
            return true;
    if(form.Function.value == 'insert')
    {
        if(trimString(form.S_Year.value)=="")
        {
            alert("年份不可為空");
            return false;
        }
        if(!CheckYear(form.S_Year))
            return false;
    }
    if(isNaN(Math.abs(changeVal(form.num1))) || isNaN(Math.abs(changeVal(form.amount1))) ||
       isNaN(Math.abs(changeVal(form.num2))) || isNaN(Math.abs(changeVal(form.amount2))) ||
       isNaN(Math.abs(changeVal(form.num3))) || isNaN(Math.abs(changeVal(form.amount3))) ||
       isNaN(Math.abs(changeVal(form.num4))) || isNaN(Math.abs(changeVal(form.amount4))) ||
       isNaN(Math.abs(changeVal(form.num5))) || isNaN(Math.abs(changeVal(form.amount5))) ||
       isNaN(Math.abs(changeVal(form.num6))) || isNaN(Math.abs(changeVal(form.amount6))) ||
       isNaN(Math.abs(changeVal(form.num7))) || isNaN(Math.abs(changeVal(form.amount7))) ||
       isNaN(Math.abs(changeVal(form.num8))) || isNaN(Math.abs(changeVal(form.amount8))) ||
       isNaN(Math.abs(changeVal(form.num9))) || isNaN(Math.abs(changeVal(form.amount9))) ||
       isNaN(Math.abs(changeVal(form.num10))) || isNaN(Math.abs(changeVal(form.amount10))))
    {
        alert("請輸入數字");
        return false;
    }
    if(!Check_Maintain(form))
    	return false;
    return true;
}
//====================================================
function doSubmitIX012W(form,myfun)
{
    form.Function.value = myfun;
    if(form.Function.value=='delete')
    {
        if(AskDelete())
            form.submit();
    }
    else if(form.Function.value=='update')
    {
        if(Check_IX012W_ShowInsert(form))
            form.submit();
    }
}
//===============================================================================
function Check_IX015W_ShowInsert(form)
{
    if(form.Function.value == 'delete')
            return true;
    if(form.Function.value == 'insert')
    {
        if(trimString(form.S_Year.value)=="")
        {
            alert("年份不可為空");
            return false;
        }
        if(!CheckYear(form.S_Year))
            return false;
    }
    if(isNaN(Math.abs(changeVal(form.num1))) || isNaN(Math.abs(changeVal(form.amount1))) ||
       isNaN(Math.abs(changeVal(form.num2))) || isNaN(Math.abs(changeVal(form.amount2))) ||
       isNaN(Math.abs(changeVal(form.num3))) || isNaN(Math.abs(changeVal(form.amount3))) ||
       isNaN(Math.abs(changeVal(form.num4))) || isNaN(Math.abs(changeVal(form.amount4))) ||
       isNaN(Math.abs(changeVal(form.num5))) || isNaN(Math.abs(changeVal(form.amount5))) ||
       isNaN(Math.abs(changeVal(form.num6))) || isNaN(Math.abs(changeVal(form.amount6))) ||
       isNaN(Math.abs(changeVal(form.num7))) || isNaN(Math.abs(changeVal(form.amount7))) ||
       isNaN(Math.abs(changeVal(form.num8))) || isNaN(Math.abs(changeVal(form.amount8))) ||
       isNaN(Math.abs(changeVal(form.num9))) || isNaN(Math.abs(changeVal(form.amount9))) ||
       isNaN(Math.abs(changeVal(form.num10))) || isNaN(Math.abs(changeVal(form.amount10))) ||
       isNaN(Math.abs(changeVal(form.num11))) || isNaN(Math.abs(changeVal(form.amount11))) ||
       isNaN(Math.abs(changeVal(form.num12))) || isNaN(Math.abs(changeVal(form.amount12))) ||
       isNaN(Math.abs(changeVal(form.num13))) || isNaN(Math.abs(changeVal(form.amount13))) ||
       isNaN(Math.abs(changeVal(form.num14))) || isNaN(Math.abs(changeVal(form.amount14))))
    {
        alert("請輸入數字");
        return false;
    }
    if(!Check_Maintain(form))
    	return false;
    return true;
}
//====================================================
function doSubmitIX015W(form,myfun)
{
    form.Function.value = myfun;
    if(form.Function.value=='delete')
    {
        if(AskDelete())
            form.submit();
    }
    else if(form.Function.value=='update')
    {
        if(Check_IX015W_ShowInsert(form))
            form.submit();
    }
}
//===============================================================================
function Check_OY006W_ShowInsert(form)
{
    if(form.Function.value == 'delete')
            return true;
    if(form.Function.value == 'insert')
    {
        if(trimString(form.S_Year.value)=="")
        {
            alert("年份不可為空");
            return false;
        }
        if(!CheckYear(form.S_Year))
            return false;
    }
    if(isNaN(Math.abs(changeVal(form.ix151))) || isNaN(Math.abs(changeVal(form.ix152))) ||
       isNaN(Math.abs(changeVal(form.ix153))) || isNaN(Math.abs(changeVal(form.ix154))) ||
       isNaN(Math.abs(changeVal(form.ix155))) || isNaN(Math.abs(changeVal(form.ix156))) ||
       isNaN(Math.abs(changeVal(form.ix157))) || isNaN(Math.abs(changeVal(form.ix158))) ||
       isNaN(Math.abs(changeVal(form.ix159))) || isNaN(Math.abs(changeVal(form.ix1513))) ||
       isNaN(Math.abs(changeVal(form.ix1511))) || isNaN(Math.abs(changeVal(form.ix1512))) ||
       isNaN(Math.abs(changeVal(form.ix1514))) || isNaN(Math.abs(changeVal(form.ix1515))))
    {
        alert("請輸入數字");
        return false;
    }
    if(!NewcheckRate(form.ix1510))
        return false;
    if(!Check_Maintain(form))
    	return false;
    return true;
}
//====================================================
function doSubmitOY006W(form,myfun)
{
    form.Function.value = myfun;
    if(form.Function.value=='delete')
    {
        if(AskDelete())
            form.submit();
    }
    else if(form.Function.value=='update')
    {
        if(Check_OY006W_ShowInsert(form))
            form.submit();
    }
}
//===維護人員畫面欄位CHECK=============
function Check_Maintain(form1)
{
   if (trimString(form1.director_name.value)=="" ||  trimString(form1.maintain_dept.value)==""
       || trimString(form1.director_tel.value)=="" || trimString(form1.maintain_tel.value)==""){
        alert("維護部門名稱.主管姓名.主管電話.承辦員電話欄位不可空白！");
        return false;
   }
   return true
}
//=======================================
function AVG_IX011(form,MyIndex)
{
	if(MyIndex==1)
	{
		if(!isNaN(Math.abs(changeVal(form.num1))) && !isNaN(Math.abs(changeVal(form.amount1))))
		{
			if(Math.abs(changeVal(form.num1))!=0 && trimString(form.amount1.value)!="")
			{
				form.avg1.value=eval(changeVal(form.amount1))/Math.abs(changeVal(form.num1));
				form.avg1.value=Math.round(eval(form.avg1.value)*100)/100;
				form.avg1.value=changeStr(form.avg1);
			}
			else
				form.avg1.value="";
		}
	}
	if(MyIndex==2)
	{
		if(!isNaN(Math.abs(changeVal(form.num2))) && !isNaN(Math.abs(changeVal(form.amount2))))
		{
			if(Math.abs(changeVal(form.num2))!=0 && trimString(form.amount2.value)!="")
			{
				form.avg2.value=eval(changeVal(form.amount2))/Math.abs(changeVal(form.num2));
				form.avg2.value=Math.round(eval(form.avg2.value)*100)/100;
				form.avg2.value=changeStr(form.avg2);
			}
			else
				form.avg2.value="";
		}
	}
	if(MyIndex==3)
	{
		if(!isNaN(Math.abs(changeVal(form.num3))) && !isNaN(Math.abs(changeVal(form.amount3))))
		{
			if(Math.abs(changeVal(form.num3))!=0 && trimString(form.amount3.value)!="")
			{
				form.avg3.value=eval(changeVal(form.amount3))/Math.abs(changeVal(form.num3));
				form.avg3.value=Math.round(eval(form.avg3.value)*100)/100;
				form.avg3.value=changeStr(form.avg3);
			}
			else
				form.avg3.value="";
		}
	}
	if(MyIndex==4)
	{
		if(!isNaN(Math.abs(changeVal(form.num4))) && !isNaN(Math.abs(changeVal(form.amount4))))
		{
			if(Math.abs(changeVal(form.num4))!=0 && trimString(form.amount4.value)!="")
			{
				form.avg4.value=eval(changeVal(form.amount4))/Math.abs(changeVal(form.num4));
				form.avg4.value=Math.round(eval(form.avg4.value)*100)/100;
				form.avg4.value=changeStr(form.avg4);
			}
			else
				form.avg4.value="";
		}
	}
	if(MyIndex==5)
	{
		if(!isNaN(Math.abs(changeVal(form.num5))) && !isNaN(Math.abs(changeVal(form.amount5))))
		{
			if(Math.abs(changeVal(form.num5))!=0 && trimString(form.amount5.value)!="")
			{
				form.avg5.value=eval(changeVal(form.amount5))/Math.abs(changeVal(form.num5));
				form.avg5.value=Math.round(eval(form.avg5.value)*100)/100;
				form.avg5.value=changeStr(form.avg5);
			}
			else
				form.avg5.value="";
		}
	}
}
//==================================================
function CountAvg_IX011(form)
{
	AVG_IX011(form,"1");
	AVG_IX011(form,"2");
	AVG_IX011(form,"3");
	AVG_IX011(form,"4");
	AVG_IX011(form,"5");
}

//serissa 2001/11/21&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
function writeCnt(form)
{
    if (isNaN(Math.abs(form.tmonth_cnt.value)))
   	    return form.tmonth_cnt.value;
    if (form.lmonth_cnt.value == '')
            form.lmonth_cnt.value = '0';
    if (form.tmonth_cnt.value == '')
            form.tmonth_cnt.value = '0';

    form.emonth_cnt1.value = eval(changeVal(form.lmonth_cnt)) + eval(changeValA(form.tmonth_cnt));
    form.emonth_cnt.value = form.emonth_cnt1.value;
    form.emonth_cnt1.value = changeStr(form.emonth_cnt1);


}
//==============================================================================
function writeRate1(form)
{
    if(isNaN(Math.abs(form.basic_rate.value)) || isNaN(Math.abs(form.pm_rate1.value)) ||
       isNaN(Math.abs(form.pm_rate2.value))   || isNaN(Math.abs(form.pm_rate3.value)) ||
       isNaN(Math.abs(form.pm_rate4.value))   || isNaN(Math.abs(form.n_pm_rate1.value)) ||
       isNaN(Math.abs(form.n_pm_rate2.value))   || isNaN(Math.abs(form.n_pm_rate3.value)) ||
       isNaN(Math.abs(form.n_pm_rate4.value)))
    {
   	     return false;
    }
    if (form.basic_rate.value == '')
            form.basic_rate.value = '0';
    if (form.pm_rate1.value == '')
            form.pm_rate1.value = '0';
    if (form.pm_rate2.value == '')
            form.pm_rate2.value = '0';
    if (form.pm_rate3.value == '')
            form.pm_rate3.value = '0';
    if (form.pm_rate4.value == '')
            form.pm_rate4.value = '0';
    if (form.n_pm_rate1.value == '')
            form.n_pm_rate1.value = '0';
    if (form.n_pm_rate2.value == '')
            form.n_pm_rate2.value = '0';
    if (form.n_pm_rate3.value == '')
            form.n_pm_rate3.value = '0';
    if (form.n_pm_rate4.value == '')
            form.n_pm_rate4.value = '0';

    form.real_rate1.value = Math.round((eval(changeValA(form.basic_rate)) + eval(changeValA(form.pm_rate1))) * 1000) / 1000;
    form.h_real_rate1.value = form.real_rate1.value;
    form.n_real_rate1.value = Math.round((eval(changeValA(form.basic_rate)) + eval(changeValA(form.n_pm_rate1))) * 1000) / 1000;
    form.h_n_real_rate1.value = form.n_real_rate1.value;

    form.real_rate2.value = Math.round((eval(changeValA(form.basic_rate)) + eval(changeValA(form.pm_rate2))) * 1000) / 1000;
    form.h_real_rate2.value = form.real_rate2.value;
    form.n_real_rate2.value = Math.round((eval(changeValA(form.basic_rate)) + eval(changeValA(form.n_pm_rate2))) * 1000) / 1000;
    form.h_n_real_rate2.value = form.n_real_rate2.value;

    form.real_rate3.value = Math.round((eval(changeValA(form.basic_rate)) + eval(changeValA(form.pm_rate3))) * 1000) / 1000;
    form.h_real_rate3.value = form.real_rate3.value;
    form.n_real_rate3.value = Math.round((eval(changeValA(form.basic_rate)) + eval(changeValA(form.n_pm_rate3))) * 1000) / 1000;
    form.h_n_real_rate3.value = form.n_real_rate3.value;

    form.real_rate4.value = Math.round((eval(changeValA(form.basic_rate)) + eval(changeValA(form.pm_rate4))) * 1000) / 1000;
    form.h_real_rate4.value = form.real_rate4.value;
    form.n_real_rate4.value = Math.round((eval(changeValA(form.basic_rate)) + eval(changeValA(form.n_pm_rate4))) * 1000) / 1000;
    form.h_n_real_rate4.value = form.n_real_rate4.value;
}
//==============================================================================
function writeAmt1(form)
{
    if (isNaN(Math.abs(form.tmonth_amt.value)))
   	    return form.tmonth_cnt.value;
    if (form.lmonth_amt.value == '')
        form.lmonth_amt.value = '0';
    if (form.tmonth_amt.value == '')
        form.tmonth_amt.value = '0';

    form.emonth_amt1.value = eval(changeVal(form.lmonth_amt)) + eval(changeValA(form.tmonth_amt));
    form.emonth_amt.value = form.emonth_amt1.value;

    if (form.neg_amt.value != '' && form.neg_amt.value != '0')
        form.fin_rate.value = changeVal(form.emonth_amt) / changeVal(form.neg_amt) * 100;

	//小數五位後去掉,避免數字太小時出現科學符號
	form.fin_rate.value = Math.round(form.fin_rate.value * Math.pow(10, 5)) / Math.pow(10, 5);

    form.emonth_amt1.value = changeStr(form.emonth_amt1);
    form.les_amt.value =   changeVal(form.neg_amt) - changeVal(form.emonth_amt);

}
//====================================================
function doSubmitWB009W(form,myfun)
{
    form.Function.value = myfun;
    if(form.Function.value=='delete')
    {
        if(AskDelete())
            form.submit();
    }
    else if(form.Function.value=='update')
    {
        if(Check_WB009W_ShowInsert(form))
            form.submit();
    }
}
//==============================================================================
function Check_WB009W_ShowInsert(form)
{
    if(form.Function.value == 'delete')
            return true;
    if(form.Function.value == 'insert')
    {
        if(trimString(form.S_Year.value)=="")
        {
            alert("年份不可為空");
            return false;
        }
        if(!CheckYear(form.S_Year))
            return false;
    }
    if(isNaN(Math.abs(changeVal(form.sign_amount_in)))      ||
       isNaN(Math.abs(changeVal(form.sign_amount_out)))     ||
       isNaN(Math.abs(changeVal(form.sign_qty_in)))         ||
       isNaN(Math.abs(changeVal(form.sign_qty_out)))        ||
       isNaN(Math.abs(changeVal(form.lend_amount_in)))      ||
       isNaN(Math.abs(changeVal(form.lend_amount_out)))     ||
       isNaN(Math.abs(changeVal(form.lend_qty_in)))         ||
       isNaN(Math.abs(changeVal(form.lend_qty_out)))        ||
       isNaN(Math.abs(changeVal(form.store_new_qty)))       ||
       isNaN(Math.abs(changeVal(form.store_disengageqty))) ||
       isNaN(Math.abs(changeVal(form.store_total_qty)))     ||
       isNaN(Math.abs(changeVal(form.fee_in_store)))        ||
       isNaN(Math.abs(changeVal(form.fee_in_lend)))         ||
       isNaN(Math.abs(changeVal(form.fee_in_else)))         ||
       isNaN(Math.abs(changeVal(form.fee_out_bank)))        ||
       isNaN(Math.abs(changeVal(form.fee_out_else)))        ||
       isNaN(Math.abs(changeVal(form.store_frozen_loan))) )
    {
        alert("請輸入數字");
        return false;
    }
    if(!Check_Maintain(form))
    	return false;
    return true;
}
//==============================================================================
function SUM_WB009W(form,MyIndex)
{
	if(MyIndex==1)
	{
		if (trimString(form.sign_amount_in.value) == '')
            form.sign_amount_in.value = '0';
        if (trimString(form.sign_amount_out.value) == '')
            form.sign_amount_out.value = '0';
        if(!isNaN(Math.abs(changeVal(form.sign_amount_in))) && !isNaN(Math.abs(changeVal(form.sign_amount_out)))  )
        {
            form.sign_amount.value = eval(changeVal(form.sign_amount_in)) + eval(changeVal(form.sign_amount_out));
	        form.sign_amount.value = changeStr(form.sign_amount);
	    }
    }
	if(MyIndex==2)
	{
        if (trimString(form.sign_qty_in.value) == '')
            form.sign_qty_in.value = '0';
        if (trimString(form.sign_qty_out.value) == '')
            form.sign_qty_out.value = '0';
        if(!isNaN(Math.abs(changeVal(form.sign_qty_in)))  && !isNaN(Math.abs(changeVal(form.sign_qty_out))) )
        {
            form.sign_qty.value = eval(changeVal(form.sign_qty_in)) + eval(changeVal(form.sign_qty_out));
	        form.sign_qty.value = changeStr(form.sign_qty);
	    }
	}
	if(MyIndex==3)
	{
        if (trimString(form.lend_amount_in.value) == '')
            form.lend_amount_in.value = '0';
        if (trimString(form.lend_amount_out.value) == '')
            form.lend_amount_out.value = '0';
        if(!isNaN(Math.abs(changeVal(form.lend_amount_in)))  && !isNaN(Math.abs(changeVal(form.lend_amount_out))) )
        {
            form.lend_amount.value = eval(changeVal(form.lend_amount_in)) + eval(changeVal(form.lend_amount_out));
	        form.lend_amount.value = changeStr(form.lend_amount);
        }
    }
	if(MyIndex==4)
	{
        if (trimString(form.lend_qty_in.value) == '')
            form.lend_qty_in.value = '0';
        if (trimString(form.lend_qty_out.value) == '')
            form.lend_qty_out.value = '0';
        if(!isNaN(Math.abs(changeVal(form.lend_qty_in)))  && ! isNaN(Math.abs(changeVal(form.lend_qty_out))) )
        {
            form.lend_qty.value = eval(changeVal(form.lend_qty_in)) + eval(changeVal(form.lend_qty_out));
	        form.lend_qty.value = changeStr(form.lend_qty);
	    }
	}
	if(MyIndex==5)
	{
        if (trimString(form.fee_in_store.value) == '')
            form.fee_in_store.value = '0';
        if (trimString(form.fee_in_lend.value) == '')
            form.fee_in_lend.value = '0';
        if (trimString(form.fee_in_else.value) == '')
            form.fee_in_else.value = '0';
        if(!isNaN(Math.abs(changeVal(form.fee_in_store)))  &&
                !isNaN(Math.abs(changeVal(form.fee_in_lend)))  &&
                !isNaN(Math.abs(changeVal(form.fee_in_else)))  )
        {
            form.fee_in.value = eval(changeVal(form.fee_in_store)) +
                                eval(changeVal(form.fee_in_lend)) +
                                eval(changeVal(form.fee_in_else));
	        form.fee_in.value = changeStr(form.fee_in);
        }
	}
	if(MyIndex==6)
	{
        if (trimString(form.fee_out_bank.value) == '')
            form.fee_out_bank.value = '0';
        if (trimString(form.fee_out_else.value) == '')
            form.fee_out_else.value = '0';
        if(!isNaN(Math.abs(changeVal(form.fee_out_else)))  && !isNaN(Math.abs(changeVal(form.store_frozen_loan))) )
        {
            form.fee_out.value = eval(changeVal(form.fee_out_bank)) + eval(changeVal(form.fee_out_else));
	        form.fee_out.value = changeStr(form.fee_out);

        }
    }
}
//====================================================
function doSubmitWB010W(form,myfun)
{
    form.Function.value = myfun;
    if(form.Function.value=='delete')
    {
        if(AskDelete())
            form.submit();
    }
    else if(form.Function.value=='update')
    {
        if(Check_WB010W_ShowInsert(form))
            form.submit();
    }
}
//==============================================================================
function Check_WB010W_ShowInsert(form)
{
    if(form.Function.value == 'delete')
            return true;
    if(form.Function.value == 'insert')
    {
        if(trimString(form.S_Year.value)=="")
        {
            alert("年份不可為空");
            return false;
        }
        if(!CheckYear(form.S_Year))
            return false;
    }
    if(isNaN(Math.abs(changeVal(form.issue_qty)))         ||
       isNaN(Math.abs(changeVal(form.effective_qty)))     ||
       isNaN(Math.abs(changeVal(form.sign_amount_in)))    ||
       isNaN(Math.abs(changeVal(form.sign_amount_out)))   ||
       isNaN(Math.abs(changeVal(form.sign_qty_in)))       ||
       isNaN(Math.abs(changeVal(form.sign_qty_out)))      ||
       isNaN(Math.abs(changeVal(form.lend_amount_in)))    ||
       isNaN(Math.abs(changeVal(form.lend_amount_out)))   ||
       isNaN(Math.abs(changeVal(form.lend_qty_in)))       ||
       isNaN(Math.abs(changeVal(form.lend_qty_out)))      ||
       isNaN(Math.abs(changeVal(form.circle_amount)))     ||
       isNaN(Math.abs(changeVal(form.circle_int_amount))) ||
       isNaN(Math.abs(changeVal(form.fee_delay_fine)))    ||
       isNaN(Math.abs(changeVal(form.fee_year)))          ||
       isNaN(Math.abs(changeVal(form.fee_lost)))          ||
       isNaN(Math.abs(changeVal(form.fee_checkbill)))     ||
       isNaN(Math.abs(changeVal(form.fee_lend)))          ||
       isNaN(Math.abs(changeVal(form.fee_sign)))          ||
       isNaN(Math.abs(changeVal(form.fee_else)))          ||
       isNaN(Math.abs(changeVal(form.fee_delay_30)))      ||
       isNaN(Math.abs(changeVal(form.fee_delay_3160)))    ||
       isNaN(Math.abs(changeVal(form.fee_delay_6190)))    ||
       isNaN(Math.abs(changeVal(form.fee_delay_91120)))   ||
       isNaN(Math.abs(changeVal(form.fee_delay_120)))     ||
       isNaN(Math.abs(changeVal(form.ar_amount)))         ||
       isNaN(Math.abs(changeVal(form.lost_frozen_amt)))   ||
       isNaN(Math.abs(changeVal(form.lost_fake_amt)))     ||
       isNaN(Math.abs(changeVal(form.frozen_return_amt))) )
    {
        alert("請輸入數字");
        return false;
    }
    if(!checkRate5(form.circle_int_rate))
        return false;
    if(!Check_Maintain(form))
    	return false;
    return true;
}
//==============================================================================
function SUM_WB010W(form,MyIndex)
{
	if(MyIndex==1)
	{
		if (trimString(form.sign_amount_in.value) == '')
            form.sign_amount_in.value = '0';
        if (trimString(form.sign_amount_out.value) == '')
            form.sign_amount_out.value = '0';
        if(!isNaN(Math.abs(changeVal(form.sign_amount_in))) && !isNaN(Math.abs(changeVal(form.sign_amount_out)))  )
        {
            form.sign_amount.value = eval(changeVal(form.sign_amount_in)) + eval(changeVal(form.sign_amount_out));
	        form.sign_amount.value = changeStr(form.sign_amount);
	    }
    }
	if(MyIndex==2)
	{
        if (trimString(form.sign_qty_in.value) == '')
            form.sign_qty_in.value = '0';
        if (trimString(form.sign_qty_out.value) == '')
            form.sign_qty_out.value = '0';
        if(!isNaN(Math.abs(changeVal(form.sign_qty_in)))  && !isNaN(Math.abs(changeVal(form.sign_qty_out))) )
        {
            form.sign_qty.value = eval(changeVal(form.sign_qty_in)) + eval(changeVal(form.sign_qty_out));
	        form.sign_qty.value = changeStr(form.sign_qty);
	    }
	}
	if(MyIndex==3)
	{
        if (trimString(form.lend_amount_in.value) == '')
            form.lend_amount_in.value = '0';
        if (trimString(form.lend_amount_out.value) == '')
            form.lend_amount_out.value = '0';
        if(!isNaN(Math.abs(changeVal(form.lend_amount_in)))  && !isNaN(Math.abs(changeVal(form.lend_amount_out))) )
        {
            form.lend_amount.value = eval(changeVal(form.lend_amount_in)) + eval(changeVal(form.lend_amount_out));
	        form.lend_amount.value = changeStr(form.lend_amount);
        }
    }
	if(MyIndex==4)
	{
        if (trimString(form.lend_qty_in.value) == '')
            form.lend_qty_in.value = '0';
        if (trimString(form.lend_qty_out.value) == '')
            form.lend_qty_out.value = '0';
        if(!isNaN(Math.abs(changeVal(form.lend_qty_in)))  && ! isNaN(Math.abs(changeVal(form.lend_qty_out))) )
        {
            form.lend_qty.value = eval(changeVal(form.lend_qty_in)) + eval(changeVal(form.lend_qty_out));
	        form.lend_qty.value = changeStr(form.lend_qty);
	    }
	}
	if(MyIndex==5)
	{
        if (trimString(form.fee_delay_fine.value) == '')
            form.fee_delay_fine.value = '0';
        if (trimString(form.fee_year.value) == '')
            form.fee_year.value = '0';
        if (trimString(form.fee_lost.value) == '')
            form.fee_lost.value = '0';
        if (trimString(form.fee_checkbill.value) == '')
            form.fee_checkbill.value = '0';
        if (trimString(form.fee_lend.value) == '')
            form.fee_lend.value = '0';
        if (trimString(form.fee_sign.value) == '')
            form.fee_sign.value = '0';
        if (trimString(form.fee_else.value) == '')
            form.fee_else.value = '0';

        if(!isNaN(Math.abs(changeVal(form.fee_delay_fine))) &&
                !isNaN(Math.abs(changeVal(form.fee_year))) &&
                !isNaN(Math.abs(changeVal(form.fee_lost))) &&
                !isNaN(Math.abs(changeVal(form.fee_checkbill))) &&
                !isNaN(Math.abs(changeVal(form.fee_lend))) &&
                !isNaN(Math.abs(changeVal(form.fee_sign))) &&
                !isNaN(Math.abs(changeVal(form.fee_else))) )
        {
            form.fee_in.value = eval(changeVal(form.fee_delay_fine)) +
                                eval(changeVal(form.fee_year)) +
                                eval(changeVal(form.fee_lost)) +
                                eval(changeVal(form.fee_checkbill)) +
                                eval(changeVal(form.fee_lend)) +
                                eval(changeVal(form.fee_sign)) +
                                eval(changeVal(form.fee_else));
	        form.fee_in.value = changeStr(form.fee_in);
        }
	}
	if(MyIndex==6)
	{
        if (trimString(form.fee_delay_30.value) == '')
            form.fee_delay_30.value = '0';
        if (trimString(form.fee_delay_3160.value) == '')
            form.fee_delay_3160.value = '0';
        if (trimString(form.fee_delay_6190.value) == '')
            form.fee_delay_6190.value = '0';
        if (trimString(form.fee_delay_91120.value) == '')
            form.fee_delay_91120.value = '0';
        if (trimString(form.fee_delay_120.value) == '')
            form.fee_delay_120.value = '0';
        if (trimString(form.ar_amount.value) == '')
            form.ar_amount.value = '0';

        if(     !isNaN(Math.abs(changeVal(form.fee_delay_30))) &&
                !isNaN(Math.abs(changeVal(form.fee_delay_3160))) &&
                !isNaN(Math.abs(changeVal(form.fee_delay_6190))) &&
                !isNaN(Math.abs(changeVal(form.fee_delay_91120))) &&
                !isNaN(Math.abs(changeVal(form.fee_delay_120))) &&
                !isNaN(Math.abs(changeVal(form.ar_amount))) )
        {
            form.fee_delay.value = eval(changeVal(form.fee_delay_30)) +
                                eval(changeVal(form.fee_delay_3160)) +
                                eval(changeVal(form.fee_delay_6190)) +
                                eval(changeVal(form.fee_delay_91120)) +
                                eval(changeVal(form.fee_delay_120)) +
                                eval(changeVal(form.ar_amount));
	        form.fee_delay.value = changeStr(form.fee_delay);
	        MyIndex = 9;
        }
    }
    if(MyIndex==7)
	{
        if (trimString(form.lost_frozen_amt.value) == '')
            form.lost_frozen_amt.value = '0';
        if (trimString(form.lost_fake_amt.value) == '')
            form.lost_fake_amt.value = '0';
        if(!isNaN(Math.abs(changeVal(form.lost_frozen_amt)))  && ! isNaN(Math.abs(changeVal(form.lost_fake_amt))) )
        {
            form.lost_total_amt.value = eval(changeVal(form.lost_frozen_amt)) + eval(changeVal(form.lost_fake_amt));
	        form.lost_total_amt.value = changeStr(form.lost_total_amt);
	        MyIndex = 8;
	    }
	}
	if(MyIndex==8)
	{
        if (trimString(form.lost_total_amt.value) == '')
            form.lost_total_amt.value = '0';
        if (trimString(form.frozen_return_amt.value) == '')
            form.frozen_return_amt.value = '0';

        if(!isNaN(Math.abs(changeVal(form.lost_total_amt)))  && ! isNaN(Math.abs(changeVal(form.frozen_return_amt))) )
        {
            form.lost_balance.value = eval(changeVal(form.lost_total_amt)) - eval(changeVal(form.frozen_return_amt));
	        form.lost_balance.value = changeStr(form.lost_balance);
	    }
	}
	if(MyIndex==9)
	 {
	        if (trimString(form.circle_amount.value) == '')
	            form.circle_amount.value = '0';
	        if (trimString(form.ar_amount.value) == '')
	            form.ar_amount.value = '0';
	        if (trimString(form.fee_delay_91120.value) == '')
	            form.fee_delay_91120.value = '0';
	        if (trimString(form.fee_delay_120.value) == '')
	            form.fee_delay_120.value = '0';
//	        if (trimString(form.fee_delay.value) == '')
//	            form.fee_delay.value = '0';

	        if(!isNaN(Math.abs(changeVal(form.circle_amount))) &&
	           !isNaN(Math.abs(changeVal(form.ar_amount))) &&
	           !isNaN(Math.abs(changeVal(form.fee_delay_91120))) &&
	           !isNaN(Math.abs(changeVal(form.fee_delay_120))) )
//	               !isNaN(Math.abs(changeVal(form.fee_delay))) )
	        {
	        	var numerator	= eval(changeVal(form.fee_delay_91120)) + eval(changeVal(form.fee_delay_120)) +
	        					  eval(changeVal(form.ar_amount));
	            var denominator = eval(changeVal(form.circle_amount)) + eval(changeVal(form.ar_amount));

	            if(denominator != 0) {
	                form.delay_circle_rate1.value = Math.round((numerator / denominator) * 1000000) / 10000;
//	                form.delay_circle_rate1.value = Math.round((changeVal(form.fee_delay) / denominator) * 1000000) / 10000;
	             	form.delay_circle_rate.value = changeStr(form.delay_circle_rate1);
	         	} else {
	             form.delay_circle_rate1.value = '';
	             form.delay_circle_rate.value = '';
	         	}
	     	}
	 }
}
function checkRate5(Rate1)
{
    if(isNaN(Math.abs(Rate1.value)))
    {
   	    alert("請輸入數字");
   	    return false;
    }
//    if(eval(Rate1.value) > 100 || eval(Rate1.value) < -99.999)
    if(eval(Rate1.value) > 100 || eval(Rate1.value) < 0)
    {
        alert("利率不可大於 100, 也不可小於 0")
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
function changeValA(T1)
{
    pos=0
    var oring=T1.value;
    pos=oring.indexOf(",");
    var c="";
    var i=0;
    oring = (oring).replace(" ","");
    if (oring.charAt(i) == "-"){
        c="-";
        oring=oring.substring(i+1,oring.length);
    }
    if (oring != 0)
        while (oring.charAt(i) == " " || oring.charAt(i) == "0")
            oring=oring.substring(i+1,oring.length);
    while (pos !=-1)
    {
        oring=(oring).replace(",","");
        pos=oring.indexOf(",");
    }

    return c+oring;
}
function changeStrA(T1)
{
    T1.value = changeValA(T1);
    //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
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
//serissa end &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&