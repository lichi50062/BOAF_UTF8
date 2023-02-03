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
            if(Math.abs(sDate) >= Math.abs(eDate))
            {
                alert("終止年月必須大於起始年月");
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
//============================================================================
function writeRate(form)
{   
    if(form.neg_amt.value!='' && form.neg_amt.value!='0' && form.emonth_amt.value!='')
       form.fin_rate.value =  changeVal(form.emonth_amt) / changeVal(form.neg_amt);
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
        if(trimString(form.S_Year.value) == '')
        {
            alert("基準日不可空白");    
            return false;
        }     
        if(!CheckYear(form.S_Year))
            return false;
        
        bdate = (form.hidebegdate[form.periodyear.selectedIndex].value).substring(0,6);
        edate = (form.hideenddate[form.periodyear.selectedIndex].value).substring(0,6);
        
        if(eval(form.S_Month.value)<10)
            mdate = (eval(form.S_Year.value)+1911) + "0" + form.S_Month.value;
        else                        
            mdate = (eval(form.S_Year.value)+1911) + "" + form.S_Month.value;
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
    
    if(isNaN(Math.abs(form.basic_rate.value)))
    {
        alert("基本放款率中請輸入數字");    
        return false;
    }    
    for(i=0;i<4;i++)
    {
        if(isNaN(Math.abs(form.fin_rate1[i].value))   || isNaN(Math.abs(form.fin_rate2[i].value)) ||
           isNaN(Math.abs(form.house_rate1[i].value)) || isNaN(Math.abs(form.house_rate2[i].value)))
        {
            alert("本年度貸款利率政策表格中請輸入數字");    
            return false;
        }    
    }
  //~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    if(eval(form.basic_rate.value) > 99.99 || eval(form.basic_rate.value) < -99.99)
    {
        alert("基本放款率不可大於 99.99, 也不可小於 -99.99")    
        return false;
    } 
    for(i=0;i<4;i++)
    {
        if(eval(form.fin_rate1[i].value) > 99.99 || eval(form.fin_rate1[i].value) < -99.99)
        {
            alert("利率不可大於 99.99, 也不可小於 -99.99")    
            return false;
        }    
        if(eval(form.fin_rate2[i].value) > 99.99 || eval(form.fin_rate2[i].value) < -99.99)
        {
            alert("利率不可大於 99.99, 也不可小於 -99.99")    
            return false;
        } 
        if(eval(form.house_rate1[i].value) > 99.99 || eval(form.house_rate1[i].value) < -99.99)
        {
            alert("利率不可大於 99.99, 也不可小於 -99.99")    
            return false;
        } 
        if(eval(form.house_rate1[i].value) > 99.99 || eval(form.house_rate1[i].value) < -99.99)
        {
            alert("利率不可大於 99.99, 也不可小於 -99.99")    
            return false;
        } 
    }
//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    if((form.basic_rate.value).indexOf(".") != -1 )
    {
        len = (form.basic_rate.value).substring((form.basic_rate.value).indexOf(".")+1,form.basic_rate.value.length);
        if(len.length > 2)
        {
            alert("基本放款率的小數點後只能有兩個位數");  
            return false;  
        }
    }        
    for(i=0;i<4;i++)    
    {
        if((form.fin_rate1[i].value).indexOf(".") != -1 )   
        {
            len = (form.fin_rate1[i].value).substring((form.fin_rate1[i].value).indexOf(".")+1,form.fin_rate1[i].value.length);
            if(len.length > 2)
            {
                alert("利率的小數點後只能有兩個位數");  
                return false;  
            }
        }
        if((form.fin_rate2[i].value).indexOf(".") != -1 )
        {
            len = (form.fin_rate2[i].value).substring((form.fin_rate2[i].value).indexOf(".")+1,form.fin_rate2[i].value.length);
            if(len.length > 2)
            {
                alert("利率的小數點後只能有兩個位數");  
                return false;  
            }    
         }
         if((form.house_rate1[i].value).indexOf(".") != -1 )
         {
            len = (form.house_rate1[i].value).substring((form.house_rate1[i].value).indexOf(".")+1,form.house_rate1[i].value.length);
            if(len.length > 2)
            {
                alert("利率的小數點後只能有兩個位數");  
                return false;  
            }    
         }
         if((form.house_rate2[i].value).indexOf(".") != -1 )
         {
            len = (form.house_rate2[i].value).substring((form.house_rate2[i].value).indexOf(".")+1,form.house_rate2[i].value.length);
            if(len.length > 2)
            {
                alert("利率的小數點後只能有兩個位數");  
                return false;  
            }        
         }
    }
    
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
    if(!CheckYear(form.baseyear))
        return false;
    if(form.baseyear.value == ''   || form.proposerno.value == ''   || form.proposername.value == '' ||
       form.subsidiary.value == '' || form.investmentno.value == '' || form.guaranteemoney.value == '')
    {
        alert("統計表欄位不可為空白");   
        return flase; 
    }   
    if(form.maintain_dept.value == '' || form.maintain_name.value == '' ||
       form.maintain_tel.value == ''  || form.director_tel.value =='')
    {
        alert("維護人基本資料表格中不可有空白欄位");    
        return false;
    }

    return true;
}
//=====================================================
function doIX017Submit(form,myfun)
{
    form.Function.value = myfun;
    if(Check_IX017W_ShowInsert(form))
        form.submit();
}
//==================================================
function check_IX011R(form)
{
    if(Check_WB006W_Range(form))
    {
        if(form.S_YEAR.value.length != 0 && form.S_MONTH.value.length != 0 && 
           form.E_YEAR.value.length != 0 && form.E_MONTH.value.length != 0)    
        {
            return true;    
        }
        else
        {
            alert("請完整輸入起迄年月");    
            return false;
        }    
    }
    return false;
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
function Check_WFX01W(form)
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
//====================================================================
function LinkToInsert(form,linkaddr)
{
    linkaddr = "/servlet/" + linkaddr;
    if(form.Function.value=='insert')
    {
        alert("請先新增外國銀行相關資料");
        return;
    }
    else
        window.open(linkaddr,'_self');
}
//========================================================
function Check_WFX01W_ShowSubR(form)
{
    if(!NewcheckRate(form.revrate))
        return false;
    return true;
}
//===========================================================================
function doWFX01WSubmit(form,myfun)
{
    form.Function.value = myfun;
    if(myfun=="delete")
    {
        if(AskDelete())
            form.submit();
    }        
    else    
    {
        if(Check_WFX01W_ShowSubR(form))
            form.submit();
    }
}
//========================================================
function Check_WFX02W(form)
{
    if(!NewcheckRate(form.captrate))
        return false;
    if(!NewcheckRate(form.roerate))
        return false;
    if(!NewcheckRate(form.cirate))
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
        var setyear = (Math.abs(form.setupdateyear.value)) + 1911;
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
    
    if(form.policymemo.value.length>126)
    {
         alert("輸入的重要政策概述不可超過126個字");
         return false;    
    }
    if(form.impmemo.value.length>126)
    {
         alert("輸入的重大事項概述不可超過126個字");
         return false;    
    }
        
    
    return true;
}
//===========================================================================
function doWFX02WSubmit(form,myfun)
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
function Check_WFX03W(form)
{
    if(trimString(form.opfund.value)=="")
    {
        alert("營運所用資金不可為空");
        return false;
    }
    return true;        
}
//===========================================================================
function doWFX03WSubmit(form,myfun)
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
//==================================================
function DeleteAction(form)
{
    form.Function.value = 'delete';
    if(AskDelete()) form.submit();
}
