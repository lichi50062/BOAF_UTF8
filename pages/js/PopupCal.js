
//popupCal begin
//=======西元年轉成民國年=========================================================
function changeMDY(oring) {
  var tms1;
  var tms2;
  var tms3;
  tms1 = Math.abs(oring.substring(0, 4)) - 1911;
  tms2 = oring.substring(4);
  oring = tms1 + tms2 + "";
  return oring;
}
//=======西元年轉成民國年, 並加入分割字元
//=========================================================
function turnDateFormat(oring, s) {
  var tms1 = changeMDY(oring);
  return dateformat(addZero(tms1, 7), s);
}
//====日期補滿(結果呈現在欄位上)======================================================
function addZeroField(DateValue) {
  if (trimString(DateValue.value) == '') {
      return;
  } else if (!isNaN(Math.abs(DateValue.value))) {
      DateValue.value = Math.abs(DateValue.value);

      var max = DateValue.maxLength;
      if ((max == 7 && DateValue.value.length == 6)
              || (max == 5 && DateValue.value.length == 4)
              || (max == 3 && DateValue.value.length == 2))
          DateValue.value = "0" + DateValue.value;
      else if ((max == 7 && DateValue.value.length == 5)
              || (max == 5 && DateValue.value.length == 3)
              || (max == 3 && DateValue.value.length == 1))
          DateValue.value = "00" + DateValue.value;
      else if (max == 2 && DateValue.value.length == 1)
          DateValue.value = "0" + DateValue.value;
  }
}
//==日期補滿(結果回傳)======================================================================
function addZero(id, num) {
  var temp = "";
  for ( var i = 0; i < num; i++) {
      temp += "0";
  }
  temp = temp + id;
  end = temp.length;
  start = end - num;

  return temp.substring(start, end);
}
//======================================================
//對日期插入分割字完
function dateformat(str, sign) {
  var date = str;
  if (str.length == 7) {
      date = str.substring(0, 3) + sign + str.substring(3, 5) + sign
              + str.substring(5, 7);
  }
  return date;
}
//數字轉字串補上千分位
function commaFormat(num) {
  num = num + "";
  var re = /(-?\d+)(\d{3})/;
  while (re.test(num)) {
    num = num.replace(re, "$1,$2");
  }
  return num;
}

//合成日期欄位,並做檢查, 把結果放入tar 的欄位, fieldstr日期欄位, 以; 做分隔字
function mergeCheckedDate(fieldstr, tar) {
if(tar == null ) return;
var date="";
var token = fieldstr.split(';');


if(document.all(token[0]).length) {
  var temp = document.all(token[0]);

  for(var i=0; i< temp.length; i++ ) {
          if(document.all(token[0])[i].value=="") {
      continue;
    }
    date += addZero(document.all(token[0])[i].value,3);
    for(var j=1; j< token.length ; j++ ) {
      date += addZero(document.all(token[j])[i].value,2);
    }

    document.all(tar)[i].value = date;
    if(!checkdate(document.all(tar)[i])) {
            document.all(tar)[i].value = "";
            document.all(token[0])[i].focus;
            return false;
    }
    //alert(document.all(tar)[i].value);
    date = "";
  }
} else {
        if(document.all(token[0]).value == "") {
          return true;
        }
        date += addZero(document.all(token[0]).value,3);
  for(var i=1; i< token.length ; i++ ) {
    date += addZero(document.all(token[i]).value,2);
  }
  document.all(tar).value = date;
  if(!checkdate(document.all(tar))) {
          document.all(tar).value = "";
//          document.all(token[0]).focus;
          return false;
  }
}
return true;
}

//======日期POP視窗共用變數=================================================
var timerID = null;
var wndCal = null;
var frm = null;
var comp = null;
var pic = null;
var ev = null;
var calrowid = null;
var calrowflag = null;

//==日期元件POP視窗================================
function popupCal(f, c, p, e, rowid, rowflag) {
  frm = f;
  comp = c;
  pic = p;
  ev = e;
  calrowid = rowid;
  calrowflag = rowflag;
  if (wndCal != null) {
      wndCal.close();
      wndCal == null;
  }
  wndCal = open("./PopupCalendar.htm", "", "width=220,height=180");
  var dynamicX = e.screenX;
  var dynamicY = e.screenY;
  if (dynamicX + 200 > screen.width) {
      dynamicX -= 200 + 25;
  }
  if (dynamicY + 180 + 25 + 25 > screen.height) {// 25,25 are title bar
      // height and fast bar

      dynamicY -= 180 + 25;// 20 is title bar height
  }
  wndCal.moveTo(dynamicX, dynamicY);
  timerID = setInterval("timerCheck()", 100);
}
//========日期POP視窗選完之後將畫面移除且將值放入欄位================================
function timerCheck() {
  if (wndCal.loadOK) {
      clearInterval(timerID);
      wndCal.setRowInfo(calrowid, calrowflag);
      wndCal.showCalendar(frm, comp, pic);
  }
}
//popupCal end

//=====檢查日期===========================================================
function checkdate(date) {
    startdate = date.value;
    startdate = trimString(startdate);
    if (isNaN(Math.abs(startdate))) {
        alert("日期不可為文字");
        return false;
    }
    if (startdate.substring(0, 2) == '00') {
        alert("日期有誤");
        return false;
    }

    // 轉換為西元年
    if (startdate.length == 7) {
        startdate = (Math.abs(startdate.substring(0, startdate.length - 4)) + 1911)
                + startdate.substring(startdate.length - 4, startdate.length);

    } else if (startdate.length >= 8) { // 如有時分則做檢查
        hour = startdate.substring(startdate.length - 4, startdate.length - 2);
        min = startdate.substring(startdate.length - 2, startdate.length);
        startdate = startdate.substring(0, startdate.length - 4);
        startdate = (Math.abs(startdate.substring(0, startdate.length - 4)) + 1911)
                + startdate.substring(startdate.length - 4, startdate.length);

        if (hour > 12 || hour < 1) {
            alert("時 應介於1至12 之間");
            return false;
        }
        if (min > 60 || min < 0) {
            alert("分 應介於0至59 之間");
            return false;
        }
    } else {
        alert("日期有誤");
        return false;
    }

    if ((Math.abs(startdate.substring(0, startdate.length - 4)) < 1921)) {
        alert("日期有誤");
        return false;
    }

    if (!fnValidDate(startdate)) {
        alert("日期有誤");
        return false;
    }
    return true;
}
function fnValidDate(dateStr){

    var leap = 28;
    if (leapYear(parseInt(dateStr.substring(0,4),10)) == 1){
        leap = 29;
    }
    var tmp = parseInt(dateStr.substring(4,6),10);

    if (tmp < 1 || tmp > 12) {
        return (false);
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

    var dtmp = parseInt(dateStr.substring(6),10);

    if(dtmp < 1 || dtmp > monthTable[tmp]) {
        return (false);
    }
    return (true);
}