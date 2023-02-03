<%
//99.07.29 first designed by 2660
//99.08.24 licno欄位加大至50 by 2295
%>

<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Date" %>

<%
  // 接收資料
  List WLX13_S_Edit = (List)request.getAttribute("WLX13_S_Edit");
  List BN01_Data = (List)request.getAttribute("BN01_Data");
  List BA01_Data = (List)request.getAttribute("BA01_Data");
  List CD034_Data = (List)request.getAttribute("CD034_Data");
  List CD035_Data = (List)request.getAttribute("CD035_Data");

  String title = (WLX13_S_Edit == null)?"新增":"維護";
  String m_year = ( request.getParameter("m_year")==null ) ? "" : (String)request.getParameter("m_year");
  int m_month = Integer.parseInt(( request.getParameter("m_month")==null ) ? "0" : (String)request.getParameter("m_month"));
  String tbank_no = ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");
  String bank_no = "";
  String seq_no = "";
  String name = "";
  String position_code = "";
  String train_place = "";
  String course_name = "";
  String course_hour = "";
  String begin_date = "";
  String end_date = "";
  String licno = "";
  String other_trainplace = "";
  
  Date today = new Date();
  String nowDay = today.getYear() + "/" + today.getMonth() + "/" + today.getDate();
  String Temp_No = "";
  String Temp_Name = "";
  Date Temp_Date = new Date();
  int Temp_Year = 0;
  int Temp_Month = 0;
  int Temp_Day = 0;

  DataObject Bean = new DataObject();  // 使用dataobject type 來做資料傳遞

  if(WLX13_S_Edit == null){
    System.out.println("WLX13_S_Edit == null");
  }else{
    // 使用dataobject type 來做資料傳遞
    Bean = (DataObject)WLX13_S_Edit.get(0); // 把選取的資料取出來
    System.out.println("WLX13_S_Edit.size()="+WLX13_S_Edit.size());
    m_year = Bean.getValue("m_year").toString();
    m_month = Integer.parseInt(Bean.getValue("m_month").toString());
    tbank_no = (String)Bean.getValue("tbank_no");
    bank_no = (String)Bean.getValue("bank_no");
    seq_no = Bean.getValue("seq_no").toString();
    name = Bean.getValue("name")==null?"":(String)Bean.getValue("name");
    position_code = (String)Bean.getValue("position_code");
    train_place = (String)Bean.getValue("train_place");
    course_name = Bean.getValue("course_name")==null?"":(String)Bean.getValue("course_name");
    course_hour = Bean.getValue("course_hour").toString();
    begin_date = Bean.getValue("begin_date").toString();
    end_date = Bean.getValue("end_date").toString();
    licno = Bean.getValue("licno")==null?"":(String)Bean.getValue("licno");
    other_trainplace = Bean.getValue("other_trainplace")==null?"":(String)Bean.getValue("other_trainplace");
  }

  // 取得FX013W的權限
  Properties permission = ( session.getAttribute("FX013W")==null ) ? new Properties() : (Properties)session.getAttribute("ZZ005W");
  if(permission == null){
    System.out.println("FX013W_List.permission == null");
  }
  else{
    System.out.println("FX013W_List.permission.size ="+permission.size());
  }
%>

<html>
  <head>
    <title>理監事基本資料維護</title>
    <link href="css/b51.css" rel="stylesheet" type="text/css">
    <script language="javascript" src="js/FX013W.js"></script>
    <script language="javascript" src="js/Common.js"></script>
    <script language="javascript" event="onresize" for="window"></script>
    <script language="JavaScript" type="text/JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
    </script>
  </head>
  <body topmargin="0" leftmargin="15">
    <form method=post>
      <input type="hidden" name="act" value="">
      <input type="hidden" name="seq_no" value="<%=seq_no %>">
      <div><img src="images/space_1.gif" width="12" height="12"><div>
      <table width="600" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="40"><img src="images/banner_bg1.gif" width="40" height="17"></td>
          <td width="520">
            <font color='#000000' size=4><b><center>信用部從業人員參加金融相關業務進修情形申報作業</center></b></font>
          </td>
          <td width="40"><img src="images/banner_bg1.gif" width="40" height="17" align="right"></td>
        </tr>
      </table>
      <div><img src="images/space_1.gif" width="12" height="12"><div>
      <table width="600" border="0" cellpadding="0" cellspacing="0">
        <tr>
<jsp:include page="getLoginUser.jsp" flush="true" /> <!-- list the login user in form-->
        </tr>
      </table>
      <table width=600 border=1 cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
        <tr class="sbody">
          <td width='30%' align='left' bgcolor='#D8EFEE' class="sbody">申報年月</td>
          <td width='70%' bgcolor='e7e7e7'>
            <input type='text' name='m_year' value="<%=m_year==""?today.getYear() - 11:m_year %>" size='3' maxlength='3' onblur='CheckYear(this)'><font color='#000000'>年
            <select name='m_month' size="1">
<%
  for (int k = 1; k < 13; k++) {
    m_month = m_month==0?today.getMonth():m_month;
%>
              <option value=<%= k %> <%=m_month==k?"selected":""%> ><%= k %></option>
<%
  }
%>
            </select>月
          </td>
        </tr>
        <tr class="sbody">
          <td width='30%' align='left' bgcolor='#D8EFEE' class="sbody">總機構代號</td>
          <td width='70%' bgcolor='e7e7e7'>
            <select name='tbank_no' size="1">
<%
  if(BN01_Data != null && BN01_Data.size() != 0) {
    for (int i = 0; i < BN01_Data.size(); i++) {
      Temp_No = (String)((DataObject)BN01_Data.get(i)).getValue("bank_no");
      Temp_Name = (String)((DataObject)BN01_Data.get(i)).getValue("bank_name");
%>
              <option value=<%=Temp_No %> <%=Temp_No.equals(tbank_no) ? "selected" : ""%> ><%=Temp_Name %></option>
<%
    }
  }
%>
            </select><font color="red" size=4>*</font>
          </td>
        </tr>
        <tr class="sbody">
          <td width='30%' align='left' bgcolor='#D8EFEE' class="sbody">總分支機構代號</td>
          <td width='70%' bgcolor='e7e7e7'>
            <select name='bank_no' size="1">
<%
  if(BA01_Data != null && BA01_Data.size() != 0){
    for (int i = 0; i < BA01_Data.size(); i++) {
      Temp_No = (String)((DataObject)BA01_Data.get(i)).getValue("bank_no");
      Temp_Name = (String)((DataObject)BA01_Data.get(i)).getValue("bank_name");
%>
              <option value=<%=Temp_No %> <%=Temp_No.equals(bank_no) ? "selected" : ""%> ><%=Temp_Name %></option>
<%
    }
  }
%>
            </select><font color="red" size=4>*</font>
          </td>
        </tr>  
        <tr class="sbody">
          <td width='30%' align='left' bgcolor='#D8EFEE' class="sbody">學員姓名</td>
          <td width='70%' bgcolor='e7e7e7'>
            <input type='text' name='name' value="<%=name %>" size=25 maxlength=5><font color="red" size=4>*</font>
          </td>
        </tr>  
        <tr class="sbody">
          <td width='30%' align='left' bgcolor='#D8EFEE'>職稱</td>
          <td width='70%' bgcolor='e7e7e7'>
            <select name='position_code' size="1">
<%
  for (int i = 0; i < CD034_Data.size(); i++) {
    Temp_No = (String)((DataObject)CD034_Data.get(i)).getValue("cmuse_id");
    Temp_Name = (String)((DataObject)CD034_Data.get(i)).getValue("cmuse_name");
%>
              <option value=<%=Temp_No %> <%=Temp_No.equals(position_code) ? "selected" : ""%> ><%=Temp_Name %></option>
<%
  }
%>
            </select><font color="red" size=4>*</font>
          </td>
        </tr>
        <tr class="sbody">
          <td width='30%' align='left' bgcolor='#D8EFEE'>訓練機構</td>
          <td width='70%' bgcolor='e7e7e7'>
            <select name='train_place' size="1">
<%
  for (int i = 0; i < CD035_Data.size(); i++) {
    Temp_No = (String)((DataObject)CD035_Data.get(i)).getValue("cmuse_id");
    Temp_Name = (String)((DataObject)CD035_Data.get(i)).getValue("cmuse_name");
%>
              <option value=<%=Temp_No %> <%=Temp_No.equals(train_place) ? "selected" : ""%> ><%=Temp_Name %></option>
<%
  }
%>
            </select><font color="red" size=4>*</font>
          </td>
        </tr>
        <tr class="sbody">
          <td width='30%' bgcolor='#D8EFEE' align='left'>課程名稱</td>
          <td width='70%' bgcolor='e7e7e7'>
            <input type='text' name='course_name' value="<%=course_name %>" size='52' maxlength='20'>
            <font color="red" size=4>*</font>
          </td>
        </tr>
        <tr class="sbody">
          <td width='30%' bgcolor='#D8EFEE' align='left'>課程時數</td>
          <td width='70%' bgcolor='e7e7e7'>
            <input type='text' name='course_hour' value="<%=course_hour %>" size='10' maxlength='20' >小時<font color="red" size=4>*</font>
          </td>
        </tr>
        <tr class="sbody">
          <td width='30%' align='left' bgcolor='#D8EFEE'>上課期間</td>
          <td width='70%' bgcolor='e7e7e7'>
<%
  if(WLX13_S_Edit != null && WLX13_S_Edit.size() != 0){
    Temp_Date = (Date)Bean.getValue("begin_date");
    Temp_Year = Temp_Date.getYear() - 11;
    Temp_Month = Temp_Date.getMonth() + 1;
    Temp_Day = Temp_Date.getDate();
  }
%>
            <input type='text' name='begin_date_y' value="<%=Temp_Year==0?"":Temp_Year %>" size='3' maxlength='3' onblur='CheckYear(this)'>
            <font color='#000000'>年
              <select id="hide1" name=begin_date_m>
                <option></option>
<%
  for (int k = 1; k < 13; k++) {
%>
                <option value=<%= k %> <%=Temp_Month==k?"selected":""%> ><%= k %></option>
<%
  }
%>
              </select>
            </font>
            <font color='#000000'>月
              <select id="hide1" name=begin_date_d>
                <option></option>
<%
  for (int k = 1; k < 32; k++) {
%>
                <option value=<%= k %> <%=Temp_Day==k?"selected":""%> ><%= k %></option>
<%
  }
%>
              </select>
            </font>
            <font color='#000000'>日</font> ~
<%
  if(WLX13_S_Edit != null && WLX13_S_Edit.size() != 0){
    Temp_Date = (Date)Bean.getValue("begin_date");
    Temp_Year = Temp_Date.getYear() - 11;
    Temp_Month = Temp_Date.getMonth() + 1;
    Temp_Day = Temp_Date.getDate();
  }
%>
            <input type='text' name='end_date_y' value="<%=Temp_Year==0?"":Temp_Year %>" size='3' maxlength='3' onblur='CheckYear(this)'>
            <font color='#000000'>年
              <select id="hide2" name='end_date_m'>
                <option></option>
<%
  for (int k = 1; k < 13; k++) {
%>
                <option value=<%= k %> <%=Temp_Month==k?"selected":""%> ><%= k %></option>
<%
  }
%>
              </select>月
              <select id="hide3" name='end_date_d'>
                <option></option>
<%
  for (int k = 1; k < 32; k++) {
%>
                <option value=<%= k %> <%=Temp_Day==k?"selected":""%> ><%= k %></option>
<%
  }
%>
              </select>日
            </font><font color="red" size=4>*</font>
          </td>
        </tr> 
        <tr class="sbody">
          <td width='30%' align='left' bgcolor='#D8EFEE'>證書字號</td>
          <td width='70%' bgcolor='e7e7e7'>
            <input type='text' name='licno' value="<%=licno %>" size='50' maxlength='50' ><font color="red" size=4>*</font>
          </td>
        </tr>	
        <tr class="sbody">
          <td width='30%' align='left' bgcolor='#D8EFEE'>其他訓練機構備註</td>
          <td width='70%' bgcolor='e7e7e7'>
            <input type='text' name='other_trainplace' value="<%=other_trainplace %>" size='31' maxlength='10' >
          </td>
        </tr>	
      </table>
      <div style="width:600" align="center" >
<jsp:include page="getMaintainUser.jsp" flush="true" /><!--載入維護者資訊 -->
    </div>
      <div><img src="images/space_1.gif" width="12" height="12"><div>
      <div style="width:600" align="center" >
        <table border="0" cellspacing="1">
          <tr>
<%
  if(WLX13_S_Edit == null){
%>
            <td width="66">
              <a href="javascript:doSubmit(this.document.forms[0],'Insert');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_confirmb.gif',1)">
              <img src="images/bt_confirm.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>
            </td>
<%
  } else {
%>
            <td width="66">
              <a href="javascript:doSubmit(this.document.forms[0],'Update');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_updateb.gif',1)">
              <img src="images/bt_update.gif" name="Image105" width="66" height="25" border="0" id="Image105"></a>
            </td>
            <td width="66">
              <a href="javascript:doSubmit(this.document.forms[0],'Delete');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image103','','images/bt_deleteb.gif',1)">
              <img src="images/bt_delete.gif" name="Image103" width="66" height="25" border="0" id="Image103"></a>
            </td>
<% } %>
            <td width="66">
              <a href="javascript:AskReset(this.document.forms[0]);" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image105','','images/bt_cancelb.gif',1)">
              <img src="images/bt_cancel.gif" name="Image105" width="66" height="25" border="0" id="Image105"></a>
            </td>
            <td width="93">
              <a href="javascript:history.back();"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image106','','images/bt_backb.gif',1)">
              <img src="images/bt_back.gif" name="Image106" width="80" height="25" border="0" id="Image106"></a>
            </td>
          </tr>
        </table>
      </div>
      <table width="600" border="0" cellpadding="1" cellspacing="1" class="sbody">
        <tr> 
          <td colspan="2"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明 : </font></font></td>
        </tr>
        <tr> 
          <td width="16">　</td>
          <td width="577">
            <ul>
              <li>確認輸入資料無誤後, 按<font color="#666666">【確定】</font>即將本表上的資料, 於資料庫中建檔。</li>
              <li>修改資料無誤後, 按<font color="#666666">【修改】</font>即將本表上的資料, 於資料庫中建檔。</li>
              <li>欲重新輸入資料, 按<font color="#666666">【取消】</font>即將本表上的資料清空。</li>                          
              <li>如放棄, 按<font color="#666666">【回上一頁】</font>即離開本程式。</li>                          
              <li>【<font color="red" size=4>*</font>】為必填欄位。</li>
            </ul>
          </td>
        </tr>
      </table>
    </form>
  </body>
</html>