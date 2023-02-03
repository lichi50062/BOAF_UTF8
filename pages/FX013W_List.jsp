<%
// 94.10.14 first designed by 2660
//100.01.27 fix sqlInjection by 2295
%>
<script language="javascript" src="js/FX013W.js"></script>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" event="onresize" for="window"></script>

<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@ page import="java.io.*" %>
<%
  // ---------------------------------------取得參數---------------------------------------------
  List FX013_list1 = (List)request.getAttribute("WLX13_S");   	// 內含每月的資料
  List FX013_list2 = (List)request.getAttribute("WLX13_S_Sum"); // 內含年月以及筆數

  Properties permission = ( session.getAttribute("FX013W") == null ) ? new Properties() : (Properties)session.getAttribute("FX013W");
  if (permission == null) {
    System.out.println("FX013W_List.permission == null");
  }else {
    System.out.println("FX013W_List.permission.size =" + permission.size());
  }

  String bank_no = ( request.getParameter("bank_no")==null ) ? "" :request.getParameter("bank_no");
  String bank_name = "";
  String lguser_name = (session.getAttribute("muser_name") == null)?"":(String)session.getAttribute("muser_name"); 
  // 登入者的tbank_no
  String tbank_no = (session.getAttribute("tbank_no") == null)?"":(String)session.getAttribute("tbank_no"); 
  // 所點選的nowtbank_no
  // 若有點選的tbank_no,則顯示所點選的tbank_no
  tbank_no = ( session.getAttribute("nowtbank_no")==null ) ? tbank_no : (String)session.getAttribute("nowtbank_no");	
  String wlx01_m_year = (Integer.parseInt(Utility.getYear()) < 100)?"99":"100"; 		
  String sqlCmd = "select bank_name from (select * from ba01 where m_year=?)BA01 where bank_no=?";
  List paramList = new ArrayList() ;
  paramList.add(wlx01_m_year);
  paramList.add(tbank_no);
  List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
  if(dbData != null && dbData.size() != 0){
    bank_name = (String)((DataObject)dbData.get(0)).getValue("bank_name");
    System.out.println("bank_name="+bank_name);
  }

  // table width
  int width = 950;//1300 pixel
%>
<html>
  <head>
    <link href="css/b51.css" rel="stylesheet" type="text/css">
    <STYLE type=text/css>td {table-layout:fixed;word-break:break-all}</Style>
    <title>信用部主任及分部主任參訓情形申報作業</title>
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
    <form method="post">
      <div><img src="images/space_1.gif" width="12" height="12"><div>
      <table width="<%=width %>" align="center" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="250"><img src="images/banner_bg1.gif" width="250" height="17"></td>
          <td width="450" align="center">
            <font color='#000000' size=4><b>信用部從業人員參加金融相關業務進修情形申報作業</b></font>
          </td>
          <td width="250"><img src="images/banner_bg1.gif" width="250" height="17"></td>
        </tr>
      </table>
      <table width="<%=width %>" border="0" cellpadding="0" cellspacing="0">
        <tr>
<jsp:include page="getLoginUser.jsp?width=950" flush="true" /> <!-- list the login user in form-->
        </tr>
      </table>
      <table width="<%=width %>" border=1 cellpadding="1" cellspacing="1" bordercolor="#3A9D99">
        <tr class="sbody">
          <td class="sbody" bgcolor="#9AD3D0" width="85" height="20"><p align="center">申報年月</p></td>
          <td class="sbody" bgcolor="#E7E7E7" width="715" valign="middle" colspan=10>
<%
  Date today = new Date(); //取得現在時間 initial 為現在時間
%>
            <input type='text' name='m_year' value="<%=today.getYear()-11%>" size='3' maxlength='3' onblur='CheckYear(this)'>	
            <font color='#000000'>年               	
            <select name='m_month' size="1">
              <option></option>
<%
  for (int k = 1; k < 13; k++) {
%>
              <option value=<%= k %> <%=today.getMonth()==k?"selected":""%> ><%= k %></option>
<%
  }
%>
            </select>月
            </font><a href="javascript:doSubmit(this.document.forms[0],'New');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image101','','images/bt_addb.gif',1)"><img src="images/bt_add.gif" name="Image101" width="66" height="25" border="0" id="Image101"></a>	
          </td>
        </tr>
        <tr class="sbody">
          <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="85px" >申報年月</td>
          <td bgcolor="#9AD3D0" colspan="2" valign="middle" align="center" width="110px">參訓課程</td>
          <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="150px">機構名稱</td>
          <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="70px" >學員姓名</td>
          <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="60px" >職稱</td>
          <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="145px">訓練機構</td> 
          <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="70px" >課程名稱</td>
          <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="60px" >課程時數</td>
          <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="110px">異動者帳號/姓名</td>
          <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="90px">異動日期</td>
        </tr>
        <tr class="sbody">
          <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="65px" >總筆數</td>
          <td bgcolor="#9AD3D0" rowspan="2" valign="middle" align="center" width="45px" >編號</td>            
        </tr>
      </table>
 <!-- -----------start to parse the form ----------------------------------------------------------- -->
 <%
  if( FX013_list2.size() == 0 ) {
 %>
      <table border="1" cellspacing="1" width="<%=width %>" bordercolor="#3A9D99">
        <tr class="sbody" bgcolor="#D8EFEE">
          <td colspan="2" align=center>尚無資料</td>
        </tr>
      </table>

<%
  } else {
				
    //備註 使用div圖層與javascript去達成展開的動作

    int count = 0;		//紀錄在哪一年哪一月之下有多少筆資料
    int displacement = 0;	//紀錄位移(因為所有的資料是放在同一個list中,所以要紀錄之前已經印過多少筆資料)
    String seq_no = "";
    String m_year = "";
    String m_month = "";

    for(int i=0; i < FX013_list2.size(); i++ ) {
      count = Integer.parseInt(((DataObject)FX013_list2.get(i)).getValue("cnt").toString());
      m_year = ((DataObject)FX013_list2.get(i)).getValue("m_year").toString();
      m_month = ((DataObject)FX013_list2.get(i)).getValue("m_month").toString();
 %> 
      <table class="sbody" width="<%=width %>" border=1 cellpadding="0" cellspacing="1" bordercolor="#3A9D99" height="20" >
        <tr class="sbody" bgcolor="#e7e7e7">
          <td width="85px"  align="center">
<%    if( count > 0 ){ %>
            <div onmousedown="fmenu<%=i %>()" style="cursor:hand">
              <u>
<%    }%>
                <%=m_year %>/<%=m_month %>
<%    if( count > 0 ){ %>
              </u>
            </div>
<%    }%>
          </td>
          <td width="65px"  align="center"><%=count%>　</td>
          <td width="45px"  align="right">　</td>
          <td width="150px" align="right">　</td>
          <td width="70px"  align="right">　</td>
          <td width="60px"  align="right">　</td>
          <td width="145px" align="right">　</td>
          <td width="70px"  align="right">　</td>
          <td width="60px"  align="right">　</td>
          <td width="110px" align="left">　</td>
          <td width="90px"  align="center">　</td>
        </tr>   
      </table>
      <table id="menu<%=i %>" style="display:none" border="1" cellspacing="1" width="<%=width %>" bordercolor="#3A9D99" >
<%
      for (int j = 0; j < count; j++){
        DataObject bean = (DataObject)FX013_list1.get(j + displacement);
        seq_no = bean.getValue("seq_no").toString();
%>
        <tr class="sbody" bgcolor="#e7e7e7">
          <td width="85px" align="center">　</td>
          <td width="65px" align="center">　</td>
          <td width="45px" align="right">
            <a href="FX013W.jsp?act=Edit&seq_no=<%=seq_no %>&m_year=<%=m_year %>&m_month=<%=m_month%>"><%=seq_no %></a>
          </td>
          <td width="150px" align="right" ><%=(String)bean.getValue("bank_name") %></td>
          <td width="70px"  align="right" ><%=(String)bean.getValue("name") %></td>
          <td width="60px"  align="right" ><%=(String)bean.getValue("position_name") %></td>
          <td width="145px" align="right" ><%=(String)bean.getValue("train_place_name") %></td>
          <td width="70px"  align="right" ><%=(String)bean.getValue("course_name") %></td>
          <td width="60px"  align="right" ><%=bean.getValue("course_hour").toString() %></td>
          <td width="110px" align="left"  ><%=bean.getValue("user_id") %>/<%=bean.getValue("user_name") %></td>
          <td width="90px"  align="center"><%=bean.getValue("update_date")==null?"　":Utility.getCHTdate((bean.getValue("update_date")).toString().substring(0, 10), 0) %></td>
        </tr> 
<%
      }
      displacement += count;
    }
  }
%>
      </table>
      <div style="width:600" align="center" >
<jsp:include page="getMaintainUser.jsp?width=950" flush="true" /><!--載入維護者資訊 -->
      </div>
      <table width="767" border="0" cellpadding="1" cellspacing="1" class="sbody">
        <tr>
          <td colspan="2" width="763"><font color='#990000'><img src="images/arrow_1.gif" width="28" height="23" align="absmiddle"><font color="#007D7D" size="3">使用說明: </font></font></td>
        </tr>
        <tr>
          <td width="16">　</td>
          <td class="sbody" width="743">
            <ul>
              <li>點選<font color="#666666">【新增】按鈕</font>可新增該年月「信用部主任及分部主任參加金融相關業務進修情形」之資料。
              <li>點選所列之[編號]可變更該申報之資料。
            </ul>
          </td>
        </tr>
      </table>
    </form>
  </body>
</html>
<!-- --------------------------------------------------- -->
<script language='JavaScript'>
<%
	DataObject tempBean;
	for(int k = 0; k < FX013_list2.size(); k++){
		tempBean = ((DataObject)FX013_list2.get(k));
%>
function fmenu<%= k %>(){
  if( menu<%= k %>.style.display == "none")
    menu<%= k %>.style.display = "block";
  else
    menu<%= k %>.style.display = "none";
}
arrpush('<%=tempBean.getValue("m_year").toString()+tempBean.getValue("m_month").toString()%>');
<%}%>
</script>
<!-- -------------------------------------------------	-->