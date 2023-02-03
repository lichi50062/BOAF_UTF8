<%
// 94.03.10	增加機構類別判斷及更改檔名FR006W為FR006W處理 add by egg
// 94.08.15  add Office 2003版Excel現僅提供「下載報表」功能 
// 94.09.05 fix 拿掉檢視報表 by 2295
//100.06.24 add 套用查詢日期模組 by 2295
//105.03.24 add for 地方主管機關使用 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.*" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@include file="./include/bank_no_hsien_id.include" %>
<%     
   // 查詢條件值 
    Map dataMap =Utility.saveSearchParameter(request);
	String report_no = "FR006W";
	showCancel_No=false;//顯示營運中/裁撤別
	showBankType=false;//顯示金融機構類別
    showCityType=false;//顯示縣市別
    showUnit=false;//顯示金額單位
    showPageSetting=true;//顯示本報表採用A4紙張直印/橫印 
	setLandscape=true;//true:橫印
    String cancel_no = "N";
	
    String bank_type = Utility.getTrimString(dataMap.get("bank_type"));	  
    String bankType = bank_type;//ym_hsien_id_unit.include用
%>

<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/HsienIDUtil.js"></script><!-- 根據查詢年月.挑選縣市別/總機構單位-->
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<script language="JavaScript" type="text/JavaScript">
<!--

function doSubmit(cnd){   

   this.document.forms[0].action = "/pages/FR006W_Excel.jsp?act="+cnd;	
   this.document.forms[0].target = '_self';
   this.document.forms[0].submit();   
}

//-->
</script>
<link href="css/b51.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0">
<form method=post action='#'>
<input type=hidden name="bank_type" value="<%=bank_type%>">
<input type='hidden' name="showTbank" value='false'>
<input type='hidden' name="showCityType" value='<%=showCityType%>'>
<input type='hidden' name="showCancel_No" value='<%=showCancel_No%>'>

<table width='600' border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
     <td>&nbsp;</td>
  </tr>
  <tr>
    <td width="20%"><img src="images/banner_bg1.gif" width="100%" height="17"></td>
    <td width="60%"><font color='#000000' size=4><b><center>信用部違反法定比率規定分析表</center></b></font></td>
    <td width="20%"><img src="images/banner_bg1.gif" width="100%" height="17"></td>
  </tr>
</table>
<%
  String nameColor="nameColor_sbody";
  String textColor="textColor_sbody";
  String bordercolor="#3A9D99";
%>
<table border=1 width='600' align=center height="65" bgcolor="#FFFFF" bordercolor="<%=bordercolor%>">
<tr class="sbody" >
    <td colspan="2" height="1" class="<%=nameColor%>">
      <div align="right">
       <input type='radio' name="act" value='download' checked>下載報表  
       <%if(Utility.getPermission(request,report_no,"P")){//Print %> 
       <a href="javascript:doSubmit('download')" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image411','','images/bt_execb.gif',1)"><img src="images/bt_exec.gif" name="Image411" width="66" height="25" border="0" id="Image41"></a>
       <%}%>
       <a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image511','','images/bt_cancelb.gif',0)"><img src="images/bt_cancel.gif" name="Image511" width="66" height="25" border="0" id="Image51"></a>
       <a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image611','','images/bt_reporthelpb.gif',1)"><img src="images/bt_reporthelp.gif" name="Image611" width="80" height="25" border="0" id="Image61"></a>
      </div>
    </td>
</tr>
<%@include file="./include/ym_hsien_id_unit_m2.include" %><!--顯示查詢日期/金融機構類別/縣市別-->
<table border="1" width="600" align="center" height="54" bgcolor="#FFFFF" bordercolor="<%=bordercolor%>">
  <tr>
  	<td class="<%=nameColor%>" colspan="2">
       <div align="center">
  		<table width="574" border="0" cellpadding="1" cellspacing="1">
  		<tr><td width="34"><img src="/pages/images/print_1.gif" width="34" height="34"></td>
            <td width="492"><font color="#CC6600">本報表採用A4紙張橫印</font></td>                              
        </tr>
        </table>
        </div>
    </td>
  </tr>  
</table> 
</table>
</form>
</body>
</html>
