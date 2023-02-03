<%
//農漁會信用部各項法定比例明細表
// 99.03.09 add 根據查詢年月.顯示所屬縣市別/總機構單位 by 2295
// 99.03.10 add bank_no_hsien_id.include為縣市別/總機構單位xml by 2295
//              HsienIDUtil.js 根據查詢年月.改變縣市別名稱/根據查詢年度.縣市別.改變總機構名稱 by 2295
// 99.04.07 add 根據查詢日期,若需顯示縣市別才連動顯示縣市別 by 2295
//              根據查詢日期/縣市別,若需顯示總機構單位才連動顯示總機構單位
// 99.04.08 add 增加營運中/裁撤別選項=false by 2295
//102.01.17 fix 查詢日期.年度 by 2295
//105.03.22 add for 地方主管機關使用 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@include file="./include/bank_no_hsien_id.include" %>
<%
	// 查詢條件值 
    Map dataMap =Utility.saveSearchParameter(request);
	String report_no = "FR0066W";
	showCancel_No=false;//顯示營運中/裁撤別
	showBankType=true;//顯示金融機構類別
    showCityType=true;//顯示縣市別
    showPageSetting=true;//顯示本報表採用A4紙張直印/橫印 
	setLandscape=false;//true:橫印
    String cancel_no = "N";
	String act = Utility.getTrimString(dataMap.get("act"));	
    String bankType = Utility.getTrimString(dataMap.get("bankType"));	
    String tbank = Utility.getTrimString(dataMap.get("tbank"));	
    String cityType = Utility.getTrimString(dataMap.get("cityType"));	
    String title = ((bankType.equals("6"))?"農會":"漁會");
    System.out.println(report_no+"_Qry.act="+act);
%>

<HTML>
<HEAD>
<TITLE>農漁會信用部各項法定比例明細表</TITLE>
<script language="javascript" src="js/FR0066W.js"></script>
<script language="javascript" src="js/HsienIDUtil.js"></script><!-- 根據查詢年月.挑選縣市別/總機構單位-->
<link href="css/b51.css" rel="stylesheet" type="text/css">
<%
  String nameColor="nameColor_sbody";
  String textColor="textColor_sbody";
  String bordercolor="#3A9D99";
%>
<BODY bgColor=#FFFFFF>
<Form name='form' method=post action='#'>
<input type='hidden' name="act" value=''>
<input type='hidden' name="rtf_bank_type" value=''>
<input type="hidden" name="wordaction" value='download' >
<input type='hidden' name="showTbank" value='true'>
<input type='hidden' name="showCityType" value='<%=showCityType%>'>
<input type='hidden' name="showCancel_No" value='<%=showCancel_No%>'>

<table width='600' border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="25%"><img src="images/banner_bg1.gif" width="100%" height="17"></td>
    <td width="50%"><font color='#000000' size=4><b><center>農漁會信用部各項法定比例明細表</center></b></font></td>
    <td width="25%"><img src="images/banner_bg1.gif" width="100%" height="17"></td>
  </tr>
</table>
<table border=1 width='600' align=center height="65" bgcolor="#FFFFF" bordercolor="<%=bordercolor%>">
<tr class="sbody">
    <td colspan="2" height="1" class="<%=nameColor%>">
      <div align="right">
       <input type='radio' name="excelaction" value='download' checked> 下載報表
       <a href="javascript:doSubmit('createRpt')" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image411','','images/bt_execb.gif',1)"><img src="images/bt_exec.gif" name="Image411" width="66" height="25" border="0" id="Image41"></a>
       <a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image511','','images/bt_cancelb.gif',0)"><img src="images/bt_cancel.gif" name="Image511" width="66" height="25" border="0" id="Image51"></a>
       <a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image611','','images/bt_reporthelpb.gif',1)"><img src="images/bt_reporthelp.gif" name="Image611" width="80" height="25" border="0" id="Image61"></a>
      </div>
    </td>
</tr>
<%@include file="./include/ym_hsien_id_unit_m2.include" %><!--顯示查詢日期/金融機構類別/縣市別-->

<tr class="sbody">
<td width="118" class="<%=nameColor%>" height="1">總機構單位</td>
<td width="416" class="<%=textColor%>" height="1">
  <select size="1" name="tbank">
  </select> </td>
</tr>
<table border="1" width="600" align="center" height="54" bgcolor="#FFFFF" bordercolor="<%=bordercolor%>">
  <tr>
  	<td class="<%=nameColor%>" colspan="2">
       <div align="center">
  		<table width="574" border="0" cellpadding="1" cellspacing="1">
  		<tr><td width="34"><img src="/pages/images/print_1.gif" width="34" height="34"></td>
            <td width="492"><font color="#CC6600">本報表採用A4紙張直印</font></td>                              
        </tr>
        </table>
        </div>
    </td>
  </tr>  
</table>   
</table>
</form>
</BODY>
<script language="JavaScript" >
<!--
setSelect(form.bankType,"<%=bankType%>");
changeCity('CityXML', form.cityType, form.S_YEAR, form);
changeTbank('TBankXML', form.tbank, form.cityType, form);
-->
</script>
</HTML>