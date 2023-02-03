<%
//94.09.05  add by 2495
//99.04.12  修改部分程式為共用 fix by 2808 
//105.03.24 add for 地方主管機關使用 by 2295
%>

<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.Calendar" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Calendar" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@include file="./include/Header.include" %>
<%@include file="./include/bank_no_hsien_id.include" %>
<%!
	private final static String report_no = "FR023W";
	private final static String ListPgName = "/pages/"+report_no+"_Qry.jsp";
	private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    private final static String nextPgName = "/pages/ActMsg.jsp";    
    private final static String RptCreatePgName = "/pages/"+report_no+"_Excel.jsp";        
    private final static String RptQryPgName = ListPgName ;
%>
<%
	showCancel_No=false;//顯示營運中/裁撤別
	showBankType=false;//顯示金融機構類別
	showCityType=false;//顯示縣市別
	showUnit=true;//顯示金額單位
	showPageSetting=true;//顯示報表列印格式
	setLandscape=false;//true:橫印
    String bank_type = Utility.getTrimString(dataMap.get("bank_type")) ;
	String bankType = bank_type ;
	String cancel_no = "" ;
 	String title=(bank_type.equals("6"))?"農(漁)會信用部各類對象存放款比率表":"農(漁)會信用部各類對象存放款比率表";
 	
%>
<script language="javascript" src="js/Common.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<script language="javascript" src="js/HsienIDUtil.js"></script><!-- 根據查詢年月.挑選縣市別/總機構單位-->
<html>
<head>
<script language="JavaScript" type="text/JavaScript">
<!--
function doSubmit(){      
   var s_year = this.document.forms[0].S_YEAR.value =="" ? 0 : eval(this.document.forms[0].S_YEAR.value ) ;
   var s_month =  this.document.forms[0].S_MONTH.value =="" ? 0 : eval(this.document.forms[0].S_MONTH.value ) ;
   //if(this.document.forms[0].S_YEAR.value <= "94" && this.document.forms[0].S_MONTH.value < "06"){
   //99.04.08 fix by 2808
   if(s_year <= 94 && s_month< 6) {
      alert('94年6月起開始受理申報資料');
      return;
   }
   if (confirm('本作業須執行約5-10秒,確定執行?')) 
   {	   
	this.document.forms[0].action = "/pages/FR023W_Excel.jsp";	
   	this.document.forms[0].target = '_self';
   	this.document.forms[0].submit();
  }       
}

//-->
</script>
<link href="css/b51.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0">
<form method=post action='#'>
<input type="hidden" name="wordaction" value='download' >
<input type='hidden' name="showTbank" value='<%=showBankType %>'>
<input type='hidden' name="showCityType" value='<%=showCityType%>'>
<input type=hidden name="bank_type" value=<%=bank_type%>>
<table width='600' border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="25%"><img src="images/banner_bg1.gif" width="100%" height="17"></td>
    <td width="50%"><font color='#000000' size=4> 
					<nobr><%=title%></nobr>
					</font>
    </td>
    <td width="25%"><img src="images/banner_bg1.gif" width="100%" height="17"></td>
  </tr>
</table>
<%
  String nameColor="nameColor_sbody";
  String textColor="textColor_sbody";
  String bordercolor="#3A9D99";
%>
<Table border=1 width='600' align=center height="65" bgcolor="#FFFFF" bordercolor="<%=bordercolor%>">
<tr class="sbody">
    <td colspan="2" height="1" class="<%=nameColor%>">
      <div align="right">
       <input type='radio' name="excelaction" value='download' checked> 下載報表
	   <%if(Utility.getPermission(request,report_no,"P")){ //Print %>
       <a href="javascript:doSubmit('createRpt')" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image411','','images/bt_execb.gif',1)"><img src="images/bt_exec.gif" name="Image411" width="66" height="25" border="0" id="Image41"></a>
       <%} %>
       <a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image511','','images/bt_cancelb.gif',0)"><img src="images/bt_cancel.gif" name="Image511" width="66" height="25" border="0" id="Image51"></a>
       <a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image611','','images/bt_reporthelpb.gif',1)"><img src="images/bt_reporthelp.gif" name="Image611" width="80" height="25" border="0" id="Image61"></a>
      </div>
    </td>
</tr>
<%@include file="./include/ym_hsien_id_unit_m2.include" %><!--顯示查詢日期/金融機構類別/縣市別-->

</Table>
<table border="1" width="600" align="center" height="54" bgcolor="#FFFFF" bordercolor="<%=bordercolor%>">
  <tr>
  	<td class="<%=nameColor%>" colspan="2">
       <div align="center">
  		<table width="555" border="0" cellpadding="1" cellspacing="1">
  		<tr><td width="34"><img src="/pages/images/print_1.gif" width="34" height="34"></td>
            <td width="222"><font color="#CC6600"><nobr>本報表採用A4紙張橫印(本報表包括全體農會及漁會的資料)</nobr></font></td>                              
        </tr>
        </table>
        </div>
    </td>
  </tr>  
</table>

</form>

</body>
</html>
