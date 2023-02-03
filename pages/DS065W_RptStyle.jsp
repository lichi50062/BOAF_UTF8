<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%
	// 查詢條件值 
    Map dataMap =Utility.saveSearchParameter(request);
	String report_no = "DS065W";
	String title = Utility.getPgName(report_no);
	String act = Utility.getTrimString(dataMap.get("act"));		
					
	String szExcelAction = (session.getAttribute("excelaction")==null)?"download":(String)session.getAttribute("excelaction");	
	String bank_type = (session.getAttribute("nowbank_type")==null)?"":(String)session.getAttribute("nowbank_type");	
	System.out.println(report_no+"_RptStyle.szExcelAction="+szExcelAction);
	
    List templateList = null;
    if(request.getAttribute("templateList") != null){
       templateList = (List)request.getAttribute("templateList");
       System.out.println("templateList.size()="+templateList.size());
    }else{
       System.out.println("templateList is null");
    }   
%>



<script language="javascript" src="js/Common.js"></script>
<script language="javascript" src="js/DSUtil.js"></script>
<script language="javascript" src="js/movesels.js"></script>
<script language="javascript" event="onresize" for="window"></script>
<html>
<head>
<script language="JavaScript" type="text/JavaScript">
<!--
function chInput(form,cnd){
    if(cnd == 'BankList'){//金融機構類別Panel時,才需檢查此項   
	   if(!checkDateS(form.S_YEAR,form.S_MONTH,'查詢年月 -起始日期','1')) return false;
	   if(!CheckQueryDate2(form.S_YEAR,form.S_MONTH,form.S_DATE,"起始日期")) return false;
    }
    //新舊無法同時列印
    if(parseInt(form.S_YEAR.value)<=99 && parseInt(form.E_YEAR.value)>=100){
        alert('縣市於100年起改用新制，新舊制無法同時輸出，請重新輸入結束日期');
        return false;
    }
	return true;
}
//-->
</script>
<link href="css/b51.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0">
<form method=post action='#' name='RptStylefrm'>
<table width="750" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
     <td>&nbsp;</td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF">
	<table width="791" border="0" align="center" cellpadding="1" cellspacing="1">        
        <tr> 
          <td width="824">
            <table width="797" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr> 
                <td width="257"><img src="images/banner_bg1.gif" width="256" height="17"></td>
                <td width="280" class="title_font"><%=title%></td>
                <td width="254"><img src="images/banner_bg1.gif" width="254" height="17"></td>
              </tr>
            </table>
          </td>
        </tr>
        <%@include file="./include/DS_RptStyle.include" %><!-- 報表格式(範本檔) -->   
    </table>
    </td>
  </tr>
</table>
<INPUT type="hidden" name=SortList><!--//btnSortList儲存已勾選的報表排序欄位名稱-->
<INPUT type="hidden" name=btnFieldList value='<%if(session.getAttribute("btnFieldList") != null) out.print((String)session.getAttribute("btnFieldList"));%>'><!--//btnFieldList儲存已勾選的報表欄位名稱-->
<INPUT type="hidden" name=BankList value='<%if(session.getAttribute("BankList") != null) out.print((String)session.getAttribute("BankList"));%>'><!--//BankList儲存已勾選的金融機構-->
<INPUT type="hidden" name=S_YEAR value='<%if(session.getAttribute("S_YEAR") != null) out.print((String)session.getAttribute("S_YEAR"));%>'><!--//95.12.04儲存已勾選的起始日期-年-->
<INPUT type="hidden" name=S_MONTH value='<%if(session.getAttribute("S_MONTH") != null) out.print((String)session.getAttribute("S_MONTH"));%>'><!--//95.12.04儲存已勾選的起始日期-月-->
<INPUT type="hidden" name=DS_bank_type value='<%if(session.getAttribute("DS_bank_type") != null) out.print((String)session.getAttribute("DS_bank_type"));%>'><!--//102.11.11-->
<INPUT type="hidden" name=loan_item value='<%if(session.getAttribute("loan_item") != null) out.print((String)session.getAttribute("loan_item"));%>'>
<INPUT type="hidden" name=bank_type value='<%=bank_type%>'>
</form>

</body>
</html>
