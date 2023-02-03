<%
//102.07.09 created N年內及目前月份經營統計資料 by 2968
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%
	// 查詢條件值 
    Map dataMap =Utility.saveSearchParameter(request);
	String report_no = "DS055W";	
	String act = Utility.getTrimString(dataMap.get("act"));		
					
	String szExcelAction = (session.getAttribute("excelaction")==null)?"download":(String)session.getAttribute("excelaction");	
	String bank_type = (session.getAttribute("bank_type")==null)?"":(String)session.getAttribute("bank_type");	
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
<link href="css/b51.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0">
<form method=post action='#' name='RptStylefrm'>
<input type='hidden' name='menuBank_type' value='<%=request.getAttribute("menuBank_type") %>'>
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
                <td width="180"><img src="images/banner_bg1.gif" width="180" height="17"></td>
                <td width="*" class="title_font"><%=Utility.getPgName(report_no)%></td>
                <td width="180"><img src="images/banner_bg1.gif" width="180" height="17"></td>
              </tr>
            </table>
          </td>
        </tr>
        <%@include file="./include/DS_RptStyle_AgriBank.include" %><!-- 報表格式(範本檔) -->   
    </table>
    </td>
  </tr>
</table>

<INPUT type="hidden" name=SortList><!--//btnSortList儲存已勾選的報表排序欄位名稱-->
<INPUT type="hidden" name=btnFieldList value='<%if(session.getAttribute("btnFieldList") != null) out.print((String)session.getAttribute("btnFieldList"));%>'><!--//btnFieldList儲存已勾選的報表欄位名稱-->
<INPUT type="hidden" name=BankList value='<%if(session.getAttribute("BankList") != null) out.print((String)session.getAttribute("BankList"));%>'><!--//BankList儲存已勾選的金融機構-->
<INPUT type="hidden" name=S_YEAR value='<%if(session.getAttribute("S_YEAR") != null) out.print((String)session.getAttribute("S_YEAR"));%>'><!--//95.12.04儲存已勾選的起始日期-年-->
<INPUT type="hidden" name=E_YEAR value='<%if(session.getAttribute("E_YEAR") != null) out.print((String)session.getAttribute("E_YEAR"));%>'><!--//95.12.04儲存已勾選的結束日期-年-->
<INPUT type="hidden" name=S_MONTH value='<%if(session.getAttribute("S_MONTH") != null) out.print((String)session.getAttribute("S_MONTH"));%>'><!--//95.12.04儲存已勾選的起始日期-月-->
<INPUT type="hidden" name=E_MONTH value='<%if(session.getAttribute("E_MONTH") != null) out.print((String)session.getAttribute("E_MONTH"));%>'><!--//95.12.04儲存已勾選的結束日期-月-->
</form>

</body>
</html>
