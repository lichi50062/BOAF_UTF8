<%
/* 94.01.26 create 農漁會信用部各項法定比例明細表
 * 99.03.10 add 取得99/100年度的縣市別/總機構資料 by 2295
 *100.06.24 fix 機構排列順序,按照直轄市在前.其他縣市在後排序 by 2295
 */
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@include file="./include/Header.include" %>
<%
	String bank_type = Utility.getTrimString(dataMap.get("bank_type"));
	if(!Utility.CheckPermission(request,report_no)){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );
	} else {
        //request.setAttribute("TBank", Utility.getALLTBank(bank_type));
        session.setAttribute("nowbank_type",bank_type);//100.06.24
    	request.setAttribute("TBank",Utility.getBankList(request) );//100.06.24按照直轄市在前.其他縣市在後排序.
    	request.setAttribute("City", Utility.getCity());
        rd = application.getRequestDispatcher( ListPgName +"?act=List&bankType="+bank_type);
    }
%>
<%@include file="./include/Tail.include" %>
<%!
private final static String report_no = "FR0066W";
private final static String ListPgName = "/pages/"+report_no+"_Qry.jsp";
private final static String LoginErrorPgName = "/pages/LoginError.jsp";
%>