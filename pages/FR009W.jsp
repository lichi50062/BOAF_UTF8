<%
/* 
 * 100.05.09 fix 信用部淨值占風險性資產比率計算表套用畫面模組 by 2295
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
        request.setAttribute("TBank", Utility.getALLTBank(bank_type));
    	request.setAttribute("City", Utility.getCity());
        rd = application.getRequestDispatcher( ListPgName +"?act=List&bankType="+bank_type);
    }
%>
<%@include file="./include/Tail.include" %>
<%!
private final static String report_no = "FR009W";
private final static String ListPgName = "/pages/"+report_no+"_Qry.jsp";
private final static String LoginErrorPgName = "/pages/LoginError.jsp";
%>