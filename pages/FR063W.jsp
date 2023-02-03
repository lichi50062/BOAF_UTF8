<%
// Copy from FR003W
// 102.01.24 add by 2968
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.Properties" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@include file="./include/Header.include" %>
<%
	String bank_type = Utility.getTrimString(dataMap.get("bank_type"));			    		
    	
	if(!Utility.CheckPermission(request,report_no)){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );        
    }else{            
    	//set next jsp  
    	rd = application.getRequestDispatcher( ListPgName + "?bank_type="+bank_type); 
    	//request.setAttribute("TBank", Utility.getALLTBank(bank_type));//100.05.03
    	//request.setAttribute("City", Utility.getCity());
    	session.setAttribute("nowbank_type",bank_type);//100.06.24
    	request.setAttribute("TBank",Utility.getBankList(request) );//按照直轄市在前.其他縣市在後排序.
        
    }        
     
%>
<%@include file="./include/Tail.include" %>
<%!
    private final static String report_no = "FR063W";
    private final static String nextPgName = "/pages/ActMsg.jsp";    
    private final static String ListPgName = "/pages/"+report_no+"_Qry.jsp";
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";    
                   
%>    