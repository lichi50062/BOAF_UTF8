<%
// 94.02.16 add by link to FR004W_Qry.jsp(農會)/FR004F_Qry.jsp(漁會) by 2295
// 95.03.28 add 單月損益表 by 2295
// 99.04.26 fix 搬移部分程式至共用 by 2808
// 99.11.09.fix 修改機構排序 by 2808
//100.05.04 fix 無法顯示漁會機構名稱 by 2295
//100.06.24 fix 機構排列順序,按照直轄市在前.其他縣市在後排序 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.Properties" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@include file="./include/Header.include" %>
<%
	String bank_type = Utility.getTrimString(dataMap.get("bank_type"));				    		
    String hasMonth = Utility.getTrimString(dataMap.get("hasMonth"));			    		
    	
	if(session.getAttribute("muser_id") == null){	
      System.out.println("FR004W login timeout");   
	  rd = application.getRequestDispatcher( "/pages/reLogin.jsp?url=LoginError.jsp?timeout=true" );         	   
      rd.forward(request,response);
    }
	if(!Utility.CheckPermission(request,report_no)){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );        
    }else{            
    	//set next jsp     		    	
    	if(bank_type.equals("6")){
           rd = application.getRequestDispatcher( bank_type6PgName + "?bank_type="+bank_type+"&hasMonth="+hasMonth);                	
        }else if(bank_type.equals("7")){
           rd = application.getRequestDispatcher( bank_type7PgName + "?bank_type="+bank_type+"&hasMonth="+hasMonth);                	
        }
    	//request.setAttribute("TBank", Utility.getALLTBank(bank_type));//100.05.04
    	session.setAttribute("nowbank_type",bank_type);//100.06.24
    	request.setAttribute("TBank",Utility.getBankList(request) );//100.06.24
    }        
     
%>
<%@include file="./include/Tail.include" %>

<%!
    private final static String report_no = "FR004W";
    private final static String nextPgName = "/pages/ActMsg.jsp";    
    private final static String bank_type6PgName = "/pages/"+report_no+"_Qry.jsp";
    private final static String bank_type7PgName = "/pages/FR004F_Qry.jsp";
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";    
         
%>    