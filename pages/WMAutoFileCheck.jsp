<%
//94.01.04 add 超過登入時間,請重新登入 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DownLoad" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="java.util.*" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="com.tradevan.util.UpdateA01" %>
<%@ page import="com.tradevan.util.UpdateA02" %>	
<%@ page import="com.tradevan.util.UpdateA03" %>	
<%@ page import="com.tradevan.util.UpdateA04" %>	
<%@ page import="com.tradevan.util.UpdateA05" %>	
<%@ page import="com.tradevan.util.UpdateA06" %>
<%@ page import="com.tradevan.util.UpdateA08" %>
<%@ page import="com.tradevan.util.UpdateA09" %>
<%@ page import="com.tradevan.util.UpdateA10" %>
<%@ page import="com.tradevan.util.UpdateA99" %>
<%@ page import="com.tradevan.util.UpdateF01" %>
<%@ page import="com.tradevan.util.UpdateB01" %>	
<%@ page import="com.tradevan.util.UpdateB03" %>	
<%@ page import="com.tradevan.util.UpdateM01" %>	
<%@ page import="com.tradevan.util.UpdateM02" %>	
<%@ page import="com.tradevan.util.UpdateM03" %>
<%@ page import="com.tradevan.util.UpdateM04" %>	
<%@ page import="com.tradevan.util.UpdateM05" %>
<%@ page import="com.tradevan.util.UpdateM06" %>	
<%@ page import="com.tradevan.util.UpdateM07" %>	
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.AutoFileCheck" %>
<%@ page import="java.util.Date" %>
<%@ page import="java.io.File" %>
<%@ page import="java.lang.Integer" %>
<%@ page import="java.util.Properties" %>
<%
   String report_no = ( request.getParameter("report_no")==null ) ? "" : (String)request.getParameter("report_no");		
   System.out.println("=============執行檢核開始===========");
   List dbData = DBManager.QueryDB_SQLParam("select * from cdshareno where cmuse_Div='001' and cmuse_id <> 'Z'",null,"");   
   if(dbData != null && dbData.size() != 0){
	  for(int i=0;i<dbData.size();i++){
	       System.out.println("請至"+AutoFileCheck.FileCheck(report_no,(String)((DataObject)dbData.get(i)).getValue("cmuse_id"))+"查看檢核結果");
	  }
	}
   System.out.println("=============執行檢核結束===========");
%>