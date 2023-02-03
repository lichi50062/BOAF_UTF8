<%
//94.01.04 add 超過登入時間,請重新登入 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DownLoad" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>	
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.AutoGenerateA11" %>
<%@ page import="java.util.Date" %>
<%@ page import="java.io.File" %>
<%@ page import="java.lang.Integer" %>
<%@ page import="java.util.Properties" %>
<%
   String report_no = ( request.getParameter("report_no")==null ) ? "" : (String)request.getParameter("report_no");		
   System.out.println("=============執行產生A11開始===========");
   AutoGenerateA11.generate("","");
	  
   System.out.println("=============執行產生A11結束===========");
%>