<%
  //created on 102.1.8  by 2968
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.*,java.io.*" %>
<%@ page import="org.apache.poi.poifs.filesystem.*,org.apache.poi.hssf.usermodel.*" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.report.RtpFR063W_Print" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Enumeration" %>

<%
	System.out.println("WMFileEdit_A11_Print.jsp Program Start...");
	response.setContentType("application/msexcel;charset=UTF-8");
	
	String act = "download" ;
	String s_year =((String)request.getParameter("s_year")==null)?"":(String)request.getParameter("s_year");
	String s_month =((String)request.getParameter("s_month")==null)?"":(String)request.getParameter("s_month");
	String bank_no =((String)request.getParameter("bank_no")==null)?"":(String)request.getParameter("bank_no");
	String bank_name =((String)request.getParameter("bank_name")==null)?"":(String)request.getParameter("bank_name");
	String Unit =((String)request.getParameter("Unit")==null)?"":(String)request.getParameter("Unit");
	String filename="農漁會信用部聯合貸款案件明細表.xls";
	if(act.equals("view")){
		response.setHeader("Content-disposition","inline; filename=view.xls");
	}else if (act.equals("download")){
		response.setHeader("Content-Disposition","attachment; filename=download.xls");
	}
%>
<%
	try{
	    String actMsg = RtpFR063W_Print.createRpt(s_year,s_month,bank_no,bank_name,Unit);
	    System.out.println("createRpt="+actMsg);
	    FileInputStream fin = null;
	   	fin = new FileInputStream(Utility.getProperties("reportDir")+System.getProperty("file.separator")+filename);
	   	System.out.println(Utility.getProperties("reportDir")+System.getProperty("file.separator")+filename);
	    ServletOutputStream out1 = response.getOutputStream();
	    byte[] line = new byte[8196];
	    int getBytes=0;
	
	    while( ((getBytes=fin.read(line,0,8196)))!=-1 ){
	       out1.write(line,0,getBytes);
	       out1.flush();
	    }
		
	    fin.close();
	    out1.close();
	    
	}catch(Exception e){
	   System.out.println(e.getMessage());
	}
	
	System.out.println("WMFileEdit_A11_Print.jsp Program End...");
	
%>






