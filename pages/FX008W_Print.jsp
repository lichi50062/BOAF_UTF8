<%
/*
  *   created on 95.8.21  by lilic0c0 2495
  *  
  */ 
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.*,java.io.*" %>
<%@ page import="org.apache.poi.poifs.filesystem.*,org.apache.poi.hssf.usermodel.*" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.report.RtpFX008W_Print" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Enumeration" %>

<%
	System.out.println("FX008W_Print.jsp Program Start...");
	
	response.setHeader("Content-Disposition","attachment; filename=download.xls");
	response.setContentType("application/msexcel;charset=UTF-8");//以上這行設定本網頁為excel格式的網頁
	
	
	String act =  "download" ;
	String bank_name = ( request.getParameter("bank_name")==null ) ? "" : (String)request.getParameter("bank_name");
	String dureassure_no = ( request.getParameter("dureassure_no")==null ) ? "" : (String)request.getParameter("dureassure_no");
	String debtname = ( request.getParameter("debtname")==null ) ? "" : (String)request.getParameter("debtname");
	String year = ( request.getParameter("year")==null ) ? "" : (String)request.getParameter("year");
	String month = ( request.getParameter("month")==null ) ? "" : (String)request.getParameter("month");
	String day = ( request.getParameter("day")==null ) ? "" : (String)request.getParameter("day");
	String dureassuresite = ( request.getParameter("dureassuresite")==null ) ? "" : (String)request.getParameter("dureassuresite");
	String accountamt = ( request.getParameter("accountamt")==null ) ? "" : (String)request.getParameter("accountamt");
	String applydelayyear_month = ( request.getParameter("applydelayyear_month")==null ) ? "" : (String)request.getParameter("applydelayyear_month");
	String applydelayreason = ( request.getParameter("applydelayreason")==null ) ? "" : (String)request.getParameter("applydelayreason");
	String damage_yn = ( request.getParameter("damage_yn")==null ) ? "" : (String)request.getParameter("damage_yn");
	String disposal_fact_yn = ( request.getParameter("disposal_fact_yn")==null ) ? "" : (String)request.getParameter("disposal_fact_yn");
	String disposal_plan_yn = ( request.getParameter("disposal_plan_yn")==null ) ? "" : (String)request.getParameter("disposal_plan_yn");
	

	String filename="信用部承受擔保品申請延長處分期限審核表.xls";
	
	
	
	if(act.equals("view")){
		//以上這行設定傳送到前端瀏覽器時的檔名為test1.xls
		//就是靠這一行，讓前端瀏覽器以為接收到一個excel檔
		response.setHeader("Content-disposition","inline; filename=view.xls");
	}else if (act.equals("download")){
		response.setHeader("Content-Disposition","attachment; filename=download.xls");
	}
%>
<%
	try{
	    String actMsg = RtpFX008W_Print.createRpt(bank_name,dureassure_no,debtname,year,month,day,dureassuresite,applydelayyear_month,accountamt,applydelayreason,damage_yn,disposal_fact_yn,disposal_plan_yn);
	    System.out.println("createRpt="+actMsg);
	    FileInputStream fin = null;
		
			System.out.println("filename="+Utility.getProperties("reportDir")+System.getProperty("file.separator")+"信用部承受擔保品申請延長處分期限審核表.xls");
	   	fin = new FileInputStream(Utility.getProperties("reportDir")+System.getProperty("file.separator")+"信用部承受擔保品申請延長處分期限審核表.xls");
		
		
	    ServletOutputStream out1 = response.getOutputStream();
	    byte[] line = new byte[8196];
	    int getBytes=0;
	
	    while( ((getBytes=fin.read(line,0,8196)))!=-1 ){
	       out1.write(line,0,getBytes);
	       out1.flush();
	    }
		
		//關閉檔案
	    fin.close();
	    out1.close();
	    
	}catch(Exception e){
	   System.out.println(e.getMessage());
	}
	
	System.out.println("FX008W_Print.jsp Program End...");
	
%>






