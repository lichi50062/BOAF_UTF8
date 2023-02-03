<%
//95.11.27 add 農(漁)會信用部法定比率分析統計表 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.report.RptFR0066WB" %>								          
<%@ page import="java.util.List" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Enumeration" %>
<%@ page import="java.util.HashMap" %>
<%@ page import="java.util.Hashtable" %>
<%@ page import="java.io.FileOutputStream" %>
<%@ page import="java.io.BufferedOutputStream" %>
<%@ page import="java.io.File" %>
<%@ page import="java.io.PrintStream" %>
<%@ page import="java.io.*" %>

<%
	response.setContentType("application/msexcel;charset=UTF-8");//以上這行設定本網頁為excel格式的網頁       
	
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");			    	
	String bank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");			    		
	String BANK_DATA = ( request.getParameter("BankListSrc")==null ) ? "" : (String)request.getParameter("BankListSrc");			  
	String BANK_NO = "";
    String BANK_NAME = "";
    if(!BANK_DATA.equals("")){
       System.out.print(":BANK_DATA="+BANK_DATA);   
       BANK_NO = (BANK_DATA.startsWith("全體"))?"ALL":BANK_DATA.substring(0,7);
       BANK_NAME = (BANK_DATA.startsWith("全體"))?BANK_DATA.substring(0,BANK_DATA.length()):BANK_DATA.substring(7,BANK_DATA.length());    
	}else{
	   System.out.println(":BANK_DATA="+BANK_DATA);   
	   BANK_NO = ( request.getParameter("BANK_NO")==null ) ? "" : (String)request.getParameter("BANK_NO");
       BANK_NAME = ( request.getParameter("BANK_NAME")==null ) ? "" : (String)request.getParameter("BANK_NAME");
	}	
    String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? "" : (String)request.getParameter("S_YEAR");
    String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? "" : (String)request.getParameter("S_MONTH");
    String Unit = ( request.getParameter("Unit")==null ) ? "1" : (String)request.getParameter("Unit");//金額單位			      
    String filename=(bank_type.equals("6")?"農會":"漁會")+"信用部法定比率分析統計表.xls";
    System.out.print("FR0066WB_Excel.S_YEAR="+S_YEAR);    
    System.out.print(":bank_type="+bank_type);   
    System.out.print(":S_MONTH="+S_MONTH);   
    System.out.print(":BANK_NO="+BANK_NO);
    System.out.print(":BANK_NAME="+BANK_NAME);
    System.out.println(":Unit="+Unit);
   
    String excelAction = ( request.getParameter("excelaction")==null ) ? "download" : (String)request.getParameter("excelaction");			    	
    
    

    if(excelAction.equals("view")){
      //以上這行設定傳送到前端瀏覽器時的檔名為test1.xls
      //就是靠這一行，讓前端瀏覽器以為接收到一個excel檔 
      response.setHeader("Content-disposition","inline; filename=view.xls");
    }else if (excelAction.equals("download")){   
      response.setHeader("Content-Disposition","attachment; filename=download.xls");
    }
    
    try{	
	    String actMsg = RptFR0066WB.createRpt(S_YEAR,S_MONTH,BANK_NO,bank_type,BANK_NAME,Unit);	    
	    System.out.println("createRpt="+actMsg);
	    System.out.println("filename="+Utility.getProperties("reportDir")+System.getProperty("file.separator")+filename);
		FileInputStream fin = new FileInputStream(Utility.getProperties("reportDir")+System.getProperty("file.separator")+filename);  		 
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
%>
    