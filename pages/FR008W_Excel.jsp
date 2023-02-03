﻿<%
//95.08.21 add 金額單位 by 2295
//95.10.02 fix 下載報表 by 2495
//  95.10.17 增加 insert A05_LOG BY 2495 	
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.*,java.io.*" %>
<%@ page import="org.apache.poi.poifs.filesystem.*,org.apache.poi.hssf.usermodel.*" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.report.FR008W_Excel" %>
<%@ page import="com.tradevan.util.DBManager" %>	
<%
   response.setContentType("application/msexcel;charset=UTF-8");//以上這行設定本網頁為excel格式的網頁
   String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");
   System.out.println("act="+act);
   String bank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");
   /*String BANK_DATA = ( request.getParameter("BANK_NO")==null ) ? "" : (String)request.getParameter("BANK_NO");
   String BANK_NO = BANK_DATA.substring(0,BANK_DATA.indexOf("/"));
   String BANK_NAME = BANK_DATA.substring(BANK_DATA.indexOf("/")+1,BANK_DATA.length());*/
   String BANK_NO = ( request.getParameter("BANK_NO")==null ) ? "" : (String)request.getParameter("BANK_NO");
   String BANK_NAME = ( request.getParameter("BANK_NAME")==null ) ? "" : (String)request.getParameter("BANK_NAME");
   String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? "" : (String)request.getParameter("S_YEAR");
   String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? "" : (String)request.getParameter("S_MONTH");
   String Unit = ( request.getParameter("Unit")==null ) ? "1" : (String)request.getParameter("Unit");//95.08.21 add 金額單位			  
   String muser_name = ( request.getParameter("muser_name")==null ) ? "" : (String)request.getParameter("muser_name");
   String muser_id = ( request.getParameter("muser_id")==null ) ? "" : (String)request.getParameter("muser_id");
   System.out.println("S_YEAR="+S_YEAR);
   System.out.println("S_MONTH="+S_MONTH);
   System.out.println("BANK_NO="+BANK_NO);
   System.out.println("BANK_NAME="+BANK_NAME);
   System.out.println("Unit="+Unit);
   /* 
   String excelAction = ( request.getParameter("excelaction")==null ) ? "" : (String)request.getParameter("excelaction");
   if(excelAction.equals("view")){
      //以上這行設定傳送到前端瀏覽器時的檔名為test1.xls
      //就是靠這一行，讓前端瀏覽器以為接收到一個excel檔
      response.setHeader("Content-disposition","inline; filename=view.xls");
   }else if (excelAction.equals("download")){
    */  
    response.setHeader("Content-Disposition","attachment; filename=download.xls");
   /*}*/

%>
<%
	try{
	    String actMsg = FR008W_Excel.createRpt(S_YEAR,S_MONTH,BANK_NO,BANK_NAME,bank_type,Unit);
	    System.out.println("createRpt="+actMsg);
	    System.out.println("filename="+Utility.getProperties("reportDir")+System.getProperty("file.separator")+"信用部淨值占風險性資產比率.xls");
		FileInputStream fin = new FileInputStream(Utility.getProperties("reportDir")+System.getProperty("file.separator")+"信用部淨值占風險性資產比率.xls");
		ServletOutputStream out1 = response.getOutputStream();
		byte[] line = new byte[8192];
		int getBytes=0;
		while( ((getBytes=fin.read(line,0,8192)))!=-1 ){
			out1.write(line,0,getBytes);
			out1.flush();
    }
			//95.10.17 增加 insert A05_LOG BY 2495 
        String acc_code="ALL",atm="0",atm_name="",user_id_c="",user_name_c="",update_type_c="L";
        INSERT_A05_LOG(S_YEAR,S_MONTH,BANK_NO,acc_code,atm,atm_name,muser_id,muser_name,update_type_c); 
	
		fin.close();
		out1.close();

	}catch(Exception e){
	   System.out.println(e.getMessage());
	}
%>
<%!
	  //95.10.17 ADD insert A05_LOG BY 2495 
	  private String INSERT_A05_LOG(String m_year,String m_month,String bank_code,String acc_code,String atm,String atm_name,String user_id_c,String user_name_c,String update_type_c) throws Exception{    	
		String sqlCmd = "";		
		String errMsg="";		
		
		try {
				List updateDBSqlList = new LinkedList();								   				   
				//insert A05_LOG===================================================		    
				sqlCmd = "INSERT INTO A05_LOG VALUES ("+m_year
				      	   + ",'" + m_month + "'" 					       
				           + ",'" + bank_code + "'" 
					   + ",'" + acc_code + "'" 					       
				           + ",'" + atm + "'" 
				           + ",'" + atm_name + "'" 
				      	   + ",'" + user_id_c + "'" 					       
				           + ",'" + user_name_c + "'" 
					   + ",sysdate" 					       
				           + ",'" + update_type_c + "')" ;				
				          			           
			 								   
				updateDBSqlList.add(sqlCmd);	
					            		            		
				if(DBManager.updateDB(updateDBSqlList)){
				errMsg = errMsg + "無法寫入A05_log資料";					
				}else{
				  	errMsg = errMsg + "無法寫入A05_log資料<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
				}    	   		
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "無法寫入A05_log資料<br>[Exception Error]:<br>"+e.getMessage();								
		}	

		return errMsg;
	}  			
%>