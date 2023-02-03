<%
//95.04.24 add 農(漁)會信用部逾期放款統計表 by 2295
//95.10.02 add 下載報表 by 2495	
//95.10.17 增加 insert A06_LOG BY 2495 
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.report.RptFR037W" %>								          
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
<%@ page import="com.tradevan.util.DBManager" %>	

<%
	response.setContentType("application/msexcel;charset=UTF-8");//以上這行設定本網頁為excel格式的網頁       
	
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");			    	
	String bank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");			    	
	/*
	String BANK_DATA = ( request.getParameter("BankListSrc")==null ) ? "" : (String)request.getParameter("BankListSrc");			  
    String BANK_NO = (BANK_DATA.startsWith("全體"))?"ALL":BANK_DATA.substring(0,7);
    String BANK_NAME = (BANK_DATA.startsWith("全體"))?BANK_DATA.substring(0,BANK_DATA.length()):BANK_DATA.substring(7,BANK_DATA.length());
  */
  String BANK_NO = ( request.getParameter("BANK_NO")==null ) ? "" : (String)request.getParameter("BANK_NO");
   String BANK_NAME = ( request.getParameter("BANK_NAME")==null ) ? "" : (String)request.getParameter("BANK_NAME");
    String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? "" : (String)request.getParameter("S_YEAR");
    String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? "" : (String)request.getParameter("S_MONTH");
    String Unit = ( request.getParameter("Unit")==null ) ? "1" : (String)request.getParameter("Unit");//金額單位			  
    String filename="農漁會信用部逾期放款統計表";
    String muser_name = ( request.getParameter("muser_name")==null ) ? "" : (String)request.getParameter("muser_name");
    String muser_id = ( request.getParameter("muser_id")==null ) ? "" : (String)request.getParameter("muser_id"); 
    System.out.print("FR037W_Excel.S_YEAR="+S_YEAR);
    System.out.print(":S_MONTH="+S_MONTH);   
    System.out.print(":BANK_NO="+BANK_NO);
    System.out.print(":BANK_NAME="+BANK_NAME);
    System.out.println(":Unit="+Unit);
   /*
    String excelAction = ( request.getParameter("excelaction")==null ) ? "" : (String)request.getParameter("excelaction");			    	
    
    

    if(excelAction.equals("view")){
      //以上這行設定傳送到前端瀏覽器時的檔名為test1.xls
      //就是靠這一行，讓前端瀏覽器以為接收到一個excel檔 
      response.setHeader("Content-disposition","inline; filename=view.xls");
    }else if (excelAction.equals("download")){   
     */ response.setHeader("Content-Disposition","attachment; filename=download.xls");
    /*}*/
    
    try{	
	    String actMsg = RptFR037W.createRpt(S_YEAR,S_MONTH,BANK_NO,bank_type,BANK_NAME,Unit);	    
	    System.out.println("createRpt="+actMsg);
	    System.out.println("filename="+Utility.getProperties("reportDir")+System.getProperty("file.separator")+filename+".xls");
		FileInputStream fin = new FileInputStream(Utility.getProperties("reportDir")+System.getProperty("file.separator")+filename+".xls");  		 
		ServletOutputStream out1 = response.getOutputStream();           
		byte[] line = new byte[8196];
		int getBytes=0;
		while( ((getBytes=fin.read(line,0,8196)))!=-1 ){		    		
			out1.write(line,0,getBytes);
			out1.flush();
				    }
		//95.10.17 增加 insert A06_LOG BY 2495 
        String acc_code="ALL",atm_3month="0",atm_6month="0",atm_1year="0",atm_2year="0",atm_over2year="0",atm_total="0",user_id_c="",user_name_c="",update_type_c="L";
        INSERT_A06_LOG(S_YEAR,S_MONTH,BANK_NO,acc_code,atm_3month,atm_6month,atm_1year,atm_2year,atm_over2year,atm_total,muser_id,muser_name,update_type_c); 

		fin.close();
		out1.close();            		      
		
	}catch(Exception e){
	   System.out.println(e.getMessage());
	}		
%>
<%!
	  //95.10.17 ADD insert A06_LOG BY 2495 
	  private String INSERT_A06_LOG(String m_year,String m_month,String bank_code,String acc_code,String atm_3month,String atm_6month,String atm_1year,String atm_2year,String atm_over2year,String atm_total,String user_id_c,String user_name_c,String update_type_c) throws Exception{    	
		String sqlCmd = "";		
		String errMsg="";		
		
		try {
				List updateDBSqlList = new LinkedList();								   				   
				//insert A06_LOG===================================================		    
				sqlCmd = "INSERT INTO A06_LOG VALUES ("+m_year
				      	   + ",'" + m_month + "'" 					       
				           + ",'" + bank_code + "'" 
					         + ",'" + acc_code + "'" 					       
				           + ",'" + atm_3month + "'" 
				           + ",'" + atm_6month + "'" 
				           + ",'" + atm_1year + "'" 
				           + ",'" + atm_2year + "'" 
				           + ",'" + atm_over2year + "'" 
				           + ",'" + atm_total + "'"
				      	   + ",'" + user_id_c + "'" 					       
				           + ",'" + user_name_c + "'" 
					         + ",sysdate" 					       
				           + ",'" + update_type_c + "')" ;				
				          			           
			 								   
				updateDBSqlList.add(sqlCmd);	
					            		            		
				if(DBManager.updateDB(updateDBSqlList)){
				errMsg = errMsg + "無法寫入A06_log資料";					
				}else{
				  	errMsg = errMsg + "無法寫入A06_log資料";
				}    	   		
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "無法寫入A06_log資料";								
		}	

		return errMsg;
	}  			
%>    