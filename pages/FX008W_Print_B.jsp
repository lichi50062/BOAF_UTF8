<%
/*
  *   created on 95.8.21  by lilic0c0 2495
  *   fix sqlInejction by 2808
  */ 
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.*,java.io.*" %>
<%@ page import="org.apache.poi.poifs.filesystem.*,org.apache.poi.hssf.usermodel.*" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.report.RtpFX008W_PrintB" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Enumeration" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.DBManager" %>

<%
	System.out.println("FX008W_Print.jsp Program Start...");
	
	response.setHeader("Content-Disposition","attachment; filename=download.xls");
	response.setContentType("application/msexcel;charset=UTF-8");//以上這行設定本網頁為excel格式的網頁
		
	String act =  "download" ;
	String bank_type = ( session.getAttribute("bank_type")==null ) ? "" : (String)session.getAttribute("bank_type");
	String bank_name = ( request.getParameter("bank_name")==null ) ? "" : (String)request.getParameter("bank_name");
	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? "" : (String)request.getParameter("S_YEAR");
	String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? "" : (String)request.getParameter("S_MONTH");
								
	 String filename="縣市政府承受擔保品申請延長處分期限審核表.xls";
	 String unit ="1";   
	 String BankList=""; 
   String bank_no=""; 
   String tbank_no="";   
   String hsien_id="";
   List bn01_BANK_NO = new LinkedList(); 
   List hsien_id_data = new LinkedList(); 
   List hsien_name_data = new LinkedList(); 
   String muser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");	          
   if(bank_type.equals("B")){
   	System.out.print("bank_type.equals : B");
   	List WTT01_BANK_NO = Get_WTT01_BANK_NO(muser_id);	 
	 	tbank_no = (((DataObject)WTT01_BANK_NO.get(0)).getValue("tbank_no")==null ) ? "" : (String)((DataObject)WTT01_BANK_NO.get(0)).getValue("tbank_no");	  
	 	bn01_BANK_NO = Get_bn01_BANK_NO(tbank_no);
	 	bank_no = (((DataObject)bn01_BANK_NO.get(0)).getValue("bank_no")==null ) ? "" : (String)((DataObject)bn01_BANK_NO.get(0)).getValue("bank_no");	
	 	hsien_id_data = Get_WLX01_hsien_id(bank_no);
	 	hsien_id = (((DataObject)hsien_id_data.get(0)).getValue("hsien_id")==null ) ? "" : (String)((DataObject)hsien_id_data.get(0)).getValue("hsien_id");
	 	hsien_name_data = Get_CD01_hsien_name(hsien_id);
	 	bank_name = (((DataObject)hsien_name_data.get(0)).getValue("hsien_name")==null ) ? "" : (String)((DataObject)hsien_name_data.get(0)).getValue("hsien_name");
	 	
	 	}
	 		
	 	if(bank_type.equals("2")){
	 	  System.out.print("bank_type.equals : 2");
	 		bank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");	 
	 		System.out.print("bank_no="+bank_no);
	 		hsien_id_data = Get_WLX01_hsien_id(bank_no);
	 		hsien_id = (((DataObject)hsien_id_data.get(0)).getValue("hsien_id")==null ) ? "" : (String)((DataObject)hsien_id_data.get(0)).getValue("hsien_id");
	 		hsien_name_data = Get_CD01_hsien_name(hsien_id);
	 		bank_name = (((DataObject)hsien_name_data.get(0)).getValue("hsien_name")==null ) ? "" : (String)((DataObject)hsien_name_data.get(0)).getValue("hsien_name");

	 	}	
	
	
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
	    String actMsg = RtpFX008W_PrintB.createRpt(S_YEAR,S_MONTH,unit,bank_name,bank_type,bank_no);	
	    System.out.println("createRpt="+actMsg);
	    FileInputStream fin = null;
		
			System.out.println("filename="+Utility.getProperties("reportDir")+System.getProperty("file.separator")+"縣市政府承受擔保品申請延長處分期限審核表.xls");
	   	fin = new FileInputStream(Utility.getProperties("reportDir")+System.getProperty("file.separator")+"縣市政府承受擔保品申請延長處分期限審核表.xls");
		
		
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
<%!
    List paramList =new ArrayList() ;
	private List Get_WTT01_BANK_NO(String muser_id){
    	//查詢條件    
    	String sqlCmd = "select * from WTT01 " + " where muser_id=? "; 	 
    	paramList.clear() ;
    	paramList.add(muser_id ) ;
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
        return dbData;
    }
    
    private List Get_bn01_BANK_NO(String bank_no){
    	//查詢條件    
    	String sqlCmd = "select * from bn01 " + " where bank_no=? "; 
    	paramList.clear() ;
    	paramList.add(bank_no ) ;
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
        return dbData;
    }
    
    private List Get_WLX01_hsien_id(String bank_no){
    	//查詢條件    
    	String sqlCmd = "select * from WLX01 " + " where m2_name=? "; 	
    	paramList.clear() ;
    	paramList.add(bank_no ) ;
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
        return dbData;
    }
    
    private List Get_CD01_hsien_name(String hsien_id){
    	//查詢條件    
    	String sqlCmd = "select * from CD01 " + " where hsien_id=? "; 	 	
    	paramList.clear() ;
    	paramList.add(hsien_id ) ;
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
        return dbData;
    }
%>





