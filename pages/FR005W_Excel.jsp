<%
//95.09.29 create  2495	
//95.10.02 add boaf報表下載 by 2495
//95.10.17 增加 insert A03_LOG BY 2495 	
//99.09.13 RptFR005W_BOAF合併至RptFR005W,套用saveSearchParameter取得參數 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.*,java.io.*" %>
<%@ page import="org.apache.poi.poifs.filesystem.*,org.apache.poi.hssf.usermodel.*" %>
<%@ page import="com.tradevan.util.Utility" %>	
<%@ page import="com.tradevan.util.report.RptFR005W" %>								          
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Enumeration" %>
<%@ page import="com.tradevan.util.DBManager" %>	

<%@include file="./include/Header.include" %>

<%
      response.setContentType("application/msexcel;charset=UTF-8");//以上這行設定本網頁為excel格式的網頁	     
      String Unit = ( request.getParameter("Unit")==null ) ? "1" : Utility.getTrimString(dataMap.get("Unit"));   
      String bank_type = Utility.getTrimString(dataMap.get("bank_type"));   
      String S_YEAR = Utility.getTrimString(dataMap.get("S_YEAR"));   
      String S_MONTH = Utility.getTrimString(dataMap.get("S_MONTH"));   
      String datestate = Utility.getTrimString(dataMap.get("datestate"));   
      String rptStyle = Utility.getTrimString(dataMap.get("rptStyle"));   
      String BANK_NO = Utility.getTrimString(dataMap.get("BANK_NO")); 
      
      String filename=(bank_type.equals("6"))?"全體農會按縣市別平均利率.xls":"全體漁會按縣市別平均利率.xls";   
      String BANK_NAME = ( request.getParameter("BANK_NAME")==null ) ? "" : (String)request.getParameter("BANK_NAME"); 
      
      
 	  //set next jsp 	
	  if(act.equals("view")){
         //以上這行設定傳送到前端瀏覽器時的檔名為test1.xls
         //就是靠這一行，讓前端瀏覽器以為接收到一個excel檔 
      	response.setHeader("Content-disposition","inline; filename=view.xls");
   	  }else if (act.equals("download")){   
      	response.setHeader("Content-Disposition","attachment; filename=download.xls");
   	  }   
   		
   	  try{		    
	  	actMsg=RptFR005W.createRpt(S_YEAR,S_MONTH,Unit,datestate,bank_type,rptStyle,BANK_NO,BANK_NAME);	    
	  	System.out.println("createRpt="+actMsg);
	  	System.out.println("filename="+Utility.getProperties("reportDir")+System.getProperty("file.separator")+filename);
		FileInputStream fin=new FileInputStream(Utility.getProperties("reportDir")+System.getProperty("file.separator")+filename);  		 
		ServletOutputStream out1=response.getOutputStream();           
		byte[] line=new byte[8196];
		int getBytes=0;
		while( ((getBytes=fin.read(line,0,8196)))!=-1 ){		    		
		 		out1.write(line,0,getBytes);
		 		out1.flush();
	  	}		   
		fin.close();
		out1.close();            		      
		//95.10.17 增加 insert A03_LOG BY 2495 
         String acc_code="ALL",atm="0",user_id_c="",user_name_c="",update_type_c="L";
         INSERT_A03_LOG(S_YEAR,S_MONTH,BANK_NO,acc_code,atm,lguser_id,lguser_name,update_type_c); 
	  }catch(Exception e){
	     System.out.println(e.getMessage());
	  }	   		
   	  request.setAttribute("actMsg",actMsg);  
%>
<%@include file="./include/Tail.include" %>

<%!
    private final static String report_no = "FR005W";  
    private final static String nextPgName="/pages/ActMsg.jsp";    
    private final static String QryPgName="/pages/"+report_no+"_Qry.jsp";        
    private final static String RptCreatePgName="/pages/"+report_no+"_Excel.jsp";        
    private final static String LoginErrorPgName="/pages/LoginError.jsp";
%>
<%!
	  //95.10.17 ADD insert A03_LOG BY 2495 
	  private String INSERT_A03_LOG(String m_year,String m_month,String bank_code,String acc_code,String atm,String user_id_c,String user_name_c,String update_type_c) throws Exception{    	
		String sqlCmd = "";		
		String errMsg="";		
		
		try {
				List updateDBSqlList = new LinkedList();								   				   
				//insert A03_LOG===================================================		    
				sqlCmd = "INSERT INTO A03_LOG VALUES ("+m_year
				      	   + ",'" + m_month + "'" 					       
				           + ",'" + bank_code + "'" 
					   + ",'" + acc_code + "'" 					       
				           + ",'" + atm + "'" 
				      	   + ",'" + user_id_c + "'" 					       
				           + ",'" + user_name_c + "'" 
					   + ",sysdate" 					       
				           + ",'" + update_type_c + "')" ;				
				          			           
			 								   
				updateDBSqlList.add(sqlCmd);	
					            		            		
				if(DBManager.updateDB(updateDBSqlList)){
				errMsg = errMsg + "無法寫入A03_log資料";					
				}else{
				  	errMsg = errMsg + "無法寫入A03_log資料";
				}    	   		
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "無法寫入A03_log資料";						
		}	

		return errMsg;
	}  			
%>