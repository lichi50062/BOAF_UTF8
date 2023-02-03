<%
//95.09.29 create  2495	
//95.10.02 add 下載報表 by 2495
//95.10.17 增加 insert A01_LOG BY 2495 						  																	
//95.12.13 fix 讀取目前的BANK_TYPE by 2295
//102.01.21 fix A02_LOG 增加amt_name by 2295
//102.02.04 fix MIS讀取參數不同 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.*,java.io.*" %>
<%@ page import="org.apache.poi.poifs.filesystem.*,org.apache.poi.hssf.usermodel.*" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.report.RptFR0066W" %>
<%@ page import="com.tradevan.util.DBManager"%>	
<%
   	System.out.println("FR006W_Excel.jsp Start...");
   	//95.12.13 fix 讀取目前的BANK_TYPE
   	String bank_type = ( request.getParameter("BANK_TYPE")==null ) ? "" : (String)request.getParameter("BANK_TYPE");//BOAF讀取
   	if(bank_type.equals("")){
	bank_type = ( request.getParameter("bankType")==null ) ? "" : (String)request.getParameter("bankType");//MIS讀取
	}
	System.out.println("bank_type ="+bank_type);

//   	response.setContentType("APPLICATION/msword;charset=Big5");//以上這行設定本網頁為excel格式的網頁
//	response.setContentType("application/octet-stream; charset=iso-8859-1");//以上這行設定本網頁為excel格式的網頁
        response.setContentType("application/octet-stream");//以上這行設定本網頁為excel格式的網頁
   	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");
   	System.out.println("act="+act);
	  String unit =request.getParameter("unit")==null?"":request.getParameter("unit");
	  System.out.println("unit = "+unit);
   	//String BANK_DATA = ( request.getParameter("BANK_NO")==null ) ? "" : (String)request.getParameter("BANK_NO");
   	//String BANK_NO = BANK_DATA.substring(0,BANK_DATA.indexOf("/"));
   	//String BANK_NAME = BANK_DATA.substring(BANK_DATA.indexOf("/")+1,BANK_DATA.length());
   	String BANK_NO = ( request.getParameter("BANK_NO")==null ) ? "" : (String)request.getParameter("BANK_NO");//BOAF讀取
   	if(BANK_NO.equals("") || BANK_NO == null){
   	   BANK_NO = ( request.getParameter("tbank")==null ) ? "" : (String)request.getParameter("tbank");//MIS讀取
	}
   	String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? "" : (String)request.getParameter("S_YEAR");
   	String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? "" : (String)request.getParameter("S_MONTH");
   	String muser_id = ( request.getParameter("muser_id")==null ) ? "" : (String)request.getParameter("muser_id");
    String muser_name = ( request.getParameter("muser_name")==null ) ? "" : (String)request.getParameter("muser_name");
   	System.out.println("S_YEAR="+S_YEAR);
   	System.out.println("S_MONTH="+S_MONTH);
   	System.out.println("muser_id="+muser_id);
    System.out.println("muser_name="+muser_name); 
   	System.out.println("BANK_NO="+BANK_NO);
   	//System.out.println("BANK_NAME="+BANK_NAME);

		/*
   	String wordAction = request.getParameter("wordaction")==null  ? "" : request.getParameter("wordaction");
   	System.out.println("wordAction.equals ="+wordAction);
   	
   	if(wordAction.equals("view")){
   	   //以上這行設定傳送到前端瀏覽器時的檔名為test1.xls
   	   //就是靠這一行，讓前端瀏覽器以為接收到一個excel檔
   	   response.setHeader("Content-disposition","inline; filename=view.rtf");
   	}else if (wordAction.equals("download")){
   	*/
   	   response.setHeader("Content-Disposition","attachment; filename=download.rtf");
   	//}
%>
<%
	try{
	    String actMsg = RptFR0066W.createRpt(S_YEAR,S_MONTH,BANK_NO,null,bank_type,null,false);
	    System.out.println("createRpt="+actMsg);
	    FileInputStream fin = null;

	    if(bank_type.equals("6")) {
	    	System.out.println("filename="+Utility.getProperties("reportDir")+System.getProperty("file.separator")+"農會各項法定比率表.rtf");
	    	fin = new FileInputStream(Utility.getProperties("reportDir")+System.getProperty("file.separator")+"農會各項法定比率表.rtf");
	    }else {
	    	System.out.println("filename="+Utility.getProperties("reportDir")+System.getProperty("file.separator")+"漁會各項法定比率表.rtf");
	    	fin = new FileInputStream(Utility.getProperties("reportDir")+System.getProperty("file.separator")+"漁會各項法定比率表.rtf");

	    }
		ServletOutputStream out1 = response.getOutputStream();
//                FileOutputStream out2 = new FileOutputStream("")
		byte[] line = new byte[1024];
		int getBytes=0;
		while( ((getBytes=fin.read(line,0,1024)))!=-1 ){
			out1.write(line,0,getBytes);
			out1.flush();
	    }
		//關閉檔案
		fin.close();
		out1.close();
    //95.10.17 增加 insert A02_LOG BY 2495 	
        String acc_code="ALL",atm="0",user_id_c="",user_name_c="",update_type_c="L";
        INSERT_A02_LOG(S_YEAR,S_MONTH,BANK_NO,acc_code,atm,muser_id,muser_name,update_type_c); 
	}catch(Exception e){
	   System.out.println(e.getMessage());
	}
	System.out.println("FR006W_Excel.jsp End...");
%>
<%!
	  //95.10.17 ADD insert A02_LOG BY 2495 
	  private String INSERT_A02_LOG(String m_year,String m_month,String bank_code,String acc_code,String atm,String user_id_c,String user_name_c,String update_type_c) throws Exception{    	
		String sqlCmd = "";		
		String errMsg="";		
		
		try {
				List updateDBSqlList = new LinkedList();								   				   
				//insert A02_LOG===================================================		    
				sqlCmd = "INSERT INTO A02_LOG VALUES ("+m_year
				      	   + ",'" + m_month + "'" 					       
				           + ",'" + bank_code + "'" 
					   + ",'" + acc_code + "'" 					       
				           + ",'" + atm + "'" 
				      	   + ",'" + user_id_c + "'" 					       
				           + ",'" + user_name_c + "'" 
					   + ",sysdate" 					       
				           + ",'" + update_type_c + "','')" ;				
				          			           
			 								   
				updateDBSqlList.add(sqlCmd);	
					            		            		
				if(DBManager.updateDB(updateDBSqlList)){
				errMsg = errMsg + "無法寫入A02_log資料";					
				}else{
				  	errMsg = errMsg + "無法寫入A02_log資料";;
				}    	   		
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "無法寫入A02_log資料<br>[Exception Error]";								
		}	

		return errMsg;
	}  			
%>