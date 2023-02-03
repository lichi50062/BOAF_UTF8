<%
//96.07.12 add by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.*,java.io.*" %>
<%@ page import="org.apache.poi.poifs.filesystem.*,org.apache.poi.hssf.usermodel.*" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.ListArray" %>	
<%@ page import="com.tradevan.util.report.*" %>								          
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Enumeration" %>

<%
   response.setContentType("application/msexcel;charset=UTF-8");//以上這行設定本網頁為excel格式的網頁	     
   String lguser_name = (session.getAttribute("muser_name") == null)?"":(String)session.getAttribute("muser_name"); 
   String tbank_no = (session.getAttribute("tbank_no") == null)?"":(String)session.getAttribute("tbank_no"); 
   
   String act = ( request.getParameter("act")==null ) ? "view" : (String)request.getParameter("act");		
   String bank_no = ( request.getParameter("BANK_NO")==null ) ? tbank_no : (String)request.getParameter("BANK_NO");		  		  
   String bank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");			  
   String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? "" : (String)request.getParameter("S_YEAR");
   String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? "" : (String)request.getParameter("S_MONTH");
   String E_YEAR = ( request.getParameter("E_YEAR")==null ) ? "" : (String)request.getParameter("E_YEAR");
   String E_MONTH = ( request.getParameter("E_MONTH")==null ) ? "" : (String)request.getParameter("E_MONTH");
   //String filename="存款帳戶分級差異化管理統計表.xls";
   //String filename="存款帳戶分級差異化管理統計表_縣市政府.xls";
   //String filename="存款帳戶分級差異化管理統計表_農漁會.xls";
   //String filename="存款帳戶分級差異化管理統計表_全國農業金庫.xls";
   String filename="存款帳戶分級差異化管理統計表";
   String BankList="";
   List BankList_data=null;
   if(request.getParameter("BankList") !=null  && !((String)request.getParameter("BankList")).equals("")){
	  BankList = (String)request.getParameter("BankList");
	  BankList_data = Utility.getReportData(BankList);
	  System.out.println("BankList_data.size()="+BankList_data.size());	  
   }
   if(bank_type.equals("2")) filename += ".xls";//農金局
   if(bank_type.equals("B")) filename += "_縣市政府.xls";//縣市政府
   if(bank_type.equals("1")) filename += "_全國農業金庫.xls";
   if(bank_type.equals("6") || bank_type.equals("7")) filename += "_農漁會.xls";
   System.out.println("S_YEAR="+S_YEAR);
   System.out.println("S_MONTH="+S_MONTH);
   System.out.println("E_YEAR="+E_YEAR);
   System.out.println("E_MONTH="+E_MONTH);
   System.out.println("act="+act);
   System.out.println("bank_type="+bank_type);
   System.out.println("BankList="+BankList); 
   
   RequestDispatcher rd = null;
   boolean doProcess = false;	
   String actMsg = "";		
            
	//set next jsp
	/* 	
	if(act.equals("Qry")){
	   rd = application.getRequestDispatcher( QryPgName + "?bank_type="+bank_type);        
	}else if(act.equals("view")){
       //以上這行設定傳送到前端瀏覽器時的檔名為test1.xls
       //就是靠這一行，讓前端瀏覽器以為接收到一個excel檔 
       	response.setHeader("Content-disposition","inline; filename=view.xls");
   	}else if (act.equals("download")){   
   	*/
       	response.setHeader("Content-Disposition","attachment; filename=download.xls");
   	//}   
   	//if(act.equals("view") || act.equals("download")){
   		try{		    
	    	if(bank_type.equals("2")){//農金局
	    	   actMsg = RptFR045W.createRpt(S_YEAR,S_MONTH,bank_type,lguser_name,BankList_data);//農金局	    
	    	}else if(bank_type.equals("B")){//地方主管機關
	    	   actMsg = RptFR045W_bank_b.createRpt(S_YEAR,S_MONTH,bank_no);//地方主管機關
	    	}else if(bank_type.equals("6") || bank_type.equals("7")){//農.漁會 
       		   actMsg = RptFR045W_bank_67.createRpt(S_YEAR,S_MONTH,E_YEAR,E_MONTH,bank_no);//農漁會
       		}else if(bank_type.equals("1")){//全國農業金庫
       		   actMsg = RptFR045W_bank_1.createRpt(S_YEAR,S_MONTH,E_YEAR,E_MONTH);//全國農業金庫
       		}
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
   	//}
   		
   	request.setAttribute("actMsg",actMsg);    	
	try {
       	//forward to next present jsp
       	rd.forward(request, response);
    } catch (NullPointerException npe) {}   
%>

<%!
    private final static String nextPgName = "/pages/ActMsg.jsp";    
    private final static String QryPgName = "/pages/FR045W_Qry.jsp";              
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
%>