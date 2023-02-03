<%
/*
//95/08/17 by 2495
//99.12.07 fix sqlInjection by 2808
//100.02.16 fix wlx01.bn01區分99/100年度   by 2479
//          fix cd01_table 查詢年度100年以前，查cd01_99，100年以後查cd01
*/            
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.*,java.io.*" %>
<%@ page import="org.apache.poi.poifs.filesystem.*,org.apache.poi.hssf.usermodel.*" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.ListArray" %>	
<%@ page import="com.tradevan.util.report.RptFR009WB" %>								          
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Enumeration" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.DBManager" %>

<%
	 System.out.print("FX009W_Print_B.jsp--------------start");
   response.setContentType("application/msexcel;charset=UTF-8");//以上這行設定本網頁為excel格式的網頁	     
   //String act = ( request.getParameter("act")==null ) ? "view" : (String)request.getParameter("act");		
   String act = "download";	  		  
   String bank_type = ( session.getAttribute("bank_type")==null ) ? "" : (String)session.getAttribute("bank_type");	    
   String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? "" : (String)request.getParameter("S_YEAR");
   String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? "" : (String)request.getParameter("S_MONTH");
   
    //100.04.13 add 查詢年度100年以前.縣市別不同===============================
 	String cd01_table = (Integer.parseInt(S_YEAR) < 100)?"cd01_99":"cd01"; 
 	String wlx01_m_year = (Integer.parseInt(S_YEAR) < 100)?"99":"100"; 
 	//=====================================================================   
   
   //String unit = ( request.getParameter("Unit")==null ) ?"1" : (String)request.getParameter("Unit");
   String unit ="1";
   String filename=(bank_type.equals("6"))?"農漁會信用部警示帳戶調查統計明細表.xls":"農漁會信用部警示帳戶調查統計明細表.xls";
   String BankList=""; 
   String bank_no=""; 
   String tbank_no="";
   String bank_name="";
   String hsien_id="";
   String hsien_name="";
   List bn01_BANK_NO = new LinkedList(); 
   List hsien_id_data = new LinkedList(); 
   List hsien_name_data = new LinkedList(); 
   List bank_data = new LinkedList(); 
   List bank_list = new LinkedList(); 
   
   String muser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");	          
   if(bank_type.equals("B")){
   	System.out.print("bank_type.equals : B");
   	List WTT01_BANK_NO = Get_WTT01_BANK_NO(muser_id);	 
	 	tbank_no = (((DataObject)WTT01_BANK_NO.get(0)).getValue("tbank_no")==null ) ? "" : (String)((DataObject)WTT01_BANK_NO.get(0)).getValue("tbank_no");	  
	 	System.out.print("tbank_no : "+tbank_no);
	 	bn01_BANK_NO = Get_bn01_BANK_NO(tbank_no);
	 	bank_no = (((DataObject)bn01_BANK_NO.get(0)).getValue("bank_no")==null ) ? "" : (String)((DataObject)bn01_BANK_NO.get(0)).getValue("bank_no");	
	 	System.out.print("bank_no : "+bank_no);
	 	hsien_id_data = Get_WLX01_hsien_id(bank_no,wlx01_m_year);
	 	hsien_id = (((DataObject)hsien_id_data.get(0)).getValue("hsien_id")==null ) ? "" : (String)((DataObject)hsien_id_data.get(0)).getValue("hsien_id");
	 	System.out.print("hsien_id : "+hsien_id);
	 	hsien_name_data = Get_CD01_hsien_name(hsien_id,cd01_table);
	 	hsien_name = (((DataObject)hsien_name_data.get(0)).getValue("hsien_name")==null ) ? "" : (String)((DataObject)hsien_name_data.get(0)).getValue("hsien_name");
    System.out.print("hsien_name : "+hsien_name);
	 	bank_data = Get_WLX01_bank_data(hsien_id);
	 	System.out.print("bank_data.size() : "+bank_data.size());
	 	for(int i=0;i<bank_data.size();i++){
	 			bank_no = (((DataObject)bank_data.get(i)).getValue("bank_no")==null ) ? "" : (String)((DataObject)bank_data.get(i)).getValue("bank_no");	 		 			
	 		 	bank_name = (((DataObject)bank_data.get(i)).getValue("bank_name")==null ) ? "" : (String)((DataObject)bank_data.get(i)).getValue("bank_name");	 		 			
	 			if(i==0)
	 			{
	 				BankList=bank_no+"+"+bank_name;
	 			}else{
	 			  BankList=BankList+","+bank_no+"+"+bank_no+bank_name;	
	 			}
	 	}
	 	  System.out.print("BankList : "+BankList);
	 		bank_list = getReportData(BankList);   
	 	}
	 		
	 	if(muser_id.equals("A111111111")||bank_type.equals("2")){
	 	  System.out.println("bank_type.equals : 2");
	 		tbank_no = ( request.getParameter("bank_no")==null ) ? "" : (String)request.getParameter("bank_no"); 
	 		System.out.println("tbank_no="+tbank_no);	 		
	 		hsien_id_data = Get_WLX01_hsien_id(tbank_no,wlx01_m_year); //Get_WLX01_hsien_id(tbank_no);
	 		hsien_id = (((DataObject)hsien_id_data.get(0)).getValue("hsien_id")==null ) ? "" : (String)((DataObject)hsien_id_data.get(0)).getValue("hsien_id");
	 		System.out.println("hsien_id : "+hsien_id);
	 		hsien_name_data = Get_CD01_hsien_name(hsien_id,cd01_table);//Get_CD01_hsien_name(hsien_id);
	 		hsien_name = (((DataObject)hsien_name_data.get(0)).getValue("hsien_name")==null ) ? "" : (String)((DataObject)hsien_name_data.get(0)).getValue("hsien_name");
	    System.out.println("hsien_name : "+hsien_name);
		 	bank_data = Get_WLX01_bank_data(hsien_id);
		 	System.out.println("bank_data.size() : "+bank_data.size());
		 	for(int i=0;i<bank_data.size();i++){
	 			bank_no = (((DataObject)bank_data.get(i)).getValue("bank_no")==null ) ? "" : (String)((DataObject)bank_data.get(i)).getValue("bank_no");	 		 			
	 		 	bank_name = (((DataObject)bank_data.get(i)).getValue("bank_name")==null ) ? "" : (String)((DataObject)bank_data.get(i)).getValue("bank_name");	 		 			
	 			if(i==0)
	 			{
	 				BankList=bank_no+"+"+bank_name;
	 			}else{
	 			  BankList=BankList+","+bank_no+"+"+bank_no+bank_name;	
	 			}
	 		}
	 	  System.out.println("BankList : "+BankList);
	 		bank_list = getReportData(BankList);   
	 	}	
  
   
   RequestDispatcher rd = null;
   boolean doProcess = false;	
   String actMsg = "";		
   //取得session資料,取得成功時,才繼續往下執行===================================================
   if(session.getAttribute("muser_id") == null){//session timeout		
      System.out.println("FR034W login timeout");   
	   rd = application.getRequestDispatcher( "/pages/reLogin.jsp?url=LoginError.jsp?timeout=true" );         	   
	   try{
          rd.forward(request,response);
       }catch(Exception e){
          System.out.println("forward Error:"+e+e.getMessage());
       }
   }else{
      doProcess = true;
   } 
        
   if(doProcess){//若muser_id資料時,表示登入成功====================================================================	
   	        
	    	//set next jsp 	
	    	if(act.equals("Qry")){
	    	   rd = application.getRequestDispatcher( QryPgName + "?bank_type="+bank_type);        
	    	}else if(act.equals("view")){
      		   //以上這行設定傳送到前端瀏覽器時的檔名為test1.xls
      		   //就是靠這一行，讓前端瀏覽器以為接收到一個excel檔 
      			response.setHeader("Content-disposition","inline; filename=view.xls");
   			}else if (act.equals("download")){   
      			response.setHeader("Content-Disposition","attachment; filename=download.xls");
   			}   
   			if(act.equals("view") || act.equals("download")){
   				try{
   					System.out.print("FX009W_Print_B.jsp--------------preexcel");	
   					bank_type="6";
	    			actMsg = RptFR009WB.createRpt(S_YEAR,S_MONTH,unit,bank_type,bank_list,hsien_name);
	    			System.out.print("FX009W_Print_B.jsp--------------endexcel");		 
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
   			}
   	/*	}*/
   		request.setAttribute("actMsg",actMsg);    	
		try {
        	//forward to next present jsp
        	rd.forward(request, response);
    	} catch (NullPointerException npe) {
    	}
    }//end of doProcess

%>


<%!
    private final static String nextPgName = "/pages/ActMsg.jsp";    
    private final static String QryPgName = "/pages/FR034W_Qry.jsp";              
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    List paramList = new ArrayList() ;
   
//將選擇的bank_list取出
	private List getReportData(String rptData){
	    List rptList = new LinkedList();
	    StringTokenizer paserData = null;
	    List rptDetail = null;
		try{
			StringTokenizer paser = new StringTokenizer(rptData.trim(),",");			
	        while (paser.hasMoreTokens()){
	            paserData = new StringTokenizer(paser.nextToken(","),"+");
	            rptDetail = new LinkedList();
	            while (paserData.hasMoreTokens()){
	                rptDetail.add(paserData.nextToken());  
	            }//end of have "+" data	            
	            rptList.add(rptDetail);
			}//end of have "," data 
		}catch(Exception e){
			System.out.println("getReportData Error:"+e+e.getMessage());
		}
		System.out.println("rptData:"+rptData);
		System.out.println("rptList:"+rptList);
		return rptList;
	}
	
	private List Get_WTT01_BANK_NO(String muser_id){
    	//查詢條件    
    	String sqlCmd = "select * from WTT01  where muser_id=? ";
    	paramList.clear( ) ;
    	paramList.add(muser_id) ;
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
        return dbData;
    }
    
    private List Get_bn01_BANK_NO(String bank_no){
    	//查詢條件    
    	String sqlCmd = "select * from bn01  where bank_no=? "; 	
    	paramList.clear() ;
    	paramList.add(bank_no) ;
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
        return dbData;
    }
    
    private List Get_WLX01_hsien_id(String bank_no ,String year){
    	//查詢條件    
    	String sqlCmd = "select * from WLX01  where m2_name=? and  m_year=?"; 
    	paramList.clear() ;
    	paramList.add(bank_no) ;
    	paramList.add(year) ;
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
        return dbData;
    }
    
    private List Get_CD01_hsien_name(String hsien_id,String cd01_table){
    	//查詢條件    
    	String sqlCmd = "select * from "+cd01_table+"  where hsien_id=? ";
    	paramList.clear() ;
    	paramList.add(hsien_id) ;
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
        return dbData;
    }
    
    private List Get_WLX01_bank_data(String hsien_id){
    	//查詢條件    
    	String sqlCmd = "select * from (select bank_no from wlx01   where hsien_id=? ) a , BA01 where BA01.BANK_NO=a.BANK_NO";
    	paramList.clear() ;
    	paramList.add(hsien_id) ;
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
        return dbData;
    }
     
%>