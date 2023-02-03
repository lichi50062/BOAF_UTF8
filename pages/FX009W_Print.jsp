<%
//95/08/17 by 2495
//99.12.07 fix sqlInjection by 2808
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.*,java.io.*" %>
<%@ page import="org.apache.poi.poifs.filesystem.*,org.apache.poi.hssf.usermodel.*" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.ListArray" %>	
<%@ page import="com.tradevan.util.report.RptFR009W" %>								          
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Enumeration" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.DBManager" %>

<%

   response.setContentType("application/msexcel;charset=UTF-8");//以上這行設定本網頁為excel格式的網頁	     
   //String act = ( request.getParameter("act")==null ) ? "view" : (String)request.getParameter("act");		
   String act = "download";	  		  
   String bank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");	    
   String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? "" : (String)request.getParameter("S_YEAR");
   String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? "" : (String)request.getParameter("S_MONTH");
   //String unit = ( request.getParameter("Unit")==null ) ?"1" : (String)request.getParameter("Unit");
   String unit ="1";
   String filename=(bank_type.equals("6"))?"農漁會信用部警示帳戶調查統計明細表.xls":"農漁會信用部警示帳戶調查統計明細表.xls";
   String BankList=""; 
   String bank_no=""; 
   String bank_name="";
   List bn01_BANK_NO = new LinkedList(); 
   List BankList_data = new LinkedList(); 
   String session_bank_type = ( session.getAttribute("bank_type")==null ) ? "" : (String)session.getAttribute("bank_type");	         
   String muser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");	 

   System.out.println("bank_type="+bank_type);	 
	 System.out.println("session_bank_type="+session_bank_type);	 
	 System.out.println("muser_id="+muser_id);	
   if(muser_id.equals("A111111111")||session_bank_type.equals("B")||session_bank_type.equals("2")){
   	 String tbank_no = ( request.getParameter("bank_no")==null ) ? "" : (String)request.getParameter("bank_no");	         
		 bn01_BANK_NO = Get_bn01_BANK_NO(tbank_no);
		 bank_no = (((DataObject)bn01_BANK_NO.get(0)).getValue("bank_no")==null ) ? "" : (String)((DataObject)bn01_BANK_NO.get(0)).getValue("bank_no");
		 bank_name = (((DataObject)bn01_BANK_NO.get(0)).getValue("bank_name")==null ) ? "" : (String)((DataObject)bn01_BANK_NO.get(0)).getValue("bank_name");
	   BankList=bank_no+"+"+bank_no+bank_name;
	   BankList_data = getReportData(BankList);	
   	  	
	 }else if(!muser_id.equals("A111111111")){
	 	 List WTT01_BANK_NO = Get_WTT01_BANK_NO(muser_id);
	 	 String tbank_no = (((DataObject)WTT01_BANK_NO.get(0)).getValue("tbank_no")==null ) ? "" : (String)((DataObject)WTT01_BANK_NO.get(0)).getValue("tbank_no");	 
	 	
		 bn01_BANK_NO = Get_bn01_BANK_NO(tbank_no);
		 bank_no = (((DataObject)bn01_BANK_NO.get(0)).getValue("bank_no")==null ) ? "" : (String)((DataObject)bn01_BANK_NO.get(0)).getValue("bank_no");
		 bank_name = (((DataObject)bn01_BANK_NO.get(0)).getValue("bank_name")==null ) ? "" : (String)((DataObject)bn01_BANK_NO.get(0)).getValue("bank_name");
	   BankList=bank_no+"+"+bank_no+bank_name;
	   BankList_data = getReportData(BankList); 
	 }
	 
   

   System.out.println("S_YEAR="+S_YEAR);
   System.out.println("S_MONTH="+S_MONTH);
   System.out.println("act="+act);
   System.out.println("bank_type="+bank_type);
   System.out.println("bank_no="+bank_no);
   System.out.println("bank_name="+bank_name);
   System.out.println("BankList="+BankList);
   System.out.println("BankList_data="+BankList_data);
 
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
   	  /*	
   		if(!CheckPermission(request)){//無權限時,導向到LoginError.jsp
        	rd = application.getRequestDispatcher( LoginErrorPgName );        
    	}else{
    	*/            
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
	    			actMsg = RptFR009W.createRpt(S_YEAR,S_MONTH,unit,bank_name,bank_type,BankList_data);	    
	    			
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
    private boolean CheckPermission(HttpServletRequest request){//檢核權限    	        		
    	    boolean CheckOK=false;
    	    HttpSession session = request.getSession();            
            Properties permission = ( session.getAttribute("FR034W")==null ) ? new Properties() : (Properties)session.getAttribute("FR034W");				                
            if(permission == null){
              System.out.println("FR34W.permission == null");
            }else{
               System.out.println("FR034W.permission.size ="+permission.size());
               
            }
            //只要有Query的權限,就可以進入畫面
        	if(permission != null && permission.get("Q") != null && permission.get("Q").equals("Y")){            
        	   CheckOK = true;//Query
        	}
        	System.out.println("CheckOk="+CheckOK);        	
        	return CheckOK;
    }   

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
    	String sqlCmd = "select * from WTT01  where muser_id=?"; 	 	
    	paramList.clear() ;
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
%>