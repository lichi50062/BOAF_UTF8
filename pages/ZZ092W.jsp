<%
// 95.09.12 create by 2495
// 96.01.04 fix 轉檔前.檢查資料是否存在資料庫中.若存在.先刪除.再轉檔 by 2295
// 96.01.10 fix 農漁會支票存款資料.加上刪除wml01 by 2295
// 96.01.16 add 是否需重新至金庫取檔 by 2295
// 96.09.28 fix 牌告利率改為月報 by 2295
//103.04.02 add 專案農貸明細資料/專案農貸貸款項目資料  by 2295
//103.05.19 add 金庫取檔使用sftp by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.List" %>
<%@ page import="java.util.ArrayList" %>
<%@ page import="java.util.LinkedList" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Enumeration" %>
<%@ page import="java.util.HashMap" %>
<%@ page import="java.util.Hashtable" %>
<%@ page import="java.util.StringTokenizer" %>
<%@ page import="com.tradevan.util.ftp.AgriBankFTP" %>

<%
	RequestDispatcher rd = null;
	String actMsg = "";	
	String alertMsg = "";	
	String webURL = "";	
	String webURL_Y = "";
	boolean doProcess = false;	

	//取得session資料,取得成功時,才繼續往下執行===================================================
	if(session.getAttribute("muser_id") == null){//session timeout	
      System.out.println("ZZ092W login timeout");   
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
	
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");						
	String m_year = ( request.getParameter("year")==null ) ? "" : (String)request.getParameter("year");	
	String m_month = ( request.getParameter("month")==null ) ? "" : (String)request.getParameter("month");	  
	String getFiles = ( request.getParameter("getFiles")==null ) ? "" : (String)request.getParameter("getFiles");	  
	
	//登入者資訊
	String lguser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
	String lguser_name = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");		
	String lguser_type = ( session.getAttribute("muser_type")==null ) ? "" : (String)session.getAttribute("muser_type");					
	String tbank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");					
	session.setAttribute("nowtbank_no",null);//94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除======
	String program_id = "";
	//金庫主機
	
	String FTP_SERVER = "59.124.54.12";	
	//String FTP_SERVER = "59.124.54.11";
 	String FTP_USER = "TNBOAF";
 	String FTP_PASSWORD = "TNBOAF123";
 	String FTP_DIRECTORY = "/out/"; 	
 	String LOCAL_DIRECTORY = "C:\\Sun\\WebServer6.1\\BOAF\\AgriBankData";
 	
	//公司測試主機
	/*
	String FTP_SERVER = "172.20.5.22";
	String FTP_USER = "pboafmgr";
	String FTP_PASSWORD = "tvpboafmgr";	
	String FTP_DIRECTORY = "/export/home/pboafmgr/test/";	\
	
	String FTP_SERVER = "10.89.8.170";
	String FTP_USER = "pdntmgr";
	String FTP_PASSWORD = "AAllen6812";
	String FTP_DIRECTORY = "APBOAF";	
	//String LOCAL_DIRECTORY = "C:\\Sun\\WebServer6.1\\BOAF\\AgriBankData";	
	String LOCAL_DIRECTORY = "D:\\workProject\\BOAF\\AgriBankData";
	*/
	String TargetFile = "";
	StringBuffer sqlCmd = new StringBuffer() ;;
	List paramList =new ArrayList () ;
	//String sqlCmd_delete=" delete ";
	StringBuffer sqlCmd_delete = new StringBuffer() ;
	sqlCmd_delete.append(" delete ") ;
	List dbData = new LinkedList();
	//List updateDBSqlList = new LinkedList();
    if(!CheckPermission(request)){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );        
    }else{            
        if(act.equals("List")){	                
           rd = application.getRequestDispatcher( ListPgName );                	              	        	        	    	            
    	}
    	if(act.equals("download")){
    	   String Star[] = request.getParameterValues("candidate");    	   
    	   System.out.println("Star.length="+Star.length);
		   if(m_month.length()==1) m_month="0"+m_month;
 		   m_year = Integer.toString(Integer.parseInt(m_year)+1911); 
 		   sqlCmd.append("select count(*) as countdata from ");
		   for(int i=0;i<Star.length;i++){						      
			   System.out.println("Star.length="+Star.length); 
			   System.out.println("Star[i]="+Star[i]);  
			   /*96.09.28 牌告利率改為月報											
			   if(Star[i].equals("INT")){
			      if(Integer.parseInt(m_month) >= 1  && Integer.parseInt(m_month) <= 3) m_month ="01";
			      if(Integer.parseInt(m_month) >= 4  && Integer.parseInt(m_month) <= 6) m_month ="02";
			      if(Integer.parseInt(m_month) >= 7  && Integer.parseInt(m_month) <= 9) m_month ="03";
			      if(Integer.parseInt(m_month) >= 10 && Integer.parseInt(m_month) <= 12) m_month ="04";			     
			   }
			   */
			   TargetFile = Star[i]+m_year+m_month+".DAT"; 
 			   if(Star[i].equals("ATM")){//農漁會金融卡發卡情形及ATM裝設情形統計	 				  
 				  sqlCmd.append(" wlx05_m_atm ");
 				  sqlCmd_delete.append(" wlx05_m_atm ");
 			   }else if(Star[i].equals("CHK")){//農漁會支票存款資料
 			      sqlCmd.append(" WLX07_M_CHECKBANK "); 
 			      sqlCmd_delete.append(" WLX07_M_CHECKBANK ");				  
 			   }else if(Star[i].equals("FGN")){//外國人存款資料
 			      sqlCmd.append(" F01 "); 				
 			      sqlCmd_delete.append(" F01 ");   				  
 			   }else if(Star[i].equals("INT")){//農漁會信用部牌告利率申報資料
 			      sqlCmd.append(" WLX_S_RATE " );
 			      sqlCmd_delete.append(" WLX_S_RATE ");
 			   }else if(Star[i].equals("FRM")){//專案農貸
 			      sqlCmd.append(" AGRI_LOAN " );
 			      sqlCmd_delete.append(" AGRI_LOAN ");   
 			   }else if(Star[i].equals("ITEM")){//貸款項目
 			      sqlCmd.append(" AGRI_LOAN_ITEM " );
 			      sqlCmd_delete.append(" AGRI_LOAN_ITEM ");    
 			   }
 			   System.out.println("TargetFile="+TargetFile); 
 			   if(!Star[i].equals("ITEM")){//貸款項目
 			      sqlCmd.append(" where m_year= ? ");
 			      sqlCmd_delete.append(" where m_year=?");
 			      paramList.add(Integer.toString(Integer.parseInt(m_year)-1911));
 			   }
 			   if(Star[i].equals("INT")){
 			      sqlCmd.append(" and m_quarter=?"); 				  
 			      sqlCmd_delete.append(" and m_quarter=?");
 			      paramList.add(m_month) ;
 			   }else if(!Star[i].equals("ITEM")){//貸款項目
 			      sqlCmd.append(" and m_month= ? "); 				  
 			      sqlCmd_delete.append(" and m_month= ? "); 		
 			      paramList.add(m_month) ;
 			   }    			   
 			   dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"countdata"); 
 			   if(dbData != null){
 				  if(!((((DataObject)dbData.get(0)).getValue("countdata")).toString()).equals("0")){
 					   System.out.println(" have "+(((DataObject)dbData.get(0)).getValue("countdata")).toString() +" datas");
 					   //updateDBSqlList.add(sqlCmd_delete);
 					   this.updDbUsesPreparedStatement(sqlCmd_delete.toString(),paramList) ;
        	      }//end of countdata    	   
    	       }//end of query data		    	       
    	       if(Star[i].equals("FGN")){//96.01.10 fix 農漁會支票存款資料.加上刪除wml01
    	    	   paramList.clear() ;
    	           paramList.add(Integer.toString(Integer.parseInt(m_year)-1911)) ;
    	           paramList.add(m_month) ;
    	           dbData = DBManager.QueryDB_SQLParam("select count(*) as countdata from wml01 where m_year=? and m_month=? and report_no='F01'",paramList,"countdata");
    	           paramList.clear() ;
    	           if(dbData != null && dbData.size() ==1){
 				      if(!((((DataObject)dbData.get(0)).getValue("countdata")).toString()).equals("0")){
 					    System.out.println(" wml01 have "+(((DataObject)dbData.get(0)).getValue("countdata")).toString() +" datas");
 					    paramList.add(Integer.toString(Integer.parseInt(m_year)-1911)) ;
 					    paramList.add(m_month) ;
 					    this.updDbUsesPreparedStatement(" delete wml01 where m_year=? and m_month=? and report_no='F01'",paramList) ;
 					    //updateDBSqlList.add(" delete wml01 where m_year="+Integer.toString(Integer.parseInt(m_year)-1911)+" and m_month="+m_month+" and report_no='F01'");
        	          }//end of countdata    	   
    	           }//end of query data		
 				   
 			   }
 			   /*if(updateDBSqlList.size() != 0){
 			      for(int delidx=0;delidx<updateDBSqlList.size();delidx++){
 			          System.out.println(updateDBSqlList.get(delidx)); 			          
 			      }
 			      if(!DBManager.updateDB(updateDBSqlList)){
    			     alertMsg += "舊有資料刪除失敗";
    			  } 
 			   }*/
    	       int result = -1;
    	       
    	       if(alertMsg.equals("")){
    	          result = AgriBankFTP.getDataFiles(FTP_SERVER, FTP_USER,FTP_PASSWORD, FTP_DIRECTORY, LOCAL_DIRECTORY, TargetFile,getFiles);
    	       }   
        	   System.out.println("result="+result);
        	   if(result==1){
        	      alertMsg += "轉檔已經完成,請至資料庫查詢!!";
        	   }else if(result==0){
        	      alertMsg += "轉檔不成功,已發錯誤信件至農業金庫管理者,請聯絡金庫人員協助解決!!";
        	   } 	
        	   webURL_Y = "/pages/ZZ092W_List.jsp";
			   request.setAttribute("alertMsg",alertMsg);
    		   request.setAttribute("webURL_Y",webURL_Y); 	
			   rd = application.getRequestDispatcher( nextPgName );	
    	   }//end of for
        }//end of download
     }
%>

<%
	try {
        //forward to next present jsp
        rd.forward(request, response);
    } catch (NullPointerException npe) {
    }
    }//end of doProcess
%>

<%!
    private final static String nextPgName = "/pages/ActMsg.jsp";        
    private final static String ListPgName = "/pages/ZZ092W_List.jsp";        
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    private boolean CheckPermission(HttpServletRequest request){//檢核權限    	    
    	    boolean CheckOK=false;
    	    HttpSession session = request.getSession();            
            Properties permission = ( session.getAttribute("ZZ092W")==null ) ? new Properties() : (Properties)session.getAttribute("ZZ092W");				                
            if(permission == null){
              System.out.println("ZZ092W.permission == null");
            }else{
               System.out.println("ZZ092W.permission.size ="+permission.size());
               
            }
            //只要有Query的權限,就可以進入畫面
        	if(permission != null && permission.get("Q") != null && permission.get("Q").equals("Y")){            
        	   CheckOK = true;//Query
        	}
        	return CheckOK;
   }
    private boolean updDbUsesPreparedStatement(String sql ,List paramList) throws Exception{
		List updateDBList = new ArrayList();//0:sql 1:data
	    List updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
		List updateDBDataList = new ArrayList();//儲存參數的List
		
		updateDBDataList.add(paramList);
		updateDBSqlList.add(sql);
		updateDBSqlList.add(updateDBDataList);
		updateDBList.add(updateDBSqlList);
		return DBManager.updateDB_ps(updateDBList) ;
	}
%>    