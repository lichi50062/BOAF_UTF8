<%
//  95.09.29 create  2495	
//  95.10.02	add boaf報表下載 by 2495
//  95.10.17 增加 insert A04_LOG BY 2495 	
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.*,java.io.*" %>
<%@ page import="org.apache.poi.poifs.filesystem.*,org.apache.poi.hssf.usermodel.*" %>
<%@ page import="com.tradevan.util.Utility" %>	
<%@ page import="com.tradevan.util.report.RptFR007WA_BOAF" %>								          
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Enumeration" %>
<%@ page import="com.tradevan.util.DBManager" %>	

<%
   response.setContentType("application/msexcel;charset=UTF-8");//以上這行設定本網頁為excel格式的網頁	     
   String act = ( request.getParameter("act")==null ) ? "view" : (String)request.getParameter("act");			  
   String Unit = ( request.getParameter("Unit")==null ) ? "1" : (String)request.getParameter("Unit");			  
   String bank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");			    
   String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? "" : (String)request.getParameter("S_YEAR");
   String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? "" : (String)request.getParameter("S_MONTH");
   String datestate = ( request.getParameter("datestate")==null ) ? "" : (String)request.getParameter("datestate");   
   String filename="農漁會信用部資產品質分析明細表.xls";
   String BANK_NO = ( request.getParameter("BANK_NO")==null ) ? "" : (String)request.getParameter("BANK_NO");
   String BANK_NAME = ( request.getParameter("BANK_NAME")==null ) ? "" : (String)request.getParameter("BANK_NAME");
   String muser_id = ( request.getParameter("muser_id")==null ) ? "" : (String)request.getParameter("muser_id");
   String muser_name = ( request.getParameter("muser_name")==null ) ? "" : (String)request.getParameter("muser_name");  
   
   System.out.println("FR007WA Start......................"); 
   System.out.println("S_YEAR="+S_YEAR);
   System.out.println("S_MONTH="+S_MONTH);
   System.out.println("act="+act);
   System.out.println("Unit="+Unit);
   System.out.println("datestate="+datestate);
   System.out.println("bank_type="+bank_type);
   System.out.println("BANK_NO="+BANK_NO);
   System.out.println("muser_id="+muser_id);
   System.out.println("muser_name="+muser_name);
   RequestDispatcher rd = null;
   boolean doProcess = false;	
   String actMsg = "";		
   //取得session資料,取得成功時,才繼續往下執行===================================================
   if(session.getAttribute("muser_id") == null){//session timeout		
      System.out.println("FR007WA login timeout");   
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
	    	//set next jsp 	
	    
	    	if(act.equals("Qry")){
	    	   rd = application.getRequestDispatcher( QryPgName + "?bank_type="+bank_type);        
	    	}else if(act.equals("view")){
      		   //以上這行設定傳送到前端瀏覽器時的檔名為test1.xls
      		   //就是靠這一行，讓前端瀏覽器以為接收到一個excel檔 
      			response.setHeader("Content-disposition","inline; filename=view.xls");
   			}else if (act.equals("download")){   
      	*/		
      	response.setHeader("Content-Disposition","attachment; filename=download.xls");
   			/*}*/   
   			
   			if(/*act.equals("view") || act.equals("download")*/true){
   				try{		    
	    			actMsg = RptFR007WA_BOAF.createRpt(S_YEAR,S_MONTH,Unit,datestate,bank_type,BANK_NO,BANK_NAME);	    
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
		    //95.10.17 增加 insert A04_LOG BY 2495 
        String acc_code="ALL",atm="0",user_id_c="",user_name_c="",update_type_c="L";
        INSERT_A04_LOG(S_YEAR,S_MONTH,BANK_NO,acc_code,atm,muser_id,muser_name,update_type_c); 
					fin.close();
					out1.close();            		      
		
				}catch(Exception e){
	   				System.out.println(e.getMessage());
				}   			
   			}
   		/*}*/
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
    private final static String QryPgName = "/pages/FR007WA_Qry.jsp";        
    private final static String RptCreatePgName = "/pages/FR007WA_Excel.jsp";        
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    private boolean CheckPermission(HttpServletRequest request){//檢核權限    	        		
    	    boolean CheckOK=false;
    	    HttpSession session = request.getSession();            
            Properties permission = ( session.getAttribute("FR007WA")==null ) ? new Properties() : (Properties)session.getAttribute("FR007WA");				                
            if(permission == null){
              System.out.println("FR007WA.permission == null");
            }else{
               System.out.println("FR007WA.permission.size ="+permission.size());
               
            }
            //只要有Query的權限,就可以進入畫面
        	if(permission != null && permission.get("Q") != null && permission.get("Q").equals("Y")){            
        	   CheckOK = true;//Query
        	}
        	System.out.println("CheckOk="+CheckOK);        	
        	return CheckOK;
    }   
%>
<%!
	  //95.10.17 ADD insert A04_LOG BY 2495 
	  private String INSERT_A04_LOG(String m_year,String m_month,String bank_code,String acc_code,String atm,String user_id_c,String user_name_c,String update_type_c) throws Exception{    	
		String sqlCmd = "";		
		String errMsg="";		
		
		try {
				List updateDBSqlList = new LinkedList();								   				   
				//insert A04_LOG===================================================		    
				sqlCmd = "INSERT INTO A04_LOG VALUES ("+m_year
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
						errMsg = errMsg + "無法寫入A04_log資料";					
				}else{
				  	errMsg = errMsg + "無法寫入A04_log資料<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
				}    	   		
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "無法寫入A04_log資料<br>[Exception Error]";								
		}	

		return errMsg;
	}  			
%>