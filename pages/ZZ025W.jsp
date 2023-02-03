<%
//93.12.20 add 權限檢核 by 2295
//         add 設定異動者資訊 by 2295
//93.12.23 add 超過登入時間,請重新登入 by 2295
//94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除 by 2295
//94.01.06 fix 所有人都有使用者資料維護的權限 by 2295
//94.01.10 fix 若Muser_Data裡的e-mail為空白時,須強制輸入基本資料,並登出.. by 2295
//         fix 將更改後的使用者e-mail.寫到session by 2295
//99.12.10 fix sqlInjection by 2808
//102.06.27 add 操作歷程寫入log by2968
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
<%
	RequestDispatcher rd = null;
	String actMsg = "";	
	String alertMsg = "";	
	String webURL = "";	
	boolean doProcess = false;	
	
	//取得session資料,取得成功時,才繼續往下執行===================================================
	if(session.getAttribute("muser_id") == null){//session timeout	
      System.out.println("ZZ025W login timeout");   
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
	String muser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");				    
	String tbank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");				    
	String lguser_id = (session.getAttribute("muser_id") == null)?"":(String)session.getAttribute("muser_id");
	String lguser_name = (session.getAttribute("muser_name") == null)?"":(String)session.getAttribute("muser_name"); 
	String lguser_type = (session.getAttribute("muser_type") == null)?"":(String)session.getAttribute("muser_type");	
	session.setAttribute("nowtbank_no",null);//94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除======
	System.out.println("act="+act);			
    
    if(session.getAttribute("muser_id") == null){	
      System.out.println("ZZ025W login timeout");   
	  rd = application.getRequestDispatcher( "/pages/reLogin.jsp?url=LoginError.jsp?timeout=true" );         	   
      rd.forward(request,response);
   }		
    if(!CheckPermission(request)){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );        
    }else{            
    	//set next jsp 	
    	if(act.equals("Edit")){        		
            List WTT01 = getWTT01(muser_id);    	      
            List MUSER_DATA = getMUSER_DATA(muser_id);    	      
            request.setAttribute("WTT01",WTT01);    	                               
            request.setAttribute("MUSER_DATA",MUSER_DATA);    	                               
    	    List bank_no_list = getBank_No(tbank_no);//組室代碼    	    
    	    request.setAttribute("bank_no",bank_no_list); 
    	    actMsg = this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,"Q");    	    
    	    if("Y".equals(actMsg)){
        	   actMsg = "相關資料寫入成功";
        	}
    	 	rd = application.getRequestDispatcher( EditPgName +"?act=Edit&test=nothing");            	
        }else if(act.equals("Update")){
    	    actMsg = UpdateDB(request,muser_id,lguser_id,lguser_name);
    	    if("Y".equals(actMsg)){
    	        actMsg = this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,"U");
    	        if("Y".equals(actMsg)){
        	       actMsg = "相關資料寫入成功";
        	    }
    	    }
        	rd = application.getRequestDispatcher( nextPgName );         
        }
    	request.setAttribute("actMsg",actMsg);    
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
	private final static String program_id = "ZZ025W";
    private final static String nextPgName = "/pages/ActMsg.jsp";    
    private final static String EditPgName = "/pages/"+program_id+"_Edit.jsp";    
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    private boolean CheckPermission(HttpServletRequest request){//檢核權限    	    
    		return true;//所有人都有使用者資料維護的權限
            /* 
    	    boolean CheckOK=false;
    	    HttpSession session = request.getSession();            
            Properties permission = ( session.getAttribute("ZZ025W")==null ) ? new Properties() : (Properties)session.getAttribute("ZZ025W");				                
            if(permission == null){
              System.out.println("ZZ025W.permission == null");
            }else{
               System.out.println("ZZ025W.permission.size ="+permission.size());
               
            }
            //只要有Query的權限,就可以進入畫面
        	if(permission != null && permission.get("Q") != null && permission.get("Q").equals("Y")){            
        	   CheckOK = true;//Query
        	}
        	return CheckOK;
        	*/
    }            
   
    //取得各組室代碼
    private List getBank_No(String tbank_no){
    		//查詢條件    
    		List paramList =new ArrayList() ;
    		String sqlCmd = "select bank_no,bank_name from bn02 "
						  + " where tbank_no=? "
    		 			  + " order by bank_no";
    		paramList.add(tbank_no) ;
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    } 
    
    //取得WTT01該使用者帳號資料
    private List getWTT01(String muser_id){
    		//查詢條件    
    		List paramList =new ArrayList() ;
    		String sqlCmd = " select * from WTT01 "
    					  + " where muser_id=? ";  
    		paramList.add(muser_id) ;
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    } 
    //取得MUSER_DATA該使用者基本資料
    private List getMUSER_DATA(String muser_id){
    		//查詢條件    
    		List paramList = new ArrayList() ;
    		String sqlCmd = " select * from MUSER_DATA "
    					  + " where muser_id=? ";  
    		paramList.add(muser_id) ;
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    }
    
        
	public String UpdateDB(HttpServletRequest request,String muser_id,String lguser_id,String lguser_name) throws Exception{    		
		String sqlCmd = "";		
		String errMsg="";		
		String muser_name=((String)request.getParameter("MUSER_NAME")==null)?"":(String)request.getParameter("MUSER_NAME");		
		String m_position=((String)request.getParameter("M_POSITION")==null)?"":(String)request.getParameter("M_POSITION");		
		String m_telno=((String)request.getParameter("M_TELNO")==null)?"":(String)request.getParameter("M_TELNO");
		String m_cellino=((String)request.getParameter("M_CELLINO")==null)?"":(String)request.getParameter("M_CELLINO");
		String m_fax=((String)request.getParameter("M_FAX")==null)?"":(String)request.getParameter("M_FAX");		
		String m_email=((String)request.getParameter("M_EMAIL")==null)?"":(String)request.getParameter("M_EMAIL");
		String m_sex=((String)request.getParameter("M_SEX")==null)?"":(String)request.getParameter("M_SEX");		
		String m_note=((String)request.getParameter("M_NOTE")==null)?"":(String)request.getParameter("M_NOTE");
		
		String director_name=((String)request.getParameter("DIRECTOR_NAME")==null)?"":(String)request.getParameter("DIRECTOR_NAME");		
		String director_telno=((String)request.getParameter("DIRECTOR_TELNO")==null)?"":(String)request.getParameter("DIRECTOR_TELNO");
		String director_cellino=((String)request.getParameter("DIRECTOR_CELLINO")==null)?"":(String)request.getParameter("DIRECTOR_CELLINO");
		String director_fax=((String)request.getParameter("DIRECTOR_FAX")==null)?"":(String)request.getParameter("DIRECTOR_FAX");		
		String director_email=((String)request.getParameter("DIRECTOR_EMAIL")==null)?"":(String)request.getParameter("DIRECTOR_EMAIL");
		String director_sex=((String)request.getParameter("DIRECTOR_SEX")==null)?"":(String)request.getParameter("DIRECTOR_SEX");				
		
		String bank_no=((String)request.getParameter("BANK_NO")==null)?" ":(String)request.getParameter("BANK_NO");				
		bank_no = bank_no.equals("")?" ":bank_no;
		String old_bank_no=((String)request.getParameter("OLD_BANK_NO")==null)?" ":(String)request.getParameter("OLD_BANK_NO");				
		old_bank_no = old_bank_no.equals("")?" ":old_bank_no;
		String subdep_id=((String)request.getParameter("SUBDEP_ID")==null)?"":(String)request.getParameter("SUBDEP_ID");				
		String old_subdep_id=((String)request.getParameter("OLD_SUBDEP_ID")==null)?"":(String)request.getParameter("OLD_SUBDEP_ID");				
		
		String user_id=lguser_id;
		String user_name=lguser_name;
		HttpSession session = request.getSession();
		try {
				List paramList =new ArrayList() ;
				sqlCmd = "SELECT * FROM MUSER_DATA WHERE muser_id=? ";			
				paramList.add(muser_id) ;
			    List data = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");		
			    paramList.clear() ;
				System.out.println("MUSER_DATA.size="+data.size());
				//組室代碼或科別有更改時,寫入WTT01/WTT01_LOG
				if((!bank_no.equals(old_bank_no)) || (!subdep_id.equals(old_subdep_id)) || (!muser_name.equals(user_name)) ){
				    //insert WTT01_LOG===================================================		    
				    sqlCmd = " INSERT INTO WTT01_LOG " 
						   + " select muser_id,muser_name,muser_password,muser_i_o,bank_type,"
						   + " tbank_no,bank_no,subdep_id,add_user,add_name,add_date,firstlogin_mark,login_mark,"
						   + " lock_mark,delete_mark,muser_type,password_update_date,password_pre,"
						   + " user_id,user_name,update_date"
						   + ",?,?,sysdate,'U'"
						   + " from WTT01"						  
						   + " WHERE muser_id=?"; 		
				    paramList.add(user_id) ;
				    paramList.add(user_name) ;
				    paramList.add(muser_id) ;
					this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
					//=========================================================================
					paramList.clear(); 
				    sqlCmd = "UPDATE WTT01 SET bank_no=?,subdep_id=?"
				           + ",muser_name=?"
				    	   + ",user_id=?"					       
					       + ",user_name=?"
					       + ",update_date=sysdate" 		            		 	
				    	   + " where muser_id=?";
				    paramList.add(bank_no) ;
				    paramList.add(subdep_id) ;
				    paramList.add(muser_name) ;
				    paramList.add(user_id) ;
				    paramList.add(user_name) ;
				    paramList.add(muser_id) ;
				    
				    this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
				}
				paramList.clear() ;
				if (data.size() == 0){//無資料時,Insert
				    sqlCmd = "INSERT INTO MUSER_DATA VALUES ('" + muser_id + "'"
					       + ",'" + m_position + "'" 					       
					       + ",'" + m_telno +"'"
					       + ",'" + m_cellino + "'"
					       + ",'" + m_fax + "'"
					       + ",'" + m_email + "'"					       
					       + ",'" + m_sex + "'"
					       + ",'" + m_note + "'" 
					       + ",'" + director_name + "'" 					       
					       + ",'" + director_telno +"'"
					       + ",'" + director_cellino + "'"
					       + ",'" + director_fax + "'"
					       + ",'" + director_email + "'"					       
					       + ",'" + director_sex + "'"
					       + ",'" + user_id +"'"					       
					       + ",'" + user_name + "'"
					       + ",sysdate)"; 		 
				    sqlCmd = " INSERT INTO MUSER_DATA VALUES ( "  ;
				    sqlCmd += "?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,sysdate ) " ;
					paramList.add(muser_id) ;
					paramList.add(m_position) ; 
					paramList.add(m_telno) ;
					paramList.add(m_cellino);
					paramList.add(m_fax) ;
					paramList.add(m_email) ;
					paramList.add(m_sex) ;
					paramList.add(m_note) ;
					paramList.add(director_name) ;
					paramList.add(director_telno) ;
					paramList.add(director_cellino) ;
					paramList.add(director_fax) ;
					paramList.add(director_email) ;
					paramList.add(director_sex) ;
					paramList.add(user_id) ;
					paramList.add(user_name) ;
				}else{//有資料時,Update    				 
				    sqlCmd = " INSERT INTO MUSER_DATA_LOG " 
						   + " select muser_id,m_position,m_telno,m_cellino,m_fax,m_email,"
						   + " m_sex,m_note,director_name,director_telno,director_cellino,"
						   + " director_fax,director_email,director_sex,user_id,user_name,update_date,"
						   + " ?,?,sysdate,'U'"
						   + " from MUSER_DATA"
						   + " WHERE muser_id=? ";	
				 	paramList.add(user_id) ;
				    paramList.add(user_name) ;
				    paramList.add(muser_id) ;
				    this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
				    paramList.clear() ;
				    sqlCmd = "UPDATE MUSER_DATA SET "
				    	   + " m_position=?" 					       
					       + ",m_telno=?"
					       + ",m_cellino=?"
					       + ",m_fax=?"
					       + ",m_email=?"					       
					       + ",m_sex=?"
					       + ",m_note=?" 
					       + ",director_name=?" 
					       + ",director_telno=?"
					       + ",director_cellino=?"
					       + ",director_fax=?"
					       + ",director_email=?"					       
					       + ",director_sex=?"					      			       
					       + ",user_id=?"					       
					       + ",user_name=?"
					       + ",update_date=sysdate" 		            		 	
				    	   + " where muser_id=?";		
				    paramList.add(m_position) ;
				    paramList.add(m_telno) ;
				    paramList.add(m_cellino) ;
				    paramList.add(m_fax) ;
				    paramList.add(m_email) ;
				    paramList.add(m_sex) ;
				    paramList.add(m_note) ;
				    paramList.add(director_name) ;
				    paramList.add(director_telno) ;
				    paramList.add(director_cellino) ;
				    paramList.add(director_fax) ;
				    paramList.add(director_email) ;
				    paramList.add(director_sex) ;
				    paramList.add(user_id) ;
				    paramList.add(user_name) ;
				    paramList.add(muser_id) ;
    	   		}
				if(this.updDbUsesPreparedStatement(sqlCmd,paramList)){
					errMsg = errMsg + "Y";	
					//94.01.10 fix 若Muser_Data裡的e-mail為空白時,須強制輸入基本資料,並登出..
					if(session.getAttribute("UpdateMuser_Data") != null && ((String)session.getAttribute("UpdateMuser_Data")).equals("true")){
					   errMsg = "更改使用者基本資料成功,請重新登入系統...<br>";
					}	
					//94.01.10 fix 將更改後的使用者e-mail.寫到session
					if(!m_email.equals((String)session.getAttribute("muser_email"))){
					    session.setAttribute("muser_email",m_email);//使用者e-mail		
					}
					if(!muser_name.equals(user_name)){
					    session.setAttribute("muser_name",muser_name);//使用者姓名		
					}    					
					if(!bank_no.equals(old_bank_no)){
					    session.setAttribute("bank_no",bank_no);//分支機構代號
					}
				}else{
				 	errMsg = errMsg + "相關資料寫入資料庫失敗 ";
				}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗<br>[Exception Error]:<br>"+e.getMessage();					
		}	

		return errMsg;
	} 	
	public String InsertWlXOPERATE_LOG(HttpServletRequest request,String lguser_id,String program_id,String update_type) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList() ;
		String errMsg="";
	    try {
	        sqlCmd.append(" INSERT INTO WlXOPERATE_LOG(muser_id,use_Date,program_id,ip_address,update_type)");
	        sqlCmd.append("                     VALUES(?,sysdate,?,?,?) ");
	        paramList.add(lguser_id);
	        paramList.add(program_id);
	        paramList.add(request.getRemoteAddr());//ipAddress
	        paramList.add(update_type);//操作類別 I-新增，U-異動，D-刪除，Q-明細，P-列印
	        if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){
				errMsg = errMsg + "Y";					
			}else{
			    errMsg = errMsg + "相關資料寫入log失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
			}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入log失敗";						
		}	

		return errMsg;
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