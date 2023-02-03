<%
// 94.11.07 first designed by 4180
// 99.04.02 fix 當為第1季時.前一季為第4季 by 2295
// 99.12.12 fix sqlInjection by 2808  
//100.01.27 fix bug by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
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
      System.out.println("FX009W login timeout");
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

    //若程序為新增資料，則從request中抓出年月以判斷是否已有資料
    String checkyear = ( request.getParameter("checkyear")==null ) ? "" : (String)request.getParameter("checkyear");
	String checkquarter = ( request.getParameter("checkquarter")==null ) ? "" : (String)request.getParameter("checkquarter");

    //若程序為修改資料，則從request中抓出年月
    String myear = ( request.getParameter("myear")==null ) ? "" : (String)request.getParameter("myear");
	String mquarter = ( request.getParameter("mquarter")==null ) ? "" : (String)request.getParameter("mquarter");

	//登入者資訊
	String lguser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");
	String lguser_name = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");
	String lguser_type = ( session.getAttribute("muser_type")==null ) ? "" : (String)session.getAttribute("muser_type");
	//======================================================================================================================
	System.out.println("act="+act);
	//fix 93.12.20 若有已點選的tbank_no,則以已點選的tbank_no為主============================================================
	String bank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");
	String nowtbank_no =  ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session
	}
	bank_no = ( session.getAttribute("nowtbank_no")==null ) ? bank_no : (String)session.getAttribute("nowtbank_no");
	//=======================================================================================================================

    if(!Utility.CheckPermission(request,"FX009W")){//無權限時,導向到LoginError.jsp   
        rd = application.getRequestDispatcher( LoginErrorPgName );
    }else{
    	//set next jsp
    	if(act.equals("new")){
    	   List dbData = getWLX09_S_WARNING(bank_no,checkyear,checkquarter);  //是否當年季已有資料
           if(dbData.size()==0){
        	  int temp = Integer.parseInt(checkquarter); //將參數轉換為上季
        	  temp = temp-1;
        	  String lastquarter = String.valueOf(temp);
        	  if(temp==0){
        	  	 lastquarter="4";
        	  	 checkyear=String.valueOf(Integer.parseInt(checkyear)-1);
        	  }
        	  dbData = getWLX09_S_WARNING(bank_no,checkyear,lastquarter);  //取得上個月資料
        	  if(dbData.size()==0){ //上月若無資料
        	 	 rd = application.getRequestDispatcher( EditPgName +"?act=new");
        	  }else{
        	 	 rd = application.getRequestDispatcher( EditPgName +"?act=new&loaddata=ok");
        	  }	 
            }else{
        	  dbData = getWLX09_S_WARNING(bank_no,checkyear,checkquarter);
    	      request.setAttribute("WLX09_S_WARNING",dbData);
			  request.setAttribute("maintainInfo","select * from WLX09_S_WARNING WHERE bank_no='" + bank_no+"'");
			  rd = application.getRequestDispatcher( EditPgName +"?act=modify");
        	}
    	}else if(act.equals("Edit")){
    	    List dbData = getWLX09_S_WARNING(bank_no,myear,mquarter);
    	    request.setAttribute("WLX09_S_WARNING",dbData);
    	    //93.12.21設定異動者資訊======================================================================
			request.setAttribute("maintainInfo","select * from WLX09_S_WARNING WHERE bank_no='" + bank_no+"'");
			//=======================================================================================================================
        	rd = application.getRequestDispatcher( EditPgName +"?act=Edit");
    	}else if(act.equals("List")){
    	    List dbData = getWLX09_S_WARNING(bank_no,"","");
            List inidate = getWLX09_S_WARNING_INI();
            List lockdate = getWLX09_S_WARNING_LOCK(bank_no);
    	    request.setAttribute("WLX09_S_WARNING",dbData);
            request.setAttribute("WLX09_INI",inidate);
            request.setAttribute("WLX09_LOCK",lockdate);
        	rd = application.getRequestDispatcher( ListPgName +"?bank_no="+bank_no);
      }else if(act.equals("Load")){
      	    String syear = ((String)request.getParameter("hyear")==null)?"":(String)request.getParameter("hyear");
	        String squarter= ((String)request.getParameter("hquarter")==null)?"":(String)request.getParameter("hquarter");
        	int temp = Integer.parseInt(squarter); //將參數轉換為上季
            temp = temp-1;
            String lyear=syear;
        	String lquarter = String.valueOf(temp);
        	if(temp==0){
        	   lquarter="4";//99.04.02 fix 當為第1季時.前一季為第4季
        	   lyear=String.valueOf(Integer.parseInt(syear)-1);
        	}//處理第一季的情況

        	List dbData = getWLX09_S_WARNING(bank_no,lyear,lquarter);  //取得上個月資料
        	request.setAttribute("WLX09_S_WARNING",dbData);
        	rd = application.getRequestDispatcher( EditPgName +"?act=new&loaddata=loaded&S_YEAR="+syear+"&S_QUARTER="+squarter);
    	}else if(act.equals("Insert")){
    	    actMsg = InsertDB(request,bank_no,lguser_id,lguser_name);
        	rd = application.getRequestDispatcher( nextPgName+"?FX=FX009W" );
    	}else if(act.equals("Update")){
    	    actMsg = UpdateDB(request,bank_no,lguser_id,lguser_name);
        	rd = application.getRequestDispatcher( nextPgName+"?FX=FX009W");
    	}else if(act.equals("Delete")){
    	    actMsg = DeleteDB(request,bank_no,lguser_id,lguser_name);
        	rd = application.getRequestDispatcher( nextPgName+"?FX=FX009W" );
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
    private final static String nextPgName = "/pages/ActMsg.jsp";
    private final static String EditPgName = "/pages/FX009W_Edit.jsp";
    private final static String ListPgName = "/pages/FX009W_List.jsp";
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    
    //94.04.01 add 主key更改為bank_no+seq_no by 2295
    private List getWLX09_S_WARNING(String bank_no, String myear, String mquarter){
    		//程序為顯示畫面，查詢條件
    		List paramList = new ArrayList() ;
    		String sqlCmd = "select * from WLX09_S_WARNING where bank_no= ? ";
    		paramList.add(bank_no) ;
    		if(!myear.equals("") && !mquarter.equals("")){
    			sqlCmd =sqlCmd + " and m_year=? and m_quarter= ?";
    			paramList.add(myear) ;
    			paramList.add(mquarter) ;
    		}
    		sqlCmd =sqlCmd + " order by M_YEAR desc, M_QUARTER desc";
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"m_year,m_quarter,warnaccount_tcnt,warnaccount_tbal,"+
            "warnaccount_remit_tcnt,warnaccount_refund_apply_cnt,warnaccount_refund_apply_amt,warnaccount_refund_cnt,warnaccount_refund_amt,user_id,user_name,update_date");
            return dbData;
    }
    private List getWLX09_S_WARNING_INI(){
        List paramList = new ArrayList();
      	String sqlCmd = "select * from WLX_APPLY_INI where REPORT_NO=?";
      	paramList.add("C06");
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"m_year,m_month");
        if(dbData.size()!= 0){
           return dbData;
        }else{
          return null;
        }  
    }
    private List getWLX09_S_WARNING_LOCK(String bank_no){
    	List paramList =new ArrayList() ;
      	String sqlCmd = " select M_YEAR,M_QUARTER from WLX_APPLY_LOCK "
 		   			  + " where BANK_CODE = ? "
 		       		  + " and REPORT_NO = 'C06'"          
 		       		  + " and((LOCK_OWN   =  'Y') or (LOCK_MGR = 'Y'))";
      	paramList.add(bank_no) ;
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"m_year,m_quarter");
        if(dbData.size()!= 0){
           return dbData;
        }else{
          return null;
        }  
    }
    public String InsertDB(HttpServletRequest request,String bank_no,String lguser_id,String lguser_name) throws Exception{
  			StringBuffer sqlCmd = new StringBuffer();
			String errMsg="";
	
			String nyear = ((String)request.getParameter("hyear")==null)?"":(String)request.getParameter("hyear");
			String nquarter= ((String)request.getParameter("hquarter")==null)?"":(String)request.getParameter("hquarter"); 	
			String warnaccount_tcnt=((String)request.getParameter("warnaccount_tcnt")==null)?"":(String)request.getParameter("warnaccount_tcnt");
			String warnaccount_tbal=((String)request.getParameter("warnaccount_tbal")==null)?"":(String)request.getParameter("warnaccount_tbal");
			String warnaccount_remit_tcnt=((String)request.getParameter("warnaccount_remit_tcnt")==null)?"":(String)request.getParameter("warnaccount_remit_tcnt");
			String warnaccount_refund_apply_cnt=((String)request.getParameter("warnaccount_refund_apply_cnt")==null)?"":(String)request.getParameter("warnaccount_refund_apply_cnt");
			String warnaccount_refund_apply_amt=((String)request.getParameter("warnaccount_refund_apply_amt")==null)?"":(String)request.getParameter("warnaccount_refund_apply_amt");
			String warnaccount_refund_cnt=((String)request.getParameter("warnaccount_refund_cnt")==null)?"":(String)request.getParameter("warnaccount_refund_cnt");
			String warnaccount_refund_amt=((String)request.getParameter("warnaccount_refund_amt")==null)?"":(String)request.getParameter("warnaccount_refund_amt");
			String user_id=lguser_id;
			String user_name=lguser_name;
			String sqlLock = "";
			String sqlLockLog = "";
			List dbLock;			
			System.out.print("WARN==="+warnaccount_refund_amt);
		try {
   			//List updateDBSqlList = new LinkedList();
   			List paramList = new ArrayList() ;
			sqlCmd.append("INSERT INTO WLX09_S_WARNING VALUES(?,?,?,?,?,?,?,?,?,?,?,?,sysdate)" );
			 paramList.add(nyear) ;
			 paramList.add(nquarter) ;
			 paramList.add(bank_no) ;
			 paramList.add(warnaccount_tcnt);
			 paramList.add(warnaccount_tbal) ;
			 paramList.add(warnaccount_remit_tcnt) ;
			 paramList.add(warnaccount_refund_apply_cnt) ;
			 paramList.add(warnaccount_refund_apply_amt) ;
			 paramList.add(warnaccount_refund_cnt) ;
			 paramList.add(warnaccount_refund_amt) ;
			 paramList.add(lguser_id) ;
			 paramList.add(lguser_name) ;
			 this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList) ;
            //updateDBSqlList.add(sqlCmd);
                
            //當寫進資料庫成功後，再寫進WLX_APPLY_LOCK檔中
            sqlCmd.setLength(0) ;
            paramList.clear() ;
			sqlLock = " select * from WLX_APPLY_LOCK aa where aa.M_YEAR =?"+
		    	      " and aa.M_QUARTER=?"+
		        	  " and aa.BANK_CODE=?"+
		          	  " and aa.REPORT_NO='C06'";
			paramList.add(nyear) ;
			paramList.add(nquarter) ;
			paramList.add(bank_no) ;
	  		dbLock = DBManager.QueryDB_SQLParam(sqlLock,paramList,"");   	  
	  		System.out.print("大小:"+String.valueOf(dbLock.size()));
	  		paramList.clear() ;
			if(dbLock.size()!=0){     
				sqlLock = " update WLX_APPLY_LOCK set USER_ID = ?"+
						  " ,USER_NAME =?"+
						  " ,UPDATE_DATE =sysdate"+
					  	  " where M_YEAR =? and M_QUARTER =? and BANK_CODE =? "+
					  	  " and REPORT_NO ='C06'";
				paramList.add(user_id) ;
				paramList.add(user_name) ;
				paramList.add(nyear) ;
				paramList.add(nquarter) ;
				paramList.add(bank_no) ;
				//updateDBSqlList.add(sqlLock);  
			}else{  //若是原本沒有鎖定資料的情況下，寫進一筆新的鎖定紀錄
				/*sqlLock = " insert into WLX_APPLY_LOCK values('"
		        	    +nyear+"','"
		            	+nquarter+"','"
		                +bank_no+"','C06','"
		                +user_id+"','"
		                +user_name+"',sysdate,'N','N','"
		                +user_id+"','"
		                +user_name+"',sysdate)"; 					  
				//updateDBSqlList.add(sqlLock); */
				sqlLock = "Insert into WLX_APPLY_LOCK values ( " ;
				sqlLock += "  ?,?,?,'C06',?,?,sysdate,'N','N',?,?,sysdate) " ;
				paramList.add(nyear) ;
				paramList.add(nquarter) ;
				paramList.add(bank_no) ;
				paramList.add(user_id) ;
				paramList.add(user_name) ;
				paramList.add(user_id) ;
				paramList.add(user_name) ;
			}    
                
			//if(DBManager.updateDB(updateDBSqlList)){
			if(this.updDbUsesPreparedStatement(sqlLock,paramList)) {
				errMsg = errMsg + "相關資料寫入資料庫成功";
			}else{
				errMsg = errMsg + "相關資料寫入資料庫失敗";
			}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}
		return errMsg;
	}
	//94.04.01 add 主key更改為bank_no+seq_no by 2295
	public String UpdateDB(HttpServletRequest request,String bank_no,String lguser_id,String lguser_name) throws Exception{
			StringBuffer sqlCmd = new StringBuffer( );
			String sqlCmdLog = "";
			String errMsg="";
			String nyear = ((String)request.getParameter("hyear")==null)?"":(String)request.getParameter("hyear");
			String nquarter= ((String)request.getParameter("hquarter")==null)?"":(String)request.getParameter("hquarter"); 	
			String warnaccount_tcnt=((String)request.getParameter("warnaccount_tcnt")==null)?"":(String)request.getParameter("warnaccount_tcnt");
			String warnaccount_tbal=((String)request.getParameter("warnaccount_tbal")==null)?"":(String)request.getParameter("warnaccount_tbal");
			String warnaccount_remit_tcnt=((String)request.getParameter("warnaccount_remit_tcnt")==null)?"":(String)request.getParameter("warnaccount_remit_tcnt");
			String warnaccount_refund_apply_cnt=((String)request.getParameter("warnaccount_refund_apply_cnt")==null)?"":(String)request.getParameter("warnaccount_refund_apply_cnt");
			String warnaccount_refund_apply_amt=((String)request.getParameter("warnaccount_refund_apply_amt")==null)?"":(String)request.getParameter("warnaccount_refund_apply_amt");
			String warnaccount_refund_cnt=((String)request.getParameter("warnaccount_refund_cnt")==null)?"":(String)request.getParameter("warnaccount_refund_cnt");
			String warnaccount_refund_amt=((String)request.getParameter("warnaccount_refund_amt")==null)?"":(String)request.getParameter("warnaccount_refund_amt");
			String user_id=lguser_id;
			String user_name=lguser_name;
			String sqlLock = "";
			String sqlLockLog = "";
			List dbLock;


		try {
			   //List updateDBSqlList = new LinkedList();
			   List paramList =new ArrayList() ;
			   sqlCmd.append("SELECT * FROM WLX09_S_WARNING WHERE bank_no=?  AND m_year=?  and m_quarter= ? ");
			   paramList.add(bank_no) ;
			   paramList.add(nyear) ;
			   paramList.add(nquarter) ;
			   List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
			   paramList.clear() ;
			   sqlCmd.setLength(0);
			   System.out.println("WLX09_S_WARNING.size="+data.size());
			   if (data.size() == 0){
				   errMsg = errMsg + "此筆資料不存在無法修改<br>";
			   }else{
				/*寫入log檔中*/
				sqlCmdLog = "INSERT INTO WLX09_S_WARNING_LOG "
		        	      +"select m_year,m_quarter,bank_no,"
		                  +"warnaccount_tcnt,warnaccount_tbal,warnaccount_remit_tcnt,"
		                  +"warnaccount_refund_apply_cnt,warnaccount_refund_apply_amt,"
		                  +"warnaccount_refund_cnt,warnaccount_refund_amt,"		          
		                  +"user_id,user_name,update_date,"
		                  +"?,?,sysdate,'U'"
		                  +" from WLX09_S_WARNING WHERE bank_no=? AND m_year=?  and m_quarter= ? ";
				paramList.add(user_id) ;
				paramList.add(user_name) ;
				paramList.add(bank_no) ;
				paramList.add(nyear);
				paramList.add(nquarter) ;
				//updateDBSqlList.add(sqlCmdLog);				
				if(!this.updDbUsesPreparedStatement(sqlCmdLog,paramList)) {				
					errMsg = errMsg + "WLX09_S_WARNING_LOG相關資料寫入資料庫失敗";
				}
				paramList.clear() ;
				sqlCmd.append("UPDATE WLX09_S_WARNING SET ") ;
				sqlCmd.append(" m_year=?");paramList.add(nyear);
				sqlCmd.append(",m_quarter=?");paramList.add(nquarter ) ;
			    sqlCmd.append(" ,bank_no=?");paramList.add(bank_no ) ;
				sqlCmd.append(" ,warnaccount_tcnt=?");paramList.add(warnaccount_tcnt ) ;
        		sqlCmd.append(" ,warnaccount_tbal=?");paramList.add(warnaccount_tbal ); 
        		sqlCmd.append(" ,warnaccount_remit_tcnt=?");paramList.add(warnaccount_remit_tcnt ) ;
        		sqlCmd.append(" ,warnaccount_refund_apply_cnt=?");paramList.add(warnaccount_refund_apply_cnt ) ;
        	    sqlCmd.append(" ,warnaccount_refund_apply_amt=?");paramList.add(warnaccount_refund_apply_amt ) ;
        		sqlCmd.append(" ,warnaccount_refund_cnt=?");paramList.add(warnaccount_refund_cnt ) ;
        		sqlCmd.append(" ,warnaccount_refund_amt=?");paramList.add(warnaccount_refund_amt ) ;
	       	    sqlCmd.append(" ,user_id=?");paramList.add(user_id ) ;
	       	    sqlCmd.append(" ,user_name=?");paramList.add(user_name ) ;
	       	    sqlCmd.append(" ,update_date=sysdate") ;
	       	    sqlCmd.append("  where bank_no=?");paramList.add(bank_no); 
	       	    sqlCmd.append("  and m_year=? and m_quarter=? ");
	       	    paramList.add(nyear) ;
	       	    paramList.add(nquarter) ;
				/*       +"m_year='"+nyear+"',m_quarter='"+nquarter+"'"
				       +",bank_no='"+bank_no+"'"
					   +",warnaccount_tcnt='"+warnaccount_tcnt+"'"
               		   +",warnaccount_tbal='"+warnaccount_tbal+"'" 
               		   +",warnaccount_remit_tcnt='"+warnaccount_remit_tcnt+"'"
               		   +",warnaccount_refund_apply_cnt='"+warnaccount_refund_apply_cnt+"'"
               		   +",warnaccount_refund_apply_amt='"+warnaccount_refund_apply_amt+"'"
               		   +",warnaccount_refund_cnt='"+warnaccount_refund_cnt+"'"
               		   +",warnaccount_refund_amt='"+warnaccount_refund_amt+"'"
   			       	   +",user_id='"+user_id+"'"
   			       	   +",user_name='"+user_name+"'"
   			       	   +",update_date=sysdate"
   			       	   +" where bank_no='"+bank_no+"' and m_year='"+nyear+"' and m_quarter='"+nquarter+"'");*/
		        //updateDBSqlList.add(sqlCmd);		        
		        if(!this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)) {				
					errMsg = errMsg + "WLX09_S_WARNING相關資料寫入資料庫失敗";
				}
		        sqlCmd.setLength(0) ;
		        paramList.clear() ;
				//當寫進資料庫成功後，再寫進WLX_APPLY_LOCK檔中
				sqlLock = " select * from WLX_APPLY_LOCK aa where aa.M_YEAR =?"+
		        	    " and aa.M_QUARTER=?"+
		          	    " and aa.BANK_CODE=?"+
		          		" and aa.REPORT_NO='C06'";
		         paramList.add(nyear) ;
		         paramList.add(nquarter) ;
		         paramList.add(bank_no) ;
	  			 dbLock = DBManager.QueryDB_SQLParam(sqlLock,paramList,"");
	  			 paramList.clear() ;
	  			System.out.print("大小:"+String.valueOf(dbLock.size()));
				if(dbLock.size()!=0){
				   sqlLockLog = " insert into WLX_APPLY_LOCK_LOG "+
					   		  " select M_YEAR, M_QUARTER, BANK_CODE, REPORT_NO, "+
						 	  " ADD_USER, ADD_NAME, ADD_DATE, LOCK_OWN, LOCK_MGR, "+
							  " USER_ID, USER_NAME,UPDATE_DATE, "+
					     	  "? ,?,sysdate,'U'"+
					     	  " from WLX_APPLY_LOCK WHERE BANK_CODE=? AND M_YEAR=? and M_QUARTER= ? ";
				    paramList.add(user_id) ;
				    paramList.add(user_name) ;
				    paramList.add(bank_no) ;
				    paramList.add(nyear);
				    paramList.add(nquarter) ;
				    //updateDBSqlList.add(sqlLockLog);				    
				    if(!this.updDbUsesPreparedStatement(sqlLockLog,paramList)) {				
					    errMsg = errMsg + "WLX_APPLY_LOCK_LOG相關資料寫入資料庫失敗";
				    }
				    paramList.clear();
				    
					sqlLock = " update WLX_APPLY_LOCK set USER_ID = ?"+
							  " ,USER_NAME =?"+
					  		  " ,UPDATE_DATE =sysdate"+
					  		  " where M_YEAR = ? and M_QUARTER = ? and BANK_CODE =? "+
					  		  " and REPORT_NO ='C06'";
					paramList.add(user_id) ;
					paramList.add(user_name);
					paramList.add(nyear) ;
					paramList.add(nquarter) ;
					paramList.add(bank_no) ;
					//updateDBSqlList.add(sqlLock);				  
				}else{  //若是原本沒有鎖定資料的情況下，寫進一筆新的鎖定紀錄
					/*sqlLock = " insert into WLX_APPLY_LOCK values('"
		            	    +nyear+"','"
		              	    +nquarter+"','"
		                    +bank_no+"','C06','"
		                    +user_id+"','"
		                    +user_name+"',sysdate,'N','N','"
		              		+user_id+"','"
		              		+user_name+"',sysdate)"; */
		            sqlLock = "insert into WLX_APPLY_LOCK values( " ;
		            sqlLock += "?,?,?,'C06',?,?,sysdate,'N','N',?,?,sysdate)" ;
					paramList.add(nyear) ;
					paramList.add(nquarter) ;
					paramList.add(bank_no) ;
					paramList.add(user_id) ;
					paramList.add(user_name) ;
					paramList.add(user_id) ;
					paramList.add(user_name) ;
					
					//updateDBSqlList.add(sqlLock);	 
			    }

				//if(DBManager.updateDB(updateDBSqlList)){
				if(this.updDbUsesPreparedStatement(sqlLock,paramList)) {
					errMsg = errMsg + "相關資料寫入資料庫成功";
				}else{
				 	errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
				}
    	       }
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}
		return errMsg;
	}
    //94.04.01 add 主key更改為bank_no+seq_no by 2295
    public String DeleteDB(HttpServletRequest request,String bank_no,String lguser_id,String lguser_name) throws Exception{
		StringBuffer sqlCmd = new StringBuffer();
		String sqlCmdLog = "";
		String errMsg="";
		String nyear = ((String)request.getParameter("hyear")==null)?"":(String)request.getParameter("hyear");
	 	String nquarter= ((String)request.getParameter("hquarter")==null)?"":(String)request.getParameter("hquarter");
		String user_id=lguser_id;
	 	String user_name=lguser_name;
		String sqlLock = "";
		String sqlLockLog = "";
		List dbLock;

		try {
			   //List updateDBSqlList = new LinkedList();
			   List paramList =new ArrayList() ;
			   sqlCmd.append("SELECT * FROM WLX09_S_WARNING WHERE bank_no=? AND m_year=? and m_quarter= ? ");
			   paramList.add(bank_no) ;
			   paramList.add(nyear) ;
			   paramList.add(nquarter) ;
			   List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
			   sqlCmd.setLength(0) ;
			   paramList.clear() ;
			   System.out.println("WLX09_S_WARNING.size="+data.size());
			   if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法刪除<br>";
			   }else{
				  if(!nyear.equals("") && !nquarter.equals("")){
				 	/*寫入log檔中*/
				  sqlCmdLog = "INSERT INTO WLX09_S_WARNING_LOG "
		          		    +"select m_year,m_quarter,bank_no,"
		            		+"warnaccount_tcnt,warnaccount_tbal,warnaccount_remit_tcnt,"
		            		+"warnaccount_refund_apply_cnt,warnaccount_refund_apply_amt,"
		            		+"warnaccount_refund_cnt,warnaccount_refund_amt,"		          
		            		+"user_id,user_name,update_date,"
		            		+"?,?,sysdate,'D'"
		            		+" from WLX09_S_WARNING WHERE bank_no=? AND m_year=?  and m_quarter= ? ";
				   paramList.add(user_id) ;
				   paramList.add(user_name) ;
				   paramList.add(bank_no) ;
				   paramList.add(nyear) ;
				   paramList.add(nquarter) ;
				   //updateDBSqlList.add(sqlCmdLog);					   
				   if(!this.updDbUsesPreparedStatement(sqlCmdLog,paramList)) {				
					    errMsg = errMsg + "WLX09_S_WARNING_LOG相關資料寫入資料庫失敗";
				   }
				   paramList.clear() ;
				   sqlCmd.append(" delete WLX09_S_WARNING where bank_no=? and m_year=? and m_quarter=? ");
				   paramList.add(bank_no) ;
				   paramList.add(nyear) ;
				   paramList.add(nquarter) ;
		           //updateDBSqlList.add(sqlCmd);		           
		           if(!this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)) {				
					    errMsg = errMsg + "WLX09_S_WARNING刪除失敗";
				   }else{
				       errMsg = errMsg + "相關資料刪除成功";
				   }
		           paramList.clear() ;
		           sqlCmd.setLength(0) ;
					//當寫進資料庫成功後，再寫進WLX_APPLY_LOCK檔中
					sqlLock = " select * from WLX_APPLY_LOCK aa where aa.M_YEAR =?"+
				          	" and aa.M_QUARTER=?"+
		          			" and aa.BANK_CODE=?"+
		          			" and aa.REPORT_NO='C06'";	
				   paramList.add(nyear) ;
				   paramList.add(nquarter) ;
				   paramList.add(bank_no) ;
	  			   dbLock = DBManager.QueryDB_SQLParam(sqlLock,paramList,"");
	  			   paramList.clear() ;	
	  				System.out.print("大小:"+String.valueOf(dbLock.size()));
					if(dbLock.size()!=0){
						sqlLockLog = " insert into WLX_APPLY_LOCK_LOG "+
								   " select M_YEAR, M_QUARTER, BANK_CODE, REPORT_NO, "+
						 		   " ADD_USER, ADD_NAME, ADD_DATE, LOCK_OWN, LOCK_MGR, "+
						 		   " USER_ID, USER_NAME,UPDATE_DATE, "+
					     		   "?,?,sysdate,'D'"+
					     		   " from WLX_APPLY_LOCK WHERE BANK_CODE=? AND M_YEAR=?  and M_QUARTER=?";		
						paramList.add(user_id) ;
						paramList.add(user_name) ;
						paramList.add(bank_no) ;
						paramList.add(nyear) ;
						paramList.add(nquarter) ;
		                //updateDBSqlList.add(sqlLockLog);
		               
		                if(!this.updDbUsesPreparedStatement(sqlLockLog,paramList)) {				
					        errMsg = errMsg + "WLX_APPLY_LOCK_LOG相關資料寫入資料庫失敗";
				   	    }
		                paramList.clear() ;
						sqlLock = " delete WLX_APPLY_LOCK "+
					  			" where M_YEAR = ? and M_QUARTER = ? and BANK_CODE =?"+
					  			" and REPORT_NO ='C06'";	
						paramList.add(nyear) ;
						paramList.add(nquarter) ;
						paramList.add(bank_no) ;
					    //updateDBSqlList.add(sqlLock);
					    
					    if(!this.updDbUsesPreparedStatement(sqlLock,paramList)) {				
					        errMsg = errMsg + "WLX_APPLY_LOCK刪除失敗";
				   	    }
					 }
		          }
				  /*if(DBManager.updateDB(updateDBSqlList)){
					 errMsg = errMsg + "相關資料刪除成功";
				  }else{
				  	errMsg = errMsg + "相關資料刪除失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
				  }*/
    	   	   }
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料刪除失敗";
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
