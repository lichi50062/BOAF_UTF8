<%
// 94/10/17 first designed by 4180
// 95.05.24 fix 獨立成統一農(漁)貸資料辦理情形維護 by 2495
// 99.12.06 fix sqlInjection by 2808
//100.01.26 fix bug by 2295
//103.01.14 fix 讀取歷史資料時,增加為起始申報年月以後才計算 by 2295
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
      System.out.println("FX007W login timeout");
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
	// 95.06.02 add by 2495
	String LoadFlag = ( request.getParameter("LoadFlag")==null ) ? "false" : (String)request.getParameter("LoadFlag");
	System.out.println("LoadFlag:"+LoadFlag);
    //若程序為新增資料，則從request中抓出年月以判斷是否已有資料
    String checkyear = ( request.getParameter("checkyear")==null ) ? "" : (String)request.getParameter("checkyear");
	String checkmonth = ( request.getParameter("checkmonth")==null ) ? "" : (String)request.getParameter("checkmonth");

    //若程序為修改資料，則從request中抓出年月
    String myear = ( request.getParameter("myear")==null ) ? "" : (String)request.getParameter("myear");
	String mmonth = ( request.getParameter("mmonth")==null ) ? "" : (String)request.getParameter("mmonth");

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

    if(!Utility.CheckPermission(request,"FX007WB")){//無權限時,導向到LoginError.jsp   
        rd = application.getRequestDispatcher( LoginErrorPgName );
    }else{
    	//set next jsp
    	if(act.equals("new")){
    	
    		List inidate = getWLX07_M_IMPORTANT_INI();//先取得初始申報資料
   		 	request.setAttribute("WLX07_INI",inidate);
   		 	List hisdata = getWLX07_M_CREDIT(bank_no);  //是否已有申報資料.增加為起始申報年月以後
   		 	request.setAttribute("WLX07_HIS",hisdata);
    		List dbData = getWLX07_M_IMPORTANT(bank_no,checkyear,checkmonth);  //是否當年月已有資料
           
            if(dbData.size()==0){
        	  int temp = Integer.parseInt(checkmonth); //將參數轉換為上月份
        	  temp = temp-1;
        	  String lastmonth = String.valueOf(temp);
        	  if(temp==0){
        	 	lastmonth="12";
        	 	checkyear=String.valueOf(Integer.parseInt(checkyear)-1);
        	  }//處理一月時的情況

        	  dbData = getWLX07_M_IMPORTANT(bank_no,checkyear,lastmonth);  //取得上個月資料
        	  if(dbData.size()==0){ //上月若無資料
        	 	 rd = application.getRequestDispatcher( EditPgName +"?act=new&loaddata=false");
        	  }else{
        	  	
        	 	  rd = application.getRequestDispatcher( EditPgName +"?act=new&loaddata=ok");
        	  }	 
        	  
            }else{
        	   dbData = getWLX07_M_IMPORTANT(bank_no,checkyear,checkmonth);
    	       request.setAttribute("WLX07_M_Credit",dbData);
			   request.setAttribute("maintainInfo","select * from WLX07_M_Credit WHERE bank_no='" + bank_no+"' and m_year="+checkyear+" and m_month="+checkmonth);
			   rd = application.getRequestDispatcher( EditPgName +"?act=modify");
        	}
    	}else if(act.equals("Edit")){
        	//93.12.21設定異動者資訊======================================================================
			request.setAttribute("maintainInfo","select * from WLX07_M_Credit WHERE bank_no='" + bank_no+"' and m_year="+myear+" and m_month="+mmonth);
			//=======================================================================================================================
    	    List dbData = getWLX07_M_IMPORTANT(bank_no,myear,mmonth);
    	    request.setAttribute("WLX07_M_Credit",dbData);
    	    //95/01/09 fix by 4180     
        	int temp = Integer.parseInt(mmonth); //將參數轉換為下月份
        	temp = temp+1;
        	String nextmonth = String.valueOf(temp);
        	if(temp==13){
        	 	nextmonth="1";
        	 	myear=String.valueOf(Integer.parseInt(myear)+1);
        	}//處理一月時的情況

        	dbData = getWLX07_M_IMPORTANT(bank_no,myear,nextmonth);  //取得下個月資料
        	if(dbData.size()==0){ //下月若無資料
        	   rd = application.getRequestDispatcher( EditPgName +"?act=Edit&&editable=ok");
        	}else{
        	 	 rd = application.getRequestDispatcher( EditPgName +"?act=Edit&&editable=false");      	
        	} 	 
    	}else if(act.equals("List")){    		  
    	      List dbData = getWLX07_M_IMPORTANT(bank_no,"","");
    	      System.out.println("dbData="+dbData.size());
            List inidate = getWLX07_M_IMPORTANT_INI();
            List lockdate = getWLX07_M_IMPORTANT_LOCK(bank_no);             
    	      request.setAttribute("WLX07_M_Credit",dbData);
            request.setAttribute("WLX07_INI",inidate);
        	  request.setAttribute("WLX07_LOCK",lockdate);
        	  rd = application.getRequestDispatcher( ListPgName +"?bank_no="+bank_no);
        }else if(act.equals("Load")){
      	    String syear = ((String)request.getParameter("hyear")==null)?"":(String)request.getParameter("hyear");
	        String smonth = ((String)request.getParameter("hmonth")==null)?"":(String)request.getParameter("hmonth");
        	int temp = Integer.parseInt(smonth); //將參數轉換為上月份
          	temp = temp-1;
            String lyear=syear;
        	String lmonth = String.valueOf(temp);
        	if(temp==0){
        		lmonth="12";
        		lyear=String.valueOf(Integer.parseInt(syear)-1);
        	}//處理一月時的情況

        	List dbData = getWLX07_M_IMPORTANT(bank_no,lyear,lmonth);  //取得上個月資料
        	request.setAttribute("WLX07_M_Credit",dbData);
        	rd = application.getRequestDispatcher( EditPgName +"?act=new&loaddata=loaded&S_YEAR="+syear+"&S_MONTH="+smonth);
    	}else if(act.equals("Insert")){
    	    actMsg = InsertDB(request,bank_no,lguser_id,lguser_name);
        	rd = application.getRequestDispatcher( nextPgName+"?FX=FX007WB" );
    	}else if(act.equals("Update")){
    	    actMsg = UpdateDB(request,bank_no,lguser_id,lguser_name);
        	rd = application.getRequestDispatcher( nextPgName+"?FX=FX007WB");
    	}else if(act.equals("Delete")){
    	    actMsg = DeleteDB(request,bank_no,lguser_id,lguser_name);
        	rd = application.getRequestDispatcher( nextPgName+"?FX=FX007WB" );
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
    private final static String EditPgName = "/pages/FX007WB_Edit.jsp";
    private final static String ListPgName = "/pages/FX007WB_List.jsp";
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
   
    //94.04.01 add 主key更改為bank_no+seq_no by 2295
    private List getWLX07_M_IMPORTANT(String bank_no, String myear, String mmonth){
    		//程序為顯示畫面，查詢條件
    		List paramList = new ArrayList() ;
    		String sqlCmd = "select * from WLX07_M_Credit where bank_no= ? ";
    		paramList.add(bank_no) ;
    		if(!myear.equals("") && !mmonth.equals("")){
    			sqlCmd =sqlCmd + " and m_year=? and m_month= ? ";
    			paramList.add(myear) ;
    			paramList.add(mmonth) ;
    		}
    		sqlCmd += " order by M_YEAR desc, M_MONTH desc";

            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"m_year,m_month,creditmonth_cnt,creditmonth_amt,credityear_cnt_acc,"+
            									   "credityear_cnt_acc,credityear_amt_acc,credit_cnt,credit_bal,overcreditmonth_cnt,overcreditmonth_amt,overcredit_cnt,overcredit_bal,update_date,creditmonth_avgrate,credityear_avgrate");     
            return dbData;
    }
    
    
    //103.01.14 add 歷史資料從起使申報年月開始計算 by 2295
    private List getWLX07_M_CREDIT(String bank_no){
    		
    		List paramList = new ArrayList() ;
    		String sqlCmd = " select a.* "
						  + " from WLX07_M_Credit a,(select * from WLX_APPLY_INI where REPORT_NO=?)b" 
						  + " where bank_no= ?"
						  + " and ((a.m_year * 100 + a.m_month) >=(b.m_year * 100 + b.m_month))";
			paramList.add("C03") ;			  
    		paramList.add(bank_no) ;    		
    		sqlCmd += " order by a.M_YEAR desc, a.M_MONTH desc";

            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"m_year,m_month,creditmonth_cnt,creditmonth_amt,credityear_cnt_acc,"+
            									   "credityear_cnt_acc,credityear_amt_acc,credit_cnt,credit_bal,overcreditmonth_cnt,overcreditmonth_amt,overcredit_cnt,overcredit_bal,update_date,creditmonth_avgrate,credityear_avgrate");     
            return dbData;
    }
    
    private List getWLX07_M_IMPORTANT_INI(){
        List paramList = new ArrayList();
      String sqlCmd = "select * from WLX_APPLY_INI where REPORT_NO=?";
      paramList.add("C03");
      List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"m_year,m_month");
      if (dbData.size()!= 0){
          return dbData;
      }else{
          return null;
      }   
    }
    private List getWLX07_M_IMPORTANT_LOCK(String bank_no){
    	List paramList = new ArrayList() ;
      	String sqlCmd = " select M_YEAR,M_QUARTER from WLX_APPLY_LOCK "
 		    		  + " where BANK_CODE = ? "
 		       		  + " and REPORT_NO = 'C03'"          
 		       		  + " and((LOCK_OWN   =  'Y') or (LOCK_MGR = 'Y'))";
 		paramList.add(bank_no) ;
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"m_year,m_quarter");
      	if (dbData.size()!= 0){
            return dbData;
        }else{
          return null;
        }  
    }
    public String InsertDB(HttpServletRequest request,String bank_no,String lguser_id,String lguser_name) throws Exception{
    StringBuffer sqlCmd = new StringBuffer();
		String errMsg="";
		String nyear = ((String)request.getParameter("hyear")==null)?"":(String)request.getParameter("hyear");
		String nmonth = ((String)request.getParameter("hmonth")==null)?"":(String)request.getParameter("hmonth");      	
    	String creditmonth_cnt=((String)request.getParameter("creditmonth_cnt")==null)?"":(String)request.getParameter("creditmonth_cnt");
		String creditmonth_amt=((String)request.getParameter("creditmonth_amt")==null)?"":(String)request.getParameter("creditmonth_amt");
		String credityear_cnt_acc=((String)request.getParameter("credityear_cnt_acc")==null)?"":(String)request.getParameter("credityear_cnt_acc");
		String credityear_amt_acc=((String)request.getParameter("credityear_amt_acc")==null)?"":(String)request.getParameter("credityear_amt_acc");
		String credit_cnt=((String)request.getParameter("credit_cnt")==null)?"":(String)request.getParameter("credit_cnt");
		String credit_bal=((String)request.getParameter("credit_bal")==null)?"":(String)request.getParameter("credit_bal");		
		String overcreditmonth_cnt=((String)request.getParameter("overcreditmonth_cnt")==null)?"":(String)request.getParameter("overcreditmonth_cnt");
		String overcreditmonth_amt=((String)request.getParameter("overcreditmonth_amt")==null)?"":(String)request.getParameter("overcreditmonth_amt");
		String overcredit_cnt=((String)request.getParameter("overcredit_cnt")==null)?"":(String)request.getParameter("overcredit_cnt");
		String overcredit_bal=((String)request.getParameter("overcredit_bal")==null)?"":(String)request.getParameter("overcredit_bal");
		String creditmonth_avgrate=((String)request.getParameter("creditmonth_avgrate")==null)?"":(String)request.getParameter("creditmonth_avgrate");
		String credityear_avgrate=((String)request.getParameter("credityear_avgrate")==null)?"":(String)request.getParameter("credityear_avgrate");
		String user_id=lguser_id;
		String user_name=lguser_name;
		String sqlLock = "";
		String sqlLockLog = "";
		List dbLock;
		
		try {
   			//List updateDBSqlList = new LinkedList();
   			List paramList =new ArrayList() ;
			/*sqlCmd.append("INSERT INTO WLX07_M_Credit VALUES('"
				   + nyear + "','"
			       + nmonth +"','"
			   	   + bank_no +"','"
        	   	   + creditmonth_cnt +"','"
        	       + creditmonth_amt +"','"
        	       + credityear_cnt_acc +"','"
			       + credityear_amt_acc +"','"
			       + credit_cnt +"','"
			       + credit_bal +"','"
			       + overcreditmonth_cnt +"','"
        	       + overcreditmonth_amt +"','"
        	       + overcredit_cnt +"','"
			       + overcredit_bal +"','"
			       + lguser_id +"','"
			       + lguser_name
			       + "',sysdate)" );*/
			sqlCmd.append(" insert into WLX07_M_Credit Values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,sysdate,?,?)");
			paramList.add( nyear) ;
			paramList.add(nmonth) ;
			paramList.add(bank_no) ;
			paramList.add(creditmonth_cnt) ;
			paramList.add(creditmonth_amt) ;
			paramList.add(credityear_cnt_acc) ;
			paramList.add(credityear_amt_acc) ;
			paramList.add(credit_cnt) ;
			paramList.add(credit_bal) ;
			paramList.add(overcreditmonth_cnt) ;
			paramList.add(overcreditmonth_amt) ;
			paramList.add(overcredit_cnt) ;
			paramList.add(overcredit_bal) ;
			paramList.add(lguser_id) ;
			paramList.add(lguser_name) ;
			paramList.add(creditmonth_avgrate) ;
			paramList.add(credityear_avgrate) ;
 			
        	//updateDBSqlList.add(sqlCmd);
            this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList) ;    
	        //當寫進資料庫成功後，再寫進WLX_APPLY_LOCK檔中
	        sqlCmd.setLength(0) ;
	        paramList.clear() ;
			sqlLock = " select * from WLX_APPLY_LOCK aa where aa.M_YEAR =? "+
		          	" and aa.M_QUARTER=?"+
		          	" and aa.BANK_CODE=?"+
		          	" and aa.REPORT_NO='C03'";
		    paramList.add(nyear) ;
		    paramList.add(nmonth) ;
		    paramList.add(bank_no) ;
	  		dbLock = DBManager.QueryDB_SQLParam(sqlLock,paramList,"");   	  
	  		System.out.print("dbLock.size():"+String.valueOf(dbLock.size()));
	  		paramList.clear() ;
			if(dbLock.size()!=0){            
				sqlLock = " update WLX_APPLY_LOCK set USER_ID = ?"+
						" ,USER_NAME =?"+
						" ,UPDATE_DATE =sysdate"+
					  	" where M_YEAR = ? and M_QUARTER = ? and BANK_CODE =? "+
					  	" and REPORT_NO ='C03'";	
				paramList.add(user_id) ;
				paramList.add(user_name) ;
				paramList.add(nyear);
				paramList.add(nmonth) ;
				paramList.add(bank_no) ;
				
				//updateDBSqlList.add(sqlLock);  
			}else{  //若是原本沒有鎖定資料的情況下，寫進一筆新的鎖定紀錄
				/*sqlLock = " insert into WLX_APPLY_LOCK values('"
		        	    +nyear+"','"
		                +nmonth+"','"
		                +bank_no+"','C03','"
		                +user_id+"','"
		                +user_name+"',sysdate,'N','N','"
		                +user_id+"','"
		                +user_name+"',sysdate)";*/
		        sqlLock = "INSERT into WLX_APPLY_LOCK Values( " ;
		        sqlLock += "?,?,?,'C03',?,?,sysdate,'N','N',?,?,sysdate )"  ;
				paramList.add(nyear) ;
				paramList.add(nmonth) ;
				paramList.add(bank_no) ;
				paramList.add(user_id) ;
				paramList.add(user_name) ; 
				paramList.add(user_id);
				paramList.add(user_name );
				//updateDBSqlList.add(sqlLock);		  
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
		StringBuffer sqlCmd = new StringBuffer() ;
		String sqlCmdLog = "";
		String errMsg="";
		String nyear = ((String)request.getParameter("hyear")==null)?"":(String)request.getParameter("hyear");
	    String nmonth = ((String)request.getParameter("hmonth")==null)?"":(String)request.getParameter("hmonth");
	
	 
    	String creditmonth_cnt=((String)request.getParameter("creditmonth_cnt")==null)?"":(String)request.getParameter("creditmonth_cnt");
		System.out.println("creditmonth_cnt="+creditmonth_cnt);
		String creditmonth_amt=((String)request.getParameter("creditmonth_amt")==null)?"":(String)request.getParameter("creditmonth_amt");
		System.out.println("creditmonth_amt="+creditmonth_amt);
		String credityear_cnt_acc=((String)request.getParameter("credityear_cnt_acc")==null)?"":(String)request.getParameter("credityear_cnt_acc");
		System.out.println("credityear_cnt_acc="+credityear_cnt_acc);
		String credityear_amt_acc=((String)request.getParameter("credityear_amt_acc")==null)?"":(String)request.getParameter("credityear_amt_acc");
		System.out.println("credityear_amt_acc="+credityear_amt_acc);
		String credit_cnt=((String)request.getParameter("credit_cnt")==null)?"":(String)request.getParameter("credit_cnt");
		System.out.println("credit_cnt="+credit_cnt);
		String credit_bal=((String)request.getParameter("credit_bal")==null)?"":(String)request.getParameter("credit_bal");		
		System.out.println("credit_bal="+credit_bal);
		String overcreditmonth_cnt=((String)request.getParameter("overcreditmonth_cnt")==null)?"":(String)request.getParameter("overcreditmonth_cnt");
		System.out.println("overcreditmonth_cnt="+overcreditmonth_cnt);
		String overcreditmonth_amt=((String)request.getParameter("overcreditmonth_amt")==null)?"":(String)request.getParameter("overcreditmonth_amt");
		System.out.println("overcreditmonth_amt="+overcreditmonth_amt);
		String overcredit_cnt=((String)request.getParameter("overcredit_cnt")==null)?"":(String)request.getParameter("overcredit_cnt");
		System.out.println("overcredit_cnt="+overcredit_cnt);
		String overcredit_bal=((String)request.getParameter("overcredit_bal")==null)?"":(String)request.getParameter("overcredit_bal");
		System.out.println("overcredit_bal="+overcredit_bal);
 		String creditmonth_avgrate=((String)request.getParameter("creditmonth_avgrate")==null)?"":(String)request.getParameter("creditmonth_avgrate");
		String credityear_avgrate=((String)request.getParameter("credityear_avgrate")==null)?"":(String)request.getParameter("credityear_avgrate");
		
 
 
		String user_id=lguser_id;
	    String user_name=lguser_name;
	    String sqlLock = "";
		String sqlLockLog = "";
		List dbLock;


		try {
			 //List updateDBSqlList = new LinkedList();
			 List paramList = new ArrayList() ;
			 sqlCmd.append("SELECT * FROM WLX07_M_Credit WHERE bank_no=? AND m_year=? and m_month=?" ) ;
			 paramList.add(bank_no) ;
			 paramList.add(nyear) ;
			 paramList.add(nmonth) ;
			 List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
			 System.out.println("WLX07_M_Credit.size="+data.size());
 			 if (data.size() == 0){
			    errMsg = errMsg + "此筆資料不存在無法修改<br>";
			 }else{/*寫入log檔中*/
			 	System.out.print("STRAT sqlCmd ..........");
			    sqlCmd.setLength(0) ;
			    paramList.clear() ;
			    
			    sqlCmdLog = "INSERT INTO WLX07_M_Credit_LOG "
						  + "select m_year,m_month,bank_no,"
		            	  +"creditmonth_cnt,creditmonth_amt,credityear_cnt_acc,credityear_amt_acc,credit_cnt,credit_bal,overcreditmonth_cnt,overcreditmonth_amt,overcredit_cnt,overcredit_bal,"
		            	  +"user_id,user_name,update_date,"
		            	  +"?,?,sysdate,'U',creditmonth_avgrate,credityear_avgrate"
		            	  +" from WLX07_M_Credit WHERE bank_no=? " 
		            	  +" AND m_year=? and m_month=?";	
			    paramList.add(user_id);
			    paramList.add(user_name);
			    paramList.add(bank_no);
			    paramList.add(nyear);
			    paramList.add(nmonth);
			    	    
				
				if(!this.updDbUsesPreparedStatement(sqlCmdLog,paramList)) {					
				   errMsg = errMsg + "WLX07_M_Credit_LOG寫入資料庫失敗";
				}
				sqlCmd.setLength(0) ;
				paramList.clear() ;
				//updateDBSqlList.add(sqlCmdLog);
				sqlCmd.append("UPDATE WLX07_M_Credit SET ") ;				       
				sqlCmd.append(" creditmonth_cnt=?");paramList.add(creditmonth_cnt ) ; 
				sqlCmd.append(",creditmonth_amt=?");paramList.add(creditmonth_amt ) ;
				sqlCmd.append(",credityear_cnt_acc=?");paramList.add(credityear_cnt_acc ) ;
				sqlCmd.append(",credityear_amt_acc=?");paramList.add(credityear_amt_acc ) ;
   			    sqlCmd.append(",credit_cnt=?");paramList.add(credit_cnt ) ;
				sqlCmd.append(",credit_bal=?");paramList.add(credit_bal );
				sqlCmd.append(",overcreditmonth_cnt=?");paramList.add(overcreditmonth_cnt ) ;
				sqlCmd.append(",overcreditmonth_amt=?");paramList.add(overcreditmonth_amt ) ;
   			    sqlCmd.append(",overcredit_cnt=?");paramList.add(overcredit_cnt ) ;
				sqlCmd.append(",overcredit_bal=?");paramList.add(overcredit_bal ) ;
				sqlCmd.append(",creditmonth_avgrate=?");paramList.add(creditmonth_avgrate ) ;
				sqlCmd.append(",credityear_avgrate=?");paramList.add(credityear_avgrate ) ;
				sqlCmd.append(",user_id=?");paramList.add(user_id ) ;
   			    sqlCmd.append(",user_name=?");paramList.add(user_name ) ;
   			    sqlCmd.append(",update_date=sysdate");
   			    sqlCmd.append(" where bank_no=?");paramList.add(bank_no);
   			   	sqlCmd.append(" and m_year=?");paramList.add(nyear);
   			   	sqlCmd.append(" and m_month=?");paramList.add(nmonth) ;
				if(!this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)) {					
				   errMsg = errMsg + "WLX07_M_Credit寫入資料庫失敗";
				}
				sqlCmd.setLength(0) ;
				paramList.clear() ;
   			    //System.out.print("updateDBSql="+sqlCmd);	
		        //updateDBSqlList.add(sqlCmd);
				//當寫進資料庫成功後，再寫進WLX_APPLY_LOCK檔中
				sqlLock = " select * from WLX_APPLY_LOCK aa where aa.M_YEAR =?"
						+ " and aa.M_QUARTER=?"
						+ " and aa.BANK_CODE=?"
						+ " and aa.REPORT_NO=?";
		        paramList.add(nyear) ;
		        paramList.add(nmonth) ;
		        paramList.add(bank_no) ;
		        paramList.add("C03") ;
	  			dbLock = DBManager.QueryDB_SQLParam(sqlLock,paramList,"");   
	  			paramList.clear() ;
	  			System.out.print("大小:"+String.valueOf(dbLock.size()));
				if(dbLock.size()!=0){
					sqlLockLog = " insert into WLX_APPLY_LOCK_LOG "
							   + " select M_YEAR, M_QUARTER, BANK_CODE, REPORT_NO, "
							   + " ADD_USER, ADD_NAME, ADD_DATE, LOCK_OWN, LOCK_MGR, "
							   + " USER_ID, USER_NAME,UPDATE_DATE, "
							   + "?,?,sysdate,'U'"
							   + " from WLX_APPLY_LOCK WHERE BANK_CODE=? AND M_YEAR=? and M_QUARTER=?";
					paramList.add(user_id) ;
					paramList.add(user_name) ;
					paramList.add(bank_no) ;
					paramList.add(nyear) ;
					paramList.add(nmonth) ;
					if(!this.updDbUsesPreparedStatement(sqlLockLog,paramList)) {					
				 		errMsg = errMsg + "WLX_APPLY_LOCK_LOG寫入資料庫失敗";
					}
					paramList.clear() ;
		           //updateDBSqlList.add(sqlLockLog);		                 
				    sqlLock = " update WLX_APPLY_LOCK set USER_ID =?"
				   		   + " ,USER_NAME =?"
				   		   + " ,UPDATE_DATE =sysdate"
				   		   + " where M_YEAR = ? and M_QUARTER = ? and BANK_CODE =?"
				   		   + " and REPORT_NO =?";		
				    paramList.add(user_id) ;
				    paramList.add(user_name) ;
				    paramList.add(nyear) ;
				    paramList.add(nmonth) ;
				    paramList.add(bank_no) ;
				    paramList.add("C03") ;
				   //updateDBSqlList.add(sqlLock);	
				   		  
				}else{  //若是原本沒有鎖定資料的情況下，寫進一筆新的鎖定紀錄
					/*sqlLock = " insert into WLX_APPLY_LOCK values('"
		            	    +nyear+"','"
		              	    +nmonth+"','"
		              	    +bank_no+"','C03','"
		              		+user_id+"','"
		              		+user_name+"',sysdate,'N','N','"
		              		+user_id+"','"
		              		+user_name+"',sysdate)"; 					  
					//  updateDBSqlList.add(sqlLock);*/
					sqlLock = "insert into WLX_APPLY_LOCK Values( " ;
					sqlLock += "?,?,?,?,?,?,sysdate,'N','N',?,?,sysdate ) " ;
					paramList.add(nyear) ;
					paramList.add(nmonth) ;
					paramList.add(bank_no) ;
					paramList.add("C03") ;
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
    	   	 }//end of 資料存在
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}
		return errMsg;
	}
    //94.04.01 add 主key更改為bank_no+seq_no by 2295
    public String DeleteDB(HttpServletRequest request,String bank_no,String lguser_id,String lguser_name) throws Exception{
		StringBuffer sqlCmd = new StringBuffer() ;
		String sqlCmdLog = "";
		String errMsg="";
		String nyear = ((String)request.getParameter("hyear")==null)?"":(String)request.getParameter("hyear");
	    String nmonth = ((String)request.getParameter("hmonth")==null)?"":(String)request.getParameter("hmonth");
		String user_id=lguser_id;
	    String user_name=lguser_name;
	    String sqlLock = "";
		String sqlLockLog = "";
		List dbLock;

		try {
			   //List updateDBSqlList = new LinkedList();
			   List paramList = new ArrayList () ;
			   sqlCmd.append("SELECT * FROM WLX07_M_Credit WHERE bank_no=? AND m_year=? and m_month=?") ;
			   paramList.add(bank_no) ;
			   paramList.add(nyear) ;
			   paramList.add(nmonth) ;
			   List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
			   sqlCmd.setLength(0) ;
			   paramList.clear() ;
			   System.out.println("WLX07_M_Credit.size="+data.size());
			   if(data.size() == 0){
				  errMsg = errMsg + "此筆資料不存在無法刪除<br>";
			   }else{
				  if(!nyear.equals("") && !nmonth.equals("")){
				 	/*寫入log檔中*/
					sqlCmdLog = "INSERT INTO WLX07_M_Credit_LOG "
							  + "select m_year,m_month,bank_no,"
		            		  + "creditmonth_cnt,creditmonth_amt,credityear_cnt_acc,credityear_amt_acc,credit_cnt,credit_bal,overcreditmonth_cnt,overcreditmonth_amt,overcredit_cnt,overcredit_bal,"		            
		            		  + "user_id,user_name,update_date,"
		            		  + "?,?,sysdate,'D',creditmonth_avgrate,credityear_avgrate"
		            		  + " from WLX07_M_Credit WHERE bank_no=? AND m_year=? and m_month= ? ";
				 	paramList.add(user_id) ;
				 	paramList.add(user_name) ;
				 	paramList.add(bank_no) ;
				 	paramList.add(nyear) ;
				 	paramList.add(nmonth) ;
				 	this.updDbUsesPreparedStatement(sqlCmdLog,paramList) ;
				 	paramList.clear() ;
					//updateDBSqlList.add(sqlCmdLog);
				    sqlCmd.append(" delete WLX07_M_Credit where bank_no=? and m_year=? and m_month=? ");
				    paramList.add(bank_no) ;
				    paramList.add(nyear) ;
				    paramList.add(nmonth) ;
				    this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList) ;
				    sqlCmd.setLength(0) ;
				    paramList.clear() ;
		            //updateDBSqlList.add(sqlCmd);
 					//當寫進資料庫成功後，再寫進WLX_APPLY_LOCK檔中
					sqlLock = " select * from WLX_APPLY_LOCK aa where aa.M_YEAR =? "
							+ " and aa.M_QUARTER=?"
							+ " and aa.BANK_CODE=?"
							+ " and aa.REPORT_NO='?";		
 					paramList.add(nyear) ;
 					paramList.add(nmonth) ;
 					paramList.add(bank_no) ;
					paramList.add("C03") ;
	  				dbLock = DBManager.QueryDB_SQLParam(sqlLock,paramList,"");   
	  				paramList.clear() ;
	  				
	  				System.out.print("大小:"+String.valueOf(dbLock.size()));
					if(dbLock.size()!=0){
						sqlLockLog = " insert into WLX_APPLY_LOCK_LOG "
								  + " select M_YEAR, M_QUARTER, BANK_CODE, REPORT_NO, "
								  + " ADD_USER, ADD_NAME, ADD_DATE, LOCK_OWN, LOCK_MGR, "
								  + " USER_ID, USER_NAME,UPDATE_DATE, "
								  + "? ,?,sysdate,'D'"
								  + " from WLX_APPLY_LOCK WHERE BANK_CODE=? AND M_YEAR=? and M_QUARTER= ? ";
						paramList.add(user_id) ;
						paramList.add(user_name) ;
						paramList.add(bank_no) ;
						paramList.add(nyear) ;
						paramList.add(nmonth) ;
		                //updateDBSqlList.add(sqlLockLog);
		                this.updDbUsesPreparedStatement(sqlLockLog,paramList) ;
		                paramList.clear() ;
						sqlLock = " delete WLX_APPLY_LOCK "
								+ " where M_YEAR = ? and M_QUARTER = ? and BANK_CODE =? "
								+ " and REPORT_NO =?";		
						paramList.add(nyear) ;
						paramList.add(nmonth) ;
						paramList.add(bank_no) ;
						paramList.add("C03") ;
						//updateDBSqlList.add(sqlLock);
						this.updDbUsesPreparedStatement(sqlLock,paramList) ;
					}
		          }//end of nyear & nmonth !- ""
		          errMsg = errMsg + "相關資料刪除成功";
				  /*if(DBManager.updateDB(updateDBSqlList)){
					 errMsg = errMsg + "相關資料刪除成功";
				  }else{
				  	errMsg = errMsg + "相關資料刪除失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
				  }*/
    	   	   }//end of 資料存在
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
