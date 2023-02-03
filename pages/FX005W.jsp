<%
//94.12.22 fix loaddata for check date by lilic0c0 4183 
//94.10.17 first designed by 4180
//96.11.28 add 金融卡本月交易次數/金融卡本月交易金額(元)/本年累計交易次數/本年累計交易金額(元) by 2295
//99.12.03 fix sqlInjection by 2808
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
      System.out.println("FX005W login timeout");
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

    if(!CheckPermission(request)){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );
    }else{
    	//set next jsp
    	if(act.equals("new")){   		
   		 	List inidate = getWLX05_M_ATM_INI();//先取得初始申報資料
   		 	request.setAttribute("WLX05_INI",inidate);
   		 	List hisdata = getWLX05_M_ATM(bank_no,"","");  //是否已有申報資料
   		 	request.setAttribute("WLX05_HIS",hisdata);
   		 	
    	 	List dbData = getWLX05_M_ATM(bank_no,checkyear,checkmonth);  //是否當年月已有資料
            if(dbData.size()==0){
        	 	int temp = Integer.parseInt(checkmonth); //將參數轉換為上月份
        	 	temp = temp-1;
        	 	String lastmonth = String.valueOf(temp);
        	 	if(temp==0){
        	  	   lastmonth="12";
        	 	   checkyear=String.valueOf(Integer.parseInt(checkyear)-1);
        	 	}//處理一月時的情況
 
        	 	dbData = getWLX05_M_ATM(bank_no,checkyear,lastmonth);  //取得上個月資料
        	 	if(dbData.size()==0) //上月若無資料
        	 	   rd = application.getRequestDispatcher( EditPgName +"?act=new&loaddata=false");
        	 	else
        	 	   rd = application.getRequestDispatcher( EditPgName +"?act=new&loaddata=ok");
        	}else{
        	   dbData = getWLX05_M_ATM(bank_no,checkyear,checkmonth);
    	       request.setAttribute("WLX05_M_ATM",dbData);
			   request.setAttribute("maintainInfo","select * from WLX05_M_ATM WHERE bank_no='" + bank_no+"'");
			   rd = application.getRequestDispatcher( EditPgName +"?act=modify");
        	}

    	}else if(act.equals("Edit")){
    		//93.12.21設定異動者資訊======================================================================
			request.setAttribute("maintainInfo","select * from WLX05_M_ATM WHERE bank_no='" + bank_no+"'");
			//=======================================================================================================================
    	    List dbData = getWLX05_M_ATM(bank_no,myear,mmonth);
    	    request.setAttribute("WLX05_M_ATM",dbData);
    	    //95/01/09 fix by 4180    
        	int temp = Integer.parseInt(mmonth); //將參數轉換為下月份
        	temp = temp+1;
        	String nextmonth = String.valueOf(temp);
        	if(temp==13){
        	   nextmonth="1";
        	   myear=String.valueOf(Integer.parseInt(myear)+1);
        	}//處理一月時的情況

        	dbData = getWLX05_M_ATM(bank_no,myear,nextmonth);  //取得下個月資料
        	if(dbData.size()==0) //下月若無資料
        	 rd = application.getRequestDispatcher( EditPgName +"?act=Edit&&editable=ok");
        	else
        	 rd = application.getRequestDispatcher( EditPgName +"?act=Edit&&editable=false");
    	}else if(act.equals("List")){
    	    List dbData = getWLX05_M_ATM(bank_no,"","");
            List inidate = getWLX05_M_ATM_INI();
            List lockdate = getWLX05_M_ATM_LOCK(bank_no);            
    	    request.setAttribute("WLX05_M_ATM",dbData);
            request.setAttribute("WLX05_INI",inidate);
            request.setAttribute("WLX05_LOCK",lockdate);
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

        	List dbData = getWLX05_M_ATM(bank_no,lyear,lmonth);  //取得上個月資料
        	request.setAttribute("WLX05_M_ATM",dbData);
        	rd = application.getRequestDispatcher( EditPgName +"?act=new&loaddata=loaded&S_YEAR="+syear+"&S_MONTH="+smonth);
    	}else if(act.equals("Insert")){
    	    actMsg = InsertDB(request,bank_no,lguser_id,lguser_name);
        	rd = application.getRequestDispatcher( nextPgName+"?FX=FX005W" );
    	}else if(act.equals("Update")){
    	    actMsg = UpdateDB(request,bank_no,lguser_id,lguser_name);
        	rd = application.getRequestDispatcher( nextPgName+"?FX=FX005W");
    	}else if(act.equals("Delete")){
    	    actMsg = DeleteDB(request,bank_no,lguser_id,lguser_name);
        	rd = application.getRequestDispatcher( nextPgName+"?FX=FX005W" );
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
    private final static String EditPgName = "/pages/FX005W_Edit.jsp";
    private final static String ListPgName = "/pages/FX005W_List.jsp";
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    private boolean CheckPermission(HttpServletRequest request){//檢核權限
    	    boolean CheckOK=false;
    	    HttpSession session = request.getSession();
            Properties permission = ( session.getAttribute("FX005W")==null ) ? new Properties() : (Properties)session.getAttribute("FX005W");
            if(permission == null){
              System.out.println("FX005W.permission == null");
            }else{
               System.out.println("FX005W.permission.size ="+permission.size());

            }
            //只要有Query的權限,就可以進入畫面
        	if(permission != null && permission.get("Q") != null && permission.get("Q").equals("Y")){
        	   CheckOK = true;//Query
        	}
        	return CheckOK;
    }
    //94.04.01 add 主key更改為bank_no+seq_no by 2295
    private List getWLX05_M_ATM(String bank_no, String myear, String mmonth){
    		//程序為顯示畫面，查詢條件
    		List paramList = new ArrayList() ;
    		String sqlCmd = "select * from WLX05_M_ATM where bank_no= ? ";
    		paramList.add(bank_no) ;
    		if(!myear.equals("") && !mmonth.equals("")){
    			sqlCmd =sqlCmd + " and m_year= ? and m_month= ? ";
    			paramList.add(myear) ;
    			paramList.add(mmonth) ;
    		}
    		sqlCmd =sqlCmd + " order by M_YEAR desc, M_MONTH desc";

		    //List dbData = DBManager.QueryDB(sqlCmd,"m_year,m_month,push_debitcard_cnt,use_debitcard_cnt,push_bincard_cnt,use_bincard_cnt,atm_cnt,month_tran_cnt,year_acctran_cnt,month_tran_amt,year_acctran_amt,canc_debitcard_cnt,canc_bincard_cnt");
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"m_year,m_month,push_debitcard_cnt,use_debitcard_cnt,push_bincard_cnt,use_bincard_cnt,atm_cnt,month_tran_cnt,year_acctran_cnt,month_tran_amt,year_acctran_amt,user_id,user_name,update_date,canc_debitcard_cnt,canc_bincard_cnt,debitcard_month_tran_cnt,debitcard_year_acctran_cnt,debitcard_month_tran_amt,debitcard_year_acctran_amt");
            return dbData;
    }
    private List getWLX05_M_ATM_INI(){
          List paramList = new ArrayList();
          String sqlCmd = "select * from WLX_APPLY_INI where REPORT_NO=?";
          paramList.add("C01");
          List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"m_year,m_month");

      	  if (dbData.size()!= 0)
              return dbData;
      	  else
             return null;
    }
    private List getWLX05_M_ATM_LOCK(String bank_no){
    	    List paramList = new ArrayList() ;
      		String sqlCmd = " select M_YEAR,M_QUARTER from WLX_APPLY_LOCK "
 		       			  + " where BANK_CODE = ?"
 		       			  + " and REPORT_NO = 'C01'"          
 		       			  + " and((LOCK_OWN   =  'Y') or (LOCK_MGR = 'Y'))"; 		       			  
 			paramList.add(bank_no) ;
        	List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"m_year,m_quarter");

      		if (dbData.size()!= 0)
          		return dbData;
      		else
          		return null;
    }
    public String InsertDB(HttpServletRequest request,String bank_no,String lguser_id,String lguser_name) throws Exception{    
        StringBuffer sqlCmd = new StringBuffer();
	    String errMsg="";
        
	    String nyear = ((String)request.getParameter("hyear")==null)?"":(String)request.getParameter("hyear");
	    String nmonth = ((String)request.getParameter("hmonth")==null)?"":(String)request.getParameter("hmonth");
        String debitcard = ((String)request.getParameter("debitcard")==null)?"":(String)request.getParameter("debitcard");
  	    String usedebitcard = ((String)request.getParameter("usedebitcard")==null)?"":(String)request.getParameter("usedebitcard");
	    String cancdebitcard = ((String)request.getParameter("cancdebitcard")==null)?"":(String)request.getParameter("cancdebitcard");
	    String bincard = ((String)request.getParameter("bincard")==null)?"":(String)request.getParameter("bincard");
	    String cancbincard = ((String)request.getParameter("cancbincard")==null)?"":(String)request.getParameter("cancbincard");
	    String usebincard = ((String)request.getParameter("usebincard")==null)?"":(String)request.getParameter("usebincard");
	    String setup_atm = ((String)request.getParameter("setup_atm")==null)?"":(String)request.getParameter("setup_atm");
	    String monthtran_cnt = ((String)request.getParameter("monthtran_cnt")==null)?"":(String)request.getParameter("monthtran_cnt");
	    String yeartran_cnt = ((String)request.getParameter("yeartran_cnt")==null)?"":(String)request.getParameter("yeartran_cnt");
	    String monthtran_amt = ((String)request.getParameter("monthtran_amt")==null)?"":(String)request.getParameter("monthtran_amt");
	    String yeartran_amt = ((String)request.getParameter("yeartran_amt")==null)?"":(String)request.getParameter("yeartran_amt");
	    String user_id=lguser_id;
	    String user_name=lguser_name;
	    String sqlLock = "";
		String sqlLockLog = "";
		List dbLock;
		
		try {
   			 //List updateDBSqlList = new LinkedList();
   			 List paramList = new ArrayList() ;
			/* sqlCmd = "INSERT INTO WLX05_M_ATM VALUES('"+ nyear+ "','"+nmonth+"','"+bank_no+"'"
                	+",'"+debitcard+"','"+usedebitcard+"','"+bincard+"','"+usebincard+"','"+setup_atm+"','"
                	+monthtran_cnt+"','"+yeartran_cnt+"','"+monthtran_amt+"','"+yeartran_amt+"','"
					+lguser_id+"','"+lguser_name+"',sysdate,'"+cancdebitcard+"','"+cancbincard+"')"; */
		    sqlCmd.append(" insert into WLX05_M_ATM values ( ") ;
		    sqlCmd.append("?,?,?,?,?,?,?,?,?,?,");
		    sqlCmd.append("?,?,?,?,sysdate,?,?)");
			paramList.add(nyear) ;
			paramList.add(nmonth) ;
			paramList.add(bank_no) ;
			paramList.add(debitcard) ;
			paramList.add(usedebitcard) ;
			paramList.add(bincard) ;
			paramList.add(usebincard) ;
			paramList.add(setup_atm) ;
			paramList.add(monthtran_cnt) ;
			paramList.add(yeartran_cnt) ;
			paramList.add(monthtran_amt) ;
			paramList.add(yeartran_amt) ;
			paramList.add(lguser_id) ;
			paramList.add(lguser_name) ;
			paramList.add(cancdebitcard) ;
			paramList.add(cancbincard) ;
                //updateDBSqlList.add(sqlCmd);
            this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList) ;
            sqlCmd.setLength(0) ;
            paramList.clear() ;
                    //當寫進資料庫成功後，再寫進WLX_APPLY_LOCK檔中
			 sqlLock = " select * from WLX_APPLY_LOCK aa where aa.M_YEAR =?"+
		     	     " and aa.M_QUARTER=?"+
		        	 " and aa.BANK_CODE=?"+
		          	 " and aa.REPORT_NO='C01'";
		     paramList.add(nyear) ;
		     paramList.add(nmonth) ;
		     paramList.add(bank_no) ;
	  		 dbLock = DBManager.QueryDB_SQLParam(sqlLock,paramList,"");   
	  
	  		System.out.print("大小:"+String.valueOf(dbLock.size()));
	  
			if(dbLock.size()!=0){    
				paramList.clear() ;
				sqlLock = " update WLX_APPLY_LOCK set USER_ID =? "+
						  " ,USER_NAME =?"+
						  " ,UPDATE_DATE =sysdate"+
					  	  " where M_YEAR = ? and M_QUARTER = ? and BANK_CODE =? "+
					  	  " and REPORT_NO ='C01'";
				paramList.add(user_id) ;
				paramList.add(user_name);
				paramList.add(nyear) ;
				paramList.add(nmonth) ;
				paramList.add(bank_no) ;
			    //updateDBSqlList.add(sqlLock);  
			}else{  //若是原本沒有鎖定資料的情況下，寫進一筆新的鎖定紀錄
				paramList.clear() ;
				sqlLock = " insert into WLX_APPLY_LOCK values(?,?,?,?,?,?,sysdate,?,?,?,?,sysdate)"; 
				paramList.add(nyear) ;
				paramList.add(nmonth) ;
				paramList.add(bank_no);
				paramList.add("C01");
				paramList.add(user_id);
				paramList.add(user_name) ;
				paramList.add("N") ;
				paramList.add("N") ;
				paramList.add(user_id) ;
				paramList.add(user_name) ;
				
				//updateDBSqlList.add(sqlLock);		  
			}
			if(this.updDbUsesPreparedStatement(sqlLock,paramList)){
			//if(DBManager.updateDB(updateDBSqlList)){
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
		StringBuffer sqlCmd = new StringBuffer();
		String sqlCmdLog = "";
		String errMsg="";
		String nyear = ((String)request.getParameter("hyear")==null)?"":(String)request.getParameter("hyear");
	  	String nmonth = ((String)request.getParameter("hmonth")==null)?"":(String)request.getParameter("hmonth");
		String debitcard = ((String)request.getParameter("debitcard")==null)?"":(String)request.getParameter("debitcard");
		String usedebitcard = ((String)request.getParameter("usedebitcard")==null)?"":(String)request.getParameter("usedebitcard");
		String cancdebitcard = ((String)request.getParameter("cancdebitcard")==null)?"":(String)request.getParameter("cancdebitcard");
		String bincard = ((String)request.getParameter("bincard")==null)?"":(String)request.getParameter("bincard");
		String cancbincard = ((String)request.getParameter("cancbincard")==null)?"":(String)request.getParameter("cancbincard");
		String usebincard = ((String)request.getParameter("usebincard")==null)?"":(String)request.getParameter("usebincard");
		String setup_atm = ((String)request.getParameter("setup_atm")==null)?"":(String)request.getParameter("setup_atm");
		String monthtran_cnt = ((String)request.getParameter("monthtran_cnt")==null)?"":(String)request.getParameter("monthtran_cnt");
		String yeartran_cnt = ((String)request.getParameter("yeartran_cnt")==null)?"":(String)request.getParameter("yeartran_cnt");
		String monthtran_amt = ((String)request.getParameter("monthtran_amt")==null)?"":(String)request.getParameter("monthtran_amt");
		String yeartran_amt = ((String)request.getParameter("yeartran_amt")==null)?"":(String)request.getParameter("yeartran_amt");

		String user_id=lguser_id;
	    String user_name=lguser_name;
	    String sqlLock = "";
		String sqlLockLog = "";
		List dbLock;

		try {
			 //List updateDBSqlList = new LinkedList();
			 List paramList = new ArrayList() ;
			 sqlCmd.append("SELECT * FROM WLX05_M_ATM WHERE bank_no=?  AND m_year=? and m_month=? ");
			 paramList.add(bank_no) ;
			 paramList.add(nyear) ;
			 paramList.add(nmonth) ;
			 List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,push_debitcard_cnt,use_debitcard_cnt,push_bincard_cnt,use_bincard_cnt,atm_cnt,month_tran_cnt,year_acctran_cnt,month_tran_amt,year_acctran_amt,user_id,user_name,update_date,canc_debitcard_cnt,canc_bincard_cnt");
			 System.out.println("WLX05_M_ATM.size="+data.size());

			 if (data.size() == 0){
				 errMsg = errMsg + "此筆資料不存在無法修改<br>";
			 }else{/*寫入log檔中*/
				 paramList.clear() ;
			     sqlCmdLog = "INSERT INTO WLX05_M_ATM_LOG "
				  		   +"select m_year,m_month,bank_no,push_debitcard_cnt,"
		                   +"use_debitcard_cnt,push_bincard_cnt,use_bincard_cnt,atm_cnt,"
		                   +"month_tran_cnt,year_acctran_cnt,month_tran_amt,year_acctran_amt,"
		                   +"user_id,user_name,update_date,"
		                   +"?,?,sysdate,'U'"
		                   +",canc_debitcard_cnt,canc_bincard_cnt"
		                   +" from WLX05_M_ATM WHERE bank_no=? AND m_year=? and m_month= ? ";
				paramList.add(user_id) ;
				paramList.add(user_name) ;
				paramList.add(bank_no) ;
				paramList.add(nyear) ;
				paramList.add(nmonth);
				this.updDbUsesPreparedStatement(sqlCmdLog,paramList) ;
				//updateDBSqlList.add(sqlCmdLog);
				sqlCmd.setLength(0) ;
				paramList.clear() ;
			    sqlCmd.append("UPDATE WLX05_M_ATM SET ");
				sqlCmd.append("m_year=?");paramList.add(nyear);
				sqlCmd.append(",m_month=?");paramList.add(nmonth);
				sqlCmd.append(",bank_no=?");paramList.add(bank_no);
				sqlCmd.append(",push_debitcard_cnt=?");paramList.add(debitcard);
				sqlCmd.append(",use_debitcard_cnt=?");paramList.add(usedebitcard );
				sqlCmd.append(",push_bincard_cnt=?");paramList.add(bincard );
				sqlCmd.append(",use_bincard_cnt=?");paramList.add(usebincard);
				sqlCmd.append(",atm_cnt=?");paramList.add(setup_atm ) ;
   				sqlCmd.append(",month_tran_cnt=?");paramList.add(monthtran_cnt ) ;
   				sqlCmd.append(",year_acctran_cnt=?");paramList.add(yeartran_cnt ) ;
   				sqlCmd.append(",month_tran_amt=?");paramList.add(monthtran_amt );
   			    sqlCmd.append(",year_acctran_amt=?");paramList.add(yeartran_amt );
   			    sqlCmd.append(",user_id=?");paramList.add(user_id ) ;
   			    sqlCmd.append(",user_name=?");paramList.add(user_name );
   			    sqlCmd.append(",update_date=sysdate") ;
   			    sqlCmd.append(",canc_debitcard_cnt=?");paramList.add(cancdebitcard ) ;
   			    sqlCmd.append(",canc_bincard_cnt=?");paramList.add(cancbincard ) ;
   			    sqlCmd.append(" where bank_no=? and m_year=? and m_month=? ");
   			    paramList.add(bank_no) ;
   			    paramList.add(nyear) ;
   			    paramList.add(nmonth) ;
				this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList) ;
				sqlCmd.setLength(0) ;
				paramList.clear() ;
		        //updateDBSqlList.add(sqlCmd);
				//當寫進資料庫成功後，再寫進WLX_APPLY_LOCK檔中
				sqlLock = " select * from WLX_APPLY_LOCK aa where aa.M_YEAR =?"+
		                " and aa.M_QUARTER=?"+
		                " and aa.BANK_CODE=?"+
		                " and aa.REPORT_NO='C01'";
		        paramList.add(nyear) ;
		        paramList.add(nmonth) ;
		        paramList.add(bank_no) ;
		        
	  			dbLock = DBManager.QueryDB_SQLParam(sqlLock,paramList,"");
	  			paramList.clear() ;
	  			
	  		    System.out.print("大小:"+String.valueOf(dbLock.size()));
				if(dbLock.size()!=0){
					sqlLockLog = " insert into WLX_APPLY_LOCK_LOG "+
						       " select M_YEAR, M_QUARTER, BANK_CODE, REPORT_NO, "+
						       " ADD_USER, ADD_NAME, ADD_DATE, LOCK_OWN, LOCK_MGR, "+
						       " USER_ID, USER_NAME,UPDATE_DATE, "+
					           "?,?,sysdate,'U'"+
					           " from WLX_APPLY_LOCK WHERE BANK_CODE=? AND M_YEAR=? and M_QUARTER= ?";
		             paramList.add(user_id) ;
		             paramList.add(user_name) ;
		             paramList.add(bank_no) ;
		             paramList.add(nyear);
		             paramList.add(nmonth);
		            //updateDBSqlList.add(sqlLockLog);
		            this.updDbUsesPreparedStatement(sqlLockLog,paramList) ;
		            paramList.clear() ;
					sqlLock = " update WLX_APPLY_LOCK set USER_ID = ? "+
					        " ,USER_NAME =?"+
					        " ,UPDATE_DATE =sysdate"+
					        " where M_YEAR = ? and M_QUARTER = ? and BANK_CODE =? "+
					        " and REPORT_NO ='C01'";
					paramList.add(user_id) ;
					paramList.add(user_name) ;
					paramList.add(nyear) ;
					paramList.add(nmonth) ;
					paramList.add(bank_no);
				    //updateDBSqlList.add(sqlLock);
				    this.updDbUsesPreparedStatement(sqlLock,paramList) ;
			    }else{  //若是原本沒有鎖定資料的情況下，寫進一筆新的鎖定紀錄
					/*sqlLock = " insert into WLX_APPLY_LOCK values('"
		                    +nyear+"','"
		                    +nmonth+"','"
		                    +bank_no+"','C01','"
		                    +user_id+"','"
		                    +user_name+"',sysdate,'N','N','"
		                    +user_id+"','"
		                    +user_name+"',sysdate)";*/ 
			        sqlLock = "insert into WLX_APPLY_LOCK Values(?,?,?,?,?,?,sysdate,?,?,?,?,sysdate)" ;
					paramList.add(nyear) ;
					paramList.add(nmonth) ;
					paramList.add(bank_no) ;
					paramList.add("C01") ;
					paramList.add(user_id) ;
					paramList.add(user_name) ;
					paramList.add("N") ;
					paramList.add("N") ;
					paramList.add(user_id) ;
					paramList.add(user_name) ;
					//updateDBSqlList.add(sqlLock);	 
				}
				//if(DBManager.updateDB(updateDBSqlList)){
				if(this.updDbUsesPreparedStatement(sqlLock,paramList)) {
					errMsg = errMsg + "相關資料寫入資料庫成功";
				}else{
					errMsg = errMsg + "相關資料寫入資料庫失敗";
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
			   List paramList = new ArrayList() ;
			   sqlCmd.append("SELECT * FROM WLX05_M_ATM WHERE bank_no=? AND m_year=? and m_month=?") ;
			   paramList.add(bank_no) ;
			   paramList.add(nyear) ;
			   paramList.add(nmonth) ;
			   List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,push_debitcard_cnt,use_debitcard_cnt,push_bincard_cnt,use_bincard_cnt,atm_cnt,month_tran_cnt,year_acctran_cnt,month_tran_amt,year_acctran_amt,user_id,user_name,update_date,canc_debitcard_cnt,canc_bincard_cnt");
			 	System.out.println("WLX05_M_ATM.size="+data.size());
			   sqlCmd.setLength(0) ;
			   paramList.clear() ;
				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法刪除<br>";
				}else{
				   if(!nyear.equals("") && !nmonth.equals("")){/*寫入log檔中*/
				      sqlCmdLog = "INSERT INTO WLX05_M_ATM_LOG "+
		                        "select m_year,m_month,bank_no,push_debitcard_cnt,"
		                        +"use_debitcard_cnt,push_bincard_cnt,use_bincard_cnt,atm_cnt,"
		                        +"month_tran_cnt,year_acctran_cnt,month_tran_amt,year_acctran_amt,"
		                        +"user_id,user_name,update_date,"
		                        +"?,?,sysdate,'D'"
		                        +",canc_debitcard_cnt,canc_bincard_cnt"
		                        +" from WLX05_M_ATM WHERE bank_no='" 
		                        + bank_no + "' AND m_year=? and m_month=?";
				     paramList.add(user_id) ;
				     paramList.add(user_name) ;
				     paramList.add(nyear) ;
				     paramList.add(nmonth) ;
					 // updateDBSqlList.add(sqlCmdLog);
					 this.updDbUsesPreparedStatement(sqlCmdLog,paramList) ;
					 paramList.clear() ;
				     sqlCmd.append(" delete WLX05_M_ATM where bank_no=? and m_year=? and m_month=? ");
				     paramList.add(bank_no) ;
				     paramList.add(nyear);
				     paramList.add(nmonth) ;
		             //updateDBSqlList.add(sqlCmd);
		             this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList) ;
		             sqlCmd.setLength(0) ;
		             paramList.clear() ;
					 //當寫進資料庫成功後，再寫進WLX_APPLY_LOCK檔中
					 sqlLock = " select * from WLX_APPLY_LOCK aa where aa.M_YEAR =? "+
		                     " and aa.M_QUARTER=? "+
		                     " and aa.BANK_CODE=? "+
		                     " and aa.REPORT_NO='C01'";
		             paramList.add(nyear) ;
		             paramList.add(nmonth);
		             paramList.add(bank_no) ;
	  				 dbLock = DBManager.QueryDB_SQLParam(sqlLock,paramList,"");   
	  				 System.out.print("大小:"+String.valueOf(dbLock.size()));
					 if(dbLock.size()!=0){
			            sqlLockLog = " insert into WLX_APPLY_LOCK_LOG "+
			            			 " select M_YEAR, M_QUARTER, BANK_CODE, REPORT_NO, "+
			            			 " ADD_USER, ADD_NAME, ADD_DATE, LOCK_OWN, LOCK_MGR, "+
			            			 " USER_ID, USER_NAME,UPDATE_DATE, "+
			            		     "? ,?,sysdate,'D'"+
			            		     " from WLX_APPLY_LOCK WHERE BANK_CODE=? AND M_YEAR=?  and M_QUARTER= ?";
			            paramList.clear() ;
			            paramList.add(user_id) ;
			            paramList.add(user_name);
			            paramList.add(bank_no) ;
			            paramList.add(nyear) ;
			            paramList.add(nmonth) ;
		                //updateDBSqlList.add(sqlLockLog);
		                this.updDbUsesPreparedStatement(sqlLockLog,paramList) ;
		                paramList.clear() ;
			            sqlLock = " delete WLX_APPLY_LOCK "+
			            		  " where M_YEAR = ? and M_QUARTER = ? and BANK_CODE =? and REPORT_NO ='C01'";
			            paramList.add(nyear) ;
			            paramList.add(nmonth) ;
			            paramList.add(bank_no) ;
					    //updateDBSqlList.add(sqlLock);
					    this.updDbUsesPreparedStatement(sqlLock,paramList) ;
					}
		           }
				   /*if(DBManager.updateDB(updateDBSqlList)){
				 	  errMsg = errMsg + "相關資料刪除成功";
				   }else{
				 	  errMsg = errMsg + "相關資料刪除失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
				   }*/
				   errMsg = errMsg + "相關資料刪除成功";
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
