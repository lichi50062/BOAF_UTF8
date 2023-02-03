<%
//94.11.07 first designed by 4180
//94.12.01 FIX period_6_fix_rate & period_6_var_rateby 2495
//96.10.03 add 增加牌告利率欄位 by 2295 
//98.03.31 add 基準利率-指標利率(月調)/基準利率(月調)/指數型房貸指標利率(月調) by 2295
//99.12.07 fix sqlInjection by 2808
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
      System.out.println("FX010W login timeout");
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

    if(!CheckPermission(request)){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );
    }else{
    	//set next jsp
    	if(act.equals("new")){
    	   List dbData = getWLX_S_RATE(bank_no,checkyear,checkquarter);  //是否當年季已有資料
           if(dbData.size()==0){
        	 	int temp = Integer.parseInt(checkquarter); //將參數轉換為上季
        	 	temp = temp-1;
        	 	String lastquarter = String.valueOf(temp);
        	 	if(temp==0){
        	 		lastquarter="4";
        	 		checkyear=String.valueOf(Integer.parseInt(checkyear)-1);
        	 	}

        	 	dbData = getWLX_S_RATE(bank_no,checkyear,lastquarter);  //取得上個月資料
        	 	if(dbData.size()==0) //上月若無資料
        	 	   rd = application.getRequestDispatcher( EditPgName +"?act=new");
        	 	else
        	 	   rd = application.getRequestDispatcher( EditPgName +"?act=new&loaddata=ok");
           } else{
        	   dbData = getWLX_S_RATE(bank_no,checkyear,checkquarter);
    	       request.setAttribute("WLX_S_RATE",dbData);
			   request.setAttribute("maintainInfo","select * from WLX_S_RATE WHERE bank_no='" + bank_no+"'");
			   rd = application.getRequestDispatcher( EditPgName +"?act=modify");
        	}
    	}else if(act.equals("Edit")){
    	    List dbData = getWLX_S_RATE(bank_no,myear,mquarter);
    	    request.setAttribute("WLX_S_RATE",dbData);
    	    //93.12.21設定異動者資訊======================================================================
			request.setAttribute("maintainInfo","select * from WLX_S_RATE WHERE bank_no='" + bank_no+"'");
			//=======================================================================================================================
        	rd = application.getRequestDispatcher( EditPgName +"?act=Edit");
    	}else if(act.equals("List")){
    	    List dbData = getWLX_S_RATE(bank_no,"","");
            List inidate = getWLX_S_RATE_INI();
            List lockdate = getWLX_S_RATE_LOCK(bank_no);
    	    request.setAttribute("WLX_S_RATE",dbData);
            request.setAttribute("WLX10_INI",inidate);
            request.setAttribute("WLX10_LOCK",lockdate);
        	rd = application.getRequestDispatcher( ListPgName +"?bank_no="+bank_no);
      }else if(act.equals("Load")){
      	    String syear = ((String)request.getParameter("hyear")==null)?"":(String)request.getParameter("hyear");
	        String squarter= ((String)request.getParameter("hquarter")==null)?"":(String)request.getParameter("hquarter");
        	int temp = Integer.parseInt(squarter); //將參數轉換為上季
          	temp = temp-1;
            String lyear=syear;
        	String lquarter = String.valueOf(temp);
        	if(temp==0){
        	   lquarter="4";
        	   lyear=String.valueOf(Integer.parseInt(syear)-1);
        	 }//處理第一季的情況

        	 List dbData = getWLX_S_RATE(bank_no,lyear,lquarter);  //取得上個月資料
        	 request.setAttribute("WLX_S_RATE",dbData);
        	 rd = application.getRequestDispatcher( EditPgName +"?act=new&loaddata=loaded&S_YEAR="+syear+"&S_QUARTER="+squarter);
    	}else if(act.equals("Insert")){
    	     actMsg = InsertDB(request,bank_no,lguser_id,lguser_name);
        	 rd = application.getRequestDispatcher( nextPgName+"?FX=FX010W" );
    	}else if(act.equals("Update")){
    	     actMsg = UpdateDB(request,bank_no,lguser_id,lguser_name);
        	 rd = application.getRequestDispatcher( nextPgName+"?FX=FX010W");
    	}else if(act.equals("Delete")){
    	     actMsg = DeleteDB(request,bank_no,lguser_id,lguser_name);
        	 rd = application.getRequestDispatcher( nextPgName+"?FX=FX010W" );
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
    private final static String EditPgName = "/pages/FX010W_Edit.jsp";
    private final static String ListPgName = "/pages/FX010W_List.jsp";
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    private boolean CheckPermission(HttpServletRequest request){//檢核權限
    	    boolean CheckOK=false;
    	    HttpSession session = request.getSession();
            Properties permission = ( session.getAttribute("FX010W")==null ) ? new Properties() : (Properties)session.getAttribute("FX010W");
            if(permission == null){
              System.out.println("FX010W.permission == null");
            }else{
               System.out.println("FX010W.permission.size ="+permission.size());

            }
            //只要有Query的權限,就可以進入畫面
        	if(permission != null && permission.get("Q") != null && permission.get("Q").equals("Y")){
        	   CheckOK = true;//Query
        	}
        	return CheckOK;
    }
    //94.04.01 add 主key更改為bank_no+seq_no by 2295
    private List getWLX_S_RATE(String bank_no, String myear, String mquarter){
			List paramList = new ArrayList() ;
    		//程序為顯示畫面，查詢條件
    		String sqlCmd = "select * from WLX_S_RATE where bank_no= ? ";
			paramList.add(bank_no) ;
    		if(!myear.equals("") && !mquarter.equals("")){
    			sqlCmd =sqlCmd + " and m_year=? and m_quarter= ? ";
    			paramList.add(myear) ;
    			paramList.add(mquarter) ;
    		}
    		sqlCmd =sqlCmd + " order by M_YEAR desc, M_QUARTER desc";

            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"m_year,m_quarter,period_1_fix_rate,period_1_var_rate,"+
            								 	   "period_3_fix_rate,period_3_var_rate,"+
                                                   "period_6_fix_rate,period_6_var_rate,"+
                                                   "period_9_fix_rate,period_9_var_rate,"+
                                                   "period_12_fix_rate,period_12_var_rate,"+
                                                   "period_24_fix_rate,period_24_var_rate,"+
                                                   "period_36_fix_rate,period_36_var_rate,"+
                                                   "deposit_12_fix_rate,deposit_12_var_rate,"+
                                                   "deposit_24_fix_rate,deposit_24_var_rate,"+
                                                   "deposit_36_fix_rate,deposit_36_var_rate,"+
                                                   "deposit_var_rate,save_var_rate,"+
                                                   "basic_pay_var_rate,period_house_var_rate,"+
                                                   "base_mark_rate,base_fix_rate,"+
                                                   "base_base_rate,"+
                                                   "base_mark_rate_month,base_base_rate_month,period_house_var_rate_month,"+ 
                                                   "user_id,user_name,update_date");
            paramList.clear()  ;
            return dbData;
    }
    private List getWLX_S_RATE_INI(){
        List paramList = new ArrayList();
      	String sqlCmd = "select * from WLX_APPLY_INI where REPORT_NO=?";
      	paramList.add("C07");
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"m_year,m_month");

        if (dbData.size()!= 0)
            return dbData;
        else
           return null;
    }
    private List getWLX_S_RATE_LOCK(String bank_no){
    	    List paramList = new ArrayList () ;
      		String sqlCmd = " select M_YEAR,M_QUARTER from WLX_APPLY_LOCK "
 		       			  + " where BANK_CODE = ? "
 		       			  + " and REPORT_NO = 'C07'"          
 		       			  + " and((LOCK_OWN   =  'Y') or (LOCK_MGR = 'Y'))"; 		       				
 			paramList.add(bank_no) ;
        	List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"m_year,m_quarter");

      		if (dbData.size()!= 0)
          		return dbData;
      		else
          		return null;
    }
    public String InsertDB(HttpServletRequest request,String bank_no,String lguser_id,String lguser_name) throws Exception{
  			StringBuffer sqlCmd = new StringBuffer() ;
			String errMsg="";

			String nyear = ((String)request.getParameter("hyear")==null)?"":(String)request.getParameter("hyear");
			String nquarter= ((String)request.getParameter("hquarter")==null)?"":(String)request.getParameter("hquarter"); 	
	        String period_1_fix_rate = ((String)request.getParameter("period_1_fix_rate")==null)?"":(String)request.getParameter("period_1_fix_rate");
			String period_1_var_rate = ((String)request.getParameter("period_1_var_rate")==null)?"":(String)request.getParameter("period_1_var_rate");
			String period_3_fix_rate = ((String)request.getParameter("period_3_fix_rate")==null)?"":(String)request.getParameter("period_3_fix_rate");
			String period_3_var_rate = ((String)request.getParameter("period_3_var_rate")==null)?"":(String)request.getParameter("period_3_var_rate");
			String period_6_fix_rate = ((String)request.getParameter("period_6_fix_rate")==null)?"":(String)request.getParameter("period_6_fix_rate");
			String period_6_var_rate = ((String)request.getParameter("period_6_var_rate")==null)?"":(String)request.getParameter("period_6_var_rate");
			String period_9_fix_rate = ((String)request.getParameter("period_9_fix_rate")==null)?"":(String)request.getParameter("period_9_fix_rate");
			String period_9_var_rate = ((String)request.getParameter("period_9_var_rate")==null)?"":(String)request.getParameter("period_9_var_rate");
			String period_12_fix_rate = ((String)request.getParameter("period_12_fix_rate")==null)?"":(String)request.getParameter("period_12_fix_rate");
			String period_12_var_rate = ((String)request.getParameter("period_12_var_rate")==null)?"":(String)request.getParameter("period_12_var_rate");
			//96.10.03增加的利率==================================================================================
			String period_24_fix_rate = ((String)request.getParameter("period_24_fix_rate")==null)?"":(String)request.getParameter("period_24_fix_rate");
			String period_24_var_rate = ((String)request.getParameter("period_24_var_rate")==null)?"":(String)request.getParameter("period_24_var_rate");
			String period_36_fix_rate = ((String)request.getParameter("period_36_fix_rate")==null)?"":(String)request.getParameter("period_36_fix_rate");
			String period_36_var_rate = ((String)request.getParameter("period_36_var_rate")==null)?"":(String)request.getParameter("period_36_var_rate");
			String deposit_12_fix_rate = ((String)request.getParameter("deposit_12_fix_rate")==null)?"":(String)request.getParameter("deposit_12_fix_rate");
			String deposit_12_var_rate = ((String)request.getParameter("deposit_12_var_rate")==null)?"":(String)request.getParameter("deposit_12_var_rate");
			String deposit_24_fix_rate = ((String)request.getParameter("deposit_24_fix_rate")==null)?"":(String)request.getParameter("deposit_24_fix_rate");
			String deposit_24_var_rate = ((String)request.getParameter("deposit_24_var_rate")==null)?"":(String)request.getParameter("deposit_24_var_rate");
			String deposit_36_fix_rate = ((String)request.getParameter("deposit_36_fix_rate")==null)?"":(String)request.getParameter("deposit_36_fix_rate");
			String deposit_36_var_rate = ((String)request.getParameter("deposit_36_var_rate")==null)?"":(String)request.getParameter("deposit_36_var_rate");
			String deposit_var_rate = ((String)request.getParameter("deposit_var_rate")==null)?"":(String)request.getParameter("deposit_var_rate");
			String save_var_rate = ((String)request.getParameter("save_var_rate")==null)?"":(String)request.getParameter("save_var_rate");
			//98.03.31增加的利率==================================================================================
			String base_mark_rate_month = ((String)request.getParameter("base_mark_rate_month")==null)?"":(String)request.getParameter("base_mark_rate_month");
			String base_base_rate_month = ((String)request.getParameter("base_base_rate_month")==null)?"":(String)request.getParameter("base_base_rate_month");
			String period_house_var_rate_month = ((String)request.getParameter("period_house_var_rate_month")==null)?"":(String)request.getParameter("period_house_var_rate_month");		
			//===================================================================================================
			String basic_pay_var_rate = ((String)request.getParameter("basic_pay_var_rate")==null)?"":(String)request.getParameter("basic_pay_var_rate");			
			String period_house_var_rate = ((String)request.getParameter("period_house_var_rate")==null)?"":(String)request.getParameter("period_house_var_rate");
			String base_mark_rate = ((String)request.getParameter("base_mark_rate")==null)?"":(String)request.getParameter("base_mark_rate");
			String base_fix_rate = ((String)request.getParameter("base_fix_rate")==null)?"":(String)request.getParameter("base_fix_rate");
			String base_base_rate = ((String)request.getParameter("base_base_rate")==null)?"":(String)request.getParameter("base_base_rate");
			String user_id=lguser_id;
			String user_name=lguser_name;
			String sqlLock = "";
			String sqlLockLog = "";
			List dbLock;		
      		//List updateDBSqlList = new LinkedList();
			List paramList = new ArrayList();
			try {   		
				paramList.add( nyear );
				paramList.add( nquarter );
				paramList.add( bank_no );
        		paramList.add( period_1_fix_rate);
				paramList.add( period_1_var_rate );
				paramList.add( period_3_fix_rate);
				paramList.add( period_3_var_rate);
				paramList.add( period_6_fix_rate);
				paramList.add( period_6_var_rate);
				paramList.add( period_9_fix_rate);
				paramList.add( period_9_var_rate);
				paramList.add( period_12_fix_rate);
				paramList.add( period_12_var_rate);
				paramList.add( basic_pay_var_rate);
				paramList.add( period_house_var_rate);
				paramList.add( base_mark_rate);
				paramList.add( base_fix_rate);
				paramList.add( base_base_rate);
				paramList.add( lguser_id);
				paramList.add( lguser_name);
				paramList.add( period_24_fix_rate);
				paramList.add( period_24_var_rate);
				paramList.add( period_36_fix_rate);
				paramList.add( period_36_var_rate);
				paramList.add( deposit_12_fix_rate);
				paramList.add( deposit_12_var_rate);
				paramList.add( deposit_24_fix_rate);
				paramList.add( deposit_24_var_rate);
				paramList.add( deposit_36_fix_rate);
				paramList.add( deposit_36_var_rate);
				paramList.add( deposit_var_rate);
				paramList.add( save_var_rate ) ;
				paramList.add( base_mark_rate_month ) ;
				paramList.add( base_base_rate_month );
				paramList.add( period_house_var_rate_month ) ;
				sqlCmd.append("INSERT INTO WLX_S_RATE VALUES('" ) ;
				sqlCmd.append(" ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,sysdate,");
				sqlCmd.append("?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
				/*		    + nyear + "','"
							+ nquarter +"','"
							+ bank_no +"','"
                            + period_1_fix_rate+ "','"
							+ period_1_var_rate+ "','"
							+ period_3_fix_rate+ "','"
							+ period_3_var_rate+ "','"
							+ period_6_fix_rate+ "','"
							+ period_6_var_rate+ "','"
							+ period_9_fix_rate+ "','"
							+ period_9_var_rate+ "','"
							+ period_12_fix_rate+ "','"
							+ period_12_var_rate+ "','"
							+ basic_pay_var_rate+ "','"
							+ period_house_var_rate+ "','"
							+ base_mark_rate+ "','"
							+ base_fix_rate+ "','"
							+ base_base_rate+ "','"
							+ lguser_id +"','"
							+ lguser_name
							+ "',sysdate,'"
							+ period_24_fix_rate+ "','" //96.10.03 add 增加的利率
							+ period_24_var_rate+ "','"
							+ period_36_fix_rate+ "','"
							+ period_36_var_rate+ "','"
							+ deposit_12_fix_rate+ "','"
							+ deposit_12_var_rate+ "','"
							+ deposit_24_fix_rate+ "','"
							+ deposit_24_var_rate+ "','"
							+ deposit_36_fix_rate+ "','"
							+ deposit_36_var_rate+ "','"
							+ deposit_var_rate+ "','"
							+ save_var_rate+ "','"
							+ base_mark_rate_month+ "','"
							+ base_base_rate_month+ "','"
							+ period_house_var_rate_month+ "'"//98.03.31 add 增加的利率
							+")");

                	updateDBSqlList.add(sqlCmd);*/
                	this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList) ;
                	paramList.clear() ;
                	sqlCmd.setLength(0) ;
                	//當寫進資料庫成功後，再寫進WLX_APPLY_LOCK檔中
					sqlLock = " select * from WLX_APPLY_LOCK aa where aa.M_YEAR =? "+
		         			  " and aa.M_QUARTER=?"+
		          			  " and aa.BANK_CODE=?"+
		          			  " and aa.REPORT_NO='C07'";
		          	paramList.add(nyear) ;
		          	paramList.add(nquarter) ;
		          	paramList.add(bank_no) ;
	  				dbLock = DBManager.QueryDB_SQLParam(sqlLock,paramList,"");   
	  				paramList.clear();
	  				System.out.print("大小:"+String.valueOf(dbLock.size()));
	  
					if(dbLock.size()!=0){            
						sqlLock = " update WLX_APPLY_LOCK set USER_ID = ?"+
							    " ,USER_NAME =?"+
					  		    " ,UPDATE_DATE =sysdate"+
					  			" where M_YEAR = ? and M_QUARTER = ? and BANK_CODE =?"+
					  			" and REPORT_NO ='C07'";
					   paramList.add(user_id) ;
					   paramList.add(user_name) ;
					   paramList.add(nyear);
					   paramList.add(nquarter) ;
					   paramList.add(bank_no) ;
					   // updateDBSqlList.add(sqlLock);  
					}else{  //若是原本沒有鎖定資料的情況下，寫進一筆新的鎖定紀錄
						sqlLock = " insert into WLX_APPLY_LOCK values(?,?,?,'C07',?,?,sysdate,'N','N',?,?,sysdate)"; 
					  paramList.add(nyear) ;
					  paramList.add(nquarter) ;
					  paramList.add(bank_no) ;
					  paramList.add(user_id) ;
					  paramList.add(user_name) ;
					  paramList.add(user_id);
					  paramList.add(user_name) ;
					  //updateDBSqlList.add(sqlLock);		  
					}		
                
                
					//if(DBManager.updateDB(updateDBSqlList)){
					if(this.updDbUsesPreparedStatement(sqlLock,paramList)){
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
			String sqlLock = "";
			String sqlLockLog = "";
			String errMsg="";
			String nyear = ((String)request.getParameter("hyear")==null)?"":(String)request.getParameter("hyear");
			String nquarter= ((String)request.getParameter("hquarter")==null)?"":(String)request.getParameter("hquarter"); 	
			String period_1_fix_rate = ((String)request.getParameter("period_1_fix_rate")==null)?"":(String)request.getParameter("period_1_fix_rate");
			String period_1_var_rate = ((String)request.getParameter("period_1_var_rate")==null)?"":(String)request.getParameter("period_1_var_rate");
			String period_3_fix_rate = ((String)request.getParameter("period_3_fix_rate")==null)?"":(String)request.getParameter("period_3_fix_rate");
			String period_3_var_rate = ((String)request.getParameter("period_3_var_rate")==null)?"":(String)request.getParameter("period_3_var_rate");
			String period_6_fix_rate = ((String)request.getParameter("period_6_fix_rate")==null)?"":(String)request.getParameter("period_6_fix_rate");
			String period_6_var_rate = ((String)request.getParameter("period_6_var_rate")==null)?"":(String)request.getParameter("period_6_var_rate");
			String period_9_fix_rate = ((String)request.getParameter("period_9_fix_rate")==null)?"":(String)request.getParameter("period_9_fix_rate");
			String period_9_var_rate = ((String)request.getParameter("period_9_var_rate")==null)?"":(String)request.getParameter("period_9_var_rate");
			String period_12_fix_rate = ((String)request.getParameter("period_12_fix_rate")==null)?"":(String)request.getParameter("period_12_fix_rate");
			String period_12_var_rate = ((String)request.getParameter("period_12_var_rate")==null)?"":(String)request.getParameter("period_12_var_rate");
			//96.10.03增加的利率==================================================================================
			String period_24_fix_rate = ((String)request.getParameter("period_24_fix_rate")==null)?"":(String)request.getParameter("period_24_fix_rate");
			String period_24_var_rate = ((String)request.getParameter("period_24_var_rate")==null)?"":(String)request.getParameter("period_24_var_rate");
			String period_36_fix_rate = ((String)request.getParameter("period_36_fix_rate")==null)?"":(String)request.getParameter("period_36_fix_rate");
			String period_36_var_rate = ((String)request.getParameter("period_36_var_rate")==null)?"":(String)request.getParameter("period_36_var_rate");
			String deposit_12_fix_rate = ((String)request.getParameter("deposit_12_fix_rate")==null)?"":(String)request.getParameter("deposit_12_fix_rate");
			String deposit_12_var_rate = ((String)request.getParameter("deposit_12_var_rate")==null)?"":(String)request.getParameter("deposit_12_var_rate");
			String deposit_24_fix_rate = ((String)request.getParameter("deposit_24_fix_rate")==null)?"":(String)request.getParameter("deposit_24_fix_rate");
			String deposit_24_var_rate = ((String)request.getParameter("deposit_24_var_rate")==null)?"":(String)request.getParameter("deposit_24_var_rate");
			String deposit_36_fix_rate = ((String)request.getParameter("deposit_36_fix_rate")==null)?"":(String)request.getParameter("deposit_36_fix_rate");
			String deposit_36_var_rate = ((String)request.getParameter("deposit_36_var_rate")==null)?"":(String)request.getParameter("deposit_36_var_rate");
			String deposit_var_rate = ((String)request.getParameter("deposit_var_rate")==null)?"":(String)request.getParameter("deposit_var_rate");
			String save_var_rate = ((String)request.getParameter("save_var_rate")==null)?"":(String)request.getParameter("save_var_rate");
			//98.03.31增加的利率==================================================================================
			String base_mark_rate_month = ((String)request.getParameter("base_mark_rate_month")==null)?"":(String)request.getParameter("base_mark_rate_month");
			String base_base_rate_month = ((String)request.getParameter("base_base_rate_month")==null)?"":(String)request.getParameter("base_base_rate_month");
			String period_house_var_rate_month = ((String)request.getParameter("period_house_var_rate_month")==null)?"":(String)request.getParameter("period_house_var_rate_month");		
			//====================================================================================================
			String basic_pay_var_rate = ((String)request.getParameter("basic_pay_var_rate")==null)?"":(String)request.getParameter("basic_pay_var_rate");
			String period_house_var_rate = ((String)request.getParameter("period_house_var_rate")==null)?"":(String)request.getParameter("period_house_var_rate");
			String base_mark_rate = ((String)request.getParameter("base_mark_rate")==null)?"":(String)request.getParameter("base_mark_rate");
			String base_fix_rate = ((String)request.getParameter("base_fix_rate")==null)?"":(String)request.getParameter("base_fix_rate");
			String base_base_rate = ((String)request.getParameter("base_base_rate")==null)?"":(String)request.getParameter("base_base_rate");
			String user_id=lguser_id;
			String user_name=lguser_name;
			List dbLock;
			System.out.print("period_6_fix_rate:"+period_6_fix_rate);
      		System.out.print("period_6_var_rate:"+period_6_var_rate);
			//List updateDBSqlList = new LinkedList();
			List paramList = new ArrayList() ;
			try {
				
			    sqlCmd.append("SELECT * FROM WLX_S_RATE WHERE bank_no=? AND m_year=? and m_quarter= ? ");
			    paramList.add(bank_no) ;
			    paramList.add(nyear) ;
			    paramList.add(nquarter) ;
			  	List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
				System.out.println("WLX_S_RATE.size="+data.size());
				sqlCmd.setLength(0) ;
				paramList.clear() ;
				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
				}else{
					/*寫入log檔中*/
					sqlCmdLog = "INSERT INTO WLX_S_RATE_LOG "
		            			+"select m_year,m_quarter,bank_no,"
		            			+"period_1_fix_rate,period_1_var_rate,"
		            			+"period_3_fix_rate,period_3_var_rate,"
		            			+"period_6_fix_rate,period_6_var_rate,"
		            			+"period_9_fix_rate,period_9_var_rate,"
		            			+"period_12_fix_rate,period_12_var_rate,"
		            			+"basic_pay_var_rate,period_house_var_rate,"
		            			+"base_mark_rate,base_fix_rate,base_base_rate,"
		            			+"user_id,user_name,update_date,"
		            			+"?,?,sysdate,'U',"
		            			+"period_24_fix_rate,period_24_var_rate," //96.10.03 add 增加的利率
							    +"period_36_fix_rate,period_36_var_rate,"
							    +"deposit_12_fix_rate,deposit_12_var_rate,"
							    +"deposit_24_fix_rate,deposit_24_var_rate,"
								+"deposit_36_fix_rate,deposit_36_var_rate,"
								+"deposit_var_rate,save_var_rate,"
								+"base_mark_rate_month,base_base_rate_month,period_house_var_rate_month "//98.03.31 add 增加的利率
		            			+" from WLX_S_RATE WHERE bank_no=? AND m_year=? and m_quarter= ?";
					paramList.add(user_id) ;
					paramList.add(user_name) ;
					paramList.add(bank_no) ;
					paramList.add(nyear) ;
					paramList.add(nquarter) ;
					//updateDBSqlList.add(sqlCmdLog);
					this.updDbUsesPreparedStatement(sqlCmdLog,paramList) ;
					paramList.clear()  ;
    				//94/12/01 FIX period_6_fix_rate & period_6_var_rateby 2495
					sqlCmd.append("UPDATE WLX_S_RATE SET ") ;
					sqlCmd.append("m_year=?,m_quarter=? ") ;
				    paramList.add(nyear) ;
					paramList.add(nquarter) ;
			        sqlCmd.append(",bank_no=?");paramList.add(bank_no ) ;
			     	sqlCmd.append(",period_1_fix_rate=?");paramList.add(period_1_fix_rate ) ;
			     	sqlCmd.append(",period_1_var_rate=?");paramList.add(period_1_var_rate ) ;
		         	sqlCmd.append(",period_3_fix_rate=?");paramList.add(period_3_fix_rate ) ;
		         	sqlCmd.append(",period_3_var_rate=?");paramList.add(period_3_var_rate ) ;
		         	sqlCmd.append(",period_6_fix_rate=?");paramList.add(period_6_fix_rate ) ;
		         	sqlCmd.append(",period_6_var_rate=?");paramList.add(period_6_var_rate ) ;
		         	sqlCmd.append(",period_9_fix_rate=?");paramList.add(period_9_fix_rate ) ;
		         	sqlCmd.append(",period_9_var_rate=?");paramList.add(period_9_var_rate ) ;
		         	sqlCmd.append(",period_12_fix_rate=?");paramList.add(period_12_fix_rate ) ;
		         	sqlCmd.append(",period_12_var_rate=?");paramList.add(period_12_var_rate ) ;
		         	sqlCmd.append(",period_24_fix_rate=?");paramList.add(period_24_fix_rate);
		         	sqlCmd.append(",period_24_var_rate=?");paramList.add(period_24_var_rate );
		         	sqlCmd.append(",period_36_fix_rate=?");paramList.add(period_36_fix_rate ) ;
		         	sqlCmd.append(",period_36_var_rate=?");paramList.add(period_36_var_rate ) ;
		         	sqlCmd.append(",deposit_12_fix_rate=?");paramList.add(deposit_12_fix_rate ) ;
		         	sqlCmd.append(",deposit_12_var_rate=?");paramList.add(deposit_12_var_rate ) ;
		         	sqlCmd.append(",deposit_24_fix_rate=?");paramList.add(deposit_24_fix_rate ) ;
		         	sqlCmd.append(",deposit_24_var_rate=?");paramList.add(deposit_24_var_rate );
		         	sqlCmd.append(",deposit_36_fix_rate=?");paramList.add(deposit_36_fix_rate );
		         	sqlCmd.append(",deposit_36_var_rate=?");paramList.add(deposit_36_var_rate );
		         	sqlCmd.append(",deposit_var_rate=?");paramList.add(deposit_var_rate ) ;
		         	sqlCmd.append(",save_var_rate=?");paramList.add(save_var_rate ) ;
		         	sqlCmd.append(",basic_pay_var_rate=?");paramList.add(basic_pay_var_rate ) ;
		         	sqlCmd.append(",period_house_var_rate=?");paramList.add(period_house_var_rate ) ;
		         	sqlCmd.append(",base_mark_rate=?");paramList.add(base_mark_rate ) ;
		         	sqlCmd.append(",base_fix_rate=?");paramList.add(base_fix_rate ) ;
		         	sqlCmd.append(",base_base_rate=?");paramList.add(base_base_rate ) ;	
		         	sqlCmd.append(",base_mark_rate_month=?");paramList.add(base_mark_rate_month); //98.03.31 add 增加的利率
		         	sqlCmd.append(",base_base_rate_month=?");paramList.add(base_base_rate_month ) ;
		         	sqlCmd.append(",period_house_var_rate_month=?");paramList.add(period_house_var_rate_month ) ;		
   	   	   		 	sqlCmd.append(",user_id=?");paramList.add(user_id ) ;
   	   	   		 	sqlCmd.append(",user_name=?");paramList.add(user_name ) ;
   	   	   		 	sqlCmd.append(",update_date=sysdate") ;
   	   	   		 	sqlCmd.append(" where bank_no=? and m_year=? and m_quarter=?" );
   	   	   		 	paramList.add(bank_no) ;
   	   	   		 	paramList.add(nyear) ;
   	   	   		 	paramList.add(nquarter) ;
						/*	 +"m_year='"+nyear+"',m_quarter='"+nquarter+"'"
							 +",bank_no='"+bank_no+"'"
							 +",period_1_fix_rate='"+period_1_fix_rate+"'"
							 +",period_1_var_rate='"+period_1_var_rate+"'"
					         +",period_3_fix_rate='"+period_3_fix_rate+"'"
					         +",period_3_var_rate='"+period_3_var_rate+"'"
					         +",period_6_fix_rate='"+period_6_fix_rate+"'"
					         +",period_6_var_rate='"+period_6_var_rate+"'"
					         +",period_9_fix_rate='"+period_9_fix_rate+"'"
					         +",period_9_var_rate='"+period_9_var_rate+"'"
					         +",period_12_fix_rate='"+period_12_fix_rate+"'"
					         +",period_12_var_rate='"+period_12_var_rate+"'"
					         +",period_24_fix_rate='"+period_24_fix_rate+"'"//96.10.03 add 增加的利率
					         +",period_24_var_rate='"+period_24_var_rate+"'"
					         +",period_36_fix_rate='"+period_36_fix_rate+"'"
					         +",period_36_var_rate='"+period_36_var_rate+"'"
					         +",deposit_12_fix_rate='"+deposit_12_fix_rate+"'"
					         +",deposit_12_var_rate='"+deposit_12_var_rate+"'"
					         +",deposit_24_fix_rate='"+deposit_24_fix_rate+"'"
					         +",deposit_24_var_rate='"+deposit_24_var_rate+"'"
					         +",deposit_36_fix_rate='"+deposit_36_fix_rate+"'"
					         +",deposit_36_var_rate='"+deposit_36_var_rate+"'"
					         +",deposit_var_rate='"+deposit_var_rate+"'"
					         +",save_var_rate='"+save_var_rate+"'"
					         +",basic_pay_var_rate='"+basic_pay_var_rate+"'"
					         +",period_house_var_rate='"+period_house_var_rate+"'"
					         +",base_mark_rate='"+base_mark_rate+"'"
					         +",base_fix_rate='"+base_fix_rate+"'"
					         +",base_base_rate='"+base_base_rate+"'"	
					         +",base_mark_rate_month='"+base_mark_rate_month+"'"//98.03.31 add 增加的利率
					         +",base_base_rate_month='"+base_base_rate_month+"'"
					         +",period_house_var_rate_month='"+period_house_var_rate_month+"'"		
   	   	   	   			     +",user_id='"+user_id+"'"
   	   	   	   			     +",user_name='"+user_name+"'"
   	   	   	   			     +",update_date=sysdate"
   	   	   	   			     +" where bank_no='"+bank_no+"' and m_year='"+nyear+"' and m_quarter='"+nquarter+"'";

		            updateDBSqlList.add(sqlCmd);*/
					this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList) ;
		            sqlCmd.setLength(0) ;
		            paramList.clear() ;
 					//當寫進資料庫成功後，再寫進WLX_APPLY_LOCK檔中
					sqlLock = " select * from WLX_APPLY_LOCK aa where aa.M_YEAR =?"+
		         			 " and aa.M_QUARTER=?"+
		         			 " and aa.BANK_CODE=?"+
		         			 " and aa.REPORT_NO='C07'";
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
					     			"?,?,sysdate,'U'"+
					     			" from WLX_APPLY_LOCK WHERE BANK_CODE=? AND M_YEAR=? and M_QUARTER=?";
		                 paramList.add(user_id) ;
		                 paramList.add(user_name) ;
		                 paramList.add(bank_no) ;
		                 paramList.add(nyear) ;
		                 paramList.add(nquarter) ;
		                 //updateDBSqlList.add(sqlLockLog);
		                 this.updDbUsesPreparedStatement(sqlLockLog,paramList) ;
		                 paramList.clear() ;
						 sqlLock = " update WLX_APPLY_LOCK set USER_ID = ?"+
					  			" ,USER_NAME =?"+
					  			" ,UPDATE_DATE =sysdate"+
					  			" where M_YEAR = ? and M_QUARTER = ? and BANK_CODE =?"+
					  			" and REPORT_NO ='C07'";
						 paramList.add(user_id) ;
		                 paramList.add(user_name) ;
		                 paramList.add(nyear) ;
		                 paramList.add(nquarter) ;
		                 paramList.add(bank_no) ;
					  	 //updateDBSqlList.add_SQLP(sqlLock);
				  
				 	}else{  //若是原本沒有鎖定資料的情況下，寫進一筆新的鎖定紀錄
						/*sqlLock = " insert into WLX_APPLY_LOCK values('"
		              			+nyear+"','"
		              			+nquarter+"','"
		              			+bank_no+"','C07','"
		              			+user_id+"','"
		              			+user_name+"',sysdate,'N','N','"
		              			+user_id+"','"
		              			+user_name+"',sysdate)";*/
		              	 sqlLock = "insert into WLX_APPLY_LOCK Values( " ;
		              	 sqlLock += "?,?,?,'C07',?,?,sysdate,'N','N',?,?,sysdate)" ;
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
	 	String nquarter= ((String)request.getParameter("hquarter")==null)?"":(String)request.getParameter("hquarter");
		String user_id=lguser_id;
	    String user_name=lguser_name;
		String sqlLock = "";
		String sqlLockLog = "";
		List dbLock;
		//List updateDBSqlList = new LinkedList();	
		List paramList =new ArrayList() ;
		try {
			   sqlCmd.append("SELECT * FROM WLX_S_RATE WHERE bank_no=? AND m_year=? and m_quarter=?") ;
			   paramList.add(bank_no) ;
			   paramList.add(nyear) ;
			   paramList.add(nquarter) ;
			   List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
			   //System.out.println("WLX_S_RATE.size="+data.size());
			   sqlCmd.setLength(0) ;
			   paramList.clear() ;
			   if (data.size() == 0){
				   errMsg = errMsg + "此筆資料不存在無法刪除<br>";
			   }else{
				   if(!nyear.equals("") && !nquarter.equals("")){
				 	  /*寫入log檔中*/
					  sqlCmdLog = "INSERT INTO WLX_S_RATE_LOG "
		                        +"select m_year,m_quarter,bank_no,"
		                        +"period_1_fix_rate,period_1_var_rate,"
		                        +"period_3_fix_rate,period_3_var_rate,"
		                        +"period_6_fix_rate,period_6_var_rate,"
		                        +"period_9_fix_rate,period_9_var_rate,"
		                        +"period_12_fix_rate,period_12_var_rate,"
		                        +"basic_pay_var_rate,period_house_var_rate,"
		                        +"base_mark_rate,base_fix_rate,base_base_rate,"
		                        +"user_id,user_name,update_date,"
		                        +"?,?,sysdate,'D',"
		                        +"period_24_fix_rate,period_24_var_rate," //96.10.03 add 增加的利率
							    +"period_36_fix_rate,period_36_var_rate,"
							    +"deposit_12_fix_rate,deposit_12_var_rate,"
							    +"deposit_24_fix_rate,deposit_24_var_rate,"
								+"deposit_36_fix_rate,deposit_36_var_rate,"
								+"deposit_var_rate,save_var_rate,"
								+"base_mark_rate_month,base_base_rate_month,period_house_var_rate_month "//98.03.31 add 增加的利率
		                        +" from WLX_S_RATE WHERE bank_no=? AND m_year=? and m_quarter= ?";
				 	  paramList.add(user_id) ;
				 	  paramList.add(user_name) ;
				 	  paramList.add(bank_no) ;
				 	  paramList.add(nyear) ;
				 	  paramList.add(nquarter) ;
				 	  this.updDbUsesPreparedStatement(sqlCmdLog,paramList) ;
				 	  paramList.clear() ;
					  //updateDBSqlList.add(sqlCmdLog);	
		
				      sqlCmd.append(" delete WLX_S_RATE where bank_no=? and m_year=? and m_quarter=?" );
				      paramList.add(bank_no) ;
				      paramList.add(nyear) ;
				      paramList.add(nquarter) ;
				      this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList) ;
				      paramList.clear() ;
				      sqlCmd.setLength(0) ;
		              //updateDBSqlList.add(sqlCmd);
		            
 					  //當寫進資料庫成功後，再寫進WLX_APPLY_LOCK檔中
					  sqlLock = " select * from WLX_APPLY_LOCK aa where aa.M_YEAR =?"+
		          			  " and aa.M_QUARTER=?"+
		          			  " and aa.BANK_CODE=?"+
		          			  " and aa.REPORT_NO='C07'";
		          	  paramList.add(nyear) ;
		          	  paramList.add(nquarter);
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
					     			" from WLX_APPLY_LOCK WHERE BANK_CODE=? AND M_YEAR=? and M_QUARTER= ? ";
		                 paramList.add(user_id) ;
		                 paramList.add(user_name) ;
		                 paramList.add(bank_no) ;
		                 paramList.add(nyear) ;
		                 paramList.add(nquarter) ;
		                 //updateDBSqlList.add(sqlLockLog);
		                 this.updDbUsesPreparedStatement(sqlLockLog,paramList ) ;
		                 paramList.clear() ;
						 sqlLock = " delete WLX_APPLY_LOCK "+
					  			 " where M_YEAR = ? and M_QUARTER = ? and BANK_CODE =?"+
					 	         " and REPORT_NO ='C07'";
					  	  paramList.add(nyear) ;
					  	  paramList.add(nquarter);
					  	  paramList.add(bank_no) ;
					  	 //updateDBSqlList.add(sqlLock);
					  	 this.updDbUsesPreparedStatement(sqlLock,paramList) ;
					  }//end of dbLock.size != 0
		         }
				 errMsg = errMsg + "相關資料刪除成功";
				/* if(DBManager.updateDB(updateDBSqlList)){
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
