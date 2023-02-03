<%
//105.10.03 create by2968
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
      System.out.println("TM006W login timeout");
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
	    	if(act.equals("Qry")){
	    		System.out.println("Qry.............");
	    	    List dbData = getTM006W_List(bank_no);
	    	    System.out.println("dbData="+dbData.size());
	    	    request.setAttribute("dbData",dbData);
	        	rd = application.getRequestDispatcher( ListPgName +"?bank_no="+bank_no);
	    	}else if(act.equals("ApplyList")){
		    		System.out.println("ApplyList.............");
		    		String acc_Tr_Type = ( request.getParameter("acc_Tr_Type")==null ) ? "" : (String)request.getParameter("acc_Tr_Type");
		    	    List dbData = getTM006W_ApplyList(bank_no,acc_Tr_Type,"");
		    	    System.out.println("dbData="+dbData.size());
		    	    request.setAttribute("dbData",dbData);
		        	rd = application.getRequestDispatcher( ApplyListPgName +"?bank_no="+bank_no+"&acc_Tr_Type="+acc_Tr_Type);
	    	}else if(act.equals("new")||act.equals("Edit")){
	    		String acc_Tr_Type = ( request.getParameter("acc_Tr_Type")==null ) ? "" : (String)request.getParameter("acc_Tr_Type");
	    		String applyDate = ( request.getParameter("applyDate")==null ) ? "" : (String)request.getParameter("applyDate");
	    		request.setAttribute("accDivList01",getAccDivList(acc_Tr_Type,"01"));
	    		request.setAttribute("accDivList02",getAccDivList(acc_Tr_Type,"02"));
	    	    List dbData = getEditInfo(acc_Tr_Type,bank_no,applyDate);
	    	    List EditData = getTM006W_ApplyList(bank_no,acc_Tr_Type,applyDate);
	    	    List dbData_Pre = getPreEditInfo(acc_Tr_Type,bank_no,applyDate);
	    	    request.setAttribute("dbData",dbData);
	    	    request.setAttribute("EditData",EditData);
	    	    request.setAttribute("dbData_Pre",dbData_Pre);
	    		rd = application.getRequestDispatcher( EditPgName +"?act="+act);
	    	}else if(act.equals("Update")){
	    		String acc_Tr_Type = ( request.getParameter("acc_Tr_Type")==null ) ? "" : (String)request.getParameter("acc_Tr_Type");
	    	    actMsg = UpdateDB(request,bank_no,lguser_id,lguser_name);
	        	rd = application.getRequestDispatcher( nextPgName+"?goPages=TM006W.jsp&act=ApplyList&bank_no="+bank_no+"&acc_Tr_Type="+acc_Tr_Type);
	    	}else if(act.equals("Delete")){
	    		String acc_Tr_Type = ( request.getParameter("acc_Tr_Type")==null ) ? "" : (String)request.getParameter("acc_Tr_Type");
	    	    actMsg = DeleteDB(request,bank_no,lguser_id,lguser_name);
	        	rd = application.getRequestDispatcher( nextPgName+"?goPages=TM006W.jsp&act=ApplyList&bank_no="+bank_no+"&acc_Tr_Type="+acc_Tr_Type);
	    	}
	    	request.setAttribute("actMsg",actMsg);
	    }

		try {
	        //forward to next present jsp
	        rd.forward(request, response);
	    } catch (NullPointerException npe) {
	    }
    }//end of doProcess
%>


<%!
    private final static String nextPgName = "/pages/ActMsg.jsp";
    private final static String EditPgName = "/pages/TM006W_Edit.jsp";
    private final static String ListPgName = "/pages/TM006W_List.jsp";
    private final static String ApplyListPgName = "/pages/TM006W_ApplyList.jsp";
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    private boolean CheckPermission(HttpServletRequest request){//檢核權限
    	    boolean CheckOK=false;
    	    HttpSession session = request.getSession();
            Properties permission = ( session.getAttribute("TM006W")==null ) ? new Properties() : (Properties)session.getAttribute("TM006W");
            if(permission == null){
              System.out.println("TM006W.permission == null");
            }else{
               System.out.println("TM006W.permission.size ="+permission.size());

            }
            //只要有Query的權限,就可以進入畫面
        	if(permission != null && permission.get("Q") != null && permission.get("Q").equals("Y")){
        	   CheckOK = true;//Query
        	}
        	return CheckOK;
    }
    
    private List getTM006W_List(String bank_no){
    	List paramList = new ArrayList() ;
    	String sqlCmd = "select loanapply_bn01.acc_tr_type,acc_tr_name ";
    	sqlCmd += " from loanapply_ncacno left join loanapply_bn01 on loanapply_ncacno.acc_tr_type = loanapply_bn01.acc_tr_type ";
    	sqlCmd += " where loanapply_bn01.bank_code=? ";
    	paramList.add(bank_no) ;
    	sqlCmd += " group by loanapply_bn01.acc_tr_type,acc_tr_name ";
    	sqlCmd += " order by loanapply_bn01.acc_tr_type ";
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"acc_tr_type,acc_tr_name");     
        return dbData;
    }
    private List getTM006W_ApplyList(String bank_no,String acc_tr_type,String applydate){
    	List paramList = new ArrayList() ;
    	String sqlCmd = "select loanapply_period.acc_tr_name,loanapply_period.applydate,loanapply_wml01.add_date,applytype,applydate_b,applydate_e,loanapply_wml01.cnt_name,loanapply_wml01.cnt_tel ";
    	sqlCmd += " from loanapply_period left join (select * from loanapply_wml01 where bank_code=?)loanapply_wml01 ";
    	sqlCmd += "   on loanapply_period.acc_tr_type=loanapply_wml01.acc_tr_type and loanapply_period.applydate=loanapply_wml01.applydate ";
    	paramList.add(bank_no) ;
    	sqlCmd += " where loanapply_period.acc_tr_type=? ";
    	paramList.add(acc_tr_type) ;
    	if(!"".equals( applydate)){
    		sqlCmd += " and loanapply_period.applydate = TO_DATE(?, 'YYYY/MM/DD') ";
    		paramList.add( applydate) ;
    	}
    	sqlCmd += " order by loanapply_period.applydate ";
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"applydate,add_date,applytype,applydate_b,applydate_e,cnt_name,cnt_tel");     
        return dbData;
    }
    private List getAccDivList(String acc_tr_type,String acc_div){
    	List paramList = new ArrayList() ;
    	//程序為顯示畫面，查詢條件
    	String sqlCmd = "select acc_tr_type,acc_tr_name,acc_div,acc_code,acc_name ";
    	sqlCmd += " from loanapply_ncacno ";
    	sqlCmd += " where acc_tr_type=? and acc_div=? ";
    	sqlCmd += " order by acc_div,acc_range ";
    	paramList.add(acc_tr_type) ;
    	paramList.add(acc_div) ;
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"acc_tr_type,acc_tr_name,acc_div,acc_code,acc_name");     
        return dbData;
    }
    //該申報基準日以前的累計資料
    private List getPreEditInfo(String acc_tr_type,String bank_code,String applydate){
    	List paramList = new ArrayList() ;
    	String sqlCmd = "select acc_div,acc_code,sum(apply_cnt) as apply_cnt_sum,sum(apply_amt) as apply_amt_sum,sum(apply_bal) as apply_bal_sum, ";
    	sqlCmd += "             sum(appr_cnt) as appr_cnt_sum,sum(appr_amt) as appr_amt_sum,sum(appr_bal) as appr_bal_sum ";
    	sqlCmd += " from loanapply_rpt ";
    	sqlCmd += " where acc_tr_type=? and bank_code=? and applydate < TO_DATE(?, 'YYYY/MM/DD') ";
    	sqlCmd += " group by acc_div,acc_code ";
    	sqlCmd += " order by acc_div,acc_code ";
    	paramList.add(acc_tr_type) ;
    	paramList.add(bank_code) ;
    	paramList.add(applydate) ;
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"acc_div,acc_code,apply_cnt_sum,apply_amt_sum,apply_bal_sum,appr_cnt_sum,appr_amt_sum,appr_bal_sum");     
        return dbData;
    }
    private List getEditInfo(String acc_tr_type,String bank_code,String applydate){
    	List paramList = new ArrayList() ;
    	String sqlCmd = "select loanapply_rpt.acc_tr_type,";//--金融協助措施報表代碼
    	sqlCmd += " acc_tr_name,";//--金融助措施名稱
    	sqlCmd += " loanapply_rpt.acc_div,";//--01:舊貸 02:新貸
    	sqlCmd += " loanapply_rpt.acc_code,";//--貸款子項別代碼
    	sqlCmd += " apply_cnt,";        //--申請件數
    	sqlCmd += " apply_amt,";        //--申請貸款金額
    	sqlCmd += " apply_bal,";        //--申請貸款餘額
    	sqlCmd += " apply_cnt_sum,";    //--申請累計件數
    	sqlCmd += " apply_amt_sum,";    //--申請累計貸款金額";
    	sqlCmd += " apply_bal_sum,";    //--申請累計貸款餘額 ";
    	sqlCmd += " appr_cnt,";        //--核准件數 ";
    	sqlCmd += " appr_amt,";        //--核准貸款金額 ";
    	sqlCmd += " appr_bal,";        //--核准貸款餘額 ";
    	sqlCmd += " appr_cnt_sum,";    //--核准累計件數 ";
    	sqlCmd += " appr_amt_sum,";    //--核准累計貸款金額 ";
    	sqlCmd += " appr_bal_sum,";    //--核准累計貸款餘額 ";
    	sqlCmd += " nonappr_cnt,";    //--不予核准件數 ";
    	sqlCmd += " nonappr_reason ";   //--不予核准原因 ";
    	sqlCmd += " from loanapply_ncacno left join loanapply_rpt ";
    	sqlCmd += " on loanapply_ncacno.acc_tr_type = loanapply_rpt.acc_tr_type and loanapply_ncacno.acc_div=loanapply_rpt.acc_div and loanapply_ncacno.acc_code=loanapply_rpt.acc_code ";
    	sqlCmd += " where loanapply_rpt.acc_tr_type=? ";
    	sqlCmd += " and bank_code=? ";
    	sqlCmd += " and applydate = TO_DATE(?, 'YYYY/MM/DD') ";
    	sqlCmd += " order by acc_div,acc_range ";
    	paramList.add(acc_tr_type) ;
    	paramList.add(bank_code) ;
    	paramList.add(applydate) ;
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"acc_tr_type,acc_tr_name,acc_div,acc_code,apply_cnt,apply_amt,apply_bal,"+
													        	  "apply_cnt_sum,apply_amt_sum,apply_bal_sum,appr_cnt,appr_amt,appr_bal,"+
													        	  "appr_cnt_sum,appr_amt_sum,appr_bal_sum,nonappr_cnt,nonappr_reason");     
        return dbData;
    }
    
	public String UpdateDB(HttpServletRequest request,String bank_no,String lguser_id,String lguser_name) throws Exception{
		StringBuffer sqlCmd = new StringBuffer() ;
		String errMsg="";
		String acc_tr_type = (request.getParameter("acc_Tr_Type")==null)?"":(String)request.getParameter("acc_Tr_Type");
		String applyDate = (request.getParameter("applyDate")==null)?"":(String)request.getParameter("applyDate");
		String cnt_Name = (request.getParameter("cnt_Name")==null)?"":(String)request.getParameter("cnt_Name");
		String cnt_Tel = (request.getParameter("cnt_Tel")==null)?"":(String)request.getParameter("cnt_Tel");
		String loan_Sbm = (request.getParameter("loan_Sbm")==null)?"":(String)request.getParameter("loan_Sbm");
		
		List updateDBList = new LinkedList();//0:sql 1:data
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		
		try {
			 //List updateDBSqlList = new LinkedList();
			 List paramList =new ArrayList() ;
			 sqlCmd.append("SELECT acc_tr_type FROM LOANAPPLY_RPT WHERE acc_tr_type=? AND bank_code=? and applydate = TO_DATE(?, 'YYYY/MM/DD') ");
			 paramList.add(acc_tr_type) ;
			 paramList.add(bank_no) ;
			 paramList.add(applyDate) ;
			 List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"acc_tr_type");
 			 if (data.size() > 0){//do update
 				updateDBSqlList = new LinkedList();
				updateDBDataList = new LinkedList();
 				sqlCmd.setLength(0); 
				sqlCmd.append("delete LOANAPPLY_RPT where acc_tr_type=? and bank_code=? and applydate = TO_DATE(?, 'YYYY/MM/DD') ");
				updateDBSqlList.add(sqlCmd.toString());
				paramList =new ArrayList() ;
				paramList.add(acc_tr_type) ;
				paramList.add(bank_no) ;
				paramList.add(applyDate) ;
				updateDBDataList.add(paramList);
		        updateDBSqlList.add(updateDBDataList);
		        updateDBList.add(updateDBSqlList);
			 }
				
			//寫入table(LOANAPPLY_RPT)金融協助措施-申報資料檔,每個貸款子項別各寫入一筆loanapply_rpt
			sqlCmd.setLength(0); 
			sqlCmd.append("insert into LOANAPPLY_RPT( ");		
			sqlCmd.append(" ACC_TR_TYPE,ACC_DIV,BANK_CODE,APPLYDATE,ACC_CODE,");
			sqlCmd.append(" APPLY_CNT,APPLY_AMT,APPLY_BAL,APPLY_CNT_SUM,APPLY_AMT_SUM,APPLY_BAL_SUM,");
			sqlCmd.append(" APPR_CNT,APPR_AMT,APPR_BAL,APPR_CNT_SUM,APPR_AMT_SUM,APPR_BAL_SUM,");
			sqlCmd.append(" NONAPPR_CNT,NONAPPR_REASON,USER_ID,USER_NAME,UPDATE_DATE  ");
			sqlCmd.append(" )values(?,?,?,TO_DATE(?, 'YYYY/MM/DD'),?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,sysdate) ");
			String[] infoList = loan_Sbm.split(";");
			for (String l:infoList) {
				//System.out.println(l);
				updateDBSqlList = new LinkedList();
				updateDBDataList = new LinkedList();
				updateDBSqlList.add(sqlCmd.toString());
				String[] tokens = l.split(",");
				paramList =new ArrayList() ;
				paramList.add(acc_tr_type) ;
				paramList.add(tokens[0].toString()) ;//ACC_DIV
				paramList.add(bank_no) ;
				paramList.add(applyDate) ;
				paramList.add(tokens[1].toString()) ;//ACC_CODE
				paramList.add(tokens[2].toString()) ;//APPLY_CNT
				paramList.add(tokens[3].toString()) ;//APPLY_AMT
				paramList.add(tokens[4].toString()) ;//APPLY_BAL
				paramList.add(tokens[5].toString()) ;//APPLY_CNT_SUM
				paramList.add(tokens[6].toString()) ;//APPLY_AMT_SUM
				paramList.add(tokens[7].toString()) ;//APPLY_BAL_SUM
				paramList.add(tokens[8].toString()) ;//APPR_CNT
				paramList.add(tokens[9].toString()) ;//APPR_AMT
				paramList.add(tokens[10].toString()) ;//APPR_BAL
				paramList.add(tokens[11].toString()) ;//APPR_CNT_SUM
				paramList.add(tokens[12].toString()) ;//APPR_AMT_SUM
				paramList.add(tokens[13].toString()) ;//APPR_BAL_SUM
				paramList.add(tokens[14].toString()) ;//NONAPPR_CNT
				String reason = tokens[15].toString();
				if("null".equals(reason))reason="";
				paramList.add(reason) ;//NONAPPR_REASON
				paramList.add(lguser_id) ;
				paramList.add(lguser_name) ;
				updateDBDataList.add(paramList);
			    updateDBSqlList.add(updateDBDataList);
			    updateDBList.add(updateDBSqlList);
			} 
			 sqlCmd.setLength(0);
			 sqlCmd.append("SELECT acc_tr_type FROM LOANAPPLY_WML01 WHERE acc_tr_type=? AND bank_code=? and applydate = TO_DATE(?, 'YYYY/MM/DD') ");
			 paramList =new ArrayList() ;
			 paramList.add(acc_tr_type) ;
			 paramList.add(bank_no) ;
			 paramList.add(applyDate) ;
			 data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"acc_tr_type");
 			 if (data.size() > 0){
 				//寫入(LOANAPPLY_WML01_LOG)金融協助措施-申報歷程紀錄檔
 		        updateDBSqlList = new LinkedList();
 				updateDBDataList = new LinkedList();
  				sqlCmd.setLength(0); 
  				sqlCmd.append("insert into loanapply_wml01_log ");
  				sqlCmd.append("select applydate,bank_code,acc_tr_type,cnt_name,cnt_tel,add_user,add_name,add_date,user_id,user_name,update_date,?,?,sysdate,'U' ");
  				sqlCmd.append(" from loanapply_wml01 where applydate = TO_DATE(?, 'YYYY/MM/DD') and bank_code=? and acc_tr_type=? ");
 				updateDBSqlList.add(sqlCmd.toString());
 				paramList =new ArrayList() ;
 				paramList.add(lguser_id) ;
 				paramList.add(lguser_name) ;
 				paramList.add(applyDate) ;
 				paramList.add(bank_no) ;
 				paramList.add(acc_tr_type) ;
 				updateDBDataList.add(paramList);
 		        updateDBSqlList.add(updateDBDataList);
 		        updateDBList.add(updateDBSqlList);
 		        //更新(LOANAPPLY_WML01)金融協助措施-申報紀錄檔
 		        updateDBSqlList = new LinkedList();
 				updateDBDataList = new LinkedList();
  				sqlCmd.setLength(0); 
 				sqlCmd.append("update LOANAPPLY_WML01 set ");
 				sqlCmd.append("  CNT_NAME=?,CNT_TEL=?,USER_ID=?,USER_NAME=?,UPDATE_DATE=sysdate ");
 				sqlCmd.append(" where applydate = TO_DATE(?, 'YYYY/MM/DD') and bank_code=? and acc_tr_type=? ");
 				updateDBSqlList.add(sqlCmd.toString());
 				paramList =new ArrayList() ;
 				paramList.add(cnt_Name) ;
 				paramList.add(cnt_Tel) ;
 				paramList.add(lguser_id) ;
 				paramList.add(lguser_name) ;
 				paramList.add(applyDate) ;
 				paramList.add(bank_no) ;
 				paramList.add(acc_tr_type) ;
 				updateDBDataList.add(paramList);
 		        updateDBSqlList.add(updateDBDataList);
 		        updateDBList.add(updateDBSqlList);
 			 }else{//do insert
 				//寫入table(LOANAPPLY_WML01)金融協助措施-申報紀錄檔
 				updateDBSqlList = new LinkedList();
 				updateDBDataList = new LinkedList();
 	 			sqlCmd.setLength(0); 
 				sqlCmd.append("insert into LOANAPPLY_WML01(");
 				sqlCmd.append("APPLYDATE,BANK_CODE,ACC_TR_TYPE,CNT_NAME,CNT_TEL,ADD_USER,ADD_NAME,ADD_DATE,USER_ID,USER_NAME,UPDATE_DATE");
 				sqlCmd.append(")values(TO_DATE(?, 'YYYY/MM/DD'),?,?,?,?,?,?,sysdate,?,?,sysdate)");
 				updateDBSqlList.add(sqlCmd.toString());
 				paramList =new ArrayList() ;
 				paramList.add(applyDate) ;
 				paramList.add(bank_no) ;
 				paramList.add(acc_tr_type) ;
 				paramList.add(cnt_Name) ;//CNT_NAME
 				paramList.add(cnt_Tel) ;//CNT_TEL
 				paramList.add(lguser_id) ;
 				paramList.add(lguser_name) ;
 				paramList.add(lguser_id) ;
 				paramList.add(lguser_name) ;
 				updateDBDataList.add(paramList);
 			    updateDBSqlList.add(updateDBDataList);
 			    updateDBList.add(updateDBSqlList);
 			 }
			if(DBManager.updateDB_ps(updateDBList)){
				errMsg = errMsg + "相關資料寫入資料庫成功";
			}else{
				errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
			}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}
		return errMsg;
	}
	public String DeleteDB(HttpServletRequest request,String bank_no,String lguser_id,String lguser_name) throws Exception{
		StringBuffer sqlCmd = new StringBuffer() ;
		String errMsg="";
		String acc_tr_type = (request.getParameter("acc_Tr_Type")==null)?"":(String)request.getParameter("acc_Tr_Type");
		String applyDate = (request.getParameter("applyDate")==null)?"":(String)request.getParameter("applyDate");
		String cnt_Name = (request.getParameter("cnt_Name")==null)?"":(String)request.getParameter("cnt_Name");
		String cnt_Tel = (request.getParameter("cnt_Tel")==null)?"":(String)request.getParameter("cnt_Tel");
		String loan_Sbm = (request.getParameter("loan_Sbm")==null)?"":(String)request.getParameter("loan_Sbm");
		
		List updateDBList = new LinkedList();//0:sql 1:data
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		
		try {
			 //List updateDBSqlList = new LinkedList();
			 List paramList =new ArrayList() ;
			 
				sqlCmd.append("delete LOANAPPLY_RPT where acc_tr_type=? and bank_code=? and applydate = TO_DATE(?, 'YYYY/MM/DD') ");
				updateDBSqlList.add(sqlCmd.toString());
				paramList =new ArrayList() ;
				paramList.add(acc_tr_type) ;
				paramList.add(bank_no) ;
				paramList.add(applyDate) ;
				updateDBDataList.add(paramList);
		        updateDBSqlList.add(updateDBDataList);
		        updateDBList.add(updateDBSqlList);
			 
			
			 sqlCmd.setLength(0);
			 sqlCmd.append("SELECT acc_tr_type FROM LOANAPPLY_WML01 WHERE acc_tr_type=? AND bank_code=? and applydate = TO_DATE(?, 'YYYY/MM/DD') ");
			 paramList =new ArrayList() ;
			 paramList.add(acc_tr_type) ;
			 paramList.add(bank_no) ;
			 paramList.add(applyDate) ;
			 List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"acc_tr_type");
 			 if (data.size() > 0){
 				//寫入(LOANAPPLY_WML01_LOG)金融協助措施-申報歷程紀錄檔
 		        updateDBSqlList = new LinkedList();
 				updateDBDataList = new LinkedList();
  				sqlCmd.setLength(0); 
  				sqlCmd.append("insert into loanapply_wml01_log ");
  				sqlCmd.append("select applydate,bank_code,acc_tr_type,cnt_name,cnt_tel,add_user,add_name,add_date,user_id,user_name,update_date,?,?,sysdate,'D' ");
  				sqlCmd.append(" from loanapply_wml01 where applydate = TO_DATE(?, 'YYYY/MM/DD') and bank_code=? and acc_tr_type=? ");
 				updateDBSqlList.add(sqlCmd.toString());
 				paramList =new ArrayList() ;
 				paramList.add(lguser_id) ;
 				paramList.add(lguser_name) ;
 				paramList.add(applyDate) ;
 				paramList.add(bank_no) ;
 				paramList.add(acc_tr_type) ;
 				updateDBDataList.add(paramList);
 		        updateDBSqlList.add(updateDBDataList);
 		        updateDBList.add(updateDBSqlList);
 		        //更新(LOANAPPLY_WML01)金融協助措施-申報紀錄檔
 		        updateDBSqlList = new LinkedList();
 				updateDBDataList = new LinkedList();
  				sqlCmd.setLength(0); 
 				sqlCmd.append("delete LOANAPPLY_WML01 ");
 				sqlCmd.append(" where applydate = TO_DATE(?, 'YYYY/MM/DD') and bank_code=? and acc_tr_type=? ");
 				updateDBSqlList.add(sqlCmd.toString());
 				paramList =new ArrayList() ;
 				paramList.add(applyDate) ;
 				paramList.add(bank_no) ;
 				paramList.add(acc_tr_type) ;
 				updateDBDataList.add(paramList);
 		        updateDBSqlList.add(updateDBDataList);
 		        updateDBList.add(updateDBSqlList);
 			 }
 			if(DBManager.updateDB_ps(updateDBList)){
				errMsg = errMsg + "相關資料刪除成功";
			}else{
				   errMsg = errMsg + "相關資料刪除失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
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
