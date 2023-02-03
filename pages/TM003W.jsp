<%
//105.09.29 create by2968
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
      System.out.println("TM003W login timeout");
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
	    	    List dbData = getTM003W_List(bank_no);
	    	    System.out.println("dbData="+dbData.size());
	    	    request.setAttribute("dbData",dbData);
	        	rd = application.getRequestDispatcher( ListPgName +"?bank_no="+bank_no);
	    	}else if(act.equals("Edit")){
	    		String acc_Tr_Type = ( request.getParameter("acc_Tr_Type")==null ) ? "" : (String)request.getParameter("acc_Tr_Type");
	    		request.setAttribute("accDivList01",getAccDivList(acc_Tr_Type,"01"));
	    		request.setAttribute("accDivList02",getAccDivList(acc_Tr_Type,"02"));
	    	    List dbData = getEditInfo(acc_Tr_Type,bank_no);
	    	    if(dbData.size()==0){
	    	    	act="new";
	    	    }
	    	    request.setAttribute("dbData",dbData);
	    		rd = application.getRequestDispatcher( EditPgName +"?act="+act);
	    	}else if(act.equals("Update")){
	    	    actMsg = UpdateDB(request,bank_no,lguser_id,lguser_name);
	        	rd = application.getRequestDispatcher( nextPgName+"?goPages=TM003W.jsp&act=Qry");
	    	}else if(act.equals("Delete")){
	    	    actMsg = DeleteDB(request,bank_no,lguser_id,lguser_name);
	        	rd = application.getRequestDispatcher( nextPgName+"?goPages=TM003W.jsp&act=Qry");
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
    private final static String EditPgName = "/pages/TM003W_Edit.jsp";
    private final static String ListPgName = "/pages/TM003W_List.jsp";
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    private boolean CheckPermission(HttpServletRequest request){//檢核權限
    	    boolean CheckOK=false;
    	    HttpSession session = request.getSession();
            Properties permission = ( session.getAttribute("TM003W")==null ) ? new Properties() : (Properties)session.getAttribute("TM003W");
            if(permission == null){
              System.out.println("TM003W.permission == null");
            }else{
               System.out.println("TM003W.permission.size ="+permission.size());

            }
            //只要有Query的權限,就可以進入畫面
        	if(permission != null && permission.get("Q") != null && permission.get("Q").equals("Y")){
        	   CheckOK = true;//Query
        	}
        	return CheckOK;
    }
    
    private List getTM003W_List(String bank_no){
    	List paramList = new ArrayList() ;
    	//程序為顯示畫面，查詢條件
    	String sqlCmd = "select loan_bn01.acc_tr_type,acc_tr_name ";
    	sqlCmd += " from loan_ncacno left join loan_bn01 on loan_ncacno.acc_tr_type = loan_bn01.acc_tr_type ";
    	sqlCmd += " where loan_bn01.bank_code=? ";
    	paramList.add(bank_no) ;
    	sqlCmd += " group by loan_bn01.acc_tr_type,acc_tr_name ";
    	sqlCmd += " order by loan_bn01.acc_tr_type ";
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"acc_tr_type,acc_tr_name");     
        return dbData;
    }
    private List getAccDivList(String acc_tr_type,String acc_div){
    	List paramList = new ArrayList() ;
    	//程序為顯示畫面，查詢條件
    	String sqlCmd = "select acc_tr_type,acc_tr_name,acc_div,acc_code,acc_name ";
    	sqlCmd += " from loan_ncacno ";
    	sqlCmd += " where acc_tr_type=? and acc_div=? ";
    	sqlCmd += " order by acc_div,acc_range ";
    	paramList.add(acc_tr_type) ;
    	paramList.add(acc_div) ;
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"acc_tr_type,acc_tr_name,acc_div,acc_code,acc_name");     
        return dbData;
    }
    private List getEditInfo(String acc_tr_type,String bank_code){
    	List paramList = new ArrayList() ;
    	//程序為顯示畫面，查詢條件
    	String sqlCmd = "select loan_rpt.acc_tr_type,acc_tr_name,loan_rpt.acc_div,loan_rpt.acc_code,acc_name,loan_cnt,loan_amt ";
    	sqlCmd += " from loan_ncacno ";
    	sqlCmd += " left join loan_rpt on loan_ncacno.acc_tr_type = loan_rpt.acc_tr_type and loan_ncacno.acc_div=loan_rpt.acc_div and loan_ncacno.acc_code=loan_rpt.acc_code ";
    	sqlCmd += " where loan_rpt.acc_tr_type=? and bank_code=? ";
    	sqlCmd += " order by acc_div,acc_range ";
    	paramList.add(acc_tr_type) ;
    	paramList.add(bank_code) ;
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"acc_tr_type,acc_tr_name,acc_div,acc_code,acc_name,loan_cnt,loan_amt");     
        return dbData;
    }
    
	public String UpdateDB(HttpServletRequest request,String bank_no,String lguser_id,String lguser_name) throws Exception{
		StringBuffer sqlCmd = new StringBuffer() ;
		String errMsg="";
		String acc_tr_type = ((String)request.getParameter("acc_Tr_Type")==null)?"":(String)request.getParameter("acc_Tr_Type");
		String loan_Sbm = ((String)request.getParameter("loan_Sbm")==null)?"":(String)request.getParameter("loan_Sbm");
		List updateDBList = new LinkedList();//0:sql 1:data
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		
		try {
			 //List updateDBSqlList = new LinkedList();
			 List paramList =new ArrayList() ;
			 sqlCmd.append("SELECT acc_tr_type FROM loan_rpt WHERE acc_tr_type=? AND bank_code=?" );
			 paramList.add(acc_tr_type) ;
			 paramList.add(bank_no) ;
			 List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"acc_tr_type");
 			 if (data.size() > 0){
 				updateDBSqlList = new LinkedList();
				updateDBDataList = new LinkedList();
 				sqlCmd.setLength(0); 
				sqlCmd.append("delete loan_rpt where acc_tr_type=? and bank_code=? ");
				updateDBSqlList.add(sqlCmd.toString());
				paramList =new ArrayList() ;
				paramList.add(acc_tr_type) ;
				paramList.add(bank_no) ;
				updateDBDataList.add(paramList);
		        updateDBSqlList.add(updateDBDataList);
		        updateDBList.add(updateDBSqlList);
			 }
				
				//寫入table(LOAN_RPT)規劃協助措施-申報資料檔,每個貸款子項別各寫入一筆loan_rpt
				sqlCmd.setLength(0); 
				sqlCmd.append("insert into LOAN_RPT ");		
				sqlCmd.append(" (ACC_TR_TYPE,ACC_DIV,BANK_CODE,ACC_CODE,LOAN_CNT,LOAN_AMT,USER_ID,USER_NAME,UPDATE_DATE) ");
				sqlCmd.append("values(?,?,?,?,?,?,?,?,sysdate) ");
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
					paramList.add(tokens[1].toString()) ;//ACC_CODE
					paramList.add(tokens[2].toString()) ;//LOAN_CNT
					paramList.add(tokens[3].toString()) ;//LOAN_AMT
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
		String acc_tr_type = ((String)request.getParameter("acc_Tr_Type")==null)?"":(String)request.getParameter("acc_Tr_Type");
		String loan_Sbm = ((String)request.getParameter("loan_Sbm")==null)?"":(String)request.getParameter("loan_Sbm");
		List updateDBList = new LinkedList();//0:sql 1:data
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		
		try {
			 List paramList =new ArrayList() ;
			 
				sqlCmd.append("delete loan_rpt where acc_tr_type=? and bank_code=? ");
				updateDBSqlList.add(sqlCmd.toString());
				paramList =new ArrayList() ;
				paramList.add(acc_tr_type) ;
				paramList.add(bank_no) ;
				updateDBDataList.add(paramList);
		        updateDBSqlList.add(updateDBDataList);
		        updateDBList.add(updateDBSqlList);
			 
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
