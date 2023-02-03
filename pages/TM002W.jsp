<%
//105.9.7 create by 2968
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.*" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@include file="./include/Header.include" %>
<%
   	String acc_Tr_Type = Utility.getTrimString(dataMap.get("acc_Tr_Type"));
	String acc_Tr_Name = Utility.getTrimString(dataMap.get("acc_Tr_Name"));
	System.out.println("*** s acc_Tr_Type="+acc_Tr_Type);
	System.out.println("*** s acc_Tr_Name="+acc_Tr_Name);
	//fix 93.12.17 從 session取出登入者資訊=======================================================================================
	String userid = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");
	String username = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");
	//======================================================================================================================
	//fix 93.12.18 若有已點選的tbank_no,則以已點選的tbank_no為主============================================================
	String bank_code = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");
	String nowtbank_no = Utility.getTrimString(dataMap.get("tbank_no"));
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session
	}
	bank_code = ( session.getAttribute("nowtbank_no")==null ) ? bank_code : (String)session.getAttribute("nowtbank_no");
	//=======================================================================================================================
	//fix 93.12.20 若有已點選的bank_type,則以已點選的bank_type為主============================================================
	String bank_type = Utility.getTrimString(dataMap.get("bank_type"));
	bank_type = ( session.getAttribute("nowbank_type")==null ) ? bank_type : (String)session.getAttribute("nowbank_type");
	//=======================================================================================================================
	if(session.getAttribute("nowbank_type") == null)
	   System.out.println("nowbank_type == null");
	else  System.out.println((String)session.getAttribute("nowbank_type"));


  	if(!Utility.CheckPermission(request,report_no)){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );
    }else{
    	//set next jsp
    	if(act.equals("List")){
        	rd = application.getRequestDispatcher( ListPgName+"?act=List&test=nothing" );
        }else if(act.equals("new")){
        	request.setAttribute("AllLoanItem", getAllLoanItem());//貸款種類別
        	request.setAttribute("AllLoanSubItem", getAllLoanSubItem());//貸款子項別
        	request.setAttribute("AllBankList", Utility.getLoanBank());//貸款經辦機構
        	request.setAttribute("AllCityList",Utility.getCity());
     	    rd = application.getRequestDispatcher( EditPgName+"?act=new&test=nothing");
    	}else if(act.equals("Edit")){
    		request.setAttribute("acc_Tr_Type", acc_Tr_Type);
        	request.setAttribute("AllLoanItem", getAllLoanItem());//貸款種類別
        	request.setAttribute("AllLoanSubItem", getAllLoanSubItem());//貸款子項別
        	request.setAttribute("AllBankList", Utility.getLoanBank());//貸款經辦機構
        	request.setAttribute("AllCityList",Utility.getCity());
	    	request.setAttribute("SelDataList", getSelItemList(acc_Tr_Type));//已挑選的舊貸展延需求、新貸需求
	    	request.setAttribute("SelBankList", getBankData(acc_Tr_Type));//已挑選的貸款經辦經機名稱
	    	rd = application.getRequestDispatcher( EditPgName+"?act=Edit&acc_Tr_Type="+acc_Tr_Type+"&test=nothing");
    	}else if(act.equals("Insert")||act.equals("Update")){
    	    actMsg = doUpdate(request,userid,username);
	   	    request.setAttribute("actMsg",actMsg);
        	rd = application.getRequestDispatcher( nextPgName +"?goPages=TM002W_List.jsp&act=List");
    	}else if(act.equals("Delete")){
    	    actMsg = doDelete(request);
	   	    request.setAttribute("actMsg",actMsg);
        	rd = application.getRequestDispatcher( nextPgName +"?goPages=TM002W_List.jsp&act=List");
    	}
    }
%>

<%@include file="./include/Tail.include" %>

<%!
    private final static String report_no = "TM002W";
    private final static String nextPgName = "/pages/ActMsg.jsp";
    private String EditPgName = "/pages/"+report_no+"_Edit.jsp";
    private final static String ListPgName = "/pages/"+report_no+"_List.jsp";
    private final static String SubListPgName = "/pages/"+report_no+"_subList.jsp";
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    
    private List getAllLoanItem(){
   		//查詢條件
   	    List dbData =null;
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		sqlCmd.append("select loan_item,loan_item_name ");
		sqlCmd.append("  from loan_item ");
		sqlCmd.append(" order by input_order ");
		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"loan_item,loan_item_name");

		return dbData;
    }
    private List getAllLoanSubItem(){
    	//查詢條件
   	    List dbData =null;
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		sqlCmd.append("select loan_item,subitem,subitem_name ");
		sqlCmd.append("  from loan_subitem ");
		sqlCmd.append(" order by loan_item,input_order ");
		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"loan_item,subitem,subitem_name");

		return dbData;
    }
    private List getSelItemList(String type){
   		//查詢條件
   	    List dbData =null;
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		sqlCmd.append("select acc_div,acc_code,acc_name,RATE_PERIOD,acc_tr_name,acc_tr_type ");
		sqlCmd.append("  from loan_ncacno ");
		sqlCmd.append(" where acc_tr_type = ? ");
		sqlCmd.append(" order by acc_div,acc_range ");
		paramList.add(type);    
		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"acc_div,acc_code,acc_name,rate_period,acc_tr_name,acc_tr_type");

		return dbData;
    }
    private List getBankData(String type){
   		//查詢條件
   	    List dbData =null;
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		sqlCmd.append("select bn01.bn_type,HSIEN_id, BN01.BANK_NO , BANK_NAME, BANK_TYPE,bn01.m_year ");
		sqlCmd.append("  from (select acc_tr_type,bn01.* from loan_bn01 left join bn01 on loan_bn01.bank_code=bn01.bank_no)BN01, WLX01 ");
		sqlCmd.append(" where BN01.BANK_NO = WLX01.BANK_NO(+) ");
		sqlCmd.append("   and bank_type in ('A','6','7') ");//--機構類別:銀行/農會/漁會
		sqlCmd.append("   and wlx01.m_year = bn01.m_year ");
		sqlCmd.append("   and bn01.m_year = 100 ");//--都固定只抓100年度的         
		sqlCmd.append("   and acc_tr_type = ? "); 
		sqlCmd.append(" order by hsien_id,bank_type,BANK_NO ");
		paramList.add(type);    
		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"bn_type,hsien_id,bank_no,bank_name,bank_type,m_year");

		return dbData;
    }
    
    
    public String doUpdate(HttpServletRequest request,String userid,String username) throws Exception{
    	String acc_Tr_Type = Utility.getTrimString(request.getParameter("acc_Tr_Type"));
    	String acc_Tr_Name = Utility.getTrimString(request.getParameter("acc_Tr_Name"));
    	String rate_Period1 = Utility.getTrimString(request.getParameter("rate_Period1"));
    	String rate_Period2 = Utility.getTrimString(request.getParameter("rate_Period2"));
    	String selSubList1 = Utility.getTrimString(request.getParameter("selSubList1"));
    	String selSubList2 = Utility.getTrimString(request.getParameter("selSubList2"));
    	String selBankList = Utility.getTrimString(request.getParameter("selBankList"));
		String sqlCmd = "";
		String errMsg="";
		try {
			List updateDBList = new LinkedList();//0:sql 1:data
			List updateDBSqlList = new LinkedList();
			List updateDBDataList = new LinkedList();//儲存參數的List
			List dataList = new ArrayList();//儲存參數的data
			
			//doInsert
			if("".equals(acc_Tr_Type)){
			    sqlCmd = "select LPAD(to_number(max(acc_tr_type))+1, 6, 0) as acc_tr_type from loan_ncacno ";
			    List dbData = DBManager.QueryDB_SQLParam(sqlCmd,new ArrayList(),"acc_tr_type"); 
			    if(dbData != null && dbData.size() != 0 ){
			    	acc_Tr_Type = (((DataObject)dbData.get(0)).getValue("acc_tr_type")).toString();
			    }
			//doUpdate
			}else{
				//將屬於該acc_tr_type的先刪除
				sqlCmd = "delete loan_ncacno where acc_tr_type= ? ";
				updateDBSqlList.add(sqlCmd);
				dataList = new ArrayList();
				dataList.add(acc_Tr_Type);
				updateDBDataList.add(dataList);
		        updateDBSqlList.add(updateDBDataList);
		        updateDBList.add(updateDBSqlList);
		        
				updateDBSqlList = new LinkedList();
				updateDBDataList = new LinkedList();
		        sqlCmd = "delete loan_bn01 where acc_tr_type= ? ";
				updateDBSqlList.add(sqlCmd);
				dataList = new ArrayList();
				dataList.add(acc_Tr_Type);
				updateDBDataList.add(dataList);
		        updateDBSqlList.add(updateDBDataList);
		        updateDBList.add(updateDBSqlList);
			}
				
			//1.寫入table(LOAN_NCACNO)規劃協助措施-會計科目檔,每個已挑選的貸款子項別各寫入一筆loan_ncacno
			sqlCmd = "INSERT INTO LOAN_NCACNO(ACC_TR_TYPE,ACC_TR_NAME,ACC_DIV,ACC_CODE,ACC_NAME,ACC_RANGE,RATE_PERIOD,USER_ID,USER_NAME,UPDATE_DATE) ";
			sqlCmd += "VALUES(?,?,?,?,(select distinct subitem_name from loan_subItem where subitem=?),LPAD(to_number(?), 10, 0),?,?,?,sysdate)";
			
			//舊貸展延需求
			if(!"".equals(selSubList1)){
				String[] tokens = selSubList1.split(";");
				int n = 1;
				for (String token:tokens) {
					updateDBSqlList = new LinkedList();
					updateDBDataList = new LinkedList();
					updateDBSqlList.add(sqlCmd);
					dataList = new ArrayList();
					dataList.add(acc_Tr_Type);
					dataList.add(acc_Tr_Name);
					dataList.add("01");
					dataList.add(token);
					dataList.add(token);
					dataList.add(n);
					dataList.add(rate_Period1);
					dataList.add(userid);
					dataList.add(username);
					updateDBDataList.add(dataList);
			        updateDBSqlList.add(updateDBDataList);
			        updateDBList.add(updateDBSqlList);
			        n++;
				}
			}
			//新貸需求
			if(!"".equals(selSubList2)){
				String[] tokens = selSubList2.split(";");
				int n = 1;
				for (String token:tokens) {
					updateDBSqlList = new LinkedList();
					updateDBDataList = new LinkedList();
					updateDBSqlList.add(sqlCmd);
					dataList = new ArrayList();
					dataList.add(acc_Tr_Type);
					dataList.add(acc_Tr_Name);
					dataList.add("02");
					dataList.add(token);
					dataList.add(token);
					dataList.add(n);
					dataList.add(rate_Period2);
					dataList.add(userid);
					dataList.add(username);
					updateDBDataList.add(dataList);
			        updateDBSqlList.add(updateDBDataList);
			        updateDBList.add(updateDBSqlList);
			        n++;
			        //System.out.println("updateDBDataList add");
				}
			}
	        //2.寫入table(LOAN_BN01)規劃協助措施-需申報機構,每個已挑選的機構代碼各寫入一筆loan_bn01
	        sqlCmd =  " INSERT INTO LOAN_BN01(ACC_TR_TYPE,BANK_CODE,user_id,user_name,update_date) ";
			sqlCmd += " VALUES(?,?,?,?,sysdate) ";
			if(!"".equals(selBankList)){
				String[] tokens = selBankList.split(";");
				for (String token:tokens) {
					updateDBSqlList = new LinkedList();
					updateDBDataList = new LinkedList();
					updateDBSqlList.add(sqlCmd);
					dataList = new ArrayList();
					dataList.add(acc_Tr_Type);
					dataList.add(token);
					dataList.add(userid);
					dataList.add(username);
					updateDBDataList.add(dataList);
			        updateDBSqlList.add(updateDBDataList);
			        updateDBList.add(updateDBSqlList);
				}
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
	
    public String doDelete(HttpServletRequest request) throws Exception{		
		String errMsg="";
		String acc_Tr_Type = Utility.getTrimString(request.getParameter("acc_Tr_Type"));
    	String acc_Tr_Name = Utility.getTrimString(request.getParameter("acc_Tr_Name"));
    	String rate_Period1 = Utility.getTrimString(request.getParameter("rate_Period1"));
    	String rate_Period2 = Utility.getTrimString(request.getParameter("rate_Period2"));
    	String selSubList1 = Utility.getTrimString(request.getParameter("selSubList1"));
    	String selSubList2 = Utility.getTrimString(request.getParameter("selSubList2"));
    	String selBankList = Utility.getTrimString(request.getParameter("selBankList"));
    	String sqlCmd = "";
		List updateDBList = new LinkedList();//0:sql 1:data
		List updateDBSqlList = new LinkedList();
		List updateDBDataList = new LinkedList();//儲存參數的List
		List dataList = new LinkedList();//儲存參數的data	
		
		try {
			//將屬於該acc_tr_type的先刪除
			sqlCmd = "delete loan_ncacno where acc_tr_type= ? ";
			updateDBSqlList.add(sqlCmd);
			dataList = new ArrayList();
			dataList.add(acc_Tr_Type);
			updateDBDataList.add(dataList);
	        updateDBSqlList.add(updateDBDataList);
	        updateDBList.add(updateDBSqlList);
	        
			updateDBSqlList = new LinkedList();
			updateDBDataList = new LinkedList();
	        sqlCmd = "delete loan_bn01 where acc_tr_type= ? ";
			updateDBSqlList.add(sqlCmd);
			dataList = new ArrayList();
			dataList.add(acc_Tr_Type);
			updateDBDataList.add(dataList);
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
	

%>
