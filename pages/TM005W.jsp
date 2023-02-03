<%
//105.9.21 create by 2968
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
    		request.setAttribute("dbData", getListInfo(""));
        	rd = application.getRequestDispatcher( ListPgName+"?act="+act+"&test=nothing" );
    	}else if(act.equals("qList")){
        	request.setAttribute("dbData", getListInfo(acc_Tr_Name));
        	request.setAttribute("sAcc_Tr_Name", acc_Tr_Name);
            rd = application.getRequestDispatcher( ListPgName+"?act="+act+"&test=nothing" );
        }else if(act.equals("new")){
        	request.setAttribute("sAcc_Tr_Name", acc_Tr_Name);
        	request.setAttribute("AllAccTr",getAllAccTr());//下拉的協助措施名稱
        	request.setAttribute("AllLoanItem", getAllLoanItem());//貸款種類別
        	request.setAttribute("AllLoanSubItem", getAllLoanSubItem());//貸款子項別
        	request.setAttribute("AllBankList", Utility.getLoanBank());//貸款經辦機構
        	request.setAttribute("AllCityList",Utility.getCity());
        	request.setAttribute("AllSelItemList", getAllSelItemList(acc_Tr_Type));//預設舊貸展延需求、新貸需求
        	request.setAttribute("AllSelDataList", getAllSelDataList(acc_Tr_Type));//已挑選的舊貸展延需求、新貸需求
        	request.setAttribute("SelPreBankList", getPreBankData(acc_Tr_Type));//預設貸款經辦經機名稱
     	    rd = application.getRequestDispatcher( EditPgName+"?act="+act+"&test=nothing");
    	}else if(act.equals("Edit")){
    		request.setAttribute("sAcc_Tr_Name", acc_Tr_Name);
    		request.setAttribute("acc_Tr_Type", acc_Tr_Type);
    		request.setAttribute("AllAccTr",getAllAccTr());//下拉的協助措施名稱
        	request.setAttribute("AllLoanItem", getAllLoanItem());//貸款種類別
        	request.setAttribute("AllLoanSubItem", getAllLoanSubItem());//貸款子項別
        	request.setAttribute("AllBankList", Utility.getLoanBank());//貸款經辦機構
        	request.setAttribute("AllCityList",Utility.getCity());
        	request.setAttribute("EditInfo",getEditInfo(acc_Tr_Type));
        	request.setAttribute("AllSelDataList", getAllSelDataList(acc_Tr_Type));//已挑選的舊貸展延需求、新貸需求
	    	request.setAttribute("SelBankList", getBankData(acc_Tr_Type));//已挑選的貸款經辦經機名稱
	    	rd = application.getRequestDispatcher( EditPgName+"?act="+act+"&acc_Tr_Type="+acc_Tr_Type+"&test=nothing");
    	}else if(act.equals("Insert")||act.equals("Update")){
    	    actMsg = doUpdate(request,userid,username);
	   	    request.setAttribute("actMsg",actMsg);
        	rd = application.getRequestDispatcher( nextPgName +"?goPages=TM005W.jsp&act=List");
    	}else if(act.equals("Delete")){
    	    actMsg = doDelete(request);
	   	    request.setAttribute("actMsg",actMsg);
        	rd = application.getRequestDispatcher( nextPgName +"?goPages=TM005W.jsp&act=List");
    	}
    }
%>

<%@include file="./include/Tail.include" %>

<%!
    private final static String report_no = "TM005W";
    private final static String nextPgName = "/pages/ActMsg.jsp";
    private String EditPgName = "/pages/"+report_no+"_Edit.jsp";
    private final static String ListPgName = "/pages/"+report_no+"_List.jsp";
    private final static String SubListPgName = "/pages/"+report_no+"_subList.jsp";
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    
    private List getListInfo(String acc_tr_name){
   		//查詢條件
   	    List dbData =null;
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		sqlCmd.append("select acc_tr_type,acc_tr_name,applytype,decode(applytype,'2','雙週報','1','週報','4','月報','') as applytype_name ");
		sqlCmd.append("  from (select acc_tr_type,acc_tr_name,applytype from loanapply_period group by acc_tr_type,acc_tr_name,applytype)loanapply_period ");
		if(!"".equals(acc_tr_name)){
			sqlCmd.append(" where acc_tr_name like ? ");
			paramList.add("%"+acc_tr_name+"%");
		}
		sqlCmd.append(" order by acc_tr_type ");
		    
		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"acc_tr_type,acc_tr_name,applytype,applytype_name");
		List rtnList = new ArrayList();
		if(dbData!=null && dbData.size()>0){
			for(int i=0;i<dbData.size();i++){
				String key = (String)((DataObject)dbData.get(i)).getValue("acc_tr_type");
				Map p = new HashMap();
				p.put("acc_tr_type", key);
				p.put("acc_tr_name", (String)((DataObject)dbData.get(i)).getValue("acc_tr_name"));
				p.put("applytype", (String)((DataObject)dbData.get(i)).getValue("applytype"));
				p.put("applytype_name", (String)((DataObject)dbData.get(i)).getValue("applytype_name"));
				//1.取得該協助措施名稱.舊貸展延需求項目
				p.put("acc_name01", getAcc_Name(key,"01"));
				//2.取得該協助措施名稱.新貸展延需求項目
				p.put("acc_name02", getAcc_Name(key,"02"));
				//3.取得該協助措施名稱.申報日期
				p.put("applydate", getApplyDate(key));
				rtnList.add(p);
			}
		}
		return rtnList;
    }
    public static String getAcc_Name(String acc_tr_type,String acc_div) {
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		sqlCmd.append("select acc_name ");
		sqlCmd.append("  from loanapply_ncacno ");
		sqlCmd.append(" where acc_tr_type=? and acc_div=? ");
		sqlCmd.append(" order by acc_range asc ");
		paramList.add(acc_tr_type);
		paramList.add(acc_div);
		List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"acc_name");
		String rtnValue="";
  		if (dbData!=null && dbData.size()>0) {
  			for(int i=0;i<dbData.size();i++){
  				if(!"".equals(rtnValue)) rtnValue+="<br>";
  				rtnValue += (String)((DataObject)dbData.get(i)).getValue("acc_name");
  			}
  		}
  		if("".equals(rtnValue))rtnValue = "&nbsp;";
  		return rtnValue;
    }
    public static String getApplyDate(String acc_tr_type) {
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		sqlCmd.append("select F_TRANSCHINESEDATE(applydate) as applydate ");
		sqlCmd.append("  from loanapply_period ");
		sqlCmd.append(" where acc_tr_type=? ");
		paramList.add(acc_tr_type);
		List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"applydate");
		String rtnValue="";
  		if (dbData!=null && dbData.size()>0) {
  			for(int i=0;i<dbData.size();i++){
  				if(!"".equals(rtnValue)) rtnValue+="<br>";
  				rtnValue += (String)((DataObject)dbData.get(i)).getValue("applydate");
  			}
  			
  		}
  		return rtnValue;
    }
    private List getAllAccTr(){
   		//查詢條件
   	    List dbData =null;
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		sqlCmd.append("select acc_tr_type,acc_tr_name ");
		sqlCmd.append("  from loan_ncacno ");
		sqlCmd.append(" group by acc_tr_type,acc_tr_name ");
		sqlCmd.append(" order by acc_tr_type ");
		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"acc_tr_type,acc_tr_name");

		return dbData;
    }
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
    private List getEditInfo(String type){
   		//查詢條件
   	    List dbData =null;
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		sqlCmd.append("select acc_tr_type,acc_tr_name,applytype ");
		sqlCmd.append(" ,(to_number(to_char(startdate,'yyyy'))-1911)||to_char(startdate,'MMDD') as startdate ");
		sqlCmd.append(" ,(to_number(to_char(enddate,'yyyy'))-1911)||to_char(enddate,'MMDD') as enddate ");
		sqlCmd.append(" ,(to_number(to_char(begindate,'yyyy'))-1911)||to_char(begindate,'MMDD') as begindate ");
		sqlCmd.append(" ,(to_number(to_char(applydate,'yyyy'))-1911)||to_char(applydate,'MMDD') as applydate ");
		sqlCmd.append("  from LOANAPPLY_PERIOD ");
		sqlCmd.append(" where acc_tr_type=? ");
		paramList.add(type); 
		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"acc_tr_type,acc_tr_name,applytype,startdate,enddate,begindate,applydate");
		return dbData;
    }
    private List getAllSelDataList(String type){
   		//查詢條件
   	    List dbData =null;
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		sqlCmd.append("select acc_div,acc_code,acc_name,acc_tr_name,acc_tr_type ");
		sqlCmd.append("  from loanapply_ncacno ");
		sqlCmd.append(" order by acc_tr_type,acc_div,acc_range ");
		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"acc_div,acc_code,acc_name,acc_tr_name,acc_tr_type");
		return dbData;
    }
    private List getAllSelItemList(String type){
   		//查詢條件
   	    List dbData =null;
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		sqlCmd.append("select acc_div,acc_code,acc_name,acc_tr_name,acc_tr_type ");
		sqlCmd.append("  from loan_ncacno ");
		sqlCmd.append(" order by acc_tr_type,acc_div,acc_range ");
		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"acc_div,acc_code,acc_name,acc_tr_name,acc_tr_type");
		return dbData;
    }
    private List getBankData(String type){
   		//查詢條件
   	    List dbData =null;
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		sqlCmd.append("select bn01.bn_type,HSIEN_id, BN01.BANK_NO , BANK_NAME, BANK_TYPE,bn01.m_year ");
		sqlCmd.append("  from (select acc_tr_type,bn01.* from loanapply_bn01 left join bn01 on loanapply_bn01.bank_code=bn01.bank_no)BN01, WLX01 ");
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
    private List getPreBankData(String type){
   		//查詢條件
   	    List dbData =null;
    	StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();	
		sqlCmd.append("select acc_tr_type,bn01.bn_type,HSIEN_id, BN01.BANK_NO , BANK_NAME, BANK_TYPE,bn01.m_year ");
		sqlCmd.append("  from (select acc_tr_type,bn01.* from loan_bn01 left join bn01 on loan_bn01.bank_code=bn01.bank_no)BN01, WLX01 ");
		sqlCmd.append(" where BN01.BANK_NO = WLX01.BANK_NO(+) ");
		sqlCmd.append("   and bank_type in ('A','6','7') ");//--機構類別:銀行/農會/漁會
		sqlCmd.append("   and wlx01.m_year = bn01.m_year ");
		sqlCmd.append("   and bn01.m_year = 100 ");//--都固定只抓100年度的         
		//sqlCmd.append("   and acc_tr_type = ? "); 
		sqlCmd.append(" order by hsien_id,bank_type,BANK_NO ");
		//paramList.add(type);    
		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"acc_tr_type,bn_type,hsien_id,bank_no,bank_name,bank_type,m_year");

		return dbData;
    }
    
    public String doUpdate(HttpServletRequest request,String userid,String username) throws Exception{
    	String acc_Tr_Type = Utility.getTrimString(request.getParameter("acc_Tr_Type"));
    	String acc_Tr_Name = Utility.getTrimString(request.getParameter("acc_Tr_Name"));
    	String applyType = Utility.getTrimString(request.getParameter("applyType"));//申報頻率
    	String startDate = Utility.getTrimString(request.getParameter("startDate"));//申報期間(起)
    	String endDate = Utility.getTrimString(request.getParameter("endDate"));//申報期間(迄)
    	String beginDate = Utility.getTrimString(request.getParameter("beginDate"));//發布日期
    	String applyDateS = Utility.getTrimString(request.getParameter("applyDateS"));//申報基準日 list
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
			    sqlCmd = "select LPAD(to_number(max(acc_tr_type))+1, 6, 0) as acc_tr_type from loanapply_ncacno ";
			    List dbData = DBManager.QueryDB_SQLParam(sqlCmd,new ArrayList(),"acc_tr_type"); 
			    if(dbData != null && dbData.size() != 0 ){
			    	acc_Tr_Type = (((DataObject)dbData.get(0)).getValue("acc_tr_type")).toString();
			    }
			//doUpdate
			}else{
				//將屬於該acc_tr_type的先刪除
				//金融協助措施-會計科目檔
				sqlCmd = "delete loanapply_ncacno where acc_tr_type= ? ";
				updateDBSqlList.add(sqlCmd);
				dataList = new ArrayList();
				dataList.add(acc_Tr_Type);
				updateDBDataList.add(dataList);
		        updateDBSqlList.add(updateDBDataList);
		        updateDBList.add(updateDBSqlList);
		        
				updateDBSqlList = new LinkedList();
				updateDBDataList = new LinkedList();
				//金融協助措施-需申報機構
		        sqlCmd = "delete loanapply_bn01 where acc_tr_type= ? ";
				updateDBSqlList.add(sqlCmd);
				dataList = new ArrayList();
				dataList.add(acc_Tr_Type);
				updateDBDataList.add(dataList);
		        updateDBSqlList.add(updateDBDataList);
		        updateDBList.add(updateDBSqlList);
		        
		        updateDBSqlList = new LinkedList();
				updateDBDataList = new LinkedList();
				//金融協助措施-申報期間設定
		        sqlCmd = "delete loanapply_period where acc_tr_type= ? ";
				updateDBSqlList.add(sqlCmd);
				dataList = new ArrayList();
				dataList.add(acc_Tr_Type);
				updateDBDataList.add(dataList);
		        updateDBSqlList.add(updateDBDataList);
		        updateDBList.add(updateDBSqlList);
			}
				
			//1.寫入table(LOANAPPLY_NCACNO)金融協助措施-會計科目檔,每個已挑選的貸款子項別各寫入一筆loanapply_ncacno
			
			sqlCmd = "INSERT INTO LOANAPPLY_NCACNO(ACC_TR_TYPE,ACC_TR_NAME,ACC_DIV,ACC_CODE,ACC_NAME,ACC_RANGE,USER_ID,USER_NAME,UPDATE_DATE) ";
			sqlCmd += "VALUES(?,?,?,?,(select distinct subitem_name from loan_subItem where subitem=?),LPAD(to_number(?), 10, 0),?,?,sysdate)";
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
					dataList.add(userid);
					dataList.add(username);
					updateDBDataList.add(dataList);
			        updateDBSqlList.add(updateDBDataList);
			        updateDBList.add(updateDBSqlList);
			        n++;
				}
			}
			if(!"".equals(selSubList2)){
				//新貸需求
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
					dataList.add(userid);
					dataList.add(username);
					updateDBDataList.add(dataList);
			        updateDBSqlList.add(updateDBDataList);
			        updateDBList.add(updateDBSqlList);
			        n++;
				}
			}
	        //2.table(LOANAPPLY_BN01)金融協助措施-需申報機構,每個已挑選的機構代碼各寫入一筆loanapply_bn01
	        if(!"".equals(selBankList)){
		        sqlCmd =  " INSERT INTO LOANAPPLY_BN01(ACC_TR_TYPE,BANK_CODE,user_id,user_name,update_date) ";
				sqlCmd += " VALUES(?,?,?,?,sysdate) ";
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
			//3.table(LOANAPPLY_PERIOD)金融協助措施-申報期間設定,每個申報基準日各寫入一筆loanapply_period
	        sqlCmd = "INSERT INTO LOANAPPLY_PERIOD(ACC_TR_TYPE,ACC_TR_NAME,APPLYTYPE,STARTDATE,ENDDATE,BEGINDATE,APPLYDATE,APPLYDATE_B,APPLYDATE_E,USER_ID,USER_NAME,UPDATE_DATE) ";
			sqlCmd += "VALUES(?,?,?,TO_DATE(?,'YYYY/MM/DD'),TO_DATE(?,'YYYY/MM/DD'),TO_DATE(?,'YYYY/MM/DD'),TO_DATE(?,'YYYY/MM/DD'),TO_DATE(?,'YYYY/MM/DD'),TO_DATE(?,'YYYY/MM/DD')-1,?,?,sysdate)";
			//申報基準日(每一筆)
			
			String[] tokens = applyDateS.split(";");
			int n = 0;
			for (String token:tokens) {
				updateDBSqlList = new LinkedList();
				updateDBDataList = new LinkedList();
				updateDBSqlList.add(sqlCmd);
				dataList = new ArrayList();
				dataList.add(acc_Tr_Type);
				dataList.add(acc_Tr_Name);
				dataList.add(applyType);
				dataList.add(transYYYYMMDD(startDate));
				dataList.add(transYYYYMMDD(endDate));
				dataList.add(transYYYYMMDD(beginDate));
				dataList.add(transYYYYMMDD(token));//UI.申報基準日(每一筆)
				if(n==0){//申報資料範圍(起) 若為第一筆申報基準日,則為UI.發布日期,其餘為前一次申報基準日
					dataList.add(transYYYYMMDD(beginDate));
				}else{
					dataList.add(transYYYYMMDD(tokens[n-1]));
				}
				dataList.add(transYYYYMMDD(token));//申報資料範圍(迄) UI.申報基準日-1 
				dataList.add(userid);
				dataList.add(username);
				updateDBDataList.add(dataList);
		        updateDBSqlList.add(updateDBDataList);
		        updateDBList.add(updateDBSqlList);
		        n++;
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
    	String applyType = Utility.getTrimString(request.getParameter("applyType"));//申報頻率
    	String startDate = Utility.getTrimString(request.getParameter("startDate"));//申報期間(起)
    	String endDate = Utility.getTrimString(request.getParameter("endDate"));//申報期間(迄)
    	String beginDate = Utility.getTrimString(request.getParameter("beginDate"));//發布日期
    	String applyDateS = Utility.getTrimString(request.getParameter("applyDateS"));//申報基準日 list
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
				//金融協助措施-會計科目檔
				sqlCmd = "delete loanapply_ncacno where acc_tr_type= ? ";
				updateDBSqlList.add(sqlCmd);
				dataList = new ArrayList();
				dataList.add(acc_Tr_Type);
				updateDBDataList.add(dataList);
		        updateDBSqlList.add(updateDBDataList);
		        updateDBList.add(updateDBSqlList);
		        
				updateDBSqlList = new LinkedList();
				updateDBDataList = new LinkedList();
				//金融協助措施-需申報機構
		        sqlCmd = "delete loanapply_bn01 where acc_tr_type= ? ";
				updateDBSqlList.add(sqlCmd);
				dataList = new ArrayList();
				dataList.add(acc_Tr_Type);
				updateDBDataList.add(dataList);
		        updateDBSqlList.add(updateDBDataList);
		        updateDBList.add(updateDBSqlList);
		        
		        updateDBSqlList = new LinkedList();
				updateDBDataList = new LinkedList();
				//金融協助措施-申報期間設定
		        sqlCmd = "delete loanapply_period where acc_tr_type= ? ";
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
	public String transYYYYMMDD(String date){
		String rtnDate = "";
		if(date.length()==7){
			rtnDate = String.valueOf(Integer.parseInt(date)+19110000);
		}
		return rtnDate;
	}
	
	

%>
