<%
//94.04.15 add 申報資料批次刪除 by 2295
//94.04.20 add 刪除WML01檢核結果 by 2295
//94.09.19 add 刪除時一併刪除WML01_LOCK(申報鎖住檔)/WML02(更新錯誤紀錄檔)/WML03(更新錯誤紀錄_描述檔) by 2295
//94.11.09 fix 抓不到月份的問題 by 2295
//95.09.08 add A06批次刪除 by 2295
//96.07.11 add A08批次刪除 by 2295
//97.01.03 add A09批次刪除 by 2295
//97.06.13 add A10批次刪除 by 2295
//99.10.08 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DownLoad" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="java.util.*" %>

<%@include file="./include/Header.include" %>
<%

	String S_YEAR = Utility.getTrimString(dataMap.get("S_YEAR"));
	String S_MONTH = Utility.getTrimString(dataMap.get("S_MONTH"));
	String Report_no = Utility.getTrimString(dataMap.get("Report_no"));
	//fix 93.12.18 若有已點選的tbank_no,則以已點選的tbank_no為主============================================================
	String bank_code = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");
	String nowtbank_no = Utility.getTrimString(dataMap.get("tbank_no"));
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session
	}
	bank_code = ( session.getAttribute("nowtbank_no")==null ) ? bank_code : (String)session.getAttribute("nowtbank_no");
	//=======================================================================================================================
    //fix 93.12.20 若有已點選的bank_type,則以已點選的bank_type為主============================================================
	String bank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");
	bank_type = ( session.getAttribute("nowbank_type")==null ) ? bank_type : (String)session.getAttribute("nowbank_type");
	//=======================================================================================================================
	String userid = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");
	String username = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");

	StringBuffer sqlCmd = new StringBuffer();
	List paramList = new ArrayList();
	StringBuffer sqlCmd_Qry = new StringBuffer();
	List paramList_Qry = new ArrayList();
	StringBuffer sqlCmd_Del = new StringBuffer();
	List paramList_Del = new ArrayList();
	StringBuffer condition = new StringBuffer();
	List paramList_condition = new ArrayList();
	String BankList = "";
    List BankList_data = null;
    List queryData = null;
    String queryCount = "0";
    String selectBank_no = "";//選取的金融機構代號

	System.out.println("WMFileDeleteBatch.act="+act);
	S_MONTH = String.valueOf(Integer.parseInt(S_MONTH));

    if(!Utility.CheckPermission(request,report_no)){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );
    }else{
    	//set next jsp
    	//將選取的金融機構代碼.縣市別寫到session=======================================
    	if(request.getParameter("BankList")	!= null && !((String)request.getParameter("BankList")).equals("")){
    	   session.setAttribute("BankList",(String)request.getParameter("BankList"));
    	}
    	if(request.getParameter("HSIEN_ID")	!= null && !((String)request.getParameter("HSIEN_ID")).equals("")){
           session.setAttribute("HSIEN_ID",(String)request.getParameter("HSIEN_ID"));
        }
    	if(act.equals("new")){
        	rd = application.getRequestDispatcher( QryPgName+"?bank_type="+bank_type );
    	}else if(act.equals("Delete")){
    	    //金融機構
			if(session.getAttribute("BankList") != null && !((String)session.getAttribute("BankList")).equals("")){
		   		BankList = (String)session.getAttribute("BankList");
		   		BankList_data = Utility.getReportData(BankList);
		   		System.out.println("BankList_data.size()="+BankList_data.size());
			}
    		List data=null;

    		if(Report_no.equals("A06")){//95.09.08 add A06
    		   sqlCmd_Qry.append(" select count(a.amt_3month) as havecount from "+Report_no+" a ");
    		}else if(Report_no.equals("A08")){//96.07.11 add A08
    		   sqlCmd_Qry.append(" select count(a.warnaccount_cnt) as havecount from "+Report_no+" a ");
    		}else if(Report_no.equals("A09")){//97.01.03 add A09
    		   sqlCmd_Qry.append(" select count(a.over_cnt) as havecount from "+Report_no+" a ");
    		}else if(Report_no.equals("A10")){//97.06.13 add A10
    		   sqlCmd_Qry.append(" select count(a.loan2_amt) as havecount from "+Report_no+" a ");
    		}else{
    		   sqlCmd_Qry.append(" select count(a.amt) as havecount from "+Report_no+" a ");
    		}

    		condition.append(" LEFT JOIN wml01_a_v c on a.bank_code = c.bank_code and ");
			condition.append("				 		      a.m_year  = c.m_year and ");
			condition.append("					  		  a.m_month = c.m_month and ");
			condition.append("							c.report_no = ?");
			condition.append(" where a.m_year=? and a.m_month=?");
			paramList_condition.add(Report_no);
			paramList_condition.add(S_YEAR);
			paramList_condition.add(S_MONTH);
			sqlCmd.append("select bank_code,wml01_lock_status,wml01_lock_lock_status from wml01_a_v ");
			sqlCmd.append(" where m_year=? and m_month=?");
			sqlCmd.append(" and   report_no = ?");
			sqlCmd.append(" and  (wml01_lock_status is not null and wml01_lock_status = 'Y'");
            sqlCmd.append("        or wml01_lock_lock_status is not null and wml01_lock_lock_status = 'Y')");
            paramList.add(S_YEAR);
            paramList.add(S_MONTH);
            paramList.add(Report_no);

			List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
			List bankcode_lock = new LinkedList();
			List bankcode = new LinkedList();
			boolean addCheck=false;
    		//金融機構代號=============================================================
            if(BankList_data != null && BankList_data.size() != 0){
               addCheck=true;
               for(int i=0;i<BankList_data.size();i++){
                   //檢核有無被鎖定,鎖定時不可刪除====================================
                   if(dbData != null && dbData.size() != 0){
                      addCheck=true;
                      for(int j=0;j<dbData.size();j++){
                          if(((String)((DataObject)dbData.get(j)).getValue("bank_code")).equals((String)((List)BankList_data.get(i)).get(0))){
                             bankcode_lock.add((String)((DataObject)dbData.get(j)).getValue("bank_code"));
                             addCheck=false;
                          }
                      }
                   }//==============================================================
                   if(addCheck){//未被鎖定
                      selectBank_no +="'"+(String)((List)BankList_data.get(i)).get(0)+"'";
                      bankcode.add((String)((List)BankList_data.get(i)).get(0));
                      if(i < BankList_data.size()-1) selectBank_no +=",";
                   }
               }
            }
            //==============================================================================
            condition.append(" and a.bank_code IN ("+selectBank_no+")");
            if(!Report_no.equals("A08") && !Report_no.equals("A09") && !Report_no.equals("A10")) sqlCmd.append(" order by acc_code ");

            sqlCmd_Qry.append(condition);
            for(int i=0;i<paramList_condition.size();i++){
                paramList_Qry.add(paramList_condition.get(i));
            }
           
            
            data = DBManager.QueryDB_SQLParam(sqlCmd_Qry.toString(),paramList_Qry,"havecount");

	    	System.out.println("data.size()="+data.size());
	    	//94.04.18 顯示已被鎖定無法刪除的bank_code
	    	if(bankcode_lock.size() != 0){
	    	   actMsg +="機構代號:";
	    	   for(int i=0;i<bankcode_lock.size();i++){
	    	      actMsg += (String)bankcode_lock.get(i);
	    	      if(i < bankcode_lock.size()-1) actMsg +=",";
	    	   }
	    	   actMsg += "已被鎖定無法刪除<br>";
	    	}
    		if(((((DataObject)data.get(0)).getValue("havecount")).toString()).equals("0")){
    		   actMsg = actMsg + "無資料可執行批次刪除";
    		}else{
    		   List updateDBList = new ArrayList();//0:sql 1:data		
	    	   List updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
			   List updateDBDataList = new ArrayList();//儲存參數的List
			   List dataList =  new ArrayList();//儲存參數的data
			   
    		   sqlCmd.delete(0,sqlCmd.length());
    		   paramList = new ArrayList();
    		   sqlCmd.append("select count(*) as count from "+Report_no + " a" + " where a.m_year=? and a.m_month=? and a.bank_code IN ("+selectBank_no+")");
    		   paramList.add(S_YEAR);
    		   paramList.add(S_MONTH);
    		   queryData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"count");
    		   queryCount = (((DataObject)queryData.get(0)).getValue("count")).toString();
    		   if(!queryCount.equals("0")){
    		      sqlCmd_Del.append(" delete from "+Report_no + " a");
    		      sqlCmd_Del.append(" where a.m_year=? and a.m_month=?");
    		      sqlCmd_Del.append(" and a.bank_code IN ("+selectBank_no+")");
    		      dataList.add(S_YEAR);
    		      dataList.add(S_MONTH);
				  updateDBDataList.add(dataList);
				  updateDBSqlList.add(sqlCmd_Del.toString());//0:欲執行的sql	
				  updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				  updateDBList.add(updateDBSqlList);		          
		       }

		       //94.04.20 add 刪除WML01(申報紀錄檔)檢核結果====================================================
		       sqlCmd.delete(0,sqlCmd.length());
    		   paramList = new ArrayList();
    		   updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
			   updateDBDataList = new ArrayList();//儲存參數的List
			   dataList =  new ArrayList();//儲存參數的data
    		   sqlCmd.append("select count(*) as count from WML01 WHERE m_year=? AND m_month=? AND bank_code IN (" +selectBank_no+") AND report_no=?");
    		   paramList.add(S_YEAR);
    		   paramList.add(S_MONTH);
    		   paramList.add(Report_no);
		       queryData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"count");
		       queryCount = (((DataObject)queryData.get(0)).getValue("count")).toString();
		       System.out.println("WML01.size()="+queryCount);
    		   if(!queryCount.equals("0")){
    		      sqlCmd_Del.delete(0,sqlCmd_Del.length());
			      sqlCmd_Del.append("INSERT INTO WML01_LOG ");
			   	  sqlCmd_Del.append(" select m_year,m_month,bank_code,report_no,input_method,add_user,add_name,add_date,common_center,upd_method,upd_code ");
			   	  sqlCmd_Del.append(",batch_no,lock_status,user_id,user_name,update_date,?,?,sysdate,'D'");
			   	  sqlCmd_Del.append(" FROM WML01 WHERE m_year=? AND m_month=?");
			   	  sqlCmd_Del.append(" AND bank_code IN (" +selectBank_no+") AND report_no=?");
			      
			      dataList.add(userid);
			      dataList.add(username);
			      dataList.add(S_YEAR);
			      dataList.add(S_MONTH);
    		      dataList.add(Report_no);
				  updateDBDataList.add(dataList);
				  
				  updateDBSqlList.add(sqlCmd_Del.toString());//0:欲執行的sql	
				  updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				  updateDBList.add(updateDBSqlList);		  			
					
				  
    		   	  updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
			   	  updateDBDataList = new ArrayList();//儲存參數的List
			      dataList =  new ArrayList();//儲存參數的data
				  sqlCmd_Del.delete(0,sqlCmd_Del.length());
				  
			      sqlCmd_Del.append("DELETE FROM WML01 WHERE m_year=? AND m_month=? AND ");
				  sqlCmd_Del.append("bank_code IN (" +selectBank_no+") AND report_no=?");
				  			      
			      dataList.add(S_YEAR);
			      dataList.add(S_MONTH);
    		      dataList.add(Report_no);
				  updateDBDataList.add(dataList);
				  
				  updateDBSqlList.add(sqlCmd_Del.toString());//0:欲執行的sql	
				  updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				  updateDBList.add(updateDBSqlList);	
		       }
		       //94.09.19 add 刪除WML02(更新錯誤紀錄檔)檢核結果====================================================
		       sqlCmd.delete(0,sqlCmd.length());    		   
    		   sqlCmd.append("select count(*) as count from WML02 WHERE m_year=? AND m_month=? AND bank_code IN (" +selectBank_no+") AND report_no=?");
    		   
    		   
		       queryData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"count");
		       queryCount = (((DataObject)queryData.get(0)).getValue("count")).toString();
		       System.out.println("WML02.size()="+queryCount);
    		   if(!queryCount.equals("0")){
    		      updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
			   	  updateDBDataList = new ArrayList();//儲存參數的List
			      dataList =  new ArrayList();//儲存參數的data
				  sqlCmd_Del.delete(0,sqlCmd_Del.length());
		          sqlCmd_Del.append(" INSERT INTO WML02_LOG ");
				  sqlCmd_Del.append(" select m_year,m_month,bank_code,report_no,cano,l_amt,r_amt,user_id,user_name,update_date");
				  sqlCmd_Del.append(",user_id,user_name,sysdate,'D'");
				  sqlCmd_Del.append(" from WML02 where m_year=? AND m_month=?");
				  sqlCmd_Del.append(" AND bank_code IN (" +selectBank_no+") AND report_no=?");
				 
			      dataList.add(S_YEAR);
			      dataList.add(S_MONTH);
    		      dataList.add(Report_no);
				  updateDBDataList.add(dataList);
				  
				  updateDBSqlList.add(sqlCmd_Del.toString());//0:欲執行的sql	
				  updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				  updateDBList.add(updateDBSqlList);				  
			      
				  
				  updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
			   	  updateDBDataList = new ArrayList();//儲存參數的List
			   	  sqlCmd_Del.delete(0,sqlCmd_Del.length());	
			      sqlCmd_Del.append("DELETE FROM WML02 WHERE m_year=? AND m_month=? AND ");
				  sqlCmd_Del.append("bank_code IN (" +selectBank_no+") AND report_no=?");
			     
				  updateDBDataList.add(dataList);
				  
				  updateDBSqlList.add(sqlCmd_Del.toString());//0:欲執行的sql	
				  updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				  updateDBList.add(updateDBSqlList);	
			   }
			   //94.09.19 add 刪除WML03(更新錯誤紀錄_描述檔)檢核結果====================================================
			   sqlCmd.delete(0,sqlCmd.length());   
			   sqlCmd.append("select count(*) as count from WML03 WHERE m_year=? AND m_month=? AND bank_code IN (" +selectBank_no+") AND report_no=?");
			   queryData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"count");
			   queryCount = (((DataObject)queryData.get(0)).getValue("count")).toString();
		       System.out.println("WML03.size()="+queryCount);
    		   if(!queryCount.equals("0")){
    		      updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
			   	  updateDBDataList = new ArrayList();//儲存參數的List
			   	  sqlCmd_Del.delete(0,sqlCmd_Del.length());	
			      sqlCmd_Del.append(" INSERT INTO WML03_LOG ");
				  sqlCmd_Del.append(" select m_year,m_month,bank_code,report_no,serial_no,remark,user_id,user_name,update_date");
				  sqlCmd_Del.append(",user_id,user_name,sysdate,'D'");
				  sqlCmd_Del.append(" from WML03 where m_year=? AND m_month=?");
				  sqlCmd_Del.append(" AND bank_code IN (" +selectBank_no+") AND report_no=?");
			      updateDBDataList.add(dataList);
				  
				  updateDBSqlList.add(sqlCmd_Del.toString());//0:欲執行的sql	
				  updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				  updateDBList.add(updateDBSqlList);
				
				  updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
			   	  updateDBDataList = new ArrayList();//儲存參數的List
			   	  sqlCmd_Del.delete(0,sqlCmd_Del.length());	
			      sqlCmd_Del.append("DELETE FROM WML03 WHERE m_year=? AND m_month=? AND ");
				  sqlCmd_Del.append("bank_code IN (" +selectBank_no+") AND report_no=?");
			      updateDBDataList.add(dataList);
				  
				  updateDBSqlList.add(sqlCmd_Del.toString());//0:欲執行的sql	
				  updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				  updateDBList.add(updateDBSqlList);
			   }
			   //94.09.19 add 刪除WML01_LOCK(申報鎖住檔)檢核結果====================================================
			   sqlCmd.delete(0,sqlCmd.length());   
			   sqlCmd.append("select count(*) as count from WML01_LOCK WHERE m_year=? AND m_month=? AND " + "bank_code IN (" +selectBank_no+") AND report_no=?");
			   queryData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"count");
			   queryCount = (((DataObject)queryData.get(0)).getValue("count")).toString();
		       System.out.println("WML01_LOCK.size()="+queryCount);
    		   if(!queryCount.equals("0")){
    		      updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
			   	  updateDBDataList = new ArrayList();//儲存參數的List
			   	  sqlCmd_Del.delete(0,sqlCmd_Del.length());	
			      sqlCmd_Del.append(" DELETE FROM WML01_LOCK WHERE m_year=? AND m_month=? AND ");
				  sqlCmd_Del.append("bank_code IN (" +selectBank_no+") AND report_no=?");
			      updateDBDataList.add(dataList);
				  
				  updateDBSqlList.add(sqlCmd_Del.toString());//0:欲執行的sql	
				  updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				  updateDBList.add(updateDBSqlList);
			   }
		       //94.04.18 顯示已被刪除成功的bank_code =======================================
		       if(bankcode.size() != 0){
	    	      actMsg +="機構代號:";
	    	      for(int i=0;i<bankcode.size();i++){
	    	          actMsg += (String)bankcode.get(i);
	    	          if(i < bankcode.size()-1) actMsg +=",";
	    	      }
	    	   }
		       //============================================================================
			   if(DBManager.updateDB_ps(updateDBList)){
				  actMsg = actMsg + "相關資料刪除成功";
			   }else{
				  actMsg = actMsg + "相關資料刪除失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
			   }
    		}
    		request.setAttribute("actMsg",actMsg);
    		rd = application.getRequestDispatcher( nextPgName );
    	}
    }

%>

<%@include file="./include/Tail.include" %>

<%!
    private final static String report_no = "WMFileDeleteBatch";
    private final static String nextPgName = "/pages/ActMsg.jsp";
    private final static String QryPgName = "/pages/"+report_no+"_Qry.jsp";
    private final static String LoginErrorPgName = "/pages/LoginError.jsp"; 
%>    