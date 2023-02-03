<%
//93.12.20 add 權限檢核 by 2295
//93.12.20 add 若有已點選的tbank_no,則以已點選的tbank_no為主 by 2295
//93.12.20 fix 若有已點選的bank_type,則以已點選的bank_type為主 by 2295
//93.12.23 add 超過登入時間,請重新登入 by 2295
//99.09.27 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DownLoad" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@include file="./include/Header.include" %>

<%
	String S_DATE = Utility.getTrimString(dataMap.get("S_DATE"));
	String E_DATE = Utility.getTrimString(dataMap.get("E_DATE"));
	String Report_no = Utility.getTrimString(dataMap.get("Report_no"));

    //fix 93.12.18 若有已點選的tbank_no,則以已點選的tbank_no為主============================================================
	String bank_code = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");
	String nowtbank_no =  ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");
	if(!nowtbank_no.equals("")){
	   session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session
	}
	bank_code = ( session.getAttribute("nowtbank_no")==null ) ? bank_code : (String)session.getAttribute("nowtbank_no");
	//=======================================================================================================================
	//fix 93.12.20 若有已點選的bank_type,則以已點選的bank_type為主============================================================
	String bank_type = Utility.getTrimString(dataMap.get("bank_type"));
	bank_type = ( session.getAttribute("nowbank_type")==null ) ? bank_type : (String)session.getAttribute("nowbank_type");
	//=======================================================================================================================

   if(!Utility.CheckPermission(request,report_no)){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );
    }else{
    	//set next jsp
    	if(act.equals("new")){
        	rd = application.getRequestDispatcher( QryPgName +"?bank_type="+bank_type+"&test=nothing");
    	}else if(act.equals("Query")){
    		List dbData = getData(request,S_DATE,E_DATE,bank_code,Report_no);
    		if(dbData == null || dbData.size() == 0){
    			actMsg = actMsg + "無資料可供查詢";
				request.setAttribute("actMsg",actMsg);
				rd = application.getRequestDispatcher( nextPgName );
			}else{
			   request.setAttribute("dbData",dbData);
        	   rd = application.getRequestDispatcher( ListPgName +"?act=Query&test=nothing");
    	    }
    	}
    }

%>
<%@include file="./include/Tail.include" %>


<%!
    private final static String report_no = "WMFileQuery";
    private final static String nextPgName = "/pages/ActMsg.jsp";
    private final static String ListPgName = "/pages/"+report_no+"_List.jsp";
    private final static String QryPgName = "/pages/"+report_no+"_Qry.jsp";
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";

    private List getData(HttpServletRequest request,String S_DATE,String E_DATE,String bank_code,String Report_no){
    		StringBuffer sqlCmd = new StringBuffer();
			List paramList = new ArrayList();

    		sqlCmd.append("SELECT distinct(m_year || '/' || substr(100 + m_month, 2)) as inputdate, report_no, m_year, m_month, add_date, input_method, user_name, upd_code FROM WML01 ");
    		sqlCmd.append("WHERE bank_code =?");
    		paramList.add(bank_code);
    		if (!Report_no.equals("ALL")){
    			sqlCmd.append(" AND report_no =?");
    			paramList.add(Report_no);
    		}
	    	if (!S_DATE.equals("")){
				sqlCmd.append(" AND (m_year * 100 + m_month) >= ?");
				paramList.add(S_DATE);
			}
	    	if (!E_DATE.equals("")){
			   sqlCmd.append(" AND (m_year * 100 + m_month) <=?");
			   paramList.add(E_DATE);
			}
    		sqlCmd.append(" ORDER BY report_no, m_year desc, m_month desc");
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,m_month,add_date");
            return dbData;
    }
%>    