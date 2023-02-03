<%
//102.07.01 created 基本資料歷程記錄查詢列印 by 2968 
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@include file="./include/Header.include" %>
<%
	String pgId = "ZZ093W";
	System.out.println("@@ZZ093W.jsp Start...");
	String excelaction = ( request.getParameter("excelaction")==null ) ? "view" : (String)request.getParameter("excelaction");
	String duringDate = ( request.getParameter("duringDate")==null ) ? "" : (String)request.getParameter("duringDate");	
	String userId = session.getAttribute("muser_id") != null ? (String) session.getAttribute("muser_id") : "" ;
	String userName = session.getAttribute("muser_name") != null ? (String) session.getAttribute("muser_name") : "" ;
	String program_id = ( request.getParameter("program_id")==null ) ? "" : (String)request.getParameter("program_id");
	String begDate = ( request.getParameter("begDate")==null ) ? "" : (String)request.getParameter("begDate");
	String endDate = ( request.getParameter("endDate")==null ) ? "" : (String)request.getParameter("endDate");
	String updType1 = ( request.getParameter("updType1")==null ) ? "" : (String)request.getParameter("updType1");
	String updType2 = ( request.getParameter("updType2")==null ) ? "" : (String)request.getParameter("updType2");
	String updType3 = ( request.getParameter("updType3")==null ) ? "" : (String)request.getParameter("updType3");
	String updType4 = ( request.getParameter("updType4")==null ) ? "" : (String)request.getParameter("updType4");
	String updType5 = ( request.getParameter("updType5")==null ) ? "" : (String)request.getParameter("updType5");
	String pbank_no = ( request.getParameter("pbank_no")==null ) ? "" : (String)request.getParameter("pbank_no");
	String bank_no = ( request.getParameter("bank_no")==null ) ? "" : (String)request.getParameter("bank_no");
	request.setAttribute("excelaction",excelaction);
    request.setAttribute("duringDate",duringDate);
    
	if(!Utility.CheckPermission(request,report_no)){//無權限時,導向到LoginError.jsp
            rd = application.getRequestDispatcher( LoginErrorPgName ); 
	        
	    } else {
    	    actMsg = request.getAttribute("actMsg") != null ? (String) request.getAttribute("actMsg") : "";
    	    request.setAttribute("getPgName",getPgName(""));
    	    System.out.println("act= "+act);
    	    if("List".equals(act)) {
    	        rd = application.getRequestDispatcher( ListPgName +"?act=List"); 
    	    } else if("Qry".equals(act)) {
    	        request.setAttribute("QueryResult",getQueryResult(request));
    	        rd = application.getRequestDispatcher( ListPgName +"?act=Qry");
    	    } else if("Report".equals(act)) {
    	    	    request.setAttribute("ReportBody",getQueryResult(request));
    	    	    rd = application.getRequestDispatcher( ReportPgName +"?act=Report");
        	}
	    }
        request.setAttribute("queryURL","ZZ093W.jsp?act=List");
    	request.setAttribute("actMsg",actMsg); 

%>
<%@include file="./include/Tail.include" %>
<%!
	private final static String report_no = "ZZ093W" ;
	private final static String nextPgName = "/pages/message.jsp";
	private final static String ConPgName = "/pages/"+report_no+".jsp";
	private final static String EditPgName = "/pages/"+report_no+"_Edit.jsp";    
	private final static String ListPgName = "/pages/"+report_no+"_List.jsp";
	private final static String ReportPgName = "/pages/"+report_no+"_Report.jsp";
	private final static String LoginErrorPgName = "/pages/LoginError.jsp";

    
    //功能名稱選單
    private List getPgName(String id){
    	//查詢條件    
		StringBuffer sqlCmd = new StringBuffer() ;
		List paramList = new ArrayList () ;
		sqlCmd.append("select program_id,program_name || '(' || program_id || ')' as pg_name "); 
		sqlCmd.append("from wtt03_1 where program_id = ? ");
		for(int i=1;i<=15;i++){
		    sqlCmd.append("or program_id = ? ");
		}
		paramList.add("FX001W") ;
		paramList.add("FX002W") ;
		paramList.add("FX003W") ;
		paramList.add("ZZ001W") ;
		paramList.add("ZZ025W") ;
		paramList.add("BR002W") ;
		paramList.add("BR003W") ;
		paramList.add("BR004W") ;
		paramList.add("BR005W") ;
		paramList.add("BR006W") ;
		paramList.add("BR007W") ;
		paramList.add("BR008W") ;
		paramList.add("BR009W") ;
		paramList.add("BR010W") ;
		paramList.add("BR021W") ;
		paramList.add("BR022W") ;
		sqlCmd.append("order by to_number(order_type) ");
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
        System.out.println(" getPgName.size= "+dbData.size());
        return dbData;
    }
    //取得查詢資料
    private List getQueryResult(HttpServletRequest request){
        int flag = 0;
        String program_id = Utility.getTrimString(request.getParameter("program_id"));
        String begDate = Utility.getTrimString(request.getParameter("begDate"));
        String endDate = Utility.getTrimString(request.getParameter("endDate"));
        String updType1 = Utility.getTrimString(request.getParameter("updType1"));
        String updType2 = Utility.getTrimString(request.getParameter("updType2"));
        String updType3 = Utility.getTrimString(request.getParameter("updType3"));
        String updType4 = Utility.getTrimString(request.getParameter("updType4"));
        String updType5 = Utility.getTrimString(request.getParameter("updType5"));
        String bank_no = Utility.getTrimString(request.getParameter("bank_no"));
        String pbank_no = Utility.getTrimString(request.getParameter("pbank_no"));
        request.setAttribute("program_id",program_id);
        request.setAttribute("pbank_no",pbank_no);
    	request.setAttribute("bank_no",bank_no);
    	request.setAttribute("begDate",begDate);
    	request.setAttribute("endDate",endDate);
    	StringBuffer sqlCmd = new StringBuffer() ;
    	List paramList = new ArrayList() ;
    	//查詢條件   
    	sqlCmd.append(" select wtt03_1.program_name || '(' || wlxoperate_log.program_id || ')' as pg_name, ");//--功能名稱
    	sqlCmd.append(" wtt01.muser_name, ");//--使用者名稱
    	sqlCmd.append(" to_char(use_date,'yyyy/mm/dd') as use_date, ");//--使用日期
    	sqlCmd.append(" ip_address, ");//--來源IP
    	sqlCmd.append(" pbank_no, ");//--總機構
    	sqlCmd.append(" bn01.bank_name as pbank_name, ");//--總機構名稱
    	sqlCmd.append(" wlxoperate_log.bank_no, ");//--分支機構
    	sqlCmd.append(" bn02.bank_name, ");//--分支機構名稱
    	sqlCmd.append(" position_code, ");//--異動職位代碼
    	sqlCmd.append(" decode(wtt03_1.program_id,'FX001W',wlx01_m.cmuse_name,'FX002W',wlx02_m.cmuse_name,'FX003W',wlx04.cmuse_name,'') as position_name, ");//--異動職位
    	sqlCmd.append(" upd_name, ");//--異動帳號/姓名
    	sqlCmd.append(" decode(update_type,'I','新增','U','修改','D','刪除','Q','明細','R','重設密碼','P','列印','') as update_type  ");//--操作類別
    	sqlCmd.append(" from wlxoperate_log "); 
    	sqlCmd.append(" left join wtt03_1 on wlxoperate_log.program_id = wtt03_1.program_id ");
    	sqlCmd.append(" left join wtt01 on wlxoperate_log.muser_id = wtt01.muser_id ");
    	sqlCmd.append(" left join (select * from bn01 where m_year=100)bn01 on wlxoperate_log.pbank_no=bn01.bank_no ");
    	sqlCmd.append(" left join (select * from bn02 where m_year=100)bn02 on wlxoperate_log.bank_no=bn02.bank_no ");
    	sqlCmd.append(" left join (select * from cdshareno where cmuse_div='005')wlx01_m on wlxoperate_log.position_code = cmuse_id ");
    	sqlCmd.append(" left join (select * from cdshareno where cmuse_div='007')wlx02_m on wlxoperate_log.position_code = cmuse_id ");
    	sqlCmd.append(" left join (select * from cdshareno where cmuse_div='008')wlx04 on wlxoperate_log.position_code = cmuse_id ");
    	sqlCmd.append(" WHERE 1=1 ");
    	if(!"".equals(program_id) && !"ZZ093W".equals(program_id)){
    	    sqlCmd.append(" and wlxoperate_log.program_id  = ? ");
    	    paramList.add(program_id) ;
    	}
    	if(!"".equals(updType1)){
    	    if(flag==0){
    	        sqlCmd.append(" AND (update_type =? ");
    	        flag = 1;
    	    }else{
    	        sqlCmd.append(" or update_type =? ");
    	    }
    	    paramList.add(updType1) ;
    	}
    	if(!"".equals(updType2)){
    	    if(flag==0){
    	        sqlCmd.append(" AND (update_type =? ");
    	        flag = 1;
    	    }else{
    	        sqlCmd.append(" or update_type =? ");
    	    }
    	    paramList.add(updType2) ;
    	}
    	if(!"".equals(updType3)){
    	    if(flag==0){
    	        sqlCmd.append(" AND (update_type =? ");
    	        flag = 1;
    	    }else{
    	        sqlCmd.append(" or update_type =? ");
    	    }
    	    paramList.add(updType3) ;
    	}
    	if(!"".equals(updType4)){
    	    if(flag==0){
    	        sqlCmd.append(" AND (update_type =? ");
    	        flag = 1;
    	    }else{
    	        sqlCmd.append(" or update_type =? ");
    	    }
    	    paramList.add(updType4) ;
    	}
    	if(!"".equals(updType5)){
    	    if(flag==0){
    	        sqlCmd.append(" AND (update_type =? ");
    	        flag = 1;
    	    }else{
    	        sqlCmd.append(" or update_type =? ");
    	    }
    	    paramList.add(updType5) ;
    	}
    	if(!"".equals(updType1)||!"".equals(updType2)||!"".equals(updType3)||!"".equals(updType4)||!"".equals(updType5)){
    		sqlCmd.append(")  ");
    	}
    	if(!"".equals(begDate) && !"".equals(endDate)){
    	    sqlCmd.append(" AND TO_CHAR(use_date ,'yyyymmdd') BETWEEN ? AND ? ");
    	    paramList.add(begDate) ;
    	    paramList.add(endDate) ;
    	}
    	if(!"".equals(pbank_no)){
    	    sqlCmd.append(" AND pbank_no like ? ");
    	    paramList.add(pbank_no+"%") ;
    	}
    	if(!"".equals(bank_no)){
    	    sqlCmd.append(" AND bank_no like ? ");
    	    paramList.add(bank_no+"%") ;
    	}
    	sqlCmd.append(" order by use_date  desc ,wlxoperate_log.program_id ");
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
        System.out.println(" getQueryResult.size= "+dbData.size());          
        return dbData;
    }
  
    
%>
