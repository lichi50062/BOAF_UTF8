<%
//94.10.17 first designed by 4180
//96.11.12 add 增加委外項目.委外範圍 by 2295
//99.12.06 fix sqlInjection by 2808
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.List" %>
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
      System.out.println("FX006W login timeout");
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
	String seq_no = (request.getParameter("seq_no")==null)?"":(String)request.getParameter("seq_no");
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

    if(!Utility.CheckPermission(request,"FX006W")){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );
    }else{
    	if(act.equals("new")){
			rd = application.getRequestDispatcher( EditPgName +"?act=new");
    	}else if(act.equals("Edit")){
    		List dbData = getWLX06_M_OUTPUSH(bank_no,seq_no);
    	    request.setAttribute("WLX06_M_OUTPUSH",dbData);
        	request.setAttribute("maintainInfo","select * from WLX06_M_OUTPUSH WHERE bank_no='" + bank_no+"'");
        	rd = application.getRequestDispatcher( EditPgName +"?act=Edit");
    	}else if(act.equals("List")){
    		List dbData = getWLX06_M_OUTPUSH(bank_no,"");    		
    		request.setAttribute("WLX06_M_OUTPUSH",dbData);    		
        	rd = application.getRequestDispatcher( ListPgName +"?bank_no="+bank_no);    	
		}else if(act.equals("Insert")){
    	    actMsg = InsertDB(request,bank_no,lguser_id,lguser_name);
        	rd = application.getRequestDispatcher( nextPgName+"?FX=FX006W" );
    	}else if(act.equals("Update")){
    	    actMsg = UpdateDB(request,bank_no,lguser_id,lguser_name);
        	rd = application.getRequestDispatcher( nextPgName+"?FX=FX006W");
    	}else if(act.equals("Delete")){
    	    actMsg = DeleteDB(request,bank_no,lguser_id,lguser_name);
        	rd = application.getRequestDispatcher( nextPgName+"?FX=FX006W" );
    	}
    	request.setAttribute("actMsg",actMsg);
    }

	try {
        	//forward to next present jsp
        	rd.forward(request, response);
    	} catch (NullPointerException npe) { }
    }//end of doProcess
%>


<%!
    private final static String nextPgName = "/pages/ActMsg.jsp";
    private final static String EditPgName = "/pages/FX006W_Edit.jsp";
    private final static String ListPgName = "/pages/FX006W_List.jsp";
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    
    //94.04.01 add 主key更改為bank_no+seq_no by 2295
    private List getWLX06_M_OUTPUSH(String bank_no,String seq_no){
    	    List paramList = new ArrayList() ;
    		//程序為顯示畫面，查詢條件
    		String sqlCmd = " select a.*, b.countdata"
					      + " from (select out_item,seq_no,outcompanyname,outcontractname,"
						  + "			   outcontracttel,bankcomplainname,bankcomplaintel,"
						  + "			   outcomment,user_id,user_name,update_date,"
					      + "              out_begin_date,out_end_date,out_range "       
 					      + "		from WLX06_M_OUTPUSH where bank_no= ? ";
				   paramList.add(bank_no) ;
				   if(!"".equals(seq_no)) {
					   sqlCmd += " and seq_no = ?" ;
					   paramList.add(seq_no) ;
				   }
 				   //sqlCmd =(seq_no.equals(""))?sqlCmd+"":sqlCmd+" and seq_no="+seq_no;
				   sqlCmd += " 		order by out_item,seq_no)a"
						  + " 		left join (select out_item,"
						  + "						  count(*) as countdata "       
						  + "				   from WLX06_M_OUTPUSH where bank_no=?"; 
				   paramList.add(bank_no) ;
				   if(!"".equals(seq_no)) {
					   sqlCmd += " and seq_no=?" ;
					   paramList.add(seq_no) ;
				   }
				   //sqlCmd =(seq_no.equals(""))?sqlCmd+"":sqlCmd+" and seq_no="+seq_no;		  
				   sqlCmd +="					group by out_item)b on a.out_item=b.out_item";
    	
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"SEQ_NO,OUTCOMPANYNAME,OUTCONTRACTNAME,OUTCONTRACTTEL,BANKCOMPLAINNAME,BANKCOMPLAINTEL,OUTCOMMENT,USER_ID,USER_NAME,UPDATE_DATE,OUT_BEGIN_DATE,OUT_END_DATE,out_item,out_range,countdata");
            return dbData;
    }
    
    
    public String InsertDB(HttpServletRequest request,String bank_no,String lguser_id,String lguser_name) throws Exception{
    	   StringBuffer sqlCmd = new StringBuffer() ;
    	   String errMsg="";

	       String companyname = ((String)request.getParameter("companyname")==null)?"":(String)request.getParameter("companyname");
	       String contractname = ((String)request.getParameter("contractname")==null)?"":(String)request.getParameter("contractname");
           String contracttel = ((String)request.getParameter("contracttel")==null)?"":(String)request.getParameter("contracttel");	
           String complainname = ((String)request.getParameter("complainname")==null)?"":(String)request.getParameter("complainname");
	       String complaintel = ((String)request.getParameter("complaintel")==null)?"":(String)request.getParameter("complaintel");
	       String comment = ((String)request.getParameter("comment")==null)?"-":(String)request.getParameter("comment");
	       String out_beg_date = ((String)request.getParameter("BEG_DATE")==null)?"-":(String)request.getParameter("BEG_DATE");
	       String out_end_date = ((String)request.getParameter("END_DATE")==null)?"-":(String)request.getParameter("END_DATE");
	       String out_item = ((String)request.getParameter("out_item")==null)?"":(String)request.getParameter("out_item");//委外項目
	       String out_range = ((String)request.getParameter("out_range")==null)?"":(String)request.getParameter("out_range");//委外範圍
	       String user_id=lguser_id;
	       String user_name=lguser_name;

		   try {
   		  	    List updateDBSqlList = new LinkedList();
   		  	    List paramList = new ArrayList() ;
   				List max_seq_number = new LinkedList();
   			    //先去取得seq number
				sqlCmd.append("select to_char(max(SEQ_NO)+1) as maxq from WLX06_M_OUTPUSH where bank_no = ? "); 
				paramList.add(bank_no) ;
				max_seq_number = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");			  
				String max_seq =(String)((DataObject)max_seq_number.get(0)).getValue("maxq");
				max_seq=(max_seq==null)?"1":max_seq;
				sqlCmd.setLength(0) ;
				paramList.clear();
		  	    /*sqlCmd.append("INSERT INTO WLX06_M_OUTPUSH VALUES('"
		               +bank_no+"','"
		               +max_seq+"','"
		               +companyname+"','"
		               +contractname+"','"
		               +contracttel+"','"
		               +complainname+"','"
		               +complaintel+"','"
		               +comment+"','"
		               +user_id+"','"
		               +user_name
		               +"',sysdate,to_date('"+out_beg_date+"','YYYY/MM/DD'),to_date('"+out_end_date+"','YYYY/MM/DD'),'"
		               +out_item+"','"//委外項目
		               +out_range+"')");//委外範圍
			  
               updateDBSqlList.add(sqlCmd);*/
		       sqlCmd.append(" insert into WLX06_M_OUTPUSH Values( ");
		       sqlCmd.append("?,?,?,?,?,?,?,?,?,?,");
		       sqlCmd.append("sysdate,to_date(?,'YYYY/MM/DD'),to_date(?,'YYYY/MM/DD'),?,?)");
		       paramList.add(bank_no) ;
		       paramList.add(max_seq);
		       paramList.add(companyname) ;
		       paramList.add(contractname) ;
		       paramList.add(contracttel) ;
		       paramList.add(complainname) ;
		       paramList.add(complaintel) ;
		       paramList.add(comment) ;
		       paramList.add(user_id) ;
		       paramList.add(user_name) ;
		       paramList.add(out_beg_date) ;
		       paramList.add(out_end_date);
		       paramList.add(out_item) ;
		       paramList.add(out_range) ;
		       if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)) {
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

	    String seq_no = ((String)request.getParameter("editseq_no")==null)?"":(String)request.getParameter("editseq_no");
	    String companyname = ((String)request.getParameter("companyname")==null)?"":(String)request.getParameter("companyname");
	    String contractname = ((String)request.getParameter("contractname")==null)?"":(String)request.getParameter("contractname");
        String contracttel = ((String)request.getParameter("contracttel")==null)?"":(String)request.getParameter("contracttel");	
        String complainname = ((String)request.getParameter("complainname")==null)?"":(String)request.getParameter("complainname");
	    String complaintel = ((String)request.getParameter("complaintel")==null)?"":(String)request.getParameter("complaintel");
	    String comment = ((String)request.getParameter("comment")==null)?"-":(String)request.getParameter("comment");
	    String out_beg_date = ((String)request.getParameter("BEG_DATE")==null)?"-":(String)request.getParameter("BEG_DATE");
	    String out_end_date = ((String)request.getParameter("END_DATE")==null)?"-":(String)request.getParameter("END_DATE");
		String out_item = ((String)request.getParameter("out_item")==null)?"":(String)request.getParameter("out_item");//委外項目
	    String out_range = ((String)request.getParameter("out_range")==null)?"":(String)request.getParameter("out_range");//委外範圍
	       
		String user_id=lguser_id;
	    String user_name=lguser_name;

		try {
				//List updateDBSqlList = new LinkedList();
				List paramList =new ArrayList() ;
				sqlCmd.append("SELECT * FROM WLX06_M_OUTPUSH WHERE bank_no=? AND seq_no=? ");
				paramList.add(bank_no) ;
				paramList.add(seq_no) ;
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
				sqlCmd.setLength(0) ;
				paramList.clear() ;
				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
				}else{

				/*寫入log檔中*/

			    sqlCmdLog = "INSERT INTO WLX06_M_OUTPUSH_LOG "
		                  +"select bank_no,seq_no,outcompanyname,"
		                  +"outcontractname,outcontracttel,bankcomplainname,bankcomplaintel,"
		                  +"outcomment,user_id,user_name,update_date,?,?,sysdate,'U',out_begin_date,out_end_date,out_item,out_range "
		                  +" from WLX06_M_OUTPUSH where bank_no=? and seq_no=? ";
		        paramList.add(user_id) ;
		        paramList.add(user_name);
		        paramList.add(bank_no);
		        paramList.add(seq_no);
		        this.updDbUsesPreparedStatement(sqlCmdLog,paramList) ;
		        //updateDBSqlList.add(sqlCmdLog);  
		        paramList.clear() ;
		        sqlCmd.setLength(0) ;
				sqlCmd.append("UPDATE WLX06_M_OUTPUSH SET ");
				sqlCmd.append("bank_no=?");paramList.add(bank_no) ;
				sqlCmd.append(",seq_no=?");paramList.add(seq_no );
     			sqlCmd.append(",outcompanyname=?");paramList.add( companyname) ;
	    		sqlCmd.append(",outcontractname=?");paramList.add(contractname );
			    sqlCmd.append(",outcontracttel=?");paramList.add(contracttel );
				sqlCmd.append(",bankcomplainname=?");paramList.add(complainname ) ;
				sqlCmd.append(",bankcomplaintel=?");paramList.add(complaintel );
				sqlCmd.append(",outcomment=?");paramList.add(comment ) ;
				sqlCmd.append(",user_id=?");paramList.add(user_id );
				sqlCmd.append(",user_name=?");paramList.add(user_name);
				sqlCmd.append(",update_date=sysdate");
				sqlCmd.append(",out_begin_date = to_date(?,'YYYY/MM/DD')");paramList.add(out_beg_date) ;
				sqlCmd.append(",out_end_date = to_date(?,'YYYY/MM/DD')");	 paramList.add(out_end_date) ;
				sqlCmd.append(",out_item=?");paramList.add(out_item);//委外項目
				sqlCmd.append(",out_range=?");paramList.add(out_range);//委外範圍
				sqlCmd.append(" where bank_no=? and seq_no=? ");
				paramList.add(bank_no) ;
				paramList.add(seq_no);
			
		        //updateDBSqlList.add(sqlCmd);
				if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){
				//if(DBManager.updateDB(updateDBSqlList)){
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
		String seq_no = ((String)request.getParameter("editseq_no")==null)?"":(String)request.getParameter("editseq_no");
		String user_id=lguser_id;
	    String user_name=lguser_name;

		
		try {
		 		//List updateDBSqlList = new LinkedList();
		 		List paramList =new ArrayList() ;
		   	    sqlCmd.append("SELECT * FROM WLX06_M_OUTPUSH WHERE bank_no=?  AND seq_no=? ");
		   	    paramList.add(bank_no)  ;
		   	    paramList.add(seq_no) ;
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
			

				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法刪除<br>";
				}else{
					paramList.clear() ;
					sqlCmdLog = "INSERT INTO WLX06_M_OUTPUSH_LOG "
		                      +"select bank_no,seq_no,outcompanyname,"
		                      +"outcontractname,outcontracttel,bankcomplainname,bankcomplaintel,"
		                      +"outcomment,user_id,user_name,update_date,?,?,sysdate,'D',out_begin_date,out_end_date,out_item,out_range  "
							  +"from WLX06_M_OUTPUSH where bank_no=? and seq_no=? ";
		        	paramList.add(user_id) ;
		        	paramList.add(user_name) ;
		        	paramList.add(bank_no) ;
		        	paramList.add(seq_no) ;
		        	//updateDBSqlList.add(sqlCmdLog);  
 					this.updDbUsesPreparedStatement(sqlCmdLog,paramList) ;
 					paramList.clear() ;
 					sqlCmd.setLength(0) ;
				    sqlCmd.append(" delete WLX06_M_OUTPUSH where bank_no=? and seq_no=?");
				    paramList.add(bank_no) ;
				    paramList.add(seq_no) ;
		            //updateDBSqlList.add(sqlCmd);
					if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){
						errMsg = errMsg + "相關資料刪除成功";
					}else{
				   		errMsg = errMsg + "相關資料刪除失敗";
					}
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
