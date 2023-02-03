<%
//97.09.18 create 簽証會計師申報 by 2295
//99.12.07 fix sqlInjection by 2808
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
<%@ page import="java.util.Enumeration" %>
<%@ page import="java.util.HashMap" %>
<%@ page import="java.util.Hashtable" %>
<%
	RequestDispatcher rd = null;
	String actMsg = "";	
	boolean doProcess = false;
	
    
	//取得session資料,取得成功時,才繼續往下執行===================================================
	if(session.getAttribute("muser_id") == null){//session timeout		
      System.out.println("FX012W login timeout");   
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
	String m_year = ( request.getParameter("m_year")==null ) ? "" : (String)request.getParameter("m_year");
	String tbank_no = ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");
	String examine = ( request.getParameter("examine")==null ) ? "" : (String)request.getParameter("examine");
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
	System.out.println("m2_name="+bank_no);
	//=======================================================================================================================

    if(!Utility.CheckPermission(request,"FX012W")){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );
    }else{
    	if(act.equals("new")){
			rd = application.getRequestDispatcher( EditPgName +"?act=new");
    	}else if(act.equals("Edit")){
    		List dbData = getBOAF_ACCOUNT(m_year,bank_no);//年度.機構代碼
    	    request.setAttribute("BOAF_ACCOUNT",dbData);
        	request.setAttribute("maintainInfo","select * from BOAF_ACCOUNT WHERE m_year="+m_year+" and bank_no='"+bank_no+"'");
        	rd = application.getRequestDispatcher( EditPgName +"?act=Edit");
    	}else if(act.equals("List")){
    		List dbData = getBOAF_ACCOUNT("",bank_no);
    		request.setAttribute("BOAF_ACCOUNT",dbData);
        	rd = application.getRequestDispatcher( ListPgName +"?bank_no="+bank_no);    	
		}else if(act.equals("Insert")){
    	    actMsg = InsertDB(request,bank_no,lguser_id,lguser_name);
        	rd = application.getRequestDispatcher( nextPgName+"?goPages=FX012W.jsp&act=List");
    	}else if(act.equals("Update")){
    	    actMsg = UpdateDB(request,bank_no,lguser_id,lguser_name);
        	rd = application.getRequestDispatcher( nextPgName+"?goPages=FX012W.jsp&act=List");
    	}else if(act.equals("Delete")){
    	    actMsg = DeleteDB(request,bank_no,lguser_id,lguser_name);
        	rd = application.getRequestDispatcher( nextPgName+"?goPages=FX012W.jsp&act=List");
		}else if(act.equals("Print")){    	    
        	rd = application.getRequestDispatcher( PrintPgName+"?m_year="+m_year+"&bank_no="+bank_no);
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
    private final static String EditPgName = "/pages/FX012W_Edit.jsp";
    private final static String ListPgName = "/pages/FX012W_List.jsp";    
    private final static String PrintPgName = "/pages/FX012W_Excel.jsp";    
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    
    /***
    * 99.12.08 fix 縣市合併年度區分
    **/
    private List getBOAF_ACCOUNT(String m_year,String bank_no){
    	List paramList = new ArrayList() ;
    	StringBuffer sqlCmd = new StringBuffer() ;
    	String s_year = "99" ;
        if(Integer.parseInt(Utility.getYear()) > 99) {
        	s_year = "100" ;
        }
		sqlCmd.append(" select BOAF_ACCOUNT.*,ba01.bank_name from BOAF_ACCOUNT "
					  + " left join (select * from ba01 where m_year=?)ba01 on BOAF_ACCOUNT.bank_no = ba01.bank_no "
					  + " where 1=1 ");
		paramList.add(s_year);
		/*String condition = "";					  
		if(!m_year.equals("")) condition += (condition.length() > 0 ? " and":"")+" m_year="+m_year;					  
		if(!bank_no.equals("")) condition += (condition.length() > 0 ? " and":"")+" BOAF_ACCOUNT.bank_no='"+bank_no+"'";					  		
		sqlCmd += condition
		       + " order by m_year,BOAF_ACCOUNT.bank_no ";
		*/		
		if(!"".equals(m_year)) {
			sqlCmd.append(" And BOAF_ACCOUNT.m_year=? ");
			paramList.add(m_year) ;
		}
		if(!"".equals(bank_no)){
			sqlCmd.append(" And BOAF_ACCOUNT.bank_no= ? ");
			paramList.add(bank_no) ;
		}
		sqlCmd.append(" order by BOAF_ACCOUNT.m_year,BOAF_ACCOUNT.bank_no") ;
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year");
        return dbData;
    }
   
    
    public String InsertDB(HttpServletRequest request,String bank_no,String lguser_id,String lguser_name) throws Exception{	
	    String sqlCmd = "";
	    String errMsg="";
	    String m_year=((String)request.getParameter("m_year")==null)?"":(String)request.getParameter("m_year");	   
	    String name=((String)request.getParameter("name")==null)?"-":(String)request.getParameter("name");			
	    String title=((String)request.getParameter("title")==null)?"-":(String)request.getParameter("title");	
	    String addr=((String)request.getParameter("addr")==null)?"-":(String)request.getParameter("addr");
	    String telno=((String)request.getParameter("telno")==null)?"-":(String)request.getParameter("telno");	    
	    String user_id=lguser_id;
	    String user_name=lguser_name;
		
		try {
   			//List updateDBSqlList = new LinkedList();
   			List paramList =new ArrayList() ;
   			sqlCmd = "SELECT * FROM BOAF_ACCOUNT WHERE m_year=? AND bank_no=? ";
   			paramList.add(m_year) ;
   			paramList.add(bank_no) ;
			List data = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");			
			paramList.clear() ;
			if (data.size() > 0){
			    errMsg = errMsg + "此筆資料已存在無法新增<br>";
			}else{
				/*sqlCmd = "INSERT INTO BOAF_ACCOUNT VALUES("
		               +m_year+",'"
		               +bank_no+"','"
		               +name+"','"
		               +title+"','"
		               +addr+"','"
		               +telno+"','"		              
		               +user_id+"','"
		               +user_name
		               +"',sysdate)";*/
		        sqlCmd = " INSERT INTO BOAF_ACCOUNT VALUES( " ;
		        sqlCmd += "?,?,?,?,?,?,?,?,sysdate )" ;
		   		paramList.add(m_year) ;
		   		paramList.add(bank_no) ;
		   		paramList.add(name) ;
		   		paramList.add(title) ;
		   		paramList.add(addr) ;
		   		paramList.add(telno);
		   		paramList.add(user_id) ;
		   		paramList.add(user_name) ;
                //updateDBSqlList.add(sqlCmd);
			    //if(DBManager.updateDB(updateDBSqlList)){
			    if(this.updDbUsesPreparedStatement(sqlCmd,paramList)) {
				   errMsg = errMsg + "相關資料寫入資料庫成功";
				}else{
			   	   errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
				}
			}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}

		return errMsg;

	}
	
	public String UpdateDB(HttpServletRequest request,String bank_no,String lguser_id,String lguser_name) throws Exception{
		String sqlCmd = "";		
		String errMsg="";
		
		String m_year = ( request.getParameter("m_year")==null ) ? "" : (String)request.getParameter("m_year");
	    String name=((String)request.getParameter("name")==null)?"-":(String)request.getParameter("name");			
	    String title=((String)request.getParameter("title")==null)?"-":(String)request.getParameter("title");	
	    String addr=((String)request.getParameter("addr")==null)?"-":(String)request.getParameter("addr");
	    String telno=((String)request.getParameter("telno")==null)?"-":(String)request.getParameter("telno");	
	    String user_id=lguser_id;
	    String user_name=lguser_name;		

		try {
				//List updateDBSqlList = new LinkedList();
				List paramList =new ArrayList() ;
				sqlCmd = "SELECT * FROM BOAF_ACCOUNT WHERE m_year=? AND bank_no=? ";
				paramList.add(m_year) ;
				paramList.add(bank_no) ;
			    List data = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
			    paramList.clear() ;
				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
				}else{
				    /*寫入log檔中*/
				    sqlCmd = "INSERT INTO BOAF_ACCOUNT_LOG "
		             	   +"select m_year,bank_no,name,title,addr,telno,"
		             	   +"user_id,user_name,update_date,?,?,sysdate,'U' from BOAF_ACCOUNT WHERE m_year=? AND bank_no=? ";
		        	paramList.add(user_id) ;
		        	paramList.add(user_name) ;
		        	paramList.add(m_year) ;
		        	paramList.add(bank_no) ;
		            //updateDBSqlList.add(sqlCmd); 
		          	
		          	if(!this.updDbUsesPreparedStatement(sqlCmd,paramList)) {						
				   		errMsg = errMsg + "BOAF_ACCOUNT_LOG相關資料寫入資料庫失敗" ;
					}
		          	paramList.clear() ;
				    sqlCmd = "UPDATE BOAF_ACCOUNT SET "				       	   
		               	   +" name=?"
		               	   +",title=?"
		               	   +",addr=?"
		               	   +",telno=?"		                  
				       	   +",user_id=?"
   					   	   +",user_name=?,update_date=sysdate"
   			           	   +" where m_year=? AND bank_no=?";
					paramList.add(name) ;
					paramList.add(title) ;
					paramList.add(addr) ;
					paramList.add(telno) ;
					paramList.add(user_id) ;
					paramList.add(user_name) ;
					paramList.add(m_year) ;
					paramList.add(bank_no) ;
		            //updateDBSqlList.add(sqlCmd);
					//if(DBManager.updateDB(updateDBSqlList)){
					if(this.updDbUsesPreparedStatement(sqlCmd,paramList)) {
						errMsg = errMsg + "相關資料寫入資料庫成功";
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗" ;
					}
    	   		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";
		}

		return errMsg;
	}
    
    public String DeleteDB(HttpServletRequest request,String bank_no,String lguser_id,String lguser_name) throws Exception{
		String sqlCmd = "";		
		String errMsg="";
		String m_year = ( request.getParameter("m_year")==null ) ? "" : (String)request.getParameter("m_year");			    
		String user_id=lguser_id;
	  	String user_name=lguser_name;
		

		try {
			 	//List updateDBSqlList = new LinkedList();
			 	List paramList = new ArrayList() ;
				sqlCmd = "SELECT * FROM BOAF_ACCOUNT WHERE m_year=? AND bank_no=? ";
				paramList.add(m_year) ;
				paramList.add(bank_no);
				
			    List data = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
				paramList.clear() ;
				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法刪除<br>";
				}else{
				    /*寫入log檔中*/
				     sqlCmd = "INSERT INTO BOAF_ACCOUNT_LOG "
		             	   +"select m_year,bank_no,name,title,addr,telno,"
		             	   +"user_id,user_name,update_date,?,?,sysdate,'D' from BOAF_ACCOUNT WHERE m_year=? AND bank_no=? ";
		        	paramList.add(user_id) ;
		        	paramList.add(user_name) ;
		        	paramList.add(m_year) ;
		        	paramList.add(bank_no) ;
		        	//updateDBSqlList.add(sqlCmd); 
 					
 					if(!this.updDbUsesPreparedStatement(sqlCmd,paramList)) {						
				   		errMsg = errMsg + "BOAF_ACCOUNT_LOG相關資料寫入資料庫失敗" ;
					}
 					paramList.clear() ;
				    sqlCmd = " delete BOAF_ACCOUNT where m_year=? AND bank_no=? ";
				    paramList.add(m_year) ;
				    paramList.add(bank_no) ;
		            //updateDBSqlList.add(sqlCmd);

					//if(DBManager.updateDB(updateDBSqlList)){
					if(this.updDbUsesPreparedStatement(sqlCmd,paramList)) {
						errMsg = errMsg + "相關資料刪除成功";
					}else{
				   		errMsg = errMsg + "相關資料刪除失敗" ;
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
