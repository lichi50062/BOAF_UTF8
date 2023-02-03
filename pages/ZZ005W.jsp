<%
//93.12.29 create by 2295
//94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除 by 2295
//94.02.01 fix sub_type/order_type為MIS系統專用
//99.12.09 fix sqlInjection by 2808
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
<%@ page import="java.util.StringTokenizer" %>

<%
	RequestDispatcher rd = null;
	String actMsg = "";
	String alertMsg = "";
	String webURL = "";
	boolean doProcess = false;

	//取得session資料,取得成功時,才繼續往下執行===================================================
	if (session.getAttribute("muser_id") == null) {//session timeout	
		System.out.println("ZZ005W login timeout");
		rd = application
				.getRequestDispatcher("/pages/reLogin.jsp?url=LoginError.jsp?timeout=true");
		try {
			rd.forward(request, response);
		} catch (Exception e) {
			System.out.println("forward Error:" + e + e.getMessage());
		}
	} else {
		doProcess = true;
	}
	if (doProcess) {//若muser_id資料時,表示登入成功====================================================================	

		String act = (request.getParameter("act") == null) ? ""
				: (String) request.getParameter("act");
		String program_id = (request.getParameter("program_id") == null) ? ""
				: (String) request.getParameter("program_id");

		System.out.println("act=" + act);

		//登入者資訊
		String lguser_id = (session.getAttribute("muser_id") == null) ? ""
				: (String) session.getAttribute("muser_id");
		String lguser_name = (session.getAttribute("muser_name") == null) ? ""
				: (String) session.getAttribute("muser_name");
		String lguser_type = (session.getAttribute("muser_type") == null) ? ""
				: (String) session.getAttribute("muser_type");
		session.setAttribute("nowtbank_no", null);//94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除======

		if (!CheckPermission(request)) {//無權限時,導向到LoginError.jsp
			rd = application.getRequestDispatcher(LoginErrorPgName);
		} else {
			//set next jsp 	    	
			if (act.equals("new")) {
				rd = application.getRequestDispatcher(EditPgName
						+ "?act=new");
			} else if (act.equals("List") || act.equals("Qry")
					|| act.equals("del")) {
				rd = application.getRequestDispatcher(ListPgName
						+ "?act=" + act);
			} else if (act.equals("Edit")) {
				List WTT03_1 = getWTT03_1(program_id);
				request.setAttribute("WTT03_1", WTT03_1);
				//93.12.20設定異動者資訊======================================================================
				request.setAttribute("maintainInfo",
						"select * from WTT03_1 WHERE program_id='"
								+ program_id + "'");
				//=======================================================================================================================		        	     	   	        	    
				rd = application.getRequestDispatcher(EditPgName
						+ "?act=Edit");
			} else if (act.equals("Insert")) {
				actMsg = InsertDB(request, lguser_id, lguser_name);
				rd = application.getRequestDispatcher(nextPgName);
			} else if (act.equals("Update")) {
				actMsg = UpdateDB(request, program_id, lguser_id,
						lguser_name);
				rd = application.getRequestDispatcher(nextPgName);
			} else if (act.equals("Delete")) {
				actMsg = DeleteDB(request);
				rd = application.getRequestDispatcher(nextPgName);
			}
			request.setAttribute("actMsg", actMsg);
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


<%!private final static String nextPgName = "/pages/ActMsg.jsp";
	private final static String EditPgName = "/pages/ZZ005W_Edit.jsp";
	private final static String ListPgName = "/pages/ZZ005W_List.jsp";
	private final static String LoginErrorPgName = "/pages/LoginError.jsp";

	private boolean CheckPermission(HttpServletRequest request) {//檢核權限    	    
		boolean CheckOK = false;
		HttpSession session = request.getSession();
		Properties permission = (session.getAttribute("ZZ005W") == null) ? new Properties()
				: (Properties) session.getAttribute("ZZ005W");
		if (permission == null) {
			System.out.println("ZZ005W.permission == null");
		} else {
			System.out.println("ZZ005W.permission.size =" + permission.size());

		}
		//只要有Query的權限,就可以進入畫面
		if (permission != null && permission.get("Q") != null
				&& permission.get("Q").equals("Y")) {
			CheckOK = true;//Query
		}
		return CheckOK;
	}

	//取得WTT03_1該程式資料
	private List getWTT03_1(String program_id) {
		//查詢條件    
		List paramList = new ArrayList();
		String sqlCmd = " select * from WTT03_1 " + " where program_id=? ";
		paramList.add(program_id);
		List dbData = DBManager.QueryDB_SQLParam(sqlCmd, paramList, "");
		return dbData;
	}

	public String InsertDB(HttpServletRequest request, String lguser_id,
			String lguser_name) throws Exception {
		String sqlCmd = "";
		String errMsg = "";

		String program_id = ((String) request.getParameter("PROGRAM_ID") == null) ? ""
				: (String) request.getParameter("PROGRAM_ID");
		String program_name = ((String) request.getParameter("PROGRAM_NAME") == null) ? ""
				: (String) request.getParameter("PROGRAM_NAME");
		String url_id = ((String) request.getParameter("URL_ID") == null) ? ""
				: (String) request.getParameter("URL_ID");
		String p_add = ((String) request.getParameter("P_ADD") == null) ? "N"
				: (String) request.getParameter("P_ADD");
		String p_delete = ((String) request.getParameter("P_DELETE") == null) ? "N"
				: (String) request.getParameter("P_DELETE");
		String p_update = ((String) request.getParameter("P_UPDATE") == null) ? "N"
				: (String) request.getParameter("P_UPDATE");
		String p_query = ((String) request.getParameter("P_QUERY") == null) ? "N"
				: (String) request.getParameter("P_QUERY");
		String p_print = ((String) request.getParameter("P_PRINT") == null) ? "N"
				: (String) request.getParameter("P_PRINT");
		String p_upload = ((String) request.getParameter("P_UPLOAD") == null) ? "N"
				: (String) request.getParameter("P_UPLOAD");
		String p_download = ((String) request.getParameter("P_DOWNLOAD") == null) ? "N"
				: (String) request.getParameter("P_DOWNLOAD");
		String p_lock = ((String) request.getParameter("P_LOCK") == null) ? "N"
				: (String) request.getParameter("P_LOCK");
		String p_other = ((String) request.getParameter("P_OTHER") == null) ? "N"
				: (String) request.getParameter("P_OTHER");
		String input_order = ((String) request.getParameter("INPUT_ORDER") == null) ? " "
				: (String) request.getParameter("INPUT_ORDER");
		String sub_type = ((String) request.getParameter("SUB_TYPE") == null) ? " "
				: (String) request.getParameter("SUB_TYPE");
		String order_type = ((String) request.getParameter("ORDER_TYPE") == null) ? " "
				: (String) request.getParameter("ORDER_TYPE");
		String user_id = lguser_id;
		String user_name = lguser_name;

		try {
			// List updateDBSqlList = new LinkedList();
			List paramList = new ArrayList();
			sqlCmd = "SELECT * FROM WTT03_1 " + " WHERE program_id=? ";
			paramList.add(program_id);
			List data = DBManager.QueryDB_SQLParam(sqlCmd, paramList, "");			
			paramList.clear();
			System.out.println("data.size=" + data.size());
			if (data.size() != 0) {
				errMsg = errMsg + "此筆程式代碼已存在無法新增<br>";
			} else {
				/* sqlCmd = "INSERT INTO WTT03_1 VALUES ('"+program_id+"'" 
				 	   + ",'"+ program_name + "'"
				    	   + ",'"+ url_id + "'"			           	      
				       + ",'" + p_add + "'" 
				       + ",'" + p_delete + "'" 
				       + ",'" + p_update + "'" 
				       + ",'" + p_query + "'" 
				       + ",'" + p_print + "'" 
				       + ",'" + p_upload + "'" 
				       + ",'" + p_download + "'" 
				       + ",'" + p_lock + "'" 
				       + ",'" + p_other + "'" 
				       + ",'" + input_order + "'" 
				       + ",'" + sub_type + "'" 
				       + ",'" + order_type + "'" 
				       + ",'" + user_id +"'"					       
				       + ",'" + user_name + "'"
				       + ",sysdate)"; 		*/
				sqlCmd = " insert into WTT03_1 Values( ?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,sysdate)";
				paramList.add(program_id);
				paramList.add(program_name);
				paramList.add(url_id);
				paramList.add(p_add);
				paramList.add(p_delete);
				paramList.add(p_update);
				paramList.add(p_query);
				paramList.add(p_print);
				paramList.add(p_upload);
				paramList.add(p_download);
				paramList.add(p_lock);
				paramList.add(p_other);
				paramList.add(input_order);
				paramList.add(sub_type);
				paramList.add(order_type);
				paramList.add(user_id);
				paramList.add(user_name);
				if (!this.updDbUsesPreparedStatement(sqlCmd, paramList)) {
					throw new Exception();
				}
			}
			errMsg = "相關資料寫入資料庫成功";
			// updateDBSqlList.add(sqlCmd); 		            	   				  

			/*if(DBManager.updateDB(updateDBSqlList)){					 
			   errMsg = errMsg + "相關資料寫入資料庫成功";					
			}else{
			   errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
			}*/
		} catch (Exception e) {
			System.out.println(e + ":" + e.getMessage());
			errMsg = errMsg + "相關資料寫入資料庫失敗";
		}

		return errMsg;
	}

	public String UpdateDB(HttpServletRequest request,String program_id,String lguser_id,String lguser_name) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer();		
		String errMsg="";				
	    String program_name=((String)request.getParameter("PROGRAM_NAME")==null)?"":(String)request.getParameter("PROGRAM_NAME");
	    String url_id=((String)request.getParameter("URL_ID")==null)?"":(String)request.getParameter("URL_ID");
	    String p_add=((String)request.getParameter("P_ADD")==null)?"N":(String)request.getParameter("P_ADD");
	    String p_delete=((String)request.getParameter("P_DELETE")==null)?"N":(String)request.getParameter("P_DELETE");
	    String p_update=((String)request.getParameter("P_UPDATE")==null)?"N":(String)request.getParameter("P_UPDATE");
	    String p_query=((String)request.getParameter("P_QUERY")==null)?"N":(String)request.getParameter("P_QUERY");
	    String p_print=((String)request.getParameter("P_PRINT")==null)?"N":(String)request.getParameter("P_PRINT");
	    String p_upload=((String)request.getParameter("P_UPLOAD")==null)?"N":(String)request.getParameter("P_UPLOAD");
	    String p_download=((String)request.getParameter("P_DOWNLOAD")==null)?"N":(String)request.getParameter("P_DOWNLOAD");
	    String p_lock=((String)request.getParameter("P_LOCK")==null)?"N":(String)request.getParameter("P_LOCK");
	    String p_other=((String)request.getParameter("P_OTHER")==null)?"N":(String)request.getParameter("P_OTHER");
	    String input_order=((String)request.getParameter("INPUT_ORDER")==null)?" ":(String)request.getParameter("INPUT_ORDER");
	    String sub_type=((String)request.getParameter("SUB_TYPE")==null)?" ":(String)request.getParameter("SUB_TYPE");
	    String order_type=((String)request.getParameter("ORDER_TYPE")==null)?" ":(String)request.getParameter("ORDER_TYPE");
	    String user_id=lguser_id;
	    String user_name=lguser_name;			   
		try {			    			    
			    //List updateDBSqlList = new LinkedList();
				List paramList = new ArrayList() ;
				sqlCmd.append( "SELECT * FROM WTT03_1 "
				    	+ " WHERE program_id=? ");
				paramList.add(program_id) ;
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");	
			    paramList.clear() ;
			    sqlCmd.setLength(0) ;
				System.out.println("data.size="+data.size());
				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
				}else{    				   
				    sqlCmd.append("UPDATE WTT03_1 SET ") ;
				    sqlCmd.append("program_name=?");paramList.add( program_name  ) ;
		            sqlCmd.append(",url_id=?");paramList.add( url_id );			           	      
				    sqlCmd.append(",p_add=?"); paramList.add( p_add  ) ; 
				    sqlCmd.append(",p_delete=?"); paramList.add( p_delete); 
				    sqlCmd.append(",p_update=?"); paramList.add( p_update); 
				    sqlCmd.append(",p_query=?"); paramList.add(p_query); 
				    sqlCmd.append(",p_print=?"); paramList.add( p_print); 
				    sqlCmd.append(",p_upload=?"); paramList.add( p_upload); 
				    sqlCmd.append(",p_download=?"); paramList.add( p_download ); 
				    sqlCmd.append(",p_lock=?"); paramList.add(p_lock ); 
				    sqlCmd.append(",p_other=?"); paramList.add( p_other); 
				    sqlCmd.append(",input_order=?"); paramList.add( input_order ); 
				    sqlCmd.append(",sub_type=?"); paramList.add( sub_type ); 
				    sqlCmd.append(",order_type=?"); paramList.add( order_type); 
				    sqlCmd.append(",user_id=?"); paramList.add( user_id );					       
				    sqlCmd.append(",user_name=?"); paramList.add( user_name );
				    sqlCmd.append(",update_date=sysdate");
				    sqlCmd.append(" WHERE program_id=? ");paramList.add(program_id) ;
					 	   	  
						   
		            //updateDBSqlList.add(sqlCmd); 		            	
	            		
					//if(DBManager.updateDB(updateDBSqlList)){
					if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)) {
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
	public String DeleteDB(HttpServletRequest request) throws Exception {
		String sqlCmd = "";
		String errMsg = "";

		try {
			//取出form裡的所有變數=================================== 
			Enumeration ep = request.getParameterNames();
			Enumeration ea = request.getAttributeNames();
			Hashtable t = new Hashtable();
			String name = "";

			for (; ep.hasMoreElements();) {
				name = (String) ep.nextElement();
				t.put(name, request.getParameter(name));
			}
			int row = Integer.parseInt((String) t.get("row"));
			System.out.println("row=" + row);
			List deleteData = new LinkedList();
			for (int i = 0; i < row; i++) {
				if (t.get("isModify_" + (i + 1)) != null) {
					deleteData.add((String) t.get("isModify_" + (i + 1)));
				}
			}
			System.out.println("deleteData.size=" + deleteData.size());

			//List updateDBSqlList = new LinkedList();
			List paramList = new ArrayList() ;
			List data = null;
			String tmpData = "";

			for (int i = 0; i < deleteData.size(); i++) {
				paramList.clear() ;
				sqlCmd = "SELECT * FROM WTT03_1 WHERE program_id=? ";
				paramList.add(deleteData.get(i)) ;
				data = DBManager.QueryDB_SQLParam(sqlCmd,paramList, "");
				paramList.clear() ;
				System.out.println("delete.size=" + data.size());
				if (data.size() != 0) {
					sqlCmd = " DELETE FROM WTT03_1 WHERE program_id=?";
					paramList.add( deleteData.get(i)) ;
					if(!this.updDbUsesPreparedStatement(sqlCmd,paramList)) {
						throw new Exception();
					}
				}
				//updateDBSqlList.add(sqlCmd);
			}
			errMsg = "相關資料寫入資料庫成功";
			/*if (DBManager.updateDB(updateDBSqlList)) {
				errMsg = errMsg + "相關資料寫入資料庫成功";
			} else {
				errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>"
						+ DBManager.getErrMsg();
			}*/
		} catch (Exception e) {
			System.out.println(e + ":" + e.getMessage());
			errMsg = errMsg + "相關資料寫入資料庫失敗";
					
		}

		return errMsg;
	}

	private boolean updDbUsesPreparedStatement(String sql, List paramList)
			throws Exception {
		List updateDBList = new ArrayList();//0:sql 1:data
		List updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
		List updateDBDataList = new ArrayList();//儲存參數的List
        for(int i =0;i<paramList.size();i++){
            System.out.println("i="+i+(String)paramList.get(i));
        }
		updateDBDataList.add(paramList);
		updateDBSqlList.add(sql);
		updateDBSqlList.add(updateDBDataList);
		updateDBList.add(updateDBSqlList);
		return DBManager.updateDB_ps(updateDBList);
	}%>    