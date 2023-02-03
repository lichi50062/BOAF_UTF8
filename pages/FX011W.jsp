<%
// 97.08.26 create 縣市政府變現性資產查核 by 2295
// 98.01.09 fix 查核人員.直接key in.不使用下拉式選單 by 2295
// 99.12.07 fix sqlInjection by 2808
//100.01.27 fix bug by 2295
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
      System.out.println("FX011W login timeout");   
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
	String muser_bank_type = ( session.getAttribute("bank_type")==null ) ? "" : (String)session.getAttribute("bank_type");			
    String muser_tbank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");
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

    if(!Utility.CheckPermission(request,"FX011W")){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );
    }else{
    	if(act.equals("new")){
			rd = application.getRequestDispatcher( EditPgName +"?act=new");
    	}else if(act.equals("Edit")){
    		List dbData = getBOAF_ASSETCHECK(m_year,tbank_no,examine);//年度.縣市政府.受檢單位
    	    request.setAttribute("BOAF_ASSETCHECK",dbData);
        	request.setAttribute("maintainInfo","select * from BOAF_ASSETCHECK WHERE m_year="+m_year+" and m2_name='"+tbank_no+"' and examine='"+examine+"'");
        	rd = application.getRequestDispatcher( EditPgName +"?act=Edit");
    	}else if(act.equals("List")){
    		List dbData = getBOAF_ASSETCHECK("",(muser_bank_type.equals("B")?muser_tbank_no:bank_no),"");
    		request.setAttribute("BOAF_ASSETCHECK",dbData);
        	rd = application.getRequestDispatcher( ListPgName +"?bank_no="+(muser_bank_type.equals("B")?muser_tbank_no:bank_no));    	
		}else if(act.equals("Insert")){
    	    actMsg = InsertDB(request,bank_no,lguser_id,lguser_name,muser_bank_type,muser_tbank_no);
        	rd = application.getRequestDispatcher( nextPgName+"?goPages=FX011W.jsp&act=List");
    	}else if(act.equals("Update")){
    	    actMsg = UpdateDB(request,bank_no,lguser_id,lguser_name,muser_bank_type,muser_tbank_no);
        	rd = application.getRequestDispatcher( nextPgName+"?goPages=FX011W.jsp&act=List");
    	}else if(act.equals("Delete")){
    	    actMsg = DeleteDB(request,bank_no,lguser_id,lguser_name,muser_bank_type,muser_tbank_no);
        	rd = application.getRequestDispatcher( nextPgName+"?goPages=FX011W.jsp&act=List");
		}else if(act.equals("Print")){    	    
        	rd = application.getRequestDispatcher( PrintPgName+"?m_year="+m_year);
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
    private final static String EditPgName = "/pages/FX011W_Edit.jsp";
    private final static String ListPgName = "/pages/FX011W_List.jsp";    
    private final static String PrintPgName = "/pages/FX011W_Excel.jsp";    
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    /***
     * 99.12.08 fix 縣市合併年度區分
     **/
    private List getBOAF_ASSETCHECK(String m_year,String m2_name,String examine){
    	List paramList =new ArrayList() ;
    	StringBuffer sqlCmd = new StringBuffer() ;
    	String s_year = "99" ;
        if(Integer.parseInt(Utility.getYear()) > 99) {
        	s_year = "100" ;
        }
		sqlCmd.append(" select BOAF_ASSETCHECK.*,ba01.bank_name from BOAF_ASSETCHECK "
					  + " left join (select * from ba01 where m_year=?) ba01 on BOAF_ASSETCHECK.examine = ba01.bank_no "
					  + " where 1=1 " );
		paramList.add(s_year) ;
		/*String condition = "";					  
		if(!m_year.equals("")) condition += (condition.length() > 0 ? " and":"")+" m_year="+m_year;					  
		if(!m2_name.equals("")) condition += (condition.length() > 0 ? " and":"")+" m2_name='"+m2_name+"'";					  
		if(!examine.equals("")) condition += (condition.length() > 0 ? " and":"")+" examine='"+examine+"'";				  
		sqlCmd += condition
		       + " order by boaf_assetcheck.m_year,check_date ";
		*/
		if(!"".equals(m_year)) {
			sqlCmd.append(" And boaf_assetcheck.m_year=? ");
			paramList.add(m_year) ;
		}
		if(!"".equals(m2_name)) {
			sqlCmd.append(" And m2_name =? ");
			paramList.add(m2_name) ;
		}
		if(!"".equals(examine)){
			sqlCmd.append(" And examine=?");
			paramList.add(examine) ;
		}
		sqlCmd.append(" order by boaf_assetcheck.m_year,check_date "); 
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"m_year,check_date,update_date");
        return dbData;
    }
   
    
    public String InsertDB(HttpServletRequest request,String bank_no,String lguser_id,String lguser_name,String muser_bank_type,String muser_tbank_no) throws Exception{	
	    String sqlCmd = "";
	    String errMsg="";
	    String m_year=((String)request.getParameter("m_year")==null)?"":(String)request.getParameter("m_year");
	    String examine=((String)request.getParameter("examine")==null)?"-":(String)request.getParameter("examine");
	    String check_date=(String)request.getParameter("CHECK_DATE");
	    String name=((String)request.getParameter("name")==null)?"-":(String)request.getParameter("name");//受檢人員			
	    String item=((String)request.getParameter("item")==null)?"-":(String)request.getParameter("item");	
	    String result=((String)request.getParameter("result")==null)?"-":(String)request.getParameter("result");
	    String content=((String)request.getParameter("content")==null)?"-":(String)request.getParameter("content");
	    String remark=((String)request.getParameter("remark")==null)?"-":(String)request.getParameter("remark");
	    String user_id=lguser_id;
	    String user_name=lguser_name;
		String check_name = "";//受檢人員
		try {
   			//List updateDBSqlList = new LinkedList();
   			List paramList =new ArrayList() ;
   			sqlCmd = "SELECT * FROM BOAF_ASSETCHECK WHERE m_year=?  AND m2_name=? and examine=? ";
   			paramList.add(m_year) ;
   			paramList.add(muser_bank_type.equals("B")?muser_tbank_no:bank_no) ;
   			paramList.add(examine) ;
			List data = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");			
			paramList.clear() ;
			if (data.size() > 0){
			    errMsg = errMsg + "此筆資料已存在無法新增<br>";
			}else{
			    /*
			    98.01.09 fix 查核人員.直接key in.不使用下拉式選單 by 2295
			    //取出form裡的所有變數=================================== 
		  		Enumeration ep = request.getParameterNames();
		  		Enumeration ea = request.getAttributeNames();
		  		Hashtable t = new Hashtable();
		  		String name = "";
		        
		  		for ( ; ep.hasMoreElements() ; ) {
			   		name = (String)ep.nextElement();
			   		t.put( name, request.getParameter(name) );			   
		  		}		  
		  		int row =Integer.parseInt((String)t.get("row"));
		  		System.out.println("row="+row);
		  		System.out.println("check_name="+check_name);
		  		
		  	    List insertData = new LinkedList();
		  	   
		  		for ( int i = 0; i < row; i++) {		  	    		  	  			  
					if ( t.get("name_" + (i+1)) != null ) {			
					  	check_name += (String)t.get("name_"+(i+1))+":"; 	  					 
					}										
		  		}	
		  		System.out.println("check_name="+check_name);
			    */
				/*sqlCmd = "INSERT INTO BOAF_ASSETCHECK VALUES("
		               +m_year+",'"
		               +(muser_bank_type.equals("B")?muser_tbank_no:bank_no)+"','"
		               +examine+"',"
		  			   +"to_date('"+check_date+"','YYYY/MM/DD'),'"
		               +name+"','"
		               +item+"','"
		               +result+"','"
		               +content+"','"
		               +remark+"','"
		               +user_id+"','"
		               +user_name
		               +"',sysdate)";*/
		        sqlCmd =" INSERT INTO BOAF_ASSETCHECK VALUES(" ;
		        sqlCmd += "?,?,?,to_date(?,'YYYY/MM/DD'),?,?,?,?,?,?,?,sysdate)" ;
		   		paramList.add(m_year) ;
		   		paramList.add(muser_bank_type.equals("B")?muser_tbank_no:bank_no) ;
		   		paramList.add(examine) ;
		   		paramList.add(check_date) ;
		   		paramList.add(name) ;
		   		paramList.add(item) ;
		   		paramList.add(result) ;
		   		paramList.add(content) ;
		   		paramList.add(remark) ;
		   		paramList.add(user_id) ;
		   		paramList.add(user_name) ;
                //updateDBSqlList.add(sqlCmd);
			    //if(DBManager.updateDB(updateDBSqlList)){
			    if(this.updDbUsesPreparedStatement(sqlCmd,paramList)){
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
	
	public String UpdateDB(HttpServletRequest request,String bank_no,String lguser_id,String lguser_name,String muser_bank_type,String muser_tbank_no) throws Exception{
		String sqlCmd = "";		
		String errMsg="";
		
		String m_year = ( request.getParameter("m_year")==null ) ? "" : (String)request.getParameter("m_year");		
	    String examine = ( request.getParameter("examine")==null ) ? "" : (String)request.getParameter("examine");
		
	    String check_date=(String)request.getParameter("CHECK_DATE");
	    String name=((String)request.getParameter("name")==null)?"-":(String)request.getParameter("name");//受檢人員			
	    String item=((String)request.getParameter("item")==null)?"-":(String)request.getParameter("item");	
	    String result=((String)request.getParameter("result")==null)?"-":(String)request.getParameter("result");
	    String content=((String)request.getParameter("content")==null)?"-":(String)request.getParameter("content");
	    String remark=((String)request.getParameter("remark")==null)?"-":(String)request.getParameter("remark");
	    String user_id=lguser_id;
	    String user_name=lguser_name;
		String check_name = "";//受檢人員

		try {
				//List updateDBSqlList = new LinkedList();
				List paramList = new ArrayList() ;
				sqlCmd = "SELECT * FROM BOAF_ASSETCHECK WHERE m_year=?  AND m2_name=? and examine=? ";
				paramList.add(m_year) ;
				paramList.add(muser_bank_type.equals("B")?muser_tbank_no:bank_no) ;
				paramList.add(examine) ;
			    List data = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
			    paramList.clear() ;
				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
				}else{
				    /*寫入log檔中*/
				    sqlCmd = "INSERT INTO BOAF_ASSETCHECK_LOG "
		             	   +"select m_year,m2_name,examine,check_date,name,item,result,"
		             	   +"content,remark,user_id,user_name,update_date,?,?,sysdate,'U' from BOAF_ASSETCHECK WHERE m_year=? AND m2_name=? and examine=? ";
		        	paramList.add(user_id) ;
		        	paramList.add(user_name) ;
		        	paramList.add(m_year) ;
		        	paramList.add(muser_bank_type.equals("B")?muser_tbank_no:bank_no);
		        	paramList.add(examine) ;
		        	this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
		        	paramList.clear() ;
		            //updateDBSqlList.add(sqlCmd);  
		          	/*
		          	98.01.09 fix 查核人員.直接key in.不使用下拉式選單 by 2295
		          	//取出form裡的所有變數=================================== 
		  		    Enumeration ep = request.getParameterNames();
		  		    Enumeration ea = request.getAttributeNames();
		  		    Hashtable t = new Hashtable();
		  		    String name = "";
		            
		  		    for ( ; ep.hasMoreElements() ; ) {
			   	    	name = (String)ep.nextElement();
			   	    	t.put( name, request.getParameter(name) );			   
		  		    }		  
		  		    int row =Integer.parseInt((String)t.get("row"));
		  		    System.out.println("row="+row);
		  	        List insertData = new LinkedList();
		  		    for ( int i = 0; i < row; i++) {		  	    		  	  			  
				    	if ( t.get("name_" + (i+1)) != null ) {			
				    	  	check_name += (String)t.get("name_"+(i+1))+":"; 	  					 
				    	}										
		  		    }	
		  		    System.out.println("check_name="+check_name);
		          	*/
				    sqlCmd = "UPDATE BOAF_ASSETCHECK SET "
				       	   + "check_date = to_date(?,'YYYY/MM/DD')"
		               	   +",name=?"
		               	   +",item=?"
		               	   +",result=?"
		               	   +",content=?"
		                   +",remark=?"
				       	   +",user_id=?"
   					   	   +",user_name=?,update_date=sysdate"
   			           	   +" where m_year=? AND m2_name=? and examine=?";
					
		            //updateDBSqlList.add(sqlCmd);
					paramList.add(check_date) ;
					paramList.add(name);
					paramList.add(item) ;
					paramList.add(result) ;
					paramList.add(content) ;
					paramList.add(remark) ;
					paramList.add(user_id) ;
					paramList.add(user_name) ;
					paramList.add(m_year) ;
					paramList.add(muser_bank_type.equals("B")?muser_tbank_no:bank_no);
					paramList.add(examine) ;
					//if(DBManager.updateDB(updateDBSqlList)){
					if(this.updDbUsesPreparedStatement(sqlCmd,paramList)){
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
    
    public String DeleteDB(HttpServletRequest request,String bank_no,String lguser_id,String lguser_name,String muser_bank_type,String muser_tbank_no) throws Exception{
		String sqlCmd = "";		
		String errMsg="";
		String m_year = ( request.getParameter("m_year")==null ) ? "" : (String)request.getParameter("m_year");		
	    String examine = ( request.getParameter("examine")==null ) ? "" : (String)request.getParameter("examine");
		String user_id=lguser_id;
	  	String user_name=lguser_name;
		

		try {
			 	//List updateDBSqlList = new LinkedList();
			 	List paramList = new ArrayList() ;
				sqlCmd = "SELECT * FROM BOAF_ASSETCHECK WHERE m_year=?  AND m2_name=? and examine=? ";
				paramList.add(m_year) ;
				paramList.add(muser_bank_type.equals("B")?muser_tbank_no:bank_no) ;
				paramList.add(examine) ;
			    List data = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
			
				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法刪除<br>";
				}else{
				    /*寫入log檔中*/
				    sqlCmd = "INSERT INTO BOAF_ASSETCHECK_LOG "
		             	   +"select m_year,m2_name,examine,check_date,name,item,result,"
		             	   +"content,remark,user_id,user_name,update_date,?,?,sysdate,'D' from BOAF_ASSETCHECK WHERE m_year=? AND m2_name=? and examine=?";
                    paramList.clear() ;		             	   
		        	paramList.add(user_id) ;
		        	paramList.add(user_name) ;
		        	paramList.add(m_year) ;
		        	paramList.add(muser_bank_type.equals("B")?muser_tbank_no:bank_no) ;
		        	paramList.add(examine) ;
		        	//updateDBSqlList.add(sqlCmd);  
 					
 					if(!this.updDbUsesPreparedStatement(sqlCmd,paramList)){						
				   		errMsg = errMsg + "BOAF_ASSETCHECK_LOG寫入資料失敗";
					}
 					paramList.clear() ;
				    sqlCmd = " delete BOAF_ASSETCHECK where m_year=? AND m2_name=? and examine=?";
				    paramList.add(m_year) ;
				    paramList.add(bank_no);
				    paramList.add(examine) ;
		            //updateDBSqlList.add(sqlCmd);

					//if(DBManager.updateDB(updateDBSqlList)){
					if(this.updDbUsesPreparedStatement(sqlCmd,paramList)){
						errMsg = errMsg + "相關資料刪除成功";
					}else{
				   		errMsg = errMsg + "相關資料刪除失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
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
