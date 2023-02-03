<%
//93.12.26 create by 2295
//94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除 by 2295
//94.01.06 fix 增加機構類別bank_type,機構代號tbank_no為查詢條件 by 2295
//94.01.12 fix 維護的金融機構類別改成程式歸屬機構類別 by 2295
//			   Z -- 管理系統
//99.12.08 fix sqlInjection by 2808
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
	String alertMsg = "";	
	String webURL = "";	
	boolean doProcess = false;	
	
	//取得session資料,取得成功時,才繼續往下執行===================================================
	if(session.getAttribute("muser_id") == null){//session timeout	
      System.out.println("ZZ003W login timeout");   
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
	String nowact = ( request.getParameter("nowact")==null ) ? "" : (String)request.getParameter("nowact");			
	String bank_type = ( request.getParameter("BANK_TYPE")==null ) ? "" : (String)request.getParameter("BANK_TYPE");					
	String tbank_no = ( request.getParameter("TBANK_NO")==null ) ? "" : (String)request.getParameter("TBANK_NO");					

	System.out.println("ZZ002W.act="+act);				
	System.out.println("ZZ002W.nowact="+nowact);	
	
	if(session.getAttribute("muser_id") == null){	
      System.out.println("ZZ002W login timeout");   
	  rd = application.getRequestDispatcher( "/pages/reLogin.jsp?url=LoginError.jsp?timeout=true" );         	   
      rd.forward(request,response);
    }
   
	//登入者資訊
	String lguser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
	String lguser_name = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");		
	String lguser_type = ( session.getAttribute("muser_type")==null ) ? "" : (String)session.getAttribute("muser_type");				
	String lguser_tbank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");				
    
    
    	
	session.setAttribute("nowtbank_no",null);//94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除======
	
    if(!CheckPermission(request)){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );        
    }else{            
    	//set next jsp 	    	    	
    	if(act.equals("Qry") || act.equals("newQry") || act.equals("delQry") || act.equals("List") || act.equals("del") || act.equals("new") || act.equals("getData")){
    	    if(lguser_type.equals("A")){
    	       bank_type = ( session.getAttribute("bank_type")==null ) ? "" : (String)session.getAttribute("bank_type");				
    	       tbank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");				
    	    }
    	    if(bank_type.equals("")){
    	       bank_type = "1";
    	    }
    	    List tbank_no_list = getTbank_No(bank_type);
    	    if(tbank_no.equals("") && (tbank_no_list.size() != 0)){
               tbank_no= (String)((DataObject)tbank_no_list.get(0)).getValue("bank_no");
            }      	   
    	    request.setAttribute("tbank_no",tbank_no_list);
    	    if( act.equals("Qry") || act.equals("newQry") || act.equals("delQry") || act.equals("del") || act.equals("new")){    	    
    	       if(act.equals("newQry")) act="new";
    	       if(act.equals("delQry")) act="del";
    	       if(act.equals("new")){
    	          List bank_typeList = getBankType(lguser_id,lguser_type);    	        	   
    	          request.setAttribute("bank_typeList",bank_typeList);
    	       }   
    	       List WTT02List = getQryResult(lguser_type,tbank_no,act,bank_type);    	    
    	       request.setAttribute("WTT02List",WTT02List);    	       
    	    }
    	    rd = application.getRequestDispatcher( ListPgName +"?act="+act+"&bank_type="+bank_type+"&tbank_no="+tbank_no);            	        	        	    	    
    	    
    	    if(nowact.equals("new") || nowact.equals("List") || nowact.equals("Qry") || nowact.equals("del")){
        	   rd = application.getRequestDispatcher( ListPgName +"?act="+nowact+"&bank_type="+bank_type+"&tbank_no="+tbank_no);        
        	}
        	   
    	}else if(act.equals("Insert")){
    	    actMsg = InsertDB(request,lguser_id,lguser_name);    	    	        	    
        	rd = application.getRequestDispatcher( nextPgName );         
        }else if(act.equals("Delete")){
    	    actMsg = DeleteDB(request,lguser_id,lguser_name);    	    	        	    
        	rd = application.getRequestDispatcher( nextPgName );                 
        }
    	request.setAttribute("actMsg",actMsg);    
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


<%!
    private final static String nextPgName = "/pages/ActMsg.jsp";    
    private final static String EditPgName = "/pages/ZZ002W_Edit.jsp";    
    private final static String ListPgName = "/pages/ZZ002W_List.jsp";        
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    private boolean CheckPermission(HttpServletRequest request){//檢核權限    	    
    	    boolean CheckOK=false;
    	    HttpSession session = request.getSession();            
            Properties permission = ( session.getAttribute("ZZ002W")==null ) ? new Properties() : (Properties)session.getAttribute("ZZ002W");				                
            if(permission == null){
              System.out.println("ZZ002W.permission == null");
            }else{
               System.out.println("ZZ002W.permission.size ="+permission.size());
               
            }
            //只要有Query的權限,就可以進入畫面
        	if(permission != null && permission.get("Q") != null && permission.get("Q").equals("Y")){            
        	   CheckOK = true;//Query
        	}
        	return CheckOK;
    }           
    
    //取得總機構代碼
    private List getTbank_No(String bank_type){
    		//查詢條件    
    		String yy = Integer.parseInt(Utility.getYear())>99 ? "100" : "99" ;
    		List paramList = new ArrayList() ;
    		String sqlCmd = " select bank_no,bank_name from bn01 "
    					  + " where bank_type=? and m_year=? "
    		    		  + " order by bank_no";
    		paramList.add(bank_type) ;
    		paramList.add(yy) ;
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    } 
   
    //取得查詢結果
    private List getQryResult(String muser_type,String tbank_no,String act,String bank_type){
    		//查詢條件        		
    		String sqlCmd = "";
    		String yy = Integer.parseInt(Utility.getYear())>99 ? "100" : "99" ;
    		List paramList = new ArrayList() ;
    		sqlCmd = " select a.bank_type,a1.cmuse_name as bank_type_name,"
    			   + " a.tbank_no,bn01.bank_name as tbank_no_name,"
    			   + " a.bank_no,bn02.bank_name as bank_no_name,"
    			   + " a.subdep_id,c1.cmuse_name as subdep_id_name,"
    			   + " a.muser_id,a.muser_name,a.muser_type";
    		if(act.equals("Qry") || act.equals("del")){
    		   sqlCmd += " ,b.bank_type as pg_bank_type , b.bank_type || c.cmuse_name as pg_type_name";
    		}	   
    		
			sqlCmd = sqlCmd	    
			       + " from wtt01 a"
				   + " LEFT JOIN cdshareno a1 on a.bank_type = a1.cmuse_id and a1.cmuse_div='001'"
				   + " LEFT JOIN (select * from bn01 where m_year=? )bn01 on a.tbank_no = bn01.BANK_NO"
				   + " LEFT JOIN (select * from bn02 where m_year=? )bn02 on a.bank_no = bn02.bank_no  "
			       + " LEFT JOIN cdshareno c1 on a.subdep_id = c1.cmuse_id and c1.cmuse_div='010' ";
			paramList.add(yy) ;
			paramList.add(yy) ;
            if(act.equals("Qry") || act.equals("del")){			       
                   //94.01.12 fix 維護的金融機構類別改成程式歸屬機構類別
				   sqlCmd = sqlCmd 
				   	      + " ,wtt02 b LEFT JOIN cdshareno c on b.bank_type = c.cmuse_id and c.cmuse_div='016'";
			}	   
				   
			sqlCmd += " where";
			if(act.equals("Qry") || act.equals("del")){	
			   sqlCmd += " a.muser_id = b.muser_id ";
			   if(!bank_type.equals("ALL")){
			      sqlCmd += " and a.bank_type =?  and ";	
			      paramList.add(bank_type) ;
			   }   
			}
			if(act.equals("new")){
			   if(!bank_type.equals("ALL")){
			      sqlCmd += " a.bank_type =?  and ";	
			      paramList.add(bank_type) ;
			   }
			}
			if(muser_type.equals("S")){//suer user只能查管理者	   
			   sqlCmd += " a.muser_type='A'";
			   if(!tbank_no.equals("ALL")){
			      sqlCmd += " and a.tbank_no = ? ";
			      paramList.add(tbank_no) ;
			   }
			}	   
			if(muser_type.equals("A")){//管理者只能查一般使用者	   
			   sqlCmd += " a.muser_type=' '";
			   sqlCmd += " and a.tbank_no = ? ";
			   paramList.add(tbank_no) ;
			}
			
			sqlCmd = sqlCmd 
				  +	" and a.delete_mark <> 'Y'"			       				  
				  + " order by a.bank_type,a.tbank_no,a.bank_no,a.subdep_id,a.muser_id";    		   
				  
		    if(act.equals("Qry") || act.equals("del")){			       		  
		        sqlCmd += ",b.bank_type";
			}
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    }
    
    
    private List getBankType(String muser_id,String muser_type){
        //查詢條件        		
    	String sqlCmd = "";
        List paramList =new ArrayList() ;
    	if(muser_type.equals("S")){
    	   //94.01.12 fix 維護的金融機構類別改成程式歸機構類別
    	   sqlCmd = "select cmuse_id as bank_type,cmuse_name from cdshareno where cmuse_div='016' order by input_order";
    	}else if(muser_type.equals("A")){
    	   sqlCmd = " select bank_type,cmuse_name"
    	   		  + " from WTT02,cdshareno"
    	   		  + " where muser_id = ? "
    	   		  + " and bank_type = cmuse_id"
    	   		  + " and cmuse_div='016'" //94.01.12 fix 維護的金融機構類別改成程式歸機構類別
    	   		  + " order by bank_type";
    	   paramList.add(muser_id) ;
    	}
    	List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
        return dbData;
    
    }
    
    
    
	public String InsertDB(HttpServletRequest request,String lguser_id,String lguser_name) throws Exception{    	
		String sqlCmd = "";		
		String errMsg="";
		String user_id=lguser_id;
	    String user_name=lguser_name;		
		String bank_type=((String)request.getParameter("maintainBANK_TYPE")==null)?" ":(String)request.getParameter("maintainBANK_TYPE");
		System.out.println("bank_type="+bank_type);
		try {			    			    
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
					if ( t.get("isModify_" + (i+1)) != null ) {					  
					 insertData.add((String)t.get("isModify_"+(i+1)));
					}										
		  		}	
		  		System.out.println("insertData.size="+insertData.size());
			    
			    //List updateDBSqlList = new LinkedList();
			    List paramList =new ArrayList() ;
			    List data = null;		    
			   
			    
			    for(int i=0;i<insertData.size();i++){
			        /*
			        sqlCmd = "SELECT * FROM WTT02 WHERE muser_id='" + (String)insertData.get(i) + "'"
				 	       + " and bank_type='Z'";					 
			    	data = DBManager.QueryDB(sqlCmd,"");		 			    				
					if(data.size() == 0){
				   	    sqlCmd = "INSERT INTO WTT02 VALUES ('" +  (String)insertData.get(i) + "'"
					    	   + ",'Z'" 
					      	   + ",'" + user_id +"'"					       
					      	   + ",'" + user_name + "'"
					      	   + ",sysdate)"; 	
				   		updateDBSqlList.add(sqlCmd); 	      
			        }
			        */
			        paramList.clear();
				    sqlCmd = "SELECT * FROM WTT02 WHERE muser_id=? "
				    	   + " and bank_type=? ";
					paramList.add((String)insertData.get(i)) ; 
					paramList.add(bank_type);
			        data = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");		 	
			        paramList.clear() ;
				    //System.out.println("insert.size="+data.size());				    
					if(data.size() != 0){
					   sqlCmd = "UPDATE WTT02 SET"					  	      
					  	      + " user_id=?"
					          + ",user_name=?"
					          + ",update_date=sysdate" 		
						      + " where muser_id=?"
						      + " and bank_type=?";		
					   paramList.add(user_id) ;
					   paramList.add(user_name) ;
					   paramList.add(insertData.get(i)) ;
					   paramList.add(bank_type) ;
					}else{    			        
			           sqlCmd = "INSERT INTO WTT02 VALUES (?"
					          + ",?" 
					          + ",?"					       
					          + ",?"
					          + ",sysdate)"; 
			           paramList.add(insertData.get(i)) ;
			           paramList.add(bank_type) ;
			           paramList.add(user_id) ;
			           paramList.add(user_name);
    	   		    }
    	   		    //updateDBSqlList.add(sqlCmd); 	
					if(this.updDbUsesPreparedStatement(sqlCmd,paramList)){					 
						   errMsg = errMsg + "相關資料寫入資料庫成功";					
					}else{
					 	   errMsg = errMsg + "相關資料寫入資料庫失敗";
					 	   throw new Exception(errMsg) ;
					}	
	            }		
	            	
			    			    
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";							
		}	

		return errMsg;
	} 
	
	
	public String DeleteDB(HttpServletRequest request,String lguser_id,String lguser_name) throws Exception{    	
		String sqlCmd = "";		
		List paramList = new ArrayList() ;
		String errMsg="";
		String user_id=lguser_id;
	    String user_name=lguser_name;				
		try {			    			    
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
		  	    List deleteData = new LinkedList();
		  		for ( int i = 0; i < row; i++) {		  	    		  	  			  
					if ( t.get("isModify_" + (i+1)) != null ) {					  
					 deleteData.add((String)t.get("isModify_"+(i+1)));
					}										
		  		}	
		  		System.out.println("deleteData.size="+deleteData.size());
			    
			    List updateDBSqlList = new LinkedList();
			    List data = null;
			    String bank_type="";
			    String muser_id="";
			    for(int i=0;i<deleteData.size();i++){
			        if(((String)deleteData.get(i)).indexOf(":") != -1){			        
			           bank_type = ((String)deleteData.get(i)).substring(0,((String)deleteData.get(i)).indexOf(":"));
			           muser_id = ((String)deleteData.get(i)).substring(((String)deleteData.get(i)).indexOf(":")+1,((String)deleteData.get(i)).length());
			           System.out.println("bank_type="+bank_type);
			           System.out.println("muser_id="+muser_id);
			        }
			        paramList.clear() ;
				    sqlCmd = "SELECT * FROM WTT02 WHERE muser_id=?"
				    	   + " and bank_type=?";
					paramList.add(muser_id) ;
					paramList.add(bank_type) ;
			        data = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");		
			        paramList.clear() ;
				    System.out.println("insert.size="+data.size());				    
					if(data.size() != 0){
					   sqlCmd = "DELETE FROM WTT02 "					  	      
						      + " where muser_id=? and bank_type=?" ;
						paramList.add(muser_id) ;
					   paramList.add(bank_type) ;
					   if(this.updDbUsesPreparedStatement(sqlCmd,paramList)){					 
						   errMsg = errMsg + "相關資料寫入資料庫成功";					
					    }else{
					 	   errMsg = errMsg + "相關資料寫入資料庫失敗";
					 	   throw new Exception(errMsg) ;
					    }
					}    	   		     
   					//updateDBSqlList.add(sqlCmd); 		            	
	            }		
	            	
			    /*if(DBManager.updateDB(updateDBSqlList)){					 
				   errMsg = errMsg + "相關資料寫入資料庫成功";					
			    }else{
			 	   errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
			    }*/				    
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = "相關資料寫入資料庫失敗";						
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