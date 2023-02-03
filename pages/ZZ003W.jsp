<%
//93.12.28 create by 2295
//94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除 by 2295
//94.02.15 fix 區分網際網路申報/MIS by 2295
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
	String menu = ( request.getParameter("menu")==null ) ? "" : (String)request.getParameter("menu");						
	String web_type = ( request.getParameter("WEB_TYPE")==null ) ? "" : (String)request.getParameter("WEB_TYPE");				
	String sys_type = ( request.getParameter("SYS_TYPE")==null ) ? "" : (String)request.getParameter("SYS_TYPE");				
	String program_type = ( request.getParameter("PROGRAM_TYPE")==null ) ? "" : (String)request.getParameter("PROGRAM_TYPE");				
	
	System.out.println("act="+act);	
   
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
    	if(!menu.equals("")){
    	   session.setAttribute("menu",menu);    	   
    	}
    	menu = (session.getAttribute("menu") == null)?"":(String)session.getAttribute("menu");   	
    	if(act.equals("new")){    	   
    	   rd = application.getRequestDispatcher( EditPgName +"?act=new&menu="+menu);                	
        }else if(act.equals("List")){
           rd = application.getRequestDispatcher( ListPgName +"?act=List&menu="+menu);                	
        }else if(act.equals("Qry")){            
    	    List WTT03_2List = getQryResult("","","");      
    	    request.setAttribute("WTT03_2List",WTT03_2List);    	     	    
    	    rd = application.getRequestDispatcher( ListPgName +"?act=Qry&menu="+menu);            	        	        	    	    
    	}else if(act.equals("DelQry")){    	    	    
    	    List WTT03_2List = getQryResult(web_type,sys_type,program_type);      	    
    	    request.setAttribute("WTT03_2List",WTT03_2List);
    	    rd = application.getRequestDispatcher( ListPgName +"?act=del&web_type="+web_type+"&sys_type="+sys_type+"&program_type="+program_type+"&menu="+menu); 
    	}else if(act.equals("del")){    	    	    
    	    List WTT03_2List = getQryResult(web_type,sys_type,program_type);    	    
    	    request.setAttribute("WTT03_2List",WTT03_2List);
    	    rd = application.getRequestDispatcher( ListPgName +"?act=del&menu="+menu);            	        	        	    	        	    
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
    private final static String EditPgName = "/pages/ZZ003W_Edit.jsp";    
    private final static String ListPgName = "/pages/ZZ003W_List.jsp";        
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    private boolean CheckPermission(HttpServletRequest request){//檢核權限    	    
    	    boolean CheckOK=false;
    	    HttpSession session = request.getSession();            
            Properties permission = ( session.getAttribute("ZZ003W")==null ) ? new Properties() : (Properties)session.getAttribute("ZZ003W");				                
            if(permission == null){
              System.out.println("ZZ003W.permission == null");
            }else{
               System.out.println("ZZ003W.permission.size ="+permission.size());
               
            }
            //只要有Query的權限,就可以進入畫面
        	if(permission != null && permission.get("Q") != null && permission.get("Q").equals("Y")){            
        	   CheckOK = true;//Query
        	}
        	return CheckOK;
    }           
    
   
    //取得查詢結果
    private List getQryResult(String web_type,String sys_type,String program_type){
    		//查詢條件        		
    		String sqlCmd = "";
    		List paramList =new ArrayList() ;
    		sqlCmd = " select a.program_id,d.program_name"
    			   + " ,a.sys_type,b.cmuse_name as sys_type_name"
    			   + " ,a.program_type,c.cmuse_name as program_type_name"
    			   + " ,a.web_type,e.cmuse_name as web_type_name"
  			       + " from wtt03_2 a,wtt03_1 d,cdshareno b,cdshareno c,cdshareno e"
  				   + " where a.program_id = d.program_id";
  			if(!web_type.equals("")){
  			   sqlCmd += " and web_type=? ";
  			   paramList.add(web_type) ;
  			}	   
  			if(!sys_type.equals("")){
  			   sqlCmd += " and sys_type=? ";
  			   paramList.add(sys_type) ;
  			}
  			if(!program_type.equals("")){
  			   sqlCmd += " and program_type=? ";
  			   paramList.add(program_type) ;
  			}
  			sqlCmd += " and (a.sys_type = b.cmuse_id and b.cmuse_div='018')"
  				   + " and (a.program_type = c.cmuse_id and c.cmuse_div='016')"
  				   + " and (a.web_type = e.cmuse_id and e.cmuse_div='017')"
  				   + " order by a.program_id,a.sys_type,a.program_type,a.web_type ";
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    }
    
	public String InsertDB(HttpServletRequest request,String lguser_id,String lguser_name) throws Exception{    	
		String sqlCmd = "";		
		String errMsg="";
		String user_id=lguser_id;
	    String user_name=lguser_name;		
	    String sys_type="";
	    String program_type="";
	    String program_id=((String)request.getParameter("PROGRAM_ID")==null)?" ":(String)request.getParameter("PROGRAM_ID");
	    String web_type=((String)request.getParameter("WEB_TYPE")==null)?" ":(String)request.getParameter("WEB_TYPE");
		
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
		  		int row_systype =Integer.parseInt((String)t.get("row_systype"));
		  		System.out.println("row_systype="+row_systype);
		  	    List systypeData = new LinkedList();
		  		for ( int i = 0; i < row_systype; i++) {		  	    		  	  			  
					if ( t.get("Systype_isModify_" + (i+1)) != null ) {					  
					 systypeData.add((String)t.get("Systype_isModify_"+(i+1)));
					}										
		  		}	
		  		System.out.println("systypeData.size="+systypeData.size());
		  		
		  		int row_programtype =Integer.parseInt((String)t.get("row_programtype"));
		  		System.out.println("row_programtype="+row_programtype);
		  	    List pgtypeData = new LinkedList();
		  		for ( int i = 0; i < row_programtype; i++) {		  	    		  	  			  
					if ( t.get("Programtype_isModify_" + (i+1)) != null ) {					  
					 pgtypeData.add((String)t.get("Programtype_isModify_"+(i+1)));
					}										
		  		}	
		  		System.out.println("pgtypeData.size="+pgtypeData.size());
			    
			    //List updateDBSqlList = new LinkedList();
			    List paramList = new ArrayList() ;
			    List data = null;
			    for(int i=0;i<systypeData.size();i++){
			      for(int j=0;j<pgtypeData.size();j++){
			    	  paramList.clear() ;
				    sqlCmd = "SELECT * FROM WTT03_2 "
				    	   + " WHERE sys_type=?"
				    	   + " and program_type=?"
					 	   + " and program_id=?"
					 	   + " and web_type=?";
				      paramList.add(systypeData.get(i)) ;
				      paramList.add(pgtypeData.get(j)) ;
				      paramList.add(program_id) ;
				      paramList.add(web_type) ;
			        data = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");		
			        paramList.clear() ;
				    System.out.println("insert.size="+data.size());				    
					if(data.size() != 0){
					   sqlCmd = "UPDATE WTT03_2 SET"					  	      
					  	      + " user_id=?"
					          + ",user_name=?"
					          + ",update_date=sysdate" 		
						      + " WHERE sys_type=?"
				    	   	  + " and program_type=?"
					 	   	  + " and program_id=?"
					 	   	  + " and web_type=?";
					   paramList.add(user_id) ;
					   paramList.add(user_name) ;
					   paramList.add(systypeData.get(i)) ;
					   paramList.add(pgtypeData.get(j)) ;
					   paramList.add(program_id) ;
					   paramList.add(web_type) ;
					}else{    			        
			           sqlCmd = "INSERT INTO WTT03_2 VALUES (?,?,?,?,?,?,sysdate)";
			           	     
			           paramList.add(systypeData.get(i)) ;
			           paramList.add(pgtypeData.get(j)) ;
			           paramList.add(program_id) ;
			           paramList.add(web_type) ;
			           paramList.add(user_id) ;
			           paramList.add(user_name) ;
    	   		    }
					if(this.updDbUsesPreparedStatement(sqlCmd,paramList)){					 
						   errMsg = "相關資料寫入資料庫成功";					
					}else{
					 	  throw new Exception() ;
					}	 
   					//updateDBSqlList.add(sqlCmd); 		            	
   				  }//end of pgtype	
	            }//end of systype		
	            	
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
	
	
	public String DeleteDB(HttpServletRequest request,String lguser_id,String lguser_name) throws Exception{    	
		String sqlCmd = "";		
		String errMsg="";
		String sys_type="";
	    String program_type="";
	    String program_id="";
		String user_id=lguser_id;
	    String user_name=lguser_name;		
	    StringTokenizer st = null;
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
			    
			    //List updateDBSqlList = new LinkedList();
			    List paramList = new ArrayList() ;
			    List data = null;
			    String tmpData="";
			    
			    for(int i=0;i<deleteData.size();i++){	
			        st = new StringTokenizer((String)deleteData.get(i),":");
     				while (st.hasMoreTokens()) {
         				sys_type = st.nextToken();
         				program_type = st.nextToken();
         				program_id = st.nextToken();
     				}		        
			        paramList.clear() ;   
				    sqlCmd = "SELECT * FROM WTT03_2 WHERE sys_type=?"
				    	   + " and program_type=?"
				    	   + " and program_id=?" ;
					paramList.add(sys_type) ;
					paramList.add(program_type) ;
					paramList.add(program_id) ;
			        data = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");		 
			        paramList.clear() ;
				    System.out.println("delete.size="+data.size());				    
					if(data.size() != 0){
					   sqlCmd = " DELETE FROM WTT03_2 "					  	      
						      + " WHERE sys_type=?"
				    	   	  + " and program_type=?"
				    	   	  + " and program_id=?" ;	    
					   paramList.add(sys_type) ;
					   paramList.add(program_type) ;
					   paramList.add(program_id) ;
					   if(this.updDbUsesPreparedStatement(sqlCmd,paramList)){					 
						   errMsg = "相關資料寫入資料庫成功";					
					    }else{
					 	  throw new Exception() ;
					    }
					}    	   		     
   					//updateDBSqlList.add(sqlCmd); 		            	
	            }		
	            	
			    /* if(DBManager.updateDB(updateDBSqlList)){					 
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