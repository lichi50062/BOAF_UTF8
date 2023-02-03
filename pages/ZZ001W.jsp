<%
// 93.12.20 add 權限檢核 by 2295
//          add 設定異動者資訊 by 2295
// 93.12.23 add 超過登入時間,請重新登入 by 2295
// 93.12.27 fix 若lguser_type="A",以登入者的bank_type及tbank_no為主 by 2295
// 94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除 by 2295
// 94.01.17 fix 重設密碼時,把firstlogin_mark='Y' by 2295
// 94.02.18 fix 若為superuser才清空bank_type.tbank_no by 2295
// 94.03.22 fix 按照總機構代號.組室.科別.帳號做sort by 2295
// 94.03.23 fix 直接將帳號刪除,一併將其所屬權限刪除 by 2295
// 94.04.07 fix WTT04有data時,才delete by 2295	                   			    	   						   		            
// 94.05.18 fix 管理者(A)只能看一般使用者,系統管理者(S)只能看管理者(A) by 2295
// 94.06.14 add 農金局的.只能看其建置的使用者 by 2295
// 96.03.22 add 重設密碼時,將密碼加密 by 2295
// 99.12.08 fix sqlInjection by 2808
//100.08.18 fix 查詢結果資料重覆 by 2295
//102.08.15 add 操作歷程寫入log by2968
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
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
      System.out.println("ZZ001W login timeout");   
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
	String muser_id = ( request.getParameter("muser_id")==null ) ? "" : (String)request.getParameter("muser_id");							
	System.out.println("act="+act);			
	System.out.println("nowact="+nowact);			
	session.setAttribute("nowtbank_no",null);//94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除======
	if(session.getAttribute("muser_id") == null){	
      System.out.println("ZZ001W login timeout");   
	  rd = application.getRequestDispatcher( "/pages/reLogin.jsp?url=LoginError.jsp?timeout=true" );         	   
      rd.forward(request,response);
    }
   
	//登入者資訊
	String lguser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
	String lguser_name = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");		
	String lguser_type = ( session.getAttribute("muser_type")==null ) ? "" : (String)session.getAttribute("muser_type");				
	
	if(lguser_type.equals("A")){
	   System.out.println("111bank_type="+session.getAttribute("bank_type"));
	   bank_type = (session.getAttribute("bank_type") == null)?"":(String)session.getAttribute("bank_type");		   
	   tbank_no = (session.getAttribute("tbank_no") == null)?"":(String)session.getAttribute("tbank_no");		   
	}
		
    if(!CheckPermission(request)){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );        
    }else{            
    	//set next jsp 	
    	if(act.equals("new") || act.equals("getData") || act.equals("List")){    	   
    	    //94.02.18 fix 若為superuser才清空bank_type.tbank_no
    	    if(lguser_type.equals("S") && act.equals("new")){
    	       bank_type="";
    	       tbank_no="";    	       
    	    }
    	    if(bank_type.equals("")){
    	       bank_type = "1";
    	    } 
    	    List tbank_no_list = getTbank_No(bank_type);
    	    System.out.println("test_tbank_no="+tbank_no);
    	    if(tbank_no.equals("") && (tbank_no_list.size() != 0)){
               tbank_no= (String)((DataObject)tbank_no_list.get(0)).getValue("bank_no");
            }  
    	    List bank_no_list = getBank_No(tbank_no);
    	    request.setAttribute("tbank_no",tbank_no_list);
    	    request.setAttribute("bank_no",bank_no_list); 
    	    if(act.equals("new")){
        	   rd = application.getRequestDispatcher( EditPgName +"?act=new");                	
        	}
        	if(act.equals("List")){
        	   rd = application.getRequestDispatcher( ListPgName +"?act=List");                	
        	}    
        	if(act.equals("getData")){
        	    if(nowact.equals("new") || nowact.equals("Edit")){
        	       rd = application.getRequestDispatcher( EditPgName +"?act="+nowact+"&bank_type="+bank_type+"&tbank_no="+tbank_no);        
        	    }
        	    if(nowact.equals("List") || nowact.equals("Qry")){
        	       rd = application.getRequestDispatcher( ListPgName +"?act="+nowact+"&bank_type="+bank_type+"&tbank_no="+tbank_no);        
        	    }   
        	}
        	       
    	}else if(act.equals("Qry")){
    	    List tbank_no_list = getTbank_No(bank_type);
    	    if(tbank_no.equals("") && (tbank_no_list.size() != 0)){
               tbank_no= (String)((DataObject)tbank_no_list.get(0)).getValue("bank_no");
            }  
    	    //List bank_no_list = getBank_No(tbank_no);
    	    request.setAttribute("tbank_no",tbank_no_list);
    	    //request.setAttribute("bank_no",bank_no_list); 
    	
    	    List WTT01List = getQryResult(bank_type,tbank_no,lguser_id,lguser_type);
    	    //List WTT01List = getQryResult("2","");
    	    request.setAttribute("WTT01List",WTT01List);
    	    rd = application.getRequestDispatcher( ListPgName +"?act=Qry&bank_type="+bank_type+"&tbank_no="+tbank_no);            	        	        	    	    
    	}else if(act.equals("Edit")){        		
            List WTT01 = getWTT01(muser_id);    	      
            request.setAttribute("WTT01",WTT01);    	       
            if(WTT01.size() != 0){    	       
               bank_type = (String)((DataObject)WTT01.get(0)).getValue("bank_type");
               tbank_no = (String)((DataObject)WTT01.get(0)).getValue("tbank_no");
            }
            List tbank_no_list = getTbank_No(bank_type);    	      
    	    List bank_no_list = getBank_No(tbank_no);
    	    request.setAttribute("tbank_no",tbank_no_list);
    	    request.setAttribute("bank_no",bank_no_list);
    	  	//操作歷程寫入log
    	    this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,muser_id,"Q");
    	 	rd = application.getRequestDispatcher( EditPgName +"?act=Edit&bank_type="+bank_type+"&tbank_no="+tbank_no);        
    	}else if(act.equals("Insert")){
    	    actMsg = InsertDB(request,lguser_id,lguser_name,lguser_type,bank_type,tbank_no);
    	    if("Y".equals(actMsg)){
    	        if(!"2".equals(bank_type)){//局外
 			       muser_id = tbank_no + (((String)request.getParameter("MUSER_ID")==null)?"":(String)request.getParameter("MUSER_ID"));
 			    }
    	      	//操作歷程寫入log
        	    actMsg = this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,muser_id,"I");
    	      	if("Y".equals(actMsg)){
    	      	    actMsg = "相關資料寫入資料庫成功";
    	      	}
    	    }   
        	rd = application.getRequestDispatcher( nextPgName );         
        }else if(act.equals("Update")){
    	    actMsg = UpdateDB(request,lguser_id,lguser_name);  
    	    if("Y".equals(actMsg)){
    	      	//操作歷程寫入log
        	    actMsg = this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,muser_id,"U");
        	    if("Y".equals(actMsg)){
    	      	    actMsg = "相關資料寫入資料庫成功";
    	      	}
    	    }   
        	rd = application.getRequestDispatcher( nextPgName );         
        }else if(act.equals("Delete")){
    	    actMsg = DeleteDB(request,lguser_id,lguser_name); 
    	    if("Y".equals(actMsg)){
    	      	//操作歷程寫入log
        	    actMsg = this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,muser_id,"D");
        	    if("Y".equals(actMsg)){
    	      	    actMsg = "相關資料刪除成功";
    	      	}
    	    }  
        	rd = application.getRequestDispatcher( nextPgName );         
        }else if(act.equals("ResetPwd")){
    	    actMsg = ResetPwd(request,lguser_id,lguser_name);  
    	    if("Y".equals(actMsg)){
    	      	//操作歷程寫入log
        	    actMsg = this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,muser_id,"R");
        	    if("Y".equals(actMsg)){
    	      	    actMsg = "相關資料寫入資料庫成功";
    	      	}
    	    }
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


<%! private final static String program_id = "ZZ001W";
    private final static String nextPgName = "/pages/ActMsg.jsp";    
    private final static String EditPgName = "/pages/"+program_id+"_Edit.jsp";    
    private final static String ListPgName = "/pages/"+program_id+"_List.jsp";        
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    private boolean CheckPermission(HttpServletRequest request){//檢核權限    	    
    	    boolean CheckOK=false;
    	    HttpSession session = request.getSession();            
            Properties permission = ( session.getAttribute("ZZ001W")==null ) ? new Properties() : (Properties)session.getAttribute("ZZ001W");				                
            if(permission == null){
              System.out.println("ZZ001W.permission == null");
            }else{
               System.out.println("ZZ001W.permission.size ="+permission.size());
               
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
    					  + " where bank_type=? and m_year=?"
    		    		  + " order by bank_no";
    		paramList.add(bank_type) ;
    		paramList.add(yy) ;
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    } 
    //取得各組室代碼
    private List getBank_No(String tbank_no){
    		//查詢條件    
    		String yy = Integer.parseInt(Utility.getYear())>99 ? "100" : "99" ;
    		List paramList = new ArrayList() ;
    		String sqlCmd = "select bank_no,bank_name from bn02 "
						  + " where tbank_no=? and m_year=? "
    		 			  + " order by bank_no";
    		paramList.add(tbank_no) ;
    		paramList.add(yy) ;
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    } 
    
    //取得WTT01該使用者帳號資料
    private List getWTT01(String muser_id){
    		//查詢條件    
    		List paramList = new ArrayList() ;
    		String sqlCmd = " select * from WTT01 "
    					  + " where muser_id=? ";  
    		paramList.add(muser_id) ;
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    } 
    //取得該管理者機構類別
    private List getBank_Type(String muser_id,String muser_type){
    		//查詢條件        		
    		String sqlCmd = "";
    		List paramList = new ArrayList() ;
    		if(muser_type.equals("S")){
     		   sqlCmd = "select * from cdshareno where cmuse_div='001' and cmuse_id <>'Z'";
     		}else{
     		   sqlCmd = "select * from cdshareno where cmuse_div='001' and cmuse_id <>'Z'" 
     		   		  + " and cmuse_id in (select bank_type from wtt01 where muser_id=? )";    
     		   paramList.add(muser_id) ;
     		}	
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    }
    
    //取得查詢結果
    private List getQryResult(String bank_type,String bank_no,String lguser_id,String lguser_type){
    		//查詢條件        		
    		String yy = Integer.parseInt(Utility.getYear())>99 ? "100" : "99" ;
    		List paramList = new ArrayList() ;
    		String sqlCmd = "";
    		sqlCmd = "select b.bank_no as tbank_no,b.bank_name as tbank_name ,a.bank_no,bn02.bank_name,a.subdep_id,c.cmuse_name, "
			       + "a.muser_id,a.muser_name,a.add_name,a.muser_type,a.firstlogin_mark, "
				   + "a.lock_mark,a.delete_mark,b.bank_type,b.bank_no "
				   + "from wtt01 a LEFT JOIN (select * from bn01 where m_year=?) b on a.tbank_no = b.bank_no "
				   + "LEFT JOIN cdshareno c on a.subdep_id = c.cmuse_id and c.cmuse_div='010' "
				   + "LEFT JOIN (select * from bn02 where m_year=?)bn02  on a.bank_no = bn02.bank_no ";
			paramList.add(yy);	   
			paramList.add(yy);	   
			if(!bank_type.equals("")){	   
			    sqlCmd = sqlCmd + " where b.bank_type=?  ";
			    paramList.add(bank_type) ;
			    if(bank_type.equals("2")){//add 94.06.14 農金局的.只能看其建置的使用者
			       sqlCmd = sqlCmd + " and a.add_user=? ";
			       paramList.add(lguser_id) ;
			    }
			}
			if(!bank_no.equals("")){//若有總機構代碼.以總機構代碼為主
			    if(!bank_no.equals("ALL")){
			        sqlCmd = sqlCmd + "and b.bank_no=? ";
			        paramList.add(bank_no) ;
			    }    
			}
			//94.05.18 fix 管理者只能看一般使用者,系統管理者只能看管理者
			if(lguser_type.equals("A")){
			   sqlCmd += " and a.muser_id <> ? "
			   		  +  " and (a.muser_type <> 'A' and a.muser_type <> 'S' or a.muser_type is null)";		
			   paramList.add(lguser_id) ;
			}else if(lguser_type.equals("S")){
			   sqlCmd += " and (a.muser_type = 'A' and a.muser_type is not null)";			   		  	   		  
			}
				   
			//sqlCmd = sqlCmd + "order by b.bank_type,b.bank_no,a.muser_id";
			//94.03.22 fix 按照總機構代號.組室.科別.帳號做sort 
			sqlCmd = sqlCmd + " order by tbank_no,a.bank_no,a.subdep_id,a.muser_id";
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    }
    
	public String InsertDB(HttpServletRequest request,String lguser_id,String lguser_name,String lguser_type,String lgbank_type,String lgtbank_no) throws Exception{    	
		String sqlCmd = "";		
		String errMsg="";		
		String muser_id=((String)request.getParameter("MUSER_ID")==null)?"":(String)request.getParameter("MUSER_ID");
		String muser_name=((String)request.getParameter("MUSER_NAME")==null)?"":(String)request.getParameter("MUSER_NAME");
		String muser_password="";
		String muser_i_o="";		
		String bank_type=((String)request.getParameter("BANK_TYPE")==null)?"":(String)request.getParameter("BANK_TYPE");
		String tbank_no=((String)request.getParameter("TBANK_NO")==null)?"":(String)request.getParameter("TBANK_NO");
		String bank_no=((String)request.getParameter("BANK_NO")==null)?" ":(String)request.getParameter("BANK_NO");
		bank_no = bank_no.equals("")?" ":bank_no;
		//fix 93.12.27 by 2295
		if(lguser_type.equals("A")){	 
	   	   bank_type = lgbank_type;
	       tbank_no = lgtbank_no;
	    }
		String subdep_id=((String)request.getParameter("SUBDEP_ID")==null)?"":(String)request.getParameter("SUBDEP_ID");
		String add_user=lguser_id;
		String add_name=lguser_name;
		String firstlogin_mark="Y";
		String login_mark="N";
		String lock_mark="N";
		String delete_mark="N";
		String muser_type=" ";
		String password_pre="";
	    String user_id=lguser_id;
	    String user_name=lguser_name;
	    
	    
		
		try {
			    List paramList =new ArrayList() ;
			    if(bank_type.equals("2")){//局內
			       muser_password=muser_id;
			       password_pre=muser_id;
			       muser_i_o="I";
			    }else{//局外
			       muser_id = tbank_no + muser_id;
			       muser_password =muser_id;
			       password_pre=muser_id;			       
			       muser_i_o="O";
			    }
			    
			    if(lguser_type.equals("S")){//super user用來新增管理者
			       muser_type = "A";
			    }
			    
				sqlCmd = "SELECT * FROM WTT01 WHERE muser_id=? ";					 
				paramList.add(muser_id) ;	 
			    List data = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");		 			    
				System.out.println("WTT01.size="+data.size());
				paramList.clear();
				if (data.size() != 0){
				    errMsg = errMsg + "此筆使用者帳號已存在無法新增<br>";
				}else{    			        
   				 sqlCmd = " INSERT INTO WTT01 VALUES ( " ;
   				 sqlCmd +="?,?,?,?,?,?,?,?,?,?,sysdate,?,?,?,?,?,sysdate,?,?,?,sysdate)";
   				 paramList.add(muser_id) ;
   				 paramList.add( muser_name );
		         paramList.add( Utility.encode(muser_password) );					       					       
		       	 paramList.add( muser_i_o );
		         paramList.add( bank_type ); 
		         paramList.add( tbank_no ); 
		         paramList.add( bank_no ); 
		         paramList.add( subdep_id ); 
		         paramList.add( add_user ); 
		         paramList.add( add_name );
		         paramList.add( firstlogin_mark );
		         paramList.add( login_mark );
		         paramList.add( lock_mark );					       
		         paramList.add( delete_mark);
		         paramList.add( muser_type ); 
		         paramList.add( password_pre );
		         paramList.add( user_id );				       
		         paramList.add( user_name );	
				 //   if(DBManager.updateDB(updateDBSqlList)){
			     if(this.updDbUsesPreparedStatement(sqlCmd,paramList)) {
					   errMsg = errMsg + "Y";					
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
	
	
	private String UpdateDB(HttpServletRequest request,String lguser_id,String lguser_name) throws Exception{    	
		String sqlCmd = "";		
		String errMsg="";		
		String muser_id=((String)request.getParameter("MUSER_ID")==null)?"":(String)request.getParameter("MUSER_ID");
		String muser_name=((String)request.getParameter("MUSER_NAME")==null)?"":(String)request.getParameter("MUSER_NAME");
		String bank_no=((String)request.getParameter("BANK_NO")==null)?" ":(String)request.getParameter("BANK_NO");
		bank_no = bank_no.equals("")?" ":bank_no;
		String subdep_id=((String)request.getParameter("SUBDEP_ID")==null)?"":(String)request.getParameter("SUBDEP_ID");		
		String lock_mark=((String)request.getParameter("LOCK_MARK")==null)?"":(String)request.getParameter("LOCK_MARK");
		String user_id=lguser_id;
	    String user_name=lguser_name;
	    System.out.println("muser_id="+muser_id);
		
		try {
				List paramList =new ArrayList() ;
				sqlCmd = "SELECT * FROM WTT01 WHERE muser_id=? ";
				paramList.add(muser_id) ;
			    List data = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");		 			    
				System.out.println("WTT01.size="+data.size());
				paramList.clear() ;
				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
				}else{    				   
				    //insert WTT01_LOG===================================================		    
				    sqlCmd = " INSERT INTO WTT01_LOG " 
						   + " select muser_id,muser_name,muser_password,muser_i_o,bank_type,"
						   + " tbank_no,bank_no,subdep_id,add_user,add_name,add_date,firstlogin_mark,login_mark,"
						   + " lock_mark,delete_mark,muser_type,password_update_date,password_pre,"
						   + " user_id,user_name,update_date"
						   + ",?,?,sysdate,'U'"
						   + " from WTT01"						  
						   + " WHERE muser_id=?";
				    paramList.add(user_id) ;
				    paramList.add(user_name) ;
				    paramList.add(muser_id) ;
				    this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
				    paramList.clear() ;
					//updateDBSqlList.add(sqlCmd);	
				    //=========================================================================
				    sqlCmd = "UPDATE WTT01 SET "
				    	   + " muser_name=?"				    	   						   
						   + ",bank_no=?" 
					       + ",subdep_id=?" 
					       + ",lock_mark=?"
					       + ",user_id=?"
					       + ",user_name=?"
					       + ",update_date=sysdate" 		            		 						       
						   + " where muser_id=?";				    	   
					paramList.add(muser_name)	;
					paramList.add(bank_no);
					paramList.add(subdep_id);
					paramList.add(lock_mark) ;
					paramList.add(user_id);
					paramList.add(user_name);
					paramList.add(muser_id);
	            		
					if(this.updDbUsesPreparedStatement(sqlCmd,paramList)) {
						errMsg = errMsg + "Y";					
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
    	   		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗<br>[Exception Error]:<br>"+e.getMessage();				
		}	

		return errMsg;
	} 
    
    private String DeleteDB(HttpServletRequest request,String lguser_id,String lguser_name) throws Exception{    	
		String sqlCmd = "";		
		String errMsg="";		
		String muser_id=((String)request.getParameter("MUSER_ID")==null)?"":(String)request.getParameter("MUSER_ID");
		String user_id=lguser_id;
	    String user_name=lguser_name;
	    
		
		try {
				List paramList =new ArrayList() ;
				sqlCmd = "SELECT * FROM WTT01 WHERE muser_id=? ";		
				paramList.add(muser_id) ;
			    List data = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");		 			    
				System.out.println("WTT01.size="+data.size());
				paramList.clear() ;
				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法刪除<br>";
				}else{    				   
				    //insert WTT01_LOG===================================================		    
				    sqlCmd = " INSERT INTO WTT01_LOG " 
						   + " select muser_id,muser_name,muser_password,muser_i_o,bank_type,"
						   + " tbank_no,bank_no,subdep_id,add_user,add_name,add_date,firstlogin_mark,login_mark,"
						   + " lock_mark,delete_mark,muser_type,password_update_date,password_pre,"
						   + " user_id,user_name,update_date"
						   + ",?,?,sysdate,'U'"
						   + " from WTT01"						  
						   + " WHERE muser_id=? "; 		
				    paramList.add(user_id) ;
				    paramList.add(user_name) ;
				    paramList.add(muser_id) ;
				    this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
				    paramList.clear() ;
				    //=========================================================================
				    /*
				    sqlCmd = "UPDATE WTT01 SET "				    	 
					       + " delete_mark='Y'"
					       + ",user_id='" + user_id +"'"
					       + ",user_name='" + user_name + "'"
					       + ",update_date=sysdate" 		            		 						       
						   + " where muser_id='"+muser_id+"'";				    	   
					*/
					//94.03.23 fix 直接將帳號刪除,一併將其所屬權限刪除================
					sqlCmd = "Delete WTT01 where muser_id=? ";		
					paramList.add(muser_id) ;
					this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
					paramList.clear() ;
		            //updateDBSqlList.add(sqlCmd); 		     
	                //94.04.07 fix WTT04有data時,才delete	                   			    	   						   		            
	            	sqlCmd = " select * from WTT04 where muser_id=?";
	                paramList.add(muser_id) ;
		  		    List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
		  		    paramList.clear() ;
		  		    if(dbData != null && dbData.size() != 0){
	            	   sqlCmd = "Delete WTT04 where muser_id=? ";
	            	   paramList.add(muser_id) ;
	            	   this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
	            	   paramList.clear() ;
		  		       //updateDBSqlList.add(sqlCmd); 	
		  		    }
		  		    //94.04.07 fix WTT04_1D有data時,才delete	                   			    	   						   		            
		  		    sqlCmd = " select * from WTT04_1D where muser_id=? "; 	
		  		    paramList.add(muser_id) ;
		  		    dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
		  		    paramList.clear() ;
		  		    if(dbData != null && dbData.size() != 0){
		  		       sqlCmd = "DELETE WTT04_1D WHERE muser_id=? ";
		  		       paramList.add(muser_id) ;
		  		       this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
				       //updateDBSqlList.add(sqlCmd); 		
				    }
	            	//===============================================================
					errMsg = errMsg + "Y";
	            	/*if(DBManager.updateDB(updateDBSqlList)){					 
						errMsg = errMsg + "相關資料寫入資料庫成功";					
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
					}*/
    	   		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗<br>[Exception Error]:<br>"+e.getMessage();						
		}	

		return errMsg;
	}
    
    private String ResetPwd(HttpServletRequest request,String lguser_id,String lguser_name) throws Exception{    	
		String sqlCmd = "";		
		String errMsg="";		
		String muser_id=((String)request.getParameter("MUSER_ID")==null)?"":(String)request.getParameter("MUSER_ID");
		String user_id=lguser_id;
	    String user_name=lguser_name;
	    String muser_password="";
		
		try {
				List paramList = new ArrayList() ;
				sqlCmd = "SELECT * FROM WTT01 WHERE muser_id=? ";
				paramList.add(muser_id) ;
			    List data = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");		 			    
				System.out.println("WTT01.size="+data.size());
				paramList.clear() ;
				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法重設密碼<br>";
				}else{    		
				    muser_password = (String)((DataObject)data.get(0)).getValue("muser_password");		   
				    //insert WTT01_LOG===================================================		    
				    sqlCmd = " INSERT INTO WTT01_LOG " 
						   + " select muser_id,muser_name,muser_password,muser_i_o,bank_type,"
						   + " tbank_no,bank_no,subdep_id,add_user,add_name,add_date,firstlogin_mark,login_mark,"
						   + " lock_mark,delete_mark,muser_type,password_update_date,password_pre,"
						   + " user_id,user_name,update_date"
						   + ",?,?,sysdate,'U'"
						   + " from WTT01"						  
						   + " WHERE muser_id=? ";
				    paramList.add(user_id) ;
				    paramList.add(user_name) ;
				    paramList.add(muser_id);
				    this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
				    paramList.clear() ;
				    //=========================================================================
				    sqlCmd = "UPDATE WTT01 SET "				    	 
					       + " muser_password=?"//96.03.22 add重設密碼時,將密碼加密
					       + ",firstlogin_mark='Y'"	 						       
					       + ",password_update_date=sysdate" 		            		 						       
					       + ",password_pre=?"
					       + ",user_id=?"
					       + ",user_name=?"
					       + ",update_date=sysdate" 		            		 						       
						   + " where muser_id=?";				    	   
					paramList.add(Utility.encode(muser_id) ) ;	   
					paramList.add(muser_password) ;
					paramList.add(user_id);
					paramList.add(user_name);
					paramList.add(muser_id) ;
					if(this.updDbUsesPreparedStatement(sqlCmd,paramList)){
						errMsg = errMsg + "Y";					
					}else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					}
    	   		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗"+e.getMessage();						
		}	

		return errMsg;
	}
    public String InsertWlXOPERATE_LOG(HttpServletRequest request,String lguser_id,String program_id,String upd_name,String update_type) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList() ;
		String muser_id=((String)request.getParameter("MUSER_ID")==null)?"":(String)request.getParameter("MUSER_ID");
		String errMsg="";
	    try {
	        sqlCmd.append(" INSERT INTO WlXOPERATE_LOG(muser_id,use_Date,program_id,ip_address,upd_name,update_type)");
	        sqlCmd.append("                     VALUES(?,sysdate,?,?,?,?) ");
	        paramList.add(lguser_id);
	        paramList.add(program_id);
	        paramList.add(request.getRemoteAddr());//ipAddress
	        paramList.add((upd_name=="")?muser_id:upd_name);//異動姓名(高階主管/負責人/理監事) 
	        paramList.add(update_type);//操作類別 I-新增，U-異動，D-刪除，Q-明細，P-列印
	        if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){
				errMsg = errMsg + "Y";					
			}else{
			    errMsg = errMsg + "相關資料寫入log失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
			}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入log失敗";						
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