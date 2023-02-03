<%
// 94.09.14 fix superBOAF控管帳號 by 2495
// 96.03.22 add 密碼解密 by 2295
// 96.03.22 add 總機構代碼使用XML by 2295
// 99.12.10 fix sqlInjection by 2808
//100.07.01 fix 查無資料 by 2295
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
      System.out.println("ZZBOAF login timeout");   
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
	System.out.println("act="+act);			
	System.out.println("nowact="+nowact);					
	System.out.println("bank_type="+bank_type);		
	System.out.println("tbank_no="+tbank_no);		
   
	//登入者資訊
	String lguser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
	String lguser_name = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");		
	String lguser_type = ( session.getAttribute("muser_type")==null ) ? "" : (String)session.getAttribute("muser_type");					
	
		
    if(!CheckPermission(request)){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );        
    }else{            
    	//set next jsp 	
    	if(act.equals("getData") || act.equals("List") || act.equals("Qry") || act.equals("unLock")){    	       	        	         	    
    	    List tbank_no_list = getTbank_No(bank_type);
    	    /*
    	    if(tbank_no.equals("") && (tbank_no_list.size() != 0)){
               tbank_no= (String)((DataObject)tbank_no_list.get(0)).getValue("bank_no");
            } */ 
    	    request.setAttribute("tbank_no",tbank_no_list);    	    
    	    List WTT01List = getQryResult(bank_type,tbank_no);    	    
    	    request.setAttribute("WTT01List",WTT01List);
        	if(act.equals("List")){
        	   rd = application.getRequestDispatcher( ListPgName +"?act=List");                	
        	}    
        	if(act.equals("Qry")){
        	   rd = application.getRequestDispatcher( ListPgName +"?act="+act+"&bank_type="+bank_type+"&tbank_no="+tbank_no);        
        	}
        	if(act.equals("getData")){        	    
        	    if(nowact.equals("List") || nowact.equals("Qry")){
        	       rd = application.getRequestDispatcher( ListPgName +"?act="+nowact+"&bank_type="+bank_type+"&tbank_no="+tbank_no);        
        	    }   
        	}

			//94.09.14 fix superBOAF控管帳號 by 2495 
        	if(act.equals("unLock")){
        	   //版本一:帳號登入與否個別控管
        	   String Star[] = request.getParameterValues("isModify");
						 for(int i=0;i<Star.length;i++)
						 {						      
									System.out.println("Star.length="+Star.length); 
									System.out.println("Star[i]="+Star[i]);       	   		
        	   		  setLoginMark(Star[i]);
        	   }    	   
        	   //版本二:帳號登入與否全部均設為N
        	   //setLoginMarkToN(); 
        	   WTT01List = getQryResult(bank_type,tbank_no);    	    
    	       request.setAttribute("WTT01List",WTT01List);
        	   rd = application.getRequestDispatcher( ListPgName +"?act="+act+"&bank_type="+bank_type+"&tbank_no="+tbank_no);        
        	}
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
    private final static String ListPgName = "/pages/ZZBOAF_List.jsp";        
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    private boolean CheckPermission(HttpServletRequest request){//檢核權限    	    
            /*
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
        	*/
        	return true;
    }           
    
    //取得總機構代碼
    private List getTbank_No(String bank_type){
    		//查詢條件    
    		String yy = Integer.parseInt(Utility.getYear()) > 99 ?"100" :"99" ;
    		List paramList = new ArrayList() ;
    		String sqlCmd = " select bank_no,bank_name from bn01 "
    					  + " where bank_type=? and m_year=?"
    		    		  + " order by bank_no";
    		paramList.add(bank_type) ;
    		paramList.add(yy) ;
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    } 
    
    //取得該管理者機構類別
    private List getBank_Type(String muser_id,String muser_type){
    		//查詢條件        		
    		String sqlCmd = "";
    		List paramList =new ArrayList() ;
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
    //2005 09 14 fix superBOAF控管帳號 by 2495 [多加  a.login_mark,]  
    private List getQryResult(String bank_type,String tbank_no){
    		//查詢條件        		
    		String sqlCmd = "";
    		String yy = Integer.parseInt(Utility.getYear()) > 99 ?"100" :"99" ;
    		List paramList = new ArrayList() ;
    		sqlCmd = "select b.bank_no as tbank_no,b.bank_name as tbank_name ,"
			       + "a.muser_id,a.muser_password,a.muser_name,a.muser_type,a.login_mark, "
				   + "b.bank_type,b.bank_no "
				   + "from wtt01 a LEFT JOIN (select * from bn01 where m_year=? )b on a.tbank_no = b.bank_no ";
    		paramList.add(yy) ;
			if(!bank_type.equals("ALL")){			
			    sqlCmd = sqlCmd + " where b.bank_type=? ";
			    paramList.add(bank_type) ;
			}
			if(!tbank_no.equals("ALL")){
			    sqlCmd = sqlCmd + " and b.bank_no=? ";
			    paramList.add(tbank_no) ;
			}
			
				   
			sqlCmd = sqlCmd + " order by b.bank_type,b.bank_no ";
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    }
	//94.09.14 fix superBOAF控管帳號 by 2495     
    private void setLoginMark(String muser_id){
      try{
	  //List updateDBSqlList = new LinkedList();  
	  List paramList =new ArrayList() ;  
	  String sqlCmd = "UPDATE WTT01 SET "
				    + " login_mark='N'"
  				    + " where muser_id=? ";			   
 	  paramList.add(muser_id) ;	  		
 	  this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
	  //updateDBSqlList.add(sqlCmd);                     
	  //DBManager.updateDB(updateDBSqlList,"");
	  }catch(Exception e){System.out.println("setLoginMark Error"+e+e.getMessage());}
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