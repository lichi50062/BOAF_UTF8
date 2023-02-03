<%
// 93.12.20 add 若有已點選的tbank_no,則以已點選的tbank_no為主 by 2295
// 		   add 權限檢核 by 2295
// 		   add 異動者資訊 by 2295
// 93.12.23 add 超過登入時間,請重新登入 by 2295
// 94.04.01 add 同一職務不可有一人以上擔任 by 2295
//          add 主key更改為bank_no+seq_no by 2295
// 94.04.06 fix 顯示裁撤日期的list by 2295
// 94.04.07 add 只統計未裁撤的分支機構 by 2295
// 94.04.12 add abdicate_code is null 也是未裁撤 by 2295
// 94.04.12 fix 不為裁撤時,將bn02.BN_TYPE='1'
// 95.06.05 add 將異動資料寫入WLX02_LOG/WLX02_M_LOG by 2295
// 99.12.03 fix sqlInjection by 2808
//100.01.26 fix 異動時,加上m_year及回傳訊息 by 2295
//102.04.24 add idn加解密  by2968
//102.06.28 add 操作歷程寫入log by2968
//102.12.18 fix 更新負責人ID資料,若無修改ID時,存入DB會變成已mask過的資料 by 2295
//103.06.24 fix 異動者資訊.增加區分99/100 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
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
      System.out.println("FX001W login timeout");   
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
	
	//登入者資訊==========================================================================================
	String lguser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
	String lguser_name = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");		
	String lguser_type = ( session.getAttribute("muser_type")==null ) ? "" : (String)session.getAttribute("muser_type");						
	//======================================================================================================================
	String act = ( request.getParameter("act")==null ) ? "" : (String)request.getParameter("act");			
	String bank_no = ( request.getParameter("bank_no")==null ) ? "" : (String)request.getParameter("bank_no");//分支機構代碼			
	String tbank_no = ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");//總機構代碼			
	String bank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");			
	String seq_no = ( request.getParameter("seq_no")==null ) ? "" : (String)request.getParameter("seq_no");			
	//String position_code = ( request.getParameter("position_code")==null ) ? "" : (String)request.getParameter("position_code");			
	//String id = ( request.getParameter("id")==null ) ? "" : (String)request.getParameter("id");				
	String nowtbank_no = "";
	System.out.println("act="+act);
	System.out.println("bank_no="+bank_no);
	System.out.println("FX002W.tbank_no="+tbank_no);
			
			
    if(!Utility.CheckPermission(request,"FX002W")){//無權限時,導向到LoginError.jsp 
        rd = application.getRequestDispatcher( LoginErrorPgName );        
    }else{            
    	//set next jsp 	
    	if(act.equals("List") || act.equals("Query")){ //for list and Query   	    
    	    //fix 93.12.20 若有已點選的tbank_no,則以已點選的tbank_no為主============================================================
			tbank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");				
			nowtbank_no =  ( request.getParameter("tbank_no")==null ) ? "" : (String)request.getParameter("tbank_no");			
			if(!nowtbank_no.equals("")){
	   			session.setAttribute("nowtbank_no",nowtbank_no);//將已點選的tbank_no寫入session	   
			}   
			tbank_no = ( session.getAttribute("nowtbank_no")==null ) ? tbank_no : (String)session.getAttribute("nowtbank_no");			
			//=======================================================================================================================	
			if(bank_type.equals("")){
	   		   bank_type=(String)session.getAttribute("nowbank_type");
		    }
		    List ba01Data = null;  
		    if(act.equals("List")){
    		   ba01Data = getBA01(tbank_no,"",bank_type);
    		}
    		if(act.equals("Query")){
    		   ba01Data = getBA01(tbank_no,bank_no,bank_type);
    		}
 			List constTypeData = getConstTypeCount(tbank_no,bank_type);
 			List HsiendIdData = getHsienIdCount(tbank_no,bank_type);
 			List bn02RevokeData = getBN02Revoke(tbank_no,bank_type);
    	    request.setAttribute("ba01Data",ba01Data);
    	    request.setAttribute("constTypeData",constTypeData);
    	    request.setAttribute("HsiendIdData",HsiendIdData);
    	    request.setAttribute("bn02RevokeData",bn02RevokeData);
        	rd = application.getRequestDispatcher( ListPgName );                
        }else if(act.equals("Edit")){ //分支機構編輯               
    	    List dbData = getWLX02(tbank_no,bank_no);
    	    List dbData1 = getWLX02_M(bank_no,"");
    	    List dbData_wlx01 = getWLX01(tbank_no);
    	    request.setAttribute("WLX02_M",dbData1);
    	    request.setAttribute("WLX02",dbData);
    	    request.setAttribute("WLX01",dbData_wlx01);
    	    //93.12.20設定異動者資訊======================================================================
    	    //103.06.24 fix 異動者資訊.增加區分99/100 by 2295
    	    String yy = Integer.parseInt(Utility.getYear())>99  ?"100" : "99" ;
			request.setAttribute("maintainInfo","select * from WLX02 WHERE tbank_no='" + tbank_no+ "' and bank_no='"+bank_no+"' and m_year="+yy);								       
			//=======================================================================================================================		        	
        	//操作歷程寫入log
			this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,tbank_no,bank_no,"Q");
			rd = application.getRequestDispatcher( EditWLX02PgName );        
        }else if(act.equals("Update")){//分支機構修改    	        	    
       	    actMsg = UpdateWLX02(request,tbank_no,bank_no,lguser_id,lguser_name);
        	if("Y".equals(actMsg)){
	       		//操作歷程寫入log
				actMsg = this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,tbank_no,bank_no,"U");
        		if("Y".equals(actMsg)){
        		    actMsg = "相關資料寫入資料庫成功";
        		}
        	}
        	rd = application.getRequestDispatcher( nextPgName );        		
    	}else if(act.equals("Revoke")){//分支機構裁撤    	    
    	    actMsg = RevokeWLX02(request,tbank_no,bank_no,lguser_id,lguser_name); 
    	    if("Y".equals(actMsg)){
	       		//操作歷程寫入log
				//actMsg = this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,tbank_no,bank_no,"U");
        		//if("Y".equals(actMsg)){
        		    actMsg = "執行裁撤分部成功";
        		//}
        	}
        	rd = application.getRequestDispatcher( nextPgName );            		
    	}else if(act.equals("newM")){//負責人新增    	    
        	rd = application.getRequestDispatcher( EditWLX02_MPgName +"?act=newM&bank_no="+bank_no);        
    	}else if(act.equals("EditM")){//負責人編輯    	    
    	    List dbData = getWLX02_M(bank_no,seq_no);
    	    request.setAttribute("WLX02_M",dbData);
    	    //93.12.20設定異動者資訊======================================================================
			request.setAttribute("maintainInfo","select * from WLX02_M WHERE bank_no='" + bank_no+"' and seq_no="+seq_no);								       
			//=======================================================================================================================		        	
        	//操作歷程寫入log
			this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,tbank_no,bank_no,"Q");
			rd = application.getRequestDispatcher( EditWLX02_MPgName +"?act=EditM");        
    	}else if(act.equals("InsertM")){//負責人新增    	        	    
       	    actMsg = InsertWLX02_M(request,bank_no,lguser_id,lguser_name);
	    	if("Y".equals(actMsg)){
	    	  	//操作歷程寫入log
				actMsg = this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,tbank_no,bank_no,"I");
				if("Y".equals(actMsg)){
        		    actMsg = "相關資料寫入資料庫成功";
        		}
	    	}
        	rd = application.getRequestDispatcher( nextPgName );        		
    	}else if(act.equals("UpdateM")){//負責人修改    	        	    
       	     actMsg = UpdateWLX02_M(request,bank_no,seq_no,lguser_id,lguser_name);
	       	 if("Y".equals(actMsg)){
	       		//操作歷程寫入log
				actMsg = this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,tbank_no,bank_no,"U");
				if("Y".equals(actMsg)){
        		    actMsg = "相關資料寫入資料庫成功";
        		}
	       	 }
        	 rd = application.getRequestDispatcher( nextPgName );        		
    	}else if(act.equals("DeleteM")){//負責人刪除    	        	    
       	     actMsg = DeleteWLX02_M(request,bank_no,seq_no,lguser_id,lguser_name);
	       	 if("Y".equals(actMsg)){
	       		//操作歷程寫入log
				actMsg = this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,tbank_no,bank_no,"D");
				if("Y".equals(actMsg)){
				    actMsg = "相關資料刪除成功";
				}
	       	 }
        	 rd = application.getRequestDispatcher( nextPgName );        		
    	}else if(act.equals("AbdicateM")){//負責人卸任    	    
    	    actMsg = AbdicateWLX02_M(request,bank_no,seq_no,lguser_id,lguser_name);
    	    if("Y".equals(actMsg)){
	       		//操作歷程寫入log
				actMsg = this.InsertWlXOPERATE_LOG(request,lguser_id,program_id,tbank_no,bank_no,"U");
        		if("Y".equals(actMsg)){
        		    actMsg = "執行卸任成功";
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


<%!
	private final static String program_id = "FX002W";
    private final static String nextPgName = "/pages/ActMsg.jsp";    
    private final static String EditWLX02PgName = "/pages/"+program_id+"_Edit.jsp";
    private final static String EditWLX02_MPgName = "/pages/"+program_id+"_EditM.jsp";
    private final static String ListPgName = "/pages/"+program_id+"_List.jsp";        
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
       
    private List getWLX01(String bank_no){
    		//查詢條件    		
    		List paramList = new ArrayList() ;
    		String sqlCmd = "select  cancel_no from WLX01 where bank_no=? ";
    		sqlCmd += " and m_year=? ";
    		paramList.add(bank_no) ;
    		paramList.add(Integer.parseInt(Utility.getYear())>99 ?"100":"99") ;
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    }
    
    private List getWLX02(String tbank_no,String bank_no){
    		//查詢條件    		
    		//String sqlCmd = "select * from WLX02,ba01 where WLX02.bank_no = ba01.bank_no and WLX02.tbank_no='"+tbank_no+"' and WLX02.bank_no='"+bank_no+"'";    		
    		List paramList = new ArrayList() ;
    		String yy = Integer.parseInt(Utility.getYear())>99 ?"100":"99" ;
    		String sqlCmd = " select * from (select * from ba01 where m_year=?)ba01 " 
						  + " LEFT JOIN (select * from wlx02 where m_year=?)WLX02 on WLX02.bank_no = ba01.bank_no "
						  + " and WLX02.tbank_no=?"
						  + " and WLX02.bank_no=?" 
						  + " where ba01.PBANK_NO=?"
						  + " and ba01.BANK_NO=? "
						  + " and ba01.BANK_KIND=? "; 
    		paramList.add(yy) ;
    		paramList.add(yy) ;
    		paramList.add(tbank_no)  ;
    		paramList.add(bank_no) ;
    		paramList.add(tbank_no); 
    		paramList.add(bank_no);
    		paramList.add("1") ;
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"setup_date,setup_no_date,chg_license_date,start_date,staff_num,open_date,cancel_date,update_date");            
            return dbData;
    }
    
    //94.04.01 主key更改為bank_no+seq_no 
    private List getWLX02_M(String bank_no,String seq_no){
    		//查詢條件    		
    		List paramList = new ArrayList() ;
    		String sqlCmd = "select * from WLX02_M,cdshareno where bank_no=? ";
    		paramList.add(bank_no) ;
    		if(!seq_no.equals("")){			
    			sqlCmd = sqlCmd + " and seq_no= ?";
    			paramList.add(seq_no) ;
    		}
    		sqlCmd = sqlCmd + "and wlx02_M.POSITION_CODE = cdshareno.CMUSE_ID and cdshareno.CMUSE_DIV='007'";
    		sqlCmd =sqlCmd + " order by position_code";				
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"birth_date,induct_date,abdicate_date,rank,seq_no");            
            return dbData;
    } 
    //94.04.07 add 只取得未裁撤的分支機構
    private List getBA01(String tbank_no,String bank_no,String bank_type){
            
    		//查詢條件
    		List paramList = new ArrayList() ;
    		String yy = Integer.parseInt(Utility.getYear())>99 ?"100":"99" ;
    		String sqlCmd = " select ba01.pbank_no,ba01.bank_no,ba01.bank_name from (select * from ba01 where m_year=?)ba01 "
    		              + " LEFT JOIN (select * from bn02 where m_year=?)bn02 on ba01.bank_no = bn02.bank_no "     		              
    					  + " where pbank_no=?"
					      + " and ba01.bank_type=? "
						  + " and ba01.bank_kind='1'"
						  + " and (bn02.bn_type <> '2' or bn02.bn_type is null)";
    		paramList.add(yy) ;
    		paramList.add(yy);
    		paramList.add(tbank_no) ;
    		paramList.add(bank_type) ;
			if(!bank_no.equals("")){
			   sqlCmd =sqlCmd + " and bank_no like ? ";
			   paramList.add(bank_no+"%") ;
			}
			List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
			
			/*
			//查詢條件    		
    		String sqlCmd = " select * from ba01 "
    					  + " where pbank_no='"+tbank_no+"'"
					      + " and bank_type='"+bank_type+"'"
						  + " and bank_kind='1'";
			if(!bank_no.equals("")){
			   sqlCmd =sqlCmd + " and bank_no like '"+bank_no+"%'";
			}			  
            List dbData = DBManager.QueryDB(sqlCmd,"");            
            return dbData;			  
            */
    }
    
    private List getBN02Revoke(String tbank_no,String bank_type){
    		//查詢條件    		
    		List paramList = new ArrayList() ;
    		String yy = Integer.parseInt(Utility.getYear()) > 99 ? "100" :"99" ;
    		String sqlCmd = " select bn02.tbank_no,bn02.bank_no,bn02.bank_name,wlx02.cancel_date from (select * from bn02 where m_year=?)bn02 "
    		              + " LEFT JOIN (select * from wlx02 where m_year=?)wlx02 on bn02.tbank_no = wlx02.tbank_no and bn02.bank_no = wlx02.bank_no "
    		 			  + " where bn02.tbank_no=?"
    		 			  + " and bn02.bank_type=?"
    		 			  + " and bn02.bn_type = '2'";
    		paramList.add(yy) ;
    		paramList.add(yy) ;
    		paramList.add(tbank_no) ;
    		paramList.add(bank_type) ;
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"cancel_date");            
            return dbData;
    }
    //94.04.07 add 只統計未裁撤的分支機構
    private List getConstTypeCount(String tbank_no,String bank_type){
    		//查詢條件    	
    		List paramList = new ArrayList() ;
    		String yy = Integer.parseInt(Utility.getYear()) > 99 ? "100" :"99" ;
    		String sqlCmd = "select sum(decode(wlx02.const_type,'1',1,0)) as const_type1 "
    					  + ",sum(decode(wlx02.CONST_TYPE,'2',1,0)) as const_type2"
    					  + ",sum(decode(wlx02.const_type,'9',1,0)) as const_type9"
    					  + ",count(*) as const_tpyecount"    					  			      
						  + " from (select * from wlx02 where m_year=?)wlx02,(select * from ba01 where m_year=?)ba01 "
					      + " where wlx02.bank_no=ba01.bank_no"
						  + " and ba01.pbank_no=? "
					      + " and ba01.bank_type=? "
						  + " and ba01.bank_kind='1'"
						  + " and (wlx02.cancel_no <> 'Y' or wlx02.cancel_no is null)";
    		paramList.add(yy);
    		paramList.add(yy);
    		paramList.add(tbank_no) ;
    		paramList.add(bank_type);
			List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"const_type1,const_type2,const_type9,const_tpyecount");            
            return dbData;
    } 
    //94.04.07 add 只統計未裁撤的分支機構
	private List getHsienIdCount(String tbank_no,String bank_type){
    		//查詢條件    		
    		List paramList = new ArrayList() ;
    		String yy = Integer.parseInt(Utility.getYear()) > 99 ? "100" :"99" ;
    		String cd01Table = Integer.parseInt(Utility.getYear()) > 99 ? "cd01" :"cd01_99" ;
 			String sqlCmd = "select wlx02.hsien_id,cd01.hsien_name"
 						  + ",sum(decode(wlx02.const_type,'1',1,0)) as const_type1"
 						  + ",sum(decode(wlx02.CONST_TYPE,'2',1,0)) as const_type2"
 						  + ",sum(decode(wlx02.const_type,'9',1,0)) as const_type9"
 						  + ",count(*) as const_typecount"
						  + " from (select * from wlx02 where m_year=?)wlx02 ,(select * from ba01 where m_year=?)ba01,"+cd01Table+" cd01"
						  +	" where wlx02.bank_no=ba01.bank_no"
						  + " and wlx02.hsien_id=cd01.hsien_id"
						  + " and ba01.pbank_no=? "
					      + " and ba01.bank_type=? "
						  + " and ba01.bank_kind='1'"
						  + " and (wlx02.cancel_no <> 'Y' or wlx02.cancel_no is null)"
						  + " group by wlx02.hsien_id,cd01.hsien_name"
						  + " order by wlx02.hsien_id,cd01.hsien_name";
 			paramList.add(yy) ;
 			paramList.add(yy);
 			paramList.add(tbank_no) ;
 			paramList.add(bank_type);
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"const_type1,const_type2,const_type9,const_typecount");            
            return dbData;
    } 
     
    	
    //100.01.26 寫入/更新資料時,增加m_year by 2295
	private String UpdateWLX02(HttpServletRequest request,String tbank_no,String bank_no,String lguser_id,String lguser_name) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer();		
		String errMsg="";		
		
		String const_type=((String)request.getParameter("CONST_TYPE")==null)?"":(String)request.getParameter("CONST_TYPE");		
		String setup_approval_unt=((String)request.getParameter("SETUP_APPROVAL_UNT")==null)?"":(String)request.getParameter("SETUP_APPROVAL_UNT");//93.12.21 add		
		String setup_date=(String)request.getParameter("SETUP_DATE");
		String setup_no=((String)request.getParameter("SETUP_NO")==null)?"":(String)request.getParameter("SETUP_NO");
		String setup_no_date=(String)request.getParameter("SETUP_NO_DATE");
		String chg_license_date=(String)request.getParameter("CHG_LICENSE_DATE");
		String chg_license_no=((String)request.getParameter("CHG_LICENSE_NO")==null)?"":(String)request.getParameter("CHG_LICENSE_NO");
		String chg_license_reason=((String)request.getParameter("CHG_LICENSE_REASON")==null)?"":(String)request.getParameter("CHG_LICENSE_REASON");
		String start_date=(String)request.getParameter("START_DATE");
		String hsien_id_area_id=((String)request.getParameter("HSIEN_ID_AREA_ID")==null)?"":(String)request.getParameter("HSIEN_ID_AREA_ID");
		String hsien_id = hsien_id_area_id.substring(0,hsien_id_area_id.indexOf("/"));
		String area_id = hsien_id_area_id.substring(hsien_id_area_id.indexOf("/")+1,hsien_id_area_id.length());
		String addr=((String)request.getParameter("ADDR")==null)?"":(String)request.getParameter("ADDR");
		String telno=((String)request.getParameter("TELNO")==null)?"":(String)request.getParameter("TELNO");
		String fax=((String)request.getParameter("FAX")==null)?"":(String)request.getParameter("FAX");
		String email=((String)request.getParameter("EMAIL")==null)?"":(String)request.getParameter("EMAIL");
		String web_site=((String)request.getParameter("WEB_SITE")==null)?"":(String)request.getParameter("WEB_SITE");
		String flag="";
		String open_date=(String)request.getParameter("OPEN_DATE");		
		String staff_num=((String)request.getParameter("STAFF_NUM") == null || ((String)request.getParameter("STAFF_NUM")).equals(""))?"0" : (String)request.getParameter("STAFF_NUM");
		String hsien_div_1=((String)request.getParameter("HSIEN_DIV")==null)?"":(String)request.getParameter("HSIEN_DIV");		
		String cancel_no=(String)request.getParameter("CANCEL_NO");
		String cancel_date=(String)request.getParameter("CANCEL_DATE");
		String user_id=lguser_id;
		String user_name=lguser_name;
		System.out.println("SETUP_NO_DATE="+setup_no_date);
		try {
			    //List updateDBSqlList = new LinkedList();
			    List paramList = new ArrayList() ;
			    String yy = Integer.parseInt(Utility.getYear()) > 99 ? "100" :"99" ;
			    sqlCmd.append("SELECT * FROM WLX02 where tbank_no=? and bank_no=? and m_year=?");				    	   						   				 
				paramList.add(tbank_no) ;
				paramList.add(bank_no);
				paramList.add(yy) ;
			   	List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");		 			    
				//System.out.println("WLX02.size="+data.size());
				
				if (data.size() == 0){//無資料時,Insert
					sqlCmd.setLength(0) ;
				    paramList.clear() ;
					sqlCmd.append("INSERT INTO WLX02 VALUES (?,?,?,?,to_date(?,'YYYY/MM/DD') ,");
					sqlCmd.append("?,to_date(?,'YYYY/MM/DD'),to_date(?,'YYYY/MM/DD'),?,?,");
					sqlCmd.append("to_date(?,'YYYY/MM/DD'),?,?,?,?,");
					sqlCmd.append("?,?,?,?,to_date(?,'YYYY/MM/DD'),");
					sqlCmd.append("?,?,?,to_date(?,'YYYY/MM/DD'),?,?,sysdate,?) ") ;
				    paramList.add(tbank_no) ;
				    paramList.add(bank_no);
					paramList.add(const_type );
	       			paramList.add(setup_approval_unt); 				    	   
					paramList.add(setup_date);  
					paramList.add(setup_no); 
					paramList.add(setup_no_date );  
					paramList.add(chg_license_date);
					paramList.add(chg_license_no ) ; 
					paramList.add(chg_license_reason ); 
					paramList.add(start_date); 
					paramList.add(hsien_id ); 
					paramList.add(area_id ); 
					paramList.add(addr ); 
					paramList.add(telno); 					       
					paramList.add( fax ); 
					paramList.add(email ); 
					paramList.add(web_site); 
					paramList.add(flag);
					paramList.add(open_date) ;
					paramList.add(staff_num); 
					paramList.add(hsien_div_1);
					paramList.add( cancel_no);
					paramList.add(cancel_date);  
					paramList.add(user_id );
					paramList.add(user_name ) ;
					paramList.add(yy);//100.01.26 add
					if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){	   				      		 
						errMsg += "相關資料寫入資料庫成功";					
					}else{
				 	    errMsg += "相關資料寫入資料庫失敗";
					}
				}else{//Update			
					//95.06.05 add 寫入WLX02_LOG ===================================================================
					sqlCmd.setLength(0) ;
				    paramList.clear() ;
                    sqlCmd.append(" INSERT INTO WLX02_LOG "
                           + " select TBANK_NO,BANK_NO,CONST_TYPE,SETUP_APPROVAL_UNT,SETUP_DATE,SETUP_NO,"
                           + "	  	  SETUP_NO_DATE,CHG_LICENSE_DATE,CHG_LICENSE_NO,CHG_LICENSE_REASON,"
				           + "		  START_DATE,HSIEN_ID,AREA_ID,ADDR,TELNO,FAX,EMAIL,WEB_SITE,FLAG,"
				           + "		  OPEN_DATE,STAFF_NUM,HSIEN_DIV_1,CANCEL_NO,CANCEL_DATE,"
						   + "		  USER_ID,USER_NAME,UPDATE_DATE, "
						   + "        ?,?,sysdate,'U'"
						   + " from WLX02"
						   + " where tbank_no=? and bank_no=? and m_year=? ");
                    paramList.add(user_id) ;
                    paramList.add(user_name) ;
                    paramList.add(tbank_no) ;
                    paramList.add(bank_no) ;
                    paramList.add(yy);//100.01.26 add
					//updateDBSqlList.add(sqlCmd);
					if(!this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){	   				      		 
							errMsg += "WLX02_LOG寫入失敗";
					}
					//==================================================================================================
				 	sqlCmd.setLength(0) ;
				    paramList.clear() ;
				    sqlCmd.append("UPDATE WLX02 SET ");		
				    sqlCmd.append(" const_type=?");paramList.add(const_type) ;
		    	    sqlCmd.append(",setup_approval_unt=?"); paramList.add( setup_approval_unt); 				    	   
			        sqlCmd.append(",setup_date = to_date(?,'YYYY/MM/DD')" );paramList.add(setup_date) ; 
			        sqlCmd.append(",setup_no=?"); paramList.add( setup_no );
			        sqlCmd.append(",setup_no_date = to_date(?,'YYYY/MM/DD')");paramList.add(setup_no_date) ;  
			        sqlCmd.append(",chg_license_date = to_date(?,'YYYY/MM/DD')");paramList.add(chg_license_date) ;  
			        sqlCmd.append(",chg_license_no=?"); paramList.add( chg_license_no); 
			        sqlCmd.append(",chg_license_reason=?"); paramList.add( chg_license_reason );
			        sqlCmd.append(",start_date = to_date(?,'YYYY/MM/DD')");paramList.add(start_date) ; 
			        sqlCmd.append(",hsien_id=?");paramList.add(hsien_id); 
			        sqlCmd.append(",area_id=?");paramList.add(area_id); 
			        sqlCmd.append(",addr=?"); paramList.add( addr); 
			        sqlCmd.append(",telno=?"); paramList.add( telno ); 					       
			        sqlCmd.append(",fax=?"); paramList.add( fax );
			        sqlCmd.append(",email=?"); paramList.add( email); 
			        sqlCmd.append(",web_site=?"); paramList.add( web_site); 
			        sqlCmd.append(",flag=?"); paramList.add( flag);
			        sqlCmd.append(",open_date=to_date(?,'YYYY/MM/DD')");paramList.add(open_date);  
			        sqlCmd.append(",staff_num=?");paramList.add(staff_num); 
			        sqlCmd.append(",hsien_div_1=?"); paramList.add( hsien_div_1 );
			        sqlCmd.append(",cancel_no=?"); paramList.add( cancel_no);
			        sqlCmd.append(",cancel_date=to_date(?,'YYYY/MM/DD')");paramList.add(cancel_date) ;  
			        sqlCmd.append(",user_id=?"); paramList.add( user_id ); 
			        sqlCmd.append(",user_name=?"); paramList.add( user_name); 
			        sqlCmd.append(",update_date=sysdate"); 		            		 					    				    
		    	    sqlCmd.append(" where tbank_no=?");paramList.add(tbank_no);
		    	    sqlCmd.append(" and bank_no=?");paramList.add(bank_no);
		    	    sqlCmd.append(" and m_year=?");paramList.add(yy);//100.01.26 add
		    	    if(!this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){	   				      		 					  
				 	   errMsg += "WLX02相關資料寫入資料庫失敗";
					}
				    	    	   
				    if(cancel_no.equals("Y")){//修改時,選裁撤
				    	sqlCmd.setLength(0) ;
				        paramList.clear() ;
					    sqlCmd.append(" UPDATE bn02 SET " 
		             		          + " BN_TYPE='2' " 
		             		          + ",user_id=? "
					                  + ",user_name=? "
                         	          + ",update_date=sysdate "
                                      + " where tbank_no=? and bank_no=? and m_year=?");
					    paramList.add(user_id) ;
					    paramList.add(user_name) ;
					    paramList.add(tbank_no) ;
					    paramList.add(bank_no);
					    paramList.add(yy);//100.01.26 add
                        //updateDBSqlList.add(sqlCmd1);
                        if(!this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){	   				      		 
							errMsg += "分支機構裁撤失敗";
					 	 }
                    }else{//94.04.12 fix 若修改時,選不裁撤,則恢復為bn_type='1'
                    	sqlCmd.setLength(0) ;
				        paramList.clear() ;
                        sqlCmd.append(" UPDATE bn02 SET " 
		             		          + " BN_TYPE='1' " 
		             		          + ",user_id=? "
					                  + ",user_name=? "
                         	          + ",update_date=sysdate "
                                      + " where tbank_no=?  and bank_no=? and m_year=?");
                        paramList.add(user_id) ;
                        paramList.add(user_name) ;
                        paramList.add(tbank_no) ;
                        paramList.add(bank_no) ;
                        paramList.add(yy);//100.01.26 add
                        //updateDBSqlList.add(sqlCmd1); 
                        if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){	   				      		 
					   		errMsg += "Y";					
					    }else{
				 	   		errMsg += "相關資料寫入資料庫失敗";
					     }
                    }
				}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";							
		}	

		return errMsg;
	}     
    //100.01.26異動時,加上m_year及回傳訊息 by 2295
	private String RevokeWLX02(HttpServletRequest request,String tbank_no,String bank_no,String lguser_id,String lguser_name) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer() ;
		List paramList = new ArrayList() ;
		String errMsg="";		
		String cancel_no=(String)request.getParameter("CANCEL_NO");
		String cancel_date=(String)request.getParameter("CANCEL_DATE");
		String user_id=lguser_id;
		String user_name=lguser_name;
		
		try {
				//List updateDBSqlList = new LinkedList();
				String yy = Integer.parseInt(Utility.getYear()) > 99 ? "100" :"99" ;
				sqlCmd.append("SELECT  tbank_no,bank_no FROM WLX02 WHERE tbank_no=? AND bank_no=? and m_year=?");
				paramList.add(tbank_no) ;
				paramList.add(bank_no);
				paramList.add(yy) ;
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");		 			    
				System.out.println("WLX04.size="+data.size());
				
				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法裁撤<br>";
				}else{    
				    //95.06.05 add 寫入WLX02_LOG ===================================================================
				    sqlCmd.setLength(0) ;
				    paramList.clear() ;
                    sqlCmd.append(" INSERT INTO WLX02_LOG "
                           + " select TBANK_NO,BANK_NO,CONST_TYPE,SETUP_APPROVAL_UNT,SETUP_DATE,SETUP_NO,"
                           + "	  	  SETUP_NO_DATE,CHG_LICENSE_DATE,CHG_LICENSE_NO,CHG_LICENSE_REASON,"
				           + "		  START_DATE,HSIEN_ID,AREA_ID,ADDR,TELNO,FAX,EMAIL,WEB_SITE,FLAG,"
				           + "		  OPEN_DATE,STAFF_NUM,HSIEN_DIV_1,CANCEL_NO,CANCEL_DATE,"
						   + "		  USER_ID,USER_NAME,UPDATE_DATE, "
						   + "        ?,?,sysdate,'R'"//revoke
						   + " from WLX02"
						   + " where tbank_no=? and bank_no=? and m_year=?");
                    paramList.add(user_id) ;
                    paramList.add(user_name) ;
                    paramList.add(tbank_no);
                    paramList.add(bank_no);
                    paramList.add(yy);//100.01.26 add
					//updateDBSqlList.add(sqlCmd); 		
					if(!this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){	   				      		 
					   errMsg += "WLX02_LOG寫入失敗";
					}
					//==================================================================================================
				 	sqlCmd.setLength(0) ;
					paramList.clear() ;
				    sqlCmd.append(" UPDATE WLX02 SET cancel_no=?,cancel_date=to_date(?,'YYYY/MM/DD')"			
				    	   + " where tbank_no=? and bank_no=? and m_year=?");
				    paramList.add(cancel_no) ;
				    paramList.add(cancel_date) ;
				    paramList.add(tbank_no) ;
				    paramList.add(bank_no) ;
				    paramList.add(yy);//100.01.26 add
				    if(!this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){	   				      		 
					   errMsg += "WLX02寫入失敗";
					}
		            //updateDBSqlList.add(sqlCmd); 		            		            		
		            sqlCmd.setLength(0) ;
					paramList.clear() ;
		            sqlCmd.append(" UPDATE bn02 SET " 
		             		+ " BN_TYPE='2' " 
		             		+ ",user_id=? "
					        + ",user_name=? "
                         	+ ",update_date=sysdate "
                            + " where tbank_no=?  and bank_no=? and m_year=?");
		            paramList.add(user_id) ;
		            paramList.add(user_name);
		            paramList.add(tbank_no) ;
		            paramList.add(bank_no) ;
		            paramList.add(yy);//100.01.26 add
		            if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){					 
						errMsg = errMsg + "Y";					
					}else{
				     	errMsg = errMsg + "執行裁撤分部失敗";
				    }
                                		            		                    
                    /*
                    sqlCmd = " INSERT INTO bn04 SELECT " +
						  	 "tbank_no, bank_no, bank_name, bn_date, bn_address, '2'," +
						  	 "bank_type, '" + User_Id + "', '' , Current " +
						  	 "FROM bn02 WHERE tbank_no ='" + Bank_No + "' AND bank_no='" + szBBANK_NO + "'";
				     */				  
    	   		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "執行裁撤分部失敗";						
		}	

		return errMsg;
	}
    
    //94.04.01 主key更改為bank_no+seq_no 
    private String InsertWLX02_M(HttpServletRequest request,String bank_no,String lguser_id,String lguser_name) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer();		
		String errMsg="";		
		String name=((String)request.getParameter("NAME")==null)?"":(String)request.getParameter("NAME");
		String birth_date=(String)request.getParameter("BIRTH_DATE");
		String degree=((String)request.getParameter("DEGREE")==null)?"":(String)request.getParameter("DEGREE");
		String sex=((String)request.getParameter("SEX")==null)?"":(String)request.getParameter("SEX");
		String telno=((String)request.getParameter("TELNO")==null)?"":(String)request.getParameter("TELNO");;
		String induct_date=(String)request.getParameter("INDUCT_DATE");
		String background=((String)request.getParameter("BACKGROUND")==null)?"":(String)request.getParameter("BACKGROUND");
		String choose_item="";
		String rank=((String)request.getParameter("RANK") == null || ((String)request.getParameter("RANK")).equals(""))?"0" : (String)request.getParameter("RANK");
		String abdicate_code=(String)request.getParameter("ABDICATE_CODE");
		String abdicate_date=(String)request.getParameter("ABDICATE_DATE");
		String email=((String)request.getParameter("EMAIL")==null)?"":(String)request.getParameter("EMAIL");;
		String user_id=lguser_id;
		String user_name=lguser_name;
		String seq_no="";
	    String id_code=((String)request.getParameter("ID_CODE")==null)?"N":(String)request.getParameter("ID_CODE");		
	    String position_code = ( request.getParameter("POSITION_CODE")==null ) ? "" : (String)request.getParameter("POSITION_CODE");			
	    String id = ( request.getParameter("ID")==null ) ? "" : Utility.encode((String)request.getParameter("ID"));			

		
		try {
				//List updateDBSqlList = new LinkedList();
				List paramList = new ArrayList() ;
				sqlCmd.append(" SELECT count(*) as have_cnt FROM WLX02_M "
				       + " WHERE bank_no=? "
				       + " AND position_code=? "
					   + " AND (abdicate_code <> 'Y' or abdicate_code is null)");
				paramList.add(bank_no) ;
				paramList.add(position_code) ;
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"have_cnt");		 			    
				System.out.println("WLX02_M.size="+data.size());
				System.out.println("have_cnt="+(((DataObject)data.get(0)).getValue("have_cnt")).toString());
				
				if(!((((DataObject)data.get(0)).getValue("have_cnt")).toString()).equals("0")){
				    errMsg = errMsg + "該職務已建檔,無法新增<br>";				    
				}else{   
					sqlCmd.setLength(0) ;
					paramList.clear() ;
				    List dbData = DBManager.QueryDB_SQLParam("SELECT to_char(wlx02_m_seqno.NEXTVAL) seq_no FROM DUAL",paramList,"");
                    seq_no = (String)((DataObject)dbData.get(0)).getValue("seq_no");	
                    paramList.add(bank_no) ;
                    paramList.add(position_code) ;
                    paramList.add(id) ;
                    paramList.add(name) ;
                    paramList.add(birth_date) ;
                    paramList.add(degree) ;
                    paramList.add(sex) ;
                    paramList.add(telno) ;
                    paramList.add(induct_date) ;
                    paramList.add(background) ;
                    paramList.add(choose_item) ;
                    paramList.add(rank) ;
                    paramList.add(abdicate_code) ; 
                    paramList.add(abdicate_date) ;
                    paramList.add(email) ;
                    paramList.add(user_id) ;
                    paramList.add(user_name) ; 
                    paramList.add(seq_no) ;
                    paramList.add(id_code) ;
                    sqlCmd.append(" insert into wlx02_m values( ?,?,?,?,to_date(?,'YYYY/MM/DD'),") ;
                    sqlCmd.append("?,?,?,to_date(?,'YYYY/MM/DD'),");
                    sqlCmd.append("?,?,?,?,to_date(?,'YYYY/MM/DD'),");
                    sqlCmd.append("?,?,?,sysdate,?,? )");
					/*sqlCmd.append("INSERT INTO WLX02_M VALUES ('" + bank_no + "','" + position_code + "','" + id +"'"
					       + ",'" + name + "'" 
					       + ",to_date('"+birth_date+"','YYYY/MM/DD')"  
					       + ",'" + degree + "'" 
					       + ",'" + sex + "'" 
					       + ",'" + telno + "'"
					       + ",to_date('"+induct_date+"','YYYY/MM/DD')" 
					       + ",'" + background +"'"
					       + ",'" + choose_item+"'"
					       + ","+rank 					       
					       + ",'" + abdicate_code + "'" 
					       + ",to_date('"+abdicate_date+"','YYYY/MM/DD')"  					       
					       + ",'" + email + "'"
					       + ",'" + user_id +"'"
					       + ",'" + user_name + "'"
					       + ",sysdate"
					       + ","+seq_no
					       + ",'"+id_code+"')"); 				            		 	
					       
		            updateDBSqlList.add(sqlCmd); 		            	
	            	*/	
					if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){					 
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
	//94.04.01 主key更改為bank_no+seq_no 
	//102.12.18 fix 更新負責人ID資料,若無修改ID時,存入DB會變成已mask過的資料 by 2295
	private String UpdateWLX02_M(HttpServletRequest request,String bank_no,String seq_no,String lguser_id,String lguser_name) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer() ;
		List paramList =new ArrayList() ;
		String errMsg="";		
		String name=((String)request.getParameter("NAME")==null)?"":(String)request.getParameter("NAME");
		String birth_date=(String)request.getParameter("BIRTH_DATE");
		String degree=((String)request.getParameter("DEGREE")==null)?"":(String)request.getParameter("DEGREE");
		String sex=((String)request.getParameter("SEX")==null)?"":(String)request.getParameter("SEX");
		String telno=((String)request.getParameter("TELNO")==null)?"":(String)request.getParameter("TELNO");;
		String induct_date=(String)request.getParameter("INDUCT_DATE");
		String background=((String)request.getParameter("BACKGROUND")==null)?"":(String)request.getParameter("BACKGROUND");
		String choose_item="";
		String rank=((String)request.getParameter("RANK") == null || ((String)request.getParameter("RANK")).equals(""))?"0" : (String)request.getParameter("RANK");
		String abdicate_code=(String)request.getParameter("ABDICATE_CODE");
		String abdicate_date=(String)request.getParameter("ABDICATE_DATE");
		String email=((String)request.getParameter("EMAIL")==null)?"":(String)request.getParameter("EMAIL");;
		String user_id=lguser_id;
		String user_name=lguser_name;
		String id_code=((String)request.getParameter("ID_CODE")==null)?"N":(String)request.getParameter("ID_CODE");			   
	    String position_code = ( request.getParameter("POSITION_CODE")==null ) ? "" : (String)request.getParameter("POSITION_CODE");			
	    String id = ( request.getParameter("ID")==null ) ? "" : Utility.encode((String)request.getParameter("ID"));			
		//102.12.18 add ==============================================================================================================
	    String ui_ID = 	( request.getParameter("ID")==null ) ? "" : (String)request.getParameter("ID");//UI上所key的ID	
	    String encode_ID = ( request.getParameter("encode_ID")==null ) ? "" : (String)request.getParameter("encode_ID");//DB取得的ID	
	    
		System.out.println("UI.id="+(String)request.getParameter("ID"));
		System.out.println("decode.id="+Utility.decode(encode_ID));
		System.out.println("id.src="+id);
		if(ui_ID.indexOf("****") != -1){ //UI上的已經是mask過的ID,且這次無變更ID資料
			id = encode_ID;
		}	
		System.out.println("id.after="+id);
		//============================================================================================================================
		
		try {
				//List updateDBSqlList = new LinkedList();
				sqlCmd.append(" SELECT count(*) as have_cnt FROM WLX02_M "
				       + " WHERE bank_no=?"
				       + " AND position_code=?"
				       + " AND seq_no <> ?"
					   + " AND (abdicate_code <> 'Y' or abdicate_code is null)");
				paramList.add(bank_no) ;
				paramList.add(position_code);
				paramList.add(seq_no);
				
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"have_cnt");		 			    
				System.out.println("WLX01_M.size="+data.size());
				
				if( Integer.parseInt((((DataObject)data.get(0)).getValue("have_cnt")).toString()) > 0 ){
				    errMsg = errMsg + "該職務已建檔<br>";				    				
			    }else{    
			    	//95.06.05 add 寫入WLX02_M_LOG ===================================================================
			    	sqlCmd.setLength(0) ;
			    	paramList.clear() ;
                    sqlCmd.append( " INSERT INTO WLX02_M_LOG "
                           + " select BANK_NO,POSITION_CODE,ID,NAME,BIRTH_DATE,DEGREE,SEX,TELNO,INDUCT_DATE,BACKGROUND,"
                           + "        CHOOSE_ITEM,RANK,ABDICATE_CODE,ABDICATE_DATE,EMAIL,USER_ID,USER_NAME,UPDATE_DATE,"
                           + "	      SEQ_NO,ID_CODE,?,?,sysdate,'U'"
						   + " from WLX02_M"
						   + " where bank_no=? and seq_no=?");
                    paramList.add(user_id) ;
                    paramList.add(user_name) ;
                    paramList.add(bank_no);
                    paramList.add(seq_no);
                    this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList) ;
					//updateDBSqlList.add(sqlCmd); 		            	
					//==================================================================================================
				 
			    	sqlCmd.setLength(0) ;
					paramList.clear() ;
				    sqlCmd.append("UPDATE WLX02_M SET ");
				    sqlCmd.append(" position_code=?");paramList.add(position_code);
				    sqlCmd.append(",id=?");paramList.add(id);
				    sqlCmd.append(",id_code=?");paramList.add(id_code) ;
				   	sqlCmd.append(",name=?");paramList.add(name);				    	   
					sqlCmd.append(",birth_date=to_date(?,'YYYY/MM/DD')");paramList.add(birth_date) ;											   					    
					sqlCmd.append(",degree=?"); paramList.add( degree);  
					sqlCmd.append(",sex=?"); paramList.add( sex); 
					sqlCmd.append(",telno=?"); paramList.add( telno);
					sqlCmd.append(",induct_date=to_date(?,'YYYY/MM/DD')");paramList.add(induct_date) ; 
					sqlCmd.append(",background=?"); paramList.add( background) ;
				    sqlCmd.append(",choose_item=?"); paramList.add( choose_item);
					sqlCmd.append(",rank=?");paramList.add(rank); 					       
					sqlCmd.append(",abdicate_code=?"); paramList.add( abdicate_code ) ; 
					sqlCmd.append(",abdicate_date=to_date(?,'YYYY/MM/DD')");paramList.add(abdicate_date) ;  					       
					sqlCmd.append(",email=?"); paramList.add( email);
					sqlCmd.append(",user_id=?"); paramList.add( user_id) ;
					sqlCmd.append(",user_name=?");paramList.add( user_name);
					sqlCmd.append(",update_date=sysdate"); 	
					sqlCmd.append(" where bank_no=?");paramList.add(bank_no); 
					sqlCmd.append(" and seq_no=?");paramList.add(seq_no);				            		 						       						   
						   
		            //updateDBSqlList.add(sqlCmd); 		            	
	            		
					if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){					 
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
    //94.04.01 主key更改為bank_no+seq_no 
    private String DeleteWLX02_M(HttpServletRequest request,String bank_no,String seq_no,String lguser_id,String lguser_name) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer();		
		String errMsg="";		
		String user_id=lguser_id;
		String user_name=lguser_name;
		
		try {
				//List updateDBSqlList = new LinkedList();
				List paramList = new ArrayList() ;
				sqlCmd.append("SELECT * FROM WLX02_M WHERE bank_no=?  AND seq_no= ? ");					 
				paramList.add(bank_no) ;
				paramList.add(seq_no) ;
				
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"rank,birth_date,induct_date,abdicate_date");		 			    
				System.out.println("WLX02_M.size="+data.size());
				
				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法刪除<br>";
				}else{    
				    //95.06.05 add 寫入WLX02_M_LOG ===================================================================
				    paramList.clear() ;
				    sqlCmd.setLength(0) ;
                    sqlCmd.append(" INSERT INTO WLX02_M_LOG "
                           + " select BANK_NO,POSITION_CODE,ID,NAME,BIRTH_DATE,DEGREE,SEX,TELNO,INDUCT_DATE,BACKGROUND,"
                           + "        CHOOSE_ITEM,RANK,ABDICATE_CODE,ABDICATE_DATE,EMAIL,USER_ID,USER_NAME,UPDATE_DATE,"
                           + "	      SEQ_NO,ID_CODE,"
						   + "        ?,?,sysdate,'D'"
						   + " from WLX02_M"
						   + " where bank_no=? and seq_no=?");
                    paramList.add(user_id) ;
                    paramList.add(user_name) ;
                    paramList.add(bank_no) ;
                    paramList.add(seq_no);
                    this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList) ;
					//updateDBSqlList.add(sqlCmd); 		            	
					//==================================================================================================
				 	sqlCmd.setLength(0) ;
					paramList.clear() ;
				    sqlCmd.append(" delete WLX02_M where bank_no=? and seq_no= ?");
				    paramList.add(bank_no) ;
				    paramList.add(seq_no);
		            //updateDBSqlList.add(sqlCmd); 		            		            		
					if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){					 
						errMsg = errMsg + "Y";					
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
	
	private String AbdicateWLX02_M(HttpServletRequest request,String bank_no,String seq_no,String lguser_id,String lguser_name) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer();		
		String errMsg="";		
		String abdicate_code=(String)request.getParameter("ABDICATE_CODE");
		String abdicate_date=(String)request.getParameter("ABDICATE_DATE");
		String user_id=lguser_id;
		String user_name=lguser_name;
		String position_code = ( request.getParameter("POSITION_CODE")==null ) ? "" : (String)request.getParameter("POSITION_CODE");			
	    String id = ( request.getParameter("encode_ID")==null ) ? "" : (String)request.getParameter("encode_ID");			
	   
		try {
				//List updateDBSqlList = new LinkedList();
				List paramList =new ArrayList() ;
				sqlCmd.append("SELECT * FROM WLX02_M WHERE bank_no=? AND position_code=? and Id=?"); 
				paramList.add(bank_no) ;
				paramList.add(position_code) ;
				paramList.add(id);
			    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"rank,birth_date,induct_date,abdicate_date");		 			    
				System.out.println("WLX02_M.size="+data.size());
				
				if (data.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法卸任<br>";
				}else{    
					//95.06.05 add 寫入WLX02_M_LOG ===================================================================
					sqlCmd.setLength(0) ;
					paramList.clear() ;
                    sqlCmd.append(" INSERT INTO WLX02_M_LOG "
                           + " select BANK_NO,POSITION_CODE,ID,NAME,BIRTH_DATE,DEGREE,SEX,TELNO,INDUCT_DATE,BACKGROUND,"
                           + "        CHOOSE_ITEM,RANK,ABDICATE_CODE,ABDICATE_DATE,EMAIL,USER_ID,USER_NAME,UPDATE_DATE,"
                           + "	      SEQ_NO,ID_CODE,"
                           + "        ?,?,sysdate,'A'"
						   + " from WLX02_M"
						   + " where bank_no=? and position_code=? and id=?");
                    paramList.add(user_id) ;
                    paramList.add(user_name) ;
                    paramList.add(bank_no) ;
                    paramList.add(position_code) ;
                    paramList.add(id) ;
					//updateDBSqlList.add(sqlCmd);
					this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList) ;
					//==================================================================================================
				    sqlCmd.setLength(0) ;
			    	paramList.clear() ;
				    sqlCmd.append(" UPDATE WLX02_M SET abdicate_code=?,abdicate_date=to_date(?,'YYYY/MM/DD')"			
				    	   + " where bank_no=? and position_code=? and id=?");
				    paramList.add(abdicate_code) ;
				    paramList.add(abdicate_date);
				    paramList.add(bank_no);
				    paramList.add(position_code) ;
				    paramList.add(id) ;
		            //updateDBSqlList.add(sqlCmd);
		            
					if(this.updDbUsesPreparedStatement(sqlCmd.toString(),paramList)){					 
						errMsg = errMsg + "Y";					
					}else{
				   		errMsg = errMsg + "執行卸任失敗";
					}
    	   		}
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "執行卸任失敗";							
		}	

		return errMsg;
	} 
	public String InsertWlXOPERATE_LOG(HttpServletRequest request,String lguser_id,String program_id,String pbank_no,String bank_no,String update_type) throws Exception{    	
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList() ;
		String errMsg="";
		String position_code=( request.getParameter("POSITION_CODE")==null ) ? "" : (String)request.getParameter("POSITION_CODE");
		String seq_no=( request.getParameter("seq_no")==null ) ? "" : (String)request.getParameter("seq_no");
	    String upd_name=( request.getParameter("NAME")==null ) ? "" : (String)request.getParameter("NAME");
	    try {
	        sqlCmd.append("select name from WLX02_M WHERE bank_no=? and seq_no=?  ");
			paramList.add(bank_no) ;
			paramList.add(seq_no) ;
		    List data = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"name");		 			    
			System.out.println("WLX02_M.size()="+data.size());
			if(data.size() > 0 && "".equals(upd_name)){
			    upd_name=Utility.getTrimString(((DataObject)data.get(0)).getValue("name"));
			}
			sqlCmd.setLength(0) ;
			paramList.clear() ;
	        sqlCmd.append(" INSERT INTO WlXOPERATE_LOG(muser_id,use_Date,program_id,ip_address,pbank_no,bank_no,position_code,upd_name,update_type)");
	        sqlCmd.append("                     VALUES(?,sysdate,?,?,?,?,?,?,?) ");
	        paramList.add(lguser_id);
	        paramList.add(program_id);
	        paramList.add(request.getRemoteAddr());//ipAddress
	        paramList.add(pbank_no);//總機構代號
	        paramList.add(bank_no);//分支機構代號
	        paramList.add(position_code);//異動職位(高階主管/負責人/理監事) 
	        paramList.add(upd_name);//異動姓名(高階主管/負責人/理監事) 
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