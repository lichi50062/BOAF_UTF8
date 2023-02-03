<%
//95.01.03 create by 2295
//95.05.24 fix C02->支票存款.C03->統一農漁貸 by 2295
//99.12.11 fix sqlInjection by 2808
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
<%@ page import="java.util.Enumeration" %>
<%@ page import="java.util.HashMap" %>
<%@ page import="java.util.Hashtable" %>
<%@ page import="java.util.StringTokenizer" %>

<%
	RequestDispatcher rd = null;
	String actMsg = "";	
	String alertMsg = "";			
	String webURL_Y = "";	
	String webURL_N = "";		
	boolean doProcess = false;	
	
	//取得session資料,取得成功時,才繼續往下執行===================================================
	if(session.getAttribute("muser_id") == null){//session timeout	
      System.out.println("ZZ032W login timeout");   
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
	
	
	System.out.println("act="+act);	
   
	//登入者資訊
	String lguser_id = ( session.getAttribute("muser_id")==null ) ? "" : (String)session.getAttribute("muser_id");		
	String lguser_name = ( session.getAttribute("muser_name")==null ) ? "" : (String)session.getAttribute("muser_name");		
	String lguser_type = ( session.getAttribute("muser_type")==null ) ? "" : (String)session.getAttribute("muser_type");				
	String lguser_tbank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");			
	session.setAttribute("nowtbank_no",null);//94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除======
	
	String bank_type = ( session.getAttribute("bank_type")==null ) ? "" : (String)session.getAttribute("bank_type");				
	String trans_type = ( request.getParameter("TRANS_TYPE")==null ) ? "" : (String)request.getParameter("TRANS_TYPE");
	String report_no = "";				
    String bank_code = ( request.getParameter("TBANK_NO")==null ) ? "" : (String)request.getParameter("TBANK_NO");				
    String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? "" : (String)request.getParameter("S_YEAR");				
    String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? "" : (String)request.getParameter("S_MONTH");				
    String upd_code = ( request.getParameter("UPD_CODE")==null ) ? "" : (String)request.getParameter("UPD_CODE");				
    String lock_status = ( request.getParameter("LOCK_STATUS")==null ) ? "" : (String)request.getParameter("LOCK_STATUS");				
    String firstStatus = ( request.getParameter("firstStatus")==null ) ? "" : (String)request.getParameter("firstStatus");			    	    
    String tmpbank_type="";
    System.out.println("ZZ032W.trans_type="+trans_type);	
    System.out.println("ZZ032W.bank_type="+bank_type);	
    System.out.println("ZZ032W.bank_code="+bank_code);	
    System.out.println("ZZ032W.S_YEAR="+S_YEAR);	
    System.out.println("ZZ032W.S_MONTH="+S_MONTH);	
    System.out.println("ZZ032W.upd_code="+upd_code);	
    System.out.println("ZZ032W.lock_status="+lock_status);	
    System.out.println("ZZ032W.tmpbank_type="+tmpbank_type);	
    if(!trans_type.equals("")){
       report_no = trans_type.substring(0,trans_type.indexOf(":"));    	    
    }   
    System.out.println("report_no="+report_no);
    if(lguser_id.equals("A111111111") || bank_type.equals("2") || bank_type.equals("1")){
       tmpbank_type = "Z";			    
	}else{
	   tmpbank_type = bank_type;				
	}	  
	
    if(!CheckPermission(request)){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );        
    }else{            
    	//set next jsp 	    	
    	if( act.equals("List")){
    	    List queryList = new ArrayList();
    	    queryList.add(report_no);
    	   
    	    List dbData = DBManager.QueryDB_SQLParam("select m_year,m_month from wlx_apply_ini where report_no=?",queryList,"m_year,m_month");
    	    if(dbData != null && dbData.size() > 0){
    	       System.out.println("wlx_apply_ini.size()="+dbData.size());
    	       if((Integer.parseInt(S_YEAR) < Integer.parseInt((((DataObject)dbData.get(0)).getValue("m_year")).toString()) )
    	       || (Integer.parseInt(S_MONTH) < Integer.parseInt((((DataObject)dbData.get(0)).getValue("m_month")).toString()) )){
    	          alertMsg = "本項目起始申報年月(季)為("+S_YEAR+","+S_MONTH+")";    	    
    	          webURL_Y = ListPgName +"?act=Qry";    
    	          System.out.println("本項目起始申報年月(季)為("+S_YEAR+","+S_MONTH+")");
    	       }else{
    	          List lockList = getQryResult(bank_type,bank_code, report_no,S_YEAR,S_MONTH,upd_code,lock_status,lguser_tbank_no,lguser_id);    	       	    
    	          request.setAttribute("lockList",lockList);    	     	        	    
    	       }
    	    }else{    	        	     
    	      List lockList = getQryResult(bank_type,bank_code, report_no,S_YEAR,S_MONTH,upd_code,lock_status,lguser_tbank_no,lguser_id);    	       	    
    	      request.setAttribute("lockList",lockList);    	     	        	    
    	    }
        	//rd = application.getRequestDispatcher( ListPgName +"?act=List");                	
        	rd = application.getRequestDispatcher( ListPgName +"?act=List&firstStatus="+firstStatus+"&TRANS_TYPE="+trans_type+"&bank_code="+bank_code+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&upd_code="+upd_code+"&lock_status="+lock_status+"&tmpbank_type="+tmpbank_type);            	        	        	    	            
        }else if(act.equals("Qry") || act.equals("goQry")){                    	        	        	    
            if(!S_YEAR.equals("")){
               S_YEAR = String.valueOf(Integer.parseInt(S_YEAR));
            } 
            if(!S_MONTH.equals("")){
                S_MONTH = String.valueOf(Integer.parseInt(S_MONTH));
            }      	    
    	    rd = application.getRequestDispatcher( ListPgName +"?act=Qry&firstStatus="+firstStatus+"&TRANS_TYPE="+trans_type+"&bank_code="+bank_code+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&upd_code="+upd_code+"&lock_status="+lock_status);            	        	        	    	    
    	}else if(act.equals("Lock")){   
    	    System.out.println("Lock.tmpbank_type="+tmpbank_type); 	
    	    actMsg = UpdateDB(request,lguser_id,lguser_name,report_no,"Y",tmpbank_type);     	        	
    	    if(actMsg.indexOf("相關資料寫入資料庫成功") != -1){
		       alertMsg = "本作業執行完成";    	           	       	
    	       webURL_Y = ListPgName +"?act=Qry&firstStatus="+firstStatus+"&TRANS_TYPE="+trans_type+"&bank_code="+bank_code+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&upd_code="+upd_code+"&lock_status="+lock_status; 	   	   	
    	    }       
    	    rd = application.getRequestDispatcher( nextPgName );    
        	//rd = application.getRequestDispatcher( nextPgName +"?goPages=ZZ032W.jsp&act=Qry&firstStatus="+firstStatus+"&TRANS_TYPE="+trans_type+"&bank_code="+bank_code+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&upd_code="+upd_code+"&lock_status="+lock_status+"&tmpbank_type="+tmpbank_type);            	        	        	    	    
        }else if(act.equals("unLock")){    	
            System.out.println("Lock.tmpbank_type="+tmpbank_type); 	
    	    actMsg = UpdateDB(request,lguser_id,lguser_name,report_no,"N",tmpbank_type);         	
    	    if(actMsg.indexOf("相關資料寫入資料庫成功") != -1){
		       alertMsg = "本作業執行完成";    	           	       	
    	       webURL_Y = ListPgName +"?act=Qry&firstStatus="+firstStatus+"&TRANS_TYPE="+trans_type+"&bank_code="+bank_code+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&upd_code="+upd_code+"&lock_status="+lock_status; 	   	   	
    	    }
    	    rd = application.getRequestDispatcher( nextPgName );    
        	//rd = application.getRequestDispatcher( nextPgName +"?goPages=ZZ032W.jsp&act=Qry&firstStatus="+firstStatus+"&TRANS_TYPE="+trans_type+"&bank_code="+bank_code+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&upd_code="+upd_code+"&lock_status="+lock_status+"&tmpbank_type="+tmpbank_type);            	        	        	    	    
        }
       
    	request.setAttribute("actMsg",actMsg);
    	request.setAttribute("alertMsg",alertMsg);
    	request.setAttribute("webURL_Y",webURL_Y);    
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
    private final static String ListPgName = "/pages/ZZ032W_List.jsp";        
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    private boolean CheckPermission(HttpServletRequest request){//檢核權限    	    
    	    boolean CheckOK=false;
    	    HttpSession session = request.getSession();            
            Properties permission = ( session.getAttribute("ZZ032W")==null ) ? new Properties() : (Properties)session.getAttribute("ZZ032W");				                
            if(permission == null){
              System.out.println("ZZ032W.permission == null");
            }else{
               System.out.println("ZZ032W.permission.size ="+permission.size());
               
            }
            //只要有Query的權限,就可以進入畫面
        	if(permission != null && permission.get("Q") != null && permission.get("Q").equals("Y")){            
        	   CheckOK = true;//Query
        	}
        	return CheckOK;
    }           
     
    //取得查詢結果
    private List getQryResult(String bank_type,String bank_code,String report_no,String S_YEAR,String S_MONTH,String upd_code,String lock_status,String tbank_no,String lguser_id){    	   
    		//查詢條件        		
    		StringBuffer sqlCmd = new StringBuffer() ;
    		String rule_1 = "";
    		String rule_2 = "";
    		String rule_3 = "";
    		String yy = Integer.parseInt(S_YEAR) > 99 ?"100" :"99" ;
    		String cd01Table = Integer.parseInt(S_YEAR) > 99 ?"cd01" :"cd01_99" ;
    		String tmpbank_type="";
    		String tmpbank_no="";
    		List <String > paramList = new ArrayList<String>() ;
    		List <String > rule_1List = new ArrayList<String>() ;
    		List <String > rule_2List = new ArrayList<String>() ;
    		List <String > rule_3List = new ArrayList<String>() ;
    		if(lguser_id.equals("A111111111") || bank_type.equals("2") || bank_type.equals("1")){
    		    tmpbank_type = "Z";
			    tmpbank_no = "9999999";
		    }else{
		        tmpbank_type = bank_type;
				tmpbank_no = tbank_no;
			}	

            rule_1 = " select * from ("
				   + " select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name,"
    			   + "		  BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type , "
    			   + "		  nvl(wlx01.CENTER_NO,' ') as CENTER_NO,"
				   + "		  nvl(wlx01.M2_NAME,' ')   as M2_NAME,"
				   + "		  wlx01.CANCEL_NO,   wlx01.CANCEL_date, "
    			   + "	      nvl(cd01.hsien_id,' ') as  hsien_id,"
				   + "		  nvl(wml01.lock_own,'N') as lock_own,"
				   + "		  nvl(wml01.lock_mgr,'N') as lock_mgr,"
				   + "		  nvl(wlx_temp.Tot_Cnt,0) as Tot_Cnt,"
				   + "		  nvl(wlx_temp.Agree_Cnt,0) as Agree_Cnt,"
				   + "		  nvl(wlx_temp.Have_MK,' ') as Have_MK,  "
				   + "		  nvl(cd01.hsien_name,'OTHER')  as  hsien_name,"
    			   + "		  cd01.FR001W_output_order     as  FR001W_output_order,"
				   + "		  nvl(((to_char(WML01.UPDATE_DATE,'yyyymmdd')-19110000)  || to_char(WML01.UPDATE_DATE,' hh24:mi')),' ')  as S_UpdateDate "
			       + "	from  (select * from "+cd01Table+"  where  hsien_id <> 'Y') cd01 "
     			   + "		   left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id "
	 			   + "		   left join (select BANK_NO, BANK_NAME,  bank_type "
	               + "					  from bn01 where bn01.bank_type in('6','7') and m_year=?)  bn01 "
			       + "	on wlx01.bank_no=bn01.bank_no ";
            rule_1List.add(yy);
            rule_1List.add(yy);
			if(report_no.equals("C04") || report_no.equals("C05")){
    		   rule_2 = " left join  (select  * FROM  wlx_apply_lock WML01 "
    		   		  + " 			  where WML01.M_Year = ?"
    		   		  +	"			    and WML01.M_quarter = ?"                    		   		  
    		   		  + " 				and WML01.Report_No = ?)  WML01 "
					  + " 			  on 	bn01.bank_no = WML01.Bank_Code "
	 			      + " left join  (select wlx_temp.bank_no  as  bank_no, Tot_Cnt,  "
       				  +	"			         Agree_Cnt,  nvl(Have_Cnt,' ') as Have_MK"
 				      +	"			  from "					 
				      + "				   (select m_year, m_quarter, bank_no,  count(*) Tot_Cnt,  "
      				  + "				    	   sum(decode(AuditResult_YN,'Y',1,0)) as Agree_Cnt "      				  
  				      +	"					FROM  wlx08_s_gage  WML01 "
  				      + "				    where WML01.M_Year = ?"
  				      + "					  and WML01.M_quarter =? "
   				      + "					group by m_year, m_quarter, bank_no "
   					  + "				    order by m_year, m_quarter, bank_no) wlx_temp "
  					  + " left join  (select bank_no, '*' as Have_Cnt from wlx08_s_gage_apply  WML01 "
                 	  + "			  where  WML01.M_Year = ?"
                 	  + "			    and  WML01.M_quarter = ? )  wlx_temp_1 "
             		  +	"			  on wlx_temp.bank_no = wlx_temp_1.bank_no)  wlx_temp "
             		  + "  on bn01.bank_no = wlx_temp.bank_no ";
    		   rule_2List.add(S_YEAR) ;
    		   rule_2List.add(String.valueOf(Integer.parseInt(S_MONTH)));
    		   rule_2List.add(report_no) ;
    		   rule_2List.add(S_YEAR) ;
    		   rule_2List.add(String.valueOf(Integer.parseInt(S_MONTH)));
    		   rule_2List.add(S_YEAR) ;
    		   rule_2List.add(String.valueOf(Integer.parseInt(S_MONTH)));
    		}else{       
		       rule_2 = " left join  (select  * FROM  wlx_apply_lock WML01 "
		    	 	  + " 			   where WML01.M_Year = ?"
		    	      + "   		     and WML01.M_quarter = ?"
		    	      + "   			 and WML01.Report_No = ?) WML01 "
				      + "    		   on 	bn01.bank_no = WML01.Bank_Code "
	 			      + " left join  (select  bank_no, 1 as Tot_Cnt, " 
                      + "		      		   1 as Agree_Cnt,  '*'  as Have_MK "
		              + "			   from ";
	 		   rule_2List.add(S_YEAR) ;
	 		   rule_2List.add(String.valueOf(Integer.parseInt(S_MONTH))) ;
	 		   rule_2List.add(report_no) ;
		       if(report_no.equals("C01")){
		          rule_2 += " wlx05_m_atm WML01 ";
		       }else if(report_no.equals("C02")){//95.05.24支票存款資料  
		          rule_2 += " WLX07_M_CHECKBANK WML01 ";
		       }else if(report_no.equals("C03")){//95.05.24農一農漁貸資料  
		          rule_2 += " WLX07_M_CREDIT WML01 ";   
		       }else if(report_no.equals("C06")){  
		          rule_2 += " WLX09_S_WARNING WML01 ";   
		       }else if(report_no.equals("C07")){  
		          rule_2 += " WLX_S_RATE  WML01 ";      
		       }		     
               rule_2 += " where  WML01.M_Year = ?";
               rule_2List.add(S_YEAR) ;
               if(report_no.equals("C01") || report_no.equals("C02") || report_no.equals("C03")){
                  rule_2 += "  and  WML01.M_Month = ? )  wlx_temp ";
		          rule_2List.add(String.valueOf(Integer.parseInt(S_MONTH))) ;
               }else if(report_no.equals("C06") || report_no.equals("C07")){
                 rule_2 += "  and  WML01.M_Quarter = ? )  wlx_temp ";
                 rule_2List.add(String.valueOf(Integer.parseInt(S_MONTH))) ;
               }
               rule_2 += " on bn01.bank_no = wlx_temp.bank_no ";
		    }
    		
    		rule_3 = " ) Temp_Output "
    			   + " where 	bank_no is not null  ";
    	    if(!bank_code.equals("")){
    	       rule_3 += " and bank_no =? ";
    	       rule_3List.add(bank_code) ;
    	    }		   
    			   //+ " 	 and (('"+bank_code+"' = '')  or ( '"+bank_code+"' <>  ' '  and  '"+bank_code+"' = BANK_NO))"
    		rule_3 += "   and ( "
				   + "			(? = 'Z') or "
				   + "		    ((? = '6' or ? = '7') and ? = BANK_NO) or "
				   + "		    (? = '8' and ? = CENTER_NO) or "
			       + "			(? = 'B' and ? = M2_NAME) "
				   + " 		 )"
				   + "   and ( "
				   + "		 (? = 'ALL')  or "//全部 
				   + "		 (? = '0' and  Have_MK = '*') or "//已申報
				   + "		 (? = '1' and  Have_MK <> '*') or "//未申報
				   + " 		 (? = '2' and  Tot_Cnt > 0 and (Tot_Cnt <> Agree_Cnt)) "//未審核
				   + "		 )"
				   + "   and ("
				   + "		 (? = 'ALL')  or "//全部
				   + "		 (? = 'Y' and  (lock_own = 'Y' or lock_mgr= 'Y')) or "//鎖定
				   + " 		 (? = 'N' and  (lock_own <> 'Y' and lock_mgr <> 'Y')) "//未鎖定
				   + "		 )"                                                                                             
				   + " order by  Bank_Type, FR001W_output_order, BANK_NO";
			rule_3List.add(tmpbank_type) ;
			rule_3List.add(tmpbank_type) ;
			rule_3List.add(tmpbank_type) ;
			rule_3List.add(tmpbank_no) ;
			rule_3List.add(tmpbank_type) ;
			rule_3List.add(tmpbank_no) ;
			rule_3List.add(tmpbank_type) ;
			rule_3List.add(tmpbank_no) ;
			rule_3List.add(upd_code) ;
			rule_3List.add(upd_code) ;
			rule_3List.add(upd_code) ;
			rule_3List.add(upd_code) ;
			rule_3List.add(lock_status) ;
			rule_3List.add(lock_status) ;
			rule_3List.add(lock_status) ;
    		//sqlCmd = rule_1 + rule_2 + rule_3;
    		sqlCmd.append(rule_1).append(rule_2).append(rule_3) ;
    		for(String s : rule_1List) {
    			paramList.add(s) ;
    		}
    		for(String s : rule_2List) {
    			paramList.add(s) ;
    		}
    		for(String s : rule_3List) {
    			paramList.add(s) ;
    		}
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"cancel_date,tot_cnt,agree_cnt");            
            return dbData;
    }
     
     
    
	public String UpdateDB(HttpServletRequest request,String lguser_id,String lguser_name,String report_no,String lock_status,String tmpbank_type) throws Exception{    	
		String sqlCmd = "";		
		String errMsg="";
		String user_id=lguser_id;
	    String user_name=lguser_name;		
	    
	    String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? "" : (String)request.getParameter("S_YEAR");				
        String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? "" : (String)request.getParameter("S_MONTH");				
	    String bank_code="";	    
	    
		
	    List paramList =new ArrayList() ;
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
		  	    List lockData = new LinkedList();
		  		for ( int i = 0; i < row; i++) {		  	    		  	  			  
					if ( t.get("isModify_" + (i+1)) != null ) {					  
					 lockData.add((String)t.get("isModify_"+(i+1)));
					}										
		  		}	
		  		System.out.println("lockData.size="+lockData.size());
		  		
			    List updateDBSqlList = new LinkedList();
			    List data = null;
			    StringTokenizer st = null;

			    for(int i=0;i<lockData.size();i++){					       			        
         			bank_code = (String)lockData.get(i);     			    
     			    System.out.println("S_YEAR = '"+S_YEAR+"'");
     			    System.out.println("S_MONTH = '"+S_MONTH+"'");
     			    System.out.println("bank_code = '"+bank_code+"'");
     			    sqlCmd = "select * from wlx_apply_lock WHERE m_year=? AND M_QUARTER=?" +
							 " AND bank_code=?  AND report_no=?";
     			    paramList.add(S_YEAR) ;
     			    paramList.add(S_MONTH) ;
     			    paramList.add(bank_code) ;
     			    paramList.add(report_no) ;
					data = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");		
					paramList.clear() ;
				    System.out.println("update.size="+data.size());				    
					if(data.size() != 0){//有資料							 
			        	sqlCmd = "INSERT INTO wlx_apply_lock_log " +					   		
					   			 " select m_year,m_quarter,bank_code,report_no,add_user,add_name,add_date,lock_own,lock_mgr,user_id,user_name,update_date"+					   			 
					   			 ",?,?,sysdate,'U'"+
					   		 	 " from wlx_apply_lock WHERE m_year=? AND M_QUARTER=?" +
							     " AND bank_code=? AND report_no=?";
					    paramList.add(lguser_id) ;
					    paramList.add(lguser_name) ;
					    paramList.add(S_YEAR) ;
					    paramList.add(S_MONTH) ;
					    paramList.add(bank_code) ;
					    paramList.add(report_no) ;
						//updateDBSqlList.add(sqlCmd);						
    		    		this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
    		    		paramList.clear() ;
						if(tmpbank_type.equals("Z")){//A111111111|農金局|全國農業金庫
						   sqlCmd = "UPDATE wlx_apply_lock set lock_mgr=? ";		
						   paramList.add(lock_status) ;
						}else{
						   sqlCmd = "UPDATE wlx_apply_lock set lock_own=? ";
						   paramList.add(lock_status) ;
						}   
						sqlCmd += " ,user_id=?,user_name=?,update_date=sysdate"
						        + " WHERE m_year=? AND M_QUARTER= ?" 
						        + " AND bank_code=? AND report_no=? ";
						paramList.add(lguser_id) ;
						paramList.add(lguser_name);
						paramList.add(S_YEAR) ;
						paramList.add(S_MONTH);
						paramList.add(bank_code) ;
						paramList.add(report_no) ;
					}else if(lock_status.equals("Y")){					    
					    sqlCmd = "Insert into wlx_apply_lock VALUES(?,?,?,?,?,?,sysdate,";
					    paramList.add(S_YEAR) ;
					    paramList.add(S_MONTH) ;
					    paramList.add(bank_code);
					    paramList.add(report_no) ;
					    paramList.add(lguser_id) ;
					    paramList.add(lguser_name);
						if(tmpbank_type.equals("Z")){//A111111111|農金局|全國農業金庫
						   sqlCmd += "'N','Y'";
						}else{
						   sqlCmd += "'Y','N'";
						}	   
						sqlCmd += ",?,?,sysdate)";
					    paramList.add(lguser_id) ;
					    paramList.add(lguser_name) ;
					}	
		            updateDBSqlList.add(sqlCmd);						   						      				    
	            }//end of for		
	            	
			    //if(DBManager.updateDB(updateDBSqlList)){
			    if(this.updDbUsesPreparedStatement(sqlCmd,paramList)) {
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