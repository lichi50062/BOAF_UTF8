<%
// 93.12.28 create by 2295
// 94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除 by 2295
// 94.01.14 加上若沒有trans_type, 則根據bank_type來判斷trans_type屬於何者
// 94.02.14 fix A111111111可以看到全部 by 2295
// 		   fix 農漁會共用中心..可以看到其所屬的機構 by 2295
// 94.02.21 fix A111111111跟農金局可鎖定全部機構 by 2295
// 94.03.07 fix 若為共用中心鎖定時,lock_stauts="C"後,再設成"Y" by 2295
// 94.05.06 fix 若為A111111111跟農金局時,根據trans_type來鎖定全部機構 by 2295
// 94.06.08 fix 只抓總機構 by 2295
// 94.09.09 add 全國農業金庫比照農金局查詢A01-A05 by 2295
// 95.02.09 add 全國農業金庫比照農金局,進行鎖定/解除鎖定 by 2295
//          add 檢核結果增加"未申報"."申報內容均為0"."A01_申報逾放金額為0" by 2295
// 95.02.20 add 檢核結果增加"A04_申報逾放金額為0" by 2295
// 95.02.21 add 報表為A01時,檢核結果才"A01_逾放金額為0" by 2295
//          add 報表為A04時,檢核結果才"A04_逾放金額為0" 
// 95.11.07 add 增加使用者姓名.電話.機構電話 by 2295
// 99.12.10 fix sqlInjection by 2808
//100.01.11 fix bug by 2295
//102.10.04 fix bug by 2295
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
	String webURL = "";	
	boolean doProcess = false;	
	
	//取得session資料,取得成功時,才繼續往下執行===================================================
	if(session.getAttribute("muser_id") == null){//session timeout	
      System.out.println("ZZ031W login timeout");   
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
	// add by 2354 2005.01.14 begin
	if(trans_type.equals("")){	//沒有trans_type, 則根據bank_type來判斷trans_type應屬於何者
		if(bank_type.equals("6") || bank_type.equals("7") || bank_type.equals("8")){//農會.漁會.農漁會共用中心
		   trans_type="1";
		}else if(bank_type.equals("4")){//農業信用保証基金
		   trans_type="2";
		}
	}
	System.out.println("trans_type="+trans_type);
	// add by 2354 2005.01.14 end
	String report_no = ( request.getParameter("REPORT_NO")==null ) ? "" : (String)request.getParameter("REPORT_NO");				
    String bank_code = ( request.getParameter("TBANK_NO")==null ) ? "" : (String)request.getParameter("TBANK_NO");				
    String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? "" : (String)request.getParameter("S_YEAR");				
    String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? "" : (String)request.getParameter("S_MONTH");				
    String upd_code = ( request.getParameter("UPD_CODE")==null ) ? "" : (String)request.getParameter("UPD_CODE");				
    String lock_status = ( request.getParameter("LOCK_STATUS")==null ) ? "" : (String)request.getParameter("LOCK_STATUS");				
    				
    if(!Utility.CheckPermission(request,"ZZ031W")){//無權限時,導向到LoginError.jsp
        rd = application.getRequestDispatcher( LoginErrorPgName );        
    }else{            
    	//set next jsp 	    	
    	if(act.equals("getData") || act.equals("List")){    	
    	    trans_type = (trans_type.equals(""))?"1":trans_type;
    	    List Report_List = getReport_List(trans_type);    	    
    	    request.setAttribute("Report_List",Report_List);    	    
    	    report_no = ((String)((DataObject)Report_List.get(0)).getValue("cmuse_name")).substring(0,((String)((DataObject)Report_List.get(0)).getValue("cmuse_name")).indexOf("_"));
    	    System.out.println("回傳第一次查出的List.report_no="+report_no);
        	rd = application.getRequestDispatcher( ListPgName +"?act=List&report_no="+report_no+"&test=nothing");                	
        }else if(act.equals("Qry")){  
            if(!S_YEAR.equals("")){
               S_YEAR = String.valueOf(Integer.parseInt(S_YEAR));
            } 
            if(!S_MONTH.equals("")){
                S_MONTH = String.valueOf(Integer.parseInt(S_MONTH));
            }  
    	    List lockList = getQryResult(trans_type,bank_type,bank_code, report_no,S_YEAR,S_MONTH,upd_code,lock_status,lguser_tbank_no,lguser_id);    	   
    	    request.setAttribute("lockList",lockList);    	     	        	    
    	    List Report_List = getReport_List(trans_type);    	    
    	    request.setAttribute("Report_List",Report_List);    	    
    	    rd = application.getRequestDispatcher( ListPgName +"?act=Qry&TRANS_TYPE="+trans_type+"&report_no="+report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&upd_code="+upd_code+"&lock_status="+lock_status+"&bank_code="+bank_code);            	        	        	    	    
    	}else if(act.equals("Lock")){    	
    	    actMsg = UpdateDB(request,lguser_id,lguser_name,bank_type,"Y");     	
        	rd = application.getRequestDispatcher( nextPgName );         
        }else if(act.equals("unLock")){    	
    	    actMsg = UpdateDB(request,lguser_id,lguser_name,bank_type,"N"); 
        	rd = application.getRequestDispatcher( nextPgName );                 
        }else if(act.equals("printUserData")){//98.10.14 add 產生申報者通訊錄 by 2295
    	    rd = application.getRequestDispatcher( RptCreatePgName +"?act=printUserData&REPORT_NO="+report_no+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH);            	        	        	    	    
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
    private final static String EditPgName = "/pages/ZZ031W_Edit.jsp";    
    private final static String ListPgName = "/pages/ZZ031W_List.jsp";  
    private final static String RptCreatePgName = "/pages/ZZ031W_Excel.jsp";        
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
            
     
    //取得總機構代碼
    private List getReport_List(String trans_type){
           List paramList = new ArrayList();
    		//查詢條件    
    		String sqlCmd = " select * from cdshareno "
    					  + " where cmuse_div= ? order by input_order";
    		if(trans_type.equals("1")){
    		    paramList.add("012");
    		}			  
    		if(trans_type.equals("2")){
    		    paramList.add("013");
    		}
    		if(trans_type.equals("3")){
    		    paramList.add("014");
    		}
    		
    		       
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");            
            return dbData;
    }
    
    //取得查詢結果
    private List getQryResult(String trans_type,String bank_type,String bank_code,String report_no,String S_YEAR,String S_MONTH,String upd_code,String lock_status,String tbank_no,String lguser_id){    	   
    		//查詢條件        		
    		String sqlCmd = "";
    		String yy = Integer.parseInt(S_YEAR)> 99 ?"100" : "99" ;
    		List paramList = new ArrayList() ;
    		sqlCmd = " select distinct a.bank_no," // 95.02.09 add distinct by 2295
       			   + " a.bank_name,"
	   			   + " b.m_year,"
       			   + " b.m_month,"
	   			   + " b.upd_code,"
	   			   + " b.input_method,"
	   			   + " b.wml01_lock_status as lock_status,"
	   			   + " b.common_center,"
	   			   + " b.upd_method  ,"
       			   + " b.wml01_lock_lock_status as wml01_lock_status,"
       			   + " decode(nvl(b.input_method,'N'),'N','N','Y') as havingdata, "
       			   + " wlx01.telno, bn01.m_telno,bn01.muser_name";//95.11.07 add 增加使用者姓名.電話.機構電話 
       		if(upd_code.equals("A01_990000_0")){//95.02.20 add A01_逾放金額為0 by 2295	   
               sqlCmd += " from (SELECT ba01.*  from  (select * from ba01 where m_year=? )ba01, "
				      +  "      (select * from a01 where a01.m_year= ?  and "
 	                  +  " 					             a01.m_month= ? and "
                      +  "                               a01.acc_code='990000' and a01.amt= 0) a01 "
					  +  " where ba01.BANK_NO = a01.bank_code )  a ";					  
       		  paramList.add(yy) ;
       		  paramList.add(S_YEAR) ;
       		  paramList.add(S_MONTH) ;
            }else if(upd_code.equals("A04_840740_0")){//95.02.20 add A04_逾放金額為0 by 2295	   					    
			   sqlCmd += " from (SELECT ba01.*  from  (select * from ba01 where m_year=? )ba01, "
				      +  "      (select * from a04 where a04.m_year = ? and "
				   	  +  "		  			             a04.m_month= ? and "
					  +  "                               a04.acc_code='840740' and a04.amt= 0) a04 "
					  +  " where ba01.BANK_NO = a04.bank_code )  a ";
			   paramList.add(yy) ;
	       		  paramList.add(S_YEAR) ;
	       		  paramList.add(S_MONTH) ;
        	}else{		   
  			   sqlCmd += " from (select * from ba01 where m_year=?) a ";
  			   paramList.add(yy) ;
  			}	   
  			 
  			sqlCmd += " left join wml01_a_v b on a.bank_no= b.bank_code and"
			       +  "					 			    b.m_year=? and"
				   +  "					  			    b.m_month=? and"
				   +  "								    b.report_no=? "
				   +  " left join (select * from wlx01 where m_year=? )wlx01 on wlx01.bank_no=a.bank_no " //95.11.07 add 增加使用者姓名.電話.機構電話
				   +  " left join (select bn01.BANK_NO, bn01.BANK_NAME,  bn01.bank_type,wtt01.muser_name,wtt01.m_telno " 
   		   		   +  "            from (select * from bn01 where m_year=?)bn01,(select wtt01.muser_id,wtt01.muser_name,muser_data.m_telno "
   		   		   +  "						  from wtt01 left join muser_data on wtt01.muser_id = muser_data.muser_id "										       		 
				   +  "					 	 )wtt01 ";
		    paramList.add(S_YEAR) ;
		    paramList.add(S_MONTH) ;
		    paramList.add(report_no) ;
		    paramList.add(yy) ;
		    paramList.add(yy) ;
			//System.out.println("lguser_id="+lguser_id);
			//System.out.println("trans_type="+trans_type);
			//System.out.println("bank_type="+bank_type);
			//if(!lguser_id.equals("A111111111")){			  
				if(trans_type.equals("1")){//A01-A05
				   if(bank_type.equals("8") || bank_type.equals("2") /*農金局*/ || bank_type.equals("1") /*全國農業金庫*/ || lguser_id.equals("A111111111") ){
				    //共用中心.94.05.06 add農金局;94.09.09 add全國農業金庫
				      if(!bank_type.equals("8")){
				   	     sqlCmd += "   where bn01.bank_type in('6','7','8')"//95.11.07 add 增加使用者姓名.電話.機構電話
		   					    +  "   and wtt01.muser_id = bn01.bank_no || '001'"											 
		   				        +  "  ) bn01 on wlx01.bank_no=bn01.bank_no " 
		   				        +  " where a.bank_type in ('6','7','8')";
		   			  }else if(bank_type.equals("8")){//95.11.07 add 增加使用者姓名.電話.機構電話
				         sqlCmd += "   where bn01.bank_no in (select bank_no from wlx01 where center_flag='Y' and center_no = ?)"				      
		   					    +  "   and wtt01.muser_id = bn01.bank_no || '001'"											 
		   				        +  "  ) bn01 on wlx01.bank_no=bn01.bank_no " 				            
		   				        +  " where a.bank_type in ('6','7','8')"
				                +  " and a.bank_no in (select bank_no from wlx01 where center_flag='Y' and center_no = ?)";
		   				 paramList.add(tbank_no) ;
		   				 paramList.add(tbank_no) ;
				      }   
				   }
				   if(!lguser_id.equals("A111111111")){
				      if(bank_type.equals("6") || bank_type.equals("7")){//農會 or 漁會只負責其所本身的總機構	
				         sqlCmd += "   where bn01.bank_type =? "//95.11.07 add 增加使用者姓名.電話.機構電話
		   					    +  "   and wtt01.muser_id = bn01.bank_no || '001'"											 
		   				        +  "  ) bn01 on wlx01.bank_no=bn01.bank_no " 				      
				                +  " where a.bank_type = ? ";
		   				paramList.add(bank_type) ;
		   				paramList.add(bank_type) ;
				      }
				   }

		    	}else if(trans_type.equals("2")){//M01-M08//94.05.06 add農金局可看農業信用保證基金//95.02.09 add全國農業金庫
		    	   if(bank_type.equals("4") || bank_type.equals("2") /*農金局*/ || bank_type.equals("1") /*全國農業金庫*/ || lguser_id.equals("A111111111")){//農業信用保證基金
		    	   	  sqlCmd += "   where bn01.bank_type in('4')"//95.11.07 add 增加使用者姓名.電話.機構電話
		   					 +  "   and wtt01.muser_id = bn01.bank_no || '001'"											 
		   				     +  "  ) bn01 on wlx01.bank_no=bn01.bank_no " 				      				         
				             +  " where a.bank_type in ('4')";
				   }
		    	}else if(trans_type.equals("3")){//B01-B03	//95.02.09 add全國農業金庫
		    	   if(bank_type.equals("2") /*農金局*/ || bank_type.equals("1") /*全國農業金庫*/ || lguser_id.equals("A111111111")){//94.05.06 add農金局可看農金局跟農民銀行 
		    	      sqlCmd += "   where bn01.bank_type in ('2','5')"//95.11.07 add 增加使用者姓名.電話.機構電話
		   					 +  "   and wtt01.muser_id = bn01.bank_no || '001'"											 
		   				     +  "  ) bn01 on wlx01.bank_no=bn01.bank_no " 				      				         
				             +  " where a.bank_type in ('2','5')";
				   }
		    	}else if(trans_type.equals("4")){//F01		    	   
		    	      sqlCmd += "   where bn01.bank_type in ('6','7')"//95.11.07 add 增加使用者姓名.電話.機構電話
		   					 +  "   and wtt01.muser_id = bn01.bank_no || '001'"											 
		   				     +  "  ) bn01 on wlx01.bank_no=bn01.bank_no " 				      				         				
				             +  " where a.bank_type in ('6','7')";//94.11.15 add 				   
		    	}	   
		    //}//end of lguser_id != 'A111111111'
		    //94.02.21 fix A111111111跟農金局可鎖定全部機構
		    /*
		    94.05.06
		    if(lguser_id.equals("A111111111") || bank_type.equals("2")){ 
		       sqlCmd += " where a.bank_no is not null ";
		    }*/
 			if(!bank_code.equals("")){	   
				sqlCmd +=" and a.bank_no=? ";
				paramList.add(bank_code) ;
			}    
   			//System.out.println("test1="+upd_code);
   			
   			if(!upd_code.equals("ALL") && !upd_code.equals("N")){//ALL全部,N未申報   			      
   			   if(upd_code.equals("A01_990000_0") || upd_code.equals("A04_840740_0")){//95.02.09 add A01_逾放金額為0 by 2295 ////95.02.20 add A04_逾放金額為0 by 2295 
  			      sqlCmd += " and upd_code is not null ";
  			   }else{
  			      sqlCmd +=" and b.upd_code = ? ";  	
  			      paramList.add(upd_code) ;
  			   }
  			}
  			
   			if(lock_status.equals("Y")){//鎖定	   
   				sqlCmd += " and (b.wml01_lock_status='Y' or b.wml01_lock_lock_status='Y')";
   			}	   
   			if(lock_status.equals("N")){//未鎖定
   				sqlCmd += " and (b.input_method is not null and b.wml01_lock_status is null and wml01_lock_lock_status is null) ";
   			}	   
   			sqlCmd += " and a.bank_kind='0' ";//94.06.08 fix只有總機構
 			sqlCmd += " order by a.bank_no,b.m_year,b.m_month";
  			
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"m_year,m_month");            
            return dbData;
    }
    
	public String UpdateDB(HttpServletRequest request,String lguser_id,String lguser_name,String bank_type,String lock_status) throws Exception{    	
		String sqlCmd = "";		
		String errMsg="";
		String user_id=lguser_id;
	    String user_name=lguser_name;		
	    
	    String S_YEAR = ( request.getParameter("S_YEAR")==null ) ? "" : (String)request.getParameter("S_YEAR");				
        String S_MONTH = ( request.getParameter("S_MONTH")==null ) ? "" : (String)request.getParameter("S_MONTH");				

	    String bank_code="";
	    String Report_no=((String)request.getParameter("REPORT_NO")==null)?" ":(String)request.getParameter("REPORT_NO");
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
     			    paramList.clear() ;
     			    sqlCmd = "select * from wml01 WHERE m_year=? AND m_month=?"  +
							 " AND bank_code=? AND report_no=?";
     			    paramList.add(S_YEAR) ;
     			    paramList.add(S_MONTH) ;
     			    paramList.add(bank_code) ;
     			    paramList.add(Report_no) ;
					data = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");		
					paramList.clear() ;
				    System.out.println("update.size="+data.size());				    
					if(data.size() != 0){//已申報							 
			        	sqlCmd = "INSERT INTO WML01_LOG " +					   		
					   			 " select m_year,m_month,bank_code,report_no,input_method,add_user,add_name,add_date,common_center,upd_method,upd_code "+
					   			 ",batch_no,lock_status,user_id,user_name,update_date,?,?,sysdate,'U'"+
					   		 	 " FROM WML01 WHERE m_year=?  AND m_month=? " +
							 	 " AND bank_code=?  AND report_no=? ";
					    paramList.add(lguser_id) ;
					    paramList.add(lguser_name) ;
					    paramList.add(S_YEAR) ;
					    paramList.add(S_MONTH) ;
					    paramList.add(bank_code) ;
					    paramList.add(Report_no) ;
						//updateDBSqlList.add(sqlCmd);
						this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
						if(bank_type.equals("1")/*95.02.09 add 全國農業金庫*/ || bank_type.equals("2") || bank_type.equals("8")){//農金局.共用中心寫到WML01_LOCK
						   paramList.clear() ;
						   sqlCmd = " select * from WML01_LOCK WHERE m_year=?  AND m_month= ?" +
							 	  " AND bank_code=?  AND report_no=? ";   
						   paramList.add(S_YEAR) ;
						   paramList.add(S_MONTH) ;
						   paramList.add(bank_code) ;
						   paramList.add(Report_no) ;
						   data = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
						   paramList.clear() ;
				           System.out.println("updateWML01_LOCK.size="+data.size());				    	 	  
				           if(data.size() != 0){
				              if(lock_status.equals("N")){//解除鎖定,將WML01
						         sqlCmd = "DELETE WML01_LOCK WHERE m_year=? AND m_month=?"  +
							 	        " AND bank_code=? AND report_no=? ";
				                 paramList.add(S_YEAR) ;
				                 paramList.add(S_MONTH) ;
				                 paramList.add(bank_code) ;
				                 paramList.add(Report_no) ;
						      }else if(lock_status.equals("Y")){//鎖定
						         if(bank_type.equals("8")){
						            lock_status = "C";//由共用中心做鎖定的
						         }
						         sqlCmd = "UPDATE WML01_LOCK SET lock_status =? WHERE m_year=? AND m_month=?" +
							     	 " AND bank_code=? AND report_no=?";
						         paramList.add(lock_status) ;
						         paramList.add(S_YEAR) ;
						         paramList.add(S_MONTH);
						         paramList.add(bank_code) ;
						         paramList.add(Report_no) ;
						      }  						      
				           }else{				              
				              if(lock_status.equals("Y")){//鎖定				              
				                 if(bank_type.equals("8")){
						            lock_status = "C";//由共用中心做鎖定的						      
						         }						      
						         sqlCmd = "Insert into WML01_LOCK VALUES(?,?,?,?,?,?,?,sysdate)";
							     lock_status="Y";//94.03.07 fix若為共用中心鎖定時,恢復為"Y"
							     paramList.add(S_YEAR) ;
							     paramList.add(S_MONTH) ;
							     paramList.add(bank_code) ;
							     paramList.add(Report_no) ;
							     paramList.add(lock_status) ;
							     paramList.add(lguser_id) ;
							     paramList.add(lguser_name) ;
						      }						      						      
				           }
				           //updateDBSqlList.add(sqlCmd);
				           this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
						}//end of bank_type in ('2','8')
						paramList.clear() ;
						if(bank_type.equals("6") || bank_type.equals("7") || bank_type.equals("4") ){//農會.漁會.農業信用保証基金寫到WML01
						   if(lock_status.equals("N")){//解除鎖定
						      sqlCmd = "UPDATE WML01 SET lock_status = null ,user_id=?,user_name=?,update_date=sysdate WHERE m_year=? AND m_month=?"  +
							 	     " AND bank_code=? AND report_no=?";
						      paramList.add(lguser_id) ;
						      paramList.add(lguser_name) ;
						      paramList.add(S_YEAR);
						      paramList.add(S_MONTH) ;
						      paramList.add(bank_code) ;
						      paramList.add(Report_no) ;
						   }else if(lock_status.equals("Y")){//鎖定
						      sqlCmd = "UPDATE WML01 SET lock_status ='Y' ,user_id=?,user_name=?,update_date=sysdate WHERE m_year=? AND m_month=?"  +
							     	 " AND bank_code=? AND report_no=?";
						   	  paramList.add(lguser_id) ;
						   	  paramList.add(lguser_name) ;
						   	  paramList.add(S_YEAR) ;
						   	  paramList.add(S_MONTH) ;
						   	  paramList.add(bank_code) ;
						   	  paramList.add(Report_no) ;
						   }	    	  
						   //updateDBSqlList.add(sqlCmd);
						   this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
						   
						}//end of bank_type in ('6','7','4')						
					}else{//未申報
					   paramList.clear() ;
					   sqlCmd = " select * from WML01_LOCK WHERE m_year=? AND m_month=?" +
							 	  " AND bank_code=? AND report_no=? ";
					   paramList.add(S_YEAR) ;
					   paramList.add(S_MONTH) ;
					   paramList.add(bank_code);
					   paramList.add(Report_no) ;
						   data = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");	
					   paramList.clear() ;
				           System.out.println("updateWML01_LOCK.size="+data.size());				    	 	  
				           if(data.size() != 0){
				              if(lock_status.equals("N")){//解除鎖定,將WML01
						         sqlCmd = "DELETE WML01_LOCK WHERE m_year=?  AND m_month= ? "+
							 	        " AND bank_code=?  AND report_no=? ";
				                 paramList.add(S_YEAR) ;
				                 paramList.add(S_MONTH);
				                 paramList.add(bank_code) ;
				                 paramList.add(Report_no) ;
						      }else if(lock_status.equals("Y")){//鎖定						         
						         sqlCmd = "UPDATE WML01_LOCK SET lock_status =? WHERE m_year=? AND m_month= ? " +
							     	 " AND bank_code=? AND report_no=? ";
						         paramList.add(lock_status) ;
						         paramList.add(S_YEAR) ;
				                 paramList.add(S_MONTH);
				                 paramList.add(bank_code) ;
				                 paramList.add(Report_no) ;
						      }  						      
				           }else{
				              if(lock_status.equals("Y")){//鎖定				                 
						         sqlCmd = "Insert into WML01_LOCK VALUES(?,?,?,?,?,?,?,sysdate)";
				              	 paramList.add(S_YEAR) ;
				              	 paramList.add(S_MONTH) ;
				              	 paramList.add(bank_code);
				              	 paramList.add(Report_no) ;
				              	 paramList.add(lock_status) ;
				              	 paramList.add(lguser_id);
				              	 paramList.add(lguser_name) ;
						      }						      
				           }
				           //updateDBSqlList.add(sqlCmd);
				           this.updDbUsesPreparedStatement(sqlCmd,paramList) ;
					}	      				    
	            }//end of for		
	            errMsg = errMsg + "相關資料寫入資料庫成功";		
			    /*if(DBManager.updateDB(updateDBSqlList)){					 
				   errMsg = errMsg + "相關資料寫入資料庫成功";					
			    }else{
			 	   errMsg = errMsg + "相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
			    }*/				    
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