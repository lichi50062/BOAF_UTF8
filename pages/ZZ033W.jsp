<%
//95.02.16 create by 2295
//95.05.05 add A01與A04總放款金額不符/A01與A04總放款金額均為0/A01或A04總放款金額尚未申報 by 2295
//95.08.16 add 增加使用者姓名.電話.機構電話 by 2295
//95.09.11 fix sql以wtt01為主取得使用者資料 by 2295
//100.01.07 fix 根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@include file="./include/Header.include" %>

<%
	
   
	//登入者資訊
	String lguser_tbank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");			
	session.setAttribute("nowtbank_no",null);//94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除======	
	String bank_type = ( session.getAttribute("bank_type")==null ) ? "" : (String)session.getAttribute("bank_type");				
	
	
	
	String bank_type_list = Utility.getTrimString(dataMap.get("BANK_TYPE_List"));
	String S_YEAR = Utility.getTrimString(dataMap.get("S_YEAR"));
	String S_MONTH = Utility.getTrimString(dataMap.get("S_MONTH"));	
	String upd_code = Utility.getTrimString(dataMap.get("UPD_CODE"));
	    
    String tmpbank_type="";    
    if(lguser_id.equals("A111111111") || bank_type.equals("2") || bank_type.equals("1")){
       tmpbank_type = "Z";			    
	}else{
	   tmpbank_type = bank_type;				
	}	  
	
    if(!Utility.CheckPermission(request,report_no)){//無權限時,導向到LoginError.jsp    
        rd = application.getRequestDispatcher( LoginErrorPgName );        
    }else{            
    	//set next jsp 	    	
    	if( act.equals("Qry")){    	  
    	   if(!S_YEAR.equals("")){
               S_YEAR = String.valueOf(Integer.parseInt(S_YEAR));
            } 
            if(!S_MONTH.equals("")){
                S_MONTH = String.valueOf(Integer.parseInt(S_MONTH));
            }      	        	    
    	    rd = application.getRequestDispatcher( ListPgName +"?act=Qry");            	        	        	    	            
    	}else if( act.equals("List")){    	          	     
    	    List zz033wList = getQryResult(bank_type_list,S_YEAR,S_MONTH,upd_code,bank_type,lguser_tbank_no,lguser_id);    	       	        	    
    	    request.setAttribute("zz033wList",zz033wList);    	     	        	    
        	rd = application.getRequestDispatcher( ListPgName +"?act=List&BANK_TYPE_List="+bank_type_list+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&upd_code="+upd_code+"&tmpbank_type="+tmpbank_type);            	        	        	    	            
        }
       
    	request.setAttribute("actMsg",actMsg);
    	request.setAttribute("alertMsg",alertMsg);
    	request.setAttribute("webURL_Y",webURL_Y);    
    }        
     
%>

<%@include file="./include/Tail.include" %>

<%!
    private final static String report_no = "ZZ033W";
    private final static String nextPgName = "/pages/ActMsg.jsp";        
    private final static String ListPgName = "/pages/"+report_no+"_List.jsp";        
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
     
    //取得查詢結果
    //100.01.07 fix 根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
    private List getQryResult(String bank_type_list,String S_YEAR,String S_MONTH,String upd_code,String bank_type,String tbank_no,String lguser_id){    	   
    		//查詢條件        		
    		String sqlCmd = "";
    		String rule_1 = "";
    		String rule_2 = "";
    		String rule_3 = "";
    		String tmpbank_type="";
    		String tmpbank_no="";
    		//100.01.07 add 查詢年度100年以前.縣市別不同===============================
  	    	String cd01_table = (Integer.parseInt(S_YEAR) < 100)?"cd01_99":""; 
  	    	String wlx01_m_year = (Integer.parseInt(S_YEAR) < 100)?"99":"100"; 
  	    	List paramList = new ArrayList(); 
  	    	//=====================================================================    
    		if(lguser_id.equals("A111111111") || bank_type.equals("2") || bank_type.equals("1")){
    		    tmpbank_type = "Z";
			    tmpbank_no = "9999999";
		    }else{
		        tmpbank_type = bank_type;
				tmpbank_no = tbank_no;
			}	
			/*
			select * from (  
			   	        	select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name,  
    			    	   	       BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type ,  
  								   nvl(wlx01.CENTER_NO,' ') as CENTER_NO, 
				    			   nvl(wlx01.M2_NAME,' ')   as M2_NAME, 
				    			   wlx01.CANCEL_no, wlx01.CANCEL_date, 
				    			   wlx01.telno, bn01.m_telno,bn01.muser_name,--增加使用者姓名.電話.機構電話
    			    			   nvl(cd01.hsien_id,' ') as  hsien_id, 
				    			   a01_990000.amt  as a01_amt_990000, 
								   A01_total.amt as a01_amt_total,
				    			   a04_840740.amt  as a04_amt_840740,
								   a04_840750.amt  as a04_amt_840750,								
				    			   a05.amt  as a05_amt, 
				    			   nvl(cd01.hsien_name,'OTHER')  as  hsien_name, 
    			    		       cd01.FR001W_output_order     as  FR001W_output_order 
				            from  (select * from cd01 where  cd01.hsien_id <> 'Y') cd01 
     			      		      left join wlx01 on wlx01.hsien_id=cd01.hsien_id       			      		      
	 			    	          left join (select bn01.BANK_NO, bn01.BANK_NAME, bn01.bank_type,wtt01.muser_name,wtt01.m_telno  --增加使用者姓名.電話.機構電話
	                					     from bn01 left join (select wtt01.muser_id,wtt01.muser_name,muser_data.m_telno from wtt01 
											                     left join muser_data
											                          on wtt01.muser_id = muser_data.muser_id										       		 
											                     )wtt01 
											                 on bn01.bank_no || '001' = wtt01.muser_id 
											 where bn01.bank_type in('6','7')											 
											 )  bn01  
			        				   on wlx01.bank_no=bn01.bank_no  
				    			  left join (select  * FROM  A01 where  
              	    						 M_Year =  94  and M_month =  8
              	    					     and  acc_code in('990000') ) A01_990000											  
		    	    			        on 	bn01.bank_no = A01_990000.Bank_Code
								  left join (select m_year,m_month,bank_code,sum(amt) as amt 
								             from (select  * FROM  A01 where  
              	    						       M_Year =  94  and M_month =  8
              	    					           and  acc_code in('120000','120800','150300') ) A01_total
										     group by m_year,m_month,bank_code
										     order by m_year,m_month,bank_code		  
										     ) A01_total											  
		    	    			       on bn01.bank_no = A01_total.Bank_Code  		  
				    			  left join (select  * FROM  A04 where  
             	    						M_Year =  94  and M_month = 8
			  	    						and acc_code in('840740') ) A04_840740  
				    					on 	bn01.bank_no = A04_840740.Bank_Code
								  left join (select  * FROM  A04 where  
             	    						M_Year =  94  and M_month = 8
			  	    						and acc_code in('840750') ) A04_840750  
				    					on 	bn01.bank_no = A04_840750.Bank_Code  		  
				    			  left join (select  bank_code, round(amt /  1000  ,3)  as amt  FROM  A05  
				    					    where M_Year =  94 and M_month = 8
             	    			            and  acc_code = '91060P') A05  
			 	    			        on 	bn01.bank_no = A05.Bank_Code
						  )	Temp_Output  where 	BANK_NO <> ' '		
			order by  Bank_Type, FR001W_output_order, BANK_NO 
			*/
            rule_1 = " select * from "
				   + " ( "
				   + "	 select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name, "
    			   + "	   	    BN01.BANK_NO, BN01.BANK_NAME, bn01.bank_type , "
    			   + " 			nvl(wlx01.CENTER_NO,' ') as CENTER_NO,"
				   + "			nvl(wlx01.M2_NAME,' ')   as M2_NAME,"
				   + "			wlx01.CANCEL_no, wlx01.CANCEL_date,"
				   + "          wlx01.telno, bn01.m_telno,bn01.muser_name,"//95.08.16 增加使用者姓名.電話.機構電話 by 2295
    			   + "			nvl(cd01.hsien_id,' ') as  hsien_id,"
				   + "			a01_990000.amt  as a01_amt_990000," //A01逾期放款
				   + "          a01_total.amt   as a01_amt_total,"  //A01總放款
				   + "          a04_840740.amt  as a04_amt_840740,"//A04逾期放款
				   + "			a04_840750.amt  as a04_amt_840750,"//A04總放款						   
				   + "			a05.amt  as a05_amt,"//A05_BIS
				   + "			nvl(cd01.hsien_name,'OTHER')  as  hsien_name,"
    			   + "		    cd01.FR001W_output_order     as  FR001W_output_order"
				   + "  from  (select * from "+cd01_table+" cd01 where  cd01.hsien_id <> 'Y') cd01"
     			   + "  		left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id "
     			   + "			left join (select bn01.BANK_NO, bn01.BANK_NAME, bn01.bank_type,"//fix sql以wtt01為主取得使用者資料 by 2295
     			   + "					          wtt01.muser_name,wtt01.m_telno"//95.08.16 增加使用者姓名.電話.機構電話 by 2295
	               + " 					   from (select * from bn01 where m_year=?)bn01 left join (select wtt01.muser_id,wtt01.muser_name,muser_data.m_telno from wtt01 "
				   + "						  	                left join muser_data"
				   + "				                                 on wtt01.muser_id = muser_data.muser_id"										       		 
				   + "					                       )wtt01 "
				   + "					              on bn01.bank_no || '001' = wtt01.muser_id"
				   + "						where bn01.bank_type in('6','7')"
				   + " 						)  bn01  "
			       + " 				 on wlx01.bank_no=bn01.bank_no " 				    		
     			   //+ "	        left join (select bn01.BANK_NO, bn01.BANK_NAME,bn01.bank_type,"
     			   //+ "                            wtt01.muser_name,muser_data.m_telno "//95.08.16 增加使用者姓名.電話.機構電話 by 2295
	               //+ " 				       from bn01,muser_data,wtt01 "
	               //+ "                     where bn01.bank_type in('6','7') "
				   //+ "					   and muser_data.muser_id = bn01.bank_no || '001' "
                   //+ "				  	   and wtt01.muser_id = muser_data.muser_id "
				   //+ "					   ) bn01 " 
     			   /*fix 95.08.16
	 			   + "	        left join (select BANK_NO, BANK_NAME,  bank_type "
	               + "					   from bn01 where bn01.bank_type in('6','7'))  bn01 "*/			       
				   + "			left join  (select  * FROM  A01 where "
              	   + "						M_Year = "+S_YEAR + " and M_month = "+String.valueOf(Integer.parseInt(S_MONTH))
              	   + "					    and  acc_code in('990000') ) a01_990000 " //A01逾期放款(990000)
		    	   + "			        on 	bn01.bank_no = a01_990000.Bank_Code "
		    	   + "          left join  (select m_year,m_month,bank_code,sum(amt) as amt "
				   + "			            from  (select  * FROM  A01 where " 
				   + "						       M_Year = "+S_YEAR + " and M_month = "+String.valueOf(Integer.parseInt(S_MONTH))              	    						      
              	   + "					           and  acc_code in('120000','120800','150300') ) a01_total "
				   + "						group by m_year,m_month,bank_code "
				   + "					    order by m_year,m_month,bank_code "		  
				   + "						) a01_total "//A01總放款(120000+120800+150300)											  
		    	   + " 			         on bn01.bank_no = a01_total.Bank_Code " 	
				   + "			left join (select  * FROM  A04 where "
             	   + "						M_Year = "+S_YEAR + " and M_month = "+String.valueOf(Integer.parseInt(S_MONTH))
			  	   + "						and acc_code in('840740') ) a04_840740 "//A04逾期放款840740
				   + "					on 	bn01.bank_no = a04_840740.Bank_Code "
				   + "			left join (select  * FROM  A04 where "
             	   + "						M_Year = "+S_YEAR + " and M_month = "+String.valueOf(Integer.parseInt(S_MONTH))
			  	   + "						and acc_code in('840750') ) a04_840750 "//A04總放款840750
				   + "					on 	bn01.bank_no = a04_840750.Bank_Code "
				   + "			left join (select  bank_code, round(amt /  1000  ,3)  as amt  FROM  A05 "
				   + "					   where M_Year = "+S_YEAR + " and M_month = "+String.valueOf(Integer.parseInt(S_MONTH))
             	   + "			            and  acc_code = '91060P') A05 "
			 	   + "			        on 	bn01.bank_no = A05.Bank_Code ";
			paramList.add(wlx01_m_year);//wlx01
			paramList.add(wlx01_m_year);//bn01
			rule_2 = " ) 	Temp_Output  where 	BANK_NO <> ' '		and "
				   + " ( "
				   + " ('"+tmpbank_type+"' = 'Z') or "
			       + " (('"+tmpbank_type+"' = '6' or '"+tmpbank_type+"' = '7') and  '"+tmpbank_no+"' = BANK_NO) or "
			       + " ('"+tmpbank_type+"' = '8' and '"+tmpbank_no+"' = CENTER_NO) or"
			       + " ('"+tmpbank_type+"' = 'B' and '"+tmpbank_no+"' = M2_NAME) "
				   + " )    and "
				   + " ( "
				   + " ('"+bank_type_list+"'= 'ALL') or "
 			       + " (('"+bank_type_list+"'= '6') and (Bank_Type = '6')) or "
				   + " (('"+bank_type_list+"'= '7') and (Bank_Type = '7')) "
				   + " )    and "
 				   + " ( "
				   + " ('"+upd_code+"'= 'ALL')  or "
			       + " ('"+upd_code+"'= '0' and  " //有任一不符
				   + " ((a01_amt_990000 is null) or (a04_amt_840740 is null) or (a01_amt_total is null) or (a04_amt_840750 is null) or (a05_amt is null) or "
  				   + " ((a01_amt_990000 is not null and a04_amt_840740 is not null) and "
				   + " (a01_amt_990000 <> a04_amt_840740))                          or "
				   + " ((a01_amt_total is not null and a04_amt_840750 is not null) and "
				   + " (a01_amt_total <> a04_amt_840750))                          or "
				   + " ((a01_amt_990000 is not null and a04_amt_840740 is not null) and "
				   + " (a01_amt_990000 = 0  and a04_amt_840740 = 0))                or "
				   + " ((a01_amt_total is not null and a04_amt_840750 is not null) and "
				   + " (a01_amt_total = 0  and a04_amt_840750 = 0))                or "
				   + " ((a05_amt is not null) and (a05_amt = 0)) "
				   + " ) "
				   + " )    OR "
				   + " ('"+upd_code+"'= '1' and  " //A01與A04申報逾放金額不符
  				   + " ((a01_amt_990000 is not null and a04_amt_840740 is not null) and "
				   + "	(a01_amt_990000 <> a04_amt_840740)) "                                           
				   + " )    OR "
				   + " ('"+upd_code+"'= '1a' and  " //A01與A04總放款金額不符
  				   + " ((a01_amt_total is not null and a04_amt_840750 is not null) and "
				   + "	(a01_amt_total <> a04_amt_840750)) "                                           
				   + " )    OR "
				   + " ('"+upd_code+"'= '2' and  " //A01與A04申報逾放金額均為0
  			       + " ((a01_amt_990000 is not null and a04_amt_840740 is not null) and "
				   + " (a01_amt_990000 = 0 and  a04_amt_840740 = 0)) "                                           
				   + " )    OR "
				   + " ('"+upd_code+"'= '2a' and  " //A01與A04總放款金額均為0
  			       + " ((a01_amt_total is not null and a04_amt_840750 is not null) and "
				   + " (a01_amt_total = 0 and  a04_amt_840750 = 0)) "  
				   + " )    OR "
				   + " ('"+upd_code+"'= '3' and  " //A01或A04逾放金額尚未申報
 				   + " (a01_amt_990000 is null or a04_amt_840740 is null) "                    
 				   + " )    OR "
				   + " ('"+upd_code+"'= '3a' and  " //A01或A04總放款金額尚未申報
 				   + " (a01_amt_total is null or a04_amt_840750 is null) "       
				   + " )    OR "
 				   + " ('"+upd_code+"'= '4' and  " //A05_BIS未申報
 				   + " (a05_amt is null) "                    
				   + " )    OR "
				   + " ('"+upd_code+"'= '5' and  " //A05_BIS申報值為0
 				   + " (a05_amt is not null and a05_amt = 0) "                    
				   + " ) "                                                                                   
				   + " ) order by  Bank_Type, FR001W_output_order, BANK_NO ";
			
    		sqlCmd = rule_1 + rule_2;
    		
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"cancel_date,a01_amt_990000,a01_amt_total,a04_amt_840740,a04_amt_840750,a05_amt");            
            return dbData;
    }
%>    