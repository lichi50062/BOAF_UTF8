<%
//102.08.06 created by 2968
//103.03.17 結合RptWR003W、RptWR004W、RptWR005W by 2968
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.report.*" %>
<%@ page import="com.tradevan.util.ftp.MyFTPClient" %>
<%@ page import="com.tradevan.util.AutoWR098W" %>
<%@ page import="java.util.*,java.io.*" %>
<%@include file="./include/Header.include" %>
<%				
	String lguser_tbank_no = ( session.getAttribute("tbank_no")==null ) ? "" : (String)session.getAttribute("tbank_no");			
	session.setAttribute("nowtbank_no",null);//94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除======	
	String bank_type = ( session.getAttribute("bank_type")==null ) ? "" : (String)session.getAttribute("bank_type");				
	
	
	String szrpt_code = Utility.getTrimString(dataMap.get("RPT_CODE"));    
	String szrpt_output_type = Utility.getTrimString(dataMap.get("szRPT_OUTPUT_TYPE"));    
	String szhsien_id = Utility.getTrimString(dataMap.get("HSIEN_ID"));    
	String S_YEAR = Utility.getTrimString(dataMap.get("S_YEAR"));    
	String S_MONTH = Utility.getTrimString(dataMap.get("S_MONTH"));    
	String szUnit = Utility.getTrimString(dataMap.get("Unit"));    
	String firstStatus = Utility.getTrimString(dataMap.get("firstStatus"));    
	String filename = Utility.getTrimString(dataMap.get("filename"));    
	String online = Utility.getTrimString(dataMap.get("online"));    
	
		
    if(!Utility.CheckPermission(request,report_no)){//無權限時,導向到LoginError.jsp//95.11.03 fix 使用Utility.CheckPermission檢核權限 by 2295
        rd = application.getRequestDispatcher( LoginErrorPgName );        
    }else{            
    	//set next jsp 	    	
    	if(act.equals("Qry") || act.equals("goQry")){                    	        	    
    	    rd = application.getRequestDispatcher( ListPgName +"?act=Qry&firstStatus="+firstStatus);            	        	        	    	    
    	}else if(act.equals("List")){                    	    
    	    //List dbData = getQryResult(bank_type,lguser_tbank_no,szrpt_code,szrpt_output_type,szhsien_id,S_YEAR,S_MONTH,E_YEAR,E_MONTH,szrpt_sort,lguser_id);    	       
    	    List dbData = getQryResult(szrpt_code,szrpt_output_type,szhsien_id,S_YEAR,S_MONTH);    	       
    	    request.setAttribute("reportList",dbData);    	     	        	        	     	    
    	    rd = application.getRequestDispatcher( ListPgName +"?act=List&RPT_CODE="+szrpt_code+"&szRPT_OUTPUT_TYPE="+szrpt_output_type+"&HSIEN_ID="+szhsien_id+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&Unit="+szUnit);            	        	        	    	        	    
    	}else if(act.equals("GenerateRpt")){		        		        	
    	    //List updateDBSqlList = GenerateRpt(request,lguser_id,lguser_name,S_YEAR,S_MONTH,szrpt_code, szrpt_output_type, szUnit,szhsien_id);    	           			                
		    actMsg = GenerateRpt(request,lguser_id,lguser_name,S_YEAR,S_MONTH,szrpt_code, szrpt_output_type, szUnit,szhsien_id);    	           			                
		   
            if(actMsg.indexOf("相關資料寫入資料庫成功") != -1){								 
		       alertMsg = "本作業執行完成";    	           	       	
    	       webURL_Y = ListPgName +"?act=Qry&firstStatus="+firstStatus;    	       	   	   	
	        }else{
		       alertMsg = "報表產生完成,但上傳至Sever未成功";    	           	       	
    	       webURL_Y = ListPgName +"?act=Qry&firstStatus="+firstStatus;    	       	   	   	    	     
         	}
         	rd = application.getRequestDispatcher( nextPgName );                       
    	}else if(act.equals("GenerateRptALL")){//95.04.21 add 全部報表發佈		
    	   if(online.equals("true")){//全部報表執行完成    	      
      	      alertMsg = "全部報表發佈作業執行完成";    	           	       	
    	      webURL_Y = ListPgName +"?act=Qry&firstStatus="+firstStatus;   
    	      rd = application.getRequestDispatcher( nextPgName );                       
    	   }else{ 
    	      rd = application.getRequestDispatcher("/pages/WMAutoWR098W.jsp?report_no=ALL&lguser_id="+lguser_id+"&lguser_name="+lguser_id+"&unit="+szUnit+"&S_YEAR="+S_YEAR+"&S_MONTH="+S_MONTH+"&online=true");         	
    	   }    	   
    	}
    	request.setAttribute("actMsg",actMsg);    
    	request.setAttribute("alertMsg",alertMsg);
    	request.setAttribute("webURL_Y",webURL_Y);
    }        
     
%>

<%@include file="./include/Tail.include" %>

<%!
    private final static String report_no = "WR098W";
    private final static String nextPgName = "/pages/ActMsg.jsp";        
    private final static String ListPgName = "/pages/"+report_no+"_List.jsp";        
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    private final static String[][] rptfilename = {												  
        											{"WR003W_ATOT", "農漁會信用部營運狀況警訊報表-與上月比較.xls"},
	      											{"WR004W_ATOT", "農漁會信用部營運狀況警訊報表-與上季比較.xls"},
	      											{"WR005W_ATOT", "農漁會信用部營運狀況警訊報表-與上年度同期比較.xls"},
	      											{"WR006W_ATOT", "農漁會信用部營運狀況警訊報表(專案農貸)-與上月比較.xls"},
	      											{"WR007W_ATOT", "農漁會信用部營運狀況警訊報表(專案農貸)-與上季比較.xls"},
	      											{"WR008W_ATOT", "農漁會信用部營運狀況警訊報表(專案農貸)-與上年度同期比較.xls"}	
											      };
        
    //取得查詢結果
    private List getQryResult(String szrpt_code,String szrpt_output_type,String szhsien_id,String S_YEAR,String S_MONTH){    	       
            //100.01.04 add 查詢年度100年以前.縣市別不同===============================
  	    	String cd01_table = (Integer.parseInt(S_YEAR) < 100)?"cd01_99":""; 
  	    	String wlx01_m_year = (Integer.parseInt(S_YEAR) < 100)?"99":"100"; 
  	    	//=====================================================================     
    		//查詢條件        		
    		String sqlCmd = "";                                                
    		String rule_1 = "";
    		String rule_2 = "";
    		String rule_3 = "";
    		String rule_4 = "";
    		List dbData = null;
    		List paramList =new ArrayList() ;
    		List paramList_rule_1 =new ArrayList() ;
    		List paramList_rule_2 =new ArrayList() ;
            sqlCmd = " select rpt_code,rpt_name,rpt_output_type,rpt_include "
		   		   + " from rpt_nof "
		   		   + " where rpt_nof.rpt_code = ?";
		   	paramList.add(szrpt_code);	   
    		dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");//報表data    
    		String szrpt_include = (String)((DataObject)dbData.get(0)).getValue("rpt_include");
    		String rpt_output_type = (String)((DataObject)dbData.get(0)).getValue("rpt_output_type");
    		System.out.print("szrpt_code="+szrpt_code);
            System.out.print(":szrpt_output_type="+szrpt_output_type);
            System.out.print(":szhsien_id="+szhsien_id);
            System.out.print(":rpt_output_type="+rpt_output_type);
            System.out.print(":szrpt_include="+szrpt_include);
            System.out.print(":S_YEAR="+S_YEAR);
            System.out.print(":S_MONTH="+S_MONTH);
    		String condition = " rpt_dirf.RPT_Code =? and "
	 					     + " rpt_dirf.M_Year = ? and "
	 					     + " rpt_dirf.M_Month = ?";
      	 	List paramList_condition =new ArrayList();		
      	 	paramList_condition.add(szrpt_code);
      	 	paramList_condition.add(S_YEAR);
      	 	paramList_condition.add(S_MONTH);
      	 	
            String condition_1 = " (select  * FROM  WML01 "
          				       + "  where WML01.M_Year = ?"
                          	   + "	and  WML01.M_Month = ?";
            List paramList_condition_1 =new ArrayList();	
            paramList_condition_1.add(S_YEAR);
      	 	paramList_condition_1.add(S_MONTH);	                          	   
            if(rpt_output_type.equals("X") && szrpt_include.equals("X") && szrpt_code.equals("FR026W_DET")){               
                  condition_1 += "	and WML01.Report_No = ?) WML01 ";
                  paramList_condition_1.add("F01");
            }else{     
                  condition_1 += "	and WML01.Report_No = ?) WML01 ";     
                  paramList_condition_1.add("A01");          
            }              	   
            if(rpt_output_type.equals("X") && szrpt_include.equals("T")){   
               if(szrpt_code.equals("FR025W_STOT") || szrpt_code.equals("FR025WA_SDET") || szrpt_code.equals("FR027W_STDET")){
                  paramList_condition_1 =new ArrayList();	
                  condition_1 = " (select 'U' as Upd_Code, Bank_Code"
 						      + " from WLX_APPLY_LOCK "
 						      + " where WLX_APPLY_LOCK.M_Year = ?"
 						      + " and   WLX_APPLY_LOCK.M_QUARTER = ?";
 				  paramList_condition_1.add(S_YEAR);	
 				  paramList_condition_1.add(String.valueOf(Integer.parseInt(S_MONTH)));     
 				  if(szrpt_code.equals("FR025W_STOT") || szrpt_code.equals("FR025WA_SDET")){
 				     condition_1 += " and   WLX_APPLY_LOCK.Report_No = ?) WML01";
 				     paramList_condition_1.add("C01");
 				  }else if(szrpt_code.equals("FR027W_STDET")){		      
 				     condition_1 += " and   WLX_APPLY_LOCK.Report_No = ?) WML01";
 				     paramList_condition_1.add("C07");
 				  } 						       				  
               }
            }              	   
            condition_1 += " on 	bn01.bank_no = WML01.Bank_Code  "                          	   
          			    +  "  ) temp_2 ";	 			
            
            if((rpt_output_type.equals("X") && szrpt_include.equals("B")) 
            || (rpt_output_type.equals("X") && szrpt_include.equals("X"))){   
                                   
            	rule_1 = " select NVL((BN01.BANK_NO || BN01.BANK_NAME), ' ')  as S_Report_Name,"
    			       + " BN01.BANK_NO, BN01.BANK_NAME,"
    				   + " bn01.bank_type as  Bank_Type,"
    				   + " decode(bn01.bank_type,'6','農會','7','漁會',' ')  as  Bank_Type_Name,"
    				   + " nvl(cd01.hsien_id,' ') as  hsien_id,"
					   + " nvl(cd01.hsien_name,'OTHER')  as  hsien_name,"
    				   + " cd01.FR001W_output_order     as  FR001W_output_order,"
   					   + " nvl(rpt_dirf.RPT_Fname,' ')  as RPT_Fname,"
   					   + " nvl(rpt_dirf.RPT_FSize,' ')  as RPT_FSize,"
   					   + " nvl(((to_char(rpt_dirf.UPDATE_DATE,'yyyymmdd')-19110000)  || to_char(rpt_dirf.UPDATE_DATE,' hh24:mi')),' ')  as S_UpdateDate "  
					   + " from  (select * from "+cd01_table+" cd01 where  cd01.hsien_id <> 'Y') cd01";
     		  
           	  // 報表發佈無法發佈新北市、臺中市、臺南市、高雄市 增加 upper(cd01.hsien_id) By 2479
     		  if(szrpt_include.equals("B")){ 			  
     		      rule_1 += " left join "
     		              + " ((select BANK_NO, BANK_NAME, 6 as bank_type "
	               		  + "   from (select * from bn01 where m_year=?)bn01 where bn01.bank_type='B') "
	               		  + " union all "
				 		  + "  (select BANK_NO, BANK_NAME, 7 as bank_type "
	               		  + "   from (select * from bn01 where m_year=?)bn01 where bn01.bank_type='B'))  bn01 "
	             		  + " on  (upper(cd01.hsien_id) = SUBSTR(bn01.bank_no,1,1) and cd01.hsien_id <> 'h')";//100.06.28排除台灣省    
	              paramList_rule_1.add(wlx01_m_year);
			      paramList_rule_1.add(wlx01_m_year);		  
     		  }else if(szrpt_include.equals("X")){ 			   			  
     		      rule_1 += " left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id " 
     		              + " left join "
     		              + " (select BANK_NO, BANK_NAME,  bank_type ";
     		              if(szrpt_output_type.equals("ALL")){
	               		     rule_1 += "   from (select * from bn01 where m_year=?)bn01 where bn01.bank_type in('6','7'))  bn01";
	               		  }else if(szrpt_output_type.equals("6")){
	               		     rule_1 += "   from (select * from bn01 where m_year=?)bn01 where bn01.bank_type in('6'))  bn01";
	               		  }else if(szrpt_output_type.equals("7")){
	               		     rule_1 += "   from (select * from bn01 where m_year=?)bn01 where bn01.bank_type in('7'))  bn01";
	               		  }	               		  
			      rule_1 += "   on wlx01.bank_no=bn01.bank_no ";               
			      paramList_rule_1.add(wlx01_m_year);
			      paramList_rule_1.add(wlx01_m_year);
     		  }			       					  
	 			  rule_1 += " left join (select * from rpt_dirf "
	 				   	  + "            where "+condition+") rpt_dirf "
	             		  + " 			 on  SUBSTR(rpt_dirf.RPT_Fname,4,7) = bn01.bank_no and SUBSTR(rpt_dirf.RPT_Fname,1,1) = bn01.bank_type "
						  + " order by  bn01.Bank_Type, BN01.BANK_NO ";
                  for(int i=0;i<paramList_condition.size();i++){
                      paramList_rule_1.add(paramList_condition.get(i));
                  }
            	   rule_2 = " select  bank_type,";
            	if(szrpt_include.equals("B")){ 			 
            	   rule_2 += " hsien_id,";
            	}else if(szrpt_include.equals("X")){ 			 
            	   rule_2 += " bank_no,";
            	}
            	   rule_2 += " count(*)  as  T_Cnt,"
		 				  + " SUM(decode(temp_2.Upd_Code,'U',1,0)) as  OK_Cnt"
						  + " from("
						  + "	    select  nvl(temp_1.bank_no,' ')   as bank_no, "
        				  + "   			nvl(temp_1.hsien_id,' ')  as hsien_id, "
	    				  + "				nvl(bn01.bank_type,' ')   as bank_type,"
						  + "				nvl(WML01.Upd_Code,' ')   as Upd_Code"
						  + "		from ("
						  +	"				select wlx01.bank_no, wlx01.hsien_id from (select * from wlx01 where m_year=?)wlx01 "
     					  + "				where ((wlx01.CANCEL_NO = 'Y' and (to_char(wlx01.cancel_date,'yyyymm')-191101) <= ?) or (wlx01.CANCEL_NO <> 'Y'))"			  
     					  + "				and    wlx01.hsien_id in (select cd01.hsien_id from "+cd01_table+" cd01 where  cd01.hsien_id <> 'Y' ";
     			paramList_rule_2.add(wlx01_m_year);		
     			paramList_rule_2.add(S_YEAR+S_MONTH);  
     			if(szrpt_include.equals("B")){ 			 		  
     			    rule_2 += " )";
     			}else if(szrpt_include.equals("X")){ 			 		  
     				rule_2 += "  and hsien_id =?)";
     				paramList_rule_2.add(szhsien_id);  
     		    }			  
					rule_2 += "	         ) temp_1  left join (select * from bn01 where m_year=?)bn01 on temp_1.bank_no = bn01.bank_no and ";
					paramList_rule_2.add(wlx01_m_year);		
						  if(szrpt_output_type.equals("ALL")){
	               		     rule_2 += " bn01.bank_type in('6','7')";
	               		  }else if(szrpt_output_type.equals("6")){
	               		     rule_2 += " bn01.bank_type in('6')";
	               		  }else if(szrpt_output_type.equals("7")){
	               		     rule_2 += " bn01.bank_type in('7')";
	               		  }
						  
          			rule_2 += " left join "+condition_1+" where bank_type <> ' '";
          			for(int i=0;i<paramList_condition_1.size();i++){
                      paramList_rule_2.add(paramList_condition_1.get(i));
                    }
                if(szrpt_include.equals("B")){ 			           				  
                   rule_2 += " group by bank_type,  hsien_id"
		  			      + " order by bank_type,  hsien_id";		  				  
                }else if(szrpt_include.equals("X")){ 			           				  
                   rule_2 += " and bank_no <> ' '"
                          + " group by bank_type,  bank_no"
		  		 	      + " order by bank_type,  bank_no";		  				  
                }
          		paramList = new ArrayList(); 	  
				   sqlCmd = " Select   MTemp_2.* "
				  		  + " from ("
				   		  + "       select   MTemp_1.* ,  nvl(temp_3.T_Cnt,0)  as  T_Cnt ,  nvl(temp_3.OK_Cnt,0) as OK_Cnt"
				   		  + "		 from ( "+rule_1+")  MTemp_1 "
				   		  + "		 left join ("+rule_2+") temp_3 on ";
				for(int i=0;i<paramList_rule_1.size();i++){
                    paramList.add(paramList_rule_1.get(i));
                }  
                for(int i=0;i<paramList_rule_2.size();i++){
                    paramList.add(paramList_rule_2.get(i));
                } 		  
				if(szrpt_include.equals("B")){ 			           				  
				   sqlCmd += " MTemp_1.bank_type = temp_3.bank_type and "			 
             	   		   + " MTemp_1.hsien_id = temp_3.hsien_id "
				   		   + " ) MTemp_2  where T_CNT > 0 and ";
				}else if(szrpt_include.equals("X")){ 		
				   sqlCmd += " MTemp_1.bank_no = temp_3.bank_no "
						  +  " ) Mtemp_2  where T_CNT > 0  or (bank_no <> ' ' ) and";	           				  
				}
				   sqlCmd += " ((? = 'ALL') or (? <> 'ALL' and  hsien_id =?))"                    
                   		  + "  and ((?='ALL') OR (?='6' and  bank_type =  '6')"
			       		  + "  OR (?='7' and  bank_type =  '7')) "                                                     			       
				          + " order by  MTemp_2.bank_type,  MTemp_2.FR001W_output_order,";
				paramList.add(szhsien_id);				          
				paramList.add(szhsien_id);
				paramList.add(szhsien_id);
				paramList.add(szrpt_output_type);
				paramList.add(szrpt_output_type);
				paramList.add(szrpt_output_type);
				if(szrpt_include.equals("B")){ 			           				  
				   sqlCmd +=" MTemp_2.hsien_id";
				}else if(szrpt_include.equals("X")){ 			           				  
				   sqlCmd +=" MTemp_2.bank_no";
				}
	        }
	        if((rpt_output_type.equals("X") && szrpt_include.equals("T")) 
            || (rpt_output_type.equals("T") && szrpt_include.equals("T"))){                                  
	            rule_1 = " select ?  as S_Report_Name,"
    				   + " ' '  as BANK_NO,   ' ' as  BANK_NAME,";
    		   	paramList_rule_1.add(szrpt_code);	   
    			if(rpt_output_type.equals("X")){ 	   
    			   rule_1 += " bn01.bank_type as  Bank_Type,"
    				      +  " decode(bn01.bank_type,'6','農會','7','漁會',' ')  as  Bank_Type_Name,";
    			}else if(rpt_output_type.equals("T")){ 	   
    			   rule_1 += " 'T' as  Bank_Type,"
    					  +  "'農漁會'    as  Bank_Type_Name,";
    			}
    			rule_1 += " ' ' as  hsien_id,"
					    + " ' '  as  hsien_name,"
    				    + " ' '  as  FR001W_output_order,"
   					    + " nvl(rpt_dirf.RPT_Fname,' ')  as RPT_Fname,"
   					    + " nvl(rpt_dirf.RPT_FSize,' ')  as RPT_FSize,"
   					    + " nvl(((to_char(rpt_dirf.UPDATE_DATE,'yyyymmdd')-19110000)  || to_char(rpt_dirf.UPDATE_DATE,' hh24:mi')),' ')  as S_UpdateDate   "
					    + " from  ";
               if(rpt_output_type.equals("X")){ 	
                  rule_1 += "(select bank_type from (select * from bn01 where m_year=?)bn01 where bn01.bank_type = '6' group by bank_type "
	   					  + " union all"
       					  + " select bank_type from (select * from bn01 where m_year=?)bn01 where bn01.bank_type = '7' group by bank_type"
	   					  + " )  bn01"         
	 					  + " left join (select * from rpt_dirf where " + condition+ " ) rpt_dirf"
	            		  + "            on  SUBSTR(rpt_dirf.RPT_Fname,1,1) =  bn01.Bank_Type "
						  + " order by  bn01.Bank_Type";   
				 paramList_rule_1.add(wlx01_m_year);	  
				 paramList_rule_1.add(wlx01_m_year);	
				 for(int i=0;i<paramList_condition.size();i++){
                      paramList_rule_1.add(paramList_condition.get(i));
                 }   
			   }else if(rpt_output_type.equals("T")){ 
			      rule_1 += "(SELECT 'T'  as  bank_type  from DUAL )  bn01 "
       					  + " LEFT JOIN "
         				  + " rpt_dirf on "+ condition;
         		  for(int i=0;i<paramList_condition.size();i++){
                      paramList_rule_1.add(paramList_condition.get(i));
                  } 		  
			   } 			  
			   if(rpt_output_type.equals("X")){ 	
			      rule_2 = " select bank_type,";	
			   }else if(rpt_output_type.equals("T")){ 	
			      rule_2 = "select 'T' as bank_type,";
			   }
			   rule_2 += " count(*)  as  T_Cnt,"
		 			   + " SUM(decode(temp_2.Upd_Code,'U',1,0)) as  OK_Cnt"
					   + " from("
					   + "		select  nvl(temp_1.bank_no,' ')   as bank_no, "
        			   + "				nvl(temp_1.hsien_id,' ')  as hsien_id, "
	    			   + "				nvl(bn01.bank_type,' ')   as bank_type,"
					   + "				nvl(WML01.Upd_Code,' ')   as Upd_Code"
					   + "		from ("
					   + "		 		select wlx01.bank_no, wlx01.hsien_id from (select * from wlx01 where m_year=?)wlx01 "
     				   + "				where ((wlx01.CANCEL_NO = 'Y'  "
	         		   + "				and (to_char(wlx01.cancel_date,'yyyymm')-191101) <= ?)   "
			  		   + "				or (wlx01.CANCEL_NO <> 'Y')) "
			  		   + "				and wlx01.hsien_id in (select cd01.hsien_id from "+cd01_table+" cd01 where  cd01.hsien_id <> 'Y')"
					   + "      ) temp_1  left join (select * from bn01 where m_year=?)bn01 on temp_1.bank_no = bn01.bank_no and ";
			   paramList_rule_2.add(wlx01_m_year);	   
			   paramList_rule_2.add(S_YEAR+S_MONTH);
			   paramList_rule_2.add(wlx01_m_year);
			   
			   if(szrpt_output_type.equals("ALL")){
	              rule_2 += " bn01.bank_type in('6','7')";
	           }else if(szrpt_output_type.equals("6")){
	              rule_2 += " bn01.bank_type in('6')";
	           }else if(szrpt_output_type.equals("7")){
	              rule_2 += " bn01.bank_type in('7')";
	           }	   					   
          	   rule_2 += " left join " + condition_1;
          	   for(int i=0;i<paramList_condition_1.size();i++){
                   paramList_rule_2.add(paramList_condition_1.get(i));
               } 	   
			   if(rpt_output_type.equals("X")){ 	
			      rule_2 += " where bank_type <> ' '"
                		  + " group by bank_type";	
			   }else if(rpt_output_type.equals("T")){ 	
			      rule_2 += " where bank_type <> ' '";
			   }
	           paramList = new ArrayList(); 
	           sqlCmd = " Select   MTemp_2.* "
					  + " from ("
					  + "       select   MTemp_1.* ,  nvl(temp_3.T_Cnt,0)  as  T_Cnt ,  nvl(temp_3.OK_Cnt,0) as OK_Cnt"
				   	  + "		 from ( "+rule_1+")  MTemp_1 "
				   	  + "		 left join ("+rule_2+") temp_3 on MTemp_1.bank_type = temp_3.bank_type"
				      + " 		            ) MTemp_2  where T_CNT > 0";   
				for(int i=0;i<paramList_rule_1.size();i++){
                    paramList.add(paramList_rule_1.get(i));
                }  
                for(int i=0;i<paramList_rule_2.size();i++){
                    paramList.add(paramList_rule_2.get(i));
                }       
			  if(rpt_output_type.equals("X")){ 	
			     sqlCmd += "  and ((?='ALL') OR (?='6' and  bank_type =  '6')"
			       		 + "  OR (?='7' and  bank_type =  '7')) ";  
			     paramList.add(szrpt_output_type);  		      
			     paramList.add(szrpt_output_type);
			     paramList.add(szrpt_output_type);			    
			  }
			  sqlCmd += " order by  MTemp_2.bank_type";  			  
	        }
    		dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"t_cnt,ok_cnt");
    		if(dbData != null && dbData.size() != 0){
    		   System.out.println("QrydbData.size="+dbData.size());  
    		}else{
    		   System.out.println("dbData is null or size = 0");  
    		}
    		
			return dbData;
    }
    
    public String GenerateRpt(HttpServletRequest request,String lguser_id,String lguser_name,String S_YEAR,String S_MONTH,String szrpt_code, String szrpt_output_type, String szUnit, String szhsien_id){    	        
        String errMsg = "";
        String putMsg = "";	
    	String BANK_NO = "";
		String BANK_NAME = "";
		String HSIEN_ID = "";		  		
		String scSrcFile = "";
		String copyOK = "";		  		
		String sqlCmd = "";
		String idxRptCode="";
		String filename="";
		String bank_type="";
		String rptLine="";
		List dbData = null;	
	    String szrpt_sort="";
	    String[] fname1;		    
        List filename_List  = new LinkedList();	
		List paramList =new ArrayList() ;		
		List<List> updateDBList = new ArrayList<List>();//0:sql 1:data		
		List updateDBSqlList = new ArrayList();
		List<List> updateDBDataList_insert_rpt_dirf= new ArrayList<List>();//儲存參數的List		
		List<List> updateDBDataList_update_rpt_dirf= new ArrayList<List>();//儲存參數的List		
		List<List> updateDBDataList_rpt_dirf_log= new ArrayList<List>();//儲存參數的List	
		List<String> dataList =  new ArrayList<String>();//儲存參數的data
    	try{
		  		String rptIP=Utility.getProperties("rptIP");			
	            String rptID=Utility.getProperties("rptID");			
	            String rptPwd=Utility.getProperties("rptPwd");	
		  		//取出form裡的所有變數
		  		List rptData = getFormData(request);		  			
		  		System.out.println("rptData.size="+rptData.size());		  					    			    
		  		//取得報表data
			    sqlCmd = " select rpt_code,rpt_name,rpt_output_type,rpt_include "
		   		   	   + " from rpt_nof "
		   		   	   + " where rpt_nof.rpt_code = ?";
		   		paramList.add(szrpt_code);   	   
    			dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");//報表data    
    			
    			String szrpt_include = (String)((DataObject)dbData.get(0)).getValue("rpt_include");
    			String dbrpt_output_type = (String)((DataObject)dbData.get(0)).getValue("rpt_output_type");
			    //新增GenerateRptDir         		
		  		File ClientRptDir = new File(Utility.getProperties("GenerateRptDir"));        
	        	if(!ClientRptDir.exists()){
         			if(!Utility.mkdirs(Utility.getProperties("GenerateRptDir"))){
         		   		errMsg+=Utility.getProperties("GenerateRptDir")+"目錄新增失敗";
         			}    
        		}
		  		if(errMsg.equals("")){			  		   
		  		   for(int i=0;i<rptData.size();i++){			
		  		      System.out.println("rptData="+(String)rptData.get(i));		       			        
		  		      rptLine = (String)rptData.get(i);
		  		      bank_type = rptLine.substring(0,rptLine.indexOf(":"));     	         			  
		  		      System.out.print("bank_type="+bank_type);
		  		      rptLine = rptLine.substring(rptLine.indexOf(":")+1,rptLine.length());
		  		      HSIEN_ID = rptLine.substring(0,rptLine.indexOf(":"));     	  
		  		      System.out.print(":HSIEN_ID="+HSIEN_ID);         			         			  
		  		      rptLine = rptLine.substring(rptLine.indexOf(":")+1,rptLine.length());
         			  filename = rptLine.substring(0,rptLine.indexOf(":"));     	         			  
         			  System.out.print(":filename="+filename);         			  
         			  rptLine = rptLine.substring(rptLine.indexOf(":")+1,rptLine.length());
         			  BANK_NO = rptLine.substring(0,rptLine.indexOf(":")+8);     	         			           			  
         			  System.out.print(":BANK_NO="+BANK_NO);         			  
         			  BANK_NAME = rptLine.substring(rptLine.indexOf(":")+8,rptLine.length());     	         			           			           			  
         			  System.out.println(":BANK_NAME="+BANK_NAME);
         			  //產生報表==========================================================================================
         			  errMsg += CreateRpt(S_YEAR,S_MONTH,BANK_NO,BANK_NAME,szrpt_code,bank_type,szUnit,HSIEN_ID);
         			  //=======================================================================================================
         			  scSrcFile = getFileName(szrpt_code,bank_type);//找出實際產生的報表檔名
         			  
         			  if(scSrcFile.indexOf(".xls") != -1){//有找到實際檔名         			     
         			        if((dbrpt_output_type.equals("X") && szrpt_include.equals("T")) || 
         			           (dbrpt_output_type.equals("T") && szrpt_include.equals("T")))
         			        {
         			           filename = bank_type;
         			           if(dbrpt_output_type.equals("T")){
         			              filename = dbrpt_output_type;
         			           }         			           
         			           
         			           filename += szrpt_include+"_";
         			           if(S_YEAR.length() == 2){
         			              filename += "0"+S_YEAR+S_MONTH+".xls";
         			           }else{
         			              filename += S_YEAR+S_MONTH+".xls";
         			           }         			                    			        
         			        }else if(dbrpt_output_type.equals("X") && szrpt_include.equals("B")){ 
         			          filename=bank_type+szrpt_include+"_"+BANK_NO+".xls";
         			        }else{
         			          filename=bank_type+bank_type+"_"+BANK_NO+".xls";
         			        }
         			        System.out.println("filename="+filename);
         			   
         			     copyOK = Utility.CopyFile(Utility.getProperties("reportDir") + System.getProperty("file.separator")+scSrcFile,Utility.getProperties("GenerateRptDir")+System.getProperty("file.separator")+filename);
         			     
         			     if(copyOK.equals("0")){
         			        File tmpFile = new File(Utility.getProperties("reportDir") + System.getProperty("file.separator")+scSrcFile);
         			        paramList = new LinkedList();
         			        sqlCmd = " select count(*) as count from rpt_dirf"
         			               + " where rpt_code=?"
         			               + " and   m_year = ?"
         			               + " and   m_month = ?"
         			               + " and   rpt_fname = ?";
         			         paramList.add(szrpt_code);      
         			         paramList.add(S_YEAR);      
         			         paramList.add(S_MONTH);      
         			         paramList.add(filename);      
         			           
         			         dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"count"); 
         			         dataList =  new ArrayList<String>();
         			         if(dbData != null && dbData.size() != 0 ){
         			            if(((((DataObject)dbData.get(0)).getValue("count")).toString()).equals("0")){
         			               dataList.add(szrpt_code);
         			               dataList.add(S_YEAR);
         			               dataList.add(S_MONTH);
         			               dataList.add(filename);
         			               dataList.add(String.valueOf(tmpFile.length()));
         			               dataList.add(lguser_id);
         			               dataList.add(lguser_name);
         			               updateDBDataList_insert_rpt_dirf.add(dataList);     
         			               System.out.println("add.insert filename="+filename);  			                       			                 			            
         			            }else{      			           
         			               dataList =  new ArrayList<String>(); 
         			               dataList.add(lguser_id);		  
         			               dataList.add(lguser_name);
         			               dataList.add(szrpt_code);
         			               dataList.add(S_YEAR);
         			               dataList.add(S_MONTH);
         			               dataList.add(filename);
         			               updateDBDataList_rpt_dirf_log.add(dataList);    
         			               System.out.println("log.insert filename="+filename);  
         			               dataList =  new ArrayList<String>(); 
         			               dataList.add(lguser_id);		  
         			               dataList.add(lguser_name);
         			               dataList.add(szrpt_code);
         			               dataList.add(S_YEAR);
         			               dataList.add(S_MONTH);
         			               dataList.add(filename);
         			               updateDBDataList_update_rpt_dirf.add(dataList);   
         			               System.out.println("update.insert filename="+filename);  		 
         			            }           			            
         			         }            			         
         			     }else{
         			        errMsg += copyOK;
         			     }
         			  }//end of scSrcFile             			 		  
         		   }//end of rptData for
         		   
         		   if(updateDBDataList_insert_rpt_dirf.size() >= 1 ){
         		      sqlCmd = "insert into rpt_dirf values(?,?,?,?,?,?,?,sysdate)"; 
         		      updateDBSqlList = new ArrayList(); 
         		      updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				
         			  updateDBSqlList.add(updateDBDataList_insert_rpt_dirf);//0:sql 1:參數List
					  updateDBList.add(updateDBSqlList);
         		   }
         		   
         		   if(updateDBDataList_rpt_dirf_log.size() >= 1 ){
         		      sqlCmd = " insert into rpt_dirf_log" 
							 + " select rpt_code,m_year,m_month,rpt_fname,rpt_fsize,user_id,user_name,update_date"
							 + ",?,?,sysdate,'U'"
							 + " from rpt_dirf"
							 + " where rpt_code=?"
         			         + " and   m_year = ?"
         			         + " and   m_month = ?"
         			         + " and   rpt_fname = ?";
         			  updateDBSqlList = new ArrayList();       
         		      updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				
         			  updateDBSqlList.add(updateDBDataList_rpt_dirf_log);//0:sql 1:參數List
					  updateDBList.add(updateDBSqlList);
         		   }
         		   
         		   if(updateDBDataList_update_rpt_dirf.size() >= 1 ){
         		      sqlCmd = " update rpt_dirf set user_id=?"
         			         + " ,user_name=?,update_date=sysdate"
         			         + " where rpt_code=?"
         			         + " and   m_year = ?"
         			         + " and   m_month = ?"
         			         + " and   rpt_fname = ?";
         			  updateDBSqlList = new ArrayList();       
         		      updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				
         			  updateDBSqlList.add(updateDBDataList_update_rpt_dirf);//0:sql 1:參數List
					  updateDBList.add(updateDBSqlList);
         		   }
         		   
         		   if(updateDBList.size() >= 1){         		       
                       System.out.println("updateDBList.size()="+updateDBList.size());
                       //for(int k=0;k<updateDBSqlList.size();k++){
                       //    System.out.println("updateDBSqlList.get("+k+")="+(String)updateDBSqlList.get(k));
                       //}               
                       System.out.print("serverRptDir="+Utility.getProperties("serverRptDir"));
    	               System.out.println(":GenerateRptDir="+Utility.getProperties("GenerateRptDir")+System.getProperty("file.separator"));                   
                       File rptDir = new File(Utility.getProperties("GenerateRptDir"));
                       File rptFile = null;
                       if(rptDir.exists() && rptDir.isDirectory()){
		                  fname1= rptDir.list(); //====列出此目錄下的所有檔案===================
		                  for(int c=0;c<fname1.length;c++){
		                      rptFile = new File(Utility.getProperties("GenerateRptDir")+System.getProperty("file.separator")+fname1[c]);
		                      if(!rptFile.isDirectory()) filename_List.add(fname1[c]);
		                  }
		                  MyFTPClient ftpC = new MyFTPClient(rptIP, rptID, rptPwd);   
		                  if(S_YEAR.length() == 2){
		                     S_YEAR = "0"+S_YEAR;
		                  }
		                  putMsg = ftpC.putFiles(Utility.getProperties("serverRptDir"), Utility.getProperties("GenerateRptDir")+System.getProperty("file.separator"),szrpt_code+"/"+S_YEAR+S_MONTH,filename_List);
		                  //putMsg = null;
		                  ftpC=null;
		                  if(putMsg == null){//上傳檔案成功		                     
		                     System.out.print("檔案上傳成功");
		                     for(int i=0;i<filename_List.size();i++){
		                         errMsg +="<br>"+(String)filename_List.get(i);
		                         rptFile = new File(Utility.getProperties("GenerateRptDir")+System.getProperty("file.separator")+(String)filename_List.get(i));
		                         if(rptFile.exists()) rptFile.delete();
		                     }
		                     errMsg +="檔案上傳成功";		  
		                     if(DBManager.updateDB_ps(updateDBList)){//寫入DB
				      	        errMsg += "<br>相關資料寫入資料庫成功";					
			       	         }else{
			 	                errMsg += "<br>相關資料寫入資料庫失敗<br>[DBManager.getErrMsg()]:<br>" + DBManager.getErrMsg();
					         }					 			
					         System.out.println(":errMsg="+errMsg);		         					            		               
                          }else{//end of 上傳檔案成功
                             errMsg += "<br>報表產生完成,但上傳至Sever未成功"+putMsg;    	           	       	
                          }
         	           }//end of rptDir存在
         		   }//end of updateDBSqlList
		        }//end of errMsg = ""		
		}catch(Exception e){
			errMsg += "Generate Report Error";  
			System.out.println("errMsg = Generate Report Error:"+e.getMessage());  
		}	
		 		
	    return errMsg;

    }
    
    private List getFormData(HttpServletRequest request){        
        List rptData = new LinkedList();
    	try{
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
		  	
		  		for ( int i = 0; i < row; i++) {		  	    		  	  			  
					if ( t.get("isModify_" + (i+1)) != null ) {					  
					     rptData.add((String)t.get("isModify_"+(i+1)));
					}										
		  		}	
		  		System.out.println("rptData.size="+rptData.size());		  			
		}catch(Exception e){
		   System.out.println("get Form Data Error:"+e.getMessage());
		   
		}  		
		return rptData;
    }
    private String getFileName(String szrpt_code,String szrpt_output_type){
    	String idxRptCode="";
    	String scSrcFile="";
    	try{
    		 //找出實際產生的報表檔名
         	 filenameLoop:
         	 for (int j = 0; j < rptfilename.length; j++) {
         	       idxRptCode = szrpt_code;       			
				   if (rptfilename[j][0].equals(idxRptCode)){
				   	scSrcFile = rptfilename[j][1];
				   	System.out.print("idxRptCode="+idxRptCode);   
				   	System.out.println(":scSrcFile="+scSrcFile);						
				   	if(szrpt_code.equals("FR002WA_STDET") || szrpt_code.equals("FR023W_ATOT") || szrpt_code.equals("FR024W_STDET") || szrpt_code.equals("FR025W_STOT") 
         		       || szrpt_code.equals("FR026W_DET") || szrpt_code.equals("FR028W_STDET") || szrpt_code.equals("FR029W_STDET") || szrpt_code.equals("FR030W_STDET") 
         		       || szrpt_code.equals("FR027W_STDET"))//95.10.31 add FR027W報表名稱區分.農.漁會 by 2295
         		       { 
         		          if(szrpt_output_type.equals("7")){//漁會
         		             scSrcFile = scSrcFile.substring(0,scSrcFile.indexOf("農會"))+"漁"+scSrcFile.substring(scSrcFile.indexOf("農會")+1,scSrcFile.length());
         		             System.out.println("new.scSrcFile="+scSrcFile);
         		          }
         		       }
         		       
         		       //95.03.15 fix FR003W/FR004W_DET/STOT改實際檔名 
         		       if(szrpt_code.equals("FR003W_DET") || szrpt_code.equals("FR003W_STOT") || szrpt_code.equals("FR004W_DET") || szrpt_code.equals("FR004W_STOT")){         			        
         		          if(szrpt_output_type.equals("7")){//漁會
         		             scSrcFile = scSrcFile.substring(0,scSrcFile.indexOf("農業"))+"漁會"+scSrcFile.substring(scSrcFile.indexOf("農業")+2,scSrcFile.length());
         		             System.out.println("new.scSrcFile="+scSrcFile);
         		          }   
         		       }   
				       break filenameLoop;
				   }//end of idxRptCode 
			 }//end of find rptfilename
	    }catch(Exception e){
	        System.out.println("getFileName error:"+e.getMessage());
	        scSrcFile = e.getMessage();
	    }		 
	    return scSrcFile;
    }
    
    private String CreateRpt(String S_YEAR,String S_MONTH,String BANK_NO,String BANK_NAME,String szrpt_code, String bank_type, String szUnit, String szhsien_id){    	  
        String errMsg = "";
    	try{
    	    String sqlCmd = "";
    	    String s_quarter = "";
    	    errMsg += RptWR098W.createRpt(szrpt_code,S_YEAR,S_MONTH,szUnit,null);
    	    /*if(szrpt_code.equals("WR003W_ATOT")){//全體農(漁)會信用部主要經營指標明細表
         	   errMsg += RptWR003W.createRpt(S_YEAR,S_MONTH,szUnit,null);
         	}else if(szrpt_code.equals("WR004W_ATOT")){//全體農(漁)會信用部主要經營指標總表
         	   errMsg += RptWR004W.createRpt(S_YEAR,S_MONTH,szUnit,null);
         	}else if(szrpt_code.equals("WR005W_ATOT")){//全體農漁會信用部主要經營指標總表
         	   errMsg += RptWR005W.createRpt(S_YEAR,S_MONTH,szUnit,null);						
         	} */  
    	}catch(Exception e){
    	   System.out.println("CreateRpt Error:"+e.getMessage());
    	   errMsg += "CreateRpt Error";
    	}
    	return errMsg;
    }
%>    