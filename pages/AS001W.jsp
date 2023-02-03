<%
//93.12.26 create by 2295
//94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除 by 2295
//94.01.07 fix ba01增加欄位 by 2295
//99.02.08 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
//99.10.15 fix 區分100年/99年總機構基本資料 by 2295
//101.07   fix 增加英文名稱修改欄位 by 2968
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.util.*" %>
<%@include file="./include/Header.include" %>
<%
	String bank_type = Utility.getTrimString(dataMap.get("BANK_TYPE"));
	String tbank_no = Utility.getTrimString(dataMap.get("TBANK_NO"));
	System.out.println(report_no+".act="+act);
	session.setAttribute("nowtbank_no",null);//94.01.05 fix 沒有Bank_List,把所點選的Bank_no清除======
	
	if(!Utility.CheckPermission(request,report_no)){//無權限時,導向到LoginError.jsp    
        rd = application.getRequestDispatcher( LoginErrorPgName );        
    }else{            
    	//set next jsp 	    	
    	if(act.equals("new")){
           rd = application.getRequestDispatcher( EditPgName +"?act=new");                	
        }else if(act.equals("List")){
           rd = application.getRequestDispatcher( ListPgName +"?act=List");                	
        }else if(act.equals("Qry")){    	    	    
    	    List BN01List = getQryResult(bank_type,tbank_no);    	    
    	    request.setAttribute("BN01List",BN01List);
    	    rd = application.getRequestDispatcher( ListPgName +"?act=Qry&bank_type="+bank_type+"&tbank_no="+tbank_no);            	        	        	    	    
    	}else if(act.equals("Edit")){        		
            List<DataObject> BN01 = Utility.getBN01(tbank_no);    	      
            DataObject bean = null;
            request.setAttribute("BN01",BN01);    	       
            if(BN01.size() != 0){   
               bean = (DataObject)BN01.get(0); 	       
               bank_type = (String)bean.getValue("bank_type");
               tbank_no = (String)bean.getValue("bank_no");
            }             
            //設定異動者資訊======================================================================
            Calendar now = Calendar.getInstance();
            String YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
	    	//99.10.15 add 查詢年度100年以前.縣市別不同===============================         
  	    	String wlx01_m_year = (Integer.parseInt(YEAR) < 100)?"99":"100"; 
        	//===================================================================== 
			request.setAttribute("maintainInfo","select * from BN01 WHERE m_year="+wlx01_m_year+" and bank_no='" + tbank_no+"'");								       
			request.setAttribute("getEng","select english from WLX01 WHERE m_year="+wlx01_m_year+" and bank_no='" + tbank_no+"'");
        	//=======================================================================================================================		   	    
    	 	rd = application.getRequestDispatcher( EditPgName +"?act=Edit&bank_type="+bank_type+"&tbank_no="+tbank_no);        
    	}else if(act.equals("Insert")){
    	    actMsg = InsertDB(request,lguser_id,lguser_name);    	    	        	    
        	rd = application.getRequestDispatcher( nextPgName );         
        }else if(act.equals("Update")){
    	    actMsg = UpdateDB(request,lguser_id,lguser_name);    	    	        	    
        	rd = application.getRequestDispatcher( nextPgName );         
        }else if(act.equals("Delete")){
    	    actMsg = DeleteDB(request,lguser_id,lguser_name);    	    	        	    
        	rd = application.getRequestDispatcher( nextPgName );                 
        }
    	request.setAttribute("actMsg",actMsg);    
    }        
     
%>

<%@include file="./include/Tail.include" %>

<%!
    private final static String report_no = "AS001W";
    private final static String nextPgName = "/pages/ActMsg.jsp";    
    private final static String EditPgName = "/pages/"+report_no+"_Edit.jsp";    
    private final static String ListPgName = "/pages/"+report_no+"_List.jsp";        
    private final static String LoginErrorPgName = "/pages/LoginError.jsp";
    
    //取得查詢結果
    private List getQryResult(String bank_type,String bank_no){    		
    		List<String> paramList = new ArrayList<String>();
    		StringBuffer sql = new StringBuffer();
    		Calendar now = Calendar.getInstance();
         	String YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
    		//99.10.14 add 查詢年度100年以前.縣市別不同===============================
	        String cd01_table = (Integer.parseInt(YEAR) < 100)?"cd01_99":""; 
	      	String wlx01_m_year = (Integer.parseInt(YEAR) < 100)?"99":"100"; 
	        //===================================================================== 
	      
    		sql.append(" select cdshareno.cmuse_id,cdshareno.cmuse_name,bn01.bank_no,bn01.bank_name,bn01.update_date ");
			sql.append(" from (select * from bn01 where m_year=?)bn01 LEFT JOIN cdshareno on bank_type = cmuse_id and cmuse_div='001'");					
			paramList.add(wlx01_m_year);
			
			if(!bank_no.equals("")){//若有總機構代碼.以總機構代碼為主
			    sql.append(" where bn01.bank_no like ?");
			    paramList.add(bank_no+"%");		
			}else{	   
			  if(!bank_type.equals("Z"))//不為全部類別
			    sql.append(" where bank_type=?");
			    paramList.add(bank_type);		
			}
			
			sql.append(" order by bank_type,bank_no ");
		
			List<DataObject> dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"update_date");
			
            return dbData;
    }
    
	public String InsertDB(HttpServletRequest request,String lguser_id,String lguser_name) throws Exception{
		String errMsg="";	
			
		String bank_no=" ";
		String bank_name=" ";
		String bn_type="1";
		String bank_type="";
		String add_user=lguser_id;
		String add_name=lguser_name;				
		String bank_b_name="";		
		String kind_1="";
		String kind_2="";
		String bn_type2="";
		String exchange_no="";		
		String user_id=lguser_id;
	    String user_name=lguser_name;
	    //99.02.06 add 		
		List<String> paramList = new ArrayList<String>();
		StringBuffer sql = new StringBuffer();		
		List<List> updateDBList = new ArrayList<List>();//0:sql 1:data		
		List updateDBSqlList = new ArrayList();
		List<List> updateDBDataList = new ArrayList<List>();//儲存參數的List
		List<String> dataList =  new ArrayList<String>();//儲存參數的data
		 
		try {
			    Calendar now = Calendar.getInstance();
			    String YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
    		    //99.10.14 add 查詢年度100年以前.縣市別不同===============================
	            String cd01_table = (Integer.parseInt(YEAR) < 100)?"cd01_99":""; 
	      	    String wlx01_m_year = (Integer.parseInt(YEAR) < 100)?"99":"100"; 
	            //===================================================================== 
	      
			    Map dataMap =Utility.saveSearchParameter(request);
			    if(dataMap != null){
			       bank_no = Utility.getTrimString(dataMap.get("BANK_NO"));
			       bank_name = Utility.getTrimString(dataMap.get("BANK_NAME"));
			       bank_type = Utility.getTrimString(dataMap.get("BANK_TYPE"));
			       bank_b_name = Utility.getTrimString(dataMap.get("BANK_B_NAME"));
			       exchange_no = Utility.getTrimString(dataMap.get("EXCHANGE_NO"));
			    }
		
				sql.append(" SELECT * FROM BA01 WHERE bank_no= ? ");		
				paramList.add(bank_no);		
		
				List<DataObject> dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"");
				System.out.println("BA01.size="+dbData.size());
				
				if (dbData.size() != 0){
				    errMsg = errMsg + "此筆總機構資料已存在無法新增<br>";
				}else{
				    //寫入BN01==========================================================================
					sql.delete(0, sql.length());
					paramList.clear();
				    sql.append(" INSERT INTO BN01 VALUES (?,?,?,?,?,?,sysdate,?,?,?,?,?,?,?,sysdate,?)");
				    
				    updateDBSqlList.add(sql.toString());//0:欲執行的sql				    
					
					dataList = new ArrayList<String>();//傳內的參數List		 	           				   
					dataList.add(bank_no); 
					dataList.add(bank_name); 
					dataList.add(bn_type); 
					dataList.add(bank_type); 
					dataList.add(add_user); 
					dataList.add(add_name); 
					dataList.add(bank_b_name); 
					dataList.add(kind_1); 
					dataList.add(kind_2); 
					dataList.add(bn_type2); 
					dataList.add(exchange_no); 
					dataList.add(user_id); 
					dataList.add(user_name); 
					dataList.add(wlx01_m_year);
					updateDBDataList.add(dataList);//1:傳內的參數List
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
					
					//寫入BA01====================================================================================
					updateDBSqlList = new ArrayList();		
					updateDBDataList = new ArrayList<List>();//儲存參數的List						
					sql.delete(0, sql.length());
					sql.append(" INSERT INTO BA01 VALUES (?,?,?,'0',?,' ',null,?,?,sysdate,?)");
							
					updateDBSqlList.add(sql.toString());//0:欲執行的sql		
					
					dataList = new ArrayList<String>();//傳內的參數List
					dataList.add(bank_no); 
					dataList.add(bank_name); 					
					dataList.add(bank_type); 
					dataList.add(bank_no); 					
					dataList.add(user_id); 
					dataList.add(user_name); 
					dataList.add(wlx01_m_year);
					updateDBDataList.add(dataList);//1:傳內的參數List
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					
    				updateDBList.add(updateDBSqlList);
    				
				    if(DBManager.updateDB_ps(updateDBList)){					 
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
	
	
	private String UpdateDB(HttpServletRequest request,String lguser_id,String lguser_name) throws Exception{    	
		String errMsg="";		
		String bank_no="";
		String bank_name="";
		String bn_type="1";
		String bank_type="";
		String add_user=lguser_id;
		String add_name=lguser_name;				
		String bank_b_name="";		
		String kind_1="";
		String kind_2="";
		String bn_type2="";
		String exchange_no="";		
		String user_id=lguser_id;
	    String user_name=lguser_name;	 
	    String english="";
	    //99.02.06 add 		
		List<String> paramList = new ArrayList<String>();
		StringBuffer sql = new StringBuffer();		
		List<List> updateDBList = new ArrayList<List>();//0:sql 1:data		
		List updateDBSqlList = new ArrayList();
		List<List> updateDBDataList = new ArrayList<List>();//儲存參數的List
		List<String> dataList =  new ArrayList<String>();//儲存參數的data
	    
		
		try {
				Map dataMap =Utility.saveSearchParameter(request);
			    if(dataMap != null){
			       bank_no = Utility.getTrimString(dataMap.get("BANK_NO"));
			       bank_name = Utility.getTrimString(dataMap.get("BANK_NAME"));
			       bank_type = Utility.getTrimString(dataMap.get("BANK_TYPE"));
			       bank_b_name = Utility.getTrimString(dataMap.get("BANK_B_NAME"));
			       exchange_no = Utility.getTrimString(dataMap.get("EXCHANGE_NO"));
			       english = Utility.getTrimString(dataMap.get("ENGLISH"));
			    }
				Calendar now = Calendar.getInstance();
				String YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
    		    //99.10.14 add 查詢年度100年以前.縣市別不同===============================
	            String cd01_table = (Integer.parseInt(YEAR) < 100)?"cd01_99":""; 
	      	    String wlx01_m_year = (Integer.parseInt(YEAR) < 100)?"99":"100"; 
	            //===================================================================== 
	      
				sql.append(" SELECT * FROM (select * from BA01 where m_year=?)BA01 WHERE bank_no= ? ");		
				paramList.add(wlx01_m_year);
				paramList.add(bank_no);		
		
				List<DataObject> dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"");
				System.out.println("BA01.size="+dbData.size());
				
				if (dbData.size() == 0){
				    errMsg = errMsg + "此筆資料不存在無法修改<br>";
				}else{    				   
				    //insert BN04_LOG===================================================	
				    sql.delete(0, sql.length());				    
					sql.append("INSERT INTO BN04_LOG select bank_no,bank_no,bank_name,bn_type,bank_type,");
					sql.append("add_user,add_name,add_date,bank_b_name,kind_1,kind_2,bn_type2,exchange_no,");
					sql.append("user_id,user_name,update_date,?,?,sysdate,'0','U' from BN01 WHERE m_year=? and bank_no=?"); 					    
				    updateDBSqlList.add(sql.toString());//0:欲執行的sql	
				    
					dataList = new ArrayList<String>();//傳內的參數List		 	           				   
					dataList.add(user_id); 
					dataList.add(user_name);
					dataList.add(wlx01_m_year); 
					dataList.add(bank_no); 
					
					updateDBDataList.add(dataList);//1:傳內的參數List
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					updateDBList.add(updateDBSqlList);
				    //UPDATE BN01=========================================================================
				    updateDBSqlList = new ArrayList();		
					updateDBDataList = new ArrayList<List>();//儲存參數的List						
					sql.delete(0, sql.length());
					
				    sql.append("UPDATE BN01 SET bank_name=?,bank_b_name=?,exchange_no=?,user_id=?,user_name=?");
					sql.append(",update_date=sysdate where bank_no=? and m_year=?");									    	   
					updateDBSqlList.add(sql.toString());//0:欲執行的sql	
					
					dataList = new ArrayList<String>();//傳內的參數List
					dataList.add(bank_name); 
					dataList.add(bank_b_name); 					
					dataList.add(exchange_no); 									
					dataList.add(user_id); 
					dataList.add(user_name); 
					dataList.add(bank_no); 	
					dataList.add(wlx01_m_year);
					updateDBDataList.add(dataList);//1:傳內的參數List
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
    				updateDBList.add(updateDBSqlList);	     				  
		              
		            //UPDATE BA01====================================================================================  		            
				    updateDBSqlList = new ArrayList();		
					updateDBDataList = new ArrayList<List>();//儲存參數的List						
					sql.delete(0, sql.length());
					sql.append("UPDATE BA01 SET bank_name=? where bank_no=? and m_year=?");					    			    	   
					updateDBSqlList.add(sql.toString());//0:欲執行的sql	
					
					dataList = new ArrayList<String>();//傳內的參數List
					dataList.add(bank_name); 
					dataList.add(bank_no); 	
					dataList.add(wlx01_m_year);	        	
	            	updateDBDataList.add(dataList);//1:傳內的參數List
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
    				updateDBList.add(updateDBSqlList);
					
					
					 //102.02.05 add 寫入WLX01_LOG ===================================================================
					updateDBSqlList = new ArrayList();		
					updateDBDataList = new ArrayList<List>();//儲存參數的List						
					sql.delete(0, sql.length());
                    sql.append(" INSERT INTO WLX01_LOG "
                           + " select BANK_NO,ENGLISH,SETUP_APPROVAL_UNT,SETUP_DATE,SETUP_NO,CHG_LICENSE_DATE,CHG_LICENSE_NO,"
                           + " 		  CHG_LICENSE_REASON,START_DATE,BUSINESS_ID,HSIEN_ID,AREA_ID,ADDR,TELNO,FAX,EMAIL,WEB_SITE,"
                           + " 		  CENTER_FLAG,CENTER_NO,STAFF_NUM,IT_HSIEN_ID,IT_AREA_ID,IT_ADDR,IT_NAME,IT_TELNO,"
                           + " 		  AUDIT_HSIEN_ID,AUDIT_AREA_ID,AUDIT_ADDR,AUDIT_NAME,AUDIT_TELNO,FLAG,OPEN_DATE,M2_NAME,"
                           + " 		  HSIEN_DIV_1,CANCEL_NO,CANCEL_DATE,USER_ID,USER_NAME,UPDATE_DATE,CREDIT_STAFF_NUM,"
                           + "        ?,?,sysdate,'U'"
						   + " from WLX01"
						   + " where bank_no=? and m_year=?");
				    updateDBSqlList.add(sql.toString());//0:欲執行的sql	   
					dataList = new ArrayList<String>();//傳內的參數List		 	           				   
					dataList.add(user_id); 
					dataList.add(user_name);
					dataList.add(bank_no);
					dataList.add(wlx01_m_year); 
					updateDBDataList.add(dataList);//1:傳內的參數List
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
    				updateDBList.add(updateDBSqlList);						                 
					
    				//UPDATE WLX01====================================================================================  		            
				    updateDBSqlList = new ArrayList();		
					updateDBDataList = new ArrayList<List>();//儲存參數的List						
					sql.delete(0, sql.length());
					sql.append("UPDATE WLX01 SET english=? WHERE bank_no=? ");					    			    	   
					updateDBSqlList.add(sql.toString());//0:欲執行的sql	
					
					dataList = new ArrayList<String>();//傳內的參數List
					dataList.add(english);
					dataList.add(bank_no); 	
	            	updateDBDataList.add(dataList);//1:傳內的參數List
					updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
    				updateDBList.add(updateDBSqlList);
						   
					if(DBManager.updateDB_ps(updateDBList)){					 
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
    
    private String DeleteDB(HttpServletRequest request,String lguser_id,String lguser_name) throws Exception{
		String errMsg="";		
		String bank_no=" ";
		String user_id=lguser_id;
	    String user_name=lguser_name;
	    //99.02.06 add 		
		List<String> paramList = new ArrayList<String>();
		StringBuffer sql = new StringBuffer();		
		List<List> updateDBList = new ArrayList<List>();//0:sql 1:data		
		List updateDBSqlList = new ArrayList();
		List<List> updateDBDataList = new ArrayList<List>();//儲存參數的List
		List<String> dataList =  new ArrayList<String>();//儲存參數的data
		
		try {
				Map dataMap =Utility.saveSearchParameter(request);
			    if(dataMap != null){
			       bank_no = Utility.getTrimString(dataMap.get("BANK_NO"));			      
			    }
				Calendar now = Calendar.getInstance();
				String YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
    		    //99.10.15 add 查詢年度100年以前.縣市別不同===============================
	            String cd01_table = (Integer.parseInt(YEAR) < 100)?"cd01_99":""; 
	      	    String wlx01_m_year = (Integer.parseInt(YEAR) < 100)?"99":"100"; 
	            //===================================================================== 
	      
				//檢查是否有建立分支機構
				sql.append("SELECT * FROM BA01 WHERE pbank_no=? and bank_kind='1' and m_year=?");		
				paramList.add(bank_no);		
				paramList.add(wlx01_m_year);
		
				List<DataObject> dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"");
				System.out.println("BA01.size="+dbData.size());
				
				if(dbData.size () != 0 ){
				   errMsg = errMsg + "已有分支機構,不可刪除<br>";
				}else{
				   sql.delete(0, sql.length());
				   sql.append("SELECT * FROM BA01 WHERE bank_no= ? and m_year=?");	
				   dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"");
				   System.out.println("BA01.size="+dbData.size());
				   
				   if (dbData.size() == 0){
				      errMsg = errMsg + "此筆資料不存在無法刪除<br>";
				   }else{    				   
				      //insert BN04_LOG===================================================		    				    				      
				      sql.delete(0, sql.length());				    
					  sql.append("INSERT INTO BN04_LOG select bank_no,bank_no,bank_name,bn_type,bank_type,");
					  sql.append("add_user,add_name,add_date,bank_b_name,kind_1,kind_2,bn_type2,exchange_no,");
					  sql.append("user_id,user_name,update_date,?,?,sysdate,'0','D' from BN01 WHERE m_year=? and bank_no=?"); 					    
				      updateDBSqlList.add(sql.toString());//0:欲執行的sql	
				      
					  dataList = new ArrayList<String>();//傳內的參數List		 	           				   
					  dataList.add(user_id); 
					  dataList.add(user_name); 
					  dataList.add(wlx01_m_year);
					  dataList.add(bank_no); 					 
					  updateDBDataList.add(dataList);//1:傳內的參數List
					  updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
					  updateDBList.add(updateDBSqlList);
				      //DELETE BN01=========================================================================
				      updateDBSqlList = new ArrayList();		
					  updateDBDataList = new ArrayList<List>();//儲存參數的List						
					  sql.delete(0, sql.length());
					  sql.append("DELETE from BN01 where bank_no = ? and m_year=?");					    			    	   
					  updateDBSqlList.add(sql.toString());//0:欲執行的sql	
					 
					  dataList = new ArrayList<String>();//傳內的參數List					 
					  dataList.add(bank_no); 
					  dataList.add(wlx01_m_year);		        	
	            	  updateDBDataList.add(dataList);//1:傳內的參數List
					  updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
    				  updateDBList.add(updateDBSqlList);
    				  //DELETE BA01=========================================================================
				      updateDBSqlList = new ArrayList();		
					  updateDBDataList = new ArrayList<List>();//儲存參數的List						
					  sql.delete(0, sql.length());
					  sql.append("DELETE from BA01 where bank_no = ? and m_year=?");					    			    	   
					  updateDBSqlList.add(sql.toString());//0:欲執行的sql	
					 
					  dataList = new ArrayList<String>();//傳內的參數List					 
					  dataList.add(bank_no); 	
					  dataList.add(wlx01_m_year);	        	
	            	  updateDBDataList.add(dataList);//1:傳內的參數List
					  updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
    				  updateDBList.add(updateDBSqlList);
    				 	
					  if(DBManager.updateDB_ps(updateDBList)){				 
						errMsg = errMsg + "相關資料寫入資料庫成功";					
					  }else{
				   		errMsg = errMsg + "相關資料寫入資料庫失敗";
					  }
    	   		   }
    	   		}//end of 無分支機構
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";						
		}	

		return errMsg;
	}
%>    