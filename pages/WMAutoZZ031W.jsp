<%
// 95.02.10 add 排程鎖定 by 2295
//100.07.21 fix 套用DAO.preparestatment,並列印轉換後的SQL by 2295
//          fix 區分100年/99年總機構基本資料 by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="com.tradevan.util.DownLoad" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="java.io.*" %>
<%@ page import="java.util.*" %>
<%@ page import="java.lang.Integer" %>
<%@ page import="java.text.SimpleDateFormat" %>
<%
   String report_no = ( request.getParameter("report_no")==null ) ? "" : (String)request.getParameter("report_no");		
   String lguser_id = ( request.getParameter("lguser_id")==null ) ? "" : (String)request.getParameter("lguser_id");		
   String bank_type = ( request.getParameter("bank_type")==null ) ? "" : (String)request.getParameter("bank_type");		
   String lock_status = ( request.getParameter("lock_status")==null ) ? "" : (String)request.getParameter("lock_status");		
   System.out.println("=============執行鎖定開始===========");
   String errMsg = UpdateDB(report_no, lguser_id,lguser_id,bank_type,lock_status);
   System.out.println("errMsg = "+errMsg);
   /*
   List dbData = DBManager.QueryDB("select * from cdshareno where cmuse_Div='001' and cmuse_id <> 'Z'","");   
   if(dbData != null && dbData.size() != 0){
	  for(int i=0;i<dbData.size();i++){
	       System.out.println("請至"+AutoFileCheck.FileCheck(report_no,(String)((DataObject)dbData.get(i)).getValue("cmuse_id"))+"查看檢核結果");
	  }
	}
	*/
   System.out.println("=============執行鎖定結束===========");
%>
<%!
public String UpdateDB(String Report_no,String lguser_id,String lguser_name,String bank_type,String lock_status) throws Exception{    	
		File logfile;
		FileOutputStream logos=null;    	
		BufferedOutputStream logbos = null;
		PrintStream logps = null;
		Date nowlog = new Date();
		SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");	     
		SimpleDateFormat logfileformat = new SimpleDateFormat("yyyyMMddHHmmss");
	    Calendar logcalendar;
	    File logDir = null;
	    
		String sqlCmd = "";		
		String errMsg="";
		String user_id=lguser_id;
	    String user_name=lguser_name;		
	    String bank_code="";
	    Calendar now = Calendar.getInstance();
	    String S_YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
   	    String S_MONTH = String.valueOf(now.get(Calendar.MONTH)+1);   //月份以0開始故加1取得實際月份;
   	    String lock_status_msg = lock_status.equals("Y")?"鎖定":"解除鎖定";
   	    		
	    List paramList = new ArrayList();
	    String cd01_table = "";
        String wlx01_m_year = "";
        
   	    if(S_MONTH.equals("1")){//若本月為1月份是..則是申報上個年度的12月份
           S_YEAR = String.valueOf(Integer.parseInt(S_YEAR) - 1);
           S_MONTH = "12";
        }else{    
           S_MONTH = String.valueOf(Integer.parseInt(S_MONTH) - 1);//申報上個月份的
        }
        
        //100.07.20 add 查詢年度100年以前.縣市別不同===============================
  	    cd01_table = (Integer.parseInt(S_YEAR) < 100)?"cd01_99":""; 
  	    wlx01_m_year = (Integer.parseInt(S_YEAR) < 100)?"99":"100"; 
  	    //=====================================================================    
        
        //List updateDBSqlList = new LinkedList();
        List data = null;	
        
        List updateDBList = new ArrayList();//0:sql 1:data
	    List updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
		List updateDBDataList = new ArrayList();//儲存參數的List		    
		List dataList =  new ArrayList();//儲存參數的data
		
		String sqlCmd_WML01_Log="";
		String sqlCmd_Delete_WML01_Lock="";
		String sqlCmd_Update_WML01_Lock="";
		String sqlCmd_Insert_WML01_Lock="";
		String sqlCmd_Update_WML01_null="";
		String sqlCmd_Update_WML01="";
		
		List updateDBDataList_WML01_Log = new ArrayList();//儲存參數的List	
		List updateDBDataList_Delete_WML01_Lock = new ArrayList();//儲存參數的List
		List updateDBDataList_Update_WML01_Lock = new ArrayList();//儲存參數的List
		List updateDBDataList_Insert_WML01_Lock = new ArrayList();//儲存參數的List
		List updateDBDataList_Update_WML01_null = new ArrayList();//儲存參數的List
		List updateDBDataList_Update_WML01 = new ArrayList();//儲存參數的List
		try {			    			    
			    logDir  = new File(Utility.getProperties("logDir"));
	            if(!logDir.exists()){
     			    if(!Utility.mkdirs(Utility.getProperties("logDir"))){
     				   System.out.println("目錄新增失敗");
     			    }    
    		    }
			    logfile = new File(logDir + System.getProperty("file.separator") + Report_no +"_ZZ031W."+ logfileformat.format(nowlog));						 
			    System.out.println("logfile filename="+logDir + System.getProperty("file.separator") + Report_no +"_ZZ031W."+ logfileformat.format(nowlog));
			    logos = new FileOutputStream(logfile,true);  		        	   
			    logbos = new BufferedOutputStream(logos);
			    logps = new PrintStream(logbos);			
			    List paramList1 = new ArrayList();
			    sqlCmd = " select muser_name from wtt01 where muser_id=?";
			    paramList1.add(user_id);
			    data = DBManager.QueryDB_SQLParam(sqlCmd,paramList1,""); 
			    if(data != null && data.size() != 0){
			       user_name = (String)((DataObject)data.get(0)).getValue("muser_name");
			    }
			    logcalendar = Calendar.getInstance(); 
			    nowlog = logcalendar.getTime();			    	
			    logps.println(logformat.format(nowlog)+"執行"+lock_status_msg+" 帳號 "+":"+user_id+user_name);		    		
			    
			    logps.flush();
			    sqlCmd = " select distinct a.bank_no, a.bank_name, b.m_year, b.m_month, b.upd_code, b.input_method, "
					   + " b.wml01_lock_status as lock_status, b.common_center, b.upd_method  , b.wml01_lock_lock_status as wml01_lock_status, "
					   + " decode(nvl(b.input_method,'N'),'N','N','Y') as havingdata  "
					   + " from (select * from ba01 where m_year=?)a left join wml01_a_v b on a.bank_no= b.bank_code and "
					   + "                       b.m_year=? and  b.m_month= ? and  b.report_no=? "
					   + " where a.bank_type in ('6','7','8') and a.bank_kind='0' "
					   + " order by a.bank_no,b.m_year,b.m_month ";
			    paramList.add(wlx01_m_year);	
				paramList.add(S_YEAR);
				paramList.add(S_MONTH);
				paramList.add(Report_no);							    				   			    
			    List lockData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"m_year,m_month");            
			    
			    
			    //已申報							 
			    sqlCmd_WML01_Log = "INSERT INTO WML01_LOG " 
			        	         + " select m_year,m_month,bank_code,report_no,input_method,add_user,add_name,add_date,common_center,upd_method,upd_code "
			        	         + ",batch_no,lock_status,user_id,user_name,update_date,?,?,sysdate,'U'"
			        	         + " FROM WML01 WHERE m_year= ? AND m_month=? AND bank_code=? AND report_no=? ";
			    //解除鎖定
			    sqlCmd_Delete_WML01_Lock = "DELETE WML01_LOCK WHERE m_year=? AND m_month=? AND bank_code=? AND report_no=? "; 
			    
				//鎖定			    
			    sqlCmd_Update_WML01_Lock = "UPDATE WML01_LOCK SET lock_status =? WHERE m_year=? AND m_month=? AND bank_code=? AND report_no=? ";
			    
			    //鎖定
			    sqlCmd_Insert_WML01_Lock = "Insert into WML01_LOCK VALUES(?,?,?,?,?,?,?,sysdate)";
			    
			    //解除鎖定(農會.漁會.農業信用保証基金寫到WML01)
			    sqlCmd_Update_WML01_null = " UPDATE WML01 SET lock_status = null ,user_id=?,user_name=?,update_date=sysdate "
						            + " WHERE m_year=? AND m_month=? AND bank_code=? AND report_no=?";
				sqlCmd_Update_WML01 = " UPDATE WML01 SET lock_status =? ,user_id=?,user_name=?,update_date=sysdate "
						        	+ " WHERE m_year=? AND m_month=? AND bank_code=? AND report_no=? ";
										 
			    
			    for(int i=0;i<lockData.size();i++){					       			                 			
			        bank_code = (String)((DataObject)lockData.get(i)).getValue("bank_no");
     			    System.out.println("bank_code = '"+bank_code+"'");
			    	logcalendar = Calendar.getInstance(); 
			    	nowlog = logcalendar.getTime();			    	
			    	logps.println(logformat.format(nowlog)+" "+"機構代號:"+bank_code+" 申報年月:"+S_YEAR+"/"+S_MONTH);		    		
			    	logps.flush();
     			    
     			    paramList = new ArrayList();
     			    sqlCmd = "select * from wml01 WHERE m_year=? AND m_month=? AND bank_code=? AND report_no=? ";
					paramList.add(S_YEAR);
					paramList.add(S_MONTH);
					paramList.add(bank_code);
					paramList.add(Report_no);													
																	 
					data = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");		 			    
				    System.out.println("update.size="+data.size());				    
				    
					if(data.size() != 0){//已申報							 
			        	//sqlCmd_WML01_Log = "INSERT INTO WML01_LOG " 
			        	//                 + " select m_year,m_month,bank_code,report_no,input_method,add_user,add_name,add_date,common_center,upd_method,upd_code "
			        	//                 + ",batch_no,lock_status,user_id,user_name,update_date,?,?,sysdate,'U'"
			        	//                 + " FROM WML01 WHERE m_year= ? AND m_month=? AND bank_code=? AND report_no=? ";
						dataList = new ArrayList();   
						dataList.add(lguser_id);
						dataList.add(lguser_name);
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(bank_code);
						dataList.add(Report_no);
						updateDBDataList_WML01_Log.add(dataList);						
												
						if(bank_type.equals("1")/*95.02.09 add 全國農業金庫*/ || bank_type.equals("2") || bank_type.equals("8")){//農金局.共用中心寫到WML01_LOCK
						   paramList = new ArrayList();
						   sqlCmd = " select * from WML01_LOCK WHERE m_year=? AND m_month=? AND bank_code=? AND report_no=? ";   
						   paramList.add(S_YEAR);
						   paramList.add(S_MONTH);
						   paramList.add(bank_code);
						   paramList.add(Report_no); 	  
						   data = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");		 			    
				           System.out.println("updateWML01_LOCK.size="+data.size());				    	 	  
				           if(data.size() != 0){
				              if(lock_status.equals("N")){//解除鎖定,將WML01
						         //sqlCmd_Update_WML01_Lock = "DELETE WML01_LOCK WHERE m_year=? AND m_month=? AND bank_code=? AND report_no=? ";
							 	 dataList = new ArrayList();            
							 	 dataList.add(S_YEAR);
							 	 dataList.add(S_MONTH);
							 	 dataList.add(bank_code);
							 	 dataList.add(Report_no);
							 	 updateDBDataList_Delete_WML01_Lock.add(dataList);
						      }else if(lock_status.equals("Y")){//鎖定
						         if(bank_type.equals("8")){
						            lock_status = "C";//由共用中心做鎖定的
						         }
						         //sqlCmd_Update_WML01_Lock = "UPDATE WML01_LOCK SET lock_status =? WHERE m_year=? AND m_month=? AND bank_code=? AND report_no=? ";
							     dataList = new ArrayList(); 
							     dataList.add(lock_status);	 
							     dataList.add(S_YEAR);
							     dataList.add(S_MONTH);
							     dataList.add(bank_code);
							     dataList.add(Report_no);
							     updateDBDataList_Update_WML01_Lock.add(dataList);
						      }  						      
				           }else{				              
				              if(lock_status.equals("Y")){//鎖定				              
				                 if(bank_type.equals("8")){
						            lock_status = "C";//由共用中心做鎖定的						      
						         }						      
						         //sqlCmd_Insert_WML01_Lock = "Insert into WML01_LOCK VALUES(?,?,?,?,?,?,?,sysdate)";
							     dataList = new ArrayList(); 
							     dataList.add(S_YEAR);	    
							     dataList.add(S_MONTH);
							     dataList.add(bank_code);
							     dataList.add(Report_no);
							     dataList.add(lock_status);
							     dataList.add(lguser_id);
							     dataList.add(lguser_name);
							     updateDBDataList_Insert_WML01_Lock.add(dataList);
							     lock_status="Y";//94.03.07 fix若為共用中心鎖定時,恢復為"Y"
						      }						      						      
				           }
				           //updateDBSqlList.add(sqlCmd);	
						}//end of bank_type in ('2','8')
						if(bank_type.equals("6") || bank_type.equals("7") || bank_type.equals("4") ){//農會.漁會.農業信用保証基金寫到WML01
						   if(lock_status.equals("N")){//解除鎖定
						      //sqlCmd_Update_WML01_null = " UPDATE WML01 SET lock_status = null ,user_id=?,user_name=?,update_date=sysdate "
						      //                    + " WHERE m_year=? AND m_month=? AND bank_code=? AND report_no=?";
							  dataList = new ArrayList(); 	     
							  dataList.add(lguser_id);
							  dataList.add(lguser_name);
							  dataList.add(S_YEAR);
							  dataList.add(S_MONTH);
							  dataList.add(bank_code);
							  dataList.add(Report_no);
							  updateDBDataList_Update_WML01_null.add(dataList);
						   }else if(lock_status.equals("Y")){//鎖定
						      //sqlCmd_Update_WML01 = " UPDATE WML01 SET lock_status =? ,user_id=?,user_name=?,update_date=sysdate "
						      //					  + " WHERE m_year=? AND m_month=? AND bank_code=? AND report_no=? ";
							  dataList = new ArrayList(); 	   	 
							  dataList.add("Y");   	 
							  dataList.add(lguser_id);
							  dataList.add(lguser_name);
							  dataList.add(S_YEAR);
							  dataList.add(S_MONTH);
							  dataList.add(bank_code);
							  dataList.add(Report_no);
							  updateDBDataList_Update_WML01.add(dataList);
						   }	    	  
						   //updateDBSqlList.add(sqlCmd);	
						}//end of bank_type in ('6','7','4')						
					}else{//未申報
					   sqlCmd = " select * from WML01_LOCK WHERE m_year=? AND m_month=? AND bank_code=? AND report_no=? ";  
					   paramList = new ArrayList();
					   paramList.add(S_YEAR);	 	 	   
					   paramList.add(S_MONTH);
					   paramList.add(bank_code);
					   paramList.add(Report_no);
					   
					   data = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");		 			    
				       System.out.println("updateWML01_LOCK.size="+data.size());				    	 	  
				       if(data.size() != 0){
				          if(lock_status.equals("N")){//解除鎖定,將WML01
				            //sqlCmd_Delete_WML01_Lock = "DELETE WML01_LOCK WHERE m_year=? AND m_month=? AND bank_code=? AND report_no=? "; 		  
						
						    dataList = new ArrayList();            
							dataList.add(S_YEAR);
							dataList.add(S_MONTH);
							dataList.add(bank_code);
							dataList.add(Report_no);
							updateDBDataList_Delete_WML01_Lock.add(dataList);   	        
						      	        
						  }else if(lock_status.equals("Y")){//鎖定						         
						    //sqlCmd_Update_WML01_Lock = "UPDATE WML01_LOCK SET lock_status =? WHERE m_year=? AND m_month=? AND bank_code=? AND report_no=? ";
							dataList = new ArrayList(); 
							dataList.add(lock_status);	 
							dataList.add(S_YEAR);
							dataList.add(S_MONTH);
							dataList.add(bank_code);
							dataList.add(Report_no);
							updateDBDataList_Update_WML01_Lock.add(dataList);     	 
						  }  						      
				       }else{
				          if(lock_status.equals("Y")){//鎖定				                 
						     //sqlCmd_Insert_WML01_Lock = "Insert into WML01_LOCK VALUES(?,?,?,?,?,?,?,sysdate)";
							 dataList = new ArrayList(); 
							 dataList.add(S_YEAR);	    
							 dataList.add(S_MONTH);
							 dataList.add(bank_code);
							 dataList.add(Report_no);
							 dataList.add(lock_status);
							 dataList.add(lguser_id);
							 dataList.add(lguser_name);
							 updateDBDataList_Insert_WML01_Lock.add(dataList);						          	    
						  }						      
				       }
				       //updateDBSqlList.add(sqlCmd);						   
					}	      				    
	            }//end of for		
	            	
	            	
	    	    
				  
			     
			   
	            	
	            if(updateDBDataList_WML01_Log.size() >= 1 ){
	               //已申報							 
			       sqlCmd_WML01_Log = "INSERT INTO WML01_LOG " 
			        	            + " select m_year,m_month,bank_code,report_no,input_method,add_user,add_name,add_date,common_center,upd_method,upd_code "
			        	            + ",batch_no,lock_status,user_id,user_name,update_date,?,?,sysdate,'U'"
			        	            + " FROM WML01 WHERE m_year= ? AND m_month=? AND bank_code=? AND report_no=? ";
	
         		   updateDBSqlList = new ArrayList(); 
         		   updateDBSqlList.add(sqlCmd_WML01_Log);//0:欲執行的sql				
         		   updateDBSqlList.add(updateDBDataList_WML01_Log);//0:sql 1:參數List
				   updateDBList.add(updateDBSqlList);
	            }
	            
	            if(updateDBDataList_Delete_WML01_Lock.size() >= 1 ){
			       //解除鎖定
			       sqlCmd_Delete_WML01_Lock = "DELETE WML01_LOCK WHERE m_year=? AND m_month=? AND bank_code=? AND report_no=? "; 
		
         		   updateDBSqlList = new ArrayList(); 
         		   updateDBSqlList.add(sqlCmd_Delete_WML01_Lock);//0:欲執行的sql				
         		   updateDBSqlList.add(updateDBDataList_Delete_WML01_Lock);//0:sql 1:參數List
				   updateDBList.add(updateDBSqlList);
	            }
	            
	            if(updateDBDataList_Update_WML01_Lock.size() >= 1 ){
			       //鎖定			    
			       sqlCmd_Update_WML01_Lock = "UPDATE WML01_LOCK SET lock_status =? WHERE m_year=? AND m_month=? AND bank_code=? AND report_no=? ";
			  
         		   updateDBSqlList = new ArrayList(); 
         		   updateDBSqlList.add(sqlCmd_Update_WML01_Lock);//0:欲執行的sql				
         		   updateDBSqlList.add(updateDBDataList_Update_WML01_Lock);//0:sql 1:參數List
				   updateDBList.add(updateDBSqlList);
	            }
		
				if(updateDBDataList_Insert_WML01_Lock.size() >= 1 ){
			       //鎖定
			       sqlCmd_Insert_WML01_Lock = "Insert into WML01_LOCK VALUES(?,?,?,?,?,?,?,sysdate)";
			   
         		   updateDBSqlList = new ArrayList(); 
         		   updateDBSqlList.add(sqlCmd_Insert_WML01_Lock);//0:欲執行的sql				
         		   updateDBSqlList.add(updateDBDataList_Insert_WML01_Lock);//0:sql 1:參數List
				   updateDBList.add(updateDBSqlList);
	            }
	            
	            if(updateDBDataList_Update_WML01_null.size() >= 1 ){
			       //解除鎖定(農會.漁會.農業信用保証基金寫到WML01)
			       sqlCmd_Update_WML01_null = " UPDATE WML01 SET lock_status = null ,user_id=?,user_name=?,update_date=sysdate "
						                    + " WHERE m_year=? AND m_month=? AND bank_code=? AND report_no=?";
			
			       updateDBSqlList = new ArrayList(); 
         		   updateDBSqlList.add(sqlCmd_Update_WML01_null);//0:欲執行的sql				
         		   updateDBSqlList.add(updateDBDataList_Update_WML01_null);//0:sql 1:參數List
				   updateDBList.add(updateDBSqlList);
	            }
	            
				if(updateDBDataList_Update_WML01.size() >= 1 ){
			       //鎖定(農會.漁會.農業信用保証基金寫到WML01)
			       sqlCmd_Update_WML01 = " UPDATE WML01 SET lock_status =? ,user_id=?,user_name=?,update_date=sysdate "
						        	+ " WHERE m_year=? AND m_month=? AND bank_code=? AND report_no=? ";
			
			       updateDBSqlList = new ArrayList(); 
         		   updateDBSqlList.add(sqlCmd_Update_WML01);//0:欲執行的sql				
         		   updateDBSqlList.add(updateDBDataList_Update_WML01);//0:sql 1:參數List
				   updateDBList.add(updateDBSqlList);
	            }
		
	             				
				if(DBManager.updateDB_ps(updateDBList)){					 
				   errMsg = errMsg + "相關資料寫入資料庫成功";					
				}else{
				   errMsg = errMsg + "相關資料寫入資料庫失敗";
				}	
			    
			    
			    if(errMsg.equals("相關資料寫入資料庫成功")){
    			   logps.println(logformat.format(nowlog)+" "+"執行"+lock_status_msg+"完成");				   
    			}else{
    			   logps.println(logformat.format(nowlog)+" "+"執行"+lock_status_msg+"失敗:"+errMsg);
    			}
			    logps.flush();		    
		}catch (Exception e){
				System.out.println(e+":"+e.getMessage());
				errMsg = errMsg + "相關資料寫入資料庫失敗";					
				logcalendar = Calendar.getInstance(); 
		        nowlog = logcalendar.getTime();   
	            logps.println(logformat.format(nowlog)+" "+"UpdateDB Error:"+e + "\n"+e.getMessage());	
		        logps.flush();		   
		}finally{
			try{
			   if (logos  != null) logos.close();
 	           if (logbos != null) logbos.close();
 	           if (logps  != null) logps.close();
			}catch(Exception ioe){
				System.out.println(ioe.getMessage());
		    }
		}	

		return errMsg;
	} 
%>	