<%
//104.07 created by 2968
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
	System.out.println("=============執行農漁會信用部人員配置月結作業開始===========");
	String errMsg = UpdateDB();
	System.out.println("errMsg = "+errMsg);
	System.out.println("=============執行農漁會信用部人員配置月結作業結束===========");
	
%>
<%!
public String UpdateDB() throws Exception{    	
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
		String errMsg= "";
	    String bank_count= "" ;
	    Calendar now = Calendar.getInstance();
	    String S_YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
   	    String S_MONTH = String.valueOf(now.get(Calendar.MONTH)+1);   //月份以0開始故加1取得實際月份;
   	 	
	    List paramList = new ArrayList();
	    String cd01_table = "";
        String wlx01_m_year = "";
        
        //100.07.20 add 查詢年度100年以前.縣市別不同===============================
  	    cd01_table = (Integer.parseInt(S_YEAR) < 100)?"cd01_99":""; 
  	    wlx01_m_year = (Integer.parseInt(S_YEAR) < 100)?"99":"100"; 
  	    //=====================================================================    
        List data = null;	
        List updateDBList = new ArrayList();//0:sql 1:data
	    List updateDBSqlList = new ArrayList();//欲執行updatedb的sql list
		List updateDBDataList = new ArrayList();//儲存參數的List		    
		List dataList =  new ArrayList();//儲存參數的data
		String sqlCmd_Insert_bn01_month="";
		
		try {			    			    
			    logDir  = new File(Utility.getProperties("logDir"));
	            if(!logDir.exists()){
     			    if(!Utility.mkdirs(Utility.getProperties("logDir"))){
     				   System.out.println("目錄新增失敗");
     			    }    
    		    }
	            
	            logfile = new File(logDir + System.getProperty("file.separator") + "AutoFR072W."+ logfileformat.format(nowlog));						 
				System.out.println("logfile filename="+logDir + System.getProperty("file.separator") +"AutoFR072W.log"+ logfileformat.format(nowlog));
				logos = new FileOutputStream(logfile,true);  		        	   
				logbos = new BufferedOutputStream(logos);
				logps = new PrintStream(logbos);	
			    
				sqlCmd = "";
			    paramList = new ArrayList();
			    sqlCmd = " select t2.bank_type,t1.bank_no,staff_num,credit_staff_num,credit_staff,skill_staff,manual_staff,temp_staff from wlx01 t1 "
			    	   + " left join (select bank_no,bank_type from  bn01 where m_year=?) t2 on t1.bank_no = t2.bank_no where t1.m_year=? ";
			    paramList.add(wlx01_m_year);	
			    paramList.add(wlx01_m_year);
			    List qlist1 = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"bank_type,bank_no,staff_num,credit_staff_num,credit_staff,skill_staff,manual_staff,temp_staff"); 
			    System.out.println("qlist1.size:"+qlist1.size());
			    
			    
			    
			    //關帳紀錄
			    paramList.clear();
			    sqlCmd = "SELECT update_date FROM rpt_month WHERE m_year = ? AND m_month= ? and report_no='FR072W' " ;
			    paramList.add(S_YEAR);
			    paramList.add(S_MONTH);
				List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"update_date");
				System.out.println("dbData.size()="+dbData.size());
	         	
				if(qlist1!=null && qlist1.size()>0){
					if(dbData != null && dbData.size() > 0){
						dataList = new ArrayList();
					    updateDBDataList = new ArrayList();
		         		updateDBSqlList = new ArrayList();
					    sqlCmd = "DELETE rpt_month WHERE m_year = ? AND m_month= ? and report_no='FR072W'  " ;
					    dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						updateDBDataList.add(dataList);
		         		updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				
		         		updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
						updateDBList.add(updateDBSqlList);
						DBManager.updateDB_ps(updateDBList);
					}
					for(int i=0;i<qlist1.size();i++){
						DataObject obj = (DataObject)qlist1.get(i);
						String bank_type = obj.getValue("bank_type")==null?"6":Utility.getTrimString(obj.getValue("bank_type"));
						String bank_no = obj.getValue("bank_no")==null?"":Utility.getTrimString(obj.getValue("bank_no"));
						String staff_num = obj.getValue("staff_num")==null?"0":obj.getValue("staff_num").toString();
						String credit_staff_num = obj.getValue("credit_staff_num")==null?"0":obj.getValue("credit_staff_num").toString();
						String credit_staff = obj.getValue("credit_staff")==null?"0":obj.getValue("credit_staff").toString();
						String skill_staff = obj.getValue("skill_staff")==null?"0":obj.getValue("skill_staff").toString();
						String manual_staff = obj.getValue("manual_staff")==null?"0":obj.getValue("manual_staff").toString();
						String temp_staff = obj.getValue("temp_staff")==null?"0":obj.getValue("temp_staff").toString();
						
					    
					    sqlCmd = "INSERT INTO rpt_month " 
					        + " ( m_year, m_month,report_no,bank_type,bank_code,acc_code,type,amt,update_date) "
					        + " VALUES ( ?, ?,'FR072W', ?, ?, ?, 0, ?,sysdate) ";
					    dataList = new ArrayList();
					    updateDBDataList = new ArrayList();
		         		updateDBSqlList = new ArrayList();
		         		updateDBList = new ArrayList(); 
					    dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(bank_type);
						dataList.add(bank_no);
						dataList.add("staff_num");//acc_code
						dataList.add(staff_num);//amt
						updateDBDataList.add(dataList);
		         		updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				
		         		updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
						updateDBList.add(updateDBSqlList);
		         		
						dataList = new ArrayList();
						updateDBDataList = new ArrayList();
		         		updateDBSqlList = new ArrayList();
					    dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(bank_type);
						dataList.add(bank_no);
						dataList.add("credit_staff_num");//acc_code
						dataList.add(credit_staff_num);//amt
						updateDBDataList.add(dataList);
		         		updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				
		         		updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
						updateDBList.add(updateDBSqlList);
		         		
						dataList = new ArrayList();
						updateDBDataList = new ArrayList();
		         		updateDBSqlList = new ArrayList();
					    dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(bank_type);
						dataList.add(bank_no);
						dataList.add("credit_staff");//acc_code
						dataList.add(credit_staff);//amt
						updateDBDataList.add(dataList);
		         		updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				
		         		updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
						updateDBList.add(updateDBSqlList);
		         		
						dataList = new ArrayList();
						updateDBDataList = new ArrayList();
		         		updateDBSqlList = new ArrayList();
					    dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(bank_type);
						dataList.add(bank_no);
						dataList.add("skill_staff");//acc_code
						dataList.add(skill_staff);//amt
						updateDBDataList.add(dataList);
		         		updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				
		         		updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
						updateDBList.add(updateDBSqlList);
		         		
						dataList = new ArrayList();
						updateDBDataList = new ArrayList();
		         		updateDBSqlList = new ArrayList();
					    dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(bank_type);
						dataList.add(bank_no);
						dataList.add("manual_staff");//acc_code
						dataList.add(manual_staff);//amt
						updateDBDataList.add(dataList);
		         		updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				
		         		updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
						updateDBList.add(updateDBSqlList);

						dataList = new ArrayList();
						updateDBDataList = new ArrayList();
		         		updateDBSqlList = new ArrayList();
					    dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(bank_type);
						dataList.add(bank_no);
						dataList.add("temp_staff");//acc_code
						dataList.add(temp_staff);//amt
						updateDBDataList.add(dataList);
		         		updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				
		         		updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
						updateDBList.add(updateDBSqlList);
		         		
						if(updateDBList.size()>0){
			         		if(DBManager.updateDB_ps(updateDBList)){					 
							   errMsg = "相關資料寫入資料庫成功";
							   System.out.println(errMsg);
							}else{
							   errMsg = "相關資料寫入資料庫失敗";
							   System.out.println(errMsg);
							}
						}
					}
					
					
				}
         		
				paramList = new ArrayList();
				sqlCmd = "SELECT update_date FROM rpt_month WHERE m_year = ? AND m_month= ? and report_no='FR072W' " ;
				paramList.add(S_YEAR);
				paramList.add(S_MONTH);
				List qData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"update_date");
	         	if(errMsg.equals("相關資料寫入資料庫成功")){
				    logps.println(logformat.format(((DataObject)qData.get(0)).getValue("update_date"))+" "+S_YEAR+"年"+S_MONTH+"月人員配置資料產生完成");
				    logps.flush();				   
	     		}else{
	     			logps.println(logformat.format(nowlog)+" "+"執行關帳失敗 :"+errMsg);
	     		}
         		
         		
         		
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
<%!

	private File logfile;
	private FileOutputStream logos=null;    	
	private BufferedOutputStream logbos = null;
	private PrintStream logps = null;
	private Date nowlog = new Date();
	private SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");	     
	private SimpleDateFormat logfileformat = new SimpleDateFormat("yyyyMMddHHmmss");
	private Calendar logcalendar;
	private File logDir = null;
%>	
	