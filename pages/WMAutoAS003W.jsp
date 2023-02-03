<%
//101.07 created on 2012/07
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
	System.out.println("=============執行農漁會信用部家數月結作業開始===========");
	String errMsg = UpdateDB();
	System.out.println("errMsg = "+errMsg);
	System.out.println("=============執行農漁會信用部家數月結作業結束===========");
	
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
	            
	            logfile = new File(logDir + System.getProperty("file.separator") + "AutoAS003W."+ logfileformat.format(nowlog));						 
				System.out.println("logfile filename="+logDir + System.getProperty("file.separator") +"AutoAS003W.log"+ logfileformat.format(nowlog));
				logos = new FileOutputStream(logfile,true);  		        	   
				logbos = new BufferedOutputStream(logos);
				logps = new PrintStream(logbos);	
			    
				sqlCmd = "";
			    paramList = new ArrayList();
			    sqlCmd = " select count(*) count from bn01 where bank_type=? and bn_type <> ? and m_year=? ";
			    paramList.add("6");
			    paramList.add("2");	
			    paramList.add(wlx01_m_year);	
			    List qlist1 = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"count"); 
			    System.out.println("農會qlist.size:"+qlist1.size());
			    sqlCmd = "";
			    paramList = new ArrayList();
			    sqlCmd = " select count(*) count from bn01 where bank_type=? and bn_type <> ? and m_year=? ";
			    paramList.add("7");
			    paramList.add("2");	
			    paramList.add(wlx01_m_year);	
			    List qlist2 = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"count");
			    System.out.println("漁會qlist.size:"+qlist2.size());
			    String count1 = ((DataObject)qlist1.get(0)).getValue("count").toString();
			    String count2 = ((DataObject)qlist2.get(0)).getValue("count").toString();
			    
			    //農(漁)會信用部月底總機構家數關帳紀錄
			    sqlCmd = "";
			    paramList = new ArrayList();
			    sqlCmd = "SELECT * FROM bn01_month WHERE m_year = ? AND m_month= ? " ;
			    paramList.add(S_YEAR);
			    paramList.add(S_MONTH);
				List dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"update_date");
				System.out.println("dbData.size()="+dbData.size());
				int flag = 0;
				if(dbData.size()>0){
				    if ("".equals(((DataObject)dbData.get(0)).getValue("update_date").toString()) ){
					    sqlCmd = "";
					    sqlCmd = "DELETE FROM bn01_month WHERE m_year = ? AND m_month = ?  " ;
					    dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						updateDBDataList.add(dataList);
		         		updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				
		         		updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
						updateDBList.add(updateDBSqlList);
		         		
						sqlCmd = "";
					    dataList = new ArrayList();
					    sqlCmd = "INSERT INTO bn01_month " 
					        	         + " ( m_year, m_month, bank_sum_6, bank_sum_7,update_date ) "
					        	         + " VALUES ( ?, ?, ?, ?,sysdate ) ";
					    dataList.add(S_YEAR);
						dataList.add(S_MONTH);
						dataList.add(count1);
						dataList.add(count2);
		         		
				    }else{
					    sqlCmd = "";
					    sqlCmd = "UPDATE bn01_month SET " 
					        	     + " bank_sum_6 = ?,"
					        	     + " bank_sum_7 = ?, "
					        	     + " update_date = sysdate  "        
					    	   + " WHERE m_year = ? AND m_month = ?  ";
						dataList.add(count1);
						dataList.add(count2);
						dataList.add(S_YEAR);
						dataList.add(S_MONTH);
				    }
				    flag=1;
				}else{
				    
				    sqlCmd = "";
				    dataList = new ArrayList();
				    sqlCmd = "INSERT INTO bn01_month " 
				        	         + " ( m_year, m_month, bank_sum_6, bank_sum_7,update_date ) "
				        	         + " VALUES ( ?, ?, ?, ?,sysdate ) ";
				    dataList.add(S_YEAR);
					dataList.add(S_MONTH);
					dataList.add(count1);
					dataList.add(count2);
					
				}
				
				updateDBDataList.add(dataList);
         		updateDBSqlList.add(sqlCmd.toString());//0:欲執行的sql				
         		updateDBSqlList.add(updateDBDataList);//0:sql 1:參數List
				updateDBList.add(updateDBSqlList);
         		System.out.println(updateDBSqlList);
         		
         		if(DBManager.updateDB_ps(updateDBList)){					 
				   errMsg = errMsg + "相關資料寫入資料庫成功";
				   System.out.println(errMsg);
				}else{
				   errMsg = errMsg + "相關資料寫入資料庫失敗";
				   System.out.println(errMsg);
				}
         		
         		sqlCmd = "";
			    paramList = new ArrayList();
			    sqlCmd = "SELECT bank_sum_6,bank_sum_7,update_date FROM bn01_month WHERE m_year = ? AND m_month= ? " ;
			    paramList.add(S_YEAR);
			    paramList.add(S_MONTH);
				List qData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"bank_sum_6,bank_sum_7,update_date");
         		if(errMsg.equals("相關資料寫入資料庫成功")){
			    	logps.println(logformat.format(((DataObject)qData.get(0)).getValue("update_date"))+" "+S_YEAR+"年"+S_MONTH+"月農會總機構家數:"+((DataObject)qData.get(0)).getValue("bank_sum_6").toString());		    		
			    	logps.println(logformat.format(((DataObject)qData.get(0)).getValue("update_date"))+" "+S_YEAR+"年"+S_MONTH+"月漁會總機構家數:"+((DataObject)qData.get(0)).getValue("bank_sum_7").toString());
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
	