<%
// 95.09.05 by 2495
// 96.10.29 add 牌告改成月報 by 2295
// 96.10.29 add 刪除舊資料 by 2295
// 96.11.05 add 牌告利率15日.抓當月份資料.5日抓上個月份資料,其他的15日抓上個月的申報資料 by 2295
// 98.07.31 fix 牌告利率5日抓上個月份資料時,若有舊資料.先刪除 by 2295
//103.04.02 add 專案農貸明細資料/專案農貸貸款項目資料  by 2295
//103.05.19 add 金庫取檔使用sftp by 2295
%>
<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<%@ page import="java.util.*,java.io.*" %>
<%@ page import="org.apache.poi.poifs.filesystem.*,org.apache.poi.hssf.usermodel.*" %>
<%@ page import="com.tradevan.util.Utility" %>
<%@ page import="com.tradevan.util.ftp.AgriBankFTP" %>
<%@ page import="java.util.Properties" %>
<%@ page import="java.util.Enumeration" %>
<%@ page import="com.tradevan.util.DBManager" %>
<%@ page import="com.tradevan.util.dao.DataObject" %>
<%@ page import="com.tradevan.util.ListArray" %>
<%
 
 //全國農業金庫主機
 String FTP_SERVER = "59.124.54.12";
 //String FTP_SERVER = "59.124.54.11";
 String FTP_USER = "TNBOAF";
 String FTP_PASSWORD = "TNBOAF123";
 String FTP_DIRECTORY = "/out/";
 String LOCAL_DIRECTORY = "C:\\Sun\\WebServer6.1\\BOAF\\AgriBankData";
 /*
 //公司測試主機
 String FTP_SERVER = "172.20.5.22";
 String FTP_USER = "pboafmgr";
 String FTP_PASSWORD = "tvpboafmgr";
 String FTP_DIRECTORY = "/export/home/pboafmgr/test/"; 
 */
 //String LOCAL_DIRECTORY = "D:\\workProject\\BOAF\\AgriBankData";
 String alertMsg = "";	
 String TargetFile = "";
 String sqlCmd = "";
 String sqlCmd_delete=" delete ";
 List dbData = new LinkedList();
 List updateDBSqlList = new LinkedList();
 int FileCount=6;
 List paramList = new ArrayList();
 Calendar now = Calendar.getInstance();
 String year  = String.valueOf(now.get(Calendar.YEAR)); //回覆值為西元年;
 String month = String.valueOf(now.get(Calendar.MONTH)+1);   //月份以0開始故加1取得實際月份;
 String dday  = String.valueOf(now.get(Calendar.DATE));
 if(month.equals("1")){//若本月為1月份是..則是申報上個年度的12月份
    year = String.valueOf(Integer.parseInt(year) - 1);
    month = "12";
 }else{    
    month = String.valueOf(Integer.parseInt(month) - 1);//申報上個月份的
 }  
 if(month.length()==1) month="0"+month;										 
 System.out.println("year="+year+":month="+month+":day="+dday); 
 
 //15號傳ATM,CHK,FGN
 if(dday.equals("15"))
 {      
	for(int i=0;i<FileCount;i++){   	    
	    sqlCmd = "select count(*) as countdata from ";
	    sqlCmd_delete=" delete ";
	    dbData = new LinkedList();
        updateDBSqlList = new LinkedList();
        paramList = new ArrayList();
 	    System.out.println("15day parser ATM.CHG.FGN.INT ");
 	    System.out.println("i="+i+":year="+year+":month="+month+":day="+dday);
 	    
		if(i==0){//農漁會金融卡發卡情形及ATM裝設情形統計
		   TargetFile = "ATM"+year+month+".DAT";
		   sqlCmd += " wlx05_m_atm ";
 		   sqlCmd_delete += " wlx05_m_atm ";
		}else if(i==1){//農漁會支票存款資料
		   TargetFile = "CHK"+year+month+".DAT";
		   sqlCmd += " WLX07_M_CHECKBANK "; 
 		   sqlCmd_delete += " WLX07_M_CHECKBANK ";		
		}else if(i==2){//外國人存款資料
		   TargetFile = "FGN"+year+month+".DAT";	
		   sqlCmd += " F01 "; 				
 		   sqlCmd_delete += " F01 ";  		   
		}else if(i==3){//農漁會信用部牌告利率申報資料
		   //96.11.05.牌告利率15日.抓當月份資料	
		   year  = String.valueOf(now.get(Calendar.YEAR)); //回覆值為西元年;
 		   month = String.valueOf(now.get(Calendar.MONTH)+1);//月份以0開始故加1取得實際月份; 		   
 		   if(month.length()==1) month="0"+month;										 
 		   System.out.println("year="+year+":month="+month+":day="+dday); 
		   TargetFile = "INT"+year+month+".DAT";//牌告改成月報  
		   sqlCmd += " WLX_S_RATE ";
 		   sqlCmd_delete += " WLX_S_RATE ";
 		}else if(i==4){//專案農貸
 		   TargetFile = "FRM"+year+month+".DAT";
 		   sqlCmd += " AGRI_LOAN "; 				
 		   sqlCmd_delete += " AGRI_LOAN ";   		
 		}else if(i==5){//貸款項目代碼
 		   TargetFile = "ITEM"+year+month+".DAT";
 		   sqlCmd += " AGRI_LOAN_ITEM "; 				
 		   sqlCmd_delete += " AGRI_LOAN_ITEM ";     
		}   
		
 		System.out.println("TargetFile="+TargetFile); 
 		if(i !=5){//貸款項目代碼
 		   sqlCmd += " where m_year=?"; 		       
 	       sqlCmd_delete += " where m_year=?";
 	       paramList.add(Integer.toString(Integer.parseInt(year)-1911));
	    }
 		if(i==3){
 		   sqlCmd += " and m_quarter=?"; 				  
 		   sqlCmd_delete += " and m_quarter=?"; 				  
 		}else{
 		   if(i !=5){//貸款項目代碼
 		      sqlCmd += " and m_month=?"; 				  
 		      sqlCmd_delete += " and m_month=?"; 	
 		   }			  
 		}    			   
 		paramList.add(month);	 
 		dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"countdata"); 
 		if(dbData != null){
 		   if(!((((DataObject)dbData.get(0)).getValue("countdata")).toString()).equals("0")){
 			  System.out.println(" have "+(((DataObject)dbData.get(0)).getValue("countdata")).toString() +" datas");
 			   updateDBSqlList.add(sqlCmd_delete); 					     		         	   						         	   		  
           }//end of countdata    	   
    	}//end of query data		    	       
    	if(i==2){//96.01.10 fix 外國人存款資料.加上刪除wml01    	   
    	   dbData = DBManager.QueryDB_SQLParam("select count(*) as countdata from wml01 where m_year=? and m_month=? and report_no='F01'",paramList,"countdata");
    	   if(dbData != null && dbData.size() ==1){
 			  if(!((((DataObject)dbData.get(0)).getValue("countdata")).toString()).equals("0")){
 			    System.out.println(" wml01 have "+(((DataObject)dbData.get(0)).getValue("countdata")).toString() +" datas"); 					    
 			    updateDBSqlList.add(" delete wml01 where m_year="+Integer.toString(Integer.parseInt(year)-1911)+" and m_month="+month+" and report_no='F01'");
        	  }//end of countdata    	   
    	   }//end of query data		 				   
 	    }
 		if(updateDBSqlList.size() != 0){
 		    for(int delidx=0;delidx<updateDBSqlList.size();delidx++){
 		         System.out.println(updateDBSqlList.get(delidx)); 			          
 		    }
 		    if(!DBManager.updateDB(updateDBSqlList)){
    	       System.out.println("舊有資料刪除失敗");
    	    }
 		}//end of delete data    
 		int result = -1;   	   
    	if(alertMsg.equals("")){
    	   result = AgriBankFTP.getDataFiles(FTP_SERVER, FTP_USER,FTP_PASSWORD, FTP_DIRECTORY, LOCAL_DIRECTORY, TargetFile,"true");
        }  
        
 	} 
 }
 //改成月報
 //5號傳INT.上月份資料. 
 if(dday.equals("5"))
 {
 	 	System.out.println( "5-INT  year.... : " + year);
 		System.out.println( "5-INT  month : " + month);
 		System.out.println( "5-INT  day : " + dday);
 	 	//month=Integer.toString(Integer.parseInt(month)+1);
 	 	/*
 	 	if (month.equals("1")||month.equals("2")||month.equals("3"))
 	 		 month ="01";
 	 	if (month.equals("4")||month.equals("5")||month.equals("6"))
 	 		 month ="02";
 	    if (month.equals("7")||month.equals("8")||month.equals("9"))
 	 		 month ="03";
 	 	if (month.equals("10")||month.equals("11")||month.equals("12"))
 	 		 month ="04";	 
 	 	*/ 
 	  
 	   
 	 	sqlCmd = "select count(*) as countdata from WLX_S_RATE where m_year=? and m_quarter=?";
 	 	paramList.add(Integer.toString(Integer.parseInt(year)-1911));
 	 	paramList.add(month);
 	 	System.out.println(sqlCmd);
 	 	//98.07.31 fix 牌告利率舊資料.加上刪除作業
    	dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"countdata");
    	System.out.println("dbData.size()="+dbData.size());
    	if(dbData != null && dbData.size() ==1){
 		   if(!((((DataObject)dbData.get(0)).getValue("countdata")).toString()).equals("0")){
 			    System.out.println(" WLX_S_RATE have "+(((DataObject)dbData.get(0)).getValue("countdata")).toString() +" datas"); 					    
 			    updateDBSqlList.add(" delete WLX_S_RATE where m_year="+Integer.toString(Integer.parseInt(year)-1911)+" and m_quarter="+month);
           }//end of countdata    	   
    	}//end of query data		 				   
 	    
 		if(updateDBSqlList.size() != 0){
 		    for(int delidx=0;delidx<updateDBSqlList.size();delidx++){
 		         System.out.println(updateDBSqlList.get(delidx)); 			          
 		    }
 		    if(!DBManager.updateDB(updateDBSqlList)){
    	       System.out.println("牌告利率舊有資料刪除失敗");
    	    }
 		}//end of delete data    
 	 	TargetFile = "INT"+year+month+".DAT";  
 	 	System.out.println( "05day parser INT TargetFile .... : " + TargetFile);
		AgriBankFTP.getDataFiles(FTP_SERVER, FTP_USER,FTP_PASSWORD, FTP_DIRECTORY, LOCAL_DIRECTORY, TargetFile,"false");
		
	}
	
%>















