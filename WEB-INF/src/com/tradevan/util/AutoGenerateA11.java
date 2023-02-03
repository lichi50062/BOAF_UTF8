/*
 * 102.12.11 create A11上月資料匯入 by 2295
 */
package com.tradevan.util;

import java.io.*;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.text.SimpleDateFormat;
import java.util.*;
import com.tradevan.util.Utility;



public class AutoGenerateA11{
    static final String driverName = "oracle.jdbc.driver.OracleDriver";
    static String dbURL = ""; 
    static String userName = ""; 
    static String userPwd = "";     
    static Connection dbConn = null;
    static ResultSet rs = null;
    static PreparedStatement pstmt = null;
    static File logfile;
    static FileOutputStream logos=null;      
    static BufferedOutputStream logbos = null;
    static PrintStream logps = null;
    static Date nowlog = new Date();
    static SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");        
    static SimpleDateFormat logfileformat = new SimpleDateFormat("yyyyMMddHHmmss");
    static Calendar logcalendar;
    static File logDir = null;
	public static void main(String[] args){
		System.out.println("AutoGenerateA11 begin");
		AutoGenerateA11 a = new AutoGenerateA11();
		a.generate("","");
		System.out.println("AutoGenerateA11 end");
	}


	public static void generate(String S_YEAR,String S_MONTH)
	{	 
	 String sqlCmd="";	   
	 
	 try{	    
	    logDir  = new File(Utility.getProperties("logDir"));
        if(!logDir.exists()){
            if(!Utility.mkdirs(Utility.getProperties("logDir"))){
               System.out.println("目錄新增失敗");
            }    
        }
        logfile = new File(logDir + System.getProperty("file.separator") + "AutoGenerateA11."+ logfileformat.format(nowlog));                       
        System.out.println("logfile filename="+logDir + System.getProperty("file.separator") +"AutoGenerateA11."+ logfileformat.format(nowlog));
        logos = new FileOutputStream(logfile,true);                         
        logbos = new BufferedOutputStream(logos);
        logps = new PrintStream(logbos);   
	    
	   	if("".equals(S_YEAR) || S_YEAR == null){    
	    Calendar now = Calendar.getInstance();
        S_YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
        S_MONTH = String.valueOf(now.get(Calendar.MONTH)+1);   //月份以0開始故加1取得實際月份;       
	   	}
	   	
        String S_YEAR_Last = "";//前一個月份.年
        String S_MONTH_Last = "";//前一個月份.月
  
        if(S_MONTH.equals("1")){//若本月為1月份是..則是申報上個年度的12月份
           S_YEAR = String.valueOf(Integer.parseInt(S_YEAR) - 1);
           S_MONTH = "12";           
        }else{    
           S_MONTH = String.valueOf(Integer.parseInt(S_MONTH) - 1);//申報上個月份的
        }
        if(S_MONTH.equals("1")){//若上月12月份是..則是匯入申報上個年度的1月份
            S_YEAR_Last = String.valueOf(Integer.parseInt(S_YEAR) - 1);
            S_MONTH_Last = "12";           
        }else{    
            S_YEAR_Last = S_YEAR;
            S_MONTH_Last = String.valueOf(Integer.parseInt(S_MONTH) - 1);//申報上個月份的
        }
        System.out.println("S_YEAR="+S_YEAR+":S_MONTH="+S_MONTH);
        System.out.println("S_YEAR_Last="+S_YEAR_Last+":S_MONTH_Last="+S_MONTH_Last);
        printLog(logps,"開始匯入"+S_YEAR_Last+"年"+S_MONTH_Last+"月份資料至"+S_YEAR+"年"+S_MONTH+"月份資料");
        dbURL = Utility.getProperties("BOAFDBURL");
        userName = Utility.getProperties("rptID");
        userPwd = Utility.getProperties("rptPwd");
        Class.forName(driverName); 
        dbConn = DriverManager.getConnection(dbURL, userName, userPwd); 
        System.out.println("Connection Successful!");
        
	    sqlCmd = " insert into wlx10_m_loan"
	           + " select ?,?,bank_no,seq_no,loan_idn,loan_name,loan_amt_sum,case_begin_year,case_begin_month,case_end_year, "
	           + " case_end_month,bank_no_max,manabank_name,loan_kind,loan_amt,loan_bal_amt,loan_type,pay_state,violate_type,loan_rate,"
	           + " new_case,user_id,user_name,update_date,case_no,case_cnt "
	           + " from wlx10_m_loan "
	           + " where m_year=? and m_month=? ";
	    pstmt = dbConn.prepareStatement(sqlCmd);
	    pstmt.setObject(1,S_YEAR);
	    pstmt.setObject(2,S_MONTH);
	    pstmt.setObject(3,S_YEAR_Last);
	    pstmt.setObject(4,S_MONTH_Last);
	    int rowCount = pstmt.executeUpdate();	    
	    if(rowCount == 0){
	        printLog(logps,"更新資料庫失敗");
	    }else{        
           System.out.println(" A11-UPDATE ok="+rowCount);    
           printLog(logps,"更新資料庫成功");
        }
        printLog(logps,"A11上月資料匯入完成");
        if(pstmt != null){//104.10.06 add
            pstmt.close();
            pstmt = null;             
        }
        if(!dbConn.isClosed()){//104.10.06
            dbConn.close();
            dbConn = null;
        }
     }catch (Exception e){
	    System.out.println(e.getMessage());
     }
    }//end of generate
	public static void printLog(PrintStream logps,String errRptMsg){
	       if(!errRptMsg.equals("")){
	          logcalendar = Calendar.getInstance(); 
	          nowlog = logcalendar.getTime();
	          logps.println(logformat.format(nowlog)+errRptMsg);
	          logps.flush();
	       }
	}
}



