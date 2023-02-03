/*
 * Created on 2017/12/28 排程產生Axx_operation_month關帳資料  by 2295
 */
package com.tradevan.util;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Calendar;


public class GenOperationMonth {
	 static final String driverName = "oracle.jdbc.driver.OracleDriver";
	 static String dbURL = "";//jdbc:oracle:thin:@172.20.5.22:1521:TBOAF"; 
	 static String userName = ""; 
	 static String userPwd = ""; 
	 static Connection dbConn = null;
	 static ResultSet rs = null;
	 static Statement stat = null;
	 static PreparedStatement pstmt = null;
	public static void main(String[] args){
		
		System.out.println("GenOperationMonth begin---"+args[0]);		 
		Calendar now = Calendar.getInstance();
		String lguser_id="BOAF000001";
		String report_no="";
		String isDebug="";
		String s_year="";
		String s_month="";
		String S_YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
	   	String S_MONTH = String.valueOf(now.get(Calendar.MONTH)+1);   //月份以0開始故加1取得實際月份;	   	    
	    
	   	if(S_MONTH.equals("1")){//若本月為1月份是..則是申報上個年度的12月份
	       S_YEAR = String.valueOf(Integer.parseInt(S_YEAR) - 1);
	       S_MONTH = "12";
	    }else{    
	       S_MONTH = String.valueOf(Integer.parseInt(S_MONTH) - 1);//申報上個月份的
	    }	
	    
	    if(args[0] != null) report_no = args[0];
	    if(args[1] != null) isDebug = args[1];	    
	    if(args.length == 4){
	       if(args[2] != null) s_year = args[2];
	       if(args[3] != null) s_month = args[3];
	    }
	    if(s_year.equals("")) s_year = S_YEAR;
	    if(s_month.equals("")) s_month = S_MONTH;  
	    
		GenOperationMonth a = new GenOperationMonth();
		/*
		String report_no="A01";
		String s_year="106";
		String s_month="11";
		String isDebug="true";
		String lguser_id="BOAF000001";
		*/
		
		System.out.println("report_no="+report_no);
		System.out.println("s_year="+s_year);
		System.out.println("s_month="+s_month);
		System.out.println("isDebug="+isDebug);
		
		
		a.exeGenOperationMonth(report_no,s_year,s_month,isDebug,lguser_id);
		System.out.println("GenOperationMonth end-----"+args[0]);
	}
	public void exeGenOperationMonth(String report_no,String s_year,String s_month,String isDebug,String lguser_id){		      	     
	    try{   		      
		       if(report_no.equals("A01")){//A01關帳資料    
		           deleteData(report_no,s_year,s_month);//刪除舊資料
		           genData(report_no,s_year,s_month);//產生A01關帳資料
		       }		       
     	}catch (Exception e){
			    System.out.println("exeGenOperationMonth Error:"+e.getMessage());
		}
    }	
	
	private void genData(String report_no,String s_year,String s_month){
        String sqlCmd = "";         
        try{
            Class.forName(driverName); 
            dbURL = Utility.getProperties("BOAFDBURL");
            userName = Utility.getProperties_conf("JDBC_USER");
            userPwd = Utility.getProperties_conf("JDBC_PASSWORD");
            dbConn = DriverManager.getConnection(dbURL, userName, userPwd); 
            System.out.println("Connection Successful!");
            
            sqlCmd = " insert into "+report_no+"_operation_month select * from "+report_no+"_operation where m_year=? and m_month=?";
            System.out.println("GenOperationMonth.sqlCmd="+sqlCmd);
            pstmt = dbConn.prepareStatement(sqlCmd);            
            pstmt.setObject(1, s_year); 
            pstmt.setObject(2, s_month);          
            int rowCount = pstmt.executeUpdate();
            System.out.println("[rowCount]="+rowCount);
            System.out.println("genData ["+report_no+"_operation_month] ??"+rowCount);
            dbConn.commit();
         }catch(Exception e){
            System.out.println("genData Error:"+e+e.getMessage());
         }finally{
            try{
                
                if (pstmt != null){
                    pstmt.close(); 
                    pstmt = null;
                }
                //if (dbConn != null){
                if(!dbConn.isClosed()){//104.10.06    
                    dbConn.close();
                    dbConn = null;
                }
            }catch(SQLException sqle){
                   System.out.println(sqle+sqle.getMessage());
            }  
         }  
    }
	
	private void deleteData(String report_no,String s_year,String s_month){
		String sqlCmd = ""; 		
		try{
		 	Class.forName(driverName); 
		 	dbURL = Utility.getProperties("BOAFDBURL");
		 	userName = Utility.getProperties_conf("JDBC_USER");
		 	userPwd = Utility.getProperties_conf("JDBC_PASSWORD");
		 	dbConn = DriverManager.getConnection(dbURL, userName, userPwd); 
		 	System.out.println("Connection Successful!");
		 	//stat = dbConn.createStatement(); 
            sqlCmd = " select count(*) as datacount from "+report_no+"_operation_month where m_year=? and m_month=?";
            System.out.println("GenOperationMonth.sqlCmd="+sqlCmd);
            pstmt = dbConn.prepareStatement(sqlCmd);            
            pstmt.setObject(1, s_year); 
            pstmt.setObject(2, s_month);
            rs = pstmt.executeQuery();     
           
            while (rs.next()) {
                if(rs.getInt("datacount") > 0){
                   System.out.println(report_no+"_operation_month multi data");
     	    	   sqlCmd = " delete "+report_no+"_operation_month where m_year=? and m_month=?";
     	    	   pstmt.clearBatch();
     	    	   pstmt.clearParameters();
     	    	   pstmt = dbConn.prepareStatement(sqlCmd); 
     	    	   pstmt.setObject(1, s_year); 
     	           pstmt.setObject(2, s_month);
     	    	   System.out.println("GenOperationMonth.sqlCmd="+sqlCmd);
     	    	   System.out.println("deleteData ["+report_no+"] ??"+pstmt.executeUpdate());
     	    	   dbConn.commit();
                } 
	        }  
		  
         }catch(Exception e){
         	System.out.println("deleteData Error:"+e+e.getMessage());
         }finally{
	        try{
	            if (rs != null){
	                rs.close();
	                rs = null;
	            }
	            //if (stat != null) stat.close();
	            if (pstmt != null){
	                pstmt.close(); 
	                pstmt = null;
	            }
	            //if (dbConn != null){
	            if(!dbConn.isClosed()){//104.10.06    
	                dbConn.close();
	                dbConn = null;
	            }
	        }catch(SQLException sqle){
	               System.out.println(sqle+sqle.getMessage());
	        }  
         }  
	}
}

