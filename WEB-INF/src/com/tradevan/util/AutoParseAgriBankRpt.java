/*
 *  99.12.17 create 金庫排程輚檔程式 by 2295 
 * 100.09.26 fix 修正Hard-Coded Password by 2295
 */
package com.tradevan.util;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.net.URL;
import java.net.URLConnection;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Calendar;


public class AutoParseAgriBankRpt {
	 static final String driverName = "oracle.jdbc.driver.OracleDriver";
	 static String dbURL = "";//jdbc:oracle:thin:@172.20.5.22:1521:TBOAF"; 
	 static String userName = ""; 
	 static String userPwd = ""; 
	 static Connection dbConn = null;
	 static ResultSet rs = null;
	 static Statement stat = null;
	
	public static void main(String[] args){
		
		System.out.println("GenerateOperation begin---"+args[0]);		 
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
	    
		GenerateOperation a = new GenerateOperation();
		/*
		report_no="A04";
		s_year="94";
		s_month="8";
		isDebug="true";
		*/
		System.out.println("report_no="+report_no);
		System.out.println("s_year="+s_year);
		System.out.println("s_month="+s_month);
		System.out.println("isDebug="+isDebug);
		
		
		a.exeGenerateOperation(report_no,s_year,s_month,isDebug,lguser_id);
		System.out.println("GenerateOperation end-----"+args[0]);
	}
	public void exeGenerateOperation(String report_no,String s_year,String s_month,String isDebug,String lguser_id){		      	     
	    try{
		       URL myURL=null;
		       URLConnection raoURL;
		       String raoInputString;
		       //String webURL = "https://ebankweb.boaf.gov.tw/";//正式
		       //String webURL = "http://localhost:82/";//測試
		       if(report_no.equals("A01")){
		           deleteData(report_no,s_year,s_month);//刪除舊資料
		           //農.漁會各別機構
		           myURL = new URL(Utility.getProperties("BOAFWebSite")+"A01InsertDB_bank_no_ALL.jsp?report_no="+report_no+"&s_year="+s_year+"&s_month="+s_month+"&isDebug="+isDebug+"&lguser_id="+lguser_id);
		           runURL(myURL);
			       //農會.縣市別小計
		           //myURL = new URL(Utility.getProperties("BOAFWebSite")+"A01InsertDB_hsien_id_6.jsp?report_no="+report_no+"&s_year="+s_year+"&s_month="+s_month+"&isDebug="+isDebug+"&lguser_id="+lguser_id);
		           myURL = new URL(Utility.getProperties("BOAFWebSite")+"A01InsertDB_hsien_id.jsp?report_no="+report_no+"&s_year="+s_year+"&s_month="+s_month+"&bank_type=6&isDebug="+isDebug+"&lguser_id="+lguser_id);
		           runURL(myURL);
		           //漁會.縣市別小計
		           //myURL = new URL(Utility.getProperties("BOAFWebSite")+"A01InsertDB_hsien_id_7.jsp?report_no="+report_no+"&s_year="+s_year+"&s_month="+s_month+"&isDebug="+isDebug+"&lguser_id="+lguser_id);
		           myURL = new URL(Utility.getProperties("BOAFWebSite")+"A01InsertDB_hsien_id.jsp?report_no="+report_no+"&s_year="+s_year+"&s_month="+s_month+"&bank_type=7&isDebug="+isDebug+"&lguser_id="+lguser_id);
		           runURL(myURL);
		           //全体農漁會.縣市別小計
		           myURL = new URL(Utility.getProperties("BOAFWebSite")+"A01InsertDB_hsien_id_ALL.jsp?report_no="+report_no+"&s_year="+s_year+"&s_month="+s_month+"&isDebug="+isDebug+"&lguser_id="+lguser_id);
		           runURL(myURL);
		       }else if(report_no.equals("A02")){
		       	   deleteData(report_no,s_year,s_month);//刪除舊資料
		           //農.漁會各別機構
		           myURL = new URL(Utility.getProperties("BOAFWebSite")+"A02InsertDB_bank_no.jsp?report_no="+report_no+"&s_year="+s_year+"&s_month="+s_month+"&isDebug="+isDebug+"&lguser_id="+lguser_id);
		           runURL(myURL);
		       }else if(report_no.equals("A03")){
		       	   deleteData(report_no,s_year,s_month);//刪除舊資料
		           //農會.各別機構 
		           myURL = new URL(Utility.getProperties("BOAFWebSite")+"A03InsertDB_bank_no_6.jsp?report_no="+report_no+"&s_year="+s_year+"&s_month="+s_month+"&isDebug="+isDebug+"&lguser_id="+lguser_id);
		           runURL(myURL);
		           //漁會.各別機構 
		           myURL = new URL(Utility.getProperties("BOAFWebSite")+"A03InsertDB_bank_no_7.jsp?report_no="+report_no+"&s_year="+s_year+"&s_month="+s_month+"&isDebug="+isDebug+"&lguser_id="+lguser_id);
		           runURL(myURL);
		           //農會.縣市別小計 
		           myURL = new URL(Utility.getProperties("BOAFWebSite")+"A03InsertDB_hsien_id_6.jsp?report_no="+report_no+"&s_year="+s_year+"&s_month="+s_month+"&isDebug="+isDebug+"&lguser_id="+lguser_id);
		           runURL(myURL);
		           //漁會.縣市別小計
		     	   myURL = new URL(Utility.getProperties("BOAFWebSite")+"A03InsertDB_hsien_id_7.jsp?report_no="+report_no+"&s_year="+s_year+"&s_month="+s_month+"&isDebug="+isDebug+"&lguser_id="+lguser_id);
		     	   runURL(myURL);
		       }else if(report_no.equals("A04")){
		       	   deleteData(report_no,s_year,s_month);//刪除舊資料
		           //農.漁會.各別機構 
		           myURL = new URL(Utility.getProperties("BOAFWebSite")+"A04InsertDB_bank_no_67.jsp?report_no="+report_no+"&s_year="+s_year+"&s_month="+s_month+"&isDebug="+isDebug+"&lguser_id="+lguser_id);
		           runURL(myURL);		           
		           //農會.縣市別小計 
		           myURL = new URL(Utility.getProperties("BOAFWebSite")+"A04InsertDB_hsien_id_6.jsp?report_no="+report_no+"&s_year="+s_year+"&s_month="+s_month+"&isDebug="+isDebug+"&lguser_id="+lguser_id);
		           runURL(myURL);
		           //漁會.縣市別小計
		     	   myURL = new URL(Utility.getProperties("BOAFWebSite")+"A04InsertDB_hsien_id_7.jsp?report_no="+report_no+"&s_year="+s_year+"&s_month="+s_month+"&isDebug="+isDebug+"&lguser_id="+lguser_id);
		     	   runURL(myURL);	   		
		     	   //全体農漁會.縣市別小計
		     	   myURL = new URL(Utility.getProperties("BOAFWebSite")+"A04InsertDB_hsien_id_ALL.jsp?report_no="+report_no+"&s_year="+s_year+"&s_month="+s_month+"&isDebug="+isDebug+"&lguser_id="+lguser_id);
		     	   runURL(myURL);
		       }		       
     	}catch (Exception e){
			    System.out.println("exeGenerateOperation Error:"+e.getMessage());
		}
    }	
	private void runURL(URL myURL){
	    String raoInputString;
	    URLConnection raoURL;
	    try{
	        raoURL=myURL.openConnection();
	        raoURL.setDoInput(true);
		    raoURL.setDoOutput(true);
		    raoURL.setUseCaches(false);
		    BufferedReader infromURL=new BufferedReader(new InputStreamReader(raoURL.getInputStream()));	     	   	     	  
	     	while((raoInputString=infromURL.readLine())!=null){
	       	       System.out.println(raoInputString);
	     	}
	     	infromURL.close();
	    }catch (Exception e){
		    System.out.println("GenerateOperation Error:"+e.getMessage());
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
		 	stat = dbConn.createStatement(); 
         
            rs = stat.executeQuery(" select count(*) as datacount from "+report_no+"_operation where m_year="+s_year+" and m_month="+s_month);         
            while (rs.next()) {
                if(rs.getInt("datacount") > 0){
                   System.out.println(report_no+"_operation multi data");
     	    	   sqlCmd = " delete "+report_no+"_operation where m_year="+s_year+" and m_month="+s_month;
     	    	   System.out.println(sqlCmd);
     	    	   System.out.println("deleteData ["+report_no+"] ??"+stat.executeUpdate(sqlCmd));
     	    	   dbConn.commit();
                } 
	        }  
		  
         }catch(Exception e){
         	System.out.println("deleteData Error:"+e+e.getMessage());
         }finally{
	        try{
	            if (rs != null) rs.close();
	            if (stat != null) stat.close(); 
	            if (dbConn != null) dbConn.close(); 
	        }catch(SQLException sqle){
	               System.out.println(sqle+sqle.getMessage());
	        }  
         }  
	}
}

