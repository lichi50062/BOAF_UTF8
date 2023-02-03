/*
 * 102.07.30 created by 2968
 * 103.03.18 為配合103.02需求，結合WR00InsertDB_bank_no_ALL.jsp、WR01InsertDB_bank_no_ALL.jsp、WR02InsertDB_bank_no_ALL.jsp  by 2968
 */
package com.tradevan.util;

import java.io.BufferedReader;
import java.io.InputStreamReader;
import java.net.URL;
import java.net.URLConnection;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
//import java.util.ArrayList;
import java.util.Calendar;
//import java.util.List;


public class GenerateOperation_WR {
	 static final String driverName = "oracle.jdbc.driver.OracleDriver";
	 static String dbURL = "";//jdbc:oracle:thin:@172.20.5.22:1521:TBOAF"; 
	 static String userName = ""; 
	 static String userPwd = ""; 
	 static Connection dbConn = null;
	 static ResultSet rs = null;
	 static Statement stat = null;
	 static PreparedStatement pstmt = null;
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
	    
		GenerateOperation_WR a = new GenerateOperation_WR();
		/*
		report_no="A04";
		s_year="94";
		s_month="8";
		isDebug="true";*/
		
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
		       /*if(report_no.equals("WR00")){
                   myURL = new URL(Utility.getProperties("BOAFWebSite")+"WR00InsertDB_bank_no_ALL.jsp?report_no="+report_no+"&s_year="+s_year+"&s_month="+s_month+"&isDebug="+isDebug+"&lguser_id="+lguser_id);
		       }else if(report_no.equals("WR01")){
		           myURL = new URL(Utility.getProperties("BOAFWebSite")+"WR01InsertDB_bank_no_ALL.jsp?report_no="+report_no+"&s_year="+s_year+"&s_month="+s_month+"&isDebug="+isDebug+"&lguser_id="+lguser_id);
		       }else if(report_no.equals("WR02")){
		           myURL = new URL(Utility.getProperties("BOAFWebSite")+"WR02InsertDB_bank_no_ALL.jsp?report_no="+report_no+"&s_year="+s_year+"&s_month="+s_month+"&isDebug="+isDebug+"&lguser_id="+lguser_id);
		       }*/
		       myURL = new URL(Utility.getProperties("BOAFWebSite")+"WRInsertDB_bank_no_ALL.jsp?report_no="+report_no+"&s_year="+s_year+"&s_month="+s_month+"&isDebug="+isDebug+"&lguser_id="+lguser_id);
		       runURL(myURL);
		       
     	}catch (Exception e){
			    System.out.println("exeGenerateOperation Error:"+e.getMessage());
		}
    }	
	private void runURL(URL myURL){
	    String raoInputString;
	    URLConnection raoURL;
	    try{
	    	System.out.println("myURL:"+myURL.toString());
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
}

