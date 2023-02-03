/*
 * Created on 2005/1/19
 *
 * TODO To change the template for this generated file go to
 * Window - Preferences - Java - Code Style - Code Templates
 */
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.text.*;
import java.util.*;
import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;
/**
 * @author 2295
 *
 * TODO To change the template for this generated type comment go to
 * Window - Preferences - Java - Code Style - Code Templates
 */
public class FR043W_Excel {
    public static String createRpt(String S_YEAR,String S_MONTH,String Unit){
		  String errMsg = "";
		  List dbData = null;
		  String sqlCmd = "";
		  Properties B01Data = new Properties();		
		  String acc_code = "";
		  String amt = "";
		try{	
			System.out.println("農業發展基金及農業天然災害救助基金貸放餘額統計.xls");
			File xlsDir = new File(Utility.getProperties("xlsDir"));        
        	File reportDir = new File(Utility.getProperties("reportDir"));       
	    	        
    		if(!xlsDir.exists()){
     			if(!Utility.mkdirs(Utility.getProperties("xlsDir"))){
     		   		errMsg +=Utility.getProperties("xlsDir")+"目錄新增失敗";
     			}    
    		}
    		if(!reportDir.exists()){
     			if(!Utility.mkdirs(Utility.getProperties("reportDir"))){
     		   		errMsg +=Utility.getProperties("reportDir")+"目錄新增失敗";
     			}    
    		}
			FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+"農業發展基金及農業天然災害救助基金貸放餘額統計.xls" );
			
	  	    //設定FileINputStream讀取Excel檔
	  		POIFSFileSystem fs = new POIFSFileSystem( finput );
	  		HSSFWorkbook wb = new HSSFWorkbook(fs);
	  		HSSFSheet sheet = wb.getSheetAt(0);//讀取第一個工作表，宣告其為sheet 
	  		HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
	        //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	        //sheet.setAutobreaks(true); //自動分頁
			
	        //設定頁面符合列印大小
	        sheet.setAutobreaks( false );
	        ps.setScale( ( short )65 ); //列印縮放百分比
                ps.setLandscape( true ); // 設定橫式
	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	  		//wb.setSheetName(0,"test");
	  		finput.close();
	  		
	  		HSSFRow row=null;//宣告一列 
	  		HSSFCell cell=null;//宣告一個儲存格  		
	  		
	  		short i=0;
	  		short y=0;	  
	  		
	        	sqlCmd = " select m_year,m_month,a.run_master_no,a.run_sub_no,a.run_next_no,"
	        	       + " (loan_cnt_year/" + Unit
	        	       + " ) as loan_cnt_year,"
	        	       + " (loan_amt_year/" + Unit
	        	       + " ) as loan_amt_year,"
	        	       + " (loan_cnt_totacc/" + Unit
	        	       + " ) as loan_cnt_totacc,"
	        	       + " (loan_amt_totacc/" + Unit
	        	       + " ) as loan_amt_totacc,"
	        	       + " (loan_cnt_bal/" + Unit
	        	       + " ) as loan_cnt_bal,"
	        	       + " (loan_amt_bal_subtot/" + Unit
	        	       + " ) as loan_amt_bal_subtot,"
	        	       + " (loan_amt_bal_fund/" + Unit
	        	       + " ) as loan_amt_bal_fund,"
	        	       + " (loan_amt_bal_bank/" + Unit
	        	       + " ) as loan_amt_bal_bank"
	        	       + " from   B02 a,b00_run_item b "
		               + " where  a.run_master_no = b.run_master_no "
		               + " and    a.run_sub_no = b.run_sub_no "
		               + " and    a.run_next_no = b.run_next_no " 
		               + " and    m_year = " + S_YEAR
		               + " and    m_month = " + S_MONTH
		               + " order by input_order " ;
		        dbData = DBManager.QueryDB(sqlCmd,"m_year,m_month,run_master_no,run_sub_no,run_next_no,loan_cnt_year,loan_amt_year,loan_cnt_totacc,loan_amt_totacc,loan_cnt_bal,loan_amt_bal_subtot,loan_amt_bal_fund,loan_amt_bal_bank");		

	       	
			
                if(dbData.size() == 0){	
	  		row=sheet.getRow(2);
	  		cell=row.getCell((short)2);	       	
	  		//設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
   	       	        cell.setCellValue(S_YEAR +"年" +S_MONTH +"月無資料存在");
	  	}else{
	  		row=sheet.getRow(2);
	  		cell=row.getCell((short)2);	       	
	  		//設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
   	       	        cell.setCellValue("農業發展基金及農業天然災害救助基金貸放餘額統計");
   	       	        
	  		row=sheet.getRow(3);
	  		cell=row.getCell((short)3);	       	
	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	        	cell.setCellValue("中 華 民 國"+S_YEAR+"年底止");
	        	
	  		row=sheet.getRow(3);
	  		cell=row.getCell((short)8);	       	
	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
   	       	        System.out.println("Unit ="+Unit);
   	       	        if (Unit.equals("1")){
   	       	           cell.setCellValue("元");
   	       	        }else if (Unit.equals("1000")){
   	       	           cell.setCellValue("千元");
   	       	        }else if (Unit.equals("10000")){
   	       	           cell.setCellValue("萬元");
   	       	        }else if (Unit.equals("1000000")){
   	       	           cell.setCellValue("百萬元");
   	       	        }else if (Unit.equals("10000000")){
   	       	           cell.setCellValue("千萬元");
   	       	        }else if (Unit.equals("100000000")){
   	       	           cell.setCellValue("億元");
   	       	        }
   	       	        
	  		//以巢狀迴圈讀取所有儲存格資料 
	  		System.out.println("total row ="+sheet.getLastRowNum());
	  		for(i=8;i<31;i++){  
	  			System.out.println("I ="+i);
	  			  		
	    			row=sheet.getRow(i);	    		
	       			cell=row.getCell((short)2);
	       			String loan_cnt_year = Utility.setCommaFormat((((DataObject)dbData.get(i-8)).getValue("loan_cnt_year")).toString());
	       			cell.setCellValue(loan_cnt_year);
	       			
	       			cell=row.getCell((short)3);
	       			String loan_amt_year = Utility.setCommaFormat((((DataObject)dbData.get(i-8)).getValue("loan_amt_year")).toString());	       			       		
	       			cell.setCellValue(loan_amt_year);
	       			
	       			cell=row.getCell((short)4);	       			 
	       			String loan_cnt_totacc = Utility.setCommaFormat((((DataObject)dbData.get(i-8)).getValue("loan_cnt_totacc")).toString());
	       			cell.setCellValue(loan_cnt_totacc);	       		
	       			
	       			cell=row.getCell((short)5);
	       			String loan_amt_totacc = Utility.setCommaFormat((((DataObject)dbData.get(i-8)).getValue("loan_amt_totacc")).toString());
	       			cell.setCellValue(loan_amt_totacc);
	       			
	       			cell=row.getCell((short)6);
	       			String loan_cnt_bal = Utility.setCommaFormat((((DataObject)dbData.get(i-8)).getValue("loan_cnt_bal")).toString());
	       			cell.setCellValue(loan_cnt_bal);
	       			
	       			cell=row.getCell((short)7);
	       			String loan_amt_bal_subtot = Utility.setCommaFormat((((DataObject)dbData.get(i-8)).getValue("loan_amt_bal_subtot")).toString());
	       			cell.setCellValue(loan_amt_bal_subtot);
	       			
	       			cell=row.getCell((short)9);
	       			String loan_amt_bal_fund = Utility.setCommaFormat((((DataObject)dbData.get(i-8)).getValue("loan_amt_bal_fund")).toString());
	       			cell.setCellValue(loan_amt_bal_fund);
	       			
	       			cell=row.getCell((short)10);
	       			String loan_amt_bal_bank = Utility.setCommaFormat((((DataObject)dbData.get(i-8)).getValue("loan_amt_bal_bank")).toString());
	       			cell.setCellValue(loan_amt_bal_bank);
	  		}
	  		
	  	}

	  		
	        FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+"農業發展基金及農業天然災害救助基金貸放餘額統計.xls");
	        wb.write(fout);
	        //儲存 
	        fout.close();
		}catch(Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}		 
}
