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
public class FR042W_Excel {
    public static String createRpt(String S_YEAR,String S_MONTH,String Unit){
		  String errMsg = "";
		  List dbData = null;
		  String sqlCmd = "";
		  Properties B01Data = new Properties();		
		  String acc_code = "";
		  String amt = "";
		try{	
			System.out.println("農業發展基金及農業天然災害基金貸款執行情形表.xls");
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
			FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+"農業發展基金及農業天然災害基金貸款執行情形表.xls" );
			
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

	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	  		//wb.setSheetName(0,"test");
	  		finput.close();
	  	
	  		HSSFRow row=null;//宣告一列 
	  		HSSFCell cell=null;//宣告一個儲存格  		
	  		
	  		short i=0;
	  		short y=0;	  
	  		
	  	
         		sqlCmd =  " select m_year,m_month,"
         		       +  " (budget_amt/" + Unit
         		       +  " ) as budget_amt,"
         		       +  " (credit_pay_amt/" + Unit
         		       +  " ) as credit_pay_amt,"
         		       +  " credit_pay_rate"
	         	       +  " from   B01 a,b00_fund_item b "
		               +  " where  a.fund_master_no = b.fund_master_no "
		               +  " and    a.fund_sub_no = b.fund_sub_no "
		               +  " and    a.fund_next_no = b.fund_next_no " 
		               +  " and    m_year = " + S_YEAR
		               +  " and    m_month = " + S_MONTH
		               +  " order by input_order " ;
	         	dbData = DBManager.QueryDB(sqlCmd,"m_year,m_month,budget_amt,credit_pay_amt,credit_pay_rate");		

                if(dbData.size() == 0){	
	  		row=sheet.getRow(0);
	  		cell=row.getCell((short)0);	       	
	  		//設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
   	       	        cell.setCellValue(S_YEAR +"年" +S_MONTH +"月無資料存在");
	  	}else{
	  		row=sheet.getRow(0);
	  		cell=row.getCell((short)0);	       	
	  		//設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
   	       	        cell.setCellValue("農業發展基金及農業天然災害基金貸款執行情形表");
   	       	        
	  		row=sheet.getRow(1);
	  		cell=row.getCell((short)1);	       	
	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	        	cell.setCellValue(S_YEAR +"年"+ S_MONTH +"月底止");	       	
			
	  		row=sheet.getRow(1);
	  		cell=row.getCell((short)4);	       	
	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
   	       	        System.out.println("Unit ="+Unit);
   	       	        if (Unit.equals("1")){
   	       	           cell.setCellValue("單位:元");
   	       	        }else if (Unit.equals("1000")){
   	       	           cell.setCellValue("單位:千元");
   	       	        }else if (Unit.equals("10000")){
   	       	           cell.setCellValue("單位:萬元");
   	       	        }else if (Unit.equals("1000000")){
   	       	           cell.setCellValue("單位:百萬元");
   	       	        }else if (Unit.equals("10000000")){
   	       	           cell.setCellValue("單位:千萬元");
   	       	        }else if (Unit.equals("100000000")){
   	       	           cell.setCellValue("單位:億元");
   	       	        }
	        	
	  		//以巢狀迴圈讀取所有儲存格資料 
	  		System.out.println("total row ="+sheet.getLastRowNum());
	  		for(i=3;i<(41);i++){
	  		    System.out.println("I ="+i);

	  		    row=sheet.getRow(i);
	       		    cell=row.getCell((short)1);
	       		    String budget_amt = Utility.setCommaFormat((((DataObject)dbData.get(i-3)).getValue("budget_amt")).toString());
	       		    /*
	       		    if (((DataObject)dbData.get(i-3)).getValue("budget_amt") != null && 
	       		   		!((((DataObject)dbData.get(i-3)).getValue("budget_amt")).toString()).equals("") ){
	       		    	budget_amt = (((DataObject)dbData.get(i-3)).getValue("budget_amt")).toString();
	       		    }*/
	       		    System.out.println("budget_amt ="+budget_amt);
	       		    cell.setCellValue(budget_amt);
	       		
	       	            cell=row.getCell((short)2);
	       		    String credit_pay_amt = Utility.setCommaFormat((((DataObject)dbData.get(i-3)).getValue("credit_pay_amt")).toString());
	       		    System.out.println("credit_pay_amt ="+credit_pay_amt);
	       		    cell.setCellValue(credit_pay_amt);
	       		
	       		    cell=row.getCell((short)3);	       			 
	       		    String credit_pay_rate = (((DataObject)dbData.get(i-3)).getValue("credit_pay_rate")).toString();
	       		    System.out.println("credit_pay_rate ="+credit_pay_rate);
	       		    cell.setCellValue(credit_pay_rate);	       		
	       		
	       		    cell=row.getCell((short)4);
	       		    String remark = (String)((DataObject)dbData.get(i-3)).getValue("remark");
	       		    System.out.println("remark ="+remark);
	       		    cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
	       		    if(remark == null || !remark.equals("")){
	       		       cell.setCellValue(remark);
	       		    }else{
	       		       cell.setCellValue("");	
	       		    }
	  		}
	  	}

	  		
	        FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+"農業發展基金及農業天然災害基金貸款執行情形表.xls");
	        wb.write(fout);
	        //儲存 
	        fout.close();
		}catch(Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}		 
}
