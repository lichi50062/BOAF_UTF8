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
public class FR041W_Excel {
    public static String createRpt(String S_YEAR,String S_MONTH,String Unit){
		  String errMsg = "";
		  List dbData1 = null;
		  List dbData2 = null;
		  List dbData3 = null;
		  List dbData4 = null;
		  String sqlCmd = "";
		  Properties B01Data = new Properties();		
		  String acc_code = "";
		  String amt = "";
		try{	
			System.out.println("農業發展基金貸款有關統計資料表.xls");
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
			FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+"農業發展基金貸款有關統計資料表.xls" );
			
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
	  		short j=0;
	  		short k=0;
	  		short l=0;
	  		short y=0;	  
	  		
                        sqlCmd 	=   " select m_year,m_month,"
                                +   " (loan_cnt_totacc/" + Unit
                                +   " ) as loan_cnt_totacc,"
                                +   " (loan_amt_totacc_fund/" + Unit
                                +   " ) as loan_amt_totacc_fund,"
                                +   " (loan_amt_totacc_bank/" + Unit
                                +   " ) as loan_amt_totacc_bank,"
                                +   " (loan_amt_totacc_tot/" + Unit
                                +   " ) as loan_amt_totacc_tot,"
                                +   " (loan_cnt_bal/" + Unit
                                +   " ) as loan_cnt_bal,"
                                +   " (loan_amt_bal_fund/" + Unit
                                +   " ) as loan_amt_bal_fund,"
                                +   " (loan_amt_bal_bank/" + Unit
                                +   " ) as loan_amt_bal_bank,"
                                +   " (loan_amt_bal_tot/" + Unit
                                +   " ) as loan_amt_bal_tot"
                                +   " from B03_1 a,b00_funs_item b "
                                +   " where a.funs_master_no=b.funs_master_no "
    				+   " and   a.funs_sub_no=b.funs_sub_no "
    				+   " and   a.funs_next_no=b.funs_next_no "
    				+   " and   m_year=" + S_YEAR
    				+   " and   m_month=" + S_MONTH
    				+   " order by input_order ";
			dbData1 = DBManager.QueryDB(sqlCmd,"m_year,m_month,loan_cnt_totacc,loan_amt_totacc_fund,loan_amt_totacc_bank,loan_amt_totacc_tot,loan_cnt_bal,loan_amt_bal_fund,loan_amt_bal_bank,loan_amt_bal_tot");
	  		
	  		sqlCmd  =   " select m_year,m_month,"
	  		        +   " (loan_amt_bal/" + Unit
	  		        +   " ) as loan_amt_bal,"
	  		        +   " (loan_amt_over/" + Unit
	  		        +   " ) as loan_amt_over,"
	  		        +   " loan_rate_over"
	  		        +   " from B03_2 a,b00_funs_item b "
	  		        +   " where a.funs_master_no=b.funs_master_no "
    				+   " and   a.funs_sub_no=b.funs_sub_no "
    				+   " and   a.funs_next_no=b.funs_next_no "
    				+   " and   m_year=" + S_YEAR
    				+   " and   m_month=" + S_MONTH
    				+   " order by input_order";
			dbData2 = DBManager.QueryDB(sqlCmd,"m_year,m_month,loan_amt_bal,loan_amt_over,loan_rate_over");
			
			sqlCmd	=   " select m_year,m_month,"
			        +   " (funo_amt/" + Unit
			        +   " ) as funo_amt,"
			        +   " funo_rate"
			        +   " from B03_3 a,b00_funo_item b "
			        +   " where a.funo_master_no=b.funo_master_no "
    				+   " and   a.funo_sub_no=b.funo_sub_no "
    				+   " and   a.funo_next_no=b.funo_next_no "
    				+   " and   m_year=" + S_YEAR
    			        +   " and   m_month=" + S_MONTH
    			        +   " order by input_order";
			dbData3 = DBManager.QueryDB(sqlCmd,"m_year,m_month,funo_amt,funo_rate");
			
			sqlCmd	=   " select m_year,m_month,"
			        +   " (machine_cnt/" + Unit
			        +   " ) as machine_cnt,"
			        +   " (machine_amt/" + Unit
			        +   " ) as machine_amt,"
			        +   " (land_cnt/" + Unit
			        +   " ) as land_cnt,"
			        +   " (land_amt/" + Unit
			        +   " ) as land_amt,"
			        +   " (house_cnt/" + Unit
			        +   " ) as house_cnt,"
			        +   " (house_amt/" + Unit
			        +   " ) as house_amt,"
			        +   " (build_cnt/" + Unit
			        +   " ) as build_cnt,"
			        +   " (build_amt/" + Unit
			        +   " ) as build_amt,"
			        +   " (tot_cnt/" + Unit
			        +   " ) as tot_cnt,"
			        +   " (tot_amt/" + Unit
			        +   " ) as tot_amt"
			        +   " from B03_4 a,b00_bank_no b "
			        +   " where a.bank_no=b.bank_no "
    				+   " and   m_year=" + S_YEAR
    				+   " and   m_month=" + S_MONTH
    				+   " order by input_order";
			dbData4 = DBManager.QueryDB(sqlCmd,"m_year,m_month,machine_cnt,machine_amt,land_cnt,land_amt,house_cnt,house_amt,build_cnt,build_amt,tot_cnt,tot_amt");
			
                if(dbData1.size() == 0){	
	  		row=sheet.getRow(0);
	  		cell=row.getCell((short)4);	       	
	  		//設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
   	       	        cell.setCellValue(S_YEAR +"年" +S_MONTH +"月無資料存在");
   	       	        
	  	}else{
	  		row=sheet.getRow(0);
	  		cell=row.getCell((short)2);	       	
	  		//設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	        	cell.setCellValue(S_YEAR +"年" + S_MONTH +"月底止");	       	
			
			row=sheet.getRow(0);
	  		cell=row.getCell((short)4);	       	
	  		//設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
	  		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
   	       	        cell.setCellValue("農業發展基金貸款有關統計資料表");
   	       	        
   	       	        row=sheet.getRow(2);
	  		cell=row.getCell((short)7);	       	
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
   	       	        
   	       	        row=sheet.getRow(23);
	  		cell=row.getCell((short)2);	       	
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
   	       	        
   	       	        row=sheet.getRow(23);
	  		cell=row.getCell((short)7);	       	
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
   	       	        
   	       	        row=sheet.getRow(35);
	  		cell=row.getCell((short)9);	       	
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
	  		//農業發展基金貸款統計表
	  		for(i=6;i<21;i++){    		
	    			System.out.println("I ="+i);
	    			row=sheet.getRow(i);	    		
	       			cell=row.getCell((short)1);
	       			//String loan_cnt_totacc = String.valueOf(Integer.parseInt((((DataObject)dbData1.get(i-6)).getValue("loan_cnt_totacc")).toString()*));
	       			String loan_cnt_totacc = Utility.setCommaFormat((((DataObject)dbData1.get(i-6)).getValue("loan_cnt_totacc")).toString());
	       			cell.setCellValue(loan_cnt_totacc);
	       			
	       			cell=row.getCell((short)2);
	       			String loan_amt_totacc_fund = Utility.setCommaFormat((((DataObject)dbData1.get(i-6)).getValue("loan_amt_totacc_fund")).toString());	       			       		
	       			cell.setCellValue(loan_amt_totacc_fund);
	       			
	       			cell=row.getCell((short)3);	       			 
	       			String loan_amt_totacc_bank = Utility.setCommaFormat((((DataObject)dbData1.get(i-6)).getValue("loan_amt_totacc_bank")).toString());
	       			cell.setCellValue(loan_amt_totacc_bank);	       		
	       			
	       			cell=row.getCell((short)4);
	       			String loan_amt_totacc_tot = Utility.setCommaFormat((((DataObject)dbData1.get(i-6)).getValue("loan_amt_totacc_tot")).toString());
	       			cell.setCellValue(loan_amt_totacc_tot);
	       			
	       			cell=row.getCell((short)5);
	       			String loan_cnt_bal = Utility.setCommaFormat((((DataObject)dbData1.get(i-6)).getValue("loan_cnt_bal")).toString());
	       			cell.setCellValue(loan_cnt_bal);
	       			
	       			cell=row.getCell((short)6);
	       			String loan_amt_bal_fund = Utility.setCommaFormat((((DataObject)dbData1.get(i-6)).getValue("loan_amt_bal_fund")).toString());
	       			cell.setCellValue(loan_amt_bal_fund);
	       			
	       			cell=row.getCell((short)7);
	       			String loan_amt_bal_bank = Utility.setCommaFormat((((DataObject)dbData1.get(i-6)).getValue("loan_amt_bal_bank")).toString());
	       			cell.setCellValue(loan_amt_bal_bank);
	       			
	       			cell=row.getCell((short)8);
	       			String loan_amt_bal_tot = Utility.setCommaFormat((((DataObject)dbData1.get(i-6)).getValue("loan_amt_bal_tot")).toString());
	       			cell.setCellValue(loan_amt_bal_tot);
	  		}
	  		
	  		//農業發展基金貸款逾期情形表
	  		for(j=25;j<30;j++){    		
	    			System.out.println("J ="+j);
	    			row=sheet.getRow(j);	    		
	       			cell=row.getCell((short)1);
	       			String loan_amt_bal = Utility.setCommaFormat((((DataObject)dbData2.get(j-25)).getValue("loan_amt_bal")).toString());
	       			cell.setCellValue(loan_amt_bal);
	       			
	       			cell=row.getCell((short)2);
	       			String loan_amt_over = Utility.setCommaFormat((((DataObject)dbData2.get(j-25)).getValue("loan_amt_over")).toString());	       			       		
	       			cell.setCellValue(loan_amt_over);
	       			
	       			cell=row.getCell((short)3);	       			 
	       			String loan_rate_over = (((DataObject)dbData2.get(j-25)).getValue("loan_rate_over")).toString();
	       			cell.setCellValue(loan_rate_over);	       		
	       			
	  		}
	  		
	  		//農業發展基金來源運用表
	  		for(k=25;k<33;k++){    		
	    			System.out.println("K ="+k);
	    			row=sheet.getRow(k);	    		
	       			cell=row.getCell((short)6);
	       			String funo_amt = Utility.setCommaFormat((((DataObject)dbData3.get(k-25)).getValue("funo_amt")).toString());
	       			cell.setCellValue(funo_amt);
	       			
	       			cell=row.getCell((short)7);
	       			String funo_rate = (((DataObject)dbData3.get(k-25)).getValue("funo_rate")).toString();	       			       		
	       			cell.setCellValue(funo_rate);
	       			
	  		}
	  		
	  		//農業發展基金貸款統計表
	  		for(l=38;l<43;l++){    		
	    			System.out.println("L ="+l);
	    			row=sheet.getRow(l);	    		
	       			cell=row.getCell((short)1);
	       			String machine_cnt = Utility.setCommaFormat((((DataObject)dbData4.get(l-38)).getValue("machine_cnt")).toString());
	       			System.out.println("machine_cnt ="+machine_cnt);
	       			cell.setCellValue(machine_cnt);
	       			
	       			cell=row.getCell((short)2);
	       			String machine_amt = Utility.setCommaFormat((((DataObject)dbData4.get(l-38)).getValue("machine_amt")).toString());	       			       		
	       			System.out.println("machine_amt ="+machine_amt);
	       			cell.setCellValue(machine_amt);
	       			
	       			cell=row.getCell((short)3);	       			 
	       			String land_cnt = Utility.setCommaFormat((((DataObject)dbData4.get(l-38)).getValue("land_cnt")).toString());
	       			System.out.println("land_cnt ="+land_cnt);
	       			cell.setCellValue(land_cnt);	       		
	       			
	       			cell=row.getCell((short)4);
	       			String land_amt = Utility.setCommaFormat((((DataObject)dbData4.get(l-38)).getValue("land_amt")).toString());
	       			System.out.println("land_amt ="+land_amt);
	       			cell.setCellValue(land_amt);
	       			
	       			cell=row.getCell((short)5);
	       			String house_cnt = Utility.setCommaFormat((((DataObject)dbData4.get(l-38)).getValue("house_cnt")).toString());
	       			System.out.println("house_cnt ="+house_cnt);
	       			cell.setCellValue(house_cnt);
	       			
	       			cell=row.getCell((short)6);
	       			String house_amt = Utility.setCommaFormat((((DataObject)dbData4.get(l-38)).getValue("house_amt")).toString());
	       			System.out.println("house_amt ="+house_amt);
	       			cell.setCellValue(house_amt);
	       			
	       			cell=row.getCell((short)7);
	       			String build_cnt = Utility.setCommaFormat((((DataObject)dbData4.get(l-38)).getValue("build_cnt")).toString());
	       			System.out.println("build_cnt ="+build_cnt);
	       			cell.setCellValue(build_cnt);
	       			
	       			cell=row.getCell((short)8);
	       			String build_amt = Utility.setCommaFormat((((DataObject)dbData4.get(l-38)).getValue("build_amt")).toString());
	       			System.out.println("build_amt ="+build_amt);
	       			cell.setCellValue(build_amt);
	       			
	       			cell=row.getCell((short)9);
	       			String tot_cnt = Utility.setCommaFormat((((DataObject)dbData4.get(l-38)).getValue("tot_cnt")).toString());
	       			System.out.println("tot_cnt ="+tot_cnt);
	       			cell.setCellValue(tot_cnt);
	       			
	       			cell=row.getCell((short)10);
	       			String tot_amt = Utility.setCommaFormat((((DataObject)dbData4.get(l-38)).getValue("tot_amt")).toString());
	       			System.out.println("tot_amt ="+tot_amt);
	       			cell.setCellValue(tot_amt);
	  		}
	  		
	  	}

	  		
	        FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+"農業發展基金貸款有關統計資料表.xls");
	        wb.write(fout);
	        //儲存 
	        fout.close();
		}catch(Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}		 
}
