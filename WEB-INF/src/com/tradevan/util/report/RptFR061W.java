/*
 *101.07.30 create  by 2968
 *106.10.16 add [M]台中商業銀行 by 2295
 */
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import java.io.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR061W {
	 	
	 public static String createRpt(String M_YEAR, String M_MONTH,String unit,String bank_code ) {    

	    String errMsg = "";
	    List dbData = null;
	    String sqlCmd = "";    
	    int rowNum=0;
	    DataObject bean = null;
	    reportUtil reportUtil = new reportUtil();
		HSSFCellStyle cs_right = null; 
		HSSFCellStyle cs_center = null;
		HSSFCellStyle cs_left = null;
	   
	  	StringBuffer sql = new StringBuffer() ;
	  	ArrayList paramList = new ArrayList() ;
	  	String u_year = "100" ;
	  	if(M_YEAR==null || Integer.parseInt(M_YEAR)<=99 ) {
	  		u_year ="99" ;
	  	}
	    try {

	      File xlsDir = new File(Utility.getProperties("xlsDir"));
	      File reportDir = new File(Utility.getProperties("reportDir"));

	      if (!xlsDir.exists()) {
	        if (!Utility.mkdirs(Utility.getProperties("xlsDir"))) {
	          errMsg += Utility.getProperties("xlsDir") + "目錄新增失敗";
	        }
	      }
	      if (!reportDir.exists()) {
	        if (!Utility.mkdirs(Utility.getProperties("reportDir"))) {
	          errMsg += Utility.getProperties("reportDir") + "目錄新增失敗";
	        }
	      }
	      FileInputStream finput = null;
	      String filename="M206_金融機構別及地區別保證案件分析表.xls";
	      //input the standard report form      
	      finput = new FileInputStream(xlsDir +System.getProperty("file.separator") +filename);
	      //設定FileINputStream讀取Excel檔
	      POIFSFileSystem fs = new POIFSFileSystem(finput);
	      HSSFWorkbook wb = new HSSFWorkbook(fs);
	      HSSFSheet sheet = wb.getSheetAt(0); //讀取第一個工作表，宣告其為sheet
	      HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
	      //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	      //sheet.setAutobreaks(true); //自動分頁

	      //設定頁面符合列印大小
	      sheet.setAutobreaks(false);
	      ps.setScale( (short) 100); //列印縮放百分比

	      ps.setPaperSize( (short) 9); //設定紙張大小 A4
	      //wb.setSheetName(0,"test");
	      finput.close();

	      HSSFRow row = null; //宣告一列
	      HSSFCell cell = null; //宣告一個儲存格

	      short i = 0;
	      short y = 0;
	      cs_right = reportUtil.getRightStyle(wb);
	      cs_center = reportUtil.getDefaultStyle(wb);
	      cs_left = reportUtil.getLeftStyle(wb);
	      
	      sql.append(" select m206.loan_unit,M206_LOAN_UNIT.in_loan_name,"
	      		 + "          guarantee_cnt_year,"
				 + "          round(guarantee_amt_year / ? ,0)  as guarantee_amt_year,"
				 + "          round(loan_amt_year / ? ,0)  as loan_amt_year,"
				 + "          guarantee_cnt_sum,"
	      	     + "          round(guarantee_amt_sum / ?,0)  as guarantee_amt_sum,"
				 + "          round(loan_amt_sum / ?,0)  as loan_amt_sum,"
				 + "          round(guarantee_bal_year / ?,0)  as guarantee_bal_year "
				 + "   from m206 "
				 + "   left join m206_loan_unit on M206.LOAN_UNIT = M206_LOAN_UNIT.LOAN_UNIT_NO, "
				 + "   (select m_year,m_month,count(*) from m206_loan_unit ");
	      if(Integer.parseInt(M_YEAR) *100 +  Integer.parseInt(M_MONTH) >= 10605 ){//106.10.16 add
	          sql.append(" where m_year=? and m_month >= ?");
	      }
	      sql.append("    group by m_year,m_month)loan_unit_year "
				 + "   where m206.m_year*100+m206.m_month >= loan_unit_year.m_year *100+loan_unit_year.m_month "
				 + "         and m206.m_year= ? "
				 + "         and m206.m_month= ? "
				 + "         and loan_unit_year.m_year = m206_loan_unit.m_year "
				 + "         and loan_unit_year.m_month = m206_loan_unit.m_month "
				 + "   order by to_number(m206_loan_unit.input_order)" );
	      
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);
	      if(Integer.parseInt(M_YEAR) *100 +  Integer.parseInt(M_MONTH) >= 10605 ){//106.05增加台中商銀
              paramList.add("106");
              paramList.add("05");
          }
	      paramList.add(M_YEAR);
	      paramList.add(M_MONTH);
	     
	      dbData =  DBManager.QueryDB_SQLParam(sql.toString(), paramList, "loan_unit,in_loan_name,guarantee_cnt_year,guarantee_amt_year,loan_amt_year,guarantee_cnt_sum,guarantee_amt_sum,loan_amt_sum,guarantee_bal_year");
	      System.out.println("dbData.size=" + dbData.size());
	     
	      //設定報表表頭資料============================================
	      row=sheet.getRow(3);
	      cell=row.getCell((short)1);	       	
	      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	      cell.setCellValue("中華民國"+M_YEAR+"年"+M_MONTH+"月"+((dbData == null || dbData.size() ==0)?"無資料存在":""));  	
	      
	      if (dbData != null || dbData.size() !=0) { 
            rowNum = 7;
            for (int k=0; k<dbData.size();k++){
	            bean = (DataObject)dbData.get(k);
                String loan_unit=(bean.getValue("loan_unit") == null)?"":(bean.getValue("loan_unit")).toString();
                String in_loan_name=(bean.getValue("in_loan_name") == null)?"":(bean.getValue("in_loan_name")).toString();
                String guarantee_cnt_year=(bean.getValue("guarantee_cnt_year") == null)?"0":(bean.getValue("guarantee_cnt_year")).toString();
                String guarantee_amt_year=(bean.getValue("guarantee_amt_year") == null)?"0":(bean.getValue("guarantee_amt_year")).toString();
                String loan_amt_year =(bean.getValue("loan_amt_year") == null)?"0":(bean.getValue("loan_amt_year")).toString();
                String guarantee_cnt_sum=(bean.getValue("guarantee_cnt_sum") == null)?"0":(bean.getValue("guarantee_cnt_sum")).toString();
                String guarantee_amt_sum=(bean.getValue("guarantee_amt_sum") == null)?"0":(bean.getValue("guarantee_amt_sum")).toString();
                String loan_amt_sum=(bean.getValue("loan_amt_sum") == null)?"0":(bean.getValue("loan_amt_sum")).toString();
                String guarantee_bal_year=(bean.getValue("guarantee_bal_year") == null)?"0":(bean.getValue("guarantee_bal_year")).toString();
                System.out.println(in_loan_name+", "+guarantee_cnt_year+", "+guarantee_amt_year+", "+loan_amt_year+", "+guarantee_cnt_sum+", "+guarantee_amt_sum+", "+loan_amt_sum+", "+guarantee_bal_year);
                
                row = sheet.createRow(rowNum);
                
                //貸款機構別
                cell = row.createCell( (short)1);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellStyle(cs_left);                    
                cell.setCellValue(in_loan_name);
                
                
                //本年度保證案件件數
                cell = row.createCell( (short)2);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellStyle(cs_right);                    
                cell.setCellValue(Utility.setCommaFormat(guarantee_cnt_year));
               
                //本年度保證案件保證金額
                cell = row.createCell( (short)3);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellStyle(cs_right);                    
                cell.setCellValue(Utility.setCommaFormat(guarantee_amt_year));
                //本年度保證案件融資金額
                cell = row.createCell( (short)4);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellStyle(cs_right);                    
                cell.setCellValue(Utility.setCommaFormat(loan_amt_year));
                //累計件數
                cell = row.createCell( (short)5);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellStyle(cs_right);                    
                cell.setCellValue(Utility.setCommaFormat(guarantee_cnt_sum));
                //累計保證金額
                cell = row.createCell( (short)6);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellStyle(cs_right);                    
                cell.setCellValue(Utility.setCommaFormat(guarantee_amt_sum));
                //累計融資金額
                cell = row.createCell( (short)7);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellStyle(cs_right);                    
                cell.setCellValue(Utility.setCommaFormat(loan_amt_sum));
                //本年度保證餘額
                cell = row.createCell( (short)8);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellStyle(cs_right);                    
                cell.setCellValue(Utility.setCommaFormat(guarantee_bal_year));
               
                rowNum++;
                
                
	      	} // end of for
	      } //end of if
	      
	      
	      FileOutputStream fout = null;     
	      fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + filename);
	     
	      HSSFFooter footer = sheet.getFooter();
	      footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
	      footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
	      wb.write(fout);
	      //儲存
	      fout.close();
	    }
	    catch (Exception e) {
	    	System.out.println("RptFR061W.createRpt Error:" + e + e.getMessage());
	    }
	    
	    return errMsg;
	  }
}
