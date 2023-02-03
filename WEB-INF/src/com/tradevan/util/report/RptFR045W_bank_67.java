/*
 *  Created on 2007/07/23 存款帳戶分級差異化管理統計表_農漁會 by 2295
 *  96.12.17 fix 修改說明文字 by 2295
 * 100.02.16 fix 根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
 * 				 使用PreparedStatement;並列印轉換後的SQL;套用QueryDB_SQLParam by 2295
 */

package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR045W_bank_67 {
  public static String createRpt(String S_YEAR, String S_MONTH,String E_YEAR,String E_MONTH,String bank_code ) {    

    String errMsg = "";
    List dbData = null;
    String sqlCmd = "";    
    int rowNum=0;
    DataObject bean = null;
    reportUtil reportUtil = new reportUtil();
	HSSFCellStyle cs_right = null; 
	HSSFCellStyle cs_center = null;
   
    String data_year="";
    String data_month="";
    String bank_name="";
    String warnaccount_cnt="";
	String limitaccount_cnt="";
	String erroraccount_cnt="";
	String otheraccount_cnt="";
	String depositaccount_tcnt="";
	String last_warnaccount_cnt="";
	String last_limitaccount_cnt="";
	String last_erroraccount_cnt="";
	String last_otheraccount_cnt="";
	String last_depositaccount_tcnt="";
	String now_warnaccount_cnt="";
	String now_limitaccount_cnt="";
	String now_erroraccount_cnt="";
	String now_otheraccount_cnt="";
	String now_depositaccount_tcnt="";
	String bn01_m_year = "";
	List paramList = new ArrayList();
    try {
      //100.02.16 add 查詢年度100年以前.縣市別不同===============================  	     
   	  bn01_m_year = (Integer.parseInt(S_YEAR) < 100)?"99":"100"; 
   	  //=====================================================================
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

      //input the standard report form      
      finput = new FileInputStream(xlsDir +System.getProperty("file.separator") +"存款帳戶分級差異化管理統計表_農漁會.xls");

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
      
      int m_year = Integer.parseInt(S_YEAR);
      int m_month = Integer.parseInt(S_MONTH);      
      int last_year = m_year;
      int last_month = m_month - 1;
      int quarter_year = m_year;
      int quarter_month = m_month;
      if (last_month == 0) { //本月如果是1月的話 上月為12月
        last_month = 12;
        last_year--;
      }
      
   	  /*
      sqlCmd = " select nowA08.bank_code,bn01.bank_name,nowA08.m_year,nowA08.m_month,"
       		 + "	    nowA08.warnaccount_cnt,"
       		 + "	    nowA08.warnaccount_cnt - lastA08.warnaccount_cnt as last_warnaccount_cnt,"        
       		 + "	 	nowA08.limitaccount_cnt,"
       		 + "	 	nowA08.limitaccount_cnt - lastA08.limitaccount_cnt as last_limitaccount_cnt,"       
       		 + " 		nowA08.erroraccount_cnt,"
       		 + " 		nowA08.erroraccount_cnt - lastA08.erroraccount_cnt as last_erroraccount_cnt,"        
       		 + "	 	nowA08.otheraccount_cnt,"
       		 + " 		nowA08.otheraccount_cnt - lastA08.otheraccount_cnt as last_otheraccount_cnt,"       
       		 + " 		nowA08.depositaccount_tcnt,"
       		 + "	 	nowA08.depositaccount_tcnt - lastA08.depositaccount_tcnt as last_depositaccount_tcnt"       		    
			 + " from(select * from a08 where m_year="+m_year+" and m_month="+m_month+")nowA08"
             + " left join (select * from a08 where m_year="+last_year+" and m_month="+last_month+")lastA08 on nowA08.bank_code = lastA08.bank_code"
			 + " left join bn01 on nowA08.bank_code = bn01.bank_no"; 
      */
      sqlCmd = " select a08.*,bank_name"
      	     + " from a08 left join (select * from bn01 where m_year=?)bn01 on a08.bank_code = bn01.bank_no"
			 + " where bank_code=?"
			 + " and to_char(a08.m_year * 100 + m_month) >= ?"
			 + " and to_char(a08.m_year * 100 + m_month) <= ?"	
			 + " order by a08.m_year,m_month ";
      paramList.add(bn01_m_year);
      paramList.add(bank_code);
      paramList.add((last_year < 100 ?"0":"")+last_year+(last_month < 10 ?"0":"")+last_month );
      paramList.add(E_YEAR+E_MONTH);
   
	  
	  

      dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"m_year,m_month,warnaccount_cnt,limitaccount_cnt,erroraccount_cnt,otheraccount_cnt,depositaccount_tcnt");

      
      System.out.println("dbData.size=" + dbData.size());
     
      //設定報表表頭資料============================================
      
      row=sheet.getRow(1);
      cell=row.getCell((short)0);	       	
      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      cell.setCellValue("民國"+S_YEAR+"年"+S_MONTH+"月底~"+E_YEAR+"年"+E_MONTH+"月底"+((dbData == null || dbData.size() ==0)?"無資料存在":""));  	
      
      row = sheet.getRow(2);
      cell = row.getCell( (short) 1);
      cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
      if(dbData != null && dbData.size() != 0){
      	 bean = (DataObject)dbData.get(0);
      	 bank_name = (String)bean.getValue("bank_name");
      }
      cell.setCellValue(bank_name);
      
     
      if (dbData == null || dbData.size() ==0) {      	      
      }else {
      	rowNum = 5;
      	bean = (DataObject)dbData.get(0);
      	data_year=(bean.getValue("m_year") == null)?"":(bean.getValue("m_year")).toString();
    	data_month=(bean.getValue("m_month") == null)?"":(bean.getValue("m_month")).toString();
    	if(Integer.parseInt(data_year) == last_year && Integer.parseInt(data_month)==last_month){
    		last_warnaccount_cnt=(bean.getValue("warnaccount_cnt") == null)?"":(bean.getValue("warnaccount_cnt")).toString();
        	last_limitaccount_cnt=(bean.getValue("limitaccount_cnt") == null)?"":(bean.getValue("limitaccount_cnt")).toString();
        	last_erroraccount_cnt=(bean.getValue("erroraccount_cnt") == null)?"":(bean.getValue("erroraccount_cnt")).toString();
        	last_otheraccount_cnt=(bean.getValue("otheraccount_cnt") == null)?"":(bean.getValue("otheraccount_cnt")).toString();
        	last_depositaccount_tcnt=(bean.getValue("depositaccount_tcnt") == null)?"":(bean.getValue("depositaccount_tcnt")).toString();        		
    	}else{
    		last_warnaccount_cnt = "0";
    		last_limitaccount_cnt="0";
    		last_erroraccount_cnt="0";
    		last_otheraccount_cnt="0";
    		last_depositaccount_tcnt="0";
    	}
    	
        for(int idx=0;idx<dbData.size();idx++){
        	//System.out.println("idx="+idx);
        	bean = (DataObject)dbData.get(idx);
        	data_year=(bean.getValue("m_year") == null)?"":(bean.getValue("m_year")).toString();
        	data_month=(bean.getValue("m_month") == null)?"":(bean.getValue("m_month")).toString();
        	warnaccount_cnt=(bean.getValue("warnaccount_cnt") == null)?"":(bean.getValue("warnaccount_cnt")).toString();
        	limitaccount_cnt=(bean.getValue("limitaccount_cnt") == null)?"":(bean.getValue("limitaccount_cnt")).toString();
        	erroraccount_cnt=(bean.getValue("erroraccount_cnt") == null)?"":(bean.getValue("erroraccount_cnt")).toString();
        	otheraccount_cnt=(bean.getValue("otheraccount_cnt") == null)?"":(bean.getValue("otheraccount_cnt")).toString();
        	depositaccount_tcnt=(bean.getValue("depositaccount_tcnt") == null)?"":(bean.getValue("depositaccount_tcnt")).toString();
        	if(Integer.parseInt(data_year) == last_year && Integer.parseInt(data_month)==last_month){
        	   continue;
        	}
        	row = sheet.createRow(rowNum);
		    for(int cellcount=0;cellcount<6;cellcount++){
	 		    cell = row.createCell( (short)cellcount);
	     		cell.setEncoding(HSSFCell.ENCODING_UTF_16);	 	
	     		if(cellcount == 0 ){
	     		   cell.setCellStyle(cs_center);
	     		}else{
	     		   cell.setCellStyle(cs_right);
	     		}
	     		if(cellcount == 0) cell.setCellValue(data_year+"年"+data_month+"月底戶數");		        
	     		if(cellcount == 1) cell.setCellValue(Utility.setCommaFormat(warnaccount_cnt));	 
	     		if(cellcount == 2) cell.setCellValue(Utility.setCommaFormat(limitaccount_cnt));	 
	     		if(cellcount == 3) cell.setCellValue(Utility.setCommaFormat(erroraccount_cnt));
	     		if(cellcount == 4) cell.setCellValue(Utility.setCommaFormat(otheraccount_cnt));
	     		if(cellcount == 5) cell.setCellValue(Utility.setCommaFormat(depositaccount_tcnt));	     		
			}//end of cellcount	 		  	  	 	
		    rowNum++;
		    row = sheet.createRow(rowNum);
		    now_warnaccount_cnt = warnaccount_cnt;
     		now_limitaccount_cnt = limitaccount_cnt;
     		now_erroraccount_cnt = erroraccount_cnt;
     		now_otheraccount_cnt = otheraccount_cnt;
     		now_depositaccount_tcnt = depositaccount_tcnt;     		
     		warnaccount_cnt = String.valueOf(Integer.parseInt(warnaccount_cnt)-Integer.parseInt(last_warnaccount_cnt));     		
     		limitaccount_cnt = String.valueOf(Integer.parseInt(limitaccount_cnt)-Integer.parseInt(last_limitaccount_cnt));     		
     		erroraccount_cnt = String.valueOf(Integer.parseInt(erroraccount_cnt)-Integer.parseInt(last_erroraccount_cnt));
     		otheraccount_cnt = String.valueOf(Integer.parseInt(otheraccount_cnt)-Integer.parseInt(last_otheraccount_cnt));
     		depositaccount_tcnt = String.valueOf(Integer.parseInt(depositaccount_tcnt)-Integer.parseInt(last_depositaccount_tcnt));
     		last_warnaccount_cnt = now_warnaccount_cnt;
     		last_limitaccount_cnt = now_limitaccount_cnt;
     		last_erroraccount_cnt = now_erroraccount_cnt;
     		last_otheraccount_cnt = now_otheraccount_cnt;
     		last_depositaccount_tcnt = now_depositaccount_tcnt;
    	    for(int cellcount=0;cellcount<6;cellcount++){
	 		    cell = row.createCell( (short)cellcount);
	     		cell.setEncoding(HSSFCell.ENCODING_UTF_16);	 	
	     		if(cellcount == 0 ){
	     		   cell.setCellStyle(cs_center);
	     		}else{
	     		   cell.setCellStyle(cs_right);
	     		}		     		 		
	     		if(cellcount == 0) cell.setCellValue("與上月底比較增減戶數(請以+,-表示)");		        
	     		if(cellcount == 1) cell.setCellValue(warnaccount_cnt.startsWith("-")?warnaccount_cnt:"+"+Utility.setCommaFormat(warnaccount_cnt));	 
	     		if(cellcount == 2) cell.setCellValue(limitaccount_cnt.startsWith("-")?limitaccount_cnt:"+"+Utility.setCommaFormat(limitaccount_cnt));	 
	     		if(cellcount == 3) cell.setCellValue(erroraccount_cnt.startsWith("-")?erroraccount_cnt:"+"+Utility.setCommaFormat(erroraccount_cnt));
	     		if(cellcount == 4) cell.setCellValue(otheraccount_cnt.startsWith("-")?otheraccount_cnt:"+"+Utility.setCommaFormat(otheraccount_cnt));
	     		if(cellcount == 5) cell.setCellValue(depositaccount_tcnt.startsWith("-")?depositaccount_tcnt:"+"+Utility.setCommaFormat(depositaccount_tcnt));	     		
			}//end of cellcount	 		  	  	 	
		    rowNum++;
        }//end of dbData        
      } //end of else dbData.size() != 0
      
      row=sheet.createRow(rowNum);
      cell=row.createCell((short)0);	       	
      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      cell.setCellValue("備註：");
      rowNum++;
      row=sheet.createRow(rowNum);
      cell=row.createCell((short)0);	       	
      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      cell.setCellValue("1.存款帳戶指支票存款、活期存款、活期儲蓄存款及定期存款帳戶。");
      rowNum++;
      row=sheet.createRow(rowNum);
      cell=row.createCell((short)0);	       	
      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      cell.setCellValue("2.本表由農漁會信用部於每月終了次月15日前上線填報");      
      rowNum++;
      row=sheet.createRow(rowNum);
      cell=row.createCell((short)0);	       	
      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      cell.setCellValue("    （原第一次書面填報為95年9月底資料）。");     
      
      FileOutputStream fout = null;     
      fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "存款帳戶分級差異化管理統計表_農漁會.xls");
     
      HSSFFooter footer = sheet.getFooter();
      footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
      footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
      wb.write(fout);
      //儲存
      fout.close();
    }
    catch (Exception e) {
      System.out.println("RptFR045W_bank_b.createRpt Error:" + e + e.getMessage());
    }
    
    return errMsg;
  }
}
