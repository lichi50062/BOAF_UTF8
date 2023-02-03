/*
 * 106.12.04 create 金融機構代號轉換清單 by 2295
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

public class RptAS003W1 { 
     
	 public static String createRpt() {    

	    String errMsg = "";	    
	    List dbData = null;	    
	    int rowNum=0;
	    DataObject bean = null;
	    reportUtil reportUtil = new reportUtil();		
	    HSSFCellStyle cs_center = null;
	  
	    String bank_no=""; //--新機構代號
	    String src_bank_no = ""; //--舊機構代號
	    String bank_name = ""; //--總機構名稱
	    String online_date = ""; //--機構代碼轉換日期
	    String trans_date = ""; //--MIS系統轉換日期
   	   
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
	     
	      //input the standard report form      
	      finput = new FileInputStream(xlsDir +System.getProperty("file.separator") +"AS003W金融機構代號轉換清單.xls");      
	      
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
	      
	      cs_center = reportUtil.getDefaultStyle(wb);	     
          
	      StringBuffer sql = new StringBuffer();
	      List paramList = new ArrayList();//傳入參數
	     
	      sql.append(" Select ba01_trans.bank_no ") ;
	      sql.append(",ba01_trans.src_bank_no ") ;
	      sql.append(",bank_name ");
	      sql.append(",to_char(trunc(ba01_trans.trans_date),'YYYY/MM/DD')trans_date");//MIS系統轉換日期
	      sql.append(",to_char(trunc(ba01_trans.online_date),'YYYY/MM/DD')online_date");//106.11.22 add 機構代號轉換日期
	      sql.append(",ba01_trans.pbank_no") ;
	      sql.append(" from ba01_trans left join ( ");
	      sql.append(" Select * from bn01 Where m_year=100 ) bn01 ")  ;
	      sql.append(" on ba01_trans.src_bank_no = bn01.ori_bank_no ") ;
	      sql.append(" Where bank_kind = '0' ") ;	      
	      sql.append(" order by online_date desc,bank_name asc");//106.12.04 add
	     	      
	      dbData = DBManager.QueryDB_SQLParam(sql.toString(),null,"trans_date,online_date");
	      System.out.println("RptAS003W1.dbData.size=" + dbData.size());	      
	      //設定報表表頭資料============================================
	     
          
	      row = sheet.getRow(0);
          cell = row.getCell( (short) 0);
          cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	      if (dbData != null && dbData.size() != 0) {	  
	          bean = (DataObject)dbData.get(0);
	         
		   	  rowNum = 1;	      	  
	      	  for(int i=0;i<dbData.size();i++){	
	      	      bean = (DataObject)dbData.get(i);
	      	      bank_no = (bean.getValue("bank_no")==null)?"":(bean.getValue("bank_no")).toString(); //--新機構代號
	      	      src_bank_no = (bean.getValue("src_bank_no")==null)?"":(bean.getValue("src_bank_no")).toString(); //--舊機構代號
	      	      bank_name = (bean.getValue("bank_name")==null)?"":(bean.getValue("bank_name")).toString(); //--總機構名稱
	      	      online_date = (bean.getValue("online_date")==null)?"":Utility.getCHTdate(Utility.getTrimString(bean.getValue("online_date")),0);//--機構代碼轉換日期	      	     
	      	      trans_date = (bean.getValue("trans_date")==null)?"":Utility.getCHTdate(Utility.getTrimString(bean.getValue("trans_date")),0);//--MIS系統轉換日期
	              
	      	      //System.out.println("bank_no="+bank_no);
	      	      //System.out.println("src_bank_no="+src_bank_no);
	      	      //System.out.println("bank_name="+bank_name);
	      	      //System.out.println("online_date="+online_date);
	      	      //System.out.println("trans_date="+trans_date);
	      	      rowNum++;
    	      	  row = sheet.createRow(rowNum);
   	    	  	  //列印各機構代號明細資料
	      	      for(int cellcount=0;cellcount<5;cellcount++){			 	      
	      	          cell=row.createCell((short)cellcount);			 		
	      	          cell.setEncoding(HSSFCell.ENCODING_UTF_16);	     
	      	          cell.setCellStyle(cs_center);
	      	          if(cellcount == 0) cell.setCellValue(bank_no);	      	          
	      	          if(cellcount == 1) cell.setCellValue(src_bank_no);	      	          
	      	          if(cellcount == 2) cell.setCellValue(bank_name);
	      	          if(cellcount == 3) cell.setCellValue(online_date);
	      	          if(cellcount == 4) cell.setCellValue(trans_date);
	      	      }//end of cellcount
	      	      
	      	  }
	      	  
	      	 
	      }else{ //end of else dbData.size() != 0	         
             cell.setCellValue("金融機構代號轉換清單   無資料存在");             
	      }
	     
	            	
	      FileOutputStream fout = null;     
	      fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "金融機構代號轉換清單.xls");
	      System.out.println(reportDir + System.getProperty("file.separator") + "金融機構代號轉換清單.xls");
	      HSSFFooter footer = sheet.getFooter();
	      footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
	      footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
	      wb.write(fout);
	      //儲存
	      fout.close();
	      System.out.println("儲存成功!");
	      
	    }catch (Exception e) {
	      System.out.println("RptAS003W.createRpt Error:" + e + e.getMessage());
	    }
	    
	    return errMsg;
	  }	 
}
