/*
 * 106.11.22 create 金融機構代號對照表 by 2295
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

public class RptAS003W { 
     
	 public static String createRpt(String bank_no) {    

	    String errMsg = "";	    
	    List dbData = null;	    
	    int rowNum=0;
	    DataObject bean = null;
	    reportUtil reportUtil = new reportUtil();
		HSSFCellStyle cs_right = null; 
		HSSFCellStyle cs_center = null;
	    HSSFCellStyle cs_left = null;
	  
	    String bank_name=""; //--機構名稱
	    String new_exchange_no = ""; //--新代碼-機構代碼
	    String src_bank_no = ""; //--舊代號-機構代號
	    String exchange_no = ""; //--舊代號-通匯代號
   	   
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
	      finput = new FileInputStream(xlsDir +System.getProperty("file.separator") +"AS003W金融機構代號對照表.xls");      
	      
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
	     
	      cs_center = getCenterStyle(wb);
	      cs_left = getLeftStyle(wb);
	     
          String wlx01_m_year = "";
	      StringBuffer sql = new StringBuffer();
	      List paramList = new ArrayList();//傳入參數
	     
	      sql.append(" select bank_name,");// --單位名稱
	      sql.append(" bank_kind,");
	      sql.append(" ba01_trans.bank_no,");//--新代碼-機構代碼
	      sql.append(" ba01_trans.bank_no as new_exchange_no,");//--新代碼-機構代碼
	      sql.append(" src_bank_no,");//--舊代號-機構代號
	      sql.append(" exchange_no");// --舊代號-通匯代號
	      sql.append(" from ba01_trans left join (select * from bn01 where m_year=100)bn01 on ba01_trans.src_bank_no=bn01.ori_bank_no");
	      sql.append(" where ba01_trans.pbank_no=?");
	      sql.append(" and bank_kind='0'");
	      sql.append(" union");
	      sql.append(" select bank_name,"); 
	      sql.append(" bank_kind,");
	      sql.append(" ba01_trans.bank_no,");
	      sql.append(" ba01_trans.bank_no as new_exchange_no,");
	      sql.append(" src_bank_no,");
	      sql.append(" exchange_no");
	      sql.append(" from ba01_trans left join (select * from bn02 where m_year=100)bn02 on ba01_trans.src_bank_no=bn02.ori_bank_no");
	      sql.append(" where ba01_trans.pbank_no=?");
	      sql.append(" and bank_kind='1'");
	      sql.append(" order by bank_kind,src_bank_no");
	      
	      paramList.add(bank_no);
	      paramList.add(bank_no);           
	      
	      dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"");
	      System.out.println("RptAS003W.dbData.size=" + dbData.size());	      
	      //設定報表表頭資料============================================
	     
          
	      row = sheet.getRow(1);
          cell = row.getCell( (short) 0);
          cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	      if (dbData != null && dbData.size() != 0) {	  
	          bean = (DataObject)dbData.get(0);
	          cell.setCellValue(String.valueOf(bean.getValue("bank_name"))+"金融機構代號對照表");
	          
		   	  rowNum = 3;	      	  
	      	  for(int i=0;i<dbData.size();i++){	
	      	      bean = (DataObject)dbData.get(i);
	      	      bank_name = (bean.getValue("bank_name")==null)?"":(bean.getValue("bank_name")).toString(); //--單位名稱
	      	      bank_no = (bean.getValue("bank_no")==null)?"":(bean.getValue("bank_no")).toString(); //--新代碼-機構代碼
	      	      new_exchange_no = (bean.getValue("new_exchange_no")==null)?"":(bean.getValue("new_exchange_no")).toString(); //--新代碼-機構代碼
	      	      src_bank_no = (bean.getValue("src_bank_no")==null)?"":(bean.getValue("src_bank_no")).toString(); //--舊代號-機構代號
	      	      exchange_no = (bean.getValue("exchange_no")==null)?"":(bean.getValue("exchange_no")).toString(); //--舊代號-通匯代號
	      	      System.out.println("bank_name="+bank_name);
	      	      System.out.println("bank_no="+bank_no);
	      	      System.out.println("new_exchange_no="+new_exchange_no);
	      	      System.out.println("src_bank_no="+src_bank_no);
	      	      System.out.println("exchange_no="+exchange_no);
	      	      rowNum++;
    	      	  row = sheet.getRow(rowNum);
   	    	  	  //列印各機構代號明細資料
	      	      for(int cellcount=0;cellcount<5;cellcount++){			 	      
	      	          cell=row.getCell((short)cellcount);			 		
	      	          cell.setEncoding(HSSFCell.ENCODING_UTF_16);	      	          
	      	          if(cellcount == 0){
	      	            //cell.setCellStyle(cs_left);
	      	            cell.setCellValue(bank_name);
	      	          }//else{
	      	           // cell.setCellStyle(cs_center);
	      	          //}	      	              
	      	          
	      	          if(cellcount == 1) cell.setCellValue(bank_no);	      	          
	      	          if(cellcount == 2) cell.setCellValue(new_exchange_no);
	      	          if(cellcount == 3) cell.setCellValue(src_bank_no);
	      	          if(cellcount == 4) cell.setCellValue(exchange_no);
	      	      }//end of cellcount
	      	      
	      	  }
	      	  
	      	  for(int i = rowNum;i<22;i++){
	      	      ++rowNum;
	      	      row = sheet.getRow(rowNum);
	      	      sheet.removeRow(row);
              }
	      	  row = sheet.getRow((short)0);
	      	  cell=row.getCell((short)0);
	      	  cell.setCellValue("");
	      }else{ //end of else dbData.size() != 0	         
             cell.setCellValue("金融機構代號對照表   無資料存在");             
	      }
	     
	            	
	      FileOutputStream fout = null;     
	      fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "金融機構代號對照表.xls");
	      System.out.println(reportDir + System.getProperty("file.separator") + "金融機構代號對照表.xls");
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
	 private static String getYYYYMMDDHHMMSS(){
	        Calendar rightNow = Calendar.getInstance();
	        String year = (new Integer(rightNow.get(Calendar.YEAR))).toString();
	        String month = (new Integer(rightNow.get(Calendar.MONTH) + 1)).toString();
	        String day = (new Integer(rightNow.get(Calendar.DAY_OF_MONTH))).toString();
	        String hour = (new Integer(rightNow.get(Calendar.HOUR_OF_DAY))).toString();
	        String minute = (new Integer(rightNow.get(Calendar.MINUTE))).toString();
	        String second = (new Integer(rightNow.get(Calendar.SECOND))).toString();

	        if (month.length() == 1) {
	            month = "0" + month;
	        }
	        if (day.length() == 1) {
	            day = "0" + day;
	        }
	        if (hour.length() == 1) {
	            hour = "0" + hour;
	        }
	        if (minute.length() == 1) {
	            minute = "0" + minute;
	        }
	        if (second.length() == 1) {
	            second = "0" + second;

	        }
	        return (year + month + day + hour + minute + second);
	    }
	    //有框內文置左
	    private static HSSFCellStyle getLeftStyle(HSSFWorkbook wb){
	            HSSFCellStyle leftStyle = wb.createCellStyle(); 
	            HSSFFont f = wb.createFont();
	            //set font 1 to 12 point type
	            f.setFontHeightInPoints((short) 13);
	            f.setFontName("標楷體");
	            
	            leftStyle = HssfStyle.setStyle( leftStyle, f,
	                                    new String[] {
	                                    "BORDER", "PHL", "PVC", "F13",
	                                    "WRAP"} );
	            return leftStyle;   
	    }    
	    //有框內文置中
	    private static HSSFCellStyle getCenterStyle(HSSFWorkbook wb){ 
	           HSSFCellStyle defaultStyle1 = wb.createCellStyle();     
	           HSSFFont f = wb.createFont();
               //set font 1 to 12 point type
               f.setFontHeightInPoints((short) 13);
               f.setFontName("標楷體");
	           defaultStyle1 = HssfStyle.setStyle( defaultStyle1, f,
	                                       new String[] {
	                                       "BORDER", "PHC", "PVC", "F13",
	                                       "WRAP"} );
	           return defaultStyle1;
	    } 
}
