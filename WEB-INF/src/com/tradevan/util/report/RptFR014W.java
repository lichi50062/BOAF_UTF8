/*
 * 98.06.23 檢舉書 by 2756
 * 102.11.19 add 原QueryDB改套用QueryDB_SQLParam by 2295    
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

public class RptFR014W {	 
	 public static String createRpt(String bank_no,String year,String mth,String day ) {    

	    String errMsg = "";	    
	    String sqlCmd = ""; 
	    String condition = "";
	    List dbData = null;	    
	    int rowNum=0;
	    DataObject bean = null;
	    reportUtil reportUtil = new reportUtil();
		HSSFCellStyle cs_right = null; 
		HSSFCellStyle cs_center = null;
	    HSSFCellStyle cs_left = null;

	    String bank_name="";//被檢舉機構
	    String ta_number="";//收文文號
	    String ta_date="";//收文日期
	    String ta_reporter="";//檢舉人
	    String ta_content="";//檢舉內容
	    String ta_ps="";//備註

	    
	    
	    String filename="";
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

	      filename="農會信用部檢舉書.xls";
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
	      
	      cs_right = reportUtil.getRightStyle(wb);
	      cs_center = reportUtil.getDefaultStyle(wb);
	      cs_left = reportUtil.getNoBorderLeftStyle(wb);
	      
	      StringBuffer sb=new StringBuffer();
	      sb.append(" select ta_number,to_char(mis_ta.ta_date,'YYYY-MM-DD') as ta_date,ta_reporter,mis_ta.bank_no,ba01.bank_name,ta_content,ta_ps");
	      sb.append(" from mis_ta left join ba01 on mis_ta.bank_no=ba01.bank_no");
	      List paramList = new ArrayList();
    	  if(!bank_no.equals("")) 
    	  {
      	    condition += (condition.length() > 0 ? " and":"")+" ba01.pbank_no=?";	
      	    paramList.add(bank_no);  
      	  }
    	  sqlCmd=sb.toString();
       	  if(condition.length() > 0) sqlCmd += " where "+condition;
       	  sqlCmd += "order by ta_date desc";
	      dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
	      System.out.println("dbData.size=" + dbData.size());

	      //設定表頭資料
	      bean = (DataObject)dbData.get(0);	      
	      row=sheet.getRow(0);
	      cell=row.getCell((short)0);
	      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	      cell.setCellValue(bean.getValue("bank_name")+"檢舉書");
	      row=sheet.getRow(1);
	      cell=row.getCell((short)0);
	      cell.setEncoding(HSSFCell.ENCODING_UTF_16);

          //置放表格內容資料          
	      if (dbData != null && dbData.size() != 0) 
	      {	          
	    	  rowNum = 2;//從表格的第二列開始放置第一筆資料(因編號0~1皆為表頭資料)
	      	  System.out.println("print detail");
	      	  for(int i=0;i<dbData.size();i++)
	      	  {
	      		bean = (DataObject)dbData.get(i);
	      		bank_name=(bean.getValue("bank_name") == null)?"":(String)bean.getValue("bank_name");
	    	    ta_number=(bean.getValue("ta_number") == null)?"":(String)bean.getValue("ta_number");
	    	    ta_date=(bean.getValue("ta_date") == null)?"":(String)bean.getValue("ta_date");
	    	    ta_reporter=(bean.getValue("ta_reporter") == null)?"":(String)bean.getValue("ta_reporter");
	    	    ta_content=(bean.getValue("ta_content") == null)?"":(String)bean.getValue("ta_content");
	    	    ta_ps=(bean.getValue("ta_ps") == null)?"":(String)bean.getValue("ta_ps");

	      		   
	               row = sheet.createRow(rowNum);
   	    	  	   //列印各機構明細資料
				   for(int cellcount=0;cellcount<=12;cellcount++){			 	      
			 	       cell=row.createCell((short)cellcount);			 		
			    	   cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			    	   cell.setCellStyle(cs_center);	

			    	   if(cellcount == 0) cell.setCellValue(ta_number);//收文文號
			    	   if(cellcount == 1) cell.setCellValue(ta_date);//收文日期	
			    	   if(cellcount == 2) cell.setCellValue(ta_reporter);//檢舉人	
			    	   if(cellcount == 3) cell.setCellValue(bank_name);//被檢舉機構
			    	   if(cellcount == 4) cell.setCellValue(ta_content);//檢舉內容
			    	   if(cellcount == 5) cell.setCellValue(ta_ps);//備註
			    	   		    	   
				   }//end of cellcount	
				   rowNum++;   
	      	  }//end of bean
	      } //end of else dbData.size() != 0
	      
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
	      System.out.println("RptFR014W.createRpt Error:" + e + e.getMessage());
	    }
	    
	    return errMsg;
	  }
}
