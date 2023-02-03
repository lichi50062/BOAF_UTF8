/*
 * 98.06.23 個別農漁會解釋函令   by 2756
 * 98.06.26 add 合併檔案/機構名稱獨立Query/無資料顯示 by 2295
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

public class RptMC011W {	 
	 public static String createRpt(String bank_no,String febxlsFlag,HSSFWorkbook wb) {    

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
	    String tbank_no_name="";//總機構名稱
	    String bank_name="";//發文對象
	    String dffe_date="";//發文日期
	    String dffe_number="";//發文文號
	    String dffe_title="";//函令主旨
	    String dffe_content="";//內容

	    
	    
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

	      filename="個別農漁會解釋函令_檢查局.xls";
	      
          if(febxlsFlag.equals("")){
          	finput = new FileInputStream(xlsDir +System.getProperty("file.separator") +filename);
          	//設定FileINputStream讀取Excel檔
          	POIFSFileSystem fs = new POIFSFileSystem(finput);
	        wb = new HSSFWorkbook(fs);
          }
          HSSFSheet sheet =null;
	  	  sheet = wb.getSheetAt(0);//讀取第一個工作表，宣告其為sheet
	  	  
	     
	      HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
	      //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	      //sheet.setAutobreaks(true); //自動分頁
	      
	      //設定頁面符合列印大小
	      sheet.setAutobreaks(false);
	      ps.setScale( (short) 100); //列印縮放百分比
 
	      ps.setPaperSize( (short) 9); //設定紙張大小 A4
	      //wb.setSheetName(0,"test");
	      if(febxlsFlag.equals("")) finput.close();

	      HSSFRow row = null; //宣告一列
	      HSSFCell cell = null; //宣告一個儲存格
	      
	      cs_right = reportUtil.getRightStyle(wb);
	      cs_center = reportUtil.getDefaultStyle(wb);
	      cs_left = reportUtil.getLeftStyle(wb);
	      List paramList = new ArrayList();
	      paramList.add(bank_no);
	      dbData = DBManager.QueryDB_SQLParam("select bank_name from ba01 where bank_no=?",paramList,"");
          if(dbData.size() != 0){
          	tbank_no_name  = (String)((DataObject)dbData.get(0)).getValue("bank_name");
          }
	      
	      StringBuffer sb=new StringBuffer();
	      sb.append(" select mis_dffe.bank_no,ba01.bank_name,((TO_CHAR(dffe_date,'yyyy')-1911)||'/'|| TO_CHAR(dffe_date,'mm/dd')) as dffe_date,");
	      sb.append(" dffe_number,dffe_title,dffe_content");
	      sb.append(" from mis_dffe left join ba01 on mis_dffe.bank_no=ba01.bank_no");
	      
    	  if(!bank_no.equals("")) 
    	  {
    	    paramList = new ArrayList();
            paramList.add(bank_no);
      	    condition += (condition.length() > 0 ? " and":"")+" ba01.pbank_no=?";		
      	  }
    	  sqlCmd=sb.toString();
       	  if(condition.length() > 0) sqlCmd += " where "+condition;
       	  sqlCmd += " order by dffe_date desc";          
           
	      dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
	      System.out.println("dbData.size=" + dbData.size());

	      //設定表頭資料	           
	      row=sheet.getRow(0);
	      cell=row.getCell((short)0);
	      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	      cell.setCellValue(tbank_no_name+"解釋函令");

          //置放表格內容資料          
	      if (dbData != null && dbData.size() != 0) 
	      {	          
	    	  rowNum = 2;//從表格的第二列開始放置第一筆資料(因編號0~1皆為表頭資料)
	      	  System.out.println("print detail");
	      	  for(int i=0;i<dbData.size();i++)
	      	  {
	      		bean = (DataObject)dbData.get(i);
	      		
	    	    bank_name=(bean.getValue("bank_name") == null)?"":(String)bean.getValue("bank_name");
	    	    dffe_date=(bean.getValue("dffe_date") == null)?"":(String)bean.getValue("dffe_date");
	    	    dffe_number=(bean.getValue("dffe_number") == null)?"":(String)bean.getValue("dffe_number");
	    	    dffe_title=(bean.getValue("dffe_title") == null)?"":(String)bean.getValue("dffe_title");
	    	    dffe_content=(bean.getValue("dffe_content") == null)?"":(String)bean.getValue("dffe_content");      			      		   
	               row = sheet.createRow(rowNum);
   	    	  	   //列印各機構明細資料
				   for(int cellcount=0;cellcount<=4;cellcount++){			 	      
			 	       cell=row.createCell((short)cellcount);			 		
			    	   cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			    	   cell.setCellStyle(cs_center);	
			    	   
			    	   if(cellcount == 0) 
			    	   {
			    		   cell.setCellStyle(cs_center);
			    		   cell.setCellValue(bank_name);//發文對象
			    	   }
			    	   else if(cellcount == 1)
			    	   { 
			    		   cell.setCellStyle(cs_center);
			    		   cell.setCellValue(dffe_date);//發文日期	
			    	   }
			    	   else if(cellcount == 2)
			    	   {
			    		   cell.setCellStyle(cs_center);
			    		   cell.setCellValue(dffe_number);//發文文號
			    	   }
			    	   else if(cellcount == 3)
			    	   { 
			    		   cell.setCellStyle(cs_left);
			    		   cell.setCellValue(dffe_title);//函令主旨
			    	   }
			    	   else if(cellcount == 4)
			    	   { 
			    		   cell.setCellStyle(cs_left);
			    		   cell.setCellValue(dffe_content);//內容
			    	   }	    	   
				   }//end of cellcount	
				   rowNum++;   
	      	  }//end of bean
	      }else{ //end of else dbData.size() != 0	      	    
	      	  cell.setCellValue(tbank_no_name+"解釋函令無資料");
	      }
	      if(febxlsFlag.equals("")){//原個別農漁會解釋函令_檢查局
	      	FileOutputStream fout = null;     
	      	fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + filename);
	     
	      	HSSFFooter footer = sheet.getFooter();
	      	footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
	      	footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
	      	wb.write(fout);
	      	//儲存
	      	fout.close();
	      }
	    }
	    catch (Exception e) {
	      System.out.println("RptMC011W.createRpt Error:" + e + e.getMessage());
	    }
	    
	    return errMsg;
	  }
}
