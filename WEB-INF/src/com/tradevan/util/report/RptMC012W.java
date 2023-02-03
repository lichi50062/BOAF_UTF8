/*
 * 98.06.23 舞弊案件   by 2756
 * 98.06.26 add 機構名稱獨立Query/無資料顯示 by 2295
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

public class RptMC012W {	 
	 public static String createRpt(String bank_no) {    

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
	    String em_number="";//來文文號
	    String em_date="";//發文日期
	    String bank_name="";//發生事件機構
	    String em_content="";//事件內容
	    
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

	      filename="舞幣案件_檢查局.xls";
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
	      cs_left = reportUtil.getLeftStyle(wb);
	      List paramList = new ArrayList();
	      paramList.add(bank_no);
	      dbData = DBManager.QueryDB_SQLParam("select bank_name from ba01 where bank_no=?",paramList,"");
          if(dbData.size() != 0){
          	tbank_no_name  = (String)((DataObject)dbData.get(0)).getValue("bank_name");
          }	      
	      
	      StringBuffer sb=new StringBuffer();
	      sb.append(" select em_number,((TO_CHAR(em_date,'yyyy')-1911)||'/'|| TO_CHAR(em_date,'mm/dd')) as em_date,mis_em.bank_no,ba01.bank_name,em_content");
	      sb.append(" from mis_em left join ba01 on mis_em.bank_no=ba01.bank_no");
	      
    	  if(!bank_no.equals("")) 
    	  {
    	    paramList = new ArrayList();
            paramList.add(bank_no);
      	    condition += (condition.length() > 0 ? " and":"")+" ba01.pbank_no=?";		
      	  }
    	  sqlCmd=sb.toString();
       	  if(condition.length() > 0) sqlCmd += " where "+condition;
       	  sqlCmd += "order by em_date desc";

	      dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
	      System.out.println("dbData.size=" + dbData.size());

	      //設定表頭資料	         
	      row=sheet.getRow(0);
	      cell=row.getCell((short)0);
	      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	      cell.setCellValue(tbank_no_name+"重大偶發事件（含舞弊案）");

          //置放表格內容資料          
	      if (dbData != null && dbData.size() != 0) 
	      {	          
	    	  rowNum = 2;//從表格的第二列開始放置第一筆資料(因編號0~1皆為表頭資料)
	      	  System.out.println("print detail");
	      	  for(int i=0;i<dbData.size();i++)
	      	  {
	      		bean = (DataObject)dbData.get(i);	      			      		
	      		bank_name=(bean.getValue("bank_name") == null)?"":(String)bean.getValue("bank_name");
	      		em_number=(bean.getValue("em_number") == null)?"":(String)bean.getValue("em_number");
	      		em_date=(bean.getValue("em_date") == null)?"":(String)bean.getValue("em_date");
	      		em_content=(bean.getValue("em_content") == null)?"":(String)bean.getValue("em_content");
     			      		   
	               row = sheet.createRow(rowNum);
   	    	  	   //列印各機構明細資料
				   for(int cellcount=0;cellcount<=3;cellcount++){			 	      
			 	       cell=row.createCell((short)cellcount);			 		
			    	   cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			    	   
			    	   if(cellcount == 0)
			    	   {
				    	   cell.setCellStyle(cs_center);
			    		   cell.setCellValue(em_number);//來文文號	
			    	   }
			    	   else if(cellcount == 1)
			    	   {			    	  
			    		   cell.setCellStyle(cs_center);
			    		   cell.setCellValue(em_date);//發文日期	
			    	   }
			    	   else if(cellcount == 2)
			    	   {
				    	   cell.setCellStyle(cs_center);
				    	   cell.setCellValue(bank_name);//發生事件機構
			    	   }
			    	   else if(cellcount == 3)
			    	   {
			    		   cell.setCellStyle(cs_left);
			    		   cell.setCellValue(em_content);//事件內容
			    	   }
	    	   
			    	   		    	   
				   }//end of cellcount	
				   rowNum++;   
	      	  }//end of bean
	      }else{//end of else dbData.size() != 0
	      	  cell.setCellValue(tbank_no_name+"重大偶發事件（含舞弊案）無資料");
	      }
	      
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
	      System.out.println("RptMC012W.createRpt Error:" + e + e.getMessage());
	    }
	    
	    return errMsg;
	  }
}
