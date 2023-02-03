/*
 * 98.06.23 限制或核准業務函令_檢查局   by 2756
 * 98.06.26 add 合併檔案/機構名稱獨立Query/無資料顯示 by 2295
 * 98.10.22 拿掉業務開辦日期 by 2295
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

public class RptMC010W {	 
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
	    String bank_name="";//受限制或核准機構
	    String loal_manager="";//主管機關
	    String loal_add_date="";//函文日期
	    String loal_number="";//限制或核准函號
	    String loal_open_date="";//業務開辦日期
	    String loal_states_name="";//狀態
	    String loal_content="";//限制或核准內容
	    String loal_ps="";//備註

	    
	    
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
	      if(febxlsFlag.equals("")){//98.06.25 原限制或核准業務函令
	      	filename="限制或核准業務函令_檢查局.xls";
	      	finput = new FileInputStream(xlsDir +System.getProperty("file.separator") +filename);
	      	//設定FileINputStream讀取Excel檔
	      	POIFSFileSystem fs = new POIFSFileSystem(finput);
	      	wb = new HSSFWorkbook(fs);
	      }
	      HSSFSheet sheet =null;
	      if(febxlsFlag.equals("")){//98.06.25 原限制或核准業務函令
	      	 sheet = wb.getSheetAt(0); //讀取第一個工作表，宣告其為sheet
	  	  }else{//98.06.25合併檔案  			
	  		 sheet = wb.getSheetAt(1);
	  	  }
	      
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
	      sb.append(" select mis_loal.bank_no,ba01.bank_name,loal_manager,decode(loal_add_date,null,'',((TO_CHAR(loal_add_date,'yyyy')-1911)||'/'|| TO_CHAR(loal_add_date,'mm/dd'))) as loal_add_date,");
	      sb.append(" loal_number,mis_select.select_name as loal_states_name,loal_states,decode(loal_open_flag,'Y','尚未開辦', decode(loal_open_date,null,'',((TO_CHAR(loal_open_date,'yyyy')-1911)||'/'|| TO_CHAR(loal_open_date,'mm/dd')))) as loal_open_date,loal_content,loal_ps");
	      sb.append(" from mis_loal left join ba01 on mis_loal.bank_no=ba01.bank_no left join (select * from mis_select where select_id='LOAL_STATES')mis_select");
	      sb.append(" on mis_loal.loal_states = mis_select.select_num");
	      
    	  if(!bank_no.equals("")) 
    	  {
    	    paramList = new ArrayList();
            paramList.add(bank_no);
      	    condition += (condition.length() > 0 ? " and":"")+" ba01.pbank_no=?";		
      	  }
    	  sqlCmd=sb.toString();
       	  if(condition.length() > 0) sqlCmd += " where "+condition;

	      dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
	      System.out.println("dbData.size=" + dbData.size());

	      //設定表頭資料	            
	      row=sheet.getRow(0);
	      cell=row.getCell((short)1);
	      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	      cell.setCellValue(tbank_no_name+"限制或核准業務函令");

          //置放表格內容資料          
	      if (dbData != null && dbData.size() != 0) 
	      {	          
	    	  rowNum = 2;//從表格的第二列開始放置第一筆資料(因編號0~1皆為表頭資料)
	      	  System.out.println("print detail");
	      	  for(int i=0;i<dbData.size();i++)
	      	  {
	      		bean = (DataObject)dbData.get(i);
	      		
	    	    bank_name=(bean.getValue("bank_name") == null)?"":(String)bean.getValue("bank_name");
	    	    loal_manager=(bean.getValue("loal_manager") == null)?"":(String)bean.getValue("loal_manager");
	    	    loal_add_date=(bean.getValue("loal_add_date") == null)?"":(String)bean.getValue("loal_add_date");
	    	    loal_number=(bean.getValue("loal_number") == null)?"":(String)bean.getValue("loal_number");
	    	    loal_open_date=(bean.getValue("loal_open_date") == null)?"":(String)bean.getValue("loal_open_date");
	    	    loal_states_name=(bean.getValue("loal_states_name") == null)?"":(String)bean.getValue("loal_states_name");
	    	    loal_content=(bean.getValue("loal_content") == null)?"":(String)bean.getValue("loal_content");
	    	    loal_ps=(bean.getValue("loal_ps") == null)?"":(String)bean.getValue("loal_ps");
	      			      		   
	               row = sheet.createRow(rowNum);
   	    	  	   //列印各機構明細資料
				   for(int cellcount=1;cellcount<=7;cellcount++){			 	      
			 	       cell=row.createCell((short)cellcount);			 		
			    	   cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			    	   cell.setCellStyle(cs_center);	
			    	   
			    	   if(cellcount == 1)
			    	   { 
			    		   cell.setCellStyle(cs_center);
			    		   cell.setCellValue(bank_name);//受限制或核准機構	
			    	   }
			    	   else if(cellcount == 2)
			    	   { 
			    		   cell.setCellStyle(cs_center);
			    		   cell.setCellValue(loal_manager);//主管機關	
			    	   }
			    	   else if(cellcount == 3)
			    	   { 
			    		   cell.setCellStyle(cs_center);
			    		   cell.setCellValue(loal_add_date);//函文日期
			    	   }
			    	   else if(cellcount == 4)
			    	   { 
			    		   cell.setCellStyle(cs_center);
			    		   cell.setCellValue(loal_number);//限制或核准函號
			    	   }
			    	   /*98.10.22 拿掉業務開辦日期
			    	   else if(cellcount == 5)
			    	   { 
			    		   cell.setCellStyle(cs_center);
			    		   cell.setCellValue(loal_open_date);//業務開辦日期
			    	   }
			    	   */
			    	   else if(cellcount == 5)
			    	   { 
			    		   cell.setCellStyle(cs_center);
			    		   cell.setCellValue(loal_states_name);//狀態
			    	   }
			    	   else if(cellcount == 6)
			    	   { 
			    		   cell.setCellStyle(cs_left);
			    		   cell.setCellValue(loal_content);//限制或核准內容
			    	   }
			    	   else if(cellcount == 7)
			    	   { 
			    		   cell.setCellStyle(cs_left);
			    		   cell.setCellValue(loal_ps);//備註			    	   
			    	   }
			    	   		    	   
				   }//end of cellcount	
				   rowNum++;   
	      	  }//end of bean
	      }else{ //end of else dbData.size() != 0
	      	  cell.setCellValue(tbank_no_name+"限制或核准業務函令無資料");
	      }
	      if(febxlsFlag.equals("")){	
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
	      System.out.println("RptMC010W.createRpt Error:" + e + e.getMessage());
	    }
	    
	    return errMsg;
	  }
}
