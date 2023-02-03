/*
 * 97.09.26 create 農漁會信用部簽証會計師明細表 by 2295
 * 99.05.26 fixed 縣市合併 & sql injection by 2808
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

public class RptFR052W {	 
	 public static String createRpt(String rptStyle,String m_year,String city_type,String bank_no) {    

	    String errMsg = "";	    
	    StringBuffer sqlCmd =  new StringBuffer();
	    List sqlCmdList = new ArrayList() ;
	    StringBuffer condition = new StringBuffer(); 
	    List conditionList = new ArrayList() ;
	    List dbData = null;	    
	    int rowNum=0;
	    DataObject bean = null;
	    DataObject bean_sub = null;
        DataObject bean_sub1 = null;
	    reportUtil reportUtil = new reportUtil();
		HSSFCellStyle cs_right = null; 
		HSSFCellStyle cs_center = null;
	    HSSFCellStyle cs_left = null;
	    String tbank="";
	    String bank_name = "";//機構名稱
	    String name="";//會計師姓名
	    String title="";//事務所名稱
	    String addr="";//地址
	    String telno="";//電話
	   
	    String filename="";
	    String u_year = "99" ;
		StringBuffer bn01= new StringBuffer() ;
		StringBuffer wlx01 = new StringBuffer() ;
		
	    try {
	    	if(!"".equals(m_year) && Integer.parseInt(m_year)>99) {
	    		u_year = "100" ;
	    	}
	    	
	    	bn01.append("(select BANK_NO,BANK_NAME,BN_TYPE,BANK_TYPE,ADD_USER,ADD_NAME,ADD_DATE,BANK_B_NAME,KIND_1,KIND_2,BN_TYPE2,EXCHANGE_NO from bn01 where m_year=? )") ;
			wlx01.append("(select BANK_NO ,ENGLISH,SETUP_APPROVAL_UNT,SETUP_DATE,SETUP_NO,CHG_LICENSE_DATE,");
			wlx01.append(" CHG_LICENSE_NO,CHG_LICENSE_REASON,START_DATE,BUSINESS_ID,HSIEN_ID,AREA_ID,ADDR,");
			wlx01.append(" TELNO,FAX,EMAIL,WEB_SITE,CENTER_FLAG,CENTER_NO,STAFF_NUM,IT_HSIEN_ID,IT_AREA_ID,");
			wlx01.append(" IT_ADDR,IT_NAME,IT_TELNO,AUDIT_HSIEN_ID,AUDIT_AREA_ID,AUDIT_ADDR,AUDIT_NAME,AUDIT_TELNO,FLAG,OPEN_DATE,M2_NAME,");
			wlx01.append(" HSIEN_DIV_1,CANCEL_NO,CANCEL_DATE from wlx01 where m_year =? )");
			
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
	      if(rptStyle.equals("0") || rptStyle.equals("1")){
	         filename = "農漁會信用部簽証會計師明細表.xls";
	      }else{
	      	 filename = "信用部歷年簽証會計師明細表.xls";
	      }
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
	      
	      sqlCmd.append(" select m_year,bn01.bank_no,bn01.bank_name,"
	      		 + " name,title,addr,telno " 
				 + " from boaf_account " 
				 + " left join ").append(bn01.toString()).append(" bn01 on boaf_account.bank_no = bn01.bank_no ");
	      sqlCmdList.add(u_year) ;
	      
	      if(!m_year.equals("")) {
            condition.append( (condition.length() > 0 ? " and":"")+" boaf_account.m_year= ?");
            conditionList.add(m_year) ;
          }
    	  if(!city_type.equals("")) {
    		condition.append( (condition.length() > 0 ? " and":"")+" boaf_account.BANK_NO in (select BANK_NO  from ").append(wlx01.toString()).append(" WLX01  where HSIEN_ID = ? ) " );
    		conditionList.add(u_year) ;
    		conditionList.add(city_type) ;
    	  }
    	  if(!bank_no.equals("")) {
    	    condition.append( (condition.length() > 0 ? " and":"")+" boaf_account.bank_no=? " );
    	    conditionList.add(bank_no) ;
    	  }
    	  if(condition.length() > 0) {
    		  sqlCmd.append( "where ").append(condition.toString() );
    	      for(int i=0 ; i<conditionList.size();i++) {
    	    	  sqlCmdList.add(conditionList.get(i)) ;
    	      }
    	  }
    	  sqlCmd.append( " ORDER BY m_year,bank_no" );
	      dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),sqlCmdList,"m_year");
	      System.out.println("dbData.size=" + dbData.size());
	      
	      //設定報表表頭資料============================================
	      System.out.println("rptStyle="+rptStyle);
      	  if(rptStyle.equals("0") || rptStyle.equals("1")){
	          row=sheet.getRow(2);
	          cell=row.getCell((short)0);	       	
	          cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	          if (dbData != null && dbData.size() != 0) {	 
	          	cell.setCellValue(m_year+"年度");	          
	          }else{
	          	cell.setCellValue(m_year+"年度無資料");	
	          }
      	  }else{
      	      row=sheet.getRow(1);
              cell=row.getCell((short)0);	       	
              cell.setEncoding(HSSFCell.ENCODING_UTF_16);         
      	      if (dbData != null && dbData.size() != 0) {	 
      	         bean = (DataObject)dbData.get(0);
  	             bank_name=(bean.getValue("bank_name") == null)?"":(String)bean.getValue("bank_name");//機構名稱
	             cell.setCellValue(bank_name+"歷年年報明細表");	          
      	      }else{
      	     	cell.setCellValue("歷年年報明細表無資料");
      	      }
      	  }  
	      if (dbData != null && dbData.size() != 0) {	          
	      	  rowNum = 4;
	      	  System.out.println("print detail");
	      	  for(int i=0;i<dbData.size();i++){
	      	      bank_name="";name ="";title="";addr="";telno="";
	      	      bean = (DataObject)dbData.get(i);
	      	      bank_name=(bean.getValue("bank_name") == null)?"":(String)bean.getValue("bank_name");//機構名稱
	      	      m_year=(bean.getValue("m_year") == null)?"":(bean.getValue("m_year")).toString();//年份
		  	      name=(bean.getValue("name") == null)?"":(String)bean.getValue("name");//會計師姓名
		  	      title=(bean.getValue("title") == null)?"":(String)bean.getValue("title");//事務所名稱		  	       
		  	      addr=(bean.getValue("addr") == null)?"":(String)bean.getValue("addr");//地址
		  	      telno=(bean.getValue("telno") == null)?"":(String)bean.getValue("telno");//電話
		  	           
	               row = sheet.createRow(rowNum);
   	    	  	   //列印各機構明細資料
				   for(int cellcount=0;cellcount<=4;cellcount++){			 	      
			 	       cell=row.createCell((short)cellcount);			 		
			    	   cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			    	   cell.setCellStyle(cs_center);	
			    	   if(rptStyle.equals("0") || rptStyle.equals("1")){
			    	      if(cellcount == 0) cell.setCellValue(bank_name);//單位名稱
			    	   }else{
			    	   	  if(cellcount == 0) cell.setCellValue(m_year);//年份
			    	   }
			    	   if(cellcount == 1) cell.setCellValue(name);//會計師姓名	
			    	   if(cellcount == 2) cell.setCellValue(title);//事務所名稱	
			    	   if(cellcount == 3) cell.setCellValue(addr);//地址	 
			    	   if(cellcount == 4) cell.setCellValue(telno);//電話
			    	   
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
	      System.out.println("RptFR052W.createRpt Error:" + e + e.getMessage());
	    }
	    
	    return errMsg;
	  }
}
