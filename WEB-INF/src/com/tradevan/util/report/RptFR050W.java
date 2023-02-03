/*
 * 97.09.08 create 農漁會信用部防治洗錢歷次函報情形明細表 by 2295
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

public class RptFR050W {	 
	 public static String createRpt(String bank_no,String s_year ) {    

	    String errMsg = "";	    
	    StringBuffer sqlCmd = new StringBuffer();
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
	    String bank_name="";//機構名稱
	    String come_date="";//來文日期
	    String come_docno="";//來文文號
	    String sn_date="";//發文日期
	    String sn_docno="";//發文文號
	    String pre_sn_docno="";//前次發文文號 
	    String content="";//處理情形
	    
	    String u_year = "99" ;
		String cd01Table = "cd01_99" ;
		StringBuffer bn01= new StringBuffer() ;
		StringBuffer wlx01 = new StringBuffer() ;
	    try {
	    	if(!"".equals(s_year) && Integer.parseInt(s_year)>99) {
	    		u_year = "100" ;
	    		cd01Table = "cd01" ;
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
	      finput = new FileInputStream(xlsDir +System.getProperty("file.separator") +"農漁會信用部防治洗錢歷次函報情形明細表.xls");

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
	   	 		
	      sqlCmd.append(" select bn01.bank_no,bn01.bank_name,"					  
				 + " ((TO_CHAR(come_date,'yyyy')-1911)||'/'|| TO_CHAR(come_date,'mm/dd')) as come_date,"
				 + " TO_CHAR(come_date,'yyyy/mm/dd') as come_date_1,"
				 + " ((TO_CHAR(sn_date,'yyyy')-1911)||'/'|| TO_CHAR(sn_date,'mm/dd')) as sn_date,"
				 + " TO_CHAR(sn_date,'yyyy/mm/dd') as sn_date_1,"
				 + " come_docno,sn_docno,pre_sn_docno,"
				 + " decode(content,'01','同意備查','02','應修改',content) as content"
				 + " from mis_moneylaunder "
				 + " left join ").append(bn01.toString()).append(" bn01 on mis_moneylaunder.bank_no = bn01.bank_no "
                 + " where mis_moneylaunder.bank_no= ? "
				 + " ORDER BY bank_no,come_date,sn_date" );
			sqlCmdList.add(u_year) ;
			sqlCmdList.add(bank_no) ;
	      dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),sqlCmdList,"come_date,come_date_1,sn_date,sn_date_1");
	      System.out.println("dbData.size=" + dbData.size());
	      
	        
	      if (dbData != null && dbData.size() != 0) {
	          //設定報表表頭資料============================================
	      	  bean = (DataObject)dbData.get(0);
	      	  bank_name=(bean.getValue("bank_name") == null)?"":(String)bean.getValue("bank_name");//機構名稱
		      row=sheet.getRow(1);
		      cell=row.getCell((short)0);	       	
		      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		      cell.setCellValue(bank_name+"防治洗錢歷次函報情形明細表");  	
		      List havePrint = new LinkedList();  
	      	  rowNum = 3;
	      	  boolean checkOK = false;
	      	  for(int i=0;i<dbData.size();i++){
	      	      bank_name="";come_date ="";come_docno="";sn_date="";sn_docno="";pre_sn_docno="";content="";
	      	      bean = (DataObject)dbData.get(i);
	      	      checkOK = false;  
	      	      come_date=(bean.getValue("come_date") == null)?"":(bean.getValue("come_date")).toString();//來文日期
		  	      come_docno=(bean.getValue("come_docno") == null)?"":(String)bean.getValue("come_docno");//來文文號
		  	      sn_date=(bean.getValue("sn_date") == null)?"":(bean.getValue("sn_date")).toString();//發文日期
		  	      sn_docno=(bean.getValue("sn_docno") == null)?"":(String)bean.getValue("sn_docno");//發文文號
		  	      pre_sn_docno=(bean.getValue("pre_sn_docno") == null)?"":(String)bean.getValue("pre_sn_docno");//前次發文文號 
		  	      content=(bean.getValue("content") == null)?"":(String)bean.getValue("content");//處理情形
		  	      for(int k=0;k<havePrint.size();k++){
	                   if((sn_docno).equals((String)havePrint.get(k))){//此筆發文文號.已印過
	                       checkOK = true;
	                   }
	              }     
	              if(!checkOK){
	                  havePrint.add(sn_docno);
	                  rowNum++;
	    	  	      row = sheet.createRow(rowNum);
   	    	  	      //列印各機構明細資料
				      for(int cellcount=0;cellcount<=4;cellcount++){			 	      
			 	          cell=row.createCell((short)cellcount);			 		
			    	      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			    	      cell.setCellStyle(cs_center);		   
			    	      if(cellcount == 0) cell.setCellValue(come_date);//來文日期	
			    	      if(cellcount == 1) cell.setCellValue(come_docno);//來文文號	
			    	      if(cellcount == 2) cell.setCellValue(sn_date);//發文日期	 
			    	      if(cellcount == 3) cell.setCellValue(sn_docno);//發文文號
			    	      if(cellcount == 4) cell.setCellValue(content);//處理情形			    	  
				      }//end of cellcount	
				      
				      for(int j=0;j<dbData.size();j++){
	                      bean_sub =(DataObject)dbData.get(j);  
	                      if(bean_sub.getValue("pre_sn_docno") != null 
	                      && ((String)bean_sub.getValue("pre_sn_docno")).equals(sn_docno)){//有第二次發文的文號
	                          System.out.println("now.sn_docno="+(String)bean.getValue("sn_docno"));  
	                          System.out.println("second.sn_docno="+(String)bean_sub.getValue("sn_docno"));                                                            
	                          System.out.println("pre_sn_docno="+(String)bean_sub.getValue("pre_sn_docno"));
	                          havePrint.add((String)bean_sub.getValue("sn_docno"));
	                          come_date=(bean_sub.getValue("come_date") == null)?"":(bean_sub.getValue("come_date")).toString();//來文日期
	        		  	      come_docno=(bean_sub.getValue("come_docno") == null)?"":(String)bean_sub.getValue("come_docno");//來文文號
	        		  	      sn_date=(bean_sub.getValue("sn_date") == null)?"":(bean_sub.getValue("sn_date")).toString();//發文日期
	        		  	      sn_docno=(bean_sub.getValue("sn_docno") == null)?"":(String)bean_sub.getValue("sn_docno");//發文文號
	        		  	      pre_sn_docno=(bean_sub.getValue("pre_sn_docno") == null)?"":(String)bean_sub.getValue("pre_sn_docno");//前次發文文號 
	        		  	      content=(bean_sub.getValue("content") == null)?"":(String)bean_sub.getValue("content");//處理情形
	        		  	      rowNum++;
	        		  	      row = sheet.createRow(rowNum);
	     	    	  	      //列印各機構明細資料
	  				          for(int cellcount=0;cellcount<=4;cellcount++){			 	      
	  			 	              cell=row.createCell((short)cellcount);			 		
	  			    	          cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	  			    	          cell.setCellStyle(cs_center);		   
	  			    	          if(cellcount == 0) cell.setCellValue(come_date);//來文日期	
	  			    	          if(cellcount == 1) cell.setCellValue(come_docno);//來文文號	
	  			    	          if(cellcount == 2) cell.setCellValue(sn_date);//發文日期	 
	  			    	          if(cellcount == 3) cell.setCellValue(sn_docno);//發文文號
	  			    	          if(cellcount == 4) cell.setCellValue(content);//處理情形			    	  
	  				          }//end of cellcount	
	  				          
	  				          for(int m=0;m<dbData.size();m++){
	                            bean_sub1 =(DataObject)dbData.get(m);  
	                            if(bean_sub1.getValue("pre_sn_docno") != null 
	                            && ((String)bean_sub1.getValue("pre_sn_docno")).equals((String)bean_sub.getValue("sn_docno"))){//有第三次發文的文號
	                               System.out.println("now.sn_docno="+(String)bean_sub.getValue("sn_docno"));  
	                               System.out.println("second.sn_docno="+(String)bean_sub1.getValue("sn_docno"));                                                            
	                               System.out.println("pre_sn_docno="+(String)bean_sub1.getValue("pre_sn_docno"));
	                               havePrint.add((String)bean_sub1.getValue("sn_docno"));
	                               come_date=(bean_sub1.getValue("come_date") == null)?"":(bean_sub1.getValue("come_date")).toString();//來文日期
	 	        		  	       come_docno=(bean_sub1.getValue("come_docno") == null)?"":(String)bean_sub1.getValue("come_docno");//來文文號
	 	        		  	       sn_date=(bean_sub1.getValue("sn_date") == null)?"":(bean_sub1.getValue("sn_date")).toString();//發文日期
	 	        		  	       sn_docno=(bean_sub1.getValue("sn_docno") == null)?"":(String)bean_sub1.getValue("sn_docno");//發文文號
	 	        		  	       pre_sn_docno=(bean_sub1.getValue("pre_sn_docno") == null)?"":(String)bean_sub1.getValue("pre_sn_docno");//前次發文文號 
	 	        		  	       content=(bean_sub1.getValue("content") == null)?"":(String)bean_sub1.getValue("content");//處理情形
	 	        		  	     
	                               rowNum++;
	                               row = sheet.createRow(rowNum);
	 	     	    	  	       //列印各機構明細資料
	 	  				           for(int cellcount=0;cellcount<=4;cellcount++){			 	      
	 	  			 	              cell=row.createCell((short)cellcount);			 		
	 	  			    	          cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	 	  			    	          cell.setCellStyle(cs_center);		   
	 	  			    	          if(cellcount == 0) cell.setCellValue(come_date);//來文日期	
	 	  			    	          if(cellcount == 1) cell.setCellValue(come_docno);//來文文號	
	 	  			    	          if(cellcount == 2) cell.setCellValue(sn_date);//發文日期	 
	 	  			    	          if(cellcount == 3) cell.setCellValue(sn_docno);//發文文號
	 	  			    	          if(cellcount == 4) cell.setCellValue(content);//處理情形			    	  
	 	  				           }//end of cellcount	
	                            }//end of if bean_sub1
	        			      }//end of for bean_sub1     
	                      }//end of if bean_sub
	                  }//end of for bean_sub             
	              }//end of 可以列印    
	      	  }//end of bean
	      } //end of else dbData.size() != 0
	      
	      
	      FileOutputStream fout = null;     
	      fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "農漁會信用部防治洗錢歷次函報情形明細表.xls");
	     
	      HSSFFooter footer = sheet.getFooter();
	      footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
	      footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
	      wb.write(fout);
	      //儲存
	      fout.close();
	    }
	    catch (Exception e) {
	      System.out.println("RptFR050W.createRpt Error:" + e + e.getMessage());
	    }
	    
	    return errMsg;
	  }
}
