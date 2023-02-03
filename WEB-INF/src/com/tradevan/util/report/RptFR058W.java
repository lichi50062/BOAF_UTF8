/*
 * 99.09.29 create 全體農漁會信用部主任及分部主任參加金融相關業務進修情形統計表 by 2295
 * 99.10.11 fix 取得未裁撤的分支機構.同一個總機構代號/總機構名稱合併儲存格 by 2295
 */
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.Region;

import java.io.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR058W {	 
	 public static String createRpt(String m_year,String m_month) {    

	    String errMsg = "";	
	    String condition = "";
	    List dbData_over16 = null;	    
	    List dbData_less16 = null;
	    int rowNum=0;
	    int rowNum_start=0;
	   	int rowNum_end=0;
	    DataObject bean = null;	    
	    reportUtil reportUtil = new reportUtil();
		HSSFCellStyle cs_right = null; 
		HSSFCellStyle cs_center = null;
	    HSSFCellStyle cs_left = null;
	    String pbank_no = "";
	    String cellname[] = {"pbank_no","pbank_name","bank_name","name","total_hour"};    
	    
	   
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
	        
	      FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+ "全體農漁會信用部主任及分部主任參訓情形統計表.xls" );

          //設定FileINputStream讀取Excel檔
          POIFSFileSystem fs = new POIFSFileSystem( finput );
          if(fs==null){System.out.println("open 範本檔失敗");} else System.out.println("open 範本檔成功");
          HSSFWorkbook wb = new HSSFWorkbook(fs);
          if(wb==null){System.out.println("open工作表失敗");}else System.out.println("open 工作表 成功");
          HSSFSheet sheet = wb.getSheetAt(0);//讀取第一個工作表，宣告其為sheet
          if(sheet==null){System.out.println("open sheet 失敗");}else System.out.println("open sheet 成功");
          HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
	     
	      //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	      //sheet.setAutobreaks(true); //自動分頁

	      //設定頁面符合列印大小
	      sheet.setAutobreaks(false);
	      ps.setScale( (short) 95); //列印縮放百分比	      
	      ps.setPaperSize( (short) 9); //設定紙張大小 A4
	      //wb.setSheetName(0,"test");	      

	      HSSFRow row = null; //宣告一列
	      HSSFCell cell = null; //宣告一個儲存格	      
	      cs_right = reportUtil.getRightStyle(wb);
	      cs_center = reportUtil.getDefaultStyle(wb);
	      cs_left = reportUtil.getNoBorderLeftStyle(wb);
	      m_year = String.valueOf(Integer.parseInt(m_year));
	      m_month = String.valueOf(Integer.parseInt(m_month));
	      	      
	      String cd01_table = "";
          String wlx01_m_year = "";
	      StringBuffer sql = new StringBuffer();
	      List paramList = new ArrayList();//傳入參數
	      //99.09.29 add 查詢年度100年以前.縣市別不同===============================
	      cd01_table = (Integer.parseInt(m_year) < 100)?"cd01_99":""; 
	      wlx01_m_year = (Integer.parseInt(m_year) < 100)?"99":"100"; 
	      //===================================================================== 
	      //總機構.分支機構滿16小時
	      sql.append(" select * ");
		  sql.append(" from ( ");
		  sql.append("       select ba01.bank_type,ba01.pbank_no,bn01.bank_name as pbank_name,ba01.bank_no,ba01.bank_name,name,total_hour ");
          sql.append("       from (select * from ba01 where m_year=? and bank_type in ('6','7') and pbank_no != '8888888')ba01 ");//99.09.29區分100年度
		  sql.append("       left join (select * from bn01 where m_year=? and bank_type in ('6','7') and bn_type <> '2' and bank_no != '8888888')bn01 on ba01.pbank_no = bn01.bank_no ");//99.09.29區分100年度
		  sql.append("       left join ");
		  sql.append("       (select bank_no,name,sum(course_hour) as total_hour ");
		  sql.append("       from wlx_trainning ");
		  sql.append("       where m_year = ? and m_month <= ?");//99.09.29		  		  
		  sql.append("       and position_code in ('1','2') ");//--信用部主任/分部主任
		  sql.append("       group by tbank_no,bank_no,name)wlx_trainning on ba01.bank_no = wlx_trainning.bank_no ");
		  sql.append("       where bn01.bank_name is not null ");
		  sql.append("       order by pbank_no,bank_no,name ");
		  sql.append("       )tmp where bank_no in (select bank_no ");
		  sql.append("                              from ");
		  sql.append("                              ( ");
		  sql.append("                              select bank_no,sum(course_hour) as total_hour ");
		  sql.append("                              from wlx_trainning ");
		  sql.append("                              where m_year = ? and m_month <= ?");//99.09.29
		  sql.append("                              and position_code in ('1','2') ");//--信用部主任/分部主任
		  sql.append("                              group by bank_no ");
		  sql.append("                              )where total_hour >= 16) ");
		  sql.append(" order by bank_type,pbank_no,bank_no,name ");  
		  paramList.add(wlx01_m_year);//99.09.29區分100年度
		  paramList.add(wlx01_m_year);//99.09.29區分100年度
		  paramList.add(m_year);
		  paramList.add(m_month);
		  paramList.add(m_year);
		  paramList.add(m_month);
		  
		  
	      dbData_over16 = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"total_hour");//99.09.29	      
	      System.out.println("dbData_over16.size=" + dbData_over16.size());
	      
	      //總機構.分支機構不滿16小時
	      sql.delete(0,sql.length());	      
	      sql.append(" select * ");
          sql.append(" from ( ");
          sql.append("       select ba01.bank_type,ba01.pbank_no,bn01.bank_name as pbank_name,ba01.bank_no,ba01.bank_name,name,total_hour ");
          
          sql.append(" from ( ");
          sql.append(" 	     select ba01.bank_no,ba01.bank_name,ba01.bank_type,ba01.pbank_no from ba01 ");
          sql.append(" 	     where bank_no not in (select bank_no from bn02 where bank_type in ('6','7') and bn_type = '2') ");//--已裁撤的分支機構	  
          sql.append(" 	     and m_year=? and ba01.bank_type in ('6','7')  and pbank_no != '8888888' ");//99.09.29區分100年度
          sql.append("       )ba01 ");          
          sql.append("       left join (select * from bn01 where m_year=? and bank_type in ('6','7') and bn_type <> '2' and bank_no != '8888888')bn01 on ba01.pbank_no = bn01.bank_no ");//99.09.29區分100年度 
          sql.append("       left join ");
          sql.append("       (select bank_no,name,sum(course_hour) as total_hour ");
          sql.append("       from wlx_trainning ");          
          sql.append("       where m_year = ? and m_month <= ? ");//99.09.29
          sql.append("       and position_code in ('1','2') ");//--信用部主任/分部主任
          sql.append("       group by tbank_no,bank_no,name)wlx_trainning on ba01.bank_no = wlx_trainning.bank_no ");
          sql.append("       where bn01.bank_name is not null ");
          sql.append("       order by pbank_no,bank_no,name ");
          sql.append("       )tmp where bank_no not in (select bank_no ");
          sql.append("                              from ");
          sql.append("                              ( ");
          sql.append("                              select bank_no,sum(course_hour) as total_hour ");
          sql.append("                              from wlx_trainning ");          
          sql.append("                              where m_year = ? and m_month <= ? "); //99.09.29
          sql.append("                              and position_code in ('1','2') ");//--信用部主任/分部主任
          sql.append("                              group by bank_no ");
          sql.append("                              )where total_hour >= 16) ");
          sql.append(" order by bank_type,pbank_no,bank_no,name ");
          	      
          dbData_less16 = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"total_hour");//99.09.29
	      System.out.println("dbData_less16.size=" + dbData_less16.size());
	      
	      
	      //設定報表表頭資料============================================
	      //列印總機構or分支機構滿16小時
	   	  row = sheet.getRow(1);
	   	  cell = row.getCell( (short) 0);
	   	  String insertValue = "";  
	   	  cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	      if (dbData_over16 != null && dbData_over16.size() != 0) {
		   	  cell.setCellValue("民國"+m_year+"年"+m_month+"月");		   	  
		   	  
		   	  rowNum = 2;	 
		   	  pbank_no = (String)((DataObject)dbData_over16.get(0)).getValue("pbank_no");		   	  
		   	  rowNum_start = rowNum+1;
		   	  rowNum_end  = rowNum+1;
		   	  System.out.println("dbData_over16  begin");
	      	  for(int i=0;i<dbData_over16.size();i++){	      	  
	      	      bean = (DataObject)dbData_over16.get(i);	      	     
	      	      rowNum++;
	    	  	  row = sheet.createRow(rowNum);
   	    	  	  //列印各機構明細資料	    	  	
				  for(int cellcount=0;cellcount<cellname.length;cellcount++){			 	      
			 	     cell=row.createCell((short)cellcount);			 	    
			    	 cell.setEncoding(HSSFCell.ENCODING_UTF_16);	
			    	 cell.setCellStyle(cs_center);
			    	 if(cellcount == 4){
			    	 	insertValue = (bean.getValue(cellname[cellcount]) == null)?"":(bean.getValue(cellname[cellcount])).toString();
			    	 }else{	
			    	 	insertValue = (bean.getValue(cellname[cellcount]) == null)?"":String.valueOf(bean.getValue(cellname[cellcount]));    
			    	 }			    	 
			    	 //System.out.println("insertValue="+insertValue);
			    	 cell.setCellValue(insertValue);			    	 
				  }//end of cellcount
				  //99.10.11 add同一個總機構代號/名稱儲存格合併
			      if(!((String)bean.getValue("pbank_no")).equals(pbank_no)){			    	 	
			    	  rowNum_end = rowNum-1;
			    	  
			    	  //System.out.println("pbank_no="+pbank_no+":rowNum_start="+rowNum_start);
			    	  //System.out.println("rowNum_end="+rowNum_end);
			    	  sheet.addMergedRegion( new Region( ( short )rowNum_start, ( short )0,( short )rowNum_end,( short )0) );
			    	  sheet.addMergedRegion( new Region( ( short )rowNum_start, ( short )1,( short )rowNum_end,( short )1) );
			    	  rowNum_start=rowNum;
			    	  pbank_no = (String)bean.getValue("pbank_no");
			      }
	      	  }//end of bean
	      	  //99.10.11 add 最後一筆merge
	      	  rowNum_end = rowNum;
	    	  sheet.addMergedRegion( new Region( ( short )rowNum_start, ( short )0,( short )rowNum_end,( short )0) );
	    	  sheet.addMergedRegion( new Region( ( short )rowNum_start, ( short )1,( short )rowNum_end,( short )1) );
	    	  
	      }else{ //end of else dbData.size() != 0
	      	 cell.setCellValue("民國"+m_year+"年"+m_month+"月無資料存在");
	      }
	     
	      //列印總機構or分支機構不滿16小時
	      sheet = wb.getSheetAt(1);//讀取第一個工作表，宣告其為sheet
	   	  row = sheet.getRow(1);
	   	  cell = row.getCell( (short) 0);	   	  
	   	  cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	      if (dbData_less16 != null && dbData_less16.size() != 0) {
		   	  cell.setCellValue("民國"+m_year+"年"+m_month+"月");
		   	  rowNum = 2;
		   	  pbank_no = (String)((DataObject)dbData_less16.get(0)).getValue("pbank_no");		   	  
		   	  rowNum_start = rowNum+1;
		   	  rowNum_end  = rowNum+1;
		   	  System.out.println("dbData_less16  begin");
	      	  for(int i=0;i<dbData_less16.size();i++){	      	      
	      	      bean = (DataObject)dbData_less16.get(i);	      	     
	      	      rowNum++;
	    	  	  row = sheet.createRow(rowNum);
	    	  	  //列印各機構明細資料	    	  	
				  for(int cellcount=0;cellcount<cellname.length;cellcount++){			 	      
			 	     cell=row.createCell((short)cellcount);			 	    
			    	 cell.setEncoding(HSSFCell.ENCODING_UTF_16);	
			    	 cell.setCellStyle(cs_center);
			    	 if(cellcount == 4){
			    	 	insertValue = (bean.getValue(cellname[cellcount]) == null)?"":(bean.getValue(cellname[cellcount])).toString();
			    	 }else{	
			    	 	insertValue = (bean.getValue(cellname[cellcount]) == null)?"":String.valueOf(bean.getValue(cellname[cellcount]));    
			    	 }	
			    	 
			    	 //System.out.println("insertValue="+insertValue);
			    	 cell.setCellValue(insertValue);			    	 
				  }//end of cellcount
				  //99.10.11 add同一個總機構代號/名稱儲存格合併
			      if(!((String)bean.getValue("pbank_no")).equals(pbank_no)){			    	 	
			    	  rowNum_end = rowNum-1;
			    	  //System.out.println("pbank_no="+pbank_no+":rowNum_start="+rowNum_start);
			    	  //System.out.println("rowNum_end="+rowNum_end);
			    	  sheet.addMergedRegion( new Region( ( short )rowNum_start, ( short )0,( short )rowNum_end,( short )0) );
			    	  sheet.addMergedRegion( new Region( ( short )rowNum_start, ( short )1,( short )rowNum_end,( short )1) );
			    	  rowNum_start=rowNum;
			    	  pbank_no = (String)bean.getValue("pbank_no");
			      }
	      	  }//end of bean
	      	  //99.10.11 add 最後一筆merge
	      	  rowNum_end = rowNum;
	    	  sheet.addMergedRegion( new Region( ( short )rowNum_start, ( short )0,( short )rowNum_end,( short )0) );
	    	  sheet.addMergedRegion( new Region( ( short )rowNum_start, ( short )1,( short )rowNum_end,( short )1) );
	    	  
	      }else{ //end of else dbData.size() != 0
	      	 cell.setCellValue("民國"+m_year+"年"+m_month+"月無資料存在");
	      }
	     
	      
	      FileOutputStream fout = null;     
	      fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "全體農漁會信用部主任及分部主任參訓情形統計表.xls");
	      HSSFFooter footer = sheet.getFooter();
	      footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
	      footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
	      wb.write(fout);
	      //儲存
	      fout.close();	      
	    }catch (Exception e) {
	      System.out.println("RptFR058W.createRpt Error:" + e + e.getMessage());
	    }
	    
	    return errMsg;
	  }	  
}
