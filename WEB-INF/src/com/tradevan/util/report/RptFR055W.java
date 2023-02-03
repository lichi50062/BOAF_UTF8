/*
 *  98.06.23 農漁會信用部理監事基本資料表  by 2756
 *  98.06.29 add 機構名稱獨立Query by 2295
 * 100.09.15 fix 1.根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
 * 				   使用PreparedStatement;並列印轉換後的SQL;套用QueryDB_SQLParam by 2295
 *               2.已卸任的不顯示 by 2295
 * 102.10.09 add IDN加密後,還原顯示 by 2295              
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

public class RptFR055W {	 
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
	    String bank_name="";
	    String position="";//職稱
	    String name="";//姓名
	    String id="";//身分證字號
	    String sex="";//性別
	    String birthday="";//生日
	    String inductDate="";//原就任日期
	    String appointNum="";//本屆屆期
	    String abdicateDate="";//卸任日期
	    String BegDate="";//本屆任期(起日)
	    String EndDate="";//本屆任期(迄日)
	    String degree="";//學歷
	    String background="";//經歷
	    String phoneNo="";//聯絡電話 
	    String filename="";
	    List paramList = new ArrayList();
	    try {
	      	
	      //100.09.15 add 查詢年度100年以前.縣市別不同===============================
		  String cd01_table = (Integer.parseInt(Utility.getYear()) < 100)?"cd01_99":""; 
		  String wlx01_m_year = (Integer.parseInt(Utility.getYear()) < 100)?"99":"100"; 
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

	      filename="理監事基本資料_檢查局.xls";
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
	      paramList.add(wlx01_m_year);
	      paramList.add(bank_no);
	      dbData = DBManager.QueryDB_SQLParam("select bank_name from ba01 where m_year=? and bank_no=?",paramList,"");
          if(dbData.size() != 0){
          	tbank_no_name  = (String)((DataObject)dbData.get(0)).getValue("bank_name");
          }	
         
	      StringBuffer sb=new StringBuffer();
	      sb.append(" select  bn01.bank_name,wlx04.position_code,cdshareno.cmuse_name as position_code_name,");
	      sb.append(" wlx04.name,wlx04.id,wlx04.sex,to_char(wlx04.birth_date,'YYYY-MM-DD') as birth_date,");
	      sb.append(" to_char(wlx04.induct_date,'YYYY-MM-DD') as induct_date,wlx04.appointed_num,");
	      sb.append(" to_char(wlx04.abdicate_date,'YYYY-MM-DD') as abdicate_date,to_char(wlx04.period_start,'YYYY-MM-DD') as period_start,");
	      sb.append(" to_char(wlx04.period_end,'YYYY-MM-DD') as period_end,wlx04.degree,wlx04.background,wlx04.telno");
	      sb.append(" from  wlx04 left join (select * from bn01 where m_year=?)bn01 on wlx04.bank_no = bn01.bank_no ");
	      sb.append(" left join cdshareno on wlx04.position_code = cdshareno.cmuse_id and cdshareno.cmuse_div='008'");
	     
    	  if(!bank_no.equals("")) {
      	    condition += (condition.length() > 0 ? " and":"")+" wlx04.bank_no in (?) and abdicate_code <> 'Y'";		
      	  }
    	  sqlCmd=sb.toString();
       	  if(condition.length() > 0) sqlCmd += " where "+condition;
       	 
	      dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"");
	      System.out.println("dbData.size=" + dbData.size());

	      //設定表頭資料
	      row=sheet.getRow(0);
	      cell=row.getCell((short)0);
	      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	      cell.setCellValue(tbank_no_name+"理監事基本資料表");
	      row=sheet.getRow(1);
	      cell=row.getCell((short)0);
	      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	      
	      //設定產生日期
	      Calendar c = Calendar.getInstance();
	      
	      
          if(dbData != null && dbData.size() != 0) 
          {	 

        	  cell.setCellValue("產生日期：中華民國"+(c.get(Calendar.YEAR)-1911)+"年"+(c.get(Calendar.MONTH)+1)+"月"+c.get(Calendar.DAY_OF_MONTH)+"日");
	      }
          else
          {
	          	cell.setCellValue("無資料");	
	      }
          //置放表格內容資料          
	      if (dbData != null && dbData.size() != 0) 
	      {	          
	    	  rowNum = 3;//從表格的第三列開始放置第一筆資料(因編號0~2皆為表頭資料)
	      	  System.out.println("print detail");
	      	  for(int i=0;i<dbData.size();i++)
	      	  {
	      		  bean = (DataObject)dbData.get(i);
	      		  bank_name=(bean.getValue("bank_name") == null)?"":(String)bean.getValue("bank_name");
	      		  position=(bean.getValue("position_code_name") == null)?"":(String)bean.getValue("position_code_name");
	      		  name=(bean.getValue("name") == null)?"":(String)bean.getValue("name");
	      		  //102.10.09 add IDN加密後,還原顯示 by 2295
	      		  id=(bean.getValue("id") == null)?"":Utility.decode((String)bean.getValue("id"));
	      		  sex=(bean.getValue("sex") == null)?"":(String)bean.getValue("sex");
	      		  birthday=(bean.getValue("birth_date") == null)?"":(String)bean.getValue("birth_date");
	      		  inductDate=(bean.getValue("induct_date") == null)?"":(String)bean.getValue("induct_date");
	      		  appointNum=(bean.getValue("appointed_num") == null)?"":(String)bean.getValue("appointed_num");
	      		  abdicateDate=(bean.getValue("abdicate_date") == null)?"":(String)bean.getValue("abdicate_date");
	      		  BegDate=(bean.getValue("period_start") == null)?"":(String)bean.getValue("period_start");
	      		  EndDate=(bean.getValue("period_end") == null)?"":(String)bean.getValue("period_end");
	      		  degree=(bean.getValue("degree") == null)?"":(String)bean.getValue("degree");
	      		  background=(bean.getValue("background") == null)?"":(String)bean.getValue("background");
	      		  phoneNo=(bean.getValue("telno") == null)?"":(String)bean.getValue("telno");
	      		   
	               row = sheet.createRow(rowNum);
   	    	  	   //列印各機構明細資料
				   for(int cellcount=0;cellcount<=12;cellcount++){			 	      
			 	       cell=row.createCell((short)cellcount);			 		
			    	   cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			    	   cell.setCellStyle(cs_center);	

			    	   if(cellcount == 0) cell.setCellValue(position);//職稱
			    	   if(cellcount == 1) cell.setCellValue(name);//姓名	
			    	   if(cellcount == 2) cell.setCellValue(id);//身分證字號	
			    	   if(cellcount == 3) cell.setCellValue(sex);//性別
			    	   if(cellcount == 4) cell.setCellValue(birthday);//生日
			    	   if(cellcount == 5) cell.setCellValue(inductDate);//就任日期
			    	   if(cellcount == 6) cell.setCellValue(appointNum);//本屆屆期	
			    	   if(cellcount == 7) cell.setCellValue(abdicateDate);//卸任日期	
			    	   if(cellcount == 8) cell.setCellValue(BegDate);//本屆任期-起始	 
			    	   if(cellcount == 9) cell.setCellValue(EndDate);//本屆任期-結束
			    	   if(cellcount == 10) cell.setCellValue(degree);//學歷
			    	   if(cellcount == 11) cell.setCellValue(background);//經歷	
			    	   if(cellcount == 12) cell.setCellValue(phoneNo);//聯絡電話				    	   
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
	      System.out.println("RptFR055W.createRpt Error:" + e + e.getMessage());
	    }
	    
	    return errMsg;
	  }
}
