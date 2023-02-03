/*
 * Created on 2007/07/24 存款帳戶分級差異化管理統計表_全國農業金庫 by 2295
 * 100.02.16 fix 根據查詢年度.100年以後取得新縣市別.100年以前取得舊縣市別 
 * 				 使用PreparedStatement;並列印轉換後的SQL;套用QueryDB_SQLParam by 2295
 */

package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;


public class RptFR045W_bank_1 {
  public static String createRpt(String S_YEAR, String S_MONTH,String E_YEAR,String E_MONTH) {    

    String errMsg = "";
    List dbData = null;
    String sqlCmd = "";  
    int rowNum=0;
    DataObject bean = null;
    reportUtil reportUtil = new reportUtil();
	HSSFCellStyle cs_right = null; 
	HSSFCellStyle cs_center = null;
	HSSFCellStyle cs_left = null; 
	HSSFCellStyle cs_noborderleft = null; 
    String m_year="";
    String m_month="";
    String warnaccount_cnt="";
	String limitaccount_cnt="";
	String erroraccount_cnt="";
	String otheraccount_cnt="";
	String depositaccount_tcnt="";
	String bn01_m_year = "";
	List paramList = new ArrayList();
    try {
      //100.02.16 add 查詢年度100年以前.縣市別不同===============================  	     
      bn01_m_year = (Integer.parseInt(S_YEAR) < 100)?"99":"100"; 
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

      //input the standard report form      
      finput = new FileInputStream(xlsDir +System.getProperty("file.separator") +"存款帳戶分級差異化管理統計表_全國農業金庫.xls");

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

      short i = 0;
      short y = 0;
      cs_right = reportUtil.getRightStyle(wb);
      cs_left = reportUtil.getLeftStyle(wb);
      cs_center = reportUtil.getDefaultStyle(wb);
      cs_noborderleft = reportUtil.getNoBorderLeftStyle(wb);
      int data_year = Integer.parseInt(E_YEAR);
      int data_month = Integer.parseInt(E_MONTH);      
      int last_year = data_year;
      int last_month = data_month - 1;
      int quarter_year = data_year;
      int quarter_month = data_month;
      if (last_month == 0) { //本月如果是1月的話 上月為12月
        last_month = 12;
        last_year--;
      }
      /*全國農業金庫與上月底之比較
       select bank_code,bn01.bank_name,m_year,m_month,
       		  warnaccount_cnt,
       		  limitaccount_cnt,
       		  erroraccount_cnt,
       		  otheraccount_cnt,
       		  depositaccount_tcnt
       from  a08 left join bn01 on a08.bank_code = bn01.bank_no
       where bank_code='0180012'
       and to_char(m_year * 100 + m_month) >= 9508
       and to_char(m_year * 100 + m_month) <= 9512  
       union
       select nowA08.bank_code,bn01.bank_name,99 as m_year,99 as m_month,       
       		  nowA08.warnaccount_cnt - lastA08.warnaccount_cnt as last_warnaccount_cnt,
       		  nowA08.limitaccount_cnt - lastA08.limitaccount_cnt as last_limitaccount_cnt,    
       		  nowA08.erroraccount_cnt - lastA08.erroraccount_cnt as last_erroraccount_cnt,    
       		  nowA08.otheraccount_cnt - lastA08.otheraccount_cnt as last_otheraccount_cnt,    
       		  nowA08.depositaccount_tcnt - lastA08.depositaccount_tcnt as last_depositaccount_tcnt       		    
       from(select * from a08 where m_year=95 and m_month=12 and bank_code='0180012')nowA08
       left join (select * from a08 where m_year=95 and m_month=11)lastA08 on nowA08.bank_code = lastA08.bank_code
       left join bn01 on nowA08.bank_code = bn01.bank_no
       order by m_year,m_month	 
       */
     
      
      sqlCmd = " select bank_code,bn01.bank_name,a08.m_year,m_month,"
		  	 + " 	    warnaccount_cnt,"
		     + "		limitaccount_cnt,"
			 + "		erroraccount_cnt,"
			 + "		otheraccount_cnt,"
			 + "		depositaccount_tcnt"
			 + " from  a08 left join (select * from bn01 where m_year=?)bn01 on a08.bank_code = bn01.bank_no"
			 + " where bank_code='0180012'"
			 + " and to_char(a08.m_year * 100 + m_month) >= ?"
			 + " and to_char(a08.m_year * 100 + m_month) <= ?"  
			 + " union "
			 + " select nowA08.bank_code,bn01.bank_name,99 as m_year,99 as month,"       
		     + "		nowA08.warnaccount_cnt - lastA08.warnaccount_cnt as last_warnaccount_cnt,"
		     + "		nowA08.limitaccount_cnt - lastA08.limitaccount_cnt as last_limitaccount_cnt,"    
			 + "		nowA08.erroraccount_cnt - lastA08.erroraccount_cnt as last_erroraccount_cnt,"    
			 + "		nowA08.otheraccount_cnt - lastA08.otheraccount_cnt as last_otheraccount_cnt,"    
			 + "		nowA08.depositaccount_tcnt - lastA08.depositaccount_tcnt as last_depositaccount_tcnt"	       		    
			 + " from(select * from a08 where m_year=? and m_month=? and bank_code='0180012')nowA08"
			 + " left join (select * from a08 where m_year=? and m_month=?)lastA08 on nowA08.bank_code = lastA08.bank_code"
			 + " left join (select * from bn01 where m_year=?)bn01 on nowA08.bank_code = bn01.bank_no"
			 + " order by m_year,m_month";	 
      paramList.add(bn01_m_year);
      paramList.add(S_YEAR+S_MONTH);
      paramList.add(E_YEAR+E_MONTH);
      paramList.add(E_YEAR);
      paramList.add(E_MONTH);
      paramList.add(String.valueOf(last_year));
      paramList.add(String.valueOf(last_month));
      paramList.add(bn01_m_year);
      
      System.out.println("sqlCmd="+sqlCmd);
      dbData = DBManager.QueryDB_SQLParam(sqlCmd,paramList,"m_year,m_month,warnaccount_cnt,limitaccount_cnt,erroraccount_cnt,otheraccount_cnt,depositaccount_tcnt");

      
      System.out.println("dbData.size=" + dbData.size());
      
      if (dbData == null || dbData.size() ==0) {
      	  row=sheet.getRow(2);
          cell=row.getCell((short)0);	       	
          cell.setEncoding(HSSFCell.ENCODING_UTF_16);
          cell.setCellValue("無資料存在");         
      }else {
      	rowNum = 6;
        for(int idx=0;idx<dbData.size();idx++){
        	//System.out.println("idx="+idx);
        	bean = (DataObject)dbData.get(idx);
        	m_year=(bean.getValue("m_year") == null)?"":(bean.getValue("m_year")).toString();
        	m_month=(bean.getValue("m_month") == null)?"":(bean.getValue("m_month")).toString();
        	warnaccount_cnt=(bean.getValue("warnaccount_cnt") == null)?"":(bean.getValue("warnaccount_cnt")).toString();
        	limitaccount_cnt=(bean.getValue("limitaccount_cnt") == null)?"":(bean.getValue("limitaccount_cnt")).toString();
        	erroraccount_cnt=(bean.getValue("erroraccount_cnt") == null)?"":(bean.getValue("erroraccount_cnt")).toString();
        	otheraccount_cnt=(bean.getValue("otheraccount_cnt") == null)?"":(bean.getValue("otheraccount_cnt")).toString();
        	//depositaccount_tcnt=(bean.getValue("depositaccount_tcnt") == null)?"":(bean.getValue("depositaccount_tcnt")).toString();
        	
        	row = sheet.createRow(rowNum);
		    for(int cellcount=0;cellcount<5;cellcount++){
	 		    cell = row.createCell( (short)cellcount);
	     		cell.setEncoding(HSSFCell.ENCODING_UTF_16);	 	
	     		if(cellcount == 0 ){
	     		   cell.setCellStyle(cs_center);
	     		}else{
	     		   cell.setCellStyle(cs_right);
	     		}
	     		if(!m_month.equals("99")){
	     			if(cellcount == 0) cell.setCellValue(m_year+"年"+m_month+"月底戶數");
	     			if(cellcount == 1 && !warnaccount_cnt.equals("")) cell.setCellValue(Utility.setCommaFormat(warnaccount_cnt));	 
	     			if(cellcount == 2 && !limitaccount_cnt.equals("")) cell.setCellValue(Utility.setCommaFormat(limitaccount_cnt));	 
	     			if(cellcount == 3 && !erroraccount_cnt.equals("")) cell.setCellValue(Utility.setCommaFormat(erroraccount_cnt));
	     			if(cellcount == 4 && !otheraccount_cnt.equals("")) cell.setCellValue(Utility.setCommaFormat(otheraccount_cnt));	     			
	     		}
			}//end of cellcount	 		  	  	 	
		    rowNum++;
        }//end of dbData        
      } //end of else dbData.size() != 0
      
      row=sheet.createRow(rowNum);
      cell=row.createCell((short)0);	       	
      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      cell.setCellValue("備註：");
      rowNum++;
      row=sheet.createRow(rowNum);
      cell=row.createCell((short)0);	       	
      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      cell.setCellValue("1.存款帳戶總戶數為(A)+(B)+(C)+(D)合計數。");
      rowNum++;
      row=sheet.createRow(rowNum);
      cell=row.createCell((short)0);	       	
      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      cell.setCellValue("2.存款帳戶總戶數指:支票存款、定期存款及活期性存款帳戶，含外匯活期存款及行員存款。");      
      rowNum++;
      row=sheet.createRow(rowNum);
      cell=row.createCell((short)0);	       	
      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      cell.setCellValue("3.本表及警示戶數異動情形分析(如附表)，應於次月15日前上線填報");
     
      rowNum++;
      rowNum++;
      row=sheet.createRow(rowNum);
      cell=row.createCell((short)0);     
      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      cell.setCellValue("附表");
      
      rowNum++;
      row = sheet.createRow(rowNum);
      for(int cellcount=0;cellcount<5;cellcount++){
		  cell = row.createCell( (short)cellcount);
 		  cell.setEncoding(HSSFCell.ENCODING_UTF_16);
 		  cell.setCellStyle(cs_center);
      }	
      cell = row.getCell( (short)0);
      cell.setCellValue("警示戶數異動情形分析");
      sheet.addMergedRegion(new Region((short) rowNum, (short) 0, (short) rowNum, (short) 4));

      rowNum++;
      printTitle(row,cell,cs_left,cs_center,rowNum,sheet,E_YEAR+"年"+E_MONTH+"月份警示帳戶新增戶數",warnaccount_cnt);      
      rowNum++;
      printTitle(row,cell,cs_left,cs_center,rowNum,sheet,"新增警示帳戶數屬上月(B)類洐生管制帳戶者",limitaccount_cnt);      
      rowNum++;
      printTitle(row,cell,cs_left,cs_center,rowNum,sheet,"新增警示帳戶數屬上月(C)類自行篩選異常者",erroraccount_cnt);      
      rowNum++;
      printTitle(row,cell,cs_left,cs_center,rowNum,sheet,"新增警示帳戶數屬上月(D)類其他帳戶者",otheraccount_cnt);     
      rowNum++;
      rowNum++;
      row=sheet.createRow(rowNum);
      cell=row.createCell((short)0);      
      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      cell.setCellValue("填表人：");
    
      cell=row.createCell((short)3);      
      cell.setEncoding(HSSFCell.ENCODING_UTF_16);
      cell.setCellValue("聯絡電話：");
      
      FileOutputStream fout = null;     
      fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "存款帳戶分級差異化管理統計表_全國農業金庫.xls");
     
      HSSFFooter footer = sheet.getFooter();
      footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
      footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
      wb.write(fout);
      //儲存
      fout.close();
    }
    catch (Exception e) {
      System.out.println("RptFR045W_bank_1.createRpt Error:" + e + e.getMessage());
    }
    
    return errMsg;
  }
  private static void printTitle(HSSFRow row,HSSFCell cell,HSSFCellStyle cs_left,HSSFCellStyle cs_center,int rowNum,HSSFSheet sheet,String title,String amt){
  	try{
  	row = sheet.createRow(rowNum);
    for(int cellcount=0;cellcount<5;cellcount++){
		  cell = row.createCell( (short)cellcount);
		  cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		  if(cellcount < 3){
		     cell.setCellStyle(cs_left);
		  }else{
		  	 cell.setCellStyle(cs_center);
		  }
    }	
    cell = row.getCell( (short)0);
    cell.setCellValue(title);
    sheet.addMergedRegion(new Region((short) rowNum, (short) 0, (short) rowNum, (short) 2));
    cell = row.getCell( (short)3);
    cell.setCellValue(Utility.setCommaFormat(amt)+"戶");
    sheet.addMergedRegion(new Region((short) rowNum, (short) 3, (short) rowNum, (short) 4));
  	}catch(Exception e){
  		System.out.println("printTitle Error:"+e+e.getMessage());
  	}
  }
}
