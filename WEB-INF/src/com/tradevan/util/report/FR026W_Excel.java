/*
 * Created on 2005/11/23 by lilic0c0 4183
 *
 * TODO To change the template for this generated file go to
 * Window - Preferences - Java - Code Style - Code Templates
 * 
 * 99.04.12 fix SQL以preparedstatement方式查詢 by 2808
 */

package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.util.*;
import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;


public class FR026W_Excel {
	
  public static String createRpt(String S_YEAR, String S_MONTH,
                                 String bank_type,String S_Bank_Name,String Unit) {
                                 	
    System.out.println("FR026W_Excel Start ...");

    String errMsg = "";
    List dbData = null;
    List paramList = new ArrayList();
    DataObject bean = null;
    String u_year = "100" ;
    if("".equals(S_YEAR) || Integer.parseInt(S_YEAR)<=99) {
    	u_year = "99" ;
    }
    try {
      String simpleNm = "在台無住所之外國新台幣存款表.xls" ;
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
      System.out.println("xlsDir:"+xlsDir) ;
      System.out.println("System.getProperty(file.separator):"+System.getProperty("file.separator")) ;
      System.out.println("open 範本路徑:"+xlsDir + System.getProperty("file.separator")+ simpleNm) ;
	  FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+ simpleNm );
	  System.out.println("Open excel 完成");
                                        
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


      //去取得本月的資料
      StringBuffer sql = new StringBuffer();
      sql.append(" select F01.M_YEAR, M_MONTH, BANK_CODE, bank_name, DEP_TYPE,  ACCT_TYPE,");
      sql.append("        sum(ACCT_CNT_TM)  AS  A_CNT,                                    ");
      sql.append("        round(sum(BAL_LM)/?,0)       AS  B_BAL,                         ");
      sql.append("        round(SUM(DEP_TM)/?,0)       AS  C_DEP,                         ");
      sql.append("        round(SUM(WTD_TM)/?,0)       AS  D_WTD,                         ");
      sql.append("        round(SUM(BAL_TM)/?,0)       AS  E_BAL                          ");
      sql.append(" from   F01 ,  bn01                                                     ");
      sql.append(" where  F01.M_YEAR  =  ?                                               ");
      sql.append(" and    F01.M_MONTH =  ?                                                ");
      sql.append(" and    F01.BANK_CODE  = ?                                      ");
      sql.append(" and    (f01.BANK_CODE =bn01.bank_no  and bn01.m_year=? and bn01.bank_type=?)");// --以bn01.m_year來取得查詢年度新機構名稱
      sql.append(" group  by  F01.M_YEAR, M_MONTH, BANK_CODE, bank_name,  DEP_TYPE,  ACCT_TYPE     ");           
      sql.append(" ORDER  by  F01.M_YEAR, M_MONTH, BANK_CODE, bank_name,  DEP_TYPE,  ACCT_TYPE     ");
      paramList.add(Unit);
      paramList.add(Unit);
      paramList.add(Unit);
      paramList.add(Unit);
      paramList.add(S_YEAR);
      paramList.add(S_MONTH);
      paramList.add(S_Bank_Name);
      paramList.add(u_year) ;
      paramList.add(bank_type);
      dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "M_YEAR,M_MONTH,A_CNT,B_BAL,C_DEP,D_WTD,E_BAL"); 

      System.out.println("dbData.size=" + dbData.size());
      
      
      
      //設定金額單位(原本是數字1000轉成千元 10000轉成萬元 以此類推)
      Unit = Utility.getUnitName(Unit) ;
      
      //設定報表表頭資料============================================
      row = sheet.getRow(0);
      cell = row.getCell( (short) 0);
      cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
      cell.setCellValue("在台無住所之外國人新台幣存款表");
      
      //設定金融機構資料============================================
      row = sheet.getRow(2);
      cell = row.getCell( (short) 0);
      cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
      cell.setCellValue("金融機構名稱 ："+((DataObject) dbData.get(0)).getValue("bank_name"));
	  cell = row.getCell( (short) 5);
      cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
      cell.setCellValue("單位：新台幣"+Unit);

      //判斷dbData.size()是不是0，是的話表示沒有資料
      if (dbData.size() == 0) {
        row = sheet.getRow(1);
        cell = row.getCell( (short) 0);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16);//設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        //設定無資料
        cell.setCellValue("民國 " + S_YEAR + " 年 " + S_MONTH + " 月 無資料存在");
      }
      else {
        //設定日期===================================================
        row = sheet.getRow(1);
        cell = row.getCell( (short) 0);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue("民國 " + S_YEAR + " 年 " + S_MONTH + " 月底");
        
        //設定儲存格資料============================================
        for(int i=0;i < dbData.size();i++){
        	bean = (DataObject) dbData.get(i);
        	row = sheet.getRow(i+5);
        	cell = row.getCell( (short) 2);
        	cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        	cell.setCellValue(bean.getValue("a_cnt").toString());
        	
        	cell = row.getCell( (short) 3);
        	cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        	cell.setCellValue(Utility.setCommaFormat( bean.getValue("b_bal").toString()));
        	
        	cell = row.getCell( (short) 4);
        	cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        	cell.setCellValue(Utility.setCommaFormat( bean.getValue("c_dep").toString()));
        	
        	cell = row.getCell( (short) 5);
        	cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        	cell.setCellValue(Utility.setCommaFormat( bean.getValue("d_wtd").toString()));
        	
        	cell = row.getCell( (short) 6);
        	cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        	cell.setCellValue(Utility.setCommaFormat( bean.getValue("e_bal").toString()));
        }//end of for

      } //end of else ((DataObject) dbData.get(0)).getValue("m_year") == null

      FileOutputStream fout = null;
      if (bank_type.equals("6")) {
    	  fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "在台無住所之外國人新台幣存款表-農會.xls");
      }
      else {
    	  fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "在台無住所之外國人新台幣存款表-漁會.xls");
      }
      HSSFFooter footer = sheet.getFooter();
      footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
      footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));	

      wb.write(fout);
      //儲存
      fout.close();
      System.out.println("儲存完成");
    }
    catch (Exception e) {
      System.out.println("createRpt Error:" + e + e.getMessage());
    }

    System.out.println("FR026W_Excel End ...");
    return errMsg;
  }
}
