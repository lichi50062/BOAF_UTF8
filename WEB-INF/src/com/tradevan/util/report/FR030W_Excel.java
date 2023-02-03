/*
 * Created on 2005/12/17 by lilic0c0 4183
 *
 * TODO To change the template for this generated file go to
 * Window - Preferences - Java - Code Style - Code Templates
 * FIX ON 20060725 BY 2495
 * 99.04.12 fix 修改SQL以preparedstatement方式查詢 by 2808
 */

package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

/**
 * 信用部支票存款戶數與餘額統計表.
 * 
 * @author 2295
 *
 * TODO To change the template for this generated type comment go to
 * Window - Preferences - Java - Code Style - Code Templates
 */
public class FR030W_Excel {
  public static String createRpt(String S_YEAR, String S_MONTH,
                                 String bank_type,String Unit) {
                                 	
    System.out.println("FR030W_Excel Start ...");
    
	System.out.println("bank_type="+bank_type);
	int u_year = S_YEAR==null? 99 :Integer.parseInt(S_YEAR)  ; //判斷縣市合併用參數
	List paramList = new ArrayList();
	if(u_year >=100 ) {
		u_year = 100 ;
	}else {
		u_year = 99 ;
	}
	DataObject bean = null;
    String errMsg = "";
    List dbData = null;
    //String sqlCmd = "";
    //Properties A02Data = new Properties();

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

      //input the standard report form
      //System.out.println("台灣區農會信用部支票存款戶數與餘額彙計表.xls");
      finput = new FileInputStream(xlsDir +System.getProperty("file.separator") +
                                   "台灣區農會信用部支票存款戶數與餘額彙計表.xls");

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

      //把月份跟年作轉換 去取得上月的月份跟年份
      int n_year = Integer.parseInt(S_YEAR);
      int n_month = Integer.parseInt(S_MONTH);
      int last_year = n_year;
      int last_month = n_month - 1;

      if (last_month == 0) { //本月如果是1月的話 上月為12月
        last_month = 12;
        last_year--;
      }
      StringBuffer sql  = new StringBuffer() ;
      sql.append(" select Sum(CheckBank_Cnt + CheckBank_Cnt_S + CheckBank_Cnt_N) as Check_cnt_TOT,           ");
      sql.append("        Round(Sum(CheckBank_Bal + CheckBank_Bal_S + CheckBank_Bal_N)/?,0) as Check_Bal_TOT,"); 
      sql.append("        Sum(CheckBank_Cnt) as Check_cnt_1,                                                 ");
      sql.append("        Round(Sum(CheckBank_Bal) /? ,0) As Check_bal_1,                                     ");
      sql.append("        Sum(CheckBank_Cnt_S) as Check_cnt_S,                                               ");
      sql.append("        Round(Sum(CheckBank_Bal_S) /?,0) as Check_bal_S,                                   ");
      sql.append("        Sum(CheckBank_Cnt_N) as Check_cnt_N,                                               ");
      sql.append("        Round(Sum(CheckBank_Bal_N) /?,0) as Check_bal_N                                    ");
      sql.append(" from  WLX07_M_CHECKBANK  AA,  (select * from bn01 where bank_type=? and m_year=? ) bN01  ");
      sql.append(" ,v_bank_location T2 ");
      sql.append(" where (AA.m_year=  ?  and AA.M_MONTH =  ? ) ");
      sql.append(" and  AA.BANK_NO=BN01.Bank_No  ");
      sql.append(" and  AA.BANK_NO=T2.Bank_No and T2.m_year=? ");
      sql.append(" ORDER BY T2.fr001w_output_order ");
      paramList.add(Unit) ;
      paramList.add(Unit) ;
      paramList.add(Unit) ;
      paramList.add(Unit) ;
      paramList.add(bank_type) ;
      paramList.add(String.valueOf(u_year)) ;
      paramList.add(S_YEAR) ;
      paramList.add(S_MONTH) ;
      paramList.add(String.valueOf(u_year)) ;
      dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "Check_cnt_TOT,Check_Bal_TOT,Check_cnt_1,Check_bal_1,"+
      									"Check_cnt_S,Check_bal_S,Check_cnt_N,Check_bal_N");
      bean = (DataObject) dbData.get(0);  //修改以bean物件取得資料 by 2808
      //去取得本月以及上月的資料 type= 1 為本月 type = 2 為上月
      
      System.out.println("dbData.size=" + dbData.size());
    
      //設定金額單位(原本是數字1000轉成千元 10000轉成萬元 以此類推)
      Unit = Utility.getUnitName(Unit);//取得單位名稱
      

      //取得當前日期
      //Calendar rightNow = Calendar.getInstance();
      //String year = String.valueOf(rightNow.get(Calendar.YEAR) - 1911);
      //String month = String.valueOf(rightNow.get(Calendar.MONTH) + 1);
      //String day = String.valueOf(rightNow.get(Calendar.DAY_OF_MONTH));


      //設定報表表頭資料============================================
      row = sheet.getRow(0);
      cell = row.getCell( (short) 0);
      cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
      if (bank_type.equals("6")) {
        cell.setCellValue("台灣區農會信用部支票存款戶數與餘額彙計表");
      }
      else {
        cell.setCellValue("台灣區漁會信用部支票存款戶數與餘額彙計表");
      }

      //判斷a1是不是null，是的話表示沒有資料
      if (bean.getValue("check_cnt_tot") == null) {
        row = sheet.getRow(1);
        cell = row.getCell( (short) 1);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16);//設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        //設定無資料
        cell.setCellValue("民國 " + S_YEAR + " 年 " + S_MONTH + " 月底 無資料存在");
      }
      else {
        //設定日期===================================================
        row = sheet.getRow(1);
        cell = row.getCell( (short) 1);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue("民國 " + S_YEAR + " 年 " + S_MONTH + " 月底");
        //設定總計===================================================
        row = sheet.getRow(3);
        cell = row.getCell( (short) 4);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( bean.getValue("check_cnt_tot").toString()));
        row = sheet.getRow(4);
        cell = row.getCell( (short) 4);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( bean.getValue("check_bal_tot").toString()));
        //正會員===================================================
        row = sheet.getRow(5);
        cell = row.getCell( (short) 4);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( bean.getValue("check_cnt_1").toString()));
        row = sheet.getRow(6);
        cell = row.getCell( (short) 4);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( bean.getValue("check_bal_1").toString()));
        //贊助會員===================================================
        row = sheet.getRow(7);
        cell = row.getCell( (short) 4);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( bean.getValue("check_cnt_s").toString()));
        row = sheet.getRow(8);
        cell = row.getCell( (short) 4);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( bean.getValue("check_bal_s").toString()));
        //非會員===================================================
        row = sheet.getRow(9);
        cell = row.getCell( (short) 4);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( bean.getValue("check_cnt_n").toString()));
        row = sheet.getRow(10);
        cell = row.getCell( (short) 4);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( bean.getValue("check_bal_n").toString()));

      } //end of else dbData.size() == 0

      FileOutputStream fout = null;
      if (bank_type.equals("6")) {
        fout = new FileOutputStream(reportDir +System.getProperty("file.separator") + "台灣區農會信用部支票存款戶數與餘額彙計表.xls");
      }
      else {
        fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "台灣區漁會信用部支票存款戶數與餘額彙計表.xls");
      }
      HSSFFooter footer = sheet.getFooter();
      footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
      footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));	

      wb.write(fout);
      //儲存
      fout.close();

    }
    catch (Exception e) {
      System.out.println("createRpt Error:" + e + e.getMessage());
    }

    System.out.println("FR030W_Excel End ...");
    return errMsg;
  }
}
