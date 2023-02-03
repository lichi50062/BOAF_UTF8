/*
 * Created on 2005/11/17 by lilic0c0 4183
 * 96.11.29 fix 調整報表格式.增加金融卡本月交易次數/金融卡本月交易金額(元)/本年累計交易次數/本年累計交易金額(元) by 2295
 * 				晶片卡+金融卡合併計算.發停/停卡/流通張數 by 2295
 * 99.12.09 fix sqlInjection by 2808
 * 101.7.26 fix 查詢頁修改+機構中英文 by 2968
 */ 

package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.util.*;
import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;


public class FR025W_Excel {
  public static String createRpt(String S_YEAR, String S_MONTH,
                                 String bank_type,String Unit) {
    System.out.println("FR025W_Excel Start ...");

    String errMsg = "";
    List dbData = null;
    StringBuffer sqlCmd = new StringBuffer();
    Properties A02Data = new Properties();
    String yy = Integer.parseInt(S_YEAR)>99 ?"100" :"99" ;
    String cd01Table =  Integer.parseInt(S_YEAR)>99 ?"cd01" :"cd01_99" ;
    List paramList = new ArrayList() ;
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
      System.out.println("全體農(漁)會信用部自動化機器彙計.xls");
      finput = new FileInputStream(xlsDir +
                                   System.getProperty("file.separator") +
                                   "全體農(漁)會信用部自動化機器彙計.xls");

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

      //去取得本月以及上月的資料 type= 1 為本月 type = 2 為上月
      sqlCmd.append("select  '1' as  a_type, "
          	 + "SUM(ATM_CNT)                                 		AS  atmcnt, "//ATM-機器台數(歷累計數)
			 + "SUM(PUSH_DebitCard_CNT + PUSH_BinCard_CNT)   		AS  push_cnt, "//金融卡-發卡張數
			 + "SUM(CANC_DebitCard_CNT + CANC_BinCard_CNT)   		AS  cancel_cnt, "//金融卡-停卡張數
			 + "SUM(USE_DebitCard_CNT  + USE_BinCard_CNT)   		AS  use_cnt, "//金融卡-流通張數 
			 + "round((SUM(DebitCard_Month_Tran_AMT)/?),0)   AS  debitcard_monthamt, "//金融卡-交易金額.本月合計
			 + "round((SUM(DebitCard_Year_AccTran_AMT)/?),0) AS  debitcard_yearamt, "//金融卡-交易金額.本年累計
          	 + "SUM(debitcard_month_tran_cnt)                       AS  debitcard_mtrancnt, "//金融卡-交易次數.本月合計
			 + "SUM(debitcard_year_acctran_cnt)                     AS  debitcard_ytrancnt, "//金融卡-交易次數.本年累計
			 + "round((SUM(Month_Tran_AMT)/?),0)      		AS  monthamt, "//ATM-交易金額.本月合計
			 + "round((SUM(Year_AccTran_AMT)/?),0)    		AS  yearamt, "//ATM-交易金額.本年累計
			 + "SUM(Month_Tran_CNT)                          		AS  mtrancnt, "//ATM-交易次數.本月合計
			 + "SUM(Year_AccTran_CNT)                        		AS  ytrancnt "//ATM-交易次數.本年累計
			 + "from  WLX05_M_ATM  ,  (select * from ba01 where m_year=?)ba01 "
			 + "where (WLX05_M_ATM.m_year=  ?"
			 + " and WLX05_M_ATM.M_MONTH = ? ) and "
			 + "(WLX05_M_ATM.BANK_NO=ba01.bank_no AND ba01.bank_type=?) "
			 + "union all "
			 + "select  '2' as  a_type, "
			 + "SUM(ATM_CNT)                                 		AS  atmcnt, "
			 + "SUM(PUSH_DebitCard_CNT + PUSH_BinCard_CNT)   		AS  push_cnt, "
			 + "SUM(CANC_DebitCard_CNT + CANC_BinCard_CNT)   		AS  cancel_cnt, "
			 + "SUM(USE_DebitCard_CNT  + USE_BinCard_CNT)   		AS  use_cnt, "
			 + "round((SUM(DebitCard_Month_Tran_AMT)/?),0)   AS  debitcard_monthamt, "
			 + "round((SUM(DebitCard_Year_AccTran_AMT)/?),0) AS  debitcard_yearamt, "
          	 + "SUM(debitcard_month_tran_cnt)                       AS  debitcard_mtrancnt, "
			 + "SUM(debitcard_year_acctran_cnt)                     AS  debitcard_ytrancnt, "
			 + "round((SUM(Month_Tran_AMT)/?),0)      		AS  monthamt, "
			 + "round((SUM(Year_AccTran_AMT)/?),0)    		AS  yearamt, "
			 + "SUM(Month_Tran_CNT)                          		AS  mtrancnt, "
			 + "SUM(Year_AccTran_CNT)                        		AS  ytrancnt "
			 + "from  WLX05_M_ATM  , (select * from ba01 where m_year=?) ba01 "
			 + "where (WLX05_M_ATM.m_year= ?" 
			 + " and WLX05_M_ATM.M_MONTH = ?) and "
			 + "(WLX05_M_ATM.BANK_NO=ba01.bank_no AND ba01.bank_type=? ) ");
      paramList.add(Unit) ;
      paramList.add(Unit) ;
      paramList.add(Unit) ;
      paramList.add(Unit) ;
      paramList.add(yy) ;
      paramList.add(String.valueOf(n_year)) ;
      paramList.add(String.valueOf(n_month));
      paramList.add(bank_type);
      paramList.add(Unit) ;
      paramList.add(Unit) ;
      paramList.add(Unit) ;
      paramList.add(Unit) ;
      paramList.add(yy) ;
      paramList.add(String.valueOf(last_year)) ;
      paramList.add(String.valueOf(last_month)) ;
      paramList.add(bank_type);
      dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"atmcnt,push_cnt,cancel_cnt,use_cnt,debitcard_monthamt,debitcard_yearamt,debitcard_mtrancnt,debitcard_ytrancnt,monthamt,yearamt,mtrancnt,ytrancnt");

      System.out.println("dbData.size=" + dbData.size());

      //設定金額單位(原本是數字1000轉成千元 10000轉成萬元 以此類推)
      if(Unit.compareTo("1") == 0 )
        Unit = "元";
      else if(Unit.compareTo("1000") == 0 )
        Unit = "千元";
      else if(Unit.compareTo("10000") == 0 )
        Unit = "萬元";
      else if(Unit.compareTo("1000000") == 0 )
        Unit = "百萬元";
      else if(Unit.compareTo("10000000") == 0 )
        Unit = "仟萬元";
      else if(Unit.compareTo("100000000") == 0 )
        Unit = "億元";

      //取得當前日期
      Calendar rightNow = Calendar.getInstance();
      String year = String.valueOf(rightNow.get(Calendar.YEAR) - 1911);
      String month = String.valueOf(rightNow.get(Calendar.MONTH) + 1);
      String day = String.valueOf(rightNow.get(Calendar.DAY_OF_MONTH));


      //設定報表表頭資料============================================
      row = sheet.getRow(0);
      cell = row.getCell( (short) 0);
      cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
      if (bank_type.equals("6")) {
        cell.setCellValue("全體農會信用部自動化服務機器(CD/ATM)彙計");
      }
      else {
        cell.setCellValue("全體漁會信用部自動化服務機器(CD/ATM)彙計");
      }

      //判斷atmcnt是不是null，是的話表示沒有資料
      if (((DataObject) dbData.get(0)).getValue("atmcnt") == null) {
      	System.out.println("atmcnt == null");
        row = sheet.getRow(1);
        cell = row.getCell( (short) 0);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16);//設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        //設定無資料
        cell.setCellValue("民國 " + S_YEAR + " 年 " + S_MONTH + " 月 無資料存在");
      }
      else {
      	int thisMonth =0;
      	int lastMonth =0;
      	DataObject bean0= (DataObject)dbData.get(0);
        DataObject bean1= (DataObject)dbData.get(1);
        //設定日期===================================================
        row = sheet.getRow(1);
        cell = row.getCell( (short) 0);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue("民國 " + S_YEAR + " 年 " + S_MONTH + " 月底");
        //設定機器台數資料============================================
        row = sheet.getRow(3);
        cell = row.getCell( (short) 4);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( ((bean0.getValue("atmcnt") == null)?0:bean0.getValue("atmcnt")).toString()) );
        cell = row.getCell( (short) 5);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( ((bean1.getValue("atmcnt") == null)?0:bean1.getValue("atmcnt")).toString()) );
        cell = row.getCell( (short) 6);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        thisMonth = Integer.parseInt( ((bean0.getValue("atmcnt") == null)?0:bean0.getValue("atmcnt")).toString() );
        lastMonth = Integer.parseInt( ((bean1.getValue("atmcnt") == null)?0:bean1.getValue("atmcnt")).toString() );
        cell.setCellValue(Utility.setCommaFormat(String.valueOf(thisMonth-lastMonth)));
        //ATM自動化機器.交易金額.本月合計============================================
        row = sheet.getRow(4);
        cell = row.getCell( (short) 3);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue("新台幣 "+Unit);
        cell = row.getCell( (short) 4);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( ((bean0.getValue("monthamt") == null)?0:bean0.getValue("monthamt")).toString()) );
        cell = row.getCell( (short) 5);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( ((bean1.getValue("monthamt") == null)?0:bean1.getValue("monthamt")).toString()) );
        cell = row.getCell( (short) 6);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        thisMonth = Integer.parseInt(((bean0.getValue("monthamt") == null)?0:bean0.getValue("monthamt")).toString());
        lastMonth = Integer.parseInt(((bean1.getValue("monthamt") == null)?0:bean1.getValue("monthamt")).toString());
        cell.setCellValue(Utility.setCommaFormat(String.valueOf(thisMonth-lastMonth)));
        //ATM自動化機器.交易金額.本年累計============================================
        row = sheet.getRow(5);
        cell = row.getCell( (short) 3);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue("新台幣 "+Unit);
        cell = row.getCell( (short) 4);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( ((bean0.getValue("yearamt") == null)?0:bean0.getValue("yearamt")).toString()) );
        cell = row.getCell( (short) 5);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( ((bean1.getValue("yearamt") == null)?0:bean1.getValue("yearamt")).toString()) );
        cell = row.getCell( (short) 6);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        thisMonth = Integer.parseInt(((bean0.getValue("yearamt") == null)?0:bean0.getValue("yearamt")).toString());
        lastMonth = Integer.parseInt(((bean1.getValue("yearamt") == null)?0:bean1.getValue("yearamt")).toString());
        cell.setCellValue(Utility.setCommaFormat(String.valueOf(thisMonth-lastMonth)));
        //ATM自動化機器.交易次數.本月合計============================================
        row = sheet.getRow(6);
        cell = row.getCell( (short) 4);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( ((bean0.getValue("mtrancnt") == null)?0:bean0.getValue("mtrancnt")).toString()) );
        cell = row.getCell( (short) 5);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( ((bean1.getValue("mtrancnt") == null)?0:bean1.getValue("mtrancnt")).toString()) );
        cell = row.getCell( (short) 6);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        thisMonth = Integer.parseInt(((bean0.getValue("mtrancnt") == null)?0:bean0.getValue("mtrancnt")).toString());
        lastMonth = Integer.parseInt(((bean1.getValue("mtrancnt") == null)?0:bean1.getValue("mtrancnt")).toString());
        cell.setCellValue(Utility.setCommaFormat(String.valueOf(thisMonth-lastMonth)));
        //ATM自動化機器.交易次數.本年累計============================================
        row = sheet.getRow(7);
        cell = row.getCell( (short) 4);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( ((bean0.getValue("ytrancnt") == null)?0:bean0.getValue("ytrancnt")).toString()) );
        cell = row.getCell( (short) 5);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( ((bean1.getValue("ytrancnt") == null)?0:bean1.getValue("ytrancnt")).toString()) );
        cell = row.getCell( (short) 6);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        thisMonth = Integer.parseInt(((bean0.getValue("ytrancnt") == null)?0:bean0.getValue("ytrancnt")).toString());
        lastMonth = Integer.parseInt(((bean1.getValue("ytrancnt") == null)?0:bean1.getValue("ytrancnt")).toString());
        cell.setCellValue(Utility.setCommaFormat(String.valueOf(thisMonth-lastMonth)));
        //金融卡-發卡張數============================================
        row = sheet.getRow(8);
        cell = row.getCell( (short) 4);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( ((bean0.getValue("push_cnt") == null)?0:bean0.getValue("push_cnt")).toString()) );
        cell = row.getCell( (short) 5);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( ((bean1.getValue("push_cnt") == null)?0:bean1.getValue("push_cnt")).toString()) );
        cell = row.getCell( (short) 6);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        thisMonth = Integer.parseInt(((bean0.getValue("push_cnt") == null)?0:bean0.getValue("push_cnt")).toString());
        lastMonth = Integer.parseInt(((bean1.getValue("push_cnt") == null)?0:bean1.getValue("push_cnt")).toString());
        cell.setCellValue(Utility.setCommaFormat(String.valueOf(thisMonth-lastMonth)));
        //金融卡-停卡張數============================================
        row = sheet.getRow(9);
        cell = row.getCell( (short) 4);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( ((bean0.getValue("cancel_cnt") == null)?0:bean0.getValue("cancel_cnt")).toString()) );
        cell = row.getCell( (short) 5);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( ((bean1.getValue("cancel_cnt") == null)?0:bean1.getValue("cancel_cnt")).toString()) );
        cell = row.getCell( (short) 6);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        thisMonth = Integer.parseInt(((bean0.getValue("cancel_cnt") == null)?0:bean0.getValue("cancel_cnt")).toString());
        lastMonth = Integer.parseInt(((bean1.getValue("cancel_cnt") == null)?0:bean1.getValue("cancel_cnt")).toString());
        cell.setCellValue(Utility.setCommaFormat(String.valueOf(thisMonth-lastMonth)));
        //金融卡-流通張數============================================
        row = sheet.getRow(10);
        cell = row.getCell( (short) 4);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( ((bean0.getValue("use_cnt") == null)?0:bean0.getValue("use_cnt")).toString()) );
        cell = row.getCell( (short) 5);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( ((bean1.getValue("use_cnt") == null)?0:bean1.getValue("use_cnt")).toString()) );
        cell = row.getCell( (short) 6);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        thisMonth = Integer.parseInt(((bean0.getValue("use_cnt") == null)?0:bean0.getValue("use_cnt")).toString());
        lastMonth = Integer.parseInt(((bean1.getValue("use_cnt") == null)?0:bean1.getValue("use_cnt")).toString());
        cell.setCellValue(Utility.setCommaFormat(String.valueOf(thisMonth-lastMonth)));
        //金融卡-交易金額.本月合計============================================
        row = sheet.getRow(11);
        cell = row.getCell( (short) 3);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue("新台幣 "+Unit);
        cell = row.getCell( (short) 4);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( ((bean0.getValue("debitcard_monthamt") == null)?0:bean0.getValue("debitcard_monthamt")).toString()) );
        cell = row.getCell( (short) 5);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( ((bean1.getValue("debitcard_monthamt") == null)?0:bean1.getValue("debitcard_monthamt")).toString()) );
        cell = row.getCell( (short) 6);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        thisMonth = Integer.parseInt(((bean0.getValue("debitcard_monthamt") == null)?0:bean0.getValue("debitcard_monthamt")).toString());
        lastMonth = Integer.parseInt(((bean1.getValue("debitcard_monthamt") == null)?0:bean1.getValue("debitcard_monthamt")).toString());
        cell.setCellValue(Utility.setCommaFormat(String.valueOf(thisMonth-lastMonth)));
        //金融卡-交易金額.本年累計============================================
        row = sheet.getRow(12);
        cell = row.getCell( (short) 3);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue("新台幣 "+Unit);
        cell = row.getCell( (short) 4);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( ((bean0.getValue("debitcard_yearamt") == null)?0:bean0.getValue("debitcard_yearamt")).toString()) );
        cell = row.getCell( (short) 5);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( ((bean1.getValue("debitcard_yearamt") == null)?0:bean1.getValue("debitcard_yearamt")).toString()) );
        cell = row.getCell( (short) 6);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        thisMonth = Integer.parseInt(((bean0.getValue("debitcard_yearamt") == null)?0:bean0.getValue("debitcard_yearamt")).toString());
        lastMonth = Integer.parseInt(((bean1.getValue("debitcard_yearamt") == null)?0:bean1.getValue("debitcard_yearamt")).toString());
        cell.setCellValue(Utility.setCommaFormat(String.valueOf(thisMonth-lastMonth)));
        //金融卡-交易次數.本月合計============================================
        row = sheet.getRow(13);       
        cell = row.getCell( (short) 4);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( ((bean0.getValue("debitcard_mtrancnt") == null)?0:bean0.getValue("debitcard_mtrancnt")).toString()) );
        cell = row.getCell( (short) 5);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( ((bean1.getValue("debitcard_mtrancnt") == null)?0:bean1.getValue("debitcard_mtrancnt")).toString()) );
        cell = row.getCell( (short) 6);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        thisMonth = Integer.parseInt(((bean0.getValue("debitcard_mtrancnt") == null)?0:bean0.getValue("debitcard_mtrancnt")).toString());
        lastMonth = Integer.parseInt(((bean1.getValue("debitcard_mtrancnt") == null)?0:bean1.getValue("debitcard_mtrancnt")).toString());
        cell.setCellValue(Utility.setCommaFormat(String.valueOf(thisMonth-lastMonth)));
        //金融卡-交易次數.本年累計============================================
        row = sheet.getRow(14);        
        cell = row.getCell( (short) 4);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( ((bean0.getValue("debitcard_ytrancnt") == null)?0:bean0.getValue("debitcard_ytrancnt")).toString()) );
        cell = row.getCell( (short) 5);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility.setCommaFormat( ((bean1.getValue("debitcard_ytrancnt") == null)?0:bean1.getValue("debitcard_ytrancnt")).toString()) );
        cell = row.getCell( (short) 6);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        thisMonth = Integer.parseInt(((bean0.getValue("debitcard_ytrancnt") == null)?0:bean0.getValue("debitcard_ytrancnt")).toString());
        lastMonth = Integer.parseInt(((bean1.getValue("debitcard_ytrancnt") == null)?0:bean1.getValue("debitcard_ytrancnt")).toString());
        cell.setCellValue(Utility.setCommaFormat(String.valueOf(thisMonth-lastMonth)));
      } //end of else dbData.size() == 0

      FileOutputStream fout = null;
      if (bank_type.equals("6")) {
        fout = new FileOutputStream(reportDir +System.getProperty("file.separator") + "全體農會信用部自動化機器彙計.xls");
      }
      else {
        fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "全體漁會信用部自動化機器彙計.xls");
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

    System.out.println("FR025W_Excel End ...");
    return errMsg;
  }
}
