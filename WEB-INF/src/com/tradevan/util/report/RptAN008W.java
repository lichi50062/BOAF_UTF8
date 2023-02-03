/*
 * Created on 2006/10/20 by ABYSS Brenda
 * 某年底全體農漁會信用部逾期放款及存款、放款、淨值備抵呆帳分析表
 *
 * 2006/12/18 by Abyss Brenda
 * 1.新增其他縣市別
 * 2.金額四捨五入至整數位
 *
 * 2006/12/22 by Abyss Brenda
 * 1.新增可區分農漁會
 * 2.新增選擇金額單位
 *
 * 2006/12/27 by Abyss Brenda
 * 1.逾放比率為0時，不需計算於本張報表中
 * 
 * fixed 99.06.07 sql injection by 2808
 * 108.06.12 fix 調整因檔名造成無法轉ods/pdf問題 by 2295
 */

package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.util.*;
import java.sql.*;

import com.tradevan.util.Utility;
import com.tradevan.util.Utility_report;
import com.tradevan.util.dao.RdbCommonDao;


public class RptAN008W {
  public static String createRpt(String S_YEAR, String bankType ,String tmUnit) {
    boolean debug = false;
    if(debug) System.out.println("RptAN008W Start ...");

    //金額單位
    String unitName = tmUnit.substring(0,tmUnit.indexOf(";"));
    String unit = tmUnit.substring(tmUnit.indexOf(";") + 1);

    // 取得當下年月資料轉換西元年到民國年
    Calendar c = Calendar.getInstance();
    //String nowYear = Integer.toString(c.get(Calendar.YEAR) - 1911);
    //String nowMonth = Integer.toString(c.get(Calendar.MONTH) + 1);
    //if(!S_YEAR.equals(nowYear)) nowMonth = "12";

    //2006-12-19修改只取得12月份資料
    int nowYear = c.get(Calendar.YEAR) - 1911;
    int qYear = Integer.parseInt(S_YEAR);
    String nowMonth = "12";

    String errMsg = "";
    Connection conn = null;
    PreparedStatement pst = null;
    ResultSet rs = null;
    StringBuffer sqlCmd = new StringBuffer () ;
	List paramList = new ArrayList () ;
	String u_year = "99" ;
	if(!"".equals(S_YEAR) && Integer.parseInt(S_YEAR) >99) {
		u_year = "100" ;
	}

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
      System.out.println("某年底全體農漁會信用部逾期放款及存款、放款、淨值備抵呆帳分析表.xls");

      finput = new FileInputStream(xlsDir +
                                   System.getProperty("file.separator") +
                                   "某年底全體農漁會信用部逾期放款及存款、放款、淨值備抵呆帳分析表.xls");

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

      //設定報表表頭資料============================================
      row = sheet.getRow(0);
      cell = row.getCell( (short) 0);
      cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
      //cell.setCellValue(S_YEAR + " 年底全體農漁會信用部逾期放款及存款、放款、淨值備抵呆帳分析表");
      //2006-12-21 變更
      String title = "全體農漁會信用部逾期放款及存款、放款、淨值備抵呆帳分析表";
      if(bankType.equals("6")){
        title = "全體農會信用部逾期放款及存款、放款、淨值備抵呆帳分析表";
      }else if(bankType.equals("7")){
        title = "全體漁會信用部逾期放款及存款、放款、淨值備抵呆帳分析表";
      }
      cell.setCellValue(S_YEAR + title);

      //設定年月及單位資料============================================
      row = sheet.getRow(1);
      cell = row.getCell( (short) 6);
      cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
      cell.setCellValue(S_YEAR +" 年 " + nowMonth + " 月底");
      cell = row.getCell( (short) 12);
      cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
      cell.setCellValue("單位：新台幣 " + unitName);

      if(qYear >= nowYear){
        row = sheet.getRow(3);
        cell = row.getCell( (short) 0);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue("查無資料");
        sheet.addMergedRegion(new Region((short)3, (short)0,(short)3, (short)13));//跨行(第幾行,開始欄位數,跨幾行,結束欄位數)
      }else{
        //取得所有縣市的單位數
        conn = (new RdbCommonDao("")).newConnection();
        if (debug) System.out.println("conn=" + conn);

        sqlCmd.append(
            "SELECT BN01.BANK_TYPE AS BANK_TYPE,A01.BANK_CODE AS BANK_CODE"
            + ",A01.ACC_CODE AS ACC_CODE,SUM(A01.AMT) AS AMOUNT "
            + "FROM A01,(select * from bn01 where m_year=?)BN01 "
            + "WHERE A01.BANK_CODE = BN01.BANK_NO "
            + "AND A01.ACC_CODE IN (?,?,?,"
            + "?,?,?,?,?) " //--會計科目
            + "AND A01.M_YEAR = ? " //--查詢年度
            + "AND A01.M_MONTH=? "
            + "AND A01.BANK_CODE !=? ");
        paramList.add(u_year);
        paramList.add("990000");
        paramList.add("220000");
        paramList.add("120000");
        paramList.add("120800");
        paramList.add("150300");
        paramList.add("310000");
        paramList.add("320000");
        paramList.add("300000");
        paramList.add( S_YEAR);
        paramList.add("12");
        paramList.add("8888888");
        if (bankType.equals("")) {
          sqlCmd.append( "AND BN01.BANK_TYPE IN (?,?) " );
          paramList.add("6") ;
          paramList.add("7") ;
        }
        else {
          sqlCmd.append( "AND BN01.BANK_TYPE = ? " );
          paramList.add(bankType) ;
        }
        sqlCmd.append( "GROUP BY BN01.BANK_TYPE,BANK_CODE,ACC_CODE" ); 
        
        if (debug) System.out.println("sqlCmd=" + sqlCmd);
        pst = conn.prepareStatement(sqlCmd.toString());
        setPreparedStatementParameter(pst,paramList) ;
        rs = pst.executeQuery();
        List dataList = new ArrayList(); //記錄每個單位的會計科目及金額
        HashMap map = new HashMap();
        String tmBankCode = "";
        String dbBankCode = "";
        while (rs.next()) {
          dbBankCode = Utility_report.getTrimString(rs.getString("BANK_CODE"));
          if (!tmBankCode.equals(dbBankCode)) {
            if (!tmBankCode.equals("")) {
              dataList.add(map);
              map = new HashMap();
            }
            tmBankCode = dbBankCode;
            map.put("BANK_TYPE",Utility_report.getTrimString(rs.getString("BANK_TYPE")));
            map.put("BANK_CODE", tmBankCode);
          }
          String accCode = rs.getString("ACC_CODE");
          String amountStr = rs.getString("AMOUNT");
          if (debug) System.out.println("dbBankCode=" + dbBankCode +
                                        ";accCode=" + accCode + ";amountStr=" +
                                        amountStr);
          map.put(accCode, Utility_report.getTrimString(amountStr, "0"));
        }

        if (map.size() > 0) {
          dataList.add(map);
        }
        else {
          row = sheet.getRow(3);
          cell = row.getCell( (short) 0);
          cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
          cell.setCellValue("查無資料");
          sheet.addMergedRegion(new Region( (short) 3, (short) 0, (short) 3,
                                           (short) 13)); //跨行(第幾行,開始欄位數,跨幾行,結束欄位數)
        }

        int t1[] = new int[6]; //家數
        long t2[] = new long[6]; //逾期放款金額
        long t3[] = new long[6]; //存款金額
        long t4[] = new long[6]; //放款總額
        long t5[] = new long[6]; //淨值金額
        long t6[] = new long[6]; //備抵呆帳金額
        int tmNo = 0;

        for (int i = 0; i < dataList.size(); i++) {
          map = (HashMap) dataList.get(i);
          long ac990000 = Long.parseLong(Utility_report.getTrimString( (String)
              map.get("990000"), "0"));
          long ac220000 = Long.parseLong(Utility_report.getTrimString( (String)
              map.get("220000"), "0"));
          long ac120000 = Long.parseLong(Utility_report.getTrimString( (String)
              map.get("120000"), "0"));
          long ac120800 = Long.parseLong(Utility_report.getTrimString( (String)
              map.get("120800"), "0"));
          long ac150300 = Long.parseLong(Utility_report.getTrimString( (String)
              map.get("150300"), "0"));
          long ac310000 = Long.parseLong(Utility_report.getTrimString( (String)
              map.get("310000"), "0"));
          long ac320000 = Long.parseLong(Utility_report.getTrimString( (String)
              map.get("320000"), "0"));
          long ac300000 = Long.parseLong(Utility_report.getTrimString( (String)
              map.get("300000"), "0"));

          //逾放比率＝逾放金額【990000】/  (放款【120000】+ 備抵呆帳-放款【120800】+備抵呆帳-催收款項【150300】)
          long amount = ac120000 + ac120800 + ac150300;
          double ratio = 0;

          if (amount == 0 || ac990000 == 0) {
            ratio = 0;
          }
          else {
            //ratio = Double.parseDouble(Utility_report.round(Long.toString(ac990000),new Long(amount).toString(), 2));
            ratio = Utility_report.round( (new Long(ac990000).doubleValue() /
                                           new Long(amount).doubleValue()) *
                                         100, 2);
          }

          /* 2007-01-02 逾放比率為零時，需計算
                     if(ratio <= 0){  //2006-12-27 新增逾放比率為零時，不需計算
            continue;
                     }
           */
          if (ratio < 5) {
            tmNo = 0;
          }
          else if (ratio < 10) {
            tmNo = 1;
          }
          else if (ratio < 15) {
            tmNo = 2;
          }
          else if (ratio < 25) {
            tmNo = 3;
          }
          else {
            tmNo = 4;
          }

          t1[tmNo]++;
          t2[tmNo] += ac990000;
          t3[tmNo] += ac220000;
          t4[tmNo] += ac120000 + ac120800 + ac150300;
          t6[tmNo] += ac120800 + ac150300;

          //合計
          t1[5]++;
          t2[5] += ac990000;
          t3[5] += ac220000;
          t4[5] += ac120000 + ac120800 + ac150300;
          t6[5] += ac120800 + ac150300;

          //BANK_TYPE  6==>農會,7==>漁會
          if ( ( (String) map.get("BANK_TYPE")).equals("6")) {
            t5[tmNo] += ac310000 + ac320000;
            t5[5] += ac310000 + ac320000;
          }
          else {
            t5[tmNo] += ac300000;
            t5[5] += ac300000;
          }
        }
        if (map.size() > 0) {
          for (int i = 0; i < 6; i++) {
            row = sheet.getRow(i + 3);
            cell = row.getCell( (short) 2);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
            cell.setCellValue(new Integer( (int) t1[i]).toString());
            cell = row.getCell( (short) 4);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
            cell.setCellValue(Utility.setCommaFormat(Utility_report.round(Long.
                toString(t2[i]), unit, 0)));
            cell = row.getCell( (short) 6);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
            cell.setCellValue(Utility.setCommaFormat(Utility_report.round(Long.
                toString(t3[i]), unit, 0)));
            cell = row.getCell( (short) 8);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
            cell.setCellValue(Utility.setCommaFormat(Utility_report.round(Long.
                toString(t4[i]), unit, 0)));
            cell = row.getCell( (short) 10);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
            cell.setCellValue(Utility.setCommaFormat(Utility_report.round(Long.
                toString(t5[i]), unit, 0)));
            cell = row.getCell( (short) 12);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
            cell.setCellValue(Utility.setCommaFormat(Utility_report.round(Long.
                toString(t6[i]), unit, 0)));
          }
        }
      }

      FileOutputStream fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "某年底全體農漁會信用部逾期放款及存款_放款_淨值備抵呆帳分析表.xls");
      HSSFFooter footer = sheet.getFooter();
      footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
      footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));

      wb.write(fout);
      //儲存
      fout.close();
    }catch (Exception e) {
      System.out.println("//RptAN008W createRpt() Have Error.....");
      e.printStackTrace();
      System.out.println(e.toString());
      System.out.println("//-------------------------------------");
    }finally {
      try {
          if(rs != null){
              rs.close();
              rs = null;//104.10.06
           }
           if(pst != null){
              pst.close();
              pst = null;//104.10.06
           }
           if(!conn.isClosed()){//104.10.06    
              conn.close();
              conn = null;
           }
      }catch (Exception sqlEx) {
        conn = null;
      }
    }

    if(debug) System.out.println("RptAN008W End ...");
    return errMsg;
  }
  private static void setPreparedStatementParameter(PreparedStatement pst,List paramList) throws Exception{
		for(int i = 0 ;i< paramList.size() ;i++) {
			pst.setString(i+1,(String)paramList.get(i)) ;
		}
	}
}
