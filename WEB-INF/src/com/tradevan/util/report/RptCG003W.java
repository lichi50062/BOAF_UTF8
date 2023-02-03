/*
 * Created on 2006/11/27 by ABYSS Brenda
 * 稽核記錄項目明細報表
 * 2008.08.29 fix 修正報表title.亂碼問題 by 2295
 */

package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.util.*;
import java.sql.*;
import com.tradevan.util.Utility;
import com.tradevan.util.Utility_report;
import com.tradevan.util.dao.RdbCommonDao;

public class RptCG003W {
  private static boolean debug = true;
  public static String createRpt(String qYear, String qMonth,
                                 String bankType, String cityType,
                                 String tbank, String tbName,
                                 String reportGroup) {
    if(debug) System.out.println("RptCG003W Start ...");

    String tbNameEng = tbName.substring(0,tbName.indexOf("@"));
    String tbNameChi = Utility.ISOtoBig5(Utility.toBig5Convert(tbName).substring(Utility.toBig5Convert(tbName).indexOf("@")+1,Utility.toBig5Convert(tbName).length()));
    String oldQYear = qYear;
    qYear = Integer.toString(Integer.parseInt(qYear) + 1911);
    cityType = Utility_report.getTrimString(cityType);

    HashMap typeMap = new HashMap();
    typeMap.put("U","異動");
    typeMap.put("D","刪除");
    typeMap.put("L","下載");
    typeMap.put("X00","下載");

    if(debug) System.out.println("tbNameEng="+tbNameEng);
    if(debug) System.out.println("tbNameCh=" + tbNameChi);

    String errMsg = "";
    Connection conn = null;
    PreparedStatement pst = null;
    ResultSet rs = null;
    String sqlCmd = null;
    String tmColumn = null;

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
      finput = new FileInputStream(xlsDir +
                                   System.getProperty("file.separator") +
                                   "稽核記錄項目明細表.xls");

      //設定FileINputStream讀取Excel檔
      POIFSFileSystem fs = new POIFSFileSystem(finput);
      HSSFWorkbook wb = new HSSFWorkbook(fs);
      HSSFSheet sheet = wb.getSheetAt(0); //讀取第一個工作表，宣告其為sheet
      HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定

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
      cell.setCellValue(tbNameChi + " 功能稽核記錄項目明細表");

      row = sheet.getRow(1);
      cell = row.getCell( (short) 1);
      cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
      cell.setCellValue(oldQYear + " 年 " + qMonth + " 月");

      conn =(new RdbCommonDao("")).newConnection();
      if(debug) System.out.println("conn="+conn);

      //取得機構名稱及類別
      HashMap bankMap = new HashMap();
      sqlCmd = "SELECT BANK_NO,BANK_NAME FROM BN01";
      pst = conn.prepareStatement(sqlCmd);
      rs = pst.executeQuery();
      while (rs.next()) {
        bankMap.put(rs.getString("BANK_NO"), rs.getString("BANK_NAME"));
      }

      //查詢符合條件的資料
      if(reportGroup.equalsIgnoreCase("A")){
        tmColumn = "A.ACC_CODE,";
      }else if(tbNameEng.equalsIgnoreCase("B01_log")){
        tmColumn = "A.FUND_MASTER_NO || A.FUND_SUB_NO || A.FUND_NEXT_NO AS TMCOLUMN,";
      }else if(tbNameEng.equalsIgnoreCase("B02_log")){
        tmColumn = "A.RUN_MASTER_NO || A.RUN_SUB_NO || A.RUN_NEXT_NO AS TMCOLUMN,";
      }else if (tbNameEng.equalsIgnoreCase("B03_1_log") ||
                tbNameEng.equalsIgnoreCase("B03_2_log")) {
        tmColumn = "A.FUNS_MASTER_NO || A.FUNS_SUB_NO || A.FUNS_NEXT_NO AS TMCOLUMN,";
      }else if(tbNameEng.equalsIgnoreCase("B03_3_log")){
        tmColumn = "A.FUNO_MASTER_NO || A.FUNO_SUB_NO || A.FUNO_NEXT_NO AS TMCOLUMN,";
      }else if (tbNameEng.equalsIgnoreCase("B03_4_log") ||
                tbNameEng.equalsIgnoreCase("BN04_log") ||
                tbNameEng.equalsIgnoreCase("WLX01_log")) {
        tmColumn = "A.BANK_NO,";
      }else if (tbNameEng.equalsIgnoreCase("BANK_CMML_log")) {
        tmColumn = "A.M_NAME,";
      }else if(tbNameEng.equalsIgnoreCase("ExDefGoodF_log") ||
               tbNameEng.equalsIgnoreCase("ExDG_HistoryF_log")){
        tmColumn = "A.REPORTNO_SEQ,";
      }else if(tbNameEng.equalsIgnoreCase("ExDisTripF_log")){
        tmColumn = "A.DISP_ID || A.EXAM_ID AS TMCOLUMN,";
      }else if(tbNameEng.equalsIgnoreCase("ExHelpItemF_log")){
        tmColumn = "A.EXAM_ITEM,";
      }else if(tbNameEng.equalsIgnoreCase("ExReportF_log")){
        tmColumn = "A.REPORTNO,";
      }else if(tbNameEng.equalsIgnoreCase("ExRtDocF_log") ||
               tbNameEng.equalsIgnoreCase("ExWamingF_log")){
        tmColumn = "A.RT_DOCNO,";
      }else if(tbNameEng.equalsIgnoreCase("ExScheduleF_log")){
        tmColumn = "A.DISP_ID,";
      }else if(tbNameEng.equalsIgnoreCase("ExSnDocF_log")){
        tmColumn = "A.SN_DOCNO,";
      }else if(tbNameEng.equalsIgnoreCase("F01_log")){
        tmColumn = "A.BANK_CODE || A.DEP_TYPE || A.ACCT_TYPE AS TMCOLUMN,";
      }else if(tbNameEng.equalsIgnoreCase("M01_log")){
        tmColumn = "A.M_YEAR || A.M_MONTH || A.GUARANTEE_ITEM_NO AS TMCOLUMN,";
      }else if(tbNameEng.equalsIgnoreCase("M02_log") ||
               tbNameEng.equalsIgnoreCase("M05_log") ||
               tbNameEng.equalsIgnoreCase("M05_TOTACC_log")){
        tmColumn = "A.M_YEAR || A.M_MONTH || A.LOAN_UNIT_NO AS TMCOLUMN,";
      }else if(tbNameEng.equalsIgnoreCase("M03_log")){
        tmColumn = "A.M_YEAR || A.M_MONTH || A.DIV_NO AS TMCOLUMN,";
      }else if(tbNameEng.equalsIgnoreCase("M03_NOTE_log") ||
               tbNameEng.equalsIgnoreCase("M05_NOTE_log")){
        tmColumn = "A.M_YEAR || A.M_MONTH || A.NOTE_NO AS TMCOLUMN,";
      }else if(tbNameEng.equalsIgnoreCase("M04_log")){
        tmColumn = "A.M_YEAR || A.M_MONTH || A.LOAN_USE_NO AS TMCOLUMN,";
      }else if(tbNameEng.equalsIgnoreCase("M06_log") ||
               tbNameEng.equalsIgnoreCase("M07_log")){
        tmColumn = "A.M_YEAR || A.M_MONTH || A.AREA_NO AS TMCOLUMN,";
      }else if(tbNameEng.equalsIgnoreCase("M08_log")){
        tmColumn = "A.M_YEAR || A.M_MONTH || A.ID_NO AS TMCOLUMN,";
      }else if(tbNameEng.equalsIgnoreCase("MUSER_DATA_log") ||
               tbNameEng.equalsIgnoreCase("WTT07") ||
               tbNameEng.equalsIgnoreCase("WTT07_ELM_log")){
        tmColumn = "A.MUSER_ID,";
      }else if(tbNameEng.equalsIgnoreCase("WLX_APPLY_LOCK_log")){
        tmColumn = "A.M_YEAR || A.M_QUARTER || A.BANK_CODE || A.REPORT_NO AS TMCOLUMN,";
      }else if(tbNameEng.equalsIgnoreCase("WLX_Notify_log")){
        tmColumn = "A.SEQ_NO,";
      }else if(tbNameEng.equalsIgnoreCase("WLX_S_RATE_log") ||
               tbNameEng.equalsIgnoreCase("WLX08_S_GAME_APPLY_log") ||
               tbNameEng.equalsIgnoreCase("WLX09_S_WARNING_log")){
        tmColumn = "A.M_YEAR || A.M_QUARTER || A.BANK_NO AS TMCOLUMN,";
      }else if(tbNameEng.equalsIgnoreCase("WLX01_M_log") ||
               tbNameEng.equalsIgnoreCase("WLX05_ATM_SETUP_log") ||
               tbNameEng.equalsIgnoreCase("WLX06_M_OUTPUSH_log")){
        tmColumn = "A.BANK_NO || A.SEQ_NO AS TMCOLUMN,";
      }else if(tbNameEng.equalsIgnoreCase("WLS02_log")){
        tmColumn = "A.TBANK_NO || A.BANK_NO AS TMCOLUMN,";
      }else if(tbNameEng.equalsIgnoreCase("WLX02_M_log") ||
               tbNameEng.equalsIgnoreCase("WLX04_log")){
        tmColumn = "A.BANK_NO || A.SEQ_NO || A.ID AS TMCOLUMN,";
      }else if(tbNameEng.equalsIgnoreCase("WLX05_M_ATM_log") ||
               tbNameEng.equalsIgnoreCase("WLX07_M_IMPORTANT_log")){
        tmColumn = "A.M_YEAR || A.M_MONTH || A.BANK_NO AS TMCOLUMN,";
      }else if(tbNameEng.equalsIgnoreCase("WLX08_S_GAME_log")){
        tmColumn = "A.M_YEAR || A.M_QUARTER || A.BANK_NO || A.SEQ_NO AS TMCOLUMN,";
      }else if(tbNameEng.equalsIgnoreCase("WML01_log")){
        tmColumn = "A.M_YEAR || A.M_MONTH || A.BANK_CODE || A.REPORT_NO AS TMCOLUMN,";
      }else if(tbNameEng.equalsIgnoreCase("WML02_log")){
        tmColumn = "A.M_YEAR || A.M_MONTH || A.BANK_CODE || A.REPORT_NO || A.CANO AS TMCOLUMN,";
      }else if(tbNameEng.equalsIgnoreCase("WML03_log")){
        tmColumn = "A.M_YEAR || A.M_MONTH || A.BANK_CODE || A.REPORT_NO || A.SERIAL_NO AS TMCOLUMN,";
      }else if(tbNameEng.equalsIgnoreCase("WTT01_log")){
        tmColumn = "A.MUSER_ID || A.BANK_TYPE AS TMCOLUMN,";
      }else if(tbNameEng.equalsIgnoreCase("WZZ07_log")){
        tmColumn = "A.BANK_NO || A.PROGRAM_ID AS TMCOLUMN,";
      }else{
        tmColumn = "";
      }

      if(!cityType.equals("")){
        sqlCmd = "SELECT "+tmColumn+"A.UPDATE_TYPE_C,A.USER_ID_C,A.USER_NAME_C,"
            + "W.TBANK_NO,A.UPDATE_DATE_C "
            + "FROM " + tbNameEng + " A,WTT01 W ,WLX01 B,CD01 C "
            + "WHERE A.USER_ID_C = W.MUSER_ID "
            + "AND W.TBANK_NO = B.BANK_NO "
            + "AND B.HSIEN_ID = C.HSIEN_ID "
            + "AND C.HSIEN_ID = '" + cityType + "' ";
        if (!bankType.equals("")) {
          sqlCmd += "AND W.BANK_TYPE = '" + bankType + "' ";
        }
        if (!tbank.equals("")) {
          sqlCmd += "AND W.TBANK_NO = '" + tbank + "' ";
        }
        sqlCmd += "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') = '" + (qYear + qMonth) + "' "
            + "ORDER BY A.UPDATE_DATE_C,A.USER_ID_C,W.TBANK_NO";
      }else if (tbNameEng.equalsIgnoreCase("MUSER_DATA_log")) {
        sqlCmd = "SELECT " + tmColumn + "A.UUPDATE_TYPE_C AS UPDATE_TYPE_C,"
            + "A.USER_ID_C,A.USER_NAME_C,W.TBANK_NO,A.UPDATE_DATE_C "
            + "FROM " + tbNameEng + " A,WTT01 W "
            + "WHERE A.USER_ID_C = W.MUSER_ID ";
        if (!bankType.equals("")) {
          sqlCmd += "AND W.BANK_TYPE = '" + bankType + "' ";
        }
        if (!tbank.equals("")) {
          sqlCmd += "AND W.TBANK_NO = '" + tbank + "' ";
        }
        sqlCmd += "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') = '" + (qYear + qMonth) + "' "
            + "ORDER BY A.UPDATE_DATE_C,A.USER_ID_C,W.TBANK_NO";
      }else if (tbNameEng.equalsIgnoreCase("WTT07")) {
        sqlCmd = "SELECT " + tmColumn + "A.Result_P AS UPDATE_TYPE_C,"
            + "A.MUSER_ID AS USER_ID_C,W.MUSER_NAME AS USER_NAME_C,"
            + "W.TBANK_NO,A.INPUT_DATE AS UPDATE_DATE_C "
            + "FROM " + tbNameEng + " A,WTT01 W "
            + "WHERE A.MUSER_ID = W.MUSER_ID ";
        if (!bankType.equals("")) {
          sqlCmd += "AND W.BANK_TYPE = '" + bankType + "' ";
        }
        if (!tbank.equals("")) {
          sqlCmd += "AND W.TBANK_NO = '" + tbank + "' ";
        }
        sqlCmd += "AND TO_CHAR(A.INPUT_DATE,'yyyymm') = '" + (qYear + qMonth) + "' "
            + "ORDER BY A.INPUT_DATE,A.MUSER_ID,W.TBANK_NO";
      }else{
        sqlCmd = "SELECT "+tmColumn+"A.UPDATE_TYPE_C,A.USER_ID_C,A.USER_NAME_C,"
            + "W.TBANK_NO,A.UPDATE_DATE_C "
            + "FROM " + tbNameEng + " A,WTT01 W "
            + "WHERE A.USER_ID_C = W.MUSER_ID ";
        if (!bankType.equals("")) {
          sqlCmd += "AND W.BANK_TYPE = '" + bankType + "' ";
        }
        if (!tbank.equals("")) {
          sqlCmd += "AND W.TBANK_NO = '" + tbank + "' ";
        }
        sqlCmd += "AND TO_CHAR(A.UPDATE_DATE_C,'yyyymm') = '" + (qYear + qMonth) + "' "
            + "ORDER BY A.UPDATE_DATE_C,A.USER_ID_C,W.TBANK_NO";
      }

      if (debug) System.out.println("sqlCmd=" + sqlCmd);
      pst = conn.prepareStatement(sqlCmd);
      rs = pst.executeQuery();

      //備份
      int num = 0;
      HSSFRow row2 = sheet.getRow(3); //宣告一列
      while (rs.next()) {
        //=== 若超過原本的規劃的長度時,需建立row
        if (sheet.getRow(num + 3) == null) {
          row = sheet.createRow(num + 3);
        }else {
          row = sheet.getRow(num + 3);
        }
        if (row.getCell( (short) 0) == null) {
          cell = row.createCell( (short) 0);
          cell.setCellStyle( (row2.getCell( (short) 0)).getCellStyle());
        }else {
          cell = row.getCell( (short) 0);
        }
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility_report.getTrimString(rs.getString("USER_ID_C")));

        if (row.getCell( (short) 1) == null) {
          cell = row.createCell( (short) 1);
          cell.setCellStyle( (row2.getCell( (short) 1)).getCellStyle());
        }else {
          cell = row.getCell( (short) 1);
        }
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility_report.getTrimString(rs.getString("USER_NAME_C")));

        if (row.getCell( (short) 2) == null) {
          cell = row.createCell( (short) 2);
          cell.setCellStyle( (row2.getCell( (short) 2)).getCellStyle());
        }else {
          cell = row.getCell( (short) 2);
        }
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        String tbankNo = Utility_report.getTrimString(rs.getString("TBANK_NO"));
        cell.setCellValue(tbankNo);

        if (row.getCell( (short) 3) == null) {
          cell = row.createCell( (short) 3);
          cell.setCellStyle( (row2.getCell( (short) 3)).getCellStyle());
        }else {
          cell = row.getCell( (short) 3);
        }
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility_report.getTrimString(bankMap.get(tbankNo)));

        if (row.getCell( (short) 4) == null) {
          cell = row.createCell( (short) 4);
          cell.setCellStyle( (row2.getCell( (short) 4)).getCellStyle());
        }else {
          cell = row.getCell( (short) 4);
        }
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(Utility_report.getTrimString(rs.getString("UPDATE_DATE_C")));

        if (row.getCell( (short) 5) == null) {
          cell = row.createCell( (short) 5);
          cell.setCellStyle( (row2.getCell( (short) 5)).getCellStyle());
        }else {
          cell = row.getCell( (short) 5);
        }
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        //2006-12-18設定，若欄位項目中資料為ALL，則顯示報表下載
        String tmValue = Utility_report.getTrimString(rs.getString(1));
        if(tmValue.equalsIgnoreCase("ALL")){
          tmValue = "報表下載";
        }
        cell.setCellValue(tmValue);

        if (row.getCell( (short) 6) == null) {
          cell = row.createCell( (short) 6);
          cell.setCellStyle( (row2.getCell( (short) 6)).getCellStyle());
        }else {
          cell = row.getCell( (short) 6);
        }
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        String updateType = Utility_report.getTrimString(rs.getString("UPDATE_TYPE_C"));
        cell.setCellValue(Utility_report.getTrimString(typeMap.get(updateType)));

        num++;
      }




      FileOutputStream fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "稽核記錄項目明細表.xls");
      HSSFFooter footer = sheet.getFooter();
      footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
      footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
      wb.write(fout);
      //儲存
      fout.close();
    }catch (Exception e) {
      System.out.println("//RptCG003W createRpt() Have Error.....");
      e.printStackTrace();
      System.out.println(e.toString());
      System.out.println("//-------------------------------------");
    }finally {
      try {
        conn.close();
      }catch (Exception sqlEx) {
        conn = null;
      }
    }

    if(debug) System.out.println("RptCG003W End ...");
    return errMsg;
  }
}
