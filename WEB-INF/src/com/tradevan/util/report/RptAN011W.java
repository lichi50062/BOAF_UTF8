/*
 * Created on 2006/10/24 by ABYSS Brenda
 * 某年底全體農漁會信用部存款、放款、資產總額、淨值、本期損益及淨值與存款總額比率排名表
 *
 * 2006/12/18 by Abyss Brenda
 * 1.新增其他縣市別
 * 2.金額四捨五入至整數位
 *
 * 2006/12/22 by Abyss Brenda
 * 1.新增選擇金額單位
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

public class RptAN011W {
  private static boolean debug = true;
  public static String createRpt(String S_YEAR,String bankType ,String tmUnit) {
    if(debug) System.out.println("RptAN011W Start ...");

    //金額單位
    String unitName = tmUnit.substring(0,tmUnit.indexOf(";"));
    String unit = tmUnit.substring(tmUnit.indexOf(";") + 1);

    // 取得當下年月資料轉換西元年到民國年
    Calendar c = Calendar.getInstance();
    int nowYear = c.get(Calendar.YEAR) - 1911;
    int qYear = Integer.parseInt(S_YEAR);

    String errMsg = "";
    Connection conn = null;
	String u_year = "99" ;
	if(!"".equals(S_YEAR) && Integer.parseInt(S_YEAR) >99) {
		u_year = "100" ;
	}
	System.out.println("u_year=" + u_year) ;
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
      String fileName = "";
      if(bankType.equals("6")){
        fileName = "某年底全體農會信用部存款、放款、資產總額、淨值、本期損益及淨值與存款總額比率排名表.xls";
        finput = new FileInputStream(xlsDir +
                                     System.getProperty("file.separator") +
                                     "某年底全體農會信用部存款、放款、資產總額、淨值、本期損益及淨值與存款總額比率排名表.xls");
      }else{
        fileName = "某年底全體漁會信用部存款、放款、資產總額、淨值、本期損益及淨值與存款總額比率排名表.xls";
        finput = new FileInputStream(xlsDir +
                                     System.getProperty("file.separator") +
                                     "某年底全體漁會信用部存款、放款、資產總額、淨值、本期損益及淨值與存款總額比率排名表.xls");
      }
      System.out.println(fileName);

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
      if(bankType.equals("6")){
        cell.setCellValue(S_YEAR + " 年底全體農會信用部存款、放款、資產總額、淨值、本期損益及淨值與存款總額比率排名表");
      }else{
        cell.setCellValue(S_YEAR + " 年底全體漁會信用部存款、放款、資產總額、淨值、本期損益及淨值與存款總額比率排名表");
      }

      //設定年月及單位資料============================================
      row = sheet.getRow(1);
      cell = row.getCell( (short) 0);
      cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
      cell.setCellValue("單位：新台幣 " + unitName + " ，%");

      //判斷是否有資料
      boolean isData = getIsData(S_YEAR, bankType);

      if(qYear >= nowYear || isData == false){
        row = sheet.getRow(3);
        cell = row.getCell( (short) 1);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue("查無資料");
        sheet.addMergedRegion(new Region((short)3, (short)1,(short)3, (short)12));//跨行(第幾行,開始欄位數,跨幾行,結束欄位數)
      }else{
        conn = (new RdbCommonDao("")).newConnection();
        if (debug) System.out.println("conn=" + conn);

        String accCode = ""; //會計科目
        //取得存款總額排名
        accCode = "'220000'";
        List listA = getAllList(S_YEAR, bankType, accCode);

        //取得放款總額排名
        if (bankType.equals("6")) {
          accCode = "'120000','120800','150300'";
        }else {
          //accCode = "'120000','20800','150300'";
          //2006-12-21
          accCode = "'120000','120800','150300'";
        }
        List listB = getAllList(S_YEAR, bankType, accCode);

        //取得資產總額排名
        //accCode = "'190000'";
        //2006-12-21
        if (bankType.equals("6")) {
          accCode = "'190000'";
        }else {
          accCode = "'100000'";
        }

        List listC = getAllList(S_YEAR, bankType, accCode);

        //取得淨值排名
        if (bankType.equals("6")) {
          accCode = "'310000','320000'";
        }else {
          accCode = "'300000'";
        }
        List listD = getAllList(S_YEAR, bankType, accCode);

        //取得本期損益排名
        accCode = "'320300'";
        List listE = getAllList(S_YEAR, bankType, accCode);

        //依淨值與存款比率取得名次
        List ratioList = null;
        if (bankType.equals("6")) {
          ratioList = getRatioListA(S_YEAR);
        }else if (bankType.equals("7")) {
          ratioList = getRatioListB(S_YEAR);
        }

        //取得所有機構簡稱
        HashMap bankMap = getBank(bankType,u_year);

        //設定儲存格資料======================
        HSSFRow row2 = sheet.getRow(3); //宣告一列

        for (int i = 0; i < listA.size(); i++) {
          //row = sheet.getRow(i + 3);
          if (sheet.getRow(i + 3) == null) {
            row = sheet.createRow(i + 3);
          }else {
            row = sheet.getRow(i + 3);
          }

          //取得存款總額排名
          HashMap map = (HashMap) listA.get(i);
          if (row.getCell( (short) 1) == null) {
            cell = row.createCell( (short) 1);
            cell.setCellStyle( (row2.getCell( (short) 1)).getCellStyle());
          }else {
            cell = row.getCell( (short) 1);
          }
          cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
          cell.setCellValue( (String)bankMap.get((String) map.get("BANK_NAME")));
          if (row.getCell( (short) 2) == null) {
            cell = row.createCell( (short) 2);
            cell.setCellStyle( (row2.getCell( (short) 2)).getCellStyle());
          }else {
            cell = row.getCell( (short) 2);
          }
          cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
          cell.setCellValue(Utility.setCommaFormat(Utility_report.round( (String) map.get("AMOUNT"), unit, 0)));

          //取得放款總額排名
          map = (HashMap) listB.get(i);
          if (row.getCell( (short) 3) == null) {
            cell = row.createCell( (short) 3);
            cell.setCellStyle( (row2.getCell( (short) 3)).getCellStyle());
          }else {
            cell = row.getCell( (short) 3);
          }
          cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
          cell.setCellValue( (String)bankMap.get((String) map.get("BANK_NAME")));
          if (row.getCell( (short) 4) == null) {
            cell = row.createCell( (short) 4);
            cell.setCellStyle( (row2.getCell( (short) 4)).getCellStyle());
          }else {
            cell = row.getCell( (short) 4);
          }
          cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
          cell.setCellValue(Utility.setCommaFormat(Utility_report.round( (String) map.get("AMOUNT"), unit, 0)));

          //取得資產總額排名
          map = (HashMap) listC.get(i);
          if (row.getCell( (short) 5) == null) {
            cell = row.createCell( (short) 5);
            cell.setCellStyle( (row2.getCell( (short) 5)).getCellStyle());
          }else {
            cell = row.getCell( (short) 5);
          }
          cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
          cell.setCellValue( (String)bankMap.get((String) map.get("BANK_NAME")));
          if (row.getCell( (short) 6) == null) {
            cell = row.createCell( (short) 6);
            cell.setCellStyle( (row2.getCell( (short) 6)).getCellStyle());
          }else {
            cell = row.getCell( (short) 6);
          }
          cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
          cell.setCellValue(Utility.setCommaFormat(Utility_report.round( (String) map.get("AMOUNT"), unit, 0)));

          //取得淨值排名
          map = (HashMap) listD.get(i);
          if (row.getCell( (short) 7) == null) {
            cell = row.createCell( (short) 7);
            cell.setCellStyle( (row2.getCell( (short) 7)).getCellStyle());
          }else {
            cell = row.getCell( (short) 7);
          }
          cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
          cell.setCellValue( (String)bankMap.get((String) map.get("BANK_NAME")));
          if (row.getCell( (short) 8) == null) {
            cell = row.createCell( (short) 8);
            cell.setCellStyle( (row2.getCell( (short) 8)).getCellStyle());
          }else {
            cell = row.getCell( (short) 8);
          }
          cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
          cell.setCellValue(Utility.setCommaFormat(Utility_report.round( (String) map.get("AMOUNT"), unit, 0)));

          //取得本期損益排名
          map = (HashMap) listE.get(i);
          if (row.getCell( (short) 9) == null) {
            cell = row.createCell( (short) 9);
            cell.setCellStyle( (row2.getCell( (short) 9)).getCellStyle());
          }else {
            cell = row.getCell( (short) 9);
          }
          cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
          cell.setCellValue( (String)bankMap.get((String) map.get("BANK_NAME")));
          if (row.getCell( (short) 10) == null) {
            cell = row.createCell( (short) 10);
            cell.setCellStyle( (row2.getCell( (short) 10)).getCellStyle());
          }else {
            cell = row.getCell( (short) 10);
          }
          cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
          cell.setCellValue(Utility.setCommaFormat(Utility_report.round( (String) map.get("AMOUNT"), unit, 0)));

          //依淨值與存款比率取得名次
          if (ratioList.size() > 0) { //程式有誤，需修改
            map = (HashMap) ratioList.get(i);
            if (row.getCell( (short) 11) == null) {
              cell = row.createCell( (short) 11);
              cell.setCellStyle( (row2.getCell( (short) 11)).getCellStyle());
            }else {
              cell = row.getCell( (short) 11);
            }
            cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
            cell.setCellValue( (String)bankMap.get((String) map.get("BANK_NAME")));
            if (row.getCell( (short) 12) == null) {
              cell = row.createCell( (short) 12);
              cell.setCellStyle( (row2.getCell( (short) 12)).getCellStyle());
            }else {
              cell = row.getCell( (short) 12);
            }
            cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
            cell.setCellValue((String) map.get("RATIO"));
          }
        }
      }

      FileOutputStream fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + fileName);
      HSSFFooter footer = sheet.getFooter();
      footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
      footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));

      wb.write(fout);
      //儲存
      fout.close();
    }catch (Exception e) {
      System.out.println("//RptAN011W createRpt() Have Error.....");
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

    if(debug) System.out.println("RptAN011W End ...");
    return errMsg;
  }

  /**
   * 依會計科目取得名次
   * @param S_YEAR String 查詢年度
   * @param bankType String 金融機構類別
   * @param accCode 會計科目 EX:'220000' OR '220000','300000'
   * @return List
   */
  public static List getAllList(String S_YEAR, String bankType,String accCode) {
    Connection conn = null;
    PreparedStatement pst = null;
    ResultSet rs = null;
    StringBuffer sqlCmd = new StringBuffer () ;
	List paramList = new ArrayList () ;
	String u_year = "99" ;
	if(!"".equals(S_YEAR) && Integer.parseInt(S_YEAR) >99) {
		u_year = "100" ;
	}
    List list = new ArrayList();
    try {
      conn = (new RdbCommonDao("")).newConnection();
      long tmAmount = 0;
      //取得存款總額排名
      sqlCmd.append("SELECT BN01.BANK_NAME,NVL(SUM(A01.AMT),0) AS AMOUNT "
          + "FROM (select * from bn01 where m_year=?)BN01 FULL OUTER JOIN A01 "
          + "ON (BN01.BANK_NO=A01.BANK_CODE "
          + "AND A01.ACC_CODE IN ("+accCode+") "         //--會計科目
          + "AND A01.M_YEAR = ? "             //--查詢年度
          + "AND A01.M_MONTH=? ) "
          + "WHERE BN01.BANK_TYPE = ? "     //--查詢農會或漁會
          + "AND BN01.BANK_NO != ?"
          + "GROUP BY BN01.BANK_NAME "
          + "ORDER BY AMOUNT DESC");
      paramList.add(u_year) ;
      paramList.add(S_YEAR) ;
      paramList.add("12") ;
      paramList.add(bankType) ;
      paramList.add("8888888") ;
      
      if(debug) System.out.println("sqlCmd=" + sqlCmd);
      pst = conn.prepareStatement(sqlCmd.toString());
      setPreparedStatementParameter(pst,paramList) ;
      rs = pst.executeQuery();
      HashMap map = null;
      while (rs.next()) {
        map = new HashMap();
        map.put("BANK_NAME",Utility_report.getTrimString(rs.getString("BANK_NAME")));
        map.put("AMOUNT", Utility_report.getTrimString(rs.getString("AMOUNT")));
        tmAmount += rs.getLong("AMOUNT");
        list.add(map);
      }
      //合計
      map = new HashMap();
      map.put("BANK_NAME", "");
      map.put("AMOUNT", Long.toString(tmAmount));
      list.add(map);
    }catch (Exception e) {
      System.out.println("//RptAN011W getAllList() Have Error.....");
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
    return list;
  }

  /**
   * 依淨值與存款比率取得名次 -- 農會
   * @param S_YEAR String 查詢年度
   * @param bankType String 金融機構類別
   * @return List
   */
  public static List getRatioListA(String S_YEAR) {
    Connection conn = null;
    PreparedStatement pst = null;
    ResultSet rs = null;
    StringBuffer sqlCmd = new StringBuffer () ;
	List paramList = new ArrayList () ;
	String u_year = "99" ;
	if(!"".equals(S_YEAR) && Integer.parseInt(S_YEAR) >99) {
		u_year = "100" ;
	}
    List list = new ArrayList();
    try {
      conn = (new RdbCommonDao("")).newConnection();
      long amountA = 0; //淨值
      long amountB = 0; //存款
      String ratio = "";  //比率

      sqlCmd.append("SELECT BN01.BANK_NAME,"
            + "(NVL((SELECT SUM(AMT) FROM A01 WHERE BANK_CODE=BN01.BANK_NO "
            + "AND ACC_CODE='310000' AND M_YEAR = ? AND M_MONTH= ? GROUP BY ACC_CODE),0) "
            + "+ NVL((SELECT SUM(AMT) FROM A01 WHERE BANK_CODE=BN01.BANK_NO "
            + "AND ACC_CODE='320000' AND M_YEAR = ? AND M_MONTH= ? GROUP BY ACC_CODE),0)) AS AMOUNT_A,"
            + "NVL((SELECT SUM(AMT) FROM A01 WHERE BANK_CODE=BN01.BANK_NO "
            + "AND ACC_CODE='220000' AND M_YEAR = ? AND M_MONTH= ?  GROUP BY ACC_CODE),0) AS AMOUNT_B "
            + "FROM (select * from bn01 where m_year=?)BN01 "
            + "WHERE BN01.BANK_TYPE = ? "
            + "AND BN01.BANK_NO != ? "
            + "ORDER BY AMOUNT_A DESC,BANK_NO");
      paramList.add(S_YEAR) ;
      paramList.add("12") ;
      paramList.add(S_YEAR) ;
      paramList.add("12") ;
      paramList.add(S_YEAR) ;
      paramList.add("12") ;
      paramList.add(u_year) ;
      paramList.add("6") ;
      paramList.add("8888888") ;
      if(debug) System.out.println("sqlCmda=" + sqlCmd);
      pst = conn.prepareStatement(sqlCmd.toString());
      setPreparedStatementParameter(pst,paramList) ;
      rs = pst.executeQuery();

      HashMap map = null;
      List dataList = new ArrayList();

      int num=0;
      while (rs.next()) {
        String dbBankName = Utility_report.getTrimString(rs.getString("BANK_NAME"));
        long dbAmountA = Utility_report.getTrimLong(rs.getString("AMOUNT_A"));
        long dbAmountB = Utility_report.getTrimLong(rs.getString("AMOUNT_B"));
        String dbRatio = Utility_report.round(
                                     Long.toString(dbAmountA * 100),
                                     Long.toString(dbAmountB), 2);
        amountA += dbAmountA;
        amountB += dbAmountB;
        //data[num] = dbRatio + "_" + dbBankName;
        dataList.add(dbRatio + "_" + dbBankName);
      }

      num = dataList.size();
      String data[] = new String[num];
      for(int i=0;i<num;i++){
        data[i] = (String)dataList.get(i);
      }

      //排序
      Arrays.sort(data);
      for(int i=data.length;i > 0;i--){
        String dbBankName = data[i-1].substring((data[i-1].indexOf("_")+1));
        String dbRatio = data[i-1].substring(0,data[i-1].indexOf("_"));
        map = new HashMap();
        map.put("BANK_NAME", dbBankName);
        map.put("RATIO", dbRatio);
        list.add(map);
      }

      //全體淨值與存款之比率
      ratio = Utility_report.round(Long.toString(amountA*100),Long.toString(amountB),2);
      map = new HashMap();
      map.put("BANK_NAME", "");
      map.put("RATIO", ratio);
      list.add(map);
    }catch (Exception e) {
      System.out.println("//RptAN011W getRatioList() Have Error.....");
      e.printStackTrace();
      System.out.println(e.toString());
      System.out.println("//-------------------------------------");
    }finally {
    	try{
    		if(pst!=null){
    		    pst.close();
    			pst = null ;
    		}
    		if(rs!=null) {
    			rs.close() ;
    			rs = null;
    		}
    		
    		if(!conn.isClosed()){//104.10.06        
    			//conn.commit() ;
    			conn.close();
    			conn = null;
    		}
    	}catch(Exception e) {
    		System.out.println("//RptAN011W getRatioList() Have Error.....");
    	    e.printStackTrace();
    	    System.out.println(e.toString());
    	    System.out.println("//-------------------------------------");
    	}
    }
    return list;
  }


  /**
   * 依淨值與存款比率取得名次 --漁會
   * @param S_YEAR String 查詢年度
   * @param bankType String 金融機構類別
   * @return List
   */
  public static List getRatioListB(String S_YEAR) {
    Connection conn = null;
    PreparedStatement pst = null;
    ResultSet rs = null;
    StringBuffer sqlCmd = new StringBuffer () ;
	List paramList = new ArrayList () ;
	String u_year = "99" ;
	if(!"".equals(S_YEAR) && Integer.parseInt(S_YEAR) >99) {
		u_year = "100" ;
	}
    List list = new ArrayList();
    try {
      conn = (new RdbCommonDao("")).newConnection();
      long amountA = 0; //淨值
      long amountB = 0; //存款
      String ratio = "";  //比率

      sqlCmd.append("SELECT BN01.BANK_NAME,"
            + "NVL((SELECT SUM(AMT) FROM A01 WHERE BANK_CODE=BN01.BANK_NO "
            + "AND ACC_CODE='300000' AND M_YEAR = ?  AND M_MONTH=? GROUP BY ACC_CODE),0) AS AMOUNT_A,"
            + "NVL((SELECT SUM(AMT) FROM A01 WHERE BANK_CODE=BN01.BANK_NO "
            + "AND ACC_CODE='220000' AND M_YEAR = ? AND M_MONTH=? GROUP BY ACC_CODE),0) AS AMOUNT_B,"
            + "NVL(((SELECT SUM(AMT) FROM A01 WHERE BANK_CODE=BN01.BANK_NO "
            + "AND ACC_CODE='300000' AND M_YEAR = ? AND M_MONTH=?  GROUP BY ACC_CODE) / "
            + "(SELECT SUM(AMT) FROM A01 WHERE BANK_CODE=BN01.BANK_NO "
            + "AND ACC_CODE='220000' AND M_YEAR =? AND M_MONTH=? GROUP BY ACC_CODE)"
            + "),0)*100 AS RATIO "
            + "FROM (select * from bn01 where m_year=?)BN01 "
            + "WHERE BN01.BANK_TYPE =?"
            + "AND BN01.BANK_NO != ? "
            + "ORDER BY RATIO DESC,BANK_NO");
      paramList.add(S_YEAR) ;
      paramList.add("12") ;
      paramList.add(S_YEAR) ;
      paramList.add("12") ;
      paramList.add(S_YEAR) ;
      paramList.add("12") ;
      paramList.add(S_YEAR) ;
      paramList.add("12") ;
      paramList.add(u_year) ;
      paramList.add("7") ;
      paramList.add("8888888") ;
      if(debug) System.out.println("sqlCmda=" + sqlCmd);
      pst = conn.prepareStatement(sqlCmd.toString());
      setPreparedStatementParameter(pst,paramList) ;
      rs = pst.executeQuery();
      HashMap map = null;
      while (rs.next()) {
        map = new HashMap();
        map.put("BANK_NAME",Utility_report.getTrimString(rs.getString("BANK_NAME")));
        String tmRatio = Utility_report.getTrimString(rs.getString("RATIO"),"0");
        map.put("RATIO", Utility_report.round(tmRatio,2));
        amountA += rs.getLong("AMOUNT_A");
        amountB += rs.getLong("AMOUNT_B");
        list.add(map);
      }
      //全體淨值與存款之比率
      ratio = Utility_report.round(Long.toString(amountA*100),Long.toString(amountB),2);
      map = new HashMap();
      map.put("BANK_NAME", "");
      map.put("RATIO", ratio);
      list.add(map);
    }catch (Exception e) {
      System.out.println("//RptAN011W getRatioList() Have Error.....");
      e.printStackTrace();
      System.out.println(e.toString());
      System.out.println("//-------------------------------------");
    }finally {
      try {
    	  if(pst!=null){
    	      pst.close();
    		  pst = null ;
    	  }
    	  if(rs!=null) {
    		  rs.close() ;
    		  rs =  null;
    	  }
    	  if(conn!=null){
    		  //conn.commit() ;
    		  conn.close(); 
    		  conn = null;
    	  }
      }catch (Exception e) {
    	  System.out.println("//RptAN011W getRatioList() Have Error.....");
          e.printStackTrace();
          System.out.println(e.toString());
          System.out.println("//-------------------------------------");
      }
    }
    return list;
  }

  /**
   * 取得機構簡稱
   * @param S_YEAR String
   * @param bankType String
   * @return List
   */
  public static HashMap getBank(String bankType,String u_year) {
    Connection conn = null;
    PreparedStatement pst = null;
    ResultSet rs = null;
    StringBuffer sqlCmd = new StringBuffer () ;
	List paramList = new ArrayList () ;
	
    HashMap map = new HashMap();
    try {
      conn = (new RdbCommonDao("")).newConnection();
      sqlCmd.append("SELECT BANK_NO, BANK_NAME "
            + "FROM BN01 "
            + "WHERE BANK_TYPE = ? AND M_YEAR=? ");
      paramList.add(bankType) ;
      paramList.add(u_year) ;
      if(debug) System.out.println("sqlCmda=" + sqlCmd);
      pst = conn.prepareStatement(sqlCmd.toString());
      setPreparedStatementParameter(pst,paramList) ;
      rs = pst.executeQuery();
      String bankNo = "";
      String bankName = "";
      String oldBankName = "";
      while (rs.next()) {
        bankNo = rs.getString("BANK_NO");
        bankName = rs.getString("BANK_NAME");
        oldBankName = rs.getString("BANK_NAME");

        if(bankName.length() > 8 && bankName.indexOf("信用部") >= 0){
          bankName = bankName.substring(3);
        }
        if (bankName.indexOf("漁會信用部") >= 0) {
          bankName = bankName.substring(0, bankName.indexOf("漁會信用部"));
        }else if(bankName.indexOf("農會信用部") >= 0) {
          bankName = bankName.substring(0, bankName.indexOf("農會信用部"));
        }
        //System.out.println("oldBankName=" + oldBankName + " bankName=" +bankName);
        map.put(oldBankName,bankName);
      }
    }catch (Exception e) {
      System.out.println("//RptAN011W getBank() Have Error.....");
      e.printStackTrace();
      System.out.println(e.toString());
      System.out.println("//-------------------------------------");
    }finally {
      try {
    	  if(pst!=null){
    	      pst.close();
    		  pst = null ;
    	  }
    	  if(rs!=null) {
    		  rs.close() ;
    		  rs = null;
    	  }
    	  if(conn!=null){
    		  //conn.commit() ;
    		  conn.close(); 
    		  conn = null;
    	  }
      }catch (Exception e) {
    	  System.out.println("//RptAN011W getCity() Have Error.....");
          e.printStackTrace();
          System.out.println(e.toString());
          System.out.println("//-------------------------------------");
      }
    }
    return map;
  }

  /**
   * 判斷是否有資料
   * @param S_YEAR String
   * @param bankType String
   * @param cityType String
   * @return boolean
   */
  public static boolean getIsData(String S_YEAR, String bankType) {
    Connection conn = null;
    PreparedStatement pst = null;
    ResultSet rs = null;
    boolean isCheck = false; //判斷是否有資料
    StringBuffer sqlCmd = new StringBuffer () ;
	List paramList = new ArrayList () ;
	String u_year = "99" ;
	if(!"".equals(S_YEAR) && Integer.parseInt(S_YEAR) >99) {
		u_year = "100" ;
	}
    try {
      conn = (new RdbCommonDao("")).newConnection();
      sqlCmd.append("SELECT COUNT(*) "
          + "FROM A01,(select  * from bn01 where m_year= ? )BN01 "
          + "WHERE A01.BANK_CODE = BN01.BANK_NO "
          + "AND A01.M_YEAR = ? "  //--查詢年度
          + "AND A01.M_MONTH = ? "
          + "AND BN01.BANK_TYPE = ? " //--查詢農會或漁會
          + "AND BN01.BANK_NO != ? ");
      paramList.add(u_year) ;
      paramList.add(S_YEAR) ;
      paramList.add("12") ;
      paramList.add(bankType) ;
      paramList.add("8888888") ;
      //System.out.println("sqlCmd=" + sqlCmd);
      pst = conn.prepareStatement(sqlCmd.toString());
      setPreparedStatementParameter(pst,paramList) ;
      rs = pst.executeQuery();
      if (rs.next()) {
        int num = rs.getInt(1);
        if(num > 0){
          isCheck = true;
        }
      }
      pst.close() ;
      pst = null;
      rs.close() ;
      rs = null;
    }catch (Exception e) {
      System.out.println("//RptAN011W getIsData() Have Error.....");
      e.printStackTrace();
      System.out.println(e.toString());
      System.out.println("//-------------------------------------");
    }finally {
        try {
      	  if(pst!=null){
      	      pst.close();
      		  pst = null ;
      	  }
      	  if(rs!=null) {
      		  rs.close() ;
      		  rs = null;
      	  }
      	  if(conn!=null){
      		  //conn.commit() ;
      		  conn.close();
      		  conn = null;
      	  }
        }catch (Exception e) {
      	  System.out.println("//RptAN011W getIsData() Have Error.....");
            e.printStackTrace();
            System.out.println(e.toString());
            System.out.println("//-------------------------------------");
        }
      }
    return isCheck;
  }
  private static void setPreparedStatementParameter(PreparedStatement pst,List paramList) throws Exception{
		for(int i = 0 ;i< paramList.size() ;i++) {
			pst.setString(i+1,(String)paramList.get(i)) ;
		}
	}
}
