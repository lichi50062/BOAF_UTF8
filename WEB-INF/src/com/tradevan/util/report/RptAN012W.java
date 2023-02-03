/*
 * Created on 2006/10/31 by ABYSS Brenda
 * 某年底個別農漁會信用部業務經營概況
 *
 * 2006/12/18 by Abyss Brenda
 * 1.新增其他縣市別
 * 2.金額四捨五入至整數位
 *
 * 2006/12/22 by Abyss Brenda
 * 1.新增選擇金額單位
 *
 * 2006/12/26 by Abyss Brenda
 * 燕貞提出存放比率計算方式修改
 * 1.(X)修正後方式總額：有邏輯判斷的部份，需先將該縣市別的所有金額合計後（含農漁會），再做判斷。
 * 2.(Y)修正後存款總額：各信用部（含農漁會）依公式計算出Y值後（SUM(a:j)-i/2)，再做合計。
 * 3.修改會計科目
 *   3.1.負債總額(農)：(210000) + (220000) + (240000) + (250000) + (260000)
 *              (漁)：(200000)
 *   3.2.淨值總額(農)：(310000) + (310300) + (310800) + (310000-310300-310800) + (320100) + (320300)
 *              (漁)：(300000)
 *   3.3.資產總額（漁）：(100000)
 *  96.04.27 fix 統一農貸:農會維持不變，漁會公式由"SUM(120401+120402)"改為"SUM(120201+120202)"
 *               其他流動負債:公式由"SUM(190000-210000-210400-210700)"改為"SUM(210000-210400-210700)
 *               淨值總額:農會公式由"SUM(300000)"改為"SUM(310000+320000)，漁會維持不變
 * 				 存款平均餘額:"SUM(220000)/12"-- 1月累加到12月
 *               放款平均餘額:"SUM(120000+120800+150300)/12-- 1月累加到12月
 *               全年度負債科目平均餘額:1月累加到12月/12  by 2295
 *               
 *  99.06.08 fixed  sql injection by 2808
 *  102.06.03 fixed sql by 2968
 *  fixed 102.08.30 外包的connection拿掉 by 2968
 */

package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.text.DecimalFormat;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptAN012W {
  private static boolean debug = false;
  public static String createRpt(String S_YEAR, String bankType,
                                 String cityType, String tmUnit) {
    if (debug) System.out.println("RptAN012W Start ...");
    String hsien_id ="";
    String area_id ="";
    String area_name ="";
    String field_220000_12_last ="";
    String field_220000_12 ="";
    String field_220000 ="";
    String field_220100 ="";
    String field_220200 ="";
    String field_220300 ="";
    String field_220400 ="";
    String field_220500 ="";
    String field_220600 ="";
    String field_220700 ="";
    String field_220800 ="";
    String field_220900 ="";
    String field_221000 ="";
    String field_990420 ="";
    String field_990620 ="";
    String field_CREDIT_12_last ="";
    String field_CREDIT_12 ="";
    String field_CREDIT ="";
    String field_120101 ="";
    String field_120102 ="";
    String field_120200_cal ="";
    String field_120401_cal ="";
    String field_120501_cal ="";
    String field_120601 ="";
    String field_120602 ="";
    String field_120603_cal ="";
    String field_120604_cal ="";
    String field_120700 ="";
    String field_990410 ="";
    String field_110000 ="";
    String field_110100 ="";
    String field_110300_cal ="";
    String field_110400 ="";
    String field_111100_cal ="";      
    String field_110000_other ="";
    String field_120000 ="";
    String field_140000 ="";
    String field_150000_other ="";
    String field_190000_cal ="";
    String field_210000 ="";
    String field_210400_cal ="";
    String field_210700 ="";
    String field_210000_other ="";
    String field_220000_220900 ="";
    String field_240100 ="";
    String field_240300_cal ="";
    String field_240220_cal ="";
    String field_240210_cal ="";
    String field_240230_cal ="";
    String field_240240_cal ="";
    String field_250000_cal ="";
    String field_210000_cal ="";
    String field_310000 ="";
    String field_310300 ="";
    String field_310800_cal ="";
    String field_310000_other ="";
    String field_320100 ="";
    String field_320300 ="";
    String field_310000_cal ="";
    String field_rate1 ="";
    String field_rate2 ="";
    String field_rate3 ="";
    String field_rate4 ="";
    String field_991110 ="";
    String field_991120 ="";
    String field_DC_RATE ="";       
    String field_rate5 ="";
    String field_rate6 ="";
    String field_rate7 ="";
    String field_rate8 ="";
    String field_rate9 ="";
    String field_rate10 ="";
    String field_rate11 ="";
    String field_base_base_rate ="";
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
      if(debug) System.out.println("某年底個別農漁會信用部業務經營概況.xls");

      finput = new FileInputStream(xlsDir +
                                   System.getProperty("file.separator") +
                                   "某年底個別農漁會信用部業務經營概況.xls");

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
      String title = " 年底個別農漁會信用部業務經營概況";
      if(bankType.equals("6")){
        title = " 年底個別農會信用部業務經營概況";
      }else if(bankType.equals("7")){
        title = " 年底個別漁會信用部業務經營概況";
      }
      cell.setCellValue(S_YEAR + title);

      //設定年月及單位資料============================================
      row = sheet.getRow(1);
      cell = row.getCell( (short) 0);
      cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
      cell.setCellValue("單位：新台幣 " + unitName + " ，%");

      if(qYear >= nowYear){
        row = sheet.getRow(2);
        cell = row.getCell( (short) 3);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue("查無資料");
        sheet.addMergedRegion(new Region((short)2, (short)3,(short)3, (short)8));//跨行(第幾行,開始欄位數,跨幾行,結束欄位數)
      }else{
          sqlCmd.setLength(0) ;
          paramList.clear() ;
          sqlCmd.append("select hsien_id,area_id,area_name, ");
          sqlCmd.append("       round( a01.field_220000_12_last/12 /?,0) as  field_220000_12_last, ");  //上年度存款平均餘額 
          sqlCmd.append("       round( a01.field_220000_12/12 /?,0)      as  field_220000_12, ");  //本年度存款平均餘額
          sqlCmd.append("       round( a01.field_220000 /?,0)    as  field_220000, ");  //本年底存款餘額
          sqlCmd.append("       round( a01.field_220100 /?,0)    as  field_220100, ");  //支票存款       
          sqlCmd.append("       round( a01.field_220200 /?,0)    as  field_220200, ");  //保付支票
          sqlCmd.append("       round( a01.field_220300 /?,0)    as  field_220300, ");  //活期存款
          sqlCmd.append("       round( a01.field_220400 /?,0)    as  field_220400, ");  //活期儲蓄存款
          sqlCmd.append("       round( a01.field_220500 /?,0)    as  field_220500, ");  //員工活期儲蓄存款
          sqlCmd.append("       round( a01.field_220600 /?,0)    as  field_220600, ");  //定期存款
          sqlCmd.append("       round( a01.field_220700 /?,0)    as  field_220700, ");  //定期儲蓄存款
          sqlCmd.append("       round( a01.field_220800 /?,0)    as  field_220800, ");  //員工定期儲蓄存款
          sqlCmd.append("       round( a01.field_220900 /?,0)    as  field_220900, ");  //公庫存款
          sqlCmd.append("       round( a01.field_221000 /?,0)    as  field_221000, ");  //本會支票
          sqlCmd.append("       round( a01.field_990420 /?,0)    as  field_990420, ");  //贊助會員存款
          sqlCmd.append("       round( a01.field_990620 /?,0)    as  field_990620, ");  //非會員存款
          sqlCmd.append("       round( a01.field_CREDIT_12_last/12 /?,0)    as  field_CREDIT_12_last, ");  //上年度放款平均餘額
          sqlCmd.append("       round( a01.field_CREDIT_12/12 /?,0)    as  field_CREDIT_12, ");  //本年度放款平均餘額
          sqlCmd.append("       round( a01.field_CREDIT /?,0)    as  field_CREDIT, ");  //本年底放款餘額
          sqlCmd.append("       round( a01.field_120101 /?,0)    as  field_120101, ");  //一般放款 // 無擔保
          sqlCmd.append("       round( a01.field_120102 /?,0)    as  field_120102, ");  //一般放款 // 擔    保
          sqlCmd.append("       round( a01.field_120200_cal /?,0)    as  field_120200_cal, ");  //貼現及透支
          sqlCmd.append("       round( a01.fieldI_XB1/?,0) as  field_120401_cal, ");  // 統一農貸
          sqlCmd.append("       round(a01.fieldI_XC1/?,0) as  field_120501_cal, ");  // 專案放款
          sqlCmd.append("       round( a01.field_120601 /?,0)    as  field_120601, ");  //農建放款
          sqlCmd.append("       round( a01.field_120602 /?,0)    as  field_120602, ");  //農機放款
          sqlCmd.append("       round( a01.field_120603_cal /?,0)    as  field_120603_cal, ");  //購地放款
          sqlCmd.append("       round( a01.field_120604_cal /?,0)    as  field_120604_cal, ");  //農宅放款
          sqlCmd.append("       round( a01.field_120700 /?,0)    as  field_120700, ");  //內部融資
          sqlCmd.append("       round( a01.field_990410 /?,0)    as field_990410, ");  //贊助會員放款總額     
          sqlCmd.append("       round( a01.field_110000 /?,0)    as field_110000, ");  //流動資產
          sqlCmd.append("       round( a01.field_110100 /?,0)    as field_110100, ");  //庫存現金
          sqlCmd.append("       round( a01.field_110300_cal /?,0)    as field_110300_cal, ");  //存放行庫
          sqlCmd.append("       round( a01.field_110400 /?,0)    as field_110400, ");  //繳存存款準備金
          sqlCmd.append("       round( a01.field_111100_cal /?,0)    as field_111100_cal, ");  //應收利息       
          sqlCmd.append("       round( (a01.field_110000-a01.field_110100-a01.field_110300_cal-a01.field_110400) /?,0) as field_110000_other, ");//其他流動資產
          sqlCmd.append("       round( a01.field_120000 /?,0)    as field_120000, ");  //放款淨額(減備抵呆帳)
          sqlCmd.append("       round( a01.fieldI_XF2/?,0) as field_140000, ");//固定資產淨額
          sqlCmd.append("       round( a01.field_150000_other/?,0) as field_150000_other, ");//其他資產(含出資及往來)       
          sqlCmd.append("       round( a01.field_190000_cal/?,0) as field_190000_cal, ");//資產總額
          sqlCmd.append("       round( a01.field_210000/?,0) as field_210000, ");//流動負債
          sqlCmd.append("       round( a01.field_210400_cal/?,0) as field_210400_cal, ");//短期借款
          sqlCmd.append("       round( a01.field_210700/?,0) as field_210700, ");//應付利息
          sqlCmd.append("       round( (a01.field_210000-a01.field_210400_cal-a01.field_210700) /?,0) as field_210000_other, ");//其他流動負債
          sqlCmd.append("       round( a01.field_220000 /?,0)    as  field_220000_220900, ");  //存款(含公庫存款)
          sqlCmd.append("       round( a01.field_240100 /?,0)    as  field_240100, ");  //長期借款
          sqlCmd.append("       round( a01.field_240300_cal /?,0)    as  field_240300_cal, ");  //借入專案放款資金
          sqlCmd.append("       round( a01.field_240220_cal /?,0)    as  field_240220_cal, ");  //借入農機放款資金
          sqlCmd.append("       round( a01.field_240210_cal /?,0)    as  field_240210_cal, ");  //借入農建放款資金
          sqlCmd.append("       round( a01.field_240230_cal /?,0)    as  field_240230_cal, ");  //借入購地放款資金
          sqlCmd.append("       round( a01.field_240240_cal /?,0)    as  field_240240_cal, ");  //借入農宅放款資金
          sqlCmd.append("       round( (a01.field_250000+a01.field_260000) /?,0)    as  field_250000_cal, ");  //其他負債(含往來)
          sqlCmd.append("       round( a01.field_210000_cal /?,0)    as  field_210000_cal, ");  //負債總額  
          sqlCmd.append("       round( a01.field_310000 /?,0)    as  field_310000, ");  //事業基金及公積        
          sqlCmd.append("       round( a01.field_310300 /?,0)    as  field_310300, ");  //法定公積
          sqlCmd.append("       round( a01.field_310800_cal /?,0)    as  field_310800_cal, ");  //統一農貸公積
          sqlCmd.append("       round( (a01.field_310000-a01.field_310300-a01.field_310800_cal) /?,0)    as  field_310000_other, ");  //其他公積
          sqlCmd.append("       round( a01.field_320100 /?,0)    as  field_320100, ");  //累積盈餘
          sqlCmd.append("       round( a01.field_320300 /?,0)    as  field_320300, ");  //本期損益
          sqlCmd.append("       round( a01.field_310000_cal /?,0)    as  field_310000_cal, ");  //淨值總額
          sqlCmd.append("       decode(a01.field_220000,0,0,round(a01.field_520100 /  (a01.field_220000_12/12) *100 ,2))  as field_rate1, ");   //存款平均利率.12個月加總
          sqlCmd.append("       decode(a01.field_CREDIT_12,0,0,round(a01.field_420100 /  (a01.field_CREDIT_12/12) *100 ,2))  as field_rate2, ");   //放款平均利率.12個月加總
          sqlCmd.append("       decode((a01.field_debt1_12-a01.field_debt2_12),0,0,round(a01.field_520000 /  ((a01.field_debt1_12-a01.field_debt2_12)/12) *100 ,2))  as field_rate3, ");   //總資金平均成本率.12個月加總
          sqlCmd.append("       decode((a01.field_debt1_12-a01.field_debt2_12),0,0,round(a01.field_420000 /  ((a01.field_debt1_12-a01.field_debt2_12)/12) *100 ,2))  as field_rate4, ");   //總資金平均收益率.12個月加總
          sqlCmd.append("       round( a01.field_991110 /?,0)    as  field_991110, ");  //每一會員放款最高額
          sqlCmd.append("       round( a01.field_991120 /?,0)    as  field_991120, ");  //會員無擔保放款最高額
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          paramList.add(unit) ;
          sqlCmd.append("       decode(a01.fieldI_Y,0,0, ");                                  
          sqlCmd.append("              round((a01.fieldI_XA ");                               
          sqlCmd.append("                  + decode(sign(a01.fieldI_XB1 - a01.fieldI_XB2),-1,0,(a01.fieldI_XB1 - a01.fieldI_XB2)) ");   
          sqlCmd.append("                  + decode(sign(a01.fieldI_XC1 - a01.fieldI_XC2),-1,0,(a01.fieldI_XC1 - a01.fieldI_XC2)) ");   
          sqlCmd.append("                  + decode(sign(a01.fieldI_XD1 - a01.fieldI_XD2),-1,0,(a01.fieldI_XD1 - a01.fieldI_XD2)) ");   
          sqlCmd.append("                  + decode(sign(a01.fieldI_XE1 - a01.fieldI_XE2),-1,0,(a01.fieldI_XE1 - a01.fieldI_XE2)) ");   
          sqlCmd.append("                  - decode(sign(a01.fieldI_XF1 - a01.fieldI_XF3 -  a01.fieldI_XF2),-1,0,(a01.fieldI_XF1 - a01.fieldI_XF3 - a01.fieldI_XF2)) ");   
          sqlCmd.append("                    )/a01.fieldI_Y * 100,2))  as  field_DC_RATE, "); //存放比率            
          sqlCmd.append("       decode(a01.field_310000_cal,0,0,round(a01.field_120700 /  a01.field_310000_cal *100 ,2))  as field_rate5, ");  //內部融資占淨值比率     
          sqlCmd.append("       decode(a01.field_990420,0,0,round(a01.field_990410 /  a01.field_990420 *100 ,2))  as field_rate6, ");   //贊助會員放款占贊助會員存款比率        
          sqlCmd.append("       decode(a01.field_220000_220900,0,0,round(a01.field_110000 /  a01.field_220000_220900 *100 ,2))  as field_rate7, ");   //流動比率
          sqlCmd.append("       decode(a01.field_310000_cal,0,0,round(a01.fieldI_XF2 /  a01.field_310000_cal *100 ,2))  as field_rate8, ");   //固定資產淨額占淨值比率
          sqlCmd.append("       decode(a01.field_220000,0,0,round(a01.field_210200 /  a01.field_220000 *100 ,2))  as field_rate9, ");   //融通資金比率
          sqlCmd.append("       decode(a01.field_210000_cal,0,0,round(a01.field_310000_cal /  a01.field_210000_cal *100 ,2))  as field_rate10, ");   //淨值占負債總額比率
          sqlCmd.append("       decode(a01.field_220000,0,0,round(a01.field_310000_cal /  a01.field_220000 *100 ,2))  as field_rate11, ");  //淨值占存款總額比率
          sqlCmd.append("       decode(a01.base_base_rate,0,0,round(a01.base_base_rate / a01.num  ,2))  as field_base_base_rate ");   //本年底基本放款利率            
          sqlCmd.append("from ( ");   
          sqlCmd.append("     select a01.hsien_id,a01.area_id,a01.area_name, ");
          sqlCmd.append("     field_220100,field_220200,field_220300,field_220400,field_220500,field_220600,field_220700,field_220800,field_220900, ");
          sqlCmd.append("     field_221000,field_120101,field_120102,field_120200_cal,field_120601,field_120602,field_120603_cal,field_120604_cal, ");
          sqlCmd.append("     field_110000,field_110100,field_110300_cal,field_110400,field_111100_cal,field_120000,field_150000_other,field_190000_cal, ");
          sqlCmd.append("     field_210000,field_210400_cal,field_210700,field_220000,field_220000_220900,field_520100,field_420100, ");
          sqlCmd.append("     field_CREDIT,field_220000_12,field_220000_12_last,field_CREDIT_12,field_CREDIT_12_last,field_210200, ");
          sqlCmd.append("     field_310000_cal,field_210000_cal,field_120700,field_240100,field_240300_cal,field_240220_cal,field_240210_cal,field_240230_cal, ");
          sqlCmd.append("     field_240240_cal,field_250000,field_260000,field_310000,field_310300,field_310800_cal,field_320100,field_320300,field_520000, ");
          sqlCmd.append("     field_debt1_12,field_debt2_12,field_420000,fieldI_XA,fieldI_XB1,fieldI_XB2,fieldI_XC1,fieldI_XC2,fieldI_XD1,fieldI_XD2,fieldI_XE1,fieldI_XE2, ");
          sqlCmd.append("     fieldI_XF1,fieldI_XF3,fieldI_XF2,fieldI_Y,field_991110,field_991120,field_990410,field_990420,field_990620,base_base_rate,num ");
          sqlCmd.append("     from ( ");//本年底資料                   
          sqlCmd.append("             select cd01.hsien_id,cd01.area_id,cd01.area_name, ");
          sqlCmd.append("                    sum(decode(a01.acc_code,'220100',amt,0))  as  field_220100, ");
          sqlCmd.append("                    sum(decode(a01.acc_code,'220200',amt,0))  as  field_220200, ");
          sqlCmd.append("                    sum(decode(a01.acc_code,'220300',amt,0))  as  field_220300, ");
          sqlCmd.append("                    sum(decode(a01.acc_code,'220400',amt,0))  as  field_220400, ");
          sqlCmd.append("                    sum(decode(a01.acc_code,'220500',amt,0))  as  field_220500, ");
          sqlCmd.append("                    sum(decode(a01.acc_code,'220600',amt,0))  as  field_220600, ");
          sqlCmd.append("                    sum(decode(a01.acc_code,'220700',amt,0))  as  field_220700, ");
          sqlCmd.append("                    sum(decode(a01.acc_code,'220800',amt,0))  as  field_220800, ");
          sqlCmd.append("                    sum(decode(a01.acc_code,'220900',amt,0))  as  field_220900, ");
          sqlCmd.append("                    sum(decode(a01.acc_code,'221000',amt,0))  as  field_221000, ");         
          sqlCmd.append("                    sum(decode(a01.acc_code,'120101',amt,0))  as  field_120101, "); 
          sqlCmd.append("                    sum(decode(a01.acc_code,'120102',amt,0))  as  field_120102, ");   
          sqlCmd.append("                    decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'120200',amt,'120301',amt,'120302',amt,0), ");
          sqlCmd.append("                                                                '7',decode(a01.acc_code,'120300',amt,'120401',amt,'120402',amt,0),0)), ");
          sqlCmd.append("                                     '103',sum(decode(a01.acc_code,'120200',amt,'120301',amt,'120302',amt,0)),0)  as field_120200_cal, ");    
          sqlCmd.append("                    sum(decode(a01.acc_code,'120601',amt,0))  as  field_120601, ");    
          sqlCmd.append("                    sum(decode(a01.acc_code,'120602',amt,0))  as  field_120602, ");          
          sqlCmd.append("                    decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'120603',amt,0), ");
          sqlCmd.append("                                                                '7',decode(a01.acc_code,'120604',amt,0),0)), ");
          sqlCmd.append("                                     '103',sum(decode(a01.acc_code,'120603',amt,0)),0)  as field_120603_cal, ");   
          sqlCmd.append("                    decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'120604',amt,0), ");
          sqlCmd.append("                                                                '7',decode(a01.acc_code,'120603',amt,0),0)), ");
          sqlCmd.append("                                     '103',sum(decode(a01.acc_code,'120604',amt,0)),0)  as field_120604_cal, ");        
          sqlCmd.append("                    sum(decode(a01.acc_code,'110000',amt,0))  as  field_110000, ");                                 
          sqlCmd.append("                    sum(decode(a01.acc_code,'110100',amt,0))  as  field_110100, ");
          sqlCmd.append("                    decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'110300',amt,0), ");
          sqlCmd.append("                                                                '7',decode(a01.acc_code,'110200',amt,0),0)), ");
          sqlCmd.append("                                     '103',sum(decode(a01.acc_code,'110300',amt,0)),0)  as field_110300_cal, ");          
          sqlCmd.append("                    sum(decode(a01.acc_code,'110400',amt,0))  as  field_110400, ");
          sqlCmd.append("                    decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'111100',amt,0), ");
          sqlCmd.append("                                                                '7',decode(a01.acc_code,'111200',amt,0),0)), ");
          sqlCmd.append("                                     '103',sum(decode(a01.acc_code,'111100',amt,0)),0)  as field_111100_cal, ");  
          sqlCmd.append("                    sum(decode(a01.acc_code,'120000',amt,0))  as  field_120000, ");     
          sqlCmd.append("                    sum(decode(a01.acc_code,'150000',amt,'130000',amt,'160000',amt,0))  as  field_150000_other, ");          
          sqlCmd.append("                    sum(decode(bank_type,'6',decode(a01.acc_code,'190000',amt,0),'7',decode(a01.acc_code,'100000',amt,0),0)) as field_190000_cal, ");    
          sqlCmd.append("                    sum(decode(a01.acc_code,'210000',amt,0))  as  field_210000, ");     
          sqlCmd.append("                    decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'210400',amt,0), ");
          sqlCmd.append("                                                                '7',decode(a01.acc_code,'210200',amt,0),0)), ");
          sqlCmd.append("                                     '103',sum(decode(a01.acc_code,'210400',amt,0)),0)  as field_210400_cal, ");           
          sqlCmd.append("                    sum(decode(a01.acc_code,'210700',amt,0))  as  field_210700, ");                                            
          sqlCmd.append("                    sum(decode(a01.acc_code,'220000',amt,0))  as  field_220000, ");
          sqlCmd.append("                    sum(decode(a01.acc_code,'220000',amt,0)) - sum(decode(a01.acc_code,'220900',amt,0))   as  field_220000_220900, ");
          sqlCmd.append("                    sum(decode(a01.acc_code,'520100',amt,0))  as  field_520100, ");
          sqlCmd.append("                    sum(decode(bank_type,'6',decode(a01.acc_code,'420100',amt,'420170',amt,0),'7',decode(a01.acc_code,'420100',amt,0),0)) as field_420100, ");
          sqlCmd.append("                    sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) as  field_CREDIT, "); 
          sqlCmd.append("                    decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'210200',amt,'210300',amt,'210400',amt,0), ");
          sqlCmd.append("                                                                '7',decode(a01.acc_code,'210100',amt,'210300',amt,'210200',amt,0),0)), ");
          sqlCmd.append("                                     '103',sum(decode(a01.acc_code,'210200',amt,'210300',amt,'210400',amt,0)),0)  as field_210200, ");                                     
          sqlCmd.append("                    sum(decode(bank_type,'6',decode(a01.acc_code,'310000',amt,'320000',amt,0),'7',decode(a01.acc_code,'300000',amt,0),0)) as field_310000_cal, ");
          sqlCmd.append("                    sum(decode(bank_type,'6',decode(a01.acc_code,'210000',amt,'240000',amt,'250000',amt,'220000',amt,'260000',amt,0),'7',decode(a01.acc_code,'200000',amt,0),0)) as field_210000_cal, ");
          sqlCmd.append("                    sum(decode(a01.acc_code,'120700',amt,0))  as  field_120700, ");
          sqlCmd.append("                    sum(decode(a01.acc_code,'240100',amt,0))  as  field_240100, ");
          sqlCmd.append("                    decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'240300',amt,0), ");
          sqlCmd.append("                                                                '7',decode(a01.acc_code,'240200',amt,0),0)), ");
          sqlCmd.append("                                     '103',sum(decode(a01.acc_code,'240300',amt,0)),0)  as field_240300_cal, ");        
          sqlCmd.append("                    decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'240220',amt,0), ");
          sqlCmd.append("                                                                '7',decode(a01.acc_code,'240320',amt,0),0)), ");
          sqlCmd.append("                                     '103',sum(decode(a01.acc_code,'240220',amt,0)),0)  as field_240220_cal, ");     
          sqlCmd.append("                    decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'240210',amt,0), ");
          sqlCmd.append("                                                                '7',decode(a01.acc_code,'240310',amt,0),0)), ");
          sqlCmd.append("                                     '103',sum(decode(a01.acc_code,'240210',amt,0)),0)  as field_240210_cal, ");           
          sqlCmd.append("                    decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'240230',amt,0), ");
          sqlCmd.append("                                                                '7',decode(a01.acc_code,'240340',amt,0),0)), ");
          sqlCmd.append("                                     '103',sum(decode(a01.acc_code,'240230',amt,0)),0)  as field_240230_cal, "); 
          sqlCmd.append("                    decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'240240',amt,0), ");
          sqlCmd.append("                                                                '7',decode(a01.acc_code,'240330',amt,0),0)), ");
          sqlCmd.append("                                     '103',sum(decode(a01.acc_code,'240240',amt,0)),0)  as field_240240_cal, "); 
          sqlCmd.append("                    sum(decode(a01.acc_code,'250000',amt,0))  as  field_250000, ");
          sqlCmd.append("                    sum(decode(a01.acc_code,'260000',amt,0))  as  field_260000, ");     
          sqlCmd.append("                    sum(decode(a01.acc_code,'310000',amt,0))  as  field_310000, ");    
          sqlCmd.append("                    sum(decode(a01.acc_code,'310300',amt,0))  as  field_310300, ");           
          sqlCmd.append("                    decode(YEAR_TYPE, '102',sum(decode(a01.acc_code,'310800',amt,0)), ");
          sqlCmd.append("                                      '103',sum(decode(bank_type,'6',decode(a01.acc_code,'310800',amt,0), ");
          sqlCmd.append("                                                                 '7',0,0)),0)  as field_310800_cal, ");  
          sqlCmd.append("                    sum(decode(a01.acc_code,'320100',amt,0))  as  field_320100, ");                                                                                                                                              
          sqlCmd.append("                    sum(decode(a01.acc_code,'320300',amt,0))  as  field_320300, ");
          sqlCmd.append("                    sum(decode(bank_type,'6',decode(a01.acc_code,'520000',amt,0),'7',decode(a01.acc_code,'500000',amt,0),0)) as field_520000, ");                      
          sqlCmd.append("                    sum(decode(bank_type,'6',decode(a01.acc_code,'420000',amt,0),'7',decode(a01.acc_code,'400000',amt,0),0)) as field_420000, ");                                                                                                                                           
          sqlCmd.append("                    decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'120101',amt,'120102',amt,'120200',amt,'120301',amt,'120302',amt,'120700',amt,'150200',amt,0), ");
          sqlCmd.append("                                                                '7',decode(a01.acc_code,'120101',amt,'120102',amt,'120300',amt,'120401',amt,'120402',amt,'120700',amt,'150200',amt,0),0)), ");
          sqlCmd.append("                                     '103',sum(decode(a01.acc_code,'120101',amt,'120102',amt,'120200',amt,'120301',amt,'120302',amt,'120700',amt,'150200',amt,0)),0)  as fieldI_XA, ");
          sqlCmd.append("                    decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'120401',amt,'120402',amt,0),'7',decode(a01.acc_code,'120201',amt,'120202',amt,0),0)), ");
          sqlCmd.append("                                     '103',sum(decode(a01.acc_code,'120401',amt,'120402',amt,0)),0) as  fieldI_XB1, ");
          sqlCmd.append("                    decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt, 0),'7',decode(a01.acc_code,'240205',amt, '310800',amt,0),0)), ");
          sqlCmd.append("                                     '103',sum(decode(bank_type,'6',decode(a01.acc_code,'240305',amt,'251300',amt,'310800',amt, 0),'7',decode(a01.acc_code,'240305',amt, '251200',amt,0),0)),0)  as fieldI_XB2, ");
          sqlCmd.append("                    sum(decode(a01.acc_code,'120501',amt,'120502',amt,0))      as fieldI_XC1, ");
          sqlCmd.append("                    decode(YEAR_TYPE,'102', sum(decode(bank_type,'6',decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0), ");                                                             
          sqlCmd.append("                                                                 '7',decode(a01.acc_code,'240201','240202',amt,'240203',amt,'240204',amt,0),0)), ");
          sqlCmd.append("                                     '103', sum(decode(a01.acc_code,'240301',amt,'240302',amt,'240303',amt,'240304',amt,0)),0) as fieldI_XC2, ");           
          sqlCmd.append("                    sum(decode(a01.acc_code,'120600',amt,0))   as fieldI_XD1, ");
          sqlCmd.append("                    decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'240200',amt,0),'7',decode(a01.acc_code,'240300',amt,0),0)), ");
          sqlCmd.append("                                     '103',sum(decode(a01.acc_code,'240200',amt,0)),0) as fieldI_XD2, ");      
          sqlCmd.append("                    sum(decode(a01.acc_code,'150100',amt,0)) as fieldI_XE1, ");                                                                               
          sqlCmd.append("                    sum(decode(a01.acc_code,'250100',amt,0)) as fieldI_XE2, ");                                                                               
          sqlCmd.append("                    sum(decode(a01.acc_code,'310000',amt,'320000',amt,0))  as fieldI_XF1, ");    
          sqlCmd.append("                    decode(YEAR_TYPE,'102',sum(decode(a01.acc_code,'310800',amt,0)), ");
          sqlCmd.append("                                     '103',sum(decode(bank_type,'6',decode(a01.acc_code,'310800',amt,0),'7',0,0)) ,0)  as fieldI_XF3, ");     
          sqlCmd.append("                    sum(decode(a01.acc_code,'140000',amt,0))  as fieldI_XF2, ");
          sqlCmd.append("                    sum(decode(a01.acc_code,'220100',amt,'220200',amt,'220300',amt,'220400',amt,'220500',amt,'220600',amt,'220700',amt,'220800',amt, '220900',amt,'221000',amt,0)) ");                                                                                                   
          sqlCmd.append("                        - round(sum(decode(a01.acc_code,'220900',amt,0))/2,0)  as fieldI_Y, ");           
          sqlCmd.append("                    sum(decode(a01.acc_code,'991110',amt,0))  as  field_991110, ");   
          sqlCmd.append("                    sum(decode(a01.acc_code,'991120',amt,0))  as  field_991120, ");                  
          sqlCmd.append("                    sum(decode(a01.acc_code,'990410',amt,0))  as  field_990410, ");   
          sqlCmd.append("                    sum(decode(a01.acc_code,'990420',amt,0))  as  field_990420, ");      
          sqlCmd.append("                    sum(decode(a01.acc_code,'990620',amt,0))  as  field_990620  ");                                                                  
          sqlCmd.append("             from (select cd01.hsien_id,cd02.area_id,cd02.area_name ");
          sqlCmd.append("                   from (select * from cd01 where hsien_id=?)cd01 left join cd02 on cd01.hsien_id = cd02.hsien_id ");
          paramList.add(cityType) ;
          sqlCmd.append("                  ) cd01  ");
          sqlCmd.append("             left join (select cd01.hsien_id,cd01.hsien_name,wlx01.area_id,cd01.fr001w_output_order,bn01.bank_type,bn01.bank_no,bn01.bank_name ");
          sqlCmd.append("                    from (select * from cd01 where hsien_id=?)cd01, (select * from wlx01 where m_year=? )wlx01, (select * from bn01 where m_year=? and bank_type in (?) and bn01.bn_type <> '2' )bn01 ");
          paramList.add(cityType) ;
          paramList.add(u_year) ;
          paramList.add(u_year) ;
          paramList.add(bankType) ;
          sqlCmd.append("                    where wlx01.hsien_id=cd01.hsien_id ");  
          sqlCmd.append("                      and wlx01.bank_no=bn01.bank_no)wlx01 on cd01.hsien_id=wlx01.hsien_id and cd01.area_id =wlx01.area_id ");    
          sqlCmd.append("             left join (select (CASE WHEN (m_year <= 102) THEN '102' ");
          sqlCmd.append("                                WHEN (m_year > 102) THEN '103' ");
          sqlCmd.append("                                ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt "); 
          sqlCmd.append("                        from (select * from a01 "); 
          sqlCmd.append("                              where m_year = ? and m_month=12 ");
          paramList.add(S_YEAR) ;
          sqlCmd.append("                              union "); 
          sqlCmd.append("                              select m_year,m_month,bank_code,acc_code,amt from a02 ");
          sqlCmd.append("                              where m_year = ? and m_month=12 ");
          paramList.add(S_YEAR) ;
          sqlCmd.append("                              and acc_code in ('990410','990420','990620','991110','991120') ");
          sqlCmd.append("                             )a01 ");   
          sqlCmd.append("                        )a01 on  wlx01.bank_no = a01.bank_code ");  
          sqlCmd.append("             group by a01.YEAR_TYPE,cd01.hsien_id,cd01.area_id,cd01.area_name,wlx01.bank_type ");
          sqlCmd.append("        ) a01 "); 
          sqlCmd.append("        left join  ( "); //本年度12個月加總                  
          sqlCmd.append("             select cd01.hsien_id,cd01.area_id,cd01.area_name, ");                                                             
          sqlCmd.append("                    sum(decode(a01.acc_code,'220000',amt,0))  as  field_220000_12, ");
          sqlCmd.append("                    sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) as  field_CREDIT_12, "); 
          sqlCmd.append("                    sum(decode(a01.acc_code,'210000',amt,'220000',amt,'240000',amt,'250000',amt,0))  as  field_debt1_12, ");       
          sqlCmd.append("                    decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'210100',amt,0)),'7',sum(decode(a01.acc_code,'211000',amt,0)),0), ");
          sqlCmd.append("                                     '103',sum(decode(a01.acc_code,'210100',amt,0)),0) +  sum(decode(a01.acc_code,'250200',amt,0)) + "); 
          sqlCmd.append("                    decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'250900',amt,0)),'7',sum(decode(a01.acc_code,'250700',amt,0)),0), ");
          sqlCmd.append("                                     '103',sum(decode(a01.acc_code,'250900',amt,0)),0) + ");    
          sqlCmd.append("                    decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'251000',amt,0)),'7',sum(decode(a01.acc_code,'250800',amt,0)),0), ");
          sqlCmd.append("                                     '103',sum(decode(a01.acc_code,'251000',amt,0)),0) + ");
          sqlCmd.append("                    decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'251800',amt,0)),'7',sum(decode(a01.acc_code,'251300',amt,0)),0), ");
          sqlCmd.append("                                     '103',decode(bank_type,'6',sum(decode(a01.acc_code,'251800',amt,0)),'7',sum(decode(a01.acc_code,'251700',amt,0)),0),0)  as field_debt2_12 ");
          sqlCmd.append("             from (select cd01.hsien_id,cd02.area_id,cd02.area_name ");
          sqlCmd.append("                   from (select * from cd01 where hsien_id=?)cd01 left join cd02 on cd01.hsien_id = cd02.hsien_id ");
          paramList.add(cityType) ;
          sqlCmd.append("                  ) cd01 "); 
          sqlCmd.append("             left join (select cd01.hsien_id,cd01.hsien_name,wlx01.area_id,cd01.fr001w_output_order,bn01.bank_type,bn01.bank_no,bn01.bank_name ");
          sqlCmd.append("                    from (select * from cd01 where hsien_id=?)cd01, (select * from wlx01 where m_year=? )wlx01, (select * from bn01 where m_year=? and bank_type in (?) and bn01.bn_type <> '2' )bn01 ");
          paramList.add(cityType) ;
          paramList.add(u_year) ;
          paramList.add(u_year) ;
          paramList.add(bankType) ;
          sqlCmd.append("                    where wlx01.hsien_id=cd01.hsien_id ");  
          sqlCmd.append("                      and wlx01.bank_no=bn01.bank_no)wlx01 on cd01.hsien_id=wlx01.hsien_id and cd01.area_id =wlx01.area_id ");    
          sqlCmd.append("             left join (select (CASE WHEN (m_year <= 102) THEN '102' ");
          sqlCmd.append("                                  WHEN (m_year > 102) THEN '103' ");
          sqlCmd.append("                                  ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt "); 
          sqlCmd.append("                   from (select * from a01 "); 
          sqlCmd.append("                         where m_year = ? "); 
          paramList.add(S_YEAR) ;
          sqlCmd.append("                         and acc_code in ('220000','120000','120800','150300','210000','240000','250000','210100','250200','250900','251000','251700','251800','211000','250700','250800','251300') ");
          sqlCmd.append("                        )a01 ");   
          sqlCmd.append("                  )a01 on  wlx01.bank_no = a01.bank_code ");  
          sqlCmd.append("             group by a01.YEAR_TYPE,cd01.hsien_id,cd01.area_id,cd01.area_name,wlx01.bank_type ");
          sqlCmd.append("        ) a01_12 on a01_12.hsien_id = a01.hsien_id and a01_12.area_id = a01.area_id ");
          sqlCmd.append("        left join  ( ");  //上年度12個月加總                  
          sqlCmd.append("             select cd01.hsien_id,cd01.area_id,cd01.area_name, ");                                                             
          sqlCmd.append("                    sum(decode(a01.acc_code,'220000',amt,0))  as  field_220000_12_last, ");
          sqlCmd.append("                    sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) as  field_CREDIT_12_last ");
          sqlCmd.append("             from (select cd01.hsien_id,cd02.area_id,cd02.area_name ");
          sqlCmd.append("                   from (select * from cd01 where hsien_id=?)cd01 left join cd02 on cd01.hsien_id = cd02.hsien_id ");
          paramList.add(cityType) ;
          sqlCmd.append("             ) cd01 "); 
          sqlCmd.append("             left join (select cd01.hsien_id,cd01.hsien_name,wlx01.area_id,cd01.fr001w_output_order,bn01.bank_type,bn01.bank_no,bn01.bank_name ");
          sqlCmd.append("                    from (select * from cd01 where hsien_id=?)cd01, (select * from wlx01 where m_year=? )wlx01, (select * from bn01 where m_year=? and bank_type in (?) and bn01.bn_type <> '2' )bn01 ");
          paramList.add(cityType) ;
          paramList.add(u_year) ;
          paramList.add(u_year) ;
          paramList.add(bankType) ;
          sqlCmd.append("                    where wlx01.hsien_id=cd01.hsien_id ");  
          sqlCmd.append("                      and wlx01.bank_no=bn01.bank_no)wlx01 on cd01.hsien_id=wlx01.hsien_id and cd01.area_id =wlx01.area_id ");    
          sqlCmd.append("             left join (select (CASE WHEN (m_year <= 102) THEN '102' ");
          sqlCmd.append("                                     WHEN (m_year > 102) THEN '103' ");
          sqlCmd.append("                                     ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt "); 
          sqlCmd.append("                        from (select * from a01 "); 
          sqlCmd.append("                              where m_year = ? "); 
          paramList.add(Integer.parseInt(S_YEAR)-1) ;
          sqlCmd.append("                              and acc_code in ('220000','120000','120800','150300') ");
          sqlCmd.append("                             )a01 ");   
          sqlCmd.append("                       )a01 on  wlx01.bank_no = a01.bank_code ");  
          sqlCmd.append("             group by a01.YEAR_TYPE,cd01.hsien_id,cd01.area_id,cd01.area_name,wlx01.bank_type ");
          sqlCmd.append("        ) a01_12_last on a01_12_last.hsien_id = a01.hsien_id and a01_12_last.area_id = a01.area_id ");
          sqlCmd.append("        left join ( ");//本年底基本放款利率%加總  
          sqlCmd.append("                  select bn01.bank_type,cd02.hsien_id,cd02.area_id, "); 
          sqlCmd.append("                         cd02.area_name,sum(w.base_base_rate) as base_base_rate, count(*) as num ");  
          sqlCmd.append("                  from wlx_s_rate w,(select * from bn01 where m_year=?)bn01,(select * from wlx01 where m_year=? )wlx01,cd02 "); 
          paramList.add(u_year) ;
          paramList.add(u_year) ;
          sqlCmd.append("                  where w.bank_no=bn01.bank_no ");  
          sqlCmd.append("                  and w.bank_no=wlx01.bank_no ");  
          sqlCmd.append("                  and wlx01.area_id = cd02.area_id ");  
          sqlCmd.append("                  and w.m_year=? ");
          sqlCmd.append("                  and bn01.bank_type=? ");
          sqlCmd.append("                  and wlx01.hsien_id=? ");
          paramList.add(S_YEAR) ;
          paramList.add(bankType) ;
          paramList.add(cityType) ;
          sqlCmd.append("                  group by bn01.bank_type,cd02.hsien_id,cd02.area_id,cd02.area_name ");
          sqlCmd.append("                  )wlx_s_rate on a01.hsien_id = wlx_s_rate.hsien_id and a01.area_id = wlx_s_rate.area_id ");
          sqlCmd.append("    )a01 ");        
          sqlCmd.append("    order by area_id ");
          
          List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,
                  "hsien_id,area_id,area_name,field_220000_12_last,field_220000_12," +
                  "field_220000,field_220100,field_220200,field_220300,field_220400,field_220500," +
                  "field_220600,field_220700,field_220800,field_220900,field_221000,field_990420,field_990620," +
                  "field_credit_12_last,field_credit_12,field_credit,field_120101,field_120102,field_120200_cal," +
                  "field_120401_cal,field_120501_cal,field_120601,field_120602,field_120603_cal,field_120604_cal," +
                  "field_120700,field_990410,field_110000,field_110100,field_110300_cal,field_110400,field_111100_cal," +      
                  "field_110000_other,field_120000,field_140000,field_150000_other,field_190000_cal,field_210000," +
                  "field_210400_cal,field_210700,field_210000_other,field_220000_220900,field_240100,field_240300_cal," +
                  "field_240220_cal,field_240210_cal,field_240230_cal,field_240240_cal,field_250000_cal,field_210000_cal," +
                  "field_310000,field_310300,field_310800_cal,field_310000_other,field_320100,field_320300,field_310000_cal," +
                  "field_rate1,field_rate2,field_rate3,field_rate4,field_991110,field_991120,field_dc_rate," +       
                  "field_rate5,field_rate6,field_rate7,field_rate8,field_rate9,field_rate10,field_rate11,field_base_base_rate");
          System.out.println("dbData.size=" + dbData.size());
          
          int rowNo=2;
          int cellNo=3;
          DataObject bean = null;
          if(dbData.size()>0){
              for (int i = 0; i < 76; i++) {
                  row = sheet.getRow(i + 2);
                  for (int j = 0; j < dbData.size(); j++) {
                    cell = row.getCell( (short) (j + 3));
                    if (cell == null) {
                      cell = row.createCell( (short) (j + 3));
                      cell.setCellStyle( (row.getCell( (short) 3)).getCellStyle());
                    }
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
                  }
                }
              for(int i=0;i<dbData.size();i++){  
                  bean = (DataObject)dbData.get(i);
                  hsien_id = (bean.getValue("hsien_id")==null)?"":(bean.getValue("hsien_id")).toString();
                  area_id = (bean.getValue("area_id")==null)?"":(bean.getValue("area_id")).toString();
                  area_name = (bean.getValue("area_name")==null)?"":(bean.getValue("area_name")).toString();
                  field_220000_12_last = (bean.getValue("field_220000_12_last")==null)?"":(bean.getValue("field_220000_12_last")).toString();
                  field_220000_12 = (bean.getValue("field_220000_12")==null)?"":(bean.getValue("field_220000_12")).toString();
                  field_220000 = (bean.getValue("field_220000")==null)?"":(bean.getValue("field_220000")).toString();
                  field_220100 = (bean.getValue("field_220100")==null)?"":(bean.getValue("field_220100")).toString();
                  field_220200 = (bean.getValue("field_220200")==null)?"":(bean.getValue("field_220200")).toString();
                  field_220300 = (bean.getValue("field_220300")==null)?"":(bean.getValue("field_220300")).toString();
                  field_220400 = (bean.getValue("field_220400")==null)?"":(bean.getValue("field_220400")).toString();
                  field_220500 = (bean.getValue("field_220500")==null)?"":(bean.getValue("field_220500")).toString();
                  field_220600 = (bean.getValue("field_220600")==null)?"":(bean.getValue("field_220600")).toString();
                  field_220700 = (bean.getValue("field_220700")==null)?"":(bean.getValue("field_220700")).toString();
                  field_220800 = (bean.getValue("field_220800")==null)?"":(bean.getValue("field_220800")).toString();
                  field_220900 = (bean.getValue("field_220900")==null)?"":(bean.getValue("field_220900")).toString();
                  field_221000 = (bean.getValue("field_221000")==null)?"":(bean.getValue("field_221000")).toString();
                  field_990420 = (bean.getValue("field_990420")==null)?"":(bean.getValue("field_990420")).toString();
                  field_990620 = (bean.getValue("field_990620")==null)?"":(bean.getValue("field_990620")).toString();
                  field_CREDIT_12_last = (bean.getValue("field_credit_12_last")==null)?"":(bean.getValue("field_credit_12_last")).toString();
                  field_CREDIT_12 = (bean.getValue("field_credit_12")==null)?"":(bean.getValue("field_credit_12")).toString();
                  field_CREDIT = (bean.getValue("field_credit")==null)?"":(bean.getValue("field_credit")).toString();
                  field_120101 = (bean.getValue("field_120101")==null)?"":(bean.getValue("field_120101")).toString();
                  field_120102 = (bean.getValue("field_120102")==null)?"":(bean.getValue("field_120102")).toString();
                  field_120200_cal = (bean.getValue("field_120200_cal")==null)?"":(bean.getValue("field_120200_cal")).toString();
                  field_120401_cal = (bean.getValue("field_120401_cal")==null)?"":(bean.getValue("field_120401_cal")).toString();
                  field_120501_cal = (bean.getValue("field_120501_cal")==null)?"":(bean.getValue("field_120501_cal")).toString();
                  field_120601 = (bean.getValue("field_120601")==null)?"":(bean.getValue("field_120601")).toString();
                  field_120602 = (bean.getValue("field_120602")==null)?"":(bean.getValue("field_120602")).toString();
                  field_120603_cal = (bean.getValue("field_120603_cal")==null)?"":(bean.getValue("field_120603_cal")).toString();
                  field_120604_cal = (bean.getValue("field_120604_cal")==null)?"":(bean.getValue("field_120604_cal")).toString();
                  field_120700 = (bean.getValue("field_120700")==null)?"":(bean.getValue("field_120700")).toString();
                  field_990410 = (bean.getValue("field_990410")==null)?"":(bean.getValue("field_990410")).toString();
                  field_110000 = (bean.getValue("field_110000")==null)?"":(bean.getValue("field_110000")).toString();
                  field_110100 = (bean.getValue("field_110100")==null)?"":(bean.getValue("field_110100")).toString();
                  field_110300_cal = (bean.getValue("field_110300_cal")==null)?"":(bean.getValue("field_110300_cal")).toString();
                  field_110400 = (bean.getValue("field_110400")==null)?"":(bean.getValue("field_110400")).toString();
                  field_111100_cal = (bean.getValue("field_111100_cal")==null)?"":(bean.getValue("field_111100_cal")).toString();      
                  field_110000_other = (bean.getValue("field_110000_other")==null)?"":(bean.getValue("field_110000_other")).toString();
                  field_120000 = (bean.getValue("field_120000")==null)?"":(bean.getValue("field_120000")).toString();
                  field_140000 = (bean.getValue("field_140000")==null)?"":(bean.getValue("field_140000")).toString();
                  field_150000_other = (bean.getValue("field_150000_other")==null)?"":(bean.getValue("field_150000_other")).toString();
                  field_190000_cal = (bean.getValue("field_190000_cal")==null)?"":(bean.getValue("field_190000_cal")).toString();
                  field_210000 = (bean.getValue("field_210000")==null)?"":(bean.getValue("field_210000")).toString();
                  field_210400_cal = (bean.getValue("field_210400_cal")==null)?"":(bean.getValue("field_210400_cal")).toString();
                  field_210700 = (bean.getValue("field_210700")==null)?"":(bean.getValue("field_210700")).toString();
                  field_210000_other = (bean.getValue("field_210000_other")==null)?"":(bean.getValue("field_210000_other")).toString();
                  field_220000_220900 = (bean.getValue("field_220000_220900")==null)?"":(bean.getValue("field_220000_220900")).toString();
                  field_240100 = (bean.getValue("field_240100")==null)?"":(bean.getValue("field_240100")).toString();
                  field_240300_cal = (bean.getValue("field_240300_cal")==null)?"":(bean.getValue("field_240300_cal")).toString();
                  field_240220_cal = (bean.getValue("field_240220_cal")==null)?"":(bean.getValue("field_240220_cal")).toString();
                  field_240210_cal = (bean.getValue("field_240210_cal")==null)?"":(bean.getValue("field_240210_cal")).toString();
                  field_240230_cal = (bean.getValue("field_240230_cal")==null)?"":(bean.getValue("field_240230_cal")).toString();
                  field_240240_cal = (bean.getValue("field_240240_cal")==null)?"":(bean.getValue("field_240240_cal")).toString();
                  field_250000_cal = (bean.getValue("field_250000_cal")==null)?"":(bean.getValue("field_250000_cal")).toString();
                  field_210000_cal = (bean.getValue("field_210000_cal")==null)?"":(bean.getValue("field_210000_cal")).toString();
                  field_310000 = (bean.getValue("field_310000")==null)?"":(bean.getValue("field_310000")).toString();
                  field_310300 = (bean.getValue("field_310300")==null)?"":(bean.getValue("field_310300")).toString();
                  field_310800_cal = (bean.getValue("field_310800_cal")==null)?"":(bean.getValue("field_310800_cal")).toString();
                  field_310000_other = (bean.getValue("field_310000_other")==null)?"":(bean.getValue("field_310000_other")).toString();
                  field_320100 = (bean.getValue("field_320100")==null)?"":(bean.getValue("field_320100")).toString();
                  field_320300 = (bean.getValue("field_320300")==null)?"":(bean.getValue("field_320300")).toString();
                  field_310000_cal = (bean.getValue("field_310000_cal")==null)?"":(bean.getValue("field_310000_cal")).toString();
                  field_rate1 = (bean.getValue("field_rate1")==null)?"":(bean.getValue("field_rate1")).toString();
                  field_rate2 = (bean.getValue("field_rate2")==null)?"":(bean.getValue("field_rate2")).toString();
                  field_rate3 = (bean.getValue("field_rate3")==null)?"":(bean.getValue("field_rate3")).toString();
                  field_rate4 = (bean.getValue("field_rate4")==null)?"":(bean.getValue("field_rate4")).toString();
                  field_991110 = (bean.getValue("field_991110")==null)?"":(bean.getValue("field_991110")).toString();
                  field_991120 = (bean.getValue("field_991120")==null)?"":(bean.getValue("field_991120")).toString();
                  field_DC_RATE = (bean.getValue("field_dc_rate")==null)?"":(bean.getValue("field_dc_rate")).toString();       
                  field_rate5 = (bean.getValue("field_rate5")==null)?"":(bean.getValue("field_rate5")).toString();
                  field_rate6 = (bean.getValue("field_rate6")==null)?"":(bean.getValue("field_rate6")).toString();
                  field_rate7 = (bean.getValue("field_rate7")==null)?"":(bean.getValue("field_rate7")).toString();
                  field_rate8 = (bean.getValue("field_rate8")==null)?"":(bean.getValue("field_rate8")).toString();
                  field_rate9 = (bean.getValue("field_rate9")==null)?"":(bean.getValue("field_rate9")).toString();
                  field_rate10 = (bean.getValue("field_rate10")==null)?"":(bean.getValue("field_rate10")).toString();
                  field_rate11 = (bean.getValue("field_rate11")==null)?"":(bean.getValue("field_rate11")).toString();
                  field_base_base_rate = (bean.getValue("field_base_base_rate")==null)?"":(bean.getValue("field_base_base_rate")).toString();
                  
                  //農漁會別
                  row = sheet.getRow(rowNo);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(hsien_id+area_id);
                  row = sheet.getRow(rowNo+1);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(area_name);
                  //存款業務-上年度存款平均餘額
                  row = sheet.getRow(rowNo+2);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_220000_12_last));
                  //存款業務-本年度存款平均餘額
                  row = sheet.getRow(rowNo+3);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_220000_12));
                  //存款業務-本年底存款餘額
                  row = sheet.getRow(rowNo+4);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_220000));
                  //存款業務-支票存款
                  row = sheet.getRow(rowNo+5);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_220100));
                  //存款業務-保付支票
                  row = sheet.getRow(rowNo+6);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_220200));
                  //存款業務-活期存款
                  row = sheet.getRow(rowNo+7);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_220300));
                  //存款業務-活期儲蓄存款
                  row = sheet.getRow(rowNo+8);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_220400));
                  //存款業務-員工活期儲蓄存款
                  row = sheet.getRow(rowNo+9);
                  cell = row.getCell((short)cellNo);
                  cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                  cell.setCellValue(fmtMicrometer(field_220500));
                  //存款業務-定期存款
                  row = sheet.getRow(rowNo+10);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_220600));
                  //存款業務-定期儲蓄存款
                  row = sheet.getRow(rowNo+11);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_220700));
                  //存款業務-員工定期儲蓄存款
                  row = sheet.getRow(rowNo+12);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_220800));
                  //存款業務-公庫存款
                  row = sheet.getRow(rowNo+13);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_220900));
                  //存款業務-本會支票
                  row = sheet.getRow(rowNo+14);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_221000));
                  //存款業務-贊助會員存款
                  row = sheet.getRow(rowNo+15);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_990420));
                  //存款業務-非會員存款
                  row = sheet.getRow(rowNo+16);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_990620));

                  //放款業務-上年度存款平均餘額
                  row = sheet.getRow(rowNo+17);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_CREDIT_12_last));
                  //放款業務-本年度存款平均餘額
                  row = sheet.getRow(rowNo+18);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_CREDIT_12));
                  //放款業務-本年底存款餘額
                  row = sheet.getRow(rowNo+19);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_CREDIT));
                  //放款業務-一般放款 -- 無擔保
                  row = sheet.getRow(rowNo+20);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_120101));
                  //放款業務-一般放款 -- 擔    保
                  row = sheet.getRow(rowNo+21);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_120102));
                  //放款業務-貼現及透支
                  row = sheet.getRow(rowNo+22);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_120200_cal));
                  //放款業務-統一農貸
                  row = sheet.getRow(rowNo+23);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_120401_cal));
                  //放款業務-專案放款
                  row = sheet.getRow(rowNo+24);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_120501_cal));
                  //放款業務-農建放款
                  row = sheet.getRow(rowNo+25);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_120601));
                  //放款業務-農機放款
                  row = sheet.getRow(rowNo+26);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_120602));
                  //放款業務-購地放款
                  row = sheet.getRow(rowNo+27);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_120603_cal));
                  //放款業務-農宅放款
                  row = sheet.getRow(rowNo+28);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_120604_cal));
                  //放款業務-內部融資
                  row = sheet.getRow(rowNo+29);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_120700));
                  //放款業務-贊助會員放款總額
                  row = sheet.getRow(rowNo+30);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_990410));
                  //資產-流動資產
                  row = sheet.getRow(rowNo+31);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_110000));
                  //資產-庫存現金
                  row = sheet.getRow(rowNo+32);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_110100));
                  //資產-存放行庫
                  row = sheet.getRow(rowNo+33);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_110300_cal));
                  //資產-繳存存款準備金
                  row = sheet.getRow(rowNo+34);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_110400));
                  //資產-應收利息
                  row = sheet.getRow(rowNo+35);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_111100_cal));
                  //資產-其他流動資產
                  row = sheet.getRow(rowNo+36);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_110000_other));
                  //資產-放款淨額(減備抵呆帳)
                  row = sheet.getRow(rowNo+37);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_120000));
                  //資產-固定資產淨額
                  row = sheet.getRow(rowNo+38);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_140000));
                  //資產-其他資產(含出資及往來)
                  row = sheet.getRow(rowNo+39);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_150000_other));
                  //資產-資產總額
                  row = sheet.getRow(rowNo+40);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_190000_cal));
                  
                  //負債-流動負債
                  row = sheet.getRow(rowNo+41);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_210000));
                  //負債-短期借款
                  row = sheet.getRow(rowNo+42);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_210400_cal));
                  //負債-應付利息
                  row = sheet.getRow(rowNo+43);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_210700));
                  //負債-其他流動負債
                  row = sheet.getRow(rowNo+44);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_210000_other));
                  //負債-存款(含公庫存款)
                  row = sheet.getRow(rowNo+45);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_220000_220900));
                  //負債-長期借款
                  row = sheet.getRow(rowNo+46);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_240100));
                  //負債-借入專案放款資金
                  row = sheet.getRow(rowNo+47);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_240300_cal));
                  //負債-借入農機放款資金
                  row = sheet.getRow(rowNo+48);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_240220_cal));
                  //負債-借入農建放款資金
                  row = sheet.getRow(rowNo+49);
                  cell = row.getCell((short)cellNo);
                  cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                  cell.setCellValue(fmtMicrometer(field_240210_cal));
                  //負債-借入購地放款資金
                  row = sheet.getRow(rowNo+50);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_240230_cal));
                  //負債-借入農宅放款資金
                  row = sheet.getRow(rowNo+51);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_240240_cal));
                  //負債-其他負債(含往來)
                  row = sheet.getRow(rowNo+52);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_250000_cal));
                  //負債-負債總額
                  row = sheet.getRow(rowNo+53);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_210000_cal));
                  
                  //淨值-事業基金及公積
                  row = sheet.getRow(rowNo+54);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_310000));
                  //淨值-法定公積
                  row = sheet.getRow(rowNo+55);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_310300));
                  //淨值-統一農貸公積
                  row = sheet.getRow(rowNo+56);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_310800_cal));
                  //淨值-其他公積
                  row = sheet.getRow(rowNo+57);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_310000_other));
                  //淨值-累積盈餘
                  row = sheet.getRow(rowNo+58);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_320100));
                  //淨值-本期損益
                  row = sheet.getRow(rowNo+59);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_320300));
                  //淨值-淨值總額
                  row = sheet.getRow(rowNo+60);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_310000_cal));

                  //成本收益-本年度存款平均利率%
                  row = sheet.getRow(rowNo+61);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(MakesUpZero(fmtMicrometer(field_rate1)));
                  //成本收益-本年度放款平均利率%
                  row = sheet.getRow(rowNo+62);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(MakesUpZero(fmtMicrometer(field_rate2)));
                  //成本收益-總資金平均成本率%
                  row = sheet.getRow(rowNo+63);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(MakesUpZero(fmtMicrometer(field_rate3)));
                  //成本收益-總資金平均收益率%
                  row = sheet.getRow(rowNo+64);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(MakesUpZero(fmtMicrometer(field_rate4)));
                 
                  //放款限額-每一會員放款最高額
                  row = sheet.getRow(rowNo+65);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_991110));
                  //放款限額-會員無擔保放款最高額
                  row = sheet.getRow(rowNo+66);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(fmtMicrometer(field_991120));
                  
                  //各項經營比率-本年底存放比率%
                  row = sheet.getRow(rowNo+67);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(MakesUpZero(fmtMicrometer(field_DC_RATE)));
                  //各項經營比率-內部融資比率%
                  row = sheet.getRow(rowNo+68);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(MakesUpZero(fmtMicrometer(field_rate5)));
                  //各項經營比率-贊助會員放款占存款率%
                  row = sheet.getRow(rowNo+69);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(MakesUpZero(fmtMicrometer(field_rate6)));
                  //各項經營比率-流動性比率%
                  row = sheet.getRow(rowNo+70);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(MakesUpZero(fmtMicrometer(field_rate7)));
                  //各項經營比率-固定資產淨額占淨值率%
                  row = sheet.getRow(rowNo+71);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(MakesUpZero(fmtMicrometer(field_rate8)));
                  //各項經營比率-融通資金比率%
                  row = sheet.getRow(rowNo+72);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(MakesUpZero(fmtMicrometer(field_rate9)));
                  //各項經營比率-淨值占負債總額比率%
                  row = sheet.getRow(rowNo+73);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(MakesUpZero(fmtMicrometer(field_rate10)));
                  //各項經營比率-淨值占存款總額比率%
                  row = sheet.getRow(rowNo+74);
                  cell = row.getCell((short)cellNo);
                  cell.setCellValue(MakesUpZero(fmtMicrometer(field_rate11)));
                  //各項經營比率-本年底基本放款利率%
                  row = sheet.getRow(rowNo+75);
                  cell = row.getCell((short)cellNo);
                  cell.setEncoding(HSSFCell.ENCODING_UTF_16); 
                  cell.setCellValue(MakesUpZero(fmtMicrometer(field_base_base_rate)));
                  cellNo++;
              }
          }else{
              row = sheet.getRow(2);
              cell = row.getCell( (short) 3);
              cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
              cell.setCellValue("查無資料");
              sheet.addMergedRegion(new Region((short)2, (short)3,(short)3, (short)8));//跨行(第幾行,開始欄位數,跨幾行,結束欄位數)
          }
      }

      FileOutputStream fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "某年底個別農漁會信用部業務經營概況.xls");
      HSSFFooter footer = sheet.getFooter();
      footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
      footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));

      wb.write(fout);
      //儲存
      fout.close();
    }catch (Exception e) {
      System.out.println("//RptAN012W createRpt() Have Error.....");
      e.printStackTrace();
      System.out.println(e.toString());
      System.out.println("//-------------------------------------");
    }

    if(debug) System.out.println("RptAN012W End ...");
    return errMsg;
  }

    /** 
     * 格式化數字為千分位顯示； 
     * @param 要格式化的數字； 
     * @return 
     */  
    public static String fmtMicrometer(String text)  {  
        DecimalFormat df = null;  
        if(text.indexOf(".") > 0)  {  
            if(text.length() - text.indexOf(".")-1 == 0)  {  
                df = new DecimalFormat("###,##0.");  
            }else if(text.length() - text.indexOf(".")-1 == 1) {  
                df = new DecimalFormat("###,##0.0");  
            }else  {  
                df = new DecimalFormat("###,##0.00");  
            }  
        }else{  
            df = new DecimalFormat("###,##0");  
        }  
        double number = 0.0;  
        try {  
             number = Double.parseDouble(text);  
        }catch (Exception e) {  
            number = 0.0;  
        }  
        return df.format(number);  
    }
    
    /** 
     * 百分比顯示到小數點2位,不足者補0； 
     * @param 要格式化的數字； 
     * @return 
     */ 
    public static String MakesUpZero(String str) {
        String[] strarray=str.split("\\.");
        if(strarray.length==2){
            if(strarray[1].length()==1){
                str = str+"0";
            }
        }else if(strarray.length==1){
            str = str+".00";
        }
        return str;
    }
    
}
