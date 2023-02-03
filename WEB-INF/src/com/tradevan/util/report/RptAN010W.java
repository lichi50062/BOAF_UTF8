/*
 * Created on 2006/10/30 by ABYSS Brenda
 * 某年底全體農漁會信用部各類財務比率
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
 * 
 * fixed 99.06.07 sql injection by 2808
 * 2012.11.05 connection 沒有全部關 by 2295
 * 2013.06.03 fixed sql by 2968
 * fixed 102.08.30 外包的connection拿掉 by 2968
 */

package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.text.DecimalFormat;
import java.util.*;
import java.sql.*;

import com.tradevan.util.DBManager;
import com.tradevan.util.Utility;
import com.tradevan.util.Utility_report;
import com.tradevan.util.dao.DataObject;
import com.tradevan.util.dao.RdbCommonDao;
import java.math.BigDecimal;

public class RptAN010W {
  private static boolean debug = false;
  public static String createRpt(String S_YEAR, String bankType ,String tmUnit) {
    if(debug) System.out.println("RptAN010W Start ...");
    String field_rate1 =""; 
    String field_DC_RATE ="";
    String field_rate2 ="";
    String field_rate3 ="";
    String field_rate4 ="";
    String field_rate5 ="";
    String field_rate6 ="";
    String field_rate7 ="";
    String field_rate8 ="";
    String field_rate9 ="";
    String field_rate10 ="";
    String field_rate11 ="";
    String field_310000 ="";
    DataObject bean = null;
    
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
      if(debug) System.out.println("某年底全體農漁會信用部各類財務比率.xls");

      finput = new FileInputStream(xlsDir +
                                   System.getProperty("file.separator") +
                                   "某年底全體農漁會信用部各類財務比率.xls");

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
      String title = " 年底全體農漁會信用部各類財務比率";
      if(bankType.equals("6")){
        title = " 年底全體農會信用部各類財務比率";
      }else if(bankType.equals("7")){
        title = " 年底全體漁會信用部各類財務比率";
      }
      cell.setCellValue(S_YEAR + title);

      //設定年月及單位資料============================================
      row = sheet.getRow(1);
      cell = row.getCell( (short) 0);
      cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
      cell.setCellValue("單位：新台幣 " + unitName + " ，%");

      if(qYear >= nowYear){
        row = sheet.getRow(0);
        cell = row.getCell( (short) 0);
        cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
        cell.setCellValue(S_YEAR + title +" 查無資料");
       // sheet.addMergedRegion(new Region((short)3, (short)0,(short)3, (short)7));//跨行(第幾行,開始欄位數,跨幾行,結束欄位數)
      }else{
        //依取得會計科目金額
        sqlCmd.setLength(0) ;
        paramList.clear() ;
        sqlCmd.append("select  decode(a01.field_220000_220900,0,0,round(a01.field_110000 /  a01.field_220000_220900 *100 ,2))  as field_rate1, ");//流動比率  
        sqlCmd.append("        decode(a01.fieldI_Y,0,0, ");                                  
        sqlCmd.append("              round((a01.fieldI_XA ");                               
        sqlCmd.append("                  + decode(sign(a01.fieldI_XB1 - a01.fieldI_XB2),-1,0,(a01.fieldI_XB1 - a01.fieldI_XB2)) ");   
        sqlCmd.append("                  + decode(sign(a01.fieldI_XC1 - a01.fieldI_XC2),-1,0,(a01.fieldI_XC1 - a01.fieldI_XC2)) ");   
        sqlCmd.append("                  + decode(sign(a01.fieldI_XD1 - a01.fieldI_XD2),-1,0,(a01.fieldI_XD1 - a01.fieldI_XD2)) ");   
        sqlCmd.append("                  + decode(sign(a01.fieldI_XE1 - a01.fieldI_XE2),-1,0,(a01.fieldI_XE1 - a01.fieldI_XE2)) ");   
        sqlCmd.append("                  - decode(sign(a01.fieldI_XF1 - a01.fieldI_XF3 -  a01.fieldI_XF2),-1,0,(a01.fieldI_XF1 - a01.fieldI_XF3 - a01.fieldI_XF2)) ");   
        sqlCmd.append("                    )/a01.fieldI_Y * 100,2))  as  field_DC_RATE, ");//存放比率            
        sqlCmd.append("        decode(a01.field_220000,0,0,round(a01.field_520100 /  (a01.field_220000/12) *100 ,2))  as field_rate2, ");//存款平均利率
        sqlCmd.append("        decode(a01.field_CREDIT,0,0,round(a01.field_420100 /  (a01.field_CREDIT/12) *100 ,2))  as field_rate3, ");//放款平均利率
        sqlCmd.append("        decode(a01.field_220000,0,0,round(a01.field_210200 /  a01.field_220000 *100 ,2))  as field_rate4, "); //融通資金比率
        sqlCmd.append("        decode(a01.field_990420,0,0,round(a01.field_990410 /  a01.field_990420 *100 ,2))  as field_rate5, ");//贊助會員放款占贊助會員放款比率
        sqlCmd.append("        decode(a01.field_310000,0,0,round(a01.fieldI_XF2 /  a01.field_310000 *100 ,2))  as field_rate6, ");//固定資產淨額占淨值比率
        sqlCmd.append("        decode(a01.field_210000,0,0,round(a01.field_310000 /  a01.field_210000 *100 ,2))  as field_rate7, ");//淨值占負債總額比率
        sqlCmd.append("        decode(a01.field_220000,0,0,round(a01.field_310000 /  a01.field_220000 *100 ,2))  as field_rate8, ");//淨值占存款總額比率
        sqlCmd.append("        decode(a01.field_310000,0,0,round(a01.field_120700 /  a01.field_310000 *100 ,2))  as field_rate9, ");//內部融資占淨值比率
        sqlCmd.append("        decode((a01.field_debt1-a01.field_debt2),0,0,round(a01.field_520000 /  ((a01.field_debt1-a01.field_debt2)/12) *100 ,2))  as field_rate10, ");//總資金平均成本率
        sqlCmd.append("        decode((a01.field_debt1-a01.field_debt2),0,0,round(a01.field_420000 /  ((a01.field_debt1-a01.field_debt2)/12) *100 ,2))  as field_rate11, ");//總資金平均收益率
        sqlCmd.append("        round( a01.field_310000 /?,0)    as  field_310000 ");//淨值       
        paramList.add(unit) ;
        sqlCmd.append("from ( ");   
        sqlCmd.append("      select SUM(field_110000) field_110000,SUM(field_220000) field_220000,SUM(field_220000_220900) field_220000_220900,SUM(field_520100) field_520100, ");
        sqlCmd.append("             SUM(field_420100) field_420100, ");
        sqlCmd.append("             SUM(field_CREDIT) field_CREDIT,SUM(field_210200) field_210200,SUM(field_990410) field_990410,SUM(field_990420) field_990420, ");
        sqlCmd.append("             SUM(field_310000) field_310000,SUM(field_210000) field_210000,SUM(field_120700) field_120700,SUM(field_520000) field_520000, ");
        sqlCmd.append("             SUM(field_debt1) field_debt1, SUM(field_debt2) field_debt2,SUM(field_420000) field_420000,SUM(fieldI_XA) fieldI_XA, ");                
        sqlCmd.append("             SUM(fieldI_XB1) fieldI_XB1, SUM(fieldI_XB2) fieldI_XB2, SUM(fieldI_XC1) fieldI_XC1,SUM(fieldI_XC2) fieldI_XC2, ");               
        sqlCmd.append("             SUM(fieldI_XD1) fieldI_XD1, SUM(fieldI_XD2) fieldI_XD2, SUM(fieldI_XE1) fieldI_XE1,SUM(fieldI_XE2) fieldI_XE2, ");               
        sqlCmd.append("             SUM(fieldI_XF1) fieldI_XF1,SUM(fieldI_XF3) fieldI_XF3,SUM(fieldI_XF2)  fieldI_XF2, SUM(fieldI_Y) fieldI_Y ");          
        sqlCmd.append("      from (  ");                  
        sqlCmd.append("             select a01.m_year, ");
        sqlCmd.append("                    sum(decode(a01.acc_code,'110000',amt,0))  as  field_110000, ");
        sqlCmd.append("                    sum(decode(a01.acc_code,'220000',amt,0))  as  field_220000, ");
        sqlCmd.append("                    sum(decode(a01.acc_code,'220000',amt,0)) - sum(decode(a01.acc_code,'220900',amt,0))   as  field_220000_220900, ");
        sqlCmd.append("                    sum(decode(a01.acc_code,'520100',amt,0))  as  field_520100, ");
        sqlCmd.append("                    sum(decode(bank_type,'6',decode(a01.acc_code,'420100',amt,'420170',amt,0),'7',decode(a01.acc_code,'420100',amt,0),0)) as field_420100, ");
        sqlCmd.append("                    sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) as  field_CREDIT, "); 
        sqlCmd.append("                    decode(YEAR_TYPE,'102',sum(decode(bank_type,'6',decode(a01.acc_code,'210200',amt,'210300',amt,'210400',amt,0), ");
        sqlCmd.append("                                                                '7',decode(a01.acc_code,'210100',amt,'210300',amt,'210200',amt,0),0)), ");
        sqlCmd.append("                                     '103',sum(decode(a01.acc_code,'210200',amt,'210300',amt,'210400',amt,0)),0)  as field_210200, ");    
        sqlCmd.append("                    sum(decode(a01.acc_code,'990410',amt,0))  as  field_990410, ");   
        sqlCmd.append("                    sum(decode(a01.acc_code,'990420',amt,0))  as  field_990420, ");                      
        sqlCmd.append("                    sum(decode(bank_type,'6',decode(a01.acc_code,'310000',amt,'320000',amt,0),'7',decode(a01.acc_code,'300000',amt,0),0)) as field_310000, ");
        sqlCmd.append("                    sum(decode(bank_type,'6',decode(a01.acc_code,'210000',amt,'240000',amt,'250000',amt,'220000',amt,'260000',amt,0),'7',decode(a01.acc_code,'200000',amt,0),0)) as field_210000, ");
        sqlCmd.append("                    sum(decode(a01.acc_code,'120700',amt,0))  as  field_120700, ");
        sqlCmd.append("                    sum(decode(bank_type,'6',decode(a01.acc_code,'520000',amt,0),'7',decode(a01.acc_code,'500000',amt,0),0)) as field_520000, ");   
        sqlCmd.append("                    sum(decode(a01.acc_code,'210000',amt,'220000',amt,'240000',amt,'250000',amt,0))  as  field_debt1, ");       
        sqlCmd.append("                    decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'210100',amt,0)),'7',sum(decode(a01.acc_code,'211000',amt,0)),0), ");
        sqlCmd.append("                                     '103',sum(decode(a01.acc_code,'210100',amt,0)),0) +  sum(decode(a01.acc_code,'250200',amt,0)) + "); 
        sqlCmd.append("                    decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'250900',amt,0)),'7',sum(decode(a01.acc_code,'250700',amt,0)),0), ");
        sqlCmd.append("                                     '103',sum(decode(a01.acc_code,'250900',amt,0)),0) + ");    
        sqlCmd.append("                    decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'251000',amt,0)),'7',sum(decode(a01.acc_code,'250800',amt,0)),0), ");
        sqlCmd.append("                                     '103',sum(decode(a01.acc_code,'251000',amt,0)),0) + ");
        sqlCmd.append("                    decode(YEAR_TYPE,'102',decode(bank_type,'6',sum(decode(a01.acc_code,'251800',amt,0)),'7',sum(decode(a01.acc_code,'251300',amt,0)),0), ");
        sqlCmd.append("                                     '103',decode(bank_type,'6',sum(decode(a01.acc_code,'251800',amt,0)),'7',sum(decode(a01.acc_code,'251700',amt,0)),0),0)  as field_debt2, ");
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
        sqlCmd.append("                        - round(sum(decode(a01.acc_code,'220900',amt,0))/2,0)  as fieldI_Y ");                                                                     
        sqlCmd.append("             from (select (CASE WHEN (m_year <= 102) THEN '102' ");
        sqlCmd.append("                                WHEN (m_year > 102) THEN '103' ");
        sqlCmd.append("                                ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt "); 
        sqlCmd.append("                   from (select * from a01 "); 
        sqlCmd.append("                         where m_year = ? and m_month=12 ");
        paramList.add(S_YEAR) ;
        sqlCmd.append("                         union "); 
        sqlCmd.append("                         select m_year,m_month,bank_code,acc_code,amt from a02 ");
        sqlCmd.append("                         where m_year = ? and m_month=12 ");
        paramList.add(S_YEAR) ;
        sqlCmd.append("                         and acc_code in ('990410','990420') ");
        sqlCmd.append("                        )a01 ");
        sqlCmd.append("                  )a01, (select * from bn01 where m_year=? ");
        paramList.add(u_year) ;
        if("".equals(bankType)){
            sqlCmd.append("             and bank_type in ('6','7') " );
        }else{
            sqlCmd.append("             and bank_type in (?) " );
            paramList.add(bankType) ;
        }
        sqlCmd.append("                            and bn01.bn_type <> '2')bn01 "); 
        sqlCmd.append("             where a01.bank_code = bn01.bank_no ");             
        sqlCmd.append("             group by a01.m_year,bn01.bank_type,YEAR_TYPE ");
        sqlCmd.append("        ) a01 ");                
        sqlCmd.append("    )a01 ");        
        List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,
                "field_rate1,field_DC_RATE,field_rate2,field_rate3,field_rate4,field_rate5," +
                "field_rate6,field_rate7,field_rate8,field_rate9,field_rate10,field_rate11,field_310000");
        System.out.println("dbData.size=" + dbData.size());
        
        int rowNo=3;
        int cellNo=2;
        if(dbData.size()>0){
            for(int i=0;i<dbData.size();i++){  
                bean = (DataObject)dbData.get(i);
                field_rate1 = (bean.getValue("field_rate1")==null)?"":(bean.getValue("field_rate1")).toString();
                field_DC_RATE = (bean.getValue("field_dc_rate")==null)?"":(bean.getValue("field_dc_rate")).toString();
                field_rate2 = (bean.getValue("field_rate2")==null)?"":(bean.getValue("field_rate2")).toString();
                field_rate3 = (bean.getValue("field_rate3")==null)?"":(bean.getValue("field_rate3")).toString();
                field_rate4 = (bean.getValue("field_rate4")==null)?"":(bean.getValue("field_rate4")).toString();
                field_rate5 = (bean.getValue("field_rate5")==null)?"":(bean.getValue("field_rate5")).toString();
                field_rate6 = (bean.getValue("field_rate6")==null)?"":(bean.getValue("field_rate6")).toString();
                field_rate7 = (bean.getValue("field_rate7")==null)?"":(bean.getValue("field_rate7")).toString();
                field_rate8 = (bean.getValue("field_rate8")==null)?"":(bean.getValue("field_rate8")).toString();
                field_rate9 = (bean.getValue("field_rate9")==null)?"":(bean.getValue("field_rate9")).toString();
                field_rate10 = (bean.getValue("field_rate10")==null)?"":(bean.getValue("field_rate10")).toString();
                field_rate11 = (bean.getValue("field_rate11")==null)?"":(bean.getValue("field_rate11")).toString();
                field_310000 = (bean.getValue("field_310000")==null)?"":(bean.getValue("field_310000")).toString();
                //流動比率  
                row = sheet.getRow(rowNo);
                cell = row.getCell((short)cellNo);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
                cell.setCellValue(MakesUpZero(fmtMicrometer(field_rate1)));
                //存放比率
                row = sheet.getRow(rowNo+1);
                cell = row.getCell((short)cellNo);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16); 
                cell.setCellValue(MakesUpZero(fmtMicrometer(field_DC_RATE)));
                //存款平均利率
                row = sheet.getRow(rowNo+2);
                cell = row.getCell((short)cellNo);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16); 
                cell.setCellValue(MakesUpZero(fmtMicrometer(field_rate2)));
                //放款平均利率
                row = sheet.getRow(rowNo+3);
                cell = row.getCell((short)cellNo);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16); 
                cell.setCellValue(MakesUpZero(fmtMicrometer(field_rate3)));
                //融通資金比率
                row = sheet.getRow(rowNo+4);
                cell = row.getCell((short)cellNo);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16); 
                cell.setCellValue(MakesUpZero(fmtMicrometer(field_rate4)));
                //贊助會員放款占 贊助會員存款比率
                row = sheet.getRow(rowNo+5);
                cell = row.getCell((short)cellNo);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellValue(MakesUpZero(fmtMicrometer(field_rate5)));
                //固定資產淨額佔淨值比率
                row = sheet.getRow(rowNo);
                cell = row.getCell((short)(cellNo+4));
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellValue(MakesUpZero(fmtMicrometer(field_rate6)));
                //淨值佔負債總額比率
                row = sheet.getRow(rowNo+1);
                cell = row.getCell((short)(cellNo+4));
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellValue(MakesUpZero(fmtMicrometer(field_rate7)));
                //淨值佔存款總額比率
                row = sheet.getRow(rowNo+2);
                cell = row.getCell((short)(cellNo+4));
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellValue(MakesUpZero(fmtMicrometer(field_rate8)));
                //內部融資占淨值比率
                row = sheet.getRow(rowNo+3);
                cell = row.getCell((short)(cellNo+4));
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellValue(MakesUpZero(fmtMicrometer(field_rate9)));
                //總資金平均成本率
                row = sheet.getRow(rowNo+4);
                cell = row.getCell((short)(cellNo+4));
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellValue(MakesUpZero(fmtMicrometer(field_rate10)));
                //總資金平均收益率
                row = sheet.getRow(rowNo+5);
                cell = row.getCell((short)(cellNo+4));
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellValue(MakesUpZero(fmtMicrometer(field_rate11)));
                //淨值
                row = sheet.getRow(rowNo+6);
                cell = row.getCell((short)(cellNo+4));
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                cell.setCellValue(fmtMicrometer(field_310000));
            }
        }else{
            row = sheet.getRow(0);
            cell = row.getCell( (short) 0);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16); //設定這個儲存格的字串要儲存雙位元,中文才能正常顯示
            cell.setCellValue(S_YEAR + title +" 查無資料");
        }
      }
      
      FileOutputStream fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "某年底全體農漁會信用部各類財務比率.xls");
      HSSFFooter footer = sheet.getFooter();
      footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
      footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));

      wb.write(fout);
      //儲存
      fout.close();
    }catch (Exception e) {
      System.out.println("//RptAN010W createRpt() Have Error.....");
      e.printStackTrace();
      System.out.println(e.toString());
      System.out.println("//-------------------------------------");
    }

    if(debug) System.out.println("RptAN010W End ...");
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
     if("0".equals(str)){
         str="0.00";
     }
     return str;
 }
}
