/*
 * 97.11.06 create 農漁會信用部逾期放款統計表_總表 by 2295
 * 98.01.09 fix 調整縮放大小 by 2295
 * 99.11.10 fix sql injection & 排序 by 2808
 * 101.08.06 fix sql 報表欄位 by 2968
 * 106.03.07 fix 調整原台灣省改為其他(包含台灣省及福建省.中華民國農會) by 2295  
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

public class RptFR053W {	 
	 public static String createRpt(String m_year,String m_month,String bank_type,String unit) {    

	    String errMsg = "";	    
	    StringBuffer sql = new StringBuffer() ;
	    String condition = "";
	    List dbData = null;	    
	    int rowNum=6;
	    DataObject bean = null;
	    DataObject bean_sub = null;
        DataObject bean_sub1 = null;
	    reportUtil reportUtil = new reportUtil();
		HSSFCellStyle cs_right = null; 
		HSSFCellStyle cs_center = null;
	    HSSFCellStyle cs_left = null;
	    String field_seq="";
	    String hsien_id="";
	    String hsien_name="";//--縣市名稱
	    String bank_code="";
	    String bank_name="";//--機構名稱
	    String amt840750="";//--本月底放款總額
	    String amt840740="";//--狹義逾期放款.逾放金額
	    String amt840740_1="";//--狹義逾期放款.較上月底增減金額		   
	    String field_over_rate_1="";//--狹義逾期放款.逾放比率
	    String field_over_rate_1b="";//--狹義逾期放款.逾放比率(上月)
		String field_over_rate_1c="";//--狹義逾期放款.逾放比率(較上月底增減比率)
		String amt840740_rate="";//--較上月底增減比率
		String amt840740_840760="";//--廣義逾期放款.逾放金額=狹義逾放金額+應予觀察放款
		String amt840740_840760_1="";//--廣義逾期放款.較上月底增減金額		
		String field_over_rate_2="";//--廣義逾期放款.逾放比率
		String field_over_rate_2b="";//--廣義逾期放款.逾放比率(上月)
		String field_over_rate_2c="";//--廣義逾期放款.逾放比率(較上月底增減比率)
		String amt840710="";//--(1)放款本金未超過清償期三個月，惟利息未按期繳納超過三個月至六個月者 
		String amt840720="";//--(2)中長期分期償還放款，未按期攤還超過三個月至六個月者 
		String amt840731="";//--(a)協議分期償還放款，協議條件符合規定，且借款戶依協議條件按期履約未滿六個月者
		String amt840732="";//--(b)已獲信用保證基金同意理賠款項或有足額存單或存款備償(須辦妥質權設定且徵得發單行拋棄抵銷權同意書)，而約定待其他債務人財產處分後再予沖償者
		String amt840733="";//--(c)已確定分配之債權，惟尚未接獲分配款者 
		String amt840734="";//--(d)債務人兼擔保品提供人死亡，於辦理繼承期間屆期而未清償之放款 
		String amt840735="";//--(e)其他  
		String amt840730="";//--小計(a)+(b)+(c)+(d)+(e)
		String amt8407X0="";//--總計(1)+(2)+(3)
		String field_over_rate_3="";//--占同日放款總餘額比率 
	    String last_year="";
	    String last_month="";
	    String unit_name="";
	    List qList = new ArrayList() ;
	    String u_year = "100" ;
	    if(100 > Integer.parseInt(m_year)) {
	    	u_year= "99" ;
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
	      finput = new FileInputStream(xlsDir +System.getProperty("file.separator") +"農漁會信用部逾期放款統計表_總表.xls");

	      //設定FileINputStream讀取Excel檔
	      POIFSFileSystem fs = new POIFSFileSystem(finput);
	      HSSFWorkbook wb = new HSSFWorkbook(fs);
	      HSSFSheet sheet = wb.getSheetAt(0); //讀取第一個工作表，宣告其為sheet
	      HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
	      //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	      //sheet.setAutobreaks(true); //自動分頁

	      //設定頁面符合列印大小
	      sheet.setAutobreaks(false);
	      ps.setScale( (short) 64); //列印縮放百分比

	      ps.setPaperSize( (short) 9); //設定紙張大小 A4
	      //wb.setSheetName(0,"test");
	      finput.close();

	      HSSFRow row = null; //宣告一列
	      HSSFCell cell = null; //宣告一個儲存格
	      
	      cs_right = reportUtil.getRightStyle(wb);
	      cs_center = reportUtil.getDefaultStyle(wb);
	      cs_left = reportUtil.getLeftStyle(wb);
	      m_year = String.valueOf(Integer.parseInt(m_year));
	      m_month = String.valueOf(Integer.parseInt(m_month));
	      if(m_month.equals("1")){//若本月為1月份是..則是上個年度的12月份
	        last_year = String.valueOf(Integer.parseInt(m_year) - 1);
	        last_month = "12";
	      }else{  
	      	last_year = String.valueOf(Integer.parseInt(m_year));
	        last_month = String.valueOf(Integer.parseInt(m_month) - 1);//上個月份的
	      }	   
	      unit_name = Utility.getUnitName(unit);//取得單位名稱
	      sql.append(" SELECT t2.hsien_id, ");
	      sql.append("        t2.hsien_name, "); //--縣市名稱
	      sql.append("        t2.fr001w_output_order, ");
	      sql.append("        bank_code, "); //--機構代號
	      sql.append("        T1.bank_name, "); //--機構名稱
	      sql.append("        1  AS  count_seq, 'A01'  as  field_seq, ");
	      sql.append("        amt840750, "); //--本月底放款總額
	      sql.append("        amt840740, "); //--狹義逾期放款.逾放金額
	      sql.append("        amt840740_1, "); //--狹義逾期放款.較上月底增減金額
	      sql.append("        field_over_rate_1, "); //--狹義逾期放款.逾放比率
	      sql.append("        field_over_rate_1b, ");
	      sql.append("        field_over_rate_1c, "); //--狹義逾期放款.逾放比率(較上月底增減比率)
	      sql.append("        amt840740_rate, ");
	      sql.append("        amt840740_840760, "); //--廣義逾期放款.逾放金額=狹義逾放金額+應予觀察放款
	      sql.append("        amt840740_840760_1, "); //--廣義逾期放款.較上月底增減金額
	      sql.append("        field_over_rate_2, "); //--廣義逾期放款.逾放比率
	      sql.append("        field_over_rate_2b, ");
	      sql.append("        field_over_rate_2c, "); //--廣義逾期放款.逾放比率(較上月底增減比率)
	      sql.append("        amt840710, "); //--(1)放款本金未超過清償期三個月，惟利息未按期繳納超過三個月至六個月者 
	      sql.append("        amt840720, "); //--(2)中長期分期償還放款，未按期攤還超過三個月至六個月者 
	      sql.append("        amt840731, "); //--(a)協議分期償還放款，協議條件符合規定，且借款戶依協議條件按期履約未滿六個月者
	      sql.append("        amt840732, "); //--(b)已獲信用保證基金同意理賠款項或有足額存單或存款備償(須辦妥質權設定且徵得發單行拋棄抵銷權同意書)，而約定待其他債務人財產處分後再予沖償者
	      sql.append("        amt840733, "); //--(c)已確定分配之債權，惟尚未接獲分配款者 
	      sql.append("        amt840734, "); //--(d)債務人兼擔保品提供人死亡，於辦理繼承期間屆期而未清償之放款 
	      sql.append("        amt840735, "); //--(e)其他  
	      sql.append("        amt840730, "); //--小計(a)+(b)+(c)+(d)+(e)
	      sql.append("        amt8407X0, "); //--總計(1)+(2)+(3)
	      sql.append("        field_over_rate_3 "); //--占同日放款總餘額比率      
	      sql.append(" FROM (SELECT a.m_year,a.m_month,a.bank_code,a.bank_name, ");
	      sql.append("              ROUND (a.amt840750 /?, 0) AS amt840750, ");
	      sql.append("              ROUND (a.amt840740 /?, 0) AS amt840740, ");
	      sql.append("              ROUND ( (a.amt840740 - b.amt840740) /?, 0) AS amt840740_1, ");
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      sql.append("              DECODE (a.amt840750, 0, 0, ROUND (a.amt840740 / a.amt840750 * 100, 2)) AS field_OVER_RATE_1, ");
	      sql.append("              DECODE (b.amt840750, 0, 0, ROUND (b.amt840740 / b.amt840750 * 100, 2)) AS field_OVER_RATE_1b, ");
	      sql.append("              DECODE (a.amt840750, 0, 0, ROUND (a.amt840740 / a.amt840750 * 100, 2)) ");
	      sql.append("              - DECODE (b.amt840750,0, 0,ROUND (b.amt840740 / b.amt840750 * 100, 2)) ");
	      sql.append("                 AS field_OVER_RATE_1c, ");
	      sql.append("              DECODE ( a.amt840740, 0, 0, ROUND ( (a.amt840740 - b.amt840740) / a.amt840740 * 100, 2)) ");
	      sql.append("                 AS amt840740_rate, ");
	      sql.append("              ROUND (a.amt840740_840760 /?, 0) AS amt840740_840760, ");
	      sql.append("              ROUND ( (a.amt840740_840760 - b.amt840740_840760) /?, 0) AS amt840740_840760_1, ");
	      qList.add(unit) ;
	      qList.add(unit) ;
	      sql.append("              DECODE (a.amt840750, 0, 0,  ROUND (a.amt840740_840760 / a.amt840750 * 100, 2))  AS field_OVER_RATE_2, ");
	      sql.append("              DECODE (b.amt840750, 0, 0,  ROUND (b.amt840740_840760 / b.amt840750 * 100, 2))  AS field_OVER_RATE_2b, ");
	      sql.append("              DECODE (a.amt840750, 0, 0,  ROUND (a.amt840740_840760 / a.amt840750 * 100, 2)) ");
	      sql.append("              - DECODE (b.amt840750, 0, 0, ROUND (b.amt840740_840760 / b.amt840750 * 100, 2)) AS field_OVER_RATE_2c, ");
	      sql.append("              ROUND (a.amt840710 / ?, 0) AS amt840710, ");
	      sql.append("              ROUND (a.amt840720 / ?, 0) AS amt840720, ");
	      sql.append("              ROUND (a.amt840731 / ?, 0) AS amt840731, ");
	      sql.append("              ROUND (a.amt840732 / ?, 0) AS amt840732, ");
	      sql.append("              ROUND (a.amt840733 / ?, 0) AS amt840733, ");
	      sql.append("              ROUND (a.amt840734 / ?, 0) AS amt840734, ");
	      sql.append("              ROUND (a.amt840735 / ?, 0) AS amt840735, ");
	      sql.append("              ROUND (a.amt840730 / ?, 0) AS amt840730, ");
	      sql.append("              ROUND (a.amt8407X0 / ?, 0) AS amt8407X0, ");
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      sql.append("              DECODE (a.amt840750, 0, 0,  ROUND (a.amt8407X0 / a.amt840750 * 100, 2))  AS field_OVER_RATE_3 ");
	      sql.append("         FROM ( SELECT a04.m_year,  a04.m_month,bank_code, bn01.bank_name, ");
	      sql.append("                       SUM (DECODE (acc_code, '840750', amt, 0)) AS amt840750, ");
	      sql.append("                       SUM (DECODE (acc_code, '840740', amt, 0)) AS amt840740, ");
	      sql.append("                       SUM (DECODE (acc_code, '840740', amt,'840760', amt,0)) AS amt840740_840760, ");
	      sql.append("                       SUM (DECODE (acc_code, '840710', amt, 0)) AS amt840710, ");
	      sql.append("                       SUM (DECODE (acc_code, '840720', amt, 0)) AS amt840720, ");
	      sql.append("                       SUM (DECODE (acc_code, '840731', amt, 0)) AS amt840731, ");
	      sql.append("                       SUM (DECODE (acc_code, '840732', amt, 0)) AS amt840732, ");
	      sql.append("                       SUM (DECODE (acc_code, '840733', amt, 0)) AS amt840733, ");
	      sql.append("                       SUM (DECODE (acc_code, '840734', amt, 0)) AS amt840734, ");
	      sql.append("                       SUM (DECODE (acc_code, '840735', amt, 0)) AS amt840735, ");
	      sql.append("                       SUM (DECODE (acc_code,'840731', amt,'840732', amt,'840733', amt,'840734', amt,'840735', amt,0)) AS amt840730, ");
	      sql.append("                       SUM (DECODE (acc_code,'840710', amt,'840720', amt,'840731', amt,'840732', amt,'840733', amt,'840734',amt,'840735', amt,0)) AS amt8407X0 ");
	      sql.append("                FROM a04,(SELECT * FROM bn01 WHERE m_year = ? AND bank_type = ? ) bn01 ");
	      sql.append("                WHERE  a04.m_year = ? AND m_month = ? AND a04.bank_code = bn01.bank_no ");
	      qList.add(u_year) ;
          qList.add(bank_type) ;
	      qList.add(m_year) ;
	      qList.add(m_month) ;
	      sql.append("                GROUP BY a04.m_year,m_month,bank_type,bank_code,bn01.bank_name ");
	      sql.append("                ORDER BY a04.m_year,m_month,bank_type,bank_code,bn01.bank_name) a ");
	      sql.append("          LEFT JOIN ");
	      sql.append("                 ( SELECT a04.m_year,m_month,bank_code,bn01.bank_name, ");
	      sql.append("                          SUM (DECODE (acc_code, '840750', amt, 0))AS amt840750, ");
	      sql.append("                          SUM (DECODE (acc_code, '840740', amt, 0))AS amt840740, ");
	      sql.append("                          SUM (DECODE (acc_code,'840740', amt,'840760', amt,0)) AS amt840740_840760 ");
	      sql.append("                   FROM a04,(SELECT * FROM bn01 WHERE m_year = ? AND bank_type = ? ) bn01 ");
	      sql.append("                   WHERE  a04.m_year = ? AND m_month = ? ");
	      qList.add(u_year) ;
	      qList.add(bank_type) ;
	      qList.add(last_year) ;
	      qList.add(last_month) ;
	      sql.append("                   AND bn01.bank_no = a04.bank_code ");
	      sql.append("                  GROUP BY a04.m_year,m_month,bank_type,bank_code,bn01.bank_name ");
	      sql.append("                  ORDER BY a04.m_year,m_month,bank_type,bank_code,bn01.bank_name) b ON a.bank_code = b.bank_code ");
	      sql.append("      ) t1,(select * from v_bank_location where m_year = ?)T2 ");
	      qList.add(u_year) ;
	      sql.append(" WHERE t1.bank_code = T2.bank_no ");
	      //sql.append(" --ORDER BY T2.fr001w_output_order, hsien_id ,  bank_code  ");
	      sql.append(" union  ");
	      //總計
	      sql.append(" SELECT  ' '  AS  hsien_id ,  ' 總   計 '   AS hsien_name,  '001' AS FR001W_output_order, ' ' AS  bank_code , ' '   AS  BANK_NAME,         "); 
	      sql.append("         COUNT(*)  AS  COUNT_SEQ, 'A99'  as  field_SEQ,      ");
	      sql.append("         ROUND (SUM(a.amt840750) / ?, 0) AS amt840750, ");
	      sql.append("         ROUND (SUM(a.amt840740) / ?, 0) AS amt840740, ");
	      qList.add(unit) ;
	      qList.add(unit) ;
	      sql.append("         ROUND ( (SUM(a.amt840740) - SUM(b.amt840740)) / ?, 0) AS amt840740_1, ");
	      qList.add(unit) ;
	      sql.append("         DECODE (SUM(a.amt840750), 0, 0, ROUND (SUM(a.amt840740) / SUM(a.amt840750) * 100, 2)) AS field_OVER_RATE_1, ");
	      sql.append("         DECODE (SUM(b.amt840750), 0, 0, ROUND (SUM(b.amt840740) / SUM(b.amt840750) * 100, 2)) AS field_OVER_RATE_1b, ");
	      sql.append("         DECODE (SUM(a.amt840750), 0, 0, ROUND (SUM(a.amt840740) / SUM(a.amt840750) * 100, 2)) ");
	      sql.append("         - DECODE (SUM(b.amt840750),0, 0,ROUND (SUM(b.amt840740) / SUM(b.amt840750) * 100, 2)) ");
	      sql.append("            AS field_OVER_RATE_1c, ");
	      sql.append("         DECODE ( SUM(a.amt840740), 0, 0, ROUND ( (SUM(a.amt840740) - SUM(b.amt840740)) / SUM(a.amt840740) * 100, 2)) ");
	      sql.append("            AS amt840740_rate, ");
	      sql.append("         ROUND (SUM(a.amt840740_840760) /?, 0) AS amt840740_840760, ");
	      sql.append("         ROUND ( (SUM(a.amt840740_840760) - SUM(b.amt840740_840760)) / ?, 0) AS amt840740_840760_1, ");
	      qList.add(unit) ;
	      qList.add(unit) ;
	      sql.append("         DECODE (SUM(a.amt840750), 0, 0,  ROUND (SUM(a.amt840740_840760) / SUM(a.amt840750) * 100, 2))  AS field_OVER_RATE_2, ");
	      sql.append("         DECODE (SUM(b.amt840750), 0, 0,  ROUND (SUM(b.amt840740_840760) / SUM(b.amt840750) * 100, 2))  AS field_OVER_RATE_2b, ");
	      sql.append("         DECODE (SUM(a.amt840750), 0, 0,  ROUND (SUM(a.amt840740_840760) / SUM(a.amt840750) * 100, 2)) ");
	      sql.append("         - DECODE (SUM(b.amt840750), 0, 0, ROUND (SUM(b.amt840740_840760) / SUM(b.amt840750) * 100, 2)) AS field_OVER_RATE_2c, ");
	      sql.append("         ROUND (SUM(a.amt840710) / ?, 0) AS amt840710, ");
	      sql.append("         ROUND (SUM(a.amt840720) / ?, 0) AS amt840720, ");
	      sql.append("         ROUND (SUM(a.amt840731) / ?, 0) AS amt840731, ");
	      sql.append("         ROUND (SUM(a.amt840732) / ?, 0) AS amt840732, ");
	      sql.append("         ROUND (SUM(a.amt840733) / ?, 0) AS amt840733, ");
	      sql.append("         ROUND (SUM(a.amt840734) / ?, 0) AS amt840734, ");
	      sql.append("         ROUND (SUM(a.amt840735) / ?, 0) AS amt840735, ");
	      sql.append("         ROUND (SUM(a.amt840730) / ?, 0) AS amt840730, ");
	      sql.append("         ROUND (SUM(a.amt8407X0) / ?, 0) AS amt8407X0, ");
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      sql.append("         DECODE (SUM(a.amt840750), 0, 0,  ROUND (SUM(a.amt8407X0) / SUM(a.amt840750) * 100, 2))  AS field_OVER_RATE_3 ");
	      sql.append(" FROM ( SELECT a04.m_year,  a04.m_month,bank_code, bn01.bank_name, ");
	      sql.append("               SUM (DECODE (acc_code, '840750', amt, 0)) AS amt840750, ");
	      sql.append("               SUM (DECODE (acc_code, '840740', amt, 0)) AS amt840740, ");
	      sql.append("               SUM (DECODE (acc_code, '840740', amt,'840760', amt,0)) AS amt840740_840760, ");
	      sql.append("               SUM (DECODE (acc_code, '840710', amt, 0)) AS amt840710, ");
	      sql.append("               SUM (DECODE (acc_code, '840720', amt, 0)) AS amt840720, ");
	      sql.append("               SUM (DECODE (acc_code, '840731', amt, 0)) AS amt840731, ");
	      sql.append("               SUM (DECODE (acc_code, '840732', amt, 0)) AS amt840732, ");
	      sql.append("               SUM (DECODE (acc_code, '840733', amt, 0)) AS amt840733, ");
	      sql.append("               SUM (DECODE (acc_code, '840734', amt, 0)) AS amt840734, ");
	      sql.append("               SUM (DECODE (acc_code, '840735', amt, 0)) AS amt840735, ");
	      sql.append("               SUM (DECODE (acc_code,'840731', amt,'840732', amt,'840733', amt,'840734', amt,'840735', amt,0)) AS amt840730, ");
	      sql.append("               SUM (DECODE (acc_code,'840710', amt,'840720', amt,'840731', amt,'840732', amt,'840733', amt,'840734',amt,'840735', amt,0)) AS amt8407X0 ");
	      sql.append("        FROM a04,(SELECT * FROM bn01 WHERE m_year = ? AND bank_type = ?) bn01 ");
	      sql.append("        WHERE  a04.m_year = ? AND m_month = ? AND a04.bank_code = bn01.bank_no ");
	      qList.add(u_year) ;
	      qList.add(bank_type) ;
	      qList.add(m_year) ;
	      qList.add(m_month) ;
	      sql.append("        GROUP BY a04.m_year,m_month,bank_type,bank_code,bn01.bank_name ");
	      sql.append("        ORDER BY a04.m_year,m_month,bank_type,bank_code,bn01.bank_name) a ");
	      sql.append(" LEFT JOIN ");
	      sql.append("        ( SELECT a04.m_year,m_month,bank_code,bn01.bank_name, ");
	      sql.append("                 SUM (DECODE (acc_code, '840750', amt, 0))AS amt840750, ");
	      sql.append("                 SUM (DECODE (acc_code, '840740', amt, 0))AS amt840740, ");
	      sql.append("                 SUM (DECODE (acc_code,'840740', amt,'840760', amt,0)) AS amt840740_840760 ");
	      sql.append("          FROM a04,(SELECT * FROM bn01 WHERE m_year = ? AND bank_type = ?) bn01 ");
	      sql.append("          WHERE  a04.m_year = ? AND m_month = ? ");
	      qList.add(u_year) ;
	      qList.add(bank_type) ;
	      qList.add(last_year) ;
	      qList.add(last_month) ;
	      sql.append("          AND bn01.bank_no = a04.bank_code ");
	      sql.append("         GROUP BY a04.m_year,m_month,bank_type,bank_code,bn01.bank_name ");
	      sql.append("         ORDER BY a04.m_year,m_month,bank_type,bank_code,bn01.bank_name) b ON a.bank_code = b.bank_code ");
	      sql.append(" LEFT JOIN  (select * from  v_bank_location where m_year=?)T2  on a.bank_code= T2.bank_no    ");
	      qList.add(u_year) ;
	      sql.append(" union  ");
	      //臺灣省合計 //106.03.07 原台灣省改為其他(含台灣省及福建省.中華民國農會)
	      sql.append(" SELECT  ' '  AS  hsien_id ,  '其他 '   AS hsien_name,  '025' AS FR001W_output_order, ' ' AS  bank_code , ' '   AS  BANK_NAME,   ");       
	      sql.append("         COUNT(*)  AS  COUNT_SEQ, 'A92'  as  field_SEQ,      ");
	      sql.append("         ROUND (SUM(a.amt840750) / ?, 0) AS amt840750, ");
	      sql.append("         ROUND (SUM(a.amt840740) / ?, 0) AS amt840740, ");
	      sql.append("         ROUND ( (SUM(a.amt840740) - SUM(b.amt840740)) / ?, 0) AS amt840740_1, ");
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      sql.append("         DECODE (SUM(a.amt840750), 0, 0, ROUND (SUM(a.amt840740) / SUM(a.amt840750) * 100, 2)) AS field_OVER_RATE_1, ");
	      sql.append("         DECODE (SUM(b.amt840750), 0, 0, ROUND (SUM(b.amt840740) / SUM(b.amt840750) * 100, 2)) AS field_OVER_RATE_1b, ");
	      sql.append("         DECODE (SUM(a.amt840750), 0, 0, ROUND (SUM(a.amt840740) / SUM(a.amt840750) * 100, 2)) ");
	      sql.append("         - DECODE (SUM(b.amt840750),0, 0,ROUND (SUM(b.amt840740) / SUM(b.amt840750) * 100, 2)) ");
	      sql.append("            AS field_OVER_RATE_1c, ");
	      sql.append("         DECODE ( SUM(a.amt840740), 0, 0, ROUND ( (SUM(a.amt840740) - SUM(b.amt840740)) / SUM(a.amt840740) * 100, 2)) ");
	      sql.append("            AS amt840740_rate, ");
	      sql.append("         ROUND (SUM(a.amt840740_840760) / ?, 0) AS amt840740_840760, ");
	      sql.append("         ROUND ( (SUM(a.amt840740_840760) - SUM(b.amt840740_840760)) / ?, 0) AS amt840740_840760_1, ");
	      qList.add(unit) ;
	      qList.add(unit) ;
	      sql.append("         DECODE (SUM(a.amt840750), 0, 0,  ROUND (SUM(a.amt840740_840760) / SUM(a.amt840750) * 100, 2))  AS field_OVER_RATE_2, ");
	      sql.append("         DECODE (SUM(b.amt840750), 0, 0,  ROUND (SUM(b.amt840740_840760) / SUM(b.amt840750) * 100, 2))  AS field_OVER_RATE_2b, ");
	      sql.append("         DECODE (SUM(a.amt840750), 0, 0,  ROUND (SUM(a.amt840740_840760) / SUM(a.amt840750) * 100, 2)) ");
	      sql.append("         - DECODE (SUM(b.amt840750), 0, 0, ROUND (SUM(b.amt840740_840760) / SUM(b.amt840750) * 100, 2)) AS field_OVER_RATE_2c, ");
	      sql.append("         ROUND (SUM(a.amt840710) / ?, 0) AS amt840710, ");
	      sql.append("         ROUND (SUM(a.amt840720) / ?, 0) AS amt840720, ");
	      sql.append("         ROUND (SUM(a.amt840731) / ?, 0) AS amt840731, ");
	      sql.append("         ROUND (SUM(a.amt840732) / ?, 0) AS amt840732, ");
	      sql.append("         ROUND (SUM(a.amt840733) / ?, 0) AS amt840733, ");
	      sql.append("         ROUND (SUM(a.amt840734) / ?, 0) AS amt840734, ");
	      sql.append("         ROUND (SUM(a.amt840735) / ?, 0) AS amt840735, ");
	      sql.append("         ROUND (SUM(a.amt840730) / ?, 0) AS amt840730, ");
	      sql.append("         ROUND (SUM(a.amt8407X0) / ?, 0) AS amt8407X0, ");
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      sql.append("         DECODE (SUM(a.amt840750), 0, 0,  ROUND (SUM(a.amt8407X0) / SUM(a.amt840750) * 100, 2))  AS field_OVER_RATE_3 ");
	      sql.append(" FROM ( SELECT a04.m_year,  a04.m_month,bank_code, bn01.bank_name, ");
	      sql.append("               SUM (DECODE (acc_code, '840750', amt, 0)) AS amt840750, ");
	      sql.append("               SUM (DECODE (acc_code, '840740', amt, 0)) AS amt840740, ");
	      sql.append("               SUM (DECODE (acc_code, '840740', amt,'840760', amt,0)) AS amt840740_840760, ");
	      sql.append("               SUM (DECODE (acc_code, '840710', amt, 0)) AS amt840710, ");
	      sql.append("               SUM (DECODE (acc_code, '840720', amt, 0)) AS amt840720, ");
	      sql.append("               SUM (DECODE (acc_code, '840731', amt, 0)) AS amt840731, ");
	      sql.append("               SUM (DECODE (acc_code, '840732', amt, 0)) AS amt840732, ");
	      sql.append("               SUM (DECODE (acc_code, '840733', amt, 0)) AS amt840733, ");
	      sql.append("               SUM (DECODE (acc_code, '840734', amt, 0)) AS amt840734, ");
	      sql.append("               SUM (DECODE (acc_code, '840735', amt, 0)) AS amt840735, ");
	      sql.append("               SUM (DECODE (acc_code,'840731', amt,'840732', amt,'840733', amt,'840734', amt,'840735', amt,0)) AS amt840730, ");
	      sql.append("               SUM (DECODE (acc_code,'840710', amt,'840720', amt,'840731', amt,'840732', amt,'840733', amt,'840734',amt,'840735', amt,0)) AS amt8407X0 ");
	      sql.append("        FROM a04,(SELECT * FROM bn01 WHERE m_year = ? AND bank_type = ?) bn01 ");
	      sql.append("        WHERE  a04.m_year = ? AND m_month = ? AND a04.bank_code = bn01.bank_no ");
	      qList.add(u_year) ;
	      qList.add(bank_type) ;
	      qList.add(m_year) ;
	      qList.add(m_month) ;
	      sql.append("        GROUP BY a04.m_year,m_month,bank_type,bank_code,bn01.bank_name ");
	      sql.append("        ORDER BY a04.m_year,m_month,bank_type,bank_code,bn01.bank_name) a ");
	      sql.append(" LEFT JOIN ");
	      sql.append("        ( SELECT a04.m_year,m_month,bank_code,bn01.bank_name, ");
	      sql.append("                 SUM (DECODE (acc_code, '840750', amt, 0))AS amt840750, ");
	      sql.append("                 SUM (DECODE (acc_code, '840740', amt, 0))AS amt840740, ");
	      sql.append("                 SUM (DECODE (acc_code,'840740', amt,'840760', amt,0)) AS amt840740_840760 ");
	      sql.append("          FROM a04,(SELECT * FROM bn01 WHERE m_year = ? AND bank_type = ?) bn01 ");
	      sql.append("          WHERE  a04.m_year = ? AND m_month = ? ");
	      qList.add(u_year) ;
	      qList.add(bank_type) ;
	      qList.add(last_year) ;
	      qList.add(last_month) ;
	      sql.append("          AND bn01.bank_no = a04.bank_code ");
	      sql.append("         GROUP BY a04.m_year,m_month,bank_type,bank_code,bn01.bank_name ");
	      sql.append("         ORDER BY a04.m_year,m_month,bank_type,bank_code,bn01.bank_name) b ON a.bank_code = b.bank_code ");
	      sql.append(" LEFT JOIN   ");
	      sql.append("        ( ");
	      sql.append("         select  v.*,cd01.Hsien_div from  (select * from v_bank_location where m_year = ?)v   ");
	      sql.append("         LEFT JOIN (select * from  cd01 where cd01.hsien_id <> 'Y') cd01 on v.hsien_id = cd01.hsien_id ");	      
	      qList.add(u_year) ;
	      sql.append("         )T2  on a.bank_code= T2.bank_no    ");
	      sql.append(" where T2.hsien_div in ('2','3')  ");//106.03.08 fix
	      sql.append(" union              ");
	      //各縣市小計
	      sql.append(" SELECT  T2.hsien_id,T2.hsien_name,T2.FR001W_output_order,'' as bank_no,'' as bank_name, ");
	      sql.append(" COUNT(*)  AS  COUNT_SEQ,'A90'  as  field_SEQ, ");
	      sql.append("         ROUND (SUM(a.amt840750) / ?, 0) AS amt840750, ");
	      sql.append("         ROUND (SUM(a.amt840740) / ?, 0) AS amt840740, ");
	      sql.append("         ROUND ( (SUM(a.amt840740) - SUM(b.amt840740)) / ?, 0) AS amt840740_1, ");
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      sql.append("         DECODE (SUM(a.amt840750), 0, 0, ROUND (SUM(a.amt840740) / SUM(a.amt840750) * 100, 2)) AS field_OVER_RATE_1, ");
	      sql.append("         DECODE (SUM(b.amt840750), 0, 0, ROUND (SUM(b.amt840740) / SUM(b.amt840750) * 100, 2)) AS field_OVER_RATE_1b, ");
	      sql.append("         DECODE (SUM(a.amt840750), 0, 0, ROUND (SUM(a.amt840740) / SUM(a.amt840750) * 100, 2)) ");
	      sql.append("         - DECODE (SUM(b.amt840750),0, 0,ROUND (SUM(b.amt840740) / SUM(b.amt840750) * 100, 2)) ");
	      sql.append("            AS field_OVER_RATE_1c, ");
	      sql.append("         DECODE ( SUM(a.amt840740), 0, 0, ROUND ( (SUM(a.amt840740) - SUM(b.amt840740)) / SUM(a.amt840740) * 100, 2)) ");
	      sql.append("            AS amt840740_rate, ");
	      sql.append("         ROUND (SUM(a.amt840740_840760) / ?, 0) AS amt840740_840760, ");
	      sql.append("         ROUND ( (SUM(a.amt840740_840760) - SUM(b.amt840740_840760)) / ?, 0) AS amt840740_840760_1, ");
	      qList.add(unit) ;
	      qList.add(unit) ;
	      sql.append("         DECODE (SUM(a.amt840750), 0, 0,  ROUND (SUM(a.amt840740_840760) / SUM(a.amt840750) * 100, 2))  AS field_OVER_RATE_2, ");
	      sql.append("         DECODE (SUM(b.amt840750), 0, 0,  ROUND (SUM(b.amt840740_840760) / SUM(b.amt840750) * 100, 2))  AS field_OVER_RATE_2b, ");
	      sql.append("         DECODE (SUM(a.amt840750), 0, 0,  ROUND (SUM(a.amt840740_840760) / SUM(a.amt840750) * 100, 2)) ");
	      sql.append("         - DECODE (SUM(b.amt840750), 0, 0, ROUND (SUM(b.amt840740_840760) / SUM(b.amt840750) * 100, 2)) AS field_OVER_RATE_2c, ");
	      sql.append("         ROUND (SUM(a.amt840710) / ?, 0) AS amt840710, ");
	      sql.append("         ROUND (SUM(a.amt840720) / ?, 0) AS amt840720, ");
	      sql.append("         ROUND (SUM(a.amt840731) / ?, 0) AS amt840731, ");
	      sql.append("         ROUND (SUM(a.amt840732) / ?, 0) AS amt840732, ");
	      sql.append("         ROUND (SUM(a.amt840733) / ?, 0) AS amt840733, ");
	      sql.append("         ROUND (SUM(a.amt840734) / ?, 0) AS amt840734, ");
	      sql.append("         ROUND (SUM(a.amt840735) / ?, 0) AS amt840735, ");
	      sql.append("         ROUND (SUM(a.amt840730) / ?, 0) AS amt840730, ");
	      sql.append("         ROUND (SUM(a.amt8407X0) / ?, 0) AS amt8407X0, ");
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      qList.add(unit) ;
	      sql.append("         DECODE (SUM(a.amt840750), 0, 0,  ROUND (SUM(a.amt8407X0) / SUM(a.amt840750) * 100, 2))  AS field_OVER_RATE_3 ");
	      sql.append(" FROM ( SELECT a04.m_year,  a04.m_month,bank_code, bn01.bank_name, ");
	      sql.append("               SUM (DECODE (acc_code, '840750', amt, 0)) AS amt840750, ");
	      sql.append("               SUM (DECODE (acc_code, '840740', amt, 0)) AS amt840740, ");
	      sql.append("               SUM (DECODE (acc_code, '840740', amt,'840760', amt,0)) AS amt840740_840760, ");
	      sql.append("               SUM (DECODE (acc_code, '840710', amt, 0)) AS amt840710, ");
	      sql.append("               SUM (DECODE (acc_code, '840720', amt, 0)) AS amt840720, ");
	      sql.append("               SUM (DECODE (acc_code, '840731', amt, 0)) AS amt840731, ");
	      sql.append("               SUM (DECODE (acc_code, '840732', amt, 0)) AS amt840732, ");
	      sql.append("               SUM (DECODE (acc_code, '840733', amt, 0)) AS amt840733, ");
	      sql.append("               SUM (DECODE (acc_code, '840734', amt, 0)) AS amt840734, ");
	      sql.append("               SUM (DECODE (acc_code, '840735', amt, 0)) AS amt840735, ");
	      sql.append("               SUM (DECODE (acc_code,'840731', amt,'840732', amt,'840733', amt,'840734', amt,'840735', amt,0)) AS amt840730, ");
	      sql.append("               SUM (DECODE (acc_code,'840710', amt,'840720', amt,'840731', amt,'840732', amt,'840733', amt,'840734',amt,'840735', amt,0)) AS amt8407X0 ");
	      sql.append("        FROM a04,(SELECT * FROM bn01 WHERE m_year = ? AND bank_type =  ?) bn01 ");
	      sql.append("        WHERE  a04.m_year = ? AND m_month = ? AND a04.bank_code = bn01.bank_no ");
	      qList.add(u_year) ;
	      qList.add(bank_type) ;
	      qList.add(m_year) ;
	      qList.add(m_month) ;
	      sql.append("        GROUP BY a04.m_year,m_month,bank_type,bank_code,bn01.bank_name ");
	      sql.append("        ORDER BY a04.m_year,m_month,bank_type,bank_code,bn01.bank_name) a ");
	      sql.append(" LEFT JOIN ");
	      sql.append("        ( SELECT a04.m_year,m_month,bank_code,bn01.bank_name, ");
	      sql.append("                 SUM (DECODE (acc_code, '840750', amt, 0))AS amt840750, ");
	      sql.append("                 SUM (DECODE (acc_code, '840740', amt, 0))AS amt840740, ");
	      sql.append("                 SUM (DECODE (acc_code,'840740', amt,'840760', amt,0)) AS amt840740_840760 ");
	      sql.append("          FROM a04,(SELECT * FROM bn01 WHERE m_year = ? AND bank_type = ?) bn01 ");
	      sql.append("          WHERE  a04.m_year = ? AND m_month = ? ");
	      qList.add(u_year) ;
	      qList.add(bank_type) ;
	      qList.add(last_year) ;
	      qList.add(last_month) ;
	      sql.append("          AND bn01.bank_no = a04.bank_code ");
	      sql.append("         GROUP BY a04.m_year,m_month,bank_type,bank_code,bn01.bank_name ");
	      sql.append("         ORDER BY a04.m_year,m_month,bank_type,bank_code,bn01.bank_name) b ON a.bank_code = b.bank_code ");
	      sql.append(" LEFT JOIN  (select * from  v_bank_location where m_year=?)T2  on a.bank_code= T2.bank_no  ");
	      qList.add(u_year) ;
	      sql.append(" GROUP BY T2.hsien_id ,T2.hsien_name,T2.FR001W_output_order ");
	      sql.append(" ORDER by FR001W_output_order,field_SEQ,hsien_id ,bank_code ");
			
	      dbData = DBManager.QueryDB_SQLParam(sql.toString(),qList,"hsien_id,hsien_name,fr001w_output_order,bank_code,bank_name," +
                                                    	      		"count_seq,field_seq,amt840750,amt840740,amt840740_1," +
                                                    	      		"field_over_rate_1,field_over_rate_1b,field_over_rate_1c," +
                                                    	      		"amt840740_rate,amt840740_840760,amt840740_840760_1," +
                                                    	      		"field_over_rate_2,field_over_rate_2b,field_over_rate_2c," +
                                                    	      		"amt840710,amt840720,amt840731,amt840732,amt840733,amt840734," +
                                                    	      		"amt840735,amt840730,amt8407X0,field_over_rate_3");
	      System.out.println("dbData.size=" + dbData.size());
	      //設定報表表頭資料============================================	
	      
	      row=sheet.getRow(0);
	      cell = row.getCell( (short) 0);
	   	  cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	   	  cell.setCellValue((bank_type.equals("6")?"農會":"漁會")+"信用部逾期放款統計表");	   	 
	   	  
	   	  row = sheet.getRow(1);
	   	  cell = row.getCell( (short) 0);	   	  
	   	  cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	      if (dbData != null && dbData.size() != 0) {
		   	  cell.setCellValue("基準日:"+m_year+"年"+m_month+"月底");		   	  
		   	  row = sheet.getRow(2);
		   	  cell = row.getCell( (short) 17);	   	  
		   	  cell.setEncoding(HSSFCell.ENCODING_UTF_16);		   	
		   	  cell.setCellValue("單位:新臺幣 " + unit_name + ",%");  
	                	     	 
		   	  int rowNum2=0;	      	  
	      	  for(int i=0;i<dbData.size();i++){	      	      
	      	      bean = (DataObject)dbData.get(i);
	      	      field_seq=(bean.getValue("field_seq") == null)?"":(String)(bean.getValue("field_seq"));
	      	      hsien_id=(bean.getValue("hsien_id") == null)?"":(bean.getValue("hsien_id")).toString();
	      	      hsien_name=(bean.getValue("hsien_name") == null)?"":(bean.getValue("hsien_name")).toString();
	      	      bank_code=(bean.getValue("bank_code") == null)?"":(bean.getValue("bank_code")).toString();
	      	      bank_name=(bean.getValue("bank_name") == null)?"":(bean.getValue("bank_name")).toString();
	      	      amt840750=(bean.getValue("amt840750") == null)?"":(bean.getValue("amt840750")).toString();//--本月底放款總額
	      	      amt840740=(bean.getValue("amt840740") == null)?"":(bean.getValue("amt840740")).toString();//--廣義逾期放款.逾放金額
	      	      amt840740_1=(bean.getValue("amt840740_1") == null)?"":(bean.getValue("amt840740_1")).toString();//--廣義逾期放款.較上月底增減金額		   
	      	      field_over_rate_1=(bean.getValue("field_over_rate_1") == null)?"":(bean.getValue("field_over_rate_1")).toString();//--狹義逾期放款.逾放比率
	      	      field_over_rate_1b=(bean.getValue("field_over_rate_1b") == null)?"":(bean.getValue("field_over_rate_1b")).toString();//--狹義逾期放款.逾放比率(上月)
	      	      field_over_rate_1c=(bean.getValue("field_over_rate_1c") == null)?"":(bean.getValue("field_over_rate_1c")).toString();//--狹義逾期放款.逾放比率(較上月底增減比率)
	      	      amt840740_rate=(bean.getValue("amt840740_rate") == null)?"":(bean.getValue("amt840740_rate")).toString();//--較上月底增減比率
	      	      amt840740_840760=(bean.getValue("amt840740_840760") == null)?"":(bean.getValue("amt840740_840760")).toString();//--逾放金額=狹義逾放金額+應予觀察放款
	      	      amt840740_840760_1=(bean.getValue("amt840740_840760_1") == null)?"":(bean.getValue("amt840740_840760_1")).toString();//--廣義逾期放款.較上月底增減金額
	      	      field_over_rate_2=(bean.getValue("field_over_rate_2") == null)?"":(bean.getValue("field_over_rate_2")).toString();//--廣義逾期放款.逾放比率
	      	      field_over_rate_2b=(bean.getValue("field_over_rate_2b") == null)?"":(bean.getValue("field_over_rate_2b")).toString();//--廣義逾期放款.逾放比率(上月)
	      	      field_over_rate_2c=(bean.getValue("field_over_rate_2c") == null)?"":(bean.getValue("field_over_rate_2c")).toString();//--廣義逾期放款.逾放比率(較上月底增減比率)
	      	      amt840710=(bean.getValue("amt840710") == null)?"":(bean.getValue("amt840710")).toString();
	      	      amt840720=(bean.getValue("amt840720") == null)?"":(bean.getValue("amt840720")).toString();
	      	      amt840731=(bean.getValue("amt840731") == null)?"":(bean.getValue("amt840731")).toString();
	      	      amt840732=(bean.getValue("amt840732") == null)?"":(bean.getValue("amt840732")).toString();
	      	      amt840733=(bean.getValue("amt840733") == null)?"":(bean.getValue("amt840733")).toString();
	      	      amt840734=(bean.getValue("amt840734") == null)?"":(bean.getValue("amt840734")).toString();
	      	      amt840735=(bean.getValue("amt840735") == null)?"":(bean.getValue("amt840735")).toString();
	      	      amt840730=(bean.getValue("amt840730") == null)?"":(bean.getValue("amt840730")).toString();
	      	      amt8407X0=(bean.getValue("amt8407x0") == null)?"":(bean.getValue("amt8407x0")).toString();
	      	      field_over_rate_3=(bean.getValue("field_over_rate_3") == null)?"":(bean.getValue("field_over_rate_3")).toString();//--占同日放款總餘額比率
	    	  	  
   	    	  	  //列印各機構明細資料
			      if(!"A99".equals(field_seq)&& !"A92".equals(field_seq)){
			         row = sheet.createRow(rowNum);
			    	 cell=row.createCell((short)1);
			    	 cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                     cell.setCellStyle(cs_right);
			    	 cell.setCellValue(Utility.setCommaFormat(amt840750));//本月底放款總額 
			    	 
			    	 cell=row.createCell((short)2);
			    	 cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                     cell.setCellStyle(cs_right);
			    	 cell.setCellValue(Utility.setCommaFormat(amt840740));//狹義逾期放款.逾放金額
			    	 
			    	 cell=row.createCell((short)3);
			    	 cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                     cell.setCellStyle(cs_right);
			    	 cell.setCellValue(Utility.setCommaFormat(amt840740_1));//狹義逾期放款.較上月底增減金額		    	  
			    	 
			    	 cell=row.createCell((short)4);
			    	 cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                     cell.setCellStyle(cs_right);
			    	 cell.setCellValue(Utility.setCommaFormat(field_over_rate_1)+"%");//狹義逾期放款.逾放比率
			    	 
			    	 cell=row.createCell((short)5);
			    	 cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                     cell.setCellStyle(cs_right);
			    	 cell.setCellValue(Utility.setCommaFormat(field_over_rate_1c)+"%");//狹義逾期放款.逾放比率(較上月底增減比率)
			    	 
			    	 cell=row.createCell((short)6);
			    	 cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                     cell.setCellStyle(cs_right);
			    	 cell.setCellValue(Utility.setCommaFormat(amt840740_840760));//廣義逾期放款.逾放金額=狹義逾放金額+應予觀察放款
			    	 
			    	 cell=row.createCell((short)7);
			    	 cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                     cell.setCellStyle(cs_right);
			    	 cell.setCellValue(Utility.setCommaFormat(amt840740_840760_1));//廣義逾期放款.較上月底增減金額
			    	 
			    	 cell=row.createCell((short)8);
			    	 cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                     cell.setCellStyle(cs_right);
			    	 cell.setCellValue(Utility.setCommaFormat(field_over_rate_2)+"%");//廣義逾期放款.逾放比率
			    	 
			    	 cell=row.createCell((short)9);
			    	 cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                     cell.setCellStyle(cs_right);
			    	 cell.setCellValue(Utility.setCommaFormat(field_over_rate_2c)+"%");//廣義逾期放款.逾放比率(較上月底增減比率)
			    	 
			    	 cell=row.createCell((short)10);
			    	 cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                     cell.setCellStyle(cs_right);
			    	 cell.setCellValue(Utility.setCommaFormat(amt840710));//(1)放款本金未超過清償期三個月，惟利息未按期繳納超過三個月至六個月者 
			    	 
			    	 cell=row.createCell((short)11);
			    	 cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                     cell.setCellStyle(cs_right);
			    	 cell.setCellValue(Utility.setCommaFormat(amt840720));//(2)中長期分期償還放款，未按期攤還超過三個月至六個月者 
			    	 
			    	 cell=row.createCell((short)12);
			    	 cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                     cell.setCellStyle(cs_right);
			    	 cell.setCellValue(Utility.setCommaFormat(amt840731));//(a)協議分期償還放款，協議條件符合規定，且借款戶依協議條件按期履約未滿六個月者
			    	 
			    	 cell=row.createCell((short)13);
			    	 cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                     cell.setCellStyle(cs_right);
			    	 cell.setCellValue(Utility.setCommaFormat(amt840732));//(b)已獲信用保證基金同意理賠款項或有足額存單或存款備償(須辦妥質權設定且徵得發單行拋棄抵銷權同意書)，而約定待其他債務人財產處分後再予沖償者
			    	 
			    	 cell=row.createCell((short)14);
			    	 cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                     cell.setCellStyle(cs_right);
			    	 cell.setCellValue(Utility.setCommaFormat(amt840733));//(c)已確定分配之債權，惟尚未接獲分配款者 
			    	 
			    	 cell=row.createCell((short)15);
			    	 cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                     cell.setCellStyle(cs_right);
			    	 cell.setCellValue(Utility.setCommaFormat(amt840734));//(d)債務人兼擔保品提供人死亡，於辦理繼承期間屆期而未清償之放款 
			    	 
			    	 cell=row.createCell((short)16);
			    	 cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                     cell.setCellStyle(cs_right);
			    	 cell.setCellValue(Utility.setCommaFormat(amt840735));//(e)其他  
			    	 
			    	 cell=row.createCell((short)17);
			    	 cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                     cell.setCellStyle(cs_right);
			    	 cell.setCellValue(Utility.setCommaFormat(amt840730));//小計(a)+(b)+(c)+(d)+(e)
			    	 
			    	 cell=row.createCell((short)18);
			    	 cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                     cell.setCellStyle(cs_right);
			    	 cell.setCellValue(Utility.setCommaFormat(amt8407X0));//總計(1)+(2)+(3)
			    	 
			    	 cell=row.createCell((short)19);
			    	 cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                     cell.setCellStyle(cs_right);
			    	 cell.setCellValue(Utility.setCommaFormat(field_over_rate_3)+"%");//占同日放款總餘額比率
			    	 
			    	 cell=row.createCell((short)0);
                     if("A01".equals(field_seq)){
                         cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                         cell.setCellStyle(cs_left);
                         cell.setCellValue(bank_name);//機構名稱
                     }else if("A90".equals(field_seq)) {
                         cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                         cell.setCellStyle(cs_left);
                         cell.setCellValue(hsien_name);//縣市名稱
                     }
                     if("A90".equals(field_seq)) {
                         rowNum++;
                         row = sheet.createRow(rowNum);
                         for(int j=1;j<20;j++){
                             cell=row.createCell((short)j);
                             cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                             cell.setCellStyle(cs_right);
                             cell.setCellValue("");
                         }
                         rowNum2++;
                     }
                     
                         rowNum++;  
                     
			    }
			}//end of bean
	      	  
	      	
              bean = (DataObject)dbData.get(0);
              field_seq=(bean.getValue("field_seq") == null)?"":(String)(bean.getValue("field_seq"));
              hsien_id=(bean.getValue("hsien_id") == null)?"":(String)(bean.getValue("hsien_id"));
              hsien_name=(bean.getValue("hsien_name") == null)?"":(bean.getValue("hsien_name")).toString();
              bank_code=(bean.getValue("bank_code") == null)?"":(bean.getValue("bank_code")).toString();
              bank_name=(bean.getValue("bank_name") == null)?"":(bean.getValue("bank_name")).toString();
              amt840750=(bean.getValue("amt840750") == null)?"":(bean.getValue("amt840750")).toString();//--本月底放款總額
              amt840740=(bean.getValue("amt840740") == null)?"":(bean.getValue("amt840740")).toString();//--廣義逾期放款.逾放金額
              amt840740_1=(bean.getValue("amt840740_1") == null)?"":(bean.getValue("amt840740_1")).toString();//--廣義逾期放款.較上月底增減金額          
              field_over_rate_1=(bean.getValue("field_over_rate_1") == null)?"":(bean.getValue("field_over_rate_1")).toString();//--狹義逾期放款.逾放比率
              field_over_rate_1b=(bean.getValue("field_over_rate_1b") == null)?"":(bean.getValue("field_over_rate_1b")).toString();//--狹義逾期放款.逾放比率(上月)
              field_over_rate_1c=(bean.getValue("field_over_rate_1c") == null)?"":(bean.getValue("field_over_rate_1c")).toString();//--狹義逾期放款.逾放比率(較上月底增減比率)
              amt840740_rate=(bean.getValue("amt840740_rate") == null)?"":(bean.getValue("amt840740_rate")).toString();//--較上月底增減比率
              amt840740_840760=(bean.getValue("amt840740_840760") == null)?"":(bean.getValue("amt840740_840760")).toString();//--逾放金額=狹義逾放金額+應予觀察放款
              amt840740_840760_1=(bean.getValue("amt840740_840760_1") == null)?"":(bean.getValue("amt840740_840760_1")).toString();//--廣義逾期放款.較上月底增減金額
              field_over_rate_2=(bean.getValue("field_over_rate_2") == null)?"":(bean.getValue("field_over_rate_2")).toString();//--廣義逾期放款.逾放比率
              field_over_rate_2b=(bean.getValue("field_over_rate_2b") == null)?"":(bean.getValue("field_over_rate_2b")).toString();//--廣義逾期放款.逾放比率(上月)
              field_over_rate_2c=(bean.getValue("field_over_rate_2c") == null)?"":(bean.getValue("field_over_rate_2c")).toString();//--廣義逾期放款.逾放比率(較上月底增減比率)
              amt840710=(bean.getValue("amt840710") == null)?"":(bean.getValue("amt840710")).toString();
              amt840720=(bean.getValue("amt840720") == null)?"":(bean.getValue("amt840720")).toString();
              amt840731=(bean.getValue("amt840731") == null)?"":(bean.getValue("amt840731")).toString();
              amt840732=(bean.getValue("amt840732") == null)?"":(bean.getValue("amt840732")).toString();
              amt840733=(bean.getValue("amt840733") == null)?"":(bean.getValue("amt840733")).toString();
              amt840734=(bean.getValue("amt840734") == null)?"":(bean.getValue("amt840734")).toString();
              amt840735=(bean.getValue("amt840735") == null)?"":(bean.getValue("amt840735")).toString();
              amt840730=(bean.getValue("amt840730") == null)?"":(bean.getValue("amt840730")).toString();
              amt8407X0=(bean.getValue("amt8407x0") == null)?"":(bean.getValue("amt8407x0")).toString();
              field_over_rate_3=(bean.getValue("field_over_rate_3") == null)?"":(bean.getValue("field_over_rate_3")).toString();//--占同日放款總餘額比率 
              
              rowNum=6+dbData.size()+rowNum2-3;
              row = sheet.createRow(rowNum);
             
             cell=row.createCell((short)0);
             cell.setEncoding(HSSFCell.ENCODING_UTF_16);
             cell.setCellStyle(cs_right);
             cell.setCellValue(hsien_name);//總計
             
             cell=row.createCell((short)1);
             cell.setEncoding(HSSFCell.ENCODING_UTF_16);
             cell.setCellStyle(cs_right);
             cell.setCellValue(Utility.setCommaFormat(amt840750));//本月底放款總額 
             
             cell=row.createCell((short)2);
             cell.setEncoding(HSSFCell.ENCODING_UTF_16);
             cell.setCellStyle(cs_right);
             cell.setCellValue(Utility.setCommaFormat(amt840740));//狹義逾期放款.逾放金額
             
             cell=row.createCell((short)3);
             cell.setEncoding(HSSFCell.ENCODING_UTF_16);
             cell.setCellStyle(cs_right);
             cell.setCellValue(Utility.setCommaFormat(amt840740_1));//狹義逾期放款.較上月底增減金額                 
             
             cell=row.createCell((short)4);
             cell.setEncoding(HSSFCell.ENCODING_UTF_16);
             cell.setCellStyle(cs_right);
             cell.setCellValue(Utility.setCommaFormat(field_over_rate_1)+"%");//狹義逾期放款.逾放比率
             
             cell=row.createCell((short)5);
             cell.setEncoding(HSSFCell.ENCODING_UTF_16);
             cell.setCellStyle(cs_right);
             cell.setCellValue(Utility.setCommaFormat(field_over_rate_1c)+"%");//狹義逾期放款.逾放比率(較上月底增減比率)
             
             cell=row.createCell((short)6);
             cell.setEncoding(HSSFCell.ENCODING_UTF_16);
             cell.setCellStyle(cs_right);
             cell.setCellValue(Utility.setCommaFormat(amt840740_840760));//廣義逾期放款.逾放金額=狹義逾放金額+應予觀察放款
             
             cell=row.createCell((short)7);
             cell.setEncoding(HSSFCell.ENCODING_UTF_16);
             cell.setCellStyle(cs_right);
             cell.setCellValue(Utility.setCommaFormat(amt840740_840760_1));//廣義逾期放款.較上月底增減金額
             
             cell=row.createCell((short)8);
             cell.setEncoding(HSSFCell.ENCODING_UTF_16);
             cell.setCellStyle(cs_right);
             cell.setCellValue(Utility.setCommaFormat(field_over_rate_2)+"%");//廣義逾期放款.逾放比率
             
             cell=row.createCell((short)9);
             cell.setEncoding(HSSFCell.ENCODING_UTF_16);
             cell.setCellStyle(cs_right);
             cell.setCellValue(Utility.setCommaFormat(field_over_rate_2c)+"%");//廣義逾期放款.逾放比率(較上月底增減比率)
             
             cell=row.createCell((short)10);
             cell.setEncoding(HSSFCell.ENCODING_UTF_16);
             cell.setCellStyle(cs_right);
             cell.setCellValue(Utility.setCommaFormat(amt840710));//(1)放款本金未超過清償期三個月，惟利息未按期繳納超過三個月至六個月者 
             
             cell=row.createCell((short)11);
             cell.setEncoding(HSSFCell.ENCODING_UTF_16);
             cell.setCellStyle(cs_right);
             cell.setCellValue(Utility.setCommaFormat(amt840720));//(2)中長期分期償還放款，未按期攤還超過三個月至六個月者 
             
             cell=row.createCell((short)12);
             cell.setEncoding(HSSFCell.ENCODING_UTF_16);
             cell.setCellStyle(cs_right);
             cell.setCellValue(Utility.setCommaFormat(amt840731));//(a)協議分期償還放款，協議條件符合規定，且借款戶依協議條件按期履約未滿六個月者
             
             cell=row.createCell((short)13);
             cell.setEncoding(HSSFCell.ENCODING_UTF_16);
             cell.setCellStyle(cs_right);
             cell.setCellValue(Utility.setCommaFormat(amt840732));//(b)已獲信用保證基金同意理賠款項或有足額存單或存款備償(須辦妥質權設定且徵得發單行拋棄抵銷權同意書)，而約定待其他債務人財產處分後再予沖償者
             
             cell=row.createCell((short)14);
             cell.setEncoding(HSSFCell.ENCODING_UTF_16);
             cell.setCellStyle(cs_right);
             cell.setCellValue(Utility.setCommaFormat(amt840733));//(c)已確定分配之債權，惟尚未接獲分配款者 
             
             cell=row.createCell((short)15);
             cell.setEncoding(HSSFCell.ENCODING_UTF_16);
             cell.setCellStyle(cs_right);
             cell.setCellValue(Utility.setCommaFormat(amt840734));//(d)債務人兼擔保品提供人死亡，於辦理繼承期間屆期而未清償之放款 
             
             cell=row.createCell((short)16);
             cell.setEncoding(HSSFCell.ENCODING_UTF_16);
             cell.setCellStyle(cs_right);
             cell.setCellValue(Utility.setCommaFormat(amt840735));//(e)其他  
             
             cell=row.createCell((short)17);
             cell.setEncoding(HSSFCell.ENCODING_UTF_16);
             cell.setCellStyle(cs_right);
             cell.setCellValue(Utility.setCommaFormat(amt840730));//小計(a)+(b)+(c)+(d)+(e)
             
             cell=row.createCell((short)18);
             cell.setEncoding(HSSFCell.ENCODING_UTF_16);
             cell.setCellStyle(cs_right);
             cell.setCellValue(Utility.setCommaFormat(amt8407X0));//總計(1)+(2)+(3)
             
             cell=row.createCell((short)19);
             cell.setEncoding(HSSFCell.ENCODING_UTF_16);
             cell.setCellStyle(cs_right);
             cell.setCellValue(Utility.setCommaFormat(field_over_rate_3)+"%");//占同日放款總餘額比率
        
             int rowNum3 = 6+dbData.size()+rowNum2;
             
             for(int i=0;i<dbData.size();i++){                
                 bean = (DataObject)dbData.get(i);
                 field_seq=(bean.getValue("field_seq") == null)?"":(String)(bean.getValue("field_seq"));
                 hsien_id=(bean.getValue("hsien_id") == null)?"":(bean.getValue("hsien_id")).toString();
                 hsien_name=(bean.getValue("hsien_name") == null)?"":(bean.getValue("hsien_name")).toString();
                 bank_code=(bean.getValue("bank_code") == null)?"":(bean.getValue("bank_code")).toString();
                 bank_name=(bean.getValue("bank_name") == null)?"":(bean.getValue("bank_name")).toString();
                 amt840750=(bean.getValue("amt840750") == null)?"":(bean.getValue("amt840750")).toString();//--本月底放款總額
                 amt840740=(bean.getValue("amt840740") == null)?"":(bean.getValue("amt840740")).toString();//--廣義逾期放款.逾放金額
                 amt840740_1=(bean.getValue("amt840740_1") == null)?"":(bean.getValue("amt840740_1")).toString();//--廣義逾期放款.較上月底增減金額          
                 field_over_rate_1=(bean.getValue("field_over_rate_1") == null)?"":(bean.getValue("field_over_rate_1")).toString();//--狹義逾期放款.逾放比率
                 field_over_rate_1b=(bean.getValue("field_over_rate_1b") == null)?"":(bean.getValue("field_over_rate_1b")).toString();//--狹義逾期放款.逾放比率(上月)
                 field_over_rate_1c=(bean.getValue("field_over_rate_1c") == null)?"":(bean.getValue("field_over_rate_1c")).toString();//--狹義逾期放款.逾放比率(較上月底增減比率)
                 amt840740_rate=(bean.getValue("amt840740_rate") == null)?"":(bean.getValue("amt840740_rate")).toString();//--較上月底增減比率
                 amt840740_840760=(bean.getValue("amt840740_840760") == null)?"":(bean.getValue("amt840740_840760")).toString();//--逾放金額=狹義逾放金額+應予觀察放款
                 amt840740_840760_1=(bean.getValue("amt840740_840760_1") == null)?"":(bean.getValue("amt840740_840760_1")).toString();//--廣義逾期放款.較上月底增減金額
                 field_over_rate_2=(bean.getValue("field_over_rate_2") == null)?"":(bean.getValue("field_over_rate_2")).toString();//--廣義逾期放款.逾放比率
                 field_over_rate_2b=(bean.getValue("field_over_rate_2b") == null)?"":(bean.getValue("field_over_rate_2b")).toString();//--廣義逾期放款.逾放比率(上月)
                 field_over_rate_2c=(bean.getValue("field_over_rate_2c") == null)?"":(bean.getValue("field_over_rate_2c")).toString();//--廣義逾期放款.逾放比率(較上月底增減比率)
                 amt840710=(bean.getValue("amt840710") == null)?"":(bean.getValue("amt840710")).toString();
                 amt840720=(bean.getValue("amt840720") == null)?"":(bean.getValue("amt840720")).toString();
                 amt840731=(bean.getValue("amt840731") == null)?"":(bean.getValue("amt840731")).toString();
                 amt840732=(bean.getValue("amt840732") == null)?"":(bean.getValue("amt840732")).toString();
                 amt840733=(bean.getValue("amt840733") == null)?"":(bean.getValue("amt840733")).toString();
                 amt840734=(bean.getValue("amt840734") == null)?"":(bean.getValue("amt840734")).toString();
                 amt840735=(bean.getValue("amt840735") == null)?"":(bean.getValue("amt840735")).toString();
                 amt840730=(bean.getValue("amt840730") == null)?"":(bean.getValue("amt840730")).toString();
                 amt8407X0=(bean.getValue("amt8407x0") == null)?"":(bean.getValue("amt8407x0")).toString();
                 field_over_rate_3=(bean.getValue("field_over_rate_3") == null)?"":(bean.getValue("field_over_rate_3")).toString();//--占同日放款總餘額比率
                 
                 if(("A90".equals(field_seq)&&("e".equals(hsien_id)||"d".equals(hsien_id)||"b".equals(hsien_id)||"A".equals(hsien_id)||"f".equals(hsien_id)||"H".equals(hsien_id))) || "A92".equals(field_seq)){
                    row = sheet.createRow(rowNum3);
                    cell=row.createCell((short)0);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellStyle(cs_right);
                    cell.setCellValue(hsien_name);//機構名稱
                    
                    cell=row.createCell((short)1);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellStyle(cs_right);
                    cell.setCellValue(Utility.setCommaFormat(amt840750));//本月底放款總額 
                    
                    cell=row.createCell((short)2);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellStyle(cs_right);
                    cell.setCellValue(Utility.setCommaFormat(amt840740));//狹義逾期放款.逾放金額
                    
                    cell=row.createCell((short)3);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellStyle(cs_right);
                    cell.setCellValue(Utility.setCommaFormat(amt840740_1));//狹義逾期放款.較上月底增減金額                 
                    
                    cell=row.createCell((short)4);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellStyle(cs_right);
                    cell.setCellValue(Utility.setCommaFormat(field_over_rate_1)+"%");//狹義逾期放款.逾放比率
                    
                    cell=row.createCell((short)5);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellStyle(cs_right);
                    cell.setCellValue(Utility.setCommaFormat(field_over_rate_1c)+"%");//狹義逾期放款.逾放比率(較上月底增減比率)
                    
                    cell=row.createCell((short)6);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellStyle(cs_right);
                    cell.setCellValue(Utility.setCommaFormat(amt840740_840760));//廣義逾期放款.逾放金額=狹義逾放金額+應予觀察放款
                    
                    cell=row.createCell((short)7);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellStyle(cs_right);
                    cell.setCellValue(Utility.setCommaFormat(amt840740_840760_1));//廣義逾期放款.較上月底增減金額
                    
                    cell=row.createCell((short)8);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellStyle(cs_right);
                    cell.setCellValue(Utility.setCommaFormat(field_over_rate_2)+"%");//廣義逾期放款.逾放比率
                    
                    cell=row.createCell((short)9);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellStyle(cs_right);
                    cell.setCellValue(Utility.setCommaFormat(field_over_rate_2c)+"%");//廣義逾期放款.逾放比率(較上月底增減比率)
                    
                    cell=row.createCell((short)10);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellStyle(cs_right);
                    cell.setCellValue(Utility.setCommaFormat(amt840710));//(1)放款本金未超過清償期三個月，惟利息未按期繳納超過三個月至六個月者 
                    
                    cell=row.createCell((short)11);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellStyle(cs_right);
                    cell.setCellValue(Utility.setCommaFormat(amt840720));//(2)中長期分期償還放款，未按期攤還超過三個月至六個月者 
                    
                    cell=row.createCell((short)12);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellStyle(cs_right);
                    cell.setCellValue(Utility.setCommaFormat(amt840731));//(a)協議分期償還放款，協議條件符合規定，且借款戶依協議條件按期履約未滿六個月者
                    
                    cell=row.createCell((short)13);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellStyle(cs_right);
                    cell.setCellValue(Utility.setCommaFormat(amt840732));//(b)已獲信用保證基金同意理賠款項或有足額存單或存款備償(須辦妥質權設定且徵得發單行拋棄抵銷權同意書)，而約定待其他債務人財產處分後再予沖償者
                    
                    cell=row.createCell((short)14);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellStyle(cs_right);
                    cell.setCellValue(Utility.setCommaFormat(amt840733));//(c)已確定分配之債權，惟尚未接獲分配款者 
                    
                    cell=row.createCell((short)15);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellStyle(cs_right);
                    cell.setCellValue(Utility.setCommaFormat(amt840734));//(d)債務人兼擔保品提供人死亡，於辦理繼承期間屆期而未清償之放款 
                    
                    cell=row.createCell((short)16);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellStyle(cs_right);
                    cell.setCellValue(Utility.setCommaFormat(amt840735));//(e)其他  
                    
                    cell=row.createCell((short)17);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellStyle(cs_right);
                    cell.setCellValue(Utility.setCommaFormat(amt840730));//小計(a)+(b)+(c)+(d)+(e)
                    
                    cell=row.createCell((short)18);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellStyle(cs_right);
                    cell.setCellValue(Utility.setCommaFormat(amt8407X0));//總計(1)+(2)+(3)
                    
                    cell=row.createCell((short)19);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellStyle(cs_right);
                    cell.setCellValue(Utility.setCommaFormat(field_over_rate_3)+"%");//占同日放款總餘額比率
                    rowNum3++;  
               }
                 
           }//end of bean
           
             bean = (DataObject)dbData.get(0);
             field_seq=(bean.getValue("field_seq") == null)?"":(String)(bean.getValue("field_seq"));
             hsien_id=(bean.getValue("hsien_id") == null)?"":(String)(bean.getValue("hsien_id"));
             hsien_name=(bean.getValue("hsien_name") == null)?"":(bean.getValue("hsien_name")).toString();
             bank_code=(bean.getValue("bank_code") == null)?"":(bean.getValue("bank_code")).toString();
             bank_name=(bean.getValue("bank_name") == null)?"":(bean.getValue("bank_name")).toString();
             amt840750=(bean.getValue("amt840750") == null)?"":(bean.getValue("amt840750")).toString();//--本月底放款總額
             amt840740=(bean.getValue("amt840740") == null)?"":(bean.getValue("amt840740")).toString();//--廣義逾期放款.逾放金額
             amt840740_1=(bean.getValue("amt840740_1") == null)?"":(bean.getValue("amt840740_1")).toString();//--廣義逾期放款.較上月底增減金額          
             field_over_rate_1=(bean.getValue("field_over_rate_1") == null)?"":(bean.getValue("field_over_rate_1")).toString();//--狹義逾期放款.逾放比率
             field_over_rate_1b=(bean.getValue("field_over_rate_1b") == null)?"":(bean.getValue("field_over_rate_1b")).toString();//--狹義逾期放款.逾放比率(上月)
             field_over_rate_1c=(bean.getValue("field_over_rate_1c") == null)?"":(bean.getValue("field_over_rate_1c")).toString();//--狹義逾期放款.逾放比率(較上月底增減比率)
             amt840740_rate=(bean.getValue("amt840740_rate") == null)?"":(bean.getValue("amt840740_rate")).toString();//--較上月底增減比率
             amt840740_840760=(bean.getValue("amt840740_840760") == null)?"":(bean.getValue("amt840740_840760")).toString();//--逾放金額=狹義逾放金額+應予觀察放款
             amt840740_840760_1=(bean.getValue("amt840740_840760_1") == null)?"":(bean.getValue("amt840740_840760_1")).toString();//--廣義逾期放款.較上月底增減金額
             field_over_rate_2=(bean.getValue("field_over_rate_2") == null)?"":(bean.getValue("field_over_rate_2")).toString();//--廣義逾期放款.逾放比率
             field_over_rate_2b=(bean.getValue("field_over_rate_2b") == null)?"":(bean.getValue("field_over_rate_2b")).toString();//--廣義逾期放款.逾放比率(上月)
             field_over_rate_2c=(bean.getValue("field_over_rate_2c") == null)?"":(bean.getValue("field_over_rate_2c")).toString();//--廣義逾期放款.逾放比率(較上月底增減比率)
             amt840710=(bean.getValue("amt840710") == null)?"":(bean.getValue("amt840710")).toString();
             amt840720=(bean.getValue("amt840720") == null)?"":(bean.getValue("amt840720")).toString();
             amt840731=(bean.getValue("amt840731") == null)?"":(bean.getValue("amt840731")).toString();
             amt840732=(bean.getValue("amt840732") == null)?"":(bean.getValue("amt840732")).toString();
             amt840733=(bean.getValue("amt840733") == null)?"":(bean.getValue("amt840733")).toString();
             amt840734=(bean.getValue("amt840734") == null)?"":(bean.getValue("amt840734")).toString();
             amt840735=(bean.getValue("amt840735") == null)?"":(bean.getValue("amt840735")).toString();
             amt840730=(bean.getValue("amt840730") == null)?"":(bean.getValue("amt840730")).toString();
             amt8407X0=(bean.getValue("amt8407x0") == null)?"":(bean.getValue("amt8407x0")).toString();
             field_over_rate_3=(bean.getValue("field_over_rate_3") == null)?"":(bean.getValue("field_over_rate_3")).toString();//--占同日放款總餘額比率 
             
            row = sheet.createRow(rowNum3);
            cell=row.createCell((short)0);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellStyle(cs_right);
            cell.setCellValue(hsien_name);//總計
            
            cell=row.createCell((short)1);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellStyle(cs_right);
            cell.setCellValue(Utility.setCommaFormat(amt840750));//本月底放款總額 
            
            cell=row.createCell((short)2);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellStyle(cs_right);
            cell.setCellValue(Utility.setCommaFormat(amt840740));//狹義逾期放款.逾放金額
            
            cell=row.createCell((short)3);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellStyle(cs_right);
            cell.setCellValue(Utility.setCommaFormat(amt840740_1));//狹義逾期放款.較上月底增減金額                 
            
            cell=row.createCell((short)4);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellStyle(cs_right);
            cell.setCellValue(Utility.setCommaFormat(field_over_rate_1)+"%");//狹義逾期放款.逾放比率
            
            cell=row.createCell((short)5);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellStyle(cs_right);
            cell.setCellValue(Utility.setCommaFormat(field_over_rate_1c)+"%");//狹義逾期放款.逾放比率(較上月底增減比率)
            
            cell=row.createCell((short)6);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellStyle(cs_right);
            cell.setCellValue(Utility.setCommaFormat(amt840740_840760));//廣義逾期放款.逾放金額=狹義逾放金額+應予觀察放款
            
            cell=row.createCell((short)7);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellStyle(cs_right);
            cell.setCellValue(Utility.setCommaFormat(amt840740_840760_1));//廣義逾期放款.較上月底增減金額
            
            cell=row.createCell((short)8);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellStyle(cs_right);
            cell.setCellValue(Utility.setCommaFormat(field_over_rate_2)+"%");//廣義逾期放款.逾放比率
            
            cell=row.createCell((short)9);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellStyle(cs_right);
            cell.setCellValue(Utility.setCommaFormat(field_over_rate_2c)+"%");//廣義逾期放款.逾放比率(較上月底增減比率)
            
            cell=row.createCell((short)10);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellStyle(cs_right);
            cell.setCellValue(Utility.setCommaFormat(amt840710));//(1)放款本金未超過清償期三個月，惟利息未按期繳納超過三個月至六個月者 
            
            cell=row.createCell((short)11);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellStyle(cs_right);
            cell.setCellValue(Utility.setCommaFormat(amt840720));//(2)中長期分期償還放款，未按期攤還超過三個月至六個月者 
            
            cell=row.createCell((short)12);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellStyle(cs_right);
            cell.setCellValue(Utility.setCommaFormat(amt840731));//(a)協議分期償還放款，協議條件符合規定，且借款戶依協議條件按期履約未滿六個月者
            
            cell=row.createCell((short)13);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellStyle(cs_right);
            cell.setCellValue(Utility.setCommaFormat(amt840732));//(b)已獲信用保證基金同意理賠款項或有足額存單或存款備償(須辦妥質權設定且徵得發單行拋棄抵銷權同意書)，而約定待其他債務人財產處分後再予沖償者
            
            cell=row.createCell((short)14);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellStyle(cs_right);
            cell.setCellValue(Utility.setCommaFormat(amt840733));//(c)已確定分配之債權，惟尚未接獲分配款者 
            
            cell=row.createCell((short)15);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellStyle(cs_right);
            cell.setCellValue(Utility.setCommaFormat(amt840734));//(d)債務人兼擔保品提供人死亡，於辦理繼承期間屆期而未清償之放款 
            
            cell=row.createCell((short)16);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellStyle(cs_right);
            cell.setCellValue(Utility.setCommaFormat(amt840735));//(e)其他  
            
            cell=row.createCell((short)17);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellStyle(cs_right);
            cell.setCellValue(Utility.setCommaFormat(amt840730));//小計(a)+(b)+(c)+(d)+(e)
            
            cell=row.createCell((short)18);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellStyle(cs_right);
            cell.setCellValue(Utility.setCommaFormat(amt8407X0));//總計(1)+(2)+(3)
            
            cell=row.createCell((short)19);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellStyle(cs_right);
            cell.setCellValue(Utility.setCommaFormat(field_over_rate_3)+"%");//占同日放款總餘額比率
	      	
	      	  
	      }else{ //end of else dbData.size() != 0
	      	 cell.setCellValue("基準日:"+m_year+"年"+m_month+"月底無資料存在");
	      }
	      
	      FileOutputStream fout =  new FileOutputStream(reportDir + System.getProperty("file.separator") + "農漁會信用部逾期放款統計表_總表.xls");
	      HSSFFooter footer = sheet.getFooter();
	      footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
	      footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
	      wb.write(fout);
	        //儲存 
	      fout.close();
	      System.out.println("儲存完成");
	    }catch (Exception e) {
	      System.out.println("RptFR053W.createRpt Error:" + e + e.getMessage());
	    }
	    
	    return errMsg;
	  }
}
