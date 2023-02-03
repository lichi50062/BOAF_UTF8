/*
 * 101.01.09 create 全體農(漁)會信用部聯合貸款案件彙總表 by 2968
 * 103.08.14 add 調整授信總金額,若無01結尾的案件,以下一筆計算授信總金額 by 2295
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

public class RptFR064W {	 
	 public static String createRpt(String m_year,String m_month,String unit,String bank_type,String febxlsFlag,HSSFWorkbook wb) {    

	    String errMsg = "";	    
	    List dbData = null;	    
	    int rowNum=0;
	    DataObject bean = null;
	    reportUtil reportUtil = new reportUtil();
		HSSFCellStyle cs_right = null; 
		HSSFCellStyle cs_center = null;
	    HSSFCellStyle cs_left = null;
	    String bank_no=""; //--機構代號 
	    String bank_name=""; //--機構名稱
	    String hsien_name="";
	    String count_seq="";
	    String field_seq="";
	    String loan_amt_sum_1="";  //--金庫.授信案總金額
	    String loan_amt_sum_2=""; //--其他.授信案總金額
	    String loan_amt_sum_all=""; //--合計.授信案總金額
	    String loan_amt_1=""; // --金庫.參貸額度             
	    String loan_amt_2=""; //--其他.參貸額度
	    String loan_amt_all=""; //--合計.參貸額度
	    String loan_bal_amt_1=""; // --金庫.實際授信餘額
	    String loan_bal_amt_2=""; //--其他.實際授信餘額
	    String loan_bal_amt_all=""; //--合計.實際授信餘額
	    String field_OVER_1=""; //--金庫.逾放金額
	    String field_OVER_2=""; //--其他.逾放金額
	    String field_OVER_all=""; //--合計.逾放金額
	    String field_OVER_RATE_1=""; //--金庫.逾放金額占實際授信餘額比率
	    String field_OVER_RATE_2=""; //--其他.逾放金額占實際授信餘額比率
	    String field_OVER_RATE_all=""; //--合計.逾放金額占實際授信餘額比率
	    String loan_type_amt_1=""; //--授信用途.購地.實際授信餘額
	    String loan_type_amt_2=""; //--授信用途.建築.實際授信餘額
	    String loan_type_amt_3=""; //--授信用途.其他.實際授信餘額
	    String loan_type_all=""; //--授信用途.合計.實際授信餘額
	    String bank_type_name =( bank_type.equals("6") )?"農會":"漁會";
	    String unit_name="";
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
	      if(febxlsFlag.equals("")){//農漁會信用部聯合貸款案件彙總表
	      	//input the standard report form      
	      	finput = new FileInputStream(xlsDir +System.getProperty("file.separator") +"農漁會信用部聯合貸款案件彙總表.xls");      
	      
	      	//設定FileINputStream讀取Excel檔
	      	POIFSFileSystem fs = new POIFSFileSystem(finput);	      
	        wb = new HSSFWorkbook(fs);
	      }else{
	      	System.out.println("格式:農漁會信用部聯合貸款案件彙總表");
	      }
	      HSSFSheet sheet = wb.getSheetAt((febxlsFlag.equals("")?0:2)); //讀取第一個工作表，宣告其為sheet
	      HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
	      //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	      //sheet.setAutobreaks(true); //自動分頁

	      //設定頁面符合列印大小
	      sheet.setAutobreaks(false);
	      ps.setScale( (short) 55); //列印縮放百分比	      
	      ps.setPaperSize( (short) 9); //設定紙張大小 A4
	      //wb.setSheetName(0,"test");
	      
	      if(febxlsFlag.equals("")) finput.close();

	      HSSFRow row = null; //宣告一列
	      HSSFCell cell = null; //宣告一個儲存格	      
	      cs_right = reportUtil.getRightStyle(wb);
	      cs_center = reportUtil.getDefaultStyle(wb);
	      cs_left = reportUtil.getLeftStyle(wb);
	      m_year = String.valueOf(Integer.parseInt(m_year));
	      m_month = String.valueOf(Integer.parseInt(m_month));
	      
	      unit_name = Utility.getUnitName(unit);//取得單位名稱
          String wlx01_m_year = "";
	      StringBuffer sql = new StringBuffer();
	      List paramList = new ArrayList();//傳入參數
	      if(m_month.length()==1){
	          m_month = "0"+m_month;
	      }
	      //99.09.10 add 查詢年度100年以前.縣市別不同===============================
	      wlx01_m_year = (Integer.parseInt(m_year) < 100)?"99":"100"; 
	      //===================================================================== 
	      sql.append("select wlx10_m_loan.hsien_id,");
	      sql.append("       wlx10_m_loan.hsien_name,");
	      sql.append("       wlx10_m_loan.FR001W_output_order,");
	      sql.append("       wlx10_m_loan.bank_no,"); //--機構代號 
	      sql.append("       wlx10_m_loan.BANK_NAME,"); //--機構名稱
	      sql.append("       wlx10_m_loan.COUNT_SEQ,"); //--聯貸件數
	      sql.append("       wlx10_m_loan.field_SEQ,");
	      sql.append("       loan_amt_sum_1,");  //--金庫.授信案總金額
	      sql.append("       loan_amt_sum_2,"); //--其他.授信案總金額
	      sql.append("       loan_amt_sum_all,"); //--合計.授信案總金額
	      sql.append("       loan_amt_1,"); // --金庫.參貸額度             
	      sql.append("       loan_amt_2,"); //--其他.參貸額度
	      sql.append("       loan_amt_all,"); //--合計.參貸額度
	      sql.append("       loan_bal_amt_1,"); // --金庫.實際授信餘額
	      sql.append("       loan_bal_amt_2,"); //--其他.實際授信餘額
	      sql.append("       loan_bal_amt_all,"); //--合計.實際授信餘額
	      sql.append("       field_OVER_1,"); //--金庫.逾放金額
	      sql.append("       field_OVER_2,"); //--其他.逾放金額
	      sql.append("       field_OVER_all,"); //--合計.逾放金額
	      sql.append("       field_OVER_RATE_1,"); //--金庫.逾放金額占實際授信餘額比率
	      sql.append("       field_OVER_RATE_2,"); //--其他.逾放金額占實際授信餘額比率
	      sql.append("       field_OVER_RATE_all,"); //--合計.逾放金額占實際授信餘額比率
	      sql.append("       loan_type_amt_1,"); //--授信用途.購地.實際授信餘額
	      sql.append("       loan_type_amt_2,"); //--授信用途.建築.實際授信餘額
	      sql.append("       loan_type_amt_3,"); //--授信用途.其他.實際授信餘額
	      sql.append("       loan_type_all "); //--授信用途.合計.實際授信餘額
	      sql.append("from ( ");
	      sql.append("       select wlx10_m_loan.hsien_id,wlx10_m_loan.hsien_name,wlx10_m_loan.FR001W_output_order,");         
	      sql.append("         wlx10_m_loan.bank_no,wlx10_m_loan.BANK_NAME,");
	      sql.append("         COUNT_SEQ, field_SEQ,");             
	      sql.append("         round(loan_amt_sum_1 /?,0) as loan_amt_sum_1,round(loan_amt_sum_2 /?,0) as loan_amt_sum_2,round(loan_amt_sum_all /?,0) as loan_amt_sum_all,"); 
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);           
	      sql.append("         round(loan_amt_1 /?,0) as loan_amt_1,round(loan_amt_2 /?,0) as loan_amt_2,round(loan_amt_all /?,0) as loan_amt_all,");
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);
	      sql.append("         round(loan_bal_amt_1 /?,0) as loan_bal_amt_1,round(loan_bal_amt_2 /?,0) as loan_bal_amt_2,round(loan_bal_amt_all /?,0) as loan_bal_amt_all,");
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);
	      sql.append("         round(field_OVER_1 /?,0) as field_OVER_1,round(field_OVER_2 /?,0) as field_OVER_2,round(field_OVER_all /?,0) as field_OVER_all,");
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);
	      sql.append("         decode(loan_bal_amt_1,0,0,round(field_OVER_1 / loan_bal_amt_1 *100 ,2)) as field_OVER_RATE_1,");        
	      sql.append("         decode(loan_bal_amt_2,0,0,round(field_OVER_2 / loan_bal_amt_2 *100 ,2)) as field_OVER_RATE_2,");   
	      sql.append("         decode(loan_bal_amt_all,0,0,round(field_OVER_all / loan_bal_amt_all *100 ,2)) as field_OVER_RATE_all,");             
	      sql.append("         round(loan_type_amt_1 /?,0)  as loan_type_amt_1,round(loan_type_amt_2 /?,0)  as loan_type_amt_2,");
	      sql.append("         round(loan_type_amt_3 /?,0)  as loan_type_amt_3, round(loan_type_all /?,0)  as loan_type_all ");
	      paramList.add(unit);
          paramList.add(unit);
          paramList.add(unit);
          paramList.add(unit);
	      sql.append("  from ("); //--總計begin 
	      sql.append("        select ' '  AS  hsien_id ,  ' 總   計 '   AS hsien_name,  '001' AS FR001W_output_order, ");                 
	      sql.append("               ' ' AS  bank_no , ' '   AS  BANK_NAME, ");               
	      sql.append("               COUNT(*) AS COUNT_SEQ, 'A99' as field_SEQ, ");      
	      sql.append("               SUM(loan_amt_sum_1) loan_amt_sum_1, SUM(loan_amt_sum_2) loan_amt_sum_2, SUM(loan_amt_sum_all) loan_amt_sum_all, ");     
	      sql.append("               SUM(loan_amt_1) loan_amt_1, SUM(loan_amt_2) loan_amt_2, SUM(loan_amt_all) loan_amt_all, ");
	      sql.append("               SUM(loan_bal_amt_1) loan_bal_amt_1, SUM(loan_bal_amt_2) loan_bal_amt_2, SUM(loan_bal_amt_all) loan_bal_amt_all, ");
	      sql.append("               SUM(field_OVER_1) field_OVER_1, SUM(field_OVER_2) field_OVER_2, SUM(field_OVER_all) field_OVER_all, ");           
	      sql.append("               SUM(loan_type_amt_1)  loan_type_amt_1,  SUM(loan_type_amt_2)  loan_type_amt_2, ");
	      sql.append("               SUM(loan_type_amt_3)  loan_type_amt_3,  SUM(loan_type_all)  loan_type_all ");      
	      sql.append("        from ( ");
	      sql.append("             select hsien_id,hsien_name,FR001W_output_order,bank_no,BANK_NAME,case_no, ");
	      sql.append("                    SUM(loan_amt_sum_1) loan_amt_sum_1 ,SUM(loan_amt_sum_2)    loan_amt_sum_2,   SUM(loan_amt_sum_all)  loan_amt_sum_all, ");    
	      sql.append("                    SUM(loan_amt_1)   loan_amt_1,   SUM(loan_amt_2)  loan_amt_2, SUM(loan_amt_all)  loan_amt_all, ");
	      sql.append("                    SUM(loan_bal_amt_1)   loan_bal_amt_1,   SUM(loan_bal_amt_2)  loan_bal_amt_2, SUM(loan_bal_amt_all)  loan_bal_amt_all, ");
	      sql.append("                    SUM(field_OVER_1)   field_OVER_1,   SUM(field_OVER_2)  field_OVER_2, SUM(field_OVER_all)  field_OVER_all, ");
	      sql.append("                    SUM(loan_type_amt_1)  loan_type_amt_1,  SUM(loan_type_amt_2)  loan_type_amt_2, ");
          sql.append("                    SUM(loan_type_amt_3)  loan_type_amt_3,  SUM(loan_type_all)  loan_type_all ");     
          sql.append("             from ( ");
          sql.append("                    select nvl(cd01.hsien_id,' ') as hsien_id, ");               
	      sql.append("                           nvl(cd01.hsien_name,'OTHER') as hsien_name, ");               
	      sql.append("                           cd01.FR001W_output_order     as FR001W_output_order, ");               
	      sql.append("                           bn01.bank_no, bn01.BANK_NAME,substr(case_no,0,7) as case_no, "); 
	      sql.append("                           0 as loan_amt_sum_1, ");//--金庫.授信案總金額 
	      sql.append("                           0 as loan_amt_sum_2, ");//--其他.授信案總金額
	      sql.append("                           0 as loan_amt_sum_all, ");//--合計.授信案總金額
	      //sql.append("                           round(sum(decode(bank_no_max,'1',decode(substr(case_no,9,1),'1',loan_amt_sum,0),0)) /1,0) as loan_amt_sum_1, "); //--金庫.授信案總金額 
	      //sql.append("                           round(sum(decode(bank_no_max,'2',decode(substr(case_no,9,1),'1',loan_amt_sum,0),0)) /1,0) as loan_amt_sum_2, "); //--其他.授信案總金額
	      //sql.append("                           round(sum(decode(substr(case_no,9,1),'1' ,loan_amt_sum,0)) /1,0) as loan_amt_sum_all, "); //--合計.授信案總金額
	      sql.append("                           round(sum(decode(bank_no_max,'1',loan_amt,0)) /1,0) as loan_amt_1, "); //--金庫.參貸額度
	      sql.append("                           round(sum(decode(bank_no_max,'2',loan_amt,0)) /1,0) as loan_amt_2, "); //--其他.參貸額度
	      sql.append("                           round(sum(loan_amt) /1,0) as loan_amt_all, "); //--合計.參貸額度
	      sql.append("                           round(sum(decode(bank_no_max,'1',loan_bal_amt,0)) /1,0) as loan_bal_amt_1, "); //--金庫.實際授信餘額
	      sql.append("                           round(sum(decode(bank_no_max,'2',loan_bal_amt,0)) /1,0) as loan_bal_amt_2, "); //--其他.實際授信餘額
	      sql.append("                           round(sum(loan_bal_amt) /1,0) as loan_bal_amt_all, "); //--合計.實際授信餘額
	      sql.append("                           round(sum(decode(pay_state,'2',decode(bank_no_max,'1',loan_bal_amt,0),0)) /1,0) as field_OVER_1, "); //--金庫.逾放金額
	      sql.append("                           round(sum(decode(pay_state,'2',decode(bank_no_max,'2',loan_bal_amt,0),0)) /1,0) as field_OVER_2, "); //--其他.逾放金額
	      sql.append("                           round(sum(decode(pay_state,'2',loan_bal_amt,0)) /1,0) as field_OVER_all, "); //--合計.逾放金額              
	      sql.append("                           round(sum(decode(loan_type,'1',loan_bal_amt,0)) /1,0) as loan_type_amt_1, "); //--授信用途.購地.實際授信餘額
	      sql.append("                           round(sum(decode(loan_type,'2',loan_bal_amt,0)) /1,0) as loan_type_amt_2, "); //--授信用途.建築.實際授信餘額
	      sql.append("                           round(sum(decode(loan_type,'3',loan_bal_amt,0)) /1,0) as loan_type_amt_3, "); //--授信用途.其他.實際授信餘額
	      sql.append("                           round(sum(decode(loan_type,'1',loan_bal_amt,'2',loan_bal_amt,'3',loan_bal_amt,0)) /1,0) as  loan_type_all "); //--授信用途.合計.實際授信餘額
	      sql.append("                    from  (select * from  cd01 where cd01.hsien_id <> 'Y' ) cd01 ");
	      sql.append("                    left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year =? and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL) ");  
	      paramList.add(wlx01_m_year);
	      sql.append("                    left join bn01 on wlx01.bank_no=bn01.bank_no and bn01.bank_type in ");
          if("".equals(bank_type)){
              sql.append("                    (?,?) ");
              paramList.add("6");
              paramList.add("7");
          }else{
              sql.append("                    (?) ");
              paramList.add(bank_type);
          }
          sql.append("                    and bn01.m_year=? and bn_type <> '2' and wlx01.m_year=? ");
	      paramList.add(wlx01_m_year);
	      paramList.add(wlx01_m_year);
	      sql.append("                    left join (select * from wlx10_m_loan where to_char(m_year * 100 + m_month) = ?) wlx10_m_loan on bn01.bank_no = wlx10_m_loan.bank_no"); 
	      paramList.add(m_year+m_month);
	      sql.append("                    group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME,substr(case_no,0,7)");
	      sql.append("                    union ");
	      sql.append("                    select nvl(cd01.hsien_id,' ') as hsien_id, ");               
          sql.append("                           nvl(cd01.hsien_name,'OTHER') as hsien_name, ");               
          sql.append("                           cd01.FR001W_output_order     as FR001W_output_order, ");               
          sql.append("                           bn01.bank_no, bn01.BANK_NAME,substr(case_no,0,7) as case_no, "); 
          sql.append("                           round(sum(decode(bank_no_max,'1',loan_amt_sum,0)) /1,0)  as loan_amt_sum_1, ");//--金庫.授信案總金額 
          sql.append("                           round(sum(decode(bank_no_max,'2',loan_amt_sum,0)) /1,0) as loan_amt_sum_2, ");//--其他.授信案總金額
          sql.append("                           round(sum(loan_amt_sum) /1,0) as loan_amt_sum_all, ");//--合計.授信案總金額
          sql.append("                           0,0,0,0,0,0,0,0,0,0,0,0,0 ");   
          sql.append("                    from  (select * from  cd01 where cd01.hsien_id <> 'Y' ) cd01 ");
          sql.append("                    left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year =? and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL) ");  
          paramList.add(wlx01_m_year);
          sql.append("                    left join bn01 on wlx01.bank_no=bn01.bank_no and bn01.bank_type in ");
          if("".equals(bank_type)){
              sql.append("                    (?,?) ");
              paramList.add("6");
              paramList.add("7");
          }else{
              sql.append("                    (?) ");
              paramList.add(bank_type);
          }
          sql.append("                    and bn01.m_year=? and bn_type <> '2' and wlx01.m_year=? ");
          paramList.add(wlx01_m_year);
          paramList.add(wlx01_m_year);
          sql.append("                    left join( ");
          sql.append("                            select wlx10_m_loan.* from wlx10_m_loan, "); 
          sql.append("                                 (select m_year,m_month,bank_no,case_grp,case_no ");
          sql.append("                                  from ( ");
          sql.append("                                  select  m_year,m_month,bank_no,substr(case_no,0,7)  as case_grp,min(case_no) as case_no ");
          sql.append("                                  from wlx10_m_loan ");
          sql.append("                                  where m_year=? and m_month=? ");
          sql.append("                                  group by m_year,m_month,bank_no,substr(case_no,0,7) ))a ");
          paramList.add(m_year);
          paramList.add(m_month);
          sql.append("                            where to_char(wlx10_m_loan.m_year * 100 + wlx10_m_loan.m_month) = ?");
          paramList.add(m_year+m_month);
          sql.append("                                    and (wlx10_m_loan.m_year= a.m_year ");
          sql.append("                                    and wlx10_m_loan.m_month=a.m_month ");
          sql.append("                                    and wlx10_m_loan.bank_no = a.bank_no ");
          sql.append("                                    and wlx10_m_loan.case_no = a.case_no) ");                    
          sql.append("                              ) wlx10_m_loan  on  bn01.bank_no = wlx10_m_loan.bank_no ");
          sql.append("                     group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME,substr(case_no,0,7) "); 
          sql.append("                   )group by hsien_id,hsien_name,FR001W_output_order,bank_no,BANK_NAME,case_no ");
	      sql.append("               ) wlx10_m_loan ");
	      sql.append("            where wlx10_m_loan.bank_no <> ' '  ");
	      sql.append("        ) wlx10_m_loan "); //--總計end                  
	      sql.append("        UNION ALL ");
	      sql.append("        select wlx10_m_loan.hsien_id, wlx10_m_loan.hsien_name, wlx10_m_loan.FR001W_output_order, wlx10_m_loan.bank_no, wlx10_m_loan.BANK_NAME, ");       
	      sql.append("           COUNT_SEQ, field_SEQ, ");
	      sql.append("           round(loan_amt_sum_1/?,0) as loan_amt_sum_1, round(loan_amt_sum_2/?,0) as loan_amt_sum_2, round(loan_amt_sum_all/?,0) as loan_amt_sum_all, ");  
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);
	      sql.append("           round(loan_amt_1/?,0) as loan_amt_1, round(loan_amt_2/?,0) as loan_amt_2, round(loan_amt_all/?,0) as loan_amt_all, ");
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);
	      sql.append("           round(loan_bal_amt_1/?,0) as loan_bal_amt_1, round(loan_bal_amt_2/?,0) as loan_bal_amt_2, round(loan_bal_amt_all/?,0) as loan_bal_amt_all, ");
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);
	      sql.append("           round(field_OVER_1/?,0) as field_OVER_1, round(field_OVER_2/?,0) as field_OVER_2, round(field_OVER_all/?,0) as field_OVER_all, ");
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);
	      sql.append("           decode(loan_bal_amt_1,0,0,round(field_OVER_1 / loan_bal_amt_1 *100 ,2)) as field_OVER_RATE_1, ");        
	      sql.append("           decode(loan_bal_amt_2,0,0,round(field_OVER_2 / loan_bal_amt_2 *100 ,2)) as field_OVER_RATE_2, ");   
	      sql.append("           decode(loan_bal_amt_all,0,0,round(field_OVER_all / loan_bal_amt_all *100 ,2)) as field_OVER_RATE_all, ");            
	      sql.append("           round(loan_type_amt_1 /?,0)  as loan_type_amt_1,round(loan_type_amt_2 /?,0)  as loan_type_amt_2, ");
	      sql.append("           round(loan_type_amt_3 /?,0)  as loan_type_amt_3, round(loan_type_all /?,0)  as loan_type_all ");
	      paramList.add(unit);
          paramList.add(unit);
          paramList.add(unit);
          paramList.add(unit);
	      sql.append("    from ");//--各別機構明細begin
	      sql.append("        ( "); 
	      sql.append("         select wlx10_m_loan.hsien_id, wlx10_m_loan.hsien_name, wlx10_m_loan.FR001W_output_order,");                 
	      sql.append("                wlx10_m_loan.bank_no, wlx10_m_loan.BANK_NAME, count(case_no) AS COUNT_SEQ, 'A01' as field_SEQ,");    
	      sql.append("                SUM(loan_amt_sum_1) loan_amt_sum_1, SUM(loan_amt_sum_2) loan_amt_sum_2, SUM(loan_amt_sum_all) loan_amt_sum_all, ");     
	      sql.append("                SUM(loan_amt_1) loan_amt_1, SUM(loan_amt_2) loan_amt_2, SUM(loan_amt_all) loan_amt_all, ");
	      sql.append("                SUM(loan_bal_amt_1) loan_bal_amt_1, SUM(loan_bal_amt_2) loan_bal_amt_2, SUM(loan_bal_amt_all) loan_bal_amt_all, ");
	      sql.append("                SUM(field_OVER_1) field_OVER_1, SUM(field_OVER_2) field_OVER_2, SUM(field_OVER_all) field_OVER_all, ");
	      sql.append("                SUM(loan_type_amt_1)  loan_type_amt_1,  SUM(loan_type_amt_2)  loan_type_amt_2, ");
	      sql.append("                SUM(loan_type_amt_3)  loan_type_amt_3,    SUM(loan_type_all)  loan_type_all ");
	      sql.append("         from ( ");
	      sql.append("              select hsien_id,hsien_name,FR001W_output_order,bank_no,BANK_NAME,case_no,");
	      sql.append("                    SUM(loan_amt_sum_1) loan_amt_sum_1 ,SUM(loan_amt_sum_2)    loan_amt_sum_2,   SUM(loan_amt_sum_all)  loan_amt_sum_all,");     
	      sql.append("                    SUM(loan_amt_1)   loan_amt_1,   SUM(loan_amt_2)  loan_amt_2, SUM(loan_amt_all)  loan_amt_all,");
	      sql.append("                    SUM(loan_bal_amt_1)   loan_bal_amt_1,   SUM(loan_bal_amt_2)  loan_bal_amt_2, SUM(loan_bal_amt_all)  loan_bal_amt_all,");
	      sql.append("                    SUM(field_OVER_1)   field_OVER_1,   SUM(field_OVER_2)  field_OVER_2, SUM(field_OVER_all)  field_OVER_all,");
	      sql.append("                    SUM(loan_type_amt_1)  loan_type_amt_1,  SUM(loan_type_amt_2)  loan_type_amt_2, ");
          sql.append("                    SUM(loan_type_amt_3)  loan_type_amt_3,    SUM(loan_type_all)  loan_type_all ");          
	      sql.append("              from( "); 
	      sql.append("                    select nvl(cd01.hsien_id,' ') as  hsien_id, ");               
	      sql.append("                           nvl(cd01.hsien_name,'OTHER') as  hsien_name, ");              
	      sql.append("                           cd01.FR001W_output_order as FR001W_output_order, ");               
	      sql.append("                           bn01.bank_no, bn01.BANK_NAME,substr(case_no,0,7) as case_no, ");
	      sql.append("                           0 as loan_amt_sum_1,");//--金庫.授信案總金額 
	      sql.append("                           0 as loan_amt_sum_2,");//--其他.授信案總金額
	      sql.append("                           0 as loan_amt_sum_all,");//--合計.授信案總金額
	      //sql.append("                           round(sum(decode(bank_no_max,'1',decode(substr(case_no,9,1),'1',loan_amt_sum,0),0)) /1,0) as loan_amt_sum_1, "); //--金庫.授信案總金額 
          //sql.append("                           round(sum(decode(bank_no_max,'2',decode(substr(case_no,9,1),'1',loan_amt_sum,0),0)) /1,0) as loan_amt_sum_2, "); //--其他.授信案總金額
          //sql.append("                           round(sum(decode(substr(case_no,9,1),'1' ,loan_amt_sum,0)) /1,0) as loan_amt_sum_all, "); //--合計.授信案總金額
          sql.append("                           round(sum(decode(bank_no_max,'1',loan_amt,0)) /1,0) as loan_amt_1, "); //--金庫.參貸額度
          sql.append("                           round(sum(decode(bank_no_max,'2',loan_amt,0)) /1,0) as loan_amt_2, "); //--其他.參貸額度
          sql.append("                           round(sum(loan_amt) /1,0) as loan_amt_all, "); //--合計.參貸額度
          sql.append("                           round(sum(decode(bank_no_max,'1',loan_bal_amt,0)) /1,0) as loan_bal_amt_1, "); //--金庫.實際授信餘額
          sql.append("                           round(sum(decode(bank_no_max,'2',loan_bal_amt,0)) /1,0) as loan_bal_amt_2, "); //--其他.實際授信餘額
          sql.append("                           round(sum(loan_bal_amt) /1,0) as loan_bal_amt_all, "); //--合計.實際授信餘額
          sql.append("                           round(sum(decode(pay_state,'2',decode(bank_no_max,'1',loan_bal_amt,0),0)) /1,0) as field_OVER_1, "); //--金庫.逾放金額
          sql.append("                           round(sum(decode(pay_state,'2',decode(bank_no_max,'2',loan_bal_amt,0),0)) /1,0) as field_OVER_2, "); //--其他.逾放金額
          sql.append("                           round(sum(decode(pay_state,'2',loan_bal_amt,0)) /1,0) as field_OVER_all, "); //--合計.逾放金額 
          sql.append("                           round(sum(decode(loan_type,'1',loan_bal_amt,0)) /1,0) as loan_type_amt_1, "); //--授信用途.購地.實際授信餘額
          sql.append("                           round(sum(decode(loan_type,'2',loan_bal_amt,0)) /1,0) as loan_type_amt_2, "); //--授信用途.建築.實際授信餘額
          sql.append("                           round(sum(decode(loan_type,'3',loan_bal_amt,0)) /1,0) as loan_type_amt_3, "); //--授信用途.其他.實際授信餘額
          sql.append("                           round(sum(decode(loan_type,'1',loan_bal_amt,'2',loan_bal_amt,'3',loan_bal_amt,0)) /1,0) as  loan_type_all "); //--授信用途.合計.實際授信餘額
          sql.append("                    from  (select * from cd01 where cd01.hsien_id <> 'Y'  ) cd01 ");
	      sql.append("                    left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year=? and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL) ");  
	      paramList.add(wlx01_m_year);
	      sql.append("                    left join bn01 on wlx01.bank_no=bn01.bank_no and bn01.bank_type in ");
	      if("".equals(bank_type)){
    	      sql.append("                    (?,?) ");
    	      paramList.add("6");
              paramList.add("7");
	      }else{
	          sql.append("                    (?) ");
              paramList.add(bank_type);
	      }
	      sql.append("                    and bn01.m_year=? and bn_type <> '2' and wlx01.m_year=? ");
	      paramList.add(wlx01_m_year);
	      paramList.add(wlx01_m_year);
	      sql.append("                    left join (select * from wlx10_m_loan where to_char(m_year * 100 + m_month) =?) wlx10_m_loan on bn01.bank_no=wlx10_m_loan.bank_no ");
	      paramList.add(m_year+m_month);
	      sql.append("                    group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME,substr(case_no,0,7)");
	      sql.append("                    union ");
	      sql.append("                    select nvl(cd01.hsien_id,' ')       as  hsien_id ,");               
	      sql.append("                    nvl(cd01.hsien_name,'OTHER') as  hsien_name,");               
	      sql.append("                    cd01.FR001W_output_order     as  FR001W_output_order,");               
	      sql.append("                    bn01.bank_no ,  bn01.BANK_NAME,  substr(case_no,0,7) as case_no,");        
	      sql.append("                    round(sum(decode(bank_no_max,'1',loan_amt_sum,0)) /1,0)  as loan_amt_sum_1,");//--金庫.授信案總金額 
	      sql.append("                    round(sum(decode(bank_no_max,'2',loan_amt_sum,0)) /1,0) as loan_amt_sum_2,");//--其他.授信案總金額
	      sql.append("                    round(sum(loan_amt_sum) /1,0) as loan_amt_sum_all,");//--合計.授信案總金額
	      sql.append("                    0,0,0,0,0,0,0,0,0,0,0,0,0 ");       
	      sql.append("                    from  (select * from cd01 where cd01.hsien_id <> 'Y'  ) cd01 ");
          sql.append("                    left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year=? and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL) ");  
          paramList.add(wlx01_m_year);
          sql.append("                    left join bn01 on wlx01.bank_no=bn01.bank_no and bn01.bank_type in ");
          if("".equals(bank_type)){
              sql.append("                    (?,?) ");
              paramList.add("6");
              paramList.add("7");
          }else{
              sql.append("                    (?) ");
              paramList.add(bank_type);
          }
          sql.append("                    and bn01.m_year=? and bn_type <> '2' and wlx01.m_year=? ");
          paramList.add(wlx01_m_year);
          paramList.add(wlx01_m_year);
          sql.append("                    left join( ");
          sql.append("                            select wlx10_m_loan.* from wlx10_m_loan, "); 
          sql.append("                                 (select m_year,m_month,bank_no,case_grp,case_no ");
          sql.append("                                  from ( ");
          sql.append("                                  select  m_year,m_month,bank_no,substr(case_no,0,7)  as case_grp,min(case_no) as case_no ");
          sql.append("                                  from wlx10_m_loan ");
          sql.append("                                  where m_year=? and m_month=? ");
          sql.append("                                  group by m_year,m_month,bank_no,substr(case_no,0,7) ))a ");
          paramList.add(m_year);
          paramList.add(m_month);
          sql.append("                            where to_char(wlx10_m_loan.m_year * 100 + wlx10_m_loan.m_month) = ?");
          paramList.add(m_year+m_month);
          sql.append("                                    and (wlx10_m_loan.m_year= a.m_year ");
          sql.append("                                    and wlx10_m_loan.m_month=a.m_month ");
          sql.append("                                    and wlx10_m_loan.bank_no = a.bank_no ");
          sql.append("                                    and wlx10_m_loan.case_no = a.case_no) ");                    
          sql.append("                              ) wlx10_m_loan  on  bn01.bank_no = wlx10_m_loan.bank_no ");
          sql.append("                     group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME,substr(case_no,0,7) "); 
          sql.append("                   )group by hsien_id,hsien_name,FR001W_output_order,bank_no,BANK_NAME,case_no ");
	      sql.append("               ) wlx10_m_loan ");             
	      sql.append("            where wlx10_m_loan.bank_no <> ' ' "); 
	      sql.append("            GROUP BY wlx10_m_loan.hsien_id,wlx10_m_loan.hsien_name,wlx10_m_loan.FR001W_output_order,wlx10_m_loan.bank_no,wlx10_m_loan.BANK_NAME");
	      sql.append("        ) wlx10_m_loan ");//--各別機構明細  
	      sql.append("        UNION ALL ");
	      sql.append("        select wlx10_m_loan.hsien_id,wlx10_m_loan.hsien_name,wlx10_m_loan.FR001W_output_order,wlx10_m_loan.bank_no,wlx10_m_loan.BANK_NAME, ");    
	      sql.append("          COUNT_SEQ, field_SEQ, "); 
	      sql.append("          round(loan_amt_sum_1 /?,0) as loan_amt_sum_1,round(loan_amt_sum_2 /?,0) as loan_amt_sum_2,round(loan_amt_sum_all /?,0) as loan_amt_sum_all, ");
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);            
	      sql.append("          round(loan_amt_1 /?,0) as loan_amt_1, round(loan_amt_2 /?,0) as loan_amt_2, round(loan_amt_all /?,0) as loan_amt_all, ");
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);
	      sql.append("          round(loan_bal_amt_1 /?,0) as loan_bal_amt_1, round(loan_bal_amt_2 /?,0) as loan_bal_amt_2, round(loan_bal_amt_all /?,0) as loan_bal_amt_all, ");
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);
	      sql.append("          round(field_OVER_1 /?,0) as field_OVER_1, round(field_OVER_2 /?,0) as field_OVER_2, round(field_OVER_all /?,0) as field_OVER_all, ");
	      paramList.add(unit);
	      paramList.add(unit);
	      paramList.add(unit);
	      sql.append("          decode(loan_bal_amt_1,0,0,round(field_OVER_1 / loan_bal_amt_1 *100 ,2)) as field_OVER_RATE_1, ");        
	      sql.append("          decode(loan_bal_amt_2,0,0,round(field_OVER_2 / loan_bal_amt_2 *100 ,2)) as field_OVER_RATE_2, ");   
	      sql.append("          decode(loan_bal_amt_all,0,0,round(field_OVER_all / loan_bal_amt_all *100 ,2)) as field_OVER_RATE_all, ");        
	      sql.append("          round(loan_type_amt_1 /?,0)  as loan_type_amt_1,round(loan_type_amt_2 /?,0)  as loan_type_amt_2, ");   
	      sql.append("          round(loan_type_amt_3 /?,0)  as loan_type_amt_3, round(loan_type_all /?,0)  as loan_type_all ");   
	      paramList.add(unit);
          paramList.add(unit);
          paramList.add(unit);
          paramList.add(unit);
	      sql.append("    from ");
	      sql.append("        (");//--縣市小計
	      sql.append("         select wlx10_m_loan.hsien_id, wlx10_m_loan.hsien_name,  wlx10_m_loan.FR001W_output_order, ' ' AS  bank_no , ' ' AS  BANK_NAME, "); 
	      sql.append("           count(case_no) AS COUNT_SEQ, 'A90' as field_SEQ, ");       
	      sql.append("           SUM(loan_amt_sum_1) loan_amt_sum_1, SUM(loan_amt_sum_2) loan_amt_sum_2, SUM(loan_amt_sum_all) loan_amt_sum_all, ");     
	      sql.append("           SUM(loan_amt_1) loan_amt_1,  SUM(loan_amt_2) loan_amt_2, SUM(loan_amt_all) loan_amt_all, ");
	      sql.append("           SUM(loan_bal_amt_1) loan_bal_amt_1, SUM(loan_bal_amt_2) loan_bal_amt_2, SUM(loan_bal_amt_all) loan_bal_amt_all, ");
	      sql.append("           SUM(field_OVER_1) field_OVER_1, SUM(field_OVER_2) field_OVER_2, SUM(field_OVER_all) field_OVER_all, ");              
	      sql.append("           SUM(loan_type_amt_1)  loan_type_amt_1,  SUM(loan_type_amt_2)  loan_type_amt_2, "); 
	      sql.append("           SUM(loan_type_amt_3)  loan_type_amt_3,    SUM(loan_type_all)  loan_type_all ");
	      sql.append("         from ( ");        
	      sql.append("              select hsien_id,hsien_name,FR001W_output_order,bank_no,BANK_NAME,case_no, ");
	      sql.append("                     SUM(loan_amt_sum_1) loan_amt_sum_1 ,SUM(loan_amt_sum_2)    loan_amt_sum_2,   SUM(loan_amt_sum_all)  loan_amt_sum_all, ");     
	      sql.append("                     SUM(loan_amt_1)   loan_amt_1,   SUM(loan_amt_2)  loan_amt_2, SUM(loan_amt_all)  loan_amt_all, ");
	      sql.append("                     SUM(loan_bal_amt_1)   loan_bal_amt_1,   SUM(loan_bal_amt_2)  loan_bal_amt_2, SUM(loan_bal_amt_all)  loan_bal_amt_all, ");
	      sql.append("                     SUM(field_OVER_1)   field_OVER_1,   SUM(field_OVER_2)  field_OVER_2, SUM(field_OVER_all)  field_OVER_all, ");
	      sql.append("                     SUM(loan_type_amt_1)  loan_type_amt_1,  SUM(loan_type_amt_2)  loan_type_amt_2, "); 
          sql.append("                     SUM(loan_type_amt_3)  loan_type_amt_3,    SUM(loan_type_all)  loan_type_all ");         
	      sql.append("              from (");
	      sql.append("                select nvl(cd01.hsien_id,' ') as hsien_id, nvl(cd01.hsien_name,'OTHER') as hsien_name, ");              
	      sql.append("                       cd01.FR001W_output_order as FR001W_output_order, bn01.bank_no, bn01.BANK_NAME,substr(case_no,0,7) as case_no, ");
	      sql.append("                       0 as loan_amt_sum_1,");//--金庫.授信案總金額 
	      sql.append("                       0 as loan_amt_sum_2,");//--其他.授信案總金額
	      sql.append("                       0 as loan_amt_sum_all,");//--合計.授信案總金額
	      //sql.append("                       round(sum(decode(bank_no_max,'1',decode(substr(case_no,9,1),'1',loan_amt_sum,0),0)) /1,0) as loan_amt_sum_1, "); //--金庫.授信案總金額 
          //sql.append("                       round(sum(decode(bank_no_max,'2',decode(substr(case_no,9,1),'1',loan_amt_sum,0),0)) /1,0) as loan_amt_sum_2, "); //--其他.授信案總金額
          //sql.append("                       round(sum(decode(substr(case_no,9,1),'1' ,loan_amt_sum,0)) /1,0) as loan_amt_sum_all, "); //--合計.授信案總金額
          sql.append("                       round(sum(decode(bank_no_max,'1',loan_amt,0)) /1,0) as loan_amt_1, "); //--金庫.參貸額度
          sql.append("                       round(sum(decode(bank_no_max,'2',loan_amt,0)) /1,0) as loan_amt_2, "); //--其他.參貸額度
          sql.append("                       round(sum(loan_amt) /1,0) as loan_amt_all, "); //--合計.參貸額度
          sql.append("                       round(sum(decode(bank_no_max,'1',loan_bal_amt,0)) /1,0) as loan_bal_amt_1, "); //--金庫.實際授信餘額
          sql.append("                       round(sum(decode(bank_no_max,'2',loan_bal_amt,0)) /1,0) as loan_bal_amt_2, "); //--其他.實際授信餘額
          sql.append("                       round(sum(loan_bal_amt) /1,0) as loan_bal_amt_all, "); //--合計.實際授信餘額
          sql.append("                       round(sum(decode(pay_state,'2',decode(bank_no_max,'1',loan_bal_amt,0),0)) /1,0) as field_OVER_1, "); //--金庫.逾放金額
          sql.append("                       round(sum(decode(pay_state,'2',decode(bank_no_max,'2',loan_bal_amt,0),0)) /1,0) as field_OVER_2, "); //--其他.逾放金額
          sql.append("                       round(sum(decode(pay_state,'2',loan_bal_amt,0)) /1,0) as field_OVER_all, "); //--合計.逾放金額                         
          sql.append("                       round(sum(decode(loan_type,'1',loan_bal_amt,0)) /1,0) as loan_type_amt_1, "); //--授信用途.購地.實際授信餘額
          sql.append("                       round(sum(decode(loan_type,'2',loan_bal_amt,0)) /1,0) as loan_type_amt_2, "); //--授信用途.建築.實際授信餘額
          sql.append("                       round(sum(decode(loan_type,'3',loan_bal_amt,0)) /1,0) as loan_type_amt_3, "); //--授信用途.其他.實際授信餘額
          sql.append("                       round(sum(decode(loan_type,'1',loan_bal_amt,'2',loan_bal_amt,'3',loan_bal_amt,0)) /1,0) as  loan_type_all "); //--授信用途.合計.實際授信餘額
          sql.append("                from  (select * from  cd01 where cd01.hsien_id <> 'Y' ) cd01 ");
	      sql.append("                left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year=? and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL) ");
	      paramList.add(wlx01_m_year);
	      sql.append("                    left join bn01 on wlx01.bank_no=bn01.bank_no and bn01.bank_type in ");
          if("".equals(bank_type)){
              sql.append("                    (?,?) ");
              paramList.add("6");
              paramList.add("7");
          }else{
              sql.append("                    (?) ");
              paramList.add(bank_type);
          }
          sql.append("                    and bn01.m_year=? and bn_type <> '2' and wlx01.m_year=? ");
	      paramList.add(wlx01_m_year);
	      paramList.add(wlx01_m_year);
	      sql.append("                left join (select * from wlx10_m_loan where to_char(m_year * 100 + m_month) =?) wlx10_m_loan  on  bn01.bank_no = wlx10_m_loan.bank_no");  
	      paramList.add(m_year+m_month);
	      sql.append("                group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME,substr(case_no,0,7)");
	      sql.append("                union"); 
	      sql.append("                select nvl(cd01.hsien_id,' ')       as  hsien_id ,");               
	      sql.append("                       nvl(cd01.hsien_name,'OTHER') as  hsien_name,");               
	      sql.append("                       cd01.FR001W_output_order     as  FR001W_output_order,");               
	      sql.append("                       bn01.bank_no ,  bn01.BANK_NAME,  substr(case_no,0,7) as case_no,");        
	      sql.append("                       round(sum(decode(bank_no_max,'1',loan_amt_sum,0)) /1,0)  as loan_amt_sum_1,");//--金庫.授信案總金額 
	      sql.append("                       round(sum(decode(bank_no_max,'2',loan_amt_sum,0)) /1,0) as loan_amt_sum_2,");//--其他.授信案總金額
	      sql.append("                       round(sum(loan_amt_sum) /1,0) as loan_amt_sum_all,");//--合計.授信案總金額
	      sql.append("                       0,0,0,0,0,0,0,0,0,0,0,0,0 ");
	      sql.append("                from  (select * from cd01 where cd01.hsien_id <> 'Y'  ) cd01 ");
          sql.append("                left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year=? and (wlx01.cancel_no <> 'Y' OR wlx01.cancel_no IS NULL) ");  
          paramList.add(wlx01_m_year);
          sql.append("                left join bn01 on wlx01.bank_no=bn01.bank_no and bn01.bank_type in ");
          if("".equals(bank_type)){
              sql.append(" (?,?) ");
              paramList.add("6");
              paramList.add("7");
          }else{
              sql.append(" (?) ");
              paramList.add(bank_type);
          }
          sql.append("                and bn01.m_year=? and bn_type <> '2' and wlx01.m_year=? ");
          paramList.add(wlx01_m_year);
          paramList.add(wlx01_m_year);
          sql.append("                left join( ");
          sql.append("                         select wlx10_m_loan.* from wlx10_m_loan, "); 
          sql.append("                            (select m_year,m_month,bank_no,case_grp,case_no ");
          sql.append("                            from ( ");
          sql.append("                                  select  m_year,m_month,bank_no,substr(case_no,0,7)  as case_grp,min(case_no) as case_no ");
          sql.append("                                  from wlx10_m_loan ");
          sql.append("                                  where m_year=? and m_month=? ");
          sql.append("                                  group by m_year,m_month,bank_no,substr(case_no,0,7) ))a ");
          paramList.add(m_year);
          paramList.add(m_month);
          sql.append("                            where to_char(wlx10_m_loan.m_year * 100 + wlx10_m_loan.m_month) = ?");
          paramList.add(m_year+m_month);
          sql.append("                                    and (wlx10_m_loan.m_year= a.m_year ");
          sql.append("                                    and wlx10_m_loan.m_month=a.m_month ");
          sql.append("                                    and wlx10_m_loan.bank_no = a.bank_no ");
          sql.append("                                    and wlx10_m_loan.case_no = a.case_no) ");                    
          sql.append("                           ) wlx10_m_loan  on  bn01.bank_no = wlx10_m_loan.bank_no ");
          sql.append("                  group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order,bn01.bank_no,bn01.BANK_NAME,substr(case_no,0,7) "); 
          sql.append("                 )group by hsien_id,hsien_name,FR001W_output_order,bank_no,BANK_NAME,case_no ");	      
	      sql.append("              ) wlx10_m_loan");
	      sql.append("         where  wlx10_m_loan.bank_no <> ' ' "); 
	      sql.append("         GROUP BY wlx10_m_loan.hsien_id ,wlx10_m_loan.hsien_name,wlx10_m_loan.FR001W_output_order ");
	      sql.append("     ) wlx10_m_loan ");//--縣市小計end
	      sql.append("  )  wlx10_m_loan ");
	      sql.append("  left join (select * from cd01 where cd01.hsien_id <> 'Y') cd01 on wlx10_m_loan.hsien_id=cd01.hsien_id ");
	      sql.append("ORDER by FR001W_output_order, field_SEQ, hsien_id, bank_no ");
	      dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"hsien_id,hsien_name,FR001W_output_order,bank_no,BANK_NAME,COUNT_SEQ,field_SEQ," +
                                                        	      		"loan_amt_sum_1,loan_amt_sum_2,loan_amt_sum_all,loan_amt_1,loan_amt_2," +
                                                        	      		"loan_amt_all,loan_bal_amt_1,loan_bal_amt_2,loan_bal_amt_all," +
                                                        	      		"field_OVER_1,field_OVER_2,field_OVER_all,field_OVER_RATE_1,field_OVER_RATE_2,field_OVER_RATE_all," +
                                                        	      		"loan_type_amt_1,loan_type_amt_2,loan_type_amt_3,loan_type_all");
	      System.out.println("dbData.size=" + dbData.size());	      
	      //設定報表表頭資料============================================
	   	  	   	  
	      row = sheet.getRow(0);
          cell = row.getCell( (short) 0);
          cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	      if (dbData != null && dbData.size() != 0) {
	          if("".equals(bank_type)){
	              cell.setCellValue(m_year+"年"+m_month+"月"+"全體農漁會信用部聯合貸款案件彙總表");
	          }else{
	              cell.setCellValue(m_year+"年"+m_month+"月"+"全體"+bank_type_name+"信用部聯合貸款案件彙總表");
	          }
		   	  row = sheet.getRow(1);
		   	  cell = row.getCell( (short) 11);	   	  
		   	  cell.setEncoding(HSSFCell.ENCODING_UTF_16);		   	
		   	  cell.setCellValue("單位:新臺幣" + unit_name + "、%");  
		   	  rowNum = 3;	      	  
	      	  for(int i=1;i<=dbData.size();i++){	
	      	      if(i==dbData.size()){
	      	        bean = (DataObject)dbData.get(0);
	      	      }else{
	      	        bean = (DataObject)dbData.get(i);
	      	      }
	      	      count_seq = (bean.getValue("count_seq")==null)?"":(bean.getValue("count_seq")).toString();
	      	      bank_no = (bean.getValue("bank_no")==null)?"":(bean.getValue("bank_no")).toString(); //--機構代號 
	      	      bank_name = String.valueOf(bean.getValue("bank_name")); 
	      	      hsien_name = String.valueOf(bean.getValue("hsien_name"));
	      	      field_seq = String.valueOf(bean.getValue("field_seq"));
	      	      loan_amt_sum_1 = (bean.getValue("loan_amt_sum_1") == null)?"0":(bean.getValue("loan_amt_sum_1")).toString();  //--金庫.授信案總金額
	              loan_amt_sum_2 = (bean.getValue("loan_amt_sum_2") == null)?"0":(bean.getValue("loan_amt_sum_2")).toString(); //--其他.授信案總金額
	              loan_amt_sum_all = (bean.getValue("loan_amt_sum_all") == null)?"0":(bean.getValue("loan_amt_sum_all")).toString(); //--合計.授信案總金額
	              loan_amt_1 = (bean.getValue("loan_amt_1") == null)?"0":(bean.getValue("loan_amt_1")).toString(); // --金庫.參貸額度             
	              loan_amt_2 = (bean.getValue("loan_amt_2") == null)?"0":(bean.getValue("loan_amt_2")).toString(); //--其他.參貸額度
	              loan_amt_all = (bean.getValue("loan_amt_all") == null)?"0":(bean.getValue("loan_amt_all")).toString(); //--合計.參貸額度
	              loan_bal_amt_1 = (bean.getValue("loan_bal_amt_1") == null)?"0":(bean.getValue("loan_bal_amt_1")).toString(); // --金庫.實際授信餘額
	              loan_bal_amt_2 = (bean.getValue("loan_bal_amt_2") == null)?"0":(bean.getValue("loan_bal_amt_2")).toString(); //--其他.實際授信餘額
	              loan_bal_amt_all = (bean.getValue("loan_bal_amt_all") == null)?"0":(bean.getValue("loan_bal_amt_all")).toString(); //--合計.實際授信餘額
	              field_OVER_1 = (bean.getValue("field_over_1") == null)?"0":(bean.getValue("field_over_1")).toString(); //--金庫.逾放金額
	              field_OVER_2 = (bean.getValue("field_over_2") == null)?"0":(bean.getValue("field_over_2")).toString(); //--其他.逾放金額
	              field_OVER_all = (bean.getValue("field_over_all") == null)?"0":(bean.getValue("field_over_all")).toString(); //--合計.逾放金額
	              field_OVER_RATE_1 = (bean.getValue("field_over_rate_1") == null)?"0":(bean.getValue("field_over_rate_1")).toString(); //--金庫.逾放金額占實際授信餘額比率
	              field_OVER_RATE_2 = (bean.getValue("field_over_rate_2") == null)?"0":(bean.getValue("field_over_rate_2")).toString(); //--其他.逾放金額占實際授信餘額比率
	              field_OVER_RATE_all = (bean.getValue("field_over_rate_all") == null)?"0":(bean.getValue("field_over_rate_all")).toString(); //--合計.逾放金額占實際授信餘額比率
	              loan_type_amt_1 = (bean.getValue("loan_type_amt_1") == null)?"0":(bean.getValue("loan_type_amt_1")).toString(); //--授信用途.購地.實際授信餘額
	              loan_type_amt_2 = (bean.getValue("loan_type_amt_2") == null)?"0":(bean.getValue("loan_type_amt_2")).toString(); //--授信用途.建築.實際授信餘額
	              loan_type_amt_3 = (bean.getValue("loan_type_amt_3") == null)?"0":(bean.getValue("loan_type_amt_3")).toString(); //--授信用途.其他.實際授信餘額
	              loan_type_all = (bean.getValue("loan_type_all") == null)?"0":(bean.getValue("loan_type_all")).toString(); //--授信用途.合計.實際授信餘額
	      	      rowNum++;
    	      	  row = sheet.createRow(rowNum);
   	    	  	  //列印各機構明細資料
	      	      for(int cellcount=0;cellcount<22;cellcount++){			 	      
	      	          cell=row.createCell((short)cellcount);			 		
	      	          cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	      	          if(cellcount>=2){
	      	              cell.setCellStyle(cs_right);
	      	          }else{
	      	              cell.setCellStyle(cs_left);
	      	          }
	      	          if(cellcount == 0) cell.setCellValue(bank_no);
	      	          if(cellcount == 1){
	      	              if(!"A01".equals(field_seq)){ 
	      	                  cell.setCellValue(hsien_name);
	      	              }else{
	      	                  cell.setCellValue(bank_name);
	      	              }
	      	          }
	      	          if(cellcount == 2) cell.setCellValue(count_seq);
	      	          if(cellcount == 3) cell.setCellValue(Utility.setCommaFormat(loan_amt_sum_1));
	      	          if(cellcount == 4) cell.setCellValue(Utility.setCommaFormat(loan_amt_sum_2));
	      	          if(cellcount == 5) cell.setCellValue(Utility.setCommaFormat(loan_amt_sum_all));		    	  
	      	          if(cellcount == 6) cell.setCellValue(Utility.setCommaFormat(loan_amt_1));
	      	          if(cellcount == 7) cell.setCellValue(Utility.setCommaFormat(loan_amt_2));
	      	          if(cellcount == 8) cell.setCellValue(Utility.setCommaFormat(loan_amt_all));
	      	          if(cellcount == 9) cell.setCellValue(Utility.setCommaFormat(loan_bal_amt_1));
	      	          if(cellcount == 10) cell.setCellValue(Utility.setCommaFormat(loan_bal_amt_2));
	      	          if(cellcount == 11) cell.setCellValue(Utility.setCommaFormat(loan_bal_amt_all));
	      	          if(cellcount == 12) cell.setCellValue(Utility.setCommaFormat(loan_type_amt_1));
	      	          if(cellcount == 13) cell.setCellValue(Utility.setCommaFormat(loan_type_amt_2));
	      	          if(cellcount == 14) cell.setCellValue(Utility.setCommaFormat(loan_type_amt_3));
	      	          if(cellcount == 15) cell.setCellValue(Utility.setCommaFormat(loan_type_all));
	      	          if(cellcount == 16) cell.setCellValue(Utility.setCommaFormat(field_OVER_1));	 
	      	          if(cellcount == 17) cell.setCellValue(Utility.setCommaFormat(field_OVER_2));
	      	          if(cellcount == 18) cell.setCellValue(Utility.setCommaFormat(field_OVER_all));
	      	          if(cellcount == 19) cell.setCellValue(field_OVER_RATE_1);              
	      	          if(cellcount == 20) cell.setCellValue(field_OVER_RATE_2);
	      	          if(cellcount == 21) cell.setCellValue(field_OVER_RATE_all);
	      	          
	      	      }//end of cellcount
	      	  }
	      	
            
	      }else{ //end of else dbData.size() != 0
	          if("".equals(bank_type)){
                  cell.setCellValue(m_year+"年"+m_month+"月"+"全體農漁會信用部聯合貸款案件彙總表   無資料存在");
              }else{
                  cell.setCellValue(m_year+"年"+m_month+"月"+"全體"+bank_type_name+"信用部聯合貸款案件彙總表   無資料存在");
              }
	      }
	     
	      if(febxlsFlag.equals("")){//原全體農漁會信用部各會員別放款金額一覽表	      	
	         FileOutputStream fout = null;     
	         fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "農漁會信用部聯合貸款案件彙總表.xls");
	     
	         HSSFFooter footer = sheet.getFooter();
	         footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
	         footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
	         wb.write(fout);
	         //儲存
	         fout.close();
	         System.out.println("儲存成功!");
	      }
	    }catch (Exception e) {
	      System.out.println("RptFR064W.createRpt Error:" + e + e.getMessage());
	    }
	    
	    return errMsg;
	  }
}
