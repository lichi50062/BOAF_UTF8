/*
 * 103.03.17 created by 2968  
 *           1.結合RptWR003W、RptWR004W、RptWR005W 
 *           2.fix (1)各項警訊指標皆調整為5名；
 *                 (2)最後一列增加顯示警示項目加總及各項數
 *                 (3)統計上月／上季／上年度同期，警訊指標皆相同，只有比較期間不同/顯示的title不同
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

public class RptWR098W {
    static String bank_code ="";
    static String bank_name ="";
    static String cnt ="";
    static String field_debit_rate_max ="";
    static String field_debit_rate_min ="";
    static String field_credit_rate_max ="";
    static String field_credit_rate_min ="";
    static String field_990610_rate_max ="";
    static String field_noassure_rate_max ="";
    static String field_992710_rate_max ="";
    static String fd_992710_990230_rate_100 =""; 
    static String diff_992710_990230_rate_max =""; 
    static String field_diff_over_max ="";
    static String field_diff_over_rate_max ="";
    static String field_diff_992530_max ="";
    static String fd_992550_992150_rate_max ="";
    static String fd_992720_992710_rate_max ="";
    static String field_diff_992610_cal_max ="";
    static String field_diff_992630_max ="";
    static String field_diff_992730_max ="";
    static String field_diff_backup_min ="";
    static String fd_diff_backup_credit_rate_min ="";
    static String diff_loan_bal_amt_max ="";
    static String diff_over6m_loan_bal_amt_max ="";
    static String diff_over6m_loan_rate_max ="";
    static String delay_loan_rate_max ="";
    static String wr_count ="0";
    static String bank_count ="0";
	static File logfile;
    static File logDir = null;
    static boolean isLoan = false;
    public static String createRpt(String szrpt_code,String s_year,String s_month,String unit,HSSFWorkbook wb){
          String errMsg = "";
          HSSFRow row=null;//宣告一列
          HSSFCell cell=null;//宣告一個儲存格
          HSSFSheet sheet =null;
          FileInputStream finput = null;
          StringBuffer sqlCmd = new StringBuffer();
          List paramList = new ArrayList();
          List dbData = null;
          DataObject bean = null;
          StringBuffer sqlCmd_sum = new StringBuffer();
          List dbData_sum = null;
          DataObject bean_sum = null;
          String wlx01_m_year = "";
          String unit_name = "";
          unit_name = Utility.getUnitName(unit);//取得單位名稱
          s_month = String.valueOf(Integer.parseInt(s_month));
          String last_year  = (Integer.parseInt(s_month) == 1) ? String.valueOf(Integer.parseInt(s_year) - 1) : s_year; 
          String last_month = (Integer.parseInt(s_month) == 1) ? "12" : String.valueOf(Integer.parseInt(s_month) - 1);
          String lastSeason_year  = (Integer.parseInt(s_month) == 3) ? String.valueOf(Integer.parseInt(s_year) - 1) : s_year; 
          String lastSeason_month = "";
          if(Integer.parseInt(s_month) <= 3 && Integer.parseInt(s_month)>=1){
              lastSeason_month = "12";
          }else if(Integer.parseInt(s_month) <= 6 && Integer.parseInt(s_month)>=4){
              lastSeason_month = "3";
          }else if(Integer.parseInt(s_month) <= 9 && Integer.parseInt(s_month)>=7){
              lastSeason_month = "6";
          }else if(Integer.parseInt(s_month) <= 12 && Integer.parseInt(s_month)>=10){
              lastSeason_month = "9";
          }
          String wr_rpt="";
          if("WR003W_ATOT".equals(szrpt_code)){
              wr_rpt="0";
          }else if("WR004W_ATOT".equals(szrpt_code)){
              wr_rpt="1";
          }else if("WR005W_ATOT".equals(szrpt_code)){
              wr_rpt="2";
          }else if("WR006W_ATOT".equals(szrpt_code)){
              wr_rpt="3";
          }else if("WR007W_ATOT".equals(szrpt_code)){
              wr_rpt="4";
          }else if("WR008W_ATOT".equals(szrpt_code)){
              wr_rpt="5";
          }
      try{
          logDir  = new File(Utility.getProperties("logDir"));
          if(!logDir.exists()){
              if(!Utility.mkdirs(Utility.getProperties("logDir"))){
                 System.out.println("目錄新增失敗");
              }    
          }
          logfile = new File(logDir + System.getProperty("file.separator") + "RptWR003W.log");
          
          System.out.println("logfile filename="+logDir + System.getProperty("file.separator") +"RptWR003W.log");
          
            //99.03.24 add 查詢年度100年以前.縣市別不同===============================
  	    	wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100"; 
  	    	//=====================================================================    
      		File xlsDir = new File(Utility.getProperties("xlsDir"));
            File reportDir = new File(Utility.getProperties("reportDir"));

            if(!xlsDir.exists()){
                if(!Utility.mkdirs(Utility.getProperties("xlsDir"))){
                    errMsg +=Utility.getProperties("xlsDir")+"目錄新增失敗";
                }
            }
            if(!reportDir.exists()){
                if(!Utility.mkdirs(Utility.getProperties("reportDir"))){
                   errMsg +=Utility.getProperties("reportDir")+"目錄新增失敗";
                }
            }
            String outputFile= "";
            if("WR003W_ATOT".equals(szrpt_code)){
                outputFile= "農漁會信用部營運狀況警訊報表-與上月比較.xls";
            }else if("WR004W_ATOT".equals(szrpt_code)){
                outputFile= "農漁會信用部營運狀況警訊報表-與上季比較.xls";
            }else if("WR005W_ATOT".equals(szrpt_code)){
                outputFile= "農漁會信用部營運狀況警訊報表-與上年度同期比較.xls";
            }else if("WR006W_ATOT".equals(szrpt_code)){
                outputFile= "農漁會信用部營運狀況警訊報表(專案農貸)-與上月比較.xls";
            }else if("WR007W_ATOT".equals(szrpt_code)){
                outputFile= "農漁會信用部營運狀況警訊報表(專案農貸)-與上季比較.xls";
            }else if("WR008W_ATOT".equals(szrpt_code)){
                outputFile= "農漁會信用部營運狀況警訊報表(專案農貸)-與上年度同期比較.xls";
            }
            String openfile="";
            if("WR003W_ATOT".equals(szrpt_code)||"WR004W_ATOT".equals(szrpt_code)||"WR005W_ATOT".equals(szrpt_code)){
                openfile="農漁會信用部營運狀況警訊報表-與上月_上季_上年度同期.xls";
                isLoan = false;
            }else{
                openfile="農漁會信用部營運狀況警訊報表_專案農貸-與上月_上季_上年度同期.xls";
                isLoan = true;
            }
            System.out.println("開啟檔:" + openfile);
            finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+ openfile );
            //設定FileINputStream讀取Excel檔
            POIFSFileSystem fs = new POIFSFileSystem( finput );
            if(fs==null){System.out.println("open 範本檔失敗");} else System.out.println("open 範本檔成功");
            wb = new HSSFWorkbook(fs);
            if(wb==null){System.out.println("open工作表失敗");}else System.out.println("open 工作表 成功");
	  		//讀取第一個工作表，宣告其為sheet
	  		sheet = wb.getSheetAt(0);
	  		//sheet = wb.getSheetAt((bank_type.equals("6")?0:1));
	  		if(sheet==null){System.out.println("open sheet 失敗");}else System.out.println("open sheet 成功");
	  		HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
	        //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	        //sheet.setAutobreaks(true); //自動分頁

	        //設定頁面符合列印大小
	        sheet.setAutobreaks( false );
	        /*
	        if(isLoan){
                ps.setScale( ( short )100 );//列印縮放百分比
            }else{
                ps.setScale( ( short )76 );
            }
            */
	        //ps.setPaperSize( ( short )8 ); //設定紙張大小 A3
	  		//wb.setSheetName(0,"test");
	  		//設定表頭 為固定 先設欄的起始再設列的起始
	        //wb.setRepeatingRowsAndColumns(0, 1, 21, 2, 3);
	        if(isLoan){
	            sqlCmd.append("select bank_code,");//--機構代號
	            sqlCmd.append("       bn01.bank_name,");//--機構名稱
	            sqlCmd.append("       count(*) as wr_count,");//--警示筆數
	            sqlCmd.append("       sum(decode(acc_code,'field_diff_loan_bal_amt',decode(wr_range_serial,1,round(amt /?,0),''),'')) as diff_loan_bal_amt_max,");//--專案農貸放款餘額.增加金額前5名
	            sqlCmd.append("       sum(decode(acc_code,'field_diff_over6m_loan_bal_amt',decode(wr_range_serial,2,round(amt /?,0),''),'')) as diff_over6m_loan_bal_amt_max,");//--專案農貸逾期放款.增加金額前5名
	            sqlCmd.append("       sum(decode(acc_code,'field_diff_over6m_loan_rate',decode(wr_range_serial,3,amt,''),'')) as diff_over6m_loan_rate_max,");//--專案農貸逾放比率.增加百分點前5名
	            sqlCmd.append("       sum(decode(acc_code,'field_delay_loan_rate',decode(wr_range_serial,4,amt,''),'')) as delay_loan_rate_max ");//--當年度累計核准延期還款件數占尚有餘額專案農貸總件數比率前5名
	            sqlCmd.append("  from wr_operation left join (select * from bn01 where m_year=?)bn01 on wr_operation.bank_code=bn01.bank_no ");
	            sqlCmd.append(" where wr_operation.m_year=? ");
	            sqlCmd.append("   and m_month=? ");
	            sqlCmd.append("   and wr_rpt=? ");//3:上月 4:上季 5:上年度同期
	            sqlCmd.append("   and warn_type=? "); 
	            sqlCmd.append(" group by bank_code,bank_name ");
	            sqlCmd.append(" order by bank_code ");
                paramList.add(unit);
                paramList.add(unit);
                paramList.add(wlx01_m_year);
                paramList.add(s_year);
                paramList.add(s_month);
                paramList.add(wr_rpt);
	            paramList.add("Y");
	            dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,
	                      "bank_code,bank_name,wr_count,diff_loan_bal_amt_max,asdiff_over6m_loan_bal_amt_max,diff_over6m_loan_rate_max,delay_loan_rate_max");
	        }else{
	            sqlCmd.append(" select bank_code,");//--機構代號
	            sqlCmd.append("        bn01.bank_name,");//--機構名稱
	            sqlCmd.append("        count(*) wr_count,");//警示筆數 102.08.26 add  
	            sqlCmd.append("        sum(decode(acc_code,'field_debit_rate',decode(wr_range_serial,1,amt,''),'')) as field_debit_rate_max,");//--存款總額.增加比率前5名
	            sqlCmd.append("        sum(decode(acc_code,'field_debit_rate',decode(wr_range_serial,2,amt,''),'')) as field_debit_rate_min,");//--存款總額.減少比率前5名
	            sqlCmd.append("        sum(decode(acc_code,'field_credit_rate',decode(wr_range_serial,3,amt,''),'')) as field_credit_rate_max,");//--放款總額.增加比率前5名
	            sqlCmd.append("        sum(decode(acc_code,'field_credit_rate',decode(wr_range_serial,4,amt,''),'')) as field_credit_rate_min,");//--放款總額.減少比率前5名
	            sqlCmd.append("        sum(decode(acc_code,'field_990610_rate',decode(wr_range_serial,5,amt,''),'')) as field_990610_rate_max,");//--非會員放款.增加比率前5名
	            //sqlCmd.append("        sum(decode(acc_code,'field_990610_rate',decode(wr_range_serial,6,amt,''),'')) as field_990610_rate_min,");//--非會員放款.減少比率前5名 103.02.26 取消
	            sqlCmd.append("        sum(decode(acc_code,'field_noassure_rate',decode(wr_range_serial,7,amt,''),'')) as field_noassure_rate_max,");//--無擔保放款.增加比率前5名
	            //sqlCmd.append("        sum(decode(acc_code,'field_noassure_rate',decode(wr_range_serial,8,amt,''),'')) as field_noassure_rate_min,");//--無擔保放款.減少比率前5名 103.02.26 取消
	            //sqlCmd.append("        sum(decode(acc_code,'field_120700_rate',decode(wr_range_serial,9,amt,''),'')) as field_120700_rate_max,");//--內部融資.增加比率前5名 103.02.26 取消
	            //sqlCmd.append("        sum(decode(acc_code,'field_992710',decode(wr_range_serial,10, round(amt /?,0) ,''),'')) as field_992710_max,");//--建築放款.前10名 103.02.26 取消
	            sqlCmd.append("        sum(decode(acc_code,'field_992710_rate',decode(wr_range_serial,11,amt,''),'')) as field_992710_rate_max,");//--建築放款.增加比率前10名
	            //sqlCmd.append("        sum(decode(acc_code,'field_992710_credit_rate',decode(wr_range_serial,12,amt,''),'')) as fd_992710_credit_rate_max,");//--建築放款/放款.前10名 103.02.26 取消
	            sqlCmd.append("        sum(decode(acc_code,'field_992710_990230_rate',decode(wr_range_serial,13,amt,''),'')) as fd_992710_990230_rate_100,");//--建築放款/上年度信用部決算淨值>=100%
	            sqlCmd.append("        sum(decode(acc_code,'field_diff_992710_990230_rate',decode(wr_range_serial,14,amt,''),'')) as diff_992710_990230_rate_max,");//--本月份之(建築放款/上年度信用部決算淨值)-上月份之(建築放款/上年度信用部決算淨值)增加百分點前10名
	            //sqlCmd.append("        sum(decode(acc_code,'field_over',decode(wr_range_serial,15,round(amt /?,0),''),'')) as field_over_max,");//--逾期放款本月前10名 103.02.26 取消
	            sqlCmd.append("        sum(decode(acc_code,'field_diff_over',decode(wr_range_serial,16,round(amt /?,0),''),'')) as field_diff_over_max,");//--逾期放款增加金額前10名
	            //sqlCmd.append("        sum(decode(acc_code,'field_over_rate',decode(wr_range_serial,17,amt,''),'')) as field_over_rate_max,");//--'逾放比率本月最高前10名 103.02.26 取消
	            sqlCmd.append("        sum(decode(acc_code,'field_diff_over_rate',decode(wr_range_serial,18,amt,''),'')) as field_diff_over_rate_max,");//--本月份逾放比率-上月份逾放比率增加百分點前10名
	            sqlCmd.append("        sum(decode(acc_code,'field_diff_992530',decode(wr_range_serial,19,round(amt /?,0),''),'')) as field_diff_992530_max,");//--逾放-非會員.增加金額前10名
	            sqlCmd.append("        sum(decode(acc_code,'field_992550_992150_rate',decode(wr_range_serial,20,amt,''),'')) as fd_992550_992150_rate_max,");//--無擔保消費性放款中之逾放/無擔保消費性放款.前5名
	            sqlCmd.append("        sum(decode(acc_code,'field_992720_992710_rate',decode(wr_range_serial,21,amt,''),'')) as fd_992720_992710_rate_max,");//--建築放款中之逾放/建築放款.前10名
	            //sqlCmd.append("        sum(decode(acc_code,'field_992610_cal',decode(wr_range_serial,22,round(amt /?,0),''),'')) as field_992610_cal_max,");//--應予觀察放款.前10名 103.02.26 取消
	            sqlCmd.append("        sum(decode(acc_code,'field_diff_992610_cal',decode(wr_range_serial,23,round(amt /?,0),''),'')) as field_diff_992610_cal_max,");//--應予觀察放款.增加金額前10名
	            //sqlCmd.append("        sum(decode(acc_code,'field_992610_cal_credit_rate',decode(wr_range_serial,24,amt,''),'')) as fd_992610_cal_credit_rate_max,");//--應予觀察放款/放款.前10名 103.02.26 取消
	            sqlCmd.append("        sum(decode(acc_code,'field_diff_992630',decode(wr_range_serial,25,round(amt /?,0),''),'')) as field_diff_992630_max,");//--應予觀察放款-非會員.增加金額前5名
	            sqlCmd.append("        sum(decode(acc_code,'field_diff_992730',decode(wr_range_serial,26,round(amt /?,0),''),'')) as field_diff_992730_max,");//--應予觀察放款-建築放款.增加金額前5名
	            sqlCmd.append("        sum(decode(acc_code,'field_diff_backup',decode(wr_range_serial,27,round(amt /?,0),''),'')) as field_diff_backup_min,");//--備抵呆帳.減少金額前10名
	            //sqlCmd.append("        sum(decode(acc_code,'field_backup_over_rate',decode(wr_range_serial,28,amt,''),'')) as backup_over_rate_min ");//--備抵呆帳/逾期放款.最低前10名 103.02.26 取消
	            sqlCmd.append("        sum(decode(acc_code,'field_diff_backup_credit_rate',decode(wr_range_serial,28,amt,''),'')) as fd_diff_backup_credit_rate_min ");//--本月份之放款覆蓋率-上月份放款覆蓋率(備抵呆帳/放款)(28).最低前5名103.03.12 add
	            sqlCmd.append("   from wr_operation left join (select * from bn01 where m_year=?)bn01 on wr_operation.bank_code=bn01.bank_no ");
	            sqlCmd.append("  where wr_operation.m_year=? and m_month=? ");
	            sqlCmd.append("    and warn_type='Y' ");
	            sqlCmd.append("    and wr_rpt=? ");//0:上月 1:上季 2:上年度同期
	            sqlCmd.append("    and wr_range_serial >0 ");
	            sqlCmd.append("  group by bank_code,bank_name ");
	            sqlCmd.append("  order by bank_code ");
	            //paramList.add(unit);
	            //paramList.add(unit);
	            //paramList.add(unit);
	            paramList.add(unit);
	            paramList.add(unit);
	            paramList.add(unit);
	            paramList.add(unit);
	            paramList.add(unit);
	            paramList.add(unit);
	            paramList.add(wlx01_m_year);
	            paramList.add(s_year);
	            paramList.add(s_month);
	            paramList.add(wr_rpt);
	            dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,
	                  "bank_code,bank_name,wr_count,field_debit_rate_max,field_debit_rate_min,field_credit_rate_max,"
	                +"field_credit_rate_min,field_990610_rate_max,field_noassure_rate_max,field_992710_rate_max,"
	                +"fd_992710_990230_rate_100,diff_992710_990230_rate_max,field_diff_over_max,field_diff_over_rate_max,"
	                +"field_diff_992530_max,fd_992550_992150_rate_max,fd_992720_992710_rate_max,field_diff_992610_cal_max,"
	                +"field_diff_992630_max,field_diff_992730_max,field_diff_backup_min,fd_diff_backup_credit_rate_min");
	        }
	        System.out.print("總表資料 共"+dbData.size()+"筆");
	        
	        sqlCmd_sum.append("select wr_count,");//--警示數目
            sqlCmd_sum.append("       count(*) as bank_count ");//--農漁會家數統計
            sqlCmd_sum.append("  from (").append(sqlCmd.toString()).append(")");
            sqlCmd_sum.append(" group by wr_count ");
            sqlCmd_sum.append(" order by wr_count desc ");
            dbData_sum = DBManager.QueryDB_SQLParam(sqlCmd_sum.toString(),paramList,"wr_count,bank_count");
            System.out.print("家數統計.size()="+dbData_sum.size());
            
	  		//建表開始--------------------------------------
	  		HSSFFont ft = wb.createFont();
            HSSFCellStyle cs = wb.createCellStyle();
            ft.setFontHeightInPoints((short)12);
            ft.setFontName("標楷體");
            cs.setFont(ft);
            cs.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
            cs.setBorderTop(HSSFCellStyle.BORDER_THIN);   
            cs.setBorderRight(HSSFCellStyle.BORDER_THIN); 
            cs.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);
            HSSFFont ft1 = wb.createFont();
            HSSFCellStyle cs1 = wb.createCellStyle();
            ft1.setFontHeightInPoints((short)12);
            ft1.setFontName("標楷體");
            cs1.setFont(ft1);
            cs1.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
            cs1.setBorderTop(HSSFCellStyle.BORDER_THIN);   
            cs1.setBorderRight(HSSFCellStyle.BORDER_THIN); 
            cs1.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            cs1.setAlignment(HSSFCellStyle.ALIGN_LEFT);
            HSSFFont ft2 = wb.createFont();
            HSSFCellStyle cs2 = wb.createCellStyle();
            ft2.setFontHeightInPoints((short)12);
            ft2.setFontName("標楷體");
            cs2.setFont(ft2);
            cs2.setBorderLeft(HSSFCellStyle.BORDER_THIN);  
            cs2.setBorderTop(HSSFCellStyle.BORDER_THIN);   
            cs2.setBorderRight(HSSFCellStyle.BORDER_THIN); 
            cs2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
            cs2.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
            
	  		row = sheet.getRow(0);
            cell = row.getCell( (short)0);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            String title = s_year+"年"+s_month+"月份農漁會信用部營運狀況警訊報表";
            if(isLoan)title +="（專案農貸部分）";
            cell.setCellValue(title);
            row = sheet.getRow(1);
            cell = row.getCell( (short)0);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            if("WR003W_ATOT".equals(szrpt_code) || "WR006W_ATOT".equals(szrpt_code)){
                cell.setCellValue("與上月份("+last_year+"年"+last_month+"月)比較");
            }else if("WR004W_ATOT".equals(szrpt_code) || "WR007W_ATOT".equals(szrpt_code)){
                cell.setCellValue("與上一季("+lastSeason_year+"年"+lastSeason_month+"月)比較");
            }else if("WR005W_ATOT".equals(szrpt_code) || "WR008W_ATOT".equals(szrpt_code)){
                cell.setCellValue("與上一年度同期("+String.valueOf(Integer.parseInt(s_year) - 1)+"年"+s_month+"月)比較");
            }
            row = sheet.getRow(2);
            cell = row.getCell( (short)0);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellValue("單位:新臺幣 " + unit_name + ",%,百分點");
            
            if(isLoan){
                row = sheet.getRow(7);
                cell = row.getCell( (short)3);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                if("WR006W_ATOT".equals(szrpt_code)){
                    cell.setCellValue("本月專案農貸放款餘額-上月專案農貸放款餘額");
                }else if("WR007W_ATOT".equals(szrpt_code)){
                    cell.setCellValue("本季專案農貸放款餘額-上季專案農貸放款餘額");
                }else if("WR008W_ATOT".equals(szrpt_code)){
                    cell.setCellValue("本年專案農貸放款餘額-上年專案農貸放款餘額");
                }
                cell = row.getCell( (short)5);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                if("WR006W_ATOT".equals(szrpt_code)){
                    cell.setCellValue("本月專案農貸逾放比率-上月專案農貸逾放比率");
                }else if("WR007W_ATOT".equals(szrpt_code)){
                    cell.setCellValue("本季專案農貸逾放比率-上季專案農貸逾放比率");
                }else if("WR008W_ATOT".equals(szrpt_code)){
                    cell.setCellValue("本年專案農貸逾放比率-上年專案農貸逾放比率");
                }
                
            }else{
                row = sheet.getRow(7);
                cell = row.getCell( (short)11);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                if("WR003W_ATOT".equals(szrpt_code)){
                    cell.setCellValue("本("+s_month+")月份之(建築放款/信用部上年度決算淨值)-上("+last_month+")月份之(建築放款/信用部上年度決算淨值)");
                }else if("WR004W_ATOT".equals(szrpt_code)){
                    cell.setCellValue("本季("+s_month+"月)之(建築放款/信用部上年度決算淨值)-上季("+lastSeason_month+"月)之(建築放款/信用部上年度決算淨值)");
                }else if("WR005W_ATOT".equals(szrpt_code)){
                    cell.setCellValue("本年("+s_year+"年"+s_month+"月)之(建築放款/信用部上年度決算淨值)-上年("+String.valueOf(Integer.parseInt(s_year) - 1)+"年"+s_month+"月)之(建築放款/信用部上年度決算淨值)");
                }
                cell = row.getCell( (short)13);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                if("WR003W_ATOT".equals(szrpt_code)){
                    cell.setCellValue("本("+s_month+")月份逾放比率-上("+last_month+")月份逾放比率");
                }else if("WR004W_ATOT".equals(szrpt_code)){
                    cell.setCellValue("本季("+s_month+"月)逾放比率-上季("+lastSeason_month+"月)逾放比率");
                }else if("WR005W_ATOT".equals(szrpt_code)){
                    cell.setCellValue("本年("+s_year+"年"+s_month+"月)逾放比率-上年("+String.valueOf(Integer.parseInt(s_year) - 1)+"年"+s_month+"月)逾放比率");
                }
                row = sheet.getRow(7);
                cell = row.getCell( (short)20);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                if("WR003W_ATOT".equals(szrpt_code)){
                    cell.setCellValue("本("+s_month+")月份放款覆蓋率-上("+last_month+")月份放款覆蓋率");
                }else if("WR004W_ATOT".equals(szrpt_code)){
                    cell.setCellValue("本季("+s_month+"月)放款覆蓋率-上季("+lastSeason_month+"月)放款覆蓋率");
                }else if("WR005W_ATOT".equals(szrpt_code)){
                    cell.setCellValue("本年("+s_year+"年"+s_month+"月)放款覆蓋率-上年("+String.valueOf(Integer.parseInt(s_year) - 1)+"年"+s_month+"月)放款覆蓋率");
                }
                
                row = sheet.getRow(9);
                cell = row.getCell( (short)10);
                cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                if("WR003W_ATOT".equals(szrpt_code)){
                    cell.setCellValue("與上("+last_month+")月份超逾100%者比較，本("+s_month+")月份增加超逾100%者");
                }else if("WR004W_ATOT".equals(szrpt_code)){
                    cell.setCellValue("與上季("+lastSeason_month+"月)超逾100%者比較，本季("+s_month+"月)增加超逾100%者");
                }else if("WR005W_ATOT".equals(szrpt_code)){
                    cell.setCellValue("與上年("+String.valueOf(Integer.parseInt(s_year) - 1)+"年"+s_month+"月)超逾100%者比較，本年("+s_year+"年"+s_month+"月)增加超逾100%者");
                }
            }
            if(dbData!=null && dbData.size()>0){
            for(int i=0;i<dbData.size();i++){
                bean = (DataObject)dbData.get(i);    
                getBeanData(isLoan,bean);
                row = sheet.createRow(10+i);
                int maxC = 20;
                if(isLoan) maxC = 6;
                for(int c=0;c<=maxC;c++){
                    row.createCell((short)c);
                    cell = row.getCell( (short)c);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(setInsertValue(isLoan,c,i+1));
                    if(c==0||c==1){
                        cell.setCellStyle(cs);
                    }else if(c==2){
                        cell.setCellStyle(cs1);
                    }else{
                        cell.setCellStyle(cs2);
                    }
                }
            }
            }
            String str = "";
            int sum = 0;
            int bank_count_sum = 0;
            if(dbData_sum.size()>0){
                for(int i=0;i<dbData_sum.size();i++){
                    bean_sum = (DataObject)dbData_sum.get(i);
                    wr_count =bean_sum.getValue("wr_count") == null?"0":(bean_sum.getValue("wr_count")).toString();
                    bank_count =bean_sum.getValue("bank_count") == null?"0":(bean_sum.getValue("bank_count")).toString();
                    sum += Integer.parseInt(wr_count)*Integer.parseInt(bank_count);
                    bank_count_sum += Integer.parseInt(bank_count);
                    if(i!=0) str +="，";
                    str += "("+wr_count+")＊"+bank_count+"家";
                }
            }
            row = sheet.createRow(10+dbData.size());
            cell = row.createCell( (short)2);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            String tatlePoint = "18項指標；";
            if(isLoan) tatlePoint = "4項指標；";
            cell.setCellValue(tatlePoint+bank_count_sum+"家列有警訊，計有"+sum+"項警訊；"+str);
            
 		    //建表結束--------------------------------------
            HSSFFooter footer = sheet.getFooter();
            footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
            footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
            FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+outputFile);
            System.out.println("儲存檔案:"+reportDir + System.getProperty("file.separator")+outputFile);
            wb.write(fout);//儲存
            fout.close();
            System.out.println("儲存完成");
            
      }catch(Exception e){
                System.out.println("createRpt Error:"+e+e.getMessage());
      }
      return errMsg;
    }//end of createRpt
    
    //取得各欄位data
    private static void getBeanData(boolean isLoan,DataObject bean){ 
    	try{
    	    bank_code =bean.getValue("bank_code") == null?"":String.valueOf(bean.getValue("bank_code"));
            bank_name =bean.getValue("bank_name") == null?"":String.valueOf(bean.getValue("bank_name"));
            cnt =bean.getValue("wr_count") == null?"":String.valueOf(bean.getValue("wr_count"));
    	    if(isLoan){
    	        diff_loan_bal_amt_max =bean.getValue("diff_loan_bal_amt_max") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("diff_loan_bal_amt_max")));//專案農貸放款餘額.增加金額前5名
                diff_over6m_loan_bal_amt_max =bean.getValue("diff_over6m_loan_bal_amt_max") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("diff_over6m_loan_bal_amt_max")));//專案農貸逾期放款.增加金額前5名
                diff_over6m_loan_rate_max =bean.getValue("diff_over6m_loan_rate_max") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("diff_over6m_loan_rate_max")));//專案農貸逾放比率.增加百分點前5名
                delay_loan_rate_max =bean.getValue("delay_loan_rate_max") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("delay_loan_rate_max")));//當年度累計核准延期還款件數占尚有餘額專案農貸總件數比率前5名
    	    }else{
                field_debit_rate_max =bean.getValue("field_debit_rate_max") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_debit_rate_max")));//存款總額.增加比率前5名
                field_debit_rate_min =bean.getValue("field_debit_rate_min") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_debit_rate_min")));//存款總額.減少比率前5名
                field_credit_rate_max =bean.getValue("field_credit_rate_max") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_credit_rate_max")));//放款總額.增加比率前5名
                field_credit_rate_min =bean.getValue("field_credit_rate_min") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_credit_rate_min")));//放款總額.減少比率前5名
                field_990610_rate_max =bean.getValue("field_990610_rate_max") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_990610_rate_max")));//非會員放款.增加比率前5名
                field_noassure_rate_max =bean.getValue("field_noassure_rate_max") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_noassure_rate_max")));//無擔保放款.增加比率前5名
                field_992710_rate_max =bean.getValue("field_992710_rate_max") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_992710_rate_max")));//建築放款.增加比率前5名 103.02.26調整為前5名
                fd_992710_990230_rate_100 =bean.getValue("fd_992710_990230_rate_100") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("fd_992710_990230_rate_100")));//建築放款/上年度信用部決算淨值>=100%
                diff_992710_990230_rate_max =bean.getValue("diff_992710_990230_rate_max") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("diff_992710_990230_rate_max")));//本月份之(建築放款/上年度信用部決算淨值)-上月份之(建築放款/上年度信用部決算淨值)增加百分點前5名103.02.26調整為前5名
                field_diff_over_max =bean.getValue("field_diff_over_max") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_diff_over_max")));//逾期放款增加金額前5名103.02.26調整為前5名
                field_diff_over_rate_max =bean.getValue("field_diff_over_rate_max") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_diff_over_rate_max")));//本月份逾放比率-上月份逾放比率增加百分點前5名103.02.26調整為前5名
                field_diff_992530_max =bean.getValue("field_diff_992530_max") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_diff_992530_max")));//逾放-非會員.增加金額前5名103.02.26調整為前5名
                fd_992550_992150_rate_max =bean.getValue("fd_992550_992150_rate_max") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("fd_992550_992150_rate_max")));//無擔保消費性放款中之逾放/無擔保消費性放款.前5名103.02.26調整為前5名
                fd_992720_992710_rate_max =bean.getValue("fd_992720_992710_rate_max") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("fd_992720_992710_rate_max")));//建築放款中之逾放/建築放款.前5名103.02.26調整為前5名
                field_diff_992610_cal_max =bean.getValue("field_diff_992610_cal_max") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_diff_992610_cal_max")));//應予觀察放款.增加金額前5名103.02.26調整為前5名
                field_diff_992630_max =bean.getValue("field_diff_992630_max") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_diff_992630_max")));//應予觀察放款-非會員.增加金額前5名
                field_diff_992730_max =bean.getValue("field_diff_992730_max") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_diff_992730_max")));//應予觀察放款-建築放款.增加金額前5名
                field_diff_backup_min =bean.getValue("field_diff_backup_min") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("field_diff_backup_min")));//備抵呆帳.減少金額前5名103.02.26調整為前5名
                fd_diff_backup_credit_rate_min =bean.getValue("fd_diff_backup_credit_rate_min") == null?"":Utility.setCommaFormat(String.valueOf(bean.getValue("fd_diff_backup_credit_rate_min")));//本月份之放款覆蓋率-上月份放款覆蓋率(備抵呆帳/放款)(28).最低前5名
    	    }
    	}catch(Exception e){
    		System.out.println("getBeanData Error:"+e+e.getMessage());
    	}
    }
   	private static String setInsertValue(boolean isLoan,int cellcount,int no){
   	    String insertValue = "";
   	    if(isLoan){
       	     if( cellcount==0 )insertValue = String.valueOf(no);
             else if( cellcount==1 )insertValue =bank_code;
             else if( cellcount==2 )insertValue =bank_name+"("+cnt+")";
             else if( cellcount==3 )insertValue =diff_loan_bal_amt_max;
             else if( cellcount==4 )insertValue =diff_over6m_loan_bal_amt_max;
             else if( cellcount==5 )insertValue =diff_over6m_loan_rate_max;
             else if( cellcount==6 )insertValue =delay_loan_rate_max;
   	    }else{
    		 if( cellcount==0 )insertValue = String.valueOf(no);
    		 else if( cellcount==1 )insertValue =bank_code;
             else if( cellcount==2 )insertValue =bank_name+"("+cnt+")";
             else if( cellcount==3 )insertValue =field_debit_rate_max;
             else if( cellcount==4 )insertValue =field_debit_rate_min;
             else if( cellcount==5 )insertValue =field_credit_rate_max;
             else if( cellcount==6 )insertValue =field_credit_rate_min;
             else if( cellcount==7 )insertValue =field_990610_rate_max;
             else if( cellcount==8 )insertValue =field_noassure_rate_max;
             else if( cellcount==9 )insertValue =field_992710_rate_max;
             else if( cellcount==10)insertValue =fd_992710_990230_rate_100;
             else if( cellcount==11)insertValue =diff_992710_990230_rate_max;
             else if( cellcount==12)insertValue =field_diff_over_max;
             else if( cellcount==13)insertValue =field_diff_over_rate_max;
             else if( cellcount==14)insertValue =field_diff_992530_max;
             else if( cellcount==15)insertValue =fd_992720_992710_rate_max;
             else if( cellcount==16)insertValue =field_diff_992610_cal_max;
             else if( cellcount==17)insertValue =field_diff_992630_max;
             else if( cellcount==18)insertValue =field_diff_992730_max;
             else if( cellcount==19)insertValue =field_diff_backup_min;
             else if( cellcount==20)insertValue =fd_diff_backup_credit_rate_min;
   	    }
   	    return insertValue;
    }
   	
}
