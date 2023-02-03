/*
 * 102.08.07 created by 2968 
 * 102.08.26 add 警示筆數 
 * 102.09.05 fix 公式A02.990320改為A02.990230
 */
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptWR003W {
    static String bank_code ="";
    static String bank_name ="";
    static String cnt ="";
    static String field_debit_rate_max ="";
    static String field_debit_rate_min ="";
    static String field_credit_rate_max ="";
    static String field_credit_rate_min ="";
    static String field_990610_rate_max ="";
    static String field_990610_rate_min ="";
    static String field_noassure_rate_max ="";
    static String field_noassure_rate_min ="";
    static String field_120700_rate_max ="";
    static String field_992710_max ="";
    static String field_992710_rate_max ="";
    static String fd_992710_credit_rate_max ="";
    static String fd_992710_990230_rate_100 ="";
    static String diff_992710_990230_rate_max ="";
    static String field_over_max ="";
    static String field_diff_over_max ="";
    static String field_over_rate_max ="";
    static String field_diff_over_rate_max ="";
    static String field_diff_992530_max ="";
    static String fd_992550_992150_rate_max ="";
    static String fd_992720_992710_rate_max ="";
    static String field_992610_cal_max ="";
    static String field_diff_992610_cal_max ="";
    static String fd_992610_cal_credit_rate_max ="";
    static String field_diff_992630_max ="";
    static String field_diff_992730_max ="";
    static String field_diff_backup_min ="";
    static String backup_over_rate_min ="";
    static String field_992710_990230_rate_avg ="";
    static String field_over_rate_avg ="";
	static File logfile;
    static File logDir = null;
    public static String createRpt(String s_year,String s_month,String unit,HSSFWorkbook wb){
          String errMsg = "";
          HSSFRow row=null;//宣告一列
          HSSFCell cell=null;//宣告一個儲存格
          HSSFSheet sheet =null;
          FileInputStream finput = null;
          StringBuffer sqlCmd = new StringBuffer();
          List paramList = new ArrayList();
          List dbData = null;
          DataObject bean = null;
          StringBuffer sqlCmd_avg = new StringBuffer();
          List paramList_avg = new ArrayList();
          List dbData_avg = null;
          DataObject bean_avg = null;
          String cd01_table = "";
          String wlx01_m_year = "";
          String unit_name = "";
          unit_name = Utility.getUnitName(unit);//取得單位名稱
          Calendar now = Calendar.getInstance();
          String nowYear  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
          String nowMonth= String.valueOf(now.get(Calendar.MONTH)+1);   //月份以0開始故加1取得實際月份;
          String nowDay = String.valueOf(now.get(Calendar.DAY_OF_MONTH));//日期
          String last_year  = (Integer.parseInt(s_month) == 1) ? String.valueOf(Integer.parseInt(s_year) - 1) : s_year; 
          String last_month = (Integer.parseInt(s_month) == 1) ? "12" : String.valueOf(Integer.parseInt(s_month) - 1);
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
  	    	cd01_table = (Integer.parseInt(s_year) < 100)?"cd01_99":""; 
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
            String openfile="農漁會信用部營運狀況警訊報表-與上月比較.xls";
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
	        ps.setScale( ( short )64 ); //列印縮放百分比
	        ps.setPaperSize( ( short )8 ); //設定紙張大小 A3
	  		//wb.setSheetName(0,"test");
	  		//設定表頭 為固定 先設欄的起始再設列的起始
	        //wb.setRepeatingRowsAndColumns(0, 1, 21, 2, 3);
            System.out.println("************** unit="+unit);
	  		sqlCmd.append(" select bank_code,");//--機構代號
	  		sqlCmd.append("        bn01.bank_name,");//--機構名稱
	  		sqlCmd.append("        count(*) cnt,");//警示筆數 102.08.26 add
	  		sqlCmd.append("        sum(decode(acc_code,'field_debit_rate',decode(wr_range_serial,1,amt,''),'')) as field_debit_rate_max,");//--存款總額.增加比率前5名
	  		sqlCmd.append("        sum(decode(acc_code,'field_debit_rate',decode(wr_range_serial,2,amt,''),'')) as field_debit_rate_min,");//--存款總額.減少比率前5名
	  		sqlCmd.append("        sum(decode(acc_code,'field_credit_rate',decode(wr_range_serial,3,amt,''),'')) as field_credit_rate_max,");//--放款總額.增加比率前5名
	  		sqlCmd.append("        sum(decode(acc_code,'field_credit_rate',decode(wr_range_serial,4,amt,''),'')) as field_credit_rate_min,");//--放款總額.減少比率前5名
	  		sqlCmd.append("        sum(decode(acc_code,'field_990610_rate',decode(wr_range_serial,5,amt,''),'')) as field_990610_rate_max,");//--非會員放款.增加比率前5名
	  		sqlCmd.append("        sum(decode(acc_code,'field_990610_rate',decode(wr_range_serial,6,amt,''),'')) as field_990610_rate_min,");//--非會員放款.減少比率前5名
	  		sqlCmd.append("        sum(decode(acc_code,'field_noassure_rate',decode(wr_range_serial,7,amt,''),'')) as field_noassure_rate_max,");//--無擔保放款.增加比率前5名
	  		sqlCmd.append("        sum(decode(acc_code,'field_noassure_rate',decode(wr_range_serial,8,amt,''),'')) as field_noassure_rate_min,");//--無擔保放款.減少比率前5名
	  		sqlCmd.append("        sum(decode(acc_code,'field_120700_rate',decode(wr_range_serial,9,amt,''),'')) as field_120700_rate_max,");//--內部融資.增加比率前5名
	  		sqlCmd.append("        sum(decode(acc_code,'field_992710',decode(wr_range_serial,10, round(amt /?,0) ,''),'')) as field_992710_max,");//--建築放款.前10名
	  		sqlCmd.append("        sum(decode(acc_code,'field_992710_rate',decode(wr_range_serial,11,amt,''),'')) as field_992710_rate_max,");//--建築放款.增加比率前10名
	  		sqlCmd.append("        sum(decode(acc_code,'field_992710_credit_rate',decode(wr_range_serial,12,amt,''),'')) as fd_992710_credit_rate_max,");//--建築放款/放款.前10名
	  		sqlCmd.append("        sum(decode(acc_code,'field_992710_990230_rate',decode(wr_range_serial,13,amt,''),'')) as fd_992710_990230_rate_100,");//--建築放款/上年度信用部決算淨值>=100%
	  		sqlCmd.append("        sum(decode(acc_code,'field_diff_992710_990230_rate',decode(wr_range_serial,14,amt,''),'')) as diff_992710_990230_rate_max,");//--本月份之(建築放款/上年度信用部決算淨值)-上月份之(建築放款/上年度信用部決算淨值)增加百分點前10名
	  		sqlCmd.append("        sum(decode(acc_code,'field_over',decode(wr_range_serial,15,round(amt /?,0),''),'')) as field_over_max,");//--逾期放款本月前10名
	  		sqlCmd.append("        sum(decode(acc_code,'field_diff_over',decode(wr_range_serial,16,round(amt /?,0),''),'')) as field_diff_over_max,");//--逾期放款增加金額前10名
	  		sqlCmd.append("        sum(decode(acc_code,'field_over_rate',decode(wr_range_serial,17,amt,''),'')) as field_over_rate_max,");//--'逾放比率本月最高前10名
	  		sqlCmd.append("        sum(decode(acc_code,'field_diff_over_rate',decode(wr_range_serial,18,amt,''),'')) as field_diff_over_rate_max,");//--本月份逾放比率-上月份逾放比率增加百分點前10名
	  		sqlCmd.append("        sum(decode(acc_code,'field_diff_992530',decode(wr_range_serial,19,round(amt /?,0),''),'')) as field_diff_992530_max,");//--逾放-非會員.增加金額前10名
	  		sqlCmd.append("        sum(decode(acc_code,'field_992550_992150_rate',decode(wr_range_serial,20,amt,''),'')) as fd_992550_992150_rate_max,");//--無擔保消費性放款中之逾放/無擔保消費性放款.前5名
	  		sqlCmd.append("        sum(decode(acc_code,'field_992720_992710_rate',decode(wr_range_serial,21,amt,''),'')) as fd_992720_992710_rate_max,");//--建築放款中之逾放/建築放款.前10名
	  		sqlCmd.append("        sum(decode(acc_code,'field_992610_cal',decode(wr_range_serial,22,round(amt /?,0),''),'')) as field_992610_cal_max,");//--應予觀察放款.前10名
	  		sqlCmd.append("        sum(decode(acc_code,'field_diff_992610_cal',decode(wr_range_serial,23,round(amt /?,0),''),'')) as field_diff_992610_cal_max,");//--應予觀察放款.增加金額前10名
	  		sqlCmd.append("        sum(decode(acc_code,'field_992610_cal_credit_rate',decode(wr_range_serial,24,amt,''),'')) as fd_992610_cal_credit_rate_max,");//--應予觀察放款/放款.前10名
	  		sqlCmd.append("        sum(decode(acc_code,'field_diff_992630',decode(wr_range_serial,25,round(amt /?,0),''),'')) as field_diff_992630_max,");//--應予觀察放款-非會員.增加金額前5名
	  		sqlCmd.append("        sum(decode(acc_code,'field_diff_992730',decode(wr_range_serial,26,round(amt /?,0),''),'')) as field_diff_992730_max,");//--應予觀察放款-建築放款.增加金額前5名
	  		sqlCmd.append("        sum(decode(acc_code,'field_diff_backup',decode(wr_range_serial,27,round(amt /?,0),''),'')) as field_diff_backup_min,");//--備抵呆帳.減少金額前10名
	  		sqlCmd.append("        sum(decode(acc_code,'field_backup_over_rate',decode(wr_range_serial,28,amt,''),'')) as backup_over_rate_min ");//--備抵呆帳/逾期放款.最低前10名
	  		sqlCmd.append("   from wr_operation left join (select * from bn01 where m_year=?)bn01 on wr_operation.bank_code=bn01.bank_no ");
	  		sqlCmd.append("  where wr_operation.m_year=? ");
	  		sqlCmd.append("    and m_month=? ");
	  		sqlCmd.append("    and warn_type='Y' and wr_rpt='0' ");
	  		sqlCmd.append("  group by bank_code,bank_name ");
	  		sqlCmd.append("  order by bank_code ");
	  		paramList.add(unit);
	  		paramList.add(unit);
	  		paramList.add(unit);
	  		paramList.add(unit);
            paramList.add(unit);
            paramList.add(unit);
            paramList.add(unit);
            paramList.add(unit);
            paramList.add(unit);
            paramList.add(wlx01_m_year);
	  		paramList.add(s_year);
	  		paramList.add(s_month);
	  		dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,
	  		             "bank_code,bank_name,cnt,field_debit_rate_max,field_debit_rate_min,field_credit_rate_max,field_credit_rate_min,"
	  		            +"field_990610_rate_max,field_990610_rate_min,field_noassure_rate_max,field_noassure_rate_min,"
	  		            +"field_120700_rate_max,field_992710_max,field_992710_rate_max,fd_992710_credit_rate_max,"
	  		            +"fd_992710_990230_rate_100,diff_992710_990230_rate_max,field_over_max,field_diff_over_max,"
	  		            +"field_over_rate_max,field_diff_over_rate_max,field_diff_992530_max,fd_992550_992150_rate_max,"
	  		            +"fd_992720_992710_rate_max,field_992610_cal_max,field_diff_992610_cal_max,fd_992610_cal_credit_rate_max,"
	  		            +"field_diff_992630_max,field_diff_992730_max,field_diff_backup_min,backup_over_rate_min");
	  		System.out.print("總表資料 共"+dbData.size()+"筆");
	  		
	  		
	  		sqlCmd_avg.append(" select sum(decode(acc_code,'field_992710_990230_rate_avg',amt,'')) as  field_992710_990230_rate_avg,");//--平均數.建築放款/上年度信用部決算淨值
            sqlCmd_avg.append("        sum(decode(acc_code,'field_over_rate_avg',amt,'')) as  field_over_rate_avg ");//--逾放比率.平均數
            sqlCmd_avg.append(" from wr_operation ");
            sqlCmd_avg.append("  where m_year=? ");
            sqlCmd_avg.append("    and m_month=? ");
            sqlCmd_avg.append("    and wr_rpt='0' ");
            paramList_avg.add(s_year);
            paramList_avg.add(s_month);
            dbData_avg = DBManager.QueryDB_SQLParam(sqlCmd_avg.toString(),paramList_avg,"field_992710_990230_rate_avg,field_over_rate_avg");
            System.out.print("平均數.size()="+dbData_avg.size());
            
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
            cell.setCellValue(s_year+"年"+s_month+"月份農漁會信用部營運狀況警訊報表");
                     
            row = sheet.getRow(1);
            cell = row.getCell( (short)0);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellValue("與上月份("+last_year+"年"+last_month+"月)比較");
            
            row = sheet.getRow(2);
            cell = row.getCell( (short)0);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellValue("單位:新臺幣 " + unit_name + ",%");
            
            if(dbData_avg.size()>0){
                bean_avg = (DataObject)dbData_avg.get(0);
                field_992710_990230_rate_avg =bean_avg.getValue("field_992710_990230_rate_avg") == null?"":Utility.setCommaFormat((bean_avg.getValue("field_992710_990230_rate_avg")).toString());//平均數.建築放款/上年度信用部決算淨值
                field_over_rate_avg =bean_avg.getValue("field_over_rate_avg") == null?"":Utility.setCommaFormat((bean_avg.getValue("field_over_rate_avg")).toString());//逾放比率.平均數
            }
            row = sheet.getRow(8);
            cell = row.getCell( (short)15);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellValue("同業平均數/"+field_992710_990230_rate_avg);
            cell = row.getCell( (short)19);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellValue("同業平均數/"+field_over_rate_avg);
            
            for(int i=0;i<dbData.size();i++){
                bean = (DataObject)dbData.get(i);    
                getBeanData(bean);
                row = sheet.createRow(11+i);
                for(int c=0;c<=30;c++){
                    row.createCell((short)c);
                    cell = row.getCell( (short)c);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(setInsertValue(c,i+1));
                    if(c==0||c==1){
                        cell.setCellStyle(cs);
                    }else if(c==2){
                        cell.setCellStyle(cs1);
                    }else{
                        cell.setCellStyle(cs2);
                    }
                }
                
            }
            
 		    //建表結束--------------------------------------
            HSSFFooter footer = sheet.getFooter();
            footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
            footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
            FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+openfile);
            wb.write(fout);//儲存
            fout.close();
            System.out.println("儲存完成");
            
      }catch(Exception e){
                System.out.println("createRpt Error:"+e+e.getMessage());
      }
      return errMsg;
    }//end of createRpt
    
    //取得各欄位data
    private static void getBeanData(DataObject bean){
    	try{
    	    bank_code =bean.getValue("bank_code") == null?"":String.valueOf(bean.getValue("bank_code"));
    	    bank_name =bean.getValue("bank_name") == null?"":String.valueOf(bean.getValue("bank_name"));
    	    cnt =bean.getValue("cnt") == null?"":String.valueOf(bean.getValue("cnt"));
    	    field_debit_rate_max =bean.getValue("field_debit_rate_max") == null?"":Utility.setCommaFormat((bean.getValue("field_debit_rate_max")).toString());//存款總額.增加比率前5名
    	    field_debit_rate_min =bean.getValue("field_debit_rate_min") == null?"":Utility.setCommaFormat((bean.getValue("field_debit_rate_min")).toString());//存款總額.減少比率前5名
    	    field_credit_rate_max =bean.getValue("field_credit_rate_max") == null?"":Utility.setCommaFormat((bean.getValue("field_credit_rate_max")).toString());//放款總額.增加比率前5名
    	    field_credit_rate_min =bean.getValue("field_credit_rate_min") == null?"":Utility.setCommaFormat((bean.getValue("field_credit_rate_min")).toString());//放款總額.減少比率前5名
    	    field_990610_rate_max =bean.getValue("field_990610_rate_max") == null?"":Utility.setCommaFormat((bean.getValue("field_990610_rate_max")).toString());//非會員放款.增加比率前5名
    	    field_990610_rate_min =bean.getValue("field_990610_rate_min") == null?"":Utility.setCommaFormat((bean.getValue("field_990610_rate_min")).toString());//非會員放款.減少比率前5名
    	    field_noassure_rate_max =bean.getValue("field_noassure_rate_max") == null?"":Utility.setCommaFormat((bean.getValue("field_noassure_rate_max")).toString());//無擔保放款.增加比率前5名
    	    field_noassure_rate_min =bean.getValue("field_noassure_rate_min") == null?"":Utility.setCommaFormat((bean.getValue("field_noassure_rate_min")).toString());//無擔保放款.減少比率前5名
    	    field_120700_rate_max =bean.getValue("field_120700_rate_max") == null?"":Utility.setCommaFormat((bean.getValue("field_120700_rate_max")).toString());//內部融資.增加比率前5名
    	    field_992710_max =bean.getValue("field_992710_max") == null?"":Utility.setCommaFormat((bean.getValue("field_992710_max")).toString());//建築放款.前10名
    	    field_992710_rate_max =bean.getValue("field_992710_rate_max") == null?"":Utility.setCommaFormat((bean.getValue("field_992710_rate_max")).toString());//建築放款.增加比率前10名
    	    fd_992710_credit_rate_max =bean.getValue("fd_992710_credit_rate_max") == null?"":Utility.setCommaFormat((bean.getValue("fd_992710_credit_rate_max")).toString());//建築放款/放款.前10名
    	    fd_992710_990230_rate_100 =bean.getValue("fd_992710_990230_rate_100") == null?"":Utility.setCommaFormat((bean.getValue("fd_992710_990230_rate_100")).toString());//建築放款/上年度信用部決算淨值>=100%
    	    diff_992710_990230_rate_max =bean.getValue("diff_992710_990230_rate_max") == null?"":Utility.setCommaFormat((bean.getValue("diff_992710_990230_rate_max")).toString());//本月份之(建築放款/上年度信用部決算淨值)-上月份之(建築放款/上年度信用部決算淨值)增加百分點前10名
    	    field_over_max =bean.getValue("field_over_max") == null?"":Utility.setCommaFormat((bean.getValue("field_over_max")).toString());//逾期放款本月前10名
    	    field_diff_over_max =bean.getValue("field_diff_over_max") == null?"":Utility.setCommaFormat((bean.getValue("field_diff_over_max")).toString());//逾期放款增加金額前10名
    	    field_over_rate_max =bean.getValue("field_over_rate_max") == null?"":Utility.setCommaFormat((bean.getValue("field_over_rate_max")).toString());//逾放比率本月最高前10名
    	    field_diff_over_rate_max =bean.getValue("field_diff_over_rate_max") == null?"":Utility.setCommaFormat((bean.getValue("field_diff_over_rate_max")).toString());//本月份逾放比率-上月份逾放比率增加百分點前10名
    	    field_diff_992530_max =bean.getValue("field_diff_992530_max") == null?"":Utility.setCommaFormat((bean.getValue("field_diff_992530_max")).toString());//逾放-非會員.增加金額前10名
    	    fd_992550_992150_rate_max =bean.getValue("fd_992550_992150_rate_max") == null?"":Utility.setCommaFormat((bean.getValue("fd_992550_992150_rate_max")).toString());//無擔保消費性放款中之逾放/無擔保消費性放款.前5名
    	    fd_992720_992710_rate_max =bean.getValue("fd_992720_992710_rate_max") == null?"":Utility.setCommaFormat((bean.getValue("fd_992720_992710_rate_max")).toString());//建築放款中之逾放/建築放款.前10名
    	    field_992610_cal_max =bean.getValue("field_992610_cal_max") == null?"":Utility.setCommaFormat((bean.getValue("field_992610_cal_max")).toString());//應予觀察放款.前10名
    	    field_diff_992610_cal_max =bean.getValue("field_diff_992610_cal_max") == null?"":Utility.setCommaFormat((bean.getValue("field_diff_992610_cal_max")).toString());//應予觀察放款.增加金額前10名
    	    fd_992610_cal_credit_rate_max =bean.getValue("fd_992610_cal_credit_rate_max") == null?"":Utility.setCommaFormat((bean.getValue("fd_992610_cal_credit_rate_max")).toString());//應予觀察放款/放款.前10名
    	    field_diff_992630_max =bean.getValue("field_diff_992630_max") == null?"":Utility.setCommaFormat((bean.getValue("field_diff_992630_max")).toString());//應予觀察放款-非會員.增加金額前5名
    	    field_diff_992730_max =bean.getValue("field_diff_992730_max") == null?"":Utility.setCommaFormat((bean.getValue("field_diff_992730_max")).toString());//應予觀察放款-建築放款.增加金額前5名
    	    field_diff_backup_min =bean.getValue("field_diff_backup_min") == null?"":Utility.setCommaFormat((bean.getValue("field_diff_backup_min")).toString());//備抵呆帳.減少金額前10名
    	    backup_over_rate_min =bean.getValue("backup_over_rate_min") == null?"":Utility.setCommaFormat((bean.getValue("backup_over_rate_min")).toString());//備抵呆帳/逾期放款.最低前10名
    	}catch(Exception e){
    		System.out.println("getBeanData Error:"+e+e.getMessage());
    	}
    }
   	private static String setInsertValue(int cellcount,int no){
   	         String insertValue = "";
    		 if( cellcount==0 )insertValue = String.valueOf(no);
    		 else if( cellcount==1 )insertValue =bank_code;
             else if( cellcount==2 )insertValue =bank_name+"("+cnt+")";
             else if( cellcount==3 )insertValue =field_debit_rate_max;
             else if( cellcount==4 )insertValue =field_debit_rate_min;
             else if( cellcount==5 )insertValue =field_credit_rate_max;
             else if( cellcount==6 )insertValue =field_credit_rate_min;
             else if( cellcount==7 )insertValue =field_990610_rate_max;
             else if( cellcount==8 )insertValue =field_990610_rate_min;
             else if( cellcount==9 )insertValue =field_noassure_rate_max;
             else if( cellcount==10 )insertValue =field_noassure_rate_min;
             else if( cellcount==11 )insertValue =field_120700_rate_max;
             else if( cellcount==12 )insertValue =field_992710_max;
             else if( cellcount==13 )insertValue =field_992710_rate_max;
             else if( cellcount==14 )insertValue =fd_992710_credit_rate_max;
             else if( cellcount==15 )insertValue =fd_992710_990230_rate_100;
             else if( cellcount==16 )insertValue =diff_992710_990230_rate_max;
             else if( cellcount==17 )insertValue =field_over_max;
             else if( cellcount==18 )insertValue =field_diff_over_max;
             else if( cellcount==19 )insertValue =field_over_rate_max;
             else if( cellcount==20 )insertValue =field_diff_over_rate_max;
             else if( cellcount==21 )insertValue =field_diff_992530_max;
             else if( cellcount==22 )insertValue =fd_992550_992150_rate_max;
             else if( cellcount==23 )insertValue =fd_992720_992710_rate_max;
             else if( cellcount==24 )insertValue =field_992610_cal_max;
             else if( cellcount==25 )insertValue =field_diff_992610_cal_max;
             else if( cellcount==26 )insertValue =fd_992610_cal_credit_rate_max;
             else if( cellcount==27 )insertValue =field_diff_992630_max;
             else if( cellcount==28 )insertValue =field_diff_992730_max;
             else if( cellcount==29 )insertValue =field_diff_backup_min;
             else if( cellcount==30 )insertValue =backup_over_rate_min;
             return insertValue;
    }
   	
}
