/*
 * Created 農漁會信用部財業務資料_動態 by 2968
 * 103.04.09 add by 2968
 *           1.原無此申報資料以空白顯示 
 *           2.(三)主要業務增加合計欄位為各上開儲存格加總,原合計調整為總額
 * 111.02.24 調整直接讀取excel格式 by 2295
 *           調整上月份資料  UI月份減1 (若為1月份則為前一年度12月份) by 2295           
 *
 */
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.awt.Font;
import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

import com.sun.corba.se.impl.javax.rmi.CORBA.Util;
import com.tradevan.util.dao.DataObject;
import com.tradevan.util.*; 

public class RptWR002W {
    
	static DecimalFormat df_md = new DecimalFormat("############0.00");//顯示小數點至第2位,不足者補0
	static File logfile;
    static FileOutputStream logos=null;      
    static BufferedOutputStream logbos = null;
    static PrintStream logps = null;
    static Date nowlog = new Date();
    static SimpleDateFormat logformat = new SimpleDateFormat("yyyy/MM/dd  HH:mm:ss  ");        
    static SimpleDateFormat logfileformat = new SimpleDateFormat("yyyyMMddHHmmss");
    static Calendar logcalendar;
    static File logDir = null;
    static String field_190000_cal = "";//--資產總額
    static String field_debit = "";//--存款
    static String field_credit = "";//--放款
    static String field_net = "";//--淨值
    static String field_over_rate = "";//--逾放比率
    static String field_over = "";//--逾期放款
    static String field_992510 = "";//--逾期放款-正會員
    static String field_992520 = "";//--逾期放款-贊助會員
    static String field_992530 = "";//--逾期放款-非會員
    static String field_992720 = "";//--建築放款中之逾期放款
    static String field_992550 = "";//--無擔保放款之逾期放款
    static String field_backup_over_rate = "";//--備抵呆帳覆蓋率=備抵呆帳/逾期放款
    static String field_captial_rate = "";//  --資本適足率
    static String field_420000_cal = "";//--收入
    static String field_520000_cal = "";//--出支
    static String field_320300 = "";//--本期損益
    static String field_992130 = "";//--存款業務.正會員
    static String field_990420 = "";//--存款業務.贊助會員
    static String field_990310 = "";//--存款業務.非會員
    static String field_debit_1 = "";//--存款業務.合計
    static String field_992140 = "";//--放款業務.放款-正會員
    static String field_990410 = "";//--放款業務.放款-贊助會員
    static String field_990610_990611 = "";//--放款業務.放款-非會員
    static String field_120700 = "";//--放款業務.放款-內部融資
    static String field_990611 = "";//--放款業務.放款-縣市政府貸款
    static String field_credit_1 = "";//--放款業務.放款-合計
    static String field_990510 = "";//--無擔保放款-非會員
    static String field_noassure = "";//--無擔保放款-合計
    static String field_992710 = "";//--建築放款合計
    static String field_sum3_1 = "";//存款業務.合計
    static String field_sum_3_2_1 = "";//放款業務.放款-合計
    static String field_sum_3_2_2 = "";//無擔保放款-合計
    public static String createRpt(String s_year,String s_month,String unit,String bank_no){
          String errMsg = "";
          String unit_name = Utility.getUnitName(unit);
          String wlx01_m_year = "";
          String m_year = "";
          String m_month = "";
          int cellNo = 0;
          reportUtil reportUtil = new reportUtil();
          FileOutputStream fileOut = null;          
          HSSFRow row=null;//宣告一列
          HSSFCell cell=null;//宣告一個儲存格
      try{
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
            String openfile="農漁會信用部財業務資料_動態.xls";//要去開啟的範本檔
            
            System.out.println("開啟檔:" + openfile);
            FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+ openfile );         
            //設定FileINputStream讀取Excel檔
            
            POIFSFileSystem fs = new POIFSFileSystem( finput );//新增一個xls unit
            if(fs==null){
                System.out.println("open 範本檔失敗");
            } else{ 
                System.out.println("open 範本檔成功");
            }
            
            HSSFWorkbook wb = new HSSFWorkbook(fs);//新增一個sheet
            if(wb==null){
                System.out.println("open工作表失敗");
            } else {
                System.out.println("open 工作表成功");
            }
            
            //對第一個sheet工作
            HSSFSheet sheet = wb.getSheetAt(0);//讀取第一個工作表，宣告其為sheet 
            if(sheet==null){
                System.out.println("open sheet 失敗");
            }else {
                System.out.println("open sheet 成功");
            }
            
            //做屬性設定
            HSSFPrintSetup ps = sheet.getPrintSetup();  //取得設定
            //sheet.setZoom(80, 100);                   //螢幕上看到的縮放大小
            //sheet.setAutobreaks(true);                //自動分頁
            
            //設定頁面符合列印大小
            sheet.setAutobreaks( false );
            ps.setScale( ( short )70 );                 //列印縮放百分比
            ps.setPaperSize( ( short )9 );              //設定紙張大小 A4
            
            //設定表頭 為固定 先設欄的起始再設列的起始
            //wb.setRepeatingRowsAndColumns(0, 1, 17, 2, 3);
            finput.close();
            HSSFFooter footer = sheet.getFooter();            
	  		if(sheet==null){System.out.println("open sheet 失敗");}else System.out.println("open sheet 成功");
	  		footer.setCenter( "Page:" + HSSFFooter.page() + " of " +
                    HSSFFooter.numPages() );                                        
	  		footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa")); 
	  		wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100";
	  		//設定表頭===============================================================================
	  		HSSFCellStyle style = wb.createCellStyle();
	        HSSFFont font = wb.createFont();
	        font.setFontHeightInPoints((short) 14);
	        style.setFont(font);
	        row=(sheet.getRow((short)0)==null)? sheet.createRow((short)0) : sheet.getRow((short)0);
            cell = (row.getCell((short)1)==null)? row.getCell((short)1) : row.getCell((short)1);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            String bank_name = getBank_name(bank_no,wlx01_m_year);
            cell.setCellValue(s_year+"年度"+s_month+"月份"+bank_name+"財業務資料");
            cell.setCellStyle(style);
            Calendar now = Calendar.getInstance();
            String nowYear  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
            String nowMonth = String.valueOf(now.get(Calendar.MONTH)+1);   //月份以0開始故加1取得實際月份;
            String nowDay   = String.valueOf(now.get(Calendar.DATE));
            insertCell("列印日期："+nowYear+"年"+nowMonth+"月"+nowDay+"日",wb,row,(short)12,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_LEFT);
            row=(sheet.getRow((short)1)==null)? sheet.createRow((short)1) : sheet.getRow((short)1);
            insertCell("單位：新台幣"+Utility.getUnitName(unit)+"、％",wb,row,(short)12,HSSFCellStyle.BORDER_NONE,HSSFCellStyle.ALIGN_RIGHT);
            
            //寫入資料===============================================================================
            
            //當月份資料   UI年UI月
            cellNo = 2;
            m_year = s_year;
            m_month = s_month;
            wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100"; 
            List mainInfo = mainInfo(s_year,s_month,unit,bank_no,wlx01_m_year);
            setDate(sheet,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,m_year,m_month);
            DataObject mainInfoBean = new DataObject();
            if(mainInfo.size()>0){
                mainInfoBean = (DataObject)mainInfo.get(0);
            }
            setValues(sheet,wb,row,mainInfoBean,cellNo);
            
            //上月份資料  UI月份減1 (若為1月份則為前一年度12月份)
            cellNo = 4;
            System.out.println("s_year="+s_year);
            System.out.println("s_month="+s_month);
            m_year = (Integer.parseInt(s_month)-1==0) ? String.valueOf(Integer.parseInt(s_year)-1) : s_year;
            m_month = (Integer.parseInt(s_month)-1==0) ? "12" : String.valueOf(Integer.parseInt(s_month)-1);
            System.out.println("m_year="+m_year);
            System.out.println("m_month="+m_month);
            if(m_month.length()==1) m_month = "0"+m_month;
            wlx01_m_year = (Integer.parseInt(m_year) < 100)?"99":"100"; 
            List lastMInfo = mainInfo(m_year,m_month,unit,bank_no,wlx01_m_year);
            setDate(sheet,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,m_year,m_month);
            DataObject lastMInfoBean = new DataObject();
            if(lastMInfo.size()>0){
                lastMInfoBean = (DataObject)lastMInfo.get(0);
            }
            setValues(sheet,wb,row,lastMInfoBean,cellNo);
            
            //UI年-1 12月份
            cellNo = 6; 
            m_year = String.valueOf(Integer.parseInt(s_year)-1);
            m_month = "12";
            wlx01_m_year = (Integer.parseInt(m_year) < 100)?"99":"100"; 
            List lastY12Info = mainInfo(m_year,m_month,unit,bank_no,wlx01_m_year);
            setDate(sheet,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,m_year,m_month);
            DataObject lastY12InfoBean = new DataObject();
            if(lastY12Info.size()>0){
                lastY12InfoBean = (DataObject)lastY12Info.get(0);
            }
            setValues(sheet,wb,row,lastY12InfoBean,cellNo);
            
            //UI年-1 9月份
            cellNo = 7;
            m_year = String.valueOf(Integer.parseInt(s_year)-1);
            m_month = "09";
            wlx01_m_year = (Integer.parseInt(m_year) < 100)?"99":"100"; 
            List lastY09Info = mainInfo(m_year,m_month,unit,bank_no,wlx01_m_year);
            setDate(sheet,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,m_year,m_month);
            DataObject lastY09InfoBean = new DataObject();
            if(lastY09Info.size()>0){
                lastY09InfoBean = (DataObject)lastY09Info.get(0);
            }
            setValues(sheet,wb,row,lastY09InfoBean,cellNo);
            
            //UI年-1 6月份
            cellNo = 8;
            m_year = String.valueOf(Integer.parseInt(s_year)-1);
            m_month = "06";
            wlx01_m_year = (Integer.parseInt(m_year) < 100)?"99":"100"; 
            List lastY06Info = mainInfo(m_year,m_month,unit,bank_no,wlx01_m_year);
            setDate(sheet,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,m_year,m_month);
            DataObject lastY06InfoBean = new DataObject();
            if(lastY06Info.size()>0){
                lastY06InfoBean = (DataObject)lastY06Info.get(0);
            }
            setValues(sheet,wb,row,lastY06InfoBean,cellNo);
            
            //上年度同期資料  UI年-1 UI月份
            cellNo = 9;
            m_year = String.valueOf(Integer.parseInt(s_year)-1);
            m_month = s_month;
            wlx01_m_year = (Integer.parseInt(m_year) < 100)?"99":"100"; 
            List lastYInfo = mainInfo(m_year,m_month,unit,bank_no,wlx01_m_year);
            setDate(sheet,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,m_year,m_month);
            DataObject lastYInfoBean = new DataObject();
            if(lastYInfo.size()>0){
                lastYInfoBean = (DataObject)lastYInfo.get(0);
            }
            setValues(sheet,wb,row,lastYInfoBean,cellNo);
            
            //UI年-1 3月份
            cellNo = 11; 
            m_year = String.valueOf(Integer.parseInt(s_year)-1);
            m_month = "03";
            wlx01_m_year = (Integer.parseInt(m_year) < 100)?"99":"100"; 
            List lastY03Info = mainInfo(m_year,m_month,unit,bank_no,wlx01_m_year);
            setDate(sheet,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,m_year,m_month);
            DataObject lastY03InfoBean = new DataObject();
            if(lastY03Info.size()>0){
                lastY03InfoBean = (DataObject)lastY03Info.get(0);
            }
            setValues(sheet,wb,row,lastY03InfoBean,cellNo);
            
            //UI年-2 12月份
            cellNo = 12;
            m_year = String.valueOf(Integer.parseInt(s_year)-2);
            m_month = "12";
            wlx01_m_year = (Integer.parseInt(m_year) < 100)?"99":"100"; 
            List last2Y12Info = mainInfo(m_year,m_month,unit,bank_no,wlx01_m_year);
            setDate(sheet,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,m_year,m_month);
            DataObject last2Y12InfoBean = new DataObject();
            if(last2Y12Info.size()>0){
                last2Y12InfoBean = (DataObject)last2Y12Info.get(0);
            }
            setValues(sheet,wb,row,last2Y12InfoBean,cellNo);
            
            //UI年-3 12月份
            cellNo = 13;
            m_year = String.valueOf(Integer.parseInt(s_year)-3);
            m_month = "12";
            wlx01_m_year = (Integer.parseInt(m_year) < 100)?"99":"100"; 
            List last3Y12Info = mainInfo(m_year,m_month,unit,bank_no,wlx01_m_year);
            setDate(sheet,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_NONE,m_year,m_month);
            DataObject last3Y12InfoBean = new DataObject();
            if(last3Y12Info.size()>0){
                last3Y12InfoBean = (DataObject)last3Y12Info.get(0);
            }
            setValues(sheet,wb,row,last3Y12InfoBean,cellNo);
            
            //UI年-4 12月份
            cellNo = 14;
            m_year = String.valueOf(Integer.parseInt(s_year)-4);
            m_month = "12";
            wlx01_m_year = (Integer.parseInt(m_year) < 100)?"99":"100"; 
            List last4Y12Info = mainInfo(m_year,m_month,unit,bank_no,wlx01_m_year);
            setDate(sheet,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.BORDER_THIN,m_year,m_month);
            DataObject last4Y12InfoBean = new DataObject();
            if(last4Y12Info.size()>0){
                last4Y12InfoBean = (DataObject)last4Y12Info.get(0);
            }
            setValues(sheet,wb,row,last4Y12InfoBean,cellNo);
           
            
            // Write the output to a file============================   
           fileOut = new FileOutputStream( Utility.getProperties("reportDir")+System.getProperty("file.separator")+"農漁會信用部財業務資料_動態.xls" );
           wb.write( fileOut );
           fileOut.close();            
           System.out.println("儲存完成");
        } catch ( Exception e ) {            
           e.printStackTrace();
           
        } finally {
           try {
               if ( fileOut != null ) {
                   fileOut.close();
               }
           } catch ( Exception e ) {
                 System.out.println(e.getMessage() );
           }
        }
      return errMsg;
    }
    
   	public static void printLog(PrintStream logps,String errRptMsg){
        if(!errRptMsg.equals("")){
           logcalendar = Calendar.getInstance(); 
           nowlog = logcalendar.getTime();
           logps.println(logformat.format(nowlog)+errRptMsg);
           logps.flush();
        }
   }
   	
   public static List mainInfo(String s_year,String s_month,String unit,String bank_no,String wlx01_m_year){
       StringBuffer sqlCmd = new StringBuffer(); 
       List paramList = new ArrayList();
       sqlCmd.append(" select a01.bank_no,a01.bank_name,");
       sqlCmd.append(" round(field_190000_cal/?,0) as field_190000_cal,");//--資產總額
       sqlCmd.append(" round(field_DEBIT/?,0) as field_DEBIT,");//--存款
       sqlCmd.append(" round(field_CREDIT/?,0) as field_CREDIT,");//--放款
       sqlCmd.append(" round(field_NET/?,0) as field_NET,");//--淨值
       sqlCmd.append(" decode(a01.field_CREDIT,0,0,round(a01.field_OVER /  a01.field_CREDIT *100 ,2))  as   field_OVER_RATE,");//--逾放比率
       sqlCmd.append(" round(field_OVER/?,0) as field_OVER,");//--逾期放款
       sqlCmd.append(" round(field_992510/?,0) as field_992510,");//--逾期放款-正會員
       sqlCmd.append(" round(field_992520/?,0) as field_992520,");//--逾期放款-贊助會員
       sqlCmd.append(" round(field_992530/?,0) as field_992530,");//--逾期放款-非會員
       sqlCmd.append(" round(field_992720/?,0) as field_992720,");//--建築放款中之逾期放款
       sqlCmd.append(" round(field_992550/?,0) as field_992550,");//--無擔保放款之逾期放款
       sqlCmd.append(" decode(a01.field_OVER,0,0,round(a01.field_BACKUP /  a01.field_OVER *100 ,2)) as   field_BACKUP_OVER_RATE,");//--備抵呆帳覆蓋率=備抵呆帳/逾期放款
       sqlCmd.append(" round(field_CAPTIAL /  1000 ,2)  as   field_CAPTIAL_RATE,");//--資本適足率
       sqlCmd.append(" round(field_420000_cal/?,0) as field_420000_cal,");//--收入
       sqlCmd.append(" round(field_520000_cal/?,0) as field_520000_cal,");//--出支
       sqlCmd.append(" round(field_320300/?,0) as field_320300,");//--本期損益
       sqlCmd.append(" round(field_992130/?,0) as field_992130,");//--存款業務.正會員
       sqlCmd.append(" round(field_990420/?,0) as field_990420,");//--存款業務.贊助會員
       sqlCmd.append(" round(field_990310/?,0) as field_990310,");//--存款業務.非會員
       sqlCmd.append(" round((field_992130+field_990420+field_990310)/ ?,0) as field_sum3_1,");//--存款業務.合計
       sqlCmd.append(" round(field_DEBIT/?,0) as field_DEBIT_1,");//--存款業務.總額 -->103.03.04 
       sqlCmd.append(" round(field_992140/?,0) as field_992140,");//--放款業務.放款-正會員
       sqlCmd.append(" round(field_990410/?,0) as field_990410,");//--放款業務.放款-贊助會員
       sqlCmd.append(" round((field_990610-field_990611)/ ?,0) as field_990610_990611,");//--放款業務.放款-非會員
       sqlCmd.append(" round(field_120700/?,0) as field_120700,");//--放款業務.放款-內部融資
       sqlCmd.append(" round(field_990611/?,0) as field_990611,");//--放款業務.放款-縣市政府貸款
       sqlCmd.append(" round((field_992140+field_990410+(field_990610-field_990611)+field_120700+field_990611)/ ?,0) as field_sum_3_2_1,");//--放款業務.放款-合計
       sqlCmd.append(" round(field_CREDIT/?,0) as field_CREDIT_1,");//--放款業務.放款-總額
       sqlCmd.append(" round(field_990510/?,0) as field_990510,");//--無擔保放款-非會員
       sqlCmd.append(" round(field_990510/?,0) as field_sum_3_2_2,");//--無擔保放款-合計
       sqlCmd.append(" round(field_NOASSURE/?,0) as field_NOASSURE,");//--無擔保放款-總額
       sqlCmd.append(" round(field_992710/?,0) as field_992710 ");//--建築放款-總額
       for(int i=1;i<=29;i++){
           paramList.add(unit);
       }
       sqlCmd.append(" from ");
       sqlCmd.append(" ( ");
       sqlCmd.append(" select  a01.bank_no ,   a01.BANK_NAME,");
       sqlCmd.append("          SUM(field_190000_cal)  field_190000_cal ,");
       sqlCmd.append("          SUM(field_420000_cal)  field_420000_cal ,");    
       sqlCmd.append("          SUM(field_520000_cal)  field_520000_cal ,");
       sqlCmd.append("          SUM(field_DEBIT)    field_DEBIT ,");
       sqlCmd.append("          SUM(field_CREDIT)   field_CREDIT,");        
       sqlCmd.append("          SUM(field_320300)   field_320300,");                 
       sqlCmd.append("          SUM(field_NET)      field_NET,");
       sqlCmd.append("          SUM(field_OVER)     field_OVER,");         
       sqlCmd.append("          SUM(field_BACKUP)   field_BACKUP,");       
       sqlCmd.append("          SUM(field_120700)   field_120700,");  
       sqlCmd.append("          SUM(field_NOASSURE) field_NOASSURE,");         
       sqlCmd.append("          SUM(field_990310)   field_990310,"); 
       sqlCmd.append("          SUM(field_990410)   field_990410,");
       sqlCmd.append("          SUM(field_990420)   field_990420,");
       sqlCmd.append("          SUM(field_990510)   field_990510,");
       sqlCmd.append("          SUM(field_990610)   field_990610,");
       sqlCmd.append("          SUM(field_990611)   field_990611,");
       sqlCmd.append("          SUM(field_CAPTIAL)  field_CAPTIAL,");        
       sqlCmd.append("          SUM(field_992130)   field_992130,");
       sqlCmd.append("          SUM(field_992140)   field_992140,");
       sqlCmd.append("          SUM(field_992510)   field_992510,");
       sqlCmd.append("          SUM(field_992520)   field_992520,");
       sqlCmd.append("          SUM(field_992530)   field_992530,");
       sqlCmd.append("          SUM(field_992550)   field_992550,");
       sqlCmd.append("          SUM(field_992710)   field_992710,");
       sqlCmd.append("          SUM(field_992720)   field_992720 ");
       sqlCmd.append(" from ");         
       sqlCmd.append(" (select  bn01.bank_no , bn01.BANK_NAME,");           
       sqlCmd.append("            round(sum(decode(bn01.bank_type,'6',decode(a01.acc_code,'190000',amt,0),'7',decode(a01.acc_code,'100000',amt,0),0)) /1,0)     as field_190000_cal,");//--資產總額
       sqlCmd.append("            round(sum(decode(a01.acc_code,'220000',amt,0)) /1,0) as field_DEBIT,");//--存款
       sqlCmd.append("            round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as  field_CREDIT,");//--放款
       sqlCmd.append("            round(sum(decode(bn01.bank_type,'6',decode(a01.acc_code,'310000',amt,'320000',amt,0),'7',decode(a01.acc_code,'300000',amt,0),0)) /1,0)     as field_NET,");//--淨值
       sqlCmd.append("            round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0) as field_OVER,");//--逾期放款
       sqlCmd.append("            round(sum(decode(bn01.bank_type,'6',decode(a01.acc_code,'420000',amt,0),'7',decode(a01.acc_code,'400000',amt,0),0)) /1,0)     as field_420000_cal,");//--收入
       sqlCmd.append("            round(sum(decode(bn01.bank_type,'6',decode(a01.acc_code,'520000',amt,0),'7',decode(a01.acc_code,'520000',amt,'522000',amt,0),0)) /1,0) as field_520000_cal,");//--支出
       sqlCmd.append("            round(sum(decode(a01.acc_code,'320300',amt,0)) /1,0) as field_320300,");//--本期損益           
       sqlCmd.append("            round(sum(decode(a01.acc_code, '120800',amt,'150300',amt,0)) /1,0) as  field_BACKUP,");//--備抵呆帳
       sqlCmd.append("            round(sum(decode(a01.acc_code,'120700',amt,0)) /1,0) as field_120700,");//--內部融資   
       sqlCmd.append("            round(sum( decode(YEAR_TYPE,'102',decode(bank_type,'6',decode(a01.acc_code, '120101',amt,'120301',amt, '120401',amt, '120501',amt,0),'7',decode(a01.acc_code, '120101',amt,'120401', amt, '120201',amt, '120501',amt,0)),");  
       sqlCmd.append("                                        '103',decode(a01.acc_code, '120101',amt,'120301',amt, '120401',amt, '120501',amt,0),0) ) /1,0) as  field_NOASSURE ");
       sqlCmd.append("          from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
       paramList.add(wlx01_m_year);
       sqlCmd.append("          left join (select (CASE WHEN (a01.m_year <= 102) THEN '102' ");                             
       sqlCmd.append("                                  WHEN (a01.m_year > 102) THEN '103' ");                             
       sqlCmd.append("                                  ELSE '00' END) as YEAR_TYPE,m_year,m_month,bank_code,acc_code,amt from a01 ");
       sqlCmd.append("                      where m_year= ?  and m_month=? "); 
       paramList.add(s_year);
       paramList.add(s_month);
       sqlCmd.append("                      ) a01  on  bn01.bank_no = a01.bank_code ");
       sqlCmd.append("           where a01.bank_code= ? ");
       paramList.add(bank_no);
       sqlCmd.append("          group by a01.m_year,a01.m_month,bn01.bank_no,bn01.BANK_NAME ");
       sqlCmd.append(" )a01,");
       sqlCmd.append(" ( select bn01.bank_no as bank_code, bn01.BANK_NAME,");
       sqlCmd.append("           round(sum(decode(acc_code,'990310',amt,0)) /1,0) as field_990310,");
       sqlCmd.append("           round(sum(decode(acc_code,'990410',amt,0)) /1,0) as field_990410,");
       sqlCmd.append("           round(sum(decode(acc_code,'990420',amt,0)) /1,0) as field_990420,");
       sqlCmd.append("           round(sum(decode(acc_code,'990510',amt,0)) /1,0) as field_990510,");
       sqlCmd.append("           round(sum(decode(acc_code,'990610',amt,0)) /1,0) as field_990610,");
       sqlCmd.append("           round(sum(decode(acc_code,'990611',amt,0)) /1,0) as field_990611 ");
       sqlCmd.append("         from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
       paramList.add(wlx01_m_year);
       sqlCmd.append("         left join (select * from a02 where m_year =? and m_month=?)a02 on bn01.bank_no = a02.bank_code ");
       sqlCmd.append("          where bank_code=? ");
       paramList.add(s_year);
       paramList.add(s_month);
       paramList.add(bank_no);
       sqlCmd.append("         group by bn01.bank_no,bn01.BANK_NAME ");
       sqlCmd.append(" ) a02,");
       sqlCmd.append(" (select bn01.bank_no as bank_code,  bn01.bank_name,");
       sqlCmd.append("           round(sum(decode(a05.acc_code,'91060P',amt,0)) /1,0) as field_CAPTIAL ");
       sqlCmd.append("        from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
       paramList.add(wlx01_m_year);
       sqlCmd.append("        left join (select * from a05 where m_year = ? and m_month=? and a05.ACC_code in ('91060P') ) a05 on  bn01.bank_no = a05.bank_code ");
       sqlCmd.append("        where a05.bank_code=? ");
       paramList.add(s_year);
       paramList.add(s_month);
       paramList.add(bank_no);
       sqlCmd.append("        group by bn01.bank_no,bn01.BANK_NAME ");
       sqlCmd.append(" ) a05,");
       sqlCmd.append(" ( select bn01.bank_no as bank_code, bn01.BANK_NAME,");
       sqlCmd.append("          round(sum(decode(a99.acc_code,'992130',amt,0)) /1,0) as field_992130,");
       sqlCmd.append("          round(sum(decode(a99.acc_code,'992140',amt,0)) /1,0) as field_992140,");
       sqlCmd.append("          round(sum(decode(a99.acc_code,'992510',amt,0)) /1,0) as field_992510,");
       sqlCmd.append("          round(sum(decode(a99.acc_code,'992520',amt,0)) /1,0) as field_992520,");
       sqlCmd.append("          round(sum(decode(a99.acc_code,'992530',amt,0)) /1,0) as field_992530,");
       sqlCmd.append("          round(sum(decode(a99.acc_code,'992550',amt,0)) /1,0) as field_992550,");
       sqlCmd.append("          round(sum(decode(a99.acc_code,'992710',amt,0)) /1,0) as field_992710,");
       sqlCmd.append("          round(sum(decode(a99.acc_code,'992720',amt,0)) /1,0) as field_992720 ");
       sqlCmd.append("         from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
       paramList.add(wlx01_m_year);
       sqlCmd.append("         left join (select * from a99 where m_year = ? and m_month=? )a99 on bn01.bank_no = a99.bank_code ");
       sqlCmd.append("          where bank_code=? ");
       paramList.add(s_year);
       paramList.add(s_month);
       paramList.add(bank_no);
       sqlCmd.append("         group by bn01.bank_no,bn01.BANK_NAME ");
       sqlCmd.append(" ) a99 ");
       sqlCmd.append(" where a01.bank_no = a02.bank_code(+) and a01.bank_no = a05.bank_code(+)  and a01.bank_no=a99.bank_code(+) ");
       sqlCmd.append(" and a01.bank_no=? ");
       paramList.add(bank_no);
       sqlCmd.append(" GROUP BY a01.bank_no,a01.BANK_NAME ");  
       sqlCmd.append(" )a01 ");

       List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"bank_no,bank_name,field_190000_cal,field_debit,field_credit,field_net,field_over_rate,"
                                                                           +"field_over,field_992510,field_992520,field_992530,field_992720,field_992550,"
                                                                           +"field_backup_over_rate,field_captial_rate,field_420000_cal,field_520000_cal,"
                                                                           +"field_320300,field_992130,field_990420,field_990310,field_sum3_1,field_debit_1,"
                                                                           +"field_992140,field_990410,field_990610_990611,field_120700,field_990611,"
                                                                           +"field_sum_3_2_1,field_credit_1,field_990510,field_sum_3_2_2,field_noassure,field_992710");
       System.out.println("dbData_mainInfo.size()="+dbData.size()); 
       return dbData;
   }
   //111.02.24 調整直接讀取excel格式
   private static void insertCell_Date(String value,HSSFWorkbook wb,HSSFRow row,int i,short leftBorder,short rightBorder){
       HSSFCell cell=(row.getCell((short)i)==null)? row.createCell((short)i) : row.getCell((short)i);
       //HSSFCellStyle cs1 = wb.createCellStyle();//111.02.24
       //HSSFCellStyle cs1 = cell.getCellStyle();//會套用原本excel所設定的格式
       //設置邊框
       //cs1.setBorderTop(HSSFCellStyle.BORDER_THIN); //上邊框
       //cs1.setBorderBottom(HSSFCellStyle.BORDER_THIN); //下邊框
       //cs1.setBorderLeft(leftBorder); //左邊框
       //cs1.setBorderRight(rightBorder); //右邊框
       //cs1.setWrapText(true);//自動換行
       //cs1.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);//垂直置中
       //cs1.setAlignment(HSSFCellStyle.ALIGN_RIGHT);//水平
       //cell.setCellStyle(cs1);
       cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
       cell.setCellValue(value);
   }
   //111.02.24 調整直接讀取excel格式
   private static void insertCell(String value,HSSFWorkbook wb,HSSFRow row,int i,short border, short alignment){
           HSSFCell cell=(row.getCell((short)i)==null)? row.createCell((short)i) : row.getCell((short)i);
           //HSSFCellStyle cs1 = wb.createCellStyle();
           //HSSFCellStyle cs1 = cell.getCellStyle();//會套用原本excel所設定的格式
           //設置邊框
           //cs1.setBorderTop(border); //上邊框
           //cs1.setBorderBottom(border); //下邊框
           //cs1.setBorderLeft(border); //左邊框
           //cs1.setBorderRight(border); //右邊框
           //cs1.setWrapText(true);//自動換行
           //cs1.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);//垂直置中
           //cs1.setAlignment(alignment);//水平
           //cell.setCellStyle(cs1);
           cell.setEncoding( HSSFCell.ENCODING_UTF_16 );
           cell.setCellValue(value);
   }
   private static void setDate(HSSFSheet sheet,HSSFWorkbook wb,HSSFRow row,int cellNo,short leftBorder,short rightBorder,String m_year,String m_month){
       row=(sheet.getRow((short)3)==null)? sheet.createRow((short)3) : sheet.getRow((short)3);
       insertCell_Date(m_year+m_month,wb,row,cellNo,leftBorder,rightBorder);
       row=(sheet.getRow((short)9)==null)? sheet.createRow((short)9) : sheet.getRow((short)9);
       insertCell_Date(m_year+m_month,wb,row,cellNo,leftBorder,rightBorder);
       row=(sheet.getRow((short)19)==null)? sheet.createRow((short)19) : sheet.getRow((short)19);
       insertCell_Date(m_year+m_month,wb,row,cellNo,leftBorder,rightBorder);
       row=(sheet.getRow((short)22)==null)? sheet.createRow((short)22) : sheet.getRow((short)22);
       insertCell_Date(m_year+m_month,wb,row,cellNo,leftBorder,rightBorder);
       row=(sheet.getRow((short)28)==null)? sheet.createRow((short)28) : sheet.getRow((short)28);
       insertCell_Date(m_year+m_month,wb,row,cellNo,leftBorder,rightBorder);
       row=(sheet.getRow((short)36)==null)? sheet.createRow((short)36) : sheet.getRow((short)36);
       insertCell_Date(m_year+m_month,wb,row,cellNo,leftBorder,rightBorder);
       row=(sheet.getRow((short)45)==null)? sheet.createRow((short)45) : sheet.getRow((short)45);
       insertCell_Date(m_year+m_month,wb,row,cellNo,leftBorder,rightBorder);
       row=(sheet.getRow((short)52)==null)? sheet.createRow((short)52) : sheet.getRow((short)52);
       insertCell_Date(m_year+m_month,wb,row,cellNo,leftBorder,rightBorder);
   }
   private static void setValues(HSSFSheet sheet,HSSFWorkbook wb,HSSFRow row,DataObject bean,int cellNo){
       field_190000_cal = (bean.getValue("field_190000_cal")==null)?"0":Utility.setCommaFormat(bean.getValue("field_190000_cal").toString());
       field_debit = (bean.getValue("field_debit")==null)?"0":Utility.setCommaFormat(bean.getValue("field_debit").toString());
       field_credit = (bean.getValue("field_credit")==null)?"0":Utility.setCommaFormat(bean.getValue("field_credit").toString());
       field_net = (bean.getValue("field_net")==null)?"0":Utility.setCommaFormat(bean.getValue("field_net").toString());
       field_over_rate = (bean.getValue("field_over_rate")==null)?"0.00":Utility.setCommaFormat(bean.getValue("field_over_rate").toString());
       field_over = (bean.getValue("field_over")==null)?"0":Utility.setCommaFormat(bean.getValue("field_over").toString());
       field_992510 = (bean.getValue("field_992510")==null)?"0":Utility.setCommaFormat(bean.getValue("field_992510").toString());
       field_992520 = (bean.getValue("field_992520")==null)?"0":Utility.setCommaFormat(bean.getValue("field_992520").toString());
       field_992530 = (bean.getValue("field_992530")==null)?"0":Utility.setCommaFormat(bean.getValue("field_992530").toString());
       field_992720 = (bean.getValue("field_992720")==null)?"0":Utility.setCommaFormat(bean.getValue("field_992720").toString());
       field_992550 = (bean.getValue("field_992550")==null)?"0":Utility.setCommaFormat(bean.getValue("field_992550").toString());
       field_backup_over_rate = (bean.getValue("field_backup_over_rate")==null)?"0.00":Utility.setCommaFormat(bean.getValue("field_backup_over_rate").toString());
       field_captial_rate = (bean.getValue("field_captial_rate")==null)?"0.00":Utility.setCommaFormat(bean.getValue("field_captial_rate").toString());
       field_420000_cal = (bean.getValue("field_420000_cal")==null)?"0":Utility.setCommaFormat(bean.getValue("field_420000_cal").toString());
       field_520000_cal = (bean.getValue("field_520000_cal")==null)?"0":Utility.setCommaFormat(bean.getValue("field_520000_cal").toString());
       field_320300 = (bean.getValue("field_320300")==null)?"0":Utility.setCommaFormat(bean.getValue("field_320300").toString());
       field_992130 = (bean.getValue("field_992130")==null)?"0":Utility.setCommaFormat(bean.getValue("field_992130").toString());
       field_990420 = (bean.getValue("field_990420")==null)?"0":Utility.setCommaFormat(bean.getValue("field_990420").toString());
       field_990310 = (bean.getValue("field_990310")==null)?"0":Utility.setCommaFormat(bean.getValue("field_990310").toString());
       field_debit_1 = (bean.getValue("field_debit_1")==null)?"0":Utility.setCommaFormat(bean.getValue("field_debit_1").toString());
       field_992140 = (bean.getValue("field_992140")==null)?"0":Utility.setCommaFormat(bean.getValue("field_992140").toString());
       field_990410 = (bean.getValue("field_990410")==null)?"0":Utility.setCommaFormat(bean.getValue("field_990410").toString());
       field_990610_990611 = (bean.getValue("field_990610_990611")==null)?"0":Utility.setCommaFormat(bean.getValue("field_990610_990611").toString());
       field_120700 = (bean.getValue("field_120700")==null)?"0":Utility.setCommaFormat(bean.getValue("field_120700").toString());
       field_990611 = (bean.getValue("field_990611")==null)?"0":Utility.setCommaFormat(bean.getValue("field_990611").toString());
       field_credit_1 = (bean.getValue("field_credit_1")==null)?"0":Utility.setCommaFormat(bean.getValue("field_credit_1").toString());
       field_990510 = (bean.getValue("field_990510")==null)?"0":Utility.setCommaFormat(bean.getValue("field_990510").toString());
       field_noassure = (bean.getValue("field_noassure")==null)?"0":Utility.setCommaFormat(bean.getValue("field_noassure").toString());
       field_992710 = (bean.getValue("field_992710")==null)?"0":Utility.setCommaFormat(bean.getValue("field_992710").toString());
       field_sum3_1 = (bean.getValue("field_sum3_1")==null)?"0":Utility.setCommaFormat(bean.getValue("field_sum3_1").toString());
       field_sum_3_2_1 = (bean.getValue("field_sum_3_2_1")==null)?"0":Utility.setCommaFormat(bean.getValue("field_sum_3_2_1").toString());
       field_sum_3_2_2 = (bean.getValue("field_sum_3_2_2")==null)?"0":Utility.setCommaFormat(bean.getValue("field_sum_3_2_2").toString());
       //一、財務狀況(一)主要資產與負債
       row=(sheet.getRow((short)4)==null)? sheet.createRow((short)4) : sheet.getRow((short)4);
       insertCell(field_190000_cal,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       row=(sheet.getRow((short)5)==null)? sheet.createRow((short)5) : sheet.getRow((short)5);
       insertCell(field_debit,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       row=(sheet.getRow((short)6)==null)? sheet.createRow((short)6) : sheet.getRow((short)6);
       insertCell(field_credit,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       row=(sheet.getRow((short)7)==null)? sheet.createRow((short)7) : sheet.getRow((short)7);
       insertCell(field_net,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       //一、財務狀況(二)資產品質
       row=(sheet.getRow((short)10)==null)? sheet.createRow((short)10) : sheet.getRow((short)10);
       insertCell(field_over_rate,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       row=(sheet.getRow((short)11)==null)? sheet.createRow((short)11) : sheet.getRow((short)11);
       insertCell(field_over,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       row=(sheet.getRow((short)12)==null)? sheet.createRow((short)12) : sheet.getRow((short)12);
       insertCell(field_992510,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       row=(sheet.getRow((short)13)==null)? sheet.createRow((short)13) : sheet.getRow((short)13);
       insertCell(field_992520,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       row=(sheet.getRow((short)14)==null)? sheet.createRow((short)14) : sheet.getRow((short)14);
       insertCell(field_992530,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       row=(sheet.getRow((short)15)==null)? sheet.createRow((short)15) : sheet.getRow((short)15);
       insertCell(field_992720,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       row=(sheet.getRow((short)16)==null)? sheet.createRow((short)16) : sheet.getRow((short)16);
       insertCell(field_992550,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       row=(sheet.getRow((short)17)==null)? sheet.createRow((short)17) : sheet.getRow((short)17);
       insertCell(field_backup_over_rate,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       //一、財務狀況(三)資本適足性
       row=(sheet.getRow((short)20)==null)? sheet.createRow((short)20) : sheet.getRow((short)20);
       insertCell(field_captial_rate,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       //二、經營績效  
       row=(sheet.getRow((short)23)==null)? sheet.createRow((short)23) : sheet.getRow((short)23);
       insertCell(field_420000_cal,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       row=(sheet.getRow((short)24)==null)? sheet.createRow((short)24) : sheet.getRow((short)24);
       insertCell(field_520000_cal,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       row=(sheet.getRow((short)25)==null)? sheet.createRow((short)25) : sheet.getRow((short)25);
       insertCell(field_320300,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       //三、主要業務(一)存款業務
       row=(sheet.getRow((short)29)==null)? sheet.createRow((short)29) : sheet.getRow((short)29);
       insertCell(field_992130,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       row=(sheet.getRow((short)30)==null)? sheet.createRow((short)30) : sheet.getRow((short)30);
       insertCell(field_990420,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       row=(sheet.getRow((short)31)==null)? sheet.createRow((short)31) : sheet.getRow((short)31);
       insertCell(field_990310,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       row=(sheet.getRow((short)32)==null)? sheet.createRow((short)32) : sheet.getRow((short)32);
       insertCell(field_sum3_1,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       row=(sheet.getRow((short)33)==null)? sheet.createRow((short)33) : sheet.getRow((short)33);
       insertCell(field_debit_1,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       //三、主要業務(二)放款業務1.放款總額
       row=(sheet.getRow((short)37)==null)? sheet.createRow((short)37) : sheet.getRow((short)37);
       insertCell(field_992140,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       row=(sheet.getRow((short)38)==null)? sheet.createRow((short)38) : sheet.getRow((short)38);
       insertCell(field_990410,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       row=(sheet.getRow((short)39)==null)? sheet.createRow((short)39) : sheet.getRow((short)39);
       insertCell(field_990610_990611,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       row=(sheet.getRow((short)40)==null)? sheet.createRow((short)40) : sheet.getRow((short)40);
       insertCell(field_120700,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       row=(sheet.getRow((short)41)==null)? sheet.createRow((short)41) : sheet.getRow((short)41);
       insertCell(field_990611,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       row=(sheet.getRow((short)42)==null)? sheet.createRow((short)42) : sheet.getRow((short)42);
       insertCell(field_sum_3_2_1,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       row=(sheet.getRow((short)43)==null)? sheet.createRow((short)43) : sheet.getRow((short)43);
       insertCell(field_credit_1,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       //三、主要業務(二)放款業務2.無擔保放款
       row=(sheet.getRow((short)48)==null)? sheet.createRow((short)48) : sheet.getRow((short)48);
       insertCell(field_990510,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       row=(sheet.getRow((short)49)==null)? sheet.createRow((short)50) : sheet.getRow((short)49);
       insertCell(field_sum_3_2_2,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       row=(sheet.getRow((short)50)==null)? sheet.createRow((short)50) : sheet.getRow((short)50);
       insertCell(field_noassure,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       //三、主要業務(二)放款業務3.建築放款
       row=(sheet.getRow((short)56)==null)? sheet.createRow((short)56) : sheet.getRow((short)56);
       insertCell("",wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
       row=(sheet.getRow((short)57)==null)? sheet.createRow((short)57) : sheet.getRow((short)57);
       insertCell(field_992710,wb,row,cellNo,HSSFCellStyle.BORDER_THIN,HSSFCellStyle.ALIGN_RIGHT);
   }
   //取得總機構代碼
   private static String getBank_name(String bank_no,String m_year){
       String bank_name = "";
       StringBuffer sqlCmd = new StringBuffer(); 
       List paramList = new ArrayList();
       sqlCmd.append(" select bank_name from bn01 where bank_no = ? and m_year = ? ");
       paramList.add(bank_no);
       paramList.add(m_year);
       List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"bank_name");
       if(dbData.size()>0){
           bank_name = Utility.getTrimString(((DataObject)dbData.get(0)).getValue("bank_name"));
       }
       return bank_name;
   }
}

