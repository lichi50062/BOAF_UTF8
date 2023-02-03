/*
    101.08 created
    101.11.14 add 存款總額/放款總額/逾放金額/逾放比率增加顯示90年底資料 by 2295  
    101.12.20 fix 利率顯示到小數點2位 by 2295			      
*/
package com.tradevan.util.report;

import org.apache.poi.hssf.model.Sheet;
import org.apache.poi.hssf.usermodel.*;
import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.io.FileInputStream;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.hssf.usermodel.*;
import java.io.FileInputStream; 
import java.io.FileOutputStream; 
import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR001WD {
    public static String createRpt(String s_year,String s_month,String unit,String bank_type){
		String errMsg = "";
		StringBuffer sqlCmd = new StringBuffer();
		List paramList = new ArrayList();
		StringBuffer sqlCmd1 = new StringBuffer();
        List paramList1 = new ArrayList(); 
		int rowNum1=0;
		int rowNum2=0;
		int j=0;
		String last_year="";
        String last_month="";
		String filename="農漁會信用部營運概況.xls";
		//===============================
	    String wlx01_m_year = (Integer.parseInt(s_year) < 100)?"99":"100"; 
	    Calendar now = Calendar.getInstance();
	    String now_YEAR  = String.valueOf(now.get(Calendar.YEAR)-1911); //回覆值為西元年故需-1911取得民國年;
	    //=====================================================================    
	   
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
    		
    		String openfile="農漁會信用部營運概況.xls";
    		System.out.println("open file "+openfile);
			FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+ openfile );
	  	    //設定FileINputStream讀取Excel檔
			POIFSFileSystem fs = new POIFSFileSystem(finput); //讀取檔案
            HSSFWorkbook wb = new HSSFWorkbook(fs);//建立活頁簿
            HSSFSheet sheet = wb.getSheetAt(0);// 讀取第一個工作表，宣告其為sheet
            HSSFSheet sheet1 = wb.getSheetAt(1);// 讀取第2個工作表，宣告其為sheet1
            HSSFSheet sheet2 = wb.getSheetAt(2);// 讀取第3個工作表，宣告其為sheet2
            HSSFPrintSetup ps = sheet.getPrintSetup(); // 取得設定
            HSSFPrintSetup ps1 = sheet1.getPrintSetup();
            HSSFPrintSetup ps2 = sheet2.getPrintSetup();
	        //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	        //sheet.setAutobreaks(true); //自動分頁

	  		//設定頁面符合列印大小
	        sheet.setAutobreaks( false );
	        ps.setScale( ( short )90 ); //列印縮放百分比
	        ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
	        sheet1.setAutobreaks( false );
            ps1.setScale( ( short )90 ); 
            ps1.setPaperSize( ( short )9 );
            ps2.setScale( ( short )90 ); 
            ps2.setPaperSize( ( short )9 ); 
	  		//wb.setSheetName(0,"test");
	  		finput.close();
	  		HSSFRow row=null;//宣告一列
	  		HSSFCell cell=null;//宣告一個儲存格
	  		
	  		reportUtil reportUtil = new reportUtil();
            HSSFCellStyle cs_right = reportUtil.getRightStyle(wb);
            HSSFCellStyle cs_center = reportUtil.getDefaultStyle(wb);
            HSSFCellStyle nb_left = reportUtil.getNoBorderLeftStyle(wb);
            HSSFCellStyle nb_center = reportUtil.getNoBorderDefaultStyle(wb);
            HSSFCellStyle nb_right = reportUtil.getNoBoderStyle(wb);//無框置右 
            
            sqlCmd.append(" select * from rpt_month ");
            sqlCmd.append("        where m_year = ? and m_month = ? ");
            sqlCmd.append("        and report_no= ? ");
            paramList.add(s_year);
            paramList.add(s_month);
            paramList.add("FR001WD");
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"");
           //如果有資料，表示該年月資料已關帳，則所有資料讀取關帳檔(rpt_month)  101.08.27 add
            
            if(s_month.equals("1")){//若本月為1月份是..則是上個年度的12月份
                last_year = String.valueOf(Integer.parseInt(s_year) - 1);
                last_month = "12";
            }else{  
                last_year = String.valueOf(Integer.parseInt(s_year));
                last_month = String.valueOf(Integer.parseInt(s_month) - 1);//上個月份的
            }
            
            
	  		//查詢年度以前只取91年度以後12月份資料統計資料
            sqlCmd.setLength(0) ;
            paramList.clear() ;
            sqlCmd.append(" select rpt_month.m_year, rpt_month.m_month, bn01_month.bank_sum_6, bn01_month.bank_sum_7,");
            sqlCmd.append("        field_190000,");//--農會總資產
            sqlCmd.append("        field_100000,");//--漁會總資產
            sqlCmd.append("        field_asset,");//--農漁會總資產
            sqlCmd.append("        round(field_asset/?,0) as field_asset_unit,");//--農漁會總資產取到億元 
            sqlCmd.append("        field_310000,");//--農會事業資金及公積 
            sqlCmd.append("        field_320000,");//--農會盈虧及損益 
            sqlCmd.append("        field_300000,");//--漁會淨值
            sqlCmd.append("        field_networth,");//--農漁會淨值
            sqlCmd.append("        round(field_networth/?,0) as field_networth_unit,");//--農漁會淨值取到億元
            sqlCmd.append("        field_220000_6,");//--農會存款 
            sqlCmd.append("        field_220000_7,");//--漁會存款 
            sqlCmd.append("        field_debit,");//--農漁會存款
            sqlCmd.append("        round(field_debit/?,0) as field_debit_unit,");//--農漁會存款總額取到億元 
            sqlCmd.append("        field_120000_6,");//--農會放款 
            sqlCmd.append("        field_120000_7,");//--漁會放款 
            sqlCmd.append("        field_120800_6,");//--農會備抵呆帳-放款 
            sqlCmd.append("        field_120800_7,");//--漁會備抵呆帳-放款 
            sqlCmd.append("        field_150300_6,");//--農會備抵呆帳-催收款項 
            sqlCmd.append("        field_150300_7,");//--漁會備抵呆帳-催收款項 
            sqlCmd.append("        field_credit,");//--農漁會放款總額
            sqlCmd.append("        round(field_credit/?,0) as field_credit_unit,");//--農漁會放款總額取到億元
            sqlCmd.append("        field_990000_6,");//--農會.狹義逾期放款金額 
            sqlCmd.append("        field_990000_7,");//--漁會.狹義逾期放款金額
            sqlCmd.append("        field_990000,");//--農漁會.狹義逾期放款金額
            sqlCmd.append("        round(field_990000/?,0) as field_990000_unit,");//--農漁會.狹義逾期放款金額.取到億元  
            sqlCmd.append("        field_840740_6,");//--農會.廣義逾期放款
            sqlCmd.append("        field_840740_7,");//--漁會.廣義逾期放款
            sqlCmd.append("        field_840740,");//--農漁會.廣義逾期放款
            sqlCmd.append("        round(field_840740/?,0) as field_840740_unit,");//--農漁會.廣義逾期放款.取到億元
            sqlCmd.append("        field_990000_rate,");//--狹義逾放比率
            sqlCmd.append("        field_840740_rate ");//--廣義逾放比率
            paramList.add(unit);
            paramList.add(unit);
            paramList.add(unit);
            paramList.add(unit);
            paramList.add(unit);
            paramList.add(unit);
            sqlCmd.append(" from ( ");
            sqlCmd.append("  select m_year,m_month,");
            sqlCmd.append("         sum(decode(bank_type,'6',decode(acc_code,'field_190000',amt,0))) as field_190000,");
            sqlCmd.append("         sum(decode(bank_type,'7',decode(acc_code,'field_100000',amt,0))) as field_100000,");
            sqlCmd.append("         sum(decode(bank_type,'ALL',decode(acc_code,'field_asset',amt,0))) as field_asset,");
            sqlCmd.append("         sum(decode(bank_type,'6',decode(acc_code,'field_310000',amt,0))) as field_310000,");
            sqlCmd.append("         sum(decode(bank_type,'6',decode(acc_code,'field_320000',amt,0))) as field_320000,");
            sqlCmd.append("         sum(decode(bank_type,'7',decode(acc_code,'field_300000',amt,0))) as field_300000,");
            sqlCmd.append("         sum(decode(bank_type,'ALL',decode(acc_code,'field_networth',amt,0))) as field_NETWorth,");
            sqlCmd.append("         sum(decode(bank_type,'6',decode(acc_code,'field_220000_6',amt,0))) as field_220000_6,");
            sqlCmd.append("         sum(decode(bank_type,'7',decode(acc_code,'field_220000_7',amt,0))) as field_220000_7,");
            sqlCmd.append("         sum(decode(bank_type,'ALL',decode(acc_code,'field_debit',amt,0))) as field_DEBIT,");   
            sqlCmd.append("         sum(decode(bank_type,'6',decode(acc_code,'field_120000_6',amt,0))) as field_120000_6,");
            sqlCmd.append("         sum(decode(bank_type,'7',decode(acc_code,'field_120000_7',amt,0))) as field_120000_7,");
            sqlCmd.append("         sum(decode(bank_type,'6',decode(acc_code,'field_120800_6',amt,0))) as field_120800_6,");
            sqlCmd.append("         sum(decode(bank_type,'7',decode(acc_code,'field_120800_7',amt,0))) as field_120800_7,");
            sqlCmd.append("         sum(decode(bank_type,'6',decode(acc_code,'field_150300_6',amt,0))) as field_150300_6,");
            sqlCmd.append("         sum(decode(bank_type,'7',decode(acc_code,'field_150300_7',amt,0))) as field_150300_7,");
            sqlCmd.append("         sum(decode(bank_type,'ALL',decode(acc_code,'field_credit',amt,0))) as field_CREDIT, ");
            sqlCmd.append("         sum(decode(bank_type,'6',decode(acc_code,'field_990000_6',amt,0))) as field_990000_6,");
            sqlCmd.append("         sum(decode(bank_type,'7',decode(acc_code,'field_990000_7',amt,0))) as field_990000_7,");
            sqlCmd.append("         sum(decode(bank_type,'ALL',decode(acc_code,'field_990000',amt,0))) as field_990000,");
            sqlCmd.append("         sum(decode(bank_type,'6',decode(acc_code,'field_840740_6',amt,0))) as field_840740_6,");
            sqlCmd.append("         sum(decode(bank_type,'7',decode(acc_code,'field_840740_7',amt,0))) as field_840740_7,");
            sqlCmd.append("         sum(decode(bank_type,'ALL',decode(acc_code,'field_840740',amt,0))) as field_840740,");
            sqlCmd.append("         sum(decode(bank_type,'ALL',decode(acc_code,'field_990000_rate',amt,0))) as field_990000_rate,");
            sqlCmd.append("         sum(decode(bank_type,'ALL',decode(acc_code,'field_840740_rate',amt,0))) as field_840740_rate ");
            sqlCmd.append("   from (select * from rpt_month where (m_year >= 90 and m_year  <= ?) and m_month=12 and report_no= ? ");
            paramList.add(Integer.parseInt(s_year)-1);
            paramList.add("FR001WD");
            sqlCmd.append("   )rpt_month ");
            sqlCmd.append("   group by m_year,m_month ");
            sqlCmd.append("  )rpt_month, ");
            sqlCmd.append("  (select * ");
            sqlCmd.append(" from bn01_month where m_month=12 )bn01_month ");
            sqlCmd.append(" where rpt_month.m_year = bn01_month.m_year(+) and rpt_month.m_month= bn01_month.m_month(+)");
            sqlCmd.append(" union");
            //查詢年度,前一月份(含)以前(關帳資料)        
            sqlCmd.append(" select  rpt_month.m_year,rpt_month.m_month,bn01_month.bank_sum_6,bn01_month.bank_sum_7,");
            sqlCmd.append("         field_190000,");//--農會總資產
            sqlCmd.append("         field_100000,");//--漁會總資產
            sqlCmd.append("         field_asset,");//--農漁會總資產
            sqlCmd.append("         round(field_asset/?,0) as field_asset_unit,");//--農漁會總資產取到億元 
            sqlCmd.append("         field_310000,");//--農會事業資金及公積 
            sqlCmd.append("         field_320000,");//--農會盈虧及損益 
            sqlCmd.append("         field_300000,");//--漁會淨值
            sqlCmd.append("         field_networth,");//--農漁會淨值
            sqlCmd.append("         round(field_networth/?,0) as field_networth_unit,");//--農漁會淨值取到億元
            sqlCmd.append("         field_220000_6,");//--農會存款 
            sqlCmd.append("         field_220000_7,");//--漁會存款 
            sqlCmd.append("         field_debit,");//--農漁會存款
            sqlCmd.append("         round(field_debit/?,0) as field_debit_unit,");//--農漁會存款總額取到億元 
            sqlCmd.append("         field_120000_6,");//--農會放款 
            sqlCmd.append("         field_120000_7,");//--漁會放款 
            sqlCmd.append("         field_120800_6,");//--農會備抵呆帳-放款 
            sqlCmd.append("         field_120800_7,");//--漁會備抵呆帳-放款 
            sqlCmd.append("         field_150300_6,");//--農會備抵呆帳-催收款項 
            sqlCmd.append("         field_150300_7,");//--漁會備抵呆帳-催收款項 
            sqlCmd.append("         field_credit,");//--農漁會放款總額
            sqlCmd.append("         round(field_credit/?,0) as field_credit_unit,");//--農漁會放款總額取到億元
            sqlCmd.append("         field_990000_6,");//--農會.狹義逾期放款金額 
            sqlCmd.append("         field_990000_7,");//--漁會.狹義逾期放款金額
            sqlCmd.append("         field_990000,");//--農漁會.狹義逾期放款金額
            sqlCmd.append("         round(field_990000/?,0) as field_990000_unit,");//--農漁會.狹義逾期放款金額.取到億元  
            sqlCmd.append("         field_840740_6,");//--農會.廣義逾期放款
            sqlCmd.append("         field_840740_7,");//--漁會.廣義逾期放款
            sqlCmd.append("         field_840740,");//--農漁會.廣義逾期放款
            sqlCmd.append("         round(field_840740/?,0) as field_840740_unit,");//--農漁會.廣義逾期放款.取到億元
            sqlCmd.append("         field_990000_rate,");//--狹義逾放比率
            sqlCmd.append("         field_840740_rate ");//--廣義逾放比率
            paramList.add(unit);
            paramList.add(unit);
            paramList.add(unit);
            paramList.add(unit);
            paramList.add(unit);
            paramList.add(unit);
            sqlCmd.append(" from (  ");
            sqlCmd.append("     select  m_year,m_month,");
            sqlCmd.append("             sum(decode(bank_type,'6',decode(acc_code,'field_190000',amt,0)))     as field_190000,");
            sqlCmd.append("             sum(decode(bank_type,'7',decode(acc_code,'field_100000',amt,0)))     as field_100000,");
            sqlCmd.append("             sum(decode(bank_type,'ALL',decode(acc_code,'field_asset',amt,0)))    as field_asset,");
            sqlCmd.append("             sum(decode(bank_type,'6',decode(acc_code,'field_310000',amt,0)))     as field_310000,");
            sqlCmd.append("             sum(decode(bank_type,'6',decode(acc_code,'field_320000',amt,0)))     as field_320000,");
            sqlCmd.append("             sum(decode(bank_type,'7',decode(acc_code,'field_300000',amt,0)))     as field_300000,");
            sqlCmd.append("             sum(decode(bank_type,'ALL',decode(acc_code,'field_networth',amt,0))) as field_NETWorth,");
            sqlCmd.append("             sum(decode(bank_type,'6',decode(acc_code,'field_220000_6',amt,0)))   as field_220000_6,");
            sqlCmd.append("             sum(decode(bank_type,'7',decode(acc_code,'field_220000_7',amt,0)))   as field_220000_7,");
            sqlCmd.append("             sum(decode(bank_type,'ALL',decode(acc_code,'field_debit',amt,0)))    as field_DEBIT, ");     
            sqlCmd.append("             sum(decode(bank_type,'6',decode(acc_code,'field_120000_6',amt,0)))   as field_120000_6,");
            sqlCmd.append("             sum(decode(bank_type,'7',decode(acc_code,'field_120000_7',amt,0)))   as field_120000_7,");
            sqlCmd.append("             sum(decode(bank_type,'6',decode(acc_code,'field_120800_6',amt,0)))   as field_120800_6,");
            sqlCmd.append("             sum(decode(bank_type,'7',decode(acc_code,'field_120800_7',amt,0)))   as field_120800_7,");
            sqlCmd.append("             sum(decode(bank_type,'6',decode(acc_code,'field_150300_6',amt,0)))   as field_150300_6,");
            sqlCmd.append("             sum(decode(bank_type,'7',decode(acc_code,'field_150300_7',amt,0)))   as field_150300_7,");
            sqlCmd.append("             sum(decode(bank_type,'ALL',decode(acc_code,'field_credit',amt,0)))   as field_CREDIT, ");     
            sqlCmd.append("             sum(decode(bank_type,'6',decode(acc_code,'field_990000_6',amt,0)))   as field_990000_6,");
            sqlCmd.append("             sum(decode(bank_type,'7',decode(acc_code,'field_990000_7',amt,0)))   as field_990000_7,");
            sqlCmd.append("             sum(decode(bank_type,'ALL',decode(acc_code,'field_990000',amt,0)))   as field_990000, ");     
            sqlCmd.append("             sum(decode(bank_type,'6',decode(acc_code,'field_840740_6',amt,0)))   as field_840740_6,");
            sqlCmd.append("             sum(decode(bank_type,'7',decode(acc_code,'field_840740_7',amt,0)))   as field_840740_7,");
            sqlCmd.append("             sum(decode(bank_type,'ALL',decode(acc_code,'field_840740',amt,0)))   as field_840740,  ");    
            sqlCmd.append("             sum(decode(bank_type,'ALL',decode(acc_code,'field_990000_rate',amt,0))) as field_990000_rate,");
            sqlCmd.append("             sum(decode(bank_type,'ALL',decode(acc_code,'field_840740_rate',amt,0))) as field_840740_rate ");
            sqlCmd.append("     from (select * from rpt_month ");
            sqlCmd.append("             where m_year = ?  and m_month <= ?  and report_no= ? ");
            if(dbData.size() == 0){
                paramList.add(last_year);
                paramList.add(last_month);
            }else{
                paramList.add(s_year);
                paramList.add(s_month);
            }
            paramList.add("FR001WD");
            sqlCmd.append("     )rpt_month ");
            sqlCmd.append("     group by m_year,m_month ");
            sqlCmd.append("     )rpt_month, ");
            sqlCmd.append("     (select * ");
            sqlCmd.append("     from bn01_month  where m_year = ? and m_month <= ? )bn01_month ");
            if(dbData.size() == 0){
                paramList.add(last_year);
                paramList.add(last_month);
            }else{
                paramList.add(s_year);
                paramList.add(s_month);
            }
            sqlCmd.append("     where rpt_month.m_year = bn01_month.m_year(+) and rpt_month.m_month= bn01_month.m_month(+) ");
            
            
            
            //若查詢年月資料，未關帳時，顯示即時資料，才加入此段SQL 101.08.27 add
            if(dbData.size() == 0){
                sqlCmd.append("     union ");
                //--查詢年度.當月份,顯示即時資料,增加計算表 101.08.06 add
                sqlCmd.append("     select  a01.m_year,a01.m_month,bn01_month.bank_sum_6,bn01_month.bank_sum_7,");
                sqlCmd.append("             field_190000,");//--農會總資產
                sqlCmd.append("             field_100000,");//--漁會總資產
                sqlCmd.append("             field_asset,");//--農漁會總資產
                sqlCmd.append("             round(field_asset/?,0) as field_asset,");//--農漁會總資產取到億元
                sqlCmd.append("             field_310000,");//--農會事業資金及公積 
                sqlCmd.append("             field_320000,");//--農會盈虧及損益 
                sqlCmd.append("             field_300000,");//--漁會淨值
                sqlCmd.append("             field_networth,");//--農漁會淨值
                sqlCmd.append("             round(field_networth/?,0) as field_networth,");//--農漁會淨值取到億元
                sqlCmd.append("             field_220000_6,");//--農會存款 
                sqlCmd.append("             field_220000_7,");//--漁會存款 
                sqlCmd.append("             field_debit,");//--農漁會存款
                sqlCmd.append("             round(field_debit/?,0) as field_debit,");//--農漁會存款總額取到億元
                sqlCmd.append("             field_120000_6,");//--農會放款 
                sqlCmd.append("             field_120000_7,");//--漁會放款 
                sqlCmd.append("             field_120800_6,");//--農會備抵呆帳-放款 
                sqlCmd.append("             field_120800_7,");//--漁會備抵呆帳-放款 
                sqlCmd.append("             field_150300_6,");//--農會備抵呆帳-催收款項 
                sqlCmd.append("             field_150300_7,");//--漁會備抵呆帳-催收款項 
                sqlCmd.append("             field_credit,");//--農漁會放款總額
                sqlCmd.append("             round(field_credit/?,0) as field_credit,");//--農漁會放款總額取到億元
                sqlCmd.append("             field_990000_6,");//--農會.狹義逾期放款金額 
                sqlCmd.append("             field_990000_7,");//--漁會.狹義逾期放款金額
                sqlCmd.append("             field_990000,");//--農漁會.狹義逾期放款金額
                sqlCmd.append("             round(field_990000/?,0) as field_990000,");//--農漁會.狹義逾期放款金額.取到億元
                sqlCmd.append("             field_840740_6,");//--農會.廣義逾期放款
                sqlCmd.append("             field_840740_7,");//--漁會.廣義逾期放款
                sqlCmd.append("             field_840740,");//--農漁會.廣義逾期放款
                sqlCmd.append("             round(field_840740/?,0) as field_840740,");//--農漁會.廣義逾期放款.取到億元
                sqlCmd.append("             decode(field_credit,0,0,round(field_990000 /  field_credit *100 ,2)) as field_990000_rate,");//--狹義逾放比率
                sqlCmd.append("             decode(field_credit,0,0,round(field_840740 /  field_credit *100 ,2)) as field_840740_rate ");//--廣義逾放比率
                paramList.add(unit);
                paramList.add(unit);
                paramList.add(unit);
                paramList.add(unit);
                paramList.add(unit);
                paramList.add(unit);
                sqlCmd.append("     from (  ");
                sqlCmd.append("     select  a01.m_year,a01.m_month,");
                sqlCmd.append("             round(sum(decode(bank_type,'6',decode(a01.acc_code,'190000',amt,0))) /1,0)     as field_190000,");
                sqlCmd.append("             round(sum(decode(bank_type,'7',decode(a01.acc_code,'100000',amt,0))) /1,0)     as field_100000,");
                sqlCmd.append("             round(sum(decode(bank_type,'6',decode(a01.acc_code,'190000',amt,0),'7',decode(a01.acc_code,'100000',amt,0))) /1,0)     as field_ASSET,");
                sqlCmd.append("             round(sum(decode(bank_type,'6',decode(a01.acc_code,'310000',amt,0))) /1,0)     as field_310000,");
                sqlCmd.append("             round(sum(decode(bank_type,'6',decode(a01.acc_code,'320000',amt,0))) /1,0)     as field_320000,");
                sqlCmd.append("             round(sum(decode(bank_type,'7',decode(a01.acc_code,'300000',amt,0))) /1,0)     as field_300000,");
                sqlCmd.append("             round(sum(decode(bank_type,'6',decode(a01.acc_code,'310000',amt,'320000',amt,0),'7',decode(a01.acc_code,'300000',amt,0))) /1,0)     as field_NETWorth,");
                sqlCmd.append("             round(sum(decode(bank_type,'6',decode(a01.acc_code,'220000',amt,0))) /1,0)     as field_220000_6,");
                sqlCmd.append("             round(sum(decode(bank_type,'7',decode(a01.acc_code,'220000',amt,0))) /1,0)     as field_220000_7,");
                sqlCmd.append("             round(sum(decode(a01.acc_code,'220000',amt,0)) /1,0) as field_DEBIT,");
                sqlCmd.append("             round(sum(decode(bank_type,'6',decode(a01.acc_code,'120000',amt,0))) /1,0)     as field_120000_6,");
                sqlCmd.append("             round(sum(decode(bank_type,'7',decode(a01.acc_code,'120000',amt,0))) /1,0)     as field_120000_7,");
                sqlCmd.append("             round(sum(decode(bank_type,'6',decode(a01.acc_code,'120800',amt,0))) /1,0)     as field_120800_6,");
                sqlCmd.append("             round(sum(decode(bank_type,'7',decode(a01.acc_code,'120800',amt,0))) /1,0)     as field_120800_7,");
                sqlCmd.append("             round(sum(decode(bank_type,'6',decode(a01.acc_code,'150300',amt,0))) /1,0)     as field_150300_6,");
                sqlCmd.append("             round(sum(decode(bank_type,'7',decode(a01.acc_code,'150300',amt,0))) /1,0)     as field_150300_7,");
                sqlCmd.append("             round(sum(decode(a01.acc_code,'120000',amt,'120800',amt,'150300',amt,0)) /1,0) as  field_CREDIT,");
                sqlCmd.append("             round(sum(decode(bank_type,'6',decode(a01.acc_code,'990000',amt,0))) /1,0)     as field_990000_6,");
                sqlCmd.append("             round(sum(decode(bank_type,'7',decode(a01.acc_code,'990000',amt,0))) /1,0)     as field_990000_7,");
                sqlCmd.append("             round(sum(decode(a01.acc_code,'990000',amt,0)) /1,0) as field_990000");
                sqlCmd.append("     from (select * from a01 ");
                sqlCmd.append("             where a01.m_year= ? ");
                sqlCmd.append("             and a01.m_month=? ");
                paramList.add(s_year);
                paramList.add(s_month);
                sqlCmd.append("             and acc_code in ('190000','100000','310000','320000','300000','220000','120000','120800','150300','990000'))a01");
                sqlCmd.append("     left join (select * from bn01 where m_year=100)bn01 on bn01.bank_no = a01.bank_code");
                sqlCmd.append("     group by a01.m_year,a01.m_month)a01,");
                sqlCmd.append("     ( ");
                sqlCmd.append("      select a04.m_year,a04.m_month,");
                sqlCmd.append("             round(sum(decode(bank_type,'6',decode(a04.acc_code,'840740',amt,'840760',amt,0))) /1,0) as field_840740_6,");
                sqlCmd.append("             round(sum(decode(bank_type,'7',decode(a04.acc_code,'840740',amt,'840760',amt,0))) /1,0) as field_840740_7,");
                sqlCmd.append("             round(sum(decode(a04.acc_code,'840740',amt,'840760',amt,0)) /1,0) as field_840740 ");
                sqlCmd.append("      from (select * from a04 where a04.m_year=? and a04.m_month=? ");
                paramList.add(s_year);
                paramList.add(s_month);
                sqlCmd.append("      and a04.ACC_code in ('840740','840760') ) a04 ");
                sqlCmd.append("      left join  (select * from bn01 where m_year=100)bn01 on bn01.bank_no = a04.bank_code");
                sqlCmd.append("      group by a04.m_year,a04.m_month)a04,");
                sqlCmd.append("      (select * ");
                sqlCmd.append("       from bn01_month where m_year=? and m_month=? )bn01_month ");
                paramList.add(s_year);
                paramList.add(s_month);
                sqlCmd.append("       where a01.m_year = a04.m_year(+) and a01.m_month = a04.m_month(+) ");
                sqlCmd.append("       and a01.m_year = bn01_month.m_year(+) and a01.m_month= bn01_month.m_month(+) ");
                sqlCmd.append("       order by m_year,m_month ");
            }else{
                sqlCmd.append("     order by m_year,m_month ");
            }
            
            List qList = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList
                    ,"m_year,m_month,bank_sum_6,bank_sum_7,field_190000,field_100000,field_asset,field_asset_unit,field_310000,field_320000,field_300000,field_networth,field_networth_unit,field_220000_6,field_220000_7,field_debit,field_debit_unit,field_120000_6,field_120000_7,field_120800_6,field_120800_7,field_150300_6,field_150300_7,field_credit,field_credit_unit,field_990000_6,field_990000_7,field_990000,field_990000_unit,field_840740_6,field_840740_7,field_840740,field_840740_unit,field_990000_rate,field_840740_rate");
            
            String unit_name = Utility.getUnitName(unit);
            String unit_EngName = getUnitEName(unit);
            row=sheet.getRow(3);
            cell=row.getCell((short)5);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellValue(unit_name+"\n"+unit_EngName);
            cell=row.getCell((short)7);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellValue(unit_name+"\n"+unit_EngName);
            row=sheet2.getRow(4);
            cell=row.getCell((short)3);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellValue(unit_name+"\n"+unit_EngName);
            cell=row.getCell((short)4);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellValue(unit_name+"\n"+unit_EngName);
            cell=row.getCell((short)5);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellValue(unit_name+"\n"+unit_EngName);
            insertRow(wb, sheet, 5, qList.size()-2);
            if (qList.size() != 0){
                HSSFCellStyle dateCellStyle=wb.createCellStyle(); 
                short df=wb.createDataFormat().getFormat("mmm-yyyy");  
                dateCellStyle.setDataFormat(df); 
                rowNum1= 4;
                for(j=0;j<qList.size();j++){
                    DataObject bean = (DataObject) qList.get(j);
                    row=(sheet.getRow(rowNum1)==null)? sheet.createRow(rowNum1) : sheet.getRow(rowNum1);
                    //row=sheet.createRow(rowNum1);
                    String m_year=((bean.getValue("m_year") == null)?0:bean.getValue("m_year")).toString();
                    String m_month=((bean.getValue("m_month") == null)?0:bean.getValue("m_month")).toString();
                    String bank_sum_6 = ((bean.getValue("bank_sum_6") == null)?0:bean.getValue("bank_sum_6")).toString();
                    String bank_sum_7 = ((bean.getValue("bank_sum_7") == null)?0:bean.getValue("bank_sum_7")).toString();
                    String field_asset_unit = ((bean.getValue("field_asset_unit") == null)?"":bean.getValue("field_asset_unit")).toString();
                    String field_networth_unit = ((bean.getValue("field_networth_unit") == null)?"":bean.getValue("field_networth_unit")).toString();
                    System.out.println(m_year+", "+m_month+", "+bank_sum_6+", "+bank_sum_7+", "+field_asset_unit+", "+field_networth_unit);
                    if("90".equals(m_year)) continue;
                    //年月別
                    cell=row.getCell((short)0);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if(!(String.valueOf(Integer.parseInt(s_year))).equals(m_year)){
                        cell.setCellValue(m_year+"年底");
                    }else{
                        cell.setCellValue(m_year+"年"+m_month+"月底");
                    }
                    String date  = (Integer.parseInt(m_year)+1911)+"/"+m_month ;
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM");
                    Date dt1 = sdf.parse( date );
                    cell=row.getCell((short)2);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(dt1);
                    
                    cell=row.getCell((short)3);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(bank_sum_6);
                    
                    cell=row.getCell((short)4);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(bank_sum_7);
                    
                    cell=row.getCell((short)5);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    sheet.addMergedRegion(new Region(rowNum1,(short)5,rowNum1,(short)6));
                    cell.setCellValue(field_asset_unit);
                    
                    cell=row.getCell((short)7);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    sheet.addMergedRegion(new Region(rowNum1,(short)7,rowNum1,(short)8));
                    cell.setCellValue(field_networth_unit);
                    
                    rowNum1++;
                }
                copyRows(sheet2, sheet, 0,7, rowNum1, wb);
                rowNum2 = rowNum1+6;
                insertRow(wb, sheet, rowNum2+1, qList.size()-2);
                for(j=0;j<qList.size();j++){
                    DataObject bean = (DataObject) qList.get(j);
                    row=(sheet.getRow(rowNum2)==null)? sheet.createRow(rowNum2) : sheet.getRow(rowNum2);
                    String m_year=((bean.getValue("m_year") == null)?"0":bean.getValue("m_year")).toString();
                    String m_month=((bean.getValue("m_month") == null)?"0":bean.getValue("m_month")).toString();
                    String field_debit_unit = ((bean.getValue("field_debit_unit") == null)?"0":bean.getValue("field_debit_unit")).toString();
                    String field_credit_unit = ((bean.getValue("field_credit_unit") == null)?"0":bean.getValue("field_credit_unit")).toString();
                    String field_990000_unit = ((bean.getValue("field_990000_unit") == null)?"0":bean.getValue("field_990000_unit")).toString();
                    String field_840740_unit = ((bean.getValue("field_840740_unit") == null)?"0":bean.getValue("field_840740_unit")).toString();
                    String field_990000_rate = ((bean.getValue("field_990000_rate") == null)?"0":bean.getValue("field_990000_rate")).toString();
                    String field_840740_rate = ((bean.getValue("field_840740_rate") == null)?"0":bean.getValue("field_840740_rate")).toString();
                    System.out.println(field_debit_unit+", "+field_credit_unit+", "+field_990000_unit+", "+field_840740_unit+", "+field_990000_rate+", "+field_840740_rate);
                    //年月別
                    cell=row.getCell((short)0);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    if(!(String.valueOf(Integer.parseInt(s_year))).equals(m_year)){
                        cell.setCellValue(m_year+"年底");
                    }else{
                        cell.setCellValue(m_year+"年"+m_month+"月底");
                    }
                    
                    String date  = (Integer.parseInt(m_year)+1911)+"/"+m_month ;
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM");
                    Date dt1 = sdf.parse( date );
                    cell=row.getCell((short)2);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(dt1);
                    
                    cell=row.getCell((short)3);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue((field_debit_unit.equals("0")?"":field_debit_unit));
                    
                    cell=row.getCell((short)4);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue((field_credit_unit.equals("0")?"":field_credit_unit));
                    
                    cell=row.getCell((short)5);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue((field_990000_unit.equals("0")?"":field_990000_unit));
                    
                    cell=row.getCell((short)6);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue((field_840740_unit.equals("0")?"":field_840740_unit));
                  
                    cell=row.getCell((short)7);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue((field_990000_rate.equals("0")?"":getPercentNumber_3(field_990000_rate)));
                    
                    cell=row.getCell((short)8);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue((field_840740_rate.equals("0")?"":getPercentNumber_3(field_840740_rate)));
                    
                    rowNum2++;
                }
                
           }
         
            if (qList.size() != 0){
                for(j=0;j<qList.size();j++){
                    DataObject bean = (DataObject) qList.get(j);
                    String field_asset_unit = ((bean.getValue("field_asset_unit") == null)?"":bean.getValue("field_asset_unit")).toString();
                    String field_networth_unit = ((bean.getValue("field_networth_unit") == null)?"":bean.getValue("field_networth_unit")).toString();
                    String field_debit_unit = ((bean.getValue("field_debit_unit") == null)?"":bean.getValue("field_debit_unit")).toString();
                    String field_credit_unit = ((bean.getValue("field_credit_unit") == null)?"":bean.getValue("field_credit_unit")).toString();
                    String field_990000_unit = ((bean.getValue("field_990000_unit") == null)?"":bean.getValue("field_990000_unit")).toString();
                    String field_840740_unit = ((bean.getValue("field_840740_unit") == null)?"":bean.getValue("field_840740_unit")).toString();
                    String field_990000_rate = ((bean.getValue("field_990000_rate") == null)?"":bean.getValue("field_990000_rate")).toString();
                    String field_840740_rate = ((bean.getValue("field_840740_rate") == null)?"":bean.getValue("field_840740_rate")).toString();
                    String field_190000=((bean.getValue("field_190000") == null)?"0":bean.getValue("field_190000")).toString();
                    String field_100000=((bean.getValue("field_100000") == null)?"0":bean.getValue("field_100000")).toString();
                    String field_asset = ((bean.getValue("field_asset") == null)?"0":bean.getValue("field_asset")).toString();
                    String field_310000=((bean.getValue("field_310000") == null)?"0":bean.getValue("field_310000")).toString();
                    String field_320000=((bean.getValue("field_320000") == null)?"0":bean.getValue("field_320000")).toString();
                    String field_300000= ((bean.getValue("field_300000") == null)?"0":bean.getValue("field_300000")).toString();
                    String field_networth = ((bean.getValue("field_networth") == null)?"0":bean.getValue("field_networth")).toString();
                    String field_220000_6=((bean.getValue("field_220000_6") == null)?"0":bean.getValue("field_220000_6")).toString();
                    String field_220000_7=((bean.getValue("field_220000_7") == null)?"0":bean.getValue("field_220000_7")).toString();
                    String field_debit= ((bean.getValue("field_debit") == null)?"0":bean.getValue("field_debit")).toString();
                    String field_120000_6=((bean.getValue("field_120000_6") == null)?"0":bean.getValue("field_120000_6")).toString();
                    String field_120000_7=((bean.getValue("field_120000_7") == null)?"0":bean.getValue("field_120000_7")).toString();
                    String field_120800_6=((bean.getValue("field_120800_6") == null)?"0":bean.getValue("field_120800_6")).toString();
                    String field_120800_7=((bean.getValue("field_120800_7") == null)?"0":bean.getValue("field_120800_7")).toString();
                    String field_150300_6=((bean.getValue("field_150300_6") == null)?"0":bean.getValue("field_150300_6")).toString();
                    String field_150300_7=((bean.getValue("field_150300_7") == null)?"0":bean.getValue("field_150300_7")).toString();
                    String field_credit = ((bean.getValue("field_credit") == null)?"0":bean.getValue("field_credit")).toString();
                    String field_990000_6=((bean.getValue("field_990000_6") == null)?"0":bean.getValue("field_990000_6")).toString();
                    String field_990000_7=((bean.getValue("field_990000_7") == null)?"0":bean.getValue("field_990000_7")).toString();
                    String field_990000=((bean.getValue("field_990000") == null)?"0":bean.getValue("field_990000")).toString();
                    String field_840740_6=((bean.getValue("field_840740_6") == null)?"0":bean.getValue("field_840740_6")).toString();
                    String field_840740_7=((bean.getValue("field_840740_7") == null)?"0":bean.getValue("field_840740_7")).toString();
                    String field_840740=((bean.getValue("field_840740") == null)?"0":bean.getValue("field_840740")).toString();
                    
                    
                    row=sheet1.getRow(5);
                    cell=row.getCell((short)7);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue("取至"+getUnitCName(unit));
                    row=sheet1.getRow(8);
                    cell=row.getCell((short)7);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue("取至"+getUnitCName(unit));
                    row=sheet1.getRow(11);
                    cell=row.getCell((short)7);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue("取至"+getUnitCName(unit));
                    row=sheet1.getRow(14);
                    cell=row.getCell((short)7);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue("取至"+getUnitCName(unit));
                    row=sheet1.getRow(17);
                    cell=row.getCell((short)7);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue("取至"+getUnitCName(unit));
                    row=sheet1.getRow(20);
                    cell=row.getCell((short)7);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue("取至"+getUnitCName(unit));
                    
                    row=sheet1.getRow(1);
                    cell=row.getCell((short)0);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_asset));
                    cell=row.getCell((short)1);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_networth));
                    cell=row.getCell((short)2);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_debit));
                    cell=row.getCell((short)3);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_credit));
                    cell=row.getCell((short)4);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_990000));
                    cell=row.getCell((short)5);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_840740));
                    cell=row.getCell((short)6);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_990000_rate));
                    cell=row.getCell((short)7);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_840740_rate));
                    
                    row=sheet1.getRow(2);
                    cell=row.getCell((short)0);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_asset_unit));
                    cell=row.getCell((short)1);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_networth_unit));
                    cell=row.getCell((short)2);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_debit_unit));
                    cell=row.getCell((short)3);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_credit_unit));
                    cell=row.getCell((short)4);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_990000_unit));
                    cell=row.getCell((short)5);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_840740_unit));
                    
                    row=sheet1.getRow(6);
                    cell=row.getCell((short)0);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_asset));
                    cell=row.getCell((short)1);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_190000));
                    cell=row.getCell((short)2);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_100000));
                    cell=row.getCell((short)7);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_asset_unit));
                    
                    row=sheet1.getRow(9);
                    cell=row.getCell((short)0);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_networth));
                    cell=row.getCell((short)1);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_310000));
                    cell=row.getCell((short)2);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_320000));
                    cell=row.getCell((short)3);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_300000));
                    cell=row.getCell((short)7);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_networth_unit));
                   
                    row=sheet1.getRow(12);
                    cell=row.getCell((short)0);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_debit));
                    cell=row.getCell((short)1);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_220000_6));
                    cell=row.getCell((short)2);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_220000_7));
                    cell=row.getCell((short)7);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_debit_unit));
                    
                    row=sheet1.getRow(15);
                    cell=row.getCell((short)0);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_credit));
                    cell=row.getCell((short)1);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_120000_6));
                    cell=row.getCell((short)2);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_120800_6));
                    cell=row.getCell((short)3);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_150300_6));
                    cell=row.getCell((short)4);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_120000_7));
                    cell=row.getCell((short)5);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_120800_7));
                    cell=row.getCell((short)6);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_150300_7));
                    cell=row.getCell((short)7);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_credit_unit));
                    
                    row=sheet1.getRow(18);
                    cell=row.getCell((short)0);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_990000));
                    cell=row.getCell((short)1);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_990000_6));
                    cell=row.getCell((short)2);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_990000_7));
                    cell=row.getCell((short)7);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_990000_unit));
                    
                    row=sheet1.getRow(21);
                    cell=row.getCell((short)0);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_840740));
                    cell=row.getCell((short)1);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_840740_6));
                    cell=row.getCell((short)2);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_840740_7));
                    cell=row.getCell((short)7);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_840740_unit));
                    
                    row=sheet1.getRow(24);
                    cell=row.getCell((short)0);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(getPercentNumber_3(field_990000_rate));
                    //System.out.println("field_990000_rate="+getPercentNumber_3(field_990000_rate));
                    cell=row.getCell((short)1);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_990000));
                    cell=row.getCell((short)2);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_credit));
                    
                    row=sheet1.getRow(27);
                    cell=row.getCell((short)0);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(getPercentNumber_3(field_840740_rate));
                    //System.out.println("field_840740_rate="+getPercentNumber_3(field_840740_rate));
                    cell=row.getCell((short)1);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_840740));
                    cell=row.getCell((short)2);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellValue(Double.parseDouble(field_credit));
                    
                }
                
                
            }
            
            
			
            // 建表結束--------------------------------------
	        HSSFFooter footer = sheet.getFooter();
	        footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));

	        FileOutputStream fout=new FileOutputStream(reportDir + System.getProperty("file.separator")+filename);
	        wb.write(fout);
	        //儲存
	        fout.close();
	        System.out.println("儲存完成");
		}catch(Exception e){
			System.out.println("FR001WD.createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}
    
    public static String getMonthName(int month){
        String unit_name="";
        try{
            if (month==1){
                unit_name="Jan-";
            }else if (month==2){                 
                unit_name="Feb-";
            }else if (month==3){
                unit_name="Mar-";
            }else if (month==4){
                unit_name="Apr-";
            }else if (month==5){
                unit_name="May-";
            }else if (month==6){
                unit_name="Jun-";
            }else if (month==7){                 
                unit_name="Jul-";
            }else if (month==8){
                unit_name="Aug-";
            }else if (month==9){
                unit_name="Sep-";
            }else if (month==10){
                unit_name="Oct-";
            }else if (month==11){
                unit_name="Nov-";
            }else if (month==12){
                unit_name="Dec-";
            }
        }catch(Exception e){
            System.out.println("getMonthName Error:"+e.getMessage());
        }
        return unit_name;   
    }
    public static String getUnitEName(String unit){
        String unit_name="";
        try{
            if (unit.equals("1")){
                unit_name="dollars";
            }else if (unit.equals("1000")){                 
                unit_name="A thousand";
            }else if (unit.equals("10000")){
                unit_name="A ten thousand";
            }else if (unit.equals("1000000")){
                unit_name="A million";
            }else if (unit.equals("10000000")){
                unit_name="A ten million";
            }else if (unit.equals("100000000")){
                unit_name="A hundred million";
            }
        }catch(Exception e){
            System.out.println("getUnitName Error:"+e.getMessage());
        }
        return unit_name;   
    }
    public static String getUnitCName(String unit){
        String unit_name="";
        try{
            if (unit.equals("1")){
                unit_name="元";
            }else if (unit.equals("1000")){                 
                unit_name="千元";
            }else if (unit.equals("10000")){
                unit_name="萬元";
            }else if (unit.equals("1000000")){
                unit_name="百萬元";
            }else if (unit.equals("10000000")){
                unit_name="千萬元";
            }else if (unit.equals("100000000")){
                unit_name="億元";
            }
        }catch(Exception e){
            System.out.println("getUnitName Error:"+e.getMessage());
        }
        return unit_name;   
    }
    
    public static void copyRows(HSSFSheet pSourceSheet, HSSFSheet pTargetSheet, int pStartRow,
            int pEndRow, int pPosition, HSSFWorkbook wb) {
           HSSFRow sourceRow = null;
           HSSFRow targetRow = null;
           HSSFCell sourceCell = null;
           HSSFCell targetCell = null;
           HSSFSheet sourceSheet = null;
           HSSFSheet targetSheet = null;
           Region region = null;
           int cType;
           int i;
           short j;
           int targetRowFrom;
           int targetRowTo;
           if ((pStartRow == -1) || (pEndRow == -1)) {
            return;
           }
           
           sourceSheet = pSourceSheet;
           targetSheet = pTargetSheet;
           //copy合併的單元格
           for (i = 0; i < sourceSheet.getNumMergedRegions(); i++) {
            region = sourceSheet.getMergedRegionAt(i);
            if ((region.getRowFrom() >= pStartRow)
              && (region.getRowTo() <= pEndRow)) {
             targetRowFrom = region.getRowFrom() - pStartRow + pPosition;
             targetRowTo = region.getRowTo() - pStartRow + pPosition;
             region.setRowFrom(targetRowFrom);
             region.setRowTo(targetRowTo);
             targetSheet.addMergedRegion(region);
            }
           }
           //設置列寬
           //如果是同一頁就不需要設置列寬，否則會有問題
           if (pSourceSheet != pTargetSheet) {
            for (i = pStartRow; i <= pEndRow; i++) {
             sourceRow = sourceSheet.getRow(i);
             if (sourceRow != null) {
              for (j = sourceRow.getFirstCellNum(); j < sourceRow
                .getLastCellNum(); j++) {
               targetSheet.setColumnWidth(j, sourceSheet
                 .getColumnWidth(j));
              }
              break;
             }
            }
           }
           //拷貝並填充數據
           for (i = pStartRow;i <= pEndRow; i++) {
            sourceRow = sourceSheet.getRow(i);
            if (sourceRow == null) {
             continue;
            }
            targetRow = targetSheet.createRow(i - pStartRow + pPosition);
            targetRow.setHeight(sourceRow.getHeight());
            for (j = sourceRow.getFirstCellNum(); j < sourceRow
              .getLastCellNum(); j++) {
             sourceCell = sourceRow.getCell(j);
             if (sourceCell == null) {
              continue;
             }
             targetCell = targetRow.createCell(j);
             targetCell.setEncoding(sourceCell.getEncoding());
             targetCell.setCellStyle(sourceCell.getCellStyle());
             cType = sourceCell.getCellType();
             targetCell.setCellType(cType);
             switch (cType) {
             case HSSFCell.CELL_TYPE_BOOLEAN:
              targetCell.setCellValue(sourceCell.getBooleanCellValue());
              break;
             case HSSFCell.CELL_TYPE_ERROR:
              targetCell
                .setCellErrorValue(sourceCell.getErrorCellValue());
              break;
             case HSSFCell.CELL_TYPE_FORMULA:
              // parseFormula
              targetCell.setCellFormula(parseFormula(sourceCell
                .getCellFormula()));
              break;
             case HSSFCell.CELL_TYPE_NUMERIC:
              targetCell.setCellValue(sourceCell.getNumericCellValue());
              break;
             case HSSFCell.CELL_TYPE_STRING:
              targetCell.setCellValue(sourceCell.getStringCellValue());
              break;
             }
            }
           }
          }
          //這是為了解決單元格中設置了函數的複製問題
     private static String parseFormula(String pPOIFormula) {
        final String cstReplaceString = "ATTR(semiVolatile)"; //$NON-NLS-1$
          StringBuffer result = null;
           int index;
           result = new StringBuffer();
           index = pPOIFormula.indexOf(cstReplaceString);
           if (index >= 0) {
            result.append(pPOIFormula.substring(0, index));
            result.append(pPOIFormula.substring(index
              + cstReplaceString.length()));
           } else {
            result.append(pPOIFormula);
           }
           return result.toString();
          }
     
    public static void insertRow(HSSFWorkbook wb, HSSFSheet sheet,  int starRow, int rows) { 
           sheet.shiftRows(starRow + 1, sheet.getLastRowNum(), rows); 
           starRow = starRow - 1;
           for (int i = 0; i < rows; i++) { 
                HSSFRow sourceRow = null;
                HSSFRow targetRow = null;
                HSSFCell sourceCell = null;
                HSSFCell targetCell = null;
                short m;
                starRow = starRow + 1;
                sourceRow = sheet.getRow(starRow);
                targetRow = sheet.createRow(starRow + 1);
                targetRow.setHeight(sourceRow.getHeight()); 
                for (m = sourceRow.getFirstCellNum(); m < sourceRow.getPhysicalNumberOfCells(); m++) {
                    sourceCell = sourceRow.getCell(m);
                    targetCell = targetRow.createCell(m);
                    targetCell.setEncoding(sourceCell.getEncoding());
                    targetCell.setCellStyle(sourceCell.getCellStyle());
                    targetCell.setCellType(sourceCell.getCellType());
                    
                }
           }       
    }
    
    public static String getPercentNumber_3(String szNumString) {
        try {
            double dDouble=Double.parseDouble(szNumString);
            DecimalFormat df_md = new DecimalFormat("0.00");
            return df_md.format(dDouble).toString();            
        }catch (Exception e) {
            System.out.println(e);return szNumString;
        }
    }
}
