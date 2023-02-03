/*
106.06.02 create 農漁會信用部重要數據 by George
106.10.25 fix 調整(續2)3.全體農漁會信用部○年至○年○月經營情形表(查詢年月及前6年底),報表月份及title顯示 by 2295
106.12.07 fix 調整(續1)/(續2)逾放比率與率本適足率位置錯置;(續1)增減比較%調整為個百分點  by 2295
106.12.19 fix 調整原讀取a01_operation改為a01_operation_month(關帳資料) by 2295
106.12.20 fix 暫調整(續2)前5年底為固定數字.只顯示當年底資料 by 2295
106.12.26 fix 調整(續1).存款-一般放款/逾放金額-一般放款先除以單位.再相減  by 2295
106.12.26 fix 調整(續3).4.合計的部份.以原本各機構.先除以單位後.再加總  by 2295
106.12.26 fix 調整(續2)3.存款年成長率/放款年成長率,當月底與前一年底金額比較 by 2295
106.12.27 fix 調整(續2)3.前5年底若為100~105年,則讀取固定資料顯示 by 2295
106.12.29 fix 調整(續5).原農漁會信用.建築放款餘額A99.992710調整為讀取A01_operation_month關帳資料 by 2295
*/
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.*;
import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR077W {

	public static String createRpt(String s_year, String s_month, String bank_type, String unit)	{

		String errMsg = "";
		String m_year = (Integer.parseInt(s_year) < 100)?"99":"100"; 
		String unitName = Utility.getUnitName(unit);
		
		reportUtil reportUtil = new reportUtil();
		StringBuffer sql = new StringBuffer () ;
		List dbData = new ArrayList();
		ArrayList paramList = new ArrayList() ;
		int num = 0; 
		
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

			String openfile = "重要數據.xls";	
			System.out.println("open file " + openfile);
			FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+ openfile );			

			//設定FileINputStream讀取Excel檔
			POIFSFileSystem fs = new POIFSFileSystem( finput );
			if (fs == null) {
				System.out.println("Open 範本檔失敗");
			} else {
				System.out.println("Open 範本檔成功");
			}
			HSSFWorkbook wb = new HSSFWorkbook(fs);
			if (wb == null) {
				System.out.println("Open 工作表失敗");
			} else {
				System.out.println("Open 工作表成功");
			}
			
			finput.close();
			
			HSSFPrintSetup ps = null;
			HSSFSheet sheet = null;
			HSSFRow row = null;  //宣告一列
			HSSFCell cell = null;//宣告一個儲存格
			
			HSSFFont ft = wb.createFont();
			HSSFFont ft1 = wb.createFont();			
			HSSFCellStyle cs = wb.createCellStyle();
			HSSFCellStyle cs_center = reportUtil.getDefaultStyle(wb);
			HSSFCellStyle cs1 = null;
			HSSFCellStyle cs2 = null;
            
			ft.setFontHeightInPoints((short)12);
			ft.setFontName("標楷體");
			ft1.setFontHeightInPoints((short)12);
			ft1.setFontName("細明體");
			cs.setFont(ft1);
			cs.setBorderTop((short)0);
			cs.setBorderLeft((short)0);
			cs.setBorderLeft((short)0);
			cs.setBorderBottom((short)0);
			
			String printTime = Utility.getDateFormat("  HH:mm:ss");
            String printDate = Utility.getDateFormat("yyyy/MM/dd");
            
			//開始寫入Excel的內容
			int rowNum = 0;
			int cellNum = 0;
			String sValue = "";
			DataObject bean = null;
			
			//續1工作表:A4直印,縮放比100%
			sheet = wb.getSheetAt(0); //讀取第一個工作表，宣告其為sheet
			if (sheet == null) {
				System.out.println("open sheet1 失敗");
			} else {
				System.out.println("open sheet1 成功");
			}
			ps = sheet.getPrintSetup(); //取得設定

			// 設定頁面符合列印大小
			sheet.setAutobreaks(false);
			ps.setScale((short)100); //列印縮放百分比
			ps.setPaperSize((short)9); //設定紙張大小 A4
			ps.setLandscape(false);//設定橫印
			
			sql.setLength(0);
			paramList.clear();
			
			//資料SQL
			sql.append(" select tbank_cnt_6, "); //農會本部家數
			sql.append("       tbank_cnt_7, "); //漁會本部家數
			sql.append("       tbank_cnt_6+tbank_cnt_7 as tbank_cnt, "); //本部家數.合計
			sql.append("       bank_cnt_6, "); //農會分部家數
			sql.append("       bank_cnt_7, "); //漁會分部家數
			sql.append("       bank_cnt_6+bank_cnt_7  as bank_cnt, "); //分部家數.合計
			sql.append("       tbank_cnt_6+bank_cnt_6 as bank_6_sum, "); //農會本部+分部家數.合計
			sql.append("       tbank_cnt_7+bank_cnt_7 as bank_7_sum, "); //漁會本部+分部家數.合計
			sql.append("       tbank_cnt_6+bank_cnt_6+ tbank_cnt_7+bank_cnt_7 as bank_sum "); //本部+分部家數.合計
			sql.append(" from rpt_business ");
			sql.append(" where m_year=? and m_month=? ");
			paramList.add(s_year);
			paramList.add(s_month);
			
			dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "tbank_cnt_6,tbank_cnt_7,tbank_cnt,bank_cnt_6,bank_cnt_7,bank_cnt,bank_6_sum,bank_7_sum,bank_sum");
			System.out.println("1.dbData.size()="+dbData.size());
			
			row = sheet.getRow(0);
			cell = row.getCell((short) 0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(s_year+"年"+s_month+"月農漁會信用部重要數據");
			
			row = sheet.getRow(1);
			cell = row.getCell((short) 2);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("列印日期：" + Utility.getCHTdate(printDate, 1)+printTime);
			
			row = sheet.getRow(4);
			cell = row.getCell((short) 0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("1.全體農漁會信用部"+s_year+"年"+s_month+"月家數" );
			
			// 列印明細/合計
			for (int datacount = 0; datacount < dbData.size(); datacount++) {
				bean = (DataObject) dbData.get(datacount);
				
				for (int counts = 0; counts < 9; counts++) {
					
					rowNum = (counts%3)+6;
					cellNum = (counts/3)+1;
					
					if (counts == 0)
						sValue = (bean.getValue("tbank_cnt_6") == null) ? "0" : String.valueOf(bean.getValue("tbank_cnt_6"));//農會本部家數
					else if (counts == 1)
						sValue = (bean.getValue("bank_cnt_6") == null) ? "0" : String.valueOf(bean.getValue("bank_cnt_6"));//農會分部家數
					else if (counts == 2)
						sValue = (bean.getValue("bank_6_sum") == null) ? "0" : String.valueOf(bean.getValue("bank_6_sum"));//農會本部+分部家數.合計
					else if (counts == 3)
						sValue = (bean.getValue("tbank_cnt_7") == null) ? "0" : String.valueOf(bean.getValue("tbank_cnt_7"));//漁會本部家數
					else if (counts == 4)
						sValue = (bean.getValue("bank_cnt_7") == null) ? "0" : String.valueOf(bean.getValue("bank_cnt_7"));//漁會分部家數
					else if (counts == 5)
						sValue = (bean.getValue("bank_7_sum") == null) ? "0" : String.valueOf(bean.getValue("bank_7_sum"));//漁會本部+分部家數.合計
					else if (counts == 6)
						sValue = (bean.getValue("tbank_cnt") == null) ? "0" : String.valueOf(bean.getValue("tbank_cnt"));//本部家數.合計
					else if (counts == 7)
						sValue = (bean.getValue("bank_cnt") == null) ? "0" : String.valueOf(bean.getValue("bank_cnt"));//分部家數.合計
					else if (counts == 8)
						sValue = (bean.getValue("bank_sum") == null) ? "0" : String.valueOf(bean.getValue("bank_sum"));//本部+分部家數.合計
					
					row = sheet.getRow(rowNum);
					cell = row.getCell((short) cellNum);
					
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs_center);
					cell.setCellValue(Utility.setCommaFormat(sValue));
				}
			}
			
			row = sheet.getRow(10);
			cell = row.getCell((short) 0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("2.全體農漁會信用部93年1月與"+s_year+"年"+s_month+"月經營情形比較表" );
			
			row = sheet.getRow(11);
			cell = row.getCell((short) 3);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("單位:"+unitName+"、百分點" );
			
			row = sheet.getRow(12);
			cell = row.getCell((short) 2);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(s_year+"年"+s_month+"月" );
			
			sql.setLength(0);
			paramList.clear();
			
			//資料SQL
			sql.append(" select round(field_ASSETS/? ,0) as field_assets, "); //資產
			paramList.add(unit);
			sql.append("       round(field_DEBT/? ,0) as field_debt,  "); //負債
			paramList.add(unit);
			sql.append("       round(field_DEBIT/? ,0) as field_debit,  "); //存款
			paramList.add(unit);
			sql.append("       round(field_CREDIT/? ,0) as field_credit, "); //放款
			paramList.add(unit);
			sql.append("       round(field_CREDIT/? ,0) - round(loan_amt/? ,0) as field_credit_normal, "); //放款-一般放款
			paramList.add(unit);
			paramList.add(unit);
			sql.append("       round(loan_amt/? ,0) as loan_amt, "); //放款-專案農貸
			paramList.add(unit);
			sql.append("       round(field_OVER/? ,0) as field_over, "); //逾放金額
			paramList.add(unit);
			sql.append("       round(field_OVER/? ,0) - round(loan_over_amt/? ,0) as field_over_normal, "); //逾放金額 -一般放款
			paramList.add(unit);
			paramList.add(unit);
			sql.append("       round(loan_over_amt/? ,0) as loan_over_amt, "); //逾放金額-專案農貸
			paramList.add(unit);
			sql.append("       field_over_rate, "); //逾放比率
			sql.append("       decode((field_CREDIT-loan_amt),0,0,round(((field_OVER-loan_over_amt)/(field_CREDIT-loan_amt))*100 ,2)) as field_over_rate_normal, "); //逾放比率-一般放款
			sql.append("       decode(loan_amt,0,0,round((loan_over_amt/loan_amt)*100 ,2)) as field_over_rate_loan, "); //逾放比率-專案農貸
			sql.append("       round(field_NET/? ,0)  as field_net , "); //淨值
			paramList.add(unit);
			sql.append("       round(field_320300 /? ,0)  as field_320300, "); //稅前純益
			paramList.add(unit);
			sql.append("       round(field_BACKUP/? ,0) as field_backup, "); //備抵呆帳
			paramList.add(unit);
			sql.append("       round(field_120700 /? ,0) as field_120700, "); //內部融資
			paramList.add(unit);
			sql.append("       field_assets_rate, "); //資產報酬率
			sql.append("       field_net_rate, "); //淨值報酬率
			sql.append("       field_backup_credit_rate, "); //備呆占放款比率=備抵呆帳/放款
			sql.append("       field_backup_over_rate,  "); //備呆占逾放比率=備抵呆帳/逾期放款
			sql.append("       field_captial_rate, "); //資本適足率
			sql.append("       bank_cnt, "); //受輔導信用部家數
			sql.append("       over_rate_count, "); //逾放比率 15%以上家數
			sql.append("       bis_count, "); //資本適足率8%以下家數
			sql.append("       loan_cnt  "); //專案農貸受益戶數
			sql.append(" from ( ");
			sql.append(" select round(sum(decode(a01.acc_code,'field_assets',a01.amt,0)) /1,0) as field_assets, "); //資產
			sql.append("   round(sum(decode(a01.acc_code,'field_debt',a01.amt,0)) /1,0) as field_debt,   "); //負債
			sql.append("    round(sum(decode(a01.acc_code,'field_debit',a01.amt,0)) /1,0) as field_debit, "); //存款
			sql.append("    round(sum(decode(a01.acc_code,'field_credit',a01.amt,0)) /1,0) as field_credit,   "); //放款
			sql.append("    round(sum(decode(a01.acc_code,'field_over',a01.amt,0)) /1,0) as field_over, "); //逾放金額
			sql.append("    sum(decode(a01.acc_code,'field_over_rate',a01.amt,0)) as field_over_rate, "); //逾放比率
			sql.append("    round(sum(decode(a01.acc_code,'field_net',a01.amt,0)) /1,0) as field_net , "); //淨值
			sql.append("   round(sum(decode(a01.acc_code,'field_320300',a01.amt,0)) /1,0)  as field_320300, "); //稅前純益
			sql.append("   round(sum(decode(a01.acc_code,'field_backup',a01.amt,0)) /1,0) as field_backup, "); //備抵呆帳
			sql.append("   round(sum(decode(a01.acc_code,'field_120700',a01.amt,0)) /1,0) as field_120700, "); //內部融資
			sql.append("        decode(sum(decode(a01.acc_code,'field_assets',a01.amt,0)),0,0,round((sum(decode(a01.acc_code,'field_320300',a01.amt,0))/sum(decode(a01.acc_code,'field_assets',a01.amt,0)))*100 ,2)) as field_assets_rate, "); //資產報酬率
			sql.append("        decode(sum(decode(a01.acc_code,'field_net',a01.amt,0)),0,0,round((sum(decode(a01.acc_code,'field_320300',a01.amt,0))/sum(decode(a01.acc_code,'field_net',a01.amt,0)))*100 ,2)) as field_net_rate, "); //淨值報酬率
			sql.append("        sum(decode(a01.acc_code,'field_backup_credit_rate',a01.amt,0))  as field_backup_credit_rate, "); //備呆占放款比率=備抵呆帳/放款
			sql.append("        sum(decode(a01.acc_code,'field_backup_over_rate',a01.amt,0)) as field_backup_over_rate,  "); //備呆占逾放比率=備抵呆帳/逾期放款
			sql.append("        sum(decode(a01.acc_code,'field_captial_rate',a01.amt,0)) as field_captial_rate  "); //資本適足率
			sql.append(" from (select * from a01_operation_month where m_year=? and m_month=? and bank_type in ('ALL') and hsien_id=' ')a01 ");
			paramList.add(s_year);
			paramList.add(s_month);
			sql.append(" )a01, (select * from rpt_business where m_year=? and m_month=? )rpt_business, ");
			paramList.add(s_year);
			paramList.add(s_month);
			sql.append(" ( select count(*) as bis_count ");
			sql.append(" from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
			paramList.add(m_year);
			sql.append(" left join (select * from a01_operation_month where m_year=? and m_month=? and bank_type in ('6','7') and bank_code !='ALL')a01 on  bn01.bank_no = a01.bank_code ");
			paramList.add(s_year);
			paramList.add(s_month);
			sql.append(" where acc_code='field_captial_rate' ");
			sql.append(" and amt < 8  )a01_bis,  "); //資本適足率8%以下家數
			sql.append(" 	(  select count(*) as over_rate_count ");
			sql.append(" 	from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
			paramList.add(m_year);
			sql.append(" 	left join (select * from a01_operation_month where m_year=? and m_month=? and bank_type in ('6','7') and bank_code !='ALL')a01 on  bn01.bank_no = a01.bank_code ");
			paramList.add(s_year);
			paramList.add(s_month);
			sql.append(" 	where acc_code='field_over_rate' and amt > 15 ");
			sql.append(" 	)a01_over_rate  "); //逾放比率15%以上家數

			dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "field_assets,field_debt,field_debit,field_credit,field_credit_normal,loan_amt,field_over,field_over_normal,loan_over_amt,field_over_rate,field_over_rate_normal,field_over_rate_loan,field_net,field_320300,field_backup,field_120700,field_assets_rate,field_net_rate,field_backup_credit_rate,field_backup_over_rate,field_captial_rate,bank_cnt,bis_count,over_rate_count,loan_cnt");
			System.out.println("2.dbData.size()="+dbData.size());
			
			// 列印明細/合計
			for (int datacount = 0; datacount < dbData.size(); datacount++) {
				bean = (DataObject) dbData.get(datacount);
				
				for (int rowcount = 13 ; rowcount < 38; rowcount++) {
					String sValue2 = "";
					row = sheet.getRow(rowcount);
					
					if (rowcount == 13){
						sValue = (bean.getValue("field_assets") == null) ? "0" : String.valueOf(bean.getValue("field_assets"));//資產
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = Utility.setCommaFormat(sValue);
						sValue2 = Utility.setCommaFormat(sValue2);
					}else if (rowcount == 14){
						sValue = (bean.getValue("field_debt") == null) ? "0" : String.valueOf(bean.getValue("field_debt"));//負債
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = Utility.setCommaFormat(sValue);
						sValue2 = Utility.setCommaFormat(sValue2);
					}else if (rowcount == 15){
						sValue = (bean.getValue("field_debit") == null) ? "0" : String.valueOf(bean.getValue("field_debit"));//存款
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = Utility.setCommaFormat(sValue);
						sValue2 = Utility.setCommaFormat(sValue2);
					}else if (rowcount == 16){
						sValue = (bean.getValue("field_credit") == null) ? "0" : String.valueOf(bean.getValue("field_credit"));//放款
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = Utility.setCommaFormat(sValue);
						sValue2 = Utility.setCommaFormat(sValue2);
					}else if (rowcount == 17){
						sValue = (bean.getValue("field_credit_normal") == null) ? "0" : String.valueOf(bean.getValue("field_credit_normal"));//放款-一般放款
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = Utility.setCommaFormat(sValue);
						sValue2 = Utility.setCommaFormat(sValue2);
					}else if (rowcount == 18){
						sValue = (bean.getValue("loan_amt") == null) ? "0" : String.valueOf(bean.getValue("loan_amt"));//放款-專案農貸
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = Utility.setCommaFormat(sValue);
						sValue2 = Utility.setCommaFormat(sValue2);
					}else if (rowcount == 19){
						sValue = (bean.getValue("field_over") == null) ? "0" : String.valueOf(bean.getValue("field_over"));//逾放金額
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = Utility.setCommaFormat(sValue);
						sValue2 = Utility.setCommaFormat(sValue2);
					}else if (rowcount == 20){
						sValue = (bean.getValue("field_over_normal") == null) ? "0" : String.valueOf(bean.getValue("field_over_normal"));//逾放金額 -一般放款
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = Utility.setCommaFormat(sValue);
						sValue2 = Utility.setCommaFormat(sValue2);
					}else if (rowcount == 21){
						sValue = (bean.getValue("loan_over_amt") == null) ? "0" : String.valueOf(bean.getValue("loan_over_amt"));//逾放金額-專案農貸
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = Utility.setCommaFormat(sValue);
						sValue2 = Utility.setCommaFormat(sValue2);
					}else if (rowcount == 22){
						sValue =  (bean.getValue("field_over_rate") == null) ? "0" : String.valueOf(bean.getValue("field_over_rate"));//逾放比率
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,true);
						sValue = String.format("%s%%",sValue);
						sValue2 = String.format("%s 個百分點",sValue2);
					}else if (rowcount == 23){
						sValue =  (bean.getValue("field_over_rate_normal") == null) ? "0" : String.valueOf(bean.getValue("field_over_rate_normal"));//逾放比率-一般放款
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,true);
						sValue = String.format("%s%%",sValue);
						sValue2 = String.format("%s 個百分點",sValue2);
					}else if (rowcount == 24){
						sValue =  (bean.getValue("field_over_rate_loan") == null) ? "0" : String.valueOf(bean.getValue("field_over_rate_loan"));//逾放比率-專案農貸
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,true);
						sValue = String.format("%s%%",sValue);
						sValue2 = String.format("%s 個百分點",sValue2);
					}else if (rowcount == 25){
						sValue = (bean.getValue("field_net") == null) ? "0" : String.valueOf(bean.getValue("field_net"));//淨值
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = Utility.setCommaFormat(sValue);
						sValue2 = Utility.setCommaFormat(sValue2);
					}else if (rowcount == 26){
						sValue = (bean.getValue("field_320300") == null) ? "0" : String.valueOf(bean.getValue("field_320300"));//稅前純益
						sValue = Utility.setCommaFormat(sValue);
					}else if (rowcount == 27){
						sValue = (bean.getValue("field_backup") == null) ? "0" : String.valueOf(bean.getValue("field_backup"));//備抵呆帳
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = Utility.setCommaFormat(sValue);
						sValue2 = Utility.setCommaFormat(sValue2);
					}else if (rowcount == 28){
						sValue = (bean.getValue("field_120700") == null) ? "0" : String.valueOf(bean.getValue("field_120700"));//內部融資
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = Utility.setCommaFormat(sValue);
						sValue2 = Utility.setCommaFormat(sValue2);
					}else if (rowcount == 29){
						sValue =  (bean.getValue("field_assets_rate") == null) ? "0" : String.valueOf(bean.getValue("field_assets_rate"));//資產報酬率
						sValue = String.format("%s%%",sValue);
					}else if (rowcount == 30){
						sValue =  (bean.getValue("field_net_rate") == null) ? "0" : String.valueOf(bean.getValue("field_net_rate"));//淨值報酬率
						sValue = String.format("%s%%",sValue);
					}else if (rowcount == 31){
						sValue =  (bean.getValue("field_backup_credit_rate") == null) ? "0" : String.valueOf(bean.getValue("field_backup_credit_rate"));//備呆占放款比率=備抵呆帳/放款
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,true);
						sValue = String.format("%s%%",sValue);
						sValue2 = String.format("%s 個百分點",sValue2);
					}else if (rowcount == 32){
						sValue =  (bean.getValue("field_backup_over_rate") == null) ? "0" : String.valueOf(bean.getValue("field_backup_over_rate"));//備呆占逾放比率=備抵呆帳/逾期放款
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,true);
						sValue = String.format("%s%%",sValue);
						sValue2 = String.format("%s 個百分點",sValue2);
					}else if (rowcount == 33){
						sValue = (bean.getValue("field_captial_rate") == null) ? "0" : String.valueOf(bean.getValue("field_captial_rate"));//資本適足率
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = String.format("%s%%",sValue);
						sValue2 = String.format("%s 個百分點",sValue2);
					}else if (rowcount == 34){
						sValue = (bean.getValue("bank_cnt") == null) ? "0" : String.valueOf(bean.getValue("bank_cnt"));//受輔導信用部家數
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = String.format("%s 家",sValue);
						sValue2 = String.format("%s 家",sValue2);
					}else if (rowcount == 35){
						sValue = (bean.getValue("over_rate_count") == null) ? "0" : String.valueOf(bean.getValue("over_rate_count"));//逾放比率 15%以上家數
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = String.format("%s 家",sValue);
						sValue2 = String.format("%s 家",sValue2);
					}else if (rowcount == 36){
						sValue = (bean.getValue("bis_count") == null) ? "0" : String.valueOf(bean.getValue("bis_count"));//資本適足率8%以下家數
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = String.format("%s 家",sValue);
						sValue2 = String.format("%s 家",sValue2);
					}else if (rowcount == 37){
						sValue = (bean.getValue("loan_cnt") == null) ? "0" : String.valueOf(bean.getValue("loan_cnt"));//專案農貸受益戶數
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = Utility.setCommaFormat(sValue);
						sValue2 = Utility.setCommaFormat(sValue2);
					}
					
					cell = row.getCell((short) 2);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellValue(sValue);
					
					if(!"".equals(sValue2)){
						cell = row.getCell((short) 3);
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell.setCellValue(sValue2);
					}
				}
			}
			
			//續2工作表:A4直印,縮放比90%
			sheet = wb.getSheetAt(1);
			if (sheet == null) {
				System.out.println("open sheet2 失敗");
			} else {
				System.out.println("open sheet2 成功");
			}
			ps = sheet.getPrintSetup(); //取得設定

			// 設定頁面符合列印大小
			sheet.setAutobreaks(false);
			ps.setScale((short)90); //列印縮放百分比
			ps.setPaperSize((short)9); //設定紙張大小 A4
			ps.setLandscape(false);//設定橫印
			
			sql.setLength(0);
			paramList.clear();
			
			row = sheet.getRow(0);
			cell = row.getCell((short) 0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(s_year+"年"+s_month+"月農漁會信用部重要數據" );
			
			row = sheet.getRow(2);
			cell = row.getCell((short) 5);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("列印日期：" + Utility.getCHTdate(printDate, 1)+printTime);
			
			row = sheet.getRow(3);
			cell = row.getCell((short) 0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("3.全體農漁會信用部"+(Integer.parseInt(s_year)-6)+"年至"+s_year+"年"+s_month+"月經營情形表" );
			
			row = sheet.getRow(4);
			cell = row.getCell((short) 7);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("單位:"+unitName );
			
			dbData = getA01Operation(unit,s_year,s_month);
			System.out.println("3.dbData.size()="+dbData.size());
			
						
			for(int datacount = 0 ; datacount < dbData.size() ; datacount++){//106.12.20 原前5年及當年度月份
			//for(int datacount = 6 ; datacount < dbData.size() ; datacount++){//106.12.20 暫調整只顯示當年度月份			
				bean = (DataObject) dbData.get(datacount);
				
				cellNum = datacount + 1;
				
				for (int rowcount = 5; rowcount < 26; rowcount++) {
					sValue = "";
					if (rowcount == 5)
						sValue = String.valueOf((Integer.parseInt(s_year)-6+datacount)+"年"+(datacount==6?s_month+"月":"")+"底");
					else if (rowcount == 6)
						sValue = Utility.setCommaFormat((bean.getValue("field_assets") == null) ? "" : String.valueOf(bean.getValue("field_assets")));//資產
					else if (rowcount == 7)
						sValue = Utility.setCommaFormat((bean.getValue("field_debt") == null) ? "" : String.valueOf(bean.getValue("field_debt")));//負債
					else if (rowcount == 8)
						sValue = Utility.setCommaFormat((bean.getValue("field_debit") == null) ? "" : String.valueOf(bean.getValue("field_debit")));//存款
					else if (rowcount == 9)
						sValue = Utility.setCommaFormat((bean.getValue("field_credit") == null) ? "" : String.valueOf(bean.getValue("field_credit")));//放款
					else if (rowcount == 10)
						sValue = (bean.getValue("field_dc_rate") == null) ? "" : String.format("%s%%",String.valueOf(bean.getValue("field_dc_rate")));//存放比率
					else if (rowcount == 11)
						sValue = Utility.setCommaFormat((bean.getValue("field_net") == null) ? "" : String.valueOf(bean.getValue("field_net")));//淨值
					else if (rowcount == 12)
						sValue = Utility.setCommaFormat((bean.getValue("field_320300") == null) ? "" : String.valueOf(bean.getValue("field_320300")));//稅前純益
					else if (rowcount == 13)
						sValue = Utility.setCommaFormat((bean.getValue("field_over") == null) ? "" : String.valueOf(bean.getValue("field_over")));//逾放金額
					else if (rowcount == 14)
						sValue = (bean.getValue("field_over_rate") == null) ? "" : String.format("%s%%",String.valueOf(bean.getValue("field_over_rate")));//逾放比率
					else if (rowcount == 15)
						sValue = Utility.setCommaFormat((bean.getValue("field_backup") == null) ? "" : String.valueOf(bean.getValue("field_backup")));//備抵呆帳
					else if (rowcount == 16)
						sValue = Utility.setCommaFormat((bean.getValue("field_120700") == null) ? "" : String.valueOf(bean.getValue("field_120700")));//內部融資
					else if (rowcount == 17)
						sValue = (bean.getValue("field_assets_rate") == null) ? "" : String.format("%s%%",String.valueOf(bean.getValue("field_assets_rate")));//資產報酬率
					else if (rowcount == 18)
						sValue = (bean.getValue("field_net_rate") == null) ? "" : String.format("%s%%",String.valueOf(bean.getValue("field_net_rate")));//淨值報酬率
					else if (rowcount == 19)
						sValue = (bean.getValue("field_backup_credit_rate") == null) ? "" : String.format("%s%%",String.valueOf(bean.getValue("field_backup_credit_rate")));//備呆占放款比率=備抵呆帳/放款
					else if (rowcount == 20)
						sValue = (bean.getValue("field_backup_over_rate") == null) ? "" : String.format("%s%%",String.valueOf(bean.getValue("field_backup_over_rate")));//備呆占逾放比率=備抵呆帳/逾期放款
					else if (rowcount == 21)
						sValue = (bean.getValue("field_debit_rate") == null) ? "" : String.format("%s%%",String.valueOf(bean.getValue("field_debit_rate")));//存款年長年率
					else if (rowcount == 22)
						sValue = (bean.getValue("field_credit_rate") == null) ? "" : String.format("%s%%",String.valueOf(bean.getValue("field_credit_rate")));//放款年成長率
					else if (rowcount == 23)
						sValue = (bean.getValue("bank_cnt") == null) ? "" : String.format("%s 家",String.valueOf(bean.getValue("bank_cnt")));//受輔導信用部家數
					else if (rowcount == 24)
						sValue = (bean.getValue("over_rate_count") == null) ? "" : String.format("%s 家",String.valueOf(bean.getValue("over_rate_count")));//逾放比率 15%以上家數
					else if (rowcount == 25)
						sValue = (bean.getValue("bis_count") == null) ? "" : String.format("%s 家",String.valueOf(bean.getValue("bis_count")));//資本適足率8%以下家數
					
					row = sheet.getRow(rowcount);
					cell = row.getCell((short) cellNum);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellValue(sValue);
				}
			}
		    String yrData[][] = {{"100","1,753,583","1,654,118","1,526,828","766,373","46.97%","99,465","4,825",
                                        "17,085","2.23%","21,710","4,529","0.28%","4.85%","2.83%","127.07%",
                                        "1.79%","3.19%","37家","16家","26家"},
                                 {"101","1,824,519","1,720,532","1,585,142","812,522","47.95%","103,987","4,885",
                                        "12,528","1.54%","23,060","4,978","0.27%","4.70%","2.84%","184.07%",
                                        "3.82%","6.02%","27家","11家","22家"},
                                 {"102","1,894,312","1,784,982","1,647,093","885,747","50.42%","109,330","4,895",
                                        "9,096","1.03%","25,116","4,367","0.26%","4.48%","2.84%","276.12%",
                                        "3.91%","9.01%","16家","5家","17家"},
                                 {"103","1,979,876","1,863,313","1,714,648","954,394","52.23%","116,563","5,571",
                                        "6,308","0.66%","26,996","4,669","0.28%","4.78%","2.83%","427.96%",
                                        "4.10%","7.75%","12家","3家","17家"},
                                 {"104","2,006,964","1,885,188","1,753,545","1,008,710","53.98%","121,776","5,839",
                                        "5,107","0.51%","29,299","5,078","0.29%","4.79%","2.90%","573.71%",
                                        "2.27%","5.69%","9家","1家","16家"},
                                 {"105","2,082,075","1,956,656","1,795,213","1,030,870","53.85%","125,419","4,770",
                                        "4,901","0.48%","31,077","5,033","0.23%","3.80%","3.01%","634.12%",
                                        "2.38%","2.20%","8家","0家","14家"}
		                         };

			//原100~105年底資料,以固定資料顯示
		    int yridx=-1;
		    for(int datacount = 0 ; datacount < dbData.size()-1 ; datacount++){//100~105年為固定資料
	            bean = (DataObject) dbData.get(datacount);
	            cellNum = datacount + 1;
	            System.out.println("datacount="+datacount+":s_year="+(Integer.parseInt(s_year)-6+datacount));
	            for(int yrcount=0;yrcount<=5;yrcount++){
	                if((Integer.parseInt(s_year)-6+datacount) == Integer.parseInt(yrData[yrcount][0])){
	                    yridx = yrcount;
	                    break;
	                }
	            }
	            if(yridx != -1){	                
	                for (int rowcount = 6; rowcount < 26; rowcount++) {	 
	                    row = sheet.getRow(rowcount);
	                    cell = row.getCell((short) cellNum);
	                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	                    sValue = yrData[yridx][rowcount-5];
	                    System.out.println("cellNum="+cellNum+":yrData["+yridx+"][0]="+ yrData[yridx][0]+":rowcount="+rowcount+".value="+sValue);
	                    cell.setCellValue(sValue);
	                }
	                yridx=-1;
	           }
			}//end of datacount
			
			//續3工作表:A4直印,縮放比96%
			sheet = wb.getSheetAt(2);
			if (sheet == null) {
				System.out.println("open sheet3 失敗");
			} else {
				System.out.println("open sheet3 成功");
			}
			ps = sheet.getPrintSetup(); //取得設定

			// 設定頁面符合列印大小
			sheet.setAutobreaks(false);
			ps.setScale((short)96); //列印縮放百分比
			ps.setPaperSize((short)9); //設定紙張大小 A4
			ps.setLandscape(false);//設定橫印
			
			sql.setLength(0);
			paramList.clear();
			
			row = sheet.getRow(0);
			cell = row.getCell((short) 0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(s_year+"年"+s_month+"月農漁會信用部重要數據" );
			
			row = sheet.getRow(1);
			cell = row.getCell((short) 5);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("列印日期：" + Utility.getCHTdate(printDate, 1)+printTime);
			
			row = sheet.getRow(3);
			cell = row.getCell((short) 0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("4.重設及新設信用部"+s_year+"年"+s_month+"月經營概況表" );
			
			sql.append(" select bn01_reset.bank_no, ");
			sql.append("        bank_name, "); //信用部
			sql.append("        to_number(nvl(output_order, '998')) as output_order, ");
			sql.append("        F_TRANSCHINESEDATE(start_date) start_date, "); //開始營運日期
			sql.append("        sum(decode(acc_code,'field_dc_rate',amt,0)) as field_dc_rate, "); //存放比率
			sql.append("        round(sum(decode(acc_code,'field_over',amt,0)) /? ,0) as field_over, "); //逾放金額
			paramList.add(1000);
			sql.append("        sum(decode(acc_code,'field_over_rate',amt,0)) as field_over_rate, "); //逾放比率
			sql.append("        round(sum(decode(acc_code,'field_320300',amt,0)) /? ,0) as field_320300, "); //本期損益
			paramList.add(1000);
			sql.append("        round(sum(decode(acc_code,'field_backup',amt,0)) /? ,0) as field_backup, "); //備抵呆帳
			paramList.add(1000);
			sql.append("        sum(decode(acc_code,'field_captial_rate',amt,0))  as field_captial_rate   "); //淨值佔風險性資產比率
			sql.append(" from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
			paramList.add(m_year);
			sql.append(" left join (select * from a01_operation_month where m_year=? and m_month=? and bank_type in ('6','7') and bank_code !='ALL')a01 on  bn01.bank_no = a01.bank_code ");
			paramList.add(s_year);
			paramList.add(s_month);
			sql.append(" ,bn01_reset ");
			sql.append(" where bn01.bank_no=bn01_reset.bank_no ");
			sql.append(" group by bn01_reset.bank_no,bank_name,output_order,setup_date,start_date,bn01_reset.add_date ");
			sql.append(" union ");
			sql.append(" select 'ALL' as bank_no,'總計' as bank_name, 999 as output_order, ");
			sql.append("        '' as start_date, ");
			sql.append("        0 as field_dc_rate, "); //存放比率
			sql.append("        round(sum(decode(acc_code,'field_over',amt,0)) /? ,0) as field_OVER, "); //逾放金額
			paramList.add(1000);
			sql.append("        decode(sum(decode(acc_code,'field_credit',amt,0)),0,0,round(sum(decode(acc_code,'field_over',amt,0)) / sum(decode(acc_code,'field_credit',amt,0)) *100 ,2)) as field_OVER_RATE, "); //逾放比率
			sql.append("        round(sum(decode(acc_code,'field_320300',amt,0)) /? ,0) as field_320300, "); //本期損益
			paramList.add(1000);
			sql.append("        round(sum(decode(acc_code,'field_backup',amt,0)) /? ,0) as field_BACKUP, "); //備抵呆帳
			paramList.add(1000);
			sql.append("        0 as field_CAPTIAL_RATE   "); //淨值佔風險性資產比率
			sql.append(" from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
			paramList.add(m_year);
			sql.append(" left join (select * from a01_operation_month where m_year=? and m_month=? and bank_type in ('6','7') and bank_code !='ALL')a01 on  bn01.bank_no = a01.bank_code ");
			paramList.add(s_year);
			paramList.add(s_month);
			sql.append(" ,bn01_reset ");
			sql.append(" where bn01.bank_no=bn01_reset.bank_no ");
			sql.append(" order by output_order ");
			
			dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "bank_no,bank_name,output_order,start_date,field_dc_rate,field_over,field_over_rate,field_320300,field_backup,field_captial_rate");
			System.out.println("4.dbData.size()="+dbData.size());
			
			int field_over=0,field_320300=0,field_backup=0;//106.12.26 add
			
			// 列印明細/合計
			for (int datacount = 0; datacount < dbData.size(); datacount++) {
				bean = (DataObject) dbData.get(datacount);
				rowNum = datacount+6;
				
				for (int cellcount = 0; cellcount < 9; cellcount++) {
					sValue = "";
					
					if (cellcount == 0 && datacount != dbData.size()-1){
						sValue = String.valueOf(datacount+1);//序號
					}else if (cellcount == 1){
						sValue = (bean.getValue("bank_name") == null) ? "0" : String.valueOf(bean.getValue("bank_name"));//信用部
					}else if (cellcount == 2 && datacount != dbData.size()-1){
						sValue = (bean.getValue("start_date") == null) ? "0" : String.valueOf(bean.getValue("start_date"));//開始營運日期
					}else if (cellcount == 3 && datacount != dbData.size()-1){
						sValue = String.format("%s%%",(bean.getValue("field_dc_rate") == null) ? "0" : String.valueOf(bean.getValue("field_dc_rate")));//存放比率
					}else if (cellcount == 4){
						sValue = Utility.setCommaFormat((bean.getValue("field_over") == null) ? "0" : String.valueOf(bean.getValue("field_over")));//逾放金額
					    if(!((String)bean.getValue("bank_name")).equals("總計")) field_over+=Integer.parseInt((bean.getValue("field_over") == null) ? "0" : String.valueOf(bean.getValue("field_over")));
					}else if (cellcount == 5){
						sValue = String.format("%s%%",(bean.getValue("field_over_rate") == null) ? "0" : String.valueOf(bean.getValue("field_over_rate")));//逾放比率
					}else if (cellcount == 6){
						sValue = Utility.setCommaFormat((bean.getValue("field_320300") == null) ? "0" : String.valueOf(bean.getValue("field_320300")));//本期損益
					    if(!((String)bean.getValue("bank_name")).equals("總計")) field_320300+=Integer.parseInt((bean.getValue("field_320300") == null) ? "0" : String.valueOf(bean.getValue("field_320300")));
					}else if (cellcount == 7){
						sValue = Utility.setCommaFormat((bean.getValue("field_backup") == null) ? "0" : String.valueOf(bean.getValue("field_backup")));//備抵呆帳
					    if(!((String)bean.getValue("bank_name")).equals("總計")) field_backup+=Integer.parseInt((bean.getValue("field_backup") == null) ? "0" : String.valueOf(bean.getValue("field_backup")));
					}else if (cellcount == 8 && datacount != dbData.size()-1){
						sValue = String.format("%s%%",(bean.getValue("field_captial_rate") == null) ? "0" : String.valueOf(bean.getValue("field_captial_rate")));//淨值佔風險性資產比率
					}
					row = sheet.createRow(rowNum);
					cell = row.createCell((short) cellcount);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs_center);
					cell.setCellValue(sValue);
				}
			}
			
			System.out.println("field_over="+field_over+":field_320300="+field_320300+":field_backup="+field_backup);
			
			rowNum = 6+dbData.size()-1;
            System.out.println("4.合計列="+rowNum);          
            row = sheet.getRow(rowNum);         
            //合計的部份.以原本各機構.先除以單位後.再加總 106.12.26 add 
            for (int cellcount = 4; cellcount < 8; cellcount++) {          
                if(cellcount != 5){
                    cell = row.createCell((short) cellcount);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellStyle(cs_center);
                }
                if (cellcount == 4){              
                    sValue = Utility.setCommaFormat(String.valueOf(field_over));                  
                }else if (cellcount == 6){
                    sValue = Utility.setCommaFormat(String.valueOf(field_320300));
                }else if (cellcount == 7){    
                    sValue = Utility.setCommaFormat(String.valueOf(field_backup));
                }
                if(cellcount != 5){
                    cell.setCellValue(sValue); 
                }
            } 
               
			
			//續4工作表:A4直印,縮放比100%
			sheet = wb.getSheetAt(3);
			if (sheet == null) {
				System.out.println("open sheet4 失敗");
			} else {
				System.out.println("open sheet4 成功");
			}
			ps = sheet.getPrintSetup(); //取得設定
			
			// 設定頁面符合列印大小
			sheet.setAutobreaks(false);
			ps.setScale((short)100); //列印縮放百分比
			ps.setPaperSize((short)9); //設定紙張大小 A4
			ps.setLandscape(false);//設定橫印
			
			row = sheet.getRow(0);
			cell = row.getCell((short) 0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(s_year+"年"+s_month+"月農漁會信用部重要數據" );
			
			row = sheet.getRow(1);
			cell = row.getCell((short) 3);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("列印日期：" + Utility.getCHTdate(printDate, 1)+printTime);
			
			row = sheet.getRow(3);
			cell = row.getCell((short) 0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("5."+s_year+"年"+s_month+"月底逾放比率超過15%及累積虧損超過信用部上年度決算淨值三分之一" );
			
			sql.setLength(0);
			paramList.clear();
			
			sql.append(" select bank_no, ");
			sql.append(" bank_name, "); //信用部
			sql.append(" amt  "); //逾放比率
			sql.append(" from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
			paramList.add(m_year);
			sql.append(" left join (select * from a01_operation_month where m_year=? and m_month=? and bank_type in ('6','7') and bank_code !='ALL')a01 on  bn01.bank_no = a01.bank_code ");
			paramList.add(s_year);
			paramList.add(s_month);
			sql.append(" where acc_code='field_over_rate' "); //逾放比率
			sql.append(" and amt > 15 ");
			sql.append(" order by amt ");

			dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "bank_no,bank_name,amt");
			System.out.println("5.dbData.size()="+dbData.size());
			
			rowNum = 3;
			
			if(dbData.size()>0){
				rowNum++;
				row = sheet.createRow(rowNum);
				num++;
				
				cell = row.createCell((short) 0);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs);
				cell.setCellValue("("+num+")逾放比率超過15%");
				
				rowNum++;
				
				row = sheet.createRow(rowNum);
				cell = row.createCell((short) 1);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("序號");
				
				cell = row.createCell((short) 2);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("信用部");
				
				cell = row.createCell((short) 3);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("逾放比率(%)");
				
				for (int datacount = 0; datacount < dbData.size(); datacount++) {
					bean = (DataObject) dbData.get(datacount);
					
					rowNum++;
					
					for (int cellcount = 1; cellcount < 4; cellcount++) {
						sValue = "";
						
						if (cellcount == 1)
							sValue = String.valueOf(datacount+1);//序號
						else if (cellcount == 2)
							sValue = (bean.getValue("bank_name") == null) ? "" : String.valueOf(bean.getValue("bank_name"));//信用部
						else if (cellcount == 3)
							sValue = String.format("%s%%",(bean.getValue("amt") == null) ? "0" : String.valueOf(bean.getValue("amt")));//逾放比率
						
						row = sheet.createRow(rowNum);
						cell = row.createCell((short) cellcount);
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell.setCellStyle(cs_center);
						cell.setCellValue(sValue);
					}
				}
			}
			
			sql.setLength(0);
			paramList.clear();
			
			sql.append(" select bank_no,bank_name, "); //信用部
			sql.append(" field_320100_rate  "); //累積虧損占信用部上年度決算淨值
			sql.append(" from ( ");
			sql.append("  select a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order,a01.bank_type, ");
			sql.append("        a01.bank_no , a01.BANK_NAME,COUNT_SEQ, field_SEQ, ");
			sql.append("        decode(field_990230-field_990240-field_992810,0,0,round(field_320100 / (field_990230-field_990240-field_992810) *100 ,2)) as field_320100_rate   "); //累積虧損占信用部上年度決算淨值之比率
			sql.append("      from (select a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order,a01.bank_type, ");
			sql.append("            a01.bank_no,a01.bank_name,1  AS  COUNT_SEQ,'A01'  as  field_SEQ, ");
			sql.append("            SUM(field_320100)   field_320100 , ");
			sql.append("            SUM(a02.field_990230) field_990230, ");
			sql.append("            SUM(a02.field_990240) field_990240, ");
			sql.append("            SUM(a99.field_992810) field_992810 ");
			sql.append("           from ( select nvl(cd01.hsien_id,' ')    as  hsien_id , ");
			sql.append("                 nvl(cd01.hsien_name,'OTHER') as  hsien_name, ");
			sql.append("                 cd01.FR001W_output_order  as FR001W_output_order,bn01.bank_type, ");
			sql.append("                 bn01.bank_no ,  bn01.bank_name, ");
			sql.append("                 round(sum(decode(a01.acc_code,'320100',amt,0)) /1,0) as field_320100 ");
			sql.append("                from  (select * from cd01 where cd01.hsien_id <> 'Y') cd01 ");
			sql.append("                left join (select * from wlx01 where m_year=? )wlx01 on wlx01.hsien_id=cd01.hsien_id ");
			paramList.add(m_year);
			sql.append("                left join (select * from bn01 where m_year=? )bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type in ('6','7') ");
			paramList.add(m_year);
			sql.append("                left join (select * from a01  where  a01.m_year  = ? and a01.m_month  = ? ) a01 on  bn01.bank_no = a01.bank_code ");
			paramList.add(s_year);
			paramList.add(s_month);
			sql.append(" 				group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order, ");
			sql.append(" 				bn01.bank_type,bn01.bank_no ,  bn01.BANK_NAME ");
			sql.append("           ) a01 ");
			sql.append("                 ,(select bank_code, ");
			sql.append("                        sum(decode(acc_code,'990230',amt,0)) as field_990230, ");
			sql.append("                        sum(decode(acc_code,'990240',amt,0)) as field_990240 ");
			sql.append("                  		from a02 where a02.m_year= ? and a02.m_month  = ? and  a02.ACC_code in ('990230','990240') ");
			paramList.add(s_year);
			paramList.add(s_month);
			sql.append("                    group by bank_code ");
			sql.append("                  ) a02 ");
			sql.append("                 ,(select bank_code, ");
			sql.append("                         sum(decode(acc_code,'992810',amt,0)) as field_992810 ");
			sql.append("                    from a99 where m_year= ?  and m_month  = ?  and  ACC_code in ('992810') ");
			paramList.add(s_year);
			paramList.add(s_month);
			sql.append("                    group by bank_code ");
			sql.append("                  ) a99 ");
			sql.append("              where   a01.bank_no=a02.bank_code(+) ");
			sql.append("              and a01.bank_no=a99.bank_code(+) ");
			sql.append("              and a01.bank_no <> ' ' ");
			sql.append("             GROUP  BY a01.hsien_id ,  a01.hsien_name,  a01.FR001W_output_order,a01.bank_type,a01.bank_no , a01.BANK_NAME ");
			sql.append("	) a01 ");
			sql.append(" )a01 ");
			sql.append(" where  field_320100_rate < -33 ");
			sql.append(" order by field_320100_rate ");

			dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "bank_no,bank_name,field_320100_rate");
			System.out.println("6.dbData.size()="+dbData.size());
			
			if(dbData.size()>0){
				rowNum++;
				row = sheet.createRow(rowNum);
				num++;
				
				cell = row.createCell((short) 0);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs);
				cell.setCellValue("("+num+")累積虧損超過信用部上年度決算淨值三分之一");
				
				rowNum++;
				
				row = sheet.createRow(rowNum);
				cell = row.createCell((short) 1);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("序號");
				
				cell = row.createCell((short) 2);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("信用部");
				
				cell = row.createCell((short) 3);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("累積虧損占信用部上年度決算淨值之比率(%)");
				
				for (int datacount = 0; datacount < dbData.size(); datacount++) {
					bean = (DataObject) dbData.get(datacount);
					
					rowNum++;
					
					for (int cellcount = 1; cellcount < 4; cellcount++) {
						sValue = "";
						
						if (cellcount == 1)
							sValue = String.valueOf(datacount+1);//序號
						else if (cellcount == 2)
							sValue = (bean.getValue("bank_name") == null) ? "" : String.valueOf(bean.getValue("bank_name"));//信用部
						else if (cellcount == 3)
							sValue = String.format("%s%%",(bean.getValue("field_320100_rate") == null) ? "0" : String.valueOf(bean.getValue("field_320100_rate")));//累積虧損占信用部上年度決算淨值之比率
						
						row = sheet.createRow(rowNum);
						cell = row.createCell((short) cellcount);
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell.setCellStyle(cs_center);
						cell.setCellValue(sValue);
					}
				}
			}
			
			rowNum++;
			if(num == 0){
				row = sheet.createRow(rowNum);
				cell = row.createCell((short) 1);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs);
				cell.setCellValue("無資料" );
			}
			
			num = 0;
			rowNum++;
			row = sheet.createRow(rowNum);
			cell = row.createCell((short) 0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(sheet.getRow(3).getCell((short) 0).getCellStyle());
			cell.setCellValue("6."+s_year+"年"+s_month+"月底農漁會信用部違反法定比率" );
			
			//(1)存放比率超過80%:field_month_dc_rate
			dbData = getA02Operation(s_year,s_month, "field_month_dc_rate");
			System.out.println("7.dbData.size()="+dbData.size());
			
			if(dbData.size()>0){
				rowNum++;
				row = sheet.createRow(rowNum);
				num++;
				
				cell = row.createCell((short) 0);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs);
				cell.setCellValue("("+num+")存放比率超過80%");
				
				rowNum++;
				
				row = sheet.createRow(rowNum);
				cell = row.createCell((short) 1);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("序號");
				
				cell = row.createCell((short) 2);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("信用部");
				
				cell = row.createCell((short) 3);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("存放比率(%)");
				
				rowNum = setData(sheet, cs_center, dbData, rowNum);
			}
			
			//(2)內部融資餘額占上年度信用部決算淨值超過60%:field_990210/(990230-990240)
			dbData = getA02Operation(s_year,s_month, "field_990210/(990230-990240)");
			System.out.println("8.dbData.size()="+dbData.size());
			
			if(dbData.size()>0){
				rowNum++;
				row = sheet.createRow(rowNum);
				num++;
				
				cell = row.createCell((short) 0);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs);
				cell.setCellValue("("+num+")內部融資餘額占上年度信用部決算淨值超過60%");
				
				rowNum++;
				
				row = sheet.createRow(rowNum);
				cell = row.createCell((short) 1);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("序號");
				
				cell = row.createCell((short) 2);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("信用部");
				
				cell = row.createCell((short) 3);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("內部融資餘額占上年度信用部決算淨值比率(%)");
				
				rowNum = setData(sheet, cs_center, dbData, rowNum);
			}
			
			//(3)贊助會員授信總額占贊助會員存款總額比率超過100%.150%.200%:field_990410/990420
			dbData = getA02Operation(s_year,s_month, "field_990410/990420");
			System.out.println("9.dbData.size()="+dbData.size());
			
			if(dbData.size()>0){
				rowNum++;
				row = sheet.createRow(rowNum);
				num++;
				
				cell = row.createCell((short) 0);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs);
				cell.setCellValue("("+num+")贊助會員授信總額占贊助會員存款總額比率超過100%、150%、200%");
				
				rowNum++;
				
				row = sheet.createRow(rowNum);
				cell = row.createCell((short) 1);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("序號");
				
				cell = row.createCell((short) 2);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("信用部");
				
				cell = row.createCell((short) 3);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("贊助會員授信總額占贊助會員存款總額比率(%)");
				
				rowNum = setData(sheet, cs_center, dbData, rowNum);
			}
			
			//(4)非會員無擔保消費性貸款總額占農(漁)會上年度全體決算淨值超過100%:field_990512/990320
			dbData = getA02Operation(s_year,s_month, "field_990512/990320");
			System.out.println("10.dbData.size()="+dbData.size());
			
			if(dbData.size()>0){
				rowNum++;
				row = sheet.createRow(rowNum);
				num++;
				
				cell = row.createCell((short) 0);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs);
				cell.setCellValue("("+num+")非會員無擔保消費性貸款總額占農(漁)會上年度全體決算淨值超過100%");
				
				rowNum++;
				
				row = sheet.createRow(rowNum);
				cell = row.createCell((short) 1);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("序號");
				
				cell = row.createCell((short) 2);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("信用部");
				
				cell = row.createCell((short) 3);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("非會員無擔保消費性貸款總額占農(漁)會上年度全體決算淨值比率(%)");
				
				rowNum = setData(sheet, cs_center, dbData, rowNum);
			}
			
			//(5)非會員授信總額占非會員存款總額比率超過100%、150%、200%:field_k/990620
			dbData = getA02Operation(s_year,s_month, "field_k/990620");
			System.out.println("11.dbData.size()="+dbData.size());
			
			if(dbData.size()>0){
				rowNum++;
				row = sheet.createRow(rowNum);
				num++;
				
				cell = row.createCell((short) 0);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs);
				cell.setCellValue("("+num+")非會員授信總額占非會員存款總額比率超過100%、150%、200%");
				
				rowNum++;
				
				row = sheet.createRow(rowNum);
				cell = row.createCell((short) 1);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("序號");
				
				cell = row.createCell((short) 2);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("信用部");
				
				cell = row.createCell((short) 3);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("非會員授信總額占非會員存款總額比率(%)");
				
				rowNum = setData(sheet, cs_center, dbData, rowNum);
			}
			
			//(6)自用住宅放款總額占定期性存款總額超過50%:field_990710/990720
			dbData = getA02Operation(s_year,s_month, "field_990710/990720");
			System.out.println("12.dbData.size()="+dbData.size());
			
			if(dbData.size()>0){
				rowNum++;
				row = sheet.createRow(rowNum);
				num++;
				
				cell = row.createCell((short) 0);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs);
				cell.setCellValue("("+num+")自用住宅放款總額占定期性存款總額超過50%");
				
				rowNum++;
				
				row = sheet.createRow(rowNum);
				cell = row.createCell((short) 1);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("序號");
				
				cell = row.createCell((short) 2);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("信用部");
				
				cell = row.createCell((short) 3);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("自用住宅放款總額占定期性存款總額比率(%)");
				
				rowNum = setData(sheet, cs_center, dbData, rowNum);
			}
			
			//(7)固定資產淨額占信用部上年度決算淨值超過100%:field_990810/(990230-990240)
			dbData = getA02Operation(s_year,s_month, "field_990810/(990230-990240)");
			System.out.println("13.dbData.size()="+dbData.size());
			
			if(dbData.size()>0){
				rowNum++;
				row = sheet.createRow(rowNum);
				num++;
				
				cell = row.createCell((short) 0);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs);
				cell.setCellValue("("+num+")固定資產淨額占信用部上年度決算淨值超過100%");
				
				rowNum++;
				
				row = sheet.createRow(rowNum);
				cell = row.createCell((short) 1);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("序號");
				
				cell = row.createCell((short) 2);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("信用部");
				
				cell = row.createCell((short) 3);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("固定資產淨額占信用部上年度決算淨值比率(%)");
				
				rowNum = setData(sheet, cs_center, dbData, rowNum);
			}
			
			//(8)外幣資產與外幣負債差額絕對值逾新台幣100萬元且占信用部上年度決算淨值超過5%;field_|990910-990920|/990230
			dbData = getA02Operation(s_year,s_month, "field_|990910-990920|/990230");
			System.out.println("14.dbData.size()="+dbData.size());
			
			if(dbData.size()>0){
				rowNum++;
				row = sheet.createRow(rowNum);
				num++;
				
				cell = row.createCell((short) 0);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs);
				cell.setCellValue("("+num+")外幣資產與外幣負債差額絕對值逾新台幣100萬元且占信用部上年度決算淨值超過5%");
				
				rowNum++;
				
				row = sheet.createRow(rowNum);
				cell = row.createCell((short) 1);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("序號");
				
				cell = row.createCell((short) 2);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("信用部");
				
				cell = row.createCell((short) 3);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("外幣資產與外幣負債差額絕對值逾新台幣100萬元且占信用部上年度決算淨值比率(萬元：%)");
				
				rowNum = setData(sheet, cs_center, dbData, rowNum);
			}
			
			//(9)理事、監事、職員及利害關係人擔保授信餘額占農(漁)會上年度全體決算淨值超過150%:field_991020/990320
			dbData = getA02Operation(s_year,s_month, "field_991020/990320");
			System.out.println("15.dbData.size()="+dbData.size());
			
			if(dbData.size()>0){
				rowNum++;
				row = sheet.createRow(rowNum);
				num++;
				
				cell = row.createCell((short) 0);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs);
				cell.setCellValue("("+num+")理事、監事、職員及利害關係人擔保授信餘額占農(漁)會上年度全體決算淨值超過150%");
				
				rowNum++;
				
				row = sheet.createRow(rowNum);
				cell = row.createCell((short) 1);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("序號");
				
				cell = row.createCell((short) 2);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("信用部");
				
				cell = row.createCell((short) 3);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("理事、監事、職員及利害關係人擔保授信餘額占農(漁)會上年度全體決算淨值比率(%)");
				
				rowNum = setData(sheet, cs_center, dbData, rowNum);
			}
			
			//(10)對鄉（鎮、市）公所授信未經其所隸屬之縣政府保證，及對直轄市、縣（市）政府投資經營之公營事業，其授信經該直轄市、縣（市）政府保證，兩者合計超過信用部上年度決算淨值:field_996114_996115/(990230-990240)
			dbData = getA02Operation(s_year,s_month, "field_996114_996115/(990230-990240)");
			System.out.println("16.dbData.size()="+dbData.size());
			
			if(dbData.size()>0){
				rowNum++;
				row = sheet.createRow(rowNum);
				num++;
				
				cell = row.createCell((short) 0);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs);
				cell.setCellValue("("+num+")對鄉（鎮、市）公所授信未經其所隸屬之縣政府保證，及對直轄市、縣（市）政府投資經營之公營事業，其授信經該直轄市、縣（市）政府保證，兩者合計超過信用部上年度決算淨值");
				
				rowNum++;
				
				row = sheet.createRow(rowNum);
				cell = row.createCell((short) 1);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("序號");
				
				cell = row.createCell((short) 2);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("信用部");
				
				cell = row.createCell((short) 3);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("對鄉（鎮、市）公所授信未經其所隸屬之縣政府保證，及對直轄市、縣（市）政府投資經營之公營事業，其授信經該直轄市、縣（市）政府保證，兩者合計超過信用部上年度決算淨值比率(%)");
				
				rowNum = setData(sheet, cs_center, dbData, rowNum);
			}
			
			//(11)資本適足率未達8%
			sql.setLength(0);
			paramList.clear();
			
			//資料SQL
			sql.append(" select bank_no,bank_name, "); //信用部名稱
			sql.append(" amt "); //BIS%
			sql.append(" from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
			paramList.add(m_year);
			sql.append(" left join (select * from a01_operation_month where m_year=? and m_month=? and bank_type in ('6','7') and bank_code !='ALL')a01 on  bn01.bank_no = a01.bank_code ");
			paramList.add(s_year);
			paramList.add(s_month);
			sql.append(" where acc_code='field_captial_rate' ");
			sql.append(" and amt < 8 "); //有違反
			sql.append(" order by amt ");
			
			dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "bank_no,bank_name,amt");
			System.out.println("17.dbData.size()="+dbData.size());

			if(dbData.size()>0){
				rowNum++;
				row = sheet.createRow(rowNum);
				num++;
				
				cell = row.createCell((short) 0);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs);
				cell.setCellValue("("+num+")資本適足率未達8%");
				
				rowNum++;
				
				row = sheet.createRow(rowNum);
				cell = row.createCell((short) 1);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("序號");
				
				cell = row.createCell((short) 2);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("信用部");
				
				cell = row.createCell((short) 3);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_center);
				cell.setCellValue("BIS(%)");
				
				rowNum = setData(sheet, cs_center, dbData, rowNum);
			}
			
			rowNum++;
			if(num == 0){
				row = sheet.createRow(rowNum);
				cell = row.createCell((short) 1);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs);
				cell.setCellValue("無資料" );
			}
			
			//續5工作表:A4直印,縮放比100%
			sheet = wb.getSheetAt(4);
			if (sheet == null) {
				System.out.println("open sheet5 失敗");
			} else {
				System.out.println("open sheet5 成功");
			}
			ps = sheet.getPrintSetup(); //取得設定

			// 設定頁面符合列印大小
			sheet.setAutobreaks(false);
			ps.setScale((short)100); //列印縮放百分比
			ps.setPaperSize((short)9); //設定紙張大小 A4
			ps.setLandscape(false);//設定橫印
			
			row = sheet.getRow(0);
			cell = row.getCell((short) 0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(s_year+"年"+s_month+"月農漁會信用部重要數據" );
			
			row = sheet.getRow(1);
			cell = row.getCell((short) 0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("列印日期：" + Utility.getCHTdate(printDate, 1)+printTime);
			
			dbData = getBuild("100000000",s_year,s_month);
			System.out.println("18.dbData.size()="+dbData.size());
			
			// 列印明細/合計
			for (int datacount = 0; datacount < dbData.size(); datacount++) {
				bean = (DataObject) dbData.get(datacount);
				
				row = sheet.getRow(4);
				cell = row.getCell((short) (1+(datacount*2)));
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				if(datacount != dbData.size()-1){
					cell.setCellValue((Integer.parseInt(s_year)-3+datacount)+"年底");
				}else{
					cell.setCellValue(s_year+"年"+s_month+"月底");
				}
				
				for (int count = 0; count < 6; count++) {
					rowNum = 6+(count/2);
					cellNum = 1+(datacount*2)+(count%2);
					
					sValue = "";
					
					if (count == 0)
						sValue  = (bean.getValue("agri_build_amt") == null) ? "" : String.format("%s 億元",String.valueOf(bean.getValue("agri_build_amt"))); //農業金庫.建築貸款餘額
					else if (count == 1)
						sValue  = (bean.getValue("field_agri_loan_rate") == null) ? "" : String.format("%s%%",String.valueOf(bean.getValue("field_agri_loan_rate"))); //農業金庫.建築貸款占放款比率
					else if (count == 2)
						sValue  = (bean.getValue("field_992710") == null) ? "" : String.format("%s 億元",String.valueOf(bean.getValue("field_992710"))); //農漁會信用.建築放款餘額
					else if (count == 3)
						sValue  = (bean.getValue("field_992710_rate") == null) ? "" : String.format("%s%%",String.valueOf(bean.getValue("field_992710_rate"))); //農漁會信用.建築貸款占放款比率
					else if (count == 4)
						sValue  = (bean.getValue("field_build_sum") == null) ? "" : String.format("%s 億元",String.valueOf(bean.getValue("field_build_sum"))); //建築貸款餘額.合計
					else if (count == 5)
						sValue  = (bean.getValue("field_buile_rate") == null) ? "" : String.format("%s%%",String.valueOf(bean.getValue("field_buile_rate"))); //合計.建築貸款占放款比率
					
					row = sheet.createRow(rowNum);
					cell = row.createCell((short) cellNum);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs_center);
					cell.setCellValue(sValue);
				}
			}
			           
			HSSFFooter footer = sheet.getFooter();
			footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
			FileOutputStream fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + openfile);
			wb.write(fout);
			// 儲存
			fout.close();
			System.out.println("儲存完成");
		} catch(Exception e) {
			System.out.println("createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}
	
	private static List getA01Operation(String unit, String s_year, String s_month){
		List dbData = new ArrayList();
		StringBuffer sql = new StringBuffer ();
		ArrayList paramList = new ArrayList();
		int startYear = Integer.parseInt(s_year);
		
		//資料SQL
		sql.append(" select nvl(round(field_assets /? ,0),0) as field_assets, "); //資產
		sql.append("           nvl(round(field_debt /? ,0),0) as field_debt, "); //負債
		sql.append("           nvl(round(a01.field_debit /? ,0),0) as field_debit, "); //存款
		sql.append("           nvl(round(a01.field_credit /? ,0),0) as field_credit,   "); //放款
		sql.append("           nvl(field_dc_rate,0) as field_dc_rate, "); //存放比率
		sql.append("           nvl(round(field_net /? ,0),0) as  field_net , "); //淨值
		sql.append("           nvl(round(field_320300 /? ,0),0)  as field_320300, "); //稅前純益
		sql.append("           nvl(round(field_over /? ,0),0)  as field_over, "); //逾放金額
		sql.append("           nvl(field_over_rate,0) as field_over_rate, "); //逾放比率
		sql.append("           nvl(round(field_backup /? ,0),0)  as  field_backup, "); //備抵呆帳
		sql.append("           nvl(round(field_120700 /? ,0),0)  as   field_120700, "); //內部融資
		sql.append("           nvl(field_assets_rate,0) as field_assets_rate, "); //資產報酬率
		sql.append(" 		   nvl(field_net_rate,0) as field_net_rate, "); //淨值報酬率
		sql.append("           nvl(field_backup_credit_rate,0) as field_backup_credit_rate, "); //備呆占放款比率=備抵呆帳/放款
		sql.append("           nvl(field_backup_over_rate,0) as field_backup_over_rate, "); //備呆占逾放比率=備抵呆帳/逾期放款
		sql.append("           nvl(decode(a01_last.field_debit,0,0,round((a01.field_debit-a01_last.field_debit)/ a01_last.field_debit*100 ,2)),0) as field_debit_rate, "); //存款年長年率
		sql.append("           nvl(decode(a01_last.field_credit,0,0,round((a01.field_credit-a01_last.field_credit)/ a01_last.field_credit*100 ,2)),0) as field_credit_rate, "); //放款年成長率
		sql.append("           nvl(bank_cnt,0) as bank_cnt, "); //受輔導信用部家數
		sql.append("           nvl(over_rate_count,0) as over_rate_count, "); //逾放比率 15%以上家數
		sql.append("           nvl(bis_count,0) as bis_count "); //資本適足率8%以下家數
		sql.append(" from( ");
		sql.append(" select round(sum(decode(a01.acc_code,'field_assets',a01.amt,0)) /1,0) as field_assets, "); //資產
		sql.append("        round(sum(decode(a01.acc_code,'field_debt',a01.amt,0)) /1,0) as field_debt,   "); //負債
		sql.append("        round(sum(decode(a01.acc_code,'field_debit',a01.amt,0)) /1,0) as field_debit, "); //存款
		sql.append("        round(sum(decode(a01.acc_code,'field_credit',a01.amt,0)) /1,0) as field_credit,   "); //放款
		sql.append("        sum(decode(a01.acc_code,'field_dc_rate',a01.amt,0)) as field_dc_rate, "); //存放比率
		sql.append("        round(sum(decode(a01.acc_code,'field_net',a01.amt,0)) /1,0) as field_net , "); //淨值
		sql.append("        round(sum(decode(a01.acc_code,'field_320300',a01.amt,0)) /1,0)  as field_320300, "); //稅前純益
		sql.append("        round(sum(decode(a01.acc_code,'field_over',a01.amt,0)) /1,0) as field_over, "); //逾放金額
		sql.append("        sum(decode(a01.acc_code,'field_over_rate',a01.amt,0)) as field_over_rate, "); //逾放比率
		sql.append("        round(sum(decode(a01.acc_code,'field_backup',a01.amt,0)) /1,0) as field_backup, "); //備抵呆帳
		sql.append("        round(sum(decode(a01.acc_code,'field_120700',a01.amt,0)) /1,0) as field_120700, "); //內部融資
		sql.append("        decode(sum(decode(a01.acc_code,'field_assets',a01.amt,0)),0,0,round((sum(decode(a01.acc_code,'field_320300',a01.amt,0))/sum(decode(a01.acc_code,'field_assets',a01.amt,0)))*100 ,2)) as field_assets_rate, "); //資產報酬率
		sql.append(" 	    decode(sum(decode(a01.acc_code,'field_net',a01.amt,0)),0,0,round((sum(decode(a01.acc_code,'field_320300',a01.amt,0))/sum(decode(a01.acc_code,'field_net',a01.amt,0)))*100 ,2)) as field_net_rate, "); //淨值報酬率
		sql.append("        sum(decode(a01.acc_code,'field_backup_credit_rate',a01.amt,0))  as field_backup_credit_rate, "); //備呆占放款比率=備抵呆帳/放款
		sql.append("        sum(decode(a01.acc_code,'field_backup_over_rate',a01.amt,0)) as field_backup_over_rate "); //, "); //備呆占逾放比率=備抵呆帳/逾期放款
		sql.append("  from (select * from a01_operation_month where m_year=? and m_month=? and bank_type in ('ALL') and hsien_id=' ')a01 ");
		sql.append(" )a01,  "); //本年底
		sql.append(" (select round(sum(decode(a01.acc_code,'field_debit',a01.amt,0)) /1,0) as field_debit, "); //存款
		sql.append("		round(sum(decode(a01.acc_code,'field_credit',a01.amt,0)) /1,0) as field_credit   "); //放款
		sql.append(" 	from (select * from a01_operation_month where m_year=? and m_month=? and bank_type in ('ALL') and hsien_id=' ')a01 ");
		sql.append(" )a01_last,  "); //上年底
		sql.append(" (select * from rpt_business where m_year=? and m_month=? )rpt_business, ");
		sql.append(" (select count(*) as bis_count ");
		sql.append(" 	from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
		sql.append(" 	left join (select * from a01_operation_month where m_year=? and m_month=? and bank_type in ('6','7') and bank_code !='ALL')a01 on  bn01.bank_no = a01.bank_code ");
		sql.append(" 	where acc_code='field_captial_rate' and amt < 8 ");
		sql.append(" )a01_bis,  "); //資本適足率8%以下家數
		sql.append(" (select count(*) as over_rate_count ");
		sql.append(" 	from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
		sql.append(" 	left join (select * from a01_operation_month where m_year=? and m_month=? and bank_type in ('6','7') and bank_code !='ALL')a01 on  bn01.bank_no = a01.bank_code ");
		sql.append(" 	where acc_code='field_over_rate' and amt > 15 ");
		sql.append(" )a01_over_rate  "); //逾放比率15%以上家數
		
		for(int i = startYear-6 ; i <= startYear; i++){//106.12.20 原當年度及前5年
		//for(int i = startYear ; i <= startYear; i++){//106.12.20  暫調整為只顯示當年度
			
			String m_year = (i < 100)?"99":"100";
			
			paramList.clear();
			
			paramList.add(unit);
			paramList.add(unit);
			paramList.add(unit);
			paramList.add(unit);
			paramList.add(unit);
			paramList.add(unit);
			paramList.add(unit);
			paramList.add(unit);
			paramList.add(unit);
			paramList.add(i);
			System.out.println("i="+i);
			if(i < startYear){
			   paramList.add("12"); 
			}else{
			   paramList.add(s_month);
			}
			paramList.add(i-1);
			//if(i < startYear){
	           paramList.add("12");//106.12.26 fix 調整當月底與前一年底金額比較  
	        //}else{
	        //   paramList.add(s_month);
	        //}
			paramList.add(i);
			if(i < startYear){
	           paramList.add("12"); 
	        }else{
	           paramList.add(s_month);
	        }
			paramList.add(m_year);
			paramList.add(i);
			if(i < startYear){
	           paramList.add("12"); 
	        }else{
	           paramList.add(s_month);
	        }
			paramList.add(m_year);
			paramList.add(i);
			if(i < startYear){
	           paramList.add("12"); 
	        }else{
	           paramList.add(s_month);
	        }
			
			List Data = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "field_assets,field_debt,field_debit,field_credit,field_dc_rate,field_net,field_320300,field_over,field_over_rate,field_backup,field_120700,field_assets_rate,field_net_rate,field_backup_credit_rate,field_backup_over_rate,field_debit_rate,field_credit_rate,bank_cnt,bis_count,over_rate_count");
			System.out.println("dbData.size()="+Data.size());
			if(Data.size()>0){
				dbData.add(Data.get(0));
			}else{
				dbData.add(new DataObject());
			}
		}
		return dbData;
	}
	
	private static List getA02Operation(String s_year, String s_month, String rangeName){
		StringBuffer sql = new StringBuffer ();
		ArrayList paramList = new ArrayList();
		String m_year = (Integer.parseInt(s_year) < 100)?"99":"100";
		
		//資料SQL
		sql.append(" select bank_no,bank_name, "); //信用部名稱
		sql.append(" amt "); //有違反的比率
		sql.append(" from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
		paramList.add(m_year);
		sql.append(" left join (select * from a02_operation where m_year=? and m_month=? )a02 on  bn01.bank_no = a02.bank_code ");
		paramList.add(s_year);
		paramList.add(s_month);
		sql.append(" where acc_code in (?) ");
		paramList.add(rangeName);
		sql.append(" and violate='Y' "); //有違反
		sql.append(" order by amt ");
		
		List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "bank_no,bank_name,amt");
		
		return dbData;
	}
	
	private static List getBuild(String unit, String s_year, String s_month){
		List dbData = new ArrayList();
		StringBuffer sql = new StringBuffer ();
		ArrayList paramList = new ArrayList();
		
		//資料SQL
		sql.append(" select nvl(round(agri_build_amt/? ,0),0) as agri_build_amt, "); //農業金庫.建築貸款餘額
		sql.append(" 	nvl(round(agri_loan_amt/? ,0),0) as agri_loan_amt, ");
		sql.append(" 	nvl(decode(agri_loan_amt,0,0,round((agri_build_amt/agri_loan_amt)*100 ,2)),0) as field_agri_loan_rate, "); //農業金庫.建築貸款占放款比率
		sql.append(" 	nvl(round(field_992710/? ,0),0) as field_992710, "); //農漁會信用.建築放款餘額
		sql.append(" 	nvl(round(field_CREDIT/? ,0),0) as field_credit, ");
		sql.append(" 	nvl(decode(field_credit,0,0,round((field_992710/field_credit)*100 ,2)),0) as field_992710_rate, "); //農漁會信用.建築貸款占放款比率
		sql.append(" 	nvl(round((agri_build_amt+field_992710)/ ? ,0),0) as field_build_sum, "); //建築貸款餘額.合計
		sql.append(" 	nvl(decode((agri_loan_amt+field_credit),0,0,round(((agri_build_amt+field_992710)/(agri_loan_amt+field_credit))*100 ,2)),0) as field_buile_rate "); //合計.建築貸款占放款比率
		sql.append(" from ( ");
		sql.append(" 	select  round(sum(decode(a01.acc_code,'field_credit',a01.amt,0)) /1,0) as field_credit,  "); //放款
		sql.append("            round(sum(decode(a01.acc_code,'field_992710',a01.amt,0)) /1,0) as field_992710 ");//--建築放款總額 106.12.29 add
		sql.append(" 	from (select * from a01_operation_month where m_year=? and m_month=? and bank_type in ('ALL') and hsien_id=' ')a01 ");
		sql.append(" )a01, (select * from rpt_business where m_year=? and m_month=?)rpt_business ");
		//106.12.29 原農漁會信用.建築放款餘額A99.992710調整為取關帳資料
		//sql.append(" (select round(sum(decode(a99.acc_code,'992710',a99.amt,0)) /1,0) as field_992710 "); //建築放款總額
		//sql.append(" 	from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
		//sql.append(" 	left join (select * from a99 where m_year=? and m_month=? and acc_code='992710')a99 on  bn01.bank_no = a99.bank_code ");
		//sql.append(" )a99 ");
		
		for(int year = Integer.parseInt(s_year)-3 ; year <=Integer.parseInt(s_year) ; year++){
			String m_year = (year< 100)?"99":"100";
			String month = "12";
			
			if(year == Integer.parseInt(s_year)) month = s_month;
			
			paramList.clear();
			
			paramList.add(unit);
			paramList.add(unit);
			paramList.add(unit);
			paramList.add(unit);
			paramList.add(unit);
			paramList.add(year);
			paramList.add(month);
			paramList.add(year);
			paramList.add(month);
			//paramList.add(m_year);
			//paramList.add(year);
			//paramList.add(month);
			
			List Data = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "agri_build_amt,agri_loan_amt,field_agri_loan_rate,field_992710,field_credit,field_992710_rate,field_build_sum,field_buile_rate");
			
			if(Data.size()>0){
				dbData.add(Data.get(0));
			}else{
				dbData.add( new DataObject());
			}
		}
		return dbData;
	}
	
	private static int setData(HSSFSheet sourceSheet, HSSFCellStyle cs, List dbData, int rowNum){
		for (int datacount = 0; datacount < dbData.size(); datacount++) {
			DataObject bean = (DataObject) dbData.get(datacount);
			
			HSSFRow row = null;  //宣告一列
			HSSFCell cell = null;//宣告一個儲存格
			
			rowNum++;
			
			for (int cellcount = 1; cellcount < 4; cellcount++) {
				String sValue = "";
				
				if (cellcount == 1)
					sValue = String.valueOf(datacount+1);//序號
				else if (cellcount == 2)
					sValue = (bean.getValue("bank_name") == null) ? "" : String.valueOf(bean.getValue("bank_name"));//信用部
				else if (cellcount == 3)
					sValue = (bean.getValue("amt") == null) ? "" : String.format("%s%%",String.valueOf(bean.getValue("amt")));
				
				row = sourceSheet.createRow(rowNum);
				cell = row.createCell((short) cellcount);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs);
				cell.setCellValue(sValue);
			}
		}
		return rowNum;
	}
	
	//cell 所要扣除的儲存格
	//fromValue原本的值
	//b是否原儲存格為數值且為百分比
	private static String cellSubMath(HSSFCell cell, String fromValue, boolean b){
		String result = "";
		Double d1 = 0.0;
		Double d2 = Double.parseDouble(fromValue);
		String str = "";
				
		switch (cell.getCellType()) {
	        case 0 :  // 數字日期型態
	        	d1 = cell.getNumericCellValue();
	            break;
	        case 1 :  // 字串型態
	        	str = cell.getStringCellValue();
	        	String s = str;
	        	//如果儲存格內除了有數值符號的文字會被判定為String
	        	s = s.indexOf("%")>0?s.split("%")[0]:s;
	        	s = s.indexOf("家")>0?s.split("家")[0]:s;
	        	d1 = Double.parseDouble(s);
	            break;
		}
		
		 if(b){
			 DecimalFormat df = new DecimalFormat("0.00");
			 result = df.format(d2-d1*100);
		 }else{
			 DecimalFormat df = new DecimalFormat("0");
			 if(str.indexOf("%")>0){
				 df = new DecimalFormat("0.00");
			 }
			 result = df.format(d2-d1);
		 }
		
		return result;
	}
		
}