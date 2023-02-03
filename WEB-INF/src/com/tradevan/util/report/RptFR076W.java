/*
106.06.02 create 經營月報表 by George
106.10.24 fix 調整附件(四).(五)依比率區間顯示其所在位置 by 2295
106.12.05 add 附件(七)增減比較%調整為個百分點 by 2295
106.12.06 add 附件(九)合計的部份.以原本各機構.先除以金額單位後.再加總  by 2295 
106.12.19 fix 附件(八)調整從107年底才開始抓DB資料,產生報表 by 2295
106.12.19 fix 調整原讀取a01_operation改為a01_operation_month(關帳資料) by 2295
//106.12.20 fix 附件(二)暫時調整為從10月份開始 by 2295
106.12.26 fix 附件(七)調整存款-一般放款/逾放金額-一般放款先除以單位.再相減  by 2295
106.01.25 ifx 附件(八)從107年底才開始抓DB資料 by 2295
*/
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.Region;

import java.io.*;
import java.text.DecimalFormat;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR076W {

	public static String createRpt(String s_year, String s_month, String bank_type, String unit)	{

		String errMsg = "";
		String m_year = (Integer.parseInt(s_year) < 100)?"99":"100"; 
		String unitName = Utility.getUnitName(unit);
		
		reportUtil reportUtil = new reportUtil();
		StringBuffer sql = new StringBuffer () ;
		ArrayList paramList = new ArrayList() ;
		
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

			String openfile = "經營月報表.xls";	
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
            
			ft.setFontHeightInPoints((short)12);
			ft.setFontName("標楷體");
			ft1.setFontHeightInPoints((short)12);
			ft1.setFontName("細明體");
			cs.setFont(ft1);
			cs.setBorderTop((short)0);
			cs.setBorderLeft((short)0);
			cs.setBorderLeft((short)0);
			cs.setBorderBottom((short)0);
			
			//開始寫入Excel的內容
			int rowNum = 0;
			int cellNum = 0;
			String sValue = "";
			DataObject bean = null;
			
			//(一) 全體農漁會信用部○年○月家數及經營情形表(A4直印,縮放比:100%)
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
			
			sql.setLength(0) ;
			
			//資料SQL
			sql.append(" select a01.bank_type, ")//6:農會 7:漁會 ALL:合計
				.append(" decode(a01.bank_type,'6',tbank_cnt_6,'7',tbank_cnt_7,'ALL',tbank_cnt_6+tbank_cnt_7,'') as tbank_cnt, ")//本部家數
				.append(" decode(a01.bank_type,'6',bank_cnt_6,'7',bank_cnt_7,'ALL',bank_cnt_6+bank_cnt_7,'') as bank_cnt, ")//分部家數
				.append(" round(sum(decode(acc_code,'field_debit',amt,0)) /?,0) as field_debit, ");//存款 
			paramList.add(unit);
			sql.append(" round(sum(decode(acc_code,'field_credit',amt,0)) /?,0) as field_credit, ");//放款
			paramList.add(unit);
			sql.append(" sum(decode(acc_code,'field_dc_rate',amt,0)) as field_dc_rate, ")//存放比率field_dc_rate
				.append(" round(sum(decode(acc_code,'field_net',amt,0)) /?,0) as field_net , ");//淨值field_net
			paramList.add(unit);
			sql.append(" round(sum(decode(acc_code,'field_320300',amt,0)) /?,0)  as field_320300, ");//稅前純益field_320300
			paramList.add(unit);
			sql.append(" round(sum(decode(acc_code,'field_over',amt,0)) /?,0)  as field_over, ");//逾放金額field_over
			paramList.add(unit);
			sql.append(" sum(decode(acc_code,'field_over_rate',amt,0)) as field_over_rate, ")//逾放比率field_over_rate
				.append(" round(sum(decode(acc_code,'field_backup',amt,0)) /?,0) as field_backup, ");//備抵呆帳field_backup
			paramList.add(unit);
			sql.append(" round(sum(decode(acc_code,'field_120700',amt,0)) /?,0) as field_120700, ");//內部融資field_120700
			paramList.add(unit);
			sql.append(" sum(decode(acc_code,'field_backup_credit_rate',amt,0)) as field_backup_credit_rate, ")//備呆占放款比率=備抵呆帳/放款 field_backup_credit_rate 
				.append(" sum(decode(acc_code,'field_backup_over_rate',amt,0)) as field_backup_over_rate ")//備呆占逾放比率=備抵呆帳/逾期放款    field_backup_over_rate       
				.append(" from (select * from a01_operation_month where m_year=? and m_month=? and bank_type in ('6','7','ALL') and hsien_id=' ')a01,(select * from rpt_business where m_year=? and m_month=? )rpt_business ");
			paramList.add(s_year);
			paramList.add(s_month);
			paramList.add(s_year);
			paramList.add(s_month);
			sql.append(" where a01.m_year=rpt_business.m_year and a01.m_month=rpt_business.m_month ")
				.append(" group by a01.bank_type,tbank_cnt_6,tbank_cnt_7,bank_cnt_6,bank_cnt_7 ");
			
			List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "bank_type,tbank_cnt,bank_cnt,field_debit,field_credit,field_dc_rate,field_net,field_320300,field_over,field_over_rate,field_backup,field_120700,field_backup_credit_rate,field_backup_over_rate");
			System.out.println("dbData.size()="+dbData.size());
			
			row = sheet.getRow(0);
			cell = row.getCell((short) 0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("附件1、全體農漁會信用部"+s_year+"年"+s_month+"月家數及經營情形表");
			
			row = sheet.getRow(1);
			cell = row.getCell((short) 3);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("單位：" + unitName);
			
			// 列印明細/合計
			for (int datacount = 0; datacount < dbData.size(); datacount++) {
				bean = (DataObject) dbData.get(datacount);
				rowNum = datacount;
				
				String bankType = (bean.getValue("bank_type") == null) ? "" : String.valueOf(bean.getValue("bank_type"));
				if(bankType.equals("6")){
					cellNum = 1;
				}else if(bankType.equals("7")){
					cellNum = 2;
				}else if(bankType.equals("ALL")){
					cellNum = 3;
				}
				
				for (int rowcount = 3; rowcount < 16; rowcount++) {
					if (rowcount == 3)
						sValue = (bean.getValue("tbank_cnt") == null) ? "0" : String.valueOf(bean.getValue("tbank_cnt"));//本部家數
					else if (rowcount == 4)
						sValue = (bean.getValue("bank_cnt") == null) ? "0" : String.valueOf(bean.getValue("bank_cnt"));//分部家數
					else if (rowcount == 5)
						sValue = (bean.getValue("field_debit") == null) ? "0" : String.valueOf(bean.getValue("field_debit"));//存款
					else if (rowcount == 6)
						sValue = (bean.getValue("field_credit") == null) ? "0" : String.valueOf(bean.getValue("field_credit"));//放款
					else if (rowcount == 7)
						sValue = String.format("%s%%",(bean.getValue("field_dc_rate") == null) ? "0" : String.valueOf(bean.getValue("field_dc_rate")));//存放比率
					else if (rowcount == 8)
						sValue = (bean.getValue("field_net") == null) ? "0" : String.valueOf(bean.getValue("field_net"));//淨值
					else if (rowcount == 9)
						sValue = (bean.getValue("field_320300") == null) ? "0" : String.valueOf(bean.getValue("field_320300"));//稅前純益
					else if (rowcount == 10)
						sValue = (bean.getValue("field_over") == null) ? "0" : String.valueOf(bean.getValue("field_over"));//逾放金額
					else if (rowcount == 11)
						sValue = String.format("%s%%",(bean.getValue("field_over_rate") == null) ? "0" : String.valueOf(bean.getValue("field_over_rate")));//逾放比率
					else if (rowcount == 12)
						sValue = (bean.getValue("field_backup") == null) ? "0" : String.valueOf(bean.getValue("field_backup"));//備抵呆帳
					else if (rowcount == 13)
						sValue = (bean.getValue("field_120700") == null) ? "0" : String.valueOf(bean.getValue("field_120700"));//內部融資
					else if (rowcount == 14)
						sValue = String.format("%s%%",(bean.getValue("field_backup_credit_rate") == null) ? "0" : String.valueOf(bean.getValue("field_backup_credit_rate")));//備呆占放款比率=備抵呆帳/放款
					else if (rowcount == 15)
						sValue = String.format("%s%%",(bean.getValue("field_backup_over_rate") == null) ? "0" : String.valueOf(bean.getValue("field_backup_over_rate")));//備呆占逾放比率=備抵呆帳/逾期放款
					
					row = sheet.getRow(rowcount);
					cell = row.getCell((short) cellNum);
					HSSFCellStyle formStyle = cell.getCellStyle();
					
					formStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
					formStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(formStyle);
					cell.setCellValue(sValue.indexOf("%")>=0?sValue:Utility.setCommaFormat(sValue));
					
				} // end of cellcount
			} // end of rowcount
			
			//(二) 全體農漁會信用部○年1～12月經營情形表(A3橫印,縮放比:95%)
			sheet = wb.getSheetAt(1);
			if (sheet == null) {
				System.out.println("open sheet2 失敗");
			} else {
				System.out.println("open sheet2 成功");
			}
			ps = sheet.getPrintSetup(); //取得設定

			// 設定頁面符合列印大小
			sheet.setAutobreaks(false);
			ps.setScale((short)95); //列印縮放百分比
			ps.setPaperSize( ( short )8 ); //設定紙張大小 A3
			ps.setLandscape(true);//設定橫印
			
			dbData = getA01Operation(unit, s_year, s_month);
			System.out.println("2.dbData.size()="+dbData.size());
			
			row = sheet.getRow(0);
			cell = row.getCell((short) 0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("附件2、全體農漁會信用部"+s_year+"年"+"1～"+s_month+"月經營情形表月家數及經營情形表");
			
			row = sheet.getRow(1);
			cell = row.getCell((short) 12);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("單位：" + unitName);
			
			cellNum = Integer.parseInt(s_month);
			
			// 列印明細/合計
			for (int datacount = 0; datacount < dbData.size(); datacount++) {//106.12.20 fix 原從1月份開始
	        //for (int datacount = 9; datacount < dbData.size(); datacount++) {//106.12.20 fix 暫時調整從10月份開始
				
				row = sheet.getRow(2);
				cell = row.getCell((short) (datacount+1));
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellValue(s_year+"年"+(datacount+1)+"月");
				
				bean = (DataObject) dbData.get(datacount);
				
				for (int rowcount = 3; rowcount < 19; rowcount++) {
					if (rowcount == 3)
						sValue = (bean.getValue("field_assets") == null) ? "0" : String.valueOf(bean.getValue("field_assets"));//資產
					else if (rowcount == 4)
						sValue = (bean.getValue("field_debt") == null) ? "0" : String.valueOf(bean.getValue("field_debt"));//負債
					else if (rowcount == 5)
						sValue = (bean.getValue("field_debit") == null) ? "0" : String.valueOf(bean.getValue("field_debit"));//存款
					else if (rowcount == 6)
						sValue = (bean.getValue("field_credit") == null) ? "0" : String.valueOf(bean.getValue("field_credit"));//放款
					else if (rowcount == 7)
						sValue = (bean.getValue("field_net") == null) ? "0" : String.valueOf(bean.getValue("field_net"));//淨值
					else if (rowcount == 8)
						sValue = (bean.getValue("field_320300") == null) ? "0" : String.valueOf(bean.getValue("field_320300"));//稅前純益
					else if (rowcount == 9)
						sValue = (bean.getValue("field_over") == null) ? "0" : String.valueOf(bean.getValue("field_over"));//逾放金額
					else if (rowcount == 10)
						sValue = String.format("%s%%",(bean.getValue("field_over_rate") == null) ? "0" : String.valueOf(bean.getValue("field_over_rate")));//逾放比率
					else if (rowcount == 11)
						sValue = (bean.getValue("field_backup") == null) ? "0" : String.valueOf(bean.getValue("field_backup"));//備抵呆帳
					else if (rowcount == 12)
						sValue = (bean.getValue("field_120700") == null) ? "0" : String.valueOf(bean.getValue("field_120700"));//內部融資
					else if (rowcount == 13)
						sValue = String.format("%s%%",(bean.getValue("field_assets_rate") == null) ? "0" : String.valueOf(bean.getValue("field_assets_rate")));//資產報酬率
					else if (rowcount == 14)
						sValue = String.format("%s%%",(bean.getValue("field_net_rate") == null) ? "0" : String.valueOf(bean.getValue("field_net_rate")));//淨值報酬率
					else if (rowcount == 15)
						sValue = String.format("%s%%",(bean.getValue("field_backup_credit_rate") == null) ? "0" : String.valueOf(bean.getValue("field_backup_credit_rate")));//備呆占放款比率=備抵呆帳/放款
					else if (rowcount == 16)
						sValue = String.format("%s%%",(bean.getValue("field_backup_over_rate") == null) ? "0" : String.valueOf(bean.getValue("field_backup_over_rate")));//備呆占逾放比率=備抵呆帳/逾期放款
					else if (rowcount == 17)
						sValue = String.format("%s%%",(bean.getValue("field_debit_rate") == null) ? "0" : String.valueOf(bean.getValue("field_debit_rate")));//存款成長率
					else if (rowcount == 18)
						sValue = String.format("%s%%",(bean.getValue("field_credit_rate") == null) ? "0" : String.valueOf(bean.getValue("field_credit_rate")));//放款成長率
					
					row = sheet.getRow(rowcount);
					cell = row.getCell((short) (datacount+1));
					HSSFCellStyle formStyle = cell.getCellStyle();
					
					formStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
					formStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(formStyle);
					cell.setCellValue(sValue.indexOf("%")>=0?sValue:Utility.setCommaFormat(sValue));
					
				} // end of cellcount
			} // end of rowcount
			
			//(三) 全體農漁會信用部各逾放比率區間之財務資料統計表(A4橫印,縮放比:90%)
			sheet = wb.getSheetAt(2);
			if (sheet == null) {
				System.out.println("open sheet3 失敗");
			} else {
				System.out.println("open sheet3 成功");
			}
			ps = sheet.getPrintSetup(); //取得設定
			
			//取得該查詢年月的最後一天
			String lastday = Utility.getCHTdate (Utility.getLastDay((Integer.parseInt(s_year)+1911)+s_month,"yyyymm").substring(0,10),1);
			
			// 設定頁面符合列印大小
			sheet.setAutobreaks(false);
			ps.setScale((short)90); //列印縮放百分比
			ps.setPaperSize((short)9); //設定紙張大小 A4
			ps.setLandscape(true);//設定橫印
			
			sql.setLength(0);
			paramList.clear();
			
			//資料SQL
			sql.append(" select dc_type, ") //TYPE0:逾放比率＜2％,TYPE1:第一區間,TYPE2:第二區間,TYPE3:第三區間,TYPE4:第四區間,TYPEALL:合計
				.append(" decode(dc_TYPE,'TYPEALL',sum(bank_code),count(*)) as  bank_count, ") //家數
				.append(" round(sum(field_DEBIT)/?,0) field_debit, "); //存款
			paramList.add(unit);
			sql.append(" round(sum(field_CREDIT)/?,0) field_credit, "); //放款
			paramList.add(unit);
			sql.append(" round(sum(field_NET)/?,0) field_net, "); //淨值
			paramList.add(unit);
			sql.append(" round(sum(field_320300)/?,0) field_320300, "); //稅前純益
			paramList.add(unit);
			sql.append(" round(sum(field_OVER)/?,0)  field_over, "); //逾放金額
			paramList.add(unit);
			sql.append(" decode(sum(field_CREDIT),0,0,round(sum(field_OVER)/sum(field_CREDIT) *100 ,2)) as field_over_rate, ") //逾放比率
				.append(" decode(sum(field_OVER) ,0,0,round(sum(field_BACKUP)/sum(field_OVER)  *100 ,2)) as field_backup_over_rate, ") //備呆占逾放比率=備抵呆帳/逾期放款  
				.append(" round(sum(field_120700)/?,0)  field_120700 "); //內部融資
			paramList.add(unit);
			sql.append(" from ( ")
				.append(" 	select bank_type,bank_code, ")
				.append(" CASE ")
				.append(" WHEN  sum(decode(acc_code,'field_over_rate',amt,0)) < 2 ")
				.append(" 	THEN 'TYPE0' ")
				.append(" WHEN (sum(decode(acc_code,'field_over_rate',amt,0)) >= 2  and sum(decode(acc_code,'field_over_rate',amt,0)) < 10) ")
				.append(" 	THEN 'TYPE1' ")
				.append(" WHEN (sum(decode(acc_code,'field_over_rate',amt,0)) >= 10  and sum(decode(acc_code,'field_over_rate',amt,0)) < 15) ")
				.append(" 	THEN 'TYPE2' ")
				.append(" WHEN (sum(decode(acc_code,'field_over_rate',amt,0)) >= 15  and sum(decode(acc_code,'field_over_rate',amt,0)) < 25) ")
				.append(" 	THEN 'TYPE3' ")
				.append(" WHEN sum(decode(acc_code,'field_over_rate',amt,0)) >= 25 ")
				.append(" 	THEN 'TYPE4' ")
				.append(" END as dc_TYPE, ") //逾放比率區間        
				.append(" round(sum(decode(acc_code,'field_debit',amt,0)) /1,0)  as field_debit, ") //存款
				.append(" round(sum(decode(acc_code,'field_credit',amt,0)) /1,0) as field_credit,   ") //放款    
				.append(" round(sum(decode(acc_code,'field_net',amt,0)) /1,0) as field_net , ") //淨值
				.append(" round(sum(decode(acc_code,'field_320300',amt,0)) /1,0) as field_320300, ") //稅前純益
				.append(" round(sum(decode(acc_code,'field_over',amt,0)) /1,0) as field_over, ") //逾放金額
				.append(" round(sum(decode(acc_code,'field_backup',amt,0)) /1,0) as field_backup, ") //備抵呆帳
				.append(" round(sum(decode(acc_code,'field_120700',amt,0)) /1,0) as field_120700 ") //內部融資
				.append(" from (select * from a01_operation_month where m_year=? and m_month=? and bank_type in ('6','7') and bank_code !='ALL')a01 ");
			paramList.add(s_year);
			paramList.add(s_month);
			sql.append(" group by bank_type,bank_code ")
				.append(" union ")
				.append(" select 'ALL',to_char(count(*)) as bank_no, ")
				.append(" 'TYPEALL' dc_TYPE, ") //逾放比率區間
				.append(" round(sum(field_DEBIT)/1,0) as field_debit, ") //存款
				.append(" round(sum(field_CREDIT)/1,0) as field_credit,   ") //放款    
				.append(" round(sum(field_NET)/1,0) as field_net , ") //淨值
				.append(" round(sum(field_320300)/1,0) as field_320300, ") //稅前純益
				.append(" round(sum(field_OVER)/1,0) as field_over, ") //逾放金額 
				.append(" round(sum(field_BACKUP)/1,0) as field_backup, ") //備抵呆帳
				.append(" round(sum(field_120700)/1,0) as field_120700 ") //內部融資
				.append(" from( ")
				.append(" select bank_type,bank_code, ")
				.append(" round(sum(decode(acc_code,'field_debit',amt,0)) /1,0)  as field_debit, ") //存款
				.append(" round(sum(decode(acc_code,'field_credit',amt,0)) /1,0) as field_credit,   ") //放款    
				.append(" round(sum(decode(acc_code,'field_net',amt,0)) /1,0) as field_net , ") //淨值
				.append(" round(sum(decode(acc_code,'field_320300',amt,0)) /1,0) as field_320300, ") //稅前純益
				.append(" round(sum(decode(acc_code,'field_over',amt,0)) /1,0) as field_over, ") //逾放金額
				.append(" round(sum(decode(acc_code,'field_backup',amt,0)) /1,0) as field_backup, ") //備抵呆帳
				.append(" round(sum(decode(acc_code,'field_120700',amt,0)) /1,0) as field_120700 ") //內部融資
				.append(" from (select * from a01_operation_month where m_year=? and m_month=? and bank_type in ('6','7') and bank_code !='ALL')a01 ");
			paramList.add(s_year);
			paramList.add(s_month);
			sql.append(" group by bank_type,bank_code)a01 ")
				.append(" )group by dc_TYPE order by dc_TYPE asc ");
			
			dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "dc_type,bank_count,field_debit,field_credit,field_net,field_320300,field_over,field_over_rate,field_backup_over_rate,field_120700");
			System.out.println("3.dbData.size()="+dbData.size());
			
			row = sheet.getRow(1);
			cell = row.getCell((short) 0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("                                 基準日："+lastday+"                    單位﹕" + unitName);
			
			// 列印明細/合計
			for (int datacount = 0; datacount < dbData.size(); datacount++) {
				bean = (DataObject) dbData.get(datacount);
				//rowNum = datacount+3;
				//TYPE0:逾放比率＜2％,TYPE1:第一區間,TYPE2:第二區間,TYPE3:第三區間,TYPE4:第四區間,TYPEALL:合計
				//106.10.24 fix 調整顯示位置 
				if("TYPE0".equals(bean.getValue("dc_type").toString())){
				   rowNum = 3; 
				}else if("TYPE1".equals(bean.getValue("dc_type").toString())){
		           rowNum = 4;
				}else if("TYPE2".equals(bean.getValue("dc_type").toString())){
	                   rowNum = 5;   
				}else if("TYPE3".equals(bean.getValue("dc_type").toString())){
                    rowNum = 6;       
				}else if("TYPE4".equals(bean.getValue("dc_type").toString())){
                    rowNum = 7;
				}else {
				    rowNum = 8;
				}           
				for (int cellcount = 1; cellcount < 10; cellcount++) {
					if (cellcount == 1)
						sValue = (bean.getValue("bank_count") == null) ? "0" : String.valueOf(bean.getValue("bank_count"));//家數
					else if (cellcount == 2)
						sValue = (bean.getValue("field_debit") == null) ? "0" : String.valueOf(bean.getValue("field_debit"));//存款
					else if (cellcount == 3)
						sValue = (bean.getValue("field_credit") == null) ? "0" : String.valueOf(bean.getValue("field_credit"));//放款
					else if (cellcount == 4)
						sValue = (bean.getValue("field_net") == null) ? "0" : String.valueOf(bean.getValue("field_net"));//淨值
					else if (cellcount == 5)
						sValue = (bean.getValue("field_320300") == null) ? "0" : String.valueOf(bean.getValue("field_320300"));//稅前純益
					else if (cellcount == 6)
						sValue = (bean.getValue("field_over") == null) ? "0" : String.valueOf(bean.getValue("field_over"));//逾放金額
					else if (cellcount == 7)
						sValue = String.format("%s%%",(bean.getValue("field_over_rate") == null) ? "0" : String.valueOf(bean.getValue("field_over_rate")));//逾放比率
					else if (cellcount == 8)
						sValue = String.format("%s%%",(bean.getValue("field_backup_over_rate") == null) ? "0" : String.valueOf(bean.getValue("field_backup_over_rate")));//備呆占逾放比率=備抵呆帳/逾期放款
					else if (cellcount == 9)
						sValue = (bean.getValue("field_120700") == null) ? "0" : String.valueOf(bean.getValue("field_120700"));//內部融資
					
					row = sheet.getRow(rowNum);
					cell = row.getCell((short) cellcount);
					HSSFCellStyle formStyle = cell.getCellStyle();
					
					formStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
					formStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(formStyle);
					cell.setCellValue(sValue.indexOf("%")>=0?sValue:Utility.setCommaFormat(sValue));
					
				} // end of cellcount
			} // end of rowcount
			
			//(四) 全體農漁會信用部各逾放比率區間之財務資料統計表(A4橫印,縮放比:100%)
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
			ps.setLandscape(true);//設定橫印
			
			sql.setLength(0);
			paramList.clear();
			
			//資料SQL
			sql.append(" select dc_type, "); //TYPE1:逾放比率 < 10%,TYPE2:10 <=逾放比率 < 15,TYPE3:15 <=逾放比率 < 25,TYPE4:逾放比率 >=25,TYPEALL:合計(家數)
			sql.append("        sum(netover_type1) as netover_type1, "); //70%以下[(淨值＋備抵)/逾放金額]
			sql.append("        sum(netover_type2) as netover_type2, "); //70%~100%
			sql.append("        sum(netover_type3) as netover_type3, "); //100%~200%
			sql.append("        sum(netover_type4) as netover_type4, "); //200%以上
			sql.append("        sum(netover_type1+netover_type2+netover_type3+netover_type4) as netover_sum  "); //合計(家數)
			sql.append(" from ( ");
			sql.append("   select bank_type,bank_code,dc_type, ");
			sql.append("          decode(netover_type,'1',1,0) as netover_type1, ");
			sql.append("          decode(netover_type,'2',1,0) as netover_type2, ");
			sql.append("          decode(netover_type,'3',1,0) as netover_type3, ");
			sql.append("          decode(netover_type,'4',1,0) as netover_type4  ");        
			sql.append("   from ");
			sql.append("   ( ");
			sql.append("     select bank_type,bank_code,dc_type, ");
			sql.append("            CASE ");
			sql.append("            WHEN sum(field_NETOVER_RATE) < 70 ");
			sql.append("                 THEN '1' ");
			sql.append("            WHEN (sum(field_NETOVER_RATE) >= 70  and sum(field_NETOVER_RATE) < 100) ");
			sql.append("                 THEN '2' ");
			sql.append("            WHEN (sum(field_NETOVER_RATE) >= 100  and sum(field_NETOVER_RATE) < 200) ");
			sql.append("                 THEN '3' ");
			sql.append("            WHEN sum(field_NETOVER_RATE) >= 200 ");
			sql.append("                 THEN '4' ");
			sql.append("            END as netover_TYPE "); //(淨值＋備抵)/逾放金額
			sql.append("     from ");
			sql.append("     ( ");
			sql.append("     select bank_type,bank_code, ");
			sql.append("            CASE ");
			sql.append("            WHEN sum(field_over_rate) < 10 ");
			sql.append("                 THEN 'TYPE1' ");
			sql.append("            WHEN (sum(field_over_rate) >= 10  and sum(field_over_rate) < 15) ");
			sql.append("                 THEN 'TYPE2' ");
			sql.append("            WHEN (sum(field_over_rate) >= 15  and sum(field_over_rate) < 25) ");
			sql.append("                 THEN 'TYPE3' ");
			sql.append("            WHEN sum(field_over_rate) >= 25 ");
			sql.append("                 THEN 'TYPE4' ");
			sql.append("            END as dc_TYPE, "); //逾放比率區間        
			sql.append("            decode(sum(field_OVER) ,0,999,round((sum(field_NET)+sum(field_BACKUP)) / sum(field_OVER) *100 ,2)) as field_NETOVER_RATE "); // (淨值＋備抵) /逾放金額               
			sql.append("     from( ");
			sql.append("     select bank_type,bank_code, ");
			sql.append("            round(sum(decode(acc_code,'field_net',amt,0)) /1,0) as field_NET, "); //淨值
			sql.append("            round(sum(decode(acc_code,'field_over',amt,0)) /1,0) as field_OVER, "); //逾放金額
			sql.append("            round(sum(decode(acc_code,'field_backup',amt,0)) /1,0) as  field_BACKUP, "); //備抵呆帳          
			sql.append("            sum(decode(acc_code,'field_over_rate',amt,0))   as field_over_rate "); //逾放比率
			sql.append("     from (select * from a01_operation_month where m_year=? and m_month=? and bank_type in ('6','7') and bank_code !='ALL')a01 ");
			paramList.add(s_year);
			paramList.add(s_month);
			sql.append("     group by bank_type,bank_code)a01 ");
			sql.append("     group by bank_type,bank_code ");
			sql.append("     )group by bank_type,bank_code,dc_type ");
			sql.append("    ) ");
			sql.append("   )group by dc_type  "); //逾放比率區間家數
			sql.append(" union "); 
			sql.append(" select 'TYPEALL' as dc_type, ");
			sql.append("        sum(netover_type1) as netover_type1, ");
			sql.append("        sum(netover_type2) as netover_type2, ");
			sql.append("        sum(netover_type3) as netover_type3, ");
			sql.append("        sum(netover_type4) as netover_type4, ");
			sql.append("        sum(netover_type1+netover_type2+netover_type3+netover_type4) as netover_sum ");
			sql.append(" from ( ");
			sql.append(" select bank_type,bank_code,dc_type, ");
			sql.append("        decode(netover_type,'1',1,0) as netover_type1, ");
			sql.append("        decode(netover_type,'2',1,0) as netover_type2, ");
			sql.append("        decode(netover_type,'3',1,0) as netover_type3, ");
			sql.append("        decode(netover_type,'4',1,0) as netover_type4  ");
			sql.append(" from ");
			sql.append("  (  ");
			sql.append("    select bank_type,bank_code,dc_type, ");
			sql.append("              CASE ");
			sql.append("           WHEN sum(field_NETOVER_RATE) < 70 ");
			sql.append("                THEN '1' ");
			sql.append("           WHEN (sum(field_NETOVER_RATE) >= 70  and sum(field_NETOVER_RATE) < 100) ");
			sql.append("                THEN '2' ");
			sql.append("           WHEN (sum(field_NETOVER_RATE) >= 100 and sum(field_NETOVER_RATE) < 200) ");
			sql.append("                THEN '3' ");
			sql.append("           WHEN sum(field_NETOVER_RATE) >= 200 ");
			sql.append("                THEN '4' ");
			sql.append("           END as netover_TYPE "); //(淨值＋備抵) /逾放金額  區間   
			sql.append("    from ");
			sql.append("    ( ");
			sql.append("     select bank_type,bank_code, ");
			sql.append("            CASE ");
			sql.append("            WHEN sum(field_over_rate) < 2 ");
			sql.append("                 THEN 'TYPE0' ");
			sql.append("            WHEN (sum(field_over_rate) >= 2  and sum(field_over_rate) < 10) ");
			sql.append("                 THEN 'TYPE1' ");
			sql.append("            WHEN (sum(field_over_rate) >= 10 and sum(field_over_rate) < 15) ");
			sql.append("                 THEN 'TYPE2' ");
			sql.append("            WHEN (sum(field_over_rate) >= 15 and sum(field_over_rate) < 25) ");
			sql.append("                 THEN 'TYPE3' ");
			sql.append("            WHEN sum(field_over_rate) >= 25 ");
			sql.append("                 THEN 'TYPE4' ");
			sql.append("            END as dc_TYPE, "); //逾放比率區間        
			sql.append("            decode(sum(field_OVER) ,0,999,round((sum(field_NET)+sum(field_BACKUP)) / sum(field_OVER) *100 ,2)) as field_NETOVER_RATE "); // (淨值＋備抵) /逾放金額               
			sql.append("     from( ");
			sql.append("     select bank_type,bank_code, ");         
			sql.append("            round(sum(decode(acc_code,'field_net',amt,0)) /1,0) as field_NET, "); //淨值
			sql.append("            round(sum(decode(acc_code,'field_over',amt,0)) /1,0) as field_OVER, "); //逾放金額
			sql.append("            round(sum(decode(acc_code,'field_backup',amt,0)) /1,0) as  field_BACKUP, "); //備抵呆帳          
			sql.append("            sum(decode(acc_code,'field_over_rate',amt,0))   as field_over_rate "); //逾放比率
			sql.append("     from (select * from a01_operation_month where m_year=? and m_month=? and bank_type in ('6','7') and bank_code !='ALL')a01 ");
			paramList.add(s_year);
			paramList.add(s_month);
			sql.append("     group by bank_type,bank_code)a01 ");
			sql.append("     group by bank_type,bank_code ");
			sql.append("    )group by bank_type,bank_code,dc_type ");
			sql.append("  ) ");
			sql.append(" ) "); //合計家數
			sql.append(" order by dc_type ");
			
			dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "dc_type,netover_type1,netover_type2,netover_type3,netover_type4,netover_sum");
			System.out.println("4.dbData.size()="+dbData.size());
			
			row = sheet.getRow(1);
			cell = row.getCell((short) 0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("基準日："+lastday);
			
			// 列印明細/合計
			for (int datacount = 0; datacount < dbData.size(); datacount++) {
				bean = (DataObject) dbData.get(datacount);
				rowNum = datacount+4;
				//TYPE1:逾放比率 < 10%,TYPE2:10 <=逾放比率 < 15,TYPE3:15 <=逾放比率 < 25,TYPE4:逾放比率 >=25,TYPEALL:合計(家數)
				//106.10.24 fix 調整顯示位置 
                if("TYPE1".equals(bean.getValue("dc_type").toString())){
                   rowNum = 4; 
                }else if("TYPE2".equals(bean.getValue("dc_type").toString())){
                   rowNum = 5;
                }else if("TYPE3".equals(bean.getValue("dc_type").toString())){
                   rowNum = 6;                  
                }else if("TYPE4".equals(bean.getValue("dc_type").toString())){
                   rowNum = 7;
                }else {
                   rowNum = 8;
                }   
				for (int cellcount = 1; cellcount < 6; cellcount++) {
					if (cellcount == 1)
						sValue = (bean.getValue("netover_type1") == null) ? "0" : String.valueOf(bean.getValue("netover_type1"));//70%以下[(淨值＋備抵)/逾放金額]
					else if (cellcount == 2)
						sValue = (bean.getValue("netover_type2") == null) ? "0" : String.valueOf(bean.getValue("netover_type2"));//70%~100%
					else if (cellcount == 3)
						sValue = (bean.getValue("netover_type3") == null) ? "0" : String.valueOf(bean.getValue("netover_type3"));//100%~200%
					else if (cellcount == 4)
						sValue = (bean.getValue("netover_type4") == null) ? "0" : String.valueOf(bean.getValue("netover_type4"));//200%以上
					else if (cellcount == 5)
						sValue = (bean.getValue("netover_sum") == null) ? "0" : String.valueOf(bean.getValue("netover_sum"));//合計(家數)
					
					row = sheet.getRow(rowNum);
					cell = row.getCell((short) cellcount);
					HSSFCellStyle formStyle = cell.getCellStyle();
					
					formStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
					formStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(formStyle);
					cell.setCellValue(sValue.indexOf("%")>=0?sValue:Utility.setCommaFormat(sValue));
					
				} 
			} 
			
			//(五)○年○月全體農漁會信用部逾放區間明細表(A3橫印,縮放比:60%)
			sheet = wb.getSheetAt(4);
			if (sheet == null) {
				System.out.println("open sheet5 失敗");
			} else {
				System.out.println("open sheet5 成功");
			}
			ps = sheet.getPrintSetup(); //取得設定
			
			// 設定頁面符合列印大小
			sheet.setAutobreaks(false);
			ps.setScale((short)60); //列印縮放百分比
			ps.setPaperSize((short)8); //設定紙張大小 A3
			ps.setLandscape(true);//設定橫印
			
			sql.setLength(0);
			paramList.clear();
			
			//資料SQL
			sql.append(" select bank_type,bank_no,bank_name, "); //單位名稱
			sql.append("        CASE ");
			sql.append("        WHEN sum(field_over_rate) < 1  ");
			sql.append("             THEN 'TYPE1' ");
			sql.append("        WHEN (sum(field_over_rate) >= 1  and sum(field_over_rate) < 15) ");
			sql.append("             THEN 'TYPE2' ");
			sql.append("        WHEN (sum(field_over_rate) >= 15  and sum(field_over_rate) < 25) ");
			sql.append("             THEN 'TYPE3' ");
			sql.append("        WHEN sum(field_over_rate) >= 25 ");
			sql.append("             THEN 'TYPE4' ");
			sql.append("        END as over_type, "); //逾放比率區間 TYPE1:0~1%(不含1%) TYPE2:1%-15%(不含15%) TYPE3:15%-25%(不含25%)  TYPE4:25%以上
			sql.append("        sum(field_over_rate) field_over_rate  "); //逾放比率        
			sql.append(" from ");
			sql.append(" ( ");
			sql.append(" select bn01.bank_type,bn01.bank_no , bn01.BANK_NAME,fr001w_output_order, ");
			sql.append("        sum(decode(a01.acc_code,'field_over_rate',amt,0))  as field_over_rate "); //逾放比率
			sql.append(" from (select * from bn01 where m_year = ? and bank_type in ('6','7') and bn_type <> '2')bn01 ");
			paramList.add(m_year);
			sql.append("      left join (select m_year,m_month,bank_code,acc_code,amt from a01_operation_month ");
			sql.append("                 where m_year=? and m_month=?  and acc_code='field_over_rate'  "); //逾放比率
			paramList.add(s_year);
			paramList.add(s_month);
			sql.append("                ) a01  on  bn01.bank_no = a01.bank_code ");
			sql.append("      left join (select bank_no,wlx01.hsien_id,fr001w_output_order from wlx01 left join cd01 on wlx01.hsien_id=cd01.hsien_id where m_year=100)wlx01 ");
			sql.append("      on bn01.bank_no=wlx01.bank_no ");
			sql.append(" group by bn01.bank_type,bn01.bank_no,bn01.BANK_NAME,fr001w_output_order ");
			sql.append(" )group by bank_type,bank_no,bank_name,fr001w_output_order ");
			sql.append(" order by over_type,field_over_rate,fr001w_output_order,bank_no ");
			
			dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "bank_type,bank_no,bank_name,over_type,field_over_rate");
			System.out.println("5.dbData.size()="+dbData.size());
			
			row = sheet.getRow(0);
			cell = row.getCell((short) 0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("附件5、"+s_year+"年"+s_month+"月全體農漁會信用部逾放區間明細表");
			
			String over_type = ""; //逾放比率區間 TYPE1:0~1%(不含1%) TYPE2:1%-15%(不含15%) TYPE3:15%-25%(不含25%)  TYPE4:25%以上
			int counts = 0; //筆數
			int num = 1; //序號
			int type1counts = 0;
			int type2counts = 0;
			int type3counts = 0;
			int type4counts = 0;
			
			for (int datacount = 0; datacount < dbData.size(); datacount++) {
				bean = (DataObject) dbData.get(datacount);
				if("TYPE1".equals(bean.getValue("over_type"))){
					type1counts++;
				}else if("TYPE2".equals(bean.getValue("over_type"))){
					type2counts++;
				}else if("TYPE3".equals(bean.getValue("over_type"))){
					type3counts++;
				}else if("TYPE4".equals(bean.getValue("over_type"))){
					type4counts++;
				}
			}
			
			int c = type1counts/70+(type1counts%70>0?1:0);
			if(type1counts == 0) c = 1;
			for(int i =0 ; i < c; i++){
				setTitle(sheet,cs_center,"TYPE1",i,type1counts,i*4,3,5);
			}
			int j = type2counts/70+(type2counts%70>0?1:0)+c;
			if(type2counts == 0) j = c+1;
			for(int i =c ; i < j; i++){
				setTitle(sheet,cs_center,"TYPE2",i-c,type2counts,i*4,3,5);
			}
			c = j;
			j += type3counts/70+(type3counts%70>0?1:0);
			if(type3counts == 0) j = c+1;
			for(int i =c ; i < j; i++){
				setTitle(sheet,cs_center,"TYPE3",i-c,type3counts,i*4,3,5);
			}
			for(int i =c ; i < j; i++){
				setTitle(sheet,cs_center,"TYPE4",i-c,type4counts,i*4,7+type3counts,5);
			}
			
			// 列印明細/合計
			for (int datacount = 0; datacount < dbData.size(); datacount++) {
				bean = (DataObject) dbData.get(datacount);
				
				if(!over_type.equals(bean.getValue("over_type"))&&!over_type.equals("")){
					num = 1; //將不同逾放比率的序號歸零
					 if("TYPE4".equals(bean.getValue("over_type"))){
						 counts = counts +4;
					 }else{
						 counts = ((counts/70)+1)*70;
					 }
				}
				
				rowNum = (counts%70)+3;
				cellNum = (counts/70)*4;
				
				for (int cellcount = 0; cellcount < 3; cellcount++) {
					if (cellcount == 0)
						sValue =  String.valueOf(num); //序號
					else if (cellcount == 1)
						sValue = (bean.getValue("bank_name") == null) ? "0" : String.valueOf(bean.getValue("bank_name")); //單位名稱
					else if (cellcount == 2)
						sValue = (bean.getValue("field_over_rate") == null) ? "0" : String.valueOf(bean.getValue("field_over_rate")); //逾放比率
						
					row = sheet.createRow(rowNum);
					cell = row.createCell((short) (cellcount+cellNum));
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs_center);
					cell.setCellValue(sValue);
				}
				
				over_type = (bean.getValue("over_type") == null) ? "" : String.valueOf(bean.getValue("over_type"));
				counts++;
				num++;
			}
			
			//(六)○年○月全體農漁會信用部BIS區間明細表(A3橫印,縮放比:65%)
			sheet = wb.getSheetAt(5);
			if (sheet == null) {
				System.out.println("open sheet6 失敗");
			} else {
				System.out.println("open sheet6 成功");
			}
			ps = sheet.getPrintSetup(); //取得設定
			
			// 設定頁面符合列印大小
			sheet.setAutobreaks(false);
			ps.setScale((short)65); //列印縮放百分比
			ps.setPaperSize((short)8); //設定紙張大小 A3
			ps.setLandscape(true);//設定橫印
			
			sql.setLength(0);
			paramList.clear();
			
			//資料SQL
			sql.append(" select bank_type,bank_no,bank_name, "); //單位名稱
			sql.append("        CASE ");
			sql.append("        WHEN sum(a05_bis) < 8 ");
			sql.append("             THEN 'TYPE1' ");
			sql.append("        WHEN (sum(a05_bis) >= 8  and sum(a05_bis) < 12) ");
			sql.append("             THEN 'TYPE2' ");
			sql.append("        WHEN sum(a05_bis) >= 12 ");
			sql.append("             THEN 'TYPE3' ");
			sql.append("        END as bis_type, "); //BIS區間 TYPE1:未達8% TYPE2: 8%~12%(不含12%) TYPE3:12%以上
			sql.append("        sum(a05_bis) a05_bis  "); //BIS(淨值佔風險性資產比率)
			sql.append(" from ");
			sql.append(" ( ");
			sql.append(" select bn01.bank_type,bn01.bank_no , bn01.BANK_NAME,fr001w_output_order, ");
			sql.append("        sum(decode(a05.acc_code,'91060P',amt,0))  as a05_bis "); //BIS(淨值佔風險性資產比率)
			sql.append(" from (select * from bn01 where m_year = 100 and bank_type in ('6','7') and bn_type <> '2')bn01 ");
			sql.append("       left join (select m_year,m_month,bank_code,acc_code,round(amt/1000,2) as amt from a05 ");
			sql.append("                   where m_year=? and m_month=?  and acc_code='91060P'  "); //bis
			paramList.add(s_year);
			paramList.add(s_month);
			sql.append("                  ) a05  on  bn01.bank_no = a05.bank_code ");
			sql.append("       left join (select bank_no,wlx01.hsien_id,fr001w_output_order from wlx01 left join cd01 on wlx01.hsien_id=cd01.hsien_id where m_year=100)wlx01 ");
			sql.append("                  on bn01.bank_no=wlx01.bank_no ");
			sql.append(" group by bn01.bank_type,bn01.bank_no,bn01.BANK_NAME,fr001w_output_order ");
			sql.append(" )group by bank_type,bank_no,bank_name,fr001w_output_order ");
			sql.append(" order by bis_type,a05_bis,fr001w_output_order,bank_no ");
			
			dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "bank_type,bank_no,bank_name,bis_type,a05_bis");
			System.out.println("6.dbData.size()="+dbData.size());
			
			row = sheet.getRow(0);
			cell = row.getCell((short) 0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("附件6、"+s_year+"年"+s_month+"月全體農漁會信用部BIS區間明細表");
			
			// 列印明細/合計
			String bis_type = ""; //BIS區間 TYPE1:未達8% TYPE2: 8%~12%(不含12%) TYPE3:12%以上
			counts = 0; //筆數
			num = 1; //序號
			type1counts = 0;
			type2counts = 0;
			type3counts = 0;
			
			for (int datacount = 0; datacount < dbData.size(); datacount++) {
				bean = (DataObject) dbData.get(datacount);
				if("TYPE1".equals(bean.getValue("bis_type"))){
					type1counts++;
				}else if("TYPE2".equals(bean.getValue("bis_type"))){
					type2counts++;
				}else if("TYPE3".equals(bean.getValue("bis_type"))){
					type3counts++;
				}
			}
			
			c = type1counts/70+(type1counts%70>0?1:0);
			if(type1counts == 0) c = 1;
			for(int i =0 ; i < c; i++){
				setTitle(sheet,cs_center,"TYPE1",i,type1counts,i*4,3,6);
			}
			j = type2counts/70+(type2counts%70>0?1:0)+c;
			if(type2counts == 0) j = c+1;
			for(int i =c ; i < j; i++){
				setTitle(sheet,cs_center,"TYPE2",i-c,type2counts,i*4,3,6);
			}
			c = j;
			j += type3counts/70+(type3counts%70>0?1:0);
			if(type3counts == 0) j = c+1;
			for(int i =c ; i < j; i++){
				setTitle(sheet,cs_center,"TYPE3",i-c,type3counts,i*4,3,6);
			}
			
			// 列印明細/合計
			for (int datacount = 0; datacount < dbData.size(); datacount++) {
				bean = (DataObject) dbData.get(datacount);
				
				if(!bis_type.equals(bean.getValue("bis_type"))&&!bis_type.equals("")){
					num = 1; //將不同逾放比率的序號歸零
					counts = ((counts/70)+1)*70;
				}
				
				rowNum = (counts%70)+3;
				cellNum = (counts/70)*4;
				
				for (int cellcount = 0; cellcount < 3; cellcount++) {
					if (cellcount == 0)
						sValue =  String.valueOf(num); //序號
					else if (cellcount == 1)
						sValue = (bean.getValue("bank_name") == null) ? "0" : String.valueOf(bean.getValue("bank_name")); //單位名稱
					else if (cellcount == 2)
						sValue = (bean.getValue("a05_bis") == null) ? "0" : String.valueOf(bean.getValue("a05_bis")); //逾放比率
						
					row = sheet.createRow(rowNum);
					cell = row.createCell((short) (cellcount+cellNum));
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs_center);
					cell.setCellValue(sValue);
				}
				
				bis_type = (bean.getValue("bis_type") == null) ? "" : String.valueOf(bean.getValue("bis_type"));
				counts++;
				num++;
			}
			
			//(七)全體農漁會信用部93年1月與○年○月經營情形比較表(A4直印,縮放比:75%)
			sheet = wb.getSheetAt(6);
			if (sheet == null) {
				System.out.println("open sheet7 失敗");
			} else {
				System.out.println("open sheet7 成功");
			}
			ps = sheet.getPrintSetup(); //取得設定
			
			// 設定頁面符合列印大小
			sheet.setAutobreaks(false);
			ps.setScale((short)75); //列印縮放百分比
			ps.setPaperSize((short)9); //設定紙張大小 A4
			ps.setLandscape(false);//設定橫印
			
			sql.setLength(0);
			paramList.clear();
			
			//資料SQL
			sql.append(" select round(field_ASSETS/?,0) as field_assets, "); //資產
			paramList.add(unit);
			sql.append("          round(field_DEBT/?,0) as field_debt,  "); //負債
			paramList.add(unit);
			sql.append("          round(field_DEBIT/?,0) as field_debit,  "); //存款
			paramList.add(unit);
			sql.append("          round(field_CREDIT/?,0) as field_credit, "); //放款
			paramList.add(unit);
			sql.append("          round(field_CREDIT/?,0) - round(loan_amt/?,0) as field_credit_normal, "); //放款-一般放款
			paramList.add(unit);
			paramList.add(unit);
			sql.append("          round(loan_amt/?,0) as loan_amt, "); //放款-專案農貸
			paramList.add(unit);
			sql.append("          round(field_OVER/?,0) as field_over, "); //逾放金額
			paramList.add(unit);
			sql.append("          round(field_OVER/?,0) - round(loan_over_amt/?,0) as field_over_normal, "); //逾放金額 -一般放款
			paramList.add(unit);
			paramList.add(unit);
			sql.append("          round(loan_over_amt/?,0) as loan_over_amt, "); //逾放金額-專案農貸
			paramList.add(unit);
			sql.append("          field_over_rate, "); //逾放比率
			sql.append("          decode((field_CREDIT-loan_amt),0,0,round(((field_OVER-loan_over_amt)/(field_CREDIT-loan_amt))*100 ,2)) as field_over_rate_normal, "); //逾放比率-一般放款
			sql.append("          decode(loan_amt,0,0,round((loan_over_amt/loan_amt)*100 ,2)) as field_over_rate_loan, "); //逾放比率-專案農貸
			sql.append("          round(field_NET/?,0)  as field_net , "); //淨值
			paramList.add(unit);
			sql.append("          round(field_320300 /?,0)  as field_320300, "); //稅前純益
			paramList.add(unit);
			sql.append("          round(field_BACKUP /?,0) as field_backup, "); //備抵呆帳
			paramList.add(unit);
			sql.append("          round(field_120700 /?,0) as field_120700, "); //內部融資
			paramList.add(unit);
			sql.append("          field_assets_rate, "); //資產報酬率
			sql.append("          field_net_rate, "); //淨值報酬率
			sql.append("          field_backup_credit_rate, "); //備呆占放款比率=備抵呆帳/放款
			sql.append("          field_backup_over_rate,  "); //備呆占逾放比率=備抵呆帳/逾期放款
			sql.append("          bank_cnt, "); //受輔導信用部家數
			sql.append("          loan_cnt  "); //專案農貸受益戶數
			sql.append(" from ( ");
			sql.append(" select round(sum(decode(a01.acc_code,'field_assets',a01.amt,0)) /1,0) as field_ASSETS, "); //資產
			sql.append("        round(sum(decode(a01.acc_code,'field_debt',a01.amt,0)) /1,0) as field_DEBT,   "); //負債
			sql.append("        round(sum(decode(a01.acc_code,'field_debit',a01.amt,0)) /1,0) as field_DEBIT, "); //存款
			sql.append("        round(sum(decode(a01.acc_code,'field_credit',a01.amt,0)) /1,0) as field_CREDIT,   "); //放款
			sql.append("        round(sum(decode(a01.acc_code,'field_over',a01.amt,0)) /1,0) as field_OVER, "); //逾放金額
			sql.append("        sum(decode(a01.acc_code,'field_over_rate',a01.amt,0)) as field_OVER_RATE, "); //逾放比率
			sql.append("        round(sum(decode(a01.acc_code,'field_net',a01.amt,0)) /1,0) as field_NET , "); //淨值
			sql.append("        round(sum(decode(a01.acc_code,'field_320300',a01.amt,0)) /1,0)  as field_320300, "); //稅前純益
			sql.append("        round(sum(decode(a01.acc_code,'field_backup',a01.amt,0)) /1,0) as field_BACKUP, "); //備抵呆帳
			sql.append("        round(sum(decode(a01.acc_code,'field_120700',a01.amt,0)) /1,0) as field_120700, "); //內部融資
			sql.append(" 	   decode(sum(decode(a01.acc_code,'field_assets',a01.amt,0)),0,0,round((sum(decode(a01.acc_code,'field_320300',a01.amt,0))/sum(decode(a01.acc_code,'field_assets',a01.amt,0)))*100 ,2)) as field_ASSETS_RATE, "); //資產報酬率
			sql.append(" 	   decode(sum(decode(a01.acc_code,'field_net',a01.amt,0)),0,0,round((sum(decode(a01.acc_code,'field_320300',a01.amt,0))/sum(decode(a01.acc_code,'field_net',a01.amt,0)))*100 ,2)) as field_NET_RATE, "); //淨值報酬率
			sql.append("        sum(decode(a01.acc_code,'field_backup_credit_rate',a01.amt,0))  as field_BACKUP_CREDIT_RATE, "); //備呆占放款比率=備抵呆帳/放款
			sql.append("        sum(decode(a01.acc_code,'field_backup_over_rate',a01.amt,0)) as field_BACKUP_OVER_RATE  "); //備呆占逾放比率=備抵呆帳/逾期放款
			sql.append(" from (select * from a01_operation_month where m_year=? and m_month=? and bank_type in ('ALL') and hsien_id=' ')a01 ");
			paramList.add(s_year);
			paramList.add(s_month);
			sql.append(" )a01, (select * from rpt_business where m_year=? and m_month=? )rpt_business ");
			paramList.add(s_year);
			paramList.add(s_month);
			
			dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "field_assets,field_debt,field_debit,field_credit,field_credit_normal,loan_amt,field_over,field_over_normal,loan_over_amt,field_over_rate,field_over_rate_normal,field_over_rate_loan,field_net,field_320300,field_backup,field_120700,field_assets_rate,field_net_rate,field_backup_credit_rate,field_backup_over_rate,bank_cnt,loan_cnt");
			System.out.println("7.dbData.size()="+dbData.size());
			
			row = sheet.getRow(0);
			cell = row.getCell((short) 0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("附件7、全體農漁會信用部93年1月與"+s_year+"年"+s_month+"月經營情形比較表");
			
			row = sheet.getRow(1);
			cell = row.getCell((short) 3);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("單位：" + unitName+"、百分點");
			
			row = sheet.getRow(2);
			cell = row.getCell((short) 2);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(s_year+"年"+s_month+"月底");
			
			// 列印明細/合計
			for (int datacount = 0; datacount < dbData.size(); datacount++) {
				bean = (DataObject) dbData.get(datacount);
				
				for (int rowcount = 3 ; rowcount < 25; rowcount++) {
					String sValue2 = "";
					row = sheet.getRow(rowcount);
					
					if (rowcount == 3){
						sValue = (bean.getValue("field_assets") == null) ? "0" : String.valueOf(bean.getValue("field_assets"));//資產
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = Utility.setCommaFormat(sValue);
						sValue2 = Utility.setCommaFormat(sValue2);
					}else if (rowcount == 4){
						sValue = (bean.getValue("field_debt") == null) ? "0" : String.valueOf(bean.getValue("field_debt"));//負債
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = Utility.setCommaFormat(sValue);
						sValue2 = Utility.setCommaFormat(sValue2);
					}else if (rowcount == 5){
						sValue = (bean.getValue("field_debit") == null) ? "0" : String.valueOf(bean.getValue("field_debit"));//存款
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = Utility.setCommaFormat(sValue);
						sValue2 = Utility.setCommaFormat(sValue2);
					}else if (rowcount == 6){
						sValue = (bean.getValue("field_credit") == null) ? "0" : String.valueOf(bean.getValue("field_credit"));//放款
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = Utility.setCommaFormat(sValue);
						sValue2 = Utility.setCommaFormat(sValue2);
					}else if (rowcount == 7){
						sValue = (bean.getValue("field_credit_normal") == null) ? "0" : String.valueOf(bean.getValue("field_credit_normal"));//放款-一般放款
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = Utility.setCommaFormat(sValue);
						sValue2 = Utility.setCommaFormat(sValue2);
					}else if (rowcount == 8){
						sValue = (bean.getValue("loan_amt") == null) ? "0" : String.valueOf(bean.getValue("loan_amt"));//放款-專案農貸
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = Utility.setCommaFormat(sValue);
						sValue2 = Utility.setCommaFormat(sValue2);
					}else if (rowcount == 9){
						sValue = (bean.getValue("field_over") == null) ? "0" : String.valueOf(bean.getValue("field_over"));//逾放金額
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = Utility.setCommaFormat(sValue);
						sValue2 = Utility.setCommaFormat(sValue2);
					}else if (rowcount == 10){
						sValue = (bean.getValue("field_over_normal") == null) ? "0" : String.valueOf(bean.getValue("field_over_normal"));//逾放金額 -一般放款
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = Utility.setCommaFormat(sValue);
						sValue2 = Utility.setCommaFormat(sValue2);
					}else if (rowcount == 11){
						sValue = (bean.getValue("loan_over_amt") == null) ? "0" : String.valueOf(bean.getValue("loan_over_amt"));//逾放金額-專案農貸
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = Utility.setCommaFormat(sValue);
						sValue2 = Utility.setCommaFormat(sValue2);
					}else if (rowcount == 12){
						sValue =  (bean.getValue("field_over_rate") == null) ? "0" : String.valueOf(bean.getValue("field_over_rate"));//逾放比率
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,true);
						sValue = String.format("%s%%",sValue);
						sValue2 = String.format("%s 個百分點",sValue2);						
					}else if (rowcount == 13){
						sValue =  (bean.getValue("field_over_rate_normal") == null) ? "0" : String.valueOf(bean.getValue("field_over_rate_normal"));//逾放比率-一般放款
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,true);
						sValue = String.format("%s%%",sValue);
						sValue2 = String.format("%s 個百分點",sValue2);						
					}else if (rowcount == 14){
						sValue =  (bean.getValue("field_over_rate_loan") == null) ? "0" : String.valueOf(bean.getValue("field_over_rate_loan"));//逾放比率-專案農貸
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,true);
						sValue = String.format("%s%%",sValue);
						sValue2 = String.format("%s 個百分點",sValue2);						
					}else if (rowcount == 15){
						sValue = (bean.getValue("field_net") == null) ? "0" : String.valueOf(bean.getValue("field_net"));//淨值
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = Utility.setCommaFormat(sValue);
						sValue2 = Utility.setCommaFormat(sValue2);
					}else if (rowcount == 16){
						sValue = (bean.getValue("field_320300") == null) ? "0" : String.valueOf(bean.getValue("field_320300"));//稅前純益
						sValue = Utility.setCommaFormat(sValue);
					}else if (rowcount == 17){
						sValue = (bean.getValue("field_backup") == null) ? "0" : String.valueOf(bean.getValue("field_backup"));//備抵呆帳
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = Utility.setCommaFormat(sValue);
						sValue2 = Utility.setCommaFormat(sValue2);
					}else if (rowcount == 18){
						sValue = (bean.getValue("field_120700") == null) ? "0" : String.valueOf(bean.getValue("field_120700"));//內部融資
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = Utility.setCommaFormat(sValue);
						sValue2 = Utility.setCommaFormat(sValue2);
					}else if (rowcount == 19){
						sValue =  (bean.getValue("field_assets_rate") == null) ? "0" : String.valueOf(bean.getValue("field_assets_rate"));//資產報酬率
						sValue = String.format("%s%%",sValue);
					}else if (rowcount == 20){
						sValue =  (bean.getValue("field_net_rate") == null) ? "0" : String.valueOf(bean.getValue("field_net_rate"));//淨值報酬率
						sValue = String.format("%s%%",sValue);
					}else if (rowcount == 21){
						sValue =  (bean.getValue("field_backup_credit_rate") == null) ? "0" : String.valueOf(bean.getValue("field_backup_credit_rate"));//備呆占放款比率=備抵呆帳/放款
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,true);
						sValue = String.format("%s%%",sValue);
						sValue2 = String.format("%s 個百分點",sValue2);						
					}else if (rowcount == 22){
						sValue =  (bean.getValue("field_backup_over_rate") == null) ? "0" : String.valueOf(bean.getValue("field_backup_over_rate"));//備呆占逾放比率=備抵呆帳/逾期放款
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,true);
						sValue = String.format("%s%%",sValue);						
						sValue2 = String.format("%s 個百分點",sValue2);						
					}else if (rowcount == 23){
						sValue = (bean.getValue("bank_cnt") == null) ? "0" : String.valueOf(bean.getValue("bank_cnt"));//受輔導信用部家數
						sValue2 = cellSubMath(row.getCell((short) 1),sValue,false);
						sValue = String.format("%s 家",sValue);
						sValue2 = String.format("%s 家",sValue2);
					}else if (rowcount == 24){
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
			
			//(八)全體農漁會信用部85年至○年經營情形表(A3橫印,縮放比:84%)
			sheet = wb.getSheetAt(7);
			if (sheet == null) {
				System.out.println("open sheet8 失敗");
			} else {
				System.out.println("open sheet8 成功");
			}
			ps = sheet.getPrintSetup(); //取得設定
			
			// 設定頁面符合列印大小
			sheet.setAutobreaks(false);
			ps.setScale((short)84); //列印縮放百分比
			ps.setPaperSize((short)8); //設定紙張大小 A3
			ps.setLandscape(true);//設定橫印
			
			sql.setLength(0);
			paramList.clear();
			
			row = sheet.getRow(0);
			cell = row.getCell((short) 0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("附件8、全體農漁會信用部85年至"+s_year+"年經營情形表");
			
			dbData = getA01Operation(unit,s_year);
			System.out.println("8.dbData.size()="+dbData.size());
			//counts = 10; //從95年底開始//106.12.19 fix
			counts = 7; //從107年底開始//106.12.19 add
			//counts = 6; //從106年底開始//107.01.24 add
			HSSFCellStyle cs2 = sheet.getRow((short)2).getCell((short)1).getCellStyle();
			HSSFCellStyle cs3 = sheet.getRow((short)3).getCell((short)1).getCellStyle();
			
			// 列印明細/合計
			for (int datacount = 0; datacount < dbData.size(); datacount++) {
				bean = (DataObject) dbData.get(datacount);
				
				cellNum = (counts%15)+1;
				//rowNum = (counts/15)*19;//106.12.19 原從95年開始
				rowNum = ((counts/15)+1)*19;//106.12.19 add 從107年開始				
				System.out.println("cellNum="+cellNum+":rowNum="+rowNum);
				
				if(cellNum == 1){
					for(int i = 2 ; i < 19 ; i++){
						row = sheet.createRow(rowNum+i);
						cell = row.createCell((short) 0);
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cs3.setWrapText(true);
						if(i ==2) cell.setCellStyle(cs2);
						else cell.setCellStyle(cs3);
						
						if(i == 3) cell.setCellValue("資產");
						if(i == 4) cell.setCellValue("負債");
						if(i == 5) cell.setCellValue("存款");
						if(i == 6) cell.setCellValue("放款");
						if(i == 7) cell.setCellValue("淨值");
						if(i == 8) cell.setCellValue("稅前純益");
						if(i == 9) cell.setCellValue("逾放金額");
						if(i == 10) cell.setCellValue("逾放比率");
						if(i == 11) cell.setCellValue("備抵呆帳");
						if(i == 12) cell.setCellValue("內部融資");
						if(i == 13) cell.setCellValue("資產報酬率");
						if(i == 14) cell.setCellValue("淨值報酬率");
						if(i == 15) cell.setCellValue("備抵呆帳占放款比率");
						if(i == 16) cell.setCellValue("備抵呆帳占逾放比率");
						if(i == 17) cell.setCellValue("存款年成長率");
						if(i == 18) cell.setCellValue("放款年成長率");
					}
				}
				
				row = sheet.createRow(rowNum+2);
				cell = row.createCell((short) cellNum);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs2);
				//cell.setCellValue((datacount+95)+"年底");//106.12.19 fix
				cell.setCellValue((datacount+107)+"年底");//106.12.19 add從107年開始
				//cell.setCellValue((datacount+106)+"年底");//107.01.24 add從106年開始
				
				for (int rowcount = 3; rowcount < 19; rowcount++) {
					if (rowcount == 3)
						sValue = (bean.getValue("field_assets") == null) ? "0" : String.valueOf(bean.getValue("field_assets"));//資產
					else if (rowcount == 4)
						sValue = (bean.getValue("field_debt") == null) ? "0" : String.valueOf(bean.getValue("field_debt"));//負債
					else if (rowcount == 5)
						sValue = (bean.getValue("field_debit") == null) ? "0" : String.valueOf(bean.getValue("field_debit"));//存款
					else if (rowcount == 6)
						sValue = (bean.getValue("field_credit") == null) ? "0" : String.valueOf(bean.getValue("field_credit"));//放款
					else if (rowcount == 7)
						sValue = (bean.getValue("field_net") == null) ? "0" : String.valueOf(bean.getValue("field_net"));//淨值
					else if (rowcount == 8)
						sValue = (bean.getValue("field_320300") == null) ? "0" : String.valueOf(bean.getValue("field_320300"));//稅前純益
					else if (rowcount == 9)
						sValue = (bean.getValue("field_over") == null) ? "0" : String.valueOf(bean.getValue("field_over"));//逾放金額
					else if (rowcount == 10)
						sValue = String.format("%s%%",(bean.getValue("field_over_rate") == null) ? "0" : String.valueOf(bean.getValue("field_over_rate")));//逾放比率
					else if (rowcount == 11)
						sValue = (bean.getValue("field_backup") == null) ? "0" : String.valueOf(bean.getValue("field_backup"));//備抵呆帳
					else if (rowcount == 12)
						sValue = (bean.getValue("field_120700") == null) ? "0" : String.valueOf(bean.getValue("field_120700"));//內部融資
					else if (rowcount == 13)
						sValue = String.format("%s%%",(bean.getValue("field_assets_rate") == null) ? "0" : String.valueOf(bean.getValue("field_assets_rate")));//資產報酬率
					else if (rowcount == 14)
						sValue = String.format("%s%%",(bean.getValue("field_net_rate") == null) ? "0" : String.valueOf(bean.getValue("field_net_rate")));//淨值報酬率
					else if (rowcount == 15)
						sValue = String.format("%s%%",(bean.getValue("field_backup_credit_rate") == null) ? "0" : String.valueOf(bean.getValue("field_backup_credit_rate")));//備呆占放款比率=備抵呆帳/放款
					else if (rowcount == 16)
						sValue = String.format("%s%%",(bean.getValue("field_backup_over_rate") == null) ? "0" : String.valueOf(bean.getValue("field_backup_over_rate")));//備呆占逾放比率=備抵呆帳/逾期放款
					else if (rowcount == 17)
						sValue = String.format("%s%%",(bean.getValue("field_debit_rate") == null) ? "0" : String.valueOf(bean.getValue("field_debit_rate")));//存款成長率
					else if (rowcount == 18)
						sValue = String.format("%s%%",(bean.getValue("field_credit_rate") == null) ? "0" : String.valueOf(bean.getValue("field_credit_rate")));//放款成長率
					
					row = sheet.createRow(rowNum+rowcount);
					cell = row.createCell((short) cellNum);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs3);
					cell.setCellValue(sValue.indexOf("%")>=0?sValue:Utility.setCommaFormat(sValue));
					
				} 
				counts++;
			}
			
			//(九)重設及新設信用部○年○月營運概況表(A3橫印,縮放比:115%)
			sheet = wb.getSheetAt(8);
			if (sheet == null) {
				System.out.println("open sheet9 失敗");
			} else {
				System.out.println("open sheet9 成功");
			}
			ps = sheet.getPrintSetup(); //取得設定
			
			// 設定頁面符合列印大小
			sheet.setAutobreaks(false);
			ps.setScale((short)115); //列印縮放百分比
			ps.setPaperSize((short)8); //設定紙張大小 A3
			ps.setLandscape(true);//設定橫印
			
			sql.setLength(0);
			paramList.clear();
			
			row = sheet.getRow(0);
			cell = row.getCell((short) 0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue("附件9、重設及新設信用部"+s_year+"年"+s_month+"月營運概況表");
			
			sql.append(" select bn01_reset.bank_no, ");
			sql.append("        bank_name, "); //信用部
			sql.append("        to_number(nvl(output_order, '998')) as output_order, ");
			sql.append("        F_TRANSCHINESEDATE(setup_date) setup_date, "); //核准設立日期
			sql.append("        F_TRANSCHINESEDATE(start_date) start_date, "); //開始營運日期
			sql.append("        F_TRANSCHINESEDATE(bn01_reset.add_date) add_date, "); //加入存款保險日期
			sql.append("        round(sum(decode(acc_code,'field_debit',amt,0)) /? ,0) as field_debit, "); //存款
			paramList.add(1000);
			sql.append("        round(sum(decode(acc_code,'field_credit',amt,0)) /? ,0) as field_credit, "); //放款
			paramList.add(1000);
			sql.append("        sum(decode(acc_code,'field_dc_rate',amt,0)) as field_dc_rate, "); //存放比率
			sql.append("        round(sum(decode(acc_code,'field_over',amt,0)) /? ,0) as field_over, "); //逾放金額
			paramList.add(1000);
			sql.append("        sum(decode(acc_code,'field_over_rate',amt,0)) as field_over_rate, "); //逾放比率
			sql.append("        round(sum(decode(acc_code,'field_320300',amt,0)) /? ,0) as field_320300, "); //本期損益
			paramList.add(1000);
			sql.append("        round(sum(decode(acc_code,'field_310000',amt,0)) /? ,0) as field_310000, "); //事業資金及公積
			paramList.add(1000);
			sql.append("        round(sum(decode(acc_code,'field_net',amt,0)) /? ,0) as field_net , "); //淨值
			paramList.add(1000);
			sql.append("        round(sum(decode(acc_code,'field_backup',amt,0)) /? ,0) as field_backup, "); //備抵呆帳
			paramList.add(1000);
			sql.append("        round(sum(decode(acc_code,'field_noassure',amt,0)) /? ,0) as field_noassure, "); //無擔保放款
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
			sql.append("        '' as setup_date,'' as start_date,'' as add_date, ");
			sql.append("        round(sum(decode(acc_code,'field_debit',amt,0)) /? ,0) as field_debit, "); //存款
			paramList.add(1000);
			sql.append("        round(sum(decode(acc_code,'field_credit',amt,0)) /? ,0) as field_credit, "); //放款
			paramList.add(1000);
			sql.append("        0 as field_dc_rate, "); //存放比率
			sql.append("        round(sum(decode(acc_code,'field_over',amt,0)) /? ,0) as field_over, "); //逾放金額
			paramList.add(1000);
			sql.append("        decode(sum(decode(acc_code,'field_credit',amt,0)),0,0,round(sum(decode(acc_code,'field_over',amt,0)) / sum(decode(acc_code,'field_credit',amt,0)) *100 ,2)) as field_over_rate, "); //逾放比率
			sql.append("        round(sum(decode(acc_code,'field_320300',amt,0)) /? ,0) as field_320300, "); //本期損益
			paramList.add(1000);
			sql.append("        round(sum(decode(acc_code,'field_310000',amt,0)) /? ,0) as field_310000, "); //事業資金及公積
			paramList.add(1000);
			sql.append("        round(sum(decode(acc_code,'field_net',amt,0)) /? ,0) as field_net , "); //淨值
			paramList.add(1000);
			sql.append("        round(sum(decode(acc_code,'field_backup',amt,0)) /? ,0) as field_backup, "); //備抵呆帳
			paramList.add(1000);
			sql.append("        round(sum(decode(acc_code,'field_noassure',amt,0)) /? ,0) as field_noassure,  "); //無擔保放款
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
			
			dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "bank_no,bank_name,output_order,setup_date,start_date,add_date,field_debit,field_credit,field_dc_rate,field_over,field_over_rate,field_320300,field_310000,field_net,field_backup,field_noassure,field_captial_rate");
			System.out.println("9.dbData.size()="+dbData.size());
			
			num = 1; //序號
			int field_debit=0,field_credit=0,field_over=0,field_320300=0,field_310000=0,field_net=0,field_backup=0,field_noassure=0;
			// 列印明細/合計
			for (int datacount = 0; datacount < dbData.size(); datacount++) {
				bean = (DataObject) dbData.get(datacount);
				
				rowNum = datacount+3;
				
				for (int cellcount = 0; cellcount < 16; cellcount++) {
					sValue = "";
					if (cellcount == 0 && datacount != dbData.size()-1)
						sValue = String.valueOf(num+datacount);//序號
					else if (cellcount == 1)
						sValue = (bean.getValue("bank_name") == null) ? "0" : String.valueOf(bean.getValue("bank_name"));//信用部
					else if (cellcount == 2 && datacount != dbData.size()-1)
						sValue = (bean.getValue("setup_date") == null) ? "0" : String.valueOf(bean.getValue("setup_date"));//核准設立日期
					else if (cellcount == 3 && datacount != dbData.size()-1)
						sValue = (bean.getValue("start_date") == null) ? "0" : String.valueOf(bean.getValue("start_date"));//開始營運日期
					else if (cellcount == 4 && datacount != dbData.size()-1)
						sValue = (bean.getValue("add_date") == null) ? "0" : String.valueOf(bean.getValue("add_date"));//加入存款保險日期
					else if (cellcount == 5){
						sValue = (bean.getValue("field_debit") == null) ? "0" : String.valueOf(bean.getValue("field_debit"));//存款
						if(!((String)bean.getValue("bank_name")).equals("總計")) field_debit+=Integer.parseInt((bean.getValue("field_debit") == null) ? "0" : String.valueOf(bean.getValue("field_debit")));
						sValue = Utility.setCommaFormat(sValue);						
					}
					else if (cellcount == 6){
						sValue = (bean.getValue("field_credit") == null) ? "0" : String.valueOf(bean.getValue("field_credit"));//放款
						if(!((String)bean.getValue("bank_name")).equals("總計")) field_credit+=Integer.parseInt((bean.getValue("field_credit") == null) ? "0" : String.valueOf(bean.getValue("field_credit")));
						sValue = Utility.setCommaFormat(sValue);
						
					}
					else if (cellcount == 7 && datacount != dbData.size()-1)
						sValue = (bean.getValue("field_dc_rate") == null) ? "0" : String.valueOf(bean.getValue("field_dc_rate"));//存放比率
					else if (cellcount == 8){
						sValue = (bean.getValue("field_over") == null) ? "0" : String.valueOf(bean.getValue("field_over"));//逾放金額
						if(!((String)bean.getValue("bank_name")).equals("總計")) field_over+=Integer.parseInt((bean.getValue("field_over") == null) ? "0" : String.valueOf(bean.getValue("field_over")));
						sValue = Utility.setCommaFormat(sValue);
						
					}
					else if (cellcount == 9)
						sValue = (bean.getValue("field_over_rate") == null) ? "0" : String.valueOf(bean.getValue("field_over_rate"));//逾放比率
					else if (cellcount == 10){
						sValue = (bean.getValue("field_320300") == null) ? "0" : String.valueOf(bean.getValue("field_320300"));//本期損益						
						if(!((String)bean.getValue("bank_name")).equals("總計")) field_320300+=Integer.parseInt((bean.getValue("field_320300") == null) ? "0" : String.valueOf(bean.getValue("field_320300")));
						sValue = Utility.setCommaFormat(sValue);
						
					}
					else if (cellcount == 11){
						sValue = (bean.getValue("field_310000") == null) ? "0" : String.valueOf(bean.getValue("field_310000"));//事業資金及公積
						if(!((String)bean.getValue("bank_name")).equals("總計")) field_310000+=Integer.parseInt((bean.getValue("field_310000") == null) ? "0" : String.valueOf(bean.getValue("field_310000")));
						sValue = Utility.setCommaFormat(sValue);
						
					}
					else if (cellcount == 12){
						sValue = (bean.getValue("field_net") == null) ? "0" : String.valueOf(bean.getValue("field_net"));//淨值
						if(!((String)bean.getValue("bank_name")).equals("總計")) field_net+=Integer.parseInt((bean.getValue("field_net") == null) ? "0" : String.valueOf(bean.getValue("field_net")));
						sValue = Utility.setCommaFormat(sValue);
						
					}
					else if (cellcount == 13){
						sValue = (bean.getValue("field_backup") == null) ? "0" : String.valueOf(bean.getValue("field_backup"));//備抵呆帳
						if(!((String)bean.getValue("bank_name")).equals("總計")) field_backup+=Integer.parseInt((bean.getValue("field_backup") == null) ? "0" : String.valueOf(bean.getValue("field_backup")));
						sValue = Utility.setCommaFormat(sValue);
					}
					else if (cellcount == 14){
						sValue = (bean.getValue("field_noassure") == null) ? "0" : String.valueOf(bean.getValue("field_noassure"));//無擔保放款
						if(!((String)bean.getValue("bank_name")).equals("總計")) field_noassure+=Integer.parseInt((bean.getValue("field_noassure") == null) ? "0" : String.valueOf(bean.getValue("field_noassure")));
						sValue = Utility.setCommaFormat(sValue);						
					}
					else if (cellcount == 15 && datacount != dbData.size()-1)
						sValue = (bean.getValue("field_captial_rate") == null) ? "0" : String.valueOf(bean.getValue("field_captial_rate"));//淨值佔風險性資產比率
					
					row = sheet.createRow(rowNum);
					cell = row.createCell((short) cellcount);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs_center);
					cell.setCellValue(sValue);
				}
				
			}
			rowNum = 3+dbData.size()-1;
			//System.out.println("9.合計列="+rowNum);			
			row = sheet.getRow(rowNum);			
            //合計的部份.已原本各機構.先除以單位後.再加總 106.12.06 fix 
            for (int cellcount = 5; cellcount < 15; cellcount++) {
                if(cellcount != 7 && cellcount != 9){
                    cell = row.createCell((short) cellcount);
                    cell.setEncoding(HSSFCell.ENCODING_UTF_16);
                    cell.setCellStyle(cs_center);
                }
                if (cellcount == 5){                     
                    sValue = Utility.setCommaFormat(String.valueOf(field_debit));                  
                }else if (cellcount == 6){
                    sValue = Utility.setCommaFormat(String.valueOf(field_credit));
                }else if (cellcount == 8){    
                    sValue = Utility.setCommaFormat(String.valueOf(field_over));
                }else if (cellcount == 10){
                    sValue = Utility.setCommaFormat(String.valueOf(field_320300));
                }else if (cellcount == 11){   
                    sValue = Utility.setCommaFormat(String.valueOf(field_310000));
                }else if (cellcount == 12){   
                    sValue = Utility.setCommaFormat(String.valueOf(field_net));
                }else if (cellcount == 13){   
                    sValue = Utility.setCommaFormat(String.valueOf(field_backup));
                }else if (cellcount == 14){  
                    sValue = Utility.setCommaFormat(String.valueOf(field_noassure));
                }
                if(cellcount != 7 && cellcount != 9){
                    cell.setCellValue(sValue); 
                }
            }    
					
			// 建表結束--------------------------------------
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
	
	private static void setTitle(HSSFSheet sourceSheet, HSSFCellStyle cs, String over_type, int num, int typecounts, int cellNum, int rowNum, int order){
		
		HSSFSheet sheet = sourceSheet;
		HSSFRow row = null;  //宣告一列
		HSSFCell cell = null;//宣告一個儲存格
		
		String title = "";
		if(order == 5){
			if("TYPE1".equals(over_type)){
				title = "逾放比率0~1%(不含1%)";
				if(num == 0) title +="  共"+typecounts+"家";
			}else if("TYPE2".equals(over_type)){
				title = "逾放比率1%-15%(不含15%)";
				if(num == 0) title +="  共"+typecounts+"家";
			}else if("TYPE3".equals(over_type)){
				title = "逾放比率15%-25%(不含25%)";
				if(num == 0) title +="  共"+typecounts+"家";
			}else if("TYPE4".equals(over_type)){
				title = "逾放比率25%以上";
				if(num == 0) title +="  共"+typecounts+"家";
			}
		}else if(order == 6){
			if("TYPE1".equals(over_type)){
				title = "未達8%";
				if(num == 0) title +="  共"+typecounts+"家";
			}else if("TYPE2".equals(over_type)){
				title = "8%~12%  (不含12%) ";
				if(num == 0) title ="8%~12%   共"+typecounts+"家 (不含12%)";
			}else if("TYPE3".equals(over_type)){
				title = "12%以上";
				if(num == 0) title +="  共"+typecounts+"家";
			}
		}
		
		sheet.setColumnWidth((short)cellNum,(short)(256*11));
		sheet.setColumnWidth((short)(cellNum+1),(short)(256*31));
		sheet.setColumnWidth((short)(cellNum+2),(short)(256*11));
		sheet.setColumnWidth((short)(cellNum+3),(short)(256*2));
		
		row = sheet.createRow(rowNum-2);
		cell = row.createCell((short)cellNum);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(cs);
		cell.setCellValue(title);
		cell = row.createCell((short)(cellNum+1));
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(cs);
		cell = row.createCell((short)(cellNum+2));
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(cs);
		
		//合併儲存格，參數分別為起始行、起始列、結束行、結束列 
		sheet.addMergedRegion(new Region((short) (rowNum-2), (short) cellNum, (short)  (rowNum-2), (short)( cellNum+2)));//合併儲存格
				
		row = sheet.createRow(rowNum-1);
		cell = row.createCell((short)cellNum);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(cs);
		cell.setCellValue("序號");
		
		cell = row.createCell((short)(cellNum+1));
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(cs);
		cell.setCellValue("單位名稱");
		
		cell = row.createCell((short)(cellNum+2));
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellStyle(cs);
		if(order == 5){
			cell.setCellValue("逾放比率");
		}else if(order == 6){
			cell.setCellValue("BIS");
		}

	}
	
	private static List getA01Operation(String unit, String s_year, String s_month){
		List dbData = new ArrayList();
		StringBuffer sql = new StringBuffer ();
		ArrayList paramList = new ArrayList();
		String year = s_year;
		String l_month = "1";
		
		for(int i = 1; i <= Integer.parseInt(s_month); i++){
		//for(int i = 10; i <= Integer.parseInt(s_month); i++){//106.12.20 fix 暫時調整為從10月份開始列印
			sql.setLength(0);
			paramList.clear();
			
			if(i == 1){
				l_month = "12";
			}else{
				l_month = String.valueOf(i-1);
			}
			
			//資料SQL
			sql.append(" select round(field_assets /?,0) as field_assets, "); //資產
			paramList.add(unit);
			sql.append("round(field_debt /?,0) as field_debt, "); //負債
			paramList.add(unit);
			sql.append("round(a01.field_debit /?,0) as field_debit, "); //存款
			paramList.add(unit);
			sql.append("round(a01.field_credit /?,0) as field_credit,   "); //放款
			paramList.add(unit);
			sql.append("round(field_net /?,0) as  field_net , "); //淨值
			paramList.add(unit);
			sql.append("round(field_320300 /?,0)  as field_320300, "); //稅前純益
			paramList.add(unit);
			sql.append("round(field_over /?,0)  as field_over, "); //逾放金額
			paramList.add(unit);
			sql.append("field_over_rate, ") //逾放比率
				.append("round(field_backup /?,0)  as  field_backup, "); //備抵呆帳
			paramList.add(unit);
			sql.append("round(field_120700 /?,0)  as   field_120700, "); //內部融資
			paramList.add(unit);
			sql.append("field_assets_rate, ") //資產報酬率
				.append("field_net_rate, ") //淨值報酬率
				.append("field_backup_credit_rate, ") //備呆占放款比率=備抵呆帳/放款
				.append("field_backup_over_rate, ") //備呆占逾放比率=備抵呆帳/逾期放款
				.append("decode(a01_last.field_debit,0,0,round((a01.field_debit-a01_last.field_debit)/ a01_last.field_debit*100 ,2)) as field_debit_rate, ") //存款成長率
				.append("decode(a01_last.field_credit,0,0,round((a01.field_credit-a01_last.field_credit)/ a01_last.field_credit*100 ,2)) as field_credit_rate ") //放款成長率       
				.append(" from ( ")
				.append(" select round(sum(decode(a01.acc_code,'field_assets',a01.amt,0)) /1,0) as field_assets, ") //資產
				.append("round(sum(decode(a01.acc_code,'field_debt',a01.amt,0)) /1,0) as field_debt,   ") //負債
				.append("round(sum(decode(a01.acc_code,'field_debit',a01.amt,0)) /1,0) as field_debit, ") //存款        
				.append("round(sum(decode(a01.acc_code,'field_credit',a01.amt,0)) /1,0) as field_credit,   ") //放款    
				.append("round(sum(decode(a01.acc_code,'field_net',a01.amt,0)) /1,0) as field_net , ") //淨值
				.append("round(sum(decode(a01.acc_code,'field_320300',a01.amt,0)) /1,0)  as field_320300, ") //稅前純益
				.append("round(sum(decode(a01.acc_code,'field_over',a01.amt,0)) /1,0) as field_over, ") //逾放金額     
				.append("sum(decode(a01.acc_code,'field_over_rate',a01.amt,0)) as field_over_rate, ") //逾放比率
				.append("round(sum(decode(a01.acc_code,'field_backup',a01.amt,0)) /1,0) as field_backup, ") //備抵呆帳
				.append("round(sum(decode(a01.acc_code,'field_120700',a01.amt,0)) /1,0) as field_120700, ") //內部融資
				.append("decode(sum(decode(a01.acc_code,'field_assets',a01.amt,0)),0,0,round((sum(decode(a01.acc_code,'field_320300',a01.amt,0))/sum(decode(a01.acc_code,'field_assets',a01.amt,0)))*100 ,2)) as field_assets_rate, ") //資產報酬率
				.append("decode(sum(decode(a01.acc_code,'field_net',a01.amt,0)),0,0,round((sum(decode(a01.acc_code,'field_320300',a01.amt,0))/sum(decode(a01.acc_code,'field_net',a01.amt,0)))*100 ,2)) as field_net_rate, ") //淨值報酬率
				.append("sum(decode(a01.acc_code,'field_backup_credit_rate',a01.amt,0))  as field_backup_credit_rate, ") //備呆占放款比率=備抵呆帳/放款
				.append("sum(decode(a01.acc_code,'field_backup_over_rate',a01.amt,0)) as field_backup_over_rate ") //, ") //備呆占逾放比率=備抵呆帳/逾期放款
				.append(" from (select * from a01_operation_month where m_year=? and m_month=? and bank_type in ('ALL') and hsien_id=' ')a01 ");
			paramList.add(year);
			paramList.add(i);
			sql.append(")a01,  ") //本月份
				.append(" ( ") 
				.append(" select round(sum(decode(a01.acc_code,'field_debit',a01.amt,0)) /1,0) as field_debit, ") //存款        
				.append(" round(sum(decode(a01.acc_code,'field_credit',a01.amt,0)) /1,0) as field_credit   ") //放款          
				.append(" from (select * from a01_operation_month where m_year=? and m_month=?  and bank_type in ('ALL') and hsien_id=' ')a01 ");
			paramList.add(i == 1?String.valueOf(Integer.parseInt(s_year)-1):year);
			paramList.add(l_month);
			sql.append(" )a01_last  "); //上月份
			
			List Data = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "field_assets,field_debt,field_debit,field_credit,field_net,field_320300,field_over,field_over_rate,field_backup,field_120700,field_assets_rate,field_net_rate,field_backup_credit_rate,field_backup_over_rate,field_debit_rate,field_credit_rate");
			dbData.add(Data.get(0));
		}
		return dbData;
	}
	
	private static List getA01Operation(String unit, String s_year){
		List dbData = new ArrayList();
		StringBuffer sql = new StringBuffer ();
		ArrayList paramList = new ArrayList();
		//106.12.19 調整107年以後,才讀取db資料
		//for(int i = 106 ; i <= Integer.parseInt(s_year) ; i++){//107.01.24 add 從106年開始列印資料
		for(int i = 107 ; i <= Integer.parseInt(s_year) ; i++){//106.12.19 add 從107年開始列印資料
		//for(int i = 95 ; i <= Integer.parseInt(s_year) ; i++){//106.12.19 原95年開始列印		
			sql.setLength(0);
			paramList.clear();
			
			//資料SQL
			sql.append(" select round(field_ASSETS /? ,0) as field_assets, "); //資產
			paramList.add(unit);
			sql.append("           round(field_DEBT /? ,0) as field_debt, "); //負債
			paramList.add(unit);
			sql.append("           round(a01.field_DEBIT /? ,0) as field_debit, "); //存款
			paramList.add(unit);
			sql.append("           round(a01.field_CREDIT /? ,0) as field_credit,   "); //放款
			paramList.add(unit);
			sql.append("           round(field_NET /? ,0) as  field_net , "); //淨值
			paramList.add(unit);
			sql.append("           round(field_320300 /? ,0)  as field_320300, "); //稅前純益
			paramList.add(unit);
			sql.append("           round(field_OVER /? ,0)  as field_over, "); //逾放金額
			paramList.add(unit);
			sql.append("           field_over_rate, "); //逾放比率
			sql.append("           round(field_BACKUP /? ,0)  as  field_backup, "); //備抵呆帳
			paramList.add(unit);
			sql.append("           round(field_120700 /? ,0)  as   field_120700, "); //內部融資
			paramList.add(unit);
			sql.append("           field_assets_rate, "); //資產報酬率
			sql.append("           field_net_rate, "); //淨值報酬率
			sql.append("           field_backup_credit_rate, "); //備呆占放款比率=備抵呆帳/放款
			sql.append("           field_backup_over_rate, "); //備呆占逾放比率=備抵呆帳/逾期放款
			sql.append("           decode(a01_last.field_DEBIT,0,0,round((a01.field_DEBIT-a01_last.field_DEBIT)/ a01_last.field_DEBIT*100 ,2)) as field_debit_rate, "); //存款成長率
			sql.append("           decode(a01_last.field_CREDIT,0,0,round((a01.field_CREDIT-a01_last.field_CREDIT)/ a01_last.field_CREDIT*100 ,2)) as field_credit_rate "); //放款成長率
			sql.append(" from ( ");
			sql.append("  select round(sum(decode(a01.acc_code,'field_assets',a01.amt,0)) /1,0) as field_ASSETS, "); //資產
			sql.append("        round(sum(decode(a01.acc_code,'field_debt',a01.amt,0)) /1,0) as field_DEBT,   "); //負債
			sql.append("        round(sum(decode(a01.acc_code,'field_debit',a01.amt,0)) /1,0) as field_DEBIT, "); //存款
			sql.append("        round(sum(decode(a01.acc_code,'field_credit',a01.amt,0)) /1,0) as field_CREDIT,   "); //放款
			sql.append("        round(sum(decode(a01.acc_code,'field_net',a01.amt,0)) /1,0) as field_NET , "); //淨值
			sql.append("        round(sum(decode(a01.acc_code,'field_320300',a01.amt,0)) /1,0)  as field_320300, "); //稅前純益
			sql.append("        round(sum(decode(a01.acc_code,'field_over',a01.amt,0)) /1,0) as field_OVER, "); //逾放金額
			sql.append("        sum(decode(a01.acc_code,'field_over_rate',a01.amt,0)) as field_OVER_RATE, "); //逾放比率
			sql.append("        round(sum(decode(a01.acc_code,'field_backup',a01.amt,0)) /1,0) as field_BACKUP, "); //備抵呆帳
			sql.append("        round(sum(decode(a01.acc_code,'field_120700',a01.amt,0)) /1,0) as field_120700, "); //內部融資
			sql.append("        decode(sum(decode(a01.acc_code,'field_assets',a01.amt,0)),0,0,round((sum(decode(a01.acc_code,'field_320300',a01.amt,0))/sum(decode(a01.acc_code,'field_assets',a01.amt,0)))*100 ,2)) as field_ASSETS_RATE, "); //資產報酬率
			sql.append("        decode(sum(decode(a01.acc_code,'field_net',a01.amt,0)),0,0,round((sum(decode(a01.acc_code,'field_320300',a01.amt,0))/sum(decode(a01.acc_code,'field_net',a01.amt,0)))*100 ,2)) as field_NET_RATE, "); //淨值報酬率
			sql.append("        sum(decode(a01.acc_code,'field_backup_credit_rate',a01.amt,0))  as field_BACKUP_CREDIT_RATE, "); //備呆占放款比率=備抵呆帳/放款
			sql.append("        sum(decode(a01.acc_code,'field_backup_over_rate',a01.amt,0)) as field_BACKUP_OVER_RATE "); //, "); //備呆占逾放比率=備抵呆帳/逾期放款
			sql.append("  from (select * from a01_operation_month where m_year=? and m_month='12' and bank_type in ('ALL') and hsien_id=' ')a01 ");
			paramList.add(i);
			sql.append(" )a01,  "); //本月份
			sql.append(" (select round(sum(decode(a01.acc_code,'field_debit',a01.amt,0)) /1,0) as field_DEBIT, "); //存款
			sql.append("        round(sum(decode(a01.acc_code,'field_credit',a01.amt,0)) /1,0) as field_CREDIT   "); //放款
			sql.append("  from (select * from a01_operation_month where m_year=? and m_month='12'  and bank_type in ('ALL') and hsien_id=' ')a01 ");
			paramList.add(i-1);
			sql.append(" )a01_last ");
			
			List Data = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "field_assets,field_debt,field_debit,field_credit,field_net,field_320300,field_over,field_over_rate,field_backup,field_120700,field_assets_rate,field_net_rate,field_backup_credit_rate,field_backup_over_rate,field_debit_rate,field_credit_rate");
			dbData.add(Data.get(0));
		}
		return dbData;
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