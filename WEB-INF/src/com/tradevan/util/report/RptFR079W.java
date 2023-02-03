/*
107.05.21 create 洗錢關鍵字_依檢查報告編號_分年度 by Ethan
*/

package com.tradevan.util.report;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFFooter;
import org.apache.poi.hssf.usermodel.HSSFPrintSetup;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import com.tradevan.util.DBManager;
import com.tradevan.util.Utility;
import com.tradevan.util.dao.DataObject;

public class RptFR079W {
	private final static int START_TITLE_ROW =4;
	
	public static String createRpt(String strYear, String strMonth, String endYear, String endMonth, String bank_type, String report_no, String[] words)	{
		String errMsg = "";
		
		List keyWords = new ArrayList();
		for(int index=0; index<words.length; index++){
			String word = ((String)words[index]);
			if(!word.isEmpty()){
				keyWords.add(word);
			}
		}
		
		System.out.println("key size="+keyWords.size());
		
		reportUtil reportUtil = new reportUtil();
		StringBuffer sql = new StringBuffer () ;
		ArrayList paramList = new ArrayList();
						
		try{
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
			
			String openfile = "洗錢關鍵字_依檢查報告編號_分年度.xls";	
			System.out.println("open file=" + openfile);
			
			FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator")+ openfile );			
			//設定FileINputStream讀取Excel檔
			POIFSFileSystem fs = new POIFSFileSystem(finput);
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
			HSSFCellStyle cs_time = reportUtil.getNoBoderStyle(wb);
			HSSFCellStyle cs_center = reportUtil.getLeftStyle(wb);
			HSSFCellStyle cs_keyword = reportUtil.getDefaultStyle(wb);
			cs_keyword.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			cs_keyword.setFillForegroundColor(HSSFColor.TAN.index);
			
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
			//int rowNum = 0;
			
			String printTime = Utility.getDateFormat("  HH:mm:ss");
            String printDate = Utility.getDateFormat("yyyy/MM/dd");			
            				
			sql.setLength(0);
			paramList.clear();	

			//資料SQL
			sql.append("SELECT a.reportno, ");//檢查報告編號 
			sql.append("b.bank_no, "); //受檢單位.單位代號
			sql.append("c.bank_name, "); //受檢單位.單位名稱
			sql.append("d.cmuse_name as ch_type, "); //檢查性質
			sql.append("((TO_CHAR(b.base_date,'yyyy')-1911)||'/'|| TO_CHAR(b.base_date,'mm/dd')) base_date, "); //檢查基準日 
			sql.append("a.reportno_seq, "); 
			sql.append("a.item_no, ");//序號
			sql.append("a.serial, ");
			sql.append("a.ex_content, "); //檢查缺失摘要
			sql.append("a.commentt, ");//檢查處理意見 
			sql.append("e.digest, ");//函覆改善情形摘要
			sql.append("(select cmuse_name from CDShareNo WHERE a.audit_result = cmuse_id AND cmuse_div ='026' ) cmuse_name ");//審核結果 
			sql.append("FROM ExDefGoodF a, ExReportF b, (select * from ba01 where m_year=100 ) c, CDShareNO d , ExDG_HistoryF e ");   
			sql.append("WHERE a.reportno(+) = b.reportno AND b.bank_no = c.bank_no ");
			sql.append("and  a.reportno = e.reportno AND a.reportno_seq = e.reportno_seq(+) ");
			sql.append("AND (b.ch_type = d.cmuse_id AND cmuse_div = '023') ");
			sql.append("AND TO_CHAR(b.base_date, 'yyyymmdd') BETWEEN ? AND ? ");
			paramList.add(Utility.getDatetoString(Utility.getFullDate(strYear, strMonth, "01"))); //檢查基準日起
			paramList.add(Utility.getDatetoString(Utility.getFullDate(endYear, endMonth, "31"))); //檢查基準日迄	
					
			if(report_no.length() != 0)
				sql.append("AND a.reportno like '%").append(report_no).append("%'");
			
			if(keyWords.size() != 0){
				sql.append(" AND (");
				for(int index = 0; index <keyWords.size(); index++){
					if(index != 0)
						sql.append(" or ");
					
					sql.append("a.ex_content like '%");
					sql.append(keyWords.get(index));
					sql.append("%'");
				}
				sql.append(") ");
			}
			
			sql.append("order by reportno, item_no");

			System.out.println("sql ="+sql.toString());		
			
			List datas = DBManager.QueryDB_SQLParam(sql.toString(),paramList, "tbank_cnt_6,tbank_cnt_7,tbank_cnt,bank_cnt_6,bank_cnt_7,bank_cnt,bank_6_sum,bank_7_sum,bank_sum");
			System.out.println("datas size()="+datas.size());
			
			
			//各年資料(年份:各年資料)
			Map datasByYear = new HashMap();
			
			//建立各年活頁
			for(int sY = Integer.parseInt(strYear); sY <= Integer.parseInt(endYear); sY++){
				datasByYear.put(String.valueOf(sY), new ArrayList());
			
				sheet = wb.cloneSheet(0);
								
				if (sheet == null) {
					System.out.println("建表單失敗");
				} else {
					System.out.println("建表成功");
				}
			}
			
			//依年份過濾資料
			for (int i =0; i<datas.size(); i++){
				Map obj = ((DataObject)datas.get(i)).getValues();
				String y = ((String)obj.get("base_date")).substring(0, 3);
				
				if(datasByYear.containsKey(y)){
					List data = (List) datasByYear.get(y);
					data.add(obj);
				}
			}
			
			int datasByYearSize= 1;	//年份數
			
			for(Object key : datasByYear.keySet()){
				//改寫各年分表名
				wb.setSheetName(datasByYearSize, key+"年");
				
				//取得各年分表及設定
				sheet = wb.getSheetAt(datasByYearSize);
				ps = sheet.getPrintSetup();
				
				//設定表頭 為固定 先設欄的起始再設列的起始
				wb.setRepeatingRowsAndColumns(datasByYearSize, 0, 9, 0, 3);
				
				// 設定頁面符合列印大小
				sheet.setAutobreaks(false);
				ps.setScale((short)62); //列印縮放百分比
				ps.setPaperSize((short)9); //設定紙張大小 A4
				ps.setLandscape(true);//設定橫印
				
				row = sheet.getRow(0);
				cell = row.getCell((short) 0);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellValue("「MIS檢查追蹤管理系統」 涉及洗錢等關鍵字﹝註1﹞之檢查意見─"+key+"年度");
				
				// 設定輸出列印日期
				row = sheet.getRow(1);
				cell = row.getCell((short) 0);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellStyle(cs_time);
				cell.setCellValue("列印日期：" + Utility.getCHTdate(printDate, 1));
								
				List list = (List)datasByYear.get(key);
				for (int index=0; index< list.size(); index++){
					Map obj = (Map) list.get(index);
					int rowNum = index+START_TITLE_ROW;

					for(int column =0; column < 10; column++){
						String val = "";
						
						if (column == 0){
							val = obj.get("reportno") == null ? "" : (String)obj.get("reportno");
						} else if (column == 1) {
							val = obj.get("bank_no") == null ? "" : (String)obj.get("bank_no");
						} else if (column == 2) {
							val = obj.get("bank_name") == null ? "" : (String)obj.get("bank_name");
						} else if (column == 3) {
							val = obj.get("ch_type") == null ? "" : (String)obj.get("ch_type");
						} else if (column == 4) {
							val = obj.get("base_date") == null ? "" : (String)obj.get("base_date");
						} else if (column == 5) {
							val = obj.get("item_no") == null ? "" : (String)obj.get("item_no");
						} else if (column == 6) {
							val = obj.get("ex_content") == null ? "" : (String)obj.get("ex_content");
						} else if (column == 7) {
							val = obj.get("commentt") == null ? "" : (String)obj.get("commentt");
						} else if (column == 8) {
							val = obj.get("digest") == null ? "" : (String)obj.get("digest");
						} else if (column == 9){
							val = obj.get("cmuse_name") == null ? "" : (String)obj.get("cmuse_name");
						}
						
						row = sheet.createRow(rowNum);
						cell = row.createCell((short) column);
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell.setCellStyle(cs_center);
						cell.setCellValue(val);
					}
				}
				datasByYearSize++;
					
				//檢查是否有關鍵字
				if (keyWords.size() != 0){
					//size+7(資料總比數+3資料起始列+4關鍵字)
					int keyWordRow = list.size()+START_TITLE_ROW+4;
					row = sheet.createRow(keyWordRow);
					cell = row.createCell((short) 0);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellValue("[註1]關鍵字");
					
					keyWordRow++;
					int dataCountInRow =0;
					for(int index =0; index<keyWords.size(); index ++){
						
						//每列4筆資料
						if(dataCountInRow == 4){
							dataCountInRow =0;
							keyWordRow++;
						}
						
						row = sheet.createRow(keyWordRow);
						cell = row.createCell((short) dataCountInRow);
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell.setCellStyle(cs_keyword);
						cell.setCellValue((String)keyWords.get(index));
						dataCountInRow++;
					}	
				}				
			}
			
			HSSFFooter footer = sheet.getFooter();
			footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));	
			
			//全部資料:A4橫印,縮放比80%
			sheet = wb.getSheetAt(0); //讀取第一個工作表，宣告其為sheet
			if (sheet == null) {
				System.out.println("open sheet1 失敗");
			} else {
				System.out.println("open sheet1 成功");
			}
			ps = sheet.getPrintSetup(); //取得設定			
			
			//設定表頭 為固定 先設欄的起始再設列的起始
			wb.setRepeatingRowsAndColumns(0, 0, 9, 0, 3);
						
			// 設定頁面符合列印大小
			sheet.setAutobreaks(false);
			ps.setScale((short)62); //列印縮放百分比
			ps.setPaperSize((short)9); //設定紙張大小 A4
			ps.setLandscape(true);//設定橫印
			
			// 設定輸出列印時間
			row = sheet.getRow(1);
			cell = row.getCell((short) 0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellStyle(cs_time);
			cell.setCellValue("列印日期：" + Utility.getCHTdate(printDate, 1));			
			
			for(int count =0; count<datas.size(); count++){
				Map obj = ((DataObject)datas.get(count)).getValues();
							
				int rowNum = count+START_TITLE_ROW;
				
				for(int column =0; column < 10; column++){
					String val = "";
					
					if (column == 0){
						val = obj.get("reportno") == null ? "" : (String)obj.get("reportno");
					} else if (column == 1) {
						val = obj.get("bank_no") == null ? "" : (String)obj.get("bank_no");
					} else if (column == 2) {
						val = obj.get("bank_name") == null ? "" : (String)obj.get("bank_name");
					} else if (column == 3) {
						val = obj.get("ch_type") == null ? "" : (String)obj.get("ch_type");
					} else if (column == 4) {
						val = obj.get("base_date") == null ? "" : (String)obj.get("base_date");
					} else if (column == 5) {
						val = obj.get("item_no") == null ? "" : (String)obj.get("item_no");
					} else if (column == 6) {
						val = obj.get("ex_content") == null ? "" : (String)obj.get("ex_content");
					} else if (column == 7) {
						val = obj.get("commentt") == null ? "" : (String)obj.get("commentt");
					} else if (column == 8) {
						val = obj.get("digest") == null ? "" : (String)obj.get("digest");
					} else if (column == 9){
						val = obj.get("cmuse_name") == null ? "" : (String)obj.get("cmuse_name");
					}
					
					row = sheet.createRow(rowNum);
					cell = row.createCell((short) column);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs_center);
					cell.setCellValue(val);
				}				
			}
			
			//檢查是否有關鍵字
			if (keyWords.size() != 0){
				//size+7(資料總比數+4資料起始列+4關鍵字)
				int keyWordRow = datas.size()+START_TITLE_ROW+4;
				row = sheet.createRow(keyWordRow);
				cell = row.createCell((short) 0);
				cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellValue("[註1]關鍵字");
				
				keyWordRow++;
				int dataCountInRow =0;
				for(int index =0; index<keyWords.size(); index ++){
					
					//每列4筆資料
					if(dataCountInRow == 4){
						dataCountInRow =0;
						keyWordRow++;
					}
					
					row = sheet.createRow(keyWordRow);
					cell = row.createCell((short) dataCountInRow);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellStyle(cs_keyword);
					cell.setCellValue((String)keyWords.get(index));
					dataCountInRow++;
				}	
			}	
			
			FileOutputStream fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + openfile);
			
			footer = sheet.getFooter();
			footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));	
			
			wb.write(fout);
			// 儲存
			fout.close();
			System.out.println("write excel success");
		
		} catch (Exception e){
			System.out.println("createRpt Error:"+e+e.getMessage());
		}
		return errMsg;
	}
	
}
