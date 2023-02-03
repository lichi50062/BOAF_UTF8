/*
  Created on 2006/03/30 by 2495
    99.04.13 fix 縣市合併SQL調整 && 修改查詢方式為preparedstatement by 2808
   103.01.16 add 臺灣省農會更名為中華民國農會增加說明 by 2295
   103.12.23 fix 調整title出現奇怪的線 by 2295
 */
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.*;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import java.io.*;
import java.util.*;
import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR036WB {
	public static String createRpt(String S_YEAR, String S_MONTH,String E_YEAR, String E_MONTH, String unit, String bank_type,
		List bank_list, String rptStyle) {
		System.out.println("start RptFR036WA------------------");
		String errMsg = "";
		String S_YEAR_LAST = "";
		String E_YEAR_LAST = ""; 
		String filename = "";
		String openfile = "";
		int rowcount = 0;
		int flag = 0;
		int temp_year = Integer.parseInt(S_YEAR);
		temp_year = temp_year - 1;
		S_YEAR_LAST = Integer.toString(temp_year);
		int tmp_year2 = Integer.parseInt(E_YEAR) ;
		tmp_year2 = tmp_year2-1 ;
		E_YEAR_LAST = Integer.toString(tmp_year2) ;
		String u_year ="100" ;
		if(S_YEAR==null || Integer.parseInt(S_YEAR) < 100) {
        	u_year = "99" ;
        }
		String bankNm = "6".equals(bank_type) ? "農" : "漁";
		String reportType = "0".equals(rptStyle) ? "總表" : "明細表";
		filename = "全體" + bankNm + "會信用部統一" + bankNm + "貸資料某二指定期間新增戶數金額比較"+ reportType + ".xls";
		openfile = filename;

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

			System.out.println("open file " + openfile);
			FileInputStream finput = new FileInputStream(xlsDir+ System.getProperty("file.separator") + openfile);

			// 設定FileINputStream讀取Excel檔
			POIFSFileSystem fs = new POIFSFileSystem(finput);
			if (fs == null) {
				System.out.println("open 範本檔失敗");
			} 
			HSSFWorkbook wb = new HSSFWorkbook(fs);
			if (wb == null) {
				System.out.println("open工作表失敗");
			} 
			HSSFSheet sheet = wb.getSheetAt(0);// 讀取第一個工作表，宣告其為sheet
			if (sheet == null) {
				System.out.println("open sheet 失敗");
			} 
			HSSFPrintSetup ps = sheet.getPrintSetup(); // 取得設定
			// sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
			// sheet.setAutobreaks(true); //自動分頁

			// 設定頁面符合列印大小
			sheet.setAutobreaks(false);
			ps.setScale((short) 70); // 列印縮放百分比

			ps.setPaperSize((short) 9); // 設定紙張大小 A4

			// 設定表頭 為固定 先設欄的起始再設列的起始
			wb.setRepeatingRowsAndColumns(0, 1, 10, 1, 3);

			finput.close();

			HSSFRow row = null;// 宣告一列
			HSSFCell cell = null;// 宣告一個儲存格

			HSSFFont ft = wb.createFont();
			HSSFCellStyle cs = wb.createCellStyle();
			HSSFFont ft2 = wb.createFont();
			HSSFCellStyle cs2 = wb.createCellStyle();
			HSSFCellStyle cs3 = wb.createCellStyle();
			HSSFFont f = wb.createFont();
			HSSFCellStyle c = wb.createCellStyle();

			row = sheet.getRow(0);
			cell = row.getCell((short) 0);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16); // 設定這個儲存格的字串要儲存雙位元,
														// 中文才能正常顯示
			cell.setCellValue( "全體" + bankNm + "會信用部統一" + bankNm + "貸資料某二指定期間新增戶數金額比較"+ reportType);
			int rowNum = 1;

			// 取得bank_list的bank_no
			String bank_id = "";
			String sqlCmd = "";
			StringBuffer sql = new StringBuffer() ;
			List paramList = new ArrayList() ;
			if (rptStyle.equals("1") && bank_list.size() > 0) {
				// 明細==========================================
				sql.append(getReportDetailSQL(u_year,bank_list)) ; //取得報表明細SQL
				paramList.add(unit );
				paramList.add(unit );
				paramList.add(unit );
				paramList.add(unit );
				paramList.add(unit );
				paramList.add(u_year     );
				paramList.add(bank_type  );
				paramList.add(u_year     );
				paramList.add(S_YEAR+S_MONTH);
				paramList.add(E_YEAR+E_MONTH);
				paramList.add(unit          );
				paramList.add(unit          );
				paramList.add(unit          );
				paramList.add(unit          );
				paramList.add(unit          );
				paramList.add(u_year        );
				paramList.add(bank_type     );
				paramList.add(u_year        );
				paramList.add(S_YEAR_LAST+S_MONTH );
				paramList.add(E_YEAR_LAST+E_MONTH );

				
				List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"hsien_id,hsien_name,fr001w_output_order,bank_no,bank_name,creditmonth_cnt,creditmonth_amt,overcreditmonth_cnt,overcreditmonth_amt,last_creditmonth_cnt,last_creditmonth_amt,last_overcreditmonth_cnt,last_overcreditmonth_amt");
				
				System.out.print("dbData.size =" + dbData.size());
				DataObject bean = null;
				String unit_name = "";

				if (dbData.size() != 0) {
					// 列印年度
					row = (sheet.getRow(1) == null) ? sheet.createRow(1): sheet.getRow(1);
					cell = row.getCell((short) 12);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);

					System.out.println("test2");
					ft.setFontName("標楷體");
					ft.setFontHeightInPoints((short) 14);
					//cs.setFont(ft);

					cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);
					//cell.setCellStyle(cs);
					unit_name = Utility.getUnitName(unit) ;
					
					if (S_MONTH.equals("0")) {
						cell.setCellValue("中華民國" + S_YEAR + "年度");
					} else {
						cell.setCellValue(" 單位：新台幣" + unit_name + "、％");
					}

					// 列印單位
					cell = row.getCell((short) 11);
					f.setFontName("標楷體");
					f.setFontHeightInPoints((short) 14);
					//cs.setFont(f);

					rowNum = 5;

				} else {
					row = (sheet.getRow(1) == null) ? sheet.createRow(1): sheet.getRow(1);
					cell = row.getCell((short) 3);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);

					f.setFontName("標楷體");
					f.setFontHeightInPoints((short) 14);
					cs.setFont(f);
					cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);
					//cell.setCellStyle(cs);
					if (S_MONTH.equals("0")) {
						cell.setCellValue("中華民國" + S_YEAR + "年度");
					} else {
						cell.setCellValue("            中華民國" + S_YEAR + "年"+ S_MONTH + "月無資料存在");
					}
					System.out.println("debug---無資料");
				}

				if (dbData != null) {
					row = (sheet.getRow(rowNum) == null) ? sheet.createRow(rowNum) : sheet.getRow(rowNum);
					for (int cellcount = 1; cellcount < 15; cellcount++) {
						cell = row.createCell((short) cellcount);
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);

						ft.setFontName("標楷體");
						ft.setFontHeightInPoints((short) 10);
						cs2.setFont(ft);
						cs2.setAlignment(HSSFCellStyle.ALIGN_CENTER);
						cell.setCellStyle(cs2);
						if (cellcount == 3 || cellcount == 6 || cellcount == 9|| cellcount == 12) {
							cell.setCellValue(S_YEAR + "年" + "\n" + S_MONTH+ "月-" + E_MONTH + "月");
						}
						if (cellcount == 4 || cellcount == 7 || cellcount == 10|| cellcount == 13) {
							cell.setCellValue(S_YEAR_LAST + "年" + "\n" + S_MONTH + "月-" + E_MONTH + "月");
						}
						if (cellcount == 5 || cellcount == 8 || cellcount == 11|| cellcount == 14) {
							cell.setCellValue("增減");
						}
					}
					for (int k = 0; k < dbData.size(); k++) {

						bean = (DataObject) dbData.get(k);
						String hsien_id = String.valueOf(bean.getValue("hsien_id"));
						String hsien_name = String.valueOf(bean.getValue("hsien_name"));
						String fr001w_output_order = String.valueOf(bean.getValue("fr001w_output_order"));
						String bank_no = String.valueOf(bean.getValue("bank_no"));
						String bank_name = String.valueOf(bean.getValue("bank_name"));
						String preyearset_creditmonth_cnt = String.valueOf(bean.getValue("creditmonth_cnt"));
						String preyearset_creditmonth_amt = String.valueOf(bean.getValue("creditmonth_amt"));
						String preyearset_overcreditmonth_cnt = String.valueOf(bean.getValue("overcreditmonth_cnt"));
						String preyearset_overcreditmonth_amt = String.valueOf(bean.getValue("overcreditmonth_amt"));
						String lastyearset_creditmonth_cnt = String.valueOf(bean.getValue("last_creditmonth_cnt"));
						String lastyearset_creditmonth_amt = String.valueOf(bean.getValue("last_creditmonth_amt"));
						String lastyearset_overcreditmonth_cnt = String.valueOf(bean.getValue("last_overcreditmonth_cnt"));
						String lastyearset_overcreditmonth_amt = String.valueOf(bean.getValue("last_overcreditmonth_amt"));

						rowNum++;
						row = (sheet.getRow(rowNum) == null) ? sheet.createRow(rowNum) : sheet.getRow(rowNum);
						for (int cellcount = 0; cellcount < 15; cellcount++) {
							cell = row.createCell((short) cellcount);
							cell.setEncoding(HSSFCell.ENCODING_UTF_16);

							cs2.setBorderTop(HSSFCellStyle.BORDER_THIN);
							cs2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
							cs2.setBorderLeft(HSSFCellStyle.BORDER_THIN);
							cs2.setBorderRight(HSSFCellStyle.BORDER_THIN);

							ft.setFontName("標楷體");
							ft.setFontHeightInPoints((short) 12);
							cs2.setFont(ft);
							cs2.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
							cell.setCellStyle(cs2);
							if (cellcount == 0) {
								cell.setCellValue(hsien_name);
							}
							if (cellcount == 1) {
								cell.setCellValue(bank_no);
							}
							if (cellcount == 2) {
								cs3.setBorderTop(HSSFCellStyle.BORDER_THIN);
								cs3.setBorderBottom(HSSFCellStyle.BORDER_THIN);
								cs3.setBorderLeft(HSSFCellStyle.BORDER_THIN);
								cs3.setBorderRight(HSSFCellStyle.BORDER_THIN);
								ft2.setFontName("標楷體");
								ft2.setFontHeightInPoints((short) 12);
								cs3.setFont(ft2);
								cs3.setAlignment(HSSFCellStyle.ALIGN_LEFT);
								cell.setCellStyle(cs3);
								cell.setCellValue(bank_name);
							}

							if (cellcount == 3) {
								preyearset_creditmonth_cnt = (preyearset_creditmonth_cnt.equals("null")) ? "0": preyearset_creditmonth_cnt;
								cell.setCellValue(preyearset_creditmonth_cnt);
							}
							if (cellcount == 4) {
								lastyearset_creditmonth_cnt = (lastyearset_creditmonth_cnt.equals("null")) ? "0": lastyearset_creditmonth_cnt;
								cell.setCellValue(lastyearset_creditmonth_cnt);

							}
							if (cellcount == 5) {
								int temp_preyearset_creditmonth_cnt = Integer.parseInt(preyearset_creditmonth_cnt);
								int temp_lastyearset_creditmonth_cnt = Integer.parseInt(lastyearset_creditmonth_cnt);
								int temp_preyearset_creditmonth_min = temp_preyearset_creditmonth_cnt- temp_lastyearset_creditmonth_cnt;
								String preyearset_creditmonth_min = Integer.toString(temp_preyearset_creditmonth_min);
								cell.setCellValue(preyearset_creditmonth_min);
							}
							if (cellcount == 6) {
								preyearset_creditmonth_amt = (preyearset_creditmonth_amt.equals("null")) ? "0": preyearset_creditmonth_amt;
								cell.setCellValue(preyearset_creditmonth_amt);

							}
							if (cellcount == 7) {
								lastyearset_creditmonth_amt = (lastyearset_creditmonth_amt.equals("null")) ? "0": lastyearset_creditmonth_amt;
								cell.setCellValue(lastyearset_creditmonth_amt);
							}
							if (cellcount == 8) {
								int temp_preyearset_creditmonth_amt = Integer.parseInt(preyearset_creditmonth_amt);
								int temp_lastyearset_creditmonth_amt = Integer.parseInt(lastyearset_creditmonth_amt);
								int temp_preyearset_creditmonth_min = temp_preyearset_creditmonth_amt- temp_lastyearset_creditmonth_amt;
								String preyearset_creditmonth_min = Integer.toString(temp_preyearset_creditmonth_min);
								cell.setCellValue(preyearset_creditmonth_min);
							}
							if (cellcount == 9) {
								preyearset_overcreditmonth_cnt = (preyearset_overcreditmonth_cnt.equals("null")) ? "0": preyearset_overcreditmonth_cnt;
								cell.setCellValue(preyearset_overcreditmonth_cnt);
							}
							if (cellcount == 10) {
								lastyearset_overcreditmonth_cnt = (lastyearset_overcreditmonth_cnt.equals("null")) ? "0": lastyearset_overcreditmonth_cnt;
								cell.setCellValue(lastyearset_overcreditmonth_cnt);
							}
							if (cellcount == 11) {
								int temp_preyearset_overcreditmonth_cnt = Integer.parseInt(preyearset_overcreditmonth_cnt);
								int temp_lastyearset_overcreditmonth_cnt = Integer.parseInt(lastyearset_overcreditmonth_cnt);
								int temp_preyearset_overcreditmonth_min = temp_preyearset_overcreditmonth_cnt- temp_lastyearset_overcreditmonth_cnt;
								String preyearset_overcreditmonth_min = Integer.toString(temp_preyearset_overcreditmonth_min);
								cell.setCellValue(preyearset_overcreditmonth_min);
							}
							if (cellcount == 12) {

								preyearset_overcreditmonth_amt = (preyearset_overcreditmonth_amt.equals("null")) ? "0": preyearset_overcreditmonth_amt;
								cell.setCellValue(preyearset_overcreditmonth_amt);

							}
							if (cellcount == 13) {
								lastyearset_overcreditmonth_amt = (lastyearset_overcreditmonth_amt.equals("null")) ? "0": lastyearset_overcreditmonth_amt;
								cell.setCellValue(lastyearset_overcreditmonth_amt);
							}
							if (cellcount == 14) {
								int temp_preyearset_overcreditmonth_amt = Integer.parseInt(preyearset_overcreditmonth_amt);
								int temp_lastyearset_overcreditmonth_amt = Integer.parseInt(lastyearset_overcreditmonth_amt);
								int temp_preyearset_overcreditmonth_min = temp_preyearset_overcreditmonth_amt- temp_lastyearset_overcreditmonth_amt;
								String preyearset_overcreditmonth_min = Integer.toString(temp_preyearset_overcreditmonth_min);
								cell.setCellValue(preyearset_overcreditmonth_min);

							}

						}
						
					}
				}
			}

			if (rptStyle.equals("0")) {
				sql.append(getReportTotalSQL(u_year,unit)) ;//取得總表SQL
				paramList.add(u_year) ;
				paramList.add(bank_type) ;
				paramList.add(u_year) ;
				paramList.add(S_YEAR+S_MONTH) ;
				paramList.add(E_YEAR+E_MONTH) ;
				paramList.add(u_year) ;
				paramList.add(bank_type) ;
				paramList.add(u_year) ;
				paramList.add(S_YEAR+S_MONTH) ;
				paramList.add(E_YEAR+E_MONTH) ;
				paramList.add(u_year) ;
				paramList.add(bank_type) ;
				paramList.add(u_year) ;
				paramList.add(S_YEAR+S_MONTH) ;
				paramList.add(E_YEAR+E_MONTH) ;
				paramList.add(u_year) ;
				paramList.add(bank_type) ;
				paramList.add(u_year) ;
				paramList.add(S_YEAR+S_MONTH) ;
				paramList.add(E_YEAR+E_MONTH) ;
				

				// 總表_各縣市
				List dbData = DBManager.QueryDB_SQLParam(sql.toString(),paramList,"hsien_id,hsien_name,fr001w_output_order,bank_no,bank_name,creditmonth_cnt,creditmonth_amt,overcreditmonth_cnt,overcreditmonth_amt,last_creditmonth_cnt,last_creditmonth_amt,last_overcreditmonth_cnt,last_overcreditmonth_amt");
				System.out.print("單月總表抓出的dbData.size =" + dbData.size());
				DataObject bean = null;
				String unit_name = "";
				if (dbData.size() > 0) {
					row = (sheet.getRow(1) == null) ? sheet.createRow(1) : sheet.getRow(1);
					cell = row.getCell((short) 10);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					
					f.setFontName("標楷體");
					f.setFontHeightInPoints((short) 10);
					cs2.setFont(f);

					
					//cell.setCellStyle(cs2);
					unit_name = Utility.getUnitName(unit) ;
					
					if (S_MONTH.equals("0")) {
						cell.setCellValue("中華民國" + S_YEAR + "年度");
					} else {
						
						cell.setCellValue(" 單位：新台幣" + unit_name + "、％ ");
					}
					rowNum = 5;
					System.out.print("總表資料dbData.size()" + dbData.size());
				} else {
					System.out.print("總表尚無資料");
					// 列印年度
					System.out.println("dbData.size()=" + dbData.size());
					row = (sheet.getRow(1) == null) ? sheet.createRow(1)
							: sheet.getRow(1);
					cell = row.getCell((short) 1);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);

					f.setFontName("標楷體");
					f.setFontHeightInPoints((short) 10);
					cs.setFont(f);
					cs.setAlignment(HSSFCellStyle.ALIGN_CENTER);
					cell.setCellStyle(cs);
					if (S_MONTH.equals("0")) {
						cell.setCellValue("中華民國" + S_YEAR + "年度");
					} else {
						cell.setCellValue("中華民國" + S_YEAR + "年" + S_MONTH+ "月無資料存在");
					}

				}
				row = (sheet.getRow(rowNum) == null) ? sheet.createRow(rowNum): sheet.getRow(rowNum);
				for (int cellcount = 1; cellcount < 13; cellcount++) {
					cell = row.createCell((short) cellcount);
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);

					ft.setFontName("標楷體");
					ft.setFontHeightInPoints((short) 10);
					cs2.setFont(ft);
					cs2.setAlignment(HSSFCellStyle.ALIGN_CENTER);
					cell.setCellStyle(cs2);
					if (cellcount == 1 || cellcount == 4 || cellcount == 7|| cellcount == 10) {
						cell.setCellValue(S_YEAR + "年" + "\n" + S_MONTH + "月-"+ E_MONTH + "月");
					}
					if (cellcount == 2 || cellcount == 5 || cellcount == 8|| cellcount == 11) {
						cell.setCellValue(S_YEAR_LAST + "年" + "\n" + S_MONTH+ "月-" + E_MONTH + "月");
					}
					if (cellcount == 3 || cellcount == 6 || cellcount == 9|| cellcount == 12) {
						cell.setCellValue("增減");
					}
				}
				System.out.println("準備列印~~~");
				if (dbData != null) {
					System.out.println("開始列印~~~");
					for (int k = 0; k < dbData.size(); k++) {
						bean = (DataObject) dbData.get(k);
						String hsien_id = String.valueOf(bean
								.getValue("hsien_id"));
						String hsien_name = String.valueOf(bean
								.getValue("hsien_name"));
						String fr001w_output_order = String.valueOf(bean
								.getValue("fr001w_output_order"));
						String bank_no = String.valueOf(bean
								.getValue("bank_no"));
						String bank_name = String.valueOf(bean
								.getValue("bank_name"));
						String preyearset_creditmonth_cnt = String.valueOf(bean
								.getValue("creditmonth_cnt"));
						String preyearset_creditmonth_amt = String.valueOf(bean
								.getValue("creditmonth_amt"));
						String preyearset_overcreditmonth_cnt = String
								.valueOf(bean.getValue("overcreditmonth_cnt"));
						String preyearset_overcreditmonth_amt = String
								.valueOf(bean.getValue("overcreditmonth_amt"));
						String lastyearset_creditmonth_cnt = String
								.valueOf(bean.getValue("last_creditmonth_cnt"));
						String lastyearset_creditmonth_amt = String
								.valueOf(bean.getValue("last_creditmonth_amt"));
						String lastyearset_overcreditmonth_cnt = String
								.valueOf(bean
										.getValue("last_overcreditmonth_cnt"));
						String lastyearset_overcreditmonth_amt = String
								.valueOf(bean
										.getValue("last_overcreditmonth_amt"));
						rowNum++;
						row = (sheet.getRow(rowNum) == null) ? sheet.createRow(rowNum) : sheet.getRow(rowNum);
						for (int cellcount = 0; cellcount < 13; cellcount++) {
							//System.out.println("cellcount=" + cellcount);
							cell = row.createCell((short) cellcount);
							cell.setEncoding(HSSFCell.ENCODING_UTF_16);

							cs2.setBorderTop(HSSFCellStyle.BORDER_THIN);
							cs2.setBorderBottom(HSSFCellStyle.BORDER_THIN);
							cs2.setBorderLeft(HSSFCellStyle.BORDER_THIN);
							cs2.setBorderRight(HSSFCellStyle.BORDER_THIN);

							ft.setFontName("標楷體");
							ft.setFontHeightInPoints((short) 12);
							cs2.setFont(ft);
							cs2.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
							cell.setCellStyle(cs2);
							if (cellcount == 0) {
								cell.setCellValue(hsien_name);
							}
							if (cellcount == 1) {
								preyearset_creditmonth_cnt = (preyearset_creditmonth_cnt
										.equals("null")) ? "0"
										: preyearset_creditmonth_cnt;
								cell.setCellValue(preyearset_creditmonth_cnt);
							}
							if (cellcount == 2) {
								lastyearset_creditmonth_cnt = (lastyearset_creditmonth_cnt
										.equals("null")) ? "0"
										: lastyearset_creditmonth_cnt;
								cell.setCellValue(lastyearset_creditmonth_cnt);

							}
							if (cellcount == 3) {
								int temp_preyearset_creditmonth_cnt = Integer
										.parseInt(preyearset_creditmonth_cnt);
								int temp_lastyearset_creditmonth_cnt = Integer
										.parseInt(lastyearset_creditmonth_cnt);
								int temp_preyearset_creditmonth_min = temp_preyearset_creditmonth_cnt
										- temp_lastyearset_creditmonth_cnt;
								String preyearset_creditmonth_min = Integer
										.toString(temp_preyearset_creditmonth_min);
								cell.setCellValue(preyearset_creditmonth_min);
							}
							if (cellcount == 4) {

								preyearset_creditmonth_amt = (preyearset_creditmonth_amt
										.equals("null")) ? "0"
										: preyearset_creditmonth_amt;
								cell.setCellValue(preyearset_creditmonth_amt);

							}
							if (cellcount == 5) {
								lastyearset_creditmonth_amt = (lastyearset_creditmonth_amt
										.equals("null")) ? "0"
										: lastyearset_creditmonth_amt;
								cell.setCellValue(lastyearset_creditmonth_amt);

							}
							if (cellcount == 6) {
								int temp_preyearset_creditmonth_amt = Integer
										.parseInt(preyearset_creditmonth_amt);
								int temp_lastyearset_creditmonth_amt = Integer
										.parseInt(lastyearset_creditmonth_amt);
								int temp_preyearset_creditmonth_min = temp_preyearset_creditmonth_amt
										- temp_lastyearset_creditmonth_amt;
								String preyearset_creditmonth_min = Integer
										.toString(temp_preyearset_creditmonth_min);
								cell.setCellValue(preyearset_creditmonth_min);
							}
							if (cellcount == 7) {

								preyearset_overcreditmonth_cnt = (preyearset_overcreditmonth_cnt
										.equals("null")) ? "0"
										: preyearset_overcreditmonth_cnt;
								cell
										.setCellValue(preyearset_overcreditmonth_cnt);

							}
							if (cellcount == 8) {
								lastyearset_overcreditmonth_cnt = (lastyearset_overcreditmonth_cnt
										.equals("null")) ? "0"
										: lastyearset_overcreditmonth_cnt;
								cell
										.setCellValue(lastyearset_overcreditmonth_cnt);
							}
							if (cellcount == 9) {
								int temp_preyearset_overcreditmonth_cnt = Integer
										.parseInt(preyearset_overcreditmonth_cnt);
								int temp_lastyearset_overcreditmonth_cnt = Integer
										.parseInt(lastyearset_overcreditmonth_cnt);
								int temp_preyearset_overcreditmonth_min = temp_preyearset_overcreditmonth_cnt
										- temp_lastyearset_overcreditmonth_cnt;
								String preyearset_overcreditmonth_min = Integer
										.toString(temp_preyearset_overcreditmonth_min);
								cell
										.setCellValue(preyearset_overcreditmonth_min);
							}
							if (cellcount == 10) {

								preyearset_overcreditmonth_amt = (preyearset_overcreditmonth_amt
										.equals("null")) ? "0"
										: preyearset_overcreditmonth_amt;
								cell
										.setCellValue(preyearset_overcreditmonth_amt);

							}
							if (cellcount == 11) {
								lastyearset_overcreditmonth_amt = (lastyearset_overcreditmonth_amt
										.equals("null")) ? "0"
										: lastyearset_overcreditmonth_amt;
								cell
										.setCellValue(lastyearset_overcreditmonth_amt);
							}
							if (cellcount == 12) {
								int temp_preyearset_overcreditmonth_amt = Integer
										.parseInt(preyearset_overcreditmonth_amt);
								int temp_lastyearset_overcreditmonth_amt = Integer
										.parseInt(lastyearset_overcreditmonth_amt);
								int temp_preyearset_overcreditmonth_min = temp_preyearset_overcreditmonth_amt
										- temp_lastyearset_overcreditmonth_amt;
								String preyearset_overcreditmonth_min = Integer
										.toString(temp_preyearset_overcreditmonth_min);
								cell
										.setCellValue(preyearset_overcreditmonth_min);

							}
						}

					}
				}

				// 總表_合計  == >SQL已合併
				
			}
			rowNum++;
            row = sheet.createRow(rowNum);
            cell = row.createCell( (short)0);
            cell.setEncoding(HSSFCell.ENCODING_UTF_16);
            cell.setCellValue("縣市別欄之其他(農會)係指原臺灣省農會，該農會於102年5月22日更名為中華民國農會。");
            
			HSSFFooter footer = sheet.getFooter();
			footer.setCenter("Page:" + HSSFFooter.page() + " of "+ HSSFFooter.numPages());
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
			FileOutputStream fout = new FileOutputStream(reportDir+ System.getProperty("file.separator") + openfile);
			wb.write(fout);
			// 儲存
			fout.close();
			System.out.println("儲存完成");

		} catch (Exception e) {
			System.out.println("createRpt Error:" + e + e.getMessage());
		}
		return errMsg;
	}

	public static void insertCell(String insertValue, int rowNum,
			int cellcount, HSSFWorkbook wb, HSSFRow row, HSSFSheet sheet,
			HSSFCell cell) {

	}
	private static String getReportTotalSQL(String u_year,String unit) {
		StringBuffer  sql = new StringBuffer();
		List paramList = new ArrayList() ;
		String cd01Table = "cd01" ;
		if("99".equals(u_year)) {
			cd01Table = "cd01_99" ;
		}
		sql.append(" select PreYearSet.hsien_id  	as  hsien_id,    ");
		sql.append("        PreYearSet.hsien_name	as	hsien_name,  ");
		sql.append("        PreYearSet.FR001W_output_order	as	FR001W_output_order,     ");
		sql.append("        PreYearSet.CreditMonth_Cnt		as	CreditMonth_Cnt,         ");
		sql.append("        PreYearSet.CreditMonth_Amt		as	CreditMonth_Amt,         ");
		sql.append("        PreYearSet.OverCreditMonth_Cnt	as	OverCreditMonth_Cnt,     ");
		sql.append("        PreYearSet.OverCreditMonth_Amt	as	OverCreditMonth_Amt,     ");
		sql.append("        LastYearSet.CreditMonth_Cnt		as	Last_CreditMonth_Cnt,    ");
		sql.append("        LastYearSet.CreditMonth_Amt		as	Last_CreditMonth_Amt,    ");
		sql.append("        LastYearSet.OverCreditMonth_Cnt	as	Last_OverCreditMonth_Cnt,");
		sql.append("        LastYearSet.OverCreditMonth_Amt	as	Last_OverCreditMonth_Amt ");
		sql.append(" from   ");
		sql.append(" (select  * from ( ");
		sql.append("  select nvl(cd01.hsien_id,' ')        as  hsien_id ,                ");
		sql.append("         nvl(cd01.hsien_name,'OTHER')  as  hsien_name,               ");
		sql.append("         cd01.FR001W_output_order      as  FR001W_output_order,      ");
		sql.append("         sum(CreditMonth_Cnt)                as  CreditMonth_Cnt,    ");
		sql.append("         Round(sum(CreditMonth_Amt)/").append(unit).append(",0)     as  CreditMonth_Amt,    ");
		sql.append("         sum(OverCreditMonth_Cnt)            as  OverCreditMonth_Cnt,"); 
		sql.append("         Round(sum(OverCreditMonth_Amt)/").append(unit).append(",0) as  OverCreditMonth_Amt ");
		sql.append("  from( ");
		sql.append("  select * from ").append(cd01Table).append(" cd01 where cd01.hsien_id <> 'Y') cd01  ");
		sql.append("  left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year =? ");
		sql.append("  left join bn01  on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? and bn01.m_year = wlx01.m_year and bn01.m_year = ? ");
		sql.append("  left join (select bank_no, ");         
		sql.append("                    sum(CreditMonth_Cnt)      as  CreditMonth_Cnt,                ");
		sql.append("                    Round(sum(CreditMonth_Amt)/").append(unit).append(",0)  as CreditMonth_Amt,          ");
		sql.append("                    Sum(CreditYear_Cnt_Acc)   as  CreditYear_Cnt_Acc,             ");
		sql.append("                    Round(sum(CreditYear_Amt_Acc)/").append(unit).append(",0)  as CreditYear_Amt_Acc,    ");
		sql.append("                    sum(Credit_Cnt) as Credit_Cnt,                                ");
		sql.append("                    Round(sum(Credit_Bal)/").append(unit).append(",0)  as Credit_Bal,                    ");
		sql.append("                    Sum(OverCreditMonth_Cnt)  as OverCreditMonth_Cnt,             ");
		sql.append("                    Round(sum(OverCreditMonth_Amt)/").append(unit).append(",0)   as  OverCreditMonth_Amt,"); 
		sql.append("                    Sum(OverCredit_Cnt)   as OverCredit_Cnt,                      ");
		sql.append("                    Round(sum(OverCredit_Bal)/").append(unit).append(",0)  as  OverCredit_Bal            ");
		sql.append("             from WLX07_M_CREDIT                                                  ");
		sql.append("             where to_char(WLX07_M_CREDIT.m_year * 100 + WLX07_M_CREDIT.m_month) >= ? ");
		sql.append("             and   to_char(WLX07_M_CREDIT.m_year * 100 + WLX07_M_CREDIT.m_month) <= ? ");
		sql.append("             group by bank_no) a01 on  bn01.bank_no = a01.bank_no                         ");
		sql.append("   group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order  ");
		sql.append("   ) a01                                                                                  ");
		sql.append("   order by  a01.FR001W_output_order, a01.hsien_id                                        ");
		sql.append("  ) PreYearSet                                                                            ");
		sql.append("  left join                                                                               ");
		sql.append(" (select  * from (                                                                        ");
		sql.append("  select nvl(cd01.hsien_id,' ')       as  hsien_id ,");
		sql.append("         nvl(cd01.hsien_name,'OTHER') as  hsien_name,");
		sql.append("         cd01.FR001W_output_order     as  FR001W_output_order, ");
		sql.append("         sum(CreditMonth_Cnt)                as  CreditMonth_Cnt,");
		sql.append("         Round(sum(CreditMonth_Amt)/").append(unit).append(",0)     as  CreditMonth_Amt, ");
		sql.append("         sum(OverCreditMonth_Cnt)            as  OverCreditMonth_Cnt, ");
		sql.append("         Round(sum(OverCreditMonth_Amt)/").append(unit).append(",0) as  OverCreditMonth_Amt ");
		sql.append("  from (select * from ").append(cd01Table).append(" cd01 where cd01.hsien_id <> 'Y') cd01 ");
		sql.append("  left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ? ");
		sql.append("  left join bn01  on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? and bn01.m_year = wlx01.m_year and bn01.m_year = ? ");
		sql.append("  left join (select bank_no,   ");
		sql.append("                    sum(CreditMonth_Cnt)      as  CreditMonth_Cnt, ");
		sql.append("                    Round(sum(CreditMonth_Amt)/").append(unit).append(",0)  as CreditMonth_Amt,");
		sql.append("                    Sum(CreditYear_Cnt_Acc)   as  CreditYear_Cnt_Acc, ");
		sql.append("                    Round(sum(CreditYear_Amt_Acc)/").append(unit).append(",0)  as CreditYear_Amt_Acc, ");
		sql.append("                    sum(Credit_Cnt) as Credit_Cnt,  ");
		sql.append("                    Round(sum(Credit_Bal)/").append(unit).append(",0)  as Credit_Bal,");
		sql.append("                    Sum(OverCreditMonth_Cnt)  as OverCreditMonth_Cnt,   ");
		sql.append("                    Round(sum(OverCreditMonth_Amt)/").append(unit).append(",0)   as  OverCreditMonth_Amt,  ");
		sql.append("                    Sum(OverCredit_Cnt)   as OverCredit_Cnt, ");
		sql.append("                    Round(sum(OverCredit_Bal)/").append(unit).append(",0)  as  OverCredit_Bal   ");
		sql.append("             from WLX07_M_CREDIT ");
		sql.append("             where to_char(WLX07_M_CREDIT.m_year * 100 + WLX07_M_CREDIT.m_month) >=  ?  ");
		sql.append("             and   to_char(WLX07_M_CREDIT.m_year * 100 + WLX07_M_CREDIT.m_month) <=  ?  ");
		sql.append("             group by bank_no) a01 on  bn01.bank_no = a01.bank_no  ");
		sql.append("   group by nvl(cd01.hsien_id,' '),nvl(cd01.hsien_name,'OTHER'),cd01.FR001W_output_order ");
		sql.append("   ) a01  ");
		sql.append("   order by  a01.FR001W_output_order, a01.hsien_id  ");
		sql.append(" ) LastYearSet    ");
		sql.append(" on PreYearSet.hsien_id=LastYearSet.hsien_id  ");
		sql.append(" union  ");
		
		sql.append(" select ' '    as  hsien_id ,                                                                                                ");
		sql.append(" 	   '合計' as hsien_name,                                                                                                   ");
		sql.append("        PreYearSet.FR001W_output_order	as	FR001W_output_order,                                                             ");
		sql.append("        PreYearSet.CreditMonth_Cnt		as	CreditMonth_Cnt,                                                                   ");
		sql.append("        PreYearSet.CreditMonth_Amt		as	CreditMonth_Amt,                                                                   ");
		sql.append("        PreYearSet.OverCreditMonth_Cnt	as	OverCreditMonth_Cnt,                                                             ");
		sql.append("        PreYearSet.OverCreditMonth_Amt	as	OverCreditMonth_Amt,                                                             ");
		sql.append("        LastYearSet.CreditMonth_Cnt	    as	Last_CreditMonth_Cnt,                                                            ");
		sql.append("        LastYearSet.CreditMonth_Amt		as	Last_CreditMonth_Amt,                                                              ");
		sql.append("        LastYearSet.OverCreditMonth_Cnt	as	Last_OverCreditMonth_Cnt,                                                        ");
		sql.append("        LastYearSet.OverCreditMonth_Amt	as	Last_OverCreditMonth_Amt                                                         ");
		sql.append(" from                                                                                                                        ");
		sql.append(" ( select ' '    as  hsien_id ,                                                                                              ");
		sql.append("          '合計' as hsien_name,                                                                                              ");
		sql.append("          '999'  as  FR001W_output_order,                                                                                    ");
		sql.append("  	     sum(CreditMonth_Cnt)                 as  CreditMonth_Cnt,                                                           ");
		sql.append("  	     Round(sum(CreditMonth_Amt)/").append(unit).append(",0)      as  CreditMonth_Amt,                                                           ");
		sql.append("  	     sum(OverCreditMonth_Cnt)             as  OverCreditMonth_Cnt,                                                       ");
		sql.append("  	     Round(sum(OverCreditMonth_Amt)/").append(unit).append(",0)  as  OverCreditMonth_Amt                                                        ");
		sql.append("   from  (select * from ").append(cd01Table).append(" cd01 where cd01.hsien_id <> 'Y') cd01                                                                ");
		sql.append("   left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ? ");
		sql.append("   left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? and bn01.m_year = wlx01.m_year and bn01.m_year = ?");
		sql.append("   left join (select bank_no,                                                                                                ");
		sql.append("  	                sum(CreditMonth_Cnt)      as  CreditMonth_Cnt,                                                           ");
		sql.append(" 	                Round(sum(CreditMonth_Amt)/").append(unit).append(",0)  as CreditMonth_Amt,                                                       ");
		sql.append(" 	                Sum(CreditYear_Cnt_Acc)   as  CreditYear_Cnt_Acc,                                                          ");
		sql.append(" 	                Round(sum(CreditYear_Amt_Acc)/").append(unit).append(",0)  as CreditYear_Amt_Acc,                                                 ");
		sql.append(" 	                sum(Credit_Cnt)  as Credit_Cnt,                                                                            ");
		sql.append("                     Round(sum(Credit_Bal)/").append(unit).append(",0)  as Credit_Bal,                                                              ");
		sql.append("                     Sum(OverCreditMonth_Cnt)  as OverCreditMonth_Cnt,                                                       ");
		sql.append("                     Round(sum(OverCreditMonth_Amt)/").append(unit).append(",0)   as  OverCreditMonth_Amt,                                          ");
		sql.append("                     Sum(OverCredit_Cnt)   as OverCredit_Cnt,                                                                ");
		sql.append("                     Round(sum(OverCredit_Bal)/").append(unit).append(",0)  as  OverCredit_Bal                                                      ");
		sql.append("  			from WLX07_M_CREDIT                                                                                                  ");
		sql.append("  			where to_char(WLX07_M_CREDIT.m_year * 100 + WLX07_M_CREDIT.m_month) >=  ?                                         ");
		sql.append("  		    and to_char(WLX07_M_CREDIT.m_year * 100 + WLX07_M_CREDIT.m_month) <=  ?                                         ");
		sql.append("  		    group by bank_no) a01  on  bn01.bank_no = a01.bank_no                                                              ");
		sql.append(" ) PreYearSet                                                                                                                ");
		sql.append(" left join                                                                                                                   ");
		sql.append(" ( select ' '    as  hsien_id ,");
		sql.append(" 	     '合計' as hsien_name, ");
		sql.append("  		 '999'  as  FR001W_output_order,  ");
		sql.append("  	     sum(CreditMonth_Cnt)                 as  CreditMonth_Cnt,                                                           ");
		sql.append("          Round(sum(CreditMonth_Amt)/").append(unit).append(",0)      as  CreditMonth_Amt,                                                          ");
		sql.append("          sum(OverCreditMonth_Cnt)             as  OverCreditMonth_Cnt,                                                      ");
		sql.append("          Round(sum(OverCreditMonth_Amt)/").append(unit).append(",0)  as  OverCreditMonth_Amt                                                       ");
		sql.append("   from  (select * from ").append(cd01Table).append(" cd01 where cd01.hsien_id <> 'Y') cd01                                                                ");
		sql.append("   left join wlx01 on wlx01.hsien_id=cd01.hsien_id and wlx01.m_year = ?                                                    ");
		sql.append("   left join bn01 on wlx01.bank_no=bn01.bank_no  and bn01.bank_type=? and bn01.m_year = wlx01.m_year and bn01.m_year = ?");
		sql.append("   left join (select bank_no,  ");
		sql.append("              sum(CreditMonth_Cnt)      as  CreditMonth_Cnt, ");
		sql.append(" 	          Round(sum(CreditMonth_Amt)/").append(unit).append(",0)  as CreditMonth_Amt,                                                       ");
		sql.append(" 	          Sum(CreditYear_Cnt_Acc)   as  CreditYear_Cnt_Acc,                                                          ");
		sql.append(" 	          Round(sum(CreditYear_Amt_Acc)/").append(unit).append(",0)  as CreditYear_Amt_Acc,                                                 ");
		sql.append(" 	                sum(Credit_Cnt)/").append(unit).append(",0  as Credit_Cnt,                                                                        ");
		sql.append("                     Round(sum(Credit_Bal)/").append(unit).append(",0)  as Credit_Bal,                                                              ");
		sql.append("                     Sum(OverCreditMonth_Cnt)  as OverCreditMonth_Cnt,                                                       ");
		sql.append("                     Round(sum(OverCreditMonth_Amt)/").append(unit).append(",0)   as  OverCreditMonth_Amt,                                          ");
		sql.append("                     Sum(OverCredit_Cnt)   as OverCredit_Cnt,                                                                ");
		sql.append("                     Round(sum(OverCredit_Bal)/").append(unit).append(",0)  as  OverCredit_Bal                                                      ");
		sql.append("  			from WLX07_M_CREDIT                                                                                                  ");
		sql.append("  			where to_char(WLX07_M_CREDIT.m_year * 100 + WLX07_M_CREDIT.m_month) >=  ?                                         ");
		sql.append("  			and to_char(WLX07_M_CREDIT.m_year * 100 + WLX07_M_CREDIT.m_month) <=  ?                                           ");
		sql.append("  			group by bank_no) a01  on  bn01.bank_no = a01.bank_no                                                                ");
		sql.append(" ) LastYearSet                                                                                                               ");
		sql.append(" on PreYearSet.hsien_id=LastYearSet.hsien_id                                                                                 ");
		sql.append(" order by fr001w_output_order                                                                                                ");
		
		return sql.toString();
	}
	/***
	 * 取得報表明細SQL.
	 * 
	 * @return
	 */
    private static String getReportDetailSQL(String u_year,List bank_list) {
    	StringBuffer  sql = new StringBuffer();
    	String cd01Table = "cd01" ;
    	if("99".equals(u_year)) {
    		cd01Table = "cd01_99" ;
    	}
    	
		sql.append(" select PreYearSet.hsien_id  	as  hsien_id,                                 ");
    	sql.append("        PreYearSet.hsien_name	as	hsien_name,                               ");
    	sql.append("        PreYearSet.FR001W_output_order	as  FR001W_output_order,              ");
    	sql.append("        PreYearSet.bank_no				as  bank_no,                          ");
    	sql.append("        PreYearSet.bank_name 			as  bank_name,                        ");
    	sql.append("        PreYearSet.CreditMonth_Cnt		as 	CreditMonth_Cnt,                  ");
    	sql.append("        PreYearSet.CreditMonth_Amt		as	CreditMonth_Amt,                  ");
    	sql.append("        PreYearSet.OverCreditMonth_Cnt	as	OverCreditMonth_Cnt,              ");
    	sql.append("        PreYearSet.OverCreditMonth_Amt	as	OverCreditMonth_Amt,              ");
    	sql.append("        LastYearSet.CreditMonth_Cnt		as	Last_CreditMonth_Cnt,             ");
    	sql.append("        LastYearSet.CreditMonth_Amt		as	Last_CreditMonth_Amt,             ");
    	sql.append("        LastYearSet.OverCreditMonth_Cnt	as	Last_OverCreditMonth_Cnt,         ");
    	sql.append("        LastYearSet.OverCreditMonth_Amt	as	Last_OverCreditMonth_Amt          ");
    	sql.append(" from                                                                                      ");
    	sql.append(" (select * from (                                                                          ");
    	sql.append("  select nvl(cd01.hsien_id,' ')       as  hsien_id ,                                       ");
    	sql.append("         nvl(cd01.hsien_name,'OTHER') as  hsien_name,                                      ");
    	sql.append("         cd01.FR001W_output_order     as  FR001W_output_order,                             ");
    	sql.append("         bn01.bank_no ,  bn01.BANK_NAME , CreditMonth_Cnt,                                 ");
    	sql.append("         Round(CreditMonth_Amt/?,0)   as CreditMonth_Amt,                                  ");
    	sql.append("         CreditYear_Cnt_Acc,                                                               ");
    	sql.append("         Round(CreditYear_Amt_Acc/?,0)  as CreditYear_Amt_Acc,                             ");
    	sql.append("         Credit_Cnt,                                                                       ");
    	sql.append("         Round(Credit_Bal/?,0)  as Credit_Bal,                                             ");
    	sql.append("         OverCreditMonth_Cnt,                                                              ");
    	sql.append("         Round(OverCreditMonth_Amt/?,0)   as  OverCreditMonth_Amt,                         ");
    	sql.append("         OverCredit_Cnt,                                                                   ");
    	sql.append("         Round(OverCredit_Bal/?,0)  as  OverCredit_Bal                                     ");
    	sql.append("   from  (select * from ").append(cd01Table).append(" cd01  where cd01.hsien_id <> 'Y') cd01                              ");
    	sql.append("   left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id    ");
    	sql.append("   left join (select * from bn01 where bank_type=? and m_year=?)bn01 on wlx01.bank_no=bn01.bank_no  ");
    	sql.append("   left join (select bank_no,                                                                                               ");
    	sql.append("                     sum(CreditMonth_Cnt)      as  CreditMonth_Cnt,                                                         ");
    	sql.append("                     Round(sum(CreditMonth_Amt)/1,0)  as CreditMonth_Amt,                                                   ");
    	sql.append("                     Sum(CreditYear_Cnt_Acc) as CreditYear_Cnt_Acc,                                                         ");
    	sql.append("                     Round(sum(CreditYear_Amt_Acc)/1,0)  as CreditYear_Amt_Acc,                                             ");
    	sql.append("                     sum(Credit_Cnt) as Credit_Cnt,                                                                         ");
    	sql.append("                     Round(sum(Credit_Bal)/1,0)  as Credit_Bal,                                                             ");
    	sql.append("                     Sum(OverCreditMonth_Cnt)  as OverCreditMonth_Cnt,                                                      ");
    	sql.append("                     Round(sum(OverCreditMonth_Amt)/1,0)   as  OverCreditMonth_Amt,                                         ");
    	sql.append("                     Sum(OverCredit_Cnt)   as OverCredit_Cnt,                                                               ");
    	sql.append("                     Round(sum(OverCredit_Bal)/1,0)  as  OverCredit_Bal                                                     ");
    	sql.append("              from WLX07_M_CREDIT                                                                                           ");
    	sql.append("              where to_char(WLX07_M_CREDIT.m_year * 100 + WLX07_M_CREDIT.m_month) >=  ?                                  ");
    	sql.append("              and   to_char(WLX07_M_CREDIT.m_year * 100 + WLX07_M_CREDIT.m_month) <=  ?                                  ");
    	sql.append("              group by bank_no) a01  on  bn01.bank_no = a01.bank_no                                                         ");
    	sql.append("  ) a01 where a01.bank_no  <>  ' '  and a01.bank_no in (").append(getSelectBankNo(bank_list)).append(") and  a01.CreditYear_Amt_Acc >= 0               ");
    	sql.append("  order by  a01.FR001W_output_order, a01.hsien_id,  a01.bank_no                                                             ");
    	sql.append(" ) PreYearSet   ");//--查詢年度區間 
    	sql.append(" left join                                                                                                                  ");
    	sql.append(" (select  * from (                                                                                                          ");
    	sql.append("  select nvl(cd01.hsien_id,' ')       as  hsien_id ,                                                                        ");
    	sql.append("         nvl(cd01.hsien_name,'OTHER') as  hsien_name,                                                                       ");
    	sql.append("         cd01.FR001W_output_order     as  FR001W_output_order,                                                              ");
    	sql.append("         bn01.bank_no ,  bn01.BANK_NAME , CreditMonth_Cnt,                                                                  ");
    	sql.append("         Round(CreditMonth_Amt/?,0)  as CreditMonth_Amt,                                                                    ");
    	sql.append("         CreditYear_Cnt_Acc,                                                                                                ");
    	sql.append("         Round(CreditYear_Amt_Acc/?,0)  as CreditYear_Amt_Acc,                                                              ");
    	sql.append("         Credit_Cnt,                                                                                                        ");
    	sql.append("         Round(Credit_Bal/?,0)  as Credit_Bal,                                                                              ");
    	sql.append("         OverCreditMonth_Cnt,                                                                                               ");
    	sql.append("         Round(OverCreditMonth_Amt/?,0)   as  OverCreditMonth_Amt,                                                          ");
    	sql.append("         OverCredit_Cnt,                                                                                                    ");
    	sql.append("         Round(OverCredit_Bal/?,0)  as  OverCredit_Bal                                                                      ");
    	sql.append("  from  (select * from ").append(cd01Table).append(" cd01 where cd01.hsien_id <> 'Y') cd01   ");
    	sql.append("  left join (select * from wlx01 where m_year=?)wlx01 on wlx01.hsien_id=cd01.hsien_id  ");
    	sql.append("  left join (select * from bn01 where bank_type=? and m_year=?)bn01 on wlx01.bank_no=bn01.bank_no ");
    	sql.append("  left join (select bank_no, ");
    	sql.append("                    sum(CreditMonth_Cnt)      as  CreditMonth_Cnt,                                                          ");
    	sql.append("                    Round(sum(CreditMonth_Amt)/1,0)  as CreditMonth_Amt,                                                    ");
    	sql.append("                    Sum(CreditYear_Cnt_Acc)   as  CreditYear_Cnt_Acc,                                                       ");
    	sql.append("                    Round(sum(CreditYear_Amt_Acc)/1,0)  as CreditYear_Amt_Acc,                                              ");
    	sql.append("                    sum(Credit_Cnt) as Credit_Cnt,                                                                          ");
    	sql.append("                    Round(sum(Credit_Bal)/1,0)  as Credit_Bal,                                                              ");
    	sql.append("                    Sum(OverCreditMonth_Cnt)  as OverCreditMonth_Cnt,                                                       ");
    	sql.append("                    Round(sum(OverCreditMonth_Amt)/1,0)   as  OverCreditMonth_Amt,                                          ");
    	sql.append("                    Sum(OverCredit_Cnt)   as OverCredit_Cnt,                                                                ");
    	sql.append("                    Round(sum(OverCredit_Bal)/1,0)  as  OverCredit_Bal                                                      ");
    	sql.append("             from WLX07_M_CREDIT                                                                                            ");
    	sql.append("             where to_char(WLX07_M_CREDIT.m_year * 100 + WLX07_M_CREDIT.m_month) >=  ?                                   ");
    	sql.append("             and   to_char(WLX07_M_CREDIT.m_year * 100 + WLX07_M_CREDIT.m_month) <=  ?                                   ");
    	sql.append("             group by bank_no) a01 on  bn01.bank_no = a01.bank_no                                                           ");
    	sql.append("  ) a01  ");
    	sql.append("  where a01.bank_no  <>  ' '  and a01.bank_no in (").append(getSelectBankNo(bank_list)).append(") and  a01.CreditYear_Amt_Acc >= 0                     ");
    	sql.append(" order by  a01.FR001W_output_order, a01.hsien_id,  a01.bank_no                                                              ");
    	sql.append(" ) LastYearSet   "); //--前一年度區間   
    	//sql.append(" on PreYearSet.hsien_id=LastYearSet.hsien_id  ");
    	sql.append(" on PreYearSet.bank_no=LastYearSet.bank_no ");
    	
    	
    	return sql.toString();
    }
    /***
     * 回傳已選的單位.
     * 
     * @param bank_list
     * @return
     */
    private static String getSelectBankNo(List bank_list) {
    	List temp;
    	String bank_id ="" ;
    	for (int v = 0; v < bank_list.size(); v++) {
			temp = (List) (bank_list.get(v));
			if("".equals(bank_id)) {
				bank_id = "'"+String.valueOf(temp.get(0))+"'"; 
			}else {
				bank_id += ",'"+String.valueOf(temp.get(0))+"'"; 
			}
			
		}
    	return bank_id ;
    }
}
