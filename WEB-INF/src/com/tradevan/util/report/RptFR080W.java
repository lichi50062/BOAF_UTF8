/*
//107.05.18 create 洗錢關鍵字報表-依縣市別 by 6417
*/
package com.tradevan.util.report;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
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

public class RptFR080W {

	private final static String ALL_CITY = "ALL";

	public static String createRpt(String s_year, String s_month, String e_year, String e_month, String city_type, String[] ex_content) {

		String errMsg = "";
		try {
			// get directory
			File xlsDir = new File(Utility.getProperties("xlsDir"));
			File reportDir = new File(Utility.getProperties("reportDir"));
			if (xlsDir.exists() == false) {
				if (!Utility.mkdirs(Utility.getProperties("xlsDir"))) {
					errMsg += Utility.getProperties("xlsDir") + "目錄新增失敗";
				}
			}
			if (reportDir.exists() == false) {
				if (!Utility.mkdirs(Utility.getProperties("reportDir"))) {
					errMsg += Utility.getProperties("reportDir") + "目錄新增失敗";
				}
			}

			// chose file
			String openfile;
			boolean isAllCity = ALL_CITY.equals(city_type);
			System.out.println("isAllCity :" + isAllCity);
			if (isAllCity) {
				openfile = "洗錢關鍵字_依縣市別_總表.xls";
			} else {
				openfile = "洗錢關鍵字_依縣市別.xls";

			}
			System.out.println("open file " + openfile);
			HSSFWorkbook workbook = null;
			try {
				// 讀取Excel範本
				FileInputStream fileInput = new FileInputStream(xlsDir + System.getProperty("file.separator") + openfile);
				System.out.println("Open " + openfile + " 範本檔成功");
				POIFSFileSystem fileSystem = new POIFSFileSystem(fileInput);
				// 創建工作簿
				workbook = new HSSFWorkbook(fileSystem);
				System.out.println("Open " + openfile + " 工作表成功");
				fileInput.close();

			} catch (FileNotFoundException e) {
				errMsg = "檔案不存在";
			}
			reportUtil report = new reportUtil();
			// 預設格式
			HSSFCellStyle cs_left = report.getLeftStyle(workbook);
			HSSFCellStyle cs_center = report.getDefaultStyle(workbook);
			HSSFCellStyle cs_keyword = report.getDefaultStyle(workbook);
			cs_keyword.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
			cs_keyword.setFillForegroundColor(HSSFColor.RED.index);
			// 關鍵字
			List<String> keyWords = new ArrayList<String>();
			for (int i = 0; i < ex_content.length; i++) {
				if (!"".equals(ex_content[i])) {
					keyWords.add(ex_content[i]);
				}
			}
			// get data
			List<Object> dbData = getData(city_type, s_year, s_month, e_year, e_month, isAllCity, keyWords);

			HSSFSheet sheet = null;
			HSSFCell cell = null;
			// write data
			DataObject bean = null;
			int rowNum = 0;

			// 判斷是查詢個別縣市or總表
			if (isAllCity) {
				String sheetName = null;
				int sheetSize = workbook.getNumberOfSheets();
				System.out.println("sheetSize=" + sheetSize);
				for (int sheetAt = 0; sheetAt < sheetSize; sheetAt++) {
					rowNum = 3;
					sheetName = workbook.getSheetName(sheetAt);
					sheet = workbook.getSheetAt(sheetAt);
					if (sheet == null) {
						System.out.println("open sheet " + sheetName + " 失敗");
					} else {
						System.out.println("open sheet " + sheetName + " 成功");
					}
					sheet.setAutobreaks(false);// 調整單元格寬度
					HSSFPrintSetup ps = sheet.getPrintSetup();
					ps.setScale((short) 68); // 列印縮放百分比
					ps.setPaperSize(HSSFPrintSetup.A4_PAPERSIZE); // 設定紙張大小 A4
					ps.setLandscape(true);// true:橫,false:縱
					printDate(sheet, workbook);

					if (dbData.size() != 0) {
						if ("總表".equals(sheetName)) {
							for (int i = 0; i < dbData.size(); i++) {
								bean = (DataObject) dbData.get(i);
								RptFR080W.writeData(bean, sheet, cell, rowNum++, cs_left, cs_center);
							}
						} else {
							for (int i = 0; i < dbData.size(); i++) {
								bean = (DataObject) dbData.get(i);
								// 寫入與表名相同的資料
								if (bean.getValue("hsien_name").equals(sheetName) || bean.getValue("bank_name").equals(sheetName)) {
									RptFR080W.writeData(bean, sheet, cell, rowNum++, cs_left, cs_center);
								}
							}
						}
					} else {
						errMsg = "無報表資料";
					}
					if (keyWords.size() != 0) {
						RptFR080W.writeKeywordData(keyWords, sheet, rowNum++, cs_keyword);
					}
					setFooter(sheet);

				}
			} else {
				rowNum = 3;
				String hsienName = getHsienName(city_type);
				// update sheet name and cell name
				sheet = workbook.getSheetAt(0);
				if (!"".equals(hsienName)) {
					workbook.setSheetName(0, hsienName);
				}
				
				if (sheet == null) {
					System.out.println("open sheet 縣市 失敗");
				} else {
					System.out.println("open sheet 縣市 成功");
				}
				sheet.setAutobreaks(false);// 調整單元格寬度
				HSSFPrintSetup ps = sheet.getPrintSetup();
				ps.setScale((short) 74); // 列印縮放百分比
				ps.setPaperSize(HSSFPrintSetup.A4_PAPERSIZE); // 設定紙張大小 A4
				ps.setLandscape(true);// true:橫,false:縱
				printDate(sheet, workbook);
				HSSFCell titleCell = sheet.getRow(0).getCell((short) 0);
				titleCell.setCellValue("「MIS檢查追蹤管理系統」 涉及洗錢等關鍵字﹝註1﹞之檢查意見─" + hsienName);
				// write data to excel
				if (dbData.size() != 0) {
					for (int i = 0; i < dbData.size(); i++) {
						bean = (DataObject) dbData.get(i);
						RptFR080W.writeData(bean, sheet, cell, rowNum++, cs_left, cs_center);
					}
				} else {
					errMsg = "無報表資料";
				}
				if (keyWords.size() != 0) {
					RptFR080W.writeKeywordData(keyWords, sheet, rowNum++, cs_keyword);
				}
				setFooter(sheet);
			}
			// write file to reportDir
			FileOutputStream outputStream = new FileOutputStream(Utility.getProperties("reportDir") + System.getProperty("file.separator") + openfile);
			workbook.write(outputStream);
			outputStream.close();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return errMsg;
	}

	/**
	 * 檢查意見資料寫入
	 * 
	 * @param bean
	 * @param sheet
	 * @param rowNum
	 * @param cs_left
	 * @param cs_center
	 * @return
	 */
	private static boolean writeData(DataObject bean, HSSFSheet sheet, HSSFCell cell, int rowNum, HSSFCellStyle cs_left, HSSFCellStyle cs_center) {
		boolean result = false;
		HSSFRow row = sheet.createRow(rowNum);
		for (int column = 0; column < 10; column++) {
			String value = "";
			row = sheet.createRow(rowNum);
			cell = row.createCell((short) column);
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			if (column == 0) {
				value = bean.getValue("bank_no") == null ? "" : (String) bean.getValue("bank_no");
				cell.setCellStyle(cs_center);
			} else if (column == 1) {
				value = bean.getValue("bank_name") == null ? "" : (String) bean.getValue("bank_name");
				cell.setCellStyle(cs_center);
			} else if (column == 2) {
				value = bean.getValue("reportno") == null ? "" : (String) bean.getValue("reportno");
				cell.setCellStyle(cs_center);
			} else if (column == 3) {
				value = bean.getValue("ch_type") == null ? "" : (String) bean.getValue("ch_type");
				cell.setCellStyle(cs_center);
			} else if (column == 4) {
				value = bean.getValue("base_date") == null ? "" : (String) bean.getValue("base_date");
				cell.setCellStyle(cs_center);
			} else if (column == 5) {
				value = bean.getValue("item_no") == null ? "" : (String) bean.getValue("item_no");
				cell.setCellStyle(cs_center);
			} else if (column == 6) {
				value = bean.getValue("ex_content") == null ? "" : (String) bean.getValue("ex_content");
				cell.setCellStyle(cs_left);
			} else if (column == 7) {
				value = bean.getValue("commentt") == null ? "" : (String) bean.getValue("commentt");
				cell.setCellStyle(cs_left);
			} else if (column == 8) {
				value = bean.getValue("digest") == null ? "" : (String) bean.getValue("digest");
				cell.setCellStyle(cs_left);
			} else if (column == 9) {
				value = bean.getValue("cmuse_name") == null ? "" : (String) bean.getValue("cmuse_name");
				cell.setCellStyle(cs_left);
			}
			cell.setCellValue(value);
		}

		return result;
	}

	/**
	 * 關鍵字資料寫入
	 * 
	 * @param exCotentList
	 *            關鍵字
	 * @param sheet
	 * @param rowNum
	 * @param cs_left
	 * @return
	 */
	private static boolean writeKeywordData(List<String> exCotentList, HSSFSheet sheet, int rowNum, HSSFCellStyle cs_left) {
		boolean result = false;
		int newRowNum = rowNum + 4;// 空四行
		int countNum = 1;
		int cellCount = 0;// cell 位置
		HSSFRow row = sheet.createRow(newRowNum++);
		HSSFCell cell = row.createCell((short) cellCount);
		cell.setCellValue("〔註1〕關鍵字");
		cs_left.setFillForegroundColor(HSSFColor.TAN.index);
		// 關鍵字每四個換行
		if (exCotentList.size() != 0) {
			row = sheet.createRow(newRowNum++);
			for (int i = 0; i < exCotentList.size(); i++) {
				cell = row.createCell((short) cellCount++);
				cell.setCellStyle(cs_left);
				cell.setCellValue(exCotentList.get(i));
				if (countNum++ % 4 == 0) {
					row = sheet.createRow(newRowNum++);
					cellCount = 0;
				}
			}
		}
		return result;
	}

	private static void setFooter(HSSFSheet sheet) {
		HSSFFooter footer = sheet.getFooter();
		footer.setCenter("Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages());
		footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
	}

	private static void printDate(HSSFSheet sheet, HSSFWorkbook workbook) {
		HSSFRow row = sheet.getRow(0);
		String printDate = Utility.getDateFormat("yyyy/MM/dd");
		HSSFCell cell = row.createCell((short) 7);
		HSSFCellStyle style = workbook.createCellStyle();
		style.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
		cell.setCellStyle(style);
		cell.setCellValue("列印日期：" + Utility.getCHTdate(printDate, 1));
	}


	private static String getHsienName(String city_type) {
		// SQL 取得縣市名稱
		StringBuffer sql2 = new StringBuffer();
		List<Object> dbData2 = new ArrayList<Object>();
		ArrayList<String> paramList2 = new ArrayList<String>();
		sql2.append(" SELECT hsien_name FROM V_BANK_LOCATION WHERE  hsien_id =? and ROWNUM = 1");
		paramList2.add(city_type);
		dbData2 = DBManager.QueryDB_SQLParam(sql2.toString(), paramList2, "HSIEN_NAME");
		String hsienName = "";
		for (int i = 0; i < dbData2.size(); i++) {
			hsienName = (String) ((DataObject) dbData2.get(i)).getValue("hsien_name");
		}

		return hsienName;
	}

	private static List<Object> getData(String city_type, String s_year, String s_month, String e_year, String e_month, boolean isAllCity, List<String> keyWords) {
		StringBuffer sql = new StringBuffer();
		sql.setLength(0);
		List<Object> dbData = new ArrayList<Object>();
		ArrayList<String> paramList = new ArrayList<String>();
		sql.append(" SELECT");
		sql.append(" c.hsien_id,");
		sql.append(" hsien_name, ");
		sql.append(" b.bank_no, ");
		sql.append(" c.bank_name, ");
		sql.append(" a.reportno,");
		sql.append(" d.cmuse_name as ch_type, ");
		sql.append(" ((TO_CHAR(b.base_date,'yyyy')-1911)||'/'|| TO_CHAR(b.base_date,'mm/dd'))  base_date, ");
		sql.append(" a.reportno_seq, ");
		sql.append(" a.item_no,");
		sql.append(" a.serial,");
		sql.append(" a.ex_content,");
		sql.append(" a.commentt,");
		sql.append(" e.digest,");
		sql.append(" (select cmuse_name from CDShareNo WHERE a.audit_result = cmuse_id AND cmuse_div ='026' )  cmuse_name ");
		sql.append(" FROM ExDefGoodF a, ExReportF b, (select * from v_bank_location where m_year=100) c, CDShareNO d , ExDG_HistoryF e ");
		sql.append(" WHERE a.reportno(+) = b.reportno AND b.bank_no = c.bank_no ");
		sql.append(" and  a.reportno = e.reportno AND a.reportno_seq = e.reportno_seq(+) ");
		sql.append(" AND (b.ch_type = d.cmuse_id AND cmuse_div = '023') ");
		sql.append(" AND TO_CHAR(b.base_date, 'yyyymm') BETWEEN ? AND ? ");
		paramList.add(s_year + s_month);
		paramList.add(e_year + e_month);
		// 查詢個別縣市
		if (!isAllCity) {
			sql.append(" and hsien_id= ? ");
			paramList.add(city_type);
		}
		if (keyWords.size() != 0) {
			sql.append(" and (");
			for (int i = 0; i < keyWords.size(); i++) {
				if (i != 0) {
					sql.append(" or a.ex_content like ? ");
				} else {
					sql.append(" a.ex_content like ? ");
				}
				paramList.add("%" + keyWords.get(i) + "%");
			}
			sql.append(" ) ");
		}
		sql.append(" order by fr001w_output_order,bank_no,reportno,item_no");

		dbData = DBManager.QueryDB_SQLParam(sql.toString(), paramList, "HSIEN_ID, HSIEN_NAME, BANK_NO, BANK_NAME, REPORTNO, CH_TYPE, BASE_DATE, ORG_BASE_DATE, REPORTNO_SEQ, ITEM_NO, SERIAL, EX_CONTENT, COMMENTT, DIGEST, CMUSE_NAME");
		System.out.println("dbData.size()=" + dbData.size());
		return dbData;
	}

}