/*
  94.03.11 fix 金額資料為零是不輸出0，改為輸出空白及欄位資料右靠處理
 110.03.12 fix 110年4月份套用新格式 by 2295
 110.05.13 fix 調整110/5套用新的格式 by 2295
 */
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.text.*;
import java.util.*;

import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;
/**
 * @author 2295
 * 
 * TODO To change the template for this generated type comment go to Window -
 * Preferences - Java - Code Style - Code Templates
 */
public class FR009W_Excel {
	final private static String[] table = {"920101", "920201", "920301", "920401", "920501", "920601", "space", "92071N", "92071P", "920710", "92072N", "92072P", "920720", "92073N", "92073P", "920730", "92074N", "92074P", "920740", "92075N", "92075P",
			"920750", "920801", "929901"};
	final private static String[] table_11004 = {"920101", "920201", "920301", "920401", "920501", "920601","920901","920902", "space", "92071N", "92071P", "920710", "92072N", "92072P", "920720", "92073N", "92073P", "920730", "92074N", "92074P", "920740", "92075N", "92075P",
		"920750", "920801", "929901"};
	final private static double[] risk      =  {0, 0, 0, 0.1, 0.2, 0.5, 0, 0, 0, 0, 0, 0, 1, 1};
	final private static double[] risk_11004 = {0, 0, 0, 0.1, 0.2, 0.45, 0.75, 1, 0, 0, 0, 0, 0, 0, 1, 1};

	public static String createRpt(String S_YEAR, String S_MONTH, String BANK_NAME, HashMap h, boolean isEmpty) {
		String errMsg = "";

		try {
			System.out.println("信用部淨值占風險性資產比率計算表.xls");
			File xlsDir = new File(Utility.getProperties("xlsDir"));
			File reportDir = new File(Utility.getProperties("reportDir"));
			System.out.println("xlsDir Path = " + xlsDir.getPath());
			System.out.println("reportDir Path = " + reportDir.getPath());
			String[] table_now = table;
			double[] risk_now = risk;
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
			FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator") + "信用部淨值占風險性資產比率計算表"+((Integer.parseInt(S_YEAR+S_MONTH) >= 11005)?"_11004":"")+".xls");

			// 設定FileINputStream讀取Excel檔
			POIFSFileSystem fs = new POIFSFileSystem(finput);
			HSSFWorkbook wb = new HSSFWorkbook(fs);
			HSSFSheet sheet = wb.getSheetAt(0);// 讀取第一個工作表，宣告其為sheet
			HSSFPrintSetup ps = sheet.getPrintSetup(); // 取得設定
			// sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
			// sheet.setAutobreaks(true); //自動分頁

			// 設定頁面符合列印大小
			sheet.setAutobreaks(false);
			ps.setScale((short) 80); // 列印縮放百分比

			ps.setPaperSize((short) 9); // 設定紙張大小 A4
			// wb.setSheetName(0,"test");
			finput.close();

			HSSFRow row = null;// 宣告一列
			HSSFCell cell = null;// 宣告一個儲存格

			String title = "";
			if (!isEmpty) {
				title = S_YEAR + "年" + S_MONTH + "月" + BANK_NAME + "淨值占風險性資產比率計算表";
			}
			else {
				title = S_YEAR + "年" + S_MONTH + "月 無資料存在";
			}

			short rowNo = 0;
			row = sheet.getRow(rowNo);
			cell = row.getCell((short) 1);
			// 設定這個儲存格的字串要儲存雙位元,中文才能正常顯示================
			cell.setEncoding(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(toBig5(title));

			String amt = "";
			rowNo = 2;
			if (!isEmpty) {
				if(Integer.parseInt(S_YEAR+S_MONTH) >= 11005){
					table_now = table_11004;
					risk_now = risk_11004;
				}
				for (int i = 0; i < table_now.length; i++) {
					System.out.println("rowNo="+rowNo+",i=" + i+",table_now["+i+"]="+table_now[i]);
					if (table_now[i].indexOf("N") > 0) {
						row = sheet.getRow(rowNo);
						amt = (String) h.get(table_now[i]);
						cell = row.getCell((short) 1);
						cell.setEncoding(HSSFCell.ENCODING_UTF_16);
						cell.setCellValue(toBig5(amt));
						//System.out.println("rowNo="+rowNo+",test-N="+toBig5(amt));
					}
					else if (table_now[i].indexOf("P") > 0) {
						
						row = sheet.getRow(rowNo);
						amt = (String) h.get(table_now[i]);
						risk_now[rowNo - 2] = toDouble(amt);						
						cell = row.getCell((short) 2);
						if (amt.equals("0") || amt.equals("0.0") ) {
							cell.setCellValue(toFormatWithoutZero(amt));
							//System.out.println("rowNo="+rowNo+",test-P="+toFormatWithoutZero(amt));
						}
						else {
							cell.setCellValue(toDouble(amt)*100 + "%");
							//System.out.println("rowNo="+rowNo+",test-P="+toDouble(amt)*100 + "%");
						}
							
					}
					else {
						//System.out.println("test-count.rowNo="+rowNo);
						row = sheet.getRow(rowNo);
						
						amt = (String) h.get(table_now[i]);						
						cell = row.getCell((short) 3);
						cell.setCellValue(toFormatWithoutZero(amt));						
						double r = 0;
						if (table_now[i].equals("929901")) {
							r = toDouble((String) h.get("910500"));							
						}
						else {
							r = toDouble(amt) * (risk_now[rowNo - 2]);							
						}
						//System.out.println("rowNo="+rowNo+",amt="+toDouble(amt)+" * Risk: "+risk_now[rowNo - 2]+"="+toFormatWithoutZero(r));
						cell = row.getCell((short) 4);
						cell.setCellValue(toFormatWithoutZero(r));
						rowNo++;
					}

				}
			}
			FileOutputStream fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + "信用部淨值占風險性資產比率計算表.xls");
			System.out.println("p = " + reportDir + System.getProperty("file.separator") + "信用部淨值占風險性資產比率計算表.xls");
			wb.write(fout);
			// 儲存
			fout.close();
		} catch (Exception e) {
			System.out.println("createRpt Error:" + e + e.getMessage());
		}
		return errMsg;
	}

	private static double toDouble(String num) {
		try {
			return Double.parseDouble(num);
		} catch (Exception e) {
			return 0;
		}
	}

	private static String toFormatWithoutZero(String s) {
		String str = Utility.setCommaFormat(s);
		if(str.equals("0") || str.equals("0.0")) {
			str = "";
		}
		return str;
	}

	private static String toFormatWithoutZero(double d) {
		d = toRound(d, 4);
		MessageFormat mf = new MessageFormat("{0,number,##,###.####}");
		Object[] obj = {new Double(d)};
		String str = mf.format(obj);
		// String str = Utility.setCommaFormat(nf.format(d));
		if(str.equals("0") || str.equals("0.0")) {
			str = "";
		}
		return   str;
	}

	private static String toBig5(String str) throws Exception {
		// return new String(str.getBytes("BIG5"),"MS950");
		return str;
	}

	// 四捨五入 num 到小數以下 p 位數
	private static double toRound(double num, int p) {
		return (Math.floor(num * (Math.pow(10, p)) + 0.5)) / Math.pow(10, p);
	}

}
