package com.tradevan.util.report;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Calendar;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFooter;
import org.apache.poi.hssf.usermodel.HSSFPrintSetup;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import com.tradevan.util.Utility;
import com.tradevan.util.dao.DataObject;

public class TM008W {
	
	private static String RPT_FILE_NAME = "辦理情形統計表.xls";
	
	private static String[] RPT_BODY_COLUMN = {"bank_code" , "bank_name" , "wml01_status" , "cnt_name" };
	
	private static HSSFCellStyle[] RPT_BODY_CELL_STYLES = new HSSFCellStyle[RPT_BODY_COLUMN.length];
	
	public static TM008W getInstance() {
		return new TM008W();
	}
	
	public String createRpt(String userName , String accTrType , String applyDate , List rptData) {
		String errMsg = "";
		try {
			//D:\\J119\\boaf-workspace\\pboaf\\MIS\\xlsDir
			String xlsDir = Utility.getProperties("xlsDir");
			//D:\\J119\\boaf-workspace\\pboaf\\MIS\\reportDir
			String reportDir = Utility.getProperties("reportDir");
			initDir(xlsDir , reportDir);
			
			FileInputStream finput = new FileInputStream(xlsDir + System.getProperty("file.separator") + RPT_FILE_NAME);
			POIFSFileSystem fs = new POIFSFileSystem(finput);
			HSSFWorkbook wb = new HSSFWorkbook(fs);
			HSSFSheet sheet = wb.getSheetAt(0);
			sheet.setAutobreaks(false);
			HSSFPrintSetup ps = sheet.getPrintSetup();
			ps.setScale( (short) 100); //列印縮放百分比
			ps.setPaperSize( (short) 9); //設定紙張大小 A4
			finput.close();
			
			setRowCell(sheet, 1, 1, accTrType, HSSFCell.ENCODING_UTF_16);
			setRowCell(sheet, 3, 1, "申報基準日:" + applyDate , HSSFCell.ENCODING_UTF_16);
			Calendar c = Calendar.getInstance();
			String printDate = "列印日期:" + (c.get(Calendar.YEAR)-1911)+"年"+(c.get(Calendar.MONTH)+1)+"月"+c.get(Calendar.DAY_OF_MONTH)+"日" +" " + Utility.getDateFormat("hh:mm:ss");
			setRowCell(sheet, 4, 2, printDate , HSSFCell.ENCODING_UTF_16);
			setRowCell(sheet, 5, 2, "列印人員:" + userName , HSSFCell.ENCODING_UTF_16);
			setBodyCellStyles(wb);
			setBody(sheet , rptData);
			String FileOutPutPath = reportDir + System.getProperty("file.separator") + RPT_FILE_NAME;
			outputReport(FileOutPutPath , sheet , wb);
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("TM008W.createRpt Error:" + e + e.getMessage());
		}
		return errMsg;
	}
	private void initDir(String... files) throws Exception {
		if(files != null) {
			for(String file : files) {
				File f = new File(file);
				if (!f.exists()) {
					Utility.mkdirs(file);
				}
			}
		}
	}
	private void setRowCell(HSSFSheet sheet , int rowNumber , int cellNumber , String cellValue , short cellStyle) {
		HSSFRow row = sheet.getRow(rowNumber);
		HSSFCell cell=row.getCell((short)cellNumber);
		cell.setEncoding(cellStyle);
		cell.setCellValue(cellValue);
	}
	private void outputReport(String FileOutPutPath , HSSFSheet sheet , HSSFWorkbook wb) throws IOException {
		FileOutputStream fout = new FileOutputStream(FileOutPutPath);
		HSSFFooter footer = sheet.getFooter();
		footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
		footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
		wb.write(fout);
		fout.close();
	}
	private void setBodyCellStyles(HSSFWorkbook wb) {
		reportUtil reportUtil = new reportUtil();
		HSSFCellStyle cs_right = reportUtil.getRightStyle(wb);
		HSSFCellStyle cs_center = reportUtil.getDefaultStyle(wb);
		HSSFCellStyle cs_left = reportUtil.getLeftStyle(wb);
		RPT_BODY_CELL_STYLES[0] = cs_center;
		RPT_BODY_CELL_STYLES[1] = cs_left;
		RPT_BODY_CELL_STYLES[2] = cs_center;
		RPT_BODY_CELL_STYLES[3] = cs_center;
	}
	private void setBody(HSSFSheet sheet , List rtpData) {
		int bodyStartRow = 7;
		if (rtpData != null && rtpData.size() != 0) {
			HSSFRow row = null;
			HSSFCell cell = null;
			for(int i = 0 ; i < rtpData.size() ; i++) {
				row = sheet.createRow(bodyStartRow);
				DataObject bean = (DataObject)rtpData.get(i);
			
				for(int j = 0 ; j < RPT_BODY_COLUMN.length ; j++ ) {
					String column = RPT_BODY_COLUMN[j];
					cell = row.createCell((short)(j + 1));
					cell.setEncoding(HSSFCell.ENCODING_UTF_16);
					cell.setCellValue(Utility.getTrimString(bean.getValue(column)));
					cell.setCellStyle(RPT_BODY_CELL_STYLES[j]);
				}
				bodyStartRow++;
			}
		}
	}
	public static void main(String[] args) {
		TM008W util = TM008W.getInstance();
		List dbData = getTestData();
		String accTrType = "";
		String applyDate = "";
		String userName = "";
		String retMsg = util.createRpt(userName , accTrType , applyDate , dbData);
		System.out.println(retMsg);
	}
	
	private static List getTestData() {
		List dbData = new LinkedList();
		DataObject bean1 = new DataObject();
		bean1.setValue("bank_code", "0130774");
		bean1.setValue("bank_name", "國泰世華銀行八德分行");
		bean1.setValue("applydate", "105/08/15");
		bean1.setValue("wml01_status", "未申報");
		bean1.setValue("cnt_name", "");
		dbData.add(bean1);
		
		DataObject bean2 = new DataObject();
		bean2.setValue("bank_code", "5030019");
		bean2.setValue("bank_name", "基隆市基隆區漁會信用部");
		bean2.setValue("applydate", "105/08/15");
		bean2.setValue("wml01_status", "未申報");
		bean2.setValue("cnt_name", "02-24623111");
		dbData.add(bean2);
		
		DataObject bean3 = new DataObject();
		bean3.setValue("bank_code", "6030016");
		bean3.setValue("bank_name", "基隆市農會信用部");
		bean3.setValue("applydate", "105/08/15");
		bean3.setValue("wml01_status", "已申報");
		bean3.setValue("cnt_name", "王小明 2312312");
		dbData.add(bean3);
		return dbData;
	}
}
