/*
 *109.08.11 create  by 6493
 */
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.Region;

import java.io.*;
import java.util.*;

import com.tradevan.util.DownLoad;
import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;

public class RptFR083W {
 	
	 @SuppressWarnings({ "null", "rawtypes" })
	public static String createRpt(String M_YEAR, String M_MONTH,String unit,String bank_code,String bank_name, String bank_type ) {    

	    String errMsg = "";
	    List dbBaseDataList = null;
	    List dbItem5DataList = null;
	    List dbItem9DataList = null;
	    String sqlCmd = "";    
	    int rowNum=0;
	    DataObject bean = null;
	    reportUtil reportUtil = new reportUtil();
	    HSSFCellStyle cs_right_grey = null; 
	    HSSFCellStyle cs_center_grey = null; 
		HSSFCellStyle cs_right = null; 
		HSSFCellStyle cs_center = null;
		HSSFCellStyle cs_left = null;
		HSSFCellStyle nb_left = null;
		HSSFCellStyle nb_right = null;
		HSSFCellStyle nb_left_wrap_text = null;
	   
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
	      FileInputStream finput = null;
	      String filename="信用部電子銀行及行動支付業務辦理情形.xls";
	      //input the standard report form      
	      finput = new FileInputStream(xlsDir +System.getProperty("file.separator") +filename);
	      //設定FileINputStream讀取Excel檔
	      POIFSFileSystem fs = new POIFSFileSystem(finput);
	      HSSFWorkbook wb = new HSSFWorkbook(fs);
	      HSSFSheet sheet_1 = getSheet(wb, 0);//讀取第一個工作表
	      HSSFSheet sheet_2 = getSheet(wb, 1);//讀取第二個工作表
	      finput.close();

	      HSSFRow row = null; //宣告一列
	      
	      
	      
	      HSSFPalette palette = wb.getCustomPalette();
	      palette.setColorAtIndex(HSSFColor.GREY_25_PERCENT.index, (byte)217, (byte)217, (byte)217);
	      HSSFColor  lightGray = palette.getColor(HSSFColor.GREY_25_PERCENT.index);
	      
	      cs_right_grey = reportUtil.getRightStyle(wb);
	      cs_right_grey.setFillForegroundColor(lightGray.getIndex());
	      cs_right_grey.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
	      
	      cs_center_grey = reportUtil.getDefaultStyle(wb);
	      cs_center_grey.setFillForegroundColor(lightGray.getIndex());
	      cs_center_grey.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
	      
	      cs_right = reportUtil.getRightStyle(wb);
	      cs_center = reportUtil.getDefaultStyle(wb);
	      cs_left = reportUtil.getLeftStyle(wb);
	      
	      nb_left = reportUtil.getNoBorderLeftStyle(wb);
	      nb_left.setWrapText(false);
	      nb_left_wrap_text = reportUtil.getNoBorderLeftStyle(wb);
	      nb_left_wrap_text.setWrapText(true);
	      
	      nb_right = reportUtil.getNoBoderStyle(wb);
	      
	      //取得報表項目1~4,6~8資料，項目5,9為動態新增項目需另外取得
	      dbBaseDataList = getBaseData(M_YEAR, M_MONTH, bank_code);
	      System.out.println("dbData.size=" + dbBaseDataList.size());
	      //設定報表表頭資料============================================
	      setSheetTitle(sheet_1,M_YEAR, M_MONTH, bank_name, dbBaseDataList);
	      setSheetTitle(sheet_2,M_YEAR, M_MONTH, bank_name, dbBaseDataList);
	      
	      if (dbBaseDataList != null || dbBaseDataList.size() !=0) {
	    	  bean = (DataObject)dbBaseDataList.get(0);
	    	  //======================工作表1======================
	    	  //-----項目1:客戶數-----
	    	  //本月開戶數
	    	  row = sheet_1.createRow(5);
	    	  setCellNumberValue(wb, row, 1, cs_right, bean.getValue("field_900100_m"));//電子銀行
	    	  setCellNumberValue(wb, row, 2, cs_right, bean.getValue("field_900200_m"));//行動支付
	    	  setCellNumberValue(wb, row, 3, cs_right, bean.getValue("field_900300_m"));//收單特店-主掃模式
	    	  setCellNumberValue(wb, row, 4, cs_right, bean.getValue("field_900400_m"));//收單特店-被掃模式
	    	  //累計開戶數
	    	  row = sheet_1.createRow(6);
	    	  setCellNumberValue(wb, row, 1, cs_right_grey, bean.getValue("field_900100_y"));//電子銀行
	    	  setCellNumberValue(wb, row, 2, cs_right_grey, bean.getValue("field_900200_y"));//行動支付
	    	  setCellNumberValue(wb, row, 3, cs_right_grey, bean.getValue("field_900300_y"));//收單特店-主掃模式
	    	  setCellNumberValue(wb, row, 4, cs_right_grey, bean.getValue("field_900400_y"));//收單特店-被掃模式
	    	  
	    	  //-----項目2:約定帳戶轉帳交易情形-----
	    	  //本月交易
	    	  row = sheet_1.createRow(10);
	    	  setCellNumberValue(wb, row, 1, cs_right, bean.getValue("field_910101_m"));//筆數
	    	  setCellNumberValue(wb, row, 2, cs_right, bean.getValue("field_910102_m"));//金額
	    	  //本年累計交易
	    	  row = sheet_1.createRow(11);
	    	  setCellNumberValue(wb, row, 1, cs_right_grey, bean.getValue("field_910101_y"));//筆數
	    	  setCellNumberValue(wb, row, 2, cs_right_grey, bean.getValue("field_910102_y"));//金額

	    	  //-----項目3:非約定帳戶轉帳交易情形-----
	    	  //本月交易
	    	  row = sheet_1.createRow(15);
	    	  setCellNumberValue(wb, row, 1, cs_right, bean.getValue("field_910201_m"));//筆數
	    	  setCellNumberValue(wb, row, 2, cs_right, bean.getValue("field_910202_m"));//金額
	    	  //本年累計交易
	    	  row = sheet_1.createRow(16);
	    	  setCellNumberValue(wb, row, 1, cs_right_grey, bean.getValue("field_910201_y"));//筆數
	    	  setCellNumberValue(wb, row, 2, cs_right_grey, bean.getValue("field_910202_y"));//金額
	    	  
	    	  //-----項目4:臺灣Pay購物交易情形-----
	    	  //--轉帳購物--
	    	  //筆數
	    	  row = sheet_1.createRow(20);
	    	  setCellNumberValue(wb, row, 2, cs_right, bean.getValue("field_910311_m"));//本月交易
	    	  setCellNumberValue(wb, row, 3, cs_right, bean.getValue("field_910311_y"));//本年累計交易
	    	  //金額
	    	  row = sheet_1.createRow(21);
	    	  setCellNumberValue(wb, row, 2, cs_right, bean.getValue("field_910312_m"));//本月交易
	    	  setCellNumberValue(wb, row, 3, cs_right, bean.getValue("field_910312_y"));//本年累計交易
	    	  
	    	  //--消費扣款--
	    	  //筆數
	    	  row = sheet_1.createRow(22);
	    	  setCellNumberValue(wb, row, 2, cs_right, bean.getValue("field_910321_m"));//本月交易
	    	  setCellNumberValue(wb, row, 3, cs_right, bean.getValue("field_910321_y"));//本年累計交易
	    	  //金額
	    	  row = sheet_1.createRow(23);
	    	  setCellNumberValue(wb, row, 2, cs_right, bean.getValue("field_910322_m"));//本月交易
	    	  setCellNumberValue(wb, row, 3, cs_right, bean.getValue("field_910322_y"));//本年累計交易
	    	  
	    	  //--繳費--
	    	  //筆數
	    	  row = sheet_1.createRow(24);
	    	  setCellNumberValue(wb, row, 2, cs_right, bean.getValue("field_910331_m"));//本月交易
	    	  setCellNumberValue(wb, row, 3, cs_right, bean.getValue("field_910331_y"));//本年累計交易
	    	  //金額
	    	  row = sheet_1.createRow(25);
	    	  setCellNumberValue(wb, row, 2, cs_right, bean.getValue("field_910332_m"));//本月交易
	    	  setCellNumberValue(wb, row, 3, cs_right, bean.getValue("field_910332_y"));//本年累計交易
	    	  
	    	  //--P2P轉帳--
	    	  //筆數
	    	  row = sheet_1.createRow(26);
	    	  setCellNumberValue(wb, row, 2, cs_right, bean.getValue("field_910341_m"));//本月交易
	    	  setCellNumberValue(wb, row, 3, cs_right, bean.getValue("field_910341_y"));//本年累計交易
	    	  //金額
	    	  row = sheet_1.createRow(27);
	    	  setCellNumberValue(wb, row, 2, cs_right, bean.getValue("field_910342_m"));//本月交易
	    	  setCellNumberValue(wb, row, 3, cs_right, bean.getValue("field_910342_y"));//本年累計交易
	    	  
	    	  //-----項目5:電子銀行-動態新增項目-----
	    	  //取得報表動態新增項目5
	    	  String ncacno = "ncacno";
	 		  String ncacno_7 = "ncacno_7";
	 		  String t1 = ncacno;
	 		  if("6".equals(bank_type)){
	 			  t1 = ncacno;
	 		  }else if("7".equals(bank_type)){
	 			  t1 = ncacno_7;
	 		  }
		      HashMap<String, Object> item5CodeAndNameMap = getItem5CodeAndName(t1);
		      dbItem5DataList = getAddItemData(M_YEAR, M_MONTH, bank_code, item5CodeAndNameMap);
	    	  List item5CodeAndNameList = (List) item5CodeAndNameMap.get("codeAndNameList");
	    	  DataObject item5CodeAndName = null;
	    	  Region region = null;
	    	  String[] item5FieldNames;
	    	  int item5Index = 29;
	    	  int itemNumber = 4;
	    	  for(int i=0; i<item5CodeAndNameList.size(); i++){
	    		  itemNumber++;
	    		  item5CodeAndName = (DataObject)item5CodeAndNameList.get(i);
	    		  row = sheet_1.createRow( item5Index++ );
	    		  setCellValue(wb, row, 0, nb_left, itemNumber + "." + item5CodeAndName.getValue("limit_type_name"));//動態新增項目Title

	    		  row = sheet_1.createRow( item5Index++ );
	    		  setCellValue(wb, row, 0, cs_center, "");
	    		  setCellValue(wb, row, 1, cs_center, "筆數");
	    		  setCellValue(wb, row, 2, cs_center, "金額");
	    		  
	    		  item5FieldNames = ((String) item5CodeAndNameMap.get((String)item5CodeAndName.getValue("limit_type"))).split(",");
	    		  row = sheet_1.createRow( item5Index++ );
	    		  setCellValue(wb, row, 0, cs_center, "本月交易");
	    		  setCellNumberValue(wb, row, 1, cs_right, dbItem5DataList==null ? "0" : ((DataObject)dbItem5DataList.get(0)).getValue(item5FieldNames[0]));
	    		  setCellNumberValue(wb, row, 2, cs_right, dbItem5DataList==null ? "0" : ((DataObject)dbItem5DataList.get(0)).getValue(item5FieldNames[2]));
	    		  row = sheet_1.createRow( item5Index++ );
	    		  setCellValue(wb, row, 0, cs_center_grey, "本年累計交易");
	    		  setCellNumberValue(wb, row, 1, cs_right_grey, dbItem5DataList==null ? "0" : ((DataObject)dbItem5DataList.get(0)).getValue(item5FieldNames[1]));
	    		  setCellNumberValue(wb, row, 2, cs_right_grey, dbItem5DataList==null ? "0" : ((DataObject)dbItem5DataList.get(0)).getValue(item5FieldNames[3]));
	    		  
	    		  region = new Region(item5Index, (short)0, item5Index, (short)4);
	    		  sheet_1.addMergedRegion(region);
	    		  row = sheet_1.createRow( item5Index++ );
	    		  row.setHeight((short)(sheet_1.getDefaultRowHeight()*2));
	    		  setCellValue(wb, row, 0, nb_left_wrap_text, "註"+itemNumber+"：新系統上線第一次申報時，由農漁會鍵入，第二個月起，由系統自動計算，如無交易應填「0」。");
	    	  }
	    	  
	    	  //======================工作表2======================
	    	  //-----項目6:特店收單交易情形-----
	    	  itemNumber++;
	    	  row = sheet_2.createRow(2);
	    	  setCellValue(wb, row, 0, nb_left, itemNumber+".特店收單交易情形");
	    	  //本月交易
	    	  row = sheet_2.createRow(5);
	    	  setCellNumberValue(wb, row, 1, cs_right, bean.getValue("field_920101_m"));//主掃模式-筆數
	    	  setCellNumberValue(wb, row, 2, cs_right, bean.getValue("field_920102_m"));//主掃模式-金額
	    	  setCellNumberValue(wb, row, 3, cs_right, bean.getValue("field_920201_m"));//被掃模式-筆數
	    	  setCellNumberValue(wb, row, 4, cs_right, bean.getValue("field_920202_m"));//被掃模式-金額
	    	  //本年累計交易
	    	  row = sheet_2.createRow(6);
	    	  setCellNumberValue(wb, row, 1, cs_right, bean.getValue("field_920101_y"));//主掃模式-筆數
	    	  setCellNumberValue(wb, row, 2, cs_right, bean.getValue("field_920102_y"));//主掃模式-金額
	    	  setCellNumberValue(wb, row, 3, cs_right, bean.getValue("field_920201_y"));//被掃模式-筆數
	    	  setCellNumberValue(wb, row, 4, cs_right, bean.getValue("field_920202_y"));//被掃模式-金額
	    	  
	    	  row = sheet_2.createRow(7);
	    	  row.setHeight((short)(sheet_1.getDefaultRowHeight()*2));
	    	  setCellValue(wb, row, 0, nb_left_wrap_text, "註"+itemNumber+"：新系統上線第一次申報時，由農漁會鍵入，第二個月起，由系統自動計算，如無交易應填「0」。");
	    	  
	    	  //-----項目7:行動支付-臺灣Pay交易情形-----
	    	  itemNumber++;
	    	  row = sheet_2.createRow(8);
	    	  setCellValue(wb, row, 0, nb_left, itemNumber+".行動支付--臺灣Pay交易情形");
	    	  //--消費扣款--
	    	  //筆數
	    	  row = sheet_2.createRow(10);
	    	  setCellNumberValue(wb, row, 2, cs_right, bean.getValue("field_920311_m"));//本月交易
	    	  setCellNumberValue(wb, row, 3, cs_right, bean.getValue("field_920311_y"));//本年累計交易
	    	  //金額
	    	  row = sheet_2.createRow(11);
	    	  setCellNumberValue(wb, row, 2, cs_right, bean.getValue("field_920312_m"));//本月交易
	    	  setCellNumberValue(wb, row, 3, cs_right, bean.getValue("field_920312_y"));//本年累計交易

	    	  //--非約定帳戶轉帳--
	    	  //筆數
	    	  row = sheet_2.createRow(12);
	    	  setCellNumberValue(wb, row, 2, cs_right, bean.getValue("field_920321_m"));//本月交易
	    	  setCellNumberValue(wb, row, 3, cs_right, bean.getValue("field_920321_y"));//本年累計交易
	    	  //金額
	    	  row = sheet_2.createRow(13);
	    	  setCellNumberValue(wb, row, 2, cs_right, bean.getValue("field_920322_m"));//本月交易
	    	  setCellNumberValue(wb, row, 3, cs_right, bean.getValue("field_920322_y"));//本年累計交易
	    	  
	    	  //--約定帳戶轉帳--
	    	  //筆數
	    	  row = sheet_2.createRow(14);
	    	  setCellNumberValue(wb, row, 2, cs_right, bean.getValue("field_920331_m"));//本月交易
	    	  setCellNumberValue(wb, row, 3, cs_right, bean.getValue("field_920331_y"));//本年累計交易
	    	  //金額
	    	  row = sheet_2.createRow(15);
	    	  setCellNumberValue(wb, row, 2, cs_right, bean.getValue("field_920332_m"));//本月交易
	    	  setCellNumberValue(wb, row, 3, cs_right, bean.getValue("field_920332_y"));//本年累計交易
	    	  
	    	  //--繳稅--
	    	  //筆數
	    	  row = sheet_2.createRow(16);
	    	  setCellNumberValue(wb, row, 2, cs_right, bean.getValue("field_920341_m"));//本月交易
	    	  setCellNumberValue(wb, row, 3, cs_right, bean.getValue("field_920341_y"));//本年累計交易
	    	  //金額
	    	  row = sheet_2.createRow(17);
	    	  setCellNumberValue(wb, row, 2, cs_right, bean.getValue("field_920342_m"));//本月交易
	    	  setCellNumberValue(wb, row, 3, cs_right, bean.getValue("field_920342_y"));//本年累計交易
	    	  
	    	  //--繳費--
	    	  //筆數
	    	  row = sheet_2.createRow(18);
	    	  setCellNumberValue(wb, row, 2, cs_right, bean.getValue("field_920351_m"));//本月交易
	    	  setCellNumberValue(wb, row, 3, cs_right, bean.getValue("field_920351_y"));//本年累計交易
	    	  //金額
	    	  row = sheet_2.createRow(19);
	    	  setCellNumberValue(wb, row, 2, cs_right, bean.getValue("field_920352_m"));//本月交易
	    	  setCellNumberValue(wb, row, 3, cs_right, bean.getValue("field_920352_y"));//本年累計交易
	    	  
	    	  row = sheet_2.createRow(20);
	    	  row.setHeight((short)(sheet_1.getDefaultRowHeight()*2));
	    	  setCellValue(wb, row, 0, nb_left_wrap_text, "註"+itemNumber+"：新系統上線第一次申報時，由農漁會鍵入，第二個月起，由系統自動計算，如無交易應填「0」。");
	    	  
	    	  //-----項目8:行動支付-LinePay交易情形-----
	    	  itemNumber++;
	    	  row = sheet_2.createRow(21);
	    	  setCellValue(wb, row, 0, nb_left, itemNumber+".行動支付--Line Pay交易情形");
	    	  //筆數
	    	  row = sheet_2.createRow(23);
	    	  setCellNumberValue(wb, row, 2, cs_right, bean.getValue("field_920471_m"));//本月交易
	    	  setCellNumberValue(wb, row, 3, cs_right, bean.getValue("field_920471_y"));//本年累計交易
	    	  //金額
	    	  row = sheet_2.createRow(24);
	    	  setCellNumberValue(wb, row, 2, cs_right, bean.getValue("field_920472_m"));//本月交易
	    	  setCellNumberValue(wb, row, 3, cs_right, bean.getValue("field_920472_y"));//本年累計交易
	    	  
	    	  row = sheet_2.createRow(25);
	    	  row.setHeight((short)(sheet_1.getDefaultRowHeight()*2));
	    	  setCellValue(wb, row, 0, nb_left_wrap_text, "註"+itemNumber+"：新系統上線第一次申報時，由農漁會鍵入，第二個月起，由系統自動計算，如無交易應填「0」。");
	    	  
	    	  //-----項目9:行動支付-動態新增項目-----
	    	  //取得報表動態新增項目9
		      HashMap<String, Object> item9CodeAndNameMap = getItem9CodeAndName(t1);
		      dbItem9DataList = getAddItemData(M_YEAR, M_MONTH, bank_code, item9CodeAndNameMap);
	    	  List item9CodeAndNameList = (List) item9CodeAndNameMap.get("codeAndNameList");
	    	  DataObject item9CodeAndName = null;
	    	  String[] item9FieldNames;
	    	  int item9Index = 26;
	    	  for(int i=0; i<item9CodeAndNameList.size(); i++){
	    		  itemNumber++;
	    		  item9CodeAndName = (DataObject)item9CodeAndNameList.get(i);
	    		  row = sheet_2.createRow( item9Index++ );
	    		  setCellValue(wb, row, 0, nb_left, itemNumber + ".行動支付--" + item9CodeAndName.getValue("limit_type_name"));//動態新增項目Title

	    		  row = sheet_2.createRow( item9Index );
	    		  region = new Region(item9Index, (short)0, item9Index, (short)1);
	    		  sheet_2.addMergedRegion(region);
	    		  item9Index++;
	    		  setCellValue(wb, row, 0, cs_center, "");
	    		  setCellValue(wb, row, 1, cs_center, "");
	    		  setCellValue(wb, row, 2, cs_center, "本月交易");
	    		  setCellValue(wb, row, 3, cs_center, "本年累計交易");
	    		  
	    		  item9FieldNames = ((String) item9CodeAndNameMap.get((String)item9CodeAndName.getValue("limit_type"))).split(",");
	    		  row = sheet_2.createRow( item9Index );
	    		  region = new Region(item9Index, (short)0, item9Index+1, (short)0);
	    		  sheet_2.addMergedRegion(region);
	    		  item9Index++;
	    		  setCellValue(wb, row, 0, cs_center, "消費扣款");
	    		  setCellValue(wb, row, 1, cs_center, "筆數");
	    		  setCellNumberValue(wb, row, 2, cs_center, dbItem9DataList==null ? "0" : ((DataObject)dbItem9DataList.get(0)).getValue(item9FieldNames[0]));
	    		  setCellNumberValue(wb, row, 3, cs_center, dbItem9DataList==null ? "0" : ((DataObject)dbItem9DataList.get(0)).getValue(item9FieldNames[1]));
	    		  row = sheet_2.createRow( item9Index++ );
	    		  setCellValue(wb, row, 1, cs_center, "金額");
	    		  setCellNumberValue(wb, row, 2, cs_center, dbItem9DataList==null ? "0" : ((DataObject)dbItem9DataList.get(0)).getValue(item9FieldNames[2]));
	    		  setCellNumberValue(wb, row, 3, cs_center, dbItem9DataList==null ? "0" : ((DataObject)dbItem9DataList.get(0)).getValue(item9FieldNames[3]));
	    		  
	    		  row = sheet_2.createRow( item9Index );
	    		  region = new Region(item9Index, (short)0, item9Index+1, (short)0);
	    		  sheet_2.addMergedRegion(region);
	    		  item9Index++;
	    		  setCellValue(wb, row, 0, cs_center, "非約定帳戶轉帳");
	    		  setCellValue(wb, row, 1, cs_center, "筆數");
	    		  setCellNumberValue(wb, row, 2, cs_center, dbItem9DataList==null ? "0" : ((DataObject)dbItem9DataList.get(0)).getValue(item9FieldNames[4]));
	    		  setCellNumberValue(wb, row, 3, cs_center, dbItem9DataList==null ? "0" : ((DataObject)dbItem9DataList.get(0)).getValue(item9FieldNames[5]));
	    		  row = sheet_2.createRow( item9Index++ );
	    		  setCellValue(wb, row, 1, cs_center, "金額");
	    		  setCellNumberValue(wb, row, 2, cs_center, dbItem9DataList==null ? "0" : ((DataObject)dbItem9DataList.get(0)).getValue(item9FieldNames[6]));
	    		  setCellNumberValue(wb, row, 3, cs_center, dbItem9DataList==null ? "0" : ((DataObject)dbItem9DataList.get(0)).getValue(item9FieldNames[7]));
	    		  
	    		  row = sheet_2.createRow( item9Index );
	    		  region = new Region(item9Index, (short)0, item9Index+1, (short)0);
	    		  sheet_2.addMergedRegion(region);
	    		  item9Index++;
	    		  setCellValue(wb, row, 0, cs_center, "約定帳戶轉帳");
	    		  setCellValue(wb, row, 1, cs_center, "筆數");
	    		  setCellNumberValue(wb, row, 2, cs_center, dbItem9DataList==null ? "0" : ((DataObject)dbItem9DataList.get(0)).getValue(item9FieldNames[8]));
	    		  setCellNumberValue(wb, row, 3, cs_center, dbItem9DataList==null ? "0" : ((DataObject)dbItem9DataList.get(0)).getValue(item9FieldNames[9]));
	    		  row = sheet_2.createRow( item9Index++ );
	    		  setCellValue(wb, row, 1, cs_center, "金額");
	    		  setCellNumberValue(wb, row, 2, cs_center, dbItem9DataList==null ? "0" : ((DataObject)dbItem9DataList.get(0)).getValue(item9FieldNames[10]));
	    		  setCellNumberValue(wb, row, 3, cs_center, dbItem9DataList==null ? "0" : ((DataObject)dbItem9DataList.get(0)).getValue(item9FieldNames[11]));
	    		  
	    		  row = sheet_2.createRow( item9Index );
	    		  region = new Region(item9Index, (short)0, item9Index+1, (short)0);
	    		  sheet_2.addMergedRegion(region);
	    		  item9Index++;
	    		  setCellValue(wb, row, 0, cs_center, "繳稅");
	    		  setCellValue(wb, row, 1, cs_center, "筆數");
	    		  setCellNumberValue(wb, row, 2, cs_center, dbItem9DataList==null ? "0" : ((DataObject)dbItem9DataList.get(0)).getValue(item9FieldNames[12]));
	    		  setCellNumberValue(wb, row, 3, cs_center, dbItem9DataList==null ? "0" : ((DataObject)dbItem9DataList.get(0)).getValue(item9FieldNames[13]));
	    		  row = sheet_2.createRow( item9Index++ );
	    		  setCellValue(wb, row, 1, cs_center, "金額");
	    		  setCellNumberValue(wb, row, 2, cs_center, dbItem9DataList==null ? "0" : ((DataObject)dbItem9DataList.get(0)).getValue(item9FieldNames[14]));
	    		  setCellNumberValue(wb, row, 3, cs_center, dbItem9DataList==null ? "0" : ((DataObject)dbItem9DataList.get(0)).getValue(item9FieldNames[15]));
	    		  
	    		  row = sheet_2.createRow( item9Index );
	    		  region = new Region(item9Index, (short)0, item9Index+1, (short)0);
	    		  sheet_2.addMergedRegion(region);
	    		  item9Index++;
	    		  setCellValue(wb, row, 0, cs_center, "繳費");
	    		  setCellValue(wb, row, 1, cs_center, "筆數");
	    		  setCellNumberValue(wb, row, 2, cs_center, dbItem9DataList==null ? "0" : ((DataObject)dbItem9DataList.get(0)).getValue(item9FieldNames[16]));
	    		  setCellNumberValue(wb, row, 3, cs_center, dbItem9DataList==null ? "0" : ((DataObject)dbItem9DataList.get(0)).getValue(item9FieldNames[17]));
	    		  row = sheet_2.createRow( item9Index++ );
	    		  setCellValue(wb, row, 0, cs_center, "");
	    		  setCellValue(wb, row, 1, cs_center, "金額");
	    		  setCellNumberValue(wb, row, 2, cs_center, dbItem9DataList==null ? "0" : ((DataObject)dbItem9DataList.get(0)).getValue(item9FieldNames[18]));
	    		  setCellNumberValue(wb, row, 3, cs_center, dbItem9DataList==null ? "0" : ((DataObject)dbItem9DataList.get(0)).getValue(item9FieldNames[19]));
	    		  
	    		  region = new Region(item9Index, (short)0, item9Index, (short)4);
	    		  sheet_2.addMergedRegion(region);
	    		  row = sheet_2.createRow( item9Index++ );
	    		  row.setHeight((short)(sheet_1.getDefaultRowHeight()*2));
	    		  setCellValue(wb, row, 0, nb_left_wrap_text, "註"+itemNumber+"：新系統上線第一次申報時，由農漁會鍵入，第二個月起，由系統自動計算，如無交易應填「0」。");
	    	  }
	      } //end of if
	      
	      
	      FileOutputStream fout = null;     
	      fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + filename);
	     
	      HSSFFooter footer = sheet_1.getFooter();
	      footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
	      footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
	      footer = sheet_2.getFooter();
	      footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
	      footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
	      wb.write(fout);
	      //儲存
	      fout.close();
	    }
	    catch (Exception e) {
	    	System.out.println("RptFR083W.createRpt Error:" + e + e.getMessage());
	    }
	    
	    return errMsg;
	  }
	 
	 @SuppressWarnings("rawtypes")
	 private static List getBaseData(String M_YEAR, String M_MONTH, String bank_code){
		 List dataList = null;
		 StringBuffer sql = new StringBuffer();
		 ArrayList<String> paramList = new ArrayList<String>();
		 sql.append(" select "
	    		   //項目1:客戶數
	      		   + "		  sum(field_900100_m) as field_900100_m, "	//--電子銀行-本月開戶數
	    		   + " 		  sum(field_900100_y) as field_900100_y, "	//--電子銀行-累計開戶數
	    		   + " 		  sum(field_900200_m) as field_900200_m, "	//--行動銀行-本月開戶數
	    		   + " 		  sum(field_900200_y) as field_900200_y, "	//--行動銀行-累計開戶數 
	    		   + " 		  sum(field_900300_m) as field_900300_m, "	//--收單特店-主掃模式-本月開戶數
	    		   + " 		  sum(field_900300_y) as field_900300_y, "	//--收單特店-主掃模式-累計開戶數
	    		   + " 		  sum(field_900400_m) as field_900400_m, "	//--收單特店-被掃模式-本月開戶數
	    		   + " 		  sum(field_900400_y) as field_900400_y, "	//--收單特店-被掃模式-累計開戶數
	    		   //項目2:約定帳戶轉帳交易情形
	    		   + " 		  sum(field_910101_m) as field_910101_m, "	//--約定帳戶轉帳交易-本月交易筆數
	    		   + " 		  sum(field_910101_y) as field_910101_y, "	//--約定帳戶轉帳交易-本年累計交易筆數
	    		   + " 		  sum(field_910102_m) as field_910102_m, "	//--約定帳戶轉帳交易-本月交易金額
	    		   + " 		  sum(field_910102_y) as field_910102_y, "	//--約定帳戶轉帳交易-本年累計交易金額
	    		   //項目3:非約定帳戶轉帳交易情形
	    		   + " 		  sum(field_910201_m) as field_910201_m, "	//--非約定帳戶轉帳交易-本月交易筆數
	    		   + " 		  sum(field_910201_y) as field_910201_y, "	//--非約定帳戶轉帳交易-本年累計交易筆數
	    		   + " 		  sum(field_910202_m) as field_910202_m, "	//--非約定帳戶轉帳交易-本月交易金額
	    		   + " 		  sum(field_910202_y) as field_910202_y, "	//--非約定帳戶轉帳交易-本年累計交易金額
	    		   //項目4:臺灣Pay購物交易情形
	    		   + " 		  sum(field_910311_m) as field_910311_m, "	//--臺灣Pay購物交易-轉帳購物-本月交易筆數
	    		   + " 		  sum(field_910311_y) as field_910311_y, "	//--臺灣Pay購物交易-轉帳購物-本年累計交易筆數
	    		   + " 		  sum(field_910312_m) as field_910312_m, "	//--臺灣Pay購物交易-轉帳購物-本月交易金額
	    		   + " 		  sum(field_910312_y) as field_910312_y, "	//--臺灣Pay購物交易-轉帳購物-本年累計交易金額
	    		   + " 		  sum(field_910321_m) as field_910321_m, "	//--臺灣Pay購物交易-消費扣款-本月交易筆數
	    		   + " 		  sum(field_910321_y) as field_910321_y, "	//--臺灣Pay購物交易-消費扣款-本年累計交易筆數
	    		   + " 		  sum(field_910322_m) as field_910322_m, "	//--臺灣Pay購物交易-消費扣款-本月交易金額
	    		   + " 		  sum(field_910322_y) as field_910322_y, "	//--臺灣Pay購物交易-消費扣款-本年累計交易金額
	    		   + " 		  sum(field_910331_m) as field_910331_m, "	//--臺灣Pay購物交易-繳費-本月交易筆數
	    		   + " 		  sum(field_910331_y) as field_910331_y, "	//--臺灣Pay購物交易-繳費-本年累計交易筆數
	    		   + " 		  sum(field_910332_m) as field_910332_m, "	//--臺灣Pay購物交易-繳費-本月交易金額
	    		   + " 		  sum(field_910332_y) as field_910332_y, "	//--臺灣Pay購物交易-繳費-本年累計交易金額
	    		   + " 		  sum(field_910341_m) as field_910341_m, "	//--臺灣Pay購物交易-P2P轉帳-本月交易筆數
	    		   + " 		  sum(field_910341_y) as field_910341_y, "	//--臺灣Pay購物交易-P2P轉帳-本年累計交易筆數
	    		   + " 		  sum(field_910342_m) as field_910342_m, "	//--臺灣Pay購物交易-P2P轉帳-本月交易金額
	    		   + " 		  sum(field_910342_y) as field_910342_y, "	//--臺灣Pay購物交易-P2P轉帳-本年累計交易金額
	    		   //項目6:特店收單交易情形
	    		   + " 		  sum(field_920101_m) as field_920101_m, "	//--特店收單-主掃模式-本月交易筆數
	    		   + " 		  sum(field_920101_y) as field_920101_y, "	//--特店收單-主掃模式-本年累計交易筆數
	    		   + " 		  sum(field_920102_m) as field_920102_m, "	//--特店收單-主掃模式-本月交易金額
	    		   + " 		  sum(field_920102_y) as field_920102_y, "	//--特店收單-主掃模式本年累計交易金額
	    		   + " 		  sum(field_920201_m) as field_920201_m, "	//--特店收單-被掃模式-本月交易筆數
	    		   + " 		  sum(field_920201_y) as field_920201_y, "	//--特店收單-被掃模式-本年累計交易筆數
	    		   + " 		  sum(field_920202_m) as field_920202_m, "	//--特店收單-被掃模式-本月交易金額
	    		   + " 		  sum(field_920202_y) as field_920202_y, "	//--特店收單-被掃模式-本年累計交易金額
	    		   //項目7:行動支付-臺灣Pay交易情形
	    		   + " 		  sum(field_920311_m) as field_920311_m, "	//--行動支付-臺灣Pay-消費扣款-本月交易筆數
	    		   + " 		  sum(field_920311_y) as field_920311_y, "	//--行動支付-臺灣Pay-消費扣款-本年累計交易筆數
	    		   + " 		  sum(field_920312_m) as field_920312_m, "	//--行動支付-臺灣Pay-消費扣款-本月交易金額
	    		   + " 		  sum(field_920312_y) as field_920312_y, "	//--行動支付-臺灣Pay-消費扣款-本年累計交易金額
	    		   + " 		  sum(field_920321_m) as field_920321_m, "	//--行動支付-臺灣Pay-非約定帳戶轉帳-本月交易筆數
	    		   + " 		  sum(field_920321_y) as field_920321_y, "	//--行動支付-臺灣Pay-非約定帳戶轉帳-本年累計交易筆數
	    		   + " 		  sum(field_920322_m) as field_920322_m, "	//--行動支付-臺灣Pay-非約定帳戶轉帳-本月交易金額
	    		   + " 		  sum(field_920322_y) as field_920322_y, "	//--行動支付-臺灣Pay-非約定帳戶轉帳-本年累計交易金額
	    		   + " 		  sum(field_920331_m) as field_920331_m, "	//--行動支付-臺灣Pay-約定帳戶轉帳-本月交易筆數
	    		   + " 		  sum(field_920331_y) as field_920331_y, "	//--行動支付-臺灣Pay-約定帳戶轉帳-本年累計交易筆數
	    		   + " 		  sum(field_920332_m) as field_920332_m, "	//--行動支付-臺灣Pay-約定帳戶轉帳-本月交易金額
	    		   + " 		  sum(field_920332_y) as field_920332_y, "	//--行動支付-臺灣Pay-約定帳戶轉帳-本年累計交易金額
	    		   + " 		  sum(field_920341_m) as field_920341_m, "	//--行動支付-臺灣Pay-繳稅-本月交易筆數
	    		   + " 		  sum(field_920341_y) as field_920341_y, "	//--行動支付-臺灣Pay-繳稅-本年累計交易筆數
	    		   + " 		  sum(field_920342_m) as field_920342_m, "	//--行動支付-臺灣Pay-繳稅-本月交易金額
	    		   + " 		  sum(field_920342_y) as field_920342_y, "	//--行動支付-臺灣Pay-繳稅-本年累計交易金額
	    		   + " 		  sum(field_920351_m) as field_920351_m, "	//--行動支付-臺灣Pay-繳費-本月交易筆數
	    		   + " 		  sum(field_920351_y) as field_920351_y, "	//--行動支付-臺灣Pay-繳費-本年累計交易筆數
	    		   + " 		  sum(field_920352_m) as field_920352_m, "	//--行動支付-臺灣Pay-繳費-本月交易金額
	    		   + " 		  sum(field_920352_y) as field_920352_y, "	//--行動支付-臺灣Pay-繳費-本年累計交易金額
	    		   //項目8:行動支付-LinePay交易情形
	    		   + " 		  sum(field_920471_m) as field_920471_m, "	//--行動支付-LinePay-LinePay儲值-本月交易筆數
	    		   + " 		  sum(field_920471_y) as field_920471_y, "	//--行動支付-LinePay-LinePay儲值-本年累計交易筆數
	    		   + " 		  sum(field_920472_m) as field_920472_m, "	//--行動支付-LinePay-LinePay儲值-本月交易金額
	    		   + " 		  sum(field_920472_y) as field_920472_y  "	//--行動支付-LinePay-LinePay儲值-本年累計交易金額
	    		   + " from  "
	    		   + " (    select decode(acc_code,'900100',month_amt,0) as field_900100_m,"	//--電子銀行-本月開戶數
	    		   + " 			   decode(acc_code,'900100',year_amt,0) as field_900100_y, "	//--電子銀行-累計開戶數
	    		   + " 			   decode(acc_code,'900200',month_amt,0) as field_900200_m,"	//--行動銀行-本月開戶數
	    		   + " 			   decode(acc_code,'900200',year_amt,0) as field_900200_y, "	//--行動銀行-累計開戶數
	    		   + " 			   decode(acc_code,'900300',month_amt,0) as field_900300_m,"	//--收單特店-主掃模式-本月開戶數
	    		   + " 			   decode(acc_code,'900300',year_amt,0) as field_900300_y, "	//--收單特店-主掃模式-累計開戶數
	    		   + " 			   decode(acc_code,'900400',month_amt,0) as field_900400_m,"	//--收單特店-被掃模式-本月開戶數
	    		   + " 			   decode(acc_code,'900400',year_amt,0) as field_900400_y, "	//--收單特店-被掃模式-累計開戶數
	    		   + " 			   decode(acc_code,'910101',month_amt,0) as field_910101_m,"	//--約定帳戶轉帳交易-本月交易筆數
	    		   + " 			   decode(acc_code,'910101',year_amt,0) as field_910101_y, "	//--約定帳戶轉帳交易-本年累計交易筆數
	    		   + " 			   decode(acc_code,'910102',month_amt,0) as field_910102_m,"	//--約定帳戶轉帳交易-本月交易金額
	    		   + " 			   decode(acc_code,'910102',year_amt,0) as field_910102_y, "	//--約定帳戶轉帳交易-本年累計交易金額
	    		   + " 			   decode(acc_code,'910201',month_amt,0) as field_910201_m,"	//--非約定帳戶轉帳交易-本月交易筆數
	    		   + " 			   decode(acc_code,'910201',year_amt,0) as field_910201_y, "	//--非約定帳戶轉帳交易-本年累計交易筆數
	    		   + " 			   decode(acc_code,'910202',month_amt,0) as field_910202_m,"	//--非約定帳戶轉帳交易-本月交易金額
	    		   + " 			   decode(acc_code,'910202',year_amt,0) as field_910202_y, "	//--非約定帳戶轉帳交易-本年累計交易金額
	    		   + " 			   decode(acc_code,'910311',month_amt,0) as field_910311_m,"	//--臺灣Pay購物交易-轉帳購物-本月交易筆數
	    		   + " 			   decode(acc_code,'910311',year_amt,0) as field_910311_y, "	//--臺灣Pay購物交易-轉帳購物-本年累計交易筆數
	    		   + " 			   decode(acc_code,'910312',month_amt,0) as field_910312_m,"	//--臺灣Pay購物交易-轉帳購物-本月交易金額
	    		   + " 			   decode(acc_code,'910312',year_amt,0) as field_910312_y, "	//--臺灣Pay購物交易-轉帳購物-本年累計交易金額
	    		   + " 			   decode(acc_code,'910321',month_amt,0) as field_910321_m,"	//--臺灣Pay購物交易-消費扣款-本月交易筆數
	    		   + " 			   decode(acc_code,'910321',year_amt,0) as field_910321_y, "	//--臺灣Pay購物交易-消費扣款-本年累計交易筆數
	    		   + " 			   decode(acc_code,'910322',month_amt,0) as field_910322_m,"	//--臺灣Pay購物交易-消費扣款-本月交易金額
	    		   + " 			   decode(acc_code,'910322',year_amt,0) as field_910322_y, "	//--臺灣Pay購物交易-消費扣款-本年累計交易金額
	    		   + " 			   decode(acc_code,'910331',month_amt,0) as field_910331_m,"	//--臺灣Pay購物交易-繳費-本月交易筆數
	    		   + " 			   decode(acc_code,'910331',year_amt,0) as field_910331_y, "	//--臺灣Pay購物交易-繳費-本年累計交易筆數
	    		   + " 			   decode(acc_code,'910332',month_amt,0) as field_910332_m,"	//--臺灣Pay購物交易-繳費-本月交易金額
	    		   + " 			   decode(acc_code,'910332',year_amt,0) as field_910332_y, "	//--臺灣Pay購物交易-繳費-本年累計交易金額
	    		   + " 			   decode(acc_code,'910341',month_amt,0) as field_910341_m,"	//--臺灣Pay購物交易-P2P轉帳-本月交易筆數
	    		   + " 			   decode(acc_code,'910341',year_amt,0) as field_910341_y, "	//--臺灣Pay購物交易-P2P轉帳-本年累計交易筆數
	    		   + " 			   decode(acc_code,'910342',month_amt,0) as field_910342_m,"	//--臺灣Pay購物交易-P2P轉帳-本月交易金額
	    		   + " 			   decode(acc_code,'910342',year_amt,0) as field_910342_y, "	//--臺灣Pay購物交易-P2P轉帳-本年累計交易金額
	    		   + " 			   decode(acc_code,'920101',month_amt,0) as field_920101_m,"	//--主掃模式-本月交易筆數
	    		   + " 			   decode(acc_code,'920101',year_amt,0) as field_920101_y, "	//--主掃模式-本年累計交易筆數
	    		   + " 			   decode(acc_code,'920102',month_amt,0) as field_920102_m,"	//--主掃模式-本月交易金額
	    		   + " 			   decode(acc_code,'920102',year_amt,0) as field_920102_y, "	//--主掃模式-本年累計交易金額
	    		   + " 			   decode(acc_code,'920201',month_amt,0) as field_920201_m,"	//--被掃模式-本月交易筆數
	    		   + " 			   decode(acc_code,'920201',year_amt,0) as field_920201_y, "	//--被掃模式-本年累計交易筆數
	    		   + " 			   decode(acc_code,'920202',month_amt,0) as field_920202_m,"	//--被掃模式-本月交易金額
	    		   + " 			   decode(acc_code,'920202',year_amt,0) as field_920202_y, "	//--被掃模式-本年累計交易金額
	    		   + " 			   decode(acc_code,'920311',month_amt,0) as field_920311_m,"	//--行動支付-臺灣Pay-消費扣款-本月交易筆數
	    		   + " 			   decode(acc_code,'920311',year_amt,0) as field_920311_y, "	//--行動支付-臺灣Pay-消費扣款-本年累計交易筆數
	    		   + " 			   decode(acc_code,'920312',month_amt,0) as field_920312_m,"	//--行動支付-臺灣Pay-消費扣款-本月交易金額
	    		   + " 			   decode(acc_code,'920312',year_amt,0) as field_920312_y, "	//--行動支付-臺灣Pay-消費扣款-本年累計交易金額
	    		   + " 			   decode(acc_code,'920321',month_amt,0) as field_920321_m,"	//--行動支付-臺灣Pay-非約定帳戶轉帳-本月交易筆數
	    		   + " 			   decode(acc_code,'920321',year_amt,0) as field_920321_y, "	//--行動支付-臺灣Pay-非約定帳戶轉帳-本年累計交易筆數
	    		   + " 			   decode(acc_code,'920322',month_amt,0) as field_920322_m,"	//--行動支付-臺灣Pay-非約定帳戶轉帳-本月交易金額
	    		   + " 			   decode(acc_code,'920322',year_amt,0) as field_920322_y, "	//--行動支付-臺灣Pay-非約定帳戶轉帳-本年累計交易金額
	    		   + " 			   decode(acc_code,'920331',month_amt,0) as field_920331_m,"	//--行動支付-臺灣Pay-約定帳戶轉帳-本月交易筆數
	    		   + " 			   decode(acc_code,'920331',year_amt,0) as field_920331_y, "	//--行動支付-臺灣Pay-約定帳戶轉帳-本年累計交易筆數
	    		   + " 			   decode(acc_code,'920332',month_amt,0) as field_920332_m,"	//--行動支付-臺灣Pay-約定帳戶轉帳-本月交易金額
	    		   + " 			   decode(acc_code,'920332',year_amt,0) as field_920332_y, "	//--行動支付-臺灣Pay-約定帳戶轉帳-本年累計交易金額
	    		   + " 			   decode(acc_code,'920341',month_amt,0) as field_920341_m,"	//--行動支付-臺灣Pay-繳稅-本月交易筆數
	    		   + " 			   decode(acc_code,'920341',year_amt,0) as field_920341_y, "	//--行動支付-臺灣Pay-繳稅-本年累計交易筆數
	    		   + " 			   decode(acc_code,'920342',month_amt,0) as field_920342_m,"	//--行動支付-臺灣Pay-繳稅-本月交易金額
	    		   + " 			   decode(acc_code,'920342',year_amt,0) as field_920342_y, "	//--行動支付-臺灣Pay-繳稅-本年累計交易金額
	    		   + " 			   decode(acc_code,'920351',month_amt,0) as field_920351_m,"	//--行動支付-臺灣Pay-繳費-本月交易筆數
	    		   + " 			   decode(acc_code,'920351',year_amt,0) as field_920351_y, "	//--行動支付-臺灣Pay-繳費-本年累計交易筆數
	    		   + " 			   decode(acc_code,'920352',month_amt,0) as field_920352_m,"	//--行動支付-臺灣Pay-繳費-本月交易金額
	    		   + " 			   decode(acc_code,'920352',year_amt,0) as field_920352_y, "	//--行動支付-臺灣Pay-繳費-本年累計交易金額
	    		   + " 			   decode(acc_code,'920471',month_amt,0) as field_920471_m,"	//--行動支付-LinePay-LinePay儲值-本月交易筆數
	    		   + " 			   decode(acc_code,'920471',year_amt,0) as field_920471_y, "	//--行動支付-LinePay-LinePay儲值-本年累計交易筆數
	    		   + " 			   decode(acc_code,'920472',month_amt,0) as field_920472_m,"	//--行動支付-LinePay-LinePay儲值-本月交易金額
	    		   + " 			   decode(acc_code,'920472',year_amt,0) as field_920472_y  "	//--行動支付-LinePay-LinePay儲值-本年累計交易金額
	    		   + " 		from a15 "
	    		   + " 		where 	  m_year=? "
	    		   + "			  and m_month=? "
	    		   + " 			  and bank_code=? "
	    		   + " )a " );
	      paramList.add(M_YEAR);
	      paramList.add(M_MONTH);
	      paramList.add(bank_code);
		 
		 String orgTypeFields = //項目1:客戶數
				 				"field_900100_m,"	//--電子銀行-本月開戶數
				 			  + "field_900100_y,"	//--電子銀行-累計開戶數
				 			  + "field_900200_m,"	//--行動銀行-本月開戶數
				 			  + "field_900200_y,"	//--行動銀行-累計開戶數 
				 			  + "field_900300_m,"	//--收單特店-主掃模式-本月開戶數
				 			  + "field_900300_y,"	//--收單特店-主掃模式-累計開戶數
				 			  + "field_900400_m,"	//--收單特店-被掃模式-本月開戶數
				 			  + "field_900400_y,"	//--收單特店-被掃模式-累計開戶數
				 			  //項目2:約定帳戶轉帳交易情形
				 			  + "field_910101_m,"	//--約定帳戶轉帳交易-本月交易筆數
				 			  + "field_910101_y,"	//--約定帳戶轉帳交易-本年累計交易筆數
				 			  + "field_910102_m,"	//--約定帳戶轉帳交易-本月交易金額
				 			  + "field_910102_y,"	//--約定帳戶轉帳交易-本年累計交易金額
				 			  //項目3:非約定帳戶轉帳交易情形
				 			  + "field_910201_m,"	//--非約定帳戶轉帳交易-本月交易筆數
				 			  + "field_910201_y,"	//--非約定帳戶轉帳交易-本年累計交易筆數
				 			  + "field_910202_m,"	//--非約定帳戶轉帳交易-本月交易金額
				 			  + "field_910202_y,"	//--非約定帳戶轉帳交易-本年累計交易金額
				 			  //項目4:臺灣Pay購物交易情形
				 			  + "field_910311_m,"	//--臺灣Pay購物交易-轉帳購物-本月交易筆數
				 			  + "field_910311_y,"	//--臺灣Pay購物交易-轉帳購物-本年累計交易筆數
				 			  + "field_910312_m,"	//--臺灣Pay購物交易-轉帳購物-本月交易金額
				 			  + "field_910312_y,"	//--臺灣Pay購物交易-轉帳購物-本年累計交易金額
				 			  + "field_910321_m,"	//--臺灣Pay購物交易-消費扣款-本月交易筆數
				 			  + "field_910321_y,"	//--臺灣Pay購物交易-消費扣款-本年累計交易筆數
				 			  + "field_910322_m,"	//--臺灣Pay購物交易-消費扣款-本月交易金額
				 			  + "field_910322_y,"	//--臺灣Pay購物交易-消費扣款-本年累計交易金額
				 			  + "field_910331_m,"	//--臺灣Pay購物交易-繳費-本月交易筆數
				 			  + "field_910331_y,"	//--臺灣Pay購物交易-繳費-本年累計交易筆數
				 			  + "field_910332_m,"	//--臺灣Pay購物交易-繳費-本月交易金額
				 			  + "field_910332_y,"	//--臺灣Pay購物交易-繳費-本年累計交易金額
				 			  + "field_910341_m,"	//--臺灣Pay購物交易-P2P轉帳-本月交易筆數
				 			  + "field_910341_y,"	//--臺灣Pay購物交易-P2P轉帳-本年累計交易筆數
				 			  + "field_910342_m,"	//--臺灣Pay購物交易-P2P轉帳-本月交易金額
				 			  + "field_910342_y,"	//--臺灣Pay購物交易-P2P轉帳-本年累計交易金額
				 			  //項目6:特店收單交易情形
				 			  + "field_920101_m,"	//--特店收單-主掃模式-本月交易筆數
				 			  + "field_920101_y,"	//--特店收單-主掃模式-本年累計交易筆數
				 			  + "field_920102_m,"	//--特店收單-主掃模式-本月交易金額
				 			  + "field_920102_y,"	//--特店收單-主掃模式本年累計交易金額
				 			  + "field_920201_m,"	//--特店收單-被掃模式-本月交易筆數
				 			  + "field_920201_y,"	//--特店收單-被掃模式-本年累計交易筆數
				 			  + "field_920202_m,"	//--特店收單-被掃模式-本月交易金額
				 			  + "field_920202_y,"	//--特店收單-被掃模式-本年累計交易金額
				 			  //項目7:行動支付-臺灣Pay交易情形
				 			  + "field_920311_m,"	//--行動支付-臺灣Pay-消費扣款-本月交易筆數
				 			  + "field_920311_y,"	//--行動支付-臺灣Pay-消費扣款-本年累計交易筆數
				 			  + "field_920312_m,"	//--行動支付-臺灣Pay-消費扣款-本月交易金額
				 			  + "field_920312_y,"	//--行動支付-臺灣Pay-消費扣款-本年累計交易金額
				 			  + "field_920321_m,"	//--行動支付-臺灣Pay-非約定帳戶轉帳-本月交易筆數
				 			  + "field_920321_y,"	//--行動支付-臺灣Pay-非約定帳戶轉帳-本年累計交易筆數
				 			  + "field_920322_m,"	//--行動支付-臺灣Pay-非約定帳戶轉帳-本月交易金額
				 			  + "field_920322_y,"	//--行動支付-臺灣Pay-非約定帳戶轉帳-本年累計交易金額
				 			  + "field_920331_m,"	//--行動支付-臺灣Pay-約定帳戶轉帳-本月交易筆數
				 			  + "field_920331_y,"	//--行動支付-臺灣Pay-約定帳戶轉帳-本年累計交易筆數
				 			  + "field_920332_m,"	//--行動支付-臺灣Pay-約定帳戶轉帳-本月交易金額
				 			  + "field_920332_y,"	//--行動支付-臺灣Pay-約定帳戶轉帳-本年累計交易金額
				 			  + "field_920341_m,"	//--行動支付-臺灣Pay-繳稅-本月交易筆數
				 			  + "field_920341_y,"	//--行動支付-臺灣Pay-繳稅-本年累計交易筆數
				 			  + "field_920342_m,"	//--行動支付-臺灣Pay-繳稅-本月交易金額
				 			  + "field_920342_y,"	//--行動支付-臺灣Pay-繳稅-本年累計交易金額
				 			  + "field_920351_m,"	//--行動支付-臺灣Pay-繳費-本月交易筆數
				 			  + "field_920351_y,"	//--行動支付-臺灣Pay-繳費-本年累計交易筆數
				 			  + "field_920352_m,"	//--行動支付-臺灣Pay-繳費-本月交易金額
				 			  + "field_920352_y,"	//--行動支付-臺灣Pay-繳費-本年累計交易金額
				 			  //項目8:行動支付-LinePay交易情形
				 			  + "field_920471_m,"	//--行動支付-LinePay-LinePay儲值-本月交易筆數
				 			  + "field_920471_y,"	//--行動支付-LinePay-LinePay儲值-本年累計交易筆數
				 			  + "field_920472_m,"	//--行動支付-LinePay-LinePay儲值-本月交易金額
				 			  + "field_920472_y ";	//--行動支付-LinePay-LinePay儲值-本年累計交易金額
		 
		 dataList =  DBManager.QueryDB_SQLParam(sql.toString(), paramList, orgTypeFields);
	     return dataList;
	 }
	 
	 @SuppressWarnings("rawtypes")
	 private static HashMap<String, Object> getItem5CodeAndName(String ncacno){
		 HashMap<String, Object> dataMap = new HashMap<String, Object>();
		 StringBuffer sql = new StringBuffer();
		 ArrayList paramList = new ArrayList();
		 
		 sql.append(" select limit_type, "		//--項目代碼
	    		   +" 		 limit_type_name " 	//--項目名稱
	    		   +" from wlx01_limit_item " 
	    		   +" left join "+ncacno+" on substr( "+ncacno+".acc_code,2,3) = limit_type " 
	    		   +" where limit_kind = '1' and limit_type not in ('101','102','103') " 
	    		   +" and acc_tr_type = 'A15' and acc_div='15' " 
	    		   +" group by limit_type, limit_type_name, output_order "
	    		   +" order by output_order ");
		 String orgTypeFields = "limit_type,limit_type_name";
		 List dataList = DBManager.QueryDB_SQLParam(sql.toString(), paramList, orgTypeFields);
		 
		 DataObject data = null;
		 String fieldCode = "";
		 String fieldNames = "";
		 StringBuffer fieldNamesBuffer = new StringBuffer();
		 for(int i=0; i<dataList.size();i++){
			 data = (DataObject)dataList.get(i);
			 fieldCode = data.getValue("limit_type") == null ? "" : data.getValue("limit_type").toString();
			 fieldNames = "field_9" + fieldCode + "01_m,"		//--電子銀行動態新增項目-本月交易筆數
				 		+ "field_9" + fieldCode + "01_y,"		//--電子銀行動態新增項目-本年累計交易筆數
				 		+ "field_9" + fieldCode + "02_m,";		//--電子銀行動態新增項目-本月交易金額
		 
			 if(i != dataList.size()-1){
				 fieldNames += "field_9" + fieldCode + "02_y,";//--電子銀行動態新增項目-本年累計交易金額
			 } else {
				 fieldNames += "field_9" + fieldCode + "02_y";//--電子銀行動態新增項目-本年累計交易金額
			 }
			 fieldNamesBuffer.append(fieldNames.trim());
			 dataMap.put(fieldCode, fieldNames.trim());
		 }
		 
		 dataMap.put("codeAndNameList", dataList);
		 dataMap.put("fieldNames", fieldNamesBuffer.toString());
		 return dataMap;
	 }
	 
	 @SuppressWarnings("rawtypes")
	 private static HashMap<String, Object> getItem9CodeAndName(String ncacno){
		 HashMap<String, Object> dataMap = new HashMap<String, Object>();
		 StringBuffer sql = new StringBuffer();
		 ArrayList paramList = new ArrayList();
		 sql.append(" select limit_type, "		//--項目代碼
				   +" 		 limit_type_name " 	//--項目名稱
				   +" from wlx01_limit_item " 
				   +" left join "+ncacno+" on substr( "+ncacno+".acc_code,2,3) = limit_type " 
				   +" where limit_kind = '2' and limit_type not in ('203','204') "
				   +" and acc_tr_type = 'A15' and acc_div='15' "
				   +" group by limit_type, limit_type_name, output_order "
				   +" order by output_order ");
		 String orgTypeFields = "limit_type,limit_type_name";
		 List dataList =  DBManager.QueryDB_SQLParam(sql.toString(), paramList, orgTypeFields);
		 
		 DataObject data = null;
		 String fieldCode = "";
		 String fieldNames = "";
		 StringBuffer fieldNamesBuffer = new StringBuffer();
		 for(int i=0; i<dataList.size();i++){
			 data = (DataObject)dataList.get(i);
			 fieldCode = data.getValue("limit_type") == null ? "" : data.getValue("limit_type").toString();
			 fieldNames = "field_9" + fieldCode + "11_m,field_9" + fieldCode + "11_y,"
					 	+ "field_9" + fieldCode + "12_m,field_9" + fieldCode + "12_y,"
					 	+ "field_9" + fieldCode + "21_m,field_9" + fieldCode + "21_y,"
					 	+ "field_9" + fieldCode + "22_m,field_9" + fieldCode + "22_y,"
					 	+ "field_9" + fieldCode + "31_m,field_9" + fieldCode + "31_y,"
					 	+ "field_9" + fieldCode + "32_m,field_9" + fieldCode + "32_y,"
					 	+ "field_9" + fieldCode + "41_m,field_9" + fieldCode + "41_y,"
					 	+ "field_9" + fieldCode + "42_m,field_9" + fieldCode + "42_y,"
					 	+ "field_9" + fieldCode + "51_m,field_9" + fieldCode + "51_y,"
					 	+ "field_9" + fieldCode + "52_m,";
		 
			 if(i != dataList.size()-1){
				 fieldNames += "field_9" + fieldCode + "52_y,";
			 } else {
				 fieldNames += "field_9" + fieldCode + "52_y";
			 }
			 fieldNamesBuffer.append(fieldNames.trim());
			 dataMap.put(fieldCode, fieldNames.trim());
		 }
		 
		 dataMap.put("codeAndNameList", dataList);
		 dataMap.put("fieldNames", fieldNamesBuffer.toString());
		 return dataMap;
	 }
	 
	 @SuppressWarnings("rawtypes")
	 private static List getAddItemData(String M_YEAR, String M_MONTH, String bank_code, HashMap<String, Object> codeAndNameMap){
		 List dataList = null;
		 StringBuffer sql = new StringBuffer();
		 ArrayList<String> paramList = new ArrayList<String>();
		 String orgTypeFields = (String)codeAndNameMap.get("fieldNames");
		 
		 System.out.println("orgTypeFields:" + orgTypeFields);
		 
		 if(!orgTypeFields.isEmpty()){
			 String[] fieldNames = orgTypeFields.split(",");
			 String fieldName = "";
			 sql.append(" select ");
			 for(int i=0; i<fieldNames.length; i++){
				 fieldName = fieldNames[i];
				 if(i == fieldNames.length-1){
					 sql.append(" sum(" + fieldName + ") as " + fieldName );
				 }else{
					 sql.append(" sum(" + fieldName + ") as " + fieldName + ", ");					 
				 }
			 }
			 sql.append(" from ( select");
			 String monthOrYear = "";
			 for(int i=0; i<fieldNames.length; i++){
				 fieldName = fieldNames[i];
				 monthOrYear = i%2==0 ? "month_amt" : "year_amt";
				 if(i == fieldNames.length-1){
					 sql.append(" decode(acc_code,'" + fieldName.substring(6, 12) + "'," + monthOrYear + ",0) as " + fieldName );
				 }else{
					 sql.append(" decode(acc_code,'" + fieldName.substring(6, 12) + "'," + monthOrYear + ",0) as " + fieldName + ", ");					 
				 }
			 }
			 sql.append(" from a15 ");
			 sql.append(" where m_year=? and m_month=? and bank_code=? ");
			 sql.append(" )a ");
			 paramList.add(M_YEAR);
		     paramList.add(M_MONTH);
		     paramList.add(bank_code);
			 
			 dataList =  DBManager.QueryDB_SQLParam(sql.toString(), paramList, orgTypeFields);
		 }
		 
	     return dataList;
	 }
	 
	 private static String filterValue(Object dataValue){
		 return Utility.setCommaFormat(dataValue == null ? "0" : dataValue.toString());
	 }
	 
	 private static void setCellValue(HSSFWorkbook wb, HSSFRow row, int cellIndex, HSSFCellStyle cellStyle, Object dataValue){
		 HSSFCell cell = null; //宣告一個儲存格
		 cell = row.createCell( (short)cellIndex);
		 cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		 
	     HSSFFont origFont = wb.getFontAt( cell.getCellStyle().getFontIndex() );
	     cellStyle.setFont(origFont);
		 
		 cell.setCellStyle(cellStyle);                    
		 cell.setCellValue( dataValue == null ? "0" : dataValue.toString() );
	 }
	 
	 private static void setCellNumberValue(HSSFWorkbook wb, HSSFRow row, int cellIndex, HSSFCellStyle cellStyle, Object dataValue){
		 HSSFCell cell = null; //宣告一個儲存格
		 cell = row.createCell( (short)cellIndex);
		 cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		 
		 HSSFFont origFont = wb.getFontAt( cell.getCellStyle().getFontIndex() );
	     cellStyle.setFont(origFont);
		 
		 cell.setCellStyle(cellStyle);                    
		 cell.setCellValue( filterValue(dataValue) );
	 }
	 
	 private static HSSFSheet getSheet(HSSFWorkbook wb, int index){
		 HSSFSheet sheet = wb.getSheetAt(index);
		 HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
	     //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
	     //sheet.setAutobreaks(true); //自動分頁

	     //設定頁面符合列印大小
	     sheet.setAutobreaks(false);
	     ps.setScale( (short) 100); //列印縮放百分比

	     ps.setPaperSize( (short) 9); //設定紙張大小 A4
	     //wb.setSheetName(0,"test");
		 return sheet;
	 }
	 
	 private static void setSheetTitle(HSSFSheet sheet,String M_YEAR,String M_MONTH,String bank_name,List dbDataList){
		 //設定報表表頭資料============================================
		 HSSFRow row=sheet.getRow(0);
		 HSSFCell cell=row.getCell((short)0);	       	
	     cell.setEncoding(HSSFCell.ENCODING_UTF_16);
	     cell.setCellValue(M_YEAR+"年度"+M_MONTH+"月份"+bank_name+"電子銀行及行動支付業務辦理情形"+((dbDataList == null || dbDataList.size() ==0)?"無資料存在":"")); 
	 }
}
