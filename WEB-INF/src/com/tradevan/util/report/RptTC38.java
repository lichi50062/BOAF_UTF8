/* 94.4.13 createed

報表--農業金融機構缺失統計表
*/
package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.*;
import org.apache.poi.hssf.record.ExtendedFormatRecord;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import java.io.*;
import java.text.*;
import java.util.*;
import com.tradevan.util.Utility;
import com.tradevan.util.DBManager;
import com.tradevan.util.dao.DataObject;
import com.tradevan.util.report.reportUtil;

/**
 * @author Administrator
 * 
 * TODO To change the template for this generated type comment go to Window -
 * Preferences - Java - Code Style - Code Templates
 */
public class RptTC38 {

	//	設定寬度
	int[] columnLen = {33, 10, 10, 10, 10, 10, 10, 10, 10};
	HSSFCellStyle defaultStyle;
	HSSFCellStyle noBorderDefaultStyle;

	public String createReport(List bodyList, String date) {
		String fileName = "TC38.xls";
		FileOutputStream fileOut = null;

		try {
			System.out.println("檢查追蹤歷程記錄表");
			File xlsDir = new File(Utility.getProperties("xlsDir"));
        	File reportDir = new File(Utility.getProperties("reportDir"));
			
    		if(!xlsDir.exists()){
     			if(!Utility.mkdirs(Utility.getProperties("xlsDir"))){
     		   		System.out.println(Utility.getProperties("xlsDir")+"目錄新增失敗");
     		   	
     			}
    		}
    		if(!reportDir.exists()){
     			if(!Utility.mkdirs(Utility.getProperties("reportDir"))){
     				System.out.println(Utility.getProperties("reportDir")+"目錄新增失敗");

     			}
    		}
			
//			Creating Cells
            HSSFWorkbook wb = new HSSFWorkbook();
            HSSFSheet sheet = wb.createSheet( "report" ); //建立sheet，及名稱
            HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
            //sheet.setZoom(80, 100); // 螢幕上看到的縮放大小
            //sheet.setAutobreaks(true); //自動分頁

            //設定頁面符合列印大小
            sheet.setAutobreaks( false );
            ps.setScale( ( short )100 ); //列印縮放百分比

            ps.setPaperSize( ( short )9 ); //設定紙張大小 A4
            ps.setLandscape( true ); // 設定橫印
            //ps.setFitWidth((short)14);

            HSSFFooter footer = sheet.getFooter();
            //HSSFCell cell;
            //HSSFCellStyle style;
            //HSSFFont font;
            //style = wb.createCellStyle();
            //Font = wb.createFont();

            //設定樣式和位置(請精減style物件的使用量，以免style物件太多excel報表無法開啟)
            this.defaultStyle = wb.createCellStyle(); //有框內文置中
            defaultStyle = HssfStyle.setStyle( defaultStyle, wb.createFont(),
                                               new String[] {
                                               "BORDER", "PHC", "PVC", "F10",
                                               "WRAP"} );
            this.noBorderDefaultStyle = wb.createCellStyle(); //無框內文置中
            noBorderDefaultStyle = HssfStyle.setStyle( noBorderDefaultStyle,
              wb.createFont(), new String[] {
              "PHC", "PVC", "F10", "WRAP"} );

            //自定需求style
            HSSFCellStyle titleStyle = wb.createCellStyle(); //標題用
            titleStyle = HssfStyle.setStyle( titleStyle, wb.createFont(),
                                             new String[] {
                                             "PHC", "PVC", "F24"} );
            HSSFCellStyle leftStyle = wb.createCellStyle(); //有框內文置左
            leftStyle = HssfStyle.setStyle( leftStyle, wb.createFont(),
                                            new String[] {
                                            "BORDER", "PHL", "PVC", "F10",
                                            "WRAP"} );
            HSSFCellStyle rightStyle = wb.createCellStyle(); //有框內文置右
            rightStyle = HssfStyle.setStyle( rightStyle, wb.createFont(),
                                             new String[] {
                                             "BORDER", "PHR", "PVC", "F10",
                                             "WRAP"} );
            HSSFCellStyle smallFontStyle = wb.createCellStyle(); //有框小字
            smallFontStyle = HssfStyle.setStyle( smallFontStyle, wb.createFont(),
                                                 new String[] {
                                                 "BORDER", "PHC", "PVC", "F08",
                                                 "WRAP"} );
            HSSFCellStyle noBoderStyle = wb.createCellStyle(); //無框置右
            noBoderStyle = HssfStyle.setStyle( noBoderStyle, wb.createFont(),
                                               new String[] {
                                               "PHR", "PVC", "F10", "WRAP"} );
            
            HSSFRow row;
            
            row = sheet.createRow( ( short )1 );
            String titleName = "檢查追蹤歷程記錄表";
            createCell( wb, row, ( short )1, titleName, titleStyle );
            createCell( wb, row, ( short )2, "", noBorderDefaultStyle );
            createCell( wb, row, ( short )3, "", noBorderDefaultStyle );
            createCell( wb, row, ( short )4, "", noBorderDefaultStyle );
            createCell( wb, row, ( short )5, "", noBorderDefaultStyle );
            createCell( wb, row, ( short )6, "", noBorderDefaultStyle );
            createCell( wb, row, ( short )7, "", noBorderDefaultStyle );
            createCell( wb, row, ( short )8, "", noBorderDefaultStyle );
            createCell( wb, row, ( short )9, "", noBorderDefaultStyle );
            sheet.addMergedRegion( new Region( ( short )1, ( short )1,
                                               ( short )1,
                                               ( short )9 ) );
            
            row = sheet.createRow( ( short )2 );
            String duringDate = "期間: " + date;
            createCell( wb, row, ( short )1, duringDate, noBorderDefaultStyle );
            createCell( wb, row, ( short )2, "", noBorderDefaultStyle );
            createCell( wb, row, ( short )3, "", noBorderDefaultStyle );
            createCell( wb, row, ( short )4, "", noBorderDefaultStyle );
            createCell( wb, row, ( short )5, "", noBorderDefaultStyle );
            createCell( wb, row, ( short )6, "", noBorderDefaultStyle );
            createCell( wb, row, ( short )7, "", noBorderDefaultStyle );
            createCell( wb, row, ( short )8, "", noBorderDefaultStyle );
            createCell( wb, row, ( short )9, "", noBorderDefaultStyle );
            sheet.addMergedRegion( new Region( ( short )2, ( short )1,
                                               ( short )2,
                                               ( short )9 ) );
            
            row = sheet.createRow( ( short )3 );

            createCell( wb, row, ( short )1, "功能名稱", defaultStyle );
            createCell( wb, row, ( short )2, "使用者名稱", defaultStyle );
            createCell( wb, row, ( short )3, "使用日期", defaultStyle );
            createCell( wb, row, ( short )4, "來源IP", defaultStyle );
            createCell( wb, row, ( short )5, "發文文號", defaultStyle );
            createCell( wb, row, ( short )6, "回文文號", defaultStyle );
            createCell( wb, row, ( short )7, "檢查報告編號", defaultStyle );
            createCell( wb, row, ( short )8, "事項序號", defaultStyle );
            createCell( wb, row, ( short )9, "操作類別", defaultStyle );
            
            wb.setRepeatingRowsAndColumns( 0, 1, 8, 1, 3 ); //設定表頭 為固定 先設欄的起始再設列的起始
            
            short rowNo = ( short )4;
            
    
            System.out.println("bodyList.size = " + bodyList.size());
            if(bodyList.size()>0){
                for(int i=0; i < bodyList.size(); i++) {
                	
                	DataObject bean = (DataObject)bodyList.get(i);
                	String pg_name= bean.getValue("pg_name") != null ? (String)bean.getValue("pg_name") : "" ;
                	String muser_name = bean.getValue("muser_name") != null ? (String)bean.getValue("muser_name") : "" ;
                	String use_date = bean.getValue("use_date") != null ? (String)bean.getValue("use_date") : "" ; 
                	String ip_address = bean.getValue("ip_address") != null ? (String)bean.getValue("ip_address") : "" ;
                	String sn_docno = bean.getValue("sn_docno") != null ? (String)bean.getValue("sn_docno") : "" ;
                    String rt_docno = bean.getValue("rt_docno") != null ? (String)bean.getValue("rt_docno") : "" ; 
                    String reportno = bean.getValue("reportno") != null ? (String)bean.getValue("reportno") : "" ;
                    String item_no = bean.getValue("item_no") != null ? (String)bean.getValue("item_no") : "" ;
                    String update_type = bean.getValue("update_type") != null ? (String)bean.getValue("update_type") : "" ; 
                	
                    int use_dateY = Integer.parseInt(use_date.substring(0,4))-1911;
                    String use_dateM = use_date.substring(5,7);
                    String use_dateD = use_date.substring(8,10);
                    use_date = use_dateY+"/"+use_dateM+"/"+use_dateD;
                	
                	row = sheet.createRow( rowNo );
                	createCell( wb, row, ( short )1, pg_name, leftStyle );
                    createCell( wb, row, ( short )2, muser_name, defaultStyle );
                    createCell( wb, row, ( short )3, use_date, defaultStyle );
                    createCell( wb, row, ( short )4, ip_address, defaultStyle );
                    createCell( wb, row, ( short )5, sn_docno, defaultStyle );
                    createCell( wb, row, ( short )6, rt_docno, defaultStyle );
                    createCell( wb, row, ( short )7, reportno, defaultStyle );
                    createCell( wb, row, ( short )8, item_no, defaultStyle );
                    createCell( wb, row, ( short )9, update_type, defaultStyle );
                	rowNo++;
                	
                }
            }else{
                row = sheet.createRow( ( short )4 );
                createCell( wb, row, ( short )1, "無相符資料", noBorderDefaultStyle );
                createCell( wb, row, ( short )2, "", noBorderDefaultStyle );
                createCell( wb, row, ( short )3, "", noBorderDefaultStyle );
                createCell( wb, row, ( short )4, "", noBorderDefaultStyle );
                createCell( wb, row, ( short )5, "", noBorderDefaultStyle );
                createCell( wb, row, ( short )6, "", noBorderDefaultStyle );
                createCell( wb, row, ( short )7, "", noBorderDefaultStyle );
                createCell( wb, row, ( short )8, "", noBorderDefaultStyle );
                createCell( wb, row, ( short )9, "", noBorderDefaultStyle );
                
                sheet.addMergedRegion( new Region( ( short )4, ( short )1,
                                                   ( short )4,
                                                   ( short )9 ) );
            }

            for ( int i = 1; i <= columnLen.length; i++ ) {
                sheet.setColumnWidth( ( short )i,
                                      ( short ) ( 256 * ( columnLen[i - 1] + 4 ) ) );
            }
            
            
//          設定涷結欄位
            //sheet.createFreezePane(0,1,0,1);
            footer.setRight( "Page:" + HSSFFooter.page() + " of " +
                             HSSFFooter.numPages() );

            // Write the output to a file
           
            fileOut = new FileOutputStream(reportDir + System.getProperty("file.separator")+"TC38.xls");
            wb.write( fileOut );
            fileOut.close();
           
	        //儲存

	        System.out.println("儲存完成");
            
			

		} catch (Exception e) {
			System.out.println("TC38_report:Exception  action=create_File  msg=" + e.getMessage());
			e.printStackTrace();
			fileName = null;
		} finally {
			try {
				if (fileOut != null) {
					fileOut.close();
				}
			} catch (Exception e) {
				System.out.println("TC38_report:Exception  action=close_File  msg=" + e.getMessage());
			}
		}

		return fileName;
	}

	private void setBorder(HSSFCellStyle style) {
		style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		style.setBorderRight(HSSFCellStyle.BORDER_THIN);
		style.setBorderTop(HSSFCellStyle.BORDER_THIN);
		style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
	}

	private void createCell(HSSFWorkbook wb, HSSFRow row, short column, String value) {
		createCell(wb, row, column, value, true);
	}

	private void createCell(HSSFWorkbook wb, HSSFRow row, short column, String value, boolean hasBorder) {
		HSSFCell cell = row.createCell(column);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellValue(value);
		int len = value.getBytes().length;
		if (len > columnLen[column - 1]) {
			columnLen[column - 1] = len;
		}
		HSSFCellStyle style = wb.createCellStyle();
		if (hasBorder) {
			setBorder(style);
			//style.setAlignment(HSSFCellStyle.ALIGN_CENTER);

		}
		cell.setCellStyle(style);
	}

	private void createCell(HSSFWorkbook wb, HSSFRow row, short column, String value, HSSFCellStyle style) {
		HSSFCell cell = row.createCell(column);
		cell.setEncoding(HSSFCell.ENCODING_UTF_16);
		cell.setCellValue(value);
		cell.setCellStyle(style);
	}

	private void createCell(HSSFWorkbook wb, HSSFRow row, short column, int value) {
		createCell(wb, row, column, String.valueOf(value), this.defaultStyle);
	}

	private void createCell(HSSFWorkbook wb, HSSFRow row, short column, int value, boolean hasBorder) {
		if (hasBorder == true) {
			createCell(wb, row, column, String.valueOf(value), this.defaultStyle);
		} else {
			createCell(wb, row, column, String.valueOf(value), this.noBorderDefaultStyle);
		}
	}

	protected static final String convertString(String s) {
		try {
			return new String(s.getBytes("ISO8859_1"), "Big5");
		} catch (UnsupportedEncodingException ueex) {
			return s;
		}
	}

	private String toBig5(String s) throws UnsupportedEncodingException {
		return new String(s.getBytes("ISO-8859-1"), "Big5");
	}
	private String toISO8859_1(String s) throws UnsupportedEncodingException {
		return new String(s.getBytes("Big5"), "ISO-8859-1");
	}

	/*
	 * p = 位置 h = 橫向 v = 垂直 c = 置中 PHL PHR PVT PVB
	 * 
	 *  
	 */

}
