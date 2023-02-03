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
public class RptTC37 {

	//	設定寬度
	int[] columnLen = {40, 5, 40, 10};
	HSSFCellStyle defaultStyle;
	HSSFCellStyle noBorderDefaultStyle;

	public String createReport(String startDate ,String endDate , String date) {
		
		String fileName = "TC37.xls";
		FileOutputStream fileOut = null;

		try {
			System.out.println("農業金融機構缺失統計表");
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
            String titleName = "農業金融機構缺失統計表";

            createCell( wb, row, ( short )1, titleName, titleStyle );
            createCell( wb, row, ( short )2, "", noBorderDefaultStyle );
            createCell( wb, row, ( short )3, "", noBorderDefaultStyle );
            createCell( wb, row, ( short )4, "", noBorderDefaultStyle );
            
            sheet.addMergedRegion( new Region( ( short )1, ( short )1,
                                               ( short )1,
                                               ( short )4 ) );
            
            row = sheet.createRow( ( short )2 );
            String duringDate = "期間: " + date;

            createCell( wb, row, ( short )1, duringDate, noBorderDefaultStyle );
            createCell( wb, row, ( short )2, "", noBorderDefaultStyle );
            createCell( wb, row, ( short )3, "", noBorderDefaultStyle );
            createCell( wb, row, ( short )4, "", noBorderDefaultStyle );
            
            sheet.addMergedRegion( new Region( ( short )2, ( short )1,
                                               ( short )2,
                                               ( short )4 ) );
            
            row = sheet.createRow( ( short )3 );

            createCell( wb, row, ( short )1, "缺失項目", defaultStyle );
            createCell( wb, row, ( short )2, "次數", defaultStyle );
            createCell( wb, row, ( short )3, "受檢單位", defaultStyle );
            createCell( wb, row, ( short )4, "備註", defaultStyle );
            
            wb.setRepeatingRowsAndColumns( 0, 1, 4, 1, 3 ); //設定表頭 為固定 先設欄的起始再設列的起始
            
            short rowNo = ( short )4;
            
            String bank = "";
            List dbData = getReportData(startDate , endDate );
            System.out.println("dbData.size = " + dbData.size());
            for(int i=0; i < dbData.size(); i++) {
            	
            	DataObject bean = (DataObject)dbData.get(i);
            	String fault_id = bean.getValue("fault_id") != null ? (String)bean.getValue("fault_id") : "" ;
            	String fault_name = bean.getValue("fault_name") != null ? (String)bean.getValue("fault_name") : "" ;
            	String cnt = bean.getValue("cnt").toString();
            	int count = Integer.parseInt(cnt);
            	for(int j=0; j< count; j++) {
            		i=i+j;
            		DataObject bean1 = (DataObject)dbData.get(i);
            		if(bank.equals("")) {
            			bank +=  (String)bean1.getValue("bank_no") + "  " + (String)bean1.getValue("bank_name") ;
            		}
            		else {
            			bank += "，\n" + (String)bean1.getValue("bank_no") + "  " + (String)bean1.getValue("bank_name") ;
            		}
            		
            	}
            	
            	row = sheet.createRow( rowNo );
            	createCell( wb, row, ( short )1, fault_name, leftStyle );
                createCell( wb, row, ( short )2, cnt, defaultStyle );
                createCell( wb, row, ( short )3, bank, leftStyle );
                createCell( wb, row, ( short )4, "", defaultStyle );
                bank = "";
   
            	rowNo++;
            	
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
           
            fileOut = new FileOutputStream(reportDir + System.getProperty("file.separator")+"TC37.xls");
            wb.write( fileOut );
            fileOut.close();
           
	        //儲存

	        System.out.println("儲存完成");
            
			

		} catch (Exception e) {
			System.out.println("TC37_report:Exception  action=create_File  msg=" + e.getMessage());
			e.printStackTrace();
			fileName = null;
		} finally {
			try {
				if (fileOut != null) {
					fileOut.close();
				}
			} catch (Exception e) {
				System.out.println("TC37_report:Exception  action=close_File  msg=" + e.getMessage());
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
	
	//取得報表資料
    private List getReportData(String startDate,String endDate){
       		//20060126 by 2495   
		    System.out.println("startDate ="+startDate);
		    System.out.println("endDate ="+endDate);
			String u_year = "99" ;
			if(!"".equals(startDate) && Integer.parseInt(startDate.substring(0,4))>2010) {
				u_year = "100" ;
			}
			StringBuffer sqlCmd = new StringBuffer() ;
			List paramList = new ArrayList();
    		//查詢條件    
    		sqlCmd.append(" select a.fault_id, c.FAULT_NAME,  b.bank_no,  ba01.bank_name, tmp.cnt "+
                            " from exdefgoodf a, exreportf  b, ExFaultF c, (select * from ba01 where m_year=? )ba01, "+
	                        " (select a.fault_id, count(*) cnt from exdefgoodf a, exreportf  b, ExFaultF c"+ 
	                        "    where a.reportno=b.reportno"+ 
	                        "    and ((to_char(b.report_In_date,'yyyymmdd') between ?  and ? )"+
	                        "    or  (to_char(b.report_En_date,'yyyymmdd') between  ? and ? ))"+
	                        "    and a.fault_id = c.fault_id"+
	                        "    group by a.fault_id"+
	                        "    order by cnt DESC) tmp "+
	                        " where a.reportno = b.reportno"+
	                        " and b.bank_no = ba01.bank_no"+
	                        " and tmp.fault_id = a.fault_id"+
	                        " and a.fault_id = c.fault_id"+
	                        " order by cnt desc, a.fault_id, b.bank_no" );
    		paramList.add(u_year) ;
    		paramList.add(startDate);
    		paramList.add(endDate);
    		paramList.add(startDate);
    		paramList.add(endDate);
	        
            List dbData = DBManager.QueryDB_SQLParam(sqlCmd.toString(),paramList,"cnt");            
            return dbData;
    }

}
