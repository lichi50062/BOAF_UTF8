/*
 * Created on 2006/10/23 by ABYSS Allen
 * 歷年來全體農漁會逾放_本期損益_淨值_資產總額一覽表
 */

package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.Region;
import org.apache.poi.poifs.filesystem.*;
import java.io.*;
import java.util.*;
import java.sql.*;

import com.tradevan.util.Utility;
import com.tradevan.util.Utility_report;
import com.tradevan.util.dao.RdbCommonDao;

public class RptAN003W {

	public static String createRpt(String startYear, String endYear, String bankType, int priceUtil) {
		String[] priceUtilStr = new String[]{"元","仟元","萬元","百萬元","仟萬元","億元"};
		boolean debug = false;
		if(debug) System.out.println("RptAN003 createRpt() Debug Start ...");
		String errMsg = "";
		List dbData = new ArrayList();
		
		StringBuffer sqlCmd = new StringBuffer() ;
		List paramList = new ArrayList () ;
		String u_year = "99" ;
		if(!"".equals(startYear) && Integer.parseInt(startYear)>99) {
			u_year = "100";
		}
		Connection conn = null;
		PreparedStatement pst = null;
		ResultSet rs = null;
		short totalColumnNums = (short)6;
		int priceIndex=1;
		if(priceUtil==1){
			priceIndex=0;
		}else if(priceUtil==1000){
			priceIndex=1;
		}else if(priceUtil==10000){
			priceIndex=2;
		}else if(priceUtil==1000000){
			priceIndex=3;
		}else if(priceUtil==10000000){
			priceIndex=4;
		}else if(priceUtil==100000000){
			priceIndex=5;
		}
		if(debug) System.out.println("priceUtil="+priceUtil+";priceIndex="+priceIndex+";priceUtilStr="+priceUtilStr[priceIndex]);
		try {
			conn =(new RdbCommonDao("")).newConnection();
			if(debug) System.out.println("conn="+conn);
			long amount9900=0, amount9900_6=0, amount9900_7=0,amount1200=0, amount1200_6=0, amount1200_7=0, amount1208=0;
			long amount1208_6=0, amount1208_7=0, amount1503=0, amount1503_6=0, amount1503_7=0, amount3203=0, amount3203_6=0;
			long amount3203_7=0, amount3100=0,amount3200=0,amount1900=0,amount3000=0,amount1000=0;
			for(int yi=Integer.parseInt(startYear); yi<=Integer.parseInt(endYear);yi++){
				//===2006/12/21 修正,先查出農會的資料
				if(!bankType.equals("7")){
					sqlCmd.setLength(0) ;
					sqlCmd.append("SELECT SUM(a.AMT) AMOUNT, a.ACC_CODE FROM A01 a, (select * from bn01 where m_year=?) b"
						+" WHERE a.BANK_CODE=b.BANK_NO AND b.BANK_TYPE= ?  AND a.M_YEAR=? AND a.M_MONTH=12"
						+" AND a.ACC_CODE IN (?,?,?,?,?,?,?,?) "
						+ "AND a.BANK_CODE != ? GROUP BY a.M_YEAR, a.ACC_CODE");
					paramList.add(u_year) ;
					paramList.add("6") ;
					paramList.add(new Integer(yi).toString()) ;
					paramList.add("990000") ;
					paramList.add("120000") ;
					paramList.add("120800") ;
					paramList.add("150300") ;
					paramList.add("320300") ;
					paramList.add("310000") ;
					paramList.add("320000") ;
					paramList.add("190000") ;
					paramList.add("8888888") ;
					if(debug) System.out.println("sqlCmd="+sqlCmd);
					pst = conn.prepareStatement(sqlCmd.toString());
					setPreparedStatementParameter(pst,paramList) ; 
					//pst.setString(1,new Integer(yi).toString());
					rs = pst.executeQuery();
					while(rs.next()){
						long valueTmp = rs.getLong(1);
						String accCodeTmp = rs.getString(2);
						if(accCodeTmp.equals("990000")){
							amount9900_6=valueTmp;
							if(debug) System.out.println(yi+";amount9900_6="+amount9900_6);
						}else if(accCodeTmp.equals("120000")){
							amount1200_6=valueTmp;
							if(debug) System.out.println(yi+";amount1200_6="+amount1200_6);
						}else if(accCodeTmp.equals("120800")){
							amount1208_6=valueTmp;
							if(debug) System.out.println(yi+";amount1208_6="+amount1208_6);
						}else if(accCodeTmp.equals("150300")){
							amount1503_6=valueTmp;
							if(debug) System.out.println(yi+";amount1503_6="+amount1503_6);
						}else if(accCodeTmp.equals("320300")){
							amount3203_6=valueTmp;
							if(debug) System.out.println(yi+";amount3203_6="+amount3203_6);
						}else if(accCodeTmp.equals("310000")){
							amount3100=valueTmp;
							if(debug) System.out.println(yi+";amount3100="+amount3100);
						}else if(accCodeTmp.equals("320000")){
							amount3200=valueTmp;
							if(debug) System.out.println(yi+";amount3200="+amount3200);
						}else if(accCodeTmp.equals("190000")){
							amount1900=valueTmp;
							if(debug) System.out.println(yi+";amount1900="+amount1900);
						}
					}
					rs.close();
					pst.close();
					paramList.clear() ;
					sqlCmd.setLength(0) ;
				}
				//===農會結束換找漁會的
				if(!bankType.equals("6")){
					sqlCmd.setLength(0) ;
					sqlCmd.append("SELECT SUM(a.AMT) AMOUNT, a.ACC_CODE FROM A01 a, (select * from bn01 where m_year=?) b"
						+" WHERE a.BANK_CODE=b.BANK_NO AND b.BANK_TYPE=? AND a.M_YEAR=? AND a.M_MONTH=12"
						+" AND a.ACC_CODE IN (?,?,?,?,?,?,?) "
						+ "AND a.BANK_CODE != ? GROUP BY a.M_YEAR, a.ACC_CODE" );
					paramList.add(u_year) ;
					paramList.add("7") ;
					paramList.add(new Integer(yi).toString()) ;
					paramList.add("990000") ;
					paramList.add("120000") ;
					paramList.add("120800") ;
					paramList.add("150300") ;
					paramList.add("320300") ;
					paramList.add("300000") ;
					paramList.add("100000") ;
					paramList.add("8888888") ;
					if(debug) System.out.println("sqlCmd="+sqlCmd);
					pst = conn.prepareStatement(sqlCmd.toString());
					setPreparedStatementParameter(pst,paramList) ;
					//pst.setString(1,new Integer(yi).toString());
					rs = pst.executeQuery();
					while(rs.next()){
						long valueTmp = rs.getLong(1);
						String accCodeTmp = rs.getString(2);
						if(accCodeTmp.equals("990000")){
							amount9900_7=valueTmp;
							if(debug) System.out.println(yi+";amount9900_7="+amount9900_7);
						}else if(accCodeTmp.equals("120000")){
							amount1200_7=valueTmp;
							if(debug) System.out.println(yi+";amount1200_7="+amount1200_7);
						}else if(accCodeTmp.equals("120800")){
							amount1208_7=valueTmp;
							if(debug) System.out.println(yi+";amount1208_7="+amount1208_7);
						}else if(accCodeTmp.equals("150300")){
							amount1503_7=valueTmp;
							if(debug) System.out.println(yi+";amount1503_7="+amount1503_7);
						}else if(accCodeTmp.equals("320300")){
							amount3203_7=valueTmp;
							if(debug) System.out.println(yi+";amount3203_7="+amount3203_7);
						}else if(accCodeTmp.equals("300000")){
							amount3000=valueTmp;
							if(debug) System.out.println(yi+";amount3000="+amount3000);
						}else if(accCodeTmp.equals("100000")){
							amount1000=valueTmp;
							if(debug) System.out.println(yi+";amount1000="+amount1000);
						}
					}
					rs.close();
					pst.close();
					paramList.clear() ;
					sqlCmd.setLength(0) ;
				}
				String data[] = new String[6];
				data[0] = yi+"";
				amount9900 = amount9900_6 + amount9900_7;
				amount1200 = amount1200_6 + amount1200_7;
				amount1208 = amount1208_6 + amount1208_7;
				amount1503 = amount1503_6 + amount1503_7;
				amount3203 = amount3203_6 + amount3203_7;
				if(amount9900>0){
					data[1] = new Long(amount9900).toString();//逾放金額
					data[1]=String.valueOf(Math.round(Double.parseDouble(data[1])/priceUtil));
					long valueTmp = amount1200+amount1208+amount1503;
					data[2] = Utility_report.round(new Long(amount9900*100).toString(), new Long(valueTmp).toString(), 2);
					data[3] = new Long(amount3203).toString();
					data[3]=String.valueOf(Math.round(Double.parseDouble(data[3])/priceUtil));
					data[4] = new Long(amount3100+amount3200+amount3000).toString();
					data[4]=String.valueOf(Math.round(Double.parseDouble(data[4])/priceUtil));
					data[5] = new Long(amount1900+amount1000).toString();
					data[5]=String.valueOf(Math.round(Double.parseDouble(data[5])/priceUtil));
				}else{
					data[1]="";
					data[2]="";
					data[3]="";
					data[4]="";
					data[5]="";
				}
				dbData.add(data);
			}
			if(debug) System.out.println("dbData.size=" + dbData.size());
			HSSFWorkbook wb = new HSSFWorkbook();
			HSSFSheet sheet = wb.createSheet("New Sheet 1");
			HSSFPrintSetup ps = sheet.getPrintSetup(); //取得列印設定
			//設定頁面符合列印大小
            sheet.setZoom(100, 100); // 螢幕上看到的縮放大小
            sheet.setAutobreaks(false); //自動分頁            
            ps.setScale((short)80); //列印縮放百分比
            ps.setPaperSize((short)9); //設定紙張大小 A4
			HSSFRow row = null; //宣告一列
			short[] columnLen = new short[]{5,16,11,16,16,16};
    		for(int ai=0; ai<columnLen.length; ai++){
    			sheet.setColumnWidth((short)ai, (short)(256*(columnLen[ai]+4)));
    		}
			HSSFDataFormat format = wb.createDataFormat();
			reportUtil reportUtil = new reportUtil();
			HSSFCellStyle titleStyle = reportUtil.getTitleStyle(wb); //標題用
			HSSFFont titleFont = wb.createFont();
			titleFont.setFontHeightInPoints((short)18);
			titleStyle.setFont(titleFont);
			HSSFFont defaultFont = wb.createFont();
			defaultFont.setFontHeightInPoints((short)12);
			HSSFCellStyle defaultStyle = reportUtil.getDefaultStyle(wb);//有框內文置中
			defaultStyle.setFont(defaultFont);
			HSSFCellStyle leftStyle = reportUtil.getLeftStyle(wb);
			leftStyle.setFont(defaultFont);
			HSSFCellStyle numberRightStyle = reportUtil.getRightStyle(wb);//有框內文置右
			numberRightStyle.setDataFormat(format.getFormat("#,##0"));
			numberRightStyle.setFont(defaultFont);
			HSSFCellStyle doubleRightStyle = reportUtil.getRightStyle(wb);//有框內文置右
			doubleRightStyle.setDataFormat(format.getFormat("#,##0.00"));
			doubleRightStyle.setFont(defaultFont);
			HSSFCellStyle noBorderRightStyle = reportUtil.getNoBoderStyle(wb);
			noBorderRightStyle.setFont(defaultFont);
			reportUtil.setDefaultStyle(defaultStyle);
			//設定報表表頭資料 開始============================================
			row=sheet.createRow(0);
			String titleStr="";
			if(bankType.equals("ALL")){
				titleStr = "歷年來全體農漁會逾放、本期損益、淨值、資產總額一覽表";
			}else if(bankType.equals("6")){
				titleStr = "歷年來全體農會逾放、本期損益、淨值、資產總額一覽表";
			}else{
				titleStr = "歷年來全體漁會逾放、本期損益、淨值、資產總額一覽表";
			}
			reportUtil.createCell( wb, row, (short)0, titleStr, titleStyle);
			for(int ci=1; ci<totalColumnNums;ci++){
				reportUtil.createCell( wb, row, (short)ci, "", titleStyle);
			}
			sheet.addMergedRegion(new Region((short)0, (short)0, (short)0, (short)(totalColumnNums-1)));
			row=sheet.createRow(1);
			row.setHeightInPoints(20.0F);
			reportUtil.createCell(wb, row, (short)0, "單位：新台幣 "+priceUtilStr[priceIndex]+",%", noBorderRightStyle);
			for(int ci=1; ci<totalColumnNums;ci++){
				reportUtil.createCell( wb, row, (short)ci, "", noBorderRightStyle);
			}
			sheet.addMergedRegion(new Region((short)1, (short)0, (short)1, (short)(totalColumnNums-1)));
			row=sheet.createRow(2);
			row.setHeightInPoints(24.0F);
			reportUtil.createCell(wb,row,(short)0, "年度", defaultStyle);
			reportUtil.createCell(wb,row,(short)1, "逾放金額", defaultStyle);
			reportUtil.createCell(wb,row,(short)2, "逾放比例", defaultStyle);
			reportUtil.createCell(wb,row,(short)3, "本期損益", defaultStyle);
			reportUtil.createCell(wb,row,(short)4, "淨值", defaultStyle);
			reportUtil.createCell(wb,row,(short)5, "資產總額", defaultStyle);
			wb.setRepeatingRowsAndColumns(0, 0, (totalColumnNums-1), 0, 2);//設為固定表頭(第幾個sheet,起始欄,終止欄,起始列,終止列)
			//設定報表表頭資料 結束============================================
			//判斷dbData.size()是不是0，是的話表示沒有資料
			if (dbData.size() == 0) {
				int rowNo=2;
				for(int yi=Integer.parseInt(startYear); yi<=Integer.parseInt(endYear);yi++){
					row=sheet.createRow(++rowNo);
					row.setHeightInPoints(30.0F);
					reportUtil.createCell(wb,row,(short)0, yi+"", defaultStyle);
					reportUtil.createCell(wb,row,(short)1, "", leftStyle);
					reportUtil.createCell(wb,row,(short)2, "", leftStyle);
					reportUtil.createCell(wb,row,(short)3, "", leftStyle);
					reportUtil.createCell(wb,row,(short)4, "", leftStyle);
					reportUtil.createCell(wb,row,(short)5, "", leftStyle);
					//sheet.addMergedRegion(new Region((short)rowNo, (short)1, (short)rowNo, (short)(totalColumnNums-1)));
				}
			}else {
				int rowNo=2;
				for(int li=0;li < dbData.size();li++){
					HSSFCell cell=null;
					row=sheet.createRow(++rowNo);
					row.setHeightInPoints(30.0F);
					String[] dataTmp = (String[])dbData.get(li);
					reportUtil.createCell(wb,row,(short)0, dataTmp[0], defaultStyle);
					if(dataTmp[1].trim().length()==0){
						reportUtil.createCell(wb,row,(short)1, "", leftStyle);
						reportUtil.createCell(wb,row,(short)2, dataTmp[2], leftStyle);
						reportUtil.createCell(wb,row,(short)3, dataTmp[3], leftStyle);
						reportUtil.createCell(wb,row,(short)4, dataTmp[4], leftStyle);
						reportUtil.createCell(wb,row,(short)5, dataTmp[5], leftStyle);
						//sheet.addMergedRegion(new Region((short)rowNo, (short)1, (short)rowNo, (short)(totalColumnNums-1)));
					}else{
						cell = row.createCell((short)1);
						cell.setCellStyle(numberRightStyle);
						cell.setCellValue(Long.parseLong(dataTmp[1]));
						cell = row.createCell((short)2);
						cell.setCellStyle(doubleRightStyle);
						cell.setCellValue(dataTmp[2]);
						cell = row.createCell((short)3);
						cell.setCellStyle(numberRightStyle);
						cell.setCellValue(Long.parseLong(dataTmp[3]));
						cell = row.createCell((short)4);
						cell.setCellStyle(numberRightStyle);
						cell.setCellValue(Long.parseLong(dataTmp[4]));
						cell = row.createCell((short)5);
						cell.setCellStyle(numberRightStyle);
						cell.setCellValue(Long.parseLong(dataTmp[5]));
					}					
				}//end of for
			} //end of else ((DataObject) dbData.get(0)).getValue("m_year") == null
			File reportDir = new File(Utility.getProperties("reportDir"));
			if (!reportDir.exists()) {
				if (!Utility.mkdirs(Utility.getProperties("reportDir"))) {
					errMsg += Utility.getProperties("reportDir") + "目錄新增失敗";
				}
			}
			FileOutputStream fout = new FileOutputStream(reportDir + System.getProperty("file.separator")+titleStr+".xls");
			HSSFFooter footer = sheet.getFooter();
			footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
			wb.write(fout);
			//儲存
			fout.close();
		}catch (Exception e) {
			System.out.println("//RptAN003 createRpt() Have Error.....");
			e.printStackTrace();
			System.out.println("//-------------------------------------");
		}finally{
			try{
			    if(rs != null){
	               rs.close();
	               rs = null;//104.10.06
	            }
	            if(pst != null){
	               pst.close();
	               pst = null;//104.10.06
	            }
	            if(!conn.isClosed()){//104.10.06    
	                conn.close();
	                conn = null;
	            }
			}catch(Exception sqlEx){
				conn=null;
			}
		}
		if(debug) System.out.println("RptAN003 createRpt() Debug End ...");
		return errMsg;
	}
	private static void setPreparedStatementParameter(PreparedStatement pst,List paramList) throws Exception{
		for(int i = 0 ;i< paramList.size() ;i++) {
			//System.out.println("i:"+(i+1)+"=========="+(String)paramList.get(i)) ;
			pst.setString(i+1,(String)paramList.get(i)) ;
		}
	}
}
