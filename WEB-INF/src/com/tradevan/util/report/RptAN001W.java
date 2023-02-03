/*
 * Created on 2006/10/18 by ABYSS Allen
 * 歷年來全體農漁會信用部簡明存款結構比較表
 * fixed 99.06.04 sql injection by 2808
 */

package com.tradevan.util.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.Region;
import java.io.*;
import java.util.*;
import java.sql.*;

import com.tradevan.util.Utility;
import com.tradevan.util.Utility_report;
import com.tradevan.util.dao.RdbCommonDao;

public class RptAN001W {

	public static String createRpt(String startYear, String endYear, String bankType, int priceUtil) {
		String[] priceUtilStr = new String[]{"元","仟元","萬元","百萬元","仟萬元","億元"};
		boolean debug = false;
		if(debug) System.out.println("RptAN001 createRpt() Debug Start ...");
		String errMsg = "";
		List dbData = new ArrayList();
		StringBuffer sqlCmd = new StringBuffer() ;
		List paramList = new ArrayList() ;
		String u_year= "99" ;
		if(!"".equals(startYear) && Integer.parseInt(startYear) > 99) {
			u_year = "100" ;
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
		try {
			conn =(new RdbCommonDao("")).newConnection();
			if(debug) System.out.println("conn="+conn);
			sqlCmd.append("SELECT SUM(a.AMT) FROM A01 a, (select * from bn01 where m_year=?) b WHERE a.BANK_CODE=b.BANK_NO AND a.M_YEAR=?"
				+" AND a.M_MONTH=12 AND a.ACC_CODE=? AND a.BANK_CODE != ? " );
			paramList.add(u_year) ;
			paramList.add("") ;
			paramList.add("220000") ;
			paramList.add("8888888") ;
			if(!bankType.equals("ALL")){
				sqlCmd.append(" AND b.BANK_TYPE=? " );
				paramList.add(bankType) ;
			}
			if(debug) System.out.println("sqlCmd="+sqlCmd);
			long sumAmount=0,comAmount=0,preSumAmount=0;
			for(int yi=Integer.parseInt(startYear); yi<=Integer.parseInt(endYear);yi++){
				String[] data = new String[6];
				pst = conn.prepareStatement(sqlCmd.toString());
				paramList.set(1, new Integer(yi).toString()) ;
				paramList.set(2, "220000") ;
				//pst.setString(1,new Integer(yi).toString());
				//pst.setString(2,"220000");
				setPreparedStatementParameter(pst,paramList) ;
				rs = pst.executeQuery();
				if(rs.next()){
					sumAmount=rs.getLong(1);//本次存款總額
				}else{
					sumAmount=0;
				}
				if(debug) System.out.println(yi+" sumAmount="+sumAmount);
				pst = conn.prepareStatement(sqlCmd.toString());
				//pst.setString(1,new Integer(yi).toString());
				//pst.setString(2,"220900");
				paramList.set(2, "220900") ;
				setPreparedStatementParameter(pst,paramList) ;
				rs = pst.executeQuery();
				if(rs.next()){
					comAmount=rs.getLong(1);//本次公庫存款
				}else{
					comAmount=0;
				}
				if(debug) System.out.println(yi+"preSumAmount="+preSumAmount+"; comAmount="+comAmount);
				if(sumAmount>0){
					data[0]=new Integer(yi).toString();//年度
					data[1]=new Long(sumAmount-comAmount).toString();
					data[1]=String.valueOf(Math.round(Double.parseDouble(data[1])/priceUtil));
					data[2]=new Long(comAmount).toString();
					data[2]=String.valueOf(Math.round(Double.parseDouble(data[2])/priceUtil));
					data[3]=new Long(sumAmount).toString();
					data[3]=String.valueOf(Math.round(Double.parseDouble(data[3])/priceUtil));
					data[4]=new Long(sumAmount-preSumAmount).toString();
					data[4]=String.valueOf(Math.round(Double.parseDouble(data[4])/priceUtil));
					if(preSumAmount>0){
						data[5]=Utility_report.round(Long.toString((sumAmount-preSumAmount)*100),Long.toString(preSumAmount),2);
					}else{
						data[5]="0.00";
					}
				}else{
					data[0]=new Integer(yi).toString();//年度
					data[1]="";
					data[2]="";
					data[3]="";
					data[4]="";
					data[5]="";
				}				
				dbData.add(data);
				preSumAmount=sumAmount;
				rs.close();
				pst.close();
			}
			if(debug) System.out.println("dbData.size=" + dbData.size());
			//===開始製作報表
			HSSFWorkbook wb = new HSSFWorkbook();
			HSSFSheet sheet = wb.createSheet("New Sheet 1");
			HSSFPrintSetup ps = sheet.getPrintSetup(); //取得設定
			//設定頁面符合列印大小
            sheet.setZoom(100, 100); // 螢幕上看到的縮放大小
            sheet.setAutobreaks(false); //自動分頁            
            ps.setScale((short)85); //列印縮放百分比
            ps.setPaperSize((short)9); //設定紙張大小 A4            
            //ps.setLandscape(true); // 設定橫印
            //ps.setFitWidth((short)1);//調整成幾頁寬
            //ps.setFitHeight((short)1);//調整成幾頁高
			//設定欄位長度----Start 
    		short[] columnLen = new short[]{5,15,15,15,15,10};    		
    		for(int ai=0; ai<columnLen.length; ai++){
    			sheet.setColumnWidth((short)ai, (short)(256*(columnLen[ai]+4)));
    		}
    		HSSFDataFormat format = wb.createDataFormat();
			reportUtil reportUtil = new reportUtil();
			HSSFCellStyle titleStyle = reportUtil.getTitleStyle(wb); //標題用
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
			HSSFRow row = null; //宣告一列
			//設定報表表頭資料 開始============================================
			row=sheet.createRow(0);
			String titleStr = "";
			if(bankType.equals("ALL")){
				titleStr = "歷年來全體農漁會信用部簡明存款結構比較表";
			}else if(bankType.equals("6")){
				titleStr = "歷年來全體農會信用部簡明存款結構比較表";
			}else{
				titleStr = "歷年來全體漁會信用部簡明存款結構比較表";
			}
			reportUtil.createCell( wb, row, (short)0, titleStr, titleStyle);
			for(int ci=1; ci<totalColumnNums;ci++){
				reportUtil.createCell( wb, row, (short)ci, "", titleStyle);
			}
			sheet.addMergedRegion(new Region((short)0, (short)0, (short)0, (short)(totalColumnNums-1)));
			row=sheet.createRow(1);
			row.setHeightInPoints(20.0F);
			reportUtil.createCell(wb, row, (short)0, "單位：新台幣 "+priceUtilStr[priceIndex], noBorderRightStyle);
			for(int ci=1; ci<totalColumnNums;ci++){
				reportUtil.createCell( wb, row, (short)ci, "", noBorderRightStyle);
			}
			sheet.addMergedRegion(new Region((short)1, (short)0, (short)1, (short)(totalColumnNums-1)));
			row=sheet.createRow(2);
			row.setHeightInPoints(28.0F);
			reportUtil.createCell(wb,row,(short)0, "年度", defaultStyle);
			reportUtil.createCell(wb,row,(short)1, "一般存款", defaultStyle);
			reportUtil.createCell(wb,row,(short)2, "公庫存款", defaultStyle);
			reportUtil.createCell(wb,row,(short)3, "存款總額", defaultStyle);
			reportUtil.createCell(wb,row,(short)4, "增減額", defaultStyle);
			reportUtil.createCell(wb,row,(short)5, "增減率", defaultStyle);
			wb.setRepeatingRowsAndColumns(0, 0, (totalColumnNums-1), 0, 2);//設為固定表頭(第幾個sheet,起始欄,終止欄,起始列,終止列)
			//設定報表表頭資料 結束============================================
			//判斷dbData.size()是不是0，是的話表示沒有資料
			if (dbData.size()==0) {
				int rowNo=2;
				for(int yi=Integer.parseInt(startYear); yi<=Integer.parseInt(endYear);yi++){
					row=sheet.createRow(++rowNo);
					row.setHeightInPoints(28.0F);
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
					row.setHeightInPoints(28.0F);
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
						cell.setCellStyle(numberRightStyle);
						cell.setCellValue(Long.parseLong(dataTmp[2]));
						cell = row.createCell((short)3);
						cell.setCellStyle(numberRightStyle);
						cell.setCellValue(Long.parseLong(dataTmp[3]));
						cell = row.createCell((short)4);
						cell.setCellStyle(numberRightStyle);
						cell.setCellValue(Long.parseLong(dataTmp[4]));
					}					
					reportUtil.createCell(wb,row,(short)5, dataTmp[5], doubleRightStyle);
				}//end of for
			} //end of else ((DataObject) dbData.get(0)).getValue("m_year") == null
			File reportDir = new File(Utility.getProperties("reportDir"));
			if (!reportDir.exists()) {
				if (!Utility.mkdirs(Utility.getProperties("reportDir"))) {
					errMsg += Utility.getProperties("reportDir") + "目錄新增失敗";
				}
			}
			FileOutputStream fout = new FileOutputStream(reportDir + System.getProperty("file.separator") + titleStr+".xls");
			HSSFFooter footer = sheet.getFooter();
			footer.setCenter( "Page:" + HSSFFooter.page() + " of " + HSSFFooter.numPages() );
			footer.setRight(Utility.getDateFormat("yyyy/MM/dd hh:mm aaa"));
			wb.write(fout);
			//儲存
			fout.close();
		}catch (Exception e) {
			System.out.println("//RptAN001 createRpt() Have Error.....");
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
		if(debug) System.out.println("RptAN001 createRpt() Debug End ...");
		return errMsg;
	}
	private static void setPreparedStatementParameter(PreparedStatement pst,List paramList) throws Exception{
		for(int i = 0 ;i< paramList.size() ;i++) {
			//System.out.println("i:"+(i+1)+"=========="+(String)paramList.get(i)) ;
			pst.setString(i+1,(String)paramList.get(i)) ;
		}
	}
}
